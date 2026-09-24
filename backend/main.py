"""
FastAPI backend for the CoA (Certificate of Analysis) -> Word template filler.

Endpoints:
  GET  /api/templates          -> list available .docx templates
  POST /api/generate           -> upload a PDF + pick a template -> AI-extract + fill -> returns a report + download id
  POST /api/shipping/read      -> upload a Purchase Order PDF -> AI-extract the shipping-instruction fields
  POST /api/shipping/generate  -> fill the Shipping Instructions .xlsx from those (user-checked) fields
  GET  /api/download/{file_id} -> download the generated .docx / .xlsx

Run with:  uvicorn main:app --reload --port 8420   (from the backend/ directory)
"""
from __future__ import annotations

import os
import shutil
import tempfile
import traceback
import uuid
from pathlib import Path
from typing import Optional
from urllib.parse import quote

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

load_dotenv()

import ai_extract
import any_doc_extract
import compare
import customers_store
import docx_fill
import forms_extract
import forms_registry
import pdf_extract
import pdf_layout_dump
import po_extract
import shipping_xlsx

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
GENERATED_DIR = BASE_DIR / "generated"
FRONTEND_DIR = BASE_DIR / "frontend"
GENERATED_DIR.mkdir(exist_ok=True)

app = FastAPI(title="CoA Template Filler")


@app.exception_handler(Exception)
async def _unhandled_exception_handler(request: Request, exc: Exception):
    # Belt-and-suspenders: every /api/* route below already wraps its own
    # risky calls in try/except -> HTTPException, which FastAPI turns into a
    # clean JSON body on its own. But that only covers exceptions the code
    # anticipated. Anything that slips through uncaught (a bug in a code path
    # nobody wrapped, a crash while building the final response dict, etc.)
    # would otherwise hit Starlette's default handler, which -- depending on
    # the exact server/proxy setup -- can render an HTML error page instead
    # of JSON. The frontend always does `await res.json()`, so an HTML body
    # there fails with a confusing "Unexpected token '<' ... is not valid
    # JSON" instead of ever showing what actually went wrong. Catch
    # everything at this top level and always hand back JSON instead.
    print("=" * 70)
    print(f"[coa-filler] UNHANDLED EXCEPTION on {request.method} {request.url.path}")
    traceback.print_exc()
    print("=" * 70)
    if request.url.path.startswith("/api/"):
        return JSONResponse(
            status_code=500,
            content={"detail": f"Unexpected server error: {exc}"},
        )
    raise exc

PRESERVE_LAYOUT_ID = "__preserve_layout__"


def _list_templates():
    # The "keep the PDF's own layout" option is a virtual entry, not a real
    # .docx file in templates/ -- it's pinned first since it needs no company
    # template selection at all.
    items = [
        {
            "id": PRESERVE_LAYOUT_ID,
            "name": "Keep original layout (no company template -- just convert to Word)",
        }
    ]
    for p in sorted(TEMPLATES_DIR.glob("*.docx")):
        items.append({"id": p.name, "name": p.stem})
    return items


@app.get("/api/templates")
def get_templates():
    return _list_templates()


# --------------------------------------------------------------------------- #
# Customers
# --------------------------------------------------------------------------- #

class CustomerIn(BaseModel):
    name: str
    full_text: str = ""


@app.get("/api/customers")
def list_customers():
    return customers_store.list_customers()


@app.post("/api/customers")
def create_customer(payload: CustomerIn):
    try:
        return customers_store.add_customer(payload.name, payload.full_text)
    except ValueError as e:
        raise HTTPException(400, str(e))


@app.put("/api/customers/{customer_id}")
def edit_customer(customer_id: str, payload: CustomerIn):
    try:
        return customers_store.update_customer(customer_id, payload.name, payload.full_text)
    except KeyError:
        raise HTTPException(404, "Customer not found")


@app.delete("/api/customers/{customer_id}")
def remove_customer(customer_id: str):
    try:
        customers_store.delete_customer(customer_id)
    except KeyError:
        raise HTTPException(404, "Customer not found")
    return {"status": "deleted"}


@app.post("/api/generate")
async def generate(
    file: UploadFile = File(...),
    template_id: str = Form(...),
    api_key: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    is_preserve_layout = template_id == PRESERVE_LAYOUT_ID
    template_path = None
    if not is_preserve_layout:
        template_path = TEMPLATES_DIR / template_id
        if not template_path.exists() or template_path.parent != TEMPLATES_DIR:
            raise HTTPException(400, f"Unknown template: {template_id}")

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a PDF file.")

    # "Keep original layout" mode is pure PDF parsing + docx reconstruction --
    # no AI call, no API key needed at all.
    if is_preserve_layout:
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                shutil.copyfileobj(file.file, tmp)
                tmp_path = tmp.name
            out_id = f"{uuid.uuid4().hex}.docx"
            out_path = GENERATED_DIR / out_id
            try:
                layout_report = pdf_layout_dump.convert_pdf_preserve_layout(tmp_path, str(out_path))
            except Exception as e:
                raise HTTPException(500, f"Layout-preserving conversion failed: {e}")
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

        base_name = Path(file.filename).stem
        return JSONResponse(
            {
                "mode": "preserve_layout",
                "file_id": out_id,
                "download_url": f"/api/download/{out_id}",
                "download_filename": f"{base_name} - converted.docx",
                "layout_report": layout_report,
            }
        )

    key = (api_key or "").strip() or os.getenv("OPENAI_API_KEY")
    if not key:
        raise HTTPException(
            400,
            "No OpenAI API key available. Either add OPENAI_API_KEY to the server's "
            ".env file, or paste a key into the 'OpenAI API key' field in the form.",
        )

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = tmp.name

        extraction = pdf_extract.extract_all(tmp_path)

        model_name = (model or "").strip() or os.getenv("OPENAI_MODEL") or ai_extract.DEFAULT_MODEL
        try:
            result = ai_extract.extract_coa_data(extraction, api_key=key, model=model_name)
        except Exception as e:
            print("=" * 70)
            print(f"[coa-filler] OpenAI extraction call failed (template={template_id}, model={model_name})")
            traceback.print_exc()
            print("=" * 70)
            # Use 500, not 502: some reverse proxies/hosting platforms intercept
            # 502 responses and replace the body with their own generic HTML
            # error page, which makes the browser see "<!DOCTYPE ..." instead of
            # this JSON error -- 500 is far less likely to be swallowed that way.
            raise HTTPException(500, f"OpenAI extraction failed: {e}")

        # Always log what the model actually returned -- visible in the terminal /
        # `docker compose logs` even if the browser UI is confusing. This is the
        # fastest way to tell "AI extracted nothing" apart from "AI extracted fine
        # but nothing matched the template's own label wording".
        print("=" * 70)
        print(f"[coa-filler] model={model_name} template={template_id}")
        print(f"[coa-filler] header_fields ({len(result.header_fields)}):", result.header_fields)
        print(f"[coa-filler] test_results ({len(result.test_results)}):", result.test_results)
        print(f"[coa-filler] signature:", result.signature)
        print("=" * 70)

        signature_bytes = ai_extract.crop_signature(extraction, result.signature)

        out_id = f"{uuid.uuid4().hex}.docx"
        out_path = GENERATED_DIR / out_id
        try:
            fill_report = docx_fill.fill_template(
                str(template_path),
                str(out_path),
                result.header_fields,
                result.test_results,
                signature_bytes,
            )
        except Exception as e:
            raise HTTPException(500, f"Filling the template failed: {e}")

        print(f"[coa-filler] fill_report: {fill_report}")

    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

    return JSONResponse(
        {
            "file_id": out_id,
            "download_url": f"/api/download/{out_id}",
            "download_filename": f"{template_path.stem} - filled.docx",
            "extraction": {
                "header_fields": result.header_fields,
                "test_results": result.test_results,
                "signature_detected": bool(result.signature.get("present", False)),
                "signature_description": result.signature.get("description", ""),
            },
            "fill_report": {
                "header_fields": fill_report.header_fields,
                "test_results": fill_report.test_results,
                "signature": fill_report.signature,
            },
        }
    )


@app.post("/api/compare")
async def compare_endpoint(
    coa_file: UploadFile = File(...),
    other_file: UploadFile = File(...),
    api_key: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    key = (api_key or "").strip() or os.getenv("OPENAI_API_KEY")
    if not key:
        raise HTTPException(
            400,
            "No OpenAI API key available. Either add OPENAI_API_KEY to the server's "
            ".env file, or paste a key into the 'OpenAI API key' field in the form.",
        )

    allowed_exts = {".pdf", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tiff"}

    tmp_paths = []
    try:
        def _save_upload(upload: UploadFile) -> str:
            ext = Path(upload.filename or "").suffix.lower()
            if ext not in allowed_exts:
                raise HTTPException(400, f"Unsupported file type '{ext}' for {upload.filename}. Allowed: PDF, DOCX, or common image formats.")
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                shutil.copyfileobj(upload.file, tmp)
                tmp_paths.append(tmp.name)
                return tmp.name

        coa_path = _save_upload(coa_file)
        other_path = _save_upload(other_file)

        try:
            coa_doc = any_doc_extract.read_document(coa_path, filename=coa_file.filename)
            other_doc = any_doc_extract.read_document(other_path, filename=other_file.filename)
        except Exception as e:
            raise HTTPException(400, f"Couldn't read one of the files: {e}")

        model_name = (model or "").strip() or os.getenv("OPENAI_MODEL") or compare.DEFAULT_MODEL
        try:
            result = compare.compare_documents(coa_doc, other_doc, api_key=key, model=model_name)
        except Exception as e:
            print("=" * 70)
            print(f"[coa-filler] OpenAI comparison call failed (model={model_name})")
            traceback.print_exc()
            print("=" * 70)
            raise HTTPException(500, f"OpenAI comparison failed: {e}")

        print("=" * 70)
        print(f"[coa-filler:compare] model={model_name}")
        print(f"[coa-filler:compare] overall_status={result.overall_status}")
        print(f"[coa-filler:compare] comparison={result.comparison}")
        print("=" * 70)

    finally:
        for p in tmp_paths:
            if os.path.exists(p):
                os.unlink(p)

    return JSONResponse(
        {
            "coa": result.coa,
            "other": result.other,
            "comparison": result.comparison,
            "overall_status": result.overall_status,
            "summary": result.summary,
            "field_labels": compare.FIELD_LABELS,
        }
    )


@app.get("/api/forms")
def list_forms():
    return forms_registry.list_forms()


@app.post("/api/forms/generate")
async def generate_form(
    customer_id: str = Form(...),
    form_id: str = Form(...),
    coa_files: list[UploadFile] = File(...),
    supplier_files: list[UploadFile] = File(default=[]),
    api_key: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    key = (api_key or "").strip() or os.getenv("OPENAI_API_KEY")
    if not key:
        raise HTTPException(
            400,
            "No OpenAI API key available. Either add OPENAI_API_KEY to the server's "
            ".env file, or paste a key into the 'OpenAI API key' field in the form.",
        )

    customer = customers_store.get_customer(customer_id)
    if not customer:
        raise HTTPException(404, f"Customer '{customer_id}' not found")

    generate_fn = forms_registry.get_generator(form_id)
    if generate_fn is None:
        raise HTTPException(400, f"Form '{form_id}' isn't wired up yet -- ask to have it added.")

    if not coa_files:
        raise HTTPException(400, "Upload at least one COA file")

    allowed_exts = {".pdf", ".docx", ".doc", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tiff"}
    tmp_paths = []
    try:
        def _save_upload(upload: UploadFile) -> str:
            ext = Path(upload.filename or "").suffix.lower()
            if ext not in allowed_exts:
                raise HTTPException(400, f"Unsupported file type '{ext}' for {upload.filename}.")
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                shutil.copyfileobj(upload.file, tmp)
                tmp_paths.append(tmp.name)
                return tmp.name

        try:
            coa_docs = [
                any_doc_extract.read_document(_save_upload(f), filename=f.filename) for f in coa_files
            ]
            supplier_docs = [
                any_doc_extract.read_document(_save_upload(f), filename=f.filename) for f in supplier_files
            ]
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(400, f"Couldn't read one of the uploaded files: {e}")

        model_name = (model or "").strip() or os.getenv("OPENAI_MODEL") or forms_extract.DEFAULT_MODEL
        try:
            shipment = forms_extract.extract_shipment_data(coa_docs, supplier_docs, api_key=key, model=model_name)
        except Exception as e:
            print("=" * 70)
            print(f"[coa-filler] OpenAI shipment-extraction call failed (model={model_name})")
            traceback.print_exc()
            print("=" * 70)
            raise HTTPException(500, f"OpenAI extraction failed: {e}")

        print("=" * 70)
        print(f"[coa-filler:forms] form={form_id} customer={customer_id} model={model_name}")
        print(f"[coa-filler:forms] shipment={shipment.raw}")
        print("=" * 70)

        with tempfile.TemporaryDirectory() as tmp_out:
            try:
                results = generate_fn(customer, shipment, Path(tmp_out))
            except Exception as e:
                raise HTTPException(500, f"Failed to fill the '{form_id}' form: {e}")

            files = {}
            for doc_key, path in results.items():
                out_id = f"{uuid.uuid4().hex}_{path.name}"
                shutil.copy(str(path), str(GENERATED_DIR / out_id))
                files[doc_key] = {
                    "download_url": f"/api/download/{out_id}",
                    "filename": path.name,
                }
    finally:
        for p in tmp_paths:
            if os.path.exists(p):
                os.unlink(p)

    return JSONResponse(
        {
            "form_id": form_id,
            "customer": customer["name"],
            "shipment": shipment.raw,
            "files": files,
        }
    )


# --------------------------------------------------------------------------- #
# Shipping Instructions (Purchase Order PDF -> Shipping Instructions .xlsx)
# --------------------------------------------------------------------------- #


class ShippingValues(BaseModel):
    """The sheet's input cells, as shown on the review form. Everything else in
    the workbook is either static text or one of the sheet's own formulas."""

    shipment_mode: str = "Air"       # A6  -> "Air Shipment" / "Sea Shipment"
    product: str = ""                # B8  (also drives B16 via the sheet)
    quantity: float = 0              # B9
    unit: str = "KG"                 # relabels A10 ("Price per KG:")
    price_per_unit: float = 0        # B10
    total: float = 0                 # C10 (kept as =B9*B10 when they agree)
    incoterms: str = ""              # C9  -> "Inco terms: ..."
    origin: str = ""                 # B12 (also drives B22 "Made in" + CCPIT/CoC rows)
    consignee: str = ""              # B28
    label_type: str = "full"         # B14 (also drives C14 + A26)
    specs: str = ""                  # B11 -- blank keeps the template's own wording


@app.post("/api/shipping/read")
async def shipping_read(
    po_file: UploadFile = File(...),
    api_key: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    """Read a Purchase Order PDF and return the fields for the review form."""
    key = (api_key or "").strip() or os.getenv("OPENAI_API_KEY")
    if not key:
        raise HTTPException(
            400,
            "No OpenAI API key available. Either add OPENAI_API_KEY to the server's "
            ".env file, or paste a key into the 'OpenAI API key' field in the form.",
        )

    ext = Path(po_file.filename or "").suffix.lower()
    if ext != ".pdf":
        raise HTTPException(400, "Upload the Purchase Order as a PDF.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            shutil.copyfileobj(po_file.file, tmp)
            tmp_path = tmp.name

        try:
            extraction = pdf_extract.extract_all(tmp_path)
        except Exception as e:
            raise HTTPException(400, f"Couldn't read that PDF: {e}")

        model_name = (model or "").strip() or os.getenv("OPENAI_MODEL") or po_extract.DEFAULT_MODEL
        try:
            po = po_extract.extract_po_data(extraction, api_key=key, model=model_name)
        except Exception as e:
            print("=" * 70)
            print(f"[coa-filler:shipping] OpenAI PO-extraction call failed (model={model_name})")
            traceback.print_exc()
            print("=" * 70)
            raise HTTPException(500, f"OpenAI extraction failed: {e}")

        print("=" * 70)
        print(f"[coa-filler:shipping] model={model_name} po={po.raw}")
        print("=" * 70)

        return JSONResponse({"po": po.as_dict(), "values": shipping_xlsx.build_values(po)})
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


@app.post("/api/shipping/generate")
def shipping_generate(values: ShippingValues):
    """Fill the Shipping Instructions workbook from the (user-checked) fields."""
    data = values.model_dump() if hasattr(values, "model_dump") else values.dict()
    out_id = f"{uuid.uuid4().hex}.xlsx"
    try:
        report = shipping_xlsx.fill_shipping_instructions(data, GENERATED_DIR / out_id)
    except Exception as e:
        print("=" * 70)
        traceback.print_exc()
        print("=" * 70)
        raise HTTPException(500, f"Filling the Shipping Instructions sheet failed: {e}")

    filename = shipping_xlsx.suggested_filename(data)
    return JSONResponse(
        {
            "download_url": f"/api/download/{out_id}?name={quote(filename)}",
            "filename": filename,
            "cells": {k: str(v) for k, v in report["cells"].items()},
            "notes": report["notes"],
        }
    )


@app.get("/api/download/{file_id}")
def download(file_id: str, name: Optional[str] = None):
    # file_id is a uuid hex + .docx/.xlsx generated by us -- reject anything else defensively.
    if "/" in file_id or "\\" in file_id or not file_id.endswith((".docx", ".xlsx")):
        raise HTTPException(400, "Invalid file id")
    path = GENERATED_DIR / file_id
    if not path.exists():
        raise HTTPException(404, "File not found (it may have already been cleaned up)")
    is_xlsx = file_id.endswith(".xlsx")
    # `name` only sets the name the browser saves it under (the Shipping
    # Instructions sheet is named after its product) -- never a file path.
    download_name = Path(name).name if name else path.name
    if is_xlsx and not download_name.lower().endswith(".xlsx"):
        download_name += ".xlsx"
    return FileResponse(
        path,
        filename=download_name,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if is_xlsx
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )


# Serve the frontend (index.html, css, js) for everything else.
app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
