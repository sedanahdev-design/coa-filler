# CoA &rarr; Word Template Filler

A local web app: upload a Certificate of Analysis (CoA) PDF, pick one of your Word
templates, and it uses OpenAI (vision + Structured Outputs) to read the PDF and
drop the data straight into the template's own layout &mdash; header fields, the
test-results table, and the signature/stamp &mdash; while leaving the template's
own logo untouched. The output is a normal, fully-editable `.docx`.

## What it does, concretely

1. You upload a PDF and choose a template (your 16 CoA templates are already in
   `templates/`, converted to `.docx` where needed).
2. The backend renders the PDF pages to images and extracts any raw text, then
   sends both to OpenAI (`gpt-4o` by default) with a strict JSON schema asking for:
   - every `Label: value` pair in the document header/footer (product name, batch
     no, dates, quantities, report numbers, etc.)
   - every row of the test-results table (parameter / specification / result)
   - the location of a signature or stamp, if visible, as a bounding box on the page
3. It copies your chosen template and:
   - finds matching `Label: value` fields anywhere in the template (paragraphs or
     table cells) and overwrites just the value, keeping the template's own
     formatting
   - finds the results table and updates the **Result** column row-by-row by
     matching on parameter name (new tests that don't exist in the template are
     appended as new rows; template rows with no match in the PDF are left as-is)
   - crops the signature/stamp out of the source PDF and drops it into the
     template's existing signature area, replacing any placeholder image found
     there (the template's own logo is never touched)
4. You download the resulting `.docx` and edit it freely in Word.

## Setup

### Option A: Docker (recommended if you have Docker installed)

```bash
cp .env.example .env      # then edit .env and add your OPENAI_API_KEY
docker compose up --build
```

Open **http://localhost:8420**. That's it &mdash; no Python install needed on your
machine.

- `templates/` and `generated/` are bind-mounted into the container (see
  `docker-compose.yml`), so dropping a new `.docx` into `templates/` shows up
  immediately after a page reload, no rebuild required. Generated `.docx` files
  land in `generated/` on your host machine too.
- To stop it: `Ctrl+C`, or `docker compose down`.
- To rebuild after changing the code: `docker compose up --build`.
- Without compose: `docker build -t coa-filler . && docker run -p 8420:8420 --env-file .env -v "$(pwd)/templates:/app/templates" -v "$(pwd)/generated:/app/generated" coa-filler`

### Option B: Plain Python

Requires Python 3.10+.

**Windows:** double-click `run.bat` (or run it from a terminal).
**macOS/Linux:** `./run.sh`

First run creates a virtual environment, installs dependencies, and creates a
`.env` file from `.env.example` for you to fill in. Open `.env` and set:

```
OPENAI_API_KEY=sk-...your-key...
```

Then run the script again. It starts the server at **http://localhost:8420** &mdash;
open that in your browser.

(Manual setup, if you'd rather not use the script:)
```bash
python3 -m venv venv
source venv/bin/activate        # venv\Scripts\activate on Windows
pip install -r requirements.txt
cp .env.example .env             # then edit .env and add your key
cd backend
uvicorn main:app --reload --port 8420
```

## Adding more templates

Drop any `.docx` file into `templates/` and reload the page &mdash; it shows up in
the dropdown automatically. `.doc` files need to be converted to `.docx` first
(e.g. open in Word and "Save As", or use LibreOffice: `soffice --headless
--convert-to docx yourfile.doc`).

## How matching works (and its limits)

This is a generic engine, not a hand-built template per company, so it uses a few
heuristics. They work well across all 16 of your existing templates in testing,
but are worth understanding:

- **Header fields** are matched by fuzzy-comparing the label text found in your
  template (e.g. "Batch No:", "AR. No:") against the label text the AI extracted
  from the source PDF. Differently-worded labels between the two documents (e.g.
  template says "Batch No" but the PDF says "Lot Number") may not match and will
  be left as the template's original sample value &mdash; the app tells you exactly
  which extracted fields didn't get placed, in the "Header fields extracted but
  not placed" section of the result.
- **Test results** are matched by parameter name (fuzzy); once a template row is
  matched to a row from the source PDF, both its Specification and Result columns
  are overwritten with the PDF's own text. This matters because the template is
  only being reused for its *layout* &mdash; it was originally filled out for a
  different batch (sometimes a completely different product), so its old
  specification text belongs to that other product, not yours.
  - Tests in the PDF with no matching row in the template are **appended** as new
    rows (placed right after the last real test row, before any trailing
    "Conclusion" line).
  - Template rows whose test doesn't appear anywhere in the source PDF are
    **deleted** from the output entirely, rather than left behind with stale
    data from whatever product the template was last used for. The result panel
    lists exactly which rows were removed, under "Template rows removed."
- **Signature/stamp placement**: the app looks for a paragraph containing
  keywords like "Signature", "Authorized", "Checked by", "Approved by", etc. If
  that area already has an image (common &mdash; these templates are usually
  exemplar documents with a real signature/stamp baked in), it replaces just that
  image. If no such keyword is found anywhere, and the template has more than one
  image, it falls back to replacing the last image in the document (signatures
  are conventionally near the bottom). If there's only one image total, it's
  assumed to be the company logo and is left alone; the result panel will tell
  you if no signature was placed so you can add it manually in Word.

Every result screen tells you exactly what was matched, what was appended, and
what was left untouched, so you always know what to double-check before sending
the document out.

## Shipping Instructions page (Purchase Order &rarr; Excel)

`/shipping.html` turns a customer Purchase Order PDF into the filled
**Shipping Instructions** Excel sheet (`forms/shipping/shipping_instructions.xlsx`).

1. Upload the PO. OpenAI reads: product name (the Description column's material
   name only &mdash; "Adapalene EP with Bacterial test as per typical" &rarr;
   "Adapalene EP"), quantity + unit, unit price, total, the Shipping terms
   (&rarr; Inco terms, and Air vs Sea), the Vendor's country (&rarr; Origin), and
   the Consignee block.
2. The page shows those values for you to correct.
3. The sheet is filled by editing its XML in place, so all styling, merged cells,
   column widths, conditional formatting and print setup survive, and only the
   *input* cells are written:

   | Cell | Filled with |
   | --- | --- |
   | A6  | `Air Shipment` / `Sea Shipment` |
   | B8  | Product |
   | B9  | Quantity |
   | A10 | `Price per <UNIT>:` |
   | B10 | Price per unit |
   | C9  | `Inco terms: ...` |
   | C10 | total &mdash; keeps the sheet's `=B9*B10` formula unless the PO's own total disagrees |
   | B11 | Specs (only if you type something; otherwise the template's wording stays) |
   | B12 | Origin |
   | B14 | Label type (`full` / `Neutral`) |
   | B28 | Consignee &amp; Notify Parties |

   Everything else is the sheet's own text or its own formulas, which keep
   working: `B22` Made in follows Origin, `A26` follows the label type, `C14`
   adds the neutral-label note, `A27`/`A29`/`B29` switch on Air vs Sea, and the
   document rows switch between CCPIT (China) and Chamber of Commerce.
   Cached formula results are recomputed on the way out and the workbook is
   flagged for full recalculation, so it reads correctly the moment it opens.

## Architecture

```
coa-filler/
  backend/
    main.py         FastAPI app (routes, orchestration)
    pdf_extract.py  PDF -> text / page images / embedded images (pdfplumber, pypdfium2, pypdf)
    ai_extract.py   OpenAI call + structured JSON schema + signature crop
    docx_fill.py    Generic label/value + results-table + signature fill engine (python-docx)
    po_extract.py   OpenAI call + schema for Purchase Orders (Shipping Instructions page)
    shipping_xlsx.py  Fills the Shipping Instructions .xlsx in place (zip + lxml), keeps its formulas
  frontend/
    index.html      Single-page vanilla JS/CSS UI
    shipping.html   Purchase Order -> Shipping Instructions Excel page
  templates/        Your Word templates (.docx)
  forms/            Per-customer document forms + the Shipping Instructions .xlsx template
  generated/        Output files land here (also served via /api/download)
  Dockerfile
  docker-compose.yml
```

Nothing here talks to any server except OpenAI's API (for the extraction step) &mdash;
PDF parsing and document generation both run entirely locally.

## Notes / things to know

- This is a local app: it runs on your machine and is only reachable at
  `localhost:8420` unless you deploy it somewhere yourself.
- Your OpenAI API key lives in your own `.env` file, or can be pasted per-request
  in the UI's "Advanced" section &mdash; it's never written to disk by the app.
  Standard OpenAI usage-based billing applies per document processed.
  Reference for keys/usage: https://platform.openai.com/docs
- Generated files accumulate in `generated/`; delete old ones periodically if you
  process many documents.
- Scanned/image-only PDFs work too (the AI reads the rendered page images), but
  very low-resolution scans may reduce accuracy of the extracted numbers &mdash;
  always double-check results before sending a document out.
