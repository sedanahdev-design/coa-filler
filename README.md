# CoA Filler

CoA Filler is a small local web application for export-documentation teams. It
reads a supplier's Certificate of Analysis (CoA) and other shipping paperwork,
uses AI to pull the data out, and produces finished, fully-editable Word
documents that follow your own templates and layouts. Everything runs on your
own machine; the only external service it contacts is the OpenAI API, and only
for the text-extraction step.

## Overview

The application is organised into four areas, reachable from the top navigation
bar once the server is running:

**PDF to Word.** Upload a Certificate of Analysis PDF and either choose one of
your saved Word templates or ask the app to keep the PDF's own layout. When a
template is chosen, the AI reads the PDF and places the header fields, the
test-results table and the signature or stamp into the template while preserving
its formatting and its company logo. When "keep original layout" is chosen, the
PDF is simply rebuilt as a Word document with no template involved and no AI
call. The result is always a normal, editable `.docx` file.

**Compare CoA.** Upload a Certificate of Analysis together with a second
document (a purchase order, a specification sheet, another CoA, an image, and so
on). The app extracts both and reports, parameter by parameter, where the two
agree and where they differ, along with an overall pass or fail summary.

**Customers.** A simple address book of customer and consignee profiles. Each
entry holds a name and a block of free text (full address, tax numbers, contact
details). These profiles feed the Document Forms area.

**Document Forms.** Select a saved customer and a forwarder or house form, then
upload the relevant CoA and supplier documents. The app extracts the shipment
details and generates the matching set of export documents for that form, such
as the commercial invoice, the packing or weight list and the certificate of
origin. Supported forms currently include Vigorous, Thosco, Biesterfeld Germany,
Biesterfeld Dubai and Sedanah Jordan.

## What is in this repository

- **backend/** – the FastAPI server and all processing logic: PDF and document
  reading, the OpenAI extraction calls, the template-filling engine, the
  comparison engine, the customer store and the per-form document generators.
- **frontend/** – the single-page browser interface, one HTML file per area
  plus shared styles. Plain HTML, CSS and JavaScript with no build step.
- **templates/** – your Word templates, one `.docx` per supplier or product
  form. Sixteen templates are included to start with. This folder is the source
  of the template dropdown in the PDF-to-Word area.
- **generated/** – output folder. Every document the app produces is written
  here and also offered as a download in the browser.
- **run.bat** / **run.sh** – one-step start scripts for Windows and for
  macOS or Linux.
- **Dockerfile** / **docker-compose.yml** – container setup for running the
  app without installing Python.
- **requirements.txt** – the Python dependency list.
- **.env.example** – a template for your local configuration file.

## Requirements

You need an OpenAI API key. Beyond that, choose one of the following:

- **Docker**, if you would rather not install Python. This is the simplest
  option on a machine that already has Docker.
- **Python 3.10 or newer**, to run the app directly.

## Configuration

Copy `.env.example` to a new file named `.env` in the project root and open it
in a text editor. Set `OPENAI_API_KEY` to your key. You may optionally set
`OPENAI_MODEL` to choose a specific model; if you leave it out, the app uses a
sensible default.

The start scripts create the `.env` file for you on their first run if it does
not already exist, so you can also just run the app once, fill in the key, and
run it again.

As an alternative to the `.env` file, the browser interface has an "Advanced"
section on each screen where you can paste a key for a single request. A key
entered this way is used only for that request and is never written to disk.

## Running the application

### Using Docker

From the project root, make sure your `.env` file exists and contains your key,
then start the app with Docker Compose. The first start builds the image, which
takes a few minutes; later starts are quick. When it is running, open
`http://localhost:8420` in your browser.

The `templates/` and `generated/` folders are shared with the container, so a
new template dropped into `templates/` appears after a page reload with no
rebuild, and generated documents also appear in `generated/` on your machine.
Stop the app with Ctrl+C, or with `docker compose down` from another terminal.
Rebuild after changing the code by starting Compose again with the build option.

### Using Python directly

On Windows, run `run.bat` by double-clicking it or launching it from a
terminal. On macOS or Linux, run `run.sh` from a terminal. The first run
creates a virtual environment, installs the dependencies and prepares the `.env`
file. After you have added your key, run the same script again. It starts the
server at `http://localhost:8420`.

If you prefer to set things up by hand, create and activate a virtual
environment, install the packages listed in `requirements.txt`, create your
`.env` file, and start the server from the `backend/` folder with Uvicorn on
port 8420.

## Adding more templates

Put any `.docx` file into the `templates/` folder and reload the page. It
appears in the template dropdown automatically, named after the file. Older
`.doc` files must be converted to `.docx` first, for example by opening the file
in Word and using "Save As", or with a LibreOffice headless conversion.

## How the template filling works, and its limits

The PDF-to-Word area uses a single generic engine rather than a hand-built
routine per supplier, so it relies on a few heuristics. These work well across
the included templates, but they are worth understanding.

**Header fields** are matched by comparing the label text in your template, such
as "Batch No" or "AR. No", against the label text the AI found in the PDF. When
the two documents word a label differently, for example "Batch No" against "Lot
Number", the field may not match and the template keeps its original sample
value. The result screen lists every extracted field that could not be placed.

**Test results** are matched by parameter name. When a template row is matched
to a row from the PDF, both the specification and the result cells are
overwritten with the PDF's text, because the template is being reused only for
its layout and its old specification text belongs to whatever product it was
last filled in for. Tests present in the PDF with no matching template row are
added as new rows. Template rows whose test does not appear in the PDF are
removed, so no stale data is left behind. The result screen lists both the added
and the removed rows.

**Signatures and stamps** are placed by looking for a paragraph near wording
such as "Signature", "Authorized", "Checked by" or "Approved by". If that area
already holds an image, only that image is replaced. If no such wording is
found and the template contains more than one image, the last image in the
document is replaced instead. If the template contains only one image, it is
assumed to be the company logo and is left untouched, and the result screen
tells you that no signature was placed so you can add it in Word.

Every result screen states exactly what was matched, what was added and what
was left alone, so you always know what to check before sending a document out.

## Notes

- The app is intended to run locally and is only reachable at
  `localhost:8420` unless you deploy it somewhere yourself.
- Your OpenAI key stays in your own `.env` file or is entered per request in the
  browser. Standard OpenAI usage-based billing applies to each document
  processed.
- Scanned or image-only PDFs are supported, since the AI reads the rendered
  page images, but very low-resolution scans can reduce the accuracy of
  extracted numbers. Always review a generated document before sending it out.
- Generated files accumulate in `generated/`. Delete old ones periodically if
  you process many documents.
