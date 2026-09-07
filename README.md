# Call Register Extractor (web app)

A tiny, self-hosted web app for the monthly job: upload your Goal Sheet
workbook + a zip of Sotax service-report PDFs, review the extracted rows in
your browser, edit anything (including filling in the blanks), then download
the finished workbook.

This is deliberately "dumb" — no lookups, no guessing, no history matching.
It copies a fixed set of fields straight off each report and leaves
everything else blank for you to fill in.

## What gets filled in vs. left blank

| Column | Filled from | Notes |
|---|---|---|
| Sr. No. | auto-numbered | continues from the last row in your sheet — the only non-literal field, since row numbering doesn't exist in a PDF |
| Location | filename | second `_`-separated token in the PDF's filename |
| Service Type | filename | reads BD/AMC/PM/Training from the filename and expands it (BD → Breakdown) |
| Customer Details | PDF | left-column customer name, as printed |
| Contact Person | PDF | as printed, no reformatting |
| Contact Email ID | PDF | as printed |
| Instrument Model | PDF | first row of the parts table only (no combining with accessory rows) |
| Serial Number | PDF | first row of the parts table |
| Call Ticket Number | PDF | "Service Contract No." field |
| Call Ticket Date | PDF | earliest signature timestamp at the bottom of the report |
| Engineer Allocated | PDF | "Service Engineer" field, with the leading numeric code stripped |

**Everything else is blank** — Phone, Classification (Local/Outstation), Call
Status, Follow-Up Remarks, all rating/formula columns, etc. Fill those in
directly in the preview table before saving, or afterward in Excel.

## Project layout

```
main.py             FastAPI app: serves the page + the two API endpoints
static/index.html   The whole frontend (plain HTML/JS, no build step)
requirements.txt    Python dependencies
```

## Run it locally

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

Then open http://localhost:8000 in a browser.

## Deploy it online (so you get a permanent link)

Any host that runs a Python web service works. **Render** has the simplest
free-tier setup:

1. Push this folder to a GitHub repo.
2. On [render.com](https://render.com), click **New → Web Service**, connect
   the repo.
3. Set:
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
4. Deploy. Render gives you a permanent `https://your-app.onrender.com` link.

(Railway and Fly.io work the same way — connect the repo, same build/start
commands, they auto-detect the Python app.)

## How it works, in short

- **`POST /api/extract`** — you upload the workbook + zip. The server reads
  each PDF (using `pdfplumber`, matching text by its position on the page so
  long customer names don't run into the neighboring column), reads the
  filename for Location/Service Type, and returns a JSON table: one row per
  report, in chronological order, with only the fields above filled in. The
  original workbook is held in server memory for up to 1 hour, keyed by a
  one-time session id — nothing is written to a file yet.
- **The browser** shows this as an editable table. Auto-filled cells are
  tinted light green; blank ones are plain white. You can edit any cell,
  delete a row, or add a blank row.
- **`POST /api/generate`** — sends your edited table back. The server
  appends it to the *original* uploaded workbook (copying the last row's
  cell formatting so the new rows look consistent), and streams the finished
  `.xlsx` back for download. Nothing else in the workbook is touched.

## Notes & limitations

- This expects the exact "Call Register" sheet layout from your existing
  workbook (36 columns, header row 1). If that layout ever changes, the
  column-index constants near the top of `main.py` (`AUTO_COLUMNS`) will
  need updating to match.
- Filenames need to follow the existing convention
  (`Company_Location_Model_Serial_Type_...date.pdf`) for Location/Service
  Type to come through — if a filename doesn't contain one of
  BD/AMC/PM/Training/CV/LU, Service Type is left blank.
- Sessions (the uploaded workbook held between Extract and Save) live in
  server memory and expire after 1 hour, or immediately after you download.
  If you take a long break mid-review, you may need to re-upload.
- Tested by running the same June → July → August workflow used previously
  and confirming the literal fields (customer, contact, email, model,
  serial, ticket number/date, engineer) match exactly what's on each report.
