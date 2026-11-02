# Data Format Fixer

Database exports often store numbers and dates as text (`"1,200"`, `"2026-03-10"`). Excel then
cannot sort, sum or filter them properly. **Data Format Fixer** takes an `.xlsx` or `.csv` file,
works out which columns hold numbers or dates, and gives you back a corrected Excel file where
those values are real Excel numbers and dates.

It uses plain Python rules. **No AI and no API key are needed or used.**

## Features

- Upload `.xlsx` or `.csv` by drag-and-drop or file picker.
- Every worksheet is analysed. Each column is classified as **Integer, Decimal, Date, DateTime,
  Text or Mixed**.
- Confident values are converted; uncertain or invalid values are kept exactly as they were and
  listed under **Warnings** with their cell addresses (for example `B7`).
- Input date order: **Auto**, **Day-first** or **Month-first**. Output date format:
  `yyyy-mm-dd`, `mm/dd/yyyy` or `dd/mm/yyyy`.
- Override the type of any column and regenerate the file.
- Ambiguous dates are flagged with an **AMBIGUOUS** tag, a notice, and a comment on the cell in
  the downloaded file.
- Download `<original_name>_corrected.xlsx`. The original file is never modified.
- Limits on file size, rows, columns, cells, sheets and processing time (configurable).

## Known limitations

- **No automatic download.** The file downloads only when you press **Download**.
- **Optional AI is not implemented.** The app never contacts any AI service.
- Saving with openpyxl keeps sheet names, order, formulas, styles and images, but **charts and
  pivot tables are dropped** (the app shows a notice), and formula results are recalculated when
  Excel opens the file.
- Only the first header row of each sheet is used.
- Dates before 1900, and dates with a time-zone offset, are kept as text.
- There is no login, so run it on your own machine only. Do not expose it to the internet as is.
- Tested with Python 3.14 and Node.js 24 on macOS. Other versions have not been tested.

## Technology and architecture

| Part | Technology |
|---|---|
| Frontend | React 19, TypeScript, Vite 7 |
| Backend | Python, FastAPI, Uvicorn |
| Excel processing | openpyxl |
| Tests | pytest, httpx (FastAPI test client) |

In development the browser talks to the Vite dev server (port 5173), which forwards `/api` requests
to the FastAPI backend (port 8000).

Each upload becomes a **job** with its own private temporary folder and a random, unguessable id.
API routes:

| Method and path | Purpose |
|---|---|
| `GET /api/config` | Upload limit and allowed extensions |
| `POST /api/jobs` | Upload a file and get the analysis and corrected file |
| `POST /api/jobs/{id}/regenerate` | Re-run with new date settings or column types |
| `GET /api/jobs/{id}/download` | Download the corrected `.xlsx` |
| `DELETE /api/jobs/{id}` | Delete the job's files |

## Folder structure

```
data-format-fixer/
├── .env.example          # settings template (copy to .env)
├── .gitignore
├── README.md
├── claude.md             # original project brief
├── requirements.md       # build instructions
├── backend/
│   ├── requirements.txt  # Python dependencies
│   ├── pytest.ini
│   ├── app/
│   │   ├── main.py       # FastAPI app and routes
│   │   ├── config.py     # settings from environment / .env
│   │   ├── detection.py  # number and date parsing rules
│   │   ├── processor.py  # column detection, conversion, warnings, xlsx output
│   │   ├── csv_reader.py # CSV encoding and delimiter handling
│   │   ├── jobs.py       # private temp folders, expiry and cleanup
│   │   └── errors.py     # user-facing errors
│   └── tests/
├── frontend/
│   ├── package.json
│   ├── vite.config.ts
│   └── src/              # App.tsx, api.ts, types.ts, styles.css, main.tsx
└── samples/
    ├── make_sample.py    # creates the demo file
    └── sample_export.xlsx  # synthetic demo data
```

## Prerequisites (macOS)

- Python 3 (`python3 --version`)
- Node.js and npm (`node --version`, `npm --version`)

Install them from <https://www.python.org/downloads/> and <https://nodejs.org/> or with Homebrew
(`brew install python node`).

## Install and set up

Run these once, from the folder that contains this README:

```bash
# Backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r backend/requirements.txt

# Frontend
cd frontend
npm install
cd ..

# Optional: custom limits
cp .env.example .env
```

## Environment variables

All settings are optional; the defaults work for local use. They are read from a `.env` file in the
project root. `.env.example` contains only placeholders and defaults, never secrets. `.env` is
ignored by Git.

| Variable | Default | Meaning |
|---|---|---|
| `MAX_UPLOAD_MB` | 25 | Largest accepted upload |
| `MAX_UNCOMPRESSED_MB` | 300 | Largest size an `.xlsx` may expand to |
| `MAX_ROWS` | 200000 | Rows per sheet |
| `MAX_COLUMNS` | 500 | Columns per sheet |
| `MAX_CELLS` | 2000000 | Cells per file |
| `MAX_SHEETS` | 50 | Worksheets per file |
| `MAX_PROCESSING_SECONDS` | 60 | Processing time limit |
| `JOB_TTL_MINUTES` | 30 | Minutes before an unused job is deleted |
| `MAX_ACTIVE_JOBS` | 50 | Jobs held at once |
| `CORS_ORIGINS` | `http://localhost:5173,http://127.0.0.1:5173` | Allowed browser origins |

## Run the app

Use two terminals, both in the project folder.

**Terminal 1: backend**

```bash
source .venv/bin/activate
cd backend
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

**Terminal 2: frontend**

```bash
cd frontend
npm run dev
```

Open **<http://localhost:5173>**. Stop either server with `Ctrl+C`.

## Run tests and build

```bash
source .venv/bin/activate
python -m pytest backend      # backend tests

cd frontend
npm run build                 # type-check and production build
```

The tests cover numeric conversion, leading zeros, long identifiers, ambiguous dates, date order
choices, mixed columns, formulas, multiple sheets, CSV encodings and delimiters, and the upload,
regenerate and download API.

## How to use

1. **Open the app** at <http://localhost:5173>. Both servers must be running.
2. **Choose date settings** (optional) before uploading:
   - **INPUT DATE ORDER**: *Auto (never guess)*, *Day-first (31/12/2026)* or
     *Month-first (12/31/2026)*.
   - **OUTPUT DATE FORMAT**: `yyyy-mm-dd`, `mm/dd/yyyy` or `dd/mm/yyyy`.
3. **Add your file.** Drop a `.xlsx` or `.csv` file onto the box, or click it to choose one, then
   press **Upload**.
4. **Wait for processing.** The status line shows *Processing*, then
   *Done · N cells converted · N kept with warnings*.
5. **Review the results.** Each sheet has a tab. For each column you see its **DETECTED** type,
   how many cells were **CONVERTED**, how many have **WARNINGS**, and a **SAMPLE**.
6. **Review Warnings.** The list at the bottom shows each kept value with its cell address and the
   reason. Ambiguous dates carry a red **AMBIGUOUS** tag.
7. **Change a column** (optional). Pick a type in **USE TYPE** (Auto, Integer, Decimal, Date,
   DateTime or Text). You can also change the date settings at the top, then press
   **Apply changes & regenerate**.
8. **Download.** Press **Download** to save `<original_name>_corrected.xlsx`. Nothing is downloaded
   automatically.
9. Press **Start over** to delete the job's files and begin again.

To try it, use `samples/sample_export.xlsx`.

### Before and after

| Cell text in your file | Column | After conversion |
|---|---|---|
| `"1,200"` | Quantity | the number 1200 (shown `1,200`) |
| `"19.99"` | Unit Price | the number 19.99 |
| `"2026-03-10"` | Order Date | a real Excel date |
| `"2026-03-11 17:45:00"` | Shipped At | a real Excel date and time |
| `"02134"` | Zip | **stays text** `02134` (leading zero) |
| `"12345678901234567890"` | Long ID | **stays text** (more than 15 digits) |
| `"2026-02-30"` | Order Date | **stays text**, listed in Warnings (not a real date) |

## How conversion works

All rules live in `backend/app/detection.py` and `backend/app/processor.py`.

- **Numbers.** Text such as `1200`, `-5`, `19.99`, `.5` or `1,200.50` becomes a real Excel number
  with a matching number format, so `1.50` still displays as `1.50`. Digits are never rounded.
- **Dates.** Text such as `2026-03-10`, `10/03/2026`, `5 March 2026` or `Mar 6, 2026` becomes a real
  Excel date, shown in your chosen output format. Values with a time (`2026-03-11 17:45:00`,
  `2:30 PM`) become date-times and keep the time.
- **Text columns.** A column that is mostly text is left unchanged.
- **Column type.** A type is chosen when at least 90% of non-empty values fit it. Otherwise the
  column is **Mixed** and nothing in it is converted.
- **Leading zeros and identifiers.** Values with leading zeros (`007`, `02134`), values starting
  with `+`, and numbers with more than 15 significant digits are kept as text, because Excel would
  change or round them. Columns whose header contains words like `phone`, `zip`, `postal`, `id`,
  `account`, `sku`, `code` or `ref` are also kept as text. If you set a column type yourself, it
  overrides the leading-zero and header rules, but numbers over 15 digits are never converted.
- **Blank cells.** Blank cells stay blank. Markers like `n/a`, `null`, `none` and `-` are treated
  as empty and left as they are.
- **Ambiguous dates.** `01/02/2026` is never guessed in **Auto**. Auto only picks an order when
  other values in the same column are unambiguous (for example `13/02/2026`); if the column mixes
  both orders, ambiguous values are kept and flagged. If you choose **Day-first** or
  **Month-first**, that choice decides ambiguous values, and values that are valid in only one
  order still convert.
- **Invalid dates.** Dates like `31/02/2026` or `2026-02-30` are kept as text and listed in
  Warnings with their cell address.
- **Ordinary numbers are never treated as dates.** `45000` stays a number.
- **Formulas.** Formula cells in an `.xlsx` are never modified.
- **Multiple worksheets.** Every sheet is processed, and sheet names and order are kept. Rows,
  columns and duplicates are never removed or reordered.
- **Original file.** Uploaded files are never overwritten; the result is a new file.

### Why CSV cannot keep Excel cell types

A CSV file is plain text. It has no cell types, number formats, formulas or styles, so everything
arrives as text. The app detects types from the text and saves the result as `.xlsx`. The encoding
(UTF-8, UTF-16 or Windows-1252) and the delimiter (comma, semicolon, tab or pipe) are detected
automatically. Text starting with `=` is written as literal text, never as a formula.

### Optional AI

Not implemented. The app does not use an AI service and has no API key setting.

## Privacy and safety

- Each upload is stored in its own private temporary folder under a random id.
- Files are deleted after `JOB_TTL_MINUTES`, when you press **Start over**, when an upload fails,
  and when the server shuts down.
- Uploaded data is not logged. Unexpected errors are logged by type only.
- File contents and extensions are validated, and size and processing limits apply.

## Troubleshooting

| Problem | What to do |
|---|---|
| `command not found: python3` or `node` | Install Python or Node.js (see Prerequisites). |
| `ModuleNotFoundError` when starting the backend | Run `source .venv/bin/activate`, then `pip install -r backend/requirements.txt`. |
| `Address already in use` | Another program uses the port. Stop it, or use `--port 8001` for the backend and update the proxy in `frontend/vite.config.ts`. |
| The page loads but upload fails | Check the backend is running at `http://127.0.0.1:8000`. |
| *The file is larger than the 25 MB limit* | Raise `MAX_UPLOAD_MB` in `.env` and restart the backend. |
| *This file is no longer available* | The job expired. Upload the file again. |
| A date was not converted | Open **Warnings**. If it says ambiguous, choose **Day-first** or **Month-first** and press **Apply changes & regenerate**. |
| Charts missing in the result | Charts and pivot tables are not preserved. Use the original file for those. |
| `.env` changes have no effect | Restart the backend. |
