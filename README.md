# Training Compliance Compiler

A browser-based tool that compiles any number of training-report CSV/Excel
exports into a single roster: one row per person, one column per training,
labeled **Complete** / **Incomplete** / **Unassigned**. It also recognizes
its own past exports, so you can track compliance history and roster changes
over time.

Runs entirely client-side (no server, no upload) using [SheetJS](https://sheetjs.com/)
to read CSV/XLSX files and [ExcelJS](https://github.com/exceljs/exceljs) to
generate a formatted `.xlsx` download.

## How it works

- **Import**: choose or drag-and-drop any number of `.csv`, `.xlsx`, or `.xls`
  files. Each file is automatically classified as either **raw data** (a
  training-system export) or a **past compilation** (a `.xlsx` this tool
  generated earlier) based on its columns — a file counts as raw data if it
  has a training-name column ("Program Title"/"Item Title") or a completion
  column ("Is Complete"/"Completion Status ID"/"Completion Status");
  otherwise it's treated as a past compilation. You can mix any combination:
  raw data alone, raw data plus one or more past compilations, or several
  past compilations with no new raw data.
- **Snapshot dates**: every past compilation gets an editable date field
  (auto-guessed from the filename, e.g. `2026-07-17` or `7.1.26`); when raw
  data is included, a single editable date applies to the newly compiled
  result. Correct these if the guess is wrong — they drive the ordering in
  the history sheet.
- **Compile (raw data)**: for each row, the training name is taken from the
  "Program Title" column, falling back to "Item Title" if Program Title is
  blank. Completion is determined from "Is Complete" = `Y`, "Completion
  Status ID" = `Online Complete`, or a "Completion Status" value that
  contains "complete". A person/training pair is **Complete** only if every
  matching row is complete, **Incomplete** if at least one row is not, and
  **Unassigned** if the person never appears with that training in any
  imported file. People are matched across raw files by "User ID" (rows
  missing a User ID are skipped). Rows are also filtered by "Legal Entity"
  (ignoring the trailing code suffix, e.g. " (42775)") — only "ENGIE North
  America Inc.", "MATEP LLC", "SoCore Energy LLC", and "SoCore Installation
  Services LLC" are kept; everything else is dropped.
- **Export**: downloads a styled `.xlsx` workbook with:
  - **Training Roster** — the current roster (the freshly compiled raw data
    if any was imported, otherwise the most recent past compilation), styled
    with a bold header, frozen top row, autofilter, and color-coded status
    cells.
  - **Compliance History** — one row per snapshot (oldest to newest by date),
    with a Complete count and % Complete for every training that appears in
    any snapshot. A single new-raw-data compile produces one row; each past
    compilation adds another.
  - **Roster Changes** — included whenever at least one past compilation was
    imported. Lists everyone present in only the oldest past compilation or
    only in the newest snapshot (the freshly compiled raw data if present,
    otherwise the most recent past compilation), tagged with which snapshot
    they came from. People are matched by User ID where available, falling
    back to name matching for older exports that predate User ID tracking.

## Running locally

No build step or install required. Serve the folder with any static file
server and open it in a browser, for example:

```
python -m http.server 8080
```

Then visit `http://localhost:8080`.

## Deploying to GitHub Pages

1. Push this repository to GitHub.
2. In the repo settings, go to **Pages**.
3. Under **Build and deployment**, set **Source** to "Deploy from a branch",
   pick the branch (e.g. `main`) and folder `/ (root)`.
4. Save. GitHub will publish the site at
   `https://<username>.github.io/<repo-name>/`.

No further configuration is needed — the site is fully static.
