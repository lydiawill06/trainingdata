# Training Compliance Compiler

A browser-based tool that compiles any number of training-report CSV/Excel
exports into a single roster: one row per person, one column per training,
labeled **Complete** / **Incomplete** / **Unassigned**.

Runs entirely client-side (no server, no upload) using [SheetJS](https://sheetjs.com/)
to read CSV/XLSX files and [ExcelJS](https://github.com/exceljs/exceljs) to
generate a formatted `.xlsx` download.

## How it works

- **Import**: choose or drag-and-drop any number of `.csv`, `.xlsx`, or `.xls`
  files exported from a training system.
- **Compile**: for each row, the training name is taken from the "Program
  Title" column, falling back to "Item Title" if Program Title is blank.
  Completion is determined from "Is Complete" = `Y`, "Completion Status ID" =
  `Online Complete`, or a "Completion Status" value that contains "complete".
  A person/training pair is **Complete** only if every matching row is
  complete, **Incomplete** if at least one row is not, and **Unassigned** if
  the person never appears with that training in any imported file. People
  are matched across files by first + last name.
- **Export**: downloads a styled `.xlsx` workbook (bold header, frozen top
  row, autofilter, color-coded status cells).

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
