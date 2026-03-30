# Price Break Transformer

A browser-based utility for converting SSI pricing review files to and from the format required by Acumatica's Sales Price import.

No installation. No server. Drop the HTML file anywhere — SharePoint, a network share, a desktop shortcut — and open it in a browser.

---

## Background

When SSI conducts pricing reviews, price break data is maintained in a **column-per-break** format: one row per item, with break quantities and prices spread across adjacent columns (`Break 1`, `Break 2`, `QTY Break 1`, `QTY Break 2`, etc.). This is a natural layout for review and editing.

Acumatica's Sales Price import requires the opposite: a **row-per-break** format, where each price break is its own row with the item repeated on every line.

Previously, this conversion was handled by a VBA macro embedded in a protected Excel workbook. The macro was slow, occasionally errored out, had broken and required repair multiple times, and had no in-house maintainer. It also hardcoded assumptions about the number of breaks and the exact column positions in the review file — making it fragile against any structural variation between review cycles.

This tool replaces that macro entirely with a self-contained HTML file that runs in any modern browser.

---

## What It Does

### Direction 1 — Review → Acumatica Upload

Converts a pricing review workbook (one row per item, breaks as columns) into a flat upload file ready for Acumatica's Sales Price import scenario.

**Input:** Any pricing review sheet containing `Break 1`...`Break N` price columns and `QTY Break 1`...`QTY Break N` quantity columns. The tool auto-detects which sheet in the workbook contains the break data, regardless of what the sheet is named.

**Output:** An `.xlsx` file with a `Sales Price WS Template` sheet — one row per break per item — plus an `Issues` sheet documenting any rows that could not be converted.

**Handles:**
- Variable number of breaks per item (1–20)
- Inconsistent break counts across items
- Files where the number of QTY columns exceeds the number of price columns (structurally normal — extra QTY columns are silently ignored)
- Multiple sheet layouts across different review cycles (column positions are detected from headers, not hardcoded)
- UOM populated via either `UOM` or `Base Unit` column depending on the file variant

### Direction 2 — Acumatica Export → Review Format

Pivots a flat Acumatica price export back into the column-per-break review layout. Useful when a review needs to start from current system prices rather than a manually assembled file.

**Input:** An Acumatica Sales Price export with `Inventory ID`, `Break Qty`, and `Pending Price` (or `Price`) columns — one row per break per item.

**Output:** An `.xlsx` file with one row per item and break quantities/prices as side-by-side columns. The number of break columns auto-adjusts to match the maximum number of breaks found in the export.

---

## Issues Sheet

Every output file contains a second sheet named **Issues**. This sheet lists any rows that were skipped or partially converted, along with a specific reason:

| Reason | What it means |
|---|---|
| `Has price(s) but no QTY for break(s): 1, 2` | A price was entered for these breaks but no corresponding break quantity. The row cannot be uploaded. The source file needs to be corrected. |
| `No valid break data` | The item row had no usable price/qty pairs at all. |
| `Missing Inventory ID` | The row had no Inventory ID and was skipped entirely. |

The Issues table is also shown inline in the browser (first 10 rows) after each conversion, displayed below the success output and download button.

> Items that produced at least some valid break rows are still included in the upload file even if they also appear in Issues (partial conversion). Items with no valid breaks at all are excluded from the upload file entirely.

---

## Supported Input Files

The tool works with the SSI pricing review format. Two structural variants are supported:

| Variant | Example | Break price cols | QTY break cols |
|---|---|---|---|
| Full review format | Fall 2022 and similar | `W`–`AP` (`Break 1`–`Break 20`) | `AQ`–`BJ` (`QTY Break 1`–`QTY Break 20`) |
| Jack's condensed format | Spring 2025 review | `N`–`AF` (`Break 1`–`Break 20`) | `AG`–`AK` (`QTY Break 1`–`QTY Break 5`) |

Both are auto-detected. Any future variant that uses the same `Break N` / `QTY Break N` header convention will work without modification.

**Files that are NOT compatible with this tool:**

- Customer/contracted pricing review files (rep-level price worksheets with customer names, price codes, and sales data) — these are a separate workflow
- Vendor price lists or cost comparison workbooks
- Any file where the input sheet has already been converted to the upload format

---

## Usage

1. Open `index.html` in a browser (Chrome, Edge, or Firefox recommended)
2. Select the direction using the tabs at the top
3. Drop your file onto the drop zone or click to browse
4. Review the log output and preview table
5. If issues are shown, review them — the upload file is still generated and items with at least one valid break are included
6. Click **Download** to save the output file
7. Review the **Issues** sheet in the downloaded file before submitting to Acumatica

---

## File Naming

Output files are named automatically based on the input filename:

- `MyPriceReview.xlsx` → `MyPriceReview_ACU_UPLOAD.xlsx`
- `AcuExport.xlsx` → `AcuExport_REVIEW.xlsx`

---

## Technical Notes

- Runs entirely client-side. No data is transmitted anywhere.
- Uses [SheetJS (xlsx)](https://sheetjs.com/) for Excel parsing and generation, loaded from the Cloudflare CDN.
- Requires an internet connection only to load the SheetJS library on first use. If needed for offline use, the library can be downloaded and referenced locally.
- Compatible with `.xlsx` and `.xlsm` input files. Password-protected or encrypted workbooks cannot be read and must be unprotected before use.
