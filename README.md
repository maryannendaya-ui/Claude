# Report Builder — offline Excel report tool

A single-file HTML tool that reads multiple Excel/CSV files, combines every sheet
into one **master sheet**, builds a **summary** (overview, records by source, key
figures, breakdowns, data checks), **verifies** the output against the raw data,
and exports a professionally formatted **.xlsx** workbook — borders, centered
alignment, styled headers, totals, frozen panes, auto-filter and print setup
included.

Runs 100% offline in any modern browser. No install, no internet, no data leaves
the computer.

## Use it

Open **`dist/ReportBuilder.html`** in a browser (double-click it), then:

1. **Load your files** — drop any number of .xlsx/.xls/.csv files. Every sheet is
   read; the header row is detected automatically (click *Preview* to adjust it).
   Sheets that contain the exact same data as another sheet are excluded
   automatically. Optionally drop a *master layout* file whose headers define the
   column names and order of the master sheet.
2. **Master sheet columns** — same-named columns across sheets are combined;
   similar names are matched automatically (shown in amber). Rename, reorder,
   combine, exclude, or fix a column's type.
3. **Summary** — breakdowns (by category, by month, …) are suggested from the
   data; add or change them freely, including cross-tabs ("spread across columns
   by") and sum/average/count/unique measures.
3b. **Connected lists** — drop reference data (e.g. approved vendor/contractor
   lists) into the *Connected lists* zone. Lists are never merged into the
   master. Each master row is matched by up to two rules (each rule = one master
   column against one list key column; exact value or a number embedded in text,
   e.g. "PRODUCT RELATED SERV (344140)" → group 344140). A chosen "bring" column
   (e.g. Vendor Name) carries the distinct matching values into the master, next
   to the match count. Coverage appears in the summary; the full list can be
   exported as its own sheet. Breakdowns support a "List unique of" measure,
   "show top N" (rest rolled into an Other row) and "ignore blank" — the
   suggested summary leads with Top Materials by purchase records plus their
   approved vendors, and an auto-written "Analyst Highlights" section states the
   key findings in plain English.
4. **Preview & export** — check both tabs, then download the formatted workbook.
   Options include S/N and source columns, skipping total rows, de-duplication,
   sorting, a totals row (live `SUBTOTAL` formulas, so filtering recalculates),
   and one extra sheet per value of a chosen column (e.g. per status or month).

Every report includes a **Verification** tab/sheet: an independent
reconciliation proving the output tallies with the raw input — every raw row
accounted for per source (kept + blank + repeated headers + totals + duplicates),
every numeric column's raw total vs master total, summary breakdown totals
re-added, extra sheets partitioning the master, and connected-list coverage.
Each check shows ✓ Tally or ✗ Off by N.

**Save settings** downloads a small JSON file; **Load settings** restores the
whole configuration next time, so a recurring report is one click.

## Mask Kit (dist/MaskKit.html)

A companion offline tool for working with confidential data:

1. **Mask** — load the confidential files; every text value becomes a stable
   token (`VENDOR_0042`), every ID/code becomes a different number with the same
   shape, consistently across all files, so lookups and connections still work
   on the masked copies. Amounts, dates, quantities and harmless enums are kept
   (all per-column choices, with a live preview). Masked files download together
   with a private key file (optionally passphrase-encrypted, AES-GCM).
2. Share the masked files for processing — e.g. build the report with the
   Report Builder, or send them to an AI assistant.
3. **Unmask** — load the key file plus the processed results; every token and
   remapped code is restored to the real value wherever it appears (cells,
   sentences, joined lists), preserving the Excel formatting of the processed
   file. Computed aggregates pass through untouched.

## Build from source

```bash
npm install exceljs@4.4.0 xlsx@0.18.5
node build.js   # inlines src/ + libraries into dist/ReportBuilder.html
```

`src/index.html`, `src/styles.css` and `src/app.js` hold the code;
`build.js` produces the single distributable file.
