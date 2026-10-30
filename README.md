# Report Builder — offline Excel report tool

A single-file HTML tool that reads multiple Excel/CSV files, combines every sheet
into one **master sheet**, builds a **summary** (overview, records by source, key
figures, breakdowns, data checks), and exports a professionally formatted
**.xlsx** workbook — borders, centered alignment, styled headers, totals,
frozen panes, auto-filter and print setup included.

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
4. **Preview & export** — check both tabs, then download the formatted workbook.
   Options include S/N and source columns, skipping total rows, de-duplication,
   sorting, a totals row (live `SUBTOTAL` formulas, so filtering recalculates),
   and one extra sheet per value of a chosen column (e.g. per status or month).

**Save settings** downloads a small JSON file; **Load settings** restores the
whole configuration next time, so a recurring report is one click.

## Build from source

```bash
npm install exceljs@4.4.0 xlsx@0.18.5
node build.js   # inlines src/ + libraries into dist/ReportBuilder.html
```

`src/index.html`, `src/styles.css` and `src/app.js` hold the code;
`build.js` produces the single distributable file.
