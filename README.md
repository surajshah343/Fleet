# TEN Leasing — Asset List Comparison

A Streamlit app that compares two dated asset-list workbooks, cleans them, works
out what was added and removed, lets you build calculated columns, and writes the
results back as new tabs inside the newer workbook.

---

## What it does

**1 · Clean** — pairs up the sheets across the two files and cleans both:

| Problem in the source files | What the app does |
|---|---|
| Trailing padding: `"MAX ATLAS        "` | Trims and collapses repeated spaces |
| Tabs, non-breaking spaces, zero-width characters | Replaced with normal spaces |
| Placeholder nulls: `NULL`, `--`, `N/A`, `#N/A`, `" "` | Set to blank |
| Money and counts stored as text: `"1,200.50"`, `"$800"`, `"(150)"` | Converted to real numbers, brackets read as negative |
| Dates stored three ways in one column family | Converted to real dates |
| `1900-01-01` used to mean "no date" (108,652 cells in the sample) | Blanked |
| Blank spacer rows and columns | Dropped |

Every step is a checkbox, and the app reports exactly how many cells each one
touched so nothing changes silently.

**2 · Additions & Deletions** — matches assets on a key column (Serial Number by
default) and produces two tables, each keeping **all** columns:

- **New Additions** — in the newer file, absent from the older one
- **Deletions** — in the older file, gone from the newer one

Whitespace and letter case are always ignored when matching, so `"978554    "`
and `"978554"` are the same asset. Duplicate and blank keys are reported rather
than quietly skewing the counts. A **Column changes** tab shows which columns were
added, dropped or merely renamed between the two files.

**3 · Calculated columns** — Excel-style maths without leaving the app:

- **Two operands** — pick a column or a number on each side and an operation:
  add, subtract, multiply, divide, power, remainder, min, max, average,
  percent of, percent change, absolute difference, text concatenate.
- **Several columns** — sum, average, min, max, product or count across many
  columns at once.
- **Custom formula** — `([$ OLV] - [NBV_Current]) / [NBV_Current] * 100`

Queue as many operations as you like; they run in order and **each new column is
available to the ones after it**, so calculations chain the way adjacent
spreadsheet columns do. You name every output column yourself.

Available in formulas: `SUM AVG MIN MAX PRODUCT ABS ROUND SQRT LN LOG10 FLOOR
CEIL IF COALESCE CONCAT ISBLANK COUNTVALUES DAYS YEAR MONTH`, plus whole-column
aggregates `COLSUM COLAVG COLMIN COLMAX COLMEDIAN COLCOUNT` for share-of-total
style columns.

Formulas are parsed and walked against a whitelist — never `eval`'d — so a typo
produces an error message and `__import__("os")` is rejected outright.

**4 · Export** — writes the new tabs into a copy of the **newer** workbook. All
original sheets are copied across untouched: formatting, formulas and pivot
tables all survive. Re-running replaces a tab of the same name instead of piling
up duplicates.

---

## Quick start

```bash
git clone https://github.com/<your-org>/ten-asset-diff.git
cd ten-asset-diff

python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

streamlit run app.py
```

The app opens at `http://localhost:8501`. Upload both workbooks in the sidebar —
the newer one is detected from the date in the filename, and you can override it.

---

## Putting it on GitHub

```bash
git init
git add .
git commit -m "Asset list comparison app"
git branch -M main
git remote add origin https://github.com/<your-org>/ten-asset-diff.git
git push -u origin main
```

`.gitignore` already excludes `.xlsx` files, so the asset lists will not be
committed — worth keeping, as they contain customer names and contract rates.

### Deploying to Streamlit Community Cloud

1. Push the repo to GitHub (public, or private with a linked account).
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Pick the repo, branch `main`, main file `app.py`.
4. Deploy.

`.streamlit/config.toml` raises the upload limit to 500 MB, which the 50–60 MB
sample files need headroom for.

> **Before deploying anything customer-facing:** the free Community Cloud tier is
> public and has roughly 1 GB of RAM — not enough for files this size, and not a
> place for contract rates and customer names. For real use, deploy internally
> (Streamlit in Docker behind SSO, Azure App Service, or similar) with at least
> 4 GB of RAM.

---

## How it is put together

```
ten-asset-diff/
├── app.py                  Streamlit UI — the four-step workflow
├── src/
│   ├── naming.py           Column/sheet name normalisation and pairing
│   ├── loader.py           Header-row detection and streaming reads
│   ├── cleaning.py         The cleaning pipeline and its audit report
│   ├── diffing.py          Additions/deletions and column drift
│   ├── calc.py             Sandboxed formula engine
│   └── exporter.py         ZIP-level sheet appending
├── tests/test_pipeline.py  38 tests
├── requirements.txt
└── .streamlit/config.toml
```

Run the tests with:

```bash
pip install pytest
python -m pytest tests/ -v
```

### Three things worth knowing about the design

**Header rows are detected, not assumed.** In the sample files the header sits on
row 2 of `USA`, `Canada` and `US`, but on row 4 of `CAN`, under an
`Amounts in CAD$` title. Each candidate row is scored on how header-like it looks
— short, distinct, non-numeric labels — and the winner is used. You can override
it per sheet if a preview looks wrong.

**Column names are matched loosely.** The US sheets say `UnitNumber` and
`NBV_Current`; the Canadian sheets say `Unit Number` and `NBV Current`. Matching
strips spaces, underscores and case, so those line up — but every output keeps
the original spelling from its own file. Sheet names are matched the same way
(`USA`↔`US`, `Canada`↔`CAN`).

**The export never loads the workbook.** `openpyxl.load_workbook()` on the 60 MB
source file needs several gigabytes and gets killed on a normal machine. So
`exporter.py` works at the ZIP level instead: an `.xlsx` is a ZIP of XML parts, so
every existing part is streamed across untouched and only three small ones are
edited (`[Content_Types].xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`).
Peak memory stays around 350 MB and the write takes about 13 seconds.

---

## Notes and limits

- **First load is slow.** Reading four sheets of ~58,000 × 122 cells takes about
  3–4 minutes total. Results are cached, so changing a formula or a key column
  afterwards is instant. Re-running cleaning with different options re-reads only
  what changed.
- **Memory.** Budget roughly 4 GB. The cleaned frames stay in session state so
  later steps do not re-read the files.
- **Formulas need cached results.** Values are read from the cached results Excel
  stores alongside each formula. If a workbook was written by a tool that never
  cached them, those cells arrive blank — the app warns you when it detects this.
  Open the file in Excel, recalculate, and save.
- **Matching key.** Serial Number is the default and was near-unique in the
  sample (2 duplicates in 58,604 US rows). You can match on several columns at
  once, and the app reports duplicate and blank keys rather than hiding them.
- **Sheets excluded by default.** Narrow tabs and anything named like a summary
  or pivot are left out of the comparison; add them manually if you need to.

---

## Sample results

Comparing `06-19-2026` against `03-06-2026`, matching on Serial Number:

| | Newer | Older | In both | Additions | Deletions |
|---|---|---|---|---|---|
| **USA ⟷ US** | 58,604 | 57,618 | 56,184 | 2,418 | 1,434 |
| **Canada ⟷ CAN** | 23,830 | 24,450 | 23,445 | 385 | 1,005 |

Cleaning the USA sheet alone trimmed whitespace from 1,333,737 cells, blanked
484,503 placeholder values and 108,652 placeholder dates, and recovered 4 numeric
and 9 date columns that had been stored as text.
