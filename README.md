# MMF Image Stack Reviewer

*Designed and Developed by Yash Kende*

A Streamlit port/extension of the `MMF_ImageLinks_By_SKU_Sequence` VBA macro.
Where the original macro grouped one marketplace's worth of SKU image rows
into a `Master Id/ Sku | Main Image | Other Image 1 | ...` table inside
Excel, this app is built for PXM **Media Stack** exports that bundle
**multiple marketplaces in one file** (Amazon US, Amazon CA, Walmart US,
Walmart CA, Best Buy US, Kohl's US, Lowe's US, etc.), and lets you review
the actual images in a browser instead of just a list of links.

## What it does

1. **Upload** the media stack file exported from PXM (`.xlsx`, `.xls`, or
   `.csv`).
2. **Pick a marketplace** from a dropdown built from the file's
   `Media Stack Group` column (falls back gracefully if the file only has
   one marketplace / no such column).
3. **Review** every SKU's images, in the same sequence order PXM assigned,
   rendered directly on the page (first image per SKU is labeled
   `Main Image`, the rest `Seq N`), each with a link to the full-resolution
   file.
4. **Download** the same wide `Master Id/ Sku | Main Image | Other Image 1 | ...`
   Excel table the original macro produced, scoped to the marketplace you
   selected.

## Input file format

The app auto-detects the following columns by header name
(case-insensitive), matching a typical PXM Media Stack export:

| Logical field | Expected header(s) |
|---|---|
| SKU | `Collection Folder`, `SKU`, `Master Id/ Sku` |
| Marketplace | `Media Stack Group`, `Marketplace`, `Channel` |
| Sequence | `Media Stack Order`, `Sequence` |
| Image URL | `Media`, `Image Link`, `Image URL` |
| Thumbnail (optional) | `Preview`, `Thumbnail` |
| Filename (optional) | `Filename` |

A row is only included if the SKU and Image URL are non-blank and the
Sequence value is numeric — the same validity rule the VBA macro used.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Then open the local URL Streamlit prints (defaults to
`http://localhost:8501`).

## Project structure

```
app.py          Streamlit UI: upload -> marketplace select -> review -> download
utils.py        Core transform logic (column detection, filtering, grouping,
                pivoting, Excel export) — the Python equivalent of the VBA
                macro's logic, unit-testable independent of Streamlit
tests/          Unit tests for utils.py
requirements.txt
```

## Notes / differences from the VBA macro

- The VBA macro operated on the *active worksheet* of the *active workbook*
  and wrote its output back into a new sheet in that same workbook. This
  app instead accepts a file upload and offers the output table as a
  downloadable `.xlsx`, since Streamlit apps don't have an "active Excel
  workbook" to write back into.
- The VBA macro only ever saw one marketplace at a time (whatever was on
  the active sheet). This app adds the marketplace dropdown so one
  multi-marketplace PXM export can be split and reviewed per marketplace
  without manually pre-splitting the file first.
- Sorting/grouping/validity rules (SKU non-blank, Image Link non-blank,
  Sequence numeric, sort by SKU then Sequence) are unchanged from the
  original macro.
