"""
mmf_core.py
-----------
Core data-transform logic for the MMF Image Stack Reviewer.

This module is a Python port of the original VBA macro
`MMF_ImageLinks_By_SKU_Sequence` (Designed and Developed by Yash Kende),
extended to support PXM "Media Stack" exports that contain more than one
marketplace in a single file (e.g. Amazon US, Amazon CA, Walmart US,
Walmart CA, Best Buy US, Kohl's US, Lowes US, ...).

The original macro read four columns (SKU, blank, blank, Sequence, Image
Link) from the active sheet, grouped rows by SKU, sorted each group by
Sequence, and pivoted the result into:

    Master Id/ Sku | Main Image | Other Image 1 | Other Image 2 | ...

This module keeps that exact grouping/sorting/pivoting behavior but adds:
    1. Flexible column name detection (PXM exports use column headers like
       "Collection Folder", "Media Stack Group", "Media Stack Order",
       "Media", "Preview", "Filename" rather than bare A/B/C/D/E).
    2. A marketplace dimension ("Media Stack Group") so a single workbook
       covering many marketplaces can be filtered down to one marketplace
       at a time before the SKU/Sequence pivot is built.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from typing import Optional

import pandas as pd

TOOL_TITLE = "Designed and Developed by Yash Kende"

# Candidate header names (case-insensitive, whitespace-insensitive) for each
# logical field we need. The first matching column found in the uploaded
# file is used. This keeps the app working across slightly different PXM
# export configurations without requiring the user to rename columns.
_COLUMN_CANDIDATES = {
    "sku": ["collection folder", "sku", "master id/ sku", "master id", "product sku"],
    "marketplace": ["media stack group", "marketplace", "channel", "media stack"],
    "sequence": ["media stack order", "sequence", "seq", "image sequence", "order"],
    "image": ["media", "image link", "image url", "full image", "image"],
    "preview": ["preview", "thumbnail", "thumb", "preview url"],
    "filename": ["filename", "file name", "image name"],
}


def _normalize(name: object) -> str:
    return str(name).strip().lower()


def _detect_columns(df: pd.DataFrame) -> dict:
    """Map logical field names -> actual column names present in df."""
    normalized = {_normalize(c): c for c in df.columns}
    detected: dict[str, Optional[str]] = {}

    for field_name, candidates in _COLUMN_CANDIDATES.items():
        found = None
        for candidate in candidates:
            if candidate in normalized:
                found = normalized[candidate]
                break
        detected[field_name] = found

    missing_required = [f for f in ("sku", "sequence", "image") if detected[f] is None]
    if missing_required:
        raise ValueError(
            "Could not find required column(s) in the uploaded file: "
            f"{', '.join(missing_required)}. Expected headers similar to "
            "'Collection Folder' (SKU), 'Media Stack Order' (Sequence) and "
            "'Media' (Image Link)."
        )
    return detected


def load_media_stack(file_or_buffer, filename: Optional[str] = None) -> pd.DataFrame:
    """Load a PXM media stack export (xlsx/xls/csv) into a normalized DataFrame.

    Returns a DataFrame with standardized columns:
        sku, marketplace, sequence, image, preview, filename
    `marketplace`, `preview` and `filename` may be all-empty if the source
    file doesn't contain those columns (e.g. a single-marketplace export).
    """
    name = (filename or getattr(file_or_buffer, "name", "") or "").lower()

    if name.endswith(".csv"):
        raw = pd.read_csv(file_or_buffer, dtype=str)
    else:
        raw = pd.read_excel(file_or_buffer, dtype=str)

    raw.columns = [str(c).strip() for c in raw.columns]
    cols = _detect_columns(raw)

    out = pd.DataFrame()
    out["sku"] = raw[cols["sku"]].astype(str).str.strip()
    out["sequence"] = pd.to_numeric(raw[cols["sequence"]], errors="coerce")
    out["image"] = raw[cols["image"]].astype(str).str.strip()
    out["marketplace"] = (
        raw[cols["marketplace"]].astype(str).str.strip() if cols["marketplace"] else ""
    )
    out["preview"] = raw[cols["preview"]].astype(str).str.strip() if cols["preview"] else ""
    out["filename"] = raw[cols["filename"]].astype(str).str.strip() if cols["filename"] else ""

    # Mirror the VBA validity check: SKU non-empty, Image non-empty, Sequence numeric.
    valid = (
        out["sku"].notna()
        & (out["sku"] != "")
        & (out["sku"].str.lower() != "nan")
        & out["image"].notna()
        & (out["image"] != "")
        & (out["image"].str.lower() != "nan")
        & out["sequence"].notna()
    )
    out = out[valid].reset_index(drop=True)

    # Blank out placeholder "nan" strings produced by astype(str) on empty cells.
    for col in ("marketplace", "preview", "filename"):
        out[col] = out[col].replace({"nan": "", "None": ""})

    return out


def has_marketplace_dimension(df: pd.DataFrame) -> bool:
    return bool(df["marketplace"].str.strip().replace("", pd.NA).notna().any())


def get_marketplaces(df: pd.DataFrame) -> list[str]:
    """Return the sorted list of distinct, non-blank marketplace values."""
    vals = sorted(v for v in df["marketplace"].unique() if v)
    return vals


def filter_by_marketplace(df: pd.DataFrame, marketplace: Optional[str]) -> pd.DataFrame:
    if not marketplace or not has_marketplace_dimension(df):
        return df
    return df[df["marketplace"] == marketplace].reset_index(drop=True)


@dataclass
class SkuImages:
    sku: str
    images: list[dict] = field(default_factory=list)  # each: sequence, image, preview, filename


def group_by_sku(df: pd.DataFrame) -> list[SkuImages]:
    """Sort by SKU asc, Sequence asc (same order as the VBA Range.Sort call)
    and group rows into one ordered image list per SKU."""
    ordered = df.sort_values(by=["sku", "sequence"], kind="stable")

    groups: dict[str, SkuImages] = {}
    for _, row in ordered.iterrows():
        sku = row["sku"]
        if sku not in groups:
            groups[sku] = SkuImages(sku=sku)
        groups[sku].images.append(
            {
                "sequence": row["sequence"],
                "image": row["image"],
                "preview": row["preview"] or row["image"],
                "filename": row["filename"],
            }
        )
    # dict preserves insertion order (Python 3.7+), and rows arrive SKU-sorted,
    # so iterating groups.values() yields SKUs in ascending order too.
    return list(groups.values())


def build_wide_table(groups: list[SkuImages]) -> pd.DataFrame:
    """Pivot grouped SKU images into the same wide layout the VBA macro
    produced: Master Id/ Sku | Main Image | Other Image 1 | Other Image 2 | ..."""
    if not groups:
        return pd.DataFrame(columns=["Master Id/ Sku", "Main Image"])

    max_images = max(len(g.images) for g in groups)
    columns = ["Master Id/ Sku", "Main Image"] + [
        f"Other Image {i}" for i in range(1, max_images)
    ]

    rows = []
    for g in groups:
        row = [g.sku] + [img["image"] for img in g.images]
        row += [""] * (len(columns) - len(row))
        rows.append(row)

    return pd.DataFrame(rows, columns=columns)


def to_excel_bytes(wide_df: pd.DataFrame, marketplace: Optional[str] = None) -> bytes:
    """Serialize the wide image-links table to an .xlsx file (in memory),
    matching the VBA output sheet name/shape, plus a small About sheet."""
    buffer = BytesIO()
    sheet_name = "output_image_links"
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        wide_df.to_excel(writer, sheet_name=sheet_name, index=False)
        about = pd.DataFrame(
            {
                "Info": [
                    TOOL_TITLE,
                    f"Marketplace: {marketplace}" if marketplace else "Marketplace: (all)",
                    f"SKUs: {len(wide_df)}",
                ]
            }
        )
        about.to_excel(writer, sheet_name="About", index=False)

        ws = writer.sheets[sheet_name]
        for col_cells in ws.columns:
            length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 60)

    return buffer.getvalue()


def build_grouped_and_table(df: pd.DataFrame) -> tuple[list[SkuImages], pd.DataFrame]:
    groups = group_by_sku(df)
    table = build_wide_table(groups)
    return groups, table
