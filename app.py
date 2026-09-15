"""
app.py
------
MMF Image Stack Reviewer (Streamlit)

Designed and Developed by Yash Kende

Workflow:
    1. Download a "Media Stack" export from PXM (covers one or more
       marketplaces: Amazon US, Amazon CA, Walmart US, Walmart CA,
       Best Buy US, Kohl's US, Lowe's US, etc).
    2. Upload that file below.
    3. Pick a marketplace.
    4. Review every SKU's images (Main Image + Other Images, in the same
       sequence order PXM assigned) directly on the page.
    5. Optionally download the same wide SKU -> image-link table the
       original MMF_ImageLinks_By_SKU_Sequence VBA macro produced, now
       scoped to the marketplace you picked.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from utils import (
    TOOL_TITLE,
    build_grouped_and_table,
    filter_by_marketplace,
    get_marketplaces,
    has_marketplace_dimension,
    load_media_stack,
    to_excel_bytes,
)

st.set_page_config(page_title="MMF Image Stack Reviewer", layout="wide")

st.title("MMF Image Stack Reviewer")
st.caption(TOOL_TITLE)

st.markdown(
    "Upload a PXM **Media Stack** export, choose a marketplace, then review "
    "every SKU's images in sequence order before they go out the door."
)

# ---------------------------------------------------------------------------
# Step 1 — Upload
# ---------------------------------------------------------------------------
st.header("1. Upload media stack file")
uploaded = st.file_uploader(
    "Media stack export (.xlsx, .xls or .csv)",
    type=["xlsx", "xls", "csv"],
)

if uploaded is None:
    st.info("Upload a media stack file to get started.")
    st.stop()


@st.cache_data(show_spinner="Reading media stack file...")
def _load(file_bytes: bytes, filename: str) -> pd.DataFrame:
    from io import BytesIO

    return load_media_stack(BytesIO(file_bytes), filename=filename)


try:
    file_bytes = uploaded.getvalue()
    df = _load(file_bytes, uploaded.name)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

if df.empty:
    st.warning("No valid SKU / Sequence / Image Link rows were found in this file.")
    st.stop()

st.success(f"Loaded {len(df)} valid image rows from **{uploaded.name}**.")

# ---------------------------------------------------------------------------
# Step 2 — Marketplace selection
# ---------------------------------------------------------------------------
st.header("2. Select marketplace")

if has_marketplace_dimension(df):
    marketplaces = get_marketplaces(df)
    marketplace = st.selectbox("Marketplace", options=marketplaces, index=0)
    scoped_df = filter_by_marketplace(df, marketplace)
else:
    st.info(
        "No marketplace column detected in this file — treating it as a "
        "single marketplace."
    )
    marketplace = None
    scoped_df = df

if scoped_df.empty:
    st.warning("No rows for this marketplace.")
    st.stop()

groups, wide_table = build_grouped_and_table(scoped_df)

# ---------------------------------------------------------------------------
# Step 3 — Summary + download
# ---------------------------------------------------------------------------
st.header("3. Review")

col1, col2, col3 = st.columns(3)
col1.metric("Marketplace", marketplace or "All")
col2.metric("SKUs", len(groups))
col3.metric("Total images", sum(len(g.images) for g in groups))

excel_bytes = to_excel_bytes(wide_table, marketplace)
st.download_button(
    "Download image-links Excel (Main Image + Other Images)",
    data=excel_bytes,
    file_name=f"output_image_links_{(marketplace or 'all').replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Preview output table"):
    st.dataframe(wide_table, use_container_width=True)

st.divider()

# ---------------------------------------------------------------------------
# Step 4 — SKU search + image grid
# ---------------------------------------------------------------------------
search = st.text_input("Filter by SKU (optional)", "").strip().lower()
visible_groups = [g for g in groups if search in g.sku.lower()] if search else groups

st.caption(f"Showing {len(visible_groups)} of {len(groups)} SKUs")

IMAGES_PER_ROW = 5

for g in visible_groups:
    st.subheader(g.sku)
    for row_start in range(0, len(g.images), IMAGES_PER_ROW):
        row_images = g.images[row_start : row_start + IMAGES_PER_ROW]
        cols = st.columns(IMAGES_PER_ROW)
        for col, img in zip(cols, row_images):
            label = "Main Image" if img["sequence"] == g.images[0]["sequence"] else f"Seq {int(img['sequence'])}"
            with col:
                st.image(img["preview"], caption=f"{label}", use_container_width=True)
                st.markdown(
                    f"[Open full image]({img['image']})"
                    + (f"  \n`{img['filename']}`" if img["filename"] else "")
                )
    st.divider()
