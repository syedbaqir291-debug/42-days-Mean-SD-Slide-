# app_fuzzy_multi_mean.py
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches

st.set_page_config(page_title="Mean (SD) / Median(Range) Table Builder (Multi-sheet)", layout="wide")

st.title("Mean (SD) / Median(Range) Table Builder (Multi-sheet)")
st.markdown("""
Upload an Excel workbook with multiple sheets.  
- You can select **one or more sheets** for outside values.  
- Similarly, select **one or more sheets** for inside values.  
The app will combine values **cell-wise with dash `-`** in the order you select them.  
Fuzzy category matching and decimal formatting are supported.
""")

uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.info("Please upload an Excel file that contains Mean,SD,Range, Median sheets.")
    st.stop()

xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
st.write("Detected sheets:", sheets)

# MULTI-select widgets
mean_sheets = st.multiselect("Select sheet(s) that contain outer values", options=sheets)
sd_sheets = st.multiselect("Select sheet(s) that contain inner values", options=sheets)

decimals = st.selectbox("Select decimal places for output", options=[0,1,2], index=0)

if not mean_sheets or not sd_sheets:
    st.warning("Please select at least one sheet for both Mean and SD.")
    st.stop()

# Read multiple sheets and concatenate cell-wise
def read_multi_sheets(file, sheet_list):
    dfs = []
    for sheet in sheet_list:
        df = pd.read_excel(file, sheet_name=sheet)
        df = df.rename(columns={df.columns[0]:"Category"}).set_index("Category")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.apply(pd.to_numeric, errors='coerce')
        dfs.append(df)
    return dfs

mean_dfs = read_multi_sheets(uploaded_file, mean_sheets)
sd_dfs = read_multi_sheets(uploaded_file, sd_sheets)

# Category order
category_order = [
    "Haematological",
    "Gynecological",
    "Urological",
    "Neurological",
    "Breast",
    "Pulmonary",
    "Gastrointestinal",
    "Head & Neck",
    "Thyroid",
    "Sarcoma",
    "Retinoblastoma",
    "Other rare tumors"
]

final_columns = [
    "First visit to acceptance",
    "Acceptance to first visit in OPD",
    "First Visit to MDT",
    "MDT to First Day of Therapy",
    "First visit to First Day of Therapy"
]

# Handle Non-specific -> Other rare tumors in all sheets
def fix_non_specific(dfs):
    fixed_dfs = []
    for df in dfs:
        df = df.copy()
        alt_names = ["Non-specific", "Non specific", "Non_specific"]
        for alt in alt_names:
            if alt in df.index:
                if "Other rare tumors" in df.index:
                    df = df.drop(index=alt)
                else:
                    df = df.rename(index={alt: "Other rare tumors"})
        fixed_dfs.append(df)
    return fixed_dfs

mean_dfs = fix_non_specific(mean_dfs)
sd_dfs = fix_non_specific(sd_dfs)

# Column mapping
expected_cols = {
    "FIRST_VISIT_TO_ACCEPT": "First visit to acceptance",
    "ACCEPT_TO_FIRST_CONSULTANT_NOT": "Acceptance to first visit in OPD",
    "CONSULTANT_NOTE_TO_MDT": "First Visit to MDT",
    "DAYS_BTW_MDT_TO_1ST_THERAPY": "MDT to First Day of Therapy",
    "FIRST_NOTE_TO_THERAPY": "First visit to First Day of Therapy"
}

def map_columns(df):
    mapped = {}
    for col in df.columns:
        col_up = str(col).strip().upper()
        if col_up in expected_cols:
            mapped[col] = expected_cols[col_up]
    return mapped

mean_maps = [map_columns(df) for df in mean_dfs]
sd_maps = [map_columns(df) for df in sd_dfs]

# Fuzzy matching helper
def fuzzy_match(cat, available_list):
    match = get_close_matches(cat, available_list, n=1, cutoff=0.5)
    if match:
        return match[0]
    else:
        return None

# Combine multiple sheets cell-wise with dash
def get_multi_values(dfs, maps, category, pretty_col):
    values = []
    for df, m in zip(dfs, maps):
        keys = [k for k,v in m.items() if v==pretty_col]
        val = pd.NA
        for k in keys:
            matched_cat = fuzzy_match(category, df.index.tolist())
            if matched_cat:
                val = df.loc[matched_cat, k]
        if pd.isna(val):
            values.append("")
        else:
            values.append(val)
    return "-".join(str(int(v)) if isinstance(v,(int,float)) and float(v).is_integer() else f"{v}" for v in values)

# Prepare final DataFrame
final_df = pd.DataFrame(index=category_order, columns=final_columns)

for cat in category_order:
    for pretty_col in final_columns:
        mean_val = get_multi_values(mean_dfs, mean_maps, cat, pretty_col)
        sd_val = get_multi_values(sd_dfs, sd_maps, cat, pretty_col)
        
        # Format decimals
        if mean_val != "":
            mean_vals = [float(v) if v!="" else pd.NA for v in mean_val.split("-")]
            mean_vals_fmt = [f"{v:.{decimals}f}" if not pd.isna(v) else "" for v in mean_vals]
            mean_val = "-".join(mean_vals_fmt)
        if sd_val != "":
            sd_vals = [float(v) if v!="" else pd.NA for v in sd_val.split("-")]
            sd_vals_fmt = [f"{v:.{decimals}f}" if not pd.isna(v) else "" for v in sd_vals]
            sd_val = "-".join(sd_vals_fmt)
        
        if mean_val=="" and sd_val=="":
            final_df.loc[cat, pretty_col] = "–"
        elif mean_val!="" and sd_val=="":
            final_df.loc[cat, pretty_col] = mean_val
        elif mean_val=="" and sd_val!="":
            final_df.loc[cat, pretty_col] = f"({sd_val})"
        else:
            final_df.loc[cat, pretty_col] = f"{mean_val} ({sd_val})"

st.markdown(f"### Final Table (Mean (SD)) — {decimals} Decimal Place(s)")
st.dataframe(final_df.reset_index().rename(columns={"index":"Category"}), use_container_width=True)

# Excel download
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mean_SD_Table")
    return output.getvalue()

excel_bytes = to_excel_bytes(final_df.reset_index().rename(columns={"index":"Category"}))
st.download_button(
    label="Download final table as Excel",
    data=excel_bytes,
    file_name="Mean_SD_Table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Ready — Multi-sheet dash combined table generated!")
