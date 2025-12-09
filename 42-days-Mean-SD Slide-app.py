# app_fuzzy_decimal.py
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches

st.set_page_config(page_title="Mean (SD) Table Builder ", layout="wide")

st.title("Mean (SD) Table Builder — PPT-ready ")
st.markdown("""
Upload an Excel workbook with two sheets:
- one sheet containing **means**
- another sheet containing **SDs**  

The app will match your category names **even if they are slightly different** from the standard category order,  
and you can choose **decimal places** for Mean and SD.
""")

uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.info("Please upload an Excel file that contains Mean and SD sheets.")
    st.stop()

# read sheet names
xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
st.write("Detected sheets:", sheets)
mean_sheet = st.selectbox("Select sheet that contains MEAN values", options=sheets, index=0)
sd_sheet = st.selectbox("Select sheet that contains SD values", options=sheets, index=min(1, len(sheets)-1))

# Decimal places option
decimals = st.selectbox("Select decimal places for output", options=[0, 1, 2], index=0)

@st.cache_data
def read_sheet(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name)
    return df

df_mean = read_sheet(uploaded_file, mean_sheet)
df_sd = read_sheet(uploaded_file, sd_sheet)

st.markdown("**Preview: Mean sheet (first 10 rows)**")
st.dataframe(df_mean.head(10))
st.markdown("**Preview: SD sheet (first 10 rows)**")
st.dataframe(df_sd.head(10))

def normalize_df(df):
    df = df.copy()
    idx_col = df.columns[0]
    df = df.rename(columns={idx_col: "Category"}).set_index("Category")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.apply(pd.to_numeric, errors="coerce")
    return df

df_mean_n = normalize_df(df_mean)
df_sd_n = normalize_df(df_sd)

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

def handle_non_specific(df_mean, df_sd):
    mean_df = df_mean.copy()
    sd_df = df_sd.copy()
    alt_names = ["Non-specific", "Non specific", "Non_specific"]
    for alt in alt_names:
        if alt in mean_df.index:
            if "Other rare tumors" in mean_df.index:
                mean_df = mean_df.drop(index=alt)
                sd_df = sd_df.drop(index=alt, errors='ignore')
            else:
                mean_df = mean_df.rename(index={alt: "Other rare tumors"})
                sd_df = sd_df.rename(index={alt: "Other rare tumors"})
    return mean_df, sd_df

df_mean_n, df_sd_n = handle_non_specific(df_mean_n, df_sd_n)

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

mean_map = map_columns(df_mean_n)
sd_map = map_columns(df_sd_n)

final_columns = [
    "First visit to acceptance",
    "Acceptance to first visit in OPD",
    "First Visit to MDT",
    "MDT to First Day of Therapy",
    "First visit to First Day of Therapy"
]

# Fuzzy match helper
def fuzzy_match(cat, available_list):
    match = get_close_matches(cat, available_list, n=1, cutoff=0.5)
    if match:
        return match[0]
    else:
        return None

final_df = pd.DataFrame(index=category_order, columns=final_columns)

def get_value(df, mapped, category, pretty_col):
    keys = [k for k,v in mapped.items() if v==pretty_col]
    for k in keys:
        try:
            matched_category = fuzzy_match(category, df.index.tolist())
            if matched_category:
                return df.loc[matched_category, k]
            else:
                return pd.NA
        except:
            return pd.NA
    return pd.NA

# Fill final_df with decimal formatting
for cat in category_order:
    for pretty_col in final_columns:
        mean_val = get_value(df_mean_n, mean_map, cat, pretty_col)
        sd_val = get_value(df_sd_n, sd_map, cat, pretty_col)

        if pd.isna(mean_val) and pd.isna(sd_val):
            final_df.loc[cat, pretty_col] = "–"
        else:
            # Format with selected decimals
            fmt = f"{{:.{decimals}f}}"
            mean_str = "" if pd.isna(mean_val) else fmt.format(mean_val)
            sd_str = "" if pd.isna(sd_val) else fmt.format(sd_val)
            if mean_str=="" and sd_str!="":
                final_df.loc[cat, pretty_col] = f"({sd_str})"
            elif mean_str!="" and sd_str=="":
                final_df.loc[cat, pretty_col] = f"{mean_str}"
            else:
                final_df.loc[cat, pretty_col] = f"{mean_str} ({sd_str})"

st.markdown(f"### Final Table (Mean (SD)) — {decimals} Decimal Place(s)")
st.dataframe(final_df.reset_index().rename(columns={"index":"Category"}), use_container_width=True)

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

st.success("Ready — table generated with selected decimal places.")
