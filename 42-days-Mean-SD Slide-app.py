# app.py
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Mean (SD) Table Builder", layout="wide")

st.title("Mean (SD) Table Builder — PPT / Excel ready")
st.markdown("""
Upload an Excel workbook with two sheets:
- one sheet containing **means** (first column with category names),
- another sheet containing **SDs** with same layout.

Select the sheets for Mean and SD, then the app will produce a table in the requested category order
with each cell formatted as: **Mean (SD)**.
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

# read both sheets into DataFrames
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

# Helper to find category column (user used header 'Sheet' or first column)
def normalize_df(df):
    df = df.copy()
    # If there's a column explicitly named like 'Sheet' or 'Category' use it as index
    possible_idx = [c for c in df.columns if str(c).strip().lower() in ("sheet", "category", "categories", "index")]
    if possible_idx:
        idx_col = possible_idx[0]
    else:
        # else take the first column as category names
        idx_col = df.columns[0]
    df = df.rename(columns={idx_col: "Category"}).set_index("Category")
    # strip whitespace from column names
    df.columns = [str(c).strip() for c in df.columns]
    # ensure numeric where possible
    df = df.apply(pd.to_numeric, errors="coerce")
    return df

df_mean_n = normalize_df(df_mean)
df_sd_n = normalize_df(df_sd)

# category order (exact order requested by user)
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

# handle 'Non-specific' mapping to 'Other rare tumors'
def handle_non_specific(df_mean, df_sd):
    mean_df = df_mean.copy()
    sd_df = df_sd.copy()

    # possible alternative names that appear in user's file
    alt_names = ["Non-specific", "Non specific", "Non_specific"]
    for alt in alt_names:
        if alt in mean_df.index:
            if "Other rare tumors" in mean_df.index:
                # If both exist, prefer 'Other rare tumors' and drop Non-specific (but warn)
                mean_df = mean_df.drop(index=alt)
                sd_df = sd_df.drop(index=alt, errors='ignore')
                st.warning(f"Both '{alt}' and 'Other rare tumors' present in Mean sheet. Dropped '{alt}' (kept 'Other rare tumors').")
            else:
                # rename Non-specific -> Other rare tumors
                mean_df = mean_df.rename(index={alt: "Other rare tumors"})
                sd_df = sd_df.rename(index={alt: "Other rare tumors"})
                st.info(f"Renamed '{alt}' to 'Other rare tumors'.")
    return mean_df, sd_df

df_mean_n, df_sd_n = handle_non_specific(df_mean_n, df_sd_n)

# Build final table with desired columns mapping (user columns)
# Expected column names (tolerant to minor variants)
expected_cols = {
    "FIRST_VISIT_TO_ACCEPT": "First visit to acceptance",
    "FIRST_VISIT_TO_ACCEPT ": "First visit to acceptance",
    "ACCEPT_TO_FIRST_CONSULTANT_NOT": "Acceptance to first visit in OPD",
    "ACCEPT_TO_FIRST_CONSULTANT_NOTE": "Acceptance to first visit in OPD",
    "ACCEPT_TO_FIRST_CONSULTANT_NOT ": "Acceptance to first visit in OPD",
    "CONSULTANT_NOTE_TO_MDT": "First Visit to MDT",
    "CONSULTANT_NOTE_TO_MDT ": "First Visit to MDT",
    "DAYS_BTW_MDT_TO_1ST_THERAPY": "MDT to First Day of Therapy",
    "DAYS_BTW_MDT_TO_1ST_THERAPY ": "MDT to First Day of Therapy",
    "FIRST_NOTE_TO_THERAPY": "First visit to First Day of Therapy",
    "FIRST_NOTE_TO_THERAPY ": "First visit to First Day of Therapy"
}

# map available mean columns to pretty headers
def map_columns(df):
    mapped = {}
    for col in df.columns:
        col_up = str(col).strip().upper()
        if col_up in expected_cols:
            mapped[col] = expected_cols[col_up]
    return mapped

mean_map = map_columns(df_mean_n)
sd_map = map_columns(df_sd_n)

# use union of mapped columns
all_mapped = list(dict.fromkeys(list(mean_map.values()) + list(sd_map.values())))

# If some expected columns missing, still include them but they will be blanks
final_columns = [
    "First visit to acceptance",
    "Acceptance to first visit in OPD",
    "First Visit to MDT",
    "MDT to First Day of Therapy",
    "First visit to First Day of Therapy"
]

# prepare a DataFrame for final output
final_df = pd.DataFrame(index=category_order, columns=final_columns)

# helper: get value for category & pretty column
def get_value(df, mapped, category, pretty_col):
    # find key in mapped where mapped[key] == pretty_col
    keys = [k for k,v in mapped.items() if v==pretty_col]
    for k in keys:
        try:
            return df.loc[category, k]
        except Exception:
            # maybe category not present
            return pd.NA
    return pd.NA

# fill final_df with Mean (SD) formatted strings
for cat in category_order:
    for pretty_col in final_columns:
        mean_val = get_value(df_mean_n, mean_map, cat, pretty_col)
        sd_val = get_value(df_sd_n, sd_map, cat, pretty_col)

        if pd.isna(mean_val) and pd.isna(sd_val):
            final_df.loc[cat, pretty_col] = "–"
        else:
            # handle NaN: replace with blank
            mean_str = "" if pd.isna(mean_val) else (str(int(mean_val)) if float(mean_val).is_integer() else f"{mean_val}")
            sd_str = "" if pd.isna(sd_val) else (str(int(sd_val)) if float(sd_val).is_integer() else f"{sd_val}")
            if mean_str=="" and sd_str!="":
                final_df.loc[cat, pretty_col] = f"({sd_str})"
            elif mean_str!="" and sd_str=="":
                final_df.loc[cat, pretty_col] = f"{mean_str}"
            else:
                final_df.loc[cat, pretty_col] = f"{mean_str} ({sd_str})"

st.markdown("### Final Table (Mean (SD)) — PPT-ready")
st.dataframe(final_df.reset_index().rename(columns={"index":"Category"}), use_container_width=True)

# Provide download button for Excel
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mean_SD_Table")
        writer.save()
    processed_data = output.getvalue()
    return processed_data

excel_bytes = to_excel_bytes(final_df.reset_index().rename(columns={"index":"Category"}))
st.download_button(
    label="Download final table as Excel",
    data=excel_bytes,
    file_name="Mean_SD_Table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Ready — copy this app to your GitHub repo and deploy on Streamlit Sharing / Cloud.")
st.markdown("""
**Notes / Tips**
- Make sure Mean and SD sheets use the same category names (or include 'Non-specific' if you want it treated as 'Other rare tumors').
- The app is tolerant to slight column-name variants but expects the key columns that were listed earlier (FIRST_VISIT_TO_ACCEPT, ACCEPT_TO_FIRST_CONSULTANT_NOT, CONSULTANT_NOTE_TO_MDT, DAYS_BTW_MDT_TO_1ST_THERAPY, FIRST_NOTE_TO_THERAPY).
- If you want help pushing this to a GitHub repo or making it prettier (colors / Excel formatting / direct PPT generation), tell me and I will provide the exact extra code.
""")
