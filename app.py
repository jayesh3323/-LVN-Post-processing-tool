import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Cleaner", layout="centered")

st.title("📊 Excel Data Cleaner")

st.write("Upload your Excel file, clean it, and download the processed version.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "csv"])

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Placeholder cleaning function.
    Replace this with your actual cleaning rules.
    """
    # Example: Drop completely empty rows
    df = df.dropna(how="all")
    return df

if uploaded_file is not None:
    # Detect file type
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    # Apply cleaning
    cleaned_df = clean_data(df)

    st.subheader("Preview of Cleaned Data")
    st.dataframe(cleaned_df.head())

    # Download button
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        cleaned_df.to_excel(writer, index=False, sheet_name="CleanedData")
    st.download_button(
        label="⬇️ Download Cleaned Excel",
        data=buffer.getvalue(),
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
