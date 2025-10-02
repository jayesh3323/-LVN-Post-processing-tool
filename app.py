import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Data Cleaner", layout="centered")
st.title("📊 Excel Data Cleaner")

st.write("Choose the data source, upload your Excel file, clean it, and download the processed file.")

# Tabs for SUUMO and HOMES
tab1, tab2 = st.tabs(["SUUMO", "HOMES"])

# ---------- SUUMO Tab ----------
with tab1:
    st.header("SUUMO Data Cleaner")
    suumo_file = st.file_uploader("Upload SUUMO Excel file", type=["xlsx", "csv"], key="suumo_uploader")
    
    def clean_suumo(df: pd.DataFrame) -> pd.DataFrame:
        """Placeholder SUUMO cleaning logic."""
        df = df.dropna(how="all")
        return df

    if suumo_file is not None:
        if suumo_file.name.endswith(".csv"):
            df_suumo = pd.read_csv(suumo_file)
        else:
            df_suumo = pd.read_excel(suumo_file)
        
        st.subheader("Preview of Uploaded SUUMO Data")
        st.dataframe(df_suumo.head())

        cleaned_suumo = clean_suumo(df_suumo)

        st.subheader("Preview of Cleaned SUUMO Data")
        st.dataframe(cleaned_suumo.head())

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            cleaned_suumo.to_excel(writer, index=False, sheet_name="CleanedSUUMO")
        st.download_button(
            label="⬇️ Download Cleaned SUUMO Excel",
            data=buffer.getvalue(),
            file_name="cleaned_suumo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------- HOMES Tab ----------
with tab2:
    st.header("HOMES Data Cleaner")
    homes_file = st.file_uploader("Upload HOMES Excel file", type=["xlsx", "csv"], key="homes_uploader")
    
    def clean_homes(df: pd.DataFrame) -> pd.DataFrame:
        """
        HOMES cleaning logic: reshape long format to wide format.
        Each company becomes one row with fields as columns.
        """
        # Fill empty repeated values for company name / URLs
        df['Company_Name'] = df['Company_Name'].fillna(method='ffill')
        df['Link_to_the_Homepage'] = df['Link_to_the_Homepage'].fillna(method='ffill')
        df['URL'] = df['URL'].fillna(method='ffill')

        # Pivot: Field1 becomes columns, Field2 becomes values
        df_wide = df.pivot_table(
            index=['Company_Name', 'Link_to_the_Homepage', 'URL'],
            columns='Field1',
            values='Field2',
            aggfunc=lambda x: ' '.join(str(v) for v in x)
        ).reset_index()

        # Reorder / select desired columns
        desired_columns = [
            'Company_Name', 'Link_to_the_Homepage', 'URL',
            '所在地', '交通', '営業時間', '定休日',
            'TEL', 'FAX', '免許番号', '所属団体名', '保証協会', '屋号'
        ]
        df_wide = df_wide[[c for c in desired_columns if c in df_wide.columns]]

        return df_wide

    if homes_file is not None:
        if homes_file.name.endswith(".csv"):
            df_homes = pd.read_csv(homes_file)
        else:
            df_homes = pd.read_excel(homes_file)
        
        st.subheader("Preview of Uploaded HOMES Data")
        st.dataframe(df_homes.head())

        cleaned_homes = clean_homes(df_homes)

        st.subheader("Preview of Cleaned HOMES Data")
        st.dataframe(cleaned_homes.head())

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            cleaned_homes.to_excel(writer, index=False, sheet_name="CleanedHOMES")
        st.download_button(
            label="⬇️ Download Cleaned HOMES Excel",
            data=buffer.getvalue(),
            file_name="cleaned_homes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
