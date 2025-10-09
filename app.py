import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO

st.set_page_config(page_title="Excel Data Cleaner", layout="centered")
st.title("📊 Excel Data Cleaner")

st.write("Choose the data source, upload one or more Excel/CSV files, clean them, and download the processed file. **Hold Ctrl (for Windows) or Cmd (for Mac) to select multiple files.**")

# Tabs for SUUMO and HOMES
tab1, tab2 = st.tabs(["SUUMO", "HOMES"])

# ---------- SUUMO Tab ----------
with tab1:
    st.header("SUUMO Data Cleaner")
    suumo_files = st.file_uploader(
        "Upload SUUMO Excel/CSV files",
        type=["xlsx", "csv"],
        key="suumo_uploader",
        accept_multiple_files=True,
        help="Select one or more files. Hold Ctrl (for Windows) or Cmd (for Mac) to select multiple files."
    )

    def clean_suumo(df: pd.DataFrame) -> pd.DataFrame:
        output_rows = []
        seen_companies = set()
    
        for _, row in df.iterrows():
            prefecture = row.get('Text', '')
            if isinstance(prefecture, str):
                prefecture = prefecture.replace('- 市区郡を選択', '').strip()
    
            # First company set
            company1 = row.get('Field1_text', '')
            if pd.notna(company1) and company1 != '' and company1 not in seen_companies:
                output_rows.append({
                    'Prefecture': prefecture,
                    'Company Name': company1,
                    'Link to Suumo Webpage': row.get('Field1_links', '') or '',
                    'TEL': row.get('Field3', '') or '',
                    'Address': row.get('Field2', '') or ''
                })
                seen_companies.add(company1)
    
            # Second company set
            company2 = row.get('Field4_text', '')
            if pd.notna(company2) and company2 != '' and company2 not in seen_companies:
                output_rows.append({
                    'Prefecture': prefecture,
                    'Company Name': company2,
                    'Link to Suumo Webpage': row.get('Field4_links', '') or '',
                    'TEL': row.get('Field6', '') or '',
                    'Address': row.get('Field5', '') or ''
                })
                seen_companies.add(company2)
    
        df_clean = pd.DataFrame(output_rows, columns=['Prefecture', 'Company Name', 'Link to Suumo Webpage', 'Address', 'TEL'])
        # Remove duplicates based on all columns
        df_clean = df_clean.drop_duplicates().reset_index(drop=True)
        return df_clean

    if suumo_files:
        dfs = []
        required_cols = ['Text', 'Field1_text', 'Field1_links', 'Field2', 'Field3', 'Field4_text', 'Field4_links', 'Field5', 'Field6']
        try:
            for suumo_file in suumo_files:
                if suumo_file.name.endswith(".csv"):
                    df = pd.read_csv(suumo_file)
                else:
                    df = pd.read_excel(suumo_file)
                # Check required columns
                if not all(col in df.columns for col in required_cols):
                    st.error(f"File '{suumo_file.name}' does not contain all required SUUMO columns: {', '.join(required_cols)}")
                    break
                dfs.append(df)
            else:  # Executes if no break occurs (all files have required columns)
                if not dfs:
                    st.error("No valid files were uploaded.")
                else:
                    df_suumo = pd.concat(dfs, ignore_index=True)
                    cleaned_suumo = clean_suumo(df_suumo)

                    # # Display prefecture row counts
                    # if not cleaned_suumo.empty:
                    #     prefecture_counts = cleaned_suumo['Prefecture'].value_counts().sort_index()
                    #     st.subheader("SUUMO Prefecture Row Counts")
                    #     for prefecture, count in prefecture_counts.items():
                    #         st.write(f"{prefecture}: {count} rows")

                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                            cleaned_suumo.to_excel(writer, index=False, sheet_name="CleanedSUUMO")
                        st.download_button(
                            label="⬇️ Download Cleaned SUUMO Excel",
                            data=buffer.getvalue(),
                            file_name="cleaned_suumo.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("No data after cleaning. Please check your input files.")
        except Exception as e:
            st.error(f"Error processing files: {str(e)}")

# ---------- HOMES Tab ----------
with tab2:
    st.header("HOMES Data Cleaner")
    homes_files = st.file_uploader(
        "Upload HOMES Excel/CSV files",
        type=["xlsx", "csv"],
        key="homes_uploader",
        accept_multiple_files=True,
        help="Select one or more files. Hold Ctrl (Windows) or Cmd (Mac) to select multiple files."
    )

    def clean_homes(df: pd.DataFrame) -> pd.DataFrame:
        """
        Cleans HOMES company data:
        - Maps input columns to required output columns
        - Removes 'map/' from Link to HOMES webpage
        - Removes duplicates
        """
        df = df.copy()
        # Rename columns to match desired output
        df = df.rename(columns={
            'Text': 'Company_Name',
            'Text1': 'Address',
            'URL': 'Homepage_URL',
            'URL1': 'Link_to_HOMES_webpage',
            'Text2': 'TEL'
        })
        # Remove trailing 'map/' from Link_to_HOMES_webpage
        df['Link_to_HOMES_webpage'] = df['Link_to_HOMES_webpage'].str.rstrip('/')
        df['Link_to_HOMES_webpage'] = df['Link_to_HOMES_webpage'].str.replace('map$', '', regex=True)
        # Remove duplicates based on all columns
        df_clean = df[['Company_Name', 'Address', 'Homepage_URL', 'Link_to_HOMES_webpage', 'TEL']].drop_duplicates().reset_index(drop=True)
        return df_clean

    if homes_files:
        dfs = []
        required_cols = ['Text', 'Text1', 'URL', 'URL1', 'Text2']
        try:
            for homes_file in homes_files:
                if homes_file.name.endswith(".csv"):
                    df = pd.read_csv(homes_file)
                else:
                    df = pd.read_excel(homes_file, engine="openpyxl")
                # Check required columns
                if not all(col in df.columns for col in required_cols):
                    st.error(f"File '{homes_file.name}' does not contain all required HOMES columns: {', '.join(required_cols)}")
                    break
                dfs.append(df)
            else:  # Executes if no break occurs (all files have required columns)
                if not dfs:
                    st.error("No valid files were uploaded.")
                else:
                    df_homes = pd.concat(dfs, ignore_index=True)
                    cleaned_homes = clean_homes(df_homes)

                    # # Display prefecture row counts
                    # if not cleaned_homes.empty:
                    #     cleaned_homes['Prefecture'] = cleaned_homes['Address'].str.extract(r'^([^都道府県]+[都道府県])')
                    #     prefecture_counts = cleaned_homes['Prefecture'].value_counts().sort_index()
                    #     st.subheader("HOMES Prefecture Row Counts")
                    #     for prefecture, count in prefecture_counts.items():
                    #         st.write(f"{prefecture}: {count} rows")

                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                            cleaned_homes.to_excel(writer, index=False, sheet_name="CleanedHOMES")
                        st.download_button(
                            label="⬇️ Download Cleaned HOMES Excel",
                            data=buffer.getvalue(),
                            file_name="cleaned_homes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("No data after cleaning. Please check your input files.")
        except Exception as e:
            st.error(f"Error processing files: {str(e)}")
