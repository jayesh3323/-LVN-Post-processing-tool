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
        output_rows = []
        seen_companies = set()
    
        for _, row in df.iterrows():
            prefecture = row['Text']
    
            # First company set
            company1 = str(row.get('Field1_text', '')).replace('市区郡を選択', '').strip()
            link1 = row.get('Field1_links', '')
            tel1 = row.get('Field2', '')
            address1 = row.get('Field3', '')
    
            if company1 and company1 not in seen_companies:
                output_rows.append({
                    'Prefecture': prefecture,
                    'Company Name': company1,
                    'Link to Suumo Webpage': link1,
                    'Address': address1,
                    'TEL': tel1
                })
                seen_companies.add(company1)
    
            # Second company set
            company2 = str(row.get('Field4_text', '')).replace('市区郡を選択', '').strip()
            link2 = row.get('Field4_links', '')
            tel2 = row.get('Field5', '')
            address2 = row.get('Field6', '')
    
            if company2 and company2 not in seen_companies:
                output_rows.append({
                    'Prefecture': prefecture,
                    'Company Name': company2,
                    'Link to Suumo Webpage': link2,
                    'Address': address2,
                    'TEL': tel2
                })
                seen_companies.add(company2)
    
        df_clean = pd.DataFrame(output_rows, columns=['Prefecture', 'Company Name', 'Link to Suumo Webpage', 'Address', 'TEL'])
        return df_clean

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
        Cleans HOMES company data:
        - Preserve first occurrence order
        - Convert company name to kanji
        - Remove duplicates and unwanted text
        """
        df = df.copy()
        # Fill repeated URLs if missing
        df['Link_to_the_Homepage'] = df['Link_to_the_Homepage'].fillna(method='ffill')
        df['URL'] = df['URL'].fillna(method='ffill')

        processed = []
        seen = {}  # preserve order

        for _, row in df.iterrows():
            company_key = row['Company_Name']
            if company_key not in seen:
                company_dict = {
                    'Company_Name': None,  # kanji
                    'Link_to_the_Homepage': row['Link_to_the_Homepage'],
                    'HOMES Webpage URL': row['URL'],
                    '所在地': '',
                    '交通': '',
                    '営業時間': '',
                    '定休日': '',
                    'TEL': '',
                    'FAX': '',
                    '免許番号': '',
                    '所属団体名': '',
                    '保証協会': '',
                    '屋号': ''
                }
                processed.append(company_dict)
                seen[company_key] = company_dict

            company_dict = seen[company_key]

            if row['Field1'] == '会社名':
                company_dict['Company_Name'] = row['Field2']
            else:
                if row['Field1'] in company_dict:
                    if company_dict[row['Field1']]:
                        company_dict[row['Field1']] += ' ' + str(row['Field2'])
                    else:
                        company_dict[row['Field1']] = row['Field2']

        df_wide = pd.DataFrame(processed)
        # Remove trailing 'map/' from HOMES Webpage URL
        df_wide['HOMES Webpage URL'] = df_wide['HOMES Webpage URL'].str.rstrip('/')
        df_wide['HOMES Webpage URL'] = df_wide['HOMES Webpage URL'].str.replace('map$', '', regex=True)
        # Remove '地図を見る' from 所在地
        df_wide['所在地'] = df_wide['所在地'].str.replace('地図を見る', '', regex=False).str.strip()
        return df_wide

    if homes_file is not None:
        if homes_file.name.endswith(".csv"):
            df_homes = pd.read_csv(homes_file)
        else:
            df_homes = pd.read_excel(homes_file, engine="openpyxl")

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
