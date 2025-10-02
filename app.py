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
    HOMES cleaning logic:
    - Preserve original row order (first occurrence of company)
    - Reshape long -> wide manually (Field1 becomes columns)
    - Use kanji from Field2 for company name
    """
    # Fill repeated URLs if missing
    df['Link_to_the_Homepage'] = df['Link_to_the_Homepage'].fillna(method='ffill')
    df['URL'] = df['URL'].fillna(method='ffill')

    # Dictionary to hold processed companies
    processed = []
    seen = {}  # To preserve original order

    for _, row in df.iterrows():
        company_key = row['Company_Name']
        if company_key not in seen:
            # Initialize a new company row
            company_dict = {
                'Company_Name': None,  # kanji will be filled later
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

        # Current company dict
        company_dict = seen[company_key]

        # If Field1 is 会社名, replace Company_Name with kanji from Field2
        if row['Field1'] == '会社名':
            company_dict['Company_Name'] = row['Field2']
        else:
            # For other fields, add the value
            if row['Field1'] in company_dict:
                # Append if value already exists
                if company_dict[row['Field1']]:
                    company_dict[row['Field1']] += ' ' + str(row['Field2'])
                else:
                    company_dict[row['Field1']] = row['Field2']

    # Convert to DataFrame
    df_wide = pd.DataFrame(processed)
    # Remove trailing 'map/' from HOMES Webpage URL
    df_wide['HOMES Webpage URL'] = df_wide['HOMES Webpage URL'].str.rstrip('/')  # remove trailing slash first
    df_wide['HOMES Webpage URL'] = df_wide['HOMES Webpage URL'].str.replace('map$', '', regex=True)

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
