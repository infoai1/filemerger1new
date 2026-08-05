import streamlit as st
import pandas as pd
from excel_reader import read_excel
import io  # For in-memory file handling

def extract_date(transaction_id):
    """Extract date from the transaction ID in 'DDMMYYYY' format."""
    if isinstance(transaction_id, str):
        parts = transaction_id.split('-')
        if len(parts) >= 2 and len(parts[1]) == 8:
            date_str = parts[1]
            return f"{date_str[0:2]}/{date_str[2:4]}/{date_str[4:]}"
    return None


def select_existing_columns(data, columns):
    """Keep only columns that exist so missing Excel fields don't crash the app."""
    existing = [col for col in columns if col in data.columns]
    missing = [col for col in columns if col not in data.columns]
    if missing:
        st.warning(f"Missing columns skipped: {', '.join(missing)}")
    return data[existing]


def get_status(row):
    """Status for mobile/address; safe when Line forms omit those columns."""
    has_address = 'Patient Address' in row.index
    has_contact = 'Contact Number' in row.index
    if not has_address and not has_contact:
        return ''

    address_missing = (not has_address) or pd.isna(row['Patient Address'])
    contact_missing = (not has_contact) or pd.isna(row['Contact Number'])

    if address_missing and contact_missing:
        return 'No Mobile and No Address'
    if address_missing:
        return 'No Address'
    if contact_missing:
        return 'No Mobile'
    return 'Data Available'


def main():
    st.title("Junaid's Excel File Merger")

    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        all_data = []

        for uploaded_file in uploaded_files:
            data = read_excel(uploaded_file)
            if data is None:
                st.error(f"Could not read {uploaded_file.name}")
                continue

            # Check file name for 'Presumptive'
            if uploaded_file.name.startswith('Presumptive'):
                data.insert(0, 'Form Type', 'P form')
                if 'Patient Transaction Id' in data.columns:
                    data.insert(1, 'Reporting Date', data['Patient Transaction Id'].apply(extract_date))

                columns_to_keep_presumptive = [
                    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
                    'Contact Number', 'Gender', 'Age', 'Patient Address', 'District',
                    'Opd Ipd', 'Provisional Diagnosis', 'Test Performed', 'Pathogen Name',
                    'Pathogen Subtype', 'Facility Name Pform', 'Latitude', 'Longitude',
                ]
                data = select_existing_columns(data, columns_to_keep_presumptive)

            # Check file name for 'Laboratory'
            elif uploaded_file.name.startswith('Laboratory'):
                data.insert(0, 'Form Type', 'L form')
                if 'Batch Submitteddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Batch Submitteddate'])

                columns_to_keep_laboratory = [
                    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
                    'Contact Number', 'Gender', 'Age', 'Patient Address', 'District',
                    'Opd Ipd', 'Confirmed Diagnosis', 'Test Performed', 'Pathogen Name',
                    'Pathogen Subtype', 'Facility Name Lform', 'Latitude', 'Longitude',
                ]
                data = select_existing_columns(data, columns_to_keep_laboratory)

            # Check file name for 'Line'
            elif uploaded_file.name.startswith('Line'):
                data.insert(0, 'Form Type', 'S Form')
                if 'Updateddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Updateddate'])

                columns_to_keep_line = [
                    'Form Type', 'Reporting Date', 'Patient Name', 'Age', 'Gender',
                    'Houseno', 'Hfname', 'Sformdiseasename', 'Wardname', 'Latitude', 'Longitude',
                ]
                data = select_existing_columns(data, columns_to_keep_line)

            else:
                st.warning(
                    f"Unrecognized file name '{uploaded_file.name}'. "
                    "Expected names starting with Presumptive, Laboratory, or Line."
                )
                continue

            data['Status of Mobile & Address'] = data.apply(get_status, axis=1)
            all_data.append(data)

        if not all_data:
            st.write("No usable files were processed.")
            return

        merged_data = pd.concat(all_data, ignore_index=True)

        if 'Patient Name' in merged_data.columns:
            merged_data['Duplicate Case'] = merged_data.duplicated(subset=['Patient Name'], keep=False)
            merged_data['Duplicate Case'] = merged_data['Duplicate Case'].map({True: 'Duplicate', False: ''})

        st.write(merged_data)

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            merged_data.to_excel(writer, index=False)

        st.download_button(
            label="Download Merged Excel",
            data=output_buffer.getvalue(),
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.write("Please upload one or more Excel files.")

if __name__ == "__main__":
    main()
