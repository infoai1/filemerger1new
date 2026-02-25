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


    
def main():
    st.title("Junaid's Excel File Merger")

    
    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
    
    # Define the get_status function here
    def get_status(row):
        if pd.isna(row['Patient Address']) and pd.isna(row['Contact Number']):
            return 'No Mobile and No Address'
        elif pd.isna(row['Patient Address']):
            return 'No Address'
        elif pd.isna(row['Contact Number']):
            return 'No Mobile'
        else:
            return 'Data Available'


    if uploaded_files:
        all_data = []

        for uploaded_file in uploaded_files:
            data = read_excel(uploaded_file)

            # Check file name for 'Presumptive'
            if uploaded_file.name.startswith('Presumptive'):
                # Add 'Form Type' and 'Reporting Date' columns
                data.insert(0, 'Form Type', 'P form')
                if 'Patient Transaction Id' in data.columns:
                    data.insert(1, 'Reporting Date', data['Patient Transaction Id'].apply(extract_date))

                # Keep only specified columns for Presumptive
                columns_to_keep_presumptive = ['Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name', 
                                               'Contact Number', 'Gender', 'Age', 'Patient Address', 'District', 
                                               'Opd Ipd', 'Provisional Diagnosis', 'Test Performed', 'Pathogen Name',
                                               'Pathogen Subtype', 'Facility Name Pform','Latitude','Longitude']
                data = data[columns_to_keep_presumptive]
                data = data.rename(columns={'Facility Name Pform': 'Facility Name'})

            # Check file name for 'Laboratory'
            elif uploaded_file.name.startswith('Laboratory'):
                # Add 'Form Type' and 'Reporting Date' columns
                data.insert(0, 'Form Type', 'L form')
                if 'Batch Submitteddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Batch Submitteddate'])

                # Keep only specified columns for Laboratory
                columns_to_keep_laboratory =['Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name', 
                                              'Contact Number', 'Gender', 'Age', 'Patient Address', 'District', 
                                              'Opd Ipd', 'Confirmed Diagnosis', 'Test Performed', 'Pathogen Name',
                                              'Pathogen Subtype', 'Facility Name Lform','Latitude','Longitude']
                data = data[columns_to_keep_laboratory]
                data = data.rename(columns={'Facility Name Lform': 'Facility Name'})

            
            # Check file name for 'Line'
            elif uploaded_file.name.startswith('Line'):
                # Add 'Form Type' column
                data.insert(0, 'Form Type', 'S Form')

                # Add 'Reporting Date' column and copy values from 'Updateddate'
                if 'Updateddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Updateddate'])

                # Keep only specified columns for Laboratory
                columns_to_keep_line = ['Form Type', 'Reporting Date', 'Patient Name', 'Age', 'Gender','Houseno','Hfname','Sformdiseasename','Wardname','Latitude','Longitude']
                data = data[columns_to_keep_line]

            
            # Add 'Status of Mobile & Address' column
            data['Status of Mobile & Address'] = data.apply(get_status, axis=1)

            
            if data is not None:
                all_data.append(data)
    

        # Merge all data
        merged_data = pd.concat(all_data, ignore_index=True)

        # Add 'Duplicate Case' column
        # Mark as 'Duplicate' if the name appears more than once
        merged_data['Duplicate Case'] = merged_data.duplicated(subset=['Patient Name'], keep=False)
        merged_data['Duplicate Case'] = merged_data['Duplicate Case'].map({True: 'Duplicate', False: ''})


        # Display merged DataFrame
        st.write(merged_data)

        # Create an in-memory buffer
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer) as writer:
            merged_data.to_excel(writer, index=False)

        # Provide a download button for the merged Excel file
        st.download_button(label="Download Merged Excel",
                           data=output_buffer.getvalue(),
                           file_name="merged_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.write("Please upload one or more Excel files.")

if __name__ == "__main__":
    main()
