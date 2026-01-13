import streamlit as st
import pandas as pd
import io
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =============================================================================
# COLUMN CONFIGURATION
# =============================================================================

# Columns required for status calculation (must exist in P and L forms)
STATUS_REQUIRED_COLUMNS = ['Patient Address', 'Contact Number']

# Column configurations for each form type
PRESUMPTIVE_COLUMNS = [
    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
    'Contact Number', 'Gender', 'Age', 'Patient Address', 'District',
    'Opd Ipd', 'Provisional Diagnosis', 'Test Performed', 'Pathogen Name',
    'Pathogen Subtype', 'Facility Name Pform', 'Latitude', 'Longitude'
]

LABORATORY_COLUMNS = [
    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
    'Contact Number', 'Gender', 'Age', 'Patient Address', 'District',
    'Opd Ipd', 'Confirmed Diagnosis', 'Test Performed', 'Pathogen Name',
    'Pathogen Subtype', 'Facility Name Lform', 'Latitude', 'Longitude'
]

LINE_COLUMNS = [
    'Form Type', 'Reporting Date', 'Patient Name', 'Age', 'Gender',
    'Houseno', 'Hfname', 'Sformdiseasename', 'Wardname', 'Latitude', 'Longitude'
]

# File type detection prefixes
FILE_PREFIXES = {
    'Presumptive': 'P form',
    'Laboratory': 'L form',
    'Line': 'S Form'
}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def read_excel(file):
    """Read an Excel file and use the second row as the header."""
    if file is not None:
        df = pd.read_excel(file, header=1)
        return df
    return None


def extract_date(transaction_id):
    """Extract date from the transaction ID in 'DDMMYYYY' format."""
    if isinstance(transaction_id, str):
        parts = transaction_id.split('-')
        if len(parts) >= 2 and len(parts[1]) == 8:
            date_str = parts[1]
            return f"{date_str[0:2]}/{date_str[2:4]}/{date_str[4:]}"
    return None


def get_status(row):
    """Determine data completeness status based on contact info."""
    has_address = not pd.isna(row.get('Patient Address'))
    has_contact = not pd.isna(row.get('Contact Number'))

    if not has_address and not has_contact:
        return 'No Mobile and No Address'
    elif not has_address:
        return 'No Address'
    elif not has_contact:
        return 'No Mobile'
    else:
        return 'Data Available'


def validate_columns(data, required_columns, file_name):
    """Check if required columns exist in the dataframe."""
    missing = [col for col in required_columns if col not in data.columns]
    if missing:
        return False, missing
    return True, []


def filter_columns(data, columns_to_keep):
    """Filter dataframe to keep only specified columns that exist."""
    existing_columns = [col for col in columns_to_keep if col in data.columns]
    return data[existing_columns]


def detect_file_type(file_name):
    """Detect the form type based on file name prefix."""
    for prefix, form_type in FILE_PREFIXES.items():
        if file_name.startswith(prefix):
            return prefix, form_type
    return None, None


# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    st.title("Junaid's Excel File Merger")

    uploaded_files = st.file_uploader(
        "Upload Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload files starting with 'Presumptive', 'Laboratory', or 'Line'"
    )

    if not uploaded_files:
        st.info("Please upload one or more Excel files.")
        st.markdown("""
        **Supported file types:**
        - `Presumptive*.xlsx` - P Form data
        - `Laboratory*.xlsx` - L Form data
        - `Line*.xlsx` - S Form data
        """)
        return

    all_data = []
    errors = []
    warnings = []

    # Progress bar for file processing
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        file_name = uploaded_file.name
        status_text.text(f"Processing: {file_name}")

        try:
            # Read the Excel file
            data = read_excel(uploaded_file)

            if data is None or data.empty:
                warnings.append(f"'{file_name}': File is empty or could not be read")
                continue

            # Detect file type
            file_type, form_type = detect_file_type(file_name)

            if file_type is None:
                warnings.append(f"'{file_name}': Unknown file type. File name must start with 'Presumptive', 'Laboratory', or 'Line'")
                continue

            # Process based on file type
            if file_type == 'Presumptive':
                data.insert(0, 'Form Type', form_type)

                if 'Patient Transaction Id' in data.columns:
                    data.insert(1, 'Reporting Date', data['Patient Transaction Id'].apply(extract_date))
                else:
                    warnings.append(f"'{file_name}': Missing 'Patient Transaction Id' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, PRESUMPTIVE_COLUMNS)

            elif file_type == 'Laboratory':
                data.insert(0, 'Form Type', form_type)

                if 'Batch Submitteddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Batch Submitteddate'])
                else:
                    warnings.append(f"'{file_name}': Missing 'Batch Submitteddate' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, LABORATORY_COLUMNS)

            elif file_type == 'Line':
                data.insert(0, 'Form Type', form_type)

                if 'Updateddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Updateddate'])
                else:
                    warnings.append(f"'{file_name}': Missing 'Updateddate' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, LINE_COLUMNS)

            # Add status column (only for forms with contact info)
            if 'Patient Address' in data.columns or 'Contact Number' in data.columns:
                # Add missing columns with NaN if needed for status calculation
                if 'Patient Address' not in data.columns:
                    data['Patient Address'] = None
                if 'Contact Number' not in data.columns:
                    data['Contact Number'] = None
                data['Status of Mobile & Address'] = data.apply(get_status, axis=1)

            all_data.append(data)
            logger.info(f"Successfully processed: {file_name} ({len(data)} rows)")

        except Exception as e:
            errors.append(f"'{file_name}': {str(e)}")
            logger.error(f"Error processing {file_name}: {e}")

        # Update progress
        progress_bar.progress((idx + 1) / len(uploaded_files))

    status_text.empty()
    progress_bar.empty()

    # Display warnings
    if warnings:
        with st.expander(f"Warnings ({len(warnings)})", expanded=False):
            for warning in warnings:
                st.warning(warning)

    # Display errors
    if errors:
        with st.expander(f"Errors ({len(errors)})", expanded=True):
            for error in errors:
                st.error(error)

    # Check if we have any data to merge
    if not all_data:
        st.error("No valid data could be processed. Please check your files and try again.")
        return

    # Merge all data
    try:
        merged_data = pd.concat(all_data, ignore_index=True)
    except Exception as e:
        st.error(f"Error merging data: {str(e)}")
        return

    # Add duplicate detection
    if 'Patient Name' in merged_data.columns:
        merged_data['Duplicate Case'] = merged_data.duplicated(subset=['Patient Name'], keep=False)
        merged_data['Duplicate Case'] = merged_data['Duplicate Case'].map({True: 'Duplicate', False: ''})

    # Display summary
    st.success(f"Successfully merged {len(merged_data)} records from {len(all_data)} file(s)")

    # Display merged DataFrame
    st.dataframe(merged_data, use_container_width=True)

    # Create downloadable Excel file
    try:
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            merged_data.to_excel(writer, index=False, sheet_name='Merged Data')

        st.download_button(
            label="Download Merged Excel",
            data=output_buffer.getvalue(),
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error creating download file: {str(e)}")


if __name__ == "__main__":
    main()
