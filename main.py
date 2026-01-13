"""
Junaid's Excel File Merger
Merges multiple Excel files from healthcare surveillance sources into a single CSV.
"""

import streamlit as st
import pandas as pd
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# =============================================================================
# COLUMN CONFIGURATION
# =============================================================================

# Columns to extract from each form type (before transformation)
PRESUMPTIVE_COLUMNS = [
    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
    'Contact Number', 'Gender', 'Age', 'Patient Address',
    'Opd Ipd', 'Provisional Diagnosis', 'Test Performed', 'Pathogen Name',
    'Pathogen Subtype', 'Facility Name Pform', 'District'
]

LABORATORY_COLUMNS = [
    'Form Type', 'Reporting Date', 'Date Of Onset', 'Patient Name',
    'Contact Number', 'Gender', 'Age', 'Patient Address',
    'Opd Ipd', 'Confirmed Diagnosis', 'Test Performed', 'Pathogen Name',
    'Pathogen Subtype', 'Facility Name Lform', 'District'
]

LINE_COLUMNS = [
    'Form Type', 'Reporting Date', 'Patient Name', 'Age', 'Gender',
    'Houseno', 'Hfname', 'Sformdiseasename', 'Wardname'
]

# Column renaming map (old name -> new name)
COLUMN_RENAME_MAP = {
    'Form Type': 'Type',
    'Reporting Date': 'reporting Datae',
    'Patient Name': 'Name of Patient'
}

# Final output column order (17 columns)
OUTPUT_COLUMNS = [
    'MSU Unique Code',
    'Type',
    'reporting Datae',
    'Date Of Onset',
    'Name of Patient',
    'Contact Number',
    'Gender',
    'Age',
    'Patient Address',
    'Ward',
    'Confirmed Diagnosis',
    'District',
    'Opd Ipd',
    'Test Performed',
    'Pathogen Name',
    'Pathogen Subtype',
    'Facility Name Lform'
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
    """
    Read an Excel file and use the second row as the header.

    Args:
        file: Uploaded file object from Streamlit

    Returns:
        DataFrame or None if file is invalid
    """
    if file is not None:
        return pd.read_excel(file, header=1)
    return None


def extract_date(transaction_id):
    """
    Extract date from the transaction ID in 'DDMMYYYY' format.

    Args:
        transaction_id: String containing transaction ID with embedded date

    Returns:
        Formatted date string (DD/MM/YYYY) or None if parsing fails
    """
    if isinstance(transaction_id, str):
        parts = transaction_id.split('-')
        if len(parts) >= 2 and len(parts[1]) == 8:
            date_str = parts[1]
            return f"{date_str[0:2]}/{date_str[2:4]}/{date_str[4:]}"
    return None


def filter_columns(data, columns_to_keep):
    """
    Filter dataframe to keep only specified columns that exist.

    Args:
        data: Input DataFrame
        columns_to_keep: List of column names to retain

    Returns:
        DataFrame with only the specified columns
    """
    existing_columns = [col for col in columns_to_keep if col in data.columns]
    return data[existing_columns]


def detect_file_type(file_name):
    """
    Detect the form type based on file name prefix.

    Args:
        file_name: Name of the uploaded file

    Returns:
        Tuple of (file_type, form_type) or (None, None) if unknown
    """
    for prefix, form_type in FILE_PREFIXES.items():
        if file_name.startswith(prefix):
            return prefix, form_type
    return None, None


def transform_to_output_format(data):
    """
    Transform merged data to the final output format.

    Performs:
    - Column renaming
    - Adding empty columns (MSU Unique Code, Ward)
    - Reordering columns to match OUTPUT_COLUMNS

    Args:
        data: Merged DataFrame

    Returns:
        DataFrame in final output format
    """
    # Rename columns
    data = data.rename(columns=COLUMN_RENAME_MAP)

    # Add empty columns
    data['MSU Unique Code'] = ''
    data['Ward'] = ''

    # Ensure all output columns exist (fill missing with empty)
    for col in OUTPUT_COLUMNS:
        if col not in data.columns:
            data[col] = ''

    # Reorder to final column order
    return data[OUTPUT_COLUMNS]


# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    """Main Streamlit application entry point."""

    st.title("Junaid's Excel File Merger")

    # File upload section
    uploaded_files = st.file_uploader(
        "Upload Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload files starting with 'Presumptive', 'Laboratory', or 'Line'"
    )

    # Show instructions if no files uploaded
    if not uploaded_files:
        st.info("Please upload one or more Excel files.")
        st.markdown("""
        **Supported file types:**
        - `Presumptive*.xlsx` - P Form data
        - `Laboratory*.xlsx` - L Form data
        - `Line*.xlsx` - S Form data
        """)
        return

    # Initialize collections for processing
    all_data = []
    errors = []
    warnings = []

    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()

    # ==========================================================================
    # FILE PROCESSING LOOP
    # ==========================================================================

    for idx, uploaded_file in enumerate(uploaded_files):
        file_name = uploaded_file.name
        status_text.text(f"Processing: {file_name}")

        try:
            # Read the Excel file
            data = read_excel(uploaded_file)

            if data is None or data.empty:
                warnings.append(f"'{file_name}': File is empty or could not be read")
                continue

            # Detect file type from filename
            file_type, form_type = detect_file_type(file_name)

            if file_type is None:
                warnings.append(
                    f"'{file_name}': Unknown file type. "
                    "File name must start with 'Presumptive', 'Laboratory', or 'Line'"
                )
                continue

            # ------------------------------------------------------------------
            # Process Presumptive files (P Form)
            # ------------------------------------------------------------------
            if file_type == 'Presumptive':
                data.insert(0, 'Form Type', form_type)

                # Extract date from transaction ID
                if 'Patient Transaction Id' in data.columns:
                    data.insert(1, 'Reporting Date',
                               data['Patient Transaction Id'].apply(extract_date))
                else:
                    warnings.append(f"'{file_name}': Missing 'Patient Transaction Id' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, PRESUMPTIVE_COLUMNS)

            # ------------------------------------------------------------------
            # Process Laboratory files (L Form)
            # ------------------------------------------------------------------
            elif file_type == 'Laboratory':
                data.insert(0, 'Form Type', form_type)

                # Use batch submission date
                if 'Batch Submitteddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Batch Submitteddate'])
                else:
                    warnings.append(f"'{file_name}': Missing 'Batch Submitteddate' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, LABORATORY_COLUMNS)

            # ------------------------------------------------------------------
            # Process Line files (S Form)
            # ------------------------------------------------------------------
            elif file_type == 'Line':
                data.insert(0, 'Form Type', form_type)

                # Use update date
                if 'Updateddate' in data.columns:
                    data.insert(1, 'Reporting Date', data['Updateddate'])
                else:
                    warnings.append(f"'{file_name}': Missing 'Updateddate' column")
                    data.insert(1, 'Reporting Date', None)

                data = filter_columns(data, LINE_COLUMNS)

            # Add processed data to collection
            all_data.append(data)
            logger.info(f"Successfully processed: {file_name} ({len(data)} rows)")

        except Exception as e:
            errors.append(f"'{file_name}': {str(e)}")
            logger.error(f"Error processing {file_name}: {e}")

        # Update progress bar
        progress_bar.progress((idx + 1) / len(uploaded_files))

    # Clear progress indicators
    status_text.empty()
    progress_bar.empty()

    # ==========================================================================
    # DISPLAY WARNINGS AND ERRORS
    # ==========================================================================

    if warnings:
        with st.expander(f"Warnings ({len(warnings)})", expanded=False):
            for warning in warnings:
                st.warning(warning)

    if errors:
        with st.expander(f"Errors ({len(errors)})", expanded=True):
            for error in errors:
                st.error(error)

    # Check if we have any data to merge
    if not all_data:
        st.error("No valid data could be processed. Please check your files and try again.")
        return

    # ==========================================================================
    # MERGE AND TRANSFORM DATA
    # ==========================================================================

    try:
        merged_data = pd.concat(all_data, ignore_index=True)
    except Exception as e:
        st.error(f"Error merging data: {str(e)}")
        return

    # Transform to final output format
    output_data = transform_to_output_format(merged_data)

    # Display summary
    st.success(f"Successfully merged {len(output_data)} records from {len(all_data)} file(s)")

    # Display merged DataFrame
    st.dataframe(output_data, use_container_width=True)

    # ==========================================================================
    # CSV DOWNLOAD
    # ==========================================================================

    try:
        csv_data = output_data.to_csv(index=False)

        st.download_button(
            label="Download Merged CSV",
            data=csv_data,
            file_name="merged_data.csv",
            mime="text/csv"
        )
    except Exception as e:
        st.error(f"Error creating download file: {str(e)}")


if __name__ == "__main__":
    main()
