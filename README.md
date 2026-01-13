# Excel File Merger

A Streamlit web application for merging healthcare/disease surveillance Excel files from different sources.

## Features

- Merge multiple Excel files into a single dataset
- Support for three form types:
  - **Presumptive** (P Form) - Presumptive case data
  - **Laboratory** (L Form) - Laboratory test results
  - **Line** (S Form) - Line-listed case data
- Automatic date extraction from transaction IDs
- Data completeness status calculation
- Duplicate patient detection
- Download merged data as Excel file

## Usage

1. Run the application:
   ```bash
   streamlit run main.py
   ```

2. Upload Excel files through the web interface
   - Files must be named with the appropriate prefix:
     - `Presumptive*.xlsx` for P Form data
     - `Laboratory*.xlsx` for L Form data
     - `Line*.xlsx` for S Form data

3. View the merged data in the browser

4. Download the merged Excel file

## File Requirements

### Presumptive Files
Must contain columns: Patient Transaction Id, Date Of Onset, Patient Name, Contact Number, Gender, Age, Patient Address, District, Opd Ipd, Provisional Diagnosis, Test Performed, Pathogen Name, Pathogen Subtype, Facility Name Pform, Latitude, Longitude

### Laboratory Files
Must contain columns: Batch Submitteddate, Date Of Onset, Patient Name, Contact Number, Gender, Age, Patient Address, District, Opd Ipd, Confirmed Diagnosis, Test Performed, Pathogen Name, Pathogen Subtype, Facility Name Lform, Latitude, Longitude

### Line Files
Must contain columns: Updateddate, Patient Name, Age, Gender, Houseno, Hfname, Sformdiseasename, Wardname, Latitude, Longitude

## Installation

```bash
pip install -r requirements.txt
```

## Development

This project includes a Dev Container configuration for VS Code / GitHub Codespaces.

## Dependencies

- streamlit
- pandas
- openpyxl
