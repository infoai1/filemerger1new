import pandas as pd

def read_excel(file):
    """Read an Excel file and use the second row as the header."""
    if file is not None:
        df = pd.read_excel(file, header=1)  # Skip the first row and use the second row as the header
        return df
    return None
