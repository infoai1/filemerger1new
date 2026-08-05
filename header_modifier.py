import pandas as pd

def modify_header(df):
    """Modify the DataFrame header."""
    if df is not None and not df.empty:
        new_header = df.iloc[0]  # Assign the first row as header
        df = df[1:]  # Remove the first row

        # Handling NaN values in the header
        new_header = new_header.fillna("Unnamed")

        # Handling duplicate column names
        new_header = rename_duplicates(new_header)

        df.columns = new_header  # Set the new header
    return df

def rename_duplicates(old_header):
    """
    Rename duplicate column names by appending an index.
    """
    new_header = []
    count = {}
    for item in old_header:
        if item in count:
            count[item] += 1
            new_header.append(f"{item}_{count[item]}")
        else:
            count[item] = 0
            new_header.append(item)
    return new_header
