import pandas as pd

# Load the Excel file
df = pd.read_excel("/Users/jayanth/Desktop/SheetShift/Invoice_20rows.xlsx", sheet_name="Invoice")

def find_column(df, column_name):
    """
    Finds a column in the DataFrame case-insensitively.
    """
    for col in df.columns:
        if col.lower() == column_name.lower():
            return col
    return None

# Calculate null values
null_counts = df.isnull().sum()
total_nulls = null_counts.sum()

# Filter out columns with zero nulls for the detailed breakdown
null_counts_filtered = null_counts[null_counts > 0]

result = {
    "has_null_values": bool(total_nulls > 0),
    "total_null_values": int(total_nulls),
    "null_values_per_column": null_counts_filtered.to_dict()
}
print(result)