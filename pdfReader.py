# Step 1: Install required libraries
# !pip install camelot-py[cv] pandas openpyxl

import camelot
import pandas as pd

# Step 2: Extract tables from the PDF using camelot with stream flavor
def pdf_to_excel(pdf_path, output_excel_path):
    # Use the 'stream' flavor to read the tables
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='5')
    
    # Step 3: Convert each table to a pandas DataFrame
    all_tables = []
    for i, table in enumerate(tables):
        df = table.df  # This converts the table to a pandas DataFrame
        all_tables.append(df)
    
    # Step 4: Write DataFrames to an Excel file
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for idx, df in enumerate(all_tables):
            # Each table gets its own sheet in the Excel file
            df.to_excel(writer, sheet_name=f'Table_{idx+1}', index=False)

    print(f"Data has been successfully written to {output_excel_path}")

# Example usage
pdf_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\2024\9 - 2024 September\TELUS-INVOICE.pdf"
output_excel_path = 'output_file.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
