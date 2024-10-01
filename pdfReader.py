import camelot
import pandas as pd

# Step 2: Extract tables from the PDF using camelot with stream flavor
def pdf_to_excel(pdf_path, output_excel_path):
    # Use the 'stream' flavor to read the tables from the PDF
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='5')
    
    # Step 3: Convert the table to a pandas DataFrame (assuming single table)
    df = tables[0].df
    
    # Step 4: Initialize a list to hold cleaned rows
    cleaned_data = []
    
    # Step 5: Iterate through the rows of the dataframe
    for i in range(0, len(df), 2):  # Step over 2 rows since the contact number is on the next row
        row = df.iloc[i]
        
        # Check if the name column is not empty or contains unexpected data
        if not row[0].strip():
            continue  # Skip rows without a valid name
        
        # Extract first and last names from column 0
        name_parts = row[0].split(maxsplit=1)
        if len(name_parts) < 1:
            continue  # Skip rows that don't have at least a first name
        
        first_name = name_parts[0]
        last_name = name_parts[1] if len(name_parts) > 1 else ''
        
        # Extract the contact number from the next row (i+1), if available
        contact_num = df.iloc[i + 1, 0] if i + 1 < len(df) else ''
        
        # Remove unwanted device-related rows based on empty values in key columns
        if not row[1].strip() or not row[3].strip():  # Skips empty or invalid rows
            continue
        
        # Add the cleaned row to the list
        cleaned_data.append([
            first_name, last_name, contact_num,
            row[1], row[2], row[3], row[4], row[5], row[6]
        ])
    
    # Convert the cleaned data to a new DataFrame with correct headers
    cleaned_df = pd.DataFrame(cleaned_data, columns=[
        'First Name', 'Last Name', 'Contact Number',
        'Starting Balance', 'Payments ($)', 'Current Balance',
        'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)', 'End Date'
    ])
    
    # Step 6: Write the cleaned DataFrame to an Excel file
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, sheet_name='Hardware', index=False)

    print(f"Data has been successfully written to {output_excel_path}")

# Example usage
pdf_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\2024\9 - 2024 September\TELUS-INVOICE.pdf"
output_excel_path = 'output_file_cleaned.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
