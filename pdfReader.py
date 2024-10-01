import camelot
import pandas as pd

# Step 2: Extract tables from the PDF using camelot with stream flavor
def pdf_to_excel(pdf_path, output_excel_path):
    # Use the 'stream' flavor to read the tables from the PDF
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='5')
    
    # Step 3: Convert the table to a pandas DataFrame (assuming single table)
    df = tables[0].df
    
    # Define keywords to identify device-related rows
    keywords = ["SAMSUNG", "GOOGLE", "IPHONE", "GALAXY", "BLACK", "TABLET"]
    
    # Step 4: Initialize a list to hold cleaned rows
    cleaned_data = []
    
    # Step 5: Iterate through the rows of the dataframe
    i = 0
    while i < len(df):
        row = df.iloc[i]
        
        # Skip device-related rows using keyword detection
        if any(keyword.lower() in row[0].lower() for keyword in keywords):
            i += 1
            continue  # Skip this row as it's a device-related row
        
        # If the row contains a valid name (first and last name)
        if row[0].strip() and len(row[0].split()) > 1:
            name_parts = row[0].split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ''
            
            # Move to the next row to get the contact number
            contact_row = df.iloc[i + 1] if i + 1 < len(df) else None
            contact_num = contact_row[0].strip() if contact_row is not None else ''
            
            # Append the cleaned row to the list (including contact number and other columns)
            cleaned_data.append([
                first_name, last_name, contact_num,
                row[1].strip(), row[2].strip(), row[3].strip(),
                row[4].strip(), row[5].strip(), row[6].strip()
            ])
            
            # Move to the next relevant record (skipping the contact row)
            i += 2
        else:
            # If the current row doesn't contain a valid name, skip it
            i += 1
    
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
output_excel_path = 'telusOUTPUT.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
