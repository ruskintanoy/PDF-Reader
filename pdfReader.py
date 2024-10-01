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
    current_record = []  # To hold partial rows

    # Step 5: Iterate through the rows of the dataframe
    for i in range(0, len(df)):
        row = df.iloc[i]
        
        # Check if the first column contains any of the keywords, skip the row if true
        if any(keyword.lower() in row[0].lower() for keyword in keywords):
            continue  # Skip this row as it's a device-related row
        
        # If the row contains a name (detected by presence of non-empty string in column 0)
        if row[0].strip() and len(row[0].split()) > 1:  # Ensure that the first column has more than one word (name)
            if current_record:  # If we already have a record in progress, save it before starting a new one
                cleaned_data.append(current_record)
            
            # Start a new record by processing name and subsequent data
            name_parts = row[0].split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ''
            current_record = [
                first_name, last_name, '',  # Start a new record with empty contact number initially
                row[1], row[2], row[3], row[4], row[5], row[6]
            ]
        
        # If the current row is likely to be a continuation of the previous record (contains a contact number)
        elif i > 0 and df.iloc[i - 1, 0].strip() and not row[0].strip():
            # Join the contact number into one value if it's split across columns
            contact_num = ' '.join([str(row[0]).strip(), str(row[1]).strip()]).strip()  # Combine first two columns for contact number
            current_record[2] = contact_num  # Update the contact number in the current record
        
        # Skip rows that are either empty or contain irrelevant info
        elif not any(row[1:].str.strip()):
            continue
    
    # Ensure the last record is also added to the cleaned data
    if current_record:
        cleaned_data.append(current_record)

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
output_excel_path = 'output_file_cleaned_v4.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
