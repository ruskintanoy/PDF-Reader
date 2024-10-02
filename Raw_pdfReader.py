import camelot
import pandas as pd
import os
import re

# Function to format contact numbers
def format_contact_number(contact_num):
    formatted_number = re.sub(r'(\d{3})[ ](\d{3}-\d{4})', r'\1-\2', contact_num)
    return formatted_number

# Function to find a PDF file in the specified folder
def find_pdf_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            return os.path.join(folder_path, file_name)
    return None

# Main function to extract data from PDF and write to Excel
def pdf_to_excel(folder_path, output_excel_path):
    pdf_path = find_pdf_in_folder(folder_path)

    if not pdf_path:
        print("No PDF file found in the folder.")
        return

    # Extract the table from a fixed page number, adjust as needed
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='25')  
    df = tables[0].df

    # Check if the table contains the "PARTIAL CHARGES ($)" column
    headers = df.iloc[0]
    has_partial_charges = "PARTIAL CHARGES ($)" in headers.values

    # Define the columns based on the existence of "PARTIAL CHARGES ($)"
    if has_partial_charges:
        columns = ['USER', 'PARTIAL CHARGES ($)', 'MONTHLY AND OTHER CHARGES ($)', 
                   'ADD-ONS ($)', 'USAGE CHARGES ($)', 'TOTAL BEFORE TAXES ($)', 
                   'TAXES ($)', 'TOTAL ($)']
    else:
        columns = ['USER', 'MONTHLY AND OTHER CHARGES ($)', 'ADD-ONS ($)', 
                   'USAGE CHARGES ($)', 'TOTAL BEFORE TAXES ($)', 
                   'TAXES ($)', 'TOTAL ($)']

    # Filter for skip conditions
    skip_conditions = ["BBAN", "FORD", "BUSINESS", "TABLET", "SUMMARY", "USER"]

    corrected_data = []
    i = 1  # Start after header row
    while i < len(df):
        row = df.iloc[i]

        # Skip rows based on the defined conditions
        if any(skip_text.lower() in row[0].lower() for skip_text in skip_conditions):
            i += 1
            continue

        # Split and process names
        if row[0].strip() and len(row[0].split()) > 1:
            name_parts = row[0].split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ''

            # Get the contact number from the next row
            contact_row = df.iloc[i + 1] if i + 1 < len(df) else None
            contact_num = contact_row[0].strip() if contact_row is not None else ''
            contact_num = format_contact_number(contact_num)

            # Handle the rest of the columns based on the existence of partial charges
            if has_partial_charges:
                corrected_data.append([
                    first_name, last_name, contact_num,
                    row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip(),
                    row[5].strip(), row[6].strip(), row[7].strip()
                ])
            else:
                corrected_data.append([
                    first_name, last_name, contact_num,
                    row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip(),
                    row[5].strip(), row[6].strip()
                ])

            i += 2  # Skip the contact row as it's already processed
        else:
            i += 1

    # Create a DataFrame with the cleaned data and dynamic columns
    cleaned_df = pd.DataFrame(corrected_data, columns=[
        'First Name', 'Last Name', 'Contact Number'] + columns[1:])

    # Write to Excel
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, index=False, header=True)

    print(f"Data written to {output_excel_path}")

# Example usage
folder_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\Telus Invoice"
output_excel_path = 'telusOUTPUT_updated.xlsx'
pdf_to_excel(folder_path, output_excel_path)
