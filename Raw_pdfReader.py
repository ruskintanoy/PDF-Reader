import camelot
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import simpledialog

def format_contact_number(contact_num):
    formatted_number = re.sub(r'(\d{3})[ ](\d{3}-\d{4})', r'\1-\2', contact_num)
    return formatted_number

def find_pdf_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            return os.path.join(folder_path, file_name)
    return None

def dash_to_zero(value):
    return '0' if value == '-' else value

def parse_pages_input(page_input):
    """Parse the user's input into a valid string for Camelot."""
    # This will allow ranges like '1-3' and individual pages like '5' and combine them
    pages = []
    for part in page_input.split(','):
        part = part.strip()
        if '-' in part:  # Handle ranges like '1-3'
            start, end = part.split('-')
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))  # Handle single pages like '5'
    
    return ','.join(map(str, pages))  # Convert to a string that Camelot accepts

def pdf_to_excel(folder_path, output_excel_path):
    # Initialize Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window as we only need the dialog

    # Prompt for pages using Tkinter's simpledialog
    page_input = simpledialog.askstring("Input", "Enter the pages you want to extract:")

    if not page_input:
        print("No pages entered.")
        return

    pages = parse_pages_input(page_input) 

    pdf_path = find_pdf_in_folder(folder_path)
    
    if not pdf_path:
        print("No PDF file found in the folder.")
        return

    tables = camelot.read_pdf(pdf_path, flavor='stream', pages=pages)  
    corrected_data = []
    skip_conditions = ["BBAN", "TABLET", "BUSINESS", "MOBILE", "SUMMARY", "ACCOUNT"] 

    for table in tables:
        df = table.df
        has_partial_charges = len(df.columns) == 8

        if has_partial_charges:
            headers = ['First Name', 'Last Name', 'Contact Number', 'Partial Charges ($)',
                       'Monthly and Other Charges ($)', 'Add-Ons ($)', 'Usage Charges ($)',
                       'Total Before Taxes ($)', 'Taxes ($)', 'Total ($)']
        else:
            headers = ['First Name', 'Last Name', 'Contact Number',
                       'Monthly and Other Charges ($)', 'Add-Ons ($)', 'Usage Charges ($)',
                       'Total Before Taxes ($)', 'Taxes ($)', 'Total ($)']

        i = 0
        while i < len(df):
            row = df.iloc[i]
            
            if any(skip_text.lower() in row[0].lower() for skip_text in skip_conditions):
                i += 1
                continue  
            
            if row[0].strip() and len(row[0].split()) > 1:
                name_parts = row[0].split(maxsplit=1)
                first_name = name_parts[0]
                last_name = name_parts[1] if len(name_parts) > 1 else ''
                
                contact_row = df.iloc[i + 1] if i + 1 < len(df) else None
                contact_num = contact_row[0].strip() if contact_row is not None else ''

                contact_num = format_contact_number(contact_num)

                if has_partial_charges:
                    corrected_data.append([
                        first_name, last_name, contact_num,
                        dash_to_zero(row[1].strip()), dash_to_zero(row[2].strip()),
                        dash_to_zero(row[3].strip()), dash_to_zero(row[4].strip()),
                        dash_to_zero(row[5].strip()), dash_to_zero(row[6].strip()), 
                        dash_to_zero(row[7].strip())
                    ])
                else:
                    corrected_data.append([
                        first_name, last_name, contact_num,
                        dash_to_zero(row[1].strip()), dash_to_zero(row[2].strip()),
                        dash_to_zero(row[3].strip()), dash_to_zero(row[4].strip()),
                        dash_to_zero(row[5].strip()), dash_to_zero(row[6].strip())
                    ])
                
                i += 2
            else:
                i += 1

    cleaned_df = pd.DataFrame(corrected_data, columns=headers)
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, index=False, header=True)

    print(f"Data written to {output_excel_path}")

# Folder path and output file path
folder_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\Telus Invoice"
output_excel_path = 'raw_output.xlsx'

# Run the function to process the PDF and export to Excel
pdf_to_excel(folder_path, output_excel_path)
