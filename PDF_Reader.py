import camelot
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import simpledialog, filedialog
from openpyxl import load_workbook

MONTH_REGEX = r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\b'

def format_contact_number(contact_num):
    formatted_number = re.sub(r'(\d{3})[ ](\d{3}-\d{4})', r'\1-\2', contact_num)
    return formatted_number

def dash_to_zero(value):
    return '0' if value == '-' else value

def convert_to_number(value):
    value = dash_to_zero(value)
    try:
        return float(value.replace(',', '').replace('$', '').strip())  
    except ValueError:
        return 0

def parse_pages_input(page_input):
    pages = []
    for part in page_input.split(','):
        part = part.strip()
        if '-' in part:  
            start, end = part.split('-')
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part)) 
    return ','.join(map(str, pages))  

def adjust_column_widths(worksheet):
    for column_cells in worksheet.columns:
        max_length = 0
        column = column_cells[0].column_letter  
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2  
        worksheet.column_dimensions[column].width = adjusted_width

def clean_data_frame(df, skip_conditions, table_type):
    cleaned_data = []
   
    if table_type.lower() == 'hardware':
        headers = ['First Name', 'Last Name', 'Contact Number',
                   'Starting Balance', 'Payments ($)', 'Current Balance',
                   'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)']
    else:
        if len(df.columns) == 8:  
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
        first_column_text = row[0].strip()

        if re.search(MONTH_REGEX, first_column_text):
            i += 1
            continue

        if "summary of mobile data sharing" in first_column_text.lower():
            print("Detected unwanted section. Stopping extraction after this.")
            break

        if any(skip_text.lower() in first_column_text.lower() for skip_text in skip_conditions):
            i += 1
            continue

        if first_column_text.strip() and len(first_column_text.split()) > 1:
            name_parts = first_column_text.split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ''

            contact_row = df.iloc[i + 1] if i + 1 < len(df) else None
            contact_num = contact_row[0].strip() if contact_row is not None else ''
            contact_num = format_contact_number(contact_num)

            if table_type.lower() == 'hardware':
                cleaned_data.append([
                    first_name, last_name, contact_num,
                    convert_to_number(row[1].strip()), convert_to_number(row[2].strip()),
                    convert_to_number(row[3].strip()), convert_to_number(row[4].strip()),
                    convert_to_number(row[5].strip())
                ])
            else:
                if len(df.columns) == 8: 
                    cleaned_data.append([
                        first_name, last_name, contact_num,
                        convert_to_number(row[1].strip()), convert_to_number(row[2].strip()),
                        convert_to_number(row[3].strip()), convert_to_number(row[4].strip()),
                        convert_to_number(row[5].strip()), convert_to_number(row[6].strip()),
                        convert_to_number(row[7].strip())
                    ])
                else: 
                    cleaned_data.append([
                        first_name, last_name, contact_num,
                        convert_to_number(row[1].strip()), convert_to_number(row[2].strip()),
                        convert_to_number(row[3].strip()), convert_to_number(row[4].strip()),
                        convert_to_number(row[5].strip()), convert_to_number(row[6].strip())
                    ])

            i += 2
        else:
            i += 1
    
    return pd.DataFrame(cleaned_data, columns=headers)

def pdf_to_excel(output_excel_path):
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if not pdf_path:
        print("No PDF file selected.")
        return

    table_type = simpledialog.askstring("Input", "Which table do you want to extract? (Hardware or Raw):")

    if not table_type or table_type.lower() not in ['hardware', 'raw']:
        print("Invalid selection. Please choose either 'Hardware' or 'Raw'.")
        return
    
    page_input = simpledialog.askstring("Input", "Enter the pages you want to extract:")

    if not page_input:
        print("No pages entered.")
        return

    pages = parse_pages_input(page_input)

    try:
        tables = camelot.read_pdf(pdf_path, flavor='stream', pages=pages)
  
        final_data = []

        if table_type.lower() == 'hardware':
            skip_conditions = ["SAMSUNG", "IPHONE", "GOOGLE", "GALAXY", "SUMMARY", "MOBILE", "SPAAR"]
        else:  
            skip_conditions = ["BBAN", "BUSINESS", "MOBILE", "SUMMARY", "ACCOUNT", "TABLET", "SPAAR"]

        for table in tables:
            df = table.df  
            cleaned_df = clean_data_frame(df, skip_conditions, table_type)
            final_data.append(cleaned_df)

        combined_df = pd.concat(final_data, ignore_index=True)

        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, header=True)

        workbook = load_workbook(output_excel_path)
        worksheet = workbook.active
        adjust_column_widths(worksheet)  
        workbook.save(output_excel_path)
        
        print(f"Data written to {output_excel_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

output_excel_path = 'telus_output.xlsx'

pdf_to_excel(output_excel_path)
