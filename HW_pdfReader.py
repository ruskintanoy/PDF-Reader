import camelot
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import simpledialog
from openpyxl import load_workbook

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

def convert_to_number(value):
    """Converts a value to a float if it's numeric, or returns 0."""
    value = dash_to_zero(value)
    try:
        return float(value.replace(',', '').replace('$', '').strip())  
    except ValueError:
        return 0

def parse_pages_input(page_input):
    """Parse the user's input into valid pages or ranges for Camelot."""
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
    """Auto-adjust the column widths based on the content of the headers."""
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

def pdf_to_excel(folder_path, output_excel_path):
    root = tk.Tk()
    root.withdraw()  

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
    skip_conditions = ["SAMSUNG", "IPHONE", "GOOGLE", "GALAXY", "SUMMARY", "MOBILE"] 
    
    for table in tables:
        df = table.df
        
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

                corrected_data.append([
                    first_name, last_name, contact_num,
                    convert_to_number(row[1].strip()), convert_to_number(row[2].strip()),
                    convert_to_number(row[3].strip()), convert_to_number(row[4].strip()),
                    convert_to_number(row[5].strip())
                ])
                
                i += 2
            else:
                i += 1
    
    cleaned_df = pd.DataFrame(corrected_data, columns=[
        'First Name', 'Last Name', 'Contact Number',
        'Starting Balance', 'Payments ($)', 'Current Balance',
        'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)'
    ])
    
    # Write DataFrame to Excel
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, index=False, header=True)

    workbook = load_workbook(output_excel_path)
    worksheet = workbook.active

    adjust_column_widths(worksheet)  
    workbook.save(output_excel_path)
    print(f"Data written to {output_excel_path}")

folder_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\Telus Invoice"
output_excel_path = 'hw_output.xlsx'

pdf_to_excel(folder_path, output_excel_path)
