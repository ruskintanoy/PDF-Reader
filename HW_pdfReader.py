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

def pdf_to_excel(folder_path, output_excel_path):

    root = tk.Tk()
    root.withdraw()  
    page_number = simpledialog.askstring("Input", "Enter the page number you want to extract:")

    if not page_number:
        print("No page number entered.")
        return

    pdf_path = find_pdf_in_folder(folder_path)
    
    if not pdf_path:
        print("No PDF file found in the folder.")
        return

    tables = camelot.read_pdf(pdf_path, flavor='stream', pages=page_number)  
    df = tables[0].df   
    skip_conditions = ["SAMSUNG", "GOOGLE", "IPHONE", "BLACK", "SUMMARY", "USER"] 
    corrected_data = []
    
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
                dash_to_zero(row[1].strip()), dash_to_zero(row[2].strip()), dash_to_zero(row[3].strip()),
                dash_to_zero(row[4].strip()), dash_to_zero(row[5].strip())
            ])
            
            i += 2
        else:
            i += 1
    
    cleaned_df = pd.DataFrame(corrected_data, columns=[
        'First Name', 'Last Name', 'Contact Number',
        'Starting Balance', 'Payments ($)', 'Current Balance',
        'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)'
    ])
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, index=False, header=False)

    print(f"Data written to {output_excel_path}")

folder_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\Telus Invoice"
output_excel_path = 'hw_output.xlsx'

pdf_to_excel(folder_path, output_excel_path)
