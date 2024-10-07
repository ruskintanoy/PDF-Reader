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

def clean_data_frame(df, skip_conditions):
    """
    Clean the DataFrame by removing unwanted sections such as 'Summary of mobile data sharing'.
    """
    cleaned_data = []
    
    for i in range(len(df)):
        row = df.iloc[i]
        first_column_text = row[0].strip().lower()
        
        # Detect and skip the "Summary of mobile data sharing" section
        if "summary of mobile data sharing" in first_column_text:
            print("Detected unwanted section. Stopping extraction after this.")
            break
        
        # Skip any row that matches skip_conditions
        if any(skip_text.lower() in first_column_text for skip_text in skip_conditions):
            continue
        
        cleaned_data.append(row)
    
    # Convert cleaned data back to DataFrame
    return pd.DataFrame(cleaned_data)

def pdf_to_excel(folder_path, output_excel_path):
    root = tk.Tk()
    root.withdraw()

    table_type = simpledialog.askstring("Input", "Which table do you want to extract? (Hardware or Raw):")

    if not table_type or table_type.lower() not in ['hardware', 'raw']:
        print("Invalid selection. Please choose either 'Hardware' or 'Raw'.")
        return
    
    page_input = simpledialog.askstring("Input", "Enter the pages you want to extract:")

    if not page_input:
        print("No pages entered.")
        return

    pages = parse_pages_input(page_input)
    pdf_path = find_pdf_in_folder(folder_path)
    
    if not pdf_path:
        print("No PDF file found in the folder.")
        return

    try:
        # Extract tables from the PDF
        tables = camelot.read_pdf(pdf_path, flavor='stream', pages=pages)
        
        # This will hold all the cleaned data across multiple tables
        final_data = []
        
        # Skip conditions to avoid unwanted sections
        if table_type.lower() == 'hardware':
            headers = ['First Name', 'Last Name', 'Contact Number',
                       'Starting Balance', 'Payments ($)', 'Current Balance',
                       'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)']
            skip_conditions = ["SAMSUNG", "IPHONE", "GOOGLE", "GALAXY", "SUMMARY", "MOBILE"]

        else:  # Raw table
            skip_conditions = ["BBAN", "BUSINESS", "MOBILE", "SUMMARY", "ACCOUNT", "TABLET"]

            has_partial_charges = True  # Set based on inspection
            if has_partial_charges:
                headers = ['First Name', 'Last Name', 'Contact Number', 'Partial Charges ($)',
                           'Monthly and Other Charges ($)', 'Add-Ons ($)', 'Usage Charges ($)',
                           'Total Before Taxes ($)', 'Taxes ($)', 'Total ($)']
            else:
                headers = ['First Name', 'Last Name', 'Contact Number',
                           'Monthly and Other Charges ($)', 'Add-Ons ($)', 'Usage Charges ($)',
                           'Total Before Taxes ($)', 'Taxes ($)', 'Total ($)']

        # Process each extracted table
        for table in tables:
            df = table.df  # Convert the table to a DataFrame
            
            # Clean the DataFrame: filter out unwanted rows and sections
            cleaned_df = clean_data_frame(df, skip_conditions)
            
            # Append cleaned data to final data
            final_data.append(cleaned_df)
        
        # Combine all cleaned tables into a single DataFrame
        combined_df = pd.concat(final_data, ignore_index=True)

        # Write the final cleaned DataFrame to Excel
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, header=True)

        # Adjust column widths for better visibility
        workbook = load_workbook(output_excel_path)
        worksheet = workbook.active
        adjust_column_widths(worksheet)  
        workbook.save(output_excel_path)
        
        print(f"Data written to {output_excel_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

folder_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\Telus Invoice"
output_excel_path = 'telus_output.xlsx'

pdf_to_excel(folder_path, output_excel_path)
