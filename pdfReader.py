import camelot
import pandas as pd

def pdf_to_excel(pdf_path, output_excel_path):
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='5')  # switch to the page you want to copy
    
    df = tables[0].df
    
    skip_conditions = ["SAMSUNG", "GOOGLE", "IPHONE", "GALAXY", "BLACK", "TABLET"]
    
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
            
            corrected_data.append([
                first_name, last_name, contact_num,
                row[1].strip(), row[2].strip(), row[3].strip(),
                row[4].strip(), row[5].strip()  
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

pdf_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\2024\9 - 2024 September\TELUS-INVOICE.pdf" # change the fle path to the correct month
output_excel_path = 'OUTPUT.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
