import camelot
import pandas as pd

def pdf_to_excel(pdf_path, output_excel_path):
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='5')
    
    df = tables[0].df
    
    keywords = ["SAMSUNG", "GOOGLE", "IPHONE", "GALAXY", "BLACK", "TABLET"]
    
    cleaned_data = []
    current_record = []  

    for i in range(0, len(df)):
        row = df.iloc[i]
        
        if any(keyword.lower() in row[0].lower() for keyword in keywords):
            continue 
        
        if row[0].strip() and len(row[0].split()) > 1:  
            if current_record: 
                cleaned_data.append(current_record)
            
            name_parts = row[0].split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ''
            current_record = [
                first_name, last_name, '',  
                row[1], row[2], row[3], row[4], row[5], row[6]
            ]
        
        elif i > 0 and df.iloc[i - 1, 0].strip() and not row[0].strip():
           
            contact_num = ' '.join([str(row[0]).strip(), str(row[1]).strip()]).strip()  
            current_record[2] = contact_num 
        
        elif not any(row[1:].str.strip()):
            continue
    
    if current_record:
        cleaned_data.append(current_record)

    cleaned_df = pd.DataFrame(cleaned_data, columns=[
        'First Name', 'Last Name', 'Contact Number',
        'Starting Balance', 'Payments ($)', 'Current Balance',
        'Starting Device Discount Balance ($)', 'Current Device Discount Balance ($)', 'End Date'
    ])
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, sheet_name='Hardware', index=False)

    print(f"Data has been successfully written to {output_excel_path}")

pdf_path = r"C:\Users\ruskin\Spaar Inc\SPAAR IT - Documents\Telus Monthly Bill\2024\9 - 2024 September\TELUS-INVOICE.pdf"
output_excel_path = 'output_file_cleaned_v4.xlsx'
pdf_to_excel(pdf_path, output_excel_path)
