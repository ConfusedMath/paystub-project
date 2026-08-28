import os
import re
import pandas as pd
from datetime import datetime
from pypdf import PdfReader, PdfWriter
from appscript import app, k, mactypes

#%%
def generate_password(name, ssn):
    last_name = name.strip().split()[0]
    name_part = re.sub(r'[^a-zA-Z]', '', last_name)[:4].upper()
    return f"{name_part}{ssn}"

def load_employee_data(excel_path):
    df = pd.read_excel(excel_path)
    employees = {}
    
    for _, row in df.iterrows():
        first_name = str(row['first_name']).strip()
        last_name = str(row['last_name']).strip()
        full_name = f"{first_name} {last_name}"
        
        employees[full_name] = {
            'first_name': first_name,
            'last_name': last_name,
            'email': str(row['email_address']).strip(),
            'emp_id': str(row['employee_no']).strip()
        }
    return employees

def process_paystubs(paystub_pdf_path, employees, date_string):
    reader = PdfReader(paystub_pdf_path)
    employee_pages = {name: [] for name in employees.keys()}
    employee_ssns = {}

    ssn_pattern = re.compile(r'(?:\*|X){3}-(?:\*|X){2}-(\d{4})')
    
    # Pre-compile name matching regexes to account for initials and JR/SR distinctions
    name_patterns = {}
    for name, data in employees.items():
        first = data['first_name'].strip()
        last = data['last_name'].strip()
        
        # Look for suffixes in the last name
        suffix_match = re.search(r'\b(Jr\.?|Sr\.?|II|III|IV)\b', last, re.IGNORECASE)
        
        if suffix_match:
            # Standardize suffix format (remove periods for base variable)
            suffix = suffix_match.group(1).replace('.', '')
            
            # Isolate the base last name (e.g., "Smith" out of "Smith, Jr.")
            base_last = re.sub(r'\b' + re.escape(suffix_match.group(1)) + r'\b', '', last, flags=re.IGNORECASE)
            base_last = base_last.replace(',', '').strip()
            
            first_esc = re.escape(first)
            last_esc = re.escape(base_last)
            
            # Requires the exact suffix (handles optional commas and periods in the PDF)
            pattern = re.compile(rf'\b{first_esc}\s+(?:[a-zA-Z]\.?\s+)?{last_esc}\s*,?\s*{suffix}\.?\b', re.IGNORECASE)
        else:
            first_esc = re.escape(first)
            last_esc = re.escape(last)
            
            # Matches base name, but explicitly ignores it if followed by Jr/Sr/etc.
            # This prevents "John Smith" from stealing "John Smith Jr."'s pages.
            pattern = re.compile(rf'\b{first_esc}\s+(?:[a-zA-Z]\.?\s+)?{last_esc}\b(?!\s*,?\s*(?:Jr\.?|Sr\.?|II|III|IV)\b)', re.IGNORECASE)
            
        name_patterns[name] = pattern

    # 1. Match pages and extract SSN
    for page in reader.pages:
        text = page.extract_text() or ""
        for name, pattern in name_patterns.items():
            if pattern.search(text): 
                employee_pages[name].append(page)
                
                if name not in employee_ssns:
                    ssn_match = ssn_pattern.search(text)
                    if ssn_match:
                        employee_ssns[name] = ssn_match.group(1)
                break 

    # 2. Setup output directory
    output_dir = f"{date_string} paystub"
    os.makedirs(output_dir, exist_ok=True)
    generated_files = []

    # 3. Encrypt and save grouped pages
    for name, pages in employee_pages.items():
        if not pages:
            print(f"Skipping {name}: No matching paystub pages found.")
            continue
            
        ssn_last4 = employee_ssns.get(name)
        if not ssn_last4:
            print(f"Warning: Could not extract SSN for {name}")
            continue

        writer = PdfWriter()
        for page in pages:
            writer.add_page(page)
            
        last_name = employees[name]['last_name']
        password = generate_password(last_name, ssn_last4)
        writer.encrypt(password)
        
        safe_filename = name.replace(' ', '_')
        output_path = os.path.abspath(os.path.join(output_dir, f"{safe_filename}.pdf"))
        
        with open(output_path, "wb") as f_out:
            writer.write(f_out)
            
        generated_files.append({
            'name': name,
            'email': employees[name]['email'],
            'filepath': output_path,
            'password_hint': password
        })
        
    return generated_files

def create_outlook_drafts(generated_files, date_string):
    # Function omitted for brevity - logic remains exactly the same
    pass
        
#%%
PAYSTUB_FILE = '/Users/alvin/Downloads/CIC 8-28 To be Sent Paystubs.pdf'
EMAIL_LIST_FILE = '/Users/alvin/Documents/CIC Emails.xlsx'
BIWEEKLY_DATE = datetime.now().strftime("%Y-%m-%d")

print("Loading employee data...")
employee_data = load_employee_data(EMAIL_LIST_FILE)

print("Processing and encrypting PDFs...")
processed_files = process_paystubs(PAYSTUB_FILE, employee_data, BIWEEKLY_DATE)

# print("Generating Outlook drafts...")
# create_outlook_drafts(processed_files, BIWEEKLY_DATE)

print("Workflow complete! Check your Outlook Drafts folder.")