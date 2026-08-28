import os
import re
import io
import zipfile
import tempfile
import subprocess
from datetime import datetime
import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter

# Page configuration
st.set_page_config(
    page_title="PMS + CIC Paystub Processor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Clean Corporate Theme, No Emojis)
st.markdown("""
<style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.02em;
        margin-bottom: 0.25rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #475569;
        margin-bottom: 1.5rem;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# Helper Functions (Exact Matching Logic from PMS+CIC_paystubs.py)
# ==============================================================================

def generate_password(last_name: str, ssn: str) -> str:
    """Generates the encrypted PDF password: First 4 letters of base last name (uppercase) + SSN last 4."""
    last_name_clean = last_name.strip().split()[0] if last_name else ""
    name_part = re.sub(r'[^a-zA-Z]', '', last_name_clean)[:4].upper()
    return f"{name_part}{ssn}"


def parse_employee_dataframe(df: pd.DataFrame):
    """Normalizes column names and parses employee roster."""
    col_map = {}
    for col in df.columns:
        norm = str(col).strip().lower().replace(" ", "_").replace("-", "_")
        col_map[norm] = col

    first_col = next((col_map[k] for k in ['first_name', 'firstname', 'first'] if k in col_map), None)
    last_col = next((col_map[k] for k in ['last_name', 'lastname', 'last'] if k in col_map), None)
    email_col = next((col_map[k] for k in ['email_address', 'email', 'email_id', 'e_mail'] if k in col_map), None)
    emp_id_col = next((col_map[k] for k in ['employee_no', 'employee_number', 'emp_id', 'emp_no', 'employee_id'] if k in col_map), None)

    if not first_col or not last_col or not email_col:
        raise ValueError("Excel roster must include 'first_name', 'last_name', and 'email_address' columns.")

    employees = {}
    for _, row in df.iterrows():
        if pd.isna(row[first_col]) or pd.isna(row[last_col]):
            continue
        first_name = str(row[first_col]).strip()
        last_name = str(row[last_col]).strip()
        full_name = f"{first_name} {last_name}"
        email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
        emp_id = str(row[emp_id_col]).strip() if emp_id_col and pd.notna(row[emp_id_col]) else ""

        employees[full_name] = {
            'first_name': first_name,
            'last_name': last_name,
            'email': email,
            'emp_id': emp_id
        }
    return employees


def process_paystub_pdf(pdf_stream, employees: dict, date_string: str, output_folder: str = None):
    """
    Parses multi-page PDF, matches pages per employee with suffix/initial handling,
    distinguishing people with the same names except for Jr / Sr / II / III / IV,
    extracts SSN, encrypts individual paystubs, and returns metadata & byte streams.
    """
    reader = PdfReader(pdf_stream)
    employee_pages = {name: [] for name in employees.keys()}
    employee_ssns = {}

    ssn_pattern = re.compile(r'(?:[\*X]{3}-[\*X]{2}-|\b\d{3}-\d{2}-)(\d{4})')

    # Pre-compile name matching regexes to account for initials and JR/SR distinctions
    # Exactly as implemented in PMS+CIC_paystubs.py
    name_patterns = {}
    suffix_flag = {}

    for name, data in employees.items():
        first = data['first_name'].strip()
        last = data['last_name'].strip()

        # Look for suffixes in the last name (or full name)
        suffix_match = re.search(r'\b(Jr\.?|Sr\.?|II|III|IV)\b', last, re.IGNORECASE) or re.search(r'\b(Jr\.?|Sr\.?|II|III|IV)\b', first, re.IGNORECASE)

        if suffix_match:
            # Standardize suffix format (remove periods for base variable)
            matched_suffix_text = suffix_match.group(1)
            suffix = matched_suffix_text.replace('.', '')

            # Isolate the base last name (e.g., "Smith" out of "Smith, Jr.")
            base_last = re.sub(r'\b' + re.escape(matched_suffix_text) + r'\b', '', last, flags=re.IGNORECASE)
            base_last = base_last.replace(',', '').strip()

            base_first = re.sub(r'\b' + re.escape(matched_suffix_text) + r'\b', '', first, flags=re.IGNORECASE)
            base_first = base_first.replace(',', '').strip()

            first_esc = re.escape(base_first)
            last_esc = re.escape(base_last)

            # Requires the exact suffix (handles optional commas and periods in the PDF)
            pattern = re.compile(rf'\b{first_esc}\s+(?:[a-zA-Z]\.?\s+)?{last_esc}\s*,?\s*{suffix}\.?\b', re.IGNORECASE)
            suffix_flag[name] = True
        else:
            first_esc = re.escape(first)
            last_esc = re.escape(last)

            # Matches base name, but explicitly ignores it if followed by Jr/Sr/etc.
            # This prevents "John Smith" from stealing "John Smith Jr."'s pages.
            pattern = re.compile(rf'\b{first_esc}\s+(?:[a-zA-Z]\.?\s+)?{last_esc}\b(?!\s*,?\s*(?:Jr\.?|Sr\.?|II|III|IV)\b)', re.IGNORECASE)
            suffix_flag[name] = False

        name_patterns[name] = pattern

    # Order name pattern evaluations: Suffix patterns first, followed by base names with negative lookahead
    sorted_names = sorted(name_patterns.keys(), key=lambda n: 0 if suffix_flag.get(n, False) else 1)

    # 1. Match pages and extract SSN
    for page in reader.pages:
        text = page.extract_text() or ""
        for name in sorted_names:
            pattern = name_patterns[name]
            if pattern.search(text):
                employee_pages[name].append(page)
                if name not in employee_ssns:
                    ssn_match = ssn_pattern.search(text)
                    if ssn_match:
                        employee_ssns[name] = ssn_match.group(1)
                break

    # 2. Encrypt and generate PDF data
    results = []
    if output_folder:
        os.makedirs(output_folder, exist_ok=True)

    for name, pages in employee_pages.items():
        emp_info = employees[name]
        safe_filename = f"{name.replace(' ', '_')}_{date_string}.pdf"

        if not pages:
            results.append({
                'name': name,
                'email': emp_info['email'],
                'emp_id': emp_info['emp_id'],
                'status': 'No Pages Found',
                'pages_count': 0,
                'ssn_last4': None,
                'password': None,
                'filename': safe_filename,
                'pdf_bytes': None,
                'filepath': None
            })
            continue

        ssn_last4 = employee_ssns.get(name)
        if not ssn_last4:
            results.append({
                'name': name,
                'email': emp_info['email'],
                'emp_id': emp_info['emp_id'],
                'status': 'Missing SSN',
                'pages_count': len(pages),
                'ssn_last4': None,
                'password': None,
                'filename': safe_filename,
                'pdf_bytes': None,
                'filepath': None
            })
            continue

        writer = PdfWriter()
        for page in pages:
            writer.add_page(page)

        # Base last name for password generation
        last_name_for_pw = emp_info['last_name']
        password = generate_password(last_name_for_pw, ssn_last4)
        writer.encrypt(password)

        pdf_bytes_io = io.BytesIO()
        writer.write(pdf_bytes_io)
        pdf_data = pdf_bytes_io.getvalue()

        output_path = None
        if output_folder:
            output_path = os.path.abspath(os.path.join(output_folder, safe_filename))
            with open(output_path, "wb") as f_out:
                f_out.write(pdf_data)

        results.append({
            'name': name,
            'email': emp_info['email'],
            'emp_id': emp_info['emp_id'],
            'status': 'Encrypted',
            'pages_count': len(pages),
            'ssn_last4': ssn_last4,
            'password': password,
            'filename': safe_filename,
            'pdf_bytes': pdf_data,
            'filepath': output_path
        })

    return results


def create_zip_archive(processed_files: list, date_string: str) -> io.BytesIO:
    """Creates an in-memory zip archive containing all encrypted PDFs and an execution summary."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        summary_rows = []
        for item in processed_files:
            summary_rows.append({
                'Name': item['name'],
                'Email': item['email'],
                'Employee ID': item['emp_id'],
                'Status': item['status'],
                'Pages': item['pages_count'],
                'Password': item['password'] if item['password'] else 'N/A',
                'Filename': item['filename']
            })
            if item['pdf_bytes']:
                zip_file.writestr(item['filename'], item['pdf_bytes'])

        summary_df = pd.DataFrame(summary_rows)
        csv_data = summary_df.to_csv(index=False)
        zip_file.writestr(f"paystub_summary_{date_string}.csv", csv_data)

    zip_buffer.seek(0)
    return zip_buffer


def generate_eml_drafts(processed_files: list, subject_template: str, body_template: str, date_string: str, sender_email: str = ""):
    """
    Generates RFC 2822 compliant .eml files with the X-Unsent: 1 header.

    Each .eml embeds the encrypted PDF as a base64 MIME attachment at the start
    so the files are fully self-contained. Double-clicking an .eml in Microsoft Outlook,
    Apple Mail, or Thunderbird opens it as an editable unsent draft.

    Returns (output_folder, reports) where output_folder is the absolute path
    to the directory containing the generated .eml files.
    """
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    from email.utils import formatdate, formataddr

    output_folder = os.path.join(tempfile.gettempdir(), f"generated_drafts_{date_string}")
    os.makedirs(output_folder, exist_ok=True)

    reports = []
    for item in processed_files:
        if item['status'] != 'Encrypted' or not item.get('pdf_bytes'):
            reports.append({
                'Name': item['name'],
                'Email': item['email'],
                'Status': 'Skipped (No PDF)',
                'Error': item.get('status')
            })
            continue

        if not item['email']:
            reports.append({
                'Name': item['name'],
                'Email': 'N/A',
                'Status': 'Skipped (No Email)',
                'Error': 'Missing email address'
            })
            continue

        try:
            subject = subject_template.format(
                name=item['name'],
                date=date_string,
                password=item.get('password', '')
            )
            body = body_template.format(
                name=item['name'],
                date=date_string,
                password=item.get('password', '')
            )

            msg = MIMEMultipart()
            if sender_email and sender_email.strip():
                msg['From'] = sender_email.strip()
            msg['Subject'] = subject
            msg['To'] = formataddr((item['name'], item['email']))
            msg['Date'] = formatdate(localtime=True)
            msg['X-Unsent'] = '1'
            msg['MIME-Version'] = '1.0'

            # 1. Attach encrypted PDF attachment first (at the start of the draft)
            pdf_part = MIMEApplication(item['pdf_bytes'], _subtype='pdf')
            pdf_part.add_header(
                'Content-Disposition', 'attachment', filename=item['filename']
            )
            msg.attach(pdf_part)

            # 2. Attach plain-text body
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Write .eml to disk
            safe_name = item['name'].replace(' ', '_')
            eml_filename = f"{safe_name}_{date_string}.eml"
            eml_path = os.path.join(output_folder, eml_filename)
            with open(eml_path, 'w') as f:
                f.write(msg.as_string())

            reports.append({
                'Name': item['name'],
                'Email': item['email'],
                'Status': 'EML Created',
                'Filename': eml_filename,
                'Error': None
            })
        except Exception as err:
            reports.append({
                'Name': item['name'],
                'Email': item['email'],
                'Status': 'Error',
                'Filename': None,
                'Error': str(err)
            })

    return output_folder, reports


# ==============================================================================
# Main Streamlit Application
# ==============================================================================

def main():
    st.markdown('<div class="main-header">Paystub Processor</div>', unsafe_allow_html=True)

    # Sidebar settings
    with st.sidebar:
        st.header("Configuration")
        pay_date = st.date_input("Pay Period Date", value=datetime.today())
        date_str = pay_date.strftime("%Y-%m-%d")

        sender_email = st.text_input(
            "Sender Email Address",
            value="",
            placeholder="payroll@company.com",
            help="Email address set in the 'From' header of the generated .eml files."
        )

        st.subheader("Email Template")
        default_subject = "Paystub for {name} - {date}"
        default_body = (
            "Hello {name},\n\n"
            "Please find your encrypted paystub attached for the pay period ending {date}.\n\n"
            "Password hint: First 4 letters of your last name (UPPERCASE) + Last 4 digits of your SSN.\n\n"
            "Best regards,\nPayroll Team"
        )
        
        subject_template = st.text_input("Email Subject", value=default_subject)
        body_template = st.text_area("Email Body", value=default_body, height=160)
        st.caption("Available variables: {name}, {date}, {password}")

    # Section 1: Uploads
    st.subheader("1. Upload Documents")
    
    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            pdf_file = st.file_uploader(
                "Combined Paystub PDF",
                type=['pdf'],
                help="Multi-page PDF containing employee paystubs"
            )
            if pdf_file:
                st.caption(f"{pdf_file.name} ({(pdf_file.size / 1024):.1f} KB)")

        with col2:
            xlsx_file = st.file_uploader(
                "Employee Roster (Excel / CSV)",
                type=['xlsx', 'xls', 'csv'],
                help="File containing first_name, last_name, and email_address"
            )
            if xlsx_file:
                st.caption(f"{xlsx_file.name} ({(xlsx_file.size / 1024):.1f} KB)")

    # Reset processed results if uploaded files changed or were removed
    current_files_signature = (
        (pdf_file.name, pdf_file.size) if pdf_file else None,
        (xlsx_file.name, xlsx_file.size) if xlsx_file else None
    )

    if 'last_files_signature' not in st.session_state:
        st.session_state.last_files_signature = current_files_signature
        st.session_state.processed_results = None
    elif st.session_state.last_files_signature != current_files_signature:
        st.session_state.last_files_signature = current_files_signature
        st.session_state.processed_results = None

    if 'processed_results' not in st.session_state:
        st.session_state.processed_results = None
    if 'processed_date' not in st.session_state:
        st.session_state.processed_date = date_str

    # Section 2: Process Action
    st.subheader("2. Process & Encrypt")
    
    can_process = pdf_file is not None and xlsx_file is not None

    if not can_process:
        st.info("Upload both the Combined Paystub PDF and Employee Roster to enable processing.")
    else:
        # Strictly process ONLY when button is clicked
        if st.button("Process Paystubs", type="primary", width="stretch"):
            with st.status("Processing paystubs...", expanded=True) as status_box:
                try:
                    status_box.write("Reading employee roster...")
                    if xlsx_file.name.endswith('.csv'):
                        df = pd.read_csv(xlsx_file)
                    else:
                        df = pd.read_excel(xlsx_file)

                    employees = parse_employee_dataframe(df)
                    if not employees:
                        status_box.update(label="No valid employee records found.", state="error")
                        st.error("No valid employee records found in roster.")
                        return

                    status_box.write("Splitting pages, distinguishing JR/SR designations, and encrypting...")
                    temp_dir = os.path.join(tempfile.gettempdir(), f"paystubs_{date_str}")

                    results = process_paystub_pdf(
                        pdf_stream=pdf_file,
                        employees=employees,
                        date_string=date_str,
                        output_folder=temp_dir
                    )

                    st.session_state.processed_results = results
                    st.session_state.processed_date = date_str
                    st.session_state.temp_dir = temp_dir

                    status_box.update(label="Paystubs successfully processed and encrypted.", state="complete", expanded=False)
                    st.toast("Processing complete.")

                except Exception as e:
                    status_box.update(label=f"Error: {str(e)}", state="error")
                    st.error(f"Error processing files: {str(e)}")

    # Section 3: Output Area (Unified Downloads & Outlook Automation)
    # Displays ONLY after the button has been clicked and processed_results is populated
    if st.session_state.processed_results is not None:
        results = st.session_state.processed_results
        curr_date = st.session_state.processed_date

        total = len(results)
        successful = sum(1 for r in results if r['status'] == 'Encrypted')
        missing_pages = sum(1 for r in results if r['status'] == 'No Pages Found')
        missing_ssn = sum(1 for r in results if r['status'] == 'Missing SSN')

        st.subheader("3. Distribution & Downloads")

        # Summary Metrics
        with st.container(border=True):
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total in Roster", total)
            m2.metric("Successfully Encrypted", successful)
            m3.metric("No Pages Found", missing_pages)
            m4.metric("Missing SSN", missing_ssn)

        # Tabbed outputs: Unified Downloads Area & Email Drafts
        tab_downloads, tab_outlook = st.tabs(["Download Center", "Email Drafts (.eml)"])

        # TAB 1: Unified Download Center
        with tab_downloads:
            with st.container(border=True):
                st.markdown("#### Download All Files")
                st.write(f"Generate a single compressed archive containing all {successful} encrypted paystubs and an execution summary manifest.")
                
                zip_data = create_zip_archive(results, curr_date)
                st.download_button(
                    label=f"Download All Paystubs (ZIP - {successful} Files)",
                    data=zip_data,
                    file_name=f"paystubs_{curr_date}.zip",
                    mime="application/zip",
                    type="primary",
                    width="stretch"
                )

            with st.container(border=True):
                st.markdown("#### Download Individual Paystubs")
                st.write("Download specific employee paystubs individually:")

                search_query = st.text_input("Filter employees by name or email", placeholder="Type name to filter...").strip().lower()

                filtered_results = [
                    r for r in results
                    if search_query in r['name'].lower() or search_query in r['email'].lower()
                ] if search_query else results

                ready_files = [r for r in filtered_results if r['status'] == 'Encrypted' and r['pdf_bytes']]

                if not ready_files:
                    st.info("No matching encrypted paystubs available to download.")
                else:
                    col_count = 3
                    cols = st.columns(col_count)
                    for idx, r in enumerate(ready_files):
                        with cols[idx % col_count]:
                            with st.container(border=True):
                                st.write(f"**{r['name']}**")
                                st.caption(f"Pages: {r['pages_count']} | SSN: {r['ssn_last4']}")
                                st.download_button(
                                    label=f"Download PDF",
                                    data=r['pdf_bytes'],
                                    file_name=r['filename'],
                                    mime="application/pdf",
                                    key=f"dl_btn_{r['name']}_{idx}",
                                    width="stretch"
                                )

                st.divider()
                st.markdown("#### Complete Status Table")
                table_data = []
                for r in results:
                    table_data.append({
                        "Name": r['name'],
                        "Email": r['email'],
                        "Status": r['status'],
                        "Pages": r['pages_count'],
                        "SSN Last 4": r['ssn_last4'] if r['ssn_last4'] else "-",
                        "Password": r['password'] if r['password'] else "-",
                        "Filename": r['filename']
                    })

                df_display = pd.DataFrame(table_data)
                st.dataframe(
                    df_display,
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "Name": st.column_config.TextColumn("Employee Name", width="medium"),
                        "Email": st.column_config.TextColumn("Email Address", width="medium"),
                        "Status": st.column_config.TextColumn("Status", width="small"),
                        "Pages": st.column_config.NumberColumn("Pages", width="small"),
                        "SSN Last 4": st.column_config.TextColumn("SSN (Last 4)", width="small"),
                        "Password": st.column_config.TextColumn("Password Hint", width="small"),
                        "Filename": st.column_config.TextColumn("Target Filename", width="medium")
                    }
                )

        # TAB 2: Email Draft (.eml) Generation
        with tab_outlook:
            with st.container(border=True):
                st.markdown("#### Email Draft Generator (.eml)")
                st.write(
                    "Generate RFC-compliant `.eml` files with the `X-Unsent: 1` header. "
                    "Double-click any `.eml` file to open it as an **editable unsent draft** "
                    "in Microsoft Outlook, Apple Mail, or Thunderbird — with the encrypted paystub already attached."
                )

                col_prev0, col_prev1, col_prev2 = st.columns([1, 1, 2])
                with col_prev0:
                    st.text_input("From Preview", value=sender_email if sender_email else "(Not configured)", disabled=True)
                with col_prev1:
                    st.text_input("Subject Preview", value=subject_template.format(name="John Doe", date=curr_date, password="DOEJ1234"), disabled=True)
                with col_prev2:
                    st.text_area("Body Preview", value=body_template.format(name="John Doe", date=curr_date, password="DOEJ1234"), height=100, disabled=True)

                if st.button("Generate .eml Draft Files", type="primary", width="stretch"):
                    with st.status("Generating .eml draft files...", expanded=True) as draft_status:
                        eml_folder, draft_reports = generate_eml_drafts(
                            processed_files=results,
                            subject_template=subject_template,
                            body_template=body_template,
                            date_string=curr_date,
                            sender_email=sender_email
                        )

                        created_count = sum(1 for r in draft_reports if r['Status'] == 'EML Created')
                        draft_status.update(
                            label=f"Generated {created_count} .eml draft files.",
                            state="complete", expanded=False
                        )
                        st.toast(f"Successfully generated {created_count} .eml drafts.")

                        st.session_state.eml_folder = eml_folder
                        st.session_state.eml_reports = draft_reports

                # Render results if they exist in session state
                if st.session_state.get('eml_reports'):
                    draft_reports = st.session_state.eml_reports
                    eml_folder = st.session_state.eml_folder
                    created_count = sum(1 for r in draft_reports if r['Status'] == 'EML Created')

                    st.info(f"**{created_count}** `.eml` files saved to: `{eml_folder}`")

                    # Action buttons row
                    col_open, col_zip = st.columns(2)
                    with col_open:
                        if st.button("Open Folder in Finder", width="stretch"):
                            subprocess.Popen(['open', eml_folder])
                    with col_zip:
                        eml_zip = io.BytesIO()
                        with zipfile.ZipFile(eml_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for r in draft_reports:
                                if r['Status'] == 'EML Created' and r.get('Filename'):
                                    eml_path = os.path.join(eml_folder, r['Filename'])
                                    if os.path.exists(eml_path):
                                        zf.write(eml_path, r['Filename'])
                        eml_zip.seek(0)
                        st.download_button(
                            label=f"Download All Drafts (ZIP - {created_count} Files)",
                            data=eml_zip,
                            file_name=f"email_drafts_{curr_date}.zip",
                            mime="application/zip",
                            width="stretch"
                        )

                    st.dataframe(
                        pd.DataFrame(draft_reports),
                        width="stretch",
                        hide_index=True,
                        column_config={
                            "Name": st.column_config.TextColumn("Name"),
                            "Email": st.column_config.TextColumn("Email"),
                            "Status": st.column_config.TextColumn("Draft Status"),
                            "Filename": st.column_config.TextColumn(".eml File"),
                            "Error": st.column_config.TextColumn("Details")
                        }
                    )


if __name__ == "__main__":
    main()
