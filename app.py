import io
import os
import re
import zipfile
import subprocess
import tempfile
import datetime
import requests
import streamlit as st
import streamlit.components.v1 as components
from docx import Document

# ==========================================
# PAGE CONFIGURATION & CSS CUSTOMIZATION
# ==========================================
st.set_page_config(
    page_title="Document Generator | Fidelity Funding",
    page_icon="📄",
    layout="wide"
)

# Initialize Login State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# Hide Streamlit's "Press Enter to submit" helper text in form inputs
st.markdown("""
    <style>
    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# Injected directly into the parent window's head to prevent 'Enter' keypress from submitting forms
components.html("""
    <script>
    const parentDoc = window.parent.document;
    if (!parentDoc.getElementById('prevent-enter-submit')) {
        const script = parentDoc.createElement('script');
        script.id = 'prevent-enter-submit';
        script.type = 'text/javascript';
        script.innerHTML = `
            document.addEventListener('keydown', function(e) {
                if ((e.key === 'Enter' || e.keyCode === 13) && e.target.tagName === 'INPUT') {
                    e.preventDefault();
                    e.stopPropagation();
                    e.stopImmediatePropagation();
                }
            }, true);
        `;
        parentDoc.head.appendChild(script);
    }
    </script>
""", height=0, width=0)

# ==========================================
# STREAMLIT SECRETS / LINK CONFIGURATION
# ==========================================
def get_secret_link(key: str) -> str:
    """Fetch URL strictly from Streamlit Secrets."""
    try:
        return st.secrets["templates"][key]
    except KeyError:
        st.error(f"Missing required secret key: `templates.{key}`. Please configure it in your Streamlit secrets.")
        return ""
    except Exception as e:
        st.error(f"Error accessing secret `templates.{key}`: {e}")
        return ""

# Standard / Base Documents
DEED_URL = get_secret_link("deed_url")
DECL_URL = get_secret_link("decl_url")

# RPS Class URLs
TEMPLATE_CONFIG = {
    "RPS_CLASSES": {
        "RPS-L | 30k | 1yr | 9.0%": {"class_code": "RPS-L", "doc_url": get_secret_link("rps_l_url")},
        "RPS-N | 30k | 2yr | 9.5%": {"class_code": "RPS-N", "doc_url": get_secret_link("rps_n_url")},
        "RPS-S | 30k | 3yr | 10.0%": {"class_code": "RPS-S", "doc_url": get_secret_link("rps_s_url")},
        "RPS-M | 50k | 1yr | 10.0%": {"class_code": "RPS-M", "doc_url": get_secret_link("rps_m_url")},
        "RPS-O | 50k | 2yr | 10.5%": {"class_code": "RPS-O", "doc_url": get_secret_link("rps_o_url")},
        "RPS-T | 50k | 3yr | 11.0%": {"class_code": "RPS-T", "doc_url": get_secret_link("rps_t_url")},
        "RPS-F | 100k | 1yr | 11.0%": {"class_code": "RPS-F", "doc_url": get_secret_link("rps_f_url")},
        "RPS-P | 100k | 2yr | 11.5%": {"class_code": "RPS-P", "doc_url": get_secret_link("rps_p_url")},
        "RPS-U | 100k | 3yr | 12.0%": {"class_code": "RPS-U", "doc_url": get_secret_link("rps_u_url")},
        "RPS-G | 250k | 1yr | 11.0%": {"class_code": "RPS-G", "doc_url": get_secret_link("rps_g_url")},
        "RPS-Q | 250k | 2yr | 12.0%": {"class_code": "RPS-Q", "doc_url": get_secret_link("rps_q_url")},
        "RPS-W | 250k | 3yr | 12.5%": {"class_code": "RPS-W", "doc_url": get_secret_link("rps_w_url")},
        "RPS-H | 500k | 1yr | 12.0%": {"class_code": "RPS-H", "doc_url": get_secret_link("rps_h_url")},
        "RPS-R | 500k | 2yr | 12.5%": {"class_code": "RPS-R", "doc_url": get_secret_link("rps_r_url")},
        "RPS-X | 500k | 3yr | 13.0%": {"class_code": "RPS-X", "doc_url": get_secret_link("rps_x_url")},
        "RPS-AA | Profit Sharing": {"class_code": "RPS-AA", "doc_url": get_secret_link("rps_aa_url")},
        "RPS-Z | 1yr | 6.0% | Semi-Annually": {"class_code": "RPS-Z", "doc_url": get_secret_link("rps_z_url")},
        "RPS-K | 1yr | 9.0% | Monthly": {"class_code": "RPS-K", "doc_url": get_secret_link("rps_k_url")},
        "RPS-Y | 1yr | 10.0% | Monthly": {"class_code": "RPS-Y", "doc_url": get_secret_link("rps_y_url")},
        "RPS-V | 1yr | 15.0% | Monthly": {"class_code": "RPS-V", "doc_url": get_secret_link("rps_v_url")},
        "RPS-V | 1yr | 15.0% | Monthly (DIRECT)": {"class_code": "RPS-V (DIRECT)", "doc_url": "https://docs.google.com/document/d/1KsdlheVf7UZSZEx5W0iP2vTRlpz3t98d/"}
    }
}

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def extract_doc_id(url: str) -> str:
    """Extract Google Doc ID from full URL."""
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else ""

def fetch_google_doc_bytes(url: str) -> bytes:
    """Download Google Doc directly as DOCX bytes."""
    doc_id = extract_doc_id(url)
    if not doc_id:
        raise ValueError(f"Invalid Google Doc URL provided: {url}")
    export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
    response = requests.get(export_url)
    response.raise_for_status()
    return response.content

def replace_placeholders_in_paragraph(paragraph, replacements: dict):
    """Replace placeholder keys in paragraph while preserving text runs and formatting."""
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def process_docx_bytes(file_bytes: bytes, replacements: dict) -> bytes:
    """Load DOCX bytes, replace placeholders globally, and return modified DOCX bytes."""
    doc_io = io.BytesIO(file_bytes)
    doc = Document(doc_io)

    # 1. Process standard paragraphs
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, replacements)

    # 2. Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, replacements)
                    
    # 3. Process headers and footers (Important for templates)
    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                replace_placeholders_in_paragraph(p, replacements)
            for table in section.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_placeholders_in_paragraph(p, replacements)
        if section.footer:
            for p in section.footer.paragraphs:
                replace_placeholders_in_paragraph(p, replacements)
            for table in section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_placeholders_in_paragraph(p, replacements)

    output_io = io.BytesIO()
    doc.save(output_io)
    return output_io.getvalue()

def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    """Convert DOCX bytes to PDF bytes using headless LibreOffice."""
    with tempfile.TemporaryDirectory() as tmpdir:
        input_docx_path = os.path.join(tmpdir, "input.docx")
        with open(input_docx_path, "wb") as f:
            f.write(docx_bytes)
        
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            input_docx_path
        ]
        
        try:
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
            output_pdf_path = os.path.join(tmpdir, "input.pdf")
            if os.path.exists(output_pdf_path):
                with open(output_pdf_path, "rb") as f:
                    return f.read()
        except Exception as e:
            st.warning(f"PDF conversion warning: {e}")
    return None

# ==========================================
# USER TIERING / ACCESS CONTROL
# ==========================================

RESTRICTED_CLASSES = ["RPS-Z", "RPS-K", "RPS-Y", "RPS-V", "RPS-V (DIRECT)"]

# Filter available options based on login status
available_rps_options = []
for rps_name, rps_info in TEMPLATE_CONFIG["RPS_CLASSES"].items():
    code = rps_info["class_code"]
    # If not logged in, skip the restricted classes (Tier 2 behavior)
    if not st.session_state.logged_in and code in RESTRICTED_CLASSES:
        continue
    available_rps_options.append(rps_name)


# ==========================================
# USER INTERFACE
# ==========================================

st.title("Fidelity Funding RPS Subscription Document Generator")
st.markdown("Automated generation of **Subscription Agreements**, **Deed of Adherence**, and **Declaration Forms**.")

# Main Input Form
with st.form("doc_generation_form"):
    st.subheader("1. Client & Investment Details")
    col1, col2 = st.columns(2)
    
    with col1:
        client_name = st.text_input("Client Name", value="")
        client_nric = st.text_input("NRIC / Passport / Reg No", value="")
        client_email = st.text_input("Client Email", value="")
        client_contact = st.text_input("Contact Number", value="+60")
        client_address = st.text_area("Correspondence Address", value="")
        
    with col2:
        selected_rps = st.selectbox("Select RPS Class", available_rps_options)
        client_dob = st.date_input(
            "Date of Birth", 
            value=datetime.date(1990, 1, 1), 
            min_value=datetime.date(1900, 1, 1), 
            max_value=datetime.date.today()
        )
        client_nationality = st.text_input("Nationality", value="")
        client_occupation = st.text_input("Occupation", value="")
        investment_amt = st.text_input("Investment Amount (RM)", value="")
        agreement_date = st.date_input("Agreement Date", value=datetime.date.today())
        stamping = st.radio("Stamping", options=["Yes", "No"], index=0, horizontal=True)

    st.subheader("2. Bank Details")
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        bank_name = st.text_input("Bank Name", value="")
    with b_col2:
        bank_acc_name = st.text_input("Bank Account Name", value="")
    with b_col3:
        bank_acc_no = st.text_input("Bank Account Number", value="")

    st.subheader("3. Witness / Advisor Information")
    w_col1, w_col2 = st.columns(2)
    with w_col1:
        witness_name = st.text_input("Witness / Advisor Name", value="")
    with w_col2:
        witness_nric = st.text_input("Witness / Advisor NRIC / Passport", value="")

    st.subheader("4. Nominee Information")
    
    nom_tabs = st.tabs(["Nominee 1", "Nominee 2", "Nominee 3", "Nominee 4"])
    nom_data = {}
    
    for i, tab in enumerate(nom_tabs, start=1):
        with tab:
            n_col1, n_col2 = st.columns(2)
            with n_col1:
                nom_data[f"<<NOM{i}_NAME>>"] = st.text_input(f"Nominee {i} Name", key=f"n_name_{i}")
                nom_data[f"<<NOM{i}_NRIC>>"] = st.text_input(f"Nominee {i} NRIC/Passport", key=f"n_nric_{i}")
                nom_data[f"<<NOM{i}_RELATIONSHIP>>"] = st.text_input(f"Nominee {i} Relationship", key=f"n_rel_{i}")
            with n_col2:
                nom_data[f"<<NOM{i}_ADDRESS>>"] = st.text_area(f"Nominee {i} Address", key=f"n_addr_{i}")
                nom_data[f"<<NOM{i}_EMAIL>>"] = st.text_input(f"Nominee {i} Email", key=f"n_email_{i}")
                nom_data[f"<<NOM{i}_PERCENTAGE>>"] = st.text_input(f"Nominee {i} Percentage (%)", key=f"n_pct_{i}")

    submit_button = st.form_submit_button("Generate PDF Documents")

# ==========================================
# VALIDATION & PROCESSING
# ==========================================

if submit_button:
    errors = []

    # 1. Mandatory Client Name Check
    if not client_name.strip():
        errors.append("Please enter the Client Name before proceeding.")

    # 2. Client Email Validation
    email_regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if client_email.strip() and not re.match(email_regex, client_email.strip()):
        errors.append("Invalid Client Email format.")

    # 3. Contact Number Validation
    clean_contact = re.sub(r"[\s\-\(\)]", "", client_contact.strip())
    if clean_contact and not re.match(r"^\+?\d{7,15}$", clean_contact):
        errors.append("Invalid Contact Number format. Please enter a valid number (e.g., +60123456789).")

    # 4. Investment Amount Validation (Numbers only)
    clean_inv_amt = investment_amt.strip().replace(",", "")
    if clean_inv_amt and not re.match(r"^\d+(\.\d+)?$", clean_inv_amt):
        errors.append("Investment Amount must contain numbers only.")

    # 5. Bank Account Number Validation (Numbers only)
    clean_bank_acc = bank_acc_no.strip().replace("-", "").replace(" ", "")
    if clean_bank_acc and not clean_bank_acc.isdigit():
        errors.append("Bank Account Number must contain numbers only.")

    # 6. Nominee Validation Logic
    total_pct = 0.0
    pct_has_error = False

    for i in range(1, 5):
        n_name = nom_data.get(f"<<NOM{i}_NAME>>", "").strip()
        n_nric = nom_data.get(f"<<NOM{i}_NRIC>>", "").strip()
        n_rel  = nom_data.get(f"<<NOM{i}_RELATIONSHIP>>", "").strip()
        n_addr = nom_data.get(f"<<NOM{i}_ADDRESS>>", "").strip()
        n_email = nom_data.get(f"<<NOM{i}_EMAIL>>", "").strip()
        pct_val_str = nom_data.get(f"<<NOM{i}_PERCENTAGE>>", "").strip()

        # Nominee Email Check
        if n_email and not re.match(email_regex, n_email):
            errors.append(f"Nominee {i} Email format is invalid.")

        # Check if ANY field for this nominee has been filled
        is_nominee_filled = any([n_name, n_nric, n_rel, n_addr, n_email, pct_val_str])
        
        pct_val = 0.0
        if pct_val_str:
            try:
                pct_val = float(pct_val_str)
                total_pct += pct_val
            except ValueError:
                errors.append(f"Nominee {i} Percentage must be a valid number.")
                pct_has_error = True

        # Rule: If any field is filled, percentage must be > 0
        if is_nominee_filled and pct_val <= 0:
            errors.append(f"Nominee {i} details are filled out, so Nominee {i} Percentage must be greater than 0%.")
            pct_has_error = True

    # Rule: Total nominee percentage must equal 100% OR 0%
    if not pct_has_error:
        if abs(total_pct - 100.0) > 0.001 and abs(total_pct - 0.0) > 0.001:
            errors.append(f"The sum of all nominee percentages must equal 100% or 0%. (Current total: {total_pct:.2f}%)")

    # Display Errors or Proceed
    if errors:
        for err in errors:
            st.error(f"⚠️ {err}")
    else:
        with st.spinner("Fetching templates and generating PDFs..."):
            
            # --- Text Formatting Rules ---
            formatted_client_name = client_name.strip().upper()
            formatted_client_nric = client_nric.strip().upper()
            formatted_witness_name = witness_name.strip().upper()
            formatted_witness_nric = witness_nric.strip().upper()
            formatted_client_address = client_address.strip().title() if client_address.strip() else ""
            formatted_client_occupation = client_occupation.strip().title() if client_occupation.strip() else ""
            formatted_client_nationality = client_nationality.strip().title() if client_nationality.strip() else ""
            formatted_agreement_date = agreement_date.strftime("%d %b %Y")
            formatted_client_dob = client_dob.strftime("%d %b %Y") if client_dob else ""
            formatted_bank_name = bank_name.strip().upper()
            formatted_bank_acc_name = bank_acc_name.strip().title()

            if clean_inv_amt:
                try:
                    formatted_investment_amt = f"{float(clean_inv_amt):,.2f}"
                except ValueError:
                    formatted_investment_amt = investment_amt
            else:
                formatted_investment_amt = ""

            formatted_nom_data = {}
            for i in range(1, 5):
                n_name = nom_data.get(f"<<NOM{i}_NAME>>", "").strip()
                n_nric = nom_data.get(f"<<NOM{i}_NRIC>>", "").strip()
                n_rel  = nom_data.get(f"<<NOM{i}_RELATIONSHIP>>", "").strip()
                n_addr = nom_data.get(f"<<NOM{i}_ADDRESS>>", "").strip()
                n_email = nom_data.get(f"<<NOM{i}_EMAIL>>", "").strip()
                n_pct  = nom_data.get(f"<<NOM{i}_PERCENTAGE>>", "").strip()

                formatted_nom_data[f"<<NOM{i}_NAME>>"] = n_name.upper() if n_name else ""
                formatted_nom_data[f"<<NOM{i}_NRIC>>"] = n_nric.upper() if n_nric else ""
                formatted_nom_data[f"<<NOM{i}_RELATIONSHIP>>"] = n_rel.title() if n_rel else ""
                formatted_nom_data[f"<<NOM{i}_ADDRESS>>"] = n_addr.title() if n_addr else ""
                formatted_nom_data[f"<<NOM{i}_EMAIL>>"] = n_email
                formatted_nom_data[f"<<NOM{i}_PERCENTAGE>>"] = n_pct

            rps_sub_url = TEMPLATE_CONFIG["RPS_CLASSES"][selected_rps]["doc_url"]

            raw_replacements = {
                "<<CLIENT_NAME>>": formatted_client_name,
                "<<CLIENT_NRIC>>": formatted_client_nric,
                "<<CLIENT_EMAIL>>": client_email,
                "<<CLIENT_CONTACT>>": client_contact,
                "<<CLIENT_ADDRESS>>": formatted_client_address,
                "<<CLIENT_DOB>>": formatted_client_dob,
                "<<CLIENT_NATIONALITY>>": formatted_client_nationality,
                "<<CLIENT_OCCUPATION>>": formatted_client_occupation,
                "<<INVESTMENT_AMT>>": formatted_investment_amt,
                "<<DATE>>": formatted_agreement_date,
                "<<STAMPING>>": stamping,
                "<<BANK_NAME>>": formatted_bank_name,
                "<<BANK_ACC_NAME>>": formatted_bank_acc_name,
                "<<BANK_ACC_NO>>": bank_acc_no,
                "<<WITNESS_NAME>>": formatted_witness_name,
                "<<WITNESS_NRIC>>": formatted_witness_nric,
                "<<RPS-CLASS>>": TEMPLATE_CONFIG["RPS_CLASSES"][selected_rps]["class_code"],
                **formatted_nom_data
            }

            replacements = {k: (v.strip() if v and str(v).strip() else "  ") for k, v in raw_replacements.items()}

            clean_file_client_name = client_name.strip().upper()
            clean_file_date = agreement_date.strftime("%Y%m%d")
            rps_code = TEMPLATE_CONFIG["RPS_CLASSES"][selected_rps]["class_code"]

            fn_sub_base = f"[{clean_file_client_name}] 1. FF {rps_code} Subscription Agreement {clean_file_date}"
            fn_deed_base = f"[{clean_file_client_name}] 2. FF Deed of Adherence {clean_file_date}"
            fn_decl_base = f"[{clean_file_client_name}] 3. FF Declaration Form (Sophisticated Investor)"

            # Flag to bypass the Deed of Adherence for RPS-V (DIRECT)
            is_direct_class = selected_rps == "RPS-V | 1yr | 15.0% | Monthly (DIRECT)"

            try:
                # 1. Fetch Google Doc templates (skip DEED_URL if Direct Class)
                sub_docx_raw = fetch_google_doc_bytes(rps_sub_url)
                decl_docx_raw = fetch_google_doc_bytes(DECL_URL)

                # 2. Populate placeholders
                sub_docx = process_docx_bytes(sub_docx_raw, replacements)
                decl_docx = process_docx_bytes(decl_docx_raw, replacements)

                # 3. Convert to PDF
                sub_pdf = convert_docx_to_pdf(sub_docx)
                decl_pdf = convert_docx_to_pdf(decl_docx)
                
                # Conditionally handle Deed of Adherence
                deed_pdf = None
                if not is_direct_class:
                    deed_docx_raw = fetch_google_doc_bytes(DEED_URL)
                    deed_docx = process_docx_bytes(deed_docx_raw, replacements)
                    deed_pdf = convert_docx_to_pdf(deed_docx)

                st.success("✅ PDF Documents generated successfully!")

                # Prepare ZIP download package (PDFs ONLY)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    if sub_pdf and decl_pdf:
                        zip_file.writestr(f"{fn_sub_base}.pdf", sub_pdf)
                        zip_file.writestr(f"{fn_decl_base}.pdf", decl_pdf)
                        if deed_pdf and not is_direct_class:
                            zip_file.writestr(f"{fn_deed_base}.pdf", deed_pdf)
                    else:
                        st.error("Failed to generate PDFs. Please check system dependencies.")

                # Single Download Button for ZIP
                st.download_button(
                    label="📦 Download All PDFs (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"[{clean_file_client_name}]_Documents.zip",
                    mime="application/zip",
                    type="primary"
                )

                st.markdown("---")
                st.subheader("📥 Individual PDF Downloads")

                col_d1, col_d2, col_d3 = st.columns(3)

                with col_d1:
                    st.write("**1. Subscription Agreement**")
                    if sub_pdf:
                        st.download_button("Download PDF", sub_pdf, f"{fn_sub_base}.pdf", "application/pdf")

                with col_d2:
                    if not is_direct_class:
                        st.write("**2. Deed of Adherence**")
                        if deed_pdf:
                            st.download_button("Download PDF", deed_pdf, f"{fn_deed_base}.pdf", "application/pdf")

                with col_d3:
                    if is_direct_class:
                        st.write("**2. Declaration Form**")
                    else:
                        st.write("**3. Declaration Form**")
                    if decl_pdf:
                        st.download_button("Download PDF", decl_pdf, f"{fn_decl_base}.pdf", "application/pdf")

            except Exception as e:
                st.error(f"Error processing Google Docs: {e}")

# ==========================================
# ADMIN LOGIN EXPANDER (BOTTOM OF APP)
# ==========================================
st.markdown("<br><br><br>", unsafe_allow_html=True)
st.markdown("---")
with st.expander("🔐 Admin Login", expanded=False):
    if not st.session_state.logged_in:
        # We use a form here to capture Enter key presses naturally for the login 
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            login_submitted = st.form_submit_button("Login")
            
            if login_submitted:
                try:
                    secret_user = st.secrets["admin"]["username"]
                    secret_pass = st.secrets["admin"]["password"]
                except KeyError:
                    st.error("Admin credentials are not correctly configured in secrets.toml.")
                    secret_user, secret_pass = None, None
                
                if secret_user and secret_pass:
                    if username == secret_user and password == secret_pass:
                        st.session_state.logged_in = True
                        st.rerun()
                    else:
                        st.error("Invalid credentials")
    else:
        st.success("You are currently logged in as an Administrator with full Tier 1 access.")
        if st.button("Logout"):
            st.session_state.logged_in = False
            st.rerun()
