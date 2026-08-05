import streamlit as st
import streamlit.components.v1 as components
import re
from datetime import datetime
from io import BytesIO
import requests
from docx import Document
import os
import subprocess

# ==========================================
# 1. PREVENT ENTER KEY SUBMISSION (JS INJECTION)
# ==========================================
components.html("""
    <script>
    const parentDoc = window.parent.document;
    if (!parentDoc.getElementById('prevent-enter-submit')) {
        const script = parentDoc.createElement('script');
        script.id = 'prevent-enter-submit';
        script.type = 'text/javascript';
        script.innerHTML = `
            document.addEventListener('keydown', function(e) {
                // Target standard input fields and intercept the Enter key
                if ((e.key === 'Enter' || e.keyCode === 13) && e.target.tagName === 'INPUT') {
                    e.preventDefault();
                    e.stopPropagation();
                    e.stopImmediatePropagation(); // Crucial for overriding React's synthetic events
                }
            }, true); // 'true' means capture phase, catching it before React does
        `;
        parentDoc.head.appendChild(script);
    }
    </script>
""", height=0, width=0)

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
def download_google_doc(url):
    """Downloads a Google Doc as a DOCX file."""
    try:
        doc_id = re.search(r"/d/([a-zA-Z0-9-_]+)", url).group(1)
        export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
        response = requests.get(export_url)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        st.error("Failed to download template. Ensure the link is correct and set to 'Anyone with the link can view'.")
        return None

def replace_text_in_doc(doc, replacements):
    """Replaces placeholders in paragraphs and tables."""
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, str(val))
                
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if key in p.text:
                            p.text = p.text.replace(key, str(val))
    return doc

def convert_to_pdf(docx_path, output_dir):
    """Converts DOCX to PDF using LibreOffice (required in packages.txt)."""
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", 
        docx_path, "--outdir", output_dir
    ], check=True)

# ==========================================
# 3. STREAMLIT UI & FORM
# ==========================================
st.title("Agreement Document Generator")

# Google Doc Template
template_url = st.text_input("Google Doc Template URL")

st.markdown("---")
st.subheader("Client & Agreement Details")

# Calendar pop-up for Agreement Date
agreement_date = st.date_input("Agreement Date", value=datetime.today())

# Removed () from titles
client_name = st.text_input("Client Name")

# Email validation
client_email = st.text_input("Client Email")
email_is_valid = True
if client_email:
    # Regex to check basic email format
    if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", client_email):
        st.warning("⚠️ Please enter a valid email address.")
        email_is_valid = False

# Default to Malaysia Country Code
contact_number = st.text_input("Contact Number", value="+60")

# Number input only for Investment Amount
investment_amount = st.number_input("Investment Amount (RM)", min_value=0.0, step=1000.0, format="%.2f")

# Stamping Dropdown
stamping = st.selectbox("Stamping", options=["Yes", "No"])

st.markdown("---")
st.subheader("Nominee Percentages")
st.write("Ensure the sum of all 4 nominees equals exactly 100%.")

# Number inputs only for Nominee Percentages
col1, col2, col3, col4 = st.columns(4)
with col1:
    nom1 = st.number_input("Nominee 1 %", min_value=0, max_value=100, value=100, step=1)
with col2:
    nom2 = st.number_input("Nominee 2 %", min_value=0, max_value=100, value=0, step=1)
with col3:
    nom3 = st.number_input("Nominee 3 %", min_value=0, max_value=100, value=0, step=1)
with col4:
    nom4 = st.number_input("Nominee 4 %", min_value=0, max_value=100, value=0, step=1)

# Real-time percentage check
total_percentage = nom1 + nom2 + nom3 + nom4
if total_percentage != 100:
    st.error(f"Total percentage is {total_percentage}%. It must equal 100%.")
else:
    st.success("Percentages total 100%.")

st.markdown("---")
st.subheader("Signatories")

# Updated Witness field names
witness_name = st.text_input("Witness / Advisor Name")
witness_nric = st.text_input("Witness / Advisor NRIC / Passport")

# ==========================================
# 4. DOCUMENT GENERATION LOGIC
# ==========================================
st.markdown("---")
if st.button("Generate Document", type="primary"):
    
    # 1. Run final validations before generating
    if not template_url:
        st.error("Please provide a Google Doc Template URL.")
    elif total_percentage != 100:
        st.error("Cannot proceed: Nominee percentages must sum to 100%.")
    elif not email_is_valid:
        st.error("Cannot proceed: Invalid client email format.")
    else:
        with st.spinner("Downloading template and generating document..."):
            
            # Map inputs to document placeholders
            # Adjust these dictionary keys if your Google Doc placeholders are named differently
            replacements = {
                "<<AGREEMENT_DATE>>": agreement_date.strftime("%d %B %Y"),
                "<<CLIENT_NAME>>": client_name,
                "<<CLIENT_EMAIL>>": client_email,
                "<<CONTACT_NUMBER>>": contact_number,
                "<<INVESTMENT_AMOUNT>>": f"{investment_amount:,.2f}",
                "<<NOMINEE_1_PERCENTAGE>>": str(nom1),
                "<<NOMINEE_2_PERCENTAGE>>": str(nom2),
                "<<NOMINEE_3_PERCENTAGE>>": str(nom3),
                "<<NOMINEE_4_PERCENTAGE>>": str(nom4),
                "<<STAMPING>>": stamping,
                "<<WITNESS_NAME>>": witness_name,
                "<<WITNESS_NRIC>>": witness_nric
            }

            # Download and process the Docx
            docx_stream = download_google_doc(template_url)
            
            if docx_stream:
                doc = Document(docx_stream)
                doc = replace_text_in_doc(doc, replacements)
                
                # Save processed DOCX locally
                temp_docx_path = "Generated_Agreement.docx"
                doc.save(temp_docx_path)
                
                # Convert to PDF
                try:
                    convert_to_pdf(temp_docx_path, ".")
                    temp_pdf_path = "Generated_Agreement.pdf"
                    
                    # Provide Download Buttons
                    st.success("Document Generated Successfully!")
                    
                    colA, colB = st.columns(2)
                    with colA:
                        with open(temp_docx_path, "rb") as file:
                            st.download_button(
                                label="Download DOCX",
                                data=file,
                                file_name="Agreement.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    with colB:
                        if os.path.exists(temp_pdf_path):
                            with open(temp_pdf_path, "rb") as file:
                                st.download_button(
                                    label="Download PDF",
                                    data=file,
                                    file_name="Agreement.pdf",
                                    mime="application/pdf"
                                )
                except Exception as e:
                    st.error(f"PDF Conversion failed. Is LibreOffice installed? Error: {e}")
