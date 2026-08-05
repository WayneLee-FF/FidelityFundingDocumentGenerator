import re
import io
from datetime import date
import requests
import streamlit as st
import streamlit.components.v1 as components
from docx import Document

# Set page configuration
st.set_page_config(page_title="Document Generator", page_icon="📄", layout="centered")

# ==============================================================================
# 1. PREVENT ENTER KEY FORM SUBMISSION INJECTION
# ==============================================================================
components.html(
    """
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
""",
    height=0,
    width=0,
)

st.title("📄 Agreement Document Generator")
st.write("Fill in the form details below to generate your customized document.")

# Helper function for Email Regex validation
EMAIL_REGEX = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"


# Helper function to replace text in docx document
def replace_placeholders(doc, replacements):
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


# ==============================================================================
# 2. FORM INTERFACE
# ==============================================================================
with st.form("agreement_form"):

    st.subheader("Client Details")
    client_name = st.text_input("Client Name")
    client_nric = st.text_input("Client NRIC / Passport")

    # Client Email with format checking
    client_email = st.text_input("Client Email")

    # Default to Malaysia Country Code (+60)
    client_contact = st.text_input("Contact Number", value="+60 ")

    st.divider()
    st.subheader("Agreement Settings")

    # Calendar Date Picker
    agreement_date = st.date_input("Agreement Date", value=date.today())

    # Investment Amount (Number input only)
    investment_amount = st.number_input(
        "Investment Amount", min_value=0.0, step=1000.0, format="%.2f"
    )

    # Stamping Selection (Yes / No)
    stamping = st.radio("Stamping Required", options=["Yes", "No"], horizontal=True)

    st.divider()
    st.subheader("Nominee Percentages (Must total 100%)")

    col1, col2 = st.columns(2)
    with col1:
        nominee_1_name = st.text_input("Nominee 1 Name")
        nominee_1_pct = st.number_input(
            "Nominee 1 Percentage (%)", min_value=0.0, max_value=100.0, value=25.0, step=1.0
        )

        nominee_2_name = st.text_input("Nominee 2 Name")
        nominee_2_pct = st.number_input(
            "Nominee 2 Percentage (%)", min_value=0.0, max_value=100.0, value=25.0, step=1.0
        )

    with col2:
        nominee_3_name = st.text_input("Nominee 3 Name")
        nominee_3_pct = st.number_input(
            "Nominee 3 Percentage (%)", min_value=0.0, max_value=100.0, value=25.0, step=1.0
        )

        nominee_4_name = st.text_input("Nominee 4 Name")
        nominee_4_pct = st.number_input(
            "Nominee 4 Percentage (%)", min_value=0.0, max_value=100.0, value=25.0, step=1.0
        )

    st.divider()
    # Modified UI Titles: Witness / Advisor
    st.subheader("Witness / Advisor Details")
    witness_name = st.text_input("Witness / Advisor Name")
    witness_nric = st.text_input("Witness / Advisor NRIC / Passport")

    submit_button = st.form_submit_button("Generate Document")

# ==============================================================================
# 3. VALIDATION AND SUBMISSION LOGIC
# ==============================================================================
if submit_button:
    errors = []

    # 1. Email Format Check (If filled)
    if client_email.strip() and not re.match(EMAIL_REGEX, client_email.strip()):
        errors.append("Invalid Email address format. Please check your input.")

    # 2. Nominee Percentage Sum Check (Must equal 100%)
    total_pct = nominee_1_pct + nominee_2_pct + nominee_3_pct + nominee_4_pct
    if abs(total_pct - 100.0) > 0.001:
        errors.append(f"The sum of all 4 nominee percentages must equal 100%. (Current total: {total_pct:.1f}%)")

    # Display errors if any rules fail
    if errors:
        for err in errors:
            st.error(err)
    else:
        st.success("Form validated successfully! Generating document...")

        # Format mappings for replacing placeholders in document template
        placeholders = {
            "<<CLIENT_NAME>>": client_name,
            "<<CLIENT_NRIC>>": client_nric,
            "<<CLIENT_EMAIL>>": client_email,
            "<<CLIENT_CONTACT>>": client_contact,
            "<<AGREEMENT_DATE>>": agreement_date.strftime("%d %B %Y"),
            "<<INVESTMENT_AMOUNT>>": f"{investment_amount:,.2f}",
            "<<STAMPING>>": stamping,  # Replaces with "Yes" or "No"
            "<<NOMINEE_1_NAME>>": nominee_1_name,
            "<<NOMINEE_1_PCT>>": f"{nominee_1_pct}%",
            "<<NOMINEE_2_NAME>>": nominee_2_name,
            "<<NOMINEE_2_PCT>>": f"{nominee_2_pct}%",
            "<<NOMINEE_3_NAME>>": nominee_3_name,
            "<<NOMINEE_3_PCT>>": f"{nominee_3_pct}%",
            "<<NOMINEE_4_NAME>>": nominee_4_name,
            "<<NOMINEE_4_PCT>>": f"{nominee_4_pct}%",
            "<<WITNESS_NAME>>": witness_name,
            "<<WITNESS_NRIC>>": witness_nric,
        }

        # Example replacement execution (uncomment when integrating template download):
        # doc = Document(template_path)
        # replace_placeholders(doc, placeholders)
