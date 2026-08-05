import streamlit as st
import streamlit.components.v1 as components
import json
import re
import io
import time
import hashlib
import hmac
import requests
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

st.set_page_config(
    page_title="Fidelity Funding Investor Relations Management",
    layout="wide"
)

# Prevent 'Enter' keypress in input fields from automatically submitting forms
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
                    e.stopImmediatePropagation(); 
                }
            }, true); 
        `;
        parentDoc.head.appendChild(script);
    }
    </script>
""", height=0, width=0)

# Custom CSS for Corporate Design, Dark Mode Contrast, and Narrow Rows
st.markdown("""
    <style>
    [data-testid="InputInstructions"] { display: none !important; }
    [data-testid="stSidebar"] { background-color: #111827 !important; }
    [data-testid="stSidebar"] * { color: #F9FAFB !important; }
    [data-testid="stSidebar"] label { color: #E5E7EB !important; font-weight: 500; }
    [data-testid="stSidebar"] div.stButton > button {
        text-align: left !important; justify-content: flex-start !important;
        border-radius: 6px !important; padding: 6px 12px !important;
        font-size: 13px !important; font-weight: 500 !important;
        margin-bottom: 2px !important; width: 100% !important;
        transition: all 0.15s ease-in-out !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="secondary"] {
        border: none !important; background-color: transparent !important; color: #D1D5DB !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {
        background-color: #1F2937 !important; color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
        border: none !important; font-weight: 600 !important; background-color: #2563EB !important; color: #FFFFFF !important;
    }
    [data-testid="column"] div.stButton > button {
        padding: 1px 8px !important; font-size: 12px !important;
        min-height: 26px !important; height: 26px !important;
        line-height: 1.1 !important; border-radius: 4px !important;
    }
    [data-testid="column"] { padding-top: 1px !important; padding-bottom: 1px !important; }
    .block-container { padding-top: 2rem !important; padding-bottom: 1rem !important; }
    </style>
""", unsafe_allow_html=True)

# CATEGORY & RPS-CLASS STRUCTURE DEFINITION
RPS_STRUCTURE = {
    "30k+ Category": ["RPS-L | 30k | 1yr | 9.0%", "RPS-N | 30k | 2yr | 9.5%", "RPS-S | 30k | 3yr | 10.0%"],
    "50k+ Category": ["RPS-M | 50k | 1yr | 10.0%", "RPS-O | 50k | 2yr | 10.5%", "RPS-T | 50k | 3yr | 11.0%"],
    "100k+ Category": ["RPS-F | 100k | 1yr | 11.0%", "RPS-P | 100k | 2yr | 11.5%", "RPS-U | 100k | 3yr | 12.0%"],
    "250k+ Category": ["RPS-G | 250k | 1yr | 11.0%", "RPS-Q | 250k | 2yr | 12.0%", "RPS-W | 250k | 3yr | 12.5%"],
    "500k+ Category": ["RPS-H | 500k | 1yr | 12.0%", "RPS-R | 500k | 2yr | 12.5%", "RPS-X | 500k | 3yr | 13.0%"],
    "Other Category": ["RPS-AA | 1yr | Profit Sharing", "RPS-K | 1yr | 9.0%", "RPS-V | 1yr | 15.0%", "RPS-Y | 1yr | 10.0%", "RPS-Z | 1yr | 6.0%"]
}

ALL_RPS_CLASSES = []
CLASS_TO_CATEGORY = {}
for cat, classes in RPS_STRUCTURE.items():
    for full_cls in classes:
        ALL_RPS_CLASSES.append(full_cls)
        CLASS_TO_CATEGORY[full_cls] = cat

HEADERS = [
    "#", "Entry Creation Date", "Full Name", "Date of Birth", "Occupation",
    "Commencement Date", "Unit of Subscription", "RPS-Class", "Agency",
    "Advisor Name", "Advisor Email", "Stamping", "NRIC / Passport",
    "Email", "Contact Number", "Address", "Bank Name", "Account Name",
    "Bank Account No", "Drive Folder Link"
]

def retry_api_call(func, max_retries=5, initial_delay=2):
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "quota" in err_str.lower() or "RESOURCE_EXHAUSTED" in err_str:
                if attempt == max_retries - 1: raise e
                time.sleep(delay)
                delay *= 2
            else:
                if attempt == max_retries - 1: raise e
                time.sleep(1)

def fetch_safe_records(sheet, expected_headers):
    try:
        return sheet.get_all_records(expected_headers=expected_headers)
    except Exception:
        all_values = sheet.get_all_values()
        if not all_values or len(all_values) <= 1: return []
        records = []
        for row in all_values[1:]:
            row_dict = {}
            for i, header_name in enumerate(expected_headers):
                row_dict[header_name] = row[i] if i < len(row) else ""
            records.append(row_dict)
        return records

@st.cache_data(ttl=180, show_spinner=False)
def load_worksheet_data(sheet_id, sheet_name, _gc, headers):
    try:
        wb = _gc.open_by_key(sheet_id)
        ws = wb.worksheet(sheet_name)
        return fetch_safe_records(ws, headers)
    except gspread.exceptions.WorksheetNotFound:
        return []

def get_or_create_worksheet(wb, title, headers):
    try:
        ws = retry_api_call(lambda: wb.worksheet(title))
    except gspread.exceptions.WorksheetNotFound:
        ws = retry_api_call(lambda: wb.add_worksheet(title=title, rows="100", cols="22"))
        retry_api_call(lambda: ws.update('A1:T1', [headers]))
        return ws
    first_row = retry_api_call(lambda: ws.row_values(1))
    if not first_row or first_row[0] != headers[0]:
        retry_api_call(lambda: ws.update('A1:T1', [headers]))
    return ws

def get_or_create_drive_subfolder(drive_service, parent_folder_id, subfolder_name):
    query = f"name='{subfolder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = retry_api_call(lambda: drive_service.files().list(q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute())
    files = results.get('files', [])
    if files: return files[0]['id']
    file_metadata = {'name': subfolder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_folder_id]}
    folder = retry_api_call(lambda: drive_service.files().create(body=file_metadata, fields='id', supportsAllDrives=True).execute())
    return folder.get('id')

def get_nested_drive_class_folder(drive_service, main_folder_id, rps_class_full_name):
    category_name = CLASS_TO_CATEGORY.get(rps_class_full_name, "Other Category")
    cat_folder_id = get_or_create_drive_subfolder(drive_service, main_folder_id, category_name)
    class_folder_id = get_or_create_drive_subfolder(drive_service, cat_folder_id, rps_class_full_name)
    return class_folder_id

def generate_auth_token(username, password, exp_timestamp):
    secret_key = f"{username}:{password}".encode('utf-8')
    msg = f"{username}:{exp_timestamp}".encode('utf-8')
    return hmac.new(secret_key, msg, hashlib.sha256).hexdigest()

def check_password():
    if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        token, exp = st.query_params.get("token"), st.query_params.get("exp")
        if token and exp:
            try:
                if time.time() < float(exp):
                    if "auth" in st.secrets:
                        if hmac.compare_digest(token, generate_auth_token(st.secrets["auth"]["username"], st.secrets["auth"]["password"], exp)):
                            st.session_state["authenticated"] = True
                            return True
            except Exception: pass
    if not st.session_state["authenticated"]:
        st.title("Fidelity Funding IR Portal - Authentication")
        user_id = st.text_input("User ID")
        password = st.text_input("Password", type="password")
        if st.button("Log In", type="primary"):
            try:
                if user_id == st.secrets["auth"]["username"] and password == st.secrets["auth"]["password"]:
                    st.session_state["authenticated"] = True
                    exp_time = str(time.time() + 604800)
                    st.query_params["token"] = generate_auth_token(user_id, password, exp_time)
                    st.query_params["exp"] = exp_time
                    st.rerun()
                else: st.error("Invalid User ID or Password.")
            except Exception as e: st.error(f"Authentication Error: {str(e)}")
        return False
    return True

@st.cache_resource
def init_google_services():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        raw_json = st.secrets.get("gcp_service_account_json") or st.secrets.get("auth", {}).get("gcp_service_account_json")
        if raw_json: creds_dict = json.loads(raw_json, strict=False)
        elif "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict: creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        else: st.stop()
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        gc = gspread.authorize(creds)
        drive_service = build('drive', 'v3', credentials=creds)
        
        return gc, drive_service, creds
    except Exception as e:
        st.error(f"Connection Error: {str(e)}")
        st.stop()

def generate_dividend_schedule_pdf(gc, drive_service, creds, folder_id, full_name, dob, occupation, address, rps_class, unit, commence_date):
    TEMPLATE_SS_ID = '1o4hXAb59BFzAhgMqMfAmAMv3tVAbgvuRm8D1ULWag0Y'
    
    wb = retry_api_call(lambda: gc.open_by_key(TEMPLATE_SS_ID))
    try:
        ws = retry_api_call(lambda: wb.worksheet(rps_class))
    except gspread.exceptions.WorksheetNotFound:
        raise Exception(f"RPS Class '{rps_class}' tab not found in the Dividend Schedule template.")
    
    numeric_unit = float(unit)
    
    # 1. <<CLIENT_NAME>> always ALL UPPER CASE
    formatted_client_name = full_name.strip().upper()
    
    # 2. <<CLIENT_ADDRESS>> and <<CLIENT_OCCUPATION>> always Title Case
    formatted_client_address = address.strip().title()
    formatted_client_occupation = occupation.strip().title()
    
    # 3. <<INVESTMENT_AMT>> always comma for every 3 digits and ends with .00 (e.g. 100,000.00)
    formatted_investment_amt = f"{numeric_unit:,.2f}"
    
    # 4. <<CLIENT_DOB>> and <<DATE>> always DD MMM YYYY format
    if isinstance(dob, (date, datetime)):
        formatted_dob = dob.strftime("%d %b %Y")
    else:
        try: formatted_dob = datetime.strptime(str(dob), "%Y-%m-%d").strftime("%d %b %Y")
        except Exception: formatted_dob = str(dob)

    if isinstance(commence_date, (date, datetime)):
        formatted_date = commence_date.strftime("%d %b %Y")
    else:
        try: formatted_date = datetime.strptime(str(commence_date), "%Y-%m-%d").strftime("%d %b %Y")
        except Exception: formatted_date = str(commence_date)

    replacements = {
        "<<CLIENT_NAME>>": formatted_client_name,
        "<<CLIENT_ADDRESS>>": formatted_client_address,
        "<<CLIENT_OCCUPATION>>": formatted_client_occupation,
        "<<INVESTMENT_AMT>>": formatted_investment_amt,
        "<<CLIENT_DOB>>": formatted_dob,
        "<<DATE>>": formatted_date
    }

    # Replace placeholders dynamically in template sheet
    try:
        matched_cells = retry_api_call(lambda: ws.findall(re.compile(r'<<.*?>>')))
        if matched_cells:
            for cell in matched_cells:
                val = cell.value
                for tag, rep in replacements.items():
                    if tag in val:
                        val = val.replace(tag, rep)
                retry_api_call(lambda: ws.update_cell(cell.row, cell.col, val))
    except Exception: pass

    # Fallback/Direct Cell updates
    try:
        retry_api_call(lambda: ws.update_acell('E7', numeric_unit))
        retry_api_call(lambda: ws.update_acell('E8', formatted_date))
    except Exception: pass
    
    time.sleep(2)
    
    if not creds.valid:
        creds.refresh(Request())
        
    url = (f"https://docs.google.com/spreadsheets/d/{TEMPLATE_SS_ID}/export?"
           f"format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&"
           f"printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid={ws.id}")
           
    headers = {'Authorization': f'Bearer {creds.token}'}
    res = requests.get(url, headers=headers)
    
    if res.status_code != 200:
        raise Exception(f"Failed to export PDF from Google Drive API: {res.text}")
        
    pdf_name = f"[{formatted_client_name}] Dividend Schedule RM{formatted_investment_amt} {formatted_date}.pdf"
    media = MediaIoBaseUpload(io.BytesIO(res.content), mimetype='application/pdf')
    
    file_metadata = {
        'name': pdf_name, 
        'parents': [folder_id]
    }
    
    pdf_file = retry_api_call(lambda: drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink',
        supportsAllDrives=True
    ).execute())
    
    return pdf_file.get('webViewLink')

def render_customer_table(wb, drive_service, target_sheets, page_title, HEADERS, sheet_id, gc):
    st.subheader(page_title)
    all_records = []
    for sheet_name in target_sheets:
        try:
            records = load_worksheet_data(sheet_id, sheet_name, gc, HEADERS)
            for idx, r in enumerate(records): all_records.append({"sheet_name": sheet_name, "sheet_row": idx + 2, "data": r})
        except Exception as e: st.error(f"Error reading worksheet '{sheet_name}': {str(e)}")

    if all_records:
        with st.form("search_form", clear_on_submit=False):
            c_in, c_btn = st.columns([5, 1])
            with c_in:
                search_input = st.text_input("Search by #, Name, NRIC/Passport, or Agency", value=st.session_state.get("current_search", ""))
            with c_btn:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                submit_search = st.form_submit_button("Search", use_container_width=True)
            
            if submit_search:
                st.session_state["current_search"] = search_input
                st.rerun()

        search_query = st.session_state.get("current_search", "").lower().strip()
        filtered = [item for item in all_records if not search_query or any(search_query in str(item["data"].get(k, "")).lower() for k in ["#", "Full Name", "NRIC / Passport", "Agency"])]

        if filtered:
            h0, h1, h2, h3, h4, h5, h6, h7 = st.columns([0.4, 0.4, 1.1, 1.8, 1.1, 1.2, 1.1, 1.3])
            h0.markdown("**Action**"); h1.markdown("**#**"); h2.markdown("**Created**"); h3.markdown("**Full Name**"); h4.markdown("**Commencement**"); h5.markdown("**Class**"); h6.markdown("**Agency**"); h7.markdown("**Drive Link**")
            st.markdown("<hr style='margin: 4px 0px; border-top: 1px solid #ddd;'>", unsafe_allow_html=True)

            for item in filtered:
                sheet_name = item["sheet_name"]
                sheet_row_num = item["sheet_row"]
                row = item["data"]
                col_id = f"{sheet_name}_{sheet_row_num}"
                is_currently_editing = st.session_state.get("editing_target") == col_id
                
                c0, c1, c2, c3, c4, c5, c6, c7 = st.columns([0.4, 0.4, 1.1, 1.8, 1.1, 1.2, 1.1, 1.3])
                with c0:
                    if st.button("Close" if is_currently_editing else "Edit", key=f"edit_btn_{col_id}"):
                        st.session_state["editing_target"] = None if is_currently_editing else col_id
                        st.rerun()

                c1.write(str(row.get("#", ""))); c2.write(str(row.get("Entry Creation Date", ""))); c3.write(str(row.get("Full Name", ""))); c4.write(str(row.get("Commencement Date", ""))); c5.write(str(row.get("RPS-Class", sheet_name))); c6.write(str(row.get("Agency", "")))
                c7.markdown(f"[Drive Folder]({row.get('Drive Folder Link')})" if row.get("Drive Folder Link", "#") != "#" else "-")

                if is_currently_editing:
                    st.markdown("---")
                    with st.container():
                        st.markdown(f"#### Edit Record #{row.get('#')} ({sheet_name}) - {row.get('Full Name')}")
                        try: commence_default = datetime.strptime(str(row.get("Commencement Date", "")), "%Y-%m-%d").date()
                        except: commence_default = date.today()

                        try: dob_default = datetime.strptime(str(row.get("Date of Birth", "")), "%Y-%m-%d").date()
                        except: dob_default = date(1990, 1, 1)
                        
                        current_rps = row.get("RPS-Class", sheet_name)
                        with st.form(f"inline_form_{col_id}", clear_on_submit=False):
                            ce1, ce2 = st.columns(2)
                            with ce1:
                                edit_full_name = st.text_input("Full Name", value=str(row.get("Full Name", "")))
                                edit_dob = st.date_input("Date of Birth", value=dob_default, key=f"edit_dob_{col_id}")
                                edit_occupation = st.text_input("Occupation", value=str(row.get("Occupation", "")))
                                edit_commence_date = st.date_input("Commencement Date", value=commence_default)
                                edit_unit_sub = st.text_input("Unit of Subscription", value=str(row.get("Unit of Subscription", "")))
                                edit_rps_class = st.selectbox("RPS-Class", ALL_RPS_CLASSES, index=ALL_RPS_CLASSES.index(current_rps) if current_rps in ALL_RPS_CLASSES else 0)
                                edit_agency = st.text_input("Agency", value=str(row.get("Agency", "")))
                                edit_advisor_name = st.text_input("Advisor Name", value=str(row.get("Advisor Name", "")))
                                edit_advisor_email = st.text_input("Advisor Email", value=str(row.get("Advisor Email", "")))
                                edit_stamping = st.radio("Stamping", ["Yes", "No"], index=0 if row.get("Stamping", "Yes") == "Yes" else 1, horizontal=True, key=f"edit_st_{col_id}")
                            with ce2:
                                edit_nric_passport = st.text_input("NRIC / Passport", value=str(row.get("NRIC / Passport", "")))
                                edit_email = st.text_input("Customer Email", value=str(row.get("Email", "")))
                                edit_contact_number = st.text_input("Contact Number", value=str(row.get("Contact Number", "")))
                                edit_address = st.text_area("Address", value=str(row.get("Address", "")))
                                edit_bank_name = st.text_input("Bank Name", value=str(row.get("Bank Name", "")))
                                edit_account_name = st.text_input("Account Name", value=str(row.get("Account Name", "")))
                                edit_bank_account_no = st.text_input("Bank Account No", value=str(row.get("Bank Account No", "")))

                            st.write(f"**Drive Folder Location:** [Open Drive Folder]({row.get('Drive Folder Link')})")
                            edit_files = st.file_uploader("Upload Additional Documents", accept_multiple_files=True, key=f"uploader_{col_id}")
                            
                            if st.form_submit_button("Confirm and Save Changes", type="primary"):
                                try:
                                    with st.spinner("Updating Google Sheet and Drive directory..."):
                                        drive_url = row.get("Drive Folder Link", "")
                                        target_folder_id = None
                                        if drive_url: 
                                            match = re.search(r'folders/([a-zA-Z0-9_-]+)', drive_url)
                                            if match: target_folder_id = match.group(1)

                                        new_commence_date = edit_commence_date.strftime("%Y-%m-%d")
                                        new_dob_date = edit_dob.strftime("%Y-%m-%d")

                                        if (str(row.get("Full Name", "")) != edit_full_name.strip() or str(row.get("Commencement Date", "")) != new_commence_date) and target_folder_id:
                                            retry_api_call(lambda: drive_service.files().update(fileId=target_folder_id, body={'name': f"{row.get('#')}. {edit_full_name.strip() or 'Unnamed'} {new_commence_date}"}, supportsAllDrives=True).execute())

                                        if edit_files and target_folder_id:
                                            for f in edit_files:
                                                retry_api_call(lambda: drive_service.files().create(body={'name': f.name, 'parents': [target_folder_id]}, media_body=MediaIoBaseUpload(io.BytesIO(f.read()), mimetype=f.type), supportsAllDrives=True).execute())

                                        updated_row = [
                                            row.get("#"), row.get("Entry Creation Date"), edit_full_name.strip(), new_dob_date,
                                            edit_occupation.strip(), new_commence_date, edit_unit_sub.strip(), edit_rps_class,
                                            edit_agency.strip(), edit_advisor_name.strip(), edit_advisor_email.strip(), edit_stamping,
                                            edit_nric_passport.strip(), edit_email.strip(), edit_contact_number.strip(), edit_address.strip(),
                                            edit_bank_name.strip(), edit_account_name.strip(), edit_bank_account_no.strip(), drive_url
                                        ]
                                        retry_api_call(lambda: get_or_create_worksheet(wb, sheet_name, HEADERS).update(f'A{sheet_row_num}:T{sheet_row_num}', [updated_row]))
                                        st.cache_data.clear()
                                        st.session_state["editing_target"] = None
                                        st.rerun()
                                except Exception as e: st.error(f"Failed to update entry: {str(e)}")

                        if st.button("Cancel Editing", key=f"cancel_{col_id}"):
                            st.session_state["editing_target"] = None
                            st.rerun()
                st.markdown("<hr style='margin: 1px 0px; border-top: 1px solid #f0f0f0;'>", unsafe_allow_html=True)
        else: st.info("No matching customer records found.")
    else: st.info("No customer records found.")

# MAIN APP ENTRY
if check_password():
    gc, drive_service, creds = init_google_services()
    
    try:
        sheet_id = st.secrets["GOOGLE_SHEET_ID"].strip()
        main_drive_folder_id = st.secrets["GOOGLE_DRIVE_FOLDER_ID"].strip()
        wb = gc.open_by_key(sheet_id)
    except Exception as e:
        st.error(f"Failed to open Google Sheet/Drive Folder ID. Details: {str(e)}")
        st.stop()

    if "nav_choice" not in st.session_state: st.session_state["nav_choice"] = "Create New Entry"
    if "current_search" not in st.session_state: st.session_state["current_search"] = ""
    if "editing_target" not in st.session_state: st.session_state["editing_target"] = None

    st.sidebar.markdown("### Navigation")
    
    for key_name, label in [("Create New Entry", "Create New Entry"), ("Customer Search", "Customer Search")]:
        if st.sidebar.button(label, key=f"nav_btn_{key_name}", use_container_width=True, type="primary" if st.session_state["nav_choice"] == key_name else "secondary"):
            st.session_state["nav_choice"] = key_name
            st.session_state["current_search"] = ""
            st.session_state["editing_target"] = None
            st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.markdown("**RPS Class Lists**")

    for cat_name, class_list in RPS_STRUCTURE.items():
        with st.sidebar.expander(cat_name, expanded=False):
            for full_cls_str in class_list:
                if st.button(full_cls_str, key=f"nav_btn_{full_cls_str}", use_container_width=True, type="primary" if st.session_state["nav_choice"] == full_cls_str else "secondary"):
                    st.session_state["nav_choice"] = full_cls_str
                    st.session_state["current_search"] = ""
                    st.session_state["editing_target"] = None
                    st.rerun()

    nav_choice = st.session_state["nav_choice"]
    st.sidebar.markdown("<br><hr>", unsafe_allow_html=True)
    if st.sidebar.button("Refresh Data", use_container_width=True, type="secondary"):
        st.cache_data.clear()
        st.rerun()
    if st.sidebar.button("Log Out", use_container_width=True, type="secondary"):
        st.session_state["authenticated"] = False
        st.query_params.clear()
        st.rerun()

    st.title("Fidelity Funding Investor Relations Management")

    if nav_choice == "Create New Entry":
        st.subheader("Add New Customer Record")

        if "create_form_key_suffix" not in st.session_state:
            st.session_state["create_form_key_suffix"] = 0

        if "create_success_info" in st.session_state:
            info = st.session_state.pop("create_success_info")
            st.success(f"Record #{info['number']} for '{info['name']}' successfully saved!")
            if info.get("pdf_success"):
                st.success("Dividend Schedule PDF successfully generated and saved to the drive folder.")
            elif info.get("pdf_error"):
                st.warning(f"Note: Profile was created, but the PDF failed to generate. Detail: {info['pdf_error']}")
            st.markdown(f"📁 **Google Drive Folder:** [Open Customer Folder in Google Drive]({info['folder_url']})")

        form_suffix = st.session_state["create_form_key_suffix"]

        with st.form(f"create_customer_form_{form_suffix}", clear_on_submit=False):
            col1, col2 = st.columns(2)
            with col1:
                full_name = st.text_input("Full Name", key=f"fn_{form_suffix}")
                dob = st.date_input("Date of Birth", value=date(1990, 1, 1), key=f"dob_{form_suffix}")
                occupation = st.text_input("Occupation", key=f"occ_{form_suffix}")
                commence_date = st.date_input("Commencement Date", value=date.today(), key=f"cd_{form_suffix}")
                unit_subscription = st.text_input("Unit of Subscription", key=f"us_{form_suffix}")
                class_options = ["Select One Class"] + ALL_RPS_CLASSES
                selected_class_full = st.selectbox("RPS-Class", class_options, key=f"sc_{form_suffix}")
                agency = st.text_input("Agency", key=f"ag_{form_suffix}")
                advisor_name = st.text_input("Advisor Name", key=f"an_{form_suffix}")
                advisor_email = st.text_input("Advisor Email", key=f"ae_{form_suffix}")
                stamping = st.radio("Stamping", ["Yes", "No"], horizontal=True, key=f"st_{form_suffix}")
            with col2:
                nric_passport = st.text_input("NRIC / Passport", key=f"np_{form_suffix}")
                email = st.text_input("Customer Email", key=f"em_{form_suffix}")
                contact_number = st.text_input("Contact Number", value="+60", key=f"cn_{form_suffix}")
                address = st.text_area("Address", key=f"addr_{form_suffix}")
                bank_name = st.text_input("Bank Name", key=f"bn_{form_suffix}")
                account_name = st.text_input("Account Name", key=f"aname_{form_suffix}")
                bank_account_no = st.text_input("Bank Account No", key=f"bno_{form_suffix}")

            st.markdown("---")
            files = st.file_uploader("Upload Customer Documents", accept_multiple_files=True, key=f"files_{form_suffix}")
            submit = st.form_submit_button("Confirm and Save Entry", type="primary")

            if submit:
                errors = []
                if selected_class_full == "Select One Class": errors.append("Please select a valid RPS-Class.")
                try: float(unit_subscription.strip())
                except ValueError: errors.append("Unit of Subscription must be a valid numeric value.")

                if errors:
                    for err in errors: st.error(f"Validation Error: {err}")
                else:
                    try:
                        rps_class_full_name = selected_class_full
                        with st.spinner(f"Creating profile folder and saving entries..."):
                            creation_date_str = date.today().strftime("%Y-%m-%d")
                            commence_date_str = commence_date.strftime("%Y-%m-%d")
                            dob_str = dob.strftime("%Y-%m-%d")
                            
                            target_ws = get_or_create_worksheet(wb, rps_class_full_name, HEADERS)
                            next_row = max(len(retry_api_call(lambda: target_ws.col_values(1))) + 1, 2)
                            entry_number = next_row - 1

                            class_drive_folder_id = get_nested_drive_class_folder(drive_service, main_drive_folder_id, rps_class_full_name)
                            folder_name = f"{entry_number}. {full_name.strip() or 'Unnamed'} {commence_date_str}"
                            
                            folder = retry_api_call(lambda: drive_service.files().create(body={'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [class_drive_folder_id]}, fields='id, webViewLink', supportsAllDrives=True).execute())
                            folder_id = folder.get('id')
                            folder_url = folder.get('webViewLink')

                            if files:
                                for f in files:
                                    retry_api_call(lambda: drive_service.files().create(body={'name': f.name, 'parents': [folder_id]}, media_body=MediaIoBaseUpload(io.BytesIO(f.read()), mimetype=f.type), supportsAllDrives=True).execute())

                        # Trigger PDF generation with formatted fields
                        pdf_success = False
                        pdf_error = None
                        try:
                            with st.spinner("Generating and uploading the Dividend Schedule PDF..."):
                                pdf_link = generate_dividend_schedule_pdf(
                                    gc, drive_service, creds, 
                                    folder_id, 
                                    full_name.strip() or "Unnamed",
                                    dob,
                                    occupation.strip(),
                                    address.strip(),
                                    rps_class_full_name, 
                                    unit_subscription.strip(), 
                                    commence_date
                                )
                                pdf_success = True
                        except Exception as e:
                            pdf_error = str(e)

                        # Final Database Write
                        new_row = [
                            entry_number, creation_date_str, full_name.strip(), dob_str,
                            occupation.strip(), commence_date_str, unit_subscription.strip(), rps_class_full_name,
                            agency.strip(), advisor_name.strip(), advisor_email.strip(), stamping,
                            nric_passport.strip(), email.strip(), contact_number.strip(), address.strip(),
                            bank_name.strip(), account_name.strip(), bank_account_no.strip(), folder_url
                        ]
                        retry_api_call(lambda: target_ws.update(f'A{next_row}:T{next_row}', [new_row]))
                        
                        # Store success details & reset form fields
                        st.session_state["create_success_info"] = {
                            "number": entry_number,
                            "name": full_name.strip() or "New Entry",
                            "folder_url": folder_url,
                            "pdf_success": pdf_success,
                            "pdf_error": pdf_error
                        }
                        st.session_state["create_form_key_suffix"] += 1
                        st.cache_data.clear()
                        st.rerun()

                    except Exception as e:
                        st.error(f"Error saving entry: {str(e)}")

    elif nav_choice == "Customer Search":
        render_customer_table(wb, drive_service, ALL_RPS_CLASSES, "Customer Search (All Categories)", HEADERS, sheet_id, gc)

    elif nav_choice in ALL_RPS_CLASSES:
        render_customer_table(wb, drive_service, [nav_choice], f"{nav_choice} Customer List ({CLASS_TO_CATEGORY.get(nav_choice, '')})", HEADERS, sheet_id, gc)
