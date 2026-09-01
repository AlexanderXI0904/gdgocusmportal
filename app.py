import json
import os
import requests
import msal
import atexit
from pathlib import Path
from flask import Flask, flash, redirect, render_template, request, session, url_for
from openpyxl import load_workbook

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
CONFIG_DIR = BASE_DIR / "config"
CONFIG_DIR.mkdir(exist_ok=True)
SETTINGS_FILE = CONFIG_DIR / "settings.json"
CACHE_FILE = CONFIG_DIR / "msal_cache.json"

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "gdgoc28usm08email26sender2501key")

# ============================================================
# AUTOMATIC CLEANUP
# ============================================================
def cleanup_uploads():
    """Wipes all temporary files in the uploads folder on shutdown/startup."""
    if UPLOAD_DIR.exists():
        for file in UPLOAD_DIR.glob("*"):
            try:
                if file.is_file():
                    file.unlink()
            except Exception:
                pass

atexit.register(cleanup_uploads)
cleanup_uploads()

DEFAULT_SETTINGS = {
    "client_id": "",
    "tenant_id": "",
    "cc_emails": [],
}

# ============================================================
# SETTINGS & CACHE (CLI-STYLE PERSISTENCE)
# ============================================================
def load_settings():
    if not SETTINGS_FILE.exists():
        save_settings(DEFAULT_SETTINGS.copy())
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            settings = json.load(f)
            settings.setdefault("client_id", "")
            settings.setdefault("tenant_id", "")
            settings.setdefault("cc_emails", [])
            return settings
    except Exception:
        return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)

def load_cache():
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, "r") as f:
                cache.deserialize(f.read())
        except Exception:
            pass
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

# ============================================================
# MICROSOFT AUTH
# ============================================================
def get_msal_app(cache=None):
    settings = load_settings()
    client_id = settings.get("client_id")
    tenant_id = settings.get("tenant_id")
    
    if not client_id or not tenant_id:
        return None
        
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

def get_access_token():
    """Silently fetches a token using the saved file cache, bypassing login if valid."""
    cache = load_cache()
    app = get_msal_app(cache)
    if not app:
        return None
        
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(["https://graph.microsoft.com/Mail.Send"], account=accounts[0])
        if result and "access_token" in result:
            save_cache(cache)
            return result["access_token"]
            
    return None

# ============================================================
# EXCEL HELPERS
# ============================================================
def get_workbook():
    filename = session.get("excel_filename")
    if not filename:
        return None

    filepath = UPLOAD_DIR / filename
    if not filepath.exists():
        session.pop("excel_filename", None)
        session.pop("original_filename", None)
        session.pop("sheet_name", None)
        return None

    return load_workbook(filepath, data_only=True)

def get_columns(sheet_name=None):
    wb = get_workbook()
    if not wb:
        return []
        
    sheet_name = sheet_name or session.get("sheet_name")
    if not sheet_name or sheet_name not in wb.sheetnames:
        return []

    ws = wb[sheet_name]
    columns = [
        str(cell.value).strip() 
        for cell in ws[1] 
        if cell.value is not None and str(cell.value).strip()
    ]
    
    wb.close()
    return columns

def get_recipients(email_column):
    wb = get_workbook()
    if not wb:
        return []
        
    sheet_name = session.get("sheet_name")
    if not sheet_name or sheet_name not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[sheet_name]
    headers = {
        str(cell.value).strip(): index 
        for index, cell in enumerate(ws[1], start=1) 
        if cell.value is not None
    }

    if email_column not in headers:
        wb.close()
        return []

    col_idx = headers[email_column]
    recipients = []
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) >= col_idx:
            val = row[col_idx - 1]
            if val:
                recipients.append(str(val).strip())

    wb.close()
    return recipients

def process_dynamic_content(text, row_data):
    if not text:
        return text
    for column, value in row_data.items():
        placeholder = f"{{{{{column}}}}}"
        replacement = "" if value is None else str(value)
        text = text.replace(placeholder, replacement)
    return text

def get_row_data(ws, row_number):
    data = {}
    for index, cell in enumerate(ws[row_number], start=1):
        header_cell = ws.cell(row=1, column=index)
        if header_cell.value is not None:
            col_name = str(header_cell.value).strip()
            if col_name:
                data[col_name] = cell.value
    return data

# ============================================================
# ROUTES
# ============================================================
@app.route("/")
def index():
    settings = load_settings()
    authenticated = bool(get_access_token())
    sheets = []
    columns = []

    wb = get_workbook()
    if wb:
        sheets = wb.sheetnames
        selected_sheet = session.get("sheet_name")
        if selected_sheet and selected_sheet in sheets:
            ws = wb[selected_sheet]
            columns = [
                str(c.value).strip() for c in ws[1] 
                if c.value is not None and str(c.value).strip()
            ]
        wb.close()

    return render_template(
        "index.html",
        page=1,
        settings=settings,
        authenticated=authenticated,
        sheets=sheets,
        selected_sheet=session.get("sheet_name"),
        columns=columns,
        excel_filename=session.get("original_filename") or session.get("excel_filename")
    )

@app.route("/save-ms-settings", methods=["POST"])
def save_ms_settings():
    settings = load_settings()
    settings["client_id"] = request.form.get("client_id", "").strip()
    settings["tenant_id"] = request.form.get("tenant_id", "").strip()
    
    save_settings(settings)
    if CACHE_FILE.exists():
        CACHE_FILE.unlink()
        
    flash("Microsoft settings saved.", "success")
    return redirect(url_for("index"))

@app.route("/authenticate")
def authenticate():
    msal_app = get_msal_app()
    if not msal_app:
        flash("Please save your Client ID and Tenant ID first.", "error")
        return redirect(url_for("index"))

    scopes = ["https://graph.microsoft.com/Mail.Send"]
    flow = msal_app.initiate_device_flow(scopes=scopes)
    
    if "user_code" not in flow:
        flash("Unable to start Microsoft authentication.", "error")
        return redirect(url_for("index"))

    session["device_flow"] = flow
    return render_template("device_login.html", flow=flow)

@app.route("/complete-authentication", methods=["POST"])
def complete_authentication():
    cache = load_cache()
    msal_app = get_msal_app(cache)
    flow = session.get("device_flow")
    
    if not msal_app or not flow:
        flash("Authentication session expired. Please try again.", "error")
        return redirect(url_for("index"))

    result = msal_app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        flash(result.get("error_description", "Authentication failed."), "error")
        return redirect(url_for("index"))

    save_cache(cache)
    session.pop("device_flow", None)
    flash("Microsoft authentication successful.", "success")
    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    if CACHE_FILE.exists():
        CACHE_FILE.unlink()
    session.pop("device_flow", None)
    flash("Signed out successfully.", "success")
    return redirect(url_for("index"))

@app.route("/upload-excel", methods=["POST"])
def upload_excel():
    if not get_access_token():
        flash("Please authenticate with Microsoft first.", "error")
        return redirect(url_for("index"))

    file = request.files.get("file")
    
    if not file or not file.filename:
        flash("Please select a valid Excel file.", "error")
        return redirect(url_for("index"))

    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in {".xlsx", ".xlsm"}:
        flash("Only .xlsx and .xlsm files are supported.", "error")
        return redirect(url_for("index"))

    for old_file in UPLOAD_DIR.glob("current_workbook*"):
        try:
            old_file.unlink()
        except OSError:
            pass

    filename = f"current_workbook{file_ext}"
    file.save(UPLOAD_DIR / filename)

    session.pop("sheet_name", None)
    session.pop("email_column", None)
    session["excel_filename"] = filename
    session["original_filename"] = file.filename
    
    flash("Excel file uploaded.", "success")
    return redirect(url_for("index"))

@app.route("/select-sheet", methods=["POST"])
def select_sheet():
    sheet_name = request.form.get("sheet_name", "").strip()
    wb = get_workbook()
    
    if not wb or sheet_name not in wb.sheetnames:
        if wb:
            wb.close()
        flash("Invalid worksheet selected.", "error")
        return redirect(url_for("index"))

    session["sheet_name"] = sheet_name
    wb.close()
    flash(f"Worksheet '{sheet_name}' selected.", "success")
    return redirect(url_for("index"))

@app.route("/save-cc", methods=["POST"])
def save_cc():
    email = request.form.get("email", "").strip().lower()
    if not email:
        return {"success": False, "message": "No email supplied."}, 400

    settings = load_settings()
    cc_emails = settings.setdefault("cc_emails", [])
    
    if email not in cc_emails:
        cc_emails.append(email)
        save_settings(settings)
        
    return {"success": True, "email": email, "cc_emails": cc_emails}

@app.route("/editor")
def editor():
    if not get_access_token():
        flash("Please authenticate first.", "error")
        return redirect(url_for("index"))
    
    sheet_name = session.get("sheet_name")
    if not sheet_name or not session.get("excel_filename"):
        flash("Please select a worksheet first.", "error")
        return redirect(url_for("index"))

    return render_template(
        "index.html",
        page=2,
        columns=get_columns(sheet_name),
        cc_emails=load_settings().get("cc_emails", []),
        sheet_name=sheet_name
    )

@app.route("/preview", methods=["POST"])
def preview():
    subject = request.form.get("subject", "").strip()
    importance = request.form.get("importance", "normal")
    html = request.form.get("html", "")
    email_col = request.form.get("email_column", "").strip()
    cc = request.form.getlist("cc")

    if not subject or not email_col:
        flash("Subject and email column are required.", "error")
        return redirect(url_for("editor"))

    wb = get_workbook()
    sheet_name = session.get("sheet_name")
    
    if not wb or not sheet_name or sheet_name not in wb.sheetnames:
        if wb:
            wb.close()
        flash("Worksheet no longer available.", "error")
        return redirect(url_for("index"))

    ws = wb[sheet_name]
    headers = {
        str(cell.value).strip(): idx 
        for idx, cell in enumerate(ws[1], start=1) if cell.value is not None
    }

    if email_col not in headers:
        wb.close()
        flash("Selected email column no longer exists.", "error")
        return redirect(url_for("editor"))

    email_idx = headers[email_col]
    
    preview_data = []
    recipients = []

    for row_number in range(2, ws.max_row + 1):
        email_val = ws.cell(row=row_number, column=email_idx).value
        if not email_val:
            continue
            
        email = str(email_val).strip()
        if not email:
            continue

        row_data = get_row_data(ws, row_number)
        
        clean_data = {k: ("" if v is None else str(v)) for k, v in row_data.items()}
        
        recipients.append(email)
        preview_data.append({
            "email": email,
            "data": clean_data
        })

    wb.close()

    if not recipients:
        flash("No valid recipients found.", "error")
        return redirect(url_for("editor"))

    return render_template(
        "index.html",
        page=3,
        subject=subject,
        importance=importance,
        html=html,
        email_column=email_col,
        cc=cc,
        sheet_name=sheet_name,
        recipient_count=len(recipients),
        recipients=recipients,
        preview_data_json=json.dumps(preview_data)  
    )

@app.route("/send-emails", methods=["POST"])
def send_emails():
    access_token = get_access_token()
    
    if not access_token:
        flash("Microsoft authentication expired. Please sign in again.", "error")
        return redirect(url_for("index"))
        
    subject_template = request.form.get("subject", "").strip()
    importance = request.form.get("importance", "normal")
    html_template = request.form.get("html", "")
    email_col = request.form.get("email_column", "").strip()
    cc_emails = request.form.getlist("cc")

    if not subject_template or not email_col:
        flash("Campaign data missing. Please try again.", "error")
        return redirect(url_for("editor"))

    wb = get_workbook()
    sheet_name = session.get("sheet_name")
    if not wb or not sheet_name or sheet_name not in wb.sheetnames:
        if wb:
            wb.close()
        flash("Worksheet no longer available.", "error")
        return redirect(url_for("index"))

    ws = wb[sheet_name]
    headers = {
        str(cell.value).strip(): idx 
        for idx, cell in enumerate(ws[1], start=1) if cell.value is not None
    }
    
    if email_col not in headers:
        wb.close()
        flash("Selected email column no longer exists.", "error")
        return redirect(url_for("editor"))

    email_idx = headers[email_col]
    graph_url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers_http = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    results = []
    sent, failed = 0, 0

    for row_number in range(2, ws.max_row + 1):
        email_val = ws.cell(row=row_number, column=email_idx).value
        if not email_val:
            continue
            
        email = str(email_val).strip()
        if not email:
            continue

        row_data = get_row_data(ws, row_number)
        
        raw_subj = process_dynamic_content(subject_template, row_data)
        subject = raw_subj.replace("\n", " ").replace("\r", "").strip()

        raw_html = process_dynamic_content(html_template, row_data)
        html_payload = f"""<!DOCTYPE html>
        <html>
        <head><meta charset="UTF-8"></head>
        <body>
            {raw_html}
        </body>
        </html>"""

        message = {
            "message": {
                "subject": subject,
                "importance": importance,
                "body": {
                    "contentType": "HTML",
                    "content": html_payload
                },
                "toRecipients": [{"emailAddress": {"address": email}}]
            },
            "saveToSentItems": True
        }

        if cc_emails:
            message["message"]["ccRecipients"] = [
                {"emailAddress": {"address": cc_email}} 
                for cc_email in cc_emails if cc_email.strip()
            ]

        try:
            resp = requests.post(graph_url, headers=headers_http, json=message, timeout=30)
            if resp.status_code == 202:
                sent += 1
                results.append({"email": email, "status": "Sent"})
            else:
                failed += 1
                results.append({"email": email, "status": "Failed"})
        except requests.RequestException:
            failed += 1
            results.append({"email": email, "status": "Failed"})

    wb.close()
    return render_template("index.html", page=4, sent=sent, failed=failed, results=results)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)