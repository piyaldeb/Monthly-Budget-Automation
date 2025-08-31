import os
import requests
import json
import re
import datetime
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ========= CONFIG ==========
ODOO_URL = "https://taps.odoo.com"
DB = "masbha-tex-taps-master-2093561"
USERNAME = "ranak@texzipperbd.com"
PASSWORD = "2326"

MODEL = "mrp.report.custom"   # Wizard model
REPORT_BUTTON_METHOD = "action_generate_xlsx_report"

REPORT_TYPE = "s_invs"
DATE_FROM = "2025-08-01"
DATE_TO = datetime.date.today().strftime("%Y-%m-%d")

BASE_DOWNLOAD_DIR = os.path.join(os.getcwd(), "odoo_reports")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# Google Sheets config
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS = int(os.environ.get("PASTE_COLUMNS", "25"))

# -------------------------
# GOOGLE SHEETS SERVICE (fixed)
# -------------------------
def get_google_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()  # return the spreadsheets resource

# -------------------------
# Update Google Sheet (fixed)
# -------------------------
def update_google_sheet_with_file(file_path, sheet_name):
    df = pd.read_excel(file_path, engine="openpyxl")
    
    # Round numeric columns
    for col in df.columns[1:PASTE_COLUMNS]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    df = df.where(pd.notnull(df), "")
    
    df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
    values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()

    service = get_google_sheets_service()

    # Clear existing content first
    service.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1:Z1000"
    ).execute()

    # Paste new values
    service.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

    # Delete local file
    if os.path.exists(file_path):
        os.remove(file_path)


# -------------------------
# Download Odoo Report
# -------------------------
def download_from_odoo():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    # Step 1: Login
    login_url = f"{ODOO_URL}/web/session/authenticate"
    login_payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    resp = session.post(login_url, json=login_payload)
    uid = resp.json().get("result", {}).get("uid")
    
    # Step 2: CSRF token
    resp = session.get(f"{ODOO_URL}/web")
    match = re.search(r'var odoo = {\s*csrf_token: "([A-Za-z0-9]+)"', resp.text)
    csrf_token = match.group(1) if match else None
    
    # Step 3: Create wizard
    create_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/create"
    resp = session.post(create_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"model": MODEL, "method": "create", "args": [{}], "kwargs": {"context": {"uid": uid}}}
    })
    wizard_id = resp.json().get("result")

    # Step 3b: Save wizard with dates
    save_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    resp = session.post(save_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {"report_type": REPORT_TYPE, "date_from": DATE_FROM, "date_to": DATE_TO}],
            "kwargs": {
                "context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [3]},
                "specification": {"report_type": {}, "date_from": {}, "date_to": {}}
            }
        }
    })
    wizard_id = resp.json().get("result", [{}])[0].get("id")

    # Step 4: Call report button
    button_url = f"{ODOO_URL}/web/dataset/call_button"
    resp = session.post(button_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": REPORT_BUTTON_METHOD,
            "args": [[wizard_id]],
            "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [1]}}
        }
    })
    report_info = resp.json().get("result")

    # Step 5: Prepare download payload
    options = {"date_from": DATE_FROM, "date_to": DATE_TO, "company_id": 3}
    context = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [3],
               "active_model": MODEL, "active_id": wizard_id, "active_ids": [wizard_id]}
    REPORT_TEMPLATE = report_info.get("report_name") or "taps_manufacturing.pi_xls_template"
    report_path = f"/report/xlsx/{REPORT_TEMPLATE}?options={json.dumps(options)}&context={json.dumps(context)}"
    download_payload = {"data": json.dumps([report_path, "xlsx"]), "context": json.dumps(context),
                        "token": "dummy-because-api-expects-one", "csrf_token": csrf_token}

    download_url = f"{ODOO_URL}/report/download"
    headers = {"X-CSRF-Token": csrf_token, "Referer": f"{ODOO_URL}/web"}
    resp = session.post(download_url, data=download_payload, headers=headers, timeout=60)

    if resp.status_code == 200 and resp.headers.get("content-type", "").startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        filename = os.path.join(BASE_DOWNLOAD_DIR, f"{REPORT_TYPE}_{DATE_FROM}_to_{DATE_TO}.xlsx")
        with open(filename, "wb") as f:
            f.write(resp.content)
        return filename
    else:
        print("❌ Download failed:", resp.text[:500])
        return None

# -------------------------
# Main
# -------------------------
def main():
    print("Starting Odoo report download and Google Sheet update...")
    downloaded_file = download_from_odoo()
    if downloaded_file:
        update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME)
        print("✅ Google Sheet updated successfully.")
    else:
        print("❌ Download failed; exiting with non-zero.")
        raise SystemExit(1)

if __name__ == "__main__":
    main()
