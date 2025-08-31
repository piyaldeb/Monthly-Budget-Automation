import os
import requests
import json
import re

# ========= CONFIG ==========
ODOO_URL = "https://taps.odoo.com"
DB = "masbha-tex-taps-master-2093561"
USERNAME = "ranak@texzipperbd.com"
PASSWORD = "2326"

MODEL = "mrp.report.custom"   # Wizard model
REPORT_BUTTON_METHOD = "action_generate_xlsx_report"

# Report type options:
#   "pi"   → PI Report
#   "pir"  → Packing Invoice Report
#   "invs" → Inventory Statement Report
REPORT_TYPE = "s_invs"

DATE_FROM = "2025-08-01"
DATE_TO   = "2025-08-31"

BASE_DOWNLOAD_DIR = os.path.join(os.getcwd(), "odoo_reports")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# Optional Google Sheets config
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS = int(os.environ.get("PASTE_COLUMNS", "25"))

# ========= START SESSION ==========
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
})

# ----------------------
# Step 1: Login
login_url = f"{ODOO_URL}/web/session/authenticate"
login_payload = {
    "jsonrpc": "2.0",
    "params": {
        "db": DB,
        "login": USERNAME,
        "password": PASSWORD
    }
}
resp = session.post(login_url, json=login_payload)
login_result = resp.json()
uid = login_result.get("result", {}).get("uid")
print("✅ Step 1: Logged in, UID =", uid)

# ----------------------
# Step 2: Extract CSRF token
resp = session.get(f"{ODOO_URL}/web")
match = re.search(r'var odoo = {\s*csrf_token: "([A-Za-z0-9]+)"', resp.text)
csrf_token = match.group(1) if match else None
print("✅ Step 2: CSRF token =", csrf_token)

# ----------------------
# Step 3: Create wizard
create_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/create"
create_payload = {
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "model": MODEL,
        "method": "create",
        "args": [{}],
        "kwargs": {"context": {"uid": uid}}
    }
}
resp = session.post(create_url, json=create_payload)
wizard_id = resp.json().get("result")
print("✅ Step 3: Wizard created, ID =", wizard_id)

# ----------------------
# Step 3b: Save wizard with report_type + dates
save_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
save_payload = {
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "model": MODEL,
        "method": "web_save",
        "args": [[], {
            "report_type": REPORT_TYPE,
            "date_from": DATE_FROM,
            "date_to": DATE_TO
        }],
        "kwargs": {
            "context": {
                "lang": "en_US",
                "tz": "Asia/Dhaka",
                "uid": uid,
                "allowed_company_ids": [3]
            },
            "specification": {
                "report_type": {},
                "date_from": {},
                "date_to": {}
            }
        }
    }
}
resp = session.post(save_url, json=save_payload)
wizard_id = resp.json().get("result", [{}])[0].get("id")
print("✅ Step 3b: Wizard saved, ID =", wizard_id)

# ----------------------
# Step 4: Call report button
button_url = f"{ODOO_URL}/web/dataset/call_button"
button_payload = {
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "model": MODEL,
        "method": REPORT_BUTTON_METHOD,
        "args": [[wizard_id]],
        "kwargs": {
            "context": {
                "lang": "en_US",
                "tz": "Asia/Dhaka",
                "uid": uid,
                "allowed_company_ids": [1]
            }
        }
    }
}
resp = session.post(button_url, json=button_payload)
report_info = resp.json().get("result")
print("✅ Step 4: Report info =", json.dumps(report_info, indent=4))

# ----------------------
# Step 5: Prepare download payload
options = {"date_from": DATE_FROM, "date_to": DATE_TO, "company_id": 3}
context = {
    "lang": "en_US",
    "tz": "Asia/Dhaka",
    "uid": uid,
    "allowed_company_ids": [3],
    "active_model": MODEL,
    "active_id": wizard_id,
    "active_ids": [wizard_id]
}

REPORT_TEMPLATE = report_info.get("report_name") or "taps_manufacturing.pi_xls_template"
report_path = f"/report/xlsx/{REPORT_TEMPLATE}?options={json.dumps(options)}&context={json.dumps(context)}"
download_payload = {
    "data": json.dumps([report_path, "xlsx"]),
    "context": json.dumps(context),
    "token": "dummy-because-api-expects-one",
    "csrf_token": csrf_token
}
print("✅ Step 5: Download payload ready")

# ----------------------
# Step 6: Attempt to download
download_url = f"{ODOO_URL}/report/download"
headers = {
    "X-CSRF-Token": csrf_token,
    "Referer": f"{ODOO_URL}/web"
}

try:
    resp = session.post(download_url, data=download_payload, headers=headers, timeout=60)
    print("✅ Step 6: Download response status =", resp.status_code)

    content_type = resp.headers.get("content-type", "")
    if resp.status_code == 200 and content_type.startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        filename = os.path.join(BASE_DOWNLOAD_DIR, f"{REPORT_TYPE}_{DATE_FROM}_to_{DATE_TO}.xlsx")
        with open(filename, "wb") as f:
            f.write(resp.content)
        print(f"✅ Step 7: Report downloaded as {filename}")
    else:
        snippet = resp.text[:500]
        print("❌ Download failed, snippet:", snippet)
except Exception as e:
    print("❌ Exception during download:", e)
