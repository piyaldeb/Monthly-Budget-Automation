import os
import re
import json
import time
import glob
import traceback
from datetime import datetime
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ===============================
# Config (env first, fallback)
# ===============================
BASE_DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

ODOO_URL   = os.environ.get("ODOO_URL",   "https://taps.odoo.com").rstrip("/")
DB         = os.environ.get("ODOO_DB",    "masbha-tex-taps-master-2093561")
USERNAME   = os.environ.get("ODOO_USERNAME", "ranak@texzipperbd.com")
PASSWORD   = os.environ.get("ODOO_PASSWORD", "2326")

# Wizard / report config
MODEL                 = os.environ.get("ODOO_WIZARD_MODEL", "mrp.report.custom")
REPORT_BUTTON_METHOD  = os.environ.get("ODOO_REPORT_BUTTON", "action_generate_xlsx_report")
REPORT_TYPE           = os.environ.get("ODOO_REPORT_TYPE",   "invs")      # 'pi' | 'pir' | 'r_invs'
COMPANY_ID            = int(os.environ.get("ODOO_COMPANY_ID", "3"))         # 3 = Metal (per your context)
TZ                    = os.environ.get("ODOO_TZ", "Asia/Dhaka")

# Dates (YYYY-MM-DD for Odoo)
DATE_FROM = os.environ.get("DATE_FROM", "2025-01-01")
DATE_TO   = os.environ.get("DATE_TO",   datetime.now().strftime("%Y-%m-%d"))

# Google Sheets
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID       = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME           = os.environ.get("SHEET_NAME", "Mt")
PASTE_COLUMNS        = int(os.environ.get("PASTE_COLUMNS", "9"))  # A:I for 9 columns

# ===============================
# Utilities
# ===============================
def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

def newest_file(directory, patterns=("*.xlsx","*.xls")):
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(directory, p)))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

# ===============================
# Google Sheets helpers
# ===============================
def get_google_sheets_service_values():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds).spreadsheets().values()

def update_google_sheet_with_file(file_path: str, sheet_name: str, paste_cols: int):
    df = pd.read_excel(file_path, engine="openpyxl")
    # Numeric rounding on columns 2..paste_cols (index 1..paste_cols-1)
    for col in df.columns[1:paste_cols]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    # Replace NaNs with blank for Sheets
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:paste_cols]

    values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()
    svc = get_google_sheets_service_values()
    # Clear A:I (or A:whatever based on paste_cols)
    last_col_letter = chr(ord('A') + paste_cols - 1)
    svc.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:{last_col_letter}").execute()
    svc.update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()
    log(f"✅ Google Sheet '{sheet_name}' updated with {len(values)-1} rows.")

# ===============================
# Odoo (requests) flow
# ===============================
def odoo_login(session: requests.Session) -> int:
    url = f"{ODOO_URL}/web/session/authenticate"
    payload = {
        "jsonrpc": "2.0",
        "params": {"db": DB, "login": USERNAME, "password": PASSWORD},
    }
    r = session.post(url, json=payload)
    r.raise_for_status()
    uid = r.json().get("result", {}).get("uid")
    if not uid:
        raise RuntimeError(f"Login failed: {r.text[:500]}")
    log(f"✅ Logged in, uid={uid}")
    return uid

def get_csrf_token(session: requests.Session) -> str:
    # Not always needed for GET download, but we fetch it for the POST fallback route
    r = session.get(f"{ODOO_URL}/web")
    r.raise_for_status()
    m = re.search(r'csrf_token:\s*"([A-Za-z0-9]+)"', r.text)
    token = m.group(1) if m else ""
    log(f"ℹ️ CSRF token: {token or '(none found)'}")
    return token

def wizard_create(session: requests.Session, uid: int) -> int:
    url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/create"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "create",
            "args": [{}],
            "kwargs": {"context": {"uid": uid}},
        },
    }
    r = session.post(url, json=payload)
    r.raise_for_status()
    wiz_id = r.json().get("result")
    if not wiz_id:
        raise RuntimeError(f"Wizard create failed: {r.text[:500]}")
    log(f"✅ Wizard created: id={wiz_id}")
    return wiz_id

def wizard_save(session: requests.Session, uid: int, report_type: str, date_from: str, date_to: str) -> int:
    url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {
                "report_type": report_type,
                "date_from": date_from,
                "date_to": date_to
            }],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": TZ,
                    "uid": uid,
                    "allowed_company_ids": [COMPANY_ID],
                },
                "specification": {"report_type": {}, "date_from": {}, "date_to": {}},
            },
        },
    }
    r = session.post(url, json=payload)
    r.raise_for_status()
    result = r.json().get("result") or []
    wiz_id = result[0].get("id") if result and isinstance(result, list) else None
    if not wiz_id:
        raise RuntimeError(f"web_save failed: {r.text[:500]}")
    log(f"✅ Wizard saved: id={wiz_id} ({report_type} {date_from}→{date_to})")
    return wiz_id

def press_report_button(session: requests.Session, uid: int, wiz_id: int) -> dict:
    url = f"{ODOO_URL}/web/dataset/call_button"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": REPORT_BUTTON_METHOD,
            "args": [[wiz_id]],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": TZ,
                    "uid": uid,
                    "allowed_company_ids": [COMPANY_ID],
                }
            },
        },
    }
    r = session.post(url, json=payload)
    r.raise_for_status()
    result = r.json().get("result") or {}
    # Expect an ir.actions.report dict
    if not result or result.get("type") not in ("ir.actions.report", "ir_actions_report"):
        log(f"⚠️ Unexpected button result (printing first 600 chars): {json.dumps(result, indent=2)[:600]}")
    else:
        log(f"✅ Button returned report action: {result.get('report_name')} ({result.get('report_type')})")
    return result

def build_report_paths(report_info: dict, date_from: str, date_to: str, uid: int, wiz_id: int):
    """
    Returns (relative_report_path, options, context) for xlsx download.
    We prefer direct GET to /report/xlsx/<report_name>?options=...&context=...
    """
    report_name = report_info.get("report_name")
    if not report_name:
        # If server didn't include it, you must set a sensible default/mapping per report_type
        # Fallback (update to the correct template on your server if needed):
        fallback = {
            # "pi":   "taps_manufacturing.pi_xls_template",
            # "pir":  "taps_manufacturing.packing_invoice_summery_separated",
            "invs":"taps_manufacturing.packing_invoice_summary",
        }
        report_name = fallback.get(REPORT_TYPE)
        if not report_name:
            raise RuntimeError("No report_name in action and no fallback mapping available. Please set one.")

    options = {"date_from": date_from, "date_to": date_to, "company_id": COMPANY_ID}
    context = {
        "lang": "en_US",
        "tz": TZ,
        "uid": uid,
        "allowed_company_ids": [COMPANY_ID],
        "active_model": MODEL,
        "active_id": wiz_id,
        "active_ids": [wiz_id],
    }
    rel = f"/report/xlsx/{report_name}?options={json.dumps(options)}&context={json.dumps(context)}"
    return rel, options, context

def try_direct_get_xlsx(session: requests.Session, relative_path: str, out_path: str) -> bool:
    url = f"{ODOO_URL}{relative_path}"
    r = session.get(url, stream=True)
    if r.status_code == 200 and r.headers.get("content-type","").startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return True
    log(f"Direct GET failed: {r.status_code} {r.headers.get('content-type','')}\n{r.text[:500]}")
    return False

def try_report_download_post(session: requests.Session, csrf_token: str, relative_path: str, out_path: str) -> bool:
    """
    Fallback to /report/download with form-encoded fields, like the web client does.
    """
    url = f"{ODOO_URL}/report/download"
    data = {
        "data": json.dumps([relative_path, "xlsx"]),
        "context": "{}",  # server doesn't strictly require here because it's in the URL already
        "token": "filetoken-123",  # any string
        "csrf_token": csrf_token or "",
    }
    headers = {"Referer": f"{ODOO_URL}/web"}
    r = session.post(url, data=data, headers=headers, timeout=120)
    if r.status_code == 200 and r.headers.get("content-type","").startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        with open(out_path, "wb") as f:
            f.write(r.content)
        return True
    log(f"/report/download failed: {r.status_code}\nHeaders: {r.headers}\nBody: {r.text[:800]}")
    return False

def download_from_odoo(date_from: str, date_to: str) -> str:
    """
    Full flow → returns the path of the downloaded XLSX file.
    """
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

    uid = odoo_login(session)
    csrf = get_csrf_token(session)

    wiz_id = wizard_create(session, uid)
    wiz_id = wizard_save(session, uid, REPORT_TYPE, date_from, date_to)
    report_info = press_report_button(session, uid, wiz_id)

    rel_path, _options, _context = build_report_paths(report_info, date_from, date_to, uid, wiz_id)
    out_name = f"{REPORT_TYPE}_{date_from}_to_{date_to}.xlsx".replace(":", "-")
    out_path = os.path.join(BASE_DOWNLOAD_DIR, out_name)

    log("⬇️  Attempting direct GET download...")
    if try_direct_get_xlsx(session, rel_path, out_path):
        log(f"✅ Downloaded via direct GET → {out_path}")
        return out_path

    log("↩️  Falling back to /report/download POST ...")
    if try_report_download_post(session, csrf, rel_path, out_path):
        log(f"✅ Downloaded via /report/download → {out_path}")
        return out_path

    raise RuntimeError("Download failed via both GET and POST.")

# ===============================
# Main
# ===============================
def main():
    try:
        log(f"Starting Odoo → Google Sheets run for {DATE_FROM} → {DATE_TO} [{REPORT_TYPE}] ...")
        xlsx_path = download_from_odoo(DATE_FROM, DATE_TO)
        update_google_sheet_with_file(xlsx_path, SHEET_NAME, PASTE_COLUMNS)
        # Clean up the local file
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)
        log("🎉 Done.Metal Sheet")
    except Exception as e:
        log("❌ Fatal error:")
        traceback.print_exc()
        raise SystemExit(1)

if __name__ == "__main__":
    main()
