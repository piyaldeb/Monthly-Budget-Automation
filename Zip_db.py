import os
import re
import json
import time
import glob
import traceback
from datetime import datetime, timedelta
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

# ===============================
# Config
# ===============================
BASE_DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

ODOO_URL   = os.getenv("ODOO_URL")
DB         = os.getenv("ODOO_DB")
USERNAME   = os.getenv("ODOO_USERNAME")
PASSWORD   = os.getenv("ODOO_PASSWORD")

MODEL                 = os.environ.get("ODOO_WIZARD_MODEL", "mrp.report.custom")
REPORT_BUTTON_METHOD  = os.environ.get("ODOO_REPORT_BUTTON", "action_generate_xlsx_report")
REPORT_TYPE           = os.environ.get("ODOO_REPORT_TYPE", "invs")
COMPANY_ID            = int(os.environ.get("ODOO_COMPANY_ID", "1"))
TZ                    = os.environ.get("ODOO_TZ", "Asia/Dhaka")

# Auto-calculate dates for yesterday's data
yesterday = datetime.now() - timedelta(days=1)
if yesterday.month != datetime.now().month:
    DATE_FROM = yesterday.replace(day=1).strftime("%Y-%m-%d")
    DATE_TO = yesterday.strftime("%Y-%m-%d")
else:
    DATE_FROM = datetime.now().replace(day=1).strftime("%Y-%m-%d")
    DATE_TO = yesterday.strftime("%Y-%m-%d")

DATE_FROM = os.environ.get("DATE_FROM", DATE_FROM)
DATE_TO = os.environ.get("DATE_TO", DATE_TO)

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID       = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME           = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS        = int(os.environ.get("PASTE_COLUMNS", "25"))
CHUNK_DAYS           = int(os.environ.get("CHUNK_DAYS", "5"))

# ===============================
# Utilities
# ===============================
def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

def get_date_chunks(start_date_str: str, end_date_str: str, chunk_days: int = 5):
    start = datetime.strptime(start_date_str, "%Y-%m-%d")
    end = datetime.strptime(end_date_str, "%Y-%m-%d")
    chunks = []
    current = start
    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)
        chunks.append((current.strftime("%Y-%m-%d"), chunk_end.strftime("%Y-%m-%d")))
        current = chunk_end + timedelta(days=1)
    return chunks

# ===============================
# Google Sheets
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
    for col in df.columns[1:paste_cols]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:paste_cols]

    svc = get_google_sheets_service_values()
    
    try:
        result = svc.get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:Z").execute()
        existing_values = result.get('values', [])
    except:
        existing_values = []
    
    if len(existing_values) <= 1:
        values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()
        svc.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        log(f"✅ Sheet '{sheet_name}' initialized with {len(values)-1} rows.")
    else:
        headers = existing_values[0]
        existing_df = pd.DataFrame(existing_values[1:], columns=headers)
        new_df = df_to_paste.copy()
        new_df.columns = df_to_paste.columns.tolist()
        key_col = df_to_paste.columns[0]
        
        if key_col in existing_df.columns and len(existing_df) > 0:
            new_keys = set(new_df[key_col].astype(str))
            existing_df = existing_df[~existing_df[key_col].astype(str).isin(new_keys)]
        
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        values = [headers] + combined_df.values.tolist()
        
        last_col_letter = chr(ord('A') + paste_cols - 1)
        svc.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:{last_col_letter}").execute()
        svc.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        log(f"✅ Sheet '{sheet_name}' updated: {len(new_df)} new/updated, {len(values)-1} total rows.")

# ===============================
# Odoo
# ===============================
def odoo_login(session: requests.Session) -> int:
    url = f"{ODOO_URL}/web/session/authenticate"
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    uid = r.json().get("result", {}).get("uid")
    if not uid:
        raise RuntimeError(f"Login failed: {r.text[:500]}")
    log(f"✅ Logged in, uid={uid}")
    return uid

def get_csrf_token(session: requests.Session) -> str:
    r = session.get(f"{ODOO_URL}/web", timeout=60)
    r.raise_for_status()
    m = re.search(r'csrf_token:\s*"([A-Za-z0-9]+)"', r.text)
    return m.group(1) if m else ""

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
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    wiz_id = r.json().get("result")
    if not wiz_id:
        raise RuntimeError(f"Wizard create failed: {r.text[:500]}")
    return wiz_id

def wizard_save(session: requests.Session, uid: int, report_type: str, date_from: str, date_to: str) -> int:
    url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {"report_type": report_type, "date_from": date_from, "date_to": date_to}],
            "kwargs": {
                "context": {"lang": "en_US", "tz": TZ, "uid": uid, "allowed_company_ids": [COMPANY_ID]},
                "specification": {"report_type": {}, "date_from": {}, "date_to": {}},
            },
        },
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    result = r.json().get("result") or []
    wiz_id = result[0].get("id") if result and isinstance(result, list) else None
    if not wiz_id:
        raise RuntimeError(f"web_save failed: {r.text[:500]}")
    log(f"✅ Wizard saved: id={wiz_id} ({date_from}→{date_to})")
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
            "kwargs": {"context": {"lang": "en_US", "tz": TZ, "uid": uid, "allowed_company_ids": [COMPANY_ID]}},
        },
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    return r.json().get("result") or {}

def build_report_paths(report_info: dict, date_from: str, date_to: str, uid: int, wiz_id: int):
    report_name = report_info.get("report_name")
    if not report_name:
        fallback = {"invs": "taps_manufacturing.packing_invoice_summary"}
        report_name = fallback.get(REPORT_TYPE)
        if not report_name:
            raise RuntimeError("No report_name available")
    
    options = {"date_from": date_from, "date_to": date_to, "company_id": COMPANY_ID}
    context = {
        "lang": "en_US", "tz": TZ, "uid": uid, "allowed_company_ids": [COMPANY_ID],
        "active_model": MODEL, "active_id": wiz_id, "active_ids": [wiz_id],
    }
    rel = f"/report/xlsx/{report_name}?options={json.dumps(options)}&context={json.dumps(context)}"
    return rel, options, context

def try_direct_get_xlsx(session: requests.Session, relative_path: str, out_path: str, timeout: int = 180) -> bool:
    url = f"{ODOO_URL}{relative_path}"
    try:
        r = session.get(url, stream=True, timeout=timeout)
        if r.status_code == 200 and r.headers.get("content-type", "").startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
            with open(out_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return True
        return False
    except:
        return False

def download_from_odoo(date_from: str, date_to: str) -> str:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    
    uid = odoo_login(session)
    wiz_id = wizard_create(session, uid)
    wiz_id = wizard_save(session, uid, REPORT_TYPE, date_from, date_to)
    report_info = press_report_button(session, uid, wiz_id)
    
    time.sleep(5)
    
    rel_path, _, _ = build_report_paths(report_info, date_from, date_to, uid, wiz_id)
    out_name = f"{REPORT_TYPE}_{date_from}_to_{date_to}.xlsx".replace(":", "-")
    out_path = os.path.join(BASE_DOWNLOAD_DIR, out_name)
    
    if try_direct_get_xlsx(session, rel_path, out_path, timeout=180):
        log(f"✅ Downloaded → {out_path}")
        return out_path
    
    raise RuntimeError("Download failed")

# ===============================
# Main
# ===============================
def main():
    date_chunks = get_date_chunks(DATE_FROM, DATE_TO, CHUNK_DAYS)
    log(f"📅 Processing {len(date_chunks)} chunks: {DATE_FROM} → {DATE_TO}")
    
    all_dataframes = []
    
    for chunk_idx, (chunk_from, chunk_to) in enumerate(date_chunks, 1):
        log(f"\n{'='*50}")
        log(f"📦 Chunk {chunk_idx}/{len(date_chunks)}: {chunk_from} → {chunk_to}")
        
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                xlsx_path = download_from_odoo(chunk_from, chunk_to)
                df = pd.read_excel(xlsx_path, engine="openpyxl")
                all_dataframes.append(df)
                
                if os.path.exists(xlsx_path):
                    os.remove(xlsx_path)
                
                log(f"✅ Chunk {chunk_idx} done ({len(df)} rows)")
                break
            except Exception as e:
                log(f"❌ Chunk {chunk_idx} attempt {attempt} failed: {e}")
                if attempt < max_retries:
                    wait = 30 * attempt
                    log(f"⏳ Retrying in {wait}s...")
                    time.sleep(wait)
                else:
                    log(f"🚨 Chunk {chunk_idx} failed after {max_retries} attempts")
                    raise
    
    if not all_dataframes:
        raise RuntimeError("No data collected")
    
    log(f"\n{'='*50}")
    log(f"📊 Combining {len(all_dataframes)} chunks...")
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    
    first_col = combined_df.columns[0]
    before = len(combined_df)
    combined_df = combined_df.drop_duplicates(subset=[first_col], keep='last')
    after = len(combined_df)
    if before > after:
        log(f"🔄 Removed {before - after} duplicates")
    
    temp_file = os.path.join(BASE_DOWNLOAD_DIR, f"combined_{DATE_FROM}_to_{DATE_TO}.xlsx")
    combined_df.to_excel(temp_file, index=False, engine="openpyxl")
    log(f"💾 Combined: {temp_file} ({len(combined_df)} rows)")
    
    log(f"☁️ Uploading to Google Sheets...")
    update_google_sheet_with_file(temp_file, SHEET_NAME, PASTE_COLUMNS)
    
    if os.path.exists(temp_file):
        os.remove(temp_file)
    
    log("🎉 Done. Zipper")

if __name__ == "__main__":
    main()