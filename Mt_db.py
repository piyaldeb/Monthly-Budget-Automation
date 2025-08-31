#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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

# -------------------------
# Config (env-first)
# -------------------------
BASE_DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# Odoo
ODOO_URL      = os.environ.get("ODOO_URL", "https://taps.odoo.com").rstrip("/")
ODOO_DB       = os.environ.get("ODOO_DB", "taps")           # <- set your DB name
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "ranak@texzipperbd.com")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "2326")
COMPANY_ID    = int(os.environ.get("ODOO_COMPANY_ID", "3"))  # allowed_company_ids, options.company_id

# Report flow (model + button + report type)
MODEL                = os.environ.get("ODOO_WIZARD_MODEL", "taps_manufacturing.pi.wizard")
REPORT_TYPE          = os.environ.get("ODOO_REPORT_TYPE", "pi_xls")
REPORT_BUTTON_METHOD = os.environ.get("ODOO_REPORT_BUTTON", "print_xlsx")
REPORT_TEMPLATE_FALLBACK = os.environ.get("ODOO_REPORT_TEMPLATE", "taps_manufacturing.pi_xls_template")

# Date window (DD/MM/YYYY)
DATE_FROM = os.environ.get("DATE_FROM", "01/08/2025")
DATE_TO   = os.environ.get("DATE_TO",   "31/08/2025")

# Google Sheets
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME     = os.environ.get("SHEET_NAME", "Mt")
PASTE_COLUMNS  = int(os.environ.get("PASTE_COLUMNS", "9"))  # Metal only needs 9 cols

# -------------------------
# Utility
# -------------------------
def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

def _list_files(d, patterns=("*.xlsx","*.xls")):
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(d, p)))
    return sorted(files, key=os.path.getmtime)

# -------------------------
# Google Sheets helpers
# -------------------------
def get_google_sheets_service():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds).spreadsheets().values()

# -------------------------
# Download Odoo Report (HTTP API flow)
# -------------------------
def download_from_odoo() -> str | None:
    """
    Logs into Odoo, configures the report wizard and downloads an XLSX.
    Returns local filename or None.
    """
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    # -- Step 1: login
    login_url = f"{ODOO_URL}/web/session/authenticate"
    login_payload = {
        "jsonrpc": "2.0",
        "params": {"db": ODOO_DB, "login": ODOO_USERNAME, "password": ODOO_PASSWORD},
    }
    log("Authenticating to Odoo…")
    r = session.post(login_url, json=login_payload, timeout=30)
    r.raise_for_status()
    uid = r.json().get("result", {}).get("uid")
    if not uid:
        log(f"❌ Login failed: {r.text[:500]}")
        return None

    # -- Step 2: fetch CSRF from /web
    log("Fetching CSRF token…")
    r = session.get(f"{ODOO_URL}/web", timeout=30)
    r.raise_for_status()
    m = re.search(r'csrf_token:\s*"([A-Za-z0-9]+)"', r.text)
    csrf_token = m.group(1) if m else None
    if not csrf_token:
        log("⚠️ CSRF token not found, proceeding without explicit header (server may still accept).")

    # -- Step 3: create wizard
    create_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/create"
    log(f"Creating wizard for model={MODEL} …")
    r = session.post(create_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"model": MODEL, "method": "create", "args": [{}], "kwargs": {"context": {"uid": uid}}},
    }, timeout=30)
    r.raise_for_status()
    wizard_id = r.json().get("result")
    if not wizard_id:
        log(f"❌ Wizard create failed: {r.text[:500]}")
        return None

    # -- Step 3b: save wizard with dates
    save_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    log(f"Saving wizard {wizard_id} with dates {DATE_FROM} → {DATE_TO} …")
    r = session.post(save_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {"report_type": REPORT_TYPE, "date_from": DATE_FROM, "date_to": DATE_TO}],
            "kwargs": {
                "context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [COMPANY_ID]},
                "specification": {"report_type": {}, "date_from": {}, "date_to": {}},
            },
        },
    }, timeout=30)
    r.raise_for_status()
    res_list = r.json().get("result") or [{}]
    saved_id = (res_list[0] or {}).get("id") if isinstance(res_list, list) else None
    wizard_id = saved_id or wizard_id

    # -- Step 4: call the report button
    button_url = f"{ODOO_URL}/web/dataset/call_button"
    log(f"Calling report button method={REPORT_BUTTON_METHOD} …")
    r = session.post(button_url, json={
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": REPORT_BUTTON_METHOD,
            "args": [[wizard_id]],
            "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [COMPANY_ID]}},
        },
    }, timeout=30)
    r.raise_for_status()
    report_info = r.json().get("result") or {}
    report_template = report_info.get("report_name") or REPORT_TEMPLATE_FALLBACK

    # -- Step 5: /report/download
    log(f"Downloading report using template={report_template} …")
    options = {"date_from": DATE_FROM, "date_to": DATE_TO, "company_id": COMPANY_ID}
    context = {
        "lang": "en_US", "tz": "Asia/Dhaka", "uid": uid,
        "allowed_company_ids": [COMPANY_ID],
        "active_model": MODEL, "active_id": wizard_id, "active_ids": [wizard_id],
    }
    report_path = f"/report/xlsx/{report_template}?options={json.dumps(options)}&context={json.dumps(context)}"
    payload = {
        "data": json.dumps([report_path, "xlsx"]),
        "context": json.dumps(context),
        "token": "dummy",
    }
    headers = {"Referer": f"{ODOO_URL}/web"}
    if csrf_token:
        headers["X-CSRF-Token"] = csrf_token

    r = session.post(f"{ODOO_URL}/report/download", data=payload, headers=headers, timeout=60)
    if r.status_code == 200 and r.headers.get("content-type", "").startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        fname = os.path.join(BASE_DOWNLOAD_DIR, f"{REPORT_TYPE}_{DATE_FROM}_to_{DATE_TO}.xlsx")
        with open(fname, "wb") as f:
            f.write(r.content)
        log(f"Saved report: {fname}")
        return fname

    log(f"❌ Download failed: {r.status_code} {r.text[:500]}")
    return None

# -------------------------
# Paste into Google Sheet
# -------------------------
def update_google_sheet_with_file(file_path, sheet_name):
    df = pd.read_excel(file_path, engine="openpyxl")
    # numeric cleanup for columns 2..PASTE_COLUMNS
    for col in df.columns[1:PASTE_COLUMNS]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
    values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()

    service = get_google_sheets_service()
    # adjust range width (A:I for 9 cols)
    last_col_letter = chr(ord('A') + PASTE_COLUMNS - 1)
    service.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:{last_col_letter}").execute()
    service.update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()
    if os.path.exists(file_path):
        os.remove(file_path)

# -------------------------
# Main
# -------------------------
def main():
    log(f"Starting report run: {DATE_FROM} → {DATE_TO}")
    try:
        fpath = download_from_odoo()
        if not fpath:
            log("Download failed; exiting non-zero.")
            raise SystemExit(1)
        update_google_sheet_with_file(fpath, SHEET_NAME)
        log("✅ Sheet updated successfully.")
    except Exception:
        log("❌ Fatal error:\n" + traceback.format_exc())
        raise SystemExit(1)

if __name__ == "__main__":
    main()
