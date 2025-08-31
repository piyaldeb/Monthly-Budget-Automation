import requests, json, re, html
from urllib.parse import urljoin
from datetime import datetime

# ========= CONFIG ==========
ODOO_URL  = "https://taps.odoo.com"
DB        = "masbha-tex-taps-master-2093561"
USERNAME  = "ranak@texzipperbd.com"
PASSWORD  = "2326"

MODEL                = "mrp.report.custom"            # Wizard model
REPORT_BUTTON_METHOD = "action_generate_xlsx_report"  # Button method
REPORT_TYPE          = "s_invs"                       # Your report type flag (saved into wizard)

DATE_FROM = "2025-08-01"
DATE_TO   = datetime.now().strftime("%Y-%m-%d")
COMPANY_ID = 3  # keep consistent everywhere

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

def dbg(label, resp, max_len=600):
    ct = resp.headers.get("content-type","")
    print(f"\n=== {label} ===\n{resp.status_code} {ct}\n{resp.text[:max_len]}")

def json_or_error(resp):
    try:
        data = resp.json()
    except Exception:
        dbg("Non-JSON response", resp)
        raise SystemExit("Server did not return JSON where expected.")
    if "error" in data:
        # Pretty print Odoo error
        err = data["error"]
        msg = err.get("message")
        name = err.get("data",{}).get("name")
        debug = err.get("data",{}).get("debug","")[:1000]
        raise SystemExit(f"Odoo JSON-RPC error: {name or ''} | {msg or ''}\n{debug}")
    return data

def find_csrf_token(html_text: str):
    # several shapes across Odoo versions
    m = re.search(r'csrf_token:\s*"([^"]+)"', html_text)
    if m: return m.group(1)
    m = re.search(r'<meta[^>]+name=["\']csrf-token["\'][^>]+content=["\']([^"\']+)["\']', html_text, re.I)
    if m: return html.unescape(m.group(1))
    m = re.search(r'<input[^>]+name=["\']csrf_token["\'][^>]+value=["\']([^"\']+)["\']', html_text, re.I)
    if m: return html.unescape(m.group(1))
    return None

def login():
    # Try JSON-RPC login a
    r = session.post(
        f"{ODOO_URL}/web/session/authenticate",
        json={"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}},
        timeout=30
    )
    if r.ok and r.headers.get("content-type","").startswith("application/json"):
        data = json_or_error(r)
        uid = (data.get("result") or {}).get("uid")
        if uid:
            rw = session.get(f"{ODOO_URL}/web", timeout=30)
            csrf = find_csrf_token(rw.text) if rw.ok else None
            return uid, csrf

    # Fallback: form login
    rg = session.get(f"{ODOO_URL}/web/login", timeout=30)
    rg.raise_for_status()
    csrf = find_csrf_token(rg.text)
    form = {"login": USERNAME, "password": PASSWORD, "db": DB}
    if csrf: form["csrf_token"] = csrf
    rp = session.post(f"{ODOO_URL}/web/login", data=form, headers={"Referer": f"{ODOO_URL}/web/login"}, timeout=30, allow_redirects=True)
    rp.raise_for_status()
    ri = session.get(f"{ODOO_URL}/web/session/get_session_info", timeout=30)
    uid = None
    if ri.ok and ri.headers.get("content-type","").startswith("application/json"):
        uid = (ri.json().get("result") or {}).get("uid")
    rw = session.get(f"{ODOO_URL}/web", timeout=30)
    csrf = find_csrf_token(rw.text) or csrf
    return uid, csrf

def rpc_call_kw(model, method, args=None, kwargs=None):
    payload = {"jsonrpc":"2.0","method":"call","params":{"model":model,"method":method,"args":args or [],"kwargs":kwargs or {}}}
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}", json=payload, timeout=60)
    data = json_or_error(r)
    return data.get("result")

def call_button(model, method, args=None, kwargs=None):
    payload = {"jsonrpc":"2.0","method":"call","params":{"model":model,"method":method,"args":args or [],"kwargs":kwargs or {}}}
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload, timeout=60)
    data = json_or_error(r)
    return data.get("result")

def main():
    print(f"▶ Run: {DATE_FROM} → {DATE_TO}")
    uid, csrf = login()
    if not uid:
        raise SystemExit("❌ Login failed — check DB/username/password.")
    print(f"✅ Logged in, uid={uid}, csrf={'yes' if csrf else 'no'}")

    # Step 3: create wizard
    wiz_id = rpc_call_kw(MODEL, "create", args=[[{}]], kwargs={"context":{"uid":uid}})
    if not wiz_id:
        raise SystemExit("❌ Wizard create returned no id.")
    print(f"✅ Wizard created: {wiz_id}")

    # Step 3b: save wizard fields (report_type + date range)
    save_res = rpc_call_kw(
        MODEL, "web_save",
        args=[[], {"report_type": REPORT_TYPE, "date_from": DATE_FROM, "date_to": DATE_TO}],
        kwargs={
            "context": {"lang":"en_US","tz":"Asia/Dhaka","uid":uid,"allowed_company_ids":[COMPANY_ID]},
            "specification": {"report_type": {}, "date_from": {}, "date_to": {}}
        }
    )
    # some servers return a list of records
    if isinstance(save_res, list) and save_res:
        wiz_id = (save_res[0] or {}).get("id", wiz_id)
    print(f"✅ Wizard saved as id={wiz_id}")

    # Step 4: press the button to prepare report
    button_res = call_button(
        MODEL, REPORT_BUTTON_METHOD,
        args=[[wiz_id]],
        kwargs={"context":{"lang":"en_US","tz":"Asia/Dhaka","uid":uid,"allowed_company_ids":[COMPANY_ID]}}
    )
    print("ℹ️ Button result:", json.dumps(button_res, indent=2))

    # Resolve download path
    report_template = None
    report_url_path = None
    if isinstance(button_res, dict):
        # new Odoo often returns a dict with "url" or "report_name"
        report_template = button_res.get("report_name")
        report_url_path = button_res.get("url")

    headers = {"Referer": f"{ODOO_URL}/web"}
    if csrf: headers["X-CSRF-Token"] = csrf

    if report_url_path:
        # Example: "/report/xlsx/module.template?options=...&context=..."
        dl_url = urljoin(ODOO_URL, report_url_path)
        print("➡ Download via returned URL:", dl_url)
        r = session.get(dl_url, headers=headers, timeout=120, allow_redirects=True)
    else:
        # Step 5: build /report/download payload
        if not report_template:
            report_template = "taps_manufacturing.pi_xls_template"  # fallback
        options = {"date_from": DATE_FROM, "date_to": DATE_TO, "company_id": COMPANY_ID}
        context = {
            "lang":"en_US","tz":"Asia/Dhaka","uid":uid,"allowed_company_ids":[COMPANY_ID],
            "active_model":MODEL,"active_id":wiz_id,"active_ids":[wiz_id]
        }
        report_path = f"/report/xlsx/{report_template}?options={json.dumps(options)}&context={json.dumps(context)}"
        payload = {"data": json.dumps([report_path, "xlsx"]), "context": json.dumps(context), "token": "dummy"}
        print("➡ POST /report/download")
        r = session.post(f"{ODOO_URL}/report/download", data=payload, headers=headers, timeout=120)

    ct = r.headers.get("content-type","")
    if r.status_code == 200 and ct.startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        fname = f"{REPORT_TYPE}_{DATE_FROM}_to_{DATE_TO}.xlsx"
        with open(fname, "wb") as f:
            f.write(r.content)
        print(f"✅ Downloaded: {fname}")
    else:
        dbg("Download failed", r)
        raise SystemExit("❌ Report download failed.")

if __name__ == "__main__":
    main()
