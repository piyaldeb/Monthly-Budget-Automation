import os
import time
import glob
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import chromedriver_autoinstaller
from google.oauth2 import service_account
from googleapiclient.discovery import build

# -------------------------
# Configuration (env can override, but defaults keep your logic)
# -------------------------
DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")
ODOO_URL = os.environ.get("ODOO_URL", "https://taps.odoo.com")
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "ranak@texzipperbd.com")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "2326")

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS = int(os.environ.get("PASTE_COLUMNS", "25"))

# Keep your exact selectors/flow (logic unchanged)
DATE_INPUT_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[2]/div/div/input"
EXPORT_BTN_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]"

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

# -------------------------
# Google Sheets Auth
# -------------------------
def get_google_sheets_service():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds).spreadsheets().values()

# -------------------------
# Wait for XLSX download (with timeout for CI)
# -------------------------
def wait_for_download_complete(download_dir, timeout=180):
    log(f"Waiting for XLSX download to complete (timeout {timeout}s)...")
    start = time.time()
    initial_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
    while time.time() - start < timeout:
        time.sleep(1)
        current_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
        new_files = current_files - initial_files
        if new_files:
            latest_file = max(new_files, key=os.path.getctime)
            time.sleep(3)  # ensure fully written
            log(f"Download completed: {latest_file}")
            return latest_file
    log("Download timed out.")
    return None

def dump_debug(driver, tag="zip_odoo_debug"):
    try:
        driver.save_screenshot(f"{tag}.png")
        with open(f"{tag}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log(f"Saved debug artifacts: {tag}.png, {tag}.html")
    except Exception:
        pass

# -------------------------
# Selenium: Download Report (logic preserved)
# -------------------------
def download_from_odoo(company="Zipper", date_from="01/01/2025", date_to=None):
    if not date_to:
        date_to = (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")

    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--log-level=3")

    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(90)
    driver.set_script_timeout(60)
    wait = WebDriverWait(driver, 40)

    # Explicit headless download permission (helps CI)
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": os.path.abspath(DOWNLOAD_DIR)}
        )
    except Exception:
        pass

    try:
        log("Opening Odoo login page...")
        driver.get(ODOO_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(ODOO_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(ODOO_PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(2)

        # Switch company (same selector)
        log(f"Switching company to '{company}'...")
        switcher = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher)
        switcher.click()
        time.sleep(1)

        target_div = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")
        ))
        target_div.click()
        time.sleep(2)

        # Navigate to MRP Reports (unchanged)
        log("Navigating to MRP REPORTS...")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(2)

        # Select "Invoice Summary" (unchanged)
        log("Selecting 'Invoice Summary'...")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//select")))
        dropdown.click()
        dropdown.send_keys("Invoice Summary")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(1)

        # Set date (unchanged hardcoded input + XPath)
        log("Setting date (as in original logic)...")
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, DATE_INPUT_XPATH)))
        date_input.clear()
        date_input.send_keys("01/08/25")  # keep your exact value to preserve logic
        date_input.send_keys(Keys.ENTER)
        time.sleep(1)

        # Export (unchanged XPath)
        log("Clicking export...")
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, EXPORT_BTN_XPATH)))
        driver.execute_script("arguments[0].click();", export_btn)

        return wait_for_download_complete(DOWNLOAD_DIR, timeout=180)

    except Exception:
        log(f"Error during Odoo interaction:\n{traceback.format_exc()}")
        dump_debug(driver, tag="zip_odoo_error")
        return None
    finally:
        driver.quit()
        log("Browser closed")

# -------------------------
# Update Google Sheet (logic preserved)
# -------------------------
def update_google_sheet_with_file(file_path, sheet_name):
    df = pd.read_excel(file_path, engine="openpyxl")
    for col in df.columns[1:PASTE_COLUMNS]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
    values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()

    service = get_google_sheets_service()
    # Keep your fixed A:Y clear/update to preserve logic
    service.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:Y").execute()
    service.update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()
    if os.path.exists(file_path):
        os.remove(file_path)

# -------------------------
# Main
# -------------------------
def main():
    date_from = "01/01/2025"
    date_to = (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")
    log("Starting Zip run...")
    downloaded_file = download_from_odoo(company="Zipper", date_from=date_from, date_to=date_to)
    if not downloaded_file:
        log("Download failed; exiting with non-zero for CI.")
        raise SystemExit(1)
    update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME)
    log("Zip sheet updated successfully.")

if __name__ == "__main__":
    main()
