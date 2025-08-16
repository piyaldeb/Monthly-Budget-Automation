import os
import time
import glob
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
# Configuration
# -------------------------
DOWNLOAD_DIR = r"C:\Users\Ranak\Documents\Odoo"
ODOO_URL = "https://taps.odoo.com"
ODOO_USERNAME = "ranak@texzipperbd.com"
ODOO_PASSWORD = "2326"

SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
SPREADSHEET_ID = '1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I'
SHEET_NAME = "Mt"  # Sheet name
PASTE_COLUMNS = 9  # Columns A-I

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# -------------------------
# Google Sheets Auth
# -------------------------
def get_google_sheets_service():
    print(f"{datetime.now()} ⏳ Authenticating Google Sheets...")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds).spreadsheets().values()
    print(f"{datetime.now()} ✅ Google Sheets authentication done.")
    return service

# -------------------------
# Wait for XLSX download
# -------------------------
def wait_for_download_complete(download_dir):
    print(f"{datetime.now()} ⏳ Waiting for XLSX download to complete...")
    initial_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
    while True:
        time.sleep(1)
        current_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
        new_files = current_files - initial_files
        if new_files:
            latest_file = max(new_files, key=os.path.getctime)
            time.sleep(5)  # ensure fully written
            print(f"{datetime.now()} ✅ Download completed: {latest_file}")
            return latest_file

# -------------------------
# Selenium: Download Report (Headless)
# -------------------------
def download_from_odoo(company="Zipper", date_from="01/01/2025", date_to=None):
    if not date_to:
        date_to = (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")

    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--log-level=3")

    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    try:
        print(f"{datetime.now()} ⏳ Opening Odoo login page...")
        driver.get(ODOO_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(ODOO_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(ODOO_PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        print(f"{datetime.now()} ✅ Logged into Odoo")
        time.sleep(4)

        print(f"{datetime.now()} ⏳ Switching company to '{company}'...")
        switcher = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher)
        switcher.click()
        time.sleep(2)

        target_div = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")
        ))
        target_div.click()
        print(f"{datetime.now()} ✅ Company switched")
        time.sleep(4)

        print(f"{datetime.now()} ⏳ Navigating to MRP Reports...")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(5)

        print(f"{datetime.now()} ⏳ Selecting 'Invoice Summary' report...")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//select")))
        dropdown.click()
        dropdown.send_keys("Invoice Summary")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(2)

        print(f"{datetime.now()} ⏳ Setting date range: {date_from} to {date_to}...")
        date_inputs = driver.find_elements(By.XPATH, "//input[contains(@class,'o_datepicker')]")
        if date_inputs:
            date_inputs[0].clear()
            date_inputs[0].send_keys(date_from)
            date_inputs[1].clear()
            date_inputs[1].send_keys(date_to)
            date_inputs[1].send_keys(Keys.ENTER)
        time.sleep(2)

        print(f"{datetime.now()} ⏳ Clicking export button...")
        export_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]")
        ))
        driver.execute_script("arguments[0].click();", export_btn)

        return wait_for_download_complete(DOWNLOAD_DIR)

    except Exception as e:
        print(f"{datetime.now()} ❌ Error during Odoo interaction:", e)
    finally:
        driver.quit()
        print(f"{datetime.now()} ✅ Browser closed")

# -------------------------
# Update Google Sheet
# -------------------------
def update_google_sheet_with_file(file_path, sheet_name):
    try:
        print(f"{datetime.now()} ⏳ Reading Excel file: {file_path}")
        df = pd.read_excel(file_path, engine="openpyxl")

        print(f"{datetime.now()} ⏳ Processing data...")
        for col in df.columns[1:PASTE_COLUMNS]:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
        df = df.where(pd.notnull(df), "")
        df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
        values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()

        service = get_google_sheets_service()
        print(f"{datetime.now()} ⏳ Clearing existing data in sheet...")
        service.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:I").execute()
        print(f"{datetime.now()} ⏳ Updating Google Sheet with new data...")
        service.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        print(f"{datetime.now()} ✅ Google Sheet '{sheet_name}' updated successfully (A-I).")

        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"{datetime.now()} 🗑️ Deleted local file: {file_path}")

    except Exception as e:
        print(f"{datetime.now()} ❌ Error updating Google Sheets: {e}")

# -------------------------
# Main
# -------------------------
def main():
    date_from = "01/01/2025"
    date_to = (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")

    print(f"{datetime.now()} ⏳ Starting Odoo download and Google Sheets update process...")
    downloaded_file = download_from_odoo(company="Metal", date_from=date_from, date_to=date_to)
    if downloaded_file:
        update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME)
    else:
        print(f"{datetime.now()} ❌ Download failed, nothing to update.")

if __name__ == "__main__":
    main()
