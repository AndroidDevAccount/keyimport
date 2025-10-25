# comics_scraper.py
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
import os
import datetime

# ---------- Config ----------
CHROME_DRIVER_PATH = "C:\\projects\\keyimport\\chromedriver.exe"
HEADLESS = True

EXCEL_FILE = "Action_Comics_Key_Issues.xlsx"

TARGET_URL = "https://www.keycollectorcomics.com/series/action-comics-2,5584/?publishedDate=bronzeAge&groupBy=issue&orderBy=publishedDate"

# ---------- Excel setup ----------
wb = Workbook()
ws = wb.active
ws.title = "Key Issues"
ws.append(["Series", "Issue #", "Price", "Key Facts", "Year"])

# ---------- Selenium setup ----------
opts = Options()
if HEADLESS:
    opts.add_argument("--headless=new")
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=opts)

# Navigate
driver.get(TARGET_URL)
time.sleep(2)  # maybe increase to wait for JS

# Find the container with all key issue rows
container = driver.find_element("css selector", "#expanded-issue-list")
rows = container.find_elements("css selector", "div.px-0")

for row in rows:
    try:
        h2 = row.find_element("tag name", "h2")
        h2_text = h2.text.strip()
        # Try to split into series and issue number (e.g., 'Action Comics #421')
        if '#' in h2_text:
            series, issue_num = h2_text.split('#', 1)
            series = series.strip()
            issue_num = issue_num.strip()
        else:
            series = h2_text
            issue_num = ""
    except NoSuchElementException:
        series = ""
        issue_num = ""
    # Get price columns
    prices = row.find_elements("css selector", "span.currency")
    low = prices[0].text.strip().replace("$","") if len(prices) > 0 else ""
    mid = prices[1].text.strip().replace("$","") if len(prices) > 1 else ""
    high = prices[2].text.strip().replace("$","") if len(prices) > 2 else ""
    info = f"Low - ${low} Mid - ${mid} High - ${high}".strip()
    # Get year from first h3 (if exists)
    h3s = row.find_elements("tag name", "h3")
    if len(h3s) > 0:
        year = h3s[0].text.strip()
    else:
        year = ""
    # Get key facts from the first h3 in the third div under the main div
    try:
        inner_divs = row.find_elements("xpath", "./div/div")
        if len(inner_divs) >= 3:
            key_facts_h3 = inner_divs[2].find_element("tag name", "h3")
            key_facts = key_facts_h3.text.strip()
        else:
            key_facts = ""
    except NoSuchElementException:
        key_facts = ""
    ws.append([series, issue_num, info, key_facts, year])
    print("Saved:", series, issue_num, info, key_facts, year)

driver.quit()
try:
    wb.save(EXCEL_FILE)
    print("Done! Check the Excel file →", EXCEL_FILE)
    os.startfile(EXCEL_FILE)
except PermissionError:
    # Save to a new file with a timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    alt_file = f"Action_Comics_Key_Issues_{timestamp}.xlsx"
    wb.save(alt_file)
    print(f"{EXCEL_FILE} is open. Saved as {alt_file} instead.")
    os.startfile(alt_file)
