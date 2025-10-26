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

SERIES_URLS = [
    ("Action Comics", "https://www.keycollectorcomics.com/series/action-comics-2,5584/?publishedDate=bronzeAge%2CcopperAge%2CmodernAge&groupBy=issue&orderBy=publishedDate"),
    ("Adventures of Superman", "https://www.keycollectorcomics.com/series/adventures-of-superman,5595/"),
    ("Alpha Flight", "https://www.keycollectorcomics.com/series/alpha-flight,5632/"),
    ("Avengers", "https://www.keycollectorcomics.com/series/avengers-2,5696/?publishedDate=bronzeAge%2CcopperAge&groupBy=issue&orderBy=publishedDate"),
    ("Avengers", "https://www.keycollectorcomics.com/series/avengers-2,5696/?page=2&publishedDate=bronzeAge%2CcopperAge&groupBy=issue&orderBy=publishedDate"),
    ("Avengers", "https://www.keycollectorcomics.com/series/avengers-2,5696/?page=3&publishedDate=bronzeAge%2CcopperAge&groupBy=issue&orderBy=publishedDate"),
    ("Batman", "https://www.keycollectorcomics.com/series/batman,5708/?publishedDate=bronzeAge%2CcopperAge%2CmodernAge&groupBy=issue&orderBy=publishedDate"),
    ("Batman", "https://www.keycollectorcomics.com/series/batman,5708/?page=2&publishedDate=bronzeAge%2CcopperAge%2CmodernAge&groupBy=issue&orderBy=publishedDate"),
    ("Batman", "https://www.keycollectorcomics.com/series/batman,5708/?page=3&publishedDate=bronzeAge%2CcopperAge%2CmodernAge&groupBy=issue&orderBy=publishedDate"),
    ("Captain America", "https://www.keycollectorcomics.com/series/captain-america,5808/"),
    ("Captain America", "https://www.keycollectorcomics.com/series/captain-america,5808/?page=2"),
    ("Captain America", "https://www.keycollectorcomics.com/series/captain-america,5808/?page=3"),
    ("Daredevil (1964)", "https://www.keycollectorcomics.com/series/daredevil-4,5913/?publishedDate=modernAge%2CbronzeAge&groupBy=issue&orderBy=publishedDate"),
    ("Daredevil (1964)", "https://www.keycollectorcomics.com/series/daredevil-4,5913/?page=2&publishedDate=modernAge%2CbronzeAge&groupBy=issue&orderBy=publishedDate"),
    ("Daredevil (Marvel Knights, 1998)", "https://www.keycollectorcomics.com/series/daredevil-2100,90688/"),
    ("Detective Comics", "https://www.keycollectorcomics.com/series/detective-comics,5968/?publishedDate=bronzeAge%2CcopperAge%2CmodernAge%2CsilverAge&groupBy=issue&orderBy=publishedDate"),
    ("Detective Comics", "https://www.keycollectorcomics.com/series/detective-comics,5968/?page=2&publishedDate=bronzeAge%2CcopperAge%2CmodernAge%2CsilverAge&groupBy=issue&orderBy=publishedDate"),
    ("Detective Comics", "https://www.keycollectorcomics.com/series/detective-comics,5968/?page=3&publishedDate=bronzeAge%2CcopperAge%2CmodernAge%2CsilverAge&groupBy=issue&orderBy=publishedDate"),
    ("Excalibur", "https://www.keycollectorcomics.com/series/excalibur,7476/"),
    ("Fantastic Four", "https://www.keycollectorcomics.com/series/fantastic-four-2,6027/"),
    ("Fantastic Four", "https://www.keycollectorcomics.com/series/fantastic-four-2,6027/?page=2"),
    ("Fantastic Four", "https://www.keycollectorcomics.com/series/fantastic-four-2,6027/?page=3"),
    ("Fantastic Four", "https://www.keycollectorcomics.com/series/fantastic-four-2,6027/?page=4"),
    ("The Flash (1958 Series)", "https://www.keycollectorcomics.com/series/flash-the,6059/"),
    ("The Flash (1958 Series)", "https://www.keycollectorcomics.com/series/flash-the,6059/?page=2"),
    ("The Flash (1987 Series)", "https://www.keycollectorcomics.com/series/flash-3,6054/"),
    ("Green Arrow (1988 Series)", "https://www.keycollectorcomics.com/series/green-arrow-2,6135/"),
    ("Green Arrow (2001 Series)", "https://www.keycollectorcomics.com/series/green-arrow-4,6136/"),
    ("Ghost Rider (1973 Series)", "https://www.keycollectorcomics.com/series/ghost-rider-2100,75238/?groupBy=issue&orderBy=publishedDate"),
    ("Ghost Rider (1990 Series)", "https://www.keycollectorcomics.com/series/ghost-rider-2101,75239/"),
    ("Green Lantern (Vol 2)", "https://www.keycollectorcomics.com/series/green-lantern-4,6148/?publishedDate=bronzeAge%2CcopperAge&groupBy=issue&orderBy=publishedDate"),
    ("Green Lantern (Vol 3)", "https://www.keycollectorcomics.com/series/green-lantern-2,6145/"),
    ("Tales of the Green Lantern Corps", "https://www.keycollectorcomics.com/series/tales-of-the-green-lantern-corps,6815/"),
    ("Green Lantern Corps", "https://www.keycollectorcomics.com/series/green-lantern-corps-2,72631/"),
    ("Incredible Hulk", "https://www.keycollectorcomics.com/series/incredible-hulk,6209/"),
    ("Incredible Hulk", "https://www.keycollectorcomics.com/series/incredible-hulk,6209/?page=2"),
    ("Incredible Hulk", "https://www.keycollectorcomics.com/series/incredible-hulk,6209/?page=3"),
    ("Iron Man", "https://www.keycollectorcomics.com/series/iron-man-2,6231/"),
    ("Iron Man", "https://www.keycollectorcomics.com/series/iron-man-2,6231/?page=2"),
    ("Justice League of America", "https://www.keycollectorcomics.com/series/justice-league-of-america,6273/"),
    ("Justice League America", "https://www.keycollectorcomics.com/series/justice-league-america-2,71553/"),
    ("Justice League International", "https://www.keycollectorcomics.com/series/justice-league-international,6272/"),
    ("Justice League Europe", "https://www.keycollectorcomics.com/series/justice-league-europe,6271/"),
    ("Marvel Tales", "https://www.keycollectorcomics.com/series/marvel-tales,6382/"),
    ("Marvel Team-Up", "https://www.keycollectorcomics.com/series/marvel-team-up,6384/"),
    ("Marvel Two-in-One", "https://www.keycollectorcomics.com/series/marvel-two-in-one,6388/"),
    ("New Mutants", "https://www.keycollectorcomics.com/series/new-mutants-2,6483/"),
    ("Superman (1939", "https://www.keycollectorcomics.com/series/superman,6783/?publishedDate=bronzeAge%2CcopperAge%2CmodernAge%2CsilverAge&groupBy=issue&orderBy=publishedDate"),
    ("Superman (1987)", "https://www.keycollectorcomics.com/series/superman-2,6784/"),
    ("Superman's Pal Jimmy Olsen", "https://www.keycollectorcomics.com/series/supermans-pal-jimmy-olsen,6797/"),
    ("The Mighty Thor", "https://www.keycollectorcomics.com/series/thor,6844/"),
    ("The Mighty Thor", "https://www.keycollectorcomics.com/series/thor,6844/?page=2"),
    ("The Mighty Thor", "https://www.keycollectorcomics.com/series/thor,6844/?page=3"),
    ("Wolverine", "https://www.keycollectorcomics.com/series/wolverine-8,77556/"),
    ("World's Finest Comics", "https://www.keycollectorcomics.com/series/worlds-finest-comics,6990/?publishedDate=silverAge%2CbronzeAge&groupBy=issue&orderBy=publishedDate"),
    ("X-Factor", "https://www.keycollectorcomics.com/series/x-factor,6996/"),
    ("X-Force", "https://www.keycollectorcomics.com/series/x-force,6998/"),
    ("X-Men", "https://www.keycollectorcomics.com/series/x-men-2,7000/"),
]

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

for series_name, TARGET_URL in SERIES_URLS:
    driver.get(TARGET_URL)
    time.sleep(2)  # maybe increase to wait for JS
    container = driver.find_element("css selector", "#expanded-issue-list")
    rows = container.find_elements("css selector", "div.px-0")
    for row in rows:
        try:
            h2 = row.find_element("tag name", "h2")
            h2_text = h2.text.strip()
            # Try to split into series and issue number (e.g., 'Action Comics #421')
            if '#' in h2_text:
                _, issue_num = h2_text.split('#', 1)
                issue_num = issue_num.strip()
            else:
                issue_num = ""
        except NoSuchElementException:
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
        ws.append([series_name, issue_num, info, key_facts, year])
        print("Saved:", series_name, issue_num, info, key_facts, year)

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
