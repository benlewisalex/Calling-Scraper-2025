from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import re
import gspread
from google.oauth2.service_account import Credentials
import time
from datetime import datetime
import os
import json

# Load confidential information from environment variables
spreadsheet_name = os.getenv('SPREADSHEET_NAME')
username = os.getenv('LDS_USERNAME')
password = os.getenv('LDS_PASSWORD')
tab_name = os.getenv('TAB_NAME')
remote_webdriver_url = os.getenv("REMOTE_SELENIUM_GRID_URL")

# Load Google Sheets API credentials from environment variable
google_creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
google_creds_dict = json.loads(google_creds_json)
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_info(google_creds_dict, scopes=scope)
client = gspread.authorize(creds)

# Open the Google Spreadsheet
sheet = client.open(spreadsheet_name).worksheet(tab_name)

try:
    # Setup Remote WebDriver with increased timeouts
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.set_capability("browserName", "chrome")

    driver = webdriver.Remote(
        command_executor=f"{remote_webdriver_url}/wd/hub",
        options=chrome_options
    )
    driver.set_page_load_timeout(60)
    driver.set_script_timeout(60)
    driver.implicitly_wait(10)

    # Open the login page
    driver.get("https://lcr.churchofjesuschrist.org/report/custom-reports-details/97fb64b2-aa70-4166-93e1-c6decd332745")
    time.sleep(2)
    print("Launched login page")
    driver.find_element(By.ID, "input28").send_keys(username)
    driver.find_element(By.XPATH, '//input[@value="Next"]').click()
    time.sleep(2)
    print("Entered username and clicked 'Next'")
    driver.find_element(By.ID, "input53").send_keys(password)
    driver.find_element(By.XPATH, '//input[@value="Verify"]').click()
    print("Entered password and clicked 'Sign In'")

    # 🚀 SKIP loading dashboard (go directly to Members with Callings report)
    driver.get("https://lcr.churchofjesuschrist.org/orgs/members-with-callings?lang=eng")
    print("Navigated directly to Callings report page")
    time.sleep(5)

    table = driver.find_element(By.XPATH, '//table[contains(@class, "table ng-scope")]')
    tbody = table.find_element(By.TAG_NAME, 'tbody')
    rows = tbody.find_elements(By.TAG_NAME, 'tr')

    all_member_data = []
    for row in rows:
        name_cell = row.find_element(By.CLASS_NAME, "first.n.fn")
        name_text = name_cell.text
        if "," in name_text:
            last_name, rest_name = name_text.split(", ", 1)
            first_name = rest_name.split(" ")[0]
            name = f"{first_name} {last_name}"
        else:
            name = name_text

        member_id_link = name_cell.find_element(By.TAG_NAME, 'a').get_attribute('href')
        member_id_match = re.search(r'member-profile/(\d+)\?lang=', member_id_link)
        member_id = member_id_match.group(1) if member_id_match else ''

        organization = row.find_element(By.CLASS_NAME, "hidden-phone.organization.ng-binding").text
        calling = row.find_element(By.CLASS_NAME, "position.ng-binding").text

        sustained_date_text = row.find_element(By.CLASS_NAME, "hidden-phone.sustained.nowrap.ng-binding").text
        try:
            sustained_date = datetime.strptime(sustained_date_text, '%d %b %Y').strftime('%Y-%m-%d')
        except ValueError:
            sustained_date = ''

        set_apart_cell = row.find_element(By.CLASS_NAME, "hidden-phone.set-apart")
        set_apart = 'Yes' if set_apart_cell.find_elements(By.TAG_NAME, 'img') else 'No'

        member_data = {
            "name": name,
            "member_id": member_id,
            "organization": organization,
            "calling": calling,
            "sustained_date": sustained_date,
            "set_apart": set_apart,
            "class_name": None
        }
        all_member_data.append(member_data)

    # Primary report
    driver.get("https://lcr.churchofjesuschrist.org/orgs/643828?lang=eng")
    time.sleep(3)
    driver.find_element(By.XPATH, '//a[@ng-click="selectAllOrgs()" and text()="All Organizations"]').click()
    time.sleep(3)

    for member in all_member_data:
        if member["calling"] in ["Primary Teacher", "Primary Activities Leader"]:
            try:
                member_links = driver.find_elements(By.XPATH, f'//a[contains(@href, "{member["member_id"]}")]')
                for member_link in member_links:
                    sub_org = member_link.find_element(By.XPATH, './ancestor::sub-org')
                    primary_class = sub_org.find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'h2').text.strip()

                    if member["calling"] == "Primary Activities Leader" and "Primary Activities" not in primary_class:
                        continue
                    if member["calling"] != "Primary Activities Leader" and "Primary Activities" in primary_class:
                        continue

                    member["class_name"] = primary_class
                    break
            except Exception as e:
                print(f"Could not find primary class for member {member['name']}: {e}")

    # Sunday School report
    driver.get("https://lcr.churchofjesuschrist.org/orgs/498982?lang=eng")
    time.sleep(3)
    driver.find_element(By.XPATH, '//a[@ng-click="selectAllOrgs()" and text()="All Organizations"]').click()
    time.sleep(3)

    for member in all_member_data:
        if member["calling"] == "Sunday School Teacher":
            try:
                member_link = driver.find_element(By.XPATH, f'//a[contains(@href, "{member["member_id"]}")]')
                sub_org = member_link.find_element(By.XPATH, './ancestor::sub-org')
                class_name = sub_org.find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'h2').text.strip()
                member["class_name"] = class_name
            except Exception as e:
                print(f"Could not find class for member {member['name']}: {e}")

    # Write to Google Sheet
    sheet.clear()
    sheet.append_rows([list(member.values()) for member in all_member_data], value_input_option='USER_ENTERED')

    now = datetime.now()
    formatted_time = now.strftime("%A, %b %dth at %I:%M%p")
    print(f"Full Run Completed Successfully: {formatted_time}")

except Exception as e:
    now = datetime.now()
    formatted_time = now.strftime("%A, %b %dth at %I:%M%p")
    print(f"An error occurred: {e}")
    print(f"Full Run Failed: {formatted_time}")

finally:
    try:
        driver.quit()
    except:
        pass
