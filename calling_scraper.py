from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from datetime import datetime
import os


# Load confidential information from Creds.txt
script_dir = os.path.dirname(os.path.abspath(__file__))
creds_path = os.path.join(script_dir, 'CallingsCreds.txt')
with open(creds_path, 'r') as f:
    creds_json_path = os.path.join(script_dir,f.readline().strip())
    spreadsheet_name = f.readline().strip()
    username = f.readline().strip()
    password = f.readline().strip()
    tab_name = f.readline().strip()

# Set up Google Sheets API credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json_path, scope)
client = gspread.authorize(creds)

# Open the Google Spreadsheet
sheet = client.open(spreadsheet_name).worksheet(tab_name)

try:
    # Setup WebDriver
    webdriver_service = Service(ChromeDriverManager().install())
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)
    driver.set_page_load_timeout(30)

    # Open the login page
    driver.get("https://lcr.churchofjesuschrist.org/report/custom-reports-details/97fb64b2-aa70-4166-93e1-c6decd332745")
    time.sleep(1)
    print("Launched login page")
    driver.find_element(By.ID, "input28").send_keys(username)  # Enter username
    driver.find_element(By.XPATH, '//input[@value="Next"]').click()  # Click "Next"
    time.sleep(1)
    print("Entered username and clicked 'Next'")
    driver.find_element(By.ID, "input53").send_keys(password)  # Enter password
    driver.find_element(By.XPATH, '//input[@value="Verify"]').click()  # Click "Sign In"
    time.sleep(1)
    print("Entered password and clicked 'Sign In'")

    # Launch Members with Callings report
    driver.get("https://lcr.churchofjesuschrist.org/orgs/members-with-callings?lang=eng")
    print("Got to report page")
    time.sleep(5)

    # Find the table and tbody elements within the specified class
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table ng-scope")]')
    tbody = table.find_element(By.TAG_NAME, 'tbody')

    # Find all tr elements within the tbody
    rows = tbody.find_elements(By.TAG_NAME, 'tr')

    # Collect all rows' data into a list
    all_member_data = []
    for row in rows:
        # Extract Name
        name_cell = row.find_element(By.CLASS_NAME, "first.n.fn")
        name_text = name_cell.text
        if "," in name_text:
            last_name, rest_name = name_text.split(", ", 1)
            first_name = rest_name.split(" ")[0]  # Only take the first part of the rest of the name
            name = f"{first_name} {last_name}"
        else:
            name = name_text

        # Extract Member ID
        member_id_link = name_cell.find_element(By.TAG_NAME, 'a').get_attribute('href')
        member_id_match = re.search(r'member-profile/(\d+)\?lang=', member_id_link)
        member_id = member_id_match.group(1) if member_id_match else ''

        # Extract Organization
        organization_cell = row.find_element(By.CLASS_NAME, "hidden-phone.organization.ng-binding")
        organization = organization_cell.text

        # Extract Calling
        calling_cell = row.find_element(By.CLASS_NAME, "position.ng-binding")
        calling = calling_cell.text

        # Extract Sustained Date
        sustained_date_cell = row.find_element(By.CLASS_NAME, "hidden-phone.sustained.nowrap.ng-binding")
        sustained_date_text = sustained_date_cell.text
        try:
            sustained_date = datetime.strptime(sustained_date_text, '%d %b %Y').strftime('%Y-%m-%d')
        except ValueError:
            sustained_date = ''

        # Extract Set Apart
        set_apart_cell = row.find_element(By.CLASS_NAME, "hidden-phone.set-apart")
        set_apart = 'Yes' if set_apart_cell.find_elements(By.TAG_NAME, 'img') else 'No'

        # Append member data
        member_data = {
            "name": name,
            "member_id": member_id,
            "organization": organization,
            "calling": calling,
            "sustained_date": sustained_date,
            "set_apart": set_apart,
            "class_name": None  # Placeholder for class name
        }
        all_member_data.append(member_data)

    # Launch the Primary Callings report
    driver.get("https://lcr.churchofjesuschrist.org/orgs/643828?lang=eng")
    print("Launched Primary callings page")
    time.sleep(2)

    # Click the "All Organizations" link
    all_orgs_link = driver.find_element(By.XPATH, '//a[@ng-click="selectAllOrgs()" and text()="All Organizations"]')
    all_orgs_link.click()
    print("Clicked 'All Organizations' link")
    time.sleep(2)

    # For each member with calling "Primary Teacher" or "Primary Activities Leader", find their class name
    for member in all_member_data:
        if member["calling"] in ["Primary Teacher", "Primary Activities Leader"]:
            try:
                # Look for the a tag with the member's member_id in the href
                member_links = driver.find_elements(By.XPATH, f'//a[contains(@href, "{member["member_id"]}")]')
                for member_link in member_links:
                    # Find the parent <sub-org> tag
                    sub_org = member_link.find_element(By.XPATH, './ancestor::sub-org')
                    # Find the first child div and the first h2 within that
                    primary_class_div = sub_org.find_element(By.TAG_NAME, 'div')
                    primary_class_h2 = primary_class_div.find_element(By.TAG_NAME, 'h2')
                    # Extract and clean the text
                    primary_class = primary_class_h2.text.strip()

                    # Skip if "Primary Activities Leader" and the class does not contain "Primary Activities"
                    if member["calling"] == "Primary Activities Leader" and "Primary Activities" not in primary_class:
                        continue

                    # Skip if not "Primary Activities Leader" and the class contains "Primary Activities"
                    if member["calling"] != "Primary Activities Leader" and "Primary Activities" in primary_class:
                        continue

                    # Save the class name to the member data
                    member["class_name"] = primary_class
                    break  # Stop once we have found a matching class name
            except Exception as e:
                print(f"Could not find primary class for member {member['name']}: {e}")

    # Launch the Sunday School Callings report
    driver.get("https://lcr.churchofjesuschrist.org/orgs/498982?lang=eng")
    print("Launched Sunday School callings page")
    time.sleep(2)

    # Click the "All Organizations" link
    all_orgs_link = driver.find_element(By.XPATH, '//a[@ng-click="selectAllOrgs()" and text()="All Organizations"]')
    all_orgs_link.click()
    print("Clicked 'All Organizations' link")
    time.sleep(2)

    # For each member with calling "Sunday School Teacher", find their class name
    for member in all_member_data:
        if member["calling"] == "Sunday School Teacher":
            try:
                # Look for the a tag with the member's member_id in the href
                member_link = driver.find_element(By.XPATH, f'//a[contains(@href, "{member["member_id"]}")]')
                # Find the parent <sub-org> tag
                sub_org = member_link.find_element(By.XPATH, './ancestor::sub-org')
                # Find the first child div and the first h2 within that
                class_div = sub_org.find_element(By.TAG_NAME, 'div')
                class_h2 = class_div.find_element(By.TAG_NAME, 'h2')
                # Extract and clean the text
                class_name = class_h2.text.strip()
                # Save the class name to the member data
                member["class_name"] = class_name
            except Exception as e:
                print(f"Could not find class for member {member['name']}: {e}")

    # Write updated data including class_name to the Google Spreadsheet
    sheet.clear()
    sheet.append_rows([list(member.values()) for member in all_member_data], value_input_option='USER_ENTERED')

    # Print successful completion message
    now = datetime.now()
    formatted_time = now.strftime("%A, %b %dth at %I:%M%p")
    print(f"Full Run Completed Successfully: {formatted_time}")

except Exception as e:
    now = datetime.now()
    formatted_time = now.strftime("%A, %b %dth at %I:%M%p")
    print(f"An error occurred: {e}")
    print(f"Full Run Failed: {formatted_time}")

finally:
    driver.quit()
