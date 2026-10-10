from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re

# ==================== INITIAL SETTINGS ====================
print("=" * 60)
print("Google Maps Lead Extractor (Unified + Fast Version)")
print("=" * 60)

# 1. Ask for number of scrolls
while True:
    try:
        scroll_times = int(input("\nHow many times should I scroll? (e.g. 5, 10, 15): "))
        if scroll_times > 0:
            break
        else:
            print("Please enter a positive number.")
    except ValueError:
        print("Please enter a valid number.")

# 2. Ask for output filename
output_filename = input("\nEnter the output Excel filename (e.g. leads_dubai.xlsx): ").strip()
if not output_filename.endswith(".xlsx"):
    output_filename += ".xlsx"

print(f"\nScroll times : {scroll_times}")
print(f"Output file  : {output_filename}")

# Ask if user wants to extract emails
while True:
    extract_email = input("\nDo you want to extract Emails from websites? (y/n): ").strip().lower()
    if extract_email in ['y', 'yes', 'n', 'no']:
        break
    print("Please enter y or n")

extract_email = extract_email in ['y', 'yes']
print(f"Extract Emails: {'Yes' if extract_email else 'No'}")
print("-" * 60)

# ==================== START BROWSER ====================
print("\nOpening Google Maps...")
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 8)

driver.get("https://www.google.com/maps")

# ==================== MANUAL SEARCH SECTION ====================
print("\n" + "=" * 60)
print("INSTRUCTIONS:")
print("1. Manually search on Google Maps (example: Restaurants in Dubai)")
print("2. Make sure the results list appears on the left side")
print("3. Come back here and press ENTER")
print("=" * 60 + "\n")

input("Ready? Press ENTER after you finished searching...")

# ==================== SCROLL + COLLECT LINKS ====================
print("\nScrolling and collecting business links...")

try:
    scrollable_div = driver.find_element(By.XPATH, "//div[@role='feed']")

    for i in range(scroll_times):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
        print(f"Scroll {i+1}/{scroll_times}")
        time.sleep(1.6)

    print("\nExtracting links...")
    results = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place')]")

    data = []
    seen_links = set()

    for item in results:
        name = item.get_attribute("aria-label")
        link = item.get_attribute("href")

        if name and link and link not in seen_links:
            seen_links.add(link)
            data.append({
                "Name": name,
                "Google Maps Link": link,
                "Phone": None,
                "Website": None,
                "Email": None
            })
            print(f"Found: {name}")

    print(f"\nTotal unique businesses found: {len(data)}")

except Exception as e:
    print(f"Error while collecting list: {e}")
    driver.quit()
    exit()

if len(data) == 0:
    print("No results found. Exiting...")
    driver.quit()
    exit()

# ==================== EXTRACT DETAILS ====================
print("\n" + "=" * 60)
print("Extracting Phone, Website and Email...")
print("=" * 60)

df = pd.DataFrame(data)

for index, row in df.iterrows():
    name = row["Name"]
    link = row["Google Maps Link"]

    print(f"\n[{index+1}/{len(df)}] {name}")

    try:
        driver.get(link)
        time.sleep(1.1)

        # ---------- Phone ----------
        phone_number = "Not Found"
        try:
            phone_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//button[contains(@data-item-id, 'phone:')]")
            ))
            phone_raw = phone_btn.get_attribute("aria-label")
            if phone_raw:
                phone_number = (phone_raw
                                .replace("Phone:", "")
                                .replace("الهاتف:", "")
                                .replace("Numéro de téléphone:", "")
                                .replace("de téléphone:", "")
                                .replace("تلفن:", "")
                                .strip())
                print(f"  Phone   : {phone_number}")
        except:
            print("  Phone   : Not Found")

        df.at[index, "Phone"] = phone_number

        # ---------- Website ----------
        website = "Not Found"
        try:
            website_elem = driver.find_element(By.CSS_SELECTOR, "a[data-item-id='authority']")
            website = website_elem.get_attribute("href")
            print(f"  Website : {website}")
        except:
            try:
                website_elem = driver.find_element(By.XPATH, "//a[contains(@data-item-id, 'authority')]")
                website = website_elem.get_attribute("href")
                print(f"  Website : {website}")
            except:
                print("  Website : Not Found")

        df.at[index, "Website"] = website

        # ---------- Email (optional) ----------
        email = "Not Found"
        if extract_email and website != "Not Found" and website.startswith("http"):
            try:
                driver.get(website)
                time.sleep(1.8)

                page_source = driver.page_source
                email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                emails_found = re.findall(email_pattern, page_source)

                blacklist = [
                    "example.com", "sentry.io", "wixpress.com", "cloudflare",
                    "schema.org", "google.com", "gstatic.com", "w3.org",
                    "yourdomain.com", "email.com", "domain.com"
                ]

                valid_emails = []
                for em in emails_found:
                    em_lower = em.lower()
                    if not any(bad in em_lower for bad in blacklist):
                        if em_lower not in [v.lower() for v in valid_emails]:
                            valid_emails.append(em)

                if valid_emails:
                    email = valid_emails[0]
                    print(f"  Email   : {email}")
                else:
                    print("  Email   : Not Found")

            except Exception:
                print("  Email   : Not Found (website error)")
        else:
            if not extract_email:
                print("  Email   : Skipped")

        df.at[index, "Email"] = email

    except Exception as e:
        print(f"  Error   : {str(e)[:70]}")
        df.at[index, "Phone"] = "Not Found"
        df.at[index, "Website"] = "Not Found"
        df.at[index, "Email"] = "Not Found"

    # Checkpoint every 10 businesses
    if (index + 1) % 10 == 0:
        df.to_excel(output_filename, index=False)
        print("  --> Progress saved")

    time.sleep(0.7)

# ==================== FINAL SAVE ====================
df.to_excel(output_filename, index=False)
driver.quit()

print("\n" + "=" * 60)
print("DONE!")
print(f"File saved as : {output_filename}")
print(f"Total leads   : {len(df)}")
print("=" * 60)