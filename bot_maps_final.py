from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd  # Don't forget to import the pandas library
import time

# 1. Start the browser
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

print("🌍 Opening Google Maps...")
driver.get("https://www.google.com/maps")

# ---------------------------------------------------------
# Manual section (Human Mode)
# ---------------------------------------------------------
print("\n" + "="*50)
print("🛑 Manager instructions:")
print("1. Search manually for (Restaurants in Dubai) or (Dentists in Riyadh)...")
print("2. Make sure the sidebar list appears.")
print("3. Come back here and press ENTER.")
print("="*50 + "\n")

input("⌨️ I'm ready.. press Enter after you're done searching...")

# ---------------------------------------------------------
# Automated section: scrolling and scraping
# ---------------------------------------------------------
print("🤖 Working... I will first load more results.")

try:
    # 1. Get the list container
    scrollable_div = driver.find_element(By.XPATH, "//div[@role='feed']")
    
    # 2. Scroll (do it 5 times; you can increase it to 10 or 20 to gather hundreds of results)
    for i in range(5):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
        print(f"🔄 Scroll {i+1}...")
        time.sleep(3)  # Waiting here is very important

    print("🛑 Scrolling finished. Now extracting the data...")
    
    # 3. Extract the data
    # Trick: we search for all links containing "google.com/maps/place"
    # because these are the business links only
    results = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place')]")

    print(f"📦 Found {len(results)} places. Extracting details...")

    data = []  # The basket

    for item in results:
        # Google stores the business name inside the aria-label attribute of the link
        name = item.get_attribute("aria-label")
        
        # The link itself
        link = item.get_attribute("href")

        # Make sure the name is not empty (sometimes hidden links appear)
        if name and link:
            # Clean the data: sometimes the name is very long
            print(f"📌 Extracting: {name}")
            
            data.append({
                'Name': name,
                'Google Maps Link': link
            })

    # 4. Save to Excel
    if len(data) > 0:
        df = pd.DataFrame(data)
        df.to_excel("google_maps_leads.xlsx", index=False)
        print(f"\n✅ Saved {len(data)} businesses to 'google_maps_leads.xlsx' successfully!")
    else:
        print("⚠️ No data was found to extract.")

except Exception as e:
    print(f"❌ An error occurred: {e}")

print("Done. Press Enter to close.")
input()
driver.quit()