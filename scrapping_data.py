from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException, NoSuchElementException
import pandas as pd
import time

# Initialize the Chrome driver with improved notification blocking
from selenium.webdriver.chrome.options import Options
chrome_options = Options()

# More aggressive notification blocking
chrome_options.add_argument('--disable-notifications')
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--disable-extensions')
prefs = {
    "profile.default_content_setting_values.notifications": 2,  # Block all notifications
    "profile.default_content_setting_values.popups": 2,         # Block all popups
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

try:
    # 1. Go to the URL
    driver.get("https://merolagani.com/CompanyDetail.aspx?symbol=KBSH#0")
    
    # Add a short delay to ensure page loads and any alerts have time to appear
    time.sleep(2)
    
    # 2. Handle any alerts that might appear
    try:
        alert = driver.switch_to.alert
        alert.dismiss()  # Dismiss the alert if it's present
        print("Alert dismissed")
    except:
        print("No alert found or already handled")
    
    # 3. Wait for the 'Price History' button to be clickable and then click it
    wait = WebDriverWait(driver, 10)
    price_history_button = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_CompanyDetail1_lnkHistoryTab"))
    )
    price_history_button.click()
    
    # 4. Handle any alerts that might appear after clicking
    try:
        alert = driver.switch_to.alert
        alert.dismiss()
        print("Post-click alert dismissed")
    except:
        print("No post-click alert found")
    
    # Initialize an empty list to store all data from all pages
    all_pages_data = []
    
    # Get the table header once
    header_row = []
    
    # Determine the number of pages
    current_page = 1
    has_more_pages = True
    
    while has_more_pages:
        print(f"Extracting data from page {current_page}...")
        
        # Wait for the table to be visible
        try:
            table_element = wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "table.table.table-bordered.table-striped.table-hover"))
            )
        except TimeoutException:
            print(f"Table not found on page {current_page}. Taking screenshot for debugging...")
            driver.save_screenshot(f"debug_screenshot_page_{current_page}.png")
            break
        
        # Extract rows from the table
        rows = table_element.find_elements(By.TAG_NAME, "tr")
        
        # Process header if on first page
        if current_page == 1 and not header_row:
            header_cells = rows[0].find_elements(By.TAG_NAME, "th")
            if header_cells:
                header_row = [cell.text.strip() for cell in header_cells]
                rows = rows[1:]  # Skip the header row in further processing
        else:
            # Skip header row on subsequent pages
            rows = rows[1:] if rows and len(rows) > 0 else rows
        
        # Parse data cell by cell for current page
        page_data = []
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if cells:
                row_data = [cell.text.strip() for cell in cells]
                page_data.append(row_data)
                all_pages_data.append(row_data)
        
        print(f"Extracted {len(page_data)} rows from page {current_page}")
        
        # Try to go to the next page
        try:
            # Look for the next page link
            next_page_links = driver.find_elements(By.CSS_SELECTOR, "a[title^='Page']")
            next_page_found = False
            
            for link in next_page_links:
                page_number = link.text.strip()
                if page_number.isdigit() and int(page_number) == current_page + 1:
                    print(f"Navigating to page {current_page + 1}...")
                    link.click()
                    current_page += 1
                    time.sleep(2)  # Wait for page to load
                    next_page_found = True
                    break
            
            if not next_page_found:
                print(f"No more pages found after page {current_page}")
                has_more_pages = False
        
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Error finding next page: {e}")
            has_more_pages = False
    
    # Convert all collected data to DataFrame and store in Excel
    if all_pages_data:
        df = pd.DataFrame(all_pages_data)
        
        # If we have headers, use them
        if header_row:
            df.columns = header_row
        
        df.to_excel("KBSH_price_history_all_pages.xlsx", index=False)
        print(f"Data from all {current_page} pages successfully saved to 'KBSH_price_history_all_pages.xlsx'.")
        print(f"Total rows collected: {len(all_pages_data)}")
    else:
        print("No data was collected.")

except Exception as e:
    print(f"An error occurred: {e}")
    # Take a screenshot for debugging
    try:
        driver.save_screenshot("error_screenshot.png")
        print("Screenshot saved as 'error_screenshot.png'")
    except:
        print("Could not save screenshot")

finally:
    driver.quit()