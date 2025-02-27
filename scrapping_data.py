from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time
import traceback

from extract_cdsc import load_symbols

# Initialize the Chrome driver with improved notification blocking
from selenium.webdriver.chrome.options import Options
chrome_options = Options()

# More aggressive notification blocking
chrome_options.add_argument('--disable-notifications')
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--start-maximized')  # Maximize window to ensure elements are visible
prefs = {
    "profile.default_content_setting_values.notifications": 2,  # Block all notifications
    "profile.default_content_setting_values.popups": 2,         # Block all popups
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
company_names = load_symbols()
# Get the table header once
header_row = []
df_final = pd.DataFrame()

try:
    for company_name in company_names:
        try:
            # 1. Go to the URL
            driver.get(f"https://merolagani.com/CompanyDetail.aspx?symbol={company_name}")
            
            # Add a short delay to ensure page loads and any alerts have time to appear
            time.sleep(3)
            
            # 2. Handle any alerts that might appear
            try:
                alert = driver.switch_to.alert
                alert.dismiss()  # Dismiss the alert if it's present
                print(f"Alert dismissed for {company_name}")
            except:
                print(f"No alert found for {company_name}")
            
            # 3. Wait for the 'Price History' button to be clickable and then click it
            wait = WebDriverWait(driver, 10)
            price_history_button = wait.until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_CompanyDetail1_lnkHistoryTab"))
            )
            # Use JavaScript to click instead of the regular click method
            driver.execute_script("arguments[0].click();", price_history_button)
            print(f"Clicked Price History button for {company_name}")
            
            # 4. Handle any alerts that might appear after clicking
            try:
                alert = driver.switch_to.alert
                alert.dismiss()
                print(f"Post-click alert dismissed for {company_name}")
            except:
                print(f"No post-click alert found for {company_name}")
            
            # Initialize an empty list to store all data from all pages
            all_pages_data = []
            
            # Determine the number of pages
            current_page = 1
            has_more_pages = True
            
            # Wait for page to fully load
            time.sleep(3)
            
            while has_more_pages:
                print(f"Extracting data from page {current_page} for {company_name}...")
                
                # Wait for the table to be visible
                try:
                    table_element = wait.until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "table.table.table-bordered.table-striped.table-hover"))
                    )
                except TimeoutException:
                    print(f"Table not found on page {current_page} for {company_name}. Taking screenshot for debugging...")
                    driver.save_screenshot(f"debug_screenshot_{company_name}_page_{current_page}.png")
                    break
                
                # Extract rows from the table
                attempts = 0
                #result = False
                while attempts < 2:
                    try:
                        rows = table_element.find_elements(By.TAG_NAME, "tr")
                        break
                    except Exception as e:
                        print('Stale Element Exception found for tr tag')
                        print(e)
                    attempts +=1
                
                # Process header if on first page
                if current_page == 1 and not header_row:
                    header_cells = rows[0].find_elements(By.TAG_NAME, "th")
                    if header_cells:
                        header_row = [cell.text.strip() for cell in header_cells]
                        header_row.append("Symbol")  # Add Symbol column
                        rows = rows[1:]  # Skip the header row in further processing
                else:
                    # Skip header row on subsequent pages
                    rows = rows[1:] if rows and len(rows) > 0 else rows
                
                # Parse data cell by cell for current page
                page_data = []
                for row in rows:
                    attempts = 0
                    result = False
                    while attempts < 2:
                        try:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            result = True
                            break
                        except Exception as e:
                            print('StaleElementException')
                            print(e)
                        attempts +=1
                    if cells:
                        row_data = [cell.text.strip() for cell in cells]
                        page_data.append(row_data)
                        all_pages_data.append(row_data)
                
                print(f"Extracted {len(page_data)} rows from page {current_page} for {company_name}")
                
                # Try to go to the next page
                try:
                    # First, scroll to the bottom of the page to ensure pagination is visible
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(1)
                    
                    # Look for the next page link
                    next_page_links = driver.find_elements(By.CSS_SELECTOR, "a[title^='Page']")
                    next_page_found = False
                    
                    for link in next_page_links:
                        page_number = link.text.strip()
                        if page_number.isdigit() and int(page_number) == current_page + 1:
                            print(f"Attempting to navigate to page {current_page + 1} for {company_name}...")
                            
                            # Try scrolling the element into view first
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
                            time.sleep(1)
                            
                            # Try to click with JavaScript (safer than direct click)
                            try:
                                driver.execute_script("arguments[0].click();", link)
                                print(f"Clicked page {current_page + 1} using JavaScript")
                                current_page += 1
                                time.sleep(3)  # Wait longer for page to load
                                next_page_found = True
                                break
                            except Exception as click_error:
                                print(f"JavaScript click failed: {click_error}")
                                # If JavaScript click fails, try ActionChains as a fallback
                                try:
                                    actions = ActionChains(driver)
                                    actions.move_to_element(link).click().perform()
                                    print(f"Clicked page {current_page + 1} using ActionChains")
                                    current_page += 1
                                    time.sleep(3)
                                    next_page_found = True
                                    break
                                except Exception as action_error:
                                    print(f"ActionChains click failed: {action_error}")
                    
                    if not next_page_found:
                        print(f"No more pages found after page {current_page} for {company_name}")
                        has_more_pages = False
                
                except (NoSuchElementException, TimeoutException) as e:
                    print(f"Error finding next page for {company_name}: {e}")
                    has_more_pages = False
            
            # Convert all collected data to DataFrame
            if all_pages_data:
                df = pd.DataFrame(all_pages_data)
                df['Symbol'] = company_name
                
                # If we have headers, use them
                if header_row and len(header_row) == df.shape[1]:
                    df.columns = header_row
                
                # Append to final dataframe
                df_final = pd.concat([df_final, df], ignore_index=True)
                
                print(f"Data from all {current_page} pages for {company_name} successfully added to DataFrame.")
                print(f"Total rows collected for {company_name}: {len(all_pages_data)}")
            else:
                print(f"No data was collected for {company_name}.")

        except Exception as e:
            print(f"An error occurred for {company_name}: {e}")
            print(traceback.format_exc())  # Print full traceback for better debugging
            # Take a screenshot for debugging
            try:
                driver.save_screenshot(f"error_screenshot_{company_name}.png")
                print(f"Screenshot saved as 'error_screenshot_{company_name}.png'")
            except Exception as ss_error:
                print(f"Could not save screenshot: {ss_error}")
            
            # Continue with the next company instead of exiting
            continue

except Exception as outer_e:
    print(f"Fatal error in main loop: {outer_e}")
    print(traceback.format_exc())

finally:
    # Save whatever data we've collected so far
    if not df_final.empty:
        try:
            df_final.to_excel("history_all_pages.xlsx", index=False)
            print(f"Data successfully saved to 'history_all_pages.xlsx'. Total rows: {len(df_final)}")
        except Exception as save_error:
            print(f"Error saving to Excel: {save_error}")
            # Try CSV as fallback
            try:
                df_final.to_csv("history_all_pages.csv", index=False)
                print(f"Data saved to CSV as fallback. Total rows: {len(df_final)}")
            except:
                print("Failed to save data in any format.")
    
    # Close the driver
    try:
        driver.quit()
        print("WebDriver closed successfully")
    except:
        print("Error closing WebDriver")