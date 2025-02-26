from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException, NoSuchElementException
import pandas as pd
import time

#initialize the chrome driver
from selenium.webdriver.chrome.options import Options

chrome_options = Options()


chrome_options.add_argument('--disable-notifications')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-popup-blocking')

prefs = {
    "profile.default_content_setting_values.notifications": 2,  # Block all notifications
    "profile.default_content_setting_values.popups": 2,
}

chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)


try:
    driver.get('https://www.sharesansar.com/cdsc')
    time.sleep(2)

    wait = WebDriverWait(driver, 10)
    all_pages_data = []
    header_row = []

    current_page = 1
    has_more_pages = True

    while has_more_pages:
        print(f'Extracting data from page {current_page}')

        #Wait for the table to be visible
        try:
            table_element = wait.until(
                EC.visibility_of_element_located((By.ID, 'myTable')))
        except TimeoutException:
            print((f'Table not found on page {current_page}. Taking screenshot of the current page'))
            driver.save_screenshot(f'debug_screensht_page_{current_page}')
            break
        
        #Extract rows from the table:
        if current_page == 1:
            header = table_element.find_elements(By.TAG_NAME, 'thead')
            for header_cells in header:
                header_cell = header_cells.find_elements(By.TAG_NAME, 'th')
            if header_cell:
                header_row = [cell.text.strip() for cell in header_cell]
        
        rows = table_element.find_elements(By.TAG_NAME, 'tr')

        rows = rows[1:] if rows and len(rows) > 0 else rows

        #Scrape the data cell by cell for current page:
        page_data = []
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if cells:
                row_data = [cell.text.strip() for cell in cells]
                page_data.append(row_data)
                all_pages_data.append(row_data)
        print(f'Extracted {len(page_data)} rows from page {current_page}')


        try:
            next_page__link = driver.find_elements(By.ID, 'myTable_next')
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'myTable_next')))
            next_page_found = False
            for link in next_page__link:
                try:
                    # Wait until the element is clickable
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'myTable_next')))
                    print(f'Navigating to page {current_page + 1}')
                    link.click()
                    current_page += 1
                    time.sleep(2)
                    next_page_found = True
                    break
                except:
                    print("Next page link is not clickable. Exiting loop.")
                    next_page_found = False
                    break  # Exit loop if the element is not clickable
        except Exception as e:
            print(e)
        
        if not next_page_found:
            print(f'No more pages found after page {current_page}')
            has_more_pages = False
    print('This is the data', all_pages_data)
    print('This is the header row', header_row)
    if all_pages_data:
        df = pd.DataFrame(all_pages_data)
        
        # If we have headers, use them
        if header_row:
            df.columns = header_row
        filename = 'CDSC.xlsx'
        df.to_excel(filename, index=False)
        print(f"Data from all {current_page} pages successfully saved to {filename}.")
        print(f"Total rows collected: {len(all_pages_data)}")
    else:
        print("No data was collected.")
except Exception as e:
    print(e)
finally:
    driver.quit()