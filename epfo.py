import os
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
from io import BytesIO
import easyocr
import time

# URL for EPFO search page
EPFO_SEARCH_URL = "https://unifiedportal-emp.epfindia.gov.in/publicPortal/no-auth/misReport/home/loadEstSearchHome"

def setup_driver(download_dir):
    """Set up the Selenium WebDriver with download directory options."""
    options = Options()
    prefs = {"download.default_directory": download_dir}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return driver

def solve_captcha(driver):
    """Solve CAPTCHA using EasyOCR."""
    reader = easyocr.Reader(lang_list=['en'], gpu=False)
    captcha_image = driver.find_element(By.ID, "capImg").screenshot_as_png
    captcha_img = Image.open(BytesIO(captcha_image))
    
    # OCR to extract text
    extracted_text = reader.readtext(np.array(captcha_img), detail=0, paragraph=True)
    if not extracted_text:
        return None
    captcha_text = extracted_text[0].replace(" ", "").upper()  # Simplify text
    return captcha_text

def search_and_download_excel(driver, company_name, download_dir):
    """Search for the company and download the Excel file."""
    driver.get(EPFO_SEARCH_URL)
    time.sleep(2)

    # Create folder if it doesn't exist
    os.makedirs(download_dir, exist_ok=True)
    file_name = os.path.join(download_dir, f"{company_name.replace(' ', '_')}.xls")

    # Check if file already exists
    if os.path.exists(file_name):
        print(f"File already exists: {file_name}")
        return file_name

    retry = True
    while retry:
        # Enter the company name
        driver.find_element(By.NAME, "estName").clear()  # Clear previous input
        driver.find_element(By.NAME, "estName").send_keys(company_name)

        # Solve the CAPTCHA
        captcha_text = solve_captcha(driver)
        if not captcha_text:
            print("Failed to extract CAPTCHA text. Retrying...")
            driver.refresh()
            time.sleep(2)
            continue

        driver.find_element(By.NAME, "captcha").clear()
        driver.find_element(By.NAME, "captcha").send_keys(captcha_text)

        # Submit the search form
        driver.find_element(By.NAME, "Search").click()

        try:
            # Wait for the Excel button to be clickable
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="example_wrapper"]/div[1]/a/span'))
            )
            print("CAPTCHA solved successfully and ready to download Excel!")

            # Click the Excel button to download
            driver.find_element(By.XPATH, '//*[@id="example_wrapper"]/div[1]/a/span').click()

            # Wait for file download to complete
            time.sleep(5)  # Adjust as needed based on file size and download speed
            print(f"File downloaded: {file_name}")
            retry = False
        except TimeoutException:
            print("Invalid CAPTCHA or Excel button not clickable. Retrying...")

    return file_name

def main():
    """Main function to search for a company and download its Excel data."""
    company_name = "TATA MOTORS"
    download_dir = os.path.join(os.getcwd(), "CompanyList")
    driver = setup_driver(download_dir)

    try:
        print(f"Searching for company: {company_name}")
        file_path = search_and_download_excel(driver, company_name, download_dir)
        print(f"Excel file saved at: {file_path}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
