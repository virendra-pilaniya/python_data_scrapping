# Importing the required libraries
import os
import io
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image, ImageEnhance
from pytesseract import pytesseract

DOWNLOAD_DIR = "data/"

def preprocess_image(image):
    contrast_enhanced_image = ImageEnhance.Contrast(image).enhance(2.0)
    
    return contrast_enhanced_image

def scrape_data(company_name: str):
    '''
    Scrape data from the EPFO website
    '''

    options = Options()
    prefs = {"download.default_directory": os.path.join(os.getcwd(), DOWNLOAD_DIR)}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)

    try:
        time.sleep(5)
        # waiting to avoid any internet speed issue.
        driver.get('https://unifiedportal-epfo.epfindia.gov.in/publicPortal/no-auth/misReport/home/loadEstSearchHome')

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'estName')))

        est_name_input = driver.find_element(By.ID, 'estName')
        # filling company name
        est_name_input.send_keys(company_name)
        
        time.sleep(5)
        
        # captcha text extraction processing
        captcha_image = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'capImg'))
        )

        captcha_location = captcha_image.location
        captcha_size = captcha_image.size
        captcha_screenshot_path = os.path.join(os.getcwd(), DOWNLOAD_DIR, "captcha_screenshot.png")
        driver.save_screenshot(captcha_screenshot_path)

        captcha_image_screenshot = Image.open(captcha_screenshot_path)

        #Croping the captcha image using the location and size
        captcha_image_cropped = captcha_image_screenshot.crop((
            captcha_location['x'],
            captcha_location['y'],
            captcha_location['x'] + captcha_size['width'],
            captcha_location['y'] + captcha_size['height']
        ))

        captcha_image_cropped = preprocess_image(captcha_image_cropped)
        captcha_image_cropped_path = os.path.join(os.getcwd(), DOWNLOAD_DIR, "captcha.png")
        captcha_image_cropped.save(captcha_image_cropped_path)

        # Reading the captcha using pytesseract 
        pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'
        captcha_text = pytesseract.image_to_string(
            Image.open(captcha_image_cropped_path),
            lang='eng'
        )
        captcha_input = driver.find_element(By.ID, 'captcha')
        captcha_input.send_keys(captcha_text)
        
     
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'searchEmployer'))
        )
        search_button.click()

        WebDriverWait(driver, 10).until(EC.alert_is_present())
        
        alert = driver.switch_to.alert
        alert.accept()
        
        view_details_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'View Details')]")))
        
        driver.find_element_by_xpath("//a[contains(text(),'View Details')]")
        
        # clicking on view details link
        view_details_link.click()

        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

        view_payment_details_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@title, 'Click to view payment details.')]/u[text()='View Payment Details']"))
        )
        
        # clicking on view payment details link
        view_payment_details_link.click()

        excel_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.dt-button.buttons-excel.buttons-html5"))
        )
        excel_button.click()

        WebDriverWait(driver, 30).until(
            lambda x: os.path.exists(os.path.join(os.getcwd(), 'Payment Details.xlsx'))
        )
        
        #downloading payment_details file
        file_path = os.path.join(os.getcwd(), DOWNLOAD_DIR, 'Payment Details.xlsx')
        os.rename(os.path.join(os.getcwd(), 'Payment Details.xlsx'), file_path)
    
    except Exception as e:
        print(f"An error occurred: {e}")
        raise 
    
    finally:
        driver.quit()

def test_scrape_data():
    '''
    Test the scraped data
    '''
    # Convert xlsx file to csv due to some issues with pandas
    from xlsx2csv import Xlsx2csv
    Xlsx2csv("data/Payment Details.xlsx", outputencoding="utf-8").convert("payment_details.csv")

    df = pd.read_csv("payment_details.csv")

    assert set(df.columns) == set(['TRRN', 'Date Of Credit', 'Amount', 'Wage Month', 'No. of Employee', 'ECR'])
    assert df['TRRN'].loc[0] == 3171702000767
    assert df['Date Of Credit'].loc[0] == '03-FEB-2017 14:35:15'
    assert df['Amount'].loc[0] == 334901
    assert df['Wage Month'].loc[0] == 'DEC-16'
    assert df['No. of Employee'].loc[0] == 83
    assert df['ECR'].loc[0] == 'YES'
    print("All tests passed!")

def main():
    print("Hello World!")

    scrape_data("MGH LOGISTICS PVT LTD")

    # Uncomment the following tests whenever scraping is completed.
    # test_scrape_data()
    # Feel free to add any edge cases which you might think are helpful

if __name__ == "__main__":
    main()