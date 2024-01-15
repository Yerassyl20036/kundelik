import os
import time
import random
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService

def get_login_credentials(excel_path):
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    credentials = [(row[2], row[3], row[4]) for row in sheet.iter_rows(min_row=2, values_only=True)]
    return credentials


def add_homework(driver, class_name):
    todays_date = datetime.now().strftime("%d.%m.%Y")
    WebDriverWait(driver, 10).until(
        EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "#journal > tbody > tr"))
    )
    homework_file = class_name.lower() + '.txt'  # Assuming class_name is like 'bio', 'eng', etc.
    with open(homework_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    rows = driver.find_elements(By.CSS_SELECTOR, "#journal > tbody > tr")
    for row in rows:
        try:
            date_field = row.find_element(By.CLASS_NAME, 'td_lessonPlanning_date')
            if todays_date in date_field.text.strip():
                # print(f"Match found for today's date: {todays_date}")
                add_homework_element = row.find_element(By.CLASS_NAME, 'lesson-planning-table-homework__addhomework_readyHomework')
                add_homework_element.click()
                time.sleep(1)  # Adjust this time as necessary
                random_homework = random.choice(lines).strip()
                active_element = driver.switch_to.active_element
                time.sleep(1)
                active_element.send_keys(random_homework)
                time.sleep(1)
                active_element.send_keys(Keys.ENTER)
        except NoSuchElementException:
            continue

def click_checkbox(driver):
    try:
        checkbox = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "held"))
        )
        driver.execute_script("arguments[0].click();", checkbox)
        # print("Checkbox clicked via JavaScript.")
    except Exception as e:
        print(f"An error occurred while clicking checkbox: {e}")


def processing(driver, url, class_name):
    driver.execute_script(f"window.open('{url}');")
    driver.switch_to.window(driver.window_handles[-1])
    grey_panel_links = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'links'))
    )
    click_checkbox(driver)
    time.sleep(1)

    all_links_in_panel = grey_panel_links.find_elements(By.TAG_NAME, 'a')
    if all_links_in_panel:
        last_link = all_links_in_panel[-1]
        last_link.click()
    time.sleep(1)
    add_homework(driver, class_name)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

def automate(driver_path, username, password, class_name):
    try:
        chrome_options = ChromeOptions()
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-gpu')
        # chrome_options.add_argument('--headless')  # Uncomment if you want to run in headless mode
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--ignore-certificate-errors')

        service = ChromeService(executable_path=driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get('https://login.kundelik.kz/login')

        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/form/div[2]/div[4]/div[1]/div[1]/label/input').send_keys(username)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/form/div[2]/div[4]/div[2]/div[1]/label/input').send_keys(password)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/form/div[2]/div[4]/div[4]/div[1]/input').click()
        print(f"Starting automation for username: {username}")
        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, '_2ZuUI'))
        )
        lesson_links = [link.get_attribute('href') for link in driver.find_elements(By.CLASS_NAME, '_2ZuUI')]
        n = 0
        for url in lesson_links:
            n = n+1
            # print("Processing link: " + str(url)) 
            processing(driver, url, class_name)
    except TimeoutException:
        print(f"User '{username}' is skipped due to timeout.")
    finally:
        driver.quit()

if __name__ == "__main__":
    current_directory = os.path.dirname(os.path.abspath(__file__))
    path_to_chromedriver = os.path.join(current_directory, 'chromedriver.exe')
    excel_path = os.path.join(current_directory, 'ЖООrest.xlsx')
    credentials = get_login_credentials(excel_path)
    total_credentials = len(credentials)
    processed_credentials = 0
    for username, password, class_name in credentials:
        automate(path_to_chromedriver, username, password, class_name)
        processed_credentials += 1
        remaining_credentials = total_credentials - processed_credentials
        print(f"Completed automation for {processed_credentials} out of {total_credentials} usernames. {remaining_credentials} usernames remaining.")