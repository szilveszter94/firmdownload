import math
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
from selenium.webdriver.common.keys import Keys

# give the chrome driver location (you can download it from "https://chromedriver.chromium.org/")
chrome_driver_path = "D:/Games/chromedriver_win32/chromedriver.exe"
## here you must provide the base link, and after the '&page=' you must add a '{}' because it's a dynamic url
base_link = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/0-9/sekcia.fcgi?sid=1172&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="
# use Chrome driver
service_path = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service_path)
# get the page
emails = []
driver.get(base_link.format(1))
time.sleep(1)
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[3]/button[3]').click()
except NoSuchElementException:
    pass
page_number_string = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[2]/ul/li[4]/span/small').text
page_number = int(page_number_string.strip("()"))
correct_page_number = math.ceil(page_number / 25)
for page_num in range(1, correct_page_number + 1): ##  the number of all pages (the +1 is because the loop run from 1 - e.g 24, but the 24 is not included)
    driver.get(base_link.format(page_num))
    time.sleep(1)
    for firm_num in range(1, 26): ##26 the actual elements on a page (26 not included, so 25)
        try:
            driver.find_element(By.XPATH, f'/html/body/div[3]/div/div[2]/div[2]/section/div[5]/div/ul/li[{firm_num}]/div/div[2]/h2/a').click()
        except NoSuchElementException:
            break
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(1)
        try:
            emailAddress = driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/section/div[4]/div[2]/a').text
            print(emailAddress)
            emails.append([emailAddress])
        except NoSuchElementException:
            print([["email", "no email"]])
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(0.5)
for row_data in emails:
    sheet.append([str(cell) for cell in row_data])
workbook.save("output1.xlsx")
driver.quit()
