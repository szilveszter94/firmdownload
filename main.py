import math
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import openpyxl

# ------------------------------------------ HERE YOU CAN CHANGE CONDITIONS ------------------------------------ #

FIRST_PAGE = 1  # set the first page, from which page you want to start the program
LAST_PAGE = False  # provide last page if you don't want to loop through all pages
PAGE_LOAD = 0.5  # set the load times (second)
FIRM_DETAIL_PAGE_LOAD = 0.5  # set the load times (second)
ITEMS_PER_PAGE = 25  # items / page
EXCEL_FILE_NAME = "email_by_name_0_9.xlsx"  # here you must provide the Excel file name
# here you must provide the base link, and after the '&page=' you must add a '{}' because it's a dynamic url
BASE_LINK = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/A/sekcia.fcgi?sid=1173&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="

# ---------------------------------------------------------------------------------------------------------------- #

# give the chrome driver location (you can download it from "https://chromedriver.chromium.org/")
chrome_driver_path = "D:/Games/chromedriver_win32/chromedriver.exe"

# use Chrome driver
service_path = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service_path)

# get the page
emails = []
driver.get(BASE_LINK.format(FIRST_PAGE))
time.sleep(PAGE_LOAD)
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[3]/button[3]').click()
except NoSuchElementException:
    pass
if not LAST_PAGE:
    page_number_string = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[2]/ul/li[4]/span/small').text
    page_number = int(page_number_string.strip("()"))
    LAST_PAGE = math.ceil(page_number / ITEMS_PER_PAGE)
for page_num in range(FIRST_PAGE,
                      LAST_PAGE + 1):  # the number of all pages (the +1 is because the loop run from 1 - e.g 24, but the 24 is not included)
    driver.get(BASE_LINK.format(page_num))
    time.sleep(PAGE_LOAD)
    for firm_num in range(1, ITEMS_PER_PAGE + 1):  # 26 the actual elements on a page (26 not included, so 25)
        try:
            driver.find_element(By.XPATH,
                                f'/html/body/div[3]/div/div[2]/div[2]/section/div[5]/div/ul/li[{firm_num}]/div/div[2]/h2/a').click()
        except NoSuchElementException:
            break
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(FIRM_DETAIL_PAGE_LOAD)
        try:
            emailAddress = driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/section/div[4]/div[2]/a').text
            print(emailAddress)
            if "@" in emailAddress:
                emails.append([emailAddress])
        except NoSuchElementException:
            print("no email")
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(FIRM_DETAIL_PAGE_LOAD)

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
# write the Excel file
for row_data in emails:
    sheet.append([str(cell) for cell in row_data])
workbook.save(EXCEL_FILE_NAME)
driver.quit()
