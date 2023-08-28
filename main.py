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
# use Chrome driver
service_path = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service_path)
# get the page
emails = []
for i in range(1, 24): ## 24 the number of all pages / now is 24 (24 not included, so 23) but you can change here
    driver.get(f"https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/0-9/sekcia.fcgi?sid=1172&so=&page={i}&desc=&shops=&kraj=&okres=&cast=&attr=")
    time.sleep(1)
    try:
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[3]/button[3]').click()
    except NoSuchElementException:
        pass
    for i in range(1, 26): ##26 the actual elements on a page (26 not included, so 25)
        driver.find_element(By.XPATH, f'/html/body/div[3]/div/div[2]/div[2]/section/div[5]/div/ul/li[{i}]/div/div[2]/h2/a').click()

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
