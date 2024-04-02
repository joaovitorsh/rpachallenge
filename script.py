import pandas as pd
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import os

# access file
excel_path = os.path.join('.', 'challenge.xlsx')
df = pd.read_excel(excel_path)

# access url
driver = webdriver.Chrome()
driver.get('https://rpachallenge.com/')

# test if acces is succesfull
assert "Rpa Challenge" == driver.title, f"Title: {driver.title}"
start = driver.find_element(By.XPATH, "//button[contains(text(), 'Start')]")
start.click()

columns =  ['First Name', 'Last Name', 'Company Name', 'Role in Company', 'Address', 'Email', 'Phone Number']
for index, row in df.iterrows():
    for column in columns:
        tag = f"//rpa1-field[@ng-reflect-dictionary-value='{column}']/div/input"
        elem = driver.find_element(By.XPATH, tag)
        elem.send_keys(row[columns.index(column)])
    
    submit = driver.find_element(By.XPATH, "//input[@value='Submit']")
    submit.click()

# catch final message
result = driver.find_element(By.CLASS_NAME, 'congratulations')
message_success = result.find_element(By.XPATH, './div[@class="message1 teal-text text-darken-2"]').text
message_rating = result.find_element(By.XPATH, './div[@class="message2"]').text

print(message_success)
print(message_rating)
