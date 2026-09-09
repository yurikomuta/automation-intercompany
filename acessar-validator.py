from selenium import webdriver
from selenium.webdriver.common.by import By
import time

#1. acessa o navigatorbr.accenture.com
url = "https://navigatorbr.accenture.com"
driver = webdriver.Chrome() 

driver.get(url)
time.sleep(8)

#2. busca o botão de login e clica nele
click_button = driver.find_element(By.XPATH, "//*[@id='loginAccenture']")
click_button.click()
time.sleep(5)
print("clicked")
time.sleep(20)

#3. aguarda os campo de usuário e senha receberem os valores

#4. 

driver.quit()
