from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from urllib.parse import urlparse
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Define wait duration
wait_duration = 10

# Define login credentials and URL
username = "atharvachoudhari06.truefunds@gmail.com"
password = "atharva123"
login_url = "https://www.valueresearchonline.com/funds/15864/icici-prudential-technology-fund-direct-plan/#other"

# Initialize WebDriver
driver = webdriver.Chrome()

# Open the login URL
driver.get(login_url)

# Wait for username field to be visible and fill it
username_field = WebDriverWait(driver, wait_duration).until(
    EC.visibility_of_element_located((By.ID, "username"))
)
username_field.send_keys(username)

# Wait for password field to be visible and fill it
password_field = WebDriverWait(driver, wait_duration).until(
    EC.visibility_of_element_located((By.ID, "password"))
)
password_field.send_keys(password)

# Wait for login button to be clickable and click it
login_button = WebDriverWait(driver, wait_duration).until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Login')]"))
)
login_button.click()
