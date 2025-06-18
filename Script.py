import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.firefox import GeckoDriverManager

df = pd.read_excel(#replace this)

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
driver.maximize_window()

driver.get(#replace this with the login page)

username = #replace this
password = # replace this

username_input = driver.find_element(By.XPATH, "//label[text()='Username']/preceding-sibling::input")
password_input = driver.find_element(By.XPATH, "//label[text()='Password']/preceding-sibling::input")

username_input.send_keys(username)
password_input.send_keys(password + Keys.RETURN)

print("Please complete the 2FA manually if prompted. You have 90 seconds...")
WebDriverWait(driver, 90).until(
    EC.url_contains("home") 
)

for index, row in df.iterrows():
    profile_id = row['profile_id']
    engagement  = row['engagement']

    print(f"\nOpening profile for {profile_id}")

    profile_url = f#replace this with page needing inputs
    driver.get(profile_url)

    try:
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//th[text()='Engagement Strategy']/following-sibling::td/select"
                )
            )
        )
        Select(dropdown_element).select_by_visible_text(engagement)
        print(f"Set Engagement Strategy to '{engagement}' for {profile_id}")
    except Exception as e:
        print(f"Could not set dropdown for {profile_id}: {e}")
    time.sleep(3)
driver.quit()
print("\nAll done!")
