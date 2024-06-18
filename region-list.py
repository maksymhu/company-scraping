from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
import time
from webdriver_manager.chrome import ChromeDriverManager
import json
import numpy as np
import regex as re
from selenium.common.exceptions import NoSuchElementException

options = Options()
options.add_argument("--lang=en-US")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-extensions")
options.add_argument("--noerrdialogs")
options.headless = False
options.add_argument("--disable-session-crashed-bubble")

prefs = {
    "profile.default_content_settings.cookies": 2,
    "profile.default_content_setting_values.notifications": 2
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options = options)
driver_wait = WebDriverWait(driver, 10)

page_url = f'https://trends.builtwith.com/websitelist/Shopware-5/Germany/'

regions = [
    "Nordrhein-Westfalen",
    "Bayern",
    "Baden-Württemberg",
    "Niedersachsen",
    "Hessen",
    "Sachsen",
    "Schleswig-Holstein",
    "Rheinland-Pfalz",
    "Berlin",
    "Hamburg",
    "Thüringen",
    "Brandenburg",
    "Sachsen-Anhalt",
    "Saarland",
    "Mecklenburg-Vorpommern",
    "Bremen"
]

page_lists = []
for region_index in range(16):
    print("Page: " + str(region_index))
    driver.get(page_url + regions[region_index])
    
    time.sleep(1)
    sub_lists = driver.find_elements(By.CSS_SELECTOR, 'form#mainForm div.container div.row:nth-child(4) div.card-body div.row div.col-md-3 div.mb-4 h5 a')
    if len(sub_lists) == 0:
        page_lists.append({'link': page_url + regions[region_index]})
    else:
        for sub_list in sub_lists:
            try:
                page_lists.append({'link': sub_list.get_attribute('href')})
            except Exception as error:
                print('finding link')

try:
    print(len(page_lists))
    page_lists_df = pd.DataFrame(page_lists)
    excel_file_name = f"region-lists/region_list.xlsx"
    page_lists_df.to_excel(excel_file_name, index=False) 
    print(f"Excel file saved as {excel_file_name}")

except Exception as error:
    print('Writing Exel')