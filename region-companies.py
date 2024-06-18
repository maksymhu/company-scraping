from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
import random
import time
import os
import math
from webdriver_manager.chrome import ChromeDriverManager
import csv
import sys
import openpyxl
import numpy as np
import regex as re
import textwrap
from selenium.common.exceptions import NoSuchElementException
import requests 
import shutil
from urllib.parse import urlparse
import asyncio
from bs4 import BeautifulSoup

options = Options()
options.add_argument("--lang=en-US")
options.add_argument("window-size=1920,1080")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-extensions")
options.add_argument("--noerrdialogs")
options.user_data_dir = "./profiles"
options.headless = False
options.add_argument("--disable-session-crashed-bubble")


prefs = {
    "profile.default_content_settings.cookies": 2,
    "profile.default_content_setting_values.notifications": 2
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options = options)
driver_wait = WebDriverWait(driver, 10)

pd.options.display.max_rows = 99900

excel_files = [{'f': 'region_list', 's': 'region_companies', 'start': 0}]

for excel_file in excel_files:
    print('excel_file ' + excel_file['f'])
    companies = []

    df = pd.read_excel('region-lists/' + excel_file['f'] + '.xlsx')

    for row_i, row in df.iterrows():
        print('page_index ' + str(row_i))
        try:
            driver.get(row['link'])
            time.sleep(1)
        except Exception as error:
            print('driver.get')
            print(error)
            driver.execute_script("window.stop();")

        company_lists = driver.find_elements(By.CSS_SELECTOR, 'form#mainForm tbody tr') 
        print(len(company_lists))
        count = 0
        for index in range(len(company_lists)):
            company_list = company_lists[index]
            monthly_revenue = company_list.find_element(By.CSS_SELECTOR, 'td:nth-child(4)').text
            print(company_list.get_attribute('class'))
            if company_list.get_attribute('class') != 'no-hover':
                website = "https://" + company_list.find_element(By.CSS_SELECTOR, 'td:nth-child(2)').text
                company_name = ""
                telephone = ""
                address = ""
                email = ""
                try:
                    detail = driver_wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'form#mainForm tbody tr:nth-child(' + str(index + 1) + ')')))
                    driver.execute_script("arguments[0].click();", detail)
                    count = count + 1
                    detail_content = ""
                    time.sleep(1)
                    while True:
                        try:
                            detail_contents = driver.find_elements(By.CSS_SELECTOR, 'form#mainForm tbody tr.no-hover')
                            detail_content = detail_contents[len(detail_contents)-1]
                            break
                        except Exception as error:
                            print('Detail finding error')
                            continue
                        time.sleep(1)
                    
                    time.sleep(1)
                    detail_content = detail_content.text
                    st = detail_content.rfind("Company Name") + 13
                    ed = detail_content.rfind("Address")
                    ed2 = detail_content.rfind("Find People")
                    if ed2 != -1:
                        ed = ed2
                    company_name = detail_content[st:ed]
                    company_name = company_name.replace("\n", " ")
                    st = detail_content.rfind("Address") + 8
                    ed = detail_content.rfind("Telephone")
                    address = detail_content[st:ed]
                    address = address.replace("\n", " ")
                    st = detail_content.rfind("Telephone") + 10
                    ed = detail_content.rfind("Contacts")
                    telephone = detail_content[st:ed]
                    telephone = telephone.replace("\n", " ")
                    st = detail_content.rfind("Compliant Emails") + 17
                    ed = detail_content.rfind("Traffic Ranking")
                    email = detail_content[st:ed]
                    email = email.replace("\n", " ")
                    if monthly_revenue != "":
                        company = {                
                            'website': website,
                            'monthly_revenue': monthly_revenue,
                            'company_name': company_name,
                            'telephone': telephone, 
                            'address': address,
                            'email': email,
                        }
                        companies.append(company)
                        print(company)
                except Exception as error:
                    print('Detail Error')
                    continue
        try:
            product_links_df = pd.DataFrame(companies)
            product_links_df.to_excel("region-companies/" + excel_file['s'] + ".xlsx", index=False)

        except Exception as error:
            print('Writing Exel')

        # try:
        #     detail = driver_wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'form#mainForm tbody tr:nth-child(' + index + ')')))
        #     driver.execute_script("arguments[0].click();", reply)
        # except Exception as error:
        #     print('Reply')
        #     continue
        
        # while True:
        #     try:
        #         email = driver_wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.reply-option-label-container span')))
        #         driver.execute_script("arguments[0].click();", email)
        #         break
        #     except Exception as error:
        #         print('Email')
        #         continue    

        # time.sleep(2)

        # email_index = driver.find_elements(By.CSS_SELECTOR, 'div.reply-email-address a')
        # email_href = email_index[0].get_attribute('href')
        # start_index = 7
        # end_index = email_href.rfind("?")
        # email_content = email_href[start_index:end_index]

        # try:
        #     price = driver.find_element(By.CSS_SELECTOR, 'h1.postingtitle span.postingtitletext span.price').text
        #     titles = driver.find_elements(By.CSS_SELECTOR, 'div.mapAndAttrs div.important span.valu')
        #     title = titles[0].text + " " + titles[1].text
        #     link_to_car = row['link']
        #     distance = driver.find_element(By.CSS_SELECTOR, 'div.mapAndAttrs div.auto_miles span.valu').text
        #     distance = distance + "KM"

        #     product = {                
        #         'title': title,
        #         'email': ' ',
        #         ' ': email_content,
        #         'miles': distance, 
        #         'price': price,
        #         'link': link_to_car,
        #         }
            
        #     products.append(product)
        # except Exception as error:
        #     print('Getting Details')
        #     print(error)
        #     continue
        
        
