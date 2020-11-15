# %%
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from tqdm import tqdm
import time
import pandas as pd
import traceback
import sys
import re

url = 'https://cloud.land.gov.taipei/ImmPrice/TruePriceA.aspx'  # 台北地政雲網站
# webdriver位置(phantomjs)
webdriver_path = 'C:\\Program Files\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe'
driver = webdriver.PhantomJS(executable_path=webdriver_path)

# webdriver位置(chromedriver)
#webdriver_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe'
#driver = webdriver.Chrome(executable_path=webdriver_path)
driver.implicitly_wait(120)

district = '信義區'
positioning_method = '路段'
road = '基隆路二段'
transactional_type = '房地(土地+建物)'

driver.get(url)  # 連接到台北地政雲不動產價格資訊\買賣實價查詢網頁
#driver.refresh()

# 選行政區
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_MasterGond').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == district:
        option.click()
        time.sleep(1)
        break

# 選定位方式
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_locate').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == positioning_method:
        option.click()
        time.sleep(1)
        break

# 選路段
element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
    (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_GondRoad')))
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_GondRoad').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == road:
        option.click()
        time.sleep(0.5)
        break

# 選起始年
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '101':
        option.click()
        time.sleep(0.5)
        break

# 選起始月
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '08':
        option.click()
        time.sleep(0.5)
        break

# 預設的結束時間是資料的最新時間，所以不需設定

# 選交易類型
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionType').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == transactional_type:
        option.click()
        time.sleep(0.5)
        break

# 點選查詢
element = driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_btn_Search').click()


# %%
