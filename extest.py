# %%
from urllib.request import urlopen
from bs4 import BeautifulSoup

html = urlopen('http://www.pythonscraping.com/pages/page3.html')
bs = BeautifulSoup(html, 'html.parser')

string = ''
for sibling in bs.find('table', {'id':'giftList'}).tr.next_siblings:
    #print([s for s in sibling.stripped_strings])
    string = string+str(sibling)

# %%
soup = BeautifulSoup(string, 'html.parser')
content = soup.find_all('tr', {'class':'gift'})

# %%
def get_prize():
    prize = []
    counter = 0
    while len(content) > counter:
        dataColumn = content[counter].find_all('td')[2].string.replace('\n', '')
        prize.append(dataColumn)
        
        counter += 1
    return prize
# %%
import pandas as pd

RoadData = pd.read_excel('路段.xlsx')


# %% 
def get_RoadList(District):
    RoadList = RoadData[District].dropna().tolist()
    return RoadList

# %%
from tqdm import tqdm
for i in tqdm(get_RoadList('松山區')):
    print(i)

# %%
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import traceback
import sys
import re

url = 'https://cloud.land.gov.taipei/ImmPrice/TruePriceA.aspx'  # 台北地政雲網站
webdriver_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe'
driver = webdriver.Chrome(executable_path=webdriver_path)
driver.implicitly_wait(120)

driver.get(url)  # 連接到台北地政雲不動產價格資訊\買賣實價查詢網頁
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_MasterGond').click()

for option in driver.find_elements_by_tag_name('option'):
    if option.text == '文山區':
        option.click()
        time.sleep(0.5)
        break

driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_locate').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '路段':
        option.click()
        time.sleep(0.5)
        break

element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
    (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_GondRoad')))
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_GondRoad').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '木柵路一段':
        option.click()
        time.sleep(0.5)
        break

# 選起始年
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '103':
        option.click()
        time.sleep(0.5)
        break
        
# 選起始月
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '11':
        option.click()
        time.sleep(0.5)
        break

driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionType').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '房地(土地+建物)':
        option.click()
        time.sleep(0.5)
        break

element = driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_btn_Search').click()

try:
    element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
                (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
    bs = BeautifulSoup(driver.page_source, 'html.parser')

    table = bs.find(id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
    page_info = table.find('td', {'colspan':'18'})

# 狀況1 資料只有一頁
    if page_info == None:
        print('只有一頁')
    elif re.search('最末頁', str(page_info)) != None:
        print('有最末頁按鈕')
    elif re.search('最末頁', str(page_info)) == None:
        print('無最末頁按鈕')

except UnexpectedAlertPresentException:
    error_info = sys.exc_info()
    error_msg = re.findall('\{.*\}', str(error_info))[0]

#except NoSuchElementException:

finally:
    driver.quit()


# %%
