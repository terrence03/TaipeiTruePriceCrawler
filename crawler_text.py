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
#webdriver_path = 'C:\\Program Files\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe'
#driver = webdriver.PhantomJS(executable_path=webdriver_path)

# webdriver位置(chromedriver)
webdriver_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe'
driver = webdriver.Chrome(executable_path=webdriver_path)
driver.implicitly_wait(120)

district = '信義區'
positioning_method = '路段'
road = '基隆路二段'
transactional_type = '房地+房地車'

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
    (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_txb_GondRoad')))
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_txb_GondRoad').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == road:
        option.click()
        time.sleep(0.5)
        break

# 選起始年
driver.find_element_by_id(
    'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear').click()
for option in driver.find_elements_by_tag_name('option'):
    if option.text == '108':
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

'''
開始爬資料，這裡會有幾種情況
1. 資料只有一頁
2. 資料有很多頁（超過10頁）
3. 資料不超過10頁
4. 查無資料
'''

try:
    element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
        (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
    bs = BeautifulSoup(driver.page_source, 'html.parser')

    table = bs.find(
        id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
    page_info = table.find('td', {'colspan': '19'})

    # 狀況1 資料只有一頁
    if page_info == None:
        print('只有一頁')
        print(page_info)
        bs = BeautifulSoup(driver.page_source, 'html.parser')
        #get_ColumnsData(bs)

    # 狀況2 資料有很多頁（超過10頁）
    elif re.search('最末頁', str(page_info)) != None:
        #print('有最末頁按鈕')
        # 等待第一頁資料出來
        element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
        bs = BeautifulSoup(driver.page_source, 'html.parser')

        # 先翻到最末頁確認總頁數
        driver.find_element_by_link_text('最末頁').click()
        element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
            (By.LINK_TEXT, '第一頁')))
        bs = BeautifulSoup(driver.page_source, 'html.parser')
        table = bs.find(
            id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
        for row in table.find('td'):
            last_page = int([s for s in row.stripped_strings][-1])
            print('最末頁' ,last_page)
    
        # 回到第一頁
        driver.find_element_by_link_text('第一頁').click()
        element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
            (By.LINK_TEXT, '最末頁')))
        time.sleep(1)

        # 從第一頁開始儲存資料
        bs = BeautifulSoup(driver.page_source, 'html.parser')
        #get_ColumnsData(bs)

        # 翻頁繼續儲存資料
        now_page = i = 1
        while i <= last_page-1:
            next_page = i + 1
            print(i)
            if next_page == 11:  # 若下一頁為11，點擊'...'的按鈕
                driver.find_element_by_link_text('...').click()
                element = WebDriverWait(driver, 60).until(expected_conditions.element_to_be_clickable(
                    (By.LINK_TEXT, '第一頁')))
                time.sleep(1)
                bs = BeautifulSoup(driver.page_source, 'html.parser')
                #get_ColumnsData(bs)

            # 若下一頁為[21,31,41,51,61,71,81,91,101,]點擊td[13]的超連接
            elif next_page in [21, 31, 41, 51, 61, 71, 81, 91, 101, 111, 121, 131, 141, 151]:
                driver.find_element_by_xpath(
                    '//*[@id = "ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice"]/tbody/tr[1]/td/table/tbody/tr/td[13]/a').click()
                element = WebDriverWait(driver, 20).until_not(
                    expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page))))
                time.sleep(1)
                bs = BeautifulSoup(driver.page_source, 'html.parser')
                #get_ColumnsData(bs)

            else:
                driver.find_element_by_link_text(
                    str(next_page)).click()  # 正常換頁
                element = WebDriverWait(driver, 120).until(
                    expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page-1))))
                time.sleep(1)
                bs = BeautifulSoup(driver.page_source, 'html.parser')
                #get_ColumnsData(bs)
        
            print(i)
            i += 1

    # 狀況3 資料不超過10頁
    elif re.search('最末頁', str(page_info)) == None:
        #print('無最末頁按鈕')
        for row in table.find('td'):
            last_page = int([s for s in row.stripped_strings][-1])
            # print(last_page)

        # 從第一頁開始儲存資料
        bs = BeautifulSoup(driver.page_source, 'html.parser')
        #get_ColumnsData(bs)

        # 翻頁繼續儲存資料
        now_page = i = 1
        while i <= last_page-1:
            next_page = i + 1
            driver.find_element_by_link_text(
                str(next_page)).click()  # 正常換頁
            #time.sleep(10)
            element = WebDriverWait(driver, 60).until(
                expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page-1))))
            time.sleep(1)
            bs = BeautifulSoup(driver.page_source, 'html.parser')
            #get_ColumnsData(bs)

            i += 1

    #print(i)
    #print(district + ' ' + road + ' 爬取完成')

except UnexpectedAlertPresentException:
    error_info = sys.exc_info()
    error_msg = re.findall('\{.*\}', str(error_info))[0]
    print(district + ' ' + road + ' 爬取遇到錯誤' + ' 錯誤訊息：' + error_msg)


# %%

# %%
