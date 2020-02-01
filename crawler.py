# %%
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import pandas as pd
import sqlite3

url = 'https://cloud.land.gov.taipei/ImmPrice/TruePriceA.aspx'  # 台北地政雲網站
# 瀏覽器位置
# Options.binary_location = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
# webdriver位置
webdriver_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe'

options = Options()
driver = webdriver.Chrome(executable_path=webdriver_path, options=options)
# driver.maximize_window()


def get_data(district, positioning_method, road, transactional_type='房地(土地+建物)'):
    '''
    輸入選擇條件
    '''
    # 先輸入篩選條件
    try:
        driver.set_page_load_timeout(20)
        driver.get(url)  # 連接到台北地政雲不動產價格資訊\買賣實價查詢網頁

        # 選行政區
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_MasterGond').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == district:
                option.click()
                time.sleep(2)
                break

        # 選定位方式
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_locate').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == positioning_method:
                option.click()
                time.sleep(2)
                break

        # 選路段
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_GondRoad').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == road:
                option.click()
                time.sleep(2)
                break

        # 選起始年
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == '101':
                option.click()
                time.sleep(2)
                break

        # 選起始月
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == '08':
                option.click()
                time.sleep(2)
                break

        # 選結束年 公開資料似乎還沒更新，最新只到10811，原本預設就是最新時間10811，所以不需設定
        '''
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndYear').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == '108':
                option.click()
                time.sleep(3)
                break

        # 選結束月
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndMonth').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == '11':
                option.click()
                time.sleep(3)
                break
        '''
        # 選交易類型
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionType').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == transactional_type:
                option.click()
                time.sleep(2)
                break

        # 點選查詢
        element = driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_btn_Search').click()
        time.sleep(10)

        # 等待第一頁資料出來
        element = WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))

        # 解析網頁內容
        # soup = BeautifulSoup(driver.page_source, 'lxml')
        # table = soup.find(
            # id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')

        # 先翻到最末頁確認總頁數
        driver.find_element_by_link_text('最末頁').click()
        time.sleep(10)
        element = WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
        soup = BeautifulSoup(driver.page_source, 'lxml')
        table = soup.find(
            id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
        for row in table.find('td'):
            last_page = int([s for s in row.stripped_strings][-1])
            # print(last_page)

        # 回到第一頁
        driver.find_element_by_link_text('第一頁').click()
        time.sleep(10)

        # 從第一頁開始儲存資料
        element = WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
        soup = BeautifulSoup(driver.page_source, 'lxml')
        table = soup.find(id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
        result = []
        for row in table.find_all('tr', {'class':'gridTable'}):
            result.append([s for s in row.stripped_strings])
                                
        # 翻頁繼續儲存資料
        now_page = i = 1
        while i <= last_page-1:
            next_page = i + 1
            if next_page == 11: # 若下一頁為11，點擊'...'的按鈕
                driver.find_element_by_link_text('...').click()
                time.sleep(30)
                soup = BeautifulSoup(driver.page_source, 'lxml')
                table = soup.find(
                    id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
                for row in table.find_all('tr', {'class': 'gridTable'}):
                    result.append([s for s in row.stripped_strings])
            elif next_page in [21, 31, 41, 51, 61, 71, 81, 91, 101]: # 若下一頁為[21,31,41,51,61,71,81,91,101,]點擊td[13]的超連接
                driver.find_element_by_xpath(
                    '//*[@id = "ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice"]/tbody/tr[1]/td/table/tbody/tr/td[13]/a').click()
                time.sleep(30)
                soup = BeautifulSoup(driver.page_source, 'lxml')
                table = soup.find(
                    id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
                for row in table.find_all('tr', {'class': 'gridTable'}):
                    result.append([s for s in row.stripped_strings])
            else:
                driver.find_element_by_link_text(str(next_page)).click() # 正常換頁
                time.sleep(30)
                soup = BeautifulSoup(driver.page_source, 'lxml')
                table = soup.find(
                id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
                for row in table.find_all('tr', {'class': 'gridTable'}):
                    result.append([s for s in row.stripped_strings])

            i += 1

        # 資料儲存
        df = pd.DataFrame(result)
    finally:
        print('爬取資料完成')

    return df

data = get_data(district='松山區', positioning_method='路段', road='八德路二段')
# 此程式是抓單一路段的資料，可以透過迴圈爬取其他路段的資料

driver.close()

# %%
# 輸出資料
data.to_excel('data.xlsx')
