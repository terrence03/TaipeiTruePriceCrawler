# %%
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
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


def crawler(district, positioning_method, road, start_year, start_month, end_year, end_month, transactional_type='房地+房地車'):  ###此處有更動###
    '''
    輸入選擇條件
    '''
    # 先輸入篩選條件
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
        (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_txb_GondRoad')))  ###此處有更動###
    driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_txb_GondRoad').click()  ###此處有更動###
    for option in driver.find_elements_by_tag_name('option'):
        if option.text == road:
            option.click()
            time.sleep(0.5)
            break
                
    # 選起始年
    select = Select(driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear'))
    select.select_by_value(str(start_year))
    time.sleep(0.5)
    
    # 選起始月
    select = Select(driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth'))
    select.select_by_value(str(start_month).zfill(2))
    time.sleep(0.5)

    # 選截止年
    select = Select(driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndYear'))
    select.select_by_value(str(end_year))
    time.sleep(0.5)
    
    # 選截止月
    select = Select(driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndMonth'))
    select.select_by_value(str(end_month).zfill(2))
    time.sleep(0.5)

    '''
    # 選起始年月和截止年月產生問題
    # 由於原來的選擇方式在選起始年月時是先點開下拉選單後選擇選項
    # 在選截止年月時的選擇會覆蓋回起始年月，因為選項種類一致
    # 這是因為option選項沒有經過定位造成的
    # 所以改用上面的寫法
     
    # 選起始年    
    driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear').click()
    for option in driver.find_elements_by_tag_name('option'):
        if option.text == str(start_year):  ###此處有更動###
            option.click()
            time.sleep(0.5)
            break
    
    # 選起始月
    driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth').click()
    for option in driver.find_elements_by_tag_name('option'):
        if option.text == str(start_month).zfill(2):  ###此處有更動###
            option.click()
            time.sleep(0.5)
            break
    
    # 選截止年
    driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndYear').click()
    for option in driver.find_elements_by_tag_name('option'):
        if option.text == str(end_year):
            option.click()
            time.sleep(1)
            break
 
    # 選截止月
    driver.find_element_by_id(
        'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndMonth').click()
    for option in driver.find_elements_by_tag_name('option'):
        if option.text == str(end_month).zfill(2):
            option.click()
            time.sleep(1)
            break
    '''        
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

        table = bs.find(id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
        page_info = table.find('td', {'colspan':'19'})  ###此處有更動###

        # 狀況1 資料只有一頁
        if page_info == None:
            #print('只有一頁')
            bs = BeautifulSoup(driver.page_source, 'html.parser')
            get_ColumnsData(bs)

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
                # print(last_page)

            # 回到第一頁
            driver.find_element_by_link_text('第一頁').click()
            element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
                (By.LINK_TEXT, '最末頁')))
            time.sleep(1)
            
            # 從第一頁開始儲存資料
            bs = BeautifulSoup(driver.page_source, 'html.parser')
            get_ColumnsData(bs)
            
            # 翻頁繼續儲存資料
            now_page = i = 1
            while i <= last_page-1:
                next_page = i + 1
                if next_page == 11:  # 若下一頁為11，點擊'...'的按鈕
                    driver.find_element_by_link_text('...').click()
                    element = WebDriverWait(driver, 60).until(expected_conditions.element_to_be_clickable(
                        (By.LINK_TEXT, '第一頁')))
                    time.sleep(1)
                    bs = BeautifulSoup(driver.page_source, 'html.parser')
                    get_ColumnsData(bs)

                # 若下一頁為[21,31,41,51,61,71,81,91,101,]點擊td[13]的超連接
                elif next_page in [21, 31, 41, 51, 61, 71, 81, 91, 101, 111, 121, 131, 141, 151]:
                    driver.find_element_by_xpath(
                        '//*[@id = "ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice"]/tbody/tr[1]/td/table/tbody/tr/td[13]/a').click()
                    element = WebDriverWait(driver, 20).until_not(
                        expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page))))
                    time.sleep(1)    
                    bs = BeautifulSoup(driver.page_source, 'html.parser')
                    get_ColumnsData(bs)

                else:
                    driver.find_element_by_link_text(
                        str(next_page)).click()  # 正常換頁
                    element = WebDriverWait(driver, 120).until(
                        expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page-1))))
                    time.sleep(1)
                    bs = BeautifulSoup(driver.page_source, 'html.parser')
                    get_ColumnsData(bs)

                i += 1

        # 狀況3 資料不超過10頁
        elif re.search('最末頁', str(page_info)) == None:
            #print('無最末頁按鈕')
            for row in table.find('td'):
                last_page = int([s for s in row.stripped_strings][-1])
                # print(last_page)
            
            # 從第一頁開始儲存資料
            bs = BeautifulSoup(driver.page_source, 'html.parser')
            get_ColumnsData(bs)

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
                get_ColumnsData(bs)

                i += 1    

        print('%s %6s 資料爬取完成' % (district,  road))

    except UnexpectedAlertPresentException:
        error_info = sys.exc_info()
        error_msg = re.findall('\{.*\}', str(error_info))[0]
        print('%s %6s 爬取遇到錯誤/錯誤訊息: %s' % (district, road, error_msg))
        # print(district + ' ' + road + ' 爬取遇到錯誤' + ' 錯誤訊息：' + error_msg)
    
   

def get_ColumnsData(bs):
    # 讀取表格
    string = ''
    for sibling in bs.find('table', {'id': 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice'}).tr.next_siblings:
        string = string + str(sibling)

    # 解析表格
    soup = BeautifulSoup(string, 'html.parser')
    content = soup.find_all('tr', {'class': 'gridTable'})

    counter = 1
    while len(content) > counter:
        # 要記錄的資料
        District = content[counter].find_all(
            'td')[1].get_text()
        Adress = content[counter].find_all(
            'td')[2].get_text()
        Date = content[counter].find_all(
            'td')[3].get_text()
        TotalPrice = content[counter].find_all(
            'td')[4].get_text()
        UnitPrice = content[counter].find_all(
            'td')[5].get_text()
        Garage = content[counter].find_all(
            'td')[6].get_text()
        BuildingArea = content[counter].find_all(
            'td')[7].get_text()
        LandArea = content[counter].find_all(
            'td')[8].get_text()
        BuildingType = content[counter].find_all(
            'td')[9].get_text()
        HouseAge = content[counter].find_all(
            'td')[10].get_text()
        Floor = content[counter].find_all(
            'td')[11].get_text()
        TransactionalType = content[counter].find_all(
            'td')[12].get_text()
        Note = content[counter].find_all(
            'td')[13].get_text()
        TransactionRecord = content[counter].find_all(
            'td')[14].get_text()

        # 將資料存入對應列表中
        District_list.append(District)
        Adress_list.append(Adress)
        Date_list.append(Date)
        TotalPrice_list.append(TotalPrice)
        UnitPrice_List.append(UnitPrice)
        Garage_list.append(Garage)
        BuildingArea_list.append(BuildingArea)
        LandArea_list.append(LandArea)
        BuildingType_list.append(BuildingType)
        HouseAge_list.append(HouseAge)
        Floor_list.append(Floor)
        TransactionalType_list.append(TransactionalType)
        Note_list.append(Note)
        TransactionRecord_list.append(TransactionRecord)

        counter += 1


def get_RoadList(District):
    RoadList = RoadData[District].dropna().tolist()

    return RoadList


# 匯入路段名稱資料
RoadData = pd.read_excel('路段.xlsx')

# 爬蟲模型
District_list = []  # 行政區
Adress_list = []  # 土地位置或建物門牌
Date_list = []  # 交易日期
TotalPrice_list = []
UnitPrice_List = []
Garage_list = []
BuildingArea_list = []
LandArea_list = []
BuildingType_list = []
HouseAge_list = []
Floor_list = []
TransactionalType_list = []
Note_list = []
TransactionRecord_list = []

column = ['行政區', '土地位置或建物門牌', '交易日期', '交易總價(萬元)', '交易單價(萬元/坪)', '單價是否含車位', '建物移轉面積(坪)',
          '土地移轉面積(坪)', '建物型態', '屋齡', '樓層別/總樓層', '交易種類', '備註事項', '歷次移轉(含過去移轉資料)']

# 設定要搜尋的行政區
District_List = ['松山區','大安區','中正區','萬華區','大同區','中山區','文山區','南港區','內湖區','士林區','北投區','信義區']
Search_District = '信義區'

# 開始爬蟲

# 測試用
# crawler(district=Search_District, positioning_method='路段', road='和平東路三段', start_year=109, start_month=7, end_year=109, end_month=9)


# 此程式是抓單一路段的資料，可以透過迴圈爬取其他路段的資料
'''
for i in tqdm(get_RoadList(Search_District)):
    crawler(district=Search_District, positioning_method='路段', road=i,
            start_year=108, start_month=8, end_year=109, end_month=9)
    time.sleep(1)
'''

# 依時間範圍爬取臺北市全區資料
def Clawler_by_Time(sy, sm, ey, em):
    '''
    sy: 起始年
    sm: 起始月
    ey: 截止年
    em: 截止月
    '''
    for district in tqdm(District_List):
        for r in (get_RoadList(district)):
            crawler(district=district, positioning_method='路段', road=r,
                    start_year=sy, start_month=sm, end_year=ey, end_month=em)
            time.sleep(1)
    print('全部資料爬取完成')

Clawler_by_Time(109, 8, 109, 9)

# 將爬下來的資料存入字典
ColumnsData = {'行政區': District_list, '土地位置或建物門牌': Adress_list,
               '交易日期': Date_list, '交易總價(萬元)': TotalPrice_list,
               '交易單價(萬元/坪)': UnitPrice_List, '單價是否含車位': Garage_list,
               '建物移轉面積(坪)': BuildingArea_list, '土地移轉面積(坪)': LandArea_list,
               '建物型態': BuildingType_list, '屋齡': HouseAge_list,
               '樓層別/總樓層': Floor_list, '交易種類': TransactionalType_list,
               '備註事項': Note_list, '歷次移轉(含過去移轉資料)': TransactionRecord_list
               }


AllData = pd.DataFrame(ColumnsData)
driver.quit()


# %%
# 輸出資料
AllData.to_excel('data.xlsx')

# %%
print(str(7).zfill(2))
# %%
