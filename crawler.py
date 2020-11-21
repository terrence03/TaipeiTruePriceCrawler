# %%
from logging import fatal
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from tqdm import tqdm_notebook
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


def crawler(district, positioning_method, road, start_year, start_month, end_year, end_month, transactional_type='房地+房地車'):  # 此處有更動###
    try:
        # 先輸入篩選條件
        driver.get(url)  # 連接到台北地政雲不動產價格資訊\買賣實價查詢網頁
        # driver.refresh()

        # 選行政區
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_MasterGond')))
        driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_MasterGond').click()
        for option in driver.find_elements_by_tag_name('option'):
            if option.text == district:
                option.click()
                time.sleep(1)
                break

        # 選定位方式
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_locate')))
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
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear')))  # 此處有更動###
        select = Select(driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartYear'))
        select.select_by_value(str(start_year))
        time.sleep(0.5)

        # 選起始月
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth')))  # 此處有更動###
        select = Select(driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionStartMonth'))
        select.select_by_value(str(start_month).zfill(2))
        time.sleep(0.5)

        # 選截止年
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndYear')))  # 此處有更動###
        select = Select(driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndYear'))
        select.select_by_value(str(end_year))
        time.sleep(0.5)

        # 選截止月
        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndMonth')))  # 此處有更動###
        select = Select(driver.find_element_by_id(
            'ContentPlaceHolder1_ContentPlaceHolder1_TruePriceSearch_ddl_TransactionEndMonth'))
        select.select_by_value(str(end_month).zfill(2))
        time.sleep(0.5)

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

        element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located(
            (By.ID, 'ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')))
        bs = BeautifulSoup(driver.page_source, 'html.parser')

        table = bs.find(
            id='ContentPlaceHolder1_ContentPlaceHolder1_gvTruePrice_A_gv_TruePrice')
        page_info = table.find('td', {'colspan': '19'})  # 此處有更動###

        # 狀況1 資料只有一頁
        if page_info == None:
            # print('只有一頁')
            bs = BeautifulSoup(driver.page_source, 'html.parser')
            get_ColumnsData(bs)

        # 狀況2 資料有很多頁（超過10頁）
        elif re.search('最末頁', str(page_info)) != None:
            # print('有最末頁按鈕')
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
            # print('無最末頁按鈕')
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
                # time.sleep(10)
                element = WebDriverWait(driver, 60).until(
                    expected_conditions.element_to_be_clickable((By.LINK_TEXT, str(next_page-1))))
                time.sleep(1)
                bs = BeautifulSoup(driver.page_source, 'html.parser')
                get_ColumnsData(bs)

                i += 1

        print(f'{district} {road:{chr(12288)}<6} 爬取成功')

    # 可能遭遇的錯誤類型
    except StaleElementReferenceException:
        error_msg = '頁面缺少可選取目標'
        print(f'{district} {road:{chr(12288)}<6} 爬取失敗 錯誤訊息: {error_msg}')
        get_ErrorList(district, road, error_msg)

    except UnexpectedAlertPresentException:
        # error_info = sys.exc_info()
        # error_msg = re.findall('\{.*\}', str(error_info))[0]
        error_msg = '查無資料'
        print(f'{district} {road:{chr(12288)}<6} 爬取失敗 錯誤訊息: {error_msg}')
        get_ErrorList(district, road, error_msg)


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

# 保留錯誤資料
def get_ErrorList(error_district, error_road, error_msg):
    ErrorList_district.append(error_district)
    ErrorList_road.append(error_road)
    ErrorList_msg.append[error_msg]


# 匯入路段名稱資料
RoadData = pd.read_excel('路段.xlsx')


def get_RoadList(District):
    RoadList = RoadData[District].dropna().tolist()

    return RoadList


# 爬蟲模型
District_list = []            # 行政區
Adress_list = []              # 土地位置或建物門牌
Date_list = []                # 交易日期
TotalPrice_list = []          # 交易總價
UnitPrice_List = []           # 交易單價
Garage_list = []              # 單價是否含車位
BuildingArea_list = []        # 建物移轉面積
LandArea_list = []            # 土地移轉面積
BuildingType_list = []        # 建物型態
HouseAge_list = []            # 屋齡
Floor_list = []               # 樓層別/總樓層
TransactionalType_list = []   # 交易種類
Note_list = []                # 備註事項
TransactionRecord_list = []   # 歷次移轉

# 錯誤列表
ErrorList_district = []
ErrorList_road = []
ErrorList_msg = []

District_List = ['松山區', '大安區', '中正區', '萬華區', '大同區',
                 '中山區', '文山區', '南港區', '內湖區', '士林區', '北投區', '信義區']


# 選擇爬蟲方式
# 依行政區與時間範圍爬取該區全路段資料
def Clawler_by_District_and_Time(select_District, select_StartYear, select_StartMonth, select_EndYear, select_EndMonth):
    '''
    select_District: 行政區
    select_StartYear: 起始年
    select_StartMonth: 起始月
    select_EndYear: 截止年
    select_EndMonth: 截止月
    '''
    for i in tqdm_notebook(get_RoadList(select_District)):
        crawler(district=select_District, positioning_method='路段', road=i,
                start_year=select_StartYear, start_month=select_StartMonth, end_year=select_EndYear, end_month=select_EndMonth)
        time.sleep(1)
    print('爬取結束')


# 依時間範圍爬取臺北市全行政區資料
def Clawler_by_Time(select_StartYear, select_StartMonth, select_EndYear, select_EndMonth):
    '''
    select_StartYear: 起始年
    select_StartMonth: 起始月
    select_EndYear: 截止年
    select_EndMonth: 截止月
    '''
    for district in tqdm_notebook(District_List, desc='全區進度'):
        for r in tqdm_notebook(get_RoadList(district), desc='當區進度'):
            crawler(district=district, positioning_method='路段', road=r,
                    start_year=select_StartYear, start_month=select_StartMonth, end_year=select_EndYear, end_month=select_EndMonth)
            time.sleep(1)
    print('爬取結束')


# 測試用
# crawler(district='松山區', positioning_method='路段', road='東寧路', start_year=109, start_month=8, end_year=109, end_month=9)


# Clawler_by_District_and_Time('松山區', 109, 8, 109, 9)
Clawler_by_Time(106, 1, 108, 12)

# 將爬下來的資料存入字典
AllData = {'行政區': District_list, '土地位置或建物門牌': Adress_list,
           '交易日期': Date_list, '交易總價(萬元)': TotalPrice_list,
           '交易單價(萬元/坪)': UnitPrice_List, '單價是否含車位': Garage_list,
           '建物移轉面積(坪)': BuildingArea_list, '土地移轉面積(坪)': LandArea_list,
           '建物型態': BuildingType_list, '屋齡': HouseAge_list,
           '樓層別/總樓層': Floor_list, '交易種類': TransactionalType_list,
           '備註事項': Note_list, '歷次移轉(含過去移轉資料)': TransactionRecord_list
           }
AllData = pd.DataFrame(AllData)

# 保留資料(錯誤資料為爬取失敗資料，需要重新爬取)
ErrorData = {'行政區': ErrorList_district,
             '路段': ErrorList_road, '錯誤訊息': ErrorList_msg}
ErrorData = pd.DataFrame(ErrorData)


driver.quit()


# %%
# 輸出資料
AllData.to_excel('AllData.xlsx', encoding='cp950', index=False)
ErrorData.to_excel('ErrorData.xlsx', encoding='cp950', index=False)
