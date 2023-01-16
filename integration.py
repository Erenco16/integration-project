from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import os
from dotenv import load_dotenv
import time
import pandas as pd
import xlwt

load_dotenv()

def driver_installer():
    driver = webdriver.Chrome(ChromeDriverManager().install())
    return driver

def extract_excel_file_from_evan(driver):
    login_url = "https://www.evan.com.tr/admin"
    driver.get(login_url)
    driver.find_element(By.CLASS_NAME, "inputName").send_keys(os.getenv("evan_username"))
    driver.find_element(By.CLASS_NAME, "inputPass").send_keys(os.getenv("evan_password"))
    driver.find_element(By.CLASS_NAME, "inputButton").click()
    driver.find_element(By.ID, "nav-10").click()
    integration_xpath = '''//*[@id="subnav-10"]/li[3]'''
    driver.find_element(By.XPATH, integration_xpath).click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="Grid_output_grid_GridData"]/div[3]/div[2]/div[8]''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="Grid_output_grid_Row6-Column5"]''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="outputForm"]/table/tbody/tr[11]/td[1]/table/tbody/tr[1]/td[1]''').click()
    driver.find_element(By.XPATH, '''//*[@id="outputForm"]/table/tbody/tr[12]/td/input''').click()
    time.sleep(10)


def extract_excel_file_from_dionaks(driver):
    login_url = "https://frankebursa.myideasoft.com/admin/auth/login?redirect_uri=%2Fadmin"
    driver.get(login_url)
    driver.find_element(By.CLASS_NAME, "inputName").send_keys(os.getenv("dionaks_username"))
    driver.find_element(By.CLASS_NAME, "inputPass").send_keys(os.getenv("dionaks_password"))
    driver.find_element(By.CLASS_NAME, "inputButton").click()
    driver.find_element(By.ID, "nav-10").click()
    integration_xpath = '''//*[@id="subnav-10"]/li[3]'''
    driver.find_element(By.XPATH, integration_xpath).click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="Grid_output_grid_Row15-Column5"]''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="outputForm"]/table/tbody/tr[11]/td[1]/table/tbody/tr[1]/td[1]''').click()
    driver.find_element(By.XPATH, '''//*[@id="outputForm"]/table/tbody/tr[12]/td/input''').click()
    time.sleep(10)


def excel_read(fpath, cname):
    excel_data = pd.read_excel(fpath)
    data = pd.DataFrame(excel_data,columns=[cname])
    stock_code_list=list()
    for code in data.values:
        if cname == "price1":
           stock_code_list.append(float(str(code[0]))*1.18)
        elif cname == "rebate":
            rebate = float(str(code[0]))
            if rebate > 99:
                rebate = rebate * 1.18
            else:
                pass
            stock_code_list.append(rebate)
        else:
            stock_code_list.append(str(code[0]))
    return stock_code_list

def mutual_list_returner(evan_fpath, dionaks_fpath, cname, price_column, rebate_column, rebate_type_column):
    evan_code_list = excel_read(evan_fpath, cname)
    dionaks_code_list = excel_read(dionaks_fpath, cname)
    evan_price_list = excel_read(evan_fpath, price_column)
    evan_rebate_list = excel_read(evan_fpath, rebate_column)
    evan_rebate_type_list = excel_read(evan_fpath, rebate_type_column)
    evan_list_general = []
    for i in range(len(evan_code_list)):
        product_tuple = (evan_code_list[i], evan_price_list[i], evan_rebate_list[i], evan_rebate_type_list[i])
        evan_list_general.append(product_tuple)

    mutual_list = []
    for p in dionaks_code_list:
        if p in evan_code_list:
            index = int(evan_code_list.index(p))
            mutual_tuple = (p, evan_list_general[index][1], evan_list_general[index][2], evan_list_general[index][3])
            mutual_list.append(mutual_tuple)
    return mutual_list

def create_dataframe_and_extract_to_xls(headers, data_list):
    data_for_excel = []
    for i in data_list:
        row_list = [i[0], i[1], i[2], i[3]]
        data_for_excel.append(row_list)

    df = pd.DataFrame(data_for_excel, columns=headers)
    fpath = "/Users/godfather/PycharmProjects/testProject/dionaks_new_franke_prices.xls"
    if os.path.exists(fpath) == True:
        os.remove(fpath)
    else:
        pass
    df.to_excel("dionaks_new_franke_prices.xls", sheet_name="products_info")

def submit_file_to_dionaks(driver):
    driver.find_element(By.ID, "nav-10").click()
    driver.find_element(By.XPATH, '''//*[@id="subnav-10"]/li[1]/a''').click()
    driver.find_element(By.ID, "integrationNavigator").click()
    time.sleep(1)
    select = Select(driver.find_element(By.XPATH, '''//*[@id="integrationNavigator"]'''))
    select.select_by_visible_text("Excel")
    time.sleep(3)
    file = driver.find_element(By.XPATH, '''//*[@id="sourceFileUploadForm"]/div/div[3]/div[1]/div/input''')
    file.send_keys(r"/Users/godfather/Desktop/integrationProject/dionaks_new_franke_prices.xls")
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="sourceFileUploadForm"]/div/div[3]/div[2]/input''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="contentWrapper"]/div/div/div[2]/div[1]/input''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="contentWrapper"]/div/div/div[2]/div[6]/div[2]/div/div[2]/div/div[1]/ul/li[3]/a''').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '''//*[@id="taxstatus1"]''').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '''//*[@id="taxstatus1"]/option[2]''').click()
    driver.find_element(By.XPATH, '''//*[@id="contentWrapper"]/div/div/div[2]/div[6]/div[2]/div/div[2]/div/div[2]/div[3]/form/div/div[6]/input''').click()
    time.sleep(10)
    driver.close()