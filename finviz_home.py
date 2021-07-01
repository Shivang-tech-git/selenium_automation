from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import sys, os,shutil, time
import pandas as pd
from datetime import datetime

#---------- XPath ------ one element ----------------
def elementFinder(element, action="", text=""):
    htmlElement = None
    waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, element)))
    wait = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, element)))
    htmlElement = driver.find_element_by_xpath(element)
    if action == "click":
        driver.execute_script("arguments[0].click();", htmlElement)
    if action == "sendKeys":
        htmlElement.clear()
        htmlElement.send_keys(text)

def finviz():
    elementFinder('//a[text()="Screener"]','click')
    elementFinder('//table[@class="screener-view-table"]/descendant::a[text()="Custom"]','click')
    try:
        elementFinder('//a[@class="filter" and contains(text(),"Settings")]','click')
    except:
        pass
    should_restart = True
    while should_restart:
        try:
            waiting = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//td[@class="filters-border"]/descendant::tr')))
            list_of_columns = driver.find_elements_by_xpath('//td[@class="filters-border"]/descendant::tr')
            should_restart = False
            for row in list_of_columns:
                columns = row.find_elements_by_xpath('./td')
                for column in columns:
                    if column.find_element_by_xpath('./input[@type="checkbox"]').get_attribute('checked') != 'true':
                        column.find_element_by_xpath('./input[@type="checkbox"]').click()
                        should_restart = True
                        break
                if should_restart == True:
                    break
        except:
            try:
                elementFinder('//button[@class="modal-elite-ad_close"]','click')
                should_restart = True
            except:
                should_restart = False

def save_table():
    df = pd.DataFrame()
    elementFinder('//select[@id="signalSelect"]')
    columns_tags = driver.find_elements_by_xpath('(//div[@id="screener-content"]/descendant::table)[4]/descendant::tr[1]/td')
    columns = [col.get_attribute('innerText') for col in columns_tags]
    columns.insert(1, 'URL')
    columns.extend(['Signal'])
    signals = pd.read_csv('signal.csv')
    signals_list =  signals['Selected signals'].tolist()
    signals_list = [x for x in signals_list if str(x) != 'nan']
    for signal in signals_list:
        elementFinder('//select[@id="signalSelect"]/option[text()="{}"]'.format(signal))
        driver.find_element_by_xpath('//select[@id="signalSelect"]/option[text()="{}"]'.format(signal)).click()
        elementFinder('//select[@id="signalSelect"]/option[text()="{}"]'.format(signal))
        data = []
        table = driver.find_element_by_xpath('(//div[@id="screener-content"]/descendant::table)[4]')
        next_page = True
        while next_page:
            table = driver.find_element_by_xpath('(//div[@id="screener-content"]/descendant::table)[4]')
            body = table.find_elements_by_xpath('./descendant::tr')
            for row in body:
                cells = row.find_elements_by_xpath('./td')
                # dont add header row
                if cells[1].get_attribute('innerText') != 'Ticker':
                    one_row = []
                    one_row = [cell.get_attribute('innerText') for cell in cells]
                    one_row.insert(1,cells[1].find_element_by_xpath('./a').get_attribute('href'))
                    one_row.extend(['{}'.format(signal)])
                    data.append(one_row)
            try:
                elementFinder('//a[@class="screener_arrow"]','click')
            except:
                next_page = False
        current_signal = pd.DataFrame(data)
        df = df.append(current_signal)
        df.to_csv('finviz {}.csv'.format(datetime.today().strftime('%Y-%m-%d')), index=False)
    df.columns = columns
    df.to_csv('finviz {}.csv'.format(datetime.today().strftime('%Y-%m-%d')), index=False)

if __name__ == '__main__':
    chromedriver_path = os.path.join(os.path.abspath(os.getcwd()), "chromedriver.exe")
    chrome_options = Options()
    chrome_options.add_argument("load-extension=C:\\Users\\Shivang Gupta\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Extensions\\pkehgijcmpdhfbdbbnkijodmdjhbjlgp\\2021.6.8_0")
    driver = webdriver.Chrome(chromedriver_path,chrome_options=chrome_options)
    # driver = webdriver.Chrome(chromedriver_path)
    print('Running Chromedriver....')
    # driver.get('https://finviz.com/screener.ashx?v=152&c=0,1,2,3,4,5,6,7,8,9'
    #            ',10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29'
    #            ',30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49'
    #            ',50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69'
    #            ',70')
    driver.get("https://finviz.com/")
    driver.maximize_window()
    driver.switch_to.window(driver.window_handles[0])
    finviz()
    save_table()
    print('Done')
