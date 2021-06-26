from selenium import webdriver
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
    elementFinder('//select[@id="signalSelect"]/option[text()="Multiple Bottom"]')
    driver.find_element_by_xpath('//select[@id="signalSelect"]/option[text()="Multiple Bottom"]').click()
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
    table = driver.find_element_by_xpath('(//div[@id="screener-content"]/descendant::table)[4]')
    cell_list = [[col.get_attribute('innerText') for col in table.find_elements_by_xpath('./descendant::tr[1]/td')]]
    next_page = True
    while next_page:
        table = driver.find_element_by_xpath('(//div[@id="screener-content"]/descendant::table)[4]')
        body = table.find_elements_by_xpath('./descendant::tr')
        for row in body:
            cells = row.find_elements_by_xpath('./td')
            cell_list.append([cell.get_attribute('innerText') for cell in cells])
        try:
            elementFinder('//a[@class="screener_arrow"]','click')
        except:
            next_page = False

    df = pd.DataFrame(cell_list)
    df = df.drop(0)
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    df['signal'] = 'Multiple Bottom'
    df.to_csv('finviz {}.csv'.format(datetime.today().strftime('%Y-%m-%d')),index=False)


if __name__ == '__main__':
    chromedriver_path = os.path.join(os.path.abspath(os.getcwd()), "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
    print('Running Chromedriver....')
    driver.get("https://finviz.com/")
    driver.maximize_window()
    finviz()
    save_table()
    print('Done')
    # df = pd.read_csv('uk-500.csv')
    # df = pd.read_excel('filename', sheet_name='Sheet1')
    # df = df.loc[df['email'].str.endswith('@gmail.com')]
    # writer = pd.ExcelWriter('test2.xlsx')
    # df.to_excel(writer, index=False)
    # writer.save()