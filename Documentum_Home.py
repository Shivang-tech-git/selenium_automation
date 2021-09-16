from msedge.selenium_tools import Edge, EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from tkinter import messagebox
import os, getpass, time
import pandas, xlwings

#---------- XPath ------ one element ----------------
def axaXlElementFinder(element, action="", text=""):
    htmlElement = None
    waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, element)))
    wait = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, element)))
    htmlElement = driver.find_element_by_xpath(element)
    if action == "click":
        driver.execute_script("arguments[0].click();", htmlElement)
    if action == "sendKeys":
        htmlElement.clear()
        htmlElement.send_keys(text)

def open_edge():
    global driver
    edgedriver_path = os.path.join(os.path.abspath(os.getcwd()), "msedgedriver.exe")
    Edge_Options = EdgeOptions()
    Edge_Options.add_experimental_option("excludeSwitches", ["enable-automation"])
    Edge_Options.add_experimental_option('useAutomationExtension', False)
    Edge_Options.use_chromium = True
    Edge_Options.add_argument('no-sandbox')
    Edge_Options.add_argument("user-data-dir=C:/Users/" + getpass.getuser() + "/AppData/Local/Microsoft/Edge/User Data")
    # Edge_Options.add_argument("headless")
    # Edge_Options.add_argument("disable-gpu")
    try:
        driver = Edge(EdgeChromiumDriverManager().install(),options=Edge_Options)
    except:
        messagebox.showinfo('Documentum tool',
                            'Please close Edge and then run the tool.')
    driver.set_page_load_timeout(60)
    driver.get('https://xlcdoc.r02.xlgs.local/XLCDoc/?ws=cdagclaimsprivate')

def create_folder(cur_folder, cur_sub_folder_name):
    action = ActionChains(driver)
    action.context_click(cur_folder).perform()
    axaXlElementFinder('//table[@class="menuTable"]/descendant::div[contains(text(),"New Folder")]')
    driver.find_element_by_xpath('//table[@class="menuTable"]/descendant::div[contains(text(),"New Folder")]').click()
    axaXlElementFinder('(//div[@class="windowBackground"])[2]')
    driver.find_element_by_xpath('//div[@class="windowBackground"]/descendant::input[@type="TEXT"]').send_keys(cur_sub_folder_name)
    driver.find_element_by_xpath('//div[@class="windowBackground"]/descendant::div[text()="CREATE"]').click()
    # driver.find_element_by_xpath('//div[@class="windowBackground"]/descendant::div[text()="CANCEL"]').click()
    wait = WebDriverWait(driver, 10).until(
        EC.invisibility_of_element_located((By.XPATH, '//td[contains(text(),"Please wait...")]')))

def scroll_to_folder(folder_name):
    while True:
        folder_found = False
        last_cell = driver.find_elements_by_xpath('//table[contains(@class,"treeCell")]/descendant::td[contains(@id,"valueCell")]').__len__()
        for row in range(2,last_cell):
            cell_name = driver.find_element_by_xpath(f'(//table[contains(@class,"treeCell")]/descendant::td[contains(@id,"valueCell")])[{row}]').get_attribute('innerText')
            if cell_name == folder_name:
                folder_found = True
                return
        if folder_found == False:
            driver.execute_script("arguments[0].scrollIntoView(true);", driver.find_element_by_xpath(f'(//table[contains(@class,"treeCell")]/descendant::td[contains(@id,"valueCell")])[{last_cell}]'))
            time.sleep(3)

def documentum_home():
    axaXlElementFinder('//table[@class="listTable"]')
    wait = WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, '//td[contains(text(),"Please wait...")]')))
    driver.find_element_by_xpath('//table[@class="listTable"]/descendant::tr[./td[contains(@class,"treeCell") and text()="CDAG Private Documents"]]').click()
    wait = WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, '//td[contains(text(),"Please wait...")]')))
    scroll_to_folder(folder_name)
    folder = driver.find_element_by_xpath(f'//table[@class="listTable"]/descendant::tr[./td[contains(@class,"treeCell") and text()="{folder_name}"]]')
    driver.execute_script("arguments[0].scrollIntoView(true);", folder)
    time.sleep(5)
    folder = driver.find_element_by_xpath(f'//table[@class="listTable"]/descendant::tr[./td[contains(@class,"treeCell") and text()="{folder_name}"]]')
    create_folder(cur_folder=folder,cur_sub_folder_name=sub_folder_name)
    time.sleep(5)
    for row in range(7,last_row + 1):
        if tool_sht.range("E{0}".format(row)).value != 'None':
            try:
                folder = driver.find_element_by_xpath(f'//table[contains(@class,"treeCell")]/descendant::td[text()="{sub_folder_name}"]')
            except:
                scroll_to_folder(sub_folder_name)
                folder = driver.find_element_by_xpath(f'//table[contains(@class,"treeCell")]/descendant::td[text()="{sub_folder_name}"]')
            driver.execute_script("arguments[0].scrollIntoView(true);", folder)
            time.sleep(5)
            try:
                folder = driver.find_element_by_xpath(f'//table[contains(@class,"treeCell")]/descendant::td[text()="{sub_folder_name}"]')
            except:
                axaXlElementFinder(f'//table[contains(@class,"treeCell")]/descendant::td[text()="{sub_folder_name}"]')
                folder = driver.find_element_by_xpath(f'//table[contains(@class,"treeCell")]/descendant::td[text()="{sub_folder_name}"]')
            sub_sub_folder_name = str(tool_sht.range("E{0}".format(row)).value)
            create_folder(cur_folder=folder,cur_sub_folder_name=sub_sub_folder_name)

    messagebox.showinfo('Documentum tool',
                        f'Created new folders for {sub_folder_name}')

if __name__ == '__main__':
    tool_sht = xlwings.Book('Documentum Tool.xlsm').sheets['Tool']
    last_row = tool_sht.range('E' + str(tool_sht.cells.last_cell.row)).end('up').row
    folder_name = str(tool_sht.range("E5").value)
    sub_folder_name = str(tool_sht.range("E6").value)
    open_edge()
    try:
        documentum_home()
    except:
        messagebox.showerror('Documentum tool',
                            "All folder's were not created because of error.")
