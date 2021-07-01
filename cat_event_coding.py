from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
import time

def axaXlElementFinder(element, action="", text=""):
    htmlElement = None
    waiting = WebDriverWait(cat_driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
    wait = WebDriverWait(cat_driver, 30).until(EC.visibility_of_element_located((By.XPATH, element)))
    htmlElement = cat_driver.find_element_by_xpath(element)
    if action == "click":
        cat_driver.execute_script("arguments[0].click();", htmlElement)
    if action == "sendKeys":
        htmlElement.clear()
        htmlElement.send_keys(text)

def cat_event_coding(driver,cat_df):
    global cat_driver
    cat_driver = driver
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]','click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]','click')
    # --------------- For each row in cat_df, find BPR --------------------
    cat_df['Status'] = ''
    for i, row in cat_df.iterrows():
        try:
            if row['BPR'] != '':
                cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
                axaXlElementFinder('//input[@type = "text" and @name = "bpr"]','sendKeys', row['BPR'])
                axaXlElementFinder('//input[@type = "submit" and @value = "Search"]', 'click')
                assign_event_exist = False
            # ----------------------- check for the bpr text and cat code text ----------------------------------
                time.sleep(5)
                axaXlElementFinder('//td[./input[@id="csrfToken"]]')
                BPR_text = driver.find_element_by_xpath('//td[./input[@id="csrfToken"]]').get_attribute('innerText')
                if BPR_text.find('There are no items for event assignment') > 0:
                    cat_df.at[i, 'Status'] = 'There are no items for event assignment'
                    continue
                else:
                    assign_event_exist = True
            # ------------------- click on assign event and search market cat. code -------------------------------
                if assign_event_exist == True:
                    axaXlElementFinder('//table[@id="viewBean"]/descendant::input[@value="Assign event"]', 'click')
                    waiting = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]')))
                    axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
                    # -------------------------------- Empty all the textboxes ---------------------------------
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventCode"]','sendKeys')
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys')
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="startDate"]', 'sendKeys')
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="narrative"]', 'sendKeys')
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="days"]', 'sendKeys')
                    axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="catCode"]', 'sendKeys', row['Advised cat.'])
                    axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]', 'click')
                    # ------------------------------ cases > no results found or multiple cat codes or one cat code ----------------
                    waiting = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))
                    if driver.find_element_by_xpath('//div[@class="popupBody"]/descendant::td[@class="tabledata"]').get_attribute('innerText') == 'No results found':
                        cat_df.at[i, 'Status'] = 'No cat. code search results found'
                        axaXlElementFinder('(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]','click')
                        axaXlElementFinder('(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]','click')
                    else:
                        try:
                            html_element = driver.find_element_by_xpath('(//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]])[2]')
                            cat_df.at[i, 'Status'] = 'Multiple cat codes exist'
                            axaXlElementFinder('(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]','click')
                            axaXlElementFinder('(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]','click')
                        except:
                            axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value="   OK   "]', 'click')
                            time.sleep(5)
                            axaXlElementFinder(('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'), 'click')
                            cat_df.at[i, 'Status'] = 'Successfully Allocated'

        #  if any row throws unexpected error, then start frames again and start from the next row.
        except:
            cat_df.at[i, 'Status'] = 'Undefined Error'
            driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]','click')

    cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    cat_driver.quit()