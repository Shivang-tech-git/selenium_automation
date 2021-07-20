from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
import time
from tkinter import messagebox

def axaXlElementFinder(element, action="", text=""):
    htmlElement = None
    waiting = WebDriverWait(pcs_driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
    wait = WebDriverWait(pcs_driver, 30).until(EC.visibility_of_element_located((By.XPATH, element)))
    htmlElement = pcs_driver.find_element_by_xpath(element)
    if action == "click":
        pcs_driver.execute_script("arguments[0].click();", htmlElement)
    if action == "sendKeys":
        htmlElement.clear()
        htmlElement.send_keys(text)


def pcs_event_refresh(driver, pcs_df, ue_pcs_df):
    global pcs_driver
    pcs_driver = driver
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
    # --------------- For each row in pcs_df dataframe --------------------
    for i, row in ue_pcs_df.iterrows():
        pcs_coding_process(df=ue_pcs_df, i=i, row=row)
        # -------------- update pcs_df with new values from ue_pcs_df --------------
        pcs_df.loc[pcs_df['BPR'] == ue_pcs_df.at[i, 'BPR'], 'Status'] = ue_pcs_df.at[i,'Status']
        pcs_df.to_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    pcs_driver.quit()

def pcs_event_coding(driver, pcs_df, blank_pcs_df):
    global pcs_driver
    pcs_driver = driver
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
    # --------------- For each row in dataframe --------------------
    for i, row in blank_pcs_df.iterrows():
        pcs_coding_process(df=blank_pcs_df, i=i,row=row)
        pcs_df.loc[pcs_df['BPR'] == blank_pcs_df.at[i, 'BPR'], 'Status'] = blank_pcs_df.at[i, 'Status']
        pcs_df.to_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    err_df = pcs_df.loc[pcs_df['Status'] == 'Undefined Error']
    # -------------- run again for undefined error -----------------
    if err_df.shape[0] > 0:
        for i, row in err_df.iterrows():
            pcs_coding_process(df=err_df, i=i, row=row)
            pcs_df.loc[pcs_df['BPR'] == err_df.at[i, 'BPR'], 'Status'] = err_df.at[i, 'Status']
            pcs_df.to_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    pcs_driver.quit()

def pcs_coding_process(df, i, row):
    try:
        if row['BPR'] != '':
            axaXlElementFinder('//input[@type = "text" and @name = "bpr"]', 'sendKeys', row['BPR'])
            axaXlElementFinder('//input[@type = "submit" and @value = "Search"]', 'click')
            assign_event_exist = False
            # ----------------------- check for the bpr text and cat code text ----------------------------------
            time.sleep(5)
            axaXlElementFinder('//td[./input[@id="csrfToken"]]')
            BPR_text = pcs_driver.find_element_by_xpath('//td[./input[@id="csrfToken"]]').get_attribute('innerText')
            if BPR_text.find('There are no items for event assignment') > 0:
                df.at[i, 'Status'] = 'There are no items for event assignment'
                return
            else:
                assign_event_exist = True
            # ------------------- click on assign event and search market cat. code -------------------------------
            if assign_event_exist == True:
                axaXlElementFinder('//table[@id="viewBean"]/descendant::input[@value="Assign event"]', 'click')
                waiting = WebDriverWait(pcs_driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]')))
                # -------------------------------- check property services code ------------------------------

                axaXlElementFinder(
                    '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(),"Claim ID :")]]/a',
                    'click')
                pcs_driver.switch_to.window(pcs_driver.window_handles[1])
                property_services = ''
                try:
                    axaXlElementFinder('//a[text()="Original bureau message"]', 'click')
                    axaXlElementFinder('//td[text()[contains(.,"Property services: ")]]')
                    if pcs_driver.find_element_by_xpath(
                            '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText') != '':
                        property_services = pcs_driver.find_element_by_xpath(
                            '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText')
                except:
                    pass
                pcs_driver.close()
                pcs_driver.switch_to.window(pcs_driver.window_handles[0])
                if property_services == '':
                    df.at[i, 'Status'] = 'Property services code not available'
                    axaXlElementFinder(
                        '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                        'click')
                    return
                axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
                # -------------------------------- Empty all the textboxes ---------------------------------
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventCode"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys',
                                   'PCS {}'.format(property_services))
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="startDate"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="narrative"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="days"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="catCode"]', 'sendKeys')
                axaXlElementFinder(
                    '//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]',
                    'click')
                # ------------------------------ cases > no results found or multiple cat codes or one cat code ----------------
                waiting = WebDriverWait(pcs_driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))
                if pcs_driver.find_element_by_xpath(
                        '//div[@class="popupBody"]/descendant::td[@class="tabledata"]').get_attribute(
                    'innerText') == 'No results found':
                    df.at[i, 'Status'] = 'No results found'
                    axaXlElementFinder(
                        '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                        'click')
                    axaXlElementFinder(
                        '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                        'click')
                else:
                    try:
                        html_element = pcs_driver.find_element_by_xpath(
                            '(//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]])[2]')
                        df.at[i, 'Status'] = 'Multiple events exist'
                        axaXlElementFinder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                            'click')
                        axaXlElementFinder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                            'click')
                    except:
                        axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value="   OK   "]',
                                           'click')
                        time.sleep(5)
                        axaXlElementFinder(
                            ('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'),
                            'click')
                        df.at[i, 'Status'] = 'Successfully Allocated'

    #  if any row throws unexpected error, then start frames again and start from the next row.
    except Exception as e:
        df.at[i, 'Status'] = 'Undefined Error'
        try:
            pcs_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]',
                               'click')
        except:
            pcs_driver.quit()
            pcs_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]',
                               'click')


