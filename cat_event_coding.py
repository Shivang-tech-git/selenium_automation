from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
from tkinter import messagebox
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

def cat_event_refresh(driver, cat_df, ue_cat_df):
    global cat_driver
    cat_driver = driver
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
    # --------------- For each row in pcs_df dataframe --------------------
    for i, row in ue_cat_df.iterrows():
        cat_coding_process(df=ue_cat_df, i=i, row=row)
        # -------------- update pcs_df with new values from ue_pcs_df --------------
        cat_df.loc[cat_df['BPR'] == ue_cat_df.at[i, 'BPR'], 'Status'] = ue_cat_df.at[i, 'Status']
        cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    cat_driver.quit()

def cat_event_coding(driver, cat_df, blank_cat_df):
    global cat_driver
    cat_driver = driver
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]','click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]','click')
    # --------------- For each row in cat_df, find BPR --------------------
    for i, row in blank_cat_df.iterrows():
        cat_coding_process(df=blank_cat_df,i=i,row=row)
        cat_df.loc[cat_df['BPR'] == blank_cat_df.at[i, 'BPR'], 'Status'] = blank_cat_df.at[i, 'Status']
        cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    err_df = cat_df.loc[cat_df['Status'] == 'Undefined Error']
    # -------------- run again for undefined error -----------------
    if err_df.shape[0] > 0:
        for i, row in err_df.iterrows():
            # -------------- update cat_df with new values from err_df --------------
            cat_coding_process(df=err_df, i=i, row=row)
            cat_df.loc[cat_df['BPR'] == err_df.at[i, 'BPR'], 'Status'] = err_df.at[i, 'Status']
            cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
    cat_driver.quit()

def cat_coding_process(df, i, row):
    try:
        if row['BPR'] != '':
            axaXlElementFinder('//input[@type = "text" and @name = "bpr"]', 'sendKeys', row['BPR'])
            axaXlElementFinder('//input[@type = "submit" and @value = "Search"]', 'click')
            assign_event_exist = False
            # ----------------------- check for the bpr text and cat code text ----------------------------------
            time.sleep(5)
            axaXlElementFinder('//td[./input[@id="csrfToken"]]')
            BPR_text = cat_driver.find_element_by_xpath('//td[./input[@id="csrfToken"]]').get_attribute('innerText')
            if BPR_text.find('There are no items for event assignment') > 0:
                df.at[i, 'Status'] = 'There are no items for event assignment'
                return
            else:
                assign_event_exist = True
            # ------------------- click on assign event and search market cat. code -------------------------------
            if assign_event_exist == True:
                axaXlElementFinder('//table[@id="viewBean"]/descendant::input[@value="Assign event"]', 'click')
                waiting = WebDriverWait(cat_driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]')))
                axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
                # -------------------------------- Empty all the textboxes ---------------------------------
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventCode"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="startDate"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="narrative"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="days"]', 'sendKeys')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="catCode"]', 'sendKeys',
                                   row['Advised cat.'])
                axaXlElementFinder(
                    '//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]', 'click')
                # ------------------------------ cases > no results found or multiple cat codes or one cat code ----------------
                waiting = WebDriverWait(cat_driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))
                if cat_driver.find_element_by_xpath(
                        '//div[@class="popupBody"]/descendant::td[@class="tabledata"]').get_attribute(
                        'innerText') == 'No results found':
                    df.at[i, 'Status'] = 'No cat. code search results found'
                    axaXlElementFinder(
                        '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                        'click')
                    axaXlElementFinder(
                        '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                        'click')
                else:
                    try:
                        html_element = cat_driver.find_element_by_xpath(
                            '(//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]])[2]')
                        df.at[i, 'Status'] = 'Multiple cat codes exist'
                        axaXlElementFinder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                            'click')
                        axaXlElementFinder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                            'click')
                    except:
                        axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value="   OK   "]', 'click')
                        time.sleep(5)
                        axaXlElementFinder(
                            ('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'),
                            'click')
                        df.at[i, 'Status'] = 'Successfully Allocated'

    #  if any row throws unexpected error, then start frames again and start from the next row.
    except:
        df.at[i, 'Status'] = 'Undefined Error'
        try:
            cat_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
        except:
            cat_driver.quit()
            cat_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
