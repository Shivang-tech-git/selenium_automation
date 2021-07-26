from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
import time

property_services = ''

def axaXlElementFinder(element, action="", text=""):
    htmlElement = None
    waiting = WebDriverWait(main_driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
    wait = WebDriverWait(main_driver, 30).until(EC.visibility_of_element_located((By.XPATH, element)))
    htmlElement = main_driver.find_element_by_xpath(element)
    if action == "click":
        main_driver.execute_script("arguments[0].click();", htmlElement)
    if action == "sendKeys":
        htmlElement.clear()
        htmlElement.send_keys(text)

def event_refresh(event, driver, main_df, ue_df):
    global main_driver, main_event
    main_driver = driver
    main_event = event
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
    # --------------- For each row in pcs_df dataframe --------------------
    for i, row in ue_df.iterrows():
        coding_process(df=ue_df, i=i, row=row)
        # -------------- update pcs_df with new values from ue_pcs_df --------------
        main_df.loc[main_df['BPR'] == ue_df.at[i, 'BPR'], 'Status'] = ue_df.at[i, 'Status']
        main_df.to_csv('{0} Output {1}.CSV'.format(event, datetime.today().strftime('%d-%m-%Y')), index=False)
    main_driver.quit()

def event_coding(event, driver, main_df, blank_df):
    global main_driver, main_event
    main_driver = driver
    main_event = event
    # --------------- navigate to outstanding movements -------------------
    axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]','click')
    axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]','click')
    # --------------- For each row in main_df, find BPR --------------------
    for i, row in blank_df.iterrows():
        coding_process(df=blank_df, i=i, row=row)
        main_df.loc[main_df['BPR'] == blank_df.at[i, 'BPR'], 'Status'] = blank_df.at[i, 'Status']
        main_df.to_csv('{0} Output {1}.CSV'.format(main_event, datetime.today().strftime('%d-%m-%Y')), index=False)
    err_df = main_df.loc[main_df['Status'] == 'Undefined Error']
    # -------------- run again for undefined error -----------------
    if err_df.shape[0] > 0:
        for i, row in err_df.iterrows():
            # -------------- update main_df with new values from err_df --------------
            coding_process(df=err_df, i=i, row=row)
            main_df.loc[main_df['BPR'] == err_df.at[i, 'BPR'], 'Status'] = err_df.at[i, 'Status']
            main_df.to_csv('{0} Output {1}.CSV'.format(main_event, datetime.today().strftime('%d-%m-%Y')), index=False)
    main_driver.quit()

def coding_process(df, i, row):
    try:
        if row['BPR'] != '':
            try:
                BPR = int(float(row['BPR']))
            except:
                BPR = row['BPR']
            axaXlElementFinder('//input[@type = "text" and @name = "bpr"]', 'sendKeys', BPR)
            axaXlElementFinder('//input[@type = "submit" and @value = "Search"]', 'click')
            assign_event_exist = False
            search_code = []
            # ----------------------- check for the bpr text and cat code text ----------------------------------
            time.sleep(5)
            axaXlElementFinder('//td[./input[@id="csrfToken"]]')
            BPR_text = main_driver.find_element_by_xpath('//td[./input[@id="csrfToken"]]').get_attribute('innerText')
            if BPR_text.find('There are no items for event assignment') > 0:
                df.at[i, 'Status'] = 'There are no items for event assignment'
                return
            else:
                assign_event_exist = True
            # ------------------- click on assign event-------------------------------
            if assign_event_exist == True:
                axaXlElementFinder('//table[@id="viewBean"]/descendant::input[@value="Assign event"]', 'click')
                waiting = WebDriverWait(main_driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]')))
                # --------------- Move to the specific event coding process ---------------------------------------
                if main_event == 'CAT':
                    search_code = ['','','','','',row['Advised cat.']]
                elif main_event == 'PCS':
                    pcs_event_code_process()
                    if (property_services == '') and (df.at[i, 'Description'] == 'Mandatory Review. Event Code VARS'):
                        search_code = ['vars', '', '', '', '', '']
                    elif property_services == '':
                        axaXlElementFinder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                            'click')
                        df.at[i, 'Status'] = 'Property services code not available'
                        return
                    else:
                        search_code = ['', 'PCS {}'.format(property_services), '', '', '', '']
                elif main_event == 'VARS':
                    search_code = ['vars', '', '', '', '', '']
                # ------------------------------ search for specific event code results ----------------------------------------
                axaXlElementFinder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventCode"]', 'sendKeys', search_code[0])
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys', search_code[1])
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="startDate"]', 'sendKeys', search_code[2])
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="narrative"]', 'sendKeys', search_code[3])
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="days"]', 'sendKeys', search_code[4])
                axaXlElementFinder('//div[@class="popupBody"]/descendant::input[@name="catCode"]', 'sendKeys', search_code[5])
                # ------------------------------ cases > no results found or multiple cat codes or one cat code ----------------
                axaXlElementFinder(
                    '//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]', 'click')

                waiting = WebDriverWait(main_driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))
                if main_driver.find_element_by_xpath(
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
                        html_element = main_driver.find_element_by_xpath(
                            '(//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]])[2]')
                        df.at[i, 'Status'] = 'Multiple events exist'
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
            main_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
        except:
            main_driver.quit()
            main_driver.get('https://theframe.catlin.com/?windowId=1')
            axaXlElementFinder('//td[@class="mainmenu" and text()="Claims"]', 'click')
            axaXlElementFinder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')

def pcs_event_code_process():
    axaXlElementFinder(
        '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(),"Claim ID :")]]/a',
        'click')
    main_driver.switch_to.window(main_driver.window_handles[1])
    property_services = ''
    try:
        axaXlElementFinder('//a[text()="Original bureau message"]', 'click')
        axaXlElementFinder('//td[text()[contains(.,"Property services: ")]]')
        if main_driver.find_element_by_xpath(
                '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText') != '':
            property_services = main_driver.find_element_by_xpath(
                '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText')
    except:
        pass
    main_driver.close()
    main_driver.switch_to.window(main_driver.window_handles[0])
