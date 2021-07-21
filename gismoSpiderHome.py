# -*- coding: utf-8 -*-
"""
Created on Thu Oct 29 01:52:53 2020
@author: X134391
"""
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from openpyxlGismo import columns, cov_list, cont_list, exclusions, details, cont_str
from datetime import datetime
import allconstants as AC
import time
import sys, os

print('Loading all procedures....')
class numformat():
    def __init__(self, num):
        num = num.replace('.','')
        num = num.replace(',', '')
        self.num = "{:,}".format(int(num))

#---------- XPath ------ one element ----------------
def axaXlGismoElementFinder(element, action, text=""):
    htmlElement = None
    try:
        waiting = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
        wait = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, element)))
        htmlElement = driver.find_element_by_xpath(element)
        if action == "click":
            driver.execute_script("arguments[0].click();", htmlElement)
        if action == "sendKeys":
            htmlElement.send_keys(text)
    except:
        input('There is a technical error. Element  --> {0}'.format(element))
# -------------------- Close other tabs -------------------------------------
def close_tabs():
    try:
        waiting = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, '//ul[@class = "tabBarItems slds-grid"]/child::li[last()]')))
        wait = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//ul[@class = "tabBarItems slds-grid"]/child::li[last()]')))
        all_tabs = len(driver.find_elements_by_xpath('//ul[@class = "tabBarItems slds-grid"]/li'))
        while all_tabs > 1:
            time.sleep(5)
            waiting = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, '//ul[@class = "tabBarItems slds-grid"]/li[2]/descendant::button[@class = "slds-button slds-button_icon-container slds-button_icon-x-small"]')))
            wait = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//ul[@class = "tabBarItems slds-grid"]/li[2]/descendant::button[@class = "slds-button slds-button_icon-container slds-button_icon-x-small"]')))
            axaXlGismoElementFinder('//ul[@class = "tabBarItems slds-grid"]/li[2]/descendant::button[@class = "slds-button slds-button_icon-container slds-button_icon-x-small"]',"click")
            axaXlGismoElementFinder('//ul[@class = "tabBarItems slds-grid"]/li[2]/descendant::li[@title = "Close Tab"]/a',"click")
            all_tabs = len(driver.find_elements_by_xpath('//ul[@class = "tabBarItems slds-grid"]/li'))
    except: pass
# -------------------- Master Details ---------------------------------------
def masterDetail():
    # ----------------- Insert territory scope ------------------------------
    if not details.cell(AC.territory_scope, AC.master_col).value is None:
        axaXlGismoElementFinder('{0}/descendant::label[text()="Territory Scope"]/following-sibling::div/child::textarea'.format(AC.masterdetails_current_tab), "sendKeys",text=details.cell(AC.territory_scope, AC.master_col).value)
    # ---------------- Insert business activity ---------------------------
    if not details.cell(AC.business_act, AC.master_col).value is None:
        axaXlGismoElementFinder('{0}/descendant::div[@data-aura-class = "cLC_CollapsibleSection cLC11_MasterDetails"][2]/descendant::textarea'.format(AC.masterdetails_current_tab), "sendKeys", text=details.cell(AC.business_act, AC.master_col).value)
    # ----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('{0}{1}'.format(AC.masterdetails_current_tab, AC.save_and_next), "click")
    except:
        input('Please review Master Details.')
#----------------- Local contacts ------------------------------------
def localContact():
    # for each country column in template
    for one_cont in cont_list:
        local_contacts_country = None
        # Expand country dropdown
        try: waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '{0}/descendant::span[contains(text(), "{1}")]'.format(AC.localContact_current_tab, one_cont[AC.local_policy]))))
        except: continue
        local_contacts_country = driver.find_element_by_xpath('{0}/descendant::span[contains(text(), "{1}")]'.format(AC.localContact_current_tab, one_cont[AC.local_policy]))
        if driver.find_element_by_xpath('{0}/descendant::span[contains(text(), "{1}")]/parent::button'.format(AC.localContact_current_tab, one_cont[AC.local_policy])).get_attribute('aria-expanded') == 'false':
            driver.execute_script("arguments[0].click();", local_contacts_country)
        new_cont = 0
        while AC.name + new_cont < len(one_cont):
            # click on add broker or add insured
            if cont_str == 'Broker':
                axaXlGismoElementFinder('{0}/descendant::span[contains(text(), "{1}")]/ancestor::li{2}'.format(AC.localContact_current_tab, one_cont[AC.local_policy], AC.add_broker_contact),'click')
            elif cont_str == 'Insured':
                axaXlGismoElementFinder('{0}/descendant::span[contains(text(), "{1}")]/ancestor::li{2}'.format(AC.localContact_current_tab, one_cont[AC.local_policy], AC.add_insured_contact), 'click')
            # input broker name
            axaXlGismoElementFinder('({0}/descendant::span[contains(text(), "{1}")]/ancestor::li/descendant::input)[{2}]'.format(AC.localContact_current_tab, one_cont[AC.local_policy], (1+new_cont)),'sendKeys',one_cont[AC.name + new_cont])
            # input broker email
            axaXlGismoElementFinder('({0}/descendant::span[contains(text(), "{1}")]/ancestor::li/descendant::input)[{2}]'.format(AC.localContact_current_tab, one_cont[AC.local_policy], (2+new_cont)),'sendKeys',one_cont[AC.email + new_cont])
            # input broker phone
            axaXlGismoElementFinder('({0}/descendant::span[contains(text(), "{1}")]/ancestor::li/descendant::input)[{2}]'.format(AC.localContact_current_tab, one_cont[AC.local_policy], (3+new_cont)),'sendKeys',one_cont[AC.phone + new_cont])
            # increase row by 3 for new contact
            new_cont += 3
    #----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('{0}{1}'.format(AC.localContact_current_tab,AC.save_and_next),"click")
    except:
        input('Please review Local contacts.')
        # ------------------- Deductions and Retentions ----------------------
def deductions():
    for field in columns:
        if field[AC.local_brokerage] == 'Y':
            try: driver.find_element_by_xpath('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[1]/descendant::input'.format(AC.deductions_current_tab, field[AC.local_policy]))
            except: continue
            axaXlGismoElementFinder('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[1]/descendant::input'.format(AC.deductions_current_tab, field[AC.local_policy]),"click")
            if not field[AC.percentage] == 'None':
                axaXlGismoElementFinder('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[2]/descendant::input'.format(AC.deductions_current_tab, field[AC.local_policy]), "sendKeys", field[AC.percentage])
            if not field[AC.flat_amount] == 'None':
                axaXlGismoElementFinder('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[3]/descendant::input'.format(AC.deductions_current_tab, field[AC.local_policy]),"sendKeys",field[AC.flat_amount])
#----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('//div[@class = "slds-tabs_default__content slds-show" and @id = "deductions"]{0}'.format(AC.save_and_next),"click")
    except:
        input('Please review Deductions and Retentions.')
# ------------------- Reinsurance ------------------------------------
def reinsurance():
    while True:
        try:
            waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'{0}/descendant::div[@class = "slds-dueling-list__options"][1]/descendant::li'.format(AC.reinsurance_current_tab))))
            reinsurance_policy_id = driver.find_elements_by_xpath('{0}/descendant::div[@class = "slds-dueling-list__options"][1]/descendant::li'.format(AC.reinsurance_current_tab))
            axaXlGismoElementFinder('{0}/descendant::div[@class = "slds-dueling-list__options"][1]/descendant::li[1]/descendant::span[@class = "slds-truncate"]'.format(AC.reinsurance_current_tab), "click")
            axaXlGismoElementFinder('{0}/descendant::div[@data-aura-class = "cLC_CollapsibleSection cLC11_Reinsurance"]/descendant::button[@title = "Move selection to Selected"]'.format(AC.reinsurance_current_tab), "click")
            reinsurance_policy_id = None
        except:
            break
    #----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('{0}{1}'.format(AC.reinsurance_current_tab,AC.save_and_next),"click")
    except:
        input('Please review Reinsurance.')
# ------------------- Exclusions -------------------------------------
def exclusions_tab():
    waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '{0}/descendant::input[@placeholder="Search Exclusions"]'.format(AC.exclusions_current_tab))))
    for exc in exclusions:
        try:
            select_exclusions = driver.find_element_by_xpath('{0}/descendant::li[@class = "slds-listbox__item"]/descendant::span[contains(text(), "{1}")]'.format(AC.exclusions_current_tab, exc))
            driver.execute_script("arguments[0].click();", select_exclusions)
            axaXlGismoElementFinder('{0}/descendant::button[@title = "Move selection to Selected"]'.format(AC.exclusions_current_tab),"click")
        except:
            print('Exclusion not available - {0}'.format(exc))
            pass
    #----------------- click on save & next------------------------------
    try:
        time.sleep(2)
        axaXlGismoElementFinder('{0}{1}'.format(AC.exclusions_current_tab,AC.save_and_next),"click")
        time.sleep(10)
    except:
        input('Please review Exclusions.')
# ------------------- Claims -----------------------------------------
def claims():
    # Internal claims handler name
    if not details.cell(AC.claim_name,AC.claims_col).value is None:
        axaXlGismoElementFinder('{0}/descendant::label[text()="Internal Claim Handler Name"]/parent::lightning-input/descendant::input'.format(AC.claims_current_tab),"sendKeys",text=details.cell(AC.claim_name,AC.claims_col).value)
    # Internal Claim Handler Email
    if not details.cell(AC.claim_email,AC.claims_col).value is None:
        axaXlGismoElementFinder('{0}/descendant::label[text()="Internal Claim Handler Email"]/parent::lightning-input/descendant::input'.format(AC.claims_current_tab),"sendKeys",text=details.cell(AC.claim_email,AC.claims_col).value)
    #----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('{0}{1}'.format(AC.claims_current_tab,AC.save_and_next),"click")
    except:
        input('Please review Claims.')
# ------------------- Currencies -------------------------------------
def currencies():
    for field in columns:
        try: driver.find_element_by_xpath('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[2]/descendant::select'.format(AC.currencies_current_tab, field[AC.local_policy]))
        except: continue
        if not field[AC.local_currency] == 'None':
            axaXlGismoElementFinder('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[2]/descendant::select'.format(AC.currencies_current_tab, field[AC.local_policy]),"sendKeys",field[AC.local_currency])
        if not field[AC.roe_date] == 'None':
            rateofexchange = driver.find_element_by_xpath('{0}/child::div[contains(text(),"{1}")]/parent::div/following-sibling::div[4]/descendant::input'.format(AC.currencies_current_tab, field[AC.local_policy]))
            rateofexchange.clear()
            x = datetime.strptime(field[AC.roe_date],"%Y-%m-%d %H:%M:%S").strftime("%d-%b-%Y")
            rateofexchange.send_keys(x)
    axaXlGismoElementFinder('//div[@class = "slds-tabs_default__content slds-show" and @id = "currencies"]/descendant::button[@title = "Get OANDA RoE"]',"click")
    #----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('//div[@class = "slds-tabs_default__content slds-show" and @id = "currencies"]{0}'.format(AC.save_and_next),"click")
    except:
        input('Please review currencies.')
# ------------------- Coverages --------------------------------------
def coverages():
    for cover in cov_list:
        try: driver.find_element_by_xpath('//div[@class = "slds-modal__container"]/descendant::span[contains(text(),"{0}")]'.format(cover[AC.local_policy]))
        except: continue
        new_cov = 0
        while AC.cov_row + new_cov < len(cover):
            # click on add coverages
            axaXlGismoElementFinder((AC.add_coverages),"click")
            # select public & product liability
            axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::span[@title = "{0}" and @class = "slds-truncate"]'.format(cover[AC.cov_row + new_cov]),"click")
            # Move selection to selected
            axaXlGismoElementFinder('(//div[@class = "slds-modal__container"]/descendant::button[@title = "Move selection to Selected"])[1]',"click")
            # Select a local policy
            axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::span[contains(text(),"{0}")]'.format(cover[AC.local_policy]),"click")
            # Move selection to selected
            axaXlGismoElementFinder('(//div[@class = "slds-modal__container"]/descendant::button[@title = "Move selection to Selected"])[2]',"click")
            if not cover[AC.limit + new_cov] == 'None':
                # Click on limits tab
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::a[@data-label = "Limits"]', "click")
                # Limit
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::label[text() = "Limit"]/parent::lightning-input/descendant::input',"sendKeys",numformat(cover[AC.limit + new_cov]).num)
                # select trigger
                if not cover[AC.policy_trigger + new_cov] == 'None':
                    driver.find_element_by_xpath('//div[@class = "slds-modal__container"]/descendant::select[@name = "Trigger"]/option[@value = "{0}"]'.format(cover[AC.policy_trigger + new_cov])).click()
                # Aggregate
                if cover[AC.policy_trigger + new_cov] == 'C - Claims Made':
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::select[@name = "Aggregation"]', "sendKeys","Per Event")
                else:
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::select[@name = "Aggregation"]',"sendKeys", "1 times")
                if cover[AC.SIR + new_cov] != 'None':
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::label[text() = "SIR"]/parent::lightning-input/descendant::input',"sendKeys", numformat(cover[AC.SIR + new_cov]).num)
            if not cover[AC.deductible + new_cov] == 'None':
                # click on deductibles
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::a[@data-label = "Deductibles"]',"click")
                # click on add deductibles
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::button[@value = "newDeductible"]',"click")
                # Insert Deductible
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "deductible"]/descendant::input[1]',"sendKeys",numformat(cover[AC.deductible + new_cov]).num)
                # Select type of deductible
                if cover[AC.local_policy][:2] != 'IT':
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "deductible"]/descendant::select',"sendKeys", "Deductible included in the limit")
                else:
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "deductible"]/descendant::select',"sendKeys","Limit in excess of deductible")
            # click on premiums
            axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::a[@data-label = "Premiums"]',"click")
            # click on add Premium
            axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::button[@value = "newPremium"]',"click")
            # Insert Premium
            if not cover[AC.coverage_premium + new_cov] == 'None':
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::input[1]',"sendKeys",numformat(cover[AC.coverage_premium + new_cov]).num)
            # check flat
            if cover[AC.flat + new_cov] == 'Y':
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::input[@type = "checkbox"][1]',"click")
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::select',"sendKeys", 'FLT')
            elif cover[AC.adjustable + new_cov] == 'Y':
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::input[@type = "checkbox"][2]',"click")
                if not cover[AC.rate + new_cov] == 'None':
                    axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::input[4]',"sendKeys",cover[AC.rate + new_cov])
                axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::div[@id = "premiums"]/descendant::select',"sendKeys", 'TO')
            # click on add
            axaXlGismoElementFinder('//div[@class = "slds-modal__container"]/descendant::button[contains(text(),"Add") and @class = "slds-button slds-button_brand"]',"click")
            time.sleep(10)
            try:
                WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, '//div[@class = "slds-modal__container"]/descendant::button[contains(text(),"Add") and @class = "slds-button slds-button_brand"]')))
            except:
                input('Please clear error and add.')
            new_cov += 10
    try:
        # ----------------- click on save & next------------------------------
        axaXlGismoElementFinder('//div[@class = "slds-tabs_default__content slds-show" and @id = "coverages"]{0}'.format(AC.save_and_next), "click")
    except:
        input('Please review Coverages.')
        # ------------------- Additional Insureds ----------------------------
def additionalInsured():
    for field in columns:
        next_ins = 0
        try: waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'{0}/descendant::span[contains(text(), "{1}")]'.format(AC.additionalInsured_current_tab,field[AC.local_policy]))))
        except: continue
        additional_insureds_country = None
        additional_insureds_country = driver.find_element_by_xpath('{0}/descendant::span[contains(text(), "{1}")]'.format(AC.additionalInsured_current_tab, field[AC.local_policy]))
        if driver.find_element_by_xpath('{0}/descendant::span[contains(text(), "{1}")]/parent::button'.format(AC.additionalInsured_current_tab, field[AC.local_policy])).get_attribute('aria-expanded') == 'false':
            driver.execute_script("arguments[0].click();", additional_insureds_country)
        while AC.additional_insureds_name + next_ins < len(field):
            axaXlGismoElementFinder('{0}/descendant::span[contains(text(), "{1}")]/ancestor::li{2}'.format(AC.additionalInsured_current_tab, field[AC.local_policy], AC.add_additional_insured),'click')
            axaXlGismoElementFinder('({0}/descendant::span[contains(text(), "{1}")]/ancestor::li/descendant::input)[{2}]'.format(AC.additionalInsured_current_tab, field[AC.local_policy], (1+next_ins)), 'sendKeys', field[AC.additional_insureds_name + next_ins])
            try:
                if not field[AC.additional_insureds_address + next_ins] == 'None':
                    axaXlGismoElementFinder('({0}/descendant::span[contains(text(), "{1}")]/ancestor::li/descendant::input)[{2}]'.format(AC.additionalInsured_current_tab, field[AC.local_policy], (2+next_ins)), 'sendKeys',field[AC.additional_insureds_address + next_ins])
            except: pass
            next_ins += 2
    #----------------- click on save & next------------------------------
    try:
        axaXlGismoElementFinder('{0}{1}'.format(AC.additionalInsured_current_tab,AC.save_and_next),"click")
    except:
        input('Please review Additional insureds.')
print('Collecting data from template....')

# ---------------- open chrome and navigate to gismo -----------------
print('Running Chromedriver....')
if getattr(sys, 'frozen', False):
  chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
  driver = webdriver.Chrome(chromedriver_path)
  print(chromedriver_path)
else:
  driver = webdriver.Chrome()
#driver.get('https://axaxlgismo.lightning.force.com/one/one.app')
driver.get("https://axaxlgismo--preprod.lightning.force.com/one/one.app")
driver.maximize_window()
# ---------------- click on login button -----------------------------
axaXlGismoElementFinder('//button[@class="button mb24 secondary wide"]',"click")
# ---------------- close all tabs ------------------------------------
close_tabs()
#----------------- click on drop down button -------------------------
axaXlGismoElementFinder('//button[@title="Show Navigation Menu"]',"click")
#----------------- click on policy search------------------------------
axaXlGismoElementFinder('//span[text()="Policy Search"]',"click")
#----------------- click on Master Policy Number ----------------------
axaXlGismoElementFinder('//span[@class="slds-form-element__label" and text()="Master policy number"]',"click")
#----------------- Insert Master Policy Number------------------------------
axaXlGismoElementFinder('//label[text()="Policy Number"]/following-sibling::div/child::input[@type="text"]',action="sendKeys",text=details.cell(AC.genius_policy,AC.master_col).value)
#----------------- click on search------------------------------
axaXlGismoElementFinder('//button[@title="Search"]', "click")
#----------------- click on results------------------------------
axaXlGismoElementFinder('//input[@type="radio" and @name="select"]', "click")
#----------------- click on create------------------------------
axaXlGismoElementFinder('//button[@class="slds-button slds-button_brand" and @title="Create"]', "click")
#----------------- click on Endorse------------------------------
axaXlGismoElementFinder('//button[@class="slds-button slds-button_brand" and text()="Endorse"]', "click")
#----------------- click on master details------------------------------
waiting = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//span[@data-index="masterDetail"]')))
axaXlGismoElementFinder('//span[@data-index="masterDetail"]',"click")
total_tabs = driver.find_elements_by_xpath('//div[@class = "slds-tabs_default__content slds-hide"]')
total = len(total_tabs) + 1
increment = 0
while True:
    time.sleep(5)
    waiting = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@class = "slds-tabs_default__content slds-show"]')))
    current_tab = driver.find_element_by_xpath('(//div[@class = "slds-tabs_default"])[1]/div[@class = "slds-tabs_default__content slds-show"]').get_attribute('id')
    print('Working on {0}...{1}% complete.'.format(current_tab,(increment/total*100)))
    if current_tab == 'masterDetail':
        masterDetail()
    if current_tab == 'additionalInsured':
        additionalInsured()
    if current_tab == 'coverages':
        coverages()
    if current_tab == 'currencies':
        currencies()
    if current_tab == 'claims':
        claims()
    if current_tab == 'exclusions':
        exclusions_tab()
    if current_tab == 'reinsurance':
        reinsurance()
    if current_tab == 'deductions':
        deductions()
    if current_tab == 'localContact':
        localContact()
    t_end = time.time() + 10
    while True:
        if current_tab != driver.find_element_by_xpath('(//div[@class = "slds-tabs_default"])[1]/div[@class = "slds-tabs_default__content slds-show"]').get_attribute('id'):
             break
        if time.time() > t_end:
            input('Please clear error and save.')
            break
    increment += 1
    if total == increment:
        print('Gismo Prefill completed for policy {0}'.format(details.cell(AC.genius_policy,AC.master_col).value))
        break





