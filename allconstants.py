# convert to pyinstaller
# cd C:\Users\X134391\Desktop
# mkdir newfoldername
# xcopy /S "C:\Users\X134391\Desktop\Projects 2020\Gismo Prefill Tool\project" "C:\Users\X134391\Desktop\chrome86"
# pyinstaller -F --add-binary "C:\Users\X134391\Desktop\prefill-01-12-21\chromedriver.exe";"." gismoSpiderHome.py


# --------------- All columns in template----------
master_col = 2
exclusions_col = 3
add_exc_col = 4
claims_col = 6
# --------------- All rows in template-------------
territory_scope = 4
business_act = 5
genius_policy = 6
# claims
claim_name = 4
claim_email = 5
# country
country = 2
local_policy = 3
# coverages
cov_row = 4
policy_trigger = 5
deductible = 6
SIR = 7
coverage_premium = 8
flat = 9
adjustable = 10
rate = 11
turnover = 12
limit = 13
# Local contacts
name = 4
email = 5
phone = 6
# currency
local_currency = 4
roe_date = 5
# Deductions
local_brokerage = 6
percentage = 7
flat_amount = 8
# additional insureds
additional_insureds_name = 9
additional_insureds_address = 10

masterdetails_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "masterDetail"]'
localContact_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "localContact"]'
add_broker_contact = '/descendant::button[contains(text(), "Add Broker Contact")]'
add_insured_contact = '/descendant::button[contains(text(), "Add Insured Contact")]'
additionalInsured_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "additionalInsured"]'
add_additional_insured = '/descendant::button[contains(text(), "Add Additional Insured")]'
currencies_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "currencies"]/descendant::div[@class = "slds-size_1-of-8 slds-grid slds-wrap"]'
add_coverages = '//div[@class = "slds-tabs_default__content slds-show" and @id = "coverages"]/descendant::button[@value = "addCoverage"]'
exclusions_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "exclusions"]'
claims_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "claims"]'
reinsurance_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "reinsurance"]'
deductions_current_tab = '//div[@class = "slds-tabs_default__content slds-show" and @id = "deductions"]/descendant::div[@class = "slds-size_1-of-6 slds-wrap slds-grid"]'
save_and_next = '/descendant::button[@class = "slds-button slds-button_brand slds-float_right" and @title = "Save & Next"]'