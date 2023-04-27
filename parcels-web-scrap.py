import csv
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
import sys



# First we check if the code is executed from Python or from the .exe, because the way to get the 
# path from where the code is executed is different in both cases.
if getattr(sys, 'frozen', False):
    # Runs from the .exe
    application_path = os.path.dirname(sys.executable)
else:
    # Runs from Python
    application_path = os.path.dirname(os.path.abspath(__file__))

# Once we have the path, we read the configuration file and obtain the parameters
with open(os.path.join(application_path, "config.json")) as jsonFile:
    jsonObject = json.load(jsonFile)
    jsonFile.close()

# Parameter reading
url = jsonObject['url']
csv_file = os.path.join(application_path, jsonObject['csv_filename'])
csv_encoding = jsonObject['csv_encoding']
selenium_webdriver_useragent = jsonObject['selenium_webdriver_useragent']
header = jsonObject['header']

iframe_xpath = jsonObject['iframe_xpath']

output_SuitNo_xpath = jsonObject['output_SuitNo_xpath']
output_ParcelNo_xpath = jsonObject['output_ParcelNo_xpath']
output_Owner_xpath = jsonObject['output_Owner_xpath']
CoOwner_xpath = jsonObject['CoOwner_xpath']
LegalDesc_xpath = jsonObject['LegalDesc_xpath']
PropertyAddress_xpath = jsonObject['PropertyAddress_xpath']
PropertyCity_xpath = jsonObject['PropertyCity_xpath']
PropertyState_xpath = jsonObject['PropertyState_xpath']
DateSold_xpath = jsonObject['DateSold_xpath']
PurchasePrice_xpath = jsonObject['PurchasePrice_xpath']
Judgment_xpath = jsonObject['Judgment_xpath']
Excess_xpath = jsonObject['Excess_xpath']
Purchaser_xpath = jsonObject['Purchaser_xpath']
PurchaserAddress_xpath = jsonObject['PurchaserAddress_xpath']
PurchaserCity_xpath = jsonObject['PurchaserCity_xpath']
PurchaserState_xpath = jsonObject['PurchaserState_xpath']
PurchaserZip_xpath = jsonObject['PurchaserZip_xpath']
HearingDate_xpath = jsonObject['HearingDate_xpath']
HearingTime_xpath = jsonObject['HearingTime_xpath']
ContinuedDate_xpath = jsonObject['ContinuedDate_xpath']
ContinuedTime_xpath = jsonObject['ContinuedTime_xpath']
Confirmed_xpath = jsonObject['Confirmed_xpath']
SetAside_xpath = jsonObject['SetAside_xpath']
Appealed_xpath = jsonObject['Appealed_xpath']
CountyPaid_xpath = jsonObject['CountyPaid_xpath']
StatePaid_xpath = jsonObject['StatePaid_xpath']
ExcessAppFiled_xpath = jsonObject['ExcessAppFiled_xpath']
ExcessDenied_xpath = jsonObject['ExcessDenied_xpath']
ExcessAppealed_xpath = jsonObject['ExcessAppealed_xpath']
ExcessPaid_xpath = jsonObject['ExcessPaid_xpath']




# Chrome driver configuration
options = Options()
options.add_argument(selenium_webdriver_useragent)
driver = webdriver.Chrome(options=options)

# Access the URL and wait for the page to load.
driver.get(url)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, iframe_xpath)))


# Create a CSV file and write the headers
with open(csv_file, "w", newline="", encoding=csv_encoding) as csv_file:
    writer = csv.writer(csv_file, delimiter="\t")

    writer.writerow(header)

    # Find the iframe
    iframe = driver.find_elements(By.XPATH, iframe_xpath)

    # Change Selenium context to point to iframe
    driver.switch_to.frame(iframe[0])

    # Now you can search and manipulate the elements inside the iframe
    # To know if we have reached the end, detect if there is a 'Next' button.
    botones_siguiente = driver.find_elements(By.XPATH, "//input[@name='Next']")
    more_cases = True
    if len(botones_siguiente) == 0:
          more_cases = False
          print('No more records')

    
    


    while more_cases:
        # Find all records on the current page and write them to the CSV file
        output_SuitNo = driver.find_elements(By.XPATH, output_SuitNo_xpath)
        output_ParcelNo = driver.find_elements(By.XPATH, output_ParcelNo_xpath)
        output_Owner = driver.find_elements(By.XPATH, output_Owner_xpath)
        CoOwner = driver.find_elements(By.XPATH, CoOwner_xpath)
        LegalDesc = driver.find_elements(By.XPATH, LegalDesc_xpath)
        PropertyAddress = driver.find_elements(By.XPATH, PropertyAddress_xpath)
        PropertyCity = driver.find_elements(By.XPATH, PropertyCity_xpath)
        PropertyState = driver.find_elements(By.XPATH, PropertyState_xpath)
        DateSold = driver.find_elements(By.XPATH, DateSold_xpath)
        PurchasePrice = driver.find_elements(By.XPATH, PurchasePrice_xpath)
        Judgment = driver.find_elements(By.XPATH, Judgment_xpath)
        Excess = driver.find_elements(By.XPATH, Excess_xpath)
        Purchaser = driver.find_elements(By.XPATH, Purchaser_xpath)
        PurchaserAddress = driver.find_elements(By.XPATH, PurchaserAddress_xpath)
        PurchaserCity = driver.find_elements(By.XPATH, PurchaserCity_xpath)
        PurchaserState = driver.find_elements(By.XPATH, PurchaserState_xpath)
        PurchaserZip = driver.find_elements(By.XPATH, PurchaserZip_xpath)
        HearingDate = driver.find_elements(By.XPATH, HearingDate_xpath)
        HearingTime = driver.find_elements(By.XPATH, HearingTime_xpath)
        ContinuedDate = driver.find_elements(By.XPATH, ContinuedDate_xpath)
        ContinuedTime = driver.find_elements(By.XPATH, ContinuedTime_xpath)
        Confirmed = driver.find_elements(By.XPATH, Confirmed_xpath)
        SetAside = driver.find_elements(By.XPATH, SetAside_xpath)
        Appealed = driver.find_elements(By.XPATH, Appealed_xpath)
        CountyPaid = driver.find_elements(By.XPATH, CountyPaid_xpath)
        StatePaid = driver.find_elements(By.XPATH, StatePaid_xpath)
        ExcessAppFiled = driver.find_elements(By.XPATH, ExcessAppFiled_xpath)
        ExcessDenied = driver.find_elements(By.XPATH, ExcessDenied_xpath)
        ExcessAppealed = driver.find_elements(By.XPATH, ExcessAppealed_xpath)
        ExcessPaid = driver.find_elements(By.XPATH, ExcessPaid_xpath)


        # Write output csv infomation
        writer.writerow([output_SuitNo[0].text, 
                         output_ParcelNo[0].text, 
                         output_Owner[0].text, 
                         CoOwner[0].text, 
                         LegalDesc[0].text, 
                         PropertyAddress[0].text, 
                         PropertyCity[0].text, 
                         PropertyState[0].text, 
                         DateSold[0].text, 
                         PurchasePrice[0].text, 
                         Judgment[0].text, 
                         Excess[0].text, 
                         Purchaser[0].text, 
                         PurchaserAddress[0].text, 
                         PurchaserCity[0].text, 
                         PurchaserState[0].text, 
                         PurchaserZip[0].text, 
                         HearingDate[0].text, 
                         HearingTime[0].text, 
                         ContinuedDate[0].text, 
                         ContinuedTime[0].text, 
                         Confirmed[0].text, 
                         SetAside[0].text, 
                         Appealed[0].text, 
                         CountyPaid[0].text, 
                         StatePaid[0].text, 
                         ExcessAppFiled[0].text, 
                         ExcessDenied[0].text, 
                         ExcessAppealed[0].text, 
                         ExcessPaid[0].text])

        # Click to obtain the following record
        botones_siguiente[0].click()

        # We see if we have reached the end
        botones_siguiente = driver.find_elements(By.XPATH, "/html/body/table/tbody/tr[1]/td/table[1]/tbody/tr/td[3]/input")
        if len(botones_siguiente) == 0:
            # No more pages, exit the loop
            more_cases = False
            print('No more records')


    # When you are done working with the iframe, be sure to return to the parent context
    driver.switch_to.default_content()

# Close the browser
driver.quit()
