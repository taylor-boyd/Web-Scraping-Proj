from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
import time
import difflib
import csv

# opening the CSV file
with open("/Users/TaylorBoyd/OneDrive - Maletis Beverage/Licensing/wa_june_licenses.csv",'r') as file:
   
  # reading the CSV file
  csvFile = csv.reader(file)
  entries = []
  temp = []
  for row in csvFile:
      temp = row
      entries.append(temp)

# initial setup 
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome() # options=chrome_options
driver.get("https://secure.dor.wa.gov/gteunauth/_/#1")

# Setup wait for later
wait = WebDriverWait(driver, 10)

# store first window
window_before = driver.window_handles[0]

# find 'Business Lookup' look and click
business_lookup = wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='l_Dd-8-1']")))
business_lookup.click()

# Recaptcha
iFrame = wait.until(EC.visibility_of_element_located((By.XPATH,"//*[@id='Dc-33']/div/div/iframe")))
driver.switch_to.frame(iFrame)
recaptcha = wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='recaptcha-anchor']")))
recaptcha.click()
driver.switch_to.default_content()

# wait for user to select recaptcha photos
time.sleep(25)

# helper function
def find_expiration(customer):
    expiration = []

    print(customer[6])

    #endorsements = driver.find_elements(By.XPATH,"//*[@id='Dc-p1']/tbody")
    endorsement_table = driver.find_elements(By.CSS_SELECTOR,"[aria-label='Endorsements']")
    e2 = WebElement
    for e in endorsement_table:
        if len(e.text) > 5:
            e2 = e
    results = e2.find_elements(By.TAG_NAME,"tr")

    index = 0
    for r in results:
        if index > 0:
            temp = r.text
            row = temp.split()
            if customer[6] in row:
                length = len(row)
                expiration.append(row[length-2])
        index += 1

    return expiration

# reformat helper function
def month_reformat(exp_date):
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    index = 1
    for m in months:
        if m in exp_date:
            exp_date = exp_date.replace(m,str(index))
        index += 1
    return exp_date

# reformat helper function
def day_reformat(exp_date):
    days = ['01-','02-','03-','04-','05-','06-','07-','08-','09-']
    index = 1
    for d in days:
        if d in exp_date:
            replacement = str(index) + '-'
            exp_date = exp_date.replace(d,replacement)
    return exp_date

# looks through results table
def select_results(customer):
    # look through 'Business name' column and select row with best match
    content = driver.find_elements(By.XPATH,"//*[@id='Dc-u1']/tbody") 
    for x in content:
        rows = x.find_elements(By.TAG_NAME,"tr")
        cols = []
        business_names = []
        for row in rows:
            col = row.find_elements(By.TAG_NAME,"td")
            cols.append(col)
            business_names.append(col[0].text)
        best_match = difflib.get_close_matches(customer[1].upper(),business_names,n=3,cutoff=0.5) # case sensitive 
        expirations = []
            
        # if at least one match was found
        if len(best_match) > 0:
            for match in best_match:
                for col in cols:
                    if col[0].text == match:
                        col[0].click()
                        time.sleep(5)
                        expirations.append(find_expiration(customer)) # find expiration date(s)
                        driver.execute_script("window.history.go(-1)") # go back a page
                        time.sleep(5)
        else:
            customer[7] = "business name not recognized"

        # if an expiration date was found
        if len(expirations) > 0:
            for exp in expirations:
                first_val = exp[0]
                for e in exp:
                    if e != first_val:
                        customer[7] = "exp date unclear"
                if customer[7] != "exp date unclear":
                    first_val = first_val.replace('-','/')
                    first_val = month_reformat(first_val)
                    first_val = day_reformat(first_val)
                    if first_val != customer[7]:
                        customer[7] = first_val
        else:
            customer[7] = "exp date not found"
    return customer

index = 0
output = []

for customer in entries:
    if index > 0:
        # find 'License #' text box, click, enter license_no, hit 'Enter'
        license_no = wait.until(EC.element_to_be_clickable((By.ID,"Dc-t")))
        license_no.send_keys(int(customer[6])) # license num
        license_no.send_keys(u'\ue007')

        # wait for results to appear
        time.sleep(5)

        # results table
        results = driver.find_element(By.XPATH,"//*[@id='fc_Dc-m1']")

        # if results table is displayed
        if results.is_displayed():
            select_results(customer)
        else:
            customer[7] = "license not found"

        if customer[7] == "license not found":
            customer[6] = '0' + str(customer[6])
            # find 'License #' text box, click, delete what's in there
            license_no2 = wait.until(EC.element_to_be_clickable((By.ID,"Dc-t")))
            license_no2.send_keys(Keys.CONTROL,'a') # license num
            license_no2.send_keys(Keys.BACKSPACE)
            license_no2.send_keys(str(customer[6])) # license num
            license_no2.send_keys(u'\ue007')
            # wait for results to appear
            time.sleep(5)

            # results table
            results = driver.find_element(By.XPATH,"//*[@id='fc_Dc-m1']")

            # if results table is displayed
            if results.is_displayed():
                select_results(customer)
            else:
                customer[7] = "license not found"
                customer[6] = int(str(customer[6])[1:])

        # find 'License #' text box, click, delete what's in there
        license_no3 = wait.until(EC.element_to_be_clickable((By.ID,"Dc-t")))
        license_no3.send_keys(Keys.CONTROL,'a') # license num
        license_no3.send_keys(Keys.BACKSPACE)

        output.append(customer)
    else:
        output.append(customer)
    index += 1

# opening the CSV file in write mode
with open("/Users/TaylorBoyd/OneDrive - Maletis Beverage/Licensing/wa_updated_licenses2.csv",'w') as f:
   
  # reading the CSV file
  writer = csv.writer(f)
  for row in output:
      writer.writerow(row)
  
# closing the driver
driver.close()

# functionality to add:
# leading zero if license # not found
# look under 'legal name' as well as 'business name' if good match for 'business name' is not found