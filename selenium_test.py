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
with open("/Users/TaylorBoyd/OneDrive - Maletis Beverage/Licensing/wa_june_expiring_licenses.csv",'r') as file:
   
  # reading the CSV file
  csvFile = csv.reader(file)
  entries = []
  temp = []
  for row in csvFile:
      temp = row
      entries.append(temp)

print(entries[1])

customer = ["Kalama Shopping Center","354414","6/30/2023"]
output = [customer[0],customer[1],customer[2]]

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

# find 'License #' text box, click, enter license_no, hit 'Enter'
license_no = wait.until(EC.element_to_be_clickable((By.ID,"Dc-t")))
license_no.send_keys(int(customer[1]))
license_no.send_keys(u'\ue007')

# wait for results to appear
time.sleep(5)

# results table
results = driver.find_element(By.XPATH,"//*[@id='fc_Dc-m1']")

# helper function
def find_expiration():
    expiration = []

    endorsements = driver.find_elements(By.XPATH,"//*[@id='Dc-p1']/tbody")
    for x in endorsements:
        rows = x.find_elements(By.XPATH,"tr")
        for row in rows:
            col = row.find_elements(By.TAG_NAME,"td")
            if col[1].text == customer[1]:
                expiration.append(col[5].text)   

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

# if results table is displayed
if results.is_displayed():
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
        best_match = difflib.get_close_matches(customer[0].upper(),business_names,n=3,cutoff=0.7) # case sensitive 
        expirations = []
        
        # if at least one match was found
        if len(best_match) > 0:
            for match in best_match:
                for col in cols:
                    if col[0].text == match:
                        col[0].click()
                        time.sleep(5)
                        expirations.append(find_expiration()) # find expiration date(s)
                        driver.execute_script("window.history.go(-1)") # go back a page
                        time.sleep(5)

        # if an expiration date was found
        if len(expirations) > 0:
            for exp in expirations:
                first_val = exp[0]
                for e in exp:
                    if e != first_val:
                        output[2] = "exp date unclear"
                if output[2] != "exp date unclear":
                    first_val = first_val.replace('-','/')
                    first_val = month_reformat(first_val)
                    first_val = day_reformat(first_val)
                    if first_val != output[2]:
                        output[2] = first_val
        else:
            output[2] = "exp date not found"

else:
    output[2] = "license not found"
    
print("Output: " + output[0] + " " + output[1] + " " + output[2])
  
# closing the driver
driver.close()