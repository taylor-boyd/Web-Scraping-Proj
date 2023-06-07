from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
import time
import difflib

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
            # output them all or compare them all to input
            # how to compare jun-30-2024 to 6/30/2024? authomatically change format ig
            for exp in expirations:
                first_val = exp[0]
                for e in exp:
                    if e != first_val:
                        output[2] = "exp date unclear"
                if output[2] != "exp date unclear":
                    first_val = first_val.replace('-','/')
                    if 'jun' in first_val: # create funcion for this also need to replace '08' with '8' and so on
                        first_val = first_val.replace('Jun','6')
                    if first_val != output[2]:
                        output[2] = first_val
        else:
            output[2] = "exp date not found"

else:
    output[2] = "license not found"
    
print("Output: " + output[0] + " " + output[1] + " " + output[2])

# how to hit return
#search_bar.send_keys(Keys.RETURN)

# printing the content of entire page
#print(driver.page_source)
  
# closing the driver
driver.close()