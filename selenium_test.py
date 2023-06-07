from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
import time

license_no1 = 12345

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

#Recaptcha
iFrame = wait.until(EC.visibility_of_element_located((By.XPATH,"//*[@id='Dc-33']/div/div/iframe")))
driver.switch_to.frame(iFrame)
recaptcha = wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='recaptcha-anchor']")))
recaptcha.click()
driver.switch_to.default_content()
time.sleep(20)

# find 'License #' text box, click, enter license_no, hit 'Enter'
license_no = wait.until(EC.element_to_be_clickable((By.ID,"Dc-t")))
license_no.send_keys(license_no1)
license_no.send_keys(u'\ue007')

time.sleep(8)

# if "//*[@id='caption2_Dc-63']" is visible then move on, if "//*[@id='caption2_Dc-62']" is visible then results were found
# and we want to know how many table rows are visible
# "//*[@id='fc_Dc-m1']" is results table path

# see if results table is displayed
results = driver.find_element(By.XPATH,"//*[@id='fc_Dc-m1']")

if results.is_displayed():
    # look through rows and select correct name
    content = driver.find_elements(By.XPATH,"//*[@id='Dc-u1']/tbody")
    for x in content:
        print('hi1')
        #for row in x:
        rows = x.find_elements(By.TAG_NAME,"td")
        for row in rows:
            print(row.text)
        print('hi2')
    # continue with steps to find expiration date
# else output that license no was not found

# how to hit return
#search_bar.send_keys(Keys.RETURN)

# printing the content of entire page
#print(driver.page_source)

#Finding number of Rows
#rowsNumber = driver.find_elements(By.XPATH, "//*[@id=\"vc_Dc-3\"]/div[5]/table")
#print(rowsNumber)
#rowCount = rowsNumber.size()
#print("No of rows in this table : " + rowCount)
  
# closing the driver
driver.close()