from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from PyQt5.QtMultimedia import *
from PyQt5.QtMultimediaWidgets import *

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement

import os
import sys
import time
import itertools 
import difflib
import csv
from pathlib import Path

class MainWindow(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.setWindowTitle('WA License Checking Tool')
        self.setGeometry(100, 100, 300, 150)

        layout = QGridLayout()
        self.setLayout(layout)

        # file selection
        file_browse = QPushButton('Browse')
        file_browse.clicked.connect(self.open_file_dialog)
        self.filename_edit = QLineEdit()

        # check licenses button
        button = QPushButton('Check Licenses')
        button.clicked.connect(self.open_license_checker)

        layout.addWidget(QLabel('File:'), 0, 0)
        layout.addWidget(self.filename_edit, 0, 1)
        layout.addWidget(file_browse, 0 ,2)
        layout.addWidget(button,1,0)

      
        self.show()

    def open_file_dialog(self):
        filename, ok = QFileDialog.getOpenFileName(
            self,
            "Select a File",  
            "",
            "csv(*.csv)"
        )
        if filename:
            path = Path(filename)
            self.filename_edit.setText(str(path))

    def open_license_checker(self):
        input_file = self.filename_edit.text()
        if '.csv' in input_file:

            # if given directory exists
            if os.path.exists(input_file):

                # opening the CSV file
                with open(input_file,'r') as file:
                    # reading the CSV file
                    csvFile = csv.reader(file)
                    entries = []
                    temp = []
                    for row in csvFile:
                        temp = row
                        entries.append(temp)
                
                def check_columns(input_headers, missing, misplaced, length):
                    col_headers = ['Customer ID', 'Customer Name', 'Sales Rep', 'Shipping Address', 'Product Group', 'Distribution Area', 'State License Num', 'License Exp Date']
                    for c in col_headers:
                        if c not in input_headers:
                            missing.append(c)
                    if len(input_headers) == 8:
                        indices = [0,1,2,3,4,5,6,7]
                        for i in indices:
                            if col_headers[i] != input_headers[i]:
                                misplaced.append(col_headers[i])
                    elif len(input_headers) < 8:
                        length = 'Columns Missing'
                    else:
                        length = 'Extra Columns Found'
                    return missing,misplaced,length

                # check for correct columns in input file
                missing = []
                misplaced = []
                length = "Correct"
                missing,misplaced,length = check_columns(entries[0],missing,misplaced, length)
                if (len(missing) < 1 and len(misplaced) < 1 and length == 'Correct'):
                    
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
                    split_file = input_file.split(".csv")
                    output_file = split_file[0] + '_checked.csv'
                    with open(output_file ,'w', newline='') as f:
                    
                        # reading the CSV file
                        writer = csv.writer(f)
                        for row in output:
                            writer.writerow(row)
                    
                    # closing the driver
                    driver.close()
                    QMessageBox.about(self, 'License Check Complete', "The check has completed and the results have been written to: " + str(output_file))

                # else output message detailing missing or misplaced columns
                else:
                    if (len(missing)>0 and len(misplaced)==0):
                        QMessageBox.about(self, 'Input File Is Incorrectly Formatted', "The following columns weren't found: " + str(missing))
                    elif (len(missing)==0 and len(misplaced)>0):
                        QMessageBox.about(self, 'Input File Is Incorrectly Formatted', "The following columns are in the wrong spot: " + str(misplaced))
                    elif (len(missing)>0 and len(misplaced)>0):
                        QMessageBox.about(self, 'Input File Is Incorrectly Formatted', "The following columns weren't found: " + str(missing) + 
                                          "\nThe following columns are in the wrong spot: " + str(misplaced))
                    elif length == 'Extra Columns Found':
                        QMessageBox.about(self, 'Input File Is Incorrectly Formatted', "Extra columns were found - remove all unnecessary columns")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())