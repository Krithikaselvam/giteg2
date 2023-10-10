from flask import Flask, render_template, request
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    

#@app.route('/upload', methods=['POST'])
#def upload():
    #uploaded_file = request.files['file_input']
    #if uploaded_file:
        #Edge_options = Options()
        #Edge_options.add_argument('--headless')
    
        driver=webdriver.Edge()
        driver.get('http://10.89.1.12:8080/tops/product')
        username_field = driver.find_element(By.XPATH,'//*[@id="j_username"]')
        password_field = driver.find_elements(By.XPATH,'//*[@id="j_password"]')[0]
        username_field.send_keys("krithikaselvam")
        password_field.send_keys("tekV2023")
        login_button = driver.find_elements(By.XPATH,'//*[@id="user-details"]/input[3]')[0]
        login_button.click()
        driver.implicitly_wait(10)
        tbody=driver.find_element(By.XPATH,'/html/body/div[1]/ul/li[5]/a')
        print(tbody.text)
        time.sleep(10)
        table = tbody.find_element(By.XPATH, '//*[@id="productTable"]')
        column_1 = table.find_elements(By.XPATH,'.//tr/td[1]')
        Testplan=[]
        for i in range(1,13):   
            table = tbody.find_element(By.XPATH, '//*[@id="productTable"]')
            column_1 = table.find_elements(By.XPATH,'.//tr/td[1]')
            
            for cell in column_1:
                Testplan.append(cell.text)
            if i<13:
                wait = WebDriverWait(driver, 10)
                my_element = wait.until(EC.element_to_be_clickable((By.ID, 'productTable_next')))
                #next_page_link = driver.find_element(By.XPATH,'//*[@id="productTable_next"]')
                my_element.click()
                time.sleep(5)
            for x in Testplan:
                if x =="":
                 Testplan.remove(x)
        return render_template('index.html', Testplan=Testplan)

        # Save or process the uploaded file as needed
        # For example: uploaded_file.save('path_to_save_file')

      

if __name__ == '__main__':
    app.run()
