import pandas as pd
import openpyxl

from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options



worksheet = ref_work_book.active
file_path = ref_work_sheet['B1'].value
wb1 = openpyxl.load_workbook(file_path)
ws1 = wb1.active
df = pd.read_excel(file_path) 
operator = ref_work_sheet['B3'].value
colum_name1 = ref_work_sheet['B2'].value
colum_name2 = ref_work_sheet['B4'].value
colum_name3 = ref_work_sheet['F1'].value
date1 = df[colum_name1]
date2 = df[colum_name2]
date1 = pd.to_datetime(df[colum_name1])#infer_datetime_format=True)
date2 = pd.to_datetime(df[colum_name2])#infer_datetime_format=True)

TP = df[colum_name3]
operators = {
    ">": lambda a, b: a > b,
    "<": lambda a, b: a < b,
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
    ">=": lambda a, b: a >= b,
    "<=": lambda a, b: a <= b,
}

def highlight_empty(x):
    if pd.isnull(x):
        return 'background-color: yellow'


if operator not in operators:
    if ">" in operator:#>5
        result = (date2 - date1).dt.days > int(ref_work_sheet['B3'].value[-1])
        highlight = lambda x: ['background-color: yellow' if v else '' for v in x]   
    if "<" in operator:
        result = (date2 - date1).dt.days < int(ref_work_sheet['B3'].value[-1])
        highlight = lambda x: ['background-color: skyblue' if v else '' for v in x] 
else:
    func = operators[operator]
    result = func(df[colum_name1], df[colum_name2])
    highlight = lambda x: ['background-color: red' if v else '' for v in x]

styled = df.style.apply(highlight,subset=pd.IndexSlice[result,[colum_name1, colum_name2]])


'''chrome_options = Options()
chrome_options.add_argument('--headless')
driver=webdriver.Chrome(options=chrome_options)
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
          Testplan.remove(x)'''
Test_plan   = ref_work_sheet['B6'].value   
start_row = 7
colnames = []
for cell in ref_work_sheet['B']:
    if cell.row >= start_row and cell.value is not None:
        colnames.append(cell.value)
'''if Test_plan in Testplan:
    styled.applymap(highlight_empty, subset=colnames)'''

start_row = 7
valnames = []
for cell in ref_work_sheet['C']:
    if cell.row >= start_row and cell.value is not None:
        valnames.append(cell.value)

colum_name3 = ref_work_sheet['F1'].value

'''def highlight_missing_value(value):
    for colum_name,val_name in zip(colnames,valnames):
        if colum_name in 
    if pd.notna(value) and value != valnames :
        return 'background-color: green'
    else:
        return '''
TP = df[colum_name3]
for testplan in TP:
    if testplan == Test_plan: 
        for column_name, value in zip(colnames, valnames):
           mask = (df[column_name] != value) & (df[column_name].notnull())
           styled.applymap(lambda x: 'background-color: green' if x != value else '', subset=pd.IndexSlice[mask, column_name])
           styled.applymap(highlight_empty, subset=colnames)
           

modified = ref_work_sheet['B5'].value
styled.to_excel(modified)





