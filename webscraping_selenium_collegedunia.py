from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
import pandas as pd
import excel2json
import pandas

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get("https://collegedunia.com/usa/college/1090-harvard-university-cambridge")
driver.implicitly_wait(10)


driver.find_element(By.XPATH, "//a[text()='Courses & Fees']").click()

#courses
courses = driver.find_elements(By.XPATH,"//h2[contains(@class,'jsx-3847118977 jsx-2133140133')]/a") #this is for anchor tag text 
for x in courses:
    print(x.text)

#fees
fees = driver.find_elements(By.XPATH,"//span[contains(@class,'jsx-3847118977 jsx-2133140133 fees font-weight-bolder')]")
for y in fees:
    print(y.text)

#duration
duration = driver.find_elements(By.XPATH,"//span[contains(@class,'jsx-3847118977 jsx-2133140133 text-capitalize')]")
for z in duration:
    print(z.text)

course_list = []
fees_list = []
duration_list= []

for c in courses:
    course_list.append(c.text)

for f in fees:
    fees_list.append(f.text)

for d in duration:
    duration_list.append(d.text)

finallist = zip(course_list,fees_list,duration_list)

wb=Workbook()
sheet1=wb.active
sheet1.title="collegeduniyawebscraping"
sheet1.append(["courses","fees","duration"])

for x in list(finallist):
    sheet1.append(x)

wb.save("collegeduniyawebscraping.xlsx")



excel_data_df = pandas.read_excel('collegeduniyawebscraping.xlsx', sheet_name='collegeduniyawebscraping')

json_str = excel_data_df.to_json()

with open("sample.json", "w") as outfile:
    outfile.write(json_str)

print('Excel Sheet to JSON:\n', json_str)