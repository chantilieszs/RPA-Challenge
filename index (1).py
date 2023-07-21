from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

df= pd.read_excel("challenge.xlsx")

listDics = []

for row in df.itertuples(index=False):
    dic = dict(zip(df.columns, row))
    listDics.append(dic)

driver = webdriver.Chrome()
driver.get("https://www.rpachallenge.com/")

time.sleep(2)

driver.find_element(By.TAG_NAME, "button").click();
for dic in listDics:
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelRole']").send_keys(dic["Role in Company"])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelLastName']").send_keys(dic["Last Name "])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelFirstName']").send_keys(dic["First Name"])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelCompanyName']").send_keys(dic["Company Name"])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelEmail']").send_keys(dic["Email"])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelAddress']").send_keys(dic["Address"])
    time.sleep(.1)
    driver.find_element(By.XPATH, "//*[@ng-reflect-name='labelPhone']").send_keys(dic["Phone Number"])
    time.sleep(.1)
    btnSubmit = driver.find_element(By.XPATH, "//input[@type='submit']").click()

resultado1 = driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[1]")
resultado2 = driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]")

time.sleep(2)

arquivo = open("result.txt", "w")
arquivo.write(str(resultado1.text))
arquivo.write(" ")
arquivo.write(str(resultado2.text))

driver.quit()

# wb = openpyxl.load_workbook("challenge.xlsx")
# name = wb.get_sheet_names();
# print(name)
# type(name)
# ws = wb.get_sheet_by_name("Sheet1")
# listDict = []
# for column in ws.columns:
#     dic = {column[0].value: column[1].value}
#     listDict.append(dic)
# print(listDict)
# firstRow = df.iloc[0]

# dic = dict(zip(df.columns, firstRow))
# labelValue = ["Address", "Company Name", "First Name", "Phone Number", "Email", "Role in Company", "Role in Company", "Last Name"]

# for value in labelValue:
#     labelElement = driver.find_element(By.XPATH, f"//label[contains(text(), '{value}')]")
#     if value == "Address":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("Rua Laranjeira")
#     if value == "Company Name":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("Fiskal Digital")
#     if value == "First Name":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("Lucas")
#     if value == "Phone Number":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("(31) 99197-8031")
#     if value == "Email":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("gelanlucas@gmail.com")
#     if value == "Role in Company":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("Developer")
#     if value == "Last Name":
#         campo_input = labelElement.find_element(By.XPATH, "./following-sibling::input")
#         campo_input.send_keys("Bueno")