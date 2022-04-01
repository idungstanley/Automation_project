
import ast, requests, datetime
from xml.dom import minidom

import selenium
from datetime import datetime, date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time as sl

import win32com.client
from win32com.client import Dispatch

import os, os.path, shutil, getpass, re, glob, socket, zipfile

def kill_by_process_name_shell(name):
    os.system("taskkill /f /im " + name)

try: 
    kill_by_process_name_shell("EXCEL.EXE")
except:
    pass

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(options=chrome_options,executable_path=r"D:\Reporting Python\chromedriver.exe")
wait = WebDriverWait(driver, 600)

file_2G = "D:/Report Automation/BW Orange Report Automation/FME Dashboard/Output/FMEDailyNUR_import.xlsx"
file_2G_name = "FMEDailyNUR_import.xlsx"

file_3G = "D:/Report Automation/BW Orange Report Automation/FME Dashboard/Output/FMEDailyNUR_3G_Import.xlsx"
file_3G_name = "FMEDailyNUR_3G_Import.xlsx"

file_4G = "D:/Report Automation/BW Orange Report Automation/FME Dashboard/Output/FMEDailyNUR_4G_Import.xlsx"
file_4G_name = "FMEDailyNUR_4G_Import.xlsx"

with open('D:\\Report Automation\\Login_details\\SAR_AMS_OWS_login.txt','r') as credentials:
        details = credentials.readlines()
        username,password = details[0].strip(),details[1].strip()

def getAttachInfo():
    
    serviceURL = 'https://15fg-saapp.teleows.com/ws/rest/15fg/RNOC_operation_inbound/get_Email_Single_attachment'
    
    title = "OBW 2G_3G_4G_NUR_Site Level"
        
    getData = 'title='+title
    
    headers = {'Content-type':'application/x-www-form-urlencoded'}
            
    response = requests.post(
        serviceURL,
        verify=False,
        auth=(username,password),
        # params=PARAMS,
        headers=headers,
        data=getData,
        timeout=None
    )

    if response:
        resultValue = ast.literal_eval(response.text)
        
        attachment_name  = resultValue['attachment_name']
        attachment_ID  = resultValue['attachment_ID']
        batch_ID = resultValue['batch_ID']
        
        print(f"""
        attachment_name  is {attachment_name}
        attachment_ID  is {attachment_ID}
        batch_ID is {batch_ID}
        """)
        
        print(resultValue)
        
        getMailParam(batch_ID,attachment_ID)

        return True

def getMailParam(batch_ID,attachment_ID):
    
    batchid = batch_ID
    attachid = attachment_ID

    myurl = "https://15fg-saapp.teleows.com/app/fileservice/get?batchId="+batchid+"&attachmentId="+attachid

    print(myurl)
    
    driver.maximize_window()
    driver.get(myurl)
    sl.sleep(2)
    driver.refresh()
    sl.sleep(5)
    driver.find_element_by_id("usernameInput").send_keys(username)
    driver.find_element_by_id("password").send_keys(password)
    driver.find_element_by_id("btn_submit").click()
    sl.sleep(5)
    print("Downloading Email Attachment")
    checkFile()
    MoveFile()
    
    driver.quit()
    
def dwnldFile():
    fpath = "C:\\Users\\"+getpass.getuser()+"\\Downloads"
    if any(File.endswith(".xlsx") for File in os.listdir(fpath)):
        print("Download File Successfully")
        return True 
    else:
        return False
    
def checkFile():
    while dwnldFile() is not True:
        print("Waiting for Download File")    
        sl.sleep(10)
        dwnldFile()

def checkold():
    folder1 = "D:\\Report Automation\\BW Orange Report Automation\\FME Dashboard\\Input File\\"
    folder2 = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\"    

    Allfolder = [folder1, folder2]

    for f_dir in Allfolder:
        for the_file in os.listdir(f_dir):
            file_path = os.path.join(f_dir, the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    print("file has been Deleted.....")
                    sl.sleep(7)  
                #elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(e)

    sl.sleep(5)

def MoveFile():
    shutil.move("C:\\Users\\"+getpass.getuser()+"\\Downloads\\OBW 2G_3G_4G_NUR_Site Level.xlsx", "D:\\Report Automation\\BW Orange Report Automation\\FME Dashboard\\Input File")
    print("file move successfully")

def ExtractData():
    macroDr = "D:\\Report Automation\\BW Orange Report Automation\\FME Dashboard\\OBW_FME_NUR_WB.xlsm"
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Application.Visible = True
    wb = xl.Workbooks.Open(os.path.abspath(macroDr))
    print("Loading NUR Data....")
    print("Extracting NUR Data -  Running")
    xl.Application.Run("'OBW_FME_NUR_WB.xlsm'!OBW_NUR_site_level")
    xl.Application.Quit()
    print("Completed SUCCESSFUL!!!")

def Upload2GNUR():

    ows2Gurl = "https://1at7-frapp.teleows.com/servicecreator/spl/FME_Performance_Dashboard/fpd_fme_performance_rating_2g_import.spl"
   
    with open("D:\\Report Automation\\Login_details\\BW_Orange_ows_login.txt",'r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.maximize_window()

    driver.get(ows2Gurl)
    sl.sleep(5)
    driver.refresh()
    sl.sleep(3)

    username_tag = 'usernameInput'
    x_arg = f'//input[contains(@id,"{username_tag}")]'
    usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    usernameInput.send_keys(username)

    password_tag = 'password'
    x_arg = f'//input[contains(@id,"{password_tag}")]'
    usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    usernameInput.send_keys(password)

    login = 'btn_submit'
    x_arg = f'//div[contains(@id,"{login}")]'
    loginBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    loginBtn.click()

    sl.sleep(3)

    attachment = '_uploadFile'
    x_arg = f'//form/input[contains(@name,"{attachment}")]'

    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    print(attachmentInput)
    attachmentInput.send_keys(file_2G)
    sl.sleep(5)

    file_path_1_array = file_2G.split('\\')
    file_path_1_title = file_path_1_array[len(file_path_1_array) - 1].split('.')[0]
    print(file_path_1_title)
    x_arg = f'//div[contains(@class,"filelist_cls")]/a[contains(text(),"{file_2G_name}")]'
    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))

    print("Uploading file attached")
    
    submit = 'ExcelImportPanel1_submit'
    x_arg = f'//div[contains(@id,"{submit}")]'
    submitBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    sl.sleep(5)

    submitBtn.click()

    importTb= 'ExcelImportPanel1_msg'
    x_arg = f'//div[contains(@id,"{importTb}")]'
    importBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))

    sl.sleep(15)
    
    print("2G NUR Import Completed...")

def Upload3GNUR():

    ows3Gurl = "https://1at7-frapp.teleows.com/servicecreator/spl/FME_Performance_Dashboard/fpd_fme_performance_rating_3g_import.spl"

    with open("D:\\Report Automation\\Login_details\\BW_Orange_ows_login.txt",'r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.maximize_window()

    driver.get(ows3Gurl)
    sl.sleep(5)
    driver.refresh()
    sl.sleep(3)

    attachment = '_uploadFile'
    x_arg = f'//form/input[contains(@name,"{attachment}")]'

    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    print(attachmentInput)
    attachmentInput.send_keys(file_3G)
    sl.sleep(5)

    file_path_1_array = file_3G.split('\\')
    file_path_1_title = file_path_1_array[len(file_path_1_array) - 1].split('.')[0]
    print(file_path_1_title)
    x_arg = f'//div[contains(@class,"filelist_cls")]/a[contains(text(),"{file_3G_name}")]'
    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))

    print("Uploading file attached")
    
    submit = 'ExcelImportPanel1_submit'
    x_arg = f'//div[contains(@id,"{submit}")]'
    submitBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    sl.sleep(5)

    submitBtn.click()

    importTb= 'ExcelImportPanel1_msg'
    x_arg = f'//div[contains(@id,"{importTb}")]'
    importBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))#

    sl.sleep(15)
    
    print("3G NUR Import Completed...")

def Upload4GNUR():

    ows4Gurl = "https://1at7-frapp.teleows.com/servicecreator/spl/FME_Performance_Dashboard/fpd_fme_performance_rating_import.spl"

    with open("D:\\Report Automation\\Login_details\\BW_Orange_ows_login.txt",'r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.maximize_window()

    driver.get(ows4Gurl)
    sl.sleep(5)
    driver.refresh()
    sl.sleep(3)

    attachment = '_uploadFile'
    x_arg = f'//form/input[contains(@name,"{attachment}")]'

    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    print(attachmentInput)
    attachmentInput.send_keys(file_4G)
    sl.sleep(5)

    file_path_1_array = file_4G.split('\\')
    file_path_1_title = file_path_1_array[len(file_path_1_array) - 1].split('.')[0]
    print(file_path_1_title)
    x_arg = f'//div[contains(@class,"filelist_cls")]/a[contains(text(),"{file_4G_name}")]'
    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))

    print("Uploading file attached")
    
    submit = 'ExcelImportPanel1_submit'
    x_arg = f'//div[contains(@id,"{submit}")]'
    submitBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    sl.sleep(5)

    submitBtn.click()

    importTb= 'ExcelImportPanel1_msg'
    x_arg = f'//div[contains(@id,"{importTb}")]'
    importBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))

    sl.sleep(15)
    
    print("4G NUR Import Completed...")

    driver.quit()



# checkold()
# getAttachInfo() 
ExtractData()
Upload2GNUR()
Upload3GNUR()
Upload4GNUR()
 
