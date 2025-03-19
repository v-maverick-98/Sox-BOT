import datetime as dt
import time

stTime = time.time()

q1 = ["Nov","Dec","Jan"]
q2 = ["Feb","Mar","Apr"]
q3 = ["May","Jun","Jul"]
q4 = ["Aug","Sep","Oct"]

# get the current month
# todays_date = dt.date.today()
todays_date = dt.date(2024,11,1) # Date to test on Q1
Year = todays_date.strftime("%y")
print(type(Year))
todays_month = todays_date.strftime("%b")
if todays_month == ("Nov" or "Dec"):
    Year = int(Year) + 1
    Year = str(Year)
print(type(Year))
fiscalYear = "FY" + Year
print(fiscalYear)
if todays_month in q1:
    crrntQ = "Q1"
elif todays_month in q2:
    crrntQ = "Q2"
elif todays_month in q3:
    crrntQ = "Q3"
elif todays_month in q4:
    crrntQ = "Q4"

print(todays_month)
print(crrntQ)

import os
import os.path
import sys

xlname = "Workday " + crrntQ + fiscalYear + " PLR Evidence"
foldername = crrntQ + fiscalYear
folderpath = r"C:\Users\sandovfa\OneDrive - Hewlett Packard Enterprise\Documents\WD Audit SOX"
subfolderpath = os.path.join(folderpath,foldername)


if not os.path.exists(subfolderpath):
    os.mkdir(subfolderpath)
    print(f"Directory '{foldername}' created successfully.")
else:
    # FileExistsError:
    print(f"Directory '{foldername}' already exists.")
# except PermissionError:
#     # print(f"Permission denied: Unable to create '{foldername}'.")
#     sys.exit()
# except Exception as e:
#     sys.exit()

xlfilename = xlname + ".xlsx"
xlpathname = os.path.join(subfolderpath,xlfilename)
print(xlpathname)

import openpyxl

sheet_to_use = "Audit Logs"
# create a new workbook
new_book = openpyxl.Workbook()
new_book.worksheets[0].title = sheet_to_use
new_book.save(xlpathname)
new_book.close()


reportNames = ["COMP202 - Rewards Planning Summary",
    "COMP202 - Rewards Planning Summary w/Holdbacks",
    "COMP55a - Base Salary Change Spend Detail",
    "COMP201 - Rewards Planning Detail",
    "COMP201 - Rewards Planning - Bonus Plan Details",
    "COMP01 - Bonus Payments",
    "COMP50 - Historical Stock Grants"]

"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP202+-+Rewards+Planning+Summary&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP202+-+Rewards+Planning+Summary+w%2FHoldbacks&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP55a+-+Base+Salary+Change+Spend+Detail&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP201+-+Rewards+Planning+Detail&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP201+-+Rewards+Planning+-+Bonus+Plan+Details&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP01+-+Bonus+Payments&state=searchCategory-all%3Adefault"
"https://wd5.myworkday.com/hpe/d/search.htmld?q=COMP50+-+Historical+Stock+Grants&state=searchCategory-all%3Adefault"

rngQ1start = ["11/01/20" + str(int(Year) - 1),
    "01/31/20" + Year]
rngQ2start = ["02/01/20" + Year,
    "04/30/20" + Year]
rngQ3start = ["05/01/20" + Year,
    "07/31/20" + Year]
rngQ4start = ["08/01/20" + Year,
    "10/31/20" + Year]

rngQsList = (rngQ1start,rngQ2start,rngQ3start,rngQ4start)

listindex = crrntQ[1:]
listindex = int(listindex) - 1
dts = rngQsList[listindex]

import glob

# files = glob.glob(r"C:\Users\sandovfa\OneDrive - Hewlett Packard Enterprise\Documents\WD Audit SOX\*.xlsx",recursive=False)

# shot = pyautogui.screenshot()
# hotkey = 'win+prtscn' # Use this combination anytime while script is running

import pyautogui as pyag
import time
import keyboard
import re
import threading
from python_goto import goto
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains 

driver = webdriver.Chrome()
driver.set_page_load_timeout(5)
# driver.get("https://www.google.com.mx")
wdURL = driver.get("https://wd5.myworkday.com/hpe/d/home.htmld")
# https://wd5.myworkday.com/hpe/d/home.htmld
# https://wd5-impl.workday.com/hpe/d/home.htmld
driver.maximize_window()
main_window = driver.current_window_handle

wait = WebDriverWait(driver, 5)

driver.implicitly_wait(5)

time.sleep(3)

pyag.leftClick(x=960, y=756)

time.sleep(3)

pyag.press('enter')
time.sleep(3)
pyag.write('19910808')
pyag.press('enter')
time.sleep(5)

from PIL import ImageGrab
import re

def SStake(fpath):
    screenShot = ImageGrab.grab()
    screenShot.save(fpath)
    screenShot.close()

pictures = []

for report in reportNames:
        
    Action = ActionChains(driver)

    time.sleep(2)
    driver.switch_to.window(main_window)
    driver.maximize_window()
    driver.implicitly_wait(3)
    elementLogo = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/div[6]/div[1]/div[2]/button")))
    # elementLogo = pyag.click()
    # //*[@id='homeButtonContainer']/button/div/img
    time.sleep(2)

    elementSrch = wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='wd-searchInput']")))
    # elementSrch = driver.find_element(By.XPATH,"//*[@id='wd-searchInput']")
    Action.click(elementSrch).perform()
    time.sleep(1)
    pyag.write(str(report),0.02)
    time.sleep(2)
    pyag.press('enter')
    time.sleep(3)

    reportNm_button = driver.find_element(By.LINK_TEXT, report)
    # /html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div/div/div/div[2]/div/div/div/ol/li[1]/div/ol/li[1]/div/div/div/h3/button
    time.sleep(2)
    pyag.press("tab")
    pyag.press("tab")
    pyag.press("tab")
    Action.move_to_element(reportNm_button).perform()
    time.sleep(2)
    # reportNm_button.click()
    threedots_button = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div/div/div/div[2]/div/div/div/ol/li[1]/div/ol/li[1]/div/div/div/h3/button")
    # threebuttons_enabled =threedots_button.is_enabled()
    # Action.move_to_element(threedots_button).perform()
    time.sleep(1)
    Action.click(threedots_button).perform()
    time.sleep(2)
    try:
        audit_button = driver.find_element(By.XPATH, '/html/body/div[10]/div[3]/div[2]/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div')
        audit_button.click()
        time.sleep(2)
    except:
        audit_button2 = driver.find_element(By.XPATH, '/html/body/div[11]/div[3]/div[2]/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div')
        audit_button2.click()
        time.sleep(2)

    auditWindow = wait.until(EC.invisibility_of_element((By.XPATH ,"/html/body/div[12]/div")))

    fromdt = dts[0]
    todt = dts[1]

    fromMomentDt = driver.find_element(By.XPATH, "/html/body/div[11]/div/div[2]/div/div[2]/div/div/ul/li[1]/div[2]/div/div/div[1]/div[2]/div[1]/div[1]/input")
    fromMomentDt.click()
    pyag.write(fromdt)
    time.sleep(1.5)
    toMomentDT = driver.find_element(By.XPATH, "/html/body/div[11]/div/div[2]/div/div[2]/div/div/ul/li[2]/div[2]/div/div/div[1]/div[2]/div[1]/div[1]/input")
    toMomentDT.click()
    pyag.write(todt)
    time.sleep(2)
    ok_Button = driver.find_element(By.XPATH, "/html/body/div[11]/div/div[3]/div/div/div[1]/div[1]/button[1]")
    ok_Button.click()
    time.sleep(2)

    noItems = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[1]/div[1]/label")
    
    if noItems:
        noRows = r'^\d{1,3}'
        rws = re.search(noRows,noItems.text)
        rws = rws.group()# noItems.text
        rws = int(rws)
        print(noItems.text)
        print(rws)
        print(type(rws))
        
        if rws == 2 :
            ssname = report + ".png"
            pngpathname = os.path.join(subfolderpath,ssname)
            print(pngpathname)
            SStake(pngpathname)

        elif rws >= 2 and rws <= 5:
            ssname = "00_" + report + ".png"
            pngpathname = os.path.join(subfolderpath,ssname)
            print(pngpathname)
            SStake(pngpathname)
            pyag.press("pagedown")  # Scroll down
            ssname = "01_" + report + ".png"    
            pngpathname = os.path.join(subfolderpath,ssname)
            print(pngpathname)
            SStake(pngpathname)

        elif rws >= 6:
            noScreenshots = round(rws / 6)
            rw = 1
            ssname = "00_" + report + ".png"
            pngpathname = os.path.join(subfolderpath,ssname)
            print(pngpathname)
            SStake(pngpathname)
            pyag.press("pagedown")  # Scroll down
            time.sleep(2)
            
            for count , screenshots in enumerate(noScreenshots,start=0):     
                ssname = count + "_" + report + ".png"
                pngpathname = os.path.join(subfolderpath,ssname)
                print(pngpathname)
                
                SStake(pngpathname)

                # iterar el número de renglón
                # xprw = "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/div/div/div[1]/div/table/tbody/tr[" + rw + "]"
                # tableRow = driver.find_element(By.XPATH, xprw) # Table
                tableRow = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/div/div/div[1]/div/table/tbody/tr[1]") # Table
                tableRow.click()
                pyag.scroll(5)
                # rw += 4 
                

    else:
        SStake(pngpathname)
        ssname =  + report + ".png"
        pngpathname = os.path.join(subfolderpath,ssname)
        print(pngpathname)

    # ssname = report + ".png"
    # pngpathname = os.path.join(folderpath,ssname)
    # print(pngpathname)

    # screenShot = ImageGrab.grab()
    # screenShot.save(pngpathname)
    # screenShot.close()

    # image1 = pyag.screenshot(report + ".png")
    # timestamp = time.strftime("%Y%m%d_%H%M%S")
    # filename = f"screenshot_{timestamp}.png"
    # pyag.KEYBOARD_KEYS
    # /html/body/div[11]/div[3]/div[2]/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div


# scroll 4 times 
# /html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/div/div/div[1]
# /html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[1]/div[1] label No of items
# //*[@id="riva-grid-api-key-uid73"]/div/div/div/div[1]

# button7location = pyautogui.locateOnScreen('calc7key.png')
# button7point = pyautogui.center(button7location)
# pyautogui.click('calc7key.png')