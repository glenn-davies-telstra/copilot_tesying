# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 18:34:29 2020

App to automate the handover email

@author: d284876
"""

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup #to parse to htmps
import xlsxwriter #to create the spreadsheets

from pathlib import Path
    
    
#Start by collecting the state data and saving to one drive for export to sharepoint
global homeDirectory
homeDirectory = str(Path.home())

#load chrome for rfq to enable downloads without confirmation, and load page
def loadBrowser():
    options = Options()
    options.headless = True
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
    })
    
    global browser
    browser = webdriver.Chrome(options=options, executable_path=r"C:\PIMS\chromedriver.exe")


    
#collect alll tge outage data from the outae board
def getOutageData():
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/")
    #wait for single signon to complete
    WebDriverWait(browser, 15).until(
        EC.visibility_of_element_located((By.XPATH, '/html/body/a/font'))
    )
    print("Outage page loaded, searching states...")
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=N01_HOTV&view=Active")
    time.sleep(1.5)
    global sourceHOTV
    sourceHOTV = browser.page_source
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=V01_VIC&view=Active")
    time.sleep(1.5)
    global sourceLOTV
    sourceLOTV = browser.page_source
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=Q01_CHTE&view=Active")
    time.sleep(1.5)
    global sourceCHTE
    sourceCHTE = browser.page_source    
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=Q02_SOPE&view=Active")
    time.sleep(1.5)
    global sourceSOPE
    sourceSOPE = browser.page_source     
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=S01_FLNH&view=Active")
    time.sleep(1.5)
    global sourceFLNH
    sourceFLNH = browser.page_source        
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=W01_PIEH&view=Active")
    time.sleep(1.5)
    global sourcePIEH
    sourcePIEH = browser.page_source         
    print("Search complete")
  






def getTestData():
    browser.get("http://hangar.in.telstra.com.au/hfc_wizard/")
    #wait for single signon to complete
    WebDriverWait(browser, 15).until(
        EC.visibility_of_element_located((By.XPATH, '/html/body/a/font'))
    )   
    #browser.get("http://hangar.in.telstra.com.au/hfc_wizard/wizard.php?state=Q02_SOPE&view=Active")
    time.sleep(1.5)
    global sourceSOPE
    sourceSOPE = browser.page_source     


#collect the outage table
def parseTable(source):
    soup = BeautifulSoup(source, 'html.parser')
    data = []
    #print(soup)
    table = soup.find('table')
                      #, attrs={'class':'lineItemsTable'})
    #print("********************")
    #print(table)
    table_body = table.find('tbody')
    rows = table_body.findAll('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele]) # Get rid of empty values
    #print("********************")
    #print(data)
    #print(soup.prettify())
    #print("********************")
    #list(soup.children)
    return data

#convert to csv file and save to....
def saveToXlsx(source):
    folder = '\\Telstra\\HFC NBOC Team - Handover\\'
    path = homeDirectory + folder
    if number == 1:
        file = 'HOTV\HOTV'
    elif number == 2:
        file = 'LOTV\LOTV';
    elif number == 3:
        file = 'CHTE\CHTE';     
    elif number == 4:
        file = 'SOPE\SOPE';     
    elif number == 5:
        file = 'FLNH\FLNH';     
    elif number == 6:
        file = 'PIEH\PIEH';             
    ext = '.xlsx'
    location = path + file + ext
    #print(location)
    with xlsxwriter.Workbook(location) as workbook:
        worksheet = workbook.add_worksheet()
        i = len(source)
        #worksheet.autofilter(0,0,i-1,6)
        worksheet.add_table(0,0,i-1,6, {'data': source, 'header_row': 0 })
 #{'data': source, 'header_row': False })
        
        #for row_num, data in enumerate(source):
            #worksheet.write_row(row_num, 0, data)
        #worksheet.autofilter(0,0,i-1,6)
    
def main():
    loadBrowser()
    #loadOutageBoard()
    #getNSW()
    getOutageData()
    browser.close()
    global number
    number = 1
    print("Saving Excel spreadsheets....")
    HOTV = parseTable(sourceHOTV)
    saveToXlsx(HOTV)
    
    number =number+1
    LOTV = parseTable(sourceLOTV)
    saveToXlsx(LOTV)
    
    number =number+1
    CHTE = parseTable(sourceCHTE)
    saveToXlsx(CHTE)
    
    number =number+1
    SOPE = parseTable(sourceSOPE)
    saveToXlsx(SOPE)
    
    number =number+1
    FLNH = parseTable(sourceFLNH)
    saveToXlsx(FLNH)
    
    number =number+1
    PIEH = parseTable(sourcePIEH)
    saveToXlsx(PIEH)

    print("Operations complete, exe closing.")
    
    pass
    
main()