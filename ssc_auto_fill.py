# -*- coding: utf-8 -*-
"""
Updated on Monday Jan 20 12:40 2020

What's new:
    1. Modified TSV200E CSS selector
    2. Updated line 127 by adding " 'HD' not in projinfo"
    3. Added MOCVD remark to show MOCVD type
    
    
@author: DanielYuan
"""

# Navigate to the Production Schedule page of the SSC Online Management System

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup


def ssc_fill(cmp_data):
    #cmp_data = {'31095': ['1-PM MOCVD A7 SYS', '2019-07-22']}
    '''Access to the SSC page, extract the SSC shipment data, and fill data'''    
    url = 'http://192.168.0.104/ssc_query/index.php/login/login.html'
    browser = getUrl(url, username='danielyuan', password='1234')
    ssc_dict = getSSC_dict(browser)
    fill_data(browser, cmp_data, ssc_dict)
        

def getUrl(url, username, password):
    '''Open the SSC page and navigate to the Production Schedule page'''  
    options = webdriver.FirefoxOptions()
    options.add_argument('-headless') 
    options.add_argument('--disable-gpu')
    browser = webdriver.Firefox(options=options)
    browser.get(url)
    browser.find_element_by_name('username').send_keys(username)
    browser.find_element_by_name('password').send_keys(password)
    browser.find_element_by_class_name('but_ie').click()
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'input'))).click()
    return browser


def getSSC_dict(browser):
    '''Get the Project IDs, ship ids and Ship Dates from the SSC and store them in a dictionary'''
    page = browser.page_source
    bs= BeautifulSoup(page,'html.parser')
    trs = bs.find('tbody').findAll('tr')    
    ssc_dict = {}
    for tr in trs:
        tds = tr.findAll('td')       
        try:
            ssc_dict[tds[0].text.split('/')[0]] = [tr['id'], tds[3].text]  # Get each tr's id        
        except KeyError:
            pass         
    return ssc_dict
        

def fill_data(browser, cmp_data, ssc_dict):
    '''Compare data and fill the SSC'''    
    for j, k in cmp_data.items():
        if j in ssc_dict.keys(): 
            if k[1] != ssc_dict[j][1]:  
                update_SSC(browser, k[1], ssc_dict[j][0])  # Do SSC update        
                print(f'\n{j} ship date was updated to {k[1]} in SSC Online Management System!')
        else:
            create_SSC(browser, j, k[0], k[1])   # Do SSC create
            print(f'\n{j} was created in SSC Online Management System!')
    browser.close()
    print('\n\nDone with the SSC data auto-filling!')

    
def update_SSC(browser, ship_date, ship_id):
    '''Update the SSC Online Management System if the ship date changes'''
    try:
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#{ship_id} > td:nth-child(11) > span:nth-child(1) > img:nth-child(1)'))).click()
    	#browser.find_element_by_css_selector(f'#{ship_id} > td:nth-child(11) > span:nth-child(1) > img:nth-child(1)').click()
    except:
    	print(f'{ship_id} failed to be found!')
    	pass
    try:
        browser.find_element_by_class_name('laydate-icon').clear()
    except:
        print(f'{ship_id} field failed to clear!')
        pass
    try:
        browser.find_element_by_class_name('laydate-icon').send_keys(ship_date)
    except:
        print(f'{ship_id} field failed to enter!')
        pass

    browser.find_element_by_css_selector('#wrap > form:nth-child(1) > input:nth-child(6)').click()
    browser.switch_to.alert.accept()
    browser.implicitly_wait(20)
    
    
def create_SSC(browser, project_id, projInfo, ship_date):
    '''Create a new shipment on the SSC Online Management System'''
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ssc_schedule > span:nth-child(2) > input:nth-child(1)'))).click()
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.NAME, 'SHIPMENT_NO'))).send_keys(project_id)
    browser.find_element_by_class_name('laydate-icon').send_keys(ship_date)
    # WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.NAME, 'REMARK'))).send_keys(projInfo)
    
    ownerOption = browser.find_element_by_name('owner')
    if project_id[-1] in ['3', '5', '6', '8']:
        ownerOption.find_element_by_xpath('/html/body/div[3]/form/ul[1]/ol[5]/li[2]/select/option[4]').click()
    else:
        ownerOption.find_element_by_xpath('/html/body/div[3]/form/ul[1]/ol[5]/li[2]/select/option[2]').click()
    
    prodType = browser.find_element_by_name('ProductType') #  make a product type selection
    if 'MOCVD' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(9)'
        ).click()
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.NAME, 'REMARK'))).send_keys(projInfo)
    elif 'TSV200' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(7)'
        ).click()
    elif 'TSV300' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(8)'
        ).click()
    elif 'AD-RIE' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(2)'
        ).click()
    elif 'D-RIE' in projInfo and 'HD' not in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(3)'
        ).click()
    elif 'HD-RIE' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(10)'
        ).click()
    elif 'SD-RIE' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(11)'
        ).click()
    elif 'DSC ICP' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(12)'
        ) .click()
    elif 'SSC ICP' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(6)'
        ).click()   
    elif 'SSC AD-RIE' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(4)'
        ).click()  
    elif 'iDEA' in projInfo:
        prodType.find_element_by_css_selector(
            '#wrap > form:nth-child(1) > ul:nth-child(1) > ol:nth-child(4) > li:nth-child(2) > select:nth-child(1) > option:nth-child(5)'
        ).click()           
    
    browser.find_element_by_css_selector(
        '#wrap > form:nth-child(1) > input:nth-child(6)'
    ).click()
    browser.switch_to.alert.accept() # get dialog box
    browser.implicitly_wait(20)
    
    
