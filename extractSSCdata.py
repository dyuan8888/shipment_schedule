# -*- coding: utf-8 -*-
"""
Created on Thu May 16 17:12:11 2019

@author: DanielYuan
"""

from selenium import webdriver
from time import sleep
import requests

browser = webdriver.Firefox()
browser.get('http://192.168.0.104/ssc_query/index.php/login/login.html')
browser.add_cookie({'name':'PHPSESSID', 'value': 'g5gpm2975lk1lsmvkadtmuffo2'})
sleep(3)
browser.refresh()

sleep(2)

browser.find_element_by_css_selector(
        '#Form > form:nth-child(1) > p:nth-child(2) > input:nth-child(2)'
    ).send_keys('danielyuan')

sleep(2)

browser.find_element_by_css_selector("#Form > form:nth-child(1) > \
                                p:nth-child(3) > input:nth-child(2)").send_keys('1234')

sleep(2)

browser.find_element_by_css_selector('.but_ie').click()

sleep(2)

browser.find_element_by_css_selector('body > div:nth-child(4) > \
                                    input[type=button]').click()


'''
page = browser.page_source
#url = 'http://192.168.0.104/ssc_query/index.php/Index/production_schedule.html'
#s r = requests.get(url)
bs= BeautifulSoup(page,'html.parser')
#bs.body
trs = bs.find('tbody').findAll('tr')
project_id = []
for i in trs:
    #id = i.find('td').text
    
    tds = i.findAll('td')
    dict = {}
    dict['Project ID'] = tds[0].text
    dict['Project Name'] = tds[2].text
    dict['Manual Completion'] = tds[4].text
    dict['Ship Date'] = tds[3].text
    #dict
    project_id.append(dict)
    
df = pd.DataFrame(project_id).iloc[1:,:]
df = df.set_index('Ship Date', drop=True)
df.head(30)
'''
