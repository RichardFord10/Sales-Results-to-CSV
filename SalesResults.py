import openpyxl
import selenium
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from getpass import getpass
import pandas as pd
import sys
import csv
import os

# configure webdriver & headless chrome
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")
driver = webdriver.Chrome(options = chrome_options, executable_path=r'C:/Users/rford/Desktop/chromedriver/chromedriver.exe')

# current day format
currentDate = datetime.today().strftime('%Y-%m-%d')

#login function


def login(user, pword = str):
    driver.get("https://#######.com/manager")
    Username = driver.find_element_by_id("bvuser")
    Password = driver.find_element_by_id("bvpass")
    Login = driver.find_element_by_xpath('//*[@id="form1"]/div/div[2]/input')
    Username.send_keys(user)
    Password.send_keys(pword)
    Login.click()
    print("Logging In...")
 
   

#sale ids
everyone_ids = [
'84433',
'84454',
'85481',
'85484',
'86843',
'85488',
'85647',
'85821',
'86212',
'86214',
'86216',
'86399',
'87035',
'86400',
'86627',
'86461',
'86629',
'86638',
'86640',
'89814'
]


catalog_ids = [
'49260',
'85060',
]

acquisition_ids = [
'61736',
'61735',
'70278',
'70279',
'70280',
'70281',
'72551',
'72552',
'74968',
'74969',
'75729',
'70271',
'70274',
'70273',
'70272',
'70277',
'70275',
'70276',
'77250',
'77252',
'77254',
'60303',
'60305',
'60308',
'60311',
'60314',
'78640',
'78641',
'77248',
'81029',
'76403',
'76405',
'76407',
'76382',
'76384',
'76391',
'76392',
'76393',
'76398',
'76400',
'76402',
'80335',
'80127',
'80131',
'80132',
'80133',
'80134',
'80136',
'80137',
'80139',
'80337',
'79441',
'80334',
'80343',
'80339',
'80340',
'80346',
'80349',
'80336',
'88017',
'88018'
]


#Result Lists
everyone_results = []
catalog_results = []
acquisition_results = []


#get sales results based on sale type ('everyone', 'catalog', 'acquisition')

def get_sale_results(sale_type):   
    driver.get("https://########.com/manager/email_logs.php")
    login(input("Enter Username: "), getpass("Enter Password: "))
    currentDate = datetime.today().strftime('%Y-%m-%d') 
    if sale_type == "everyone":
        print('Gathering Email Sales Data...')
        for everyone_id in everyone_ids:                
            driver.find_element_by_xpath('/html/body/form[1]/div/select/option[@value={}]'.format(everyone_id)).click()
            driver.find_element_by_xpath('/html/body/form[1]/div/input[1]').submit()
            WebDriverWait(driver, 1)
            title = driver.find_element_by_xpath('/html/body/h2[2]').text
            sold = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[4]').text
            cost = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[6]').text[1:]
            price = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[7]').text[1:]
            profit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[8]').text[1:]
            addSales = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[9]').text[1:]
            addProfit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[10]').text[1:] 
            everyone_results.append([title, sold, cost, price, profit, addSales, addProfit, everyone_id, currentDate, sale_type])
            print([title, sold, cost, price, profit, addSales, addProfit, everyone_id, currentDate, sale_type])
        print("Email Data Gathered")
        with open(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_eve_results.csv', 'w+', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(everyone_results)
            print('Finished daily_everyone_results.csv')
    elif sale_type == "catalog":
        print('Gathering Catalog Sales Data...')
        for catalog_id in catalog_ids:
            driver.find_element_by_xpath('/html/body/form[2]/div/select/option[@value={}]'.format(catalog_id)).click()
            driver.find_element_by_xpath('/html/body/form[2]/div/input[1]').submit()
            WebDriverWait(driver, 1)
            title = driver.find_element_by_xpath('/html/body/h2[2]').text
            sold = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[4]').text
            cost = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[6]').text[1:]
            price = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[7]').text[1:]
            profit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[8]').text[1:]
            addSales = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[9]').text[1:]
            addProfit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[10]').text[1:] 
            catalog_results.append([title, sold, cost, price, profit, addSales, addProfit, catalog_id, currentDate, sale_type])
            print([title, sold, cost, price, profit, addSales, addProfit, catalog_id, currentDate, sale_type])
        print("Catalog Data Gathered")
        with open(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_cat_results.csv', 'w+', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(catalog_results)
            print('Finished daily_catalog_results.csv')
    elif sale_type == "acquisition":
        print('Gathering Acquisition Sales Data...')
        for acquisition_id in acquisition_ids:
            driver.find_element_by_xpath('/html/body/form[2]/div/select/option[@value={}]'.format(acquisition_id)).click()
            driver.find_element_by_xpath('/html/body/form[2]/div/input[1]').submit()
            WebDriverWait(driver, 1)
            title = driver.find_element_by_xpath('/html/body/h2[2]').text
            sold = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[4]').text
            cost = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[6]').text[1:]
            price = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[7]').text[1:]
            profit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[8]').text[1:]
            addSales = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[9]').text[1:]
            addProfit = driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[last()]/td[10]').text[1:] 
            acquisition_results.append([title, sold, cost, price, profit, addSales, addProfit, acquisition_id, currentDate, sale_type])
            print([title, sold, cost, price, profit, addSales, addProfit, acquisition_id, currentDate, sale_type])
        print("Acquisition Data Gathered")
        with open(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_acq_results.csv', 'w+', newline='', encoding="utf-8") as file:
            writer = csv.writer(file) 
            writer.writerows(acquisition_results)
            print('Finished daily_acquisition_results.csv')

# ↓ Covert CSV's to Excel Workbook Sheets ↓

def xlsx_report():
    with pd.ExcelWriter(r'C:\Users\rford\Desktop\Richard Ford\Results\results_{}.xlsx'.format(currentDate), engine="openpyxl", encoding="utf-8") as xlwriter:# pylint: disable=abstract-class-instantiated
        df = pd.read_csv(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_eve_results.csv')
        df.to_excel(xlwriter, sheet_name = 'Everyone', index = False)
        df2 = pd.read_csv(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_cat_results.csv')
        df2.to_excel(xlwriter, sheet_name = 'Catalog', index = False)
        df3 = pd.read_csv(r'C:\Users\rford\Desktop\Richard Ford\Results\daily_acq_results.csv')
        df3.to_excel(xlwriter, sheet_name = 'Acquisition', index = False)
        xlwriter.save()
        print('XLSX Complete')






get_sale_results('everyone')
get_sale_results('catalog')
get_sale_results('acquisition')
driver.quit()
xlsx_report()
 