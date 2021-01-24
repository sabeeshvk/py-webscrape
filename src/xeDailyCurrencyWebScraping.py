from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import pyodbc
import requests
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date 
import time
from sqlalchemy import create_engine
import urllib
driver = webdriver.Chrome(
    "C:/Users/h366319/AppData/Local/Programs/Python/chromedriver.exe")
currencyConvURL="https://www.xe.com/currencyconverter/convert/?Amount=1&From=CNY&To=USD"
driver.implicitly_wait(40)
driver.get(currencyConvURL)
currencyValClass=driver.find_element_by_class_name('converterresult-toAmount')
print(currencyValClass.text)
server = 'IE4LDTJ1M6N62.global.ds.honeywell.com' 
database = 'WebScrapping' 
username = 'sa' 
password = 'sseteam@2020' 
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()
cursor.execute("INSERT INTO WebScrapping.dbo.xe_ExchangeRates (ExchangeDate,CurrencyFrom,CurrencyTO,CurrencyValue) VALUES ('"+ str(date.today()) +"','CNY','USD','"+currencyValClass.text+"')")
cursor.commit()                
driver.quit()
