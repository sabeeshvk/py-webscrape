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

# Save to DB
driver.quit()
