import xlrd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

#Get date ahead (currentdate + 1)
next_day = datetime.datetime.today() + datetime.timedelta(days=1) #Gives you the year, month, day, hour, and minute

#Let's open up the excel file to have access to the different values needed for the tracker
loc = ('Flight Path Predictions.xlsx')
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0) #
num = sheet.cell_value(1,11)
sheet_names = wb.sheet_names()
driver = webdriver.Chrome(executable_path= 'C:/Users/sonap/Downloads/chromedriver_win32/chromedriver.exe')
driver.get('http://predict.habhub.org/')
driver.implicitly_wait(10)

##############################LAUNCH SITE####################################
element = Select(driver.find_element_by_id('site'))
element.select_by_visible_text('Other')

##############################LATITUDE/LONGITUDE#############################
#******Latitude
element = driver.find_element_by_id('lat')
element.clear() #get rid of pre-existing input
num = sheet.cell_value(1, 11)
element.send_keys(str(num))
#******Longitude
element = driver.find_element_by_id('lon')
element.clear() #get rid of pre-existing input
num = sheet.cell_value(1, 12)
element.send_keys(str(num))

##############################LAUNCH ALTITUDE#################################
element = driver.find_element_by_id('initial_alt')
element.clear() #get rid of pre-existing input
element.send_keys('100000')

#############################LAUNCH TIME########################################
#This needs to be synchronized with the location and time (implement later)
#******Hour
element = driver.find_element_by_id('hour')
element.clear() #get rid of pre-existing input
element.send_keys('12')
#******Min
element = driver.find_element_by_id('min')
element.clear() #get rid of pre-existing input
element.send_keys('00')

#############################LAUNCH DATE########################################
#******Day
element = driver.find_element_by_id('day')
element.clear() #get rid of pre-existing input
element.send_keys(next_day.day)
#******Month
element = Select(driver.find_element_by_id('month'))
element.select_by_visible_text(month[next_day.month-1])
#******Year
element = driver.find_element_by_id('year')
element.clear() #get rid of pre-existing input
element.send_keys(next_day.year)

##############################ASCENT RATE#######################################
element = driver.find_element_by_id('ascent')
element.clear() #get rid of pre-existing input
element.send_keys('9.5')

##############################BURST ALTITUDE####################################
element = driver.find_element_by_id('burst')
element.clear() #get rid of pre-existing input
element.send_keys('33000')

###############################DESCENT RATE#####################################
eleme = driver.find_element_by_id('drag')
eleme.clear()
eleme.send_keys('6.0')

#SUBMIT QUERY##
element = driver.find_element_by_id("run_pred_btn")
element.click()

#Download the KML file to Open in Google Earth
element = driver.find_element_by_link_text("KML")
file = element.click()
