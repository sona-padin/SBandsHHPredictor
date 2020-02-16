import xlrd
import os
import datetime
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

#To be used for later when filling forms and creating a new folder
month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

#Create a new folder/directory for predicitions for today
today = datetime.datetime.now()
new_folder_name = month[today.month -1] + str(today.day) + str(today.year)
path = 'C:\\ShadowBandsResearch\\' + new_folder_name

#Try to create a directory
try:
    os.mkdir(path)
except FileExistsError:
    print("Caution...directory exists.")

#Get date ahead (currentdate + 1)
next_day = datetime.datetime.today() + datetime.timedelta(days=1) #Gives you the year, month, day, hour, and minute

#Let's open up the excel file to have access to the different values needed for the tracker
loc = ('Flight Path Predictions.xlsx')
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
num = sheet.cell_value(1,11)

#Specify download path for the KML files
options = webdriver.ChromeOptions()
prefs = {"download.default_directory" : path}
options.add_experimental_option("prefs", prefs)

#Start instance of Chrome browsing with configuration specified above
driver = webdriver.Chrome(executable_path='C:/Users/sonap/Downloads/chromedriver_win32/chromedriver.exe', options=options)
driver.get('http://predict.habhub.org/')
driver.implicitly_wait(10)

#Run predictions for each location on the eclipse path
for x in range(1,13):
    ##############################LAUNCH SITE####################################
    element = Select(driver.find_element_by_id('site'))
    element.select_by_visible_text('Other')

    ##############################LATITUDE/LONGITUDE#############################
    #******Latitude
    element = driver.find_element_by_id('lat')
    element.clear() #get rid of pre-existing input
    num = sheet.cell_value(x, 11)
    element.send_keys(str(num))
    #******Longitude
    element = driver.find_element_by_id('lon')
    element.clear() #get rid of pre-existing input
    num = sheet.cell_value(x, 12)
    element.send_keys(str(num))

    ##############################LAUNCH ALTITUDE#################################
    element = driver.find_element_by_id('initial_alt')
    element.clear() #get rid of pre-existing input
    element.send_keys('100000')

    #############################LAUNCH TIME########################################
    #Read local time from datasheet and tokenize it
    time_from_sheet = xlrd.xldate_as_tuple(sheet.cell_value(x, 5), 0)
    #******Hour
    element = driver.find_element_by_id('hour')
    element.clear() #get rid of pre-existing input
    element.send_keys(time_from_sheet[3])
    #******Min
    element = driver.find_element_by_id('min')
    element.clear() #get rid of pre-existing input
    element.send_keys(time_from_sheet[4])

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
    element.click()

#End browser session and dispose chrome webdriver
driver.quit()

#Remove all yellow pins from each of the .kml files
files = [f for f in glob.glob(path + "**/*.kml", recursive=True)]

for f in files:
    with open(f, "r") as file:
        lines = file.readlines()
        file.close()
    with open(f, "w") as file:
        for x, line in enumerate(lines):
            if(x == 78):
                file.write('</Document></kml>')
                file.close()
            if(x > 78):
                continue
            else:
                file.write(line)
