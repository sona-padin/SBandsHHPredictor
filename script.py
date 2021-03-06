import xlrd
import os, fnmatch
import datetime
import glob
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

#To be used for later when filling forms and creating a new folder
month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

#To store the names of the downloaded KML files. This is a work around to rename the files efficiently
old_files = []
#Create a new folder/directory for predicitions for today (Note: Since predicition files are a day ahead, I labeled folders to specify the actual day of the prediction)
today = datetime.datetime.now() + datetime.timedelta(days=1)
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
    element.send_keys('0')

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
    element.send_keys('6.0')

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
    #Rename KML file to the city which the prediction is based on
    kml_file_name =(element.get_attribute("href")).split('=')[1]
    old_file_name = path + '\\' + kml_file_name + '.kml'
    old_files.append(old_file_name)
    element.click() #Download file

time.sleep(5)

#End browser session and dispose chrome webdriver
driver.quit()

#Rename all files to the city in which the prediciton is based on and remove all yellow pins from KML file
for x, f in enumerate(old_files):
    location_name = sheet.cell_value((x+1), 0)
    new_file_name = path + '\\' + sheet.cell_value((x+1), 0)+ '.kml'
    try:
        os.rename(f, new_file_name)
        with open(new_file_name, "r") as file:
            lines = file.readlines()
            file.close()
        with open(new_file_name, "w") as file:
            for y, line in enumerate(lines):
                if(y == 4):
                    a = line.split("<description><![CDATA[Flight data for flight ")
                    b = a[1].split("<br>")
                    new_line = "<description><![CDATA[Flight data for flight " + sheet.cell_value((x+1), 0) + " <br>" + b[1]
                    file.write(new_line)
                elif(y == 183):
                    file.write('</Document></kml>')
                    file.close()
                    break
                else:
                    file.write(line)
    except FileNotFoundError:
        print("Re-run script and delete directory. All prediction files have not successfully downloaded. Note script will continue.")

print("DONE")
