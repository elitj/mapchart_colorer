"""
Created on Tue Oct 15 15:17:09 2019

@author: Eli Johnson
"""

# A word of caution when entering counties: in Virginia there are 3 independent
# cities with the same names as distinct counties: Fairfax, Richmond and 
# Roanoke.  Bedford also shares its name with a county and was previously
# independent.  Take care that the way these are handled in your data is the
# same way they are handled by mapchart.com (or just manually correct them).

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException as WDE, StaleElementReferenceException as SERE, NoSuchElementException as NSEE, TimeoutException#, ElementNotInteractableException
from pyexcel_ods import get_data

stateDict = {
"Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR",
"California":"CA","Colorado":"CO","Connecticut":"CT","Delaware":"DE",
"Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID","Illinois":"IL",
"Indiana":"IN","Iowa":"IA","Kansas":"KS","Kentucky":"KY","Louisiana":"LA",
"Maine":"ME","Maryland":"MD","Massachusetts":"MA","Michigan":"MI",
"Minnesota":"MN","Mississippi":"MS","Missouri":"MO","Montana":"MT",
"Nebraska":"NE","Nevada":"NV","New Hampshire":"NH","New Jersey":"NJ",
"New Mexico":"NM","New York":"NY","North Carolina":"NC","North Dakota":"ND",
"Ohio":"OH","Oklahoma":"OK","Oregon":"OR","Pennsylvania":"PA",
"Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD","Tennessee":"TN",
"Texas":"TX","Utah":"UT","Vermont":"VT","Virginia":"VA","Washington":"WA",
"West Virginia":"WV","Wisconsin":"WI","Wyoming":"WY",
"District of Columbia":"DC"
}

COLOR_PICK_ATTEMPTS = 3

def scroll_shim(passed_in_driver, object): #created by user Cynic on Stackoverflow
        x = object.location['x']
        y = object.location['y']
        scroll_by_coord = 'window.scrollTo(%s,%s);' % (
            x,
            y
        )
        scroll_nav_out_of_way = 'window.scrollBy(0, -120);'
        passed_in_driver.execute_script(scroll_by_coord)
        passed_in_driver.execute_script(scroll_nav_out_of_way)
        
def hex_to_rgb(hexa):
    return str(tuple(int(hexa[i:i+2], 16) for i in (1, 3, 5)))

def color_pick(hex_color): #argument is hexadecimal string of the form '#xxxxxx'
    try:
        colorchoose = browser.find_element_by_class_name('sp-preview-inner')
        scroll_shim(browser,colorchoose)
        colorchoose.click()
        if browser.find_element_by_class_name('sp-palette-toggle').text == "more":
            browser.find_element_by_class_name('sp-palette-toggle').click()
        colorbox = browser.find_element_by_class_name('sp-input')
        colorbox.click()
        colorbox.clear()
        colorbox.send_keys(hex_color)
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "sp-choose")))
        browser.find_element_by_class_name('sp-choose').click()
        try:
            assert browser.find_element_by_class_name('sp-preview-inner').get_attribute("style")=="background-color: rgb{};".format(hex_to_rgb(hex_color))
        except AssertionError:
            global COLOR_PICK_ATTEMPTS
            print("Color ",hex_color, "which is ",hex_to_rgb(hex_color), " does not match style ", browser.find_element_by_class_name('sp-preview-inner').get_attribute("style"))
            if COLOR_PICK_ATTEMPTS > 0:
                COLOR_PICK_ATTEMPTS -= 1
                color_pick(hex_color)
            else:
                print("Too many failures.  Exiting")
                browser.quit()
                exit()
    except Exception as ex:
        print("Error while picking color ", hex_color)
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)
        browser.quit()
        exit()

#    color_counties is passed an .ods sheet with at least 3 columns containing 
#    county names, state names, and numeric data separating counties by desired
#    color.  In order to be colored, the county's data column must be between
#    min_val (inclusive) and max_val (exclusive).  Defaults to expecting 
#    unabbreviated state names, otherwise set state_is_abbreviated to True.
def color_counties(ods_county_sheet,min_val,max_val,county_column=0,state_column=1,data_column=2,state_is_abbreviated = False):
    element = browser.find_element_by_class_name('select2-selection__rendered')        
    scroll_shim(browser,element)
    for entry in sheet:
        if entry[data_column]>=min_val and entry[data_column]<max_val:
            if state_is_abbreviated: 
                state = entry[state_column]
            else: 
                state = stateDict[entry[state_column]]
            county = entry[county_column]
            text_to_enter = county+" ("+state+")"
            element = browser.find_element_by_class_name('select2-selection__rendered')
            element.click()
            try:
                element2 = browser.find_element_by_class_name('select2-search__field')
            except NSEE:
                element.click()
                element2 = browser.find_element_by_class_name('select2-search__field')
            element2.clear()
            element2.send_keys(text_to_enter)
            element2.send_keys(Keys.RETURN)
            try:
                assert browser.find_element_by_class_name('select2-selection__rendered').get_attribute("title")==text_to_enter, "foo"
                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.ID, "colorFromSearch")))
                browser.find_element_by_id('colorFromSearch').click()
            except AssertionError:
                print("'{}' did not get entered.  Usually a spelling discrepancy.".format(text_to_enter))
            except SERE:
                print("Error while entering {}, (stale element)".format(text_to_enter))
                print("This most commonly occurs when the string to enter is a substring of another county.")
                #browser.close()
            except TimeoutException:
                print("Error while entering {}, {} (timeout waiting for element)".format(text_to_enter))
    

# color_map takes two lists, a list of hex color strings of the form '#xxxxxx',
# and a descending list of numbers (your data) that mark the boundaries between
# each color.  Note that threshold_list should be longer than hex_color_list by 
# 1 to account for the upper and lower bounds of data.  Anything above or below
# these numbers will go uncolored.
# The other arguments are passed to color_counties and explained there.
def color_map(hex_color_list,threshold_list,sheet,cty_col=0,st_col=1,data_col=2,is_abbrev=False):
    assert len(hex_color_list)+1==len(threshold_list),"Threshold list should be 1 longer than color list."
    for i in range(len(hex_color_list)):
        color_pick(hex_color_list[i])
        color_counties(sheet,threshold_list[i+1],threshold_list[i],cty_col,st_col,data_col,is_abbrev)
        print("Finished with color ",hex_color_list[i])
    if input("Enter any key to close (Note: this will close the map window):")!="":
        browser.quit()
        

# Sample color palette:    
red6 = '#990e0e' #x<-10000
red5 = '#cf0e22' #-10000<x<-5000
red4 = '#fb3343' #-5000<x<-1000
red3 = '#f9897f' #-1000<x<-500
red2 = '#fdb6b6' #-500<x<-200
red1 = '#ffdddd' #-200<x<0
blue1 = '#e5f5f9' #0<x<200
blue2 = '#c7eae5' #200<x<500
blue3 = '#92c5de' #500<x<1000
blue4 = '#2585c4' #1000<x<5000
blue5 = '#1555ac' #5000<x<10000
blue6 = '#023061' #10000<x

# Sample color list and threshold list:
colors = [blue6,blue5,blue4,blue3,blue2,blue1,red1,red2,red3,red4,red5,red6]
borders = [2000000,10000,5000,1800,800,300,0,-300,-800,-1800,-5000,-10000,-20000]


# Edit .ods path and sheet name below:
#data = get_data("C:/path/to/your_data.ods")
#sheet = data["your_sheet_name"]
data = get_data("C:/Users/User/Documents/States/by_trend.ods")
sheet = data["by_trend"]
sheet.pop(0)                       #Comment out if first row is not headers

try:
    browser = webdriver.Firefox()
    browser.get("https://mapchart.net/usa-counties.html")
except WDE:
    print("Unable to open Firefox")
    browser.quit()
    exit()
    
color_map(colors,borders,sheet,0,1,18)
