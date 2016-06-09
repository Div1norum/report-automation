# Tableau Virtual Training Reports Automation Script

# Author: Elliott Stam (elliott.stam@interworks.com)
# Version 1.2 -- April 10, 2016
# IW Team: Cascade

# Requirements: python 2.7, firefox browser
# Requirements: selenium packages, pandas
# Recommended: Anaconda (nice for step-by-step debugging & practice)

# Imports
from selenium import webdriver
import re
import time
import pandas as pd
from pandas import DataFrame
import glob
import os

# ****************************
# ***** DEFINE FUNCTIONS *****
# ****************************

# Define the function to obtain user input

def userCourse():
    courseNum = raw_input("Enter the number for your course: ")
    try:
        int(courseNum)
    except:
        print("Value entered is not valid.\n")
        return userCourse()
    else:
        if int(courseNum) > 0 and int(courseNum) < 7:
            return int(courseNum)
        else:
            print("Value entered is not valid.\n")
            userCourse()

# Define function to navigate to reports page and select date range

def goReports():
    driver.get("https://tableaulive.webex.com/mw3000/mywebex/report/usagereport.do?action=first&reportType=UsageReport&siteurl=tableaulive")
    time.sleep(0.5)
    
    # Add time selection here...
    
    driver.find_element_by_name("display").click()
    time.sleep(0.5)

# Define the course link text based on user's course input
def courseText(course):
    if course == 1:
        return "Tableau Desktop Fundamentals AM"
    elif course == 2:
        return "Tableau Desktop Fundamentals PM"
    elif course == 3:
        return "Tableau Desktop Advanced AM"
    elif course == 4:
        return "Tableau Desktop Advanced PM"
    elif course == 5:
        return "Tableau Desktop Server Administration AM"
    elif course == 6:
        return "Tableau Desktop Server Administration PM"

# Define function to get confIDs...

def getConfids():
    
    confids = []

    for link in links:
        href = link.get_attribute('href')
        pattern = re.compile(r'(\d{10})')
        confid = pattern.findall(href)
        confids.append(confid)
        
    newconfids = []

    for i in confids:
        for x in i:
            newconfid = x.encode("utf-8")
            newconfids.append(newconfid)
            
    return newconfids

# Define function to generate reports

def makeReports():

    for confNum in newconfids:
        confidLink = "https://tableaulive.webex.com/mw3000/mywebex/report/sessiondetail.do?siteurl=tableaulive&confID=" + confNum + "&confName=Tableau+Desktop+Server+Administration+AM+EST&regCount=-1"
        driver.get(confidLink)
        print "Grabbing report for confid %s" % confNum
        time.sleep(2)
        reportLink = "https://tableaulive.webex.com/mw3000/mywebex/report/exportdetail.do?siteurl=tableaulive&regCount=-1&confID=" + confNum + "&confName=Tableau%20Desktop%20Server%20Administration%20AM%20EST&currentGMT=Eastern%20Daylight%20Time%20%28New%20York%2C%20GMT-04%3A00%29"
        driver.get(reportLink)
        time.sleep(2)
        print "Report for confid %s has been generated." % confNum
        
# Define function to concatenate files and output one master file

def finalReport():
    filepath = r'C:\Users\estam\Desktop\virtualReports'
    outfile = r'C:\Users\estam\Desktop\virtuals.csv'

    os.chdir(filepath)
    fileList = glob.glob("*.csv")
    dfList = []
    print "\n"
    for filename in fileList:
        print(filename + " successfully saved")
        df = pd.read_csv(filename, header =4, encoding = 'utf-16', delimiter = '\t')
        dfList.append(df)
    concatDf = pd.concat(dfList, axis = 0)

# Sort the participants by email address
    concatDf = concatDf.sort((['Email']), ascending = True)

# Remove the trainer's attendance details
    concatDf = concatDf[concatDf.Email != (login+'@tableau.com')]

# Output the concatenated report
    concatDf.to_csv(outfile, index = None)

# Delete the individual attendance reports
# Note: this cleans the folder for next time
    for filename in fileList:
        os.remove(filename)
        print filename + " successfully removed"

# **********************************************
# ***** CALL FUNCTIONS AND GENERATE REPORT *****
# **********************************************

# Program Announcemnt Header

print """

 **********************************************
 **********************************************
 *** For automating dem virtual attendances ***
 ***   Cuz ain't nobody got time for that   ***
 **********************************************
 **********************************************

 Abbreviated Course ID List:
 1 = Desktop Fundamentals AM
 2 = Desktop Fundamentals PM
 3 = Desktop Advanced AM
 4 = Desktop Advanced PM
 5 = Server Admin AM
 6 = Server Admin PM\n
"""

# Initial User Prompt/Input
course = userCourse()
login = raw_input('Enter the Tableau login: \n')
pword = raw_input('Enter the password: \n')

# Create a webdriver profile and open a browser
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", 'C:\Users\estam\Desktop\\virtualReports')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel,text/plain,text/csv,application/csv,application/download,application/octet-stream")
profile.set_preference("toolkit.startup.max_resumed_crashes", "-1")

driver = webdriver.Firefox(firefox_profile=profile)
driver.get("https://tableaulive.webex.com/mw3000/mywebex/default.do?siteurl=tableaulive&service=10")
time.sleep(6)

# Enter the login information for the website
driver.switch_to.frame('mainFrame')
loginid = driver.find_element_by_id('mwx-ipt-username')
loginid.send_keys(login)
time.sleep(1)
passwd = driver.find_element_by_id('mwx-ipt-password')
passwd.send_keys(pword)
time.sleep(2)
passwd.submit()

# Navigate to the reports area of the website
goReports()

# Store the links relevant to the Tableau course selected earlier
links = driver.find_elements_by_partial_link_text(courseText(course))

# Retrieve the conference IDs for each training session
newconfids = getConfids()

# Use the conference IDs to download each report
makeReports()

# Close the web browser
driver.quit()

# Combine all of the files into one final report file
finalReport()

print("\nFinal report is now available.\n")