#########################################################################################
# Author: Yifan Gao                                                                     #
# This script runs web scraping technique on ANA website, obtain available flight data  #
# with award ticket, and finally send all results during a specific period to a list of #
# email addresses.                                                                      #
# Known bugs:                                                                           #
# 1. The script runs periodically but will stop on clicking next day button on the page #
# where you can see detail round trip flight information.                               #
# 2. For some dates, there is simply no results matching search criteria. The script    #
# will fail if such instance is encountered at searching page.                          #
# Prerequisites:                                                                        #
# 1. It is recommended to run Python 3.7 for this script. When running Python 3.6, I    #
# received an error with lxml parser so Python 3.6 does not work. I did not test other  #
# versions so they may work as well.                                                    #
# 2. Selenium needs browser driver to initialize web page so please download browser    #
# driver to your work directory. In this script, I use Chrome driver. Here is the url   #
# to download: https://chromedriver.storage.googleapis.com/index.html?path=2.45/        #
# 3. This script is written in Windows so it's recommended to run in Windows. It may    #
# work in MacOS or Linux.                                                               #
# 4. Install package selenium, BeautifulSoup4 and lxml before executing the script.     #
#########################################################################################
import selenium.webdriver as wd
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import datetime as dt
from datetime import timedelta
import re
import pandas as pd
import lxml
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tabulate import tabulate
import numpy as np
import cv2
from PIL import ImageGrab

def ana_data_inquiry(fromcity,tocity,depdate,retdate,interval,headcount,travelclass):
    """This function simulates user behavior when booking award ticket on ANA website
       and return a list that shows available award tickets based on input arguments.
       It encompasses four pages when scraping data - first page where you see 'Award
       Booking' button, second page where you see login information, third page where
       you see a list of search criteria you can enter for searching a flight, and
       finally fourth page where you can see detail round trip flight information.
       After landing to the fourth page, the function will first collect information on
       current date, then click next departure date, collect information, and click
       next return date and collect information. The function will repeat such N times
       (N == internal).
       Please note this function only works with round trip mode.

    fromcity: city name where round trip starts, and it must be a string. Hypothetically
    this should support all cities that ANA and its partners fly to/form. For this script
    I only test with US cities that have ANA flight.
    tocity: city name where round trip ends, and it must be a string. Hypothetically
    this should support all cities that ANA and its partners fly to/form. For this script
    I only test with Tokyo.
    depdate: departure date and it must be a string. Format should be '%Y-%m-%d'. The year
    must be current year.
    retdate: return date and it must be a string. Format should be '%Y-%m-%d'. The year
    must be current year. The ret date must be at least seven days later than dep date.
    interval: number of days that search runs from start date and it must be an integer.
    If the dep date is '2018-03-01', ret date is '2018-04-01' and interval is 10, the
    function will run until dep date and ret date roll to '2018-03-11' and '2018-04-11'.
    headcount: number of adult passengers. Must be between 1 and 9, both inclusive.
    travelclass: boarding class of round trip and it must be a string. Select one from
    'Economy', 'Premium Economy', 'Business', and 'First'.
    """

    # firefox_options = wd.FirefoxOptions()
    # # firefox_options.add_argument('--headless')
    # firefox_options.add_argument('--no-sandbox')
    # firefox_options.add_argument('--allow-running-insecure-content')
    # firefox_options.add_argument('--ignore-certificate-errors')
    # # capabilities = DesiredCapabilities.CHROME.copy()
    # # capabilities['acceptSslCerts'] = True
    # # capabilities['acceptInsecureCerts'] = True
    # global driver
    # driver = wd.Firefox(executable_path='F:/Seeds/WebDriver/Firefox/geckodriver.exe')

    # global driver
    # driver = wd.Edge('F:/Seeds/WebDriver/Edge/MicrosoftWebDriver.exe')

    ## Raise valueerror if return date is within six days from departure date.
    if (dt.datetime.strptime(retdate, '%Y-%m-%d') - dt.datetime.strptime(depdate, '%Y-%m-%d')).days <= 6:
        raise ValueError("The return date must be at least seven days later from departure date.")
    # ## Raise valueerror if return/departure date is not pulled from current year.
    # if (dt.datetime.now().year != dt.datetime.strptime(depdate, '%Y-%m-%d').year) | (dt.datetime.now().year != dt.datetime.strptime(retdate, '%Y-%m-%d').year):
    #     raise ValueError("The departure or return date must enter in current year.")
    ## Initialize chrome driver and set the mode to headless so that the broswer will run in the back end.
    chrome_options = wd.ChromeOptions()
    chrome_options.add_argument('--incognito')
    # chrome_options.add_argument('--headless')
    ## Replace your own directory that stores chromedriver in the first argument.
    global driver
    driver = wd.Chrome('F:/Seeds/WebDriver/Chrome/chromedriver.exe',
                       chrome_options=chrome_options)
    # driver.maximize_window()
    ## Set window size equal to your monitor's max window size so that all the values can be clicked in headless node.
    ## Adjust this accordingly with your own resolution ratio.
    driver.set_window_size(1920, 1080)

    ### Land on the first page, enter username and password to login to the third page.
    driver.get('https://www.ana.co.jp/en/us/')
    driver.find_element_by_link_text('Award Booking').click()
    ## Replace your own ANA membership number
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accountNumber')))
    driver.find_element_by_id('accountNumber').send_keys('4395927665')
    # driver.find_element_by_id('accountNumber').send_keys('4396496135')
    ## Replace your own ANA password
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'password')))
    driver.find_element_by_id('password').send_keys('3340cjz826')
    time.sleep(0.6)
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'amcMemberLogin')))
    driver.find_element_by_id('amcMemberLogin').click()
    ## Add this to allow some flexibility around capitalizing letters for cities.
    if (fromcity == 'Chicago') | (fromcity == 'chicago') | (fromcity == 'CHICAGO'):
        realfrom = 'Chicago(All)'
    elif (fromcity == 'New York') | (fromcity == 'new york') | (fromcity == 'NEW YORK'):
        realfrom = 'New York(All)'
    elif (fromcity == 'Washington,D.C.') | (fromcity == 'washington,D.C.') | (fromcity == 'WASHINGTON,D.C.'):
        realfrom = 'Washington,D.C.(All)'
    elif (fromcity == 'Houston') | (fromcity == 'houston') | (fromcity == 'HOUSTON'):
        realfrom = 'Houston'
    elif (fromcity == 'Seattle') | (fromcity == 'seattle') | (fromcity == 'SEATTLE'):
        realfrom = 'Seattle'
    elif (fromcity == 'San Francisco') | (fromcity == 'san francisco') | (fromcity == 'SAN FRANCISCO'):
        realfrom = 'San Francisco'
    elif (fromcity == 'Los Angeles') | (fromcity == 'los angeles') | (fromcity == 'LOS ANGELES'):
        realfrom = 'Los Angeles'
    elif (fromcity == 'San Jose') | (fromcity == 'san jose') | (fromcity == 'SAN JOSE'):
        realfrom = 'San Jose(SJC - California)'
    elif (fromcity == 'Tokyo') | (fromcity == 'tokyo') | (fromcity == 'TOKYO'):
        realfrom = 'Tokyo(All)'
    elif (fromcity == 'Shanghai') | (fromcity == 'shanghai') | (fromcity == 'SHANGHAI'):
        realfrom = 'Shanghai (All)'
    elif (fromcity == 'Beijing') | (fromcity == 'beijing') | (fromcity == 'BEIJING'):
        realfrom = 'Beijing'
    elif (fromcity == 'Hong Kong') | (fromcity == 'hong kong') | (fromcity == 'HONG KONG'):
        realfrom = 'Hong Kong'

    if (tocity == 'Chicago') | (tocity == 'chicago') | (tocity == 'CHICAGO'):
        realto = 'Chicago(All)'
    elif (tocity == 'New York') | (tocity == 'new york') | (tocity == 'NEW YORK'):
        realto = 'New York(All)'
    elif (tocity == 'Washington,D.C.') | (tocity == 'washington,D.C.') | (tocity == 'WASHINGTON,D.C.'):
        realto = 'Washington,D.C.(All)'
    elif (tocity == 'Houston') | (tocity == 'houston') | (tocity == 'HOUSTON'):
        realto = 'Houston'
    elif (tocity == 'Seattle') | (tocity == 'seattle') | (tocity == 'SEATTLE'):
        realto = 'Seattle'
    elif (tocity == 'San Francisco') | (tocity == 'san francisco') | (tocity == 'SAN FRANCISCO'):
        realto = 'San Francisco'
    elif (tocity == 'Los Angeles') | (tocity == 'los angeles') | (tocity == 'LOS ANGELES'):
        realto = 'Los Angeles'
    elif (tocity == 'San Jose') | (tocity == 'san jose') | (tocity == 'SAN JOSE'):
        realto = 'San Jose(SJC - California)'
    elif (tocity == 'Tokyo') | (tocity == 'tokyo') | (tocity == 'TOKYO'):
        realto = 'Tokyo(All)'
    elif (tocity == 'Shanghai') | (tocity == 'shanghai') | (tocity == 'SHANGHAI'):
        realto = 'Shanghai (All)'
    elif (tocity == 'Beijing') | (tocity == 'beijing') | (tocity == 'BEIJING'):
        realto = 'Beijing'
    elif (tocity == 'Hong Kong') | (tocity == 'hong kong') | (tocity == 'HONG KONG'):
        realto = 'Hong Kong'




    ### Land on the third page and enter necessary information to proceed
    driver.find_element_by_id('departureAirportCode:field_pctext').send_keys(realfrom)
    ## Pause 0.6 seconds to allow dropdown list appears from departure city
    time.sleep(0.6)
    driver.find_element_by_id('departureAirportCode:field_pctext').send_keys(Keys.ENTER)
    driver.find_element_by_id('arrivalAirportCode:field_pctext').send_keys(realto)
    ## Pause 0.6 seconds to allow dropdown list appears from return city
    time.sleep(0.6)
    driver.find_element_by_id('arrivalAirportCode:field_pctext').send_keys(Keys.ENTER)

    driver.find_element_by_id('awardDepartureDate:field_pctext').click()
    ## Because calendar window is dynamic so code below ensures date entered can be found
    ## in the calendar window.
    try:
        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate + '"]').click()
    except:
        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
        try:
            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate + '"]').click()
        except:
            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
            try:
                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate + '"]').click()
            except:
                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate + '"]').click()

    driver.find_element_by_id('awardReturnDate:field_pctext').click()
    ## Because calendar window is dynamic so code below ensures date entered can be found
    ## in the calendar window.
    try:
        driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate + '"]').click()
    except:
        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
        try:
            driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate + '"]').click()
        except:
            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
            try:
                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate + '"]').click()
            except:
                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate + '"]').click()

    boardingclass = Select(driver.find_element_by_id('boardingClass'))
    boardingclass.select_by_visible_text(travelclass)
    passenger = Select(driver.find_element_by_id('adult:count'))
    passenger.select_by_visible_text(str(headcount))
    driver.find_element_by_class_name('btnFloat').click()

    ### Land on the fourth page and start collecting available flight. It will first click next departure date,
    ### collect information, and click next return date and collect information. The function will repeat such
    ### N times(N == internal).

    ## Wait for at most 5 seconds until 'Select an Award Type' banner appears. If it appears less than 5 seconds
    ## the script will proceed immediately.
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
    anadata = []

    for i in range(interval):
        # try:
        #     driver.find_element_by_xpath('//*[@for="selectAward01"]').click()
        # except:
        #     pass
        # dep_adjust = 0 if i == 0 else dep_adjust1
        # ret_adjust = 0 if i == 0 else ret_adjust1

        ## Convert driver page to a soup object (lxml) so that it's easier to parse
        soup = BeautifulSoup(driver.page_source, 'lxml')
        ## Find current dep date, ret date, next dep date and ret date
        curdepday = dt.datetime.strptime(soup.find_all('div', attrs='selectItineraryOutbound')[0].get_text('|', strip=True).split('|')[1] + ' ' + str(dt.datetime.strptime(depdate, '%Y-%m-%d').year),'%b %d %Y').strftime('%Y-%m-%d')
        curretday = dt.datetime.strptime(soup.find_all('div', attrs='selectItineraryInbound')[0].get_text('|', strip=True).split('|')[1] + ' ' + str(dt.datetime.strptime(retdate, '%Y-%m-%d').year),'%b %d %Y').strftime('%Y-%m-%d')
        nextdepday = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=1)).strftime('%b %#d')
        nextretday = (dt.datetime.strptime(curretday, '%Y-%m-%d') + timedelta(days=1)).strftime('%b %#d')
        # curdepday = (dt.datetime.strptime(depdate, '%Y-%m-%d') + timedelta(days=i+dep_adjust)).strftime('%Y-%m-%d')
        # curretday = (dt.datetime.strptime(retdate, '%Y-%m-%d') + timedelta(days=i+ret_adjust)).strftime('%Y-%m-%d')
        # nextdepday = (dt.datetime.strptime(depdate, '%Y-%m-%d') + timedelta(days=i + dep_adjust + 1)).strftime('%b %#d')
        # nextretday = (dt.datetime.strptime(retdate, '%Y-%m-%d') + timedelta(days=i + ret_adjust + 1)).strftime('%b %#d')

        ## Convert round trip section to a soup object (lxml) so that it's easier to parse
        triptable = soup.find_all('td', attrs='selectItineraryDetail')
        for j in range(len(triptable)):
            ## Only available ticket will return empty list
            if triptable[j].find_all('p') == []:
                fromairport = triptable[j].find_all('div', attrs='airportDeparture')[0].get_text()
                toairport = triptable[j].find_all('div', attrs='airportArrival')[0].get_text()
                if '(' in fromairport:
                    fromcity1 = fromairport.split('(')[0]
                else:
                    fromcity1 = fromairport
                if '(' in toairport:
                    tocity1 = toairport.split('(')[0]
                else:
                    tocity1 = toairport
                ## this step is to eliminate condition where from or to city is a layover city. For example,
                ## Chicago-Vancouver-Tokyo' if ticket is available between Vancouver-Tokyo it will be removed
                ## because it's not a direct flight from Chicago to Tokyo.
                if (fromcity1.lower() not in [fromcity.lower(),tocity.lower()])|(tocity1.lower() not in [fromcity.lower(),tocity.lower()]):
                    pass
                else:
                    if triptable[j].find_all('div', attrs=re.compile('outbound')) != []:
                        flightdate = curdepday
                    elif triptable[j].find_all('div', attrs=re.compile('inbound')) != []:
                        flightdate = curretday
                    flight = triptable[j].find_all('div', attrs='detailInformation')[0].get_text('|', strip=True).split('|')[0]
                    ## Write flight date, from airpot, to airport and flight number to a list
                    rowdata = [flightdate, fromairport, toairport, flight]
                    anadata.append(rowdata)
                # ## Filter out other star alliance airlines and focus only on ANA available ticket.
                if 'NH' in triptable[j].find_all('div', attrs='detailInformation')[0].get_text('|', strip=True).split('|')[0]:
                    fromairport = triptable[j].find_all('div', attrs='airportDeparture')[0].get_text()
                    toairport = triptable[j].find_all('div', attrs='airportArrival')[0].get_text()
                    if '(' in fromairport:
                        fromcity1 = fromairport.split('(')[0]
                    else:
                        fromcity1 = fromairport
                    if '(' in toairport:
                        tocity1 = toairport.split('(')[0]
                    else:
                        tocity1 = toairport
                    ## this step is to eliminate condition where from or to city is a layover city. For example,
                    ## Chicago-Vancouver-Tokyo' if ticket is available between Vancouver-Tokyo it will be removed
                    ## because it's not a direct flight from Chicago to Tokyo.
                    if (fromcity1.lower() not in [fromcity.lower(),tocity.lower()])|(tocity1.lower() not in [fromcity.lower(),tocity.lower()]):
                        pass
                    else:
                        if triptable[j].find_all('div', attrs=re.compile('outbound')) != []:
                            flightdate = curdepday
                        elif triptable[j].find_all('div', attrs=re.compile('inbound')) != []:
                            flightdate = curretday
                        flight = triptable[j].find_all('div', attrs='detailInformation')[0].get_text('|', strip=True).split('|')[0]
                        ## Write flight date, from airpot, to airport and flight number to a list
                        rowdata = [flightdate, fromairport, toairport, flight]
                        anadata.append(rowdata)
            else:
                pass

        ## Wait for at most 15 seconds until next depart day arrow button is clickable. If it's clickable less than 15 seconds
        ## the script will proceed immediately.
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Depart Next day"]'))).click()


        try:
            ## Wait for at most 20 seconds until next depart date appears. If it appears less than 15 seconds
            ## the script will proceed immediately.
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[text()=' + '"' + nextdepday + '"]')))
            # dep_adjust1 = max(0,dep_adjust)
        except:
            ## This try/except code accommendate scenario when clicking next depart day arrow button and no results matching
            ## search criteria. If such instance happens it will ignore the error and jump to next date and rerun with new
            ## date. The code will repeat four times at most if error ccntinues to appear. This script could be improved with
            ## while loop. I was lazy so just wrote in this way.
            try:
                driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                driver.find_element_by_class_name('btnLabel').click()
                driver.find_element_by_id('awardDepartureDate:field_pctext').click()
                depdate1 = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=2)).strftime('%Y-%m-%d')
                depdate2 = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=3)).strftime('%Y-%m-%d')
                depdate3 = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=4)).strftime('%Y-%m-%d')
                depdate4 = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=5)).strftime('%Y-%m-%d')
                try:
                    driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate1 + '"]').click()
                    # dep_adjust1 = max(1, dep_adjust)
                except:
                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate1 + '"]').click()
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate1 + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate1 + '"]').click()
                driver.find_element_by_id('awardReturnDate:field_pctext').click()
                try:
                    driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                except:
                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()

                driver.find_element_by_xpath('//*[contains(@onclick,"confirmChargeableSeatsReleaseDialog")]').click()
                try:
                    WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                except:
                    driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                    driver.find_element_by_id('awardDepartureDate:field_pctext').click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate2 + '"]').click()
                        # dep_adjust1 = max(2, dep_adjust)
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate2 + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate2 + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate2 + '"]').click()
                    driver.find_element_by_id('awardReturnDate:field_pctext').click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                    driver.find_element_by_class_name('btnFloat').click()
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                    except:
                        driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                        driver.find_element_by_id('awardDepartureDate:field_pctext').click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate3 + '"]').click()
                            # dep_adjust1 = max(3, dep_adjust)
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate3 + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate3 + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate3 + '"]').click()
                        driver.find_element_by_id('awardReturnDate:field_pctext').click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                        driver.find_element_by_class_name('btnFloat').click()
                        try:
                            WebDriverWait(driver, 1).until(
                                EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                        except:
                            driver.find_element_by_xpath(
                                '//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                            driver.find_element_by_id('awardDepartureDate:field_pctext').click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate4 + '"]').click()
                                # dep_adjust1 = max(4, dep_adjust)
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate4 + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    try:
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate4 + '"]').click()
                                    except:
                                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdate4 + '"]').click()
                            driver.find_element_by_id('awardReturnDate:field_pctext').click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    try:
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                                    except:
                                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + curretday + '"]').click()
                            driver.find_element_by_class_name('btnFloat').click()
                            WebDriverWait(driver, 1).until(
                                EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
            except:
                driver.quit()
                ana_data_inquiry(fromcity, tocity, depdate, retdate, interval, headcount, travelclass)



        ## Wait for at most 15 seconds until next return day arrow button is clickable. If it's clickable less than 15 seconds
        ## the script will proceed immediately.
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Return Next day"]'))).click()

        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[text()=' + '"' + nextretday + '"]')))
            # ret_adjust1 = max(0,ret_adjust)
        except:
            ## This try/except code accommendate scenario when clicking next return day arrow button and no results matching
            ## search criteria. If such instance happens it will ignore the error and jump to next date and rerun with new
            ## date. The code will repeat four times at most if error ccntinues to appear. This script could be improved with
            ## while loop. I was lazy so just wrote in this way.
            try:
                driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                driver.find_element_by_class_name('btnLabel').click()
                depdatenext = (dt.datetime.strptime(curdepday, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
                retdate1 = (dt.datetime.strptime(curretday, '%Y-%m-%d') + timedelta(days=2)).strftime('%Y-%m-%d')
                retdate2 = (dt.datetime.strptime(curretday, '%Y-%m-%d') + timedelta(days=3)).strftime('%Y-%m-%d')
                retdate3 = (dt.datetime.strptime(curretday, '%Y-%m-%d') + timedelta(days=4)).strftime('%Y-%m-%d')
                retdate4 = (dt.datetime.strptime(curretday, '%Y-%m-%d') + timedelta(days=5)).strftime('%Y-%m-%d')
                driver.find_element_by_id('awardDepartureDate:field_pctext').click()
                try:
                    driver.find_element_by_xpath('//*[@abbr=' + '"' + depdatenext + '"]').click()
                except:
                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + depdatenext + '"]').click()
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdatenext + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + depdatenext + '"]').click()
                driver.find_element_by_id('awardReturnDate:field_pctext').click()
                try:
                    driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate1 + '"]').click()
                    # ret_adjust1 = max(1,ret_adjust)
                except:
                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate1 + '"]').click()
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate1 + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate1 + '"]').click()
                driver.find_element_by_xpath('//*[contains(@onclick,"confirmChargeableSeatsReleaseDialog")]').click()
                try:
                    WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                except:
                    driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                    driver.find_element_by_id('awardReturnDate:field_pctext').click()
                    try:
                        driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate2 + '"]').click()
                        # ret_adjust1 = max(2,ret_adjust)
                    except:
                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate2 + '"]').click()
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate2 + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate2 + '"]').click()
                    driver.find_element_by_class_name('btnFloat').click()
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                    except:
                        driver.find_element_by_xpath('//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                        driver.find_element_by_id('awardReturnDate:field_pctext').click()
                        try:
                            driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate3 + '"]').click()
                            # ret_adjust1 = max(3,ret_adjust)
                        except:
                            driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate3 + '"]').click()
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate3 + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate3 + '"]').click()
                        driver.find_element_by_class_name('btnFloat').click()
                        try:
                            WebDriverWait(driver, 1).until(
                                EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
                        except:
                            driver.find_element_by_xpath(
                                '//*[contains(@aria-controls,"cmnErrorMessageWindow")]').click()
                            driver.find_element_by_id('awardReturnDate:field_pctext').click()
                            try:
                                driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate4 + '"]').click()
                                # ret_adjust1 = max(4,ret_adjust)
                            except:
                                driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                try:
                                    driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate4 + '"]').click()
                                except:
                                    driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                    try:
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate4 + '"]').click()
                                    except:
                                        driver.find_elements_by_xpath('//*[text()="Next 3 months"]')[0].click()
                                        driver.find_element_by_xpath('//*[@abbr=' + '"' + retdate4 + '"]').click()
                            driver.find_element_by_class_name('btnFloat').click()
                            WebDriverWait(driver, 1).until(
                                EC.presence_of_element_located((By.XPATH, '//*[text()="Select an Award Type"]')))
            except:
                driver.quit()
                ana_data_inquiry(fromcity, tocity, depdate, retdate, interval, headcount, travelclass)


    global glob_headcount, glob_travelclass
    glob_headcount = locals()['headcount']
    glob_travelclass = locals()['travelclass']
    driver.quit()
    return anadata


def fetch_and_send():
    """This function runs ana_data_inquiry function multiple times for various cities
       and combine results as a pandas data frame. After that, it sends pandas df to
       a list of recipients through email.

    """
    # data1 = ana_data_inquiry('Tokyo', 'Chicago', '2019-06-05', '2019-06-15', 150, 1, 'Business')
    # finaldata = pd.DataFrame(data1,
    #                          columns=['Date', 'From', 'To', 'Flight'])

    # data1 = ana_data_inquiry('Chicago', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data2 = ana_data_inquiry('New York', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data3 = ana_data_inquiry('Seattle', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data4 = ana_data_inquiry('San Francisco', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data5 = ana_data_inquiry('San Jose', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data6 = ana_data_inquiry('Los Angeles', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data7 = ana_data_inquiry('Houston', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data8 = ana_data_inquiry('Washington,D.C.', 'Tokyo', '2019-05-20', '2019-06-08', 7, 1,'Business')
    # data9 = ana_data_inquiry('Chicago', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data10 = ana_data_inquiry('New York', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data11 = ana_data_inquiry('Seattle', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data12 = ana_data_inquiry('San Francisco', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data13 = ana_data_inquiry('San Jose', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data14 = ana_data_inquiry('Los Angeles', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data15 = ana_data_inquiry('Houston', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data16 = ana_data_inquiry('Washington,D.C.', 'Shanghai', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data17 = ana_data_inquiry('Chicago', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data18 = ana_data_inquiry('New York', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data19 = ana_data_inquiry('Seattle', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data20 = ana_data_inquiry('San Francisco', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data21 = ana_data_inquiry('San Jose', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data22 = ana_data_inquiry('Los Angeles', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data23 = ana_data_inquiry('Houston', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data24 = ana_data_inquiry('Washington,D.C.', 'Beijing', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data25 = ana_data_inquiry('Chicago', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data26 = ana_data_inquiry('New York', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data27 = ana_data_inquiry('Seattle', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data28 = ana_data_inquiry('San Francisco', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data29 = ana_data_inquiry('San Jose', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data30 = ana_data_inquiry('Los Angeles', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data31 = ana_data_inquiry('Houston', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    # data32 = ana_data_inquiry('Washington,D.C.', 'Hong Kong', '2019-05-20', '2019-06-08', 7, 1, 'Business')
    #
    # finaldata = pd.DataFrame(data1+data2+data3+data4+data5+data6+data7+data8+data9+data10+data11+data12+data13+data14+
    #                          data15+data16+data17+data18+data19+data20+data21+data22+data23+data24+data25+data26+data27+
    #                          data28+data29+data30+data31+data32, columns=['Date', 'From', 'To', 'Flight'])

    # data1 = ana_data_inquiry('Chicago', 'Tokyo', '2020-01-01', '2020-01-08', 41, 1,'Business')
    # data2 = ana_data_inquiry('New York', 'Tokyo', '2020-01-05', '2020-01-12', 37, 1,'Business')
    # data3 = ana_data_inquiry('Seattle', 'Tokyo', '2020-01-01', '2020-01-08', 41, 1,'Business')
    # data4 = ana_data_inquiry('San Francisco', 'Tokyo', '2020-01-05', '2020-01-12', 37, 1,'Business')
    # data5 = ana_data_inquiry('San Jose', 'Tokyo', '2020-01-01', '2020-01-08', 41, 1,'Business')
    # data6 = ana_data_inquiry('Los Angeles', 'Tokyo', '2020-01-05', '2020-01-12', 37, 1,'Business')
    # data7 = ana_data_inquiry('Houston', 'Tokyo', '2020-01-05', '2020-01-12', 37, 1,'Business')
    # data8 = ana_data_inquiry('Washington,D.C.', 'Tokyo', '2020-01-05', '2020-01-12', 37, 1,'Business')
    #
    # finaldata = pd.DataFrame(data1+data2+data3+data4+data5+data6+data7+data8,
    #                          columns=['Date', 'From', 'To', 'Flight'])
    # finaldata = pd.DataFrame.drop_duplicates(finaldata)

    data1 = ana_data_inquiry('Chicago', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2,'Business')
    data2 = ana_data_inquiry('New York', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2,'Business')
    data3 = ana_data_inquiry('Seattle', 'Tokyo', '2020-03-28', '2020-04-04', 55, 2,'Business')
    data4 = ana_data_inquiry('San Francisco', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2,'Business')
 #   data5 = ana_data_inquiry('San Jose', 'Tokyo', '2020-03-30', '2020-04-06', 60, 2,'Business')
    data6 = ana_data_inquiry('Los Angeles', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2,'Business')
    data7 = ana_data_inquiry('Houston', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2,'Business')
    data8 = ana_data_inquiry('Washington,D.C.', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2, 'Business')

    finaldata = pd.DataFrame(data1+data2+data3+data4+data6+data7+data8,
                             columns=['Date', 'From', 'To', 'Flight'])
    finaldata = pd.DataFrame.drop_duplicates(finaldata)

    # data3 = ana_data_inquiry('Seattle', 'Tokyo', '2020-03-28', '2020-04-04', 60, 2, 'Business')
    # finaldata = pd.DataFrame(data3,
    #                          columns=['Date', 'From', 'To', 'Flight'])

    # data1 = ana_data_inquiry('Chicago', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2,'Business')
    # data2 = ana_data_inquiry('Chicago', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data3 = ana_data_inquiry('New York', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2,'Business')
    # data4 = ana_data_inquiry('New York', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data5 = ana_data_inquiry('Seattle', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2,'Business')
    # data6 = ana_data_inquiry('Seattle', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data7 = ana_data_inquiry('San Francisco', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2, 'Business')
    # data8 = ana_data_inquiry('San Francisco', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data9 = ana_data_inquiry('San Jose', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2, 'Business')
    # data10 = ana_data_inquiry('San Jose', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data11 = ana_data_inquiry('Los Angeles', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2, 'Business')
    # data12 = ana_data_inquiry('Los Angeles', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data13 = ana_data_inquiry('Houston', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2, 'Business')
    # data14 = ana_data_inquiry('Houston', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    # data15 = ana_data_inquiry('Washington,D.C.', 'Tokyo', '2020-01-01', '2020-01-08', 52, 2, 'Business')
    # data16 = ana_data_inquiry('Washington,D.C.', 'Tokyo', '2020-02-23', '2020-03-01', 8, 2, 'Business')
    #
    # finaldata = pd.DataFrame(data1+data2+data3+data4+data5+data6+data7+data8+data9+data10+data11+data12+data13+data14+data15+data16,
    #                          columns=['Date', 'From', 'To', 'Flight'])

    ## Sender and receiver's emails
    fromaddr = 'anaawards123@gmail.com'
    toaddr = 'anaawards123@gmail.com'
    ## List of email recipients to bcc
    # bccaddr = ['mars_gyf@hotmail.com','lena1211@icloud.com']
    bccaddr = ['mars_gyf@hotmail.com']
    text_format = """
    Hello,

    Here is your ANA data:

    {table}

    Regards,

    Your ANA friend"""

    html_format = """
    <html><body><p>Hello,</p>
    <p>Here is your ANA data:</p>
    {table}
    <p>Regards,</p>
    <p>Your ANA friend</p>
    </body></html>
    """

    ## Set up both text and html formats so that recipient email can parse the emaii body intelligently.
    text_format = text_format.format(
        table=tabulate(finaldata, headers=['Date', 'From', 'To', 'Flight'], tablefmt="grid", showindex=False))
    html_format = html_format.format(
        table=tabulate(finaldata, headers=['Date', 'From', 'To', 'Flight'], tablefmt="html", showindex=False))
    msg = MIMEMultipart(
        "alternative", None, [MIMEText(text_format), MIMEText(html_format, 'html')])
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Bcc'] = ','.join(bccaddr)
    msg['Subject'] = 'ANA Flight Availability - ' + glob_travelclass + ' Class (' +str(glob_headcount) + ' passenger(s))'
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    ## Sender email's login credentials
    server.login('anaawards123@gmail.com', 'valenti290854098')
    server.sendmail(fromaddr, bccaddr, msg.as_string())
    server.quit()

if __name__ == '__main__':
    while True:
        fetch_and_send()
        time.sleep(300)

# if __name__ == '__main__':
#     try:
#         # four character code object for video writer
#         # fourcc = cv2.VideoWriter_fourcc(*'XVID')
#         # # video writer object
#         # out = cv2.VideoWriter("output.avi", fourcc, 8.0, (1920, 1080))
#         while True:
#             ## Run fetch_and_send periodically. When one execution is finished, pause 5 mins and rerun.
#             # capture computer screen
#             # img = ImageGrab.grab()
#             # # convert image to numpy array
#             # img_np = np.array(img)
#             # # convert color space from BGR to RGB
#             # frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
#             # # show image on OpenCV frame
#             # out.write(frame)
#             # cv2.imshow("Screen", frame)
#             # cv2.waitKey(1)
#             fetch_and_send()
#             time.sleep(300)
#     except:
#         # driver.quit()
#         # out.release()
#         # cv2.destroyAllWindows()
#         # fromaddr = 'anaawards123@gmail.com'
#         # toaddr = 'anaawards123@gmail.com'
#         # ## List of email recipients to bcc
#         # bccaddr = ['mars_gyf@hotmail.com']
#         # text_format = """
#         #     Check"""
#         #
#         # html_format = """
#         #     Check
#         #     """
#         # ## Set up both text and html formats so that recipient email can parse the emaii body intelligently.
#         # text_format = text_format
#         # html_format = html_format
#         # msg = MIMEMultipart(
#         #     "alternative", None, [MIMEText(text_format), MIMEText(html_format, 'html')])
#         # msg['From'] = fromaddr
#         # msg['To'] = toaddr
#         # msg['Bcc'] = ','.join(bccaddr)
#         # msg['Subject'] = 'Big Check'
#         # server = smtplib.SMTP('smtp.gmail.com', 587)
#         # server.ehlo()
#         # server.starttls()
#         # server.ehlo()
#         # ## Sender email's login credentials
#         # server.login('anaawards123@gmail.com', 'ichigo290854098')
#         # server.sendmail(fromaddr, bccaddr, msg.as_string())
#         # server.quit()
#         driver.save_screenshot('screenshot ' + dt.datetime.now().strftime('%H%M%S%m%d%Y') + '.png')


# # four character code object for video writer
# fourcc = cv2.VideoWriter_fourcc(*'XVID')
# # video writer object
# out = cv2.VideoWriter("output.avi", fourcc, 8.0, (1920, 1080))
#
# while True:
#     # capture computer screen
#     img = ImageGrab.grab()
#     # convert image to numpy array
#     img_np = np.array(img)
#     # convert color space from BGR to RGB
#     frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
#     # show image on OpenCV frame
#     out.write(frame)
#     cv2.imshow("Screen", frame)
#     # write frame to video writer
#     if cv2.waitKey(1) == 27:
#         break
#
# out.release()
# cv2.destroyAllWindows()