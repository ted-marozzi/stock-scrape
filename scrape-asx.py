# Web scaping
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from bs4 import BeautifulSoup

# Google
import gspread
from oauth2client.service_account import ServiceAccountCredentials



import datetime

URL_COL = 11
URL_ROW = 6

PRICE_COL = 6
PRICE_ROW = 6

TIME_ROW = 6
TIME_COL = 7


def parse(url):
    #CHROME_PATH = '/usr/bin/google-chrome'
    CHROMEDRIVER_PATH = 'C:/Users/Maroz/Documents/chromedriver_win32/chromedriver.exe'
    WINDOW_SIZE = "1920,1080"

    chrome_options = Options()  
    chrome_options.add_argument("--headless")  
    chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
    #chrome_options.binary_location = CHROME_PATH

    
    driver = webdriver.Chrome(CHROMEDRIVER_PATH, options=chrome_options)
    print(url)
    driver.get(url)
    try:
        WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.CSS_SELECTOR,"span.ng-binding")))
    except:
        print("Exception")

    finally:
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        driver.quit()
        return soup
        

def getSharePrice(asx_soup):
    sharePrice = asx_soup.find_all("span", {"class": "ng-binding"})

    sharePrice = [_.text for _ in sharePrice]
    sharePrice = sharePrice[1]

    return sharePrice


def accessSheet():
    
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('secret.json', scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.

    
    # sheet = client.open("Investments").get_worksheet(1)
    worksheetList = client.open("Investments").worksheets()
    
    return worksheetList


def getUrlsSheet(worksheetList):

    for sheet in worksheetList:
        if sheet.title == "Stats":

            urls = sheet.col_values(URL_COL)
            numStocks = len(urls)-3
            urls = urls[URL_ROW-1:URL_ROW+numStocks]
            statsSheet = sheet
        elif sheet.title == "Price Log":
            priceLogSheet = sheet
    return urls, statsSheet, priceLogSheet



def insertSharePrice(urls, statsSheet, priceLogSheet):
    i = 0


    # Get length of wealth col
    colLen = len(priceLogSheet.col_values(1))
    for url in urls:
        shareSoup = parse(url)
        sharePrice = getSharePrice(shareSoup)

        statsSheet.update_cell(PRICE_ROW + i, PRICE_COL, sharePrice)
        priceLogSheet.update_cell(colLen+1,2+i,sharePrice)
  
        currentDT = str(datetime.datetime.now())
        currentDate = datetime.datetime.today().strftime('%d/%m/%y')
       
        

        statsSheet.update_cell(TIME_ROW + i, TIME_COL, currentDT)
        priceLogSheet.update_cell(colLen+1,1,currentDate)
        i = i + 1




#Get url from sheet
def main():
    
    worksheetList = accessSheet()

    (urls, statsSheet, priceLogSheet) = getUrlsSheet(worksheetList)

    insertSharePrice(urls, statsSheet, priceLogSheet)


if __name__ == '__main__':
    main()
