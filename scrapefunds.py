'''
Author : Bhishan Bhandari
bbhishan@gmail.com

Porgram Dependencies 
selenium	to install type in terminal/command pip install selenium 

openpyxl 	to install type in terminal/command pip install openpyxl

'''

from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook

def parse_row(td_element):
    '''
	params: table data element <td> 
	Extracts the link of the fund, fund name, "STATE", "AUM", "2014Q4", "2015Q1", "CLIENTS", "EMPLOYEES"
	Makes a list of the above mentioned data and returns it.
    '''
    row = []
    fund_a= td_element.find_element_by_tag_name("a")
    row.append(fund_a.get_attribute("href"))
    row.append(fund_a.text)
    td = td_element.find_elements_by_tag_name("td")
    for i in range(1, len(td)):
        row.append(td[i].text)
    return row


def main():
    '''
	Uses openpyxl to instantiate a excel workbook. Uses selenium to open browser instance. Maximizes the browser so that the table data is not hidden.
	Selects filter to display all data in a single page.
	Iterates over all the table data <td> element and passes to the parse_row(td_element) method for further extraction.
	Appends the row returned by the parse_row(td_element) method to the excel sheet. 
    '''
    wb = Workbook()
    ws = wb.active    
    browser = webdriver.Chrome()
    browser.maximize_window()
    browser.get("http://www.octafinance.com/hedge-funds/hedge-funds-list/")
    time.sleep(20)
    #body = browser.find_element_by_tag_name('body')
    #body.send_keys(Keys.ESCAPE)        
    
    req_content_elem = browser.find_element_by_class_name("entry-content")
    select_elem = req_content_elem.find_element_by_tag_name("select")
    for option in select_elem.find_elements_by_tag_name("option"):
        if option.text == "All":
            option.click()
    time.sleep(20)
    

    odds = browser.find_elements_by_class_name("odd")
    evens = browser.find_elements_by_class_name("even")
    ws.append(["LINK", "FUND", "STATE", "AUM", "2014Q4", "2015Q1", "CLIENTS", "EMPLOYEES"])
    
    for odd, even in zip(odds, evens):
        try:
            odd_row = parse_row(odd)
            ws.append(odd_row)
            even_row = parse_row(even)
            ws.append(even_row)
        except:
            print "couldn't parse tow rows."
        wb.save("scraped.xlsx")
            


if __name__ == '__main__':
    main()
