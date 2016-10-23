from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook

def parse_row(td_element):
    row = []
    fund_a= td_element.find_element_by_tag_name("a")
    row.append(fund_a.get_attribute("href"))
    row.append(fund_a.text)
    td = td_element.find_elements_by_tag_name("td")
    for i in range(1, len(td)):
        row.append(td[i].text)
    return row


def main():
    wb = Workbook()
    ws = wb.active    
    browser = webdriver.Chrome()
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
