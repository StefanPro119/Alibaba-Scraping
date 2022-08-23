from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas


driver = webdriver.Chrome(ChromeDriverManager().install())


driver.maximize_window()
time.sleep(2)

def is_it_less_or_equal_than_2005(year, element, company_name, main_products):
    if year <= 2005:
        dictionary = {}
        dictionary['Company Name'] = company_name
        dictionary['Year Established'] = year
        dictionary['Main Product'] = main_products.text
        dictionary['Web Page'] = element
        return dictionary


def find_links(list):
    links = driver.find_elements(By.XPATH, '//*[@id="root"]/div/section[3]/div[2]/div/div[1]/div[1]/div[2]/h3/a')
    for link in links:

        """ When it collects 15 items it breaks the for loop"""
        if len(list) >= 15:
            break

        element = link.get_attribute('href')
        company_name = link.text

        """It will open a new TAB in a new window with given link (like middle button on a mouse)"""
        driver.execute_script(f"window.open('{element}')")

        """It will switch to window that was previously opened"""
        driver.switch_to.window(driver.window_handles[-1])

        time.sleep(5)
        main_products = driver.find_elements(By.CLASS_NAME, 'content-value')[2]
        year_established = driver.find_elements(By.CLASS_NAME, 'content-value')[5]
        if len(year_established.text) > 4:
            year_established = driver.find_elements(By.CLASS_NAME, 'content-value')[6]
            year = int(year_established.text)
            dictionary = is_it_less_or_equal_than_2005(year, element, company_name, main_products)
        else:
            year = int(year_established.text)
            dictionary = is_it_less_or_equal_than_2005(year, element, company_name, main_products)


        if dictionary == None:
            pass
        else:
            list.append(dictionary)

        driver.close()
        driver.switch_to.window(driver.window_handles[0])



""" SEARCH_TEXT is item which you want to scrape from Alibaba Website 
( if you have more than one word, include "_" between words instead of "space" ) """
SEARCH_TEXT = input('\nWhat do you want to scrape? ( include "_" between words instead of "space" ): \n')
name_of_the_file = input('\nName the excel file: \n')

def main_program():
    for page in range(1, 10):
        driver.get(f"https://www.alibaba.com/trade/search?IndexArea=product_en&SearchText={SEARCH_TEXT}&page={page}&Country=CN&ta=y&param_order=CNTRY-CN,CAT-ISO&tab=verifiedManufactory&companyAuthTag=ISO&f0=y")
        time.sleep(2)
        find_links(list)

        """ When it collects 15 items it breaks the for loop"""
        if len(list) >= 15:
            break

list = []
main_program()

""" Make an Excel sheet """
df = pandas.DataFrame(list)

def make_hyperlink(value):
    for i in list:
        url = i['Web Page']
        return '=HYPERLINK("%s", "%s")' % (url.format(value), value)

df['Web Page'] = df['Web Page'].apply(lambda x: make_hyperlink(x))

df.to_word(f'{name_of_the_file}.xlsx')
