"""
    An automation script to extract the data of the company from the "Ministry of Corporate Affairs" government site.
    Given a valid Company name, a user should be able to extract the data in the tabular format, of the first company
    name appeared related to the input provided.
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from pathlib import Path
from time import sleep
import pandas as pd

#Local Path for Chrome driver
DRIVER_PATH = "C:\Projects\web-scrapping\chrome_driver\chromedriver.exe"
driver = webdriver.Chrome(executable_path=DRIVER_PATH)


def find_company_data(name):
    #The URL of the site we are automating
    driver.get("https://www.mca.gov.in/")

    #For maximizing the Window size
    driver.maximize_window()

    #For closing guest Pop-Up on the web Page
    #pop_up_closing = driver.find_element_by_xpath("//*[@id='imgModalPopup']/div/div/div/p/button/span")
    #pop_up_closing.click()

    #For Hovering over menu
    action = ActionChains(driver)

    menu_option1 = driver.find_element_by_xpath("//*[@id='main_nav']/ul/li[6]/a")
    sleep(2)

    hover1 = action.move_to_element(menu_option1)
    hover1.perform()

    menu_option2 = driver.find_element_by_xpath("//*[@id='main_nav']/ul/li[6]/ul/li[3]/a")
    sleep(2)
    hover2 = action.move_to_element(menu_option2)
    hover2.perform()

    view_option = driver.find_element_by_xpath("//*[@id='main_nav']/ul/li[6]/ul/li[3]/ul/li[1]/a")
    sleep(2)
    view_option.click()

    search_icon = driver.find_element_by_xpath("//div//td//img")
    sleep(2)
    search_icon.click()
    sleep(2)

    company_name_textbox = driver.find_element_by_xpath("//*[@id='searchcompanyname']")
    sleep(1)
    company_name_textbox.send_keys(name)

    search_button = driver.find_element_by_xpath("//*[@id='findcindata']")
    search_button.click()
    sleep(2)

    company_path = driver.find_element_by_xpath("//*[@id='cinlist']/tbody/tr[1]/td[1]/a")
    sleep(2)
    company_path.click()

    #For Page Scrolling
    driver.execute_script("window.scrollTo(0, 120)")
    sleep(1)

    captcha_fill_box = driver.find_element_by_xpath("//*[@id='userEnteredCaptcha']")
    captcha_fill_box.click()

    #CAPTCHA filling time
    sleep(8)

    submit_button = driver.find_element_by_xpath("//*[@id='companyLLPMasterData_0']")
    sleep(2)
    submit_button.click()

    #For Scrolling Operation
    driver.execute_script("window.scrollTo(0, 270)")
    sleep(5)

    rows_path = driver.find_elements_by_xpath("//*[@id='resultTab1']/tbody/tr")
    sleep(2)

    #Counting number of Rows
    row_len = len(rows_path)

    cols_path = driver.find_elements_by_xpath("//*[@id='resultTab1']/tbody/tr[1]/td")
    sleep(2)

    #Counting number of Columns
    col_len = len(cols_path)

    r = 1
    res_list = []
    result = pd.DataFrame()

    while 1:
        try:
            #Extracting data from the result
            col_1 = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[" + str(r) + "]/td[1]").text
            col_2 = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[" + str(r) + "]/td[2]").text

            #Assigning data to two separate columns
            table_dict = {'Query': col_1, 'Answer': col_2}

            #Adding all data in a table
            res_list.append(table_dict)
            result = pd.DataFrame(res_list)
            r += 1

        #When there is no more data to extract
        except NoSuchElementException:

            #First company name as a result after user-input
            company_name = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[2]/td[2]").text

            #File location where data getting saved in a csv format
            filepath = Path(fr'C:\Projects\web-scrapping\data\{company_name}.csv')
            result.to_csv(filepath)
            break

    print("\n==>>FILE SAVED.....")

    #Printing rows and columns of extracted data
    print("Rows :" + str(row_len))
    print("Columns :" + str(col_len))

    #Closing the browser
    driver.close()


if __name__ == '__main__':
    #Company name whose data to be extracted
    company = input("Enter company name : ")

    print("\nFetching Data.........")

    #Reminder for the user to fill the CAPTCHA, which isn't automated to fill up automatically
    print("REMINDER : Be on the browser to fill the captcha....")
    find_company_data(company)
