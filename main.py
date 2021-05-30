from time import sleep
from selenium import webdriver
from selenium.webdriver import ActionChains

#Local Path for Chrome driver
DRIVER_PATH = "C:\Projects\web-scrapping\chrome_driver\chromedriver.exe"
driver = webdriver.Chrome(executable_path=DRIVER_PATH)


def find_company_data(name):
    driver.get('http://www.mca.gov.in/')
    driver.maximize_window()

    pop_up_closing = driver.find_element_by_xpath("//*[@id='imgModalPopup']/div/div/div/p/button/span")
    pop_up_closing.click()

    #For Hovering over menu
    action = ActionChains(driver)

    menu_option1 = driver.find_element_by_xpath("//*[@id='main_nav']/ul/li[5]/a")
    hover1 = action.move_to_element(menu_option1)
    hover1.perform()

    menu_option2 = driver.find_element_by_xpath('//*[@id="main_nav"]/ul/li[5]/ul/li[3]/a')
    hover2 = action.move_to_element(menu_option2)
    hover2.perform()

    view_option = driver.find_element_by_xpath("//*[@id='main_nav']/ul/li[5]/ul/li[3]/ul/li[1]/a")
    view_option.click()

    search_icon = driver.find_element_by_xpath("//div//td//img")
    search_icon.click()
    sleep(2)

    company_name_textbox = driver.find_element_by_xpath("//*[@id='searchcompanyname']")
    company_name_textbox.send_keys(name)

    search_button = driver.find_element_by_xpath("//*[@id='findcindata']")
    search_button.click()
    sleep(2)

    company_path = driver.find_element_by_xpath("//*[@id='cinlist']/tbody/tr[1]/td[1]/a")
    company_path.click()

    #For Page Scrolling
    driver.execute_script("window.scrollTo(0, 120)")
    sleep(1)

    captcha_fill_box = driver.find_element_by_xpath("//*[@id='userEnteredCaptcha']")
    captcha_fill_box.click()
    sleep(8)

    submit_button = driver.find_element_by_xpath("//*[@id='companyLLPMasterData_0']")
    submit_button.click()

    driver.execute_script("window.scrollTo(0, 270)")
    sleep(5)

    rows_path = driver.find_elements_by_xpath("//*[@id='resultTab1']/tbody/tr")
    row_len = len(rows_path)

    cols_path = driver.find_elements_by_xpath("//*[@id='resultTab1']/tbody/tr[1]/td")
    col_len = len(cols_path)

    with open(f'C:\Projects\web-scrapping\data\{name}.txt', 'w') as f:
        for i in range(2, row_len + 1):
            for j in range(1, col_len + 1):
                d = driver.find_element_by_xpath(
                    "//*[@id='resultTab1']/tbody/tr[" + str(i) + "]/td[" + str(j) + "]").text
                f.write(d + "   ")
            f.write("\n")
        print(f'\n==>>FILE SAVED.....{name}')
        print("\nRows :" + str(row_len))
        print("Columns :" + str(col_len))
    driver.close()


if __name__ == '__main__':
    company_name = input("Enter company name : ")
    print("\nFetching Data.........")
    print("REMINDER : Be on the browser to fill the captcha....")
    find_company_data(company_name)
