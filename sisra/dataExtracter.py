from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome, ChromeOptions
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
import xlsxwriter
import pathlib


username = input("Enter username: ")
password = input("Enter password: ")


print("")
print("---What data do you need(y/n)---")
year_11_data = input("Year 11 data? ")
year_10_data = input("Year 10 data? ")
year_9_data = input("Year 9 data? ")
year_8_data = input("Year 8 data? ")
year_7_data = input("Year 7 data? ")

opts = ChromeOptions()
opts.add_argument("--window-size=1500,1500")
opts.add_argument("--headless")
driver_path = (pathlib.Path(__file__).parent / 'chromedriver').resolve()
driver = webdriver.Chrome(str(driver_path), options=opts)
driver.implicitly_wait(5) # Just in case
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 30)


def change_filters(filter, row_number, year, amount_of_classes, table_start):
    global filters_button, filter_menu_button, uncheck_all_button, submit_filters_button
    
    filters_button = driver.find_element(By.CSS_SELECTOR, ".filterbuttonholder .tabbutton:nth-of-type(3)")
    filters_button.click()
    
    filter_menu_button = driver.find_element(By.CSS_SELECTOR, ".filters .inp_180")
    filter_menu_button.click()
    
    uncheck_all_button = driver.find_element(By.CSS_SELECTOR, ".modalSubmit #checkClear")
    uncheck_all_button.click()
    
    if filter == "none":
        pass

    elif filter == "disadvantaged":
        disadvantaged_true = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_12_T")
        disadvantaged_true.click()

    elif filter == "non-disadvantaged":
        disadvantaged_false = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_12_F")
        disadvantaged_false.click()

    elif filter == "hpa":
        if year == 10 or year == 11 or year == 7:
            hpa_upper_higher = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_-1_Upper_High")
            
        if year  == 9 or year == 8:
            hpa_upper_higher = driver.find_element(By.CSS_SELECTOR, "[id='ch_-2_Upper_Higher Banding']")
            
        hpa_upper_higher.click()

    elif filter == "hpa/dis":
        if year == 10 or year == 11 or year == 7:
            hpa_upper_higher = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_-1_Upper_High")
            
        if year  == 9 or year == 8:
            hpa_upper_higher = driver.find_element(By.CSS_SELECTOR, "[id='ch_-2_Upper_Higher Banding']")
            
        hpa_upper_higher.click()
        disadvantaged_true = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_12_T")
        disadvantaged_true.click()

    elif filter == "SEN":     
        sen_e = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_7_E")
        driver.execute_script("arguments[0].scrollIntoView();", sen_e)
        sen_e.click()
        sen_k = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_7_K")
        driver.execute_script("arguments[0].scrollIntoView();", sen_k)
        sen_k.click()

    elif filter == "non-SEN":
        non_sen = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_7_N")
        driver.execute_script("arguments[0].scrollIntoView();", non_sen)
        non_sen.click()

    elif filter == "WBRI":
        wbri = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_6_WBRI")
        wbri.click()

    elif filter == "AKPN":
        akpn = driver.find_element(By.CSS_SELECTOR, ".liSelected #ch_6_AKPN")
        akpn.click()
        
    submit_filters_button = driver.find_element(By.CSS_SELECTOR, ".modalSubmit .green")
    submit_filters_button.click()
    
    try:
        total_score_box = driver.find_element(By.CSS_SELECTOR, "tfoot tr td:nth-of-type(13)")
        print(f"{year}_{amount_of_classes} {filter}: {total_score_box.text}")
        worksheet.write(row_number + table_start, amount_of_classes + 1, f"{total_score_box.text}") 

    except NoSuchElementException:
        print(f"{year}_{amount_of_classes} {filter}: N/A")
        worksheet.write(row_number + table_start, amount_of_classes + 1, f"N/A") 





def login():
    #load login page
    driver.get('https://www.sisraanalytics.co.uk/Account/Login');

    user_name_box = driver.find_element(By.CSS_SELECTOR, "#LogIn_UserName")
    password_box = driver.find_element(By.CSS_SELECTOR, "#LogIn_Password")

    #type username and password then press submit
    user_name_box.send_keys(username)
    password_box.send_keys(password)
    print("submitted login data")
    password_box.submit()
    time.sleep(4)
    
    try:
        print("clicking logout of other sessions button")
        submit_button = driver.find_element(By.CSS_SELECTOR, "#fm_LogIn button")
        submit_button.click()

    except NoSuchElementException:
        pass
    
    time.sleep(1)
    
login()



def search_pages(year, table_start):
    #global year
    #driver.get(url);
    amount_of_classes = 5;


    worksheet.write(table_start, 0, f'Year {year}')
    worksheet.write(table_start, 1, f'Year {year} overall')
    worksheet.write(table_start, 2, f'{year}-1')
    worksheet.write(table_start, 3, f'{year}-2')
    worksheet.write(table_start, 4, f'{year}-3')
    worksheet.write(table_start, 5, f'{year}-4')
    worksheet.write(table_start, 6, f'{year}-5')
    
    worksheet.write(table_start + 1, 0, "Progress VA")
    worksheet.write(table_start + 2, 0, "Dis")
    worksheet.write(table_start + 3, 0, "Non-dis")
    worksheet.write(table_start + 4, 0, "HPA")
    worksheet.write(table_start + 5, 0, "HPA/Dis")
    worksheet.write(table_start + 6, 0, "SEND")
    worksheet.write(table_start + 7, 0, "Non-SEND")
    worksheet.write(table_start + 8, 0, "WBRJ")
    worksheet.write(table_start + 9, 0, "AKPN")
    
    #body_var = driver.find_element(By.CSS_SELECTOR, "body")
    #submit()
    webdriver.ActionChains(driver).send_keys(Keys.RETURN).perform()
    time.sleep(3)
    
    
    

    #options_button.click()
    
    while amount_of_classes >= 0:
        class_changing_box = driver.find_element(By.CSS_SELECTOR, f"#ReportOptions_TchGrp_ID option:nth-of-type({amount_of_classes+1})")
        options_button = driver.find_element(By.CSS_SELECTOR, ".filterbuttonholder .tabbutton:nth-of-type(2)")
        options_button.click()
        time.sleep(3)
        class_changing_box.click()
        time.sleep(5)

        
        def cycle_through_filters():

            #no filter
            change_filters("none", 1, year, amount_of_classes, table_start)

                
            #disadvantaged
            change_filters("disadvantaged", 2, year, amount_of_classes, table_start)

            
            #non disadvantaged
            change_filters("non-disadvantaged", 3, year, amount_of_classes, table_start)
            

            #Hpa upper/higher
            change_filters("hpa", 4, year, amount_of_classes, table_start)

            
            #hpa and dis
            change_filters("hpa/dis", 5, year, amount_of_classes, table_start)


            #SEN e and k
            change_filters("SEN", 6, year, amount_of_classes, table_start)


            #non SEN
            change_filters("non-SEN", 7, year, amount_of_classes, table_start)


            #WBRI ethnic code
            change_filters("WBRI", 8, year, amount_of_classes, table_start)


            #AKPN ethnic code
            change_filters("AKPN", 9, year, amount_of_classes, table_start)

        cycle_through_filters()
        amount_of_classes-=1
    
    
    #print(total_score_box.text)


    



driver.get("https://www.sisraanalytics.co.uk/ReportsHome")

if year_11_data.lower() == "y":
    year_11_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(3):not(.lvrDDL .year)")
    year_11_link.click()
    latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_11 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
    latest_assesment.click()
    go_button = driver.find_element(By.CSS_SELECTOR, ".toClass6599 .button")
    go_button.click()
    search_pages(11, 0)

driver.get("https://www.sisraanalytics.co.uk/ReportsHome")

if year_10_data.lower() == "y":
    year_10_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(4):not(.lvrDDL .year")
    year_10_link.click()
    latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_10 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
    latest_assesment.click()
    go_button = driver.find_element(By.CSS_SELECTOR, ".toClass65101 .button")
    go_button.click()
    search_pages(10, 11)

driver.get("https://www.sisraanalytics.co.uk/ReportsHome")

if year_9_data.lower() == "y":
    year_9_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(5):not(.lvrDDL .year")
    year_9_link.click()
    latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_9 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
    latest_assesment.click()
    go_button = driver.find_element(By.CSS_SELECTOR, ".toClass65102 .button")
    go_button.click()
    search_pages(9, 22)

driver.get("https://www.sisraanalytics.co.uk/ReportsHome")

if year_8_data.lower() == "y":
    year_8_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(6):not(.lvrDDL .year")
    year_8_link.click()
    latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_8 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
    latest_assesment.click()
    go_button = driver.find_element(By.CSS_SELECTOR, ".toClass65103 .button")
    go_button.click()
    search_pages(8, 33)

driver.get("https://www.sisraanalytics.co.uk/ReportsHome")
    
if year_7_data.lower() == "y":
    year_7_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(7):not(.lvrDDL .year")
    year_7_link.click()
    latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_7 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
    latest_assesment.click()
    go_button = driver.find_element(By.CSS_SELECTOR, ".toClass65104 .button")
    go_button.click()
    search_pages(7, 44)


'''
if year_11_url.lower() != "n/a" and year_11_url.lower() != "n" :
    search_pages(year_11_url, 11, 0)

if year_10_url.lower() != "n/a" and year_10_url.lower() != "n" :
    search_pages(year_10_url, 10, 11)

if year_9_url.lower() != "n/a" and year_9_url.lower() != "n" :
    search_pages(year_9_url, 9, 22)

if year_8_url.lower() != "n/a" and year_8_url.lower() != "n" :
    search_pages(year_8_url, 8, 33)

if year_7_url.lower() != "n/a" and year_7_url.lower() != "n" :
    search_pages(year_7_url, 7, 44)
'''

workbook.close()
print("workbook closed")

time.sleep(99999999)

driver.quit()




