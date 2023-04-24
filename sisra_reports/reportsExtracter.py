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
from selenium.common.exceptions import NoSuchElementException
from getpass import getpass

username = input("Enter username: ")
password = getpass("Enter password: ")

subjectValue = int(input("What subjectValue are you finding results for? (combined=440, physics=471. chemistry=470, biology=469) "))
total_amount_of_classes = input("How many classes are there? ")


print("")
print("---What data do you need(y/n)---")
year_11_data = input("Year 11 data? ")
year_10_data = input("Year 10 data? ")


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


def change_filters(filter, row_number, year, amount_of_classes, table_start, class_name):
    global filters_button, filter_menu_button, uncheck_all_button, submit_filters_button, counter_for_percentage, amount_of_tasks
    
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
        total_score_box = driver.find_element(By.CSS_SELECTOR, ".headlines:nth-of-type(18) tbody tr:nth-of-type(2) td:nth-of-type(2)")
        print(f"{class_name} {filter}: {total_score_box.text}")
        worksheet.write(row_number + table_start, amount_of_classes, f"{total_score_box.text}") 
        counter_for_percentage+=1

    except NoSuchElementException:
        print(f"{class_name} {filter}: N/A")
        worksheet.write(row_number + table_start, amount_of_classes, f"N/A") 
        counter_for_percentage+=1





def login():
    #load login page
    driver.get('https://www.sisraanalytics.co.uk/Account/Login');
    try:
        user_name_box = driver.find_element(By.CSS_SELECTOR, "#LogIn_UserName")
        password_box = driver.find_element(By.CSS_SELECTOR, "#LogIn_Password")

    except NoSuchElementException:
        print(f"Error: Login inputs not available or HTML data changed")
        exit(1)

    try:
        #type username and password then press submit
        user_name_box.send_keys(username)
        password_box.send_keys(password)
        print("submitted login data")
        password_box.submit()
        time.sleep(4)

    except NoSuchElementException:
        print(f"Error: Login details incorrect or some other error")
        exit(2)

    
    try:
        print("clicking logout of other sessions button")
        submit_button = driver.find_element(By.CSS_SELECTOR, "#fm_LogIn button")
        submit_button.click()

    except NoSuchElementException:
        pass
    
    time.sleep(1)
    
login()



def search_pages(year, table_start):
    global total_amount_of_classes
    #global year
    #driver.get(url);
    amount_of_classes = int(total_amount_of_classes);


    worksheet.write(table_start, 0, f'Year {year}')
    #worksheet.write(table_start, 1, f'Year {year} overall')

    
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
    
    
    counter_for_line = 0
	
    #options_button.click()
    
    while amount_of_classes > 0:
        #global table_start, counter_for_line
        options_tab = driver.find_element(By.CSS_SELECTOR, "[data-tab='option']")
        options_tab.click()
        faculty_selection = driver.find_element(By.CSS_SELECTOR, "[name='ReportOptions.Faculty_ID'] [value='2']")
        faculty_selection.click()	
        qualification_selection = driver.find_element(By.CSS_SELECTOR, f"[name='ReportOptions.Qual_ID'] [value='{subjectValue}']")
        qualification_selection.click()
        class_selection = driver.find_element(By.CSS_SELECTOR, f"[name='ReportOptions.TchGrp_ID'] option:nth-of-type({amount_of_classes})")
        class_name = class_selection.text
        print(f"---{class_name}---")
        class_selection.click()
        time.sleep(0.5)
        
        worksheet.write(table_start, amount_of_classes, f'class: {class_name}')
        
    

             
        def cycle_through_filters():

            #no filter
            try:
                change_filters("none", 1, year, amount_of_classes, table_start, class_name)
            
            except NoSuchElementException:
                print(f"Error: Could not find year {year} data: no filter")
                
            #disadvantaged
            try:
                change_filters("disadvantaged", 2, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data: dis filter")


            
            #non disadvantaged
            try:
                change_filters("non-disadvantaged", 3, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data non dis filter")

            

            #Hpa upper/higher
            try:
                change_filters("hpa", 4, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data upper/higher filter")


            
            #hpa and dis
            try:
                change_filters("hpa/dis", 5, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data hpa and dis filter")



            #SEN e and k
            try:
                change_filters("SEN", 6, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data SEN filter")



            #non SEN
            try:
                change_filters("non-SEN", 7, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data non SEN filter")



            #WBRI ethnic code
            try:
                change_filters("WBRI", 8, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data WBRI filter")



            #AKPN ethnic code
            try:
                change_filters("AKPN", 9, year, amount_of_classes, table_start, class_name)

            except NoSuchElementException:
                print(f"Error: Could not find year {year} data AKPN filter")
            
            

        cycle_through_filters()
        amount_of_classes-=1
        counter_for_line+=1
    
    #print(total_score_box.text)


amount_of_tasks = 0

if year_11_data.lower == "y":
	amount_of_tasks+=1    

if year_10_data.lower == "y":
	amount_of_tasks+=1    

print(amount_of_tasks)
amount_of_tasks*=total_amount_of_classes # multiply by the amount of classes there are
amount_of_tasks*=9 # multiply by the amount of tasks per class

counter_for_percentage = 1




driver.get("https://www.sisraanalytics.co.uk/ReportsHome")
print("On reports page")

if year_11_data.lower() == "y":
	try:
		year_11_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(3):not(.lvrDDL .year)")
		year_11_link.click()
		print("On year 11 section")
		latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_11 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
		latest_assesment.click()
		print("On most recent assessment")
		take_me_to_qualtification_class = driver.find_element(By.CSS_SELECTOR, ".active .toReports")
		take_me_to_qualtification_class.click()
		print("Clicked 'take me to the reports'")

		whole_cohort = driver.find_element(By.CSS_SELECTOR, ".active .EAPRptBtn div:nth-of-type(1) a")
		whole_cohort.click()
		print("Clicked 'Whole Cohort'")
		time.sleep(1)	
		search_pages(11, 0)

	except NoSuchElementException:
		print("Error when looking for year 11")

driver.get("https://www.sisraanalytics.co.uk/ReportsHome")

if year_10_data.lower() == "y":
	try:
		year_10_link = driver.find_element(By.CSS_SELECTOR, ".year:nth-of-type(4):not(.lvrDDL .year)")
		year_10_link.click()
		print("On year 10 section")
		latest_assesment = driver.find_element(By.CSS_SELECTOR, ".pubGrp_10 .eapPub:nth-of-type(1)  .fakea:nth-of-type(1)")
		latest_assesment.click()
		print("On most recent assessment")
		#new
		take_me_to_qualtification_class = driver.find_element(By.CSS_SELECTOR, ".active .toClass")
		take_me_to_qualtification_class.click()
		print("Clicked 'take me to qualification/class'")
		qualification_id = driver.find_element(By.CSS_SELECTOR, ".active #Qual_ID")
		qualification_id.click()
		print("Clicked subjects drop down menu")
		subject_button = driver.find_element(By.CSS_SELECTOR, f'.active [value="{subjectValue}"]') 
		subject_button.click()
		print("Clicked on your subjec")

		time.sleep(1)
		go_button = driver.find_element(By.CSS_SELECTOR, ".active .EAPRptBtn .button")
		driver.execute_script("arguments[0].scrollIntoView();", go_button)
		go_button.click()
		print("Clicked go button 1")
		time.sleep(1)
		#go_button.click()
		#print("Clicked go button 2")
		search_pages(10, 11)

	except NoSuchElementException:
		print("Error when looking for year 10")



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


driver.quit()

exit(3)


