# Import dependencies
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
import xlsxwriter

URL = 'https://uk.indeed.com'
# Creating a webdriver instance
service = Service(executable_path="C:\Development\chromedriver.exe")
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=service, options=options)

job_list = ["Data Architect", " Big data architect", "Data scientist ", "Azure consultant "]
loc = 'United Kingdom'
location = True

# open webdriver for each job name
for job in job_list:

    result = []
    driver.get(URL)
    sleep(1)
    job_title = driver.find_element(By.NAME, 'q')
    job_title.send_keys(job)
    if location:
        location = driver.find_element(By.NAME, 'l')
        location.send_keys(loc)
    else:
        None
    find_job_button = driver.find_element(By.XPATH, "//button[@class='yosegi-InlineWhatWhere-primaryButton']")
    find_job_button.click()
    sleep(2)

    # Reject cookies
    try:
        reject_button = driver.find_element(By.XPATH, "//button[@id='onetrust-reject-all-handler']")
        reject_button.click()
        close_button = driver.find_element(By.XPATH, "//button[@class='icl-CloseButton icl-Card-close']")
        close_button.click()
    except:
        None
    sleep(1)
    # retrieve the html data
    page_source = driver.page_source

    # soup the html
    soup = BeautifulSoup(page_source, 'html.parser')

    # scraping process
    # take all the company that have web page
    links = soup.find_all('a', 'turnstileLink companyOverviewLink')
    link_list = []
    for link in links:
        link_list.append(URL+link['href'])

    for link in link_list:
        driver.get(link)
        sleep(1)
        try:
            company_name = driver.find_element(By.XPATH, "//div[@itemprop='name']").text
        except:
            company_name = None
        try:
            company_city = driver.find_element(By.XPATH, "//span[@class='css-smaipe e1wnkr790']").text
        except:
            company_city = None
        try:
            website = driver.find_element(By.XPATH, "//a[@data-tn-element='companyLink[]']").get_attribute('href')
        except:
            website = None
        try:
            ceo_name = driver.find_element(By.XPATH, "//span[@class='css-1w0iwyp e1wnkr790']").text
        except:
            ceo_name = None

        data = {
            "Company name": company_name,
            "Company City": company_city,
            "Website": website,
            "CEO or CTO or Financial manager Name": ceo_name,
            "Contact person email": None,
            "Contact person mobile number": None,
            "Company email": None,
            "Company mobile number": None,
        }
        result.append(data)
    print(job, result)
    # Create workbook
    workbook = xlsxwriter.Workbook(f"{job}_indeed.com.xlsx")
    # Create Excel sheet for each job name
    worksheet = workbook.add_worksheet(job)
    worksheet.write(0, 0, 'Company name')
    worksheet.write(0, 1, 'Company City')
    worksheet.write(0, 2, 'Website')
    worksheet.write(0, 3, 'CEO or CTO or Financial manager Name')
    worksheet.write(0, 4, 'Contact person email')
    worksheet.write(0, 5, 'Contact person mobile number')
    worksheet.write(0, 6, 'Company email')
    worksheet.write(0, 7, 'Company mobile number')

    for index, entry in enumerate(result):
        worksheet.write(index+1, 0, entry["Company name"])
        worksheet.write(index+1, 1, entry["Company City"])
        worksheet.write(index+1, 2, entry["Website"])
        worksheet.write(index+1, 3, entry["CEO or CTO or Financial manager Name"])
        worksheet.write(index+1, 4, entry["Contact person email"])
        worksheet.write(index+1, 5, entry["Contact person mobile number"])
        worksheet.write(index+1, 6, entry["Company email"])
        worksheet.write(index+1, 7, entry["Company mobile number"])
    workbook.close()
    location = False

