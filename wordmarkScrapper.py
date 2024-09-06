import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import requests
import time
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# Read input data from Excel
input_file = 'C:\Vicky\Personal Project space\LawStream WebScrapper\input_data.xlsx'  
output_file = 'C:\Vicky\Personal Project space\LawStream WebScrapper\output_data.xlsx'  

input_data = pd.read_excel(input_file)

scraped_data = []

chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')  # Ignore SSL certificate errorss

# Set up Selenium WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)



# Iterate through each row in the input data
for index, row in input_data.iterrows():
    type = row['Type']
    if type == 'Wordmark':
        wordmark = row['Wordmark'] 
        tm_class = str(row['Class']) 

        # Open the website
        driver.get("https://tmrsearch.ipindia.gov.in/tmrpublicsearch/frmmain.aspx")

        # Wait until the input fields are present
        wait = WebDriverWait(driver, 20)

        # Extract cookies from Selenium session to use in the request
        cookies = driver.get_cookies()
        session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}
        print(session_cookies)

        # Define headers for the request
        headers = {
            'accept': 'application/json, text/javascript, */*; q=0.01',
            'accept-encoding': 'gzip, deflate, br, zstd',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json; charset=UTF-8',
            'origin': 'https://tmrsearch.ipindia.gov.in',
            'referer': 'https://tmrsearch.ipindia.gov.in/',
            'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Mobile Safari/537.36',
            'x-requested-with': 'XMLHttpRequest'
        }

        try:
            time.sleep(10)
            # Wait until the Wordmark input is available and then find the element
            wordmark_input = wait.until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_TBWordmark')))
            class_input = driver.find_element(By.ID, 'ContentPlaceHolder1_TBClass')
            captcha_input = driver.find_element(By.ID, 'ContentPlaceHolder1_captcha1')
            

            # Enter the search terms
            wordmark_input.clear()
            class_input.clear()
            wordmark_input.send_keys(wordmark)
            class_input.send_keys(tm_class)

            time.sleep(30)

            # URL for GetCaptcha request
            captcha_url = 'https://tmrsearch.ipindia.gov.in/tmrpublicsearch/frmmain.aspx/GetCaptcha'

            # Path to your downloaded certificate
            certificate_path = 'C:\Vicky\Personal Project space\LawStream WebScrapper\combined_certificates.pem'

            wait.until(EC.presence_of_element_located((By.XPATH, "//img[@title='Captcha Audio']")))

            # Send the POST request to the GetCaptcha endpoint
            response = requests.post(captcha_url, headers=headers, cookies=session_cookies, json={}, verify=certificate_path)

            response.raise_for_status()  # Check for HTTP errors

            # Check if the request was successful
            if response.status_code == 200:
                captcha_solution = response.json()['d']
                print(f"Captcha Solution: {captcha_solution}")
            else:
                print(f"Failed to retrieve CAPTCHA solution. Status Code: {response.status_code}")
                driver.quit()
                exit()
            
            captcha_input.clear()
            captcha_input.send_keys(captcha_solution)

            # Wait for manual CAPTCHA solving or for it to be solved by an automated service
            time.sleep(10)  # Adjust time as necessary

            # Submit the form
            submit_button = driver.find_element(By.XPATH, "//input[@type='submit']")  # Adjust to actual submit button ID
            submit_button.click()

            # Wait for the results to load
            wait.until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_MGVSearchResult')))  # Update this ID with the actual results container ID

            # Scrape the data using BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Example: Finding all rows in the trademark table
            trademark_rows = soup.find_all('td')  # Adjust to actual class or identifier

            if not trademark_rows:
                print("No trademark rows found. Check if the page structure has changed or if the data is loading dynamically.")
                continue

            # Iterate over each row and extract information
            for row in trademark_rows:
                wordmark_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblsimiliarmark_')})
                proprietor_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_LblVProprietorName_')})
                applicationNumber_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblapplicationnumber_')})
                class_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblsearchclass_')})
                status_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_Label6_')})

                # Ensure the span is found before attempting to get text
                if wordmark_span and proprietor_span and applicationNumber_span and class_span and status_span:
                    wordmark_text = wordmark_span.get_text(strip=True)
                    proprietor_text = proprietor_span.get_text(strip=True)
                    applicationNumber_text = applicationNumber_span.get_text(strip=True)
                    class_text = class_span.get_text(strip=True)
                    status_text = status_span.get_text(strip=True)

                    scraped_data.append({'Searched Wordmark': wordmark, 'Class': tm_class, 'Wordmark': wordmark_text, 'Proprietor': proprietor_text, 'Application Number': applicationNumber_text, 'Class': class_text, 'Status': status_text})
                else:
                    print("No wordmark found in this row. Check the row's HTML structure.")
            
        except Exception as e:
            print(f"An error occurred: {e}")
            scraped_data.append({'Wordmark': wordmark, 'Class': tm_class, 'Result': f"Error: {str(e)}"})

    # --------------------------------------This part is under development----------------------------------------
    if type == 'Vienna Code':
        vienna_code = row['Vienna']  # Replace with your column name in Excel
        tm_class = str(row['Class'])  # Replace with your column name in Excel

        # Open the website
        driver.get("https://tmrsearch.ipindia.gov.in/tmrpublicsearch/frmmain.aspx")

        # Wait until the input fields are present
        wait = WebDriverWait(driver, 10)

        try:
            # Wait until the Wordmark input is available and then find the element
            search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//option[@value='VC']")))
            search_input.click()

            time.sleep(5)

            close_checkbox = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='ajax__validatorcallout_innerdiv']")))
            close_checkbox.click()
            time.sleep(5)


            vienna_input = wait.until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_TBVienna')))
            class_input = driver.find_element(By.ID, 'ContentPlaceHolder1_TBClass')

            # Enter the search terms
            vienna_input.clear()
            class_input.clear()
            vienna_input.send_keys(wordmark)
            class_input.send_keys(tm_class)

            # Handle CAPTCHA (you may need to pause here and solve it manually or use a third-party CAPTCHA-solving service)
            print(f"Please solve the CAPTCHA manually for Vienna Code: {vienna_code} and Class: {tm_class}...")

            # Wait for manual CAPTCHA solving or for it to be solved by an automated service
            time.sleep(15)  # Adjust time as necessary

            # Submit the form
            submit_button = driver.find_element(By.XPATH, "//input[@type='submit']")  # Adjust to actual submit button ID
            submit_button.click()

            # Wait for the results to load
            wait.until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_MGVSearchResult')))  # Update this ID with the actual results container ID

            # Scrape the data using BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Example: Finding all rows in the trademark table
            trademark_rows = soup.find_all('tbody')  # Adjust to actual class or identifier

            if not trademark_rows:
                print("No trademark rows found. Check if the page structure has changed or if the data is loading dynamically.")
                continue

            # Iterate over each row and extract information
            for row in trademark_rows:
                wordmark_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblsimiliarmark_')})
                proprietor_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_LblVProprietorName_')})
                applicationNumber_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblapplicationnumber_')})
                class_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_lblsearchclass_')})
                status_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_Label6_')})
                vienna_span = row.find('span', {'id': lambda x: x and x.startswith('ContentPlaceHolder1_MGVSearchResult_LblViennaCode_')})

                # Ensure the span is found before attempting to get text
                if wordmark_span and proprietor_span and applicationNumber_span and class_span and status_span:
                    wordmark_text = wordmark_span.get_text(strip=True)
                    proprietor_text = proprietor_span.get_text(strip=True)
                    applicationNumber_text = applicationNumber_span.get_text(strip=True)
                    class_text = class_span.get_text(strip=True)
                    status_text = status_span.get_text(strip=True)
                    vienna_text = vienna_span.get_text(strip=True)

                    scraped_data.append({'Searched Wordmark': wordmark, 'Class': tm_class, 'Wordmark': wordmark_text, 'Proprietor': proprietor_text, 'Application Number': applicationNumber_text, 'Class': class_text, 'Status': status_text})
                else:
                    print("No wordmark found in this row. Check the row's HTML structure.")
            
        except Exception as e:
            print(f"An error occurred: {e}")
            scraped_data.append({'Wordmark': wordmark, 'Class': tm_class, 'Result': f"Error: {str(e)}"})
    # -------------------------------------------------End of part---------------------------------------------------------

# Close the WebDriver
driver.quit()

# Convert the scraped data to a DataFrame
output_df = pd.DataFrame(scraped_data)
output_df = output_df.drop_duplicates(keep="first")

# Write the output data to a new Excel file using ExcelWriter and openpyxl engine
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    output_df.to_excel(writer, index=False)

print(f"Scraping complete! Results saved to {output_file}.")

