# Trademark Information Web Scraper

This project is a web scraper designed to automate the process of extracting trademark information from the Indian Trademark Registry website (https://tmrsearch.ipindia.gov.in). The scraper navigates through the website, solves CAPTCHA challenges, and extracts data based on user-specified search criteria.


### Project Description: Trademark Information Web Scraper

This project involves developing a Python-based web scraper to automate the extraction of trademark information from the official website of the Indian Trademark Registry. The scraper is designed to handle both "Wordmark" and "Vienna Code" (Under development) searches and is capable of bypassing CAPTCHAs through a manual intervention step. The data extracted from the website is saved in an Excel file for further analysis.

### Key Features
**Input Handling**: The script reads input data from an Excel file (input_data.xlsx) that contains details about trademarks to be searched, including the type of search (Wordmark or Vienna Code), the trademark itself, and its associated class.

**Web Automation**: The scraper uses Selenium WebDriver to automate the browser and interact with the web elements, such as input fields and buttons, to perform searches on the website.

**Captcha Handling**: The script addresses CAPTCHA challenges by making a POST request to the website’s CAPTCHA service and then waits for manual input to solve it.

**Data Extraction**: Once the search results are loaded, the script uses BeautifulSoup to parse the HTML and extract relevant trademark information such as Wordmark, Proprietor, Application Number, Class, Status, etc.

**Error Handling and Logging**: The scraper is equipped with error handling mechanisms to catch exceptions and log errors for failed searches or changes in website structure.

**Output Generation**: The extracted data is written to an output Excel file (output_data.xlsx), providing a structured format for further use.

**Dynamic Content Handling**: The scraper waits for dynamically loaded content using Selenium's WebDriverWait and Expected Conditions, ensuring the data is fully loaded before extraction.

## Installation

### Prerequisites

- Python 3.9.19 or higher
- ChromeDriver compatible with your version of Google Chrome
- [pip](https://pip.pypa.io/en/stable/installation/) for managing Python packages

### Setup Instructions

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/Oblivion254/Trade-mark-WebScrapper.git
   cd trademark-webscraper
2. **Install Requirements.txt**
    ```bash
    pip install -r requirements.txt
3. **Prepare Input Data**:
    Ensure the input_data.xlsx file is present in the root directory of the project. This file should contain columns for "Type", "Wordmark", "Class", and other relevant fields.
4. **Run the Scraper**:
    Execute the Python script to start the scraping process:
    ```bash
    python wordmarkScrapper.py
5. **Review Output Data**:
    The scraped data will be saved in output_data.xlsx in the root directory.


### Usage Notes
Make sure to have the latest version of Chrome and the appropriate ChromeDriver installed.
Ensure a stable internet connection to access the website and perform scraping.
The project assumes that the CAPTCHA can be manually solved or through a third-party service.

### Future Improvements
Implement automatic CAPTCHA solving using third-party services.
Add support for scraping additional trademark details or other related information.
Implement a more robust error-handling mechanism for dynamically changing web elements.

### Contact
Feel free to contact Vignesh Mani @ vickyaadhi254@gmail.com for further inquiries.
