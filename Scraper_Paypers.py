from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import time
import os

# Set the working directory to the desired directory
directory = r'C:\Users\enisr\OneDrive\Israel\Israel\7 PSR\News scraper'
os.chdir(directory)

# Configure Selenium to use a web driver
driver = webdriver.Chrome(r"C:\Users\enisr\OneDrive\Israel\Israel\7 PSR\News scraper\chromedriver.exe")

# Define the base URL
base_url = 'https://thepaypers.com/news/all'

# Navigate to the base URL
driver.get(base_url)

# Find the maximum page number
page_links = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div#ctl00_MainPlaceHolder_pageLinks a'))
)
max_page = int(page_links[-2].text.strip())

# Lists to store the data
titles = []
categories = []
dates = []
locations = []
article_links = []  # List to store the links to news articles

# Loop through the page numbers and scrape data from each page
for page_number in range(1, min(max_page, 10) + 1):
    # Wait for the page to load
    time.sleep(2)  # Adjust the duration as needed
    
    # Get the HTML content of the page
    page_html = driver.page_source
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_html, 'html.parser')
    
    # Find all the news items on the page
    news_items = soup.find_all('td')
    
    # Loop over the news items and extract the data
    for item in news_items:
        # Extract the title and link to the article
        title_element = item.find('h3')
        if title_element:
            title = title_element.text.strip()
            article_link = 'https://thepaypers.com' + title_element.find('a')['href']  # Extract the article link
            titles.append(title)
            article_links.append(article_link)
        
        # Extract the category, date, and location
        info_element = item.find('span', class_='source')
        if info_element:
            info_text = info_element.text.strip()
            parts = info_text.split('|')
            if len(parts) == 3:
                category = parts[0].strip()
                date = parts[1].strip()
                location = parts[2].strip()
                categories.append(category)
                dates.append(date)
                locations.append(location)
    
    # Click on the numbered button for the next page if not on the last page
    if page_number < max_page:
        next_button = driver.find_element(By.ID, f'ctl00_MainPlaceHolder_page{page_number+1:02d}Nav')
        driver.execute_script("arguments[0].click();", next_button)
    
    # Wait for the next page to load
    time.sleep(1)  

# Close the web driver
driver.quit()

# Generate the file name with today's date
today = datetime.today().strftime("%Y-%m-%d")
file_name = f'news_Paypers_{today}.xlsx'

# Create a DataFrame from the data
df = pd.DataFrame({'Title': titles, 'Category': categories, 'Date': dates, 'Location': locations, 'Link': article_links})

# Save the DataFrame to an Excel file
df.to_excel(file_name, index=False)

print(f'Data saved to {file_name}')
