from selenium import webdriver
import os
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

# Set the working directory to the desired directory
directory = r'C:\Users\enisr\OneDrive\Israel\Israel\7 PSR\News scraper'
os.chdir(directory)

# Check if the combined file exists in the directory
combined_file = 'Scraped_news.xlsx'
if not os.path.exists(combined_file):
    # Create an empty file if it doesn't exist
    open(combined_file, 'w').close()

# Configure Selenium to use a web driver
driver = webdriver.Chrome(r"C:\Users\enisr\OneDrive\Israel\Israel\7 PSR\News scraper\chromedriver.exe")

# Define the URLs to visit
urls = [
    'https://www.finextra.com/latest-news/payments',
    'https://www.finextra.com/latest-news/payments?page=2',
    'https://www.finextra.com/latest-news/payments?page=3',
    'https://www.finextra.com/latest-news/payments?page=4',
    'https://www.finextra.com/latest-news/payments?page=5',
    'https://www.finextra.com/latest-news/payments?page=6',
    'https://www.finextra.com/latest-news/payments?page=7',
    'https://www.finextra.com/latest-news/payments?page=8',
    'https://www.finextra.com/latest-news/payments?page=9'
]

# Define an empty list to store the news
news_list = []

# Loop through the URLs and scrape the news
for url in urls:
    # Navigate to the URL
    driver.get(url)

    # Wait for the news items to load
    driver.implicitly_wait(10)

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all the news items on the page
    news_items = soup.find_all('div', {'class': 'module--story'})

    # Loop over the news items and append them to the news list
    for item in news_items:
        title = item.find('h4', {'class': ''}).find('a').text.strip()
        link = 'https://www.finextra.com' + item.find('h4').find('a')['href']
        date = item.find('span', {'class': 'news-date'}).text.strip() if item.find('span', {'class': 'news-date'}) else ''

        # Append the news item to the news list
        news_list.append({'Title': title, 'Link': link, 'Date': date})

        print(f"Title: {title} - Date: {date}")

# Close the web driver
driver.quit()

# Convert the news list to a pandas DataFrame
df = pd.DataFrame(news_list)

# Generate the file name with today's date
today = datetime.today().strftime("%Y-%m-%d")
file_name = f'news_Finextra_{today}.xlsx'

# Save the DataFrame to an Excel file
if not df.empty:
    df.to_excel(file_name, index=False)
    print(f'Data saved to {file_name}')
else:
    print('No data found')

print('Fin.')
