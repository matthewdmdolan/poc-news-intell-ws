import subprocess
import os
import pandas as pd
from Scraper_classification import categories, companies, countries, firm_categories
import openpyxl
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows

# Set the working directory to the desired directory
directory = r'C:\Users\enisr\OneDrive\Israel\Israel\7 PSR\News scraper'
os.chdir(directory)

# Run Scraper_Paypers.py
subprocess.run(['python', 'Scraper_Paypers.py'], check=True)

# Run Scraper_Finextra.py
subprocess.run(['python', 'Scraper_Finextra.py'], check=True)

print('Both scrapers have finished running')

# Compile Finextra news
finextra_files = [f for f in os.listdir() if f.startswith('news_Finextra_')]
finextra_data = pd.concat([pd.read_excel(file) for file in finextra_files], ignore_index=True)

# Compile Paypers news
paypers_files = [f for f in os.listdir() if f.startswith('news_Paypers_')]
paypers_data = pd.concat([pd.read_excel(file) for file in paypers_files], ignore_index=True)

# Generate the file name with today's date for the combined news
today = pd.to_datetime('today').strftime('%Y-%m-%d')
combined_file = f'News_combined_{today}.xlsx'

# Combine Finextra and Paypers news
combined_data = pd.concat([finextra_data, paypers_data])

# Remove duplicates based on 'Title' column
combined_data.drop_duplicates(subset='Title', inplace=True)

# Add empty columns for Categories, Companies, and Countries
combined_data['Categories'] = ""
combined_data['Companies'] = ""
combined_data['Countries'] = ""
combined_data['Firm 1 category'] = ""
combined_data['Firm 2 category'] = ""
combined_data['Firm 3 category'] = ""
combined_data['Overlap'] = ""

# Function to check for keywords and populate the respective columns
def check_keywords(row):
    title = row['Title']
    categories_list = []
    for category, keywords in categories.items():
        for keyword in keywords:
            if keyword.lower() in title.lower():
                categories_list.append(category)
                break
    row['Categories'] = ', '.join(categories_list)

    firms = set()  # Create a set to store unique firms
    for company, tags in companies.items():
        for tag in tags:
            if tag.lower() in title.lower():
                firms.add(company)  # Use a set to store unique company names

    # Join the unique found firms with a comma and populate the 'Companies' column
    row['Companies'] = ', '.join(firms)

    for i, firm in enumerate(firms):
        firm = firm.strip()
        if firm in firm_categories:
            firm_category = firm_categories[firm]
            if i == 0:
                row['Firm 1 category'] = ', '.join(firm_category)
            elif i == 1:
                row['Firm 2 category'] = ', '.join(firm_category)
            elif i == 2:
                row['Firm 3 category'] = ', '.join(firm_category)

    # Check for repeated firm categories
    firm_categories_list = [row['Firm 1 category'], row['Firm 2 category'], row['Firm 3 category']]
    repeated_categories = [cat for cat in firm_categories_list if firm_categories_list.count(cat) > 1]
    row['Overlap'] = ', '.join(set(repeated_categories))

    for country, names in countries.items():
        for name in names:
            if name.lower() in title.lower():
                row['Countries'] = country
                break

    return row

# Apply the function to populate the columns
combined_data = combined_data.apply(check_keywords, axis=1)

# Convert the 'Date' column to datetime data type
combined_data['Date'] = pd.to_datetime(combined_data['Date'], errors='coerce').dt.date

# Save the combined news data
combined_data.to_excel(combined_file, index=False)
print(f'Combined news saved to {combined_file}')

# Reopen the output file
wb = openpyxl.load_workbook(combined_file)
ws = wb.active

# Iterate over the 'Dates' column
for row in dataframe_to_rows(combined_data, index=False, header=False):
    for cell in row:
        if isinstance(cell, pd.Timestamp):
            # Change the number format to "dd-mm-yyyy"
            cell.number_format = 'dd-mm-yyyy'

# Save the modified data back to the output file
wb.save(combined_file)
wb.close()

print(f'Output file updated: {combined_file}')
