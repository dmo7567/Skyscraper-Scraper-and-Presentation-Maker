import requests
import pandas as pd
from bs4 import BeautifulSoup
import re

# Prompt the user to enter a URL
url = input("Enter the URL: ")

# Send a GET request to the URL
response = requests.get(url)

# Create BeautifulSoup object and find the table with the specified ID
soup = BeautifulSoup(response.text, 'html.parser')
table = soup.find('table', id='table-combined-base')

# Extract table headers
headers = ['#', 'Building Name', 'City', 'Status', 'Completion', 'Height', 'Floors', 'Material', 'Function']

# Extract table rows
rows = []
for tr in table.find_all('tr'):
    row_data = []
    for cell in tr.find_all('td'):
        if 'bg-center' in cell.get('class', ''):
            status_div = cell.find('div', {'data-tippy-content': True})
            status = status_div['data-tippy-content'] if status_div else ''
            if status == '-':
                status = '??'
            row_data.append(status)
        else:
            cell_text = cell.get_text(strip=True)
            if cell_text == '-':
                cell_text = '??'
            row_data.append(cell_text)
    if row_data:
        rows.append(row_data)

# Create a pandas DataFrame from the extracted data
df = pd.DataFrame(rows, columns=headers)

# Remove the unwanted URL column
df = df.drop(columns=['#'])

# Remove non-numerical characters from the Height column, except for decimal point '.'
df['Height'] = df['Height'].apply(lambda x: re.sub(r'[^\d.]', '', x) if x else x)

# Round the Height column to the nearest integer
df['Height'] = df['Height'].apply(lambda x: round(float(x)) if x else x)

# Convert Height from meters to feet and add a new column
df['Height (ft)'] = df['Height'].apply(lambda x: round(float(x) * 3.28084) if x else None)

# Save the DataFrame to an Excel file
df.to_excel('skyscraper_data_modified.xlsx', index=False)

print("Data successfully scraped and saved to skyscraper_data_modified.xlsx.")
