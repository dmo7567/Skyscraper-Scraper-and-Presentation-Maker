import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook

# Function to extract image URL
def extract_image_url(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    img_element = soup.find('img', class_='object-cover w-full h-full')
    if img_element:
        return img_element['src']
    else:
        return "No image found"

# Prompt the user to enter a URL
url = input("Enter the URL: ")

# Send a GET request to the URL
response = requests.get(url)

# Create BeautifulSoup object and find the table with the specified ID
soup = BeautifulSoup(response.text, 'html.parser')
table = soup.find('table', id='table-combined-base')

# Extract table rows
rows = []
base_url = "https://www.skyscrapercenter.com/"
for tr in table.find_all('tr'):
    name_element = tr.find('a')
    if name_element:
        href_value = urljoin(base_url, name_element['href'])
        image_url = extract_image_url(href_value)
        rows.append({'Image URL': image_url})

# Create a pandas DataFrame from the extracted data
df = pd.DataFrame(rows)

# Create an Excel workbook
workbook = Workbook()
sheet = workbook.active

# Save the DataFrame to the Excel sheet
for idx, row in df.iterrows():
    sheet.cell(row=idx+1, column=1).value = row['Image URL']

# Save the workbook as an Excel file
workbook.save('image_urls.xlsx')

print("Image URLs successfully extracted and saved to image_urls.xlsx.")
