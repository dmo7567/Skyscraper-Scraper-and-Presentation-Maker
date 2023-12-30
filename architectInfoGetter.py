import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

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
        response = requests.get(href_value)
        soup = BeautifulSoup(response.content, "html.parser")
        preceding_div = soup.find("div", class_="pl-4 subcategory lg:w-1/3 w-1/2 text-lg")
        if preceding_div:
            next_div = preceding_div.find_next_sibling("div")
            h5_element = next_div.find("h5")
            a_element = h5_element.find("a")
            text = a_element.text.strip() if a_element else "??"
        else:
            text = "??"
        rows.append({'Architect': text})

# Create a pandas DataFrame from the extracted data
df = pd.DataFrame(rows)

# Save the DataFrame to an Excel file
df.to_excel('extracted_text.xlsx', index=False)

print("Data successfully extracted and saved to extracted_text.xlsx.")
