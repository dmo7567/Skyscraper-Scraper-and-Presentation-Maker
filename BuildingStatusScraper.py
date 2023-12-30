import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Prompt the user to enter a URL
url = input("Enter the URL: ")

# Send a GET request to retrieve the webpage content
response = requests.get(url)
html_content = response.content

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, "html.parser")

# Find the table element with the specified ID
table = soup.find("table", id="table-combined-base")

# Create a new Excel workbook and select the active sheet
workbook = Workbook()
sheet = workbook.active

# Find all rows in the table (excluding the header row)
rows = table.find_all("tr")[1:]

# Iterate over each row and extract the status value
for row in rows:
    # Find all div elements within the row
    divs = row.find_all("div", class_=lambda x: x and x.startswith("status-"))

    # Extract the status values from the divs
    status_values = [div.get("data-tippy-content") for div in divs]

    # Add the status values to the Excel spreadsheet
    sheet.append(status_values)

# Save the Excel spreadsheet
workbook.save("status_values.xlsx")
