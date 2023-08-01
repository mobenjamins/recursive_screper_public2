import os
import sys
from urllib.parse import urlparse
import itertools
import json
import difflib
import requests
from datetime import datetime

from PyPDF2 import PdfReader
from io import BytesIO

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from base_script import *

# init a driver instance
driver = init_driver()

def read_data_from_sheet(spreadsheet_url, sheet_index):
    print(f'Reading data from sheet number {sheet_index+1} ...')
    
    DIR = os.path.dirname(os.path.realpath(__file__))
    parent_dir = os.path.dirname(DIR)
    root_dir = os.path.dirname(parent_dir)
    sys.path.append(parent_dir)

    # Get the JSON credentials file for the service account
    credentials_path = 'lucid-cocoa-375621-2dc04e9671cb.json'
    with open(credentials_path, "r") as file:
        secrets = json.load(file)

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(secrets, ["https://www.googleapis.com/auth/spreadsheets"])

    # Connect to the Google Sheets API using the credentials
    client = gspread.authorize(credentials)

    # Open the spreadsheet using the URL
    spreadsheet = client.open_by_url(spreadsheet_url)

    # Get the desired worksheet in the spreadsheet
    worksheet = spreadsheet.get_worksheet(sheet_index)

    # Now worksheet is defined, you can use it to extract values
    column_b_values = worksheet.col_values(2)[1:]  # Exclude the header
    column_d_values = worksheet.col_values(4)[1:]
    column_e_values = worksheet.col_values(5)[1:]
    column_f_values = worksheet.col_values(6)[1:]
    column_g_values = worksheet.col_values(7)[1:]

    
    # Create a list of dictionaries
    data_list = []
    for b, d, e, f, g in itertools.zip_longest(column_b_values, column_d_values, column_e_values, column_f_values, column_g_values, fillvalue=""):
        data_dict = {b: [d, e, f, g]}
        data_list.append(data_dict)
    
    cleaned_data_list = []
    
    for data_dict in data_list:
        cleaned_dict = {}
        for key, phrases in data_dict.items():
            cleaned_phrases = [phrase.replace('\n', ' ') for phrase in phrases]
            cleaned_dict[key] = cleaned_phrases
        cleaned_data_list.append(cleaned_dict)

    return cleaned_data_list

def check_changes(old_data, new_data, url):
    if url in old_data and old_data[url] != new_data:
        print(f'CHANGED ==> {url}')
        old_data[url] = new_data  # Replace the old data
    else:
        print('SAME :)')
    
def update_check_date(spreadsheet_url, sheet_index, row, date_format="%Y-%m-%d"):
    # Get the date as a string in the provided format
    date_str = datetime.now().strftime(date_format)

    # Connect to the Google Sheets API
    DIR = os.path.dirname(os.path.realpath(__file__))
    parent_dir = os.path.dirname(DIR)
    root_dir = os.path.dirname(parent_dir)
    sys.path.append(parent_dir)
    credentials_path = 'lucid-cocoa-375621-2dc04e9671cb.json'
    with open(credentials_path, "r") as file:
        secrets = json.load(file)
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(secrets, ["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(credentials)

    # Open the spreadsheet and get the worksheet
    spreadsheet = client.open_by_url(spreadsheet_url)
    worksheet = spreadsheet.get_worksheet(sheet_index)

    # Find the first column labeled "Checked", or if not present, the first empty column
    headers = worksheet.row_values(1)  # Assuming first row is the headers
    try:
        column_index = headers.index("Checked") + 1  # Add one due to 1-indexing in gspread
    except ValueError:
        # "Checked" not found, find first empty column
        column_index = headers.index("") + 1 if "" in headers else len(headers) + 1

    # If a new column is created, update the header to "Checked"
    if column_index == len(headers) + 1:
        worksheet.update_cell(1, column_index, "Checked")  # Update the header

    # Update the cell in the chosen column for the given row
    worksheet.update_cell(row + 2, column_index, date_str)  # row + 2 to account for header row and 1-indexing

    
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1NR_EH89z9akyjJWslGYrUXN_gt3pxkbzFhW7FduVCdI/edit#gid=0'

# load the previous data from json file
old_data = {}
if os.path.exists('data.json'):
    with open('data.json', 'r') as f:
        old_data = json.load(f)
        
for sheet_index in range(0, 9):
    data_list = read_data_from_sheet(spreadsheet_url, sheet_index)
    
    for line in data_list:
        url = list(line.keys())[0]
        print(f'\n[{data_list.index(line) + 1}] Working on {url} ...')
        
        data = {} # dictionary to hold url and associated text
        
        if '.pdf' in url[-4:]:
            # Send a GET request to the URL
            response = requests.get(url)

            # Create a BytesIO object from the content of the response
            file = BytesIO(response.content)

            # Create a PDF reader object
            pdf = PdfReader(file)

            # Initialize an empty string for the content of the PDF
            body_text = ""

            # Iterate over the pages in the PDF and extract the text
            for i in range(len(pdf.pages)):
                body_text += pdf.pages[i].extract_text()
            
            # check if the data is different
            check_changes(old_data, body_text, url)
            
        else:
            try:
                # access the url
                go_to(driver, url)
            except:
                continue
            
            try:
                # get text from the first page
                body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                data[url] = body_text

                # check if the data is different
                check_changes(old_data, body_text, url)
                    
                # other pages links
                domain = urlparse(url).netloc # extract the domain of the main url
                elements = driver.find_elements(By.TAG_NAME, "a") # find all the anchor tags
                try:
                    same_domain_urls = [el.get_attribute("href") for el in elements if urlparse(el.get_attribute("href")).netloc == domain] # extract urls that have the same domain
                    same_domain_urls = list(set(same_domain_urls))

                    # get text from other pages
                    for each_href in same_domain_urls:
                        go_to(driver, each_href)
                        body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                        data[each_href] = body_text
                        
                        # check if the data is different
                        check_changes(old_data, body_text, each_href)
                            
                        if same_domain_urls.index(each_href) == 30:
                            break
                
                except StaleElementReferenceException:
                    pass
            
            except NoSuchElementException:
                pass
        
        update_check_date(spreadsheet_url, sheet_index, data_list.index(line), date_format="%Y-%m-%d")
        
# save data to json file after all sheets have been processed
with open('data.json', 'w') as f:
    json.dump(old_data, f, indent=4)
