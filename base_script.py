import json
import os
import pprint
import random
import requests
import smtplib
import sys
import time
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
import string
import socket
import glob

from PIL import Image, ImageFilter
from anticaptchaofficial.recaptchav2proxyless import *
from bs4 import BeautifulSoup
from butler import Client
from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.common.exceptions import (WebDriverException, ElementNotInteractableException, NoSuchElementException, NoSuchWindowException, StaleElementReferenceException, TimeoutException, ElementNotVisibleException, InsecureCertificateException, JavascriptException)
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from twocaptcha import TwoCaptcha
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.proxy import Proxy, ProxyType
from urllib3.exceptions import MaxRetryError
from selenium.webdriver.chrome.service import Service

import gspread
import openpyxl
import pytesseract
from mailtm import Email
from time import sleep

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import email

# New modules integration
import os
import socket
from selenium.webdriver.common.proxy import Proxy, ProxyType

from fillpdf import fillpdfs

import platform ; headless = platform.system() != "Windows"

x = {
    'First name': 'John',
    'Last name': 'Smith',
    'Company': 'ABC Corporation',
    'Telephone number': '5551234456',
    'Email': 'amine.elasri@thecozm.com',
    'Address 1': '123 Main Street',
    'Address 2': 'Apt 4B',
    'City': 'Anytown',
    'Postcode': 'PO16 7GZ',
    'Country': 'United States',
    'Day of birth (DD)': '1',
    'Month of birth (MM)': '1',
    'Year of birth (YYYY)': '1980',
    'Place of birth': 'Anytown',
    'Nationality': 'American',
    'National Insurance Number': 'QQ123456A',
}

# Classes for custom exceptions ####### EXPERIMENTAL ###################################################
class DateOfBirthError(Exception):
    pass

class TimeoutError(Exception):
    pass
#############################################################################################

# Proof function
def proof(driver, country, debug=False, send=False):
    screenshot_name = f'screenshot_{country}.png'
    driver.save_screenshot(screenshot_name) ; sleep(2)
    
    if debug == False:
        reason = 'Automation Proof'
    elif debug == True:
        reason = 'Debug'
    
    # Get the current date and time & Format it as a string
    now = datetime.now()
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
    
    if send == True:
        try:
            send_email('amine.elasri@thecozm.com', f'{country} {reason}', f'{timestamp_str}', [screenshot_name])
        except (WebDriverException, MaxRetryError) as e:
            pass
    return screenshot_name
                
# Password generation function
def generate_password(min_length=12, max_length=15):
    length = random.randint(min_length, max_length)
    chars = string.ascii_letters + string.digits + string.punctuation
    password = ''.join(random.choice(chars) for _ in range(length))
    return password

# Email sending function
def send_email(to_address, subject, body, attachments=None):
    from_address = "amine.elasri@thecozm.com"

    # Wrap plain text lines in <p> tags
    body_lines = body.split('\n')
    body_lines = [f"<p>{line.strip()}</p>" if line.strip() else line for line in body_lines]
    wrapped_body = "\n".join(body_lines)

    # Create message container - the correct MIME type is multipart/related.
    msg = MIMEMultipart('related')
    msg['Subject'] = subject
    msg['From'] = from_address
    msg['To'] = to_address

    password = 'gjbpnwjxhcailzby'

    # Attach HTML body part into message container
    wrapped_body = f"""
        <html>
        <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .container {{ max-width: 600px; margin: auto; }}
            .header {{ background-color: #f1f1f1; padding: 20px; }}
            .content {{ padding: 20px; }}
            .footer {{ background-color: #f1f1f1; padding: 10px; text-align: center; }}
        </style>
        </head>
        <body>
        <div class="container">
        
            <div class="content">
                {wrapped_body}
            </div>
            
            <div class="footer">
                <img src="cid:logo" alt="Company Logo" style="max-width: 100px;"/>
                
                <p>&copy; {datetime.now().year} The Cozm. All rights reserved.</p>
            </div>
            
        </div>
        </body>
        </html>
    """
    
    msg.attach(MIMEText(wrapped_body, 'html'))
    
    # Attach the company logo as an inline image
    thecozm_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cozm_logo.png")
    with open(thecozm_logo, "rb") as logo_file:
        logo_image = MIMEImage(logo_file.read())
        logo_image.add_header('Content-ID', '<logo>')
        logo_image.add_header('Content-Disposition', 'inline', filename="The Cozm Logo")  # Add this line
        msg.attach(logo_image)

    # Attach files if any
    if attachments:
        for file_path in attachments:
            with open(file_path, 'rb') as file:
                file_name = os.path.basename(file_path)
                file_attachment = MIMEBase('application', 'octet-stream')
                file_attachment.set_payload(file.read())

            # Encode the file attachment and add headers
            encoders.encode_base64(file_attachment)
            file_attachment.add_header('Content-Disposition', f'attachment; filename={file_name}')
            msg.attach(file_attachment)

    # Sending the email using a context manager
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo()
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_address, msg.as_string())

# Date formatting function
def format_date(day, month, year, date_format):
    # Convert day, month, and year to integers
    day = int(day)
    month = int(month)
    year = int(year)
    
    # Create a datetime object using the provided day, month, and year
    date_obj = datetime(year, month, day)
    
    # Format the date using strftime() with the given format
    formatted_date = date_obj.strftime(date_format)
    
    return formatted_date

# Free Port Finder function
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def init_driver(headless=False, proxy=False):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("disable-infobars")
    #options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('ignore-certificate-errors')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--window-size=1920,1080')
    
    download_path = os.path.join(os.path.expanduser("~"), "Downloads")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_path
    })
    
    if headless:
        options.add_argument("--headless")
    #options.add_extension(r'configs/ReCaptcha-Solver.crx')
        
    # Print the download directory
    #print("Download directory:", download_path)

    port = find_free_port()

    service = Service(executable_path=ChromeDriverManager().install())
    service.port = port
    service.start()

    if proxy == False:
        driver = webdriver.Chrome(service=service, options=options)
    else:
        # proxy feature
        prox = Proxy()
        prox.proxy_type = ProxyType.MANUAL
        prox.http_proxy = proxy

        capabilities = webdriver.DesiredCapabilities.CHROME
        prox.add_to_capabilities(capabilities)

        driver = webdriver.Chrome(service=service, options=options, desired_capabilities=capabilities)
    
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_path}}
    command_result = driver.execute("send_command", params)
    
    options.add_experimental_option("prefs", {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    return driver

# Select from dropdown function
def select_actionchains(driver, id, value):
    # adding ability to wait
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.element_to_be_clickable((By.ID, id)))
    
    # Create an ActionChains object
    actions = ActionChains(driver)  
    dropdown = driver.find_element(By.ID, id)
    actions.move_to_element(dropdown).perform()
    actions.click(dropdown).perform()
    
    select = Select(dropdown)
    sleep(1)
    try:
        select.select_by_visible_text(value)
    except:
        select.select_by_index(1)

def wait(driver, xpath):
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))

# Write function
def write(driver, value, element, timeout=10):
    wait = WebDriverWait(driver, timeout)

    if element.startswith('#'): # Write to an element by id
        element_id = element[1:]
        wait.until(EC.element_to_be_clickable((By.ID, element_id)))
        element = driver.find_element(By.ID, element_id)
        
    elif element.startswith('.'): # Write to an element by class name
        element_class = element[1:]
        if ' ' in element_class:
            element_class = element_class.replace(' ', '.')
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, element_class)))
        element = driver.find_element(By.CLASS_NAME, element_class)
    
    elif element.startswith('/'): # Write to an element by xpath
        element_xpath = element
        wait.until(EC.element_to_be_clickable((By.XPATH, element_xpath)))
        element = driver.find_element(By.XPATH, element_xpath)

    elif element.startswith('@'): # Write to an element by name
        element_name = element[1:]
        wait.until(EC.element_to_be_clickable((By.NAME, element_name)))
        element = driver.find_element(By.NAME, element_name)

    else:
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//input[contains(text(), "{element}")]')))
        element = driver.find_element(By.XPATH, f'//input[contains(text(), "{element}")]')

    actions = ActionChains(driver)
    #actions.click(element).send_keys(value).perform()
    actions.move_to_element(element).click()
    element.clear()
    actions.send_keys(value).perform()

def click(driver, element, timeout=10):
    wait = WebDriverWait(driver, timeout)
    
    if element.startswith('#'): 
        element_id = element[1:]
        button_wait = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
        element = driver.find_element(By.ID, element_id)
        
    elif element.startswith('.'): 
        element_class = element[1:]
        if ' ' in element_class:
            element_class = element_class.replace(' ', '.')
        button_wait = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, element_class)))
        element = driver.find_element(By.CLASS_NAME, element_class)
    
    elif element.startswith('/'):
        element_xpath = element
        button_wait = wait.until(EC.element_to_be_clickable((By.XPATH, element_xpath)))
        element = driver.find_element(By.XPATH, element_xpath)

    elif element.startswith('@'): # Write to an element by name
        element_name = element[1:]
        button_wait = wait.until(EC.element_to_be_clickable((By.NAME, element_name)))
        element = driver.find_element(By.NAME, element_name)

    else:
        try:
            button_wait = wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[text()='{element}']")))
            element = driver.find_element(By.XPATH, f"//*[text()='{element}']")
        except:
            radio_button = driver.find_element(By.XPATH, "//label[text()='" + element + "']/preceding-sibling::input[@type='radio' or contains(text(), '" + element + "')]")
            radio_button.click()
    
    actions = ActionChains(driver)
    try:
        actions.move_to_element(element).click().perform()
    except:
        try:
            element.click()
        except:
            pass

# Wait until a text is present Function
def wait_until(driver, text, timeout=20):
    start_time = time.time()
    while time.time() < start_time + timeout:
        if text in driver.page_source:
            return True
    return False

def go_to(driver, website, max_retries=3, sleep_time=1):
    for i in range(max_retries):
        try:
            driver.get(website)
            return
        except (WebDriverException, MaxRetryError) as e:
            print(f"Attempt {i+1} failed. Retrying in {sleep_time} seconds...")
            time.sleep(sleep_time)
    raise Exception(f"Unable to reach {website} after {max_retries} attempts.")
