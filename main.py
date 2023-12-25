import contextlib
from playwright.sync_api import  sync_playwright,Page
import time
import os
import pandas as pd

from google.oauth2 import service_account
import gspread
import csv
import glob
# from selectolax.parser import HTMLParser

def upload_data_to_google_sheet():
    scope = ['https://www.googleapis.com/auth/drive']
    service_account_path = os.path.join(os.getcwd(), "helium-10-387902-eb8fa25fda81.json")
    creds = service_account.Credentials.from_service_account_file(service_account_path, scopes=scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open("Arbor Note")
    worksheet = spreadsheet.worksheet('report')
    df = pd.read_excel('result.xlsx',header=None)
    df=df.astype(str)
    values = df.values.tolist()
    worksheet.clear()
    worksheet.update(values)
    print("excel file uploaded successfully to Google Sheets.")



def login(page:Page):
    page.goto("https://app.arbor-note.com/ng/#/login",timeout=100000)
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Username").fill("jaime@gomezltc.com")
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Password").fill("1gj7107A1204")
    page.wait_for_timeout(1000)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_timeout(8000)

def upload_excel(page:Page):
    page.wait_for_timeout(60000)
    with page.expect_download(timeout=60000) as download_info:
        page.wait_for_selector("#Grid_excelexport").click(force=True,timeout=60000)
    download = download_info.value
    download.save_as(os.path.join(os.getcwd(), 'result.xlsx'))
    upload_data_to_google_sheet()
    os.remove(os.path.join(os.getcwd(), 'result.xlsx'))

def Automation_Process():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        login(page)
        page.goto("https://app.arbor-note.com/ng/#/sales/report",timeout=100000)
        upload_excel(page)

def main():
    Automation_Process()
    
    
    
if __name__ == '__main__':
    main()