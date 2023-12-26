from playwright.sync_api import  sync_playwright,Page
import os
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
import gspread
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError



def upload_sheet_pdf(file_name):
    scope = ['https://www.googleapis.com/auth/drive']
    service_account_path = os.path.join(os.getcwd(), "helium-10-387902-eb8fa25fda81.json")
    creds = service_account.Credentials.from_service_account_file(service_account_path, scopes=scope)
    FOLDER_ID = '1aXNw0CQhAY2ObP2bPQOjuyq27co4FSTc'
    PDF_FILE_PATH = file_name
    drive_service = build('drive', 'v3', credentials=creds)
    # Create a MediaFileUpload object for the PDF file
    file_metadata = {'name': os.path.basename(PDF_FILE_PATH),
                    'parents': [FOLDER_ID]}
    media = MediaFileUpload(PDF_FILE_PATH, mimetype='application/pdf')
    # Upload the file to the folder
    try:
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"File ID: {file.get('id')}")
    except HttpError as error:
        print(f"An error occurred: {error}")

def upload_data_to_google_sheet():
    scope = ['https://www.googleapis.com/auth/drive']
    service_account_path = os.path.join(os.getcwd(), "helium-10-387902-eb8fa25fda81.json")
    creds = service_account.Credentials.from_service_account_file(service_account_path, scopes=scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open("Arbor Note")
    worksheet = spreadsheet.worksheet('report')
    df = pd.read_excel('result.xlsx',header=None)
    df.fillna('', inplace=True)
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

def upload_pdfs(page:Page):
    table_rows=page.query_selector_all("colgroup#content-GridcolGroup + tbody > tr")
    for row in table_rows:
        date=row.query_selector("td:nth-child(1)").text_content()
        if date!=datetime.today().strftime('%m/%d/%Y'):
            break
        proposal_number=row.query_selector("td:nth-child(2)").text_content()
        project_name=row.query_selector("td:nth-child(5)").text_content()
        with page.expect_popup() as page2_info:
            row.query_selector("td:nth-child(2) a").click()
        page2 = page2_info.value
        page2.wait_for_timeout(5000)
        with page2.expect_download(timeout=60000) as download_info:
            page2.query_selector("body > div > div > div:nth-child(1) > div > div > div:nth-child(1) > div:nth-child(1) > a").click(modifiers=['Alt'])
        download = download_info.value
        pdf_name=f'{proposal_number}_{project_name}.pdf'
        download.save_as(os.path.join(os.getcwd(), pdf_name))
        upload_sheet_pdf(pdf_name)
        
    
    
def Automation_Process():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        login(page)
        page.goto("https://app.arbor-note.com/ng/#/sales/report",timeout=100000)
        upload_excel(page)
        upload_pdfs(page)

def main():
    Automation_Process()
    
    
    
if __name__ == '__main__':
    main()