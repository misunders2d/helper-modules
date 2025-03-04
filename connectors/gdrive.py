import pickle
from io import BytesIO, StringIO
import os.path
import pandas as pd

import gdown, gspread
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload, MediaIoBaseUpload

# Update the scopes to include access to Shared Drives
SCOPES = ['https://www.googleapis.com/auth/drive']

def gdownload(file_id):
    buf = BytesIO()
    _ = gdown.download(id = file_id, output = buf)
    buf.seek(0)
    return buf

def connect(scope = 'files'):
    creds = None
    if os.path.exists('.secrets/gdrive.pickle'):
        with open('.secrets/gdrive.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'connectors/gdrive.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('.secrets/gdrive.pickle', 'wb') as token:
            pickle.dump(creds, token)
    if scope == 'files':
        service = build('drive', 'v3', credentials=creds)
    elif scope == 'gspread':
        service = gspread.authorize(creds)
    return service

def delete_file(file_id, service = connect()):
    try:
        service.files().delete(
            fileId=file_id,
            supportsAllDrives=True
            ).execute()
        print(f"File with ID: {file_id} has been deleted successfully.")
    except HttpError as error:
        print(f"An error occurred while deleting the file: {error}")

def download_gspread(service = connect(scope = 'gspread'),
                     spreadsheet_id = None,
                     sheet_id = None,
                     header = 1
                     ):
    if not spreadsheet_id:
        raise BaseException('Spreadsheet ID not indicated')
    sheet = 0 if not sheet_id else sheet_id
    book = service.open_by_key(spreadsheet_id)
    sheet = book.get_worksheet_by_id(sheet)
    data = pd.DataFrame(sheet.get_all_records(head=header))   
    return data
    
def create_folder(folder_name, parent_folder):
    # Create a folder
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder]
    }
    folder = connect().files().create(
        body=file_metadata,
        fields='id',
        supportsAllDrives=True
        ).execute()
    print('Folder ID: {}'.format(folder.get('id')))
    return folder.get('id')

def list_shared_drives(service = connect()):
    shared_drives = []
    page_token = None
    while True:
        try:
            results = service.drives().list(
                pageSize=50,
                fields="nextPageToken, drives(id, name)",
                pageToken=page_token
            ).execute()
            shared_drives.extend(results.get('drives', []))
            page_token = results.get('nextPageToken', None)
            if page_token is None:
                break
        except HttpError as error:
            print(f"An error occurred while listing shared drives: {error}")
            break
    return shared_drives

def list_folders(service, drive_id):
    folders = []
    page_token = None
    while True:
        try:
            results = service.files().list(
                q="mimeType='application/vnd.google-apps.folder'",
                spaces='drive',
                fields="nextPageToken, files(id, name)",
                pageSize=50,
                driveId=drive_id,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                corpora='drive',
                pageToken=page_token
            ).execute()
            folders.extend(results.get('files', []))
            page_token = results.get('nextPageToken', None)
            if page_token is None:
                break
        except HttpError as error:
            print(f"An error occurred while listing folders: {error}")
            break
    return folders

def list_files_in_folder(folder_id, drive_id, service = connect()):
    files = {}
    page_token = None
    while True:
        try:
            results = service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                spaces='drive',
                fields="nextPageToken, files(id, name, mimeType)",
                pageSize=50,
                driveId=drive_id,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                corpora='drive',
                pageToken=page_token
            ).execute()
            items = results.get('files', [])
            if not items:
                print('No files found.')
            else:
                for item in items:
                    files[item['name']] = {'id':item['id'],'type':item['mimeType']}
            page_token = results.get('nextPageToken', None)
            if page_token is None:
                break
        except HttpError as error:
            print(f"An error occurred while listing files: {error}")
            break
    return files

def find_file_id(folder_id, drive_id, filename):
    result = list_files_in_folder(folder_id, drive_id)
    return result.get(filename,{}).get('id')
    

def upload_file(file_path, parent_folder_id, service = connect()):
    try:
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [parent_folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, 
                                      fields='id', supportsAllDrives=True
                                      ).execute()
        return file
    except HttpError as error:
        print(f"An error occurred while uploading the file: {error}")
        return None

def create_file(file_bytes,file_name,parent_folder,
                mimetype = None, service = connect()
                ):
    file_metadata = {
        'name':file_name,
        'parents':[parent_folder]
        }
    media = MediaIoBaseUpload(file_bytes, mimetype = mimetype, resumable=True)

    updated_file = service.files().create(
        body=file_metadata,
        media_body=media,
        # driveId=drive_id,
        supportsAllDrives=True
    ).execute()

    print('File updated with ID: {}'.format(updated_file.get('id')))
    return

def replace_file(file_id, new_file_bytes, mimetype, service = connect()):
    media = MediaIoBaseUpload(new_file_bytes, mimetype = mimetype, resumable=True)

    updated_file = service.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True
    ).execute()

    print('File updated with ID: {}'.format(updated_file.get('id')))

def download_file(file_id, service = connect()):
    # Request the file
    request = service.files().get_media(fileId=file_id)
    
    # Create a BytesIO object to receive the file content
    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")

    # Write the content to a file
    buf.seek(0)
    return buf

def upload_folder(service, folder_path, parent_folder_id, drive_id):
    try:
        folder_metadata = {
            'name': os.path.basename(folder_path),
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = service.files().create(body=folder_metadata, fields='id', 
                                        supportsAllDrives=True, 
                                        driveId=drive_id).execute()
        
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                upload_file(service, item_path, folder['id'], drive_id)
            elif os.path.isdir(item_path):
                upload_folder(service, item_path, folder['id'], drive_id)
        
        return folder
    except HttpError as error:
        print(f"An error occurred while uploading the folder: {error}")
        return None