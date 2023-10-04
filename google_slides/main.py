import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import socket
import pandas as pd
socket.setdefaulttimeout(60*10)

ROOT_FOLDER_NAME = 'test_folder'

# Delete the token.json when the scopes is changed!
# When searching it does include the bin for some reason

def authenticate_google(token="token.json", scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/script.projects']):
    if os.path.exists(token):
        creds = Credentials.from_authorized_user_file(token, scopes)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret.json', scopes)
        creds = flow.run_local_server(port=0)
        with open(token, 'w') as token:
            token.write(creds.to_json())
    drive_service = build('drive', 'v3', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    script_service = build('script', 'v1', credentials=creds)
    return drive_service, slides_service, script_service

def generate_pdfs(drive_service, script_service, slides='ALL'):
    parent_folder_id = drive_service.files().list(q=f"name='{ROOT_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    slides_folder_id = drive_service.files().list(q=f"name='slides' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    slides = drive_service.files().list(q=f"'{slides_folder_id}' in parents and mimeType='application/vnd.google-apps.presentation'").execute()['files']
    slides = [pres['id'] for pres in slides] if slides == 'ALL' else [pres['id'] for pres in slides if pres in slides]

    # find the apps script and execute it
    query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.script'"
    response = drive_service.files().list(q=query).execute()
    script_id = response['files'][0]['id']
    response = script_service.scripts().run(scriptId=script_id, body={
        'function': 'exportPDFsFromFolder',
        'devMode': True,
        'parameters':[
            parent_folder_id,
            slides_folder_id,
            slides
        ]
    }).execute()

def get_pdfs(drive_service, slides='ALL'):
    parent_folder_id = drive_service.files().list(q=f"name='{ROOT_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    pdf_folder_id = drive_service.files().list(q=f"name='PDFs' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    files = drive_service.files().list(q=f"'{pdf_folder_id}' in parents and mimeType='application/pdf'").execute()
    for file in files.get('files', []):
        file_id = file['id']
        file_name = file['name']
        if not file_name.lower().endswith('.pdf'):
            file_name += '.pdf'
        file_name = f"./pdfs/{file_name}"
        if not os.path.exists(file_name):
            request = drive_service.files().get_media(fileId=file_id)
            fh = open(file_name, 'wb')
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Downloaded {int(status.progress() * 100)}% of {file_name}")
            fh.close()

def get_speaker_notes(drive_service, slides_service, slides='ALL'): #This function might also be used to properly make the svg's (Because slides may be skipped, but are included in the pdfs)
    parent_folder_id = drive_service.files().list(q=f"name='{ROOT_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    slides_folder_id = drive_service.files().list(q=f"name='slides' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'").execute()['files'][0]['id']
    slides = drive_service.files().list(q=f"'{slides_folder_id}' in parents and mimeType='application/vnd.google-apps.presentation'").execute()['files']
    slides = [pres['id'] for pres in slides] if slides == 'ALL' else [pres['id'] for pres in slides if pres in slides]
    

    for slide_id in slides:
        notes_dict = {"slide_num":[], "content": [], "style":[]}
        gslide = slides_service.presentations().get(presentationId=slide_id).execute()
        for index, slide in enumerate(gslide['slides']):
            slide_properties = slide['slideProperties']
            notes = slide_properties['notesPage']
            for element in notes['pageElements']:
                if 'text' in element['shape'].keys():
                    for telement in element['shape']['text']['textElements']:
                        if 'textRun' in telement.keys():
                            notes_dict['slide_num'].append(index)
                            notes_dict['content'].append(telement['textRun']['content'])
                            notes_dict['style'].append(telement['textRun']['style'])
        pd.DataFrame.from_dict(notes_dict).to_csv(f"./notes/{gslide['title']}.csv", index=False)

if __name__ == '__main__':
    if not os.path.exists("./pdfs"):
        os.mkdir("./pdfs")
    if not os.path.exists("./notes"):
        os.mkdir("./notes")
    drive_service, slides_service, script_service = authenticate_google()
    # generate_pdfs(drive_service, script_service)
    # get_pdfs(drive_service)
    get_speaker_notes(drive_service, slides_service)

    # Inject Peter's script here

