import os
import pickle
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# Load student data
data = pd.read_excel('students.xlsx')
print(f"Data: {data}")

# Setup Google Drive
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secrets.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('drive', 'v3', credentials=creds)

# Get the template file
template_id = '1tJtenb_7f5kvhhNA9qXFDH_l5xIGRSXbXzZ9Lqb464M'  # Update with your 'test master' file ID

# Specify the ID of the folder where files should be stored
folder_id = '15Q75aSfv0zp-vKQ4sUn7yArrG4l8Lb2Y'  # Update with your 'test' folder ID

# For each student
for index, row in data.iterrows():
    print(f"Processing: {row}")

    # Copy the template and rename it
    copied_file = service.files().copy(
        fileId=template_id,
        body={'name': f"test_{row['name']}", 'parents': [folder_id]}  # Rename the file here
    ).execute()

    print(f"File copied for {row['name']}")

    # Share the file with the student
    permission = {
        'type': 'user',
        'role': 'reader',
        'emailAddress': row['email']
    }
    service.permissions().create(
        fileId=copied_file['id'],
        body=permission,
        fields='id',
    ).execute()

    print(f"File shared with {row['email']}")
