import base64
import pickle
import os.path
import shelve
import time
from email.mime.text import MIMEText
import sys

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import dropbox

# If modifying these scopes, delete the file token.pickle.
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.compose']
GOOGLE_TOKEN_FILE = 'token.pickle'
DROPBOX_TOKEN_FILE = 'dropbox.pickle'
SETTINGS_FILE = 'settings.shelve'
THREAD_ID_SET = 'read_threads'


def authenticate_google():
    credentials = None
    if os.path.exists(GOOGLE_TOKEN_FILE):
        with open(GOOGLE_TOKEN_FILE, 'rb') as token:
            credentials = pickle.load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)
    return credentials


def authenticate_dropbox():
    credentials = None
    if os.path.exists(DROPBOX_TOKEN_FILE):
        with open(DROPBOX_TOKEN_FILE, 'rb') as token:
            credentials = pickle.load(token)
    if not credentials:
        print("Input dropbox token:")
        credentials = sys.stdin.readline().rstrip()
        with open(DROPBOX_TOKEN_FILE, 'wb') as token:
            pickle.dump(credentials, token)
    return credentials


def main():
    global service
    global settings
    global read_thread_set
    global dropbox_client

    google_credentials = authenticate_google()
    dropbox_credentials = authenticate_dropbox()
    settings = shelve.open(SETTINGS_FILE, writeback=True)
    if THREAD_ID_SET in settings:
        read_thread_set = settings[THREAD_ID_SET]
    else:
        read_thread_set = set()
        settings[THREAD_ID_SET] = read_thread_set
        settings.sync()

    service = build('gmail', 'v1', credentials=google_credentials)
    dropbox_client = dropbox.Dropbox(dropbox_credentials)

    # Call the Gmail API
    labels_results = service.users().labels().list(userId='me').execute()
    if 'labels' not in labels_results:
        print('No labels found.')
        return
    labels = labels_results['labels']

    test_label_id = None
    for label in labels:
        if label['name'] == 'Air Flow Data':
            test_label_id = label['id']
    if not test_label_id:
        return

    while True:
        print("getting new emails...")
        mail_infos_results = service.users().messages().list(userId='me', labelIds=[test_label_id]).execute()
        if 'messages' not in mail_infos_results:
            print('No Messages found.')
        else:
            mail_infos = mail_infos_results['messages']

            for mail_id in mail_infos:
                process_mail_id(mail_id)
        print('Finished processing')
        time.sleep(60)


def process_mime_part(mail_part, mail_id):
    global upload_metadata
    if mail_part["filename"]:
        if mail_part["filename"].endswith(".xlsx"):
            attachment_content_result = \
                service.users().messages().attachments().get(
                    userId='me',
                    messageId=mail_id['id'],
                    id=mail_part['body']['attachmentId']
                ).execute()
            attachment_string = attachment_content_result['data']
            attachment_data = base64.urlsafe_b64decode(attachment_string.encode('UTF-8'))
            metadata = dropbox_client.files_alpha_upload(attachment_data, "/" + mail_part['filename'],
                                                         autorename=True)
            upload_metadata += [metadata]
    sub_mail_parts = mail_part.get('parts', [])
    if sub_mail_parts:
        for sub_mail_part in sub_mail_parts:
            process_mime_part(sub_mail_part, mail_id)


def process_mail_id(mail_id):
    global service
    global settings
    global read_thread_set
    global dropbox_client

    if mail_id['threadId'] in read_thread_set:
        return
    print("new emails found")
    global upload_metadata
    upload_metadata = []
    mail_results = service.users().messages().get(userId='me', id=mail_id['id']).execute()
    mail_part = mail_results.get('payload', {})
    process_mime_part(mail_part, mail_id)
    if upload_metadata:
        response_text = "%d file(s) uploaded to dropbox! \n" % len(upload_metadata)
        response_text += str(upload_metadata)
    else:
        response_text = "No file detected"
    reply_body = create_reply(mail_results, response_text)

    try:
        message = (service.users().messages().send(userId='me', body=reply_body)
                   .execute())
        print('Message Id: %s' % message['id'])

        print("email replied")
        read_thread_set.add(mail_id['threadId'])
        settings[THREAD_ID_SET] = read_thread_set
        settings.sync()
        return message
    except HttpError as error:
        print('An error occurred: %s' % error)


def get_mail_header(section_name, mail):
    return \
        next((mail_header['value'] for mail_header in mail['headers'] if mail_header['name'] == section_name), None)


def create_reply(original_mail_result, message_text):
    original_mail = original_mail_result['payload']
    message = MIMEText(message_text)
    original_from = get_mail_header('From', original_mail)
    message['To'] = original_from
    original_cc = get_mail_header('Cc', original_mail)
    if original_cc:
        message['Cc'] = original_cc
    message['Subject'] = 'Re: ' + get_mail_header('Subject', original_mail)
    original_in_reply_to = get_mail_header('In-Reply-To', original_mail)
    original_references = get_mail_header('References', original_mail)
    original_msg_id = get_mail_header('Message-ID', original_mail)
    message['In-Reply-To'] = original_msg_id
    new_references = ""
    if original_references:
        new_references += original_references + " "
    if original_in_reply_to:
        new_references += original_in_reply_to + " "
    message['References'] = new_references
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode(), "threadId": original_mail_result['threadId']}


if __name__ == '__main__':
    main()
