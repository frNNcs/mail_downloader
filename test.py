from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import base64
import datetime
from dateutil import parser
from console_progressbar import ProgressBar
import json
import email
from bs4 import BeautifulSoup


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
# MAIL = 'jormarcdis@gmail.com'
MAIL = 'oficinavirtual@edelap.com.ar'
QUERY = [
    'from:{}'.format(MAIL),
    # 'has:attachment',
]

def main():
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    results = service.users().messages().list(userId='me', q=QUERY).execute()
    messages = results.get('messages', [])

    i = 0
    pb = ProgressBar(total=len(messages), decimals=3, length=50, fill='#', zfill='-',prefix='0', suffix=len(messages))

    if len(messages) > 0:
        mail_folder = str(os.getcwd()) + '/media/' + MAIL + '/'

        try:
            os.mkdir(mail_folder)
        except:
            print('Carpeta existente')

        mail_raw_folder = mail_folder + 'mails/'
        try:
            os.mkdir(mail_raw_folder)

        except:
            print('Carpeta de mails existente')

        try:
            mails_downloaded_file = open(mail_folder + "mails_downloaded.json", "r")
            mails_downloaded = mails_downloaded_file.read()
            mails_downloaded = json.loads(mails_downloaded)
            mails_downloaded = mails_downloaded['ids']
            if mails_downloaded == None:
                mails_downloaded = []
            mails_downloaded_file.close()
        except:
            mails_downloaded = []

        for message in messages:
            i += 1
            pb.print_progress_bar(i)

            # Descargar adjuntos.
            if not message['id'] in mails_downloaded:
                msg = service.users().messages().get(userId='me', id=message['id']).execute()

                header = msg.get('payload').get('headers')
                for item in header:
                    if item['name'] == 'Date':
                        date = parser.parse(item['value'])
                        date = str(date.date())

                parts = msg.get('payload').get('parts')
                all_parts = []
                if parts:
                    for p in parts:
                        if p.get('parts'):
                            all_parts.extend(p.get('parts'))
                        else:
                            all_parts.append(p)

                    att_parts = [p for p in all_parts if 'attachmentId' in p['body']]
                    filenames = [p['filename'] for p in att_parts]

                    folder =  mail_folder + date + ' - ' + msg['id']
                    try:
                        os.mkdir(folder)
                    except:
                        mails_downloaded.append(message['id'])
                    else:
                        for part,filename in zip(att_parts,filenames):
                            data = part['body'].get('data')
                            attachmentId = part['body'].get('attachmentId')
                            if not data:
                                att = service.users().messages().attachments().get(
                                        userId='me',
                                        id=attachmentId,
                                        messageId=message['id']).execute()
                                data = att['data']

                            str_file = base64.urlsafe_b64decode(data.encode('UTF-8'))
                            f = open(folder + '/' + filename, "wb")
                            f.write(str_file)
                            f.close()

                # Descargar emails.
                message = service.users().messages().get(userId='me', id=message['id'], format='raw').execute()
                msg_str = base64.urlsafe_b64decode(message['raw'].encode("utf-8"))
                mime_msg = email.message_from_string(msg_str.decode("utf-8", "replace"))
                html_mail = base64.urlsafe_b64decode(mime_msg._payload)

                soup = BeautifulSoup(html_mail, 'html.parser')
                tds = soup.find_all('td')
                f = open(mail_raw_folder + date + ' - ' + msg['id'] + '.html', "wb")
                f.write(html_mail)
                f.close()
                mails_downloaded.append(message['id'])

            mails_downloaded_file = open(mail_folder + '/' + "mails_downloaded.json", "w")
            mails_downloaded_file.write(json.dumps({'ids': mails_downloaded}))
            mails_downloaded_file.close()

    else:
        print('No se han encontrado mails')

if __name__ == '__main__':
    main()