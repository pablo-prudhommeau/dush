from __future__ import (unicode_literals, absolute_import, print_function, division)

import base64
import configparser
import datetime
import io
import logging
import os
import re
from time import sleep

import coloredlogs
from borb.pdf.pdf import PDF
from borb.toolkit.text.simple_text_extraction import SimpleTextExtraction
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


def authenticate():
    scopes = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/gmail.modify'
    ]
    credentials = None
    if os.path.exists('secrets/token.json'):
        credentials = Credentials.from_authorized_user_file('secrets/token.json', scopes)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('config/credentials.json', scopes)
            credentials = flow.run_local_server()
        with open('secrets/token.json', 'w') as token:
            token.write(credentials.to_json())
    credentials = Credentials.from_authorized_user_file('secrets/token.json', scopes)
    return credentials


def upload_file_to_google_drive(attachment, filename):
    media = MediaIoBaseUpload(io.BytesIO(attachment), mimetype='application/pdf')
    credentials = authenticate()
    drive = build('drive', 'v3', credentials=credentials)
    config = get_config()

    logging.info('Uploading file to Google Drive - filename [' + filename + '], googleDriveParentFolderId [' +
                 config['google.drive']['ParentFolderId'] + ']')
    drive.files().create(
        body={'parents': [config['google.drive']['ParentFolderId']], 'name': filename},
        media_body=media,
        fields='id'
    ).execute()
    logging.info('File successfully uploaded to Google Drive - filename [' + filename + ']')


def list_invoice_emails():
    credentials = authenticate()
    gmail = build('gmail', 'v1', credentials=credentials)
    response = gmail.users().messages().list(q='facture has:attachment in:inbox', userId='me').execute()

    logging.debug('Scanning eligible emails containing Leroy Merlin invoice...')
    if 'messages' in response:
        for msg_id in response['messages']:
            message = gmail.users().messages().get(id=msg_id['id'], userId='me').execute()
            invoice = gmail.users().messages().attachments().get(userId='me', messageId=message['id'],
                                                                 id=message['payload']['parts'][1]['body'][
                                                                     'attachmentId']).execute()
            logging.info('Get message details from Gmail - messageId [' + message['id'] + ']')
            attachment = bytearray(base64.urlsafe_b64decode(invoice['data'].encode('UTF-8')))
            filename = compute_invoice_filename(attachment)
            upload_file_to_google_drive(attachment, filename)
            config = get_config()
            gmail.users().messages().modify(id=msg_id['id'], userId='me',
                                            body={'addLabelIds': [config['google.gmail']['ArchiveLabelId']],
                                                  'removeLabelIds': ['INBOX']}).execute()
    else:
        logging.debug('No eligible email found')


def compute_invoice_filename(attachment_byte_array):
    simple_text_extraction = SimpleTextExtraction()
    PDF.loads(io.BytesIO(attachment_byte_array), [simple_text_extraction])
    index = 0
    line_groups = []
    last_item = {}
    invoice_total_ttc = None
    invoice_date = None
    invoice_id = None
    while index <= simple_text_extraction._current_page:
        lines = simple_text_extraction.get_text_for_page(index).split('\n')
        for line in lines:
            match_item_line = re.match('^([0-9]+) ([0-9]{8}) (.*?)( Tx TVA .*?)?$', line)
            if match_item_line:
                item_id = match_item_line.group(2)
                item_designation = match_item_line.group(3)
                last_item = {'id': item_id, 'designation': item_designation, 'total_price': -1}
                line_groups.append(last_item)
            match_price_line = re.match('^(-?[0-9]+\.[0-9]+) € (-?[0-9]+\.[0-9]+) € ([0-9]+) (-?[0-9]+\.[0-9]+) €',
                                        line)
            if match_price_line:
                last_item['total_price'] = float(match_price_line.group(4))
            match_invoice_total_ttc = re.match('^Total TTC (-?[0-9]+\.[0-9]+) €$', line)
            if match_invoice_total_ttc:
                invoice_total_ttc = float(match_invoice_total_ttc.group(1))
            match_invoice_date = re.match('^Exemplaire client / Date d\'émission : (.*?)$', line)
            if match_invoice_date:
                invoice_date = datetime.datetime.strptime(match_invoice_date.group(1), '%d/%m/%Y')
            match_invoice_id = re.match('(AVOIR|FACTURE) N° ([0-9]+)', line)
            if match_invoice_id:
                invoice_id = match_invoice_id.group(2)
        index += 1
    line_groups = sorted(line_groups, key=lambda x: x['total_price'], reverse=True)
    pdf_name = 'Leroy Merlin - ' + invoice_date.strftime('%Y_%m_%d') + ' - ' + invoice_id + ' - ' \
               + str(invoice_total_ttc) + '€ ('
    index = 0
    item_details_number = len(line_groups) if len(line_groups) < 3 else 3
    while index < item_details_number:
        pdf_name += line_groups[index]['designation'] + ' - ' + str(line_groups[index]['total_price']) + '€'
        if index < item_details_number - 1:
            pdf_name += ' | '
        index += 1
    if len(line_groups) > 3:
        pdf_name += ', ...'
    pdf_name += ').pdf'
    return pdf_name


def get_config():
    config = configparser.ConfigParser()
    config.read('config/config.ini')
    return config


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)
    coloredlogs.install()

    logging.info('Starting dush...')

    config = get_config()
    logging.info('Starting scanner scheduler every ' + config['default']['SchedulerIntervalInSeconds'] + ' seconds...')

    while True:
        list_invoice_emails()
        sleep(int(config['default']['SchedulerIntervalInSeconds']))
