from __future__ import (unicode_literals, absolute_import, print_function, division)

import base64
import configparser
import datetime
import getopt
import io
import json
import logging
import os
import re
import sys
from time import sleep

import coloredlogs
import requests
from borb.pdf.pdf import PDF
from borb.toolkit.text.simple_text_extraction import SimpleTextExtraction
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from humanfriendly.terminal import usage


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
            attachment = base64.urlsafe_b64decode(invoice['data'].encode('UTF-8'))
            attachment_byte_array = bytearray(attachment)
            filename = compute_invoice_filename('', attachment_byte_array)
            upload_file_to_google_drive(attachment, filename)
            config = get_config()
            gmail.users().messages().modify(id=msg_id['id'], userId='me',
                                            body={'addLabelIds': [config['google.gmail']['ArchiveLabelId']],
                                                  'removeLabelIds': ['INBOX', 'UNREAD']}).execute()
    else:
        logging.debug('No eligible email found')


def compute_invoice_filename(original_file_name, attachment_byte_array):
    simple_text_extraction = SimpleTextExtraction()
    try:
        PDF.loads(io.BytesIO(attachment_byte_array), [simple_text_extraction])
    except AssertionError:
        logging.error('Unable to read PDF - originalFileName [' + original_file_name + ']')
        return

    page_index = 0
    line_groups = []
    last_item = {}
    invoice_total_ttc = None
    formatted_invoice_date = None
    invoice_id = None
    is_credit = False
    while page_index <= simple_text_extraction._current_page:
        lines = simple_text_extraction.get_text_for_page(page_index).split('\n')
        for line_index, line in enumerate(lines):
            match_item_line = re.match(
                '^([0-9]+) ([0-9]{8}) (.*?)( -?[0-9]+\.?([0-9]+)? -?[0-9]+\.?([0-9]+)? -?[0-9]+\.?([0-9]+)? (-?[0-9]+\.?([0-9]+)?))?( Tx TVA .*?)?$',
                line)
            if match_item_line:
                item_id = match_item_line.group(2)
                item_designation = match_item_line.group(3)
                total_price = -1
                if match_item_line.group(4) is not None:
                    total_price = abs(float(match_item_line.group(8)))
                last_item = {'id': item_id, 'designation': item_designation, 'total_price': total_price}
                line_groups.append(last_item)
            match_price_line = re.match(
                '^-?[0-9]+\.?([0-9]+)? €? -?[0-9]+\.?([0-9]+)? €? -?[0-9]+\.?([0-9]+)? (-?[0-9]+\.?([0-9]+)?) €?',
                line)
            if match_price_line:
                last_item['total_price'] = abs(float(match_price_line.group(4)))
            match_invoice_total_ttc = re.match('^Total TTC (-?[0-9]+\.[0-9]+) €$', line)
            if match_invoice_total_ttc:
                invoice_total_ttc = abs(float(match_invoice_total_ttc.group(1)))
            match_invoice_total_ttc_single = re.match('^Total TTC$', line)
            if match_invoice_total_ttc_single:
                invoice_total_ttc_line = lines[line_index - 10]
                if is_number(invoice_total_ttc_line):
                    invoice_total_ttc = abs(float(lines[line_index - 10]))
                else:
                    logging.warning(
                        'Invoice has invalid Total TTC amount - originalFileName [' + original_file_name + ']')
            match_invoice_date = re.match('^(.*?)([0-9]+/[0-9]+/[0-9]+)$', line)
            if match_invoice_date:
                invoice_date = datetime.datetime.strptime(match_invoice_date.group(2), '%d/%m/%Y')
                formatted_invoice_date = invoice_date.strftime('%Y_%m_%d')
            match_invoice_id = re.match('((AVOIR )|FACTURE )(.*?)([0-9]+)( DUPLICATA)?', line)
            if match_invoice_id:
                invoice_id = match_invoice_id.group(4)
                if match_invoice_id.group(2) is not None:
                    is_credit = True
        page_index += 1
    line_groups = sorted(line_groups, key=lambda x: x['total_price'], reverse=True)
    pdf_name = 'Leroy Merlin'
    if formatted_invoice_date is not None:
        pdf_name += ' - ' + formatted_invoice_date
    if invoice_id is not None:
        pdf_name += ' - ' + invoice_id
    if is_credit is True:
        pdf_name += ' - Avoir'
    if invoice_total_ttc is not None:
        pdf_name += ' - ' + str(invoice_total_ttc) + '€'
    if len(line_groups) > 0:
        pdf_name += ' ('
        page_index = 0
        item_details_number = len(line_groups) if len(line_groups) < 3 else 3
        while page_index < item_details_number:
            pdf_name += line_groups[page_index]['designation'] + ' - ' + str(
                line_groups[page_index]['total_price']) + '€'
            if page_index < item_details_number - 1:
                pdf_name += ' | '
            page_index += 1
        if len(line_groups) > 3:
            pdf_name += ', ...'
        pdf_name += ')'
    if formatted_invoice_date is None and invoice_id is None and is_credit is False and invoice_total_ttc is None:
        pdf_name += ' - @Non catégorisé'
    pdf_name += '.pdf'
    return pdf_name


def get_config():
    config = configparser.ConfigParser()
    config.read('config/config.ini')
    return config


def is_number(s):
    try:
        complex(s)
    except ValueError:
        return False
    return True


def launch_api_scanner():
    config = get_config()
    logging.info(
        'Starting API scanner scheduler every ' + config['default']['ApiSchedulerIntervalInSeconds'] + ' seconds...')

    s = requests.session()
    data = {
        'client_id': "mwuEYVQmdT0MhwPsvUxA",
        'email': "pablo.prudhommeau@gmail.com",
        'password': "G2kc6$i9g7bRC6OU#Ko#",
        'scope': "openid profile email offline_access address events phone full_write LMFR.MOVE.CUSTOMER PickupPlanning.appointments community shortenUrl:write LMFR.MOVE.SALES LMFR.MOVE.LOYALTY FRLM-WAC.rw-partner-creds shortenUrl:read move SCOPE_CUSTOMER customer:management webuser:read webuser:write"
    }
    url = "https://authentication.leroymerlin.fr/identity/v1/password/login"
    r = s.post(url, data=data)
    y = json.loads(r.content)
    r3 = s.get(
        'https://authentication.leroymerlin.fr/identity/v1/password/callback' +
        '?client_id=mwuEYVQmdT0MhwPsvUxA&response_type=code'
        + '&scope=openid%20profile%20email%20offline_access%20address%20events%20phone%20full_write%20LMFR.MOVE.CUSTOMER%20PickupPlanning.appointments%20community%20shortenUrl%3Awrite%20LMFR.MOVE.SALES%20LMFR.MOVE.LOYALTY%20FRLM-WAC.rw-partner-creds%20shortenUrl%3Aread%20move%20SCOPE_CUSTOMER%20customer%3Amanagement%20webuser%3Aread%20webuser%3Awrite'
        + '&state=' + 'c3RhdGU9OWZ1cTMzbWU1RURkVTNMakRwMlFvcDlBS0Q5YnRMbTJwU19ma2tYTmVvYz0mY2xpZW50X2lkPW13dUVZVlFtZFQwTWh3UHN2VXhBJnJlZGlyZWN0X3VyaT1odHRwczovL3d3dy5sZXJveW1lcmxpbi5mci9hdXRoZW50aWNhdGlvbi9vYXV0aDIvY2FsbGJhY2smc2NvcGU9b3BlbmlkIHByb2ZpbGUgZW1haWwgb2ZmbGluZV9hY2Nlc3MgYWRkcmVzcyBldmVudHMgcGhvbmUgZnVsbF93cml0ZSBMTUZSLk1PVkUuQ1VTVE9NRVIgUGlja3VwUGxhbm5pbmcuYXBwb2ludG1lbnRzIGNvbW11bml0eSBzaG9ydGVuVXJsOndyaXRlIExNRlIuTU9WRS5TQUxFUyBMTUZSLk1PVkUuTE9ZQUxUWSBGUkxNLVdBQy5ydy1wYXJ0bmVyLWNyZWRzIHNob3J0ZW5Vcmw6cmVhZCBtb3ZlIFNDT1BFX0NVU1RPTUVSIGN1c3RvbWVyOm1hbmFnZW1lbnQgd2VidXNlcjpyZWFkIHdlYnVzZXI6d3JpdGU='
        + '&redirect_uri=https%3A%2F%2Fwww.leroymerlin.fr%2Fauthentication%2Foauth2%2Fcallback'
        + '&tkn=' + y['tkn'],
        allow_redirects=False,
        headers={
            'accept-language': 'fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
        })
    logging.info(r3.next.url)
    r4 = s.get(r3.next.url,
               headers={
                   'accept-language': 'fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7',
                   'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
               })
    r2 = s.get(
        'https://www.leroymerlin.fr/customer-space-module/services/customer-after-sales-backend/me/receipts?page=0&size=10&yearOffset=0',
        headers={
            'accept-language': 'fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
        })
    logging.info(r2.content)


# while True:
#    list_invoice_emails()
#    sleep(int(config['default']['ApiSchedulerIntervalInSeconds']))


def launch_email_box_scanner():
    config = get_config()
    logging.info('Starting email box scanner scheduler every ' + config['default'][
        'GmailSchedulerIntervalInSeconds'] + ' seconds...')

    while True:
        list_invoice_emails()
        sleep(int(config['default']['GmailSchedulerIntervalInSeconds']))


def launch_manual_invoice_upload():
    logging.info('Starting manual invoice upload...')

    for filename in os.listdir('invoices'):
        with open(os.path.join('invoices', filename), "rb") as f:
            content = bytearray(f.read())
            filename = compute_invoice_filename(filename, content)
            if filename is not None:
                upload_file_to_google_drive(content, filename)


def usage():
    print('Usage: ' + sys.argv[0] + ' [--mode MANUAL|GMAIL|API]')


def main(argv):
    try:
        opts, args = getopt.getopt(argv, "hm:", ["help", "mode"])
    except getopt.GetoptError:
        usage()
        sys.exit(2)
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()
        elif opt in ("-m", "--mode"):
            if arg == 'API':
                launch_api_scanner()
                return
            elif arg == 'MANUAL':
                launch_manual_invoice_upload()
                return
            elif arg == 'GMAIL':
                launch_email_box_scanner()
                return
            else:
                usage()
                return
    usage()
    return


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)
    coloredlogs.install()

    main(sys.argv[1:])
