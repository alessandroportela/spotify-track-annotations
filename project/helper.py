import string
import os

import time

import httplib2
import urllib.request, urllib.parse, urllib.error

from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

dir_path = os.path.abspath(os.path.dirname(__file__))

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRET_PATH = os.path.join(dir_path, "./client_secret.json")

OUTPUT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1NBTRBYQRFiCVt0ZfbfCVZgeKdjOWzVhmMAZALtxeqAQ/edit#gid=0"

def get_credentials(credential_path):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credential_path, SCOPES)
    return credentials

def build_service(credentials):
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    return service

def append_list_to_spreadsheet(row):
    sheet_url = OUTPUT_SHEET_URL
    sheet_name = "Sheet1"
    credentials = get_credentials(CLIENT_SECRET_PATH)
    service = build_service(credentials)
    sheet_path = urllib.parse.urlsplit(sheet_url).path
    sheet_id = sheet_path.split("/")[3]
    body = {
        "values": [row]
    }
    sheet_range = "Sheet1!A2:E"

    result = service.spreadsheets().values().append(spreadsheetId=sheet_id, range=sheet_range,
        valueInputOption="RAW", body=body).execute()

    time.sleep(1) #To make sure we don't make too many calls to the Sheets API

def update_cell_in_output_spreadsheet(row_number, value):
    sheet_url = OUTPUT_SHEET_URL
    sheet_name = "Sheet1"
    credentials = get_credentials(CLIENT_SECRET_PATH)
    service = build_service(credentials)
    sheet_path = urllib.parse.urlsplit(sheet_url).path
    sheet_id = sheet_path.split("/")[3]
    body = {
        "values": [value]
    }
    sheet_range = "Sheet1!F" + row_number
    result = service.spreadsheets().values().update(
    spreadsheetId=sheet_id, range=sheet_range,
    valueInputOption="USER_ENTERED", body=body).execute()

    time.sleep(1) #To make sure we don't make too many calls to the Sheets API

# Pulls from ProspectNow so we can parse Apt, #, Suite rows
def pull_all_output_rows():
    sheet_url = OUTPUT_SHEET_URL
    sheet_name = "Sheet1"
    sheet_path = urllib.parse.urlsplit(sheet_url).path
    sheet_id = sheet_path.split("/")[3]
    credentials = get_credentials(CLIENT_SECRET_PATH)
    service = build_service(credentials)
    sheet_range = "Sheet1!A2:E" #TODO

    rows = service.spreadsheets().values().get(
    spreadsheetId=sheet_id, range=sheet_range).execute()

    return rows

def clear_range(start_column_label="", start_row_label="", end_column_label="", end_row_label=""):
    sheet_url = OUTPUT_SHEET_URL
    sheet_name = "Sheet1"
    credentials = get_credentials(CLIENT_SECRET_PATH)
    service = build_service(credentials)
    sheet_path = urllib.parse.urlsplit(sheet_url).path
    sheet_id = sheet_path.split("/")[3]

    sheet_range = "Sheet1!" + start_column_label + start_row_label + ":" + end_column_label + end_row_label

    body = {}

    result = service.spreadsheets().values().clear(spreadsheetId=sheet_id, range=sheet_range, body=body).execute()