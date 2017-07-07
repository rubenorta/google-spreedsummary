from __future__ import print_function
import httplib2
import os, sys

from datetime import datetime
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from pprint import pprint

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SUMMARY_PERIOD = datetime.strptime('01/06/2017','%d/%m/%Y')

SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_id.json'
APPLICATION_NAME = 'TR-Reporting'

TR_FILES = { 'Alex' : '1kpZrtZjTL0fHKbfhlRI1xvPyGoyNQ2lYhnJvwQ269ic',
    'Alexandra': '1GgB2ULrJgmQkij03wggvvL0LLOWxXfEKlADLtnIzpn0',
    'Carlos': '1rsevQKmiSXR9_YHTKc7FaW1lClmjHW9fDPJn_TKX584',
    'David': '1ouohX2MAA7L8gXCdX-PUB-TM1a-Evo3EhBh_2cr1dRA',
    'Eloy': '1zTcuP0RWqCnHNqSAGj1t0e3k8u2p69YmhLtKqdBjFGM',
    'Fernando': '1uHlDAiBpaJhfE-MU7Xh4eWZzn6DAlqmjpF8aWMf3BAw',
    'Ivan': '1KEnimcQriyLoTE-Ql1GKzMkZtpCA-FNHqJqb5jUiS6g',
    'Jose Angel': '17I0-uUPTbBIz1CpKjgePA1N22i58pMSgM2PtIkyxiJo',
    'Rocio': '1G-rop97G1Kr4oo2tQu6ub5H-xmHXNbM9DH397MuFWIs',
    'Sergio':'1m3SnXO149P6y_VEuZCWLCFVbrQbw_2pHxwSHkHy-r6w'}

ROW_LEN = 5

def get_credentials():

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials-good')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def is_processable(sheet,summary):
    return (sheet.month == summary.month)

def summary(data,row):

    if len(row) != ROW_LEN:
        return data

    projects = data['projects']
    countries = data['countries']
    key_proj = row[1]
    key_coun = row[3] if row[4] else 'TODOS'
    hours = row[4] if row[4] else '0'
    projects[key_proj] = round(projects.get(key_proj, 0) + float(hours), 2)
    countries[key_coun] = round(countries.get(key_coun, 0) + float(hours), 2)
    
    return { 'projects' :  projects, 'countries' : countries }

def is_current_row(row):
    if len(row) != ROW_LEN:
        return False
    else:
        row_period = datetime.strptime(row[0],'%d/%m/%Y') 
        return (row_period.month == SUMMARY_PERIOD.month)

def process_sheet(service, spreadsheetId):

    data = {'projects' : {}, 'countries' : {} }

    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')

    for sheet in sheets:
        sheet_props = sheet.get('properties')
        title = sheet_props.get('title')
        rowCount = sheet_props.get('gridProperties').get('rowCount')

        rangeName = title + "!D6:D7" + str(rowCount)
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()

        if (result.get('values')):
            
            start_period = result.get('values', [])[0][0]
            start_period = datetime.strptime(start_period,'%d/%m/%Y')
            end_period = result.get('values', [])[1][0]
            end_period = datetime.strptime(end_period,'%d/%m/%Y')

            if (is_processable(start_period,SUMMARY_PERIOD) or is_processable(end_period,SUMMARY_PERIOD)):
                print("...periods " + str(start_period) + " a " + str(end_period) )

                rangeName = title + "!B14:F" + str(rowCount)
                result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
    
                values = result.get('values', [])
                if not values:
                    print("No data found.")
                else:
                    for row in values:
                        if row:
                            if is_current_row(row):
                                data = summary(data,row)        
    print("Resultados:")
    pprint(data)




def main():

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())

    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    for name, spreadsheetId in TR_FILES.iteritems():
        print(name + " ", end = '')
        process_sheet(service, spreadsheetId);

if __name__ == '__main__':
    main()
