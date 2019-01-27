from __future__ import print_function
import pickle
import os.path
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Change this to your own folder if you want this to work:    
# Make sure that it's pointing to the correct account folder. 
SV_FOLDER = 'C:\Fast Games\World of Retailcraft\World of Warcraft\_retail_\WTF\Account\SILVE\SavedVariables'



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range for the spreadsheet. The credentials you're using must have access to the sheet!
SHEET_ID = '1TaZOz7bSNzUAXsfsEZnNiw_OVl6xZoSLICJQr-FO30o'
RANGE_NAME = 'DataImport'



##################################################################################################################

# Handles authentication and returns the authenticated service.
def getAuth():
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service

# Parse the TSM4 SavedVariables to grab cached guild bank info.
def parseData(filepath):
    flag = False
    data = open(filepath, errors="ignore")

    itemId = []
    quantity = []

    for line in data:
        # when we get to the part of the file with the guildvault
        if '["f@Horde - Arthas@internalData@guildVaults"]' in line:
            flag = True
        
        if flag == True:
            print(line)
            # break out at end of structure
            if "}" in line:
                break
            
            if "i:" in line: 
                match = re.search(r'([0-9]+)"] = ([0-9]+)', line)
                itemId.append(match[1])
                quantity.append(match[2])
    
    body = {'values': [itemId, quantity]}

    return body





def main():
    # prepare data
    #itemId = [1, 3, 7, 9, 15]
    #quantity = [20, 15, 10, 5, 1]

    #body = {'values': [itemId, quantity]}
    body = parseData(SV_FOLDER + '\TradeSkillMaster.lua')
    service = getAuth()
    # Call the Sheets API
    sheet = service.spreadsheets()

    result = sheet.values().update(spreadsheetId=SHEET_ID, 
                valueInputOption='USER_ENTERED', range=RANGE_NAME, body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))
    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
    #                            range=SAMPLE_RANGE_NAME).execute()
    #values = result.get('values', [])

  
    #for row in values:
        # Print columns A and E, which correspond to indices 0 and 4.
    #    print('%s, %s' % (row[0], row[4]))

if __name__ == '__main__':
    main()