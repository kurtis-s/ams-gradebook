import argparse
import gspread
import httplib2

from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client import tools

SCOPE = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
CLIENT_SECRETS_JSON_PATH = 'client_secrets.json'
CREDENTIALS_FILE_PATH = 'credentials_file'

def _refresh_token(credentials):
    if credentials.access_token_expired:
        http = httplib2.Http()
        credentials.refresh(http)

    return credentials

def _authorize_application(storage):
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    flags = parser.parse_args()
    flow = flow_from_clientsecrets(CLIENT_SECRETS_JSON_PATH, scope=SCOPE)

    authorized_credentials = tools.run_flow(flow, storage, flags)

    return authorized_credentials

def get_credentials():
    """Returns an instantiated Credentials object.  Refreshes the Credential's
    access token if the token is expired.
    
    If no existing credentials file is found, attempt to authorize the application
    with Google.  After the application has been authorized, the new Credential is
    stored in CREDENTIALS_FILE_PATH.
    """
    storage = Storage(CREDENTIALS_FILE_PATH)
    credentials = storage.get()

    #storage.get() returns None if there is no credentials file
    if credentials is None:
        credentials = _authorize_application(storage)

    _refresh_token(credentials)

    return credentials

if __name__=='__main__':
    credentials = get_credentials()
    gc = gspread.authorize(credentials)
    #wks = gc.open("gradetestbook").Sheet1
    wks = gc.open_by_key('1SqrL7FigyTy9jhZ9pEDBYplRQaXnDf_Iz-8-MT1LN7o').sheet1
    print "value in B2"
    print wks.acell('B2').value
    print "done"
