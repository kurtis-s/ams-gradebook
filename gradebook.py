import argparse
import gspread
import httplib2

from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client import tools

class Authorizor(object):
    SCOPE = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
    CLIENT_SECRETS_JSON_PATH = 'client_secrets.json'
    CREDENTIALS_FILE_PATH = 'credentials_file'

    def __init__(self):
        storage = Storage(self.CREDENTIALS_FILE_PATH)
        self.storage = storage
        self.credentials = storage.get()

    def _refresh_token(self):
        if self.credentials.access_token_expired:
            http = httplib2.Http()
            self.credentials.refresh(http)
            self.storage.put(self.credentials)

    def _authorize_application(self):
        parser = argparse.ArgumentParser(parents=[tools.argparser])
        flags = parser.parse_args()
        flow = flow_from_clientsecrets(self.CLIENT_SECRETS_JSON_PATH, scope=SCOPE)

        authorized_credentials = tools.run_flow(flow, self.storage, flags)

        return authorized_credentials

    def get_credentials(self):
        """Returns an instantiated Credentials object.  Refreshes the Credential's
        access token if the token is expired.
        
        If no existing credentials file is found, attempt to authorize the application
        with Google.  After the application has been authorized, the new Credential is
        stored in CREDENTIALS_FILE_PATH.
        """
        #storage.get() returns None if there is no credentials file
        if self.credentials is None:
            credentials = self._authorize_application(storage)

        self._refresh_token()

        return self.credentials

if __name__=='__main__':
    authorizor = Authorizor()
    credentials = authorizor.get_credentials()
    gc = gspread.authorize(credentials)
    #wks = gc.open("gradetestbook").Sheet1
    wks = gc.open_by_key('1SqrL7FigyTy9jhZ9pEDBYplRQaXnDf_Iz-8-MT1LN7o').sheet1
    print "value in B2"
    print wks.acell('B2').value
    print "done"
