from credentials import CLIENT_ID, CLIENT_SECRET
from oauth2client.client import OAuth2WebServerFlow
from flask import Flask, request
import requests
from threading import Thread
from urlparse import urlparse

REDIRECT_URI = 'http://localhost:5000' #Flask uses port 5000 by default
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
USER_AGENT = 'ams-gradebook/1.0'

flaskapp = Flask(__name__)
@flaskapp.route("/")
def getauthcode():
    print request.args.get('code')
    return request.args.get('code')

flow = OAuth2WebServerFlow(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        redirect_uri=REDIRECT_URI,
        scope=SCOPE,
        user_agent=USER_AGENT)

if __name__=='__main__':
    flask_thread = Thread(target=flaskapp.run)
    flask_thread.daemon = True
    flask_thread.start()
    auth_uri = flow.step1_get_authorize_url()
    auth_response = requests.get(auth_uri)
    if 'accounts.google.com' == urlparse(auth_response.history[-1].url).netloc:
        print "Please authorize the app here:\n{}".format(auth_response.history[-1].url)
    print auth_response.url
