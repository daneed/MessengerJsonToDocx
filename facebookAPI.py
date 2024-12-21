from pyfacebook import GraphAPI
from pynani import Messenger

import os 
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

APP_ID = "958399926186795"
APP_SECRET = "6ddd78f577caf07922ab4f51856e1a7d"

def get_app_access_token():
    client = GraphAPI(app_id=APP_ID, app_secret=APP_SECRET,  application_only_auth=True)
    return client.access_token

__name__ == '__main__' and print(get_app_access_token())


token = "6ddd78f577caf07922ab4f51856e1a7d"
api = GraphAPI(app_id=APP_ID, app_secret=APP_SECRET, oauth_flow=True)
api.get_authorization_url()
api.exchange_user_access_token(response="url redirected")
pass
#api.get_authorization_url()
p = api.get_object( 'me/conversations')

pass