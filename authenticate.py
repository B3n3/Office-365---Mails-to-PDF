from O365 import Account
import credentials

credentials = (credentials.clientID, credentials.clientSecret)

account = Account(credentials)
if account.authenticate(scopes=['basic', 'message_all']):
    print('Authenticated!')

