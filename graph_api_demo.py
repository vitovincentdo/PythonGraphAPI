import webbrowser
import msal
from msal import ConfidentialClientApplication, PublicClientApplication

APPLIACTION_ID = 'fc674c45-7f72-4cad-8f16-fe7098dbab14'
CLIENT_SECRET = '0758Q~PIX3hAdcIchI73vAJuxS_3WJhtcKetlcgE'
authority_url = 'https://login.microsoftonline.com/59daf140-4aee-4b77-80f4-4ea8bec86c2e'
base_url = 'https://graph.microsoft.com/v1.0/'

endpoint = base_url + 'me'
SCOPES = ['User.Read']

client = ConfidentialClientApplication(client_id=APPLIACTION_ID, client_credential=CLIENT_SECRET,authority=authority_url)
authorization_url = client.get_authorization_request_url(SCOPES)
print(authority_url)
webbrowser.open(authorization_url)