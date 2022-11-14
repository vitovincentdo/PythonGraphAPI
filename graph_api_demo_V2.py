import msal
import jwt
import json
import sys
import requests
from datetime import datetime
from msal_extensions import *

graphURI = 'https://graph.microsoft.com'
tenantID = '59daf140-4aee-4b77-80f4-4ea8bec86c2e'
authority = 'https://login.microsoftonline.com/' + tenantID
# authority = 'https://login.microsoftonline.com/common/oauth2/authorize?resource=https://bcaoffice365.sharepoint.com'
clientID = 'b80e63c4-5843-44da-badc-83a20c33dff1'
clientSecret = 'ZoV8Q~fBGqAfI9QY2ALHmMLf7zOfOM7D3lTYZa4t'
scope = ["Sites.Read.All", "Sites.ReadWrite.All", "Sites.Manage.All"]
username = 'vito_vincentdo@bca.co.id'
result = None
tokenExpiry = None

def msal_persistence(location, fallback_to_plaintext=False):
    """Build a suitable persistence instance based your current OS"""
    if sys.platform.startswith('win'):
        return FilePersistenceWithDataProtection(location)
    if sys.platform.startswith('darwin'):
        return KeychainPersistence(location, "my_service_name", "my_account_name")
    return FilePersistence(location)

def msal_cache_accounts(clientID, authority):
    # Accounts
    persistence = msal_persistence("token_cache.bin")
    print("Is this MSAL persistence cache encrypted?", persistence.is_encrypted)
    cache = PersistedTokenCache(persistence)
    
    app = msal.PublicClientApplication(
        client_id=clientID, authority=authority, token_cache=cache)
    accounts = app.get_accounts()
    print(accounts)
    return accounts

def msal_delegated_refresh(clientID, scope, authority, account):
    persistence = msal_persistence("token_cache.bin")
    cache = PersistedTokenCache(persistence)
    
    app = msal.PublicClientApplication(
        client_id=clientID, authority=authority, token_cache=cache)
    result = app.acquire_token_silent_with_error(
        scopes=scope, account=account)
    return result

def msal_delegated_refresh_force(clientID, scope, authority, account):
    persistence = msal_persistence("token_cache.bin")
    cache = PersistedTokenCache(persistence)
    
    app = msal.PublicClientApplication(
        client_id=clientID, authority=authority, token_cache=cache)
    result = app.acquire_token_silent_with_error(
        scopes=scope, account=account, force_refresh=True)
    return result

def msal_delegated_device_flow(clientID, scope, authority):
    print("Initiate Device Code Flow to get an AAD Access Token.")
    print("Open a browser window and paste in the URL below and then enter the Code. CTRL+C to cancel.")
    
    persistence = msal_persistence("token_cache.bin")
    cache = PersistedTokenCache(persistence)
    
    app = msal.PublicClientApplication(client_id=clientID, authority=authority, token_cache=cache)
    flow = app.initiate_device_flow(scopes=scope)

    if "user_code" not in flow:
        raise ValueError("Fail to create device flow. Err: %s" % json.dumps(flow, indent=4))

    print(flow["message"])
    sys.stdout.flush()

    result = app.acquire_token_by_device_flow(flow)
    return result

def msal_jwt_expiry(accessToken):
    decodedAccessToken = jwt.decode(accessToken, verify=False)
    accessTokenFormatted = json.dumps(decodedAccessToken, indent=2)

    # Token Expiry
    tokenExpiry = datetime.fromtimestamp(int(decodedAccessToken['exp']))
    print("Token Expires at: " + str(tokenExpiry))
    return tokenExpiry

def msgraph_request(resource, requestHeaders):
    # Request
    results = requests.get(resource, headers=requestHeaders).json()
    return results

accounts = msal_cache_accounts(clientID, authority)

if accounts:
    for account in accounts:
        if account['username'] == username:
            myAccount = account
            print("Found account in MSAL Cache: " + account['username'])
            print("Obtaining a new Access Token using the Refresh Token")
            result = msal_delegated_refresh(clientID, scope, authority, myAccount)

            if result is None:
                # Get a new Access Token using the Device Code Flow
                result = msal_delegated_device_flow(clientID, scope, authority)
            else:
                if result["access_token"]:
                    msal_jwt_expiry(result["access_token"])                    
else:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)

    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

# Query AAD Users based on voice query using DisplayName
print(result['access_token'])
print(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/34bf6b0f-621e-4159-b916-d16e153678bd/items")
requestHeaders = {'Authorization': 'Bearer ' + result["access_token"],'Content-Type': 'application/json'}
queryResults = msgraph_request(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/34bf6b0f-621e-4159-b916-d16e153678bd/items",requestHeaders)

print(json.dumps(queryResults, indent=2))

# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])   

print(json.dumps(queryResults, indent=2))