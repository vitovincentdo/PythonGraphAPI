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
# username = 'U545820@bca.co.id'
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
    results = []
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

    results.append(app.acquire_token_by_device_flow(flow))
    results.append(app.get_accounts()[0])

    return results

def msal_jwt_expiry(accessToken):
    decodedAccessToken = jwt.decode(accessToken, verify=False)
    accessTokenFormatted = json.dumps(decodedAccessToken, indent=2)

    # Token Expiry
    tokenExpiry = datetime.fromtimestamp(int(decodedAccessToken['exp']))
    print("Token Expires at: " + str(tokenExpiry))
    return tokenExpiry

def msgraph_request_get(resource, requestHeaders):
    # Request
    results = requests.get(resource, headers=requestHeaders).json()
    return results

def msgraph_request_post(resource, requestHeaders, jsonObject):
    # Request
    results = requests.post(resource, headers=requestHeaders, json=jsonObject).json()
    return results

def msgraph_request_patch(resource, requestHeaders, newData):
    # Request
    results = requests.patch(resource, headers=requestHeaders, data=newData).json()
    return results

def msgraph_request_delete(resource, requestHeaders):
    # Request
    results = requests.delete(resource, headers=requestHeaders)
    return results

accounts = msal_cache_accounts(clientID, authority)
# print(accounts)

if accounts:
    for account in accounts:
        # if account['username'] == username:
            # print(account)
            myAccount = account
            print("Found account in MSAL Cache: " + account['username'])
            print("Obtaining a new Access Token using the Refresh Token")
            result = msal_delegated_refresh(clientID, scope, authority, myAccount)

            if result is None:
                # Get a new Access Token using the Device Code Flow
                result = msal_delegated_device_flow(clientID, scope, authority)
                result[0]
            else:
                if result["access_token"]:
                    msal_jwt_expiry(result["access_token"])
else:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
    print(result)
    myAccount=result[1]
    if result[0]["access_token"]:
        msal_jwt_expiry(result[0]["access_token"])


# print(result)
# Query AAD Users based on voice query using DisplayName
# print(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/db38cd50-ad17-49c6-a06f-99d9ed17ffdf/items?expand=fields")
requestHeaders = {'Authorization': 'Bearer ' + result[0]["access_token"],'Content-Type': 'application/json', 'Accept': 'application/json'}
print('======================================')
print('Get All Master List License Data')
print('======================================')
queryResults = msgraph_request_get(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/69C54909-36C1-483A-9C3C-006B41CD6355/items?expand=fields&$top=10000",requestHeaders)

# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

print(json.dumps(queryResults, indent=2))
print('======================================\n')

print('===========================================')
print('Get Single Item From Master List License')
print('===========================================')
queryResults = msgraph_request_get(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/69C54909-36C1-483A-9C3C-006B41CD6355/items/1?expand=fields",requestHeaders)


# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

print(json.dumps(queryResults, indent=2))
print('===========================================\n')

print('===========================================')
print('Create New Item Into Master List License')
print('===========================================')

TestData = {
    "fields": {
        "Title": "GIMP",
        "field_2": "Open Source (GIMP Development Team",
        "field_3": "2.10.30",
        "field_4": "Free",
        "field_5": "-",
        "field_6": "Image Editor",
        "field_7": "Open file .tiff, convert .tiff ke .pdf dan sebaliknya, edit file .tiff. Bisa menjadi alternatif aplikasi MODI",
        "field_8": "29-Mar-22",
        "field_9": "Hendro Purwanto",
        "field_10": "Yes",
        "field_11": "ALL BCA",
        "field_12": ""
    }
}

queryResults = msgraph_request_post(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/69C54909-36C1-483A-9C3C-006B41CD6355/items",requestHeaders,TestData)


# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

print(json.dumps(queryResults, indent=2))
print('===========================================\n')

print('===========================================')
print('Update Item Into Master List License')
print('===========================================')

newData = {
    "Title": "Beyond Compare (Test)",
    "field_5": "- (Test)"
}

print(json.dumps(newData))

queryResults = msgraph_request_patch(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/69C54909-36C1-483A-9C3C-006B41CD6355/items/1/fields",requestHeaders,json.dumps(newData))


# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

print(json.dumps(queryResults, indent=2))
print('===========================================\n')

print('===========================================')
print('Delete an Item From Master List License')
print('===========================================')

queryResults = msgraph_request_delete(graphURI + "/v1.0/sites/bcaoffice365.sharepoint.com,3983b9b4-f1ea-4771-91cb-fbbcab357e2a,ff059f1c-9c7c-49d4-ac63-b73d11cd1c1f/lists/69C54909-36C1-483A-9C3C-006B41CD6355/items/9",requestHeaders)


# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])

print(queryResults)
print('===========================================\n')