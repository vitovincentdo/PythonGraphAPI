import msal
import jwt
import json
import sys
import requests
import webbrowser
from datetime import datetime
from msal_extensions import *

graphURI = 'https://bcaoffice365.sharepoint.com/sites'
tenantID = '59daf140-4aee-4b77-80f4-4ea8bec86c2e'
authority = 'https://login.microsoftonline.com/' + tenantID
clientID = 'b80e63c4-5843-44da-badc-83a20c33dff1'
scope = ["Sites.Read.All", "Sites.ReadWrite.All", "Sites.Manage.All"]
username = 'vito_vincentdo@bca.co.id'
result = None
tokenExpiry = None
# testAT = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2JjYW9mZmljZTM2NS5zaGFyZXBvaW50LmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzU5ZGFmMTQwLTRhZWUtNGI3Ny04MGY0LTRlYThiZWM4NmMyZS8iLCJpYXQiOjE2NjgxNTkzODgsIm5iZiI6MTY2ODE1OTM4OCwiZXhwIjoxNjY4MTY0ODQzLCJhY3IiOiIxIiwiYWlvIjoiQVhRQWkvOFRBQUFBUU12ZGtvK2xCcHA4Ri93aE9iL1BXZENBU1RyS1pud20vL1dKeHRFQmVRWW4xYUx4ZHlJZkYwbEJqOWJwWlM4dys4dFBUMmZ0TDVDUmM5V1Q3SmNpbFNzcFJ1bTl0SmRXRzJkUFN3MWNoSGw2QWZITTlwWE9IUCtaQXB1RGNQekxjRXhtQnBUdGFvNkNPeTJKaUNITHNBPT0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6IlBsYXlHcmFwaEFQSSIsImFwcGlkIjoiYjgwZTYzYzQtNTg0My00NGRhLWJhZGMtODNhMjBjMzNkZmYxIiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJ2aW5jZW50ZG8iLCJnaXZlbl9uYW1lIjoiVklUTyIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjIwMi42LjIxMy4yMyIsIm5hbWUiOiJWSVRPIFZJTkNFTlRETyIsIm9pZCI6IjMyYTMzYzYxLTVlZWItNDA0MS1hYjEzLTQyZGRmOWQ1YTFlZiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS04NjI1Mjk5ODEtNTk0MDQ3Nzg3LTExMzYyNjM4NjAtMzY1MDg2IiwicHVpZCI6IjEwMDMyMDAwQUE2ODA4RDgiLCJyaCI6IjAuQVVrQVFQSGFXZTVLZDB1QTlFNm92c2hzTGdNQUFBQUFBUEVQemdBQUFBQUFBQUJKQUxVLiIsInNjcCI6IkZpbGVzLlJlYWQgRmlsZXMuUmVhZC5BbGwgRmlsZXMuUmVhZFdyaXRlIEZpbGVzLlJlYWRXcml0ZS5BbGwgR3JvdXAuUmVhZC5BbGwgb2ZmbGluZV9hY2Nlc3Mgb3BlbmlkIHByb2ZpbGUgU2l0ZXMuRnVsbENvbnRyb2wuQWxsIFNpdGVzLk1hbmFnZS5BbGwgU2l0ZXMuUmVhZC5BbGwgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUZWFtLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkIiwic2lkIjoiNmM2NDhlNGYtNzdkYS00NzQ1LTlkMDEtOWJkNzI5ZjU3NTY4Iiwic3ViIjoiOHlnZFBpRGhxVXFMNnktU0xLYm10bEJjR1RKN2VMVTFJLTRFbGdkRTZuQSIsInRpZCI6IjU5ZGFmMTQwLTRhZWUtNGI3Ny04MGY0LTRlYThiZWM4NmMyZSIsInVuaXF1ZV9uYW1lIjoidml0b192aW5jZW50ZG9AYmNhLmNvLmlkIiwidXBuIjoiVTA2ODg2M0BiY2EuY28uaWQiLCJ1dGkiOiJzQjJtb3RQOUpFLXB1VHR3V0FIM0FBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJmMmVmOTkyYy0zYWZiLTQ2YjktYjdjZi1hMTI2ZWU3NGM0NTEiLCJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXX0.G50J0k5fcP7yjkn8npkQkKL-RVXZhW8tps9jiOgjNTnsvLD7--16kpoQKsQyOXdouXfg3hOqEzTBoF59IkNa6YaGt-Hg0d1qRnBYCYNDY0Y6pzGm-T0MipY2-2cvX6cs1UIV2eVgb5cSNy87rNJUZhUGZ_tvMiIZ8m39gk4BJXdEGnz3JrYjyx_hrmsvF6u8PvLbcgkY7QowG0FgJyCqv5USsjxlcqqTadcQPRgLVQDmkKqbPoZ2x_gQqT3Kpk01cGNHYNiphm5rvA7mYAn7caiCxctvAergNex2JfV_YVWMVRwpl17HCxeXeRcHGK-QlqifECPb3_wK7pt79LuCVw'

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
    print(flow)
    webbrowser.open(flow['verification_uri'])

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
    results = requests.get(resource, headers=requestHeaders)
    print(results)
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
# print(testAT)
print(graphURI + "/imo_b/_api/web/lists/getbytitle('List Employee')/items")
requestHeaders = {'Authorization': 'Bearer ' + result["access_token"],'Accept': 'application/json; odata=verbose'}
queryResults = msgraph_request(graphURI + "sites/imo_b/_api/web/lists/getbytitle('List Employee')/items",requestHeaders)

print(queryResults)
# print(json.dumps(queryResults))
# print(json.dumps(queryResults, indent=2))

# Force Token Refresh
result =  msal_delegated_refresh_force(clientID, scope, authority, myAccount)
if result is None:
    # Get a new Access Token using the Device Code Flow
    result = msal_delegated_device_flow(clientID, scope, authority)
else:
    if result["access_token"]:
        msal_jwt_expiry(result["access_token"])   

# print(json.dumps(queryResults, indent=2))