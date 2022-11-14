from O365 import Account
credentials = ('b80e63c4-5843-44da-badc-83a20c33dff1', 'ZoV8Q~fBGqAfI9QY2ALHmMLf7zOfOM7D3lTYZa4t')

# the default protocol will be Microsoft Graph
# the default authentication method will be "on behalf of a user"

account = Account(credentials,tenant_id='59daf140-4aee-4b77-80f4-4ea8bec86c2e')
if account.authenticate(scopes=['https://graph.microsoft.com/Sites.Manage.All', 'https://graph.microsoft.com/Sites.Read.All']):
   print('Authenticated!')
   print(account)

sp_site = account.sharepoint().get_site('bcaoffice365.sharepoint.com','/sites/imo_b')
sp_site_list = sp_site.get_list_by_name('Lokasi Penyimpanan')
sp_list_items = sp_site_list.get_items()
sp_item_by_id = sp_site_list.get_item_by_id('27')
# print(sp_list_items)
# print(sp_item_by_id)

# 'basic' adds: 'offline_access' and 'https://graph.microsoft.com/User.Read'
# 'message_all' adds: 'https://graph.microsoft.com/Mail.ReadWrite' and 'https://graph.microsoft.com/Mail.Send'