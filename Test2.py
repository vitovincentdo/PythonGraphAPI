from shareplum import Site
from shareplum import Office365

authcookie = Office365('https://bcaoffice365.sharepoint.com', username='vito_vincentdo@bca.co.id', password='tyrex11').GetCookies()
site = Site('https://bcaoffice365.sharepoint.com/sites/imo_b/', authcookie=authcookie)