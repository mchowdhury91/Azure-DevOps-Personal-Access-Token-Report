import os
from getpass import getpass

ORGANIZATION = ""
AAD_TENANT_ID = ""
API_BASE_URL = "https://vssps.dev.azure.com/"

CLIENT_ID = "" # Application (client) ID of app registration

#CLIENT_SECRET = "" # Placeholder - for use ONLY during testing.
# In a production app, we recommend you use a more secure method of storing your secret,
# like Azure Key Vault. Or, use an environment variable as described in Flask's documentation:
# https://flask.palletsprojects.com/en/1.1.x/config/#configuring-from-environment-variables
CLIENT_SECRET = getpass("Enter the app client secret: ")
if not CLIENT_SECRET:
    raise ValueError("Enter a valid CLIENT_SECRET or pull from a secure location such as Azure KeyVault")

AUTHORITY = "https://login.microsoftonline.com/" + AAD_TENANT_ID  # For multi-tenant app
# AUTHORITY = "https://login.microsoftonline.com/Enter_the_Tenant_Name_Here"

REDIRECT_PATH = "/getAToken"  # Used for forming an absolute URL to your redirect URI.
                              # The absolute URL must match the redirect URI you set
                              # in the app's registration in the Azure portal.

# You can find more Microsoft Graph API endpoints from Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer
#PATSENDPOINT = API_BASE_URL + ORGANIZATION + '/_apis/Tokens/Pats?api-version=6.1-preview&displayFilterOption=All'  # This resource requires no admin consent
USERPATENDPOINT = API_BASE_URL + ORGANIZATION + '/_apis/tokenadmin/personalaccesstokens/{subjectDescriptor}?api-version=6.1-preview.1'
USERLISTENDPOINT = API_BASE_URL + ORGANIZATION + '/_apis/graph/users?api-version=6.0-preview.1'

# You can find the proper permission names from this document
# https://docs.microsoft.com/en-us/graph/permissions-reference
SCOPE = ["499b84ac-1321-427f-aa17-267ca6975798/.default"]

SESSION_TYPE = "filesystem"  # Specifies the token cache should be stored in server-side session
