"""
Downloads the loan spreadsheet from SharePoint.
Uses ClientCredential (app-based auth) instead of UserCredential,
because Microsoft 365 has deprecated legacy username/password auth
(the "binary security token" flow), causing the old approach to fail.

Required GitHub Secrets:
  SHAREPOINT_CLIENT_ID     - Azure App Registration Application (client) ID
    SHAREPOINT_CLIENT_SECRET - Azure App Registration client secret
    """
import os, sys

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

client_id     = os.environ.get("SHAREPOINT_CLIENT_ID")
client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")

if not all([client_id, client_secret]):
        print("ERROR: SHAREPOINT_CLIENT_ID or SHAREPOINT_CLIENT_SECRET secret is missing.")
        print("These replace the old SHAREPOINT_USER / SHAREPOINT_PASS secrets.")
        print("Register an Azure App and grant it SharePoint Sites.Read.All permission.")
        sys.exit(1)

site_url = "https://apexfunding.sharepoint.com/sites/ApexFunding"

print(f"Connecting to: {site_url}")
print(f"Using client_id: {client_id[:8]}...")

credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(site_url).with_credentials(credentials)

paths_to_try = [
        "/sites/ApexFunding/Shared Documents/Loan Pipeline Checklist.xlsx",
        "/sites/ApexFunding/Shared Documents/General/Loan Pipeline Checklist.xlsx",
        "/sites/ApexFunding/Documents/Loan Pipeline Checklist.xlsx",
]

downloaded = False
for path in paths_to_try:
        try:
                    print(f"Trying: {path}")
                    f_obj = ctx.web.get_file_by_server_relative_url(path)
                    with open("spreadsheet.xlsx", "wb") as f:
                                    f_obj.download(f)
                                ctx.execute_query()
                    size = os.path.getsize("spreadsheet.xlsx")
                    if size > 5000:
                                    print(f"Success! Downloaded {size:,} bytes -> spreadsheet.xlsx")
                                    downloaded = True
                                    break
        else:
                        print(f"  File too small ({size} bytes), trying next...")
except Exception as e:
        print(f"  Failed: {e}")

if not downloaded:
        print("ERROR: Could not find spreadsheet in SharePoint.")
        sys.exit(1)
