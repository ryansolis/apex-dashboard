"""
Downloads the loan spreadsheet from SharePoint.
Uses Office365-REST-Python-Client for proper Microsoft auth.
"""
import os, sys, re, subprocess

subprocess.check_call([sys.executable, "-m", "pip", "install", "Office365-REST-Python-Client", "-q"])

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

user = os.environ.get("SHAREPOINT_USER")
pwd  = os.environ.get("SHAREPOINT_PASS")

if not all([user, pwd]):
    print("ERROR: SHAREPOINT_USER or SHAREPOINT_PASS secret is missing.")
    sys.exit(1)

host_match = re.match(r"(https://[^/]+)", "https://apexfunding.sharepoint.com")
site_url = "https://apexfunding.sharepoint.com/sites/ApexFunding"

print(f"Connecting to: {site_url}")
print(f"As user: {user}")

credentials = UserCredential(user, pwd)
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
