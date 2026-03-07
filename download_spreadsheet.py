"""
Downloads the loan spreadsheet from SharePoint using Microsoft credentials.
Credentials are stored as GitHub Secrets — never in this file.
"""
import os
import sys
import subprocess

# Install required library
subprocess.check_call([sys.executable, "-m", "pip", "install", "Office365-REST-Python-Client", "-q"])

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

user = os.environ.get("SHAREPOINT_USER")
pwd  = os.environ.get("SHAREPOINT_PASS")
url  = os.environ.get("SHAREPOINT_FILE_URL", "")

if not all([user, pwd, url]):
    print("ERROR: SHAREPOINT_USER, SHAREPOINT_PASS, or SHAREPOINT_FILE_URL secret is missing.")
    sys.exit(1)

# Extract SharePoint site URL
import re
site_match = re.match(r"(https://[^/]+/sites/[^/]+)", url)
root_match  = re.match(r"(https://[^/]+)", url)
site_url = site_match.group(1) if site_match else (root_match.group(1) if root_match else None)

if not site_url:
    print(f"ERROR: Could not parse SharePoint site URL from: {url}")
    sys.exit(1)

print(f"Connecting to: {site_url} as {user}")

credentials = UserCredential(user, pwd)
ctx = ClientContext(site_url).with_credentials(credentials)

# Common paths to try
paths_to_try = [
    "/sites/ApexFunding/Shared Documents/Loan Pipeline Checklist.xlsx",
    "/sites/ApexFunding/Shared Documents/General/Loan Pipeline Checklist.xlsx",
    "/sites/ApexFunding/Documents/Loan Pipeline Checklist.xlsx",
    "/sites/ApexFunding/Shared%20Documents/Loan%20Pipeline%20Checklist.xlsx",
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
        if size > 1000:
            print(f"Success! Downloaded {size:,} bytes → spreadsheet.xlsx")
            downloaded = True
            break
    except Exception as e:
        print(f"  Failed: {e}")

if not downloaded:
    print("ERROR: Could not find the spreadsheet in SharePoint.")
    print("Please check the file location and permissions.")
    sys.exit(1)
