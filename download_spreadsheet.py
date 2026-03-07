"""
Downloads the loan spreadsheet from SharePoint.
Credentials are stored as GitHub Secrets — never in this file.
"""
import os
import sys
import urllib.request
import urllib.error

url  = os.environ.get("SHAREPOINT_URL")
user = os.environ.get("SHAREPOINT_USER")
pwd  = os.environ.get("SHAREPOINT_PASS")

if not all([url, user, pwd]):
    print("ERROR: SHAREPOINT_URL, SHAREPOINT_USER, or SHAREPOINT_PASS secret is missing.")
    print("Add them in: Settings → Secrets and variables → Actions")
    sys.exit(1)

print(f"Downloading spreadsheet from SharePoint...")

# Build a password manager for basic/NTLM-style auth
password_mgr = urllib.request.HTTPPasswordMgrWithDefaultRealm()
password_mgr.add_password(None, url, user, pwd)
auth_handler = urllib.request.HTTPBasicAuthHandler(password_mgr)
opener = urllib.request.build_opener(auth_handler)

try:
    with opener.open(url, timeout=30) as response:
        data = response.read()
    with open("spreadsheet.xlsx", "wb") as f:
        f.write(data)
    print(f"Downloaded {len(data):,} bytes → spreadsheet.xlsx")
except urllib.error.HTTPError as e:
    print(f"HTTP Error {e.code}: {e.reason}")
    print("Check that your SharePoint URL and credentials are correct.")
    sys.exit(1)
except Exception as e:
    print(f"Download failed: {e}")
    sys.exit(1)
