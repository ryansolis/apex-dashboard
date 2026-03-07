"""
Downloads the loan spreadsheet from SharePoint.
Uses requests + Microsoft authentication.
"""
import os, sys, subprocess

subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
import requests

user = os.environ.get("SHAREPOINT_USER")
pwd  = os.environ.get("SHAREPOINT_PASS")

if not all([user, pwd]):
    print("ERROR: SHAREPOINT_USER or SHAREPOINT_PASS secret is missing.")
    sys.exit(1)

# Direct download URL for the file (using SharePoint REST API format)
SITE    = "https://apexfunding.sharepoint.com/sites/ApexFunding"
# Try multiple possible paths for the file
FILE_PATHS = [
    "/sites/ApexFunding/_api/web/GetFileByServerRelativeUrl('/sites/ApexFunding/Shared%20Documents/Loan%20Pipeline%20Checklist.xlsx')/$value",
    "/sites/ApexFunding/_api/web/GetFileByServerRelativeUrl('/sites/ApexFunding/Shared%20Documents/General/Loan%20Pipeline%20Checklist.xlsx')/$value",
    "/sites/ApexFunding/_api/web/GetFileByServerRelativeUrl('/sites/ApexFunding/Documents/Loan%20Pipeline%20Checklist.xlsx')/$value",
]

print(f"Authenticating as {user}...")

# Get auth digest token via SharePoint's contextinfo endpoint
session = requests.Session()
session.auth = (user, pwd)

# Try form digest approach first
ctx_resp = session.post(
    f"{SITE}/_api/contextinfo",
    headers={"Accept": "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose"}
)

if ctx_resp.status_code == 200:
    print("Authentication successful!")
    downloaded = False
    for path in FILE_PATHS:
        url = f"https://apexfunding.sharepoint.com{path}"
        print(f"Trying: {path.split('/')[-1]}")
        r = session.get(url, headers={"Accept": "application/json;odata=verbose"})
        if r.status_code == 200 and len(r.content) > 5000:
            with open("spreadsheet.xlsx", "wb") as f:
                f.write(r.content)
            print(f"Downloaded {len(r.content):,} bytes → spreadsheet.xlsx")
            downloaded = True
            break
        else:
            print(f"  Status {r.status_code}, size {len(r.content)}")
    if not downloaded:
        print("Could not find file. Listing available files...")
        list_url = f"{SITE}/_api/web/GetFolderByServerRelativeUrl('/sites/ApexFunding/Shared Documents')/Files"
        r = session.get(list_url, headers={"Accept": "application/json;odata=verbose"})
        if r.status_code == 200:
            files = r.json().get("d", {}).get("results", [])
            print(f"Files found: {[f['Name'] for f in files]}")
        sys.exit(1)
else:
    print(f"Auth failed with status {ctx_resp.status_code}")
    print("Note: If your account uses MFA or SSO, basic auth won't work.")
    print("Response:", ctx_resp.text[:500])
    sys.exit(1)
