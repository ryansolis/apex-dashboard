"""
Diagnose SharePoint app-only auth with explicit token + API checks.

Usage:
  python auth_diagnose.py

Reads environment variables (or .env):
  SHAREPOINT_CLIENT_ID
  SHAREPOINT_CLIENT_SECRET
Optional:
  SHAREPOINT_SITE_URL (default: https://apexfunding.sharepoint.com/sites/ApexFunding)
"""
import base64
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path


def load_dotenv(path=".env"):
    env_path = Path(path)
    if not env_path.exists():
        return False
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val
    return True


def decode_jwt_payload(token):
    parts = token.split(".")
    if len(parts) < 2:
        return {}
    payload = parts[1]
    payload += "=" * ((4 - len(payload) % 4) % 4)
    raw = base64.urlsafe_b64decode(payload.encode("utf-8"))
    return json.loads(raw.decode("utf-8"))


def post_form(url, data):
    body = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))


def get(url, bearer):
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {bearer}")
    req.add_header("Accept", "application/json;odata=nometadata")
    with urllib.request.urlopen(req) as resp:
        return resp.status, resp.read().decode("utf-8")


def main():
    if load_dotenv(".env"):
        print("Loaded local .env file.")

    client_id = (os.environ.get("SHAREPOINT_CLIENT_ID") or "").strip()
    client_secret = (os.environ.get("SHAREPOINT_CLIENT_SECRET") or "").strip()
    site_url = (
        os.environ.get("SHAREPOINT_SITE_URL")
        or "https://apexfunding.sharepoint.com/sites/ApexFunding"
    ).strip().rstrip("/")

    if not client_id or not client_secret:
        print("ERROR: Missing SHAREPOINT_CLIENT_ID or SHAREPOINT_CLIENT_SECRET.")
        sys.exit(1)

    host = urllib.parse.urlparse(site_url).netloc
    tenant = host.split(".")[0]
    token_url = f"https://login.microsoftonline.com/{tenant}.onmicrosoft.com/oauth2/v2.0/token"

    print(f"Client ID prefix: {client_id[:8]}…")
    print(f"Secret length: {len(client_secret)}")
    print(f"Site URL: {site_url}")
    print(f"Token endpoint: {token_url}")

    try:
        token_json = post_form(
            token_url,
            {
                "client_id": client_id,
                "client_secret": client_secret,
                "scope": f"https://{host}/.default",
                "grant_type": "client_credentials",
            },
        )
    except Exception as exc:
        print(f"\nTOKEN REQUEST FAILED: {exc}")
        print("This means credentials/tenant endpoint are not accepted.")
        sys.exit(2)

    access_token = token_json.get("access_token")
    if not access_token:
        print("\nTOKEN RESPONSE MISSING access_token")
        print(json.dumps(token_json, indent=2))
        sys.exit(3)

    claims = decode_jwt_payload(access_token)
    print("\nToken acquired.")
    print(f"aud: {claims.get('aud')}")
    print(f"tid: {claims.get('tid')}")
    print(f"appid/azp: {claims.get('appid') or claims.get('azp')}")
    roles = claims.get("roles") or []
    print(f"roles: {roles}")

    api_url = f"{site_url}/_api/web?$select=Title,Url"
    try:
        status, body = get(api_url, access_token)
        print(f"\nSharePoint API OK (status {status})")
        print(body[:500])
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="ignore")
        print(f"\nSharePoint API FAILED (status {exc.code})")
        print(raw[:1200])
        if exc.code == 401:
            print(
                "\n401 here means token accepted by Entra but rejected by SharePoint authorization/policy."
            )
        sys.exit(4)
    except Exception as exc:
        print(f"\nSharePoint API request error: {exc}")
        sys.exit(5)


if __name__ == "__main__":
    main()
