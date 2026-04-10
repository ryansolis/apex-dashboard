"""
Downloads the loan spreadsheet from SharePoint.
Uses ClientCredential (app-based auth) instead of UserCredential,
because Microsoft 365 has deprecated legacy username/password auth
(the "binary security token" flow), causing the old approach to fail.

Required environment / GitHub Secrets:
  SHAREPOINT_CLIENT_ID     - Azure App Registration Application (client) ID
  SHAREPOINT_CLIENT_SECRET - Azure App Registration client secret
"""
import os
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse

from office365.sharepoint.client_context import ClientContext

def load_dotenv(path=".env"):
    """Very small .env loader for local runs (no extra dependency)."""
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


def parse_paths(raw):
    if not raw:
        return []
    return [p.strip() for p in raw.split("|") if p.strip()]


def warn_cert_expiry():
    """
    Warn when configured cert expiry is near.
    This is metadata-only and does not inspect key files.
    """
    raw_exp = (os.environ.get("SHAREPOINT_CERT_EXPIRES_ON") or "").strip()
    if not raw_exp:
        return
    try:
        expires_on = datetime.strptime(raw_exp, "%Y-%m-%d").date()
    except ValueError:
        print(
            "WARNING: SHAREPOINT_CERT_EXPIRES_ON should be YYYY-MM-DD "
            f"(got '{raw_exp}')."
        )
        return

    days_left = (expires_on - date.today()).days
    threshold_raw = (os.environ.get("SHAREPOINT_CERT_WARN_DAYS") or "30").strip()
    try:
        threshold = int(threshold_raw)
    except ValueError:
        threshold = 30

    if days_left < 0:
        print(
            f"WARNING: Configured certificate expired {-days_left} day(s) ago "
            f"({expires_on.isoformat()}). Rotate immediately."
        )
    elif days_left <= threshold:
        print(
            f"WARNING: Certificate expires in {days_left} day(s) "
            f"on {expires_on.isoformat()}. Rotate soon."
        )


def get_tenant_from_site_url(site_url):
    host = urlparse(site_url).netloc
    if not host:
        return ""
    return host.split(".")[0]


def normalize_tenant_id(raw_tenant):
    """
    Normalize tenant identifier for MSAL authority.
    Accepts:
      - tenant GUID
      - verified domain (contoso.onmicrosoft.com or custom domain)
      - short name (contoso) -> contoso.onmicrosoft.com
    """
    tenant = (raw_tenant or "").strip()
    if not tenant:
        return ""
    # GUID format
    if tenant.count("-") == 4 and len(tenant) >= 32:
        return tenant
    # Domain already provided
    if "." in tenant:
        return tenant
    # Short tenant alias from hostname prefix
    return f"{tenant}.onmicrosoft.com"


def build_context(site_url, client_id, client_secret):
    """
    Build SharePoint client context.
    If certificate variables are present, try certificate auth first.
    Falls back to client secret auth to keep CI/backward compatibility.
    """
    cert_path = (os.environ.get("SHAREPOINT_CERT_PATH") or "").strip()
    cert_thumbprint = (os.environ.get("SHAREPOINT_CERT_THUMBPRINT") or "").strip()
    tenant = normalize_tenant_id(os.environ.get("SHAREPOINT_TENANT_ID"))
    if not tenant:
        tenant = normalize_tenant_id(get_tenant_from_site_url(site_url))

    if cert_path and cert_thumbprint:
        try:
            print(f"Auth mode: certificate ({cert_path})")
            print(f"Using tenant authority: {tenant}")
            return ClientContext(site_url).with_client_certificate(
                tenant=tenant,
                client_id=client_id,
                thumbprint=cert_thumbprint,
                cert_path=cert_path,
            )
        except Exception as e:
            print(f"Certificate auth init failed, falling back to client secret: {e}")

    if client_secret:
        print("Auth mode: client secret")
        return ClientContext(site_url).with_client_credentials(client_id, client_secret)

    print(
        "ERROR: No usable auth method configured.\n"
        "Provide either:\n"
        "  • SHAREPOINT_CLIENT_ID + SHAREPOINT_CLIENT_SECRET, or\n"
        "  • SHAREPOINT_CLIENT_ID + SHAREPOINT_CERT_PATH + SHAREPOINT_CERT_THUMBPRINT"
    )
    sys.exit(1)


def discover_file_path(ctx, target_file_name, roots):
    """
    Try to discover a server-relative file path by crawling folder trees.
    Returns discovered path or "".
    """
    seen = set()
    queue = list(roots)
    max_visits = 250
    visits = 0

    while queue and visits < max_visits:
        current = queue.pop(0)
        if current in seen:
            continue
        seen.add(current)
        visits += 1
        try:
            folder = ctx.web.get_folder_by_server_relative_url(current).expand(
                ["Files", "Folders"]
            )
            ctx.load(folder)
            ctx.execute_query()
        except Exception:
            continue

        for f in folder.files:
            if (f.name or "").strip().lower() == target_file_name.lower():
                return f.serverRelativeUrl

        for sub in folder.folders:
            sub_url = (sub.serverRelativeUrl or "").strip()
            if not sub_url:
                continue
            # Skip hidden/system folders to keep traversal fast.
            lowered = sub_url.lower()
            if lowered.endswith("/forms") or "/_catalogs/" in lowered:
                continue
            queue.append(sub_url)

    return ""


_loaded = load_dotenv(".env")
if _loaded:
    print("Loaded local .env file.")

warn_cert_expiry()

client_id = (os.environ.get("SHAREPOINT_CLIENT_ID") or "").strip()
client_secret = (os.environ.get("SHAREPOINT_CLIENT_SECRET") or "").strip()

if not client_id:
    print("ERROR: Set environment variable before running:")
    print('  $env:SHAREPOINT_CLIENT_ID    = "<Application (client) ID from Azure app Overview>"')
    print("Or create a local .env file (ignored by git).")
    sys.exit(1)

print(f"Using app id prefix {client_id[:8]}…")
if client_secret:
    print(f"Client secret length: {len(client_secret)} chars.")

site_url = (
    os.environ.get("SHAREPOINT_SITE_URL")
    or "https://apexfunding.sharepoint.com/sites/ApexFunding"
).strip()
site_path = (urlparse(site_url).path or "").rstrip("/")
target_file_name = (os.environ.get("SHAREPOINT_FILE_NAME") or "").strip()
if not target_file_name:
    target_file_name = "Loan Pipeline Checklist.xlsx"

print(f"Connecting to: {site_url}")

ctx = build_context(site_url, client_id, client_secret)

paths_to_try = parse_paths(os.environ.get("SHAREPOINT_FILE_PATHS"))
if not paths_to_try:
    paths_to_try = [
        f"{site_path}/Shared Documents/{target_file_name}",
        f"{site_path}/Shared Documents/General/{target_file_name}",
        f"{site_path}/Documents/{target_file_name}",
        f"{site_path}/Documents/General/{target_file_name}",
    ]

# Preflight auth check before trying file paths (helps isolate 401 root cause)
try:
    ctx.web.get().execute_query()
    print("SharePoint auth preflight OK.")
except Exception as e:
    print(f"Auth preflight failed: {e}")
    print("If this is 401, the issue is app auth/consent, not file path.")

downloaded = False
saw_401 = False
for path in paths_to_try:
    try:
        print(f"Trying: {path}")
        f_obj = ctx.web.get_file_by_server_relative_url(path)
        with open("spreadsheet.xlsx", "wb") as f:
            f_obj.download(f)
            # execute_query must run while file handle is open
            ctx.execute_query()
        size = os.path.getsize("spreadsheet.xlsx")
        if size > 5000:
            print(f"Success! Downloaded {size:,} bytes -> spreadsheet.xlsx")
            downloaded = True
            break
        print(f"  File too small ({size} bytes), trying next...")
    except Exception as e:
        err = str(e)
        print(f"  Failed: {e}")
        if "401" in err or "Unauthorized" in err:
            saw_401 = True

if not downloaded:
    if saw_401:
        print(
            "\nERROR: SharePoint returned 401 Unauthorized (app credentials or permissions).\n"
            "Checklist:\n"
            "  • Ensure SHAREPOINT_CLIENT_ID is the exact app where consent was granted.\n"
            "  • Client ID and secret match the Azure app registration (no extra spaces).\n"
            "  • Secret is current — Azure client secrets expire; create a new one if needed.\n"
            "  • API permissions: under SharePoint (not only Graph) add Application →\n"
            "    Sites.Read.All (or Sites.Selected + grant this site in SharePoint admin).\n"
            "    Then Grant admin consent.\n"
            "  • Admin consent has been granted for the app in Entra ID.\n"
            "  • Tenant allows app-only access to this site collection.\n"
            "\n(401 is not 'file not found' — the server rejected authentication before path checks.)\n"
        )
    else:
        # Attempt automatic path discovery as final fallback.
        print("Known paths failed; attempting automatic discovery...")
        roots = [
            f"{site_path}/Shared Documents",
            f"{site_path}/Documents",
        ]
        discovered = discover_file_path(ctx, target_file_name, roots)
        if discovered:
            print(f"Discovered file path: {discovered}")
            try:
                f_obj = ctx.web.get_file_by_server_relative_url(discovered)
                with open("spreadsheet.xlsx", "wb") as f:
                    f_obj.download(f)
                    # execute_query must run while file handle is open
                    ctx.execute_query()
                size = os.path.getsize("spreadsheet.xlsx")
                if size > 5000:
                    print(f"Success! Downloaded {size:,} bytes -> spreadsheet.xlsx")
                    sys.exit(0)
            except Exception as e:
                print(f"Discovered path download failed: {e}")

        print("ERROR: Could not download spreadsheet from any known/discovered path.")
        print(
            "Set SHAREPOINT_FILE_PATHS in .env with the exact server-relative path(s), separated by '|'."
        )
    sys.exit(1)
