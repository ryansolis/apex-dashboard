"""
Convert a PFX certificate bundle into PEM files for SharePoint certificate auth.

Outputs:
  - private key PEM (required by download_spreadsheet.py cert auth mode)
  - certificate PEM (optional, informational)

Usage:
  python convert_pfx_to_pem.py --pfx certs/apex-dashboard-sp-private.pfx --out-dir certs
"""
import argparse
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Convert PFX to PEM private key/cert.")
    parser.add_argument("--pfx", required=True, help="Path to .pfx file")
    parser.add_argument("--out-dir", default="certs", help="Output directory for PEM files")
    parser.add_argument(
        "--password",
        default=None,
        help="PFX password (omit to prompt securely)",
    )
    args = parser.parse_args()

    try:
        from cryptography.hazmat.primitives import serialization
        from cryptography.hazmat.primitives.serialization import pkcs12
    except Exception:
        raise SystemExit(
            "Missing dependency: cryptography\n"
            "Install with: python -m pip install cryptography"
        )

    pfx_path = Path(args.pfx)
    if not pfx_path.exists():
        raise SystemExit(f"PFX file not found: {pfx_path}")

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.password is None:
        import getpass

        password = getpass.getpass("Enter PFX password: ")
    else:
        password = args.password

    data = pfx_path.read_bytes()
    private_key, cert, _additional = pkcs12.load_key_and_certificates(
        data, password.encode("utf-8")
    )

    if private_key is None:
        raise SystemExit("No private key found in PFX.")
    if cert is None:
        raise SystemExit("No certificate found in PFX.")

    key_out = out_dir / "apex-dashboard-sp-private-key.pem"
    cert_out = out_dir / "apex-dashboard-sp-certificate.pem"

    key_out.write_bytes(
        private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    cert_out.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

    print(f"Created private key PEM: {key_out}")
    print(f"Created certificate PEM: {cert_out}")
    print("Use the private key PEM path in SHAREPOINT_CERT_PATH.")


if __name__ == "__main__":
    main()
