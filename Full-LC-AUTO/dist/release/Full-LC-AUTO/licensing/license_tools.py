from __future__ import annotations

import argparse
import base64
import json
import re
import shutil
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa

DEFAULT_KEY_DIRECTORY = Path('C:/Full-LC-AUTO-License-Keys')
DEFAULT_PRIVATE_KEY = DEFAULT_KEY_DIRECTORY / 'license_private_key.pem'
DEFAULT_PUBLIC_KEY = DEFAULT_KEY_DIRECTORY / 'license_public_key.pem'
DEFAULT_APP_PUBLIC_KEY = Path(__file__).resolve().parent / 'license_public_key.pem'
DEFAULT_LICENSES_JSON = Path(__file__).resolve().parent / 'licenses.json'
DEFAULT_LICENSES_SIG = Path(__file__).resolve().parent / 'licenses.sig'
DEFAULT_CUSTOMER_LICENSE_JSON = Path(__file__).resolve().parent / 'customer_license.json'


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def prompt_path(prompt_text: str, default: Path) -> Path:
    raw_value = input(f"{prompt_text} [{default}]: ").strip()
    return Path(raw_value) if raw_value else default


def prompt_yes_no(prompt_text: str, default: bool = True) -> bool:
    suffix = 'Y/n' if default else 'y/N'
    raw_value = input(f"{prompt_text} [{suffix}]: ").strip().lower()
    if not raw_value:
        return default
    return raw_value in {'y', 'yes'}


def copy_public_key_to_app(public_key_path: Path, app_public_key_path: Path) -> None:
    ensure_parent(app_public_key_path)
    shutil.copy2(public_key_path, app_public_key_path)


def load_licenses_payload(json_path: Path) -> dict[str, object]:
    try:
        payload = json.loads(json_path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return {'version': 1, 'licenses': []}
    if not isinstance(payload, dict):
        raise SystemExit('licenses.json must contain a JSON object.')
    licenses = payload.get('licenses')
    if licenses is None:
        payload['licenses'] = []
    elif not isinstance(licenses, list):
        raise SystemExit('licenses.json field "licenses" must be a list.')
    return payload


def save_licenses_payload(json_path: Path, payload: dict[str, object]) -> None:
    ensure_parent(json_path)
    json_path.write_text(json.dumps(payload, indent=2), encoding='utf-8', newline='\n')


def write_customer_license_file(customer_license_path: Path, license_key: str, customer_name: str) -> None:
    ensure_parent(customer_license_path)
    payload = {
        'license_key': license_key.strip(),
        'customer_name': customer_name.strip(),
    }
    customer_license_path.write_text(json.dumps(payload, indent=2), encoding='utf-8', newline='\n')


def build_auto_license_key(licenses: list[object], customer_name: str) -> str:
    cleaned_name = re.sub(r'[^A-Za-z0-9]+', '-', customer_name.strip().upper()).strip('-') or 'CUSTOMER'
    prefix = cleaned_name[:20]
    max_suffix = 0
    for entry in licenses:
        if not isinstance(entry, dict):
            continue
        existing_key = str(entry.get('license_key', '')).strip().upper()
        match = re.fullmatch(rf'{re.escape(prefix)}-(\d{{3}})', existing_key)
        if match:
            max_suffix = max(max_suffix, int(match.group(1)))
    return f'{prefix}-{max_suffix + 1:03d}'


def add_license_entry(
    json_path: Path,
    license_key: str | None,
    customer_name: str,
    machine_id: str,
    expiry_date: str,
    allowed_version: str,
    status: str,
) -> str:
    payload = load_licenses_payload(json_path)
    licenses = payload['licenses']
    normalized_key = (license_key or '').strip() or build_auto_license_key(licenses, customer_name)
    for entry in licenses:
        if isinstance(entry, dict) and str(entry.get('license_key', '')).strip() == normalized_key:
            raise SystemExit(f'License key already exists: {normalized_key}')

    licenses.append(
        {
            'license_key': normalized_key,
            'customer_name': customer_name.strip(),
            'status': status.strip().lower(),
            'expiry_date': expiry_date.strip(),
            'machine_id': machine_id.strip(),
            'allowed_version': allowed_version.strip(),
        }
    )
    save_licenses_payload(json_path, payload)
    return normalized_key


def generate_keypair(private_out: Path, public_out: Path) -> None:
    ensure_parent(private_out)
    ensure_parent(public_out)
    private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    private_bytes = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )
    public_bytes = private_key.public_key().public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo,
    )
    private_out.write_bytes(private_bytes)
    public_out.write_bytes(public_bytes)


def sign_json(private_key_path: Path, json_path: Path, signature_path: Path) -> None:
    ensure_parent(signature_path)
    private_key = serialization.load_pem_private_key(private_key_path.read_bytes(), password=None)
    payload = json_path.read_bytes()
    signature = private_key.sign(payload, padding.PKCS1v15(), hashes.SHA256())
    signature_path.write_text(base64.b64encode(signature).decode('ascii'), encoding='utf-8')


def verify_signed_file(public_key_path: Path, json_path: Path, signature_path: Path) -> None:
    public_key = serialization.load_pem_public_key(public_key_path.read_bytes())
    payload = json_path.read_bytes()
    signature_text = signature_path.read_text(encoding='utf-8').strip()
    signature = base64.b64decode(signature_text.encode('ascii'), validate=True)
    public_key.verify(signature, payload, padding.PKCS1v15(), hashes.SHA256())


def run_interactive() -> None:
    print('License tools')
    print('1. Create keypair in C:')
    print('2. Validate a signed file')
    print('3. Sign a JSON file')
    print('4. Add a license entry')
    print('5. Add a license entry and sign')
    print('6. Add a license entry, sign, and write customer license file')
    choice = input('Choose 1, 2, 3, 4, 5, or 6: ').strip()

    if choice == '1':
        private_out = prompt_path('Private key output path', DEFAULT_PRIVATE_KEY)
        public_out = prompt_path('Public key output path', DEFAULT_PUBLIC_KEY)
        generate_keypair(private_out, public_out)
        print(f'Private key saved to: {private_out}')
        print(f'Public key saved to: {public_out}')
        if prompt_yes_no(f'Copy public key into app repo at {DEFAULT_APP_PUBLIC_KEY}?', True):
            copy_public_key_to_app(public_out, DEFAULT_APP_PUBLIC_KEY)
            print(f'App public key updated at: {DEFAULT_APP_PUBLIC_KEY}')
        return

    if choice == '2':
        public_key_path = prompt_path('Public key path', DEFAULT_PUBLIC_KEY)
        json_path = prompt_path('Signed JSON path', DEFAULT_LICENSES_JSON)
        signature_path = prompt_path('Signature path', DEFAULT_LICENSES_SIG)
        try:
            verify_signed_file(public_key_path, json_path, signature_path)
        except FileNotFoundError as exc:
            raise SystemExit(f'Missing file: {exc.filename}') from exc
        except InvalidSignature as exc:
            raise SystemExit('Signature verification failed.') from exc
        except Exception as exc:
            raise SystemExit(f'Verification failed: {exc}') from exc
        print('Signature is valid.')
        return

    if choice == '3':
        private_key_path = prompt_path('Private key path', DEFAULT_PRIVATE_KEY)
        json_path = prompt_path('JSON path to sign', DEFAULT_LICENSES_JSON)
        signature_path = prompt_path('Signature output path', DEFAULT_LICENSES_SIG)
        sign_json(private_key_path, json_path, signature_path)
        print(f'Signature saved to: {signature_path}')
        return

    if choice == '4':
        json_path = prompt_path('licenses.json path', DEFAULT_LICENSES_JSON)
        license_key = input('License key [auto]: ').strip()
        customer_name = input('Customer name: ').strip()
        machine_id = input('Machine ID: ').strip()
        expiry_date = input('Expiry date (YYYY-MM-DD): ').strip()
        allowed_version = input('Allowed app version [1.0.0]: ').strip() or '1.0.0'
        status = input('Status [active]: ').strip() or 'active'
        created_key = add_license_entry(
            json_path=json_path,
            license_key=license_key,
            customer_name=customer_name,
            machine_id=machine_id,
            expiry_date=expiry_date,
            allowed_version=allowed_version,
            status=status,
        )
        print(f'License key: {created_key}')
        print(f'License added to: {json_path}')
        print('Re-sign licenses.json before using this license.')
        return

    if choice == '5':
        json_path = prompt_path('licenses.json path', DEFAULT_LICENSES_JSON)
        signature_path = prompt_path('Signature output path', DEFAULT_LICENSES_SIG)
        private_key_path = prompt_path('Private key path', DEFAULT_PRIVATE_KEY)
        license_key = input('License key [auto]: ').strip()
        customer_name = input('Customer name: ').strip()
        machine_id = input('Machine ID: ').strip()
        expiry_date = input('Expiry date (YYYY-MM-DD): ').strip()
        allowed_version = input('Allowed app version [1.0.0]: ').strip() or '1.0.0'
        status = input('Status [active]: ').strip() or 'active'
        created_key = add_license_entry(
            json_path=json_path,
            license_key=license_key,
            customer_name=customer_name,
            machine_id=machine_id,
            expiry_date=expiry_date,
            allowed_version=allowed_version,
            status=status,
        )
        sign_json(private_key_path, json_path, signature_path)
        print(f'License key: {created_key}')
        print(f'License added to: {json_path}')
        print(f'Signature saved to: {signature_path}')
        return

    if choice == '6':
        json_path = prompt_path('licenses.json path', DEFAULT_LICENSES_JSON)
        signature_path = prompt_path('Signature output path', DEFAULT_LICENSES_SIG)
        private_key_path = prompt_path('Private key path', DEFAULT_PRIVATE_KEY)
        customer_license_path = prompt_path('customer_license.json path', DEFAULT_CUSTOMER_LICENSE_JSON)
        license_key = input('License key [auto]: ').strip()
        customer_name = input('Customer name: ').strip()
        machine_id = input('Machine ID: ').strip()
        expiry_date = input('Expiry date (YYYY-MM-DD): ').strip()
        allowed_version = input('Allowed app version [1.0.0]: ').strip() or '1.0.0'
        status = input('Status [active]: ').strip() or 'active'
        created_key = add_license_entry(
            json_path=json_path,
            license_key=license_key,
            customer_name=customer_name,
            machine_id=machine_id,
            expiry_date=expiry_date,
            allowed_version=allowed_version,
            status=status,
        )
        sign_json(private_key_path, json_path, signature_path)
        write_customer_license_file(customer_license_path, created_key, customer_name)
        print(f'License key: {created_key}')
        print(f'License added to: {json_path}')
        print(f'Signature saved to: {signature_path}')
        print(f'Customer license file saved to: {customer_license_path}')
        return

    raise SystemExit('Invalid choice. Expected 1, 2, 3, 4, 5, or 6.')


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='License signing tools')
    subparsers = parser.add_subparsers(dest='command')

    generate_parser = subparsers.add_parser('generate-keypair')
    generate_parser.add_argument('--private-out', default=str(DEFAULT_PRIVATE_KEY))
    generate_parser.add_argument('--public-out', default=str(DEFAULT_PUBLIC_KEY))
    generate_parser.add_argument('--copy-public-to-app', action='store_true')
    generate_parser.add_argument('--app-public-out', default=str(DEFAULT_APP_PUBLIC_KEY))

    sign_parser = subparsers.add_parser('sign-json')
    sign_parser.add_argument('--private-key', default=str(DEFAULT_PRIVATE_KEY))
    sign_parser.add_argument('--json', required=True)
    sign_parser.add_argument('--sig', required=True)

    verify_parser = subparsers.add_parser('verify-json')
    verify_parser.add_argument('--public-key', default=str(DEFAULT_PUBLIC_KEY))
    verify_parser.add_argument('--json', required=True)
    verify_parser.add_argument('--sig', required=True)

    add_license_parser = subparsers.add_parser('add-license')
    add_license_parser.add_argument('--json', default=str(DEFAULT_LICENSES_JSON))
    add_license_parser.add_argument('--license-key')
    add_license_parser.add_argument('--customer-name', required=True)
    add_license_parser.add_argument('--machine-id', required=True)
    add_license_parser.add_argument('--expiry-date', required=True)
    add_license_parser.add_argument('--allowed-version', default='1.0.0')
    add_license_parser.add_argument('--status', default='active')

    add_license_sign_parser = subparsers.add_parser('add-license-and-sign')
    add_license_sign_parser.add_argument('--json', default=str(DEFAULT_LICENSES_JSON))
    add_license_sign_parser.add_argument('--sig', default=str(DEFAULT_LICENSES_SIG))
    add_license_sign_parser.add_argument('--private-key', default=str(DEFAULT_PRIVATE_KEY))
    add_license_sign_parser.add_argument('--license-key')
    add_license_sign_parser.add_argument('--customer-name', required=True)
    add_license_sign_parser.add_argument('--machine-id', required=True)
    add_license_sign_parser.add_argument('--expiry-date', required=True)
    add_license_sign_parser.add_argument('--allowed-version', default='1.0.0')
    add_license_sign_parser.add_argument('--status', default='active')

    add_license_bundle_parser = subparsers.add_parser('add-license-sign-and-write-customer')
    add_license_bundle_parser.add_argument('--json', default=str(DEFAULT_LICENSES_JSON))
    add_license_bundle_parser.add_argument('--sig', default=str(DEFAULT_LICENSES_SIG))
    add_license_bundle_parser.add_argument('--private-key', default=str(DEFAULT_PRIVATE_KEY))
    add_license_bundle_parser.add_argument('--customer-license-out', default=str(DEFAULT_CUSTOMER_LICENSE_JSON))
    add_license_bundle_parser.add_argument('--license-key')
    add_license_bundle_parser.add_argument('--customer-name', required=True)
    add_license_bundle_parser.add_argument('--machine-id', required=True)
    add_license_bundle_parser.add_argument('--expiry-date', required=True)
    add_license_bundle_parser.add_argument('--allowed-version', default='1.0.0')
    add_license_bundle_parser.add_argument('--status', default='active')
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if not args.command:
        run_interactive()
        return

    if args.command == 'generate-keypair':
        private_out = Path(args.private_out)
        public_out = Path(args.public_out)
        generate_keypair(private_out, public_out)
        print(f'Private key saved to: {private_out}')
        print(f'Public key saved to: {public_out}')
        if args.copy_public_to_app:
            app_public_out = Path(args.app_public_out)
            copy_public_key_to_app(public_out, app_public_out)
            print(f'App public key updated at: {app_public_out}')
        return

    if args.command == 'sign-json':
        sign_json(Path(args.private_key), Path(args.json), Path(args.sig))
        print(f'Signature saved to: {args.sig}')
        return

    if args.command == 'verify-json':
        try:
            verify_signed_file(Path(args.public_key), Path(args.json), Path(args.sig))
        except InvalidSignature as exc:
            raise SystemExit('Signature verification failed.') from exc
        print('Signature is valid.')
        return

    if args.command == 'add-license':
        created_key = add_license_entry(
            json_path=Path(args.json),
            license_key=args.license_key,
            customer_name=args.customer_name,
            machine_id=args.machine_id,
            expiry_date=args.expiry_date,
            allowed_version=args.allowed_version,
            status=args.status,
        )
        print(f'License key: {created_key}')
        print(f'License added to: {args.json}')
        print('Re-sign licenses.json before using this license.')
        return

    if args.command == 'add-license-and-sign':
        json_path = Path(args.json)
        signature_path = Path(args.sig)
        created_key = add_license_entry(
            json_path=json_path,
            license_key=args.license_key,
            customer_name=args.customer_name,
            machine_id=args.machine_id,
            expiry_date=args.expiry_date,
            allowed_version=args.allowed_version,
            status=args.status,
        )
        sign_json(Path(args.private_key), json_path, signature_path)
        print(f'License key: {created_key}')
        print(f'License added to: {args.json}')
        print(f'Signature saved to: {args.sig}')
        return

    if args.command == 'add-license-sign-and-write-customer':
        json_path = Path(args.json)
        signature_path = Path(args.sig)
        customer_license_out = Path(args.customer_license_out)
        created_key = add_license_entry(
            json_path=json_path,
            license_key=args.license_key,
            customer_name=args.customer_name,
            machine_id=args.machine_id,
            expiry_date=args.expiry_date,
            allowed_version=args.allowed_version,
            status=args.status,
        )
        sign_json(Path(args.private_key), json_path, signature_path)
        write_customer_license_file(customer_license_out, created_key, args.customer_name)
        print(f'License key: {created_key}')
        print(f'License added to: {args.json}')
        print(f'Signature saved to: {args.sig}')
        print(f'Customer license file saved to: {args.customer_license_out}')
        return

    raise SystemExit(f'Unknown command: {args.command}')


if __name__ == '__main__':
    main()
