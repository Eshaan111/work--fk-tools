from __future__ import annotations

import argparse
import base64
import json
import re
import shutil
import sys
import winreg
from datetime import date, timedelta
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa

def get_runtime_directory() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


RUNTIME_DIRECTORY = get_runtime_directory()
DEFAULT_KEY_DIRECTORY = Path('C:/Full-LC-AUTO-License-Keys')
DEFAULT_PRIVATE_KEY = DEFAULT_KEY_DIRECTORY / 'license_private_key.pem'
DEFAULT_PUBLIC_KEY = DEFAULT_KEY_DIRECTORY / 'license_public_key.pem'
DEFAULT_LICENSING_DIRECTORY = (
    RUNTIME_DIRECTORY / 'licensing'
    if (RUNTIME_DIRECTORY / 'licensing').exists()
    else RUNTIME_DIRECTORY
)
DEFAULT_APP_PUBLIC_KEY = DEFAULT_LICENSING_DIRECTORY / 'license_public_key.pem'
DEFAULT_LICENSES_JSON = DEFAULT_LICENSING_DIRECTORY / 'licenses.json'
DEFAULT_LICENSES_SIG = DEFAULT_LICENSING_DIRECTORY / 'licenses.sig'
DEFAULT_CUSTOMER_LICENSE_JSON = RUNTIME_DIRECTORY / 'run_helpers' / 'system' / '.client' / 'session.json'
DEFAULT_TO_COPY_JSON = DEFAULT_LICENSING_DIRECTORY / 'to copy.json'


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


def get_machine_id() -> str:
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
        value, _ = winreg.QueryValueEx(key, "MachineGuid")
    machine_id = str(value).strip()
    if not machine_id:
        raise SystemExit('Windows MachineGuid is empty.')
    return machine_id


def resolve_machine_id(machine_id: str | None) -> str:
    normalized = (machine_id or '').strip()
    if not normalized or normalized.lower() == 'auto':
        return get_machine_id()
    return normalized


def prompt_machine_id() -> str:
    detected = get_machine_id()
    raw_value = input(f"Machine ID [{detected}]: " ).strip()
    return raw_value or detected


def get_default_expiry_date() -> str:
    return (date.today() + timedelta(days=30)).isoformat()


def resolve_expiry_date(expiry_date: str | None) -> str:
    normalized = (expiry_date or '').strip()
    return normalized or get_default_expiry_date()


def prompt_expiry_date() -> str:
    detected = get_default_expiry_date()
    raw_value = input(f"Expiry date (YYYY-MM-DD) [{detected}]: " ).strip()
    return raw_value or detected


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


def write_json_file(output_path: Path, payload: object) -> None:
    ensure_parent(output_path)
    output_path.write_text(json.dumps(payload, indent=2), encoding='utf-8', newline='\n')


def write_customer_license_file(customer_license_path: Path, license_key: str, customer_name: str) -> None:
    payload = {
        'license_key': license_key.strip(),
        'customer_name': customer_name.strip(),
    }
    write_json_file(customer_license_path, payload)


def write_to_copy_file(output_path: Path, entry_payload: dict[str, str]) -> None:
    write_json_file(output_path, entry_payload)


def build_license_entry(
    license_key: str | None,
    customer_name: str,
    machine_id: str,
    expiry_date: str,
    allowed_version: str,
    status: str,
    existing_licenses: list[object] | None = None,
) -> tuple[str, dict[str, str]]:
    existing_licenses = existing_licenses or []
    normalized_key = (license_key or '').strip() or build_auto_license_key(existing_licenses, customer_name)
    for entry in existing_licenses:
        if isinstance(entry, dict) and str(entry.get('license_key', '')).strip() == normalized_key:
            raise SystemExit(f'License key already exists: {normalized_key}')
    payload = {
        'license_key': normalized_key,
        'customer_name': customer_name.strip(),
        'status': status.strip().lower(),
        'expiry_date': expiry_date.strip(),
        'machine_id': machine_id.strip(),
        'allowed_version': allowed_version.strip(),
    }
    return normalized_key, payload


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
    normalized_key, entry_payload = build_license_entry(
        license_key=license_key,
        customer_name=customer_name,
        machine_id=machine_id,
        expiry_date=expiry_date,
        allowed_version=allowed_version,
        status=status,
        existing_licenses=licenses,
    )

    licenses.append(entry_payload)
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
    print('4. Print current machine ID')
    print('5. Print license JSON entry and write the customer license file')
    choice = input('Choose 1, 2, 3, 4, or 5: ').strip()

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
        print(f'Current machine ID: {get_machine_id()}')
        return

    if choice == '5':
        json_path = prompt_path('licenses.json path', DEFAULT_LICENSES_JSON)
        to_copy_path = prompt_path('to copy.json path', DEFAULT_TO_COPY_JSON)
        customer_license_path = prompt_path('Customer license file path', DEFAULT_CUSTOMER_LICENSE_JSON)
        license_key = input('License key [auto]: ').strip()
        customer_name = input('Customer name: ').strip()
        machine_id = prompt_machine_id()
        expiry_date = prompt_expiry_date()
        allowed_version = input('Allowed app version [1.0.0]: ').strip() or '1.0.0'
        status = input('Status [active]: ').strip() or 'active'
        existing_licenses = load_licenses_payload(json_path).get('licenses', [])
        created_key, entry_payload = build_license_entry(
            license_key=license_key,
            customer_name=customer_name,
            machine_id=machine_id,
            expiry_date=expiry_date,
            allowed_version=allowed_version,
            status=status,
            existing_licenses=existing_licenses if isinstance(existing_licenses, list) else [],
        )
        write_to_copy_file(to_copy_path, entry_payload)
        write_customer_license_file(customer_license_path, created_key, customer_name)
        print(f'License key: {created_key}')
        print('Add this entry to licenses.json:')
        print(json.dumps(entry_payload, indent=2))
        print(f'To-copy file saved to: {to_copy_path}')
        print(f'Customer license file saved to: {customer_license_path}')
        return

    raise SystemExit('Invalid choice. Expected 1, 2, 3, 4, or 5.')


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

    machine_id_parser = subparsers.add_parser('print-machine-id')

    add_license_bundle_parser = subparsers.add_parser('print-license-entry')
    add_license_bundle_parser.add_argument('--json', default=str(DEFAULT_LICENSES_JSON))
    add_license_bundle_parser.add_argument('--to-copy-out', default=str(DEFAULT_TO_COPY_JSON))
    add_license_bundle_parser.add_argument('--customer-license-out', default=str(DEFAULT_CUSTOMER_LICENSE_JSON))
    add_license_bundle_parser.add_argument('--license-key')
    add_license_bundle_parser.add_argument('--customer-name', required=True)
    add_license_bundle_parser.add_argument('--machine-id', default='auto')
    add_license_bundle_parser.add_argument('--expiry-date')
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

    if args.command == 'print-machine-id':
        print(get_machine_id())
        return

    if args.command == 'print-license-entry':
        json_path = Path(args.json)
        existing_licenses = load_licenses_payload(json_path).get('licenses', [])
        created_key, entry_payload = build_license_entry(
            license_key=args.license_key,
            customer_name=args.customer_name,
            machine_id=resolve_machine_id(args.machine_id),
            expiry_date=resolve_expiry_date(args.expiry_date),
            allowed_version=args.allowed_version,
            status=args.status,
            existing_licenses=existing_licenses if isinstance(existing_licenses, list) else [],
        )
        write_to_copy_file(Path(args.to_copy_out), entry_payload)
        write_customer_license_file(Path(args.customer_license_out), created_key, args.customer_name)
        print(f'License key: {created_key}')
        print(json.dumps(entry_payload, indent=2))
        print(f'To-copy file saved to: {args.to_copy_out}')
        print(f'Customer license file saved to: {args.customer_license_out}')
        return

    raise SystemExit(f'Unknown command: {args.command}')


if __name__ == '__main__':
    main()
