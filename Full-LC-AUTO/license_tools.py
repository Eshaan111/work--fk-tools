from __future__ import annotations

import argparse
import base64
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa

DEFAULT_KEY_DIRECTORY = Path('C:/Full-LC-AUTO-License-Keys')
DEFAULT_PRIVATE_KEY = DEFAULT_KEY_DIRECTORY / 'license_private_key.pem'
DEFAULT_PUBLIC_KEY = DEFAULT_KEY_DIRECTORY / 'license_public_key.pem'


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def prompt_path(prompt_text: str, default: Path) -> Path:
    raw_value = input(f"{prompt_text} [{default}]: ").strip()
    return Path(raw_value) if raw_value else default


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
    choice = input('Choose 1, 2, or 3: ').strip()

    if choice == '1':
        private_out = prompt_path('Private key output path', DEFAULT_PRIVATE_KEY)
        public_out = prompt_path('Public key output path', DEFAULT_PUBLIC_KEY)
        generate_keypair(private_out, public_out)
        print(f'Private key saved to: {private_out}')
        print(f'Public key saved to: {public_out}')
        return

    if choice == '2':
        public_key_path = prompt_path('Public key path', DEFAULT_PUBLIC_KEY)
        json_path = prompt_path('Signed JSON path', Path('licenses.json'))
        signature_path = prompt_path('Signature path', Path('licenses.sig'))
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
        json_path = prompt_path('JSON path to sign', Path('licenses.json'))
        signature_path = prompt_path('Signature output path', Path('licenses.sig'))
        sign_json(private_key_path, json_path, signature_path)
        print(f'Signature saved to: {signature_path}')
        return

    raise SystemExit('Invalid choice. Expected 1, 2, or 3.')


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='License signing tools')
    subparsers = parser.add_subparsers(dest='command')

    generate_parser = subparsers.add_parser('generate-keypair')
    generate_parser.add_argument('--private-out', default=str(DEFAULT_PRIVATE_KEY))
    generate_parser.add_argument('--public-out', default=str(DEFAULT_PUBLIC_KEY))

    sign_parser = subparsers.add_parser('sign-json')
    sign_parser.add_argument('--private-key', default=str(DEFAULT_PRIVATE_KEY))
    sign_parser.add_argument('--json', required=True)
    sign_parser.add_argument('--sig', required=True)

    verify_parser = subparsers.add_parser('verify-json')
    verify_parser.add_argument('--public-key', default=str(DEFAULT_PUBLIC_KEY))
    verify_parser.add_argument('--json', required=True)
    verify_parser.add_argument('--sig', required=True)
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if not args.command:
        run_interactive()
        return

    if args.command == 'generate-keypair':
        generate_keypair(Path(args.private_out), Path(args.public_out))
        print(f'Private key saved to: {args.private_out}')
        print(f'Public key saved to: {args.public_out}')
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

    raise SystemExit(f'Unknown command: {args.command}')


if __name__ == '__main__':
    main()
