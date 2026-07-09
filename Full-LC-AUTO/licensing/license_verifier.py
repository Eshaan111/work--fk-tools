from __future__ import annotations

import base64
import json
import time
import urllib.error
import urllib.request
import winreg
from dataclasses import dataclass
from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding


@dataclass
class LicenseValidationResult:
    license_key: str
    customer_name: str
    machine_id: str
    expires_on: str
    source: str


class LicenseValidationError(RuntimeError):
    pass


class RemoteBundleDownloadError(LicenseValidationError):
    pass


def resolve_path(project_root: Path, path_value: str | None) -> Path:
    if not path_value:
        raise LicenseValidationError("Missing required license path configuration.")
    candidate = Path(path_value).expanduser()
    if candidate.is_absolute():
        return candidate
    return (project_root / candidate).resolve()


def fetch_url(url: str) -> tuple[bytes, dict[str, str]]:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Full-LC-AUTO-LicenseVerifier/1.0",
            "Accept": "*/*",
            "Cache-Control": "no-cache",
        },
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return response.read(), dict(response.headers.items())


def fetch_text(url: str) -> str:
    payload, _ = fetch_url(url)
    return payload.decode("utf-8")


def fetch_bytes(url: str) -> bytes:
    payload, _ = fetch_url(url)
    return payload


def load_public_key(public_key_path: Path):
    return serialization.load_pem_public_key(public_key_path.read_bytes())


def verify_signature(payload: bytes, signature_text: str, public_key) -> None:
    try:
        signature = base64.b64decode(signature_text.encode("ascii"), validate=True)
    except Exception as exc:
        raise LicenseValidationError(f"License signature is not valid base64: {exc}") from exc
    try:
        public_key.verify(signature, payload, padding.PKCS1v15(), hashes.SHA256())
    except InvalidSignature as exc:
        raise LicenseValidationError("License signature verification failed.") from exc


def get_machine_id() -> str:
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
        value, _ = winreg.QueryValueEx(key, "MachineGuid")
    machine_id = str(value).strip()
    if not machine_id:
        raise LicenseValidationError("Windows MachineGuid is empty.")
    return machine_id


def parse_iso_date(value: str) -> datetime:
    normalized = value.strip()
    if not normalized:
        raise LicenseValidationError("License expiry date is missing.")
    if len(normalized) == 10:
        normalized = normalized + "T23:59:59"
    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"
    return datetime.fromisoformat(normalized)


def parse_retry_after_seconds(retry_after_value: str | None) -> float | None:
    if not retry_after_value:
        return None
    normalized = retry_after_value.strip()
    if not normalized:
        return None
    if normalized.isdigit():
        return max(0.0, float(normalized))
    try:
        retry_at = parsedate_to_datetime(normalized)
    except Exception:
        return None
    now = datetime.now(retry_at.tzinfo) if retry_at.tzinfo else datetime.now()
    return max(0.0, (retry_at - now).total_seconds())


def get_cache_paths(project_root: Path, license_config: dict[str, object]) -> tuple[Path, Path]:
    cache_dir_value = str(license_config.get("cache_directory", "run_helpers/system/.license_cache"))
    cache_dir = resolve_path(project_root, cache_dir_value)
    return cache_dir / "licenses.json", cache_dir / "licenses.sig"


def write_cached_bundle(
    project_root: Path,
    license_config: dict[str, object],
    licenses_bytes: bytes,
    signature_text: str,
) -> None:
    cache_json_path, cache_sig_path = get_cache_paths(project_root, license_config)
    cache_json_path.parent.mkdir(parents=True, exist_ok=True)
    cache_json_path.write_bytes(licenses_bytes)
    cache_sig_path.write_text(signature_text.strip() + "\n", encoding="utf-8")


def load_cached_bundle(project_root: Path, license_config: dict[str, object]) -> tuple[bytes, str, str]:
    cache_json_path, cache_sig_path = get_cache_paths(project_root, license_config)
    try:
        return cache_json_path.read_bytes(), cache_sig_path.read_text(encoding="utf-8").strip(), "cached"
    except Exception as exc:
        raise LicenseValidationError(f"Could not load cached hosted license files: {exc}") from exc


def load_remote_bundle(remote_licenses_url: str, remote_signature_url: str) -> tuple[bytes, str, str]:
    last_error: Exception | None = None
    for attempt in range(2):
        try:
            licenses_bytes, _ = fetch_url(remote_licenses_url)
            signature_bytes, _ = fetch_url(remote_signature_url)
            return licenses_bytes, signature_bytes.decode("utf-8").strip(), "remote"
        except urllib.error.HTTPError as exc:
            last_error = exc
            if exc.code != 429 or attempt == 1:
                break
            retry_after_seconds = parse_retry_after_seconds(exc.headers.get("Retry-After"))
            sleep_seconds = retry_after_seconds if retry_after_seconds is not None else 2.0
            time.sleep(min(max(sleep_seconds, 0.0), 5.0))
        except Exception as exc:
            last_error = exc
            break
    raise RemoteBundleDownloadError(f"Could not download hosted license files: {last_error}")


def load_local_bundle(project_root: Path, license_config: dict[str, object]) -> tuple[bytes, str, str]:
    licenses_path = resolve_path(project_root, str(license_config.get("local_licenses_path", "licenses.json")))
    signature_path = resolve_path(project_root, str(license_config.get("local_signature_path", "licenses.sig")))
    try:
        return licenses_path.read_bytes(), signature_path.read_text(encoding="utf-8").strip(), "local"
    except Exception as exc:
        raise LicenseValidationError(f"Could not load local license files: {exc}") from exc


def resolve_license_record_from_bundle(
    licenses_bytes: bytes,
    signature_text: str,
    public_key,
    license_key: str,
) -> dict[str, object]:
    verify_signature(licenses_bytes, signature_text, public_key)
    try:
        licenses_payload = json.loads(licenses_bytes.decode("utf-8"))
    except Exception as exc:
        raise LicenseValidationError(f"licenses.json is not valid UTF-8 JSON: {exc}") from exc
    return find_license_record(licenses_payload, license_key)


def find_license_record(licenses_payload: dict[str, object], license_key: str) -> dict[str, object]:
    licenses = licenses_payload.get("licenses", [])
    if not isinstance(licenses, list):
        raise LicenseValidationError("Hosted licenses.json has an invalid structure.")
    for entry in licenses:
        if isinstance(entry, dict) and str(entry.get("license_key", "")).strip() == license_key:
            return entry
    raise LicenseValidationError("License key was not found in licenses.json.")


def validate_license(project_root: Path, app_config: dict[str, object]) -> LicenseValidationResult:
    shared = app_config.get("shared")
    if not isinstance(shared, dict):
        raise LicenseValidationError("App config is missing shared settings.")
    license_config = shared.get("license")
    machine_id = get_machine_id()
    if not isinstance(license_config, dict) or not bool(license_config.get("enabled", False)):
        return LicenseValidationResult("disabled", "License disabled", machine_id, "", "disabled")

    customer_license_path = resolve_path(project_root, str(license_config.get("customer_license_path", "")))
    public_key_path = resolve_path(project_root, str(license_config.get("public_key_path", "")))
    remote_licenses_url = str(license_config.get("remote_licenses_url", "")).strip()
    remote_signature_url = str(license_config.get("remote_signature_url", "")).strip()
    app_version = str(license_config.get("app_version", "")).strip()
    allow_local_fallback = bool(license_config.get("allow_local_fallback", False))

    if not remote_licenses_url or not remote_signature_url:
        raise LicenseValidationError("Remote license URLs are not configured.")

    try:
        customer_payload = json.loads(customer_license_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise LicenseValidationError(f"Could not read customer license file at {customer_license_path}: {exc}") from exc
    license_key = str(customer_payload.get("license_key", "")).strip()
    if not license_key:
        raise LicenseValidationError(f"Customer license file at {customer_license_path} does not contain a license_key.")

    public_key = load_public_key(public_key_path)
    try:
        licenses_bytes, signature_text, source = load_remote_bundle(remote_licenses_url, remote_signature_url)
        license_record = resolve_license_record_from_bundle(licenses_bytes, signature_text, public_key, license_key)
        write_cached_bundle(project_root, license_config, licenses_bytes, signature_text)
    except RemoteBundleDownloadError as exc:
        try:
            licenses_bytes, signature_text, source = load_cached_bundle(project_root, license_config)
            license_record = resolve_license_record_from_bundle(licenses_bytes, signature_text, public_key, license_key)
        except LicenseValidationError as cached_exc:
            if not allow_local_fallback:
                raise LicenseValidationError(f"{exc} | Cache unavailable: {cached_exc}") from exc
            licenses_bytes, signature_text, source = load_local_bundle(project_root, license_config)
            license_record = resolve_license_record_from_bundle(licenses_bytes, signature_text, public_key, license_key)
    except LicenseValidationError:
        raise

    status = str(license_record.get("status", "")).strip().lower()
    if status != "active":
        raise LicenseValidationError(f"License status is '{status or 'unknown'}'.")

    expected_machine_id = str(license_record.get("machine_id", "")).strip()
    if expected_machine_id and expected_machine_id != machine_id:
        raise LicenseValidationError("This license is not approved for the current machine.")

    expires_on = str(license_record.get("expiry_date", "")).strip()
    expires_at = parse_iso_date(expires_on)
    now = datetime.now(expires_at.tzinfo) if expires_at.tzinfo else datetime.now()
    if expires_at < now:
        raise LicenseValidationError("This license has expired.")

    allowed_version = str(license_record.get("allowed_version", "")).strip()
    if allowed_version and app_version and allowed_version != app_version:
        raise LicenseValidationError(
            f"This license allows app version {allowed_version}, but the current app version is {app_version}."
        )

    customer_name = str(license_record.get("customer_name", customer_payload.get("customer_name", "Licensed Customer"))).strip() or "Licensed Customer"
    return LicenseValidationResult(license_key, customer_name, machine_id, expires_on, source)
