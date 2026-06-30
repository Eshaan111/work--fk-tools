# Packaging And Remote Updates

## Packaging shape

Use an `onedir` release, not a single-file exe.

That matters because this product depends on editable files beside the app:

- `config.json`
- `assets/`
- `data inputs/`
- `json_LC_creation/`
- Excel workbooks such as `successful-run-record.xlsx`

With `onedir`, the exe stays stable while customer-specific JSON and Excel files remain editable outside the exe.

## Release outputs

`build_exe.ps1` builds three Windows executables:

- `Full-LC-AUTO.exe`
- `Full-LC-AUTO-Setup.exe`
- `Full-LC-AUTO-UpdateCenter.exe`

It then prepares a distributable folder at:

`Full-LC-AUTO/dist/Full-LC-AUTO`

## Build prerequisites

Install runtime dependencies plus PyInstaller in your virtual environment:

```powershell
cd Full-LC-AUTO
..\.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller
```

If that venv path differs on your machine, point the script to the correct Python.

## Remote update model

The safest v1 is not direct remote control. It is bundle-based remote updates.

Flow:

1. You edit the customer copy of `config.json`, JSON flow files, or Excel inputs in your admin environment.
2. Open `Full-LC-AUTO-UpdateCenter.exe`.
3. Select the changed JSON/Excel files or whole folders.
4. Create a bundle zip.
5. Send the zip to the customer.
6. The customer opens `Full-LC-AUTO-UpdateCenter.exe` and applies the bundle.

When a bundle is applied:

- files are replaced only inside the app folder
- previous versions are backed up under `customer_updates/backups/`
- hashes are verified after write
- a text report is written under `customer_updates/reports/`

## What this supports

- updating `config.json`
- updating flow JSON files in `json_LC_creation/`
- updating asset JSON files in `assets/`
- updating Excel and CSV intake files in `data inputs/`

## What this does not do yet

- live cloud push
- automatic customer sync over internet
- license enforcement
- signed update bundles
- differential patches

Those can come in the next productization step once this release workflow feels solid.
