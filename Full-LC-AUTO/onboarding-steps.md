# Onboarding Steps

## 1. On The Client Machine

Go to the release folder beside the `.exe`.

Get the machine ID:

```powershell
C:\work--fk-tools\.venv\Scripts\python.exe C:\work--fk-tools\Full-LC-AUTO\licensing\license_tools.py print-machine-id
```

Or run the interactive tool:

```powershell
C:\work--fk-tools\.venv\Scripts\python.exe C:\work--fk-tools\Full-LC-AUTO\licensing\license_tools.py
```

Then choose:

- `4. Print current machine ID`

Next, generate the customer license file and the JSON block to add to the hosted license list:

```powershell
C:\work--fk-tools\.venv\Scripts\python.exe C:\work--fk-tools\Full-LC-AUTO\licensing\license_tools.py print-license-entry --customer-name "Customer Name"
```

This will:

- auto-detect the machine ID unless overridden
- auto-generate a license key unless provided
- default expiry to 30 days from today unless overridden
- write `licensing/customer_license.json`
- print the JSON entry to paste into `licensing/licenses.json`

Copy the printed JSON block.

## 2. On Your Admin Machine

Open:

- `Full-LC-AUTO/licensing/licenses.json`

Paste the new license entry into the `licenses` array.

Sign the updated file:

```powershell
.\.venv\Scripts\python.exe Full-LC-AUTO\licensing\license_tools.py sign-json --private-key "C:\Full-LC-AUTO-License-Keys\license_private_key.pem" --json Full-LC-AUTO\licensing\licenses.json --sig Full-LC-AUTO\licensing\licenses.sig
```

Verify locally:

```powershell
.\.venv\Scripts\python.exe Full-LC-AUTO\licensing\license_tools.py verify-json --public-key Full-LC-AUTO\licensing\license_public_key.pem --json Full-LC-AUTO\licensing\licenses.json --sig Full-LC-AUTO\licensing\licenses.sig
```

Expected result:

- `Signature is valid.`

Commit and push:

```powershell
git add Full-LC-AUTO/licensing/licenses.json Full-LC-AUTO/licensing/licenses.sig
git commit -m "Add client license"
git push origin license-based-lc
```

## 3. Check GitHub Hosted License Bundle

Run:

```powershell
$base = "https://raw.githubusercontent.com/Eshaan111/work--fk-tools/license-based-lc/Full-LC-AUTO/licensing"
$temp = "$env:TEMP\full-lc-license-check"
New-Item -ItemType Directory -Force -Path $temp | Out-Null
$stamp = [guid]::NewGuid().ToString()

Invoke-WebRequest "$base/licenses.json?ts=$stamp" -OutFile "$temp\licenses.json"
Invoke-WebRequest "$base/licenses.sig?ts=$stamp" -OutFile "$temp\licenses.sig"
Invoke-WebRequest "$base/license_public_key.pem?ts=$stamp" -OutFile "$temp\license_public_key.pem"

.\.venv\Scripts\python.exe Full-LC-AUTO\licensing\license_tools.py verify-json --public-key "$temp\license_public_key.pem" --json "$temp\licenses.json" --sig "$temp\licenses.sig"
```

Expected result:

- `Signature is valid.`

## 4. Validate On The Client Machine

From the release folder beside the `.exe`:

```powershell
$config = Get-Content ".\config.json" -Raw | ConvertFrom-Json
$baseJson = "$($config.shared.license.remote_licenses_url)".Trim()
$baseSig  = "$($config.shared.license.remote_signature_url)".Trim()
$temp = "$env:TEMP\full-lc-license-check-client"
New-Item -ItemType Directory -Force -Path $temp | Out-Null
$stamp = [guid]::NewGuid().ToString()

Invoke-WebRequest ($baseJson + "?ts=" + $stamp) -OutFile "$temp\licenses.json"
Invoke-WebRequest ($baseSig + "?ts=" + $stamp) -OutFile "$temp\licenses.sig"

C:\work--fk-tools\.venv\Scripts\python.exe C:\work--fk-tools\Full-LC-AUTO\licensing\license_tools.py verify-json --public-key ".\licensing\license_public_key.pem" --json "$temp\licenses.json" --sig "$temp\licenses.sig"
```

Expected result:

- `Signature is valid.`

If this passes, the hosted signature/public-key path is correct for that release folder.

## 5. Launch The App

From the client release folder:

```powershell
.\Full-LC-AUTO.exe
```

The app should start if:

- the client `licensing/customer_license.json` matches a hosted license row
- the hosted license row is `active`
- the machine ID matches
- the license is not expired
- the app version matches

## Notes

- Do not commit the private key.
- Do commit the public key.
- Always sign after editing `licensing/licenses.json`.
- Production mode depends on the hosted GitHub files being reachable.
- The release folder should keep:
  - `licensing/customer_license.json`
  - `licensing/license_public_key.pem`
- The release folder should not rely on local `licenses.json` or `licenses.sig` in production.
