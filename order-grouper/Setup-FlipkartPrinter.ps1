[CmdletBinding()]
param(
    [string]$PhysicalPrinterName = "TSC TTP-244 Pro",
    [switch]$SkipApplicationInstall
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$CaptureQueueName = "PDFCreator"
$PreservedPdfQueueName = "PDFCreator Original"
$CapturePortName = "FlipkartRawCapturePort"
$CaptureHost = "127.0.0.1"
$CapturePortNumber = 9100

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Green
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run this setup as Administrator. Use 'Run Machine Setup.cmd' or right-click PowerShell and choose Run as administrator."
    }
}

function Install-WingetPackage {
    param(
        [Parameter(Mandatory)] [string]$Id,
        [Parameter(Mandatory)] [string]$DisplayName
    )

    $listed = (& winget list --id $Id --exact --source winget --accept-source-agreements 2>$null | Out-String)
    if ($listed -match [regex]::Escape($Id)) {
        Write-Host "$DisplayName is already installed." -ForegroundColor DarkGreen
        return
    }

    Write-Host "Installing $DisplayName..."
    & winget install --id $Id --exact --source winget --silent `
        --accept-package-agreements --accept-source-agreements `
        --disable-interactivity
    if ($LASTEXITCODE -ne 0) {
        throw "winget could not install $DisplayName (exit code $LASTEXITCODE)."
    }
}

function Get-PrinterByName {
    param([string]$Name)
    return Get-Printer -Name $Name -ErrorAction SilentlyContinue
}

function Assert-QzTrayInstalled {
    $knownExecutables = @(
        (Join-Path $env:ProgramFiles "QZ Tray\qz-tray.exe"),
        (Join-Path $env:ProgramFiles "QZ Tray\qz-tray-console.exe")
    )
    if ($knownExecutables | Where-Object { Test-Path -LiteralPath $_ }) {
        Write-Host "QZ Tray is installed." -ForegroundColor DarkGreen
        return
    }

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        $listed = (& winget list --id "QZIndustries.QZTray" --exact --source winget --accept-source-agreements 2>$null | Out-String)
        if ($listed -match "QZIndustries.QZTray") {
            Write-Host "QZ Tray is installed." -ForegroundColor DarkGreen
            return
        }
    }

    throw "QZ Tray was not detected. Install QZ Tray first, then rerun this setup."
}

Assert-Administrator

$logDirectory = Join-Path $PSScriptRoot "setup-logs"
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$logPath = Join-Path $logDirectory ("machine-setup-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
Start-Transcript -Path $logPath | Out-Null

try {
    Write-Host "Flipkart Batch Printer - machine setup" -ForegroundColor Green
    Write-Host "Python and Python packages are intentionally not installed by this script."

    Import-Module PrintManagement -ErrorAction Stop

    Write-Step "Verifying preinstalled QZ Tray"
    Assert-QzTrayInstalled

    if (-not $SkipApplicationInstall) {
        Write-Step "Checking Windows Package Manager"
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            throw "winget is unavailable. Install or update Microsoft App Installer, then run this setup again."
        }

        Write-Step "Installing PDFCreator Free"
        Install-WingetPackage -Id "Avanquestpdfforge.PDFCreator-Free" -DisplayName "PDFCreator Free"
        Start-Sleep -Seconds 3
    }

    Write-Step "Detecting the physical TSC printer"
    $physicalPrinter = Get-PrinterByName -Name $PhysicalPrinterName

    if (-not $physicalPrinter) {
        $candidates = @(
            Get-Printer | Where-Object {
                $_.Name -like "*TTP-244*Pro*" -and
                $_.PortName -ne $CapturePortName
            }
        )

        if ($candidates.Count -eq 1) {
            $candidate = $candidates[0]
            Write-Host "Creating the required physical queue name from detected queue '$($candidate.Name)'."
            Add-Printer -Name $PhysicalPrinterName `
                -DriverName $candidate.DriverName `
                -PortName $candidate.PortName `
                -Datatype "RAW"
            $physicalPrinter = Get-PrinterByName -Name $PhysicalPrinterName
        }
        elseif ($candidates.Count -gt 1) {
            $candidateNames = ($candidates.Name -join ", ")
            throw "Multiple TTP-244 Pro queues were found ($candidateNames). Rename the correct physical queue to '$PhysicalPrinterName' and rerun setup."
        }
    }

    if (-not $physicalPrinter) {
        throw @"
The physical '$PhysicalPrinterName' queue was not detected.
Install and connect the TTP-244 Pro first, confirm it appears in Windows
Printers & scanners, and then run this setup again.
"@
    }

    if ($physicalPrinter.PortName -eq $CapturePortName) {
        throw "The physical printer is incorrectly attached to the capture port. It must use its real USB port."
    }
    $physicalDriverName = [string]$physicalPrinter.DriverName
    Write-Host "Physical queue: $($physicalPrinter.Name)" -ForegroundColor DarkGreen
    Write-Host "Physical driver: $physicalDriverName"
    Write-Host "Physical port: $($physicalPrinter.PortName)"

    Write-Step "Preserving the real PDFCreator virtual printer"
    $namedPdfCreator = Get-PrinterByName -Name $CaptureQueueName
    $captureAlreadyConfigured = (
        $namedPdfCreator -and
        $namedPdfCreator.PortName -eq $CapturePortName -and
        $namedPdfCreator.DriverName -eq $physicalDriverName
    )

    if ($namedPdfCreator -and -not $captureAlreadyConfigured) {
        $preservedQueue = Get-PrinterByName -Name $PreservedPdfQueueName
        if ($preservedQueue) {
            throw "Both '$CaptureQueueName' and '$PreservedPdfQueueName' already exist, but '$CaptureQueueName' is not the expected capture queue. Resolve the printer-name conflict manually and rerun setup."
        }
        Rename-Printer -Name $CaptureQueueName -NewName $PreservedPdfQueueName
        Write-Host "Renamed the real PDFCreator queue to '$PreservedPdfQueueName'."
    }
    elseif (-not $namedPdfCreator -and -not (Get-PrinterByName -Name $PreservedPdfQueueName)) {
        Write-Warning "PDFCreator is installed, but its virtual printer queue was not found. The capture path can still work; repair PDFCreator later if normal PDF output is needed."
    }

    Write-Step "Creating the loopback RAW capture port"
    $capturePort = Get-PrinterPort -Name $CapturePortName -ErrorAction SilentlyContinue
    if ($capturePort) {
        $existingHost = [string]$capturePort.PrinterHostAddress
        $existingPort = [int]$capturePort.PortNumber
        if (($existingHost -and $existingHost -ne $CaptureHost) -or
            ($existingPort -and $existingPort -ne $CapturePortNumber)) {
            throw "Port '$CapturePortName' already exists but points to $existingHost`:$existingPort instead of $CaptureHost`:$CapturePortNumber. It was not changed."
        }
        Write-Host "Capture port already exists and is compatible."
    }
    else {
        Add-PrinterPort -Name $CapturePortName `
            -PrinterHostAddress $CaptureHost `
            -PortNumber $CapturePortNumber `
            -SNMP 0
        Write-Host "Created $CapturePortName -> $CaptureHost`:$CapturePortNumber."
    }

    Write-Step "Creating the Flipkart/QZ capture printer"
    $captureQueue = Get-PrinterByName -Name $CaptureQueueName
    if ($captureQueue) {
        if ($captureQueue.PortName -ne $CapturePortName -or
            $captureQueue.DriverName -ne $physicalDriverName) {
            throw "Printer '$CaptureQueueName' exists but is not the expected capture queue. It was not changed."
        }
        Write-Host "Capture queue is already configured."
    }
    else {
        Add-Printer -Name $CaptureQueueName `
            -DriverName $physicalDriverName `
            -PortName $CapturePortName `
            -Datatype "RAW"
        $captureQueue = Get-PrinterByName -Name $CaptureQueueName
    }

    if (-not $captureQueue) {
        throw "Windows did not create the '$CaptureQueueName' capture queue."
    }

    Write-Step "Final verification"
    $physicalPrinter = Get-PrinterByName -Name $PhysicalPrinterName
    $captureQueue = Get-PrinterByName -Name $CaptureQueueName
    if ($physicalPrinter.PortName -eq $CapturePortName) {
        throw "Safety check failed: physical printer points to the capture port."
    }
    if ($captureQueue.PortName -ne $CapturePortName) {
        throw "Safety check failed: capture queue points to the wrong port."
    }

    Write-Host ""
    Write-Host "Machine setup completed successfully." -ForegroundColor Green
    Write-Host "Physical printer : $($physicalPrinter.Name) -> $($physicalPrinter.PortName)"
    Write-Host "Capture printer  : $($captureQueue.Name) -> $CaptureHost`:$CapturePortNumber"
    Write-Host "PDF printer      : $PreservedPdfQueueName"
    Write-Host "Setup log        : $logPath"
    Write-Host ""
    Write-Host "Next: start the batch-printer application, confirm LISTENING, then print from Flipkart."
}
finally {
    Stop-Transcript | Out-Null
}
