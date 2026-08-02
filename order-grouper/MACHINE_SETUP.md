# New computer setup

## Prerequisite

Install QZ Tray first. Also install and connect the physical TSC TTP-244 Pro,
confirm that Windows has a working physical printer queue, and print a Windows
test page. The setup script does not install QZ Tray or install/guess the
hardware driver or USB port.

## Run setup

Copy this project folder to the new computer, then double-click:

```text
Run Machine Setup.cmd
```

Accept the Windows administrator/UAC prompt. The elevated setup will:

1. Verify that QZ Tray is already installed without changing it.
2. Install PDFCreator Free using winget package
   `Avanquestpdfforge.PDFCreator-Free`.
3. Verify the physical `TSC TTP-244 Pro` queue and use its actual driver and
   physical port.
4. Preserve PDFCreator's real PDF printer as `PDFCreator Original`.
5. Create RAW TCP port `FlipkartRawCapturePort` at `127.0.0.1:9100` with SNMP
   disabled.
6. Create the Flipkart/QZ-visible capture queue named `PDFCreator`, using the
   detected TSC driver and the loopback capture port.
7. Verify that the physical printer was not pointed at the capture port.

Setup logs are written under `setup-logs`.

The script is designed to be rerun. It keeps compatible existing configuration
and stops instead of overwriting conflicting printer queues or ports.

## Intentionally excluded for now

The setup does not install Python, pip packages, a virtual environment, or a
Python runtime. Until the application is packaged as a distribution build, the
new computer still needs a compatible existing Python environment to run the
current `.pyw` application. The current development launcher contains a
machine-specific Python path and will be replaced for the distribution release.

To rerun only printer configuration without reinstalling/checking applications,
open an elevated PowerShell window in the project directory and use:

```powershell
.\Setup-FlipkartPrinter.ps1 -SkipApplicationInstall
```
