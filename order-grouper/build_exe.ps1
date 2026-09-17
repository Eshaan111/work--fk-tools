$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "FlipkartBatchPrinter" `
    --hidden-import win32print `
    "print_batch_gui.pyw"

Write-Host "Built: $projectRoot\dist\FlipkartBatchPrinter.exe"
