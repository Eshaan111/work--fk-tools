$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distRoot = Join-Path $projectRoot "dist"
$buildRoot = Join-Path $projectRoot "build"
$runtimeDistRoot = Join-Path $distRoot "Full-LC-AUTO"
$setupDistRoot = Join-Path $distRoot "Full-LC-AUTO-Setup"
$updateDistRoot = Join-Path $distRoot "Full-LC-AUTO-UpdateCenter"
$releaseBase = Join-Path $distRoot "release"
$releaseRoot = Join-Path $releaseBase "Full-LC-AUTO"
$venvPython = Join-Path $projectRoot "..\.venv\Scripts\python.exe"

if (-not (Test-Path $venvPython)) {
    throw "Virtual environment python not found at $venvPython"
}

$buildArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--distpath", $distRoot,
    "--workpath", $buildRoot,
    "--specpath", $projectRoot
)

& $venvPython -m PyInstaller @buildArgs --name "Full-LC-AUTO" "$projectRoot\main_ui_themed.py"
& $venvPython -m PyInstaller @buildArgs --name "Full-LC-AUTO-Setup" "$projectRoot\setup.py"
& $venvPython -m PyInstaller @buildArgs --name "Full-LC-AUTO-UpdateCenter" "$projectRoot\customer_update.py"

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null

Get-ChildItem -Path $runtimeDistRoot | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $releaseRoot -Recurse -Force
}

Copy-Item -Path (Join-Path $setupDistRoot "Full-LC-AUTO-Setup.exe") -Destination $releaseRoot -Force
Copy-Item -Path (Join-Path $updateDistRoot "Full-LC-AUTO-UpdateCenter.exe") -Destination $releaseRoot -Force

$copyTargets = @(
    "config.json",
    "ONBOARDING_README.md",
    "PACKAGING_AND_REMOTE_UPDATES.md",
    "image_folder_insight.py",
    "app_paths.py",
    "customer_update.py",
    "setup.py",
    "successful-run-record.xlsx",
    "assets",
    "data inputs",
    "json_LC_creation",
    "run_helpers"
)

foreach ($target in $copyTargets) {
    $source = Join-Path $projectRoot $target
    if (Test-Path $source) {
        Copy-Item -Path $source -Destination $releaseRoot -Recurse -Force
    }
}

Write-Host "Release folder ready at $releaseRoot"
