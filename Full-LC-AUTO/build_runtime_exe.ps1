$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distRoot = Join-Path $projectRoot 'dist'
$buildRoot = Join-Path $projectRoot 'build'
$runtimeDistRoot = Join-Path $distRoot 'Full-LC-AUTO'
$releaseBase = Join-Path $distRoot 'release'
$releaseRoot = Join-Path $releaseBase 'Full-LC-AUTO'
$venvPython = Join-Path $projectRoot '..\.venv\Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    throw "Virtual environment python not found at $venvPython"
}

$buildArgs = @(
    '--noconfirm',
    '--clean',
    '--windowed',
    '--name', 'Full-LC-AUTO',
    '--distpath', $distRoot,
    '--workpath', $buildRoot,
    '--specpath', $projectRoot,
    (Join-Path $projectRoot 'main_ui_themed.py')
)

& $venvPython -m PyInstaller @buildArgs

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null

Get-ChildItem -Path $runtimeDistRoot | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $releaseRoot -Recurse -Force
}

$copyTargets = @(
    'config.json',
    'customer_license.json',
    'license_public_key.pem',
    'licenses.json',
    'licenses.sig',
    'image_folder_insight.xlsx',
    'successful-run-record.xlsx',
    'assets',
    'data inputs',
    'json_LC_creation',
    'run_helpers',
    'snapshots'
)

foreach ($target in $copyTargets) {
    $source = Join-Path $projectRoot $target
    if (Test-Path $source) {
        Copy-Item -Path $source -Destination $releaseRoot -Recurse -Force
    }
}

Write-Host "Runtime release ready at $releaseRoot"
