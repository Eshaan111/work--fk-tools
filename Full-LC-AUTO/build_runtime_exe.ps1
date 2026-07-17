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
    '--collect-submodules', 'selenium.webdriver',
    '--collect-submodules', 'urllib3',
    '--collect-submodules', 'websocket',
    '--collect-submodules', 'trio',
    '--collect-submodules', 'trio_websocket',
    '--hidden-import', 'selenium.webdriver.firefox.webdriver',
    '--hidden-import', 'selenium.webdriver.firefox.service',
    '--hidden-import', 'selenium.webdriver.firefox.options',
    (Join-Path $projectRoot 'main_tab_based.py')
)

$licenseToolBuildArgs = @(
    '--noconfirm',
    '--clean',
    '--onefile',
    '--console',
    '--name', 'license_tools',
    '--distpath', $distRoot,
    '--workpath', $buildRoot,
    '--specpath', $projectRoot,
    (Join-Path $projectRoot 'licensing\license_tools.py')
)

& $venvPython -m PyInstaller @buildArgs
& $venvPython -m PyInstaller @licenseToolBuildArgs

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null

Get-ChildItem -Path $runtimeDistRoot | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $releaseRoot -Recurse -Force
}

$licenseToolExe = Join-Path $distRoot 'license_tools.exe'
if (Test-Path $licenseToolExe) {
    Copy-Item -Path $licenseToolExe -Destination (Join-Path $releaseRoot 'license_tools.exe') -Force
}
else {
    throw "License tools exe was not produced at $licenseToolExe"
}

$copyTargets = @(
    'config.json',
    'licensing',
    'insights',
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



$productionConfigPath = Join-Path $releaseRoot 'config.json'
if (Test-Path $productionConfigPath) {
    $configJson = Get-Content -Path $productionConfigPath -Raw | ConvertFrom-Json
    if ($null -ne $configJson.shared -and $null -ne $configJson.shared.license) {
        $configJson.shared.license.allow_local_fallback = $false
        $configJson | ConvertTo-Json -Depth 100 | Set-Content -Path $productionConfigPath
    }
}

foreach ($target in @(
    'licenses.json',
    'licenses.sig',
    'customer_license.json',
    'image_folder_insight.xlsx',
    'licensing\licenses.json',
    'licensing\licenses.sig',
    'licensing\customer_license.json'
)) {
    $stalePath = Join-Path $releaseRoot $target
    if (Test-Path $stalePath) {
        Remove-Item -LiteralPath $stalePath -Force
    }
}

Write-Host "Runtime release ready at $releaseRoot"
Write-Host "License tools exe ready at $(Join-Path $releaseRoot 'license_tools.exe')"
