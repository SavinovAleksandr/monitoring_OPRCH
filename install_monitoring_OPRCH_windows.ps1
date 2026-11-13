#Requires -Version 5.1
<#
  OPRCH Windows bootstrapper, pinned to a reviewed GitHub revision.
  Downloads a repository snapshot (or accepts a locally downloaded archive),
  copies the bundled Windows fixes, then builds and smoke-tests a NEW workbook
  in local desktop Microsoft Excel. No Python or Git required.

  Examples:
    powershell -NoProfile -File .\install_monitoring_OPRCH_windows.ps1
    powershell -NoProfile -File .\install_monitoring_OPRCH_windows.ps1 -SourceZip C:\Downloads\monitoring_OPRCH-main.zip
    powershell -NoProfile -File .\install_monitoring_OPRCH_windows.ps1 -ExistingRepo C:\Work\monitoring_OPRCH
#>
param(
    [string] $SourceZip = '',
    [string] $ExistingRepo = '',
    [switch] $SkipSmokeTest
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($SourceZip -and $ExistingRepo) {
    throw 'Specify either -SourceZip or -ExistingRepo, not both.'
}
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'This installer requires Windows with desktop Microsoft Excel.'
}

$bundleDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoName = 'monitoring_OPRCH'
$revision = 'a3c663d099de9a6b51efc1c888ffc8e87679048b'
$workParent = Join-Path $bundleDir 'Результат_ОПРЧ_Windows'
if (Test-Path -LiteralPath $workParent) {
    $workParent += '_' + (Get-Date -Format 'yyyyMMdd_HHmmss')
}
New-Item -ItemType Directory -Path $workParent -Force | Out-Null

$repo = ''
if ($ExistingRepo) {
    $repo = (Resolve-Path -LiteralPath $ExistingRepo).Path
    if (-not (Test-Path -LiteralPath (Join-Path $repo 'OPRCH_Monitor.bas'))) {
        throw "Not an OPRCH repo: $repo"
    }
    # Never modify the user's existing clone: copy source files to an isolated working folder.
    $localRepo = Join-Path $workParent 'monitoring_OPRCH'
    New-Item -ItemType Directory -Path $localRepo | Out-Null
    Get-ChildItem -LiteralPath $repo -Force | Where-Object {
        $_.Name -ne '.git' -and $_.Name -notlike '.oprch_windows_backup_*'
    } | Copy-Item -Destination $localRepo -Recurse -Force
    $repo = $localRepo
} else {
    $zipPath = $SourceZip
    if (-not $zipPath) {
        $zipPath = Join-Path $workParent 'Исходный_репозиторий.zip'
        $url = "https://codeload.github.com/SavinovAleksandr/$repoName/zip/$revision"
        Write-Host "Downloading pinned repository revision $revision ..."
        try {
            # TLS 1.2 is needed on older Windows PowerShell 5.1 systems.
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        } catch {
            throw ('Unable to download GitHub repository: ' + $_.Exception.Message +
                '. You may manually download the repository ZIP and pass -SourceZip <path>.')
        }
    }
    $zipPath = (Resolve-Path -LiteralPath $zipPath).Path
    if ((Get-Item -LiteralPath $zipPath).Length -lt 50000) {
        throw "Repository archive is unusually small: $zipPath"
    }
    $expanded = Join-Path $workParent 'repo_archive'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $expanded -Force
    $files = @(Get-ChildItem -LiteralPath $expanded -Filter 'OPRCH_Monitor.bas' -Recurse -File)
    if ($files.Count -ne 1) {
        throw "Expected exactly one OPRCH_Monitor.bas in the repository ZIP; got $($files.Count)"
    }
    $repo = $files[0].Directory.FullName
}

$helpers = @(
    'apply_windows_fixes.py',
    'make_windows_workbook.ps1',
    'windows_patch_manifest.json',
    'README_WINDOWS.md'
)
foreach ($helper in $helpers) {
    $src = Join-Path $bundleDir $helper
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { throw "Missing bundled helper: $src" }
    Copy-Item -LiteralPath $src -Destination (Join-Path $repo $helper) -Force
}
foreach ($name in @(
    'OPRCH_Common.bas','OPRCH_Engine.bas','OPRCH_Station.bas',
    'OPRCH_Sheets.bas','OPRCH_Charts.bas','OPRCH_Monitor.bas',
    'OPRCH_Install.bas','Шаблон_мониторинг_ОПРЧ.xlsx'
)) {
    if (-not (Test-Path -LiteralPath (Join-Path $repo $name) -PathType Leaf)) {
        throw "Repository source is missing: $name"
    }
}

Write-Host "Preparing patched VBA sources in: $repo"
# Run the bundled native Windows Excel build. It also patches sources and rebuilds CP1251 files.
$builderPath = Join-Path $repo 'make_windows_workbook.ps1'
if ($SkipSmokeTest) {
    & $builderPath -SkipSmokeTest
} else {
    & $builderPath
}
if (-not $?) { throw 'make_windows_workbook.ps1 failed' }

$finished = @(Get-ChildItem -LiteralPath $repo -Filter 'Шаблон_мониторинг_ОПРЧ_Windows*.xlsm' -File)
if ($finished.Count -lt 1) { throw 'Windows .xlsm was not produced.' }

# The result folder already contains all six sources, CP1251 modules,
# Config, original template, installed workbook and documentation.
Write-Host ''
Write-Host 'SUCCESS. Finished workbook:'
$finished | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { Write-Host $_.FullName }
Write-Host 'All source files, configuration and installation instructions are in:'
Write-Host $repo
Write-Host 'If your company requires audit/verification, review the Log and Summary in Excel.'
