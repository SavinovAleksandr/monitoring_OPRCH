#Requires -Version 5.1
<#
.SYNOPSIS
  Build a native Excel-for-Windows macro-enabled workbook from the checked VBA .bas files.
.DESCRIPTION
  Requires desktop Microsoft Excel and the explicitly enabled Excel Trust Center setting:
  "Trust access to the VBA project object model".
  Creates a new *_Windows.xlsm, never changes the source .xlsx or existing .xlsm.
  Applies bundled fixes and regenerates Windows-1251 modules automatically; Python is optional.
.EXAMPLE
  powershell -NoProfile -File .\make_windows_workbook.ps1
.EXAMPLE
  powershell -NoProfile -File .\make_windows_workbook.ps1 -SkipSmokeTest
#>
param(
    [switch] $SkipSmokeTest
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $repo 'Шаблон_мониторинг_ОПРЧ.xlsx'
$modules = @('Common', 'Engine', 'Station', 'Sheets', 'Charts', 'Monitor')

if (-not [Environment]::OSVersion.Platform.ToString().StartsWith('Win')) {
    throw 'This script requires Microsoft Windows and desktop Excel.'
}
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "Missing source workbook: $source"
}
foreach ($module in $modules) {
    $path = Join-Path $repo "OPRCH_${module}_win1251.bas"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing Windows-1251 module: $path."
    }
}

# Apply the same audited, exact-text patches as apply_windows_fixes.py.
# PowerShell can do this without Python installed on the Windows PC.
$manifestPath = Join-Path $repo 'windows_patch_manifest.json'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Missing patch manifest: $manifestPath"
}
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$patchedSources = @{}
$changed = New-Object 'System.Collections.Generic.List[string]'
foreach ($entry in $manifest.patches) {
    $name = [string]$entry.file
    $path = Join-Path $repo $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing source: $path" }
    $original = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $sourceText = $original.Replace("`r`n", "`n").Replace("`r", "`n")
    foreach ($replacement in $entry.replacements) {
        $old = [string]$replacement.old
        $new = [string]$replacement.new
        if ($sourceText.Contains($old)) {
            $sourceText = $sourceText.Replace($old, $new)
        } elseif (-not $sourceText.Contains($new)) {
            throw "Unexpected source revision in ${name}; refusing an unsafe patch."
        }
    }
    $patchedSources[$name] = $sourceText
    if ($sourceText -cne $original) { $changed.Add($name) }
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$cp1251 = [System.Text.Encoding]::GetEncoding(1251,
    [System.Text.EncoderFallback]::ExceptionFallback,
    [System.Text.DecoderFallback]::ExceptionFallback)
$binaryFiles = @{}
foreach ($module in $modules) {
    $name = "OPRCH_${module}.bas"
    if (-not $patchedSources.ContainsKey($name)) {
        $patchedSources[$name] = [System.IO.File]::ReadAllText((Join-Path $repo $name),
            [System.Text.Encoding]::UTF8).Replace("`r`n", "`n").Replace("`r", "`n")
    }
    $lines = $patchedSources[$name] -split "`n"
    $slim = ($lines | Where-Object { -not $_.TrimStart().StartsWith("'") }) -join "`r`n"
    $slim = $slim.TrimEnd() + "`r`n"
    $data = $cp1251.GetBytes($slim)
    if ($data.Length -ge 64000) { throw "VBA module too large: $name ($($data.Length) bytes)." }
    $winName = "OPRCH_${module}_win1251.bas"
    $binaryFiles[$winName] = $data
    $winPath = Join-Path $repo $winName
    if (-not (Test-Path -LiteralPath $winPath) -or
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($winPath)) -cne [Convert]::ToBase64String($data)) {
        $changed.Add($winName)
    }
}
$installText = $patchedSources['OPRCH_Install.bas']
$binaryFiles['OPRCH_Install_win1251.bas'] = $cp1251.GetBytes($installText.Replace("`n", "`r`n"))
$installerPath = Join-Path $repo 'OPRCH_Install_win1251.bas'
if (-not (Test-Path -LiteralPath $installerPath) -or
    [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($installerPath)) -cne [Convert]::ToBase64String($binaryFiles['OPRCH_Install_win1251.bas'])) {
    $changed.Add('OPRCH_Install_win1251.bas')
}

# A small number of modules is involved; backups are cheap and never overwrite existing files.
if ($changed.Count -gt 0) {
    $backupDir = Join-Path $repo ('.oprch_windows_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -ItemType Directory -Path $backupDir | Out-Null
    foreach ($name in $changed) {
        $filePath = Join-Path $repo $name
        if (Test-Path -LiteralPath $filePath) { Copy-Item -LiteralPath $filePath -Destination (Join-Path $backupDir $name) }
    }
    foreach ($name in $patchedSources.Keys) {
        if ($changed.Contains($name)) {
            [System.IO.File]::WriteAllText((Join-Path $repo $name), $patchedSources[$name], $utf8NoBom)
        }
    }
    foreach ($name in $binaryFiles.Keys) {
        if ($changed.Contains($name)) { [System.IO.File]::WriteAllBytes((Join-Path $repo $name), $binaryFiles[$name]) }
    }
    Write-Host "Sources repaired and backed up to $backupDir"
} else {
    Write-Host 'VBA sources already patched; no files modified.'
}

# Check continuation counts in every generated module, before opening Excel.
foreach ($module in $modules) {
    $vbaText = [System.IO.File]::ReadAllText((Join-Path $repo "OPRCH_${module}_win1251.bas"), $cp1251)
    $continuations = 0
    foreach ($line in ($vbaText -split "`r?`n")) {
        if ($line -match '\s_\s*$') { $continuations++ }
        else {
            if ($continuations -gt 24) {
                throw "VBA compile error: >24 line continuations in OPRCH_$module."
            }
            $continuations = 0
        }
    }
}
$monitorText = [System.IO.File]::ReadAllText((Join-Path $repo 'OPRCH_Monitor_win1251.bas'), $cp1251)
if (-not $monitorText.Contains('gOPRCHQuiet = previousQuiet') -or
    -not $monitorText.Contains('AppendLog "ERROR", stepName')) {
    throw 'The patched monitor VBA module is invalid.'
}

$out = Join-Path $repo 'Шаблон_мониторинг_ОПРЧ_Windows.xlsm'
if (Test-Path -LiteralPath $out) {
    $out = Join-Path $repo ("Шаблон_мониторинг_ОПРЧ_Windows_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsm')
}

$excel = $null
$wb = $null
try {
    Write-Host 'Starting desktop Microsoft Excel...'
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw 'Desktop Microsoft Excel COM Automation is not available. Excel for the web cannot run this installer.'
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $wb = $excel.Workbooks.Open($source)
    try {
        $project = $wb.VBProject
        $null = $project.VBComponents.Count
    } catch {
        throw ('Excel denied access to the VBA project. Enable File > Options > Trust Center > ' +
               'Trust Center Settings > Macro Settings > Trust access to the VBA project object model, ' +
               'then restart Excel. Do not enable all macros globally.')
    }

    foreach ($module in $modules) {
        $path = Join-Path $repo "OPRCH_${module}_win1251.bas"
        Write-Host "Importing OPRCH_$module ..."
        $imported = $project.VBComponents.Import($path)
        if ($imported.Name -ne "OPRCH_$module") {
            throw "Unexpected VBA component name: $($imported.Name)"
        }
    }

    # Workbook_Open is optional for running the calculation, but restores the buttons
    # on subsequent opens. Use the actual workbook CodeName rather than assuming it.
    $bookModule = $project.VBComponents.Item($wb.CodeName).CodeModule
    $existingCode = ''
    if ($bookModule.CountOfLines -gt 0) {
        $existingCode = $bookModule.Lines(1, $bookModule.CountOfLines)
    }
    if ($existingCode -notmatch '(?i)Private\s+Sub\s+Workbook_Open\s*\(') {
        $startupCode = "`r`nPrivate Sub Workbook_Open()`r`n" +
                       "    On Error Resume Next`r`n" +
                       "    InitOPRCHUISilent`r`n" +
                       "End Sub`r`n"
        $bookModule.AddFromString($startupCode)
    }

    # Save before building the buttons: OnAction must point to the final workbook name.
    $wb.SaveAs($out, 52)  # xlOpenXMLWorkbookMacroEnabled
    $macroBase = "'" + $wb.Name.Replace("'", "''") + "'!"
    $null = $excel.Run($macroBase + 'InitOPRCHUI')

    if (-not $SkipSmokeTest) {
        Write-Host 'Running AnalyzeOPRCH on the bundled demo data...'
        $null = $excel.Run($macroBase + 'AnalyzeOPRCH')
        $summary = $wb.Worksheets.Item('Summary')
        if ([string]$summary.Cells.Item(1, 1).Text -ne 'Станция') {
            throw 'Smoke test failed: the Summary header was not created.'
        }
        if ([string]::IsNullOrWhiteSpace([string]$summary.Cells.Item(2, 1).Text)) {
            throw 'Smoke test failed: Summary has no generator or station results.'
        }
        $log = $wb.Worksheets.Item('Log')
        $last = $log.Cells.Item($log.Rows.Count, 1).End(-4162).Row  # xlUp
        for ($row = 2; $row -le $last; $row++) {
            if ([string]$log.Cells.Item($row, 1).Text -eq 'ERROR') {
                throw ("Smoke test ERROR on Log row ${row}: " + [string]$log.Cells.Item($row, 3).Text)
            }
        }
        Write-Host 'Smoke test PASS: Summary and Log checked.'
    } else {
        Write-Host 'Smoke test skipped; the VBA code was not executed.'
    }

    $wb.Save()
    Write-Host "SUCCESS: $out"
} catch {
    Write-Error ("Windows Excel build/test failed: " + $_.Exception.Message)
    throw
} finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) | Out-Null } catch { }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
    }
    if ($null -ne $excel) {
        try { $excel.EnableEvents = $true; $excel.DisplayAlerts = $true; $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
