# Batch-exports design-time GridEX snapshots from the original Janus samples.
#
# For every sample .vbp: enables the OpenGEXAddin in vbaddin.ini, launches the
# VB6 IDE with the project, waits for the add-in to drop <sample>.done (or
# .err) into test\snapshots, then kills the IDE. Requires the original
# GridEX20.ocx registered and the add-in built (make.bat).
#
# Usage: powershell -File export.ps1 [-SampleFilter "Unbound*"]
param(
    [string]$SampleFilter = "*",
    [int]$TimeoutSec = 90
)

$ErrorActionPreference = "Stop"
$sVb6 = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\VB98\VB6.EXE"
$sSamplesRoot = "${env:ProgramFiles(x86)}\Janus Systems Components\Janus GridEX 2000\Samples"
$sOutDir = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "test\snapshots"
$aExclude = @('DAO Samples', 'Using RDS', 'Html & ADC Sample', 'Using DataEnvironment', 'ButtonBar Sample', 'Access Sample', 'BACKUP')

function Set-AddinLoad([int]$lValue) {
    foreach ($sIni in "$env:WINDIR\vbaddin.ini", "$env:LOCALAPPDATA\VirtualStore\Windows\vbaddin.ini") {
        if (Test-Path $sIni) {
            try {
                $aText = @(Get-Content $sIni)
                if ($aText -match '^OpenGEXAddin\.cConnect=') {
                    $aText = $aText -replace '^OpenGEXAddin\.cConnect=\d+', "OpenGEXAddin.cConnect=$lValue"
                } else {
                    $lIdx = [array]::IndexOf($aText, '[Add-Ins32]')
                    if ($lIdx -lt 0) { $aText += '[Add-Ins32]'; $lIdx = $aText.Count - 1 }
                    $aText = $aText[0..$lIdx] + "OpenGEXAddin.cConnect=$lValue" + $(if ($lIdx + 1 -lt $aText.Count) { $aText[($lIdx + 1)..($aText.Count - 1)] })
                }
                Set-Content $sIni $aText
            } catch { Write-Warning "Cannot update $sIni : $_" }
        }
    }
}

if (-not (Test-Path $sVb6)) { throw "VB6.EXE not found: $sVb6" }
if (-not (Test-Path $sSamplesRoot)) { throw "Original samples not found: $sSamplesRoot" }
New-Item -ItemType Directory -Force $sOutDir | Out-Null

$aVbps = Get-ChildItem $sSamplesRoot -Recurse -Filter *.vbp | Where-Object {
    $sRel = $_.FullName.Substring($sSamplesRoot.Length)
    -not ($aExclude | Where-Object { $sRel -like "*$_*" })
} | Sort-Object FullName

$aFailed = @()
$lDone = 0
Set-AddinLoad 1
try {
    foreach ($oVbp in $aVbps) {
        $sSample = (Split-Path -Leaf (Split-Path -Parent $oVbp.FullName)) -replace '[^A-Za-z0-9]+', '-'
        if ($sSample -notlike $SampleFilter) { continue }
        Remove-Item "$sOutDir\$sSample.done", "$sOutDir\$sSample.err" -ErrorAction SilentlyContinue
        $env:OPENGEX_SNAPSHOT_DIR = $sOutDir
        $env:OPENGEX_SAMPLE = $sSample
        Write-Host "Exporting $sSample ..." -NoNewline
        $oProc = Start-Process $sVb6 -ArgumentList "`"$($oVbp.FullName)`"" -PassThru
        $dDeadline = (Get-Date).AddSeconds($TimeoutSec)
        while ((Get-Date) -lt $dDeadline -and -not (Test-Path "$sOutDir\$sSample.done") -and -not (Test-Path "$sOutDir\$sSample.err")) {
            Start-Sleep -Milliseconds 500
        }
        Stop-Process -Id $oProc.Id -Force -ErrorAction SilentlyContinue
        if (Test-Path "$sOutDir\$sSample.done") {
            Write-Host " OK: $(Get-Content "$sOutDir\$sSample.done" -TotalCount 1)"
            $lDone++
        } elseif (Test-Path "$sOutDir\$sSample.err") {
            Write-Host " FAILED: $(Get-Content "$sOutDir\$sSample.err" -TotalCount 1)"
            $aFailed += $sSample
        } else {
            Write-Host " TIMEOUT"
            $aFailed += $sSample
        }
        Remove-Item "$sOutDir\$sSample.done", "$sOutDir\$sSample.err" -ErrorAction SilentlyContinue
    }
} finally {
    Set-AddinLoad 0
    Remove-Item Env:\OPENGEX_SNAPSHOT_DIR, Env:\OPENGEX_SAMPLE -ErrorAction SilentlyContinue
}

Write-Host "Exported $lDone sample project(s), $($aFailed.Count) failed"
if ($aFailed.Count) {
    Write-Host ("Failed: " + ($aFailed -join ', '))
    exit 1
}
