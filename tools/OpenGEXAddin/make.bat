@echo off
setlocal
set "VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
pushd "%~dp0"

powershell -NoProfile -Command "Get-ChildItem *.vbp,*.cls,*.bas | ForEach-Object { $t = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::Default); $f = $t -replace '(?<!\r)\n', ([string][char]13 + [char]10); if ($f -ne $t) { [IO.File]::WriteAllText($_.FullName, $f, [Text.Encoding]::Default); Write-Host ('Fixed ' + $_.Name) } }"

"%VB6%" /make OpenGEXAddin.vbp /out errors.log
if errorlevel 1 (
    type errors.log
    exit /b 1
)

rem --- ensure Add-In Manager entry (manual load); export.ps1 flips it to 1 for batch runs
powershell -NoProfile -Command "foreach ($ini in \"$env:WINDIR\vbaddin.ini\", \"$env:LOCALAPPDATA\VirtualStore\Windows\vbaddin.ini\") { if (Test-Path $ini) { try { $t = @(Get-Content $ini); if (-not ($t -match '^OpenGEXAddin\.cConnect=')) { $i = [array]::IndexOf($t, '[Add-Ins32]'); if ($i -ge 0) { $t = $t[0..$i] + 'OpenGEXAddin.cConnect=0' + $t[($i+1)..($t.Count-1)]; Set-Content $ini $t; Write-Host ('Registered in ' + $ini) } } } catch { Write-Host ('Skipped ' + $ini + ': ' + $_.Exception.Message) } } }"
echo Successfully built OpenGEXAddin.dll
