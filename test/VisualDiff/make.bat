@echo off
setlocal
set "VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
pushd "%~dp0"

powershell -NoProfile -Command "Get-ChildItem *.vbp,*.frm,*.bas | ForEach-Object { $t = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::Default); $f = $t -replace '(?<!\r)\n', ([string][char]13 + [char]10); if ($f -ne $t) { [IO.File]::WriteAllText($_.FullName, $f, [Text.Encoding]::Default); Write-Host ('Fixed ' + $_.Name) } }"

"%VB6%" /make VisualDiff.vbp /out errors.log
if errorlevel 1 (
    type errors.log
    exit /b 1
)
if exist VisualDiff.out.txt del VisualDiff.out.txt
start /wait VisualDiff.exe verify
if not exist VisualDiff.out.txt (
    echo FAILED: VisualDiff.out.txt not produced
    exit /b 1
)
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    type VisualDiff.out.txt
    exit /b 1
)
echo VisualDiff verify PASSED
