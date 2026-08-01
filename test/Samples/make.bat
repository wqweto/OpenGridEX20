@echo off
setlocal
set "VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
pushd "%~dp0"

powershell -NoProfile -Command "Get-ChildItem *.vbp,*.frm,*.bas,*.cls | ForEach-Object { $t = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::Default); $f = $t -replace '(?<!\r)\n', ([string][char]13 + [char]10); if ($f -ne $t) { [IO.File]::WriteAllText($_.FullName, $f, [Text.Encoding]::Default); Write-Host ('Fixed ' + $_.Name) } }"

rem --- everything a run produces goes to output\, incl. the compiler log,
rem --- so the project folder holds sources only
if not exist output md output
"%VB6%" /make Samples.vbp /out "%~dp0output\build.log"
if errorlevel 1 (
    type output\build.log
    exit /b 1
)
if exist output\Samples.out.txt del output\Samples.out.txt
rem --- run under a watchdog: a wedged sample must not hang the loop
powershell -NoProfile -Command "$p = Start-Process '.\Samples.exe' -PassThru; if (-not $p.WaitForExit(120000)) { $p.Kill(); Write-Host 'TIMEOUT: killed after 120s'; exit 1 }; exit 0"
if errorlevel 1 exit /b 1
if not exist output\Samples.out.txt (
    echo FAILED: output\Samples.out.txt not produced
    exit /b 1
)
findstr /C:"RESULT: PASSED" output\Samples.out.txt >nul || (
    type output\Samples.out.txt
    exit /b 1
)
echo Sample smoke run PASSED
