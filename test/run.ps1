<#
.SYNOPSIS
Runs one of the test executables on a desktop of its own, under a watchdog.

Every harness here shows real top-level windows -- a scenario window per
capture, a host form per ported sample -- which makes a run flicker across
whatever the machine is doing. A separate desktop (CreateDesktopW +
STARTUPINFO.lpDesktop) keeps them out of the way: they are still real windows
with real WM_PAINT, so VisualDiff's BitBlt captures exactly what it captured
before, but nothing reaches the interactive desktop.

The DPI pass is chosen by __COMPAT_LAYER, which CreateProcessW inherits from
this process -- Start-Process would not, which is what made both VisualDiff
passes run at system DPI until 2b24692.

.EXAMPLE
run.ps1 -Exe ..\VisualDiff\VisualDiff.exe -Arguments "verify 03*" -Layer DPIUNAWARE
run.ps1 -Exe ..\Samples\Samples.exe -TimeoutMs 120000
#>
param(
    [Parameter(Mandatory = $true)][string]$Exe,
    [string]$Arguments = '',
    [string]$Layer = '',
    [int]$TimeoutMs = 300000,
    [switch]$SameDesktop
)
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class NativeRun {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STARTUPINFO {
        public int cb;
        public string lpReserved, lpDesktop, lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2;
        public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION {
        public IntPtr hProcess, hThread;
        public int dwProcessId, dwThreadId;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateDesktopW(string desktop, string device, IntPtr devmode, int flags, uint access, IntPtr sa);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool CreateProcessW(string app, System.Text.StringBuilder cmd, IntPtr pa, IntPtr ta,
        bool inherit, uint flags, IntPtr env, string dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr h);
}
'@

$env:__COMPAT_LAYER = $Layer
$exe = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $Exe))
$cmd = New-Object Text.StringBuilder ('"{0}" {1}' -f $exe, $Arguments)

$hDesktop = [IntPtr]::Zero
$name = ''
if (-not $SameDesktop) {
    #--- GENERIC_ALL; a failure here is not fatal, the run just becomes visible
    #--- [NullString]::Value, not $null: PowerShell marshals $null to a string
    #--- parameter as an empty string, and CreateDesktop rejects that with 87
    $hDesktop = [NativeRun]::CreateDesktopW('OpenGEXVisualDiff', [NullString]::Value, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
    if ($hDesktop -eq [IntPtr]::Zero) {
        Write-Host ('WARNING: CreateDesktop failed ({0}), running on the current desktop' -f [Runtime.InteropServices.Marshal]::GetLastWin32Error())
    } else {
        $name = 'OpenGEXVisualDiff'
    }
}

$si = New-Object NativeRun+STARTUPINFO
$si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
$si.lpDesktop = $name
$pi = New-Object NativeRun+PROCESS_INFORMATION
if (-not [NativeRun]::CreateProcessW($exe, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IO.Path]::GetDirectoryName($exe), [ref] $si, [ref] $pi)) {
    throw ('CreateProcess failed: {0}' -f [Runtime.InteropServices.Marshal]::GetLastWin32Error())
}

$proc = [Diagnostics.Process]::GetProcessById($pi.dwProcessId)
$exit = 0
if (-not $proc.WaitForExit($TimeoutMs)) {
    $proc.Kill()
    Write-Host ('TIMEOUT: killed after {0}s' -f ([int]($TimeoutMs / 1000)))
    $exit = 1
}
[void][NativeRun]::CloseHandle($pi.hThread)
[void][NativeRun]::CloseHandle($pi.hProcess)
if ($hDesktop -ne [IntPtr]::Zero) {
    [void][NativeRun]::CloseDesktop($hDesktop)
}
exit $exit
