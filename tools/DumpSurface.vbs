'=========================================================================
'
' Open GridEX 2000 Control
' Dumps the public typelib surface of a built OCX via TLI (tlbinf32.dll)
' for comparison against doc\GridEX20.idl. Run with 32-bit cscript:
'
'   %WINDIR%\SysWOW64\cscript.exe //NoLogo DumpSurface.vbs [path\to\file.ocx]
'
'=========================================================================
Option Explicit
Dim oTli, oTlb, oInfo, oMem, lIdx, lJdx, sFile

If WScript.Arguments.Count > 0 Then
    sFile = WScript.Arguments(0)
Else
    sFile = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\..\src\OpenGridEX20.ocx"
End If
Set oTli = CreateObject("TLI.TLIApplication")
Set oTlb = oTli.TypeLibInfoFromFile(sFile)
WScript.Echo "Library: " & oTlb.Name & " v" & oTlb.MajorVersion & "." & oTlb.MinorVersion & " GUID=" & oTlb.GUID
WScript.Echo ""
WScript.Echo "=== CoClasses (" & oTlb.CoClasses.Count & ") ==="
For lIdx = 1 To oTlb.CoClasses.Count
    WScript.Echo oTlb.CoClasses.Item(lIdx).Name
Next
WScript.Echo ""
WScript.Echo "=== Enums (" & oTlb.Constants.Count & ") ==="
For lIdx = 1 To oTlb.Constants.Count
    Set oInfo = oTlb.Constants.Item(lIdx)
    For lJdx = 1 To oInfo.Members.Count
        Set oMem = oInfo.Members.Item(lJdx)
        WScript.Echo oInfo.Name & "." & oMem.Name & " = " & oMem.Value
    Next
Next
WScript.Echo ""
WScript.Echo "=== Interfaces (event dispinterfaces have __ prefix) ==="
For lIdx = 1 To oTlb.Interfaces.Count
    Set oInfo = oTlb.Interfaces.Item(lIdx)
    WScript.Echo ""
    WScript.Echo oInfo.Name & " (" & oInfo.Members.Count & " members)"
    For lJdx = 1 To oInfo.Members.Count
        Set oMem = oInfo.Members.Item(lJdx)
        WScript.Echo "  " & oMem.Name & " invkind=" & oMem.InvokeKind & " dispid=" & oMem.MemberId
    Next
Next
