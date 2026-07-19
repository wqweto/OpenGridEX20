Attribute VB_Name = "mdMain"
'=========================================================================
'
' Open GridEX 2000 Control
' VisualDiff runner: record goldens from the original control, self-test
' the harness (original vs recorded golden) or verify our control against
' the golden corpus. Usage: VisualDiff.exe [record|selftest|verify] [mask]
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_PROGID_ORIGINAL           As String = "GridEX20.GridEX"
Private Const STR_PROGID_OURS               As String = "OpenGridEX20.GridEX"

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim aArgs()         As String
    Dim sMode           As String
    Dim sMask           As String
    Dim sProgId         As String
    Dim vFile           As Variant
    Dim sName           As String
    Dim vDoc            As Variant
    Dim oForm           As frmHost
    Dim baBits()        As Byte
    Dim baGolden()      As Byte
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lGoldenW        As Long
    Dim lGoldenH        As Long
    Dim lDiff           As Long
    Dim sReport         As String
    Dim sGolden         As String

    On Error GoTo EH
    aArgs = Split(Command$)
    If UBound(aArgs) >= 0 Then
        sMode = LCase$(aArgs(0))
    End If
    If LenB(sMode) = 0 Then
        sMode = "verify"
    End If
    If UBound(aArgs) >= 1 Then
        sMask = aArgs(1)
    Else
        sMask = "*"
    End If
    TestInit App.Path & "\VisualDiff.out.txt"
    If sMode = "verify" Then
        sProgId = STR_PROGID_OURS
    Else
        sProgId = STR_PROGID_ORIGINAL
    End If
    If Not pvAddLicense(sProgId) Then
        TestSkip "no design license for " & sProgId
        Exit Sub
    End If
    For Each vFile In EnumFiles(App.Path & "\scenarios", "*.json")
        sName = Mid$(vFile, InStrRev(vFile, "\") + 1)
        sName = Left$(sName, Len(sName) - Len(".json"))
        If sName Like sMask Then
            JsonParse ReadTextFile(CStr(vFile)), vDoc
            Set oForm = New frmHost
            If Not oForm.RunScenario(sProgId, C2Obj(vDoc), baBits, lWidth, lHeight) Then
                Assert sMode & " " & sName & " (capture failed)", False
            Else
                sGolden = App.Path & "\golden\" & sName & ".bmp"
                Select Case sMode
                Case "record"
                    SaveBmp sGolden, lWidth, lHeight, baBits
                    Assert "record " & sName, True
                Case "selftest", "verify"
                    If Not LoadBmp(sGolden, lGoldenW, lGoldenH, baGolden) Then
                        Assert sMode & " " & sName & " (no golden)", False
                    ElseIf lGoldenW <> lWidth Or lGoldenH <> lHeight Then
                        Assert sMode & " " & sName & " (golden size mismatch)", False
                    Else
                        lDiff = DiffBits(baGolden, baBits, lWidth, lHeight, sReport)
                        If lDiff > 0 Then
                            If LenB(Dir$(App.Path & "\output", vbDirectory)) = 0 Then
                                MkDir App.Path & "\output"
                            End If
                            SaveBmp App.Path & "\output\" & sName & ".actual.bmp", lWidth, lHeight, baBits
                        End If
                        Assert sMode & " " & sName & IIf(lDiff > 0, " diff=" & lDiff & "px " & sReport, vbNullString), lDiff = 0
                    End If
                End Select
            End If
            Unload oForm
            Set oForm = Nothing
        End If
    Next
    TestsDone
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    Assert "Unhandled error &H" & Hex$(Err.Number) & " " & Err.Description & " in " & sName, False
    TestsDone
End Sub

Private Function pvAddLicense(sProgId As String) As Boolean
    Const FUNC_NAME     As String = "pvAddLicense"

    On Error GoTo EH
    Licenses.Add sProgId
    pvAddLicense = True
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function
