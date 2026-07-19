Attribute VB_Name = "mdExport"
'=========================================================================
'
' Open GridEX 2000 Control
' Design-time snapshot export over the active VB6 IDE project
'
' Walks all form designers in the active project, snapshots every original
' (GridEX20.*) or open (OpenGridEX20.*) GridEX/GEXPreview control via the
' shared engine (mdSnapshot.bas) augmented with a "raw" section holding the
' flattened propbag keys parsed from the .frm text, and drops one JSON per
' control into OPENGEX_SNAPSHOT_DIR. Writes <sample>.done or <sample>.err
' marker consumed by the export.ps1 batch driver.
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Methods
'=========================================================================

Public Sub ExportActiveProject(oVBE As VBIDE.VBE)
    Const FUNC_NAME     As String = "ExportActiveProject"
    Dim sOutDir         As String
    Dim sSample         As String
    Dim oComp           As VBIDE.VBComponent
    Dim lCount          As Long
    Dim sErrors         As String
    Dim sSummary        As String

    On Error GoTo EH
    sOutDir = Environ$("OPENGEX_SNAPSHOT_DIR")
    If LenB(sOutDir) = 0 Then
        Exit Sub
    End If
    sSample = Environ$("OPENGEX_SAMPLE")
    If LenB(sSample) = 0 Then
        sSample = "Manual"
    End If
    If oVBE.ActiveVBProject Is Nothing Then
        Err.Raise vbObjectError, , "No active project in IDE"
    End If
    For Each oComp In oVBE.ActiveVBProject.VBComponents
        If oComp.Type = vbext_ct_VBForm Then
            lCount = lCount + pvExportForm(oComp, sOutDir, sSample, sErrors)
        End If
    Next
    sSummary = lCount & " control(s) exported from " & oVBE.ActiveVBProject.FileName
    If LenB(sErrors) <> 0 Then
        sSummary = sSummary & " (errors:" & sErrors & ")"
    End If
    pvSaveText sOutDir & "\" & sSample & ".done", sSummary
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    pvSaveText sOutDir & "\" & sSample & ".err", "Error &H" & Hex$(Err.Number) & " " & Err.Description & " [" & FUNC_NAME & "]"
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Function pvExportForm(oComp As VBIDE.VBComponent, sOutDir As String, sSample As String, sErrors As String) As Long
    Const FUNC_NAME     As String = "pvExportForm"
    Dim oForm           As VBIDE.VBForm
    Dim oCtl            As VBIDE.VBControl

    On Error GoTo EH
    Set oForm = oComp.Designer
    For Each oCtl In oForm.VBControls
        pvExportForm = pvExportForm + pvExportControl(oCtl, oComp, sOutDir, sSample, sErrors)
    Next
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    '--- damaged forms (e.g. missing .frx) cannot open a designer; salvage
    '--- the textual propbag keys from the .frm instead
    pvExportForm = pvExportFormRaw(oComp, sOutDir, sSample, sErrors)
End Function

Private Function pvExportFormRaw(oComp As VBIDE.VBComponent, sOutDir As String, sSample As String, sErrors As String) As Long
    Const FUNC_NAME     As String = "pvExportFormRaw"
    Dim aLines()        As String
    Dim lIdx            As Long
    Dim sLine           As String
    Dim aParts()        As String
    Dim sClass          As String
    Dim sCtlName        As String
    Dim oJson           As Object

    On Error GoTo EH
    aLines = Split(pvReadText(oComp.FileNames(1)), vbCrLf)
    For lIdx = 0 To UBound(aLines)
        sLine = Trim$(aLines(lIdx))
        If LCase$(sLine) Like "begin gridex20.* *" Or LCase$(sLine) Like "begin opengridex20.* *" Then
            aParts = Split(sLine, " ")
            sClass = Split(aParts(1), ".")(1)
            Select Case sClass
            Case "GridEX", "GEXPreview"
                sCtlName = aParts(2)
                Set oJson = Nothing
                JsonValue(oJson, "/$schema") = "opengex-snapshot/1"
                JsonValue(oJson, "class") = sClass
                JsonValue(oJson, "mode") = "raw-only"
                JsonValue(oJson, "sample") = sSample
                JsonValue(oJson, "form") = oComp.Name
                JsonValue(oJson, "control") = sCtlName
                JsonValue(oJson, "progid") = aParts(1)
                Set JsonValue(oJson, "raw") = pvParseRawKeys(oComp.FileNames(1), sCtlName)
                pvSaveText sOutDir & "\" & sSample & "_" & oComp.Name & "_" & sCtlName & ".json", JsonDump(oJson) & vbCrLf
                pvExportFormRaw = pvExportFormRaw + 1
                sErrors = sErrors & " " & oComp.Name & "!" & sCtlName & ": raw-only (designer failed)"
            End Select
        End If
    Next
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    sErrors = sErrors & " " & oComp.Name & ": &H" & Hex$(Err.Number) & " " & Err.Description
End Function

Private Function pvExportControl(oCtl As VBIDE.VBControl, oComp As VBIDE.VBComponent, sOutDir As String, sSample As String, sErrors As String) As Long
    Const FUNC_NAME     As String = "pvExportControl"
    Dim aParts()        As String
    Dim sClass          As String
    Dim sCtlName        As String
    Dim oJson           As Object
    Dim oRaw            As Object

    On Error GoTo EH
    aParts = Split(oCtl.ProgId, ".")
    If UBound(aParts) >= 1 Then
        Select Case LCase$(aParts(0))
        Case "gridex20", "opengridex20"
            sClass = aParts(1)
            Select Case sClass
            Case "GridEX", "GEXPreview"
                sCtlName = oCtl.Properties("Name")
                Set oJson = JsonParseObject(SnapshotToJson(oCtl.ControlObject, sClass, False))
                JsonValue(oJson, "sample") = sSample
                JsonValue(oJson, "form") = oComp.Name
                JsonValue(oJson, "control") = sCtlName
                JsonValue(oJson, "progid") = oCtl.ProgId
                Set oRaw = pvParseRawKeys(oComp.FileNames(1), sCtlName)
                If Not oRaw Is Nothing Then
                    Set JsonValue(oJson, "raw") = oRaw
                End If
                pvSaveText sOutDir & "\" & sSample & "_" & oComp.Name & "_" & sCtlName & ".json", JsonDump(oJson) & vbCrLf
                pvExportControl = 1
            End Select
        End Select
    End If
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    sErrors = sErrors & " " & oComp.Name & "!" & sCtlName & ": &H" & Hex$(Err.Number) & " " & Err.Description
End Function

Private Function pvParseRawKeys(sFrmFile As String, sCtlName As String) As Object
    Dim aLines()        As String
    Dim lIdx            As Long
    Dim sLine           As String
    Dim lDepth          As Long
    Dim aPath()         As String
    Dim lPos            As Long
    Dim sKey            As String
    Dim oJson           As Object
    Dim aParts()        As String

    aLines = Split(pvReadText(sFrmFile), vbCrLf)
    '--- locate the control block i.e. `Begin GridEX20.GridEX GridEX1`
    For lIdx = 0 To UBound(aLines)
        sLine = Trim$(aLines(lIdx))
        If sLine Like "Begin *.* " & sCtlName Then
            Exit For
        End If
    Next
    If lIdx > UBound(aLines) Then
        Exit Function
    End If
    '--- force-create json *object* upfront so it dumps as {} even if empty
    JsonValue(oJson, vbNullString) = Empty
    ReDim aPath(0 To 32) As String
    For lIdx = lIdx + 1 To UBound(aLines)
        sLine = Trim$(aLines(lIdx))
        Select Case True
        Case sLine Like "BeginProperty *"
            aParts = Split(Mid$(sLine, Len("BeginProperty ") + 1), " ")
            aPath(lDepth) = aParts(0)
            lDepth = lDepth + 1
        Case sLine Like "Begin *"
            aParts = Split(sLine, " ")
            aPath(lDepth) = aParts(UBound(aParts))
            lDepth = lDepth + 1
        Case sLine = "EndProperty", sLine = "End"
            If lDepth = 0 Then
                Exit For
            End If
            lDepth = lDepth - 1
        Case Else
            lPos = InStr(sLine, "=")
            If lPos > 1 Then
                sKey = RTrim$(Left$(sLine, lPos - 1))
                If lDepth > 0 Then
                    ReDim Preserve aPath(0 To lDepth - 1) As String
                    sKey = Join(aPath, ".") & "." & sKey
                    ReDim Preserve aPath(0 To 32) As String
                End If
                JsonValue(oJson, sKey) = LTrim$(Mid$(sLine, lPos + 1))
            End If
        End Select
    Next
    Set pvParseRawKeys = oJson
End Function

Private Function pvReadText(sFile As String) As String
    Dim lFile           As Long

    lFile = FreeFile
    Open sFile For Binary Access Read As #lFile
    pvReadText = Space$(LOF(lFile))
    Get #lFile, , pvReadText
    Close #lFile
End Function

Private Sub pvSaveText(sFile As String, sText As String)
    Dim lFile           As Long

    lFile = FreeFile
    Open sFile For Output As #lFile
    Print #lFile, sText;
    Close #lFile
End Sub
