Attribute VB_Name = "mdUtils"
'=========================================================================
'
' Open GridEX 2000 Control
' Shared helper procedures for tools and tests
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long

'=========================================================================
' Functions
'=========================================================================

Public Function C2Obj(Value As Variant) As Object
    If IsObject(Value) Then
        Set C2Obj = Value
    End If
End Function

Public Function C2Dbl(Value As Variant) As Double
    Dim vDest           As Variant

    If VarType(Value) = vbDouble Then
        C2Dbl = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbDouble) = 0 Then
        C2Dbl = vDest
    End If
End Function

Public Function C2Lng(Value As Variant) As Long
    Dim vDest           As Variant

    If VarType(Value) = vbLong Then
        C2Lng = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbLong) = 0 Then
        C2Lng = vDest
    End If
End Function

'--- Variant assignment which keeps objects as references i.e. does not
'--- collapse them through their default property on plain let-assign
Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function ReadTextFile(sFile As String) As String
    Dim lFile           As Long

    lFile = FreeFile
    Open sFile For Binary Access Read As #lFile
    ReadTextFile = Space$(LOF(lFile))
    Get #lFile, , ReadTextFile
    Close #lFile
End Function

Public Sub WriteTextFile(sFile As String, sText As String)
    Dim lFile           As Long

    lFile = FreeFile
    Open sFile For Output As #lFile
    Print #lFile, sText;
    Close #lFile
End Sub

Public Function EnumFiles(sFolder As String, Optional sMask As String = "*") As Collection
    Dim sPath           As String
    Dim sFile           As String

    sPath = sFolder
    If LenB(sPath) <> 0 Then
        If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If
    End If
    Set EnumFiles = New Collection
    sFile = Dir$(sPath & sMask)
    Do While LenB(sFile) <> 0
        If sFile <> "." And sFile <> ".." Then
            EnumFiles.Add sPath & sFile
        End If
        sFile = Dir$
    Loop
End Function

Public Sub LogError(sMessage As String, Optional ByVal lLine As Long)
    Static bDisabled    As Boolean
    Dim lFile           As Long
    Dim sPath           As String

    '--- an unhandled error in an event pops a modal dialog which wedges an
    '--- automated run until the watchdog kills it, so every handler reports
    '--- here instead. Reporting can never take the run down itself: a
    '--- failure to write just disables the log for the rest of the session
    If lLine <> 0 Then
        sMessage = sMessage & " at line " & lLine
    End If
    Debug.Print sMessage
    If bDisabled Then
        Exit Sub
    End If
    On Error GoTo EH
    sPath = App.Path & "\output"
    If LenB(Dir$(sPath, vbDirectory)) = 0 Then
        MkDir sPath
    End If
    lFile = FreeFile
    Open sPath & "\errors.log" For Append As #lFile
    Print #lFile, Format$(Now, "yyyy-mm-dd hh:nn:ss ") & sMessage
    Close #lFile
    Exit Sub
EH:
    bDisabled = True
End Sub
