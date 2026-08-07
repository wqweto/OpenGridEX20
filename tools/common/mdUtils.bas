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

'--- the lParam shape the window messages take, low word first
Public Function MakeDWord(ByVal lLoWord As Long, ByVal lHiWord As Long) As Long
    MakeDWord = (lLoWord And &HFFFF&) Or (lHiWord * &H10000)
End Function

'--- the same lParam from a point in twips, which is the unit a test shares with
'--- the widths it sets: a 1500 twip column is clicked in the middle at 750
'--- whatever the screen is, where a pixel count only lands there at one scale
Public Function TwipsDWord(ByVal lX As Long, ByVal lY As Long) As Long
    TwipsDWord = MakeDWord(lX \ Screen.TwipsPerPixelX, lY \ Screen.TwipsPerPixelY)
End Function

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

'--- every file a test run produces lands in a single output folder next to
'--- the exe, so the project folder holds sources only and the whole lot can
'--- be wiped (or gitignored) as one
Public Function FileExists(sFile As String) As Boolean
    FileExists = (LenB(Dir$(sFile, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) <> 0)
End Function

Public Function OutputFile(sName As String) As String
    Dim sPath           As String

    sPath = App.Path & "\output"
    If LenB(Dir$(sPath, vbDirectory)) = 0 Then
        MkDir sPath
    End If
    OutputFile = sPath & "\" & sName
End Function

'--- the pair every handler goes through, PushError at the callsite and
'--- PopPrintError behind it. Splitting them is what keeps Erl usable: read
'--- inside the reporter it would be the reporter's own line, so it has to
'--- be captured in the frame that failed, which an argument does
Public Function PushError() As Variant
    PushError = Array(Err.Number, Err.Source, Err.Description, Erl)
End Function

Public Sub PopPrintError(LocalErr As Variant, sModule As String, sFunction As String)
    LogError "Critical error: " & LocalErr(2) & " &H" & Hex$(LocalErr(0)) & " [" & GetErrorSource(sModule, sFunction, LocalErr(3)) & "]"
End Sub

Public Function GetErrorSource(sModule As String, sFunction As String, Optional ByVal ErrLine As Long) As String
    GetErrorSource = sModule _
        & IIf(LenB(sFunction) <> 0, "." & sFunction, vbNullString) _
        & IIf(ErrLine = 0, vbNullString, "(" & ErrLine & ")")
End Function

Public Sub LogError(sMessage As String, Optional ByVal lLine As Long)
    Static bDisabled    As Boolean
    Dim lFile           As Long

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
    lFile = FreeFile
    Open OutputFile("errors.log") For Append As #lFile
    Print #lFile, Format$(Now, "yyyy-mm-dd hh:nn:ss ") & sMessage
    Close #lFile
    Exit Sub
EH:
    bDisabled = True
End Sub
