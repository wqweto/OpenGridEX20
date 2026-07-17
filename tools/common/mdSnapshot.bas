Attribute VB_Name = "mdSnapshot"
'=========================================================================
'
' Open GridEX 2000 Control
' Profile-driven object model snapshot engine (see tools\common\SCHEMA.md)
'
' Walks a GridEX/GEXPreview object model late-bound via CallByName using the
' generated profiles in mdProfiles.bas and produces a JSON snapshot document
' (mdJson.bas collections). Works against both the original Janus control
' and Open GridEX instances.
'
'=========================================================================
Option Explicit

Public Function SnapshotToJson(oObj As Object, sClass As String, Optional ByVal bRuntime As Boolean) As String
    Dim oJson           As Object

    JsonValue(oJson, "/$schema") = "opengex-snapshot/1"
    JsonValue(oJson, "class") = sClass
    JsonValue(oJson, "mode") = IIf(bRuntime, "runtime", "persist")
    Set JsonValue(oJson, "props") = SnapshotObject(oObj, sClass, vbNullString, bRuntime)
    SnapshotToJson = JsonDump(oJson) & vbCrLf
End Function

Public Function SnapshotObject(oObj As Object, sClass As String, sItemClass As String, ByVal bRuntime As Boolean) As Object
    Dim uProps()        As UcsProfileProp
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim oJson           As Object
    Dim oErrors         As Object
    Dim oItems          As Object

    lCount = ProfileForClass(sClass, uProps)
    For lIdx = 0 To lCount - 1
        If bRuntime Or Not uProps(lIdx).bRuntimeOnly Then
            pvSnapshotProp oObj, uProps(lIdx), oJson, oErrors, bRuntime
        End If
    Next
    '--- parameterized specials
    If sClass = "JSPrinterProperties" Then
        pvSnapshotPositional oObj, "HeaderString", 1, 3, oJson, oErrors
        pvSnapshotPositional oObj, "FooterString", 1, 3, oJson, oErrors
    End If
    '--- collection items
    If LenB(sItemClass) <> 0 Then
        Set oItems = pvSnapshotItems(oObj, sItemClass, oErrors, vbNullString, bRuntime)
        If Not oItems Is Nothing Then
            Set JsonValue(oJson, "items") = oItems
        End If
    End If
    '--- accumulated read errors (oErrors created on first recorded error only)
    If Not oErrors Is Nothing Then
        Set JsonValue(oJson, "/$errors") = oErrors
    End If
    Set SnapshotObject = oJson
End Function

Private Sub pvSnapshotProp(oObj As Object, uProp As UcsProfileProp, oJson As Object, oErrors As Object, ByVal bRuntime As Boolean)
    Dim vValue          As Variant
    Dim oSub            As Object
    Dim oItems          As Object

    On Error Resume Next
    Select Case uProp.eKind
    Case ucsPrkScalar, ucsPrkEnum, ucsPrkVariant
        vValue = CallByName(oObj, uProp.sProp, VbGet)
        If Err.Number <> 0 Then
            JsonValue(oErrors, uProp.sProp) = Err.Number
        ElseIf IsObject(vValue) Or IsEmpty(vValue) Then
            JsonValue(oJson, uProp.sProp) = Null
        Else
            JsonValue(oJson, uProp.sProp) = vValue
        End If
    Case ucsPrkFont
        Set oSub = CallByName(oObj, uProp.sProp, VbGet)
        If Err.Number <> 0 Then
            JsonValue(oErrors, uProp.sProp) = Err.Number
        ElseIf oSub Is Nothing Then
            JsonValue(oJson, uProp.sProp) = Null
        Else
            On Error GoTo 0
            Set JsonValue(oJson, uProp.sProp) = pvFontToJson(oSub, uProp.sProp, oErrors)
        End If
    Case ucsPrkPicture
        Set oSub = CallByName(oObj, uProp.sProp, VbGet)
        If Err.Number <> 0 Then
            JsonValue(oErrors, uProp.sProp) = Err.Number
        Else
            On Error GoTo 0
            pvWritePicture oSub, uProp.sProp, oJson, oErrors
        End If
    Case ucsPrkClass, ucsPrkCollection
        Set oSub = CallByName(oObj, uProp.sProp, VbGet)
        If Err.Number <> 0 Then
            JsonValue(oErrors, uProp.sProp) = Err.Number
        ElseIf oSub Is Nothing Then
            JsonValue(oJson, uProp.sProp) = Null
        ElseIf uProp.eKind = ucsPrkClass Or pvCollHasOwnProps(uProp.sTypeName) Then
            On Error GoTo 0
            Set JsonValue(oJson, uProp.sProp) = SnapshotObject(oSub, uProp.sTypeName, uProp.sItemClass, bRuntime)
        Else
            '--- collections w/o own persistable props dump as plain arrays
            On Error GoTo 0
            Set oItems = pvSnapshotItems(oSub, uProp.sItemClass, oErrors, uProp.sProp, bRuntime)
            If Not oItems Is Nothing Then
                Set JsonValue(oJson, uProp.sProp) = oItems
            End If
        End If
    End Select
    On Error GoTo 0
End Sub

Private Function pvSnapshotItems(oColl As Object, sItemClass As String, oErrors As Object, sErrPrefix As String, ByVal bRuntime As Boolean) As Object
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim oItem           As Object
    Dim vValue          As Variant
    Dim oItems          As Object

    On Error Resume Next
    lCount = CallByName(oColl, "Count", VbGet)
    If Err.Number <> 0 Then
        If LenB(sErrPrefix) = 0 Then
            JsonValue(oErrors, "Count") = Err.Number
        Else
            JsonValue(oErrors, sErrPrefix) = Err.Number
        End If
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    For lIdx = 1 To lCount
        On Error Resume Next
        If sItemClass = "String" Then
            vValue = CallByName(oColl, "Item", VbGet, lIdx)
            If Err.Number <> 0 Then
                JsonValue(oErrors, sErrPrefix & "[" & lIdx & "]") = Err.Number
                On Error GoTo 0
            Else
                On Error GoTo 0
                JsonValue(oItems, "-1") = vValue
            End If
        Else
            Set oItem = CallByName(oColl, "Item", VbGet, lIdx)
            If Err.Number <> 0 Then
                JsonValue(oErrors, sErrPrefix & "[" & lIdx & "]") = Err.Number
                On Error GoTo 0
            ElseIf oItem Is Nothing Then
                On Error GoTo 0
                JsonValue(oItems, "-1") = Null
            Else
                On Error GoTo 0
                Set JsonValue(oItems, "-1") = SnapshotObject(oItem, sItemClass, vbNullString, bRuntime)
            End If
        End If
    Next
    Set pvSnapshotItems = oItems
End Function

Private Function pvCollHasOwnProps(sClass As String) As Boolean
    Dim uProps()        As UcsProfileProp
    Dim lCount          As Long
    Dim lIdx            As Long

    lCount = ProfileForClass(sClass, uProps)
    For lIdx = 0 To lCount - 1
        If Not uProps(lIdx).bRuntimeOnly Then
            pvCollHasOwnProps = True
            Exit For
        End If
    Next
End Function

Private Sub pvSnapshotPositional(oObj As Object, sProp As String, ByVal lFrom As Long, ByVal lTo As Long, oJson As Object, oErrors As Object)
    Dim lPos            As Long
    Dim vValue          As Variant
    Dim oSub            As Object

    '--- force-create json *object* upfront or the digit keys below would
    '--- auto-create an index-shifted json *array* instead
    JsonValue(oSub, vbNullString) = Empty
    For lPos = lFrom To lTo
        On Error Resume Next
        vValue = CallByName(oObj, sProp, VbGet, lPos)
        If Err.Number <> 0 Then
            JsonValue(oErrors, sProp & "[" & lPos & "]") = Err.Number
        Else
            JsonValue(oSub, lPos) = vValue
        End If
        On Error GoTo 0
    Next
    Set JsonValue(oJson, sProp) = oSub
End Sub

Private Function pvFontToJson(oFont As Object, sKey As String, oErrors As Object) As Object
    Dim oJson           As Object

    On Error Resume Next
    JsonValue(oJson, "Name") = oFont.Name
    JsonValue(oJson, "Size") = CDbl(oFont.Size)
    JsonValue(oJson, "Bold") = oFont.Bold
    JsonValue(oJson, "Italic") = oFont.Italic
    JsonValue(oJson, "Underline") = oFont.Underline
    JsonValue(oJson, "Strikethrough") = oFont.Strikethrough
    JsonValue(oJson, "Charset") = CLng(oFont.Charset)
    If Err.Number <> 0 Then
        JsonValue(oErrors, sKey) = Err.Number
    End If
    On Error GoTo 0
    Set pvFontToJson = oJson
End Function

Private Sub pvWritePicture(oPic As Object, sKey As String, oJson As Object, oErrors As Object)
    Dim sTempFile       As String
    Dim baData()        As Byte
    Dim lFile           As Long
    Dim oSub            As Object

    On Error Resume Next
    If oPic Is Nothing Then
        JsonValue(oJson, sKey) = Null
        Exit Sub
    End If
    If oPic.Handle = 0 Then
        JsonValue(oJson, sKey) = Null
        Exit Sub
    End If
    sTempFile = Environ$("TEMP") & "\opengex_pic.tmp"
    SavePicture oPic, sTempFile
    If Err.Number <> 0 Then
        JsonValue(oErrors, sKey) = Err.Number
        On Error GoTo 0
        Exit Sub
    End If
    lFile = FreeFile
    Open sTempFile For Binary Access Read As #lFile
    ReDim baData(0 To LOF(lFile) - 1) As Byte
    Get #lFile, , baData
    Close #lFile
    Kill sTempFile
    JsonValue(oSub, "type") = CLng(oPic.Type)
    JsonValue(oSub, "data") = pvToBase64(baData)
    Set JsonValue(oJson, sKey) = oSub
    If Err.Number <> 0 Then
        JsonValue(oErrors, sKey) = Err.Number
    End If
    On Error GoTo 0
End Sub

Private Function pvToBase64(baData() As Byte) As String
    Static oNode        As Object

    If oNode Is Nothing Then
        Set oNode = CreateObject("MSXML2.DOMDocument").createElement("b64")
        oNode.DataType = "bin.base64"
    End If
    oNode.nodeTypedValue = baData
    pvToBase64 = Replace(Replace(oNode.Text, vbCr, vbNullString), vbLf, vbNullString)
    '--- release picture bytes held by the cached node between calls
    oNode.nodeTypedValue = vbNullString
End Function
