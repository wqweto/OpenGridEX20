Attribute VB_Name = "mdImport"
'=========================================================================
'
' Open GridEX 2000 Control
' Profile-driven object model import from snapshot JSON (see SCHEMA.md)
'
' Applies a snapshot document produced by mdSnapshot.bas onto a live
' GridEX/GEXPreview object model late-bound via CallByName. Collection
' items are re-created through each collection's public Add method.
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

'--- non-failing Variant -> String coercion (Null/Empty/objects -> "")
Public Function C2Str(Value As Variant) As String
    Dim vDest           As Variant

    If VarType(Value) = vbString Then
        C2Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C2Str = vDest
    End If
End Function

Public Sub ImportObject(oObj As Object, sClass As String, oJson As Object)
    Dim uProps()        As UcsProfileProp
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim vValue          As Variant

    If oObj Is Nothing Or oJson Is Nothing Then
        Exit Sub
    End If
    lCount = ProfileForClass(sClass, uProps)
    '--- three passes: fonts first because font changes recalc dependent
    '--- metrics (e.g. RowHeight/ColumnHeaderHeight) which explicit scalar
    '--- values must then override; scalars before compound props so gate
    '--- props (e.g. HasValueList) are in effect for dependents
    For lIdx = 0 To lCount - 1
        Select Case uProps(lIdx).eKind
        Case ucsPrkFont
            pvImportProp oObj, uProps(lIdx), oJson
        End Select
    Next
    For lIdx = 0 To lCount - 1
        Select Case uProps(lIdx).eKind
        Case ucsPrkScalar, ucsPrkEnum, ucsPrkVariant
            pvImportProp oObj, uProps(lIdx), oJson
        End Select
    Next
    For lIdx = 0 To lCount - 1
        Select Case uProps(lIdx).eKind
        Case ucsPrkFont, ucsPrkScalar, ucsPrkEnum, ucsPrkVariant
        Case Else
            pvImportProp oObj, uProps(lIdx), oJson
        End Select
    Next
    '--- parameterized specials
    If sClass = "JSPrinterProperties" Then
        For lPos = 1 To 3
            vValue = JsonValue(oJson, "HeaderString/" & lPos)
            If Not IsEmpty(vValue) Then
                CallByName oObj, "HeaderString", VbLet, lPos, vValue
            End If
            vValue = JsonValue(oJson, "FooterString/" & lPos)
            If Not IsEmpty(vValue) Then
                CallByName oObj, "FooterString", VbLet, lPos, vValue
            End If
        Next
    End If
End Sub

Private Sub pvImportProp(oObj As Object, uProp As UcsProfileProp, oJson As Object)
    Const FUNC_NAME     As String = "pvImportProp"
    Dim vValue          As Variant
    Dim oSub            As Object
    Dim oValue          As Object

    On Error GoTo EH
    '--- let-assigning an object-holding Variant invokes its default
    '--- property, so always tunnel through AssignVariant
    AssignVariant vValue, JsonValue(oJson, uProp.sProp)
    pvTrace uProp.sClass & "." & uProp.sProp & " kind=" & uProp.eKind & " vt=" & VarType(vValue)
    If IsEmpty(vValue) Or IsNull(vValue) Then
        Exit Sub
    End If
    Select Case uProp.eKind
    Case ucsPrkScalar, ucsPrkEnum, ucsPrkVariant
        If uProp.bCanWrite Then
            CallByName oObj, uProp.sProp, VbLet, vValue
        End If
    Case ucsPrkFont
        If IsObject(vValue) Then
            Set oSub = CallByName(oObj, uProp.sProp, VbGet)
            If Not oSub Is Nothing Then
                Set oValue = vValue
                pvImportFont oSub, oValue
            End If
        End If
    Case ucsPrkPicture
        If IsObject(vValue) And uProp.bCanWrite Then
            Set oValue = vValue
            Set oSub = pvPictureFromJson(oValue)
            If Not oSub Is Nothing Then
                CallByName oObj, uProp.sProp, VbSet, oSub
            End If
        End If
    Case ucsPrkClass
        If IsObject(vValue) Then
            Set oSub = CallByName(oObj, uProp.sProp, VbGet)
            Set oValue = vValue
            ImportObject oSub, uProp.sTypeName, oValue
        End If
    Case ucsPrkCollection
        If IsObject(vValue) Then
            Set oSub = CallByName(oObj, uProp.sProp, VbGet)
            If Not oSub Is Nothing Then
                Set oValue = vValue
                If JsonObjectType(oValue) = "array" Then
                    pvImportItems oSub, uProp.sItemClass, oValue
                Else
                    '--- wrapped collection: own props first, then items
                    ImportObject oSub, uProp.sTypeName, oValue
                    AssignVariant vValue, JsonValue(oValue, "items")
                    If IsObject(vValue) Then
                        Set oValue = vValue
                        pvImportItems oSub, uProp.sItemClass, oValue
                    End If
                End If
            End If
        End If
    End Select
    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, Err.Description & " (" & uProp.sClass & "." & uProp.sProp & ") [" & FUNC_NAME & "]"
End Sub

Private Sub pvImportItems(oColl As Object, sItemClass As String, oItems As Object)
    Const FUNC_NAME     As String = "pvImportItems"
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim vItem           As Variant
    Dim oItemJson       As Object
    Dim oItem           As Object

    On Error GoTo EH
    CallByName oColl, "Clear", VbMethod
    lCount = JsonValue(oItems, "-1")
    For lIdx = 0 To lCount - 1
        AssignVariant vItem, JsonValue(oItems, lIdx)
        If sItemClass = "String" Then
            If Not (IsEmpty(vItem) Or IsNull(vItem)) Then
                CallByName oColl, "Add", VbMethod, vItem
            End If
        ElseIf IsObject(vItem) Then
            Set oItemJson = vItem
            Set oItem = pvAddItem(oColl, sItemClass, oItemJson)
            If Not oItem Is Nothing Then
                ImportObject oItem, sItemClass, oItemJson
            End If
        End If
    Next
    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, Err.Description & " (" & sItemClass & " item " & lIdx & ") [" & FUNC_NAME & "]"
End Sub

Private Function pvAddItem(oColl As Object, sItemClass As String, oItemJson As Object) As Object
    Dim oPicture        As Object
    Dim vValue          As Variant

    Select Case sItemClass
    Case "JSColumn"
        Set pvAddItem = CallByName(oColl, "Add", VbMethod, C2Str(JsonValue(oItemJson, "Caption")), _
                JsonValue(oItemJson, "ColumnType"), JsonValue(oItemJson, "EditType"), C2Str(JsonValue(oItemJson, "Key")))
    Case "JSFmtCondition"
        Set pvAddItem = CallByName(oColl, "Add", VbMethod, JsonValue(oItemJson, "ColIndex"), _
                JsonValue(oItemJson, "Operator"), JsonValue(oItemJson, "Value1"), JsonValue(oItemJson, "Value2"), _
                C2Str(JsonValue(oItemJson, "Key")))
    Case "JSGroup", "JSSortKey"
        Set pvAddItem = CallByName(oColl, "Add", VbMethod, JsonValue(oItemJson, "ColIndex"), JsonValue(oItemJson, "SortOrder"))
    Case "JSValueItem"
        Set pvAddItem = CallByName(oColl, "Add", VbMethod, JsonValue(oItemJson, "Value"), _
                C2Str(JsonValue(oItemJson, "Text")), JsonValue(oItemJson, "IconIndex"))
    Case "JSFormatStyle"
        '--- system styles survive the collection Clear and cannot be
        '--- re-added, so upsert: reuse an existing style by name
        Set pvAddItem = pvExistingItem(oColl, C2Str(JsonValue(oItemJson, "Name")))
        If pvAddItem Is Nothing Then
            Set pvAddItem = CallByName(oColl, "Add", VbMethod, C2Str(JsonValue(oItemJson, "Name")))
        End If
    Case "JSGridImage"
        AssignVariant vValue, JsonValue(oItemJson, "Picture")
        If IsObject(vValue) Then
            Set oPicture = vValue
            Set oPicture = pvPictureFromJson(oPicture)
            If Not oPicture Is Nothing Then
                Set pvAddItem = CallByName(oColl, "Add", VbMethod, oPicture)
            End If
        End If
    End Select
End Function

Private Function pvExistingItem(oColl As Object, sName As String) As Object
    '--- non-raising "Item by key or Nothing"
    On Error Resume Next
    Set pvExistingItem = CallByName(oColl, "Item", VbGet, sName)
    On Error GoTo 0
End Function

Private Sub pvTrace(sText As String)
    Dim sFile           As String
    Dim lFile           As Long

    '--- diagnostics only: set OPENGEX_IMPORT_TRACE to a file path
    sFile = Environ$("OPENGEX_IMPORT_TRACE")
    If LenB(sFile) <> 0 Then
        lFile = FreeFile
        Open sFile For Append As #lFile
        Print #lFile, sText
        Close #lFile
    End If
End Sub

Private Sub pvImportFont(oFont As Object, oJson As Object)
    Dim vValue          As Variant

    vValue = JsonValue(oJson, "Name")
    If Not IsEmpty(vValue) Then
        oFont.Name = vValue
    End If
    vValue = JsonValue(oJson, "Size")
    If Not IsEmpty(vValue) Then
        oFont.Size = vValue
    End If
    vValue = JsonValue(oJson, "Bold")
    If Not IsEmpty(vValue) Then
        oFont.Bold = vValue
    End If
    vValue = JsonValue(oJson, "Italic")
    If Not IsEmpty(vValue) Then
        oFont.Italic = vValue
    End If
    vValue = JsonValue(oJson, "Underline")
    If Not IsEmpty(vValue) Then
        oFont.Underline = vValue
    End If
    vValue = JsonValue(oJson, "Strikethrough")
    If Not IsEmpty(vValue) Then
        oFont.Strikethrough = vValue
    End If
    vValue = JsonValue(oJson, "Charset")
    If Not IsEmpty(vValue) Then
        oFont.Charset = vValue
    End If
End Sub

Private Function pvPictureFromJson(oJson As Object) As Object
    Dim vData           As Variant
    Dim baData()        As Byte
    Dim sTempFile       As String
    Dim lFile           As Long

    vData = JsonValue(oJson, "data")
    If IsEmpty(vData) Or IsNull(vData) Then
        Exit Function
    End If
    baData = pvFromBase64(CStr(vData))
    sTempFile = Environ$("TEMP") & "\opengex_pic.tmp"
    lFile = FreeFile
    Open sTempFile For Binary Access Write As #lFile
    Put #lFile, , baData
    Close #lFile
    Set pvPictureFromJson = LoadPicture(sTempFile)
    Kill sTempFile
End Function

Private Function pvFromBase64(sText As String) As Byte()
    Static oNode        As Object

    If oNode Is Nothing Then
        Set oNode = CreateObject("MSXML2.DOMDocument").createElement("b64")
        oNode.DataType = "bin.base64"
    End If
    oNode.Text = sText
    pvFromBase64 = oNode.nodeTypedValue
    oNode.Text = vbNullString
End Function
