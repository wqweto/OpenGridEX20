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
