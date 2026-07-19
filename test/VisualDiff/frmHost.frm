VERSION 5.00
Begin VB.Form frmHost
   Caption         =   "VisualDiff host"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' Scenario host: creates the control under test at runtime (late bound,
' no compile-time OCX reference), applies scenario props through the
' shared import engine, feeds unbound data via ObjectEvent and captures
' the client area to 24bpp DIB bits
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oExt           As VBControlExtender
Attribute m_oExt.VB_VarHelpID = -1
Private m_oUnboundRows              As Object

'=========================================================================
' Methods
'=========================================================================

Public Function RunScenario(sProgId As String, oScenario As Object, baBits() As Byte, lWidth As Long, lHeight As Long) As Boolean
    Const FUNC_NAME     As String = "RunScenario"
    Dim sApply          As String
    Dim lHwnd           As Long

    On Error GoTo EH
    lWidth = C2Lng(JsonValue(oScenario, "width"))
    If lWidth = 0 Then
        lWidth = 400
    End If
    lHeight = C2Lng(JsonValue(oScenario, "height"))
    If lHeight = 0 Then
        lHeight = 260
    End If
    Move Left, Top, Width - ScaleWidth * Screen.TwipsPerPixelX + lWidth * Screen.TwipsPerPixelX, Height - ScaleHeight * Screen.TwipsPerPixelY + lHeight * Screen.TwipsPerPixelY
    Set m_oUnboundRows = C2Obj(JsonValue(oScenario, "unbound/rows"))
    sApply = C2Str(JsonValue(oScenario, "apply"))
    If sApply = "after-show" Then
        Show
        DoEvents
        pvCreateAndApply sProgId, oScenario
    Else
        pvCreateAndApply sProgId, oScenario
        Show
    End If
    DoEvents
    '--- capture the control window itself; fall back to the form client
    lHwnd = pvControlHwnd()
    If lHwnd = 0 Then
        lHwnd = hWnd
    End If
    RunScenario = pvCaptureStable(lHwnd, lWidth, lHeight, baBits)
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function

Private Function pvControlHwnd() As Long
    Const FUNC_NAME     As String = "pvControlHwnd"

    On Error GoTo EH
    pvControlHwnd = C2Lng(CallByName(m_oExt.Object, "hWnd", VbGet))
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function

Private Function pvCaptureStable(ByVal lHwnd As Long, lWidth As Long, lHeight As Long, baBits() As Byte) As Boolean
    Dim lRetry          As Long
    Dim baPrev()        As Byte
    Dim sReport         As String

    '--- captures can race pending paints: capture repeatedly until two
    '--- consecutive shots are identical
    If Not CaptureWindowClient(lHwnd, lWidth, lHeight, baPrev) Then
        Exit Function
    End If
    For lRetry = 1 To 20
        Sleep 50
        DoEvents
        If Not CaptureWindowClient(lHwnd, lWidth, lHeight, baBits) Then
            Exit Function
        End If
        If DiffBits(baPrev, baBits, lWidth, lHeight, sReport) = 0 Then
            pvCaptureStable = True
            Exit Function
        End If
        baPrev = baBits
    Next
End Function

'=========================================================================
' Functions
'=========================================================================

Private Sub pvCreateAndApply(sProgId As String, oScenario As Object)
    Dim oProps          As Object
    Dim sClass          As String
    Dim lItemCount      As Long

    Set m_oExt = Controls.Add(sProgId, "ctlGrid")
    m_oExt.Move 0, 0, ScaleWidth, ScaleHeight
    m_oExt.Visible = True
    Set oProps = C2Obj(JsonValue(oScenario, "props"))
    If Not oProps Is Nothing Then
        sClass = Split(sProgId, ".")(1)
        ImportObject m_oExt.Object, sClass, oProps
    End If
    '--- unbound data: DataMode/ItemCount applied outside the profile
    '--- import as the original control may reject them at runtime
    If Not m_oUnboundRows Is Nothing Then
        lItemCount = C2Lng(JsonValue(m_oUnboundRows, "-1"))
        pvTrySet "DataMode", 99
        pvTrySet "ItemCount", lItemCount
        pvTryCall "Rebind"
    End If
End Sub

Private Sub pvTrySet(sProp As String, ByVal vValue As Variant)
    Const FUNC_NAME     As String = "pvTrySet"

    On Error GoTo EH
    CallByName m_oExt.Object, sProp, VbLet, vValue
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Sub

Private Sub pvTryCall(sMethod As String)
    Const FUNC_NAME     As String = "pvTryCall"

    On Error GoTo EH
    CallByName m_oExt.Object, sMethod, VbMethod
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oExt_ObjectEvent(Info As EventInfo)
    Const FUNC_NAME     As String = "m_oExt_ObjectEvent"
    Dim lRowIndex       As Long
    Dim oValues         As Object
    Dim oRow            As Object
    Dim nIdx            As Integer

    On Error GoTo EH
    If Info.Name = "UnboundReadData" Then
        If Not m_oUnboundRows Is Nothing Then
            lRowIndex = C2Lng(Info.EventParameters("RowIndex").Value)
            Set oValues = C2Obj(Info.EventParameters("Values").Value)
            Set oRow = C2Obj(JsonValue(m_oUnboundRows, lRowIndex - 1))
            If Not oValues Is Nothing And Not oRow Is Nothing Then
                For nIdx = 1 To C2Lng(CallByName(oValues, "ColCount", VbGet))
                    CallByName oValues, "Value", VbLet, nIdx, JsonValue(oRow, nIdx - 1)
                Next
            End If
        End If
    End If
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Sub
