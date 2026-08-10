Attribute VB_Name = "mdIPAO"
'=========================================================================
'
' Open GridEX 2000 Control
' IInPlaceActiveObject implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdIPAO"

'=========================================================================
' Light-weight object definition
'=========================================================================

Public Type UcsIPAOHook
    lpVTable        As LongPtr
    IPAORealPtr     As LongPtr
    CtlPtr          As LongPtr
    ThisPtr         As LongPtr
    CtlName         As String
End Type

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As Any, iid2 As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As LongPtr) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As LongPtr

Private Type VBGUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const S_OK                      As Long = 0
Private Const NULL_PTR                  As Long = 0

Private IID_IOleInPlaceActiveObject As VBGUID
Private m_uVTable                   As UcsIPAOHookVTable
Private m_lpVTable                  As LongPtr

Private Type UcsIPAOHookVTable
    VTable(0 To 9)  As LongPtr
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunction As String)
    PopRaiseError PushError, MODULE_NAME, sFunction
End Sub

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub InitIPAO(uHook As UcsIPAOHook, oCtl As Object)
    Const FUNC_NAME     As String = "InitIPAO"
    Dim oIPAOReal       As IOleInPlaceActiveObject
    Dim oExt            As VBControlExtender
    
    On Error GoTo EH
    Set oExt = GetExtendedControl(oCtl)
    With uHook
        If Not oExt Is Nothing Then
            .CtlName = TypeName(oExt.Parent) & "." & oExt.Name
        End If
        Set oIPAOReal = oCtl
        .IPAORealPtr = ObjPtr(oIPAOReal)
        .CtlPtr = ObjPtr(oCtl)
        .lpVTable = pvGetVTable
        .ThisPtr = VarPtr(uHook)
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")"
End Sub

Public Sub TerminateIPAO(uHook As UcsIPAOHook)
    Const FUNC_NAME     As String = "TerminateIPAO"
    
    On Error GoTo EH
    If uHook.ThisPtr = 0 Then
        Exit Sub
    End If
    With uHook
        If .IPAORealPtr <> 0 Then
            pvSetIPAO .CtlPtr, 0
            .IPAORealPtr = 0
            .CtlPtr = 0
        End If
        .ThisPtr = 0
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")"
End Sub

Public Sub SetIPAO(uHook As UcsIPAOHook)
    Const FUNC_NAME     As String = "SetIPAO"
    
    On Error GoTo EH
    If uHook.ThisPtr = 0 Then
        Exit Sub
    End If
    pvSetIPAO uHook.CtlPtr, uHook.ThisPtr
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")"
End Sub

Public Sub RestoreIPAO(uHook As UcsIPAOHook)
    Const FUNC_NAME     As String = "RestoreIPAO"
    
    On Error GoTo EH
    If uHook.ThisPtr = 0 Then
        Exit Sub
    End If
    pvSetIPAO uHook.CtlPtr, uHook.IPAORealPtr
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")"
End Sub

'= private =================================================================

Private Sub pvSetIPAO(ByVal lCtlPtr As LongPtr, ByVal lActiveObjPtr As LongPtr)
    Const FUNC_NAME         As String = "pvSetIPAO"
    Dim oCtl                As Object
    Dim pOleObject          As IOleObject
    Dim pOleInPlaceSite     As IOleInPlaceSite
    Dim pOleInPlaceFrame    As IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
    Dim rcPos               As RECT
    Dim rcClip              As RECT
    Dim uFrameInfo          As OLEINPLACEFRAMEINFO
       
    On Error GoTo EH
    If lCtlPtr = 0 Then
        Exit Sub
    End If
    Set oCtl = pvToObject(lCtlPtr)
    If Not TypeOf oCtl Is IOleObject Then
        Exit Sub
    End If
    Set pOleObject = oCtl
    If pOleObject.GetClientSite(pOleInPlaceSite) <> S_OK Then
        Exit Sub
    End If
    If pOleInPlaceSite Is Nothing Then
        Exit Sub
    End If
    On Error Resume Next '--- checked
    pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)
    On Error GoTo EH
    If Not pOleInPlaceFrame Is Nothing Then
        pOleInPlaceFrame.SetActiveObject lActiveObjPtr, vbNullString
    End If
    If Not pOleInPlaceUIWindow Is Nothing Then
        pOleInPlaceUIWindow.SetActiveObject lActiveObjPtr, vbNullString
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvGetVTable() As Long
    Dim STR_RELEASE_THUNK       As String: STR_RELEASE_THUNK = "i1QkBItCBIsIUP9RCMIEAA==" ' 13.5.2020 20:15:19
    Const RELEASE_THUNK_SIZE    As Long = 16
    
    If m_lpVTable = 0 Then
        '--- init guid
        With IID_IOleInPlaceActiveObject
           .Data1 = &H117
           .Data4(0) = &HC0
           .Data4(7) = &H46
        End With
        With m_uVTable
            .VTable(0) = VBA.CLng(AddressOf QueryInterface)
            .VTable(1) = VBA.CLng(AddressOf AddRef)
            .VTable(2) = ThunkAllocate(STR_RELEASE_THUNK, RELEASE_THUNK_SIZE)
            .VTable(3) = VBA.CLng(AddressOf GetWindow)
            .VTable(4) = VBA.CLng(AddressOf ContextSensitiveHelp)
            .VTable(5) = VBA.CLng(AddressOf TranslateAccelerator)
            .VTable(6) = VBA.CLng(AddressOf OnFrameWindowActivate)
            .VTable(7) = VBA.CLng(AddressOf OnDocWindowActivate)
            .VTable(8) = VBA.CLng(AddressOf ResizeBorder)
            .VTable(9) = VBA.CLng(AddressOf EnableModeless)
        End With
        m_lpVTable = CoTaskMemAlloc(LenB(m_uVTable))
        Call CopyMemory(ByVal m_lpVTable, m_uVTable, LenB(m_uVTable))
    End If
    pvGetVTable = m_lpVTable
End Function

Private Function pvToObject(ByVal lPtr As LongPtr) As Object
    Call vbaObjSetAddref(pvToObject, lPtr)
End Function

Private Function pvToIOleIPAO(ByVal lPtr As LongPtr) As IOleInPlaceActiveObject
    Call vbaObjSetAddref(pvToIOleIPAO, lPtr)
End Function

Private Function GetExtendedControl(oCtl As IUnknown) As VBControlExtender
    Const FUNC_NAME     As String = "GetExtendedControl"
    Dim pOleObject      As IOleObject
    Dim pOleControlSite As IOleControlSite
    
    On Error GoTo EH
    If oCtl Is Nothing Then
        Exit Function
    End If
    If Not TypeOf oCtl Is IOleObject Then
        Exit Function
    End If
    Set pOleObject = oCtl
    If pOleObject.GetClientSite(pOleControlSite) <> S_OK Then
        Exit Function
    End If
    If Not pOleControlSite Is Nothing Then
        Set GetExtendedControl = pOleControlSite.GetExtendedControl
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= interface implemenattion ================================================

Private Function AddRef(This As UcsIPAOHook) As Long
    Const FUNC_NAME     As String = "AddRef"
    
    On Error GoTo EH
    AddRef = pvToIOleIPAO(This.IPAORealPtr).AddRef
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

'Private Function Release(This As UcsIPAOHook) As Long
'    Const FUNC_NAME     As String = "Release"
'
'    On Error GoTo EH
'    Release = pvToIOleIPAO(This.IPAORealPtr).Release
'    Exit Function
'EH:
'    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
'End Function

Private Function QueryInterface(This As UcsIPAOHook, riid As VBGUID, pObjPtr As LongPtr) As Long
    Const FUNC_NAME     As String = "QueryInterface"
    
    On Error GoTo EH
    If IsEqualGUID(riid, IID_IOleInPlaceActiveObject) Then
        pObjPtr = This.ThisPtr
        AddRef This
        QueryInterface = 0
    Else
        QueryInterface = pvToIOleIPAO(This.IPAORealPtr).QueryInterface(ByVal VarPtr(riid), pObjPtr)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function GetWindow(This As UcsIPAOHook, hWnd As LongPtr) As Long
    Const FUNC_NAME     As String = "GetWindow"
    
    On Error GoTo EH
    GetWindow = pvToIOleIPAO(This.IPAORealPtr).GetWindow(hWnd)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function ContextSensitiveHelp(This As UcsIPAOHook, ByVal fEnterMode As Long) As Long
    Const FUNC_NAME     As String = "ContextSensitiveHelp"
    
    On Error GoTo EH
    ContextSensitiveHelp = pvToIOleIPAO(This.IPAORealPtr).ContextSensitiveHelp(fEnterMode)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function TranslateAccelerator(This As UcsIPAOHook, uMsg As APIMSG) As Long
    Const FUNC_NAME     As String = "TranslateAccelerator"
    Dim oGrid           As GridEX
    Dim bHandled        As Boolean
    Dim uOrigMsg        As APIMSG
    Dim pIPAOReal       As IOleInPlaceActiveObject
    
    On Error GoTo EH
    If This.CtlPtr <> 0 Then
        Set oGrid = pvToObject(This.CtlPtr)
        uOrigMsg = uMsg
        bHandled = oGrid.frBeforeTranslateAccel(uMsg)
        Set oGrid = Nothing
    End If
    If bHandled Then
        TranslateAccelerator = S_OK
    Else
        '--- skip refcounting on This.IPAORealPtr
        Call CopyMemory(pIPAOReal, This.IPAORealPtr, 4)
        TranslateAccelerator = pIPAOReal.TranslateAccelerator(ByVal VarPtr(uMsg))
        Call CopyMemory(pIPAOReal, NULL_PTR, 4)
    End If
    If This.CtlPtr <> 0 Then
        Set oGrid = pvToObject(This.CtlPtr)
        oGrid.frAfterTranslateAccel uOrigMsg
        Set oGrid = Nothing
    End If
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function OnFrameWindowActivate(This As UcsIPAOHook, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnFrameWindowActivate"
    
    On Error GoTo EH
    OnFrameWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnFrameWindowActivate(fActivate)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function OnDocWindowActivate(This As UcsIPAOHook, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnDocWindowActivate"
    
    On Error GoTo EH
    OnDocWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnDocWindowActivate(fActivate)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function ResizeBorder(This As UcsIPAOHook, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    Const FUNC_NAME     As String = "ResizeBorder"
    
    On Error GoTo EH
    ResizeBorder = pvToIOleIPAO(This.IPAORealPtr).ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function EnableModeless(This As UcsIPAOHook, ByVal fEnable As Long) As Long
    Const FUNC_NAME     As String = "EnableModeless"
    
    On Error GoTo EH
    EnableModeless = pvToIOleIPAO(This.IPAORealPtr).EnableModeless(fEnable)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function
