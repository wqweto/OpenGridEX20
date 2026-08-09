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

Public Type IPAOHookStruct
    lpVTable        As Long     'VTable pointer
    IPAORealPtr     As Long     'Weak-ref for forwarding calls
    CtlPtr          As Long     'Weak-ref for making Friend calls
    ThisPtr         As Long
    CtlName         As String
End Type

'=========================================================================
' API
'=========================================================================

'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Function IsEqualGUID Lib "ole32" (iid1 As Any, iid2 As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

Private Type VBGUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Private Type UcsIPAOHookVTable
    VTable(0 To 9)  As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

'Private Const S_FALSE                       As Long = 1
Private Const S_OK                          As Long = 0
Private Const NULL_PTR                      As Long = 0

Private IID_IOleInPlaceActiveObject     As VBGUID
Private m_uVTable                       As UcsIPAOHookVTable

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

Public Sub InitIPAO(IPAOHookStruct As IPAOHookStruct, oCtl As Object)
    Const FUNC_NAME     As String = "InitIPAO"
    Dim oIPAOReal       As IOleInPlaceActiveObject
    Dim oExt            As VBControlExtender
    
    On Error GoTo EH
    Set oExt = GetExtendedControl(oCtl)
    With IPAOHookStruct
        If Not oExt Is Nothing Then
            .CtlName = TypeName(oExt.Parent) & "." & oExt.Name
        End If
        Set oIPAOReal = oCtl
        .IPAORealPtr = ObjPtr(oIPAOReal)
        .CtlPtr = ObjPtr(oCtl)
        .lpVTable = pvGetVTable
        .ThisPtr = VarPtr(IPAOHookStruct)
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(IPAOHookStruct.CtlName=" & IPAOHookStruct.CtlName & ")"
End Sub

Public Sub TerminateIPAO(IPAOHookStruct As IPAOHookStruct)
    Const FUNC_NAME     As String = "TerminateIPAO"
    Dim oIPAOReal       As IOleInPlaceActiveObject
    Dim oCtl            As Object
    
    On Error GoTo EH
    If IPAOHookStruct.ThisPtr = 0 Then
        Exit Sub
    End If
    With IPAOHookStruct
        If .IPAORealPtr <> 0 Then
            Set oIPAOReal = pvToIOleIPAO(.IPAORealPtr)
            Set oCtl = pvToObject(.CtlPtr)
            pvSetIPAO .CtlPtr, 0
            .IPAORealPtr = 0
            .CtlPtr = 0
            Set oIPAOReal = Nothing
            Set oCtl = Nothing
        End If
        .ThisPtr = 0
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(IPAOHookStruct.CtlName=" & IPAOHookStruct.CtlName & ")"
End Sub

Public Sub SetIPAO(IPAOHookStruct As IPAOHookStruct)
    Const FUNC_NAME     As String = "SetIPAO"
    
    On Error GoTo EH
    If IPAOHookStruct.ThisPtr = 0 Then
        Exit Sub
    End If
    pvSetIPAO IPAOHookStruct.CtlPtr, IPAOHookStruct.ThisPtr
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(IPAOHookStruct.CtlName=" & IPAOHookStruct.CtlName & ")"
End Sub

Public Sub RestoreIPAO(IPAOHookStruct As IPAOHookStruct)
    Const FUNC_NAME     As String = "RestoreIPAO"
    
    On Error GoTo EH
    If IPAOHookStruct.ThisPtr = 0 Then
        Exit Sub
    End If
    pvSetIPAO IPAOHookStruct.CtlPtr, IPAOHookStruct.IPAORealPtr
    Exit Sub
EH:
    RaiseError FUNC_NAME & "(IPAOHookStruct.CtlName=" & IPAOHookStruct.CtlName & ")"
End Sub

Private Sub pvSetIPAO(ByVal lCtlPtr As Long, ByVal pvActiveObj As Long)
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
    '--- moje client site da e not available
    If TypeOf oCtl Is IOleObject Then
        Set pOleObject = oCtl
        If pOleObject.GetClientSite(pOleInPlaceSite) = S_OK And Not pOleInPlaceSite Is Nothing Then
            '--- note: moje da grymne s Access Violation
            On Error Resume Next '--- checked
            pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)
            On Error GoTo EH
            If Not pOleInPlaceFrame Is Nothing Then
                pOleInPlaceFrame.SetActiveObject pvActiveObj, vbNullString
            End If
            If Not pOleInPlaceUIWindow Is Nothing Then '-- And Not m_bMouseActivate
                pOleInPlaceUIWindow.SetActiveObject pvActiveObj, vbNullString
            End If
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

' = private =================================================================

Private Function pvGetVTable() As Long
    Dim STR_RELEASE_THUNK       As String: STR_RELEASE_THUNK = "i1QkBItCBIsIUP9RCMIEAA==" ' 13.5.2020 20:15:19
    Const RELEASE_THUNK_SIZE    As Long = 16

    ' Set up the vTable for the interface and return a pointer to it
    With m_uVTable
        If .VTable(0) = 0 Then
            .VTable(0) = VBA.CLng(AddressOf QueryInterface)
            .VTable(1) = VBA.CLng(AddressOf AddRef)
            .VTable(2) = pvThunkAllocate(STR_RELEASE_THUNK, RELEASE_THUNK_SIZE)
            .VTable(3) = VBA.CLng(AddressOf GetWindow)
            .VTable(4) = VBA.CLng(AddressOf ContextSensitiveHelp)
            .VTable(5) = VBA.CLng(AddressOf TranslateAccelerator)
            .VTable(6) = VBA.CLng(AddressOf OnFrameWindowActivate)
            .VTable(7) = VBA.CLng(AddressOf OnDocWindowActivate)
            .VTable(8) = VBA.CLng(AddressOf ResizeBorder)
            .VTable(9) = VBA.CLng(AddressOf EnableModeless)
            '--- init guid
            With IID_IOleInPlaceActiveObject
               .Data1 = &H117
               .Data4(0) = &HC0
               .Data4(7) = &H46
            End With
        End If
        pvGetVTable = VarPtr(.VTable(0))
    End With
End Function

Private Function pvToObject(ByVal lPtr As Long) As Object
    Call vbaObjSetAddref(pvToObject, lPtr)
End Function

Private Function pvToIOleIPAO(ByVal lPtr As Long) As IOleInPlaceActiveObject
    Call vbaObjSetAddref(pvToIOleIPAO, lPtr)
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

' = interface implemenattion ================================================

Private Function AddRef(This As IPAOHookStruct) As Long
    Const FUNC_NAME     As String = "AddRef"
    
    On Error GoTo EH
    AddRef = pvToIOleIPAO(This.IPAORealPtr).AddRef
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

'Private Function Release(This As IPAOHookStruct) As Long
'    Const FUNC_NAME     As String = "Release"
'
'    On Error GoTo EH
'    Release = pvToIOleIPAO(This.IPAORealPtr).Release
'    Exit Function
'EH:
'    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
'End Function

Private Function QueryInterface(This As IPAOHookStruct, riid As VBGUID, pvObj As Long) As Long
    Const FUNC_NAME     As String = "QueryInterface"
    
    On Error GoTo EH
    If IsEqualGUID(riid, IID_IOleInPlaceActiveObject) Then
        pvObj = This.ThisPtr
        AddRef This
        QueryInterface = 0
    Else
        QueryInterface = pvToIOleIPAO(This.IPAORealPtr).QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function GetWindow(This As IPAOHookStruct, phwnd As Long) As Long
    Const FUNC_NAME     As String = "GetWindow"
    
    On Error GoTo EH
    GetWindow = pvToIOleIPAO(This.IPAORealPtr).GetWindow(phwnd)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function ContextSensitiveHelp(This As IPAOHookStruct, ByVal fEnterMode As Long) As Long
    Const FUNC_NAME     As String = "ContextSensitiveHelp"
    
    On Error GoTo EH
    ContextSensitiveHelp = pvToIOleIPAO(This.IPAORealPtr).ContextSensitiveHelp(fEnterMode)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function TranslateAccelerator(This As IPAOHookStruct, uMsg As APIMSG) As Long
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

Private Function OnFrameWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnFrameWindowActivate"
    
    On Error GoTo EH
    OnFrameWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnFrameWindowActivate(fActivate)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function OnDocWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnDocWindowActivate"
    
    On Error GoTo EH
    OnDocWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnDocWindowActivate(fActivate)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function ResizeBorder(This As IPAOHookStruct, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    Const FUNC_NAME     As String = "ResizeBorder"
    
    On Error GoTo EH
    ResizeBorder = pvToIOleIPAO(This.IPAORealPtr).ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function EnableModeless(This As IPAOHookStruct, ByVal fEnable As Long) As Long
    Const FUNC_NAME     As String = "EnableModeless"
    
    On Error GoTo EH
    EnableModeless = pvToIOleIPAO(This.IPAORealPtr).EnableModeless(fEnable)
    Exit Function
EH:
    PrintError FUNC_NAME & "(This.CtlName=" & This.CtlName & ")"
End Function

Private Function GetExtendedControl(oCtl As IUnknown) As VBControlExtender
    Const FUNC_NAME     As String = "GetExtendedControl"
    Dim pOleObject      As IOleObject
    Dim pOleControlSite As IOleControlSite
    
    On Error GoTo EH
    If Not oCtl Is Nothing Then
        If TypeOf oCtl Is IOleObject Then
            Set pOleObject = oCtl
            If pOleObject.GetClientSite(pOleControlSite) = S_OK And Not pOleControlSite Is Nothing Then
                Set GetExtendedControl = pOleControlSite.GetExtendedControl
            End If
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function
