Attribute VB_Name = "mdGlobals"
'=========================================================================
'
' Open GridEX 2000 Control
' Shared helper procedures
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

#If Not VBA And Not TWINBASIC Then
Public Enum LongPtr
    [_]
End Enum
#End If

#If Win64 Then
Public Const PTR_SIZE                   As Long = 8
Public Const NULL_PTR                   As LongPtr = 0
#Else
Public Const PTR_SIZE                   As Long = 4
Public Const NULL_PTR                   As Long = 0
#End If
'--- IObjectSafety
Public Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = 1
Public Const INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = 2
'--- Windows messages
Public Const WM_SETTEXT                 As Long = &HC
Public Const WM_GETTEXT                 As Long = &HD
Public Const WM_GETTEXTLENGTH           As Long = &HE
Public Const WM_PAINT                   As Long = &HF
Public Const WM_ERASEBKGND              As Long = &H14
Public Const WM_CANCELMODE              As Long = &H1F
Public Const WM_SETCURSOR               As Long = &H20
Public Const WM_MOUSEACTIVATE           As Long = &H21
Public Const WM_SETFONT                 As Long = &H30
Public Const EM_SETSEL                  As Long = &HB1
Public Const EM_LIMITTEXT               As Long = &HC5
Public Const EM_SETMARGINS              As Long = &HD3
Public Const EM_CHARFROMPOS             As Long = &HD7
Public Const WM_KEYDOWN                 As Long = &H100
Public Const WM_KEYUP                   As Long = &H101
Public Const WM_CHAR                    As Long = &H102
Public Const WM_COMMAND                 As Long = &H111
Public Const WM_TIMER                   As Long = &H113
Public Const WM_HSCROLL                 As Long = &H114
Public Const WM_VSCROLL                 As Long = &H115
Public Const WM_CTLCOLORSCROLLBAR       As Long = &H137
Public Const WM_MOUSEMOVE               As Long = &H200
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
Public Const WM_LBUTTONDBLCLK           As Long = &H203
Public Const EN_CHANGE                  As Long = &H300
'--- Edit control margins
Public Const EC_LEFTMARGIN              As Long = 1
'--- Window styles
Public Const ES_LEFT                    As Long = 0
Public Const ES_CENTER                  As Long = 1
Public Const ES_RIGHT                   As Long = 2
Public Const ES_MULTILINE               As Long = 4
Public Const ES_AUTOVSCROLL             As Long = &H40
Public Const ES_AUTOHSCROLL             As Long = &H80
Public Const WS_HSCROLL                 As Long = &H100000
Public Const WS_VSCROLL                 As Long = &H200000
Public Const WS_VISIBLE                 As Long = &H10000000
Public Const WS_CHILD                   As Long = &H40000000
'--- Window placement
Public Const SW_HIDE                    As Long = 0
Public Const SWP_NOSIZE                 As Long = 1
Public Const SWP_NOMOVE                 As Long = 2
Public Const SWP_NOZORDER               As Long = 4
Public Const SWP_FRAMECHANGED           As Long = &H20
Public Const SWP_SHOWWINDOW             As Long = &H40
Public Const SWP_HIDEWINDOW             As Long = &H80
'--- Window long offsets
Public Const GWL_STYLE                  As Long = -16
'--- Mouse keys
Public Const MK_LBUTTON                 As Long = &H1
Public Const MK_RBUTTON                 As Long = &H2
Public Const MK_SHIFT                   As Long = &H4
Public Const MK_CONTROL                 As Long = &H8
'--- Mouse activate
Public Const MA_NOACTIVATE              As Long = 3
'--- Cursors
Public Const IDC_SIZEWE                 As Long = 32644
Public Const IDC_SIZEALL                As Long = 32646
'--- Scroll bar codes
Public Const SB_LINELEFT                As Long = 0
Public Const SB_LINEUP                  As Long = 0
Public Const SB_LINEDOWN                As Long = 1
Public Const SB_LINERIGHT               As Long = 1
Public Const SB_VERT                    As Long = 1
Public Const SB_CTL                     As Long = 2
Public Const SB_PAGELEFT                As Long = 2
Public Const SB_PAGEUP                  As Long = 2
Public Const SB_PAGEDOWN                As Long = 3
Public Const SB_PAGERIGHT               As Long = 3
Public Const SB_THUMBPOSITION           As Long = 4
Public Const SB_THUMBTRACK              As Long = 5
Public Const SB_LEFT                    As Long = 6
Public Const SB_RIGHT                   As Long = 7
'--- Scroll info flags
Public Const SIF_RANGE                  As Long = 1
Public Const SIF_PAGE                   As Long = 2
Public Const SIF_POS                    As Long = 4
Public Const SIF_TRACKPOS               As Long = &H10
'--- System metrics
Public Const SM_CXVSCROLL               As Long = 2
Public Const SM_CYHSCROLL               As Long = 3
Public Const SM_CXDOUBLECLK             As Long = 36
Public Const SM_CYDOUBLECLK             As Long = 37
Public Const LOGPIXELSY                 As Long = 90
Public Const SM_REMOTESESSION           As Long = &H1000
'--- DrawText flags
Public Const DT_CENTER                  As Long = 1
Public Const DT_RIGHT                   As Long = 2
Public Const DT_VCENTER                 As Long = 4
Public Const DT_WORDBREAK               As Long = &H10
Public Const DT_SINGLELINE              As Long = &H20
Public Const DT_NOCLIP                  As Long = &H100
Public Const DT_CALCRECT                As Long = &H400
Public Const DT_NOPREFIX                As Long = &H800
Public Const DT_END_ELLIPSIS            As Long = &H8000&
'--- DrawFrameControl flags
Public Const DFCS_BUTTONCHECK           As Long = &H0
Public Const DFCS_SCROLLLEFT            As Long = 2
Public Const DFCS_SCROLLRIGHT           As Long = 3
Public Const DFC_SCROLL                 As Long = 3
Public Const DFC_BUTTON                 As Long = 4
Public Const DFCS_BUTTONPUSH            As Long = &H10
Public Const DFCS_INACTIVE              As Long = &H100
Public Const DFCS_CHECKED               As Long = &H400
Public Const DFCS_FLAT                  As Long = &H4000
'--- DrawEdge flags
Public Const EDGE_SUNKEN                As Long = &HA
Public Const BF_RECT                    As Long = &HF
'--- Pens, brushes and raster ops
Public Const PS_SOLID                   As Long = 0
Public Const PS_DASH                    As Long = 1
Public Const OPAQUE                     As Long = 2
Public Const PS_DOT                     As Long = 2
Public Const DSTINVERT                  As Long = &H550009
Public Const PATINVERT                  As Long = &H5A0049
Public Const SRCCOPY                    As Long = &HCC0020
Public Const PATCOPY                    As Long = &HF00021
'--- VB extender
Public Const EBMODE_DESIGN              As Long = 0

Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Public Declare Function OleTranslateColor Lib "olepro32" (ByVal clrOle As OLE_COLOR, ByVal hPal As Long, clrRef As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function IntersectClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRICW) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpString As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, ByVal fRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Type POINTAPI
    X                       As Long
    Y                       As Long
End Type

Public Type SIZEAPI
    cx                      As Long
    cy                      As Long
End Type

Public Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Public Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(0 To 31)    As Byte
End Type

Public Type SCROLLINFO
    cbSize                  As Long
    fMask                   As Long
    nMin                    As Long
    nMax                    As Long
    nPage                   As Long
    nPos                    As Long
    nTrackPos               As Long
End Type

Public Type TEXTMETRICW
    tmHeight                As Long
    tmAscent                As Long
    tmDescent               As Long
    tmInternalLeading       As Long
    tmExternalLeading       As Long
    tmAveCharWidth          As Long
    tmMaxCharWidth          As Long
    tmWeight                As Long
    tmOverhang              As Long
    tmDigitizedAspectX      As Long
    tmDigitizedAspectY      As Long
    tmFirstChar             As Integer
    tmLastChar              As Integer
    tmDefaultChar           As Integer
    tmBreakChar             As Integer
    tmItalic                As Byte
    tmUnderlined            As Byte
    tmStruckOut             As Byte
    tmPitchAndFamily        As Byte
    tmCharSet               As Byte
End Type

Public Const LIB_NAME               As String = "OpenGridEX20"

'=========================================================================
' Functions
'=========================================================================

Public Function Clamp(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long) As Long
    Clamp = lValue
    If Clamp < lMin Then
        Clamp = lMin
    ElseIf Clamp > lMax Then
        Clamp = lMax
    End If
End Function

Public Function MakeDWord(ByVal lLoWord As Long, ByVal lHiWord As Long) As Long
    MakeDWord = (lLoWord And &HFFFF&) Or ((lHiWord And &H7FFF&) * &H10000)
    If (lHiWord And &H8000&) <> 0 Then
        MakeDWord = MakeDWord Or &H80000000
    End If
End Function

'--- the halves as LOWORD and HIWORD give them: unsigned
Public Function LoWord(ByVal lValue As Long) As Long
    LoWord = lValue And &HFFFF&
End Function

Public Function HiWord(ByVal lValue As Long) As Long
    HiWord = (lValue \ &H10000) And &HFFFF&
End Function

'--- and the coordinate pair an lParam carries, which is signed
Public Function GetXLParam(ByVal lParam As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, lParam, 2)
    GetXLParam = nWord
End Function

Public Function GetYLParam(ByVal lParam As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, ByVal VarPtr(lParam) + 2, 2)
    GetYLParam = nWord
End Function

Public Function C2Dbl(vValue As Variant) As Double
    Const VARIANT_ALPHABOOL As Long = 2
    Dim vDest           As Variant

    If VarType(vValue) = vbDouble Then
        C2Dbl = vValue
    ElseIf VariantChangeType(vDest, vValue, VARIANT_ALPHABOOL, vbDouble) = 0 Then
        C2Dbl = vDest
    End If
End Function

Public Function C2Str(vValue As Variant) As String
    Const VARIANT_ALPHABOOL As Long = 2
    Dim vDest           As Variant

    If VarType(vValue) = vbString Then
        C2Str = vValue
    ElseIf VariantChangeType(vDest, vValue, VARIANT_ALPHABOOL, vbString) = 0 Then
        C2Str = vDest
    End If
End Function

Public Function C2Lng(vValue As Variant) As Long
    Const VARIANT_ALPHABOOL As Long = 2
    Dim vDest           As Variant

    If VarType(vValue) = vbLong Then
        C2Lng = vValue
    ElseIf VariantChangeType(vDest, vValue, VARIANT_ALPHABOOL, vbLong) = 0 Then
        C2Lng = vDest
    End If
End Function

Public Function C2Date(vValue As Variant) As Date
    Const VARIANT_ALPHABOOL As Long = 2
    Dim vDest           As Variant

    If VarType(vValue) = vbDate Then
        C2Date = vValue
    ElseIf VariantChangeType(vDest, vValue, VARIANT_ALPHABOOL, vbDate) = 0 Then
        C2Date = vDest
    End If
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    '--- a Variant holding an object needs Set, and one holding anything else
    '--- refuses it, so every copy of a cell value goes through here
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function ToPixels(ByVal lTwips As Long) As Long
    ToPixels = (lTwips + Screen.TwipsPerPixelY \ 2) \ Screen.TwipsPerPixelY
End Function

Public Function ToTwips(ByVal lPixels As Long) As Long
    ToTwips = lPixels * Screen.TwipsPerPixelY
End Function

Public Function FontTextMetrics(pFont As IFont, Optional ByVal hDC As Long) As TEXTMETRICW
    Dim hScreenDC       As Long
    Dim hPrevFont       As Long

    '--- a caller in the middle of painting already has a DC; without one the
    '--- screen's does, as the metrics depend on the font and the DPI only
    If hDC = 0 Then
        hScreenDC = GetDC(0)
        hDC = hScreenDC
    End If
    hPrevFont = SelectObject(hDC, pFont.hFont)
    Call GetTextMetrics(hDC, FontTextMetrics)
    Call SelectObject(hDC, hPrevFont)
    If hScreenDC <> 0 Then
        Call ReleaseDC(0, hScreenDC)
    End If
End Function

Public Function NewStdFont() As StdFont
    Dim oFont           As New StdFont

    '--- the single place a default font is minted: only the typeface is
    '--- pinned, the size is whatever StdFont itself defaults to, so no
    '--- point size is hardcoded anywhere in the control
    oFont.Name = "MS Sans Serif"
    Set NewStdFont = oFont
End Function

Public Function CloneFont(pFont As IFont) As StdFont
    '--- IFont.Clone copies every attribute in one call and hands back an
    '--- independent font, so a container's ambient font is never aliased
    If Not pFont Is Nothing Then
        pFont.Clone CloneFont
    End If
End Function

Public Function SearchCollection(ByVal pCol As IVBCollection, Index As Variant, Optional RetVal As Variant) As Boolean
    If Not pCol Is Nothing Then
        SearchCollection = (pCol.Item(Index, RetVal) >= 0)
    End If
End Function

Public Function ToHex(baData() As Byte) As String
    Dim lPtr            As LongPtr
    Dim lFirst          As Long
    Dim lIdx            As Long
    Dim sByte           As String

    Call CopyMemory(lPtr, ByVal ArrPtr(baData), PTR_SIZE)
    If lPtr <> 0 Then
        lFirst = LBound(baData)
        ToHex = String$((UBound(baData) - lFirst + 1) * 2, 48)
        For lIdx = 0 To UBound(baData) - lFirst
            sByte = LCase$(Hex$(baData(lIdx + lFirst)))
            Mid$(ToHex, lIdx * 2 + 3 - Len(sByte)) = sByte
        Next
    End If
End Function

'= error handling ========================================================

Public Function PushError() As Variant
    PushError = Array(Err.Number, Err.Source, Err.Description, Erl)
End Function

Public Sub PopRaiseError(LocalErr As Variant, sModule As String, sFunction As String)
    Err.Raise LocalErr(0), GetErrorSource(sModule, sFunction, LocalErr(2)) & vbCrLf & LocalErr(1), LocalErr(2)
End Sub

Public Sub PopPrintError(LocalErr As Variant, sModule As String, sFunction As String)
    LogError "Critical error: " & LocalErr(2) & " &H" & Hex$(LocalErr(0)) & " [" & GetErrorSource(sModule, sFunction, LocalErr(3)) & "]"
End Sub

Public Function GetErrorSource(sModule As String, sFunction As String, Optional ByVal ErrLine As Long) As String
    GetErrorSource = LIB_NAME & "." & sModule _
        & IIf(LenB(sFunction) <> 0, "." & sFunction, vbNullString) _
        & IIf(ErrLine = 0, vbNullString, "(" & ErrLine & ")")
End Function

Public Sub LogError(sMessage As String, Optional ByVal lLine As Long)
    Static bDisabled    As Boolean
    Dim nFile           As Integer

    '--- an unhandled error in an event surfaces as a modal dialog that wedges
    '--- an automated run, so every handler reports here instead. Logging can
    '--- never take the control down: a failure disables it for the session
    If lLine <> 0 Then
        sMessage = sMessage & " at line " & lLine
    End If
    Debug.Print sMessage
    If bDisabled Then
        Exit Sub
    End If
    On Error GoTo EH
    nFile = FreeFile
    Open Environ$("TEMP") & "\OpenGridEX20.log" For Append As #nFile
    Print #nFile, Format$(Now, "yyyy-mm-dd hh:nn:ss ") & sMessage
    Close #nFile
    Exit Sub
EH:
    bDisabled = True
End Sub

Public Function SetTrue(bValue As Boolean) As Boolean
    bValue = True
    SetTrue = True
End Function

Public Sub AssignWeakRef(oDst As Object, ByVal oSrc As Object)
    Dim bInIde          As Boolean: Debug.Assert SetTrue(bInIde)

    If bInIde Then
        Set oDst = oSrc
    Else
        Call CopyMemory(oDst, oSrc, PTR_SIZE)
    End If
End Sub
