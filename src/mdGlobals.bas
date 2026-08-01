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

Public Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = 1
Public Const INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = 2

'--- painting
Public Const OPAQUE                     As Long = 2
Public Const PS_SOLID                   As Long = 0
Public Const PATINVERT                  As Long = &H5A0049
Public Const PATCOPY                    As Long = &HF00021
Public Const PS_DASH                    As Long = 1
Public Const PS_DOT                     As Long = 2
Public Const DT_CENTER                  As Long = 1
Public Const DT_RIGHT                   As Long = 2
Public Const DT_VCENTER                 As Long = 4
Public Const DT_SINGLELINE              As Long = &H20
Public Const DT_CALCRECT                As Long = &H400
Public Const DT_NOPREFIX                As Long = &H800
Public Const DT_NOCLIP                  As Long = &H100
Public Const GWL_STYLE                  As Long = -16
Public Const WS_VSCROLL                 As Long = &H200000
Public Const WS_HSCROLL                 As Long = &H100000
Public Const SB_VERT                    As Long = 1
Public Const SB_CTL                     As Long = 2
Public Const SIF_RANGE                  As Long = 1
Public Const SIF_PAGE                   As Long = 2
Public Const SIF_POS                    As Long = 4
Public Const SWP_NOSIZE                 As Long = 1
Public Const SWP_NOMOVE                 As Long = 2
Public Const SWP_NOZORDER               As Long = 4
Public Const SWP_FRAMECHANGED           As Long = &H20
Public Const SIF_TRACKPOS               As Long = &H10
Public Const SM_CXVSCROLL               As Long = 2
Public Const SM_CYHSCROLL               As Long = 3
Public Const WM_VSCROLL                 As Long = &H115
Public Const WM_KEYDOWN                  As Long = &H100
Public Const WM_LBUTTONDOWN              As Long = &H201
Public Const WM_LBUTTONUP                As Long = &H202
Public Const WM_LBUTTONDBLCLK            As Long = &H203
Public Const WM_MOUSEMOVE                As Long = &H200
Public Const WM_CHAR                     As Long = &H102
Public Const MK_LBUTTON                  As Long = &H1
Public Const MK_RBUTTON                  As Long = &H2
Public Const MK_SHIFT                    As Long = &H4
Public Const MK_CONTROL                  As Long = &H8
Public Const SB_LINEUP                  As Long = 0
Public Const SB_LINEDOWN                As Long = 1
Public Const SB_PAGEUP                  As Long = 2
Public Const SB_PAGEDOWN                As Long = 3
Public Const SB_THUMBPOSITION           As Long = 4
Public Const SB_THUMBTRACK              As Long = 5
Public Const EBMODE_DESIGN              As Long = 0
Public Const WM_MOUSEACTIVATE           As Long = &H21
Public Const WM_CTLCOLORSCROLLBAR       As Long = &H137
Public Const MA_NOACTIVATE              As Long = 3
Public Const DFC_SCROLL                 As Long = 3
Public Const DFCS_SCROLLLEFT            As Long = 2
Public Const DFCS_SCROLLRIGHT           As Long = 3
Public Const DFCS_INACTIVE              As Long = &H100
Public Const DFC_BUTTON                 As Long = 4
Public Const DFCS_BUTTONPUSH            As Long = &H10
Public Const EDGE_SUNKEN                As Long = &HA
Public Const BF_RECT                    As Long = &HF

Public Type POINTAPI
    X                       As Long
    Y                       As Long
End Type

Public Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
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

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
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
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRICW) As Long
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
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'=========================================================================
' Functions
'=========================================================================

'--- metric props are stored internally in pixels and exposed in twips;
'--- conversion snaps to the nearest whole pixel like the original
Public Function ToPixels(ByVal lTwips As Long) As Long
    ToPixels = (lTwips + Screen.TwipsPerPixelY \ 2) \ Screen.TwipsPerPixelY
End Function

Public Function ToTwips(ByVal lPixels As Long) As Long
    ToTwips = lPixels * Screen.TwipsPerPixelY
End Function

Public Function FontTextHeight(oFont As Font) As Long
    Dim pFont           As IFont
    Dim hDC             As Long
    Dim hPrevFont       As Long
    Dim uTm             As TEXTMETRICW

    Set pFont = oFont
    hDC = GetDC(0)
    hPrevFont = SelectObject(hDC, pFont.hFont)
    Call GetTextMetrics(hDC, uTm)
    Call SelectObject(hDC, hPrevFont)
    Call ReleaseDC(0, hDC)
    FontTextHeight = uTm.tmHeight
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
