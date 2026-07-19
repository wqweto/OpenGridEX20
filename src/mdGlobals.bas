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
Public Const TRANSPARENT                As Long = 1
Public Const OPAQUE                     As Long = 2
Public Const PS_SOLID                   As Long = 0
Public Const PS_DASH                    As Long = 1
Public Const PS_DOT                     As Long = 2
Public Const DT_CENTER                  As Long = 1
Public Const DT_RIGHT                   As Long = 2
Public Const DT_VCENTER                 As Long = 4
Public Const DT_SINGLELINE              As Long = &H20
Public Const DT_CALCRECT                As Long = &H400
Public Const DT_NOPREFIX                As Long = &H800

Public Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
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
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRICW) As Long

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

Public Function SearchCollection(ByVal pCol As IVBCollection, Index As Variant, Optional RetVal As Variant) As Boolean
    If Not pCol Is Nothing Then
        SearchCollection = (pCol.Item(Index, RetVal) >= 0)
    End If
End Function
