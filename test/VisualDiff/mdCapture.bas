Attribute VB_Name = "mdCapture"
'=========================================================================
'
' Open GridEX 2000 Control
' Screen capture to 24bpp DIB bits, BMP file I/O and pixel diff report
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const SRCCOPY                       As Long = &HCC0020
Private Const DIB_RGB_COLORS                As Long = 0
Private Const BI_RGB                        As Long = 0

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                  As Long
    biWidth                 As Long
    biHeight                As Long
    biPlanes                As Integer
    biBitCount              As Integer
    biCompression           As Long
    biSizeImage             As Long
    biXPelsPerMeter         As Long
    biYPelsPerMeter         As Long
    biClrUsed               As Long
    biClrImportant          As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As Any, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap As Long, ByVal hPalette As Long, hGdipBmp As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal fileName As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal fileName As Long, hGdipBmp As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal hGdipBmp As Long, hBmpReturn As Long, ByVal background As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long, lplpVoid As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function GetObjectApi Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP
    bmType                  As Long
    bmWidth                 As Long
    bmHeight                As Long
    bmWidthBytes            As Long
    bmPlanes                As Integer
    bmBitsPixel             As Integer
    bmBits                  As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function CaptureWindowClient(ByVal hWnd As Long, lWidth As Long, lHeight As Long, baBits() As Byte) As Boolean
    Dim uRect           As RECT
    Dim hDCWnd          As Long
    Dim hDCMem          As Long
    Dim hBmp            As Long
    Dim hPrevBmp        As Long
    Dim uInfo           As BITMAPINFOHEADER

    '--- blit from the window own client DC at (0,0): no screen coordinate
    '--- mapping involved so DPI virtualization cannot offset the capture
    Call GetClientRect(hWnd, uRect)
    lWidth = uRect.Right
    lHeight = uRect.Bottom
    If lWidth <= 0 Or lHeight <= 0 Then
        Exit Function
    End If
    hDCWnd = GetDC(hWnd)
    hDCMem = CreateCompatibleDC(hDCWnd)
    hBmp = CreateCompatibleBitmap(hDCWnd, lWidth, lHeight)
    hPrevBmp = SelectObject(hDCMem, hBmp)
    Call BitBlt(hDCMem, 0, 0, lWidth, lHeight, hDCWnd, 0, 0, SRCCOPY)
    Call SelectObject(hDCMem, hPrevBmp)
    With uInfo
        .biSize = Len(uInfo)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
    End With
    ReDim baBits(0 To pvStride(lWidth) * lHeight - 1) As Byte
    CaptureWindowClient = (GetDIBits(hDCMem, hBmp, 0, lHeight, baBits(0), uInfo, DIB_RGB_COLORS) = lHeight)
    Call DeleteObject(hBmp)
    Call DeleteDC(hDCMem)
    Call ReleaseDC(hWnd, hDCWnd)
End Function

Public Sub SaveBmp(sFile As String, ByVal lWidth As Long, ByVal lHeight As Long, baBits() As Byte)
    Dim uInfo           As BITMAPINFOHEADER
    Dim baHeader(0 To 13) As Byte
    Dim lFile           As Long
    Dim lSize           As Long

    With uInfo
        .biSize = Len(uInfo)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = pvStride(lWidth) * lHeight
    End With
    lSize = 14 + Len(uInfo) + uInfo.biSizeImage
    baHeader(0) = Asc("B")
    baHeader(1) = Asc("M")
    Call CopyMemory(baHeader(2), lSize, 4)
    lSize = 14 + Len(uInfo)
    Call CopyMemory(baHeader(10), lSize, 4)
    If LenB(Dir$(sFile)) <> 0 Then
        Kill sFile
    End If
    lFile = FreeFile
    Open sFile For Binary Access Write As #lFile
    Put #lFile, , baHeader
    Put #lFile, , uInfo
    Put #lFile, , baBits
    Close #lFile
End Sub

Public Function LoadBmp(sFile As String, lWidth As Long, lHeight As Long, baBits() As Byte) As Boolean
    Dim baHeader(0 To 13) As Byte
    Dim uInfo           As BITMAPINFOHEADER
    Dim lFile           As Long

    If LenB(Dir$(sFile)) = 0 Then
        Exit Function
    End If
    lFile = FreeFile
    Open sFile For Binary Access Read As #lFile
    Get #lFile, , baHeader
    Get #lFile, , uInfo
    If uInfo.biBitCount = 24 And uInfo.biCompression = BI_RGB And uInfo.biHeight > 0 Then
        lWidth = uInfo.biWidth
        lHeight = uInfo.biHeight
        ReDim baBits(0 To pvStride(lWidth) * lHeight - 1) As Byte
        Get #lFile, , baBits
        LoadBmp = True
    End If
    Close #lFile
End Function

Public Function DiffBits(baExpected() As Byte, baActual() As Byte, ByVal lWidth As Long, ByVal lHeight As Long, sReport As String) As Long
    Dim lStride         As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lOffset         As Long
    Dim lMinX           As Long
    Dim lMinY           As Long
    Dim lMaxX           As Long
    Dim lMaxY           As Long

    sReport = vbNullString
    If UBound(baExpected) <> UBound(baActual) Then
        DiffBits = lWidth * lHeight
        sReport = "size mismatch"
        Exit Function
    End If
    lStride = pvStride(lWidth)
    lMinX = lWidth
    lMinY = lHeight
    For lY = 0 To lHeight - 1
        For lX = 0 To lWidth - 1
            lOffset = lY * lStride + lX * 3
            If baExpected(lOffset) <> baActual(lOffset) _
                    Or baExpected(lOffset + 1) <> baActual(lOffset + 1) _
                    Or baExpected(lOffset + 2) <> baActual(lOffset + 2) Then
                DiffBits = DiffBits + 1
                If lX < lMinX Then
                    lMinX = lX
                End If
                If lX > lMaxX Then
                    lMaxX = lX
                End If
                If lY < lMinY Then
                    lMinY = lY
                End If
                If lY > lMaxY Then
                    lMaxY = lY
                End If
            End If
        Next
    Next
    If DiffBits > 0 Then
        '--- DIB rows are bottom-up: report in top-down window coords
        sReport = "bbox=(" & lMinX & "," & lHeight - 1 - lMaxY & ")-(" & lMaxX & "," & lHeight - 1 - lMinY & ")"
    End If
End Function

Private Function pvStride(ByVal lWidth As Long) As Long
    pvStride = (lWidth * 3 + 3) And Not 3
End Function

Public Function SavePng(sFile As String, ByVal lWidth As Long, ByVal lHeight As Long, baBits() As Byte) As Boolean
    Dim uInfo           As BITMAPINFOHEADER
    Dim hBmp            As Long
    Dim lPtr            As Long

    With uInfo
        .biSize = Len(uInfo)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
    End With
    hBmp = CreateDIBSection(0, uInfo, DIB_RGB_COLORS, lPtr, 0, 0)
    If hBmp <> 0 Then
        Call CopyMemory(ByVal lPtr, baBits(0), UBound(baBits) + 1)
        If LenB(Dir$(sFile)) <> 0 Then
            Kill sFile
        End If
        SavePng = SaveBitmapAsPng(hBmp, sFile)
        Call DeleteObject(hBmp)
    End If
End Function

Public Function LoadPng(sFile As String, lWidth As Long, lHeight As Long, baBits() As Byte) As Boolean
    Dim uStartup(0 To 3) As Long
    Dim hToken          As Long
    Dim hGdipBmp        As Long
    Dim hBmp            As Long
    Dim uBmp            As BITMAP
    Dim hDCWnd          As Long
    Dim uInfo           As BITMAPINFOHEADER

    If LenB(Dir$(sFile)) = 0 Then
        Exit Function
    End If
    uStartup(0) = 1
    If GdiplusStartup(hToken, uStartup(0), 0) <> 0 Then
        Exit Function
    End If
    If GdipCreateBitmapFromFile(StrPtr(sFile), hGdipBmp) = 0 Then
        If GdipCreateHBITMAPFromBitmap(hGdipBmp, hBmp, &HFFFFFF) = 0 Then
            If GetObjectApi(hBmp, Len(uBmp), uBmp) <> 0 Then
                lWidth = uBmp.bmWidth
                lHeight = uBmp.bmHeight
                With uInfo
                    .biSize = Len(uInfo)
                    .biWidth = lWidth
                    .biHeight = lHeight
                    .biPlanes = 1
                    .biBitCount = 24
                    .biCompression = BI_RGB
                End With
                ReDim baBits(0 To pvStride(lWidth) * lHeight - 1) As Byte
                hDCWnd = GetDC(0)
                LoadPng = (GetDIBits(hDCWnd, hBmp, 0, lHeight, baBits(0), uInfo, DIB_RGB_COLORS) = lHeight)
                Call ReleaseDC(0, hDCWnd)
            End If
            Call DeleteObject(hBmp)
        End If
        Call GdipDisposeImage(hGdipBmp)
    End If
    Call GdiplusShutdown(hToken)
End Function

Public Function SaveBitmapAsPng(ByVal hDib As Long, ByVal sFile As String) As Boolean
    Dim uStartup(0 To 3) As Long
    Dim hToken          As Long
    Dim hBitmap         As Long
    Dim uEncoder(0 To 3) As Long

    uStartup(0) = 1
    If GdiplusStartup(hToken, uStartup(0), 0) <> 0 Then
        Exit Function
    End If
    If GdipCreateBitmapFromHBITMAP(hDib, 0, hBitmap) = 0 Then
        uEncoder(0) = &H557CF406: uEncoder(1) = &H11D31A04      '--- {557CF406-1A04-11D3-9A73-0000F81EF32E}
        uEncoder(2) = &H739A&: uEncoder(3) = &H2EF31EF8
        If GdipSaveImageToFile(hBitmap, StrPtr(sFile), uEncoder(0), ByVal 0) = 0 Then
            SaveBitmapAsPng = True
        End If
        Call GdipDisposeImage(hBitmap)
    End If
    Call GdiplusShutdown(hToken)
End Function
