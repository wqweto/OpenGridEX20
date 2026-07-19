VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSPrinterProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates printer and print job properties to be used when a GridEX control is to going be printed."
'=========================================================================
'
' Open GridEX 2000 Control
' Encapsulates printer and print job properties to be used when a GridEX
' control is going to be printed
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_nCopies                   As Integer
Private m_bCollate                  As Boolean
Private m_nPrintQuality             As Integer
Private m_eColorMode                As jgexPPColorModeConstants
Private m_eOrientation              As jgexPPOrientationConstants
Private m_ePaperBin                 As jgexPPPaperBinConstants
Private m_nPaperSize                As Integer
Private m_lPaperWidth               As Long
Private m_lPaperHeight              As Long
Private m_lLeftMargin               As Long
Private m_lTopMargin                As Long
Private m_lRightMargin              As Long
Private m_lBottomMargin             As Long
Private m_bRepeatHeaders            As Boolean
Private m_bFitColumns               As Boolean
Private m_bTranslateColors          As Boolean
Private m_lHeaderDistance           As Long
Private m_lFooterDistance           As Long
Private m_aHeaderStrings(1 To 3)    As String
Private m_aFooterStrings(1 To 3)    As String
Private m_sDocumentName             As String
Private m_bRepeatFrozenCols         As Boolean
Private m_bPrintPreviewRows         As Boolean
Private m_lCardColumnsPerPage       As Long
Private m_oPageHeaderFont           As Font
Private m_oPageFooterFont           As Font
Private m_bPrintProgressDialog      As Boolean
Private m_bTransparentBackground    As Boolean
Private m_eMeasurementUnits         As jgexMeasurementUnitsConstants
Private m_bPageSetupCanceled        As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Copies() As Integer
Attribute Copies.VB_Description = "Returns/sets the number of copies to be printed."
    Copies = m_nCopies
End Property

Public Property Let Copies(ByVal nValue As Integer)
    m_nCopies = nValue
End Property

Public Property Get Collate() As Boolean
Attribute Collate.VB_Description = "Determines whether collation should be used when printing multiple copies."
    Collate = m_bCollate
End Property

Public Property Let Collate(ByVal bValue As Boolean)
    m_bCollate = bValue
End Property

Public Property Get PrintQuality() As Integer
Attribute PrintQuality.VB_Description = "Returns/sets the printer resolution."
    PrintQuality = m_nPrintQuality
End Property

Public Property Let PrintQuality(ByVal nValue As Integer)
    m_nPrintQuality = nValue
End Property

Public Property Get ColorMode() As jgexPPColorModeConstants
Attribute ColorMode.VB_Description = "Returns/sets a value that determines whether a document should be printed in color or monochrome."
    ColorMode = m_eColorMode
End Property

Public Property Let ColorMode(ByVal eValue As jgexPPColorModeConstants)
    m_eColorMode = eValue
End Property

Public Property Get DeviceName() As String
Attribute DeviceName.VB_Description = "Returns the name of the device a driver supports."
    On Error GoTo EH
    DeviceName = Printer.DeviceName
EH:
End Property

Public Property Get DriverName() As String
Attribute DriverName.VB_Description = "Returns the name of the printer's driver."
    On Error GoTo EH
    DriverName = Printer.DriverName
EH:
End Property

Public Property Get Orientation() As jgexPPOrientationConstants
Attribute Orientation.VB_Description = "Determines whether documents are printed in portrait or landscape mode."
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eValue As jgexPPOrientationConstants)
    m_eOrientation = eValue
End Property

Public Property Get PaperBin() As jgexPPPaperBinConstants
Attribute PaperBin.VB_Description = "Returns/sets a value indicating the default paper bin on the printer from which paper is fed when printing."
    PaperBin = m_ePaperBin
End Property

Public Property Let PaperBin(ByVal eValue As jgexPPPaperBinConstants)
    m_ePaperBin = eValue
End Property

Public Property Get PaperSize() As Integer
Attribute PaperSize.VB_Description = "Returns/sets the paper size for the current printer. "
    PaperSize = m_nPaperSize
End Property

Public Property Let PaperSize(ByVal nValue As Integer)
    m_nPaperSize = nValue
End Property

Public Property Get PaperWidth() As Long
Attribute PaperWidth.VB_Description = "Returns/sets the physical width, in twips, of the paper set up for the printing device."
    PaperWidth = m_lPaperWidth
End Property

Public Property Let PaperWidth(ByVal lValue As Long)
    m_lPaperWidth = lValue
End Property

Public Property Get PaperHeight() As Long
Attribute PaperHeight.VB_Description = "Returns/sets the physical height, in twips, of the paper set up for the printing device."
    PaperHeight = m_lPaperHeight
End Property

Public Property Let PaperHeight(ByVal lValue As Long)
    m_lPaperHeight = lValue
End Property

Public Property Get LeftMargin() As Long
Attribute LeftMargin.VB_Description = "Returns/sets the width, in twips, of the left margin."
    LeftMargin = m_lLeftMargin
End Property

Public Property Let LeftMargin(ByVal lValue As Long)
    m_lLeftMargin = lValue
End Property

Public Property Get TopMargin() As Long
Attribute TopMargin.VB_Description = "Returns/sets the height, in twips, of the top margin."
    TopMargin = m_lTopMargin
End Property

Public Property Let TopMargin(ByVal lValue As Long)
    m_lTopMargin = lValue
End Property

Public Property Get RightMargin() As Long
Attribute RightMargin.VB_Description = "Returns/sets the width, in twips, of the right margin."
    RightMargin = m_lRightMargin
End Property

Public Property Let RightMargin(ByVal lValue As Long)
    m_lRightMargin = lValue
End Property

Public Property Get BottomMargin() As Long
Attribute BottomMargin.VB_Description = "Returns/sets the height, in twips, of the bottom margin."
    BottomMargin = m_lBottomMargin
End Property

Public Property Let BottomMargin(ByVal lValue As Long)
    m_lBottomMargin = lValue
End Property

Public Property Get ClientWidth() As Long
Attribute ClientWidth.VB_Description = "Returns the width, in twips, of the printable area in a page."
    ClientWidth = m_lPaperWidth - m_lLeftMargin - m_lRightMargin
End Property

Public Property Get ClientHeight() As Long
Attribute ClientHeight.VB_Description = "Returns the height, in twips, of the printable area in a page."
    ClientHeight = m_lPaperHeight - m_lTopMargin - m_lBottomMargin
End Property

Public Property Get RepeatHeaders() As Boolean
Attribute RepeatHeaders.VB_Description = "Determines whether column headers should appear on each page."
    RepeatHeaders = m_bRepeatHeaders
End Property

Public Property Let RepeatHeaders(ByVal bValue As Boolean)
    m_bRepeatHeaders = bValue
End Property

Public Property Get FitColumns() As Boolean
Attribute FitColumns.VB_Description = "DEtermines whether the GridEX should scale printed output to fit all visible columns in a page."
    FitColumns = m_bFitColumns
End Property

Public Property Let FitColumns(ByVal bValue As Boolean)
    m_bFitColumns = bValue
End Property

Public Property Get TranslateColors() As Boolean
Attribute TranslateColors.VB_Description = "Determines whether system colors should be translated to black and white colors."
    TranslateColors = m_bTranslateColors
End Property

Public Property Let TranslateColors(ByVal bValue As Boolean)
    m_bTranslateColors = bValue
End Property

Public Property Get HeaderDistance() As Long
Attribute HeaderDistance.VB_Description = "Returns/sets the distance, in twips, from the top edge of the paper to the top edge of the header."
    HeaderDistance = m_lHeaderDistance
End Property

Public Property Let HeaderDistance(ByVal lValue As Long)
    m_lHeaderDistance = lValue
End Property

Public Property Get FooterDistance() As Long
Attribute FooterDistance.VB_Description = "Returns/sets the distance, in twips, from the bottom edge of the paper to the top edge of the footer."
    FooterDistance = m_lFooterDistance
End Property

Public Property Let FooterDistance(ByVal lValue As Long)
    m_lFooterDistance = lValue
End Property

Public Property Get HeaderString(ByVal Position As jgexHeaderFooterPositionConstants) As String
Attribute HeaderString.VB_Description = "Returns/sets the string displayed in the page header."
    HeaderString = m_aHeaderStrings(Position)
End Property

Public Property Let HeaderString(ByVal Position As jgexHeaderFooterPositionConstants, ByVal sValue As String)
    m_aHeaderStrings(Position) = sValue
End Property

Public Property Get FooterString(ByVal Position As jgexHeaderFooterPositionConstants) As String
Attribute FooterString.VB_Description = "Returns/sets the string displayed in the page footer"
    FooterString = m_aFooterStrings(Position)
End Property

Public Property Let FooterString(ByVal Position As jgexHeaderFooterPositionConstants, ByVal sValue As String)
    m_aFooterStrings(Position) = sValue
End Property

Public Property Get DocumentName() As String
Attribute DocumentName.VB_Description = "Returns/sets the name of the document to be printed."
    DocumentName = m_sDocumentName
End Property

Public Property Let DocumentName(ByVal sValue As String)
    m_sDocumentName = sValue
End Property

Public Property Get RepeatFrozenCols() As Boolean
Attribute RepeatFrozenCols.VB_Description = "Determines whether frozen columns should appear on each page."
    RepeatFrozenCols = m_bRepeatFrozenCols
End Property

Public Property Let RepeatFrozenCols(ByVal bValue As Boolean)
    m_bRepeatFrozenCols = bValue
End Property

Public Property Get PrintPreviewRows() As Boolean
Attribute PrintPreviewRows.VB_Description = "Determines whether the control prints preview rows or not."
    PrintPreviewRows = m_bPrintPreviewRows
End Property

Public Property Let PrintPreviewRows(ByVal bValue As Boolean)
    m_bPrintPreviewRows = bValue
End Property

Public Property Get CardColumnsPerPage() As Long
Attribute CardColumnsPerPage.VB_Description = "Returns/sets the number of card columns to be printed in a page."
    CardColumnsPerPage = m_lCardColumnsPerPage
End Property

Public Property Let CardColumnsPerPage(ByVal lValue As Long)
    m_lCardColumnsPerPage = lValue
End Property

Public Property Get PageHeaderFont() As Font
Attribute PageHeaderFont.VB_Description = "Returns/sets a Font object used to draw page headers in a document."
    If m_oPageHeaderFont Is Nothing Then
        Set m_oPageHeaderFont = pvDefaultFont()
    End If
    Set PageHeaderFont = m_oPageHeaderFont
End Property

Public Property Set PageHeaderFont(ByVal oValue As Font)
    Set m_oPageHeaderFont = oValue
End Property

Public Property Get PageFooterFont() As Font
Attribute PageFooterFont.VB_Description = "Returns/sets a Font object used to draw page footers in a document."
    If m_oPageFooterFont Is Nothing Then
        Set m_oPageFooterFont = pvDefaultFont()
    End If
    Set PageFooterFont = m_oPageFooterFont
End Property

Public Property Set PageFooterFont(ByVal oValue As Font)
    Set m_oPageFooterFont = oValue
End Property

Public Property Get PrintProgressDialog() As Boolean
Attribute PrintProgressDialog.VB_Description = "Determines whether print progress dialog should be displayed when GridEX control is printing a document."
    PrintProgressDialog = m_bPrintProgressDialog
End Property

Public Property Let PrintProgressDialog(ByVal bValue As Boolean)
    m_bPrintProgressDialog = bValue
End Property

Public Property Get TransparentBackground() As Boolean
Attribute TransparentBackground.VB_Description = "Determines whether background color should be used to when printing."
    TransparentBackground = m_bTransparentBackground
End Property

Public Property Let TransparentBackground(ByVal bValue As Boolean)
    m_bTransparentBackground = bValue
End Property

Public Property Get MeasurementUnits() As jgexMeasurementUnitsConstants
Attribute MeasurementUnits.VB_Description = "Returns/sets measurement units to be used in the Page Setup dialog."
    MeasurementUnits = m_eMeasurementUnits
End Property

Public Property Let MeasurementUnits(ByVal eValue As jgexMeasurementUnitsConstants)
    m_eMeasurementUnits = eValue
End Property

Public Property Get PageSetupCanceled() As Boolean
Attribute PageSetupCanceled.VB_Description = "Returns True if user pressed cancel in the Page Setup dialog."
    PageSetupCanceled = m_bPageSetupCanceled
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function ChangePrinter(ByVal DeviceName As String, ByVal DriverName As String) As Boolean
Attribute ChangePrinter.VB_Description = "Changes printer's device name and driver."
    Dim oPrinter        As Printer

    For Each oPrinter In Printers
        If StrComp(oPrinter.DeviceName, DeviceName, vbTextCompare) = 0 Then
            Set Printer = oPrinter
            ChangePrinter = True
            Exit For
        End If
    Next
End Function

Public Sub PageSetup(ByVal hWndOwner As Long)
Attribute PageSetup.VB_Description = "Displays the page setup dialog."
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Function pvDefaultFont() As Font
    Dim oFont           As New StdFont

    oFont.Name = "MS Sans Serif"
    oFont.Size = 8.25
    Set pvDefaultFont = oFont
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_nCopies = 1
    m_eColorMode = jgexPPCMMonochrome
    m_eOrientation = jgexPPPortrait
    m_ePaperBin = jgexPPAuto
    m_nPrintQuality = -3
    m_lLeftMargin = 1440
    m_lTopMargin = 1440
    m_lRightMargin = 1440
    m_lBottomMargin = 1440
    m_bRepeatHeaders = True
    m_lHeaderDistance = 720
    m_lFooterDistance = 720
    m_bPrintProgressDialog = True
    m_lCardColumnsPerPage = 2
    m_bPrintPreviewRows = True
End Sub
