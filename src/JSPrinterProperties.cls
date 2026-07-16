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

Public Property Get Copies() As Integer
Attribute Copies.VB_Description = "Returns/sets the number of copies to be printed."
End Property

Public Property Let Copies(ByVal nValue As Integer)
End Property

Public Property Get Collate() As Boolean
Attribute Collate.VB_Description = "Determines whether collation should be used when printing multiple copies."
End Property

Public Property Let Collate(ByVal bValue As Boolean)
End Property

Public Property Get PrintQuality() As Integer
Attribute PrintQuality.VB_Description = "Returns/sets the printer resolution."
End Property

Public Property Let PrintQuality(ByVal nValue As Integer)
End Property

Public Property Get ColorMode() As jgexPPColorModeConstants
Attribute ColorMode.VB_Description = "Returns/sets a value that determines whether a document should be printed in color or monochrome."
End Property

Public Property Let ColorMode(ByVal eValue As jgexPPColorModeConstants)
End Property

Public Property Get DeviceName() As String
Attribute DeviceName.VB_Description = "Returns the name of the device a driver supports."
End Property

Public Property Get DriverName() As String
Attribute DriverName.VB_Description = "Returns the name of the printer's driver."
End Property

Public Property Get Orientation() As jgexPPOrientationConstants
Attribute Orientation.VB_Description = "Determines whether documents are printed in portrait or landscape mode."
End Property

Public Property Let Orientation(ByVal eValue As jgexPPOrientationConstants)
End Property

Public Property Get PaperBin() As jgexPPPaperBinConstants
Attribute PaperBin.VB_Description = "Returns/sets a value indicating the default paper bin on the printer from which paper is fed when printing."
End Property

Public Property Let PaperBin(ByVal eValue As jgexPPPaperBinConstants)
End Property

Public Property Get PaperSize() As Integer
Attribute PaperSize.VB_Description = "Returns/sets the paper size for the current printer. "
End Property

Public Property Let PaperSize(ByVal nValue As Integer)
End Property

Public Property Get PaperWidth() As Long
Attribute PaperWidth.VB_Description = "Returns/sets the physical width, in twips, of the paper set up for the printing device."
End Property

Public Property Let PaperWidth(ByVal lValue As Long)
End Property

Public Property Get PaperHeight() As Long
Attribute PaperHeight.VB_Description = "Returns/sets the physical height, in twips, of the paper set up for the printing device."
End Property

Public Property Let PaperHeight(ByVal lValue As Long)
End Property

Public Property Get LeftMargin() As Long
Attribute LeftMargin.VB_Description = "Returns/sets the width, in twips, of the left margin."
End Property

Public Property Let LeftMargin(ByVal lValue As Long)
End Property

Public Property Get TopMargin() As Long
Attribute TopMargin.VB_Description = "Returns/sets the height, in twips, of the top margin."
End Property

Public Property Let TopMargin(ByVal lValue As Long)
End Property

Public Property Get RightMargin() As Long
Attribute RightMargin.VB_Description = "Returns/sets the width, in twips, of the right margin."
End Property

Public Property Let RightMargin(ByVal lValue As Long)
End Property

Public Property Get BottomMargin() As Long
Attribute BottomMargin.VB_Description = "Returns/sets the height, in twips, of the bottom margin."
End Property

Public Property Let BottomMargin(ByVal lValue As Long)
End Property

Public Property Get ClientWidth() As Long
Attribute ClientWidth.VB_Description = "Returns the width, in twips, of the printable area in a page."
End Property

Public Property Get ClientHeight() As Long
Attribute ClientHeight.VB_Description = "Returns the height, in twips, of the printable area in a page."
End Property

Public Property Get RepeatHeaders() As Boolean
Attribute RepeatHeaders.VB_Description = "Determines whether column headers should appear on each page."
End Property

Public Property Let RepeatHeaders(ByVal bValue As Boolean)
End Property

Public Property Get FitColumns() As Boolean
Attribute FitColumns.VB_Description = "DEtermines whether the GridEX should scale printed output to fit all visible columns in a page."
End Property

Public Property Let FitColumns(ByVal bValue As Boolean)
End Property

Public Property Get TranslateColors() As Boolean
Attribute TranslateColors.VB_Description = "Determines whether system colors should be translated to black and white colors."
End Property

Public Property Let TranslateColors(ByVal bValue As Boolean)
End Property

Public Property Get HeaderDistance() As Long
Attribute HeaderDistance.VB_Description = "Returns/sets the distance, in twips, from the top edge of the paper to the top edge of the header."
End Property

Public Property Let HeaderDistance(ByVal lValue As Long)
End Property

Public Property Get FooterDistance() As Long
Attribute FooterDistance.VB_Description = "Returns/sets the distance, in twips, from the bottom edge of the paper to the top edge of the footer."
End Property

Public Property Let FooterDistance(ByVal lValue As Long)
End Property

Public Property Get HeaderString(ByVal Position As jgexHeaderFooterPositionConstants) As String
Attribute HeaderString.VB_Description = "Returns/sets the string displayed in the page header."
End Property

Public Property Let HeaderString(ByVal Position As jgexHeaderFooterPositionConstants, ByVal sValue As String)
End Property

Public Property Get FooterString(ByVal Position As jgexHeaderFooterPositionConstants) As String
Attribute FooterString.VB_Description = "Returns/sets the string displayed in the page footer"
End Property

Public Property Let FooterString(ByVal Position As jgexHeaderFooterPositionConstants, ByVal sValue As String)
End Property

Public Property Get DocumentName() As String
Attribute DocumentName.VB_Description = "Returns/sets the name of the document to be printed."
End Property

Public Property Let DocumentName(ByVal sValue As String)
End Property

Public Property Get RepeatFrozenCols() As Boolean
Attribute RepeatFrozenCols.VB_Description = "Determines whether frozen columns should appear on each page."
End Property

Public Property Let RepeatFrozenCols(ByVal bValue As Boolean)
End Property

Public Property Get PrintPreviewRows() As Boolean
Attribute PrintPreviewRows.VB_Description = "Determines whether the control prints preview rows or not."
End Property

Public Property Let PrintPreviewRows(ByVal bValue As Boolean)
End Property

Public Property Get CardColumnsPerPage() As Long
Attribute CardColumnsPerPage.VB_Description = "Returns/sets the number of card columns to be printed in a page."
End Property

Public Property Let CardColumnsPerPage(ByVal lValue As Long)
End Property

Public Property Get PageHeaderFont() As Font
Attribute PageHeaderFont.VB_Description = "Returns/sets a Font object used to draw page headers in a document."
End Property

Public Property Set PageHeaderFont(ByVal oValue As Font)
End Property

Public Property Get PageFooterFont() As Font
Attribute PageFooterFont.VB_Description = "Returns/sets a Font object used to draw page footers in a document."
End Property

Public Property Set PageFooterFont(ByVal oValue As Font)
End Property

Public Property Get PrintProgressDialog() As Boolean
Attribute PrintProgressDialog.VB_Description = "Determines whether print progress dialog should be displayed when GridEX control is printing a document."
End Property

Public Property Let PrintProgressDialog(ByVal bValue As Boolean)
End Property

Public Property Get TransparentBackground() As Boolean
Attribute TransparentBackground.VB_Description = "Determines whether background color should be used to when printing."
End Property

Public Property Let TransparentBackground(ByVal bValue As Boolean)
End Property

Public Property Get MeasurementUnits() As jgexMeasurementUnitsConstants
Attribute MeasurementUnits.VB_Description = "Returns/sets measurement units to be used in the Page Setup dialog."
End Property

Public Property Let MeasurementUnits(ByVal eValue As jgexMeasurementUnitsConstants)
End Property

Public Property Get PageSetupCanceled() As Boolean
Attribute PageSetupCanceled.VB_Description = "Returns True if user pressed cancel in the Page Setup dialog."
End Property

Public Function ChangePrinter(ByVal DeviceName As String, ByVal DriverName As String) As Boolean
Attribute ChangePrinter.VB_Description = "Changes printer's device name and driver."
End Function

Public Sub PageSetup(ByVal hWndOwner As Long)
Attribute PageSetup.VB_Description = "Displays the page setup dialog."
End Sub
