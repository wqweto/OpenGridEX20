VERSION 5.00
Begin VB.UserControl GEXPreview
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
End
Attribute VB_Name = "GEXPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Janus GridEX 2000 Print Preview Control (DAO 3.6 & ADO 2.x)"
'=========================================================================
'
' Open GridEX 2000 Control
' Print preview control for the GridEX control
'
'=========================================================================
Option Explicit
DefObj A-Z

Implements IObjectSafety

'=========================================================================
' Public events
'=========================================================================

Public Event OnCloseClick()
Attribute OnCloseClick.VB_Description = "Occurs when the Close button, in the control's tool bar, has been clicked."
Public Event OnPrintClick(ByVal UsePrintSetupDlg As JSRetBoolean)
Attribute OnPrintClick.VB_Description = "Occurs when the Print button, in the control's tool bar, has been clicked."
Public Event BeforePaginating()
Attribute BeforePaginating.VB_Description = "Occurs before paginating the document."
Public Event AfterPaginating()
Attribute AfterPaginating.VB_Description = "Occurs after paginating the document."
Public Event ZoomChanged()
Attribute ZoomChanged.VB_Description = "Occurs after the Zoom property changes."
Public Event PageChanged()

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lTotalPages               As Long
Private m_lCurrentPage              As Long
Private m_oToolbarFont              As Font
Private m_eZoom                     As jgexZoomConstants
Private m_bToolbarVisible           As Boolean
Private m_sPageSetupText            As String
Private m_sPrintText                As String
Private m_sCloseButtonText          As String
Private m_clrBackColor              As OLE_COLOR

'=========================================================================
' Properties
'=========================================================================

Public Property Get TotalPages() As Long
Attribute TotalPages.VB_Description = "Returns the number of pages in a document."
    TotalPages = m_lTotalPages
End Property

Public Property Get CurrentPage() As Long
Attribute CurrentPage.VB_Description = "Returns or sets the page displayed."
    CurrentPage = m_lCurrentPage
End Property

Public Property Let CurrentPage(ByVal lValue As Long)
    m_lCurrentPage = lValue
End Property

Public Property Get ToolbarFont() As Font
Attribute ToolbarFont.VB_Description = "Returns/sets a Font object used in the toolbar."
    Set ToolbarFont = m_oToolbarFont
End Property

Public Property Set ToolbarFont(ByVal oValue As Font)
    Set m_oToolbarFont = oValue
End Property

Public Property Get Zoom() As jgexZoomConstants
Attribute Zoom.VB_Description = "Determines how GEXPreview control should display pages."
    Zoom = m_eZoom
End Property

Public Property Let Zoom(ByVal eValue As jgexZoomConstants)
    m_eZoom = eValue
End Property

Public Property Get ToolbarVisible() As Boolean
Attribute ToolbarVisible.VB_Description = "Determines whether the tool bar is displayed."
    ToolbarVisible = m_bToolbarVisible
End Property

Public Property Let ToolbarVisible(ByVal bValue As Boolean)
    m_bToolbarVisible = bValue
End Property

Public Property Get PageSetupText() As String
Attribute PageSetupText.VB_Description = "Returns/sets the text displayed in the <Page Setup> button."
    PageSetupText = m_sPageSetupText
End Property

Public Property Let PageSetupText(ByVal sValue As String)
    m_sPageSetupText = sValue
End Property

Public Property Get PrintText() As String
Attribute PrintText.VB_Description = "Returns/sets the text displayed in the <Print> button."
    PrintText = m_sPrintText
End Property

Public Property Let PrintText(ByVal sValue As String)
    m_sPrintText = sValue
End Property

Public Property Get CloseButtonText() As String
Attribute CloseButtonText.VB_Description = "Returns/sets the text displayed in the <Close> button."
    CloseButtonText = m_sCloseButtonText
End Property

Public Property Let CloseButtonText(ByVal sValue As String)
    m_sCloseButtonText = sValue
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the handle of a GEXPreview control."
    hWnd = UserControl.hWnd
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Retusns/sets the background color of the control."
    BackColor = m_clrBackColor
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
    m_clrBackColor = lValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub NextPage()
Attribute NextPage.VB_Description = "Displays the next page in a document."
End Sub

Public Sub PreviousPage()
Attribute PreviousPage.VB_Description = "Displays the previous page in a document."
End Sub

Public Sub Repaginate()
Attribute Repaginate.VB_Description = "Forces the recalculation of the layout for pages in a document."
End Sub

'=========================================================================
' Interface IObjectSafety
'=========================================================================

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, pdwSupportedOptions As Long, pdwEnabledOptions As Long)
    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
    pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    Set m_oToolbarFont = New StdFont
    m_oToolbarFont.Name = "MS Sans Serif"
    m_oToolbarFont.Size = 8.25
    m_clrBackColor = vbButtonFace
    m_bToolbarVisible = True
    m_lCurrentPage = 1
    m_sPageSetupText = "Page Setup..."
    m_sPrintText = "Print..."
    m_sCloseButtonText = "Close"
End Sub