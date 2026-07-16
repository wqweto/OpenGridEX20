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

Private Const INTERFACESAFE_FOR_UNTRUSTED_CALLER   As Long = 1
Private Const INTERFACESAFE_FOR_UNTRUSTED_DATA     As Long = 2

'=========================================================================
' Properties
'=========================================================================

Public Property Get TotalPages() As Long
Attribute TotalPages.VB_Description = "Returns the number of pages in a document."
End Property

Public Property Get CurrentPage() As Long
Attribute CurrentPage.VB_Description = "Returns or sets the page displayed."
End Property

Public Property Let CurrentPage(ByVal lValue As Long)
End Property

Public Property Get ToolbarFont() As Font
Attribute ToolbarFont.VB_Description = "Returns/sets a Font object used in the toolbar."
End Property

Public Property Set ToolbarFont(ByVal oValue As Font)
End Property

Public Property Get Zoom() As jgexZoomConstants
Attribute Zoom.VB_Description = "Determines how GEXPreview control should display pages."
End Property

Public Property Let Zoom(ByVal eValue As jgexZoomConstants)
End Property

Public Property Get ToolbarVisible() As Boolean
Attribute ToolbarVisible.VB_Description = "Determines whether the tool bar is displayed."
End Property

Public Property Let ToolbarVisible(ByVal bValue As Boolean)
End Property

Public Property Get PageSetupText() As String
Attribute PageSetupText.VB_Description = "Returns/sets the text displayed in the <Page Setup> button."
End Property

Public Property Let PageSetupText(ByVal sValue As String)
End Property

Public Property Get PrintText() As String
Attribute PrintText.VB_Description = "Returns/sets the text displayed in the <Print> button."
End Property

Public Property Let PrintText(ByVal sValue As String)
End Property

Public Property Get CloseButtonText() As String
Attribute CloseButtonText.VB_Description = "Returns/sets the text displayed in the <Close> button."
End Property

Public Property Let CloseButtonText(ByVal sValue As String)
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the handle of a GEXPreview control."
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Retusns/sets the background color of the control."
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
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
' IObjectSafety interface
'=========================================================================

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, pdwSupportedOptions As Long, pdwEnabledOptions As Long)
    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
    pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub
