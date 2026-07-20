VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFormatStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains format settings."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains format settings
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_sName                     As String
Private m_clrBackColor              As OLE_COLOR
Private m_clrForeColor              As OLE_COLOR
Private m_bFontBold                 As Boolean
Private m_bFontItalic               As Boolean
Private m_bFontUnderline            As Boolean
Private m_bFontStrikeThru           As Boolean
Private m_cyFontSize                As Currency
Private m_sFontName                 As String
Private m_nFontCharset              As Integer
Private m_eTextAlignment            As jgexAlignmentConstants
Private m_oBackgroundPicture        As Picture
Private m_oPicture                  As Picture
Private m_eDrawModeBackGround       As jgexDrawModePictureBackgroundConstants
Private m_ePictureHorzAlignment     As jgexHorzPictureAlignmentConstants
Private m_ePictureVertAlignment     As jgexVertPictureAlignmentConstants
Private m_ePictureDrawMode          As jgexPictureDrawModeConstants
Private m_bIsSystem                 As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns a string that uniquely identifies a member in a JSFormatStyles collection."
    Name = m_sName
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color."
    BackColor = m_clrBackColor
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
    m_clrBackColor = lValue
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color."
    ForeColor = m_clrForeColor
End Property

Public Property Let ForeColor(ByVal lValue As OLE_COLOR)
    m_clrForeColor = lValue
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets font bold style."
    FontBold = m_bFontBold
End Property

Public Property Let FontBold(ByVal bValue As Boolean)
    m_bFontBold = bValue
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets font italic style."
    FontItalic = m_bFontItalic
End Property

Public Property Let FontItalic(ByVal bValue As Boolean)
    m_bFontItalic = bValue
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets font underline style."
    FontUnderline = m_bFontUnderline
End Property

Public Property Let FontUnderline(ByVal bValue As Boolean)
    m_bFontUnderline = bValue
End Property

Public Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_Description = "Returns/sets font strikethrough style."
    FontStrikeThru = m_bFontStrikeThru
End Property

Public Property Let FontStrikeThru(ByVal bValue As Boolean)
    m_bFontStrikeThru = bValue
End Property

Public Property Get FontSize() As Currency
Attribute FontSize.VB_Description = "Returns/sets the size of the font."
    FontSize = m_cyFontSize
End Property

Public Property Let FontSize(ByVal cyValue As Currency)
    m_cyFontSize = cyValue
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Returns/sets the typeface name of the font."
    FontName = m_sFontName
End Property

Public Property Let FontName(ByVal sValue As String)
    m_sFontName = sValue
End Property

Public Property Get FontCharset() As Integer
Attribute FontCharset.VB_Description = "Returns/sets the character set used in the font."
    FontCharset = m_nFontCharset
End Property

Public Property Let FontCharset(ByVal nValue As Integer)
    m_nFontCharset = nValue
End Property

Public Property Get TextAlignment() As jgexAlignmentConstants
Attribute TextAlignment.VB_Description = "Determines the alignment of text in a JSFormatStyle object."
    TextAlignment = m_eTextAlignment
End Property

Public Property Let TextAlignment(ByVal eValue As jgexAlignmentConstants)
    m_eTextAlignment = eValue
End Property

Public Property Get BackgroundPicture() As Picture
Attribute BackgroundPicture.VB_Description = "Returns/sets the background picture."
    Set BackgroundPicture = m_oBackgroundPicture
End Property

Public Property Set BackgroundPicture(ByVal oValue As Picture)
    Set m_oBackgroundPicture = oValue
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a picture of a JSFormatStyle object."
    Set Picture = m_oPicture
End Property

Public Property Set Picture(ByVal oValue As Picture)
    Set m_oPicture = oValue
End Property

Public Property Get DrawModeBackGroundPicture() As jgexDrawModePictureBackgroundConstants
Attribute DrawModeBackGroundPicture.VB_Description = "Determines the way a background picture is drawn."
    DrawModeBackGroundPicture = m_eDrawModeBackGround
End Property

Public Property Let DrawModeBackGroundPicture(ByVal eValue As jgexDrawModePictureBackgroundConstants)
    m_eDrawModeBackGround = eValue
End Property

Public Property Get PictureHorzAlignment() As jgexHorzPictureAlignmentConstants
Attribute PictureHorzAlignment.VB_Description = "Determines the horizontal alignment of a picture in a JSFormatStyle object."
    PictureHorzAlignment = m_ePictureHorzAlignment
End Property

Public Property Let PictureHorzAlignment(ByVal eValue As jgexHorzPictureAlignmentConstants)
    m_ePictureHorzAlignment = eValue
End Property

Public Property Get PictureVertAlignment() As jgexVertPictureAlignmentConstants
Attribute PictureVertAlignment.VB_Description = "Determines the vertical alignment of a picture in a JSFormatStyle object."
    PictureVertAlignment = m_ePictureVertAlignment
End Property

Public Property Let PictureVertAlignment(ByVal eValue As jgexVertPictureAlignmentConstants)
    m_ePictureVertAlignment = eValue
End Property

Public Property Get PictureDrawMode() As jgexPictureDrawModeConstants
Attribute PictureDrawMode.VB_Description = "Determines how GridEX control draws cells with pictures."
    PictureDrawMode = m_ePictureDrawMode
End Property

Public Property Let PictureDrawMode(ByVal eValue As jgexPictureDrawModeConstants)
    m_ePictureDrawMode = eValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Reset()
Attribute Reset.VB_Description = "Clears all properties in a JSFormatStyle object."
    m_clrBackColor = vbWindowBackground
    m_clrForeColor = vbWindowText
    m_bFontBold = False
    m_bFontItalic = False
    m_bFontUnderline = False
    m_bFontStrikeThru = False
    m_cyFontSize = 8.25
    m_sFontName = "MS Sans Serif"
    m_nFontCharset = pvDefaultCharset()
    m_eTextAlignment = jgexAlignLeft
    Set m_oBackgroundPicture = Nothing
    Set m_oPicture = Nothing
    m_eDrawModeBackGround = jgexDMCenter
    m_ePictureHorzAlignment = jgexHPALeft
    m_ePictureVertAlignment = jgexVPATop
    m_ePictureDrawMode = jgexTextOnly
End Sub

Friend Sub frInit(sName As String)
    m_sName = sName
End Sub

Friend Property Get frIsSystem() As Boolean
    frIsSystem = m_bIsSystem
End Property

Friend Property Let frIsSystem(ByVal bValue As Boolean)
    m_bIsSystem = bValue
End Property

'=========================================================================
' Functions
'=========================================================================

Private Function pvDefaultCharset() As Integer
    Dim oFont           As New StdFont

    '--- original defaults to the system locale charset (e.g. 204 Cyrillic)
    pvDefaultCharset = oFont.Charset
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Reset
End Sub
