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

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns a string that uniquely identifies a member in a JSFormatStyles collection."
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color."
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color."
End Property

Public Property Let ForeColor(ByVal lValue As OLE_COLOR)
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets font bold style."
End Property

Public Property Let FontBold(ByVal bValue As Boolean)
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets font italic style."
End Property

Public Property Let FontItalic(ByVal bValue As Boolean)
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets font underline style."
End Property

Public Property Let FontUnderline(ByVal bValue As Boolean)
End Property

Public Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_Description = "Returns/sets font strikethrough style."
End Property

Public Property Let FontStrikeThru(ByVal bValue As Boolean)
End Property

Public Property Get FontSize() As Currency
Attribute FontSize.VB_Description = "Returns/sets the size of the font."
End Property

Public Property Let FontSize(ByVal cyValue As Currency)
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Returns/sets the typeface name of the font."
End Property

Public Property Let FontName(ByVal sValue As String)
End Property

Public Property Get FontCharset() As Integer
Attribute FontCharset.VB_Description = "Returns/sets the character set used in the font."
End Property

Public Property Let FontCharset(ByVal nValue As Integer)
End Property

Public Property Get TextAlignment() As jgexAlignmentConstants
Attribute TextAlignment.VB_Description = "Determines the alignment of text in a JSFormatStyle object."
End Property

Public Property Let TextAlignment(ByVal eValue As jgexAlignmentConstants)
End Property

Public Property Get BackgroundPicture() As Picture
Attribute BackgroundPicture.VB_Description = "Returns/sets the background picture."
End Property

Public Property Set BackgroundPicture(ByVal oValue As Picture)
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a picture of a JSFormatStyle object."
End Property

Public Property Set Picture(ByVal oValue As Picture)
End Property

Public Property Get DrawModeBackGroundPicture() As jgexDrawModePictureBackgroundConstants
Attribute DrawModeBackGroundPicture.VB_Description = "Determines the way a background picture is drawn."
End Property

Public Property Let DrawModeBackGroundPicture(ByVal eValue As jgexDrawModePictureBackgroundConstants)
End Property

Public Property Get PictureHorzAlignment() As jgexHorzPictureAlignmentConstants
Attribute PictureHorzAlignment.VB_Description = "Determines the horizontal alignment of a picture in a JSFormatStyle object."
End Property

Public Property Let PictureHorzAlignment(ByVal eValue As jgexHorzPictureAlignmentConstants)
End Property

Public Property Get PictureVertAlignment() As jgexVertPictureAlignmentConstants
Attribute PictureVertAlignment.VB_Description = "Determines the vertical alignment of a picture in a JSFormatStyle object."
End Property

Public Property Let PictureVertAlignment(ByVal eValue As jgexVertPictureAlignmentConstants)
End Property

Public Property Get PictureDrawMode() As jgexPictureDrawModeConstants
Attribute PictureDrawMode.VB_Description = "Determines how GridEX control draws cells with pictures."
End Property

Public Property Let PictureDrawMode(ByVal eValue As jgexPictureDrawModeConstants)
End Property

Public Sub Reset()
Attribute Reset.VB_Description = "Clears all properties in a JSFormatStyle object."
End Sub
