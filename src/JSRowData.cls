VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSRowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a container for data, row and cell based formats."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a container for data, row and cell based formats
'
'=========================================================================
Option Explicit

Public Property Get Value(ByVal ColIndex As Integer) As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_Description = "Returns/sets the value of a column."
End Property

Public Property Let Value(ByVal ColIndex As Integer, ByVal vntValue As Variant)
End Property

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Returns the bookmark of the row."
End Property

Public Property Get ColCount() As Integer
Attribute ColCount.VB_Description = "Returns the number of columns."
End Property

Public Property Get GroupLevel() As Integer
Attribute GroupLevel.VB_Description = "Returns the group level of a row if it is a group row or 0 otherwise."
End Property

Public Property Get RowIndex() As Long
Attribute RowIndex.VB_Description = "Returns the row index of a row."
End Property

Public Property Get IconIndex(ByVal ColIndex As Integer) As Integer
Attribute IconIndex.VB_Description = "Returns/sets the index of the JSGridImage displayed in a cell."
End Property

Public Property Let IconIndex(ByVal ColIndex As Integer, ByVal nValue As Integer)
End Property

Public Property Get DisplayValue(ByVal ColIndex As Integer) As String
Attribute DisplayValue.VB_Description = "Returns/sets the display value of a column."
End Property

Public Property Let DisplayValue(ByVal ColIndex As Integer, ByVal sValue As String)
End Property

Public Property Get RowType() As jgexRowTypeConstants
Attribute RowType.VB_Description = "Returns a value that represents the type of a JSRowData object."
End Property

Public Property Get RowStyle() As String
Attribute RowStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in all the cells in a row."
End Property

Public Property Let RowStyle(ByVal sValue As String)
End Property

Public Property Get CellStyle(ByVal ColIndex As Integer) As String
Attribute CellStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in a cell."
End Property

Public Property Let CellStyle(ByVal ColIndex As Integer, ByVal sValue As String)
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records in a JSRowData object representing a group row."
End Property

Public Property Get RowHeight() As Long
Attribute RowHeight.VB_Description = "Returns/sets the height, in twips, of a row."
End Property

Public Property Let RowHeight(ByVal lValue As Long)
End Property

Public Property Get GroupCaption() As String
Attribute GroupCaption.VB_Description = "Returns/sets the caption of a group row."
End Property

Public Property Let GroupCaption(ByVal sValue As String)
End Property

Public Property Get GroupIconIndex() As Integer
Attribute GroupIconIndex.VB_Description = "Returns/sets the icon of a group row."
End Property

Public Property Let GroupIconIndex(ByVal nValue As Integer)
End Property

Public Property Get CellPicture(ByVal ColIndex As Integer) As Picture
Attribute CellPicture.VB_Description = "Returns/sets the foreground picture of a cell."
End Property

Public Property Set CellPicture(ByVal ColIndex As Integer, ByVal oValue As Picture)
End Property

Public Property Get PreviewRowVisible() As Boolean
Attribute PreviewRowVisible.VB_Description = "Determines whether a preview row should be displayed in a row."
End Property

Public Property Let PreviewRowVisible(ByVal bValue As Boolean)
End Property

Public Function GetBookmarks() As Variant
Attribute GetBookmarks.VB_Description = "Returns an array of record bookmarks in a group row."
End Function

Public Function GetRowIndexes() As Variant
Attribute GetRowIndexes.VB_Description = "Returns an array of row indexes in a group row."
End Function

Public Function GetSubTotal(ByVal ColIndex As Integer, ByVal AggregateFunction As jgexAggregateFunctionConstants) As Variant
Attribute GetSubTotal.VB_Description = "Returns the aggregate function result of the set of values contained in a specified column on a group."
End Function
