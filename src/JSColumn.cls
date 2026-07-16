VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a column within a GridEX control."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a column within a GridEX control
'
'=========================================================================
Option Explicit

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a string that uniquely identifies a column in a collection."
End Property

Public Property Get HeaderAlignment() As jgexAlignmentConstants
Attribute HeaderAlignment.VB_Description = "Returns/sets a value that determines the alignment of text in the column's header."
End Property

Public Property Let HeaderAlignment(ByVal eValue As jgexAlignmentConstants)
End Property

Public Property Get HeaderIcon() As Integer
Attribute HeaderIcon.VB_Description = "Returns/sets the index of the JSGridImage displayed in the column's header."
End Property

Public Property Let HeaderIcon(ByVal nValue As Integer)
End Property

Public Property Get DefaultIcon() As Integer
Attribute DefaultIcon.VB_Description = "Returns/sets the index of the JSGridImage displayed in the column's cells."
End Property

Public Property Let DefaultIcon(ByVal nValue As Integer)
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of the JSColumn in the JSColumns collection."
End Property

Public Property Get ColPosition() As Integer
Attribute ColPosition.VB_Description = "Returns/sets the position of the column."
End Property

Public Property Let ColPosition(ByVal nValue As Integer)
End Property

Public Property Get ValueList() As JSValueList
Attribute ValueList.VB_Description = "Returns the JSValueList collection for the column."
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value indicating whether an object is visible or hidden."
End Property

Public Property Let Visible(ByVal bValue As Boolean)
End Property

Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns/sets the width of a column."
End Property

Public Property Let Width(ByVal lValue As Long)
End Property

Public Property Get DataField() As String
Attribute DataField.VB_Description = "Returns/sets a value that represents a field in the Recordset."
End Property

Public Property Let DataField(ByVal sValue As String)
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
End Property

Public Property Let Caption(ByVal sValue As String)
End Property

Public Property Get AllowSizing() As Boolean
Attribute AllowSizing.VB_Description = "Returns/sets a value indicating Determines whether a user can resize the column."
End Property

Public Property Let AllowSizing(ByVal bValue As Boolean)
End Property

Public Property Get ColumnType() As jgexColumnTypeConstants
Attribute ColumnType.VB_Description = "Returns/sets a value that indicates how column's contents are displayed."
End Property

Public Property Let ColumnType(ByVal eValue As jgexColumnTypeConstants)
End Property

Public Property Get CardCaption() As Boolean
Attribute CardCaption.VB_Description = "Determines whether the column's value will be displayed in the card caption."
End Property

Public Property Let CardCaption(ByVal bValue As Boolean)
End Property

Public Property Get EditType() As jgexEditTypeConstants
Attribute EditType.VB_Description = "Returns/sets a value that indicates how column is edited by the user."
End Property

Public Property Let EditType(ByVal eValue As jgexEditTypeConstants)
End Property

Public Property Get HasValueList() As Boolean
Attribute HasValueList.VB_Description = "Controls whether the column has a list of values."
End Property

Public Property Let HasValueList(ByVal bValue As Boolean)
End Property

Public Property Get SortOrder() As jgexSortOrderConstants
Attribute SortOrder.VB_Description = "Returns a value that represents the sort applied to a column."
End Property

Public Property Get SortType() As jgexSortTypeConstants
Attribute SortType.VB_Description = "Determines the way that values in the column are sorted and/or grouped."
End Property

Public Property Let SortType(ByVal eValue As jgexSortTypeConstants)
End Property

Public Property Get ShowCaptionInCardView() As Boolean
Attribute ShowCaptionInCardView.VB_Description = "Controls whether the caption of a column is displayed before its value in the card body."
End Property

Public Property Let ShowCaptionInCardView(ByVal bValue As Boolean)
End Property

Public Property Get IsGrouped() As Boolean
Attribute IsGrouped.VB_Description = "Returns a value that indicates whether the Column is grouped."
End Property

Public Property Get Format() As String
Attribute Format.VB_Description = "Returns/sets a value indicating the format string for the Column."
End Property

Public Property Let Format(ByVal sValue As String)
End Property

Public Property Get TextAlignment() As jgexAlignmentConstants
Attribute TextAlignment.VB_Description = "Returns/sets a value that determines the alignment of text in the column."
End Property

Public Property Let TextAlignment(ByVal eValue As jgexAlignmentConstants)
End Property

Public Property Get GroupFormat() As String
Attribute GroupFormat.VB_Description = "Returns/sets a value indicating the format string for the value of the group rows when the control is grouped by the column."
End Property

Public Property Let GroupFormat(ByVal sValue As String)
End Property

Public Property Get GroupPrefix() As String
Attribute GroupPrefix.VB_Description = "Returns/sets the string displayed in a group row before the group value."
End Property

Public Property Let GroupPrefix(ByVal sValue As String)
End Property

Public Property Get GroupEmptyStringCaption() As String
Attribute GroupEmptyStringCaption.VB_Description = "Returns/sets the string displayed in a group row when the value grouped is an empty string."
End Property

Public Property Let GroupEmptyStringCaption(ByVal sValue As String)
End Property

Public Property Get CardIcon() As Boolean
Attribute CardIcon.VB_Description = "Determines whether the column's icon will be displayed in the card caption."
End Property

Public Property Let CardIcon(ByVal bValue As Boolean)
End Property

Public Property Get FetchIcon() As Boolean
Attribute FetchIcon.VB_Description = "Returns/sets whether the FetchIcon event will be fired for the column."
End Property

Public Property Let FetchIcon(ByVal bValue As Boolean)
End Property

Public Property Get FetchData() As Boolean
Attribute FetchData.VB_Description = "Returns/sets whether the FetchData event will be fired for the column."
End Property

Public Property Let FetchData(ByVal bValue As Boolean)
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Returns/sets an expression that stores any extra data needed for your program. "
End Property

Public Property Let Tag(ByVal sValue As String)
End Property

Public Property Get DataChanged() As Boolean
Attribute DataChanged.VB_Description = "Returns/sets a value indicating that data in the column has changed by some process other than by retrieving data from the current record."
End Property

Public Property Let DataChanged(ByVal bValue As Boolean)
End Property

Public Property Get ReplaceValues() As Boolean
Attribute ReplaceValues.VB_Description = "Returns/sets whether the column values are replaced for using the ValueList for the column."
End Property

Public Property Let ReplaceValues(ByVal bValue As Boolean)
End Property

Public Property Get DefaultValue() As Variant
Attribute DefaultValue.VB_Description = "Returns/sets the default value for the column in a new record."
End Property

Public Property Let DefaultValue(ByVal vntValue As Variant)
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a field."
End Property

Public Property Let MaxLength(ByVal lValue As Long)
End Property

Public Property Get NullBehavior() As jgexNullBehaviorConstants
Attribute NullBehavior.VB_Description = "Determines how empty strings are written in the database."
End Property

Public Property Let NullBehavior(ByVal eValue As jgexNullBehaviorConstants)
End Property

Public Property Get ButtonStyle() As jgexButtonStyleConstants
Attribute ButtonStyle.VB_Description = "Determines if a column will show a button when the user enters edit mode."
End Property

Public Property Let ButtonStyle(ByVal eValue As jgexButtonStyleConstants)
End Property

Public Property Get Selectable() As Boolean
Attribute Selectable.VB_Description = "Controls whether cell can be selected."
End Property

Public Property Let Selectable(ByVal bValue As Boolean)
End Property

Public Property Get CellStyle() As String
Attribute CellStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in all the cells in a column."
End Property

Public Property Let CellStyle(ByVal sValue As String)
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "determines whether text in a cell is wordwrapped."
End Property

Public Property Let WordWrap(ByVal bValue As Boolean)
End Property

Public Property Get AggregateFunction() As jgexAggregateFunctionConstants
Attribute AggregateFunction.VB_Description = "Returns/sets the aggregate function to be shown in group footers for a column."
End Property

Public Property Let AggregateFunction(ByVal eValue As jgexAggregateFunctionConstants)
End Property

Public Property Get MinRowsInCardView() As Long
Attribute MinRowsInCardView.VB_Description = "Returns or sets the minimum number of rows displayed in a word wrap column."
End Property

Public Property Let MinRowsInCardView(ByVal lValue As Long)
End Property

Public Property Get MaxRowsInCardView() As Long
Attribute MaxRowsInCardView.VB_Description = "Returns or sets the maximum number of rows displayed in a word wrap column."
End Property

Public Property Let MaxRowsInCardView(ByVal lValue As Long)
End Property

Public Property Get DropDownControl() As Object
Attribute DropDownControl.VB_Description = "Returns/sets the GridEX control that acts as a drop down list for the column."
End Property

Public Property Set DropDownControl(ByVal oValue As Object)
End Property

Public Property Get HeaderToolTip() As String
Attribute HeaderToolTip.VB_Description = "Returns/sets the text displayed in a tool tip when users moves the mouse pointer over a column header."
End Property

Public Property Let HeaderToolTip(ByVal sValue As String)
End Property

Public Property Get HeaderStyle() As String
Attribute HeaderStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in the column header."
End Property

Public Property Let HeaderStyle(ByVal sValue As String)
End Property

Public Property Get TotalRowFormat() As String
Attribute TotalRowFormat.VB_Description = "Returns/sets the format string for the aggregate function result of a column in the group footer row."
End Property

Public Property Let TotalRowFormat(ByVal sValue As String)
End Property

Public Property Get TotalRowPrefix() As String
Attribute TotalRowPrefix.VB_Description = "Returns/sets the string displayed in a group footer row before the aggregate function result."
End Property

Public Property Let TotalRowPrefix(ByVal sValue As String)
End Property

Public Sub AutoSize()
Attribute AutoSize.VB_Description = "Adjusts the width of a column to accommodate the longest text."
End Sub
