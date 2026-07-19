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
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_sKey                      As String
Private m_eHeaderAlignment          As jgexAlignmentConstants
Private m_nHeaderIcon               As Integer
Private m_nDefaultIcon              As Integer
Private m_nIndex                    As Integer
Private m_nColPosition              As Integer
Private m_oValueList                As JSValueList
Private m_bVisible                  As Boolean
Private m_lWidth                    As Long
Private m_sDataField                As String
Private m_sCaption                  As String
Private m_bAllowSizing              As Boolean
Private m_eColumnType               As jgexColumnTypeConstants
Private m_bCardCaption              As Boolean
Private m_eEditType                 As jgexEditTypeConstants
Private m_bHasValueList             As Boolean
Private m_eSortOrder                As jgexSortOrderConstants
Private m_eSortType                 As jgexSortTypeConstants
Private m_bShowCaptionInCardView    As Boolean
Private m_bIsGrouped                As Boolean
Private m_sFormat                   As String
Private m_eTextAlignment            As jgexAlignmentConstants
Private m_sGroupFormat              As String
Private m_sGroupPrefix              As String
Private m_sGroupEmptyStringCaption  As String
Private m_bCardIcon                 As Boolean
Private m_bFetchIcon                As Boolean
Private m_bFetchData                As Boolean
Private m_sTag                      As String
Private m_bDataChanged              As Boolean
Private m_bReplaceValues            As Boolean
Private m_vDefaultValue             As Variant
Private m_lMaxLength                As Long
Private m_eNullBehavior             As jgexNullBehaviorConstants
Private m_eButtonStyle              As jgexButtonStyleConstants
Private m_bSelectable               As Boolean
Private m_sCellStyle                As String
Private m_bWordWrap                 As Boolean
Private m_eAggregateFunction        As jgexAggregateFunctionConstants
Private m_lMinRowsInCardView        As Long
Private m_lMaxRowsInCardView        As Long
Private m_oDropDownControl          As Object
Private m_sHeaderToolTip            As String
Private m_sHeaderStyle              As String
Private m_sTotalRowFormat           As String
Private m_sTotalRowPrefix           As String

'=========================================================================
' Properties
'=========================================================================

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a string that uniquely identifies a column in a collection."
    Key = m_sKey
End Property

Public Property Get HeaderAlignment() As jgexAlignmentConstants
Attribute HeaderAlignment.VB_Description = "Returns/sets a value that determines the alignment of text in the column's header."
    HeaderAlignment = m_eHeaderAlignment
End Property

Public Property Let HeaderAlignment(ByVal eValue As jgexAlignmentConstants)
    m_eHeaderAlignment = eValue
End Property

Public Property Get HeaderIcon() As Integer
Attribute HeaderIcon.VB_Description = "Returns/sets the index of the JSGridImage displayed in the column's header."
    HeaderIcon = m_nHeaderIcon
End Property

Public Property Let HeaderIcon(ByVal nValue As Integer)
    m_nHeaderIcon = nValue
End Property

Public Property Get DefaultIcon() As Integer
Attribute DefaultIcon.VB_Description = "Returns/sets the index of the JSGridImage displayed in the column's cells."
    DefaultIcon = m_nDefaultIcon
End Property

Public Property Let DefaultIcon(ByVal nValue As Integer)
    m_nDefaultIcon = nValue
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of the JSColumn in the JSColumns collection."
    Index = m_nIndex
End Property

Public Property Get ColPosition() As Integer
Attribute ColPosition.VB_Description = "Returns/sets the position of the column."
    ColPosition = m_nColPosition
End Property

Public Property Let ColPosition(ByVal nValue As Integer)
    m_nColPosition = nValue
End Property

Public Property Get ValueList() As JSValueList
Attribute ValueList.VB_Description = "Returns the JSValueList collection for the column."
    '--- original returns Nothing until HasValueList is set
    If m_bHasValueList Then
        If m_oValueList Is Nothing Then
            Set m_oValueList = New JSValueList
        End If
        Set ValueList = m_oValueList
    End If
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value indicating whether an object is visible or hidden."
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal bValue As Boolean)
    m_bVisible = bValue
End Property

Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns/sets the width of a column."
    Width = ToTwips(m_lWidth)
End Property

Public Property Let Width(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original
    m_lWidth = ToPixels(lValue)
End Property

Public Property Get DataField() As String
Attribute DataField.VB_Description = "Returns/sets a value that represents a field in the Recordset."
    DataField = m_sDataField
End Property

Public Property Let DataField(ByVal sValue As String)
    m_sDataField = sValue
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
    Caption = m_sCaption
End Property

Public Property Let Caption(ByVal sValue As String)
    m_sCaption = sValue
End Property

Public Property Get AllowSizing() As Boolean
Attribute AllowSizing.VB_Description = "Returns/sets a value indicating Determines whether a user can resize the column."
    AllowSizing = m_bAllowSizing
End Property

Public Property Let AllowSizing(ByVal bValue As Boolean)
    m_bAllowSizing = bValue
End Property

Public Property Get ColumnType() As jgexColumnTypeConstants
Attribute ColumnType.VB_Description = "Returns/sets a value that indicates how column's contents are displayed."
    ColumnType = m_eColumnType
End Property

Public Property Let ColumnType(ByVal eValue As jgexColumnTypeConstants)
    m_eColumnType = eValue
End Property

Public Property Get CardCaption() As Boolean
Attribute CardCaption.VB_Description = "Determines whether the column's value will be displayed in the card caption."
    CardCaption = m_bCardCaption
End Property

Public Property Let CardCaption(ByVal bValue As Boolean)
    m_bCardCaption = bValue
End Property

Public Property Get EditType() As jgexEditTypeConstants
Attribute EditType.VB_Description = "Returns/sets a value that indicates how column is edited by the user."
    EditType = m_eEditType
End Property

Public Property Let EditType(ByVal eValue As jgexEditTypeConstants)
    m_eEditType = eValue
End Property

Public Property Get HasValueList() As Boolean
Attribute HasValueList.VB_Description = "Controls whether the column has a list of values."
    HasValueList = m_bHasValueList
End Property

Public Property Let HasValueList(ByVal bValue As Boolean)
    m_bHasValueList = bValue
End Property

Public Property Get SortOrder() As jgexSortOrderConstants
Attribute SortOrder.VB_Description = "Returns a value that represents the sort applied to a column."
    SortOrder = m_eSortOrder
End Property

Public Property Get SortType() As jgexSortTypeConstants
Attribute SortType.VB_Description = "Determines the way that values in the column are sorted and/or grouped."
    SortType = m_eSortType
End Property

Public Property Let SortType(ByVal eValue As jgexSortTypeConstants)
    m_eSortType = eValue
End Property

Public Property Get ShowCaptionInCardView() As Boolean
Attribute ShowCaptionInCardView.VB_Description = "Controls whether the caption of a column is displayed before its value in the card body."
    ShowCaptionInCardView = m_bShowCaptionInCardView
End Property

Public Property Let ShowCaptionInCardView(ByVal bValue As Boolean)
    m_bShowCaptionInCardView = bValue
End Property

Public Property Get IsGrouped() As Boolean
Attribute IsGrouped.VB_Description = "Returns a value that indicates whether the Column is grouped."
    IsGrouped = m_bIsGrouped
End Property

Public Property Get Format() As String
Attribute Format.VB_Description = "Returns/sets a value indicating the format string for the Column."
    Format = m_sFormat
End Property

Public Property Let Format(ByVal sValue As String)
    m_sFormat = sValue
End Property

Public Property Get TextAlignment() As jgexAlignmentConstants
Attribute TextAlignment.VB_Description = "Returns/sets a value that determines the alignment of text in the column."
    TextAlignment = m_eTextAlignment
End Property

Public Property Let TextAlignment(ByVal eValue As jgexAlignmentConstants)
    m_eTextAlignment = eValue
End Property

Public Property Get GroupFormat() As String
Attribute GroupFormat.VB_Description = "Returns/sets a value indicating the format string for the value of the group rows when the control is grouped by the column."
    GroupFormat = m_sGroupFormat
End Property

Public Property Let GroupFormat(ByVal sValue As String)
    m_sGroupFormat = sValue
End Property

Public Property Get GroupPrefix() As String
Attribute GroupPrefix.VB_Description = "Returns/sets the string displayed in a group row before the group value."
    GroupPrefix = m_sGroupPrefix
End Property

Public Property Let GroupPrefix(ByVal sValue As String)
    m_sGroupPrefix = sValue
End Property

Public Property Get GroupEmptyStringCaption() As String
Attribute GroupEmptyStringCaption.VB_Description = "Returns/sets the string displayed in a group row when the value grouped is an empty string."
    GroupEmptyStringCaption = m_sGroupEmptyStringCaption
End Property

Public Property Let GroupEmptyStringCaption(ByVal sValue As String)
    m_sGroupEmptyStringCaption = sValue
End Property

Public Property Get CardIcon() As Boolean
Attribute CardIcon.VB_Description = "Determines whether the column's icon will be displayed in the card caption."
    CardIcon = m_bCardIcon
End Property

Public Property Let CardIcon(ByVal bValue As Boolean)
    m_bCardIcon = bValue
End Property

Public Property Get FetchIcon() As Boolean
Attribute FetchIcon.VB_Description = "Returns/sets whether the FetchIcon event will be fired for the column."
    FetchIcon = m_bFetchIcon
End Property

Public Property Let FetchIcon(ByVal bValue As Boolean)
    m_bFetchIcon = bValue
End Property

Public Property Get FetchData() As Boolean
Attribute FetchData.VB_Description = "Returns/sets whether the FetchData event will be fired for the column."
    FetchData = m_bFetchData
End Property

Public Property Let FetchData(ByVal bValue As Boolean)
    m_bFetchData = bValue
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Returns/sets an expression that stores any extra data needed for your program. "
    Tag = m_sTag
End Property

Public Property Let Tag(ByVal sValue As String)
    m_sTag = sValue
End Property

Public Property Get DataChanged() As Boolean
Attribute DataChanged.VB_Description = "Returns/sets a value indicating that data in the column has changed by some process other than by retrieving data from the current record."
    DataChanged = m_bDataChanged
End Property

Public Property Let DataChanged(ByVal bValue As Boolean)
    m_bDataChanged = bValue
End Property

Public Property Get ReplaceValues() As Boolean
Attribute ReplaceValues.VB_Description = "Returns/sets whether the column values are replaced for using the ValueList for the column."
    ReplaceValues = m_bReplaceValues
End Property

Public Property Let ReplaceValues(ByVal bValue As Boolean)
    m_bReplaceValues = bValue
End Property

Public Property Get DefaultValue() As Variant
Attribute DefaultValue.VB_Description = "Returns/sets the default value for the column in a new record."
    DefaultValue = m_vDefaultValue
End Property

Public Property Let DefaultValue(ByVal vntValue As Variant)
    m_vDefaultValue = vntValue
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a field."
    MaxLength = m_lMaxLength
End Property

Public Property Let MaxLength(ByVal lValue As Long)
    m_lMaxLength = lValue
End Property

Public Property Get NullBehavior() As jgexNullBehaviorConstants
Attribute NullBehavior.VB_Description = "Determines how empty strings are written in the database."
    NullBehavior = m_eNullBehavior
End Property

Public Property Let NullBehavior(ByVal eValue As jgexNullBehaviorConstants)
    m_eNullBehavior = eValue
End Property

Public Property Get ButtonStyle() As jgexButtonStyleConstants
Attribute ButtonStyle.VB_Description = "Determines if a column will show a button when the user enters edit mode."
    ButtonStyle = m_eButtonStyle
End Property

Public Property Let ButtonStyle(ByVal eValue As jgexButtonStyleConstants)
    m_eButtonStyle = eValue
End Property

Public Property Get Selectable() As Boolean
Attribute Selectable.VB_Description = "Controls whether cell can be selected."
    Selectable = m_bSelectable
End Property

Public Property Let Selectable(ByVal bValue As Boolean)
    m_bSelectable = bValue
End Property

Public Property Get CellStyle() As String
Attribute CellStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in all the cells in a column."
    CellStyle = m_sCellStyle
End Property

Public Property Let CellStyle(ByVal sValue As String)
    m_sCellStyle = sValue
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "determines whether text in a cell is wordwrapped."
    WordWrap = m_bWordWrap
End Property

Public Property Let WordWrap(ByVal bValue As Boolean)
    m_bWordWrap = bValue
End Property

Public Property Get AggregateFunction() As jgexAggregateFunctionConstants
Attribute AggregateFunction.VB_Description = "Returns/sets the aggregate function to be shown in group footers for a column."
    AggregateFunction = m_eAggregateFunction
End Property

Public Property Let AggregateFunction(ByVal eValue As jgexAggregateFunctionConstants)
    m_eAggregateFunction = eValue
End Property

Public Property Get MinRowsInCardView() As Long
Attribute MinRowsInCardView.VB_Description = "Returns or sets the minimum number of rows displayed in a word wrap column."
    MinRowsInCardView = m_lMinRowsInCardView
End Property

Public Property Let MinRowsInCardView(ByVal lValue As Long)
    m_lMinRowsInCardView = lValue
End Property

Public Property Get MaxRowsInCardView() As Long
Attribute MaxRowsInCardView.VB_Description = "Returns or sets the maximum number of rows displayed in a word wrap column."
    MaxRowsInCardView = m_lMaxRowsInCardView
End Property

Public Property Let MaxRowsInCardView(ByVal lValue As Long)
    m_lMaxRowsInCardView = lValue
End Property

Public Property Get DropDownControl() As Object
Attribute DropDownControl.VB_Description = "Returns/sets the GridEX control that acts as a drop down list for the column."
    Set DropDownControl = m_oDropDownControl
End Property

Public Property Set DropDownControl(ByVal oValue As Object)
    Set m_oDropDownControl = oValue
End Property

Public Property Get HeaderToolTip() As String
Attribute HeaderToolTip.VB_Description = "Returns/sets the text displayed in a tool tip when users moves the mouse pointer over a column header."
    HeaderToolTip = m_sHeaderToolTip
End Property

Public Property Let HeaderToolTip(ByVal sValue As String)
    m_sHeaderToolTip = sValue
End Property

Public Property Get HeaderStyle() As String
Attribute HeaderStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in the column header."
    HeaderStyle = m_sHeaderStyle
End Property

Public Property Let HeaderStyle(ByVal sValue As String)
    m_sHeaderStyle = sValue
End Property

Public Property Get TotalRowFormat() As String
Attribute TotalRowFormat.VB_Description = "Returns/sets the format string for the aggregate function result of a column in the group footer row."
    TotalRowFormat = m_sTotalRowFormat
End Property

Public Property Let TotalRowFormat(ByVal sValue As String)
    m_sTotalRowFormat = sValue
End Property

Public Property Get TotalRowPrefix() As String
Attribute TotalRowPrefix.VB_Description = "Returns/sets the string displayed in a group footer row before the aggregate function result."
    TotalRowPrefix = m_sTotalRowPrefix
End Property

Public Property Let TotalRowPrefix(ByVal sValue As String)
    m_sTotalRowPrefix = sValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub AutoSize()
Attribute AutoSize.VB_Description = "Adjusts the width of a column to accommodate the longest text."
End Sub

Friend Sub frInit(ByVal nIndex As Integer, sCaption As String, ByVal eColumnType As jgexColumnTypeConstants, ByVal eEditType As jgexEditTypeConstants, sKey As String)
    m_nIndex = nIndex
    m_nColPosition = nIndex
    m_sCaption = sCaption
    m_eColumnType = eColumnType
    m_eEditType = eEditType
    m_sKey = sKey
End Sub

Friend Sub frSetIndex(ByVal nIndex As Integer)
    m_nIndex = nIndex
End Sub

Friend Sub frSetSortOrder(ByVal eValue As jgexSortOrderConstants)
    m_eSortOrder = eValue
End Sub

Friend Sub frSetIsGrouped(ByVal bValue As Boolean)
    m_bIsGrouped = bValue
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_bVisible = True
    m_bAllowSizing = True
    m_bSelectable = True
    m_bCardCaption = False
    m_bShowCaptionInCardView = True
    m_eEditType = jgexEditTextBox
    m_eSortType = jgexSortTypeString
    m_lWidth = 1000
    m_nHeaderIcon = -1
    m_nDefaultIcon = -1
    m_lMinRowsInCardView = 1
    m_lMaxRowsInCardView = 1
End Sub
