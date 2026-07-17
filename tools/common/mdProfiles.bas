Attribute VB_Name = "mdProfiles"
'=========================================================================
'
' Open GridEX 2000 Control
' Property profiles for the snapshot engine
'
' GENERATED FILE -- DO NOT EDIT. Re-generate with tools\GenProfiles.ps1
'
' Skipped parameterized properties (handled in mdSnapshot.bas):
'   JSColumns.Item(ByVal vntIndexKey As Variant)
'   JSDataObjectFiles.Item(ByVal Index As Long)
'   JSFmtConditions.Item(ByVal Index As Variant)
'   JSFormatStyles.Item(ByVal Index As Variant)
'   JSGridImages.Item(ByVal Index As Integer)
'   JSGroups.Item(ByVal Index As Integer)
'   JSPrinterProperties.HeaderString(ByVal Position As jgexHeaderFooterPositionConstants)
'   JSPrinterProperties.FooterString(ByVal Position As jgexHeaderFooterPositionConstants)
'   JSRowData.Value(ByVal ColIndex As Integer)
'   JSRowData.IconIndex(ByVal ColIndex As Integer)
'   JSRowData.DisplayValue(ByVal ColIndex As Integer)
'   JSRowData.CellStyle(ByVal ColIndex As Integer)
'   JSRowData.CellPicture(ByVal ColIndex As Integer)
'   JSSelectedItems.Item(ByVal Index As Long)
'   JSSortKeys.Item(ByVal Index As Integer)
'   JSValueList.Item(ByVal Index As Long)
'   JSValueList.ItemByValue(ByVal Value As Variant)
'   GridEX.RowSelected(ByVal RowPosition As Long)
'   GridEX.RowBookmark(ByVal RowIndex As Long)
'   GridEX.Value(ByVal ColIndex As Integer)
'   GridEX.RowExpanded(ByVal RowPosition As Long)
'
'=========================================================================
Option Explicit

Public Enum UcsPropKindEnum
    ucsPrkScalar
    ucsPrkEnum
    ucsPrkVariant
    ucsPrkFont
    ucsPrkPicture
    ucsPrkObject
    ucsPrkClass
    ucsPrkCollection
End Enum

Public Type UcsProfileProp
    sClass                          As String
    sProp                           As String
    eKind                           As UcsPropKindEnum
    sTypeName                       As String
    sItemClass                      As String
    bCanWrite                       As Boolean
    bRuntimeOnly                    As Boolean
End Type

Private m_uProps()                  As UcsProfileProp
Private m_lCount                    As Long

Public Sub ProfilesInit()
    If m_lCount > 0 Then
        Exit Sub
    End If
    ReDim m_uProps(0 To 266) As UcsProfileProp
    pvAdd "JSColumn", "Key", ucsPrkScalar, "String", "", False, False
    pvAdd "JSColumn", "HeaderAlignment", ucsPrkEnum, "jgexAlignmentConstants", "", True, False
    pvAdd "JSColumn", "HeaderIcon", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSColumn", "DefaultIcon", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSColumn", "Index", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSColumn", "ColPosition", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSColumn", "ValueList", ucsPrkCollection, "JSValueList", "JSValueItem", False, False
    pvAdd "JSColumn", "Visible", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "Width", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSColumn", "DataField", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "Caption", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "AllowSizing", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "ColumnType", ucsPrkEnum, "jgexColumnTypeConstants", "", True, False
    pvAdd "JSColumn", "CardCaption", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "EditType", ucsPrkEnum, "jgexEditTypeConstants", "", True, False
    pvAdd "JSColumn", "HasValueList", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "SortOrder", ucsPrkEnum, "jgexSortOrderConstants", "", False, True
    pvAdd "JSColumn", "SortType", ucsPrkEnum, "jgexSortTypeConstants", "", True, False
    pvAdd "JSColumn", "ShowCaptionInCardView", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "IsGrouped", ucsPrkScalar, "Boolean", "", False, True
    pvAdd "JSColumn", "Format", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "TextAlignment", ucsPrkEnum, "jgexAlignmentConstants", "", True, False
    pvAdd "JSColumn", "GroupFormat", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "GroupPrefix", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "GroupEmptyStringCaption", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "CardIcon", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "FetchIcon", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "FetchData", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "Tag", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "DataChanged", ucsPrkScalar, "Boolean", "", True, True
    pvAdd "JSColumn", "ReplaceValues", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "DefaultValue", ucsPrkVariant, "Variant", "", True, False
    pvAdd "JSColumn", "MaxLength", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSColumn", "NullBehavior", ucsPrkEnum, "jgexNullBehaviorConstants", "", True, False
    pvAdd "JSColumn", "ButtonStyle", ucsPrkEnum, "jgexButtonStyleConstants", "", True, False
    pvAdd "JSColumn", "Selectable", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "CellStyle", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "WordWrap", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSColumn", "AggregateFunction", ucsPrkEnum, "jgexAggregateFunctionConstants", "", True, False
    pvAdd "JSColumn", "MinRowsInCardView", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSColumn", "MaxRowsInCardView", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSColumn", "DropDownControl", ucsPrkObject, "Object", "", True, True
    pvAdd "JSColumn", "HeaderToolTip", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "HeaderStyle", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "TotalRowFormat", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumn", "TotalRowPrefix", ucsPrkScalar, "String", "", True, False
    pvAdd "JSColumns", "Count", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSDataObject", "Files", ucsPrkCollection, "JSDataObjectFiles", "String", False, True
    pvAdd "JSDataObjectFiles", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSFmtCondition", "FormatStyle", ucsPrkClass, "JSFormatStyle", "", False, False
    pvAdd "JSFmtCondition", "ColIndex", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSFmtCondition", "Operator", ucsPrkEnum, "jgexConditionOperatorConstants", "", True, False
    pvAdd "JSFmtCondition", "Value1", ucsPrkVariant, "Variant", "", True, False
    pvAdd "JSFmtCondition", "Value2", ucsPrkVariant, "Variant", "", True, False
    pvAdd "JSFmtCondition", "Index", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSFmtCondition", "Key", ucsPrkScalar, "String", "", False, False
    pvAdd "JSFmtConditions", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSFmtConditions", "GroupCondition", ucsPrkClass, "JSFmtCondition", "", False, False
    pvAdd "JSFmtConditions", "ApplyGroupCondition", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFmtConditions", "GroupConditionCountTitle", ucsPrkScalar, "String", "", True, False
    pvAdd "JSFmtConditions", "ShowGroupConditionCount", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFormatStyle", "Name", ucsPrkScalar, "String", "", False, False
    pvAdd "JSFormatStyle", "BackColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "JSFormatStyle", "ForeColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "JSFormatStyle", "FontBold", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFormatStyle", "FontItalic", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFormatStyle", "FontUnderline", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFormatStyle", "FontStrikeThru", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSFormatStyle", "FontSize", ucsPrkScalar, "Currency", "", True, False
    pvAdd "JSFormatStyle", "FontName", ucsPrkScalar, "String", "", True, False
    pvAdd "JSFormatStyle", "FontCharset", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSFormatStyle", "TextAlignment", ucsPrkEnum, "jgexAlignmentConstants", "", True, False
    pvAdd "JSFormatStyle", "BackgroundPicture", ucsPrkPicture, "Picture", "", True, False
    pvAdd "JSFormatStyle", "Picture", ucsPrkPicture, "Picture", "", True, False
    pvAdd "JSFormatStyle", "DrawModeBackGroundPicture", ucsPrkEnum, "jgexDrawModePictureBackgroundConstants", "", True, False
    pvAdd "JSFormatStyle", "PictureHorzAlignment", ucsPrkEnum, "jgexHorzPictureAlignmentConstants", "", True, False
    pvAdd "JSFormatStyle", "PictureVertAlignment", ucsPrkEnum, "jgexVertPictureAlignmentConstants", "", True, False
    pvAdd "JSFormatStyle", "PictureDrawMode", ucsPrkEnum, "jgexPictureDrawModeConstants", "", True, False
    pvAdd "JSFormatStyles", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSGridImage", "Picture", ucsPrkPicture, "Picture", "", False, False
    pvAdd "JSGridImage", "Index", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSGridImages", "Count", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSGridImages", "hImageList", ucsPrkScalar, "Long", "", True, True
    pvAdd "JSGroup", "Index", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSGroup", "SortOrder", ucsPrkEnum, "jgexSortOrderConstants", "", True, False
    pvAdd "JSGroup", "ColIndex", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSGroups", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSPrinterProperties", "Copies", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSPrinterProperties", "Collate", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "PrintQuality", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSPrinterProperties", "ColorMode", ucsPrkEnum, "jgexPPColorModeConstants", "", True, False
    pvAdd "JSPrinterProperties", "DeviceName", ucsPrkScalar, "String", "", False, True
    pvAdd "JSPrinterProperties", "DriverName", ucsPrkScalar, "String", "", False, True
    pvAdd "JSPrinterProperties", "Orientation", ucsPrkEnum, "jgexPPOrientationConstants", "", True, False
    pvAdd "JSPrinterProperties", "PaperBin", ucsPrkEnum, "jgexPPPaperBinConstants", "", True, False
    pvAdd "JSPrinterProperties", "PaperSize", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSPrinterProperties", "PaperWidth", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "PaperHeight", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "LeftMargin", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "TopMargin", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "RightMargin", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "BottomMargin", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "ClientWidth", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSPrinterProperties", "ClientHeight", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSPrinterProperties", "RepeatHeaders", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "FitColumns", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "TranslateColors", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "HeaderDistance", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "FooterDistance", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "DocumentName", ucsPrkScalar, "String", "", True, False
    pvAdd "JSPrinterProperties", "RepeatFrozenCols", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "PrintPreviewRows", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "CardColumnsPerPage", ucsPrkScalar, "Long", "", True, False
    pvAdd "JSPrinterProperties", "PageHeaderFont", ucsPrkFont, "Font", "", True, False
    pvAdd "JSPrinterProperties", "PageFooterFont", ucsPrkFont, "Font", "", True, False
    pvAdd "JSPrinterProperties", "PrintProgressDialog", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "TransparentBackground", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSPrinterProperties", "MeasurementUnits", ucsPrkEnum, "jgexMeasurementUnitsConstants", "", True, False
    pvAdd "JSPrinterProperties", "PageSetupCanceled", ucsPrkScalar, "Boolean", "", False, True
    pvAdd "JSRetBoolean", "Value", ucsPrkScalar, "Boolean", "", True, True
    pvAdd "JSRetInteger", "Value", ucsPrkScalar, "Integer", "", True, True
    pvAdd "JSRetVariant", "Value", ucsPrkVariant, "Variant", "", True, True
    pvAdd "JSRowData", "Bookmark", ucsPrkVariant, "Variant", "", False, True
    pvAdd "JSRowData", "ColCount", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSRowData", "GroupLevel", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSRowData", "RowIndex", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSRowData", "RowType", ucsPrkEnum, "jgexRowTypeConstants", "", False, True
    pvAdd "JSRowData", "RowStyle", ucsPrkScalar, "String", "", True, True
    pvAdd "JSRowData", "RecordCount", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSRowData", "RowHeight", ucsPrkScalar, "Long", "", True, True
    pvAdd "JSRowData", "GroupCaption", ucsPrkScalar, "String", "", True, True
    pvAdd "JSRowData", "GroupIconIndex", ucsPrkScalar, "Integer", "", True, True
    pvAdd "JSRowData", "PreviewRowVisible", ucsPrkScalar, "Boolean", "", True, True
    pvAdd "JSSelectedItem", "Bookmark", ucsPrkVariant, "Variant", "", False, True
    pvAdd "JSSelectedItem", "RowIndex", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSSelectedItem", "RowPosition", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSSelectedItem", "RowType", ucsPrkEnum, "jgexRowTypeConstants", "", False, True
    pvAdd "JSSelectedItems", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSSortKey", "Index", ucsPrkScalar, "Integer", "", False, True
    pvAdd "JSSortKey", "ColIndex", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSSortKey", "SortOrder", ucsPrkEnum, "jgexSortOrderConstants", "", True, False
    pvAdd "JSSortKeys", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSValueItem", "Index", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSValueItem", "Text", ucsPrkScalar, "String", "", True, False
    pvAdd "JSValueItem", "Value", ucsPrkVariant, "Variant", "", True, False
    pvAdd "JSValueItem", "IconIndex", ucsPrkScalar, "Integer", "", True, False
    pvAdd "JSValueItem", "Visible", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "JSValueList", "Count", ucsPrkScalar, "Long", "", False, True
    pvAdd "JSValueList", "VisibleCount", ucsPrkScalar, "Long", "", False, True
    pvAdd "GridEX", "FrozenColumns", ucsPrkScalar, "Integer", "", True, False
    pvAdd "GridEX", "RowHeight", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "hWndEdit", ucsPrkScalar, "Long", "", False, True
    pvAdd "GridEX", "OLEDropMode", ucsPrkEnum, "jgexOleDropModeConstants", "", True, False
    pvAdd "GridEX", "ADORecordset", ucsPrkObject, "Object", "", True, True
    pvAdd "GridEX", "FmtConditions", ucsPrkCollection, "JSFmtConditions", "JSFmtCondition", False, False
    pvAdd "GridEX", "ForeColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "RowColorEven", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "RowColorOdd", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "RowCount", ucsPrkScalar, "Long", "", False, True
    pvAdd "GridEX", "SortKeys", ucsPrkCollection, "JSSortKeys", "JSSortKey", False, False
    pvAdd "GridEX", "Col", ucsPrkScalar, "Integer", "", True, True
    pvAdd "GridEX", "BackColorRowGroup", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "Groups", ucsPrkCollection, "JSGroups", "JSGroup", False, False
    pvAdd "GridEX", "RecordNavigator", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "BorderStyle", ucsPrkEnum, "jgexBorderStyleConstants", "", True, False
    pvAdd "GridEX", "Columns", ucsPrkCollection, "JSColumns", "JSColumn", False, False
    pvAdd "GridEX", "GroupByBoxVisible", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "BackColorGBBox", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "BackColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "ColumnHeaderHeight", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "ColumnHeaders", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "BackColorHeader", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "MaskColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "DefaultColumnWidth", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "View", ucsPrkEnum, "jgexViewConstants", "", True, False
    pvAdd "GridEX", "ImageWidth", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "ImageHeight", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "GridImages", ucsPrkCollection, "JSGridImages", "JSGridImage", False, False
    pvAdd "GridEX", "RowHeaders", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "AllowAddNew", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "NewRowPos", ucsPrkEnum, "jgexNewRowPosConstants", "", True, False
    pvAdd "GridEX", "ItemCount", ucsPrkScalar, "Long", "", True, True
    pvAdd "GridEX", "DataMode", ucsPrkEnum, "jgexDataModeConstants", "", True, False
    pvAdd "GridEX", "LeftCol", ucsPrkScalar, "Integer", "", True, True
    pvAdd "GridEX", "ColumnHeaderFont", ucsPrkFont, "Font", "", True, False
    pvAdd "GridEX", "Font", ucsPrkFont, "Font", "", True, False
    pvAdd "GridEX", "FirstItem", ucsPrkScalar, "Long", "", True, True
    pvAdd "GridEX", "GridLines", ucsPrkEnum, "jgexGridLinesConstants", "", True, False
    pvAdd "GridEX", "GridLinesColor", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "BackColorBkg", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "CardSpacing", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "CardWidth", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "Row", ucsPrkScalar, "Long", "", True, True
    pvAdd "GridEX", "ErrorText", ucsPrkScalar, "String", "", True, True
    pvAdd "GridEX", "AllowEdit", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "DataChanged", ucsPrkScalar, "Boolean", "", True, True
    pvAdd "GridEX", "hWnd", ucsPrkScalar, "Long", "", False, True
    pvAdd "GridEX", "AllowDelete", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "GroupByBoxInfoText", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "ForeColorHeader", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "ForeColorRowGroup", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "BackColorInfoText", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "ForeColorInfoText", ucsPrkScalar, "OLE_COLOR", "", True, False
    pvAdd "GridEX", "AutomaticArrange", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "Connect", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "DatabaseName", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "Exclusive", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "RecordSource", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "RecordsetType", ucsPrkEnum, "jgexRecordsetTypeConstants", "", True, False
    pvAdd "GridEX", "AllowColumnDrag", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "SelectionStyle", ucsPrkEnum, "jgexSelectionStyleConstants", "", True, False
    pvAdd "GridEX", "AllowCardSizing", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "CardCaptionPrefix", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "CardBorders", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "ContinuousScroll", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "Recordset", ucsPrkObject, "Object", "", True, True
    pvAdd "GridEX", "FullyLoaded", ucsPrkScalar, "Boolean", "", False, True
    pvAdd "GridEX", "LockType", ucsPrkEnum, "jgexLockTypeConstants", "", True, False
    pvAdd "GridEX", "Options", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "ReadOnly", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "UseEvenOddColor", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "HeaderStyle", ucsPrkEnum, "jgexHeaderStyleConstants", "", True, False
    pvAdd "GridEX", "HideSelection", ucsPrkEnum, "jgexHideSelectionConstants", "", True, False
    pvAdd "GridEX", "MultiSelect", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "SelectedItems", ucsPrkCollection, "JSSelectedItems", "JSSelectedItem", False, True
    pvAdd "GridEX", "DetectRowDrag", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "CalendarTodayText", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "CalendarNoneText", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "EditMode", ucsPrkEnum, "jgexEditModeConstants", "", True, True
    pvAdd "GridEX", "SelStart", ucsPrkScalar, "Long", "", True, True
    pvAdd "GridEX", "SelLength", ucsPrkScalar, "Long", "", True, True
    pvAdd "GridEX", "SelText", ucsPrkScalar, "String", "", True, True
    pvAdd "GridEX", "CursorLocation", ucsPrkEnum, "jgexCursorLocationConstants", "", True, False
    pvAdd "GridEX", "TabKeyBehavior", ucsPrkEnum, "jgexTabKeyBehaviorConstants", "", True, False
    pvAdd "GridEX", "Enabled", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "ColumnAutoResize", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "FormatStyles", ucsPrkCollection, "JSFormatStyles", "JSFormatStyle", False, False
    pvAdd "GridEX", "PrinterProperties", ucsPrkClass, "JSPrinterProperties", "", False, False
    pvAdd "GridEX", "PreviewRowLines", ucsPrkScalar, "Integer", "", True, False
    pvAdd "GridEX", "PreviewColumn", ucsPrkVariant, "Variant", "", True, False
    pvAdd "GridEX", "GroupFooterStyle", ucsPrkEnum, "jgexGroupFooterStyleConstants", "", True, False
    pvAdd "GridEX", "ShowEmptyFields", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "ActAsDropDown", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "BoundColumnIndex", ucsPrkVariant, "Variant", "", True, False
    pvAdd "GridEX", "ReplaceColumnIndex", ucsPrkVariant, "Variant", "", True, False
    pvAdd "GridEX", "GridLineStyle", ucsPrkEnum, "jgexGridLineStyleConstants", "", True, False
    pvAdd "GridEX", "EmptyRows", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "Redraw", ucsPrkScalar, "Boolean", "", True, True
    pvAdd "GridEX", "DefaultGroupMode", ucsPrkEnum, "jgexDefaultGroupModeConstants", "", True, False
    pvAdd "GridEX", "HoldSortSettings", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "RecordNavigatorString", ucsPrkScalar, "String", "", True, False
    pvAdd "GridEX", "ShowToolTips", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "ScrollToolTips", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "ScrollToolTipColumn", ucsPrkVariant, "Variant", "", True, False
    pvAdd "GridEX", "AutomaticSort", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GridEX", "PreviewRowIndent", ucsPrkScalar, "Long", "", True, False
    pvAdd "GridEX", "AllowRowSizing", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GEXPreview", "TotalPages", ucsPrkScalar, "Long", "", False, True
    pvAdd "GEXPreview", "CurrentPage", ucsPrkScalar, "Long", "", True, True
    pvAdd "GEXPreview", "ToolbarFont", ucsPrkFont, "Font", "", True, False
    pvAdd "GEXPreview", "Zoom", ucsPrkEnum, "jgexZoomConstants", "", True, False
    pvAdd "GEXPreview", "ToolbarVisible", ucsPrkScalar, "Boolean", "", True, False
    pvAdd "GEXPreview", "PageSetupText", ucsPrkScalar, "String", "", True, False
    pvAdd "GEXPreview", "PrintText", ucsPrkScalar, "String", "", True, False
    pvAdd "GEXPreview", "CloseButtonText", ucsPrkScalar, "String", "", True, False
    pvAdd "GEXPreview", "hWnd", ucsPrkScalar, "Long", "", False, True
    pvAdd "GEXPreview", "BackColor", ucsPrkScalar, "OLE_COLOR", "", True, False
End Sub

Public Function ProfileForClass(sClass As String, uProps() As UcsProfileProp) As Long
    Dim lIdx            As Long

    ProfilesInit
    ReDim uProps(0 To m_lCount - 1) As UcsProfileProp
    For lIdx = 0 To m_lCount - 1
        If m_uProps(lIdx).sClass = sClass Then
            uProps(ProfileForClass) = m_uProps(lIdx)
            ProfileForClass = ProfileForClass + 1
        End If
    Next
End Function

Private Sub pvAdd(sClass As String, sProp As String, ByVal eKind As UcsPropKindEnum, sTypeName As String, sItemClass As String, ByVal bCanWrite As Boolean, ByVal bRuntimeOnly As Boolean)
    With m_uProps(m_lCount)
        .sClass = sClass
        .sProp = sProp
        .eKind = eKind
        .sTypeName = sTypeName
        .sItemClass = sItemClass
        .bCanWrite = bCanWrite
        .bRuntimeOnly = bRuntimeOnly
    End With
    m_lCount = m_lCount + 1
End Sub
