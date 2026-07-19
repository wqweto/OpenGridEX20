VERSION 5.00
Begin VB.UserControl GridEX
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
End
Attribute VB_Name = "GridEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Janus GridEX 2000 Control (DAO 3.6 & ADO 2.x)"
'=========================================================================
'
' Open GridEX 2000 Control
' Main grid control (DAO 3.6 & ADO 2.x)
'
'=========================================================================
Option Explicit
DefObj A-Z

Implements IObjectSafety

'=========================================================================
' Public events
'=========================================================================

Public Event BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As JSRetBoolean)
Attribute BeforeColEdit.VB_Description = "Occurs before the user enters edit mode by typing a character or clicking in a cell."
Public Event AfterColEdit(ByVal ColIndex As Integer)
Attribute AfterColEdit.VB_Description = "Occurs after data in the current cell is edited."
Public Event BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As JSRetBoolean)
Attribute BeforeColUpdate.VB_Description = "Occurs before data is moved from the cell to the control's copy buffer."
Public Event Error(ByVal ErrNumber As Long, ByVal DisplayMessage As JSRetBoolean)
Attribute Error.VB_Description = "Occurs as the result of a data access error."
Public Event BeforeUpdate(ByVal Cancel As JSRetBoolean)
Attribute BeforeUpdate.VB_Description = "Occurs before data is committed from to the database."
Public Event AfterColUpdate(ByVal ColIndex As Integer)
Attribute AfterColUpdate.VB_Description = "Occurs after data is moved from a cell  to the control's copy buffer."
Public Event AfterUpdate()
Attribute AfterUpdate.VB_Description = "Occurs after changes made by the user have been written to the database."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of the current cell have changed."
Public Event ColumnHeaderClick(ByVal Column As JSColumn)
Attribute ColumnHeaderClick.VB_Description = "Occurs when the user clicks on a column header."
Public Event GroupByBoxHeaderClick(ByVal Group As JSGroup)
Attribute GroupByBoxHeaderClick.VB_Description = "Occurs when the user clicks on a header in the group by box."
Public Event BeforeColMove(ByVal Column As JSColumn, ByVal NewPosition As Integer, ByVal Cancel As JSRetBoolean)
Attribute BeforeColMove.VB_Description = "Occurs before a column position changes."
Public Event AfterColMove()
Attribute AfterColMove.VB_Description = "Occurs after the user has moved a column into a new position."
Public Event BeforeColumnDrag(ByVal Column As JSColumn, ByVal Cancel As JSRetBoolean)
Attribute BeforeColumnDrag.VB_Description = "Occurs before the user begins a drag operation with a header in the column headers row."
Public Event BeforeGroupDrag(ByVal Group As JSGroup, ByVal Cancel As JSRetBoolean)
Attribute BeforeGroupDrag.VB_Description = "Occurs before the user begins a drag operation with a group by box header."
Public Event AfterGroupChange()
Attribute AfterGroupChange.VB_Description = "Occurs after the user has changed, added or removed a group."
Public Event BeforeGroupChange(ByVal Group As JSGroup, ByVal ChangeOperation As jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As JSRetBoolean)
Attribute BeforeGroupChange.VB_Description = "Occurs before a change in the group settings is committed."
Public Event AfterDelete()
Attribute AfterDelete.VB_Description = "Occurs after the user deletes the selected record."
Public Event BeforeDelete(ByVal Cancel As JSRetBoolean)
Attribute BeforeDelete.VB_Description = "Occurs before a delete operation begins."
Public Event FetchData(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal Value As JSRetVariant)
Attribute FetchData.VB_Description = "Occurs when unbound column data is needed for display."
Public Event FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As JSRetInteger)
Attribute FetchIcon.VB_Description = "Fetches the  icons index for cells."
Public Event UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
Attribute UnboundDelete.VB_Description = "Occurs in Unbound mode, whenever a row of data is deleted."
Public Event UnboundAddNew(ByVal NewRowBookmark As JSRetVariant, ByVal Values As JSRowData)
Attribute UnboundAddNew.VB_Description = "Occurs in an unbound GridEX control when a new row is added to it."
Public Event UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
Attribute UnboundUpdate.VB_Description = "Occurs when an unbound GridEX control has an entire row of modified data to be written in the data set."
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the control."
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button two times over the control."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occur when the user presses a key."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occur when the user releases a key."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases a mouse button."
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Public Event FirstItemChange()
Attribute FirstItemChange.VB_Description = "Occurs when the FirstItem property changes."
Public Event LeftColChange()
Attribute LeftColChange.VB_Description = "Occurs when the LeftCol property changes."
Public Event RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
Attribute RowColChange.VB_Description = "Occurs when the current cell changes to a different cell."
Public Event ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As JSRetBoolean)
Attribute ColResize.VB_Description = "Occurs when a user resizes a column in a GridEX control."
Public Event RowResize(ByVal NewRowHeight As Long, ByVal Cancel As JSRetBoolean)
Attribute RowResize.VB_Description = "Occurs when the user resizes rows."
Public Event CardResize(ByVal NewCardWidth As Long, ByVal Cancel As JSRetBoolean)
Attribute CardResize.VB_Description = "Occurs when the user resizes cards."
Public Event DropList(ByVal ColIndex As Integer)
Attribute DropList.VB_Description = "Occurs before a drop down list appears."
Public Event ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal Value As Variant)
Attribute ListSelected.VB_Description = "Occurs after the user has selected an entry in a drop down list."
Public Event UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
Attribute UnboundReadData.VB_Description = "Occurs whenever an unbound GridEX control requires data for display, sorting or grouping."
Public Event RowDrag(ByVal Button As Integer, ByVal Shift As Integer)
Attribute RowDrag.VB_Description = "Occurs when the user attempts to drag the selected rows."
Public Event InitCustomEdit(ByVal ColIndex As Integer, ByVal EditBackColor As OLE_COLOR, ByVal EditForeColor As OLE_COLOR, ByVal EditFont As Font)
Attribute InitCustomEdit.VB_Description = "Occurs when an edit operation is about to begin."
Public Event ShowCustomEdit(ByVal ColIndex As Integer, ByVal EditLeft As Single, ByVal EditTop As Single, ByVal EditWidth As Single, ByVal EditHeight As Single, ByVal EditVisible As Boolean)
Attribute ShowCustomEdit.VB_Description = "Occurs when a custom edit column is about to hide or show the cell editor."
Public Event EndCustomEdit(ByVal ColIndex As Integer)
Attribute EndCustomEdit.VB_Description = "Occurs when a custom edit operation is ended."
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs when a drag action was either performed or canceled."
Public Event OLEDragDrop(Data As JSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when the source component determines that a drop can occur."
Public Event OLEDragOver(Data As JSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when one component is dragged over another."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As JSDataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs on a source component when a target component requests data that was not provided during the OLEStartDrag event."
Public Event OLEStartDrag(Data As JSDataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when the OLEDrag method is performed and, data formats and drop effects need to be specified."
Public Event BeforeDeleteEX(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Cancel As JSRetBoolean)
Attribute BeforeDeleteEX.VB_Description = "Occurs before a selected record is deleted in a GridEX control."
Public Event ColButtonClick(ByVal ColIndex As Integer)
Attribute ColButtonClick.VB_Description = "Occurs when the user presses a button in a cell."
Public Event RowFormat(RowBuffer As JSRowData)
Attribute RowFormat.VB_Description = "Occurs when The GridEX control loads a row."
Public Event BeforePrinting(ByVal nPages As Long, ByVal Cancel As JSRetBoolean)
Attribute BeforePrinting.VB_Description = "Occurs before the GridEX control prints its contents."
Public Event BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
Attribute BeforePrintPage.VB_Description = "Occurs before the GridEX prints a page."
Public Event PrintingProgress(ByVal PrintProgress As Single, ByVal Cancel As JSRetBoolean)
Attribute PrintingProgress.VB_Description = "Occurs continuously while the GridEX control is printing a document."
Public Event NotInList(ByVal ColIndex As Integer, ByVal NewData As String, ByVal CancelUpdate As JSRetBoolean)
Attribute NotInList.VB_Description = "Occurs when the text entered by the user in a combo column is not present in the list."
Public Event SelectionChange()
Attribute SelectionChange.VB_Description = "Occurs when the user selects a different range of rows."

'=========================================================================
' Public enums
'=========================================================================

Public Enum jgexButtonStyleConstants
    jgexNoButton = 0
    jgexButtonEllipsis = 1
    jgexButtonDownArrow = 2
End Enum

Public Enum jgexColumnTypeConstants
    jgexText = 0
    jgexIcon = 1
    jgexIconAndText = 2
    jgexCheckBox = 3
End Enum

Public Enum jgexEditTypeConstants
    jgexEditCustom = -1
    jgexEditNone = 0
    jgexEditTextBox = 1
    jgexEditCheckBox = 2
    jgexEditDropDown = 3
    jgexEditCalendarDropDown = 4
    jgexEditCombo = 5
End Enum

Public Enum jgexNullBehaviorConstants
    jgexNBAutomatic = 0
    jgexNBNull = 1
    jgexNBEmptyString = 2
End Enum

Public Enum jgexHorzPictureAlignmentConstants
    jgexHPALeft = 0
    jgexHPACenter = 1
    jgexHPARight = 2
    jgexHPALeftOfText = 3
    jgexHPARightOfText = 4
End Enum

Public Enum jgexVertPictureAlignmentConstants
    jgexVPATop = 0
    jgexVPACenter = 1
    jgexVPABottom = 2
    jgexVPATopOfText = 3
    jgexVPABottomOfText = 4
End Enum

Public Enum jgexPictureDrawModeConstants
    jgexTextOnly = 0
    jgexPictureOnly = 1
    jgexPictureAndText = 2
End Enum

Public Enum jgexPaperSizeConstants
    jgexPSLetter = 1
    jgexPSLetterSmall = 2
    jgexPSTabloid = 3
    jgexPSLedger = 4
    jgexPSLegal = 5
    jgexPSStatement = 6
    jgexPSExecutive = 7
    jgexPSA3 = 8
    jgexPSA4 = 9
    jgexPSA4Small = 10
    jgexPSA5 = 11
    jgexPSB4 = 12
    jgexPSB5 = 13
    jgexPSFolio = 14
    jgexPSQuarto = 15
    jgexPS10x14 = 16
    jgexPS11x17 = 17
    jgexPSNote = 18
    jgexPSEnv9 = 19
    jgexPSEnv10 = 20
    jgexPSEnv11 = 21
    jgexPSEnv12 = 22
    jgexPSEnv14 = 23
    jgexPSCSheet = 24
    jgexPSDSheet = 25
    jgexPSESheet = 26
    jgexPSEnvDL = 27
    jgexPSEnvC5 = 28
    jgexPSEnvC3 = 29
    jgexPSEnvC4 = 30
    jgexPSEnvC6 = 31
    jgexPSEnvC65 = 32
    jgexPSEnvB4 = 33
    jgexPSEnvB5 = 34
    jgexPSEnvB6 = 35
    jgexPSEnvItaly = 36
    jgexPSEnvMonarch = 37
    jgexPSEnvPersonal = 38
    jgexPSFanfoldUS = 39
    jgexPSFanfoldStdGerman = 40
    jgexPSFanfoldLglGerman = 41
    jgexPSUser = 256
End Enum

Public Enum jgexMeasurementUnitsConstants
    jgexMUInches = 0
    jgexMUMilimeters = 1
End Enum

Public Enum jgexHeaderFooterPositionConstants
    jgexHFCenter = 1
    jgexHFLeft = 2
    jgexHFRight = 3
End Enum

Public Enum jgexPPPrintQualityConstants
    jgexPPDraft = -1
    jgexPPLow = -2
    jgexPPMedium = -3
    jgexPPHigh = -4
End Enum

Public Enum jgexPPColorModeConstants
    jgexPPCMMonochrome = 1
    jgexPPCMColor = 2
End Enum

Public Enum jgexPPOrientationConstants
    jgexPPPortrait = 1
    jgexPPLandscape = 2
End Enum

Public Enum jgexPPPaperBinConstants
    jgexPPUpper = 1
    jgexPPLower = 2
    jgexPPMiddle = 3
    jgexPPManual = 4
    jgexPPEnvelope = 5
    jgexPPEnvManual = 6
    jgexPPAuto = 7
    jgexPPTractor = 8
    jgexPPSmallFmt = 9
    jgexPPLargeFmt = 10
    jgexPPLargeCapacity = 11
    jgexPPCassete = 14
End Enum

Public Enum jgexZoomConstants
    jgexZoomCurrentSize = 0
    jgexZoomOnePage = 1
    jgexZoomTwoPages = 2
End Enum

Public Enum jgexGridLinesConstants
    jgexGLBoth = -1
    jgexGLNone = 0
    jgexGLVertical = 1
    jgexGLHorizontal = 2
End Enum

Public Enum jgexDefaultGroupModeConstants
    jgexDGMExpanded = 0
    jgexDGMCollapsed = 1
End Enum

Public Enum jgexGridLineStyleConstants
    jgexGLSSolid = 0
    jgexGLSDashes = 1
    jgexGLSSmallDots = 2
End Enum

Public Enum jgexDrawModePictureBackgroundConstants
    jgexDMCenter = 0
    jgexDMTile = 1
    jgexDMStretch = 2
End Enum

Public Enum jgexAppearanceConstants
    jgexAppearanceFlat = 0
    jgexAppearanceSingle3DRaised = 1
    jgexAppearanceSingle3DSunken = 2
    jgexAppearance3DRaised = 3
    jgexAppearance3DSunken = 4
End Enum

Public Enum jgexAggregateFunctionConstants
    jgexAggregateNone = 0
    jgexCount = 1
    jgexSum = 2
    jgexAvg = 3
    jgexMin = 4
    jgexMax = 5
    jgexStdDev = 6
    jgexValueCount = 7
End Enum

Public Enum jgexDataFormatConstants
    jgexCFText = 1
    jgexCFBitmap = 2
    jgexCFMetaFile = 3
    jgexCFEMetafile = 14
    jgexCFDIB = 8
    jgexCFPalette = 9
    jgexCFFiles = 15
    jgexCFRTF = &HFFFFBF01
End Enum

Public Enum jgexOLEDropEffectConstants
    jgexDropEffectNone = 0
    jgexDropEffectCopy = 1
    jgexDropEffectMove = 2
    jgexDropEffectScroll = &H80000000
End Enum

Public Enum jgexTabKeyBehaviorConstants
    jgexColumnNavigation = 0
    jgexControlNavigation = 1
End Enum

Public Enum jgexEditModeConstants
    jgexEditModeOff = 0
    jgexEditModeOn = 1
End Enum

Public Enum jgexOleDropModeConstants
    jgexOleDropNone = 0
    jgexOleDropManual = 1
End Enum

Public Enum jgexHeaderStyleConstants
    jgexHSDouble3D = 0
    jgexHSNoBorder = 1
    jgexHSSingleFlat = 2
    jgexHSSingle3D = 3
End Enum

Public Enum jgexHitTestConstants
    jgexHTNoWhere = 0
    jgexHTGroupByBox = 1
    jgexHTColumnHeader = 2
    jgexHTRowHeader = 3
    jgexHTNewRow = 4
    jgexHTCell = 5
    jgexHTCard = 6
    jgexHTBackGround = 7
End Enum

Public Enum jgexSortOrderConstants
    jgexSortDescending = -1
    jgexSortNone = 0
    jgexSortAscending = 1
End Enum

Public Enum jgexSortTypeConstants
    jgexSortTypeString = 1
    jgexSortTypeNumeric = 2
    jgexSortTypeDate = 3
    jgexSortTypeDateTime = 4
    jgexSortTypeTime = 5
End Enum

Public Enum jgexSelectionStyleConstants
    jgexEntireRow = 0
    jgexSingleCell = 1
End Enum

Public Enum jgexConditionOperatorConstants
    jgexEqual = 0
    jgexNotEqual = 1
    jgexGreaterThan = 2
    jgexLessThan = 3
    jgexGreaterThanOrEqualTo = 4
    jgexLessThanOrEqualTo = 5
    jgexBetween = 6
    jgexNotBetween = 7
    jgexContains = 8
    jgexNotContains = 9
End Enum

Public Enum jgexLockTypeConstants
    jgexLockReadOnly = 1
    jgexLockPessimistic = 2
    jgexLockOptimistic = 3
    jgexLockBatchOptimistic = 4
End Enum

Public Enum jgexRecordsetTypeConstants
    jgexRSDAOTable = 1
    jgexRSDAODynaset = 2
    jgexRSDAOSnapshot = 4
    jgexRSADOKeyset = 1
    jgexRSADOStatic = 3
End Enum

Public Enum jgexViewConstants
    jgexTable = 0
    jgexCard = 1
End Enum

Public Enum jgexGroupChange
    jgexGroupInsert = 0
    jgexGroupDelete = 1
    jgexGroupMove = 2
End Enum

Public Enum jgexNewRowPosConstants
    jgexTop = 0
    jgexBottom = 1
End Enum

Public Enum jgexDataModeConstants
    jgexDAO = 0
    jgexADO = 1
    jgexUnbound = 99
End Enum

Public Enum jgexBorderStyleConstants
    jgexNone = 0
    jgexFixed = 1
    jgexSingle3D = 2
    jgexFlat = 3
End Enum

Public Enum jgexAlignmentConstants
    jgexAlignLeft = 0
    jgexAlignCenter = 1
    jgexAlignRight = 2
End Enum

Public Enum jgexRowTypeConstants
    jgexRowTypeRecord = 0
    jgexRowTypeGroupHeader = 1
    jgexRowTypeGroupFooter = 2
End Enum

Public Enum jgexHideSelectionConstants
    jgexHideSelection = 0
    jgexHighLightInactive = 1
    jgexHighLightNormal = 2
End Enum

Public Enum jgexCursorLocationConstants
    jgexUseServer = 2
    jgexUseClient = 3
End Enum

Public Enum jgexGroupFooterStyleConstants
    jgexNoGroupFooter = 0
    jgexCaptionGroupFooter = 1
    jgexTotalsGroupFooter = 2
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_nFrozenColumns            As Integer
Private m_lRowHeight                As Long
Private m_eOLEDropMode              As jgexOleDropModeConstants
Private m_oADORecordset             As Object
Private m_oFmtConditions            As JSFmtConditions
Private m_clrForeColor              As OLE_COLOR
Private m_clrRowColorEven           As OLE_COLOR
Private m_clrRowColorOdd            As OLE_COLOR
Private m_oSortKeys                 As JSSortKeys
Private m_nCol                      As Integer
Private m_clrBackColorRowGroup      As OLE_COLOR
Private m_oGroups                   As JSGroups
Private m_bRecordNavigator          As Boolean
Private m_eBorderStyle              As jgexBorderStyleConstants
Private m_oColumns                  As JSColumns
Private m_bGroupByBoxVisible        As Boolean
Private m_clrBackColorGBBox         As OLE_COLOR
Private m_clrBackColor              As OLE_COLOR
Private m_lColumnHeaderHeight       As Long
Private m_bColumnHeaders            As Boolean
Private m_clrBackColorHeader        As OLE_COLOR
Private m_clrMaskColor              As OLE_COLOR
Private m_lDefaultColumnWidth       As Long
Private m_eView                     As jgexViewConstants
Private m_lImageWidth               As Long
Private m_lImageHeight              As Long
Private m_oGridImages               As JSGridImages
Private m_bRowHeaders               As Boolean
Private m_bAllowAddNew              As Boolean
Private m_eNewRowPos                As jgexNewRowPosConstants
Private m_lItemCount                As Long
Private m_eDataMode                 As jgexDataModeConstants
Private m_nLeftCol                  As Integer
Private m_oColumnHeaderFont         As Font
Private m_oFont                     As Font
Private m_lFirstItem                As Long
Private m_eGridLines                As jgexGridLinesConstants
Private m_clrGridLinesColor         As OLE_COLOR
Private m_clrBackColorBkg           As OLE_COLOR
Private m_lCardSpacing              As Long
Private m_lCardWidth                As Long
Private m_lRow                      As Long
Private m_sErrorText                As String
Private m_bAllowEdit                As Boolean
Private m_bDataChanged              As Boolean
Private m_bAllowDelete              As Boolean
Private m_sGroupByBoxInfoText       As String
Private m_clrForeColorHeader        As OLE_COLOR
Private m_clrForeColorRowGroup      As OLE_COLOR
Private m_clrBackColorInfoText      As OLE_COLOR
Private m_clrForeColorInfoText      As OLE_COLOR
Private m_bAutomaticArrange         As Boolean
Private m_sConnect                  As String
Private m_sDatabaseName             As String
Private m_bExclusive                As Boolean
Private m_sRecordSource             As String
Private m_eRecordsetType            As jgexRecordsetTypeConstants
Private m_bAllowColumnDrag          As Boolean
Private m_eSelectionStyle           As jgexSelectionStyleConstants
Private m_bAllowCardSizing          As Boolean
Private m_sCardCaptionPrefix        As String
Private m_bCardBorders              As Boolean
Private m_bContinuousScroll         As Boolean
Private m_oRecordset                As Object
Private m_bFullyLoaded              As Boolean
Private m_eLockType                 As jgexLockTypeConstants
Private m_lOptions                  As Long
Private m_bReadOnly                 As Boolean
Private m_bUseEvenOddColor          As Boolean
Private m_eHeaderStyle              As jgexHeaderStyleConstants
Private m_eHideSelection            As jgexHideSelectionConstants
Private m_bMultiSelect              As Boolean
Private m_oSelectedItems            As JSSelectedItems
Private m_bDetectRowDrag            As Boolean
Private m_sCalendarTodayText        As String
Private m_sCalendarNoneText         As String
Private m_eEditMode                 As jgexEditModeConstants
Private m_lSelStart                 As Long
Private m_lSelLength                As Long
Private m_sSelText                  As String
Private m_eCursorLocation           As jgexCursorLocationConstants
Private m_eTabKeyBehavior           As jgexTabKeyBehaviorConstants
Private m_bColumnAutoResize         As Boolean
Private m_oFormatStyles             As JSFormatStyles
Private m_oPrinterProperties        As JSPrinterProperties
Private m_nPreviewRowLines          As Integer
Private m_vPreviewColumn            As Variant
Private m_eGroupFooterStyle         As jgexGroupFooterStyleConstants
Private m_bShowEmptyFields          As Boolean
Private m_bActAsDropDown            As Boolean
Private m_vBoundColumnIndex         As Variant
Private m_vReplaceColumnIndex       As Variant
Private m_eGridLineStyle            As jgexGridLineStyleConstants
Private m_bEmptyRows                As Boolean
Private m_bRedraw                   As Boolean
Private m_eDefaultGroupMode         As jgexDefaultGroupModeConstants
Private m_bHoldSortSettings         As Boolean
Private m_sRecordNavigatorString    As String
Private m_bShowToolTips             As Boolean
Private m_bScrollToolTips           As Boolean
Private m_vScrollToolTipColumn      As Variant
Private m_bAutomaticSort            As Boolean
Private m_lPreviewRowIndent         As Long
Private m_bAllowRowSizing           As Boolean

'--- per-row virtual storage behind JSRowData wrappers
Private Type UcsCellData
    Value                   As Variant
    IconIndex               As Integer
    DisplayValue            As String
    CellStyle               As String
    CellPicture             As Picture
End Type

Private Type UcsRowData
    Bookmark                As Variant
    RowType                 As jgexRowTypeConstants
    GroupLevel              As Integer
    RecordCount             As Long
    RowHeight               As Long
    RowStyle                As String
    GroupCaption            As String
    GroupIconIndex          As Integer
    PreviewRowVisible       As Boolean
    Fetched                 As Boolean
    CellCount               As Integer
    Cells()                 As UcsCellData
    RowData                 As JSRowData
End Type

Private m_aRows()                   As UcsRowData
Private m_lRowsUBound               As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get FrozenColumns() As Integer
Attribute FrozenColumns.VB_Description = "Returns/sets the number of fixed columns at the left of the control."
    FrozenColumns = m_nFrozenColumns
End Property

Public Property Let FrozenColumns(ByVal nValue As Integer)
    m_nFrozenColumns = nValue
End Property

Public Property Get RowHeight() As Long
    RowHeight = m_lRowHeight
End Property

Public Property Let RowHeight(ByVal lValue As Long)
    m_lRowHeight = lValue
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns the handle of the cell editor."
End Property

Public Property Get OLEDropMode() As jgexOleDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/sets how OLE drop operations are handled."
    OLEDropMode = m_eOLEDropMode
End Property

Public Property Let OLEDropMode(ByVal eValue As jgexOleDropModeConstants)
    m_eOLEDropMode = eValue
End Property

Public Property Get ADORecordset() As Object
Attribute ADORecordset.VB_Description = "Returns/sets an ADO Recordset object defined by control's properties or by an existing ADO Recordset."
    Set ADORecordset = m_oADORecordset
End Property

Public Property Set ADORecordset(ByVal oValue As Object)
    Set m_oADORecordset = oValue
End Property

Public Property Get FmtConditions() As JSFmtConditions
Attribute FmtConditions.VB_Description = "Returns the JSFmtConditions collection of the control."
    Set FmtConditions = m_oFmtConditions
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text in cells."
    ForeColor = m_clrForeColor
End Property

Public Property Let ForeColor(ByVal lValue As OLE_COLOR)
    m_clrForeColor = lValue
End Property

Public Property Get RowColorEven() As OLE_COLOR
Attribute RowColorEven.VB_Description = "Returns/sets the background color for even rows."
    RowColorEven = m_clrRowColorEven
End Property

Public Property Let RowColorEven(ByVal lValue As OLE_COLOR)
    m_clrRowColorEven = lValue
End Property

Public Property Get RowColorOdd() As OLE_COLOR
Attribute RowColorOdd.VB_Description = "Returns/sets the background color for odd rows."
    RowColorOdd = m_clrRowColorOdd
End Property

Public Property Let RowColorOdd(ByVal lValue As OLE_COLOR)
    m_clrRowColorOdd = lValue
End Property

Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Returns the count of rows."
    RowCount = m_lItemCount
End Property

Public Property Get RowSelected(ByVal RowPosition As Long) As Boolean
Attribute RowSelected.VB_Description = "Returns/set whether a row is selected or not."
End Property

Public Property Let RowSelected(ByVal RowPosition As Long, ByVal bValue As Boolean)
End Property

Public Property Get SortKeys() As JSSortKeys
Attribute SortKeys.VB_Description = "Returns the JSSortKeys collection of the control."
    Set SortKeys = m_oSortKeys
End Property

Public Property Get Col() As Integer
Attribute Col.VB_Description = "Returns or sets the active column."
    Col = m_nCol
End Property

Public Property Let Col(ByVal nValue As Integer)
    Dim nLastCol        As Integer

    If m_nCol <> nValue Then
        nLastCol = m_nCol
        m_nCol = nValue
        RaiseEvent RowColChange(m_lRow, nLastCol)
    End If
End Property

Public Property Get BackColorRowGroup() As OLE_COLOR
Attribute BackColorRowGroup.VB_Description = "Returns/sets background color of the group rows."
    BackColorRowGroup = m_clrBackColorRowGroup
End Property

Public Property Let BackColorRowGroup(ByVal lValue As OLE_COLOR)
    m_clrBackColorRowGroup = lValue
End Property

Public Property Get Groups() As JSGroups
Attribute Groups.VB_Description = "Returns the JSGroups collection of the control."
    Set Groups = m_oGroups
End Property

Public Property Get RecordNavigator() As Boolean
Attribute RecordNavigator.VB_Description = "Determines whether the record navigator is visible or hidden."
    RecordNavigator = m_bRecordNavigator
End Property

Public Property Let RecordNavigator(ByVal bValue As Boolean)
    m_bRecordNavigator = bValue
End Property

Public Property Get BorderStyle() As jgexBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal eValue As jgexBorderStyleConstants)
    m_eBorderStyle = eValue
End Property

Public Property Get Columns() As JSColumns
Attribute Columns.VB_Description = "Returns the JSColumns collection."
    Set Columns = m_oColumns
End Property

Public Property Get GroupByBoxVisible() As Boolean
Attribute GroupByBoxVisible.VB_Description = "Determines whether the group by box is displayed."
    GroupByBoxVisible = m_bGroupByBoxVisible
End Property

Public Property Let GroupByBoxVisible(ByVal bValue As Boolean)
    m_bGroupByBoxVisible = bValue
End Property

Public Property Get BackColorGBBox() As OLE_COLOR
Attribute BackColorGBBox.VB_Description = "Returns/sets background color of the group by box."
    BackColorGBBox = m_clrBackColorGBBox
End Property

Public Property Let BackColorGBBox(ByVal lValue As OLE_COLOR)
    m_clrBackColorGBBox = lValue
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the grid and cards background color."
    BackColor = m_clrBackColor
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
    m_clrBackColor = lValue
End Property

Public Property Get ColumnHeaderHeight() As Long
Attribute ColumnHeaderHeight.VB_Description = "Returns/sets the height of the column header row."
    ColumnHeaderHeight = m_lColumnHeaderHeight
End Property

Public Property Let ColumnHeaderHeight(ByVal lValue As Long)
    m_lColumnHeaderHeight = lValue
End Property

Public Property Get ColumnHeaders() As Boolean
Attribute ColumnHeaders.VB_Description = "Determines whether column headers are displayed."
    ColumnHeaders = m_bColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal bValue As Boolean)
    m_bColumnHeaders = bValue
End Property

Public Property Get BackColorHeader() As OLE_COLOR
Attribute BackColorHeader.VB_Description = "Returns/sets background color of column and row headers."
    BackColorHeader = m_clrBackColorHeader
End Property

Public Property Let BackColorHeader(ByVal lValue As OLE_COLOR)
    m_clrBackColorHeader = lValue
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color used to create masks for images."
    MaskColor = m_clrMaskColor
End Property

Public Property Let MaskColor(ByVal lValue As OLE_COLOR)
    m_clrMaskColor = lValue
End Property

Public Property Get DefaultColumnWidth() As Long
Attribute DefaultColumnWidth.VB_Description = "Returns/sets a value indicating the default column width used for new columns."
    DefaultColumnWidth = m_lDefaultColumnWidth
End Property

Public Property Let DefaultColumnWidth(ByVal lValue As Long)
    m_lDefaultColumnWidth = lValue
End Property

Public Property Get View() As jgexViewConstants
Attribute View.VB_Description = "Determines the way records are displayed."
    View = m_eView
End Property

Public Property Let View(ByVal eValue As jgexViewConstants)
    m_eView = eValue
End Property

Public Property Get ImageWidth() As Long
Attribute ImageWidth.VB_Description = "Returns/sets the width of JSGridImage objects."
    ImageWidth = m_lImageWidth
End Property

Public Property Let ImageWidth(ByVal lValue As Long)
    m_lImageWidth = lValue
End Property

Public Property Get ImageHeight() As Long
Attribute ImageHeight.VB_Description = "Returns/sets the height of JSGridImage objects."
    ImageHeight = m_lImageHeight
End Property

Public Property Let ImageHeight(ByVal lValue As Long)
    m_lImageHeight = lValue
End Property

Public Property Get GridImages() As JSGridImages
Attribute GridImages.VB_Description = "Returns the JSGridImages collection used by the control."
    Set GridImages = m_oGridImages
End Property

Public Property Get RowHeaders() As Boolean
Attribute RowHeaders.VB_Description = "Determines whether the row headers are displayed."
    RowHeaders = m_bRowHeaders
End Property

Public Property Let RowHeaders(ByVal bValue As Boolean)
    m_bRowHeaders = bValue
End Property

Public Property Get AllowAddNew() As Boolean
Attribute AllowAddNew.VB_Description = "Determines whether the user can add records."
    AllowAddNew = m_bAllowAddNew
End Property

Public Property Let AllowAddNew(ByVal bValue As Boolean)
    m_bAllowAddNew = bValue
End Property

Public Property Get NewRowPos() As jgexNewRowPosConstants
Attribute NewRowPos.VB_Description = "Returns/sets the position of the row where the user can add new records."
    NewRowPos = m_eNewRowPos
End Property

Public Property Let NewRowPos(ByVal eValue As jgexNewRowPosConstants)
    m_eNewRowPos = eValue
End Property

Public Property Get ItemCount() As Long
Attribute ItemCount.VB_Description = "Returns/sets the number of items."
    ItemCount = m_lItemCount
End Property

Public Property Let ItemCount(ByVal lValue As Long)
    m_lItemCount = lValue
End Property

Public Property Get DataMode() As jgexDataModeConstants
Attribute DataMode.VB_Description = "Returns/sets a value representing the data retrieval mode."
    DataMode = m_eDataMode
End Property

Public Property Let DataMode(ByVal eValue As jgexDataModeConstants)
    m_eDataMode = eValue
End Property

Public Property Get LeftCol() As Integer
Attribute LeftCol.VB_Description = "Returns/sets the left-most visible column."
    LeftCol = m_nLeftCol
End Property

Public Property Let LeftCol(ByVal nValue As Integer)
    m_nLeftCol = nValue
End Property

Public Property Get ColumnHeaderFont() As Font
Attribute ColumnHeaderFont.VB_Description = "Returns/sets  the font used in column headers."
    Set ColumnHeaderFont = m_oColumnHeaderFont
End Property

Public Property Set ColumnHeaderFont(ByVal oValue As Font)
    Set m_oColumnHeaderFont = oValue
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
Attribute Font.VB_Description = "Returns/sets a Font object."
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal oValue As Font)
    Set m_oFont = oValue
End Property

Public Property Get FirstItem() As Long
Attribute FirstItem.VB_Description = "Returns/sets the row position of the first visible row or card."
    FirstItem = m_lFirstItem
End Property

Public Property Let FirstItem(ByVal lValue As Long)
    If m_lFirstItem <> lValue Then
        m_lFirstItem = lValue
        RaiseEvent FirstItemChange
    End If
End Property

Public Property Get GridLines() As jgexGridLinesConstants
Attribute GridLines.VB_Description = "Determines whether the control will draw lines between cells."
    GridLines = m_eGridLines
End Property

Public Property Let GridLines(ByVal eValue As jgexGridLinesConstants)
    m_eGridLines = eValue
End Property

Public Property Get GridLinesColor() As OLE_COLOR
Attribute GridLinesColor.VB_Description = "Returns/sets the color used to draw grid lines."
    GridLinesColor = m_clrGridLinesColor
End Property

Public Property Let GridLinesColor(ByVal lValue As OLE_COLOR)
    m_clrGridLinesColor = lValue
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets the background color of the area outside the grid or card."
    BackColorBkg = m_clrBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal lValue As OLE_COLOR)
    m_clrBackColorBkg = lValue
End Property

Public Property Get CardSpacing() As Long
Attribute CardSpacing.VB_Description = "Returns/sets the horizontal and vertical space between cards."
    CardSpacing = m_lCardSpacing
End Property

Public Property Let CardSpacing(ByVal lValue As Long)
    m_lCardSpacing = lValue
End Property

Public Property Get CardWidth() As Long
Attribute CardWidth.VB_Description = "Returns/sets the width of a card."
    CardWidth = m_lCardWidth
End Property

Public Property Let CardWidth(ByVal lValue As Long)
    m_lCardWidth = lValue
End Property

Public Property Get RowBookmark(ByVal RowIndex As Long) As Variant
Attribute RowBookmark.VB_Description = "Returns/sets a value containing a bookmark for a row."
    pvEnsureRow RowIndex
    RowBookmark = m_aRows(RowIndex).Bookmark
End Property

Public Property Let RowBookmark(ByVal RowIndex As Long, ByVal vntValue As Variant)
    pvEnsureRow RowIndex
    m_aRows(RowIndex).Bookmark = vntValue
End Property

Public Property Get Row() As Long
Attribute Row.VB_Description = "Returns/sets the current row/card position."
    Row = m_lRow
End Property

Public Property Let Row(ByVal lValue As Long)
    Dim lLastRow        As Long

    If m_lRow <> lValue Then
        lLastRow = m_lRow
        m_lRow = lValue
        RaiseEvent RowColChange(lLastRow, m_nCol)
    End If
End Property

Public Property Get ErrorText() As String
Attribute ErrorText.VB_Description = "Returns/sets the error message string from the underlying data source. "
    ErrorText = m_sErrorText
End Property

Public Property Let ErrorText(ByVal sValue As String)
    m_sErrorText = sValue
End Property

Public Property Get AllowEdit() As Boolean
Attribute AllowEdit.VB_Description = "Determines wheter the user can edit records. "
    AllowEdit = m_bAllowEdit
End Property

Public Property Let AllowEdit(ByVal bValue As Boolean)
    m_bAllowEdit = bValue
End Property

Public Property Get DataChanged() As Boolean
Attribute DataChanged.VB_Description = "Returns/sets a value indicating that the data has been changed by some process other than that of retrieving data from the current record."
    DataChanged = m_bDataChanged
End Property

Public Property Let DataChanged(ByVal bValue As Boolean)
    m_bDataChanged = bValue
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the handle of the control."
    hWnd = UserControl.hWnd
End Property

Public Property Get AllowDelete() As Boolean
Attribute AllowDelete.VB_Description = "Determines whether the user can delete records."
    AllowDelete = m_bAllowDelete
End Property

Public Property Let AllowDelete(ByVal bValue As Boolean)
    m_bAllowDelete = bValue
End Property

Public Property Get GroupByBoxInfoText() As String
Attribute GroupByBoxInfoText.VB_Description = "Returns/sets the text displayed in the group by box when no groups are set."
    GroupByBoxInfoText = m_sGroupByBoxInfoText
End Property

Public Property Let GroupByBoxInfoText(ByVal sValue As String)
    m_sGroupByBoxInfoText = sValue
End Property

Public Property Get ForeColorHeader() As OLE_COLOR
Attribute ForeColorHeader.VB_Description = "Returns/sets the foreground color used to display text in headers."
    ForeColorHeader = m_clrForeColorHeader
End Property

Public Property Let ForeColorHeader(ByVal lValue As OLE_COLOR)
    m_clrForeColorHeader = lValue
End Property

Public Property Get ForeColorRowGroup() As OLE_COLOR
Attribute ForeColorRowGroup.VB_Description = "Returns/sets the foreground color used to display text in group rows."
    ForeColorRowGroup = m_clrForeColorRowGroup
End Property

Public Property Let ForeColorRowGroup(ByVal lValue As OLE_COLOR)
    m_clrForeColorRowGroup = lValue
End Property

Public Property Get BackColorInfoText() As OLE_COLOR
Attribute BackColorInfoText.VB_Description = "Returns/sets the background color of the rectangle surrounding the information text displayed in the group by box "
    BackColorInfoText = m_clrBackColorInfoText
End Property

Public Property Let BackColorInfoText(ByVal lValue As OLE_COLOR)
    m_clrBackColorInfoText = lValue
End Property

Public Property Get ForeColorInfoText() As OLE_COLOR
Attribute ForeColorInfoText.VB_Description = "Returns/sets the foreground color used to display the information text in the group by box."
    ForeColorInfoText = m_clrForeColorInfoText
End Property

Public Property Let ForeColorInfoText(ByVal lValue As OLE_COLOR)
    m_clrForeColorInfoText = lValue
End Property

Public Property Get AutomaticArrange() As Boolean
Attribute AutomaticArrange.VB_Description = "Enables automatic sorting and groupping of records after commit changes made by the user."
    AutomaticArrange = m_bAutomaticArrange
End Property

Public Property Let AutomaticArrange(ByVal bValue As Boolean)
    m_bAutomaticArrange = bValue
End Property

Public Property Get Connect() As String
Attribute Connect.VB_Description = "Returns/sets a value that represents the connect parameter used to open the database."
    Connect = m_sConnect
End Property

Public Property Let Connect(ByVal sValue As String)
    m_sConnect = sValue
End Property

Public Property Get DatabaseName() As String
Attribute DatabaseName.VB_Description = "Returns/sets the name and location of the source of data."
    DatabaseName = m_sDatabaseName
End Property

Public Property Let DatabaseName(ByVal sValue As String)
    m_sDatabaseName = sValue
End Property

Public Property Get Exclusive() As Boolean
Attribute Exclusive.VB_Description = "Determines whether the database is opened for single- or multi-user access."
    Exclusive = m_bExclusive
End Property

Public Property Let Exclusive(ByVal bValue As Boolean)
    m_bExclusive = bValue
End Property

Public Property Get RecordSource() As String
Attribute RecordSource.VB_Description = "Returns/sets the underlying table, SQL statement, or QueryDef object for the Recordset."
    RecordSource = m_sRecordSource
End Property

Public Property Let RecordSource(ByVal sValue As String)
    m_sRecordSource = sValue
End Property

Public Property Get RecordsetType() As jgexRecordsetTypeConstants
Attribute RecordsetType.VB_Description = "Returns/sets a value indicating the type of recordset object."
    RecordsetType = m_eRecordsetType
End Property

Public Property Let RecordsetType(ByVal eValue As jgexRecordsetTypeConstants)
    m_eRecordsetType = eValue
End Property

Public Property Get AllowColumnDrag() As Boolean
Attribute AllowColumnDrag.VB_Description = "Determines whether the user can drag columns."
    AllowColumnDrag = m_bAllowColumnDrag
End Property

Public Property Let AllowColumnDrag(ByVal bValue As Boolean)
    m_bAllowColumnDrag = bValue
End Property

Public Property Get SelectionStyle() As jgexSelectionStyleConstants
Attribute SelectionStyle.VB_Description = "Returns/sets the style used to highlight a selected cell."
    SelectionStyle = m_eSelectionStyle
End Property

Public Property Let SelectionStyle(ByVal eValue As jgexSelectionStyleConstants)
    m_eSelectionStyle = eValue
End Property

Public Property Get AllowCardSizing() As Boolean
Attribute AllowCardSizing.VB_Description = "Determines whether the user can resize cards."
    AllowCardSizing = m_bAllowCardSizing
End Property

Public Property Let AllowCardSizing(ByVal bValue As Boolean)
    m_bAllowCardSizing = bValue
End Property

Public Property Get CardCaptionPrefix() As String
Attribute CardCaptionPrefix.VB_Description = "Returns/sets the text that will appear before the caption in every card."
    CardCaptionPrefix = m_sCardCaptionPrefix
End Property

Public Property Let CardCaptionPrefix(ByVal sValue As String)
    m_sCardCaptionPrefix = sValue
End Property

Public Property Get CardBorders() As Boolean
Attribute CardBorders.VB_Description = "Returns/sets whether the control will show borders in cards."
    CardBorders = m_bCardBorders
End Property

Public Property Let CardBorders(ByVal bValue As Boolean)
    m_bCardBorders = bValue
End Property

Public Property Get ContinuousScroll() As Boolean
Attribute ContinuousScroll.VB_Description = "Determines whether the control should scroll its contents while the user moves the scroll box along the vertical scroll bar."
    ContinuousScroll = m_bContinuousScroll
End Property

Public Property Let ContinuousScroll(ByVal bValue As Boolean)
    m_bContinuousScroll = bValue
End Property

Public Property Get Recordset() As Object
Attribute Recordset.VB_Description = "Returns/sets a Recordset object defined by control's properties or by an existing Recordset."
    Set Recordset = m_oRecordset
End Property

Public Property Set Recordset(ByVal oValue As Object)
    Set m_oRecordset = oValue
End Property

Public Property Get Value(ByVal ColIndex As Integer) As Variant
Attribute Value.VB_Description = "Returns/sets the value of a column in the current row."
End Property

Public Property Let Value(ByVal ColIndex As Integer, ByVal vntValue As Variant)
End Property

Public Property Get FullyLoaded() As Boolean
Attribute FullyLoaded.VB_Description = "Returns True if the control has loaded all bookmark records in the recordset."
    FullyLoaded = m_bFullyLoaded
End Property

Public Property Get LockType() As jgexLockTypeConstants
Attribute LockType.VB_Description = "Returns or sets the type of locking for the Recordset."
    LockType = m_eLockType
End Property

Public Property Let LockType(ByVal eValue As jgexLockTypeConstants)
    m_eLockType = eValue
End Property

Public Property Get Options() As Long
Attribute Options.VB_Description = "Returns/sets a value that specifies one or more characteristics of the Recordset object."
    Options = m_lOptions
End Property

Public Property Let Options(ByVal lValue As Long)
    m_lOptions = lValue
End Property

Public Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "Returns/sets a value indicating whether the Database object must be opened as read only."
    ReadOnly = m_bReadOnly
End Property

Public Property Let ReadOnly(ByVal bValue As Boolean)
    m_bReadOnly = bValue
End Property

Public Property Get UseEvenOddColor() As Boolean
Attribute UseEvenOddColor.VB_Description = "Determines wheter the control uses alternating row colors."
    UseEvenOddColor = m_bUseEvenOddColor
End Property

Public Property Let UseEvenOddColor(ByVal bValue As Boolean)
    m_bUseEvenOddColor = bValue
End Property

Public Property Get HeaderStyle() As jgexHeaderStyleConstants
Attribute HeaderStyle.VB_Description = "Returns/sets the display style for headers."
    HeaderStyle = m_eHeaderStyle
End Property

Public Property Let HeaderStyle(ByVal eValue As jgexHeaderStyleConstants)
    m_eHeaderStyle = eValue
End Property

Public Property Get HideSelection() As jgexHideSelectionConstants
Attribute HideSelection.VB_Description = "Determines how selected items will be displayed when the control loses focus."
    HideSelection = m_eHideSelection
End Property

Public Property Let HideSelection(ByVal eValue As jgexHideSelectionConstants)
    m_eHideSelection = eValue
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Determines whether a user can make multiple selections."
    MultiSelect = m_bMultiSelect
End Property

Public Property Let MultiSelect(ByVal bValue As Boolean)
    m_bMultiSelect = bValue
End Property

Public Property Get SelectedItems() As JSSelectedItems
Attribute SelectedItems.VB_Description = "Returns the JSSelectedItems collection in the control."
    Set SelectedItems = m_oSelectedItems
End Property

Public Property Get DetectRowDrag() As Boolean
Attribute DetectRowDrag.VB_Description = "Determines wheter dragging of the selected rows are detected."
    DetectRowDrag = m_bDetectRowDrag
End Property

Public Property Let DetectRowDrag(ByVal bValue As Boolean)
    m_bDetectRowDrag = bValue
End Property

Public Property Get CalendarTodayText() As String
Attribute CalendarTodayText.VB_Description = "Returns/sets the text displayed in the calendar drop-down 'Today' button."
    CalendarTodayText = m_sCalendarTodayText
End Property

Public Property Let CalendarTodayText(ByVal sValue As String)
    m_sCalendarTodayText = sValue
End Property

Public Property Get CalendarNoneText() As String
Attribute CalendarNoneText.VB_Description = "Returns/sets the text displayed in the calendar drop-down 'None' button."
    CalendarNoneText = m_sCalendarNoneText
End Property

Public Property Let CalendarNoneText(ByVal sValue As String)
    m_sCalendarNoneText = sValue
End Property

Public Property Get EditMode() As jgexEditModeConstants
Attribute EditMode.VB_Description = "Returns/sets whether the cell editor is active."
    EditMode = m_eEditMode
End Property

Public Property Let EditMode(ByVal eValue As jgexEditModeConstants)
    m_eEditMode = eValue
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected in a cell editor; indicates the position of the insertion point if no text is selected."
    SelStart = m_lSelStart
End Property

Public Property Let SelStart(ByVal lValue As Long)
    m_lSelStart = lValue
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected. "
    SelLength = m_lSelLength
End Property

Public Property Let SelLength(ByVal lValue As Long)
    m_lSelLength = lValue
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text in the cell editor."
    SelText = m_sSelText
End Property

Public Property Let SelText(ByVal sValue As String)
    m_sSelText = sValue
End Property

Public Property Get RowExpanded(ByVal RowPosition As Long) As Boolean
Attribute RowExpanded.VB_Description = "Returns/sets whether a group row is expanded or collapsed."
End Property

Public Property Let RowExpanded(ByVal RowPosition As Long, ByVal bValue As Boolean)
End Property

Public Property Get CursorLocation() As jgexCursorLocationConstants
Attribute CursorLocation.VB_Description = "Returns/sets the location of the cursor engine (ADO mode only)."
    CursorLocation = m_eCursorLocation
End Property

Public Property Let CursorLocation(ByVal eValue As jgexCursorLocationConstants)
    m_eCursorLocation = eValue
End Property

Public Property Get TabKeyBehavior() As jgexTabKeyBehaviorConstants
Attribute TabKeyBehavior.VB_Description = "Determines the behavior of the tab key."
    TabKeyBehavior = m_eTabKeyBehavior
End Property

Public Property Let TabKeyBehavior(ByVal eValue As jgexTabKeyBehaviorConstants)
    m_eTabKeyBehavior = eValue
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Determines whether the control can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bValue As Boolean)
    UserControl.Enabled = bValue
End Property

Public Property Get ColumnAutoResize() As Boolean
Attribute ColumnAutoResize.VB_Description = "Determines whether visible columns should be automatically sized to fit on the control's client width."
    ColumnAutoResize = m_bColumnAutoResize
End Property

Public Property Let ColumnAutoResize(ByVal bValue As Boolean)
    m_bColumnAutoResize = bValue
End Property

Public Property Get FormatStyles() As JSFormatStyles
Attribute FormatStyles.VB_Description = "Returns the JSFormatStyles collection."
    Set FormatStyles = m_oFormatStyles
End Property

Public Property Get PrinterProperties() As JSPrinterProperties
Attribute PrinterProperties.VB_Description = "Returns the JSPrinterProperties object of the control."
    Set PrinterProperties = m_oPrinterProperties
End Property

Public Property Get PreviewRowLines() As Integer
Attribute PreviewRowLines.VB_Description = "Returns/sets the number of lines to be displayed in preview rows."
    PreviewRowLines = m_nPreviewRowLines
End Property

Public Property Let PreviewRowLines(ByVal nValue As Integer)
    m_nPreviewRowLines = nValue
End Property

Public Property Get PreviewColumn() As Variant
Attribute PreviewColumn.VB_Description = "Returns/sets the index or key of the column to be displayed as the preview row."
    PreviewColumn = m_vPreviewColumn
End Property

Public Property Let PreviewColumn(ByVal vntValue As Variant)
    m_vPreviewColumn = vntValue
End Property

Public Property Get GroupFooterStyle() As jgexGroupFooterStyleConstants
Attribute GroupFooterStyle.VB_Description = "Determines the style of group footers. "
    GroupFooterStyle = m_eGroupFooterStyle
End Property

Public Property Let GroupFooterStyle(ByVal eValue As jgexGroupFooterStyleConstants)
    m_eGroupFooterStyle = eValue
End Property

Public Property Get ShowEmptyFields() As Boolean
Attribute ShowEmptyFields.VB_Description = "Determines whether the control should display empty fields in cards."
    ShowEmptyFields = m_bShowEmptyFields
End Property

Public Property Let ShowEmptyFields(ByVal bValue As Boolean)
    m_bShowEmptyFields = bValue
End Property

Public Property Get ActAsDropDown() As Boolean
Attribute ActAsDropDown.VB_Description = "Determines whether the control will behave as a drop down list for a column in another GridEX control."
    ActAsDropDown = m_bActAsDropDown
End Property

Public Property Let ActAsDropDown(ByVal bValue As Boolean)
    m_bActAsDropDown = bValue
End Property

Public Property Get BoundColumnIndex() As Variant
Attribute BoundColumnIndex.VB_Description = "Returns/sets the index or key of the column used to supply a data value to another GridEX control."
    BoundColumnIndex = m_vBoundColumnIndex
End Property

Public Property Let BoundColumnIndex(ByVal vntValue As Variant)
    m_vBoundColumnIndex = vntValue
End Property

Public Property Get ReplaceColumnIndex() As Variant
Attribute ReplaceColumnIndex.VB_Description = "Returns/sets the index or key of the column that replaces Id values in a DropDown GridEX Control."
    ReplaceColumnIndex = m_vReplaceColumnIndex
End Property

Public Property Let ReplaceColumnIndex(ByVal vntValue As Variant)
    m_vReplaceColumnIndex = vntValue
End Property

Public Property Get GridLineStyle() As jgexGridLineStyleConstants
Attribute GridLineStyle.VB_Description = "Determines the grid lines style."
    GridLineStyle = m_eGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal eValue As jgexGridLineStyleConstants)
    m_eGridLineStyle = eValue
End Property

Public Property Get EmptyRows() As Boolean
Attribute EmptyRows.VB_Description = "Determines whether empty rows below the last row should be displayed. "
    EmptyRows = m_bEmptyRows
End Property

Public Property Let EmptyRows(ByVal bValue As Boolean)
    m_bEmptyRows = bValue
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Determines whether drawing is enabled or not."
    Redraw = m_bRedraw
End Property

Public Property Let Redraw(ByVal bValue As Boolean)
    m_bRedraw = bValue
End Property

Public Property Get DefaultGroupMode() As jgexDefaultGroupModeConstants
Attribute DefaultGroupMode.VB_Description = "Determines how the control groups records. "
    DefaultGroupMode = m_eDefaultGroupMode
End Property

Public Property Let DefaultGroupMode(ByVal eValue As jgexDefaultGroupModeConstants)
    m_eDefaultGroupMode = eValue
End Property

Public Property Get HoldSortSettings() As Boolean
Attribute HoldSortSettings.VB_Description = "Determines whether the group and sort settings will be held when rebinding the control."
    HoldSortSettings = m_bHoldSortSettings
End Property

Public Property Let HoldSortSettings(ByVal bValue As Boolean)
    m_bHoldSortSettings = bValue
End Property

Public Property Get RecordNavigatorString() As String
Attribute RecordNavigatorString.VB_Description = "Returns/sets the text displayed in the record navigator."
    RecordNavigatorString = m_sRecordNavigatorString
End Property

Public Property Let RecordNavigatorString(ByVal sValue As String)
    m_sRecordNavigatorString = sValue
End Property

Public Property Get ShowToolTips() As Boolean
Attribute ShowToolTips.VB_Description = "Determines whether the control should show tool tips over cells where text is not entirely displayed."
    ShowToolTips = m_bShowToolTips
End Property

Public Property Let ShowToolTips(ByVal bValue As Boolean)
    m_bShowToolTips = bValue
End Property

Public Property Get ScrollToolTips() As Boolean
Attribute ScrollToolTips.VB_Description = "Determines whether the control should show tool tips while the user scrolls vertically."
    ScrollToolTips = m_bScrollToolTips
End Property

Public Property Let ScrollToolTips(ByVal bValue As Boolean)
    m_bScrollToolTips = bValue
End Property

Public Property Get ScrollToolTipColumn() As Variant
Attribute ScrollToolTipColumn.VB_Description = "Returns/sets the index or key of the column used to get values for scroll tool tips."
    ScrollToolTipColumn = m_vScrollToolTipColumn
End Property

Public Property Let ScrollToolTipColumn(ByVal vntValue As Variant)
    m_vScrollToolTipColumn = vntValue
End Property

Public Property Get AutomaticSort() As Boolean
Attribute AutomaticSort.VB_Description = "Determines whether the control should sort values automatically when the user clicks on a column header."
    AutomaticSort = m_bAutomaticSort
End Property

Public Property Let AutomaticSort(ByVal bValue As Boolean)
    m_bAutomaticSort = bValue
End Property

Public Property Get PreviewRowIndent() As Long
Attribute PreviewRowIndent.VB_Description = "Returns/sets the left indent, in twips, for text in the preview row."
    PreviewRowIndent = m_lPreviewRowIndent
End Property

Public Property Let PreviewRowIndent(ByVal lValue As Long)
    m_lPreviewRowIndent = lValue
End Property

Public Property Get AllowRowSizing() As Boolean
Attribute AllowRowSizing.VB_Description = "Determines whether users can resize rows."
    AllowRowSizing = m_bAllowRowSizing
End Property

Public Property Let AllowRowSizing(ByVal bValue As Boolean)
    m_bAllowRowSizing = bValue
End Property

Friend Property Get frRowColCount() As Integer
    frRowColCount = m_oColumns.Count
End Property

Friend Property Get frRowValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Variant
    pvFetchRow lRowIndex
    frRowValue = m_aRows(lRowIndex).Cells(nColIndex).Value
End Property

Friend Property Let frRowValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal vntValue As Variant)
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).Value = vntValue
End Property

Friend Property Get frRowIconIndex(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Integer
    pvFetchRow lRowIndex
    frRowIconIndex = m_aRows(lRowIndex).Cells(nColIndex).IconIndex
End Property

Friend Property Let frRowIconIndex(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal nValue As Integer)
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).IconIndex = nValue
End Property

Friend Property Get frRowDisplayValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    pvFetchRow lRowIndex
    frRowDisplayValue = m_aRows(lRowIndex).Cells(nColIndex).DisplayValue
End Property

Friend Property Let frRowDisplayValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal sValue As String)
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).DisplayValue = sValue
End Property

Friend Property Get frRowCellStyle(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    pvFetchRow lRowIndex
    frRowCellStyle = m_aRows(lRowIndex).Cells(nColIndex).CellStyle
End Property

Friend Property Let frRowCellStyle(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal sValue As String)
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).CellStyle = sValue
End Property

Friend Property Get frRowCellPicture(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Picture
    pvFetchRow lRowIndex
    Set frRowCellPicture = m_aRows(lRowIndex).Cells(nColIndex).CellPicture
End Property

Friend Property Set frRowCellPicture(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal oValue As Picture)
    pvFetchRow lRowIndex
    Set m_aRows(lRowIndex).Cells(nColIndex).CellPicture = oValue
End Property

Friend Property Get frRowBookmark(ByVal lRowIndex As Long) As Variant
    pvEnsureRow lRowIndex
    frRowBookmark = m_aRows(lRowIndex).Bookmark
End Property

Friend Property Get frRowType(ByVal lRowIndex As Long) As jgexRowTypeConstants
    pvEnsureRow lRowIndex
    frRowType = m_aRows(lRowIndex).RowType
End Property

Friend Property Get frRowGroupLevel(ByVal lRowIndex As Long) As Integer
    pvEnsureRow lRowIndex
    frRowGroupLevel = m_aRows(lRowIndex).GroupLevel
End Property

Friend Property Get frRowRecordCount(ByVal lRowIndex As Long) As Long
    pvEnsureRow lRowIndex
    frRowRecordCount = m_aRows(lRowIndex).RecordCount
End Property

Friend Property Get frRowHeight(ByVal lRowIndex As Long) As Long
    pvEnsureRow lRowIndex
    frRowHeight = m_aRows(lRowIndex).RowHeight
End Property

Friend Property Let frRowHeight(ByVal lRowIndex As Long, ByVal lValue As Long)
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).RowHeight = lValue
End Property

Friend Property Get frRowStyle(ByVal lRowIndex As Long) As String
    pvEnsureRow lRowIndex
    frRowStyle = m_aRows(lRowIndex).RowStyle
End Property

Friend Property Let frRowStyle(ByVal lRowIndex As Long, ByVal sValue As String)
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).RowStyle = sValue
End Property

Friend Property Get frRowGroupCaption(ByVal lRowIndex As Long) As String
    pvEnsureRow lRowIndex
    frRowGroupCaption = m_aRows(lRowIndex).GroupCaption
End Property

Friend Property Let frRowGroupCaption(ByVal lRowIndex As Long, ByVal sValue As String)
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).GroupCaption = sValue
End Property

Friend Property Get frRowGroupIconIndex(ByVal lRowIndex As Long) As Integer
    pvEnsureRow lRowIndex
    frRowGroupIconIndex = m_aRows(lRowIndex).GroupIconIndex
End Property

Friend Property Let frRowGroupIconIndex(ByVal lRowIndex As Long, ByVal nValue As Integer)
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).GroupIconIndex = nValue
End Property

Friend Property Get frRowPreviewRowVisible(ByVal lRowIndex As Long) As Boolean
    pvEnsureRow lRowIndex
    frRowPreviewRowVisible = m_aRows(lRowIndex).PreviewRowVisible
End Property

Friend Property Let frRowPreviewRowVisible(ByVal lRowIndex As Long, ByVal bValue As Boolean)
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).PreviewRowVisible = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub LoadLayoutFromURL(ByVal URL As String, Optional Rebind As Boolean = True, Optional ByVal DatabaseName As String)
Attribute LoadLayoutFromURL.VB_Description = "Loads a previously saved GridEX layout from an URL."
End Sub

Public Sub SetFocusRecordNavigator()
Attribute SetFocusRecordNavigator.VB_Description = "Set focus to the specific record box in the record navigator."
End Sub

Public Sub Refetch(Optional HoldSortSettings As Variant)
Attribute Refetch.VB_Description = "Forces a GridEX control to refresh its contents without re-opening the recordset. "
    pvResetRows False
    pvApplyHoldSort HoldSortSettings
End Sub

Public Sub PrintGrid(Optional UsePrintSetupDlg As Boolean, Optional PrintSelectedItems As Boolean)
Attribute PrintGrid.VB_Description = "Prints the contents of a GridEX control."
End Sub

Public Sub LoadLayout(ByVal FileName As String, Optional Rebind As Boolean = True, Optional ByVal DatabaseName As String)
Attribute LoadLayout.VB_Description = "Loads a previously saved GridEX layout from a file."
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Initiates an OLE drag/drop operation."
End Sub

Public Sub CollapseAll()
Attribute CollapseAll.VB_Description = "Collapses all group rows."
End Sub

Public Sub RefreshSort()
Attribute RefreshSort.VB_Description = "Forces the re-sort of the records."
End Sub

Public Function RowIndex(ByVal RowPosition As Long) As Long
Attribute RowIndex.VB_Description = "Returns the original index of a row."
    '--- no grouping/sorting yet: position maps 1:1 to original index;
    '--- group rows will return 0 once grouping lands
    If RowPosition >= 1 And RowPosition <= RowCount Then
        If Not IsGroupItem(RowPosition) Then
            RowIndex = RowPosition
        End If
    End If
End Function

Public Sub RefreshGroups(Optional ByVal AllCollapsed As Boolean)
Attribute RefreshGroups.VB_Description = "Forces recalculation of groups."
End Sub

Public Sub EnsureVisible(Optional ByVal Row As Long, Optional ByVal Col As Integer)
Attribute EnsureVisible.VB_Description = "Ensures visibility of a cell."
End Sub

Public Sub Rebind(Optional HoldSortSettings As Variant)
Attribute Rebind.VB_Description = "Forces re-creation of the recordset."
    pvResetRows True
    pvApplyHoldSort HoldSortSettings
End Sub

Public Function IsGroupItem(ByVal Row As Long) As Boolean
Attribute IsGroupItem.VB_Description = "Returns True if the specified row is a group row."
    If Row >= 1 And Row <= m_lRowsUBound Then
        IsGroupItem = (m_aRows(Row).RowType <> jgexRowTypeRecord)
    End If
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a GridEX control."
End Sub

Public Sub HoldFields()
Attribute HoldFields.VB_Description = "Prevents recalculation of the column layout when the recordset is created."
End Sub

Public Sub ClearFields()
Attribute ClearFields.VB_Description = "Restores the default layout columns."
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_Description = "Displays the About box for the control."
End Sub

Public Sub LoadEntireRecordset()
Attribute LoadEntireRecordset.VB_Description = "Loads the bookmarks of all records in the underlying Recordset."
End Sub

Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Expands all group rows."
End Sub

Public Function HitTest(ByVal X As Long, ByVal Y As Long) As jgexHitTestConstants
Attribute HitTest.VB_Description = "Returns the part of a GridEX control that contains the specified point."
End Function

Public Function ColFromPoint(ByVal X As Long, ByVal Y As Long) As JSColumn
Attribute ColFromPoint.VB_Description = "Returns a JSColumn object that contains a given point."
End Function

Public Function RowFromPoint(ByVal X As Long, ByVal Y As Long) As Long
Attribute RowFromPoint.VB_Description = "Returns the row that contains the given point."
End Function

Public Sub MoveToBookmark(ByVal vBookmark As Variant)
Attribute MoveToBookmark.VB_Description = "Sets the current row as the one that matches the Bookmark."
End Sub

Public Sub MoveToRowIndex(ByVal RowIndex As Long)
Attribute MoveToRowIndex.VB_Description = "Sets the current row as the one that matches the index."
End Sub

Public Sub Update()
Attribute Update.VB_Description = "Commit changes made to the current row writing to the database and re-position the record in case that needed."
End Sub

Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Sets the first open record as the current row."
End Sub

Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Sets the last open record as the current row."
End Sub

Public Sub MovePrevious()
Attribute MovePrevious.VB_Description = "Moves to the previous open record."
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Moves to the next open record."
End Sub

Public Sub MoveRelative(ByVal nRows As Long)
Attribute MoveRelative.VB_Description = "Moves the current row to the specified position."
End Sub

Public Sub Delete()
Attribute Delete.VB_Description = "Deletes selected rows."
End Sub

Public Sub SearchNewRecords()
Attribute SearchNewRecords.VB_Description = "Searches for records added after all bookmarks have been loaded."
End Sub

Public Sub RefreshRowBookmark(ByVal Bookmark As Variant)
Attribute RefreshRowBookmark.VB_Description = "Refreshes data of the record that matches the Bookmark."
    Dim lIdx            As Long

    For lIdx = 1 To m_lRowsUBound
        If Not IsEmpty(m_aRows(lIdx).Bookmark) Then
            If m_aRows(lIdx).Bookmark = Bookmark Then
                pvResetRow lIdx, False
                Exit For
            End If
        End If
    Next
End Sub

Public Sub RefreshRowIndex(ByVal RowIndex As Long)
Attribute RefreshRowIndex.VB_Description = "Refreshes data of the record that matches the index."
    If RowIndex >= 1 And RowIndex <= m_lRowsUBound Then
        pvResetRow RowIndex, False
    End If
End Sub

Public Function GroupRowLevel(ByVal RowPosition As Long) As Integer
Attribute GroupRowLevel.VB_Description = "Returns the level of a group row."
End Function

Public Sub ExpandSelection()
Attribute ExpandSelection.VB_Description = "Expands and selects children records of selected group rows."
End Sub

Public Function GetClipString(Optional ByVal IncludeHeaders As Boolean, Optional ByVal ColumnDelimeter As String, Optional ByVal RowDelimeter As String) As String
Attribute GetClipString.VB_Description = "Returns the contents of the selected rows."
End Function

Public Function CellVisible(Optional RowPosition As Long, Optional ColPosition As Integer) As Boolean
Attribute CellVisible.VB_Description = "Returns True if a cell is visible."
End Function

Public Function CellLeft(Optional RowPosition As Long, Optional ColPosition As Integer) As Long
Attribute CellLeft.VB_Description = "Returns the left of a cell."
End Function

Public Function CellTop(Optional RowPosition As Long, Optional ColPosition As Integer) As Long
Attribute CellTop.VB_Description = "Returns the top of a cell."
End Function

Public Function CellWidth(Optional RowPosition As Long, Optional ColPosition As Integer) As Long
Attribute CellWidth.VB_Description = "Returns the width of a cell."
End Function

Public Function CellHeight(Optional RowPosition As Long, Optional ColPosition As Integer) As Long
Attribute CellHeight.VB_Description = "Returns the height of a cell."
End Function

Public Function Find(ByVal ColIndex As Integer, ByVal Operator As jgexConditionOperatorConstants, ByVal Value1 As Variant, Optional ByVal Value2 As Variant, Optional Start As Long = -1, Optional SearchDirection As Long) As Boolean
Attribute Find.VB_Description = "Sets the current row to the one that matches the specified criteria."
End Function

Public Sub SaveLayout(ByVal FileName As String)
Attribute SaveLayout.VB_Description = "Saves the current GridEX layout in a file."
End Sub

Public Sub PrintPreview(PreviewControl As Object, Optional PrintSelectedItems As Boolean)
Attribute PrintPreview.VB_Description = "Sends the contents of a GridEX control to be printed to a GEXPreview control."
End Sub

Public Function GroupFromPoint(ByVal X As Long, ByVal Y As Long) As JSGroup
Attribute GroupFromPoint.VB_Description = "Returns the JSGroup object in the specified coordinates."
End Function

Public Function LayoutString(Optional UnicodeCompression As Boolean) As String
Attribute LayoutString.VB_Description = "Returns a string that contains the GridEX control layout settings."
End Function

Public Sub LoadLayoutString(LayoutString As String, Optional Rebind As Boolean, Optional DatabaseName As String)
Attribute LoadLayoutString.VB_Description = "Loads a previously saved GridEX layout from a layout string."
End Sub

Public Function GetRowData(ByVal RowPosition As Long) As JSRowData
Attribute GetRowData.VB_Description = "Returns a JSRowData object representing a row."
    pvEnsureRow RowPosition
    With m_aRows(RowPosition)
        If .RowData Is Nothing Then
            Set .RowData = New JSRowData
            .RowData.frInit Me, RowPosition
        End If
        Set GetRowData = .RowData
    End With
End Function

'=========================================================================
' Functions
'=========================================================================

Private Sub pvEnsureRow(ByVal lRowIndex As Long)
    Dim lIdx            As Long

    If lRowIndex > m_lRowsUBound Then
        ReDim Preserve m_aRows(1 To lRowIndex) As UcsRowData
        For lIdx = m_lRowsUBound + 1 To lRowIndex
            With m_aRows(lIdx)
                .RowHeight = m_lRowHeight
                .PreviewRowVisible = True
            End With
        Next
        m_lRowsUBound = lRowIndex
    End If
End Sub

Private Sub pvEnsureRowCells(ByVal lRowIndex As Long)
    Dim nColCount       As Integer

    pvEnsureRow lRowIndex
    nColCount = m_oColumns.Count
    With m_aRows(lRowIndex)
        If .CellCount < nColCount Then
            ReDim Preserve .Cells(1 To nColCount) As UcsCellData
            .CellCount = nColCount
        End If
    End With
End Sub

Private Sub pvFetchRow(ByVal lRowIndex As Long)
    Dim oRowData        As JSRowData

    pvEnsureRowCells lRowIndex
    '--- in unbound mode cell data is supplied lazily by the client app
    '--- through the UnboundReadData event, once per row until reset
    If m_eDataMode <> jgexUnbound Then
        Exit Sub
    End If
    If lRowIndex > m_lItemCount Then
        Exit Sub
    End If
    If Not m_aRows(lRowIndex).Fetched Then
        '--- mark first so reads from inside the handler do not re-fire
        m_aRows(lRowIndex).Fetched = True
        Set oRowData = GetRowData(lRowIndex)
        RaiseEvent UnboundReadData(lRowIndex, m_aRows(lRowIndex).Bookmark, oRowData)
    End If
End Sub

Private Sub pvResetRow(ByVal lRowIndex As Long, ByVal bFullReset As Boolean)
    With m_aRows(lRowIndex)
        .Fetched = False
        If .CellCount > 0 Then
            ReDim .Cells(1 To .CellCount) As UcsCellData
        End If
        If bFullReset Then
            .Bookmark = Empty
            .RowType = jgexRowTypeRecord
            .GroupLevel = 0
            .RecordCount = 0
            .RowHeight = m_lRowHeight
            .RowStyle = vbNullString
            .GroupCaption = vbNullString
            .GroupIconIndex = 0
            .PreviewRowVisible = True
        End If
    End With
End Sub

Private Sub pvResetRows(ByVal bFullReset As Boolean)
    Dim lIdx            As Long

    For lIdx = 1 To m_lRowsUBound
        pvResetRow lIdx, bFullReset
    Next
End Sub

Private Sub pvApplyHoldSort(HoldSortSettings As Variant)
    Dim bHold           As Boolean

    If IsMissing(HoldSortSettings) Then
        bHold = m_bHoldSortSettings
    Else
        bHold = CBool(HoldSortSettings)
    End If
    If Not bHold Then
        m_oSortKeys.Clear
        m_oGroups.Clear
    End If
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
    Set m_oColumns = New JSColumns
    Set m_oFmtConditions = New JSFmtConditions
    Set m_oGridImages = New JSGridImages
    Set m_oGroups = New JSGroups
    Set m_oSortKeys = New JSSortKeys
    Set m_oSelectedItems = New JSSelectedItems
    Set m_oFormatStyles = New JSFormatStyles
    Set m_oPrinterProperties = New JSPrinterProperties
    Set m_oFont = New StdFont
    m_oFont.Name = "MS Sans Serif"
    m_oFont.Size = 8.25
    Set m_oColumnHeaderFont = New StdFont
    m_oColumnHeaderFont.Name = "MS Sans Serif"
    m_oColumnHeaderFont.Size = 8.25
    m_clrBackColor = vbWindowBackground
    m_clrForeColor = vbWindowText
    m_clrBackColorHeader = vbButtonFace
    m_clrForeColorHeader = vbButtonText
    m_clrBackColorBkg = vbApplicationWorkspace
    m_clrBackColorGBBox = vbButtonFace
    m_clrBackColorInfoText = vbButtonFace
    m_clrForeColorInfoText = vbButtonText
    m_clrBackColorRowGroup = vbButtonFace
    m_clrForeColorRowGroup = vbButtonText
    m_clrRowColorEven = vbWindowBackground
    m_clrRowColorOdd = vbWindowBackground
    m_clrGridLinesColor = vbButtonFace
    m_clrMaskColor = &HC0C0C0
    m_eGridLines = jgexGLBoth
    m_eGridLineStyle = jgexGLSSolid
    m_eHeaderStyle = jgexHSDouble3D
    m_eView = jgexTable
    m_eSelectionStyle = jgexEntireRow
    m_eHideSelection = jgexHighLightInactive
    m_eNewRowPos = jgexTop
    m_eTabKeyBehavior = jgexColumnNavigation
    m_eDefaultGroupMode = jgexDGMExpanded
    m_eGroupFooterStyle = jgexNoGroupFooter
    m_eDataMode = jgexADO
    m_eRecordsetType = jgexRSADOKeyset
    m_eLockType = jgexLockOptimistic
    m_eCursorLocation = jgexUseClient
    m_eBorderStyle = jgexSingle3D
    m_lDefaultColumnWidth = 1000
    m_lColumnHeaderHeight = 285
    m_lImageWidth = 16
    m_lImageHeight = 16
    m_lCardWidth = 3000
    m_lCardSpacing = 200
    m_nPreviewRowLines = 2
    m_bColumnHeaders = True
    m_bRowHeaders = True
    m_bAllowEdit = True
    m_bAllowColumnDrag = True
    m_bAllowCardSizing = True
    m_bCardBorders = True
    m_bShowEmptyFields = True
    m_bShowToolTips = True
    m_bScrollToolTips = True
    m_bAutomaticSort = True
    m_bAutomaticArrange = True
    m_bHoldSortSettings = True
    m_bRedraw = True
    m_sCalendarTodayText = "Today"
    m_sCalendarNoneText = "None"
    m_sGroupByBoxInfoText = "Drag a column header here to group by that column."
End Sub

Private Sub UserControl_Terminate()
    Dim lIdx            As Long

    '--- detach outstanding JSRowData wrappers so their weak owner
    '--- pointers cannot dangle past the control lifetime
    For lIdx = 1 To m_lRowsUBound
        If Not m_aRows(lIdx).RowData Is Nothing Then
            m_aRows(lIdx).RowData.frTerm
        End If
    Next
End Sub