VERSION 5.00
Begin VB.UserControl GridEX 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2415
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
   Begin VB.HScrollBar hsbGrid 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
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
Private Const MODULE_NAME As String = "GridEX"

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

Private Const CHROME_COL_W              As Long = 18
Private Const GROUP_INDENT_W            As Long = 16
Private Const GROUP_BOX_W               As Long = 12
Private Const CHIP_LEFT                 As Long = 8
Private Const CHIP_TOP                  As Long = 7
Private Const CHIP_PAD                  As Long = 40
Private Const CHIP_GAP                  As Long = 8
Private Const CHECK_BOX_W               As Long = 11
Private Const CHECK_BOX_H               As Long = 12
Private Const CHECK_BOX_CLR             As Long = &H696969
Private Const CHECK_BOX_CLR_CUR         As Long = &H646464
Private Const DIVIDER_GRAB_W            As Long = 3
Private Const MIN_COL_W                 As Long = 10
Private Const ROWSEL_ZONE_W             As Long = 5
Private Const DRAG_SCROLL_ID            As Long = 1
Private Const DRAG_SCROLL_MS            As Long = 50

'--- property backing variables
Private m_lFrozenColumns            As Long
Private m_lRowHeight                As Long
Private m_eOLEDropMode              As jgexOleDropModeConstants
Private m_oADORecordset             As Object
Private m_oFmtConditions            As JSFmtConditions
Private m_clrForeColor              As OLE_COLOR
Private m_clrRowColorEven           As OLE_COLOR
Private m_clrRowColorOdd            As OLE_COLOR
Private m_oSortKeys                 As JSSortKeys
Private m_lCol                      As Long
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
Private m_lLeftCol                  As Long
Private WithEvents m_oColumnHeaderFont As StdFont
Attribute m_oColumnHeaderFont.VB_VarHelpID = -1
Private WithEvents m_oFont          As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_lFirstItem                As Long
Private m_eGridLines                As jgexGridLinesConstants
Private m_clrGridLinesColor         As OLE_COLOR
Private m_clrBackColorBkg           As OLE_COLOR
Private m_lCardSpacing              As Long
Private m_lCardWidth                As Long
Private m_lRow                      As Long
Private m_sErrorText                As String
Private m_bAllowEdit                As Boolean
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
Private m_lPreviewRowLines          As Long
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
Private m_hWndEdit                  As Long
'--- not backing property
Private m_hWndHScroll               As Long
Private m_bBandVisible              As Boolean
Private m_bRowHeightSet             As Boolean
Private m_bScrollUpdating           As Boolean
Private m_oRowData                  As JSRowData
Private m_pSubclassPic              As IUnknown
Private m_pSubclassCtl              As IUnknown
Private m_lSelAnchor                As Long
Private m_bCurRowDeselected         As Boolean
Private m_pDataModel                As IDataModel
Private m_aWindow()                 As JSRowData
Private m_lWindowFirst              As Long
Private m_lWindowCount              As Long
Private m_bWindowDirty              As Boolean
Private m_lVersion                  As Long
Private m_lHoldRowIndex             As Long
Private m_bInPendCommit             As Boolean
Private m_bInSet                    As Boolean
Private m_bEditing                  As Boolean
Private m_lEditRow                  As Long
Private m_lEditCol                  As Long
Private m_sEditOldValue             As String
Private m_bInEditSetup              As Boolean
Private m_bClickOpenedEdit          As Boolean
Private m_bDragSelect               As Boolean
Private m_hCurRowSel                As Long
Private m_bGridFocus                As Boolean
Private m_bEatClick                 As Boolean
Private m_oSizeCol                  As JSColumn
Private m_oDragCol                  As JSColumn
Private m_oDropCol                  As JSColumn
Private m_oDragGroup                As JSGroup
Private m_oDropGroup                As JSGroup
Private m_bDropAfter                As Boolean
Private m_lDragStartX               As Long
Private m_lDragStartY               As Long
Private m_bDragging                 As Boolean
Private m_bDropInGBox               As Boolean
Private m_lDragScroll               As Long
Private m_lSizeStartX               As Long
Private m_lSizeStartW               As Long
Private m_pSubclassEdit             As IUnknown
Private m_hBufDC                    As Long
Private m_hBufBmp                   As Long
Private m_hBufOldBmp                As Long
Private m_lBufLastY                 As Long
Private m_lBufLastRow               As Long
Private m_uIPAO                     As UcsIPAOHook

Private Type UcsNavLayout
    BandTop                 As Long
    BandH                   As Long
    Prefix                  As String
    Middle                  As String
    PrefixX                 As Long
    BtnFirst                As RECT
    BtnPrev                 As RECT
    Box                     As RECT
    MiddleX                 As Long
    BtnNext                 As RECT
    BtnLast                 As RECT
    Width                   As Long
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunction As String)
    PopRaiseError PushError, MODULE_NAME, sFunction
End Sub

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get FrozenColumns() As Integer
Attribute FrozenColumns.VB_Description = "Returns/sets the number of fixed columns at the left of the control."
    FrozenColumns = m_lFrozenColumns
End Property

Public Property Let FrozenColumns(ByVal nValue As Integer)
    m_lFrozenColumns = nValue
    pvInvalidate
End Property

Public Property Get RowHeight() As Long
    RowHeight = ToTwips(m_lRowHeight)
End Property

Public Property Let RowHeight(ByVal lValue As Long)
    '--- an explicit height survives later font changes
    m_lRowHeight = ToPixels(lValue)
    m_bRowHeightSet = True
    pvInvalidate
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns the handle of the cell editor."
    hWndEdit = m_hWndEdit
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
    pvInvalidate SkipScroll:=True
End Property

Public Property Get RowColorEven() As OLE_COLOR
Attribute RowColorEven.VB_Description = "Returns/sets the background color for even rows."
    RowColorEven = m_clrRowColorEven
End Property

Public Property Let RowColorEven(ByVal lValue As OLE_COLOR)
    m_clrRowColorEven = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get RowColorOdd() As OLE_COLOR
Attribute RowColorOdd.VB_Description = "Returns/sets the background color for odd rows."
    RowColorOdd = m_clrRowColorOdd
End Property

Public Property Let RowColorOdd(ByVal lValue As OLE_COLOR)
    m_clrRowColorOdd = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Returns the count of rows."
    '--- with grouping on this counts the group rows as well: the original
    '--- reports 7 for five rows in two groups, and Row lands on 2 for the
    '--- first record because position 1 is a group row
    pvSyncProjection
    RowCount = m_pDataModel.RowCount
End Property

Public Property Get RowSelected(ByVal RowPosition As Long) As Boolean
Attribute RowSelected.VB_Description = "Returns/set whether a row is selected or not."
    RowSelected = pvIsRowSelected(RowPosition)
End Property

Public Property Let RowSelected(ByVal RowPosition As Long, ByVal bValue As Boolean)
    If bValue Then
        If Not pvIsRowSelected(RowPosition) Then
            pvAddSel RowPosition
            RaiseEvent SelectionChange
            pvInvalidate SkipScroll:=True
        End If
    Else
        If pvIsRowSelected(RowPosition) Then
            m_oSelectedItems.RemoveRowPosition RowPosition
            RaiseEvent SelectionChange
            pvInvalidate SkipScroll:=True
        End If
    End If
End Property

Public Property Get SortKeys() As JSSortKeys
Attribute SortKeys.VB_Description = "Returns the JSSortKeys collection of the control."
    Set SortKeys = m_oSortKeys
End Property

Public Property Get Col() As Integer
Attribute Col.VB_Description = "Returns or sets the active column."
    Col = m_lCol
End Property

Public Property Let Col(ByVal nValue As Integer)
    Dim lLastCol        As Long
    Dim lMax            As Long
    Dim oCol            As JSColumn

    lMax = pvVisibleColCount()
    If nValue > lMax Then
        nValue = lMax
    End If
    If nValue < 1 And lMax > 0 Then
        nValue = 1
    End If
    If nValue >= 1 Then
        Set oCol = pvColByPosition(nValue)
        If oCol Is Nothing Then
            nValue = 0
        ElseIf Not oCol.Selectable Then
            nValue = 0
        End If
    End If
    If m_lCol <> nValue Then
        pvEditEnd
        lLastCol = m_lCol
        m_lCol = nValue
        pvInvalidate SkipScroll:=True
        RaiseEvent RowColChange(m_lRow, lLastCol)
    End If
End Property

Public Property Get BackColorRowGroup() As OLE_COLOR
Attribute BackColorRowGroup.VB_Description = "Returns/sets background color of the group rows."
    BackColorRowGroup = m_clrBackColorRowGroup
End Property

Public Property Let BackColorRowGroup(ByVal lValue As OLE_COLOR)
    m_clrBackColorRowGroup = lValue
    pvInvalidate SkipScroll:=True
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
    pvInvalidate
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
    pvInvalidate
End Property

Public Property Get BackColorGBBox() As OLE_COLOR
Attribute BackColorGBBox.VB_Description = "Returns/sets background color of the group by box."
    BackColorGBBox = m_clrBackColorGBBox
End Property

Public Property Let BackColorGBBox(ByVal lValue As OLE_COLOR)
    m_clrBackColorGBBox = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the grid and cards background color."
    BackColor = m_clrBackColor
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
    m_clrBackColor = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get ColumnHeaderHeight() As Long
Attribute ColumnHeaderHeight.VB_Description = "Returns/sets the height of the column header row."
    ColumnHeaderHeight = ToTwips(m_lColumnHeaderHeight)
End Property

Public Property Let ColumnHeaderHeight(ByVal lValue As Long)
    m_lColumnHeaderHeight = ToPixels(lValue)
    pvInvalidate
End Property

Public Property Get ColumnHeaders() As Boolean
Attribute ColumnHeaders.VB_Description = "Determines whether column headers are displayed."
    ColumnHeaders = m_bColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal bValue As Boolean)
    m_bColumnHeaders = bValue
    pvInvalidate
End Property

Public Property Get BackColorHeader() As OLE_COLOR
Attribute BackColorHeader.VB_Description = "Returns/sets background color of column and row headers."
    BackColorHeader = m_clrBackColorHeader
End Property

Public Property Let BackColorHeader(ByVal lValue As OLE_COLOR)
    m_clrBackColorHeader = lValue
    pvInvalidate SkipScroll:=True
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
    DefaultColumnWidth = ToTwips(m_lDefaultColumnWidth)
End Property

Public Property Let DefaultColumnWidth(ByVal lValue As Long)
    m_lDefaultColumnWidth = ToPixels(lValue)
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
    pvInvalidate
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
    '--- the record count is the one thing the model cannot see change on its
    '--- own, so setting it is what makes it re-read the source
    m_pDataModel.Refresh
    m_lWindowCount = 0
    m_bWindowDirty = True
    frNotifySortChanged
End Property

Public Property Get DataMode() As jgexDataModeConstants
Attribute DataMode.VB_Description = "Returns/sets a value representing the data retrieval mode."
    DataMode = m_eDataMode
End Property

Public Property Let DataMode(ByVal eValue As jgexDataModeConstants)
    If m_eDataMode <> eValue Then
        m_eDataMode = eValue
        pvCreateDataModel
    End If
End Property

Public Property Get LeftCol() As Integer
Attribute LeftCol.VB_Description = "Returns/sets the left-most visible column."
    LeftCol = m_lLeftCol
End Property

Public Property Let LeftCol(ByVal nValue As Integer)
    '--- scrolling stops once the last column reaches the right edge, so a
    '--- value past that clamps -- the frozen block is not scrolled over and
    '--- takes its width out of the strip the rest has to fill
    nValue = Clamp(nValue, 1, pvMaxLeftCol())
    If m_lLeftCol <> nValue Then
        m_lLeftCol = nValue
        pvInvalidate
        RaiseEvent LeftColChange
    End If
End Property

Public Property Get ColumnHeaderFont() As Font
Attribute ColumnHeaderFont.VB_Description = "Returns/sets  the font used in column headers."
    Set ColumnHeaderFont = m_oColumnHeaderFont
End Property

Public Property Set ColumnHeaderFont(ByVal oValue As Font)
    Set m_oColumnHeaderFont = oValue
    m_oColumnHeaderFont_FontChanged vbNullString
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal oValue As Font)
    Set m_oFont = oValue
    m_oFont_FontChanged vbNullString
End Property

Public Property Get FirstItem() As Long
Attribute FirstItem.VB_Description = "Returns/sets the row position of the first visible row or card."
    FirstItem = m_lFirstItem
End Property

Public Property Let FirstItem(ByVal lValue As Long)
    If lValue < 1 Then
        lValue = 1
    ElseIf lValue > RowCount And RowCount > 0 Then
        lValue = RowCount
    End If
    If m_lFirstItem <> lValue Then
        m_lFirstItem = lValue
        pvInvalidate
        RaiseEvent FirstItemChange
    End If
End Property

Public Property Get GridLines() As jgexGridLinesConstants
Attribute GridLines.VB_Description = "Determines whether the control will draw lines between cells."
    GridLines = m_eGridLines
End Property

Public Property Let GridLines(ByVal eValue As jgexGridLinesConstants)
    m_eGridLines = eValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get GridLinesColor() As OLE_COLOR
Attribute GridLinesColor.VB_Description = "Returns/sets the color used to draw grid lines."
    GridLinesColor = m_clrGridLinesColor
End Property

Public Property Let GridLinesColor(ByVal lValue As OLE_COLOR)
    m_clrGridLinesColor = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets the background color of the area outside the grid or card."
    BackColorBkg = m_clrBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal lValue As OLE_COLOR)
    m_clrBackColorBkg = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get CardSpacing() As Long
Attribute CardSpacing.VB_Description = "Returns/sets the horizontal and vertical space between cards."
    CardSpacing = ToTwips(m_lCardSpacing)
End Property

Public Property Let CardSpacing(ByVal lValue As Long)
    m_lCardSpacing = ToPixels(lValue)
End Property

Public Property Get CardWidth() As Long
Attribute CardWidth.VB_Description = "Returns/sets the width of a card."
    CardWidth = ToTwips(m_lCardWidth)
End Property

Public Property Let CardWidth(ByVal lValue As Long)
    m_lCardWidth = ToPixels(lValue)
End Property

Public Property Get RowBookmark(ByVal RowIndex As Long) As Variant
Attribute RowBookmark.VB_Description = "Returns/sets a value containing a bookmark for a row."
    RowBookmark = m_pDataModel.RowBookmark(RowIndex)
End Property

Public Property Let RowBookmark(ByVal RowIndex As Long, ByVal vntValue As Variant)
    m_pDataModel.RowBookmark(RowIndex) = vntValue
End Property

Public Property Get Row() As Long
    pvSyncProjection
    Row = m_lRow
End Property

Public Property Let Row(ByVal lValue As Long)
    '--- a value past the block lands on its nearest row, as Col does on its
    '--- last column -- without this the marquee leaves the block entirely and
    '--- the selection gains a row that is not there
    If RowCount = 0 Then
        Exit Property
    End If
    lValue = Clamp(lValue, 1, RowCount)
    If m_lRow <> lValue Then
        pvEditEnd
        pvEditCommit
        pvSetRow lValue
        '--- an assignment from outside collapses the selection onto the new
        '--- row, silently: navigation and drag go through pvSetRow instead
        '--- and apply their own selection, keeping SelectionChange ordering
        m_oSelectedItems.Clear
        pvAddSel m_lRow
        m_lSelAnchor = m_lRow
        pvInvalidate SkipScroll:=True
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
    '--- the current row's wrapper is writable for exactly as long as it is the
    '--- pending row, so the flag the buffer already carries is the answer
    If Not m_oRowData Is Nothing Then
        DataChanged = m_oRowData.frAllowUpdate
    End If
End Property

Public Property Let DataChanged(ByVal bValue As Boolean)
    If Not m_oRowData Is Nothing Then
        m_oRowData.frAllowUpdate = bValue
    End If
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the handle of the control."
    '--- the grid surface, as the original exposes: its inner window, not
    '--- the outer control that also holds the scrollbar band
    hWnd = picGrid.hWnd
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
    pvInvalidate SkipScroll:=True
End Property

Public Property Get ForeColorHeader() As OLE_COLOR
Attribute ForeColorHeader.VB_Description = "Returns/sets the foreground color used to display text in headers."
    ForeColorHeader = m_clrForeColorHeader
End Property

Public Property Let ForeColorHeader(ByVal lValue As OLE_COLOR)
    m_clrForeColorHeader = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get ForeColorRowGroup() As OLE_COLOR
Attribute ForeColorRowGroup.VB_Description = "Returns/sets the foreground color used to display text in group rows."
    ForeColorRowGroup = m_clrForeColorRowGroup
End Property

Public Property Let ForeColorRowGroup(ByVal lValue As OLE_COLOR)
    m_clrForeColorRowGroup = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get BackColorInfoText() As OLE_COLOR
Attribute BackColorInfoText.VB_Description = "Returns/sets the background color of the rectangle surrounding the information text displayed in the group by box "
    BackColorInfoText = m_clrBackColorInfoText
End Property

Public Property Let BackColorInfoText(ByVal lValue As OLE_COLOR)
    m_clrBackColorInfoText = lValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get ForeColorInfoText() As OLE_COLOR
Attribute ForeColorInfoText.VB_Description = "Returns/sets the foreground color used to display the information text in the group by box."
    ForeColorInfoText = m_clrForeColorInfoText
End Property

Public Property Let ForeColorInfoText(ByVal lValue As OLE_COLOR)
    m_clrForeColorInfoText = lValue
    pvInvalidate SkipScroll:=True
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
    If m_oRowData Is Nothing Then
        Exit Property
    End If
    AssignVariant Value, m_oRowData.Value(ColIndex)
End Property

Public Property Let Value(ByVal ColIndex As Integer, ByVal vntValue As Variant)
    '--- buffered on the current row rather than written through: the write
    '--- lands in the wrapper the page paints, so it shows at once, the row
    '--- moving commits it and Escape drops it. A group row has no record to
    '--- write to, and a column past the last one is not one to grow
    If m_oRowData Is Nothing Then
        Exit Property
    End If
    If m_oRowData.RowIndex <= 0 Or ColIndex < 1 Or ColIndex > m_oColumns.Count Then
        Exit Property
    End If
    m_oRowData.frAllowUpdate = True
    m_oRowData.Value(ColIndex) = vntValue
    pvInvalidate SkipScroll:=True
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
    pvInvalidate SkipScroll:=True
End Property

Public Property Get HeaderStyle() As jgexHeaderStyleConstants
Attribute HeaderStyle.VB_Description = "Returns/sets the display style for headers."
    HeaderStyle = m_eHeaderStyle
End Property

Public Property Let HeaderStyle(ByVal eValue As jgexHeaderStyleConstants)
    m_eHeaderStyle = eValue
    pvInvalidate
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
    pvSyncProjection
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
    RowExpanded = m_pDataModel.RowExpanded(RowPosition)
End Property

Public Property Let RowExpanded(ByVal RowPosition As Long, ByVal bValue As Boolean)
    m_pDataModel.RowExpanded(RowPosition) = bValue
    pvRecalcVisible
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
    pvInvalidate
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
    PreviewRowLines = m_lPreviewRowLines
End Property

Public Property Let PreviewRowLines(ByVal nValue As Integer)
    m_lPreviewRowLines = nValue
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
    frNotifySortChanged
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
    If Not m_bActAsDropDown Then
        Err.Raise 393 '--- available only when the control acts as dropdown
    End If
    BoundColumnIndex = m_vBoundColumnIndex
End Property

Public Property Let BoundColumnIndex(ByVal vntValue As Variant)
    If Not m_bActAsDropDown Then
        Err.Raise 393
    End If
    m_vBoundColumnIndex = vntValue
End Property

Public Property Get ReplaceColumnIndex() As Variant
Attribute ReplaceColumnIndex.VB_Description = "Returns/sets the index or key of the column that replaces Id values in a DropDown GridEX Control."
    If Not m_bActAsDropDown Then
        Err.Raise 393 '--- available only when the control acts as dropdown
    End If
    ReplaceColumnIndex = m_vReplaceColumnIndex
End Property

Public Property Let ReplaceColumnIndex(ByVal vntValue As Variant)
    If Not m_bActAsDropDown Then
        Err.Raise 393
    End If
    m_vReplaceColumnIndex = vntValue
End Property

Public Property Get GridLineStyle() As jgexGridLineStyleConstants
Attribute GridLineStyle.VB_Description = "Determines the grid lines style."
    GridLineStyle = m_eGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal eValue As jgexGridLineStyleConstants)
    m_eGridLineStyle = eValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get EmptyRows() As Boolean
Attribute EmptyRows.VB_Description = "Determines whether empty rows below the last row should be displayed. "
    EmptyRows = m_bEmptyRows
End Property

Public Property Let EmptyRows(ByVal bValue As Boolean)
    m_bEmptyRows = bValue
    pvInvalidate SkipScroll:=True
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Determines whether drawing is enabled or not."
    Redraw = m_bRedraw
End Property

Public Property Let Redraw(ByVal bValue As Boolean)
    If m_bRedraw = bValue Then
        Exit Property
    End If
    m_bRedraw = bValue
    '--- everything suppressed while it was off lands in a single repaint
    If bValue Then
        pvInvalidate
    End If
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
    pvInvalidate
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

'= friend ================================================================

Friend Property Get frDefaultColumnWidthPx() As Long
    frDefaultColumnWidthPx = m_lDefaultColumnWidth
End Property

'= private ===============================================================

Private Property Get pvAddressOfSubclassProc() As GridEX
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
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
    m_pDataModel.Refresh
    pvApplyHoldSort HoldSortSettings
    m_lWindowCount = 0
    m_bWindowDirty = True
    pvInvalidate
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
    m_pDataModel.SetAllExpanded False
    pvRecalcVisible
End Sub

Public Sub RefreshSort()
Attribute RefreshSort.VB_Description = "Forces the re-sort of the records."
    '--- the collections mark the order stale as they are edited, so this only
    '--- has to bring the rebuild forward from the next paint
    m_pDataModel.RefreshSort
    pvSyncProjection
    pvInvalidate
End Sub

Public Function RowIndex(ByVal RowPosition As Long) As Long
Attribute RowIndex.VB_Description = "Returns the original index of a row."
    Dim lRowIndex       As Long

    lRowIndex = m_pDataModel.RowIndex(RowPosition)
    If lRowIndex > 0 Then
        RowIndex = lRowIndex
    End If
End Function

Public Sub RefreshGroups(Optional ByVal AllCollapsed As Boolean)
Attribute RefreshGroups.VB_Description = "Forces recalculation of groups."
    '--- the groups are rebuilt from scratch, which resets every expand box
    '--- to DefaultGroupMode unless this call overrides it
    m_pDataModel.RefreshGroups AllCollapsed
    pvSyncProjection
    pvInvalidate
End Sub

Public Sub EnsureVisible(Optional ByVal Row As Long, Optional ByVal Col As Integer)
Attribute EnsureVisible.VB_Description = "Ensures visibility of a cell."
    Dim lVisible        As Long
    Dim oCol            As JSColumn

    '--- no arg means the current row; the row comes in first so the column
    '--- walk below asks about a cell that is on the page
    If Row < 1 Then
        Row = m_lRow
    End If
    If Row < 1 Or RowCount = 0 Then
        Exit Sub
    End If
    lVisible = pvVisibleRows()
    If Row < m_lFirstItem Then
        FirstItem = Row
    ElseIf Row > m_lFirstItem + lVisible - 1 Then
        FirstItem = Row - lVisible + 1
    End If
    If Col >= 1 Then
        Set oCol = pvColByPosition(Col)
        If Not oCol Is Nothing Then
            pvEnsureColVisible Row, oCol.Index
        End If
    End If
End Sub

Public Sub Rebind(Optional HoldSortSettings As Variant)
Attribute Rebind.VB_Description = "Forces re-creation of the recordset."
    m_bInSet = True
    m_pDataModel.Refresh
    pvApplyHoldSort HoldSortSettings
    m_lWindowCount = 0
    m_bWindowDirty = True
    '--- binding positions the current cell on the first row/column and
    '--- selects it (matches the original after a bind)
    m_oSelectedItems.Clear
    If RowCount > 0 Then
        m_lRow = 1
        m_lFirstItem = 1
        pvAddSel 1
        m_lSelAnchor = 1
    Else
        m_lRow = 0
        m_lFirstItem = 0
        m_lSelAnchor = 0
    End If
    pvSyncRowData
    '--- a bind starts with the whole row selected rather than a cell in it,
    '--- which is what Col = 0 means -- the original's first RowColChange
    '--- reports LastCol=0 for exactly that reason
    m_lCol = 0
    m_bInSet = False
    '--- a bind places the marquee itself, so nothing is carried over: the
    '--- hold is dropped first, or the sync below would put the row back on
    '--- whatever record it was on before the bind
    m_lHoldRowIndex = 0
    pvSyncProjection
    pvInvalidate
End Sub

Public Function IsGroupItem(ByVal Row As Long) As Boolean
Attribute IsGroupItem.VB_Description = "Returns True if the specified row is a group row."
    IsGroupItem = (m_pDataModel.RowIndex(Row) < 0)
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a GridEX control."
    pvInvalidate
End Sub

Public Sub HoldFields()
Attribute HoldFields.VB_Description = "Prevents recalculation of the column layout when the recordset is created."
End Sub

Public Sub ClearFields()
Attribute ClearFields.VB_Description = "Restores the default layout columns."
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Displays the About box for the control."
Attribute AboutBox.VB_UserMemId = -552
End Sub

Public Sub LoadEntireRecordset()
Attribute LoadEntireRecordset.VB_Description = "Loads the bookmarks of all records in the underlying Recordset."
End Sub

Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Expands all group rows."
    m_pDataModel.SetAllExpanded True
    pvRecalcVisible
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
    Dim oCancel         As JSRetBoolean
    Dim oItem           As JSSelectedItem
    Dim aIdx()          As Long
    Dim avBmk()         As Variant
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lLand           As Long

    '--- probed with Del in row mode: one BeforeDelete for the operation, a
    '--- vetoable BeforeDeleteEX per record, and the reprojection lands the
    '--- currency from nowhere -- RowColChange carries LastRow=0 -- before
    '--- AfterDelete closes it out
    If RowCount = 0 Or m_oSelectedItems.Count = 0 Then
        Exit Sub
    End If
    ReDim aIdx(0 To m_oSelectedItems.Count) As Long
    ReDim avBmk(0 To m_oSelectedItems.Count) As Variant
    For Each oItem In m_oSelectedItems
        If oItem.frInitRowIndex > 0 Then
            aIdx(lCount) = oItem.frInitRowIndex
            '--- from the model rather than the item: an unbound selection is
            '--- collected before its bookmarks exist, and the original hands
            '--- BeforeDeleteEX a live one -- probed, a record the client never
            '--- bookmarked answers with its row index
            AssignVariant avBmk(lCount), m_pDataModel.RowBookmark(aIdx(lCount))
            If DataIsBlank(avBmk(lCount)) Then
                avBmk(lCount) = aIdx(lCount)
            End If
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        Exit Sub
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeDelete(oCancel)
    If oCancel.Value Then
        Exit Sub
    End If
    pvEditEnd bCancel:=True
    '--- the hold as well: a delete renumbers, so the held index would remap
    '--- to whatever record slid into the hole and re-land the currency early
    lLand = m_lRow
    m_lRow = 0
    m_lHoldRowIndex = 0
    For lIdx = 0 To lCount - 1
        Set oCancel = New JSRetBoolean
        RaiseEvent BeforeDeleteEX(aIdx(lIdx), avBmk(lIdx), oCancel)
        If Not oCancel.Value Then
            lPos = m_pDataModel.GetRowPosition(aIdx(lIdx))
            If lPos >= 1 Then
                m_pDataModel.Delete lPos
            End If
        End If
    Next
    m_oSelectedItems.Clear
    m_lWindowCount = 0
    m_bWindowDirty = True
    pvSyncProjection
    lLand = pvClampRow(lLand)
    pvPopulateWindow
    If lLand >= 1 Then
        pvUpdateSelection lLand, 0, False
        pvSetRow lLand
    End If
    pvUpdateScrollBars
    pvInvalidate
    RaiseEvent AfterDelete
End Sub

Public Sub SearchNewRecords()
Attribute SearchNewRecords.VB_Description = "Searches for records added after all bookmarks have been loaded."
End Sub

Public Sub RefreshRowBookmark(ByVal Bookmark As Variant)
Attribute RefreshRowBookmark.VB_Description = "Refreshes data of the record that matches the Bookmark."
    Dim lRowIndex       As Long

    lRowIndex = m_pDataModel.GetRowIndex(Bookmark)
    If lRowIndex = 0 Then
        '--- the original leaves the current row alone and says so
        Err.Raise vbObjectError + 119, , "Not a valid Bookmark."
    End If
    m_pDataModel.RefreshRowIndex lRowIndex
    m_lWindowCount = 0
    m_bWindowDirty = True
    pvInvalidate
End Sub

Public Sub RefreshRowIndex(ByVal RowIndex As Long)
Attribute RefreshRowIndex.VB_Description = "Refreshes data of the record that matches the index."
    m_pDataModel.RefreshRowIndex RowIndex
    m_lWindowCount = 0
    m_bWindowDirty = True
    pvInvalidate
End Sub

Public Function GroupRowLevel(ByVal RowPosition As Long) As Integer
Attribute GroupRowLevel.VB_Description = "Returns the level of a group row."
    Dim oRowData        As JSRowData

    Set oRowData = pvRowDataAt(RowPosition)
    If Not oRowData Is Nothing Then
        GroupRowLevel = oRowData.GroupLevel
    End If
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
    Set GetRowData = pvRowDataAt(RowPosition)
    If GetRowData Is Nothing Then
        Set GetRowData = New JSRowData
        GetRowData.frReset m_oColumns.Count
    End If
End Function

'= friend ================================================================

Friend Sub frNotifySortChanged()
    m_pDataModel.RefreshGroups False
    m_lWindowCount = 0
    m_bWindowDirty = True
    If Not m_bInSet Then
        pvInvalidate
    End If
End Sub

Friend Sub frNotifyGroupsChanged()
    frNotifySortChanged
    pvSyncProjection
    pvScrollCurrentToTop
End Sub

Friend Function frColIsGrouped(ByVal lColIndex As Long) As Boolean
    Dim oGroup          As JSGroup

    For Each oGroup In m_oGroups
        If oGroup.ColIndex = lColIndex Then
            frColIsGrouped = True
            Exit For
        End If
    Next
End Function

Friend Function frColSortOrder(ByVal lColIndex As Long) As jgexSortOrderConstants
    Dim oGroup          As JSGroup
    Dim oKey            As JSSortKey

    For Each oGroup In m_oGroups
        If oGroup.ColIndex = lColIndex Then
            frColSortOrder = oGroup.SortOrder
            Exit Function
        End If
    Next
    For Each oKey In m_oSortKeys
        If oKey.ColIndex = lColIndex Then
            frColSortOrder = oKey.SortOrder
            Exit Function
        End If
    Next
End Function

Friend Sub frColMove(oCol As JSColumn, ByVal lNewPos As Long)
    Dim lOldPos         As Long
    Dim oItem           As JSColumn

    lNewPos = Clamp(lNewPos, 1, m_oColumns.Count)
    lOldPos = oCol.ColPosition
    If lNewPos = lOldPos Then
        Exit Sub
    End If
    For Each oItem In m_oColumns
        If oItem.ColPosition > lOldPos And oItem.ColPosition <= lNewPos Then
            oItem.frColPosition = oItem.ColPosition - 1
        ElseIf oItem.ColPosition >= lNewPos And oItem.ColPosition < lOldPos Then
            oItem.frColPosition = oItem.ColPosition + 1
        End If
    Next
    oCol.frColPosition = lNewPos
End Sub

Friend Sub frColAutoSize(oCol As JSColumn)
    Dim oCancel         As JSRetBoolean
    Dim lWidth          As Long

    lWidth = pvAutoSizeWidth(oCol.Index)
    If lWidth < MIN_COL_W Then
        lWidth = MIN_COL_W
    End If
    m_lSizeStartW = pvColWidth(oCol)
    oCol.frWidthPx = lWidth
    lWidth = oCol.Width
    oCol.frWidthPx = m_lSizeStartW
    Set oCancel = New JSRetBoolean
    RaiseEvent ColResize(oCol.Index, lWidth, oCancel)
    If Not oCancel.Value Then
        oCol.Width = lWidth
    End If
    pvInvalidate
End Sub

Friend Sub frRaiseUnboundReadData(ByVal lRowIndex As Long, vBookmark As Variant, oValues As JSRowData)
    RaiseEvent UnboundReadData(lRowIndex, vBookmark, oValues)
End Sub

Friend Sub frRaiseUnboundAddNew(oNewRowBookmark As JSRetVariant, oValues As JSRowData)
    RaiseEvent UnboundAddNew(oNewRowBookmark, oValues)
End Sub

Friend Sub frRaiseUnboundUpdate(ByVal lRowIndex As Long, vBookmark As Variant, oValues As JSRowData)
    If m_bInPendCommit Then
        Exit Sub
    End If
    RaiseEvent UnboundUpdate(lRowIndex, vBookmark, oValues)
End Sub

Friend Sub frRaiseUnboundDelete(ByVal lRowIndex As Long, vBookmark As Variant)
    RaiseEvent UnboundDelete(lRowIndex, vBookmark)
End Sub

Friend Function frRaiseFetchData(ByVal lRowIndex As Long, ByVal lColIndex As Long, vBookmark As Variant) As Variant
    Dim oValue          As JSRetVariant

    Set oValue = New JSRetVariant
    RaiseEvent FetchData(lRowIndex, lColIndex, vBookmark, oValue)
    AssignVariant frRaiseFetchData, oValue.Value
End Function

Friend Function frRaiseFetchIcon(ByVal lRowIndex As Long, ByVal lColIndex As Long, vBookmark As Variant) As Long
    Dim oIconIndex      As JSRetInteger

    Set oIconIndex = New JSRetInteger
    RaiseEvent FetchIcon(lRowIndex, lColIndex, vBookmark, oIconIndex)
    frRaiseFetchIcon = oIconIndex.Value
End Function

Friend Function frGetRowIndex(vBookmark As Variant) As Long
    frGetRowIndex = m_pDataModel.GetRowIndex(vBookmark)
End Function

Friend Function frGetRowPosition(ByVal lRowIndex As Long) As Long
    pvSyncProjection
    frGetRowPosition = m_pDataModel.GetRowPosition(lRowIndex)
End Function

Friend Function frBeforeTranslateAccel(uMsg As APIMSG) As Boolean
    Dim nKeyCode        As Integer

    If uMsg.Message <> WM_KEYDOWN Then
        Exit Function
    End If
    nKeyCode = LoWordInt(uMsg.wParam)
    If uMsg.hWnd = hWnd Or uMsg.hWnd = UserControl.hWnd Then
        Select Case nKeyCode
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyEscape, vbKeyReturn
            frBeforeTranslateAccel = True
        Case vbKeyTab
            If (pvShiftState() And vbCtrlMask) <> 0 Then
                frBeforeTranslateAccel = pvSiteTranslateKey(vbKeyTab, pvShiftState() And vbShiftMask)
                Exit Function
            End If
            frBeforeTranslateAccel = (m_eTabKeyBehavior = jgexColumnNavigation) And m_bAllowEdit
        End Select
    ElseIf uMsg.hWnd = m_hWndEdit And m_hWndEdit <> 0 Then
        Select Case nKeyCode
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight, vbKeyReturn, vbKeyEscape
            frBeforeTranslateAccel = True
        Case vbKeyTab
            If (pvShiftState() And vbCtrlMask) <> 0 Then
                pvEditEnd
                frBeforeTranslateAccel = pvSiteTranslateKey(vbKeyTab, pvShiftState() And vbShiftMask)
                Exit Function
            End If
            frBeforeTranslateAccel = (m_eTabKeyBehavior = jgexColumnNavigation)
        End Select
    End If
    If frBeforeTranslateAccel Then
        If uMsg.hWnd = m_hWndEdit And m_hWndEdit <> 0 Then
            Call SendMessage(m_hWndEdit, WM_KEYDOWN, uMsg.wParam, uMsg.lParam)
        Else
            Call SendMessage(hWnd, WM_KEYDOWN, uMsg.wParam, uMsg.lParam)
        End If
    End If
End Function

Friend Sub frAfterTranslateAccel(uMsg As APIMSG)
End Sub

'= private ===============================================================

Private Function pvSiteTranslateKey(ByVal lKeyCode As Long, ByVal lShift As Long) As Boolean
    Const FUNC_NAME     As String = "pvSiteTranslateKey"
    Const S_OK                                  As Long = 0
    Const S_FALSE                               As Long = 1
    Dim pOleObject      As IOleObject
    Dim pOleControlSite As IOleControlSite
    Dim uMsg            As APIMSG
    Dim oParent         As Object

    On Error GoTo EH
    With uMsg
        .hWnd = UserControl.hWnd
        .Message = WM_KEYDOWN
        .wParam = lKeyCode
        .lParam = &H10001
    End With
    If Not TypeOf Me Is IOleObject Then
        Exit Function
    End If
    Set pOleObject = Me
    If pOleObject.GetClientSite(pOleControlSite) <> S_OK Then
        Exit Function
    End If
    If pOleControlSite Is Nothing Then
        Exit Function
    End If
    If pOleControlSite.TranslateAccelerator(VarPtr(uMsg), lShift) = S_OK Then
        pvSiteTranslateKey = True
        Exit Function
    End If
    Set oParent = UserControl.Parent
    If oParent Is Nothing Then
        Exit Function
    End If
    If Not TypeOf oParent Is IOleObject Then
        Exit Function
    End If
    Set pOleObject = oParent
    Set pOleControlSite = Nothing
    If pOleObject.GetClientSite(pOleControlSite) <> S_OK Then
        Exit Function
    End If
    If pOleControlSite Is Nothing Then
        Exit Function
    End If
    pvSiteTranslateKey = (pOleControlSite.TranslateAccelerator(VarPtr(uMsg), lShift) = S_OK)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvCreateDataModel()
    Dim oUnbound        As cUnboundDataModel
    Dim oAdo            As cAdoDataModel

    If Not m_pDataModel Is Nothing Then
        m_pDataModel.Terminate
        Set m_pDataModel = Nothing
    End If
    '--- jgexDAO gets the unbound model too: DAO binding is out of scope, and
    '--- a control in that mode with no recordset behaves as an empty unbound
    '--- one rather than raising
    Select Case m_eDataMode
    Case jgexADO
        Set oAdo = New cAdoDataModel
        oAdo.frInit Me
        Set m_pDataModel = oAdo
    Case Else
        Set oUnbound = New cUnboundDataModel
        oUnbound.frInit Me
        Set m_pDataModel = oUnbound
    End Select
End Sub

Private Sub pvPopulateWindow()
    Dim lFirst          As Long
    Dim lCount          As Long
    Dim aPrev()         As JSRowData
    Dim lPrevCount      As Long
    Dim lIdx            As Long
    Dim lRowIndex         As Long
    Dim oCarry          As JSRowData
    Dim lJdx            As Long

    lFirst = m_lFirstItem
    If lFirst < 1 Then
        lFirst = 1
    End If
    '--- one past the bottom, so the partly shown row a scroll leaves at the
    '--- edge is buffered like any other
    lCount = pvVisibleRows() + 1
    If lFirst + lCount - 1 > RowCount Then
        lCount = RowCount - lFirst + 1
    End If
    If lCount < 0 Then
        lCount = 0
    End If
    If Not m_bWindowDirty Then
        If lFirst = m_lWindowFirst And lCount = m_lWindowCount Then
            Exit Sub
        End If
    End If
    If m_lWindowCount > 0 Then
        aPrev = m_aWindow
        lPrevCount = m_lWindowCount
    End If
    Erase m_aWindow
    m_bWindowDirty = False
    m_lWindowFirst = lFirst
    m_lWindowCount = lCount
    If lCount = 0 Then
        Exit Sub
    End If
    ReDim m_aWindow(1 To lCount) As JSRowData
    For lIdx = 1 To lCount
        '--- the row being edited keeps the instance it is buffered in:
        '--- re-reading it would drop what it holds, and a scroll does not
        '--- commit a row
        If lFirst + lIdx - 1 = m_lRow And DataChanged Then
            Set m_aWindow(lIdx) = m_oRowData
            If lPrevCount = 0 Then
                RaiseEvent RowFormat(m_aWindow(lIdx))
            End If
        Else
            Set oCarry = Nothing
            If lPrevCount > 0 Then
                lRowIndex = m_pDataModel.RowIndex(lFirst + lIdx - 1)
                If lRowIndex <> 0 Then
                    For lJdx = 1 To lPrevCount
                        If Not aPrev(lJdx) Is Nothing Then
                            If aPrev(lJdx).frInitRowIndex = lRowIndex Then
                                Set oCarry = aPrev(lJdx)
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            If Not oCarry Is Nothing Then
                Set m_aWindow(lIdx) = oCarry
            Else
                Set m_aWindow(lIdx) = m_pDataModel.GetRowData(lFirst + lIdx - 1)
                RaiseEvent RowFormat(m_aWindow(lIdx))
            End If
        End If
    Next
    pvSyncRowData
End Sub

Private Sub pvRemapCurrent()
    Dim lPos            As Long
    Dim oItem           As JSSelectedItem

    If m_lHoldRowIndex <> 0 Then
        lPos = m_pDataModel.GetRowPosition(m_lHoldRowIndex)
        If lPos >= 1 Then
            m_lRow = lPos
            pvSyncRowData
        End If
    End If
    For Each oItem In m_oSelectedItems
        oItem.frRowPosition = m_pDataModel.GetRowPosition(oItem.frInitRowIndex)
    Next
End Sub

Private Sub pvSyncProjection()
    Dim lCount          As Long
    Dim lVersion        As Long
    Dim lRowIndex       As Long

    '--- reading the count is what forces a stale projection to rebuild
    lCount = m_pDataModel.RowCount
    lVersion = m_pDataModel.Version
    If lVersion <> m_lVersion Then
        m_lVersion = lVersion
        pvRemapCurrent
        m_lWindowCount = 0
        m_bWindowDirty = True
    End If
    lRowIndex = m_pDataModel.RowIndex(m_lRow)
    If lRowIndex <> 0 Then
        m_lHoldRowIndex = lRowIndex
    End If
End Sub

Private Sub pvSyncRowData()
    Dim oRowData        As JSRowData

    Set oRowData = pvRowDataAt(m_lRow)
    If Not oRowData Is m_oRowData And DataChanged Then
        Err.Raise vbObjectError, , "Internal error: DataChanged=" & DataChanged
    End If
    Set m_oRowData = oRowData
End Sub

Private Function pvRowDataAt(ByVal lRowPosition As Long) As JSRowData
    Set pvRowDataAt = pvWindowRow(lRowPosition)
    If pvRowDataAt Is Nothing Then
        If lRowPosition >= 1 And lRowPosition <= m_pDataModel.RowCount Then
            Set pvRowDataAt = m_pDataModel.GetRowData(lRowPosition)
        End If
    End If
End Function

Private Function pvWindowRow(ByVal lPos As Long) As JSRowData
    If lPos >= m_lWindowFirst And lPos < m_lWindowFirst + m_lWindowCount Then
        Set pvWindowRow = m_aWindow(lPos - m_lWindowFirst + 1)
    End If
End Function

Private Sub pvRecalcVisible()
    Dim lMax            As Long

    lMax = m_pDataModel.RowCount - pvVisibleRows() + 1
    If lMax < 1 Then
        lMax = 1
    End If
    If m_lFirstItem > lMax Then
        FirstItem = lMax
    End If
    pvRemapCurrent
    m_bWindowDirty = True
    pvInvalidate
End Sub

Private Sub pvScrollCurrentToTop()
    Dim lMax            As Long

    If m_lRow < 1 Or m_pDataModel.RowCount = 0 Then
        Exit Sub
    End If
    lMax = m_pDataModel.RowCount - pvVisibleRows() + 1
    If lMax < 1 Then
        lMax = 1
    End If
    FirstItem = Clamp(m_lRow, 1, lMax)
End Sub

Private Function pvIsGroupRow(ByVal lPos As Long) As Boolean
    pvIsGroupRow = (m_pDataModel.RowIndex(lPos) < 0)
End Function

Private Sub pvClearCol()
    Dim lLastCol        As Long

    If m_lCol <> 0 Then
        pvEditEnd
        lLastCol = m_lCol
        m_lCol = 0
        pvInvalidate SkipScroll:=True
        RaiseEvent RowColChange(m_lRow, lLastCol)
    End If
End Sub

Private Function pvRowSelZoneAt(ByVal lX As Long, ByVal lY As Long) As Boolean
    Dim lZone           As Long
    Dim lRow            As Long

    '--- the row headers when they show, a probed 5 pixel strip when they do
    '--- not: either way the strip belongs to the row as a whole
    lZone = pvRowHeaderWidth()
    If lZone = 0 Then
        lZone = ROWSEL_ZONE_W
    End If
    If lX < 0 Or lX >= lZone Then
        Exit Function
    End If
    If m_lRowHeight <= 0 Or lY < pvRowsTop() Then
        Exit Function
    End If
    lRow = m_lFirstItem + (lY - pvRowsTop()) \ m_lRowHeight
    If lRow >= 1 And lRow <= RowCount Then
        pvRowSelZoneAt = Not pvIsGroupRow(lRow)
    End If
End Function

Private Function pvIsGroupHeader(ByVal lPos As Long) As Boolean
    Dim oRowData        As JSRowData

    '--- a header and not a footer, which shares the negative row index
    If pvIsGroupRow(lPos) Then
        pvPopulateWindow
        Set oRowData = pvWindowRow(lPos)
        If Not oRowData Is Nothing Then
            pvIsGroupHeader = (oRowData.RowType = jgexRowTypeGroupHeader)
        End If
    End If
End Function

Private Function pvIsBlank(vValue As Variant) As Boolean
    Select Case VarType(vValue)
    Case vbEmpty, vbNull, vbError
        pvIsBlank = True
    Case vbString
        pvIsBlank = (LenB(vValue) = 0)
    End Select
End Function

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

'--- painting

Private Sub pvOnPaint(ByVal hWnd As Long)
    Const FUNC_NAME     As String = "pvOnPaint"
    Dim uPS             As PAINTSTRUCT
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lTop            As Long
    Dim hMemDC          As Long

    On Error GoTo EH
    Call BeginPaint(hWnd, uPS)
    If hWnd = picGrid.hWnd Then
        pvPaint uPS.hDC, uPS.rcPaint
    Else
        '--- the band along the bottom edge is the whole of what this window
        '--- paints, so a repaint that does not reach it has nothing to do
        lWidth = UserControl.ScaleWidth
        lHeight = GetSystemMetrics(SM_CYHSCROLL)
        lTop = UserControl.ScaleHeight - lHeight
        If pvNeedRepaint(uPS.rcPaint, lTop, lHeight) Then
            hMemDC = pvBufferInit(uPS.hDC, lWidth, lHeight)
            pvBufferBand hMemDC, lTop, lWidth, lHeight
            pvPaintNavigator hMemDC
            pvBufferFlush uPS.hDC, hMemDC, lTop, lWidth, lHeight
            pvBufferTermiante
        End If
    End If
QH:
    Call EndPaint(hWnd, uPS)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvPaint(ByVal hDC As Long, uClip As RECT)
    Dim lY              As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hMemDC          As Long
    Dim lBandH          As Long

    pvSyncProjection
    lWidth = picGrid.ScaleWidth
    lHeight = pvGroupByBoxHeight()
    If m_bColumnHeaders And m_lColumnHeaderHeight > lHeight Then
        lHeight = m_lColumnHeaderHeight
    End If
    If m_lRowHeight + 1 > lHeight Then
        lHeight = m_lRowHeight + 1
    End If
    hMemDC = pvBufferInit(hDC, lWidth, lHeight)
    If m_bGroupByBoxVisible Then
        lBandH = pvGroupByBoxHeight()
        If pvNeedRepaint(uClip, lY, lBandH) Then
            pvBufferBand hMemDC, lY, lWidth, lBandH
            pvPaintGroupByBox hMemDC
            pvBufferFlush hDC, hMemDC, lY, lWidth, lBandH
        End If
        lY = lY + lBandH
    End If
    If m_bColumnHeaders Then
        lBandH = m_lColumnHeaderHeight
        If pvNeedRepaint(uClip, lY, lBandH) Then
            pvBufferBand hMemDC, lY, lWidth, lBandH
            pvPaintHeaders hMemDC, lY
            pvBufferFlush hDC, hMemDC, lY, lWidth, lBandH
        End If
        lY = lY + lBandH
    End If
    pvPaintRows hDC, hMemDC, lY, uClip
    pvBufferTermiante
End Sub

Private Sub pvPaintGroupByBox(ByVal hDC As Long)
    Dim lBoxH           As Long
    Dim lTotalH         As Long
    Dim uRect           As RECT
    Dim hPrevFont       As Long
    Dim uMetrics        As TEXTMETRICW
    Dim lIdx            As Long
    Dim oGroup          As JSGroup
    Dim sCaption        As String
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lChipW          As Long
    Dim lChipH          As Long
    Dim lElbowTop       As Long

    hPrevFont = pvSelectFont(hDC, m_oFont)
    lBoxH = m_lColumnHeaderHeight + 4
    lTotalH = pvGroupByBoxHeight()
    pvFillRect hDC, 0, 0, picGrid.ScaleWidth, lTotalH, m_clrBackColorGBBox
    If m_oGroups.Count > 0 Then
        uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
        pvLayoutGroupChips hDC:=hDC
        For Each oGroup In m_oGroups
            lIdx = lIdx + 1
            If oGroup.ColIndex >= 1 And oGroup.ColIndex <= m_oColumns.Count Then
                sCaption = m_oColumns.Item(oGroup.ColIndex).Caption
                lLeft = oGroup.frChipRect.Left
                lTop = oGroup.frChipRect.Top
                lChipW = oGroup.frChipRect.Right - lLeft
                lChipH = oGroup.frChipRect.Bottom - lTop
                pvFillRect hDC, lLeft, lTop, lLeft + lChipW, lTop + lChipH, m_clrBackColorHeader
                pvLine hDC, lLeft, lTop, lLeft + lChipW - 1, lTop, vb3DHighlight, PS_SOLID
                pvLine hDC, lLeft, lTop, lLeft, lTop + lChipH - 1, vb3DHighlight, PS_SOLID
                pvLine hDC, lLeft, lTop + lChipH - 2, lLeft + lChipW - 1, lTop + lChipH - 2, vb3DShadow, PS_SOLID
                pvLine hDC, lLeft + lChipW - 2, lTop, lLeft + lChipW - 2, lTop + lChipH - 1, vb3DShadow, PS_SOLID
                pvLine hDC, lLeft, lTop + lChipH - 1, lLeft + lChipW, lTop + lChipH - 1, vb3DDKShadow, PS_SOLID
                pvLine hDC, lLeft + lChipW - 1, lTop, lLeft + lChipW - 1, lTop + lChipH, vb3DDKShadow, PS_SOLID
                pvDrawText hDC, sCaption, lLeft + 2, lTop + 1, lLeft + lChipW - 2, lTop + lChipH, _
                    m_clrForeColorHeader, m_clrBackColorHeader, jgexAlignLeft, lLeft + 2, lLeft + lChipW - 2
                If oGroup.SortOrder <> 0 Then
                    pvPaintSortGlyph hDC, lLeft + 2 + pvTextWidth(hDC, sCaption) + 4, _
                        lTop + 3 + uMetrics.tmHeight \ 2 + 4, oGroup.SortOrder
                End If
                If oGroup Is m_oDragGroup Then
                    Call PatBlt(hDC, lLeft, lTop + 1, lChipW - 1, lChipH - 2, DSTINVERT)
                End If
                If lIdx < m_oGroups.Count Then
                    lElbowTop = lTop + m_lColumnHeaderHeight \ 2 + CLng((3 * lChipH - 4) / 4)
                    pvLine hDC, lLeft + lChipW - 5, lTop + lChipH, lLeft + lChipW - 5, lElbowTop + 1, vb3DDKShadow, PS_SOLID
                    pvLine hDC, lLeft + lChipW - 5, lElbowTop, lLeft + lChipW + CHIP_GAP, lElbowTop, vb3DDKShadow, PS_SOLID
                End If
            End If
        Next
    Else
        Call DrawText(hDC, StrPtr(m_sGroupByBoxInfoText), Len(m_sGroupByBoxInfoText), uRect, DT_SINGLELINE Or DT_CALCRECT)
        pvFillRect hDC, 4, 5, 12 + uRect.Right, 5 + lBoxH, m_clrBackColorInfoText
        pvDrawText hDC, m_sGroupByBoxInfoText, 7, 5, 7 + uRect.Right, 5 + lBoxH, m_clrForeColorInfoText, m_clrBackColorInfoText, jgexAlignLeft, 7, 7 + uRect.Right
    End If
    If m_bDropInGBox And pvGBoxDropMark(lLeft, lTop) Then
        lChipH = FontTextMetrics(m_oColumnHeaderFont, hDC).tmHeight + 6
        pvFillRect hDC, lLeft - 2, lTop, lLeft + 1, lTop + lChipH, vbRed
    End If
    Call SelectObject(hDC, hPrevFont)
End Sub

Private Sub pvLayoutGroupChips(Optional ByVal hDC As Long)
    Dim hOwnDC          As Long
    Dim hPrevFont       As Long
    Dim uMetrics        As TEXTMETRICW
    Dim lChipH          As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim oGroup          As JSGroup
    Dim uRect           As RECT

    If hDC = 0 Then
        hOwnDC = GetDC(picGrid.hWnd)
        hDC = hOwnDC
        hPrevFont = pvSelectFont(hDC, m_oFont)
    End If
    uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
    lChipH = uMetrics.tmHeight + 6
    lLeft = CHIP_LEFT
    lTop = CHIP_TOP
    For Each oGroup In m_oGroups
        uRect.Left = 0
        uRect.Top = 0
        uRect.Right = 0
        uRect.Bottom = 0
        If oGroup.ColIndex >= 1 And oGroup.ColIndex <= m_oColumns.Count Then
            uRect.Left = lLeft
            uRect.Top = lTop
            uRect.Right = lLeft + pvTextWidth(hDC, m_oColumns.Item(oGroup.ColIndex).Caption) + CHIP_PAD
            uRect.Bottom = lTop + lChipH
            lLeft = uRect.Right + CHIP_GAP
            lTop = lTop + m_lColumnHeaderHeight \ 2
        End If
        oGroup.frChipRect = uRect
    Next
    If hOwnDC <> 0 Then
        Call SelectObject(hDC, hPrevFont)
        Call ReleaseDC(picGrid.hWnd, hOwnDC)
    End If
End Sub

Private Sub pvPaintHeaders(ByVal hDC As Long, ByVal lY As Long)
    Dim lHdrH           As Long
    Dim uMetrics        As TEXTMETRICW
    Dim lX              As Long
    Dim oCol            As JSColumn
    Dim lW              As Long
    Dim hPrevFont       As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim lMarkX          As Long
    Dim bMark           As Boolean

    lHdrH = m_lColumnHeaderHeight
    lMarkX = -1
    pvFillRect hDC, 0, lY, picGrid.ScaleWidth, lY + lHdrH, picGrid.BackColor
    hPrevFont = pvSelectFont(hDC, m_oColumnHeaderFont)
    Select Case m_eHeaderStyle
    Case jgexHSDouble3D
        pvLine hDC, 0, lY, picGrid.ScaleWidth, lY, vb3DHighlight, PS_SOLID
        pvLine hDC, 0, lY + lHdrH - 2, picGrid.ScaleWidth, lY + lHdrH - 2, vb3DShadow, PS_SOLID
        pvLine hDC, 0, lY + lHdrH - 1, picGrid.ScaleWidth, lY + lHdrH - 1, vb3DDKShadow, PS_SOLID
    Case jgexHSSingleFlat
        pvLine hDC, 0, lY, picGrid.ScaleWidth, lY, vb3DDKShadow, PS_SOLID
        pvLine hDC, 0, lY + lHdrH - 1, picGrid.ScaleWidth, lY + lHdrH - 1, vb3DDKShadow, PS_SOLID
    Case jgexHSSingle3D
        pvLine hDC, 0, lY, picGrid.ScaleWidth, lY, vb3DHighlight, PS_SOLID
        pvLine hDC, 0, lY + lHdrH - 1, picGrid.ScaleWidth, lY + lHdrH - 1, vb3DShadow, PS_SOLID
    End Select
    If m_bRowHeaders Then
        pvPaintHeaderCell hDC, 0, lY, CHROME_COL_W, lHdrH, vbNullString, jgexAlignLeft
    End If
    lX = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        lW = pvColWidth(oCol)
        pvPaintHeaderCell hDC, lX, lY, lW, lHdrH, oCol.Caption, oCol.HeaderAlignment
        If frColSortOrder(oCol.Index) <> 0 Then
            uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
            pvPaintSortGlyph hDC, lX + 2 + pvTextWidth(hDC, oCol.Caption) + 4, _
                lY + (lHdrH - uMetrics.tmHeight + 1) \ 2 + uMetrics.tmHeight \ 2 + 4, frColSortOrder(oCol.Index)
        End If
        If oCol Is m_oDragCol Then
            Call PatBlt(hDC, lX, lY + 1, lW - 1, lHdrH - 2, DSTINVERT)
        End If
        If m_bDragging And oCol Is m_oDropCol And Not oCol Is pvGetDragColumn() Then
            If m_oDragCol Is Nothing Then
                bMark = True
            Else
                bMark = pvDropMoves(m_oDragCol.ColPosition, oCol.ColPosition, m_bDropAfter)
            End If
            If bMark Then
                lMarkX = lX
                If m_bDropAfter Then
                    lMarkX = lX + lW
                End If
            End If
        End If
        lX = lX + lW
    Next
    If pvGroupIndent() > 0 Then
        pvFillRect hDC, IIf(m_bRowHeaders, CHROME_COL_W, 1), lY + 1, pvBlockLeft() + 1, lY + lHdrH - 2, m_clrBackColorHeader
    End If
    If lX < picGrid.ScaleWidth Then
        pvPaintHeaderCell hDC, lX, lY, picGrid.ScaleWidth - lX + 2, lHdrH, vbNullString, jgexAlignLeft
    End If
    If lMarkX >= 0 Then
        pvFillRect hDC, lMarkX - 2, lY, lMarkX + 1, lY + lHdrH, vbRed
    End If
    Call SelectObject(hDC, hPrevFont)
End Sub

Private Sub pvPaintSortGlyph(ByVal hDC As Long, ByVal lX As Long, ByVal lBottom As Long, ByVal eOrder As jgexSortOrderConstants)
    Const GLYPH_W       As Long = 8
    Const GLYPH_H       As Long = 7
    Dim lRow            As Long
    Dim lEdge           As Long
    Dim lStep           As Long
    Dim lTop            As Long

    lTop = lBottom - GLYPH_H + 1
    For lRow = 0 To GLYPH_H - 1
        lEdge = lRow
        If eOrder = jgexSortDescending Then
            lEdge = GLYPH_H - 1 - lRow
        End If
        lStep = (lEdge + 1) \ 2
        pvSetPixel hDC, lX + GLYPH_W \ 2 - 1 - lStep, lTop + lRow, vb3DShadow
        pvSetPixel hDC, lX + GLYPH_W \ 2 + lStep, lTop + lRow, vb3DHighlight
        If lEdge Mod 2 = 1 Then
            pvSetPixel hDC, lX + GLYPH_W \ 2 - lStep, lTop + lRow, vb3DShadow
            pvSetPixel hDC, lX + GLYPH_W \ 2 - 1 + lStep, lTop + lRow, vb3DHighlight
        End If
    Next
    If eOrder = jgexSortDescending Then
        pvLine hDC, lX, lTop, lX + GLYPH_W - 1, lTop, vb3DShadow, PS_SOLID
    Else
        pvLine hDC, lX + 1, lBottom, lX + GLYPH_W, lBottom, vb3DHighlight, PS_SOLID
    End If
End Sub

Private Sub pvSetPixel(ByVal hDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal clrColor As OLE_COLOR)
    pvFillRect hDC, lX, lY, lX + 1, lY + 1, clrColor
End Sub

Private Sub pvPaintHeaderCell(ByVal hDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal lW As Long, ByVal lH As Long, sCaption As String, ByVal eAlign As jgexAlignmentConstants)
    Select Case m_eHeaderStyle
    Case jgexHSNoBorder
        pvFillRect hDC, lX, lY, lX + lW, lY + lH, m_clrBackColorHeader
    Case jgexHSSingleFlat
        pvFillRect hDC, lX, lY + 1, lX + lW - 1, lY + lH - 1, m_clrBackColorHeader
        pvLine hDC, lX + lW - 1, lY, lX + lW - 1, lY + lH, vb3DDKShadow, PS_SOLID
    Case jgexHSSingle3D
        pvFillRect hDC, lX + 1, lY + 1, lX + lW - 1, lY + lH - 1, m_clrBackColorHeader
        pvLine hDC, lX, lY, lX, lY + lH - 1, vb3DHighlight, PS_SOLID
        pvLine hDC, lX + lW - 1, lY, lX + lW - 1, lY + lH, vb3DShadow, PS_SOLID
    Case Else
        pvFillRect hDC, lX + 1, lY + 1, lX + lW - 2, lY + lH - 2, m_clrBackColorHeader
        pvLine hDC, lX, lY, lX, lY + lH - 2, vb3DHighlight, PS_SOLID
        pvLine hDC, lX + lW - 2, lY, lX + lW - 2, lY + lH - 2, vb3DShadow, PS_SOLID
        pvLine hDC, lX + lW - 1, lY, lX + lW - 1, lY + lH, vb3DDKShadow, PS_SOLID
    End Select
    If LenB(sCaption) <> 0 Then
        pvDrawText hDC, sCaption, lX + 2, lY, lX + lW - 2, lY + lH, m_clrForeColorHeader, m_clrBackColorHeader, eAlign, lX + 2, lX + lW - 2, bEllipsis:=True
    End If
End Sub

Private Sub pvPaintRows(ByVal hDC As Long, ByVal hMemDC As Long, ByVal lY As Long, uClip As RECT)
    Dim lRowH           As Long
    Dim lHdrW           As Long
    Dim lTotalW         As Long
    Dim lRow            As Long
    Dim lFirst          As Long
    Dim lRowTop         As Long
    Dim lRowsBottom     As Long
    Dim lCum            As Long
    Dim hPrevFont       As Long
    Dim lBandTop        As Long
    Dim lLineR          As Long
    Dim lPainted        As Long
    Dim lExtra          As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long

    lHdrW = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        lTotalW = lTotalW + pvColWidth(m_oColumns.ItemByPosition(vOrder(lIdx)))
    Next
    pvSyncProjection
    pvPopulateWindow
    lFirst = m_lFirstItem
    If lFirst < 1 Then
        lFirst = 1
    End If
    If m_eGridLines = jgexGLVertical Then
        lExtra = 1
    End If
    lRowTop = lY
    lRowsBottom = lY
    lRow = lFirst
    Do While lRowTop < picGrid.ScaleHeight
        lRowH = m_lRowHeight
        If lRowH <= 0 Then
            Exit Do
        End If
        If lRow > RowCount And Not m_bEmptyRows Then
            Exit Do
        End If
        lBandTop = lRowTop - 1
        If lBandTop < 0 Then
            lBandTop = 0
        End If
        If lRow <= RowCount Then
            lPainted = lPainted + 1
            lRowsBottom = lRowTop + lRowH
        End If
        If pvNeedRepaint(uClip, lBandTop, lRowTop + lRowH - lBandTop) Then
            pvBufferBand hMemDC, lBandTop, picGrid.ScaleWidth, lRowTop + lRowH - lBandTop
            hPrevFont = pvSelectFont(hMemDC, m_oFont)
            pvFillRect hMemDC, lHdrW + lTotalW, lRowTop, picGrid.ScaleWidth, lRowTop + lRowH, m_clrBackColorBkg
            If lRow > RowCount Then
                If m_bRowHeaders Then
                    pvFillRect hMemDC, 0, lRowTop, CHROME_COL_W, lRowTop + lRowH, m_clrBackColor
                    pvPaintRowHeader hMemDC, lRowTop, lRowH, CHROME_COL_W, False, EmptyRow:=True
                End If
                pvFillRect hMemDC, lHdrW, lRowTop, lHdrW + lTotalW, lRowTop + lRowH, m_clrBackColor
                If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLHorizontal Then
                    If lRowTop + lRowH - 1 < picGrid.ScaleHeight Then
                        lLineR = lHdrW + lTotalW
                        If m_eGridLines = jgexGLHorizontal Then
                            lLineR = lLineR + 1
                        End If
                        pvLine hMemDC, lHdrW, lRowTop + lRowH - 1, lLineR, lRowTop + lRowH - 1, m_clrGridLinesColor, pvPenStyle()
                    End If
                End If
            ElseIf pvIsGroupRow(lRow) Then
                pvPaintGroupRow hMemDC, lRow, lRowTop, lRowH, lHdrW + lTotalW, lY
            Else
                pvPaintDataRow hMemDC, lRow, lRowTop, lRowH, lHdrW, lTotalW
            End If
            If lRow = m_lRow And lRow <= RowCount And m_bGridFocus Then
                pvPaintRowMarquee hMemDC, lRowTop, lRowH, lHdrW, lTotalW
            End If
            pvPaintRowRules hMemDC, lRow, lRowTop, lRowH, lHdrW, vOrder
            pvBufferFlush hDC, hMemDC, lBandTop, picGrid.ScaleWidth, lRowTop + lRowH - lBandTop
            Call SelectObject(hMemDC, hPrevFont)
        End If
        lRowTop = lRowTop + lRowH
        lRow = lRow + 1
    Loop
    If lExtra > 0 And lPainted > 0 And m_bRowHeaders Then
        pvLine hDC, 0, lRowsBottom, CHROME_COL_W, lRowsBottom, vb3DDKShadow, PS_SOLID
    End If
    If m_bEmptyRows Then
        lRowsBottom = picGrid.ScaleHeight
    Else
        pvFillRect hDC, 0, lRowsBottom + lExtra, lHdrW, picGrid.ScaleHeight, m_clrBackColorBkg
        pvFillRect hDC, lHdrW, lRowsBottom, picGrid.ScaleWidth, picGrid.ScaleHeight, m_clrBackColorBkg
        If m_oGroups.Count > 0 And lPainted > 0 Then
            pvLine hDC, 0, lRowsBottom - 1, lHdrW, lRowsBottom - 1, m_clrGridLinesColor, pvPenStyle()
        End If
        If lExtra > 0 And lPainted > 0 And m_oGroups.Count = 0 Then
            For lIdx = 0 To pvOrderMax(vOrder)
                lCum = lCum + pvColWidth(m_oColumns.ItemByPosition(vOrder(lIdx)))
                pvLine hDC, lHdrW + lCum - 1, lRowsBottom, lHdrW + lCum - 1, lRowsBottom + lExtra, m_clrGridLinesColor, pvPenStyle(), DashAnchor:=lRowsBottom - m_lRowHeight
            Next
        End If
    End If
    If m_bBandVisible Then
        pvLine hDC, 0, picGrid.ScaleHeight - 1, picGrid.ScaleWidth, picGrid.ScaleHeight - 1, UserControl.BackColor, PS_SOLID
    End If
End Sub

Private Sub pvPaintRowRules(ByVal hDC As Long, ByVal lRow As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, vOrder As Variant)
    Dim lIdx            As Long
    Dim lCum            As Long
    Dim lX              As Long

    If m_eGridLines <> jgexGLBoth And m_eGridLines <> jgexGLVertical Then
        Exit Sub
    End If
    If lRow <= RowCount Then
        If pvIsGroupRow(lRow) Then
            Exit Sub
        End If
    End If
    For lIdx = 0 To pvOrderMax(vOrder)
        lCum = lCum + pvColWidth(m_oColumns.ItemByPosition(vOrder(lIdx)))
        lX = lHdrW + lCum - 1
        pvLine hDC, lX, lRowTop, lX, lRowTop + lRowH, m_clrGridLinesColor, pvPenStyle()
    Next
End Sub

Private Sub pvPaintDataRow(ByVal hDC As Long, ByVal lRow As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal lTotalW As Long)
    Dim bSelected       As Boolean
    Dim bCurrentRow     As Boolean
    Dim bSelOutline     As Boolean
    Dim bSelInactive    As Boolean
    Dim lOutR           As Long
    Dim clrBack         As OLE_COLOR
    Dim clrText         As OLE_COLOR
    Dim lFillL          As Long
    Dim lFillR          As Long
    Dim clrCellBack     As OLE_COLOR
    Dim clrCellText     As OLE_COLOR
    Dim lPos            As Long
    Dim lX              As Long
    Dim oCol            As JSColumn
    Dim lW              As Long
    Dim sText           As String
    Dim lClipR          As Long
    Dim lVLine          As Long
    Dim lMarqueeR       As Long
    Dim bHLine          As Boolean
    Dim clrHLine        As OLE_COLOR
    Dim clrHGap         As OLE_COLOR
    Dim lPenH           As Long
    Dim lLineR          As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long

    bSelected = pvIsRowSelected(lRow) Or (m_lRow >= 1 And lRow = m_lRow And Not m_bCurRowDeselected)
    If bSelected And Not m_bGridFocus Then
        Select Case m_eHideSelection
        Case jgexHideSelection
            bSelected = False
            bSelOutline = True
        Case jgexHighLightInactive
            bSelInactive = True
        End Select
    End If
    bCurrentRow = bSelected And lRow = m_lRow And m_oSelectedItems.Count <= 1
    If bSelInactive Then
        clrBack = vbButtonFace
        clrText = m_clrForeColor
    ElseIf bSelected Then
        pvSelColors clrBack, clrText
    ElseIf m_bUseEvenOddColor Then
        If lRow Mod 2 = 0 Then
            clrBack = m_clrRowColorEven
        Else
            clrBack = m_clrRowColorOdd
        End If
        clrText = m_clrForeColor
    Else
        clrBack = m_clrBackColor
        clrText = m_clrForeColor
    End If
    If m_bRowHeaders Then
        pvFillRect hDC, 0, lRowTop, CHROME_COL_W, lRowTop + lRowH, m_clrBackColor
        If m_bEmptyRows Then
            pvPaintRowHeader hDC, lRowTop, lRowH, CHROME_COL_W, False, EmptyRow:=True
        End If
        pvPaintRowHeader hDC, lRowTop, lRowH, CHROME_COL_W, (lRow = m_lRow)
    End If
    If pvGroupIndent() > 0 Then
        pvFillRect hDC, pvRowHeaderWidth(), lRowTop, lHdrW, lRowTop + lRowH, m_clrBackColorRowGroup
        pvPaintIndentRules hDC, lRowTop, lRowH, pvGroupIndent()
    End If
    If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLVertical Then
        lVLine = 1
    End If
    lMarqueeR = -1
    If lRow = m_lRow Then
        lMarqueeR = lHdrW + lTotalW - 2
        If lMarqueeR > picGrid.ScaleWidth - 1 Then
            lMarqueeR = picGrid.ScaleWidth - 1
        End If
    End If
    bHLine = (m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLHorizontal)
    lLineR = lHdrW + lTotalW
    If m_eGridLines = jgexGLHorizontal Then
        lLineR = lLineR + 1
    End If
    clrHLine = m_clrGridLinesColor
    lPenH = pvPenStyle()
    If lRow = RowCount Then
        clrHLine = vb3DDKShadow
        lPenH = PS_SOLID
    End If
    clrHGap = m_clrBackColorBkg
    If lPenH = PS_DASH Then
        clrHGap = m_clrBackColor
    End If
    lX = lHdrW
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        lPos = pvVisiblePosition(vOrder(lIdx))
        lW = pvColWidth(oCol)
        pvFillRect hDC, lX, lRowTop, lX + lW - lVLine, lRowTop + pvRowContentH(lRowH), clrBack
        If bCurrentRow And m_lCol = lPos Then
            lFillL = lX
            lFillR = lX + lW
            If lFillL <= lHdrW Then
                lFillL = lHdrW + 1
            End If
            If lFillR > pvMarqueeRight(lHdrW, lTotalW) Then
                lFillR = pvMarqueeRight(lHdrW, lTotalW)
            End If
            If lFillR > lX + lW - lVLine Then
                lFillR = lX + lW - lVLine
            End If
            pvFillRect hDC, lFillL, lRowTop + 1, lFillR, lRowTop + pvRowContentH(lRowH) - 1, m_clrBackColor
            clrCellBack = m_clrBackColor
            clrCellText = m_clrForeColor
        Else
            clrCellBack = clrBack
            clrCellText = clrText
        End If
        If oCol.ColumnType = jgexCheckBox Then
            '--- a checkbox column draws its state instead of its text, and
            '--- the box on the current cell is outlined a shade darker
            pvPaintCheckBox hDC, lX + (lW - CHECK_BOX_W) \ 2 - 1, lRowTop + (pvRowContentH(lRowH) - CHECK_BOX_H) \ 2, _
                pvIsChecked(pvWindowRow(lRow), oCol.Index), (bCurrentRow And m_lCol = lPos)
            sText = vbNullString
        Else
            sText = pvCellText(pvWindowRow(lRow), oCol.Index)
        End If
        If LenB(sText) <> 0 Then
            lClipR = lX + lW
            If lMarqueeR >= 0 And lClipR > lMarqueeR Then
                lClipR = lMarqueeR
            End If
            pvDrawText hDC, sText, lX + 2, lRowTop, lX + lW - 3, lRowTop + pvRowContentH(lRowH), clrCellText, clrCellBack, oCol.TextAlignment, lX, lClipR, oCol.WordWrap, bEllipsis:=True
        End If
        If bHLine Then
            pvLine hDC, lX, lRowTop + lRowH - 1, lX + lW, lRowTop + lRowH - 1, clrHLine, lPenH, DashAnchor:=lLineR - 1, GapColor:=clrHGap
        End If
        lX = lX + lW
    Next
    '--- the pixel past the block that no cell owns
    If bHLine And lLineR > lX Then
        pvLine hDC, lX, lRowTop + lRowH - 1, lLineR, lRowTop + lRowH - 1, clrHLine, lPenH, DashAnchor:=lLineR - 1, GapColor:=clrHGap
    End If
    If bSelOutline Then
        lOutR = lHdrW + lTotalW - 2
        If lOutR > picGrid.ScaleWidth - 1 Then
            lOutR = picGrid.ScaleWidth - 1
        End If
        pvFillRect hDC, lHdrW, lRowTop, lOutR + 1, lRowTop + 1, vbHighlight
        pvFillRect hDC, lHdrW, lRowTop + pvRowContentH(lRowH) - 1, lOutR + 1, lRowTop + pvRowContentH(lRowH), vbHighlight
        pvFillRect hDC, lHdrW, lRowTop, lHdrW + 1, lRowTop + pvRowContentH(lRowH), vbHighlight
        pvFillRect hDC, lOutR, lRowTop, lOutR + 1, lRowTop + pvRowContentH(lRowH), vbHighlight
    End If
End Sub

Private Sub pvPaintRowMarquee(ByVal hDC As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal lTotalW As Long)
    Dim hBrush          As Long
    Dim hPrevBrush      As Long
    Dim lRight          As Long
    Dim lBottom         As Long
    Dim lCum            As Long
    Dim oCol            As JSColumn
    Dim lW              As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lStartB         As Long
    Dim lDxRight        As Long
    Dim lPass           As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim lLeft           As Long
    Dim bGroupRow       As Boolean

    bGroupRow = pvIsGroupRow(m_lRow)
    lLeft = lHdrW
    If bGroupRow Then
        lLeft = 0
    End If
    lRight = lHdrW + lTotalW - 1
    If Not bGroupRow Then
        If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLVertical Then
            lRight = lRight - 1
        End If
    End If
    If lRight > picGrid.ScaleWidth - 1 Then
        lRight = picGrid.ScaleWidth - 1
    End If
    lBottom = lRowTop + pvRowContentH(lRowH) - 1
    lStartB = 1
    If (lBottom - lRowTop) Mod 2 = 1 Then
        lStartB = 0
    End If
    For lPass = 0 To 1
        lDxRight = -1
        If lPass = 0 Then
            hBrush = CreateSolidBrush(vbWhite)
        Else
            hBrush = CreateSolidBrush(pvColor(m_clrForeColor))
        End If
        hPrevBrush = SelectObject(hDC, hBrush)
        lCum = lLeft
        vOrder = pvColOrder()
        For lIdx = 0 To pvOrderMax(vOrder)
            lW = 0
            If bGroupRow Then
                '--- the whole block in one go, on the first pass round only
                If lIdx = 0 Then
                    lW = lRight - lLeft + 1
                End If
            Else
                Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
                If oCol.Visible Then
                    lW = pvColWidth(oCol)
                End If
            End If
            If lW > 0 Then
                For lX = lCum + 1 - lPass To lCum + lW - 1 Step 2
                    If lX <= lRight Then
                        Call PatBlt(hDC, lX, lRowTop, 1, 1, PATINVERT)
                    End If
                Next
                For lX = lCum + lStartB + lPass To lCum + lW - 1 Step 2
                    If lX <= lRight Then
                        Call PatBlt(hDC, lX, lBottom, 1, 1, PATINVERT)
                    End If
                Next
                If lRight >= lCum And lRight < lCum + lW Then
                    lDxRight = lRight - lCum
                End If
                lCum = lCum + lW
            End If
        Next
        '--- outer vertical edges, phased by their own column-relative x
        For lY = lRowTop + 1 + lPass To lBottom - 1 Step 2
            Call PatBlt(hDC, lLeft, lY, 1, 1, PATINVERT)
        Next
        If lDxRight >= 0 Then
            lY = lRowTop + 1 + lPass
            If lDxRight Mod 2 = 1 Then
                lY = lRowTop + 2 + lPass
            End If
            Do While lY <= lBottom - 1
                Call PatBlt(hDC, lRight, lY, 1, 1, PATINVERT)
                lY = lY + 2
            Loop
        End If
        Call SelectObject(hDC, hPrevBrush)
        Call DeleteObject(hBrush)
    Next
End Sub

Private Sub pvPaintGroupRow(ByVal hDC As Long, ByVal lPos As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lRight As Long, ByVal lBlockTop As Long)
    Dim oRowData        As JSRowData
    Dim bFooter         As Boolean
    Dim clrBack         As OLE_COLOR
    Dim clrText         As OLE_COLOR
    Dim lIndent         As Long
    Dim uBox            As RECT
    Dim lBoxTop         As Long
    Dim lBoxLeft        As Long
    Dim lTextTop        As Long
    Dim lTextLeft       As Long
    Dim lLineLeft       As Long
    Dim uMetrics        As TEXTMETRICW

    Set oRowData = pvWindowRow(lPos)
    If oRowData Is Nothing Then
        Exit Sub
    End If
    bFooter = (oRowData.RowType = jgexRowTypeGroupFooter)
    lIndent = pvRowIndent(lPos)
    If pvIsRowSelected(lPos) Or (m_lRow >= 1 And lPos = m_lRow) Then
        pvSelColors clrBack, clrText
    Else
        clrBack = m_clrBackColorRowGroup
        clrText = m_clrForeColorRowGroup
    End If
    pvFillRect hDC, 0, lRowTop, lRight, lRowTop + lRowH, clrBack
    pvFillRect hDC, 0, lRowTop + lRowH - 1, pvBlockLeft() - 1, lRowTop + lRowH, m_clrBackColorRowGroup
    pvPaintIndentRules hDC, lRowTop, lRowH, lIndent
    pvLine hDC, pvBlockLeft() - 1, lRowTop + lRowH - 1, lRight, lRowTop + lRowH - 1, m_clrGridLinesColor, pvPenStyle()
    If lRowTop > lBlockTop Then
        lLineLeft = pvRowHeaderWidth() + lIndent - 1
        If bFooter Then
            lLineLeft = pvBlockLeft() - 1
        End If
        pvLine hDC, lLineLeft, lRowTop - 1, lRight, lRowTop - 1, m_clrGridLinesColor, pvPenStyle()
    End If
    If pvGroupBoxRect(lPos, lRowTop, uBox) Then
        lBoxLeft = uBox.Left
        lBoxTop = uBox.Top
    End If
    If bFooter Then
        If m_eGroupFooterStyle = jgexTotalsGroupFooter Then
            pvPaintGroupTotals hDC, oRowData, lRowTop, lRowH, clrBack, clrText
            Exit Sub
        End If
    Else
        pvPaintGroupBox hDC, lBoxLeft, lBoxTop, m_pDataModel.RowExpanded(lPos)
    End If
    uMetrics = FontTextMetrics(m_oFont, hDC)
    lTextTop = lRowTop + (lRowH - 1 - uMetrics.tmHeight) \ 2
    lTextLeft = lBoxLeft + GROUP_BOX_W + 2
    If Not oRowData.frGroupPrefixed Then
        lTextLeft = lTextLeft + pvTextWidth(hDC, " ")
    End If
    pvDrawText hDC, oRowData.GroupCaption, lTextLeft, lTextTop, lRight, lTextTop + uMetrics.tmHeight, _
        clrText, clrBack, jgexAlignLeft, lTextLeft, lRight
End Sub

Private Sub pvPaintGroupTotals(ByVal hDC As Long, oRowData As JSRowData, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal clrBack As OLE_COLOR, ByVal clrText As OLE_COLOR)
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lW              As Long
    Dim oCol            As JSColumn
    Dim vValue          As Variant
    Dim sText           As String
    Dim lTextTop        As Long
    Dim vOrder          As Variant
    Dim uMetrics        As TEXTMETRICW

    '--- each column reads its own aggregate over the group, laid out on the
    '--- cell origins a record uses so the totals line up under their values
    uMetrics = FontTextMetrics(m_oFont, hDC)
    lTextTop = lRowTop + (lRowH - 1 - uMetrics.tmHeight) \ 2
    lX = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oCol.Visible Then
            lW = pvColWidth(oCol)
            If oCol.AggregateFunction <> jgexAggregateNone Then
                vValue = oRowData.Value(oCol.Index)
                sText = vbNullString
                If Not IsEmpty(vValue) Then
                    If LenB(oCol.TotalRowFormat) <> 0 Then
                        sText = Format$(vValue, oCol.TotalRowFormat)
                    Else
                        sText = CStr(vValue)
                    End If
                End If
                sText = oCol.TotalRowPrefix & sText
                pvDrawText hDC, sText, lX + 2, lTextTop, lX + lW - 3, lTextTop + uMetrics.tmHeight, _
                    clrText, clrBack, oCol.TextAlignment, lX, lX + lW - 3
            End If
            lX = lX + lW
        End If
    Next
End Sub

Private Sub pvPaintGroupBox(ByVal hDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal bExpanded As Boolean)
    Dim lMid            As Long

    pvFillRect hDC, lX + 1, lY + 1, lX + GROUP_BOX_W - 2, lY + GROUP_BOX_W - 2, vbButtonFace
    pvLine hDC, lX, lY, lX + GROUP_BOX_W - 1, lY, vb3DHighlight, PS_SOLID
    pvLine hDC, lX, lY, lX, lY + GROUP_BOX_W - 1, vb3DHighlight, PS_SOLID
    pvLine hDC, lX + GROUP_BOX_W - 2, lY, lX + GROUP_BOX_W - 2, lY + GROUP_BOX_W - 1, vb3DShadow, PS_SOLID
    pvLine hDC, lX, lY + GROUP_BOX_W - 2, lX + GROUP_BOX_W - 1, lY + GROUP_BOX_W - 2, vb3DShadow, PS_SOLID
    pvLine hDC, lX + GROUP_BOX_W - 1, lY, lX + GROUP_BOX_W - 1, lY + GROUP_BOX_W, vb3DDKShadow, PS_SOLID
    pvLine hDC, lX, lY + GROUP_BOX_W - 1, lX + GROUP_BOX_W, lY + GROUP_BOX_W - 1, vb3DDKShadow, PS_SOLID
    lMid = lY + (GROUP_BOX_W - 2) \ 2
    pvFillRect hDC, lX + 3, lMid, lX + GROUP_BOX_W - 4, lMid + 1, m_clrForeColorRowGroup
    If Not bExpanded Then
        pvFillRect hDC, lX + (GROUP_BOX_W - 2) \ 2, lY + 3, lX + (GROUP_BOX_W - 2) \ 2 + 1, lY + GROUP_BOX_W - 4, m_clrForeColorRowGroup
    End If
End Sub

Private Sub pvPaintRowHeader(ByVal hDC As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal bCurrent As Boolean, Optional ByVal EmptyRow As Boolean)
    Dim lIdx            As Long
    Dim lC              As Long

    If EmptyRow Then
        lC = lRowH - 1
    Else
        Select Case m_eGridLines
        Case jgexGLVertical
            lC = lRowH
        Case jgexGLHorizontal
            lC = lRowH - 2
        Case Else
            lC = lRowH - 1
        End Select
    End If
    pvFillRect hDC, 1, lRowTop + 1, lHdrW - 2, lRowTop + lC - 1, m_clrBackColorHeader
    pvLine hDC, 0, lRowTop, lHdrW, lRowTop, vb3DHighlight, PS_SOLID
    pvLine hDC, 0, lRowTop, 0, lRowTop + lC - 1, vb3DHighlight, PS_SOLID
    pvLine hDC, 0, lRowTop + lC - 1, lHdrW, lRowTop + lC - 1, vb3DShadow, PS_SOLID
    If lC < lRowH Then
        pvLine hDC, 0, lRowTop + lC, lHdrW, lRowTop + lC, vb3DDKShadow, PS_SOLID
    End If
    pvLine hDC, lHdrW - 2, lRowTop, lHdrW - 2, lRowTop + lC - 1, vb3DShadow, PS_SOLID
    pvLine hDC, lHdrW - 1, lRowTop, lHdrW - 1, lRowTop + lC, vb3DDKShadow, PS_SOLID
    '--- current row arrow marker
    If bCurrent Then
        For lIdx = 0 To 5
            pvLine hDC, 6 + lIdx, lRowTop + 3 + lIdx, 6 + lIdx, lRowTop + 15 - lIdx, vbButtonText, PS_SOLID
        Next
    End If
End Sub

Private Sub pvPaintCheckBox(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal bChecked As Boolean, ByVal bCurrent As Boolean)
    Dim lIdx            As Long
    Dim vRuns           As Variant
    Dim clrBorder       As OLE_COLOR
    Dim lOffX           As Long
    Dim lOffY           As Long

    clrBorder = CHECK_BOX_CLR
    If bCurrent Then
        clrBorder = CHECK_BOX_CLR_CUR
    End If
    pvFillRect hDC, lLeft, lTop, lLeft + CHECK_BOX_W, lTop + CHECK_BOX_H, vbWhite
    pvFillRect hDC, lLeft, lTop, lLeft + CHECK_BOX_W, lTop + 1, clrBorder
    pvFillRect hDC, lLeft, lTop + CHECK_BOX_H - 1, lLeft + CHECK_BOX_W, lTop + CHECK_BOX_H, clrBorder
    pvFillRect hDC, lLeft, lTop, lLeft + 1, lTop + CHECK_BOX_H, clrBorder
    pvFillRect hDC, lLeft + CHECK_BOX_W - 1, lTop, lLeft + CHECK_BOX_W, lTop + CHECK_BOX_H, clrBorder
    If Not bChecked Then
        Exit Sub
    End If
    Select Case pvScreenDpi()
    Case Is >= 144
        lOffX = 3
        lOffY = 8
        vRuns = Array(0, 0, 0, 0, 6, 6, 1, 0, 1, 1, 5, 6, 2, 0, 2, 2, 4, 6, 3, 0, 6)
    Case Is >= 120
        lOffX = 3
        lOffY = 6
        vRuns = Array(0, 6, 6, 1, 5, 6, 2, 0, 0, 2, 4, 6, 3, 0, 1, 3, 3, 5, 4, 0, 4, 5, 1, 3, 6, 2, 2)
    Case Else
        lOffX = 2
        lOffY = 4
        vRuns = Array(0, 6, 6, 1, 5, 6, 2, 0, 0, 2, 4, 6, 3, 0, 1, 3, 3, 5, 4, 0, 4, 5, 1, 3, 6, 2, 2)
    End Select
    For lIdx = 0 To UBound(vRuns) - 2 Step 3
        '--- a run landing on the bottom line draws over it, one past the box
        '--- is dropped: at 120dpi the tick sits low enough to do both
        If lOffY + vRuns(lIdx) < CHECK_BOX_H Then
            pvFillRect hDC, lLeft + lOffX + vRuns(lIdx + 1), lTop + lOffY + vRuns(lIdx), _
                lLeft + lOffX + vRuns(lIdx + 2) + 1, lTop + lOffY + vRuns(lIdx) + 1, vbBlack
        End If
    Next
End Sub

Private Function pvNeedRepaint(uClip As RECT, ByVal lTop As Long, ByVal lHeight As Long) As Boolean
    pvNeedRepaint = (uClip.Bottom > lTop And uClip.Top < lTop + lHeight)
End Function

Private Function pvBufferInit(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
    pvBufferInit = hDC
    pvBufferTermiante
    m_lBufLastY = -1
    m_lBufLastRow = 0
    If lWidth <= 0 Or lHeight <= 0 Then
        Exit Function
    End If
#If FORCE_BUFFER = 0 Then
    If GetSystemMetrics(SM_REMOTESESSION) <> 0 Then
        Exit Function
    End If
#End If
    m_hBufDC = CreateCompatibleDC(hDC)
    If m_hBufDC = 0 Then
        Exit Function
    End If
    m_hBufBmp = CreateCompatibleBitmap(hDC, lWidth, lHeight)
    If m_hBufBmp = 0 Then
        pvBufferTermiante
        Exit Function
    End If
    m_hBufOldBmp = SelectObject(m_hBufDC, m_hBufBmp)
    pvBufferInit = m_hBufDC
End Function

Private Sub pvBufferBand(ByVal hMemDC As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    If m_hBufDC = 0 Or hMemDC <> m_hBufDC Then
        Exit Sub
    End If
    If m_lBufLastY = lTop Then
        Call SetViewportOrgEx(m_hBufDC, 0, 0, 0)
        Call BitBlt(m_hBufDC, 0, 0, lWidth, 1, m_hBufDC, 0, m_lBufLastRow, SRCCOPY)
    End If
    Call SetViewportOrgEx(m_hBufDC, 0, -lTop, 0)
    Call SelectClipRgn(m_hBufDC, 0)
    Call IntersectClipRect(m_hBufDC, 0, lTop, lWidth, lTop + lHeight)
End Sub

Private Sub pvBufferFlush(ByVal hDC As Long, ByVal hMemDC As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    If hMemDC = hDC Then
        Exit Sub
    End If
    Call BitBlt(hDC, 0, lTop, lWidth, lHeight, hMemDC, 0, lTop, SRCCOPY)
    m_lBufLastRow = lHeight - 1
    m_lBufLastY = lTop + lHeight - 1
End Sub

Private Sub pvBufferTermiante()
    If m_hBufOldBmp <> 0 Then
        Call SelectObject(m_hBufDC, m_hBufOldBmp)
        m_hBufOldBmp = 0
    End If
    If m_hBufBmp <> 0 Then
        Call DeleteObject(m_hBufBmp)
        m_hBufBmp = 0
    End If
    If m_hBufDC <> 0 Then
        Call DeleteDC(m_hBufDC)
        m_hBufDC = 0
    End If
End Sub

Private Function pvRowIndent(ByVal lPos As Long) As Long
    If pvIsGroupRow(lPos) Then
        pvRowIndent = (pvWindowRow(lPos).GroupLevel - 1) * GROUP_INDENT_W
    Else
        pvRowIndent = pvGroupIndent()
    End If
End Function

Private Sub pvPaintIndentRules(ByVal hDC As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lIndent As Long)
    Dim lIdx            As Long

    '--- one vertical rule per level the row is nested in
    For lIdx = GROUP_INDENT_W To lIndent Step GROUP_INDENT_W
        pvLine hDC, pvRowHeaderWidth() + lIdx - 1, lRowTop, pvRowHeaderWidth() + lIdx - 1, lRowTop + lRowH, m_clrGridLinesColor, pvPenStyle()
    Next
End Sub

Private Function pvRowHeaderWidth() As Long
    If m_bRowHeaders Then
        pvRowHeaderWidth = CHROME_COL_W
    End If
End Function

Private Function pvBlockLeft() As Long
    If m_bRowHeaders Then
        pvBlockLeft = CHROME_COL_W
    End If
    pvBlockLeft = pvBlockLeft + pvGroupIndent()
End Function

Private Function pvGroupIndent() As Long
    pvGroupIndent = m_oGroups.Count * GROUP_INDENT_W
End Function

Private Function pvIsChecked(oRowData As JSRowData, ByVal lColIndex As Long) As Boolean
    Dim vValue          As Variant

    If oRowData Is Nothing Then
        Exit Function
    End If
    vValue = oRowData.Value(lColIndex)
    If Not IsObject(vValue) And Not IsArray(vValue) Then
        If Not pvIsBlank(vValue) Then
            pvIsChecked = CBool(vValue)
        End If
    End If
End Function

Private Function pvMarqueeRight(ByVal lHdrW As Long, ByVal lTotalW As Long) As Long
    pvMarqueeRight = lHdrW + lTotalW - 1
    If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLVertical Then
        pvMarqueeRight = pvMarqueeRight - 1
    End If
End Function

Private Function pvScreenDpi() As Long
    Dim hScreenDC       As Long

    hScreenDC = GetDC(0)
    pvScreenDpi = GetDeviceCaps(hScreenDC, LOGPIXELSY)
    Call ReleaseDC(0, hScreenDC)
End Function

Private Function pvCellText(oRowData As JSRowData, ByVal lColIndex As Long) As String
    Dim vValue          As Variant

    If oRowData Is Nothing Then
        Exit Function
    End If
    If oRowData.frCellDirty(lColIndex) Then
        AssignVariant vValue, oRowData.Value(lColIndex)
        If Not IsArray(vValue) Then
            Select Case VarType(vValue)
            Case vbEmpty, vbNull, vbObject, vbError
            Case Else
                pvCellText = CStr(vValue)
            End Select
        End If
        Exit Function
    End If
    pvCellText = oRowData.DisplayValue(lColIndex)
    If LenB(pvCellText) = 0 Then
        vValue = oRowData.Value(lColIndex)
        Select Case VarType(vValue)
        Case vbEmpty, vbNull, vbObject, vbError
        Case Else
            If Not IsArray(vValue) Then
                pvCellText = CStr(vValue)
            End If
        End Select
    End If
End Function

Private Function pvRowContentH(ByVal lRowH As Long) As Long
    pvRowContentH = lRowH
    If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLHorizontal Then
        pvRowContentH = lRowH - 1
    End If
End Function

Private Function pvPenStyle() As Long
    Select Case m_eGridLineStyle
    Case jgexGLSDashes
        pvPenStyle = PS_DASH
    Case jgexGLSSmallDots
        pvPenStyle = PS_DOT
    Case Else
        pvPenStyle = PS_SOLID
    End Select
End Function

Private Sub pvInvalidate(Optional ByVal SkipScroll As Boolean)
    Const FUNC_NAME     As String = "pvInvalidate"

    On Error GoTo EH
    If Not m_bRedraw Then
        Exit Sub
    End If
    If Not SkipScroll Then
        pvUpdateScrollBars
    End If
    pvLayoutEditor
    Call InvalidateRect(picGrid.hWnd, 0, 0)
    If m_bRecordNavigator Then
        Call InvalidateRect(UserControl.hWnd, 0, 0)
        If hsbGrid.Visible Then
            Call InvalidateRect(hsbGrid.hWnd, 0, 1)
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvInvalidateHeaders()
    Dim uRect           As RECT

    If Not m_bRedraw Then
        Exit Sub
    End If
    uRect.Right = picGrid.ScaleWidth
    uRect.Bottom = pvRowsTop()
    Call InvalidateRect(picGrid.hWnd, VarPtr(uRect), 0)
End Sub

Private Function pvTopHeight() As Long
    pvTopHeight = pvGroupByBoxHeight()
    If m_bColumnHeaders Then
        pvTopHeight = pvTopHeight + m_lColumnHeaderHeight
    End If
End Function

Private Function pvVisibleRows() As Long
    If m_lRowHeight > 0 Then
        pvVisibleRows = (picGrid.ScaleHeight - pvTopHeight()) \ m_lRowHeight
    End If
    If pvVisibleRows < 1 Then
        pvVisibleRows = 1
    End If
End Function

Private Function pvFirstCol() As Long
    pvFirstCol = m_lLeftCol
    If pvFirstCol < 1 Then
        pvFirstCol = 1
    End If
    If pvFirstCol > m_oColumns.Count Then
        pvFirstCol = m_oColumns.Count
    End If
End Function

Private Function pvScrollableWidth() As Long
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim lFrozen         As Long

    pvScrollableWidth = picGrid.ScaleWidth - pvBlockLeft()
    lFrozen = pvFrozenCount()
    For lIdx = 1 To lFrozen
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            pvScrollableWidth = pvScrollableWidth - pvColWidth(oCol)
        End If
    Next
End Function

Private Function pvVisibleColCount() As Long
    Dim oCol            As JSColumn

    For Each oCol In m_oColumns
        If oCol.Visible Then
            pvVisibleColCount = pvVisibleColCount + 1
        End If
    Next
End Function

Private Function pvVisiblePosition(ByVal lPosition As Long) As Long
    Dim lIdx            As Long

    For lIdx = 1 To lPosition
        If m_oColumns.ItemByPosition(lIdx).Visible Then
            pvVisiblePosition = pvVisiblePosition + 1
        End If
    Next
End Function

Private Function pvClampRow(ByVal lRow As Long) As Long
    pvClampRow = Clamp(lRow, 1, RowCount)
End Function

Private Function pvHitScrollBar() As Boolean
    Dim uPt             As POINTAPI

    If m_hWndHScroll = 0 Then
        Exit Function
    End If
    Call GetCursorPos(uPt)
    pvHitScrollBar = (WindowFromPoint(uPt.X, uPt.Y) = m_hWndHScroll)
End Function

Private Function pvNavLayout(uNav As UcsNavLayout) As Boolean
    Dim lBtnW           As Long
    Dim lBtnTop         As Long
    Dim lBtnH           As Long
    Dim lBoxW           As Long
    Dim hDC             As Long
    Dim hPrevFont       As Long
    Dim lX              As Long
    Dim vSplit          As Variant

    If Not m_bRecordNavigator Then
        Exit Function
    End If
    uNav.BandH = GetSystemMetrics(SM_CYHSCROLL)
    uNav.BandTop = UserControl.ScaleHeight - uNav.BandH
    If uNav.BandH <= 0 Then
        Exit Function
    End If
    vSplit = Split(m_sRecordNavigatorString & "|", "|")
    uNav.Prefix = vSplit(0)
    uNav.Middle = vSplit(1) & " " & RowCount
    lBtnW = uNav.BandH + 1
    lBtnTop = uNav.BandTop + 1
    lBtnH = uNav.BandH - 1
    hDC = GetDC(UserControl.hWnd)
    hPrevFont = pvSelectFont(hDC, m_oFont)
    lX = 4
    uNav.PrefixX = lX
    lX = lX + pvTextWidth(hDC, uNav.Prefix) + 4
    pvSetRect uNav.BtnFirst, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    pvSetRect uNav.BtnPrev, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW + 4
    lBoxW = pvTextWidth(hDC, "9999999") + 4
    pvSetRect uNav.Box, lX, uNav.BandTop, lX + lBoxW, uNav.BandTop + uNav.BandH + 4
    lX = lX + lBoxW + 4
    uNav.MiddleX = lX
    lX = lX + pvTextWidth(hDC, uNav.Middle) + 8
    pvSetRect uNav.BtnNext, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    pvSetRect uNav.BtnLast, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    uNav.Width = lX
    Call SelectObject(hDC, hPrevFont)
    Call ReleaseDC(UserControl.hWnd, hDC)
    pvNavLayout = True
End Function

Private Sub pvSetRect(uRect As RECT, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
    uRect.Left = lLeft
    uRect.Top = lTop
    uRect.Right = lRight
    uRect.Bottom = lBottom
End Sub

Private Function pvTextWidth(ByVal hDC As Long, sText As String) As Long
    Dim uSize           As SIZEAPI

    Call GetTextExtentPoint32(hDC, StrPtr(sText), Len(sText), uSize)
    pvTextWidth = uSize.cx
End Function

Private Sub pvOnNavigatorClick(ByVal lX As Long, ByVal lY As Long)
    Dim uNav            As UcsNavLayout

    '--- painted buttons, so hit-testing is done here rather than by a control
    If Not pvNavLayout(uNav) Then
        Exit Sub
    End If
    If pvPtInRect(uNav.BtnFirst, lX, lY) Then
        Row = 1
    ElseIf pvPtInRect(uNav.BtnPrev, lX, lY) Then
        If m_lRow > 1 Then
            Row = m_lRow - 1
        End If
    ElseIf pvPtInRect(uNav.BtnNext, lX, lY) Then
        If m_lRow < RowCount Then
            Row = m_lRow + 1
        End If
    ElseIf pvPtInRect(uNav.BtnLast, lX, lY) Then
        Row = RowCount
    Else
        Exit Sub
    End If
    EnsureVisible m_lRow
End Sub

Private Function pvPtInRect(uRect As RECT, ByVal lX As Long, ByVal lY As Long) As Boolean
    pvPtInRect = (lX >= uRect.Left And lX < uRect.Right And lY >= uRect.Top And lY < uRect.Bottom)
End Function

Private Sub pvNavButton(ByVal hDC As Long, uBtn As RECT)
    pvFillRect hDC, uBtn.Left, uBtn.Top, uBtn.Right, uBtn.Bottom, UserControl.BackColor
    pvLine hDC, uBtn.Left, uBtn.Top, uBtn.Right, uBtn.Top, vb3DHighlight, PS_SOLID
    pvLine hDC, uBtn.Left, uBtn.Top, uBtn.Left, uBtn.Bottom, vb3DHighlight, PS_SOLID
    pvLine hDC, uBtn.Right - 2, uBtn.Top + 1, uBtn.Right - 2, uBtn.Bottom - 1, vb3DShadow, PS_SOLID
    pvLine hDC, uBtn.Left + 1, uBtn.Bottom - 2, uBtn.Right - 1, uBtn.Bottom - 2, vb3DShadow, PS_SOLID
    pvLine hDC, uBtn.Right - 1, uBtn.Top, uBtn.Right - 1, uBtn.Bottom, vb3DDKShadow, PS_SOLID
    pvLine hDC, uBtn.Left, uBtn.Bottom - 1, uBtn.Right, uBtn.Bottom - 1, vb3DDKShadow, PS_SOLID
End Sub

Private Sub pvNavArrow(ByVal hDC As Long, uBtn As RECT, ByVal bRight As Boolean, ByVal bDisabled As Boolean, ByVal bEnd As Boolean, ByVal lBandH As Long)
    Dim lApex           As Long
    Dim lCenter         As Long
    Dim lHalf           As Long
    Dim lOfs            As Long

    lHalf = lBandH \ 4
    If bEnd Then
        lOfs = lBandH \ 4 + lBandH \ 7 + lBandH \ 5
    Else
        lOfs = (uBtn.Right - uBtn.Left - lHalf - 1) \ 2 + 1
    End If
    lCenter = uBtn.Top + (uBtn.Bottom - uBtn.Top) \ 2 - 1
    If bRight Then
        lApex = uBtn.Right - 1 - lOfs
    Else
        lApex = uBtn.Left + lOfs
    End If
    If bDisabled Then
        pvNavTriangle hDC, lApex + 1, lCenter + 1, bRight, vb3DHighlight, lHalf
        pvNavTriangle hDC, lApex, lCenter, bRight, vb3DShadow, lHalf
    Else
        pvNavTriangle hDC, lApex, lCenter, bRight, vbBlack, lHalf
    End If
End Sub

Private Sub pvNavTriangle(ByVal hDC As Long, ByVal lApex As Long, ByVal lCenter As Long, ByVal bRight As Boolean, ByVal clrFill As OLE_COLOR, ByVal lHalf As Long)
    Dim lIdx            As Long
    Dim lX              As Long

    For lIdx = 0 To lHalf
        If bRight Then
            lX = lApex - lIdx
        Else
            lX = lApex + lIdx
        End If
        pvFillRect hDC, lX, lCenter - lIdx, lX + 1, lCenter + lIdx + 1, clrFill
    Next
End Sub

Private Sub pvPaintNavigator(ByVal hDC As Long)
    Dim uNav            As UcsNavLayout
    Dim hPrevFont       As Long
    Dim uBox            As RECT
    Dim lAtFirst        As Long
    Dim lAtLast         As Long
    Dim lBarX           As Long
    Dim lBarW           As Long
    Dim lBarY           As Long
    Dim lBarHalf        As Long
    Dim lTextH          As Long
    Dim lTextTop        As Long

    pvFillRect hDC, 0, UserControl.ScaleHeight - GetSystemMetrics(SM_CYHSCROLL), UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.BackColor
    If Not m_bRecordNavigator Then
        Exit Sub
    End If
    If Not pvNavLayout(uNav) Then
        Exit Sub
    End If
    pvFillRect hDC, 0, uNav.BandTop, uNav.Width, uNav.BandTop + uNav.BandH, UserControl.BackColor
    lAtFirst = 0
    If m_lRow <= 1 Then
        lAtFirst = DFCS_INACTIVE
    End If
    lAtLast = 0
    If m_lRow >= RowCount Then
        lAtLast = DFCS_INACTIVE
    End If
    pvNavButton hDC, uNav.BtnFirst
    pvNavButton hDC, uNav.BtnPrev
    pvNavButton hDC, uNav.BtnNext
    pvNavButton hDC, uNav.BtnLast
    pvNavArrow hDC, uNav.BtnFirst, False, False, True, uNav.BandH
    pvNavArrow hDC, uNav.BtnPrev, False, (m_lRow <= 1), False, uNav.BandH
    pvNavArrow hDC, uNav.BtnNext, True, (m_lRow >= RowCount), False, uNav.BandH
    pvNavArrow hDC, uNav.BtnLast, True, False, True, uNav.BandH
    lBarX = (uNav.BandH + 2) \ 4
    lBarW = uNav.BandH \ 7
    lBarHalf = uNav.BandH \ 4
    lBarY = uNav.BtnFirst.Top + (uNav.BtnFirst.Bottom - uNav.BtnFirst.Top) \ 2 - 1 - lBarHalf
    pvFillRect hDC, uNav.BtnFirst.Left + lBarX, lBarY, uNav.BtnFirst.Left + lBarX + lBarW, lBarY + 2 * lBarHalf + 1, vbBlack
    pvFillRect hDC, uNav.BtnLast.Right - 1 - lBarX - lBarW, lBarY, uNav.BtnLast.Right - 1 - lBarX, lBarY + 2 * lBarHalf + 1, vbBlack
    uBox = uNav.Box
    Call DrawEdge(hDC, uBox, EDGE_SUNKEN, BF_RECT)
    pvFillRect hDC, uNav.Box.Left + 2, uNav.Box.Top + 2, uNav.Box.Right - 2, uNav.Box.Bottom - 2, vbWindowBackground
    hPrevFont = pvSelectFont(hDC, m_oFont)
    lTextH = FontTextMetrics(m_oFont, hDC).tmHeight
    lTextTop = uNav.BandTop + (uNav.BandH - lTextH + 1) \ 2
    pvDrawText hDC, uNav.Prefix, uNav.PrefixX, lTextTop, uNav.PrefixX + pvTextWidth(hDC, uNav.Prefix), lTextTop + lTextH, vbButtonText, UserControl.BackColor, jgexAlignLeft, uNav.PrefixX, uNav.Width
    pvDrawText hDC, uNav.Middle, uNav.MiddleX, lTextTop, uNav.MiddleX + pvTextWidth(hDC, uNav.Middle), lTextTop + lTextH, vbButtonText, UserControl.BackColor, jgexAlignLeft, uNav.MiddleX, uNav.Width
    pvDrawText hDC, CStr(m_lRow), uNav.Box.Left + 2, uNav.Box.Top + 2, uNav.Box.Right - 4, uNav.Box.Top + 2 + lTextH, m_clrForeColor, vbWindowBackground, jgexAlignRight, uNav.Box.Left + 2, uNav.Box.Right - 2
    Call SelectObject(hDC, hPrevFont)
End Sub

Private Sub pvLayoutGrid()
    Dim lBandH          As Long

    If m_bBandVisible Then
        lBandH = GetSystemMetrics(SM_CYHSCROLL)
    End If
    picGrid.Move 0, 0, Clamp(UserControl.ScaleWidth, lMin:=0), Clamp(UserControl.ScaleHeight - lBandH, lMin:=0)
End Sub

Private Sub pvLayoutHScroll(ByVal bNeedH As Boolean)
    Dim lBandH          As Long
    Dim lNavW           As Long
    Dim uNav            As UcsNavLayout

    If Not bNeedH Then
        hsbGrid.Visible = False
        m_hWndHScroll = 0
        Exit Sub
    End If
    lBandH = GetSystemMetrics(SM_CYHSCROLL)
    If pvNavLayout(uNav) Then
        lNavW = uNav.Width
    End If
    hsbGrid.Move lNavW, UserControl.ScaleHeight - lBandH, Clamp(picGrid.ScaleWidth - lNavW, lMin:=0), lBandH
    hsbGrid.Visible = True
    m_hWndHScroll = hsbGrid.hWnd
End Sub

Private Sub pvUpdateScrollBars()
    Dim lTopH           As Long
    Dim lAvailH         As Long
    Dim lAvailW         As Long
    Dim lFullH          As Long
    Dim lFullW          As Long
    Dim lHdrW           As Long
    Dim lColsW          As Long
    Dim lStyle          As Long
    Dim lNewStyle       As Long
    Dim bNeedV          As Boolean
    Dim bNeedH          As Boolean
    Dim uSi             As SCROLLINFO

    If m_bScrollUpdating Then
        Exit Sub
    End If
    m_bScrollUpdating = True
    lTopH = pvTopHeight()
    lHdrW = pvBlockLeft()
    lColsW = lHdrW + pvTotalColWidth()
    lFullH = picGrid.ScaleHeight
    If m_bBandVisible Then
        lFullH = lFullH + GetSystemMetrics(SM_CYHSCROLL)
    End If
    lFullW = picGrid.ScaleWidth
    If (GetWindowLong(picGrid.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0 Then
        lFullW = lFullW + GetSystemMetrics(SM_CXVSCROLL)
    End If
    lAvailH = lFullH - lTopH
    If m_bRecordNavigator Then
        lAvailH = lAvailH - GetSystemMetrics(SM_CYHSCROLL)
    End If
    lAvailW = lFullW
    bNeedV = (m_lRowHeight > 0 And RowCount > 0 And RowCount * m_lRowHeight > lAvailH)
    bNeedH = (lColsW > lAvailW)
    If bNeedV Then
        lAvailW = lFullW - GetSystemMetrics(SM_CXVSCROLL)
        bNeedH = (lColsW > lAvailW)
    End If
    If bNeedH And Not m_bRecordNavigator Then
        lAvailH = lAvailH - GetSystemMetrics(SM_CYHSCROLL)
        bNeedV = (m_lRowHeight > 0 And RowCount > 0 And RowCount * m_lRowHeight > lAvailH)
        If bNeedV Then
            lAvailW = lFullW - GetSystemMetrics(SM_CXVSCROLL)
        End If
    End If
    lStyle = GetWindowLong(picGrid.hWnd, GWL_STYLE)
    lNewStyle = lStyle And Not WS_VSCROLL And Not WS_HSCROLL
    If bNeedV Then
        lNewStyle = lNewStyle Or WS_VSCROLL
    End If
    If lNewStyle <> lStyle Then
        Call SetWindowLong(picGrid.hWnd, GWL_STYLE, lNewStyle)
        Call SetWindowPos(picGrid.hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    End If
    If bNeedV Then
        With uSi
            .cbSize = Len(uSi)
            .fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
            .nMax = RowCount - 1
            If m_lRowHeight > 0 Then
                .nPage = lAvailH \ m_lRowHeight
            End If
            m_lFirstItem = Clamp(m_lFirstItem, 1, RowCount - .nPage + 1)
            .nPos = m_lFirstItem - 1
        End With
        Call SetScrollInfo(picGrid.hWnd, SB_VERT, uSi, 1)
    End If
    m_bBandVisible = bNeedH Or m_bRecordNavigator
    pvLayoutGrid
    pvLayoutHScroll bNeedH
    If bNeedH Then
        With uSi
            .cbSize = Len(uSi)
            .fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
            .nMin = pvFrozenCount() + 1
            .nMax = pvMaxLeftCol()
            If .nMax < .nMin Then
                .nMax = .nMin
            End If
            .nPage = 1
            m_lLeftCol = Clamp(m_lLeftCol, .nMin, .nMax)
            .nPos = m_lLeftCol
        End With
        Call SetScrollInfo(hsbGrid.hWnd, SB_CTL, uSi, 1)
    ElseIf m_lLeftCol > 1 Then
        m_lLeftCol = 1
    End If
    m_bScrollUpdating = False
End Sub

Private Function pvTotalColWidth() As Long
    Dim oCol            As JSColumn

    For Each oCol In m_oColumns
        If oCol.Visible Then
            pvTotalColWidth = pvTotalColWidth + pvColWidth(oCol)
        End If
    Next
End Function

Private Function pvColWidth(oCol As JSColumn) As Long
    Dim lIdx            As Long
    Dim oItem           As JSColumn
    Dim lTotal          As Long
    Dim lAvail          As Long
    Dim lCum            As Long
    Dim lPrev           As Long

    If Not m_bColumnAutoResize Then
        pvColWidth = ToPixels(oCol.Width)
        Exit Function
    End If
    lAvail = picGrid.ScaleWidth - pvBlockLeft()
    For lIdx = 1 To m_oColumns.Count
        Set oItem = m_oColumns.ItemByPosition(lIdx)
        If oItem.Visible Then
            lTotal = lTotal + ToPixels(oItem.Width)
        End If
    Next
    If lTotal <= 0 Or lAvail <= 0 Then
        pvColWidth = ToPixels(oCol.Width)
        Exit Function
    End If
    For lIdx = 1 To m_oColumns.Count
        Set oItem = m_oColumns.ItemByPosition(lIdx)
        If oItem.Visible Then
            lPrev = (lCum * lAvail + lTotal \ 2) \ lTotal
            lCum = lCum + ToPixels(oItem.Width)
            If oItem Is oCol Then
                pvColWidth = (lCum * lAvail + lTotal \ 2) \ lTotal - lPrev
                Exit Function
            End If
        End If
    Next
    pvColWidth = ToPixels(oCol.Width)
End Function

Private Function pvColOrder() As Variant
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim aOrder()        As Long
    Dim lCount          As Long
    Dim lFrozen         As Long

    ReDim aOrder(0 To m_oColumns.Count) As Long
    lFrozen = pvFrozenCount()
    For lIdx = 1 To lFrozen
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            aOrder(lCount) = lIdx
            lCount = lCount + 1
        End If
    Next
    For lIdx = pvFirstCol() To m_oColumns.Count
        If lIdx > lFrozen Then
            Set oCol = m_oColumns.ItemByPosition(lIdx)
            If oCol.Visible Then
                aOrder(lCount) = lIdx
                lCount = lCount + 1
            End If
        End If
    Next
    If lCount = 0 Then
        pvColOrder = Array()
        Exit Function
    End If
    ReDim Preserve aOrder(0 To lCount - 1) As Long
    pvColOrder = aOrder
End Function

Private Function pvOrderMax(vOrder As Variant) As Long
    '--- -1 for an empty order, so the paint loops simply do not run
    pvOrderMax = -1
    If IsArray(vOrder) Then
        pvOrderMax = UBound(vOrder)
    End If
End Function

Private Function pvFrozenCount() As Long
    '--- more frozen columns than there are leaves nothing to scroll
    pvFrozenCount = m_lFrozenColumns
    If pvFrozenCount < 0 Then
        pvFrozenCount = 0
    End If
    If pvFrozenCount > m_oColumns.Count Then
        pvFrozenCount = m_oColumns.Count
    End If
End Function

Private Function pvColsFitFrom(ByVal lWidth As Long, ByVal lStart As Long) As Long
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim lCum            As Long

    For lIdx = lStart To m_oColumns.Count
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            lCum = lCum + pvColWidth(oCol)
            If lCum >= lWidth Then
                Exit Function
            End If
            pvColsFitFrom = pvColsFitFrom + 1
        End If
    Next
    If pvColsFitFrom < 1 Then
        pvColsFitFrom = 1
    End If
End Function

Private Function pvMaxLeftCol() As Long
    Dim lWidth          As Long
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim lCum            As Long

    lWidth = pvScrollableWidth()
    pvMaxLeftCol = m_oColumns.Count
    For lIdx = m_oColumns.Count To pvFrozenCount() + 1 Step -1
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            lCum = lCum + pvColWidth(oCol)
            If lCum >= lWidth Then
                Exit For
            End If
            pvMaxLeftCol = lIdx
        End If
    Next
    If pvMaxLeftCol < 1 Then
        pvMaxLeftCol = 1
    End If
End Function

Private Function pvColsFitBefore(ByVal lWidth As Long, ByVal lStart As Long) As Long
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim lCum            As Long

    For lIdx = lStart - 1 To pvFrozenCount() + 1 Step -1
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            lCum = lCum + pvColWidth(oCol)
            If lCum >= lWidth Then
                Exit For
            End If
            pvColsFitBefore = pvColsFitBefore + 1
        End If
    Next
    If pvColsFitBefore < 1 Then
        pvColsFitBefore = 1
    End If
End Function

Private Function pvColor(ByVal clrValue As OLE_COLOR) As Long
    Call OleTranslateColor(clrValue, 0, pvColor)
End Function

Private Sub pvSelColors(clrBack As OLE_COLOR, clrFore As OLE_COLOR)
    Dim oStyle          As JSFormatStyle

    Set oStyle = m_oFormatStyles.frItemOrNothing("SelectedRow")
    If oStyle Is Nothing Then
        clrBack = vbHighlight
        clrFore = vbHighlightText
    Else
        clrBack = oStyle.BackColor
        clrFore = oStyle.ForeColor
    End If
End Sub

Private Sub pvStampedLine(ByVal hDC As Long, ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long, ByVal lAnchor As Long, ByVal lPeriod As Long, ByVal lOn As Long, ByVal clrLine As OLE_COLOR, ByVal clrGap As OLE_COLOR)
    Dim hBrush          As Long
    Dim hPrevBrush      As Long
    Dim lIdx            As Long

    '--- the gaps go down first: what shows between the marks is the gap colour
    '--- rather than whatever the run is drawn over
    If lY1 = lY2 Then
        pvFillRect hDC, lX1, lY1, lX2, lY1 + 1, clrGap
    Else
        pvFillRect hDC, lX1, lY1, lX1 + 1, lY2, clrGap
    End If
    hBrush = CreateSolidBrush(pvColor(clrLine))
    hPrevBrush = SelectObject(hDC, hBrush)
    If lY1 = lY2 Then
        For lIdx = lX1 To lX2 - 1
            If Abs(lIdx - lAnchor) Mod lPeriod < lOn Then
                Call PatBlt(hDC, lIdx, lY1, 1, 1, PATCOPY)
            End If
        Next
    Else
        For lIdx = lY1 To lY2 - 1
            If Abs(lIdx - lAnchor) Mod lPeriod < lOn Then
                Call PatBlt(hDC, lX1, lIdx, 1, 1, PATCOPY)
            End If
        Next
    End If
    Call SelectObject(hDC, hPrevBrush)
    Call DeleteObject(hBrush)
End Sub

Private Sub pvFillRect(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal clrFill As OLE_COLOR)
    Dim uRect           As RECT
    Dim hBrush          As Long

    If lRight <= lLeft Or lBottom <= lTop Then
        Exit Sub
    End If
    uRect.Left = lLeft
    uRect.Top = lTop
    uRect.Right = lRight
    uRect.Bottom = lBottom
    hBrush = CreateSolidBrush(pvColor(clrFill))
    Call FillRect(hDC, uRect, hBrush)
    Call DeleteObject(hBrush)
End Sub

Private Sub pvLine(ByVal hDC As Long, ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long, ByVal clrLine As OLE_COLOR, ByVal lPenStyle As Long, Optional ByVal DashAnchor As Long = -1, Optional ByVal GapColor As Long = -1)
    Dim hPen            As Long
    Dim hPrevPen        As Long
    Dim clrGap          As OLE_COLOR

    If lPenStyle = PS_DOT Or lPenStyle = PS_DASH Then
        clrGap = m_clrBackColorBkg
        If GapColor <> -1 Then
            clrGap = GapColor
        End If
        If lPenStyle = PS_DOT Then
            pvStampedLine hDC, lX1, lY1, lX2, lY2, IIf(lY1 = lY2, lY1, lX1), 2, 1, clrLine, clrGap
        Else
            If DashAnchor < 0 Then
                DashAnchor = IIf(lY1 = lY2, lX2 - 1, lY1)
            End If
            pvStampedLine hDC, lX1, lY1, lX2, lY2, DashAnchor, 6, 3, clrLine, clrGap
        End If
        Exit Sub
    End If
    hPen = CreatePen(lPenStyle, 1, pvColor(clrLine))
    hPrevPen = SelectObject(hDC, hPen)
    Call MoveToEx(hDC, lX1, lY1, 0)
    Call LineTo(hDC, lX2, lY2)
    Call SelectObject(hDC, hPrevPen)
    Call DeleteObject(hPen)
End Sub

Private Sub pvDrawText(ByVal hDC As Long, sText As String, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal clrText As OLE_COLOR, ByVal clrBack As OLE_COLOR, ByVal eAlign As jgexAlignmentConstants, ByVal ClipLeft As Long, ByVal ClipRight As Long, Optional ByVal bWordWrap As Boolean, Optional ByVal bEllipsis As Boolean)
    Dim uRect           As RECT
    Dim lFlags          As Long
    Dim lSaved          As Long

    uRect.Left = lLeft
    uRect.Top = lTop
    uRect.Right = lRight
    uRect.Bottom = lBottom
    If bWordWrap Then
        uRect.Top = lTop + 1
        uRect.Bottom = lBottom - 1
        lFlags = DT_WORDBREAK Or DT_NOPREFIX
    Else
        lFlags = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
        If bEllipsis Then
            lFlags = lFlags Or DT_END_ELLIPSIS
        End If
    End If
    Select Case eAlign
    Case jgexAlignCenter
        lFlags = lFlags Or DT_CENTER
    Case jgexAlignRight
        lFlags = lFlags Or DT_RIGHT
    End Select
    lSaved = SaveDC(hDC)
    Call IntersectClipRect(hDC, ClipLeft, uRect.Top, ClipRight, uRect.Bottom)
    lFlags = lFlags Or DT_NOCLIP
    Call SetBkMode(hDC, OPAQUE)
    Call SetBkColor(hDC, pvColor(clrBack))
    Call SetTextColor(hDC, pvColor(clrText))
    Call DrawText(hDC, StrPtr(sText), Len(sText), uRect, lFlags)
    Call RestoreDC(hDC, lSaved)
End Sub

Private Function pvSelectFont(ByVal hDC As Long, pFont As IFont) As Long
    pvSelectFont = SelectObject(hDC, pFont.hFont)
End Function

Private Sub pvInheritAmbientFont()
    Const FUNC_NAME     As String = "pvInheritAmbientFont"

    On Error GoTo EH
    Set m_oFont = CloneFont(Ambient.Font)
    Set m_oColumnHeaderFont = CloneFont(Ambient.Font)
    m_oFont_FontChanged vbNullString
    m_oColumnHeaderFont_FontChanged vbNullString
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvSubclass()
    Set m_pSubclassPic = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.ControlSubclassProc(0, 0, 0, 0, 0))
    Set m_pSubclassCtl = InitSubclassingThunk(UserControl.hWnd, Me, pvAddressOfSubclassProc.ControlSubclassProc(0, 0, 0, 0, 0))
End Sub

Public Function EditSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute EditSubclassProc.VB_MemberFlags = "40"
    #If hWnd Then '--- touch args
    #End If
    Dim nKeyCode        As Integer
    Dim nShift          As Integer

    Select Case wMsg
    Case WM_KEYDOWN
        nKeyCode = LoWordInt(wParam)
        If pvEditKeyLeaves(nKeyCode) Then
            Call SendMessage(picGrid.hWnd, WM_KEYDOWN, wParam, lParam)
            Handled = True
        Else
            nShift = pvShiftState()
            RaiseEvent KeyDown(nKeyCode, nShift)
            Select Case nKeyCode
            Case vbKeyReturn
                pvEditEnd
                pvEditCommit
                If m_lRow < RowCount Then
                    pvNavigate m_lRow + 1, m_lCol, 0, False
                End If
                Handled = True
            Case vbKeyEscape
                pvEditEnd bCancel:=True
                Handled = True
            End Select
        End If
    Case WM_CHAR
        RaiseEvent KeyPress(LoWordInt(wParam))
    Case WM_KEYUP
        RaiseEvent KeyUp(LoWordInt(wParam), pvShiftState())
    Case WM_SETFOCUS
        If Not m_bGridFocus Then
            m_bGridFocus = True
            pvInvalidate SkipScroll:=True
        End If
    Case WM_KILLFOCUS
        If wParam <> picGrid.hWnd And wParam <> UserControl.hWnd Then
            pvEditFocusLost
            If m_bGridFocus Then
                m_bGridFocus = False
                pvInvalidate SkipScroll:=True
            End If
        End If
    End Select
End Function

Public Function ControlSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute ControlSubclassProc.VB_MemberFlags = "40"
    Dim nKeyCode        As Integer
    Dim nShift          As Integer
    Dim nKeyAscii       As Integer
    Dim oGroup          As JSGroup
    Dim oSizeCol        As JSColumn
    Dim oEditCol        As JSColumn

    If hWnd = UserControl.hWnd Then
        Select Case wMsg
        Case WM_LBUTTONDOWN
            pvOnNavigatorClick GetXLParam(lParam), GetYLParam(lParam)
        Case WM_MOUSEACTIVATE
            If pvHitScrollBar() Then
                ControlSubclassProc = MA_NOACTIVATE
                Handled = True
            End If
        Case WM_PAINT
            pvOnPaint hWnd
            Handled = True
        Case WM_ERASEBKGND
            If m_bRecordNavigator Then
                ControlSubclassProc = 1
                Handled = True
            End If
        Case WM_CTLCOLORSCROLLBAR
            Handled = True
        Case WM_HSCROLL
            If lParam = hsbGrid.hWnd Then
                pvOnHScroll LoWord(wParam), HiWord(wParam)
                Handled = True
            End If
        End Select
        GoTo QH
    End If
    Select Case wMsg
    Case WM_PAINT
        pvOnPaint hWnd
        Handled = True
    Case WM_ERASEBKGND
        ControlSubclassProc = 1
        Handled = True
    Case WM_SETCURSOR
        If pvSetCursor() Then
            ControlSubclassProc = 1
            Handled = True
        End If
    Case WM_CANCELMODE
        If Not m_oSizeCol Is Nothing Then
            pvEndColSize bCancel:=True
        End If
        If Not m_oDragCol Is Nothing Or Not m_oDragGroup Is Nothing Then
            pvEndColDrag bCancel:=True
        End If
    Case WM_TIMER
        If wParam = DRAG_SCROLL_ID Then
            LeftCol = m_lLeftCol + m_lDragScroll
            Handled = True
        End If
    Case WM_SETFOCUS
        If Not m_bGridFocus Then
            m_bGridFocus = True
            pvInvalidate SkipScroll:=True
        End If
    Case WM_KILLFOCUS
        If m_bGridFocus And (wParam <> m_hWndEdit Or m_hWndEdit = 0) Then
            m_bGridFocus = False
            pvInvalidate SkipScroll:=True
        End If
    Case WM_VSCROLL
        pvOnVScroll LoWord(wParam), HiWord(wParam)
        Handled = True
    Case WM_KEYDOWN
        nKeyCode = LoWordInt(wParam)
        nShift = pvShiftState()
        RaiseEvent KeyDown(nKeyCode, nShift)
        pvOnKeyDown nKeyCode, nShift
    Case WM_LBUTTONDOWN
        pvOnLButtonDown GetXLParam(lParam), GetYLParam(lParam), pvMouseShift(wParam)
        RaiseEvent MouseDown(vbLeftButton, pvMouseShift(wParam), GetXLParam(lParam) * Screen.TwipsPerPixelX, GetYLParam(lParam) * Screen.TwipsPerPixelY)
        If m_hWndEdit <> 0 Then
            ControlSubclassProc = CallNextSubclassProc(m_pSubclassPic, hWnd, wMsg, wParam, lParam)
            Handled = True
            Call SetFocusApi(m_hWndEdit)
        End If
    Case WM_LBUTTONUP
        m_bClickOpenedEdit = False
        If Not m_oSizeCol Is Nothing Then
            pvEndColSize bCancel:=False
        ElseIf m_bDragging Then
            pvOnColDrag GetXLParam(lParam), GetYLParam(lParam)
            pvEndColDrag bCancel:=False
        Else
            If Not m_oDragGroup Is Nothing Then
                Set oGroup = m_oDragGroup
                Set m_oDragGroup = Nothing
                pvInvalidateHeaders
                RaiseEvent GroupByBoxHeaderClick(oGroup)
                If m_bAutomaticSort Then
                    pvAutoSort m_oColumns.Item(oGroup.ColIndex)
                End If
            End If
            If Not m_oDragCol Is Nothing Then
                Set oSizeCol = m_oDragCol
                Set m_oDragCol = Nothing
                pvInvalidateHeaders
                RaiseEvent ColumnHeaderClick(oSizeCol)
                If m_bAutomaticSort Then
                    pvAutoSort oSizeCol
                End If
            End If
            RaiseEvent MouseUp(vbLeftButton, pvMouseShift(wParam), GetXLParam(lParam) * Screen.TwipsPerPixelX, GetYLParam(lParam) * Screen.TwipsPerPixelY)
            If m_bEatClick Then
                m_bEatClick = False
            Else
                RaiseEvent Click
            End If
        End If
    Case WM_LBUTTONDBLCLK
        Set oSizeCol = pvColDividerAt(GetXLParam(lParam), GetYLParam(lParam))
        If oSizeCol Is Nothing Then
            If pvGroupRowDblClk(GetYLParam(lParam)) Then
                m_bEatClick = True
            End If
            RaiseEvent DblClick
        Else
            frColAutoSize oSizeCol
        End If
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(IIf(m_oSizeCol Is Nothing, pvMouseButton(wParam), 0), pvMouseShift(wParam), GetXLParam(lParam) * Screen.TwipsPerPixelX, GetYLParam(lParam) * Screen.TwipsPerPixelY)
        If Not m_oSizeCol Is Nothing Then
            pvOnColSize GetXLParam(lParam)
        ElseIf (Not m_oDragCol Is Nothing Or Not m_oDragGroup Is Nothing) And (wParam And MK_LBUTTON) <> 0 Then
            pvOnColDrag GetXLParam(lParam), GetYLParam(lParam)
        ElseIf (wParam And MK_LBUTTON) <> 0 Then
            pvOnMouseDrag GetYLParam(lParam)
        End If
    Case WM_CHAR
        nKeyAscii = LoWordInt(wParam)
        If nKeyAscii >= 32 And m_hWndEdit = 0 And Not m_bEditing Then
            Set oEditCol = pvColByPosition(m_lCol)
            If Not oEditCol Is Nothing Then
                If oEditCol.EditType <> jgexEditTextBox Then
                    Set oEditCol = Nothing
                ElseIf Not pvEditInit(m_lRow, m_lCol, bSelectAll:=True) Then
                    Set oEditCol = Nothing
                ElseIf m_hWndEdit = 0 Then
                    Set oEditCol = Nothing
                End If
            End If
        End If
        If oEditCol Is Nothing Then
            RaiseEvent KeyPress(nKeyAscii)
        Else
            Call SendMessage(m_hWndEdit, WM_CHAR, wParam, lParam)
        End If
    Case WM_KEYUP
        RaiseEvent KeyUp(LoWordInt(wParam), pvShiftState())
    Case WM_COMMAND
        If lParam = m_hWndEdit And m_hWndEdit <> 0 Then
            If (wParam \ &H10000) = EN_CHANGE And Not m_bInEditSetup Then
                RaiseEvent Change
            End If
        End If
    End Select
QH:
    '--- note: performance optimization for design-time subclassing
    If Not Handled And ThunkPrivateData(m_pSubclassPic) = EBMODE_DESIGN Then
        Handled = True
        ControlSubclassProc = CallNextSubclassProc(m_pSubclassPic, hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Function pvShiftState() As Integer
    If GetKeyState(vbKeyShift) < 0 Then
        pvShiftState = pvShiftState Or vbShiftMask
    End If
    If GetKeyState(vbKeyControl) < 0 Then
        pvShiftState = pvShiftState Or vbCtrlMask
    End If
    If GetKeyState(vbKeyMenu) < 0 Then
        pvShiftState = pvShiftState Or vbAltMask
    End If
End Function

Private Function pvMouseShift(ByVal wParam As Long) As Integer
    If (wParam And MK_SHIFT) <> 0 Then
        pvMouseShift = pvMouseShift Or vbShiftMask
    End If
    If (wParam And MK_CONTROL) <> 0 Then
        pvMouseShift = pvMouseShift Or vbCtrlMask
    End If
    If GetKeyState(vbKeyMenu) < 0 Then
        pvMouseShift = pvMouseShift Or vbAltMask
    End If
End Function

Private Function pvMouseButton(ByVal wParam As Long) As Integer
    If (wParam And MK_LBUTTON) <> 0 Then
        pvMouseButton = pvMouseButton Or vbLeftButton
    End If
    If (wParam And MK_RBUTTON) <> 0 Then
        pvMouseButton = pvMouseButton Or vbRightButton
    End If
End Function

Private Sub pvOnMouseDrag(ByVal lY As Long)
    Dim lTopHdr         As Long
    Dim lRow            As Long

    If Not m_bMultiSelect Or m_lRowHeight <= 0 Or m_bClickOpenedEdit Or Not m_bDragSelect Then
        Exit Sub
    End If
    lTopHdr = pvTopHeight()
    If lY < lTopHdr Then
        Exit Sub
    End If
    lRow = pvClampRow(m_lFirstItem + (lY - lTopHdr) \ m_lRowHeight)
    If lRow <> m_lRow Then
        pvSetRow lRow
        pvSetRangeSel m_lSelAnchor, lRow
        EnsureVisible m_lRow
    End If
End Sub

Private Function pvColDividerAt(ByVal lX As Long, ByVal lY As Long) As JSColumn
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim oCol            As JSColumn
    Dim lCum            As Long

    If Not m_bColumnHeaders Or lY < pvGroupByBoxHeight() Or lY >= pvRowsTop() Then
        Exit Function
    End If
    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oCol.Visible Then
            lCum = lCum + pvColWidth(oCol)
            If Abs(lX - lCum) <= DIVIDER_GRAB_W Then
                If oCol.AllowSizing Then
                    Set pvColDividerAt = oCol
                End If
                Exit Function
            End If
        End If
    Next
End Function

Private Function pvSetCursor() As Boolean
    Dim uPt             As POINTAPI
    Dim lCursor         As Long
    Dim hCursor         As Long

    If m_bDragging Then
        lCursor = IDC_SIZEALL
    ElseIf Not m_oSizeCol Is Nothing Then
        lCursor = IDC_SIZEWE
    Else
        Call GetCursorPos(uPt)
        Call ScreenToClient(picGrid.hWnd, uPt)
        If Not pvColDividerAt(uPt.X, uPt.Y) Is Nothing Then
            lCursor = IDC_SIZEWE
        ElseIf pvRowSelZoneAt(uPt.X, uPt.Y) Then
            hCursor = pvRowSelCursor()
        End If
    End If
    If lCursor <> 0 Then
        hCursor = LoadCursor(0, lCursor)
    End If
    If hCursor <> 0 Then
        Call SetCursor(hCursor)
        pvSetCursor = True
    End If
End Function

Private Function pvRowSelCursor() As Long
    Dim uInfo           As ICONINFO
    Dim uBmp            As BITMAP
    Dim hDCSrc          As Long
    Dim hDCDst          As Long
    Dim hMask           As Long
    Dim hColor          As Long
    Dim hScreen         As Long

    '--- the original ships its right-pointing arrow as a resource; ours is
    '--- the stock arrow mirrored on the fly, hotspot and all -- the mask
    '--- (and the colour plane, where one exists) blitted at negative width
    If m_hCurRowSel = 0 Then
        Call GetIconInfo(LoadCursor(0, IDC_ARROW), uInfo)
        Call GetObjectAPI(uInfo.hbmMask, LenB(uBmp), uBmp)
        hDCSrc = CreateCompatibleDC(0)
        hDCDst = CreateCompatibleDC(0)
        hMask = CreateBitmap(uBmp.bmWidth, uBmp.bmHeight, 1, 1, ByVal 0&)
        Call SelectObject(hDCSrc, uInfo.hbmMask)
        Call SelectObject(hDCDst, hMask)
        Call StretchBlt(hDCDst, 0, 0, uBmp.bmWidth, uBmp.bmHeight, hDCSrc, uBmp.bmWidth - 1, 0, -uBmp.bmWidth, uBmp.bmHeight, SRCCOPY)
        If uInfo.hbmColor <> 0 Then
            hScreen = GetDC(0)
            hColor = CreateCompatibleBitmap(hScreen, uBmp.bmWidth, uBmp.bmHeight)
            Call ReleaseDC(0, hScreen)
            Call SelectObject(hDCSrc, uInfo.hbmColor)
            Call SelectObject(hDCDst, hColor)
            Call StretchBlt(hDCDst, 0, 0, uBmp.bmWidth, uBmp.bmHeight, hDCSrc, uBmp.bmWidth - 1, 0, -uBmp.bmWidth, uBmp.bmHeight, SRCCOPY)
        End If
        Call DeleteDC(hDCSrc)
        Call DeleteDC(hDCDst)
        Call DeleteObject(uInfo.hbmMask)
        If uInfo.hbmColor <> 0 Then
            Call DeleteObject(uInfo.hbmColor)
        End If
        uInfo.fIcon = 0
        uInfo.xHotspot = uBmp.bmWidth - 1 - uInfo.xHotspot
        uInfo.hbmMask = hMask
        uInfo.hbmColor = hColor
        m_hCurRowSel = CreateIconIndirect(uInfo)
        Call DeleteObject(hMask)
        If hColor <> 0 Then
            Call DeleteObject(hColor)
        End If
    End If
    pvRowSelCursor = m_hCurRowSel
End Function

Private Sub pvOnColSize(ByVal lX As Long)
    Dim lWidth          As Long

    pvSetCursor
    lWidth = m_lSizeStartW + lX - m_lSizeStartX
    If lWidth < MIN_COL_W Then
        lWidth = MIN_COL_W
    End If
    If lWidth <> pvColWidth(m_oSizeCol) Then
        m_oSizeCol.frWidthPx = lWidth
        pvInvalidate
    End If
End Sub

Private Function pvAutoSizeWidth(ByVal lColIndex As Long) As Long
    Dim lIdx            As Long
    Dim hDC             As Long
    Dim hPrevFont       As Long
    Dim lWidth          As Long
    Dim oRowData        As JSRowData

    pvSyncProjection
    pvPopulateWindow
    hDC = GetDC(picGrid.hWnd)
    If m_bColumnHeaders Then
        hPrevFont = pvSelectFont(hDC, m_oColumnHeaderFont)
        '--- nine, where a value takes five: probed against the original, which
        '--- fits a caption into four pixels more than the text it draws
        pvAutoSizeWidth = pvTextWidth(hDC, m_oColumns.Item(lColIndex).Caption) + 9
        Call SelectObject(hDC, hPrevFont)
    End If
    hPrevFont = pvSelectFont(hDC, m_oFont)
    For lIdx = 1 To m_lWindowCount
        Set oRowData = m_aWindow(lIdx)
        If Not oRowData Is Nothing Then
            If oRowData.RowIndex > 0 Then
                lWidth = pvTextWidth(hDC, pvCellText(oRowData, lColIndex)) + 5
                If lWidth > pvAutoSizeWidth Then
                    pvAutoSizeWidth = lWidth
                End If
            End If
        End If
    Next
    Call SelectObject(hDC, hPrevFont)
    Call ReleaseDC(picGrid.hWnd, hDC)
End Function

Private Function pvGBoxDropMark(lLeft As Long, lTop As Long) As Boolean
    If Not m_oDropGroup Is Nothing Then
        If m_oDropGroup Is m_oDragGroup Then
            Exit Function
        End If
        If Not m_oDragGroup Is Nothing Then
            If Not pvDropMoves(m_oDragGroup.Index, m_oDropGroup.Index, m_bDropAfter) Then
                Exit Function
            End If
        End If
        lLeft = m_oDropGroup.frChipRect.Left
        If m_bDropAfter Then
            lLeft = m_oDropGroup.frChipRect.Right
        End If
        lTop = m_oDropGroup.frChipRect.Top
        pvGBoxDropMark = True
        Exit Function
    End If
    If Not m_oDragGroup Is Nothing Then
        Exit Function
    End If
    lLeft = CHIP_LEFT
    lTop = CHIP_TOP
    pvGBoxDropMark = True
End Function

Private Function pvGetDragColumn() As JSColumn
    If Not m_oDragCol Is Nothing Then
        Set pvGetDragColumn = m_oDragCol
    ElseIf Not m_oDragGroup Is Nothing Then
        If m_oDragGroup.ColIndex >= 1 And m_oDragGroup.ColIndex <= m_oColumns.Count Then
            Set pvGetDragColumn = m_oColumns.Item(m_oDragGroup.ColIndex)
        End If
    End If
End Function

Private Sub pvOnColDrag(ByVal lX As Long, ByVal lY As Long)
    Dim oCancel         As JSRetBoolean
    Dim lScroll         As Long
    Dim bNewInGBox      As Boolean
    Dim oNewGroup       As JSGroup
    Dim bNewAfter       As Boolean
    Dim oNewCol         As JSColumn

    If Not m_bDragging Then
        If Not m_bAllowColumnDrag Then
            Exit Sub
        End If
        If Abs(lX - m_lDragStartX) <= GetSystemMetrics(SM_CXDOUBLECLK) \ 2 And Abs(lY - m_lDragStartY) <= GetSystemMetrics(SM_CYDOUBLECLK) \ 2 Then
            Exit Sub
        End If
        If Not m_oDragGroup Is Nothing Then
            Set oCancel = New JSRetBoolean
            RaiseEvent BeforeGroupDrag(m_oDragGroup, oCancel)
            If oCancel.Value Then
                Set m_oDragGroup = Nothing
                pvInvalidateHeaders
                Exit Sub
            End If
        ElseIf Not m_oDragCol Is Nothing Then
            Set oCancel = New JSRetBoolean
            RaiseEvent BeforeColumnDrag(m_oDragCol, oCancel)
            If oCancel.Value Then
                Set m_oDragCol = Nothing
                pvInvalidateHeaders
                Exit Sub
            End If
        End If
        m_bDragging = True
        Call SetCapture(picGrid.hWnd)
        pvInvalidateHeaders
    End If
    pvSetCursor
    lScroll = 0
    If lX < 0 Then
        lScroll = -1
    ElseIf lX >= picGrid.ScaleWidth Then
        lScroll = 1
    End If
    If lScroll <> m_lDragScroll Then
        m_lDragScroll = lScroll
        If lScroll = 0 Then
            Call KillTimer(picGrid.hWnd, DRAG_SCROLL_ID)
        Else
            Call SetTimer(picGrid.hWnd, DRAG_SCROLL_ID, DRAG_SCROLL_MS, 0)
            LeftCol = m_lLeftCol + m_lDragScroll
        End If
    End If
    If lScroll = 0 Then
        If m_bGroupByBoxVisible And lY >= 0 And lY < pvGroupByBoxHeight() Then
            bNewInGBox = True
            pvGroupDropTarget lX, oNewGroup, bNewAfter
        ElseIf lY >= 0 And lY < picGrid.ScaleHeight Then
            pvColAtX lX, oNewCol, bNewAfter
        End If
    End If
    If Not oNewCol Is m_oDropCol Or Not oNewGroup Is m_oDropGroup _
            Or bNewInGBox <> m_bDropInGBox Or bNewAfter <> m_bDropAfter Then
        Set m_oDropCol = oNewCol
        Set m_oDropGroup = oNewGroup
        m_bDropInGBox = bNewInGBox
        m_bDropAfter = bNewAfter
        pvInvalidateHeaders
    End If
End Sub

Private Sub pvEndColDrag(ByVal bCancel As Boolean)
    Dim oCol            As JSColumn
    Dim oTargetCol      As JSColumn
    Dim oGroup          As JSGroup
    Dim oTargetGroup    As JSGroup
    Dim bAfter          As Boolean
    Dim oCancel         As JSRetBoolean
    Dim lNewPos         As Long

    Set oCol = m_oDragCol
    Set oTargetCol = m_oDropCol
    Set oGroup = m_oDragGroup
    Set oTargetGroup = m_oDropGroup
    bAfter = m_bDropAfter
    Set m_oDragCol = Nothing
    Set m_oDropCol = Nothing
    Set m_oDragGroup = Nothing
    Set m_oDropGroup = Nothing
    If Not m_bDragging Then
        pvInvalidateHeaders
        Exit Sub
    End If
    m_bDragging = False
    Call ReleaseCapture
    If m_lDragScroll <> 0 Then
        Call KillTimer(picGrid.hWnd, DRAG_SCROLL_ID)
        m_lDragScroll = 0
    End If
    If bCancel Then
        GoTo QH
    End If
    If Not oGroup Is Nothing Then
        If m_bDropInGBox Then
            If Not oTargetGroup Is Nothing Then
                pvMoveGroup oGroup, oTargetGroup, bAfter
            End If
        Else
            pvDeleteGroup oGroup, oTargetCol, bAfter
        End If
        GoTo QH
    End If
    If m_bDropInGBox Then
        pvAddGroup oCol, oTargetGroup, bAfter
        GoTo QH
    End If
    If oTargetCol Is Nothing Or oTargetCol Is oCol Then
        GoTo QH
    End If
    lNewPos = pvDropPosition(oCol.ColPosition, oTargetCol.ColPosition, bAfter)
    If lNewPos = oCol.ColPosition Then
        GoTo QH
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeColMove(oCol, lNewPos, oCancel)
    If Not oCancel.Value Then
        frColMove oCol, lNewPos
        RaiseEvent AfterColMove
    End If
QH:
    m_bDropInGBox = False
    m_bDropAfter = False
    '--- the header that was inverted has to come back either way
    pvInvalidate
End Sub

Private Function pvDropMoves(ByVal lOld As Long, ByVal lTarget As Long, ByVal bAfter As Boolean) As Boolean
    pvDropMoves = (lOld <> lTarget) And (pvDropPosition(lOld, lTarget, bAfter) <> lOld)
End Function

Private Function pvDropPosition(ByVal lOld As Long, ByVal lTarget As Long, ByVal bAfter As Boolean) As Long
    pvDropPosition = lTarget
    If bAfter Then
        pvDropPosition = pvDropPosition + 1
    End If
    If lOld < pvDropPosition Then
        pvDropPosition = pvDropPosition - 1
    End If
End Function

Private Sub pvAddGroup(oCol As JSColumn, oTargetGroup As JSGroup, ByVal bAfter As Boolean)
    Dim oGroup          As JSGroup
    Dim oCancel         As JSRetBoolean
    Dim lNewPos         As Long

    If oCol Is Nothing Then
        Exit Sub
    End If
    If frColIsGrouped(oCol.Index) Then
        Exit Sub
    End If
    lNewPos = m_oGroups.Count + 1
    If Not oTargetGroup Is Nothing Then
        lNewPos = oTargetGroup.Index
        If bAfter Then
            lNewPos = lNewPos + 1
        End If
    End If
    Set oGroup = New JSGroup
    oGroup.frInit Me, lNewPos, oCol.Index, jgexSortAscending
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeGroupChange(oGroup, jgexGroupInsert, lNewPos, oCancel)
    If oCancel.Value Then
        Exit Sub
    End If
    m_oGroups.Add oCol.Index, jgexSortAscending, Index:=lNewPos
    RaiseEvent AfterGroupChange
End Sub

Private Sub pvMoveGroup(oGroup As JSGroup, oTargetGroup As JSGroup, ByVal bAfter As Boolean)
    Dim oCancel         As JSRetBoolean
    Dim lOldPos         As Long
    Dim lNewPos         As Long
    Dim eSortOrder      As jgexSortOrderConstants
    Dim lColIndex       As Long

    If oGroup Is Nothing Or oTargetGroup Is Nothing Then
        Exit Sub
    End If
    If oTargetGroup Is oGroup Then
        Exit Sub
    End If
    lOldPos = oGroup.Index
    lNewPos = pvDropPosition(lOldPos, oTargetGroup.Index, bAfter)
    If lNewPos = lOldPos Then
        Exit Sub
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeGroupChange(oGroup, jgexGroupMove, lNewPos, oCancel)
    If oCancel.Value Then
        Exit Sub
    End If
    lColIndex = oGroup.ColIndex
    eSortOrder = oGroup.SortOrder
    m_oGroups.Remove lOldPos
    m_oGroups.Add lColIndex, eSortOrder, Index:=lNewPos
    RaiseEvent AfterGroupChange
End Sub

Private Sub pvDeleteGroup(oGroup As JSGroup, oTargetCol As JSColumn, ByVal bAfter As Boolean)
    Dim oCancel         As JSRetBoolean
    Dim oCol            As JSColumn
    Dim lNewPos         As Long

    If oGroup Is Nothing Then
        Exit Sub
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeGroupChange(oGroup, jgexGroupDelete, oGroup.Index, oCancel)
    If oCancel.Value Then
        Exit Sub
    End If
    If oGroup.ColIndex >= 1 And oGroup.ColIndex <= m_oColumns.Count Then
        Set oCol = m_oColumns.Item(oGroup.ColIndex)
    End If
    m_oGroups.Remove oGroup.Index
    RaiseEvent AfterGroupChange
    If oTargetCol Is Nothing Or oCol Is Nothing Or oTargetCol Is oCol Then
        Exit Sub
    End If
    lNewPos = pvDropPosition(oCol.ColPosition, oTargetCol.ColPosition, bAfter)
    If lNewPos = oCol.ColPosition Then
        Exit Sub
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeColMove(oCol, lNewPos, oCancel)
    If Not oCancel.Value Then
        frColMove oCol, lNewPos
        RaiseEvent AfterColMove
    End If
End Sub

Private Sub pvEndColSize(ByVal bCancel As Boolean)
    Dim oCol            As JSColumn
    Dim oCancel         As JSRetBoolean
    Dim lWidth          As Long
    Dim lNewWidth       As Long

    Set oCol = m_oSizeCol
    Set m_oSizeCol = Nothing
    Call ReleaseCapture
    lWidth = pvColWidth(oCol)
    lNewWidth = oCol.Width
    oCol.frWidthPx = m_lSizeStartW
    If Not bCancel Then
        Set oCancel = New JSRetBoolean
        RaiseEvent ColResize(oCol.Index, lNewWidth, oCancel)
        bCancel = oCancel.Value
    End If
    If Not bCancel Then
        oCol.frWidthPx = lWidth
    End If
    pvInvalidate
End Sub

Private Function pvColAtX(ByVal lX As Long, oCol As JSColumn, Optional bAfter As Boolean) As Long
    Dim lCum            As Long
    Dim oItem           As JSColumn
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim lW              As Long

    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        lW = pvColWidth(oItem)
        lCum = lCum + lW
        If lX < lCum Then
            Set oCol = oItem
            bAfter = (lX >= lCum - lW \ 2)
            pvColAtX = pvVisiblePosition(vOrder(lIdx))
            Exit Function
        End If
    Next
End Function

Private Sub pvGroupDropTarget(ByVal lX As Long, oGroup As JSGroup, bAfter As Boolean)
    Dim oItem           As JSGroup

    pvLayoutGroupChips
    For Each oItem In m_oGroups
        With oItem.frChipRect
            If .Right > .Left Then
                If Not oGroup Is Nothing And lX < .Left Then
                    Exit For
                End If
                Set oGroup = oItem
            End If
        End With
    Next
    If Not oGroup Is Nothing Then
        bAfter = (lX >= (oGroup.frChipRect.Left + oGroup.frChipRect.Right) \ 2)
    End If
End Sub

Private Function pvGroupAtPoint(ByVal lX As Long, ByVal lY As Long, oGroup As JSGroup, Optional bAfter As Boolean) As Long
    Dim lIdx            As Long
    Dim oItem           As JSGroup

    pvLayoutGroupChips
    For Each oItem In m_oGroups
        lIdx = lIdx + 1
        With oItem.frChipRect
            If lX >= .Left And lX < .Right And lY >= .Top And lY < .Bottom And .Right > .Left Then
                Set oGroup = oItem
                bAfter = (lX >= (.Left + .Right) \ 2)
                pvGroupAtPoint = lIdx
                Exit For
            End If
        End With
    Next
End Function

Private Function pvEditInit(ByVal lPos As Long, ByVal lCol As Long, ByVal bSelectAll As Boolean, Optional ByVal lClickX As Long = -1, Optional ByVal lClickY As Long = -1) As Boolean
    Dim oCol            As JSColumn
    Dim oCancel         As JSRetBoolean
    Dim lRowIndex       As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lW              As Long
    Dim lStyle          As Long
    Dim lCaret          As Long
    Dim lTextTop        As Long
    Dim lEditH          As Long
    Dim pFont           As IFont
    Dim uMetrics        As TEXTMETRICW

    '--- a cell edits only where the control and the column both allow it:
    '--- EditType is jgexEditNone until a column asks for an editor
    If Not m_bAllowEdit Or m_bEditing Then
        Exit Function
    End If
    lRowIndex = m_pDataModel.RowIndex(lPos)
    If lRowIndex <= 0 Then
        Exit Function
    End If
    Set oCol = pvColByPosition(lCol)
    If oCol Is Nothing Then
        Exit Function
    End If
    If Not pvCellRect(lPos, oCol.Index, lX, lY, lW) Then
        Exit Function
    End If
    If lClickX >= 0 Then
        lClickX = lClickX - lX
        lClickY = lClickY - lY
    End If
    EnsureVisible lPos
    pvEnsureColVisible lPos, oCol.Index
    If Not pvCellRect(lPos, oCol.Index, lX, lY, lW) Then
        Exit Function
    End If
    If lClickX >= 0 Then
        lClickX = lClickX + lX
        lClickY = lClickY + lY
    End If
    '--- the client gets its veto before the editor appears
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeColEdit(oCol.Index, oCancel)
    If oCancel.Value Then
        Exit Function
    End If
    If oCol.EditType = jgexEditCheckBox Then
        Value(oCol.Index) = Not pvIsChecked(pvRowDataAt(lPos), oCol.Index)
        pvInvalidate SkipScroll:=True
        RaiseEvent Change
        pvEditInit = True
        Exit Function
    End If
    If oCol.EditType <> jgexEditTextBox Then
        m_bEditing = True
        m_lEditRow = lRowIndex
        m_lEditCol = oCol.Index
        m_sEditOldValue = pvCellText(pvRowDataAt(lPos), oCol.Index)
        pvEditInit = True
        Exit Function
    End If
    m_bInEditSetup = True
    m_bEditing = True
    m_lEditRow = lRowIndex
    m_lEditCol = oCol.Index
    m_sEditOldValue = pvCellText(pvRowDataAt(lPos), oCol.Index)
    lStyle = WS_CHILD Or WS_VISIBLE
    Select Case oCol.TextAlignment
    Case jgexAlignRight
        lStyle = lStyle Or ES_RIGHT
    Case jgexAlignCenter
        lStyle = lStyle Or ES_CENTER
    Case Else
        lStyle = lStyle Or ES_LEFT
    End Select
    If oCol.WordWrap Then
        '--- a wrapping editor scrolls down instead of across, and without
        '--- that it refuses what will not fit in the lines it has
        lStyle = lStyle Or ES_MULTILINE Or ES_AUTOVSCROLL
    Else
        lStyle = lStyle Or ES_AUTOHSCROLL
    End If
    uMetrics = FontTextMetrics(m_oFont)
    lTextTop = lY + (pvRowContentH(m_lRowHeight) - uMetrics.tmHeight) \ 2
    lEditH = uMetrics.tmHeight
    If oCol.WordWrap Then
        lTextTop = lY + 1
        lEditH = pvRowContentH(m_lRowHeight) - 2
    End If
    m_hWndEdit = CreateWindowEx(0, StrPtr("EDIT"), 0, lStyle, lX + 1, lTextTop, lW - 4, lEditH, hWnd, 0, App.hInstance, ByVal 0&)
    If m_hWndEdit = 0 Then
        m_bEditing = False
        m_bInEditSetup = False
        Exit Function
    End If
    Set pFont = m_oFont
    Call SendMessage(m_hWndEdit, WM_SETFONT, pFont.hFont, 1)
    Call SendMessage(m_hWndEdit, EM_SETMARGINS, EC_LEFTMARGIN, 1)
    Call SendMessage(m_hWndEdit, EM_LIMITTEXT, oCol.MaxLength, 0)
    Call SendMessage(m_hWndEdit, WM_SETTEXT, 0, ByVal StrPtr(m_sEditOldValue))
    If bSelectAll Then
        Call SendMessage(m_hWndEdit, EM_SETSEL, 0, -1)
    ElseIf lClickY >= 0 Then
        If lClickY - lTextTop >= uMetrics.tmHeight Then
            lCaret = Len(m_sEditOldValue)
        Else
            lCaret = SendMessage(m_hWndEdit, EM_CHARFROMPOS, 0, MakeDWord(lClickX - lX, lClickY - lTextTop))
            If lCaret = -1 Then
                '--- the point fell outside the editor, which answers with the end
                lCaret = Len(m_sEditOldValue)
            Else
                lCaret = lCaret And &HFFFF&
            End If
        End If
        Call SendMessage(m_hWndEdit, EM_SETSEL, lCaret, lCaret)
    Else
        Call SendMessage(m_hWndEdit, EM_SETSEL, Len(m_sEditOldValue), Len(m_sEditOldValue))
    End If
    Set m_pSubclassEdit = InitSubclassingThunk(m_hWndEdit, Me, pvAddressOfSubclassProc.EditSubclassProc(0, 0, 0, 0, 0))
    Call SetFocusApi(m_hWndEdit)
    m_bInEditSetup = False
    pvEditInit = True
End Function

Private Sub pvEditFocusLost()
    Dim sText           As String

    If Not m_bEditing Then
        Exit Sub
    End If
    sText = pvEditText()
    m_bEditing = False
    pvEditTerminate
    If sText <> m_sEditOldValue Then
        If Not m_oRowData Is Nothing Then
            m_oRowData.frAllowUpdate = True
            m_oRowData.Value(m_lEditCol) = sText
        End If
    End If
    pvInvalidate SkipScroll:=True
End Sub

Private Function pvEditText() As String
    Dim lLen            As Long

    If m_hWndEdit = 0 Then
        '--- a windowless session had nothing typed into it
        pvEditText = m_sEditOldValue
        Exit Function
    End If
    lLen = SendMessage(m_hWndEdit, WM_GETTEXTLENGTH, 0, 0)
    pvEditText = String$(lLen + 1, 0)
    lLen = SendMessage(m_hWndEdit, WM_GETTEXT, lLen + 1, ByVal StrPtr(pvEditText))
    pvEditText = Left$(pvEditText, lLen)
End Function

Private Sub pvEditTerminate()
    If m_hWndEdit = 0 Then
        Exit Sub
    End If
    TerminateSubclassingThunk m_pSubclassEdit, Me
    Set m_pSubclassEdit = Nothing
    Call DestroyWindow(m_hWndEdit)
    m_hWndEdit = 0
End Sub

'--- the row half of leaving a cell, and separate from the cell's own half on
'--- purpose: whether an editor was open, and whether what was typed in it
'--- changed anything, says nothing about the row -- a column edited earlier
'--- can be pending behind an editor closed on the same text
Private Sub pvEditCommit()
    Dim oCancel         As JSRetBoolean

    '--- what the buffer holds goes to storage and the client hears about the
    '--- row rather than the cell.
    '--- Re-entrant only through the client: a handler of the events
    '--- UpdateRowData raises can move the row, which comes back here with the
    '--- write half-done
    If Not DataChanged Or m_bInPendCommit Then
        Exit Sub
    End If
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeUpdate(oCancel)
    If oCancel.Value Then
        Exit Sub
    End If
    '--- the buffer is the row, so the model writes back from it directly and
    '--- the page is already showing what was written: repopulating here would
    '--- re-raise RowFormat for every visible row, where the original raises it
    '--- once, for this one
    m_bInPendCommit = True
    m_pDataModel.UpdateRowData m_oRowData
    m_bInPendCommit = False
    '--- the writes are storage's now, so the wrapper keeps them and the
    '--- decoration RowFormat gave it. Dropping the permission to write is what
    '--- drops the marks with it, and it is read-only again the moment it stops
    '--- being the pending row: the page hands that same wrapper to the client
    m_oRowData.frAllowUpdate = False
    '--- the row that was just written is handed back to be decorated before
    '--- the client is told the update is done -- recorded from the original in
    '--- 058, 068 and 071
    RaiseEvent RowFormat(m_oRowData)
    RaiseEvent AfterUpdate
End Sub

Private Sub pvEditCancel()
    If Not DataChanged Then
        Exit Sub
    End If
    '--- the pending row is the current row: the wrapper is taken for m_lRow
    '--- wherever the row moves, so its position is the one already in hand
    Set m_oRowData = m_pDataModel.GetRowData(m_lRow)
    If m_lRow >= m_lWindowFirst And m_lRow < m_lWindowFirst + m_lWindowCount Then
        Set m_aWindow(m_lRow - m_lWindowFirst + 1) = m_oRowData
    End If
    pvInvalidate SkipScroll:=True
    RaiseEvent RowFormat(m_oRowData)
End Sub

Private Sub pvEditEnd(Optional ByVal bCancel As Boolean)
    Dim oCancel         As JSRetBoolean
    Dim lCol            As Long
    Dim lRowIndex       As Long
    Dim sText           As String
    Dim oRowData        As JSRowData

    If Not m_bEditing Then
        Exit Sub
    End If
    lCol = m_lEditCol
    lRowIndex = m_lEditRow
    sText = pvEditText()
    m_bEditing = False
    pvEditTerminate
    Set oRowData = m_oRowData
    '--- committing runs the update trio the original raises in this order:
    '--- the cell first, then the row, with a repaint between the two halves
    If Not bCancel And sText <> m_sEditOldValue Then
        Set oCancel = New JSRetBoolean
        RaiseEvent BeforeColUpdate(lRowIndex, lCol, m_sEditOldValue, oCancel)
        If oCancel.Value Then
            GoTo QH
        End If
        If Not oRowData Is Nothing Then
            oRowData.frAllowUpdate = True
            oRowData.Value(lCol) = sText
        End If
        RaiseEvent AfterColUpdate(lCol)
    Else
        '--- a cancelled edit says only that the session ended. It repaints the
        '--- cell it was covering when that cancel is the whole of the row's
        '--- edit: with other columns still buffered the row stays dirty and
        '--- keeps its RowFormat until whatever resolves it
        If sText <> m_sEditOldValue And Not DataChanged Then
            RaiseEvent RowFormat(m_oRowData)
        End If
    End If
QH:
    pvInvalidate SkipScroll:=True
    RaiseEvent AfterColEdit(lCol)
    '--- the editor window is gone by now, so the grid takes the focus back
    '--- through the API rather than through VB, which raises when the control
    '--- is not in a state to take it
    Call SetFocusApi(hWnd)
End Sub

Private Function pvGroupByBoxHeight() As Long
    If m_bGroupByBoxVisible Then
        pvGroupByBoxHeight = m_lColumnHeaderHeight + 14
        If m_oGroups.Count > 1 Then
            pvGroupByBoxHeight = CLng(m_lColumnHeaderHeight * (m_oGroups.Count + 1) / 2) + 14
        End If
    End If
End Function

Private Function pvRowsTop() As Long
    pvRowsTop = pvGroupByBoxHeight()
    If m_bColumnHeaders Then
        pvRowsTop = pvRowsTop + m_lColumnHeaderHeight
    End If
End Function

Private Function pvColByPosition(ByVal lPos As Long) As JSColumn
    Dim lIdx            As Long
    Dim lSeen           As Long
    Dim oItem           As JSColumn

    For lIdx = 1 To m_oColumns.Count
        Set oItem = m_oColumns.ItemByPosition(lIdx)
        If oItem.Visible Then
            lSeen = lSeen + 1
            If lSeen = lPos Then
                Set pvColByPosition = oItem
                Exit Function
            End If
        End If
    Next
End Function

Private Function pvCellRect(ByVal lPos As Long, ByVal lColIndex As Long, lX As Long, lY As Long, lW As Long) As Boolean
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim oItem           As JSColumn
    Dim lCum            As Long

    '--- where the cell sits in the control, which is where the editor goes.
    '--- Keyed by the column rather than by its place among the visible ones,
    '--- since scrolling moves that place while the cell stays the same cell
    If m_lRowHeight <= 0 Or lPos < m_lFirstItem Then
        Exit Function
    End If
    lY = pvRowsTop() + (lPos - m_lFirstItem) * m_lRowHeight
    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oItem.Visible Then
            If oItem.Index = lColIndex Then
                lX = lCum
                lW = pvColWidth(oItem)
                pvCellRect = True
                Exit Function
            End If
            lCum = lCum + pvColWidth(oItem)
        End If
    Next
End Function

'--- the editor follows its cell: the grid scrolling under it moves it, a
'--- column half off the right edge clips it, and a cell scrolled out of the
'--- view altogether takes it out of sight rather than leaving it floating
Private Function pvEditorRect(lX As Long, lTop As Long, lW As Long, lH As Long) As Boolean
    Dim lPos            As Long
    Dim lY              As Long
    Dim lCellW          As Long
    Dim oCol            As JSColumn
    Dim uMetrics        As TEXTMETRICW

    '--- a selection of more than one row is a different mode: the box goes
    '--- away rather than sitting on whichever of them holds the current cell
    If m_oSelectedItems.Count > 1 Then
        Exit Function
    End If
    lPos = m_pDataModel.GetRowPosition(m_lEditRow)
    If lPos < 1 Then
        Exit Function
    End If
    If Not pvCellRect(lPos, m_lEditCol, lX, lY, lCellW) Then
        Exit Function
    End If
    Set oCol = m_oColumns.Item(m_lEditCol)
    uMetrics = FontTextMetrics(m_oFont)
    lTop = lY + (pvRowContentH(m_lRowHeight) - uMetrics.tmHeight) \ 2
    lH = uMetrics.tmHeight
    If oCol.WordWrap Then
        lTop = lY + 1
        lH = pvRowContentH(m_lRowHeight) - 2
    End If
    lX = lX + 1
    lW = lCellW - 4
    '--- what is left of the cell inside the grid surface: a column half off
    '--- the right edge leaves a narrower box, and a part-row at the bottom
    '--- keeps its own, clipped by the window it sits in. Only a cell with
    '--- nothing left inside the surface at all goes away
    If lW > picGrid.ScaleWidth - lX Then
        lW = picGrid.ScaleWidth - lX
    End If
    pvEditorRect = (lW > 0 And lTop < picGrid.ScaleHeight)
End Function

'--- a column half off the right edge is walked in one at a time; one that
'--- scrolled off the left is not in the order at all, so the view walks back
'--- to it. Both loops are bounded by the column count, since LeftCol clamps
'--- and would otherwise leave them turning
Private Sub pvEnsureColVisible(ByVal lPos As Long, ByVal lColIndex As Long)
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lW              As Long

    For lIdx = 1 To m_oColumns.Count
        If pvCellRect(lPos, lColIndex, lX, lY, lW) Or m_lLeftCol <= 1 Then
            Exit For
        End If
        LeftCol = m_lLeftCol - 1
    Next
    For lIdx = 1 To m_oColumns.Count
        If Not pvCellRect(lPos, lColIndex, lX, lY, lW) Then
            Exit For
        End If
        If lX + lW <= picGrid.ScaleWidth Or lX <= pvBlockLeft() Then
            Exit For
        End If
        LeftCol = m_lLeftCol + 1
    Next
End Sub

Private Sub pvLayoutEditor()
    Dim lX              As Long
    Dim lTop            As Long
    Dim lW              As Long
    Dim lH              As Long

    If m_hWndEdit = 0 Then
        Exit Sub
    End If
    If pvEditorRect(lX, lTop, lW, lH) Then
        Call SetWindowPos(m_hWndEdit, 0, lX, lTop, lW, lH, SWP_NOZORDER Or SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(m_hWndEdit, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_HIDEWINDOW)
    End If
End Sub

Private Sub pvAutoSort(oCol As JSColumn)
    Dim oGroup          As JSGroup

    '--- what the original documents client code used to write in the two
    '--- events: a grouped column flips the order of its group, any other one
    '--- becomes the only sort key, ascending unless it already was
    For Each oGroup In m_oGroups
        If oGroup.ColIndex = oCol.Index Then
            If oGroup.SortOrder = jgexSortAscending Then
                oGroup.SortOrder = jgexSortDescending
            Else
                oGroup.SortOrder = jgexSortAscending
            End If
            Exit Sub
        End If
    Next
    If frColSortOrder(oCol.Index) = jgexSortAscending Then
        m_oSortKeys.Clear
        m_oSortKeys.Add oCol.Index, jgexSortDescending
    Else
        m_oSortKeys.Clear
        m_oSortKeys.Add oCol.Index, jgexSortAscending
    End If
End Sub

Private Sub pvOnLButtonDown(ByVal lX As Long, ByVal lY As Long, ByVal lShift As Long)
    Dim lTopGbox        As Long
    Dim lTopHdr         As Long
    Dim lPos            As Long
    Dim oCol            As JSColumn
    Dim oGroup          As JSGroup
    Dim lRow            As Long

    m_bDragSelect = False
    lTopGbox = pvGroupByBoxHeight()
    lTopHdr = pvRowsTop()
    '--- the divider takes the click before the header under it does, or every
    '--- resize would sort the column it started on
    Set m_oSizeCol = pvColDividerAt(lX, lY)
    If Not m_oSizeCol Is Nothing Then
        m_lSizeStartX = lX
        m_lSizeStartW = pvColWidth(m_oSizeCol)
        Call SetCapture(picGrid.hWnd)
        Exit Sub
    End If
    If lY < lTopGbox Then
        '--- group-by box: a chip stands for its column, so clicking one sorts
        '--- exactly as clicking that column's header does
        lPos = pvGroupAtPoint(lX, lY, oGroup)
        If Not oGroup Is Nothing Then
            '--- the press only remembers the chip, as a press on a header only
            '--- remembers the header: what it turns out to be is decided when
            '--- the button comes up, since a chip dragged out of its place is a
            '--- move rather than a click that sorts
            Set m_oDragGroup = oGroup
            m_lDragStartX = lX
            m_lDragStartY = lY
            pvInvalidateHeaders
        End If
    ElseIf lY < lTopHdr Then
        '--- column header band
        '--- the press only remembers the header: what it turns out to be is
        '--- decided when the button comes up, since a header dragged onto
        '--- another column is a move rather than a click that sorts
        lPos = pvColAtX(lX, oCol)
        If Not oCol Is Nothing Then
            Set m_oDragCol = oCol
            m_lDragStartX = lX
            m_lDragStartY = lY
            pvInvalidateHeaders
        End If
    ElseIf m_lRowHeight > 0 Then
        '--- data area cell
        lRow = m_lFirstItem + (lY - lTopHdr) \ m_lRowHeight
        If lRow >= 1 And lRow <= RowCount Then
            m_bDragSelect = True
            If pvIsGroupRow(lRow) Then
                pvNavigate lRow, 0, lShift, (lShift And vbCtrlMask) <> 0
                pvGroupBoxToggle lRow, lX, lY
            ElseIf pvRowSelZoneAt(lX, lY) Then
                pvNavigate lRow, 0, lShift, (lShift And vbCtrlMask) <> 0
                pvClearCol
            Else
                lPos = pvColAtX(lX, oCol)
                If lPos >= 1 Then
                    If Not m_bAllowEdit Or Not oCol.Selectable Then
                        pvNavigate lRow, 0, lShift, (lShift And vbCtrlMask) <> 0
                    Else
                        pvNavigate lRow, lPos, lShift, (lShift And vbCtrlMask) <> 0
                        If (lShift And (vbCtrlMask Or vbShiftMask)) = 0 Then
                            m_bClickOpenedEdit = pvEditInit(lRow, lPos, False, lX, lY)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub pvToggleRow(ByVal lRow As Long)
    RowExpanded(lRow) = Not m_pDataModel.RowExpanded(lRow)
    pvSyncProjection
    pvPopulateWindow
End Sub

Private Function pvGroupBoxRect(ByVal lPos As Long, ByVal lRowTop As Long, uRect As RECT) As Boolean
    If Not pvIsGroupRow(lPos) Then
        Exit Function
    End If
    If pvWindowRow(lPos) Is Nothing Then
        Exit Function
    End If
    uRect.Left = pvRowHeaderWidth() + pvRowIndent(lPos) + (GROUP_INDENT_W - GROUP_BOX_W) \ 2
    uRect.Top = lRowTop + (m_lRowHeight - 1 - GROUP_BOX_W) \ 2
    uRect.Right = uRect.Left + GROUP_BOX_W
    uRect.Bottom = uRect.Top + GROUP_BOX_W
    pvGroupBoxRect = True
End Function

Private Function pvGroupBoxToggle(ByVal lRow As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
    Dim oRowData        As JSRowData
    Dim uRect           As RECT

    pvPopulateWindow
    '--- only a header carries the box a footer's indent merely leaves room for
    Set oRowData = pvWindowRow(lRow)
    If oRowData Is Nothing Then
        Exit Function
    End If
    If oRowData.RowType <> jgexRowTypeGroupHeader Then
        Exit Function
    End If
    If Not pvGroupBoxRect(lRow, pvRowsTop() + (lRow - m_lFirstItem) * m_lRowHeight, uRect) Then
        Exit Function
    End If
    If Not pvPtInRect(uRect, lX, lY) Then
        Exit Function
    End If
    pvToggleRow lRow
    pvGroupBoxToggle = True
End Function

Private Function pvGroupRowDblClk(ByVal lY As Long) As Boolean
    Dim lRow            As Long
    Dim oRowData        As JSRowData

    If m_lRowHeight <= 0 Or lY < pvRowsTop() Then
        Exit Function
    End If
    lRow = m_lFirstItem + (lY - pvRowsTop()) \ m_lRowHeight
    If lRow < 1 Or lRow > RowCount Then
        Exit Function
    End If
    pvPopulateWindow
    Set oRowData = pvWindowRow(lRow)
    If oRowData Is Nothing Then
        Exit Function
    End If
    If oRowData.RowType <> jgexRowTypeGroupHeader Then
        Exit Function
    End If
    pvToggleRow lRow
    pvGroupRowDblClk = True
End Function

Private Sub pvOnKeyDown(ByVal lKeyCode As Long, ByVal lShift As Long)
    Dim oRowData        As JSRowData
    Dim lPos            As Long
    Dim oCol            As JSColumn

    Select Case lKeyCode
    Case vbKeyDown, vbKeyReturn
        If m_lRow < RowCount Then
            pvNavigate m_lRow + 1, m_lCol, lShift, False
        End If
    Case vbKeyUp
        If m_lRow > 1 Then
            pvNavigate m_lRow - 1, m_lCol, lShift, False
        End If
    Case vbKeyRight
        '--- probed on a group header: Right expands a collapsed level and
        '--- steps into the first child of an expanded one
        If pvIsGroupRow(m_lRow) Then
            If pvIsGroupHeader(m_lRow) Then
                If Not m_pDataModel.RowExpanded(m_lRow) Then
                    pvToggleRow m_lRow
                ElseIf m_lRow < RowCount Then
                    pvNavigate m_lRow + 1, m_lCol, lShift, False
                End If
            End If
        Else
            lPos = pvNextSelectableCol(m_lCol, 1)
            If lPos >= 1 Then
                Col = lPos
                EnsureVisible m_lRow, m_lCol
            End If
        End If
    Case vbKeyLeft
        '--- probed: Left collapses an expanded header quietly -- a collapse
        '--- exposes nothing, so no RowFormat -- and stays put on a collapsed one
        If pvIsGroupRow(m_lRow) Then
            If pvIsGroupHeader(m_lRow) Then
                If m_pDataModel.RowExpanded(m_lRow) Then
                    pvToggleRow m_lRow
                End If
            End If
        Else
            lPos = pvNextSelectableCol(m_lCol, -1)
            If lPos >= 1 Then
                Col = lPos
                EnsureVisible m_lRow, m_lCol
            End If
        End If
    Case vbKeyF2
        pvEditInit m_lRow, m_lCol, bSelectAll:=True
    Case vbKeyDelete
        If m_lCol >= 1 And m_bAllowEdit Then
            Set oCol = pvColByPosition(m_lCol)
            If Not oCol Is Nothing Then
                If oCol.EditType = jgexEditTextBox Then
                    If pvEditInit(m_lRow, m_lCol, bSelectAll:=True) Then
                        If m_hWndEdit <> 0 Then
                            Call SendMessage(m_hWndEdit, WM_CLEAR, 0, 0)
                        End If
                    End If
                End If
            End If
        ElseIf m_bAllowDelete Then
            Delete
        End If
    Case vbKeySpace
        '--- a deliberate step past the original, which ignores Space here:
        '--- the header's level toggles, a footer stays as it is
        If pvIsGroupHeader(m_lRow) Then
            pvToggleRow m_lRow
        End If
    Case vbKeyTab
        If m_bAllowEdit And (lShift And vbCtrlMask) = 0 Then
            If (lShift And vbShiftMask) <> 0 Then
                lPos = pvNextSelectableCol(m_lCol, -1)
            Else
                lPos = pvNextSelectableCol(m_lCol, 1)
            End If
            If lPos >= 1 Then
                Col = lPos
                EnsureVisible m_lRow, m_lCol
                pvEditInit m_lRow, m_lCol, bSelectAll:=True
            End If
        End If
    Case vbKeyPageDown
        pvNavigate pvClampRow(m_lRow + pvVisibleRows()), m_lCol, lShift, False
    Case vbKeyPageUp
        pvNavigate pvClampRow(m_lRow - pvVisibleRows()), m_lCol, lShift, False
    Case vbKeyHome
        pvNavigate pvClampRow(1), m_lCol, lShift, False
    Case vbKeyEnd
        pvNavigate pvClampRow(RowCount), m_lCol, lShift, False
    Case vbKeyEscape
        If Not m_oSizeCol Is Nothing Then
            pvEndColSize bCancel:=True
        ElseIf Not m_oDragCol Is Nothing Or Not m_oDragGroup Is Nothing Then
            pvEndColDrag bCancel:=True
        Else
            pvEditCancel
        End If
    End Select
End Sub

Private Function pvNextSelectableCol(ByVal lFrom As Long, ByVal lStep As Long) As Long
    Dim lPos            As Long
    Dim oCol            As JSColumn

    lPos = lFrom + lStep
    Do While lPos >= 1 And lPos <= pvVisibleColCount()
        Set oCol = pvColByPosition(lPos)
        If Not oCol Is Nothing Then
            If oCol.Selectable Then
                pvNextSelectableCol = lPos
                Exit Function
            End If
        End If
        lPos = lPos + lStep
    Loop
End Function

Private Function pvEditKeyLeaves(ByVal nKeyCode As Integer) As Boolean
    Dim lStart          As Long
    Dim lEnd            As Long

    Select Case nKeyCode
    Case vbKeyTab
        pvEditKeyLeaves = True
    Case vbKeyDown, vbKeyUp
        pvEditKeyLeaves = ((GetWindowLong(m_hWndEdit, GWL_STYLE) And ES_MULTILINE) = 0)
    Case vbKeyLeft, vbKeyRight
        Call SendMessage(m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd))
        If lStart = lEnd Then
            If nKeyCode = vbKeyRight Then
                pvEditKeyLeaves = (lEnd = GetWindowTextLength(m_hWndEdit))
            ElseIf (GetWindowLong(m_hWndEdit, GWL_STYLE) And ES_MULTILINE) = 0 Then
                pvEditKeyLeaves = (lStart = 0)
            End If
        End If
    End Select
End Function

Private Sub pvSetRow(ByVal lValue As Long)
    Dim lLastRow        As Long

    '--- moves the current row without touching the selection: the row is
    '--- painted selected, so it still has to repaint
    If m_lRow <> lValue Then
        lLastRow = m_lRow
        m_lRow = lValue
        pvSyncRowData
        If m_pDataModel.RowIndex(m_lRow) <> 0 Then
            m_lHoldRowIndex = m_pDataModel.RowIndex(m_lRow)
        End If
        pvInvalidate SkipScroll:=True
        RaiseEvent RowColChange(lLastRow, m_lCol)
    End If
End Sub

'--- moves the current cell and updates the selection accordingly
Private Sub pvNavigate(ByVal lRow As Long, ByVal lCol As Long, ByVal lShift As Long, ByVal bCtrlToggle As Boolean)
    Dim bRowChanging    As Boolean

    '--- an open editor commits before any of it, ahead of SelectionChange
    bRowChanging = (lRow <> m_lRow And lRow >= 1 And lRow <= RowCount)
    pvEditEnd
    If bRowChanging Then
        pvEditCommit
    End If
    If bRowChanging Then
        If Not pvIsGroupRow(lRow) Then
            RaiseEvent RowFormat(pvRowDataAt(lRow))
        End If
    End If
    If lRow >= 1 And lRow <= RowCount Then
        pvUpdateSelection lRow, lShift, bCtrlToggle
        pvSetRow lRow
    Else
        pvUpdateSelection m_lRow, lShift, bCtrlToggle
    End If
    If lCol >= 1 Then
        Col = lCol
    End If
    EnsureVisible m_lRow
End Sub

Private Sub pvUpdateSelection(ByVal lRow As Long, ByVal lShift As Long, ByVal bCtrlToggle As Boolean)
    Dim bSame           As Boolean

    If lRow < 1 Or lRow > RowCount Then
        Exit Sub
    End If
    If m_bMultiSelect And bCtrlToggle Then
        If pvIsRowSelected(lRow) Then
            m_oSelectedItems.RemoveRowPosition lRow
        Else
            pvAddSel lRow
        End If
        m_bCurRowDeselected = Not pvIsRowSelected(lRow)
        m_lSelAnchor = lRow
        pvInvalidate SkipScroll:=True
        RaiseEvent SelectionChange
    ElseIf m_bMultiSelect And (lShift And vbShiftMask) <> 0 Then
        pvSetRangeSel m_lSelAnchor, lRow
        m_bCurRowDeselected = False
    Else
        bSame = (m_oSelectedItems.Count = 1)
        If bSame Then
            bSame = (m_oSelectedItems.Item(1).RowPosition = lRow)
        End If
        m_oSelectedItems.Clear
        pvAddSel lRow
        m_lSelAnchor = lRow
        m_bCurRowDeselected = False
        pvInvalidate SkipScroll:=True
        If Not bSame Then
            RaiseEvent SelectionChange
        End If
    End If
End Sub

Private Function pvIsRowSelected(ByVal lPos As Long) As Boolean
    Dim oItem           As JSSelectedItem

    For Each oItem In m_oSelectedItems
        If oItem.RowPosition = lPos Then
            pvIsRowSelected = True
            Exit Function
        End If
    Next
End Function

Private Sub pvAddSel(ByVal lPos As Long)
    Dim lRow            As Long

    If lPos < 1 Or lPos > m_pDataModel.RowCount Then
        Exit Sub
    End If
    lRow = m_pDataModel.RowIndex(lPos)
    m_oSelectedItems.frAdd lPos, m_pDataModel.RowBookmark(lRow), lRow
End Sub

Private Sub pvSetRangeSel(ByVal lFrom As Long, ByVal lTo As Long)
    Dim lLo             As Long
    Dim lHi             As Long
    Dim lIdx            As Long

    If lFrom = 0 Then
        lFrom = lTo
    End If
    If lFrom <= lTo Then
        lLo = lFrom
        lHi = lTo
    Else
        lLo = lTo
        lHi = lFrom
    End If
    m_oSelectedItems.Clear
    For lIdx = lLo To lHi
        pvAddSel lIdx
    Next
    pvInvalidate SkipScroll:=True
    RaiseEvent SelectionChange
End Sub

Private Sub pvOnVScroll(ByVal lCode As Long, ByVal lPos As Long)
    Dim uSi             As SCROLLINFO
    Dim lPage           As Long

    Select Case lCode
    Case SB_LINEUP
        FirstItem = m_lFirstItem - 1
    Case SB_LINEDOWN
        FirstItem = m_lFirstItem + 1
    Case SB_PAGEUP, SB_PAGEDOWN
        uSi.cbSize = Len(uSi)
        uSi.fMask = SIF_PAGE
        Call GetScrollInfo(picGrid.hWnd, SB_VERT, uSi)
        lPage = uSi.nPage
        If lPage < 1 Then
            lPage = 1
        End If
        If lCode = SB_PAGEUP Then
            FirstItem = m_lFirstItem - lPage
        Else
            FirstItem = m_lFirstItem + lPage
        End If
    Case SB_THUMBPOSITION, SB_THUMBTRACK
        If lCode = SB_THUMBTRACK And Not m_bContinuousScroll Then
            Exit Sub
        End If
        If lPos = 0 Then
            uSi.cbSize = Len(uSi)
            uSi.fMask = SIF_TRACKPOS
            Call GetScrollInfo(picGrid.hWnd, SB_VERT, uSi)
            lPos = uSi.nTrackPos
        End If
        FirstItem = lPos + 1
    End Select
End Sub

Private Sub pvOnHScroll(ByVal lCode As Long, ByVal lPos As Long)
    Dim uSi             As SCROLLINFO

    If m_bScrollUpdating Then
        Exit Sub
    End If
    Select Case lCode
    Case SB_LINELEFT
        LeftCol = m_lLeftCol - 1
    Case SB_LINERIGHT
        LeftCol = m_lLeftCol + 1
    Case SB_PAGELEFT
        LeftCol = m_lLeftCol - pvColsFitBefore(pvScrollableWidth(), m_lLeftCol)
    Case SB_PAGERIGHT
        LeftCol = m_lLeftCol + pvColsFitFrom(pvScrollableWidth(), m_lLeftCol)
    Case SB_THUMBPOSITION, SB_THUMBTRACK
        If lCode = SB_THUMBTRACK And Not m_bContinuousScroll Then
            Exit Sub
        End If
        If lPos = 0 Then
            uSi.cbSize = Len(uSi)
            uSi.fMask = SIF_TRACKPOS
            Call GetScrollInfo(hsbGrid.hWnd, SB_CTL, uSi)
            lPos = uSi.nTrackPos
        End If
        LeftCol = Clamp(lPos, 1, m_oColumns.Count)
    Case SB_LEFT
        LeftCol = 1
    Case SB_RIGHT
        LeftCol = m_oColumns.Count
    End Select
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
' Control events
'=========================================================================

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"

    '--- a freshly placed control starts with two default empty columns
    On Error GoTo EH
    pvInheritAmbientFont
    m_oColumns.Add(vbNullString).Width = ToTwips(m_lDefaultColumnWidth)
    m_oColumns.Add(vbNullString).Width = ToTwips(m_lDefaultColumnWidth)
    pvSubclass
    If Ambient.UserMode Then
        InitIPAO m_uIPAO, Me
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"

    On Error GoTo EH
    pvInheritAmbientFont
    pvSubclass
    If Ambient.UserMode Then
        InitIPAO m_uIPAO, Me
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_EnterFocus()
    Const FUNC_NAME     As String = "UserControl_EnterFocus"
    
    On Error GoTo EH
    SetIPAO m_uIPAO
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_Resize()
    Const FUNC_NAME     As String = "UserControl_Resize"

    On Error GoTo EH
    pvLayoutGrid
    pvInvalidate
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_Hide()
    Const FUNC_NAME     As String = "UserControl_Hide"
    Dim bInIde          As Boolean: Debug.Assert SetTrue(bInIde)

    On Error GoTo EH
    If bInIde Then
        If Ambient.UserMode And EventsFrozen Then
            UserControl_Terminate
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Const FUNC_NAME     As String = "m_oFont_FontChanged"

    '--- default row height follows the data font unless explicitly set
    On Error GoTo EH
    If Not m_bRowHeightSet Then
        m_lRowHeight = FontTextMetrics(m_oFont).tmHeight + 3
        If m_lRowHeight < 19 Then
            m_lRowHeight = 19
        End If
    End If
    pvInvalidate
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oColumnHeaderFont_FontChanged(ByVal PropertyName As String)
    Const FUNC_NAME     As String = "m_oColumnHeaderFont_FontChanged"

    '--- header height always follows the header font
    On Error GoTo EH
    m_lColumnHeaderHeight = FontTextMetrics(m_oColumnHeaderFont).tmHeight + 6
    pvInvalidate
    Exit Sub
EH:
    PrintError FUNC_NAME
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
    m_oSortKeys.frInit Me
    m_oGroups.frInit Me
    m_oColumns.frInit Me
    Set m_oSelectedItems = New JSSelectedItems
    m_oSelectedItems.frInit Me
    Set m_oFormatStyles = New JSFormatStyles
    Set m_oPrinterProperties = New JSPrinterProperties
    Set m_oFont = NewStdFont()
    Set m_oColumnHeaderFont = NewStdFont()
    m_clrBackColor = vbWindowBackground
    m_clrForeColor = vbWindowText
    m_clrBackColorHeader = vbButtonFace
    m_clrForeColorHeader = vbButtonText
    m_clrBackColorBkg = vbWindowBackground
    m_clrBackColorGBBox = vb3DShadow
    m_clrBackColorInfoText = vbButtonFace
    m_clrForeColorInfoText = vb3DShadow
    m_clrBackColorRowGroup = vbButtonFace
    m_clrForeColorRowGroup = vbButtonText
    m_clrRowColorEven = &HC1D7B0
    m_clrRowColorOdd = &HBFFFFF
    m_clrGridLinesColor = vb3DShadow
    m_clrMaskColor = &HC0C0C0
    m_eGridLines = jgexGLBoth
    m_eGridLineStyle = jgexGLSSolid
    m_eHeaderStyle = jgexHSDouble3D
    m_eView = jgexTable
    m_eSelectionStyle = jgexEntireRow
    m_eHideSelection = jgexHideSelection
    m_eNewRowPos = jgexTop
    m_eTabKeyBehavior = jgexColumnNavigation
    m_eDefaultGroupMode = jgexDGMExpanded
    m_eGroupFooterStyle = jgexNoGroupFooter
    m_eDataMode = jgexDAO
    m_eRecordsetType = jgexRSDAODynaset
    m_eLockType = jgexLockOptimistic
    m_eCursorLocation = jgexUseServer
    m_eBorderStyle = jgexFixed
    m_lDefaultColumnWidth = 100
    m_lLeftCol = 1
    m_oColumnHeaderFont_FontChanged vbNullString
    m_oFont_FontChanged vbNullString
    m_lImageWidth = 16
    m_lImageHeight = 16
    m_lCardWidth = 250
    m_lCardSpacing = 12
    m_lPreviewRowIndent = 600
    m_bColumnHeaders = True
    m_bGroupByBoxVisible = True
    m_bAllowEdit = True
    m_bAllowColumnDrag = True
    m_bAllowCardSizing = True
    m_bCardBorders = True
    m_bShowEmptyFields = True
    m_bAutomaticArrange = True
    m_bRedraw = True
    m_bGridFocus = True
    m_sCalendarTodayText = "Today"
    m_sCalendarNoneText = "None"
    m_sGroupByBoxInfoText = "Drag a column header here to group by that column."
    m_sRecordNavigatorString = "Record:|of"
    pvCreateDataModel
End Sub

Private Sub UserControl_Terminate()
    Dim vErr            As Variant
    
    vErr = PushError
    pvEditTerminate
    If m_hCurRowSel <> 0 Then
        Call DestroyIcon(m_hCurRowSel)
        m_hCurRowSel = 0
    End If
    TerminateIPAO m_uIPAO
    TerminateSubclassingThunk m_pSubclassPic, Me
    TerminateSubclassingThunk m_pSubclassCtl, Me
    pvBufferTermiante
    '--- cleanup weakrefs
    If Not m_oGroups Is Nothing Then
        m_oGroups.frTerminate
        Set m_oGroups = Nothing
    End If
    If Not m_oSortKeys Is Nothing Then
        m_oSortKeys.frTerminate
        Set m_oSortKeys = Nothing
    End If
    If Not m_oColumns Is Nothing Then
        m_oColumns.frTerminate
        Set m_oColumns = Nothing
    End If
    If Not m_oSelectedItems Is Nothing Then
        m_oSelectedItems.frTerminate
        Set m_oSelectedItems = Nothing
    End If
    If Not m_pDataModel Is Nothing Then
        m_pDataModel.Terminate
        Set m_pDataModel = Nothing
    End If
    PopError vErr
End Sub
