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
Private m_pSubclassEdit             As IUnknown

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
End Property

Public Property Get RowHeight() As Long
    RowHeight = ToTwips(m_lRowHeight)
End Property

Public Property Let RowHeight(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original; an
    '--- explicit height survives later font changes
    m_lRowHeight = ToPixels(lValue)
    m_bRowHeightSet = True
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

    '--- setting a value past the visible columns selects the last one
    lMax = pvVisibleColCount()
    If nValue > lMax Then
        nValue = lMax
    End If
    If nValue < 1 And lMax > 0 Then
        nValue = 1
    End If
    If m_lCol <> nValue Then
        pvEndEdit
        lLastCol = m_lCol
        m_lCol = nValue
        '--- the current cell is drawn out of the selected row it sits in, so
        '--- moving between columns is a repaint and not only an event
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
    ColumnHeaderHeight = ToTwips(m_lColumnHeaderHeight)
End Property

Public Property Let ColumnHeaderHeight(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original
    m_lColumnHeaderHeight = ToPixels(lValue)
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
    DefaultColumnWidth = ToTwips(m_lDefaultColumnWidth)
End Property

Public Property Let DefaultColumnWidth(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original
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
    Dim lMax            As Long

    '--- scrolling stops once the last column reaches the right edge, so a
    '--- value past that clamps -- the frozen block is not scrolled over and
    '--- takes its width out of the strip the rest has to fill
    lMax = pvVisibleColCount() - pvColsFitFrom(pvScrollableWidth(), pvFrozenCount() + 1) + 1
    If lMax < 1 Then
        lMax = 1
    End If
    nValue = Clamp(nValue, 1, lMax)
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
    CardSpacing = ToTwips(m_lCardSpacing)
End Property

Public Property Let CardSpacing(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original
    m_lCardSpacing = ToPixels(lValue)
End Property

Public Property Get CardWidth() As Long
Attribute CardWidth.VB_Description = "Returns/sets the width of a card."
    CardWidth = ToTwips(m_lCardWidth)
End Property

Public Property Let CardWidth(ByVal lValue As Long)
    '--- stored in pixels; snapped to nearest like the original
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
    Dim lLastRow        As Long

    If m_lRow <> lValue Then
        pvEndEdit
        pvCommitRow
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

'--- the level a group row sits at, 0 for a record: read off the page buffer
'--- where the position is on it, and off a wrapper minted for the occasion
'--- where it is not -- which is what a client asking about an offscreen row
'--- comes to
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
    '--- a sort moves rows around, so this is where the client app asks what
    '--- it is actually looking at; a group row answers 0
    RowIndex = m_pDataModel.RowIndex(RowPosition)
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

    '--- no arg means the current row; horizontal scroll not implemented yet
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
End Sub

Public Sub Rebind(Optional HoldSortSettings As Variant)
Attribute Rebind.VB_Description = "Forces re-creation of the recordset."
    m_bInSet = True
    m_pDataModel.Refresh
    pvApplyHoldSort HoldSortSettings
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
    IsGroupItem = (m_pDataModel.RowIndex(Row) = 0 And Row >= 1 And Row <= m_pDataModel.RowCount)
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
    m_bWindowDirty = True
    pvInvalidate
End Sub

Public Sub RefreshRowIndex(ByVal RowIndex As Long)
Attribute RefreshRowIndex.VB_Description = "Refreshes data of the record that matches the index."
    m_pDataModel.RefreshRowIndex RowIndex
    m_bWindowDirty = True
    pvInvalidate
End Sub

Public Function GroupRowLevel(ByVal RowPosition As Long) As Integer
Attribute GroupRowLevel.VB_Description = "Returns the level of a group row."
    Dim lSlot           As Long

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
    m_bWindowDirty = True
    If Not m_bInSet Then
        pvInvalidate
    End If
End Sub

Friend Function frColIsGrouped(ByVal lColIndex As Long) As Boolean
    Dim lIdx            As Long

    For lIdx = 1 To m_oGroups.Count
        If m_oGroups.Item(lIdx).ColIndex = lColIndex Then
            frColIsGrouped = True
            Exit For
        End If
    Next
End Function

Friend Function frColSortOrder(ByVal lColIndex As Long) As jgexSortOrderConstants
    Dim lIdx            As Long
    Dim oKey            As JSSortKey
    Dim oGroup          As JSGroup

    '--- grouping sorts by the column too, and the original marks the header
    '--- with the same arrow an explicit sort key gets
    For lIdx = 1 To m_oGroups.Count
        Set oGroup = m_oGroups.Item(lIdx)
        If oGroup.ColIndex = lColIndex Then
            frColSortOrder = oGroup.SortOrder
            Exit Function
        End If
    Next
    For lIdx = 1 To m_oSortKeys.Count
        Set oKey = m_oSortKeys.Item(lIdx)
        If oKey.ColIndex = lColIndex Then
            frColSortOrder = oKey.SortOrder
            Exit Function
        End If
    Next
End Function

Friend Sub frRaiseUnboundReadData(ByVal lRowIndex As Long, vBookmark As Variant, oValues As JSRowData)
    RaiseEvent UnboundReadData(lRowIndex, vBookmark, oValues)
End Sub

Friend Sub frRaiseUnboundAddNew(oNewRowBookmark As JSRetVariant, oValues As JSRowData)
    RaiseEvent UnboundAddNew(oNewRowBookmark, oValues)
End Sub

Friend Sub frRaiseUnboundUpdate(ByVal lRowIndex As Long, vBookmark As Variant, oValues As JSRowData)
    '--- an edit commit reaches the model the same way an explicit Update
    '--- does, but the original says nothing to the client about it: it puts
    '--- the row in its own store and reports only AfterUpdate. Pinned by the
    '--- event goldens of 058, 068 and 071
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

'= private ===============================================================

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
    Dim lIdx            As Long

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
    '--- the wrappers are dropped rather than terminated: a buffer carries its
    '--- own row and never reads back through the control, so one a client
    '--- kept goes on answering for the row it was filled with
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
        Else
            Set m_aWindow(lIdx) = m_pDataModel.GetRowData(lFirst + lIdx - 1)
        End If
        RaiseEvent RowFormat(m_aWindow(lIdx))
    Next
    pvSyncRowData
End Sub

Private Sub pvRemapCurrent()
    Dim oItem           As JSSelectedItem

    If m_lHoldRowIndex > 0 Then
        m_lRow = m_pDataModel.GetRowPosition(m_lHoldRowIndex)
        pvSyncRowData
    End If
    For Each oItem In m_oSelectedItems
        oItem.frSetPosition m_pDataModel.GetRowPosition(oItem.RowIndex)
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
        m_bWindowDirty = True
    End If
    '--- a group row reports no index, and holding 0 would lose the record the
    '--- marquee came from, so the last record seen stands
    lRowIndex = m_pDataModel.RowIndex(m_lRow)
    If lRowIndex > 0 Then
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

    '--- an expand or a collapse only reprojects: the sort order underneath
    '--- stays exactly as it was
    '--- collapsing takes rows away, so a view scrolled down into them has to
    '--- come back up rather than leave the block short of its own bottom
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

Private Function pvIsGroupRow(ByVal lPos As Long) As Boolean
    If lPos >= 1 And lPos <= m_pDataModel.RowCount Then
        pvIsGroupRow = (m_pDataModel.RowIndex(lPos) = 0)
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

Private Sub pvPaint(ByVal hDC As Long)
    Dim lY              As Long

    '--- a sort the caller set up since the last paint takes effect here
    pvSyncProjection
    If m_bGroupByBoxVisible Then
        pvPaintGroupByBox hDC, lY
    End If
    If m_bColumnHeaders Then
        pvPaintHeaders hDC, lY
    End If
    pvPaintRows hDC, lY
End Sub

Private Sub pvPaintGroupByBox(ByVal hDC As Long, lY As Long)
    Dim lBoxH           As Long
    Dim lTotalH         As Long
    Dim uRect           As RECT
    Dim hPrevFont       As Long

    hPrevFont = pvSelectFont(hDC, m_oFont)
    lBoxH = m_lColumnHeaderHeight + 4
    lTotalH = pvGroupByBoxHeight()
    pvFillRect hDC, 0, lY, picGrid.ScaleWidth, lY + lTotalH, m_clrBackColorGBBox
    If m_oGroups.Count > 0 Then
        '--- the grouped columns take the box over, one chip each
        pvPaintGroupChips hDC, lY
    Else
        '--- info box is sized by the info text extent
        Call DrawText(hDC, StrPtr(m_sGroupByBoxInfoText), Len(m_sGroupByBoxInfoText), uRect, DT_SINGLELINE Or DT_CALCRECT)
        pvFillRect hDC, 4, lY + 5, 12 + uRect.Right, lY + 5 + lBoxH, m_clrBackColorInfoText
        pvDrawText hDC, m_sGroupByBoxInfoText, 7, lY + 5, 7 + uRect.Right, lY + 5 + lBoxH, m_clrForeColorInfoText, m_clrBackColorInfoText, jgexAlignLeft, 7, 7 + uRect.Right
    End If
    Call SelectObject(hDC, hPrevFont)
    lY = lY + lTotalH
End Sub

Private Function pvChipStagger() As Long
    '--- each chip after the first steps down half a header row -- taken from
    '--- the stored header height rather than re-measured, which is what keeps
    '--- it in step with the band it sits in
    pvChipStagger = m_lColumnHeaderHeight \ 2
End Function

Private Sub pvLayoutGroupChips(ByVal hDC As Long, ByVal lY As Long)
    Dim lIdx            As Long
    Dim oGroup          As JSGroup
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lChipH          As Long
    Dim uRect           As RECT
    Dim uMetrics        As TEXTMETRICW

    '--- a chip per group level, sized off its column caption with the room
    '--- the original leaves for the drop affordance, each stepping right past
    '--- the one before it and down by half a header row -- the staircase the
    '--- original draws. Every level keeps the rectangle it landed on so the
    '--- painter and the hit-test never disagree about where a chip is
    uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
    lChipH = uMetrics.tmHeight + 6
    lLeft = CHIP_LEFT
    lTop = lY + CHIP_TOP
    For lIdx = 1 To m_oGroups.Count
        Set oGroup = m_oGroups.Item(lIdx)
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
            lTop = lTop + pvChipStagger()
        End If
        oGroup.frChipRect = uRect
    Next
End Sub

Private Sub pvPaintGroupChips(ByVal hDC As Long, ByVal lY As Long)
    Dim lIdx            As Long
    Dim oGroup          As JSGroup
    Dim sCaption        As String
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lChipH          As Long
    Dim lElbowTop       As Long
    Dim lChipW          As Long
    Dim uMetrics        As TEXTMETRICW

    '--- a raised button carrying the column caption and the same sort arrow
    '--- its header shows, on the rectangle the layout gave it
    uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
    pvLayoutGroupChips hDC, lY
    For lIdx = 1 To m_oGroups.Count
        Set oGroup = m_oGroups.Item(lIdx)
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
                '--- same rule as the column header, measured from the
                '--- chip's own text top
                pvPaintSortGlyph hDC, lLeft + 2 + pvTextWidth(hDC, sCaption) + 4, _
                    lTop + 3 + uMetrics.tmHeight \ 2 + 4, oGroup.SortOrder
            End If
            '--- consecutive levels are joined by an elbow: down out of the
            '--- chip above, then right into the one below it, meeting it a
            '--- line below its top edge
            If lIdx < m_oGroups.Count Then
                '--- it meets the next chip three quarters of the way down
                '--- that chip, less a pixel -- nothing to do with the font.
                '--- Measured off the original over seven chip heights (19,
                '--- 22, 26, 25, 30, 31 and 37px: three faces at two or three
                '--- DPIs) the leg lands at 13, 16, 18, 18, 22, 22 and 27.
                '--- Three of those are exactly .5 and all three go the way
                '--- CLng rounds -- to even -- which is why no \ form of this
                '--- fits: 15.5 has to give 16 while 18.5 has to give 18
                lElbowTop = lTop + pvChipStagger() + CLng((3 * lChipH - 4) / 4)
                pvLine hDC, lLeft + lChipW - 5, lTop + lChipH, lLeft + lChipW - 5, lElbowTop + 1, vb3DDKShadow, PS_SOLID
                pvLine hDC, lLeft + lChipW - 5, lElbowTop, lLeft + lChipW + CHIP_GAP, lElbowTop, vb3DDKShadow, PS_SOLID
            End If
        End If
    Next
End Sub

Private Sub pvPaintHeaders(ByVal hDC As Long, lY As Long)
    Dim lHdrH           As Long
    Dim uMetrics        As TEXTMETRICW
    Dim lX              As Long
    Dim oCol            As JSColumn
    Dim lW              As Long
    Dim hPrevFont       As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long

    lHdrH = m_lColumnHeaderHeight
    hPrevFont = pvSelectFont(hDC, m_oColumnHeaderFont)
    '--- band frame lines first; cell edges painted after overwrite them
    '--- at the cell boundaries
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
    '--- corner cell above the row headers column
    If m_bRowHeaders Then
        pvPaintHeaderCell hDC, 0, lY, CHROME_COL_W, lHdrH, vbNullString, jgexAlignLeft
    End If
    lX = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        lW = pvColWidth(oCol)
        pvPaintHeaderCell hDC, lX, lY, lW, lHdrH, oCol.Caption, oCol.HeaderAlignment
        '--- a sorted column carries the arrow right after its caption,
        '--- sitting on its baseline
        If frColSortOrder(oCol.Index) <> 0 Then
            uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
            '--- the arrow centres on the caption rather than sitting on its
            '--- baseline: the two coincide at the default font, which is why
            '--- 16, 19 and 25 pixel fonts were needed to tell them apart
            pvPaintSortGlyph hDC, lX + 2 + pvTextWidth(hDC, oCol.Caption) + 4, _
                lY + (lHdrH - uMetrics.tmHeight + 1) \ 2 + uMetrics.tmHeight \ 2 + 4, frColSortOrder(oCol.Index)
        End If
        lX = lX + lW
    Next
    If pvGroupIndent() > 0 Then
        '--- the indent has no header cell of its own: the band shows through,
        '--- so the first column's left highlight goes back to background --
        '--- all but the band's own highlight down the very left edge
        pvFillRect hDC, IIf(m_bRowHeaders, CHROME_COL_W, 1), lY + 1, pvBlockLeft() + 1, lY + lHdrH - 2, m_clrBackColorHeader
    End If
    '--- filler cell up to the right edge (its right border is clipped off)
    If lX < picGrid.ScaleWidth Then
        pvPaintHeaderCell hDC, lX, lY, picGrid.ScaleWidth - lX + 2, lHdrH, vbNullString, jgexAlignLeft
    End If
    Call SelectObject(hDC, hPrevFont)
    lY = lY + lHdrH
End Sub

Private Sub pvPaintSortGlyph(ByVal hDC As Long, ByVal lX As Long, ByVal lBottom As Long, ByVal eOrder As jgexSortOrderConstants)
    Const GLYPH_W       As Long = 8
    Const GLYPH_H       As Long = 7
    Dim lRow            As Long
    Dim lEdge           As Long
    Dim lStep           As Long
    Dim lTop            As Long

    '--- an engraved triangle: shadow down the left edge, highlight down the
    '--- right one and along the base. A fixed 8x7 bitmap -- the original
    '--- draws the same pixels at 120dpi as at 96 -- sitting on the caption
    '--- baseline, which is what keeps it in place when the header grows
    lTop = lBottom - GLYPH_H + 1
    For lRow = 0 To GLYPH_H - 1
        lEdge = lRow
        If eOrder = jgexSortDescending Then
            lEdge = GLYPH_H - 1 - lRow
        End If
        '--- the edges spread one pixel every second row, and the row that
        '--- steps carries the pixel it stepped away from as well
        lStep = (lEdge + 1) \ 2
        pvSetPixel hDC, lX + GLYPH_W \ 2 - 1 - lStep, lTop + lRow, vb3DShadow
        pvSetPixel hDC, lX + GLYPH_W \ 2 + lStep, lTop + lRow, vb3DHighlight
        If lEdge Mod 2 = 1 Then
            pvSetPixel hDC, lX + GLYPH_W \ 2 - lStep, lTop + lRow, vb3DShadow
            pvSetPixel hDC, lX + GLYPH_W \ 2 - 1 + lStep, lTop + lRow, vb3DHighlight
        End If
    Next
    If eOrder = jgexSortDescending Then
        '--- pointing down the flat edge is on top and in shadow, the mirror
        '--- of the highlighted base the ascending arrow stands on
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
        '--- DT_VCENTER over the full cell height, not an inset rect: same
        '--- result at 96dpi, a pixel higher at 120dpi
        pvDrawText hDC, sCaption, lX + 2, lY, lX + lW - 2, lY + lH, m_clrForeColorHeader, m_clrBackColorHeader, eAlign, lX + 2, lX + lW - 2
    End If
End Sub

Private Sub pvPaintRows(ByVal hDC As Long, ByVal lY As Long)
    Dim lRowH           As Long
    Dim lHdrW           As Long
    Dim lTotalW         As Long
    Dim lRow            As Long
    Dim lFirst          As Long
    Dim lRowTop         As Long
    Dim lRowsBottom     As Long
    Dim lCum            As Long
    Dim lCumF           As Long
    Dim lW              As Long
    Dim oCol            As JSColumn
    Dim hPrevFont       As Long
    Dim uRect           As RECT
    Dim lPainted        As Long
    Dim lExtra          As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long

    lRowH = m_lRowHeight
    lHdrW = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        lTotalW = lTotalW + pvColWidth(m_oColumns.ItemByPosition(vOrder(lIdx)))
    Next
    '--- background right of the columns down to the bottom
    pvFillRect hDC, lHdrW + lTotalW, lY, picGrid.ScaleWidth, picGrid.ScaleHeight, m_clrBackColorBkg
    pvSyncProjection
    pvPopulateWindow
    lFirst = m_lFirstItem
    If lFirst < 1 Then
        lFirst = 1
    End If
    If lRowH > 0 Then
        hPrevFont = pvSelectFont(hDC, m_oFont)
        For lRow = lFirst To RowCount
            lRowTop = lY + (lRow - lFirst) * lRowH
            If lRowTop >= picGrid.ScaleHeight Then
                Exit For
            End If
            If pvIsGroupRow(lRow) Then
                pvPaintGroupRow hDC, lRow, lRowTop, lRowH, lHdrW, lHdrW + lTotalW, lY
            Else
                pvPaintDataRow hDC, lRow, lRowTop, lRowH, lHdrW, lTotalW
            End If
            lPainted = lPainted + 1
        Next
        Call SelectObject(hDC, hPrevFont)
    End If
    lRowsBottom = lY + lPainted * lRowH
    '--- under vertical-only gridlines the row header's border pair ends one
    '--- line below the block, and the gridlines run down to meet it, while
    '--- the cells themselves still stop at the block bottom
    If m_eGridLines = jgexGLVertical Then
        lExtra = 1
    End If
    If m_bEmptyRows Then
        '--- empty rows continue the grid to the bottom edge
        pvFillRect hDC, lHdrW, lRowsBottom, lHdrW + lTotalW, picGrid.ScaleHeight, m_clrBackColor
        If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLHorizontal Then
            lRowTop = lRowsBottom
            Do While lRowTop + lRowH - 1 < picGrid.ScaleHeight
                pvLine hDC, lHdrW, lRowTop + lRowH - 1, lHdrW + lTotalW, lRowTop + lRowH - 1, m_clrGridLinesColor, pvPenStyle()
                lRowTop = lRowTop + lRowH
            Loop
        End If
        lRowsBottom = picGrid.ScaleHeight
    Else
        '--- background below the last row
        pvFillRect hDC, 0, lRowsBottom + lExtra, lHdrW, picGrid.ScaleHeight, m_clrBackColorBkg
        pvFillRect hDC, lHdrW, lRowsBottom, lHdrW + lTotalW, picGrid.ScaleHeight, m_clrBackColorBkg
        '--- the line closing the block runs the whole width when the rows
        '--- are grouped, indent included
        If m_oGroups.Count > 0 And lPainted > 0 Then
            pvLine hDC, 0, lRowsBottom - 1, lHdrW, lRowsBottom - 1, m_clrGridLinesColor, pvPenStyle()
        End If
    End If
    If lPainted > 0 Or m_bEmptyRows Then
        '--- focus marquee on the current row; the XOR runs against the DC
        '--- background color which the original keeps at BackColor, and
        '--- vertical gridlines paint over the marquee dots
        If m_lRow >= lFirst And m_lRow < lFirst + lPainted Then
            pvPaintRowMarquee hDC, lY + (m_lRow - lFirst) * lRowH, lRowH, lHdrW, lTotalW
        End If
        '--- vertical gridlines over the rows block
        If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLVertical Then
            For lIdx = 0 To pvOrderMax(vOrder)
                lCum = lCum + pvColWidth(m_oColumns.ItemByPosition(vOrder(lIdx)))
                If m_oGroups.Count > 0 Then
                    '--- a group row spans the block, so the column rules
                    '--- break at every one of them
                    For lRow = lFirst To lFirst + lPainted - 1
                        If Not pvIsGroupRow(lRow) Then
                            pvLine hDC, lHdrW + lCum - 1, lY + (lRow - lFirst) * lRowH, _
                                lHdrW + lCum - 1, lY + (lRow - lFirst + 1) * lRowH, m_clrGridLinesColor, pvPenStyle()
                        End If
                    Next
                ElseIf m_eGridLineStyle = jgexGLSDashes Then
                    pvDashedVLine hDC, lHdrW + lCum - 1, lY, lRowsBottom + lExtra, m_clrGridLinesColor, lY, lRowH, lRowsBottom
                Else
                    pvLine hDC, lHdrW + lCum - 1, lY, lHdrW + lCum - 1, lRowsBottom + lExtra, m_clrGridLinesColor, pvPenStyle()
                End If
            Next
        End If
    End If
    '--- separator strip the original draws along the bottom edge when the
    '--- band takes over the last client row -- the horizontal scrollbar puts
    '--- it there, and so does a record navigator with no scrollbar beside it
    If m_bBandVisible Then
        pvLine hDC, 0, picGrid.ScaleHeight - 1, picGrid.ScaleWidth, picGrid.ScaleHeight - 1, m_clrBackColorHeader, PS_SOLID
    End If
End Sub

Private Sub pvPaintDataRow(ByVal hDC As Long, ByVal lRow As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal lTotalW As Long)
    Dim bSelected       As Boolean
    Dim bCurrentRow     As Boolean
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
    Dim lMarqueeR       As Long
    Dim lLineR          As Long
    Dim vOrder          As Variant
    Dim lIdx            As Long

    '--- the current row is always shown selected, as in the original
    bSelected = pvIsRowSelected(lRow) Or (m_lRow >= 1 And lRow = m_lRow)
    '--- a block of rows has no current cell to lift out of it: what is
    '--- selected paints as one, and only a selection of one row singles the
    '--- cell the marquee is on out of it
    bCurrentRow = bSelected And lRow = m_lRow And m_oSelectedItems.Count <= 1
    If bSelected Then
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
        '--- the header cell can be shorter than the row, and what shows below
        '--- it is the grid background rather than the row's own color
        pvFillRect hDC, 0, lRowTop, CHROME_COL_W, lRowTop + lRowH, m_clrBackColor
        '--- the record-selector arrow marks the current row only
        pvPaintRowHeader hDC, lRowTop, lRowH, CHROME_COL_W, (lRow = m_lRow)
    End If
    If pvGroupIndent() > 0 Then
        '--- the indent is group chrome, painted like a group row rather than
        '--- like the grid background beside the columns, with a rule at every
        '--- level boundary the record is nested behind
        pvFillRect hDC, pvRowHeaderWidth(), lRowTop, lHdrW, lRowTop + lRowH, m_clrBackColorRowGroup
        pvPaintIndentRules hDC, lRowTop, lRowH, pvGroupIndent()
    End If
    pvFillRect hDC, lHdrW, lRowTop, lHdrW + lTotalW, lRowTop + lRowH, clrBack
    '--- text of the current row clips inside the marquee, so a cell running
    '--- past the client edge stops short of the marquee's right border
    lMarqueeR = -1
    If lRow = m_lRow Then
        lMarqueeR = lHdrW + lTotalW - 2
        If lMarqueeR > picGrid.ScaleWidth - 1 Then
            lMarqueeR = picGrid.ScaleWidth - 1
        End If
    End If
    lX = lHdrW
    lPos = 0
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oCol.Visible Then
            lPos = lPos + 1
            lW = pvColWidth(oCol)
            '--- inside a selected row the current cell keeps the plain colors:
            '--- Col = 0 selects the whole row and nothing is singled out, but
            '--- once a cell is current the original lifts it out of the block
            If bCurrentRow And m_lCol = lPos Then
                '--- inset where the marquee runs, so the row's selection colour
                '--- still shows in the band its XOR checkerboard inverts
                '--- against -- that is the row's own edge, not the boundary
                '--- between two cells, which no marquee follows
                lFillL = lX
                lFillR = lX + lW
                If lFillL <= lHdrW Then
                    lFillL = lHdrW + 1
                End If
                If lFillR > pvMarqueeRight(lHdrW, lTotalW) Then
                    lFillR = pvMarqueeRight(lHdrW, lTotalW)
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
                pvDrawText hDC, sText, lX + 2, lRowTop, lX + lW - 3, lRowTop + pvRowContentH(lRowH), clrCellText, clrCellBack, oCol.TextAlignment, lX, lClipR, oCol.WordWrap
            End If
            lX = lX + lW
        End If
    Next
    If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLHorizontal Then
        '--- with no vertical gridline claiming the block's last pixel column
        '--- the horizontal line runs one further right, as the marquee does
        lLineR = lHdrW + lTotalW
        If m_eGridLines = jgexGLHorizontal Then
            lLineR = lLineR + 1
        End If
        If lRow = RowCount Then
            '--- the line under the last data row is drawn dark
            pvLine hDC, lHdrW, lRowTop + lRowH - 1, lLineR, lRowTop + lRowH - 1, vb3DDKShadow, PS_SOLID
        Else
            pvLine hDC, lHdrW, lRowTop + lRowH - 1, lLineR, lRowTop + lRowH - 1, m_clrGridLinesColor, pvPenStyle()
        End If
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

    '--- XOR checkerboard anchored per column: a border pixel is inverted
    '--- when (x - column left) + (y - row top) is odd, which a single
    '--- row-wide DrawFocusRect only reproduces on even column boundaries.
    '--- Every border pixel is inverted, alternating the mask between white
    '--- and ForeColor like a two-color pattern brush would -- the ForeColor
    '--- pass is invisible at the default black, which is why it only shows
    '--- up once a scenario sets the property
    '--- the vertical gridline claims the last pixel column of the block, so
    '--- the marquee stops one short of it -- with vertical lines off it runs
    '--- all the way out, mirroring pvRowContentH on the other axis
    '--- a group row spans the block as a single run: it starts at the very
    '--- left edge, tree indent included, no column rule breaks the pattern
    '--- and with no vertical gridline to yield to it runs the full width out
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

Private Sub pvPaintGroupRow(ByVal hDC As Long, ByVal lPos As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lLeft As Long, ByVal lRight As Long, ByVal lBlockTop As Long)
    Dim oRowData        As JSRowData
    Dim bFooter         As Boolean
    Dim clrBack         As OLE_COLOR
    Dim clrText         As OLE_COLOR
    Dim lIndent         As Long
    Dim lBoxTop         As Long
    Dim lBoxLeft        As Long
    Dim lTextTop        As Long
    Dim lTextLeft       As Long
    Dim lLineLeft       As Long
    Dim uMetrics        As TEXTMETRICW

    '--- a group row spans the whole block in header colour, with the expand
    '--- box in the chrome column at its left and the group value beside it,
    '--- and takes the selection colours whenever the marquee is on it
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
    '--- the rules of the levels this group sits inside run through it
    pvPaintIndentRules hDC, lRowTop, lRowH, lIndent
    '--- the row's last line is the separator shared with whatever follows,
    '--- so it starts where the records start
    pvLine hDC, pvBlockLeft() - 1, lRowTop + lRowH - 1, lRight, lRowTop + lRowH - 1, m_clrGridLinesColor, pvPenStyle()
    '--- the line above it starts at this group's own edge instead: a level
    '--- one group opens the full width, a nested one only its own part. A
    '--- footer closes a group rather than opening one, so the rule above it
    '--- belongs to the records and starts where they do
    If lRowTop > lBlockTop Then
        lLineLeft = pvRowHeaderWidth() + lIndent - 1
        If bFooter Then
            lLineLeft = pvBlockLeft() - 1
        End If
        pvLine hDC, lLineLeft, lRowTop - 1, lRight, lRowTop - 1, m_clrGridLinesColor, pvPenStyle()
    End If
    lBoxLeft = pvRowHeaderWidth() + lIndent + (GROUP_INDENT_W - GROUP_BOX_W) \ 2
    lBoxTop = lRowTop + (lRowH - 1 - GROUP_BOX_W) \ 2
    '--- a footer carries no expand box: it closes the group rather than
    '--- opening it, and under the totals style it reads across the columns
    '--- instead of carrying the caption
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
    '--- the caption is drawn a space past the expand box, plus the two pixel
    '--- margin text keeps everywhere else -- this tracks the font at any dpi
    '--- the caption sits two pixels past the expand box behind a leading
    '--- space, which a prefix supplies itself -- so a prefixed level starts
    '--- one space earlier and its value still lands where a plain one does
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

    '--- raised, like a small button: it keeps the button face whatever the
    '--- group row colour is, while the sign follows ForeColorRowGroup
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

Private Function pvRowIndent(ByVal lPos As Long) As Long
    '--- a record sits inside every level; a group row sits inside the levels
    '--- above its own
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
    '--- headers, rows and hit-testing all start the column block here: the
    '--- row header column if there is one, plus a chrome column per group
    '--- level for the tree indent
    If m_bRowHeaders Then
        pvBlockLeft = CHROME_COL_W
    End If
    pvBlockLeft = pvBlockLeft + pvGroupIndent()
End Function

Private Function pvGroupIndent() As Long
    '--- one chrome column per group level, which is what the records are
    '--- pushed right by
    pvGroupIndent = m_oGroups.Count * GROUP_INDENT_W
End Function

Private Sub pvPaintRowHeader(ByVal hDC As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal bCurrent As Boolean)
    Dim lIdx            As Long
    Dim lC              As Long

    '--- a Double3D header-like cell per row. Where its border pair lands is
    '--- an empirical table, one entry per GridLines value, each measured off
    '--- a golden -- no mechanical rule was found that explains why the
    '--- single-direction modes pull it in opposite directions
    Select Case m_eGridLines
    Case jgexGLVertical
        lC = lRowH
    Case jgexGLHorizontal
        lC = lRowH - 2
    Case Else
        lC = lRowH - 1
    End Select
    pvFillRect hDC, 1, lRowTop + 1, lHdrW - 2, lRowTop + lC - 1, m_clrBackColorHeader
    pvLine hDC, 0, lRowTop, lHdrW, lRowTop, vb3DHighlight, PS_SOLID
    pvLine hDC, 0, lRowTop, 0, lRowTop + lC - 1, vb3DHighlight, PS_SOLID
    pvLine hDC, 0, lRowTop + lC - 1, lHdrW, lRowTop + lC - 1, vb3DShadow, PS_SOLID
    pvLine hDC, 0, lRowTop + lC, lHdrW, lRowTop + lC, vb3DDKShadow, PS_SOLID
    pvLine hDC, lHdrW - 2, lRowTop, lHdrW - 2, lRowTop + lC - 1, vb3DShadow, PS_SOLID
    pvLine hDC, lHdrW - 1, lRowTop, lHdrW - 1, lRowTop + lC, vb3DDKShadow, PS_SOLID
    '--- current row arrow marker
    If bCurrent Then
        For lIdx = 0 To 5
            pvLine hDC, 6 + lIdx, lRowTop + 3 + lIdx, 6 + lIdx, lRowTop + 15 - lIdx, vbButtonText, PS_SOLID
        Next
    End If
End Sub

Private Function pvIsChecked(oRowData As JSRowData, ByVal lColIndex As Long) As Boolean
    Dim vValue          As Variant

    If oRowData Is Nothing Then
        Exit Function
    End If
    '--- a toggle buffers on the row like any other edit, and the buffer is the
    '--- row, so this reads what was toggled without going anywhere else
    vValue = oRowData.Value(lColIndex)
    If Not IsObject(vValue) And Not IsArray(vValue) Then
        If Not pvIsBlank(vValue) Then
            pvIsChecked = CBool(vValue)
        End If
    End If
End Function

Private Function pvMarqueeRight(ByVal lHdrW As Long, ByVal lTotalW As Long) As Long
    '--- where the current row's marquee runs down: the vertical gridline
    '--- claims the block's last pixel column, so it stops one short of it
    pvMarqueeRight = lHdrW + lTotalW - 1
    If m_eGridLines = jgexGLBoth Or m_eGridLines = jgexGLVertical Then
        pvMarqueeRight = pvMarqueeRight - 1
    End If
End Function

Private Sub pvPaintCheckBox(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal bChecked As Boolean, ByVal bCurrent As Boolean)
    Dim lIdx            As Long
    Dim vRuns           As Variant
    Dim clrBorder       As OLE_COLOR
    Dim lOffX           As Long
    Dim lOffY           As Long

    '--- a fixed 11x12 box at either dpi, like the rest of the control's
    '--- chrome: white inside with one grey line around it, a shade darker
    '--- on the cell the marquee is on
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
    '--- the box stays 11x12 at every scale but the tick in it is the system's
    '--- own and grows with the screen, so it sits further in as that happens:
    '--- two pixels at 96, three at 120. At 144 the mark has outgrown the box
    '--- and only its top four rows are inside it, both arms already there and
    '--- the vertex below -- which is why it is not the 96 shape scaled, whose
    '--- top row is one pixel on the right. The font does not move it -- a 12pt
    '--- cell font at 96 leaves it exactly where an 8.25pt one does -- and
    '--- neither DrawFrameControl nor a plain dpi ratio reproduces any of them,
    '--- so these are the ones the recordings show, a scale each
    '--- row, first column, last column
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
    '--- a cell with an uncommitted write paints what was written, not the
    '--- display value the row was decorated with before it
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
    '--- the horizontal gridline takes the row's last pixel line, so text and
    '--- the focus marquee center in one pixel less -- with horizontal lines
    '--- off the original uses the whole row height instead
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

    '--- tolerate refresh before the control window exists
    On Error GoTo EH
    '--- a batch of changes under Redraw = False paints once, when it is
    '--- turned back on
    If Not m_bRedraw Then
        Exit Sub
    End If
    If Not SkipScroll Then
        pvUpdateScrollBars
    End If
    pvLayoutEditor
    Call InvalidateRect(picGrid.hWnd, 0, 0)
    If m_bRecordNavigator Then
        Call InvalidateRect(UserControl.hWnd, 0, 1)
        If hsbGrid.Visible Then
            Call InvalidateRect(hsbGrid.hWnd, 0, 1)
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvTopHeight() As Long
    '--- the group-by box grows a staircase step per level past the first, so
    '--- the room left for rows shrinks with them -- which is what puts a
    '--- vertical scrollbar on a grouped grid that would otherwise just fit
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
    '--- the left-most visible column: horizontal scrolling moves whole
    '--- columns, so painting and hit-testing simply start here
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

    '--- the client strip the scrollable columns share, i.e. what is left of
    '--- it once the row header and the frozen block have taken their part
    pvScrollableWidth = picGrid.ScaleWidth
    If m_bRowHeaders Then
        pvScrollableWidth = pvScrollableWidth - 18
    End If
    lFrozen = pvFrozenCount()
    For lIdx = 1 To lFrozen
        Set oCol = m_oColumns.ItemByPosition(lIdx)
        If oCol.Visible Then
            pvScrollableWidth = pvScrollableWidth - pvColWidth(oCol)
        End If
    Next
End Function

Private Function pvVisibleColCount() As Long
    Dim lIdx            As Long

    For lIdx = 1 To m_oColumns.Count
        If m_oColumns.ItemByPosition(lIdx).Visible Then
            pvVisibleColCount = pvVisibleColCount + 1
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
    '--- taken from the metric rather than from picGrid: the grid surface is
    '--- laid out to leave exactly this much, and deriving it back the other
    '--- way picks up rounding from the Move
    uNav.BandH = GetSystemMetrics(SM_CYHSCROLL)
    uNav.BandTop = UserControl.ScaleHeight - uNav.BandH
    If uNav.BandH <= 0 Then
        Exit Function
    End If
    '--- RecordNavigatorString holds the two literals, pipe separated
    vSplit = Split(m_sRecordNavigatorString & "|", "|")
    uNav.Prefix = vSplit(0)
    uNav.Middle = vSplit(1) & " " & RowCount
    '--- gaps measured off the original: prefix at 4, then 2 before the first
    '--- button pair, 4 before the record box, 5 after it and 6 after the
    '--- "of N" text -- buttons are one pixel taller than the band is high
    lBtnW = uNav.BandH + 1
    lBtnTop = uNav.BandTop + 1
    lBtnH = uNav.BandH - 1
    '--- asked for rather than read off the control: HasDC is False on both
    '--- surfaces, so hDC answers for the paint in hand and nothing else, and
    '--- this layout is wanted from a hit test as much as from the paint
    hDC = GetDC(UserControl.hWnd)
    hPrevFont = pvSelectFont(hDC, m_oFont)
    lX = 4
    uNav.PrefixX = lX
    lX = lX + pvTextWidth(hDC, uNav.Prefix) + 4
    pvSetRect uNav.BtnFirst, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    pvSetRect uNav.BtnPrev, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW + 4
    '--- the box holds seven digits plus the 2px client border on each side,
    '--- which is why 5, 50 and 500 records all render identically: the
    '--- original's TextBox is 46px at 96dpi, 53 at 120 and 67 at 144, i.e.
    '--- exactly the width of "9999999" (42, 49 and 63) grown by the border.
    '--- Measured, not derived from tmAveCharWidth: that averages the whole
    '--- charset (5, 7 and 8 px here) and does not track the digit advance
    lBoxW = pvTextWidth(hDC, "9999999") + 4
    '--- the box is taller than the band and clipped by it, like the original's
    '--- TextBox (53x24 inside a 22px band at 120dpi)
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

    '--- GetTextExtentPoint32, which is what VB6's own TextWidth calls, not
    '--- DrawText with DT_CALCRECT: the two agree on the bitmap faces but part
    '--- company on TrueType, where CALCRECT drops the last glyph's overhang
    '--- -- "Region" in Segoe UI 14 measures 58 that way against the 59 the
    '--- original lays its group-by chip out with
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
    '--- a single-pixel white edge on top and left, shadow then dark shadow
    '--- down the right and bottom -- DrawFrameControl doubles the light edge
    pvFillRect hDC, uBtn.Left, uBtn.Top, uBtn.Right, uBtn.Bottom, m_clrBackColorHeader
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

    '--- 5 wide, 9 tall: the apex is a single pixel and each column towards
    '--- the base grows by one row either side
    '--- Bottom is exclusive, so the last row is Bottom - 1
    '--- 5x9 at 96dpi, 6x11 at 120dpi. The end buttons sit their glyph clear
    '--- of the bar; the inner pair centre theirs in the face
    lHalf = lBandH \ 4
    If bEnd Then
        lOfs = lBandH \ 4 + lBandH \ 7 + lBandH \ 5
    Else
        lOfs = (uBtn.Right - uBtn.Left - lHalf - 1) \ 2 + 1
    End If
    '--- the glyph centres a row above the button's own middle, which only
    '--- parts company with it once the band height turns odd: 247 either way
    '--- at 96 and 120dpi, 242 rather than 243 at 144
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

    If Not pvNavLayout(uNav) Then
        Exit Sub
    End If
    pvFillRect hDC, 0, uNav.BandTop, uNav.Width, uNav.BandTop + uNav.BandH, m_clrBackColorHeader
    '--- the pair pointing at an end greys out once the current row is there
    lAtFirst = 0
    If m_lRow <= 1 Then
        lAtFirst = DFCS_INACTIVE
    End If
    lAtLast = 0
    If m_lRow >= RowCount Then
        lAtLast = DFCS_INACTIVE
    End If
    '--- faces from the system, glyphs drawn here: DrawFrameControl's arrow is
    '--- 4x7 where the original's is 5x9, and the end buttons shift theirs
    '--- right/left by two to clear the bar that marks first and last
    pvNavButton hDC, uNav.BtnFirst
    pvNavButton hDC, uNav.BtnPrev
    pvNavButton hDC, uNav.BtnNext
    pvNavButton hDC, uNav.BtnLast
    pvNavArrow hDC, uNav.BtnFirst, False, False, True, uNav.BandH
    pvNavArrow hDC, uNav.BtnPrev, False, (m_lRow <= 1), False, uNav.BandH
    pvNavArrow hDC, uNav.BtnNext, True, (m_lRow >= RowCount), False, uNav.BandH
    pvNavArrow hDC, uNav.BtnLast, True, False, True, uNav.BandH
    '--- bar metrics scale with the band: 2px wide at 4 in from the edge at
    '--- 96dpi, 3px at 5 in at 120dpi and 3px at 7 in at 144
    lBarX = (uNav.BandH + 2) \ 4
    lBarW = uNav.BandH \ 7
    '--- the bar is exactly as tall as the arrow beside it and shares its
    '--- centre -- 9, 11 and 13 rows at 96, 120 and 144dpi
    lBarHalf = uNav.BandH \ 4
    lBarY = uNav.BtnFirst.Top + (uNav.BtnFirst.Bottom - uNav.BtnFirst.Top) \ 2 - 1 - lBarHalf
    pvFillRect hDC, uNav.BtnFirst.Left + lBarX, lBarY, uNav.BtnFirst.Left + lBarX + lBarW, lBarY + 2 * lBarHalf + 1, vbBlack
    pvFillRect hDC, uNav.BtnLast.Right - 1 - lBarX - lBarW, lBarY, uNav.BtnLast.Right - 1 - lBarX, lBarY + 2 * lBarHalf + 1, vbBlack
    uBox = uNav.Box
    Call DrawEdge(hDC, uBox, EDGE_SUNKEN, BF_RECT)
    pvFillRect hDC, uNav.Box.Left + 2, uNav.Box.Top + 2, uNav.Box.Right - 2, uNav.Box.Bottom - 2, vbWindowBackground
    hPrevFont = pvSelectFont(hDC, m_oFont)
    '--- the labels centre on the font box, not the band: rounding the half
    '--- pixel up is what keeps them on the original's row at every scale
    lTextH = FontTextMetrics(m_oFont, hDC).tmHeight
    lTextTop = uNav.BandTop + (uNav.BandH - lTextH + 1) \ 2
    pvDrawText hDC, uNav.Prefix, uNav.PrefixX, lTextTop, uNav.PrefixX + pvTextWidth(hDC, uNav.Prefix), lTextTop + lTextH, m_clrForeColorHeader, m_clrBackColorHeader, jgexAlignLeft, uNav.PrefixX, uNav.Width
    pvDrawText hDC, uNav.Middle, uNav.MiddleX, lTextTop, uNav.MiddleX + pvTextWidth(hDC, uNav.Middle), lTextTop + lTextH, m_clrForeColorHeader, m_clrBackColorHeader, jgexAlignLeft, uNav.MiddleX, uNav.Width
    '--- the record number sits at the top of the box's interior rather than
    '--- centring in what is left of the band, which the two only tell apart
    '--- once the band height turns odd: a pixel lower at 144dpi
    pvDrawText hDC, CStr(m_lRow), uNav.Box.Left + 2, uNav.Box.Top + 2, uNav.Box.Right - 4, uNav.Box.Top + 2 + lTextH, m_clrForeColor, vbWindowBackground, jgexAlignRight, uNav.Box.Left + 2, uNav.Box.Right - 2
    Call SelectObject(hDC, hPrevFont)
End Sub

Private Sub pvLayoutGrid()
    Dim lBandH          As Long

    '--- the grid surface owns everything above the scrollbar band, so its
    '--- own WS_VSCROLL stops where the band starts -- which is what makes
    '--- the vertical thumb geometry match the original
    If m_bBandVisible Then
        lBandH = GetSystemMetrics(SM_CYHSCROLL)
    End If
    picGrid.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - lBandH
End Sub

Private Sub pvLayoutHScroll(ByVal bNeedH As Boolean, ByVal bNeedV As Boolean)
    Dim lBandH          As Long
    Dim lNavW           As Long
    Dim uNav            As UcsNavLayout

    '--- a VB6 HScrollBar rather than a Win32 SCROLLBAR child: the original
    '--- is a VB6 control too, so the runtime draws both with the same code
    '--- and the shaft dither matches by construction
    If Not bNeedH Then
        hsbGrid.Visible = False
        m_hWndHScroll = 0
        Exit Sub
    End If
    lBandH = GetSystemMetrics(SM_CYHSCROLL)
    '--- the band stops at the client width, leaving the usual corner gap
    '--- under the vertical bar, as the original's does
    If pvNavLayout(uNav) Then
        lNavW = uNav.Width
    End If
    hsbGrid.Move lNavW, UserControl.ScaleHeight - lBandH, picGrid.ScaleWidth - lNavW, lBandH
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
    Dim lPage           As Long

    If m_bScrollUpdating Then
        Exit Sub
    End If
    m_bScrollUpdating = True
    lTopH = pvTopHeight()
    If m_bRowHeaders Then
        lHdrW = 18
    End If
    lColsW = lHdrW + pvTotalColWidth()
    '--- this routine has to decide the same way whether or not it has run
    '--- before, so it starts from the surface neither bar has touched: the
    '--- last pass already took the band out of picGrid's height and the
    '--- vertical bar out of its width, and both have to be added back or the
    '--- second call charges for them twice -- it then invents a horizontal
    '--- bar for columns that do fit and pays for it out of the vertical page,
    '--- which is a thumb short by a row
    lFullH = picGrid.ScaleHeight
    If m_bBandVisible Then
        lFullH = lFullH + GetSystemMetrics(SM_CYHSCROLL)
    End If
    lFullW = picGrid.ScaleWidth
    If (GetWindowLong(picGrid.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0 Then
        lFullW = lFullW + GetSystemMetrics(SM_CXVSCROLL)
    End If
    '--- both scrollbars interact: each one steals space from the other. The
    '--- band is a single strip the navigator and the horizontal bar share, so
    '--- it costs its height once -- the navigator holds it open on its own,
    '--- and then the bar moving in is free
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
    '--- only the vertical bar is a window style. The horizontal one is a
    '--- child SCROLLBAR so it can be positioned precisely: the record
    '--- navigator shares its strip and a non-client bar always spans the
    '--- whole edge. The original does exactly this -- WS_VSCROLL on the grid
    '--- window, a child scrollbar control in the band below it
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
    '--- the band hosts the navigator as well, so it can be there without a
    '--- horizontal scrollbar
    m_bBandVisible = bNeedH Or m_bRecordNavigator
    '--- the band appearing takes height away from the grid surface
    pvLayoutGrid
    pvLayoutHScroll bNeedH, bNeedV
    If bNeedH Then
        lPage = pvColsFitFrom(pvScrollableWidth(), pvFrozenCount() + 1)
        If lPage < 1 Then
            lPage = 1
        End If
        '--- the vertical block is done with it by now, and every field this
        '--- one reads is written below
        With uSi
            .cbSize = Len(uSi)
            .fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
            .nMin = pvFrozenCount() + 1
            .nMax = pvVisibleColCount() - lPage + 1
            If .nMax < .nMin Then
                .nMax = .nMin
            End If
            .nPage = 1
            m_lLeftCol = Clamp(m_lLeftCol, .nMin, .nMax)
            .nPos = m_lLeftCol
        End With
        Call SetScrollInfo(hsbGrid.hWnd, SB_CTL, uSi, 1)
    ElseIf m_lLeftCol > 1 Then
        '--- they all fit now, so nothing is scrolled off at all
        m_lLeftCol = 1
    End If
    m_bScrollUpdating = False
End Sub

Private Function pvTotalColWidth() As Long
    Dim lIdx            As Long
    Dim oCol            As JSColumn

    For lIdx = 1 To m_oColumns.Count
        Set oCol = m_oColumns.ItemByPosition(lIdx)
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

    '--- ColumnAutoResize stretches the visible columns to fill the client
    '--- width, keeping their proportions: the boundaries land on the rounded
    '--- running total, so the last column absorbs the rounding
    If Not m_bColumnAutoResize Then
        pvColWidth = ToPixels(oCol.Width)
        Exit Function
    End If
    lAvail = picGrid.ScaleWidth
    If m_bRowHeaders Then
        lAvail = lAvail - 18
    End If
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

    '--- painting and hit-testing walk the columns in this order: the frozen
    '--- ones pinned at the left, then the scrollable rest from LeftCol. Both
    '--- blocks skip hidden columns, so callers never test Visible again
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
            '--- strictly less: a column ending exactly on the edge does not
            '--- count as visible, which is what decides the thumb size when
            '--- the columns happen to fill the client precisely
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

    '--- selected-row colors come from the SelectedRow FormatStyle, with a
    '--- system-highlight fallback if the style has been removed
    Set oStyle = m_oFormatStyles.frItemOrNothing("SelectedRow")
    If oStyle Is Nothing Then
        clrBack = vbHighlight
        clrFore = vbHighlightText
    Else
        clrBack = oStyle.BackColor
        clrFore = oStyle.ForeColor
    End If
End Sub

Private Sub pvDashedVLine(ByVal hDC As Long, ByVal lX As Long, ByVal lY1 As Long, ByVal lY2 As Long, ByVal clrLine As OLE_COLOR, ByVal lRowsTop As Long, ByVal lRowH As Long, ByVal lWrapEnd As Long)
    Dim hBrush          As Long
    Dim hPrevBrush      As Long
    Dim lIdx            As Long
    Dim lOfs            As Long

    '--- dashes run 3 on / 3 off and the phase restarts at every row top, so a
    '--- row height that is not a multiple of 6 (19px at 96dpi) leaves a run of
    '--- four across the row boundary, exactly as the original draws it
    pvFillRect hDC, lX, lY1, lX + 1, lY2, m_clrBackColorBkg
    hBrush = CreateSolidBrush(pvColor(clrLine))
    hPrevBrush = SelectObject(hDC, hBrush)
    For lIdx = lY1 To lY2 - 1
        lOfs = lIdx - lRowsTop
        If lRowH > 0 Then
            If lIdx < lWrapEnd Then
                lOfs = lOfs Mod lRowH
            Else
                '--- past the last row the phase keeps running from that row's
                '--- top instead of restarting, which is what decides whether
                '--- the block's closing line is drawn at all
                lOfs = lIdx - (lWrapEnd - lRowH)
            End If
        End If
        If lOfs Mod 6 < 3 Then
            Call PatBlt(hDC, lX, lIdx, 1, 1, PATCOPY)
        End If
    Next
    Call SelectObject(hDC, hPrevBrush)
    Call DeleteObject(hBrush)
End Sub

Private Sub pvDottedLine(ByVal hDC As Long, ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long, ByVal clrLine As OLE_COLOR)
    Dim hBrush          As Long
    Dim hPrevBrush      As Long
    Dim lIdx            As Long

    '--- the whole run is laid down in the control background first: a dotted
    '--- gridline's gaps show it, not the row color underneath
    If lY1 = lY2 Then
        pvFillRect hDC, lX1, lY1, lX2, lY1 + 1, m_clrBackColorBkg
    Else
        pvFillRect hDC, lX1, lY1, lX1 + 1, lY2, m_clrBackColorBkg
    End If
    hBrush = CreateSolidBrush(pvColor(clrLine))
    hPrevBrush = SelectObject(hDC, hBrush)
    If lY1 = lY2 Then
        For lIdx = lX1 To lX2 - 1
            If (lIdx + lY1) Mod 2 = 0 Then
                Call PatBlt(hDC, lIdx, lY1, 1, 1, PATCOPY)
            End If
        Next
    Else
        For lIdx = lY1 To lY2 - 1
            If (lX1 + lIdx) Mod 2 = 0 Then
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

Private Sub pvLine(ByVal hDC As Long, ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long, ByVal clrLine As OLE_COLOR, ByVal lPenStyle As Long)
    Dim hPen            As Long
    Dim hPrevPen        As Long

    '--- GDI's cosmetic PS_DOT renders 3-on/3-off, while the original's dotted
    '--- gridlines are 1-on/1-off anchored on odd coordinates -- the same
    '--- absolute-parity pattern the focus marquee uses -- so they are stamped
    '--- as explicit pixels instead
    If lPenStyle = PS_DOT Then
        pvDottedLine hDC, lX1, lY1, lX2, lY2, clrLine
        Exit Sub
    End If
    hPen = CreatePen(lPenStyle, 1, pvColor(clrLine))
    hPrevPen = SelectObject(hDC, hPen)
    Call MoveToEx(hDC, lX1, lY1, 0)
    Call LineTo(hDC, lX2, lY2)
    Call SelectObject(hDC, hPrevPen)
    Call DeleteObject(hPen)
End Sub

Private Sub pvDrawText(ByVal hDC As Long, sText As String, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal clrText As OLE_COLOR, ByVal clrBack As OLE_COLOR, ByVal eAlign As jgexAlignmentConstants, ByVal ClipLeft As Long, ByVal ClipRight As Long, Optional ByVal bWordWrap As Boolean)
    Dim uRect           As RECT
    Dim lFlags          As Long
    Dim lSaved          As Long

    uRect.Left = lLeft
    uRect.Top = lTop
    uRect.Right = lRight
    uRect.Bottom = lBottom
    '--- wrapped text starts at the top of its cell, a pixel in, since it is
    '--- the lines it breaks into that fill the room rather than one line
    '--- centred in it
    If bWordWrap Then
        uRect.Top = lTop + 1
        uRect.Bottom = lBottom - 1
        lFlags = DT_WORDBREAK Or DT_NOPREFIX
    Else
        lFlags = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
    End If
    Select Case eAlign
    Case jgexAlignCenter
        lFlags = lFlags Or DT_CENTER
    Case jgexAlignRight
        lFlags = lFlags Or DT_RIGHT
    End Select
    '--- text is positioned in its own rect but clipped separately: data
    '--- cells clip to the whole cell so ClearType fringes bleed into the
    '--- inset like the original's, headers clip to the text rect
    lSaved = SaveDC(hDC)
    '--- clipped to the rect the text was laid into, which for wrapped text is
    '--- inset: DT_NOCLIP would otherwise let a line's opaque background run
    '--- out over the marquee band below it
    Call IntersectClipRect(hDC, ClipLeft, uRect.Top, ClipRight, uRect.Bottom)
    lFlags = lFlags Or DT_NOCLIP
    '--- opaque so glyph anti-aliasing blends against the known cell
    '--- background exactly like the original
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

    '--- the control takes its fonts from the container, as the original
    '--- does -- which is why the same scenario renders differently under a
    '--- host form using MS Sans Serif and one using Tahoma. Swapping the
    '--- object does not raise FontChanged, so the font-derived row and
    '--- header heights are recalculated explicitly
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
    '--- the grid surface carries input, painting and its own WS_VSCROLL,
    '--- while the outer control is the scrollbar band's parent and so the
    '--- one that receives WM_CTLCOLORSCROLLBAR for the child bar
    Set m_pSubclassPic = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.ControlSubclassProc(0, 0, 0, 0, 0))
    Set m_pSubclassCtl = InitSubclassingThunk(UserControl.hWnd, Me, pvAddressOfSubclassProc.ControlSubclassProc(0, 0, 0, 0, 0))
End Sub

Private Sub pvUnsubclass()
    TerminateSubclassingThunk m_pSubclassPic, Me
    TerminateSubclassingThunk m_pSubclassCtl, Me
End Sub

Public Function EditSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute EditSubclassProc.VB_MemberFlags = "40"
    Dim nKeyCode        As Integer
    Dim nShift          As Integer

    '--- the in-place editor's own messages: the grid never sees these, so
    '--- the key events a client expects are raised from here
    Select Case wMsg
    Case WM_KEYDOWN
        nKeyCode = CInt(wParam And &HFFFF&)
        nShift = pvShiftState()
        RaiseEvent KeyDown(nKeyCode, nShift)
        Select Case nKeyCode
        Case vbKeyReturn
            '--- it steps to the next row, so the row goes through with the cell
            pvEndEdit
            pvCommitRow
            If m_lRow < RowCount Then
                pvNavigate m_lRow + 1, m_lCol, 0, False
            End If
            Handled = True
        Case vbKeyEscape
            pvEndEdit bCancel:=True
            Handled = True
        End Select
    Case WM_CHAR
        RaiseEvent KeyPress(CInt(wParam And &HFFFF&))
    Case WM_KEYUP
        RaiseEvent KeyUp(CInt(wParam And &HFFFF&), pvShiftState())
    End Select
End Function

Public Function ControlSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute ControlSubclassProc.VB_MemberFlags = "40"
    Dim nKeyCode        As Integer
    Dim nShift          As Integer
    Dim nKeyAscii       As Integer

    If hWnd = UserControl.hWnd Then
        '--- the outer control is the band: navigator clicks plus the child
        '--- scrollbar's colour and activation messages
        Select Case wMsg
        Case WM_LBUTTONDOWN
            pvOnNavigatorClick GetXLParam(lParam), GetYLParam(lParam)
        Case WM_MOUSEACTIVATE
            If pvHitScrollBar() Then
                ControlSubclassProc = MA_NOACTIVATE
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
    Case WM_ERASEBKGND
        ControlSubclassProc = 1
        Handled = True
    Case WM_VSCROLL
        pvOnVScroll LoWord(wParam), HiWord(wParam)
        Handled = True
    Case WM_KEYDOWN
        nKeyCode = CInt(wParam And &HFFFF&)
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
        RaiseEvent MouseUp(vbLeftButton, pvMouseShift(wParam), GetXLParam(lParam) * Screen.TwipsPerPixelX, GetYLParam(lParam) * Screen.TwipsPerPixelY)
        RaiseEvent Click
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(pvMouseButton(wParam), pvMouseShift(wParam), GetXLParam(lParam) * Screen.TwipsPerPixelX, GetYLParam(lParam) * Screen.TwipsPerPixelY)
        If (wParam And MK_LBUTTON) <> 0 Then
            pvOnMouseDrag GetYLParam(lParam)
        End If
    Case WM_CHAR
        nKeyAscii = CInt(wParam And &HFFFF&)
        RaiseEvent KeyPress(nKeyAscii)
    Case WM_KEYUP
        '--- a key that closed the editor releases over the grid, since the
        '--- editor window is gone by then
        RaiseEvent KeyUp(CInt(wParam And &HFFFF&), pvShiftState())
    Case WM_COMMAND
        '--- the editor is a child of the grid, so what it has to say about
        '--- itself arrives here: EN_CHANGE is the Change the client hears,
        '--- minus the one the control makes putting the value in
        If lParam = m_hWndEdit And m_hWndEdit <> 0 Then
            If (wParam \ &H10000) = EN_CHANGE And Not m_bInEditSetup Then
                RaiseEvent Change
            End If
        End If
    End Select
QH:
'    '--- note: performance optimization for design-time subclassing
'    If Not Handled And ThunkPrivateData(m_pSubclassPic) = EBMODE_DESIGN Then
'        Handled = True
'        ControlSubclassProc = CallNextSubclassProc(m_pSubclassPic, hWnd, wMsg, wParam, lParam)
'    End If
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

    '--- dragging with the left button held extends a multi-select range
    '--- from the mouse-down anchor to the row under the cursor
    If Not m_bMultiSelect Or m_lRowHeight <= 0 Or m_bClickOpenedEdit Then
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

Private Function pvColAtX(ByVal lX As Long, oCol As JSColumn) As Long
    Dim lCum            As Long
    Dim lPos            As Long
    Dim oItem           As JSColumn
    Dim vOrder          As Variant
    Dim lIdx            As Long

    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        lPos = lPos + 1
        lCum = lCum + pvColWidth(oItem)
        If lX < lCum Then
            Set oCol = oItem
            pvColAtX = lPos
            Exit Function
        End If
    Next
End Function

Private Function pvGroupAtPoint(ByVal lX As Long, ByVal lY As Long, oGroup As JSGroup) As Long
    Dim lIdx            As Long
    Dim oItem           As JSGroup
    Dim hDC             As Long
    Dim hPrevFont       As Long

    '--- laid out again rather than read off the last paint: a control that
    '--- has not painted yet still hit-tests, and a caption or font changed
    '--- since would otherwise leave the rectangles behind. It costs one text
    '--- measurement per group level, on a click
    hDC = GetDC(picGrid.hWnd)
    hPrevFont = pvSelectFont(hDC, m_oFont)
    pvLayoutGroupChips hDC, 0
    Call SelectObject(hDC, hPrevFont)
    Call ReleaseDC(picGrid.hWnd, hDC)
    For lIdx = 1 To m_oGroups.Count
        Set oItem = m_oGroups.Item(lIdx)
        With oItem.frChipRect
            If lX >= .Left And lX < .Right And lY >= .Top And lY < .Bottom And .Right > .Left Then
                Set oGroup = oItem
                pvGroupAtPoint = lIdx
                Exit For
            End If
        End With
    Next
End Function

Private Function pvBeginEdit(ByVal lPos As Long, ByVal lCol As Long, ByVal bSelectAll As Boolean, Optional ByVal lClickX As Long = -1, Optional ByVal lClickY As Long = -1) As Boolean
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
    If oCol.EditType <> jgexEditTextBox And oCol.EditType <> jgexEditCheckBox Then
        Exit Function
    End If
    '--- the client gets its veto before the editor appears
    Set oCancel = New JSRetBoolean
    RaiseEvent BeforeColEdit(oCol.Index, oCancel)
    If oCancel.Value Then
        Exit Function
    End If
    If Not pvCellRect(lPos, oCol.Index, lX, lY, lW) Then
        Exit Function
    End If
    If oCol.EditType = jgexEditCheckBox Then
        '--- a checkbox has no editor to show: the click is the edit, so the
        '--- value flips there and then and the client hears one Change.
        '--- Anywhere in the cell counts -- 064 clicks well left of the box and
        '--- still toggles, and a point a row above it toggles at 144dpi too
        Value(oCol.Index) = Not pvIsChecked(pvRowDataAt(lPos), oCol.Index)
        pvInvalidate SkipScroll:=True
        RaiseEvent Change
        pvBeginEdit = True
        Exit Function
    End If
    '--- the cell opens the box, so it has to be all the way inside the view
    '--- first: a click on a row or a column hanging off the edge scrolls it
    '--- in, and the click is carried along with it so the caret still lands
    '--- where it was put
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
    m_bInEditSetup = True
    m_bEditing = True
    m_lEditRow = lRowIndex
    m_lEditCol = oCol.Index
    m_sEditOldValue = pvCellText(pvRowDataAt(lPos), oCol.Index)
    '--- the editor is a native EDIT made for this cell and thrown away with
    '--- it: every style a column can ask for is fixed when the window is
    '--- created, so a fresh one per session is what lets the column decide
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
    '--- it sits one pixel inside the cell, so the row's selection colour and
    '--- the marquee running along it still show around the box
    '--- the editor's text has to land where the painted cell's would, and an
    '--- EDIT draws at the top of its client, so the window itself takes the
    '--- place the text centres into: the row keeps 19 pixels at either dpi
    '--- while the font grows, which is the whole of the difference
    uMetrics = FontTextMetrics(m_oFont)
    lTextTop = lY + (pvRowContentH(m_lRowHeight) - uMetrics.tmHeight) \ 2
    lEditH = uMetrics.tmHeight
    If oCol.WordWrap Then
        '--- a wrapping editor takes the whole cell, since it is the room it
        '--- has that decides how many lines the text breaks into
        lTextTop = lY + 1
        lEditH = pvRowContentH(m_lRowHeight) - 2
    End If
    '--- it ends where the painted cell's text is clipped, at lX + lW - 3, and
    '--- not a pixel further: two more would let "the quick brown fox jumps
    '--- over the lazy dog" break a word later in a wrapping cell, and the
    '--- click that opens the editor then lands on a character rather than
    '--- past the end of the last line
    m_hWndEdit = CreateWindowEx(0, StrPtr("EDIT"), 0, lStyle, lX + 1, lTextTop, lW - 4, lEditH, hWnd, 0, App.hInstance, ByVal 0&)
    If m_hWndEdit = 0 Then
        m_bEditing = False
        m_bInEditSetup = False
        Exit Function
    End If
    '--- through IFont, which is where the handle lives -- StdFont does not
    '--- carry hFont on its own interface
    Set pFont = m_oFont
    Call SendMessage(m_hWndEdit, WM_SETFONT, pFont.hFont, 1)
    '--- after the font, which resets it: the text keeps the same two pixel
    '--- margin a painted cell gives it, one of which the box itself is
    Call SendMessage(m_hWndEdit, EM_SETMARGINS, EC_LEFTMARGIN, 1)
    Call SendMessage(m_hWndEdit, EM_LIMITTEXT, oCol.MaxLength, 0)
    Call SendMessage(m_hWndEdit, WM_SETTEXT, 0, ByVal StrPtr(m_sEditOldValue))
    If bSelectAll Then
        Call SendMessage(m_hWndEdit, EM_SETSEL, 0, -1)
    ElseIf lClickY >= 0 Then
        '--- the click that opened the editor carries on into it: the caret
        '--- lands on the character under the point, which is what decides
        '--- where typing goes and, in a wrapping cell, which line shows.
        '--- Only the first line answers though -- the original hit-tests the
        '--- cell the way it does an unwrapped one, so a point below that line
        '--- is simply past the text and takes the caret to the end. Probed
        '--- against the original on a 3-line wrapping cell: clicking line 0
        '--- gives the character there at both dpi, while line 1 gives the end
        '--- at 96 (where the line is 13px and the point falls on it) and the
        '--- line 0 character at 120 (16px, where the same point does not)
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
    pvBeginEdit = True
End Function

Private Function pvEditText() As String
    Dim lLen            As Long

    If m_hWndEdit = 0 Then
        Exit Function
    End If
    lLen = SendMessage(m_hWndEdit, WM_GETTEXTLENGTH, 0, 0)
    pvEditText = String$(lLen + 1, 0)
    lLen = SendMessage(m_hWndEdit, WM_GETTEXT, lLen + 1, ByVal StrPtr(pvEditText))
    pvEditText = Left$(pvEditText, lLen)
End Function

Private Sub pvDestroyEditor()
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
Private Sub pvCommitRow()
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

Private Sub pvCancelRow()
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

Private Sub pvEndEdit(Optional ByVal bCancel As Boolean)
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
    pvDestroyEditor
    '--- committing runs the update trio the original raises in this order:
    '--- the cell first, then the row, with a repaint between the two halves
    If Not bCancel And sText <> m_sEditOldValue Then
        Set oCancel = New JSRetBoolean
        RaiseEvent BeforeColUpdate(lRowIndex, lCol, m_sEditOldValue, oCancel)
        If oCancel.Value Then
            GoTo QH
        End If
        If Not m_oRowData Is Nothing Then
            m_oRowData.frAllowUpdate = True
            m_oRowData.Value(lCol) = sText
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
            '--- the box grew to hold the staircase, so the bands below it did
            '--- move down. The staircase is measured whole and rounded once,
            '--- not accumulated a truncated half-chip at a time the way the
            '--- chips themselves step down: at a 25px chip the original's box
            '--- is 52 high (CLng of 37.5, to even) where per-level truncation
            '--- would give 51, and 19 and 31px chips agree with both at 42
            '--- and 60. The chip tops do truncate -- 9, 12 and 15 -- so the
            '--- two cannot share pvChipStagger
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
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim lSeen           As Long
    Dim oItem           As JSColumn

    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
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
    Dim lIdx            As Long
    Dim oGroup          As JSGroup

    '--- what the original documents client code used to write in the two
    '--- events: a grouped column flips the order of its group, any other one
    '--- becomes the only sort key, ascending unless it already was
    For lIdx = 1 To m_oGroups.Count
        Set oGroup = m_oGroups.Item(lIdx)
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

    lTopGbox = pvGroupByBoxHeight()
    lTopHdr = pvRowsTop()
    If lY < lTopGbox Then
        '--- group-by box: a chip stands for its column, so clicking one sorts
        '--- exactly as clicking that column's header does
        lPos = pvGroupAtPoint(lX, lY, oGroup)
        If Not oGroup Is Nothing Then
            RaiseEvent GroupByBoxHeaderClick(oGroup)
            If m_bAutomaticSort Then
                pvAutoSort m_oColumns.Item(oGroup.ColIndex)
            End If
        End If
    ElseIf lY < lTopHdr Then
        '--- column header band
        lPos = pvColAtX(lX, oCol)
        If Not oCol Is Nothing Then
            RaiseEvent ColumnHeaderClick(oCol)
            If m_bAutomaticSort Then
                pvAutoSort oCol
            End If
        End If
    ElseIf m_lRowHeight > 0 Then
        '--- data area cell
        lRow = m_lFirstItem + (lY - lTopHdr) \ m_lRowHeight
        lPos = pvColAtX(lX, oCol)
        If lRow >= 1 And lRow <= RowCount And lPos >= 1 Then
            pvNavigate lRow, lPos, lShift, (lShift And vbCtrlMask) <> 0
            '--- clicking a cell opens its editor, which is what BeforeColEdit
            '--- announces -- and it happens before the client sees MouseDown.
            '--- The click belongs to the editor from there on: the button is
            '--- still down over the grid, and without this the first move
            '--- would drag a selection out of the cell being edited
            '--- Ctrl+click is a selection gesture and nothing else: it picks
            '--- the row up or puts it down, and no editor opens under it
            If (lShift And vbCtrlMask) = 0 Then
                m_bClickOpenedEdit = pvBeginEdit(lRow, lPos, False, lX, lY)
            End If
        End If
    End If
End Sub

Private Sub pvOnKeyDown(ByVal lKeyCode As Long, ByVal lShift As Long)
    Select Case lKeyCode
    Case vbKeyDown
        If m_lRow < RowCount Then
            pvNavigate m_lRow + 1, m_lCol, lShift, False
        End If
    Case vbKeyUp
        If m_lRow > 1 Then
            pvNavigate m_lRow - 1, m_lCol, lShift, False
        End If
    Case vbKeyRight
        If m_lCol < pvVisibleColCount() Then
            Col = m_lCol + 1
        End If
    Case vbKeyLeft
        If m_lCol > 1 Then
            Col = m_lCol - 1
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
        '--- the grid only sees Escape once no editor holds the focus, so this
        '--- is the second press: the first closes the cell, this one drops
        '--- every column the row has buffered since it became current
        pvCancelRow
    End Select
End Sub

Private Sub pvSetRow(ByVal lValue As Long)
    Dim lLastRow        As Long

    '--- moves the current row without touching the selection: the row is
    '--- painted selected, so it still has to repaint
    If m_lRow <> lValue Then
        lLastRow = m_lRow
        m_lRow = lValue
        pvSyncRowData
        '--- the record the marquee is on, so a reprojection can put it back
        If m_pDataModel.RowIndex(m_lRow) > 0 Then
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
    pvEndEdit
    If bRowChanging Then
        pvCommitRow
    End If
    '--- the row being landed on is decorated as it becomes current, whether or
    '--- not anything was committed on the way out: probed with 075, a plain
    '--- row move raises it with nothing else around it. A column move does not
    '--- -- 067 has no RowFormat at all
    If bRowChanging Then
        RaiseEvent RowFormat(pvRowDataAt(lRow))
    End If
    '--- the selection lands on the new row before the move is announced,
    '--- which is the order the original raises the two events in
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
        m_lSelAnchor = lRow
        pvInvalidate SkipScroll:=True
        RaiseEvent SelectionChange
    ElseIf m_bMultiSelect And (lShift And vbShiftMask) <> 0 Then
        pvSetRangeSel m_lSelAnchor, lRow
    Else
        '--- re-selecting the row that already is the whole selection changes
        '--- nothing, and the original stays quiet about it -- a click that only
        '--- moves the column raises RowColChange and no more
        bSame = (m_oSelectedItems.Count = 1)
        If bSame Then
            bSame = (m_oSelectedItems.Item(1).RowPosition = lRow)
        End If
        m_oSelectedItems.Clear
        pvAddSel lRow
        m_lSelAnchor = lRow
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

    '--- the item remembers which row it is, not just where it sits, so a
    '--- re-sort can move it
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
        '--- while the thumb is being dragged the contents follow only with
        '--- ContinuousScroll, otherwise they wait for the button release
        If lCode = SB_THUMBTRACK And Not m_bContinuousScroll Then
            Exit Sub
        End If
        '--- the position rides in the message, as it does for any range that
        '--- fits 16 bits -- reading it back is only needed past that
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
    '--- what the bar carries is column numbers, written onto the window by
    '--- pvUpdateScrollBars, so a position out of the message is a column and
    '--- needs no mapping back
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
        '--- while the thumb is being dragged the columns follow only with
        '--- ContinuousScroll, otherwise they wait for the button release
        If lCode = SB_THUMBTRACK And Not m_bContinuousScroll Then
            Exit Sub
        End If
        '--- the position rides in the message, as it does for any range that
        '--- fits 16 bits -- reading it back is only needed past that
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
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"

    On Error GoTo EH
    pvInheritAmbientFont
    pvSubclass
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

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"

    '--- the navigator lives in the band, which is the outer control's own
    '--- client area -- picGrid covers everything above it
    On Error GoTo EH
    If m_bRecordNavigator Then
        pvPaintNavigator UserControl.hDC
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub picGrid_Paint()
    Const FUNC_NAME     As String = "picGrid_Paint"

    On Error GoTo EH
    pvPaint picGrid.hDC
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
    m_sCalendarTodayText = "Today"
    m_sCalendarNoneText = "None"
    m_sGroupByBoxInfoText = "Drag a column header here to group by that column."
    m_sRecordNavigatorString = "Record:|of"
    pvCreateDataModel
End Sub

Private Sub UserControl_Terminate()
    pvDestroyEditor
    pvUnsubclass
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
    If Not m_pDataModel Is Nothing Then
        m_pDataModel.Terminate
        Set m_pDataModel = Nothing
    End If
End Sub
