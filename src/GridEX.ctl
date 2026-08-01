VERSION 5.00
Begin VB.UserControl GridEX 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
   Begin VB.HScrollBar hsbGrid
      Height          =   255
      Left            =   0
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
Private m_hScrollH                  As Long
Private m_bHScroll                  As Boolean
Private m_bBand                     As Boolean
Private WithEvents m_oColumnHeaderFont As StdFont
Attribute m_oColumnHeaderFont.VB_VarHelpID = -1
Private WithEvents m_oFont          As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_bRowHeightSet             As Boolean
Private m_bScrollUpdating           As Boolean
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
Private m_pSubclassPic              As IUnknown
Private m_pSubclassCtl              As IUnknown
Private m_lSelAnchor                As Long
Private m_aRows()                   As UcsRowData
Private m_lRowsUBound               As Long
Private m_aOrder()                  As Long
Private m_lOrderCount               As Long
Private m_bSortDirty                As Boolean
Private m_bInSet                    As Boolean
Private m_bInOrdering               As Boolean
Private m_lSortKeys                 As Long
Private m_aSortCol()                As Integer
Private m_aSortDir()                As Long
Private m_aSortType()               As Long
Private m_aSortVals()               As Variant
Private m_aGroupRow()               As UcsGroupRow
Private m_lGroupRowsUBound          As Long
Private m_lGroupCols                As Long
Private m_aVisible()                As Long
Private m_lVisibleCount             As Long
Private m_bEditing                  As Boolean
Private m_lEditRow                  As Long
Private m_nEditCol                  As Integer
Private m_sEditOldValue             As String
Private m_bInEditSetup              As Boolean
Private m_hWndEdit                  As Long
Private m_pSubclassEdit             As IUnknown

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

Private Type UcsGroupRow
    Level                   As Integer
    Caption                 As String
    RecordCount             As Long
    Collapsed               As Boolean
    IconIndex               As Integer
    Prefixed                As Boolean
    Footer                  As Boolean
    FirstSlot               As Long
    LastSlot                As Long
    RowData                 As JSRowData
End Type

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
    pvEnsureOrder
    RowCount = m_lItemCount
    If m_lOrderCount > 0 Then
        RowCount = m_lVisibleCount
    End If
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
            pvInvalidate
        End If
    Else
        If pvIsRowSelected(RowPosition) Then
            m_oSelectedItems.RemoveRowPosition RowPosition
            RaiseEvent SelectionChange
            pvInvalidate
        End If
    End If
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
    Dim nMax            As Integer

    '--- setting a value past the visible columns selects the last one
    nMax = pvVisibleColCount()
    If nValue > nMax Then
        nValue = nMax
    End If
    If nValue < 1 And nMax > 0 Then
        nValue = 1
    End If
    If m_nCol <> nValue Then
        nLastCol = m_nCol
        m_nCol = nValue
        '--- the current cell is drawn out of the selected row it sits in, so
        '--- moving between columns is a repaint and not only an event
        pvInvalidate
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
    '--- rows appearing or going invalidates the order the same way a key does
    frSortChanged
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
    Dim nMax            As Integer

    '--- scrolling stops once the last column reaches the right edge, so a
    '--- value past that clamps -- the frozen block is not scrolled over and
    '--- takes its width out of the strip the rest has to fill
    nMax = pvVisibleColCount() - pvVisibleColsInWidth(pvScrollableWidth(), pvFrozenCount() + 1) + 1
    If nMax < 1 Then
        nMax = 1
    End If
    If nValue > nMax Then
        nValue = nMax
    End If
    If nValue < 1 Then
        nValue = 1
    End If
    If m_nLeftCol <> nValue Then
        m_nLeftCol = nValue
        '--- painting starts at the left-most column, so scrolling repaints
        pvInvalidate
        pvUpdateScrollBars
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
        RaiseEvent FirstItemChange
        pvInvalidate
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
    pvEnsureRow RowIndex
    RowBookmark = m_aRows(RowIndex).Bookmark
End Property

Public Property Let RowBookmark(ByVal RowIndex As Long, ByVal vntValue As Variant)
    pvEnsureRow RowIndex
    m_aRows(RowIndex).Bookmark = vntValue
End Property

Public Property Get Row() As Long
    pvEnsureOrder
Attribute Row.VB_Description = "Returns/sets the current row/card position."
    Row = m_lRow
End Property

Public Property Let Row(ByVal lValue As Long)
    Dim lLastRow        As Long

    If m_lRow <> lValue Then
        pvSetRow lValue
        '--- an assignment from outside collapses the selection onto the new
        '--- row, silently: navigation and drag go through pvSetRow instead
        '--- and apply their own selection, keeping SelectionChange ordering
        m_oSelectedItems.Clear
        pvAddSel m_lRow
        m_lSelAnchor = m_lRow
        pvInvalidate
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
    pvEnsureOrder
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
    Dim lSlot           As Long

    pvEnsureOrder
    lSlot = pvGroupSlot(RowPosition)
    If lSlot > 0 Then
        RowExpanded = Not m_aGroupRow(lSlot).Collapsed
    End If
End Property

Public Property Let RowExpanded(ByVal RowPosition As Long, ByVal bValue As Boolean)
    Dim lSlot           As Long

    pvEnsureOrder
    lSlot = pvGroupSlot(RowPosition)
    If lSlot > 0 Then
        m_aGroupRow(lSlot).Collapsed = Not bValue
        pvRecalcVisible
    End If
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

Friend Property Get frRowColCount() As Integer
    frRowColCount = m_oColumns.Count
End Property

Friend Property Get frRowValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Variant
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    frRowValue = m_aRows(lRowIndex).Cells(nColIndex).Value
End Property

Friend Property Let frRowValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal vntValue As Variant)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).Value = vntValue
End Property

Friend Property Get frRowIconIndex(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Integer
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    frRowIconIndex = m_aRows(lRowIndex).Cells(nColIndex).IconIndex
End Property

Friend Property Let frRowIconIndex(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal nValue As Integer)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).IconIndex = nValue
End Property

Friend Property Get frRowDisplayValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    frRowDisplayValue = m_aRows(lRowIndex).Cells(nColIndex).DisplayValue
End Property

Friend Property Let frRowDisplayValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal sValue As String)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).DisplayValue = sValue
End Property

Friend Property Get frRowCellStyle(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    frRowCellStyle = m_aRows(lRowIndex).Cells(nColIndex).CellStyle
End Property

Friend Property Let frRowCellStyle(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal sValue As String)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    m_aRows(lRowIndex).Cells(nColIndex).CellStyle = sValue
End Property

Friend Property Get frRowCellPicture(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Picture
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    Set frRowCellPicture = m_aRows(lRowIndex).Cells(nColIndex).CellPicture
End Property

Friend Property Set frRowCellPicture(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal oValue As Picture)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvFetchRow lRowIndex
    Set m_aRows(lRowIndex).Cells(nColIndex).CellPicture = oValue
End Property

Friend Property Get frRowBookmark(ByVal lRowIndex As Long) As Variant
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowBookmark = m_aRows(lRowIndex).Bookmark
End Property

Friend Property Get frRowIndex(ByVal lRowIndex As Long) As Long
    '--- a group row is not one of the client app's records, so it has no
    '--- index to report, exactly as RowIndex answers for its position
    If lRowIndex > 0 Then
        frRowIndex = lRowIndex
    End If
End Property

Friend Property Get frRowType(ByVal lRowIndex As Long) As jgexRowTypeConstants
    If lRowIndex < 0 Then
        If m_aGroupRow(-lRowIndex).Footer Then
            frRowType = jgexRowTypeGroupFooter
        Else
            frRowType = jgexRowTypeGroupHeader
        End If
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowType = m_aRows(lRowIndex).RowType
End Property

Friend Property Get frRowGroupLevel(ByVal lRowIndex As Long) As Integer
    If lRowIndex < 0 Then
        frRowGroupLevel = m_aGroupRow(-lRowIndex).Level
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowGroupLevel = m_aRows(lRowIndex).GroupLevel
End Property

Friend Property Get frRowRecordCount(ByVal lRowIndex As Long) As Long
    If lRowIndex < 0 Then
        frRowRecordCount = m_aGroupRow(-lRowIndex).RecordCount
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowRecordCount = m_aRows(lRowIndex).RecordCount
End Property

Friend Property Get frRowHeight(ByVal lRowIndex As Long) As Long
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowHeight = m_aRows(lRowIndex).RowHeight
End Property

Friend Property Let frRowHeight(ByVal lRowIndex As Long, ByVal lValue As Long)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).RowHeight = lValue
End Property

Friend Property Get frRowStyle(ByVal lRowIndex As Long) As String
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowStyle = m_aRows(lRowIndex).RowStyle
End Property

Friend Property Let frRowStyle(ByVal lRowIndex As Long, ByVal sValue As String)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).RowStyle = sValue
End Property

Friend Property Get frRowGroupCaption(ByVal lRowIndex As Long) As String
    If lRowIndex < 0 Then
        frRowGroupCaption = m_aGroupRow(-lRowIndex).Caption
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowGroupCaption = m_aRows(lRowIndex).GroupCaption
End Property

Friend Property Let frRowGroupCaption(ByVal lRowIndex As Long, ByVal sValue As String)
    If lRowIndex < 0 Then
        '--- the client app can retitle a group row, which is why the caption
        '--- is stored rather than recomputed at paint time
        m_aGroupRow(-lRowIndex).Caption = sValue
        pvInvalidate
        Exit Property
    End If
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).GroupCaption = sValue
End Property

Friend Property Get frRowGroupIconIndex(ByVal lRowIndex As Long) As Integer
    If lRowIndex < 0 Then
        frRowGroupIconIndex = m_aGroupRow(-lRowIndex).IconIndex
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowGroupIconIndex = m_aRows(lRowIndex).GroupIconIndex
End Property

Friend Property Let frRowGroupIconIndex(ByVal lRowIndex As Long, ByVal nValue As Integer)
    If lRowIndex < 0 Then
        m_aGroupRow(-lRowIndex).IconIndex = nValue
        pvInvalidate
        Exit Property
    End If
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).GroupIconIndex = nValue
End Property

Friend Property Get frRowPreviewRowVisible(ByVal lRowIndex As Long) As Boolean
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    frRowPreviewRowVisible = m_aRows(lRowIndex).PreviewRowVisible
End Property

Friend Property Let frRowPreviewRowVisible(ByVal lRowIndex As Long, ByVal bValue As Boolean)
    If lRowIndex < 0 Then
        Exit Property
    End If
    pvEnsureRow lRowIndex
    m_aRows(lRowIndex).PreviewRowVisible = bValue
End Property

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
    pvResetRows False
    pvApplyHoldSort HoldSortSettings
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
    pvSetAllCollapsed True
End Sub

Public Sub RefreshSort()
Attribute RefreshSort.VB_Description = "Forces the re-sort of the records."
    '--- the collections mark the order stale as they are edited, so this only
    '--- has to bring the rebuild forward from the next paint
    frSortChanged
    pvEnsureOrder
    pvInvalidate
End Sub

Public Function RowIndex(ByVal RowPosition As Long) As Long
Attribute RowIndex.VB_Description = "Returns the original index of a row."
    '--- a sort moves rows around, so this is where the client app asks what
    '--- it is actually looking at; group rows will return 0 once grouping lands
    If RowPosition >= 1 And RowPosition <= RowCount Then
        If Not IsGroupItem(RowPosition) Then
            RowIndex = pvDataRow(RowPosition)
        End If
    End If
End Function

Public Sub RefreshGroups(Optional ByVal AllCollapsed As Boolean)
Attribute RefreshGroups.VB_Description = "Forces recalculation of groups."
    '--- the groups are rebuilt from scratch, which resets every expand box
    '--- to DefaultGroupMode unless this call overrides it
    frSortChanged
    pvEnsureOrder
    If AllCollapsed Then
        pvSetAllCollapsed True
    Else
        pvInvalidate
    End If
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
    pvResetRows True
    pvApplyHoldSort HoldSortSettings
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
    '--- a bind starts with the whole row selected rather than a cell in it,
    '--- which is what Col = 0 means -- the original's first RowColChange
    '--- reports LastCol=0 for exactly that reason
    m_nCol = 0
    m_bInSet = False
    pvInvalidate
End Sub

Public Function IsGroupItem(ByVal Row As Long) As Boolean
Attribute IsGroupItem.VB_Description = "Returns True if the specified row is a group row."
    pvEnsureOrder
    IsGroupItem = pvIsGroupRow(Row)
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
    pvSetAllCollapsed False
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
    Dim lSlot           As Long

    pvEnsureOrder
    lSlot = pvGroupSlot(RowPosition)
    If lSlot > 0 Then
        GroupRowLevel = m_aGroupRow(lSlot).Level
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
    Dim lSlot           As Long

    '--- a group row keeps its wrapper on the group itself and is addressed
    '--- by the negative of its slot, so every frRow accessor can tell the
    '--- two apart; a record answers on the data row its position holds
    pvEnsureOrder
    lSlot = pvGroupSlot(RowPosition)
    If lSlot > 0 Then
        With m_aGroupRow(lSlot)
            If .RowData Is Nothing Then
                Set .RowData = New JSRowData
                .RowData.frInit Me, -lSlot
            End If
            Set GetRowData = .RowData
        End With
        Exit Function
    End If
    If m_lOrderCount > 0 Then
        RowPosition = pvDataRow(RowPosition)
    End If
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
' Methods
'=========================================================================

Friend Sub frSortChanged(Optional ByVal bInvalidate As Boolean)
    m_bSortDirty = True
    If bInvalidate And Not m_bInSet Then
        pvInvalidate
    End If
End Sub

Friend Property Get frDefaultColumnWidthPx() As Long
    frDefaultColumnWidthPx = m_lDefaultColumnWidth
End Property

Friend Function frRowSubTotal(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal eFunc As jgexAggregateFunctionConstants) As Variant
    '--- only a group row aggregates, over the records it covers
    If lRowIndex < 0 Then
        With m_aGroupRow(-lRowIndex)
            frRowSubTotal = pvAggregate(.FirstSlot, .LastSlot, nColIndex, eFunc)
        End With
    End If
End Function

Friend Function frColIsGrouped(ByVal nColIndex As Integer) As Boolean
    Dim lIdx            As Long

    For lIdx = 1 To m_oGroups.Count
        If m_oGroups.Item(lIdx).ColIndex = nColIndex Then
            frColIsGrouped = True
            Exit For
        End If
    Next
End Function

Private Function pvSlot(ByVal lPos As Long) As Long
    '--- a display position addresses one of the rows currently on show,
    '--- which is a slot in the full order the sort built
    If m_lVisibleCount > 0 And lPos >= 1 And lPos <= m_lVisibleCount Then
        pvSlot = m_aVisible(lPos)
    End If
End Function

Private Function pvSlotPos(ByVal lSlot As Long) As Long
    Dim lIdx            As Long

    '--- and back: a slot inside a collapsed group has no display position
    '--- of its own, so it answers with the group row that hides it
    For lIdx = m_lVisibleCount To 1 Step -1
        If m_aVisible(lIdx) <= lSlot Then
            pvSlotPos = lIdx
            Exit Function
        End If
    Next
End Function

Private Function pvDataRow(ByVal lPos As Long) As Long
    Dim lSlot           As Long

    pvEnsureOrder
    lSlot = pvSlot(lPos)
    If lSlot > 0 Then
        pvDataRow = m_aOrder(lSlot)
    ElseIf m_lOrderCount = 0 Then
        pvDataRow = lPos
    End If
End Function

Private Function pvDataRowPos(ByVal lRowIndex As Long) As Long
    Dim lIdx            As Long

    '--- and back: the current row follows its data across a re-sort, and
    '--- lands on the group row when the data it was on is collapsed away
    pvDataRowPos = lRowIndex
    If m_lOrderCount > 0 Then
        For lIdx = 1 To m_lOrderCount
            If m_aOrder(lIdx) = lRowIndex Then
                pvDataRowPos = pvSlotPos(lIdx)
                Exit Function
            End If
        Next
    End If
End Function

Private Sub pvEnsureOrder()
    If Not m_bSortDirty Or m_bInOrdering Then
        Exit Sub
    End If
    m_bSortDirty = False
    pvBuildOrder
End Sub

Private Sub pvBuildOrder()
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim aTemp()         As Long
    Dim aRows()         As Long
    Dim oItem           As JSSelectedItem

    m_bInOrdering = True
    lRow = pvDataRow(m_lRow)
    If (m_oSortKeys.Count = 0 And m_oGroups.Count = 0) Or m_lItemCount = 0 Then
        Erase m_aOrder
        pvEraseGroupRows
        m_lOrderCount = 0
    Else
        pvDecorate
        ReDim aRows(1 To m_lItemCount) As Long
        For lIdx = 1 To m_lItemCount
            aRows(lIdx) = lIdx
        Next
        ReDim aTemp(1 To m_lItemCount) As Long
        pvMergeSort aRows, aTemp, 1, m_lItemCount
        If m_lGroupCols = 0 Then
            m_aOrder = aRows
            m_lOrderCount = m_lItemCount
            pvEraseGroupRows
        Else
            pvBuildGroupRows aRows
        End If
        '--- the keys are only worth their memory while sorting
        Erase m_aSortVals
        m_lSortKeys = 0
    End If
    pvBuildVisible
    '--- the marquee and the selection stay on the rows they were on,
    '--- wherever those land
    m_lRow = pvDataRowPos(lRow)
    For Each oItem In m_oSelectedItems
        oItem.frSetPosition pvDataRowPos(oItem.RowIndex)
    Next
    m_bInOrdering = False
End Sub

Private Sub pvBuildGroupRows(aRows() As Long)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lPos            As Long
    Dim lFrom           As Long
    Dim lRoom           As Long
    Dim aStart()        As Long

    '--- the sorted rows are walked once and a group row is emitted wherever
    '--- a group key changes, so the display is group row, its records, the
    '--- next group row and so on
    '--- room for a header and a footer per level per row, which is what a
    '--- corpus of all-distinct keys would come to
    lRoom = m_lItemCount + 2 * m_lItemCount * m_lGroupCols
    ReDim m_aOrder(1 To lRoom) As Long
    '--- a plain ReDim drops the wrappers the previous order handed out, so
    '--- they get detached first, exactly as an Erase has to
    pvEraseGroupRows
    ReDim m_aGroupRow(1 To lRoom) As UcsGroupRow
    m_lGroupRowsUBound = lRoom
    ReDim aStart(1 To m_lGroupCols) As Long
    For lIdx = 1 To m_lItemCount
        '--- the first level whose key changed opens a new group row here and
        '--- at every level below it, so a change high up restarts the ones
        '--- nested inside it
        lFrom = 0
        For lJdx = 1 To m_lGroupCols
            If lIdx = 1 Then
                lFrom = 1
                Exit For
            End If
            If pvCompareValues(m_aSortVals(lJdx, aRows(lIdx)), m_aSortVals(lJdx, aRows(lIdx - 1)), _
                    m_aSortType(lJdx)) <> 0 Then
                lFrom = lJdx
                Exit For
            End If
        Next
        If lFrom > 0 Then
            '--- whatever this change closes gets its footer before the next
            '--- level opens, innermost first
            pvCloseGroups lFrom, lPos, aStart
            For lJdx = lFrom To m_lGroupCols
                lPos = lPos + 1
                m_aOrder(lPos) = 0
                With m_aGroupRow(lPos)
                    .Level = lJdx
                    .Caption = pvGroupCaption(aRows(lIdx), m_aSortCol(lJdx))
                    .Prefixed = (LenB(m_oColumns.Item(m_aSortCol(lJdx)).GroupPrefix) <> 0)
                    .RecordCount = 0
                    .Collapsed = (m_eDefaultGroupMode = jgexDGMCollapsed)
                End With
                aStart(lJdx) = lPos
            Next
        End If
        For lJdx = 1 To m_lGroupCols
            m_aGroupRow(aStart(lJdx)).RecordCount = m_aGroupRow(aStart(lJdx)).RecordCount + 1
        Next
        lPos = lPos + 1
        m_aOrder(lPos) = aRows(lIdx)
    Next
    '--- the last group of every level is closed by running out of rows
    pvCloseGroups 1, lPos, aStart
    m_lOrderCount = lPos
    ReDim Preserve m_aOrder(1 To lPos) As Long
    ReDim Preserve m_aGroupRow(1 To lPos) As UcsGroupRow
    m_lGroupRowsUBound = lPos
End Sub

Private Sub pvCloseGroups(ByVal lFromLevel As Long, lPos As Long, aStart() As Long)
    Dim lJdx            As Long
    Dim lHdr            As Long

    '--- closing a level records the span of rows it covered -- what the
    '--- aggregates and GetSubTotal are computed over -- and lays down its
    '--- footer row when the control is showing them
    For lJdx = m_lGroupCols To lFromLevel Step -1
        lHdr = aStart(lJdx)
        If lHdr > 0 Then
            m_aGroupRow(lHdr).FirstSlot = lHdr + 1
            m_aGroupRow(lHdr).LastSlot = lPos
            If m_eGroupFooterStyle <> jgexNoGroupFooter Then
                lPos = lPos + 1
                m_aOrder(lPos) = 0
                With m_aGroupRow(lPos)
                    .Level = lJdx
                    .Caption = m_aGroupRow(lHdr).Caption
                    .Prefixed = m_aGroupRow(lHdr).Prefixed
                    .RecordCount = m_aGroupRow(lHdr).RecordCount
                    .Collapsed = m_aGroupRow(lHdr).Collapsed
                    .Footer = True
                    .FirstSlot = m_aGroupRow(lHdr).FirstSlot
                    .LastSlot = m_aGroupRow(lHdr).LastSlot
                End With
            End If
        End If
    Next
End Sub

Private Sub pvEraseDataRows()
    Dim lIdx            As Long

    For lIdx = 1 To m_lRowsUBound
        If Not m_aRows(lIdx).RowData Is Nothing Then
            m_aRows(lIdx).RowData.frTerminate
        End If
    Next
    Erase m_aRows
    m_lRowsUBound = 0
End Sub

Private Sub pvEraseGroupRows()
    Dim lIdx            As Long

    For lIdx = 1 To m_lGroupRowsUBound
        If Not m_aGroupRow(lIdx).RowData Is Nothing Then
            m_aGroupRow(lIdx).RowData.frTerminate
        End If
    Next
    Erase m_aGroupRow
    m_lGroupRowsUBound = 0
End Sub

Private Sub pvBuildVisible()
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lHidden         As Long

    '--- everything under a collapsed group row drops out of the display
    '--- until a group row at that level or above shows up again
    m_lVisibleCount = 0
    Erase m_aVisible
    If m_lOrderCount = 0 Then
        Exit Sub
    End If
    ReDim m_aVisible(1 To m_lOrderCount) As Long
    For lIdx = 1 To m_lOrderCount
        If lHidden > 0 And m_aOrder(lIdx) = 0 Then
            '--- a collapsed group hides its own footer along with its records,
            '--- so only a header at that level -- or anything from a level
            '--- further out, footers included -- brings the display back
            If m_aGroupRow(lIdx).Level < lHidden Then
                lHidden = 0
            ElseIf m_aGroupRow(lIdx).Level = lHidden And Not m_aGroupRow(lIdx).Footer Then
                lHidden = 0
            End If
        End If
        If lHidden = 0 Then
            lPos = lPos + 1
            m_aVisible(lPos) = lIdx
            If m_aOrder(lIdx) = 0 Then
                If m_aGroupRow(lIdx).Collapsed Then
                    lHidden = m_aGroupRow(lIdx).Level
                End If
            End If
        End If
    Next
    m_lVisibleCount = lPos
    ReDim Preserve m_aVisible(1 To lPos) As Long
End Sub

Private Sub pvRecalcVisible()
    Dim lSlot           As Long
    Dim oItem           As JSSelectedItem

    '--- an expand or a collapse only reprojects: the sort order underneath
    '--- stays exactly as it was
    lSlot = pvSlot(m_lRow)
    pvBuildVisible
    If lSlot > 0 Then
        m_lRow = pvSlotPos(lSlot)
    End If
    For Each oItem In m_oSelectedItems
        oItem.frSetPosition pvDataRowPos(oItem.RowIndex)
    Next
    pvInvalidate
End Sub

Private Sub pvSetAllCollapsed(ByVal bValue As Boolean)
    Dim lIdx            As Long

    pvEnsureOrder
    For lIdx = 1 To m_lOrderCount
        If m_aOrder(lIdx) = 0 Then
            m_aGroupRow(lIdx).Collapsed = bValue
        End If
    Next
    pvRecalcVisible
End Sub

Private Function pvIsGroupRow(ByVal lPos As Long) As Boolean
    Dim lSlot           As Long

    lSlot = pvSlot(lPos)
    If lSlot > 0 Then
        pvIsGroupRow = (m_aOrder(lSlot) = 0)
    End If
End Function

Private Function pvGroupSlot(ByVal lPos As Long) As Long
    Dim lSlot           As Long

    '--- non-zero only for a display position that holds a group row
    lSlot = pvSlot(lPos)
    If lSlot > 0 Then
        If m_aOrder(lSlot) = 0 Then
            pvGroupSlot = lSlot
        End If
    End If
End Function

Private Sub pvDecorate()
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim oKey            As JSSortKey
    Dim oGroup          As JSGroup
    Dim nCol            As Integer
    Dim nOrder          As Integer
    Dim vValue          As Variant

    '--- every key of every row read once, up front. Reaching through
    '--- SortKeys/Columns and the row cache from inside the comparator costs
    '--- eight COM calls a comparison, which is what actually dominates a
    '--- sort -- the comparator below only touches plain arrays
    '--- grouping sorts too: the group columns come first, in order, and the
    '--- sort keys refine within a group
    m_lGroupCols = m_oGroups.Count
    m_lSortKeys = m_lGroupCols + m_oSortKeys.Count
    ReDim m_aSortCol(1 To m_lSortKeys) As Integer
    ReDim m_aSortDir(1 To m_lSortKeys) As Long
    ReDim m_aSortType(1 To m_lSortKeys) As Long
    For lIdx = 1 To m_lSortKeys
        If lIdx <= m_lGroupCols Then
            Set oKey = Nothing
            Set oGroup = m_oGroups.Item(lIdx)
            nCol = oGroup.ColIndex
            nOrder = oGroup.SortOrder
        Else
            Set oKey = m_oSortKeys.Item(lIdx - m_lGroupCols)
            nCol = oKey.ColIndex
            nOrder = oKey.SortOrder
        End If
        If nCol >= 1 And nCol <= m_oColumns.Count Then
            m_aSortCol(lIdx) = nCol
            m_aSortType(lIdx) = m_oColumns.Item(nCol).SortType
        End If
        '--- the enum is +1/-1 already, so the direction multiplies straight
        '--- into the comparison result
        m_aSortDir(lIdx) = IIf(nOrder = jgexSortDescending, -1, 1)
    Next
    ReDim m_aSortVals(1 To m_lSortKeys, 1 To m_lItemCount) As Variant
    For lRow = 1 To m_lItemCount
        '--- sorting needs every row's key, so the lazy fetch is forced here
        pvFetchRow lRow
        For lIdx = 1 To m_lSortKeys
            If m_aSortCol(lIdx) > 0 Then
                vValue = frRowValue(lRow, m_aSortCol(lIdx))
                '--- an object in a cell has no ordering, so it sorts blank
                If Not IsObject(vValue) Then
                    m_aSortVals(lIdx, lRow) = vValue
                End If
            End If
        Next
    Next
End Sub

Private Sub pvInsertionSort(aIdx() As Long, ByVal lFirst As Long, ByVal lLast As Long)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lRow            As Long

    '--- stable as long as equal elements stop the shift, which is what the
    '--- <= 0 does here
    For lIdx = lFirst + 1 To lLast
        lRow = aIdx(lIdx)
        For lJdx = lIdx - 1 To lFirst Step -1
            If pvCompareRows(aIdx(lJdx), lRow) <= 0 Then
                Exit For
            End If
            aIdx(lJdx + 1) = aIdx(lJdx)
        Next
        aIdx(lJdx + 1) = lRow
    Next
End Sub

Private Sub pvMergeSort(aIdx() As Long, aTemp() As Long, ByVal lFirst As Long, ByVal lLast As Long)
    Const MIN_RUN       As Long = 16
    Dim lMid            As Long
    Dim lLeft           As Long
    Dim lRight          As Long
    Dim lIdx            As Long

    '--- merge sort because it is stable: rows with equal keys keep the order
    '--- the client app supplied them in
    If lLast - lFirst < MIN_RUN Then
        '--- insertion sort wins on short runs and ends the recursion early
        pvInsertionSort aIdx, lFirst, lLast
        Exit Sub
    End If
    lMid = (lFirst + lLast) \ 2
    pvMergeSort aIdx, aTemp, lFirst, lMid
    pvMergeSort aIdx, aTemp, lMid + 1, lLast
    '--- already in order across the seam, which is the common case for data
    '--- that arrives sorted: nothing to merge
    If pvCompareRows(aIdx(lMid), aIdx(lMid + 1)) <= 0 Then
        Exit Sub
    End If
    lLeft = lFirst
    lRight = lMid + 1
    For lIdx = lFirst To lLast
        If lLeft > lMid Then
            aTemp(lIdx) = aIdx(lRight)
            lRight = lRight + 1
        ElseIf lRight > lLast Then
            aTemp(lIdx) = aIdx(lLeft)
            lLeft = lLeft + 1
        ElseIf pvCompareRows(aIdx(lRight), aIdx(lLeft)) < 0 Then
            aTemp(lIdx) = aIdx(lRight)
            lRight = lRight + 1
        Else
            aTemp(lIdx) = aIdx(lLeft)
            lLeft = lLeft + 1
        End If
    Next
    For lIdx = lFirst To lLast
        aIdx(lIdx) = aTemp(lIdx)
    Next
End Sub

Private Function pvCompareRows(ByVal lRow1 As Long, ByVal lRow2 As Long) As Long
    Dim lIdx            As Long

    For lIdx = 1 To m_lSortKeys
        pvCompareRows = pvCompareValues(m_aSortVals(lIdx, lRow1), m_aSortVals(lIdx, lRow2), _
            m_aSortType(lIdx)) * m_aSortDir(lIdx)
        If pvCompareRows <> 0 Then
            Exit Function
        End If
    Next
End Function

Private Function pvCompareValues(vValue1 As Variant, vValue2 As Variant, ByVal eSortType As jgexSortTypeConstants) As Long
    Dim bEmpty1         As Boolean
    Dim bEmpty2         As Boolean
    Dim dbl1            As Double
    Dim dbl2            As Double

    '--- an empty cell sorts before anything, whatever the type
    bEmpty1 = pvIsBlank(vValue1)
    bEmpty2 = pvIsBlank(vValue2)
    If bEmpty1 Or bEmpty2 Then
        pvCompareValues = -bEmpty2 - -bEmpty1
        Exit Function
    End If
    Select Case eSortType
    Case jgexSortTypeNumeric, jgexSortTypeDate, jgexSortTypeDateTime, jgexSortTypeTime
        dbl1 = C2Dbl(vValue1)
        dbl2 = C2Dbl(vValue2)
        If dbl1 < dbl2 Then
            pvCompareValues = -1
        ElseIf dbl1 > dbl2 Then
            pvCompareValues = 1
        End If
    Case Else
        pvCompareValues = StrComp(CStr(vValue1), CStr(vValue2), vbTextCompare)
    End Select
End Function

Private Function pvIsBlank(vValue As Variant) As Boolean
    Select Case VarType(vValue)
    Case vbEmpty, vbNull, vbError
        pvIsBlank = True
    Case vbString
        pvIsBlank = (LenB(vValue) = 0)
    End Select
End Function

Private Sub pvEnsureRow(ByVal lRowIndex As Long)
    Dim lIdx            As Long

    If lRowIndex > m_lRowsUBound Then
        ReDim Preserve m_aRows(1 To lRowIndex) As UcsRowData
        For lIdx = m_lRowsUBound + 1 To lRowIndex
            With m_aRows(lIdx)
                .RowHeight = ToTwips(m_lRowHeight)
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
            .RowHeight = ToTwips(m_lRowHeight)
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
    '--- the values the order was built from are gone, so it is stale
    frSortChanged
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

'--- painting

Private Sub pvPaint(ByVal hDC As Long)
    Dim lY              As Long

    '--- a sort the caller set up since the last paint takes effect here
    pvEnsureOrder
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
    lTotalH = lBoxH + 10
    If m_oGroups.Count > 1 Then
        '--- the box grows to hold the staircase of chips
        lTotalH = lTotalH + (m_oGroups.Count - 1) * pvChipStagger()
    End If
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
                lElbowTop = lTop + pvChipStagger() + uMetrics.tmHeight
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
    Dim nIdx            As Integer
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
        If pvColSortOrder(oCol) <> 0 Then
            uMetrics = FontTextMetrics(m_oColumnHeaderFont, hDC)
            '--- the arrow centres on the caption rather than sitting on its
            '--- baseline: the two coincide at the default font, which is why
            '--- 16, 19 and 25 pixel fonts were needed to tell them apart
            pvPaintSortGlyph hDC, lX + 2 + pvTextWidth(hDC, oCol.Caption) + 4, _
                lY + (lHdrH - uMetrics.tmHeight + 1) \ 2 + uMetrics.tmHeight \ 2 + 4, pvColSortOrder(oCol)
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

Private Function pvColSortOrder(oCol As JSColumn) As jgexSortOrderConstants
    Dim lIdx            As Long
    Dim oKey            As JSSortKey
    Dim oGroup          As JSGroup

    '--- grouping sorts by the column too, and the original marks the header
    '--- with the same arrow an explicit sort key gets
    For lIdx = 1 To m_oGroups.Count
        Set oGroup = m_oGroups.Item(lIdx)
        If oGroup.ColIndex = oCol.Index Then
            pvColSortOrder = oGroup.SortOrder
            Exit Function
        End If
    Next
    For lIdx = 1 To m_oSortKeys.Count
        Set oKey = m_oSortKeys.Item(lIdx)
        If oKey.ColIndex = oCol.Index Then
            pvColSortOrder = oKey.SortOrder
            Exit Function
        End If
    Next
End Function

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
    Dim nIdx            As Integer
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
        If m_lGroupCols > 0 And lPainted > 0 Then
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
                If m_lGroupCols > 0 Then
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
    If m_bBand Then
        pvLine hDC, 0, picGrid.ScaleHeight - 1, picGrid.ScaleWidth, picGrid.ScaleHeight - 1, m_clrBackColorHeader, PS_SOLID
    End If
End Sub

Private Sub pvPaintDataRow(ByVal hDC As Long, ByVal lRow As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal lHdrW As Long, ByVal lTotalW As Long)
    Dim bSelected       As Boolean
    Dim clrBack         As OLE_COLOR
    Dim clrText         As OLE_COLOR
    Dim lFillL          As Long
    Dim lFillR          As Long
    Dim clrCellBack     As OLE_COLOR
    Dim clrCellText     As OLE_COLOR
    Dim nPos            As Integer
    Dim lX              As Long
    Dim nIdx            As Integer
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
    nPos = 0
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oCol = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oCol.Visible Then
            nPos = nPos + 1
            lW = pvColWidth(oCol)
            '--- inside a selected row the current cell keeps the plain colors:
            '--- Col = 0 selects the whole row and nothing is singled out, but
            '--- once a cell is current the original lifts it out of the block
            If bSelected And m_nCol = nPos And lRow = m_lRow Then
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
                    pvIsChecked(pvDataRow(lRow), oCol.Index), (bSelected And m_nCol = nPos And lRow = m_lRow)
                sText = vbNullString
            Else
                sText = pvCellText(pvDataRow(lRow), oCol.Index)
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
    Dim nIdx            As Integer
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
    Dim lSlot           As Long
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
    lSlot = pvSlot(lPos)
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
        If m_aGroupRow(lSlot).Footer Then
            lLineLeft = pvBlockLeft() - 1
        End If
        pvLine hDC, lLineLeft, lRowTop - 1, lRight, lRowTop - 1, m_clrGridLinesColor, pvPenStyle()
    End If
    lBoxLeft = pvRowHeaderWidth() + lIndent + (GROUP_INDENT_W - GROUP_BOX_W) \ 2
    lBoxTop = lRowTop + (lRowH - 1 - GROUP_BOX_W) \ 2
    '--- a footer carries no expand box: it closes the group rather than
    '--- opening it, and under the totals style it reads across the columns
    '--- instead of carrying the caption
    If m_aGroupRow(lSlot).Footer Then
        If m_eGroupFooterStyle = jgexTotalsGroupFooter Then
            pvPaintGroupTotals hDC, lSlot, lRowTop, lRowH, clrBack, clrText
            Exit Sub
        End If
    Else
        pvPaintGroupBox hDC, lBoxLeft, lBoxTop, Not m_aGroupRow(lSlot).Collapsed
    End If
    uMetrics = FontTextMetrics(m_oFont, hDC)
    lTextTop = lRowTop + (lRowH - 1 - uMetrics.tmHeight) \ 2
    '--- the caption is drawn a space past the expand box, plus the two pixel
    '--- margin text keeps everywhere else -- this tracks the font at any dpi
    '--- the caption sits two pixels past the expand box behind a leading
    '--- space, which a prefix supplies itself -- so a prefixed level starts
    '--- one space earlier and its value still lands where a plain one does
    lTextLeft = lBoxLeft + GROUP_BOX_W + 2
    If Not m_aGroupRow(lSlot).Prefixed Then
        lTextLeft = lTextLeft + pvTextWidth(hDC, " ")
    End If
    pvDrawText hDC, m_aGroupRow(lSlot).Caption, lTextLeft, lTextTop, lRight, lTextTop + uMetrics.tmHeight, _
        clrText, clrBack, jgexAlignLeft, lTextLeft, lRight
End Sub

Private Sub pvPaintGroupTotals(ByVal hDC As Long, ByVal lSlot As Long, ByVal lRowTop As Long, ByVal lRowH As Long, ByVal clrBack As OLE_COLOR, ByVal clrText As OLE_COLOR)
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
                With m_aGroupRow(lSlot)
                    vValue = pvAggregate(.FirstSlot, .LastSlot, oCol.Index, oCol.AggregateFunction)
                End With
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
        pvRowIndent = (m_aGroupRow(pvSlot(lPos)).Level - 1) * GROUP_INDENT_W
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
    pvGroupIndent = m_lGroupCols * GROUP_INDENT_W
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

Private Function pvAggregate(ByVal lFirst As Long, ByVal lLast As Long, ByVal nColIndex As Integer, ByVal eFunc As jgexAggregateFunctionConstants) As Variant
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim lCount          As Long
    Dim dblSum          As Double
    Dim dblSquares      As Double
    Dim dblValue        As Double
    Dim vValue          As Variant
    Dim vMin            As Variant
    Dim vMax            As Variant

    '--- one walk over the records of a group serves every function: the
    '--- counts take any value, the rest only the ones that are numbers
    For lIdx = lFirst To lLast
        If lIdx >= 1 And lIdx <= m_lOrderCount Then
            lRow = m_aOrder(lIdx)
            '--- nested group rows and footers sit in the span too, and are
            '--- not records of it
            If lRow > 0 Then
                If eFunc = jgexCount Then
                    lCount = lCount + 1
                Else
                    vValue = frRowValue(lRow, nColIndex)
                    If Not pvIsBlank(vValue) And Not IsObject(vValue) Then
                        lCount = lCount + 1
                        If IsEmpty(vMin) Then
                            vMin = vValue
                            vMax = vValue
                        Else
                            If pvCompareValues(vValue, vMin, jgexSortTypeString) < 0 Then
                                vMin = vValue
                            End If
                            If pvCompareValues(vValue, vMax, jgexSortTypeString) > 0 Then
                                vMax = vValue
                            End If
                        End If
                        If IsNumeric(vValue) Then
                            dblValue = CDbl(vValue)
                            dblSum = dblSum + dblValue
                            dblSquares = dblSquares + dblValue * dblValue
                        End If
                    End If
                End If
            End If
        End If
    Next
    Select Case eFunc
    Case jgexCount, jgexValueCount
        pvAggregate = lCount
    Case jgexSum
        pvAggregate = dblSum
    Case jgexAvg
        If lCount > 0 Then
            pvAggregate = dblSum / lCount
        End If
    Case jgexMin
        pvAggregate = vMin
    Case jgexMax
        pvAggregate = vMax
    Case jgexStdDev
        '--- population deviation over the values that were numbers
        If lCount > 1 Then
            dblValue = (dblSquares - dblSum * dblSum / lCount) / lCount
            If dblValue > 0 Then
                pvAggregate = Sqr(dblValue)
            Else
                pvAggregate = 0
            End If
        End If
    End Select
End Function

Private Function pvIsChecked(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Boolean
    Dim vValue          As Variant

    vValue = frRowValue(lRowIndex, nColIndex)
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
    '--- the tick is the same seven runs of pixels at either dpi, but it sits
    '--- further into the box as the screen scales: two pixels in at 96, three
    '--- at 120. The font does not move it -- a 12pt cell font at 96 leaves it
    '--- exactly where an 8.25pt one does -- and neither DrawFrameControl nor
    '--- a plain dpi ratio reproduces it, so the two offsets are the ones the
    '--- recordings show and a third scale needs a third recording
    lOffX = 2
    lOffY = 4
    If pvScreenDpi() >= 120 Then
        lOffX = 3
        lOffY = 6
    End If
    '--- row, first column, last column
    vRuns = Array(0, 6, 6, 1, 5, 6, 2, 0, 0, 2, 4, 6, 3, 0, 1, 3, 3, 5, 4, 0, 4, 5, 1, 3, 6, 2, 2)
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

Private Function pvGroupCaption(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    Dim oCol            As JSColumn
    Dim vValue          As Variant
    Dim sText           As String

    '--- the value a group row shows: GroupFormat re-formats it -- for the
    '--- caption only, the original still breaks groups on the raw value --
    '--- an empty or Null one gives way to GroupEmptyStringCaption, and
    '--- GroupPrefix leads whichever of the two ends up there, joined with no
    '--- separator of its own -- a prefix brings its own spacing
    Set oCol = m_oColumns.Item(nColIndex)
    vValue = frRowValue(lRowIndex, nColIndex)
    Select Case VarType(vValue)
    Case vbEmpty, vbNull, vbObject, vbError
    Case Else
        If Not IsArray(vValue) Then
            If LenB(oCol.GroupFormat) <> 0 Then
                sText = Format$(vValue, oCol.GroupFormat)
            Else
                sText = pvCellText(lRowIndex, nColIndex)
            End If
        End If
    End Select
    If LenB(sText) = 0 Then
        sText = oCol.GroupEmptyStringCaption
    End If
    If LenB(oCol.GroupPrefix) <> 0 Then
        sText = oCol.GroupPrefix & " " & sText
    End If
    pvGroupCaption = sText
End Function

Private Function pvCellText(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As String
    Dim vValue          As Variant

    pvCellText = frRowDisplayValue(lRowIndex, nColIndex)
    If LenB(pvCellText) = 0 Then
        vValue = frRowValue(lRowIndex, nColIndex)
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

Private Sub pvInvalidate()
    Const FUNC_NAME     As String = "pvInvalidate"

    '--- tolerate refresh before the control window exists
    On Error GoTo EH
    '--- a batch of changes under Redraw = False paints once, when it is
    '--- turned back on
    If Not m_bRedraw Then
        Exit Sub
    End If
    pvUpdateScrollBars
    picGrid.Refresh
    If m_bRecordNavigator Then
        UserControl.Refresh
    End If
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Function pvTopHeight() As Long
    If m_bGroupByBoxVisible Then
        pvTopHeight = m_lColumnHeaderHeight + 14
    End If
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

Private Function pvFirstCol() As Integer
    '--- the left-most visible column: horizontal scrolling moves whole
    '--- columns, so painting and hit-testing simply start here
    pvFirstCol = m_nLeftCol
    If pvFirstCol < 1 Then
        pvFirstCol = 1
    End If
    If pvFirstCol > m_oColumns.Count Then
        pvFirstCol = m_oColumns.Count
    End If
End Function

Private Function pvScrollableWidth() As Long
    Dim nIdx            As Integer
    Dim oCol            As JSColumn
    Dim nFrozen         As Integer

    '--- the client strip the scrollable columns share, i.e. what is left of
    '--- it once the row header and the frozen block have taken their part
    pvScrollableWidth = picGrid.ScaleWidth
    If m_bRowHeaders Then
        pvScrollableWidth = pvScrollableWidth - 18
    End If
    nFrozen = pvFrozenCount()
    For nIdx = 1 To nFrozen
        Set oCol = m_oColumns.ItemByPosition(nIdx)
        If oCol.Visible Then
            pvScrollableWidth = pvScrollableWidth - pvColWidth(oCol)
        End If
    Next
End Function

Private Function pvVisibleColCount() As Integer
    Dim nIdx            As Integer

    For nIdx = 1 To m_oColumns.Count
        If m_oColumns.ItemByPosition(nIdx).Visible Then
            pvVisibleColCount = pvVisibleColCount + 1
        End If
    Next
End Function

Private Function pvClampRow(ByVal lRow As Long) As Long
    pvClampRow = lRow
    If pvClampRow < 1 Then
        pvClampRow = 1
    ElseIf pvClampRow > RowCount Then
        pvClampRow = RowCount
    End If
End Function

Private Function pvHitScrollBar() As Boolean
    Dim uPt             As POINTAPI

    If m_hScrollH = 0 Then
        Exit Function
    End If
    Call GetCursorPos(uPt)
    pvHitScrollBar = (WindowFromPoint(uPt.X, uPt.Y) = m_hScrollH)
End Function

Private Function pvNavLayout(uNav As UcsNavLayout) As Boolean
    Dim lBtnW           As Long
    Dim lBtnTop         As Long
    Dim lBtnH           As Long
    Dim lBoxW           As Long
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
    hPrevFont = pvSelectFont(UserControl.hDC, m_oFont)
    lX = 4
    uNav.PrefixX = lX
    lX = lX + pvTextWidth(UserControl.hDC, uNav.Prefix) + 4
    pvSetRect uNav.BtnFirst, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    pvSetRect uNav.BtnPrev, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW + 4
    '--- the box is a fixed size -- 5, 50 and 500 records render identically --
    '--- but it does not scale linearly with dpi: 46px at 96 and 53px at 120,
    '--- which this fits. Two scales cannot identify the formula uniquely, so
    '--- it wants confirming against a 144dpi golden
    lBoxW = FontTextMetrics(m_oFont).tmHeight + uNav.BandH + 16
    '--- the box is taller than the band and clipped by it, like the original's
    '--- TextBox (53x24 inside a 22px band at 120dpi)
    pvSetRect uNav.Box, lX, uNav.BandTop, lX + lBoxW, uNav.BandTop + uNav.BandH + 4
    lX = lX + lBoxW + 4
    uNav.MiddleX = lX
    lX = lX + pvTextWidth(UserControl.hDC, uNav.Middle) + 8
    pvSetRect uNav.BtnNext, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    pvSetRect uNav.BtnLast, lX, lBtnTop, lX + lBtnW, lBtnTop + lBtnH
    lX = lX + lBtnW
    uNav.Width = lX
    Call SelectObject(UserControl.hDC, hPrevFont)
    pvNavLayout = True
End Function

Private Sub pvSetRect(uRect As RECT, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
    uRect.Left = lLeft
    uRect.Top = lTop
    uRect.Right = lRight
    uRect.Bottom = lBottom
End Sub

Private Function pvTextWidth(ByVal hDC As Long, sText As String) As Long
    Dim uRect           As RECT

    Call DrawText(hDC, StrPtr(sText), Len(sText), uRect, DT_SINGLELINE Or DT_CALCRECT Or DT_NOPREFIX)
    pvTextWidth = uRect.Right
End Function

Private Function pvNavigatorWidth() As Long
    Dim uNav            As UcsNavLayout

    If pvNavLayout(uNav) Then
        pvNavigatorWidth = uNav.Width
    End If
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
    lCenter = (uBtn.Top + uBtn.Bottom - 1) \ 2
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
    '--- 96dpi, 3px at 5 in at 120dpi
    lBarX = uNav.BandH \ 4
    lBarW = uNav.BandH \ 7
    lBarY = uNav.BandH \ 5
    pvFillRect hDC, uNav.BtnFirst.Left + lBarX, uNav.BtnFirst.Top + lBarY, uNav.BtnFirst.Left + lBarX + lBarW, uNav.BtnFirst.Bottom - lBarY - 1, vbBlack
    pvFillRect hDC, uNav.BtnLast.Right - 1 - lBarX - lBarW, uNav.BtnLast.Top + lBarY, uNav.BtnLast.Right - 1 - lBarX, uNav.BtnLast.Bottom - lBarY - 1, vbBlack
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
    pvDrawText hDC, CStr(m_lRow), uNav.Box.Left + 2, uNav.Box.Top + 2, uNav.Box.Right - 4, uNav.BandTop + uNav.BandH - 2, m_clrForeColor, vbWindowBackground, jgexAlignRight, uNav.Box.Left + 2, uNav.Box.Right - 2
    Call SelectObject(hDC, hPrevFont)
End Sub

Private Sub pvLayoutGrid()
    Dim lBandH          As Long

    '--- the grid surface owns everything above the scrollbar band, so its
    '--- own WS_VSCROLL stops where the band starts -- which is what makes
    '--- the vertical thumb geometry match the original
    If m_bBand Then
        lBandH = GetSystemMetrics(SM_CYHSCROLL)
    End If
    picGrid.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - lBandH
End Sub

Private Sub pvLayoutHScroll(ByVal bNeedH As Boolean, ByVal bNeedV As Boolean)
    Dim lBandH          As Long

    '--- a VB6 HScrollBar rather than a Win32 SCROLLBAR child: the original
    '--- is a VB6 control too, so the runtime draws both with the same code
    '--- and the shaft dither matches by construction
    If Not bNeedH Then
        hsbGrid.Visible = False
        m_hScrollH = 0
        Exit Sub
    End If
    lBandH = GetSystemMetrics(SM_CYHSCROLL)
    '--- the band stops at the client width, leaving the usual corner gap
    '--- under the vertical bar, as the original's does
    hsbGrid.Move pvNavigatorWidth(), UserControl.ScaleHeight - lBandH, picGrid.ScaleWidth - pvNavigatorWidth(), lBandH
    hsbGrid.Visible = True
    m_hScrollH = hsbGrid.hWnd
End Sub

Private Sub pvUpdateScrollBars()
    Dim lTopH           As Long
    Dim lAvailH         As Long
    Dim lAvailW         As Long
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
    '--- both scrollbars interact: each one steals space from the other
    lAvailH = picGrid.ScaleHeight - lTopH
    lAvailW = picGrid.ScaleWidth
    bNeedV = (m_lRowHeight > 0 And RowCount > 0 And RowCount * m_lRowHeight > lAvailH)
    bNeedH = (lColsW > lAvailW)
    If bNeedV Then
        lAvailW = lAvailW - GetSystemMetrics(SM_CXVSCROLL)
        bNeedH = (lColsW > lAvailW)
    End If
    If bNeedH Then
        lAvailH = lAvailH - GetSystemMetrics(SM_CYHSCROLL)
        bNeedV = (m_lRowHeight > 0 And RowCount > 0 And RowCount * m_lRowHeight > lAvailH)
        If bNeedV Then
            lAvailW = picGrid.ScaleWidth - GetSystemMetrics(SM_CXVSCROLL)
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
            If m_lFirstItem > 0 Then
                .nPos = m_lFirstItem - 1
            End If
        End With
        Call SetScrollInfo(picGrid.hWnd, SB_VERT, uSi, 1)
    End If
    m_bHScroll = bNeedH
    '--- the band hosts the navigator as well, so it can be there without a
    '--- horizontal scrollbar
    m_bBand = bNeedH Or m_bRecordNavigator
    '--- the band appearing takes height away from the grid surface
    pvLayoutGrid
    pvLayoutHScroll bNeedH, bNeedV
    If bNeedH Then
        '--- the bar scrolls the columns after the frozen block, so both its
        '--- range and its position count from there
        lPage = pvVisibleColsInWidth(pvScrollableWidth(), pvFrozenCount() + 1)
        If lPage < 1 Then
            lPage = 1
        End If
        hsbGrid.LargeChange = 1
        hsbGrid.Min = 0
        hsbGrid.Max = pvVisibleColCount() - pvFrozenCount() - lPage
        If hsbGrid.Max < 1 Then
            hsbGrid.Max = 1
        End If
        If m_nLeftCol - pvFrozenCount() - 1 >= 0 And m_nLeftCol - pvFrozenCount() - 1 <= hsbGrid.Max Then
            hsbGrid.Value = m_nLeftCol - pvFrozenCount() - 1
        End If
        pvSetHScrollInfo
    End If
    m_bScrollUpdating = False
End Sub

Private Sub pvSetHScrollInfo()
    Dim uSi             As SCROLLINFO

    '--- VB6 maps a scrollbar's Min..Max onto 0..32767 before handing it to
    '--- Windows, and the thumb it computes from that lands a pixel off the
    '--- original's. The original keeps the column numbers themselves in the
    '--- scroll info (min=first scrollable, max=last valid LeftCol, page=1),
    '--- read back with GetScrollInfo, so set those over VB6's mapping
    uSi.cbSize = Len(uSi)
    uSi.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
    uSi.nMin = pvFrozenCount() + 1
    uSi.nMax = pvVisibleColCount() - pvVisibleColsInWidth(pvScrollableWidth(), pvFrozenCount() + 1) + 1
    If uSi.nMax < uSi.nMin Then
        uSi.nMax = uSi.nMin
    End If
    uSi.nPage = 1
    uSi.nPos = m_nLeftCol
    Call SetScrollInfo(hsbGrid.hWnd, SB_CTL, uSi, 0)
    hsbGrid.Refresh
End Sub

Private Function pvTotalColWidth() As Long
    Dim nIdx            As Integer
    Dim oCol            As JSColumn

    For nIdx = 1 To m_oColumns.Count
        Set oCol = m_oColumns.ItemByPosition(nIdx)
        If oCol.Visible Then
            pvTotalColWidth = pvTotalColWidth + pvColWidth(oCol)
        End If
    Next
End Function

Private Function pvColWidth(oCol As JSColumn) As Long
    Dim nIdx            As Integer
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
    For nIdx = 1 To m_oColumns.Count
        Set oItem = m_oColumns.ItemByPosition(nIdx)
        If oItem.Visible Then
            lTotal = lTotal + ToPixels(oItem.Width)
        End If
    Next
    If lTotal <= 0 Or lAvail <= 0 Then
        pvColWidth = ToPixels(oCol.Width)
        Exit Function
    End If
    For nIdx = 1 To m_oColumns.Count
        Set oItem = m_oColumns.ItemByPosition(nIdx)
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
    Dim nIdx            As Integer
    Dim oCol            As JSColumn
    Dim aOrder()        As Integer
    Dim lCount          As Long
    Dim nFrozen         As Integer

    '--- painting and hit-testing walk the columns in this order: the frozen
    '--- ones pinned at the left, then the scrollable rest from LeftCol. Both
    '--- blocks skip hidden columns, so callers never test Visible again
    ReDim aOrder(0 To m_oColumns.Count) As Integer
    nFrozen = pvFrozenCount()
    For nIdx = 1 To nFrozen
        Set oCol = m_oColumns.ItemByPosition(nIdx)
        If oCol.Visible Then
            aOrder(lCount) = nIdx
            lCount = lCount + 1
        End If
    Next
    For nIdx = pvFirstCol() To m_oColumns.Count
        If nIdx > nFrozen Then
            Set oCol = m_oColumns.ItemByPosition(nIdx)
            If oCol.Visible Then
                aOrder(lCount) = nIdx
                lCount = lCount + 1
            End If
        End If
    Next
    If lCount = 0 Then
        pvColOrder = Array()
        Exit Function
    End If
    ReDim Preserve aOrder(0 To lCount - 1) As Integer
    pvColOrder = aOrder
End Function

Private Function pvOrderMax(vOrder As Variant) As Long
    '--- -1 for an empty order, so the paint loops simply do not run
    pvOrderMax = -1
    If IsArray(vOrder) Then
        pvOrderMax = UBound(vOrder)
    End If
End Function

Private Function pvFrozenCount() As Integer
    '--- more frozen columns than there are leaves nothing to scroll
    pvFrozenCount = m_nFrozenColumns
    If pvFrozenCount < 0 Then
        pvFrozenCount = 0
    End If
    If pvFrozenCount > m_oColumns.Count Then
        pvFrozenCount = m_oColumns.Count
    End If
End Function

Private Function pvVisibleColsInWidth(ByVal lWidth As Long, Optional ByVal nStart As Integer = 1) As Long
    Dim nIdx            As Integer
    Dim oCol            As JSColumn
    Dim lCum            As Long

    '--- how many whole columns fit in lWidth starting at nStart
    For nIdx = nStart To m_oColumns.Count
        Set oCol = m_oColumns.ItemByPosition(nIdx)
        If oCol.Visible Then
            lCum = lCum + pvColWidth(oCol)
            '--- strictly less: a column ending exactly on the edge does not
            '--- count as visible, which is what decides the thumb size when
            '--- the columns happen to fill the client precisely
            If lCum >= lWidth Then
                Exit Function
            End If
            pvVisibleColsInWidth = pvVisibleColsInWidth + 1
        End If
    Next
    If pvVisibleColsInWidth < 1 Then
        pvVisibleColsInWidth = 1
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
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
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
            pvEndEdit True
            '--- committing steps to the next row, as the original does
            If m_lRow < RowCount Then
                pvNavigate m_lRow + 1, m_nCol, 0, False
            End If
            Handled = True
        Case vbKeyEscape
            pvEndEdit False
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
            pvOnNavigatorClick pvLoWord(lParam), pvHiWord(lParam)
        Case WM_MOUSEACTIVATE
            If pvHitScrollBar() Then
                ControlSubclassProc = MA_NOACTIVATE
                Handled = True
            End If
        Case WM_CTLCOLORSCROLLBAR
            Handled = True
        End Select
        GoTo QH
    End If
    Select Case wMsg
    Case WM_VSCROLL
        pvOnVScroll wParam And &HFFFF&, (wParam And &HFFFF0000) \ &H10000
        Handled = True
    Case WM_KEYDOWN
        nKeyCode = CInt(wParam And &HFFFF&)
        nShift = pvShiftState()
        RaiseEvent KeyDown(nKeyCode, nShift)
        pvOnKeyDown nKeyCode, nShift
    Case WM_LBUTTONDOWN
        '--- the cell change and the editor it opens are announced before the
        '--- click itself reaches the client, which is the order the original
        '--- raises them in. The mouse events carry container units, as it
        '--- reports them too: a click 38 pixels in comes out as 570 twips
        pvOnLButtonDown pvLoWord(lParam), pvHiWord(lParam), pvMouseShift(wParam)
        RaiseEvent MouseDown(vbLeftButton, pvMouseShift(wParam), pvLoWord(lParam) * Screen.TwipsPerPixelX, pvHiWord(lParam) * Screen.TwipsPerPixelY)
        '--- the grid's own click handling takes the focus for itself, which
        '--- would kill the caret an in-place editor just created, so the
        '--- default runs here and the editor takes the focus back after it
        If m_hWndEdit <> 0 Then
            ControlSubclassProc = CallNextSubclassProc(m_pSubclassPic, hWnd, wMsg, wParam, lParam)
            Handled = True
            Call SetFocusApi(m_hWndEdit)
        End If
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton, pvMouseShift(wParam), pvLoWord(lParam) * Screen.TwipsPerPixelX, pvHiWord(lParam) * Screen.TwipsPerPixelY)
        RaiseEvent Click
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(pvMouseButton(wParam), pvMouseShift(wParam), pvLoWord(lParam) * Screen.TwipsPerPixelX, pvHiWord(lParam) * Screen.TwipsPerPixelY)
        If (wParam And MK_LBUTTON) <> 0 Then
            pvOnMouseDrag pvHiWord(lParam)
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

    '--- dragging with the left button held extends a multi-select range
    '--- from the mouse-down anchor to the row under the cursor
    If Not m_bMultiSelect Or m_lRowHeight <= 0 Then
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

Private Function pvMakeDWord(ByVal lLoWord As Long, ByVal lHiWord As Long) As Long
    '--- the lParam shape the window messages take, low word first
    pvMakeDWord = (lLoWord And &HFFFF&) Or (lHiWord * &H10000)
End Function

Private Function pvLoWord(ByVal lValue As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, lValue, 2)
    pvLoWord = nWord
End Function

Private Function pvHiWord(ByVal lValue As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, ByVal VarPtr(lValue) + 2, 2)
    pvHiWord = nWord
End Function

Private Function pvColAtX(ByVal lX As Long, oCol As JSColumn) As Integer
    Dim lCum            As Long
    Dim nIdx            As Integer
    Dim nPos            As Integer
    Dim oItem           As JSColumn
    Dim vOrder          As Variant
    Dim lIdx            As Long

    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        nPos = nPos + 1
        lCum = lCum + pvColWidth(oItem)
        If lX < lCum Then
            Set oCol = oItem
            pvColAtX = nPos
            Exit Function
        End If
    Next
End Function

Private Function pvGroupAtPoint(ByVal lX As Long, ByVal lY As Long, oGroup As JSGroup) As Integer
    Dim lIdx            As Long
    Dim oItem           As JSGroup
    Dim hDC             As Long
    Dim hPrevFont       As Long

    '--- laid out again rather than read off the last paint: a control that
    '--- has not painted yet still hit-tests, and a caption or font changed
    '--- since would otherwise leave the rectangles behind. It costs one text
    '--- measurement per group level, on a click
    hDC = picGrid.hDC
    hPrevFont = pvSelectFont(hDC, m_oFont)
    pvLayoutGroupChips hDC, 0
    Call SelectObject(hDC, hPrevFont)
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

Private Function pvBeginEdit(ByVal lPos As Long, ByVal nCol As Integer, ByVal bSelectAll As Boolean, Optional ByVal lClickX As Long = -1, Optional ByVal lClickY As Long = -1) As Boolean
    Dim oCol            As JSColumn
    Dim oCancel         As JSRetBoolean
    Dim lRowIndex       As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lW              As Long
    Dim lStyle          As Long
    Dim lBoxTop         As Long
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
    lRowIndex = pvDataRow(lPos)
    If lRowIndex <= 0 Then
        Exit Function
    End If
    Set oCol = pvColByPosition(nCol)
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
    If Not pvCellRect(lPos, nCol, lX, lY, lW) Then
        Exit Function
    End If
    If oCol.EditType = jgexEditCheckBox Then
        '--- a checkbox has no editor to show: the click is the edit, so the
        '--- value flips there and then and the client hears one Change. Only
        '--- a click level with the box counts, which is the whole of why the
        '--- same point toggles at 96dpi and not at 120: the taller bands
        '--- there put the row lower and leave the point above the box
        lBoxTop = lY + (pvRowContentH(m_lRowHeight) - CHECK_BOX_H) \ 2
        If lClickY >= lBoxTop And lClickY < lBoxTop + CHECK_BOX_H Then
            frRowValue(lRowIndex, oCol.Index) = Not pvIsChecked(lRowIndex, oCol.Index)
            RaiseEvent Change
            pvInvalidate
        End If
        pvBeginEdit = True
        Exit Function
    End If
    m_bInEditSetup = True
    m_bEditing = True
    m_lEditRow = lRowIndex
    m_nEditCol = oCol.Index
    m_sEditOldValue = pvCellText(lRowIndex, oCol.Index)
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
    m_hWndEdit = CreateWindowEx(0, StrPtr("EDIT"), 0, lStyle, lX + 1, lTextTop, lW - 2, lEditH, hWnd, 0, App.hInstance, ByVal 0&)
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
        '--- where typing goes and, in a wrapping cell, which line shows
        lCaret = SendMessage(m_hWndEdit, EM_CHARFROMPOS, 0, pvMakeDWord(lClickX - lX, lClickY - lTextTop))
        If lCaret = -1 Then
            '--- the point fell outside the editor, which answers with the end
            lCaret = Len(m_sEditOldValue)
        Else
            lCaret = lCaret And &HFFFF&
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

Private Sub pvEndEdit(ByVal bCommit As Boolean)
    Dim oCancel         As JSRetBoolean
    Dim nCol            As Integer
    Dim lRowIndex       As Long
    Dim sText           As String
    Dim oRowData        As JSRowData

    If Not m_bEditing Then
        Exit Sub
    End If
    nCol = m_nEditCol
    lRowIndex = m_lEditRow
    sText = pvEditText()
    m_bEditing = False
    pvDestroyEditor
    '--- committing runs the update trio the original raises in this order:
    '--- the cell first, then the row, with a repaint between the two halves
    If bCommit And sText <> m_sEditOldValue Then
        Set oCancel = New JSRetBoolean
        RaiseEvent BeforeColUpdate(lRowIndex, nCol, m_sEditOldValue, oCancel)
        If oCancel.Value Then
            RaiseEvent AfterColEdit(nCol)
            pvSetFocusBack
            Exit Sub
        End If
        frRowValue(lRowIndex, nCol) = sText
        Set oRowData = GetRowData(pvDataRowPos(lRowIndex))
        RaiseEvent AfterColUpdate(nCol)
        RaiseEvent AfterColEdit(nCol)
        Set oCancel = New JSRetBoolean
        RaiseEvent BeforeUpdate(oCancel)
        RaiseEvent RowFormat(oRowData)
        If Not oCancel.Value Then
            '--- no UnboundUpdate here: the original commits the cell into its
            '--- own buffer and says only AfterUpdate about it
            RaiseEvent AfterUpdate
        End If
        RaiseEvent RowFormat(oRowData)
    Else
        '--- a cancelled edit repaints the cell it was covering and says only
        '--- that the session ended
        Set oRowData = GetRowData(pvDataRowPos(lRowIndex))
        RaiseEvent RowFormat(oRowData)
        RaiseEvent AfterColEdit(nCol)
    End If
    pvSetFocusBack
    pvInvalidate
End Sub

Private Sub pvSetFocusBack()
    '--- the editor window is gone by now, so the grid takes the focus back
    '--- through the API rather than through VB, which raises when the
    '--- control is not in a state to take it
    Call SetFocusApi(hWnd)
End Sub

Private Function pvGroupByBoxHeight() As Long
    If m_bGroupByBoxVisible Then
        pvGroupByBoxHeight = m_lColumnHeaderHeight + 14
        If m_oGroups.Count > 1 Then
            '--- the box grew to hold the staircase, so the bands below it did
            '--- move down -- pvPaintGroupByBox sizes it the same way
            pvGroupByBoxHeight = pvGroupByBoxHeight + (m_oGroups.Count - 1) * pvChipStagger()
        End If
    End If
End Function

Private Function pvRowsTop() As Long
    pvRowsTop = pvGroupByBoxHeight()
    If m_bColumnHeaders Then
        pvRowsTop = pvRowsTop + m_lColumnHeaderHeight
    End If
End Function

Private Function pvColByPosition(ByVal nPos As Integer) As JSColumn
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim nSeen           As Integer
    Dim oItem           As JSColumn

    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oItem.Visible Then
            nSeen = nSeen + 1
            If nSeen = nPos Then
                Set pvColByPosition = oItem
                Exit Function
            End If
        End If
    Next
End Function

Private Function pvCellRect(ByVal lPos As Long, ByVal nCol As Integer, lX As Long, lY As Long, lW As Long) As Boolean
    Dim vOrder          As Variant
    Dim lIdx            As Long
    Dim nSeen           As Integer
    Dim oItem           As JSColumn
    Dim lCum            As Long

    '--- where the cell sits in the control, which is where the editor goes
    If m_lRowHeight <= 0 Or lPos < m_lFirstItem Then
        Exit Function
    End If
    lY = pvRowsTop() + (lPos - m_lFirstItem) * m_lRowHeight
    lCum = pvBlockLeft()
    vOrder = pvColOrder()
    For lIdx = 0 To pvOrderMax(vOrder)
        Set oItem = m_oColumns.ItemByPosition(vOrder(lIdx))
        If oItem.Visible Then
            nSeen = nSeen + 1
            If nSeen = nCol Then
                lX = lCum
                lW = pvColWidth(oItem)
                pvCellRect = True
                Exit Function
            End If
            lCum = lCum + pvColWidth(oItem)
        End If
    Next
End Function

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
    If pvColSortOrder(oCol) = jgexSortAscending Then
        m_oSortKeys.Clear
        m_oSortKeys.Add oCol.Index, jgexSortDescending
    Else
        m_oSortKeys.Clear
        m_oSortKeys.Add oCol.Index, jgexSortAscending
    End If
End Sub

Private Sub pvOnLButtonDown(ByVal lX As Long, ByVal lY As Long, ByVal nShift As Integer)
    Dim lTopGbox        As Long
    Dim lTopHdr         As Long
    Dim nPos            As Integer
    Dim oCol            As JSColumn
    Dim oGroup          As JSGroup
    Dim lRow            As Long

    lTopGbox = pvGroupByBoxHeight()
    lTopHdr = pvRowsTop()
    If lY < lTopGbox Then
        '--- group-by box: a chip stands for its column, so clicking one sorts
        '--- exactly as clicking that column's header does
        nPos = pvGroupAtPoint(lX, lY, oGroup)
        If Not oGroup Is Nothing Then
            RaiseEvent GroupByBoxHeaderClick(oGroup)
            If m_bAutomaticSort Then
                pvAutoSort m_oColumns.Item(oGroup.ColIndex)
            End If
        End If
    ElseIf lY < lTopHdr Then
        '--- column header band
        nPos = pvColAtX(lX, oCol)
        If Not oCol Is Nothing Then
            RaiseEvent ColumnHeaderClick(oCol)
            If m_bAutomaticSort Then
                pvAutoSort oCol
            End If
        End If
    ElseIf m_lRowHeight > 0 Then
        '--- data area cell
        lRow = m_lFirstItem + (lY - lTopHdr) \ m_lRowHeight
        nPos = pvColAtX(lX, oCol)
        If lRow >= 1 And lRow <= RowCount And nPos >= 1 Then
            pvNavigate lRow, nPos, nShift, (nShift And vbCtrlMask) <> 0
            '--- clicking a cell opens its editor, which is what BeforeColEdit
            '--- announces -- and it happens before the client sees MouseDown
            pvBeginEdit lRow, nPos, False, lX, lY
        End If
    End If
End Sub

Private Sub pvOnKeyDown(ByVal nKeyCode As Integer, ByVal nShift As Integer)
    Select Case nKeyCode
    Case vbKeyDown
        If m_lRow < RowCount Then
            pvNavigate m_lRow + 1, m_nCol, nShift, False
        End If
    Case vbKeyUp
        If m_lRow > 1 Then
            pvNavigate m_lRow - 1, m_nCol, nShift, False
        End If
    Case vbKeyRight
        If m_nCol < pvVisibleColCount() Then
            Col = m_nCol + 1
        End If
    Case vbKeyLeft
        If m_nCol > 1 Then
            Col = m_nCol - 1
        End If
    Case vbKeyPageDown
        pvNavigate pvClampRow(m_lRow + pvVisibleRows()), m_nCol, nShift, False
    Case vbKeyPageUp
        pvNavigate pvClampRow(m_lRow - pvVisibleRows()), m_nCol, nShift, False
    Case vbKeyHome
        pvNavigate pvClampRow(1), m_nCol, nShift, False
    Case vbKeyEnd
        pvNavigate pvClampRow(RowCount), m_nCol, nShift, False
    End Select
End Sub

Private Sub pvSetRow(ByVal lValue As Long)
    Dim lLastRow        As Long

    '--- moves the current row without touching the selection: the row is
    '--- painted selected, so it still has to repaint
    If m_lRow <> lValue Then
        lLastRow = m_lRow
        m_lRow = lValue
        pvInvalidate
        RaiseEvent RowColChange(lLastRow, m_nCol)
    End If
End Sub

'--- moves the current cell and updates the selection accordingly
Private Sub pvNavigate(ByVal lRow As Long, ByVal nCol As Integer, ByVal nShift As Integer, ByVal bCtrlToggle As Boolean)
    '--- the selection lands on the new row before the move is announced,
    '--- which is the order the original raises the two events in
    If lRow >= 1 And lRow <= RowCount Then
        pvUpdateSelection lRow, nShift, bCtrlToggle
        pvSetRow lRow
    Else
        pvUpdateSelection m_lRow, nShift, bCtrlToggle
    End If
    If nCol >= 1 Then
        Col = nCol
    End If
    EnsureVisible m_lRow
End Sub

Private Sub pvUpdateSelection(ByVal lRow As Long, ByVal nShift As Integer, ByVal bCtrlToggle As Boolean)
    If lRow < 1 Or lRow > RowCount Then
        Exit Sub
    End If
    If m_bMultiSelect And bCtrlToggle Then
        pvToggleSel lRow
    ElseIf m_bMultiSelect And (nShift And vbShiftMask) <> 0 Then
        pvSetRangeSel m_lSelAnchor, lRow
    Else
        pvSetSingleSel lRow
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
    lRow = pvDataRow(lPos)
    pvEnsureRow lRow
    m_oSelectedItems.frAdd lPos, m_aRows(lRow).Bookmark, lRow
End Sub

Private Sub pvSetSingleSel(ByVal lPos As Long)
    Dim bSame           As Boolean

    '--- re-selecting the row that already is the whole selection changes
    '--- nothing, and the original stays quiet about it -- a click that only
    '--- moves the column raises RowColChange and no more
    bSame = (m_oSelectedItems.Count = 1)
    If bSame Then
        bSame = (m_oSelectedItems.Item(1).RowPosition = lPos)
    End If
    m_oSelectedItems.Clear
    pvAddSel lPos
    m_lSelAnchor = lPos
    If Not bSame Then
        RaiseEvent SelectionChange
    End If
    pvInvalidate
End Sub

Private Sub pvToggleSel(ByVal lPos As Long)
    If pvIsRowSelected(lPos) Then
        m_oSelectedItems.RemoveRowPosition lPos
    Else
        pvAddSel lPos
    End If
    m_lSelAnchor = lPos
    RaiseEvent SelectionChange
    pvInvalidate
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
    RaiseEvent SelectionChange
    pvInvalidate
End Sub

Private Function pvClampCol(ByVal nValue As Integer) As Integer
    pvClampCol = nValue
    If pvClampCol < 1 Then
        pvClampCol = 1
    End If
    If pvClampCol > m_oColumns.Count Then
        pvClampCol = m_oColumns.Count
    End If
End Function

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
    '--- both heights follow their font, so they are DPI-dependent (19px
    '--- and 19/22px at 96/120dpi); derive them instead of hardcoding
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
End Sub

Private Sub UserControl_InitProperties()
    '--- a freshly placed control starts with two default empty columns
    pvInheritAmbientFont
    m_oColumns.Add(vbNullString).Width = ToTwips(m_lDefaultColumnWidth)
    m_oColumns.Add(vbNullString).Width = ToTwips(m_lDefaultColumnWidth)
    pvSubclass
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pvInheritAmbientFont
    pvSubclass
End Sub

Private Sub hsbGrid_Change()
    pvOnHScroll
End Sub

Private Sub hsbGrid_Scroll()
    pvOnHScroll
End Sub

Private Sub pvOnHScroll()
    If m_bScrollUpdating Then
        Exit Sub
    End If
    LeftCol = pvClampCol(hsbGrid.Value + pvFrozenCount() + 1)
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    '--- default row height follows the data font unless explicitly set
    If Not m_bRowHeightSet Then
        m_lRowHeight = FontTextMetrics(m_oFont).tmHeight + 3
        If m_lRowHeight < 19 Then
            m_lRowHeight = 19
        End If
    End If
    pvInvalidate
End Sub

Private Sub m_oColumnHeaderFont_FontChanged(ByVal PropertyName As String)
    '--- header height always follows the header font
    m_lColumnHeaderHeight = FontTextMetrics(m_oColumnHeaderFont).tmHeight + 6
    pvInvalidate
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
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub picGrid_Paint()
    Const FUNC_NAME     As String = "picGrid_Paint"

    On Error GoTo EH
    pvPaint picGrid.hDC
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub UserControl_Resize()
    pvLayoutGrid
    pvInvalidate
End Sub

Private Sub UserControl_Terminate()
    pvDestroyEditor
    pvUnsubclass
    '--- detach outstanding JSRowData wrappers, of both kinds, so their weak
    '--- owner pointers cannot dangle past the control lifetime
    pvEraseDataRows
    pvEraseGroupRows
    '--- and the collections that point back for their change notifications
    m_oGroups.frTerminate
    m_oSortKeys.frTerminate
    m_oColumns.frTerminate
End Sub
