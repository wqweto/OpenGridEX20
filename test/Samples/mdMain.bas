Attribute VB_Name = "mdMain"
'=========================================================================
'
' Open GridEX 2000 Control
' Smoke runner for the ported Janus unbound samples: loads each form,
' lets it paint, and asserts the grid actually took the sample's data
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdMain"

'=========================================================================
' API
'=========================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim sError          As String

    On Error GoTo EH
    TestInit OutputFile("Samples.out.txt")
    '--- the five unbound samples of the original, ported by re-pointing the
    '--- reference and calling the setup Sub generated from their snapshot.
    '--- Expected row counts come from the sample data itself: Products has
    '--- 77 rows, the three hand-built datasets have 3 each
    pvRunSample frmUnbound1, "Unbound 1", 77, 10
    pvRunSample frmUnbound2, "Unbound 2", 77, 13
    pvRunSample frmUnboundArray, "Unbound Array", 3, 3
    pvRunSample frmUnboundUDT, "Unbound Array UDTs", 3, 3
    pvRunSample frmUnboundCol, "Unbound Collection", 3, 3
QH:
    TestsDone
    pvUnloadAll
    Exit Sub
EH:
    '--- PrintError resets Err, so the assert has to read it first
    sError = "&H" & Hex$(Err.Number) & " " & Err.Description
    PrintError FUNC_NAME
    Assert "Unhandled error " & sError, False
    GoTo QH
End Sub

Private Sub pvRunSample(oForm As Form, sName As String, ByVal lRows As Long, ByVal lCols As Long)
    Const FUNC_NAME     As String = "pvRunSample"
    Dim oGrid           As GridEX

    '--- a sample that raises on load must not take the whole run with it
    On Error GoTo EH
    Load oForm
    oForm.Show
    pvSettle
    Set oGrid = oForm.Controls("GridEX1")
    AssertEquals sName & ": RowCount", lRows, oGrid.RowCount
    AssertEquals sName & ": Columns.Count", lCols, oGrid.Columns.Count
    Assert sName & ": row 1 read through UnboundReadData", LenB(oGrid.GetRowData(1).Value(1) & "") <> 0
    Unload oForm
    Exit Sub
EH:
    PrintError FUNC_NAME
    Assert sName & " loads", False
End Sub

Private Sub pvSettle()
    Dim lIdx            As Long

    '--- paint and any load-time timer of the control get a chance to run
    For lIdx = 1 To 10
        DoEvents
        Sleep 1
    Next
End Sub

Private Sub pvUnloadAll()
    Dim lIdx            As Long

    For lIdx = Forms.Count - 1 To 0 Step -1
        Unload Forms(lIdx)
    Next
End Sub
