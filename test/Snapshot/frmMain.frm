VERSION 5.00
Object = "{24F4AB9F-37F4-43D4-B383-FB6CD721B629}#1.0#0"; "OpenGridEX20.ocx"
Begin VB.Form frmMain
   Caption         =   "Snapshot"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin OpenGridEX20.GridEX GridEX1
      Height          =   2000
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
   Begin OpenGridEX20.GEXPreview GEXPreview1
      Height          =   2000
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' Snapshot engine smoke test: dumps both hosted controls to JSON files
'
'=========================================================================
Option Explicit

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"

    On Error GoTo EH
    WriteTextFile App.Path & "\GridEX.json", SnapshotToJson(GridEX1.Object, "GridEX", False)
    WriteTextFile App.Path & "\GEXPreview.json", SnapshotToJson(GEXPreview1.Object, "GEXPreview", False)
QH:
    Unload Me
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    WriteTextFile App.Path & "\SnapshotError.txt", "Error &H" & Hex$(Err.Number) & " " & Err.Description
    GoTo QH
End Sub
