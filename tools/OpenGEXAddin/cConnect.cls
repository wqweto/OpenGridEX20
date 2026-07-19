VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' Open GridEX 2000 Control
' Export add-in IDE connector: batch-exports GridEX design-time snapshots
' when OPENGEX_SNAPSHOT_DIR environment variable is set (see export.ps1)
'
'=========================================================================
Option Explicit
DefObj A-Z

Implements AddInDesignerObjects.IDTExtensibility2

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oVBE                      As VBIDE.VBE

'=========================================================================
' Interface IDTExtensibility2
'=========================================================================

Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set m_oVBE = Application
    If ConnectMode = ext_cm_AfterStartup Then
        ExportActiveProject m_oVBE
    End If
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)
    ExportActiveProject m_oVBE
End Sub

Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Set m_oVBE = Nothing
End Sub

Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)
End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
End Sub
