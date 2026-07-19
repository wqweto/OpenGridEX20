VERSION 5.00
Object = "{24F4AB9F-37F4-43D4-B383-FB6CD721B629}#1.0#0"; "OpenGridEX20.ocx"
Begin VB.Form frmWeak
   Caption         =   "WeakRef host"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   3000
   ScaleWidth      =   4500
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
      Height          =   800
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1411
   End
End
Attribute VB_Name = "frmWeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' Disposable host form for the JSRowData weak reference test and for
' fresh control instances in the snapshot corpus round-trip test
'
'=========================================================================
Option Explicit
DefObj A-Z
