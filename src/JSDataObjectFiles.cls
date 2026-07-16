VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSDataObjectFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of filenames used by the JSDataObject object (jgexCFFiles format only)."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of filenames used by the JSDataObject object
'
'=========================================================================
Option Explicit

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns an specific file name."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of file names in the collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Sub Add(ByVal FileName As String, Optional ByVal Index As Variant)
Attribute Add.VB_Description = "Adds a filename to the Files collection of a JSDataObject object (jgexCFFiles format only)."
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes an specific file name."
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all file names in the collection."
End Sub
