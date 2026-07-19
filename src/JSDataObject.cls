VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Used for specifying formats and data that will be supported for OLE drag and drop operations."
'=========================================================================
'
' Open GridEX 2000 Control
' Used for specifying formats and data that will be supported for OLE
' drag and drop operations
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cData                     As Collection
Private m_oFiles                    As JSDataObjectFiles

'=========================================================================
' Properties
'=========================================================================

Public Property Get Files() As JSDataObjectFiles
Attribute Files.VB_Description = "Returns a JSDataObjectFiles collection, which in turn contains a list of all filenames used by a JSDataObject object."
    If m_oFiles Is Nothing Then
        Set m_oFiles = New JSDataObjectFiles
    End If
    Set Files = m_oFiles
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Clear()
Attribute Clear.VB_Description = "Deletes the contents of the JSDataObject object."
    Set m_cData = New Collection
    If Not m_oFiles Is Nothing Then
        m_oFiles.Clear
    End If
End Sub

Public Function GetData(ByVal Format As Integer) As Variant
Attribute GetData.VB_Description = "Returns data from a JSDataObject object in the form of a variant."
    SearchCollection m_cData, "#" & Format, RetVal:=GetData
End Function

Public Function GetFormat(ByVal Format As Integer) As Boolean
Attribute GetFormat.VB_Description = "Returns a value indicating whether an item in the JSDataObject object matches a specified format. "
    GetFormat = SearchCollection(m_cData, "#" & Format)
End Function

Public Sub SetData(Optional ByVal Value As Variant, Optional ByVal Format As Variant)
Attribute SetData.VB_Description = "Inserts data into a JSDataObject object using the specified data format."
    Dim sKey            As String

    If IsMissing(Format) Then
        sKey = "#" & jgexCFText
    Else
        sKey = "#" & CInt(Format)
    End If
    If SearchCollection(m_cData, sKey) Then
        m_cData.Remove sKey
    End If
    If Not IsMissing(Value) Then
        m_cData.Add Value, sKey
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cData = New Collection
End Sub
