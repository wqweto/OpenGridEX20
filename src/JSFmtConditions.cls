VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFmtConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of FmtCondition objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of FmtCondition objects
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cItems                    As Collection
Private m_oGroupCondition           As JSFmtCondition
Private m_bApplyGroupCondition      As Boolean
Private m_sGroupConditionCountTitle As String
Private m_bShowGroupConditionCount  As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(ByVal Index As Variant) As JSFmtCondition
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSFmtCondition of  the collection either by index or by key."
    Set Item = m_cItems.Item(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
    Count = m_cItems.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Public Property Get GroupCondition() As JSFmtCondition
Attribute GroupCondition.VB_Description = "Returns a JSFmtCondition object that is used through the group rows."
    If m_oGroupCondition Is Nothing Then
        Set m_oGroupCondition = New JSFmtCondition
        m_oGroupCondition.frInit 0, vbNullString
    End If
    Set GroupCondition = m_oGroupCondition
End Property

Public Property Get ApplyGroupCondition() As Boolean
Attribute ApplyGroupCondition.VB_Description = "Controls whether the GroupCondition property is valid."
    ApplyGroupCondition = m_bApplyGroupCondition
End Property

Public Property Let ApplyGroupCondition(ByVal bValue As Boolean)
    m_bApplyGroupCondition = bValue
End Property

Public Property Get GroupConditionCountTitle() As String
Attribute GroupConditionCountTitle.VB_Description = "Returns/sets the text displayed in a group row when one or more rows in the group meet the criteria specified in the JSGroupCondition object's properties."
    GroupConditionCountTitle = m_sGroupConditionCountTitle
End Property

Public Property Let GroupConditionCountTitle(ByVal sValue As String)
    m_sGroupConditionCountTitle = sValue
End Property

Public Property Get ShowGroupConditionCount() As Boolean
Attribute ShowGroupConditionCount.VB_Description = "Controls whether the count of rows, that meet the criteria in GroupCondition property settings, will be displayed in a group row."
    ShowGroupConditionCount = m_bShowGroupConditionCount
End Property

Public Property Let ShowGroupConditionCount(ByVal bValue As Boolean)
    m_bShowGroupConditionCount = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add(ByVal ColIndex As Integer, Optional ByVal Operator As jgexConditionOperatorConstants = jgexEqual, Optional ByVal Value1 As Variant, Optional ByVal Value2 As Variant, Optional ByVal Key As String) As JSFmtCondition
Attribute Add.VB_Description = "Adds a JSFmtCondition object to the collection and returns a reference to the newly created object."
    Set Add = New JSFmtCondition
    Add.frInit m_cItems.Count + 1, Key
    If IsMissing(Value1) Then
        Value1 = Empty
    End If
    Add.SetCondition ColIndex, Operator, Value1, Value2
    If LenB(Key) <> 0 Then
        m_cItems.Add Add, Key
    Else
        m_cItems.Add Add
    End If
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from the JSFmtConditions collection."
    m_cItems.Remove Index
    pvReindex
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
    Set m_cItems = New Collection
End Sub

Private Sub pvReindex()
    Dim lIdx            As Long
    Dim oItem           As JSFmtCondition

    For lIdx = 1 To m_cItems.Count
        Set oItem = m_cItems.Item(lIdx)
        oItem.frSetIndex lIdx
    Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cItems = New Collection
End Sub
