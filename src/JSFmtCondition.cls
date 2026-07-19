VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFmtCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains the criteria that rows must meet to appear in a different format."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains the criteria that rows must meet to appear in a different format
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oFormatStyle              As JSFormatStyle
Private m_nColIndex                 As Integer
Private m_eOperator                 As jgexConditionOperatorConstants
Private m_vValue1                   As Variant
Private m_vValue2                   As Variant
Private m_nIndex                    As Integer
Private m_sKey                      As String

'=========================================================================
' Properties
'=========================================================================

Public Property Get FormatStyle() As JSFormatStyle
Attribute FormatStyle.VB_Description = "Returns the JSFormatStyle object for the JSFmtCondition."
    If m_oFormatStyle Is Nothing Then
        Set m_oFormatStyle = New JSFormatStyle
        m_oFormatStyle.frInit vbNullString
    End If
    Set FormatStyle = m_oFormatStyle
End Property

Public Property Get ColIndex() As Integer
Attribute ColIndex.VB_Description = "Returns/sets the index of the column  to be compared."
    ColIndex = m_nColIndex
End Property

Public Property Let ColIndex(ByVal nValue As Integer)
    m_nColIndex = nValue
End Property

Public Property Get Operator() As jgexConditionOperatorConstants
Attribute Operator.VB_Description = "Returns/sets the operator used for comparison."
    Operator = m_eOperator
End Property

Public Property Let Operator(ByVal eValue As jgexConditionOperatorConstants)
    m_eOperator = eValue
End Property

Public Property Get Value1() As Variant
Attribute Value1.VB_Description = "Returns/sets the value that is compared with column values."
    Value1 = m_vValue1
End Property

Public Property Let Value1(ByVal vntValue As Variant)
    m_vValue1 = vntValue
End Property

Public Property Get Value2() As Variant
Attribute Value2.VB_Description = "Returns/sets the value that is compared with column values."
    Value2 = m_vValue2
End Property

Public Property Let Value2(ByVal vntValue As Variant)
    m_vValue2 = vntValue
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
    Index = m_nIndex
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a string that uniquely identifies a member in a collection."
    Key = m_sKey
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub SetCondition(ByVal ColIndex As Integer, ByVal Operator As jgexConditionOperatorConstants, ByVal Value1 As Variant, Optional ByVal Value2 As Variant)
Attribute SetCondition.VB_Description = "Sets the condition parameters."
    m_nColIndex = ColIndex
    m_eOperator = Operator
    m_vValue1 = Value1
    If Not IsMissing(Value2) Then
        m_vValue2 = Value2
    End If
End Sub

Friend Sub frInit(ByVal nIndex As Integer, sKey As String)
    m_nIndex = nIndex
    m_sKey = sKey
End Sub

Friend Sub frSetIndex(ByVal nIndex As Integer)
    m_nIndex = nIndex
End Sub
