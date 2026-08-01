VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim mLastName As String
Dim mFirstName As String
Dim mAge As Integer



Public Property Let Age(ByVal RHS As Integer)

    mAge = RHS
    
End Property

Public Property Get Age() As Integer

    Age = mAge
    
End Property

Public Property Let FirstName(ByVal RHS As String)

    mFirstName = RHS
    
End Property

Public Property Get FirstName() As String

    FirstName = mFirstName

End Property




Public Property Get LastName() As String

    LastName = mLastName

End Property

Public Property Let LastName(ByVal RHS As String)

    mLastName = RHS
    
End Property

