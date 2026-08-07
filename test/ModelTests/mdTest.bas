Attribute VB_Name = "mdTest"
'=========================================================================
'
' Open GridEX 2000 Control
' Minimal test assertion helpers writing results to a log file
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lFile                     As Long
Private m_lPassed                   As Long
Private m_lFailed                   As Long

'=========================================================================
' Functions
'=========================================================================

Public Sub TestInit(sFile As String)
    m_lFile = FreeFile
    Open sFile For Output As #m_lFile
End Sub

Public Sub Assert(sName As String, ByVal bCond As Boolean)
    If bCond Then
        m_lPassed = m_lPassed + 1
        Print #m_lFile, "PASS " & sName
    Else
        m_lFailed = m_lFailed + 1
        Print #m_lFile, "FAIL " & sName
    End If
End Sub

Public Sub AssertEquals(sName As String, vExpected As Variant, vActual As Variant)
    Dim bEqual          As Boolean

    '--- the comparison stays the lenient Variant =, which is what lets a
    '--- Long assert against an Integer property -- but a value it cannot
    '--- compare (a Null, an array) has to fail this one assert rather than
    '--- raise through Form_Load and silently skip every test after it
    On Error GoTo EH
    bEqual = (vExpected = vActual)
    If bEqual Then
        m_lPassed = m_lPassed + 1
        Print #m_lFile, "PASS " & sName
    Else
        m_lFailed = m_lFailed + 1
        Print #m_lFile, "FAIL " & sName & " expected=<" & pvToText(vExpected) & "> actual=<" & pvToText(vActual) & ">"
    End If
    Exit Sub
EH:
    m_lFailed = m_lFailed + 1
    Print #m_lFile, "FAIL " & sName & " not comparable: expected=<" & pvToText(vExpected) & "> actual=<" & pvToText(vActual) & ">"
End Sub

Private Function pvToText(vValue As Variant) As String
    On Error GoTo EH
    If IsObject(vValue) Then
        pvToText = "#object"
    ElseIf IsArray(vValue) Then
        pvToText = "#array"
    ElseIf IsNull(vValue) Then
        pvToText = "#null"
    Else
        pvToText = vValue & vbNullString
    End If
    Exit Function
EH:
    pvToText = "#?"
End Function

Public Sub TestsDone()
    If m_lFailed = 0 Then
        Print #m_lFile, "RESULT: PASSED (" & m_lPassed & " tests)"
    Else
        Print #m_lFile, "RESULT: FAILED (" & m_lFailed & " of " & m_lPassed + m_lFailed & " tests)"
    End If
    Close #m_lFile
End Sub

Public Sub TestSkip(sReason As String)
    Print #m_lFile, "RESULT: PASSED (0 tests, skipped: " & sReason & ")"
    Close #m_lFile
End Sub
