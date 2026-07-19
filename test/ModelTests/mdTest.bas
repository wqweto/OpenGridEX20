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
    If vExpected = vActual Then
        m_lPassed = m_lPassed + 1
        Print #m_lFile, "PASS " & sName
    Else
        m_lFailed = m_lFailed + 1
        Print #m_lFile, "FAIL " & sName & " expected=<" & vExpected & "> actual=<" & vActual & ">"
    End If
End Sub

Public Sub TestsDone()
    If m_lFailed = 0 Then
        Print #m_lFile, "RESULT: PASSED (" & m_lPassed & " tests)"
    Else
        Print #m_lFile, "RESULT: FAILED (" & m_lFailed & " of " & m_lPassed + m_lFailed & " tests)"
    End If
    Close #m_lFile
End Sub
