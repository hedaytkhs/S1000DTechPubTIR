VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsS1000DIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'
' MRJ BRÇÃíËã`Ç…è]Ç¢ Issue/Inwork NumberÇèàóùÇ∑ÇÈÉNÉâÉX
'
'
' Hideaki Takahashi
' ///// 2014/03/18 /////
'
'
'
'=========================================================

Option Explicit

Dim inWork_ As String * 2
Dim issueNumber_ As String * 3
'Dim issueType_ As S1000D_IssueType
Dim WorkAndIssueNumber_ As String * 6

Property Let WorkAndIssueNumber(ByVal sWorkAndIssueNumber As String)
    WorkAndIssueNumber_ = Left(sWorkAndIssueNumber, 6)
    Call DevideWorkAndIssueNumber
End Property

Property Get inWork() As String
    inWork = inWork_
End Property

Property Get WorkAndIssueNumber() As String
    WorkAndIssueNumber = WorkAndIssueNumber_
End Property

Property Get issueNumber() As String
    issueNumber = issueNumber_
End Property

Private Sub Class_Initialize()
    inWork_ = "01"
    issueNumber_ = "000"
    Call GetWorkAndIssueNumber
End Sub

Private Sub GetWorkAndIssueNumber()
    WorkAndIssueNumber_ = issueNumber_ & "-" & inWork_
End Sub

Private Sub DevideWorkAndIssueNumber()
    issueNumber_ = Left(WorkAndIssueNumber_, 3)
    inWork_ = Mid(WorkAndIssueNumber_, 5, 2)
End Sub

