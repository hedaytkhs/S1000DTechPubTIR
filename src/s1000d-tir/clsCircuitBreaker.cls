VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCircuitBreaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim CBNumber_ As String
Dim CBName_ As String
Dim CBClass_ As String
Dim Comment_ As String
Dim Source_ As String
Dim PowerSource_ As String
Dim ConnectedBUS_ As String
Dim CBPLocation_ As String
Dim System_ As String
Dim MACSDocument_ As String

Dim FoundInSMDS_ As Boolean
Dim FoundInMACS_ As Boolean
Dim RefDMFound_ As Boolean

Property Get CBNumber() As String
    CBNumber = CBNumber_
End Property

Property Let CBNumber(ByRef sArg As String)
    CBNumber_ = sArg
End Property

Property Get CBName() As String
    CBName = CBName_
End Property

Property Let CBName(ByRef sArg As String)
    CBName_ = sArg
End Property

Property Get CBClass() As String
    CBClass = CBClass_
End Property

Property Let CBClass(ByRef sArg As String)
    CBClass_ = sArg
End Property

Property Get Comment() As String
    Comment = Comment_
End Property

Property Let Comment(ByRef sArg As String)
    Comment_ = sArg
End Property

Property Get Source() As String
    Source = Source_
End Property

Property Let Source(ByRef sArg As String)
    Source_ = sArg
End Property

Property Get PowerSource() As String
    PowerSource = PowerSource_
End Property

Property Let PowerSource(ByRef sArg As String)
    PowerSource_ = sArg
End Property

Property Get ConnectedBUS() As String
    ConnectedBUS = ConnectedBUS_
End Property

Property Let ConnectedBUS(ByRef sArg As String)
    ConnectedBUS_ = sArg
End Property

Property Get CBPLocation() As String
    CBPLocation = CBPLocation_
End Property

Property Let CBPLocation(ByRef sArg As String)
    CBPLocation_ = sArg
End Property

Property Get System() As String
    System = System_
End Property

Property Let System(ByRef sArg As String)
    System_ = sArg
End Property

Property Get MACSDocument() As String
    MACSDocument = MACSDocument_
End Property

Property Let MACSDocument(ByRef sArg As String)
    MACSDocument_ = sArg
End Property

Property Get FoundInSMDS() As Boolean
    FoundInSMDS = FoundInSMDS_
End Property

Property Let FoundInSMDS(ByRef Flg As Boolean)
    FoundInSMDS_ = Flg
End Property

Property Get FoundInMACS() As Boolean
    FoundInMACS = FoundInMACS_
End Property

Property Let FoundInMACS(ByRef Flg As Boolean)
    FoundInMACS_ = Flg
End Property

Property Get RefDMFound() As Boolean
    RefDMFound = RefDMFound_
End Property

Property Let RefDMFound(ByRef Flg As Boolean)
    RefDMFound_ = Flg
End Property

Private Sub Class_Initialize()
    FoundInSMDS_ = False
    FoundInMACS_ = False
    RefDMFound_ = False
End Sub


