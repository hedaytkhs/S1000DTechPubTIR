VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim MiTNumber_ As String
Dim ToolNumber_ As String
Dim ToolName_ As String
Dim ToolLongName_ As String
Dim ManufactureCode_ As String
Dim Comment_ As String
Dim Source_ As String
Dim Remarks_ As String

Dim FoundInSMDS_ As Boolean
Dim FoundInMACS_ As Boolean
Dim RefDMFound_ As Boolean

Dim ActiveItem_ As Boolean
Dim RevisionSequence_ As Long
Dim ShortName_ As String


Property Get MiTNumber() As String
    MiTNumber = MiTNumber_
End Property

Property Let MiTNumber(ByRef sArg As String)
    MiTNumber_ = sArg
End Property

Property Get ToolNumber() As String
    ToolNumber = ToolNumber_
End Property

Property Let ToolNumber(ByRef sArg As String)
    ToolNumber_ = sArg
End Property

Property Get ToolName() As String
    ToolName = ToolName_
End Property

Property Let ToolName(ByRef sArg As String)
    ToolName_ = sArg
End Property

Property Get ToolLongName() As String
    ToolLongName = ToolLongName_
End Property

Property Let ToolLongName(ByRef sArg As String)
    ToolLongName_ = sArg
End Property

Property Get ManufactureCode() As String
    ManufactureCode = ManufactureCode_
End Property

Property Let ManufactureCode(ByRef sArg As String)
    ManufactureCode_ = sArg
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

Property Get Remarks() As String
    Remarks = Remarks_
End Property

Property Let Remarks(ByRef sArg As String)
    Remarks_ = sArg
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

Property Get ActiveItem() As Boolean
    ActiveItem = ActiveItem_
End Property

Property Let ActiveItem(ByRef flgArg As Boolean)
    ActiveItem_ = flgArg
End Property

Property Get RevisionSequence() As Long
    RevisionSequence = RevisionSequence_
End Property

Property Let RevisionSequence(ByRef Arg As Long)
    RevisionSequence_ = Arg
End Property

Property Get ShortName() As String
    ShortName = ShortName_
End Property

Property Let ShortName(ByRef sArg As String)
    ShortName_ = sArg
End Property

Private Sub Class_Initialize()
    FoundInSMDS_ = False
    FoundInMACS_ = False
    RefDMFound_ = False
    ActiveItem_ = False
End Sub



