VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccessPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim PanelNumber_ As String
Dim Description_ As String
Dim FoundInSMDS_ As Boolean
Dim RefDMFound_ As Boolean

Property Get PanelNumber() As String
    PanelNumber = PanelNumber_
End Property

Property Let PanelNumber(ByRef sArg As String)
    PanelNumber_ = sArg
End Property

Property Get Description() As String
    Description = Description_
End Property

Property Let Description(ByRef sArg As String)
    Description_ = sArg
End Property

Property Get FoundInSMDS() As Boolean
    FoundInSMDS = FoundInSMDS_
End Property

Property Let FoundInSMDS(ByRef Flg As Boolean)
    FoundInSMDS_ = Flg
End Property

Property Get RefDMFound() As Boolean
    RefDMFound = RefDMFound_
End Property

Property Let RefDMFound(ByRef Flg As Boolean)
    RefDMFound_ = Flg
End Property

Private Sub Class_Initialize()
    FoundInSMDS_ = False
    RefDMFound_ = False
End Sub

