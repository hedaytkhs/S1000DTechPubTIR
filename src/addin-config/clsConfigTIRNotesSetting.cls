VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigTIRNotesSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'8 óÒñ⁄Ç…TIR NotesÉÅÅ[Éãê›íËÇï€ë∂
Private Const XlsColNotesMail As Long = 8

Private Enum XlsCellSendTo
    Row = 1
    col = XlsColNotesMail
End Enum

Private Enum XlsCellCopyTo
    Row = 2
    col = XlsColNotesMail
End Enum

Private Enum XlsCellReceiverFirstName
    Row = 3
    col = XlsColNotesMail
End Enum

Private Enum XlsCellSenderFirstName
    Row = 4
    col = XlsColNotesMail
End Enum

Private Enum XlsCellIntegrationFileBaseFolder
    Row = 5
    col = XlsColNotesMail
End Enum

Private Enum XlsCellTIRMaxsaBaseFolder
    Row = 6
    col = XlsColNotesMail
End Enum


Dim SendTo_ As String
Dim CopyTo_ As String
Dim ReceiverFirstName_ As String
Dim SenderFirstName_ As String
Dim IntegrationFileBaseFolder_ As String
Dim TIRMaxsaBaseFolder_ As String

Dim ShtConfig As Worksheet

Property Let SendTo(ByVal sSendTo As String)
    SendTo_ = sSendTo
    ShtConfig.Cells(XlsCellSendTo.Row, XlsCellSendTo.col).Value = SendTo_
End Property

Property Get SendTo() As String
    SendTo = ShtConfig.Cells(XlsCellSendTo.Row, XlsCellSendTo.col).Value
End Property

Property Get CopyTo() As String
    CopyTo = ShtConfig.Cells(XlsCellCopyTo.Row, XlsCellCopyTo.col).Value
End Property

Property Let CopyTo(ByVal sCopyTo As String)
    CopyTo_ = sCopyTo
    ShtConfig.Cells(XlsCellCopyTo.Row, XlsCellCopyTo.col).Value = CopyTo_
End Property

Property Get ReceiverFirstName() As String
    ReceiverFirstName = ShtConfig.Cells(XlsCellReceiverFirstName.Row, XlsCellReceiverFirstName.col).Value
End Property

Property Let ReceiverFirstName(ByVal sName As String)
    ReceiverFirstName_ = sName
    ShtConfig.Cells(XlsCellReceiverFirstName.Row, XlsCellReceiverFirstName.col).Value = ReceiverFirstName_
End Property


Property Get SenderFirstName() As String
    SenderFirstName = ShtConfig.Cells(XlsCellSenderFirstName.Row, XlsCellSenderFirstName.col).Value
End Property

Property Let SenderFirstName(ByVal sName As String)
    SenderFirstName_ = sName
    ShtConfig.Cells(XlsCellSenderFirstName.Row, XlsCellSenderFirstName.col).Value = SenderFirstName_
End Property


Property Get TIRMaxsaBaseFolder() As String
    TIRMaxsaBaseFolder = ShtConfig.Cells(XlsCellTIRMaxsaBaseFolder.Row, XlsCellTIRMaxsaBaseFolder.col).Value
End Property

Property Let TIRMaxsaBaseFolder(ByVal sName As String)
    TIRMaxsaBaseFolder_ = sName
    ShtConfig.Cells(XlsCellTIRMaxsaBaseFolder.Row, XlsCellTIRMaxsaBaseFolder.col).Value = TIRMaxsaBaseFolder_
End Property

Property Get IntegrationFileBaseFolder() As String
    IntegrationFileBaseFolder = ShtConfig.Cells(XlsCellIntegrationFileBaseFolder.Row, XlsCellIntegrationFileBaseFolder.col).Value
End Property

Property Let IntegrationFileBaseFolder(ByVal sName As String)
    IntegrationFileBaseFolder_ = sName
    ShtConfig.Cells(XlsCellIntegrationFileBaseFolder.Row, XlsCellIntegrationFileBaseFolder.col).Value = IntegrationFileBaseFolder_
End Property

Function Save() As Boolean
    ThisWorkbook.Save
End Function

Private Sub Class_Initialize()
    Set ShtConfig = ThisWorkbook.Worksheets(CNST_XLS_SHEET_CONFIG)
    With ShtConfig
        SendTo_ = .Cells(XlsCellSendTo.Row, XlsCellSendTo.col).Value
        CopyTo_ = .Cells(XlsCellCopyTo.Row, XlsCellCopyTo.col).Value
        ReceiverFirstName_ = .Cells(XlsCellReceiverFirstName.Row, XlsCellReceiverFirstName.col).Value
        SenderFirstName_ = .Cells(XlsCellSenderFirstName.Row, XlsCellSenderFirstName.col).Value
        IntegrationFileBaseFolder_ = .Cells(XlsCellIntegrationFileBaseFolder.Row, XlsCellIntegrationFileBaseFolder.col).Value
        TIRMaxsaBaseFolder_ = ShtConfig.Cells(XlsCellTIRMaxsaBaseFolder.Row, XlsCellTIRMaxsaBaseFolder.col).Value
    End With
End Sub

Private Sub Class_Terminate()
    Set ShtConfig = Nothing
End Sub




