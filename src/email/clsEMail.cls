VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    NOTES Mail���쐬���M���邽�߂̃N���X
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

'---------------------------------------------------------------------------------
' NOTES���[���쐬�p�ݒ�
'---------------------------------------------------------------------------------
Const EMBED_ATTACHMENT As Integer = 1454
Const MAIL_SEND = 1
Const MAIL_SAVE = 2

Dim wkNSes As Object    ' lotus.NOTESSESSION
Dim wkNDB As Object     ' lotus.NOTESDATABASE

Dim wkNDoc As Object    ' lotus.NOTESDOCUMENT
Dim wkNRtItem As Object ' lotus.NOTESRICHTEXTITEM
Dim wkNAtt As Object    ' lotus.NOTESEMBEDDEDOBJECT

'---------------------------------------------------------------------------------
' �^�C�g��
Dim Subject_ As String

' ����
Dim SendTo_ As String

' CC
Dim CopyTo_ As String
Dim CopyToIsEmpty_ As Boolean

' BCC
Dim BlindCopyTo_ As String
Dim BlindCopyToIsEmpty_ As Boolean


' �{��
Dim Body_ As String
' ����
Dim Signature_ As String


' Sender�͓ǂݎ���p
'Property Get Sender() As String
''    Sender = objSender.CurrentWindowsUserName
'End Property

' �^�C�g��
Property Let Subject(ByVal sSubject As String)
    Subject_ = sSubject
End Property

Property Get Subject() As String
    Subject = Subject_
End Property

' ����
Property Let SendTo(ByVal sSendTo As String)
    SendTo_ = sSendTo
End Property

Property Get SendTo() As String
    SendTo = SendTo_
End Property

' CC
Property Let CopyTo(ByVal sCopyTo As String)
    CopyTo_ = sCopyTo
End Property

Property Get CopyTo() As String
    CopyTo = CopyTo_
End Property

Property Get CopyToIsEmpty() As Boolean
    CopyToIsEmpty = CopyToIsEmpty_
End Property


' BCC
Property Let BlindCopyTo(ByVal sBlindCopyTo As String)
    BlindCopyTo_ = sBlindCopyTo
    Call CheckBlindCopyTo
End Property

Property Get BlindCopyTo() As String
    BlindCopyTo = BlindCopyTo_
End Property

Property Get BlindCopyToIsEmpty() As Boolean
    BlindCopyToIsEmpty = BlindCopyToIsEmpty_
End Property


' �{��
Property Let Body(ByVal sBody As String)
    Body_ = sBody
End Property

Property Get Body() As String
    Body = Body_
End Property


' ����

Property Let Signature(ByVal sSignature As String)
    Signature_ = sSignature
End Property

Property Get Signature() As String
    Signature = Signature_
End Property


' �h���t�g�ۑ�

'-------------------------------------------------------------------
'  Notes���N��������ԂŎ��s
'-------------------------------------------------------------------
Public Function SaveNotesMailDraft() As Boolean
    
    Call CrateNotesMail
    
    Dim mailMode As Integer
    '-----------------------------------
    ' ***  ���샂�[�h ***
    '   �ۑ��FMAIL_SAVE
    '   ���M�FMAIL_SEND
    '-----------------------------------
    mailMode = MAIL_SAVE
 
    '-----------------------------------
    ' ***  �ۑ� / ���M ***
    '-----------------------------------
    Select Case mailMode
    Case MAIL_SEND
        wkNDoc.Send False
    Case MAIL_SAVE
        wkNDoc.Save True, False
    End Select
End Function

' Notes���[�����M
Public Function SendNotesMail() As Boolean
    Call CrateNotesMail
    
    Dim mailMode As Integer
    '-----------------------------------
    ' ***  ���샂�[�h ***
    '   �ۑ��FMAIL_SAVE
    '   ���M�FMAIL_SEND
    '-----------------------------------
    mailMode = MAIL_SEND
 
    '-----------------------------------
    ' ***  �ۑ� / ���M ***
    '-----------------------------------
    Select Case mailMode
    Case MAIL_SEND
        wkNDoc.Send False
    Case MAIL_SAVE
        wkNDoc.Save True, False
    End Select
End Function

Private Sub CrateNotesMail()
Dim AttFName As String
        
    wkNDB.OpenMail
    
    Set wkNDoc = wkNDB.CreateDocument()

    '-----------------------------------
    ' ***  ���� ***
    '-----------------------------------
    wkNDoc.Subject = Subject_
 
    '-----------------------------------
    ' ***  ���� ***
    '-----------------------------------
    wkNDoc.SendTo = SendTo_
 
    '-----------------------------------
    ' ***  CC ***
    '-----------------------------------
    If Not CopyToIsEmpty_ Then
        wkNDoc.CopyTo = CopyTo_
    End If
     
    '-----------------------------------
    ' ***  BCC ***
    '-----------------------------------
    If Not BlindCopyToIsEmpty_ Then
        wkNDoc.BlindCopyTo = BlindCopyTo_
    End If

    '-----------------------------------
    ' ***  �{�� ***
    '-----------------------------------
    Set wkNRtItem = wkNDoc.CreateRichTextItem("BODY")
    ' �{�����Z�b�g����
    With wkNRtItem
        .APPENDTEXT Body_
        .ADDNEWLINE 2
    End With

    '-----------------------------------
    ' ***  ���� ***
    '-----------------------------------
    ' �������Z�b�g����
    With wkNRtItem
        .APPENDTEXT Signature_
        .ADDNEWLINE 2
    End With

End Sub

' �l�`�F�b�N
Property Get CheckCopyTo() As Boolean
    If CopyTo_ <> "" Then
        CopyToIsEmpty_ = False
    Else
        CopyToIsEmpty_ = True
    End If
End Property

Private Sub CheckBlindCopyTo()
    If BlindCopyTo_ <> "" Then
        BlindCopyToIsEmpty_ = False
    Else
        BlindCopyToIsEmpty_ = True
    End If
End Sub


Private Sub Class_Initialize()
    Set wkNSes = CreateObject("Notes.NotesSession")
    Set wkNDB = wkNSes.GetDatabase("", "")
    
    Subject_ = ""
    SendTo_ = ""
    CopyTo_ = ""
    BlindCopyTo_ = ""
    Body_ = ""
    Signature_ = ""
    
End Sub
