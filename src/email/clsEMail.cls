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
'    NOTES Mailを作成送信するためのクラス
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

'---------------------------------------------------------------------------------
' NOTESメール作成用設定
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
' タイトル
Dim Subject_ As String

' 宛先
Dim SendTo_ As String

' CC
Dim CopyTo_ As String
Dim CopyToIsEmpty_ As Boolean

' BCC
Dim BlindCopyTo_ As String
Dim BlindCopyToIsEmpty_ As Boolean


' 本文
Dim Body_ As String
' 署名
Dim Signature_ As String


' Senderは読み取り専用
'Property Get Sender() As String
''    Sender = objSender.CurrentWindowsUserName
'End Property

' タイトル
Property Let Subject(ByVal sSubject As String)
    Subject_ = sSubject
End Property

Property Get Subject() As String
    Subject = Subject_
End Property

' 宛先
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


' 本文
Property Let Body(ByVal sBody As String)
    Body_ = sBody
End Property

Property Get Body() As String
    Body = Body_
End Property


' 署名

Property Let Signature(ByVal sSignature As String)
    Signature_ = sSignature
End Property

Property Get Signature() As String
    Signature = Signature_
End Property


' ドラフト保存

'-------------------------------------------------------------------
'  Notesを起動した状態で実行
'-------------------------------------------------------------------
Public Function SaveNotesMailDraft() As Boolean
    
    Call CrateNotesMail
    
    Dim mailMode As Integer
    '-----------------------------------
    ' ***  動作モード ***
    '   保存：MAIL_SAVE
    '   送信：MAIL_SEND
    '-----------------------------------
    mailMode = MAIL_SAVE
 
    '-----------------------------------
    ' ***  保存 / 送信 ***
    '-----------------------------------
    Select Case mailMode
    Case MAIL_SEND
        wkNDoc.Send False
    Case MAIL_SAVE
        wkNDoc.Save True, False
    End Select
End Function

' Notesメール送信
Public Function SendNotesMail() As Boolean
    Call CrateNotesMail
    
    Dim mailMode As Integer
    '-----------------------------------
    ' ***  動作モード ***
    '   保存：MAIL_SAVE
    '   送信：MAIL_SEND
    '-----------------------------------
    mailMode = MAIL_SEND
 
    '-----------------------------------
    ' ***  保存 / 送信 ***
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
    ' ***  件名 ***
    '-----------------------------------
    wkNDoc.Subject = Subject_
 
    '-----------------------------------
    ' ***  宛先 ***
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
    ' ***  本文 ***
    '-----------------------------------
    Set wkNRtItem = wkNDoc.CreateRichTextItem("BODY")
    ' 本文をセットする
    With wkNRtItem
        .APPENDTEXT Body_
        .ADDNEWLINE 2
    End With

    '-----------------------------------
    ' ***  署名 ***
    '-----------------------------------
    ' 署名をセットする
    With wkNRtItem
        .APPENDTEXT Signature_
        .ADDNEWLINE 2
    End With

End Sub

' 値チェック
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

