Attribute VB_Name = "mdlNotesMail"
Option Explicit

Type tpNotesMail
    SendTo As String
    CopyTo As String
    Subject As String
    TIRTypeText As String
    IntegrationFileExists As Boolean
    IntegrationFileDate As String
    IntegrationFileName As String
    ItemCount As Long
    ReceiverFirstName As String
    SenderName As String
    IntegrationFileBaseInMAXSA As String
End Type

Private mTIRNotesMail As tpNotesMail

Private Function GetDateWithHyphen(ByVal sDate As String) As String
    GetDateWithHyphen = Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Right(sDate, 2)
End Function

Public Sub CreateNotesEMailDraft(ByRef TIRType As ESRDFileCategory)
    On Error GoTo errHandler
    Dim objMail As New clsEMail
    Dim sSendTo  As String, sCopyTo As String
    Dim iRet As Integer
    
    Dim NotesSetting As clsConfigTIRNotesSetting
    Set NotesSetting = New clsConfigTIRNotesSetting
    With mTIRNotesMail
        .SendTo = NotesSetting.SendTo
        .CopyTo = NotesSetting.CopyTo
        .ReceiverFirstName = NotesSetting.ReceiverFirstName
        .SenderName = NotesSetting.SenderFirstName
        .IntegrationFileBaseInMAXSA = NotesSetting.TIRMaxsaBaseFolder
    End With
    
    Set NotesSetting = Nothing
    
    If TIRType = Tools Then
        Dim TIRToolsSetting As clsConfigTIRTools
        Set TIRToolsSetting = New clsConfigTIRTools
        
        With mTIRNotesMail
            .IntegrationFileExists = TIRToolsSetting.IntegrationFileExists
            .IntegrationFileDate = TIRToolsSetting.IntegrationFileDate
            .TIRTypeText = "TIR-Tools Integration File"
            .Subject = "Sending " & .TIRTypeText & " " & GetDateWithHyphen(TIRToolsSetting.IntegrationFileDate)
            .IntegrationFileName = TIRToolsSetting.IntegrationFileName
            .ItemCount = TIRToolsSetting.ItemCount
        End With
        Set TIRToolsSetting = Nothing
    
    ElseIf TIRType = SUPPLIES Then
        Dim TIRSuppliesSetting As clsConfigTIRSupplies
        Set TIRSuppliesSetting = New clsConfigTIRSupplies
        With mTIRNotesMail
            .IntegrationFileExists = TIRSuppliesSetting.IntegrationFileExists
            .IntegrationFileDate = TIRSuppliesSetting.IntegrationFileDate
            .TIRTypeText = "TIR-Supplies Integration File"
            .Subject = "Sending " & .TIRTypeText & " " & GetDateWithHyphen(TIRSuppliesSetting.IntegrationFileDate)
            .IntegrationFileName = TIRSuppliesSetting.IntegrationFileName
            .ItemCount = TIRSuppliesSetting.ItemCount
        End With
        
        Set TIRSuppliesSetting = Nothing
    
    ElseIf TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        With mTIRNotesMail
            .IntegrationFileExists = TIREnterpriseSetting.IntegrationFileExists
            .IntegrationFileDate = TIREnterpriseSetting.IntegrationFileDate
            .TIRTypeText = "TIR-Enterprise Integration File"
            .Subject = "Sending " & .TIRTypeText & " " & GetDateWithHyphen(TIREnterpriseSetting.IntegrationFileDate)
            .IntegrationFileName = TIREnterpriseSetting.IntegrationFileName
            .ItemCount = TIREnterpriseSetting.ItemCount
        End With
        Set TIREnterpriseSetting = Nothing
    
    ElseIf TIRType = Zones Then
        Dim TIRZonesSetting As clsConfigTIRZones
        Set TIRZonesSetting = New clsConfigTIRZones
        
        Set TIRZonesSetting = Nothing
    
    ElseIf TIRType = AccessPoints Then
        Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
        Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
        
        Set TIRAccessPointsSetting = Nothing
    
    ElseIf TIRType = CircuitBreakers Then
        Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
        Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
        
        Set TIRCircuitBreakersSetting = Nothing
    
    End If

    Const MsgTitle As String = "TIR登録依頼メールを作成"
    If Not mTIRNotesMail.IntegrationFileExists Then
        MsgBox "Integrarion Fileが存在しません.", vbExclamation + vbOKOnly, MsgTitle
        Exit Sub
    ElseIf mTIRNotesMail.ItemCount < 1 Then
        MsgBox "登録アイテムが存在しません.", vbExclamation + vbOKOnly, MsgTitle
        Exit Sub
    Else
        iRet = MsgBox(mTIRNotesMail.IntegrationFileName & " の登録依頼メールを作成しますか？", vbInformation + vbYesNoCancel, MsgTitle)
        If iRet <> vbYes Then
            Exit Sub
        End If
    End If

    With objMail
        
        '-----------------------------------
        ' ***  宛先 ***
        '-----------------------------------
        .SendTo = mTIRNotesMail.SendTo
 
        '-----------------------------------
        ' ***  CC ***
        '-----------------------------------
        .CopyTo = mTIRNotesMail.CopyTo
        
        '-----------------------------------
        ' ***  BCC ***
        '-----------------------------------
        ' NOT USED FOR EXTRA ICN INFORMATION MAIL TO CE
'        .BlindCopyTo=""
        
        '-----------------------------------
        ' ***  件名 ***
        '-----------------------------------
        .Subject = mTIRNotesMail.Subject


        '-----------------------------------
        ' ***  冒頭句 ***
        '-----------------------------------
'        .BeginningSentence = "***  冒頭句 ***" & vbCrLf & vbCrLf


        '-----------------------------------
        ' ***  本文 ***
        '-----------------------------------
        .Body = GetBodyText(mTIRNotesMail)

        '-----------------------------------
        ' ***  署名 ***
        '-----------------------------------
        .Signature = "Best regards," & vbCrLf & mTIRNotesMail.SenderName

        
        '-----------------------------------
        ' ***  ファイル添付 ***
        '-----------------------------------

        
        '-----------------------------------
        ' ***  生成テキスト挿入 ***
        '-----------------------------------
        
        '-----------------------------------
        ' ***  動作モード ***
        '   保存：SaveNotesMailDraft
        '   送信：SendNotesMail
        '-----------------------------------
        .SaveNotesMailDraft
    End With
    
    MsgBox "TIR 登録依頼メールをNotesのドラフトフォルダ内に作成しました.", vbInformation + vbOKOnly, "TIR 登録依頼メール"

    Set objMail = Nothing
    Exit Sub
errHandler:
    Set objMail = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function GetBodyText(ByRef TIRNotesMail As tpNotesMail) As String
GetBodyText = "Dear " & TIRNotesMail.ReceiverFirstName & "-san," & vbCrLf & vbCrLf & _
                "MITAC would like to ask you import a " & TIRNotesMail.TIRTypeText & " stored in Maxsa G:." & vbCrLf & _
                "Please import this TIR integration file into both SMDS Production and Test." & vbCrLf & vbCrLf & _
                StrAddPathSeparator(TIRNotesMail.IntegrationFileBaseInMAXSA) & TIRNotesMail.IntegrationFileDate & "\" & TIRNotesMail.IntegrationFileName & vbCrLf & vbCrLf & _
                TIRNotesMail.TIRTypeText & vbCrLf & _
                "Number of item : " & CStr(TIRNotesMail.ItemCount) & vbCrLf & vbCrLf & _
                "Prerequisite for importing this TIR: No" & vbCrLf & vbCrLf & _
                "Prerequisite for importing this TIR: Yes" & vbCrLf & _
                "NOTE: CAGE Code XXXXX should be registered before importing this " & TIRNotesMail.TIRTypeText & " Integration file." & vbCrLf & _
                "Please import Integration_DB_TIR_Enterprise_YYYYMMDD_HHMM.csv before importing this integration file." & vbCrLf & vbCrLf & _
                "If there is any problems or comments, feel free to ask me." & vbCrLf
End Function


