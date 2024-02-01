VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    DMCを操作するためのクラス
'
'
'    DMCの要素分解、結合、値チェック
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

' DMC全体
' 2014/03/25
' DMCにURNを含める
' Inwork-Issue 及びLanguage-Country拡張文字列に対応させる
Dim DMC_ As String

' DMCの各要素
'Dim URN_ As String * 3
Dim ModelIdentCode_ As String
Dim systemDiffCode_ As String * 1
Dim SystemCode_ As String * 2
Dim subSystemCode_ As String * 1
Dim subSubSystemCode_ As String * 1
Dim assyCode_ As String * 2
Dim disassyCode_ As String * 2
Dim disassyCodeVariant_ As String * 1
Dim infoCode_ As String * 3
Dim infoCodeVariant_ As String * 1
Dim itemLocationCode_ As String * 1
' SNS 8ケタ
Dim SNS_ As String * 8
Dim IsValidDMC_ As Boolean

Dim WorkAndIssueNumber_ As String * 6
Dim LanguageAndCountry_ As String * 5
Dim TechName_ As String
Dim InfoName_ As String

Const cDMCSeparator = "-"

Property Let DMC(ByVal sDMC As String)
    DMC_ = sDMC
    Call DivedeDMC
    Call UpdateDMC
End Property

Property Let ModelIdentCode(ByVal sModelIdentCode As String)
    Select Case sModelIdentCode
        Case "MRJ", "PW1000G"
            IsValidDMC_ = True
'        Exit Function
        Case "MWG", "MBR", "S1000D"
            IsValidDMC_ = False
        Case Else
            IsValidDMC_ = False
    End Select
    
    ModelIdentCode_ = sModelIdentCode
    Call UpdateDMC
End Property

Property Let systemDiffCode(ByVal ssystemDiffCode As String)
    If Not ssystemDiffCode Like "[a-zA-Z]" Then IsValidDMC_ = False
    systemDiffCode_ = Left(ssystemDiffCode, 1)
    Call UpdateDMC
End Property

Property Let SystemCode(ByVal sSystemCode As String)
    If Not IsNumeric(sSystemCode) Then
        IsValidDMC_ = False
    End If
    SystemCode_ = Left(sSystemCode, 2)
    Call UpdateSNS
    Call UpdateDMC
End Property
    
Property Let subSystemCode(ByVal ssubSystemCode As String)
    If Not IsNumeric(ssubSystemCode) Then
        IsValidDMC_ = False
    End If
    subSystemCode_ = Left(ssubSystemCode, 1)
    Call UpdateSNS
    Call UpdateDMC
End Property
    
Property Let subSubSystemCode(ByVal ssubSubSystemCode As String)
    If Not IsNumeric(ssubSubSystemCode) Then
        IsValidDMC_ = False
    End If
    subSubSystemCode_ = Left(ssubSubSystemCode, 1)
    Call UpdateSNS
    Call UpdateDMC
End Property
    
Property Let assyCode(ByVal sassyCode As String)
    If Not IsNumeric(sassyCode) Then
        IsValidDMC_ = False
    End If
    assyCode_ = Left(sassyCode, 2)
    Call UpdateSNS
    Call UpdateDMC
End Property
    
Property Let disassyCode(ByVal sdisassyCode As String)
    If Not IsNumeric(sdisassyCode) Then
        IsValidDMC_ = False
    End If
    disassyCode_ = Left(sdisassyCode, 2)
    Call UpdateDMC
End Property
    
Property Let disassyCodeVariant(ByVal sdisassyCodeVariant As String)
    If Not sdisassyCodeVariant Like "[a-zA-Z]" Then IsValidDMC_ = False
    disassyCodeVariant_ = Left(sdisassyCodeVariant, 1)
    Call UpdateDMC
End Property

Property Let infoCode(ByVal sinfoCode As String)
    If Not IsNumeric(sinfoCode) Then
        IsValidDMC_ = False
    End If
    infoCode_ = Left(sinfoCode, 3)
    Call UpdateDMC
End Property

Property Let infoCodeVariant(ByVal sinfoCodeVariant As String)
    If Not sinfoCodeVariant Like "[a-zA-Z]" Then IsValidDMC_ = False
    infoCodeVariant_ = Left(sinfoCodeVariant, 1)
    Call UpdateDMC
End Property

Property Let itemLocationCode(ByVal sitemLocationCode As String)
    If Not sitemLocationCode Like "[a-zA-Z]" Then IsValidDMC_ = False
    itemLocationCode_ = Left(sitemLocationCode, 1)
    Call UpdateDMC
End Property

Property Let TechName(ByVal ArgText As String)
    TechName_ = ArgText
End Property

Property Let InfoName(ByVal ArgText As String)
    InfoName_ = ArgText
End Property

Property Get DMC() As String
    DMC = DMC_
End Property

Property Get ModelIdentCode() As String
    ModelIdentCode = ModelIdentCode_
End Property

Property Get systemDiffCode() As String
    systemDiffCode = systemDiffCode_
End Property

Property Get SystemCode() As String
    SystemCode = SystemCode_
End Property

Property Get subSystemCode() As String
    subSystemCode = subSystemCode_
End Property

Property Get subSubSystemCode() As String
    subSubSystemCode = subSubSystemCode_
End Property

Property Get assyCode() As String
    assyCode = assyCode_
End Property

Property Get SNS() As String
    SNS = SNS_
End Property

Property Get disassyCode() As String
    disassyCode = disassyCode_
End Property

Property Get disassyCodeVariant() As String
    disassyCodeVariant = disassyCodeVariant_
End Property

Property Get infoCode() As String
    infoCode = infoCode_
End Property

Property Get infoCodeVariant() As String
    infoCodeVariant = infoCodeVariant_
End Property

Property Get itemLocationCode() As String
    itemLocationCode = itemLocationCode_
End Property

Property Get WorkAndIssueNumber() As String
    WorkAndIssueNumber = WorkAndIssueNumber_
End Property

Property Get LangueAndCountry() As String
    LangueAndCountry = LanguageAndCountry_
End Property

Property Get IsValidDMC() As Boolean
    IsValidDMC = IsValidDMC_
End Property

Property Get TechName() As String
     TechName = TechName_
End Property

Property Get InfoName() As String
     InfoName = InfoName_
End Property


Private Sub DivedeDMC()
    Dim vntDMCPart As Variant
    Dim tmp As Variant
    
    If InStr(DMC_, "-") = 0 Then
        IsValidDMC_ = False
        Exit Sub
    End If
    
    ' 区切り文字"_"でDMCとExtensionを分離
    vntDMCPart = Split("" & DMC_, "_")
    
    'vntDMCPart(0): DMC
    'vntDMCPart(1): Extension(Inwork-IssueNo)
    'vntDMCPart(2): Extension(Langue-Country)
    
    tmp = Split("" & vntDMCPart(0), "-")
    If UBound(tmp) < 7 Then
        IsValidDMC_ = False
        Exit Sub
    End If
    Dim offset As Long
    offset = UBound(tmp) - 7
    
    ModelIdentCode_ = "" & tmp(offset + 0)
    systemDiffCode_ = Left("" & tmp(offset + 1), 1)
    SystemCode_ = Left("" & tmp(offset + 2), 2)
    
    subSystemCode_ = Left("" & tmp(offset + 3), 1)
    subSubSystemCode_ = Mid("" & tmp(offset + 3), 2, 1)
    
    assyCode_ = Left("" & tmp(offset + 4), 2)
    Call UpdateSNS
    
    disassyCode_ = Left("" & tmp(offset + 5), 2)
    disassyCodeVariant_ = Mid("" & tmp(offset + 5), 3, 1)
    
    infoCode_ = Left("" & tmp(offset + 6), 3)
    infoCodeVariant_ = Mid("" & tmp(offset + 6), 4, 1)
    
    itemLocationCode_ = Left("" & tmp(offset + 7), 1)
    
    If UBound(vntDMCPart) > 0 Then
        WorkAndIssueNumber_ = Left("" & vntDMCPart(1), 6)
        If UBound(vntDMCPart) > 1 Then
            LanguageAndCountry_ = Left("" & vntDMCPart(2), 5)
        End If
    End If
End Sub

Private Sub Class_Initialize()
    
'    'URNの定義から"DMC"を取得
'    Dim mURN As New clsS1000DURN
'    mURN.URNCategory = S1000D_URN.DMC
'    URN_ = mURN.URN
    
    ModelIdentCode_ = "MRJ"
    systemDiffCode_ = "A"
    SystemCode_ = "00"
    subSystemCode_ = "0"
    subSubSystemCode_ = "0"
    assyCode_ = "00"
    disassyCode_ = "00"
    disassyCodeVariant_ = "A"
    infoCode_ = "000"
    infoCodeVariant_ = "A"
    itemLocationCode_ = "A"
    Call UpdateSNS
    Call UpdateDMC
    IsValidDMC_ = True
    
    'デフォルト値を使用
    Dim mIssueNo As New clsS1000DIssue
    WorkAndIssueNumber_ = mIssueNo.WorkAndIssueNumber
    Dim mLangugeAndCountry As New clsS1000DLanguage
    LanguageAndCountry_ = mLangugeAndCountry.LangueAndCountry
    
'    Set mURN = Nothing
    Set mIssueNo = Nothing
    Set mLangugeAndCountry = Nothing
End Sub

Private Sub UpdateSNS()
SNS_ = SystemCode_ & cDMCSeparator & _
        subSystemCode_ & subSubSystemCode_ & cDMCSeparator & _
        assyCode_
End Sub

Private Sub UpdateDMC()
DMC_ = ModelIdentCode_ & cDMCSeparator & _
        systemDiffCode_ & cDMCSeparator & _
        SystemCode_ & cDMCSeparator & _
        subSystemCode_ & subSubSystemCode_ & cDMCSeparator & _
        assyCode_ & cDMCSeparator & _
        disassyCode_ & disassyCodeVariant_ & cDMCSeparator & _
        infoCode_ & infoCodeVariant_ & cDMCSeparator & _
        itemLocationCode_
End Sub
