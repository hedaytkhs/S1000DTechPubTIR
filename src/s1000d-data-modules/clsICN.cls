VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsICN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'
'    ICNを操作するためのクラス
'
'
'    ICNの要素分解、結合、値チェック
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

'ICN全体
Dim ICN_ As String

'ICNの各要素
Dim URN_ As String * 3
Dim ModelIdentCode_ As String
Dim systemDiffCode_ As String * 1

Dim SNS_ As String * 6

' SNS分解
Dim SystemCode_ As String * 2
Dim subSystemCode_ As String * 1
Dim subSubSystemCode_ As String * 1
Dim assyCode_ As String * 2
Dim disassyCode_ As String * 2
Dim disassyCodeVariant_ As String * 1
Dim infoCode_ As String * 3
Dim infoCodeVariant_ As String * 1
Dim RPC_ As String * 1
Dim OriginatorCode As String * 5
Dim itemLocationCode_ As String * 1
Dim IsValidICN_ As Boolean

Const cICNSeparator = "-"

Property Let ICN(ByVal sICN As String)
    ICN_ = sICN
    Call DivedeICN
End Property

Property Let ModelIdentCode(ByVal sModelIdentCode As String)
    Select Case sModelIdentCode
        Case "MRJ", "MWG", "MBR", "S1000D", "PW1000G"
            IsValidICN_ = True
'        Exit Function
        Case Else
            IsValidICN_ = False
    End Select
    
    ModelIdentCode_ = sModelIdentCode
    Call UpdateICN
End Property

Property Let systemDiffCode(ByVal ssystemDiffCode As String)
    If Not ssystemDiffCode Like "[a-zA-Z]" Then IsValidICN_ = False
    systemDiffCode_ = Left(ssystemDiffCode, 1)
    Call UpdateICN
End Property

Property Let SystemCode(ByVal sSystemCode As String)
    If Not IsNumeric(sSystemCode) Then
        IsValidICN_ = False
    End If
    SystemCode_ = Left(sSystemCode, 2)
    Call UpdateSNS
    Call UpdateICN
End Property
    
Property Let subSystemCode(ByVal ssubSystemCode As String)
    If Not IsNumeric(ssubSystemCode) Then
        IsValidICN_ = False
    End If
    subSystemCode_ = Left(ssubSystemCode, 1)
    Call UpdateSNS
    Call UpdateICN
End Property
    
Property Let subSubSystemCode(ByVal ssubSubSystemCode As String)
    If Not IsNumeric(ssubSubSystemCode) Then
        IsValidICN_ = False
    End If
    subSubSystemCode_ = Left(ssubSubSystemCode, 1)
    Call UpdateSNS
    Call UpdateICN
End Property
    
Property Let assyCode(ByVal sassyCode As String)
    If Not IsNumeric(sassyCode) Then
        IsValidICN_ = False
    End If
    assyCode_ = Left(sassyCode, 2)
    Call UpdateSNS
    Call UpdateICN
End Property
    
Property Let disassyCode(ByVal sdisassyCode As String)
    If Not IsNumeric(sdisassyCode) Then
        IsValidICN_ = False
    End If
    disassyCode_ = Left(sdisassyCode, 2)
    Call UpdateICN
End Property
    
Property Let disassyCodeVariant(ByVal sdisassyCodeVariant As String)
    If Not sdisassyCodeVariant Like "[a-zA-Z]" Then IsValidICN_ = False
    disassyCodeVariant_ = Left(sdisassyCodeVariant, 1)
    Call UpdateICN
End Property

Property Let infoCode(ByVal sinfoCode As String)
    If Not IsNumeric(sinfoCode) Then
        IsValidICN_ = False
    End If
    infoCode_ = Left(sinfoCode, 3)
    Call UpdateICN
End Property

Property Let infoCodeVariant(ByVal sinfoCodeVariant As String)
    If Not sinfoCodeVariant Like "[a-zA-Z]" Then IsValidICN_ = False
    infoCodeVariant_ = Left(sinfoCodeVariant, 1)
    Call UpdateICN
End Property

Property Let itemLocationCode(ByVal sitemLocationCode As String)
    If Not sitemLocationCode Like "[a-zA-Z]" Then IsValidICN_ = False
    itemLocationCode_ = Left(sitemLocationCode, 1)
    Call UpdateICN
End Property

Property Get ICN() As String
    ICN = ICN_
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



'Property Get disassyCode() As String
'    disassyCode = disassyCode_
'End Property
'
'Property Get disassyCodeVariant() As String
'    disassyCodeVariant = disassyCodeVariant_
'End Property
'
'Property Get infoCode() As String
'    infoCode = infoCode_
'End Property
'
'Property Get infoCodeVariant() As String
'    infoCodeVariant = infoCodeVariant_
'End Property
'
'Property Get itemLocationCode() As String
'    itemLocationCode = itemLocationCode_
'End Property

Property Get IsValidICN() As String
    IsValidICN = IsValidICN_
End Property

Private Sub DivedeICN()
    Dim tmp As Variant
    
    tmp = Split(ICN_, "-")
    If UBound(tmp) < 9 Then
        IsValidICN_ = False
        Exit Sub
    End If
    
    ModelIdentCode_ = "" & tmp(1)
    systemDiffCode_ = Left("" & tmp(2), 1)
    SNS_ = Left("" & tmp(3), 6)
    
    If Len(SNS_) = 6 Then
        SystemCode_ = Left(SNS_, 2)
        subSystemCode_ = Mid("" & tmp(3), 1)
        subSubSystemCode_ = Mid("" & tmp(3), 2, 1)
        assyCode_ = Mid("" & tmp(4), 2)
    Else
        IsValidICN_ = False
    End If
    Call UpdateSNS
    disassyCode_ = Left("" & tmp(5), 2)
    disassyCodeVariant_ = Mid("" & tmp(5), 3, 1)
    infoCode_ = Left("" & tmp(6), 3)
    infoCodeVariant_ = Mid("" & tmp(6), 4, 1)
    itemLocationCode_ = Left("" & tmp(7), 1)
End Sub

Private Sub Class_Initialize()
    
    'URNの定義から"ICN"を取得
    Dim mURN As New clsS1000DURN
    mURN.URNCategory = S1000D_URN.ICN
    URN_ = mURN.URN
    
    ModelIdentCode_ = "MRJ"
    systemDiffCode_ = "A"
    SystemCode_ = "00"
    subSystemCode_ = "0"
    subSubSystemCode_ = "0"
    assyCode_ = "00"
    Call UpdateSNS
    Call UpdateICN
    IsValidICN_ = True
End Sub

Private Sub UpdateSNS()
    SNS_ = SystemCode_ & subSystemCode_ & subSubSystemCode_ & assyCode_
End Sub

Private Sub UpdateICN()
ICN_ = URN_ & cICNSeparator & ModelIdentCode_ & cICNSeparator & _
        systemDiffCode_ & cICNSeparator & _
        SystemCode_ & cICNSeparator & _
        subSystemCode_ & subSubSystemCode_ & cICNSeparator & _
        assyCode_ & cICNSeparator & _
        disassyCode_ & disassyCodeVariant_ & cICNSeparator & _
        infoCode_ & infoCodeVariant_ & cICNSeparator & _
        itemLocationCode_
End Sub
