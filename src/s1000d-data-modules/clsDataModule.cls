VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    DMの各種ファイル名を操作するためのクラス
'
'    TT用の各種ファイル名及びパスを設定、生成、取得することができる
'
' 注意事項：TT検証フォルダのパスを設定しておかないと各種パスの値が不正な値となってしまう。
' TT検証フォルダのパス
' Dim TTBaseFolderPath_ As String
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

'DMC全体
Dim DMC_ As String

' To Do
'PWのSDCの反映要

'Dim inWork_ As String * 2
'Dim issueNumber_ As String * 3
'Dim issueType_ As S1000D_IssueType
Dim WorkAndIssueNumber_ As String * 6
Dim DMLanguage_ As String
Dim URN_ As String * 3
Dim filename_ As String * 64
Dim WorkNumber_ As String * 2
Dim inWork_ As String * 2
Dim issueNumber_ As String * 3

' TT検証フォルダのパス
Dim TTBaseFolderPath_ As String
Dim DMCPath_ As String
Dim DMCIssuePath_ As String


'=========================================================
' TT作業フォルダ関連情報
'=========================================================
Property Let TTBaseFolderPath(ByVal sTTBaseFolder As String)
    TTBaseFolderPath_ = sTTBaseFolder
    '値チェック
    '関連情報を更新
    Call UpdateDMCPath
    Call UpdateDMCIssuePath
End Property

Property Get TTBaseFolderPath() As String
    TTBaseFolderPath = TTBaseFolderPath_
End Property

Property Get DMCPath() As String
    If Left(DMC_, 4) = "DMC-" Then
        DMC_ = Mid(DMC_, 5, Len(DMC_) - 4)
    End If
    DMCPath_ = StrAddPathSeparator(TTBaseFolderPath_) & DMC_
    DMCPath = DMCPath_
End Property


Property Get DF4TTTemplateFilePath() As String
    DF4TTTemplateFilePath = StrAddPathSeparator(TTBaseFolderPath_) & "DF4TT_template\DMC-MRJ-A-xx-xx-xx-xxx-xxxx-A_yyy-zz_TT-01_template_v1.00.xlsx"
'    DF4TTTemplateFilePath = DMCPath_
End Property


Property Get DMCIssuePath() As String
    DMCIssuePath_ = StrAddPathSeparator(DMCPath_) & WorkAndIssueNumber_
    DMCIssuePath = DMCIssuePath_
End Property


Property Get DMCIssueInitialDMPath() As String
    DMCIssueInitialDMPath = StrAddPathSeparator(DMCIssuePath_) & "Initial_DM"
End Property

Property Get DMCIssueEngDataPath() As String
    DMCIssueEngDataPath = StrAddPathSeparator(DMCIssuePath_) & "01_Engineering_Data"
End Property

Property Get DMCIssueSTEPath() As String
    DMCIssueSTEPath = StrAddPathSeparator(DMCIssuePath_) & "02_STE"
End Property

Property Get DMCIssueXMLPath() As String
    DMCIssueXMLPath = StrAddPathSeparator(DMCIssuePath_) & "03_XML"
End Property

Property Get DMCIssueIllustPath() As String
    DMCIssueIllustPath = StrAddPathSeparator(DMCIssuePath_) & "04_Illustration"
End Property

Property Get DMFileName()
    DMFileName = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_" & UCase(DMLanguage_) & ".PDF"
End Property

'=========================================================
' TT作業フォルダ関連情報
'=========================================================
Property Get getTTCommentPDFFileName()
    getTTCommentPDFFileName = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_" & UCase(DMLanguage_) & "_TT-" & WorkNumber_ & ".PDF"
End Property

Property Get getDF4TTFileName()
    getDF4TTFileName = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_TT-" & WorkNumber_ & ".xlsx"
End Property


'=========================================================
' DMCの要素分解及びファイル名取得用情報
'=========================================================

Property Let DMC(ByVal sDMC As String)
    DMC_ = sDMC
    Call getDMFileName
End Property

Property Get DMC() As String
    DMC = DMC_
End Property

Property Let WorkNumber(ByVal sWorkNumber As String)
    If IsDigit(sWorkNumber) Then
        Dim lWorkNumber, lWrkNo As Long
        lWorkNumber = CLng(sWorkNumber)
        lWrkNo = (lWorkNumber Mod 100)
        WorkNumber_ = FORMAT(lWrkNo, "00")
    Else
        'デフォルト値をセットする
        WorkNumber_ = "01"
    End If
End Property

Property Get WorkNumber() As String
    WorkNumber = WorkNumber_
End Property

Property Let WorkAndIssueNumber(ByVal sWorkAndIssueNumber As String)
    WorkAndIssueNumber_ = sWorkAndIssueNumber
    Call getDMFileName
End Property

Property Get WorkAndIssueNumber() As String
    WorkAndIssueNumber = WorkAndIssueNumber_
End Property

Property Get inWork() As String
    inWork = inWork_
End Property

Property Get issueNumber() As String
    issueNumber = issueNumber_
End Property

Property Get filename() As String
    filename = filename_
End Property

Private Sub getDMFileName()
    filename_ = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_" & UCase(DMLanguage_) & ".PDF"
End Sub

Private Sub UpdateDMCPath()
    DMCPath_ = StrAddPathSeparator(TTBaseFolderPath_) & DMC_
End Sub
Private Sub UpdateDMCIssuePath()
    DMCIssuePath_ = StrAddPathSeparator(DMCPath_) & WorkAndIssueNumber_
End Sub

Private Sub Class_Initialize()
    
    'URNの定義から"DMC"を取得
    Dim mURN As New clsS1000DURN
    mURN.URNCategory = S1000D_URN.DMC
    URN_ = mURN.URN
    
    Dim mDMC As New clsDMC
    DMC_ = mDMC.DMC
    
    'デフォルト値を使用
    Dim mIssueNo As New clsS1000DIssue
    WorkAndIssueNumber_ = mIssueNo.WorkAndIssueNumber
    Dim mLangugeAndCountry As New clsS1000DLanguage
    DMLanguage_ = mLangugeAndCountry.LangueAndCountry
    
    DMCPath_ = ""
    DMCIssuePath_ = ""
    filename_ = ""
    
    Call getDMFileName
    
    Set mDMC = Nothing
    Set mIssueNo = Nothing
    Set mLangugeAndCountry = Nothing
End Sub



