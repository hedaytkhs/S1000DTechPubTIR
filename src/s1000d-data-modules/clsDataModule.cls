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
'    DM�̊e��t�@�C�����𑀍삷�邽�߂̃N���X
'
'    TT�p�̊e��t�@�C�����y�уp�X��ݒ�A�����A�擾���邱�Ƃ��ł���
'
' ���ӎ����FTT���؃t�H���_�̃p�X��ݒ肵�Ă����Ȃ��Ɗe��p�X�̒l���s���Ȓl�ƂȂ��Ă��܂��B
' TT���؃t�H���_�̃p�X
' Dim TTBaseFolderPath_ As String
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

'DMC�S��
Dim DMC_ As String

' To Do
'PW��SDC�̔��f�v

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

' TT���؃t�H���_�̃p�X
Dim TTBaseFolderPath_ As String
Dim DMCPath_ As String
Dim DMCIssuePath_ As String


'=========================================================
' TT��ƃt�H���_�֘A���
'=========================================================
Property Let TTBaseFolderPath(ByVal sTTBaseFolder As String)
    TTBaseFolderPath_ = sTTBaseFolder
    '�l�`�F�b�N
    '�֘A�����X�V
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
' TT��ƃt�H���_�֘A���
'=========================================================
Property Get getTTCommentPDFFileName()
    getTTCommentPDFFileName = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_" & UCase(DMLanguage_) & "_TT-" & WorkNumber_ & ".PDF"
End Property

Property Get getDF4TTFileName()
    getDF4TTFileName = URN_ & "-" & DMC_ & "_" & WorkAndIssueNumber_ & "_TT-" & WorkNumber_ & ".xlsx"
End Property


'=========================================================
' DMC�̗v�f�����y�уt�@�C�����擾�p���
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
        '�f�t�H���g�l���Z�b�g����
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
    
    'URN�̒�`����"DMC"���擾
    Dim mURN As New clsS1000DURN
    mURN.URNCategory = S1000D_URN.DMC
    URN_ = mURN.URN
    
    Dim mDMC As New clsDMC
    DMC_ = mDMC.DMC
    
    '�f�t�H���g�l���g�p
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


