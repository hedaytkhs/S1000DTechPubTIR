Attribute VB_Name = "mdlEngSrcMetadata"
'****************************************************************************************
'
'    EngineeringSource���^�f�[�^���쐬���邽�߂̃��W���[��
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab���t�pEngineeringSource���^�f�[�^���쐬"


Public Enum EngSrcCategoryAllowableValue
    DWG = 1
    IPC = 2
    Specifacation = 3
    EngineeringDocument = 4
    TIR = 5
    Wiring = 6
    ConfigurationData = 7
    ChangeRequestChangeNotice = 8
    VendorInfo = 9
    DraftDM = 9
    ToolInformation = 10
    NDTData = 11
    SRMData = 12
    TechnicalDraft = 13
    OtherFiles = 14
End Enum

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub EngineeringSource���^�f�[�^���쐬(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub

    Dim AuthorMetadataSheet As New clsMetadataSheet
    If Not AuthorMetadataSheet.IsValidMetadateSheet(Author) Then
        MsgBox "�L����Engineering Source Metadata�p�f�[�^���w�肳��Ă��܂���." & vbCrLf & "Engineering Source Metadata�̍쐬�𒆎~���܂�.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    
    Dim lLastRow As Long
    lLastRow = AuthorMetadataSheet.LastRow
    Dim lCurrentRow, lCurrentCol As Long
    
    Dim AuthorMetadatafile As New clsMetadataFile
    Dim sMetadataFilename, sMetadataFullPath As String
    AuthorMetadatafile.FileCategory = ESRDFileCategory.Author
    sMetadataFilename = AuthorMetadatafile.filename
    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim AuthorMetadata As New clsEngSrcMetadata
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long
    
    trgTs.WriteLine AuthorMetadata.GetTitleRow

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + AuthorMetadata.MetadataItems.Count - 1
            AuthorMetadata.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print AuthorMetadata.GetMetadataRow
#End If
        
        trgTs.WriteLine AuthorMetadata.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    MsgBox "Saab���t�pEngineeringSource���^�f�[�^��ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set AuthorMetadataSheet = Nothing
    Set AuthorMetadatafile = Nothing
    Set AuthorMetadata = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set AuthorMetadataSheet = Nothing
    Set AuthorMetadatafile = Nothing
    Set AuthorMetadata = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

#If (DEBUG_MODE = 0) Then

Sub WiringMetadatafileTest()
    On Error GoTo errHandler
    Dim WiringMetadatafile As New clsMetadataFile
    WiringMetadatafile.FileCategory = WireList
    With WiringMetadatafile
        Debug.Print .BaseName
        Debug.Print .FileCategoryName
        Debug.Print .filename
        .DMC = "MRJ-A-20-70-03-00A-254A-A"
        Debug.Print .filename
    End With
    
    Set WiringMetadatafile = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

#End If
