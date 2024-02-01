Attribute VB_Name = "mdlTIRCircuitBreaker"
'****************************************************************************************
'
'    CircuitBreakerTIR�pIntegration File���쐬���邽�߂̃��W���[��
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab���t�pCircuitBreakerTIR�pIntegration File���쐬"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub TIRCircuitBreakers�pIntegrationFile���쐬(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRCircuitBreakersExcelSheet As New clsMetadataSheet
    If Not TIRCircuitBreakersExcelSheet.IsValidMetadateSheet(CircuitBreakers) Then
        MsgBox "�L����TIR-CircuitBreakers�p�f�[�^���w�肳��Ă��܂���." & vbCrLf & "TIR-Integration File�̍쐬�𒆎~���܂�.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    Dim lLastRow As Long
    lLastRow = TIRCircuitBreakersExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
    Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
    
    Dim TIRCircuitBreakersIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    
    TIRCircuitBreakersIntegrationFile.FileCategory = ESRDFileCategory.CircuitBreakers
    sMetadataFilename = TIRCircuitBreakersIntegrationFile.filename
    
    '�w��t�H���_����Integrarion File���쐬����
    sMetadataFullPath = StrAddPathSeparator(TIRCircuitBreakersSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIRCircuitBreakersSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim TIRCircuitBreakers As New clsTIRCircuitBreaker
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    
    trgTs.WriteLine TIRCircuitBreakers.GetTitleRow
    TIRCount = 0
    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + TIRCircuitBreakers.MetadataItems.Count - 1
            TIRCircuitBreakers.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print TIRCircuitBreakers.GetMetadataRow
#End If
        TIRCount = TIRCount + 1
        
        TIRCircuitBreakersExcelSheet.PutMetadataFilename ESRDFileCategory.CircuitBreakers, sMetadataFilename, lCurrentRow
        
        trgTs.WriteLine TIRCircuitBreakers.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIRCircuitBreakersSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIRCircuitBreakersIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With

    If TIRCircuitBreakersSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIRCircuitBreakersSetting = Nothing
    
    MsgBox "Saab���t�pTIRCircuitBreakers��ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRCircuitBreakersExcelSheet = Nothing
    Set TIRCircuitBreakersIntegrationFile = Nothing
    Set TIRCircuitBreakers = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRCircuitBreakersExcelSheet = Nothing
    Set TIRCircuitBreakersIntegrationFile = Nothing
    Set TIRCircuitBreakers = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub



