Attribute VB_Name = "mdlTIREnterprise"
'****************************************************************************************
'
'    EnterpriseTIR���쐬���邽�߂̃��W���[��
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab���t�pEnterpriseTIR���쐬"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub EnterpriseTIR���쐬(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIREnterpriseExcelSheet As New clsMetadataSheet
    If Not TIREnterpriseExcelSheet.IsValidMetadateSheet(Enterprise) Then
        MsgBox "�L����TIR-Enterprise�p�f�[�^���w�肳��Ă��܂���." & vbCrLf & "TIR-Integration File�̍쐬�𒆎~���܂�.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    Dim lLastRow As Long
    lLastRow = TIREnterpriseExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIREnterpriseSetting As clsConfigTIREnterprise
    Set TIREnterpriseSetting = New clsConfigTIREnterprise
    
    Dim TIREnterpriseIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    TIREnterpriseIntegrationFile.FileCategory = ESRDFileCategory.Enterprise
    sMetadataFilename = TIREnterpriseIntegrationFile.filename
    
    '�w��t�H���_����Integrarion File���쐬����
    sMetadataFullPath = StrAddPathSeparator(TIREnterpriseSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIREnterpriseSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
#If False Then
    'UNICODE encoded
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting, True)
#Else

    'ANSI encoded
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting, False)
#End If
    
    Dim EnterpriseTIR As New clsEnterpriseTIR
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    
    trgTs.WriteLine EnterpriseTIR.GetTitleRow
    TIRCount = 0

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + EnterpriseTIR.MetadataItems.Count - 1
            EnterpriseTIR.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print EnterpriseTIR.GetMetadataRow
#End If
        TIRCount = TIRCount + 1
        
        TIREnterpriseExcelSheet.PutMetadataFilename ESRDFileCategory.Enterprise, sMetadataFilename, lCurrentRow
        
        trgTs.WriteLine EnterpriseTIR.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIREnterpriseSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIREnterpriseIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With
    
    If TIREnterpriseSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIREnterpriseSetting = Nothing
    
    MsgBox "Saab���t�pCSV�t�@�C����ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIREnterpriseExcelSheet = Nothing
    Set TIREnterpriseIntegrationFile = Nothing
    Set EnterpriseTIR = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIREnterpriseExcelSheet = Nothing
    Set TIREnterpriseIntegrationFile = Nothing
    Set EnterpriseTIR = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

Sub �`�F�b�N�p�SEnterpriseTIR���쐬()
    On Error GoTo errHandler
    Const cColValidRecord = 21
    Dim TIREnterpriseExcelSheet As New clsMetadataSheet
    Dim lLastRow As Long
    lLastRow = TIREnterpriseExcelSheet.LastRow
    Dim lCurrentRow, lCurrentCol As Long
    
    Dim TIREnterpriseIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename, sMetadataFullPath As String
    TIREnterpriseIntegrationFile.FileCategory = Enterprise
    sMetadataFilename = TIREnterpriseIntegrationFile.filename
    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim EnterpriseTIR As New clsEnterpriseTIR
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long
    
    trgTs.WriteLine EnterpriseTIR.GetTitleRow

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit + 1 To cMetadataColInit + EnterpriseTIR.MetadataItems.Count
            EnterpriseTIR.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print EnterpriseTIR.GetMetadataRow
#End If
        
            If Cells(lCurrentRow, cColValidRecord).Value = "y" Then
                trgTs.WriteLine EnterpriseTIR.GetMetadataRow
            End If
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    MsgBox "Saab���t�pEnterpriseTIR��ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIREnterpriseExcelSheet = Nothing
    Set TIREnterpriseIntegrationFile = Nothing
    Set EnterpriseTIR = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIREnterpriseExcelSheet = Nothing
    Set TIREnterpriseIntegrationFile = Nothing
    Set EnterpriseTIR = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

