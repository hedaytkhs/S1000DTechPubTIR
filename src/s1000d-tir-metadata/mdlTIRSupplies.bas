Attribute VB_Name = "mdlTIRSupplies"
'****************************************************************************************
'
'    SuppliesTIR���쐬���邽�߂̃��W���[��
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab���t�pSuppliesTIR���쐬"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub SuppliesTIR���쐬(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRSuppliesExcelSheet As New clsMetadataSheet
    If Not TIRSuppliesExcelSheet.IsValidMetadateSheet(SUPPLIES) Then
        MsgBox "�L����TIR-Supplies�p�f�[�^���w�肳��Ă��܂���." & vbCrLf & "TIR-Integration File�̍쐬�𒆎~���܂�.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    Dim lLastRow As Long
    lLastRow = TIRSuppliesExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIRSuppliesSetting As clsConfigTIRSupplies
    Set TIRSuppliesSetting = New clsConfigTIRSupplies
    
    Dim TIRSuppliesIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    TIRSuppliesIntegrationFile.FileCategory = ESRDFileCategory.SUPPLIES
    sMetadataFilename = TIRSuppliesIntegrationFile.filename
    
    '�w��t�H���_����Integrarion File���쐬����
    sMetadataFullPath = StrAddPathSeparator(TIRSuppliesSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIRSuppliesSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim SuppliesTIR As New clsSuppliesTIR
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    Dim IntegrationFileRowText As tpIntegrationRowTextWithResult
    
    trgTs.WriteLine SuppliesTIR.GetTitleRow
    TIRCount = 0

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + SuppliesTIR.MetadataItems.Count - 1
            SuppliesTIR.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print SuppliesTIR.GetMetadataRow.Text
#End If
        TIRCount = TIRCount + 1
        
        TIRSuppliesExcelSheet.PutMetadataFilename ESRDFileCategory.SUPPLIES, sMetadataFilename, lCurrentRow
        
        IntegrationFileRowText = SuppliesTIR.GetMetadataRow
        trgTs.WriteLine IntegrationFileRowText.Text
                
        ShowMissingVendorCode IntegrationFileRowText
                
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIRSuppliesSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIRSuppliesIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With
    
    If TIRSuppliesSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIRSuppliesSetting = Nothing
    
    MsgBox "Saab���t�pSuppliesTIR��ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRSuppliesExcelSheet = Nothing
    Set TIRSuppliesIntegrationFile = Nothing
    Set SuppliesTIR = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRSuppliesExcelSheet = Nothing
    Set TIRSuppliesIntegrationFile = Nothing
    Set SuppliesTIR = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub



