Attribute VB_Name = "mdlTIRSupplies"
'****************************************************************************************
'
'    SuppliesTIRを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab送付用SuppliesTIRを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub SuppliesTIRを作成(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRSuppliesExcelSheet As New clsMetadataSheet
    If Not TIRSuppliesExcelSheet.IsValidMetadateSheet(SUPPLIES) Then
        MsgBox "有効なTIR-Supplies用データが指定されていません." & vbCrLf & "TIR-Integration Fileの作成を中止します.", vbCritical + vbOKOnly, cAppName
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
    
    '指定フォルダ内にIntegrarion Fileを作成する
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

    'EOFのテキストを追加する
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
    
    MsgBox "Saab送付用SuppliesTIRを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
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



