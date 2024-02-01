Attribute VB_Name = "mdlTIRTools"
'****************************************************************************************
'
'    ToolsTIRを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab送付用ToolsTIRを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub ToolsTIRを作成(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRToolsExcelSheet As New clsMetadataSheet
    If Not TIRToolsExcelSheet.IsValidMetadateSheet(Tools) Then
        MsgBox "有効なTIR-Tools用データが指定されていません." & vbCrLf & "TIR-Integration Fileの作成を中止します.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    
    Dim lLastRow As Long
    lLastRow = TIRToolsExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIRToolsSetting As clsConfigTIRTools
    Set TIRToolsSetting = New clsConfigTIRTools
    
    Dim TIRToolsIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    TIRToolsIntegrationFile.FileCategory = ESRDFileCategory.Tools
    sMetadataFilename = TIRToolsIntegrationFile.filename

    '指定フォルダ内にIntegrarion Fileを作成する
    sMetadataFullPath = StrAddPathSeparator(TIRToolsSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIRToolsSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim ToolsTIR As New clsToolsTIR
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    Dim IntegrationFileRowText As tpIntegrationRowTextWithResult
    
    trgTs.WriteLine ToolsTIR.GetTitleRow
    TIRCount = 0
    
    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + ToolsTIR.MetadataItems.Count - 1
            ToolsTIR.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print ToolsTIR.GetMetadataRow.Text
#End If
        TIRCount = TIRCount + 1
        
        TIRToolsExcelSheet.PutMetadataFilename ESRDFileCategory.Tools, sMetadataFilename, lCurrentRow
        
        IntegrationFileRowText = ToolsTIR.GetMetadataRow
        trgTs.WriteLine IntegrationFileRowText.Text
        
        ShowMissingVendorCode IntegrationFileRowText
        
        DoEvents
    Next lCurrentRow

    'EOFのテキストを追加する
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIRToolsSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIRToolsIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With
    
    If TIRToolsSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIRToolsSetting = Nothing
    
    MsgBox "Saab送付用ToolsTIRを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRToolsExcelSheet = Nothing
    Set TIRToolsIntegrationFile = Nothing
    Set ToolsTIR = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRToolsExcelSheet = Nothing
    Set TIRToolsIntegrationFile = Nothing
    Set ToolsTIR = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub



