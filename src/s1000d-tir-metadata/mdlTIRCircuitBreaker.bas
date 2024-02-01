Attribute VB_Name = "mdlTIRCircuitBreaker"
'****************************************************************************************
'
'    CircuitBreakerTIR用Integration Fileを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab送付用CircuitBreakerTIR用Integration Fileを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub TIRCircuitBreakers用IntegrationFileを作成(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRCircuitBreakersExcelSheet As New clsMetadataSheet
    If Not TIRCircuitBreakersExcelSheet.IsValidMetadateSheet(CircuitBreakers) Then
        MsgBox "有効なTIR-CircuitBreakers用データが指定されていません." & vbCrLf & "TIR-Integration Fileの作成を中止します.", vbCritical + vbOKOnly, cAppName
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
    
    '指定フォルダ内にIntegrarion Fileを作成する
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

    'EOFのテキストを追加する
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
    
    MsgBox "Saab送付用TIRCircuitBreakersを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
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



