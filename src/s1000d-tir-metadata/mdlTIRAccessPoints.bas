Attribute VB_Name = "mdlTIRAccessPoints"
'****************************************************************************************
'
'    AccessPoints用Integration Fileを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab送付用AccessPoints用Integration Fileを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub TIRAccessPoints用IntegrationFileを作成(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub

    Dim TIRAccessPointsExcelSheet As New clsMetadataSheet
    If Not TIRAccessPointsExcelSheet.IsValidMetadateSheet(AccessPoints) Then
        MsgBox "有効なTIR-AccessPoints用データが指定されていません." & vbCrLf & "TIR-Integration Fileの作成を中止します.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    Dim lLastRow As Long
    lLastRow = TIRAccessPointsExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
    Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
    
    Dim TIRAccessPointsIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    
    TIRAccessPointsIntegrationFile.FileCategory = ESRDFileCategory.AccessPoints
    sMetadataFilename = TIRAccessPointsIntegrationFile.filename
    
    '指定フォルダ内にIntegrarion Fileを作成する
    sMetadataFullPath = StrAddPathSeparator(TIRAccessPointsSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIRAccessPointsSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim TIRAccessPoints As New clsTIRAccessPoints
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    
    trgTs.WriteLine TIRAccessPoints.GetTitleRow
    TIRCount = 0
    
    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + TIRAccessPoints.MetadataItems.Count - 1
            TIRAccessPoints.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print TIRAccessPoints.GetMetadataRow
#End If
        TIRCount = TIRCount + 1
        
        TIRAccessPointsExcelSheet.PutMetadataFilename ESRDFileCategory.AccessPoints, sMetadataFilename, lCurrentRow
        
        trgTs.WriteLine TIRAccessPoints.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOFのテキストを追加する
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIRAccessPointsSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIRAccessPointsIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With
    
    If TIRAccessPointsSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIRAccessPointsSetting = Nothing
    
    MsgBox "Saab送付用TIRAccessPointsを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRAccessPointsExcelSheet = Nothing
    Set TIRAccessPointsIntegrationFile = Nothing
    Set TIRAccessPoints = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRAccessPointsExcelSheet = Nothing
    Set TIRAccessPointsIntegrationFile = Nothing
    Set TIRAccessPoints = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub




