Attribute VB_Name = "mdlTIRZones"
'****************************************************************************************
'
'    Zones用Integration Fileを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "Saab送付用Zones用Integration Fileを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub TIRZones用IntegrationFileを作成(ByRef myButton As IRibbonControl)
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim TIRZonesExcelSheet As New clsMetadataSheet
    If Not TIRZonesExcelSheet.IsValidMetadateSheet(Zones) Then
        MsgBox "有効なTIR-Zones用データが指定されていません." & vbCrLf & "TIR-Integration Fileの作成を中止します.", vbCritical + vbOKOnly, cAppName
        Exit Sub
    End If
    Dim lLastRow As Long
    lLastRow = TIRZonesExcelSheet.LastRow
    Dim lCurrentRow As Long, lCurrentCol As Long
    
    Dim TIRZonesSetting As clsConfigTIRZones
    Set TIRZonesSetting = New clsConfigTIRZones
    
    Dim TIRZonesIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename As String, sMetadataFullPath As String, sBackupFilePath As String
    
    TIRZonesIntegrationFile.FileCategory = ESRDFileCategory.Zones
    sMetadataFilename = TIRZonesIntegrationFile.filename
    
    '指定フォルダ内にIntegrarion Fileを作成する
    sMetadataFullPath = StrAddPathSeparator(TIRZonesSetting.IntegrationFileFolder) & sMetadataFilename
    sBackupFilePath = StrAddPathSeparator(TIRZonesSetting.IntegrationFileBackupFolder) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim TIRZones As New clsTIRZones
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long, TIRCount As Long
    
    trgTs.WriteLine TIRZones.GetTitleRow
    TIRCount = 0
    
    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + TIRZones.MetadataItems.Count - 1
            TIRZones.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print TIRZones.GetMetadataRow
#End If
        TIRCount = TIRCount + 1
        
        TIRZonesExcelSheet.PutMetadataFilename ESRDFileCategory.Zones, sMetadataFilename, lCurrentRow
        
        trgTs.WriteLine TIRZones.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOFのテキストを追加する
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    With TIRZonesSetting
        .IntegrationFilePath = sMetadataFullPath
        .IntegrationFileName = sMetadataFilename
        .IntegrationFileDate = TIRZonesIntegrationFile.FileDate
        .ItemCount = TIRCount
        .Save
    End With
    
    If TIRZonesSetting.IntegrationFileBackupFolderExists Then
        fso.CopyFile sMetadataFullPath, sBackupFilePath, True
    End If
    
    Set TIRZonesSetting = Nothing
    
    MsgBox "Saab送付用TIRZonesを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRZonesExcelSheet = Nothing
    Set TIRZonesIntegrationFile = Nothing
    Set TIRZones = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TIRZonesExcelSheet = Nothing
    Set TIRZonesIntegrationFile = Nothing
    Set TIRZones = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub





