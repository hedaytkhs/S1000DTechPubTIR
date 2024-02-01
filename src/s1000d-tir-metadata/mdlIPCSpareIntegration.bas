Attribute VB_Name = "mdlIPCSpareIntegration"
'****************************************************************************************
'
'    IPCSpareIntegrationFile���쐬���邽�߂̃��W���[��
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Private Const cAppName = "IPC-Spare Integration File���쐬"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub IPCSpareIntegrationFile���쐬()
    On Error GoTo errHandler
    Dim ATAPartNumberExcelSheet As New clsMetadataSheet
    Dim lLastRow As Long
    lLastRow = ATAPartNumberExcelSheet.LastRow
    Dim lCurrentRow, lCurrentCol As Long
    
    Dim IPCSpareIntegrationFile As New clsMetadataFile
    Dim sMetadataFilename, sMetadataFullPath As String
    IPCSpareIntegrationFile.FileCategory = ESRDFileCategory.IPCSpare
    sMetadataFilename = IPCSpareIntegrationFile.filename
    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    Dim IPCSpareIntegration As New clsIPCSpareIntegration
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long
    
    trgTs.WriteLine IPCSpareIntegration.GetTitleRow

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + IPCSpareIntegration.MetadataItems.Count - 1
            IPCSpareIntegration.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print IPCSpareIntegration.GetMetadataRow
#End If
        
        trgTs.WriteLine IPCSpareIntegration.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOF�̃e�L�X�g��ǉ�����
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    MsgBox "Saab���t�pIPC-Spare Integration File��ۑ����܂���!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set ATAPartNumberExcelSheet = Nothing
    Set IPCSpareIntegrationFile = Nothing
    Set IPCSpareIntegration = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set ATAPartNumberExcelSheet = Nothing
    Set IPCSpareIntegrationFile = Nothing
    Set IPCSpareIntegration = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

