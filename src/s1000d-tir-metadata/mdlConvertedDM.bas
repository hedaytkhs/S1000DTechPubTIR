Attribute VB_Name = "mdlConvertedDM"
'****************************************************************************************
'
'    ConvertedDM用メタデータを作成するためのモジュール
'
'    MRJ Technical Publication Tool
'
'
' Modified by Keiji Motomura 2015/4/2
'
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit


Private Const cAppName = "Saab送付用 Converted DM用メタデータを作成"

Private Const cMetadataColInit = 1
Private Const cMetadataRowInit = 2

Sub ConvertedDM用メタデータを作成()
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim ConvertedDMMetadataSheet As New clsMetadataSheet
    Dim lLastRow As Long
    lLastRow = ConvertedDMMetadataSheet.LastRow
    Dim lCurrentRow, lCurrentCol As Long
    
    Dim ConvertedDMMetadatafile As New clsMetadataFile
    Dim sMetadataFilename, sMetadataFullPath As String
    ConvertedDMMetadatafile.FileCategory = ESRDFileCategory.ConvertedDM
    sMetadataFilename = ConvertedDMMetadatafile.filename
    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
'   20150401本村改修（単なるミスと思われる）
'    Dim ConvertedDMMetadata As New clsEngSrcMetadata
    Dim ConvertedDMMetadata As New clsConvertedDM
    
    Dim sFileCategoryTitle As String
    Dim sFileCategory As String
    Dim lItemCnt As Long
    
    trgTs.WriteLine ConvertedDMMetadata.GetTitleRow

    For lCurrentRow = cMetadataRowInit To lLastRow
        lItemCnt = 1
        For lCurrentCol = cMetadataColInit To cMetadataColInit + ConvertedDMMetadata.MetadataItems.Count - 1
            ConvertedDMMetadata.MetadataItems(lItemCnt).Value = GetValidCharForESRD(Cells(lCurrentRow, lCurrentCol).Value)
            lItemCnt = lItemCnt + 1
        Next lCurrentCol

#If (DEBUG_MODE = 0) Then
    Debug.Print ConvertedDMMetadata.GetMetadataRow
#End If
        
        trgTs.WriteLine ConvertedDMMetadata.GetMetadataRow
        DoEvents
    Next lCurrentRow

    'EOFのテキストを追加する
    trgTs.WriteLine cESRD_EOF
 
    trgTs.Close
    
    MsgBox "Saab送付用 Converted DM用メタデータを保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set ConvertedDMMetadataSheet = Nothing
    Set ConvertedDMMetadatafile = Nothing
    Set ConvertedDMMetadata = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set ConvertedDMMetadataSheet = Nothing
    Set ConvertedDMMetadatafile = Nothing
    Set ConvertedDMMetadata = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub
