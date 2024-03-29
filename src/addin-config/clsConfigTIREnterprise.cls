VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigTIREnterprise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'3 列目にTIR-Enterprise用設定を保存
Private Const XlsColTIREnterprise As Long = 3
'7 列目にTIR共通設定を保存
Private Const XlsColTIRCommon As Long = 7

'3 列目にTIR -Enterprise用設定を保存
Private Enum XlsCellMACSTopFolder
    Row = 1
    col = XlsColTIREnterprise
End Enum

Private Enum XlsCellLatestExcelPath
    Row = 2
    col = XlsColTIREnterprise
End Enum

Private Enum XlsCellTIRDatabasePath
    Row = 3
    col = XlsColTIREnterprise
End Enum

Private Enum XlsCellIHS_URL
    Row = 4
    col = XlsColTIREnterprise
End Enum

'直前に作成したIntegrarion File情報
'チェックプログラム、メール作成に使用
'パス
Private Enum XlsCellIntegrationFilePath
    Row = 5
    col = XlsColTIREnterprise
End Enum
        
'ファイル名
Private Enum XlsCellIntegrationFileName
    Row = 6
    col = XlsColTIREnterprise
End Enum
'日付 yyyymmdd
Private Enum XlsCellIntegrationFileDate
    Row = 7
    col = XlsColTIREnterprise
End Enum
'アイテム数
Private Enum XlsCellItemCount
    Row = 8
    col = XlsColTIREnterprise
End Enum

Private Enum XlsCellIntegrationBackupFolder
    Row = 9
    col = XlsColTIREnterprise
End Enum


'7 列目にTIR共通設定を保存
Private Enum XlsCellIntegrationFileFolder
    Row = 1
    col = XlsColTIRCommon
End Enum

Private Enum XlsCellIntegrationFileCheckPGPath
    Row = 2
    col = XlsColTIRCommon
End Enum


Dim MacsTopFolder_ As String
Dim LatestExcelPath_ As String
Dim IntegrationFileFolder_ As String
Dim IntegrationFileBackupFolder_ As String
Dim IntegrationFileCheckPGPath_ As String
Dim IntegrationFilePath_ As String
Dim IntegrationFileName_ As String
Dim IntegrationFileDate_ As String
Dim ItemCount_ As Long

Dim TIRDatabasePath_ As String
Dim IHS_URL_ As String

Dim ShtConfig As Worksheet

Property Let MACSTopFolder(ByVal sMACSTopFolder As String)
    MacsTopFolder_ = sMACSTopFolder
    ShtConfig.Cells(XlsCellMACSTopFolder.Row, XlsCellMACSTopFolder.col).Value = MacsTopFolder_
End Property

Property Get MACSTopFolder() As String
    MACSTopFolder = ShtConfig.Cells(XlsCellMACSTopFolder.Row, XlsCellMACSTopFolder.col).Value
End Property

Property Let LatestExcelPath(ByVal sLatestExcelPath As String)
    LatestExcelPath_ = sLatestExcelPath
    ShtConfig.Cells(XlsCellLatestExcelPath.Row, XlsCellLatestExcelPath.col).Value = LatestExcelPath_
End Property

Property Get LatestExcelPath() As String
    LatestExcelPath = ShtConfig.Cells(XlsCellLatestExcelPath.Row, XlsCellLatestExcelPath.col).Value
End Property

Property Let TIRDatabasePath(ByVal sTIRDatabasePath As String)
    TIRDatabasePath_ = sTIRDatabasePath
    ShtConfig.Cells(XlsCellTIRDatabasePath.Row, XlsCellTIRDatabasePath.col).Value = TIRDatabasePath_
End Property

Property Get TIRDatabasePath() As String
    TIRDatabasePath = ShtConfig.Cells(XlsCellTIRDatabasePath.Row, XlsCellTIRDatabasePath.col).Value
End Property

Property Let IHS_URL(ByVal sIHS_URL As String)
    IHS_URL_ = sIHS_URL
    ShtConfig.Cells(XlsCellIHS_URL.Row, XlsCellIHS_URL.col).Value = IHS_URL_
End Property

Property Get IHS_URL() As String
    IHS_URL = ShtConfig.Cells(XlsCellIHS_URL.Row, XlsCellIHS_URL.col).Value
End Property


Property Let IntegrationFileFolder(ByVal sFolder As String)
    IntegrationFileFolder_ = sFolder
    ShtConfig.Cells(XlsCellIntegrationFileFolder.Row, XlsCellIntegrationFileFolder.col).Value = IntegrationFileFolder_
End Property

Property Get IntegrationFileFolder() As String
    IntegrationFileFolder = ShtConfig.Cells(XlsCellIntegrationFileFolder.Row, XlsCellIntegrationFileFolder.col).Value
End Property

Property Let IntegrationFileBackupFolder(ByVal sFolder As String)
    IntegrationFileBackupFolder_ = sFolder
    ShtConfig.Cells(XlsCellIntegrationBackupFolder.Row, XlsCellIntegrationBackupFolder.col).Value = IntegrationFileBackupFolder_
End Property

Property Get IntegrationFileBackupFolder() As String
    IntegrationFileBackupFolder = ShtConfig.Cells(XlsCellIntegrationBackupFolder.Row, XlsCellIntegrationBackupFolder.col).Value
End Property

Property Let IntegrationFileCheckPGPath(ByVal sPath As String)
    IntegrationFileCheckPGPath_ = sPath
    ShtConfig.Cells(XlsCellIntegrationFileCheckPGPath.Row, XlsCellIntegrationFileCheckPGPath.col).Value = IntegrationFileCheckPGPath_
End Property

Property Get IntegrationFileCheckPGPath() As String
    IntegrationFileCheckPGPath = ShtConfig.Cells(XlsCellIntegrationFileCheckPGPath.Row, XlsCellIntegrationFileCheckPGPath.col).Value
End Property

Property Let IntegrationFilePath(ByVal sPath As String)
    IntegrationFilePath_ = sPath
    ShtConfig.Cells(XlsCellIntegrationFilePath.Row, XlsCellIntegrationFilePath.col).Value = IntegrationFilePath_
End Property

Property Get IntegrationFilePath() As String
    IntegrationFilePath = ShtConfig.Cells(XlsCellIntegrationFilePath.Row, XlsCellIntegrationFilePath.col).Value
End Property

Property Let IntegrationFileName(ByVal sFilename As String)
    IntegrationFileName_ = sFilename
    ShtConfig.Cells(XlsCellIntegrationFileName.Row, XlsCellIntegrationFileName.col).Value = IntegrationFileName_
End Property

Property Get IntegrationFileName() As String
    IntegrationFileName = ShtConfig.Cells(XlsCellIntegrationFileName.Row, XlsCellIntegrationFileName.col).Value
End Property

Property Let IntegrationFileDate(ByVal sDate As String)
    IntegrationFileDate_ = sDate
    ShtConfig.Cells(XlsCellIntegrationFileDate.Row, XlsCellIntegrationFileDate.col).Value = IntegrationFileDate_
End Property

Property Get IntegrationFileDate() As String
    IntegrationFileDate = ShtConfig.Cells(XlsCellIntegrationFileDate.Row, XlsCellIntegrationFileDate.col).Value
End Property

Property Let ItemCount(ByVal lCount As Long)
    ItemCount_ = lCount
    ShtConfig.Cells(XlsCellItemCount.Row, XlsCellItemCount.col).Value = ItemCount_
End Property

Property Get ItemCount() As Long
    ItemCount = ShtConfig.Cells(XlsCellItemCount.Row, XlsCellItemCount.col).Value
End Property

Function Save() As Boolean
    ThisWorkbook.Save
End Function

Function DeleteIntegrationFile() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim iRet As Integer
    iRet = MsgBox("前回作成したIntegration Fileを削除しますか?", vbExclamation + vbYesNoCancel, "Integration File削除")
    If iRet = vbYes Then
        If fso.FileExists(IntegrationFilePath_) Then
            fso.DeleteFile (IntegrationFilePath_)
            Dim sLogFilePath As String
            sLogFilePath = Replace(IntegrationFilePath_, ".csv", ".log")
            If fso.FileExists(sLogFilePath) Then
                fso.DeleteFile (sLogFilePath)
            End If
            MsgBox "IntegrationFileを削除しました.", vbInformation + vbOKOnly, "Integration File削除"
        Else
            MsgBox "削除対象のIntegrationFileが見つかりません.", vbInformation + vbOKOnly, "Integration File削除"
        End If
        Set fso = Nothing
        
        ItemCount = 0
        IntegrationFilePath = ""
        IntegrationFileName = ""
        IntegrationFileDate = ""
        ThisWorkbook.Save
        DoEvents
    End If
End Function

Function IntegrationFileFolderExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IntegrationFileFolderExists = fso.FolderExists(IntegrationFileFolder_)
    Set fso = Nothing
End Function

Function IntegrationFileBackupFolderExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IntegrationFileBackupFolderExists = fso.FolderExists(IntegrationFileBackupFolder_)
    Set fso = Nothing
End Function

Function IntegrationFileCheckPGExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IntegrationFileCheckPGExists = fso.FileExists(IntegrationFileCheckPGPath_)
    Set fso = Nothing
End Function

Function TIRDatabaseExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    TIRDatabaseExists = fso.FileExists(TIRDatabasePath_)
    Set fso = Nothing
End Function

Function MACSTopFolderExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    MACSTopFolderExists = fso.FolderExists(MacsTopFolder_)
    Set fso = Nothing
End Function

Function LatestExcelExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    LatestExcelExists = fso.FileExists(LatestExcelPath_)
    Set fso = Nothing
End Function

Function IntegrationFileExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IntegrationFileExists = fso.FileExists(IntegrationFilePath_)
    Set fso = Nothing
End Function

Private Sub Class_Initialize()
    Set ShtConfig = ThisWorkbook.Worksheets(CNST_XLS_SHEET_CONFIG)
    With ShtConfig
        MacsTopFolder_ = .Cells(XlsCellMACSTopFolder.Row, XlsCellMACSTopFolder.col).Value
        LatestExcelPath_ = .Cells(XlsCellLatestExcelPath.Row, XlsCellLatestExcelPath.col).Value
        IntegrationFileFolder_ = .Cells(XlsCellIntegrationFileFolder.Row, XlsCellIntegrationFileFolder.col).Value
        IntegrationFileBackupFolder_ = .Cells(XlsCellIntegrationBackupFolder.Row, XlsCellIntegrationBackupFolder.col).Value
        IntegrationFileCheckPGPath_ = .Cells(XlsCellIntegrationFileCheckPGPath.Row, XlsCellIntegrationFileCheckPGPath.col).Value
        IntegrationFilePath_ = .Cells(XlsCellIntegrationFilePath.Row, XlsCellIntegrationFilePath.col).Value
        IntegrationFileName_ = .Cells(XlsCellIntegrationFileName.Row, XlsCellIntegrationFileName.col).Value
        IntegrationFileDate_ = .Cells(XlsCellIntegrationFileDate.Row, XlsCellIntegrationFileDate.col).Value
        ItemCount_ = .Cells(XlsCellItemCount.Row, XlsCellItemCount.col).Value
    End With
    TIRDatabasePath_ = ShtConfig.Cells(XlsCellTIRDatabasePath.Row, XlsCellTIRDatabasePath.col).Value
End Sub

Private Sub Class_Terminate()
    Set ShtConfig = Nothing
End Sub





