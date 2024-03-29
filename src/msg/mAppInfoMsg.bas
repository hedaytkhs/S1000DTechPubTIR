Attribute VB_Name = "mAppInfoMsg"

'****************************************************************************************
'
'    Application名、表示メッセージ等の定数取得用のモジュール
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

#Const DEBUG_MODE = 1

'---------------------------------------------------------------------------------
'Applicationの状態管理用のユーザー定義型
'---------------------------------------------------------------------------------
 Type tpArgument
    IsCancelled As Boolean
    IsCompleted As Boolean
    LogText As String
    ErrNumber As Long
    ErrDescription As String
    LogFilePath As String
    LogFileCreated As Boolean
    ItemCnt As Long
End Type


 Enum APP_CNST_ID
    Version = 100
    LastModified = 101
    AppName = 1
    AppNameTIRTools = 111
    AppNameTIRSupplies = 112
    AppNameTIREnterprise = 113
    AppNameTIRZones = 114
    AppNameTIRAccessPoints = 115
    AppNameTIRCircuitBreakers = 116
    AppNameVendorSearch = 120
    AppNameDMFolderOpen = 121
    AppNameZonesSearch = 122
    AppNameAccessPointsSearch = 123
    AppNameToolsSearch = 124
    AppNameSuppliesSearch = 125
    AppNameCircuitBreakersSearch = 126
    AppNameShowReferencedDMs = 127
    AppNameShowReferencingDMs = 128
    AppNameTIRSearch = 200
    
    ErrNumber = 2
    ErrDescription = 3
End Enum

 Function GetAppCNST(CNST_ID As APP_CNST_ID) As String
    Dim sRet As String
    Select Case CNST_ID
        Case APP_CNST_ID.Version: sRet = "Version 007.01"
        Case APP_CNST_ID.LastModified: sRet = "2016.10.19"
        Case APP_CNST_ID.AppName: sRet = "TIR Integration File作成ツール"
        Case APP_CNST_ID.AppNameTIRTools: sRet = "TIR Tools Integration File作成"
        Case APP_CNST_ID.AppNameTIRSupplies: sRet = "TIR Supplies Integration File作成"
        Case APP_CNST_ID.AppNameTIREnterprise: sRet = "TIR Enterprise Integration File作成"
        Case APP_CNST_ID.AppNameTIRZones: sRet = "TIR Zones Integration File作成"
        Case APP_CNST_ID.AppNameTIRAccessPoints: sRet = "TIR AccessPoints Integration File作成"
        Case APP_CNST_ID.AppNameTIRCircuitBreakers: sRet = "TIR CircuitBreakers Integration File作成"
        Case APP_CNST_ID.AppNameTIRSearch: sRet = "TIR登録情報検索"
        Case APP_CNST_ID.AppNameVendorSearch: sRet = "Vendor Code TIR登録情報検索"
        Case APP_CNST_ID.AppNameDMFolderOpen: sRet = "Open DM Folder"
        Case APP_CNST_ID.AppNameShowReferencedDMs: sRet = "Show Referenced DMs"
        Case APP_CNST_ID.AppNameShowReferencingDMs: sRet = "Show Referencing DMs"
        Case APP_CNST_ID.AppNameZonesSearch: sRet = "Zone TIR登録情報検索"
        Case APP_CNST_ID.AppNameAccessPointsSearch: sRet = "AccessPoints TIR登録情報検索"
        Case APP_CNST_ID.AppNameToolsSearch: sRet = "Tools TIR登録情報検索"
        Case APP_CNST_ID.AppNameSuppliesSearch: sRet = "Supplies TIR登録情報検索"
        Case APP_CNST_ID.AppNameCircuitBreakersSearch: sRet = "CircuitBreakers TIR登録情報検索"
        Case APP_CNST_ID.ErrNumber: sRet = vbCrLf & "ErrNumber: "
        Case APP_CNST_ID.ErrDescription: sRet = vbCrLf & "ErrDescription: "
    End Select
    GetAppCNST = sRet
End Function



