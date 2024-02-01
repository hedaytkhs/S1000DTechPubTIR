Attribute VB_Name = "mdlCommon"
'=========================================================
'
' Common Fucntion for MRJ Technical Publication Tools
'
'
' Version 003
' ///// 2014/04/17 /////
'
' Created
'
' by Hideaki Takahashi
'
' ///// Version History /////
'
'
'=========================================================
Option Explicit

#If (DEBUG_MODE = 0) _
 Or (DEBUG_MODE = 1) _
 Or (DEBUG_MODE = 10) _
 Or (DEBUG_MODE = 11) _
 Or (DEBUG_MODE = 12) _
 Or (DEBUG_MODE = 13) _
 Or (DEBUG_MODE = 90) Then

' **************************************************
' Japanese Language Environment Only
' **************************************************

' 2014/04/17
' LogFileをテキストエディタで開いて表示するSubプロシージャを追加
'
' 2014/04/02 OSの言語に依存しないように修正(したつもり)未検証
' デバッグモード時に新コードを利用(テスト中)
' デバッグモード時以外は旧コードを使用
'
' INI File操作用関数の追加(ReadIni) 2014/04/01
' INI File操作用関数の追加
' Windows Userｍ名取得関数を共通化
' 指定した名称のExcel Workbookが開かれているかを確認する関数を共通化
'
'
'
'　配布前の注意事項
' VBAプロジェクトの条件付き引数が適切に設定されていることを確認する。


#ElseIf (DEBUG_MODE = 100) _
 Or (DEBUG_MODE = 101) Then
 
' **************************************************
' English Language Environment Only
' **************************************************

' **************************************************
' MaxsaApp
' **************************************************
' DEBUG_MODE = 100 : For English Version Debugging
' DEBUG_MODE = 101: Ror MITAC Release Version
' **************************************************

#ElseIf (DEBUG_MODE = 110) _
 Or (DEBUG_MODE = 111) Then

' **************************************************
' Note:
' **************************************************
'
'
'
'
'
'
'
'

#End If




' **************************************************
' FOLLOWING PROCEDURES ARE AVAILABLE IN ANY DEBUG MODE
' **************************************************


Public Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Declare Function DeleteUrlCacheEntry Lib "wininet" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" _
(ByVal lpPathName As String) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                                        (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                                        (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'=========================================================
' Ini File
'=========================================================

Public Declare Function GetPrivateProfileIntA Lib "kernel32" (ByVal strSection As String, ByVal strKey As String, ByVal nDefault As Long, ByVal strFileName As String) As Long
Public Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal strApp As String, ByVal strKey As String, ByVal strDefault As String, ByVal strRead As String, ByVal nReadSize As Long, ByVal strFile As String) As Long

Public Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal strApp As String, ByVal strKey As String, ByVal strData As String, ByVal strFileName As String) As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" _
                         Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileInt Lib "kernel32" _
                         Alias "GetPrivateProfileIntA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As String, _
                          ByVal nDefault As Long, _
                          ByVal lpFileName As String) As Long



Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type



'=========================================================
' Ini File
'=========================================================
Public Function ReadIni(ByVal FName As String, ByVal sName As String, _
           ByVal KName As String, ByVal Default As String) As String
    Dim RtnCD As Long
    Dim RtnStr As String
    
    RtnStr = Space$(256)
    RtnCD = GetPrivateProfileString(sName, KName, Default, RtnStr, 255, _
    FName)
    
    If RtnCD > 0 Then
        If InStr(RtnStr, Chr$(0)) > 0 Then
            ReadIni = Left$(RtnStr, InStr(RtnStr, Chr$(0)) - 1)
        Else
            ReadIni = ""
        End If
    Else
        ReadIni = Default
    End If
    
End Function


'=========================================================
' Log Fileを開く
'=========================================================
Public Sub ShowResultLog(sMetadataFullPath As String)
    On Error GoTo errHandler
    Dim WSH As Object
    Set WSH = CreateObject("Wscript.Shell")
    WSH.Run Chr(34) & sMetadataFullPath & Chr(34), 3
    Set WSH = Nothing
    Exit Sub
errHandler:
    Set WSH = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub


Function GetFolder(Optional Msg) As String
    On Error GoTo errHandler
    Dim bInfo As BROWSEINFO, pPath As String
    Dim r As Long, X As Long, pos As Integer
    bInfo.pidlRoot = 0&
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select Folder"
    Else
        bInfo.lpszTitle = Msg
    End If
    bInfo.ulFlags = &H1
    X = SHBrowseForFolder(bInfo)
    pPath = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal pPath)
    If r Then
        pos = InStr(pPath, Chr$(0))
        GetFolder = Left(pPath, pos - 1)
    Else
        GetFolder = ""
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'=========================================================
' WindowsUserName
'=========================================================
Public Function GetCurrentWindowsUserName() As String
    On Error GoTo errHandler
' No Need to set reference to "Windows Script Host Object Model"
    Dim WshNetworkObject As Object
    Set WshNetworkObject = CreateObject("WScript.Network")
    
#If False Then
' Need to set reference to "Windows Script Host Object Model"
'    Dim WshNetworkObject As IWshRuntimeLibrary.WshNetwork
'    Set WshNetworkObject = New IWshRuntimeLibrary.WshNetwork
'    Debug.Print "UserName は" & WshNetworkObject.username
#End If
    
    GetCurrentWindowsUserName = WshNetworkObject.UserName
        
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'=========================================================
' FileSystemObject
'=========================================================

'=========================================================
' DeleteFile
'=========================================================
Public Sub DeleteFile(sDelFilePath As String)
    On Error GoTo errHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
 
    If fso.FileExists(sDelFilePath) Then
        fso.DeleteFile sDelFilePath
    End If
 
    Set fso = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub


'=========================================================
' Path
'=========================================================
Public Function StrAddPathSeparator(sPath As String) As String
    On Error GoTo errHandler
    Dim sPathSeparator As String
    sPathSeparator = Application.PathSeparator
'    Const cStrSeparator = "\"
    If Right(sPath, 1) <> sPathSeparator Then
        StrAddPathSeparator = sPath & sPathSeparator
    Else
        StrAddPathSeparator = sPath
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

' Old Version
' This works only Japanese Language Environment
'=========================================================
' Path
'=========================================================
'Public Function StrAddPathSeparator(sPath As String) As String
'    On Error GoTo errHandler
'    Const cStrSeparator = "\"
'    If Right(sPath, 1) <> "\" Then
'        StrAddPathSeparator = sPath & cStrSeparator
'    Else
'        StrAddPathSeparator = sPath
'    End If
'    Exit Function
'errHandler:
'    MsgBox Err.Number & ":" & Err.Description
'End Function






'*******************************************************
' Get Folder path from File full path
'*******************************************************
Public Function GetFolderPathFromFilepath(sFileFullPath As String) As String
    On Error GoTo errHandler
    Dim pos As Long
    pos = InStrRev(sFileFullPath, "\", , vbTextCompare)
    GetFolderPathFromFilepath = Left(sFileFullPath, pos)
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'*******************************************************
' Get Folder name from File full path
'*******************************************************
Public Function GetFolderNameFromFilepath(sFileFullPath As String) As String
    On Error GoTo errHandler
    Dim pos As Long
    pos = Len(sFileFullPath) - InStrRev(sFileFullPath, "\", , vbTextCompare)
    GetFolderNameFromFilepath = Right(sFileFullPath, pos)
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'*******************************************************
' Get extension from File full path
'*******************************************************
Public Function FullPath2Extension(strFPath As String) As String
    On Error GoTo errHandler
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FullPath2Extension = fso.GetExtensionName(strFPath)
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'*******************************************************
' Get file name from File full path
'*******************************************************
Public Function FullPath2FileName(strFPath As String)
On Error GoTo errHandler
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
FullPath2FileName = fso.GetFileName(strFPath)
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'*******************************************************
'
'*******************************************************
Public Function IsDigit(ByVal Value As String) As Boolean
On Error GoTo errHandler
    Dim K As Long

    If Len(Value) = 0 Then
        IsDigit = False
        Exit Function
    End If

    For K = 1 To Len(Value)
        If Not Mid(Value, K, 1) Like "[0-9]" Then Exit Function
    Next K

    IsDigit = True

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Public Sub SaveFolderExist(sFolderPath As String, sSaveFolderName As String)
On Error GoTo errHandler
    Dim fso As Object
    Dim sSaveFolderPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sFolderPath) = False Then
        fso.CreateFolder (sFolderPath)
    End If
    sSaveFolderPath = StrAddPathSeparator(sFolderPath) & sSaveFolderName
    If fso.FolderExists(sSaveFolderPath) = False Then
        fso.CreateFolder (sSaveFolderPath)
    End If
 
    Set fso = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub


'=========================================================
' ExcelSheet
'=========================================================
Public Function CommonGetLastRowColA(oSheet As Worksheet) As Long '  To Get Last row number in col A
On Error GoTo errHandler
   CommonGetLastRowColA = oSheet.Range("A65536").End(xlUp).Row
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function CommonGetLastRow(oSheet As Worksheet, sCol As String) As Long '  To Get Last row number in col A
On Error GoTo errHandler
   CommonGetLastRow = oSheet.Range(sCol & "65536").End(xlUp).Row
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function WorkbookIsOpen(sWorkbookName As String) As Boolean
On Error GoTo errHandler
    Dim wkbk As Workbook
    WorkbookIsOpen = False
    For Each wkbk In Workbooks
        If wkbk.name = sWorkbookName Then
            WorkbookIsOpen = True
            Exit Function
        End If
        DoEvents
    Next wkbk
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function ExistExcelSheet(sSheetName As String) As Boolean
On Error GoTo errHandler
    Dim ws As Worksheet
    ExistExcelSheet = False
    For Each ws In Worksheets
        If ws.name = sSheetName Then
            ExistExcelSheet = True
            Exit Function
        End If
        DoEvents
    Next ws
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Sub DeleteWorksheet(sSheetName As String)
On Error GoTo errHandler
     Application.DisplayAlerts = False
     Worksheets(sSheetName).Delete
     Application.DisplayAlerts = True
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub


Function getAlphabetFromColumn(col As Long) As String
On Error GoTo errHandler
 
    getAlphabetFromColumn = getSimpleAlpha("", col)
 
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function getSimpleAlpha(prefix As String, c As Long) As String
On Error GoTo errHandler
    Dim alpha As String
    Dim leftover As Long
    Dim residues As Long
 
    leftover = c \ 26
    If leftover > 0 Then
        residues = c Mod 26
        alpha = getSimpleAlpha(getSimpleAlpha("", leftover), residues)
    Else
        residues = 0
        alpha = Chr(c + 64)
    End If
 
    getSimpleAlpha = prefix & alpha
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'=========================================================
' Search
'=========================================================
Public Function CntStr(ByVal s As String, ByVal org As String) As Integer
On Error GoTo errHandler
  Dim i As Integer
  Dim j As Integer
  Dim K As Integer


  K = Len(org)
  i = 1
  j = 0
  Do
    i = InStr(i, s, org)
    If i > 0 Then
     i = i + K
     j = j + 1
    End If
    DoEvents
  Loop Until i = 0

  CntStr = j
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function CountString(TextString As String, findString As String) As Long
On Error GoTo errHandler
    Dim N As Long, cnt As Long
    N = InStr(1, TextString, findString)
    Do While N > 0
        cnt = cnt + 1
        N = InStr(N + 1, TextString, findString)
    Loop
    CountString = cnt
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function



'=========================================================
' WebQuery
'=========================================================

Public Function URLEncode(Value As String) As String
    On Error GoTo errHandler
    Dim sc, js As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JavaScript"
    Set js = sc.CodeObject
    URLEncode = js.encodeURIComponent(Value)

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'=========================================================
' FileDownload
'=========================================================







