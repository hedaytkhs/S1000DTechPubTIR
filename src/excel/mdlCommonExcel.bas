Attribute VB_Name = "mdlCommonExcel"
'****************************************************************************************
'
'    Excelシート操作用共通モジュール
'
'    MRJ Technical Publication Tool
'
' クリップボードの操作を含むため
' Microsoft Forms Object 2.0 Libralyの参照設定が必要
'
'    Hideaki Takahashi
'
'****************************************************************************************

Option Explicit

'//////////////////////////////////////////////////////////////
'
'For Activesheet
'
'2行目E列にデータがあると仮定して最後の行、列を取得
'//////////////////////////////////////////////////////////////


Public Function GetLastColOfThisSheet(oSheet As Worksheet) As Long
On Error GoTo errHandler
    Dim ActiveRange As Range
    Dim lCellsCnt As Long
    Set ActiveRange = oSheet.Range("E2").CurrentRegion
    lCellsCnt = ActiveRange.Cells.Count
    GetLastColOfThisSheet = ActiveRange.Cells(lCellsCnt).Column
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function GetLastRowOfThisSheet(oSheet As Worksheet) As Long
On Error GoTo errHandler
    Dim ActiveRange As Range
    Dim lCellsCnt As Long
    Set ActiveRange = oSheet.Range("E2").CurrentRegion
    lCellsCnt = ActiveRange.Cells.Count
    GetLastRowOfThisSheet = ActiveRange.Cells(lCellsCnt).Row
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function GetLastAddressOfThisSheet(oSheet As Worksheet) As Long
On Error GoTo errHandler
    Dim ActiveRange As Range
    Dim lCellsCnt As Long
    Set ActiveRange = oSheet.Range("E2").CurrentRegion
    lCellsCnt = ActiveRange.Cells.Count
    GetLastAddressOfThisSheet = ActiveRange.Cells(lCellsCnt).Address
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function ColNum2Txt(lngColNum As Long) As String

  On Error GoTo errHandler

  Dim strAddr As String

  strAddr = Cells(1, lngColNum).Address(False, False)
  ColNum2Txt = Left(strAddr, Len(strAddr) - 1)

  Exit Function

errHandler:

  ColNum2Txt = ""

End Function


Public Sub RemoveCRFromExcelSheet(oSheet As Worksheet)
On Error GoTo errHandler
'
' CSV保存前のExcelファイルからセル内改行を削除する
'

'   セル内改行は、半角スペースで置換する
    oSheet.Cells.Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub



'//////////////////////////////////////////////////////////////
'
' クリップボードにタグなどの文字列を格納する
'
' 2014/03/20
' Hideaki Takahashi
'
' XMLタグ
' PSDRの連結済みの文字列など
'//////////////////////////////////////////////////////////////
Public Sub CopyStringIntoClipboard(sPasteString As String)
On Error GoTo errHandler

    ' 文字列のチェック
    '/////  NO CHECK  /////
    Dim bError As Boolean
    bError = False
        
    If bError Then
        ' エラーメッセージを表示
        MsgBox "クリップボードに値を格納できません!" & vbCrLf _
            & vbCrLf & sPasteString, vbOKOnly + vbExclamation, "MRJ Technical Publication Support Tools"
 
    End If
    
    ' クリップボードへ格納
    Dim objClipboard As New DataObject
    With objClipboard
        .SetText sPasteString
        .PutInClipboard
    End With
    
    ' メッセージを表示
    MsgBox "クリップボードに次の文字列を格納しました!" & vbCrLf _
            & vbCrLf & sPasteString, vbOKOnly + vbInformation, "MRJ Technical Publication Support Tools"

    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

