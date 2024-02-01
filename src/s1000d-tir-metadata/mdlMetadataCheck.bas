Attribute VB_Name = "mdlMetadataCheck"
Option Explicit

''''' Check String ''''''
Private Const cAPPTitle = "ESRD Metadata - Engineering Source Data"

Public Function GetValidCharForESRD(ByVal sMetadataString As String) As String
    sMetadataString = MetadataRemoveDoubleSpaces(sMetadataString)
    sMetadataString = MetadataRemoveEndWithSpace(sMetadataString)
    sMetadataString = RemoveSpaceOnlyValue(sMetadataString)
    sMetadataString = StrConv(sMetadataString, vbNarrow)
    
    If MultiBiteAlert(sMetadataString) Then
        MsgBox "ESRD上使用できない文字が含まれています" & vbCrLf & "文字化けの可能性がありますので確認してください" & vbCrLf & vbCrLf & "「" & sMetadataString & "」", vbExclamation, cAPPTitle
    End If
    
    GetValidCharForESRD = sMetadataString
End Function

'全角文字のチェック
' 戻り値：全角文字あり--->true 、全角文字なし--->false
'
' 最終更新日：2011/6/16
' Saab Metadata旧プログラムから再利用
' 作成　高橋秀明
'
Public Function MultiBiteAlert(strUnicode As String) As Boolean
On Error GoTo errHandler
    Dim strANSI    As String
    Dim myLen      As Integer
    Dim myLenB     As Integer
    
    If strUnicode = "" Then
        Exit Function
    End If

    strANSI = StrConv(strUnicode, vbFromUnicode)
    
    myLen = Len(strUnicode)
    myLenB = LenB(strANSI)
    
    If InStr(1, strUnicode, Chr(63)) > 0 Then
'        MsgBox "？の文字が含まれています" & vbCrLf & "文字化けの可能性がありますので確認してください" & vbCrLf & vbCrLf & "「" & strUnicode & "」 -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    End If
    
    If myLen * 2 = myLenB Then
'        MsgBox "全角文字が含まれています！" & vbCrLf & "不要な文字を削除してください。" & vbCrLf & vbCrLf & "「" & strUnicode & "」 -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    ElseIf myLen = myLenB Then
'        MsgBox "半角文字だけです"
        MultiBiteAlert = False
        Exit Function
    Else
'        MsgBox "全角と半角が混じっています" & vbCrLf & "不要な文字を削除してください。" & vbCrLf & vbCrLf & "「" & strUnicode & "」 -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    End If
        
    
'    If Ascii_chk(strUnicode) = True Then
'        MsgBox "ASCII Code以外の文字が使われています。" & vbCrLf & "ASCII Codeの文字を使用してください。" & vbCrLf & vbCrLf & "「" & strUnicode & "」 -->" & rs.Address, vbExclamation, cAppName
'        MultiBiteAlert = True
'        Exit Function
'    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function RemoveSpaceOnlyValue(ByVal sMetadataText As String) As String
On Error GoTo errHandler
    If sMetadataText = "　" Or sMetadataText = " " Then
        MsgBox "半角スペース、全角スペースのみのセルを削除します！" & vbCrLf & vbCrLf & "「" & sMetadataText & "」", vbOKOnly, cAPPTitle
        sMetadataText = ""
    End If
    RemoveSpaceOnlyValue = sMetadataText
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function MetadataRemoveDoubleSpaces(ByVal sMetadataText As String) As String
On Error GoTo errHandler
    If InStr(1, sMetadataText, "  ", vbTextCompare) > 0 Then
        MsgBox "Double spaces appear in metadata text." & vbCrLf & "Please check again and correct text!" & vbCrLf & "Metadata text: """ & sMetadataText & """", vbOKOnly + vbExclamation, cAPPTitle
        sMetadataText = Replace(sMetadataText, "  ", " ")
    End If
    MetadataRemoveDoubleSpaces = sMetadataText
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function MetadataRemoveEndWithSpace(ByVal sMetadataText As String) As String
On Error GoTo errHandler
    If Right(sMetadataText, 1) = " " Then
        MsgBox "Metadata text ends with a space." & vbCrLf & "Please check again and correct text!" & vbCrLf & "Metadata text: """ & sMetadataText & """", vbOKOnly + vbExclamation, cAPPTitle
        sMetadataText = Left(sMetadataText, Len(sMetadataText) - 1)
    End If
    MetadataRemoveEndWithSpace = sMetadataText
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function
