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
        MsgBox "ESRD��g�p�ł��Ȃ��������܂܂�Ă��܂�" & vbCrLf & "���������̉\��������܂��̂Ŋm�F���Ă�������" & vbCrLf & vbCrLf & "�u" & sMetadataString & "�v", vbExclamation, cAPPTitle
    End If
    
    GetValidCharForESRD = sMetadataString
End Function

'�S�p�����̃`�F�b�N
' �߂�l�F�S�p��������--->true �A�S�p�����Ȃ�--->false
'
' �ŏI�X�V���F2011/6/16
' Saab Metadata���v���O��������ė��p
' �쐬�@�����G��
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
'        MsgBox "�H�̕������܂܂�Ă��܂�" & vbCrLf & "���������̉\��������܂��̂Ŋm�F���Ă�������" & vbCrLf & vbCrLf & "�u" & strUnicode & "�v -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    End If
    
    If myLen * 2 = myLenB Then
'        MsgBox "�S�p�������܂܂�Ă��܂��I" & vbCrLf & "�s�v�ȕ������폜���Ă��������B" & vbCrLf & vbCrLf & "�u" & strUnicode & "�v -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    ElseIf myLen = myLenB Then
'        MsgBox "���p���������ł�"
        MultiBiteAlert = False
        Exit Function
    Else
'        MsgBox "�S�p�Ɣ��p���������Ă��܂�" & vbCrLf & "�s�v�ȕ������폜���Ă��������B" & vbCrLf & vbCrLf & "�u" & strUnicode & "�v -->" & Rs.Address, vbExclamation, cAppName
        MultiBiteAlert = True
        Exit Function
    End If
        
    
'    If Ascii_chk(strUnicode) = True Then
'        MsgBox "ASCII Code�ȊO�̕������g���Ă��܂��B" & vbCrLf & "ASCII Code�̕������g�p���Ă��������B" & vbCrLf & vbCrLf & "�u" & strUnicode & "�v -->" & rs.Address, vbExclamation, cAppName
'        MultiBiteAlert = True
'        Exit Function
'    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function RemoveSpaceOnlyValue(ByVal sMetadataText As String) As String
On Error GoTo errHandler
    If sMetadataText = "�@" Or sMetadataText = " " Then
        MsgBox "���p�X�y�[�X�A�S�p�X�y�[�X�݂̂̃Z�����폜���܂��I" & vbCrLf & vbCrLf & "�u" & sMetadataText & "�v", vbOKOnly, cAPPTitle
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
