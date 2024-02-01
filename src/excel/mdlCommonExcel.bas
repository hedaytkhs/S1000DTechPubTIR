Attribute VB_Name = "mdlCommonExcel"
'****************************************************************************************
'
'    Excel�V�[�g����p���ʃ��W���[��
'
'    MRJ Technical Publication Tool
'
' �N���b�v�{�[�h�̑�����܂ނ���
' Microsoft Forms Object 2.0 Libraly�̎Q�Ɛݒ肪�K�v
'
'    Hideaki Takahashi
'
'****************************************************************************************

Option Explicit

'//////////////////////////////////////////////////////////////
'
'For Activesheet
'
'2�s��E��Ƀf�[�^������Ɖ��肵�čŌ�̍s�A����擾
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
' CSV�ۑ��O��Excel�t�@�C������Z�������s���폜����
'

'   �Z�������s�́A���p�X�y�[�X�Œu������
    oSheet.Cells.Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub



'//////////////////////////////////////////////////////////////
'
' �N���b�v�{�[�h�Ƀ^�O�Ȃǂ̕�������i�[����
'
' 2014/03/20
' Hideaki Takahashi
'
' XML�^�O
' PSDR�̘A���ς݂̕�����Ȃ�
'//////////////////////////////////////////////////////////////
Public Sub CopyStringIntoClipboard(sPasteString As String)
On Error GoTo errHandler

    ' ������̃`�F�b�N
    '/////  NO CHECK  /////
    Dim bError As Boolean
    bError = False
        
    If bError Then
        ' �G���[���b�Z�[�W��\��
        MsgBox "�N���b�v�{�[�h�ɒl���i�[�ł��܂���!" & vbCrLf _
            & vbCrLf & sPasteString, vbOKOnly + vbExclamation, "MRJ Technical Publication Support Tools"
 
    End If
    
    ' �N���b�v�{�[�h�֊i�[
    Dim objClipboard As New DataObject
    With objClipboard
        .SetText sPasteString
        .PutInClipboard
    End With
    
    ' ���b�Z�[�W��\��
    MsgBox "�N���b�v�{�[�h�Ɏ��̕�������i�[���܂���!" & vbCrLf _
            & vbCrLf & sPasteString, vbOKOnly + vbInformation, "MRJ Technical Publication Support Tools"

    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

