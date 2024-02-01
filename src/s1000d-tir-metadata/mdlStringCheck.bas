Attribute VB_Name = "mdlStringCheck"
Option Explicit

'//////////////////////////////////////////////////////////
'
' 文字列チェック機能
'
'//////////////////////////////////////////////////////////


Type tpChechCell
    Row As Long
    Column As Long
    Text As String
'    Msg As String
'    ExistError As Boolean
'    SuggestedText As String
End Type

Private Const cAppName = "TextChecker for MRJ Technical Publication"


Sub TextCheker()
    On Error GoTo errHandler
    
    Dim TextCheckLogFile As New clsMetadataFile
    Dim sMetadataFilename, sMetadataFullPath As String
    TextCheckLogFile.FileCategory = ESRDFileCategory.Errorlog
    sMetadataFilename = ActiveWorkbook.name & "_" & Replace(TextCheckLogFile.filename, ".csv", ".txt")
    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFilename
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
    
    
    'EOFのテキストを追加する
    trgTs.WriteLine "/////////////////////////////////////////////"
    trgTs.WriteLine "// " & cAppName
    trgTs.WriteLine "// チェック結果"
    trgTs.WriteLine "//"
    trgTs.WriteLine "/////////////////////////////////////////////" & vbCrLf
    
    Dim sCheckResult As String
    sCheckResult = StringCheck()
    
    trgTs.WriteLine sCheckResult

    trgTs.Close
    

    MsgBox "Text Checkerのチェック結果を保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
    
    Dim WSH As Object
    Set WSH = CreateObject("Wscript.Shell")
    WSH.Run Chr(34) & sMetadataFullPath & Chr(34), 3
    Set WSH = Nothing
    
    Set trgTs = Nothing
    Set fso = Nothing

    Set TextCheckLogFile = Nothing
    Exit Sub
errHandler:
    Set trgTs = Nothing
    Set fso = Nothing

    Set TextCheckLogFile = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub


'Sub TextCheker2()
'    On Error GoTo ErrHandler
'
'    Dim TextCheckLogFile As New clsMetadataFile
'    Dim sMetadataFileName, sMetadataFullPath As String
'    TextCheckLogFile.FileCategory = ESRDFileCategory.Errorlog
'    sMetadataFileName = ActiveWorkbook.name & "_" & Replace(TextCheckLogFile.FileName, ".csv", ".txt")
'    sMetadataFullPath = StrAddPathSeparator(ActiveWorkbook.Path) & sMetadataFileName
'
'    Dim fso As New FileSystemObject
'    Dim trgTs As TextStream
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set trgTs = fso.CreateTextFile(sMetadataFullPath, ForWriting)
'
'
'    'EOFのテキストを追加する
'    trgTs.WriteLine "/////////////////////////////////////////////"
'    trgTs.WriteLine "// " & cAppName
'    trgTs.WriteLine "// チェック結果"
'    trgTs.WriteLine "//"
'    trgTs.WriteLine "/////////////////////////////////////////////" & vbCrLf
'
'    Dim sCheckResult As String
'    sCheckResult = StringCheck2()
'
'    trgTs.WriteLine sCheckResult
'
'    trgTs.Close
'
'
'    MsgBox "Text Checkerのチェック結果を保存しました!" & vbCrLf & sMetadataFullPath, vbOKOnly, cAppName
'
'    Dim WSH As Object
'    Set WSH = CreateObject("Wscript.Shell")
'    WSH.Run Chr(34) & sMetadataFullPath & Chr(34), 3
'    Set WSH = Nothing
'
'    Set trgTs = Nothing
'    Set fso = Nothing
'
'    Set TextCheckLogFile = Nothing
'    Exit Sub
'ErrHandler:
'    Set trgTs = Nothing
'    Set fso = Nothing
'
'    Set TextCheckLogFile = Nothing
'    MsgBox Err.Number & ":" & Err.Description
'End Sub


Private Function StringCheck() As String
On Error GoTo errHandler
    
    Dim tpInitCell As tpChechCell
    With tpInitCell
        .Row = Selection(1).Row
        .Column = Selection(1).Column
    End With
    
    Dim tpEndCell As tpChechCell
    With tpEndCell
        .Row = Selection(Selection.Count).Row
        .Column = Selection(Selection.Count).Column
    End With

    Dim iCnt, jCnt As Long
    Dim sMsgCheckResult As String
    Dim bErrExists As Boolean
    bErrExists = False
    
    Dim tpCurrentCell As tpChechCell
    Dim sMsg As String
    Dim bExistError As Boolean
    Dim sSuggestedText As String
    Dim sErrorCheckLog As String
    
    For iCnt = tpInitCell.Column To tpEndCell.Column
        For jCnt = tpInitCell.Row To tpEndCell.Row
            With tpCurrentCell
                .Column = iCnt
                .Row = jCnt
                .Text = Cells(jCnt, iCnt)
            End With
            
            sMsg = ""
            bExistError = False
            sSuggestedText = ""
            
            
'            Debug.Print Cells(jCnt, iCnt)
            If ExistSpaceOnlyValue(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            If ExistDoubleSpaces(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            If EndWithSpace(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            If BeginWithSpace(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            If MultiBiteCheck(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            If ExistLineFeed(tpCurrentCell, sMsg, bExistError) Then
                ' RemoveXXXXXX
            End If
            
            
            If bExistError Then
                Dim lCnt As Long
                Dim sFixedText As String
                Dim ErrCnt As Long
                ErrCnt = 0
                sFixedText = ""
                For lCnt = 1 To Len(tpCurrentCell.Text) - 1
            '        Debug.Print Asc(Mid(tpCurrentCell.Text, lCnt, 1)) & " " & Mid(tpCurrentCell.Text, lCnt, 1)
                    If CLng(Asc(Mid(tpCurrentCell.Text, lCnt, 1))) = 63 Then
                        sFixedText = sFixedText & "?"
                        ErrCnt = ErrCnt + 1
                    Else
                        sFixedText = sFixedText & Mid(tpCurrentCell.Text, lCnt, 1)
                    End If
                Next lCnt
                
                If ErrCnt > 0 Then
'                    For lCnt = ErrCnt To 1 Step -1
                        sFixedText = sFixedText & Right(tpCurrentCell.Text, ErrCnt)
'                    Next lCnt
                End If
                sErrorCheckLog = sErrorCheckLog & vbCrLf & vbCrLf & "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
                "値：""" & sFixedText & """" & vbCrLf & sMsg
'                sErrorCheckLog = sErrorCheckLog & vbCrLf & vbCrLf & "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
'                "値：""" & tpCurrentCell.Text & """" & vbCrLf & sMsg
                'MsgBox "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
                "値：""" & tpCurrentCell.Text & """" & vbCrLf & sMsg _
                , vbOKOnly + vbExclamation, "TextChecker for MRJ Technical Publication"
            End If
        Next jCnt
    Next iCnt
    
    StringCheck = sErrorCheckLog
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'Private Sub Test()
'On Error GoTo ErrHandler
'    Dim buf As Range
'    Set buf = Application.InputBox(Prompt:="文字列チェックの対象セルの範囲を選択してください。", Type:=8, Title:="Text Checker")
'
'    Debug.Print buf.Address
'
'    Exit Sub
'ErrHandler:
'    MsgBox Err.Number & ":" & Err.Description
'End Sub
'
'Private Function StringCheck2() As String
'On Error GoTo ErrHandler
'
'    Dim buf As Range
'    Set buf = Application.InputBox(Prompt:="セルを選択してください。", Type:=8, Title:="Text Checker")
'
'    Dim tpInitCell As tpChechCell
'    With tpInitCell
'        .Row = Selection(1).Row
'        .Column = Selection(1).Column
'    End With
'
'    Dim tpEndCell As tpChechCell
'    With tpEndCell
'        .Row = Selection(Selection.Count).Row
'        .Column = Selection(Selection.Count).Column
'    End With
'
'    Dim iCnt, jCnt As Long
'    Dim sMsgCheckResult As String
'    Dim bErrExists As Boolean
'    bErrExists = False
'
'    Dim tpCurrentCell As tpChechCell
'    Dim sMsg As String
'    Dim bExistError As Boolean
'    Dim sSuggestedText As String
'    Dim sErrorCheckLog As String
'
'    For iCnt = tpInitCell.Column To tpEndCell.Column
'        For jCnt = tpInitCell.Row To tpEndCell.Row
'            With tpCurrentCell
'                .Column = iCnt
'                .Row = jCnt
'                .Text = Cells(jCnt, iCnt)
'            End With
'
'            sMsg = ""
'            bExistError = False
'            sSuggestedText = ""
'
'
''            Debug.Print Cells(jCnt, iCnt)
'            If ExistSpaceOnlyValue(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'            If ExistDoubleSpaces(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'            If EndWithSpace(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'            If BeginWithSpace(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'            If MultiBiteCheck(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'            If ExistLineFeed(tpCurrentCell, sMsg, bExistError) Then
'                ' RemoveXXXXXX
'            End If
'
'
'            If bExistError Then
'                Dim lCnt As Long
'                Dim sFixedText As String
'                Dim ErrCnt As Long
'                ErrCnt = 0
'                sFixedText = ""
'                For lCnt = 1 To Len(tpCurrentCell.Text) - 1
'            '        Debug.Print Asc(Mid(tpCurrentCell.Text, lCnt, 1)) & " " & Mid(tpCurrentCell.Text, lCnt, 1)
'                    If CLng(Asc(Mid(tpCurrentCell.Text, lCnt, 1))) = 63 Then
'                        sFixedText = sFixedText & "?"
'                        ErrCnt = ErrCnt + 1
'                    Else
'                        sFixedText = sFixedText & Mid(tpCurrentCell.Text, lCnt, 1)
'                    End If
'                Next lCnt
'
'                If ErrCnt > 0 Then
''                    For lCnt = ErrCnt To 1 Step -1
'                        sFixedText = sFixedText & Right(tpCurrentCell.Text, ErrCnt)
''                    Next lCnt
'                End If
'                sErrorCheckLog = sErrorCheckLog & vbCrLf & vbCrLf & "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
'                "値：""" & sFixedText & """" & vbCrLf & sMsg
''                sErrorCheckLog = sErrorCheckLog & vbCrLf & vbCrLf & "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
''                "値：""" & tpCurrentCell.Text & """" & vbCrLf & sMsg
'                'MsgBox "セル: " & ColNum2Txt(tpCurrentCell.Column) & "列" & tpCurrentCell.Row & "行" & vbCrLf & _
'                "値：""" & tpCurrentCell.Text & """" & vbCrLf & sMsg _
'                , vbOKOnly + vbExclamation, "TextChecker for MRJ Technical Publication"
'            End If
'        Next jCnt
'    Next iCnt
'
'    StringCheck = sErrorCheckLog
'    Exit Function
'ErrHandler:
'    MsgBox Err.Number & ":" & Err.Description
'End Function

Public Function ExistSpaceOnlyValue(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    ExistSpaceOnlyValue = False
    If tpCurrentCell.Text = "　" Or tpCurrentCell.Text = " " Then
        sMsg = sMsg & vbCrLf & "セルの値が、半角スペースもしくは全角スペースのみになっています！" & vbCrLf
        bExistError = True
        ExistSpaceOnlyValue = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function ExistDoubleSpaces(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    ExistDoubleSpaces = False
    If InStr(1, tpCurrentCell.Text, "  ", vbTextCompare) > 0 Then
        sMsg = sMsg & vbCrLf & "連続した半角スペースがセルに含まれています。" & vbCrLf
        bExistError = True
        ExistDoubleSpaces = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function EndWithSpace(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    EndWithSpace = False
    If Right(tpCurrentCell.Text, 1) = " " Then
        sMsg = sMsg & vbCrLf & "セルの終わりに半角スペースが含まれています。" & vbCrLf
        bExistError = True
        EndWithSpace = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function BeginWithSpace(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    BeginWithSpace = False
    If Left(tpCurrentCell.Text, 1) = " " Then
        sMsg = sMsg & vbCrLf & "セルの先頭に半角スペースが含まれています。" & vbCrLf
        bExistError = True
        BeginWithSpace = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function ExistLineFeed(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    ExistLineFeed = False
    If InStr(1, tpCurrentCell.Text, vbLf) > 0 Then
        sMsg = sMsg & "セルの中に改行が含まれています。" & vbCrLf
        bExistError = True
        ExistLineFeed = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function MultiBiteCheck(tpCurrentCell As tpChechCell, ByRef sMsg As String, ByRef bExistError As Boolean) As Boolean
On Error GoTo errHandler
    Dim strANSI    As String
    Dim myLen      As Integer
    Dim myLenB     As Integer
    
    If tpCurrentCell.Text = "" Then
        Exit Function
    End If

    strANSI = StrConv(tpCurrentCell.Text, vbFromUnicode)
    
    myLen = Len(tpCurrentCell.Text)
    myLenB = LenB(strANSI)
    
    If InStr(1, tpCurrentCell.Text, Chr(63)) > 0 Then
        sMsg = sMsg & vbCrLf & "文字化けの可能性がありますので確認してください" & vbCrLf
        bExistError = True
        MultiBiteCheck = True
    End If
    
    Dim lCnt As Long
    For lCnt = 1 To Len(tpCurrentCell.Text) - 1
'        Debug.Print Asc(Mid(tpCurrentCell.Text, lCnt, 1)) & " " & Mid(tpCurrentCell.Text, lCnt, 1)
        If CLng(Asc(Mid(tpCurrentCell.Text, lCnt, 1))) = 63 Then
            sMsg = sMsg & vbCrLf & "文字化けの可能性がありますので確認してください" & vbCrLf
            bExistError = True
            MultiBiteCheck = True
        End If
    Next lCnt
    
    If myLen * 2 = myLenB Then
        sMsg = sMsg & vbCrLf & "全角文字が含まれています！" & vbCrLf & "不要な文字を削除してください。" & vbCrLf
        bExistError = True
        MultiBiteCheck = True
    ElseIf myLen = myLenB Then
'        MsgBox "半角文字だけです"
        If Not bExistError Then
            MultiBiteCheck = False
        End If
    Else
        sMsg = sMsg & vbCrLf & "全角と半角が混じっています" & vbCrLf & "不要な文字を削除してください。" & vbCrLf
        bExistError = True
        MultiBiteCheck = True
    End If
        
    
'    If Ascii_chk(strUnicode) = True Then
'        MsgBox "ASCII Code以外の文字が使われています。" & vbCrLf & "ASCII Codeの文字を使用してください。" & vbCrLf & vbCrLf & "「" & strUnicode & "」 -->" & rs.Address, vbExclamation, cAppName
'        bExistError = True
'        MultiBiteCheck = True
'        Exit Function
'    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

