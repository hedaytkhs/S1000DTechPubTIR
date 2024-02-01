Attribute VB_Name = "VendorCodeCheck"

Private Const cCheckType_TIR_SUPPLIES = 1
Private Const cCheckType_TIR_TOOLS = 2
Private Const cCheckType_SPARE_INTEGRATION = 3

Private Const cVendorCol_TIR_SUPPLIES = 7
Private Const cVendorCol_TIR_TOOLS = 6
Private Const cCheckTypeSPARE_INTEGRATION = 7

Private Const cVendorColNum = 3


Public Type tpIntegrationRowTextWithResult
    Text As String
    VendorCode As String * 5
    VendorCodeNotFoundInTIREnterprise As Boolean
    CouldNotSearchTIREnterpriseDB As Boolean
End Type

Type tpVendorTIR_Enterprise
    VendorCode As String
    Checked As Boolean
    ID As Long
End Type

Public Sub ShowMissingVendorCode(ByRef RowTextWithResult As tpIntegrationRowTextWithResult)
    With RowTextWithResult
        If .CouldNotSearchTIREnterpriseDB Then
            MsgBox "TIR-Enterprise データベースに接続できませんでした." & vbCrLf & _
                    "Integration File内のVendor CodeがTIR-Enterpriseに登録されているか手動で確認してください.", vbExclamation + vbOKOnly, "Vendorコードチェックエラー"
            Exit Sub
        End If
        If .VendorCodeNotFoundInTIREnterprise Then
            MsgBox .VendorCode & " は Enterprise TIRに登録されていません." & vbCrLf & _
                   "TIR-Enterprise に " & .VendorCode & " を追加してからこのTIRを送付してください.", vbExclamation + vbOKOnly, "Vendorコード登録エラー"
            
        End If
    End With
End Sub

Public Function CheckMissingVendorCode(ByRef tpVendor As tpIntegrationRowTextWithResult) As tpIntegrationRowTextWithResult
On Error GoTo errHandler
    Dim TIRSetting As clsConfigTIREnterprise
    Set TIRSetting = New clsConfigTIREnterprise
    
    ' DB検索できなかった場合
    If Not TIRSetting.TIRDatabaseExists Then
        tpVendor.CouldNotSearchTIREnterpriseDB = True
        Exit Function
    End If
    

    Dim adoCon As New ADODB.Connection
    Dim adoRs As ADODB.Recordset

    adoCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TIRSetting.TIRDatabasePath & ";"
    Set TIRSetting = Nothing
    
    adoCon.Open
    
    Set adoRs = adoCon.Execute("SELECT * FROM TBL_ID_CAGE_NON_CAGE_LIST WHERE VENDOR_CODE LIKE '" & tpVendor.VendorCode & "'")
    
'検索対象のVendorCodeのレコードの有無を取得
    If Not adoRs.EOF Then
        tpVendor.CouldNotSearchTIREnterpriseDB = False
        tpVendor.VendorCodeNotFoundInTIREnterprise = False
    Else
        tpVendor.CouldNotSearchTIREnterpriseDB = False
        tpVendor.VendorCodeNotFoundInTIREnterprise = True
    End If
    
    adoRs.Close
    adoCon.Close
    
    Set adoRs = Nothing
    Set adoCon = Nothing
    
    CheckMissingVendorCode = tpVendor
    Exit Function
errHandler:
' DB検索できなかった場合
    tpVendor.CouldNotSearchTIREnterpriseDB = True
    MsgBox Err.Number & ":" & Err.Description
End Function



Public Sub VendorCodeCheck_TIR_SUPPLIES()
    VendorCodeCheck cCheckType_TIR_SUPPLIES
End Sub

Sub VendorCodeCheck_TIR_TOOLS()
    VendorCodeCheck cCheckType_TIR_TOOLS
End Sub

Sub VendorCodeCheck_SPARE_INTEGRATION()
    VendorCodeCheck cCheckType_SPARE_INTEGRATION
End Sub


Sub VendorCodeCheck(lCheckType As Long)
On Error GoTo errHandler
    Dim lCheckCol As Long
    Dim oSheet As Worksheet
    
    lCheckCol = GetCheckVendorCol(lCheckType)
    
    Dim sTIR_EnterprisePath As String
    sTIR_EnterprisePath = GetTIR_EnterpriseFilePath
    
    Dim TIR_EnterpriseVendor() As tpVendorTIR_Enterprise
    TIR_EnterpriseVendor() = GetTIR_EnterpriseArray(sTIR_EnterprisePath)
    Set oSheet = ActiveSheet
    CompareExcelsheetWithTIR_Enterprise oSheet, lCheckCol, TIR_EnterpriseVendor, sTIR_EnterprisePath
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function GetCheckVendorCol(lCheckType As Long) As Long
On Error GoTo errHandler
    Select Case lCheckType
    Case cCheckType_TIR_SUPPLIES
        GetCheckVendorCol = cVendorCol_TIR_SUPPLIES
        Exit Function
    Case cCheckType_TIR_TOOLS
        GetCheckVendorCol = cVendorCol_TIR_TOOLS
        Exit Function
    Case cCheckType_SPARE_INTEGRATION
        GetCheckVendorCol = cCheckTypeSPARE_INTEGRATION
        Exit Function
    End Select
    GetCheckVendorCol = 1
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Function EffectiveRow(oSheet As Worksheet, lVendorCol As Long) As Long
   EffectiveRow = oSheet.Cells(Application.Rows.Count, lVendorCol).End(xlUp).Row
End Function

Private Sub CompareExcelsheetWithTIR_Enterprise(oSheet As Worksheet, lVendorCol As Long, VendorTIR_Enterprise() As tpVendorTIR_Enterprise, sTIR_EnterprisePath As String)
On Error GoTo errHandler
    Dim bExistUnknownVendorCode As Boolean
    bExistUnknownVendorCode = False
    Dim arrUnknownVendorCode() As String
    Dim lItemCount As Long
    lItemCount = 0
    
    FormWait.Label1.Caption = "このExcelシートに記載されたVendor CodeがTIR_Enterpriseにあるかをチェックしています。"
    FormWait.Label1.Font.Size = 11
    FormWait.Label2.Caption = "読み込まれた TIR Enterprise: " & vbCrLf & sTIR_EnterprisePath
    FormWait.Label3.Caption = ""
    
    FormWait.Show vbModeless
    
    'Copy
    CopyExcelSheetTemporaly
    
    'Sort
    Dim lLastRow As Long
    lLastRow = EffectiveRow(Worksheets("Sorting"), lVendorCol)
    SortByVendorCode Worksheets("Sorting"), lLastRow, lVendorCol
    
    'Check
    Dim lRow As Long
    Dim sTrgVendorCode As String
    Dim sPrevRowVendor As String
    For lRow = 2 To lLastRow
        sTrgVendorCode = Worksheets("Sorting").Cells(lRow, lVendorCol)
        sPrevRowVendor = Worksheets("Sorting").Cells(lRow - 1, lVendorCol)
        If sTrgVendorCode <> sPrevRowVendor Then
            If Not IsVendorExistInEnterpriseTIR(sTrgVendorCode, VendorTIR_Enterprise()) Then
                MsgBox "VendorCode : " & sTrgVendorCode & " は、TIR-Enterpriseに登録されていません！"
                FormWait.Label3.Caption = "VendorCode : " & sTrgVendorCode & " は、TIR-Enterpriseに登録されていません！"
                lItemCount = lItemCount + 1
                ReDim Preserve arrUnknownVendorCode(lItemCount)
                arrUnknownVendorCode(lItemCount) = sTrgVendorCode
                bExistUnknownVendorCode = True
            End If
        End If
    Next lRow
    
    'DeleteTemporalyExcelSheet
    Application.DisplayAlerts = False
    Worksheets("Sorting").Delete
    Application.DisplayAlerts = True
    If Not bExistUnknownVendorCode Then
        MsgBox "TIR-Entepriseに未登録のVendorCodeはありませんでした。"
    Else
        Dim sUnknownList As String
        sUnknownList = ""
        
        Dim lCnt As Long
        For lCnt = 1 To UBound(arrUnknownVendorCode)
            sUnknownList = sUnknownList & "VendorCode : " & arrUnknownVendorCode(lCnt) & vbCrLf
        Next lCnt
        
        MsgBox "TIR-Entepriseに未登録のVendorCodeがあります." & "TIR-Entepriseを確認し、未登録のVendor Codeを先にSMDSへ登録する必要があります。" & vbCrLf & vbCrLf & sUnknownList, vbCritical + vbOKOnly
    End If
    
    FormWait.Hide
    ActiveWorkbook.Save
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function IsVendorExistInEnterpriseTIR(sTrgVendorCode As String, VendorTIR_Enterprise() As tpVendorTIR_Enterprise) As Boolean
On Error GoTo errHandler
    Dim lSeachVendorCodePos As Long
    IsVendorExistInEnterpriseTIR = False
    For lSeachVendorCodePos = 1 To UBound(VendorTIR_Enterprise)
        If Not VendorTIR_Enterprise(lSeachVendorCodePos).Checked Then
                If UCase(sTrgVendorCode) = UCase(VendorTIR_Enterprise(lSeachVendorCodePos).VendorCode) Then
'                    MsgBox "VendorCode " & sTrgVendorCode & " is found in TIR_Enterprise"
                    FormWait.Label3.Caption = "VendorCode " & sTrgVendorCode & " is found in TIR_Enterprise" & vbCrLf & _
                                                "TIR Enterprise ID: " & VendorTIR_Enterprise(lSeachVendorCodePos).ID & " Vendor Code: " & VendorTIR_Enterprise(lSeachVendorCodePos).VendorCode
'                    VendorTIR_Enterprise(lSeachVendorCodePos).Checked = True
                    IsVendorExistInEnterpriseTIR = True
                Exit For
             End If
        End If
        DoEvents
    Next lSeachVendorCodePos
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub SortByVendorCode(oSheet As Worksheet, lLastRow As Long, lVendorCol As Long)
On Error GoTo errHandler
    Dim sVendorCol As String
    sVendorCol = getAlphabetFromColumn(lVendorCol)
    Dim sLastCol As String
    sLastCol = getAlphabetFromColumn(Application.Columns.Count)
    
    Columns(sVendorCol & ":" & sVendorCol).Select
    Range("A1:" & sLastCol & lLastRow).Sort Key1:=Range(sVendorCol & "2"), Order1:=xlAscending, Header:= _
        xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub CopyExcelSheetTemporaly()
On Error GoTo errHandler
    ActiveSheet.Select
    ActiveSheet.Copy before:=Worksheets(1)
    Sheets(1).Select
    Sheets(1).name = "Sorting"
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function GetTIR_EnterpriseFilePath() As String
On Error GoTo errHandler
    Dim sFilePath As Variant
    Dim vntGetFileName As Variant
    Dim bExistAuthorForm As Boolean
    bExistAuthorForm = False
    
    sFilePath = _
        Application.GetOpenFilename( _
             FileFilter:="Saab送付用 TIR-Enterpriseデータ(*.csv),*.csv" _
             , FilterIndex:=1 _
           , Title:="Saab送付用 TIR-Enterpriseデータ(タブ区切り)を選択してください。" _
           , MultiSelect:=False _
            )
    GetTIR_EnterpriseFilePath = sFilePath
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function



Private Function GetTIR_EnterpriseArray(sTIR_EnterprisePath As String) As tpVendorTIR_Enterprise()
On Error GoTo errHandler
    Dim arrReturn() As tpVendorTIR_Enterprise
    Dim fso, dollarSeparated_file As Object
    Dim sFolder, CurrentLineText, sAuthorFilePath As String
    Dim bExistEOF As Boolean
    Dim TSVFields() As String
    Dim lItemCount As Long
    Dim lLineCount As Long
    
    'AuthorMetadataファイルを開く
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dollarSeparated_file = fso.OpenTextFile(sTIR_EnterprisePath)
    sFolder = fso.GetParentFolderName(sTIR_EnterprisePath)
    
    bExistEOF = False
    CurrentLineText = ""
    lItemCount = 0
    lLineCount = 0
    
    Do Until (dollarSeparated_file.AtEndOfStream)
        CurrentLineText = dollarSeparated_file.ReadLine
        If InStr(1, CurrentLineText, "EOF", 1) Then
            bExistEOF = True
            Exit Do
        End If
        
        lLineCount = lLineCount + 1
        
        'ドル区切りでフィールドを分ける
        TSVFields = Split(CurrentLineText, "$")
        
        'Author File
        'タイトル行でなく、EOFでもなく、Author Fileである
        If IsVendorCode(lLineCount, TSVFields) Then
            lItemCount = lItemCount + 1
            ReDim Preserve arrReturn(lItemCount)
            arrReturn(lItemCount).VendorCode = TSVFields(cVendorColNum)
            arrReturn(lItemCount).ID = TSVFields(2)
            arrReturn(lItemCount).Checked = False
        End If
    
        DoEvents
    Loop
    
    GetTIR_EnterpriseArray = arrReturn
    Set metadata_file = Nothing
    Set fso = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function IsVendorCode(lItemCount As Long, TSVFields() As String) As Boolean
On Error GoTo errHandler
    IsVendorCode = False
    If lItemCount <> 1 And Len(TSVFields(cVendorColNum)) = 5 Then
        IsVendorCode = True
    ElseIf TSVFields(cVendorColNum) = "-" Then
        IsVendorCode = True
    End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function
