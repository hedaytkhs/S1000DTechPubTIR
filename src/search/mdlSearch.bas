Attribute VB_Name = "mdlSearch"
Option Explicit

Private m_tbxVendorSearch As String
Private m_tbxOpenDMFolder As String
Private m_tbxReferencedDMsFrom As String
Private m_tbxReferencingDMsTo As String


Private m_tbxZonesSearch As String
Private m_tbxAccessPointsSearch As String
Private m_tbxToolsSearch As String
Private m_tbxSuppliesSearch As String
Private m_tbxCircuitBreakersSearch As String

Public Enum SearchType
    ShowInfo = 0
    SendToClipboard = 1
    ShowMultiItem = 2
End Enum

Public Enum CBSearchType
    CBNumber = 1
    CBName = 2
    PowerSource = 3
    ConnectedBUS = 4
    CBPLocation = 5
    System = 6
    Unknown = 99
End Enum

Public Enum ToolSearchType
    MiTNumber = 1
    ToolNumber = 2
    ToolName = 3
    ManufactureCode = 4
    Unknown = 99
End Enum

Public Enum ConsumableSearchType
    NameOrNumber = 1
    Unknown = 99
End Enum

Public Type tpCBSearch
    CBSearch As CBSearchType
    Keyword As String
End Type

Public Type tpToolSearch
    ToolSearch As ToolSearchType
    Keyword As String
End Type

Private tpCBSearchType As tpCBSearch


'***********************************************************************
' TIR-Supplies Search
'***********************************************************************
Sub btnSearchSupplies_onAction(ByRef myButton As IRibbonControl)
    Call SuppliesSearchFromTextBox
End Sub

Sub SuppliesSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxSuppliesSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call SuppliesSearchFromTextBox
    End If
End Sub

Private Sub SuppliesSearchFromTextBox()
    Unload frmConsumableInfo
    Call SearchConsumablesByKeyword(m_tbxSuppliesSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub SearchConsumablesByKeyword(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmConsumableInfo
    If Len(Keyword) = 0 Then Exit Sub
    
    Dim SearchConditions As clsConsumable
    Dim SearchResult As New Collection
    Set SearchConditions = New clsConsumable
    Dim DBQryConsumables As New clsDBQryConsumableInfo
    
    With SearchConditions
        .SupplyNumber = Keyword
        .SupplyName = Keyword
    End With
    
    Call DBQryConsumables.SearchConsumablesByKeyword(SearchConditions, SearchResult)
    
    If SearchResult.Count > 1 Then
        Call ShowConsumableSearchResult(SearchResult)
    ElseIf SearchResult.Count = 1 Then
        Call SearchConsumableInfoBySupplyNumber(SearchResult.Item(1).SupplyNumber, ShowInfo)
    Else
        MsgBox Keyword & " を検索しましたが、 " & vbCrLf & "検索条件に合致するConsumableは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameSuppliesSearch)
    End If
    
    Set DBQryConsumables = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub


Public Function SearchConsumableInfoBySupplyNumber(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    Dim DBQryConsumables As New clsDBQryConsumableInfo
    Dim RefDMs As New Collection
    Dim RevisionHistory As New Collection
    Dim sEnteredText As String

    Dim ConsumableInfo As clsConsumable
    Set ConsumableInfo = New clsConsumable

    ConsumableInfo.SupplyNumber = SearchString
    With DBQryConsumables
        If .MACSDatabaseNotFound Then
            MsgBox "MACS資料用TIR-DBが見つかりません." & vbCrLf & "MACS資料用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameSuppliesSearch)
            Exit Function
        ElseIf .SMDSDatabaseNotFound Then
            MsgBox "SMDS確認用TIR-DBが見つかりません." & vbCrLf & "SMDS確認用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameSuppliesSearch)
            Exit Function
        End If
        .SearchSupplyInfoFromTIRDB ConsumableInfo
        .SearchRefDMsBySupplyNumber ConsumableInfo, RefDMs
#If False Then
        .GetRevisionHistoryBySupplyNumber ConsumableInfo, RevisionHistory
#End If
    End With
    If ConsumableInfo.FoundInMACS Then
        If SearchType = 0 Then
            Call ShowFormWithConsumableInfo(ConsumableInfo, RefDMs, RevisionHistory)
        ElseIf SearchType = 1 Then
            Call SendConsumableInfoToClipboard(ConsumableInfo, RefDMs)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR MACS資料上に登録されていません.", GetAppCNST(AppNameSuppliesSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Then Exit Function
        SearchConsumableInfoBySupplyNumber sEnteredText, SearchType
    End If

    Set RefDMs = Nothing
    Set RevisionHistory = Nothing
    Set DBQryConsumables = Nothing
    Set ConsumableInfo = Nothing

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Private Function ShowConsumableSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long
If SearchResult.Count = 0 Then
    MsgBox "該当するConsumableが見つかりません."
ElseIf SearchResult.Count > 0 Then
    With frmListViewTIRSupplies
        i = 0
        .ListViewTIR.ListItems.Clear
        For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.SupplyNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.SupplyNumberType
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.SupplyName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.SupplyLongName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.LocallySuppliedFlag
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.ManufactureCode
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Source
            i = i + 1
        Next
        .Repaint
        .Show
    End With
End If

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function



Private Function ShowFormWithConsumableInfo(ByRef ConsumableInfo As clsConsumable, ByRef RefDMs As Collection, ByRef RevisionHistory As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant, vntRevisionHistory As Variant
    Dim i As Long, j As Long
    With frmConsumableInfo
        .txtSupplyNumber = ConsumableInfo.SupplyNumber
        .txtSupplyName = ConsumableInfo.SupplyName
        .txtLongName = ConsumableInfo.SupplyLongName
        .txtLocallySuppliedFlag = ConsumableInfo.LocallySuppliedFlag
        .txtManufactureCode.Value = ConsumableInfo.ManufactureCode
        .txtSource.Value = ConsumableInfo.Source

        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
         j = 0
         .ListViewRevHistory.ListItems.Clear
        For Each vntRevisionHistory In RevisionHistory
            If vntRevisionHistory.ActiveItem = True Then
                .ListViewRevHistory.ListItems.Add , , "YES"
            Else
                .ListViewRevHistory.ListItems.Add , , "NO"
            End If
'            .ListViewRevHistory.ListItems.Add , , vntRevisionHistory.ActiveItem
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.RevisionSequence
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.SupplyNumber
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.SupplyName
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ManufactureCode
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ShortName
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.SupplyLongName
            j = j + 1
        Next
        .btnSearch.Enabled = False
        .btnCopy.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendConsumableInfoToClipboard(ByRef ConsumableInfo As clsConsumable, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    With ConsumableInfo
    buf = "Supply Number: " & .SupplyNumber & vbCrLf _
            & "Supply Name: " & .SupplyName & vbCrLf _
            & "SupplyNumber Type: " & .SupplyNumberType & vbCrLf _
            & "Supply Long Name: " & .SupplyLongName & vbCrLf _
            & "ManufactureCode: " & .ManufactureCode & vbCrLf _
            & "Locally Supplied Flag: " & .LocallySuppliedFlag & vbCrLf _
            & "Comment: " & .Comment & vbCrLf _
            & "Source: " & .Source & vbCrLf _
            & "Remarks: " & .Remarks & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    End With
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' TIR-Tools Search
'***********************************************************************
Sub ToolSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxToolsSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call ToolSearchFromTextBox
    End If
End Sub

Sub btnSearchTools_onAction(ByRef myButton As IRibbonControl)
    Call ToolSearchFromTextBox
End Sub

Private Sub ToolSearchFromTextBox()
    Unload frmToolInfo
   
    Call ToolsByKeyword(m_tbxToolsSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub ToolsByKeyword(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmToolInfo
    If Len(Keyword) = 0 Then Exit Sub
    Call SearchToolInfo(GetToolTIRSearchType(Keyword).Keyword, 0, GetToolTIRSearchType(Keyword).ToolSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Function GetToolTIRSearchType(ByVal Keyword As String) As tpToolSearch
    On Error GoTo errHandler
    With GetToolTIRSearchType
        .ToolSearch = MiTNumber
        .Keyword = Keyword
    End With
    
    With GetToolTIRSearchType
        .ToolSearch = ToolNumber
        .Keyword = Keyword
    End With
    
    If Len(Keyword) = 7 And UCase(Left(Keyword, 3)) = "MIT" And IsNumeric(Mid(Keyword, 4, 4)) Then
        GetToolTIRSearchType.ToolSearch = 1
        GetToolTIRSearchType.Keyword = Mid(Keyword, 4, 4)
        Exit Function
    ElseIf Len(Keyword) = 4 And IsNumeric(Keyword) Then
        GetToolTIRSearchType.ToolSearch = 1
        GetToolTIRSearchType.Keyword = Keyword
        Exit Function
    Else
        GetToolTIRSearchType.ToolSearch = 1
'        GetToolTIRSearchType.ToolSearch = 2
        GetToolTIRSearchType.Keyword = Keyword
        Exit Function
    
    End If
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchToolInfo(ByRef SearchString As String, ByVal ActionSelected As SearchType, ByVal SearchCategory As ToolSearchType) As Boolean
    On Error GoTo errHandler
    Dim SearchConditions As clsTool
    Dim SearchResult As New Collection
    Set SearchConditions = New clsTool
    
    Select Case SearchCategory
    Case 1
        SearchConditions.MiTNumber = SearchString
        Call SearchToolsByMiTNumber(SearchConditions, SearchResult)
        Exit Function
    Case 2
        SearchConditions.ToolNumber = SearchString
        Call SearchToolInfoByToolNumber(SearchString, ActionSelected)
'        Call SearchToolInfoByToolNumber(SearchString, ShowInfo)
        Exit Function
    Case 99
    End Select
    
    Set SearchConditions = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchToolsByMiTNumber(ByRef ToolSearchConditions As clsTool, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim DBQryTools As New clsDBQryToolInfo
    
    Call DBQryTools.SearchToolsByMiTNumber(ToolSearchConditions, SearchResult)
    
    If SearchResult.Count > 1 Then
        Call ShowToolsSearchResult(SearchResult)
    ElseIf SearchResult.Count = 1 Then
        Call SearchToolInfoByToolNumber(SearchResult.Item(1).ToolNumber, ShowInfo)
    Else
        MsgBox ToolSearchConditions.MiTNumber & " を検索しましたが、 " & vbCrLf & "検索条件に合致するToolは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
    End If

    
    Set DBQryTools = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchToolInfoByToolNumber(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    Dim DBQryTool As New clsDBQryToolInfo
    Dim RefDMs As New Collection
    Dim RevisionHistory As New Collection
    Dim sEnteredText As String

    Dim ToolInfo As clsTool
    Set ToolInfo = New clsTool

    ToolInfo.ToolNumber = SearchString
    With DBQryTool
        If .MACSDatabaseNotFound Then
            MsgBox "MACS資料用TIR-DBが見つかりません." & vbCrLf & "MACS資料用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
            Exit Function
        ElseIf .SMDSDatabaseNotFound Then
            MsgBox "SMDS確認用TIR-DBが見つかりません." & vbCrLf & "SMDS確認用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
            Exit Function
        End If
        .SearchToolInfoFromTIRDB ToolInfo
        .SearchRefDMsByToolNumber ToolInfo, RefDMs
        .GetRevisionHistoryByToolNumber ToolInfo, RevisionHistory
    End With
    If ToolInfo.FoundInMACS Then
        If SearchType = 0 Then
            Call ShowFormWithToolInfo(ToolInfo, RefDMs, RevisionHistory)
        ElseIf SearchType = 1 Then
            Call SendToolInfoToClipboard(ToolInfo, RefDMs)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR MACS資料上に登録されていません.", GetAppCNST(AppNameToolsSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Then Exit Function
        SearchToolInfoByToolNumber sEnteredText, SearchType
    End If

    Set RevisionHistory = Nothing
    Set RefDMs = Nothing
    Set DBQryTool = Nothing
    Set ToolInfo = Nothing

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Private Function ShowToolsSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long
If SearchResult.Count = 0 Then
    MsgBox "該当するToolが見つかりません."
ElseIf SearchResult.Count > 0 Then
    With frmListViewTIRTOOLS
        i = 0
        .ListViewTIR.ListItems.Clear
        For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.ToolNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.MiTNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.ToolName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.ToolLongName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.ManufactureCode
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Source
            i = i + 1
        Next
        .Repaint
        .Show
    End With
End If
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithToolInfo(ByRef ToolInfo As clsTool, ByRef RefDMs As Collection, ByRef RevisionHistory As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim vntRevisionHistory As Variant
    Dim i As Long, j As Long
    With frmToolInfo
        .txtMiTNumber.Value = ToolInfo.MiTNumber
        .txtToolNumber.Value = ToolInfo.ToolNumber
        .txtToolName.Value = ToolInfo.ToolName
        .txtToolLongName.Value = ToolInfo.ToolLongName
        .txtManufactureCode.Value = ToolInfo.ManufactureCode
        .txtSource.Value = ToolInfo.Source
        
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
         j = 0
         .ListViewRevHistory.ListItems.Clear
        For Each vntRevisionHistory In RevisionHistory
            If vntRevisionHistory.ActiveItem = True Then
                .ListViewRevHistory.ListItems.Add , , "YES"
            Else
                .ListViewRevHistory.ListItems.Add , , "NO"
            End If
'            .ListViewRevHistory.ListItems.Add , , vntRevisionHistory.ActiveItem
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.RevisionSequence
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ToolNumber
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ToolName
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ManufactureCode
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ShortName
            .ListViewRevHistory.ListItems(j + 1).ListSubItems.Add , , vntRevisionHistory.ToolLongName
            j = j + 1
        Next
       
        .btnSearch.Enabled = False
        .btnCopy.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendToolInfoToClipboard(ByRef ToolInfo As clsTool, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    With ToolInfo
    buf = "MiT Number: " & .MiTNumber & vbCrLf _
            & "Tool Number: " & .ToolNumber & vbCrLf _
            & "Tool Name: " & .ToolName & vbCrLf _
            & "Tool Long Name: " & .ToolLongName & vbCrLf _
            & "ManufactureCode: " & .ManufactureCode & vbCrLf _
            & "Comment: " & .Comment & vbCrLf _
            & "Source: " & .Source & vbCrLf _
            & "Remarks: " & .Remarks & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    End With
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With
   
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'***********************************************************************
' TIR-Circuit Breakers Search
'***********************************************************************
Sub btnSearchCircuitBreakers_onAction(ByRef myButton As IRibbonControl)
    Call CircuitBreakersSearchFromTextBox
End Sub

Sub CircuitBreakersSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxCircuitBreakersSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call CircuitBreakersSearchFromTextBox
    End If
End Sub

Private Sub CircuitBreakersSearchFromTextBox()
    Unload frmCBInfo
    Call CircuitBreakersByKeyword(m_tbxCircuitBreakersSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub CircuitBreakersByKeyword(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmCBInfo
    If Len(Keyword) = 0 Then Exit Sub
    Call SearchCircuitBreakerInfo(GetCBTIRSearchType(Keyword).Keyword, 0, GetCBTIRSearchType(Keyword).CBSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Function GetCBTIRSearchType(ByVal Keyword As String) As tpCBSearch
    On Error GoTo errHandler
    With GetCBTIRSearchType
        .CBSearch = 99
        .Keyword = Keyword
    End With
    
    If Len(Keyword) = 7 And Left(Keyword, 2) = "CB" And IsNumeric(Mid(Keyword, 3, 5)) Then
        GetCBTIRSearchType.CBSearch = 1
        Exit Function
    ElseIf Len(Keyword) = 4 Then
        If Keyword = "LICC" Or Keyword = "RICC" Or Keyword = "EICC" Or Keyword = "AICC" Then
            GetCBTIRSearchType.CBSearch = 3
            Exit Function
        End If
    ElseIf Len(Keyword) = 5 And IsNumeric(Keyword) Then
        GetCBTIRSearchType.CBSearch = 1
        GetCBTIRSearchType.Keyword = "CB" & Keyword
        Exit Function
    ElseIf Len(Keyword) = 5 And (Keyword = "SPDA1" Or Keyword = "SPDA2") Then
        GetCBTIRSearchType.CBSearch = 3
        Exit Function
    ElseIf Mid(Keyword, 2, 1) = "-" Then
        If Keyword = "L-CBP" Or Keyword = "R-CBP" Then
            GetCBTIRSearchType.CBSearch = 3
            Exit Function
        End If
        GetCBTIRSearchType.CBSearch = 5
        Exit Function
    ElseIf Len(Keyword) < 8 And Left(Keyword, 2) = "CB" And IsNumeric(Right(Keyword, Len(Keyword) - 2)) Then
        GetCBTIRSearchType.CBSearch = 5
        Exit Function
    ElseIf Keyword = "EXT PNL" Then
        GetCBTIRSearchType.CBSearch = 3
        Exit Function
    ElseIf InStr(1, Keyword, " ") = 3 Then
        If Left(Keyword, 2) = "AC" Or Left(Keyword, 2) = "DC" Then
            GetCBTIRSearchType.CBSearch = 4
            Exit Function
        End If
    ElseIf Keyword = "MAIN BATT" Or Keyword = "APU BATT" Or Keyword = "ESS TRU PWR" Then
        GetCBTIRSearchType.CBSearch = 4
        Exit Function
    End If

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchCircuitBreakerInfo(ByRef SearchString As String, ByVal SearchType As SearchType, ByVal CBSearchType As CBSearchType) As Boolean
    On Error GoTo errHandler
    Dim CBSearchConditions As clsCircuitBreaker
    Dim SearchResult As New Collection
    Set CBSearchConditions = New clsCircuitBreaker
    
    Select Case CBSearchType
    Case 1
        SearchCircuitBreakerInfoByCBNumber SearchString, SearchType
        Exit Function
    Case 2
        CBSearchConditions.CBName = SearchString
        Call AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
        Exit Function
    Case 3
        CBSearchConditions.PowerSource = SearchString
        Call AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
        Exit Function
    Case 4
        CBSearchConditions.ConnectedBUS = SearchString
        Call AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
        Exit Function
    Case 5
        CBSearchConditions.CBPLocation = SearchString
        Call AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
        Exit Function
    Case 6
        CBSearchConditions.System = SearchString
        Call AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
        Exit Function
    Case 99
    End Select
    
    Set CBSearchConditions = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function AdvancedSearchCircuitBreaker(ByRef CBSearchConditions As clsCircuitBreaker, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim DBQryCircuitBreakers As New clsDBQryCBInfo
    
    Call DBQryCircuitBreakers.AdvancedSearchCircuitBreaker(CBSearchConditions, SearchResult)
    
'    If SearchResult.Count = 0 Then
'        MsgBox "検索条件に合致するCircuit Breakerは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameCircuitBreakersSearch)
'        Exit Function
'    Else
        Call ShowCircuitBreakersSearchResult(SearchResult)
'    End If
    
    Set DBQryCircuitBreakers = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowCircuitBreakersSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long
    With frmListViewTIRCB
        
        i = 0
        .ListViewTIR.ListItems.Clear
       For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.CBNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.CBName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.CBClass
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.PowerSource
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.ConnectedBUS
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.CBPLocation
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.System
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Comment
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Source
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.MACSDocument
            i = i + 1
        Next
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchCircuitBreakerInfoByCBNumber(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    Dim DBQryCircuitBreakers As New clsDBQryCBInfo
    Dim RefDMs As New Collection
    Dim sEnteredText As String

    Dim CBInfo As clsCircuitBreaker
    Set CBInfo = New clsCircuitBreaker

    CBInfo.CBNumber = SearchString
    With DBQryCircuitBreakers
        If .MACSDatabaseNotFound Then
            MsgBox "MACS資料用TIR-DBが見つかりません." & vbCrLf & "MACS資料用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameCircuitBreakersSearch)
            Exit Function
        ElseIf .SMDSDatabaseNotFound Then
            MsgBox "SMDS確認用TIR-DBが見つかりません." & vbCrLf & "SMDS確認用TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameCircuitBreakersSearch)
            Exit Function
        End If
        .SearchCircuitBreakerInfoFromTIRDB CBInfo
        .SearchRefDMsByCBNumber CBInfo, RefDMs
    End With
    If CBInfo.FoundInMACS Then
        If SearchType = 0 Then
            Call ShowFormWithCBInfo(CBInfo, RefDMs)
        ElseIf SearchType = 1 Then
            Call SendCBInfoToClipboard(CBInfo, RefDMs)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR MACS資料上に登録されていません.", GetAppCNST(AppNameCircuitBreakersSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Then Exit Function
        SearchCircuitBreakerInfoByCBNumber sEnteredText, SearchType
    End If

    Set RefDMs = Nothing
    Set DBQryCircuitBreakers = Nothing
    Set CBInfo = Nothing

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithCBInfo(ByRef CBInfo As clsCircuitBreaker, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim i As Long
    With frmCBInfo
        .txtCBNumber.Value = CBInfo.CBNumber
        .cboCBClass.Value = CBInfo.CBClass
        .txtCBName.Value = CBInfo.CBName
        .txtCBPLocation.Value = CBInfo.CBPLocation
        .txtComment.Value = CBInfo.Comment
        .txtConnectedBUS.Value = CBInfo.ConnectedBUS
        .txtMACSDocument.Value = CBInfo.MACSDocument
        .txtSource.Value = CBInfo.Source
        .cboPowerSource = CBInfo.PowerSource
        .cboSystem = CBInfo.System
        
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
        .btnSearch.Enabled = False
        .btnCopy.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendCBInfoToClipboard(ByRef CBInfo As clsCircuitBreaker, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    buf = "CB Number: " & CBInfo.CBNumber & vbCrLf _
            & "CB Name: " & CBInfo.CBName & vbCrLf _
            & "CB Class: " & CBInfo.CBClass & vbCrLf _
            & "PowerSource: " & CBInfo.PowerSource & vbCrLf _
            & "ConnectedBUS: " & CBInfo.ConnectedBUS & vbCrLf _
            & "CBPLocation: " & CBInfo.CBPLocation & vbCrLf _
            & "Comment: " & CBInfo.Comment & vbCrLf _
            & "Source: " & CBInfo.Source & vbCrLf _
            & "MACSDocument: " & CBInfo.MACSDocument & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

#If False Then
    MsgBox "CB Information were sent to Clipboard." & vbCrLf & vbCrLf & buf, vbInformation + vbOKOnly, "SendCBInfoToClipboard"
#End If
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' TIR-Zone Search
'***********************************************************************

Sub btnSearchZones_onAction(ByRef myButton As IRibbonControl)
    Call ZonesSearchFromTextBox
End Sub

Sub ZonesSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxZonesSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call ZonesSearchFromTextBox
    End If
End Sub

Private Sub ZonesSearchFromTextBox()
    Unload frmZoneInfo
    Call ZonesSearchByZoneNumber(m_tbxZonesSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub ZonesSearchByZoneNumber(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmZoneInfo
    If Len(Keyword) = 0 Then Exit Sub
    Dim SearchConditions As clsZone
    Dim SearchResult As New Collection
    Set SearchConditions = New clsZone
    
    SearchConditions.ZoneNumber = Keyword
    Call SearchZonesByKeyword(SearchConditions, SearchResult)
    
    Set SearchConditions = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Function SearchZonesByKeyword(ByRef SearchConditions As clsZone, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim DBQryZones As New clsDBQryZoneInfo

    Call DBQryZones.SearchToolsByKeyword(SearchConditions, SearchResult)

    If SearchResult.Count > 1 Then
        Call ShowZoneSearchResult(SearchResult)
    ElseIf SearchResult.Count = 1 Then
        Call SearchZoneInfo(SearchResult.Item(1).ZoneNumber, ShowInfo)
    Else
        MsgBox SearchConditions.ZoneNumber & " を検索しましたが、 " & vbCrLf & "検索条件に合致するZoneは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
    End If

    Set DBQryZones = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchZoneInfo(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    
    Dim DBQryTIRZones As New clsDBQryZoneInfo
    Dim RefDMs As New Collection
    Dim sEnteredText As String
    
    Dim ZoneInfo As clsZone
    Set ZoneInfo = New clsZone
    
    ZoneInfo.ZoneNumber = SearchString
    With DBQryTIRZones
        .SearchZoneInfoFromTIRDB ZoneInfo
        .SearchRefDMsByZoneNumber ZoneInfo, RefDMs
        If .DatabaseNotFound Then
            MsgBox "TIR-DBが見つかりません." & vbCrLf & "TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameZonesSearch)
            Exit Function
        End If
    End With
    If ZoneInfo.FoundInSMDS Then
        If SearchType = ShowInfo Then
            Call ShowFormWithZoneInfo(ZoneInfo, RefDMs)
        ElseIf SearchType = SendToClipboard Then
            Call SendZoneInfoToClipboard(ZoneInfo, RefDMs)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR DB上に登録されていません.", GetAppCNST(AppNameZonesSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Then Exit Function
        SearchZoneInfo sEnteredText, SearchType
    End If
    
    Set RefDMs = Nothing
    Set DBQryTIRZones = Nothing
    Set ZoneInfo = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowZoneSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long

If SearchResult.Count = 0 Then
    MsgBox "該当するZoneが見つかりません."
ElseIf SearchResult.Count > 0 Then
    With frmListViewTIRZONE
        i = 0
        .ListViewTIR.ListItems.Clear
        For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.ZoneNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Description
            i = i + 1
        Next
        .Repaint
        .Show
    End With
End If
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithZoneInfo(ByRef ZoneInfo As clsZone, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim i As Long
    With frmZoneInfo
        .txtZoneNumber.Value = ZoneInfo.ZoneNumber
        .txtZoneDescription = ZoneInfo.Description
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
        .btnCopy.Enabled = True
        .btnSearch.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendZoneInfoToClipboard(ByRef ZoneInfo As clsZone, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    buf = "Zone Number: " & ZoneInfo.ZoneNumber & vbCrLf _
            & "Description: " & ZoneInfo.Description & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

'***********************************************************************
' TIR-AccessPoints Search
'***********************************************************************

Sub btnSearchAccessPoints_onAction(ByRef myButton As IRibbonControl)
    Call AccessPointsSearchFromTextBox
End Sub

Sub AccessPointsSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxAccessPointsSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call AccessPointsSearchFromTextBox
    End If
End Sub

Private Sub AccessPointsSearchFromTextBox()
    Unload frmAPInfo
    Call AccessPointsSearchByPanelNumber(m_tbxAccessPointsSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub AccessPointsSearchByPanelNumber(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmAPInfo
    If Len(Keyword) = 0 Then Exit Sub
    Dim SearchConditions As clsAccessPanel
    Dim SearchResult As New Collection
    Set SearchConditions = New clsAccessPanel
    
    SearchConditions.PanelNumber = Keyword
    Call SearchAccessPointsByKeyword(SearchConditions, SearchResult)
    
    Set SearchConditions = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Function SearchAccessPointsByKeyword(ByRef SearchConditions As clsAccessPanel, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim DBQryAccessPoints As New clsDBQryAPInfo

    Call DBQryAccessPoints.SearchAccessPanelsByKeyword(SearchConditions, SearchResult)

    If SearchResult.Count > 1 Then
        Call ShowAccessPanelSearchResult(SearchResult)
    ElseIf SearchResult.Count = 1 Then
        Call SearchAccessPointsInfo(SearchResult.Item(1).PanelNumber, ShowInfo)
    Else
        MsgBox SearchConditions.PanelNumber & " を検索しましたが、 " & vbCrLf & "検索条件に合致するAccess Panelは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
    End If

    Set DBQryAccessPoints = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowAccessPanelSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long

If SearchResult.Count = 0 Then
    MsgBox "該当するPanelが見つかりません."
ElseIf SearchResult.Count > 0 Then
    With frmListViewTIRAP
        i = 0
        .ListViewTIR.ListItems.Clear
        For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.PanelNumber
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Description
            i = i + 1
        Next
        .Repaint
        .Show
    End With
End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function
    
Public Function SearchAccessPointsInfo(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    
    Dim DBQryTIRAccessPoints As New clsDBQryAPInfo
    Dim RefDMs As New Collection
    Dim sEnteredText As String
    
    Dim PanelInfo As clsAccessPanel
    Set PanelInfo = New clsAccessPanel
    
    PanelInfo.PanelNumber = SearchString
    With DBQryTIRAccessPoints
        If .DatabaseNotFound Then
            MsgBox "TIR-DBが見つかりません." & vbCrLf & "TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameZonesSearch)
            Exit Function
        End If
        .SearchPanelInfoFromTIRDB PanelInfo
        .SearchRefDMsByPanelNumber PanelInfo, RefDMs
    End With
    If PanelInfo.FoundInSMDS Then
        If SearchType = 0 Then
            Call ShowFormWithPanelInfo(PanelInfo, RefDMs)
        ElseIf SearchType = 1 Then
            Call SendPanelInfoToClipboard(PanelInfo, RefDMs)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR DB上に登録されていません.", GetAppCNST(AppNameAccessPointsSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Then Exit Function
        SearchAccessPointsInfo sEnteredText, SearchType
    End If
    
    Set RefDMs = Nothing
    Set DBQryTIRAccessPoints = Nothing
    Set PanelInfo = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithPanelInfo(ByRef PanelInfo As clsAccessPanel, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim i As Long
    With frmAPInfo
        .txtPanelNumber = PanelInfo.PanelNumber
        .txtPanelDescription = PanelInfo.Description
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
        .btnCopy.Enabled = True
        .btnSearch.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendPanelInfoToClipboard(ByRef PanelInfo As clsAccessPanel, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    buf = "Panel Number: " & PanelInfo.PanelNumber & vbCrLf _
            & "Description: " & PanelInfo.Description & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

#If False Then
    MsgBox "Panel Information were sent to Clipboard." & vbCrLf & vbCrLf & buf, vbInformation + vbOKOnly, "SendZoneInfoToClipboard"
#End If
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' ShowReferencedDMsFrom DMC
'***********************************************************************

Sub btnSearchReferencedDMsFrom_onAction(ByRef myButton As IRibbonControl)
    Call ShowReferencedDMsFromTextBox
End Sub

Sub RefDMFromTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxReferencedDMsFrom = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call ShowReferencedDMsFromTextBox
    End If
End Sub

Private Sub ShowReferencedDMsFromTextBox()
    Call DMCCheckAndShowReferencedDMs(m_tbxReferencedDMsFrom, ShowInfo)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub ReferencedDMSearchCurrentCell()
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim CellCount As Long
    For CellCount = 1 To Selection.Count
        Call DMCCheckAndShowReferencedDMs(CStr(Selection(CellCount).Cells.Value), ShowInfo)
    Next CellCount
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub DMCCheckAndShowReferencedDMs(ByRef Keyword As String, ByVal SearchType As SearchType)
    On Error GoTo errHandler
'    Dim sEnteredText As String
    Dim DMCInSelectedXlsCell As clsDMC
    Dim DMCInInputText As clsDMC
    Set DMCInSelectedXlsCell = New clsDMC
    Set DMCInInputText = New clsDMC
    On Error Resume Next
    
    With DMCInSelectedXlsCell
        .DMC = Keyword
        If Not .IsValidDMC Then
'            sEnteredText = InputBox("選択されたセルの値はMRJ用の正しいDMCではありません." & vbCrLf & "正しいDMCを入力してください.", GetAppCNST(AppNameShowReferencedDMs), Keyword)
'            If Keyword = sEnteredText Then Exit Sub
'            With DMCInInputText
'                .DMC = sEnteredText
'                If Not .IsValidDMC Then Exit Sub
'            End With
'            SearchReferencedDMsFromSMDSDB sEnteredText, SearchType
            Exit Sub
        End If
    End With
        
    Call SearchReferencedDMsFromSMDSDB(Keyword, SearchType)
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    Exit Sub
errHandler:
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function SearchReferencedDMsFromSMDSDB(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    
    Dim SubjectDMC As clsDMC
    Set SubjectDMC = New clsDMC
    Dim DBQryRefDMs As New clsDBQryRefDM
    Dim RefDMs As New Collection
    Set DBQryRefDMs = New clsDBQryRefDM
    SubjectDMC.DMC = SearchString
    
    With DBQryRefDMs
        Call .SearchDMInfoOfSubjectDMC(SubjectDMC)
        Call .SearchRefDMs(SubjectDMC, RefDMs)
    End With
    If SearchType = ShowInfo Then
        Call ShowFormWithReferencedDMs(SubjectDMC, RefDMs)
    ElseIf SearchType = SendToClipboard Then
        Call SendReferencedDMsToClipboard(SubjectDMC, RefDMs)
    End If
    
    Set SubjectDMC = Nothing
    Set DBQryRefDMs = Nothing
    Set RefDMs = Nothing
    Exit Function
errHandler:
    Set RefDMs = Nothing
    Set DBQryRefDMs = Nothing
    Set SubjectDMC = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithReferencedDMs(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim i As Long
    With frmReferencedDMs
        .txtDMC = DMInfo.DMC
        .txtTechName = DMInfo.TechName
        .txtInfoName = DMInfo.InfoName
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
        .btnCopy.Enabled = True
'        .btnSearch.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendReferencedDMsToClipboard(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    buf = "DMC: " & DMInfo.DMC & vbCrLf _
            & "Techname: " & DMInfo.TechName & vbCrLf _
            & "Infoname: " & DMInfo.InfoName & vbCrLf & vbCrLf _
            & "Referenced DMs: " & vbCrLf
    
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function




'***********************************************************************
' ShowReferencingDMsTo DMC
'***********************************************************************

Sub btnSearchReferencingDMsTo_onAction(ByRef myButton As IRibbonControl)
    Call ShowReferencingDMsToTextBox
End Sub

Sub RefDMToTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxReferencingDMsTo = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call ShowReferencingDMsToTextBox
    End If
End Sub

Private Sub ShowReferencingDMsToTextBox()
    Call DMCCheckAndShowReferencingDMs(m_tbxReferencingDMsTo, ShowInfo)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub ReferencingDMSearchCurrentCell()
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim CellCount As Long
    For CellCount = 1 To Selection.Count
        Call DMCCheckAndShowReferencingDMs(CStr(Selection(CellCount).Cells.Value), 0)
    Next CellCount
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub DMCCheckAndShowReferencingDMs(ByRef Keyword As String, ByVal SearchType As SearchType)
    On Error GoTo errHandler
'    Dim sEnteredText As String
    Dim DMCInSelectedXlsCell As clsDMC
    Dim DMCInInputText As clsDMC
    Set DMCInSelectedXlsCell = New clsDMC
    Set DMCInInputText = New clsDMC
    On Error Resume Next
    
    With DMCInSelectedXlsCell
        .DMC = Keyword
        If Not .IsValidDMC Then
'            sEnteredText = InputBox("選択されたセルの値はMRJ用の正しいDMCではありません." & vbCrLf & "正しいDMCを入力してください.", GetAppCNST(AppNameShowReferencingDMs), Keyword)
'            If Keyword = sEnteredText Then Exit Sub
'            With DMCInInputText
'                .DMC = sEnteredText
'                If Not .IsValidDMC Then Exit Sub
'            End With
'            Call SearchDMsReferencingSubjectDMCFromSMDSDB(sEnteredText, SearchType)
            Exit Sub
        End If
    End With
        
    Call SearchDMsReferencingSubjectDMCFromSMDSDB(Keyword, SearchType)
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    Exit Sub
errHandler:
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Function SearchDMsReferencingSubjectDMCFromSMDSDB(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    
    Dim SubjectDMC As clsDMC
    Set SubjectDMC = New clsDMC
    Dim DBQryRefDMs As New clsDBQryRefDM
    Dim RefDMs As New Collection
    Set DBQryRefDMs = New clsDBQryRefDM
    SubjectDMC.DMC = SearchString
    
    With DBQryRefDMs
        Call .SearchDMInfoOfSubjectDMC(SubjectDMC)
        Call .SearchReferencingDMs(SubjectDMC, RefDMs)
    End With
    If SearchType = ShowInfo Then
        Call ShowFormWithDMsReferencingSubjectDMC(SubjectDMC, RefDMs)
    ElseIf SearchType = SendToClipboard Then
        Call SendReferencingDMsToClipboard(SubjectDMC, RefDMs)
    End If
    
    Set SubjectDMC = Nothing
    Set DBQryRefDMs = Nothing
    Set RefDMs = Nothing
    Exit Function
errHandler:
    Set RefDMs = Nothing
    Set DBQryRefDMs = Nothing
    Set SubjectDMC = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithDMsReferencingSubjectDMC(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntDMRef As Variant
    Dim i As Long
    With frmReferencingDMs
        .txtDMC = DMInfo.DMC
        .txtTechName = DMInfo.TechName
        .txtInfoName = DMInfo.InfoName
        .ListViewRefDMs.ListItems.Clear
        i = 0
        .ListViewRefDMs.ListItems.Clear
        For Each vntDMRef In RefDMs
            .ListViewRefDMs.ListItems.Add , , vntDMRef.DMC
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.TechName
            .ListViewRefDMs.ListItems(i + 1).ListSubItems.Add , , vntDMRef.InfoName
            i = i + 1
        Next
        .btnCopy.Enabled = True
'        .btnSearch.Enabled = True
        .Repaint
        .Show
    End With
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function SendReferencingDMsToClipboard(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    Dim vntDMRef As Variant
    buf = "DMC: " & DMInfo.DMC & vbCrLf _
            & "Techname: " & DMInfo.TechName & vbCrLf _
            & "Infoname: " & DMInfo.InfoName & vbCrLf & vbCrLf _
            & "Referencing DMs: " & vbCrLf
    
    For Each vntDMRef In RefDMs
        With vntDMRef
            buf = buf & .DMC & ", " & .TechName & " - " & .InfoName & vbCrLf
        End With
    Next
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' Open DM Folder
'***********************************************************************

Sub btnOpenDMFolder_onAction(ByRef myButton As IRibbonControl)
    Call OpenDMFolderFromTextBox
End Sub

Sub DMCTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxOpenDMFolder = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call OpenDMFolderFromTextBox
    End If
End Sub

Private Sub OpenDMFolderFromTextBox()
Call DMCCheckAndOpenDMFolder(m_tbxOpenDMFolder)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub DMSearchCurrentCell()
    On Error GoTo errHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim CellCount As Long
    For CellCount = 1 To Selection.Count
        If Selection(CellCount).Cells.Value = "" Then
            Exit For
        End If
        Call DMCCheckAndOpenDMFolder(CStr(Selection(CellCount).Cells.Value))
    Next CellCount
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub DMCCheckAndOpenDMFolder(ByRef Keyword As String)
    On Error GoTo errHandler
'    Dim sEnteredText As String
    Dim DMCInSelectedXlsCell As clsDMC
    Dim DMCInInputText As clsDMC
    Set DMCInSelectedXlsCell = New clsDMC
    Set DMCInInputText = New clsDMC
    On Error Resume Next
    
    With DMCInSelectedXlsCell
        .DMC = Keyword
        If Not .IsValidDMC Then
'            sEnteredText = InputBox("選択されたセルの値はMRJ用の正しいDMCではありません." & vbCrLf & "正しいDMCを入力してください.", GetAppCNST(AppNameDMFolderOpen), Keyword)
'            If Keyword = sEnteredText Then Exit Sub
'            With DMCInInputText
'                .DMC = sEnteredText
'                If Not .IsValidDMC Then Exit Sub
'            End With
'            OpenDMTTFolder sEnteredText
            Exit Sub
        End If
    End With
        
    Call OpenDMTTFolder(Keyword)
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    Exit Sub
errHandler:
    Set DMCInSelectedXlsCell = Nothing
    Set DMCInInputText = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Sub


Public Function OpenDMTTFolder(ByRef SearchString As String) As Boolean
    On Error GoTo errHandler
    
    Dim Datamodule As clsDataModule
    Set Datamodule = New clsDataModule
    Dim DMFolderSetting As clsConfigDatamodule
    Set DMFolderSetting = New clsConfigDatamodule
    
    With Datamodule
        .TTBaseFolderPath = DMFolderSetting.TTBaseFolderPath
        .DMC = SearchString
    End With
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(Datamodule.DMCPath) Then
        Call Shell("explorer """ & Datamodule.DMCPath & "", vbNormalFocus)
    Else
        MsgBox SearchString & "の保存フォルダは存在しません.", vbOKOnly + vbInformation, GetAppCNST(AppNameDMFolderOpen)
    End If
    
    Set Datamodule = Nothing
    Set DMFolderSetting = Nothing
    Exit Function
errHandler:
    Set Datamodule = Nothing
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' Vendor Search
'***********************************************************************

Sub btnSearchVendor_onAction(ByRef myButton As IRibbonControl)
    Call VenvorSearchFromTextBox
End Sub

Sub VendorSearchTextOnChange(ByRef myButton As IRibbonControl, ByRef Text As String)
    m_tbxVendorSearch = Text
    If Trim(Text) = "" Then
        Exit Sub
    Else
        Call VenvorSearchFromTextBox
    End If
End Sub

Private Sub VenvorSearchFromTextBox()
    Call VenvorSearchByVendorCode(m_tbxVendorSearch)
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub VenvorSearchCurrentCell()
    If ActiveSheet Is Nothing Then Exit Sub
    Call VenvorSearchByVendorCode(CStr(Selection(1).Cells.Value))
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Sub VenvorSearchByVendorCode(ByRef Keyword As String)
    On Error GoTo errHandler
    Unload frmVendorInfo
    If Len(Keyword) = 0 Then Exit Sub
    Dim SearchConditions As clsVendorCode
    Dim SearchResult As New Collection
    Set SearchConditions = New clsVendorCode
    
    SearchConditions.VendorCode = Keyword
    SearchConditions.VendorName = Keyword
    Call SearchVendorsByKeyword(SearchConditions, SearchResult)
    
    Set SearchConditions = Nothing
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Public Function SearchVendorsByKeyword(ByRef SearchConditions As clsVendorCode, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim DBQryVendorInfo As New clsDBQryVendorInfo

    Call DBQryVendorInfo.SearchVendorsByKeyword(SearchConditions, SearchResult)

    If SearchResult.Count > 1 Then
        Call ShowVendorSearchResult(SearchResult)
    ElseIf SearchResult.Count = 1 Then
        Call ShowVendorInfo(SearchResult.Item(1).VendorCode, ShowInfo)
    Else
        MsgBox SearchConditions.VendorCode & " を検索しましたが、 " & vbCrLf & "検索条件に合致するVendorは見つかりませんでした.", vbOKOnly + vbInformation, GetAppCNST(AppNameToolsSearch)
    End If

    Set DBQryVendorInfo = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowVendorSearchResult(ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    Dim vntTIRItem As Variant
    Dim i As Long

If SearchResult.Count = 0 Then
    MsgBox "該当するPanelが見つかりません."
ElseIf SearchResult.Count > 0 Then
    With frmListViewTIREnterprise
        i = 0
        .ListViewTIR.ListItems.Clear
        For Each vntTIRItem In SearchResult
            .ListViewTIR.ListItems.Add , , vntTIRItem.VendorCode
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.AlternateCode
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.VendorName
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Country
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.City
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Street
            .ListViewTIR.ListItems(i + 1).ListSubItems.Add , , vntTIRItem.Source
            i = i + 1
        Next
        .Repaint
        .Show
    End With
End If
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function ShowVendorInfo(ByRef SearchString As String, ByVal SearchType As SearchType) As Boolean
    On Error GoTo errHandler
    
    Dim DBQryTIREnterprise As New clsDBQryVendorInfo
    Dim sEnteredText As String
    
    Dim VendorInfo As clsVendorCode
    Set VendorInfo = New clsVendorCode
    
    VendorInfo.VendorCode = SearchString
    With DBQryTIREnterprise
        If .DatabaseNotFound Then
            MsgBox "TIR-DBが見つかりません." & vbCrLf & "TIR-DBのパスが正しいか確認してください.", vbOKOnly + vbInformation, GetAppCNST(AppNameVendorSearch)
            Exit Function
        End If
        .SearchVendorInfoFromTIRDB VendorInfo
    End With
    
    If VendorInfo.HasVendorInfo Then
        If SearchType = ShowInfo Then
            Call ShowFormWithVendorInfo(VendorInfo)
        ElseIf SearchType = SendToClipboard Then
            Call SendVendorInfoToClipboard(VendorInfo)
        End If
    Else
        sEnteredText = InputBox("""" & SearchString & """ はTIR DB上に登録されていません.", GetAppCNST(AppNameVendorSearch), SearchString)
        If SearchString = sEnteredText Or Len(sEnteredText) = 0 Or Len(sEnteredText) <> 5 Then Exit Function
        ShowVendorInfo sEnteredText, SearchType
    End If
    
    Set DBQryTIREnterprise = Nothing
    Set VendorInfo = Nothing

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function ShowFormWithVendorInfo(ByRef VendorInfo As clsVendorCode) As Boolean
    With frmVendorInfo
        .txtVendorCode.Value = VendorInfo.VendorCode
        .txtVendorName.Value = VendorInfo.VendorName
        .txtCity.Value = VendorInfo.City
        .txtComment.Value = VendorInfo.Comment
        .txtCountry.Value = VendorInfo.Country
        .txtEMail.Value = VendorInfo.EMail
        .txtFAX.Value = VendorInfo.FAX
        .txtPhone.Value = VendorInfo.PhoneNumber
        .txtSource.Value = VendorInfo.Source
        .txtStreet.Value = VendorInfo.Street
        .txtURL.Value = VendorInfo.URL
        .txtZipCode = VendorInfo.ZIPCode
        .btnCopy.Enabled = True
        .btnSearch.Enabled = True
        .Show
    End With
End Function

Private Function SendVendorInfoToClipboard(ByRef VendorInfo As clsVendorCode) As Boolean
    On Error GoTo errHandler
    Dim buf As String
    buf = "Vendor Code: " & VendorInfo.VendorCode & vbCrLf _
            & "AlternateCode: " & VendorInfo.AlternateCode & vbCrLf _
            & "Vendor Name: " & VendorInfo.VendorName & vbCrLf _
            & "Country: " & VendorInfo.Country & vbCrLf _
            & "City: " & VendorInfo.City & vbCrLf _
            & "Street: " & VendorInfo.Street & vbCrLf _
            & "ZIP Code: " & VendorInfo.ZIPCode & vbCrLf _
            & "PhoneNumber: " & VendorInfo.PhoneNumber & vbCrLf _
            & "FAX: " & VendorInfo.FAX & vbCrLf _
            & "URL: " & VendorInfo.URL & vbCrLf _
            & "Source: " & VendorInfo.Source
    
    With New MSForms.DataObject
        .SetText buf
        .PutInClipboard
    End With

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


'***********************************************************************
' Creating and Updating Search Menu
'***********************************************************************

#If False Then
Private Sub ShowContextMenuName()
    Dim ContextMenus As CommandBars
    Dim ContextMenu As Variant
    Set ContextMenus = Application.CommandBars

    For Each ContextMenu In ContextMenus
        Debug.Print ContextMenu.name
    Next
    Set ContextMenus = Nothing
End Sub
#End If

Public Sub AddToCellMenuTechpubSearch()
    Dim ContextMenus As CommandBars
    Dim ContextMenu As Variant
    Set ContextMenus = Application.CommandBars
    Dim MenuVendorSearch As CommandBarControl
            
    Call DeleteFromCellMenuTechpubSearch
    For Each ContextMenu In ContextMenus
        If ContextMenu.name = "Cell" Or ContextMenu.name = "Query" Or ContextMenu.name = "List Range Popup" Or ContextMenu.name = "Query Layout" Then
            Set MenuVendorSearch = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)
            With MenuVendorSearch
                .Caption = "MRJ Techpub Search (&X)"
                .Tag = "Tag_CELL_CTRL_TIR_Search"
        
                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "DMSearchCurrentCell"
                    .FaceId = 32
                    .Caption = "Open Datamodule Folder for TT-Verification(&F)"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "ReferencedDMSearchCurrentCell"
                    .FaceId = 620
                    .Caption = "Show Referenced DMs from Current Cell(&R)"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "ReferencingDMSearchCurrentCell"
                    .FaceId = 313
                    .Caption = "Show DMs Referencing subject DMC(&T)"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "VenvorSearchCurrentCell"
                    .FaceId = 101
                    .Caption = "Search VendorCode(&V)"
                End With
    
            End With
            Set MenuVendorSearch = Nothing
        End If
    Next
    Set ContextMenus = Nothing
End Sub

Private Sub DeleteFromCellMenuTechpubSearch()
    Dim ContextMenus As CommandBars
    Dim ContextMenu As Variant
    Set ContextMenus = Application.CommandBars
    Dim MenuVendorSearch As CommandBarControl
    Dim ctrl As CommandBarControl
    For Each ContextMenu In ContextMenus
        If ContextMenu.name = "Cell" Or ContextMenu.name = "Query" Or ContextMenu.name = "List Range Popup" Or ContextMenu.name = "Query Layout" Then
            For Each ctrl In ContextMenu.Controls
                If ctrl.Tag = "Tag_CELL_CTRL_TIR_Search" Then
                    ctrl.Delete
                End If
            Next ctrl
        End If
    Next
    Set ContextMenus = Nothing
End Sub

