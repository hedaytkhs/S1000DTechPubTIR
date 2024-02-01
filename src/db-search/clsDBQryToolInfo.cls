VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryToolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIRTools_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim RsRefRevisionHistory_ As ADODB.Recordset

Dim SMDS_DBConnection_ As ADODB.Connection
Dim MACS_DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC
Dim ItemRevision_ As clsTool
Dim TIR_Tool_Item_ As clsTool

Private Enum ErrType
    Connected = 0
    DB_NOT_FOUND = 1
End Enum
 
Dim SMDS_DB_Status As ErrType
Dim MACS_DB_Status As ErrType

Property Get SMDSDatabaseNotFound() As Boolean
    SMDSDatabaseNotFound = False
    If SMDS_DB_Status = 1 Then
        SMDSDatabaseNotFound = True
    End If
End Property

Property Get MACSDatabaseNotFound() As Boolean
    MACSDatabaseNotFound = False
    If MACS_DB_Status = 1 Then
        MACSDatabaseNotFound = True
    End If
End Property

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsTool) As String
    strSQL_ = "SELECT * FROM TIR_TOOLS WHERE ("
    Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False

    If SearchConditions.MiTNumber <> "" Then
'       strSQL_ = strSQL_ & "((TIR_TOOLS.[MIT NUMBER])='" & SearchConditions.MiTNumber & "') AND ((TIR_TOOLS.[ActiveItem])=YES)"
        strSQL_ = strSQL_ & "((TIR_TOOLS.[MIT NUMBER])='" & SearchConditions.MiTNumber & "') AND ((TIR_TOOLS.[ActiveItem])=YES)) OR " & _
                "(((TIR_TOOLS.[Tool Number]) Like '%" & SearchConditions.MiTNumber & "%') AND ((TIR_TOOLS.[ActiveItem])=YES)) OR " & _
                "(((TIR_TOOLS.[Tool Name]) Like '%" & SearchConditions.MiTNumber & "%') AND ((TIR_TOOLS.[ActiveItem])=YES)"
        flgOtherConditionExists = True
    End If

    strSQL_ = strSQL_ & ");"

    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function SearchToolsByMiTNumber(ByRef SearchConditions As clsTool, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    SearchToolsByMiTNumber = False
    If MACS_DB_Status = 1 Then Exit Function

    Set RsTIRTools_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(SearchConditions)

    RsTIRTools_.Open strSQL_, MACS_DBConnection_, adOpenStatic

    Do Until RsTIRTools_.EOF
        Set TIR_Tool_Item_ = New clsTool

        With TIR_Tool_Item_
            .FoundInMACS = True
            .MiTNumber = "" & RsTIRTools_.Fields("MIT NUMBER")
            .ToolNumber = "" & RsTIRTools_.Fields("Tool Number")
            .ToolName = "" & RsTIRTools_.Fields("Tool Name")
            .ToolLongName = "" & RsTIRTools_.Fields("Tool Long Name")
            .ManufactureCode = "" & RsTIRTools_.Fields("Manufacture Code")
            .Comment = "" & RsTIRTools_.Fields("Comment")
            .Remarks = "" & RsTIRTools_.Fields("Remarks")
            .Source = "" & RsTIRTools_.Fields("Source")
        End With
        SearchResult.Add TIR_Tool_Item_
        Set TIR_Tool_Item_ = Nothing
        SearchToolsByMiTNumber = True
        RsTIRTools_.MoveNext
        DoEvents
    Loop

    RsTIRTools_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function



Public Function SearchRefDMsByToolNumber(ByRef ToolInfo As clsTool, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMsByToolNumber = False
    If SMDS_DB_Status = 1 Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TOOLS_REF_DM_LIST_ALL_IC WHERE (((TOOLS_REF_DM_LIST_ALL_IC.[Tool Number])='" & ToolInfo.ToolNumber & "'));"

    RsRefDMs_.Open strSQL_, SMDS_DBConnection_, adOpenDynamic

    Do Until RsRefDMs_.EOF
        Set RefDMC_ = New clsDMC
        With RefDMC_
            .DMC = "" & RsRefDMs_.Fields("DMC")
            .TechName = "" & RsRefDMs_.Fields("TECHNAME")
            .InfoName = "" & RsRefDMs_.Fields("INFONAME")
        End With
        RefDMs.Add RefDMC_
        Set RefDMC_ = Nothing
        RsRefDMs_.MoveNext
        SearchRefDMsByToolNumber = True
        DoEvents
    Loop

    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function GetRevisionHistoryByToolNumber(ByRef ToolInfo As clsTool, ByRef RevisionHistory As Collection) As Boolean
    On Error GoTo errHandler
    GetRevisionHistoryByToolNumber = False
    If SMDS_DB_Status = 1 Then Exit Function
    Set RsRefRevisionHistory_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TIR_TOOLS WHERE (((TIR_TOOLS.[Tool Number])='" & ToolInfo.ToolNumber & "')) ORDER BY TIR_TOOLS.[Revision SEQ];"

    RsRefRevisionHistory_.Open strSQL_, SMDS_DBConnection_, adOpenDynamic

    Do Until RsRefRevisionHistory_.EOF
        Set ItemRevision_ = New clsTool
        With ItemRevision_
            .MiTNumber = "" & RsRefRevisionHistory_.Fields("MIT NUMBER")
            .ToolNumber = "" & RsRefRevisionHistory_.Fields("Tool Number")
            .ToolName = "" & RsRefRevisionHistory_.Fields("Tool Name")
            .ToolLongName = "" & RsRefRevisionHistory_.Fields("Tool Long Name")
            .ManufactureCode = "" & RsRefRevisionHistory_.Fields("Manufacture Code")
            .ActiveItem = "" & RsRefRevisionHistory_.Fields("ActiveItem")
            .RevisionSequence = "" & RsRefRevisionHistory_.Fields("Revision SEQ")
            .ShortName = "" & RsRefRevisionHistory_.Fields("Short Name")
        
        End With
        RevisionHistory.Add ItemRevision_
        Set ItemRevision_ = Nothing
        RsRefRevisionHistory_.MoveNext
        GetRevisionHistoryByToolNumber = True
        DoEvents
    Loop

    RsRefRevisionHistory_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchToolInfoFromTIRDB(ByRef ToolInfo As clsTool) As Boolean
    On Error GoTo errHandler
    SearchToolInfoFromTIRDB = False

    If MACS_DB_Status = 1 Then Exit Function

    Set RsTIRTools_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TIR_TOOLS WHERE (((TIR_TOOLS.[Tool Number])='" & ToolInfo.ToolNumber & "') AND ((TIR_TOOLS.[ActiveItem])=YES));"

    RsTIRTools_.Open strSQL_, MACS_DBConnection_, adOpenDynamic

    If RsTIRTools_.EOF = False Then
        With ToolInfo
            .FoundInMACS = True
            .MiTNumber = "" & RsTIRTools_.Fields("MIT NUMBER")
            .ToolNumber = "" & RsTIRTools_.Fields("Tool Number")
            .ToolName = "" & RsTIRTools_.Fields("Tool Name")
            .ToolLongName = "" & RsTIRTools_.Fields("Tool Long Name")
            .ManufactureCode = "" & RsTIRTools_.Fields("Manufacture Code")
            .Comment = "" & RsTIRTools_.Fields("Comment")
            .Remarks = "" & RsTIRTools_.Fields("Remarks")
            .Source = "" & RsTIRTools_.Fields("Source")
        End With
        SearchToolInfoFromTIRDB = True
    End If

    RsTIRTools_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Dim SMDS_DB_PATH As String
    Dim TIR_DB_PATH As String
    Dim TIRSetting As clsConfigTIRCircuitBreakers
    Set TIRSetting = New clsConfigTIRCircuitBreakers
    SMDS_DB_PATH = TIRSetting.SMDSTIRDatabasePath
    TIR_DB_PATH = TIRSetting.MACSTIRDatabasePath
    Set TIRSetting = Nothing
    Dim fso As New FileSystemObject
    If Not fso.FileExists(SMDS_DB_PATH) Then
        SMDS_DB_Status = 1
        Exit Sub
    End If
    If Not fso.FileExists(TIR_DB_PATH) Then
        MACS_DB_Status = 1
        Exit Sub
    End If
    
    Set SMDS_DBConnection_ = New ADODB.Connection
    SMDS_DBConnection_.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" _
                        & "Data Source=" & SMDS_DB_PATH & ""
    SMDS_DBConnection_.Open
    SMDS_DB_Status = 0
    
    Set MACS_DBConnection_ = New ADODB.Connection
    MACS_DBConnection_.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" _
                        & "Data Source=" & TIR_DB_PATH & ""
    MACS_DBConnection_.Open
    MACS_DB_Status = 0
     Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If SMDS_DB_Status = 0 Then
        SMDS_DBConnection_.Close
    End If
    Set SMDS_DBConnection_ = Nothing
    
    If MACS_DB_Status = 0 Then
        MACS_DBConnection_.Close
    End If
    
    Set MACS_DBConnection_ = Nothing
     Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub










