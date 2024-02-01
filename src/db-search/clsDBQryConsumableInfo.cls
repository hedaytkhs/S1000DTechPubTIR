VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryConsumableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIRSupplies_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim RsRefRevisionHistory_ As ADODB.Recordset

Dim SMDS_DBConnection_ As ADODB.Connection
Dim MACS_DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC
Dim ItemRevision_ As clsConsumable
Dim TIR_Supply_Item_ As clsConsumable

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

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsConsumable) As String
    strSQL_ = "SELECT * FROM TIR_SUPPLIES WHERE ("

Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False

    If SearchConditions.SupplyName <> "" And SearchConditions.SupplyNumber <> "" Then
        strSQL_ = strSQL_ & "((TIR_SUPPLIES.[Supply Number]) Like '%" & SearchConditions.SupplyNumber & "%') AND ((TIR_SUPPLIES.[ActiveItem])=YES)) OR (((TIR_SUPPLIES.[Supply Name]) Like '%" & SearchConditions.SupplyName & "%') AND ((TIR_SUPPLIES.[ActiveItem])=YES)"
        flgOtherConditionExists = True
    End If

    strSQL_ = strSQL_ & ");"

    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function SearchConsumablesByKeyword(ByRef SearchConditions As clsConsumable, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    SearchConsumablesByKeyword = False
    If MACS_DB_Status = 1 Then Exit Function

    Set RsTIRSupplies_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(SearchConditions)

    RsTIRSupplies_.Open strSQL_, MACS_DBConnection_, adOpenStatic
    Do Until RsTIRSupplies_.EOF
        Set TIR_Supply_Item_ = New clsConsumable

        With TIR_Supply_Item_
            .FoundInMACS = True
            .SupplyName = "" & RsTIRSupplies_.Fields("Supply Name")
            .SupplyNumber = "" & RsTIRSupplies_.Fields("Supply Number")
            .SupplyNumberType = "" & RsTIRSupplies_.Fields("Supply Number Type")
            .SupplyLongName = "" & RsTIRSupplies_.Fields("Long Name保管用")
            .LocallySuppliedFlag = "" & RsTIRSupplies_.Fields("Locally Supplied Flag")
            .ManufactureCode = "" & RsTIRSupplies_.Fields("Manufacturer Code")
            .Comment = "" & RsTIRSupplies_.Fields("Comment")
            .Remarks = "" & RsTIRSupplies_.Fields("備考")
            .Source = "" & RsTIRSupplies_.Fields("Source")
        End With
        SearchResult.Add TIR_Supply_Item_
        Set TIR_Supply_Item_ = Nothing
        SearchConsumablesByKeyword = True
        RsTIRSupplies_.MoveNext
        DoEvents
    Loop

    RsTIRSupplies_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchRefDMsBySupplyNumber(ByRef SupplyInfo As clsConsumable, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMsBySupplyNumber = False
    If SMDS_DB_Status = 1 Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM SUPPLIES_REF_DM_LIST WHERE (((SUPPLIES_REF_DM_LIST.[Supply Number])='" & SupplyInfo.SupplyNumber & "'));"

    RsRefDMs_.Open strSQL_, SMDS_DBConnection_, adOpenStatic
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
        SearchRefDMsBySupplyNumber = True
        DoEvents
    Loop

    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function GetRevisionHistoryBySupplyNumber(ByRef SupplyInfo As clsConsumable, ByRef RevisionHistory As Collection) As Boolean
    On Error GoTo errHandler
    GetRevisionHistoryBySupplyNumber = False
    If SMDS_DB_Status = 1 Then Exit Function
    Set RsRefRevisionHistory_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TIR_SUPPLIES WHERE (((TIR_SUPPLIES.[Supply Number])='" & SupplyInfo.SupplyNumber & "')) ORDER BY TIR_SUPPLIES.[Revision SEQ];"
'    strSQL_ = "SELECT * FROM TIR_SUPPLIES WHERE (((TIR_SUPPLIES.[Supply Number])='" & SupplyInfo.SupplyNumber & "'));"
'    strSQL_ = "SELECT * FROM SUPPLIES_REF_DM_LIST WHERE (((SUPPLIES_REF_DM_LIST.[Supply Number])='" & SupplyInfo.SupplyNumber & "'));"
    RsRefRevisionHistory_.Open strSQL_, SMDS_DBConnection_, adOpenStatic

    Do Until RsRefRevisionHistory_.EOF
        Set ItemRevision_ = New clsConsumable
        With ItemRevision_
            .SupplyNumber = "" & RsRefRevisionHistory_.Fields("Supply Number")
            .SupplyName = "" & RsRefRevisionHistory_.Fields("Supply Name")
            .SupplyLongName = "" & RsRefRevisionHistory_.Fields("Long Name保管用")
            .ManufactureCode = "" & RsRefRevisionHistory_.Fields("Manufacturer Code")
            .ActiveItem = "" & RsRefRevisionHistory_.Fields("ActiveItem")
            .RevisionSequence = "" & RsRefRevisionHistory_.Fields("Revision SEQ")
            .ShortName = "" & RsRefRevisionHistory_.Fields("Short Name")
        End With
        RevisionHistory.Add ItemRevision_
        Set ItemRevision_ = Nothing
        RsRefRevisionHistory_.MoveNext
        GetRevisionHistoryBySupplyNumber = True
        DoEvents
    Loop

    RsRefRevisionHistory_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchSupplyInfoFromTIRDB(ByRef SupplyInfo As clsConsumable) As Boolean
    On Error GoTo errHandler
    SearchSupplyInfoFromTIRDB = False

    If MACS_DB_Status = 1 Then Exit Function

    Set RsTIRSupplies_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TIR_SUPPLIES WHERE (((TIR_SUPPLIES.[Supply Number])='" & SupplyInfo.SupplyNumber & "') AND ((TIR_SUPPLIES.[ActiveItem])=YES));"

    RsTIRSupplies_.Open strSQL_, MACS_DBConnection_, adOpenDynamic

    If RsTIRSupplies_.EOF = False Then
        With SupplyInfo
            .FoundInMACS = True
            .SupplyName = "" & RsTIRSupplies_.Fields("Supply Name")
            .SupplyNumber = "" & RsTIRSupplies_.Fields("Supply Number")
            .SupplyNumberType = "" & RsTIRSupplies_.Fields("Supply Number Type")
            .SupplyLongName = "" & RsTIRSupplies_.Fields("Long Name保管用")
            .LocallySuppliedFlag = "" & RsTIRSupplies_.Fields("Locally Supplied Flag")
            .ManufactureCode = "" & RsTIRSupplies_.Fields("Manufacturer Code")
            .Comment = "" & RsTIRSupplies_.Fields("Comment")
            .Remarks = "" & RsTIRSupplies_.Fields("備考")
            .Source = "" & RsTIRSupplies_.Fields("Source")
        End With
        SearchSupplyInfoFromTIRDB = True
    End If

    RsTIRSupplies_.Close

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












