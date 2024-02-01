VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryCBInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIRCircuitBreakers_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim SMDS_DBConnection_ As ADODB.Connection
Dim MACS_DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC
Dim TIR_CB_Item_ As clsCircuitBreaker

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

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsCircuitBreaker) As String
    strSQL_ = "SELECT * FROM TIR_CIRCUIT_BREAKERS WHERE ("
    Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False
    
    If SearchConditions.CBClass <> "" Then
        strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[CB Class])='" & SearchConditions.CBClass & "')"
        flgOtherConditionExists = True
    End If
    If SearchConditions.PowerSource <> "" Then
        If Not flgOtherConditionExists Then
            strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[Power Source])='" & SearchConditions.PowerSource & "')"
            flgOtherConditionExists = True
        Else
            strSQL_ = strSQL_ & "AND ((TIR_CIRCUIT_BREAKERS.[Power Source])='" & SearchConditions.PowerSource & "')"
        End If
    End If
    If SearchConditions.System <> "" Then
        If Not flgOtherConditionExists Then
            strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[System])='" & SearchConditions.System & "')"
            flgOtherConditionExists = True
        Else
            strSQL_ = strSQL_ & "AND ((TIR_CIRCUIT_BREAKERS.[System])='" & SearchConditions.System & "')"
        End If
    End If
    If SearchConditions.CBName <> "" Then
        If Not flgOtherConditionExists Then
            strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[CB Name])='" & SearchConditions.CBName & "')"
            flgOtherConditionExists = True
        Else
            strSQL_ = strSQL_ & "AND ((TIR_CIRCUIT_BREAKERS.[CB Name])='" & SearchConditions.CBName & "')"
        End If
    End If
    If SearchConditions.CBPLocation <> "" Then
        If Not flgOtherConditionExists Then
            strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[CBP Location])='" & SearchConditions.CBPLocation & "')"
            flgOtherConditionExists = True
        Else
            strSQL_ = strSQL_ & "AND ((TIR_CIRCUIT_BREAKERS.[CBP Location])='" & SearchConditions.CBPLocation & "')"
        End If
    End If
    If SearchConditions.ConnectedBUS <> "" Then
        If Not flgOtherConditionExists Then
            strSQL_ = strSQL_ & "((TIR_CIRCUIT_BREAKERS.[Connected BUS])='" & SearchConditions.ConnectedBUS & "')"
            flgOtherConditionExists = True
        Else
            strSQL_ = strSQL_ & "AND ((TIR_CIRCUIT_BREAKERS.[Connected BUS])='" & SearchConditions.ConnectedBUS & "')"
        End If
    End If
    
    strSQL_ = strSQL_ & ");"
            
    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function AdvancedSearchCircuitBreaker(ByRef CBSearchConditions As clsCircuitBreaker, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    AdvancedSearchCircuitBreaker = False
    If MACS_DB_Status = 1 Then Exit Function
    
    Set RsTIRCircuitBreakers_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(CBSearchConditions)
    
    RsTIRCircuitBreakers_.Open strSQL_, MACS_DBConnection_, adOpenStatic
    
    Do Until RsTIRCircuitBreakers_.EOF
        Set TIR_CB_Item_ = New clsCircuitBreaker
        
        With TIR_CB_Item_
            .FoundInMACS = True
            .CBNumber = "" & RsTIRCircuitBreakers_.Fields("CB Number")
            .CBName = "" & RsTIRCircuitBreakers_.Fields("CB Name")
            .CBClass = "" & RsTIRCircuitBreakers_.Fields("CB Class")
            .CBPLocation = "" & RsTIRCircuitBreakers_.Fields("CBP Location")
            .Comment = "" & RsTIRCircuitBreakers_.Fields("Comment")
            .ConnectedBUS = "" & RsTIRCircuitBreakers_.Fields("Connected BUS")
            .MACSDocument = "" & RsTIRCircuitBreakers_.Fields("MACS Document")
            .PowerSource = "" & RsTIRCircuitBreakers_.Fields("Power Source")
            .Source = "" & RsTIRCircuitBreakers_.Fields("Source")
            .System = "" & RsTIRCircuitBreakers_.Fields("System")
        End With
        SearchResult.Add TIR_CB_Item_
        Set TIR_CB_Item_ = Nothing
        AdvancedSearchCircuitBreaker = True
        RsTIRCircuitBreakers_.MoveNext
        DoEvents
    Loop
   
    RsTIRCircuitBreakers_.Close
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function



Public Function SearchRefDMsByCBNumber(ByRef CBInfo As clsCircuitBreaker, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMsByCBNumber = False
    If SMDS_DB_Status = 1 Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM CIRCUIT_BREAKERS_REF_DM_LIST WHERE (((CIRCUIT_BREAKERS_REF_DM_LIST.[CB number])='" & CBInfo.CBNumber & "'));"

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
        SearchRefDMsByCBNumber = True
        DoEvents
    Loop

    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Public Function SearchCircuitBreakerInfoFromTIRDB(ByRef CBInfo As clsCircuitBreaker) As Boolean
    On Error GoTo errHandler
    SearchCircuitBreakerInfoFromTIRDB = False
    
    If MACS_DB_Status = 1 Then Exit Function
    
    Set RsTIRCircuitBreakers_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TIR_CIRCUIT_BREAKERS WHERE (((TIR_CIRCUIT_BREAKERS.[CB Number])='" & CBInfo.CBNumber & "'));"
    
    RsTIRCircuitBreakers_.Open strSQL_, MACS_DBConnection_, adOpenDynamic
    
    If RsTIRCircuitBreakers_.EOF = False Then
        With CBInfo
            .FoundInMACS = True
            .CBNumber = "" & RsTIRCircuitBreakers_.Fields("CB Number")
            .CBName = "" & RsTIRCircuitBreakers_.Fields("CB Name")
            .CBClass = "" & RsTIRCircuitBreakers_.Fields("CB Class")
            .CBPLocation = "" & RsTIRCircuitBreakers_.Fields("CBP Location")
            .Comment = "" & RsTIRCircuitBreakers_.Fields("Comment")
            .ConnectedBUS = "" & RsTIRCircuitBreakers_.Fields("Connected BUS")
            .MACSDocument = "" & RsTIRCircuitBreakers_.Fields("MACS Document")
            .PowerSource = "" & RsTIRCircuitBreakers_.Fields("Power Source")
            .Source = "" & RsTIRCircuitBreakers_.Fields("Source")
            .System = "" & RsTIRCircuitBreakers_.Fields("System")
            
            SearchCircuitBreakerInfoFromTIRDB = True
        End With
    End If
    
    RsTIRCircuitBreakers_.Close
    
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








