VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryZoneInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIRZones_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC
Dim TIR_Zone_Item_  As clsZone

Private Enum ErrType
    Connected = 0
    DB_NOT_FOUND = 1
End Enum
 
Dim DB_Status As ErrType

Property Get DatabaseNotFound() As Boolean
    DatabaseNotFound = False
    If DB_Status = DB_NOT_FOUND Then
        DatabaseNotFound = True
    End If
End Property

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsZone) As String
    strSQL_ = "SELECT * FROM GP_ZONES WHERE ("
    Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False

    If SearchConditions.ZoneNumber <> "" Then
        strSQL_ = strSQL_ & "((GP_ZONES.[GPELEMENTID]) Like '" & SearchConditions.ZoneNumber & "%')) OR " & _
                "(((GP_ZONES.[NOMEN]) Like '%" & SearchConditions.ZoneNumber & "%')"
        flgOtherConditionExists = True
    End If

    strSQL_ = strSQL_ & ");"

    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function SearchToolsByKeyword(ByRef SearchConditions As clsZone, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    SearchToolsByKeyword = False
    If DB_Status = DB_NOT_FOUND Then Exit Function

    Set RsTIRZones_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(SearchConditions)

    RsTIRZones_.Open strSQL_, DBConnection_, adOpenStatic

    Do Until RsTIRZones_.EOF
        Set TIR_Zone_Item_ = New clsZone

        With TIR_Zone_Item_
            .ZoneNumber = "" & RsTIRZones_.Fields("GPELEMENTID")
            .Description = "" & RsTIRZones_.Fields("NOMEN")
        End With
        SearchResult.Add TIR_Zone_Item_
        Set TIR_Zone_Item_ = Nothing
        SearchToolsByKeyword = True
        RsTIRZones_.MoveNext
        DoEvents
    Loop

    RsTIRZones_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchRefDMsByZoneNumber(ByRef ZoneInfo As clsZone, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMsByZoneNumber = False
    If DB_Status = DB_NOT_FOUND Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM ZONE_REF_DM_LIST_ALL_IC WHERE (((ZONE_REF_DM_LIST_ALL_IC.[ZONE])='" & ZoneInfo.ZoneNumber & "'));"
    
    RsRefDMs_.Open strSQL_, DBConnection_, adOpenStatic
    
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
        SearchRefDMsByZoneNumber = True
        DoEvents
    Loop
    
    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Public Function SearchZoneInfoFromTIRDB(ByRef ZoneInfo As clsZone) As Boolean
    On Error GoTo errHandler
    SearchZoneInfoFromTIRDB = False
    
    If DB_Status = DB_NOT_FOUND Then Exit Function
    
    Set RsTIRZones_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM GP_ZONES WHERE (((GP_ZONES.[GPELEMENTID])='" & ZoneInfo.ZoneNumber & "'));"
    
    RsTIRZones_.Open strSQL_, DBConnection_, adOpenDynamic
    
    If RsTIRZones_.EOF = False Then
        With ZoneInfo
            .FoundInSMDS = True
            .ZoneNumber = "" & RsTIRZones_.Fields("GPELEMENTID")
            .Description = "" & RsTIRZones_.Fields("NOMEN")
            
            SearchZoneInfoFromTIRDB = True
        End With
    End If
    
    RsTIRZones_.Close
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Dim DB_PATH As String
    Dim TIRSetting As clsConfigTIRZones
    Set TIRSetting = New clsConfigTIRZones
    DB_PATH = TIRSetting.SMDSTIRDatabasePath
    Set TIRSetting = Nothing
    Dim fso As New FileSystemObject
    If Not fso.FileExists(DB_PATH) Then
        DB_Status = DB_NOT_FOUND
        Exit Sub
    End If
    
    Set DBConnection_ = New ADODB.Connection
    DBConnection_.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" _
                        & "Data Source=" & DB_PATH & ""
    DBConnection_.Open
    DB_Status = Connected
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub Class_Terminate()
    If DB_Status = Connected Then
        DBConnection_.Close
    End If
    
    Set DBConnection_ = Nothing
End Sub




