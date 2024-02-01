VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryAPInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIRAccessPoints_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC
Dim TIR_Panel_Item_ As clsAccessPanel

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

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsAccessPanel) As String
    strSQL_ = "SELECT * FROM GP_ACCESS_PANELS WHERE ("
    Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False

    If SearchConditions.PanelNumber <> "" Then
        strSQL_ = strSQL_ & "((GP_ACCESS_PANELS.[GPELEMENTID]) Like '" & SearchConditions.PanelNumber & "%')) OR " & _
                "(((GP_ACCESS_PANELS.[NOMEN]) Like '%" & SearchConditions.PanelNumber & "%')"
        flgOtherConditionExists = True
    End If

    strSQL_ = strSQL_ & ");"

    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function SearchAccessPanelsByKeyword(ByRef SearchConditions As clsAccessPanel, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    SearchAccessPanelsByKeyword = False
    If DB_Status = DB_NOT_FOUND Then Exit Function

    Set RsTIRAccessPoints_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(SearchConditions)

    RsTIRAccessPoints_.Open strSQL_, DBConnection_, adOpenStatic

    Do Until RsTIRAccessPoints_.EOF
        Set TIR_Panel_Item_ = New clsAccessPanel

        With TIR_Panel_Item_
            .PanelNumber = "" & RsTIRAccessPoints_.Fields("GPELEMENTID")
            .Description = "" & RsTIRAccessPoints_.Fields("NOMEN")
        End With
        SearchResult.Add TIR_Panel_Item_
        Set TIR_Panel_Item_ = Nothing
        SearchAccessPanelsByKeyword = True
        RsTIRAccessPoints_.MoveNext
        DoEvents
    Loop

    RsTIRAccessPoints_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchRefDMsByPanelNumber(ByRef PanelInfo As clsAccessPanel, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMsByPanelNumber = False
    If DB_Status = DB_NOT_FOUND Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM ACCESSPOINT_REF_DM_LIST_ALL_IC WHERE (((ACCESSPOINT_REF_DM_LIST_ALL_IC.[PANEL NUMBER])='" & PanelInfo.PanelNumber & "'));"
    
    RsRefDMs_.Open strSQL_, DBConnection_, adOpenDynamic
    
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
        SearchRefDMsByPanelNumber = True
        DoEvents
    Loop
    
    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Public Function SearchPanelInfoFromTIRDB(ByRef PanelInfo As clsAccessPanel) As Boolean
    On Error GoTo errHandler
    SearchPanelInfoFromTIRDB = False
    
    If DB_Status = DB_NOT_FOUND Then Exit Function
    
    Set RsTIRAccessPoints_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM GP_ACCESS_PANELS WHERE (((GP_ACCESS_PANELS.[GPELEMENTID])='" & PanelInfo.PanelNumber & "'));"
    
    RsTIRAccessPoints_.Open strSQL_, DBConnection_, adOpenStatic
    
    If RsTIRAccessPoints_.EOF = False Then
        With PanelInfo
            .FoundInSMDS = True
            .PanelNumber = "" & RsTIRAccessPoints_.Fields("GPELEMENTID")
            .Description = "" & RsTIRAccessPoints_.Fields("NOMEN")
            
            SearchPanelInfoFromTIRDB = True
        End With
    End If
    
    RsTIRAccessPoints_.Close
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Dim DB_PATH As String
    Dim TIRSetting As clsConfigTIRAccessPoints
    Set TIRSetting = New clsConfigTIRAccessPoints
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






