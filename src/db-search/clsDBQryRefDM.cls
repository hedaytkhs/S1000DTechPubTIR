VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryRefDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsSubjectDMC_ As ADODB.Recordset
Dim RsRefDMs_ As ADODB.Recordset
Dim DBConnection_ As ADODB.Connection
Dim strSQL_ As String
Dim RefDMC_ As clsDMC

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

Public Function SearchRefDMs(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchRefDMs = False
    If DB_Status = DB_NOT_FOUND Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM DMC_REF_DM_LIST WHERE (((DMC_REF_DM_LIST.[DMC(HOST)])='" & DMInfo.DMC & "'));"
    
    RsRefDMs_.Open strSQL_, DBConnection_, adOpenStatic
    
    Do Until RsRefDMs_.EOF
        Set RefDMC_ = New clsDMC
        With RefDMC_
            .DMC = "" & RsRefDMs_.Fields("DMC(TARGET)")
            .TechName = "" & RsRefDMs_.Fields("TECHNAME(TARGET)")
            .InfoName = "" & RsRefDMs_.Fields("INFONAME(TARGET)")
        End With
        RefDMs.Add RefDMC_
        Set RefDMC_ = Nothing
        RsRefDMs_.MoveNext
        SearchRefDMs = True
        DoEvents
    Loop
    
    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchReferencingDMs(ByRef DMInfo As clsDMC, ByRef RefDMs As Collection) As Boolean
    On Error GoTo errHandler
    SearchReferencingDMs = False
    If DB_Status = DB_NOT_FOUND Then Exit Function
    Set RsRefDMs_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM DMC_REF_DM_LIST WHERE (((DMC_REF_DM_LIST.[DMC(TARGET)])='" & DMInfo.DMC & "'));"
    
    RsRefDMs_.Open strSQL_, DBConnection_, adOpenStatic
    
    Do Until RsRefDMs_.EOF
        Set RefDMC_ = New clsDMC
        With RefDMC_
            .DMC = "" & RsRefDMs_.Fields("DMC(HOST)")
            .TechName = "" & RsRefDMs_.Fields("TECHNAME(HOST)")
            .InfoName = "" & RsRefDMs_.Fields("INFONAME(HOST)")
        End With
        RefDMs.Add RefDMC_
        Set RefDMC_ = Nothing
        RsRefDMs_.MoveNext
        SearchReferencingDMs = True
        DoEvents
    Loop
    
    RsRefDMs_.Close
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function


Public Function SearchDMInfoOfSubjectDMC(ByRef DMInfo As clsDMC) As Boolean
    On Error GoTo errHandler
    SearchDMInfoOfSubjectDMC = False
    
    If DB_Status = DB_NOT_FOUND Then Exit Function
    
    Set RsSubjectDMC_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM AEC_DMOD WHERE (((AEC_DMOD.[DMC])='" & DMInfo.DMC & "'));"
    
    RsSubjectDMC_.Open strSQL_, DBConnection_, adOpenDynamic
    
    If RsSubjectDMC_.EOF = False Then
        With DMInfo
            .TechName = "" & RsSubjectDMC_.Fields("TECHNAME")
            .InfoName = "" & RsSubjectDMC_.Fields("INFONAME")
            
            SearchDMInfoOfSubjectDMC = True
        End With
    End If
    
    RsSubjectDMC_.Close
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Dim DB_PATH As String
    Dim DBSetting As clsConfigTIRZones
    Set DBSetting = New clsConfigTIRZones
    DB_PATH = DBSetting.SMDSTIRDatabasePath
    Set DBSetting = Nothing
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






