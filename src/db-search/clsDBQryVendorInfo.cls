VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBQryVendorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RsTIREnterprise_ As ADODB.Recordset
Dim DBConnection_ As ADODB.Connection
Dim TIR_Vendor_Item_ As clsVendorCode
Dim strSQL_ As String

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

Private Function GetSQLStatementForAdvancedSearch(ByRef SearchConditions As clsVendorCode) As String
    strSQL_ = "SELECT * FROM TBL_UPDATED_TIR_ENTERPRISE_ALL WHERE ("
    Dim flgOtherConditionExists     As Boolean
    flgOtherConditionExists = False

    If SearchConditions.VendorCode <> "" Then
        strSQL_ = strSQL_ & "((TBL_UPDATED_TIR_ENTERPRISE_ALL.[Vendor Code]) Like '%" & SearchConditions.VendorCode & "%')) OR (((TBL_UPDATED_TIR_ENTERPRISE_ALL.[Alternate Code]) Like '%" & SearchConditions.VendorCode & "%')) OR " & _
                "(((TBL_UPDATED_TIR_ENTERPRISE_ALL.[Vendor Name]) Like '%" & SearchConditions.VendorName & "%')"
        flgOtherConditionExists = True
    End If

    strSQL_ = strSQL_ & ");"

    GetSQLStatementForAdvancedSearch = strSQL_
End Function

Public Function SearchVendorsByKeyword(ByRef SearchConditions As clsVendorCode, ByRef SearchResult As Collection) As Boolean
    On Error GoTo errHandler
    SearchVendorsByKeyword = False
    If DB_Status = DB_NOT_FOUND Then Exit Function

    Set RsTIREnterprise_ = New ADODB.Recordset
    strSQL_ = GetSQLStatementForAdvancedSearch(SearchConditions)

    RsTIREnterprise_.Open strSQL_, DBConnection_, adOpenStatic

    Do Until RsTIREnterprise_.EOF
        Set TIR_Vendor_Item_ = New clsVendorCode

        With TIR_Vendor_Item_
            .VendorCode = "" & RsTIREnterprise_.Fields("Vendor Code")
            .AlternateCode = "" & RsTIREnterprise_.Fields("Alternate Code")
            .AlternateCodeType = "" & RsTIREnterprise_.Fields("Alternate Code Type")
            .VendorName = "" & RsTIREnterprise_.Fields("Vendor Name")
            .BusinessUnitName = "" & RsTIREnterprise_.Fields("Business Unit Name")
            .City = "" & RsTIREnterprise_.Fields("TIR City")
            .Country = "" & RsTIREnterprise_.Fields("Country")
            .ZIPCode = "" & RsTIREnterprise_.Fields("Zip Code")
            .Street = "" & RsTIREnterprise_.Fields("Street")
            .PhoneNumber = "" & RsTIREnterprise_.Fields("Phone Number")
            .FAX = "" & RsTIREnterprise_.Fields("FAX Number")
            .EMail = "" & RsTIREnterprise_.Fields("E-mail")
            .URL = "" & RsTIREnterprise_.Fields("URL")
            .Comment = "" & RsTIREnterprise_.Fields("Comment")
            .Source = "" & RsTIREnterprise_.Fields("Source")
        End With
        SearchResult.Add TIR_Vendor_Item_
        Set TIR_Vendor_Item_ = Nothing
        SearchVendorsByKeyword = True
        RsTIREnterprise_.MoveNext
        DoEvents
    Loop

    RsTIREnterprise_.Close

    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Public Function SearchVendorInfoFromTIRDB(ByRef VendorInfo As clsVendorCode) As Boolean
    On Error GoTo errHandler
    SearchVendorInfoFromTIRDB = False
    
    If DB_Status = DB_NOT_FOUND Then Exit Function
    
    Set RsTIREnterprise_ = New ADODB.Recordset
    strSQL_ = "SELECT * FROM TBL_UPDATED_TIR_ENTERPRISE_ALL WHERE (((TBL_UPDATED_TIR_ENTERPRISE_ALL.[Vendor Code])='" & VendorInfo.VendorCode & "'));"
    
    RsTIREnterprise_.Open strSQL_, DBConnection_, adOpenStatic
    
    If RsTIREnterprise_.EOF = False Then
        With VendorInfo
            .AlternateCode = "" & RsTIREnterprise_.Fields("Alternate Code")
            .AlternateCodeType = "" & RsTIREnterprise_.Fields("Alternate Code Type")
            .VendorName = "" & RsTIREnterprise_.Fields("Vendor Name")
            .BusinessUnitName = "" & RsTIREnterprise_.Fields("Business Unit Name")
            .City = "" & RsTIREnterprise_.Fields("TIR City")
            .Country = "" & RsTIREnterprise_.Fields("Country")
            .ZIPCode = "" & RsTIREnterprise_.Fields("Zip Code")
            .Street = "" & RsTIREnterprise_.Fields("Street")
            .PhoneNumber = "" & RsTIREnterprise_.Fields("Phone Number")
            .FAX = "" & RsTIREnterprise_.Fields("FAX Number")
            .EMail = "" & RsTIREnterprise_.Fields("E-mail")
            .URL = "" & RsTIREnterprise_.Fields("URL")
            .Comment = "" & RsTIREnterprise_.Fields("Comment")
            .Source = "" & RsTIREnterprise_.Fields("Source")
            
            SearchVendorInfoFromTIRDB = True
        End With
    End If
    
    RsTIREnterprise_.Close
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Class_Initialize()
    Dim DB_PATH As String
    Dim TIRSetting As clsConfigTIREnterprise
    Set TIRSetting = New clsConfigTIREnterprise
    DB_PATH = TIRSetting.TIRDatabasePath
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
End Sub

Private Sub Class_Terminate()
    If DB_Status = Connected Then
        DBConnection_.Close
    End If
    
    Set DBConnection_ = Nothing
End Sub


