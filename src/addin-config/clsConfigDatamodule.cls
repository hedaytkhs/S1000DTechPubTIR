VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigDatamodule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'11 óÒñ⁄Ç…DMåüçıópê›íËÇï€ë∂
Private Const XlsColDM As Long = 11

Private Enum XlsCellTTBaseFolderPath
    Row = 1
    col = XlsColDM
End Enum

Dim TTBaseFolderPath_ As String
Dim ShtConfig As Worksheet

Property Let TTBaseFolderPath(ByVal sTTBaseFolder As String)
    TTBaseFolderPath_ = sTTBaseFolder
    ShtConfig.Cells(XlsCellTTBaseFolderPath.Row, XlsCellTTBaseFolderPath.col).Value = TTBaseFolderPath_
End Property

Property Get TTBaseFolderPath() As String
    TTBaseFolderPath = ShtConfig.Cells(XlsCellTTBaseFolderPath.Row, XlsCellTTBaseFolderPath.col).Value
End Property

Function Save() As Boolean
    ThisWorkbook.Save
End Function

Function TTBaseFolderExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    TTBaseFolderExists = fso.FolderExists(TTBaseFolderPath_)
    Set fso = Nothing
End Function

Private Sub Class_Initialize()
    Set ShtConfig = ThisWorkbook.Worksheets(CNST_XLS_SHEET_CONFIG)
    With ShtConfig
        TTBaseFolderPath_ = .Cells(XlsCellTTBaseFolderPath.Row, XlsCellTTBaseFolderPath.col).Value
    End With
End Sub

Private Sub Class_Terminate()
    Set ShtConfig = Nothing
End Sub



