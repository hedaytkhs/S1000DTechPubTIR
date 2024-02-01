VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetadataFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim FileCategoryIndex_ As ESRDFileCategory
Dim FileBaseName_ As String
Dim MetadataFilename_ As String
Dim FileCategoryName_ As String
Dim MetadateFileDate_ As String * 8
Dim DMC_ As String * 25

'FileCategory
Property Let FileCategory(ByVal IdxFileCategory As ESRDFileCategory)
    FileCategoryIndex_ = IdxFileCategory
    FileCategoryName_ = GetFileCategoryName(IdxFileCategory)
    Call GetFileBaseName
    Call GetMetadataFileName
End Property

Property Let DMC(ByVal sDMC As String)
    DMC_ = sDMC
    Call GetFileBaseName
    Call GetMetadataFileName
End Property

Property Get DMC() As String
    DMC = DMC_
End Property

'MetadataFileName取得
Property Get filename() As String
    filename = MetadataFilename_
End Property

'FileBaseName取得
Property Get BaseName() As String
    BaseName = FileBaseName_
End Property


'FileCategoryLabel取得
Property Get FileCategoryName() As String
    FileCategoryName = FileCategoryName_
End Property

'FileDate取得
Property Get FileDate() As String
    FileDate = MetadateFileDate_
End Property


'Header取得

'Metadataの行を取得
'Metadataの行を追加


'EOFを追加

'Metadataを保存
Private Sub GetMetadataFileName()
    Dim wRet As String
    Dim sDate As String
    Dim sTime As String
    sDate = Replace(FORMAT(Now, "Short Date"), "/", "")
    sTime = Replace(FORMAT(Now, "Short Time"), ":", "")
    MetadateFileDate_ = sDate
'    Select Case FileCategoryIndex_
'    Case EquipmentList, WireList, PlugAndReceptacleList, TerminalList, SpliceList, EarthPointList: wRet = FileBaseName_ & "_" & DMC_ & sDate & "_" & sTime & ".csv"
'    Case Else: wRet = FileBaseName_ & sDate & "_" & sTime & ".csv"
'    End Select
    wRet = FileBaseName_ & "_" & sDate & "_" & sTime & ".csv"
    
    MetadataFilename_ = wRet
End Sub

Private Sub GetFileBaseName()
    Const cMetadata = "_Metadata"
    Const cIntegration = "Integration"
    Const cDB = "_DB"
    Const cTIR = "_TIR"
    Const cXML = "_XML"
    
    Dim wRet As String
    Select Case FileCategoryIndex_
    Case Author: wRet = "Author" & cMetadata
    Case Illustration: wRet = cIntegration & cMetadata & cDB & "_Illustration"
    Case ConvertedDM: wRet = "ConvertedDM" & cMetadata
    Case SUPPLIES: wRet = cIntegration & cDB & cTIR & "_Supplies"
    Case Tools: wRet = cIntegration & cDB & cTIR & "_Tools"
    Case Enterprise: wRet = cIntegration & cDB & cTIR & "_Enterprise"
    Case CircuitBreakers: wRet = cIntegration & cDB & cTIR & "_Circuit_Breakers"
    Case Zones: wRet = cIntegration & cDB & cTIR & "_Zones"
    Case AccessPoints: wRet = cIntegration & cDB & cTIR & "_Access_Points"
    Case IPCSpare: wRet = cIntegration & cDB & "_IPC-Spares"
    Case EquipmentList: wRet = cIntegration & cXML & "_Wiring_EL" & "_" & DMC_
    Case WireList: wRet = cIntegration & cXML & "_Wiring_WL" & "_" & DMC_
    Case PlugAndReceptacleList: wRet = cIntegration & cXML & "_Wiring_HL-PR" & "_" & DMC_
    Case TerminalList: wRet = cIntegration & cXML & "_Wiring_HL-T" & "_" & DMC_
    Case SpliceList: wRet = cIntegration & cXML & "_Wiring_HL-S" & "_" & DMC_
    Case EarthPointList: wRet = cIntegration & cXML & "_Wiring_HL-EP" & "_" & DMC_
    Case Errorlog: wRet = "ErrorLog_TextCheck"
    Case TTStatusUpdateLog: wRet = "_作業管理ツール実行結果_"
    Case Else: wRet = "unknown"
    End Select
    
    FileBaseName_ = wRet
End Sub

' mdlESRD内に移動
'Private Sub GetFileCategoryName()
'    Dim wRet As String
'    Select Case FileCategoryIndex_
'    Case Author: wRet = "Author"
'    Case Illustration: wRet = "Illustration"
'    Case ConvertedDM: wRet = "ConvertedDM"
'    Case Supplies: wRet = "Supplies"
'    Case Tools: wRet = "Tools"
'    Case Enterprise: wRet = "Enterprise"
'    Case CircuitBreakers: wRet = "Circuit breakers"
'    Case Zones: wRet = "Zones"
'    Case AccessPoints: wRet = "Access-points"
'    Case IPCSpare: wRet = "Part"
'    Case EquipmentList: wRet = "EquipmentList"
'    Case WireList: wRet = "WireList"
'    Case PlugAndReceptacleList: wRet = "Plug&ReceptacleList"
'    Case TerminalList: wRet = "TerminalList"
'    Case SpliceList: wRet = "SpliceList"
'    Case EarthPointList: wRet = "EarthPointList"
'    Case Else: wRet = "unknown"
'    End Select
'
'    FileCategoryName_ = wRet
'End Sub

Private Sub Class_Initialize()
    Dim mDMC As New clsDMC
    DMC_ = mDMC.DMC
    Set mDMC = Nothing
End Sub
