Attribute VB_Name = "mdlESRD"
'****************************************************************************************
'
'    ESRDで定義されているの各要素を列挙型として定義するためのモジュール
'
'    オートコンプリートを使用して
'    主に、プログラムの記述を楽にし、可読性を高めるために用いる
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
'#Const DEBUG_MODE = 1

Public Enum EngSrcMetadataColumn
    FileCategory = 1
    filename = 2
    Comments = 3
    FileIssue = 4
    FileTitle = 5
    FileFormat = 6
    EngineeringSourceCategory = 7
    ResponsibleDepartment = 8
    AircraftModel = 9
    VridgeRReference = 10
    Zone = 11
    AccessPoint = 12
    PartNumber = 13
    ExportControl = 14
    Active = 15
    ChangeNumber = 16
    OriginalEngineeringSourceID = 17
End Enum

Public Enum ConvertedDMColumn
    [File Category] = 1
    [File Name] = 2
    [File Issue] = 3
    [File Title] = 4
    [File Format] = 5
    [Included In Manual] = 6
    [Aircraft Model] = 7
    [Export Control] = 8
    Comments = 9
    Active = 10
    [Change Number] = 11
End Enum

Public Enum EnterpriseTIRMetadataColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    VendorCode = 4
    AlternateCode = 5
    AlternateCodeType = 6
    VendorName = 7
    BusinessUnitName = 8
    City = 9
    Country = 10
    ZIPCode = 11
    Street = 12
    PhoneNumber = 13
    FAX = 14
    EMail = 15
    URL = 16
    Comments = 17
    Source = 18
End Enum

Public Enum SuppliesTIRColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    SupplyNumber = 4
    SupplyNumberType = 5
    SupplyName = 6
    ManufacturerCode = 7
    ShortName = 8
    LocallySuppliedFlag = 9
    Comment = 10
    Source = 11
End Enum

Public Enum ToolsTIRColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    ToolNumber = 4
    ToolName = 5
    ManufacturerCode = 6
    ShortName = 7
    AlternateToolNumber = 8
    AlternateToolDescription = 9
    OverLengthPartNumber = 10
    ProcurementData = 11
    Remarks = 12
    Comment = 13
    Source = 14
End Enum


Public Enum CircuitBreakersTIRColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    CBNumber = 4
    CBName = 5
    CBClass = 6
    Comment = 7
    Source = 8
End Enum

'Public Enum CircuitBreakersTIRColumn
'    TIRType = 1
'    Status = 2
'    ItemIdentifier = 3
'    ZoneNumber = 4
'    ZoneDescription = 5
'    Applicability = 6
'    Comment = 7
'    Source = 8
'End Enum

Public Enum AccessPointsTIRColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    AccessPointNumber = 4
    ZoneRef = 5
    AccessPointName = 6
    AccessPointType = 7
    Applicability = 8
    Comment = 9
    Source = 10
End Enum

Public Enum ZonesTIRColumn
    TIRType = 1
    Status = 2
    ItemIdentifier = 3
    ZoneNumber = 4
    zonedescription = 5
    Applicability = 6
    Comment = 7
    Source = 8
End Enum

Public Enum IPCSpareIntegrationColumn
    [Type_] = 1
    Status = 2
    [Part Nbr] = 3
    [name] = 4
    SCD = 5
    OPN = 6
    [Vendor Code] = 7
    [Internal Notes] = 8
    Comment = 9
    Source = 10
End Enum


Public Enum MetadataStatus
    Edit = 0
    [New] = 1
    Updated = 2
    Official = 3
    Obsolete = 4
    Dbg = 5
End Enum



Public Enum ESRDErrorCode
    NoError = 0
    FileCategoryError = 1
    
End Enum

Public Enum ESRDFileCategory
    'Engineering Source Document
    Author = 1
    'Illustration
    Illustration = 2
    'Technical Draft for Operationg Manuals
    ConvertedDM = 3
    'TIR Integration Files
    SUPPLIES = 4
    Tools = 5
    Enterprise = 6
    CircuitBreakers = 7
    Zones = 8
    AccessPoints = 9
    'IPC-Spare Integration File
    IPCSpare = 10
    'Wiring integration Files
    EquipmentList = 11
    WireList = 12
    PlugAndReceptacleList = 13
    TerminalList = 14
    SpliceList = 15
    EarthPointList = 16
    
    Errorlog = 17
    TTStatusUpdateLog = 18
End Enum

Public Enum ESRDRequirementForElement
    Optionaltext = 0
    ConditionalText = 1
    MandatoryText = 2
    ValidVendorCode = 3
End Enum




'
'
'Type tpIRPackage
'    IRNo As String
'    ICN As String
'    Title As String
'    SheetNo As String
'    TmpMetaFilename As String
'    TmpMetaFolder As String
'    SendingMetaFilename As String
'    SendingMetaFolder As String
'End Type


Public Const cESRD_MetadataSeparator = "$"
Public Const cESRD_EOF = "EOF"



Public Function GetFileCategoryName(FileCategoryIndex As ESRDFileCategory) As String
    Dim wRet As String
    Select Case FileCategoryIndex
    Case Author: wRet = "Author"
    Case Illustration: wRet = "Illustration"
    Case ConvertedDM: wRet = "ConvertedDM"
    Case SUPPLIES: wRet = "Supplies"
    Case Tools: wRet = "Tools"
    Case Enterprise: wRet = "Enterprise"
    Case CircuitBreakers: wRet = "Circuit breakers"
    Case Zones: wRet = "Zones"
    Case AccessPoints: wRet = "Access-points"
    Case IPCSpare: wRet = "Part"
    Case EquipmentList: wRet = "EquipmentList"
    Case WireList: wRet = "WireList"
    Case PlugAndReceptacleList: wRet = "Plug&ReceptacleList"
    Case TerminalList: wRet = "TerminalList"
    Case SpliceList: wRet = "SpliceList"
    Case EarthPointList: wRet = "EarthPointList"
    Case Else: wRet = "unknown"
    End Select
    
    GetFileCategoryName = wRet
End Function




#If (DEBUG_MODE = 1) Then

'Public Enum IntegrationType
'    IPCSpare
'End Enum
'
'Public Enum TIRType
'    Supplies = 1
'    Tools = 2
'    Enterprise = 3
'    CircuitBreakers = 4
'    Zones = 5
'    AccessPoints = 6
'End Enum

#End If
