Attribute VB_Name = "mdlMRJBRDefinition"
'=========================================================
'
'    を操作するためのクラス
'
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'
' 改訂履歴
'
'
' Created by
' Hideaki Takahashi
' 2014/04/04
'
'  Revision History
' 2014/04/05
' MRJ BR のChapter/Paragraphの種類を定義
'
'=========================================================
Option Explicit

Public Enum MRJ_BR_LevelType
    Chapter = 1
    Paragraph = 2
End Enum

Public Enum MRJ_BR_MAX_LEVEL
    MRJ_BR_MAX_LEVEL_CHAPTER = 6
    MRJ_BR_MAX_LEVEL_PARAGRAPH = 6
End Enum

Public Enum MRJ_BR_LanguageCode
    SimplifiedTechnicalEnglish = 1
    English = 2
End Enum

Public Enum MRJ_BR_ContryCode
    UnitedStates = 1
    Japan = 2
End Enum

Public Enum MRJ_BR_ModelIdentCode
    'MRJBRに記載されている
    MRJ = 1
    MBR = 2
    MWG = 3
    ' Maxsaに存在するがMRJ BRには未記載
    PW1000G = 11
    S1000D = 10
    OTHER = 99
End Enum

'MACS-G09-068Kに記載されている名称に従って定義
Public Enum MRJ_BR_RPCCode
    'MRJBRに記載されている名称をプログラム内でも使用するため列挙型を定義しておく
    MITSUBISHI_AIRCRAFT_CORPORATION = 1
    AEROSPACE_INDUSTRIAL_DEVELOPMENT = 2
    Eurocopter_Deutschland_GmbH = 3
    GOODRICH_LIGHTING_SYSTEMS_GMBH = 4
    HAMILTON_SUNDSTRAND_CORP = 5
    
    INTERTECHNIQUE = 6
    HEATH_TECNA = 7
    KORRY_ELECTRONICS_CO = 8
    LMI_AEROSPACE_INC = 9
    GOODRICH_SENSOR_SYSTEMS = 10
    
    NABTESCO_CORP = 11
    PRATT_AND_WHITNEY = 12
    PARKER_HANNIFIN_CORP = 13
    ROCKWELL_COLLINS_INC = 14
    SPIRIT = 15
    
    SUMITOMO_PRECISION_PRODUCTS_CO_LTD = 16
    SAINT_GOBAIN_PERFORMANCE_PLASTICS = 17
    OTHER = 99
End Enum

'MACS-G09-068Kに記載されている名称に従って定義
Public Enum MRJ_BR_OriginatorCode
    'MRJBRに記載されている名称をプログラム内でも使用するため列挙型を定義しておく
    MONOGRAM_SYSTEMS_INC = 1
    AVTECH_CORP = 2
    SPIRIT = 3
    ROCKWELL_COLLINS_INC_DIV_COMMERCIAL_SYSTEMS = 4
    ROHR_INC_DBA_GOODRICH_AEROSTRUCTURES_GROUP_DIV_GOODRICH_AEROSTRUCTURES = 5
    
    PPG_INDUSTRIES_INC_DIV_AEROSPACE = 6
    AMPHENOL_PCD_INC = 7
    GOODRICH_SENSOR_SYSTEMS = 8
    ROSEMOUNT_AEROSPACE_INC_DBA_GOODRICH_SENSOR_SYSTEMS_DIV_GOODRICH_SENSOR_SYSTEMS = 9
    ROSEMOUNT_AEROSPACE_INC_DBA_ROSEMOUNT_AEROSPACE_DIVISION_DIV_GOODRICH_SENSOR_SYSTEMS = 10
    
    HONEYWELL_INC_COMMERCIAL_AVIATION_SYSTEMS = 11
    HAMILTON_SUNDSTRAND_CORPORATION_01 = 12
    PRATT_AND_WHITNEY = 13
    KORRY_ELECTRONICS_CO = 14
    GKN_AEROSPACE_TRANSPARENCY_SYSTEMS_INC = 15
    
    SAINT_GOBAIN_PERFORMANCE_PLASTICS = 16
    PARKER_HANNIFIN_CORPORATION = 17
    LMI_AEROSPACE_INC = 18
    TELEDYNE_CONTROLS = 19
    HAMILTON_SUNDSTRAND_CORPORATION_02 = 20
    
    Eurocopter_Deutschland_GmbH = 21
    MTU_AERO_ENGINES_GMBH = 22
    GOODRICH_LIGHTING_SYSTEMS_GMBH = 23
    DELTA_INC = 24
    ESPA = 25
    
    ZODIAC_INTERTECHNIQUE_FUEL_AND_INERTING_SYSTEMS = 26
    ZODIAC_INTERTECHNIQUE_SYSTEMS_MONITORING_AND_MANAGEMENT = 27
    FALGAYRAS = 28
    DAHER_AEROSPACE_SITE_DE_LUCEAU = 29
    DAHER_AEROSPAC = 30
    
    SENIOR_AEROSPACE_BWT = 31
    ULTRA_ELECTRONICS_LTD = 32
    NABTESCO_CORP = 33
    KOITO_MANUFACTURING_CO_LTD = 34
    AEROSPACE_INDUSTRIAL_DEVELOPMENT = 35
    
    SUMITOMO_PRECISION_PRODUCTS_CO_LTD = 36
    MITSUBISHI_AIRCRAFT_CORPORATION = 37
End Enum


Public Function GetOriginatorCode(ORC_Index As MRJ_BR_OriginatorCode) As String
    Dim wRet As String
    Select Case ORC_Index
'    Case MONOGRAM_SYSTEMS_INC = "29780"
'    Case AVTECH_CORP = "30242"
''    Case SPIRIT = "4ATM5"
'    Case ROCKWELL_COLLINS_INC_DIV_COMMERCIAL_SYSTEMS = "4V792"
'    Case ROHR_INC_DBA_GOODRICH_AEROSTRUCTURES_GROUP_DIV_GOODRICH_AEROSTRUCTURES = "51563"
'
'    Case PPG_INDUSTRIES_INC_DIV_AEROSPACE = "53117"
'    Case AMPHENOL_PCD_INC = "58982"
''    Case GOODRICH_SENSOR_SYSTEMS = "59885"
'    Case ROSEMOUNT_AEROSPACE_INC_DBA_GOODRICH_SENSOR_SYSTEMS_DIV_GOODRICH_SENSOR_SYSTEMS = "59885"
'    Case ROSEMOUNT_AEROSPACE_INC_DBA_ROSEMOUNT_AEROSPACE_DIVISION_DIV_GOODRICH_SENSOR_SYSTEMS = "60678"
'
'    Case HONEYWELL_INC_COMMERCIAL_AVIATION_SYSTEMS = "65507"
'    Case HAMILTON_SUNDSTRAND_CORPORATION_01 = "73030"
''    Case PRATT_AND_WHITNEY = "77445"
''    Case KORRY_ELECTRONICS_CO = "81590"
'    Case GKN_AEROSPACE_TRANSPARENCY_SYSTEM0S_INC = "86175"
'
'    Case SAINT_GOBAIN_PERFORMANCE_PLASTICS = "86228"
'    Case PARKER_HANNIFIN_CORPORATION = "93835"
'    Case LMI_AEROSPACE_INC = "98465"
'    Case TELEDYNE_CONTROLS = "98571"
'    Case HAMILTON_SUNDSTRAND_CORPORATION_02 = "99167"
'
'    Case Eurocopter_Deutschland_GmbH = "C0417"
'    Case MTU_AERO_ENGINES_GMBH = "D3009"
'    Case GOODRICH_LIGHTING_SYSTEMS_GMBH = "D8095"
'    Case DELTA_INC = "DLT01"
'    Case ESPA = "F0215"
'
'    Case ZODIAC_INTERTECHNIQUE_FUEL_AND_INERTING_SYSTEMS = "F0422"
'    Case ZODIAC_INTERTECHNIQUE_SYSTEMS_MONITORING_AND_MANAGEMENT = "F0553"
'    Case FALGAYRAS = "F6914"
'    Case DAHER_AEROSPACE_SITE_DE_LUCEAU = "FA7W9"
'    Case DAHER_AEROSPAC = "FAMW4"
'
'    Case SENIOR_AEROSPACE_BWT = "K2962"
'    Case ULTRA_ELECTRONICS_LTD = "K8081"
'    Case NABTESCO_CORP = "S4980"
'    Case KOITO_MANUFACTURING_CO_LTD = "S5006"
'    Case AEROSPACE_INDUSTRIAL_DEVELOPMENT = "S7549"
'
'    Case SUMITOMO_PRECISION_PRODUCTS_CO_LTD = "SG215"
'    Case MITSUBISHI_AIRCRAFT_CORPORATION = "SJZ51"
'    Case Else: wRet = "unknown"
    End Select
    GetOriginatorCode = wRet
End Function




Public Function GetRPCCode(RPC_Index As MRJ_BR_RPCCode) As String
    Dim wRet As String
    Select Case RPC_Index
'    Case MITSUBISHI_AIRCRAFT_CORPORATION: wRet = "A"
'    Case AEROSPACE_INDUSTRIAL_DEVELOPMENT: wRet = "B"
'    Case Eurocopter_Deutschland_GmbH: wRet = "E"
'    Case GOODRICH_LIGHTING_SYSTEMS_GMBH: wRet = "G"
'    Case HAMILTON_SUNDSTRAND_CORP: wRet = "H"
'    Case INTERTECHNIQUE: wRet = "I"
'    Case HEATH_TECNA: wRet = "J"
'    Case KORRY_ELECTRONICS_CO: wRet = "K"
'    Case LMI_AEROSPACE_INC: wRet = "L"
'    Case GOODRICH_SENSOR_SYSTEMS: wRet = "M"
'    Case NABTESCO_CORP: wRet = "N"
'    Case PRATT_AND_WHITNEY: wRet = "P"
'    Case PARKER_HANNIFIN_CORP: wRet = "Q"
'    Case ROCKWELL_COLLINS_INC: wRet = "R"
'    Case SPIRIT: wRet = "S"
'    Case SUMITOMO_PRECISION_PRODUCTS_CO_LTD: wRet = "T"
'    Case SAINT_GOBAIN_PERFORMANCE_PLASTICS: wRet = "U"
'    Case Else: wRet = "unknown"
    End Select
    
    GetRPCCode = wRet
End Function

''MRJ BR MIクラスを使って書き直す
'
'    Select Case sModelIdentCode
'        Case "MRJ", "PW1000G"
'            IsValidDMC_ = True
''        Exit Function
'        Case "MWG", "MBR", "S1000D"
'            IsValidDMC_ = False
'        Case Else
'            IsValidDMC_ = False
'    End Select

Public Function GetMI(MICodeIndex As MRJ_BR_ModelIdentCode) As String
    Dim wRet As String
    Select Case MICodeIndex
    Case MRJ: wRet = "MRJ"
    Case MBR: wRet = "MBR"
    Case MWG: wRet = "MWG"
    Case PW1000G: wRet = "PW1000G"
    Case S1000D: wRet = "S1000D"
    Case Else: wRet = "unknown"
    End Select
    
    GetMI = wRet
End Function
