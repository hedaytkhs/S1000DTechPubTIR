Attribute VB_Name = "mdlS1000DCommon"
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
' ///// 2014/04/03/////
' 作成開始
'
' Created by
' Hideaki Takahashi
'
'
'=========================================================
Option Explicit

Public Enum S1000D_URN
    DMC = 1
    DME = 2
    PMC = 3
    PME = 4
    SMC = 5
    SME = 6
    CSN = 7
    ICN = 8
    COM = 9
    DDN = 10
    DML = 11
    UPF = 12
    UPE = 13
End Enum

Public Enum S1000D_FORMAT
    XML = 1
    CGM = 2
    CG4 = 3
    TIF = 4
    JPG = 5
    PNG = 6
    GIF = 7
    pdf = 8
End Enum

Public Enum S1000D_IssueType
    [New] = 1
    Changed = 2
    Revised = 3
End Enum



Public Function GetURN(URNIndex As S1000D_URN) As String
    Dim wRet As String
    Select Case URNIndex
    Case DMC: wRet = "DMC"
    Case DME: wRet = "DME"
    Case PMC: wRet = "PMC"
    Case PME: wRet = "PME"
    Case SMC: wRet = "SMC"
    Case SME: wRet = "SME"
    Case CSN: wRet = "CSN"
    Case ICN: wRet = "ICN"
    Case COM: wRet = "COM"
    Case DDN: wRet = "DDN"
    Case DML: wRet = "DML"
    Case UPF: wRet = "UPF"
    Case UPE: wRet = "UPE"
'    Case Else: wRet = "unknown"
    End Select
    
    GetURN = wRet
End Function

Public Function GetS1000DExtension(S1000D_FORMATIndex As S1000D_FORMAT) As String
    Dim wRet As String
    Select Case S1000D_FORMATIndex
    Case XML: wRet = "XML"
    Case CGM: wRet = "CGM"
    Case CG4: wRet = "CG4"
    Case TIF: wRet = "TIF"
    Case JPG: wRet = "JPG"
    Case PNG: wRet = "PNG"
    Case GIF: wRet = "GIF"
    Case pdf: wRet = "PDF"
'    Case Else: wRet = "unknown"
    End Select
    
    GetS1000DExtension = wRet
End Function

