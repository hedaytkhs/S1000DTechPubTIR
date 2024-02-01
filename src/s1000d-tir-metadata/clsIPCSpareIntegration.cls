VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIPCSpareIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    ESRDのIPCSpareIntegrationの定義に従ってメタデータを操作するためのクラス
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

Public MetadataItem As clsESRDMetadataElement
Public MetadataItems As Collection
Public IsApplicapleMetadata As Boolean


Public Function GetTitleRow() As String
    Dim sRet As String
    sRet = ""
    For Each MetadataItem In MetadataItems
        sRet = sRet & MetadataItem.name & cESRD_MetadataSeparator
    Next MetadataItem
    GetTitleRow = sRet
End Function

Public Function GetMetadataRow() As String
    Dim sRet As String
    sRet = ""
    For Each MetadataItem In MetadataItems
        sRet = sRet & MetadataItem.Value & cESRD_MetadataSeparator
    Next MetadataItem

#If (DEBUG_MODE = 0) Then
    Debug.Print sRet
#End If
    
    GetMetadataRow = sRet
End Function

Private Sub Class_Initialize()
    Set MetadataItems = New Collection
    
    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Type"
        .Value = "Part"
        .columnID = IPCSpareIntegrationColumn.Type_
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Status"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.Status
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Part Nbr"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.[Part Nbr]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Name"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.name
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "SCD"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.SCD
        .requirement = Optionaltext
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "OPN"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.OPN
        .requirement = Optionaltext
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Vendor Code"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.[Vendor Code]
        .requirement = ValidVendorCode
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Internal Notes"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.[Internal Notes]
        .requirement = Optionaltext
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comment"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.Comment
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Source"
        .Value = ""
        .columnID = IPCSpareIntegrationColumn.Source
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub



