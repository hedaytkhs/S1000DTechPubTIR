VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConvertedDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'****************************************************************************************
'
'    ESRDのConverted DMの定義に従ってメタデータを操作するためのクラス
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

Public MetadataItem As clsESRDMetadataElement
Public MetadataItems As Collection


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
    GetMetadataRow = sRet
End Function

Private Sub Class_Initialize()
    Set MetadataItems = New Collection

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Category"
        .Value = "ConvertedDM"
        .columnID = ConvertedDMColumn.[File Category]
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Name"
        .Value = ""
        .columnID = ConvertedDMColumn.[File Name]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Issue"
        .Value = ""
        .columnID = ConvertedDMColumn.[File Issue]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Title"
        .Value = ""
        .columnID = ConvertedDMColumn.[File Title]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Format"
        .Value = ""
        .columnID = ConvertedDMColumn.[File Format]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Included In Manual"
        .Value = ""
        .columnID = ConvertedDMColumn.[Included In Manual]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Aircraft Model"
        .Value = ""
        .columnID = ConvertedDMColumn.[Aircraft Model]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Export Control"
        .Value = ""
        .columnID = ConvertedDMColumn.[Export Control]
        .requirement = ConditionalText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comments"
        .Value = ""
        .columnID = ConvertedDMColumn.Comments
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Active"
        .Value = ""
        .columnID = ConvertedDMColumn.Active
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Change Number"
        .Value = ""
        .columnID = ConvertedDMColumn.[Change Number]
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub


