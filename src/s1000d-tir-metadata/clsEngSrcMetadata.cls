VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEngSrcMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    ESRDのEngineeringSourceの定義に従ってメタデータを操作するためのクラス
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
        .name = "File Category"
        .Value = "Author"
        .columnID = EngSrcMetadataColumn.FileCategory
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Name"
        .Value = ""
        .columnID = EngSrcMetadataColumn.filename
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Issue"
        .Value = ""
        .columnID = EngSrcMetadataColumn.FileIssue
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Title"
        .Value = ""
        .columnID = EngSrcMetadataColumn.FileTitle
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "File Format"
        .Value = ""
        .columnID = EngSrcMetadataColumn.FileFormat
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Engineering Source Category"
        .Value = ""
        .columnID = EngSrcMetadataColumn.EngineeringSourceCategory
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Responsible Department"
        .Value = ""
        .columnID = EngSrcMetadataColumn.ResponsibleDepartment
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Aircraft Model"
        .Value = ""
        .columnID = EngSrcMetadataColumn.AircraftModel
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "VridgeR Reference"
        .Value = ""
        .columnID = EngSrcMetadataColumn.VridgeRReference
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Zone"
        .Value = ""
        .columnID = EngSrcMetadataColumn.Zone
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Access Point"
        .Value = ""
        .columnID = EngSrcMetadataColumn.AccessPoint
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Part Number"
        .Value = ""
        .columnID = EngSrcMetadataColumn.PartNumber
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Export Control"
        .Value = ""
        .columnID = EngSrcMetadataColumn.ExportControl
        .requirement = ConditionalText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comments"
        .Value = ""
        .columnID = EngSrcMetadataColumn.Comments
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Active"
        .Value = ""
        .columnID = EngSrcMetadataColumn.Active
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Change Number"
        .Value = ""
        .columnID = EngSrcMetadataColumn.ChangeNumber
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Original Engineering Source ID"
        .Value = ""
        .columnID = EngSrcMetadataColumn.OriginalEngineeringSourceID
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub

