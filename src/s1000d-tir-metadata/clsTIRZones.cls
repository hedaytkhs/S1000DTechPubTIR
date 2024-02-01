VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTIRZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
        .name = "TIR Type"
        .Value = "Zones"
        .columnID = ZonesTIRColumn.TIRType
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Status"
        .Value = ""
        .columnID = ZonesTIRColumn.Status
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Item Identifier"
        .Value = ""
        .columnID = ZonesTIRColumn.ItemIdentifier
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Zone Number"
        .Value = ""
        .columnID = ZonesTIRColumn.ZoneNumber
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Zone Description"
        .Value = ""
        .columnID = ZonesTIRColumn.zonedescription
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Applicability"
        .Value = ""
        .columnID = ZonesTIRColumn.Applicability
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem



    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comment"
        .Value = ""
        .columnID = ZonesTIRColumn.Comment
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Source"
        .Value = ""
        .columnID = ZonesTIRColumn.Source
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub







