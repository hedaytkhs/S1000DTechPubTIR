VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTIRCircuitBreaker"
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



'Public Enum CircuitBreakersTIRColumn
'    TIRType = 1
'    Status = 2
'    ItemIdentifier = 3
'    CBNumber = 4
'    CBName = 5
'    CBClass = 6
'    Comment = 7
'    Source = 8
'End Enum


Private Sub Class_Initialize()
    Set MetadataItems = New Collection
    
    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "TIR Type"
        .Value = "Circuit breakers"
        .columnID = CircuitBreakersTIRColumn.TIRType
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Status"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.Status
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Item Identifier"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.ItemIdentifier
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "CB Number"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.CBNumber
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "CB Name"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.CBName
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "CB Class"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.CBClass
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comment"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.Comment
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Source"
        .Value = ""
        .columnID = CircuitBreakersTIRColumn.Source
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub





