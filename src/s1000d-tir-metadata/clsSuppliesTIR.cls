VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSuppliesTIR"
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

Public Function GetMetadataRow() As tpIntegrationRowTextWithResult
    Dim tpRet As tpIntegrationRowTextWithResult
    tpRet.Text = ""
    For Each MetadataItem In MetadataItems
        tpRet.Text = tpRet.Text & MetadataItem.Value & cESRD_MetadataSeparator

        'Vendor Codeìoò^É`ÉFÉbÉN
        If MetadataItem.columnID = SuppliesTIRColumn.ManufacturerCode Then
            tpRet.VendorCode = MetadataItem.Value
            CheckMissingVendorCode tpRet
        End If
    
    Next MetadataItem

    GetMetadataRow = tpRet
End Function

Private Sub Class_Initialize()
    Set MetadataItems = New Collection
    
    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "TIR Type"
        .Value = "Supplies"
        .columnID = SuppliesTIRColumn.TIRType
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Status"
        .Value = ""
        .columnID = SuppliesTIRColumn.Status
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Item Identifier"
        .Value = ""
        .columnID = SuppliesTIRColumn.ItemIdentifier
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Supply Number"
        .Value = ""
        .columnID = SuppliesTIRColumn.SupplyNumber
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Supply Number Type"
        .Value = ""
        .columnID = SuppliesTIRColumn.SupplyNumberType
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Supply Name"
        .Value = ""
        .columnID = SuppliesTIRColumn.SupplyName
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Manufacturer Code"
        .Value = ""
        .columnID = SuppliesTIRColumn.ManufacturerCode
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Short Name"
        .Value = ""
        .columnID = SuppliesTIRColumn.ShortName
        .requirement = Optionaltext
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Locally Supplied Flag"
        .Value = ""
        .columnID = SuppliesTIRColumn.LocallySuppliedFlag
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comment"
        .Value = ""
        .columnID = SuppliesTIRColumn.Comment
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Source"
        .Value = ""
        .columnID = SuppliesTIRColumn.Source
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub



