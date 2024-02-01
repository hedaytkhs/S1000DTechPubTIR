VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolsTIR"
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

        'Vendor Code登録チェック
        If MetadataItem.columnID = ToolsTIRColumn.ManufacturerCode Then
            tpRet.VendorCode = MetadataItem.Value
            CheckMissingVendorCode tpRet
        End If
    
    Next MetadataItem

    GetMetadataRow = tpRet
End Function

Function CheckVendorCode(ByRef sVendorCode As String, ByRef MetadataItem As clsESRDMetadataElement) As Boolean
    MsgBox sVendorCode & " をチェックします.", vbInformation + vbOKOnly & "VendorCodeチェック"
End Function

Private Sub Class_Initialize()
    Set MetadataItems = New Collection
    
    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "TIR Type"
        .Value = "Tools"
        .columnID = ToolsTIRColumn.TIRType
        .requirement = MandatoryText
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Status"
        .Value = ""
        .columnID = ToolsTIRColumn.Status
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Item Identifier"
        .Value = ""
        .columnID = ToolsTIRColumn.ItemIdentifier
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Tool Number"
        .Value = ""
        .columnID = ToolsTIRColumn.ToolNumber
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Tool Name"
        .Value = ""
        .columnID = ToolsTIRColumn.ToolName
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Manufacture Code"
        .Value = ""
        .columnID = ToolsTIRColumn.ManufacturerCode
        .requirement = MandatoryText
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Short Name"
        .Value = ""
        .columnID = ToolsTIRColumn.ShortName
        .requirement = Optionaltext
        .IsAllowableValue = False
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Alternate Tool Number"
        .Value = ""
        .columnID = ToolsTIRColumn.AlternateToolNumber
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Alternate Tool Description"
        .Value = ""
        .columnID = ToolsTIRColumn.AlternateToolDescription
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Over length part number"
        .Value = ""
        .columnID = ToolsTIRColumn.OverLengthPartNumber
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Procurement data"
        .Value = ""
        .columnID = ToolsTIRColumn.ProcurementData
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Remarks"
        .Value = ""
        .columnID = ToolsTIRColumn.Remarks
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Comment"
        .Value = ""
        .columnID = ToolsTIRColumn.Comment
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem

    Set MetadataItem = New clsESRDMetadataElement
    With MetadataItem
        .name = "Source"
        .Value = ""
        .columnID = ToolsTIRColumn.Source
        .requirement = Optionaltext
        .IsAllowableValue = True
    End With
    MetadataItems.Add MetadataItem
End Sub

Private Sub Class_Terminate()
    Set MetadataItems = Nothing
End Sub




