VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetadataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum XlsCellTIRInit
    Row = 2
    col = 1
End Enum
Private Enum TIRFilenameColumn
    Tools = 17
    SUPPLIES = 13
    Enterprise = 19
    Zones = 9
    AccessPoints = 11
    CircuitBreakers = 17
End Enum

Dim MetadataSheet As Worksheet

Dim LastCol_ As Long
Dim LastRow_ As Long
Dim InitCol_ As Long
Dim InitRow_ As Long

Dim MetadataFilename_ As String

Property Let InitCol(ByVal lInitCol As Long)
    InitCol_ = lInitCol
End Property

Property Let InitRow(ByVal lInitRow As Long)
    InitRow_ = lInitRow
End Property

Property Get LastCol() As Long
    LastCol = LastCol_
End Property

Property Get LastRow() As Long
    LastRow = LastRow_
End Property

Property Let MetadataFilename(ByVal sFilename As String)
    MetadataFilename_ = sFilename
End Property

Property Get MetadataFilename() As String
    MetadataFilename = MetadataFilename_
End Property

Public Function PutMetadataFilename(ByRef TIRType As ESRDFileCategory, ByRef sMetadataFilename As String, ByRef lCurrentRow As Long) As Boolean
    If TIRType = ESRDFileCategory.Tools Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Tools).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Tools).Value = MetadataFilename_
        End If
    ElseIf TIRType = ESRDFileCategory.SUPPLIES Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.SUPPLIES).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.SUPPLIES).Value = MetadataFilename_
        End If
    ElseIf TIRType = ESRDFileCategory.Enterprise Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Enterprise).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Enterprise).Value = MetadataFilename_
        End If
    ElseIf TIRType = ESRDFileCategory.Zones Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Zones).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.Zones).Value = MetadataFilename_
        End If
    ElseIf TIRType = ESRDFileCategory.AccessPoints Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.AccessPoints).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.AccessPoints).Value = MetadataFilename_
        End If
    ElseIf TIRType = ESRDFileCategory.CircuitBreakers Then
        If sMetadataFilename <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.CircuitBreakers).Value = sMetadataFilename
        ElseIf MetadataFilename_ <> "" Then
            MetadataSheet.Cells(lCurrentRow, TIRFilenameColumn.CircuitBreakers).Value = MetadataFilename_
        End If
    End If
End Function

Public Function IsValidMetadateSheet(ByRef TIRType As ESRDFileCategory) As Boolean
    If TIRType = ESRDFileCategory.Tools Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Tools" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.SUPPLIES Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Supplies" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.Enterprise Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Enterprise" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.Zones Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Zones" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.AccessPoints Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Access-points" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.CircuitBreakers Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "CircuitBreakers" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    ElseIf TIRType = ESRDFileCategory.Author Then
        If MetadataSheet.Cells(XlsCellTIRInit.Row, XlsCellTIRInit.col).Value <> "Author" Then
            IsValidMetadateSheet = False
            Exit Function
        Else
            IsValidMetadateSheet = True
            Exit Function
        End If
    End If
End Function

Private Sub Class_Initialize()
    If ActiveSheet Is Nothing Then Exit Sub
    Set MetadataSheet = ActiveSheet
    LastCol_ = GetLastColOfThisSheet(MetadataSheet)
    LastRow_ = GetLastRowOfThisSheet(MetadataSheet)
End Sub

Private Sub Class_Terminate()
    Set MetadataSheet = Nothing
End Sub
