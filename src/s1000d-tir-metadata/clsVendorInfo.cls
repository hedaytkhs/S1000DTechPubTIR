VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVendorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Vendor_ As clsVendorCode
Dim InfoVendors_ As Collection

Property Get ItemCount() As Long
    ItemCount = InfoVendors_.Count
End Property

Public Function AddVendorInfo(ByRef VendorInfo As clsVendorCode) As String
    InfoVendors_.Add VendorInfo
End Function


Private Sub Class_Initialize()
    Set Vendor_ = New clsVendorCode
    Set InfoVendors_ = New Collection
End Sub

Private Sub Class_Terminate()
    Set Vendor_ = Nothing
    Set InfoVendors_ = Nothing
End Sub

