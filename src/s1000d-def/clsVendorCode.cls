VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVendorCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    ベンダーコード情報の受け渡し時に情報を保持するためのクラス
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

Dim VendorCode_ As String * 5
Dim AlternateCode_ As String
Dim AlternateCodeType_ As String
Dim VendorName_ As String
Dim BusinessUnitName_ As String
Dim City_ As String
Dim Country_ As String
Dim ZIPCode_ As String
Dim Street_ As String
Dim PhoneNumber_ As String
Dim FAX_ As String
Dim EMail_ As String
Dim URL_ As String
Dim Comment_ As String
Dim Source_ As String
Dim HasVendorInfo_ As Boolean

Property Let HasVendorInfo(ByRef RSStatus As Boolean)
    HasVendorInfo_ = RSStatus
End Property

Property Get HasVendorInfo() As Boolean
    HasVendorInfo = HasVendorInfo_
End Property

Property Get VendorCode() As String
    VendorCode = VendorCode_
End Property

Property Let VendorCode(ByRef sVendorCode As String)
    VendorCode_ = sVendorCode
End Property

Property Get AlternateCode() As String
    AlternateCode = AlternateCode_
End Property

Property Let AlternateCode(ByRef sArgument As String)
    AlternateCode_ = sArgument
End Property

Property Get AlternateCodeType() As String
    AlternateCodeType = AlternateCodeType_
End Property

Property Let AlternateCodeType(ByRef sArgument As String)
    AlternateCodeType_ = sArgument
End Property

Property Get VendorName() As String
    VendorName = VendorName_
End Property

Property Let VendorName(ByRef sArgument As String)
    VendorName_ = sArgument
    HasVendorInfo_ = True
End Property

Property Get BusinessUnitName() As String
    BusinessUnitName = BusinessUnitName_
End Property

Property Let BusinessUnitName(ByRef sArgument As String)
    BusinessUnitName_ = sArgument
End Property

Property Get City() As String
    City = City_
End Property

Property Let City(ByRef sArgument As String)
    City_ = sArgument
End Property

Property Get Country() As String
    Country = Country_
End Property

Property Let Country(ByRef sArgument As String)
    Country_ = sArgument
End Property

Property Get ZIPCode() As String
    ZIPCode = ZIPCode_
End Property

Property Let ZIPCode(ByRef sArgument As String)
    ZIPCode_ = sArgument
End Property

Property Get Street() As String
    Street = Street_
End Property

Property Let Street(ByRef sArgument As String)
    Street_ = sArgument
End Property

Property Get PhoneNumber() As String
    PhoneNumber = PhoneNumber_
End Property

Property Let PhoneNumber(ByRef sArgument As String)
    PhoneNumber_ = sArgument
End Property

Property Get FAX() As String
    FAX = FAX_
End Property

Property Let FAX(ByRef sArgument As String)
    FAX_ = sArgument
End Property

Property Get EMail() As String
    EMail = EMail_
End Property

Property Let EMail(ByRef sArgument As String)
    EMail_ = sArgument
End Property

Property Get URL() As String
    URL = URL_
End Property

Property Let URL(ByRef sArgument As String)
    URL_ = sArgument
End Property

Property Get Comment() As String
    Comment = Comment_
End Property

Property Let Comment(ByRef sArgument As String)
    Comment_ = sArgument
End Property

Property Get Source() As String
    Source = Source_
End Property

Property Let Source(ByRef sArgument As String)
    Source_ = sArgument
End Property

Private Sub Class_Initialize()
    HasVendorInfo_ = False
End Sub

