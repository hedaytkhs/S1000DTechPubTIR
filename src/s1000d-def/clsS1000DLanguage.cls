VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsS1000DLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'
' MRJ BRÇÃíËã`Ç…è]Ç¢ Langage Code/Country CodeÇèàóùÇ∑ÇÈÉNÉâÉX
'
'
' Hideaki Takahashi
' ///// 2014/03/18 /////
'
'
'
'=========================================================

Option Explicit

Dim LanguageCodeIndex_ As MRJ_BR_LanguageCode
Dim ContryCodeIndex_ As MRJ_BR_ContryCode

Dim countryIsoCode_ As String * 2
Dim languageIsoCode_ As String * 2
'For attribute languageIsoCode, the language code specified in ISO 639 must be used
'when delivered. The valid language code for MRJ is specified in Table 4.
Dim LangueAndCountry_ As String * 5

Property Get LangueAndCountry() As String
    LangueAndCountry = LangueAndCountry_
End Property

Property Get countryIsoCode() As String
    countryIsoCode = countryIsoCode_
End Property

Property Get languageIsoCode() As String
    languageIsoCode = languageIsoCode_
End Property


Private Sub Class_Initialize()
    LanguageCodeIndex_ = MRJ_BR_LanguageCode.SimplifiedTechnicalEnglish
    ContryCodeIndex_ = MRJ_BR_ContryCode.UnitedStates
    Call GetCountryIsoCode
    Call GetLanguageIsoCode
    Call GetLangueAndCountry
End Sub

Private Sub GetLangueAndCountry()
    LangueAndCountry_ = UCase(languageIsoCode_ & "-" & countryIsoCode_)
End Sub

Private Sub GetCountryIsoCode()
    Dim wRet As String
    Select Case ContryCodeIndex_
    Case Japan: wRet = "JP"
    Case UnitedStates: wRet = "US"
    Case Else: wRet = ""
    End Select
    countryIsoCode_ = wRet
End Sub

Private Sub GetLanguageIsoCode()
    Dim wRet As String
    Select Case LanguageCodeIndex_
    Case SimplifiedTechnicalEnglish: wRet = "sx"
    Case English: wRet = "en"
    Case Else: wRet = ""
    End Select
    languageIsoCode_ = wRet
End Sub

