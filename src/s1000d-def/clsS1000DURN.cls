VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsS1000DURN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'
' S1000DÇÃíËã`Ç…è]Ç¢ URNÇèàóùÇ∑ÇÈÉNÉâÉX
'
'
' Hideaki Takahashi
' ///// 2014/03/18 /////
'
'
'
'=========================================================

Option Explicit

Dim URN_Index_ As S1000D_URN
Dim URN_ As String * 3

Property Let URNCategory(ByVal URN_Index As S1000D_URN)
    URN_Index_ = URN_Index
    URN_ = GetURN(URN_Index)
End Property

Property Get URN() As String
    URN = URN_
End Property

