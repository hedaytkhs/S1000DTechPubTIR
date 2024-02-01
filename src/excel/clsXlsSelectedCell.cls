VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXlsSelectedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Row As Long
Public Column As Long
Dim Text_ As String


Property Let Text(ByVal sText As String)
    Text_ = sText
End Property

Property Get Text() As String
    Text = Text_
End Property

