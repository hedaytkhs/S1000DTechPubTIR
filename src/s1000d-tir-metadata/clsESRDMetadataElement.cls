VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsESRDMetadataElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'
'    ESRD�̃��^�f�[�^�v�f�𑀍삷�邽�߂̃N���X
'
'    ���^�f�[�^�̖��́A�l�`�F�b�N�Ȃǂ̋��ʗv�f���܂Ƃ߂�
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************

Public columnID As Long
Public name As String
Public Value As String
Public requirement As ESRDRequirementForElement
Public IsAllowableValue As Boolean
