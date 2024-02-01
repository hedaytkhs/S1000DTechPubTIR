Attribute VB_Name = "mdlMenu"
Option Explicit

Public TIR_Type_Text As String, strCaption As String

Private Const CNST_FormCaption As String = "�̐ݒ�"
Private Const CNST_TIR_Prefix As String = "TIR-"
    
#If False Then
    Private Const CNST_IHS_Login_URL_DefaultValue As String = "file://ycab5s21/workspc/MITAC_HP_Contents/MRJ_Link_071_�h�g�r�i�C�O�K��j/MITAC_ERC_201408.html"
#End If

Sub DM�����c�[���̃o�[�W��������\������(ByRef myButton As IRibbonControl)
    MsgBox GetAppCNST(Version) & vbCrLf & "�ŏI�X�V���F " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppNameDMFolderOpen)
End Sub

Sub TIR�����c�[���̃o�[�W��������\������(ByRef myButton As IRibbonControl)
    MsgBox GetAppCNST(Version) & vbCrLf & "�ŏI�X�V���F " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppNameTIRSearch)
End Sub

Sub TIR�c�[���̃o�[�W��������\������(ByRef myButton As IRibbonControl)
    Call ShowVersion
End Sub

Private Sub ShowVersion()
    MsgBox GetAppCNST(Version) & vbCrLf & "�ŏI�X�V���F " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppName)
End Sub


Sub ToolsTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIRTop�t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsTIRIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersIntegrationFile�`�F�b�N���N������(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIRIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIRIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIRIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIRIntegrationFile�ۑ��t�H���_���J��(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIR���M���[�����쐬(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.CircuitBreakers)
End Sub

Sub ToolsTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub ToolsSupplies�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIR�ŐV�G�N�Z�����J��(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.CircuitBreakers)
End Sub

'Sub TIRTools�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Tools)
'End Sub
'
'Sub TIRSupplies�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Supplies)
'End Sub
'
'Sub TIREnterprise�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Enterprise)
'End Sub
'
'Sub TIRZones�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Zones)
'End Sub
'
'Sub TIRAccessPoints�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.AccessPoints)
'End Sub
'
'Sub TIRCircuitBreakers�̌�����ʂ�\������(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.CircuitBreakers)
'End Sub


Sub TIRTools�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Tools)
End Sub

Sub TIRSupplies�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.SUPPLIES)
End Sub

Sub TIREnterprise�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Enterprise)
End Sub

Sub TIRZones�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Zones)
End Sub

Sub TIRAccessPoints�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.AccessPoints)
End Sub

Sub TIRCircuitBreakers�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.CircuitBreakers)
End Sub

Sub TIR��ʐݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    frmCommonSettings.Show
End Sub

Sub DM���؃t�H���_�̐ݒ��ʂ�\������(ByRef myButton As IRibbonControl)
    frmTTSettings.Show
End Sub

Private Sub OpenTIRTopFolder(ByRef TIRType As ESRDFileCategory)
    On Error GoTo errHandler
    Dim strFolder As String
    If TIRType = Tools Then
        Dim TIRToolsSetting As clsConfigTIRTools
        Set TIRToolsSetting = New clsConfigTIRTools
        If TIRToolsSetting.MACSTopFolderExists Then
            strFolder = TIRToolsSetting.MACSTopFolder
        Else
            Call ShowSetting(Tools)
            Exit Sub
        End If
        Set TIRToolsSetting = Nothing
    
    ElseIf TIRType = SUPPLIES Then
        Dim TIRSuppliesSetting As clsConfigTIRSupplies
        Set TIRSuppliesSetting = New clsConfigTIRSupplies
        If TIRSuppliesSetting.MACSTopFolderExists Then
            strFolder = TIRSuppliesSetting.MACSTopFolder
        Else
            Call ShowSetting(SUPPLIES)
            Exit Sub
        End If
        Set TIRSuppliesSetting = Nothing
    
    ElseIf TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        If TIREnterpriseSetting.MACSTopFolderExists Then
            strFolder = TIREnterpriseSetting.MACSTopFolder
        Else
            Call ShowSetting(Enterprise)
            Exit Sub
        End If
        Set TIREnterpriseSetting = Nothing
    
    ElseIf TIRType = Zones Then
        Dim TIRZonesSetting As clsConfigTIRZones
        Set TIRZonesSetting = New clsConfigTIRZones
        If TIRZonesSetting.MACSTopFolderExists Then
            strFolder = TIRZonesSetting.MACSTopFolder
        Else
            Call ShowSetting(Zones)
            Exit Sub
        End If
        Set TIRZonesSetting = Nothing
    
    ElseIf TIRType = AccessPoints Then
        Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
        Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
        If TIRAccessPointsSetting.MACSTopFolderExists Then
            strFolder = TIRAccessPointsSetting.MACSTopFolder
        Else
            Call ShowSetting(AccessPoints)
            Exit Sub
        End If
        Set TIRAccessPointsSetting = Nothing
    
    ElseIf TIRType = CircuitBreakers Then
        Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
        Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
        If TIRCircuitBreakersSetting.MACSTopFolderExists Then
            strFolder = TIRCircuitBreakersSetting.MACSTopFolder
        Else
            Call ShowSetting(CircuitBreakers)
            Exit Sub
        End If
        Set TIRCircuitBreakersSetting = Nothing
    
    End If
    
    Call Shell("explorer """ & strFolder & "", vbNormalFocus)
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub OpenIntegrationFileFolder(ByRef TIRType As ESRDFileCategory)
    On Error GoTo errHandler
    Dim strFolder As String
    If TIRType = Tools Then
        Dim TIRToolsSetting As clsConfigTIRTools
        Set TIRToolsSetting = New clsConfigTIRTools
        If TIRToolsSetting.IntegrationFileFolderExists Then
            strFolder = TIRToolsSetting.IntegrationFileFolder
        Else
            Call ShowSetting(Tools)
            Exit Sub
        End If
        Set TIRToolsSetting = Nothing
    
    ElseIf TIRType = SUPPLIES Then
        Dim TIRSuppliesSetting As clsConfigTIRSupplies
        Set TIRSuppliesSetting = New clsConfigTIRSupplies
        If TIRSuppliesSetting.IntegrationFileFolderExists Then
            strFolder = TIRSuppliesSetting.IntegrationFileFolder
        Else
            Call ShowSetting(SUPPLIES)
            Exit Sub
        End If
        Set TIRSuppliesSetting = Nothing
    
    ElseIf TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        If TIREnterpriseSetting.IntegrationFileFolderExists Then
            strFolder = TIREnterpriseSetting.IntegrationFileFolder
        Else
            Call ShowSetting(Enterprise)
            Exit Sub
        End If
        Set TIREnterpriseSetting = Nothing
    
    ElseIf TIRType = Zones Then
        Dim TIRZonesSetting As clsConfigTIRZones
        Set TIRZonesSetting = New clsConfigTIRZones
        If TIRZonesSetting.IntegrationFileFolderExists Then
            strFolder = TIRZonesSetting.IntegrationFileFolder
        Else
            Call ShowSetting(Zones)
            Exit Sub
        End If
        Set TIRZonesSetting = Nothing
    
    ElseIf TIRType = AccessPoints Then
        Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
        Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
        If TIRAccessPointsSetting.IntegrationFileFolderExists Then
            strFolder = TIRAccessPointsSetting.IntegrationFileFolder
        Else
            Call ShowSetting(AccessPoints)
            Exit Sub
        End If
        Set TIRAccessPointsSetting = Nothing
    
    ElseIf TIRType = CircuitBreakers Then
        Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
        Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
        If TIRCircuitBreakersSetting.IntegrationFileFolderExists Then
            strFolder = TIRCircuitBreakersSetting.IntegrationFileFolder
        Else
            Call ShowSetting(CircuitBreakers)
            Exit Sub
        End If
        Set TIRCircuitBreakersSetting = Nothing
    
    End If
    
    Call Shell("explorer """ & strFolder & "", vbNormalFocus)
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub ShowSetting(ByRef TIRType As ESRDFileCategory)
    Dim iRet As Integer
    iRet = MsgBox("�ݒ肳��Ă���t�H���_�E�t�@�C����������܂���." & vbCrLf & _
                "�t�H���_�E�t�@�C���̐ݒ���s���܂���?", vbExclamation + vbYesNoCancel, GetAppCNST(AppName))
    If iRet = vbYes Then
        If TIRType = Tools Then
            Call UpdateCaptionAndShowForm(Tools)
        ElseIf TIRType = SUPPLIES Then
            Call UpdateCaptionAndShowForm(SUPPLIES)
        ElseIf TIRType = Enterprise Then
            Call UpdateCaptionAndShowForm(Enterprise)
        ElseIf TIRType = Zones Then
            Call UpdateCaptionAndShowForm(Zones)
        ElseIf TIRType = AccessPoints Then
            Call UpdateCaptionAndShowForm(AccessPoints)
        ElseIf TIRType = CircuitBreakers Then
            Call UpdateCaptionAndShowForm(CircuitBreakers)
        End If
    End If
End Sub

Private Sub CheckIntegrationFile(ByRef TIRType As ESRDFileCategory)
On Error GoTo errHandler
    Dim strTIRIntegrationFilePath As String, strCheckPGPath As String
    If TIRType = Tools Then
        Dim TIRToolsSetting As clsConfigTIRTools
        Set TIRToolsSetting = New clsConfigTIRTools
        With TIRToolsSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIRToolsSetting = Nothing
    
    ElseIf TIRType = SUPPLIES Then
        Dim TIRSuppliesSetting As clsConfigTIRSupplies
        Set TIRSuppliesSetting = New clsConfigTIRSupplies
        With TIRSuppliesSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIRSuppliesSetting = Nothing
    
    ElseIf TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        With TIREnterpriseSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIREnterpriseSetting = Nothing
    
    ElseIf TIRType = Zones Then
        Dim TIRZonesSetting As clsConfigTIRZones
        Set TIRZonesSetting = New clsConfigTIRZones
        With TIRZonesSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIRZonesSetting = Nothing
    
    ElseIf TIRType = AccessPoints Then
        Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
        Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
        With TIRAccessPointsSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIRAccessPointsSetting = Nothing
    
    ElseIf TIRType = CircuitBreakers Then
        Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
        Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
        With TIRCircuitBreakersSetting
            If .IntegrationFileCheckPGExists Then
                strCheckPGPath = .IntegrationFileCheckPGPath
                If Not .IntegrationFileExists Then
                    MsgBox "�`�F�b�N�Ώۂ�IntegrationFile������܂���.", vbExclamation + vbOKOnly, "Integration File�`�F�b�N"
                    Exit Sub
                Else
                    strTIRIntegrationFilePath = .IntegrationFilePath
                End If
            Else
                Call ShowSetting(Tools)
                Exit Sub
            End If
        End With
        Set TIRCircuitBreakersSetting = Nothing
    
    End If
    
    Call Shell("WScript.exe """ & strCheckPGPath & """ " & strTIRIntegrationFilePath & "")
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub OpenTIRXlsFile(ByRef TIRType As ESRDFileCategory)
On Error GoTo errHandler
    Dim strTIRXlsPath As String
    If TIRType = Tools Then
        Dim TIRToolsSetting As clsConfigTIRTools
        Set TIRToolsSetting = New clsConfigTIRTools
        If TIRToolsSetting.LatestExcelExists Then
             strTIRXlsPath = TIRToolsSetting.LatestExcelPath
        Else
            Call ShowSetting(Tools)
            Exit Sub
        End If
        Set TIRToolsSetting = Nothing
    
    ElseIf TIRType = SUPPLIES Then
        Dim TIRSuppliesSetting As clsConfigTIRSupplies
        Set TIRSuppliesSetting = New clsConfigTIRSupplies
        If TIRSuppliesSetting.LatestExcelExists Then
             strTIRXlsPath = TIRSuppliesSetting.LatestExcelPath
        Else
            Call ShowSetting(SUPPLIES)
            Exit Sub
        End If
        Set TIRSuppliesSetting = Nothing
    
    ElseIf TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        If TIREnterpriseSetting.LatestExcelExists Then
             strTIRXlsPath = TIREnterpriseSetting.LatestExcelPath
        Else
            Call ShowSetting(Enterprise)
            Exit Sub
        End If
        Set TIREnterpriseSetting = Nothing
    
    ElseIf TIRType = Zones Then
        Dim TIRZonesSetting As clsConfigTIRZones
        Set TIRZonesSetting = New clsConfigTIRZones
        If TIRZonesSetting.LatestExcelExists Then
             strTIRXlsPath = TIRZonesSetting.LatestExcelPath
        Else
            Call ShowSetting(Zones)
            Exit Sub
        End If
        Set TIRZonesSetting = Nothing
    
    ElseIf TIRType = AccessPoints Then
        Dim TIRAccessPointsSetting As clsConfigTIRAccessPoints
        Set TIRAccessPointsSetting = New clsConfigTIRAccessPoints
        If TIRAccessPointsSetting.LatestExcelExists Then
             strTIRXlsPath = TIRAccessPointsSetting.LatestExcelPath
        Else
            Call ShowSetting(AccessPoints)
            Exit Sub
        End If
        Set TIRAccessPointsSetting = Nothing
    
    ElseIf TIRType = CircuitBreakers Then
        Dim TIRCircuitBreakersSetting As clsConfigTIRCircuitBreakers
        Set TIRCircuitBreakersSetting = New clsConfigTIRCircuitBreakers
        If TIRCircuitBreakersSetting.LatestExcelExists Then
             strTIRXlsPath = TIRCircuitBreakersSetting.LatestExcelPath
        Else
            Call ShowSetting(CircuitBreakers)
            Exit Sub
        End If
        Set TIRCircuitBreakersSetting = Nothing
    
    End If
    
    Call OpenXlsFile(strTIRXlsPath)
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Sub IHS�փ��O�C������(ByRef myButton As IRibbonControl)
    Call IHSLogIn
End Sub

Private Sub IHSLogIn()
On Error GoTo errHandler

    Dim objIE As Object
    Dim str_Url As String
    
    Dim TIREnterpriseSetting As clsConfigTIREnterprise
    Set TIREnterpriseSetting = New clsConfigTIREnterprise
    str_Url = TIREnterpriseSetting.IHS_URL
    Set TIREnterpriseSetting = Nothing

#If False Then
    str_Url = CNST_IHS_Login_URL_DefaultValue
#End If
    
    Set objIE = CreateObject("InternetExplorer.Application")

    If Err.Number = 0 Then
        objIE.Navigate str_Url
        objIE.Visible = True
    Else
        GoTo errHandler
    End If

    Exit Sub
errHandler:
    Application.StatusBar = ("�G���[!�F " & Err.Number & ":" & Err.Description)
    MsgBox Err.Number & ":" & Err.Description
End Sub

Sub EnterpriseTIR�pDB���J��(ByRef myButton As IRibbonControl)
    Call OpenTIRDatabase(ESRDFileCategory.Enterprise)
End Sub

Private Sub OpenTIRDatabase(ByRef TIRType As ESRDFileCategory)
On Error GoTo errHandler
    Dim strTIRDatabasePath As String
    If TIRType = Enterprise Then
        Dim TIREnterpriseSetting As clsConfigTIREnterprise
        Set TIREnterpriseSetting = New clsConfigTIREnterprise
        If TIREnterpriseSetting.TIRDatabaseExists Then
             strTIRDatabasePath = TIREnterpriseSetting.TIRDatabasePath
        Else
            Call ShowSetting(Enterprise)
            Exit Sub
        End If
        Set TIREnterpriseSetting = Nothing
    End If
    
    Shell "msaccess.exe " & strTIRDatabasePath & "", vbNormalFocus
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub OpenXlsFile(ByRef XlsPath As String)
On Error GoTo errHandler

    Dim fso As Object
    Dim XlsWbName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    XlsWbName = fso.GetFile(XlsPath).name

    Dim wb As Workbook
    Dim bWbOpen As Boolean
    bWbOpen = False
    For Each wb In Workbooks
        If wb.name = XlsWbName Then
            bWbOpen = True
        End If
    Next wb
    If Not bWbOpen Then
        Workbooks.Open XlsPath, Notify:=False
    Else
        Workbooks(XlsWbName).Activate
    End If
    
    Set fso = Nothing
    
    Exit Sub
errHandler:
    MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub UpdateCaptionAndShowForm(ByVal TIRType As ESRDFileCategory)
    TIR_Type_Text = CNST_TIR_Prefix & GetFileCategoryName(TIRType)
    strCaption = TIR_Type_Text & CNST_FormCaption
    Select Case TIRType
    Case Tools:    frmTIRToolsSetting.Show
    Case SUPPLIES:    frmTIRSuppliesSetting.Show
    Case Enterprise:    frmTIREnterpriseSetting.Show
    Case Zones:    frmTIRZonesSetting.Show
    Case AccessPoints:    frmTIRAccessPointsSetting.Show
    Case CircuitBreakers:    frmTIRCircuitBreakersSetting.Show
    Case Else
    End Select
End Sub
