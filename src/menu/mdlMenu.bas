Attribute VB_Name = "mdlMenu"
Option Explicit

Public TIR_Type_Text As String, strCaption As String

Private Const CNST_FormCaption As String = "の設定"
Private Const CNST_TIR_Prefix As String = "TIR-"
    
#If False Then
    Private Const CNST_IHS_Login_URL_DefaultValue As String = "file://ycab5s21/workspc/MITAC_HP_Contents/MRJ_Link_071_ＩＨＳ（海外規約）/MITAC_ERC_201408.html"
#End If

Sub DM検索ツールのバージョン情報を表示する(ByRef myButton As IRibbonControl)
    MsgBox GetAppCNST(Version) & vbCrLf & "最終更新日： " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppNameDMFolderOpen)
End Sub

Sub TIR検索ツールのバージョン情報を表示する(ByRef myButton As IRibbonControl)
    MsgBox GetAppCNST(Version) & vbCrLf & "最終更新日： " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppNameTIRSearch)
End Sub

Sub TIRツールのバージョン情報を表示する(ByRef myButton As IRibbonControl)
    Call ShowVersion
End Sub

Private Sub ShowVersion()
    MsgBox GetAppCNST(Version) & vbCrLf & "最終更新日： " & GetAppCNST(LastModified), vbInformation + vbOKOnly, GetAppCNST(AppName)
End Sub


Sub ToolsTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIRTopフォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRTopFolder(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsTIRIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersIntegrationFileチェックを起動する(ByRef myButton As IRibbonControl)
    Call CheckIntegrationFile(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIRIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIRIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIRIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIRIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIRIntegrationFile保存フォルダを開く(ByRef myButton As IRibbonControl)
    Call OpenIntegrationFileFolder(ESRDFileCategory.CircuitBreakers)
End Sub


Sub ToolsTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIR送信メールを作成(ByRef myButton As IRibbonControl)
    Call CreateNotesEMailDraft(ESRDFileCategory.CircuitBreakers)
End Sub

Sub ToolsTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Tools)
End Sub

Sub SuppliesTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub ToolsSupplies最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.SUPPLIES)
End Sub

Sub EnterpriseTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Enterprise)
End Sub

Sub ZonesTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.Zones)
End Sub

Sub AccessPointsTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.AccessPoints)
End Sub

Sub CircuitBreakersTIR最新エクセルを開く(ByRef myButton As IRibbonControl)
    Call OpenTIRXlsFile(ESRDFileCategory.CircuitBreakers)
End Sub

'Sub TIRToolsの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Tools)
'End Sub
'
'Sub TIRSuppliesの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Supplies)
'End Sub
'
'Sub TIREnterpriseの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Enterprise)
'End Sub
'
'Sub TIRZonesの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.Zones)
'End Sub
'
'Sub TIRAccessPointsの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.AccessPoints)
'End Sub
'
'Sub TIRCircuitBreakersの検索画面を表示する(ByRef myButton As IRibbonControl)
'    Call SearchTIRItem(ESRDFileCategory.CircuitBreakers)
'End Sub


Sub TIRToolsの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Tools)
End Sub

Sub TIRSuppliesの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.SUPPLIES)
End Sub

Sub TIREnterpriseの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Enterprise)
End Sub

Sub TIRZonesの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.Zones)
End Sub

Sub TIRAccessPointsの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.AccessPoints)
End Sub

Sub TIRCircuitBreakersの設定画面を表示する(ByRef myButton As IRibbonControl)
    Call UpdateCaptionAndShowForm(ESRDFileCategory.CircuitBreakers)
End Sub

Sub TIR一般設定画面を表示する(ByRef myButton As IRibbonControl)
    frmCommonSettings.Show
End Sub

Sub DM検証フォルダの設定画面を表示する(ByRef myButton As IRibbonControl)
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
    iRet = MsgBox("設定されているフォルダ・ファイルが見つかりません." & vbCrLf & _
                "フォルダ・ファイルの設定を行いますか?", vbExclamation + vbYesNoCancel, GetAppCNST(AppName))
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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
                    MsgBox "チェック対象のIntegrationFileがありません.", vbExclamation + vbOKOnly, "Integration Fileチェック"
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

Sub IHSへログインする(ByRef myButton As IRibbonControl)
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
    Application.StatusBar = ("エラー!： " & Err.Number & ":" & Err.Description)
    MsgBox Err.Number & ":" & Err.Description
End Sub

Sub EnterpriseTIR用DBを開く(ByRef myButton As IRibbonControl)
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
