Public Class frmCheckSetup

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optCheckSetup))

    'Set initial values
    epl.SetValue("CheckControlNumbers", "Y")
    epl.SetValue("CheckControlTables", "Y")
    epl.SetValue("CheckTraderFinancialControlData", "Y")
    epl.SetValue("CheckTelephoneNumbers", "Y")
    epl.SetValue("CheckCurrencyTables", "Y")
    epl.SetValue("CheckPPSchedulesExist", "Y")
    epl.SetValue("CheckCCDDClaimDates", "Y")
    epl.SetValue("CheckNominalAccountCodes", "Y")
    epl.SetValue("DropTempTables", "N")
    epl.SetValue("CheckDuplicateGroupValues", "Y")
  End Sub

  Private Sub frmCheckSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      'Set default path for log file
      epl.SetValue("ErrorLogFile", Application.StartupPath & "\chksetup.log")

      If AppValues.ControlValue(AppValues.ControlValues.auto_pay_claim_date_method) = "N" Then
        'Next Payment Due dates will be used so no claim dates to check
        epl.EnableControl("CheckCCDDClaimDates", False)
        epl.SetValue("CheckCCDDClaimDates", "N")
      End If

      If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.nominal_account_validation) Then
        epl.EnableControl("CheckNominalAccountCodes", False)
        epl.SetValue("CheckNominalAccountCodes", "N")
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Try
      Dim vList As New ParameterList(True)
      epl.AddValuesToList(vList)
      vList.Remove("ErrorLogFile") 'Log file path not reqd as web service creates it own temp file
      Dim vFileName As String = epl.GetValue("ErrorLogFile")
      vFileName = DataHelper.CheckSetup(vList, vFileName)
      ShowInformationMessage(InformationMessages.ImCheckSetupComplete)
      Process.Start(vFileName)
      Me.Close()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub eplExport_ButtonClicked(ByVal sender As Object, ByVal pParameterName As String) Handles epl.ButtonClicked
    Dim vDLG As New SaveFileDialog

    With vDLG
      .Title = ControlText.LblErrorLogFile      'Error Log File
      .Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
      .DefaultExt = "log"
      .OverwritePrompt = True
      .FileName = epl.FindTextBox("ErrorLogFile").Text
      If .ShowDialog = System.Windows.Forms.DialogResult.OK Then epl.SetValue("ErrorLogFile", .FileName)
    End With
  End Sub

End Class