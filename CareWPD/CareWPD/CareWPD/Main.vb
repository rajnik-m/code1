Module Main
  <System.STAThread()> _
  Public Sub Main()
    Application.EnableVisualStyles()
    Application.DoEvents()

    Settings.Save = New Settings.SaveSettingsDelegate(AddressOf SaveSettings)
    Settings.Upgrade = New Settings.UpgradeSettingsDelegate(AddressOf UpgradeSettings)
    GetMySettings()
    If My.Application.IsNetworkDeployed Then
      AppValues.Init(My.Application.CommandLineArgs, My.Application.IsNetworkDeployed, My.Application.Deployment.UpdateLocation)
    Else
      AppValues.Init(My.Application.CommandLineArgs, False, Nothing)
    End If
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psConnecting)
    If DataHelper.CheckVersion(DataHelper.CareServiceTypes.WPDWebServices) Then
      DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
      Dim vRun As Boolean = False
      If DataHelper.AuthenticatedUser.Length > 0 Then
        Try
          DataHelper.Login("WPD", DataHelper.AuthenticatedUser, "none", AppValues.Database, DataHelper.AuthenticatedUser)
          vRun = True
        Catch vEx As CareException
          If vEx.ErrorNumber <> CareException.ErrorNumbers.enLoginFailed Then
            Throw
          End If
        End Try
      End If
      If Not vRun Then
        Dim vForm As New frmLogin("WPD")
        If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then vRun = True
        vForm = Nothing
      End If
      If vRun Then
        DataHelper.ShowProgress(frmProgress.ProgressStatuses.psInitialising)
        Try
          'CheckWebControls is now done at login time on the server
          Dim vDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtWebSites, New ParameterList(True))
          If vDataSet.Tables.Count > 0 Then
            Dim vSelectForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litWebSites)
            DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
            Dim vDialogResult As DialogResult = vSelectForm.ShowDialog
            If vDialogResult = System.Windows.Forms.DialogResult.OK Then
              Dim vWebNumber As Integer = IntegerValue(DataHelper.GetTableFromDataSet(vDataSet).Rows(vSelectForm.SelectedRow).Item("WebNumber").ToString)
              Application.Run(New frmMain(vWebNumber))
            ElseIf vDialogResult = DialogResult.No Then
              Application.Run(New frmMain(0))
            End If
          Else
            Application.Run(New frmMain(0))
          End If
        Catch vEx As Exception
          DataHelper.HandleException(vEx)
        End Try
      End If
    End If
  End Sub

  Private Sub GetMySettings()
    Settings.AllowDatabaseSelection = My.Settings.AllowDatabaseSelection
    Settings.DATABASE = My.Settings.DATABASE
    Settings.WebServiceTimeout = My.Settings.WebServiceTimeout
    Settings.ConfirmDelete = My.Settings.ConfirmDelete
  End Sub

  Private Sub SaveSettings()
    My.Settings.Save()
  End Sub
  Private Sub UpgradeSettings()
    My.Settings.Upgrade()
  End Sub

End Module
