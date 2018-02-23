Public Class frmPreferences

  Private Const MIN_WS_TIMEOUT As Integer = 20
  Private Const MAX_WS_TIMEOUT As Integer = 3600

  Private mvAllowAppearanceSettings As Boolean
  Private mvAllowFontSettings As Boolean

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    Try
      SetControlTheme()
      MainHelper.SetMDIParent(Me)
      mvAllowAppearanceSettings = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciPreferencesModifyAppearance)
      mvAllowFontSettings = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciPreferencesModifyFonts)
      Me.tabNotification.Text = ControlText.TbpNotification
      Me.lblPollingInterval.Text = ControlText.LblPollingInterval
      Me.chkNotifyMeetings.Text = ControlText.LblNotifyMeetings
      Me.chkNotifyDeadlines.Text = ControlText.LblNotifyDeadline
      Me.chkNotifyDocuments.Text = ControlText.LblNotifyDocuments
      Me.chkNotifyActions.Text = ControlText.LblNotifyActions
      Me.tabDisplay.Text = ControlText.TbpDisplay
      Me.chkPlainEditPanel.Text = ControlText.LblPlainBackground
      Me.lblBackgroundImageLayout.Text = ControlText.LblBackgroundImageLayout
      Me.lblBackgroundImage.Text = ControlText.LblBackgroundImage
      Me.tabGeneral.Text = ControlText.TbpGeneral
      Me.lblWebServicesTimeout.Text = ControlText.LblWebServiceTimeout
      Me.lblHistoryDays.Text = ControlText.LblDaysToKeepHistory
      Me.tabConfirmation.Text = ControlText.TbpConfirmation
      Me.chkConfirmCancel.Text = ControlText.LblConfirmCancel
      Me.chkConfirmDelete.Text = ControlText.LblConfirmDelete
      Me.chkConfirmInsert.Text = ControlText.LblConfirmInsert
      Me.chkConfirmUpdate.Text = ControlText.LblConfirmUpdate
      Me.chkFinderResultsMsgBox.text = ControlText.lblFinderResultsMsgBox

      Me.chkHideHistoricNetwork.Text = ControlText.lblHideHistoricNetwork

      Me.Text = ControlText.FrmUserPreferences

      Dim vLayout() As LookupItem = { _
      New LookupItem(CStr(ImageLayout.None), [Enum].GetName(GetType(ImageLayout), ImageLayout.None)), _
      New LookupItem(CStr(ImageLayout.Center), [Enum].GetName(GetType(ImageLayout), ImageLayout.Center)), _
      New LookupItem(CStr(ImageLayout.Stretch), [Enum].GetName(GetType(ImageLayout), ImageLayout.Stretch)), _
      New LookupItem(CStr(ImageLayout.Tile), [Enum].GetName(GetType(ImageLayout), ImageLayout.Tile)), _
      New LookupItem(CStr(ImageLayout.Zoom), [Enum].GetName(GetType(ImageLayout), ImageLayout.Zoom))}
      cboBackgroundImageLayout.DisplayMember = "LookupDesc"
      cboBackgroundImageLayout.ValueMember = "LookupCode"
      cboBackgroundImageLayout.DataSource = vLayout
      AddHandler txtPollingInterval.KeyPress, AddressOf Utilities.NumericKeyPressHandler
      AddHandler txtHistoryDays.KeyPress, AddressOf Utilities.NumericKeyPressHandler
      AddHandler txtWebServicesTimeout.KeyPress, AddressOf Utilities.NumericKeyPressHandler
      AddHandler txtTaskPollingInterval.KeyPress, AddressOf Utilities.NumericKeyPressHandler

      chkNotifyActions.Checked = Settings.NotifyActions
      chkNotifyDocuments.Checked = Settings.NotifyDocuments
      chkNotifyDeadlines.Checked = Settings.NotifyDeadlines
      chkNotifyMeetings.Checked = Settings.NotifyMeetings
      txtPollingInterval.Text = Settings.NotificationPollingMinutes.ToString
      txtTaskPollingInterval.Text = Settings.TaskNotificationPollingSeconds.ToString
      txtHistoryDays.Text = Settings.HistoryDays.ToString
      txtWebServicesTimeout.Text = Settings.WebServiceTimeout.ToString
      txtBackgroundImage.Text = Settings.BackGroundImage
      cboBackgroundImageLayout.SelectedValue = CStr(Settings.BackgroundImageLayout)
      chkPlainEditPanel.Checked = Settings.PlainEditPanel

      chkConfirmCancel.Checked = Settings.ConfirmCancel
      chkConfirmInsert.Checked = Settings.ConfirmInsert
      chkConfirmUpdate.Checked = Settings.ConfirmUpdate
      chkConfirmDelete.Checked = Settings.ConfirmDelete
      chkFinderResultsMsgBox.Checked = Settings.FinderResultsMsgBox

      chkTabIntoDisplayPanel.Checked = Settings.TabIntoDisplayPanel
      chkTabIntoHeaderPanel.Checked = Settings.TabIntoHeaderPanel
      chkErrorsAsMsgbox.Checked = Settings.ShowErrorsAsMsgBox
      chkHideHistoricNetwork.Checked = Settings.HideHistoricNetwork
      If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_dashboard) AndAlso AppValues.IsDashboardLicensed AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDashboardView) Then
        chkDisplayDashboardAtLogin.Checked = Settings.DisplayDashboardAtLogin
      Else
        chkDisplayDashboardAtLogin.Checked = False
        chkDisplayDashboardAtLogin.Enabled = False
      End If

      cmdApply.Visible = False
      cmdSaveAs.Visible = False

      'Save the default schemes into the DB
      DisplayTheme.SaveSchemeToDataBase()

      GetThemes(Settings.SchemeID)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GetThemes(ByVal vSchemeId As Integer)
    Dim vParameters As New ParameterList(True)
    vParameters("Default") = "Y"
    Dim vThemes As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtUserSchemes, vParameters)
    If vThemes Is Nothing Then
      cboSchemes.DataSource = Nothing
      cboSchemes.Enabled = False
    Else
      cboSchemes.Enabled = True
      cboSchemes.DisplayMember = "UserSchemeDesc"
      cboSchemes.ValueMember = "UserSchemeId"
      cboSchemes.DataSource = vThemes
      If vSchemeId > 0 Then
        SelectComboBoxItem(cboSchemes, vSchemeId.ToString)
      Else
        cboSchemes.SelectedIndex = 0
      End If
    End If
  End Sub

  Private Function SetValues() As Boolean
    Dim vInValid As Boolean
    Dim vWSTimeout As Integer = IntegerValue(txtWebServicesTimeout.Text)
    If vWSTimeout < MIN_WS_TIMEOUT Or vWSTimeout > MAX_WS_TIMEOUT Then
      ShowInformationMessage(InformationMessages.ImWEBServicesTimeoutInvalid, MIN_WS_TIMEOUT.ToString, MAX_WS_TIMEOUT.ToString)
      vInValid = True
    End If
    If Not vInValid Then
      Settings.NotifyActions = chkNotifyActions.Checked
      Settings.NotifyDocuments = chkNotifyDocuments.Checked
      Settings.NotifyDeadlines = chkNotifyDeadlines.Checked
      Settings.NotifyMeetings = chkNotifyMeetings.Checked

      Settings.ConfirmCancel = chkConfirmCancel.Checked
      Settings.ConfirmInsert = chkConfirmInsert.Checked
      Settings.ConfirmUpdate = chkConfirmUpdate.Checked
      Settings.ConfirmDelete = chkConfirmDelete.Checked

      Dim vMinutes As Integer = IntegerValue(txtPollingInterval.Text)
      If vMinutes <> Settings.NotificationPollingMinutes Then
        Settings.NotificationPollingMinutes = vMinutes
        MainHelper.SetNotificationTime()
      End If

      Dim vSeconds As Integer = IntegerValue(txtTaskPollingInterval.Text)
      If vSeconds <> Settings.TaskNotificationPollingSeconds Then
        Settings.TaskNotificationPollingSeconds = vSeconds
        MainHelper.SetTaskNotificationTimer()
      End If

      Settings.HistoryDays = IntegerValue(txtHistoryDays.Text)
      DataHelper.WebServicesTimeout = IntegerValue(txtWebServicesTimeout.Text)
      Settings.BackGroundImage = txtBackgroundImage.Text
      Settings.BackgroundImageLayout = CType(cboBackgroundImageLayout.SelectedValue, ImageLayout)
      MainHelper.SetBackgroundImage(Settings.BackGroundImage, Settings.BackgroundImageLayout)

      Settings.PlainEditPanel = chkPlainEditPanel.Checked

      Settings.TabIntoDisplayPanel = chkTabIntoDisplayPanel.Checked
      Settings.TabIntoHeaderPanel = chkTabIntoHeaderPanel.Checked
      Settings.ShowErrorsAsMsgBox = chkErrorsAsMsgbox.Checked
      Settings.HideHistoricNetwork = chkHideHistoricNetwork.Checked
      Settings.FinderResultsMsgBox = chkFinderResultsMsgBox.Checked
      If chkDisplayDashboardAtLogin.Enabled Then Settings.DisplayDashboardAtLogin = chkDisplayDashboardAtLogin.Checked 'Only Save this if it was enabled

      If cboSchemes.DataSource IsNot Nothing Then
        Settings.SchemeID = CInt(cboSchemes.SelectedValue)
      End If

      SaveAppearanceSettings()
      DisplayTheme.PlainEditPanelTheme = Settings.PlainEditPanel
      SaveFontSettings()
      MainHelper.SetControlColors()
      tim.Interval = 100
      tim.Enabled = True
      Return True
    End If
  End Function

  Private Sub SaveAppearanceSettings()
    If mvAllowAppearanceSettings Then

      DisplayTheme.MatchesThemeSettings = False
    End If
  End Sub

  Private Sub SaveFontSettings()
    If mvAllowFontSettings Then
      DisplayTheme.MatchesThemeSettings = False
    End If
  End Sub

  Private Sub SetDefaults()
    chkNotifyActions.Checked = Settings.NotifyActions
    chkNotifyDocuments.Checked = Settings.NotifyDocuments
    chkNotifyDeadlines.Checked = Settings.NotifyDeadlines
    chkNotifyMeetings.Checked = Settings.NotifyMeetings

    chkConfirmCancel.Checked = Settings.ConfirmCancel
    chkConfirmInsert.Checked = Settings.ConfirmInsert
    chkConfirmUpdate.Checked = Settings.ConfirmUpdate
    chkConfirmDelete.Checked = Settings.ConfirmDelete

    txtPollingInterval.Text = Settings.NotificationPollingMinutes.ToString
    txtTaskPollingInterval.Text = Settings.TaskNotificationPollingSeconds.ToString
    txtBackgroundImage.Text = ""
    txtHistoryDays.Text = Settings.HistoryDays.ToString
    txtWebServicesTimeout.Text = Settings.WebServiceTimeout.ToString
    cboBackgroundImageLayout.SelectedValue = CStr(ImageLayout.None)
    chkPlainEditPanel.Checked = False
    chkTabIntoDisplayPanel.Checked = False
    chkTabIntoHeaderPanel.Checked = False
    chkErrorsAsMsgbox.Checked = False
    chkHideHistoricNetwork.Checked = False
    chkDisplayDashboardAtLogin.Checked = False
    chkFinderResultsMsgBox.Checked = False
    DisplayTheme.MatchesThemeSettings = False
    If cboSchemes.DataSource IsNot Nothing Then cboSchemes.SelectedIndex = 0
    SetDefaultAppearance()
    'SetDefaultFonts()
  End Sub

  Private Sub SetDefaultAppearance()
    Dim vDT As New DisplayThemeSettings   'Save the current display theme settings
    DisplayTheme.Init()                   'Initialising the DisplayTheme will reset it to default values
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    If SetValues() Then
      DialogResult = System.Windows.Forms.DialogResult.OK
      Me.Close()
    Else
      DialogResult = System.Windows.Forms.DialogResult.None
    End If
  End Sub

  Private Sub cmdDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDefaults.Click
    SetDefaults()
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdBackgroundImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBackgroundImage.Click
    With cmd
      .Title = "Get Background Image"
      .Filter = "Bitmap Files (*.bmp)|*.bmp|JPEG Files (*.jpg)|*.jpg"

      If txtBackgroundImage.Text.Length > 0 Then
        Dim vInfo As New System.IO.FileInfo(txtBackgroundImage.Text)
        Select Case vInfo.Extension
          Case ".bmp"
            .FilterIndex = 1
          Case ".jpg"
            .FilterIndex = 2
        End Select
      Else
        .FilterIndex = 2
        .DefaultExt = "jpg"
      End If
      .FileName = txtBackgroundImage.Text
      If .ShowDialog = System.Windows.Forms.DialogResult.OK Then txtBackgroundImage.Text = .FileName
    End With
  End Sub

  Private Sub cmdApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdApply.Click
    Try
      SetValues()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Protected Overloads Sub frmCardMaintenance_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SystemColorsChanged
    'csFormBackColor.Color = DisplayTheme.FormBackColor
    'csSplitterBackColor.Color = DisplayTheme.SplitterBackColor
    'csGridBackColor.Color = DisplayTheme.GridBackAreaColor
    'csPanelBackColor.Color = DisplayTheme.SelectionPanelTreeBackColor
    'csButtonPanelBackColor.Color = DisplayTheme.ButtonPanelBackColor
    'csGridHyperlinkColor.Color = DisplayTheme.GridHyperlinkColor
    'pteDisplayPanel.InitFromPanelTheme(DisplayTheme.DisplayPanelTheme)
    'pteEditPanel.InitFromPanelTheme(DisplayTheme.EditPanelTheme)
    'pteSelectionPanel.InitFromPanelTheme(DisplayTheme.SelectionPanelTheme)
    'pteDisplayLabel.InitFromPanelTheme(DisplayTheme.DisplayLabelTheme)
    'pteDisplayData.InitFromPanelTheme(DisplayTheme.DisplayDataTheme)
    'pteDashboardHeading.InitFromPanelTheme(DisplayTheme.DashboardHeadingTheme)
    'pteToolbar.InitFromPanelTheme(DisplayTheme.ToolbarTheme)
  End Sub

  Private Sub tab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
    'If tab.SelectedTab Is tabAppearance Or tab.SelectedTab Is tabFont Then
    '  cmdSaveAs.Enabled = True
    'Else
    '  cmdSaveAs.Enabled = False
    'End If
  End Sub

  'Private Sub cmdSaveAs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSaveAs.Click
  '  Try
  '    Dim vSaveFontTheme As Boolean = False
  '    If tab.SelectedTab Is tabFont Then vSaveFontTheme = True

  '    Dim vDefaults As New ParameterList
  '    If vSaveFontTheme Then
  '      If cboFonts.SelectedIndex > 0 Then
  '        vDefaults("ThemeName") = DirectCast(cboFonts.SelectedItem, DataRowView).Item("ItemDesc").ToString()
  '      Else
  '        vDefaults("ThemeName") = "New Font Theme"
  '      End If
  '    Else
  '      If cboAppearance.SelectedIndex > 0 Then
  '        vDefaults("ThemeName") = DirectCast(cboAppearance.SelectedItem, DataRowView).Item("ItemDesc").ToString
  '      Else
  '        vDefaults("ThemeName") = "New Appearance Theme"
  '      End If
  '    End If
  '    Dim vList As ParameterList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optThemeName, Nothing, vDefaults, "Save Theme As")
  '    If vList IsNot Nothing Then
  '      Dim vMS As MemoryStream = Nothing
  '      Try
  '        vMS = New MemoryStream()
  '        If vSaveFontTheme Then
  '          Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
  '          Dim vSettings As New FontThemeSettings
  '          vSettings.FormFont = fsForm.SelectedFont
  '          vSettings.GridFont = fsGrid.SelectedFont
  '          vSettings.SelectionPanelFont = fsSelectionPanel.SelectedFont
  '          vSettings.NavigationPanelFont = fsNavigationPanel.SelectedFont
  '          vSettings.DisplayLabelFont = fsDisplayLabel.SelectedFont
  '          vSettings.DisplayItemFont = fsDisplayItem.SelectedFont
  '          vSettings.DashboardHeadingFont = fsDashboardHeading.SelectedFont
  '          vXS.Serialize(vMS, vSettings)
  '          Dim vXMLList As New ParameterList(True)
  '          vXMLList("XmlDataType") = "FT"               'Font Theme
  '          vXMLList("ItemDesc") = vList("ThemeName")    'Theme Name
  '          Dim vReader As New System.IO.StreamReader(vMS)
  '          vMS.Position = 0
  '          vXMLList("ItemXml") = vReader.ReadToEnd()
  '          Dim vReturnList As ParameterList = DataHelper.AddXMLDataItem(vXMLList)
  '          GetFontThemes(vReturnList.IntegerValue("XmlItemNumber"))
  '        Else
  '          Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
  '          Dim vAppearance As New DisplayThemeSettings
  '          vAppearance.FormBackColor = csFormBackColor.Color
  '          vAppearance.SplitterBackColor = csSplitterBackColor.Color
  '          vAppearance.GridBackAreaColor = csGridBackColor.Color
  '          vAppearance.SelectionPanelTreeBackColor = csPanelBackColor.Color
  '          vAppearance.ButtonPanelBackColor = csButtonPanelBackColor.Color
  '          vAppearance.GridHyperlinkColor = csGridHyperlinkColor.Color
  '          vAppearance.HeaderBackgroundSameAsForm = chkHeaderBackgroundSameAsForm.Checked
  '          vAppearance.UnderlineHyperlinks = chkUnderlineHyperlinks.Checked
  '          pteDisplayPanel.UpdatePanelTheme(vAppearance.DisplayPanelThemeSettings)
  '          pteSelectionPanel.UpdatePanelTheme(vAppearance.SelectionPanelThemeSettings)
  '          pteEditPanel.UpdatePanelTheme(vAppearance.EditPanelThemeSettings)
  '          pteDisplayLabel.UpdatePanelTheme(vAppearance.DisplayLabelThemeSettings)
  '          pteDisplayData.UpdatePanelTheme(vAppearance.DisplayDataThemeSettings)
  '          pteDashboardHeading.UpdatePanelTheme(vAppearance.DashboardHeadingThemeSettings)
  '          pteToolbar.UpdatePanelTheme(vAppearance.ToolbarThemeSettings)
  '          vXS.Serialize(vMS, vAppearance)
  '          Dim vXMLList As New ParameterList(True)
  '          vXMLList("XmlDataType") = "AT"               'Appearance Theme
  '          vXMLList("ItemDesc") = vList("ThemeName")    'Theme Name
  '          Dim vReader As New System.IO.StreamReader(vMS)
  '          vMS.Position = 0
  '          vXMLList("ItemXml") = vReader.ReadToEnd()
  '          Dim vReturnList As ParameterList = DataHelper.AddXMLDataItem(vXMLList)
  '          GetAppearanceThemes(vReturnList.IntegerValue("XmlItemNumber"))
  '        End If
  '        vMS.Close()
  '      Catch vEx As Exception
  '        DataHelper.HandleException(vEx)
  '      Finally
  '        If vMS IsNot Nothing Then vMS.Close()
  '      End Try
  '    End If
  '  Catch vEx As Exception
  '    DataHelper.HandleException(vEx)
  '  End Try
  'End Sub

  

  'Private Sub SetFont(ByVal vSelectedSchemeRow As DataRowView)
  '  Try
  '    Dim vXML As String = vSelectedSchemeRow.Item("ItemXml").ToString
  '    Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
  '    Dim vFontSettings As FontThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), FontThemeSettings)
  '    DisplayTheme.FormFont = vFontSettings.FormFont
  '    DisplayTheme.GridFont = vFontSettings.GridFont
  '    DisplayTheme.SelectionPanelFont = vFontSettings.SelectionPanelFont
  '    DisplayTheme.NavigationPanelFont = vFontSettings.NavigationPanelFont
  '    DisplayTheme.DisplayLabelFont = vFontSettings.DisplayLabelFont
  '    DisplayTheme.DisplayItemFont = vFontSettings.DisplayItemFont
  '    DisplayTheme.DashboardHeadingFont = vFontSettings.DashboardHeadingFont
  '    DisplayTheme.MatchesThemeSettings = False
  '  Catch vEx As Exception
  '    DataHelper.HandleException(vEx)
  '  End Try
  'End Sub

  'Private Sub SetAppearance(ByVal vSelectedSchemeRow As DataRowView)
  '  Try
  '    Dim vRow As DataRowView = DirectCast(cboAppearance.SelectedItem, DataRowView)
  '    Dim vXML As String = vRow.Item("ItemXml").ToString
  '    Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
  '    Dim vAppearance As DisplayThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), DisplayThemeSettings)
  '    csFormBackColor.Color = vAppearance.FormBackColor
  '    csSplitterBackColor.Color = vAppearance.SplitterBackColor
  '    csGridBackColor.Color = vAppearance.GridBackAreaColor
  '    csPanelBackColor.Color = vAppearance.SelectionPanelTreeBackColor
  '    csButtonPanelBackColor.Color = vAppearance.ButtonPanelBackColor
  '    csGridHyperlinkColor.Color = vAppearance.GridHyperlinkColor
  '    chkHeaderBackgroundSameAsForm.Checked = vAppearance.HeaderBackgroundSameAsForm
  '    chkUnderlineHyperlinks.Checked = vAppearance.UnderlineHyperlinks
  '    pteDisplayPanel.InitFromPanelTheme(vAppearance.DisplayPanelThemeSettings)
  '    pteSelectionPanel.InitFromPanelTheme(vAppearance.SelectionPanelThemeSettings)
  '    pteEditPanel.InitFromPanelTheme(vAppearance.EditPanelThemeSettings)
  '    pteDisplayLabel.InitFromPanelTheme(vAppearance.DisplayLabelThemeSettings)
  '    pteDisplayData.InitFromPanelTheme(vAppearance.DisplayDataThemeSettings)
  '    pteDashboardHeading.InitFromPanelTheme(vAppearance.DashboardHeadingThemeSettings)
  '    pteToolbar.InitFromPanelTheme(vAppearance.ToolbarThemeSettings)

  '    DisplayTheme.FormBackColor = csFormBackColor.Color
  '    DisplayTheme.SplitterBackColor = csSplitterBackColor.Color
  '    DisplayTheme.GridBackAreaColor = csGridBackColor.Color
  '    DisplayTheme.SelectionPanelTreeBackColor = csPanelBackColor.Color
  '    DisplayTheme.ButtonPanelBackColor = csButtonPanelBackColor.Color
  '    DisplayTheme.HeaderBackgroundSameAsForm = chkHeaderBackgroundSameAsForm.Checked
  '    pteDisplayPanel.UpdatePanelTheme(DisplayTheme.DisplayPanelTheme)
  '    pteSelectionPanel.UpdatePanelTheme(DisplayTheme.SelectionPanelTheme)
  '    pteEditPanel.UpdatePanelTheme(DisplayTheme.EditPanelTheme)
  '    pteDisplayLabel.UpdatePanelTheme(DisplayTheme.DisplayLabelTheme)
  '    pteDisplayData.UpdatePanelTheme(DisplayTheme.DisplayDataTheme)
  '    pteDashboardHeading.UpdatePanelTheme(DisplayTheme.DashboardHeadingTheme)
  '    DisplayTheme.GridHyperlinkColor = csGridHyperlinkColor.Color
  '    DisplayTheme.UnderlineHyperlinks = chkUnderlineHyperlinks.Checked
  '    pteToolbar.UpdatePanelTheme(DisplayTheme.ToolbarTheme)
  '    DisplayTheme.MatchesThemeSettings = False

  '  Catch vEx As Exception
  '    DataHelper.HandleException(vEx)
  '  End Try
  'End Sub

  Private Sub cboAppearance_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    'If DisplayTheme.MatchesThemeSettings OrElse Me.ActiveControl Is sender Then
    '  If cboAppearance.SelectedIndex > 0 Then
    '    Try
    '      Dim vRow As DataRowView = DirectCast(cboAppearance.SelectedItem, DataRowView)
    '      Dim vXML As String = vRow.Item("ItemXml").ToString
    '      Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
    '      Dim vAppearance As DisplayThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), DisplayThemeSettings)
    '      csFormBackColor.Color = vAppearance.FormBackColor
    '      csSplitterBackColor.Color = vAppearance.SplitterBackColor
    '      csGridBackColor.Color = vAppearance.GridBackAreaColor
    '      csPanelBackColor.Color = vAppearance.SelectionPanelTreeBackColor
    '      csButtonPanelBackColor.Color = vAppearance.ButtonPanelBackColor
    '      csGridHyperlinkColor.Color = vAppearance.GridHyperlinkColor
    '      chkHeaderBackgroundSameAsForm.Checked = vAppearance.HeaderBackgroundSameAsForm
    '      chkUnderlineHyperlinks.Checked = vAppearance.UnderlineHyperlinks
    '      pteDisplayPanel.InitFromPanelTheme(vAppearance.DisplayPanelThemeSettings)
    '      pteSelectionPanel.InitFromPanelTheme(vAppearance.SelectionPanelThemeSettings)
    '      pteEditPanel.InitFromPanelTheme(vAppearance.EditPanelThemeSettings)
    '      pteDisplayLabel.InitFromPanelTheme(vAppearance.DisplayLabelThemeSettings)
    '      pteDisplayData.InitFromPanelTheme(vAppearance.DisplayDataThemeSettings)
    '      pteDashboardHeading.InitFromPanelTheme(vAppearance.DashboardHeadingThemeSettings)
    '      pteToolbar.InitFromPanelTheme(vAppearance.ToolbarThemeSettings)
    '    Catch vEx As Exception
    '      DataHelper.HandleException(vEx)
    '    End Try
    '  ElseIf cboAppearance.SelectedIndex = 0 Then
    '    SetDefaultAppearance()
    '  End If
    'End If
  End Sub

  Private Sub tim_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tim.Tick
    Me.Refresh()
    tim.Enabled = False
  End Sub

  Private Sub cmdDesign_Click(sender As Object, e As EventArgs) Handles cmdDesign.Click
    Me.Close()
    Dim vFrmThemes As New frmThemes()
    vFrmThemes.Show()
  End Sub

End Class
