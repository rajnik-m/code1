Imports System.Xml
Imports System.Xml.Linq

Public Class frmThemes

  Private Const MIN_WS_TIMEOUT As Integer = 20
  Private Const MAX_WS_TIMEOUT As Integer = 3600
  Private Const DEFAULT_SCHEME As String = "CDBNETCL.BlueScheme.xml"

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

      Me.Text = ControlText.FrmUserThemes

      Dim vLayout() As LookupItem = { _
      New LookupItem(CStr(ImageLayout.None), [Enum].GetName(GetType(ImageLayout), ImageLayout.None)), _
      New LookupItem(CStr(ImageLayout.Center), [Enum].GetName(GetType(ImageLayout), ImageLayout.Center)), _
      New LookupItem(CStr(ImageLayout.Stretch), [Enum].GetName(GetType(ImageLayout), ImageLayout.Stretch)), _
      New LookupItem(CStr(ImageLayout.Tile), [Enum].GetName(GetType(ImageLayout), ImageLayout.Tile)), _
      New LookupItem(CStr(ImageLayout.Zoom), [Enum].GetName(GetType(ImageLayout), ImageLayout.Zoom))}

      'Make the color selectors share the same color dialog
      csFormBackColor.ColorDialog = cdlg
      csSplitterBackColor.ColorDialog = cdlg
      csGridBackColor.ColorDialog = cdlg
      csPanelBackColor.ColorDialog = cdlg
      csButtonPanelBackColor.ColorDialog = cdlg
      csGridHyperlinkColor.ColorDialog = cdlg
      pteDisplayPanel.ColorDialog = cdlg
      pteEditPanel.ColorDialog = cdlg
      pteSelectionPanel.ColorDialog = cdlg
      pteDisplayLabel.ColorDialog = cdlg
      pteDisplayData.ColorDialog = cdlg
      pteDashboardHeading.ColorDialog = cdlg
      pteToolbar.ColorDialog = cdlg

      If mvAllowAppearanceSettings Then
        csFormBackColor.Color = DisplayTheme.FormBackColor
        csSplitterBackColor.Color = DisplayTheme.SplitterBackColor
        csGridBackColor.Color = DisplayTheme.GridBackAreaColor
        csPanelBackColor.Color = DisplayTheme.SelectionPanelTreeBackColor
        csButtonPanelBackColor.Color = DisplayTheme.ButtonPanelBackColor
        csGridHyperlinkColor.Color = DisplayTheme.GridHyperlinkColor
        chkHeaderBackgroundSameAsForm.Checked = DisplayTheme.HeaderBackgroundSameAsForm
        chkUnderlineHyperlinks.Checked = DisplayTheme.UnderlineHyperlinks
        pteDisplayPanel.InitFromPanelTheme(DisplayTheme.DisplayPanelTheme)
        pteEditPanel.InitFromPanelTheme(DisplayTheme.EditPanelTheme)
        pteSelectionPanel.InitFromPanelTheme(DisplayTheme.SelectionPanelTheme)
        pteDisplayLabel.InitFromPanelTheme(DisplayTheme.DisplayLabelTheme)
        pteDisplayData.InitFromPanelTheme(DisplayTheme.DisplayDataTheme)
        pteDashboardHeading.InitFromPanelTheme(DisplayTheme.DashboardHeadingTheme)
        pteToolbar.InitFromPanelTheme(DisplayTheme.ToolbarTheme)
      End If

      If mvAllowFontSettings Then
        fsForm.SelectedFont = DisplayTheme.FormFont
        fsGrid.SelectedFont = DisplayTheme.GridFont
        fsSelectionPanel.SelectedFont = DisplayTheme.SelectionPanelFont
        fsNavigationPanel.SelectedFont = DisplayTheme.NavigationPanelFont
        fsDisplayLabel.SelectedFont = DisplayTheme.DisplayLabelFont
        fsDisplayItem.SelectedFont = DisplayTheme.DisplayItemFont
        fsDashboardHeading.SelectedFont = DisplayTheme.DashboardHeadingFont
      End If

      If mvAllowAppearanceSettings = False AndAlso mvAllowFontSettings = False Then
        cmdApply.Visible = False
        cmdSaveAs.Visible = False
      End If
      GetFontThemes(Settings.FontThemeID)
      GetAppearanceThemes(Settings.AppearanceThemeID)
      GetSchemes()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GetAppearanceThemes(ByVal pItemNumber As Integer)
    Dim vList As New ParameterList(True)
    vList("XmlDataType") = "AT"   'Appearance themes
    Dim vAppearance As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vList)

    If vAppearance Is Nothing Then
      cboAppearance.DataSource = Nothing
      'cboAppearance.Enabled = False
    Else
      cboAppearance.Enabled = True
      cboAppearance.DisplayMember = "ItemDesc"
      cboAppearance.ValueMember = "XmlItemNumber"
      cboAppearance.DataSource = vAppearance
      If pItemNumber > 0 Then
        SelectComboBoxItem(cboAppearance, pItemNumber.ToString)
      Else
        cboAppearance.SelectedIndex = 0
      End If
    End If
  End Sub

  Private Sub GetSchemes()
    Dim vList As New ParameterList(True)
    Dim vSchemes As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtUserSchemes)

    If vSchemes Is Nothing Then
      cboSchemes.DataSource = Nothing
      ' cboSchemes.Enabled = False
    Else
      cboSchemes.Enabled = True
      cboSchemes.DisplayMember = "UserSchemeDesc"
      cboSchemes.ValueMember = "UserSchemeId"
      cboSchemes.DataSource = vSchemes
    End If
  End Sub

  Private Sub GetFontThemes(ByVal pItemNumber As Integer)
    Dim vList As New ParameterList(True)
    vList("XmlDataType") = "FT"   'Font themes
    Dim vFonts As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vList)
    If vFonts Is Nothing Then
      cboFonts.DataSource = Nothing
      ' cboFonts.Enabled = False
    Else
      cboFonts.Enabled = True
      cboFonts.DisplayMember = "ItemDesc"
      cboFonts.ValueMember = "XmlItemNumber"
      cboFonts.DataSource = vFonts
      If pItemNumber > 0 Then
        SelectComboBoxItem(cboFonts, pItemNumber.ToString)
      Else
        cboFonts.SelectedIndex = 0
      End If
    End If
  End Sub

  Private Function SetValues() As Boolean
    Dim vInValid As Boolean
    If Not vInValid Then
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
      DisplayTheme.FormBackColor = csFormBackColor.Color
      DisplayTheme.SplitterBackColor = csSplitterBackColor.Color
      DisplayTheme.GridBackAreaColor = csGridBackColor.Color
      DisplayTheme.SelectionPanelTreeBackColor = csPanelBackColor.Color
      DisplayTheme.ButtonPanelBackColor = csButtonPanelBackColor.Color
      DisplayTheme.HeaderBackgroundSameAsForm = chkHeaderBackgroundSameAsForm.Checked
      pteDisplayPanel.UpdatePanelTheme(DisplayTheme.DisplayPanelTheme)
      pteSelectionPanel.UpdatePanelTheme(DisplayTheme.SelectionPanelTheme)
      pteEditPanel.UpdatePanelTheme(DisplayTheme.EditPanelTheme)
      pteDisplayLabel.UpdatePanelTheme(DisplayTheme.DisplayLabelTheme)
      pteDisplayData.UpdatePanelTheme(DisplayTheme.DisplayDataTheme)
      pteDashboardHeading.UpdatePanelTheme(DisplayTheme.DashboardHeadingTheme)
      DisplayTheme.GridHyperlinkColor = csGridHyperlinkColor.Color
      DisplayTheme.UnderlineHyperlinks = chkUnderlineHyperlinks.Checked
      pteToolbar.UpdatePanelTheme(DisplayTheme.ToolbarTheme)
      DisplayTheme.MatchesThemeSettings = False
    End If
  End Sub

  Private Sub SaveFontSettings()
    If mvAllowFontSettings Then
      DisplayTheme.FormFont = fsForm.SelectedFont
      DisplayTheme.GridFont = fsGrid.SelectedFont
      DisplayTheme.SelectionPanelFont = fsSelectionPanel.SelectedFont
      DisplayTheme.NavigationPanelFont = fsNavigationPanel.SelectedFont
      DisplayTheme.DisplayLabelFont = fsDisplayLabel.SelectedFont
      DisplayTheme.DisplayItemFont = fsDisplayItem.SelectedFont
      DisplayTheme.DashboardHeadingFont = fsDashboardHeading.SelectedFont
      DisplayTheme.MatchesThemeSettings = False
    End If
  End Sub

  Private Sub SetDefaults()
    DisplayTheme.MatchesThemeSettings = False
    If cboSchemes.DataSource IsNot Nothing Then cboSchemes.SelectedIndex = 0
    cboAppearance.SelectedIndex = 0
    cboFonts.SelectedIndex = 0
    SetDefaultAppearance()
    SetDefaultFonts()
  End Sub

  Private Sub SetDefaultAppearance()
    Dim vDT As New DisplayThemeSettings   'Save the current display theme settings

    Dim vDefaultTheme As String = GetResourceTextFile(DEFAULT_SCHEME)

    Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
    Dim vAppearance As DisplayThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vDefaultTheme)), DisplayThemeSettings)
    DisplayTheme.Init(vAppearance)                   'Initialising the DisplayTheme will reset it to default values

    ''Set the form up from the default settings
    csFormBackColor.Color = DisplayTheme.FormBackColor
    csSplitterBackColor.Color = DisplayTheme.SplitterBackColor
    csGridBackColor.Color = DisplayTheme.GridBackAreaColor
    csPanelBackColor.Color = DisplayTheme.SelectionPanelTreeBackColor
    csButtonPanelBackColor.Color = DisplayTheme.ButtonPanelBackColor
    csGridHyperlinkColor.Color = DisplayTheme.GridHyperlinkColor
    chkHeaderBackgroundSameAsForm.Checked = DisplayTheme.HeaderBackgroundSameAsForm
    chkUnderlineHyperlinks.Checked = DisplayTheme.UnderlineHyperlinks
    'Initialise all the paneltheme editors from the themes
    pteDisplayPanel.InitFromPanelTheme(DisplayTheme.DisplayPanelTheme)
    pteEditPanel.InitFromPanelTheme(DisplayTheme.EditPanelTheme)
    pteSelectionPanel.InitFromPanelTheme(DisplayTheme.SelectionPanelTheme)
    pteDisplayLabel.InitFromPanelTheme(DisplayTheme.DisplayLabelTheme)
    pteDisplayData.InitFromPanelTheme(DisplayTheme.DisplayDataTheme)
    pteDashboardHeading.InitFromPanelTheme(DisplayTheme.DashboardHeadingTheme)
    pteToolbar.InitFromPanelTheme(DisplayTheme.ToolbarTheme)
    DisplayTheme.Init(vDT)
  End Sub

  Private Sub SetDefaultFonts()
    Dim vDefaultTheme As String = GetFontDetailsFromResource(DEFAULT_SCHEME)
    Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
    Dim vFonts As FontThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vDefaultTheme)), FontThemeSettings)
    fsForm.SelectedFont = vFonts.FormFont
    fsGrid.SelectedFont = vFonts.GridFont
    fsSelectionPanel.SelectedFont = vFonts.SelectionPanelFont
    fsNavigationPanel.SelectedFont = vFonts.NavigationPanelFont
    fsDisplayLabel.SelectedFont = vFonts.DisplayLabelFont
    fsDisplayItem.SelectedFont = vFonts.DisplayItemFont
    fsDashboardHeading.SelectedFont = vFonts.DashboardHeadingFont
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


  Private Sub cmdApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdApply.Click
    Try
      SetValues()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Protected Overloads Sub frmCardMaintenance_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SystemColorsChanged
    csFormBackColor.Color = DisplayTheme.FormBackColor
    csSplitterBackColor.Color = DisplayTheme.SplitterBackColor
    csGridBackColor.Color = DisplayTheme.GridBackAreaColor
    csPanelBackColor.Color = DisplayTheme.SelectionPanelTreeBackColor
    csButtonPanelBackColor.Color = DisplayTheme.ButtonPanelBackColor
    csGridHyperlinkColor.Color = DisplayTheme.GridHyperlinkColor
    pteDisplayPanel.InitFromPanelTheme(DisplayTheme.DisplayPanelTheme)
    pteEditPanel.InitFromPanelTheme(DisplayTheme.EditPanelTheme)
    pteSelectionPanel.InitFromPanelTheme(DisplayTheme.SelectionPanelTheme)
    pteDisplayLabel.InitFromPanelTheme(DisplayTheme.DisplayLabelTheme)
    pteDisplayData.InitFromPanelTheme(DisplayTheme.DisplayDataTheme)
    pteDashboardHeading.InitFromPanelTheme(DisplayTheme.DashboardHeadingTheme)
    pteToolbar.InitFromPanelTheme(DisplayTheme.ToolbarTheme)
  End Sub

  Private Sub tab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
    If tab.SelectedTab Is tabAppearance Then 'Or tab.SelectedTab Is tabFont Then
      cmdSaveAs.Enabled = True
    Else
      cmdSaveAs.Enabled = False
    End If
  End Sub

  Private Sub cmdSaveAs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSaveAs.Click
    Try
      Dim vDefaults As New ParameterList
      ErrorProvider.SetError(cboSchemes, "")
      Dim vMS As MemoryStream = Nothing
      If cboSchemes.Text.Length > 0 Then
        Try
          Dim vSchemeDescription As String = String.Empty
          If cboSchemes.SelectedIndex > 0 Then
            vSchemeDescription = cboSchemes.DisplayMember
          Else
            vSchemeDescription = cboSchemes.Text
          End If
          Dim vFontId As Integer = SaveFontDetails(If(cboFonts.Text.Length > 0, cboFonts.Text, vSchemeDescription))
          Dim vAppearanceId As Integer = SaveAppearanceDetails(If(cboAppearance.Text.Length > 0, cboAppearance.Text, vSchemeDescription))
          SaveSchemeDetails(vFontId, vAppearanceId, vSchemeDescription)

          'BR19173 - Changes to ensure settings are Saved to the User.config file
          If cboSchemes.DataSource IsNot Nothing Then
            Settings.SchemeID = CInt(cboSchemes.SelectedValue)
          End If
          Settings.AppearanceThemeID = vAppearanceId
          Settings.FontThemeID = vFontId
          Settings.Save()

          GetSchemes()
        Catch vEx As Exception
          DataHelper.HandleException(vEx)
        Finally
          If vMS IsNot Nothing Then vMS.Close()
        End Try
      Else
        ErrorProvider.SetError(cboSchemes, InformationMessages.ImFieldMandatory)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub SetFonts(ByVal pFontNumber As Integer)
    'Load the selected font theme
    If DisplayTheme.MatchesThemeSettings Then
      Try
        Dim vXMLList As New ParameterList(True)
        vXMLList("XmlDataType") = "FT"               'Font Theme
        vXMLList("ItemItemNumber") = Convert.ToString(pFontNumber)
        Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vXMLList)
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
          Dim vRow As DataRowView = vDataTable.DefaultView(0)   'DirectCast(vDataTable.Rows(0), DataRowView)
          Dim vXML As String = vRow.Item("ItemXml").ToString
          Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
          Dim vFontSettings As FontThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), FontThemeSettings)
          fsForm.SelectedFont = vFontSettings.FormFont
          fsGrid.SelectedFont = vFontSettings.GridFont
          fsSelectionPanel.SelectedFont = vFontSettings.SelectionPanelFont
          fsNavigationPanel.SelectedFont = vFontSettings.NavigationPanelFont
          fsDisplayLabel.SelectedFont = vFontSettings.DisplayLabelFont
          fsDisplayItem.SelectedFont = vFontSettings.DisplayItemFont
          fsDashboardHeading.SelectedFont = vFontSettings.DashboardHeadingFont
        End If
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    Else 'If cboFonts.SelectedIndex = 0 Then
      SetDefaultFonts()
    End If

  End Sub


  Private Sub cboFonts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFonts.SelectedIndexChanged
    'Load the selected font theme
    'If DisplayTheme.MatchesThemeSettings OrElse Me.ActiveControl Is sender Then
    If cboFonts.SelectedIndex > 0 Then
      Try
        Dim vRow As DataRowView = DirectCast(cboFonts.SelectedItem, DataRowView)
        Dim vXML As String = vRow.Item("ItemXml").ToString
        Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
        Dim vFontSettings As FontThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), FontThemeSettings)
        fsForm.SelectedFont = vFontSettings.FormFont
        fsGrid.SelectedFont = vFontSettings.GridFont
        fsSelectionPanel.SelectedFont = vFontSettings.SelectionPanelFont
        fsNavigationPanel.SelectedFont = vFontSettings.NavigationPanelFont
        fsDisplayLabel.SelectedFont = vFontSettings.DisplayLabelFont
        fsDisplayItem.SelectedFont = vFontSettings.DisplayItemFont
        fsDashboardHeading.SelectedFont = vFontSettings.DashboardHeadingFont
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    ElseIf cboFonts.SelectedIndex = 0 Then
      SetDefaultFonts()
    End If
    'End If
  End Sub

  Private Sub cboAppearance_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAppearance.SelectedIndexChanged
    'If DisplayTheme.MatchesThemeSettings OrElse Me.ActiveControl Is sender Then
    If cboAppearance.SelectedIndex > 0 Then
      Try
        Dim vRow As DataRowView = DirectCast(cboAppearance.SelectedItem, DataRowView)
        Dim vXML As String = vRow.Item("ItemXml").ToString
        Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
        Dim vAppearance As DisplayThemeSettings = DirectCast(vXS.Deserialize(New System.IO.StringReader(vXML)), DisplayThemeSettings)
        csFormBackColor.Color = vAppearance.FormBackColor
        csSplitterBackColor.Color = vAppearance.SplitterBackColor
        csGridBackColor.Color = vAppearance.GridBackAreaColor
        csPanelBackColor.Color = vAppearance.SelectionPanelTreeBackColor
        csButtonPanelBackColor.Color = vAppearance.ButtonPanelBackColor
        csGridHyperlinkColor.Color = vAppearance.GridHyperlinkColor
        chkHeaderBackgroundSameAsForm.Checked = vAppearance.HeaderBackgroundSameAsForm
        chkUnderlineHyperlinks.Checked = vAppearance.UnderlineHyperlinks
        pteDisplayPanel.InitFromPanelTheme(vAppearance.DisplayPanelThemeSettings)
        pteSelectionPanel.InitFromPanelTheme(vAppearance.SelectionPanelThemeSettings)
        pteEditPanel.InitFromPanelTheme(vAppearance.EditPanelThemeSettings)
        pteDisplayLabel.InitFromPanelTheme(vAppearance.DisplayLabelThemeSettings)
        pteDisplayData.InitFromPanelTheme(vAppearance.DisplayDataThemeSettings)
        pteDashboardHeading.InitFromPanelTheme(vAppearance.DashboardHeadingThemeSettings)
        pteToolbar.InitFromPanelTheme(vAppearance.ToolbarThemeSettings)
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    ElseIf cboAppearance.SelectedIndex = 0 Then
      SetDefaultAppearance()
    End If
    'End If
  End Sub

  Private Function SaveFontDetails(ByVal pFontDesc As String) As Integer
    Dim vMS As MemoryStream = Nothing
    Dim vFontId As Integer = 0
    Try
      vMS = New MemoryStream()
      Dim vXS As New Xml.Serialization.XmlSerializer(GetType(FontThemeSettings))
      Dim vSettings As New FontThemeSettings
      vSettings.FormFont = fsForm.SelectedFont
      vSettings.GridFont = fsGrid.SelectedFont
      vSettings.SelectionPanelFont = fsSelectionPanel.SelectedFont
      vSettings.NavigationPanelFont = fsNavigationPanel.SelectedFont
      vSettings.DisplayLabelFont = fsDisplayLabel.SelectedFont
      vSettings.DisplayItemFont = fsDisplayItem.SelectedFont
      vSettings.DashboardHeadingFont = fsDashboardHeading.SelectedFont
      vXS.Serialize(vMS, vSettings)
      Dim vXMLList As New ParameterList(True)
      vXMLList("XmlDataType") = "FT"               'Font Theme
      vXMLList("ItemDesc") = pFontDesc    'Theme Name
      Dim vReader As New System.IO.StreamReader(vMS)
      vMS.Position = 0
      vXMLList("ItemXml") = vReader.ReadToEnd()
      Dim vReturnList As ParameterList = DataHelper.AddXMLDataItem(vXMLList)
      vFontId = vReturnList.IntegerValue("XmlItemNumber")
      GetFontThemes(vFontId)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      If vMS IsNot Nothing Then vMS.Close()
    End Try
    Return vFontId

  End Function


  Private Function SaveAppearanceDetails(ByVal pAppearanceDesc As String) As Integer
    Dim vMS As MemoryStream = Nothing
    Dim vAppearanceId As Integer = 0
    Try
      vMS = New MemoryStream()
      Dim vXS As New Xml.Serialization.XmlSerializer(GetType(DisplayThemeSettings))
      Dim vAppearance As New DisplayThemeSettings
      vAppearance.FormBackColor = csFormBackColor.Color
      vAppearance.SplitterBackColor = csSplitterBackColor.Color
      vAppearance.GridBackAreaColor = csGridBackColor.Color
      vAppearance.SelectionPanelTreeBackColor = csPanelBackColor.Color
      vAppearance.ButtonPanelBackColor = csButtonPanelBackColor.Color
      vAppearance.GridHyperlinkColor = csGridHyperlinkColor.Color
      vAppearance.HeaderBackgroundSameAsForm = chkHeaderBackgroundSameAsForm.Checked
      vAppearance.UnderlineHyperlinks = chkUnderlineHyperlinks.Checked
      pteDisplayPanel.UpdatePanelTheme(vAppearance.DisplayPanelThemeSettings)
      pteSelectionPanel.UpdatePanelTheme(vAppearance.SelectionPanelThemeSettings)
      pteEditPanel.UpdatePanelTheme(vAppearance.EditPanelThemeSettings)
      pteDisplayLabel.UpdatePanelTheme(vAppearance.DisplayLabelThemeSettings)
      pteDisplayData.UpdatePanelTheme(vAppearance.DisplayDataThemeSettings)
      pteDashboardHeading.UpdatePanelTheme(vAppearance.DashboardHeadingThemeSettings)
      pteToolbar.UpdatePanelTheme(vAppearance.ToolbarThemeSettings)
      vXS.Serialize(vMS, vAppearance)
      Dim vXMLList As New ParameterList(True)
      vXMLList("XmlDataType") = "AT"               'Appearance Theme
      vXMLList("ItemDesc") = pAppearanceDesc   'Theme Name
      Dim vReader As New System.IO.StreamReader(vMS)
      vMS.Position = 0
      vXMLList("ItemXml") = vReader.ReadToEnd()
      Dim vReturnList As ParameterList = DataHelper.AddXMLDataItem(vXMLList)
      vAppearanceId = vReturnList.IntegerValue("XmlItemNumber")
      GetAppearanceThemes(vAppearanceId)
      vMS.Close()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      If vMS IsNot Nothing Then vMS.Close()
    End Try
    Return vAppearanceId
  End Function

  Private Sub tim_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tim.Tick
    Me.Refresh()
    tim.Enabled = False
  End Sub

  Private Sub cboThemes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSchemes.SelectedIndexChanged
    If DisplayTheme.MatchesThemeSettings OrElse Me.ActiveControl Is sender Then
      If cboSchemes.DataSource IsNot Nothing Then
        Try

          Dim vRow As DataRowView = DirectCast(cboSchemes.SelectedItem, DataRowView)

          'Dim vXML As String = vRow.Item("ItemXml").ToString

          If vRow IsNot Nothing Then
            SetAppearanceDropDown(vRow)
            SetFontDropDown(vRow)
          End If
        
        Catch vEx As Exception
          DataHelper.HandleException(vEx)
        End Try
      ElseIf cboSchemes.SelectedIndex = 0 Then
        SetDefaultAppearance()
        SetDefaultFonts()
      End If
    End If
  End Sub

  Private Sub SetAppearanceDropDown(ByVal pRow As DataRowView)
    Dim vDataTable As DataTable = DirectCast(cboAppearance.DataSource, DataTable)
    Dim vFound As EnumerableRowCollection(Of DataRow) = From vDataRow In vDataTable.AsEnumerable Where vDataRow.Field(Of String)(0) = pRow("AppearanceXmlItemNumber").ToString Select vDataRow

    cboAppearance.SelectedIndex = DirectCast(cboAppearance.DataSource, DataTable).Rows.IndexOf(vFound(0))
  End Sub

  Private Sub SetFontDropDown(ByVal pRow As DataRowView)
    Dim vDataTable As DataTable = DirectCast(cboFonts.DataSource, DataTable)
    Dim vFound As EnumerableRowCollection(Of DataRow) = From vDataRow In vDataTable.AsEnumerable Where vDataRow.Field(Of String)(0) = pRow("FontXmlItemNumber").ToString Select vDataRow

    cboFonts.SelectedIndex = DirectCast(cboFonts.DataSource, DataTable).Rows.IndexOf(vFound(0))
  End Sub

  Private Sub SaveSchemeDetails(ByVal pFontId As Integer, ByVal pAppeareance As Integer, ByVal pSchemeDescription As String)
    If pFontId > 0 OrElse pAppeareance > 0 Then
      Try
        Dim vParam As New ParameterList(True)
        vParam.Add("FontXmlItemNumber", pFontId)
        vParam.Add("AppearanceXmlItemNumber", pAppeareance)
        vParam.Add("UserSchemeDesc", pSchemeDescription)
        vParam.Add("IsDefault", "N")
        Dim vResult As ParameterList = DataHelper.AddUserScheme(vParam)
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    End If
  End Sub

End Class
