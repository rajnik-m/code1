Public Class frmTableMaintenance
  Inherits PersistentForm

#Region "Private Members"

  Private Const MAX_ROWS As Integer = 500

  Private mvMaintenanceTables As DataTable
  Private mvParams As ParameterList
  Private mvSubTable1 As String = String.Empty
  Private mvSubTable2 As String = String.Empty
  Private mvSubTable3 As String = String.Empty
  Private mvSubTable4 As String = String.Empty
  Private mvDataSet As DataSet
  Private mvParent As frmTableMaintenance
  Private mvCurrentTable As String = String.Empty
  Private mvLastTable As Integer = -1
  Private mvLastGroup As Integer = -1
  Private mvRightsModified As Boolean


  Private mvTestMode As Boolean
  Private mvTableEntryForm As frmTableEntry
  Private mvLastSystemTable As String = ""

  'This list is used to hold the criteria required to display the subfrom
  'All sub form operations should use this criteria
  Private mvCriteria As ParameterList
#Region "Fields required for Criteria processing on the Scores Table"
  Private mvSelectionSet As Integer
  Private mvfrmGenMGen As frmGenMGen
  Private mvStartRow As Integer
  Private mvLastStartRow As Integer
  Private WithEvents mvFrmEditCriteria As frmEditCriteria
  Private mvMailingInfo As MailingInfo
  Private mvCriteriaSet As Integer
  Private mvMailingTypeCode As String = ""
#End Region

#End Region

#Region "Constructor"

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub Initialise(ByVal pForm As frmTableMaintenance, ByVal pTable As String, ByVal pList As ParameterList, ByVal pDataTable As DataTable)
    mvParent = pForm
    mvCurrentTable = pTable
    mvCriteria = pList
    mvMaintenanceTables = pDataTable
    mvParams = New ParameterList(True)
  End Sub

#End Region

#Region "Control Events"

  Private Sub cboGroups_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      If mvMaintenanceTables Is Nothing Then
        mvRightsModified = True 'Will prevent the loading of data for the table when the table combo is bound
        GetMaintenanceTables()
      End If
      If mvMaintenanceTables IsNot Nothing Then
        ApplyGroupFilter()
        If cboTables.Items.Count > 0 Then
          cboTables.SelectedIndex = -1 'Force a refresh if the group has changed
          cboTables.SelectedIndex = 0
        Else
          mvLastTable = -1
          cmdShowTable.Enabled = False
          cmdNew.Enabled = False
          cmdSelect.Enabled = False
        End If
        cboGroups.Focus()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cboTables_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTables.SelectedIndexChanged
    Try
      If cboTables.SelectedIndex > -1 Then
        'Prevent SelectTable from being called recursively when we update the table notes
        If Not (cboTables.SelectedIndex = mvLastTable AndAlso cboGroups.SelectedIndex = mvLastGroup) Then SelectTable()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmTableMaintenance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      If mvCurrentTable.Length > 0 Then
        If Not mvParent Is Nothing Then
          Me.Top = mvParent.Top
          Me.Left = mvParent.Left
          Me.Width = mvParent.Width
          Me.Height = mvParent.Height
          mvParent.Enabled = False
        End If

        'Hide the tables combobox and caption
        lblTableDesc.Visible = False
        cboTables.Visible = False
        lblGroups.Visible = False
        cboGroups.Visible = False

        DisplaySpreadsheet(mvCurrentTable)
        ShowTable(mvCurrentTable, True)
      Else
        GetMaintenanceGroups()
        ShowTable(String.Empty, False)
      End If
      SafeSetFocus(Me.cboTables) 'BR19621 - Explicitly set Focus to tables combo
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
        'ShowInformationMessage(GetInformationMessage(InformationMessages.ImTooManyItems, MAX_ROWS.ToString()))
        ShowTable(mvCurrentTable, False)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Sub cmdShowTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowTable.Click
    If Not cboTables.SelectedIndex = -1 Then
      Try
        SetMore(False)
        CheckAdministratorNotesSaved()
        ResetParameterList()
        ShowTable(mvCurrentTable, True)
        If dgr.Visible Then dgr.Select() 'Move focus from the combobox
      Catch vException As CareException
        If vException.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
          'ShowInformationMessage(GetInformationMessage(InformationMessages.ImTooManyItems, MAX_ROWS.ToString()))
          ShowTable(mvCurrentTable, False)
        Else
          DataHelper.HandleException(vException)
        End If
      End Try
    End If
  End Sub

  Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
    Try
      SetMore(False)
      CheckAdministratorNotesSaved()
      Dim vParams As New ParameterList(True)
      Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmSelect, mvCurrentTable, vParams)
      'The table entry form will return a dialog result = OK if select criteria was entered
      If Not vResult = System.Windows.Forms.DialogResult.Cancel Then
        'During select dont validate lookups
        mvParams = vParams
        mvParams("IgnoreLookup") = "Y"
        ShowTable(mvCurrentTable, True)
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
        'ShowInformationMessage(GetInformationMessage(InformationMessages.ImTooManyItems, MAX_ROWS.ToString()))
        ShowTable(mvCurrentTable, False)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Function CheckSystemMaintenance() As Boolean
    If cboGroups.SelectedValue IsNot Nothing AndAlso cboGroups.SelectedValue.ToString = "SM" AndAlso mvCurrentTable <> mvLastSystemTable AndAlso mvTestMode = False Then
      If ShowQuestion(QuestionMessages.QmSystemInternal, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then Return True
      mvLastSystemTable = mvCurrentTable
    End If
    Return False
  End Function


  Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      'Set amendedBy & amended on
      Dim vParams As New ParameterList(True)
      vParams("MaintenanceTableName") = mvCurrentTable
      If CheckSystemMaintenance() Then Exit Sub
      Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, mvCurrentTable, vParams, True)
      If vResult = System.Windows.Forms.DialogResult.OK Then
        'Clear the lookup cache data so that we see the newly added records
        DataHelper.ClearCachedLookupData()
        'Clear Contact and Organisation Groups cache data so that we can see the newly added organisation group
        If mvCurrentTable = "organisation_groups" Then
          DataHelper.ClearContactAndOrgGroups()
        End If
        If mvCurrentTable = "maintenance_users" OrElse mvCurrentTable = "maintenance_departments" Then
          'Permissions have been modified...re-fetch the mainteance data
          mvRightsModified = True
          GetMaintenanceTables(mvCurrentTable)
        End If
        mvStartRow = mvLastStartRow
        ShowTable(mvCurrentTable, True)
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
        ' BR14755 On adding a new row don't want to display the too many items msg, just add the entry
        'ShowInformationMessage(GetInformationMessage(InformationMessages.ImTooManyItems, MAX_ROWS.ToString()))
        ShowTable(mvCurrentTable, False)
      Else
        DataHelper.HandleException(vException)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAmend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAmend.Click
    Try
      AmendData()
    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
        ShowTable(mvCurrentTable, False)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtAdminNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdminNotes.TextChanged
    If cmdSave.Visible Then cmdSave.Enabled = True
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    Try
      CheckAdministratorNotesSaved()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgr.RowDoubleClicked
    Try
      If cmdAmend.Enabled Then AmendData()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      If CheckSystemMaintenance() Then Exit Sub
      If ConfirmDelete() Then
        DeleteRow(False)
      End If
    Catch vCareEx As CareException
      Select Case vCareEx.ErrorNumber
        Case CareException.ErrorNumbers.enRecordCannotBeDeleted, CareException.ErrorNumbers.enPrimaryAndRelatedAttributesDontMatch, _
           CareException.ErrorNumbers.enPrimaryRateForMembership, CareException.ErrorNumbers.enPackCannotBeDeleted
          ShowWarningMessage(vCareEx.Message)
        Case CareException.ErrorNumbers.enDeleteParentIfUniqueEntry
          'Display confirmation mesg to user and call delete if required
          If ShowQuestion(vCareEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            DeleteRow(True)
          End If
        Case Else
          DataHelper.HandleException(vCareEx)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vSFD As New SaveFileDialog
      With vSFD
        .Title = ControlText.DlgTitleSaveListAs
        .Filter = "Spreadsheet Files (*.xls)|*.xls|Tab Separated Files (*.tsv)|*.tsv|Comma Separated Files (*.csv)|*.csv"
        .DefaultExt = "xls"
        .FileName = ""
        .CheckPathExists = True
        .OverwritePrompt = True
        If .ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then dgr.SaveList(.FileName)
      End With
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

#End Region

#Region "Private Methods"

  Private Sub InitialiseControls()
    SetControlTheme()
    Me.Text = ControlText.FrmTableMaintenance
    SettingsName = "TableMaintenance"
    MainHelper.SetMDIParent(Me)

    'Lock txtUser.Text if user does not have access rights
    txtAdminNotes.Enabled = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciTableMaintenanceAdminNotes)
  End Sub

  Private Sub GetMaintenanceGroups()
    Dim vMaintenanceGroups As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMaintenanceGroups)
    cboGroups.DisplayMember = "MaintenanceGroupDesc"
    cboGroups.ValueMember = "MaintenanceGroup"
    cboGroups.DataSource = vMaintenanceGroups
    AddHandler cboGroups.SelectedIndexChanged, AddressOf cboGroups_SelectedIndexChanged
    SelectComboBoxItem(cboGroups, "AL") 'Default to "All"

    If mvMaintenanceTables IsNot Nothing Then
      Dim vGroupFilter As String = ""
      Dim vRows() As DataRow = mvMaintenanceTables.Select("MaintenanceGroup = 'SM'")
      If vRows.Length = 0 Then vGroupFilter = "'SM'"
      vRows = mvMaintenanceTables.Select("MaintenanceGroup = 'SI'")
      If vRows.Length = 0 Then
        If vGroupFilter.Length > 0 Then vGroupFilter &= ","
        vGroupFilter &= "'SI'"
      End If
      If vGroupFilter.Length > 0 Then vMaintenanceGroups.DefaultView.RowFilter = String.Format("MaintenanceGroup NOT IN ({0})", vGroupFilter)
    End If
  End Sub

  ''' <summary>
  ''' Display table notes and reset the buttons on the form. 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SelectTable()
    If Not mvRightsModified Then
      Dim vTable As String = String.Empty
      mvLastTable = cboTables.SelectedIndex
      mvLastGroup = cboGroups.SelectedIndex
      If mvLastTable > -1 Then
        CheckAdministratorNotesSaved()
        Dim vDataRow As DataRow() = mvMaintenanceTables.Select(String.Format("TableName = '{0}'", cboTables.SelectedValue.ToString))
        If vDataRow.Length > 0 Then
          vTable = cboTables.SelectedValue.ToString
          txtTableNotes.Text = vDataRow(0).Item("TableNotes").ToString.Replace(Chr(10).ToString, Environment.NewLine)
          txtDefaultValues.Text = vDataRow(0).Item("DefaultNotes").ToString.Replace(Chr(10).ToString, Environment.NewLine)
          txtAdminNotes.Text = vDataRow(0).Item("AdministratorNotes").ToString.Replace(Chr(10).ToString, Environment.NewLine)
        Else
          mvLastTable = -1
          mvLastGroup = -1
        End If
      End If

      cboTables.Enabled = False
      mvSubTable1 = String.Empty
      mvSubTable2 = String.Empty
      mvSubTable3 = String.Empty
      mvSubTable4 = String.Empty
      ShowTable(String.Empty, False)
      mvCurrentTable = vTable
      If mvCurrentTable.Length > 0 Then
        DisplaySpreadsheet(mvCurrentTable)
        cboTables.Enabled = True
      Else
        cmdShowTable.Enabled = False
        cmdExport.Enabled = True
        cmdSelect.Enabled = False
      End If
      SafeSetFocus(cboTables)
    End If
    mvRightsModified = False
  End Sub

  ''' <summary>
  ''' Initialise the spreadsheet according to the selected table
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DisplaySpreadsheet(ByVal pTable As String)
    cmdShowTable.Enabled = True
    cmdSelect.Enabled = True
    mvSubTable1 = String.Empty
    mvSubTable2 = String.Empty
    Select Case pTable
      Case "activities"
        mvSubTable1 = "activity_values"
        cmdSubTable1.Text = ControlText.CmdValues
        mvSubTable2 = "activity_users"
        cmdSubTable2.Text = ControlText.CmdUsers
      Case "activity_groups"
        mvSubTable1 = "activity_group_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "activity_values"
        mvSubTable1 = "activity_value_users"
        cmdSubTable1.Text = ControlText.CmdUsers
        'Case "application_names"
        '  mvSubTable1 = "search_areas"
        '  cmdSubTable1.Text = ControlText.CmdAreas
      Case "bank_accounts"
        mvSubTable1 = "bank_account_departments"
        cmdSubTable1.Text = ControlText.CmdDepartments
      Case "branches"
        mvSubTable1 = "branch_postcodes"
        cmdSubTable1.Text = ControlText.CmdPostcodes
        mvSubTable2 = "branches_historical"
        cmdSubTable2.Text = ControlText.CmdHistorical
      Case "branch_postcodes"
        mvSubTable1 = "branch_postcodes_move"
        cmdSubTable1.Text = ControlText.CmdMove
      Case "bunches"
        mvSubTable1 = "bunch_topics"
        cmdSubTable1.Text = ControlText.CmdTopics
      Case "companies"
        mvSubTable1 = "company_controls"
        cmdSubTable1.Text = ControlText.CmdControls
        mvSubTable2 = "bank_accounts"
        cmdSubTable2.Text = ControlText.CmdBankACs
        'Case "config_names"   
        '  mvSubTable1 = "config"
        '  cmdSubTable1.Text = ControlText.CmdValues
      Case "custom_data_sets"
        mvSubTable1 = "custom_data_set_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "custom_forms"
        mvSubTable1 = "custom_form_controls"
        cmdSubTable1.Text = ControlText.CmdControls
      Case "custom_finders"
        mvSubTable1 = "custom_finder_controls"
        cmdSubTable1.Text = ControlText.CmdControls
      Case "event_fee_bands"
        mvSubTable1 = "event_fee_band_dates"
        cmdSubTable1.Text = ControlText.CmdDates
        mvSubTable2 = "event_fee_band_discounts"
        cmdSubTable2.Text = ControlText.CmdDiscounts
      Case "event_fee_levels"
        mvSubTable1 = "event_adult_fee_levels"
        cmdSubTable1.Text = ControlText.CmdAdultFees
      Case "expenditure_groups"
        mvSubTable1 = "product_groups"
        cmdSubTable1.Text = ControlText.CmdProducts
      Case "geographical_region_types"
        mvSubTable1 = "geographical_regions"
        cmdSubTable1.Text = ControlText.CmdRegions
      Case "geographical_regions"
        mvSubTable1 = "geographical_region_postcodes"
        cmdSubTable1.Text = ControlText.CmdPostcodes
      Case "geographical_region_postcodes"
        mvSubTable1 = "geo_region_postcode_move"
        cmdSubTable1.Text = ControlText.CmdMove
      Case "incentive_schemes"
        mvSubTable1 = "incentive_scheme_reasons"
        cmdSubTable1.Text = ControlText.CmdReasons
      Case "incentive_scheme_reasons"
        mvSubTable1 = "incentive_scheme_products"
        cmdSubTable1.Text = ControlText.CmdProducts
      Case "lookup_groups"
        mvSubTable1 = "lookup_group_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "mailing_templates"
        mvSubTable1 = "mailing_template_paragraphs"
        cmdSubTable1.Text = ControlText.CmdParagraphs
        mvSubTable2 = "mailing_template_documents"
        cmdSubTable2.Text = ControlText.CmdDocuments
      Case "marketing_controls"
        mvSubTable1 = pTable     'special case
        cmdSubTable1.Text = ControlText.CmdCriteria
      Case "membership_types"
        mvSubTable1 = "membership_entitlement"
        cmdSubTable1.Text = ControlText.CmdEntitlement
        mvSubTable2 = "membership_prices"
        cmdSubTable2.Text = ControlText.CmdPrices
      Case "performances"
        mvSubTable1 = pTable     'special case
        cmdSubTable1.Text = ControlText.CmdCriteria
      Case "personnel"
        mvSubTable1 = "personnel_subjects"
        cmdSubTable1.Text = ControlText.CmdSubjects
        mvSubTable2 = "personnel_venues"
        cmdSubTable2.Text = ControlText.CmdVenues
      Case "po_authorisation_levels"
        mvSubTable1 = "po_authorisation_users"
        cmdSubTable1.Text = ControlText.CmdUsers
      Case "post_points"
        mvSubTable1 = "post_point_recipients"
        cmdSubTable1.Text = ControlText.CmdRecipients
      Case "product_nominal_accounts"
        mvSubTable1 = "rate_nominal_accounts"
        cmdSubTable1.Text = ControlText.CmdSuffixes
      Case "products"
        mvSubTable1 = "rates"
        cmdSubTable1.Text = ControlText.CmdRates
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_stock_multiple_warehouses) Then
          mvSubTable2 = "product_warehouses"
          cmdSubTable2.Text = ControlText.CmdWarehouses
        End If
      Case "rates"
        mvSubTable1 = "packed_products"
        cmdSubTable1.Text = ControlText.CmdPacks
        Dim vList As New ParameterList(True)
        Dim vTable As DataTable
        vList("product") = "TEST"
        vList("rate") = "TEST"
        vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRateModifiers, vList)
        If vTable Is Nothing OrElse Not vTable.Columns.Contains("NotExists") Then
          mvSubTable2 = "rate_modifiers"
          cmdSubTable2.Text = ControlText.CmdModifiers
        End If
      Case "relationship_groups"
        mvSubTable1 = "relationship_group_details"
        cmdSubTable1.Text = ControlText.CmdDetails
        'Case "report_parameters"
        '  mvSubTable1 = "fp_controls"
        '  cmdSubTable1.Text = ControlText.CmdControls
        'Case "report_sections"
        '  mvSubTable1 = "report_items"
        '  cmdSubTable1.Text = ControlText.CmdItems
        'Case "reports"
        '  mvSubTable1 = "report_sections"
        '  cmdSubTable1.Text = ControlText.CmdSections
        '  mvSubTable2 = "report_parameters"
        '  cmdSubTable2.Text = ControlText.CmdParams
      Case "search_areas"
        mvSubTable1 = "selection_control"
        cmdSubTable1.Text = ControlText.CmdControl
      Case "selection_control"
        mvSubTable1 = "selection_control_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "service_controls"
        mvSubTable1 = "service_control_start_days"
        cmdSubTable1.Text = ControlText.CmdStartDays
      Case "service_products"
        mvSubTable1 = "service_start_days"
        cmdSubTable1.Text = ControlText.CmdStartDays
      Case "scores"
        mvSubTable1 = "scoring_details"
        cmdSubTable1.Text = ControlText.CmdDetails
        mvSubTable2 = pTable   'special case
        cmdSubTable2.Text = ControlText.CmdCriteria
      Case "scoring_details"
        mvSubTable1 = "selection_control"
        cmdSubTable1.Text = ControlText.CmdControl
      Case "skill_levels"
        mvSubTable1 = "skill_fee_levels"
        cmdSubTable1.Text = ControlText.CmdFeeLevels
      Case "sub_topics"
        mvSubTable1 = "sub_topic_paragraphs"
        cmdSubTable1.Text = ControlText.CmdParagraphs
      Case "suppression_groups"
        mvSubTable1 = "suppression_group_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "topic_groups"
        mvSubTable1 = "topic_group_details"
        cmdSubTable1.Text = ControlText.CmdDetails
      Case "topics"
        mvSubTable1 = "sub_topics"
        cmdSubTable1.Text = ControlText.CmdSubTopics
      Case "relationships"
        Dim vList As New ParameterList(True)
        Dim vTable As DataTable
        vList("Relationship") = "TEST"
        vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRelationshipStatuses, vList)
        If vTable Is Nothing OrElse Not vTable.Columns.Contains("NotExists") Then
          mvSubTable1 = "relationship_statuses"
          cmdSubTable1.Text = ControlText.CmdStatuses
        End If
      Case "surveys"
        mvSubTable1 = "survey_versions"
        mvSubTable2 = "survey_questions"
        mvSubTable3 = "survey_contact_groups"
        ' "duplicate_survey" is not a table. it is a command text for Duplicate Button
        mvSubTable4 = "duplicate_survey"
        cmdSubTable1.Text = ControlText.CmdVersions 'Versions
        cmdSubTable2.Text = ControlText.CmdQuestions 'Questions
        cmdSubTable3.Text = ControlText.CmdGroups 'Groups
        cmdSubTable4.Text = ControlText.CmdDuplicate 'Duplicate
      Case "survey_questions"
        mvSubTable1 = "survey_answers"
        cmdSubTable1.Text = ControlText.CmdAnswers 'Answers
      Case "vat_rates"
        mvSubTable1 = "vat_rate_history"
        cmdSubTable1.Text = ControlText.CmdHistory  'History
      Case "web_documents"
        mvSubTable1 = "web_document_topics"
        cmdSubTable1.Text = ControlText.CmdTopics  'Topics
      Case "exam_cert_reprint_types"
        mvSubTable1 = "exam_cert_reprint_type_items"
        cmdSubTable1.Text = "Items" 'ControlText.CmdRates
      Case "workstream_groups"
        mvSubTable1 = "workstream_group_actions"
        cmdSubTable1.Text = ControlText.CmdActionTemplates
        mvSubTable2 = "workstream_group_outcomes"
        cmdSubTable2.Text = ControlText.CmdOutcomes
      Case Else
    End Select

    DisplayCaption(False, String.Empty, 0) 'Table Contents (No Records Selected)
    GetAccessRights(pTable)
  End Sub

  ''' <summary>
  ''' Creates a ParameterList that contains only the ConnectionData
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ResetParameterList()
    If mvParams Is Nothing Then
      mvParams = New ParameterList(True)
    Else
      mvParams.Clear()
      mvParams.AddConnectionData()
    End If
  End Sub

  ''' <summary>
  ''' Set the displayed details on the frmTableMaint form, either
  ''' at their initial values, or to re/display the current table, prior to or after modification
  ''' </summary>
  ''' <param name="pTable"></param>
  ''' <param name="pReadData"></param>
  ''' <remarks></remarks>
  Private Sub ShowTable(ByVal pTable As String, ByVal pReadData As Boolean)
    If pReadData Then
      FillSpreadsheet(pTable)
      If dgr.ColumnCount < 1 Then
        ShowWarningMessage(InformationMessages.ImNoEditableColumns)
        pReadData = False    'clear itself again
      Else
        GetAccessRights(pTable)
        SetCommands()
        If dgr.RowCount < 1 AndAlso mvTestMode = False Then ShowInformationMessage(InformationMessages.ImNoRecordsFound)
      End If
    End If

    If pReadData = False Then
      'clean up to the initial state
      dgr.Clear()
      ResetParameterList()
      DisplayCaption(False, String.Empty, 0)
      SetCommands()
    End If

    txtTableNotes.Visible = Not pReadData
    txtDefaultValues.Visible = Not pReadData
    txtAdminNotes.Visible = Not pReadData
    dgr.Visible = pReadData
  End Sub

  ''' <summary>
  ''' Displays the data in the grid
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub FillSpreadsheet(ByVal pTable As String)
    mvCurrentTable = pTable
    mvParams("MaintenanceTableName") = pTable

    'Should only be true when loading a sub form
    If Not mvCriteria Is Nothing Then
      'Add the criteria from the parent form to the param list
      For Each vParam As String In mvCriteria.Keys
        mvParams(vParam) = mvCriteria(vParam)
      Next
    End If
    If mvStartRow > 0 Then
      mvParams("StartRow") = mvStartRow.ToString
      mvParams("NumberOfRows") = MAX_ROWS.ToString
      mvLastStartRow = mvStartRow
    Else
      If mvParams.ContainsKey("StartRow") Then
        mvParams("StartRow") = mvStartRow.ToString
        'mvParams.Remove("NumberOfRows")
      End If
      mvLastStartRow = 0
    End If
    Try
      mvParams("SmartClient") = "Y"
      mvDataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstTableMaintenanceData, CareServices.XMLTableDataSelectionTypes), mvParams)
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enCountGreaterThanNumberOfRows Then
        If mvTestMode OrElse ShowQuestion(GetInformationMessage(QuestionMessages.QmShowRecordsWithMore, MAX_ROWS.ToString), MessageBoxButtons.OKCancel) = DialogResult.OK Then
          mvParams("StartRow") = mvStartRow.ToString
          mvParams("NumberOfRows") = MAX_ROWS.ToString
          mvDataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstTableMaintenanceData, CareServices.XMLTableDataSelectionTypes), mvParams)
        Else
          Throw vEx
        End If
      Else
        Throw vEx
      End If
    End Try
    dgr.AutoSetRowHeight = True
    dgr.Populate(mvDataSet)
    If dgr.RowCount > MAX_ROWS Then dgr.RowCount = MAX_ROWS
    Dim vWhere As String = BuildSelectWhere()
    Dim vTable As DataTable = mvDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      DisplayCaption(True, vWhere, vTable.Rows.Count)
      SetMore(vTable.Rows.Count > MAX_ROWS)
    Else
      DisplayCaption(True, vWhere, 0)
    End If
  End Sub

  ''' <summary>
  ''' Enable/disable New/Amend/Delete based on the user permissions
  ''' </summary>
  ''' <param name="pTable"></param>
  ''' <remarks></remarks>
  Private Sub GetAccessRights(ByVal pTable As String)
    Dim vInsert As Boolean
    Dim vUpdate As Boolean
    Dim vDelete As Boolean

    Dim vTable As String = pTable
    'We don't want to expose lookup_group_details as a maintainable table in it's own right
    'but only through using the Details button from lookup_group
    'so it needs to inherit it's access rights from lookup_groups
    If vTable = "lookup_group_details" Then vTable = "lookup_groups"
    'similarly access to packed_products should be restricted based on access to products
    If vTable = "packed_products" Then vTable = "products"

    Dim vDataRow As DataRow() = mvMaintenanceTables.Select(String.Format("TableName = '{0}'", vTable))
    If vDataRow.Length > 0 Then
      vInsert = vDataRow(0).Item("PrivInsert").ToString = "Y"     'New
      vUpdate = vDataRow(0).Item("PrivUpdate").ToString = "Y"     'Amend
      vDelete = vDataRow(0).Item("PrivDelete").ToString = "Y"     'Delete
    End If

    cmdNew.Enabled = vInsert
    cmdAmend.Enabled = vUpdate
    cmdDelete.Enabled = vDelete
  End Sub

  ''' <summary>
  ''' Toggles the buttons on the form
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SetCommands()
    Dim vTable As Boolean
    Dim vAmend As Boolean
    Dim vDelete As Boolean
    Dim vSub1 As Boolean
    Dim vSub2 As Boolean
    Dim vSub3 As Boolean
    Dim vSub4 As Boolean
    Dim vSave As Boolean

    If dgr.RowCount < 1 Then
      'Save Table New Select Close
      vTable = True
      vAmend = False
      vDelete = False
      vSave = True
    Else
      'Amend Delete New Select Close Export
      vTable = False
      vAmend = True
      vDelete = True
      vSub1 = Len(mvSubTable1) > 0
      vSub2 = Len(mvSubTable2) > 0
      vSub3 = Len(mvSubTable3) > 0
      vSub4 = Len(mvSubTable4) > 0
    End If

    cmdShowTable.Visible = vTable
    cmdExport.Visible = Not vTable
    cmdAmend.Visible = vAmend
    cmdDelete.Visible = vDelete
    cmdSubTable1.Visible = vSub1
    cmdSubTable2.Visible = vSub2
    cmdSubTable3.Visible = vSub3
    cmdSubTable4.Visible = vSub4
    cmdSave.Visible = vSave
    cmdSave.Enabled = False
    bpl.RepositionButtons()
  End Sub

  ''' <summary>
  ''' Display the appropriate caption based on the data that is currently displayed
  ''' </summary>
  ''' <param name="pGridVisible"></param>
  ''' <param name="pCriteria"></param>
  ''' <param name="pRecCount"></param>
  ''' <remarks></remarks>
  Private Sub DisplayCaption(ByVal pGridVisible As Boolean, ByVal pCriteria As String, ByVal pRecCount As Integer)
    If pGridVisible Then
      Dim vRecordCount As String
      If mvStartRow > 0 OrElse pRecCount > MAX_ROWS Then
        vRecordCount = String.Format("{0}-{1}", mvStartRow + 1, pRecCount + mvStartRow - 1)
        If pRecCount > MAX_ROWS Then vRecordCount &= "+"
      Else
        vRecordCount = pRecCount.ToString
      End If
      If pCriteria.Length > 0 Then
        lblContents.Text = String.Format(ControlText.LblTableMaintenanceSelectRecordCount, vRecordCount, pCriteria)
      Else
        Dim vTableName As String = GetDescription(mvCurrentTable)
        lblContents.Text = String.Format(ControlText.LblTableMaintenanceRecordCount, vTableName, vRecordCount.ToString)
      End If
    Else
      lblContents.Text = ControlText.LblTableMaintenanceNoRecordsSelected
    End If
  End Sub

  Private Function GetDescription(ByVal pString As String) As String
    Return StrConv(Replace(pString, "_", " "), VbStrConv.ProperCase)
  End Function

  ''' <summary>
  ''' Populates the combo box with a list of tables that the user has access to
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetMaintenanceTables(Optional ByVal pSelectedTable As String = "")
    mvMaintenanceTables = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMaintenanceTables)
    If mvMaintenanceTables Is Nothing Then
      For Each control As Control In Me.Controls
        control.Visible = False
      Next
      ShowInformationMessage(String.Format(InformationMessages.ImNoMaintenanceTablesForUser, AppValues.Logname))    'No tables may be maintained by user:
      Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
    Else
      Dim vSelectedItem As DataRowView = Nothing
      If pSelectedTable.Length > 0 Then
        mvMaintenanceTables.DefaultView.RowFilter = "TableName = '" & pSelectedTable & "'"
        If mvMaintenanceTables.DefaultView.Count = 1 Then
          vSelectedItem = mvMaintenanceTables.DefaultView.Item(0)
          mvMaintenanceTables.DefaultView.RowFilter = ""
        Else
          mvMaintenanceTables.DefaultView.RowFilter = ""
          mvRightsModified = False
        End If
      End If

      cboTables.DisplayMember = "TableNameDesc"
      cboTables.ValueMember = "TableName"
      cboTables.DataSource = mvMaintenanceTables.DefaultView
      If vSelectedItem IsNot Nothing Then cboTables.SelectedItem = vSelectedItem
    End If
  End Sub

  Private Sub ApplyGroupFilter()
    If cboGroups.SelectedValue.ToString = "AL" Then
      mvMaintenanceTables.DefaultView.RowFilter = String.Format("MaintenanceGroup NOT IN ('SI','SM')")
    Else
      mvMaintenanceTables.DefaultView.RowFilter = String.Format("MaintenanceGroup = '{0}'", cboGroups.SelectedValue)
    End If
  End Sub

  ''' <summary>
  ''' Display Table Entry form to edit the data and refresh the grid
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub AmendData()
    Dim vParams As New ParameterList(True)
    vParams("MaintenanceTableName") = mvCurrentTable
    If CheckSystemMaintenance() Then Exit Sub
    GetGridValues(vParams, True)
    'Set amendedBy & amended on
    vParams("AmendedBy") = AppValues.Logname
    vParams("AmendedOn") = Today.ToShortDateString
    Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, mvCurrentTable, vParams)
    If vResult = System.Windows.Forms.DialogResult.OK Then
      DataHelper.ClearCachedLookupData()
      If mvCurrentTable = "maintenance_users" OrElse mvCurrentTable = "maintenance_departments" Then
        'Permissions have been modified...re-fetch the maintenance data
        'mvParams sometimes gets cleared so need to re-set it
        Dim vSaveParams As New ParameterList()
        vSaveParams.FillFromValueList(mvParams.ValueList)
        mvRightsModified = True
        GetMaintenanceTables(mvCurrentTable)
        mvParams = vSaveParams
      End If
      mvStartRow = mvLastStartRow
      ShowTable(mvCurrentTable, True)
    End If
  End Sub

  ''' <summary>
  ''' Add the values for the current row to the parameter list
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetGridValues(ByVal pParams As ParameterList, Optional ByVal pAddNullValues As Boolean = False)
    'Get value for the currentRow
    Dim vValue As String = String.Empty
    For vIndex As Integer = 0 To dgr.ColumnCount - 1
      vValue = dgr.GetValue(dgr.CurrentRow, vIndex)
      If pAddNullValues Then
        pParams(ProperName(dgr.ColumnName(vIndex))) = vValue
      Else
        If vValue.Trim.Length > 0 Then pParams(ProperName(dgr.ColumnName(vIndex))) = vValue
      End If
    Next
  End Sub

  ''' <summary>
  ''' Displays the table entry form in Select/Add/Edit mode and sets the caption
  ''' </summary>
  ''' <param name="pEditMode"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DisplayTableEntry(ByVal pEditMode As CareNetServices.XMLTableMaintenanceMode, ByVal pTable As String, ByVal pParams As ParameterList, Optional ByVal pAddMore As Boolean = False) As DialogResult
    Dim vForm As New frmTableEntry(pEditMode, pTable, pParams, mvCriteria, pAddMore)
    'set the form caption and whether the Add More menu is visible
    Select Case pEditMode
      Case CareNetServices.XMLTableMaintenanceMode.xtmmNew
        vForm.Text = ControlText.FrmAddTo & GetDescription(mvCurrentTable)
      Case CareNetServices.XMLTableMaintenanceMode.xtmmAmend
        vForm.Text = ControlText.FrmAmend & GetDescription(mvCurrentTable)
      Case CareNetServices.XMLTableMaintenanceMode.xtmmSelect
        vForm.Text = ControlText.FrmSelectForm & GetDescription(mvCurrentTable)
    End Select
    mvTableEntryForm = vForm
    vForm.TableMaintenance = True
    Return vForm.ShowDialog()
    mvTableEntryForm = Nothing
  End Function

  ''' <summary>
  ''' Check if admin notes are entered/modified and prompt to save
  ''' </summary>
  ''' <remarks></remarks>
  Private Function CheckAdministratorNotesSaved() As Boolean
    Dim vResult As Boolean = True
    If cmdSave.Enabled Then
      If ConfirmInsert() Then
        Dim vParams As New ParameterList(True)
        vParams("TableName") = mvCurrentTable
        vParams("AdministratorNotes") = txtAdminNotes.Text.Trim
        DataHelper.UpdateTableNote(vParams)
        'Update the tables in the datatable to avoid fetching the maintenance data again
        Dim vDataRow As DataRow() = mvMaintenanceTables.Select(String.Format("TableName = '{0}'", mvCurrentTable))
        vDataRow(0).Item("AdministratorNotes") = txtAdminNotes.Text.Trim
        cmdSave.Enabled = False
      End If
    End If
    Return vResult
  End Function

  ''' <summary>
  ''' Delete a data row
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DeleteRow(ByVal pConfirmDelete As Boolean)
    Dim vParams As New ParameterList(True)
    vParams("MaintenanceTableName") = mvCurrentTable
    vParams("TableMaintenance") = "Y"
    If pConfirmDelete Then vParams("ConfirmDelete") = "Y"
    'Add the data from the row to the param list
    GetGridValues(vParams, True)
    DataHelper.DeleteTableMaintenanceData(vParams)
    DataHelper.ClearCachedLookupData()
    'Clear Contact and Organisation Groups cache data so that we would not see the removed organisation group
    If mvCurrentTable = "organisation_groups" Then
      DataHelper.ClearContactAndOrgGroups()
    End If
    If mvCurrentTable = "maintenance_users" OrElse mvCurrentTable = "maintenance_departments" Then
      'Permissions have been modified...re-fetch the mainteance data
      mvRightsModified = True
      GetMaintenanceTables(mvCurrentTable)
    End If
    ShowTable(mvCurrentTable, True)
  End Sub

  ''' <summary>
  ''' Builds the where clause to be displayed above the grid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildSelectWhere() As String
    Dim vWhere As New StringBuilder()
    If Not mvParams Is Nothing Then
      Dim vOperator As String
      Dim vValue As String

      For index As Integer = 0 To dgr.ColumnCount - 1
        Select Case dgr.ColumnName(index)
          Case "amended_by", "amended_on", "created_by", "created_on"
            '
          Case Else
            'While displaying a sub form use the criteria from the parent form to get the data
            'but dont display the filter criteria string
            'Add the criteria from the parent form to the param list
            Dim vProperName As String = ProperName(dgr.ColumnName(index))
            If vProperName = "OrderNumber" Then vProperName = "PaymentPlanNumber"
            If mvParams.Contains(vProperName) _
              AndAlso (mvCriteria Is Nothing OrElse (Not mvCriteria.ContainsKey(vProperName))) Then
              vOperator = "="
              vValue = mvParams(vProperName)
              If vValue.Contains("*") Then
                vOperator = "like"
                vValue = vValue.Replace("*"c, "%"c)
              End If
              If vWhere.Length > 0 Then vWhere.Append(" AND ")
              vWhere.AppendFormat("{0} {1} '{2}'", dgr.ColumnName(index), vOperator, vValue)
            End If
        End Select
      Next
    End If
    Return vWhere.ToString
  End Function

#End Region

  Private Sub SafeSetFocus(ByVal pControl As Control)
    If pControl.Enabled AndAlso pControl.Visible Then pControl.Focus()
  End Sub

  Private Sub cmdSubTable1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubTable1.Click
    Try
      ShowSubForm(mvSubTable1)
    Catch vCareEx As CareException
      Select Case vCareEx.ErrorNumber
        Case CareException.ErrorNumbers.enNoBranchPostcodesSelectedForMove, CareException.ErrorNumbers.enNoGeographicalRegionPostcodesSelectedForMove
          ShowWarningMessage(vCareEx.Message)
        Case Else
          DataHelper.HandleException(vCareEx)
      End Select
    End Try
  End Sub

  Private Sub cmdSubTable2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubTable2.Click
    Try
      ShowSubForm(mvSubTable2)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSubTable3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubTable3.Click
    Try
      ShowSubForm(mvSubTable3)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSubTable4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubTable4.Click
    Try
      ShowSubForm(mvSubTable4)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ShowSubForm(ByVal pTable As String)
    Dim vForm As frmTableMaintenance
    Dim vParams As New ParameterList(True)

    Select Case pTable
      Case "performances", "scores", "marketing_controls"
        ProcessCriteriaSet(pTable)
      Case "branch_postcodes_move"
        MoveBranchPostcode()
      Case "branches_historical"
        MakeBranchHistorical()
      Case "geo_region_postcode_move"
        MoveRegionPostcode()
      Case "duplicate_survey"
        DuplicateSurvey()
      Case Else
        Select Case pTable
          Case "activity_group_details"
            AddParameter(vParams, "activity_group")
          Case "activity_values", "activity_users"
            AddParameter(vParams, "activity")
          Case "activity_value_users"
            AddParameter(vParams, "activity")
            AddParameter(vParams, "activity_value")
          Case "bank_account_departments"
            AddParameter(vParams, "bank_account")
          Case "bank_accounts"
            AddParameter(vParams, "company")
          Case "branch_postcodes"
            AddParameter(vParams, "branch")
          Case "bunch_topics"
            AddParameter(vParams, "bunch")
          Case "company_controls"
            AddParameter(vParams, "company")
            'Case "config"
            '  AddParameter(vParams, "config_name")
          Case "custom_data_set_details"
            AddParameter(vParams, "custom_data_set")
          Case "custom_finder_controls"
            AddParameter(vParams, "custom_finder")
          Case "custom_form_controls"
            AddParameter(vParams, "custom_form")
          Case "event_fee_band_dates", "event_fee_band_discounts"
            AddParameter(vParams, "event_fee_band")
          Case "event_adult_fee_levels"
            AddParameter(vParams, "event_fee_level")
            'Case "fp_controls"                            'Coming from report_parameters
            '  AddReportParameters(vParams)
          Case "geographical_regions"
            AddParameter(vParams, "geographical_region_type")
          Case "geographical_region_postcodes"
            AddParameter(vParams, "geographical_region_type")
            AddParameter(vParams, "geographical_region")
          Case "incentive_scheme_reasons"
            AddParameter(vParams, "incentive_scheme")
          Case "incentive_scheme_products"
            AddParameter(vParams, "incentive_scheme")
            AddParameter(vParams, "reason_for_despatch")
          Case "mailing_template_paragraphs"
            AddParameter(vParams, "mailing_template")
          Case "membership_entitlement"
            AddParameter(vParams, "membership_type")
          Case "membership_prices"
            AddParameter(vParams, "membership_type")
          Case "packed_products"
            If CheckPackProduct() Then
              AddParameter(vParams, "product")
              AddParameter(vParams, "rate")
            Else
              ShowWarningMessage(InformationMessages.ImPacksOnlyAvailableForPackProducts)
              Exit Sub
            End If
          Case "personnel_subjects", "personnel_venues"
            AddParameter(vParams, "contact_number")
          Case "po_authorisation_users"
            AddParameter(vParams, "po_authorisation_level")
          Case "post_point_recipients"
            AddParameter(vParams, "post_point")
          Case "product_groups"
            AddParameter(vParams, "expenditure_group")
          Case "product_warehouses"
            If CheckStockProductForWarehouse() Then
              AddParameter(vParams, "product")
            Else
              ShowWarningMessage(InformationMessages.ImWarehousesOnlyAvailableForStockProducts)
              Exit Sub
            End If
          Case "rates"
            AddParameter(vParams, "product")
          Case "relationship_group_details"
            AddParameter(vParams, "relationship_group")
            'Case "report_sections"
            '  AddParameter(vParams, "report_number")
            'Case "report_items"
            '  AddParameter(vParams, "report_number")
            '  AddParameter(vParams, "section_number")
            'Case "report_parameters"
            '  AddParameter(vParams, "report_number")
          Case "search_areas"
            AddParameter(vParams, "application_name")
          Case "selection_control"
            AddParameter(vParams, "search_area")
            AddParameter(vParams, "application_name")
          Case "selection_control_details"
            AddParameter(vParams, "search_area")
            AddParameter(vParams, "application_name")
            AddParameter(vParams, "c_o")
          Case "service_control_start_days"
            AddParameter(vParams, "contact_group")
          Case "service_start_days"
            AddParameter(vParams, "contact_number")
          Case "scoring_details"
            AddParameter(vParams, "score")
          Case "skill_fee_levels"
            AddParameter(vParams, "skill_level")
          Case "sub_topics"
            AddParameter(vParams, "topic")
          Case "sub_topic_paragraphs"
            AddParameter(vParams, "topic")
            AddParameter(vParams, "sub_topic")
          Case "suppression_group_details"
            AddParameter(vParams, "suppression_group")
          Case "mailing_template_documents"
            AddParameter(vParams, "mailing_template")
          Case "rate_nominal_accounts"
            AddParameter(vParams, "nominal_account")
          Case "lookup_group_details"
            AddParameter(vParams, "lookup_group")
          Case "topic_group_details"
            AddParameter(vParams, "topic_group")
          Case "relationship_statuses"
            AddParameter(vParams, "relationship")
          Case "survey_versions"
            AddParameter(vParams, "survey_number")
          Case "survey_questions"
            AddParameter(vParams, "survey_number")
          Case "survey_contact_groups"
            AddParameter(vParams, "survey_number")
          Case "survey_answers"
            AddParameter(vParams, "survey_question_number")
          Case "rate_modifiers"
            AddParameter(vParams, "product")
            AddParameter(vParams, "rate")
          Case "vat_rate_history"
            AddParameter(vParams, "vat_rate")
          Case "web_documents"
            AddParameter(vParams, "web_document_number")
          Case "web_document_topics"
            AddParameter(vParams, "web_document_number")
          Case "web_document_extensions"
            AddParameter(vParams, "web_document_extension")
          Case "exam_cert_reprint_type_items"
            AddParameter(vParams, "exam_cert_reprint_type")
          Case "workstream_group_actions"
            AddParameter(vParams, "workstream_group")
          Case "workstream_group_outcomes"
            AddParameter(vParams, "workstream_group")
        End Select
        vForm = New frmTableMaintenance
        vForm.Initialise(Me, pTable, vParams, mvMaintenanceTables)
        vForm.Show()
    End Select
  End Sub

  Private Function CheckPackProduct() As Boolean
    Dim vPackProd As Boolean
    Dim vProduct As String = GetGridValue("product")
    If vProduct.Length > 0 Then
      Dim vList As New ParameterList(True)
      vList("Product") = vProduct
      Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
      vPackProd = vDataRow IsNot Nothing AndAlso BooleanValue(vDataRow.Item("PackProduct").ToString)
    End If
    Return vPackProd
  End Function

  Private Function CheckStockProductForWarehouse() As Boolean
    Return BooleanValue(GetGridValue("stock_item"))
  End Function

  ''' <summary>
  ''' Gets column value for the currently selected row
  ''' </summary>
  ''' <param name="pColumn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetGridValue(ByVal pColumn As String) As String
    Return dgr.GetValue(dgr.CurrentRow, pColumn)
  End Function

  Private Sub MoveRegionPostcode()
    Dim vMoveAllInList As Boolean = False

    If BuildSelectWhere().Length > 0 Then vMoveAllInList = ShowQuestion(QuestionMessages.QmMoveAllPostcodesInList, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes

    Dim vOldRegionType As String = GetGridValue("geographical_region_type")
    Dim vOldRegion As String = GetGridValue("geographical_region")
    Dim vOldPostcode As String = GetGridValue("postcode")

    If vMoveAllInList Then
      'Check if all the postcodes belong to the same region type
      For vIndex As Integer = 0 To dgr.RowCount - 1
        If vOldRegionType <> dgr.GetValue(vIndex, "geographical_region_type") Then
          ShowInformationMessage(InformationMessages.ImDifferentGeoRegTypes)
          Return
        End If
      Next
    End If

    Dim vDefaults As New ParameterList()
    vDefaults("RestrictRegionType") = vOldRegionType
    vDefaults("ExcludeRegion") = vOldRegion
    vDefaults("Postcode") = vOldPostcode
    vDefaults("ParentForm") = "frmTableMaintenance"

    Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptMoveRegion, vDefaults)
    'If the user clicked 'cancel', param count will be 0
    If vParams.Count > 0 Then
      vParams("OldGeographicalRegionType") = vOldRegionType
      vParams("GeographicalRegionType") = vOldRegionType
      vParams("OldGeographicalRegion") = vOldRegion
      vParams("OldPostcode") = vOldPostcode
      If vMoveAllInList Then
        'Add the select parameters to the list
        vParams("MoveAll") = "Y"
        If mvParams.ContainsKey("GeographicalRegionType") Then vParams("SelectedGeographicalRegionType") = mvParams("GeographicalRegionType")
        If mvParams.ContainsKey("GeographicalRegion") Then vParams("SelectedGeographicalRegion") = mvParams("GeographicalRegion")
        If mvParams.ContainsKey("Postcode") Then vParams("SelectedPostcode") = mvParams("Postcode")
      End If

      DataHelper.MoveRegion(vParams)
      ShowTable(mvCurrentTable, True)
    End If
  End Sub

  Private Sub MoveBranchPostcode()
    Dim vMoveAllInList As Boolean = False
    Dim vDefaults As New ParameterList()
    vDefaults("ExcludeBranch") = GetGridValue("branch")

    If BuildSelectWhere().Length > 0 Then vMoveAllInList = ShowQuestion(QuestionMessages.QmMoveAllPostcodesInList, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes

    Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptMoveBranch, vDefaults)
    'If the user clicked 'cancel', param count will be 0
    If vParams.Count > 0 Then
      vParams("OutwardPostcode") = GetGridValue("outward_postcode")
      If vMoveAllInList Then
        vParams("MoveAll") = "Y"
        If mvParams.ContainsKey("Branch") Then vParams("SelectedBranch") = mvParams("Branch")
        If mvParams.ContainsKey("OutwardPostcode") Then vParams("SelectedOutwardPostcode") = mvParams("OutwardPostcode")
      End If

      DataHelper.MoveBranchPostcode(vParams)
      ShowTable(mvCurrentTable, True)
    End If
  End Sub

  ''' <summary>
  ''' Makes a branch historical. If the branch was previously used a new branch needs to be selected
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub MakeBranchHistorical()
    Dim vParams As New ParameterList
    Dim vPrompt As Boolean

    If ShowQuestion(QuestionMessages.QmMakeBranchHistorical, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then    'Make Branch Historical
      'Check prompt for new branch. i.e If the branch is in use then display a from to select a new branch
      Dim vBranchParam As New ParameterList(True)
      vBranchParam("Branch") = GetGridValue("branch")
      vBranchParam = DataHelper.CheckPromptForNewBranch(vBranchParam)
      vPrompt = BooleanValue(vBranchParam("PromptForNewBranch"))

      If vPrompt Then
        vBranchParam.Clear()
        vBranchParam("ExcludeBranch") = GetGridValue("branch")
        vParams = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptMoveBranch, vBranchParam)
      End If

      If vParams.Count > 0 OrElse Not vPrompt Then
        If vParams.Count = 0 Then vParams.AddConnectionData()
        vParams("OldBranch") = GetGridValue("branch")
        DataHelper.MakeBranchHistorical(vParams)
        ShowTable(mvCurrentTable, True)    're-reads & displays the amended table
      End If
    End If
  End Sub

  Private Sub ProcessCriteriaSet(ByVal pTable As String)
    Dim vCriteriaSet As Integer
    Dim vCode As String = String.Empty
    Dim vDesc As String = String.Empty
    Dim vAddCSet As Boolean
    Dim vCriteriaSetParam As New ParameterList(True)
    vCriteriaSetParam("MaintenanceTableName") = pTable

    If GetGridValue("criteria_set").Length > 0 Then
      vCriteriaSet = CInt(GetGridValue("criteria_set"))
    End If

    If dgr.GetValue(dgr.CurrentDataRow, "performance").Length > 0 Then
      vDesc = "Performance " & GetGridValue("performance")
      vCode = "PA"
      vCriteriaSetParam("Performance") = dgr.GetValue(dgr.CurrentDataRow, "performance")
    ElseIf dgr.GetValue(dgr.CurrentDataRow, "score").Length > 0 Then
      vDesc = "Score " & dgr.GetValue(dgr.CurrentDataRow, "score")
      vCode = "SA"
      vCriteriaSetParam("Score") = dgr.GetValue(dgr.CurrentDataRow, "score")
    End If

    If pTable = "marketing_controls" Then
      vDesc = "Standard Exclusions"                         'NoTranslate
      vCode = "SE"
      For vCol As Integer = 0 To dgr.ColumnCount - 1
        Dim vName As String = dgr.ColumnName(vCol)
        Select Case vName
          Case "amended_on", "amended_by", "created_on", "created_by"
            'Ignore
          Case Else
            Dim vValue As String = dgr.GetValue(dgr.CurrentDataRow, vName)
            vCriteriaSetParam("Old" & ProperName(vName)) = vValue
        End Select
      Next
    End If

    If vCriteriaSet = 0 Then
      If ShowQuestion(QuestionMessages.QmCreateNewSelectionSet, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        vAddCSet = True
      End If
    Else
      vCriteriaSetParam.IntegerValue("CriteriaSet") = vCriteriaSet
      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCriteriaSet, vCriteriaSetParam) = 0 Then vAddCSet = True
    End If

    If vAddCSet Then
      Dim vInsertParams As New ParameterList(True)
      vInsertParams("CriteriaSetDesc") = vDesc
      vInsertParams("ApplicationName") = vCode
      Dim vReturnList As ParameterList = DataHelper.AddCriteriaSet(vInsertParams)
      vCriteriaSet = vReturnList.IntegerValue("CriteriaSetNumber")
      vCriteriaSetParam.IntegerValue("CriteriaSet") = vCriteriaSet
    End If

    vCriteriaSetParam.IntegerValue("CriteriaSet") = vCriteriaSet
    If vCriteriaSet > 0 Then
      mvMailingInfo = New MailingInfo
      mvMailingInfo.Init(vCode, vCriteriaSet)
      mvFrmEditCriteria = New frmEditCriteria(mvMailingInfo, vDesc)
      mvFrmEditCriteria.ShowDialog()

      vCriteriaSetParam.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
      vCriteriaSet = mvMailingInfo.CriteriaSet
      mvMailingInfo.CriteriaRows = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCriteriaSetDetails, vCriteriaSetParam)
      If mvMailingInfo.CriteriaRows > 0 Then
        vCriteriaSetParam.IntegerValue("CriteriaSet") = vCriteriaSet
        DataHelper.UpdateTableMaintenanceData(vCriteriaSetParam)
        dgr.SetValue(dgr.CurrentDataRow, "criteria_set", vCriteriaSet.ToString)
      Else
        vCriteriaSetParam = New ParameterList(True)
        vCriteriaSetParam("MaintenanceTableName") = pTable
        vCriteriaSetParam("DataProtMailingSupp") = dgr.GetValue(dgr.CurrentDataRow, "data_prot_mailing_supp")
        vCriteriaSetParam.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        DataHelper.DeleteCriteriaSet(vCriteriaSetParam)
        dgr.SetValue(dgr.CurrentDataRow, "criteria_set", "")
      End If
    End If
  End Sub

  ''' <summary>
  ''' Reads the current values from the grid and adds them to the parameter list
  ''' </summary>
  ''' <param name="pParams"></param>
  ''' <param name="pAttr"></param>
  ''' <remarks></remarks>
  Private Sub AddParameter(ByVal pParams As ParameterList, ByVal pAttr As String)
    Dim vFound As Boolean = False
    Dim vValue As String = dgr.GetValue(dgr.CurrentRow, pAttr)
    If Not String.IsNullOrEmpty(vValue) Then
      If mvCurrentTable = "product_nominal_accounts" AndAlso pAttr = "nominal_account" Then
        pParams("ProductNominalAccount") = vValue
      Else
        pParams(ProperName(pAttr)) = vValue
      End If
      vFound = True
    End If

    If Not vFound And mvCurrentTable = "scoring_details" Then
      pParams(ProperName(pAttr)) = "SM"
    End If
  End Sub

  Private Sub frmTableMaintenance_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      CheckAdministratorNotesSaved()
      If Not mvParent Is Nothing Then
        mvParent.Enabled = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  ''' <summary>
  ''' Duplicates the survey
  ''' </summary>
  ''' <remarks></remarks>
  ''' 
  Private Sub DuplicateSurvey()
    Try
      Dim vForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateSurvey, Nothing, Nothing)
      Dim vList As ParameterList = New ParameterList(True)
      Dim vReturnList As ParameterList = Nothing
      If vForm.ShowDialog() = DialogResult.OK Then
        vReturnList = vForm.ReturnList
        vList("NewSurveyName") = vReturnList("SurveyName")
        vList("SurveyNumber") = dgr.GetValue(dgr.CurrentRow, "survey_number")
        vReturnList = New ParameterList
        vReturnList = DataHelper.DuplicateSurvey(vList)
        ShowTable(mvCurrentTable, True)
      End If
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(vEx.Message)
      Else
        DataHelper.HandleException(vEx)
      End If
    End Try
  End Sub

  Private Sub cmdMore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdMore.Click
    ShowTable(mvCurrentTable, True)
  End Sub

  Private Sub SetMore(ByVal pVisible As Boolean)
    cmdMore.Visible = pVisible
    bpl.RepositionButtons()
    If pVisible Then
      mvStartRow += MAX_ROWS
    Else
      mvStartRow = 0
    End If
  End Sub
#Region "Criteria EventHandlers"
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pMailingSelection"></param>
  ''' <param name="pCriteriaSet"></param>
  ''' <param name="pSuccess"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding eventshandler in GeneralMailing, this is a copy of that event handler</remarks>
  Private Sub mvFrmEditCriteria_ProcessMailingCriteria(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByRef pSuccess As Boolean) Handles mvFrmEditCriteria.ProcessMailingCriteria
    mvMailingInfo.ProcessMailingCriteria(pCriteriaSet, True, False, Nothing, pSuccess)
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pMailingSelection"></param>
  ''' <param name="pCriteriaSet"></param>
  ''' <param name="pProcessVariables"></param>
  ''' <param name="pEditSegmentCriteria"></param>
  ''' <param name="pList"></param>
  ''' <param name="pSuccess"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding eventshandler in GeneralMailing, this is a copy of that event handler</remarks>
  Private Sub mvFrmEditCriteria_ProcessMailingCriteriaWithOptional(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As ParameterList, ByRef pSuccess As Boolean) Handles mvFrmEditCriteria.ProcessMailingCriteriaWithOptional
    mvMailingInfo.ProcessMailingCriteriaWithOptional(pMailingSelection, pCriteriaSet, pProcessVariables, pEditSegmentCriteria, pList, pSuccess)
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pRunPhase"></param>
  ''' <param name="pList"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding eventshandler in GeneralMailing, this is a copy of that event handler</remarks>
  Private Sub mvFrmEditCriteria_ProcessSelection(ByVal pRunPhase As String, ByVal pList As ParameterList) Handles mvFrmEditCriteria.ProcessSelection
    Try
      Dim vStoreGvRev As Integer
      If pList Is Nothing Then pList = New ParameterList(True)
      Dim vCountOnly As Boolean

      mvMailingInfo.GenerateStatus = MailingInfo.MailingGenerateResult.mgrNone
      mvCriteriaSet = mvMailingInfo.CriteriaSet
      mvSelectionSet = mvMailingInfo.SelectionSet
      vCountOnly = (pRunPhase = "Count")
      If mvMailingInfo.CriteriaRows = 0 Then
        If mvMailingInfo.Revision > 0 Then
          If mvMailingInfo.SelectionSet > 0 Then mvMailingInfo.SelectionCount = GetMailingSelectionCount()
          If mvMailingInfo.SelectionCount > 0 Then
            mvfrmGenMGen = New frmGenMGen(mvMailingTypeCode, mvMailingInfo, mvSelectionSet)
            mvfrmGenMGen.ShowDialog()
          End If
        Else
          ShowInformationMessage(InformationMessages.ImNoCriteria)    'No criteria entered
        End If
      Else
        If mvMailingInfo.Revision = 0 Then
          'There is no selection set
          vStoreGvRev = mvMailingInfo.Revision
          mvMailingInfo.Revision = mvMailingInfo.Revision + 1
          pList("SelectionSetNumber") = mvSelectionSet.ToString
          pList("Revision") = mvMailingInfo.Revision.ToString
          pList("ApplicationName") = AppValues.MailingApplicationCode(mvMailingInfo.TaskType) 'mvMailingInfo.MailingTypeCode.ToString
          pList("RunPhase") = pRunPhase
          pList("CriteriaSet") = mvCriteriaSet.ToString
          pList("ExclusionCriteria") = mvMailingInfo.ExclusionCriteriaSet.ToString
          pList("OrgMailTo") = mvMailingInfo.OrganisationMailTo
          pList("OrgMailWhere") = mvMailingInfo.OrganisationMailWhere
          pList("OrgLabelName") = mvMailingInfo.OrganisationLabelName
          pList("OrgAddressUsage") = mvMailingInfo.OrganisationAddressUsage
          pList("OrgRoles") = mvMailingInfo.OrganisationRoles
          pList("OrgIncludeHistoricRoles") = IIf(mvMailingInfo.IncludeHistoricRoles, "Y", "N").ToString
          pList("BypassCount") = IIf(mvMailingInfo.BypassCriteriaCount, "Y", "N").ToString
          pList("GeneralMailing") = "Y"

          If vCountOnly Then
            Dim vResults As ParameterList = DataHelper.ProcessMailingCount(pList)
            If vResults.Contains("MailingCount") Then mvMailingInfo.SelectionCount = vResults.IntegerValue("MailingCount") 'GetMailingSelectionCount()
            mvMailingInfo.Revision = 0
          Else
            DataHelper.GetMailingFile(pList, "", False, True)
          End If
          If mvMailingInfo.Revision = 0 Then
            'MouseNormal()
            mvMailingInfo.Revision = vStoreGvRev
          Else
            mvMailingInfo.SelectionCount = GetMailingSelectionCount()
            'MouseNormal()
            If mvMailingInfo.SelectionCount > 0 Then
              mvfrmGenMGen = New frmGenMGen(mvMailingTypeCode, mvMailingInfo, mvSelectionSet)
              mvfrmGenMGen.ShowDialog()
            Else
              mvMailingInfo.Revision = vStoreGvRev
            End If
          End If
        Else
          Dim vForm As frmGenMMerge = New frmGenMMerge(mvMailingInfo, pList)
          vForm.ShowDialog()
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function GetMailingSelectionCount() As Integer
    'Return the number of selected records
    Dim vSelectedRecords As Integer
    vSelectedRecords = mvMailingInfo.GetMailingSelectionCount(mvMailingInfo.SelectionSet, mvMailingInfo.Revision, mvMailingTypeCode)
    ShowInformationMessage(InformationMessages.ImContactSelected, vSelectedRecords.ToString)     '%s contacts selected
    Return vSelectedRecords
  End Function
#End Region
#Region "Test Code for System Maintenance"

  Private Sub lblGroups_Click(sender As System.Object, e As System.EventArgs) Handles lblGroups.DoubleClick
    If System.Diagnostics.Debugger.IsAttached Then
      SelectComboBoxItem(cboGroups, "SM")
      mvTestMode = True
      Dim vBGWorker As New System.ComponentModel.BackgroundWorker
      AddHandler vBGWorker.DoWork, AddressOf DoBackGroundWork
      AddHandler vBGWorker.ProgressChanged, AddressOf DoProgressChanged
      vBGWorker.WorkerSupportsCancellation = True
      vBGWorker.WorkerReportsProgress = True
      vBGWorker.RunWorkerAsync(Me)
      For vIndex As Integer = 0 To cboTables.Items.Count - 1
        cboTables.SelectedIndex = vIndex
        cmdShowTable_Click(cmdShowTable, New System.EventArgs)
        Me.Refresh()
        cmdAmend_Click(Me, New System.EventArgs)
        Me.Refresh()
      Next
      vBGWorker.CancelAsync()
    End If
    mvTestMode = False
  End Sub

  Private Sub DoBackGroundWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
    Dim vWorker As System.ComponentModel.BackgroundWorker = CType(sender, System.ComponentModel.BackgroundWorker)

    While vWorker.CancellationPending = False
      If mvTableEntryForm IsNot Nothing Then
        vWorker.ReportProgress(1, mvTableEntryForm)
      End If
      System.Threading.Thread.Sleep(1000)
    End While
    e.Cancel = True
  End Sub

  Private Sub DoProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)
    Dim vTableEntry As frmTableEntry = DirectCast(e.UserState, frmTableEntry)
    If Not vTableEntry.IsDisposed AndAlso vTableEntry.Text <> "Testing" Then    'Stop it being recursively called
      vTableEntry.Text = "Testing"
      vTableEntry.epl.DataChanged = True
      vTableEntry.cmdOK.PerformClick()
    End If

  End Sub

#End Region

End Class
