Imports System.Linq
Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class frmImport

  Private mvNoRead As Boolean
  Private mvImportForm As frmImport
  Private mvIsPAFSupported As Boolean
  Private mvValidateHonorifics As Boolean
  Private mvImportFilename As String
  Private mvMultipleDataImportRuns As Boolean
  Private mvMainDefFileName As String
  Private mvMainImportFileName As String
  Private mvSelectedImportType As Integer
  Private mvSelectedImportTypeDesc As String
  Private mvLoadedMultiple As Boolean
  Private mvInitializingFromDef As Boolean
  Private mvImportDefinitionsChanged As Boolean
  Private mvImportTypeChanged As Boolean
  Private mvSelected As Integer
  Private mvDataImport As DataSet 'will contain multiple tables that are used to represent the Data Import class and its properties.
  Private mvMasterDataImport As DataSet
  Private mvDefaultsDataSet As DataSet 'used to bind the defaults grid
  Private mvMaintenanceTables As DataTable 'holds tables that the user has access to including those through his dept 
  Private mvTempMaintAttrCols As DataTable 'Used to store temp col mappings while re-opening a file that has already been mapped
  Private mvMapAttribute As DataRow 'will hold ref to the current col being mapped
  Private mvSystemMaintenanceTables As Integer
  Private mvDefaultImportFolder As String = String.Empty
  Private mvDefaultMapFolder As String = String.Empty
  Private mvDefaultDefFolder As String = String.Empty
  Private mvOverWriteDefFile As Boolean
  Private WithEvents mvTaskInfo As frmTaskInfo = Nothing
  Private mvAttrItems As New List(Of LookupItem)
  'Private mvDefAttrItems As New List(Of LookupItem)

  Private mvCanDisplayForm As Boolean = True  'not to display import form when the user selects a def file and does not select the data files

  'holds a set of values that have changed for each of the data tables. these values
  'can then be passed on to the web service to initialise the data import class before
  'invoking any methods.
  Private mvChangedValues As Dictionary(Of String, List(Of String))
  Private mvJobID As String

  'Table & Datasets
  Private Const DATA_IMPORT As String = "DataImport"
  Private Const DATA_IMPORT_PARAMS As String = "ImportParameters"
  Private Const DATA_IMPORT_FILE As String = "ImportFile"
  Private Const DEFAULTS_ROW As String = "ImportDefaults"
  Private Const DATA_IMPORT_ATTRS As String = "ImportAttributes"
  Private Const IMPORT_OPTIONS As String = "ImportOptions"
  Private Const DEFAULTS_COL As String = "ImportDefaultsColumn"
  Private Const ATTRIBUTE_COL As String = "AttributeColumns"
  Private Const MAPPED_ATTRIBUTE_COLUMNS As String = "MappedAttributeColumns" 'Collection that is part of Mappedttributes
  Private Const MAPPED_ATTRIBUTES As String = "MappedAttributes"
  Private Const MAP_ATTR_COLS As String = "MapAttrColumns" 'Fields to which the columns are mapped along with the heading
  Private Const DEF_FILE_PARAMS As String = "DefFileParams" 'Holds the available import types when loading from a def file
  Private Const VALIDATE_IMPORT As String = "Validate" 'Stores the result from the data import validation
  Private Const MAIN As String = "mvDataImport"
  Private Const MASTER As String = "mvMasterDataImport"
  Private Const ATTR_DATE_FORMAT As String = "AttributeDateFormat"
  Private Const BLANK_HEADING As String = " " 'Single empty space to prevent the grid from setting up the default column name
  Private Property RefreshingAttributes As Boolean

  Public Sub New(ByVal pFileName As String, Optional ByVal pForm As frmImport = Nothing, Optional ByVal pDataImport As DataSet = Nothing, Optional ByVal pChangedValues As Dictionary(Of String, List(Of String)) = Nothing)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvImportFilename = pFileName
    mvImportForm = pForm
    mvMasterDataImport = pDataImport
    mvChangedValues = pChangedValues
    mvMainImportFileName = String.Empty
    If mvImportForm Is Nothing AndAlso mvImportFilename.Substring(mvImportFilename.Length - 4, 4).ToLower = ".def" Then
      mvCanDisplayForm = InitForMultipleImport()  'Do not display the form (when calling Form.Show) if the user does not select the data file
      If GetBooleanValue(DATA_IMPORT, "MultipleDataImportRuns") AndAlso mvMainImportFileName.Length = 0 Then
        'Set the MainImportFileName when using MultipleDataImportsRuns and the data file name is in the definition file
        'This resolves the issue when selecting the second or any other data import type was displaying the incorrect data in the import file grid
        Dim vImportFile As String = GetValue(DATA_IMPORT_PARAMS, "FileName")
        If vImportFile.Length > 0 Then mvMainImportFileName = vImportFile
      End If
    End If
    mvSelected = -1
    mvSelectedImportType = -1
    If mvChangedValues Is Nothing Then mvChangedValues = New Dictionary(Of String, List(Of String))
    If String.IsNullOrWhiteSpace(pFileName) Then
      Dim vParams As New ParameterList(True)
      vParams.AddSystemColumns()
      vParams.Add("ConfigName", "default_import_directory")
      mvDefaultDefFolder = DataHelper.GetConfigValue(vParams).Tables("DataRow").Rows(0)("ConfigValue").ToString
    Else
      mvDefaultDefFolder = Path.GetDirectoryName(Path.GetFullPath(pFileName))
    End If
    mvDefaultImportFolder = mvDefaultDefFolder
    mvDefaultMapFolder = mvDefaultMapFolder
    LookupItemComparer = New ImportAttributeComparer()
    InitialiseControls()
  End Sub

  Private Property LookupItemComparer As IComparer(Of LookupItem)

  Public ReadOnly Property SelectedImportType() As Integer
    Get
      Dim vSelectedImpType As Integer = -1
      If mvImportForm Is Nothing Then
        If TypeOf cboType.SelectedItem Is LookupItem Then
          vSelectedImpType = IntegerValue(DirectCast(cboType.SelectedItem, LookupItem).LookupCode)
        Else
          vSelectedImpType = IntegerValue(cboType.SelectedValue)
        End If
      Else
        vSelectedImpType = mvImportForm.SelectedImportType
      End If
      Return vSelectedImpType
    End Get
  End Property

  Public ReadOnly Property MappedAttribute() As DataRow
    Get
      Return mvMapAttribute
    End Get
  End Property

  Public Property DataImportDS As DataSet
    Get
      Return mvDataImport
    End Get
    Set(value As DataSet)
      mvDataImport = value
    End Set
  End Property

  Public Property DefaultsDS As DataSet
    Get
      Return mvDefaultsDataSet
    End Get
    Set(value As DataSet)
      mvDefaultsDataSet = value
    End Set
  End Property

  Private Sub InitialiseControls()
    SetControlTheme()
    MainHelper.SetMDIParent(Me)
    tabMain.SetItemSizes()  'Adjust the widths of the tabs
    tabSub.SetItemSizes()   'Adjust the widths of the tabs

    dgr.SetCellsEditable(True) 'Allow columns to be selected.
    dgr.AllowSorting = False
    dgrDefaults.AllowSorting = False
    dgr.ContextMenuStrip = dgrMenuStrip

    AddHandler txtControlNumberBlockSize.KeyPress, AddressOf Utilities.IntegerKeyPressHandler
    AddHandler txtNumberOfDays.KeyPress, AddressOf Utilities.IntegerKeyPressHandler

    InitTextLookupBox(txtOrgNumber, "organisations", "organisation_number")

    cboAttrs.DataSource = mvAttrItems
    cboAttrs.DisplayMember = "LookupDesc"
    cboAttrs.ValueMember = "LookupCode"
    'cboDefAttrs.DataSource = mvDefAttrItems
    cboDefAttrs.DataSource = mvAttrItems
    cboAttrs.DisplayMember = "LookupDesc"
    cboAttrs.ValueMember = "LookupCode"


    If mvImportForm IsNot Nothing Then
      PrepareAttributeForm()
      'Lock the parent form while the sub form is open
      'mvImportForm.Enabled = False
    End If
  End Sub

  Private Sub cboType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      'Dim vFileName As String
      Dim vInitMulti As Boolean
      Dim vTypeChanged As Boolean
      Dim vNoRead As Boolean

      'ClearMaps
      If DataImportDS IsNot Nothing AndAlso Not GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad") Then
        If DataImportDS.Tables.Contains(MAP_ATTR_COLS) Then DataImportDS.Tables(MAP_ATTR_COLS).Rows.Clear()
      End If

      dgrDefaults.Clear()
      cboTables.SelectedIndex = -1

      SetControls(False)

      If mvMultipleDataImportRuns AndAlso (mvSelectedImportType <> SelectedImportType) Then
        vInitMulti = True
        SaveMultipleImportFile(mvMainDefFileName, False)
        mvImportFilename = ExtractDefForMultiImport(SelectedImportType)
        SetValue(DATA_IMPORT_PARAMS, "FirstLoad", True)
        If mvMainImportFileName.Length > 0 Then SetValue(DATA_IMPORT_PARAMS, "FileName", mvMainImportFileName)
        GetImportDataSets(mvImportFilename, , False)
        ResetFinancialHistoryOptions()
        SetOptionTabDefaults()
        mvInitializingFromDef = True
        vNoRead = mvNoRead
        InitFormParameters(False)
        mvNoRead = vNoRead
        vTypeChanged = True
      End If

      ImportFileRead(False, Not vTypeChanged)

      'Set the options based on the import type selected
      SetImportOptions()

      If mvInitializingFromDef = False AndAlso vInitMulti = False Then optMIRecordsSuccFromPrevImport.Checked = True
      If vTypeChanged Then mvInitializingFromDef = False
      mvSelectedImportType = IntegerValue(SelectedImportType) ' vType
      mvSelectedImportTypeDesc = cboType.Text
      SetOptionTabs()
      If SelectedImportType = 18 Then chkValCodes.Enabled = False 'ditCMT
      SetValue(DATA_IMPORT_PARAMS, "FirstLoad", "N")
      chkIgnore.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "IgnoreFirstRow")

      If GetValue(DATA_IMPORT_PARAMS, "Separator") = "Fixed" Then
        DisableCheckBox(chkIgnore)
      Else
        chkIgnore.Enabled = True
      End If

      SetValue(DATA_IMPORT_PARAMS, "FirstLoad", "N")
      mvImportTypeChanged = True
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GetMaintenanceGroups()
    GetMaintenanceTables()
    If cboType.Text = "Table Import" Then
      Dim vMaintenanceGroups As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMaintenanceGroups)
      cboGroups.DisplayMember = "MaintenanceGroupDesc"
      cboGroups.ValueMember = "MaintenanceGroup"
      If mvSystemMaintenanceTables = 0 Then
        vMaintenanceGroups.DefaultView.RowFilter = "MaintenanceGroup NOT IN ('SI','SM')"
      Else
        vMaintenanceGroups.DefaultView.RowFilter = "MaintenanceGroup <> 'SI'"
      End If
      cboGroups.DataSource = vMaintenanceGroups
      SelectComboBoxItem(cboGroups, "AL") 'Default to "All"
    End If
  End Sub

  Private Sub cboGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroups.SelectedIndexChanged
    If cboGroups.SelectedItem IsNot Nothing AndAlso cboGroups.SelectedValue.ToString = "SM" Then
      If ShowQuestion(QuestionMessages.QmSystemInternalImport, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.Cancel Then
        SelectComboBoxItem(cboGroups, "AL") 'Default to "All"
        Exit Sub
      End If
    End If
    If cboGroups.SelectedValue.ToString = "AL" Then
      mvMaintenanceTables.DefaultView.RowFilter = String.Format("MaintenanceGroup NOT IN ('SI','SM')")
    Else
      mvMaintenanceTables.DefaultView.RowFilter = String.Format("MaintenanceGroup = '{0}'", cboGroups.SelectedValue)
    End If
    cboTables.SelectedIndex = -1
    If cboTables.Items.Count > 0 Then cboTables.SelectedIndex = 0
  End Sub

  Private Sub GetMaintenanceTables()
    If cboType.Text = "Table Import" Then
      Dim vList As New ParameterList(True)
      vList("PrivInsert") = "Y"
      If mvMaintenanceTables Is Nothing Then mvMaintenanceTables = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMaintenanceTables, vList)
      If mvMaintenanceTables Is Nothing Then
        ShowInformationMessage(String.Format(InformationMessages.ImNoMaintenanceTablesForUser, AppValues.Logname)) 'No tables may be maintained by user:
        mvSystemMaintenanceTables = 0
      Else
        Dim vRows() As DataRow = mvMaintenanceTables.Select("MaintenanceGroup = 'SM'")
        mvSystemMaintenanceTables = vRows.Length
      End If
      mvMaintenanceTables.DefaultView.RowFilter = String.Format("MaintenanceGroup NOT IN ('SI','SM')")
      cboTables.DisplayMember = "TableNameDesc"
      cboTables.ValueMember = "TableName"
      cboTables.DataSource = mvMaintenanceTables
      If cboTables.Items.Count > 0 Then cboTables.SelectedIndex = 0
    End If
  End Sub

  Private Sub SetOptionTabs()
    Dim vCaption As String
    Dim vTabsPerPage As Integer

    pnlConAndOrg.Visible = False
    pnlPayment.Visible = False
    pnlAddressUpdate.Visible = False
    pnlTableImport.Visible = False
    pnlDocument.Visible = False
    pnlStock.Visible = False
    pnlBankTransactions.Visible = False

    Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType")
      Case 0  'ditContactOrganisation
        vTabsPerPage = 2
        vCaption = "Contact and Organisation Import Options"
        pnlConAndOrg.Visible = True
      Case 8  'ditFinancialHistory
        vTabsPerPage = 2
        vCaption = "Payment Import Options"
        pnlPayment.Visible = True
      Case 9  'ditAddressUpdate
        vTabsPerPage = 2
        vCaption = "Address Update Import Options"
        pnlAddrUpdate.Visible = True
      Case 12 'ditTableImport
        vTabsPerPage = 2
        vCaption = "Table Import Options"
        pnlTableImport.Visible = True
      Case 3  'ditCommsLog
        vTabsPerPage = 2
        vCaption = "Document Import Options"
        pnlDocument.Visible = True
      Case 14 'ditStock
        vTabsPerPage = 2
        vCaption = "Stock Import Options"
        pnlStock.Visible = True
      Case 16 'ditEventBookingAndDelegates
        vTabsPerPage = 2
        vCaption = "Event Booking and Delegates Import Options"
        pnlPayment.Visible = True
        DisableCheckBox(chkGiftAidRecords)
        DisableCheckBox(chkProcessIncentives)
        DisableCheckBox(chkMatchSchPayment)
        lblNumberOfDays.Enabled = False
      Case 19 'ditStockditBankTransactions
        vTabsPerPage = 2
        vCaption = ControlText.FrmImportBankTransactionsOptions 'Bank Transactions Import Options
        pnlBankTransactions.Visible = True
      Case 6 ' ditMailingHistory BR19057
        vTabsPerPage = 1
        vCaption = ""
        chkCreateCMD.Checked = False
        chkCMDSupp.Checked = False
      Case Else
        vTabsPerPage = 1
        vCaption = ""
    End Select
    If mvMultipleDataImportRuns Then
      vTabsPerPage = vTabsPerPage + 1
      mvLoadedMultiple = True
    Else
      If vTabsPerPage = 2 Then vTabsPerPage = 3
      mvLoadedMultiple = False
    End If

    If vTabsPerPage = 1 Then
      tabSub.TabPages.Remove(tbpMultImpRuns)
      tabSub.TabPages.Remove(tbpCustomOpt)
    Else
      If mvMultipleDataImportRuns Then
        'Display multiple import run options
        If Not tabSub.TabPages.Contains(tbpMultImpRuns) Then tabSub.TabPages.Add(tbpMultImpRuns)
        'if there are only 2 tabs then remove the custom options tab that may have been displayed for one of the
        'import types contained in the def file
        If vTabsPerPage > 2 Then
          If Not tabSub.TabPages.Contains(tbpCustomOpt) Then tabSub.TabPages.Add(tbpCustomOpt)
        Else
          tabSub.TabPages.Remove(tbpCustomOpt)
        End If
      Else
        tabSub.TabPages.Remove(tbpMultImpRuns)
        If Not tabSub.TabPages.Contains(tbpCustomOpt) Then tabSub.TabPages.Add(tbpCustomOpt)
      End If
    End If
    tbpCustomOpt.Text = vCaption
    tabSub.SelectedTab = tbpGeneralOpt
    tabSub.SetItemSizes()
  End Sub

  Private Sub SetImportOptions()
    EnableDeDupOptions(GetBooleanValue(IMPORT_OPTIONS, "EnableDeDupOptions"))
    EnableDeDup(GetBooleanValue(IMPORT_OPTIONS, "EnableDeDup"), GetBooleanValue(IMPORT_OPTIONS, "EnableDeDupAddress"))
    EnableAddressUpdateOptions(GetBooleanValue(IMPORT_OPTIONS, "EnableAddressUpdateOptions"))
    EnableConOrgOptions(GetBooleanValue(IMPORT_OPTIONS, "EnableConOrgOptions"))
    EnableEmployeeUpdate(GetBooleanValue(IMPORT_OPTIONS, "EnableEmployeeUpdate"))
    ShowEmployeeUpdate(GetBooleanValue(IMPORT_OPTIONS, "ShowEmployeeUpdate"))
    EnableFinancialHistoryOptions(GetBooleanValue(IMPORT_OPTIONS, "EnableFinancialHistoryOptions"))
    chkSoundexDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "SoundexDedup")
    optPaymentsFinHistory.Checked = (GetIntegerValue(DATA_IMPORT, "PaymentImportType") = 0)
    If GetBooleanValue(IMPORT_OPTIONS, "EnableUpdateOptions") Then EnableUpdateOptions()
    chkValCodes.Enabled = GetBooleanValue(IMPORT_OPTIONS, "EnableValCodes")
    chkValCodes.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ValidateCodes")
    chkCreateCMD.Enabled = GetBooleanValue(IMPORT_OPTIONS, "EnableCreateCMD")
    chkLogDups.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogDups")
    If GetBooleanValue(IMPORT_OPTIONS, "EnableLogDeDupAudit") Then
      chkLogDedupAudit.Enabled = True
      chkLogDedupAudit.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogDedupAudit")
    End If
    If GetBooleanValue(IMPORT_OPTIONS, "EnableLogDup") Then
      chkLogDups.Enabled = True
      chkLogDups.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogDups")
    End If
    chkExtRefDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ExtRefDeDup")
    chkForeInitDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ForeInitDeDup")
    chkTitleDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "TitleDeDup")
    chkNumberDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "NumberDeDup")
    chkAddressDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AddressDedup")
    chkDear.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Dear")
    chkDefSupp.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DefaultSupp")
    chkProcessIncentives.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ProcessIncentives")
    chkGiftAidRecords.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UnclaimedGiftAidRecords")
    chkMatchSchPayment.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "MatchScheduledPayment")
    chkLogConversion.Enabled = GetBooleanValue(IMPORT_OPTIONS, "EnableLogConversions")
    If chkLogConversion.Enabled Then
      chkLogConversion.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogConversions")
    Else
      chkLogConversion.Checked = False
    End If
    If Not GetBooleanValue(DATA_IMPORT_PARAMS, "ProcessIncentives") Then DisableCheckBox(chkProcessIncentives)
    If Not GetBooleanValue(DATA_IMPORT_PARAMS, "MatchScheduledPayment") Then DisableCheckBox(chkMatchSchPayment)
    lblNumberOfDays.Enabled = GetBooleanValue(IMPORT_OPTIONS, "EnableNumberOfDays")

    If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 12 Then 'dimTableImport
      If GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad") AndAlso GetValue(DATA_IMPORT_PARAMS, "DefFileName").Length > 0 Then
        SelectComboBoxItem(cboTables, GetValue(DATA_IMPORT_PARAMS, "TableImportTable")) '  txtTable = vTable
        SetValue(DATA_IMPORT_PARAMS, "FirstLoad", False)
      End If
    End If
    chkDASImport.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DASImport")
  End Sub

  ''' <summary>
  ''' Updates the value in the dataset and adds an entry in the Changed values collection
  ''' to indicate that the value has changed since the first time it was loaded
  ''' </summary>
  ''' <param name="pTableName"></param>
  ''' <param name="pField"></param>
  ''' <param name="pValue"></param>
  ''' <param name="pDataSet"></param>
  ''' <remarks></remarks>
  Private Sub SetValue(ByVal pTableName As String, ByVal pField As String, ByVal pValue As Object, Optional ByVal pDataSet As String = MAIN)
    Dim vDataSet As DataSet = Nothing
    If pDataSet = MAIN Then
      vDataSet = DataImportDS
    Else
      vDataSet = mvMasterDataImport
    End If

    If TypeOf pValue Is Boolean Then pValue = CBoolYN(CBool(pValue))
    If vDataSet IsNot Nothing Then
      Dim vCurrentValue As String = Nothing
      Select Case pTableName
        Case DATA_IMPORT
          vCurrentValue = vDataSet.Tables(DATA_IMPORT).Rows(0)(pField).ToString
          vDataSet.Tables(DATA_IMPORT).Rows(0)(pField) = pValue
        Case DATA_IMPORT_PARAMS
          vCurrentValue = vDataSet.Tables(DATA_IMPORT_PARAMS).Rows(0)(pField).ToString
          vDataSet.Tables(DATA_IMPORT_PARAMS).Rows(0)(pField) = pValue
      End Select
      If vCurrentValue <> pValue.ToString Then SetValueChanged(pTableName, pField, pDataSet)
    End If
  End Sub

  Private Sub SetValueChanged(ByVal pTableName As String, ByVal pField As String, ByVal pDataSet As String)
    'Keeps track of the values that have changed since the last time
    'they were loaded from the server
    Dim vTableName As String = pTableName
    If mvImportForm IsNot Nothing AndAlso pDataSet = MAIN Then vTableName = String.Format("MAIN_{0}", pTableName)
    If mvChangedValues.ContainsKey(vTableName) Then
      If Not mvChangedValues(vTableName).Contains(pField) Then mvChangedValues(vTableName).Add(pField)
    Else
      mvChangedValues.Add(vTableName, New List(Of String))
      mvChangedValues(vTableName).Add(pField)
    End If
  End Sub

  Private Sub chkActivity_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkActivity.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Activity", chkActivity.Checked)
  End Sub

  Private Sub chkNameGatheringIncentives_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkNameGatheringIncentives.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "GenerateNameGatheringIncentives", chkNameGatheringIncentives.Checked)
  End Sub

  Private Sub chkAddPosition_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddPosition.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "AddPosition", chkAddPosition.Checked)
  End Sub

  Private Sub chkAddressDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddressDedup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "AddressDeDup", chkAddressDedup.Checked)
  End Sub

  Private Sub chkAddTransactions_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddTransactions.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "AddTransactions", chkAddTransactions.Checked)
  End Sub

  Private Sub chkBankDetailsDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkBankDetailsDedup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "BankDetailsDedup", chkBankDetailsDedup.Checked)
  End Sub

  Private Sub chkCacheMailsort_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCacheMailsort.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CacheMailsortData", chkCacheMailsort.Checked)
  End Sub

  Private Sub chkCacheMailsortAddr_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCacheMailsortAddr.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CacheMailsortData", chkCacheMailsortAddr.Checked)
  End Sub

  Private Sub chkCaps_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCaps.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Caps", chkCaps.Checked)
  End Sub

  Private Sub chkCMDSupp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCMDSupp.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CreateCMDForWarningSupp", chkCMDSupp.Checked)
  End Sub

  Private Sub chkControlNumbers_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkControlNumbers.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ControlNumbers", chkControlNumbers.Checked)
  End Sub

  Private Sub chkCreateAct_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCreateAct.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CreateActivityForProduct", chkCreateAct.Checked)
  End Sub

  Private Sub chkCreateCMD_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCreateCMD.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CreateCMD", chkCreateCMD.Checked)
    If chkCreateCMD.Checked Then
      chkCMDSupp.Enabled = True
    Else
      chkCMDSupp.Enabled = False
      chkCMDSupp.Checked = False
    End If
  End Sub

  Private Sub chkCreateGridRefs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCreateGridRefs.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "CreateGridReferences", chkCreateGridRefs.Checked)
  End Sub

  Private Sub chkCtrlNo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkCtrlNo.CheckedChanged
    txtLookupDefValue.Text = String.Empty
    txtDefValue.Text = String.Empty
    txtLookupDefValue.Visible = chkCtrlNo.Checked
    txtDefValue.Visible = Not chkCtrlNo.Checked
    If chkCtrlNo.Checked Then chkIncPerLine.Checked = False
  End Sub

  Private Sub chkDear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkDear.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Dear", chkDear.Checked)
  End Sub

  Private Sub chkDefAddrFromUnknown_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkDefAddrFromUnknown.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "DefaultAddrFromUnknown", chkDefAddrFromUnknown.Checked)
  End Sub

  Private Sub chkDefSupp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkDefSupp.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "DefaultSupp", chkDefSupp.Checked)
  End Sub

  Private Sub chkDupAsError_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkDupAsError.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LoadDuplicates", chkDupAsError.Checked)
  End Sub

  Private Sub chkEmailDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkEmailDedup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "EmailDedup", chkEmailDedup.Checked)
  End Sub

  Private Sub chkEmployee_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkEmployee.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "EmployeeLoad", chkEmployee.Checked)
    If chkEmployee.Checked Then
      DisableCheckBox(chkUpdate)
      DisableCheckBox(chkUpdateAll)
      txtOrgNumber.Enabled = True
      txtOrgNumber.Focus()
    Else
      txtOrgNumber.Enabled = False
      chkUpdate.Enabled = True
      chkUpdateAll.Enabled = True
    End If
  End Sub

  Private Sub chkExtractAddr_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkExtractAddr.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ExtractAddress", chkExtractAddr.Checked)
  End Sub

  Private Sub chkExtRefDeDup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkExtRefDeDup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ExtRefDeDup", chkExtRefDeDup.Checked)
  End Sub

  Private Sub chkForeInitDeDup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkForeInitDeDup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ForeInitDeDup", chkForeInitDeDup.Checked)
  End Sub

  Private Sub chkGiftAidRecords_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkGiftAidRecords.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UnclaimedGiftAidRecords", chkGiftAidRecords.Checked)
  End Sub

  Private Sub chkIgnore_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkIgnore.CheckedChanged
    If mvImportForm Is Nothing Then
      SetValue(DATA_IMPORT_PARAMS, "IgnoreFirstRow", chkIgnore.Checked)
    Else
      mvImportForm.MappedAttribute("MapIgnoreFirstRow") = CBoolYN(chkIgnore.Checked)
    End If
  End Sub

  Private Sub chkIncPerLine_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkIncPerLine.CheckedChanged
    txtLookupDefValue.Text = String.Empty
    If chkIncPerLine.Checked Then chkCtrlNo.Checked = False
  End Sub

  Private Sub chkLogConversion_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkLogConversion.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LogConversions", chkLogConversion.Checked)
  End Sub

  Private Sub chkLogCreate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkLogCreate.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LogCreate", chkLogCreate.Checked)
  End Sub

  Private Sub chkLogDups_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkLogDups.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LogDups", chkLogDups.Checked)
  End Sub

  Private Sub chkLogDedupAudit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkLogDedupAudit.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LogDedupAudit", chkLogDedupAudit.Checked)
  End Sub

  Private Sub chkLogWarn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkLogWarn.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "LogWarnings", chkLogWarn.Checked)
  End Sub

  Private Sub chkEmptyBeforeImport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkEmptyBeforeImport.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "EmptyBeforeImport", chkEmptyBeforeImport.Checked)
  End Sub

  Private Sub chkMatchSchPayment_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkMatchSchPayment.CheckedChanged
    If chkMatchSchPayment.Checked = False Then txtNumberOfDays.Text = String.Empty
    txtNumberOfDays.Enabled = chkMatchSchPayment.Checked
    lblNumberOfDays.Enabled = chkMatchSchPayment.Checked
    SetValue(DATA_IMPORT_PARAMS, "MatchScheduledPayment", chkMatchSchPayment.Checked)
  End Sub

  Private Sub chkNoFromFile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkNoFromFile.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UseNumbersFromFile", chkNoFromFile.Checked)
    If chkNoFromFile.Checked Then
      DisableCheckBox(chkProcessIncentives)
      chkAddTransactions.Enabled = True
      chkProcessIncentives.Checked = False
    Else
      chkAddTransactions.Checked = False
      DisableCheckBox(chkAddTransactions)
      If GetIntegerValue(DATA_IMPORT, "PaymentImportType") = 1 OrElse GetIntegerValue(DATA_IMPORT, "PaymentImportType") = 3 Then chkProcessIncentives.Enabled = True
    End If
  End Sub

  Private Sub chkNoIndexes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkNoIndexes.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "RemoveIndexes", chkNoIndexes.Checked)
  End Sub

  Private Sub chkAmendmentHistory_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAmendmentHistory.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "AmendmentHistory", chkAmendmentHistory.Checked)
  End Sub

  Private Sub chkNumberDeDup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkNumberDeDup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "NumberDeDup", chkNumberDeDup.Checked)
  End Sub

  Private Sub chkOrgAddressPotDup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkOrgAddressPotDup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "OrgAddressPotDup", chkOrgAddressPotDup.Checked)
  End Sub

  Private Sub chkPAFAddress_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkPAFAddress.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "PAFAddress", chkPAFAddress.Checked)
    If chkPAFAddress.Checked Then
      chkRePostcode.Enabled = True
      chkCreateGridRefs.Enabled = AppValues.IsGridReferencesSupported
    Else
      DisableCheckBox(chkRePostcode)
      DisableCheckBox(chkCreateGridRefs)
    End If
  End Sub

  Private Sub chkProcessIncentives_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkProcessIncentives.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ProcessIncentives", chkProcessIncentives.Checked)
    If chkProcessIncentives.Checked Then
      DisableCheckBox(chkNoFromFile)
      DisableCheckBox(chkAddTransactions)
      If chkNoFromFile.Checked = True Then
        chkNoFromFile.Checked = False
        chkAddTransactions.Checked = False
      End If
    Else
      chkNoFromFile.Enabled = True
    End If
  End Sub

  Private Sub chkReference_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkReference.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "DefaultReferencetoBNTN", chkReference.Checked)
  End Sub

  Private Sub chkReplaceQuestionMark_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkReplaceQuestionMark.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ReplaceQuestionMark", chkReplaceQuestionMark.Checked)
    If chkReplaceQuestionMark.Checked Then
      txtReplaceQuestionMarkWith.Enabled = True
    Else
      txtReplaceQuestionMarkWith.Enabled = False
    End If
    txtReplaceQuestionMarkWith.Text = String.Empty
  End Sub

  Private Sub chkRePostcode_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkRePostcode.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "RePostcode", chkRePostcode.Checked)
  End Sub

  Private Sub chkSkipZeroAmt_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkSkipZeroAmt.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "SkipZeroAmounts", chkSkipZeroAmt.Checked)
  End Sub

  Private Sub chkSurnameFirst_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkSurnameFirst.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "SurnameFirst", chkSurnameFirst.Checked)
  End Sub

  Private Sub chkTitleDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkTitleDeDup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "TitleDeDup", chkTitleDeDup.Checked)
  End Sub

  Private Sub chkSoundexDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkSoundexDedup.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "SoundexDeDup", chkSoundexDedup.Checked)
  End Sub

  Private Sub chkUpdate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdate.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Update", chkUpdate.Checked)
    If chkUpdate.Checked Then
      chkUpdateAll.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "UpdateAll", "N")
      DisableCheckBox(chkEmployee)
    Else
      If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") <> 12 Then chkEmployee.Enabled = True '12 - ditTableImport
    End If
  End Sub

  Private Sub chkUpdateAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateAll.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UpdateAll", chkUpdateAll.Checked)
    If chkUpdateAll.Checked Then
      chkUpdate.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "Update", "N")
      DisableCheckBox(chkEmployee)
    Else
      If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") <> 12 Then chkEmployee.Enabled = True '12 - ditTableImport
    End If
  End Sub

  Private Sub chkUpdateAllDoc_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateAllDoc.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UpdateAll", chkUpdateAllDoc.Checked)
    If chkUpdateAllDoc.Checked Then
      chkUpdateDoc.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "Update", "N")
    End If
  End Sub

  Private Sub chkUpdateAllTableImport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateAllTableImport.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UpdateAll", chkUpdateAllTableImport.Checked)
    If chkUpdateAllTableImport.Checked = True Then
      chkUpdateTableImport.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "Update", "N")
    End If
  End Sub

  Private Sub chkUpdateDoc_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateDoc.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Update", chkUpdateDoc.Checked)
    If chkUpdateDoc.Checked Then
      chkUpdateAllDoc.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "UpdateAll", "N")
    End If
  End Sub

  Private Sub chkUpdateTableImport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateTableImport.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "Update", chkUpdateTableImport.Checked)
    If chkUpdateTableImport.Checked Then
      chkUpdateAllTableImport.Checked = False
      SetValue(DATA_IMPORT_PARAMS, "UpdateAll", "N")
    End If
  End Sub

  Private Sub chkUpdateWithNull_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkUpdateWithNull.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "UpdateWithNull", chkUpdateWithNull.Checked)
  End Sub

  Private Sub chkValCodes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkValCodes.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "ValidateCodes", chkValCodes.Checked)
  End Sub

  Private Sub chkDasImport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkDASImport.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "DASImport", chkDASImport.Checked)
  End Sub

  Private Sub InitTextLookupBox(ByVal pTextLookupBox As TextLookupBox, ByVal pTable As String, ByVal pAttr As String)
    Dim vParams As New ParameterList(True)
    Dim vList As New ParameterList(True)
    vParams("TableName") = pTable
    vParams("FieldName") = pAttr
    Dim vMaintAttrs As ParameterList = DataHelper.GetMaintenanceData(vParams)
    If pAttr = "honorifics" Then
      vMaintAttrs("ValidationTable") = "honorifics"
      vMaintAttrs("ValidationAttribute") = "honorific"
    End If
    Dim vPanelItem As New PanelItem(pTextLookupBox, pAttr)
    vPanelItem.InitFromMaintenanceData(vMaintAttrs)
    pTextLookupBox.ComboBox.DataSource = Nothing 'Clear prev datasource(if any) 
    pTextLookupBox.Tag = vPanelItem
    'BR 21289
    If pAttr.Equals("vat_rate", StringComparison.InvariantCultureIgnoreCase) Then
      vList("ForMaintenance") = "Y"
      pTextLookupBox.Init(vPanelItem, True, False)
      pTextLookupBox.FillComboWithRestriction(vList)
    Else
      pTextLookupBox.Init(vPanelItem, False, False)
    End If
  End Sub

  Private Sub frmImport_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Try
      'Unlock the parent form
      If mvImportForm IsNot Nothing Then mvImportForm.Enabled = True
      If mvTaskInfo IsNot Nothing Then mvTaskInfo.StopTimer()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub frmImport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If mvImportForm Is Nothing Then
        cboType.DisplayMember = "ImportTypeDesc"
        cboType.ValueMember = "ImportType"
        cboType.DataSource = DataHelper.GetImportTypes(New ParameterList(True))
      End If

      mvValidateHonorifics = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_validate_honorifics)

      If mvImportForm Is Nothing Then
        'set up the source/data_source text boxes
        InitTextLookupBox(txtSource, "contacts", "source")
        InitTextLookupBox(txtDataSource, "contact_external_links", "data_source")
      End If

      'for the default tab
      dtpckValue.Visible = False
      txtLookupDefValue.Visible = True
      cboPatternValue.Visible = False
      HideTableControls()
      If mvImportForm Is Nothing Then
        lblKey.Visible = False
        chkKey.Visible = False
        cmdOk.Enabled = True
        cmdTest.Visible = True
        lblMapValue.Visible = False
        optMapValueNull.Visible = False
        optMapValueLookup.Visible = False
      Else
        lblKey.Visible = True
        chkKey.Visible = True
        cmdOk.Enabled = False
        cmdTest.Visible = False
        lblMapValue.Visible = True
        optMapValueNull.Visible = True
        optMapValueLookup.Visible = True
      End If
      SetControlNumAccess()
      ResetAll()

      mvIsPAFSupported = AppValues.IsPAFSupported

      'Set up the data import form after reading the import file
      SetImportFileName()
      DoCount()

      If mvImportForm Is Nothing Then
        SetControls(True)
        SetImportOptions()
        SetOptionTabs()
      Else
        mvImportForm.MappedAttribute("FileName") = mvImportFilename
        mvImportForm.MappedAttribute("MapNoOfCols") = dgr.ColumnCount
      End If

      If cboSeparator.Text = "Fixed" Then DisableCheckBox(chkIgnore)

      AddHandler cboSeparator.SelectedIndexChanged, AddressOf cboSeparator_SelectedIndexChanged
      AddHandler cboType.SelectedIndexChanged, AddressOf cboType_SelectedIndexChanged
      AddHandler cboTables.SelectedIndexChanged, AddressOf cboTables_SelectedIndexChanged

      If dgr.ColumnCount > 0 Then dgr.SelectColumn(dgr.ColumnName(0))

    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enUNCPathOnly
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enWriteAccessDenied
          ShowInformationMessage(vEx.Message, mvImportFilename)
        Case Else
          DataHelper.HandleException(vEx)
      End Select
      'BR14995 Can't do me.close in form load
      'Me.Close() 'Close form as it has not loaded properly
      Me.BeginInvoke(New MethodInvoker(AddressOf CloseIt))
    End Try
  End Sub

  Private Sub CloseIt()
    'BR14995
    Me.Close()
  End Sub
  Private Sub AttrsAdd(ByVal pItem As String, ByVal pIndex As String, pDeferRefresh As Boolean)
    Dim vDataSet As String = MAIN
    If mvImportForm IsNot Nothing Then vDataSet = MASTER

    Dim vRow As DataRow = SelectRow(vDataSet, DATA_IMPORT_ATTRS, "ID = '{0}'", pIndex)

    If vRow("AttributeName").ToString.Length > 0 Then  'This will prevent empty validation items from being added to the list
      mvAttrItems.Add(New LookupItem(pIndex, pItem))
      'mvDefAttrItems.Add(New LookupItem(pIndex, pItem))
      AttrsSort()
      If Not pDeferRefresh Then
        AttrsRefresh()
      End If
    End If
  End Sub

  Private Sub AttrsRemove(ByVal pIndex As Integer, ByVal pDefault As Boolean, pDeferFresh As Boolean)
    If pIndex >= 0 Then
      If Not pDefault AndAlso pIndex + 1 >= mvAttrItems.Count Then cboAttrs.DroppedDown = False 'BR16045: This will prevent InvalidArgument error when coming from keypress
      mvAttrItems.RemoveAt(pIndex)
      'mvDefAttrItems.RemoveAt(pIndex)
      If Not pDeferFresh Then
        AttrsRefresh()
      End If
    End If
  End Sub

  Private Sub AttrsClear()
    mvAttrItems.Clear()
    'mvDefAttrItems.Clear()
    AttrsRefresh()
  End Sub

  Private Sub AttrsRefresh()
    'utter rubbish from WinForms, you have to reset the data source to update the bindings
    'TODO change the binding to a DataTable instead of a collection
    RefreshingAttributes = True

    Dim vSelectedAttr As Object = cboAttrs.SelectedItem
    cboAttrs.DataSource = Nothing
    cboAttrs.DataSource = mvAttrItems
    cboAttrs.SelectedItem = vSelectedAttr

    Dim vSelectedDefAttr As Object = cboDefAttrs.SelectedItem
    cboDefAttrs.DataSource = Nothing
    cboDefAttrs.DataSource = mvAttrItems
    cboDefAttrs.SelectedItem = vSelectedDefAttr

    RefreshingAttributes = False
  End Sub

  Private Sub EnableConOrgOptions(ByVal pEnabled As Boolean)
    If pEnabled Then
      EnableDeDupOptions(optDedupNone.Checked = False)
      If mvIsPAFSupported Then
        chkPAFAddress.Enabled = pEnabled
        If chkPAFAddress.Checked Then
          chkRePostcode.Enabled = True
          chkCreateGridRefs.Enabled = AppValues.IsGridReferencesSupported
        Else
          DisableCheckBox(chkRePostcode)
          DisableCheckBox(chkCreateGridRefs)
        End If
      Else
        DisableCheckBox(chkPAFAddress)
        DisableCheckBox(chkRePostcode)
        DisableCheckBox(chkCreateGridRefs)
      End If
    Else
      EnableDeDupOptions(pEnabled)
      DisableCheckBox(chkPAFAddress)
      DisableCheckBox(chkRePostcode)
      DisableCheckBox(chkCreateGridRefs)
    End If
    lblUpdateSub.Enabled = pEnabled
    chkUpdate.Enabled = pEnabled
    chkUpdateAll.Enabled = pEnabled
    chkActivity.Enabled = pEnabled
    chkNameGatheringIncentives.Enabled = pEnabled
    chkDear.Enabled = pEnabled
    chkCaps.Enabled = pEnabled
    chkLogDups.Enabled = pEnabled
    chkLogDedupAudit.Enabled = pEnabled
    If IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) <> 5 Then  'BR18235-Do Not overwrite CheckBox values for Suppressions Import.
      If Not chkLogDups.Enabled Then chkLogDups.Checked = False
      If Not chkLogDedupAudit.Enabled Then chkLogDedupAudit.Checked = False
    End If
    chkSurnameFirst.Enabled = pEnabled
    chkNoIndexes.Enabled = GetBooleanValue(DATA_IMPORT, "EnableRemoveIndexes")
    chkAmendmentHistory.Enabled = GetBooleanValue(DATA_IMPORT, "EnableAmendmentHistory")
    chkCacheMailsort.Enabled = pEnabled
  End Sub

  Private Sub EnableDeDup(ByVal pEnabled As Boolean, ByVal pEnableAddress As Boolean)
    optDedupFull.Enabled = pEnabled
    optDedupAddressOnly.Enabled = pEnabled
    optDedupNone.Enabled = pEnabled
    'select a default option if nothing is selected
    If ((Not pEnableAddress) AndAlso (Not optDedupNone.Checked)) Then optDedupFull.Checked = True
    optDedupAddressOnly.Enabled = CBool(IIf(pEnabled, pEnableAddress, False))
  End Sub

  Private Sub EnableDeDupOptions(ByVal pEnabled As Boolean)
    chkExtRefDeDup.Enabled = pEnabled
    chkNumberDeDup.Enabled = pEnabled
    chkForeInitDeDup.Enabled = pEnabled
    chkTitleDeDup.Enabled = pEnabled
    chkSoundexDedup.Enabled = pEnabled
    chkOrgAddressPotDup.Enabled = pEnabled
    chkAddressDedup.Enabled = pEnabled
    chkBankDetailsDedup.Enabled = pEnabled
    chkEmailDedup.Enabled = pEnabled
  End Sub

  Private Sub EnableEmployeeUpdate(ByVal pEnabled As Boolean)
    chkEmployee.Enabled = pEnabled
    txtOrgNumber.Enabled = pEnabled
  End Sub

  Private Sub InitForm()
    If GetValue(DATA_IMPORT_PARAMS, "DefFileName").Length > 0 Then mvInitializingFromDef = True
    InitFormParameters()
    InitFormGrid()
    InitFormDefault()
    mvInitializingFromDef = False
  End Sub

  Private Sub InitFormDefault()

    dgrDefaults.Clear()
    dgrDefaults.Populate(DefaultsDS)

    'Remove attributes that have defaults from the Attributes and Default Attributes Combo lists
    For vRow As Integer = 0 To dgrDefaults.RowCount - 1
      With DataImportDS.Tables(DEFAULTS_ROW)
        If .Rows(vRow)("Attribute").ToString.Length > 0 Then
          'BR17268 - Rework clearing items from attribute combo boxes. Removing attributes no longer causes an exception.
          'Dim vAttr As LookupItem = New LookupItem(.Rows(vRow)("ID").ToString, .Rows(vRow)("Attribute").ToString)
          'cboAttrs.Items.Remove(vAttr)
          'cboDefAttrs.Items.Remove(vAttr)
          Dim vSearchFunction As New Func(Of LookupItem, Boolean)(Function(vItem) vItem.LookupCode = .Rows(vRow)("ID").ToString() AndAlso vItem.LookupDesc = .Rows(vRow)("Attribute").ToString())
          Dim vAttr As LookupItem = mvAttrItems.FirstOrDefault(vSearchFunction)
          If vAttr IsNot Nothing Then mvAttrItems.Remove(vAttr)
        End If
      End With
    Next
  End Sub

  Private Function GetValue(ByVal pTable As String, ByVal pField As String, Optional ByVal pDataSet As String = MAIN) As String
    Dim vDataSet As DataSet = DataImportDS
    If pDataSet = MASTER Then vDataSet = mvMasterDataImport
    Dim vVal As String = String.Empty
    If vDataSet IsNot Nothing Then vVal = vDataSet.Tables(pTable).Rows(0)(pField).ToString()
    Return vVal
  End Function

  Private Function GetIntegerValue(ByVal pTable As String, ByVal pField As String, Optional ByVal pDataSet As String = MAIN) As Integer
    Return IntegerValue(GetValue(pTable, pField, pDataSet))
  End Function

  Private Function GetBooleanValue(ByVal pTable As String, ByVal pField As String, Optional ByVal pDataSet As String = MAIN) As Boolean
    Return BooleanValue(GetValue(pTable, pField, pDataSet))
  End Function

  Private Sub InitFormParameters(Optional ByVal pGetImportType As Boolean = True)
    'If you add any more separators, You need to amend Public Property Separators
    RemoveHandler cboSeparator.SelectedIndexChanged, AddressOf cboSeparator_SelectedIndexChanged
    cboSeparator.DataSource = GetValue(DATA_IMPORT, "Separators").Split("\"c)

    mvNoRead = True
    If pGetImportType Then
      SelectComboBoxItem(cboType, GetValue(DATA_IMPORT_PARAMS, "DataImportType"))
      mvSelectedImportType = GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType")
      mvSelectedImportTypeDesc = cboType.Text
    End If

    If mvImportForm Is Nothing Then
      cboSeparator.Text = GetValue(DATA_IMPORT_PARAMS, "Separator")
    Else
      cboSeparator.Text = GetValue(DATA_IMPORT, "MapSeparator", MASTER)
    End If
    mvNoRead = False
    txtSource.Text = GetValue(DATA_IMPORT_PARAMS, "Source")
    txtDataSource.Text = GetValue(DATA_IMPORT_PARAMS, "DataSource")
    chkEmptyBeforeImport.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "EmptyBeforeImport")
    chkExtractAddr.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ExtractAddress")
    If mvImportForm Is Nothing Then
      chkIgnore.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "IgnoreFirstRow")
    Else
      'chkIgnore.Checked = GetBooleanValue(DATA_IMPORT, "MapIgnoreFirstRow", MASTER)
      'If GetIntegerValue(DATA_IMPORT, "MapReturnValueType", MASTER) = 0 Then
      chkIgnore.Checked = BooleanValue(mvImportForm.MappedAttribute("MapIgnoreFirstRow").ToString)
      If IntegerValue(mvImportForm.MappedAttribute("MapReturnValueType")) = 0 Then
        optMapValueNull.Checked = True
      Else
        optMapValueLookup.Checked = True
      End If
    End If
    Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "Dedup")
      Case 0
        optDedupFull.Checked = True
      Case 1
        optDedupAddressOnly.Checked = True
      Case 2
        optDedupNone.Checked = True
    End Select
    SetOptionTabDefaults()
    chkEmployee.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "EmployeeLoad")
    txtOrgNumber.Text = GetValue(DATA_IMPORT_PARAMS, "OrgNumber")
    chkExtRefDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ExtRefDeDup")
    chkNumberDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "NumberDeDup")
    chkTitleDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "TitleDeDup")
    chkForeInitDeDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ForeInitDeDup")
    chkAddressDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AddressDeDup")
    chkBankDetailsDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "BankDetailsDedup")
    chkEmailDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "EmailDedup")
    chkSoundexDedup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "SoundexDeDup")
    chkOrgAddressPotDup.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "OrgAddressPotDup")
    chkOrgNamePostCodeAddress.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "OrgNamePostCodeAddressDup")
    chkExclUnkAdd.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DeDupExclUnkAdd")
    Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType")
      Case 0 'ditContactOrganisation
        chkUpdate.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Update")
        chkUpdateAll.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UpdateAll")
        chkUpdateWithNull.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UpdateWithNull")
        chkCacheMailsort.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CacheMailsortData")
      Case 12 'ditTableImport
        chkUpdateTableImport.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Update")
        chkUpdateAllTableImport.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UpdateAll")
      Case 3 'ditCommsLog
        chkUpdateDoc.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Update")
        chkUpdateAllDoc.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UpdateAll")
      Case 9 'ditAddressUpdate
        chkCacheMailsortAddr.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CacheMailsortData")
    End Select
    chkActivity.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Activity")
    chkNameGatheringIncentives.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "GenerateNameGatheringIncentives")
    chkDear.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Dear")
    chkDefSupp.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DefaultSupp")
    chkDefAddrFromUnknown.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DefaultAddrFromUnknown")
    chkCaps.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "Caps")
    chkSurnameFirst.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "SurnameFirst")
    chkAddPosition.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AddPosition")
    chkPAFAddress.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "PAFAddress")
    chkRePostcode.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "RePostcode")
    If IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) = 8 OrElse IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) = 16 Then 'If .Parameters.DataImportType = ditFinancialHistory Or .Parameters.DataImportType = ditEventBookingAndDelegates Then
      Select Case GetIntegerValue(DATA_IMPORT, "PaymentImportType")
        Case 0
          optPaymentsFinHistory.Checked = True
        Case 1
          optPaymentsPostedToCB.Checked = True
        Case 2
          optPaymentsPostedToNominal.Checked = True
        Case 3
          optPaymentsUnposted.Checked = True
      End Select
      If IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) = 16 Then 'If .Parameters.DataImportType = ditEventBookingAndDelegates Then
        DisableCheckBox(chkGiftAidRecords)
        chkNoFromFile.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UseNumbersFromFile")
        chkAddTransactions.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AddTransactions")
        chkReference.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DefaultReferencetoBNTN")
        DisableCheckBox(chkProcessIncentives)
        DisableCheckBox(chkMatchSchPayment)
        lblNumberOfDays.Enabled = False
      Else
        chkGiftAidRecords.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UnclaimedGiftAidRecords")
        chkNoFromFile.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "UseNumbersFromFile")
        chkAddTransactions.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AddTransactions")
        chkReference.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DefaultReferencetoBNTN")
        If GetBooleanValue(DATA_IMPORT_PARAMS, "ProcessIncentives") AndAlso Not GetBooleanValue(DATA_IMPORT_PARAMS, "UseNumbersFromFile") AndAlso
        (GetIntegerValue(DATA_IMPORT, "PaymentImportType") = 1 OrElse GetIntegerValue(DATA_IMPORT, "PaymentImportType") = 3) Then
          chkProcessIncentives.Checked = True
        Else
          chkProcessIncentives.Checked = False
        End If
        chkMatchSchPayment.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "MatchScheduledPayment")
        txtNumberOfDays.Text = GetValue(DATA_IMPORT_PARAMS, "ScheduledPmtNoOfDays")
      End If
      chkSkipZeroAmt.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "SkipZeroAmounts")
      chkCreateAct.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CreateActivityForProduct")
    ElseIf AppValues.DefaultCountryCode <> "UK" Then
      DisableCheckBox(chkDear)
    End If

    If mvImportForm Is Nothing Then
      AddAttributesFromDataSet()
    End If

    Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "StockUpdateType")
      Case 0
        optStockUpdate.Checked = True
      Case 1
        optStockSet.Checked = True
    End Select

    If pGetImportType Then chkDupAsError.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LoadDuplicates")
    chkCreateGridRefs.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CreateGridReferences")
    'mvDataImport.Parameters.UserName = gvEnv.User.Logname 
    If Not mvIsPAFSupported Then chkPAFAddress.Checked = False
    chkDASImport.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "DASImport")
    chkAllowBlankForOrgName.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AllowBlankOrganisation")
  End Sub

  Private Sub AddAttributesFromDataSet()
    AttrsClear()
    If DataImportDS.Tables.Contains(DATA_IMPORT_ATTRS) Then
      Dim vDefaultsTable As DataTable = New DataTable
      vDefaultsTable.Columns.Add("ID", GetType(String))
      If DefaultsDS.Tables.Contains("DataRow") Then
        vDefaultsTable = DefaultsDS.Tables("DataRow")
      End If

      For Each vRow As DataRow In DataImportDS.Tables(DATA_IMPORT_ATTRS).Rows
        Dim vCanAdd As Boolean = True
        Dim vID As String = vRow("ID").ToString()
        'Allow the attribute if it's not already specified as a default
        vCanAdd = Not vDefaultsTable.Rows.Cast(Of DataRow).Any(Function(vItem) vItem("ID").ToString() = vID)
        If vCanAdd Then AttrsAdd(vRow("AttributeNameDesc").ToString, vID, True)
      Next
      AttrsRefresh()
    End If
  End Sub

  Private Sub InitFormGrid()
    Dim vTable As DataTable = Nothing
    Dim vRow As DataRow = Nothing

    dgr.MaxGridRows = GetIntegerValue(DATA_IMPORT, "ImportFileRows")

    'Fill the grid with the records
    dgr.Populate(DataImportDS.Tables(DATA_IMPORT_FILE))
    dgr.AdjustColumnHeaders()

    If mvImportForm Is Nothing Then
      'Set column headings for previously mapped columns
      vTable = DataImportDS.Tables(MAP_ATTR_COLS)
      If vTable IsNot Nothing Then
        For Each vRow In vTable.Rows
          dgr.ColumnHeading(IntegerValue(vRow("ColumnNumber"))) = vRow("Heading").ToString
        Next
      End If

      'remove mapped attributes from available list
      vTable = DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS)
      If vTable IsNot Nothing Then
        Dim vAttributeColRow As DataRow
        For Each vRow In vTable.Rows
          RemoveFromList(vRow("ID").ToString)
          'All of the existing mapped attribute columns should exist in AttributeColumns table with ColumnIndex set to -1
          If Not DataImportDS.Tables.Contains(ATTRIBUTE_COL) Then DataImportDS.Tables.Add(CreateTable(ATTRIBUTE_COL))
          vAttributeColRow = DataImportDS.Tables(ATTRIBUTE_COL).NewRow()
          vAttributeColRow("AttributeIndex") = vRow("ID")
          vAttributeColRow("AttributeDesc") = SelectRow("mvDataImport", DATA_IMPORT_ATTRS, "ID = '{0}'", vRow("ID").ToString)("AttributeNameDesc")
          vAttributeColRow("ColumnIndex") = -1
          DataImportDS.Tables(ATTRIBUTE_COL).Rows.Add(vAttributeColRow)
        Next
      End If
      'Set the columns mapping for normal attributes
      vTable = DataImportDS.Tables(ATTRIBUTE_COL)
      If vTable IsNot Nothing Then
        Dim vAttrIndex As String
        Dim vAttr As DataRow = Nothing
        Dim vDuplicateColumns As List(Of Integer) = ListDuplicateColumnIndex(vTable)
        For vIndex As Integer = 0 To vTable.Rows.Count - 1
          If vIndex < vTable.Rows.Count Then
            If IntegerValue(vTable.Rows(vIndex)("ColumnIndex")) >= 0 Then
              vAttrIndex = vTable.Rows(vIndex)("AttributeIndex").ToString
              vAttr = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vAttrIndex)
              dgr.ActiveColumn = IntegerValue(vTable.Rows(vIndex)("ColumnIndex"))
              SelectComboBoxItem(cboAttrs, vAttrIndex, True)
              If vIndex > 0 AndAlso vAttrIndex = vTable.Rows(vIndex - 1)("AttributeIndex").ToString Then
                'A preceeding vTable row has been deleted via SelectComboxItem() and so adjust vIndex
                vIndex -= 1
              End If
            End If
          End If
        Next
        'Now remove mapping for columns with shared ColumnIndex value - it is not possible to uniquely identify these columns
        If (Not vDuplicateColumns Is Nothing) AndAlso vDuplicateColumns.Count > 0 Then
          For Each vDuplicateColumnIndex In vDuplicateColumns
            dgr.ActiveColumn = vDuplicateColumnIndex
            RemoveAttributeMapping(True)
          Next
          AttrsRefresh()
        End If
      End If
    Else
      'Check if this file already has a column mapped
      If mvImportForm.MappedAttribute("MappedColOfFile").ToString.Length > 0 Then
        dgr.ColumnHeading(IntegerValue(mvImportForm.MappedAttribute("MappedColOfFile"))) = "Map:" & mvImportForm.MappedAttribute("FileName").ToString
        DisableCheckBox(chkKey)
        cmdOk.Enabled = True
      End If

      'if there are any fields that have been mapped then set them
      If mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS) IsNot Nothing Then
        Dim vMappedCols As DataRow() = SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}'", mvImportForm.MappedAttribute("ColumnNumber"))
        For Each vMappedCol As DataRow In vMappedCols
          'The server always send the new attribute columns that match the mapped file headings. Only map an attribute column when its not already mapped,
          'otherwise remove it from this MappedAttributeColumns table.
          If SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, String.Format("ID = '{0}'", vMappedCol("ID")) & " AND MappedColumnNumber <> '{0}'", mvImportForm.MappedAttribute("ColumnNumber")).Length = 0 Then
            dgr.ActiveColumn = IntegerValue(vMappedCol("ColumnNumber"))
            SelectComboBoxItem(cboAttrs, vMappedCol("ID").ToString)
          Else
            mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vMappedCol)
          End If
        Next
      End If
    End If

    'Row headers to display the row number
    dgr.SetRowHeaderVisible()
    For vIndex As Integer = 1 To dgr.RowCount
      dgr.SetRowHeaderValue(vIndex - 1, 0, vIndex.ToString)
    Next
    dgr.SetPreferredRowHeaderWidth(0)

    'Set Column Widths
    For vindex As Integer = 0 To dgr.ColumnCount
      dgr.SetPreferredColumnWidth(0)
    Next
  End Sub

  Private Sub RemoveFromList(ByVal pItemCode As String)
    Dim vItem As LookupItem = mvAttrItems.FirstOrDefault(Function(vFindItem) vFindItem.LookupCode = pItemCode)
    If vItem IsNot Nothing Then
      Dim vIdx As Integer = mvAttrItems.IndexOf(vItem)
      If vIdx >= 0 Then
        AttrsRemove(vIdx, False, False)
      End If
    End If
  End Sub

  Private Sub SetImportFileName()
    If mvImportForm Is Nothing Then
      If mvImportFilename.Substring(mvImportFilename.Length - 4, 4).ToLower = ".def" Then
        'If we have imported a def file then set up form so that import type can only be changed for multiple import runs
        SetUpForMultipleImport()
        'If mvMainImportFileName.Length > 0 Then mvDataImport.FileName = mvMainImportFileName
      Else
        GetImportDataSets(mvImportFilename)
      End If
    Else
      mvImportForm.MappedAttribute("FileName") = mvImportFilename
      GetImportDataSets(mvImportFilename, True)
    End If

    InitForm()
  End Sub

  Private Function InitForMultipleImport() As Boolean
    Try
      mvMainDefFileName = mvImportFilename
      GetImportDataSets(mvImportFilename)
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enExternalFilenameInvalid
          'Select a file if the import file was left blank in the def file. 
          SelectImportFile()
          If mvMainImportFileName.Length = 0 Then
            Return False
          End If
          InitForMultipleImport()
        Case CareException.ErrorNumbers.enWriteAccessDenied, CareException.ErrorNumbers.enImportFileManatoryForMultiple
          ShowInformationMessage(vEx.Message, mvImportFilename)
          Return False
        Case Else
          Throw vEx
      End Select
    End Try
    Return True
  End Function

  Private Sub GetImportDataSets(ByVal pFileName As String, Optional ByVal pIncludeImportType As Boolean = False, Optional ByVal pSplitDefFile As Boolean = True)
    Dim vParams As ParameterList = Nothing

    Dim vTables As String = String.Format("{0},{1}", DATA_IMPORT, DATA_IMPORT_PARAMS)
    If mvImportForm IsNot Nothing Then
      'Initialise the data import class with the help of the master data import
      vParams = GetParameterList(vTables, MASTER)
      If GetValue(DATA_IMPORT_PARAMS, "DefFileName", MASTER).Length > 0 Then
        vParams("MasterImportFileName") = GetValue(DATA_IMPORT_PARAMS, "DefFileName", MASTER)
        vParams("MainImportFileName") = GetValue(DATA_IMPORT_PARAMS, "FileName", MASTER)  'Send the main data file name so that the correct items are read
      Else
        vParams("MasterImportFileName") = GetValue(DATA_IMPORT_PARAMS, "FileName", MASTER)
      End If
    Else
      vParams = GetParameterList(vTables, MAIN)
    End If

    vParams("IgnoreUnknownParameters") = CBoolYN(True)
    vParams("FileName") = pFileName
    vParams("NoRead") = CBoolYN(mvNoRead)
    'Prevents the server from splitting up the def file every time a call is made
    vParams("SplitDefFile") = CBoolYN(pSplitDefFile)
    If mvMainImportFileName.Length > 0 Then vParams("MainImportFileName") = mvMainImportFileName
    If pIncludeImportType Then
      If mvImportForm Is Nothing Then
        vParams("ImportType") = cboType.Text
      Else
        vParams("ImportType") = mvImportForm.cboType.Text
      End If
    End If
    Dim vDataSet As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaInit)
    CopyTablesFromDataSet(vDataSet)
    RefreshDefaults()
  End Sub

  ''' <summary>
  ''' Keeps the defaults dataset in sync with the data import dataset and rebinds the grid
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub RefreshDefaults()
    Dim vTable As DataTable = DataImportDS.Tables(DEFAULTS_ROW)
    If vTable Is Nothing Then
      'Create a table to hold the default values
      vTable = CreateTable(DEFAULTS_ROW)
      'Add the table to the dataset
      DataImportDS.Tables.Add(vTable)
    End If

    'Create a new dataset from the two tables containing the defaults data
    'which can be directly bound to the grid
    If DefaultsDS Is Nothing Then
      DefaultsDS = New DataSet
    Else
      If DefaultsDS.Tables.Contains("DataRow") Then DefaultsDS.Tables.Remove("DataRow")
    End If

    Dim vCopy As DataTable = vTable.Copy 'Create a copy as a table cannot be present in two datasets
    vCopy.TableName = "DataRow"
    DefaultsDS.Tables.Add(vCopy)

    If Not DefaultsDS.Tables.Contains("Column") Then
      vTable = DataImportDS.Tables(DEFAULTS_COL)
      If vTable IsNot Nothing Then
        DataImportDS.Tables.Remove(vTable)
        vTable.TableName = "Column"
        'vTable.Columns("Size").ColumnName = "Width" 'grid expects a column name of Width
        DefaultsDS.Tables.Add(vTable)
      End If
    End If

    'Bind the grid
    dgrDefaults.Populate(DefaultsDS)
  End Sub

  Public Sub ResetAll()
    dtpckValue.Value = Today
    txtLookupDefValue.Text = String.Empty
    cboPatternValue.Text = String.Empty
    dtpckValue.Visible = False
    txtLookupDefValue.Visible = False
    cboPatternValue.Visible = False
    lblElse.Visible = False
    DisableCheckBox(chkIncPerLine)
    DisableCheckBox(chkCtrlNo)
    chkIncPerLine.Visible = False
    chkCtrlNo.Visible = False
    txtDefValue.Visible = False
    txtDefValue.Text = String.Empty
  End Sub

  Public Sub HideTableControls()
    cboTables.Visible = False
    cboTables.Enabled = False
    lblTableDesc.Visible = False  'lblTable(2)
    cboGroups.Visible = False
    lblGroups.Visible = False
    DisableCheckBox(chkEmptyBeforeImport)
    'options enabled
    chkControlNumbers.Enabled = True
    chkValCodes.Enabled = True
  End Sub

  Public Sub EnableUpdateOptions()
    chkUpdate.Enabled = True
    chkUpdateAll.Enabled = True
    lblUpdateSub.Enabled = True
  End Sub

  Private Sub EnableFinancialHistoryOptions(ByVal pEnabled As Boolean)
    optPaymentsFinHistory.Enabled = pEnabled
    optPaymentsPostedToCB.Enabled = pEnabled
    optPaymentsPostedToNominal.Enabled = pEnabled
  End Sub

  Private Sub ShowEmployeeUpdate(ByVal pEnabled As Boolean)
    chkEmployee.Visible = pEnabled
    txtOrgNumber.Visible = pEnabled
    lblOrganisation.Visible = pEnabled
  End Sub

  Private Sub EnableAddressUpdateOptions(ByVal pEnabled As Boolean)
    chkExtractAddr.Enabled = pEnabled
    chkCacheMailsortAddr.Enabled = pEnabled
  End Sub

  Private Sub DisableCheckBox(ByVal pControl As CheckBox)
    pControl.Checked = False
    If Not mvInitializingFromDef Then
      pControl.Enabled = False
    End If
  End Sub

  Private Sub SetControls(pFromLoad As Boolean)
    Select Case SelectedImportType
      Case 12   'ditTableImport
        ClearMappingTables(pFromLoad)
        GetMaintenanceGroups()
        SetControlNumAccess()
        'to reset the checkbox when switching between import types(as there would be no def file at this point)
        If Not GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad") Then
          chkControlNumbers.Checked = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.import_check_control_default, True)
        End If
        If mvImportForm Is Nothing Then ShowTableControls()
        DisableSourceDet()
      Case Else
        HideTableControls()
        SetControlNumAccess()
        'to reset the checkbox when switching between import types(as there would be no def file at this point)
        If Not GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad") Then
          chkControlNumbers.Checked = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.import_check_control_default, True)
        End If
        EnableSourceDet()
    End Select
  End Sub

  Public Sub ShowTableControls()
    cboTables.Visible = True
    cboTables.Enabled = True
    lblTableDesc.Visible = True
    cboGroups.Visible = True
    cboGroups.Enabled = True
    lblGroups.Visible = True
    chkEmptyBeforeImport.Enabled = True
    'options disabled
    DisableCheckBox(chkControlNumbers)
    DisableCheckBox(chkValCodes)
  End Sub

  Public Sub DisableSourceDet()
    txtSource.Text = String.Empty
    txtDataSource.Text = String.Empty
    txtSource.Enabled = False
    txtDataSource.Enabled = False
    lblSource.Enabled = False
    lblDataSource.Enabled = False
  End Sub

  Public Sub EnableSourceDet()
    txtSource.Enabled = True
    txtDataSource.Enabled = True
    lblSource.Enabled = True
    lblDataSource.Enabled = True
  End Sub

  Private Sub SetControlNumAccess()
    'TODO" And gvEnv.OptionEnabled(vItemID) Then
    chkControlNumbers.Enabled = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciImportCheckControlsCheckBox)
  End Sub

  Private Sub SetOptionTabDefaults()
    chkControlNumbers.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ControlNumbers")

    If GetBooleanValue(DATA_IMPORT, "EnableRemoveIndexes") Then
      chkNoIndexes.Enabled = True
      chkNoIndexes.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "RemoveIndexes")
    Else
      DisableCheckBox(chkNoIndexes)
    End If
    If GetBooleanValue(DATA_IMPORT, "EnableAmendmentHistory") Then
      chkAmendmentHistory.Enabled = True
      chkAmendmentHistory.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "AmendmentHistory")
    Else
      DisableCheckBox(chkAmendmentHistory)
    End If

    chkValCodes.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ValidateCodes")
    If GetBooleanValue(IMPORT_OPTIONS, "EnableCreateCMD") Then chkCreateCMD.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CreateCMD") 'BR19057
    chkCreateCMD_Click(Nothing, Nothing)
    If chkCMDSupp.Enabled Then chkCMDSupp.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "CreateCMDForWarningSupp") 'BR19057
    chkReplaceQuestionMark.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "ReplaceQuestionMark")
    chkReplaceQuestionMark_Click(Nothing, Nothing)
    txtReplaceQuestionMarkWith.Text = GetValue(DATA_IMPORT_PARAMS, "ReplaceQuestionMarkWith")
    If GetIntegerValue(DATA_IMPORT_PARAMS, "ControlNumberBlockSize") > 0 Then
      txtControlNumberBlockSize.Text = GetIntegerValue(DATA_IMPORT_PARAMS, "ControlNumberBlockSize").ToString
    Else
      txtControlNumberBlockSize.Text = ControlText.TxtHundred
    End If
    chkLogCreate.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogCreate")
    chkLogWarn.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogWarnings")
    chkLogDups.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogDups")
    chkLogConversion.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogConversions")
    chkLogDedupAudit.Checked = GetBooleanValue(DATA_IMPORT_PARAMS, "LogDedupAudit")
    Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "MIRecords")
      Case 0
        optMIRecordsSuccFromPrevImport.Checked = True
      Case 1
        optMIRecordsSuccFromFirstImport.Checked = True
      Case 2
        optMIRecordsOriginalFile.Checked = True
    End Select
  End Sub

  Private Sub ResetFinancialHistoryOptions()
    chkGiftAidRecords.Enabled = True
    chkProcessIncentives.Enabled = True
    chkCreateAct.Enabled = True
    chkMatchSchPayment.Enabled = True
    txtNumberOfDays.Enabled = True
    lblNumberOfDays.Enabled = True

    chkGiftAidRecords.Checked = False
    chkProcessIncentives.Checked = False
    chkCreateAct.Checked = False
    chkMatchSchPayment.Checked = False
    txtNumberOfDays.Text = String.Empty
    If IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) = 8 OrElse IntegerValue(GetValue(DATA_IMPORT_PARAMS, "DataImportType")) = 16 Then 'If .Parameters.DataImportType = ditFinancialHistory Or .Parameters.DataImportType = ditEventBookingAndDelegates Then
      'Do not set optPayments*.Checked values here as this will overwrite a copy of the def file values.
      ' - optPayments*.Checked values are set using a copy of .def file values in InitFormParameters() 
    Else
      optPaymentsFinHistory.Checked = False
      optPaymentsPostedToCB.Checked = False
      optPaymentsPostedToNominal.Checked = False
      optPaymentsUnposted.Checked = False
    End If
  End Sub

  Private Sub optPayments_Click(ByVal sender As Object, ByVal e As EventArgs) Handles optPaymentsFinHistory.CheckedChanged, optPaymentsPostedToCB.CheckedChanged, optPaymentsPostedToNominal.CheckedChanged, optPaymentsUnposted.CheckedChanged
    Dim vIndex As Integer
    Select Case CType(sender, Control).Name
      Case optPaymentsFinHistory.Name
        vIndex = 0
      Case optPaymentsPostedToCB.Name
        vIndex = 1
      Case optPaymentsPostedToNominal.Name
        vIndex = 2
      Case optPaymentsUnposted.Name
        vIndex = 3
    End Select

    SetValue(DATA_IMPORT, "PaymentImportType", vIndex)

    If vIndex = 2 OrElse vIndex = 0 Then
      chkGiftAidRecords.Enabled = True
      chkCreateAct.Enabled = True
      DisableCheckBox(chkProcessIncentives)
    Else
      DisableCheckBox(chkGiftAidRecords)
      DisableCheckBox(chkCreateAct)
      chkProcessIncentives.Enabled = True
      If Not mvInitializingFromDef Then chkProcessIncentives.Checked = True
    End If

    If vIndex = 1 OrElse vIndex = 3 Then
      'Posted to Cash Book and Un-posted Transactions
      chkMatchSchPayment.Enabled = True
    Else
      DisableCheckBox(chkMatchSchPayment)
      txtNumberOfDays.Text = String.Empty
      txtNumberOfDays.Enabled = False
      lblNumberOfDays.Enabled = False
    End If

    If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 16 Then ' ditEventBookingAndDelegates 
      chkGiftAidRecords.Checked = False
      chkProcessIncentives.Checked = False
      chkMatchSchPayment.Checked = False
      DisableCheckBox(chkGiftAidRecords)
      DisableCheckBox(chkProcessIncentives)
      DisableCheckBox(chkMatchSchPayment)
      lblNumberOfDays.Enabled = False
    End If
    'vsePayments.Refresh()
  End Sub

  Private Sub optDedup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles optDedupAddressOnly.CheckedChanged, optDedupFull.CheckedChanged, optDedupNone.CheckedChanged
    Dim vIndex As Integer
    Select Case CType(sender, Control).Name
      Case "optDedupFull"
        vIndex = 0
        If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 12 Then 'ditTableImport
          EnableUpdateOptions()
        Else
          EnableDeDupOptions(True)
        End If
      Case "optDedupAddressOnly"
        vIndex = 1
        EnableDeDupOptions(True)
      Case "optDedupNone"
        vIndex = 2
        EnableDeDupOptions(False)
        If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 12 Then DisableUpdateOptions()
    End Select
    SetValue(DATA_IMPORT_PARAMS, "Dedup", vIndex)
  End Sub

  Private Sub optMapValue_Click(ByVal sender As Object, ByVal e As EventArgs) Handles optMapValueLookup.CheckedChanged, optMapValueNull.CheckedChanged
    Dim vIndex As Integer
    Select Case CType(sender, Control).Name
      Case "optMapValueNull"
        vIndex = 0
      Case "optMapValueLookup"
        vIndex = 1
    End Select
    If mvImportForm IsNot Nothing Then mvImportForm.MappedAttribute("MapReturnValueType") = vIndex
  End Sub

  Private Sub optMIRecords_Click(ByVal sender As Object, ByVal e As EventArgs) Handles optMIRecordsOriginalFile.CheckedChanged, optMIRecordsSuccFromFirstImport.CheckedChanged, optMIRecordsSuccFromPrevImport.CheckedChanged
    Dim vIndex As Integer
    Select Case CType(sender, Control).Name
      Case "optMIRecordsSuccFromPrevImport"
        vIndex = 0
      Case "optMIRecordsSuccFromFirstImport"
        vIndex = 1
      Case "optMIRecordsOriginalFile"
        vIndex = 2
    End Select
    SetValue(DATA_IMPORT_PARAMS, "MIRecords", vIndex)
  End Sub

  Private Sub optStock_Click(ByVal sender As Object, ByVal e As EventArgs) Handles optStockSet.CheckedChanged, optStockUpdate.CheckedChanged
    Dim vIndex As Integer
    Select Case CType(sender, Control).Name
      Case "optStockUpdate"
        vIndex = 0
      Case "optStockSet"
        vIndex = 1
    End Select
    SetValue(DATA_IMPORT_PARAMS, "StockUpdateType", vIndex)
  End Sub

  Public Sub DisableUpdateOptions()
    DisableCheckBox(chkUpdate)
    DisableCheckBox(chkUpdateAll)
    lblUpdateSub.Enabled = False
  End Sub

  Private Sub ImportFileRead(ByVal pSplitDefFile As Boolean, Optional ByVal pSetImportType As Boolean = True)
    If Not mvNoRead OrElse pSetImportType = False Then
      ClearMappingTables(False)
      If mvImportForm Is Nothing Then
        GetImportDataSets(mvImportFilename, pSetImportType, pSplitDefFile)

        If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 12 Then 'ditTableImport
          Exit Sub
        End If
      Else
        'Only read the import file as the seperator has changed
        Dim vParams As New ParameterList(True)
        vParams("FileName") = mvImportFilename
        vParams("Separator") = cboSeparator.Text
        vParams("IgnoreUnknownParameters") = CBoolYN(True)
        Dim vDataSet As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaImportFileRead)
        CopyTablesFromDataSet(vDataSet)
      End If

      'Add items to cboAttrs only when changing the Import Type and not initialising the form
      If mvNoRead = False AndAlso mvImportForm Is Nothing Then
        AddAttributesFromDataSet()
      End If

      InitFormGrid()
      If mvMultipleDataImportRuns Then InitFormDefault()
    End If
  End Sub

  ''' <summary>
  ''' Clear attribtues,defaults,map attributes
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ClearMappingTables(pFromLoad As Boolean)
    If Not pFromLoad Then
      If DataImportDS.Tables(ATTRIBUTE_COL) IsNot Nothing Then DataImportDS.Tables(ATTRIBUTE_COL).Rows.Clear()
      If DataImportDS.Tables(DEFAULTS_ROW) IsNot Nothing Then DataImportDS.Tables(DEFAULTS_ROW).Rows.Clear()
    End If
    If DataImportDS.Tables(MAPPED_ATTRIBUTES) IsNot Nothing Then DataImportDS.Tables(MAPPED_ATTRIBUTES).Rows.Clear()
    If DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS) IsNot Nothing Then DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Clear()
    If DataImportDS.Tables(MAP_ATTR_COLS) IsNot Nothing Then DataImportDS.Tables(MAP_ATTR_COLS).Rows.Clear()
  End Sub

  ''' <summary>
  ''' Creates a parameter list of the changed values for the mentioned tables
  ''' </summary>
  ''' <param name="pTables">Comma seperated list of tables</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetParameterList(ByVal pTables As String, Optional ByVal pDataSet As String = MAIN) As ParameterList
    Dim vDataSet As DataSet = DataImportDS
    If pDataSet = MASTER Then vDataSet = mvMasterDataImport
    Dim vParams As New ParameterList(True)
    vParams("IgnoreUnknownParameters") = CBoolYN(True)
    vParams("FileName") = mvImportFilename
    'If cboType.SelectedIndex > -1 Then vParams("ImportType") = cboType.Text
    If vDataSet IsNot Nothing Then
      For Each vTable As String In pTables.Split(","c)
        If mvImportForm IsNot Nothing AndAlso pDataSet = MAIN Then vTable = String.Format("MAIN_{0}", vTable)
        Dim vDataTable As DataTable = vDataSet.Tables(vTable)
        If vDataTable IsNot Nothing Then
          Dim vAddToList As Boolean = True
          Dim vStream As New IO.MemoryStream
          Dim vWriter As New Xml.XmlTextWriter(vStream, Nothing)
          If vTable = DATA_IMPORT OrElse vTable = DATA_IMPORT_PARAMS Then
            'For Data Import and Import Params only return the changed values
            If mvChangedValues.ContainsKey(vTable) Then
              For Each vCol As String In mvChangedValues(vTable)
                vWriter.WriteElementString(vCol, vDataTable.Rows(0)(vCol).ToString)
              Next
            Else
              vAddToList = False
            End If
          Else
            vWriter.WriteStartElement("ResultSet")
            For Each vRow As DataRow In vDataTable.Rows
              vWriter.WriteStartElement("DataRow")
              If vTable = DEFAULTS_ROW Then Debug.Print("")
              For Each vCol As DataColumn In vDataTable.Columns
                vWriter.WriteElementString(vCol.ColumnName, vRow.Item(vCol.ColumnName).ToString)
              Next vCol
              vWriter.WriteEndElement()
            Next vRow
            vWriter.WriteEndElement()
          End If
          If vAddToList Then
            vWriter.Flush()
            Dim vReader As New System.IO.StreamReader(vStream)
            vStream.Position = 0
            vParams(vTable) = vReader.ReadToEnd()
          Else
            vWriter.Close()
            vStream.Dispose()
          End If
        End If
      Next
    End If
    Return vParams
  End Function

  Private Sub txtNumberOfDays_Change(ByVal sender As Object, ByVal e As EventArgs) Handles txtNumberOfDays.TextChanged
    SetValue(DATA_IMPORT_PARAMS, "ScheduledPmtNoOfDays", txtNumberOfDays.Text)
  End Sub

  Private Sub txtControlNumberBlockSize_Change(ByVal sender As Object, ByVal e As EventArgs) Handles txtControlNumberBlockSize.TextChanged
    SetValue(DATA_IMPORT_PARAMS, "ControlNumberBlockSize", txtControlNumberBlockSize.Text)
  End Sub

  Private Sub txtControlNumberBlockSize_LostFocus(ByVal sender As Object, ByVal e As EventArgs) Handles txtControlNumberBlockSize.LostFocus
    If IntegerValue(txtControlNumberBlockSize.Text) < 1 OrElse IntegerValue(txtControlNumberBlockSize.Text) > 100 Then
      ShowWarningMessage(InformationMessages.ImControlNumberBlockSize)
      txtControlNumberBlockSize.Focus()
    End If
  End Sub

  Private Sub cboTables_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vUpdate As Boolean
    Dim vDataSet As DataSet = Nothing

    If cboTables.SelectedIndex > -1 Then
      SetValue(DATA_IMPORT_PARAMS, "TableImportTable", cboTables.SelectedValue)
      vUpdate = GetUpdateRights(cboTables.SelectedValue.ToString)

      If vUpdate Then
        EnableDeDup(True, True)
        optDedupAddressOnly.Checked = False
        optDedupAddressOnly.Enabled = False
      Else
        EnableDeDup(False, False)
        DisableCheckBox(chkUpdateTableImport)
        DisableCheckBox(chkUpdateAllTableImport)
        lblUpdateSub.Enabled = False
      End If

      If Not GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad") Then dgrDefaults.Clear()

      If mvImportForm Is Nothing Then
        AttrsClear()
      End If

      If (GetValue(DATA_IMPORT_PARAMS, "DefFileName").Length = 0 AndAlso mvImportForm Is Nothing) OrElse (Not GetBooleanValue(DATA_IMPORT_PARAMS, "FirstLoad")) Then
        Dim vTables As String = String.Format("{0},{1}", DATA_IMPORT, DATA_IMPORT_PARAMS)
        Dim vParams As ParameterList = GetParameterList(vTables)
        vParams("ImportType") = "Table Import"
        If mvMainImportFileName IsNot Nothing Then vParams.AddItemIfValueSet("MainImportFileName", mvMainImportFileName)
        vDataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaReadImportFileAndAttributes)
        CopyTablesFromDataSet(vDataSet)
        AddAttributesFromDataSet()
        InitFormGrid()
      End If

      If GetBooleanValue(DATA_IMPORT_PARAMS, "IgnoreFirstRow") Then chkIgnore.Checked = True Else chkIgnore.Checked = False
    Else
      SetValue(DATA_IMPORT_PARAMS, "TableImportTable", String.Empty)
      SetValue(DATA_IMPORT_PARAMS, "TableImportGroup", String.Empty)
    End If
  End Sub

  ''' <summary>
  ''' Adds the tables from the dataset into mvDataImportDataSet that is copy of the
  ''' Data Import class on the server. Will replace existing tables and add the new ones.
  ''' </summary>
  ''' <param name="pDataSet"></param>
  ''' <remarks></remarks>
  Private Sub CopyTablesFromDataSet(ByVal pDataSet As DataSet)
    If pDataSet IsNot Nothing Then
      If DataImportDS Is Nothing Then DataImportDS = New DataSet

      'Check if there are any default column mappings. 
      'eg.when we map a new csv file that has column headers
      If mvImportForm IsNot Nothing Then
        If Not BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
          'special case for mapped attribtue. Copy the values to the map attribute in the parent form
          If pDataSet.Tables.Contains(MAPPED_ATTRIBUTES) Then
            For Each vCol As DataColumn In pDataSet.Tables(MAPPED_ATTRIBUTES).Columns
              mvImportForm.MappedAttribute(vCol.ColumnName) = pDataSet.Tables(MAPPED_ATTRIBUTES).Rows(0)(vCol)
            Next
            pDataSet.Tables.Remove(MAPPED_ATTRIBUTES)
          End If

          'special case for mapped attribtue columns. Add rows to MAPPED_ATTRIBUTE_COLUMNS
          If pDataSet.Tables.Contains(MAPPED_ATTRIBUTE_COLUMNS) Then
            For Each vRow As DataRow In pDataSet.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows
              mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).ImportRow(vRow)
            Next
            pDataSet.Tables.Remove(MAPPED_ATTRIBUTE_COLUMNS)
          End If
        Else
          'Ignore the tables as these would contain the defaults
          If pDataSet.Tables.Contains(MAPPED_ATTRIBUTES) Then pDataSet.Tables.Remove(MAPPED_ATTRIBUTES)
          If pDataSet.Tables.Contains(MAPPED_ATTRIBUTE_COLUMNS) Then pDataSet.Tables.Remove(MAPPED_ATTRIBUTE_COLUMNS)
        End If
      End If

      'Add remaining tables
      While pDataSet.Tables.Count > 0
        CopyTable(pDataSet, DataImportDS, pDataSet.Tables(0).TableName)
      End While
    End If
  End Sub

  ''' <summary>
  ''' Copies a table from one dataset to another.
  ''' </summary>
  ''' <param name="pSource"></param>
  ''' <param name="pDest"></param>
  ''' <param name="pTable"></param>
  ''' <param name="pNewName"></param>
  ''' <remarks></remarks>
  Private Sub CopyTable(ByVal pSource As DataSet, ByVal pDest As DataSet, ByVal pTable As String, Optional ByVal pNewName As String = "")
    If pSource IsNot Nothing AndAlso pDest IsNot Nothing Then
      Dim vTable As DataTable = pSource.Tables(pTable)
      If pDest.Tables.Contains(IIf(pNewName.Length = 0, pTable, pNewName).ToString) Then
        pDest.Tables.Remove(vTable.TableName) 'Remove the old table
      End If
      pSource.Tables.Remove(pTable)  'Remove from temp dataset before adding to avoid an error
      If pNewName.Length > 0 Then vTable.TableName = pNewName
      pDest.Tables.Add(vTable)
    End If
  End Sub

  Private Function GetUpdateRights(ByVal pTable As String) As Boolean
    Dim vUpdate As Boolean

    Dim vTable As String = pTable
    'We don't want to expose lookup_group_details as a maintainable table in it's own right
    'but only through using the Details button from lookup_group
    'so it needs to inherit it's access rights from lookup_groups
    If vTable = "lookup_group_details" Then vTable = "lookup_groups"
    'similarly access to packed_products should be restricted based on access to products
    If vTable = "packed_products" Then vTable = "products"

    Dim vDataRow As DataRow() = mvMaintenanceTables.Select(String.Format("TableName = '{0}'", vTable))
    If vDataRow.Length > 0 Then
      vUpdate = BooleanValue(vDataRow(0)("PrivUpdate").ToString)    'Amend
    End If

    Return vUpdate
  End Function

  ''' <summary>
  ''' Adds a mapping between the attributes and the columns in the file
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub AddAttributeMapping()
    'Create a table to hold the mapped columns if it doesn't exist
    If Not DataImportDS.Tables.Contains(ATTRIBUTE_COL) Then DataImportDS.Tables.Add(CreateTable(ATTRIBUTE_COL))

    Dim vLookupItem As LookupItem = DirectCast(cboAttrs.SelectedItem, LookupItem)
    Dim vRow As DataRow = SelectRow(MAIN, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vLookupItem.LookupCode)
    'while loading from a def file the table would already contain rows for the mapped fields
    Dim vAddNew As Boolean = False
    If vRow Is Nothing Then
      vRow = DataImportDS.Tables(ATTRIBUTE_COL).NewRow() 'Add a row for the newly mapped column
      vAddNew = True
    End If
    SetColumnHeader(vLookupItem.LookupCode, vLookupItem.LookupDesc)

    vRow("AttributeIndex") = vLookupItem.LookupCode
    vRow("AttributeDesc") = vLookupItem.LookupDesc
    vRow("ColumnIndex") = dgr.ActiveColumn
    If vAddNew Then DataImportDS.Tables(ATTRIBUTE_COL).Rows.Add(vRow)
  End Sub

  ''' <summary>
  ''' Sets the column header to the attribute desc. Also adds the date format if required.
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SetColumnHeader(ByVal pAttributeIndex As String, ByVal pAttributeDesc As String)
    Dim vRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", pAttributeIndex)
    If PanelItem.GetFieldType(vRow("Type").ToString) = PanelItem.FieldTypes.cftDate OrElse
       (vRow("AttributeName").ToString = "expiry_date" AndAlso
        vRow("TableName").ToString <> "orders") OrElse
        vRow("AttributeName").ToString = "valid_date" Then
      If mvImportForm Is Nothing Then
        vRow = SelectRow(MAIN, ATTR_DATE_FORMAT, "ID = '{0}'", pAttributeIndex)
      Else
        vRow = SelectRow(MASTER, ATTR_DATE_FORMAT, "ID = '{0}'", pAttributeIndex)
      End If
      'all date attributes should have any entry in the date format table 
      Dim vDateFormat As String = AppValues.DateFormat
      If vRow IsNot Nothing Then vDateFormat = vRow("DateFormat").ToString
      pAttributeDesc = String.Format("{0}{1}({2})", pAttributeDesc, Environment.NewLine, vDateFormat)
    End If
    dgr.ColumnHeading(dgr.ActiveColumn) = pAttributeDesc
    dgr.AdjustColumnHeaders()
  End Sub

  Private Sub RemoveAttributeMapping(pDeferRefresh As Boolean)
    If DataImportDS.Tables.Contains(ATTRIBUTE_COL) Then
      'Remove the mapping for the current column
      Dim vRow As DataRow = SelectRow(MAIN, ATTRIBUTE_COL, "ColumnIndex = '{0}'", dgr.ActiveColumn)
      If vRow IsNot Nothing Then

        'Add the item to the list of available attributes
        AttrsAdd(vRow("AttributeDesc").ToString, vRow("AttributeIndex").ToString, pDeferRefresh)
        DataImportDS.Tables(ATTRIBUTE_COL).Rows.Remove(vRow)
        'Clear text from grid header
        dgr.ColumnHeading(dgr.ActiveColumn) = BLANK_HEADING
      End If
    End If
  End Sub

  Private Function SelectRow(ByVal pDataSet As String, ByVal pTable As String, ByVal pCriteria As String, ByVal ParamArray pValues As Object()) As DataRow
    Dim vRow As DataRow = Nothing
    Dim vRows As DataRow() = SelectRows(pDataSet, pTable, pCriteria, pValues)
    If vRows IsNot Nothing AndAlso vRows.Length > 0 Then vRow = vRows(0)
    Return vRow
  End Function

  Private Function SelectRows(ByVal pDataSet As String, ByVal pTable As String, ByVal pCriteria As String, ByVal ParamArray pValues As Object()) As DataRow()
    Dim vDataSet As DataSet = Nothing
    Dim vRows As DataRow() = Nothing
    If pDataSet = MASTER Then
      vDataSet = mvMasterDataImport
    Else
      vDataSet = DataImportDS
    End If
    If vDataSet.Tables.Contains(pTable) Then vRows = vDataSet.Tables(pTable).Select(String.Format(pCriteria, pValues))
    Return vRows
  End Function

  Private Function CreateTable(ByVal pTable As String) As DataTable
    Dim vTable As DataTable = Nothing
    Select Case pTable
      Case ATTRIBUTE_COL
        vTable = New DataTable(ATTRIBUTE_COL)
        vTable.Columns.Add("AttributeIndex")
        vTable.Columns.Add("AttributeDesc")
        vTable.Columns.Add("ColumnIndex")
      Case DEFAULTS_ROW
        vTable = New DataTable(DEFAULTS_ROW)
        vTable.Columns.Add("ID")
        vTable.Columns.Add("Attribute")
        vTable.Columns.Add("Value")
      Case MAPPED_ATTRIBUTES
        vTable = New DataTable(MAPPED_ATTRIBUTES)
        vTable.Columns.Add("ColumnNumber") 'col num in the main file
        vTable.Columns.Add("FileName")
        vTable.Columns.Add("MapNoOfCols")
        vTable.Columns.Add("MapIgnoreFirstRow")
        vTable.Columns.Add("MapSeparator")
        vTable.Columns.Add("MapReturnValueType")
        vTable.Columns.Add("MappedColOfFile") 'col in the 2nd file that is mapped to the main one
        vTable.Columns.Add("MapExistsAlready")
      Case MAPPED_ATTRIBUTE_COLUMNS
        vTable = New DataTable(MAPPED_ATTRIBUTE_COLUMNS)
        vTable.Columns.Add("MappedColumnNumber") 'Used to link with MAPPED_ATTRIBUTES table
        vTable.Columns.Add("ID")
        vTable.Columns.Add("ColumnNumber")
      Case MAP_ATTR_COLS
        vTable = New DataTable(MAP_ATTR_COLS)
        vTable.Columns.Add("ColumnNumber")
        vTable.Columns.Add("Heading")
      Case ATTR_DATE_FORMAT
        vTable = New DataTable(ATTR_DATE_FORMAT)
        vTable.Columns.Add("ID")
        vTable.Columns.Add("DateFormat")
    End Select
    Return vTable
  End Function

  Private Sub cboAttrs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAttrs.SelectedIndexChanged

    If Not RefreshingAttributes Then
      Dim vHeading As String = String.Empty
      Dim vRow As DataRow = Nothing

      If cboAttrs.SelectedIndex >= 0 Then
        If dgr.ActiveColumn > -1 Then
          vHeading = dgr.ColumnHeading(dgr.ActiveColumn).Trim 'Current column heading
          If mvImportForm Is Nothing Then
            If vHeading.Length > 0 Then 'Col is already mapped 
              If vHeading.Length <= 4 OrElse (vHeading.Length > 4 AndAlso vHeading.Substring(0, 4) <> "Map:") Then
                'Remove column allocation to an attribute !!
                RemoveAttributeMapping(False)
                'Add new column allocation
                AddAttributeMapping()
              End If
            Else 'No mapping for column...create a new one
              AddAttributeMapping()
            End If

            'Remove the item from the list of available attributes
            AttrsRemove(cboAttrs.SelectedIndex, False, False)
          Else
            If vHeading.Length > 0 Then
              'Remove column allocation to an attribute !! KEEP IN SYNC WITH grd_DoubleClick
              If vHeading.Length > 4 AndAlso vHeading.Substring(0, 4) = "Map:" Then
                dgr.ColumnHeading(dgr.ActiveColumn) = BLANK_HEADING
                chkKey.Checked = False
                chkKey.Enabled = True
                cmdOk.Enabled = False
                mvImportForm.MappedAttribute("MappedColOfFile") = 0
              Else
                vRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}' AND ColumnNumber = '{1}'", mvImportForm.MappedAttribute("ColumnNumber"), dgr.ActiveColumn)
                If vRow Is Nothing AndAlso mvImportForm.MappedAttribute("ColumnNumber").ToString.Length = 0 Then
                  'MappedColumnNumber field in the DataTable is null so try again
                  vRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber IS NULL AND ColumnNumber = '{0}'", dgr.ActiveColumn)
                End If
                If vRow IsNot Nothing Then
                  Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vRow("ID"))
                  mvImportForm.AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
                  AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)

                  mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vRow)
                End If
              End If
            End If

            'Create an entry for the newly mapped attribute column
            Dim vLookupItem As LookupItem = DirectCast(cboAttrs.SelectedItem, LookupItem)
            Dim vID As Integer = IntegerValue(vLookupItem.LookupCode)

            SetColumnHeader(vLookupItem.LookupCode, vLookupItem.LookupDesc)

            'If we are modifying an existing mapped atribute then add the changes to a temp
            'table which would allow for them to be undone if the user cancels
            If BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
              'Create a temp table to hold the mapped attribute columns 
              If mvTempMaintAttrCols Is Nothing Then mvTempMaintAttrCols = CreateTable(MAPPED_ATTRIBUTE_COLUMNS)
              Dim vRows() As DataRow = mvTempMaintAttrCols.Select(String.Format("ID = '{0}'", vID))
              If vRows.Length > 0 Then
                vRow = vRows(0)
              Else
                vRow = mvTempMaintAttrCols.NewRow
                mvTempMaintAttrCols.Rows.Add(vRow)
              End If
              vRow("MappedColumnNumber") = mvImportForm.MappedAttribute("ColumnNumber")
              vRow("ID") = vID
              vRow("ColumnNumber") = dgr.ActiveColumn
            Else
              vRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "ID = '{0}'", vID)
              If vRow Is Nothing Then
                vRow = mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).NewRow
                mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Add(vRow)
              End If
              vRow("MappedColumnNumber") = mvImportForm.MappedAttribute("ColumnNumber")
              vRow("ID") = vID
              vRow("ColumnNumber") = dgr.ActiveColumn
            End If

            vRow = SelectRow(MASTER, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vID)
            'We only need to remove the attribute from the mvImportForm.cboAttrs when the selected item is not in AttributeColumns table
            'and we are adding a new map. The existing mapped attribute items will be in AttribtuesColumns with ColumnIndex set to -1 and
            'the attribute will be removed from mvImportForm.cboAttrs on saving the mapped attribute using mvTempMaintAttrCols
            If vRow Is Nothing AndAlso BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) = False Then mvImportForm.AttrsRemove(cboAttrs.SelectedIndex, False, False)
            AttrsRemove(cboAttrs.SelectedIndex, False, False)
          End If
        End If
        cboAttrs.SelectedIndex = -1
      End If
    End If

  End Sub

  Private Sub txtReplaceQuestionMarkWith_Change(ByVal sender As Object, ByVal e As EventArgs) Handles txtReplaceQuestionMarkWith.TextChanged
    SetValue(DATA_IMPORT_PARAMS, "ReplaceQuestionMarkWith", txtReplaceQuestionMarkWith.Text)
  End Sub

  Private Sub cboSeparator_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If mvImportForm Is Nothing Then
        SetValue(DATA_IMPORT_PARAMS, "Separator", cboSeparator.Text)
      Else
        mvImportForm.MappedAttribute("MapSeparator") = cboSeparator.Text
        'Double click on all the column headers to remove any mapping
        For vIndex As Integer = 0 To dgr.ColumnCount - 1
          dgr_ColumnHeaderDoubleClicked(Nothing, vIndex, 0, 0)
        Next
      End If

      ImportFileRead(False)
      If Not mvImportForm Is Nothing Then mvImportForm.MappedAttribute("MapNoOfCols") = dgr.ColumnCount

      If GetValue(DATA_IMPORT_PARAMS, "Separator") = "Fixed" Then
        chkIgnore.Checked = False
        DisableCheckBox(chkIgnore)
      Else
        chkIgnore.Enabled = True
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub txtSource_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSource.Validated
    Dim vValue As String = String.Empty
    If txtSource.IsValid Then
      vValue = txtSource.Text
      SetValue(DATA_IMPORT_PARAMS, "Source", vValue)
    Else
      'MsgBox(XLAT("You must enter a valid value"))
    End If
  End Sub

  Private Sub txtDataSource_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataSource.Validated
    Dim vValue As String = String.Empty
    If txtDataSource.IsValid Then
      vValue = txtDataSource.Text
      SetValue(DATA_IMPORT_PARAMS, "DataSource", vValue)
    Else
      'MsgBox(XLAT("You must enter a valid value"))
    End If
  End Sub

  Private Sub cboDefAttrs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDefAttrs.SelectedIndexChanged
    If Not RefreshingAttributes Then
      Try
        Dim vPattern As String
        Dim vPatternArray() As String
        Dim vIndex As Integer
        Dim vNoDefault As Boolean
        Dim vContinue As Boolean = True

        If cboDefAttrs.SelectedIndex > -1 Then
          ResetAll()
          cboPatternValue.Text = String.Empty
          Dim vDataRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", DirectCast(cboDefAttrs.SelectedItem, LookupItem).LookupCode)
          Dim vDefaultAttributeName As String = Me.cboDefAttrs.SelectedItem.ToString()
          'BR20559 - Allow Rate to be defaulted if Product is defaulted or Product is mapped in the import file
          Select Case vDefaultAttributeName
            Case "Product"
              'Defaulting product, if rate already defaulted, remove the rate default and inform the user
              Dim vAttrRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "AttributeName = '{0}'", "rate")
              Dim vDefaultAttrRow As DataRow = SelectRow(MAIN, DEFAULTS_ROW, "ID = '{0}'", vAttrRow("ID").ToString)
              If vDefaultAttrRow IsNot Nothing Then
                RemoveDefaultAttribute(vAttrRow("ID").ToString, vAttrRow("AttributeNameDesc").ToString)
                ShowInformationMessage(InformationMessages.ImImportDefaultProductRemoveRate)
              End If
              If IntegerValue(vDataRow("RestrictionIndex")) > -1 Then
                If FindRestrictionRow() = -1 Then
                  Dim vRestAttr As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vDataRow("RestrictionIndex"))
                  ShowInformationMessage(InformationMessages.ImDependantDefaultValue, vDataRow("AttributeNameDesc").ToString, vRestAttr("AttributeNameDesc").ToString)
                  vContinue = False
                End If
              End If
            Case "Rate"
              'Defaulting Rate
              Dim vIsCMT As Boolean = (Me.SelectedImportType.Equals(18))  '18 = ditCMT
              Dim vAttrRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "AttributeName = '{0}'", If(vIsCMT = True, "membership_type", "product"))
              Dim vDefaultAttrRow As DataRow = SelectRow(MAIN, DEFAULTS_ROW, "ID = '{0}'", vAttrRow("ID").ToString)
              If vDefaultAttrRow Is Nothing Then
                'The Product / MembershipType is not defaulted
                Dim vMappedAttrRow As DataRow = SelectRow(MAIN, ATTRIBUTE_COL, "AttributeDesc = '{0}'", If(vIsCMT = True, "Membership Type", "Product"))
                If vMappedAttrRow Is Nothing Then
                  'The Product / MembershipType is not mapped to a column in the import file and is not defaulted
                  ShowInformationMessage(If(vIsCMT = True, InformationMessages.ImImportRateMembershipType, InformationMessages.ImImportRateProduct))
                  vContinue = False
                Else
                  'The Product / MembershipType is mapped to a column in the import file inform the user that validation cannot happen until the import is run
                  ShowInformationMessage(If(vIsCMT = True, InformationMessages.ImDependantDefaultValueImportRateMTFile, InformationMessages.ImDependantDefaultValueImportRateException))
                End If
              Else
                If vIsCMT Then
                  'Membership Type is defaulted cannot restrict the Rate, inform the user that validation cannot happen until the import is run
                  ShowInformationMessage(InformationMessages.ImDependantDefaultValueImportRateMTDefault)
                Else
                  'Product is defaulted so restrict the Rate
                  If IntegerValue(vDataRow("RestrictionIndex")) > -1 Then
                    If FindRestrictionRow() = -1 Then
                      Dim vRestAttr As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vDataRow("RestrictionIndex"))
                      ShowInformationMessage(InformationMessages.ImDependantDefaultValue, vDataRow("AttributeNameDesc").ToString, vRestAttr("AttributeNameDesc").ToString)
                      vContinue = False
                    End If
                  End If
                End If
              End If
            Case Else
              'If the value needs to be restricted then check if the user has specified a 
              'default for the restriction attribute
              If IntegerValue(vDataRow("RestrictionIndex")) > -1 Then
                If FindRestrictionRow() = -1 Then
                  Dim vRestAttr As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vDataRow("RestrictionIndex"))
                  ShowInformationMessage(InformationMessages.ImDependantDefaultValue, vDataRow("AttributeNameDesc").ToString, vRestAttr("AttributeNameDesc").ToString)
                  vContinue = False
                End If
              End If
          End Select

          If vContinue Then
            Select Case PanelItem.GetFieldType(vDataRow("Type").ToString)
              Case PanelItem.FieldTypes.cftCharacter, PanelItem.FieldTypes.cftMemo
                txtLookupDefValue.Visible = True
                txtLookupDefValue.MaxLength = IntegerValue(vDataRow("EntryLength"))
                vNoDefault = False
                cmdDefaultAdd.Enabled = True
                chkIncPerLine.Visible = False
                chkCtrlNo.Visible = False
              Case PanelItem.FieldTypes.cftNumeric, PanelItem.FieldTypes.cftLong, PanelItem.FieldTypes.cftInteger
                txtLookupDefValue.Visible = True
                txtLookupDefValue.MaxLength = IntegerValue(vDataRow("EntryLength"))
                vNoDefault = False
                cmdDefaultAdd.Enabled = True
                chkIncPerLine.Visible = True
                chkIncPerLine.Enabled = True
                chkCtrlNo.Visible = True
                chkCtrlNo.Enabled = True
              Case PanelItem.FieldTypes.cftDate, PanelItem.FieldTypes.cftTime
                dtpckValue.Visible = True
                dtpckValue.Value = Today
                cmdDefaultAdd.Enabled = True
                vNoDefault = False
                chkIncPerLine.Visible = False
                chkCtrlNo.Visible = False
              Case Else
                lblElse.Visible = True
                cmdDefaultAdd.Enabled = False
                vNoDefault = True
                chkIncPerLine.Visible = False
                chkCtrlNo.Visible = False
            End Select

            If Not vNoDefault Then
              vPattern = vDataRow("Pattern").ToString
              If vPattern.Length > 0 Then
                cboPatternValue.Visible = True
                txtLookupDefValue.Visible = False
                dtpckValue.Visible = False
                vPattern = Mid(vPattern, 2, vPattern.Length - 2)
                cboPatternValue.Items.Clear()
                If vPattern.Contains("|") Then
                  vPatternArray = vPattern.Split("|"c)
                  For vIndex = 0 To vPatternArray.Length - 1
                    cboPatternValue.Items.Add(vPatternArray(vIndex))
                  Next
                Else
                  For vIndex = 0 To vPattern.Length - 1
                    cboPatternValue.Items.Add(vPattern.Substring(vIndex, 1))
                  Next
                End If
              End If

              'Set up a simple textbox or a textlookup box based on the attribute info
              If txtLookupDefValue.Visible Then

                If vDataRow("ValidationTable").ToString.Length > 0 Then
                  InitTextLookupBox(txtLookupDefValue, vDataRow("TableName").ToString, vDataRow("AttributeName").ToString)

                Else
                  If vDataRow("AttributeName").ToString = "honorifics" AndAlso mvValidateHonorifics Then
                    InitTextLookupBox(txtLookupDefValue, vDataRow("TableName").ToString, vDataRow("AttributeName").ToString)
                  Else
                    'simple text box
                    txtLookupDefValue.Visible = False
                    txtDefValue.Visible = True
                    txtDefValue.MaxLength = txtLookupDefValue.MaxLength
                  End If
                End If

                If chkCtrlNo.Visible Then
                  Dim vPanelItem As New PanelItem(txtLookupDefValue, "control_number_type")
                  vPanelItem.SetValidationData("control_numbers", "control_number_type")
                  txtLookupDefValue.ComboBox.DataSource = Nothing 'Clear prev datasource(if any) 
                  txtLookupDefValue.Init(vPanelItem, False, False)
                End If
              End If
            End If
          End If
        End If
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    End If
  End Sub

  Private Sub txtLookupDefValue_GetInitialCodeRestrictions(ByVal sender As System.Object, ByVal pParameterName As System.String, ByRef pList As CDBNETCL.ParameterList) Handles txtLookupDefValue.GetInitialCodeRestrictions
    Try
      Dim vRow As Integer = FindRestrictionRow()
      pList = New ParameterList(True)
      If vRow >= 0 Then
        GetRestrictions(pList)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GetRestrictions(ByVal pList As ParameterList)
    'Find the restriction attribute in the grid for default values and use that
    'value to restrict the current selection
    Dim vRow As Integer = FindRestrictionRow()
    If vRow >= 0 Then
      Dim vAttrRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", DirectCast(cboDefAttrs.SelectedItem, LookupItem).LookupCode)
      Dim vRestAttr As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vAttrRow("RestrictionIndex"))
      Dim vAttrName As String = IIf(vRestAttr("ValidationAttribute").ToString.Length > 0, vRestAttr("ValidationAttribute"), vRestAttr("AttributeName")).ToString
      pList(ProperName(vAttrName)) = dgrDefaults.GetValue(vRow, "Value")
    End If
  End Sub

  Private Sub txtDefValue_GetCodeRestrictions(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pList As CDBNETCL.ParameterList) Handles txtLookupDefValue.GetCodeRestrictions
    Try
      GetRestrictions(pList)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' Checks if the defaults grid contains a value for the restriction attribute
  ''' </summary>
  ''' <returns>The row number at which the restriction value is found. -1 = Not found</returns>
  ''' <remarks></remarks>
  Private Function FindRestrictionRow() As Integer
    Dim vRow As Integer = -1
    Dim vLookupItem As LookupItem = CType(cboDefAttrs.SelectedItem, LookupItem)
    Dim vDefaultAttr As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vLookupItem.LookupCode)
    Dim vRestrAttr As String = vDefaultAttr("RestrictionAttribute").ToString
    If vRestrAttr.Length > 0 Then
      vRow = dgrDefaults.FindRow("ID", vDefaultAttr("RestrictionIndex").ToString)
    End If
    Return vRow
  End Function

  Private Sub txtNumberOfDays_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNumberOfDays.Leave
    Try
      If IntegerValue(txtNumberOfDays.Text) < 0 OrElse IntegerValue(txtNumberOfDays.Text) > 14 Then
        ShowInformationMessage(InformationMessages.ImTableImportNoOfDaysRange) 'Must be between 0 and 14
        txtNumberOfDays.Focus()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub txtOrgNumber_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrgNumber.Validated
    Try
      If txtOrgNumber.Text.Length > 0 Then
        If Not txtOrgNumber.IsValid Then
          ShowInformationMessage(InformationMessages.ImOrganisationNotFound)
          SetValue(DATA_IMPORT_PARAMS, "OrgNumber", String.Empty)
          SetValue(DATA_IMPORT_PARAMS, "OrgName", String.Empty)
          tabSub.SelectedTab = tbpCustomOpt
          txtOrgNumber.Focus()
        Else
          SetValue(DATA_IMPORT_PARAMS, "OrgNumber", txtOrgNumber.Text)
          SetValue(DATA_IMPORT_PARAMS, "OrgName", txtOrgNumber.Description)
          chkEmployee.Checked = True
        End If
      Else
        If chkEmployee.Checked = True Then
          chkEmployee.Checked = False
          SetValue(DATA_IMPORT_PARAMS, "EmployeeLoad", "N")
          SetValue(DATA_IMPORT_PARAMS, "OrgNumber", String.Empty)
          SetValue(DATA_IMPORT_PARAMS, "OrgName", String.Empty)
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdDefaultAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDefaultAdd.Click
    Try
      Dim vValue As String = String.Empty
      Dim vValue1 As String = String.Empty
      Dim vMsg As String
      Dim vLookup As Boolean

      If cboDefAttrs.Text.Length > 0 Then
        vMsg = ValidateDefaultValue()
        If vMsg.Length = 0 Then
          vLookup = txtLookupDefValue.Visible
          Dim vDataRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", DirectCast(cboDefAttrs.SelectedItem, LookupItem).LookupCode)
          Select Case PanelItem.GetFieldType(vDataRow("Type").ToString)
            Case PanelItem.FieldTypes.cftCharacter, PanelItem.FieldTypes.cftMemo, PanelItem.FieldTypes.cftBulk, PanelItem.FieldTypes.cftFile
              vValue = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString
              vValue1 = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString
            Case PanelItem.FieldTypes.cftLong, PanelItem.FieldTypes.cftNumeric, PanelItem.FieldTypes.cftInteger
              If chkIncPerLine.Checked Then
                vValue = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString
                vValue1 = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString & ", Increment Per Line"
              ElseIf chkCtrlNo.Checked Then
                vValue = txtLookupDefValue.Text
                vValue1 = "Control Number: " & txtLookupDefValue.Text
              Else
                vValue = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString
                vValue1 = IIf(vLookup, txtLookupDefValue.Text, txtDefValue.Text).ToString
              End If
            Case PanelItem.FieldTypes.cftDate
              vValue = dtpckValue.Value.ToString(AppValues.DateFormat)
              vValue1 = dtpckValue.Value.ToString(AppValues.DateFormat)
          End Select

          If vDataRow("Pattern").ToString <> "" Then
            vValue = cboPatternValue.Text
            vValue1 = cboPatternValue.Text
          End If

          If cboDefAttrs.SelectedIndex > -1 Then
            If vValue.Length > 0 Then
              AddDefaultAttribute(vValue1)
              ResetAll()
            Else
              ShowInformationMessage(InformationMessages.ImDefaultValueRequired)
              txtDefValue.Focus()
            End If
          Else
            ShowInformationMessage(InformationMessages.ImSelectAttribute)
            cboDefAttrs.Focus()
          End If
        Else
          If vMsg <> "none" Then ShowInformationMessage(vMsg)
        End If
        cboDefAttrs.SelectedIndex = -1
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Function ValidateDefaultValue() As String
    Dim vMsg As String = String.Empty

    If txtLookupDefValue.Text.Length > 0 Then
      Dim vDataRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", DirectCast(cboDefAttrs.SelectedItem, LookupItem).LookupCode)
      If (vDataRow("Type").ToString = "I" _
          OrElse vDataRow("Type").ToString = "N" _
          OrElse vDataRow("Type").ToString = "L") _
          And chkCtrlNo.Checked Then
        If Not txtLookupDefValue.IsValid Then vMsg = "You must enter a valid value"
      Else
        If Not txtLookupDefValue.IsValid Then
          If ShowQuestion(QuestionMessages.QmUseInvalidValue, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
            vMsg = "none"
          End If
        End If
      End If
    End If
    Return vMsg
  End Function

  Private Sub dgrDefaults_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgrDefaults.RowDoubleClicked

    If pRow >= 0 Then
      'Remove item from grid
      Dim vAttr As String = dgrDefaults.GetValue(pRow, "Attribute")
      If vAttr.Length > 0 Then
        Dim vIndex As String = dgrDefaults.GetValue(pRow, "ID")
        RemoveDefaultAttribute(vIndex, vAttr)

        'Check if this was a restriction attribute for another value
        'and remove that as well 
        Dim vAttrRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "RestrictionIndex = '{0}'", vIndex)
        If vAttrRow IsNot Nothing Then
          RemoveDefaultAttribute(vAttrRow("ID").ToString, vAttrRow("AttributeNameDesc").ToString)
        End If
      End If
    End If
    ResetAll()
  End Sub

  ''' <summary>
  ''' Adds a default value for a field and binds the grid
  ''' </summary>
  ''' <param name="pValue"></param>
  ''' <remarks></remarks>
  Public Sub AddDefaultAttribute(ByVal pValue As String)
    'Adding a row for the default value in the ImportDefaults table of the dataset
    Dim vLookupItem As LookupItem = DirectCast(cboDefAttrs.SelectedItem, LookupItem)
    Dim vRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vLookupItem.LookupCode)

    ''while loading from a def file the table would already contain rows for the mapped fields
    'vRow = SelectRow(MAIN, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vAttrIndex)
    'If vRow Is Nothing Then
    '  vRow = mvDataImport.Tables(ATTRIBUTE_COL).NewRow() 'Add a row for the newly mapped column
    '  vAddNew = True
    'End If

    Dim vTable As DataTable = DataImportDS.Tables(DEFAULTS_ROW)
    Dim vDefaultRow As DataRow = vTable.NewRow
    vDefaultRow("ID") = vLookupItem.LookupCode
    vDefaultRow("Attribute") = vLookupItem.LookupDesc
    vDefaultRow("Value") = pValue
    vTable.Rows.Add(vDefaultRow)

    RefreshDefaults()

    AttrsRemove(cboDefAttrs.SelectedIndex, True, False)
    txtDefValue.Text = String.Empty
    txtLookupDefValue.Text = String.Empty
    dtpckValue.Value = Today
  End Sub

  ''' <summary>
  ''' Removes a default value for a field and binds the grid
  ''' </summary>
  ''' <param name="pIndex"></param>
  ''' <remarks></remarks>
  Public Sub RemoveDefaultAttribute(ByVal pIndex As String, ByVal pAttr As String)
    Dim vDataRow As DataRow = SelectRow(MAIN, DEFAULTS_ROW, "ID = '{0}'", pIndex)
    If vDataRow IsNot Nothing Then
      DataImportDS.Tables(DEFAULTS_ROW).Rows.Remove(vDataRow)
      AttrsAdd(pAttr, pIndex, False)
      RefreshDefaults()
    End If
  End Sub

  Private Sub dgr_ColumnHeaderDoubleClicked(ByVal sender As System.Object, ByVal pColumn As System.Int32, ByVal pX As System.Int32, ByVal pY As System.Int32) Handles dgr.ColumnHeaderDoubleClicked
    Dim vHeading As String = String.Empty
    Dim vRow As DataRow = Nothing
    If pColumn >= 0 Then
      vHeading = dgr.ColumnHeading(pColumn).Trim
      If vHeading.Length > 0 Then
        If mvImportForm Is Nothing Then
          If vHeading.Length > 4 AndAlso vHeading.Substring(0, 4) = "Map:" Then
            DeleteMap()
          Else
            'Remove column allocation to an attribute  
            RemoveAttributeMapping(False)
          End If
        Else
          If vHeading.Length > 4 AndAlso vHeading.Substring(0, 4) = "Map:" Then
            'dgr.ColumnHeading(pColumn) = BLANK_HEADING
            chkKey.Checked = False
            chkKey.Enabled = True
            cmdOk.Enabled = False
          End If
          Dim vCriteria As String = Nothing
          If mvImportForm.MappedAttribute("ColumnNumber").ToString.Length > 0 Then
            vCriteria = "MappedColumnNumber = '{0}' AND ColumnNumber = '{1}'"
          Else
            'User has clicked cancel on a new map file
            vCriteria = "(MappedColumnNumber = '' OR MappedColumnNumber IS NULL) AND ColumnNumber = '{1}'"
          End If

          'If we are modifying an existing mapped atribute then add the changes to a temp
          'table which would allow for them to be undone if the user cancels
          If BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
            Dim vRows() As DataRow = mvTempMaintAttrCols.Select(String.Format(vCriteria, mvImportForm.MappedAttribute("ColumnNumber"), pColumn))
            If vRows.Length > 0 Then
              Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vRows(0)("ID"))
              'Only remove the attribute from cboAttrs (mapped import form) as it will be removed from mvImportForm.cboAttrs on saving the mapped attribute.
              AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
              mvTempMaintAttrCols.Rows.Remove(vRows(0))
            End If
          Else
            vRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, vCriteria, mvImportForm.MappedAttribute("ColumnNumber"), pColumn)
            If vRow IsNot Nothing Then
              Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vRow("ID"))
              mvImportForm.AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
              AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
              mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vRow)
            End If
          End If

          'Clear text from grid header
          dgr.ColumnHeading(dgr.ActiveColumn) = BLANK_HEADING

        End If
      End If
    End If
  End Sub

  ''' <summary>
  ''' Deletes the mapped attribute for the active column in the grid
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DeleteMap()
    Dim vRow As DataRow = SelectRow(MAIN, MAPPED_ATTRIBUTES, "ColumnNumber = '{0}'", dgr.ActiveColumn)
    Dim vAttr As DataRow = Nothing
    If vRow IsNot Nothing Then
      DataImportDS.Tables(MAPPED_ATTRIBUTES).Rows.Remove(vRow)
      Dim vMappedCols As DataRow() = SelectRows(MAIN, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}'", dgr.ActiveColumn)
      For Each vMappedCol As DataRow In vMappedCols
        vAttr = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vMappedCol("ID"))
        AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, True)
        'Remove the row from AttributeColumns table having column index -1 to keep the data in sync
        vRow = SelectRows(MAIN, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vMappedCol("ID"))(0)
        If vRow IsNot Nothing Then DataImportDS.Tables(ATTRIBUTE_COL).Rows.Remove(vRow)
        DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vMappedCol)
      Next
      AttrsRefresh()
      vAttr = SelectRow(MAIN, MAP_ATTR_COLS, "ColumnNumber = '{0}'", dgr.ActiveColumn)
      If vAttr IsNot Nothing Then DataImportDS.Tables(MAP_ATTR_COLS).Rows.Remove(vAttr)
      dgr.ColumnHeading(dgr.ActiveColumn) = BLANK_HEADING
    End If
  End Sub

  Private Sub dgrMenuStrip_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dgrMenuStrip.Opening
    'Display context menu only if the user clicked on the header
    If dgr.MenuClickOnHeader Then
      Dim vHeading As String = String.Empty

      If mvImportForm IsNot Nothing Then
        dgrMenuMapAttribute.Visible = False
      Else
        dgrMenuMapAttribute.Visible = True
      End If

      vHeading = dgr.ColumnHeading(dgr.ActiveColumn).Trim

      If vHeading.Length = 0 Then
        If mvImportForm Is Nothing Then
          dgrMenuMapAttribute.Enabled = True
          dgrMenuDateFormat.Enabled = False
        Else
          dgrMenuDateFormat.Enabled = False
        End If
      Else
        If vHeading.Length > 4 AndAlso vHeading.Substring(0, 4) = "Map:" AndAlso mvImportForm Is Nothing Then
          dgrMenuMapAttribute.Enabled = True
          dgrMenuDateFormat.Enabled = False

          If DataImportDS.Tables(MAP_ATTR_COLS) Is Nothing Then CreateTable(MAP_ATTR_COLS)
          Dim vColRow As DataRow = SelectRow(MAIN, MAP_ATTR_COLS, "ColumnNumber = '{0}'", dgr.ActiveColumn)
          If vColRow Is Nothing Then
            Dim vRow As DataRow = DataImportDS.Tables(MAP_ATTR_COLS).NewRow
            vRow("ColumnNumber") = dgr.ActiveColumn
            vRow("Heading") = vHeading
            DataImportDS.Tables(MAP_ATTR_COLS).Rows.Add(vRow)
          Else
            vColRow("Heading") = vHeading
          End If
        Else
          dgrMenuMapAttribute.Enabled = False
          If mvImportForm Is Nothing Then
            Dim vColRow As DataRow = SelectRow(MAIN, MAP_ATTR_COLS, "ColumnNumber = '{0}'", dgr.ActiveColumn)
            If vColRow IsNot Nothing Then DataImportDS.Tables(MAP_ATTR_COLS).Rows.Remove(vColRow)
            Dim vRow As DataRow = SelectRow(MAIN, ATTRIBUTE_COL, "ColumnIndex = '{0}'", dgr.ActiveColumn)
            Dim vAttrIndex As String = vRow("AttributeIndex").ToString
            'Find the attribute that is mapped to this column
            Dim vDataRow As DataRow = SelectRow(MAIN, DATA_IMPORT_ATTRS, "ID = '{0}'", vAttrIndex)
            If PanelItem.GetFieldType(vDataRow("Type").ToString) = PanelItem.FieldTypes.cftDate OrElse
            (vDataRow("AttributeName").ToString = "expiry_date" AndAlso
             vDataRow("TableName").ToString <> "orders") OrElse
             vDataRow("AttributeName").ToString = "valid_date" Then
              dgrMenuDateFormat.Enabled = True
            Else
              dgrMenuDateFormat.Enabled = False
            End If
          Else
            Dim vMapAttrCol As DataRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "ColumnNumber = '{0}'", dgr.ActiveColumn)
            If vMapAttrCol IsNot Nothing Then
              Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vMapAttrCol("ID"))
              If PanelItem.GetFieldType(vAttr("Type").ToString) = PanelItem.FieldTypes.cftDate OrElse
              (vAttr("AttributeName").ToString = "expiry_date" AndAlso
              vAttr("TableName").ToString <> "orders") OrElse
              vAttr("AttributeName").ToString = "valid_date" Then
                dgrMenuDateFormat.Enabled = True
              Else
                dgrMenuDateFormat.Enabled = False
              End If
            End If
          End If
        End If
      End If
    Else
      e.Cancel = True
    End If
  End Sub

  Private Sub dgrMenuMapAttribute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrMenuMapAttribute.Click
    Try
      MapAttribute()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub dgrMenuDateFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrMenuDateFormat.Click
    Dim vFrmDateFormat As frmDateFormat = Nothing
    Dim vRow As DataRow = Nothing
    Dim vAttrRow As DataRow = Nothing
    Dim vDateFormat As String = String.Empty
    Dim vDataSet As String = String.Empty
    Dim vIndex As String = String.Empty
    Try
      'Get the mapping for the current column
      If mvImportForm Is Nothing Then
        vDataSet = MAIN
        vRow = SelectRow(MAIN, ATTRIBUTE_COL, "ColumnIndex = '{0}'", dgr.ActiveColumn)
        If Not DataImportDS.Tables.Contains(ATTR_DATE_FORMAT) Then DataImportDS.Tables.Add(CreateTable(ATTR_DATE_FORMAT))
        vIndex = vRow("AttributeIndex").ToString
        vRow = SelectRow(vDataSet, ATTR_DATE_FORMAT, "ID = '{0}'", vIndex)
        vAttrRow = SelectRow(vDataSet, DATA_IMPORT_ATTRS, "ID = '{0}'", vIndex)
        If vRow Is Nothing Then
          vRow = DataImportDS.Tables(ATTR_DATE_FORMAT).NewRow
          vRow("ID") = vIndex
          vRow("DateFormat") = SelectRow(vDataSet, ATTR_DATE_FORMAT, "ID = '{0}'", vIndex)("DateFormat")
          DataImportDS.Tables(ATTR_DATE_FORMAT).Rows.Add(vRow)
        End If
      Else
        vDataSet = MASTER
        vRow = SelectRow(MAIN, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}' AND ColumnNumber = '{1}'", mvImportForm.dgr.ActiveColumn, dgr.ActiveColumn)
        If Not mvMasterDataImport.Tables.Contains(ATTR_DATE_FORMAT) Then mvMasterDataImport.Tables.Add(CreateTable(ATTR_DATE_FORMAT))
        vRow = SelectRow(vDataSet, ATTR_DATE_FORMAT, "ID = '{0}'", vRow("AttributeIndex"))
        vAttrRow = SelectRow(vDataSet, DATA_IMPORT_ATTRS, "ID = '{0}'", vIndex)
        If vRow Is Nothing Then
          vRow = mvMasterDataImport.Tables(ATTR_DATE_FORMAT).NewRow
          vRow("ID") = vIndex
          vRow("DateFormat") = SelectRow(vDataSet, ATTR_DATE_FORMAT, "ID = '{0}'", vIndex)("DateFormat")
          mvMasterDataImport.Tables(ATTR_DATE_FORMAT).Rows.Add(vRow)
        End If
      End If

      Dim vExpiryDate As Boolean = (vAttrRow("AttributeName").ToString = "expiry_date" AndAlso
                                          vAttrRow("TableName").ToString <> "orders") _
                                          OrElse vAttrRow("AttributeName").ToString = "valid_date"

      If vRow IsNot Nothing Then vDateFormat = vRow("DateFormat").ToString
      vFrmDateFormat = New frmDateFormat(vDateFormat, vExpiryDate)
      vFrmDateFormat.ShowDialog()
      If vFrmDateFormat.DialogResult = System.Windows.Forms.DialogResult.OK Then RefreshDateFormat(vFrmDateFormat.DateFormat, vRow)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' Changes the date format for the mapped column
  ''' </summary>
  ''' <param name="pDateFormat">The new date format</param>
  ''' <param name="pRow">The row containing the prev date format</param>
  ''' <remarks></remarks>
  Public Sub RefreshDateFormat(ByVal pDateFormat As String, ByVal pRow As DataRow)
    dgr.ColumnHeading(dgr.ActiveColumn) = dgr.ColumnHeading(dgr.ActiveColumn).Replace(pRow("DateFormat").ToString, pDateFormat)
    pRow("DateFormat") = pDateFormat
  End Sub

  Private Sub MapAttribute()
    Dim vImportForm As frmImport
    Dim vOFD As New OpenFileDialog
    Dim vImportFile As String = String.Empty
    Dim vSelectFile As Boolean = True

    If DataImportDS.Tables(MAP_ATTR_COLS) Is Nothing Then DataImportDS.Tables.Add(CreateTable(MAP_ATTR_COLS))

    'Check if the column has been mapped previously
    Dim vRow As DataRow = SelectRow(MAIN, MAP_ATTR_COLS, "ColumnNumber = '{0}'", dgr.ActiveColumn)
    If vRow Is Nothing Then
      'No mapping exists....allow the user to select a file to map with
      With vOFD
        Do While vSelectFile
          .Title = ControlText.OfdSelectImportFile
          .CheckFileExists = True
          .CheckPathExists = True
          .FileName = ""
          .Filter = "CSV Files (*.csv)|*.csv|Fixed Format Files (*.fff)|*.fff|All Files (*.*)|*.*"
          .FilterIndex = 1
          If Not String.IsNullOrWhiteSpace(mvDefaultMapFolder) Then
            .InitialDirectory = mvDefaultMapFolder
            mvDefaultMapFolder = String.Empty
          End If
          If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
            If vOFD.FileName = String.Empty Then
              ShowWarningMessage(InformationMessages.ImImportFileNotFound)
            Else
              mvDefaultMapFolder = Path.GetDirectoryName(Path.GetFullPath(vOFD.FileName))
              vSelectFile = False
            End If
          Else
            vSelectFile = False
            Exit Sub
          End If
        Loop
      End With
    End If

    If DataImportDS.Tables(MAPPED_ATTRIBUTES) Is Nothing Then DataImportDS.Tables.Add(CreateTable(MAPPED_ATTRIBUTES))

    For Each vMappedAttribute As DataRow In DataImportDS.Tables(MAPPED_ATTRIBUTES).Rows
      If IntegerValue(vMappedAttribute("ColumnNumber")) <> dgr.ActiveColumn AndAlso vMappedAttribute("FileName").ToString = vOFD.FileName Then
        ShowErrorMessage(InformationMessages.ImFileAlreadyMapped)
        Exit Sub
      End If
    Next

    If vRow Is Nothing Then
      vImportFile = vOFD.FileName
      CreateMapAttribute()
      mvMapAttribute("MapSeparator") = ","
    Else
      vImportFile = vRow("Heading").ToString.Substring(4, vRow("Heading").ToString.Length - 4).Trim
      AssignExistingMapAttribute()
    End If

    vImportForm = New frmImport(vImportFile, Me, DataImportDS, mvChangedValues)
    vImportForm.Show()
  End Sub

  Private Sub AssignExistingMapAttribute()
    Dim vTable As DataTable = DataImportDS.Tables(MAPPED_ATTRIBUTES)
    Dim vOriginalRow As DataRow = SelectRow(MAIN, MAPPED_ATTRIBUTES, "ColumnNumber = '{0}'", dgr.ActiveColumn)
    'create clone of the original row so that the values are preserved
    'even if the user clicks cancel
    mvMapAttribute = vTable.NewRow
    mvMapAttribute.ItemArray = vOriginalRow.ItemArray
    mvMapAttribute("MapExistsAlready") = CBoolYN(True)
  End Sub

  Private Sub CreateMapAttribute()
    If DataImportDS.Tables(MAPPED_ATTRIBUTES) Is Nothing Then DataImportDS.Tables.Add(CreateTable(MAPPED_ATTRIBUTES))
    mvMapAttribute = DataImportDS.Tables(MAPPED_ATTRIBUTES).NewRow
    If DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS) Is Nothing Then DataImportDS.Tables.Add(CreateTable(MAPPED_ATTRIBUTE_COLUMNS))
  End Sub

  Public Sub PrepareAttributeForm()
    Dim vAdd As Boolean
    Dim vAttrs() As DataRow = Nothing
    AttrsClear()
    'Add the columns that have not already been mapped or for which defaults have not been setup
    For Each vRow As DataRow In mvMasterDataImport.Tables(DATA_IMPORT_ATTRS).Rows
      If vRow("AttributeName").ToString.Length > 0 Then

        'mapped to attribtue
        vAttrs = SelectRows(MASTER, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vRow("ID"))
        vAdd = mvMasterDataImport.Tables(ATTRIBUTE_COL) Is Nothing OrElse vAttrs.Length = 0
        'default value
        vAttrs = SelectRows(MASTER, DEFAULTS_ROW, "ID = '{0}'", vRow("ID").ToString)
        If vAdd Then vAdd = mvMasterDataImport.Tables(DEFAULTS_ROW) Is Nothing OrElse vAttrs.Length = 0
        'mapped to a attribute in a diff file
        vAttrs = SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "ID = '{0}'", vRow("ID"))
        If Not vAdd Then vAdd = mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS) IsNot Nothing AndAlso (vAttrs.Length > 0 AndAlso vAttrs(0)("MappedColumnNumber").ToString = mvImportForm.MappedAttribute("ColumnNumber").ToString)
        If vAdd Then AttrsAdd(vRow("AttributeNameDesc").ToString, vRow("ID").ToString, True)
      End If
    Next
    AttrsRefresh()

    'Adjust the positions of the visible controls to fill in the empty spaces
    chkIgnore.Top = mvImportForm.chkIgnore.Top + (500 \ AppValues.TwipsConversionY)
    lblColumn.Top = mvImportForm.lblColumn.Top - (600 \ AppValues.TwipsConversionY)
    cboAttrs.Top = mvImportForm.cboAttrs.Top - (600 \ AppValues.TwipsConversionY)
    lblAttribute.Top = mvImportForm.lblAttribute.Top - (600 \ AppValues.TwipsConversionY)
    cboSeparator.Top = mvImportForm.cboSeparator.Top - (600 \ AppValues.TwipsConversionY)
    lblSeperator.Top = mvImportForm.lblSeperator.Top - (600 \ AppValues.TwipsConversionY)
    lblKey.Top = mvImportForm.lblKey.Top - (600 \ AppValues.TwipsConversionY)
    chkKey.Top = mvImportForm.chkKey.Top - (600 \ AppValues.TwipsConversionY)

    cboType.Visible = False
    cboTables.Visible = False
    grpDedup.Visible = False
    lblDataImportType.Visible = False
    lblTableDesc.Visible = False
    cboGroups.Visible = False
    lblGroups.Visible = False

    lblSource.Visible = False
    lblDataSource.Visible = False
    txtSource.Visible = False
    txtDataSource.Visible = False

    tbpData.Text = String.Empty
    tabMain.TabPages.Remove(tbpDefaults)
    tabMain.TabPages.Remove(tbpOptions)
  End Sub

  Private Sub chkKey_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkKey.CheckedChanged
    If chkKey.Checked = True Then
      If dgr.ActiveColumn >= 0 Then
        If dgr.ColumnHeading(dgr.ActiveColumn).Trim.Length > 0 Then
          'Remove column allocation to an attribute !! 
          Dim vRow As DataRow = SelectRow(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}' AND ColumnNumber = '{1}'", mvImportForm.MappedAttribute("ColumnNumber"), dgr.ActiveColumn)
          If vRow IsNot Nothing Then
            Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vRow("ID"))
            mvImportForm.AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
            AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, False)
            mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vRow)
          End If
        End If

        dgr.ColumnHeading(dgr.ActiveColumn) = "Map:" & mvImportForm.MappedAttribute("FileName").ToString
        mvImportForm.MappedAttribute("MappedColOfFile") = dgr.ActiveColumn
      End If
      DisableCheckBox(chkKey)
      cmdOk.Enabled = True
    End If
  End Sub

  Private Sub SaveMapAttribute()
    If mvMasterDataImport.Tables(MAPPED_ATTRIBUTES) Is Nothing Then mvMasterDataImport.Tables.Add(CreateTable(MAPPED_ATTRIBUTES))
    Dim vTable As DataTable = mvMasterDataImport.Tables(MAPPED_ATTRIBUTES)
    Dim vRow As DataRow = SelectRow(MASTER, MAPPED_ATTRIBUTES, "ColumnNumber = '{0}'", mvImportForm.MappedAttribute("ColumnNumber"))
    If vRow IsNot Nothing Then
      'We have previously mapped this file. Remove the old mapping.
      vTable.Rows.Remove(vRow)
      'Remove previously mapped columns of this file
      If mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS) IsNot Nothing Then
        Dim vMappedRows As DataRow() = SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '{0}'", mvImportForm.MappedAttribute("ColumnNumber"))
        For Each vMappedRow As DataRow In vMappedRows
          'Remove entry from attr cols table if the user has removed mapping to this field.
          'Fields that are mapped to different files will have ColumnIndex as -1 in the master import attr cols collection
          If mvTempMaintAttrCols Is Nothing OrElse mvTempMaintAttrCols.Select(String.Format("ID = '{0}'", vMappedRow("ID"))).Length = 0 Then
            mvMasterDataImport.Tables(ATTRIBUTE_COL).Rows.Remove(SelectRow(MASTER, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vMappedRow("ID")))
            'Add the attribute back in mvImportForm.cboAttrs on removing an existing mapped attribute
            If BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
              Dim vAttr As DataRow = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vMappedRow("ID"))
              mvImportForm.AttrsAdd(vAttr("AttributeNameDesc").ToString, vAttr("ID").ToString, True)
            End If
          End If
          mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vMappedRow)
        Next
        mvImportForm.AttrsRefresh()
      End If
    End If
    mvMasterDataImport.Tables(MAPPED_ATTRIBUTES).Rows.Add(mvImportForm.MappedAttribute)
    If BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
      'Add columns as well if the mapping already existed
      If mvTempMaintAttrCols IsNot Nothing Then
        For Each vMappedCol As DataRow In mvTempMaintAttrCols.Rows
          mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).ImportRow(vMappedCol)
          'Add a row in AttributeColumns for any new mapped attribute with ColumnIndex -1 to keep the data in sync
          'and remove the attribute from mvImportForm.cboAttrs
          If SelectRow(MASTER, ATTRIBUTE_COL, "AttributeIndex = '{0}'", vMappedCol("ID").ToString) Is Nothing Then
            vRow = mvMasterDataImport.Tables(ATTRIBUTE_COL).NewRow()
            vRow("AttributeIndex") = vMappedCol("ID")
            vRow("AttributeDesc") = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vMappedCol("ID").ToString)("AttributeNameDesc")
            vRow("ColumnIndex") = -1
            mvMasterDataImport.Tables(ATTRIBUTE_COL).Rows.Add(vRow)
            RemoveFromList(vMappedCol("ID").ToString)  'Remove the attribute from attributes list
          End If
        Next
      End If
    Else
      'update the mapped attribute columns with the mapped column number
      Dim vRows As DataRow() = SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '' OR MappedColumnNumber IS NULL")
      If vRows.Length > 0 Then
        For Each vRow In vRows
          vRow("MappedColumnNumber") = mvImportForm.MappedAttribute("ColumnNumber")
          'Add a row in AttributeColumns for all new mapped attributes with ColumnIndex -1 to keep the data in sync.
          'The attributes have already been removed from mvImportForm.cboAttrs
          If Not mvMasterDataImport.Tables.Contains(ATTRIBUTE_COL) Then mvMasterDataImport.Tables.Add(CreateTable(ATTRIBUTE_COL))
          Dim vID As String = vRow("ID").ToString
          vRow = mvMasterDataImport.Tables(ATTRIBUTE_COL).NewRow() 'Add a row for the newly mapped column
          vRow("AttributeIndex") = vID
          vRow("AttributeDesc") = SelectRow(MASTER, DATA_IMPORT_ATTRS, "ID = '{0}'", vID)("AttributeNameDesc")
          vRow("ColumnIndex") = -1
          mvMasterDataImport.Tables(ATTRIBUTE_COL).Rows.Add(vRow)
        Next
      End If
    End If
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Dim vIndex As Integer
    If mvImportForm IsNot Nothing Then
      If Not BooleanValue(mvImportForm.MappedAttribute("MapExistsAlready").ToString) Then
        For vIndex = 0 To dgr.ColumnCount - 1
          dgr_ColumnHeaderDoubleClicked(Nothing, vIndex, 0, 0)
        Next
      End If

      'If mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS) IsNot Nothing Then
      '  Dim vRows() As DataRow = SelectRows(MASTER, MAPPED_ATTRIBUTE_COLUMNS, "MappedColumnNumber = '' OR MappedColumnNumber IS NULL")
      '  For Each vRow As DataRow In vRows
      '    mvMasterDataImport.Tables(MAPPED_ATTRIBUTE_COLUMNS).Rows.Remove(vRow)
      '  Next
      'End If
    End If
    Me.Close()
  End Sub

  Private Sub tabMain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tabMain.SelectedIndexChanged
    Select Case DirectCast(sender, CDBNETCL.TabControl).SelectedTab.Name
      Case tbpData.Name
        cmdTest.Visible = True
        cmdOk.Visible = True
        cmdCancel.Visible = True
        cmdSave.Visible = False
      Case tbpOptions.Name
        cmdTest.Visible = False
        cmdOk.Visible = False
        cmdCancel.Visible = False
        cmdSave.Visible = True
      Case tbpDefaults.Name
        cmdTest.Visible = False
        cmdOk.Visible = False
        cmdCancel.Visible = False
        cmdSave.Visible = False
    End Select
    bpl.RepositionButtons()
  End Sub

  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    Dim vResult As DialogResult
    Dim vDefFileName As String = String.Empty
    Dim vAppend As Boolean
    Dim vSFD As New SaveFileDialog
    Try
      vResult = System.Windows.Forms.DialogResult.No
      If mvMultipleDataImportRuns And mvImportDefinitionsChanged Then
        'User has changed the sequences of the import definitions
        'Check this is correct before attempting to save
        If ValidateMultipleImportList() = False Then vResult = System.Windows.Forms.DialogResult.Cancel
      End If

      If vResult <> vbCancel Then
        If Not mvMultipleDataImportRuns Then vSFD.OverwritePrompt = False

        'limit selection to a UNC path
        Dim vSelectFile As Boolean = True
        With vSFD
          Do While vSelectFile
            .Title = ControlText.SfdImportDef
            .FileName = GetValue(DATA_IMPORT_PARAMS, "FileName").Substring(0, GetValue(DATA_IMPORT_PARAMS, "FileName").Length - 4) & ".def"
            .DefaultExt = ".def"
            .Filter = "DEF Files (*.def)|*.def"
            .CheckPathExists = True
            If Not String.IsNullOrWhiteSpace(mvDefaultDefFolder) Then
              .InitialDirectory = mvDefaultDefFolder
              mvDefaultDefFolder = String.Empty
            End If
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
              If .FileName.StartsWith("\\") Then
                mvDefaultDefFolder = Path.GetDirectoryName(Path.GetFullPath(vSFD.FileName))
                vSelectFile = False
              Else
                'only UNC files are supported allow the user to reselect a file
                ShowInformationMessage(InformationMessages.ImUNCPathOnly)
              End If
            Else
              vSelectFile = False
              Exit Sub
            End If
          Loop
        End With
        vDefFileName = vSFD.FileName

        If mvMultipleDataImportRuns = False Then
          If mvMainDefFileName <> vDefFileName Or (String.Equals(mvMainDefFileName, vDefFileName) And Not mvOverWriteDefFile) Then
            'Do not prompt to over-write if saving to the same file (and import type not changed)
            If My.Computer.FileSystem.FileExists(vDefFileName) Then
              If mvMainDefFileName Is Nothing Then
                vResult = ShowQuestion(String.Format(QuestionMessages.QmDefFileExistsAdd, vDefFileName, Environment.NewLine), MessageBoxButtons.YesNoCancel)
                'Yes:     Append definition
                'No:      Over-write definition
                'Cancel:  Do not Save
                'If we are going to append, then check to this import definition not already in file
                If vResult = System.Windows.Forms.DialogResult.Yes AndAlso (mvMainDefFileName <> vDefFileName) Then vResult = CheckSingleImportFile(vDefFileName)
                vAppend = (vResult = System.Windows.Forms.DialogResult.Yes)

                If String.Equals(mvMainDefFileName, vDefFileName) Then vAppend = False
                If String.Equals(mvMainDefFileName, vDefFileName) AndAlso (vResult = System.Windows.Forms.DialogResult.Yes Or vResult = System.Windows.Forms.DialogResult.No) Then
                  mvOverWriteDefFile = True
                Else
                  mvOverWriteDefFile = False
                End If
              Else
                'Since we have loaded / saved an import definition, only overwriting existing definition file is allowed 
                vResult = ShowQuestion(InformationMessages.ImDefFileExistsOverwrite, MessageBoxButtons.OKCancel, vDefFileName, Environment.NewLine)
              End If
            End If
          ElseIf mvImportTypeChanged Then
            'We have loaded / saved an import definition and now the import type has been changed
            vResult = ShowQuestion(InformationMessages.ImDefFileExistsOverwrite, MessageBoxButtons.OKCancel, vDefFileName, Environment.NewLine)
          End If
        Else
          'If the file has changed then it is over-written - user has already been warned
          vAppend = (mvMainDefFileName = vDefFileName)
        End If
      End If

      If vResult <> vbCancel Then
        Dim vTables As String = String.Format("{0},{1},{2},{3},{4},{5},{6}", DATA_IMPORT, DATA_IMPORT_PARAMS, ATTRIBUTE_COL, DEFAULTS_ROW, MAPPED_ATTRIBUTES, MAPPED_ATTRIBUTE_COLUMNS, ATTR_DATE_FORMAT)
        If (vAppend OrElse mvMultipleDataImportRuns) Then
          SaveMultipleImportFile(vDefFileName, True)
        Else
          'Just over-write the existing file
          SaveDefFile(vDefFileName)
        End If
        mvMainDefFileName = vDefFileName
        mvImportTypeChanged = False
        ShowInformationMessage(InformationMessages.ImFileCreatedSuccessfully)
      End If
    Catch vCareException As CareException
      If vCareException.ErrorNumber = CareException.ErrorNumbers.enImportFileManatoryForMultiple Then
        ShowErrorMessage(vCareException.Message)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Function SaveDefFile(Optional ByVal pDefFileName As String = "") As String
    If pDefFileName.Length = 0 Then
      'Most likely the user is scheduling an import directly from a csv file.
      Dim vImportFile As String = GetValue(DATA_IMPORT_PARAMS, "FileName")
      pDefFileName = vImportFile.Substring(0, vImportFile.LastIndexOf("."c)) & ".def"
    End If

    Dim vTables As String = String.Format("{0},{1},{2},{3},{4},{5},{6}", DATA_IMPORT, DATA_IMPORT_PARAMS, ATTRIBUTE_COL, DEFAULTS_ROW, MAPPED_ATTRIBUTES, MAPPED_ATTRIBUTE_COLUMNS, ATTR_DATE_FORMAT)
    Dim vParams As ParameterList = GetParameterList(vTables)
    If GetValue(DATA_IMPORT_PARAMS, "DefFileName").Length > 0 Then
      vParams("FileName") = GetValue(DATA_IMPORT_PARAMS, "DefFileName")
      'Pass the MainImportFileName so that the system reads the correct attributes for mapped attributes on saving the def file
      If DataImportDS.Tables(MAPPED_ATTRIBUTES) IsNot Nothing AndAlso DataImportDS.Tables(MAPPED_ATTRIBUTES).Rows.Count > 0 Then
        vParams("MainImportFileName") = mvMainImportFileName
      End If
    Else
      vParams("FileName") = GetValue(DATA_IMPORT_PARAMS, "FileName")
    End If
    vParams("SaveDataFileName") = GetValue(DATA_IMPORT_PARAMS, "FileName")
    vParams("DefFileName") = pDefFileName
    vParams("IncludeSectionHeaders") = CBoolYN(True)
    vParams("AppendFile") = CBoolYN(False)
    vParams("SaveFileName") = CBoolYN(chkNoFileName.Checked = False)
    vParams("ImportType") = cboType.Text
    DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaSaveDefinitionFile)
    Return pDefFileName
  End Function

  Private Function ValidateMultipleImportList() As Boolean
    Dim vConIndex As Integer
    Dim vIndex As Integer
    Dim vSave As Boolean = True
    Dim vType As Integer
    Dim vLookupItem As LookupItem = Nothing

    vConIndex = -1
    For vIndex = 0 To lstMultiImportSelected.Items.Count - 1
      vLookupItem = CType(lstMultiImportSelected.Items(vIndex), LookupItem)
      vType = IntegerValue(vLookupItem.LookupCode)
      If vType = 0 Then vConIndex = vIndex 'ditContactOrganisation 
    Next

    If vConIndex > -1 And vConIndex <> 0 Then
      'Error - Contact / Org must be first
      vSave = False
      ShowWarningMessage(InformationMessages.ImInvalidImportTypeSeq, lstMultiImportSelected.Items(vConIndex).ToString, Environment.NewLine)
    End If

    Return vSave
  End Function

  Private Sub SetUpForMultipleImport()
    mvMultipleDataImportRuns = GetBooleanValue(DATA_IMPORT, "MultipleDataImportRuns")

    If mvMultipleDataImportRuns Then
      mvImportDefinitionsChanged = True   'Force the current records to be validated
      PopulateMultipleImportList()
    End If
  End Sub

  Private Sub PopulateMultipleImportList()
    lstMultiImportSelected.Items.Clear()
    If cboType.DataSource Is Nothing Then
      cboType.Items.Clear()
    Else
      cboType.DataSource = Nothing
    End If
    If mvMultipleDataImportRuns Then
      Dim vTable As DataTable = DataImportDS.Tables(DEF_FILE_PARAMS)
      Dim vItem As LookupItem = Nothing
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          vItem = New LookupItem(vRow("ImportType").ToString, vRow("ImportTypeDesc").ToString)
          lstMultiImportSelected.Items.Add(vItem)
          cboType.Items.Add(vItem)
        Next
      End If
      lstMultiImportSelected.SelectedIndex = -1
      If cboType.Items.Count > 0 Then cboType.SelectedIndex = 0
      mvSelectedImportType = SelectedImportType
      mvSelectedImportTypeDesc = cboType.Text
    End If
  End Sub

  Private Function ExtractDefForMultiImport(ByVal pImportType As Integer) As String
    Dim vFileName As String = String.Empty
    'Return the name of the temp def file for the selected type
    Dim vRow As DataRow = SelectRow(MAIN, DEF_FILE_PARAMS, "ImportType = '{0}'", pImportType)
    If vRow IsNot Nothing Then vFileName = vRow("FileName").ToString
    Return vFileName
  End Function

  Private Sub SelectImportFile()
    Dim vFolder As String
    Dim vFileName As String
    Dim vOFD As New OpenFileDialog

    vFolder = My.Computer.FileSystem.GetParentPath(mvMainDefFileName)
    If Not vFolder.EndsWith("\") Then vFolder = vFolder & "\"

    Dim vRetry As Boolean = False 'Show the dialoug again when the user selects a def file
    Do
      With vOFD
        .Title = ControlText.OfdSelectImportFile
        .DefaultExt = ".csv"
        .Filter = "Import Files(*.csv;*.txt)|*.csv;*.txt|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1
        .CheckFileExists = True
        .CheckPathExists = True
        If Not String.IsNullOrWhiteSpace(mvDefaultImportFolder) Then
          .InitialDirectory = mvDefaultImportFolder
          mvDefaultImportFolder = String.Empty
        End If
        .ShowDialog()
      End With

      vFileName = vOFD.FileName
      If vFileName.Length > 0 AndAlso vFileName.Substring(vFileName.Length - 4, 4) = ".def" Then
        ShowWarningMessage(InformationMessages.ImDefFileSelectionInvalid)
        vFileName = String.Empty
        vRetry = True
      Else
        vRetry = False
      End If
    Loop While vRetry
    If vFileName.Length > 0 Then
      mvDefaultImportFolder = Path.GetDirectoryName(Path.GetFullPath(vOFD.FileName))
    End If
    mvMainImportFileName = vFileName
    chkNoFileName.Checked = True
  End Sub

  Private Sub cmdAddDefinition_Click(ByVal pSender As Object, ByVal e As EventArgs) Handles cmdAddDefinition.Click
    Try
      Dim vIndex As Integer = lstMultiImportAvailable.SelectedIndex
      If vIndex >= 0 Then
        lstMultiImportSelected.Items.Add(lstMultiImportAvailable.SelectedItem)
        lstMultiImportAvailable.Items.Remove(lstMultiImportAvailable.SelectedItem)
        If vIndex >= lstMultiImportAvailable.Items.Count Then vIndex = vIndex - 1
        lstMultiImportAvailable.SelectedIndex = vIndex
        If vIndex < 0 Then cmdAddDefinition.Enabled = False
        cmdRemoveDefinition.Enabled = (lstMultiImportSelected.Items.Count > 0)
        mvImportDefinitionsChanged = True
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdRemoveDefinition_Click(ByVal pSender As Object, ByVal e As EventArgs) Handles cmdRemoveDefinition.Click
    Try
      Dim vIndex As Integer = lstMultiImportSelected.SelectedIndex
      If vIndex >= 0 Then
        lstMultiImportAvailable.Items.Add(lstMultiImportSelected.SelectedItem)
        lstMultiImportSelected.Items.Remove(lstMultiImportSelected.SelectedItem)
        If vIndex >= lstMultiImportSelected.Items.Count Then vIndex = vIndex - 1
        lstMultiImportSelected.SelectedIndex = vIndex
        If vIndex < 0 Then cmdRemoveDefinition.Enabled = False
        cmdAddDefinition.Enabled = True
        mvImportDefinitionsChanged = True
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub lstMultiImportAvailable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMultiImportAvailable.Click
    If lstMultiImportAvailable.SelectedIndex > 0 Then cmdAddDefinition.Enabled = True
  End Sub

  Private Sub lstMultiImportAvailable_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMultiImportAvailable.DoubleClick
    cmdAddDefinition_Click(Me, New EventArgs)
  End Sub

  Private Sub lstMultiImportSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMultiImportSelected.Click
    cmdRemoveDefinition.Enabled = True
  End Sub

  Private Sub lstMultiImportSelected_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMultiImportSelected.DoubleClick
    cmdRemoveDefinition_Click(Me, New EventArgs)
  End Sub

  Private Sub lstMultiImportSelected_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstMultiImportSelected.MouseDown
    mvSelected = lstMultiImportSelected.SelectedIndex
  End Sub

  Private Sub lstMultiImportSelected_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstMultiImportSelected.MouseUp
    mvSelected = -1
  End Sub

  Private Sub lstMultiImportSelected_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstMultiImportSelected.MouseMove
    Dim vIndex As Integer = lstMultiImportSelected.SelectedIndex
    Dim vValue As LookupItem

    If (vIndex <> mvSelected) AndAlso mvSelected > -1 Then
      vValue = DirectCast(lstMultiImportSelected.Items(mvSelected), LookupItem)
      lstMultiImportSelected.Items.RemoveAt(mvSelected)
      lstMultiImportSelected.Items.Insert(vIndex, vValue)
      mvSelected = vIndex
      mvImportDefinitionsChanged = True
    End If
  End Sub

  ''' <summary>
  ''' Checks if the import type that the user is trying to save is already defined
  ''' in the selected file. The check is not performed for multiple data import runs.
  ''' </summary>
  ''' <param name="pNewDefFileName"></param>
  ''' <returns>
  ''' vbYes:    Append to existing file
  ''' vbNo:     Over-write the file
  ''' vbCancel: Cancel the Save
  ''' </returns>
  ''' <remarks></remarks>
  Private Function CheckSingleImportFile(ByVal pNewDefFileName As String) As DialogResult
    Dim vResult As DialogResult = System.Windows.Forms.DialogResult.Yes
    Dim vFound As Boolean
    Dim vRow As DataRow() = Nothing

    If mvMultipleDataImportRuns = False Then
      'This is a new import type being added to an existing def file
      'So first see if this is already a multiple import run file
      Dim vParams As New ParameterList(True)
      vParams("IgnoreUnknownParameters") = CBoolYN(True)
      vParams("FileName") = pNewDefFileName
      vParams("SplitDefFile") = CBoolYN(True)
      Dim vDataSet As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaCheckSingleImportFile)
      Dim vTable As DataTable = vDataSet.Tables(DEF_FILE_PARAMS)
      If vTable IsNot Nothing Then
        'See if current type already defined
        vRow = vTable.Select(String.Format("ImportType = '{0}'", SelectedImportType))
        If vRow.Length > 0 Then vFound = True
        If vTable.Rows.Count > 1 Then
          CopyTable(vDataSet, DataImportDS, DEF_FILE_PARAMS)
          mvMultipleDataImportRuns = True
        End If
      End If

      If vFound Then
        vResult = ShowQuestion(InformationMessages.ImImportAlreadyDefined, MessageBoxButtons.OKCancel, vRow(0)("ImportTypeDesc").ToString)
        If vResult = System.Windows.Forms.DialogResult.OK Then
          'If multi-import file then append, otherwise over-write
          vResult = DirectCast(IIf(mvMultipleDataImportRuns, System.Windows.Forms.DialogResult.Yes, System.Windows.Forms.DialogResult.No), System.Windows.Forms.DialogResult)
        End If
      End If
    End If
    Return vResult
  End Function

  ''' <summary>
  ''' Save the current import definition into the selected file
  ''' </summary>
  ''' <param name="pNewDefFileName">The file to save to</param>
  ''' <param name="pSaveMainDefFile">
  ''' True : Save the changes to the original file
  ''' False: Save the changes to a temp file
  ''' </param>
  ''' <remarks></remarks>
  Private Sub SaveMultipleImportFile(ByVal pNewDefFileName As String, ByVal pSaveMainDefFile As Boolean)
    If mvImportDefinitionsChanged Then
      ReorderMultipleImportDefinitionFile()
    End If

    Dim vImportType As String = cboType.Text
    If SelectedImportType <> mvSelectedImportType Then vImportType = mvSelectedImportTypeDesc 'Just changed the combo
    Dim vTables As String = String.Format("{0},{1},{2},{3},{4},{5},{6},{7}", DATA_IMPORT, DATA_IMPORT_PARAMS, ATTRIBUTE_COL, DEFAULTS_ROW, MAPPED_ATTRIBUTES, MAPPED_ATTRIBUTE_COLUMNS, DEF_FILE_PARAMS, ATTR_DATE_FORMAT)
    Dim vParams As ParameterList = GetParameterList(vTables)
    vParams("FileName") = IIf(GetValue(DATA_IMPORT_PARAMS, "DefFileName").Length > 0, GetValue(DATA_IMPORT_PARAMS, "DefFileName"), GetValue(DATA_IMPORT_PARAMS, "FileName")).ToString
    vParams("DefFileName") = pNewDefFileName
    vParams("NoFileName") = CBoolYN(chkNoFileName.Checked)
    vParams("MainImportFileName") = mvMainImportFileName
    vParams("SaveMainDefFile") = CBoolYN(pSaveMainDefFile)
    'If Not pSaveMainDefFile Then vParams("ImportType") = mvSelectedImportTypeDesc
    vParams("ImportType") = mvSelectedImportTypeDesc
    vParams("LoadedMultiple") = CBoolYN(mvLoadedMultiple)
    Dim vResult As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaSaveDefinitionFile)
    CopyTable(vResult, DataImportDS, DEF_FILE_PARAMS)
    mvImportFilename = ExtractDefForMultiImport(SelectedImportType)
    SetValue(DATA_IMPORT_PARAMS, "DefFileName", mvImportFilename)
    mvMultipleDataImportRuns = True
    SetValue(DATA_IMPORT, "MultipleDataImportRuns", CBoolYN(True))
  End Sub

  ''' <summary>
  ''' Set the import types in the same order as the list
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ReorderMultipleImportDefinitionFile()
    Dim vNewTypes As New List(Of DataRow)
    Dim vTypes As DataTable = DataImportDS.Tables(DEF_FILE_PARAMS)
    Dim vType As DataRow = Nothing
    Dim vNewType As LookupItem = Nothing
    Dim vRow As DataRow = Nothing

    'Create a list that holds the data rows in the order that they appear in the list
    For vIndex As Integer = 0 To lstMultiImportSelected.Items.Count - 1
      vNewType = DirectCast(lstMultiImportSelected.Items(vIndex), LookupItem)
      vType = SelectRow(MAIN, DEF_FILE_PARAMS, "ImportType = '{0}'", vNewType.LookupCode)
      vRow = vTypes.NewRow
      vRow.ItemArray = vType.ItemArray
      vNewTypes.Add(vRow)
    Next

    'remove existing rows and add the new ones
    vTypes.Rows.Clear()
    For Each vType In vNewTypes
      vTypes.Rows.Add(vType)
    Next
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Try
      If mvImportForm Is Nothing Then
        SetValue(DATA_IMPORT, "Mode", 2) 'dimImport
        'Belt and Braces
        If ValidateControls() Then
          Dim vParams As New ParameterList(True)
          vParams("ShowStatus") = "Y"
          Dim vReply As DialogResult = FormHelper.ScheduleTask(vParams)

          If vReply <> vbCancel Then
            If mvMultipleDataImportRuns Then
              SaveMultipleImportFile(mvMainDefFileName, True)
              vParams("DefinitionFile") = mvMainDefFileName
            Else
              vParams("DefinitionFile") = mvImportFilename
            End If

            If vReply = System.Windows.Forms.DialogResult.Yes Then
              ScheduleImport(vParams)
            Else
              If vParams.ContainsKey("ShowTaskStatus") AndAlso vParams("ShowTaskStatus") = "Y" Then
                If mvTaskInfo Is Nothing Then mvTaskInfo = New frmTaskInfo(CareNetServices.TaskJobTypes.tjtDataImport)
                SplitControl.Panel2Collapsed = False
                mvTaskInfo.StartTimer()
                mvTaskInfo.DoRefresh()
              End If
              DoImport()
            End If
          End If
        End If
      Else
        mvImportForm.dgr.ColumnHeading(mvImportForm.dgr.ActiveColumn) = "Map:" & mvImportForm.MappedAttribute("FileName").ToString
        mvImportForm.MappedAttribute("ColumnNumber") = mvImportForm.dgr.ActiveColumn
        SaveMapAttribute()
        Me.Close()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
      ImportComplete() 'Resets the controls so the user can try again
    End Try
  End Sub

  Private Function ValidateControls() As Boolean
    Dim vValid As Boolean
    Dim vMsg As String = String.Empty

    'First check we have the basics
    If (Not txtSource.IsValid OrElse txtSource.Text.Length = 0) AndAlso GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") <> 12 Then 'ditTableImport
      ShowWarningMessage(InformationMessages.ImSourceCodeRequired)
    ElseIf ValidateDataSource(True) AndAlso (Not txtDataSource.IsValid OrElse txtDataSource.Text.Length = 0) Then
      ShowWarningMessage(InformationMessages.ImDataSourceRequired)
    Else
      vValid = True
    End If

    'Now validate that we have enough attributes to proceed
    If vValid Then
      If vValid AndAlso GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 8 Then 'ditFinancialHistory
        If txtNumberOfDays.Text.Length > 0 Then
          If IntegerValue(txtNumberOfDays.Text) < 0 OrElse IntegerValue(txtNumberOfDays.Text) > 14 Then
            vValid = False
            ShowWarningMessage(InformationMessages.ImTableImportNoOfDaysRange)
          End If
        End If
      End If

      If vValid Then
        If IntegerValue(txtControlNumberBlockSize.Text) < 1 OrElse IntegerValue(txtControlNumberBlockSize.Text) > 100 Then
          vValid = False
          ShowWarningMessage(InformationMessages.ImControlNumberBlockSize)
        End If
      End If

      If vValid AndAlso mvImportDefinitionsChanged Then
        vValid = ValidateMultipleImportList()
      End If

      If vValid Then
        If mvMultipleDataImportRuns Then
          'Need to validate each import type in the file
          SaveMultipleImportFile(mvMainDefFileName, False)   'Save what is on the screen first
          Dim vParams As New ParameterList(True)
          vParams("IgnoreUnknownParameters") = CBoolYN(True)
          vParams("MultipleImportValidate") = CBoolYN(True)
          For Each vRow As DataRow In DataImportDS.Tables(DEF_FILE_PARAMS).Rows
            If mvMainImportFileName.Length > 0 Then
              vParams("FileName") = mvMainImportFileName
            Else
              vParams("FileName") = vRow("FileName").ToString
            End If
            Dim vDataSet As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaValidate)
            Dim vResult As DataTable = vDataSet.Tables(VALIDATE_IMPORT)
            vValid = BooleanValue(vResult.Rows(0)("Valid").ToString)
            If vValid = False Then Exit For
          Next

          'If vValid = False Then cboType.Text = vParam.Name 'Force reloading of the correct temp def file
        Else
          Dim vParams As New ParameterList(True)
          Dim vTables As String = String.Format("{0},{1},{2},{3},{4},{5}", DATA_IMPORT, DATA_IMPORT_PARAMS, ATTRIBUTE_COL, MAPPED_ATTRIBUTES, MAPPED_ATTRIBUTE_COLUMNS, DEFAULTS_ROW)
          vParams = GetParameterList(vTables)
          vParams("ImportType") = cboType.Text
          If mvMainImportFileName.Length > 0 Then vParams("MainImportFileName") = mvMainImportFileName
          Dim vDataSet As DataSet = DataHelper.InitDataImport(vParams, CareNetServices.DataImportAction.diaValidate)
          Dim vResult As DataTable = vDataSet.Tables(VALIDATE_IMPORT)
          vValid = BooleanValue(vResult.Rows(0)("Valid").ToString)
          vMsg = vResult.Rows(0)("Message").ToString
        End If
        If Not vValid Then
          ErrorSetFocus(vMsg, cboAttrs)
        Else
          If vValid AndAlso vMsg.Length > 0 Then ShowInformationMessage(vMsg)
        End If
      End If
    End If
    Return vValid
  End Function

  Private Sub ErrorSetFocus(ByVal pMsg As String, ByVal pControl As Control)
    ShowWarningMessage(pMsg)
    If pControl.Enabled AndAlso pControl.Visible Then pControl.Focus()
  End Sub

  Private Function ValidateDataSource(ByVal pCheckRequiredOnly As Boolean) As Boolean
    Dim vValidate As Boolean
    Dim vIndex As Integer

    With DataImportDS
      For Each vRow As DataRow In DataImportDS.Tables(DATA_IMPORT_ATTRS).Rows
        vIndex = IntegerValue(vRow("ID"))
        If vRow("AttributeName").ToString = "external_reference" Then
          If pCheckRequiredOnly Then
            If GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") <> 12 Then 'ditTableImport
              'TODO: No dependant attribute property on validation item. Maybe its added after conversion
              'If IsSet(vIndex) AndAlso Not IsSet(.Attributes(vIndex).DependantAttributes(1).ID) Then
              If IsSet(vIndex) Then
                vValidate = True
                Exit For
              End If
            End If
          Else
            If IsSet(vIndex) Then
              vValidate = True
              Exit For
            End If
          End If
        End If
      Next
      Return vValidate
    End With
  End Function

  ''' <summary>
  ''' Checks if the attribtue is mapped or if a default value is specified
  ''' </summary>
  ''' <param name="pIndex">Attribute ID</param>
  ''' <returns>True : if the attribute is in use</returns>
  ''' <remarks></remarks>
  Public Function IsSet(ByVal pIndex As Integer) As Boolean
    Dim vSet As Boolean
    Dim vRow As DataRow = Nothing
    'Check if attribute is mapped to a column on the file
    vRow = SelectRow(MAIN, ATTRIBUTE_COL, "AttributeIndex = '{0}'", pIndex)
    If vRow IsNot Nothing Then
      vSet = True
    Else
      'Check if a default value is specified
      vRow = SelectRow(MAIN, DEFAULTS_ROW, "ID = '{0}'", pIndex)
      If vRow IsNot Nothing Then
        vSet = True
      Else
        'Check if the attribute is mapped to some other file
        If DataImportDS.Tables(MAPPED_ATTRIBUTE_COLUMNS) Is Nothing Then
          vRow = SelectRow(MAIN, MAPPED_ATTRIBUTE_COLUMNS, "ID = '{0}'", pIndex)
          If vRow IsNot Nothing Then vSet = True
        End If
      End If
    End If
    Return vSet
  End Function

  Public Sub DoCount()
    Try
      cmdOk.Visible = False
      cmdCancel.Enabled = False
      cmdStop.Visible = True
      cmdStop.Enabled = True
      cmdTest.Enabled = False
      bpl.RepositionButtons()

      Dim vParams As New ParameterList(True)
      vParams("JobName") = FormHelper.GetTaskJobTypeName(CareServices.TaskJobTypes.tjtDataImport)
      vParams("FileName") = mvImportFilename
      If mvMasterDataImport Is Nothing Then
        vParams("Seperator") = GetValue(DATA_IMPORT_PARAMS, "Separator")
      Else
        vParams("Seperator") = GetValue(DATA_IMPORT, "MapSeparator", MASTER)
      End If
      vParams("CountOnly") = CBoolYN(True)
      mvJobID = GetUniqueJobId()
      vParams("JobId") = mvJobID
      SetValue(DATA_IMPORT, "Mode", 0) 'dimCount
      If mvMainImportFileName.Length > 0 Then vParams("MainImportFileName") = mvMainImportFileName
      'Start counting async
      Dim vProcessor As New AsyncProcessHandler(CareServices.TaskJobTypes.tjtDataImport, vParams)
      AddHandler vProcessor.ProcessCompleted, AddressOf ProcessJobCompleted
      vProcessor.ProcessJob()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub ProcessJobCompleted(ByVal pJob As AsyncProcessHandler)
    Dim vRow As DataRow = Nothing
    Dim vResult As String = String.Empty
    If pJob.ResultTable IsNot Nothing AndAlso pJob.ResultTable.Rows.Count > 0 Then
      vRow = pJob.ResultTable.Rows(0)
      vResult = vRow("ResultStatus").ToString
    End If

    Select Case GetIntegerValue(DATA_IMPORT, "Mode")
      Case 0 'dimCount
        If vResult.Length > 0 AndAlso Not vResult.Contains("ERROR") Then
          If mvImportForm Is Nothing Then
            Me.Text = ControlText.FrmImport
          Else
            Me.Text = ControlText.FrmImportMapFile
          End If
          Dim vRowCount As String() = vResult.Split(":"c) 'result format - TaskName: Count eg. DataImport:4
          SetValue(DATA_IMPORT, "Rows", vRowCount(1))
          Me.Text = Me.Text & String.Format("  {0}  {1} Rows", GetValue(DATA_IMPORT_PARAMS, "FileName"), vRowCount(1))
        End If

        CountCompleted()
        SetValue(DATA_IMPORT, "Mode", 2) 'dimImport

      Case 2, 3 'dimImport,dimImportTest
        If vResult.Length > 0 Then ShowInformationMessage(vResult)
        ImportComplete()
    End Select
  End Sub

  Private Sub ImportComplete()
    cmdCancel.Enabled = True
    cmdOk.Visible = True
    cmdStop.Visible = False
    cmdTest.Visible = True
    bpl.RepositionButtons()
    'Stop the timer and refresh the status labels 
    If mvTaskInfo IsNot Nothing Then
      mvTaskInfo.StopTimer()
      lblJobNumber.Text = ""
      lblStatus.Text = ""
    End If
  End Sub

  ''' <summary>
  ''' Resets the controls once the count operation has completed or aborted
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CountCompleted()
    cmdOk.Visible = True
    cmdStop.Visible = False
    cmdCancel.Enabled = True
    cmdTest.Enabled = True
    bpl.RepositionButtons()
  End Sub

  Private Sub DoImport()
    Dim vReturn As DialogResult
    Dim vDefFileName As String = String.Empty
    Dim vExtractDup As Boolean
    Dim vExtractUnProcPayments As Boolean

    cmdCancel.Enabled = False
    cmdOk.Visible = False
    If GetIntegerValue(DATA_IMPORT, "Mode") = 3 Then 'dimImportTest
      cmdStop.Text = ControlText.CmdStopTest
    Else
      cmdStop.Text = ControlText.CmdStopImport
    End If
    cmdStop.Visible = True
    cmdTest.Visible = False
    bpl.RepositionButtons()
    vDefFileName = GetValue(DATA_IMPORT_PARAMS, "DefFileName")

    If mvMultipleDataImportRuns Then
      If mvMainDefFileName.Length > 0 Then vDefFileName = mvMainDefFileName
    End If

    Dim vParams As New ParameterList(True)
    vParams("DefinitionFile") = vDefFileName
    If GetIntegerValue(DATA_IMPORT, "Mode") = 3 Then 'dimImportTest
      vParams("Mode") = "Test"
    Else
      vParams("Mode") = "Import"
    End If

    If mvMultipleDataImportRuns Then
      For Each vImportType As DataRow In DataImportDS.Tables(DEF_FILE_PARAMS).Rows
        Select Case IntegerValue(vImportType("ImportType"))
          Case 7, 18  'ditPaymentPlan = 7, ditCMT = 18
            vExtractUnProcPayments = True
        End Select
        If vExtractUnProcPayments Then Exit For
      Next
    Else
      Select Case GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType")
        Case 7, 18  'ditPaymentPlan = 7, ditCMT = 18
          vExtractUnProcPayments = True
      End Select
    End If
    If vExtractUnProcPayments Then vReturn = ShowQuestion(QuestionMessages.QmExtractUnProcPayErrorRecords, MessageBoxButtons.YesNo)
    If vReturn = System.Windows.Forms.DialogResult.Yes Then
      vParams("UnProcessedPaymentsErrorFile") = "Yes"
      vExtractUnProcPayments = True
    Else
      vParams("UnProcessedPaymentsErrorFile") = "No"
    End If

    vReturn = ShowQuestion(QuestionMessages.QmExtractErrorRecords, MessageBoxButtons.YesNo)
    If vReturn = System.Windows.Forms.DialogResult.Yes Then
      vParams("ErrorFile") = "Yes"
    Else
      vParams("ErrorFile") = "No"
    End If

    If mvMultipleDataImportRuns Then
      For Each vImportType As DataRow In DataImportDS.Tables(DEF_FILE_PARAMS).Rows
        Select Case IntegerValue(vImportType("ImportType"))
          Case 0, 2, 5, 12 'ditContactOrganisation, ditActivity, ditSuppression, ditTableImport
            'Always ask this question during multiple import runs if it contains 
            'any of the above import types.
            'The overhead of loading all the imports just to check one value seems a bit high
            vExtractDup = True
        End Select
        If vExtractDup Then Exit For
      Next
    Else
      'ditContactOrganisation, ditActivity, ditSuppression, ditTableImport
      If (GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 0 _
      OrElse GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 2 _
      OrElse GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 5 _
      OrElse GetIntegerValue(DATA_IMPORT_PARAMS, "DataImportType") = 12) _
      AndAlso GetBooleanValue(DATA_IMPORT_PARAMS, "LogDups") Then
        vExtractDup = True
      End If
    End If

    If vExtractDup Then vReturn = ShowQuestion(QuestionMessages.QmExtractDuplicateRecords, MessageBoxButtons.YesNo)
    If vReturn = System.Windows.Forms.DialogResult.Yes Then
      vParams("DupFile") = "Yes"
    Else
      vParams("DupFile") = "No"
    End If

    vParams("CountRecords") = "No"
    vParams("NoOfRows") = GetValue(DATA_IMPORT, "Rows")
    vParams("CountingAborted") = IIf(GetBooleanValue(DATA_IMPORT, "CountingAborted"), "Yes", "No").ToString

    If mvMultipleDataImportRuns Then
      SetValue(DATA_IMPORT_PARAMS, "DefFileName", mvMainDefFileName)
      vParams("TempDefFile") = mvMainDefFileName
    Else
      SetValue(DATA_IMPORT_PARAMS, "FirstLoad", "Y") ' Will override IgnoreFirstRow if set to N, IgnoreFirstRow critical when creating Temp Def file. BR17795
      Dim vTables As String = String.Format("{0},{1},{2},{3},{4},{5}", DATA_IMPORT, DATA_IMPORT_PARAMS, ATTRIBUTE_COL, DEFAULTS_ROW, MAPPED_ATTRIBUTES, MAPPED_ATTRIBUTE_COLUMNS)
      Dim vList As ParameterList = GetParameterList(vTables)
      vList("FileName") = IIf(vDefFileName.Length > 0, vDefFileName, GetValue(DATA_IMPORT_PARAMS, "FileName")).ToString
      vList("DefFileName") = "TEMP_FILE" 'server will create a temp file and save to that. Result will contain the name of the temp file.
      vList("IncludeSectionHeaders") = CBoolYN(True)
      vList("AppendFile") = CBoolYN(False)
      vList("SaveFileName") = CBoolYN(True)
      vList("ImportType") = cboType.Text
      If mvMainImportFileName.Length > 0 Then vList("MainImportFileName") = mvMainImportFileName
      Dim vDataSet As DataSet = DataHelper.InitDataImport(vList, CareNetServices.DataImportAction.diaSaveDefinitionFile)
      SetValue(DATA_IMPORT_PARAMS, "DefFileName", vDefFileName)
      'Get the temp file location on the server
      vParams("TempDefFile") = vDataSet.Tables(0).Rows(0)("FileName").ToString
    End If
    If mvMainImportFileName.Length > 0 Then vParams("MainImportFileName") = mvMainImportFileName

    vParams("JobName") = FormHelper.GetTaskJobTypeName(CareServices.TaskJobTypes.tjtDataImport)
    mvJobID = GetUniqueJobId()
    vParams("JobId") = mvJobID
    Dim vProcessor As New AsyncProcessHandler(CareServices.TaskJobTypes.tjtDataImport, vParams)
    AddHandler vProcessor.ProcessCompleted, AddressOf ProcessJobCompleted
    vProcessor.ProcessJob()

    'Catch vEx As Exception
    '    If Err = 32755 Then
    '      'Cancel from MakeErrorFile
    '      Resume DoImportMEFCancel
    '    End If
  End Sub

  ''' <summary>
  ''' This function should be used to generate a unique id for a job that is started
  ''' asynchronously. This id should be used to abort the job if necessary.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetUniqueJobId() As String
    Return DataHelper.UserInfo.Logname & DateTime.Now.ToString("ddMMyyyyHHmmss")
  End Function

  Private Sub cmdStop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdStop.Click
    Try
      Dim vParams As New ParameterList(True)
      vParams("JobId") = mvJobID
      Dim vResult As ParameterList = DataHelper.AbortJob(vParams)
      If BooleanValue(vResult("AbortRequested")) Then
        ShowInformationMessage(InformationMessages.ImAbortRequested)
        SetValue(DATA_IMPORT, "CountingAborted", True)
        CountCompleted()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdTest_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdTest.Click
    Try
      SetValue(DATA_IMPORT, "Mode", 3) 'dimImportTest
      If ValidateControls() Then
        Dim vParams As New ParameterList(True)
        Dim vReply As DialogResult = FormHelper.ScheduleTask(vParams)

        If vReply <> System.Windows.Forms.DialogResult.Cancel Then
          If mvMultipleDataImportRuns Then
            SaveMultipleImportFile(mvMainDefFileName, True)
            vParams("DefinitionFile") = mvMainDefFileName
          Else
            vParams("DefinitionFile") = mvImportFilename
          End If

          If vReply = System.Windows.Forms.DialogResult.Yes Then
            ScheduleImport(vParams)
          Else
            DoImport()
          End If
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
      ImportComplete() 'Resets the controls so the user can try again
    End Try
  End Sub

  Private Sub ScheduleImport(ByVal pParams As ParameterList)
    Dim vDefaults As ParameterList = New ParameterList(True)
    Dim vParamList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.TaskJobTypes.tjtDataImport, vDefaults)
    If vParamList.Count > 0 Then 'user has not clicked cancel
      For Each vKey As String In pParams.Keys
        vParamList(vKey) = pParams(vKey)
      Next

      If Not vParamList("DefinitionFile").EndsWith(".def") Then
        'Def file has not been saved and the user has clicked schedule.
        'Save a def file with the same name. This is required to run later at the scheduled time.
        vParamList("DefinitionFile") = SaveDefFile()
      End If
      vParamList("DefinitionFile") = FileNameInQuotes(vParamList("DefinitionFile"))
      If GetIntegerValue(DATA_IMPORT, "Mode") = 3 Then 'dimImportTest
        vParamList("Mode") = "Test"
      Else
        vParamList("Mode") = "Import"
      End If
      vParamList("ErrorFile") = "Yes"
      vParamList("DupFile") = "Yes"
      vParamList("UnProcessedPaymentsErrorFile") = "Yes"
      If BooleanValue(vParamList("CountRecords")) Then
        vParamList("CountRecords") = "Yes"
        vParamList("NoOfRows") = GetValue(DATA_IMPORT, "Rows")
        vParamList("CountingAborted") = "No"
      Else
        vParamList("CountRecords") = "No"
        vParamList("NoOfRows") = "1"
        vParamList("CountingAborted") = "Yes"
      End If
      If mvMainImportFileName.Length > 0 Then vParamList("MainImportFileName") = FileNameInQuotes(mvMainImportFileName)
      FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDataImport, vParamList, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysSchedule, False)
    End If
  End Sub

  Public Overloads Sub Show()
    If mvCanDisplayForm Then
      MyBase.Show()
    Else
      Me.Dispose()
    End If
  End Sub

  Private Sub RefreshStatus_Click(ByVal sender As System.Object, ByVal pJobNumber As String, ByVal pJobStatus As String) Handles mvTaskInfo.RefreshStatus
    lblJobNumber.Text = pJobNumber
    lblStatus.Text = pJobStatus
  End Sub

  Private Sub chkOrgNamePostCodeAddress_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkOrgNamePostCodeAddress.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "OrgNamePostCodeAddressDup", chkOrgNamePostCodeAddress.Checked)
  End Sub
  Private Sub chkAllowBlankForOrgName_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAllowBlankForOrgName.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "AllowBlankOrganisation", chkAllowBlankForOrgName.Checked)
  End Sub
  Private Sub chkExclUnkAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkExclUnkAdd.CheckedChanged
    SetValue(DATA_IMPORT_PARAMS, "DeDupExclUnkAdd", chkExclUnkAdd.Checked)
  End Sub

  ''' <summary>
  ''' Encloses a string in Quotes
  ''' </summary>
  ''' <param name="pFileName">String to Enclose</param>
  ''' <returns>Enclosed string</returns>
  ''' <remarks>If the string is already enclosed in quotes it is returned unchanged</remarks>
  Private Function FileNameInQuotes(ByVal pFileName As String) As String
    Dim vReturn As String
    If pFileName.StartsWith("""") AndAlso pFileName.EndsWith("""") Then
      vReturn = pFileName
    Else
      vReturn = """" & pFileName & """"
    End If
    Return vReturn
  End Function
  Private Function ListDuplicateColumnIndex(ByVal pTable As DataTable) As List(Of Integer)
    'Returns a list of ColumnIndex values that exist in more than 1 row
    Dim vDuplicateIndex As New List(Of Integer)
    For vIndex As Integer = 0 To pTable.Rows.Count - 2
      If IntegerValue(pTable.Rows(vIndex)("ColumnIndex")) >= 0 Then
        For vIndex2 As Integer = vIndex + 1 To pTable.Rows.Count - 1
          If IntegerValue(pTable.Rows(vIndex2)("ColumnIndex")) = IntegerValue(pTable.Rows(vIndex)("ColumnIndex")) Then
            If vDuplicateIndex.Count = 0 Then
              vDuplicateIndex.Add(IntegerValue(pTable.Rows(vIndex)("ColumnIndex")))
            Else
              If Not vDuplicateIndex.Contains(IntegerValue(pTable.Rows(vIndex)("ColumnIndex"))) Then
                vDuplicateIndex.Add(IntegerValue(pTable.Rows(vIndex)("ColumnIndex")))
              End If
            End If
          End If
        Next vIndex2
      End If
    Next vIndex
    Return vDuplicateIndex
  End Function

  Private Sub AttrsSort()
    mvAttrItems.Sort(LookupItemComparer)
  End Sub

  Public Class ImportAttributeComparer
    Implements IComparer(Of LookupItem)


    Public Function Compare(vSource As LookupItem, vOther As LookupItem) As Integer Implements IComparer(Of LookupItem).Compare
      Dim vResult As Integer = 0
      Dim vSourceText As String = vSource.LookupDesc
      Dim vOtherDesc As String = vOther.LookupDesc
      Dim vTokenMatch As String = "\(\w+\)", vNumberMatch As String = "\d+"
      Dim vSourceToken As String = Regex.Match(vSourceText, vTokenMatch).Value
      Dim vOtherToken As String = Regex.Match(vOtherDesc, vTokenMatch).Value
      Dim vSourceSetNo As String = Regex.Match(vSourceToken, vNumberMatch).Value
      Dim vOtherSetNo As String = Regex.Match(vOtherToken, vNumberMatch).Value
      Dim vSourceSet As Integer = 0
      Dim vOtherSet As Integer = 0
      If IsNumeric(vSourceSetNo) OrElse IsNumeric(vOtherSetNo) Then
        Integer.TryParse(vSourceSetNo, vSourceSet)
        Integer.TryParse(vOtherSetNo, vOtherSet)
        vResult = vSourceSet.CompareTo(vOtherSet)
      End If

      If vResult = 0 Then
        vResult = vSource.ToString().CompareTo(vOther.ToString())
      End If

      Return vResult
    End Function
  End Class

  Private Sub lnkHelp_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkHelp.LinkClicked
    Dim vHelpUri As New UriBuilder(DataHelper.HelpBaseURL)
    vHelpUri.Path += "/mergedProjects/sc_sysadm_guide/_landing_pad/import_section_toc.htm"
    Dim vForm As New frmBrowser(vHelpUri.ToString(), False, True)
    vForm.Show()
  End Sub
End Class

