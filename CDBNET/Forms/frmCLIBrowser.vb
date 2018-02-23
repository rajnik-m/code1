Public Class frmCLIBrowser

  Private mvEditPanelInfo As EditPanelInfo
  Private mvPhoneNumber As String
  Private mvStartTime As Date

  Public Sub New(ByVal pDataSet As DataSet, ByVal pOrganisations As Boolean, ByVal pPhoneNumber As String)
    ' This call is required by the designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pDataSet, pOrganisations)
    mvPhoneNumber = pPhoneNumber
    mvStartTime = Now
  End Sub

  Private Sub InitialiseControls(ByVal pDataSet As DataSet, ByVal pOrganisations As Boolean)
    SetControlTheme()
    SettingsName = "CLIBrowser"
    mvEditPanelInfo = New EditPanelInfo(CareNetServices.FunctionParameterTypes.fptCLIBrowser)
    For Each vPanelItem As PanelItem In mvEditPanelInfo.PanelItems
      vPanelItem.Mandatory = False
    Next
    epl.Init(mvEditPanelInfo)
    If epl.Caption.Length > 0 Then Me.Text = epl.Caption
    bpl.RepositionButtons()
    spl.SplitterDistance = spl.Height - epl.RequiredHeight

    PopulateFromDataSet(pDataSet, pOrganisations)
  End Sub

  Private Sub PopulateFromDataSet(ByVal pDataSet As DataSet, ByVal pOrganisations As Boolean)
    cboOrganisations.DisplayMember = "Name"
    cboOrganisations.ValueMember = "OrganisationNumber"
    If pOrganisations Then
      cboOrganisations.DataSource = pDataSet.Tables("DataRow")
    Else
      cboOrganisations.DataSource = Nothing
      cboAddresses.DataSource = Nothing
      If pDataSet IsNot Nothing Then
        dgr.Populate(pDataSet)
        dgr.DisableCustomise()
        dgr.Focus()
      Else
        dgr.Clear()
        epl.Focus()
      End If
    End If
    cmdSelectOrg.Enabled = pOrganisations
    cmdSelect.Enabled = (dgr.RowCount > 0)
  End Sub

  Private Sub cboOrganisations_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboOrganisations.SelectedIndexChanged
    cboAddresses.DisplayMember = "AddressLine"
    cboAddresses.ValueMember = "AddressNumber"
    If cboOrganisations.SelectedIndex >= 0 Then
      Dim vContactNumber As Integer = IntegerValue(cboOrganisations.SelectedValue.ToString)
      DataHelper.FillComboWithContactData(cboAddresses, CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vContactNumber, True)
      If cboAddresses.DataSource IsNot Nothing Then DirectCast(cboAddresses.DataSource, DataTable).Rows(0).Item("AddressLine") = "<All Addresses>"
    Else
      cboAddresses.DataSource = Nothing
    End If
    cmdSelectOrg.Enabled = (cboOrganisations.SelectedIndex >= 0)
  End Sub

  Private Sub cboAddresses_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAddresses.SelectedIndexChanged
    Dim vAllowSelectContact As Boolean = True
    If cboOrganisations.SelectedIndex >= 0 AndAlso cboAddresses.SelectedIndex >= 0 Then
      Dim vContactNumber As Integer = IntegerValue(cboOrganisations.SelectedValue.ToString)
      Dim vAddressNumber As Integer = IntegerValue(cboAddresses.SelectedValue.ToString)
      Dim vList As New ParameterList(True)
      If vAddressNumber > 0 Then vList.IntegerValue("AddressNumber") = vAddressNumber
      vList("Current") = "Y"
      Dim vDataSet As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vContactNumber, vList)
      dgr.Populate(vDataSet)
      dgr.DisableCustomise()
      If dgr.RowCount = 0 Then vAllowSelectContact = False
    Else
      dgr.Clear()
    End If
    cmdSelect.Enabled = vAllowSelectContact
  End Sub

  Private Sub cmdClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    epl.Clear()
  End Sub

  Private Sub cmdFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFind.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vCount As Integer = vList.Count
      epl.AddValuesToList(vList, False, EditPanel.AddNullValueTypes.anvtNone)
      If vList.Count = vCount Then
        ShowInformationMessage(InformationMessages.ImInsufficientDetails)
      Else
        If vList.ContainsKey("Surname") Then
          Dim vDataSet As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
          PopulateFromDataSet(vDataSet, False)
        ElseIf vList.ContainsKey("Name") Then
          Dim vDataSet As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftOrganisations, vList)
          PopulateFromDataSet(vDataSet, True)
        End If
      End If
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Private Sub cmdNew_Click(ByVal sender As Object, ByVal e As System.EventArgs)

  End Sub

  Private Sub cmdSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
    'Select Contact
    Try
      If dgr.RowCount > 0 Then
        If dgr.CurrentDataRow >= 0 Then
          Dim vContactNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentDataRow, dgr.GetColumn("ContactNumber")))
          ProcessContactSelected(vContactNumber)
        End If
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdSelectOrg_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSelectOrg.Click
    'Select Organisation
    Try
      If cboOrganisations.SelectedIndex >= 0 Then
        Dim vOrgNumber As Integer = IntegerValue(cboOrganisations.SelectedValue.ToString)
        ProcessContactSelected(vOrgNumber)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub dgr_ContactSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    Try
      ProcessContactSelected(pContactNumber)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub dgr_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgr.DoubleClick

  End Sub

  Private Sub ProcessContactSelected(ByVal pContactNumber As Integer)
    Dim vForm As Form = FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, pContactNumber, False, True)
    If vForm IsNot Nothing Then
      Dim vTCRList As New ParameterList
      vTCRList("Direction") = "I"
      vTCRList("Precis") = GetInformationMessage(InformationMessages.ImTcrPrecis, mvPhoneNumber, mvStartTime.ToString(AppValues.DateFormat), mvStartTime.ToString("HH:mm:ss"))
      Dim vTCRForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, vTCRList)
      vTCRForm.RelatedContact = DirectCast(vForm, frmCardSet).ContactInfo()
      vTCRForm.Show()
    End If
    Me.Close()
  End Sub

  Private Sub epl_ButtonClicked(ByVal sender As Object, ByVal pParameterName As String) Handles epl.ButtonClicked
    Select Case pParameterName
      Case "FindContact"
        Dim vContactNumber As Integer = FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftContacts, Nothing, Me)
        If vContactNumber > 0 Then ProcessContactSelected(vContactNumber)
      Case "FindOrganisation"
        Dim vOrganisationNumber As Integer = FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftOrganisations, Nothing, Me)
        If vOrganisationNumber > 0 Then
          Dim vList As New ParameterList(True)
          vList.IntegerValue("OrganisationNumber") = vOrganisationNumber
          Dim vDataSet As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftOrganisations, vList)
          PopulateFromDataSet(vDataSet, True)
        End If
    End Select
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    If pValue.Length > 0 Then
      Me.AcceptButton = cmdFind
    Else
      Dim vList As New ParameterList
      epl.AddValuesToList(vList)
      If vList.Count > 0 Then
        Me.AcceptButton = cmdFind
      Else
        Me.AcceptButton = cmdSelect
      End If
    End If
  End Sub

  Private Sub frmCLIBrowser_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If cboOrganisations.Items.Count > 0 Then
      cboOrganisations.Focus()
    ElseIf dgr.RowCount > 0 Then
      dgr.Focus()
    Else
      Me.AcceptButton = Nothing
      epl.Focus()
    End If
  End Sub

End Class
