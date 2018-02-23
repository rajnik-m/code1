Public Partial Class AddContact
  Inherits CareWebControl
  Implements ICareParentWebControl
  Private mvOrganisationAddress As Integer

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvUsesHiddenContactNumber = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddContact, tblDataEntry)
      AddHiddenField("HiddenAddressNumber")
      AddHiddenField("HiddenSalutation")
      AddHiddenField("HiddenPreferredForename")
      AddHiddenField("HiddenOldForename")
      AddHiddenField("HiddenSurnamePrefix")
      AddHiddenField("HiddenSurname") 'existing surname value
      AddHiddenField("HiddenOldSurname") 'surname value before the most recent update
      AddHiddenField("HiddenSurname2") 'Used for capitalisation only
      AddHiddenField("HiddenTitle")
      If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.use_ajax_for_contact_names, False) Then
        AddHandlersAndTriggers(tblDataEntry)
      Else
        'On Change of Surname, HiddenSurname2 and Surname (for capitalisation) fields should be updated
        AddTextChangedHandler("Surname")
        AddAsyncPostBackTrigger("HiddenSurname2,Surname", "Surname", PostBackTriggerEventTypes.TextChanged)
      End If
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub AddCustomValidator(ByVal pHTMLTable As HtmlTable)
    Dim vControl As Control = FindControlByName(tblDataEntry, "StatusReason")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "1", "Status Reason must be entered")
    End If
  End Sub

  Public Overrides Sub ServerValidate(ByVal sender As Object, ByVal args As ServerValidateEventArgs)
    args.IsValid = ValidateContactStatus()
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    Dim vReturnList As ParameterList
    If vContactNumber > 0 Then
      Dim vParameterList As ParameterList = GetAddContactParameterList(BooleanValue(DefaultParameters("SetHistoric").ToString) = False)
      vParameterList("ContactNumber") = vContactNumber
      AddUserParameters(vParameterList)
      vReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vParameterList)
      Dim vAddressNumber As Integer = GetHiddenAddressNumber()
      If vAddressNumber > 0 Then vReturnList("AddressNumber") = vAddressNumber
    Else
      vReturnList = AddNewContact()
      If Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0 Then vReturnList("AddressNumber") = mvOrganisationAddress
      Session("CurrentContactNumber") = IntegerValue(vReturnList("ContactNumber").ToString)
      Session("CurrentAddressNumber") = IntegerValue(vReturnList("AddressNumber").ToString)
    End If
    'Get comm numbers info
    SetContactCommNumbersInfo(IntegerValue(vReturnList("ContactNumber").ToString), False)
    'Save communication numbers - Only process Additional Numbers when adding a new contact or updating a contact with historic flag not set
    Dim vAdditionalNumbersOnly As Boolean = vContactNumber = 0
    SaveContactCommNumbers(mvCommNumbers, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), vAdditionalNumbersOnly, Not vAdditionalNumbersOnly)
    ProcessChildControls(vReturnList)
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Sub SetDefaults()
    If Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0 Then
      'Get Organisation Details
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = Session("SelectedOrganisationNumber").ToString
      Dim vDataTable As DataTable = GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
      If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
        With vDataTable.Rows(0)
          SetTextBoxText("Name", .Item("OrganisationName").ToString)
          SetTextBoxText("Address", .Item("Address").ToString)
          SetControlEnabled("Name", False)
          SetControlEnabled("Address", False)
          mvOrganisationAddress = IntegerValue(vDataTable.Rows(0).Item("AddressNumber").ToString)
        End With
      End If
    End If
  End Sub
End Class
