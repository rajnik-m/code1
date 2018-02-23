Public Partial Class AddRelatedContact
  Inherits CareWebControl
  Implements ICareChildWebControl
  Implements ICareParentWebControl

  Dim mvOldRelationship As String = ""

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvHandlesLinks = True
    mvUsesHiddenContactNumber = True
    mvHideHistoricLinks = BooleanValue(DefaultParameters.OptionalValue("HideHistoricLinks"))
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddRelatedContact, tblDataEntry)
      AddHiddenField("HiddenAddressNumber")
      AddHiddenField("OldRelationship")
      AddHiddenField("OldRelationshipStatus")
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
        AddAsyncPostBackTrigger("Relationship,RelationshipStatus", "Relationship", PostBackTriggerEventTypes.SelectedIndexChanged)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub AddCustomValidator(ByVal pHTMLTable As HtmlTable)
    Dim vControl As Control = FindControlByName(tblDataEntry, "Relationship")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "1", "A Surname and Relationship must be entered")
    End If
    vControl = FindControlByName(tblDataEntry, "StatusReason")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "2", "Status Reason must be entered")
    End If
    vControl = FindControlByName(tblDataEntry, "Status")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "3", "Status must be entered")
    End If
  End Sub

  Public Overrides Sub ServerValidate(ByVal sender As Object, ByVal args As ServerValidateEventArgs)
    Dim vCFV As CustomValidator = TryCast(sender, CustomValidator)
    If vCFV.ID = "cfv1" Then        'Relationship
      args.IsValid = (GetTextBoxText("Surname").Length > 0 AndAlso GetDropDownValue("Relationship").Length > 0) OrElse (GetTextBoxText("Surname").Length = 0 AndAlso GetDropDownValue("Relationship").Length = 0)
    ElseIf vCFV.ID = "cfv2" Then    'StatusReason
      args.IsValid = ValidateContactStatus()
    ElseIf vCFV.ID = "cfv3" Then    'Status
      args.IsValid = (GetTextBoxText("Surname").Length > 0 AndAlso GetDropDownValue("Relationship").Length > 0) AndAlso ((DataHelper.ConfigurationValueOption(DataHelper.ConfigurationValues.cd_status_mandatory) = True AndAlso GetDropDownValue("Status").Length > 0) OrElse (DataHelper.ConfigurationValueOption(DataHelper.ConfigurationValues.cd_status_mandatory) = False)) OrElse (GetTextBoxText("Surname").Length = 0 AndAlso GetDropDownValue("Relationship").Length = 0)
    End If
  End Sub

  Public Overrides Sub ProcessLinkSelection(ByVal pRow As DataRow)
    SetHiddenText("OldRelationship", pRow("RelationshipCode").ToString)
    SetHiddenText("OldRelationshipStatus", pRow("RelationshipStatus").ToString)
    SetDropDownText("Relationship", pRow("RelationshipCode").ToString)
    If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
      Dim vDropDownList As DropDownList = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList)
      If pRow("RelationshipCode").ToString.Trim.Length > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("Relationship") = pRow("RelationshipCode").ToString
        vDropDownList.DataTextField = "RelationshipStatusDesc"
        vDropDownList.DataValueField = "RelationshipStatus"
        DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtRelationshipStatuses, vDropDownList, True, vList, "Relationship Is Null OR Relationship = '" & pRow("RelationshipCode").ToString & "'")
        SetDropDownText("RelationshipStatus", pRow("RelationshipStatus").ToString)
      Else
        'ToDo clear values
      End If
    End If
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    Dim vAddressNumber As Integer = GetHiddenAddressNumber()
    Dim vSalutation As String = GetHiddenText("HiddenSalutation")
    Dim vHiddenPreferredForename As String = GetHiddenText("HiddenPreferredForename")
    Dim vRelationship As String = GetDropDownValue("Relationship")
    Dim vRelationshipStaus As String = ""
    Dim vDefaultRelationship As String = ""

    If (Me.FindControl("RelationshipStatus")) IsNot Nothing Then vRelationshipStaus = GetDropDownValue("RelationshipStatus")
    If InitialParameters.ContainsKey("DefaultRelationship") Then vDefaultRelationship = InitialParameters("DefaultRelationship").ToString

    Dim vAddDefaultRelationship As Boolean = (vDefaultRelationship.Length > 0)
    Dim vAddRelationship As Boolean = (vRelationship.Length > 0)
    Dim vOldRelationship As String = GetHiddenText("OldRelationship")
    Dim vOldRelationshipStatus As String = GetHiddenText("OldRelationshipStatus")

    If vContactNumber > 0 Then
      'We have previously created the relationships and now the relatinship may have changed (the default relationship cannot change)
      vAddDefaultRelationship = False   'Never want to re-create this for the same contact
      vAddRelationship = False
      If vOldRelationship.Length > 0 AndAlso ((vRelationship.Length > 0 AndAlso String.Equals(vOldRelationship, vRelationship, StringComparison.CurrentCultureIgnoreCase) = False) _
      OrElse (String.Equals(vOldRelationshipStatus, vRelationshipStaus, StringComparison.CurrentCultureIgnoreCase) = False)) Then
        'Either both old & new relationship codes are set and they have changed
        'OR the relationship status has changed
        'Therefore update the curent relationship and do not change the default relationship
        'vAddRelationship = False
      ElseIf (vOldRelationship.Length = 0 AndAlso vRelationship.Length > 0) Then
        'The old relationship was historic and so not displayed
        'Therefore we need to create a new relationship
        vAddRelationship = True
      End If
    End If

    If vContactNumber > 0 Then
      If vAddRelationship = False _
      AndAlso ((vRelationship.Length > 0 AndAlso String.Equals(vOldRelationship, vRelationship, StringComparison.CurrentCultureIgnoreCase) = False) _
         OrElse String.Equals(vOldRelationshipStatus, vRelationshipStaus, StringComparison.CurrentCultureIgnoreCase) = False) Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber2") = pList("ContactNumber")
        vList("ContactNumber") = vContactNumber
        vList("Relationship") = vRelationship
        vList("OldRelationship") = vOldRelationship
        If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
          ' vList("OldRelationshipStatus") = vOldRelationshipStatus
          vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
        End If
        DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
      End If
      Dim vParameterList As ParameterList = GetAddContactParameterList(BooleanValue(DefaultParameters("SetHistoric").ToString) = False)
      If Session("CurrentAddressNumber") IsNot Nothing AndAlso vAddressNumber = IntegerValue(Session("CurrentAddressNumber").ToString) Then
        'Same address as current contact so should have been blank
        Dim vAddressList As New ParameterList(HttpContext.Current)
        Dim vItems() As String = {"Address", "Town", "County", "Postcode", "Country", "PafStatus"}
        For Each vItem As String In vItems
          If vParameterList.ContainsKey(vItem) Then
            vAddressList(vItem) = vParameterList(vItem)
            vParameterList.Remove(vItem)
          End If
        Next
        If vAddressList.ContainsKey("Address") And vAddressList.ContainsKey("Town") Then
          vAddressList("ContactNumber") = vContactNumber
          vAddressList("Default") = "Y"
          If Not vAddressList.Contains("Country") Then
            'Country dropdown not on page so just default to UK
            vAddressList("Country") = "UK"
          End If
          Dim vAddReturnList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vAddressList)
          vAddressNumber = IntegerValue(vAddReturnList("AddressNumber").ToString)
        End If
      End If
      vParameterList("ContactNumber") = vContactNumber
      vParameterList.Add("AmendedOn", TodaysDate())
      Dim vReturnList As ParameterList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vParameterList)
      'Get comm numbers info
      SetContactCommNumbersInfo(vContactNumber, False)
      'Save communication numbers - Only process Additional Numbers when updating a contact with historic flag not set
      SaveContactCommNumbers(mvCommNumbers, vContactNumber, vAddressNumber, BooleanValue(DefaultParameters("SetHistoric").ToString) = False, False)
      ProcessChildControls(vReturnList)
    End If
    If (vAddDefaultRelationship = True OrElse vAddRelationship = True) Then
      Dim vAddContact As Boolean = True
      If vRelationship.Length > 0 OrElse vDefaultRelationship.Length > 0 Then
        Dim vReturnList As New ParameterList(HttpContext.Current)
        If vContactNumber > 0 Then
          vAddContact = False
          vReturnList("ContactNumber") = vContactNumber.ToString
        Else
          Dim vContactList As ParameterList = GetAddContactParameterList()
          If vContactList.ContainsKey("Surname") Then
            If Not vContactList.ContainsKey("Address") AndAlso Not vContactList.ContainsKey("Town") Then
              vContactList("AddressNumber") = pList("AddressNumber")
            End If
            vReturnList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vContactList)
          End If
        End If
        If vReturnList.ContainsKey("ContactNumber") Then
          Dim vList As New ParameterList(HttpContext.Current)
          Dim vLinkReturnList As ParameterList
          If vRelationship.Length > 0 AndAlso vAddRelationship = True Then
            vList("ContactNumber") = vReturnList("ContactNumber")
            vList("ContactNumber2") = pList("ContactNumber")
            vList("Relationship") = vRelationship
            If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
              vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
            End If
            vLinkReturnList = DataHelper.AddLink(vList)
            If vLinkReturnList.ContainsKey("ComplimentaryRelationship") Then
              vList("Relationship") = vLinkReturnList("ComplimentaryRelationship")
              vList("ContactNumber") = pList("ContactNumber")
              vList("ContactNumber2") = vReturnList("ContactNumber")
              If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
                vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
              End If
              DataHelper.AddLink(vList)
            End If
          End If
          If vDefaultRelationship.Length > 0 AndAlso vAddDefaultRelationship = True Then
            If vRelationship <> vDefaultRelationship Then
              vList("ContactNumber") = vReturnList("ContactNumber")
              vList("ContactNumber2") = pList("ContactNumber")
              vList("Relationship") = vDefaultRelationship
              If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
                vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
              End If
              vLinkReturnList = DataHelper.AddLink(vList)
              If vLinkReturnList.ContainsKey("ComplimentaryRelationship") Then
                vList("Relationship") = vLinkReturnList("ComplimentaryRelationship")
                vList("ContactNumber") = pList("ContactNumber")
                vList("ContactNumber2") = vReturnList("ContactNumber")
                If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
                  vList("RelationshipStatus") = vRelationshipStaus
                End If
                DataHelper.AddLink(vList)
              End If
            End If
          End If
          If vAddContact Then
            'Get comm numbers info
            SetContactCommNumbersInfo(IntegerValue(vReturnList("ContactNumber").ToString), False)
            'Save additional numbers only
            SaveContactCommNumbers(mvCommNumbers, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), True, False)
            ProcessChildControls(vReturnList)
          End If
        End If
      End If
    End If
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Public Overrides Sub ClearControls()
    ClearControls(False)
  End Sub

  Public Overrides Sub ClearControls(ByVal pClearLabels As Boolean)
    MyBase.ClearControls(pClearLabels)
    ClearHiddenControls()
  End Sub
  Private Sub ClearHiddenControls()
    SetHiddenText("OldRelationship", String.Empty)
    SetHiddenText("HiddenAddressNumber", String.Empty)
    SetHiddenText("OldRelationship", String.Empty)
    SetHiddenText("OldRelationshipStatus", String.Empty)
    SetHiddenText("HiddenSalutation", String.Empty)
    SetHiddenText("HiddenPreferredForename", String.Empty)
    SetHiddenText("HiddenOldForename", String.Empty)
    SetHiddenText("HiddenSurnamePrefix", String.Empty)
    SetHiddenText("HiddenSurname", String.Empty)
    SetHiddenText("HiddenOldSurname", String.Empty)
    SetHiddenText("HiddenSurname2", String.Empty)
    SetHiddenText("HiddenTitle", String.Empty)

  End Sub
End Class