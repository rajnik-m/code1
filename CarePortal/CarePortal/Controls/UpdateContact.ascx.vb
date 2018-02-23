Public Partial Class UpdateContact
    Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub


  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateContact, tblDataEntry)
      AddHiddenField("HiddenSalutation")
      AddHiddenField("HiddenPreferredForename")
      AddHiddenField("HiddenOldForename")
      AddHiddenField("HiddenSurnamePrefix")
      AddHiddenField("HiddenSurname") 'existing surname value
      AddHiddenField("HiddenOldSurname") 'surname value before the most recent update
      AddHiddenField("HiddenSurname2") 'Used for capitalisation only
      AddHiddenField("HiddenTitle")
      AddHiddenField("HiddenSex")
      AddHiddenField("HiddenInitials")
      AddHiddenField("HiddenLabelName")
      If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.use_ajax_for_contact_names, False) Then
        AddHandlersAndTriggers(tblDataEntry)
      Else
        'On Change of Surname, HiddenSurname2 and Surname (for capitalisation) fields should be updated
        AddTextChangedHandler("Surname")
        AddAsyncPostBackTrigger("HiddenSurname2,Surname", "Surname", PostBackTriggerEventTypes.TextChanged)
      End If

      If Not IsPostBack Then
        For Each vCareWebControl As CareWebControl In Me.PageCareControls
          vCareWebControl.ClearControls(True)
        Next
        Session("CurrentContactNumber") = 0
        Session("CurrentAddressNumber") = 0
        Dim vList As New ParameterList(HttpContext.Current)
        'Session("CurrentContactNumber") = UserContactNumber().ToString
        vList("ContactNumber") = GetContactNumberFromParentGroup().ToString
        Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList)
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(vTable) 'DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(SelectData.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
        If vRow IsNot Nothing Then
          If vRow("OwnershipAccessLevel").ToString = "W" Then
            ProcessContactSelection(vTable)
          End If
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

    Public Overrides Sub ProcessSubmit()
    Try
      Dim vParams As ParameterList = GetAddContactParameterList()
      'Set Default values to get record save
      vParams("ContactNumber") = GetContactNumberFromParentGroup().ToString
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, GetContactNumberFromParentGroup))
      vParams("Status") = vRow("Status").ToString
      vParams("StatusDate") = vRow("StatusDate").ToString
      vParams("SourceDate") = vRow("SourceDate").ToString
      vParams("Address") = vRow("Address").ToString

      AddUserParameters(vParams)
      Dim vResultList As ParameterList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vParams)
      SubmitChildControls(vResultList)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try

    End Sub

End Class