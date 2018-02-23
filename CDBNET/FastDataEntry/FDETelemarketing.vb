Friend Class FDETelemarketing
  Inherits CareFDEControl

  Private mvProcessCancel As Boolean = False  'Used to reset the data in telemarketing_contacts table if the user decides to cancel the transaction after getting the contact

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
  End Sub

  Friend Overrides Sub SetDefaults()
    epl.FillDeferredCombos(epl)
    MyBase.SetDefaults()
    mvProcessCancel = False

    Dim vWebBrowser As WebBrowser = DirectCast(FindControl(epl, "HtmlText"), WebBrowser)
    If vWebBrowser.Document Is Nothing OrElse Not mvDefaultList.ContainsKey("Campaign") Then vWebBrowser.Navigate("about:blank") 'This is to initialise the Document object
    vWebBrowser.Document.OpenNew(False)  'This is to clear any data in a Document object
    vWebBrowser.Refresh() 'If the control is populated once then although calling OpenNew clears the data but sometime does not refresh it

    RaiseValueChangedEvent(epl, "Campaign", "") 'the value is ignored
    RaiseValueChangedEvent(epl, "Segment", "")  'the value is ignored
    RaiseEnableOtherModulesEvent(Me, False)
    With epl.PanelInfo.PanelItems
      .Item("TopicDataSheet").Mandatory = False
      .Item("Outcome").Mandatory = True
      .Item("Precis").Mandatory = Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.phone_allow_no_text)
    End With
    epl.FindPanelControl(Of DateTimePicker)("CallBackTime").ShowCheckBox = False
    epl.ResetCallTimer()
    DirectCast(FindControl(epl, "TopicDataSheet", False), TopicDataSheet).Clear()
  End Sub

  Friend Overrides Sub GetCodeRestrictions(ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList)
    MyBase.GetCodeRestrictions(pParameterName, pList)
    Select Case pParameterName
      Case "Campaign"
        pList("Telemarketing") = "Y"
      Case "Appeal"
        pList("Campaign") = epl.GetValue("Campaign")
        pList("Telemarketing") = "Y"
      Case "Segment"
        pList("Campaign") = epl.GetValue("Campaign")
        pList("Appeal") = epl.GetValue("Appeal")
        pList("Telemarketing") = "Y"
    End Select
  End Sub

  Friend Overrides Sub ButtonClicked(ByVal pParameterName As String)
    Try
      MyBase.ButtonClicked(pParameterName)
      Select Case pParameterName
        Case "GetContact"
          Dim vSelectSet As String = epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("SelectionSet")
          Dim vTable As DataTable = Nothing
          If vSelectSet.Length > 0 Then vTable = DataHelper.SelectTelemarketingContact(IntegerValue(vSelectSet))
          If vTable Is Nothing OrElse vTable.Rows.Count = 0 Then
            ShowInformationMessage(InformationMessages.ImNoContactFound)
          Else
            mvProcessCancel = True
            epl.EnableControlList("Campaign,Appeal,Segment", False)
            RaiseSelectedContactChangedEvent(Me, IntegerValue(vTable.Rows(0)("ContactNumber")))
            epl.EnableControl("AbandonCall", mvContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse AndAlso mvContactInfo.OwnershipAccessLevel >= AppValues.CommunicationsAccessLevel)
            epl.EnableControl("Dial", mvContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse AndAlso mvContactInfo.OwnershipAccessLevel >= AppValues.CommunicationsAccessLevel)
            epl.EnableControl("GetContact", False)
          End If
        Case "AbandonCall"
          DataHelper.UpdateTelemarketingContact(CareNetServices.TelemarketingUpdateType.AbandonCall, IntegerValue(epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("SelectionSet")), mvContactInfo.ContactNumber)
          mvProcessCancel = False
          epl.EnableControl("GetContact", True)
          epl.EnableControlList("AbandonCall,Dial", False)
          epl.EnableControl("Campaign", Not mvDefaultList.ContainsKey("Campaign"))
          epl.EnableControl("Appeal", Not mvDefaultList.ContainsKey("Appeal"))
          epl.EnableControl("Segment", Not mvDefaultList.ContainsKey("Segment"))
        Case "Dial"
          'We can add this handler on initialising this module but that would try to initialise the Corebridge object
          RemoveHandler PhoneApplication.PhoneInterface.NewCallStarted, AddressOf NewCallStarted  'Remove the handler in case the call was not successful
          AddHandler PhoneApplication.PhoneInterface.NewCallStarted, AddressOf NewCallStarted
          PhoneApplication.PhoneInterface.DialNumber(mvContactInfo, False)
      End Select
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Private Sub NewCallStarted()
    Try
      RemoveHandler PhoneApplication.PhoneInterface.NewCallStarted, AddressOf NewCallStarted  'Remove the handler so that another FDE application can have a unique handle
      epl.StartCallTimer(True)
      DataHelper.UpdateTelemarketingContact(CareNetServices.TelemarketingUpdateType.Dial, IntegerValue(epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("SelectionSet")), mvContactInfo.ContactNumber)
      RaiseFormClosingAllowedEvent(False) 'This will allow the form to be closed only when the user has submitted the data
      epl.EnableControlList("Dial,AbandonCall", False)
      epl.EnableControlList("Outcome,TopicGroup,TopicDataSheet,CallTime,EndCall,TotalTime,Complete,DocumentSubject,Precis,Save", True)
      RaiseEnableOtherModulesEvent(Me, True)
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Friend ReadOnly Property ProcessCancel As Boolean
    Get
      Return mvProcessCancel
    End Get
  End Property

End Class
