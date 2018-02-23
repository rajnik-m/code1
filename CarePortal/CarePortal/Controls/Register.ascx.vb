Partial Public Class Register
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctRegister, tblDataEntry, "", "DirectNumber,MobileNumber,Position,Name")
      SetControlVisible("WarningMessage1", False)
      SetControlVisible("WarningMessage2", False)
      SetControlVisible("WarningMessage3", False)
      If HttpContext.Current.User.Identity.IsAuthenticated Then
        ShowMessageOnlyFromLabel("WarningMessage1", "You are already registered on this site and do not need to register")
      End If
      'Check Query String Parameters  EA,AN,ON for Email Address,Address Number,Organisation Number 
      CheckQueryParameters()
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vParams As New ParameterList(HttpContext.Current)
        Dim vEMailIsUserName As Boolean
        Dim vConfirmationRequired As Boolean

        If DefaultParameters.OptionalValue("ConfirmationRequired") = "Y" Then vConfirmationRequired = True
        If InitialParameters.OptionalValue("EmailAddressIsUserName") = "Y" Then vEMailIsUserName = True
        Dim vUserName As String = ""
        If vEMailIsUserName Then
          vUserName = GetTextBoxText("EMailAddress")
        Else
          vUserName = GetTextBoxText("UserName")
        End If
        vParams("UserName") = vUserName
        AddOptionalTextBoxValue(vParams, "Password")
        'Check to see if user name is in use
        If CanUserRegister(vParams) Then
          Dim vAddContactList As ParameterList = GetAddContactParameterList()
          vAddContactList("UserName") = vUserName
          If GetTextBoxText("SecurityQuestion") IsNot Nothing AndAlso GetTextBoxText("SecurityQuestion").Length <> 0 Then
            vAddContactList("SecurityQuestion") = GetTextBoxText("SecurityQuestion")
            vAddContactList("SecurityAnswer") = GetTextBoxText("SecurityAnswer")
          End If
          AddOptionalTextBoxValue(vAddContactList, "Password")
          If DefaultParameters.OptionalValue("Status").Length > 0 Then vAddContactList("Status") = DefaultParameters("Status")
          If DefaultParameters.OptionalValue("DeDuplicate") = "Y" Then vAddContactList("DeDuplicate") = "Y"

          If (Request.QueryString("AN") IsNot Nothing AndAlso Request.QueryString("AN").Length > 0) Then
            vAddContactList("AddressNumber") = Request.QueryString("AN")
            Dim vDDL As DropDownList = TryCast(FindControlByName(tblDataEntry, "OrganisationAddress"), DropDownList)
            If vDDL IsNot Nothing Then
              Dim vAddressNumber As String = GetDropDownValue("OrganisationAddress")
              If Not String.IsNullOrEmpty(vAddressNumber) Then
                If IntegerValue(vAddressNumber) > 0 Then
                  vAddContactList("AddressNumber") = vAddressNumber     'Override using the address number passed in (default address for the organisation)
                Else
                  vAddContactList.Remove("AddressNumber")               'Create a new address for the organisation
                End If
                If (Request.QueryString("ON") IsNot Nothing AndAlso Request.QueryString("ON").Length > 0) Then
                  vAddContactList("OrganisationNumber") = Request.QueryString("ON")
                End If
              End If
            End If
          End If

          'Here we need to check if we would dedup to an existing contact that is a registered user
          If vAddContactList.ContainsKey("DeDuplicate") AndAlso vAddContactList.ContainsKey("Surname") AndAlso vAddContactList.ContainsKey("EMailAddress") Then
            vAddContactList("DeDuplicate") = "C"   'Just do the check
            Dim vDeDupCheckReturnList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vAddContactList)
            vAddContactList("DeDuplicate") = "Y"   'Now do it properly
          End If

          If vConfirmationRequired Then vAddContactList("ConfirmationRequired") = "Y"
          If GetTextBoxText("ConfirmEMailAddress") IsNot Nothing AndAlso GetTextBoxText("ConfirmEMailAddress").Length <> 0 Then
            vAddContactList("ConfirmEMailAddress") = GetTextBoxText("ConfirmEMailAddress")
          End If


          Dim vRedirectToDeDupOrgPage As Boolean = False
          If vAddContactList.OptionalValue("Name").Length > 0 AndAlso DefaultParameters.OptionalValue("DeDuplicateOrganisationPage").Length > 0 AndAlso
            (Request.QueryString("AN") Is Nothing OrElse Request.QueryString("AN").Length = 0) AndAlso
            (Request.QueryString("ON") Is Nothing OrElse Request.QueryString("ON").Length = 0) Then
            '1) : Only redirect to dedup page when the Organisation Name and DeDuplicateOrganisationPage is provided and 
            '            there are no AN and ON query string parameters
            vRedirectToDeDupOrgPage = True
            '2) : Do not redirect when we have already redirected to dedup page and the user clicks Back
            '     and submits the page without changing the Name field value
            If Session("AddContactList") IsNot Nothing AndAlso DirectCast(Session("AddContactList"), ParameterList)("Name").ToString.ToLower = vAddContactList("Name").ToString.ToLower Then
              vRedirectToDeDupOrgPage = False
              Session.Remove("AddContactList")
              Session.Remove("EmailParams")
              Session.Remove("ContentParams")
            End If
            '3) : Only redirect when duplicate organisations are found
            If vRedirectToDeDupOrgPage Then
              vRedirectToDeDupOrgPage = False
              Dim vDuplicateList As New ParameterList(HttpContext.Current)
              vDuplicateList("Name") = vAddContactList("Name")
              Dim vDuplicateOrgs As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtDuplicateOrganisationsForRegistration, vDuplicateList)
              If vDuplicateOrgs IsNot Nothing AndAlso vDuplicateOrgs.Rows.Count > 0 Then
                Session("AddContactList") = vAddContactList
                vRedirectToDeDupOrgPage = True
              End If
            End If
          End If
          Dim vList As ParameterList = Nothing
          If vRedirectToDeDupOrgPage = False Then
            If vAddContactList.ContainsKey("ConfirmEMailAddress") Then vAddContactList.Remove("ConfirmEMailAddress")
            vList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vAddContactList)
          End If

          If vConfirmationRequired Then
            Dim vEmailParams As New ParameterList(HttpContext.Current)
            vEmailParams("StandardDocument") = DefaultParameters("StandardDocument")
            vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
            vEmailParams("Name") = DefaultParameters("Name")
            Dim vContentParams As New ParameterList()
            vContentParams("EMail") = vAddContactList("EMailAddress")
            Dim vRegLink As New StringBuilder
            vRegLink.Append(New UriBuilder(Request.Url.Scheme, Request.Url.Host, Request.Url.Port, Request.Url.LocalPath).Uri.AbsoluteUri)
            vRegLink.Append("?pn=")
            vRegLink.Append(DefaultParameters("ConfirmationPage").ToString)
            vRegLink.Append("&UserName=")
            Dim vEP As New EncryptionProvider
            vRegLink.Append(vEP.Encrypt(vUserName))
            vContentParams("RegistrationLink") = vRegLink.ToString
            If vRedirectToDeDupOrgPage Then
              Session("EmailParams") = vEmailParams
              Session("ContentParams") = vContentParams
            Else
              DataHelper.ProcessBulkEMail(vContentParams.ToCSVFile, vEmailParams, True)
            End If
          ElseIf vRedirectToDeDupOrgPage = False Then
            Session("ContactNumber") = vList("ContactNumber")
            Session("AddressNumber") = vList("AddressNumber")

            Dim vLoginParams As New ParameterList(HttpContext.Current)
            vLoginParams("UserName") = vParams("UserName")
            vLoginParams("Password") = vParams("Password")
            vList = DataHelper.LoginRegisteredUser(vParams)
            Session("RegisteredUserName") = vList("UserName")
            If vList.Contains("ContactNumber") Then Session("UserContactNumber") = vList("ContactNumber").ToString
            SetAuthentication(vList, True)
          End If
          If vRedirectToDeDupOrgPage Then
            'Pass the query string parameter RegURL. Do not use ReturnURL as clicking Back in dedup page will call GoToSubmit page which checks for ReturnURL
            ProcessRedirect(String.Format("default.aspx?pn={0}&RegURL={1}", DefaultParameters("DeDuplicateOrganisationPage"), Request.Url))
          Else
            GoToSubmitPage()
          End If
        End If
      Catch vCareEx As CareException
        Select Case vCareEx.ErrorNumber
          Case CareException.ErrorNumbers.enDuplicateRecord
            SetLabelTextFromLabel("Message", "WarningMessage3", "You appear to be already registered with a different user name. Please try to login")
          Case CareException.ErrorNumbers.enUserNameAlreadyInUse
            SetLabelTextFromLabel("Message", "WarningMessage2", "The User Name you entered is already in use. Please choose a different one")
          Case Else
            ProcessError(vCareEx)
        End Select
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub CheckQueryParameters()
    If (Request.QueryString("EA") IsNot Nothing AndAlso Request.QueryString("EA").Length > 0) Then
      If FindControlByName(Me, "ConfirmEMailAddress") IsNot Nothing Then
        FindControlByName(Me, "ConfirmEMailAddress").Parent.Parent.Visible = False
      End If
      If FindControlByName(Me, "EMailAddress") IsNot Nothing Then
        SetTextBoxText("EMailAddress", Request.QueryString("EA"))
        SetControlEnabled("EMailAddress", False)
      End If
    End If
    Dim vAddressNumber As Integer
    If (Request.QueryString("AN") IsNot Nothing AndAlso Request.QueryString("AN").Length > 0) Then
      vAddressNumber = IntegerValue(Request.QueryString("AN"))
      ShowAddressFields(False)
    End If
    If (Request.QueryString("ON") IsNot Nothing AndAlso Request.QueryString("ON").Length > 0) Then
      If FindControlByName(Me, "Name") IsNot Nothing Then
        Dim pList As New ParameterList(HttpContext.Current)
        pList("OrganisationNumber") = Request.QueryString("ON")
        Dim vDataTable As DataTable
        Dim vResult As String
        vResult = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftOrganisations, pList)
        vDataTable = GetDataTable(vResult)
        If Not vDataTable Is Nothing Then
          PopulateFromSessionData()
          SetTextBoxText("Name", vDataTable.Rows(0).Item("Name").ToString())
          SetControlEnabled("Name", False)

          Dim vDDL As DropDownList = TryCast(FindControlByName(tblDataEntry, "OrganisationAddress"), DropDownList)
          If vDDL IsNot Nothing Then
            pList("ContactNumber") = Request.QueryString("ON")
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, pList)
            If vTable IsNot Nothing Then
              vTable.Rows.InsertAt(vTable.NewRow(), 0)
              vTable.Rows(0).Item("AddressLine") = "<New Address>"
              vTable.Rows(0).Item("AddressNumber") = "0"
              vTable.Rows(0).Item("Historical") = ""
              vTable.DefaultView.RowFilter = "Historical  = ''"
              vDDL.DataValueField = "AddressNumber"
              vDDL.DataTextField = "AddressLine"
              vDDL.DataSource = vTable
              vDDL.DataBind()
              If vAddressNumber > 0 Then SetDropDownText("OrganisationAddress", vAddressNumber.ToString)
              vDDL.AutoPostBack = True
              AddHandler vDDL.SelectedIndexChanged, AddressOf SelectedAddressChanged
            End If
          End If
        End If
      End If
    Else
      'We don't have an organisation selected
      If Not InWebPageDesigner() Then
        SetParentParentVisible("OrganisationAddress", False)
      End If
      PopulateFromSessionData()
    End If
  End Sub

  Private Sub PopulateFromSessionData()
    If IsPostBack = False AndAlso Session("AddContactList") IsNot Nothing Then
      'The user has clicked Back button in dedup page. Set all the fields from the data saved in the Session.
      Dim vAddContactsList As ParameterList = DirectCast(Session("AddContactList"), ParameterList)
      For Each vName As String In vAddContactsList.Keys
        Select Case vName
          Case "Title", "LabelNameFormatCode", "Country", "Sex", "Status"
            SetDropDownText(vName, vAddContactsList(vName).ToString)
          Case Else
            SetTextBoxText(vName, vAddContactsList(vName).ToString)
        End Select
      Next
    End If
  End Sub

  Private Sub SelectedAddressChanged(pSender As Object, e As EventArgs)
    Dim vDDL As DropDownList = TryCast(pSender, DropDownList)
    If vDDL IsNot Nothing AndAlso vDDL.SelectedIndex = 0 Then ShowAddressFields(True)
  End Sub

  Private Sub ShowAddressFields(pVisible As Boolean)
    SetParentParentVisible("PostcoderPostcode", pVisible)
    If FindControlByName(Me, "PostcoderAddress") Is Nothing Then
      SetParentParentVisible("Address", pVisible)
      SetParentParentVisible("Town", pVisible)
      SetParentParentVisible("County", pVisible)
      SetParentParentVisible("Postcode", pVisible)
      SetParentParentVisible("Country", pVisible)
    Else
      If FindControlByName(Me, "PostcoderAddress") IsNot Nothing Then FindControlByName(Me, "PostcoderAddress").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Address") IsNot Nothing Then FindControlByName(Me, "Address").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Town") IsNot Nothing Then FindControlByName(Me, "Town").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "County") IsNot Nothing Then FindControlByName(Me, "County").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Postcode") IsNot Nothing Then FindControlByName(Me, "Postcode").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Country") IsNot Nothing Then FindControlByName(Me, "Country").Parent.Parent.Parent.Parent.Visible = pVisible
    End If
  End Sub

  Private Function CanUserRegister(pParams As ParameterList) As Boolean
    Try
      Dim vResultList As ParameterList = DataHelper.LoginRegisteredUser(pParams)
    Catch vCareEx As CareException
      Select Case vCareEx.ErrorNumber
        Case CareException.ErrorNumbers.enUserDoesNotExist
          Return True
        Case Else
          Select Case vCareEx.Message
            Case "Invalid password for user name already in use"
              Throw New CareException("User name already in use", CareException.ErrorNumbers.enUserNameAlreadyInUse)
            Case Else
              Throw vCareEx
          End Select
      End Select
    End Try
  End Function

End Class