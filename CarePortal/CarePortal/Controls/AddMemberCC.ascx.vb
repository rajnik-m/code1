Partial Public Class AddMemberCC
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvMembershipType As String = ""
  Private mvContactNumber As Integer
  Private mvStartDate As String = ""
  Private mvAuthorisationService As IAuthorisationService

  Public Const OK As String = "OK"
  Public Const INVALID As String = "INVALID"
  Public Const ABORT As String = "ABORT"
  Public Const PENDING As String = "PENDING"
  Public Const NOTAUTHED As String = "NOTAUTHED"
  Public Const REJECTED As String = "REJECTED"

  Private mvUseTokens As Boolean

  
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overridable Sub InitControls()
    InitialiseControls(CareNetServices.WebControlTypes.wctAddMemberCC, tblDataEntry, "CreditCardNumber,CardExpiryDate", "DirectNumber,MobileNumber")
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      SupportsOnlineCCAuthorisation = True

      InitControls()
      ServiceType = GetAuthorisationType()

      If ServiceType <> AuthorisationService.None Then
        mvContactNumber = GetContactNumberFromParentGroup()
        'Hide any warning messages
        Dim vWarning1 As Control = FindControlByName(Me, "WarningMessage1")
        If vWarning1 IsNot Nothing Then vWarning1.Visible = False
        SetLabelText("PageError", "")

        If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" Then
          PopulatePayerDDL() 'Populate Payer Drop down
          PopulateOrganisationDDL() 'Populate organisation drop down
        End If

        If Request.QueryString("MT") IsNot Nothing AndAlso Request.QueryString("MT").Length > 0 Then mvMembershipType = Request.QueryString("MT")
        If Request.QueryString("SD") IsNot Nothing AndAlso Request.QueryString("SD").Length > 0 Then
          mvStartDate = Request.QueryString("SD").ToString
        ElseIf Session.Contents.Item("StartDate") IsNot Nothing AndAlso Session("StartDate").ToString.Length > 0 Then
          mvStartDate = Session("StartDate").ToString
        Else
          mvStartDate = ""
        End If

        Select Case ServiceType
          Case AuthorisationService.TnsHosted
            If InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
              CheckTNSResponse()        'Called for payment of provisional transaction items
              SetControlStateForTnsHosted()
            Else
              SetDefaults()
            End If
          Case AuthorisationService.SagePayHosted
            If InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
              If Request.QueryString("Status") Is Nothing OrElse
              (Request.QueryString("Status") = INVALID OrElse
              Request.QueryString("Status") = ABORT OrElse
              Request.QueryString("Status") = PENDING OrElse
              Request.QueryString("Status") = NOTAUTHED OrElse
              Request.QueryString("Status") = REJECTED) Then
                mvAuthorisationService = AutorisationServiceFactory.GetAuthorisationService(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type))
                'If Not IsPostBack Then
                SetControlStateForSagePay()
                If mvUseTokens Then PopulateListBox()
                If Request.QueryString("StatusDetail") IsNot Nothing AndAlso Request.QueryString("Status") IsNot Nothing Then
                  SetErrorLabel(Request.QueryString("StatusDetail"))
                End If
                'End If
                SetDefaults()
              ElseIf Request.QueryString("Status") IsNot Nothing AndAlso Request.QueryString("Status") = OK Then
                SetControlStateForSagePay()
                SetDefaults()
                CreateMembership()
                ClearCacheData()
              Else
                SetLabelText("PageError", "Credit Card Authorisation service type is Invalid. Make sure that Online Credit Card configuration is set and try again.")
              End If
            Else
              SetDefaults()
            End If
          Case Else
            SetDefaults()
        End Select
      Else
        SetLabelText("PageError", "Credit Card Authorisation service type is Invalid. Make sure that Online Credit Card configuration is set and try again.")
      End If


    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  ''' <summary>
  ''' Method to populate Payer Drop Down with the hardcoded values of "user" and "Organisation" with "User" set as default
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PopulatePayerDDL()
    Dim vPayerDDL As DropDownList = TryCast(FindControlByName(Me, "Payer"), DropDownList)
    If vPayerDDL IsNot Nothing Then
      vPayerDDL.Items.Add(New ListItem("User", "U"))
      vPayerDDL.Items.Add(New ListItem("Organisation", "O"))
      vPayerDDL.SelectedValue = "U"
    End If
  End Sub
  ''' <summary>
  ''' Method to populate all the current organisation where logged in user has current position 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PopulateOrganisationDDL()
    Dim vOrganisationDDL As DropDownList = TryCast(FindControlByName(Me, "Organisation"), DropDownList)
    If vOrganisationDDL IsNot Nothing Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = UserContactNumber()
      vList("Current") = "Y"
      Dim vContactPositionDT As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vList)
      If vContactPositionDT IsNot Nothing Then
        For Each vRows As DataRow In vContactPositionDT.Rows
          vOrganisationDDL.Items.Add(New ListItem(vRows("ContactName").ToString, (vRows("ContactNumber").ToString + "," + vRows("AddressNumber").ToString)))
        Next
      End If
    End If
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Sub SetDefaults()
    If mvMembershipType.Length = 0 Then mvMembershipType = InitialParameters("MembershipType").ToString
    SetTextBoxText("MembershipType", mvMembershipType)
    SetLookupItem(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, "MembershipType", mvMembershipType)
    'mvContactNumber
    If Not InWebPageDesigner() Then SetMemberBalance(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_cc), mvMembershipType, mvContactNumber, mvStartDate)
  End Sub

  Private Sub SetControlStateForSagePay()
    Dim vControlDictionary As New Dictionary(Of String, Boolean)
    vControlDictionary.Add("CreditCardType", False)
    vControlDictionary.Add("CreditCardNumber", False)
    vControlDictionary.Add("CardExpiryDate", False)
    vControlDictionary.Add("gatewayCardExpiryDateYear", False)
    vControlDictionary.Add("IssueNumber", False)
    vControlDictionary.Add("CardStartDate", False)
    vControlDictionary.Add("SecurityCode", False)
    vControlDictionary.Add("TokenDesc", False)


    Dim vParams As New ParameterList(HttpContext.Current)
    vParams("BatchCategory") = If(DefaultParameters.ContainsKey("BatchCategory"), DefaultParameters("BatchCategory").ToString, String.Empty)
    Dim vResult As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams))

    If vResult Is Nothing OrElse String.IsNullOrEmpty(vResult("UseTokens").ToString) OrElse String.Compare(vResult("UseTokens").ToString, "N", True) = 0 Then
      vControlDictionary.Add("TokenList", False)
      vControlDictionary.Add("CreateToken", False)
    Else
      mvUseTokens = True
    End If

    For Each vItem As KeyValuePair(Of String, Boolean) In vControlDictionary
      SetControlVisible(vItem.Key, vItem.Value)
      If FindControlByName(Me, vItem.Key) IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent.Parent IsNot Nothing Then FindControlByName(Me, vItem.Key).Parent.Parent.Visible = vItem.Value
    Next
  End Sub

  Private Sub SetControlStateForTnsHosted()
    Dim vControlDictionary As New Dictionary(Of String, Boolean)
    vControlDictionary.Add("TokenList", False)
    vControlDictionary.Add("CreateToken", False)
    vControlDictionary.Add("TokenDesc", False)

    For Each vItem As KeyValuePair(Of String, Boolean) In vControlDictionary
      SetControlVisible(vItem.Key, vItem.Value)
      If FindControlByName(Me, vItem.Key) IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent.Parent IsNot Nothing Then FindControlByName(Me, vItem.Key).Parent.Parent.Visible = vItem.Value
    Next
  End Sub


  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    CreateMembership()
  End Sub

  Private Sub CreateMembership()
    Dim vValid As Boolean
    If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED OrElse
      ServiceType = AuthorisationService.SagePayHosted Then
      vValid = True
    Else
      vValid = IsValid()
    End If

    If vValid Then
      Try
        SetErrorLabel("")

        Dim vReturnList As New ParameterList
        Dim vSkipProcessing As Boolean
        With vReturnList
          If ServiceType <> AuthorisationService.SagePayHosted OrElse (ServiceType = AuthorisationService.SagePayHosted And Request.QueryString("Status") Is Nothing) Then
            vReturnList = AddNewContact(GetHiddenContactNumber() > 0)
          ElseIf ServiceType = AuthorisationService.SagePayHosted AndAlso
            Request.QueryString("Status") IsNot Nothing AndAlso
            Cache("ContactNumber") IsNot Nothing AndAlso
            Cache("AddressNumber") IsNot Nothing Then
            vReturnList.Add("ContactNumber", Cache("ContactNumber").ToString)
            vReturnList.Add("AddressNumber", Cache("AddressNumber").ToString)
          End If
        End With

        If ServiceType = AuthorisationService.SagePayHosted AndAlso String.Compare(Request.QueryString("Status"), OK) <> 0 Then
          vSkipProcessing = Not GetSessionDetailsForSagePay(If(GetHiddenContactNumber() > 0, GetHiddenContactNumber(), Convert.ToInt32(vReturnList("ContactNumber"))))
        End If

        Dim vPayList As New ParameterList(HttpContext.Current)
        Dim vPayerDDL As DropDownList = TryCast(FindControlByName(Me, "Payer"), DropDownList)

        If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" AndAlso vPayerDDL IsNot Nothing AndAlso ParentGroup.Length = 0 Then
          If vPayerDDL.SelectedValue.ToUpper = "U" Then
            vPayList("ContactNumber") = UserContactNumber()
            vPayList("AddressNumber") = UserAddressNumber()
          Else
            vPayList("ContactNumber") = vReturnList("ContactNumber")
            vPayList("AddressNumber") = vReturnList("AddressNumber")
          End If
        Else
          vPayList("ContactNumber") = vReturnList("ContactNumber")
          vPayList("AddressNumber") = vReturnList("AddressNumber")
        End If

        Dim vList As New ParameterList(HttpContext.Current)
        AddMemberParameters(vList, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), GetPaymentMethod, mvMembershipType, mvStartDate)

        If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" AndAlso ParentGroup.Length = 0 Then
          vList("MemberContactNumber") = vReturnList("ContactNumber").ToString
          vList("MemberAddressNumber") = vReturnList("AddressNumber").ToString
        End If

        PreMemberCreation(vList)

        Dim vMemberList As New ParameterList
        vMemberList = DataHelper.AddMember(vList)

        AddGiftAidDeclaration(IntegerValue(vList("PayerContactNumber").ToString))

        'Now need to take the payment
        vPayList("PaymentPlanNumber") = vMemberList("PaymentPlanNumber")
        vPayList("Amount") = GetTextBoxText("Balance")
        vPayList("BankAccount") = DefaultParameters("BankAccount")
        vPayList("Source") = DefaultParameters("Source")
        AddUserParameters(vPayList)

        If ServiceType = AuthorisationService.SagePayHosted And Not vSkipProcessing Then
          AddCardDetailsForSagePay(vPayList)
        Else
          If Not vSkipProcessing Then AddCCParameters(vPayList)
        End If

        If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
          vPayList.Add("TnsSession", Session("SessionID"))
          If DefaultParameters.ContainsKey("BatchCategory") Then vPayList("BatchCategory") = DefaultParameters("BatchCategory")
        End If

        PrePaymentPlanPayment(vPayList)

        Try
          vPayList("DeletePaymentPlan") = "Y"

          'If the Online Authorisation flag is not set then do not go for online authorisation even if user has entered security code
          ' when the online authorisation is set for TNS Hosted
          If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso
            InitialParameters.OptionalValue("OnlineCCAuthorisation") = "N" AndAlso
            vPayList.ContainsKey("GetAuthorisation") AndAlso vPayList("GetAuthorisation").ToString = "Y" Then vPayList("GetAuthorisation") = "N"

          If Not vSkipProcessing Then DataHelper.AddPaymentPlanPayment(vPayList)
        Catch vEx As ThreadAbortException
          Throw vEx
        Catch vEx As CareException
          SetErrorLabel(vEx.Message)
          SetHiddenText("HiddenContactNumber", vReturnList("ContactNumber").ToString)
          SetHiddenText("HiddenAddressNumber", vReturnList("AddressNumber").ToString)
          vSkipProcessing = True

          'In case of error validate the card details again as users might have changed the details
          If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
            ClearSessionForTnsValues(True)
            CheckTNSResponse()
          End If
        End Try
        If vSkipProcessing = False Then
          ProcessChildControls(vReturnList)
          If SubmitItemUrl.Length > 0 Then
            Dim vSubmitParams As New StringBuilder
            With vSubmitParams
              .Append("MT=")
              .Append(mvMembershipType)
              .Append("&MN=")
              .Append(vMemberList("MemberNumber").ToString)
              .Append("&MSN=")
              .Append(vMemberList("MembershipNumber").ToString)
              If vMemberList.ContainsKey("CardExpiryDate") AndAlso vMemberList("CardExpiryDate").ToString.Length > 0 Then
                .Append("&CED=")
                .Append(vMemberList("CardExpiryDate").ToString)
              End If
              If vMemberList.ContainsKey("MemberNumber2") Then
                .Append("&MN2=")
                .Append(vMemberList("MemberNumber2").ToString)
                .Append("&MSN2=")
                .Append(vMemberList("MembershipNumber2").ToString)
                If vMemberList("CardExpiryDate2").ToString.Length > 0 Then
                  .Append("&CED2=")
                  .Append(vMemberList("CardExpiryDate2").ToString)
                End If
              End If
              .Append("SD=")
              .Append(mvStartDate)
            End With
            GoToSubmitPage(vSubmitParams.ToString)
          Else
            GoToSubmitPage()
          End If
        End If
      Catch vEX As ThreadAbortException
        Throw vEX
      Catch vException As CareException
        Select Case vException.ErrorNumber
          Case CareException.ErrorNumbers.enSoleMembership
            Dim vWarning1 As Control = FindControlByName(Me, "WarningMessage1")
            If vWarning1 IsNot Nothing Then vWarning1.Visible = True Else ProcessError(vException)
            'in case of error validate the card details again as users might have changed the details
            If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
              ClearSessionForTnsValues(True)
              CheckTNSResponse()
            End If

          Case CareException.ErrorNumbers.enSingleMembershipOnly
            SetLabelText("PageError", vException.Message)
            'in case of error validate the card details again as users might have changed the details
            If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
              ClearSessionForTnsValues(True)
              CheckTNSResponse()
            End If
          Case Else
            ProcessError(vException)
        End Select
      Catch vEx As Exception
        ProcessError(vEx)
      Finally
        If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then ClearSessionForTnsValues(True)
      End Try
    End If
  End Sub

  ''' <summary>
  ''' Get all the details card details recieved from Sage Pay
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub AddCardDetailsForSagePay(pParams As ParameterList)
    With pParams
      .Add("BatchNumber", Session("BatchNumber"))
      .Add("TransactionNumber", Session("TransactionNumber"))
      .Add("GetAuthorisation", "Y")
      .Add("NoClaimRequired", "N")
      .Add("BatchCategory", CStr(If(DefaultParameters.OptionalValue("BatchCategory").Length > 0, DefaultParameters.OptionalValue("BatchCategory"), "")))
      .Add("CardDigits", Request.QueryString("Last4Digit"))
      If Request.QueryString("Token") IsNot Nothing AndAlso Not String.IsNullOrEmpty(Request.QueryString("Token")) Then .Add("TokenId", Request.QueryString("Token"))
      If Cache("TokenDescription") IsNot Nothing Then
        .Add("TokenDesc", Cache("TokenDescription"))
        If Cache("TokenDesc") IsNot Nothing Then Cache.Remove("TokenDescription")
      End If
      .Add("VendorCode", Request.QueryString("VendorTxCode"))
      .Add("CardNumberNotRequired", "Y")
      .Add("CardExpiryDate", Request.QueryString("ExpiryDate"))
      .Add("CreditCardType", Request.QueryString("CardType"))
      .Remove("BatchNumber")
      .Remove("TransactionNumber")
    End With
  End Sub

  Protected Overridable Sub PreMemberCreation(ByVal pList As ParameterList)
    'None for this module
  End Sub

  Protected Overridable Sub PrePaymentPlanPayment(ByVal pList As ParameterList)
    'None for this module
  End Sub

  Protected Overridable Function GetPaymentMethod() As String
    Return DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_cc)
  End Function

  Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
    If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" Then
      If Session("ReturnUrl") Is Nothing Then
        If Session("ActionLink") IsNot Nothing Then Session.Remove("ActionLink")
        GetSessionDetailsForTNS()
      End If
      If Session("ReturnUrl") IsNot Nothing Then AddHiddenField(tblDataEntry, Session("ReturnUrl").ToString, "gatewayReturnURL")
      AddHiddenField(tblDataEntry, Request.Url.AbsoluteUri, "AddMemberCC")
    End If
  End Sub

  ''' <summary>
  ''' This function will check the response from the TNS and show the error message to the user (If any) and restore the field values. 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CheckTNSResponse()
    If Session("FormErrorContents") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.INVALID_FIELD_VALUES Then

      'Set all the fields with the values came back from TNS
      SetFieldValues()
      'Display Error 
      DisplayError()
      'Get session details for TNS
      GetSessionDetailsForTNS()
      'Rename the controls on the page to the name required by TNS
      RenameControlsForTNS(tblDataEntry)
      'Clear any information stored in the session
      ClearSessionForTnsValues(True)
    ElseIf Session("FormErrorContents") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.SESSION_EXPIRED Then
      'Set the field values back to what was entered when the authorisation fails
      'so that user do not have to enter all the values again
      SetFieldValues()
      'clear the session from then session details and get a new session 
      ClearSessionValue("Actionlink")
      'Get the session details from TNS
      GetSessionDetailsForTNS()
      'Add link to the return page for TNS server to
      'return with all the card validation details
      AddReturnLinkForTns(tblDataEntry)
      'Rename the controls on the page to the name required by TNS
      RenameControlsForTNS(tblDataEntry)
      'Set the error lable for TNS	
      SetErrorLabel("Session expired for TNS")
      'Get session details for TNS
      ClearSessionForTnsValues(True)
    ElseIf Session("FormErrorCode") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.SUCCESSFUL Then
      SetFieldValues()

      SetDefaults()

      CreateMembership()

    Else
      If Session("SessionURL") Is Nothing Then
        GetSessionDetailsForTNS()
        SetDefaults()
        RenameControlsForTNS(tblDataEntry)
      End If
    End If
  End Sub
  ''' <summary>
  ''' Call server to get the TNS session details and store it the session for use. All the information passed in the
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetSessionDetailsForTNS()
    'Only get new session whne the session is expired  else use the same session as card details will 
    'already be validate in the session and will not need resending
    If Session("ActionLink") Is Nothing Then
      Dim vParams As New ParameterList(HttpContext.Current)
      If DefaultParameters.ContainsKey("BatchCategory") Then vParams("BatchCategory") = DefaultParameters("BatchCategory")
      vParams("CarePortal") = "Y"
      Dim vResult As DataRow = Nothing
      Try
        vResult = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams))
      Catch ex As Exception
        SetErrorLabel(ex.Message)
        'mvSkipProcessing = True
      End Try
      If vResult IsNot Nothing AndAlso vResult("GatewayFormUrl") IsNot Nothing Then
        Dim vIndex As Integer = vResult.Item("GatewayFormUrl").ToString.Length - vResult.Item("GatewayFormUrl").ToString.LastIndexOf("/"c)
        Session("SessionID") = vResult.Item("GatewayFormUrl").ToString().Substring(vResult.Item("GatewayFormUrl").ToString.LastIndexOf("/"c) + 1, vIndex - 1)
        Session("ActionLink") = vResult.Item("GatewayFormUrl").ToString + "?charset=UTF-8"
        Me.Page.Form.Action = vResult.Item("GatewayFormUrl").ToString + "?charset=UTF-8"
        Me.Page.Form.Method = "Post"
        Session("ReturnUrl") = vResult("ReturnUrl").ToString
      End If
    Else
      Me.Page.Form.Action = Session("ActionLink").ToString
      Me.Page.Form.Method = "Post"
    End If
  End Sub

  ''' <summary>
  ''' Makes a call to the sever to open up the session details 
  ''' </summary>
  ''' <remarks></remarks>
  Private Function GetSessionDetailsForSagePay(pContactNumber As Integer) As Boolean
    Dim vReturnValue As String = ""
    Dim vParameters As New ParameterList(HttpContext.Current)
    StoreSagePaySessionDetails()

    With vParameters
      Dim vTokenList As ListBox = TryCast(Me.FindControl("TokenList"), ListBox)
      If vTokenList IsNot Nothing AndAlso vTokenList.SelectedIndex > -1 Then
        .Add("Token", vTokenList.SelectedValue.ToString)
      ElseIf Me.FindControl("CreateToken") IsNot Nothing AndAlso TryCast(Me.FindControl("CreateToken"), CheckBox).Checked Then
        .Add("CreateToken", "Y")
        Cache("TokenDescription") = GetTextBoxText("TokenDesc")
      Else
        'Token facility is not switched ON
      End If
      .Add("ContactNumber", pContactNumber)
      .Add("AddressNumber", GetContactAddress(pContactNumber))
      .Add("Amount", GetTextBoxText("Balance"))
      .Add("Product", mvMembershipType.ToString)
      If (DefaultParameters.ContainsKey("BatchCategory")) Then .Add("BatchCategory", DefaultParameters("BatchCategory").ToString)
      .Add("MakeRequest", "Y")
      .Add("Description", GetTextBoxText("MembershipType"))
    End With

    Dim vMechantDetails As ParameterList = mvAuthorisationService.CheckConnection(vParameters)
    If vMechantDetails IsNot Nothing Then
      Cache("VendorName") = vMechantDetails("MerchantId").ToString
      Response.Redirect(vMechantDetails("GatewayFormUrl").ToString)
      Return True
    Else
      SetErrorLabel(vReturnValue)
      Return False
    End If
  End Function

  Private Sub StoreSagePaySessionDetails()
    If Request.QueryString("Status") Is Nothing Then Cache("AddMemberCCUrl") = Request.Url
  End Sub

  Protected Overrides Function GetAuthorisationType() As AuthorisationService
    Dim vAuthorisationType As AuthorisationService = AuthorisationService.None
    Select Case DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type).ToUpper
      Case AuthorisationService.TnsHosted.GetServiceName()
        vAuthorisationType = AuthorisationService.TnsHosted
      Case AuthorisationService.SagePayHosted.GetServiceName()
        vAuthorisationType = AuthorisationService.SagePayHosted
      Case AuthorisationService.CommsXL.GetServiceName()
        vAuthorisationType = AuthorisationService.CommsXL
      Case AuthorisationService.SecureCXL.GetServiceName()
        vAuthorisationType = AuthorisationService.SecureCXL
      Case AuthorisationService.ProtX.GetServiceName()
        vAuthorisationType = AuthorisationService.ProtX
      Case Else
        vAuthorisationType = AuthorisationService.None
    End Select
    Return vAuthorisationType
  End Function

  Private Sub ClearCacheData()
    If Cache("ContactNumber") IsNot Nothing Then Cache.Remove("ContactNumber")
    If Cache("AddressNumber") IsNot Nothing Then Cache.Remove("AddressNumber")
    If Cache("AddMemberCCUrl") IsNot Nothing Then Cache.Remove("AddMemberCCUrl")
  End Sub

End Class