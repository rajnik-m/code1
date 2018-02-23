Imports System.Net
Imports System.IO

Partial Public Class ProcessPayment
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvGrossTotal As String = ""
  Private mvCardNumber As String = ""
  Private mvCardIssueNumber As String = ""
  Private mvStartDate As String = ""
  Private mvExpiryDate As String = ""
  Private mvCSCValue As String = ""
  Private mvCardType As String = ""
  Private mvBatchNumber As String = ""
  Private mvTransactionNumber As String = ""
  Private mvSkipProcessing As Boolean
  Private mvIsZeroAmount As Boolean
  Private mvIsBasketEmpty As Boolean
  Private mvProductList As StringBuilder
  Private mvContactNumber As Integer
  Private mvResponseFromTNS As TNSResult = TNSResult.NONE
  Private mvAuthorisationService As IAuthorisationService
  Private mvControlDictionary As Dictionary(Of String, Boolean)
  Private mvUseTokens As Boolean
  Dim mvContactCompany As String = String.Empty
  Dim mvSalesAccount As String = String.Empty

  Public Const OK As String = "OK"
  Public Const INVALID As String = "INVALID"
  Public Const ABORT As String = "ABORT"
  Public Const PENDING As String = "PENDING"
  Public Const NOTAUTHED As String = "NOTAUTHED"
  Public Const REJECTED As String = "REJECTED"


  Public Sub New()
    mvNeedsAuthentication = True
    mvProductList = New StringBuilder
  End Sub


  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim vCardSale As Boolean

    Try

      If InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI" Then
        vCardSale = True
        InitialiseControls(CareNetServices.WebControlTypes.wctProcessPayment, tblDataEntry, "CreditCardNumber,CardExpiryDate")
      Else
        vCardSale = False
        InitialiseControls(CareNetServices.WebControlTypes.wctProcessPayment, tblDataEntry)
      End If


      Dim vList As New ParameterList(HttpContext.Current)
      mvContactNumber = If(Request.QueryString("ContactNumber") IsNot Nothing, IntegerValue(Request.QueryString("ContactNumber")), UserContactNumber)

      'Get service type 
      ServiceType = GetAuthorisationType()

      'If the authorisation service type is SagePay
      If ServiceType = AuthorisationService.SagePayHosted AndAlso (InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI") Then
        mvAuthorisationService = AutorisationServiceFactory.GetAuthorisationService(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type))
        If Request.QueryString("Status") Is Nothing OrElse
          (Request.QueryString("Status") = INVALID OrElse
          Request.QueryString("Status") = ABORT OrElse
          Request.QueryString("Status") = PENDING OrElse
          Request.QueryString("Status") = NOTAUTHED OrElse
          Request.QueryString("Status") = REJECTED) Then
          If Not IsPostBack Then SetControlState()
          If mvUseTokens Then PopulateListBox()

        ElseIf Request.QueryString("Status") IsNot Nothing AndAlso Request("Status") = OK Then
          ' Confirm payment
          SetControlState()
          Dim vResult As ParameterList = ConfirmPaymentUsingSagePay()
          If vResult IsNot Nothing AndAlso vResult.Count > 0 Then
            ClearCacheData()
            SendConfirmationEmails(vResult)
          End If
        Else
          SetErrorLabel("Unknow Response from SagePay. Please check the transaction details by login on to the Sage Pay website.")
        End If
      Else
        HideSagePayControls()
      End If

      'Called from Trader
      If Request.QueryString("Trader") IsNot Nothing AndAlso Request.QueryString("Trader") = "Y" AndAlso vCardSale Then
        If FindControlByName(Me, "GrossAmount") IsNot Nothing Then
          TryCast(Me.FindControl("GrossAmount"), TextBox).Text = Double.Parse(Request.QueryString("Amount")).ToString("#.##")
        End If
        Session("Trader") = "Y"
        HideAmountField()
        Session("BatchCategory") = Request.QueryString("BatchCategory").ToString

        If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED AndAlso vCardSale AndAlso _
          (InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI") Then _
          CheckTNSResponse()

      ElseIf Session("FormErrorCode") IsNot Nothing Then 'Redirected from Return.Aspx
        CheckTNSResponse()

      ElseIf GetShoppingBasketTransaction(mvContactNumber, vList) AndAlso Request.QueryString("ContactNumber") Is Nothing Then 'Called for payment of provisional transaction items
        GetTransactionDetails(vList)
        If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED AndAlso vCardSale AndAlso _
          (InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI") Then _
          CheckTNSResponse()

        If ServiceType = AuthorisationService.SagePayHosted AndAlso
          Request.QueryString("StatusDetail") IsNot Nothing AndAlso Request.QueryString("Status") IsNot Nothing AndAlso
          String.Compare(Request.QueryString("Status"), OK, True) <> 0 Then
          SetErrorLabel(Request.QueryString("StatusDetail"))
        End If
      Else
        If Request.QueryString("ContactNumber") Is Nothing AndAlso Request.QueryString("Amount") Is Nothing Then
          SetErrorLabel("Shopping Basket is Empty")
          ClearAmountFields()
          HideControls()
          mvIsBasketEmpty = True
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    MakePayment()
  End Sub

  Private Sub GetCreditCustomerDetails(ByRef pContactCompany As String, ByRef pSalesAccount As String)
    Dim vParamList As ParameterList = New ParameterList(HttpContext.Current)

    AddPayerInfo(vParamList, False)
    vParamList("Company") = InitialParameters.OptionalValue("Company").Trim

    'Check if the contact is a credit sales customer if not then create "CreateCreditCustomer" check box is checked 
    'else display an error
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vParamList)
    If vDataTable Is Nothing Then
      If DefaultParameters.OptionalValue("CreateCreditCustomer") = "Y" Then
        Dim vCreditCustParams As New ParameterList(HttpContext.Current)
        AddPayerInfo(vCreditCustParams)
        vCreditCustParams("Company") = InitialParameters.OptionalValue("Company")
        vCreditCustParams("CreditCategory") = DefaultParameters.OptionalValue("CreditCategory")
        Try
          Dim vResult As ParameterList = DataHelper.AddCreditCustomer(vCreditCustParams)
          pSalesAccount = vResult("SalesLedgerAccount").ToString
          pContactCompany = InitialParameters.OptionalValue("Company")
        Catch vEx As ThreadAbortException
          Throw vEx
        Catch vEx As CareException
          SetErrorLabel(vEx.Message)
          mvSkipProcessing = True
        End Try
      Else
        SetErrorLabel("Contact is not a Credit Sales Customer")
        mvSkipProcessing = True
      End If
    Else
      Dim vRow As DataRow() = vDataTable.Select()
      pContactCompany = vRow(0)("Company").ToString
      pSalesAccount = vRow(0)("SalesLedgerAccount").ToString
    End If

  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Function GetDetailsforProtX() As ParameterList
    Dim vParameterList As New ParameterList(HttpContext.Current)
    vParameterList("Amount") = mvGrossTotal
    vParameterList("CreditCardNumber") = mvCardNumber
    vParameterList("CardStartDate") = mvStartDate
    vParameterList("CardExpiryDate") = mvExpiryDate
    vParameterList("IssueNumber") = mvCardIssueNumber
    vParameterList("AuthorisationCode") = mvCSCValue
    vParameterList("CreditCardType") = mvCardType
    vParameterList("BatchNumber") = mvBatchNumber
    vParameterList("TransactionNumber") = mvTransactionNumber
    AddPayerInfo(vParameterList)
    vParameterList("CardNumberNotRequired") = "Y"
    vParameterList("GetAuthorisation") = "Y"
    vParameterList("IsWebTransaction") = "Y"
    vParameterList("NoClaimRequired") = "N"
    AddOptionalTextBoxValue(vParameterList, "Reference")
    AddOptionalTextBoxValue(vParameterList, "Notes")
    If DefaultParameters.ContainsKey("BatchCategory") Then vParameterList("BatchCategory") = DefaultParameters("BatchCategory")

    ''Test URL https://test.sagepay.com/gateway/service/vspdirect-register.vsp
    ''Live URL https://live.sagepay.com/gateway/service/vspdirect-register.vsp
    Return vParameterList
  End Function


  Private Function GetDetailsforCommsXL() As ParameterList
    Dim vCommsXLDetails As New ParameterList(HttpContext.Current)
    vCommsXLDetails("BatchNumber") = mvBatchNumber
    vCommsXLDetails("TransactionNumber") = mvTransactionNumber
    AddOptionalTextBoxValue(vCommsXLDetails, "CreditCardNumber")
    AddOptionalTextBoxValue(vCommsXLDetails, "CardExpiryDate")
    AddOptionalTextBoxValue(vCommsXLDetails, "IssueNumber")
    AddOptionalTextBoxValue(vCommsXLDetails, "CardStartDate")
    AddOptionalTextBoxValue(vCommsXLDetails, "SecurityCode")
    AddOptionalTextBoxValue(vCommsXLDetails, "Reference")
    AddOptionalTextBoxValue(vCommsXLDetails, "Notes")
    vCommsXLDetails("GetAuthorisation") = "Y"
    AddPayerInfo(vCommsXLDetails)
    vCommsXLDetails("CreditCardType") = mvCardType.ToString
    If DefaultParameters.ContainsKey("BatchCategory") Then vCommsXLDetails("BatchCategory") = DefaultParameters("BatchCategory")
    Return vCommsXLDetails
  End Function

  Private Sub GetSessionDetailsForTNS()
    'Only get new session whne the session is expired  else use the same session as card details will 
    'already be validate in the session and will not need resending
    If Session("ActionLink") Is Nothing Then
      Dim vParams As New ParameterList(HttpContext.Current)
      If Session("BatchCategory") IsNot Nothing AndAlso Session("BatchCategory").ToString.Length > 0 Then
        vParams("BatchCategory") = Session("BatchCategory")
      Else
        If DefaultParameters.ContainsKey("BatchCategory") Then vParams("BatchCategory") = DefaultParameters("BatchCategory")
      End If

      vParams("CarePortal") = "Y"
      Dim vResult As DataRow = Nothing
      Try
        vResult = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams))
      Catch ex As Exception
        SetErrorLabel(ex.Message)
        mvSkipProcessing = True
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
  ''' Add all details required for TNS transaction to the collection so that all can be
  ''' pass to the sever for transaction
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetDetailsForTNS() As ParameterList
    Dim vTNSDetails As New ParameterList(HttpContext.Current)
    Dim vExpiryDate As String = String.Empty
    'Add all the necessary values to the collection
    vTNSDetails("CreditCardType") = mvCardType.ToString
    AddOptionalTextBoxValue(vTNSDetails, "CreditCardNumber")
    AddOptionalTextBoxValue(vTNSDetails, "IssueNumber")
    AddOptionalTextBoxValue(vTNSDetails, "CardStartDate")
    AddOptionalTextBoxValue(vTNSDetails, "SecurityCode")
    AddOptionalTextBoxValue(vTNSDetails, "Reference")
    AddOptionalTextBoxValue(vTNSDetails, "Notes")
    vTNSDetails.Add("TnsSession", Session("SessionID"))

    If FindControlByName(Me, "CardExpiryDate") IsNot Nothing AndAlso FindControlByName(Me, "gatewayCardExpiryDateYear") IsNot Nothing Then
      Try
        vExpiryDate = TryCast(Me.FindControl("CardExpiryDate"), TextBox).Text + TryCast(Me.FindControl("gatewayCardExpiryDateYear"), TextBox).Text.Substring(2, 2)
      Catch ex As Exception
        vExpiryDate = ""
      End Try
    End If
    vTNSDetails.Add("CardExpiryDate", vExpiryDate)

    'if not called from trader application then only add the following else 
    'not required and may cause issue
    Dim vTrader As Boolean = If(Session("Trader") IsNot Nothing AndAlso Session("Trader").ToString = "Y", True, False)

    If Not vTrader Then
      vTNSDetails.Add("BatchNumber", Session("BatchNumber"))
      vTNSDetails.Add("TransactionNumber", Session("TransactionNumber"))
      vTNSDetails("GetAuthorisation") = "Y"
      vTNSDetails("NoClaimRequired") = "N"
      vTNSDetails("BatchCategory") = CStr(If(DefaultParameters.OptionalValue("BatchCategory").Length > 0, DefaultParameters.OptionalValue("BatchCategory"), ""))
      AddPayerInfo(vTNSDetails)
    End If
    Return vTNSDetails
  End Function


  Private Function MakePaymentUsingCreditCard(ByVal pCreditCardDetails As ParameterList) As ParameterList
    Dim vResult As ParameterList = Nothing
    If ((pCreditCardDetails.ContainsKey("BatchNumber") AndAlso CInt(pCreditCardDetails("BatchNumber")) > 0) AndAlso _
      (pCreditCardDetails.ContainsKey("TransactionNumber") AndAlso CInt(pCreditCardDetails("TransactionNumber")) > 0)) Then
      Try
        vResult = DataHelper.ConfirmCardSaleTransaction(pCreditCardDetails)
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vEx As CareException
        SetErrorLabel(vEx.Message)

        'If this page is using TNS then we have to get the session information and start all the authorisation process again
        If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED AndAlso _
         InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI" Then _
         CheckTNSResponse()
      End Try
    Else
      SetErrorLabel("Invalid Transaction Number")
    End If
    Return vResult
  End Function

  ''' <summary>
  ''' Confirms the transaction once the Credit Card is authorised using Sage Pay
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ConfirmPaymentUsingSagePay() As ParameterList
    Dim vTransactionDetails As New ParameterList(HttpContext.Current)
    With vTransactionDetails
      .Add("BatchNumber", Session("BatchNumber"))
      .Add("TransactionNumber", Session("TransactionNumber"))
      .Add("GetAuthorisation", "Y")
      .Add("NoClaimRequired", "N")
      .Add("BatchCategory", CStr(If(DefaultParameters.OptionalValue("BatchCategory").Length > 0, DefaultParameters.OptionalValue("BatchCategory"), "")))
      .Add("CardDigits", Request.QueryString("Last4Digit"))
      If Request.QueryString("Token") IsNot Nothing AndAlso Not String.IsNullOrEmpty(Request.QueryString("Token")) Then .Add("TokenId", Request.QueryString("Token"))
      If Cache("TokenDescription") IsNot Nothing Then
        .Add("TokenDesc", Cache("TokenDescription"))
        Cache.Remove("TokenDescription")
      End If
      .Add("VendorCode", Request.QueryString("VendorTxCode"))
      .Add("CardNumberNotRequired", "Y")
      .Add("CardExpiryDate", Request.QueryString("ExpiryDate"))
      AddPayerInfo(vTransactionDetails)
    End With

    If InitialParameters.OptionalValue("PaymentType") = "CI" Then
      If Not vTransactionDetails.ContainsKey("Company") Then vTransactionDetails("Company") = mvContactCompany
      If Not vTransactionDetails.ContainsKey("SalesLedgerAccount") Then vTransactionDetails("SalesLedgerAccount") = mvSalesAccount
      Return ConfirmCreditSalesTransaction(vTransactionDetails)
    Else
      Return MakePaymentUsingCreditCard(vTransactionDetails)
    End If

  End Function

  Private Sub SendConfirmationMailToPayer(ByVal pResultList As ParameterList)
    Dim vEmailAddress As String = String.Empty
    Dim vSalutation As String = String.Empty
    Dim vLabelName As String = String.Empty

    Dim vParam As New ParameterList(HttpContext.Current)
    AddPayerInfo(vParam, False)
    vEmailAddress = GetEmailAddress(IntegerValue(vParam("ContactNumber").ToString))

    Dim vContactDetails As String = DataHelper.SelectContactData(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vParam)
    Dim vContactDetailsTable As DataTable = GetDataTable(vContactDetails)
    If vContactDetailsTable IsNot Nothing Then
      For Each vDataRow As DataRow In vContactDetailsTable.Rows
        vSalutation = vDataRow("Salutation").ToString
        vLabelName = vDataRow("LabelName").ToString
      Next
    End If
    Dim vContentParams As New ParameterList
    vContentParams("GrossAmount") = DoubleValue(pResultList("Amount").ToString).ToString("0.00")
    vContentParams("TransactionReference") = pResultList("BatchNumber").ToString + "/" + pResultList("TransactionNumber").ToString
    vContentParams("Salutation") = vSalutation
    vContentParams("LabelName") = vLabelName
    vContentParams("ProductList") = mvProductList.ToString
    vContentParams("NetAmount") = DoubleValue(NetAmount.ToString).ToString("0.00")
    vContentParams("VatAmount") = DoubleValue(VatAmount.ToString).ToString("0.00")
    vContentParams("EMail") = vEmailAddress

    'Default Parameters Set from WPD
    Dim vEmailParams As New ParameterList(HttpContext.Current)
    vEmailParams("StandardDocument") = InitialParameters.OptionalValue("PayerDocument")
    vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
    vEmailParams("Name") = DefaultParameters("Name")
    DataHelper.ProcessBulkEMail(vContentParams.ToCSVFile, vEmailParams, True)
  End Sub

  Private Function GetEmailAddress(ByVal pContactNumber As Integer) As String
    Dim vParams As New ParameterList(HttpContext.Current)
    vParams("ContactNumber") = pContactNumber
    Dim vResult As String = DataHelper.SelectContactData(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactEMailAddresses, vParams)
    Dim vDataTable As DataTable = GetDataTable(vResult)
    Dim vEmailAddress As String = String.Empty
    If Not vDataTable Is Nothing Then
      Dim vRows() As DataRow
      vRows = vDataTable.Select("PreferredMethod LIKE 'Y%'")
      If vRows.Length = 1 Then
        vEmailAddress = vRows(0).Item("EMailAddress").ToString
      Else
        vRows = vDataTable.Select("DeviceDefault LIKE 'Y%'")
        If vRows.Length > 0 Then
          vEmailAddress = vRows(0).Item("EMailAddress").ToString
        Else
          vEmailAddress = vDataTable.Rows(0).Item("EMailAddress").ToString
        End If
      End If
    End If
    Return vEmailAddress
  End Function

  Private Sub SendConfirmationMailsToEventBookers(ByVal pParams As ParameterList)
    Try
      Dim vSendToBooker As Boolean = InitialParameters.OptionalValue("BookingDocument").Length > 0
      Dim vSendToDelegate As Boolean = InitialParameters.OptionalValue("DelegateDocument").Length > 0

      If vSendToBooker OrElse vSendToDelegate Then
        If Not pParams.ContainsKey("Database") Then pParams.AddConectionData(HttpContext.Current)
        Dim vResult As String = DataHelper.SelectContactData(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, pParams)
        Dim vDataTable As DataTable = GetDataTable(vResult)

        If vDataTable IsNot Nothing Then
          For Each vDataRow As DataRow In vDataTable.Rows
            Dim vBookerContactNumber As Integer = IntegerValue(vDataRow("ContactNumber").ToString)
            'Get event session info
            Dim vBookingSessionInfo As String = GetEventSessionInformation(vDataRow("BookingNumber").ToString)
            'Get the end date
            Dim vEventEndDate As String = String.Empty
            Dim vEventInfo As New ParameterList(HttpContext.Current)
            vEventInfo("EventNumber") = vDataRow("EventNumber")
            Dim vEventResult As String = DataHelper.SelectEventData(CarePortal.CareNetServices.XMLEventDataSelectionTypes.xedtEventInformation, vEventInfo)
            Dim vEventTable As DataTable = GetDataTable(vEventResult)
            If vEventTable IsNot Nothing Then
              vEventEndDate = vEventTable.Rows(0)("EndDate").ToString
            End If

            'Process the delegates if required
            Dim vDelegateList As StringBuilder = New StringBuilder
            Dim vParams As New ParameterList(HttpContext.Current)
            vParams("ContactNumber") = vBookerContactNumber
            vParams("EventNumber") = vDataRow("EventNumber")
            vParams("BookingNumber") = vDataRow("BookingNumber")
            'Get the booking delegates
            Dim vDelegateResult As String = DataHelper.SelectEventData(CarePortal.CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates, vParams)
            Dim vDelegateDataTable As DataTable = GetDataTable(vDelegateResult)
            If vDelegateDataTable IsNot Nothing Then
              Dim vDelegatesInfoCollection As New List(Of ParameterList)
              For Each vDelegateDataRow As DataRow In vDelegateDataTable.Rows
                'Get the information for each delegate
                Dim vDelegateMainInfo As New ParameterList(HttpContext.Current)
                vDelegateMainInfo("ContactNumber") = vDelegateDataRow("ContactNumber")
                Dim vDelegateInfo As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vDelegateMainInfo)
                Dim vDelegateInfoTable As DataTable = GetDataTable(vDelegateInfo)
                Dim vDelegateSalutation As String = String.Empty
                Dim vDelegateName As String = String.Empty
                If vDelegateInfoTable IsNot Nothing Then
                  vDelegateName = vDelegateInfoTable.Rows(0)("LabelName").ToString
                  vDelegateSalutation = vDelegateInfoTable.Rows(0)("Salutation").ToString
                End If
                vDelegateList.Append(vDelegateName)
                If vDelegateList.Length > 0 Then vDelegateList.Append("<BR>")

                If vSendToDelegate Then
                  Dim vDelegateEmail As String = GetEmailAddress(IntegerValue(vDelegateDataRow("ContactNumber").ToString))
                  If Not String.IsNullOrWhiteSpace(vDelegateEmail) Then
                    vDelegateMainInfo("Salutation") = vDelegateSalutation
                    vDelegateMainInfo("LabelName") = vDelegateName
                    vDelegateMainInfo("EMail") = vDelegateEmail
                    vDelegateMainInfo("EventStartDate") = vDataRow("StartDate").ToString
                    vDelegateMainInfo("EventEndDate") = vEventEndDate
                    vDelegateMainInfo("EventStartTime") = vDataRow("StartTime").ToString
                    vDelegateMainInfo("EventEndTime") = vDataRow("EndTime").ToString
                    vDelegateMainInfo("SessionList") = vBookingSessionInfo
                    vDelegateMainInfo("EventName") = vDataRow("EventDesc").ToString
                    vDelegateMainInfo("EventReference") = vDataRow("EventReference").ToString
                    vDelegatesInfoCollection.Add(vDelegateMainInfo)
                  End If
                End If
              Next
              If vSendToDelegate AndAlso vDelegatesInfoCollection.Count > 0 Then
                Dim vEmailParams As New ParameterList(HttpContext.Current)
                vEmailParams("StandardDocument") = InitialParameters.OptionalValue("DelegateDocument")
                vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
                vEmailParams("Name") = DefaultParameters("Name")
                DataHelper.ProcessBulkEMail(GetCSVFile(vDelegatesInfoCollection), vEmailParams, True)
              End If
            End If

            If vSendToBooker Then
              Dim vBookerEmailAddress As String = GetEmailAddress(vBookerContactNumber)
              If Not String.IsNullOrWhiteSpace(vBookerEmailAddress) Then
                Dim vSenderInfo As New ParameterList(HttpContext.Current)
                vSenderInfo("ContactNumber") = vBookerContactNumber
                Dim vBookerName As String = String.Empty
                Dim vSalutation As String = String.Empty
                Dim vLabelName As String = String.Empty
                Dim vBookerInfo As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vSenderInfo)
                Dim vBookerInfoTable As DataTable = GetDataTable(vBookerInfo)
                If vBookerInfoTable IsNot Nothing Then
                  vBookerName = vBookerInfoTable.Rows(0)("LabelName").ToString
                  vSalutation = vBookerInfoTable.Rows(0)("Salutation").ToString
                  vLabelName = vBookerInfoTable.Rows(0)("LabelName").ToString
                End If
                vSenderInfo("EMailAddress") = DefaultParameters("EMailAddress") '
                vSenderInfo("Name") = DefaultParameters("Name")
                vSenderInfo("StandardDocument") = InitialParameters.OptionalValue("BookingDocument")
                Dim vContentInformation As New ParameterList(HttpContext.Current)
                vContentInformation("EMail") = vBookerEmailAddress
                vContentInformation("Salutation") = vSalutation
                vContentInformation("LabelName") = vLabelName
                vContentInformation("EventStartDate") = vDataRow("StartDate").ToString
                vContentInformation("EventEndDate") = vEventEndDate
                vContentInformation("EventStartTime") = vDataRow("StartTime").ToString
                vContentInformation("EventEndTime") = vDataRow("EndTime").ToString
                vContentInformation("SessionList") = vBookingSessionInfo
                vContentInformation("DelegateList") = vDelegateList
                vContentInformation("EventName") = vDataRow("EventDesc").ToString
                vContentInformation("EventReference") = vDataRow("EventReference").ToString
                vContentInformation("TransactionReference") = pParams("BatchNumber").ToString + "/" + pParams("TransactionNumber").ToString
                Dim vContentInfo As New ParameterList(HttpContext.Current)
                DataHelper.ProcessBulkEMail(vContentInformation.ToCSVFile, vSenderInfo, True)
              End If
            End If
          Next
        End If
      End If
    Catch vEx As CareException
      ProcessError(vEx)
    End Try
  End Sub

  Private Function GetCSVFile(ByVal pDelegateContent As List(Of ParameterList)) As String
    Dim vFileName As String = My.Computer.FileSystem.GetTempFileName()
    Dim vStreamWriter As StreamWriter = New StreamWriter(vFileName, False)
    Dim vAddSeparator As Boolean
    For Each vParameterList As ParameterList In pDelegateContent
      For Each vItem As DictionaryEntry In vParameterList
        If vAddSeparator Then vStreamWriter.Write(",")
        vStreamWriter.Write(vItem.Key)
        vAddSeparator = True
      Next
      vAddSeparator = False
      vStreamWriter.WriteLine()
      For Each vItem As DictionaryEntry In vParameterList
        If vAddSeparator Then vStreamWriter.Write(",")
        vStreamWriter.Write("""")
        vStreamWriter.Write(vItem.Value)
        vStreamWriter.Write("""")
        vAddSeparator = True
      Next
    Next
    vStreamWriter.WriteLine()
    vStreamWriter.Close()
    Return vFileName
  End Function

  Private Function GetEventSessionInformation(ByVal pEventBookingNumber As String) As String
    Dim vSessionInfo As New StringBuilder
    Dim pParams As New ParameterList(HttpContext.Current)
    pParams("BookingNumber") = pEventBookingNumber
    Dim vResult As String = DataHelper.SelectEventData(CarePortal.CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingSessions, pParams)
    Dim vDataTable As DataTable = GetDataTable(vResult)

    If vDataTable IsNot Nothing Then
      For Each vDataRow As DataRow In vDataTable.Rows
        vSessionInfo.Append(vDataRow("SessionDesc"))
        vSessionInfo.Append("<BR>")
      Next
    End If
    Return vSessionInfo.ToString
  End Function

  Private Sub GetContactInformation(ByVal pContactNumber As String, ByVal pList As ParameterList)
    Dim pParams As New ParameterList(HttpContext.Current)
    pParams("ContactNumber") = pContactNumber
    Dim vResult As String = DataHelper.SelectContactData(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, pParams)
    Dim vDataTable As DataTable = GetDataTable(vResult)
  End Sub

  Private Sub ClearAmountFields()
    VatAmount = Nothing
    NetAmount = Nothing
    SetTextBoxText("GrossAmount", "")
    SetTextBoxText("VatAmount", "")
    SetTextBoxText("NetAmount", "")
  End Sub

  Function FindField(ByVal fieldName As String, ByVal postResponse As String) As String
    Dim vItems() As String
    Dim vIndex As Integer
    Dim vItem As String
    Dim vReponse As String = String.Empty

    vItems = Split(postResponse, Chr(13))
    For vIndex = LBound(vItems) To UBound(vItems)
      vItem = Replace(vItems(vIndex), Chr(10), "")
      If InStr(vItem, fieldName & "=") = 1 Then
        ' found
        vReponse = Right(vItem, Len(vItem) - Len(fieldName) - 1)
        Exit For
      End If
    Next
    Return vReponse
  End Function
  Private Sub HideControls()
    If FindControlByName(Me, "Title") IsNot Nothing Then
      FindControlByName(Me, "Title").Parent.Parent.Visible = False
      FindControlByName(Me, "Title").Visible = False
    End If

    If FindControlByName(Me, "Forenames") IsNot Nothing Then
      FindControlByName(Me, "Forenames").Parent.Parent.Visible = False
      FindControlByName(Me, "Forenames").Visible = False
    End If

    If FindControlByName(Me, "Surname") IsNot Nothing Then
      FindControlByName(Me, "Surname").Parent.Parent.Visible = False
      FindControlByName(Me, "Surname").Visible = False
    End If

    If FindControlByName(Me, "EMailAddress") IsNot Nothing Then
      FindControlByName(Me, "EMailAddress").Parent.Parent.Visible = False
      FindControlByName(Me, "EMailAddress").Visible = False
    End If

    If FindControlByName(Me, "ConfirmEMailAddress") IsNot Nothing Then
      FindControlByName(Me, "ConfirmEMailAddress").Parent.Parent.Visible = False
      FindControlByName(Me, "ConfirmEMailAddress").Visible = False
    End If

    If FindControlByName(Me, "PostcoderPostcode") IsNot Nothing Then
      FindControlByName(Me, "PostcoderPostcode").Parent.Parent.Visible = False
      FindControlByName(Me, "PostcoderPostcode").Visible = False
    End If

    If FindControlByName(Me, "PostcoderAddress") IsNot Nothing Then
      FindControlByName(Me, "PostcoderAddress").Parent.Parent.Parent.Parent.Visible = False
      FindControlByName(Me, "PostcoderAddress").Visible = False
    End If

    If FindControlByName(Me, "Address") IsNot Nothing Then
      FindControlByName(Me, "Address").Visible = False
      FindControlByName(Me, "Address").Parent.Parent.Parent.Parent.Visible = False
    End If

    If FindControlByName(Me, "Town") IsNot Nothing Then
      FindControlByName(Me, "Town").Parent.Parent.Parent.Parent.Visible = False
      FindControlByName(Me, "Town").Visible = False
    End If

    If FindControlByName(Me, "County") IsNot Nothing Then
      FindControlByName(Me, "County").Parent.Parent.Parent.Parent.Visible = False
      FindControlByName(Me, "County").Visible = False
    End If

    If FindControlByName(Me, "Country") IsNot Nothing Then
      FindControlByName(Me, "Country").Parent.Parent.Parent.Parent.Visible = False
      FindControlByName(Me, "Country").Visible = False
    End If

    If FindControlByName(Me, "CreditCardType") IsNot Nothing Then
      FindControlByName(Me, "CreditCardType").Parent.Parent.Visible = False
      FindControlByName(Me, "CreditCardType").Visible = False
    End If

    If FindControlByName(Me, "CreditCardNumber") IsNot Nothing Then
      FindControlByName(Me, "CreditCardNumber").Parent.Parent.Visible = False
      FindControlByName(Me, "CreditCardNumber").Visible = False
    End If

    If FindControlByName(Me, "CardExpiryDate") IsNot Nothing Then
      FindControlByName(Me, "CardExpiryDate").Parent.Parent.Visible = False
      FindControlByName(Me, "CardExpiryDate").Visible = False
    End If

    If FindControlByName(Me, "IssueNumber") IsNot Nothing Then
      FindControlByName(Me, "IssueNumber").Parent.Parent.Visible = False
      FindControlByName(Me, "IssueNumber").Visible = False
    End If

    If FindControlByName(Me, "CardStartDate") IsNot Nothing Then
      FindControlByName(Me, "CardStartDate").Parent.Parent.Visible = False
      FindControlByName(Me, "CardStartDate").Visible = False
    End If

    If FindControlByName(Me, "SecurityCode") IsNot Nothing Then
      FindControlByName(Me, "SecurityCode").Parent.Parent.Visible = False
      FindControlByName(Me, "SecurityCode").Visible = False
    End If

    If FindControlByName(Me, "gatewayCardExpiryDateYear") IsNot Nothing Then
      FindControlByName(Me, "gatewayCardExpiryDateYear").Parent.Parent.Visible = False
      FindControlByName(Me, "gatewayCardExpiryDateYear").Visible = False
    End If

  End Sub

  ''' <summary>
  ''' Following method wil hide the amount fields (VatAmount, NetAmount) which are not required
  ''' this page is shown in the trader application
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub HideAmountField()
    If FindControlByName(Me, "VatAmount") IsNot Nothing Then
      FindControlByName(Me, "VatAmount").Parent.Parent.Visible = False
      FindControlByName(Me, "VatAmount").Visible = False
    End If

    If FindControlByName(Me, "NetAmount") IsNot Nothing Then
      FindControlByName(Me, "NetAmount").Parent.Parent.Visible = False
      FindControlByName(Me, "NetAmount").Visible = False
    End If
  End Sub

  Private Sub AddPayerInfo(ByVal pList As ParameterList, Optional ByVal pAddAddressNumber As Boolean = True)
    If ParentGroup.Length > 0 Then

      Dim vContactNumber As Integer = GetContactNumberFromParentGroup()
      pList("ContactNumber") = vContactNumber
      pList("AddressNumber") = GetContactAddress(vContactNumber)
    Else
      If Session("PayerContactNumber") IsNot Nothing AndAlso Session("PayerAddressNumber") IsNot Nothing Then
        pList("ContactNumber") = Session("PayerContactNumber")
        If pAddAddressNumber Then pList("AddressNumber") = Session("PayerAddressNumber")
      Else
        pList("ContactNumber") = mvContactNumber 'UserContactNumber()
        If pAddAddressNumber Then pList("AddressNumber") = UserAddressNumber()
      End If
    End If
  End Sub

  ''' <summary>
  ''' This method makes the payment using one of the following supported payment method
  ''' Depending on the configuration
  ''' COMMSXL, SECURECXL,PROTX,TNSPAY 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub MakePayment()
    Dim vValid As Boolean
    If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED Then
      vValid = IsValid()
    Else
      vValid = True
    End If

    If vValid And Not mvIsBasketEmpty AndAlso Not mvSkipProcessing Then
      Dim vResult As ParameterList = Nothing
      Try

        If mvIsZeroAmount Then
          Dim vCardSalesParam As New ParameterList(HttpContext.Current)
          vCardSalesParam("BatchNumber") = mvBatchNumber
          vCardSalesParam("TransactionNumber") = mvTransactionNumber
          AddPayerInfo(vCardSalesParam)
          AddOptionalTextBoxValue(vCardSalesParam, "Reference")
          AddOptionalTextBoxValue(vCardSalesParam, "Notes")
          If DefaultParameters.ContainsKey("BatchCategory") Then vCardSalesParam("BatchCategory") = DefaultParameters("BatchCategory")
          Try
            vResult = DataHelper.ConfirmCashSaleTransaction(vCardSalesParam)
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vEx As CareException
            SetErrorLabel(vEx.Message)
          End Try
        ElseIf InitialParameters.OptionalValue("PaymentType") = "CC" Then
          If Me.FindControl("GrossAmount") IsNot Nothing Then mvGrossTotal = TryCast(Me.FindControl("GrossAmount"), TextBox).Text
          If Me.FindControl("CreditCardNumber") IsNot Nothing Then mvCardNumber = TryCast(Me.FindControl("CreditCardNumber"), TextBox).Text
          If Me.FindControl("IssueNumber") IsNot Nothing Then mvCardIssueNumber = TryCast(Me.FindControl("IssueNumber"), TextBox).Text
          If Me.FindControl("CardStartDate") IsNot Nothing Then mvStartDate = TryCast(Me.FindControl("CardStartDate"), TextBox).Text
          If Me.FindControl("CardExpiryDate") IsNot Nothing Then mvExpiryDate = TryCast(Me.FindControl("CardExpiryDate"), TextBox).Text
          If Me.FindControl("SecurityCode") IsNot Nothing Then mvCSCValue = TryCast(Me.FindControl("SecurityCode"), TextBox).Text
          If Me.FindControl("CreditCardType") IsNot Nothing Then mvCardType = TryCast(Me.FindControl("CreditCardType"), DropDownList).SelectedValue

          Dim vCardDetails As ParameterList = Nothing
          Select Case InitialParameters.OptionalValue("PaymentService").Trim.ToUpper
            Case "PAYPAL"
              'vResult = MakePaymentUsingPayPal()
            Case "PROTX"
              vResult = MakePaymentUsingCreditCard(GetDetailsforProtX())
            Case "COMMSXL"
              vResult = MakePaymentUsingCreditCard(GetDetailsforCommsXL())
            Case AuthorisationService.TnsHosted.GetServiceName()
              vResult = MakePaymentUsingCreditCard(GetDetailsForTNS())
            Case AuthorisationService.SagePayHosted.GetServiceName()
              vResult = GetDetailsForSagePay()
            Case Else
              SetErrorLabel("Payment Option not supported. Value for payment_authorisation_services configuration option is not set")
          End Select
        ElseIf InitialParameters.OptionalValue("PaymentType").Trim = "CS" Then
          GetCreditCustomerDetails(mvContactCompany, mvSalesAccount)
          If mvSkipProcessing = False Then
            'BatchNumber,TransactionNumber,ContactNumber,AddressNumber,Company,SalesLedgerAccount
            Dim vCreditSalesParam As New ParameterList(HttpContext.Current)
            vCreditSalesParam("BatchNumber") = mvBatchNumber
            vCreditSalesParam("TransactionNumber") = mvTransactionNumber
            AddPayerInfo(vCreditSalesParam)
            vCreditSalesParam("Company") = mvContactCompany
            vCreditSalesParam("SalesLedgerAccount") = mvSalesAccount
            vCreditSalesParam("CreateInvoice") = "Y"
            AddOptionalTextBoxValue(vCreditSalesParam, "Reference")
            AddOptionalTextBoxValue(vCreditSalesParam, "Notes")
            If DefaultParameters.ContainsKey("BatchCategory") Then vCreditSalesParam("BatchCategory") = DefaultParameters("BatchCategory")
            Try
              vResult = DataHelper.ConfirmCreditSaleTransaction(vCreditSalesParam)
            Catch vEx As ThreadAbortException
              Throw vEx
            Catch vEx As CareException
              SetErrorLabel(vEx.Message)
              mvSkipProcessing = True
            End Try
          End If
        ElseIf InitialParameters.OptionalValue("PaymentType").Trim = "CI" Then
          Dim vCreditAndCardSalesParam As ParameterList = Nothing
          GetCreditCustomerDetails(mvContactCompany, mvSalesAccount)
          If mvSkipProcessing = False Then
            If Me.FindControl("GrossAmount") IsNot Nothing Then mvGrossTotal = TryCast(Me.FindControl("GrossAmount"), TextBox).Text
            If Me.FindControl("CreditCardNumber") IsNot Nothing Then mvCardNumber = TryCast(Me.FindControl("CreditCardNumber"), TextBox).Text
            If Me.FindControl("IssueNumber") IsNot Nothing Then mvCardIssueNumber = TryCast(Me.FindControl("IssueNumber"), TextBox).Text
            If Me.FindControl("CardStartDate") IsNot Nothing Then mvStartDate = TryCast(Me.FindControl("CardStartDate"), TextBox).Text

            If InitialParameters.OptionalValue("PaymentService") = TNSHOSTED AndAlso FindControlByName(Me, "gatewayCardExpiryDateYear") IsNot Nothing Then
              If Me.FindControl("CardExpiryDate") IsNot Nothing Then mvExpiryDate = TryCast(Me.FindControl("CardExpiryDate"), TextBox).Text + TryCast(Me.FindControl("gatewayCardExpiryDateYear"), TextBox).Text.Substring(2, 2)
            Else
              If Me.FindControl("CardExpiryDate") IsNot Nothing Then mvExpiryDate = TryCast(Me.FindControl("CardExpiryDate"), TextBox).Text
            End If

            If Me.FindControl("SecurityCode") IsNot Nothing Then mvCSCValue = TryCast(Me.FindControl("SecurityCode"), TextBox).Text
            If Me.FindControl("CreditCardType") IsNot Nothing Then mvCardType = TryCast(Me.FindControl("CreditCardType"), DropDownList).SelectedValue

            Select Case InitialParameters.OptionalValue("PaymentService").Trim.ToUpper
              Case "PAYPAL"
                '   vCreditAndCardSalesParam = MakePaymentUsingPayPal()
              Case "PROTX"
                vCreditAndCardSalesParam = GetDetailsforProtX()
              Case "COMMSXL"
                vCreditAndCardSalesParam = GetDetailsforCommsXL()
              Case "TNSHOSTED"
                vCreditAndCardSalesParam = GetDetailsForTNS()
              Case AuthorisationService.SagePayHosted.GetServiceName()
                vCreditAndCardSalesParam = GetDetailsForSagePay()
              Case Else
                SetErrorLabel("Payment Option not supported. Value for payment_authorisation_services configuration option is not set")
            End Select
            vCreditAndCardSalesParam("Company") = mvContactCompany
            vCreditAndCardSalesParam("SalesLedgerAccount") = mvSalesAccount
            vResult = ConfirmCreditSalesTransaction(vCreditAndCardSalesParam)
          End If
        Else
          SetErrorLabel("Payment Type not supported.")
        End If
        SendConfirmationEmails(vResult)
      Catch vEX As ThreadAbortException
        Throw vEX
      Catch vException As Exception
        ProcessError(vException)
      Finally
        If InitialParameters.OptionalValue("PaymentService") = TNSHOSTED Then ClearSessionForTnsValues(True)
      End Try
    End If
  End Sub

  Private Function ConfirmCreditSalesTransaction(pCreditAndCardSalesParam As ParameterList) As ParameterList
    Dim vResult As New ParameterList
    Try
      AddOptionalTextBoxValue(pCreditAndCardSalesParam, "Reference")
      AddOptionalTextBoxValue(pCreditAndCardSalesParam, "Notes")
      vResult = DataHelper.ConfirmCreditAndCardSaleTransaction(pCreditAndCardSalesParam)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As CareException
      SetErrorLabel(vEx.Message)
      mvSkipProcessing = True
    End Try
    Return vResult
  End Function

  ''' <summary>
  ''' This function will send confrimation emails to the payers
  ''' </summary>
  ''' <param name="pResult"></param>
  ''' <remarks></remarks>
  Private Sub SendConfirmationEmails(pResult As ParameterList)
    If pResult IsNot Nothing AndAlso pResult.Count > 0 Then
      Session("BatchNumber") = pResult("BatchNumber").ToString
      Session("TransactionNumber") = pResult("TransactionNumber").ToString
      If pResult.ContainsKey("AuthorisationCode") Then pResult.Remove("AuthorisationCode")
      'Send confirmation mail to the payer
      If InitialParameters.OptionalValue("PayerDocument").Length > 0 Then SendConfirmationMailToPayer(pResult)
      'Send mail to the event bookers and booking delegates if any
      If Session("EventBookings") IsNot Nothing AndAlso BooleanValue(Session("EventBookings").ToString) Then
        pResult("ContactNumber") = mvContactNumber.ToString  'UserContactNumber().ToString
        SendConfirmationMailsToEventBookers(pResult)
        Session.Remove("EventBookings")
      End If
      'Remove the payer data from the session when we have successfully confirmed/processed the transaction
      Session.Remove("PayerContactNumber")
      Session.Remove("PayerAddressNumber")
      If Not InWebPageDesigner() Then GoToSubmitPage()
    End If
  End Sub


  ''' <summary>
  ''' This function will get the transaction details  
  ''' </summary>
  ''' <param name="pList"></param>
  ''' <remarks></remarks>
  Private Sub GetTransactionDetails(ByVal pList As ParameterList)
    Dim vAmtTextBox As TextBox = TryCast(Me.FindControl("GrossAmount"), TextBox)
    Dim vVatAmtTextBox As TextBox = TryCast(Me.FindControl("VatAmount"), TextBox)
    Dim vNetAmtTextBox As TextBox = TryCast(Me.FindControl("NetAmount"), TextBox)
    mvBatchNumber = pList("BatchNumber").ToString
    mvTransactionNumber = pList("TransactionNumber").ToString
    Session("BatchNumber") = mvBatchNumber
    Session("TransactionNumber") = mvTransactionNumber
    Dim vResult As String = DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, pList)
    Dim vDataTable As DataTable
    vDataTable = GetDataTable(vResult)
    Dim vAmount As Double = 0
    Dim vVatAmount As Double = 0
    For Each vDataRow As DataRow In vDataTable.Rows
      vAmount = vAmount + CType((vDataRow.Item("Amount").ToString), Decimal)
      If vDataRow.Item("VatAmount").ToString.Length > 0 Then vVatAmount = vVatAmount + CType((vDataRow.Item("VatAmount").ToString), Decimal)
      If vDataRow.Item("ItemType").ToString = "E" Then
        Session("EventBookings") = "Y"
      End If
      If vDataRow.Item("ItemType").ToString = "P" Then
        mvProductList.Append(vDataRow.Item("ProductDesc").ToString)
        mvProductList.Append("<BR>")
      End If
    Next
    If vAmount = 0 And vVatAmount = 0 AndAlso (InitialParameters.OptionalValue("PaymentType") = "CC" OrElse InitialParameters.OptionalValue("PaymentType") = "CI") Then
      HideControls()
    End If
    Dim vNetAmount As Double = vAmount - vVatAmount

    If vAmtTextBox IsNot Nothing Then vAmtTextBox.Text = vAmount.ToString("0.00")
    If vVatAmtTextBox IsNot Nothing Then vVatAmtTextBox.Text = vVatAmount.ToString("0.00")
    If vNetAmtTextBox IsNot Nothing Then vNetAmtTextBox.Text = vNetAmount.ToString("0.00")
    If vAmount = 0 And vVatAmount = 0 Then mvIsZeroAmount = True
    NetAmount = vNetAmount
    VatAmount = vVatAmount
    SetErrorLabel("")   'Specifically clear the error label as sometimes it already has a value
  End Sub

  ''' <summary>
  ''' This function will check the response from the TNS and show the error message to the user (If any) and restore the field values. 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CheckTNSResponse()
    If Session("FormErrorContents") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.INVALID_FIELD_VALUES Then
      If Session("Trader") IsNot Nothing AndAlso Session("Trader").ToString = "Y" Then HideAmountField()
      DisplayError() 'Display Error 
      SetFieldValues() 'Set all the fields with the values
      GetSessionDetailsForTNS() 'Get session details for TNS
      RenameControlsForTNS(tblDataEntry)
      ClearSessionForTnsValues(True)
    ElseIf Session("FormErrorContents") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.SESSION_EXPIRED Then

      If Session("Trader") IsNot Nothing AndAlso BooleanValue(Session("Trader").ToString) Then HideAmountField()
      SetFieldValues()
      ClearSessionValue("Actionlink") 'clear the session from then session details and get a new session 
      GetSessionDetailsForTNS()
      AddReturnLinkForTns(tblDataEntry)
      RenameControlsForTNS(tblDataEntry)
      ClearSessionForTnsValues(True)
      SetErrorLabel("Session expired for TNS")
    ElseIf Session("FormErrorCode") IsNot Nothing AndAlso CInt(Session("FormErrorCode").ToString) = TNSResult.SUCCESSFUL Then
      SetFieldValues()
      ClearSessionForTnsValues(False)

      'If this module is called from trader application then pass all the information using query string
      'else do the payment
      If Session("Trader") IsNot Nothing AndAlso BooleanValue(Session("Trader").ToString) Then
        AddTnsDetailsToQueryString()
        ClearSessionValue("Trader")
      Else
        MakePayment()
      End If
      ClearSessionForTnsValues(True)
    Else
      If Session("SessionURL") Is Nothing Then
        GetSessionDetailsForTNS()
        RenameControlsForTNS(tblDataEntry)
      End If
    End If

  End Sub

  ''' <summary>
  ''' Add hidden field to store the return url. TNS will use this url to pass the Card Authorisation information
  ''' back to the website
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
    If InitialParameters.OptionalValue("PaymentService") = TNSHOSTED AndAlso Not InWebPageDesigner() Then
      If Not mvIsBasketEmpty OrElse Session("Trader") IsNot Nothing AndAlso Session("Trader").ToString = "Y" Then
        If Session("ReturnUrl") Is Nothing Then
          If Session("ActionLink") IsNot Nothing Then Session.Remove("ActionLink")
          GetSessionDetailsForTNS()
        End If
        If Session("ReturnUrl") IsNot Nothing Then AddHiddenField(tblDataEntry, Session("ReturnUrl").ToString, "gatewayReturnURL")
        Dim vBatchCategory As String = CStr(If(Session("BatchCategory") IsNot Nothing, Session("BatchCategory"), ""))
        If DefaultParameters.OptionalValue("BatchCategory").Length > 0 OrElse vBatchCategory.Length > 0 Then AddHiddenField(tblDataEntry, DefaultParameters.OptionalValue("BatchCategory"), "BatchCategory")
      End If
    End If
  End Sub

  ''' <summary>
  ''' Create query string to pass the card details authentication information back to trader application
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub AddTnsDetailsToQueryString()
    Dim vParams As ParameterList = GetDetailsForTNS()
    Dim vSubmitParams As New StringBuilder
    With vSubmitParams
      For Each vKey As String In vParams.Keys
        Dim vNewKey As String = "&" + vKey
        .Append(vNewKey)
        .Append("=")
        .Append(vParams(vKey))
      Next
      'Pass Success to the trader application to process the transaction
      .Append("&Result=SUCCESS")
    End With
    GoToSubmitPage(vSubmitParams.ToString)
  End Sub

  Private Property NetAmount() As Nullable(Of Double)
    Get
      If Session("NetAmount") IsNot Nothing Then
        Return DoubleValue(Session("NetAmount").ToString)
      End If
      Return 0
    End Get
    Set(value As Nullable(Of Double))
      If value.HasValue Then
        Session("NetAmount") = value
      Else
        Session.Remove("NetAmount")
      End If
    End Set
  End Property

  Private Property VatAmount() As Nullable(Of Double)
    Get
      If Session("VatAmount") IsNot Nothing Then
        Return DoubleValue(Session("VatAmount").ToString)
      End If
      Return 0
    End Get
    Set(value As Nullable(Of Double))
      If value.HasValue Then
        Session("VatAmount") = value
      Else
        Session.Remove("VatAmount")
      End If
    End Set
  End Property

#Region "SagePay"
  Private Sub SetControlState()
    mvControlDictionary = New Dictionary(Of String, Boolean)
    mvControlDictionary.Add("CreditCardType", False)
    mvControlDictionary.Add("CreditCardNumber", False)
    mvControlDictionary.Add("CardExpiryDate", False)
    mvControlDictionary.Add("gatewayCardExpiryDateYear", False)
    mvControlDictionary.Add("IssueNumber", False)
    mvControlDictionary.Add("CardStartDate", False)
    mvControlDictionary.Add("SecurityCode", False)
    mvControlDictionary.Add("Reference", False)

    Dim vParams As New ParameterList(HttpContext.Current)
    vParams("BatchCategory") = If(DefaultParameters.ContainsKey("BatchCategory"), DefaultParameters("BatchCategory").ToString, String.Empty)
    Dim vResult As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams))

    If vResult Is Nothing OrElse String.IsNullOrEmpty(vResult("UseTokens").ToString) OrElse String.Compare(vResult("UseTokens").ToString, "N", True) = 0 Then
      mvControlDictionary.Add("TokenList", False)
      mvControlDictionary.Add("CreateToken", False)
    Else
      mvUseTokens = True
    End If

    mvControlDictionary.Add("Notes", False)
    mvControlDictionary.Add("TokenDesc", False)

    For Each vItem As KeyValuePair(Of String, Boolean) In mvControlDictionary
      SetControlVisible(vItem.Key, vItem.Value)
      If FindControlByName(Me, vItem.Key) IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent.Parent IsNot Nothing Then FindControlByName(Me, vItem.Key).Parent.Parent.Visible = vItem.Value
    Next
  End Sub

  Private Function GetDetailsForSagePay() As ParameterList
    Dim vReturnValue As String = String.Empty
    StoreSagePaySessionDetails()
    Dim vParameters As New ParameterList(HttpContext.Current)

    With vParameters
      Dim vTokenList As ListBox = TryCast(Me.FindControl("TokenList"), ListBox)
      If vTokenList IsNot Nothing AndAlso vTokenList.SelectedIndex > -1 Then
        .Add("Token", vTokenList.SelectedValue.ToString)
      ElseIf Me.FindControl("CreateToken") IsNot Nothing AndAlso TryCast(Me.FindControl("CreateToken"), CheckBox).Checked Then
        .Add("CreateToken", "Y")
        Cache("TokenDescription") = GetTextBoxText("TokenDesc")
      Else
        'Token is not switched ON
      End If
      .Add("ContactNumber", mvContactNumber)
      .Add("AddressNumber", GetContactAddress(mvContactNumber))
      .Add("Amount", GetTextBoxText("GrossAmount")) 'TryCast(Me.FindControl("GrossAmount"), TextBox).Text)
      .Add("Description", If(mvProductList.Length > 0, mvProductList.ToString, "UNKNOWN"))
      If (DefaultParameters.ContainsKey("BatchCategory")) Then .Add("BatchCategory", DefaultParameters("BatchCategory").ToString)
      .Add("MakeRequest", "Y")
    End With

    Dim vMechantDetails As ParameterList = mvAuthorisationService.CheckConnection(vParameters)
    If vMechantDetails IsNot Nothing Then
      Cache("VendorName") = vMechantDetails("MerchantId").ToString
      Response.Redirect(vMechantDetails("GatewayFormUrl").ToString)
    Else
      SetErrorLabel(vReturnValue)
    End If
    Return New ParameterList()
  End Function

  Private Sub StoreSagePaySessionDetails()
    If Request.QueryString("Status") Is Nothing Then Cache("ProcessPaymentPageUrl") = Request.Url
  End Sub

  Protected Overrides Function GetAuthorisationType() As AuthorisationService
    Dim vAuthorisationType As AuthorisationService = AuthorisationService.None
    Select Case InitialParameters.OptionalValue("PaymentService").Trim.ToUpper
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
    If Cache("ProcessPaymentPageUrl") IsNot Nothing Then Cache.Remove("ProcessPaymentPageUrl")
  End Sub
#End Region
End Class
