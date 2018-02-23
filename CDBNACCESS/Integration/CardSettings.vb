''' <summary>
''' This card will read the Card Setting information from DataBase \ Configuration and set the details required  
''' </summary>
''' <remarks></remarks>
Public Class CardSettings

#Region "Private Variables"
  Private mvEnv As CDBEnvironment = Nothing
  Private mvTimeOut As Integer = 0
  Private mvGatewayHostAddressHostAddress As String = String.Empty
  Private mvMerchantID As String = String.Empty
  Private mvCardSale As CardSale = Nothing
  Private mvMerchant As String = String.Empty
  Private mvMerchantDetails As DataRow = Nothing
  Private mvGatewayActionAddress As String = String.Empty
  Private mvReturnAddress As String = String.Empty
  Private mvGatewayPassword As String = String.Empty
  Private mvBatchCategory As String = String.Empty
  Private mvMerchantRetailNumber As String = String.Empty
#End Region

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pEnv"></param>
  ''' <param name="pCardSale"></param>
  ''' <param name="pBatchCategory"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pCardSale As CardSale, ByVal pBatchCategory As String, ByVal pMerchantRetailNumber As String)
    mvMerchantRetailNumber = pMerchantRetailNumber
    mvEnv = pEnv
    mvCardSale = pCardSale
    mvBatchCategory = pBatchCategory
    Init()
  End Sub

#Region "Private Function"

  ''' <summary>
  ''' Method to initialise the merchant details  
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub Init()
    Dim vDT As DataTable = New SQLStatement(mvEnv.Connection, "merchant_id,access_code,user_name,user_password,gateway_password,gateway_host_url,gateway_form_url,card_details_page_url,return_url,additional_parameters", "merchant_details", New CDBFields(New CDBField("merchant_retail_number", MerchantRetailNumber))).GetDataTable()
    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then IncorrectSetup()
    mvMerchantDetails = vDT.Rows(0)
  End Sub

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub IncorrectSetup()
    Select Case CardAuthorisationServiceType
      Case CreditCardAuthorisation.OnlineAuthorisationTypes.TnsHosted
        RaiseError(DataAccessErrors.daeTNSHostedPaymentNotSetUp)
      Case CreditCardAuthorisation.OnlineAuthorisationTypes.SagePayHosted
        RaiseError(DataAccessErrors.daeSagePayHostedNotSetup)
      Case Else
        '
    End Select
  End Sub

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetMerchantRetailNumber() As String
    Dim vMerchantRetailNumber As String = String.Empty
    Dim vDataTable As CDBDataTable = Nothing
    Dim vWhereFields As New CDBFields()

    If String.IsNullOrEmpty(mvMerchantRetailNumber) Then
      If mvBatchCategory.Length > 0 Then
        vWhereFields.Add("batch_category", mvBatchCategory)
        vDataTable = New CDBDataTable(mvEnv, New SQLStatement(mvEnv.Connection, "merchant_retail_number", "batch_categories", vWhereFields))
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then vMerchantRetailNumber = vDataTable.Rows(0).Item("merchant_retail_number")
      End If

      If vMerchantRetailNumber.Length = 0 Then
        vMerchantRetailNumber = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMerchantRetailNumber)
        If vMerchantRetailNumber.Length = 0 Then IncorrectSetup()
      End If
    End If
    Return vMerchantRetailNumber

  End Function
#End Region


#Region "Properties"
  ''' <summary>
  ''' Property to read the user defined TimeOut time from the configuration. Application will use this before timing out if the service provider's
  ''' server does not respond back with the validation results 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TimeOut As Integer
    Get
      If mvTimeOut = 0 Then mvTimeOut = IntegerValue(mvEnv.GetConfig("fp_cc_authorisation_timeout", "195"))
      Return mvTimeOut
    End Get
  End Property

  ''' <summary>
  ''' Property to read the host url inforation for the authorisation service provider
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property GatewayHostAddress() As String
    Get
      If mvGatewayHostAddressHostAddress = String.Empty Then
        If mvMerchantDetails Is Nothing Then IncorrectSetup()
        mvGatewayHostAddressHostAddress = mvMerchantDetails("gateway_host_url").ToString
      End If
      Return mvGatewayHostAddressHostAddress
    End Get
  End Property

  ''' <summary>
  ''' Property to read the Merchant ID from the batch categories or merchant details table. 
  ''' If this is not set then application will raise an error. 
  ''' </summary>
  ''' <value></value>
  ''' <returns>Merchant ID</returns>
  ''' <remarks></remarks>
  Public ReadOnly Property MerchantID() As String
    Get
      If mvMerchantID = String.Empty Then
        If mvMerchantDetails Is Nothing Then IncorrectSetup()
        mvMerchantID = mvMerchantDetails("merchant_id").ToString
      End If
      Return mvMerchantID
    End Get
  End Property

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property GatewayActionAddress() As String
    Get
      If mvGatewayActionAddress = String.Empty Then
        If mvMerchantDetails Is Nothing Then IncorrectSetup()
        mvGatewayActionAddress = mvMerchantDetails("gateway_form_address").ToString
      End If
      Return mvGatewayActionAddress
    End Get
  End Property
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ReturnAddress() As String
    Get
      If mvReturnAddress = String.Empty Then
        If mvMerchantDetails Is Nothing Then IncorrectSetup()
        mvReturnAddress = mvMerchantDetails("return_url").ToString
      End If
      Return mvReturnAddress
    End Get
  End Property

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property GatewayPassword() As String
    Get
      If mvGatewayPassword = String.Empty Then
        If mvMerchantDetails Is Nothing Then IncorrectSetup()
        mvGatewayPassword = mvMerchantDetails("gateway_password").ToString
      End If
      Return mvGatewayPassword
    End Get
  End Property

  Public ReadOnly Property MerchantRetailNumber As String
    Get
      If mvMerchantRetailNumber.Length = 0 Then mvMerchantRetailNumber = GetMerchantRetailNumber()
      Return mvMerchantRetailNumber
    End Get
  End Property

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property CardAuthorisationServiceType As CreditCardAuthorisation.OnlineAuthorisationTypes
    Get
      Dim vCardAuthorisationType As CreditCardAuthorisation.OnlineAuthorisationTypes
      Dim vType As String = mvEnv.GetConfig("fp_cc_authorisation_type")
      Select Case vType
        Case "CSXL210FE"
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.CommsXL
        Case "SCXLVPCSCP"
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.SecureCXL
        Case "PROTX"
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.ProtX
        Case "TNSHOSTED"
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.TnsHosted
        Case "SAGEPAYHOSTED"
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.SagePayHosted
        Case Else
          vCardAuthorisationType = CreditCardAuthorisation.OnlineAuthorisationTypes.None
          RaiseError(DataAccessErrors.daeInvalidConfig, "fp_cc_authorisation_type")
      End Select
      Return vCardAuthorisationType
    End Get
  End Property

  ''' <summary>
  ''' Currency code from the financial controls
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks>Currently only support GBP</remarks>
  Public ReadOnly Property Currency As String
    Get
      Return If(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode).Length > 0, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode), "GBP")
    End Get
  End Property

  Public ReadOnly Property AdditionalParameters As String
    Get
      Return Me.mvMerchantDetails.Item("additional_parameters").ToString()
    End Get
  End Property

#End Region
End Class
