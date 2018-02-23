Imports System.Net
Imports Advanced.LanguageExtensions

Public Class SagePayHostedPayment
  Implements ICardAuthorisationService

  Private mvAddress As Uri
  Private mvRequestFields As New SortedList(Of String, String)
  Private mvResponseFields As String
  Private mvTimeout As Integer
  Private mvCardSettings As CardSettings
  Private Const MAX_SAGEPAY_NAME_LENGTH As Integer = 20 ' Set by Sage Pay
  Private Const IGNORE3DSECURE As String = "2"
  Private mvOverrideableParameters As New List(Of String)(New String() {"Vendor",
                                                                        "Description",
                                                                        "BillingFirstNames",
                                                                        "BillingPostcode",
                                                                        "BillingCountry",
                                                                        "DeliveryFirstNames",
                                                                        "DeliveryPostcode",
                                                                        "DeliveryCountry"})
  Private mvMustNotAddParameters As New List(Of String)(New String() {"Token"})

  Public ReadOnly Property OverrideableParameters As List(Of String)
    Get
      Return mvOverRideableParameters
    End Get

  End Property

  Public ReadOnly Property MustNotAddParameters As List(Of String)
    Get
      Return mvMustNotAddParameters
    End Get
  End Property

  Public Sub New(ByVal pUrl As String, ByVal pTimeOut As Integer)
    mvAddress = New Uri(pUrl)
    mvTimeout = pTimeOut
  End Sub

  Public Sub New(ByVal pCardSettings As CardSettings)
    mvCardSettings = pCardSettings
  End Sub

  Public Sub AddDigitalOrderField(ByVal pKey As String, ByVal pValue As String)
    mvRequestFields.Add(pKey, pValue)
  End Sub

  Public Function SendRequest(pRequest As String) As Boolean Implements ICardAuthorisationService.SendRequest
    Try
      'Create the web vRequest  
      Dim vRequest As HttpWebRequest = DirectCast(WebRequest.Create(mvCardSettings.GatewayHostAddress), HttpWebRequest)

      'Set type to POST  
      vRequest.Method = "POST"
      vRequest.ContentType = "application/x-www-form-urlencoded"
      vRequest.Timeout = mvCardSettings.TimeOut * 1000

      'Create a byte array of the data we want to send  
      Dim vByteData As Byte() = UTF8Encoding.UTF8.GetBytes(pRequest) '(GetRequestRaw)

      'Set the content length in the vRequest headers  
      vRequest.ContentLength = vByteData.Length

      'Write data
      Using vPostStream As System.IO.Stream = vRequest.GetRequestStream

        vPostStream.Write(vByteData, 0, vByteData.Length)
      End Using
      'Get response  
      Using vResponse As HttpWebResponse = TryCast(vRequest.GetResponse, HttpWebResponse)
        'Get the response stream  
        Dim vReader As New System.IO.StreamReader(vResponse.GetResponseStream)
        'Console application output
        mvResponseFields = vReader.ReadToEnd()
      End Using
      Return True
    Catch vEx As Exception
      Return False 'vEx.Message
    End Try
  End Function

  Private Function GetRequestRaw() As String
    Dim vData As New StringBuilder()
    For Each vPair As KeyValuePair(Of String, String) In mvRequestFields
      If vPair.Value.Length > 0 Then
        If (vData.Length > 0) Then vData.Append("&")
        vData.Append(vPair.Key & "=" & System.Web.HttpUtility.UrlEncode(vPair.Value))
      End If
    Next
    Return vData.ToString
  End Function

  Public Function GetCardType(ByVal pCardType As String) As String
    Dim vCardType As String
    Select Case pCardType
      Case "V", "O"
        'VISA		
        vCardType = "VISA"
      Case "A"
        'Mastercard
        vCardType = "MC"
      Case "SW"
        'Switch Card
        vCardType = "SWITCH"
      Case "DE"
        'Delta Card
        vCardType = "DELTA"
      Case "AX"
        'American Express
        vCardType = "AMEX"
      Case "SO"
        vCardType = "SOLO"
      Case "DC"
        vCardType = "DC"
      Case "JCB"
        vCardType = "JCB"
      Case "M"
        vCardType = "MAESTRO"
      Case Else
        vCardType = ""
    End Select
    Return vCardType
  End Function

  Function GetResultField(ByVal pFieldName As String) As String
    Dim vItems() As String
    Dim vIndex As Integer
    Dim vItem As String
    Dim vReponse As String = String.Empty

    vItems = Split(mvResponseFields, Chr(13))
    For vIndex = LBound(vItems) To UBound(vItems)
      vItem = Replace(vItems(vIndex), Chr(10), "")
      If InStr(vItem, pFieldName & "=") = 1 Then
        ' found
        vReponse = Right(vItem, Len(vItem) - Len(pFieldName) - 1)
        Exit For
      End If
    Next
    Return vReponse
  End Function

  Public Function GetErrorCode(pResult As String) As String Implements ICardAuthorisationService.GetErrorCode
    Throw New NotSupportedException("ICardAuthorisationService.GetErrorCode not supported for SagePay")
  End Function

  Public Function GetRawResponseData() As String Implements ICardAuthorisationService.GetRawResponseData
    Throw New NotSupportedException("ICardAuthorisationService.GetRawResponseData not supported for SagePay")
  End Function

  Public Function GetRequestData() As String Implements ICardAuthorisationService.GetRequestData
    Dim vParams As New ParameterList
    Return ParseRequestData(vParams)
  End Function

  Public Function GetRequestData(pSession As String, pTransactionID As Integer, pAmount As String) As String Implements ICardAuthorisationService.GetRequestData
    Throw New NotSupportedException("ICardAuthorisationService.GetRequestData not supported for SagePay")
  End Function

  Public Function GetResponseData() As ParameterList Implements ICardAuthorisationService.GetResponseData
    Dim vItems() As String
    Dim vIndex As Integer
    Dim vReponse As String = String.Empty
    Dim vResponseList As New ParameterList()

    vItems = Split(mvResponseFields, Chr(13))
    For vIndex = LBound(vItems) To UBound(vItems)
      If vItems(vIndex).Contains("="c) Then
        vResponseList.Add(vItems(vIndex).Split("="c)(0).Trim, vItems(vIndex).Split("="c)(1).Trim)
      End If
    Next
    Return vResponseList

  End Function

  ''' <summary>
  ''' This function formats the request data into format required by SAGEPAYHOSTED
  ''' </summary>
  ''' <returns>Request data in the format require by the SAGEPAYHOSTED</returns>
  ''' <remarks></remarks>
  Private Function ParseRequestData(ByVal pParams As ParameterList) As String
    Dim vData As New StringBuilder()
    For Each vPair As DictionaryEntry In pParams
      If vPair.Value.ToString.Length > 0 Then
        If (vData.Length > 0) Then vData.Append("&")
        vData.Append(vPair.Key.ToString & "=" & System.Web.HttpUtility.UrlEncode(vPair.Value.ToString))
      End If
    Next
    Return vData.ToString
  End Function

  Public Function GetRequestData(pParameterList As ParameterList) As String Implements ICardAuthorisationService.GetRequestData
    Dim mvRequestFields As New SortedList(Of String, String)
    Dim vBackOffice As Boolean = If(pParameterList.ContainsKey("SmartClient") AndAlso String.Compare(pParameterList("SmartClient"), "Y", True) = 0, True, False)
    mvRequestFields.Add("VPSProtocol", "3.00")
    mvRequestFields.Add("TxType", "PAYMENT")
    mvRequestFields.Add("Vendor", mvCardSettings.MerchantID)
    mvRequestFields.Add("VendorTxCode", pParameterList("TransactionNumber").ToString)
    mvRequestFields.Add("Amount", pParameterList("Amount").ToString)
    mvRequestFields.Add("Currency", mvCardSettings.Currency)
    mvRequestFields.Add("Description", pParameterList("Description").ToString)
    mvRequestFields.Add("NotificationURL", mvCardSettings.ReturnAddress)
    mvRequestFields.Add("BillingSurname", TruncateString(pParameterList("ContactSurname").ToString, MAX_SAGEPAY_NAME_LENGTH))
    mvRequestFields.Add("BillingFirstNames", TruncateString(FirstWord(pParameterList("ContactFirstname").ToString), MAX_SAGEPAY_NAME_LENGTH))
    mvRequestFields.Add("BillingAddress1", pParameterList("Address1").ToString)
    mvRequestFields.Add("BillingAddress2", pParameterList("Address2").ToString)
    mvRequestFields.Add("BillingCity", pParameterList("City").ToString)
    mvRequestFields.Add("BillingPostcode", pParameterList("PostCode").ToString)
    mvRequestFields.Add("BillingCountry", pParameterList("Country").ToString)
    mvRequestFields.Add("DeliverySurname", TruncateString(pParameterList("ContactSurname").ToString, MAX_SAGEPAY_NAME_LENGTH))
    mvRequestFields.Add("DeliveryFirstNames", TruncateString(FirstWord(pParameterList("ContactFirstname").ToString), MAX_SAGEPAY_NAME_LENGTH))
    mvRequestFields.Add("DeliveryAddress1", pParameterList("Address1").ToString)
    mvRequestFields.Add("DeliveryAddress2", pParameterList("Address2").ToString)
    mvRequestFields.Add("DeliveryCity", pParameterList("City").ToString)
    mvRequestFields.Add("DeliveryPostcode", pParameterList("PostCode").ToString)
    mvRequestFields.Add("DeliveryCountry", pParameterList("Country").ToString)
    If vBackOffice Then
      mvRequestFields.Add("Apply3DSecure", IGNORE3DSECURE)
    End If
    If pParameterList.ContainsKey("Token") Then
      mvRequestFields.Add("StoreToken", "1")
      mvRequestFields.Add("Token", pParameterList("Token"))
    ElseIf pParameterList.ContainsKey("CreateToken") AndAlso String.Compare(pParameterList("CreateToken"), "Y", True) = 0 Then
      mvRequestFields.Add("CreateToken", "1")
    Else
      'no tokens
    End If
    Dim vAdditionalParameters As SortedList(Of String, String) = GetAdditionalParameters()
    mvRequestFields = MergeParameters(mvRequestFields, vAdditionalParameters)
    Dim vData As New StringBuilder()
    For Each vPair As KeyValuePair(Of String, String) In mvRequestFields
      If vPair.Value.Length > 0 Then
        If (vData.Length > 0) Then vData.Append("&")
        vData.Append(vPair.Key & "=" & System.Web.HttpUtility.UrlEncode(vPair.Value))
      End If
    Next
    Return vData.ToString
  End Function
  Private Function GetAdditionalParameters() As SortedList(Of String, String)
    Dim vParameters As New SortedList(Of String, String)
    If Not String.IsNullOrEmpty(Me.mvCardSettings.AdditionalParameters) Then
      Dim vAdditionalParameters As String() = Split(Me.mvCardSettings.AdditionalParameters, ",")
      For Each vParamter As String In vAdditionalParameters
        Dim vParm() As String = Split(vParamter, "=")
        vParameters.Add(vParm(0), vParm(1))
      Next
    End If
    Return vParameters
  End Function
  Private Function MergeParameters(ByVal pOriginalParameters As SortedList(Of String, String), ByVal pAdditionalParameters As SortedList(Of String, String)) As SortedList(Of String, String)
    Dim vMergedParameters As SortedList(Of String, String) = pOriginalParameters
    For Each vAdditionalParameter As KeyValuePair(Of String, String) In pAdditionalParameters
      If vMergedParameters.ContainsKey(vAdditionalParameter.Key) Then
        If vMergedParameters(vAdditionalParameter.Key).IsNullOrWhitespace AndAlso OverrideableParameters.Contains(vAdditionalParameter.Key) Then
          vMergedParameters(vAdditionalParameter.Key) = vAdditionalParameter.Value
        End If
      Else
        If Not MustNotAddParameters.Contains(vAdditionalParameter.Key) Then
          vMergedParameters.Add(vAdditionalParameter.Key, vAdditionalParameter.Value)
        End If
      End If
    Next
    Return vMergedParameters
  End Function
End Class


