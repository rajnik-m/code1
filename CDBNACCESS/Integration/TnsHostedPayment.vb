Imports System.Net
Namespace Access

  Public Class TnsHostedPayment
    Implements ICardAuthorisationService

#Region "Private variables"
    Private mvCardSettings As CardSettings = Nothing
    Dim mvResponseString As String = String.Empty
    Private Const TRANS_START_RANGE As Long = 10000000000
    Private Const MOTO As String = "MOTO"
#End Region


    ''' <summary>
    ''' Constructor with CardSettings
    ''' </summary>
    ''' <param name="pCardSettings"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pCardSettings As CardSettings)
      mvCardSettings = pCardSettings
    End Sub

#Region "Public functions"

    ''' <summary>
    ''' Sends the request to Get the Session ID from the TNS to start the Authentication Process. This method then also used to
    ''' send the transaction details to the TNS for actual payment once the card details are authorised
    ''' </summary>
    ''' <returns>True if connection is successful else false</returns>
    ''' <remarks></remarks>
    Public Function SendRequest(ByVal pRequestData As String) As Boolean Implements ICardAuthorisationService.SendRequest
      Try
        Dim vTrace As New TraceSource("TnsHosted")
        'Create the web vRequest 
        Dim vRequest As HttpWebRequest = DirectCast(WebRequest.Create(mvCardSettings.GatewayHostAddress), HttpWebRequest)

        'Set type to POST  
        vRequest.Method = "POST"
        vRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
        vRequest.Timeout = mvCardSettings.TimeOut * 1000

        Dim vCredentials As String = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes("" + ":" + mvCardSettings.GatewayPassword))
        vRequest.Headers.Add("Authorization", "Basic " + vCredentials)

        'Create a byte array of the data we want to send  
        Dim vByteData As Byte() = UTF8Encoding.UTF8.GetBytes(pRequestData)

        vTrace.TraceInformation(pRequestData)

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
          mvResponseString = vReader.ReadToEnd()
        End Using
        Return CheckResponse()
      Catch vEx As Exception
        mvResponseString = vEx.Message
        Return False 'vEx.Message
      End Try
    End Function

    ''' <summary>
    ''' Parse the Tns resposne data and store in the key value pair (ParameterList)
    ''' </summary>
    ''' <returns>Tns response data as parameterList</returns>
    ''' <remarks></remarks>
    Public Function GetResponseData() As ParameterList Implements ICardAuthorisationService.GetResponseData
      If mvResponseString.Length = 0 Then SendRequest(GetRequestData)

      Dim vResponsePair As New ParameterList
      If mvResponseString.Length > 0 Then
        Dim vResponses As String() = mvResponseString.Split("&"c)
        For Each vField As String In vResponses
          Dim vResponseFields As String() = vField.Split("="c)
          vResponsePair.Add(vResponseFields(0), System.Web.HttpUtility.UrlDecode(vResponseFields(1)))
        Next
      End If
      If vResponsePair.Count > 0 Then vResponsePair.Add("MerchantRetailNumber", mvCardSettings.MerchantRetailNumber)
      Return vResponsePair
    End Function

    ''' <summary>
    ''' Returns the unformatted response data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRawResponseData() As String Implements ICardAuthorisationService.GetRawResponseData
      Return mvResponseString
    End Function

    ''' <summary>
    ''' Get the response data for the specified field
    ''' </summary>
    ''' <param name="pFieldName">response field name</param>
    ''' <returns>response field value</returns>
    ''' <remarks></remarks>
    Public Function GetResponseData(ByVal pFieldName As String) As String
      Dim vResultField As String = String.Empty
      Dim vParams As ParameterList = GetResponseData()
      If vParams.ContainsKey(pFieldName) Then vResultField = vParams(pFieldName).ToString
      Return vResultField
    End Function

    ''' <summary>
    ''' Creates the TNS request data. Transaction request will be sent to TNS for Pay operation to charge the amount from the card
    ''' </summary>
    ''' <param name="pSession">TNS session ID used when card details were send for authentication to TNS</param>
    ''' <param name="pTransactionID">Unique Transaction ID number should be more than TRANS_START_RANGE</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRequestData(ByVal pSession As String, ByVal pTransactionId As Integer, ByVal pAmount As String) As String Implements ICardAuthorisationService.GetRequestData
      Dim vRequest As New ParameterList
      Dim vTransactionId As Long = TRANS_START_RANGE + pTransactionId
      vRequest.Add("apiOperation", "PAY")
      vRequest.Add("order.id", vTransactionId.ToString)
      vRequest.Add("transaction.amount", pAmount)
      vRequest.Add("transaction.currency", mvCardSettings.Currency)
      vRequest.Add("transaction.reference", vTransactionId)
      vRequest.Add("transaction.id", vTransactionId)
      vRequest.Add("merchant", mvCardSettings.MerchantID)
      vRequest.Add("apiPassword", mvCardSettings.GatewayPassword)
      vRequest.Add("transaction.source", MOTO)
      vRequest.Add("cardDetails.session", pSession)
      Return ParseRequestData(vRequest)
    End Function

    ''' <summary>
    ''' This function adds the fields required to send the request to TNS 
    ''' </summary>
    ''' <returns>Key value pair of Requested data</returns>
    ''' <remarks></remarks>
    Public Function GetRequestData() As String Implements ICardAuthorisationService.GetRequestData
      Dim vParams As New ParameterList
      vParams("apiOperation") = "CREATE_SESSION"
      vParams("merchant") = mvCardSettings.MerchantID
      vParams("apiPassword") = mvCardSettings.GatewayPassword
      Return ParseRequestData(vParams)
    End Function

    ''' <summary>
    ''' Tns response codes 
    ''' </summary>
    ''' <param name="pResult">Result string </param>
    ''' <returns>Response code</returns>
    ''' <remarks></remarks>
    Public Function GetErrorCode(ByVal pResult As String) As String Implements ICardAuthorisationService.GetErrorCode
      Dim vResponseCode As String = ""
      Select Case pResult
        Case "SUCCESS"
          vResponseCode = "00"
        Case "DECLINED"
          vResponseCode = "05"
        Case "REFERRED"
          vResponseCode = "01"
        Case "INVALIDCARD"
          vResponseCode = "14"
        Case "INVALIDEXPIRY"
          vResponseCode = "54"
        Case "FAILURE"
          vResponseCode = "55"
        Case Else
          vResponseCode = "99"
      End Select
      Return vResponseCode
    End Function
#End Region

#Region "Private functions"


    ''' <summary>
    ''' This function formats the request data into format required by TNS
    ''' </summary>
    ''' <returns>Request data in the format require by the TNS</returns>
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

    ''' <summary>
    ''' Check the response string for success status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckResponse() As Boolean
      If mvResponseString.Length > 0 AndAlso GetResponseData("result").ToLower = "success" Then
        Return True
      Else
        Return False
      End If
    End Function
#End Region

    Public Function GetRequestData(pParameterList As ParameterList) As String Implements ICardAuthorisationService.GetRequestData
      Throw New NotSupportedException("ICardAuthorisationService.GetRequestData not supported for TNS Hosted")
    End Function
  End Class
End Namespace
