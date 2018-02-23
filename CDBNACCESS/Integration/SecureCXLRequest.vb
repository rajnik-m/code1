Imports System.Net

Public Class SecureCXLRequest

  Private mvAddress As Uri
  Private mvRequestFields As New SortedList(Of String, String)(New VPCStringComparer)
  Private mvResponseFields As New SortedList(Of String, String)(New VPCStringComparer)
  Private mvTimeout As Integer

  Private Shared mvSensitiveFields As IList(Of String) = New List(Of String)(New String() {"vpc_CardNum",
                                                                                           "vpc_CardExp",
                                                                                           "vpc_CardSecurityCode",
                                                                                           "vpc_AVS_Street01",
                                                                                           "vpc_AVS_PostCode",
                                                                                           "vpc_AVS_Country",
                                                                                           "vpc_AVS_City"}).AsReadOnly

  Public Structure MerchantDetails
    Public Shared MerchantRetailNo As String = ""
    Public Shared MerchantID As String = ""
    Public Shared AccessCode As String = ""
    Public Shared UserName As String = ""
    Public Shared UserPassword As String = ""
    Public Shared Timeout As Integer

    Public Shared Sub Init(ByVal pRecordSet As CDBRecordSet)
      If pRecordSet.Fetch Then
        'MerchantRetailNo = pRecordSet.Fields("merchant_retail_number").Value
        MerchantID = pRecordSet.Fields("merchant_id").Value
        AccessCode = pRecordSet.Fields("access_code").Value
        UserName = pRecordSet.Fields("user_name").Value
        UserPassword = pRecordSet.Fields("user_password").Value
      End If
      pRecordSet.CloseRecordSet()
    End Sub
  End Structure

  Public Sub New(ByVal pURL As String, ByVal pAPIVersion As String)
    Me.New(pURL, pAPIVersion, False)
  End Sub

  Public Sub New(ByVal pURL As String, ByVal pAPIVersion As String, ByVal pAddUserDetails As Boolean)
    mvAddress = New Uri(pURL)
    AddDigitalOrderField("vpc_Version", pAPIVersion)
    AddDigitalOrderField("vpc_AccessCode", MerchantDetails.AccessCode)
    AddDigitalOrderField("vpc_Merchant", MerchantDetails.MerchantID)
    If pAddUserDetails Then
      AddDigitalOrderField("vpc_User", MerchantDetails.UserName)
      AddDigitalOrderField("vpc_Password", MerchantDetails.UserPassword)
    End If
    mvTimeout = MerchantDetails.Timeout
  End Sub

  Public Sub AddDigitalOrderField(ByVal pKey As String, ByVal pValue As String)
    mvRequestFields.Add(pKey, pValue)
  End Sub

  Public Function SendRequest() As String
    Dim vTraceSource As New TraceSource("SecureCXL")
    Try
      'Create the web vRequest  
      Dim vRequest As HttpWebRequest = DirectCast(WebRequest.Create(mvAddress), HttpWebRequest)
      Dim vMessage As String = "Sending SecureCXL request: "
      For Each vField As KeyValuePair(Of String, String) In mvRequestFields
        vMessage &= vField.Key & " = '" & If(mvSensitiveFields.Contains(vField.Key), New String("*"c, vField.Value.Length), vField.Value) & "', "
      Next vField
      vMessage = vMessage.Substring(0, vMessage.Length - 2) & "."
      vTraceSource.TraceInformation(vMessage)
      'Set type to POST  
      vRequest.Method = "POST"
      vRequest.ContentType = "application/x-www-form-urlencoded"
      vRequest.Timeout = mvTimeout * 1000

      'Create a byte array of the data we want to send  
      Dim vByteData As Byte() = UTF8Encoding.UTF8.GetBytes(GetRequestRaw)

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
        Dim vRawResponse As String = vReader.ReadToEnd()
        Dim vResponses As String() = vRawResponse.Split("&"c)
        For Each vResponseField As String In vResponses
          Dim vField As String() = vResponseField.Split("="c)
          mvResponseFields.Add(vField(0), System.Web.HttpUtility.UrlDecode(vField(1)))
        Next
      End Using
      vMessage = "Recieved SecureCXL response: "
      For Each vField As KeyValuePair(Of String, String) In mvResponseFields
        vMessage &= vField.Key & " = '" & If(mvSensitiveFields.Contains(vField.Key), New String("*"c, vField.Value.Length), vField.Value) & "', "
      Next vField
      vMessage = vMessage.Substring(0, vMessage.Length - 2) & "."
      vTraceSource.TraceInformation(vMessage)
      Return ""
    Catch vEx As Exception
      vTraceSource.TraceEvent(TraceEventType.Error, 0, vEx.ToString)
      Return vEx.Message
    Finally
      vTraceSource.Close()
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

  Public Function GetResultField(ByVal pKey As String) As String
    Return GetResultField(pKey, "")
  End Function

  Public Function GetResultField(ByVal pKey As String, ByVal pDefValue As String) As String
    'String value
    Dim vValue As String = ""
    If mvResponseFields.TryGetValue(pKey, vValue) Then
      Return vValue
    Else
      Return pDefValue
    End If
  End Function

  ''' <summary>
  ''' Customised Compare Class
  ''' </summary>
  ''' <para>
  ''' The Virtual Payment Client need to use an Ordinal comparison to Sort on 
  ''' the field names to create the MD5 Signature for validation of the message. 
  ''' This class provides a Compare method that is used to allow the sorted list 
  ''' to be ordered imports an Ordinal comparison.
  ''' </para>
  ''' <remarks></remarks>
  Private Class VPCStringComparer
    Implements System.Collections.Generic.IComparer(Of String)

    ''' <summary>
    ''' Compare method imports Ordinal comparison
    ''' </summary>
    ''' <param name="x">The first string in the comparison.</param>
    ''' <param name="y">The second string in the comparison.</param>
    ''' <returns></returns>
    ''' <remarks>An int containing the result of the comparison.</remarks>
    Public Function Compare(ByVal x As String, ByVal y As String) As Integer Implements System.Collections.Generic.IComparer(Of String).Compare
      'Return if we are comparing the same object or one of the objects is null, since we don't need to go any further.
      Dim vValue1 As String = x.ToString
      Dim vValue2 As String = y.ToString
      If vValue1 = vValue2 Then Return 0
      If vValue1 = "" Then Return -1
      If vValue2 = "" Then Return 1

      'Get the CompareInfo object to use for comparing
      If vValue1.Length > 0 AndAlso vValue2.Length > 0 Then
        'Compare imports an Ordinal Comparison.
        Return System.Globalization.CompareInfo.GetCompareInfo("en-GB").Compare(vValue1, vValue2, System.Globalization.CompareOptions.Ordinal)
      End If
      Throw New ArgumentException("vValue1 and vValue2 should be strings.")
    End Function

  End Class

End Class

Public Class SecureCXLCodes

  Public Shared Function GetResponseCodeDescription(ByVal pResponseCode As String) As String
    If pResponseCode Is Nothing OrElse String.Compare(pResponseCode, "null", True) = 0 OrElse pResponseCode.Length = 0 Then
      Return "null response"
    Else
      Select Case pResponseCode
        Case "0"
          Return "Transaction Successful"
        Case "1"
          Return "Transaction Declined"
        Case "2"
          Return "Bank Declined Transaction"
        Case "3"
          Return "No Reply from Bank"
        Case "4"
          Return "Expired Card"
        Case "5"
          Return "Insufficient Funds"
        Case "6"
          Return "Error Communicating with Bank"
        Case "7"
          Return "Payment Server detected an error"
        Case "8"
          Return "Transaction Type Not Supported"
        Case "9"
          Return "Bank declined transaction (Do not contact Bank)"
        Case "A"
          Return "Transaction Aborted"
        Case "B"
          Return "Transaction Declined - Contact the Bank"
        Case "C"
          Return "Transaction Cancelled"
        Case "D"
          Return "Deferred transaction has been received and is awaiting processing"
        Case "E"
          Return "Issuer Returned a Referral Response"
        Case "F"
          Return "3-D Secure Authentication failed"
        Case "I"
          Return "Card Security Code verification failed"
        Case "L"
          Return "Shopping Transaction Locked (Please try the transaction again later)"
        Case "N"
          Return "Cardholder is not enrolled in Authentication scheme"
        Case "P"
          Return "Transaction has been received by the Payment Adaptor and is being processed"
        Case "R"
          Return "Transaction was not processed - Reached limit of retry attempts allowed"
        Case "S"
          Return "Duplicate SessionID"
        Case "T"
          Return "Address Verification Failed"
        Case "U"
          Return "Card Security Code Failed"
        Case "V"
          Return "Address Verification and Card Security Code Failed"
        Case Else
          Return "Unable to be determined"
      End Select
    End If
  End Function

  Public Shared Function GetCSCDescription(ByVal pCSCResultCode As String) As String
    If pCSCResultCode IsNot Nothing AndAlso Not pCSCResultCode.Length = 0 Then
      If [String].Compare(pCSCResultCode, "Unsupported", True) = 0 Then
        Return "CSC not supported or there was no CSC data provided"
      Else
        Select Case pCSCResultCode
          Case "M"
            Return "Exact code match"
          Case "S"
            Return "Merchant has indicated that CSC is not present on the card (MOTO situation)"
          Case "P"
            Return "Code not processed"
          Case "U"
            Return "Card issuer is not registered and/or certified"
          Case "N"
            Return "Code invalid or not matched"
          Case Else
            Return "Unable to be determined"
        End Select
      End If
    Else
      Return ""
    End If
  End Function

  Public Shared Function GetAVSDescription(ByVal pAVSResultCode As String) As String
    If pAVSResultCode IsNot Nothing AndAlso Not pAVSResultCode.Length = 0 Then
      If [String].Compare(pAVSResultCode, "Unsupported", True) = 0 Then
        Return "AVS not supported or there was no AVS data provided"
      Else
        Select Case pAVSResultCode
          Case "X"
            Return "Exact match - address and 9 digit ZIP/postal code"
          Case "Y"
            Return "Exact match - address and 5 digit ZIP/postal code"
          Case "S"
            Return "Service not supported or address not verified (international transaction)"
          Case "G"
            Return "Issuer does not participate in AVS (international transaction)"
          Case "A"
            Return "Address match only"
          Case "W"
            Return "9 digit ZIP/postal code matched, Address not Matched"
          Case "Z"
            Return "5 digit ZIP/postal code matched, Address not Matched"
          Case "R"
            Return "Issuer system is unavailable"
          Case "U"
            Return "Address unavailable or not verified"
          Case "E"
            Return "Address and ZIP/postal code not provided"
          Case "N"
            Return "Address and ZIP/postal code not matched"
          Case "0"
            Return "AVS not requested"
          Case Else
            Return "Unable to be determined"
        End Select
      End If
    Else
      Return ""
    End If
  End Function

End Class