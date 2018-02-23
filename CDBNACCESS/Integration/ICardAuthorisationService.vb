Namespace Access
  ''' <summary>
  ''' All the authorisatiion services should implement this interface for consistency 
  ''' </summary>
  ''' <remarks>Only TnsHostedPayment class is implementing this at the moment''' </remarks>
  Public Interface ICardAuthorisationService
    Function SendRequest(ByVal pRequestData As String) As Boolean
    Function GetResponseData() As ParameterList
    Function GetRawResponseData() As String
    Function GetRequestData() As String
    Function GetRequestData(ByVal pSession As String, ByVal pTransactionID As Integer, ByVal pAmount As String) As String
    Function GetRequestData(ByVal pParameterList As ParameterList) As String
    Function GetErrorCode(ByVal pResult As String) As String
  End Interface
End Namespace

