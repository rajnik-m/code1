''' <summary>
''' Create an instance of the Card Authorisation service 
''' </summary>
''' <remarks></remarks>
Public Class CardAuthorisationServiceFactory


  Public Shared Function GetAuthorisationServiceProvider(ByVal pEnv As CDBEnvironment) As ICardAuthorisationService
    Return GetAuthorisationServiceProvider(pEnv, "")
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pEnv"></param>
  ''' <param name="pBatchCategory"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function GetAuthorisationServiceProvider(ByVal pEnv As CDBEnvironment, ByVal pBatchCategory As String) As ICardAuthorisationService
    Return GetAuthorisationServiceProvider(pEnv, Nothing, pBatchCategory, "")
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pEnv"></param>
  ''' <param name="pBatchCategory"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function GetAuthorisationServiceProvider(ByVal pEnv As CDBEnvironment, ByVal pBatchCategory As String, ByVal pMerchantRetailNumber As String) As ICardAuthorisationService
    Return GetAuthorisationServiceProvider(pEnv, Nothing, pBatchCategory, pMerchantRetailNumber)
  End Function
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pEnv"></param>
  ''' <param name="pCardSale"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function GetAuthorisationServiceProvider(ByVal pEnv As CDBEnvironment, ByVal pCardSale As CardSale) As ICardAuthorisationService
    Return GetAuthorisationServiceProvider(pEnv, pCardSale, "", "")
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pEnv"></param>
  ''' <param name="pCardSale"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function GetAuthorisationServiceProvider(ByVal pEnv As CDBEnvironment, ByVal pCardSale As CardSale, ByVal pBatchCategory As String, ByVal pMerchantRetailNumber As String) As ICardAuthorisationService
    Dim vAuthorisationProvider As ICardAuthorisationService = Nothing
    Dim vCardSettings As New CardSettings(pEnv, pCardSale, pBatchCategory, pMerchantRetailNumber)
    Select Case vCardSettings.CardAuthorisationServiceType
      Case CreditCardAuthorisation.OnlineAuthorisationTypes.TnsHosted
        vAuthorisationProvider = New TnsHostedPayment(vCardSettings)
      Case CreditCardAuthorisation.OnlineAuthorisationTypes.SagePayHosted
        vAuthorisationProvider = New SagePayHostedPayment(vCardSettings)
      Case Else
        vAuthorisationProvider = Nothing
    End Select
    Return vAuthorisationProvider
  End Function
End Class