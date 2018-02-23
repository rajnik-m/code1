Public Class AutorisationServiceFactory
  Public Shared Function GetAuthorisationService(service As String) As IAuthorisationService
    Select Case service

      Case "SAGEPAYHOSTED"
        Return New SagePayHostedService()
      Case Else
        Return Nothing
    End Select

  End Function
End Class
