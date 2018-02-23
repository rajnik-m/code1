Public Interface IAuthorisationService

  Function GetPaymentDetails() As ParameterList
  Function CheckConnection(ByVal pParameters As ParameterList) As ParameterList
End Interface
