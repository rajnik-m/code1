Public Class SagePayHostedService
  Implements IAuthorisationService

  Public Function GetPaymentDetails() As ParameterList Implements IAuthorisationService.GetPaymentDetails
    Return New ParameterList()
  End Function

  Public Function CheckConnection(ByVal pParameterList As ParameterList) As ParameterList Implements IAuthorisationService.CheckConnection
    Dim vResult As DataRow = Nothing
    Dim vParameterList As New ParameterList(HttpContext.Current)
    Try
      vResult = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, pParameterList))
      If vResult IsNot Nothing AndAlso vResult.Table.Columns.Contains("GatewayFormUrl") AndAlso vResult("GatewayFormUrl") IsNot Nothing Then
        vParameterList.Add("GatewayFormUrl", vResult("GatewayFormUrl").ToString())
        vParameterList.Add("MerchantId", vResult("MerchantId").ToString())
        Return vParameterList
      Else
        Return Nothing
      End If
    Catch vEx As Exception
      Return Nothing
    End Try
  End Function
End Class

