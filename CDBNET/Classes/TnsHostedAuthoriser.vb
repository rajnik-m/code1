<AuthorisationMethod("TNSHOSTED")>
Public Class TnsHostedAuthoriser
  Inherits WebBasedCardAuthoriser

  Friend Sub New(browser As WebBrowser)
    MyBase.New(browser)
  End Sub

  Protected Overrides Sub OnWebBrowserNavigated()
    If Me.ReturnParameter("Result").Equals("success", StringComparison.InvariantCultureIgnoreCase) AndAlso Not Me.IsAuthorised Then
      Me.VendorCode = Me.ReturnParameter("TnsSession")
      Me.CardType = Me.ReturnParameter("CreditCardType")
      Me.CardDigits = Me.ReturnParameter("CreditCardNumber")
      Me.CardExpiry = Me.ReturnParameter("CardExpiryDate")
      Me.AuthCode = Me.ReturnParameter("SecurityCode")
      Me.IsCancelled = False
      Me.IsAuthorised = True
      Me.RaiseProcesesingComplete()
    End If
  End Sub

  Public Overrides Sub RequestAuthorisation(contactNumber As Integer, addressNumber As Integer, transactionType As String, transactionAmount As Integer, batchCategory As String, merchantRetailNumber As String)
    Me.AuthCode = String.Empty
    Me.CardDigits = String.Empty
    Me.CardExpiry = String.Empty
    Me.CardType = String.Empty
    Me.IsAuthorised = False
    Me.Status = String.Empty
    Me.StatusDetail = String.Empty
    Me.VendorCode = String.Empty

    Dim params As New ParameterList(True)

    params.Add("BatchCategory", batchCategory)
    params.Add("Amount", String.Format("{0:F2}", transactionAmount / 100))
    params.Add("ContactNumber", contactNumber)
    params.Add("AddressNumber", addressNumber)
    params.Add("Description", transactionType)
    params.Add("MakeRequest", "Y")
    params.Add("SmartClient", "Y")

    Try
      Dim result As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, params)
      If result IsNot Nothing AndAlso
         result("CardDetailsPageUrl") IsNot Nothing Then
        Dim cardDetailsUrl As New StringBuilder(result("CardDetailsPageUrl").ToString)
        cardDetailsUrl.Append("&")
        cardDetailsUrl.Append("Trader=")
        cardDetailsUrl.Append("Y")
        cardDetailsUrl.Append("&")
        cardDetailsUrl.Append("ContactNumber=")
        cardDetailsUrl.Append(params("ContactNumber"))
        cardDetailsUrl.Append("&")
        cardDetailsUrl.Append("Amount=")
        cardDetailsUrl.Append(params("Amount"))
        cardDetailsUrl.Append("&")
        cardDetailsUrl.Append("BatchCategory=")
        cardDetailsUrl.Append(params("BatchCategory"))

        Me.Browser.Stop()
        Me.Browser.Navigate(cardDetailsUrl.ToString)
        Me.Browser.Refresh(WebBrowserRefreshOption.Completely)
        Me.Browser.AllowNavigation = True
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enTNSHostedPaymentNotSetUp, CareException.ErrorNumbers.enSagePayHostedNotSetup,
          CareException.ErrorNumbers.enConnectionFailure, CareException.ErrorNumbers.enInvalidRequest
          ShowErrorMessage(vEx.Message)
      End Select
    End Try
  End Sub

  Public Overrides Sub SetServerValues(list As ParameterList)
    list("TnsSession") = Me.VendorCode
    list("CreditCardType") = Me.CardType
    list("CreditCardNumber") = Me.CardDigits
    list("CardExpiryDate") = Me.CardExpiry
    list("SecurityCode") = Me.AuthCode
  End Sub
End Class
