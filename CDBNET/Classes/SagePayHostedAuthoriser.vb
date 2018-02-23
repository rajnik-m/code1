<AuthorisationMethod("SAGEPAYHOSTED")>
Public Class SagePayHostedAuthoriser
  Inherits WebBasedCardAuthoriser

  Friend Sub New(browser As WebBrowser)
    MyBase.New(browser)
  End Sub

  Protected Overrides Sub OnWebBrowserNavigated()
    If Not String.IsNullOrWhiteSpace(Me.ReturnParameter("Status")) Then
      Me.Status = Me.ReturnParameter("Status").ToUpper
      Me.StatusDetail = Me.ReturnParameter("StatusDetail")
      Me.VendorCode = Me.ReturnParameter("VendorTxCode")
      Me.CardType = Me.ReturnParameter("CardType")
      Me.CardDigits = Me.ReturnParameter("Last4Digit")
      Me.Token = Me.ReturnParameter("Token")
      Me.CardExpiry = Me.ReturnParameter("ExpiryDate")
      Me.AuthCode = Me.ReturnParameter("BankAuthCode")
      Select Case Me.Status
        Case "OK"
          Me.IsCancelled = False
          Me.IsAuthorised = True
          Me.RaiseProcesesingComplete()
        Case "ABORT"
          Me.IsCancelled = True
          Me.IsAuthorised = False
          Me.RaiseProcesesingComplete()
        Case Else
          Me.IsCancelled = False
          Me.IsAuthorised = False
          Me.RaiseProcesesingComplete()
      End Select
    End If
  End Sub

  Public Overrides Sub RequestAuthorisation(contactNumber As Integer,
                                            addressNumber As Integer,
                                            transactionType As String,
                                            transactionAmount As Integer,
                                            batchCategory As String,
                                            merchantRetailNumber As String)

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
    params.Add("MerchantRetailNumber", merchantRetailNumber)

    If Me.CreateToken Then
      params("CreateToken") = "Y"
    ElseIf Not String.IsNullOrWhiteSpace(Me.Token) Then
      params.Add("Token", Me.Token)
    End If

    Try
      Dim result As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, params)
      If result IsNot Nothing AndAlso result("GatewayFormUrl") IsNot Nothing Then
        Me.Browser.Navigate(result("GatewayFormUrl").ToString)
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
    list("Token") = Me.Token
    list("CreditCardType") = Me.CardType
    list("CardDigits") = Me.CardDigits
    list("CardExpiryDate") = Me.CardExpiry
    list("SagePayStatus") = Me.Status
    list("StatusDetail") = Me.StatusDetail
    list("VendorCode") = Me.VendorCode
  End Sub

End Class
