Partial Public Class AddMemberCI
  Inherits AddMemberCC

  Protected Overrides Sub InitControls()
    InitialiseControls(CareNetServices.WebControlTypes.wctAddMemberCI, tblDataEntry, "CreditCardNumber,CardExpiryDate", "DirectNumber,MobileNumber")
  End Sub

  Protected Overrides Sub PreMemberCreation(pList As ParameterList)
    MyBase.PreMemberCreation(pList)

    'We are about to take the payment so check for a credit customer record
    Dim vList As New ParameterList(HttpContext.Current)
    vList("Company") = DefaultParameters("Company")
    vList("ContactNumber") = pList("PayerContactNumber")
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vList)
    If vDataTable Is Nothing Then
      Dim vCreditCustParams As New ParameterList(HttpContext.Current)
      vCreditCustParams("ContactNumber") = pList("PayerContactNumber")
      vCreditCustParams("AddressNumber") = pList("PayerAddressNumber")
      vCreditCustParams("Company") = DefaultParameters("Company")
      vCreditCustParams("CreditCategory") = DefaultParameters("CreditCategory")
      Dim vResult As ParameterList = DataHelper.AddCreditCustomer(vCreditCustParams)
    End If
  End Sub

  Protected Overrides Sub PrePaymentPlanPayment(pList As ParameterList)
    MyBase.PrePaymentPlanPayment(pList)
    pList("CcWithInvoice") = "Y"
  End Sub

  Protected Overrides Function GetPaymentMethod() As String
    Return DataHelper.ControlValue(DataHelper.ControlTables.credit_sales_controls, DataHelper.ControlValues.payment_method)
  End Function

End Class