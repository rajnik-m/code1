Partial Public Class AddPaymentPlanDD
  Inherits CareWebControl
  Implements ICareParentWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddPaymentPlanDD, tblDataEntry, "AccountName", "DirectNumber,MobileNumber")
      SetAmountOrBalance("Balance")
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vReturnList As ParameterList = AddNewContact()
    Dim vList As New ParameterList(HttpContext.Current)
    AddPaymentPlanParameters(vList, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_dd))
    vList("Product") = InitialParameters("Product")
    vList("Rate") = InitialParameters("Rate")
    AddOptionalTextBoxValue(vList, "Balance")
    AddDDParameters(vList)
    Dim vResult As ParameterList = DataHelper.AddPaymentPlan(CareNetServices.ppType.pptDD, vList)
    AddGiftAidDeclaration(IntegerValue(vList("PayerContactNumber").ToString))
    ProcessChildControls(vReturnList)
    If vResult.ContainsKey("DirectDebitNumber") AndAlso vResult("DirectDebitNumber").ToString.Length > 0 Then Session("DirectDebitNumber") = vResult("DirectDebitNumber").ToString
    If vResult.ContainsKey("FrequencyAmount") AndAlso vResult("FrequencyAmount").ToString.Length > 0 Then Session("DirectDebitFrequencyAmount") = vResult("FrequencyAmount").ToString
    If DefaultParameters.ContainsKey("PaymentFrequency") AndAlso DefaultParameters("PaymentFrequency").ToString.Length > 0 Then Session("DirectDebitPaymentFrequency") = DefaultParameters("PaymentFrequency").ToString
    If vResult.ContainsKey("DirectDebitClaimDate") AndAlso vResult("DirectDebitClaimDate").ToString.Length > 0 Then Session("DirectDebitClaimDate") = vResult("DirectDebitClaimDate").ToString
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub
End Class