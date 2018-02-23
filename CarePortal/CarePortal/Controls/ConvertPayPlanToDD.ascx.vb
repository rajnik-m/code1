Public Class ConvertPayPlanToDD
  Inherits CareWebControl

  Private mvPaymentPlanNumber As Integer
  Private mvContactNumber As Integer
  Private mvAddressNumber As Integer

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctConvertPayPlanToDD, tblDataEntry)
      If Not InWebPageDesigner() Then
        If Session("SelectedPaymentPlan") IsNot Nothing AndAlso Session("SelectedPaymentPlan").ToString.Length > 0 Then
          mvPaymentPlanNumber = IntegerValue(Session("SelectedPaymentPlan").ToString)
        Else
          RaiseError(DataAccessErrors.daeSessionValueNotSet, "SelectedPaymentPlan")
        End If
        If Session("SelectedPaymentPlanContactNumber") IsNot Nothing AndAlso Session("SelectedPaymentPlanContactNumber").ToString.Length > 0 Then
          mvContactNumber = IntegerValue(Session("SelectedPaymentPlanContactNumber").ToString)
        Else
          RaiseError(DataAccessErrors.daeSessionValueNotSet, "SelectedPaymentPlanContactNumber")
        End If
        GetPaymentPlan()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vList As New ParameterList(HttpContext.Current)
        vList("BankAccount") = DefaultParameters("BankAccount")
        vList("PaymentPlanNumber") = mvPaymentPlanNumber.ToString
        vList("AutoPayContactNumber") = mvContactNumber.ToString
        vList("AutoPayAddressNumber") = mvAddressNumber.ToString
        vList("BranchName") = "UNKNOWN"
        vList("UserID") = UserContactNumber()
        vList("CarePortal") = "Y"
        AddDDParameters(vList)
        vList("AutoPaySource") = DefaultParameters("Source")

        Dim vReturnList As ParameterList = DataHelper.AddDirectDebit(vList)
        If vReturnList.ContainsKey("AutoPaymentNumber") AndAlso vReturnList("AutoPaymentNumber").ToString.Length > 0 Then Session("DirectDebitNumber") = vReturnList("AutoPaymentNumber").ToString
        If vReturnList.ContainsKey("PaymentFrequency") AndAlso vReturnList("PaymentFrequency").ToString.Length > 0 Then Session("DirectDebitPaymentFrequency") = vReturnList("PaymentFrequency").ToString
        If vReturnList.ContainsKey("FrequencyAmount") AndAlso vReturnList("FrequencyAmount").ToString.Length > 0 Then Session("DirectDebitFrequencyAmount") = vReturnList("FrequencyAmount").ToString
        If vReturnList.ContainsKey("DirectDebitClaimDate") AndAlso vReturnList("DirectDebitClaimDate").ToString.Length > 0 Then Session("DirectDebitClaimDate") = vReturnList("DirectDebitClaimDate").ToString

        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vError As Exception
        ProcessError(vError)
      End Try
    End If
  End Sub

  Private Sub GetPaymentPlan()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("PaymentPlanNumber") = mvPaymentPlanNumber.ToString
    vList("ContactNumber") = mvContactNumber.ToString
    vList("CancellationReason") = ""
    vList("DirectDebit") = "N"
    vList("StandingOrder") = "N"
    vList("CCCA") = "N"

    Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, vList)
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(vTable)
    If vRow Is Nothing Then
      ShowMessageOnlyFromLabel("WarningMessage1", "Cannot find payment plan to convert to direct debit")
    Else
      SetTextBoxText("FrequencyAmount", vRow("FrequencyAmount").ToString)
      mvAddressNumber = IntegerValue(vRow("AddressNumber").ToString)
    End If
  End Sub

End Class



