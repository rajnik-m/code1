Partial Public Class ModifyDD
  Inherits CareWebControl

  Private mvPaymentPlanNumber As Integer
  Private mvRenewalAmount As Double

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctModifyDD, tblDataEntry)
      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, UserContactNumber)
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          If vRow("DirectDebitStatus").ToString.StartsWith("Y") AndAlso vRow("CancellationReason").ToString.Length = 0 Then
            mvPaymentPlanNumber = IntegerValue(vRow("PaymentPlanNumber").ToString)
            mvRenewalAmount = CDbl(vRow("RenewalAmount"))
            Exit For
          End If
        Next
      End If
      If mvPaymentPlanNumber > 0 Then
        SetTextBoxText("RenewalAmount", mvRenewalAmount.ToString("#.00"))
      Else
        ShowMessageOnly("Cannot find an existing Direct Debit to Modify")
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vList As New ParameterList(HttpContext.Current)
        vList("PaymentPlanNumber") = mvPaymentPlanNumber
        vList("Balance") = GetTextBoxText("Balance")
        vList("Product") = InitialParameters("Product")
        vList("Rate") = InitialParameters("Rate")
        AddDefaultParameters(vList)
        DataHelper.UpdatePaymentPlan(CareNetServices.XMLPaymentPlanUpdateTypes.xpputAddDetailLine, vList)
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
End Class