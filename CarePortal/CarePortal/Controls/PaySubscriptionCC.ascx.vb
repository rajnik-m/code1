Partial Public Class PaySubscriptionCC
  Inherits CareWebControl

  Dim mvPaymentPlanNumber As Integer
  Dim mvNextPaymentDue As String
  Dim mvNextPaymentAmount As String

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      SupportsOnlineCCAuthorisation = True
      InitialiseControls(CareNetServices.WebControlTypes.wctRenewSubscriptionCC, tblDataEntry)

      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = UserContactNumber()
      Dim vTable As DataTable = DataHelper.GetNextPaymentData(vList)
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          If vRow("MembershipType").ToString.Length = 0 Then
            mvPaymentPlanNumber = IntegerValue(vRow("PaymentPlanNumber").ToString)
            mvNextPaymentDue = vRow("NextPaymentDue").ToString
            mvNextPaymentAmount = vRow("NextPaymentAmount").ToString
            Exit For
          End If
        Next
      End If
      If mvPaymentPlanNumber > 0 Then
        SetTextBoxText("NextPaymentDue", mvNextPaymentDue)
        SetTextBoxText("NextPaymentAmount", mvNextPaymentAmount)
      Else
        ShowMessageOnly("You have no subscriptions that need payments at this time")
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
        SetErrorLabel("")
        Dim vPayList As New ParameterList(HttpContext.Current)
        vPayList("ContactNumber") = UserContactNumber()
        vPayList("AddressNumber") = UserAddressNumber()
        vPayList("PaymentPlanNumber") = mvPaymentPlanNumber
        vPayList("Amount") = mvNextPaymentAmount
        vPayList("BankAccount") = DefaultParameters("BankAccount")
        vPayList("Source") = DefaultParameters("Source")
        AddCCParameters(vPayList)
        AddUserParameters(vPayList)
        Dim vSkipProcessing As Boolean
        Try
          DataHelper.AddPaymentPlanPayment(vPayList)
        Catch vEx As ThreadAbortException
          Throw vEx
        Catch vEx As CareException
          SetErrorLabel(vEx.Message)
          vSkipProcessing = True
        End Try
        If vSkipProcessing = False Then GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

End Class