Public Class SelectPayPlanForDD
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelecttPayPlanForDD, tblDataEntry)
      GetPaymentPlans()
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
      Catch vEx As ThreadAbortException
        Throw vEx
      End Try
    End If
  End Sub

  Private Sub GetPaymentPlans()
    Dim vDGR As DataGrid = CType(Me.FindControl("PayPlanData"), DataGrid)

    Dim vList As New ParameterList(HttpContext.Current)
    Dim vContactNumber As Integer = GetContactNumberFromParentGroup()
    vList("ContactNumber") = vContactNumber.ToString
    vList("CancellationReason") = ""
    vList("DirectDebit") = "N"
    vList("StandingOrder") = "N"
    vList("CCCA") = "N"

    vList("SystemColumns") = "Y"
    vList.Add("WPD", "Y")
    vList("WebPageItemNumber") = Me.WebPageItemNumber

    Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, vList)
    DataHelper.FillGrid(vResult, vDGR, "UnprocessedPayments = 0", , True, If(InitialParameters.ContainsKey("HyperlinkText"), InitialParameters("HyperlinkText").ToString, "Convert"))
    If vDGR.Items.Count <= 0 Then
      vDGR.Visible = False
    ElseIf vDGR.Items.Count = 1 Then
      Dim vOptionPos As Integer = GetDataGridItemIndex(vDGR, "PaymentPlanNumber")
      Session("SelectedPaymentPlan") = vDGR.Items(0).Cells(vOptionPos).Text
      Session("SelectedPaymentPlanContactNumber") = GetContactNumberFromParentGroup()
      GoToSubmitPage()
    Else
      DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
      vDGR.Visible = True
    End If
  End Sub


  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Session("SelectedPaymentPlan") = e.CommandArgument
    Session("SelectedPaymentPlanContactNumber") = GetContactNumberFromParentGroup()
    GoToSubmitPage()
  End Sub

End Class



