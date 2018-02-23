Public Class DisplayTransactions
    Inherits CareWebControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDisplayTransactions, tblDataEntry, "", "")
      Dim vDGR As DataGrid = CType(Me.FindControl("TransactionData"), DataGrid)
      Dim vList As New ParameterList(HttpContext.Current)
      'vList("DocumentColumns") = "Y"
      vList("ContactNumber") = UserContactNumber()
      vList("SystemColumns") = "Y"
      vList("CarePortal") = "Y"

      vList.Add("WebPageItemNumber", Me.WebPageItemNumber)
      If Not vList.Contains("WPD") Then vList.Add("WPD", "Y")

      If (InitialParameters.Contains("PaymentMethod")) Then vList.Add("PaymentMethod", InitialParameters("PaymentMethod"))
      If (InitialParameters.Contains("BatchCategory")) Then vList.Add("BatchCategory", InitialParameters("BatchCategory"))

      Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vList)

      If Request.QueryString("RN") IsNot Nothing AndAlso InitialParameters.Contains("PrintReceiptPage") Then
        Dim vDt As DataTable = GetDataTable(vResult, True)
        Dim vRow As Integer = IntegerValue(Request.QueryString("RN"))
        Session("TransactionNumber") = vDt.Rows(vRow).Item("TransactionNumber")
        Session("BatchNumber") = vDt.Rows(vRow).Item("BatchNumber")
        ProcessRedirect("default.aspx?pn=" & InitialParameters("PrintReceiptPage").ToString)
      Else
        DataHelper.FillGrid(vResult, vDGR, "")

        If InitialParameters.Contains("PrintReceiptPage") AndAlso vDGR.Items.Count > 0 Then
          Dim vColumn As New BoundColumn()
          vDGR.Columns.AddAt(0, vColumn)
          vDGR.DataBind()

          For vRow As Integer = 0 To vDGR.Items.Count - 1
            If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
              vDGR.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&RN={2}'"">", InitialParameters("HyperlinkText").ToString(), Me.WebPageNumber, vRow)
            Else
              vDGR.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&RN={1}'>{2}</a>", Me.WebPageNumber, vRow, InitialParameters("HyperlinkText").ToString())
            End If
          Next
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
    End Sub

    Public Sub New()
        mvNeedsAuthentication = True
    End Sub
End Class
