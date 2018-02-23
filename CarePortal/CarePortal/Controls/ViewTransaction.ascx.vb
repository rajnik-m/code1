Partial Public Class ViewTransaction
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Private mvBatchNumber As String = String.Empty
  Private mvTransactionNumber As String = String.Empty
  Private mvLineNumber As String = String.Empty
  Private mvAmount As Double
  Private mvVATAmount As Double

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctViewTransaction, tblDataEntry)
      If Request.QueryString("Mode") IsNot Nothing AndAlso Request.QueryString("Mode") = "Delete" Then
        Dim vDeleteList As New ParameterList(HttpContext.Current)
        If Request.QueryString("BN") IsNot Nothing AndAlso Request.QueryString("BN").Length > 0 Then
          vDeleteList("BatchNumber") = Request.QueryString("BN")
        End If
        If Request.QueryString("TN") IsNot Nothing AndAlso Request.QueryString("TN").Length > 0 Then
          vDeleteList("TransactionNumber") = Request.QueryString("TN")
        End If
        If Request.QueryString("LN") IsNot Nothing AndAlso Request.QueryString("LN").Length > 0 Then
          vDeleteList("LineNumber") = Request.QueryString("LN")
        End If
        DeleteTransactionLine(vDeleteList)
      End If
      Dim vBaseList As BaseDataList = TryCast(Me.FindControl("EventTransactionData"), BaseDataList)
      Dim vLabel As Label = TryCast(Me.FindControl("WarningMessage"), Label)
      Dim vAmtTextBox As TextBox = TryCast(Me.FindControl("GrossAmount"), TextBox)
      Dim vVatAmtTextBox As TextBox = TryCast(Me.FindControl("VatAmount"), TextBox)
      Dim vNetAmtTextBox As TextBox = TryCast(Me.FindControl("NetAmount"), TextBox)
      Dim vList As New ParameterList(HttpContext.Current)
      If GetShoppingBasketTransaction(UserContactNumber, vList) Then
        mvBatchNumber = vList("BatchNumber").ToString
        mvTransactionNumber = vList("TransactionNumber").ToString
        vList("SystemColumns") = "Y"
        vList("WebPageItemNumber") = Me.WebPageItemNumber
        If vBaseList IsNot Nothing Then
          Dim vResult As String = DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, vList)
          DataHelper.FillGrid(vResult, vBaseList, "")
          vLabel.Visible = False
          If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
            Dim vDataGrid As DataGrid = CType(vBaseList, DataGrid)
            Dim vAmount As Decimal = 0
            Dim vVatAmount As Decimal = 0
            Dim vAmountPos As Integer = 0
            Dim vVatAmountPos As Integer = 0
            Dim vItemTypePos As Integer = 0
            Dim vLineNumberPos As Integer = 0
            Dim vProductPos As Integer = 0
            Dim vRatePos As Integer = 0
            Dim vPage As String = String.Empty
            Dim vEditColumn As New BoundColumn()
            Dim vDeleteColumn As New BoundColumn()

            vEditColumn.HeaderText = ""
            vDataGrid.Columns.AddAt(0, vEditColumn)
            vDeleteColumn.HeaderText = ""
            vDataGrid.Columns.AddAt(1, vDeleteColumn)
            vDataGrid.DataBind()

            For vCount As Integer = 0 To vDataGrid.Columns.Count - 1
              Dim vBoundColumn As TemplateColumn = TryCast(vDataGrid.Columns(vCount), TemplateColumn)
              If vBoundColumn IsNot Nothing Then
                If DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "Amount" Then
                  vAmountPos = vCount
                ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "VatAmount" Then
                  vVatAmountPos = vCount
                ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "ItemType" Then
                  vItemTypePos = vCount
                ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "LineNumber" Then
                  vLineNumberPos = vCount
                ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "Product" Then
                  vProductPos = vCount
                End If
              End If
            Next
            If InitialParameters.Contains("ProductUpdatePage") Then vPage = InitialParameters("ProductUpdatePage").ToString
            For vCount As Integer = 0 To vDataGrid.Items.Count - 1
              If vPage.Length = 0 Then
                vDataGrid.Columns(0).Visible = False
              Else
                If DirectCast(vDataGrid.Items(vCount).Cells(vItemTypePos).Controls(0), ITextControl).Text = "P" Then
                  If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                    vDataGrid.Items(vCount).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&BN={2}&TN={3}&LN={4}&PR={5}'"">", InitialParameters("HyperlinkText1").ToString, vPage, vList("BatchNumber").ToString, vList("TransactionNumber").ToString, DirectCast(vDataGrid.Items(vCount).Cells(vLineNumberPos).Controls(0), ITextControl).Text, DirectCast(vDataGrid.Items(vCount).Cells(vProductPos).Controls(0), ITextControl).Text)
                  Else
                    vDataGrid.Items(vCount).Cells(0).Text = "<a href='default.aspx?pn=" & vPage & "&BN=" & vList("BatchNumber").ToString & "&TN=" & vList("TransactionNumber").ToString & "&LN=" & DirectCast(vDataGrid.Items(vCount).Cells(vLineNumberPos).Controls(0), ITextControl).Text & "&PR=" & DirectCast(vDataGrid.Items(vCount).Cells(vProductPos).Controls(0), ITextControl).Text & "'>" & InitialParameters("HyperlinkText1").ToString & "</a>"
                  End If
                End If
              End If
              If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                vDataGrid.Items(vCount).Cells(1).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&BN={2}&TN={3}&LN={4}&Mode=Delete'"">", InitialParameters("HyperlinkText2").ToString, Request.QueryString("pn"), vList("BatchNumber").ToString, vList("TransactionNumber").ToString, DirectCast(vDataGrid.Items(vCount).Cells(vLineNumberPos).Controls(0), ITextControl).Text)
              Else
                vDataGrid.Items(vCount).Cells(1).Text = "<a href='default.aspx?pn=" & Request.QueryString("pn") & "&BN=" & vList("BatchNumber").ToString & "&TN=" & vList("TransactionNumber").ToString & "&LN=" & DirectCast(vDataGrid.Items(vCount).Cells(vLineNumberPos).Controls(0), ITextControl).Text & "&Mode=Delete'>" & InitialParameters("HyperlinkText2").ToString & "</a>"
              End If
              mvAmount = mvAmount + CDec(DirectCast(vDataGrid.Items(vCount).Cells(vAmountPos).Controls(0), ITextControl).Text)
              If DirectCast(vDataGrid.Items(vCount).Cells(vVatAmountPos).Controls(0), ITextControl).Text <> "&nbsp;" Then mvVATAmount = mvVATAmount + CDec(DirectCast(vDataGrid.Items(vCount).Cells(vVatAmountPos).Controls(0), ITextControl).Text)
            Next
          End If
          SetTextBoxText("GrossAmount", Format$(mvAmount, "Fixed"))
          SetTextBoxText("VatAmount", Format$(mvVATAmount, "Fixed"))
          SetTextBoxText("NetAmount", Format$(mvAmount - mvVATAmount, "Fixed"))
        End If
      Else
        vBaseList.Visible = False
        vAmtTextBox.Visible = False
        vVatAmtTextBox.Visible = False
        vNetAmtTextBox.Visible = False
        vLabel.Visible = True
        SetParentParentVisible("GrossAmount", False)
        SetParentParentVisible("VatAmount", False)
        SetParentParentVisible("NetAmount", False)
        SetParentVisible("Submit", False)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not InWebPageDesigner() Then
      GoToSubmitPage()
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vPage As String = InitialParameters.OptionalValue("ProductUpdatePage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)
      Dim vCount As Integer = e.Item.Controls.Count
      Dim vEditLink As New Literal
      Dim vDeleteLink As New Literal

      mvAmount = mvAmount + DoubleValue(vDrv.Row("Amount").ToString)
      mvVATAmount = mvVATAmount + DoubleValue(vDrv.Row("VatAmount").ToString)

      If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
        If vDrv.Row("ItemType").ToString = "P" Then
          vEditLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&BN={2}&TN={3}&LN={4}&PR={5}'"">", InitialParameters("HyperlinkText1").ToString, vPage, mvBatchNumber, mvTransactionNumber, vDrv.Row("LineNumber").ToString, vDrv.Row("Product").ToString)
        End If
        vDeleteLink.Text = String.Format("&nbsp;&nbsp;&nbsp;<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='default.aspx?pn={1}&BN={2}&TN={3}&LN={4}&MODE=Delete'"">", InitialParameters("HyperlinkText2").ToString, Request.QueryString("pn"), mvBatchNumber, mvTransactionNumber, vDrv.Row("LineNumber").ToString)
      Else
        If vDrv.Row("ItemType").ToString = "P" Then
          vEditLink.Text = "<a href='default.aspx?pn=" & vPage & "&BN=" & mvBatchNumber & "&TN=" & mvTransactionNumber & "&LN=" & vDrv.Row("LineNumber").ToString & "&PR=" & vDrv.Row("Product").ToString & "'>" & InitialParameters("HyperlinkText1").ToString & "</a>"
        End If
        vDeleteLink.Text = "&nbsp;&nbsp;&nbsp;<a href='default.aspx?pn=" & Request.QueryString("pn") & "&BN=" & mvBatchNumber & "&TN=" & mvTransactionNumber & "&LN=" & vDrv.Row("LineNumber").ToString & "&Mode=Delete'>" & InitialParameters("HyperlinkText2").ToString & "</a>"
      End If
      If vCount > 0 Then
        e.Item.Controls(vCount - 1).Parent.Controls.Add(vEditLink)
        e.Item.Controls(vCount - 1).Parent.Controls.Add(vDeleteLink)
      End If
    End If
  End Sub

  Private Sub DeleteTransactionLine(ByVal pList As ParameterList)
    Try
      DataHelper.DeleteItem(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, pList)
      ProcessRedirect(String.Format("Default.aspx?pn={0}", Request.QueryString("pn")))
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub
End Class