Public Class ProductSelection
  Inherits CareWebControl
  Private mvHyperLinkText As String = ""
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectProducts, tblDataEntry)
      If InitialParameters.ContainsKey("HyperlinkText") Then mvHyperLinkText = InitialParameters("HyperlinkText").ToString
      If Me.FindControl("SearchProduct") IsNot Nothing Then
        CType(Me.FindControl("SearchProduct"), TextBox).MaxLength = 100
        If Request.QueryString("Product") IsNot Nothing Then
          CType(Me.FindControl("SearchProduct"), TextBox).Text = Request.QueryString("Product")
        End If
      End If
      If Not IsPostBack Then FindProducts()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub FindProducts()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vQueryString As String = ""
    If GetTextBoxText("SearchProduct").Length > 0 Then
      vList("SearchProduct") = "*" & GetTextBoxText("SearchProduct") & "*"
      vQueryString = "&Product=" & GetTextBoxText("SearchProduct")
    End If
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    If InitialParameters.ContainsKey("SalesGroup") Then vList("SalesGroup") = InitialParameters("SalesGroup").ToString
    If InitialParameters.ContainsKey("SecondaryGroup") Then vList("SecondaryGroup") = InitialParameters("SecondaryGroup").ToString
    If InitialParameters.ContainsKey("ProductCategory") Then vList("ProductCategory") = InitialParameters("ProductCategory").ToString
    If UserContactNumber() > 0 Then vList("ContactNumber") = UserContactNumber()

    If Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("PR").Length > 0 Then vList("Product") = Request.QueryString("PR")
    Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebProducts, vList)
    Dim vBaseList As BaseDataList = TryCast(Me.FindControl("ProductData"), BaseDataList)
    Dim vImageName As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.web_product_image_name)
    Dim vSalePageNumber As String = ""
    If InitialParameters.ContainsKey("ProductSalePage") Then vSalePageNumber = InitialParameters("ProductSalePage").ToString
    If vImageName.Length = 0 Then vImageName = "Product{0}.png"
    If vBaseList IsNot Nothing Then
      DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebProducts, vBaseList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString), 0, False, vQueryString)
      If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
        Dim vDGR As DataGrid = CType(vBaseList, DataGrid)
        Dim vImagePos As Integer = -1
        Dim vSelectPos As Integer = -1
        Dim vRatePos As Integer = -1
        Dim vProductPos As Integer = -1
        If vSalePageNumber.Length > 0 AndAlso vDGR.Columns(0).HeaderText <> "Select" Then
          Dim vColumn As New TemplateColumn()
          vColumn.HeaderText = ""
          vDGR.Columns.AddAt(0, vColumn)
          vDGR.DataBind()
        End If
        For vCount As Integer = 0 To vDGR.Columns.Count - 1
          Dim vBoundColumn As TemplateColumn = DirectCast(vDGR.Columns(vCount), TemplateColumn)
          If vBoundColumn.HeaderText = "" Then
            vSelectPos = vCount
          ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "Product" Then
            vProductPos = vCount
          ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "Rate" Then
            vRatePos = vCount
          ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "ProductImage" Then
            vImagePos = vCount
          End If
        Next
        Dim vPath As String = ""

        For vRow As Integer = 0 To vDGR.Items.Count - 1
          If vImagePos >= 0 Then
            vPath = "Images/Products/" & vImageName.Replace("{0}", DirectCast(vDGR.Items(vRow).Cells(vProductPos).Controls(0), ITextControl).Text)
            'Call the GetImage which checks whether Image is available or not.
            vDGR.Items(vRow).Cells(vImagePos).Text = GetImage(vPath, DirectCast(vDGR.Items(vRow).Cells(vProductPos).Controls(0), ITextControl).Text, "Images/Products/Default.png", "ProductImage")
          End If
          If vSelectPos >= 0 Then
            If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
              vDGR.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & """ onclick=""location.href='Default.aspx?pn={0}&PR={1}&RA={2}'"">", vSalePageNumber, DirectCast(vDGR.Items(vRow).Cells(vProductPos).Controls(0), ITextControl).Text, DirectCast(vDGR.Items(vRow).Cells(vRatePos).Controls(0), ITextControl).Text)
            Else
              vDGR.Items(vRow).Cells(vSelectPos).Text = "<a href=default.aspx?pn=" & vSalePageNumber & "&PR=" & DirectCast(vDGR.Items(vRow).Cells(vProductPos).Controls(0), ITextControl).Text & "&RA=" & DirectCast(vDGR.Items(vRow).Cells(vRatePos).Controls(0), ITextControl).Text & ">" & mvHyperLinkText & "</a>"
            End If
          End If
        Next
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vSalePageNumber As String = InitialParameters.OptionalValue("ProductSalePage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)

      'Add the Select link at the end
      If vSalePageNumber.Length > 0 Then
        Dim vCount As Integer = e.Item.Controls.Count
        Dim vSelectLink As New Literal
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & " "" onclick=""location.href='Default.aspx?pn={0}&PR={1}&RA={2}'"">", vSalePageNumber, vDrv.Row("Product"), vDrv.Row("Rate"))
        Else
          vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&PR={1}&RA={2}'>" & mvHyperLinkText & "</a>", vSalePageNumber, vDrv.Row("Product"), vDrv.Row("Rate"))
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
      End If
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      FindProducts()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
End Class