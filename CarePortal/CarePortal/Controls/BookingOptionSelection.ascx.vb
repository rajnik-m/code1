Public Class BookingOptionSelection
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectBookingOptions, tblDataEntry)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      FindBookingOption(vList, False)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub FindBookingOption(ByVal pList As ParameterList, ByVal pShowCount As Boolean)
    Dim vBookingPage As String = ""
    Dim vEventNumber As String = ""
    If InitialParameters.Contains("BookingPage") Then vBookingPage = InitialParameters("BookingPage").ToString
    If InitialParameters.ContainsKey("EventNumber") Then
      pList("EventNumber") = InitialParameters("EventNumber").ToString
      vEventNumber = pList("EventNumber").ToString
    Else
      If InWebPageDesigner() Then pList("DocumentColumns") = "Y"
    End If
    If UserContactNumber() > 0 Then pList("ContactNumber") = UserContactNumber()
    If Request.QueryString("EN") IsNot Nothing Then
      pList("EventNumber") = Request.QueryString("EN")
      vEventNumber = pList("EventNumber").ToString
    End If
    If vEventNumber.Length = 0 Then pList("DocumentColumns") = "Y"

    Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebBookingOptions, pList)
    Dim vBaseList As BaseDataList = TryCast(Me.FindControl("BookingOption"), BaseDataList)
    Dim vDGR As DataGrid = Nothing
    If vBaseList IsNot Nothing Then
      SetControlVisible("WarningMessage", False)
      DataHelper.FillGrid(vResult, vBaseList)
      If vBookingPage.Length > 0 Then
        If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
          Dim vColumn As New BoundColumn()
          vDGR = CType(vBaseList, DataGrid)
          vColumn.HeaderText = ""
          vDGR.Columns.AddAt(0, vColumn)
          vDGR.DataBind()
          Dim vOptionPos As Integer = GetDataGridItemIndex(vDGR, "OptionNumber")  'Get the OptionNumber position by looking at the column name and not the header text
          Dim vRatePos As Integer = GetDataGridItemIndex(vDGR, "Rate")  'Get the Rate position by looking at the column name and not the header text
          Dim vSelectPos As Integer = GetDataGridItemIndex(vDGR, "")
          If vDGR.Items.Count = 0 Then
            SetControlVisible("WarningMessage", True)
            vDGR.Visible = False
          ElseIf vDGR.Items.Count = 1 AndAlso vEventNumber.Length > 0 AndAlso Not InWebPageDesigner() Then
            ProcessRedirect("default.aspx?pn=" & vBookingPage & "&EN=" & vEventNumber.ToString & "&OP=" & DirectCast(vDGR.Items(0).Cells(vOptionPos).Controls(0), ITextControl).Text & "&RA=" & DirectCast(vDGR.Items(0).Cells(vRatePos).Controls(0), ITextControl).Text)
          Else
            For vRow As Integer = 0 To vDGR.Items.Count - 1
              If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                vDGR.Items(vRow).Cells(vSelectPos).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&EN={2}&OP={3}&RA={4}'"">", InitialParameters("HyperlinkText").ToString, vBookingPage, vEventNumber.ToString, DirectCast(vDGR.Items(vRow).Cells(vOptionPos).Controls(0), ITextControl).Text, DirectCast(vDGR.Items(vRow).Cells(vRatePos).Controls(0), ITextControl).Text)
              Else
                vDGR.Items(vRow).Cells(vSelectPos).Text = "<a href='default.aspx?pn=" & vBookingPage & "&EN=" & vEventNumber.ToString & "&OP=" & DirectCast(vDGR.Items(vRow).Cells(vOptionPos).Controls(0), ITextControl).Text & "&RA=" & DirectCast(vDGR.Items(vRow).Cells(vRatePos).Controls(0), ITextControl).Text & "'>" & InitialParameters("HyperlinkText").ToString & "</a>"
              End If
            Next
          End If
        Else
          Dim vDataList As DataList = CType(vBaseList, DataList)
          If vDataList.Items.Count = 0 Then
            SetControlVisible("WarningMessage", True)
            vDataList.Visible = False
          ElseIf vDataList.Items.Count = 1 And Not InWebPageDesigner() Then
            Dim vDataSet As DataSet = TryCast(vDataList.DataSource, DataSet)
            ProcessRedirect(String.Format("Default.aspx?pn={0}&EN={1}&OP={2}&RA={3}", vBookingPage, vEventNumber, vDataSet.Tables("DataRow").Rows(0)("OptionNumber"), vDataSet.Tables("DataRow").Rows(0)("Rate")))
          End If
        End If
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vBookingPage As String = InitialParameters.OptionalValue("BookingPage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)

      'Add the Select link at the end
      If vBookingPage.Length > 0 Then
        Dim vEventNumber As String = String.Empty
        If InitialParameters.ContainsKey("EventNumber") Then
          vEventNumber = InitialParameters("EventNumber").ToString
        ElseIf Request.QueryString("EN") IsNot Nothing Then
          vEventNumber = Request.QueryString("EN")
        End If

        Dim vCount As Integer = e.Item.Controls.Count
        Dim vSelectLink As New Literal
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&EN={2}&OP={3}&RA={4}'"">", InitialParameters("HyperlinkText").ToString, vBookingPage, vEventNumber, vDrv.Row("OptionNumber"), vDrv.Row("Rate"))
        Else
          vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&EN={1}&OP={2}&RA={3}'>{4}</a>", vBookingPage, vEventNumber, vDrv.Row("OptionNumber"), vDrv.Row("Rate"), InitialParameters("HyperlinkText").ToString)
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
      End If
    End If
  End Sub
End Class