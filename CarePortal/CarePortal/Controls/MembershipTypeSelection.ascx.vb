Public Class MembershipTypeSelection
  Inherits CareWebControl

  Private mvHyperLink1 As String = ""
  Private mvHyperLink2 As String = ""
  Private mvString As String = String.Empty
  Private mvLocationLink As String = String.Empty

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctSelectMembershipTypes, tblDataEntry)
    If InitialParameters.ContainsKey("HyperlinkText1") Then mvHyperLink1 = InitialParameters("HyperlinkText1").ToString
    If InitialParameters.ContainsKey("HyperlinkText2") Then mvHyperLink2 = InitialParameters("HyperlinkText2").ToString
    FindMembershipTypes()
  End Sub

  Private Sub FindMembershipTypes()
    Dim vContactNumber As Integer = GetContactNumberFromParentGroup()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber

    If vContactNumber = 0 Then
      If Session.Contents.Item("UserContactNumber") IsNot Nothing AndAlso IntegerValue(Session("UserContactNumber").ToString) > 0 Then
        vList("UserContactNumber") = IntegerValue(Session("UserContactNumber").ToString)
      End If
    Else
      vList("UserContactNumber") = vContactNumber
    End If

    Dim vStartDate As String = GetMembershipStartDate()

    If InitialParameters.ContainsKey("LookupGroup") Then vList("LookupGroup") = InitialParameters("LookupGroup").ToString
    Dim vBaseList As BaseDataList = CType(FindControlByName(tblDataEntry, "MembershipTypeData"), BaseDataList)
    DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebMembershipTypes, vBaseList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString))
    'Only for display grids. Data list select column will be handled seperately
    If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
      Dim vDataGrid As DataGrid = CType(vBaseList, DataGrid)
      Dim vDirectDebitPageNumber As String = InitialParameters.OptionalValue("DirectDebitSalePage").ToString
      Dim vCreditCardPageNumber As String = InitialParameters.OptionalValue("CreditCardSalePage").ToString
      Dim vDataSet As DataSet = TryCast(vDataGrid.DataSource, DataSet)
      'if the membership col is found then create a hyperlink to the sale page
      If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then

        If vDirectDebitPageNumber.Length > 0 AndAlso vCreditCardPageNumber.Length > 0 Then
          If vDataGrid.Columns.Count = vDataSet.Tables("DataRow").Columns.Count Then
            Dim vDebitColumn As New BoundColumn()
            vDataGrid.Columns.AddAt(0, vDebitColumn)
            Dim vCreditColumn As New BoundColumn()
            vDataGrid.Columns.AddAt(1, vCreditColumn)
          End If
        Else
          If vDataGrid.Columns.Count = vDataSet.Tables("DataRow").Columns.Count Then
            If vDirectDebitPageNumber.Length > 0 Then
              Dim vColumn As New BoundColumn()
              'vColumn.HeaderText = "Credit Card"
              vDataGrid.Columns.AddAt(0, vColumn)
            ElseIf vCreditCardPageNumber.Length > 0 Then
              Dim vColumn As New BoundColumn()
              'vColumn.HeaderText = "Credit Card"
              vDataGrid.Columns.AddAt(0, vColumn)
            End If
          End If
        End If

        vBaseList.DataBind()

        For vRow As Integer = 0 To vDataGrid.Items.Count - 1
          If vDirectDebitPageNumber.Length > 0 Then
            If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink1 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}&SD={2}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink1 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            Else
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}&SD={2}'>" & mvHyperLink1 & "</a>&nbsp;&nbsp;&nbsp;", vDirectDebitPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}'>" & mvHyperLink1 & "</a>&nbsp;&nbsp;&nbsp;", vDirectDebitPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            End If
          End If
          If vDirectDebitPageNumber.Length = 0 AndAlso vCreditCardPageNumber.Length > 0 Then
            If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}&SD={2}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            Else
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}&SD={2}'>" & mvHyperLink2 & "</a>&nbsp;&nbsp;&nbsp;", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}'>" & mvHyperLink2 & "</a>&nbsp;&nbsp;&nbsp;", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            End If
          ElseIf vCreditCardPageNumber.Length > 0 Then
            If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(1).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}&SD={2}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(1).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}'"">", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            Else
              If vStartDate.Length > 0 Then
                vDataGrid.Items(vRow).Cells(1).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}&SD={2}'>" & mvHyperLink2 & "</a>&nbsp;&nbsp;&nbsp;", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"), vStartDate)
              Else
                vDataGrid.Items(vRow).Cells(1).Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}'>" & mvHyperLink2 & "</a>&nbsp;&nbsp;&nbsp;", vCreditCardPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("MembershipType"))
              End If
            End If
          End If
        Next
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vDirectDebitPageNumber As String = InitialParameters.OptionalValue("DirectDebitSalePage").ToString
      Dim vCreditCardPageNumber As String = InitialParameters.OptionalValue("CreditCardSalePage").ToString
      'Add a select link at the end
      Dim vCount As Integer = e.Item.Controls.Count
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)

      Dim vStartDate As String = GetMembershipStartDate()

      If vDirectDebitPageNumber.Length > 0 Then
        Dim vSelectLink As New Literal
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          If vStartDate.Length > 0 Then
            vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" runat=""server"" value=""" & mvHyperLink1 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}&SD={2}'"">", vDirectDebitPageNumber, vDrv.Row("MembershipType"), vStartDate)
          Else
            vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" runat=""server"" value=""" & mvHyperLink1 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}'"">", vDirectDebitPageNumber, vDrv.Row("MembershipType"))
          End If
        Else
          If vStartDate.Length > 0 Then
            vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}&SD={2}'>" & mvHyperLink1 & "</a>", vDirectDebitPageNumber, vDrv.Row("MembershipType"), vStartDate)
          Else
            vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&MT={1}'>" & mvHyperLink1 & "</a>", vDirectDebitPageNumber, vDrv.Row("MembershipType"))
          End If

        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
      End If

      If vCreditCardPageNumber.Length > 0 Then
        Dim vSelectLink As New Literal

        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          If vStartDate.Length > 0 Then
            vSelectLink.Text = String.Format("&nbsp;&nbsp;&nbsp;<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}&SD={2}'"">", vCreditCardPageNumber, vDrv.Row("MembershipType"), vStartDate)
          Else
            vSelectLink.Text = String.Format("&nbsp;&nbsp;&nbsp;<input type=""button"" class=""Button"" value=""" & mvHyperLink2 & """ onclick=""location.href='Default.aspx?pn={0}&MT={1}'"">", vCreditCardPageNumber, vDrv.Row("MembershipType"))
          End If
        Else
          If vStartDate.Length > 0 Then
            vSelectLink.Text = String.Format("&nbsp;&nbsp;&nbsp;<a href='Default.aspx?pn={0}&MT={1}&SD={2}'>" & mvHyperLink2 & " </a>", vCreditCardPageNumber, vDrv.Row("MembershipType"), vStartDate)
          Else
            vSelectLink.Text = String.Format("&nbsp;&nbsp;&nbsp;<a href='Default.aspx?pn={0}&MT={1}'>" & mvHyperLink2 & " </a>", vCreditCardPageNumber, vDrv.Row("MembershipType"))
          End If
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
      End If
    End If
  End Sub

  Private Function GetMembershipStartDate() As String
    Dim vResult As String = String.Empty

    If FindControlByName(tblDataEntry, "StartDateList") IsNot Nothing AndAlso ParentGroup.Length > 0 Then
      vResult = TryCast(FindControlByName(tblDataEntry, "StartDateList"), DropDownList).SelectedValue
    Else
      vResult = Date.Now.ToShortDateString
    End If
    Return vResult
  End Function

  Public Overrides Sub MembershipStartDateChangeHandler()
    FindMembershipTypes()
  End Sub
End Class