Public Class DisplayRelatedOrganisations
  Inherits CareWebControl
  Dim mvOrganisationView As String
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If Request.QueryString("MROP") Is Nothing Then
        InitialiseControls(CareNetServices.WebControlTypes.wctRelatedOrganisations, tblDataEntry, "", "")
        If Me.FindControl("WarningMessage") IsNot Nothing Then Me.FindControl("WarningMessage").Visible = False
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ViewType") = "X"

        Dim vOrganisationViewTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAllViews, vList)
        If vOrganisationViewTable.Rows.Count > 0 Then
          mvOrganisationView = vOrganisationViewTable.Rows(0).Item("ViewName").ToString
          vList = New ParameterList(HttpContext.Current)
          If InitialParameters.Contains("ShowChildren") Then vList("ShowChildren") = InitialParameters("ShowChildren").ToString
          If Request.QueryString("ON") IsNot Nothing Then vList("OrganisationNumber") = Request.QueryString("ON")
          vList("ContactNumber") = UserContactNumber()
          vList("WebPageItemNumber") = Me.WebPageItemNumber
          vList("SystemColumns") = "Y"
          Dim vOrganisationDataList As BaseDataList
          Dim vCount As Long
          If FindControlByName(Me, "OrganisationData") IsNot Nothing Then
            vOrganisationDataList = CType(FindControlByName(Me, "OrganisationData"), BaseDataList)
            vCount = DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftRelatedOrganisations, vOrganisationDataList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString), , False, )
            If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
              Dim vDataGrid As DataGrid = CType(vOrganisationDataList, DataGrid)
              Dim vOrganisationNumberPos As Integer
              Dim vChildCountPos As Integer
              Dim vSelectColumn As New BoundColumn()
              Dim vExpandColumn As New BoundColumn()
              Dim vOrgNo As String
              Dim vChildCount As Integer
              Dim vUrlText As String = ""

              vDataGrid.Columns.AddAt(0, vSelectColumn)
              vDataGrid.Columns.AddAt(1, vExpandColumn)
              vDataGrid.DataBind()
              For vColCount As Integer = 0 To vDataGrid.Columns.Count - 1
                Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vColCount), BoundColumn)
                If vBoundColumn.DataField = "OrganisationNumber" Then
                  vOrganisationNumberPos = vColCount
                ElseIf vBoundColumn.DataField = "ChildCount" Then
                  vChildCountPos = vColCount
                End If
              Next
              If InitialParameters.Contains("ManageRelatedOrganisationPage") Then
                For vRow As Integer = 0 To vDataGrid.Items.Count - 1
                  vOrgNo = vDataGrid.Items(vRow).Cells(vOrganisationNumberPos).Text
                  If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                    vDataGrid.Items(vRow).Cells(0).Text = String.Format("&nbsp;&nbsp;<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&ON={2}&MROP=1'"">", InitialParameters("HyperlinkText1").ToString, WebPageNumber, vOrgNo)
                  Else
                    vDataGrid.Items(vRow).Cells(0).Text = "<a href='default.aspx?pn=" & WebPageNumber & "&ON=" & vOrgNo & "&MROP=1'>" & InitialParameters("HyperlinkText1").ToString & "</a>&nbsp &nbsp"
                  End If
                Next
              Else
                vDataGrid.Columns(0).Visible = False
              End If
              If InitialParameters("ShowChildren").ToString = "C" Then
                For vRow As Integer = 0 To vDataGrid.Items.Count - 1
                  vOrgNo = vDataGrid.Items(vRow).Cells(vOrganisationNumberPos).Text
                  vChildCount = CType(vDataGrid.Items(vRow).Cells(vChildCountPos).Text, Integer)
                  If vChildCount > 0 Then
                    If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                      vDataGrid.Items(vRow).Cells(1).Text = String.Format("&nbsp;&nbsp;<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&ON={2}'"">", InitialParameters("HyperlinkText2").ToString, WebPageNumber, vOrgNo)
                    Else
                      vDataGrid.Items(vRow).Cells(1).Text = "<a href='default.aspx?pn=" & WebPageNumber & "&ON=" & vOrgNo & "'>" & InitialParameters("HyperlinkText2").ToString & "</a>&nbsp &nbsp"
                    End If
                  End If
                Next
              Else
                vDataGrid.Columns(1).Visible = False
              End If
            End If
          End If
        Else
          If Me.FindControl("WarningMessage") IsNot Nothing Then Me.FindControl("WarningMessage").Visible = True
        End If
      ElseIf Request.QueryString("MROP") IsNot Nothing AndAlso Request.QueryString("MROP") IsNot Nothing AndAlso InitialParameters.Contains("ManageRelatedOrganisationPage") Then
        Session("SelectedOrganisationNumber") = Request.QueryString("ON")
        ProcessRedirect("default.aspx?pn=" & InitialParameters("ManageRelatedOrganisationPage").ToString)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vManageRelatedOrganisationPage As String = InitialParameters.OptionalValue("ManageRelatedOrganisationPage").ToString
      Dim vCount As Integer = e.Item.Controls.Count
      Dim vSelectLink As New Literal
      Dim vExpandLink As New Literal
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)
      If vManageRelatedOrganisationPage.Length > 0 Then
        'Add a select link at the end
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&ON={2}&MROP=1'"">&nbsp &nbsp", InitialParameters("HyperlinkText1").ToString, WebPageNumber, vDrv.Row("OrganisationNumber"))
        Else
          vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&ON={1}&MROP=1'>{2}</a>&nbsp &nbsp", WebPageNumber, vDrv.Row("OrganisationNumber"), InitialParameters("HyperlinkText1").ToString)
        End If
      End If
      If InitialParameters("ShowChildren").ToString = "C" Then
        Dim vChildCount As Integer = CType(vDrv.Row("ChildCount"), Integer)
        If vChildCount > 0 Then
          If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
            vExpandLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&ON={2}'"">&nbsp &nbsp", InitialParameters("HyperlinkText2").ToString, WebPageNumber, vDrv.Row("OrganisationNumber"))
          Else
            vExpandLink.Text = String.Format("<a href='Default.aspx?pn={0}&ON={1}'>{2}</a>&nbsp &nbsp", WebPageNumber, vDrv.Row("OrganisationNumber"), InitialParameters("HyperlinkText2").ToString)
          End If
        End If
      End If
      If vCount > 0 Then
        e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
        e.Item.Controls(vCount - 1).Parent.Controls.Add(vExpandLink)
      End If
    End If
  End Sub
End Class