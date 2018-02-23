Partial Public Class EventSelection
  Inherits CareWebControl

  Private mvHyperLinkText As String = ""
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectEvents, tblDataEntry)
      If InitialParameters.ContainsKey("HyperlinkText") Then mvHyperLinkText = InitialParameters("HyperlinkText").ToString
      If Me.FindControl("SearchEvent") IsNot Nothing Then
        CType(Me.FindControl("SearchEvent"), TextBox).MaxLength = 100
        If Request.QueryString("Event") IsNot Nothing Then
          CType(Me.FindControl("SearchEvent"), TextBox).Text = Request.QueryString("Event")
        End If
      End If
      If Not IsPostBack Then FindEvents()
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

  Private Sub FindEvents()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vQueryString As String = ""
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    If GetTextBoxText("SearchEvent").Length > 0 Then
      vList("SearchEvent") = "*" & GetTextBoxText("SearchEvent") & "*"
      vQueryString = "&Event=" & GetTextBoxText("SearchEvent")
    End If
    Dim vBaseList As BaseDataList = TryCast(Me.FindControl("EventData"), BaseDataList)
    Dim vEditPageNumber As Integer = 0
    Dim vBookingPage As String = ""
    If InitialParameters.ContainsKey("Topic") Then vList("Topic") = InitialParameters("Topic").ToString
    If Request.QueryString("TO") IsNot Nothing AndAlso Request.QueryString("TO").Length > 0 Then vList("Topic") = Request.QueryString("TO")
    If Request.QueryString("EN") IsNot Nothing AndAlso Request.QueryString("EN").Length > 0 Then vList("EventNumber") = Request.QueryString("EN")
    If InitialParameters.ContainsKey("BookingPage") Then vEditPageNumber = IntegerValue(InitialParameters("BookingPage").ToString)
    If vBaseList IsNot Nothing Then
      Dim vRowCount As Integer = DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebEvents, vBaseList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString), vEditPageNumber, False, vQueryString)
      'Only for display grids. Data list select columns will be handled seperately
      If vRowCount > 0 Then
        If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
          Dim vDGR As DataGrid = CType(vBaseList, DataGrid)
          If InitialParameters.Contains("BookingPage") Then vBookingPage = InitialParameters("BookingPage").ToString
          Dim vEventPos As Integer
          Dim vSelectPos As Integer
          Dim vColumn As New TemplateColumn()
          'Book Event column.
          vColumn.HeaderText = ""
          vDGR.Columns.AddAt(0, vColumn)
          vDGR.DataBind()
          For vCount As Integer = 0 To vDGR.Columns.Count - 1
            Dim vBoundColumn As TemplateColumn = DirectCast(vDGR.Columns(vCount), TemplateColumn)
            If vBoundColumn.HeaderText = "" Then
              vSelectPos = vCount
            ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "EventNumber" Then
              vEventPos = vCount
            End If
          Next
          If vBookingPage.Length = 0 AndAlso vSelectPos >= 0 Then
            vDGR.Columns(vSelectPos).Visible = False
          Else
            For vRow As Integer = 0 To vDGR.Items.Count - 1
              If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                vDGR.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & """ onclick=""location.href='Default.aspx?pn={0}&EN={1}'"">", vBookingPage, DirectCast(vDGR.Items(vRow).Cells(vEventPos).Controls(0), ITextControl).Text)
              Else
                vDGR.Items(vRow).Cells(vSelectPos).Text = "<a href='default.aspx?pn=" & vBookingPage & "&EN=" & DirectCast(vDGR.Items(vRow).Cells(vEventPos).Controls(0), ITextControl).Text & "'>" & mvHyperLinkText & "</a>"
              End If
            Next
          End If
          'For displaying event image
          For vCount As Integer = 1 To vDGR.Columns.Count - 1
            Dim vBoundColumn As TemplateColumn = TryCast(vDGR.Columns(vCount), TemplateColumn)
            Dim vPath As String = ""
            If vBoundColumn IsNot Nothing AndAlso
               TypeOf vBoundColumn.ItemTemplate Is DisplayTemplate AndAlso
               DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "EventImage" Then
              For vRow As Integer = 0 To vDGR.Items.Count - 1
                vPath = "Images/Events/" & DirectCast(vDGR.Items(vRow).Cells(vCount).Controls(0), ITextControl).Text
                'Call the GetImage which checks whether Image is available or not.
                vDGR.Items(vRow).Cells(vCount).Text = GetImage(vPath, DirectCast(vDGR.Items(vRow).Cells(vCount).Controls(0), ITextControl).Text, "Images/Events/Default.png", "EventImage")
              Next
              Exit For
            End If
          Next
          vBaseList.Visible = True
        End If
        DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
      Else
        If GetTextBoxText("SearchEvent").Length > 0 Then
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        Else
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
          DirectCast(Me.FindControl("SearchEvent"), TextBox).Visible = False
          DirectCast(Me.FindControl("SearchEvent"), TextBox).Parent.Parent.Visible = False
          DirectCast(Me.FindControl("Search"), Button).Visible = False
        End If
        DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        vBaseList.Visible = False
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vBookingPage As String = InitialParameters.OptionalValue("BookingPage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)

      'Add the book event link at the end
      If vBookingPage.Length > 0 Then
        Dim vCount As Integer = e.Item.Controls.Count
        Dim vBookEventLink As New Literal
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vBookEventLink.Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & """ onclick=""location.href='Default.aspx?pn={0}&EN={1}'"">", vBookingPage, vDrv.Row("EventNumber"))
        Else
          vBookEventLink.Text = String.Format("<a href='Default.aspx?pn={0}&EN={1}'>" & mvHyperLinkText & "</a>", vBookingPage, vDrv.Row("EventNumber"))
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vBookEventLink)
      End If
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        FindEvents()
      Catch vEx As ThreadAbortException
        Throw vEx
      End Try
    End If
  End Sub

End Class
