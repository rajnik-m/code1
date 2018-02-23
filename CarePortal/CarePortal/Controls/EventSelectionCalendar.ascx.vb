Partial Public Class EventSelectionCalendar
  Inherits CareWebControl
  Dim mvTable As DataTable

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If Not IsPostBack Then
      Dim vStartDate As DateTime = New DateTime(dateTime.Today.Year, dateTime.Today.Month, 1)
      Dim vList As New ParameterList(HttpContext.Current)
      If InitialParameters("Topic") IsNot Nothing Then
        vList("Topic") = InitialParameters("Topic").ToString
      End If
      vList("StartDate") = vStartDate
      vList("EndDate") = vStartDate.AddMonths(1)
      mvTable = GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebEvents, vList))
    End If
  End Sub

  Protected Sub calEventCalendar_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles calEventCalendar.DayRender
    If mvTable IsNot Nothing Then
      For Each vDr As DataRow In mvTable.Rows
        If CDate(vDr("StartDate")) = CDate(e.Day.Date) Then
          Dim vHyperLink As New HyperLink
          vHyperLink.Text = vDr("EventDesc").ToString
          If InitialParameters("BookingPage") IsNot Nothing AndAlso Not String.IsNullOrEmpty(InitialParameters("BookingPage").ToString) Then
            vHyperLink.NavigateUrl = String.Format("default.aspx?pn={0}&EN={1}", InitialParameters("BookingPage"), vDr("EventNumber").ToString) 'String.Format("default.aspx?pn={0}&EV={1}", PageNumber, EventNumber)
          End If
          e.Cell.Controls.Add(New LiteralControl("<br />"))
          e.Cell.Controls.Add(vHyperLink)
        End If
      Next
    End If
  End Sub

  Protected Sub calEventCalendar_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles calEventCalendar.VisibleMonthChanged
    Dim vStartDate As DateTime = New DateTime(e.NewDate.Year, e.NewDate.Month, 1)
    Dim vList As New ParameterList(HttpContext.Current)
    If InitialParameters("Topic") IsNot Nothing Then
      vList("Topic") = InitialParameters("Topic").ToString
    End If
    vList("StartDate") = vStartDate
    vList("EndDate") = vStartDate.AddMonths(1)
    mvTable = GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebEvents, vList))
  End Sub
End Class