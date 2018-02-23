Public Partial Class DatePicker
    Inherits System.Web.UI.Page

  Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    If Not IsPostBack Then
      Dim vValue As String = Request.QueryString("Value")
      If vValue <> "" Then
        If IsDate(vValue) Then
          calCalendar.SelectedDate = CDate(vValue)
          calCalendar.VisibleDate = CDate(vValue)
        End If
      End If
    End If
  End Sub

  Private Sub Calendar_DayRender(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles calCalendar.DayRender
    Dim hl As New HyperLink
    hl.Text = CType(e.Cell.Controls(0), LiteralControl).Text
    hl.NavigateUrl = "javascript:SetDate('" & e.Day.Date.ToShortDateString & "')"
    e.Cell.Controls.Clear()
    e.Cell.Controls.Add(hl)
  End Sub

End Class