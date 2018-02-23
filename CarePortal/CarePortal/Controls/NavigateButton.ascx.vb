Partial Public Class NavigateButton
  Inherits CareWebControl

  Private mvPageNumber As Integer


  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctNavigateButton, tblDataEntry)
      If Request.QueryString("pn") IsNot Nothing AndAlso Request.QueryString("pn").Length > 0 Then mvPageNumber = IntegerValue(Request.QueryString("pn"))
      Dim vHyperLink As New Literal
      If InitialParameters.OptionalValue("HyperlinkFormat") = "H" Then
        vHyperLink.Text = String.Format("<a href='Default.aspx?pn={0}&fpn={1}'>" & DirectCast(Me.FindControl("Navigate"), Button).Text.ToString() & "</a>", mvSubmitItemNumber, mvPageNumber)
        Me.Controls.Add(vHyperLink)
        DirectCast(Me.FindControl("Navigate"), Button).Visible = False
      End If

      If InitialParameters.OptionalValue("AccessView").Length > 0 Then
        Dim vViews() As String
        vViews = FindViewsOfUser.Split(CChar(","))
        For Each vView As String In vViews
          If vView.ToUpper = InitialParameters.OptionalValue("AccessView").ToUpper Then
            If InitialParameters.OptionalValue("HyperlinkFormat") = "H" Then
              vHyperLink.Visible = True
            Else
              DirectCast(Me.FindControl("Navigate"), Button).Visible = True
            End If
            Exit For
          Else
            If InitialParameters.OptionalValue("HyperlinkFormat") = "H" Then
              If Not InWebPageDesigner() Then vHyperLink.Visible = False
            End If
            If Not InWebPageDesigner() Then DirectCast(Me.FindControl("Navigate"), Button).Visible = False
          End If
        Next
      End If

      If vHyperLink.Visible = False Then
        Me.Controls.Remove(tblDataEntry)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Function FindViewsOfUser() As String
    Dim vViews As String = String.Empty
    If HttpContext.Current.User.Identity.IsAuthenticated Then
      If Not TypeOf (HttpContext.Current.User.Identity) Is System.Security.Principal.WindowsIdentity Then
        Dim vIdentity As FormsIdentity = CType(HttpContext.Current.User.Identity, FormsIdentity)
        If vIdentity.Ticket.UserData.Length > 0 Then
          Dim vItems As String() = vIdentity.Ticket.UserData.Split("|"c)
          If vItems.Length > 4 Then 'Check if viewname exists in userdata
            'Split again to get a list of views that the user belongs to
            vViews = vItems(4).ToString()
          End If
        End If
      End If
    End If
    Return vViews
  End Function
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      GoToSubmitPage("&fpn=" & mvPageNumber.ToString)    'Include the PageNumber for the page we are coming from
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
End Class
