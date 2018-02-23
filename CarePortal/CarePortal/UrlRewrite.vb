Public Class UrlRewrite
  Implements IHttpModule

  Public Sub Dispose() Implements System.Web.IHttpModule.Dispose
    '
  End Sub

  Public Sub Init(ByVal context As System.Web.HttpApplication) Implements System.Web.IHttpModule.Init
    AddHandler context.AuthorizeRequest, AddressOf OnAuthorizeRequest
  End Sub

  Private Sub OnAuthorizeRequest(ByVal sender As Object, ByVal e As EventArgs)
    Dim vApp As HttpApplication = CType(sender, HttpApplication)

    If vApp.Request.RawUrl.ToLower.EndsWith(".aspx") = True AndAlso
      vApp.Request.RawUrl.ToLower.Contains("returnpage.aspx") OrElse vApp.Request.RawUrl.CaseInsensitiveCompare("notification.aspx", StringComparison.CurrentCultureIgnoreCase) Then

    Else

      If vApp.Request.RawUrl.ToLower.EndsWith(".aspx") = True AndAlso vApp.Request.RawUrl.ToLower.Contains("default.aspx") = False AndAlso vApp.Request.RawUrl.ToLower.Contains("showerrors.aspx") = False Then
        'Retrieve the last part of the URL which should be the "friendly url" and find a Web Page with this value
        Dim vPos As Integer = vApp.Request.RawUrl.LastIndexOf("/")
        Dim vFriendlyUrl As String = ""
        If vPos >= 0 Then vFriendlyUrl = vApp.Request.RawUrl.Substring(vPos + 1)
        If vFriendlyUrl.Length > 0 Then
          Dim vList As New ParameterList(HttpContext.Current)
          vList("WebNumber") = System.Web.Configuration.WebConfigurationManager.AppSettings("WebNumber").ToString
          vList("FriendlyUrl") = vFriendlyUrl
          Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.SelectWebDataTable(CareNetServices.XMLWebDataSelectionTypes.wstPages, vList))      'Get the FriendlyUrl for all WebPages)
          If vRow IsNot Nothing Then
            Dim vQueryString As String = "pn=" & vRow.Item("WebPageNumber").ToString
            vApp.Context.RewritePath("Default.aspx", "", vQueryString)
          Else
            Throw New HttpException(404, "Page not found")
          End If
        End If
      End If
    End If
  End Sub
End Class
