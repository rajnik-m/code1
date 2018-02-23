Partial Public Class Logout
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctLogout, tblDataEntry)
      Dim vCookie As HttpCookie = Response.Cookies(FormsAuthentication.FormsCookieName)
      If vCookie IsNot Nothing Then
        vCookie.Expires = DateTime.Now.AddDays(-1)
        Response.SetCookie(vCookie)
        Session.Abandon()

        'Single Sign-Off
        Dim vSingleSignOffURL As String = GetCustomConfigItem("CustomConfiguration/SingleSignOffURL", True)
        If vSingleSignOffURL.Length > 0 Then
          Dim vReturnURL As String = Request.QueryString("ReturnURL")
          If vReturnURL Is Nothing OrElse vReturnURL.Length = 0 Then
            If SubmitItemUrl.Length > 0 Then
              vReturnURL = SubmitItemUrl
            ElseIf SubmitItemNumber > 0 Then
              vReturnURL = String.Format("{0}?pn={1}", If(String.IsNullOrEmpty(Request.Url.Query), Request.Url.AbsoluteUri, Request.Url.AbsoluteUri.Replace(Request.Url.Query, "")), SubmitItemNumber)
            End If
          End If
          'BR18438 
          'Debug.WriteLine("PRE BR18438 fix: Logout vReturnURL: " & vReturnURL)
          vReturnURL = Server.UrlEncode(vReturnURL)
          'BR18438
          'Debug.WriteLine("POST BR18438 fix: Logout vReturnURL: " & vReturnURL)
          RedirectViaWhiteList(String.Format(vSingleSignOffURL, vReturnURL))
        End If
      End If
      If Not InWebPageDesigner() Then
        GoToSubmitPage()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
End Class