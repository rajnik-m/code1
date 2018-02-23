Partial Public Class Login
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctLogin, tblDataEntry)
    If FindControlByName(Me, "Message") IsNot Nothing Then FindControlByName(Me, "Message").Visible = False
    Dim vMemberNumber As String
    If Request.Form.Item("MemberNumber") IsNot Nothing Then
      vMemberNumber = Request.Form.Item("MemberNumber")
    Else
      vMemberNumber = String.Empty
    End If
    If Request.Form.Item("UserName") IsNot Nothing AndAlso Request.Form.Item("Password") IsNot Nothing Then
      ProcessLogin(Request.Form.Item("UserName"), Request.Form.Item("Password"), vMemberNumber)
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Me.Page.IsValid Then
      ProcessLogin(GetTextBoxText("UserName"), GetTextBoxText("Password"), GetTextBoxText("MemberNumber"))
    End If
  End Sub

  Private Sub ProcessLogin(ByVal pUserName As String, ByVal pPassword As String, ByVal pMemberNumber As String)
    Try
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("UserName") = pUserName
      vParams("Password") = pPassword
      If Not String.IsNullOrEmpty(pMemberNumber) Then
        vParams("MemberNumber") = pMemberNumber
      End If
      Dim vList As ParameterList
      If Request.QueryString("Type") = "True" Then
        vList = DataHelper.Login(vParams)
      Else
        vList = DataHelper.LoginRegisteredUser(vParams)
        'In Case both User name and Member Number are enabled and submitted then check username is  valid
        If Not String.IsNullOrEmpty(pMemberNumber) And Not String.IsNullOrEmpty(pUserName) Then
          If vList("UserName").ToString <> vParams("UserName").ToString Then Throw New Exception("")
        End If
        Session("RegisteredUserName") = vList("UserName")
      End If
      If vList.Contains("ContactNumber") Then Session("UserContactNumber") = vList("ContactNumber").ToString
      SetAuthentication(vList)

      'Check for Single Sign On
      Dim vProcessSingleSignOn As Boolean = GetCustomConfigItem("CustomConfiguration/SingleSignOnURL", True).Length > 0 AndAlso GetCustomConfigItem("CustomConfiguration/SingleSignOnKey").Length > 0

      'if update details page is set by user
      If UpdateDetailsPageNumber > 0 Then
        'Checking update frequency config option (by default will be 0) 
        Dim vUpdateFrequency As Integer = IntegerValue(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_update_details_freq))
        Dim vLastUpdate As Integer

        'if last details updated date is not set or null then set flag to zero
        'else set value to the no. of days since the last update was done
        If String.IsNullOrEmpty(vList("LastUpdatedOn").ToString) Then
          vLastUpdate = -1
        Else
          vLastUpdate = CInt(DateDiff(DateInterval.Day, CDate(vList("LastUpdatedOn")), DateTime.Now))
        End If
        If vUpdateFrequency = 1 AndAlso vLastUpdate <> -1 Then
          'do nothing
          'Update Details are their so no details required
          'Update only once
        ElseIf vLastUpdate = -1 OrElse vLastUpdate > vUpdateFrequency OrElse vUpdateFrequency = 0 Then
          'if return url is their then paas it on to udate details page
          If Request.QueryString("ReturnURL") Is Nothing Then
            Session("UpdateDetailsURL") = String.Format("Default.aspx?pn={0}", UpdateDetailsPageNumber.ToString)
          Else
            Session("UpdateDetailsURL") = String.Format("Default.aspx?pn={0}&ReturnURL={1}", UpdateDetailsPageNumber.ToString, Server.UrlEncode(Request.QueryString("ReturnURL").ToString))
          End If
          If vProcessSingleSignOn = False Then ProcessRedirect(Session("UpdateDetailsURL").ToString)
        End If
      End If

      If vProcessSingleSignOn Then ProcessSingleSignOn(vList)

      Dim vReturnURL As String = Request.Params("ReturnURL")
      If vReturnURL Is Nothing Then
        GoToSubmitPage()
      Else
        RedirectViaWhiteList(vReturnURL)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      'Default message: Failed to Login. Please Check Login Credentials then try again
      If FindControlByName(Me, "Message") IsNot Nothing Then FindControlByName(Me, "Message").Visible = True
    End Try
  End Sub

End Class
