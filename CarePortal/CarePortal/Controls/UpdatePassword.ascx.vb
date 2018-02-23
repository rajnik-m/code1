Partial Public Class UpdatePassword
  Inherits CareWebControl

  Private mvContactNumber As Integer
  Private mvUserName As String

  Public Sub New()
    mvNeedsAuthentication = False   'To account for being used as a reset password page
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    InitialiseControls(CareNetServices.WebControlTypes.wctUpdatePassword, tblDataEntry)
    mvContactNumber = 0
    Dim vEMail As String = Request.QueryString("UserName")
    If vEMail IsNot Nothing AndAlso vEMail.Length > 0 Then
      Dim vEP As New EncryptionProvider
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("UserName") = vEP.Decrypt(vEMail.Replace(vbCr, "").Replace(vbLf, ""))
      vParams("Password") = "Unknown"
      vParams("PasswordKnown") = "N"
      vParams("IsCurrent") = "N"          'Allow non current and locked out accounts to still update password
      'Check to see if user name is in use
      Dim vResultList As ParameterList = DataHelper.LoginRegisteredUser(vParams)
      If vResultList.Contains("Password") Then
        mvContactNumber = IntegerValue(vResultList("ContactNumber").ToString)
        mvUserName = vParams("UserName").ToString
      End If
    ElseIf Session("UserContactNumber") IsNot Nothing Then
      mvContactNumber = IntegerValue(Session("UserContactNumber").ToString)
      mvUserName = Session("RegisteredUserName").ToString
    End If

    If FindControlByName(Me, "SecurityQuestion") IsNot Nothing AndAlso mvContactNumber > 0 Then
      Dim vContactData As New DataTable()
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("ContactNumber") = mvContactNumber.ToString
      vParams("SystemColumns") = "N"
      vContactData = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers, vParams)
      If Not String.IsNullOrEmpty(vContactData.Rows(0).Item("SecurityQuestion").ToString) Then SetTextBoxText("SecurityQuestion", vContactData.Rows(0).Item("SecurityQuestion").ToString)
      If Not String.IsNullOrEmpty(vContactData.Rows(0).Item("SecurityAnswer").ToString) Then SetTextBoxText("SecurityAnswer", vContactData.Rows(0).Item("SecurityAnswer").ToString)
    End If
    If mvContactNumber <= 0 AndAlso Not InWebPageDesigner() Then ProcessRedirect(String.Format("default.aspx?pn={0}&ReturnURL={1}&Type={2}", LoginPageNumber, Server.UrlEncode(Request.Url.ToString), False))
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        If String.IsNullOrEmpty(mvUserName) Then Throw New Exception("Registered User is not logged in. Cannot change password")
        Dim vParams As New ParameterList(HttpContext.Current)
        vParams("OldUserName") = mvUserName
        Dim vValue As String = GetTextBoxText("NewPassword")
        vParams("Password") = vValue
        If Me.FindControl("SecurityQuestion") IsNot Nothing AndAlso Me.FindControl("SecurityQuestion").Visible Then vParams("SecurityQuestion") = GetTextBoxText("SecurityQuestion")
        If Me.FindControl("SecurityAnswer") IsNot Nothing AndAlso Me.FindControl("SecurityAnswer").Visible Then vParams("SecurityAnswer") = GetTextBoxText("SecurityAnswer")
        vParams("LockedOut") = ""
        vParams("LoginAttempts") = 0
        Dim vResultList As ParameterList = DataHelper.UpdateRegisteredUser(vParams)
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vEx As CareException
        If vEx.ErrorNumber = CareException.ErrorNumbers.enPasswordPreviouslyUsedInHistory Then
          If FindControlByName(tblDataEntry, "PageError") IsNot Nothing Then
            SetLabelText("PageError", vEx.Message)
            Me.FindControl("PageError").Visible = True
          Else
            'Can't find PageError label so must display error page
            ProcessError(vEx)
          End If
        Else
          ProcessError(vEx)
        End If
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

End Class