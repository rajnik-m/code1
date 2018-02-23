Public Partial Class EMailPassword
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctForgottenPassword, tblDataEntry)
    SetControlVisible("WarningMessage1", False)
    SetParentParentVisible("SecurityQuestion", False)
    SetParentParentVisible("SecurityAnswer", False)
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vParams As New ParameterList(HttpContext.Current)
        AddOptionalTextBoxValue(vParams, "UserName")
        vParams("Password") = "Unknown"
        vParams("PasswordKnown") = "N"
        vParams("IsCurrent") = "N"          'Allow non current and locked out accounts to still get a password update email
        'Check to see if user name is in use
        Dim vResultList As ParameterList = DataHelper.LoginRegisteredUser(vParams)
        If vResultList.Contains("Password") Then
          Dim vQuestion As String = vResultList("SecurityQuestion").ToString
          Dim vAnswer As String = vResultList("SecurityAnswer").ToString
          If GetTextBoxText("SecurityQuestion").Length = 0 Then
            If vResultList("SecurityQuestion").ToString.Length = 0 Then
              SendEmailWithPassword(vResultList("EmailAddress").ToString)
            ElseIf vResultList("SecurityQuestion").ToString.Length <> 0 Then
              SetControlVisible("Message", False)
              SetParentParentVisible("UserName", False)
              SetParentParentVisible("SecurityQuestion", True)
              DirectCast(FindControlByName(Me, "SecurityQuestion"), TextBox).ReadOnly = True
              SetTextBoxText("SecurityQuestion", vResultList("SecurityQuestion").ToString)
              SetParentParentVisible("SecurityAnswer", True)
            End If
          ElseIf GetTextBoxText("SecurityQuestion").Length <> 0 Then
            If GetTextBoxText("SecurityQuestion") = vQuestion AndAlso GetTextBoxText("SecurityAnswer") = vAnswer Then
              SendEmailWithPassword(vResultList("EmailAddress").ToString)
            Else
              'Failed to answer security question redirect to login
              ProcessRedirect(String.Format("Default.aspx?pn={0}", LoginPageNumber.ToString))
            End If
          End If
        Else
          SetLabelTextFromLabel("Message", "WarningMessage1", "Invalid User Name. Please try again")
        End If
      Catch vCareEx As CareException
        Select Case vCareEx.ErrorNumber
          Case CareException.ErrorNumbers.enUserDoesNotExist
            SetLabelTextFromLabel("Message", "WarningMessage1", "Invalid User Name. Please try again")
        End Select
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub SendEmailWithPassword(pEmailAddress As String)
    Dim vContentParams As New ParameterList
    If DefaultParameters.ContainsKey("ResetPasswordPage") Then
      Dim vRegLink As New StringBuilder
      vRegLink.Append(New UriBuilder(Request.Url.Scheme, Request.Url.Host, Request.Url.Port, Request.Url.LocalPath).Uri.AbsoluteUri)
      vRegLink.Append("?pn=")
      vRegLink.Append(DefaultParameters("ResetPasswordPage").ToString)
      vRegLink.Append("&UserName=")
      Dim vEP As New EncryptionProvider
      vRegLink.Append(vEP.Encrypt(GetTextBoxText("UserName")))
      vContentParams("ResetPasswordLink") = vRegLink.ToString
    End If
    vContentParams("EMail") = pEmailAddress
    'BR19442 removed password
    Dim vEmailParams As New ParameterList(HttpContext.Current)
    vEmailParams("StandardDocument") = DefaultParameters("StandardDocument")
    vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
    vEmailParams("Name") = DefaultParameters("Name")
    DataHelper.ProcessBulkEMail(vContentParams.ToCSVFile, vEmailParams, True)
    GoToSubmitPage()
  End Sub

End Class