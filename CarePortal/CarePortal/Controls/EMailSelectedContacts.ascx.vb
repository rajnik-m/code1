Public Class EMailSelectedContacts
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vSendEmailLink As New Literal
      InitialiseControls(CareNetServices.WebControlTypes.wctEmailSelectedContacts, tblDataEntry)
      If Not InWebPageDesigner() AndAlso Session("SelectedEmailAddresses") Is Nothing Then
        Throw New Exception("Email Addresses not found")
      End If
      If InitialParameters.OptionalValue("HyperlinkText").Length > 0 Then
        Dim sb As StringBuilder = New StringBuilder("<a href='mailto:")
        If Not InWebPageDesigner() Then sb.Append(Session("SelectedEmailAddresses").ToString)
        sb.Append("'>")
        sb.Append(InitialParameters("HyperlinkText").ToString)
        sb.Append("</a>")
        vSendEmailLink.Text = sb.ToString()
        Me.Controls.AddAt(0, vSendEmailLink)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not InWebPageDesigner() Then
      GoToSubmitPage()
    End If
  End Sub
End Class