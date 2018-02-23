Public Class IMAPEMail
  Inherits EmailInterface

  Public Overrides Function CanEMail(Optional ByVal pDownloadMail As Boolean = False) As Boolean

  End Function

  Public Overrides Function GetAttachmentPathName(ByVal pMsgID As String, ByVal pIndex As Integer) As String
    Return ""
  End Function

  Public Overrides Function GetInBox() As System.Collections.ArrayList
    Return Nothing
  End Function

  Public Overrides Sub MarkRead(ByVal pMsg As EMailMessage)

  End Sub

  Public Overloads Overrides Function ProcessAction(ByVal pMsg As EMailMessage, ByVal pAction As EmailInterface.EMailActions) As Boolean

  End Function

  Public Overrides Function SendMail(ByVal pForm As System.Windows.Forms.Form, ByVal pOptions As EmailInterface.SendEmailOptions, ByVal pSubject As String, ByVal pMessage As String, ByVal pEmailAddress As String, Optional ByVal pAttachments As Microsoft.VisualBasic.Collection = Nothing, Optional ByVal pCCList As String = "") As Boolean

  End Function

  Public Overrides Function SendHtmlEmail(pMailMessage As Net.Mail.MailMessage) As Boolean
    Throw New NotSupportedException("Not yet implemented")
  End Function
End Class
