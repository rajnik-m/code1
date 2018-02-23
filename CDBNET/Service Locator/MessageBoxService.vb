Imports CDBNETCL.My.Resources

Public Class MessageBoxService
  Implements IMessageBoxService

  Public Sub ShowError(pSource As String, pMessage As String, pStackTrace As String) Implements IMessageBoxService.ShowError
    Utilities.ShowError(pSource, pMessage, pStackTrace)
  End Sub

  Public Sub ShowErrorMessage(ByVal pMessage As String, Optional ByVal pParam1 As String = "", Optional ByVal pParam2 As String = "") Implements IMessageBoxService.ShowErrorMessage
    Utilities.ShowErrorMessage(pMessage, pParam1, pParam2)
  End Sub

  Public Sub ShowInformationMessage(ByVal pMessage As String, Optional ByVal pParam1 As String = "", Optional ByVal pParam2 As String = "") Implements IMessageBoxService.ShowInformationMessage
    Utilities.ShowInformationMessage(pMessage, pParam1, pParam2)
  End Sub
  Public Sub ShowInformationMessage(ByVal pMessage As String, ByVal pParam1 As String, ByVal pParam2 As String, ByVal pParam3 As String) Implements IMessageBoxService.ShowInformationMessage
    Utilities.ShowInformationMessage(pMessage, pParam1, pParam2, pParam3)
  End Sub

  Public Function ShowQuestion(ByVal pQuestion As String, ByVal pButtons As MessageBoxButtons, Optional ByVal pParam1 As String = "", Optional ByVal pParam2 As String = "", Optional ByVal pParam3 As String = "") As DialogResult Implements IMessageBoxService.ShowQuestion
    Return Utilities.ShowQuestion(pQuestion, pButtons, pParam1, pParam2, pParam3)
  End Function
  Public Function ShowQuestion(ByVal pQuestion As String, ByVal pButtons As MessageBoxButtons, ByVal pDefaultButton As MessageBoxDefaultButton, Optional ByVal pParam1 As String = "", Optional ByVal pParam2 As String = "") As DialogResult Implements IMessageBoxService.ShowQuestion
    Return Utilities.ShowQuestion(pQuestion, pButtons, pDefaultButton, pParam1, pParam2)
  End Function

  Public Sub ShowWarningMessage(ByVal pMessage As String, Optional ByVal pParam1 As String = "", Optional ByVal pParam2 As String = "") Implements IMessageBoxService.ShowWarningMessage
    Utilities.ShowWarningMessage(pMessage, pParam1, pParam2)
  End Sub

  Public Sub HandleException(ByVal ex As Exception) Implements IMessageBoxService.HandleException
    DataHelper.HandleException(ex)
  End Sub

  Public Sub HandleException(ByVal parent As Form, ByVal ex As Exception) Implements IMessageBoxService.HandleException
    DataHelper.HandleException(parent, ex)
  End Sub
End Class
