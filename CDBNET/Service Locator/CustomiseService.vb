Public Class CustomiseService
  Implements ICustomiseService

  Public Function Execute(webparms As ParameterList) As DialogResult Implements ICustomiseService.Execute
    Dim vFrmDisplayList As New frmDisplayList(CDBNETCL.frmDisplayList.ListUsages.CustomiseDisplayList, webparms)
    Return vFrmDisplayList.ShowDialog()
  End Function

End Class
