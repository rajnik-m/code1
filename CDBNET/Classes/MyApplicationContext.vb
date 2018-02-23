Public Class MyApplicationContext
  Inherits ApplicationContext


  Public Sub New(ByVal pForm As Form)
    Try
      MainForm = pForm
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
      Application.Exit()
    End Try
  End Sub

End Class
