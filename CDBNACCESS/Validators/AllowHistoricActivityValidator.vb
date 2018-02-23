Public Class AllowHistoricActivityValidator : Inherits AllowHistoricValidator

  Public Sub New(pEnv As CDBEnvironment, pActivity As String)
    MyBase.New(pEnv, "activities", New CDBFields(New CDBField("activity", pActivity)))
  End Sub

End Class
