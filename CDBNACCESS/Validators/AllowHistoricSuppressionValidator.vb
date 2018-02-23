Public Class AllowHistoricSuppressionValidator : Inherits AllowHistoricValidator
  Public Sub New(pEnv As CDBEnvironment, pSuppression As String)
    MyBase.New(pEnv, "mailing_suppressions", New CDBFields(New CDBField("mailing_suppression", pSuppression)))
  End Sub
End Class
