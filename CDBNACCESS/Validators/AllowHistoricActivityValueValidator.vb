Public Class AllowHistoricActivityValueValidator : Inherits AllowHistoricValidator
  Public Sub New(pEnvironment As CDBEnvironment, pActivity As String, pActivityValue As String)
    MyBase.New(pEnvironment, "activity_values", New CDBFields({New CDBField("activity", pActivity), New CDBField("activity_value", pActivityValue)}))
  End Sub

End Class
