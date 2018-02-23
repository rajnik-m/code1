Public Class SystemUser

  Private mvUser As CDBUser = Nothing

  Private Sub New(pUser As CDBUser)
    mvUser = pUser
  End Sub

  Private Sub New(pEnv As CDBEnvironment, pLogname As String)
    mvUser = New CDBUser(pEnv)
    mvUser.Init(pLogname)
  End Sub

  Public Shared Function GetCurrent(pEnv As CDBEnvironment) As SystemUser
    Return New SystemUser(pEnv.User)
  End Function

  Public Shared Function GetByLogname(pEnv As CDBEnvironment, pLogname As String) As SystemUser
    Return New SystemUser(pEnv, pLogname)
  End Function

  Public ReadOnly Property EmailAddress() As System.Net.Mail.MailAddress
    Get
      Dim vResult As System.Net.Mail.MailAddress = Nothing
      Try
        vResult = New Net.Mail.MailAddress(New SQLStatement(mvUser.Environment.Connection,
                                                            "co.""number""", "contacts c",
                                                             New CDBFields({New CDBField("c.contact_number", mvUser.ContactNumber, CDBField.FieldWhereOperators.fwoEqual),
                                                                            New CDBField("co.valid_from", CDBField.FieldTypes.cftDate, TodaysDateAndTime, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual),
                                                                            New CDBField("co.valid_to", CDBField.FieldTypes.cftDate, TodaysDateAndTime, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual),
                                                                            New CDBField("d.email", "Y", CDBField.FieldWhereOperators.fwoEqual)}),
                                                            "CASE WHEN co.preferred_method = 'Y' THEN 0 WHEN co.device_default = 'Y' THEN 1  WHEN co.device_default = 'N' THEN 2 ELSE 3 END",
                                                            New AnsiJoins({New AnsiJoin("communications co", "co.contact_number", "c.contact_number", AnsiJoin.AnsiJoinTypes.InnerJoin),
                                                                           New AnsiJoin("devices d", "d.device", "co.device", AnsiJoin.AnsiJoinTypes.InnerJoin)})).GetValue)
      Catch ex As Exception
      End Try
      Return vResult
    End Get
  End Property
End Class
