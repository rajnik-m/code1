Public Class ExamCentreAccreditationFilter

  Public Sub New(pEnv As CDBEnvironment, ByVal pExamCentres As List(Of Integer), Optional ByVal pAccreditationFilter As ExamAccreditationFilterTypes = ExamAccreditationFilterTypes.eafAllowRegistration)
    ExamCentres = pExamCentres
    Environment = pEnv
    AccreditationFilter = pAccreditationFilter
  End Sub

  Private Property ExamCentres As List(Of Integer)
  Private Property Environment As CDBEnvironment
  Private Property AccreditationFilter As ExamAccreditationFilterTypes


  Public Function GetUnaccreditedCentres() As List(Of Integer)

    Dim vRtn As List(Of Integer) = Nothing

    If ExamCentres IsNot Nothing AndAlso ExamCentres.Count > 0 Then
      Dim vJoins As New AnsiJoins

      vJoins.AddLeftOuterJoin("exam_accreditation_statuses", "exam_accreditation_statuses.accreditation_status", "exam_centres.accreditation_status")

      Dim vWhere As New CDBFields
      Dim vExamCentres As String = String.Join(",", ExamCentres)
      vWhere.Add("exam_centre_id", vExamCentres, CDBField.FieldWhereOperators.fwoIn)

      If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      Else
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      End If
      vWhere.Add(Environment.Connection.DBIsNull("exam_centres.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhere.Add(Environment.Connection.DBIsNull("exam_centres.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhere.Add("exam_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)


      Dim vSQL As New SQLStatement(Environment.Connection, "exam_centres.exam_centre_id", "exam_centres", vWhere, "", vJoins)

      vRtn = BuildList(vSQL)

    End If

    Return vRtn

  End Function
  Public Function GetUnaccreditedCentresForUnit(pExamUnitId As Integer) As List(Of Integer)

    Dim vRtn As List(Of Integer) = Nothing

    If ExamCentres IsNot Nothing AndAlso ExamCentres.Count > 0 Then
      Dim vJoins As New AnsiJoins
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_centre_id", "exam_centres.exam_centre_id")
      vJoins.Add("exam_unit_links", "exam_centre_units.exam_unit_link_id", "exam_unit_links.exam_unit_link_id")
      vJoins.Add("exam_units", "exam_unit_links.exam_unit_id_2", "exam_units.exam_unit_id")
      vJoins.AddLeftOuterJoin("exam_accreditation_statuses", "exam_accreditation_statuses.accreditation_status", "exam_unit_links.accreditation_status")

      Dim vWhere As New CDBFields
      vWhere.Add("exam_units.exam_unit_id", pExamUnitId)
      Dim vExamCentres As String = String.Join(",", ExamCentres)
      vWhere.Add("exam_centres.exam_centre_id", vExamCentres, CDBField.FieldWhereOperators.fwoIn)

      If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      Else
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      End If
      vWhere.Add(Environment.Connection.DBIsNull("exam_unit_links.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhere.Add(Environment.Connection.DBIsNull("exam_unit_links.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhere.Add("exam_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)

      Dim vSQL As New SQLStatement(Environment.Connection, "exam_centres.exam_centre_id", "exam_centres", vWhere, "", vJoins)

      vRtn = BuildList(vSQL)

    End If

    Return vRtn

  End Function

  Public Function GetUnaccreditedCentresForCentreUnit(pExamUnitId As Integer, pSessionId As Integer) As List(Of Integer)

    Dim vRtn As List(Of Integer) = Nothing

    If ExamCentres IsNot Nothing AndAlso ExamCentres.Count > 0 Then
      Dim vJoins As New AnsiJoins
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_centre_id", "exam_centres.exam_centre_id")
      vJoins.Add("exam_unit_links", "exam_centre_units.exam_unit_link_id", "exam_unit_links.exam_unit_link_id")
      If pSessionId > 0 Then
        vJoins.Add("exam_units", "exam_unit_links.exam_unit_id_2", "exam_units.exam_base_unit_id")
      Else
        vJoins.Add("exam_units", "exam_unit_links.exam_unit_id_2", "exam_units.exam_unit_id")
      End If
      vJoins.AddLeftOuterJoin("exam_accreditation_statuses centre_unit_accreditation_statuses", "centre_unit_accreditation_statuses.accreditation_status", "exam_centre_units.accreditation_status")

      Dim vWhere As New CDBFields
      vWhere.Add("exam_units.exam_unit_id", pExamUnitId)
      Dim vExamCentres As String = String.Join(",", ExamCentres)
      vWhere.Add("exam_centres.exam_centre_id", vExamCentres, CDBField.FieldWhereOperators.fwoIn)

      If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
        vWhere.Add(Environment.Connection.DBIsNull("centre_unit_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      Else
        vWhere.Add(Environment.Connection.DBIsNull("centre_unit_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      End If
      vWhere.Add(Environment.Connection.DBIsNull("exam_centre_units.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhere.Add(Environment.Connection.DBIsNull("exam_centre_units.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhere.Add("centre_unit_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)

      Dim vSQL As New SQLStatement(Environment.Connection, "exam_centres.exam_centre_id", "exam_centres", vWhere, "", vJoins)

      vRtn = BuildList(vSQL)

    End If

    Return vRtn

  End Function


  Private Function BuildList(vSQL As SQLStatement) As List(Of Integer)

    Dim vRtn As New List(Of Integer)

    Dim vUnaccreditedCentres As DataTable = vSQL.GetDataTable()
    For Each vRow As DataRow In vUnaccreditedCentres.Rows
      vRtn.Add(CInt(vRow("exam_centre_id")))
    Next

    Return vRtn

  End Function

End Class
