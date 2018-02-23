<Flags>
Public Enum ExamAccreditationFilterTypes As Integer
  eafAllowRegistration = 1
  eafAllowResultEntry = 2
End Enum

Public Class ExamAccreditationFilter

  Public Sub New(pEnv As CDBEnvironment, ByVal pExamUnitLinkIds As Dictionary(Of Integer, Integer), Optional ByVal pAccreditationFilter As ExamAccreditationFilterTypes = ExamAccreditationFilterTypes.eafAllowRegistration)
    ExamUnitLinks = pExamUnitLinkIds
    Environment = pEnv
    AccreditationFilter = pAccreditationFilter
  End Sub

  Private Property ExamUnitLinks As Dictionary(Of Integer, Integer)
  Private Property Environment As CDBEnvironment
  Private Property AccreditationFilter As ExamAccreditationFilterTypes


  Public Function GetUnaccreditedUnits() As List(Of Integer)

    Dim vRtn As New List(Of Integer)

    If Me.ExamUnitLinks.Count > 0 Then
      Dim vJoins As New AnsiJoins

      vJoins.AddLeftOuterJoin("exam_accreditation_statuses", "exam_accreditation_statuses.accreditation_status", "exam_unit_links.accreditation_status")

      Dim vWhere As New CDBFields
      Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)
      vWhere.Add("exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)

      If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      Else
        vWhere.Add(Environment.Connection.DBIsNull("exam_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      End If
      vWhere.Add(Environment.Connection.DBIsNull("exam_unit_links.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhere.Add(Environment.Connection.DBIsNull("exam_unit_links.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhere.Add("exam_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)



      Dim vSQL As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_unit_links", vWhere, "", vJoins)

      vRtn = BuildHierarchicalList(vSQL)

    End If

    Return vRtn

  End Function

  Public Function GetUnitsAtUnaccreditedCentre(ByVal pCentreCode As String, ByVal pSessionCode As String) As List(Of Integer)
    Dim vCenterId As Integer
    Dim vSessionId As Integer
    Dim vWhere As New CDBFields

    vWhere.Add("exam_centres.exam_centre_code", pCentreCode)
    Dim vSQLCentre As New SQLStatement(Environment.Connection, "exam_centre_id", "exam_centres", vWhere)
    vCenterId = CInt(vSQLCentre.GetDataTable().Rows(0)("exam_centre_id").ToString)

    If pSessionCode <> "0" Then
      vWhere.Clear()
      vWhere.Add("exam_sessions.exam_session_code", pSessionCode)
      Dim vSQLSession As New SQLStatement(Environment.Connection, "exam_session_id", "exam_sessions", vWhere)
      vCenterId = CInt(vSQLSession.GetDataTable().Rows(0)("exam_session_id").ToString)
    End If

    Return GetUnitsAtUnaccreditedCentre(vCenterId, vSessionId)

  End Function

  Public Function GetUnitsAtUnaccreditedCentre(pCentreId As Integer, pSessionId As Integer) As List(Of Integer)

    Dim vJoins As New AnsiJoins
    If pSessionId > 0 Then
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_unit_link_id", "exam_unit_links.base_unit_link_id")
    Else
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_unit_link_id", "exam_unit_links.exam_unit_link_id")
    End If
    vJoins.Add("exam_centres", "exam_centres.exam_centre_id", "exam_centre_units.exam_centre_id")
    vJoins.AddLeftOuterJoin("exam_accreditation_statuses centre_accreditation_statuses", "centre_accreditation_statuses.accreditation_status", "exam_centres.accreditation_status")

    Dim vWhere As New CDBFields
    vWhere.Add("exam_centres.exam_centre_id", pCentreId)
    Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)
    vWhere.Add("exam_unit_links.exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)

    If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
      vWhere.Add(Environment.Connection.DBIsNull("centre_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
    Else
      vWhere.Add(Environment.Connection.DBIsNull("centre_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
    End If
    vWhere.Add(Environment.Connection.DBIsNull("exam_centres.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
    vWhere.Add(Environment.Connection.DBIsNull("exam_centres.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
    vWhere.Add("centre_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)


    Dim vSQL As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_unit_links", vWhere, "", vJoins)

    Dim vRtn As List(Of Integer) = BuildHierarchicalList(vSQL)

    Return vRtn

  End Function
  Public Function GetUnaccreditedCentreUnits(pCentreId As Integer, pSessionId As Integer) As List(Of Integer)

    Dim vJoins As New AnsiJoins
    If pSessionId > 0 Then
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_unit_link_id", "exam_unit_links.base_unit_link_id")
    Else
      vJoins.Add("exam_centre_units", "exam_centre_units.exam_unit_link_id", "exam_unit_links.exam_unit_link_id")
    End If
    vJoins.AddLeftOuterJoin("exam_accreditation_statuses centre_unit_accreditation_statuses", "centre_unit_accreditation_statuses.accreditation_status", "exam_centre_units.accreditation_status")

    Dim vWhere As New CDBFields
    vWhere.Add("exam_centre_units.exam_centre_id", pCentreId)
    Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)
    vWhere.Add("exam_unit_links.exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)

    If AccreditationFilter = ExamAccreditationFilterTypes.eafAllowResultEntry Then
      vWhere.Add(Environment.Connection.DBIsNull("centre_unit_accreditation_statuses.allow_result_entry", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
    Else
      vWhere.Add(Environment.Connection.DBIsNull("centre_unit_accreditation_statuses.allow_registration", "'N'"), "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
    End If
    vWhere.Add(Environment.Connection.DBIsNull("exam_centre_units.accreditation_valid_from", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
    vWhere.Add(Environment.Connection.DBIsNull("exam_centre_units.accreditation_valid_to", Environment.Connection.DBDate()), CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
    vWhere.Add("centre_unit_accreditation_statuses.ignore_accreditation_validity", "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)


    Dim vSQL As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_unit_links", vWhere, "", vJoins)

    Dim vRtn As List(Of Integer) = BuildHierarchicalList(vSQL)

    Return vRtn

  End Function

  ''' <summary>
  ''' Returns a list of ExamUnitLinkIds built from a SQL Statement returning all un-accredited units.  The function also includes all child records of unaccredited parents
  ''' </summary>
  ''' <param name="vSQL">SQL Statement that will return a table of all unaccredited unit links.  Must contain a column named exam_unit_link_id</param>
  ''' <returns>A List of all ExamUnitLinkIds that are in the ExamUnitLinks Dictionary (from the Constructor) and that are also returned by the unaccredited rows returned by the SQL Statement</returns>
  ''' <remarks>All ExamUnitLinkIds that are returned by the SQL Statement will be returned.  All descendents of those ExamUnitLinkIds will also be returned</remarks>
  Private Function BuildHierarchicalList(vSQL As SQLStatement) As List(Of Integer)

    Dim vRtn As New List(Of Integer)

    Dim vUnaccreditedUnits As DataTable = vSQL.GetDataTable()
    For Each vRow As DataRow In vUnaccreditedUnits.Rows
      vRtn.Add(CInt(vRow("exam_unit_link_id")))
    Next

    'Exclude all children of units that are excluded
    For Each vEntry As KeyValuePair(Of Integer, Integer) In ExamUnitLinks
      If vEntry.Value > 0 AndAlso vRtn.Contains(vEntry.Value) Then 'The value is the parent unit link Id - If the parent is in the return list then the child will be in the return list.  All descendents of the parent need to be returned
        vRtn.Add(vEntry.Key) 'The Key is the exam unit link id.  Normally the key should be the parent and the value should be the unit, but that's not possible as a keys have to be unique.  A two-dimensional list is what you'd need for that but it's not worth the hassle here.
      End If
    Next

    Return vRtn

  End Function

End Class
