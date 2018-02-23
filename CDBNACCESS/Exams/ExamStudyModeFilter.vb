Imports System.Xml
Imports System.Linq

Public Class ExamStudyModeFilter

  Public Sub New(pEnv As CDBEnvironment, ByVal pExamUnitLinkIds As Dictionary(Of Integer, Integer))
    ExamUnitLinks = pExamUnitLinkIds
    Environment = pEnv
  End Sub

  Private Property ExamUnitLinks As Dictionary(Of Integer, Integer)
  Private Property Environment As CDBEnvironment

  Public Function GetUnitsForStudyMode(ByVal pStudyMode As String, pSessionId As Integer) As List(Of Integer)

    Dim vRtn As List(Of Integer)

    If pStudyMode.Length > 0 Then
      vRtn = GetUnitsForNamedStudyMode(pStudyMode, pSessionId)
    Else
      vRtn = GetUnitsForNoStudyMode(pSessionId)
    End If

    Return vRtn

  End Function

  Private Function GetUnitsForNamedStudyMode(pStudyMode As String, pSessionId As Integer) As List(Of Integer)

    Dim vJoins As New AnsiJoins
    If pSessionId > 0 Then
      vJoins.Add("exam_unit_links", "exam_unit_links.base_unit_link_id", "exam_unit_study_modes.exam_unit_link_id") 'Join will only be valid if the passed unit links are session-based
    Else
      vJoins.Add("exam_unit_links", "exam_unit_links.exam_unit_link_id", "exam_unit_study_modes.exam_unit_link_id") 'Join will only be valid if the passed unit links are non-session-based
    End If

    Dim vWhere As New CDBFields

    vWhere.Add("exam_unit_study_modes.study_mode", pStudyMode)

    Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)
    vWhere.Add("exam_unit_links.exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)

    Dim vSQL As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_unit_study_modes", vWhere, "", vJoins)

    Dim vRtn As List(Of Integer) = BuildHierarchicalList(vSQL)

    Return vRtn

  End Function

  Private Function GetUnitsForNoStudyMode(ByVal pSessionId As Integer) As List(Of Integer)

    'The exam_tree SQL below creates a complete tree from leaf to root for every bottom-level unit it finds.  The tree_id is the exam_unit_link_id of the leaf.
    'It gets this so that it can evaluate the study mode for the whole exam tree, as Study modes can be attached anywhere in the exam tree, but only once.
    Dim vSQLString As String = _
      <exam_tree_sql>
WITH exam_tree (tree_id, exam_unit_link_id)
as
(
	SELECT exam_unit_link_id tree_id, exam_unit_link_id
	FROM exam_unit_links
  JOIN exam_units
    ON exam_units.exam_unit_id = exam_unit_links.exam_unit_id_2
	WHERE exam_units.exam_session_id is null
  AND NOT EXISTS
	(
		SELECT parent_unit_link_id
		FROM exam_unit_links parent_unit_links
		WHERE parent_unit_link_id > 0
		AND parent_unit_links.parent_unit_link_id = exam_unit_links.exam_unit_link_id
	)
UNION ALL
SELECT exam_tree.tree_id, exam_unit_links.parent_unit_link_id
FROM exam_unit_links
	JOIN exam_tree
		ON exam_tree.exam_unit_link_id = exam_unit_links.exam_unit_link_id
)
SELECT bottom_leaf_link.exam_unit_link_id
FROM exam_tree
	LEFT JOIN exam_unit_study_modes ON exam_unit_study_modes.exam_unit_link_id = exam_Tree.exam_unit_link_id
		and exam_unit_study_modes.exam_unit_link_id is null
	JOIN exam_unit_links
		ON exam_unit_links.{0} = exam_tree.exam_unit_link_id
	JOIN exam_unit_links bottom_leaf_link
		ON bottom_leaf_link.{1} = exam_tree.tree_id
WHERE exam_unit_links.exam_unit_link_id in (0{2})
GROUP BY bottom_leaf_link.exam_unit_link_id
      </exam_tree_sql>.Value

    Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)

    If pSessionId > 0 Then
      vSQLString = String.Format(vSQLString, "base_unit_link_id", "base_unit_link_id", vExamUnitLinks)
    Else
      vSQLString = String.Format(vSQLString, "exam_unit_link_id", "exam_unit_link_id", vExamUnitLinks)
    End If

    vSQLString = String.Format(vSQLString, vExamUnitLinks)

    Dim vSQL As New SQLStatement(Environment.Connection, vSQLString)

    Dim vRtn As List(Of Integer) = BuildHierarchicalList(vSQL)
    Return vRtn

  End Function

  Private Function BuildHierarchicalList(vSQL As SQLStatement) As List(Of Integer)

    Dim vRtn As New List(Of Integer)

    Dim vUnits As DataTable = vSQL.GetDataTable()
    For Each vRow As DataRow In vUnits.Rows
      vRtn.Add(CInt(vRow("exam_unit_link_id")))
    Next

    Dim vChildren As New List(Of Integer)
    Dim vParents As New List(Of Integer)

    'Return all children and parents of units that are included in the SQL.  
    'Any unit that is valid for a Study Mode also includes all the unit's parents and children for that Study Mode.  This is because Study Modes can only be set on one level.

    'We need to sort the exam links to ensure that we capture all units in the tree.  When evaluating a child, we must ensure that all its parents have been evaluated first
    Dim vSortedLinks As IEnumerable(Of KeyValuePair(Of Integer, Integer)) = From vEntry In ExamUnitLinks Order By vEntry.Value

    'First traverse the tree from the top down(parent to child), to get all children of included units
    For Each vEntry As KeyValuePair(Of Integer, Integer) In vSortedLinks
      'If the exam node's parent is included then include it
      If vEntry.Value > 0 AndAlso vRtn.Contains(vEntry.Value) Then
        vRtn.Add(vEntry.Key)
      End If
    Next

    vSortedLinks = vSortedLinks.Reverse()

    'Now traverse the tree from the bottom up(child to parent), to get all parents of included units
    For Each vEntry As KeyValuePair(Of Integer, Integer) In vSortedLinks
      'If the exam node is included then include its parent
      If vEntry.Value > 0 AndAlso vRtn.Contains(vEntry.Key) Then
        vRtn.Add(vEntry.Value)
      End If
    Next

    Return vRtn
  End Function

  Function GetCentreUnitsForStudyMode(pStudyMode As String, pCentreId As Integer, pSessionId As Integer) As List(Of Integer)

    Dim vRtn As List(Of Integer)

    If pStudyMode.Length > 0 Then
      vRtn = GetCentreUnitsForNamedStudyMode(pStudyMode, pCentreId, pSessionId)
    Else
      vRtn = GetUnitsForNoStudyMode(pSessionId)
    End If

    Return vRtn

  End Function

  Private Function GetCentreUnitsForNamedStudyMode(pStudyMode As String, pCentreId As Integer, pSessionId As Integer) As List(Of Integer)

    Dim vJoins As New AnsiJoins
    vJoins.Add("exam_centre_units", "exam_centre_units.exam_centre_unit_id", "exam_centre_unit_study_modes.exam_centre_unit_link_id")
    If pSessionId > 0 Then
      vJoins.Add("exam_unit_links", "exam_unit_links.base_unit_link_id", "exam_centre_units.exam_unit_link_id")
    Else
      vJoins.Add("exam_unit_links", "exam_unit_links.exam_unit_link_id", "exam_centre_units.exam_unit_link_id")
    End If

    Dim vWhere As New CDBFields
    vWhere.Add("exam_centre_unit_study_modes.study_mode", pStudyMode)
    vWhere.Add("exam_centre_units.exam_centre_id", pCentreId)

    Dim vExamUnitLinks As String = String.Join(",", ExamUnitLinks.Keys)
    vWhere.Add("exam_unit_links.exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)

    Dim vSQL As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_centre_unit_study_modes", vWhere, "", vJoins)


    'Join for any centre units that have no study modes selected.  This means that whatever study modes were selected at session sero level will be available.
    'Warning: this is a UNION query with a Not Exists sub-query so it's a bit messy with the objects
    Dim vUnionJoins As New AnsiJoins
    If pSessionId > 0 Then
      vUnionJoins.Add("exam_unit_links", "exam_unit_links.base_unit_link_id", "exam_unit_study_modes.exam_unit_link_id")
    Else
      vUnionJoins.Add("exam_unit_links", "exam_unit_links.exam_unit_link_id", "exam_unit_study_modes.exam_unit_link_id")
    End If

    Dim vUnionWhere As New CDBFields
    vUnionWhere.Add("exam_unit_study_modes.study_mode", pStudyMode)

    vUnionWhere.Add("exam_unit_links.exam_unit_link_id", vExamUnitLinks, CDBField.FieldWhereOperators.fwoIn)


    Dim vNotExistsSQL As String = _
      <SQL>
          SELECT exam_centre_units.exam_unit_link_id
	        FROM exam_centre_unit_study_modes 
		        JOIN exam_centre_units
			        ON exam_centre_units.exam_centre_unit_id = exam_centre_unit_study_modes.exam_centre_unit_link_id
	        WHERE {0}
          AND exam_centre_units.exam_centre_id = {1}
      </SQL>.Value
    If pSessionId > 0 Then
      vNotExistsSQL = String.Format(vNotExistsSQL, "exam_centre_units.exam_unit_link_id = exam_unit_links.base_unit_link_id", pCentreId)
    Else
      vNotExistsSQL = String.Format(vNotExistsSQL, "exam_centre_units.exam_unit_link_id = exam_unit_links.exam_unit_link_id", pCentreId)
    End If
    vUnionWhere.Add("Exclude", vNotExistsSQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)

    Dim vUnion As New SQLStatement(Environment.Connection, "exam_unit_links.exam_unit_link_id", "exam_unit_study_modes", vUnionWhere, "", vUnionJoins)

    vSQL.AddUnion(vUnion)

    Dim vRtn As List(Of Integer) = BuildHierarchicalList(vSQL)

    Return vRtn


  End Function

End Class
