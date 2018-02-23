Imports System.Reflection
Imports System.Linq

Namespace Access.Deduplication
  Public MustInherit Class DedupDataGeneratorBase
    Implements IDedupDataGenerator

    Private mvParent As DedupDataSelection
    Private mvEnv As CDBEnvironment

    Public Property Environment As CDBEnvironment
      Get
        Dim vEnv As CDBEnvironment = Nothing
        If mvEnv IsNot Nothing Then vEnv = mvEnv
        If vEnv Is Nothing AndAlso Me.Parent IsNot Nothing Then vEnv = Me.Parent.Environment
        Return vEnv
      End Get
      Private Set(value As CDBEnvironment)
        mvEnv = value
      End Set
    End Property

    Public Sub New()

    End Sub

    Public MustOverride ReadOnly Property Rules As List(Of DedupRule) Implements IDedupDataGenerator.Rules

    Public MustOverride Function GenerateSQLStatement(pRule As Access.Deduplication.DedupRule) As SQLStatement Implements IDedupDataGenerator.GenerateSQLStatement

    Public Overridable Property Parent As DedupDataSelection Implements IDedupDataGenerator.Parent
      Get
        Return mvParent
      End Get
      Set(value As DedupDataSelection)
        mvParent = value
      End Set
    End Property
    Public Overridable Property Code As String Implements IDedupDataGenerator.Code

    Public Overridable Property Headings As String Implements IDedupDataGenerator.Headings

    Public Overridable Property RequiredItems As String Implements IDedupDataGenerator.RequiredItems

    Public Overridable Property ResultColumns As String Implements IDedupDataGenerator.ResultColumns

    Public Overridable Property SelectColumns As String Implements IDedupDataGenerator.SelectColumns

    Public Overridable Property Widths As String Implements IDedupDataGenerator.Widths

    Public Overridable Property Connection As CDBConnection Implements IDedupDataGenerator.Connection

    Protected Overridable Sub ApplySpecialTransforms(pEnv As CDBEnvironment, pField As CDBField)

    End Sub

    Protected Overridable Sub GenerateSQLDedupClause(pEnv As CDBEnvironment, vResult As SQLStatement, pRule As DedupRule, pParameters As CDBParameters)
      If pRule.Clauses.All(Function(pClause) pParameters.HasValue(pClause.Parameter)) Then 'all the clauses in the rule must have a value, otherwise the rule is not applicable.  If you want a rule with an 'OR', create 2 rules
        pRule.Clauses.ForEach(
        Sub(vClause)
          If pParameters.HasValue(vClause.Parameter) Then
            Dim vJoins As List(Of AnsiJoin) = DedupDataSelection.CreateJoins(vClause)
            vResult.JoinFields.AddRange(vJoins)
            Dim vWhereField As CDBField = DedupDataSelection.CreateWhere(vClause, pParameters(vClause.Parameter).Value, pEnv.Connection)
            ApplySpecialTransforms(pEnv, vWhereField)
            vResult.WhereFields.Add(vWhereField)
          End If
        End Sub)
      End If
    End Sub
  End Class
End Namespace
