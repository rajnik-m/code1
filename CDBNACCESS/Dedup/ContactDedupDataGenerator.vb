Imports System.Reflection
Imports System.Linq

Namespace Access.Deduplication
  <EnumEquivalent(DedupDataSelection.DataSelectionTypes.Contacts)>
  Public Class ContactDedupDataGenerator
    Inherits DedupDataGeneratorBase

    Private mvParent As DedupDataSelection
    Private mvEnv As CDBEnvironment



    Public Sub New()
      MyBase.New()
      Init()
    End Sub

    Protected Sub Init()
      Dim vPrimaryList As Boolean = False

      Me.ResultColumns = "ContactNumber,ContactName,Forenames,Sex,DateOfBirth,Department,Address,Postcode,StatusCode,OwnershipGroup,OwnershipAccessLevel,MatchRule,MatchValue,MatchRank,RankValue"
      Me.SelectColumns = "ContactNumber,ContactName,Forenames,Sex,DateOfBirth,Department,Address,Postcode,StatusCode,OwnershipGroup,OwnershipAccessLevel,MatchRule,MatchValue,MatchRank"
      Me.Headings = "Number,Name,Forenames,Sex,Date of Birth,Department,Main Address,Main Postcode,Status,Ownership Group,Ownership Access Level,Match Rule, Match Value, Match Rank"
      Me.RequiredItems = "ContactNumber,RankValue,OwnershipAccessLevel"
      Me.Code = "DDCO" 'De-Dupe-COntact
      Dim vSelectItems() As String = Me.SelectColumns.Split(","c)
      Dim vWidths As New StringBuilder
      Dim vSeparator As String = String.Empty
      For vIndex As Integer = 0 To vSelectItems.Length - 1
        vWidths.Append(vSeparator)
        vSeparator = ","
        vWidths.Append("1000")
      Next
      Me.Widths = vWidths.ToString()

    End Sub

    Public Overrides ReadOnly Property Rules As List(Of DedupRule)
      Get
        Dim vList As New List(Of DedupRule)
        Dim vCurrentAssembly As Assembly = Assembly.GetExecutingAssembly()

        Dim vSerializer As New Xml.Serialization.XmlSerializer(vList.GetType())
        Using vSourceStream As IO.Stream = vCurrentAssembly.GetManifestResourceStream("CARE.ContactDedupRules.xml")
          If vSourceStream IsNot Nothing Then
            Dim vOutput As Object = vSerializer.Deserialize(vSourceStream)
            If vOutput IsNot Nothing AndAlso TypeOf vOutput Is List(Of DedupRule) Then
              vList = DirectCast(vOutput, List(Of DedupRule))
            End If
          End If
        End Using
        Return vList
      End Get
    End Property

    Public Overrides Function GenerateSQLStatement(pRule As Access.Deduplication.DedupRule) As SQLStatement

      Dim vResult As SQLStatement = GenerateSQLCoreStatement(Me.Environment, pRule, Parent.Parameters)

      GenerateSQLDedupClause(Me.Environment, vResult, pRule, Parent.Parameters)

      Return vResult
    End Function

    Protected Overrides Sub ApplySpecialTransforms(pEnv As CDBEnvironment, pField As CDBField)

      If pField IsNot Nothing AndAlso pField.Name IsNot Nothing Then
        Dim vFieldName As String = pField.Name.Split("."c).LastOrDefault()

        'Phone Number Transforms
        Dim vPhoneNumberFields As New List(Of String) From {"cli_number"}
        If vPhoneNumberFields.Contains(vFieldName) AndAlso pField.WhereOperator = CDBField.FieldWhereOperators.fwoEqual Then
          ApplyCliNoTransform(pEnv, pField)
        End If
      End If

    End Sub

    Private Sub ApplyCliNoTransform(pEnv As CDBEnvironment, pField As CDBField)
      pField.Value = Communication.ExtractCliNumber(pEnv, pField.Value)
    End Sub

    Private Shared Function GenerateSQLCoreStatement(pEnv As CDBEnvironment, pRule As DedupRule, pParams As CDBParameters) As SQLStatement
      Dim vSQLWhere As New CDBFields
      vSQLWhere.Add("contacts.contact_type", "C")
      Dim vSelectColumns As String = "contacts.contact_number,contacts.label_name,contacts.forenames,contacts.sex,contacts.date_of_birth,contacts.department,addresses.address,addresses.postcode,contacts.status,ownership_group_users.ownership_group,ownership_group_users.ownership_access_level"
      vSelectColumns += String.Format(",'{0}' As MatchRule,", ResourceTranslateString(pRule.Description))
      Dim vMatchColumn As String = ""
      For Each pClause As DedupClause In pRule.Clauses
        Dim vSeparator As String = If(String.IsNullOrWhiteSpace(pClause.TableAlias), "", ".")
        vMatchColumn += String.Format("{0}{1}{2}{3}' '{3}", pClause.TableAlias,
                                                            vSeparator,
                                                            If(pEnv.Connection.IsSpecialColumn(pClause.Attribute), pEnv.Connection.DBSpecialCol(pClause.Attribute), pClause.Attribute),
                                                            pEnv.Connection.ConcatonateOperator)
      Next
      vSelectColumns += vMatchColumn + "'' as MatchValue,"
      vSelectColumns += String.Format("'{0}' As MatchRank,", pRule.RuleRank.ToString())
      vSelectColumns += String.Format("{0} As RankOrder", Convert.ChangeType(pRule.RuleRank, pRule.RuleRank.GetTypeCode()))

      Dim vSQLJoins As New AnsiJoins
      vSQLJoins.Add("addresses", "addresses.address_number", "contacts.address_number")

      vSQLJoins.Add("ownership_group_users", "ownership_group_users.ownership_group", "contacts.ownership_group")
      vSQLWhere.Add("ownership_group_users.logname", pEnv.User.Logname)
      vSQLWhere.Add("ownership_group_users.valid_from", Date.Today, CDBField.FieldWhereOperators.fwoLessThanEqual)
      vSQLWhere.Add("ownership_group_users.valid_to", Date.Today, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

      If pParams.Exists("ContactGroup") Then
        If pParams("ContactGroup").Value = pEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
          vSQLWhere.Add("contact_group", pParams("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
        Else
          vSQLWhere.Add("contact_group", pParams("ContactGroup").Value)
        End If
      End If

      Dim vResult As New SQLStatement(pEnv.Connection, vSelectColumns, "contacts", vSQLWhere, "contacts.contact_type DESC, contacts.surname, contacts.contact_number", vSQLJoins)
      Return vResult
    End Function


  End Class

End Namespace
