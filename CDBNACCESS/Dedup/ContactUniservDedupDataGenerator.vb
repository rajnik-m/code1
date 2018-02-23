Imports System.Reflection
Imports System.Linq

Namespace Access.Deduplication
  <EnumEquivalent(DedupDataSelection.DataSelectionTypes.Contacts)>
  <EnumEquivalent(DedupDataSelection.DataSelectionTypes.Uniserv)>
  Public Class ContactUniservDedupDataGenerator
    Inherits DedupDataGeneratorBase

    Private mvParent As DedupDataSelection
    Private pEnv As CDBEnvironment
    Private UniservLock As New Object


    Public Sub New()
      MyBase.New()
      Init()
    End Sub

    Protected Sub Init()
      Dim vPrimaryList As Boolean = False
      Me.ResultColumns = "ContactNumber,Surname,Forenames,Postcode,Town,Address,DateOfBirth,Status,ContactType,OwnershipAccessLevel,OwnershipGroup,MatchRule,MatchValue,MatchRank,RankValue"
      Me.SelectColumns = "ContactNumber,Surname,Forenames,Postcode,Town,Address,DateOfBirth,Status,ContactType,OwnershipAccessLevel,OwnershipGroup,MatchRule,MatchValue,MatchRank,RankValue"
      Me.Headings = "Number,Surname,Forenames,Main Postcode,Town, Country,Address,Date of Birth,Status,Contact Type,Ownership Access Level,Ownership Group,Match Rule, Match Value, Match Rank"
      Me.RequiredItems = "ContactNumber,RankValue,OwnershipAccessLevel"
      Me.Code = "DDCU" 'De-Dupe-Contact-Uniserv
      Dim vHeadingsItems() As String = Me.Headings.Split(","c)
      Dim vWidths As New StringBuilder
      Dim vSeparator As String = String.Empty
      For vIndex As Integer = 0 To vHeadingsItems.Length - 1
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
        Using vSourceStream As IO.Stream = vCurrentAssembly.GetManifestResourceStream("CARE.ContactUniservDedupRules.xml")
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
    ''' <summary>
    ''' Generate the SQL required to implement a UniservContactDedupRule
    ''' </summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pRule">From ContactUniservdedupRules</param>
    ''' <param name="pParams"></param>
    ''' <returns>An SQL statement that implements the Rule Union with an SQL statement that will find contacts identified by Uniserv.</returns>
    ''' <remarks>If Uniserv does not find any contacts there is no Union. Try not to change this code, manipulate the rules first.</remarks>
    Private Function GenerateSQLCoreStatement(pEnv As CDBEnvironment, pRule As DedupRule, pParams As CDBParameters) As SQLStatement

      Dim vSQL As SQLStatement
      Dim vUniservContactNumbers As String = ""  'CSV of Contact Numbers returned by UNISERV
      Dim vUniservErrorNumber As Integer = 0     'UNISERV Error Number
      Dim vStreetNo As String
      Dim vDataTable As DataTable = New DataTable()

      If pParams.ContainsKey("Address") Then
        If pParams("Address").ToString().Contains(vbLf) Then
          pParams("Address").Value = pParams("Address").Value.Substring(0, pParams("Address").Value.IndexOf(vbLf)).Trim()
        End If
      End If

      vSQL = GenerateCommonSQL(pEnv, pRule, pParams)
      Me.GenerateSQLDedupClause(pEnv, vSQL, pRule, pParams)
      vDataTable = vSQL.GetDataTable 'Datable containing contacts and addresses that match the rule, used by to compare with Uniserv.

      vStreetNo = pEnv.UniservInterface.GetStreetNo(pParams.OptionalValue("Address", ""))
      vUniservErrorNumber = pEnv.UniservInterface.FindContact(pParams.OptionalValue("Forenames", ""),
                                                           pParams.OptionalValue("Surname", ""),
                                                           "",
                                                           vStreetNo,
                                                           pParams.OptionalValue("Address", ""),
                                                           pParams.OptionalValue("Town", ""), pParams.OptionalValue("Postcode", ""),
                                                           pParams.OptionalValue("Country", ""),
                                                           "",
                                                           vUniservContactNumbers) 'Uniserv return a CSV in vUniservContactNumbers
      If vUniservErrorNumber = 0 Then
        If Not String.IsNullOrEmpty(vUniservContactNumbers) Then
          'Remove any contact numbers form the Uniserv contact numbers that exist in the results of the Query that matched the rule.
          Dim vContactNumbers As New ArrayList(vUniservContactNumbers.Split(CChar(",")))
          For Each vRow As DataRow In vDataTable.Rows
            If vContactNumbers.Contains(vRow("contact_number").ToString()) Then
              vContactNumbers.RemoveAt(vContactNumbers.IndexOf(vRow("contact_number").ToString()))
            End If
          Next
          vUniservContactNumbers = String.Join(",", vContactNumbers.ToArray())
        End If
        If vUniservContactNumbers.Length > 0 Then
          'If there are any Contact numbers left in the Uniserv contact number, get the contact and address details for the contacts
          Dim vUniservSQL As SQLStatement = GenerateUniservSQL(pEnv, pRule, pParams, vUniservContactNumbers)
          vSQL.AddUnion(vUniservSQL)
        End If
      End If
      Return vSQL
    End Function

    Private Function CheckContactNameAttrs(pEnv As CDBEnvironment, ByRef pAttrs As String) As String
      Dim vContact As New Contact(pEnv)
      Dim vItems As New CDBParameters
      Dim vContactItems As New CDBParameters
      Dim vParam As CDBParameter

      vContact.Init()
      vContactItems.InitFromUniqueList(Replace(vContact.GetRecordSetFieldsName, "c.contact_number,", ""))
      vItems.InitFromUniqueList(pAttrs)
      For Each vParam In vContactItems
        If Not vItems.Exists((vParam.Name)) Then vItems.Add((vParam.Name), CDBField.FieldTypes.cftCharacter, vParam.Value)
      Next vParam
      Return vItems.ItemList
    End Function

    Private Function RemoveBlankItems(ByVal pItems As String) As String
      While pItems.Contains(",,")
        pItems = pItems.Replace(",,", ",")
      End While
      'BR 8771: Remove any last comma from end of line in case blank item(s) were at end:
      If pItems.EndsWith(",") Then pItems = pItems.Substring(0, pItems.Length - 1)
      Return pItems
    End Function
    ''' <summary>
    ''' Generate SQL that will implement the pRule using the values passed in pParams
    ''' </summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pRule"></param>
    ''' <param name="pParams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GenerateCommonSQL(pEnv As CDBEnvironment, pRule As DedupRule, pParams As CDBParameters) As SQLStatement

      Dim vSQLJoins As New AnsiJoins
      Dim vSQLWhere As New CDBFields

      Dim vSelectColumns As String = "contacts.contact_number,surname,forenames,addresses.postcode,addresses.town,addresses.country,addresses.address_line1,date_of_birth,contacts.status,contacts.contact_type,ownership_group_users.ownership_access_level,ownership_group_users.ownership_group"
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
      Me.SelectColumns = vSelectColumns

      vSQLJoins.Add("contact_addresses", "contact_addresses.contact_number", "contacts.contact_number")
      vSQLJoins.Add("addresses", "addresses.address_number", "contact_addresses.address_number")
      vSQLJoins.Add("ownership_group_users", "ownership_group_users.ownership_group", "contacts.ownership_group")
      vSQLWhere = New CDBFields
      vSQLWhere.Add("ownership_group_users.logname", pEnv.User.Logname)
      vSQLWhere.Add("ownership_group_users.valid_from", Date.Today, CDBField.FieldWhereOperators.fwoLessThanEqual)
      vSQLWhere.Add("ownership_group_users.valid_to", Date.Today, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

      If pParams.Exists("ContactGroup") Then
        If pParams("ContactGroup").Value = pEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
          vSQLWhere.Add("contacts.contact_group", pParams("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
        Else
          vSQLWhere.Add("contacts.contact_group", pParams("ContactGroup").Value)
        End If
      End If
      Return New SQLStatement(pEnv.Connection, vSelectColumns, "contacts", vSQLWhere, String.Empty, vSQLJoins)
    End Function
    ''' <summary>
    ''' Generate similar SQL to GenerateCommonSQL, but just get the Uniserv contacts.
    ''' </summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pRule">Used to provide match data, not contact and addresses</param>
    ''' <param name="pParams"></param>
    ''' <param name="pUniservContactNumbers">CSV string of contact numbers returned by Uniserv</param>
    ''' <returns>This function does not implement a rule, it selects contacts identified by Uniserv when it had the same information as the rule.</returns>
    ''' <remarks></remarks>
    Private Function GenerateUniservSQL(pEnv As CDBEnvironment, pRule As DedupRule, pParams As CDBParameters, pUniservContactNumbers As String) As SQLStatement

      Dim vSQLJoins As New AnsiJoins
      Dim vSQLWhere As New CDBFields
      Dim vSelectColumns As String = "contacts.contact_number,surname,forenames,addresses.postcode,addresses.town,addresses.country,addresses.address_line1,date_of_birth,contacts.status,contacts.contact_type,ownership_group_users.ownership_access_level,ownership_group_users.ownership_group"

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

      vSQLJoins.Add("contact_addresses", "contact_addresses.contact_number", "contacts.contact_number")
      vSQLJoins.Add("addresses", "addresses.address_number", "contact_addresses.address_number")
      vSQLJoins.Add("ownership_group_users", "ownership_group_users.ownership_group", "contacts.ownership_group")

      vSQLWhere = New CDBFields
      vSQLWhere.Add("ownership_group_users.logname", pEnv.User.Logname)
      vSQLWhere.Add("ownership_group_users.valid_from", Date.Today, CDBField.FieldWhereOperators.fwoLessThanEqual)
      vSQLWhere.Add("ownership_group_users.valid_to", Date.Today, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

      If pParams.Exists("ContactGroup") Then
        If pParams("ContactGroup").Value = pEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
          vSQLWhere.Add("contacts.contact_group", pParams("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
        Else
          vSQLWhere.Add("contacts.contact_group", pParams("ContactGroup").Value)
        End If
      End If
      vSQLWhere.Add("contacts.contact_number", CDBField.FieldTypes.cftInteger, pUniservContactNumbers, CDBField.FieldWhereOperators.fwoIn)
      Return New SQLStatement(pEnv.Connection, vSelectColumns, "contacts", vSQLWhere, String.Empty, vSQLJoins)
    End Function



  End Class

End Namespace
