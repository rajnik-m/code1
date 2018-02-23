Namespace Access

  Public Class OrganisationGroup
    Inherits ContactOrOrganisationGroup

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum OrganisationGroupFields
      AllFields = 0
      OrganisationGroup
      OrganisationGroupDesc
      Client
      Name
      SequenceNo
      RgbValue
      TabPrefix
      OrganisationNumber
      AddressNumber
      HiddenAttributes
      NamedAttributes
      UseHouseNames
      GraphActivity
      UnknownAddress
      UnknownTown
      AllAddressesUnknown
      PrimaryRelationship
      PositionActivityPrompt
      PositionRelationshipPrompt
      NameFormat
      LastUsedId
      CustomTableNames
      ViewInContactCard
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("organisation_group")
        .Add("organisation_group_desc")
        .Add("client")
        .Add("name")
        .Add("sequence_no", CDBField.FieldTypes.cftInteger)
        .Add("rgb_value", CDBField.FieldTypes.cftLong)
        .Add("tab_prefix")
        .Add("organisation_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("hidden_attributes", CDBField.FieldTypes.cftMemo)
        .Add("named_attributes", CDBField.FieldTypes.cftMemo)
        .Add("use_house_names")
        .Add("graph_activity")
        .Add("unknown_address")
        .Add("unknown_town")
        .Add("all_addresses_unknown")
        .Add("primary_relationship")
        .Add("position_activity_prompt")
        .Add("position_relationship_prompt")
        .Add("name_format")
        .Add("last_used_id", CDBField.FieldTypes.cftInteger)
        .Add("custom_table_names")
        .Add("view_in_contact_card")
        .Item(OrganisationGroupFields.OrganisationGroup).PrimaryKey = True

        .Item(OrganisationGroupFields.Client).PrimaryKey = True

        .Item(OrganisationGroupFields.PositionActivityPrompt).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks)
        .Item(OrganisationGroupFields.PositionRelationshipPrompt).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks)
        .Item(OrganisationGroupFields.CustomTableNames).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataOrgGroupCustomTables)
        .Item(OrganisationGroupFields.NameFormat).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cbdGroupDefaultNameFormat)
        .Item(OrganisationGroupFields.LastUsedId).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cbdGroupDefaultNameFormat)
        .Item(OrganisationGroupFields.ViewInContactCard).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbOrganisationGroupsViewInContactCard)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "og"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "organisation_groups"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public Overrides ReadOnly Property EntityGroupCode() As String
      Get
        Return mvClassFields(OrganisationGroupFields.OrganisationGroup).Value
      End Get
    End Property
    Public Overrides ReadOnly Property EntityGroupDesc() As String
      Get
        Return mvClassFields(OrganisationGroupFields.OrganisationGroupDesc).Value
      End Get
    End Property
    Public Overrides ReadOnly Property Client() As String
      Get
        Return mvClassFields(OrganisationGroupFields.Client).Value
      End Get
    End Property
    Public Overrides ReadOnly Property Name() As String
      Get
        Return mvClassFields(OrganisationGroupFields.Name).Value
      End Get
    End Property
    Public Overrides ReadOnly Property SequenceNo() As Integer
      Get
        Return mvClassFields(OrganisationGroupFields.SequenceNo).IntegerValue
      End Get
    End Property
    Public Overrides ReadOnly Property RgbValue() As Integer
      Get
        Return mvClassFields(OrganisationGroupFields.RgbValue).IntegerValue
      End Get
    End Property
    Public Overrides ReadOnly Property TabPrefix() As String
      Get
        Return mvClassFields(OrganisationGroupFields.TabPrefix).Value
      End Get
    End Property
    Public Overrides ReadOnly Property OrganisationNumber() As Integer
      Get
        Return mvClassFields(OrganisationGroupFields.OrganisationNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(OrganisationGroupFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public Overrides ReadOnly Property HiddenAttributes() As String
      Get
        Return mvClassFields(OrganisationGroupFields.HiddenAttributes).Value
      End Get
    End Property
    Public Overrides ReadOnly Property NamedAttributes() As String
      Get
        Return mvClassFields(OrganisationGroupFields.NamedAttributes).Value
      End Get
    End Property
    Public Overrides ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(OrganisationGroupFields.AmendedBy).Value
      End Get
    End Property
    Public Overrides ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(OrganisationGroupFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property UseHouseNames() As String
      Get
        Return mvClassFields(OrganisationGroupFields.UseHouseNames).Value
      End Get
    End Property
    Public ReadOnly Property GraphActivity() As String
      Get
        Return mvClassFields(OrganisationGroupFields.GraphActivity).Value
      End Get
    End Property
    Public Overrides Property UnknownAddress() As String
      Get
        Return mvClassFields(OrganisationGroupFields.UnknownAddress).Value
      End Get
      Set(ByVal value As String)
        mvClassFields(OrganisationGroupFields.UnknownAddress).Value = value
      End Set
    End Property
    Public Overrides Property UnknownTown() As String
      Get
        Return mvClassFields(OrganisationGroupFields.UnknownTown).Value
      End Get
      Set(ByVal value As String)
        mvClassFields(OrganisationGroupFields.UnknownTown).Value = value
      End Set
    End Property
    Public Overrides ReadOnly Property AllAddressesUnknown() As Boolean
      Get
        Return mvClassFields(OrganisationGroupFields.AllAddressesUnknown).Bool
      End Get
    End Property
    Public Overrides ReadOnly Property PrimaryRelationship() As String
      Get
        Return mvClassFields(OrganisationGroupFields.PrimaryRelationship).Value
      End Get
    End Property
    Public Overrides ReadOnly Property PositionActivityPrompt() As Boolean
      Get
        Return mvClassFields(OrganisationGroupFields.PositionActivityPrompt).Bool
      End Get
    End Property
    Public Overrides ReadOnly Property PositionRelationshipPrompt() As Boolean
      Get
        Return mvClassFields(OrganisationGroupFields.PositionRelationshipPrompt).Bool
      End Get
    End Property
    Public Overrides ReadOnly Property NameFormat() As String
      Get
        Return mvClassFields(OrganisationGroupFields.NameFormat).Value.ToString
      End Get
    End Property
    Public Overrides ReadOnly Property LastUsedId() As Integer
      Get
        Return mvClassFields(OrganisationGroupFields.LastUsedId).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CustomTableNames() As String
      Get
        Return mvClassFields(OrganisationGroupFields.CustomTableNames).Value
      End Get
    End Property
    Public ReadOnly Property ViewInContactCard As String
      Get
        Return mvClassFields(OrganisationGroupFields.ViewInContactCard).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Shared ReadOnly Property DefaultGroupCode() As String
      Get
        Return "ORG"
      End Get
    End Property

    Public Overrides ReadOnly Property EntityGroupType() As EntityGroupTypes
      Get
        Return EntityGroupTypes.egtOrganisation
      End Get
    End Property

    Public Overrides ReadOnly Property UseEventPricingMatrix() As Boolean
      Get
        Return False
      End Get
    End Property

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(OrganisationGroupFields.OrganisationGroup).Value = OrganisationGroup.DefaultGroupCode
      mvClassFields.Item(OrganisationGroupFields.OrganisationGroupDesc).Value = "Organisations"
      mvClassFields.Item(OrganisationGroupFields.Name).Value = "Organisation"
      mvClassFields.Item(OrganisationGroupFields.RgbValue).IntegerValue = &HFFFF00      'CYAN
      mvClassFields.Item(OrganisationGroupFields.TabPrefix).Value = "oci"
    End Sub

    Public Function SelectGroupsSQL() As SQLStatement
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("client", mvEnv.ClientCode)
      If mvEnv.GetConfigOption("option_contact_group_users") Then
        vWhereFields.Add("department", mvEnv.User.Department)
        vWhereFields.Add("og.organisation_group", CDBField.FieldTypes.cftInteger, "cgu.contact_group")
        Return New SQLStatement(mvEnv.Connection, GetRecordSetFields(), DatabaseTableName & " og, contact_group_users cgu", vWhereFields, "sequence_no")
      Else
        Return New SQLStatement(mvEnv.Connection, GetRecordSetFields(), mvClassFields.TableNameAndAlias, vWhereFields, "sequence_no")
      End If
    End Function

    Public Overrides ReadOnly Property DefaultGroup() As Boolean
      Get
        Return EntityGroupCode = DefaultGroupCode
      End Get
    End Property

    Public Overrides ReadOnly Property ViewOrganisationInContactCard As Boolean
      Get
        Return mvClassFields.Item(OrganisationGroupFields.ViewInContactCard).Bool
      End Get
    End Property

#End Region
  End Class
End Namespace
