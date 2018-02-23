Namespace Access

  Public Class MaintenanceAttribute
    Inherits CARERecord

    Protected Enum MaintenanceAttributeFields
      AllFields = 0
      AttributeName
      AttributeNameDesc
      TableName
      Type
      EntryLength
      ToCase
      NullsInvalid
      MinimumValue
      MaximumValue
      DomainValues
      Pattern
      ValidationTable
      ValidationAttribute
      RestrictionAttribute
      Maintenance
      PrimaryKey
      SequenceNumber
      AttributeNotes
    End Enum

    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("attribute_name").PrimaryKey = True
        .Add("attribute_name_desc")
        .Add("table_name").PrimaryKey = True
        .Add("type")
        .Add("entry_length", CDBField.FieldTypes.cftInteger)
        .Add("case").SpecialColumn = True
        .Add("nulls_invalid")
        .Add("minimum_value")
        .Add("maximum_value")
        .Add("domain_values", CDBField.FieldTypes.cftMemo)
        .Add("pattern")
        .Add("validation_table")
        .Add("validation_attribute")
        .Add("restriction_attribute")
        .Add("maintenance")
        .Add("primary_key")
        .Add("sequence_number", CDBField.FieldTypes.cftInteger)
        .Add("attribute_notes", CDBField.FieldTypes.cftMemo)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ma"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "maintenance_attributes"
      End Get
    End Property

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '---------------------------------------------------------------------------------------
    'Public Property procedures follow
    '---------------------------------------------------------------------------------------

    Public ReadOnly Property TableName() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.TableName).Value
      End Get
    End Property

    Public ReadOnly Property AttributeName() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.AttributeName).Value
      End Get
    End Property

    Public Property AttributeNameDesc() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.AttributeNameDesc).Value
      End Get
      Set(ByVal pValue As String)
        mvClassFields(MaintenanceAttributeFields.AttributeNameDesc).Value = pValue
      End Set
    End Property

    Public ReadOnly Property FieldType() As CDBField.FieldTypes
      Get
        Return CDBField.GetFieldType(mvClassFields(MaintenanceAttributeFields.Type).Value)
      End Get
    End Property

    Public ReadOnly Property PrimaryKey() As Boolean
      Get
        Return mvClassFields(MaintenanceAttributeFields.PrimaryKey).Bool
      End Get
    End Property

    Public ReadOnly Property DataType() As String
      Get
        Select Case FieldType
          Case CDBField.FieldTypes.cftLong, CDBField.FieldTypes.cftInteger
            Return "Integer"
          Case CDBField.FieldTypes.cftNumeric
            Return "Double"
          Case Else
            Return "String"
        End Select
      End Get
    End Property

    Public ReadOnly Property EntryLength() As Integer
      Get
        Return mvClassFields(MaintenanceAttributeFields.EntryLength).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MaximumValue() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.MaximumValue).Value
      End Get
    End Property

    Public ReadOnly Property MinimumValue() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.MinimumValue).Value
      End Get
    End Property

    Public ReadOnly Property NullsInvalid() As Boolean
      Get
        Return mvClassFields(MaintenanceAttributeFields.NullsInvalid).Bool
      End Get
    End Property

    Public ReadOnly Property ToCase() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.ToCase).Value
      End Get
    End Property

    Public ReadOnly Property Type() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.Type).Value
      End Get
    End Property

    Public ReadOnly Property DomainValues() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.DomainValues).Value
      End Get
    End Property
    Public ReadOnly Property Pattern() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.Pattern).Value
      End Get
    End Property
    Public ReadOnly Property RestrictionAttribute() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.RestrictionAttribute).Value
      End Get
    End Property
    Public ReadOnly Property ValidationTable() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.ValidationTable).Value
      End Get
    End Property
    Public ReadOnly Property ValidationAttribute() As String
      Get
        Return mvClassFields(MaintenanceAttributeFields.ValidationAttribute).Value
      End Get
    End Property


    Public Sub InitForAdditionalInfo(ByVal pNeedCase As Boolean)
      MyBase.Init()
      mvClassFields(MaintenanceAttributeFields.TableName).InDatabase = False
      mvClassFields(MaintenanceAttributeFields.AttributeName).InDatabase = False
      mvClassFields(MaintenanceAttributeFields.AttributeNotes).InDatabase = False
      mvClassFields(MaintenanceAttributeFields.SequenceNumber).InDatabase = False
      If pNeedCase = False Then
        mvClassFields(MaintenanceAttributeFields.ToCase).InDatabase = False
      End If
    End Sub

  End Class
End Namespace
