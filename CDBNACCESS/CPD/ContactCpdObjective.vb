Namespace Access

  Public Class ContactCpdObjective
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ContactCpdObjectiveFields
      AllFields = 0
      CpdObjectiveNumber
      ContactCpdPeriodNumber
      CpdObjectiveDesc
      CpdCategoryType
      CpdCategory
      CreatedBy
      CreatedOn
      SupervisorAccepted
      Notes
      CompletionDate
      TargetDate
      SupervisorContactNumber
      LongDescription
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("cpd_objective_number", CDBField.FieldTypes.cftLong)
        .Add("contact_cpd_period_number", CDBField.FieldTypes.cftLong)
        .Add("cpd_objective_desc")
        .Add("cpd_category_type")
        .Add("cpd_category")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)
        .Add("supervisor_accepted")
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("completion_date", CDBField.FieldTypes.cftDate)
        .Add("target_date", CDBField.FieldTypes.cftDate)
        .Add("supervisor_contact_number", CDBField.FieldTypes.cftLong)
        .Add("long_description", CDBField.FieldTypes.cftMemo)

        .Item(ContactCpdObjectiveFields.CpdObjectiveNumber).PrimaryKey = True

        .SetControlNumberField(ContactCpdObjectiveFields.CpdObjectiveNumber, "OJ")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cco"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_cpd_objectives"
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
    Public ReadOnly Property CpdObjectiveNumber() As Integer
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CpdObjectiveNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactCpdPeriodNumber() As Integer
      Get
        Return mvClassFields(ContactCpdObjectiveFields.ContactCpdPeriodNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CpdObjectiveDesc() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CpdObjectiveDesc).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategoryType() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CpdCategoryType).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategory() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CpdCategory).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property SupervisorAccepted() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.SupervisorAccepted).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property CompletionDate() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.CompletionDate).Value
      End Get
    End Property
    Public ReadOnly Property TargetDate() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.TargetDate).Value
      End Get
    End Property
    Public ReadOnly Property SupervisorContactNumber() As Integer
      Get
        Return mvClassFields(ContactCpdObjectiveFields.SupervisorContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LongDescription() As String
      Get
        Return mvClassFields(ContactCpdObjectiveFields.LongDescription).Value
      End Get
    End Property
#End Region


  End Class
End Namespace
