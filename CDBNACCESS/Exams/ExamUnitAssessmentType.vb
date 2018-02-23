Namespace Access

  Public Class ExamUnitAssessmentType
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ExamUnitAssessmentTypeFields
      AllFields = 0
      ExamUnitAssessmentTypeId
      ExamUnitId
      ExamAssessmentType
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("exam_unit_assessment_type_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_assessment_type")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamUnitAssessmentTypeFields.ExamUnitAssessmentTypeId).PrimaryKey = True
        .Item(ExamUnitAssessmentTypeFields.ExamUnitAssessmentTypeId).PrefixRequired = True
        .SetControlNumberField(ExamUnitAssessmentTypeFields.ExamUnitAssessmentTypeId, "XAT")

        .SetUniqueField(ExamUnitAssessmentTypeFields.ExamUnitId)
        .Item(ExamUnitAssessmentTypeFields.ExamAssessmentType).PrefixRequired = True
        .SetUniqueField(ExamUnitAssessmentTypeFields.ExamAssessmentType)
        .Item(ExamUnitAssessmentTypeFields.CreatedBy).PrefixRequired = True
        .Item(ExamUnitAssessmentTypeFields.CreatedOn).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "euat"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_unit_assessment_types"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'IRecordCreate
'--------------------------------------------------
    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ExamUnitAssessmentType(mvEnv)
    End Function
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ExamUnitAssessmentTypeId() As Integer
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.ExamUnitAssessmentTypeId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamAssessmentType() As String
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.ExamAssessmentType).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamUnitAssessmentTypeFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGeneratedCode"

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      If mvClassFields("exam_assessment_type").ValueChanged Then
        mvClassFields.CheckRecordExists(mvEnv)
      End If
    End Sub
#End Region

  End Class
End Namespace