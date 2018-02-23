Namespace Access

  Public Class ExamStudentEligibility
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ExamStudentEligibilityFields
      AllFields = 0
      ExamUnitEligibilityCheckId
      ContactNumber
      Proven
      ProvedDate
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
        .Add("exam_unit_eligibility_check_id", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftInteger)
        .Add("proven")
        .Add("proved_date", CDBField.FieldTypes.cftDate)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamStudentEligibilityFields.ExamUnitEligibilityCheckId).PrimaryKey = True
        .Item(ExamStudentEligibilityFields.ExamUnitEligibilityCheckId).PrefixRequired = True
        .SetControlNumberField(ExamStudentEligibilityFields.ExamUnitEligibilityCheckId, "XEC")


        .Item(ExamStudentEligibilityFields.ContactNumber).PrimaryKey = True
        .Item(ExamStudentEligibilityFields.ContactNumber).PrefixRequired = True
        .SetControlNumberField(ExamStudentEligibilityFields.ContactNumber, "XEC")

        .Item(ExamStudentEligibilityFields.CreatedBy).PrefixRequired = True
        .Item(ExamStudentEligibilityFields.CreatedOn).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ese"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_student_eligibility"
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
      Return New ExamStudentEligibility(mvEnv)
    End Function
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ExamUnitEligibilityCheckId() As Integer
      Get
        Return mvClassFields(ExamStudentEligibilityFields.ExamUnitEligibilityCheckId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ExamStudentEligibilityFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Proven() As Boolean
      Get
        Return mvClassFields(ExamStudentEligibilityFields.Proven).Bool
      End Get
    End Property
    Public ReadOnly Property ProvedDate() As String
      Get
        Return mvClassFields(ExamStudentEligibilityFields.ProvedDate).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamStudentEligibilityFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamStudentEligibilityFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamStudentEligibilityFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamStudentEligibilityFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace