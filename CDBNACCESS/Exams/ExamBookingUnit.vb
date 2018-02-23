Namespace Access

  Public Class ExamBookingUnit
    Inherits CARERecord
    Implements IRecordCreate

    Private mvExamUnit As ExamUnit
    Private mvExamUnitLink As ExamUnitLink

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ExamBookingUnitFields
      AllFields = 0
      ExamBookingUnitId
      ExamBookingId
      ExamUnitId
      ExamUnitLinkId
      ExamScheduleId
      ExamPersonnelId
      ContactNumber
      AddressNumber
      ExamCandidateNumber
      DeskNumber
      BatchNumber
      TransactionNumber
      LineNumber
      AttemptNumber
      ExamStudentUnitStatus
      RawMark
      OriginalMark
      ModeratedMark
      TotalMark
      OriginalGrade
      ModeratedGrade
      TotalGrade
      OriginalResult
      ModeratedResult
      TotalResult
      EntryDate
      ExpiryDate
      DoneDate
      ExpirySession
      CancellationReason
      CancellationSource
      CancelledBy
      CancelledOn
      CourseStartDate
      ExamAssessmentLanguage
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
        .Add("exam_booking_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_booking_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_schedule_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_personnel_id", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftInteger)
        .Add("address_number", CDBField.FieldTypes.cftInteger)
        .Add("exam_candidate_number")
        .Add("desk_number", CDBField.FieldTypes.cftInteger)
        .Add("batch_number", CDBField.FieldTypes.cftInteger)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("line_number", CDBField.FieldTypes.cftInteger)
        .Add("attempt_number", CDBField.FieldTypes.cftInteger)
        .Add("exam_student_unit_status")
        .Add("raw_mark", CDBField.FieldTypes.cftNumeric)
        .Add("original_mark", CDBField.FieldTypes.cftNumeric)
        .Add("moderated_mark", CDBField.FieldTypes.cftNumeric)
        .Add("total_mark", CDBField.FieldTypes.cftNumeric)
        .Add("original_grade")
        .Add("moderated_grade")
        .Add("total_grade")
        .Add("original_result")
        .Add("moderated_result")
        .Add("total_result")
        .Add("entry_date", CDBField.FieldTypes.cftDate)
        .Add("expiry_date", CDBField.FieldTypes.cftDate)
        .Add("done_date", CDBField.FieldTypes.cftDate)
        .Add("expiry_session", CDBField.FieldTypes.cftInteger)
        .Add("cancellation_reason")
        .Add("cancellation_source")
        .Add("cancelled_by")
        .Add("cancelled_on", CDBField.FieldTypes.cftDate)
        .Add("course_start_date", CDBField.FieldTypes.cftDate)
        .Add("exam_assessment_language")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamBookingUnitFields.ExamBookingUnitId).PrimaryKey = True
        .Item(ExamBookingUnitFields.ExamBookingUnitId).PrefixRequired = True
        .SetControlNumberField(ExamBookingUnitFields.ExamBookingUnitId, "XSU")

        .SetUniqueField(ExamBookingUnitFields.ExamBookingId)
        .SetUniqueField(ExamBookingUnitFields.ExamUnitId)
        .SetUniqueField(ExamBookingUnitFields.ExamUnitLinkId)
        .Item(ExamBookingUnitFields.ExamStudentUnitStatus).PrefixRequired = True
        .Item(ExamBookingUnitFields.CreatedBy).PrefixRequired = True
        .Item(ExamBookingUnitFields.CreatedOn).PrefixRequired = True

        .Item(ExamBookingUnitFields.CancellationReason).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation)
        .Item(ExamBookingUnitFields.CancellationSource).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation)
        .Item(ExamBookingUnitFields.CancelledBy).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation)
        .Item(ExamBookingUnitFields.CancelledOn).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation)
        .Item(ExamBookingUnitFields.CourseStartDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading)
        .Item(ExamBookingUnitFields.ExamAssessmentLanguage).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading)
        .Item(ExamBookingUnitFields.ExamUnitLinkId).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ebu"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_booking_units"
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
      Return New ExamBookingUnit(mvEnv)
    End Function
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ExamBookingUnitId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamBookingUnitId).IntegerValue
      End Get
    End Property
    Public Property ExamBookingId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamBookingId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ExamBookingUnitFields.ExamBookingId).IntegerValue = value
        If mvExamBooking IsNot Nothing AndAlso Me.ExamBooking.ExamBookingId <> value Then
          Me.ExamBooking = Nothing
        End If
      End Set
    End Property
    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamScheduleId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamScheduleId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamPersonnelId() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamPersonnelId).IntegerValue
      End Get
    End Property
    Public Property ContactNumber As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ContactNumber).IntegerValue
      End Get
      Friend Set(value As Integer)
        mvClassFields(ExamBookingUnitFields.ContactNumber).IntegerValue = value
      End Set
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCandidateNumber() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamCandidateNumber).Value
      End Get
    End Property
    Public ReadOnly Property DeskNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.DeskNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LineNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.LineNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AttemptNumber() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.AttemptNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamStudentUnitStatus() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.ExamStudentUnitStatus).Value
      End Get
    End Property

    Public ReadOnly Property RawMark() As Double
      Get
        Return mvClassFields(ExamBookingUnitFields.RawMark).DoubleValue
      End Get
    End Property

    Public ReadOnly Property OriginalMark() As Double
      Get
        Return mvClassFields(ExamBookingUnitFields.OriginalMark).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ModeratedMark() As Double
      Get
        Return mvClassFields(ExamBookingUnitFields.ModeratedMark).DoubleValue
      End Get
    End Property
    Public ReadOnly Property TotalMark() As Double
      Get
        Return mvClassFields(ExamBookingUnitFields.TotalMark).DoubleValue
      End Get
    End Property
    Public ReadOnly Property OriginalGrade() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.OriginalGrade).Value
      End Get
    End Property
    Public ReadOnly Property ModeratedGrade() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.ModeratedGrade).Value
      End Get
    End Property
    Public ReadOnly Property TotalGrade() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.TotalGrade).Value
      End Get
    End Property
    Public ReadOnly Property OriginalResult() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.OriginalResult).Value
      End Get
    End Property
    Public ReadOnly Property ModeratedResult() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.ModeratedResult).Value
      End Get
    End Property
    Public ReadOnly Property TotalResult() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.TotalResult).Value
      End Get
    End Property
    Public ReadOnly Property EntryDate() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.EntryDate).Value
      End Get
    End Property
    Public ReadOnly Property ExpiryDate() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.ExpiryDate).Value
      End Get
    End Property
    Public ReadOnly Property DoneDate() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.DoneDate).Value
      End Get
    End Property
    Public ReadOnly Property ExpirySession() As Integer
      Get
        Return mvClassFields(ExamBookingUnitFields.ExpirySession).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CancellationReason() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CancellationReason).Value
      End Get
    End Property
    Public ReadOnly Property CancellationSource() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CancelledOn).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamBookingUnitFields.AmendedOn).Value
      End Get
    End Property
#End Region

    Public Property ExamUnit As ExamUnit
      Get
        If mvExamUnit Is Nothing AndAlso Me.ExamUnitId > 0 Then
          Me.ExamUnit = Me.GetRelatedInstance(Of ExamUnit)({ExamBookingUnitFields.ExamUnitId})
        End If
        Return mvExamUnit
      End Get
      Private Set(value As ExamUnit)
        mvExamUnit = value
      End Set
    End Property
    Public Property ExamUnitLink As ExamUnitLink
      Get
        If mvExamUnitLink Is Nothing AndAlso Me.ExamUnitLinkId > 0 Then
          Me.ExamUnitLink = Me.GetRelatedInstance(Of ExamUnitLink)({ExamBookingUnitFields.ExamUnitLinkId})
        End If
        Return mvExamUnitLink
      End Get
      Private Set(value As ExamUnitLink)
        mvExamUnitLink = value
      End Set
    End Property

  End Class
End Namespace
