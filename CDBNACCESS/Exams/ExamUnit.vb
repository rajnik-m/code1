Namespace Access

  Public Class ExamUnit
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ExamUnitFields
      AllFields = 0
      ExamUnitId
      ExamSessionId
      ExamUnitCode
      ExamUnitDescription
      ExamBaseUnitId
      Subject
      SkillLevel
      ExamUnitType
      ExamUnitStatus
      SequenceNumber
      SessionBased
      ValidFrom
      ValidTo
      Product
      Rate
      DateApproved
      RegistrationDate
      QcfLevel
      NumberOfCredits
      NvqCode
      SvqCode
      UnitTimeLimit
      TimeLimitType
      MinimumStudents
      MaximumStudents
      StudentCount
      MinimumAge
      AllowBookings
      ExamMarkType
      MarkFactor
      AwardingBody
      ExamUnitReplacedById
      AllowExemptions
      ExemptionMark
      ExamMarkerStatus
      PapersPerMarker
      IncludeView
      ExcludeView
      AllowDowngrade
      WebPublish
      Notes
      ActivityGroup
      IsGradingEndpoint
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
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_session_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_code")
        .Add("exam_unit_description")
        .Add("exam_base_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("subject")
        .Add("skill_level")
        .Add("exam_unit_type")
        .Add("exam_unit_status")
        .Add("sequence_number", CDBField.FieldTypes.cftInteger)
        .Add("session_based")
        .Add("valid_from", CDBField.FieldTypes.cftDate)
        .Add("valid_to", CDBField.FieldTypes.cftDate)
        .Add("product")
        .Add("rate")
        .Add("date_approved", CDBField.FieldTypes.cftDate)
        .Add("registration_date", CDBField.FieldTypes.cftDate)
        .Add("qcf_level", CDBField.FieldTypes.cftInteger)
        .Add("number_of_credits", CDBField.FieldTypes.cftNumeric)
        .Add("nvq_code")
        .Add("svq_code")
        .Add("unit_time_limit", CDBField.FieldTypes.cftInteger)
        .Add("time_limit_type")
        .Add("minimum_students", CDBField.FieldTypes.cftInteger)
        .Add("maximum_students", CDBField.FieldTypes.cftInteger)
        .Add("student_count", CDBField.FieldTypes.cftInteger)
        .Add("minimum_age", CDBField.FieldTypes.cftInteger)
        .Add("allow_bookings")
        .Add("exam_mark_type")
        .Add("mark_factor", CDBField.FieldTypes.cftNumeric)
        .Add("awarding_body", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_replaced_by_id", CDBField.FieldTypes.cftInteger)
        .Add("allow_exemptions")
        .Add("exemption_mark", CDBField.FieldTypes.cftNumeric)
        .Add("exam_marker_status")
        .Add("papers_per_marker", CDBField.FieldTypes.cftInteger)
        .Add("include_view")
        .Add("exclude_view")
        .Add("allow_downgrade")
        .Add("web_publish")
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("activity_group")
        .Add("is_grading_endpoint")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)


        .Item(ExamUnitFields.ExamUnitId).PrimaryKey = True
        .Item(ExamUnitFields.ExamUnitId).PrefixRequired = True
        .SetControlNumberField(ExamUnitFields.ExamUnitId, "XU")


        .SetUniqueField(ExamUnitFields.ExamSessionId)
        .SetUniqueField(ExamUnitFields.ExamUnitCode)
        .Item(ExamUnitFields.Subject).PrefixRequired = True
        .Item(ExamUnitFields.SkillLevel).PrefixRequired = True
        .Item(ExamUnitFields.ExamUnitType).PrefixRequired = True
        .Item(ExamUnitFields.ExamUnitStatus).PrefixRequired = True
        .Item(ExamUnitFields.Product).PrefixRequired = True
        .Item(ExamUnitFields.Rate).PrefixRequired = True
        .Item(ExamUnitFields.CreatedBy).PrefixRequired = True
        .Item(ExamUnitFields.CreatedOn).PrefixRequired = True

      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "eu"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_units"
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
      Return New ExamUnit(mvEnv)
    End Function

    '--------------------------------------------------
    'AddDeleteCheckItems
    '--------------------------------------------------
    Public Overrides Sub AddDeleteCheckItems()
      AddDeleteCheckItem("exam_booking_units", "exam_unit_id", "an Exam Booking")
      AddDeleteCheckItem("exam_schedule", "exam_unit_id", "an Exam Schedule")
      AddDeleteCheckItem("exam_student_unit_header", "exam_unit_id", "a Student Summary")
      AddDeleteCheckItem("exam_units", "exam_base_unit_id", "a Session Unit")
      AddDeleteCheckItem("exam_unit_links", "exam_unit_id_1", "a Child Unit")
      'There is an additional check that is more complex and cannot be handled by AddDeleteCheckItem.  See GetFirstDeleteCheckItem overridden method

      AddCascadeDeleteItem("exam_unit_assessment_types", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_eligibility_checks", "exam_unit_id")
      AddCascadeDeleteItem("exam_centre_units", "exam_unit_id")
      AddCascadeDeleteItem("exam_exemption_units", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_grades", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_personnel", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_prerequisites", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_products", "exam_unit_id")
      AddCascadeDeleteItem("exam_unit_links", "exam_unit_id_2")
    End Sub

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvExamUnitLink = Nothing
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamSessionId() As Integer
      Get
        Return mvClassFields(ExamUnitFields.ExamSessionId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitCode() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitCode).Value
      End Get
    End Property
    Public ReadOnly Property ExamUnitDescription() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitDescription).Value
      End Get
    End Property
    Public ReadOnly Property ExamBaseUnitId() As Integer
      Get
        Return mvClassFields(ExamUnitFields.ExamBaseUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Subject() As String
      Get
        Return mvClassFields(ExamUnitFields.Subject).Value
      End Get
    End Property
    Public ReadOnly Property SkillLevel() As String
      Get
        Return mvClassFields(ExamUnitFields.SkillLevel).Value
      End Get
    End Property
    Public ReadOnly Property ExamUnitType() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitType).Value
      End Get
    End Property
    Public ReadOnly Property ExamUnitStatus() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitStatus).Value
      End Get
    End Property
    Public ReadOnly Property SequenceNumber() As Integer
      Get
        Return mvClassFields(ExamUnitFields.SequenceNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property SessionBased() As Boolean
      Get
        Return mvClassFields(ExamUnitFields.SessionBased).Bool
      End Get
    End Property
    Public ReadOnly Property ValidFrom() As String
      Get
        Return mvClassFields(ExamUnitFields.ValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property ValidTo() As String
      Get
        Return mvClassFields(ExamUnitFields.ValidTo).Value
      End Get
    End Property
    Public ReadOnly Property ProductCode() As String
      Get
        Return mvClassFields(ExamUnitFields.Product).Value
      End Get
    End Property
    Public ReadOnly Property RateCode() As String
      Get
        Return mvClassFields(ExamUnitFields.Rate).Value
      End Get
    End Property
    Public ReadOnly Property DateApproved() As String
      Get
        Return mvClassFields(ExamUnitFields.DateApproved).Value
      End Get
    End Property
    Public ReadOnly Property RegistrationDate() As String
      Get
        Return mvClassFields(ExamUnitFields.RegistrationDate).Value
      End Get
    End Property
    Public ReadOnly Property QcfLevel() As Integer
      Get
        Return mvClassFields(ExamUnitFields.QcfLevel).IntegerValue
      End Get
    End Property
    Public ReadOnly Property NumberOfCredits() As Double
      Get
        Return mvClassFields(ExamUnitFields.NumberOfCredits).DoubleValue
      End Get
    End Property
    Public ReadOnly Property NvqCode() As String
      Get
        Return mvClassFields(ExamUnitFields.NvqCode).Value
      End Get
    End Property
    Public ReadOnly Property SvqCode() As String
      Get
        Return mvClassFields(ExamUnitFields.SvqCode).Value
      End Get
    End Property
    Public ReadOnly Property UnitTimeLimit() As Integer
      Get
        Return mvClassFields(ExamUnitFields.UnitTimeLimit).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TimeLimitType() As String
      Get
        Return mvClassFields(ExamUnitFields.TimeLimitType).Value
      End Get
    End Property
    Public ReadOnly Property MinimumStudents() As Integer
      Get
        Return mvClassFields(ExamUnitFields.MinimumStudents).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MaximumStudents() As Integer
      Get
        Return mvClassFields(ExamUnitFields.MaximumStudents).IntegerValue
      End Get
    End Property
    Public ReadOnly Property StudentCount() As Integer
      Get
        Return mvClassFields(ExamUnitFields.StudentCount).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MinimumAge() As Integer
      Get
        Return mvClassFields(ExamUnitFields.MinimumAge).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AllowBookings() As Boolean
      Get
        Return mvClassFields(ExamUnitFields.AllowBookings).Bool
      End Get
    End Property
    Public ReadOnly Property ExamMarkType() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamMarkType).Value
      End Get
    End Property
    Public ReadOnly Property MarkFactor() As Double
      Get
        Return mvClassFields(ExamUnitFields.MarkFactor).DoubleValue
      End Get
    End Property
    Public ReadOnly Property AwardingBody() As Integer
      Get
        Return mvClassFields(ExamUnitFields.AwardingBody).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitReplacedById() As Integer
      Get
        Return mvClassFields(ExamUnitFields.ExamUnitReplacedById).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AllowExemptions() As Boolean
      Get
        Return mvClassFields(ExamUnitFields.AllowExemptions).Bool
      End Get
    End Property
    Public ReadOnly Property ExemptionMark() As Double
      Get
        Return mvClassFields(ExamUnitFields.ExemptionMark).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ExamMarkerStatus() As String
      Get
        Return mvClassFields(ExamUnitFields.ExamMarkerStatus).Value
      End Get
    End Property
    Public ReadOnly Property PapersPerMarker() As Integer
      Get
        Return mvClassFields(ExamUnitFields.PapersPerMarker).IntegerValue
      End Get
    End Property
    Public ReadOnly Property IncludeView() As String
      Get
        Return mvClassFields(ExamUnitFields.IncludeView).Value
      End Get
    End Property
    Public ReadOnly Property ExcludeView() As String
      Get
        Return mvClassFields(ExamUnitFields.ExcludeView).Value
      End Get
    End Property
    Public ReadOnly Property AllowDowngrade() As String
      Get
        Return mvClassFields(ExamUnitFields.AllowDowngrade).Value
      End Get
    End Property
    Public ReadOnly Property WebPublish() As Boolean
      Get
        Return mvClassFields(ExamUnitFields.WebPublish).Bool
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(ExamUnitFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property ActivityGroup() As String
      Get
        Return mvClassFields(ExamUnitFields.ActivityGroup).Value
      End Get
    End Property
    Public ReadOnly Property IsGradingEndpoint() As Boolean
      Get
        Return mvClassFields(ExamUnitFields.IsGradingEndpoint).Bool
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamUnitFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamUnitFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamUnitFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamUnitFields.AmendedOn).Value
      End Get
    End Property

#End Region
  End Class
End Namespace