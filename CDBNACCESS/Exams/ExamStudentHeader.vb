Imports Advanced.Data.Merge
Imports Advanced.Data.Merge.ComparisonStrategy
Imports Advanced.Data.Merge.Strategy

Namespace Access

  <MergeStrategy(MergeStrategyType.MergeData)>
  Public Class ExamStudentHeader
    Inherits CARERecord
    Implements IRecordCreate

    Private mvFirstSession As ExamSession
    Private mvLastSession As ExamSession
    Private mvContact As Contact
    Private mvExamStudentUnitHeaders As IEnumerable(Of ExamStudentUnitHeader)

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Friend Enum ExamStudentHeaderFields
      AllFields = 0
      ExamStudentHeaderId
      ExamUnitId
      ExamUnitLinkId
      ContactNumber
      FirstSessionId
      LastSessionId
      LastMarkedDate
      LastGradedDate
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
        .Add("exam_student_header_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftInteger)
        .Add("first_session_id", CDBField.FieldTypes.cftInteger)
        .Add("last_session_id", CDBField.FieldTypes.cftInteger)
        .Add("last_marked_date", CDBField.FieldTypes.cftTime)
        .Add("last_graded_date", CDBField.FieldTypes.cftTime)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamStudentHeaderFields.ExamStudentHeaderId).PrimaryKey = True
        .Item(ExamStudentHeaderFields.ExamStudentHeaderId).PrefixRequired = True
        .SetControlNumberField(ExamStudentHeaderFields.ExamStudentHeaderId, "XSH")

        .SetUniqueField(ExamStudentHeaderFields.ExamUnitId)
        .SetUniqueField(ExamStudentHeaderFields.ExamUnitLinkId)
        .SetUniqueField(ExamStudentHeaderFields.ContactNumber)
        .Item(ExamStudentHeaderFields.CreatedBy).PrefixRequired = True
        .Item(ExamStudentHeaderFields.CreatedOn).PrefixRequired = True
      End With

    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "esh"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_student_header"
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
      Return New ExamStudentHeader(mvEnv)
    End Function
    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ExamStudentHeaderId As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.ExamStudentHeaderId).IntegerValue
      End Get
    End Property

    <MergeComparer(1)>
    Public ReadOnly Property ExamUnitId As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property

    Public Property ContactNumber As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.ContactNumber).IntegerValue
      End Get
      Set(value As Integer)
        mvClassFields(ExamStudentHeaderFields.ContactNumber).IntegerValue = value
        If value = 0 OrElse (mvContact IsNot Nothing AndAlso Me.Contact.ContactNumber <> value) Then
          Me.Contact = Nothing
        End If
      End Set
    End Property

    <MergeParent()>
    Public Property Contact As Contact
      Get
        If mvContact Is Nothing Then Me.Contact = Me.GetRelatedInstance(Of Contact)({ExamStudentHeaderFields.ContactNumber})
        Return mvContact
      End Get
      Set(value As Contact)
        mvContact = value
        If value IsNot Nothing AndAlso Me.ContactNumber <> value.ContactNumber Then
          Me.ContactNumber = value.ContactNumber
        End If
      End Set
    End Property

    Public Property FirstSessionId As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.FirstSessionId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ExamStudentHeaderFields.FirstSessionId).IntegerValue = value
      End Set
    End Property

    Public Property LastSessionId As Integer
      Get
        Return mvClassFields(ExamStudentHeaderFields.LastSessionId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ExamStudentHeaderFields.LastSessionId).IntegerValue = value
      End Set
    End Property

    Public Property LastMarkedDate As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.LastMarkedDate).Value
      End Get
      Set(value As String)
        mvClassFields(ExamStudentHeaderFields.LastMarkedDate).Value = value
      End Set
    End Property

    <MergeField(MergeFieldComparisonOperator.IsGreatest)>
    Public Property LastMarkedDateAsDate As DateTime
      Get
        Dim vResult As DateTime = Nothing
        DateTime.TryParse(Me.LastMarkedDate, vResult)
        Return vResult
      End Get
      Set(value As DateTime)
        Me.LastMarkedDate = value.ToString(CAREDateTimeFormat())
      End Set
    End Property
    <MergeField(MergeFieldComparisonOperator.IsGreatest)> Public Property LastGradedDateAsDate As DateTime
      Get
        Dim vResult As DateTime = Nothing
        DateTime.TryParse(Me.LastGradedDate, vResult)
        Return vResult
      End Get
      Set(value As DateTime)
        Me.LastGradedDate = value.ToString(CAREDateTimeFormat())
      End Set
    End Property

    Public Property LastGradedDate As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.LastGradedDate).Value
      End Get
      Set(value As String)
        mvClassFields(ExamStudentHeaderFields.LastGradedDate).Value = value
      End Set
    End Property

    <MergeDependentField("CreatedOnAsDate")>
    Public Property CreatedBy As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.CreatedBy).Value
      End Get
      Set(value As String)
        mvClassFields(ExamStudentHeaderFields.CreatedBy).Value = value
      End Set
    End Property
    Public Property CreatedOn As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.CreatedOn).Value
      End Get
      Set(value As String)
        mvClassFields(ExamStudentHeaderFields.CreatedOn).Value = value
      End Set
    End Property
    <MergeField(MergeFieldComparisonOperator.IsSmallest)>
    Public Property CreatedOnAsDate As DateTime
      Get
        Dim vResult As DateTime = Nothing
        DateTime.TryParse(Me.CreatedOn, vResult)
        Return vResult
      End Get
      Set(value As DateTime)
        Me.CreatedOn = value.ToString(CAREDateFormat())
      End Set
    End Property
    Public ReadOnly Property AmendedBy As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn As String
      Get
        Return mvClassFields(ExamStudentHeaderFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"
    Sub ResetFirstLastSession()
      'NB: When setting the last session after a booking is created, the last session points to the last session that was booked, not the last session in chronological order.
      'When resetting the first and last session (currently after a booking is cancelled), we are attempting the reset the first and last session in the same way.

      Dim vInnerFields As String = "MIN(exam_bookings.exam_booking_id) FirstBookingId, MAX(exam_bookings.exam_booking_id) LastBookingId"

      Dim vInnerJoins As New AnsiJoins
      vInnerJoins.Add("exam_units", "exam_units.exam_unit_id", "exam_bookings.exam_unit_id")

      Dim vInnerWhereFields As New CDBFields
      vInnerWhereFields.Add("exam_bookings.contact_number", ContactNumber)
      vInnerWhereFields.Add("exam_units.exam_base_unit_id", ExamUnitId)
      vInnerWhereFields.Add("exam_units.exam_session_id", 0, CDBField.FieldWhereOperators.fwoNotEqual)
      vInnerWhereFields.Add("exam_bookings.cancellation_reason") 'ignore cancelled bookings
      Dim vFirstAndLastBookingSQL As New SQLStatement(mvEnv.Connection, vInnerFields, "exam_bookings", vInnerWhereFields, "", vInnerJoins)

      Dim vJoins As New AnsiJoins
      vJoins.Add("exam_bookings first_booking", "first_booking.exam_booking_id", "booking_aggregates.FirstBookingId")
      vJoins.Add("exam_bookings last_booking", "last_booking.exam_booking_id", "booking_aggregates.LastBookingId")

      Dim vFields As String = "first_booking.exam_session_id FirstSessionId, last_booking.exam_session_id LastSessionId"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "(" + vFirstAndLastBookingSQL.SQL + ") booking_aggregates", Nothing, "", vJoins)
      Dim vDT As DataTable = mvEnv.Connection.GetDataTable(vSQL)

      Dim vFirstSessionId As String = "" 'default value will be null.  Should only default if there are no bookings.
      Dim vLastSessionId As String = "" 'see above for vFirstSessionId
      If vDT.Rows.Count > 0 Then
        vFirstSessionId = vDT.Rows(0)("FirstSessionId").ToString()
        vLastSessionId = vDT.Rows(0)("LastSessionId").ToString()
      End If
      mvClassFields(ExamStudentHeaderFields.FirstSessionId).Value = vFirstSessionId
      mvClassFields(ExamStudentHeaderFields.LastSessionId).Value = vLastSessionId
      Me.Save()
    End Sub

    <MergeField(MergeFieldComparisonOperator.IsSmallest)>
    Public Property FirstSession As ExamSession
      Get
        If mvFirstSession Is Nothing AndAlso Me.FirstSessionId > 0 Then
          'Can't user the normal Me.GetRelatedInstance because the field is called first_session_id, and GetRelatedInstance expects exam_session_id
          Me.FirstSession = CARERecordFactory.Instance.GetInstanceByPrimaryKey(Of ExamSession)(Me.Environment, Me.FirstSessionId.ToString())
        End If
        Return mvFirstSession
      End Get
      Set(value As ExamSession)
        mvFirstSession = value
        Dim vFirstSessionId As Integer = If(value Is Nothing, 0, value.ExamSessionId)
        If Me.FirstSessionId <> vFirstSessionId Then
          Me.FirstSessionId = vFirstSessionId
        End If
      End Set
    End Property

    <MergeField(MergeFieldComparisonOperator.IsGreatest)>
    Public Property LastSession As ExamSession
      Get
        If mvLastSession Is Nothing AndAlso Me.LastSessionId > 0 Then
          'Can't user the normal Me.GetRelatedInstance because the field is called first_session_id, and GetRelatedInstance expects exam_session_id
          Me.LastSession = CARERecordFactory.Instance.GetInstanceByPrimaryKey(Of ExamSession)(Me.Environment, Me.LastSessionId.ToString())
        End If
        Return mvLastSession
      End Get
      Set(value As ExamSession)
        mvLastSession = value
        Dim vNewLastSessionId As Integer = If(value Is Nothing, 0, value.ExamSessionId)
        If Me.LastSessionId <> vNewLastSessionId Then
          Me.LastSessionId = vNewLastSessionId
        End If
      End Set
    End Property

    <MergeList()>
    Public ReadOnly Property ExamStudentUnitHeaders As IEnumerable(Of ExamStudentUnitHeader)
      Get
        If mvExamStudentUnitHeaders Is Nothing Then
          mvExamStudentUnitHeaders = Me.GetRelatedList(Of ExamStudentUnitHeader)({ExamStudentHeaderFields.ExamStudentHeaderId})
        End If
        Return mvExamStudentUnitHeaders
      End Get
    End Property

#End Region
  End Class
End Namespace
