Imports System.Linq
Namespace Access

  Public Class ExamBooking
    Inherits CARERecord
    Implements IRecordCreate

    Private mvExamSession As ExamSession

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ExamBookingFields
      AllFields = 0
      ExamBookingId
      ContactNumber
      AddressNumber
      ExamSessionId
      ExamCentreId
      ExamUnitId
      ExamUnitLinkId
      Amount
      BatchNumber
      TransactionNumber
      CancellationReason
      CancellationSource
      CancelledBy
      CancelledOn
      SpecialRequirements
      StudyMode
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
        .Add("exam_booking_id", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftInteger)
        .Add("address_number", CDBField.FieldTypes.cftInteger)
        .Add("exam_session_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("amount", CDBField.FieldTypes.cftNumeric)
        .Add("batch_number", CDBField.FieldTypes.cftInteger)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("cancellation_reason")
        .Add("cancellation_source")
        .Add("cancelled_by")
        .Add("cancelled_on", CDBField.FieldTypes.cftDate)
        .Add("special_requirements")
        .Add("study_mode")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamBookingFields.ExamBookingId).PrimaryKey = True
        .Item(ExamBookingFields.ExamBookingId).PrefixRequired = True
        .SetControlNumberField(ExamBookingFields.ExamBookingId, "XBK")

        .Item(ExamBookingFields.CreatedBy).PrefixRequired = True
        .Item(ExamBookingFields.CreatedOn).PrefixRequired = True
        .Item(ExamBookingFields.BatchNumber).PrefixRequired = True
        .Item(ExamBookingFields.TransactionNumber).PrefixRequired = True
        .Item(ExamBookingFields.CancellationReason).PrefixRequired = True
        .Item(ExamBookingFields.CancellationSource).PrefixRequired = True
        .Item(ExamBookingFields.StudyMode).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "eb"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_bookings"
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
      Return New ExamBooking(mvEnv)
    End Function

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ExamBookingId() As Integer
      Get
        Return mvClassFields(ExamBookingFields.ExamBookingId).IntegerValue
      End Get
    End Property
    Public Property ContactNumber As Integer
      Get
        Return mvClassFields(ExamBookingFields.ContactNumber).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ExamBookingFields.ContactNumber).IntegerValue = value
        If value = 0 OrElse (mvContact IsNot Nothing AndAlso Me.Contact.ContactNumber <> value) Then
          Me.Contact = Nothing
        End If
        Me.ExamBookingUnits.ToList().ForEach(Sub(vEBU) vEBU.ContactNumber = Me.ContactNumber)
      End Set
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(ExamBookingFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamSessionId() As Integer
      Get
        Return mvClassFields(ExamBookingFields.ExamSessionId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ExamBookingFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields(ExamBookingFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamBookingFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamBookingFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamBookingFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamBookingFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamBookingFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property Amount() As Double
      Get
        Return mvClassFields(ExamBookingFields.Amount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(ExamBookingFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(ExamBookingFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CancellationReason() As String
      Get
        Return mvClassFields(ExamBookingFields.CancellationReason).Value
      End Get
    End Property
    Public ReadOnly Property CancellationSource() As String
      Get
        Return mvClassFields(ExamBookingFields.CancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        Return mvClassFields(ExamBookingFields.CancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn() As String
      Get
        Return mvClassFields(ExamBookingFields.CancelledOn).Value
      End Get
    End Property
    Public ReadOnly Property SpecialRequirements() As Boolean
      Get
        Return mvClassFields(ExamBookingFields.SpecialRequirements).Bool
      End Get
    End Property
    Public ReadOnly Property StudyMode() As String
      Get
        Return mvClassFields(ExamBookingFields.StudyMode).Value
      End Get
    End Property

    Public Property ExamSession As ExamSession
      Get
        If mvExamSession Is Nothing AndAlso Me.ExamSessionId > 0 Then
          Me.ExamSession = Me.GetRelatedInstance(Of ExamSession)({ExamBookingFields.ExamSessionId})
        End If
        Return mvExamSession
      End Get
      Set(value As ExamSession)
        mvExamSession = value
      End Set
    End Property


#End Region

  End Class
End Namespace
