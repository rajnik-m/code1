

Namespace Access
  Public Class EventDelegate

    Public Enum EventDelegateRecordSetTypes 'These are bit values
      edrtNumbers = &H1S
      edrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventDelegateFields
      edfAll = 0
      edfEventNumber
      edfBookingNumber
      edfContactNumber
      edfAddressNumber
      edfAttended
      edfPosition
      edfOrganisationName
      edfAmendedBy
      edfAmendedOn
      edfEventDelegateNumber
      edfCandidateNumber
      edfPledgedAmount
      edfDonationTotal
      edfSponsorshipTotal
      edfBookingPaymentAmount
      edfOtherPaymentsTotal
      edfSequenceNumber
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvContact As Contact
    Private mvActivities As CDBCollection
    Private mvLinks As CollectionList(Of DelegateLink)

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      Dim vFinancialAnalysis As Boolean

      vFinancialAnalysis = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis)

      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "delegates"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("booking_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("attended")
          .Add("position")
          .Add("organisation_name")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("event_delegate_number", CDBField.FieldTypes.cftLong)
          .Add("candidate_number", CDBField.FieldTypes.cftLong)
          .Add("pledged_amount", CDBField.FieldTypes.cftNumeric)
          .Add("donation_total", CDBField.FieldTypes.cftNumeric)
          .Add("sponsorship_total", CDBField.FieldTypes.cftNumeric)
          .Add("booking_payment_amount", CDBField.FieldTypes.cftNumeric)
          .Add("other_payments_total", CDBField.FieldTypes.cftNumeric)
          .Add("sequence_number", CDBField.FieldTypes.cftLong)
          .Item(EventDelegateFields.edfPledgedAmount).InDatabase = vFinancialAnalysis
          .Item(EventDelegateFields.edfDonationTotal).InDatabase = vFinancialAnalysis
          .Item(EventDelegateFields.edfSponsorshipTotal).InDatabase = vFinancialAnalysis
          .Item(EventDelegateFields.edfBookingPaymentAmount).InDatabase = vFinancialAnalysis
          .Item(EventDelegateFields.edfOtherPaymentsTotal).InDatabase = vFinancialAnalysis

          .Item(EventDelegateFields.edfSequenceNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDelegateSequenceNumber)
        End With

        mvClassFields.Item(EventDelegateFields.edfEventNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventDelegateFields.edfBookingNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventDelegateFields.edfContactNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(EventDelegateFields.edfEventNumber).PrefixRequired = True
        mvClassFields.Item(EventDelegateFields.edfBookingNumber).PrefixRequired = True
        mvClassFields.Item(EventDelegateFields.edfContactNumber).PrefixRequired = True
        mvClassFields.Item(EventDelegateFields.edfAddressNumber).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvContact = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventDelegateFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(EventDelegateFields.edfAttended).Value = "" Then mvClassFields.Item(EventDelegateFields.edfAttended).Bool = False
      mvClassFields.Item(EventDelegateFields.edfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventDelegateFields.edfAmendedBy).Value = mvEnv.User.UserID
      If mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).InDatabase And mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).IntegerValue = 0 Then
        mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).Value = CStr(mvEnv.GetControlNumber("ED"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventDelegateRecordSetTypes) As String
      Dim vFields As String = ""

      'Modify below to add each recordset type as required
      If pRSType = EventDelegateRecordSetTypes.edrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ed")
      Else
        If (pRSType And EventDelegateRecordSetTypes.edrtNumbers) = EventDelegateRecordSetTypes.edrtNumbers Then
          vFields = "ed.event_number,ed.booking_number,ed.contact_number,ed.address_number,ed.attended,event_delegate_number,candidate_number,ed.sequence_number"
        End If
      End If
      Return vFields
    End Function

    Public Sub InitForImport(ByVal pEnv As CDBEnvironment, ByRef pContact As Contact, Optional ByVal pPledgedAmount As Double = 0)
      mvEnv = pEnv
      mvContact = pContact
      mvClassFields.Item(EventDelegateFields.edfPledgedAmount).Value = CStr(pPledgedAmount)
    End Sub
    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pBookingNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0, Optional ByRef pEventDelegateNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      'The original WHERE clause was constructed with the expectation that pEventNumber, pBookingNumber & pContactNumber would all be supplied, so an Assert was added as a coding integrity check
      If pEventNumber > 0 Or pBookingNumber > 0 Or pContactNumber > 0 Then System.Diagnostics.Debug.Assert(pEventNumber > 0 And pBookingNumber > 0 And pContactNumber > 0, "")

      mvEnv = pEnv
      If (pEventNumber > 0 And pBookingNumber > 0 And pContactNumber > 0) Or pEventDelegateNumber > 0 Then
        With vWhereFields
          If pEventDelegateNumber > 0 Then
            .Add("event_delegate_number", CDBField.FieldTypes.cftLong, pEventDelegateNumber)
          Else
            .Add("event_number", CDBField.FieldTypes.cftLong, pEventNumber)
            .Add("booking_number", CDBField.FieldTypes.cftLong, pBookingNumber)
            .Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
          End If
        End With
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventDelegateRecordSetTypes.edrtAll) & " FROM delegates ed WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventDelegateRecordSetTypes.edrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub
    Public Sub CalculateTotals()
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vNotInSQL As String

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        '(1) Donation total
        mvClassFields(EventDelegateFields.edfDonationTotal).Value = CStr(0)
        vWhereFields.Add("fh.contact_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfContactNumber).IntegerValue)
        vWhereFields.Add("fhd.batch_number", CDBField.FieldTypes.cftLong, "fh.batch_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("fhd.transaction_number", CDBField.FieldTypes.cftLong, "fh.transaction_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.product", CDBField.FieldTypes.cftLong, "fhd.product", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.donation", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("es.event_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfEventNumber).IntegerValue)
        vWhereFields.Add("es.source", CDBField.FieldTypes.cftLong, "fhd.source")
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM(fhd.amount)  AS  fhd_total FROM financial_history fh, financial_history_details fhd, products p, event_sources es WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          mvClassFields(EventDelegateFields.edfDonationTotal).Value = FixedFormat(vRecordSet.Fields(1).DoubleValue)
        End If
        vRecordSet.CloseRecordSet()

        '(2) Sponsorship Total
        vWhereFields.Clear()
        mvClassFields(EventDelegateFields.edfSponsorshipTotal).Value = CStr(0)
        vWhereFields.Add("fh.contact_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfContactNumber).IntegerValue)
        vWhereFields.Add("fhd.batch_number", CDBField.FieldTypes.cftLong, "fh.batch_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("fhd.transaction_number", CDBField.FieldTypes.cftLong, "fh.transaction_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.product", CDBField.FieldTypes.cftLong, "fhd.product", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.sponsorship_event", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("es.event_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfEventNumber).IntegerValue)
        vWhereFields.Add("es.source", CDBField.FieldTypes.cftLong, "fhd.source")
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM(fhd.amount)  AS  fhd_total FROM financial_history fh, financial_history_details fhd, products p, event_sources es WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          mvClassFields(EventDelegateFields.edfSponsorshipTotal).Value = FixedFormat(vRecordSet.Fields(1).DoubleValue)
        End If
        vRecordSet.CloseRecordSet()

        '(3) Booking Payment Amount
        vWhereFields.Clear()
        mvClassFields(EventDelegateFields.edfBookingPaymentAmount).Value = CStr(0)
        vWhereFields.Add("eb.booking_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfBookingNumber).IntegerValue)
        vWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("fh.batch_number", CDBField.FieldTypes.cftLong, "eb.batch_number")
        vWhereFields.Add("fh.transaction_number", CDBField.FieldTypes.cftLong, "eb.transaction_number")
        vWhereFields.Add("fhd.batch_number", CDBField.FieldTypes.cftLong, "fh.batch_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("fhd.transaction_number", CDBField.FieldTypes.cftLong, "fh.transaction_number", CDBField.FieldWhereOperators.fwoEqual)
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT fhd.amount,eb.quantity FROM event_bookings eb,financial_history fh, financial_history_details fhd WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          mvClassFields(EventDelegateFields.edfBookingPaymentAmount).Value = FixedFormat(vRecordSet.Fields(1).DoubleValue / vRecordSet.Fields(2).IntegerValue)
        End If
        vRecordSet.CloseRecordSet()

        '(4) Other Payments Total
        vWhereFields.Clear()
        mvClassFields(EventDelegateFields.edfOtherPaymentsTotal).Value = CStr(0)
        vWhereFields.Add("fh.contact_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfContactNumber).IntegerValue)
        vWhereFields.Add("fhd.batch_number", CDBField.FieldTypes.cftLong, "fh.batch_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("fhd.transaction_number", CDBField.FieldTypes.cftLong, "fh.transaction_number", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.product", CDBField.FieldTypes.cftLong, "fhd.product", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("p.donation", CDBField.FieldTypes.cftCharacter, "N", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("sponsorship_event", CDBField.FieldTypes.cftCharacter, "N", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("es.event_number", CDBField.FieldTypes.cftLong, mvClassFields(EventDelegateFields.edfEventNumber).IntegerValue)
        vWhereFields.Add("es.source", CDBField.FieldTypes.cftLong, "fhd.source")
        vNotInSQL = "SELECT batch_number from event_bookings eb WHERE eb.batch_number = fhd.batch_number AND eb.transaction_number = fhd.transaction_number AND eb.line_number = fhd.line_number"
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM(fhd.amount)  AS  fhd_total FROM financial_history fh, financial_history_details fhd, products p, event_sources es WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND fh.batch_number NOT IN (" & vNotInSQL & ")")
        If vRecordSet.Fetch() = True Then
          mvClassFields(EventDelegateFields.edfOtherPaymentsTotal).Value = FixedFormat(vRecordSet.Fields(1).DoubleValue)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventDelegateRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventDelegateFields.edfEventNumber, vFields)
        .SetItem(EventDelegateFields.edfBookingNumber, vFields)
        .SetItem(EventDelegateFields.edfContactNumber, vFields)
        .SetItem(EventDelegateFields.edfAttended, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventDelegateRecordSetTypes.edrtNumbers) = EventDelegateRecordSetTypes.edrtNumbers Then
          .SetItem(EventDelegateFields.edfAddressNumber, vFields)
          .SetOptionalItem(EventDelegateFields.edfEventDelegateNumber, vFields)
          .SetOptionalItem(EventDelegateFields.edfCandidateNumber, vFields)
          .SetOptionalItem(EventDelegateFields.edfSequenceNumber, vFields)
        End If
        If (pRSType And EventDelegateRecordSetTypes.edrtAll) = EventDelegateRecordSetTypes.edrtAll Then
          .SetItem(EventDelegateFields.edfPosition, vFields)
          .SetItem(EventDelegateFields.edfOrganisationName, vFields)
          .SetItem(EventDelegateFields.edfAmendedBy, vFields)
          .SetItem(EventDelegateFields.edfAmendedOn, vFields)
          .SetOptionalItem(EventDelegateFields.edfPledgedAmount, vFields)
          .SetOptionalItem(EventDelegateFields.edfDonationTotal, vFields)
          .SetOptionalItem(EventDelegateFields.edfSponsorshipTotal, vFields)
          .SetOptionalItem(EventDelegateFields.edfBookingPaymentAmount, vFields)
          .SetOptionalItem(EventDelegateFields.edfOtherPaymentsTotal, vFields)
        End If
      End With
    End Sub

    Public Sub InitFromBooking(ByVal pEventBooking As EventBooking, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pPosition As String, ByVal pOrganisationName As String, Optional ByVal pEventDelegateNumber As Integer = 0, Optional ByVal pPledgedAmount As Double = 0, Optional ByVal pSequenceNumber As String = "")

      mvClassFields.Item(EventDelegateFields.edfEventNumber).Value = CStr(pEventBooking.EventNumber)
      mvClassFields.Item(EventDelegateFields.edfBookingNumber).Value = CStr(pEventBooking.BookingNumber)
      mvClassFields.Item(EventDelegateFields.edfContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(EventDelegateFields.edfAddressNumber).Value = CStr(pAddressNumber)
      mvClassFields.Item(EventDelegateFields.edfPosition).Value = pPosition
      mvClassFields.Item(EventDelegateFields.edfOrganisationName).Value = pOrganisationName
      mvClassFields.Item(EventDelegateFields.edfAttended).Bool = False
      mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).Value = CStr(pEventDelegateNumber)
      mvClassFields.Item(EventDelegateFields.edfPledgedAmount).Value = CStr(pPledgedAmount)
      If pSequenceNumber.Length > 0 Then mvClassFields.Item(EventDelegateFields.edfSequenceNumber).Value = pSequenceNumber
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventDelegateFields.edfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      If CanDelete() Then
        vWhereFields.Add("event_delegate_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).IntegerValue)
        vUpdateFields.Add("event_delegate_number", CDBField.FieldTypes.cftLong)

        mvEnv.Connection.DeleteRecordsMultiTable("delegate_activities,delegate_links", vWhereFields)
        mvEnv.Connection.UpdateRecords("event_pis", vUpdateFields, vWhereFields, False)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
          mvEnv.Connection.DeleteRecords("delegate_sessions", vWhereFields, False)
        End If
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      End If
    End Sub

    Public Function HasSessionActivities() As Boolean
      HasSessionActivities = mvEnv.Connection.GetCount("session_bookings sb,session_activities sa,contact_categories cc", Nothing, "sb.booking_number = " & BookingNumber & " AND sb.session_number = sa.session_number AND cc.activity = sa.activity AND cc.activity_value = sa.activity_value AND cc.contact_number = " & ContactNumber) > 0
    End Function

    Public Sub DeleteSessionActivities()
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("activity")
      vWhereFields.Add("activity_value")
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT sa.activity, sa.activity_value FROM session_bookings sb,session_activities sa,contact_categories cc WHERE sb.booking_number = " & BookingNumber & " AND sb.session_number = sa.session_number AND cc.activity = sa.activity AND cc.activity_value = sa.activity_value AND cc.contact_number = " & ContactNumber)
      While vRecordSet.Fetch() = True
        vWhereFields(2).Value = vRecordSet.Fields(1).Value
        vWhereFields(3).Value = vRecordSet.Fields(2).Value
        mvEnv.Connection.DeleteRecords("contact_categories", vWhereFields)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Create(ByRef pEventBooking As EventBooking, ByRef pParams As CDBParameters)
      mvClassFields.Item(EventDelegateFields.edfEventNumber).Value = CStr(pEventBooking.EventNumber)
      mvClassFields.Item(EventDelegateFields.edfBookingNumber).Value = CStr(pEventBooking.BookingNumber)
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      If pParams.Exists("ContactNumber") Then mvClassFields.Item(EventDelegateFields.edfContactNumber).Value = CStr(pParams("ContactNumber").IntegerValue)
      If pParams.Exists("AddressNumber") Then mvClassFields.Item(EventDelegateFields.edfAddressNumber).Value = CStr(pParams("AddressNumber").IntegerValue)
      If pParams.Exists("Position") Then mvClassFields.Item(EventDelegateFields.edfPosition).Value = pParams("Position").Value
      If pParams.Exists("OrganisationName") Then mvClassFields.Item(EventDelegateFields.edfOrganisationName).Value = pParams("OrganisationName").Value
      If pParams.Exists("PledgedAmount") Then mvClassFields.Item(EventDelegateFields.edfPledgedAmount).Value = pParams("PledgedAmount").Value
      If pParams.HasValue("Attended") Then mvClassFields.Item(EventDelegateFields.edfAttended).Value = pParams("Attended").Value
      If pParams.Exists("SequenceNumber") Then mvClassFields.Item(EventDelegateFields.edfSequenceNumber).Value = pParams("SequenceNumber").Value
    End Sub

    ''' <summary>Populate Activities and Links collections with data for this Delegate.</summary>
    Public Sub SetDelegateActivitiesAndLinks()
      mvActivities = New CDBCollection()
      mvLinks = New CollectionList(Of DelegateLink)
      If mvExisting Then
        Dim vWhereFields As New CDBFields(New CDBField("event_delegate_number", EventDelegateNumber))
        'Get DelegateActivities
        Dim vDelegateActivity As New DelegateActivity(mvEnv)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vDelegateActivity.GetRecordSetFields(), "delegate_activities da", vWhereFields)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          vDelegateActivity = New DelegateActivity(mvEnv)
          vDelegateActivity.InitFromRecordSet(vRS)
          mvActivities.Add(vDelegateActivity, vDelegateActivity.DelegateActivityNumber.ToString)
        End While
        vRS.CloseRecordSet()
        'Get DelegateLinks
        Dim vDelegateLink As New DelegateLink(mvEnv)
        vSQLStatement = New SQLStatement(mvEnv.Connection, vDelegateLink.GetRecordSetFields(), "delegate_links dl", vWhereFields)
        vRS = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          vDelegateLink = New DelegateLink(mvEnv)
          vDelegateLink.InitFromRecordSet(vRS)
          mvLinks.Add(vDelegateLink.DelegateLinkNumber.ToString, vDelegateLink)
        End While
        vRS.CloseRecordSet()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Public ReadOnly Property Contact() As Contact
      Get
        If mvContact Is Nothing Then
          mvContact = New Contact(mvEnv)
          mvContact.Init(ContactNumber)
        End If
        Contact = mvContact
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(EventDelegateFields.edfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventDelegateFields.edfAddressNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventDelegateFields.edfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventDelegateFields.edfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Attended() As Boolean
      Get
        Attended = mvClassFields.Item(EventDelegateFields.edfAttended).Bool
      End Get
    End Property

    Public ReadOnly Property BookingNumber() As Integer
      Get
        BookingNumber = mvClassFields.Item(EventDelegateFields.edfBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(EventDelegateFields.edfContactNumber).IntegerValue
      End Get
    End Property
    Public Property PledgedAmount() As String
      Get
        PledgedAmount = mvClassFields.Item(EventDelegateFields.edfPledgedAmount).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventDelegateFields.edfPledgedAmount).Value = Value
      End Set
    End Property
    Public ReadOnly Property DonationTotal() As String
      Get
        DonationTotal = mvClassFields.Item(EventDelegateFields.edfDonationTotal).Value
      End Get
    End Property
    Public ReadOnly Property SponsorshipTotal() As String
      Get
        SponsorshipTotal = mvClassFields.Item(EventDelegateFields.edfSponsorshipTotal).Value
      End Get
    End Property
    Public ReadOnly Property BookingPaymentAmount() As String
      Get
        BookingPaymentAmount = mvClassFields.Item(EventDelegateFields.edfBookingPaymentAmount).Value
      End Get
    End Property
    Public ReadOnly Property OtherPaymentsTotal() As String
      Get
        OtherPaymentsTotal = mvClassFields.Item(EventDelegateFields.edfOtherPaymentsTotal).Value
      End Get
    End Property
    Public Property CandidateNumber() As Integer
      Get
        CandidateNumber = mvClassFields.Item(EventDelegateFields.edfCandidateNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventDelegateFields.edfCandidateNumber).IntegerValue = Value
      End Set
    End Property

    Public Property Activities() As CDBCollection
      Get
        If mvActivities Is Nothing Then mvActivities = New CDBCollection
        Activities = mvActivities
      End Get
      Set(ByVal Value As CDBCollection)
        If mvActivities Is Nothing Then mvActivities = New CDBCollection
        mvActivities = value
      End Set
    End Property
  
    Public ReadOnly Property EventDelegateNumber() As Integer
      Get
        EventDelegateNumber = mvClassFields.Item(EventDelegateFields.edfEventDelegateNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvClassFields.Item(EventDelegateFields.edfSequenceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventDelegateFields.edfEventNumber).IntegerValue
      End Get
    End Property

    Public Property OrganisationName() As String
      Get
        OrganisationName = mvClassFields.Item(EventDelegateFields.edfOrganisationName).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventDelegateFields.edfOrganisationName).Value = Value
      End Set
    End Property

    Public Property Position() As String
      Get
        Position = mvClassFields.Item(EventDelegateFields.edfPosition).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventDelegateFields.edfPosition).Value = Value
      End Set
    End Property

    ''' <summary>Gets a <see cref="CollectionList(Of DelegateLink)">CollectionList</see>  of Delegate Links for this Delegate.</summary>
    ''' <returns><see cref="CollectionList(Of DelegateLink)">CollectionList</see>  of Delegate Links</returns>
    Public ReadOnly Property Links() As CollectionList(Of DelegateLink)
      Get
        If mvLinks Is Nothing Then mvLinks = New CollectionList(Of DelegateLink)
        Return mvLinks
      End Get
    End Property

    Private Function CanDelete() As Boolean
      Dim vWhereFields As New CDBFields
      Dim vCanDelete As Boolean

      vCanDelete = True
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPIS) Then
        vWhereFields.Add("event_delegate_number", CDBField.FieldTypes.cftLong, EventDelegateNumber)
        vWhereFields.Add("reconciled_on", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("amount", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        If mvEnv.Connection.GetCount("event_pis", vWhereFields) > 0 Then
          RaiseError(DataAccessErrors.daeCannotDeleteDelegateAsPaidPIS)
        End If
      End If
      CanDelete = vCanDelete
    End Function
  End Class
End Namespace
