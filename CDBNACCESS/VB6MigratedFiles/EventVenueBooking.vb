

Namespace Access
  Public Class EventVenueBooking

    Public Enum EventVenueBookingRecordSetTypes 'These are bit values
      evbrtAll = &HFFS
      'ADD additional recordset types here
      evbrtVenueInfo = &H100S
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventVenueBookingFields
      evbfAll = 0
      evbfVenueBookingNumber
      evbfEventNumber
      evbfVenue
      evbfVenueReference
      evbfAmount
      evbfPaymentDate
      evbfDepositAmount
      evbfDepositDate
      evbfFullAmount
      evbfFullPaymentDate
      evbfConfirmedBy
      evbfConfirmedOn
      evbfAmendedBy
      evbfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvVenueDesc As String
    Private mvLocation As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "event_venue_bookings"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("venue_booking_number", CDBField.FieldTypes.cftLong)
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("venue")
          .Add("venue_reference")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("payment_date", CDBField.FieldTypes.cftDate)
          .Add("deposit_amount", CDBField.FieldTypes.cftNumeric)
          .Add("deposit_date", CDBField.FieldTypes.cftDate)
          .Add("full_amount", CDBField.FieldTypes.cftNumeric)
          .Add("full_payment_date", CDBField.FieldTypes.cftDate)
          .Add("confirmed_by")
          .Add("confirmed_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(EventVenueBookingFields.evbfVenueBookingNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventVenueBookingFields.evbfVenueBookingNumber).PrefixRequired = True
        mvClassFields.Item(EventVenueBookingFields.evbfEventNumber).PrefixRequired = True
        mvClassFields.Item(EventVenueBookingFields.evbfVenue).PrefixRequired = True
        mvClassFields.Item(EventVenueBookingFields.evbfAmendedBy).PrefixRequired = True
        mvClassFields.Item(EventVenueBookingFields.evbfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventVenueBookingFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventVenueBookingFields.evbfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventVenueBookingFields.evbfAmendedBy).Value = mvEnv.User.Logname
      If mvClassFields.Item(EventVenueBookingFields.evbfVenueBookingNumber).IntegerValue = 0 Then
        mvClassFields.Item(EventVenueBookingFields.evbfVenueBookingNumber).Value = CStr(mvEnv.GetControlNumber("VB"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventVenueBookingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And EventVenueBookingRecordSetTypes.evbrtAll) = EventVenueBookingRecordSetTypes.evbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "evb")
      End If
      If (pRSType And EventVenueBookingRecordSetTypes.evbrtVenueInfo) = EventVenueBookingRecordSetTypes.evbrtVenueInfo Then
        vFields = vFields & ",venue_desc,location"
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pVenueBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pVenueBookingNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventVenueBookingRecordSetTypes.evbrtAll) & " FROM event_venue_bookings evb WHERE venue_booking_number = " & pVenueBookingNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventVenueBookingRecordSetTypes.evbrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventVenueBookingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventVenueBookingFields.evbfVenueBookingNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventVenueBookingRecordSetTypes.evbrtAll) = EventVenueBookingRecordSetTypes.evbrtAll Then
          .SetItem(EventVenueBookingFields.evbfEventNumber, vFields)
          .SetItem(EventVenueBookingFields.evbfVenue, vFields)
          .SetItem(EventVenueBookingFields.evbfVenueReference, vFields)
          .SetItem(EventVenueBookingFields.evbfAmount, vFields)
          .SetItem(EventVenueBookingFields.evbfPaymentDate, vFields)
          .SetItem(EventVenueBookingFields.evbfDepositAmount, vFields)
          .SetItem(EventVenueBookingFields.evbfDepositDate, vFields)
          .SetItem(EventVenueBookingFields.evbfFullAmount, vFields)
          .SetItem(EventVenueBookingFields.evbfFullPaymentDate, vFields)
          .SetItem(EventVenueBookingFields.evbfConfirmedBy, vFields)
          .SetItem(EventVenueBookingFields.evbfConfirmedOn, vFields)
          .SetItem(EventVenueBookingFields.evbfAmendedBy, vFields)
          .SetItem(EventVenueBookingFields.evbfAmendedOn, vFields)
        End If
        If (pRSType And EventVenueBookingRecordSetTypes.evbrtVenueInfo) = EventVenueBookingRecordSetTypes.evbrtVenueInfo Then
          mvVenueDesc = vFields("venue_desc").Value
          mvLocation = vFields("location").Value
        End If
      End With
    End Sub

    Public Sub InitFromEvent(ByVal pEnv As CDBEnvironment, ByVal pEvent As CDBEvent)
      mvEnv = pEnv
      InitClassFields()
      With pEvent
        mvClassFields.Item(EventVenueBookingFields.evbfEventNumber).Value = CStr(.EventNumber)
        mvClassFields.Item(EventVenueBookingFields.evbfVenue).Value = .Venue
        mvClassFields.Item(EventVenueBookingFields.evbfVenueReference).Value = .VenueReference
        mvClassFields.Item(EventVenueBookingFields.evbfConfirmedBy).Value = .VenueConfirmedBy
        mvClassFields.Item(EventVenueBookingFields.evbfConfirmedOn).Value = .VenueConfirmed
      End With
    End Sub

    Public Sub InitFromVenueBooking(ByRef pVenueBooking As EventVenueBooking, ByVal pEvent As CDBEvent)
      mvClassFields.Item(EventVenueBookingFields.evbfEventNumber).Value = CStr(pEvent.EventNumber)
      mvClassFields.Item(EventVenueBookingFields.evbfVenue).Value = pVenueBooking.Venue
      mvClassFields.Item(EventVenueBookingFields.evbfVenueReference).Value = pVenueBooking.VenueReference
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventVenueBookingFields.evbfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(EventVenueBookingFields.evbfEventNumber).Value = pParams("EventNumber").Value
        .Item(EventVenueBookingFields.evbfVenue).Value = pParams("Venue").Value
      End With
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Dim vEvent As New CDBEvent(mvEnv)

      With mvClassFields
        If pParams.Exists("Venue") Then .Item(EventVenueBookingFields.evbfVenue).Value = pParams("Venue").Value
        If pParams.Exists("VenueReference") Then .Item(EventVenueBookingFields.evbfVenueReference).Value = pParams("VenueReference").Value
        If pParams.Exists("Balance") Then .Item(EventVenueBookingFields.evbfAmount).Value = pParams("Balance").Value
        If pParams.Exists("BalancePaidDate") Then .Item(EventVenueBookingFields.evbfPaymentDate).Value = pParams("BalancePaidDate").Value
        If pParams.Exists("Deposit") Then .Item(EventVenueBookingFields.evbfDepositAmount).Value = pParams("Deposit").Value
        If pParams.Exists("DepositPaidDate") Then .Item(EventVenueBookingFields.evbfDepositDate).Value = pParams("DepositPaidDate").Value
        If pParams.Exists("TotalAmount") Then .Item(EventVenueBookingFields.evbfFullAmount).Value = pParams("TotalAmount").Value
        If pParams.Exists("DueDate") Then .Item(EventVenueBookingFields.evbfFullPaymentDate).Value = pParams("DueDate").Value
        If pParams.Exists("ConfirmedBy") Then .Item(EventVenueBookingFields.evbfConfirmedBy).Value = pParams("ConfirmedBy").Value
        If pParams.Exists("ConfirmedOn") Then .Item(EventVenueBookingFields.evbfConfirmedOn).Value = pParams("ConfirmedOn").Value
        If Len(ConfirmedOn) > 0 And Len(ConfirmedBy) = 0 Then .Item(EventVenueBookingFields.evbfConfirmedBy).Value = mvEnv.User.UserID
      End With
      If mvExisting Then
        vEvent.Init(EventNumber)
        If vEvent.Existing Then
          If VenueBookingNumber = CDbl(vEvent.BaseSession.VenueBookingNumber) Then 'Default venue booking
            If Venue <> vEvent.Venue Then 'Has changed
              vEvent.BaseSession.SetLocationFromVenue(Venue)
              vEvent.BaseSession.Save()
            End If
            vEvent.SetVenueFromBooking(Me)
            vEvent.Save()
          End If
        End If
      End If
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventVenueBookingFields.evbfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventVenueBookingFields.evbfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(EventVenueBookingFields.evbfAmount).Value
      End Get
    End Property

    Public ReadOnly Property ConfirmedBy() As String
      Get
        ConfirmedBy = mvClassFields.Item(EventVenueBookingFields.evbfConfirmedBy).Value
      End Get
    End Property

    Public ReadOnly Property ConfirmedOn() As String
      Get
        ConfirmedOn = mvClassFields.Item(EventVenueBookingFields.evbfConfirmedOn).Value
      End Get
    End Property

    Public ReadOnly Property DepositAmount() As String
      Get
        DepositAmount = mvClassFields.Item(EventVenueBookingFields.evbfDepositAmount).Value
      End Get
    End Property

    Public ReadOnly Property DepositDate() As String
      Get
        DepositDate = mvClassFields.Item(EventVenueBookingFields.evbfDepositDate).Value
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventVenueBookingFields.evbfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property FullAmount() As String
      Get
        FullAmount = mvClassFields.Item(EventVenueBookingFields.evbfFullAmount).Value
      End Get
    End Property

    Public ReadOnly Property FullPaymentDate() As String
      Get
        FullPaymentDate = mvClassFields.Item(EventVenueBookingFields.evbfFullPaymentDate).Value
      End Get
    End Property

    Public ReadOnly Property PaymentDate() As String
      Get
        PaymentDate = mvClassFields.Item(EventVenueBookingFields.evbfPaymentDate).Value
      End Get
    End Property

    Public ReadOnly Property Venue() As String
      Get
        Venue = mvClassFields.Item(EventVenueBookingFields.evbfVenue).Value
      End Get
    End Property

    Public ReadOnly Property VenueBookingNumber() As Integer
      Get
        VenueBookingNumber = mvClassFields.Item(EventVenueBookingFields.evbfVenueBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property VenueReference() As String
      Get
        VenueReference = mvClassFields.Item(EventVenueBookingFields.evbfVenueReference).Value
      End Get
    End Property

    Public ReadOnly Property VenueDescription() As String
      Get
        VenueDescription = mvVenueDesc
      End Get
    End Property

    Public ReadOnly Property VenueLocation() As String
      Get
        VenueLocation = mvLocation
      End Get
    End Property
  End Class
End Namespace
