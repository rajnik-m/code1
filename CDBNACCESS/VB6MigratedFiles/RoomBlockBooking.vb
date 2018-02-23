

Namespace Access
  Public Class RoomBlockBooking

    Public Enum RoomBlockBookingRecordSetTypes 'These are bit values
      rbbrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum RoomBlockBookingFields
      rbbfAll = 0
      rbbfBlockBookingNumber
      rbbfOrganisationNumber
      rbbfAddressNumber
      rbbfContactNumber
      rbbfBookedDate
      rbbfFromDate
      rbbfToDate
      rbbfRoomType
      rbbfNumberOfRooms
      rbbfNightsAvailable
      rbbfRackRate
      rbbfAgreedRate
      rbbfReleaseDate
      rbbfConfirmedDate
      rbbfProduct
      rbbfRate
      rbbfNotes
      rbbfAmendedBy
      rbbfAmendedOn
    End Enum

    Public Enum RoomBookingDeleteAllowedStatuses
      rbdasRoomsBooked
      rbdasOtherEvent
      rbdasOK
    End Enum

    Private mvEventRoomLink As EventRoomLink
    Private mvOrganisation As Organisation
    Private mvEnforceAllocation As String = ""

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "room_block_bookings"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("block_booking_number", CDBField.FieldTypes.cftLong)
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("booked_date", CDBField.FieldTypes.cftDate)
          .Add("from_date", CDBField.FieldTypes.cftDate)
          .Add("to_date", CDBField.FieldTypes.cftDate)
          .Add("room_type")
          .Add("number_of_rooms", CDBField.FieldTypes.cftInteger)
          .Add("nights_available", CDBField.FieldTypes.cftInteger)
          .Add("rack_rate", CDBField.FieldTypes.cftNumeric)
          .Add("agreed_rate", CDBField.FieldTypes.cftNumeric)
          .Add("release_date", CDBField.FieldTypes.cftDate)
          .Add("confirmed_date", CDBField.FieldTypes.cftDate)
          .Add("product")
          .Add("rate")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(RoomBlockBookingFields.rbbfBlockBookingNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(RoomBlockBookingFields.rbbfBookedDate).Value = TodaysDate()
      mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).Value = TodaysDate()
      mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).Value = TodaysDate()
    End Sub

    Private Sub SetValid(ByVal pField As RoomBlockBookingFields)
      'Add code here to ensure all values are valid before saving
      If BlockBookingNumber = 0 Then mvClassFields.Item(RoomBlockBookingFields.rbbfBlockBookingNumber).Value = CStr(mvEnv.GetControlNumber("BB"))
      'vNights = DateDiff("d", mvClassFields.Item(rbbfFromDate).Value, mvClassFields.Item(rbbfToDate).Value)
      'mvClassFields.Item(rbbfNightsAvailable).Value = vNights
      mvClassFields.Item(RoomBlockBookingFields.rbbfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(RoomBlockBookingFields.rbbfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As RoomBlockBookingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = RoomBlockBookingRecordSetTypes.rbbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "rbb")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBlockBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBlockBookingNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(RoomBlockBookingRecordSetTypes.rbbrtAll) & " FROM room_block_bookings rbb WHERE block_booking_number = " & pBlockBookingNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, RoomBlockBookingRecordSetTypes.rbbrtAll)
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
    Public Function DeleteAllowedStatus() As RoomBookingDeleteAllowedStatuses
      Dim vWhereFields As New CDBFields

      With mvEnv.Connection
        vWhereFields = New CDBFields
        vWhereFields.Add("block_booking_number", CDBField.FieldTypes.cftLong, BlockBookingNumber)
        If .GetCount("contact_room_bookings", vWhereFields) > 0 Then
          DeleteAllowedStatus = RoomBookingDeleteAllowedStatuses.rbdasRoomsBooked 'Rooms from this booking have been taken\r\n\r\nRoom booking cannot be deleted
        Else
          vWhereFields.Add("event_number", EventRoomLink.EventNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          If .GetCount("event_room_links", vWhereFields) > 0 Then
            DeleteAllowedStatus = RoomBookingDeleteAllowedStatuses.rbdasOtherEvent 'This booking is used by another event\r\n\r\nRoom booking cannot be deleted
          Else
            DeleteAllowedStatus = RoomBookingDeleteAllowedStatuses.rbdasOK
          End If
        End If
      End With
    End Function

    Public Function ChangeAllowed() As Boolean
      Dim vWhereFields As New CDBFields

      With mvEnv.Connection
        vWhereFields = New CDBFields
        vWhereFields.Add("block_booking_number", CDBField.FieldTypes.cftLong, BlockBookingNumber)
        ChangeAllowed = .GetCount("contact_room_bookings", vWhereFields) = 0
      End With
    End Function

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As RoomBlockBookingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(RoomBlockBookingFields.rbbfBlockBookingNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And RoomBlockBookingRecordSetTypes.rbbrtAll) = RoomBlockBookingRecordSetTypes.rbbrtAll Then
          .SetItem(RoomBlockBookingFields.rbbfOrganisationNumber, vFields)
          .SetItem(RoomBlockBookingFields.rbbfAddressNumber, vFields)
          .SetItem(RoomBlockBookingFields.rbbfContactNumber, vFields)
          .SetItem(RoomBlockBookingFields.rbbfBookedDate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfFromDate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfToDate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfRoomType, vFields)
          .SetItem(RoomBlockBookingFields.rbbfNumberOfRooms, vFields)
          .SetItem(RoomBlockBookingFields.rbbfNightsAvailable, vFields)
          .SetItem(RoomBlockBookingFields.rbbfRackRate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfAgreedRate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfReleaseDate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfConfirmedDate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfProduct, vFields)
          .SetItem(RoomBlockBookingFields.rbbfRate, vFields)
          .SetItem(RoomBlockBookingFields.rbbfNotes, vFields)
          .SetItem(RoomBlockBookingFields.rbbfAmendedBy, vFields)
          .SetItem(RoomBlockBookingFields.rbbfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      EventRoomLink.Delete()
    End Sub

    Public ReadOnly Property EventRoomLink() As EventRoomLink
      Get
        If mvEventRoomLink Is Nothing Then
          If BlockBookingNumber > 0 Then
            mvEventRoomLink = New EventRoomLink
            mvEventRoomLink.Init(mvEnv, BlockBookingNumber)
          End If
        End If
        EventRoomLink = mvEventRoomLink
      End Get
    End Property
    Public ReadOnly Property Organisation() As Organisation
      Get
        If mvOrganisation Is Nothing Then
          If OrganisationNumber > 0 Then
            mvOrganisation = New Organisation(mvEnv)
            mvOrganisation.Init(OrganisationNumber)
          End If
        End If
        Organisation = mvOrganisation
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(RoomBlockBookingFields.rbbfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AgreedRate() As Double
      Get
        AgreedRate = mvClassFields.Item(RoomBlockBookingFields.rbbfAgreedRate).DoubleValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(RoomBlockBookingFields.rbbfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(RoomBlockBookingFields.rbbfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BlockBookingNumber() As Integer
      Get
        BlockBookingNumber = mvClassFields.Item(RoomBlockBookingFields.rbbfBlockBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BookedDate() As String
      Get
        BookedDate = mvClassFields.Item(RoomBlockBookingFields.rbbfBookedDate).Value
      End Get
    End Property

    Public ReadOnly Property ConfirmedDate() As String
      Get
        ConfirmedDate = mvClassFields.Item(RoomBlockBookingFields.rbbfConfirmedDate).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(RoomBlockBookingFields.rbbfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property FromDate() As String
      Get
        FromDate = mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).Value
      End Get
    End Property

    Public ReadOnly Property NightsAvailable() As Integer
      Get
        NightsAvailable = mvClassFields.Item(RoomBlockBookingFields.rbbfNightsAvailable).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(RoomBlockBookingFields.rbbfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property NumberOfRooms() As Integer
      Get
        NumberOfRooms = mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(RoomBlockBookingFields.rbbfOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(RoomBlockBookingFields.rbbfProduct).Value
      End Get
    End Property

    Public ReadOnly Property RackRate() As Double
      Get
        RackRate = mvClassFields.Item(RoomBlockBookingFields.rbbfRackRate).DoubleValue
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(RoomBlockBookingFields.rbbfRate).Value
      End Get
    End Property

    Public ReadOnly Property ReleaseDate() As String
      Get
        ReleaseDate = mvClassFields.Item(RoomBlockBookingFields.rbbfReleaseDate).Value
      End Get
    End Property

    Public ReadOnly Property RoomType() As String
      Get
        RoomType = mvClassFields.Item(RoomBlockBookingFields.rbbfRoomType).Value
      End Get
    End Property
    Public ReadOnly Property EnforceAllocation() As Boolean
      Get
        If mvEnforceAllocation.Length = 0 Then
          mvEnforceAllocation = mvEnv.Connection.GetValue("SELECT enforce_allocation FROM room_types WHERE room_type = '" & mvClassFields.Item(RoomBlockBookingFields.rbbfRoomType).Value & "'")
        End If
        EnforceAllocation = mvEnforceAllocation = "Y"
      End Get
    End Property
    Public ReadOnly Property ToDate() As String
      Get
        ToDate = mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).Value
      End Get
    End Property
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vNights As Integer
      Dim vNightsChange As Integer

      SetValid(RoomBlockBookingFields.rbbfAll)

      vNights = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).Value), CDate(mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).Value)))

      If mvExisting = False Then
        mvClassFields.Item(RoomBlockBookingFields.rbbfNightsAvailable).Value = CStr(vNights * CDbl(mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).Value))
      Else
        If mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).Value <> mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).SetValue Or mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).Value <> mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).SetValue Or mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).Value <> mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).SetValue Then
          vNightsChange = CInt((CDbl(mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).Value) - CDbl(mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).SetValue)) * vNights)
          vNightsChange = CInt(vNightsChange + ((vNights - DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(mvClassFields.Item(RoomBlockBookingFields.rbbfFromDate).SetValue), CDate(mvClassFields.Item(RoomBlockBookingFields.rbbfToDate).SetValue))) * CDbl(mvClassFields.Item(RoomBlockBookingFields.rbbfNumberOfRooms).SetValue)))
          mvClassFields.Item(RoomBlockBookingFields.rbbfNightsAvailable).Value = CStr(CDbl(mvClassFields.Item(RoomBlockBookingFields.rbbfNightsAvailable).Value) + vNightsChange)
        End If
      End If

      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      Dim vMsg As String = ""
      Dim vNewNumberOfRooms As Integer

      With mvClassFields
        If pParams.Exists("OrganisationNumber") Then .Item(RoomBlockBookingFields.rbbfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(RoomBlockBookingFields.rbbfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("ContactNumber") Then .Item(RoomBlockBookingFields.rbbfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("BookedOn") Then .Item(RoomBlockBookingFields.rbbfBookedDate).Value = pParams("BookedOn").Value
        If pParams.Exists("FromDate") Then .Item(RoomBlockBookingFields.rbbfFromDate).Value = pParams("FromDate").Value
        If pParams.Exists("ToDate") Then .Item(RoomBlockBookingFields.rbbfToDate).Value = pParams("ToDate").Value
        If pParams.Exists("RoomType") Then .Item(RoomBlockBookingFields.rbbfRoomType).Value = pParams("RoomType").Value
        If pParams.Exists("RackRate") Then .Item(RoomBlockBookingFields.rbbfRackRate).Value = pParams("RackRate").Value
        If pParams.Exists("AgreedRate") Then .Item(RoomBlockBookingFields.rbbfAgreedRate).Value = pParams("AgreedRate").Value
        If pParams.Exists("ReleaseDate") Then .Item(RoomBlockBookingFields.rbbfReleaseDate).Value = pParams("ReleaseDate").Value
        If pParams.Exists("ConfirmedOn") Then .Item(RoomBlockBookingFields.rbbfConfirmedDate).Value = pParams("ConfirmedOn").Value
        If pParams.Exists("Product") Then .Item(RoomBlockBookingFields.rbbfProduct).Value = pParams("Product").Value
        If pParams.Exists("Rate") Then .Item(RoomBlockBookingFields.rbbfRate).Value = pParams("Rate").Value
        If pParams.Exists("Notes") Then .Item(RoomBlockBookingFields.rbbfNotes).Value = pParams("Notes").Value

        If mvExisting Then
          If (.Item(RoomBlockBookingFields.rbbfOrganisationNumber).ValueChanged Or .Item(RoomBlockBookingFields.rbbfAddressNumber).ValueChanged Or .Item(RoomBlockBookingFields.rbbfContactNumber).ValueChanged Or .Item(RoomBlockBookingFields.rbbfFromDate).ValueChanged Or .Item(RoomBlockBookingFields.rbbfToDate).ValueChanged Or .Item(RoomBlockBookingFields.rbbfRoomType).ValueChanged Or .Item(RoomBlockBookingFields.rbbfProduct).ValueChanged Or .Item(RoomBlockBookingFields.rbbfRate).ValueChanged) Then
            If Not ChangeAllowed() Then RaiseError(DataAccessErrors.daeBookingsMadeCannotChange)
          End If
          If pParams.Exists("NumberOfRooms") Then
            vNewNumberOfRooms = pParams("NumberOfRooms").IntegerValue
            If vNewNumberOfRooms < NumberOfRooms Then
              'Number of rooms will be reduced check that it is not less than number already sold
              If CheckRoomBooking(FromDate, ToDate, NumberOfRooms - vNewNumberOfRooms, 0, vMsg) = False Then RaiseError(DataAccessErrors.daeRoomsBookedExceedsQuantity)
            End If
            .Item(RoomBlockBookingFields.rbbfNumberOfRooms).Value = pParams("NumberOfRooms").Value
          End If
        Else
          If pParams.Exists("NumberOfRooms") Then .Item(RoomBlockBookingFields.rbbfNumberOfRooms).Value = pParams("NumberOfRooms").Value
        End If
      End With
    End Sub

    Function CheckRoomBooking(ByVal pFromDate As String, ByVal pToDate As String, ByVal pQty As Integer, ByVal pRoomBookingNumber As Integer, ByRef pErrorMsg As String) As Boolean
      Dim vCanBook As Boolean
      Dim vNumberAvailable As Integer
      Dim vRestrict As String = ""
      Dim vRecordSet As CDBRecordSet
      Dim vRoomDate As String = ""
      Dim vNoBookings As Boolean

      'Find out if we can take this room booking
      vCanBook = True
      If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pFromDate), CDate(pToDate)) <= 0 Then
        pErrorMsg = (ProjectText.String18780) 'At least 1 night must be selected for the Booking
        vCanBook = False
      ElseIf DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pFromDate), CDate(FromDate)) > 0 Then
        pErrorMsg = String.Format(ProjectText.String25624, FromDate) 'Selected rooms not available before %s
        vCanBook = False
      ElseIf DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(ToDate), CDate(pToDate)) > 0 Then
        pErrorMsg = String.Format(ProjectText.String25625, ToDate) 'Selected rooms not available after %s
        vCanBook = False
      End If
      If vCanBook Then
        If EnforceAllocation Then
          If pRoomBookingNumber > 0 Then vRestrict = " AND crb.room_booking_number <> " & pRoomBookingNumber
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM (number_of_rooms)  AS  rooms_sum FROM contact_room_bookings crb WHERE crb.block_booking_number = " & BlockBookingNumber & " AND crb.cancellation_reason IS NULL AND from_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pFromDate) & " AND to_date" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pToDate) & vRestrict)
          If vRecordSet.Fetch() = True Then
            vNumberAvailable = NumberOfRooms - IntegerValue(vRecordSet.Fields("rooms_sum").Value)
            If vNumberAvailable < pQty Then
              vRoomDate = pFromDate & " to " & pToDate
              vCanBook = False
            End If
          Else
            vNoBookings = True
          End If
          vRecordSet.CloseRecordSet()
        Else
          If pRoomBookingNumber > 0 Then vRestrict = " AND rbl.room_booking_number <> " & pRoomBookingNumber
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT count(*)  AS record_count, room_date FROM contact_room_bookings crb, room_booking_links rbl WHERE crb.block_booking_number = " & BlockBookingNumber & " AND crb.room_booking_number = rbl.room_booking_number AND room_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pFromDate) & " AND room_date" & mvEnv.Connection.SQLLiteral("<", CDBField.FieldTypes.cftDate, pToDate) & vRestrict & " GROUP by room_date")
          vNoBookings = True
          While vRecordSet.Fetch() = True And vCanBook
            vNoBookings = False
            vNumberAvailable = NumberOfRooms - (vRecordSet.Fields(1).IntegerValue)
            If vNumberAvailable < pQty Then
              vRoomDate = vRecordSet.Fields(2).Value
              vCanBook = False
            End If
          End While
          vRecordSet.CloseRecordSet()
        End If
        If vNoBookings Then
          If NumberOfRooms < pQty Then
            vRoomDate = FromDate
            vCanBook = False
          End If
        End If
        If vCanBook = False Then
          If vNumberAvailable = 0 Then
            pErrorMsg = String.Format(ProjectText.String25626, vRoomDate) 'No rooms available on %s
          Else
            pErrorMsg = String.Format(ProjectText.String25627, CStr(vNumberAvailable), vRoomDate) 'Only %s rooms available on %s
          End If
        End If
      End If
      CheckRoomBooking = vCanBook
    End Function
  End Class
End Namespace
