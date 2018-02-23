

Namespace Access
  Public Class EventAccommodationBooking

    Public Enum EventAccommodationBookingRecordSetTypes 'These are bit values
      eabrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventAccommodationBookingFields
      eabAll = 0
      eabEventNumber
      eabRoomBookingNumber
      eabBlockBookingNumber
      eabContactNumber
      eabAddressNumber
      eabBookedDate
      eabFromDate
      eabToDate
      eabNumberOfRooms
      eabConfirmedDate
      eabBookingStatus
      eabNotes
      eabRate
      eabAmount
      eabBatchNumber
      eabTransactionNumber
      eabLineNumber
      eabCancellationReason
      eabCancelledBy
      eabCancelledOn
      eabAmendedBy
      eabAmendedOn
      eabSalesContactNumber
      eabCancellationSource
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvBatch As Batch
    Private mvBT As BatchTransaction
    Private mvBTA As BatchTransactionAnalysis

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "contact_room_bookings"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftInteger)
          .Add("room_booking_number", CDBField.FieldTypes.cftLong)
          .Add("block_booking_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("booked_date", CDBField.FieldTypes.cftDate)
          .Add("from_date", CDBField.FieldTypes.cftDate)
          .Add("to_date", CDBField.FieldTypes.cftDate)
          .Add("number_of_rooms", CDBField.FieldTypes.cftInteger)
          .Add("confirmed_date", CDBField.FieldTypes.cftDate)
          .Add("booking_status")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("rate")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("sales_contact_number", CDBField.FieldTypes.cftLong)
          .Add("cancellation_source")

          .Item(EventAccommodationBookingFields.eabRoomBookingNumber).SetPrimaryKeyOnly()

          .Item(EventAccommodationBookingFields.eabContactNumber).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabAddressNumber).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabBookedDate).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabFromDate).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabToDate).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabNumberOfRooms).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabConfirmedDate).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabNotes).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabRate).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabBlockBookingNumber).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabAmendedOn).PrefixRequired = True
          .Item(EventAccommodationBookingFields.eabAmendedBy).PrefixRequired = True
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvBatch = Nothing
      mvBT = Nothing
      mvBTA = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventAccommodationBookingFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(EventAccommodationBookingFields.eabRoomBookingNumber).IntegerValue = 0 Then mvClassFields.Item(EventAccommodationBookingFields.eabRoomBookingNumber).IntegerValue = mvEnv.GetControlNumber("RB")
      mvClassFields.Item(EventAccommodationBookingFields.eabAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventAccommodationBookingFields.eabAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventAccommodationBookingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventAccommodationBookingRecordSetTypes.eabrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "crb")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pRoomBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pRoomBookingNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventAccommodationBookingRecordSetTypes.eabrtAll) & " FROM contact_room_bookings crb WHERE room_booking_number = " & pRoomBookingNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventAccommodationBookingRecordSetTypes.eabrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventAccommodationBookingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventAccommodationBookingFields.eabRoomBookingNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventAccommodationBookingRecordSetTypes.eabrtAll) = EventAccommodationBookingRecordSetTypes.eabrtAll Then
          .SetItem(EventAccommodationBookingFields.eabEventNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabBlockBookingNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabContactNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabAddressNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabBookedDate, vFields)
          .SetItem(EventAccommodationBookingFields.eabFromDate, vFields)
          .SetItem(EventAccommodationBookingFields.eabToDate, vFields)
          .SetItem(EventAccommodationBookingFields.eabNumberOfRooms, vFields)
          .SetItem(EventAccommodationBookingFields.eabConfirmedDate, vFields)
          .SetItem(EventAccommodationBookingFields.eabBookingStatus, vFields)
          .SetItem(EventAccommodationBookingFields.eabNotes, vFields)
          .SetItem(EventAccommodationBookingFields.eabRate, vFields)
          .SetItem(EventAccommodationBookingFields.eabAmount, vFields)
          .SetItem(EventAccommodationBookingFields.eabBatchNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabTransactionNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabLineNumber, vFields)
          .SetItem(EventAccommodationBookingFields.eabCancellationReason, vFields)
          .SetItem(EventAccommodationBookingFields.eabCancelledBy, vFields)
          .SetItem(EventAccommodationBookingFields.eabCancelledOn, vFields)
          .SetItem(EventAccommodationBookingFields.eabAmendedBy, vFields)
          .SetItem(EventAccommodationBookingFields.eabAmendedOn, vFields)
          .SetItem(EventAccommodationBookingFields.eabSalesContactNumber, vFields)
          .SetOptionalItem(EventAccommodationBookingFields.eabCancellationSource, vFields)
        End If
      End With
    End Sub

    Public Sub SetTransactionInfo(ByVal pEnv As CDBEnvironment, ByVal pBookingNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pSalesContactNumber As Integer, Optional ByVal pResetStatusCS As Boolean = False)
      'This routine initialises the object and sets the primary key values
      'To act as though it has just read the data
      'This is used so we can update it without reading it first
      Init(pEnv)
      If pResetStatusCS = True Then BookingStatus = EventBooking.EventBookingStatuses.ebsBookedCreditSale
      mvClassFields(EventAccommodationBookingFields.eabRoomBookingNumber).SetValue = CStr(pBookingNumber)
      mvExisting = True
      mvClassFields(EventAccommodationBookingFields.eabBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(EventAccommodationBookingFields.eabTransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields(EventAccommodationBookingFields.eabLineNumber).IntegerValue = pLineNumber
      If pSalesContactNumber > 0 Then mvClassFields(EventAccommodationBookingFields.eabSalesContactNumber).IntegerValue = pSalesContactNumber
    End Sub

    Public Sub Save(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      SetValid(EventAccommodationBookingFields.eabAll)
      If Not Existing Then AddBookingLinks()
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub
    Public Sub ModifyBooking(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pBlockBookingNo As Integer, ByRef pNumberOfRooms As Integer, ByRef pFrom As String, ByRef pTo As String, ByRef pStatus As EventBooking.EventBookingStatuses, ByRef pRateCode As String)
      Dim vChangedBooking As Boolean

      If (pNumberOfRooms <> NumberOfRooms) Or pFrom <> FromDate Or pTo <> ToDate Or pBlockBookingNo <> BlockBookingNumber Then
        vChangedBooking = True
        RemoveBookingLinks()
      End If
      mvClassFields.Item(EventAccommodationBookingFields.eabContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(EventAccommodationBookingFields.eabAddressNumber).Value = CStr(pAddressNumber)
      mvClassFields.Item(EventAccommodationBookingFields.eabBlockBookingNumber).Value = CStr(pBlockBookingNo)
      mvClassFields.Item(EventAccommodationBookingFields.eabNumberOfRooms).Value = CStr(pNumberOfRooms)
      mvClassFields.Item(EventAccommodationBookingFields.eabFromDate).Value = pFrom
      mvClassFields.Item(EventAccommodationBookingFields.eabToDate).Value = pTo
      BookingStatus = pStatus
      mvClassFields.Item(EventAccommodationBookingFields.eabRate).Value = pRateCode
      If vChangedBooking Then AddBookingLinks()
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      CancelOrDelete(True)
    End Sub

    Public Sub Cancel(ByRef pCancellationReason As String, ByRef pCancellationSource As String, ByVal pCancellationAmount As Double)
      CancelOrDelete(False, pCancellationReason, pCancellationSource, pCancellationAmount)
    End Sub

    Private Sub AddBookingLinks()
      Dim vStartDate As Date
      Dim vNights As Integer
      Dim vQuantity As Integer
      Dim vCount As Integer
      Dim vInsertFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vEnforceAllocation As Boolean
      Dim vIndex As Integer
      Dim vRoomID As Integer

      'BC 5163: Ref DUK Room Share mod
      'If Enforce Allocation flag is set on the Room Type, create one room_block_booking per room per capacity of Room Type.
      'All records will have date set to start date & display of these records will show the date range of the Booking.
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT enforce_allocation,capacity FROM room_block_bookings rbb,room_types rt WHERE rbb.block_booking_number = " & BlockBookingNumber & " AND rt.room_type = rbb.room_type")
      If vRecordSet.Fetch() = True Then
        vEnforceAllocation = vRecordSet.Fields(1).Bool
        If vEnforceAllocation Then
          With vInsertFields
            .AddAmendedOnBy(mvEnv.User.UserID)
            .Add("room_booking_number", CDBField.FieldTypes.cftLong, RoomBookingNumber)
            .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
            .Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
            .Add("room_date", CDBField.FieldTypes.cftDate, FromDate)
            .Add("room_id", CDBField.FieldTypes.cftInteger)
            .Add("room_booking_link_number", CDBField.FieldTypes.cftLong)
            vRoomID = 1
            For vRoomID = 1 To NumberOfRooms
              .Item("room_id").Value = CStr(vRoomID)
              For vIndex = 1 To CInt(vRecordSet.Fields(2).Value) 'capacity
                .Item("room_booking_link_number").Value = CStr(mvEnv.GetControlNumber("RL"))
                mvEnv.Connection.InsertRecord("room_booking_links", vInsertFields)
              Next
            Next
          End With
          vCount = CInt(NumberOfRooms * DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FromDate), CDate(ToDate)))
        End If
      End If
      vRecordSet.CloseRecordSet()

      If Not vEnforceAllocation Then
        'Not Enforcing Allocation, so Original format: one room_block_booking per room per night
        vStartDate = CDate(FromDate)
        vNights = 0
        While DateDiff(Microsoft.VisualBasic.DateInterval.Day, vStartDate, CDate(ToDate)) > 0
          'Add the default room user - namely the booker of the event
          With vInsertFields
            .Clear()
            .Add("room_booking_number", CDBField.FieldTypes.cftLong, RoomBookingNumber)
            .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
            .Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
            .Add("room_date", vStartDate)
            .Add("room_id", CDBField.FieldTypes.cftLong, vNights + 1)
            .AddAmendedOnBy(mvEnv.User.UserID)
            .Add("room_booking_link_number", CDBField.FieldTypes.cftLong)
            vQuantity = NumberOfRooms
            While vQuantity > 0
              .Item("room_id").Value = CStr(vNights + 1)
              .Item("room_booking_link_number").Value = CStr(mvEnv.GetControlNumber("RL"))
              mvEnv.Connection.InsertRecord("room_booking_links", vInsertFields)
              vQuantity = vQuantity - 1
              vNights = vNights + 1
            End While
          End With
          vStartDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, vStartDate)
        End While
        vCount = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FromDate), CDate(ToDate)) * NumberOfRooms)
      End If

      'Update the nights available on the block booking
      vWhereFields.Add("block_booking_number", CDBField.FieldTypes.cftLong, BlockBookingNumber)
      vUpdateFields.Add("nights_available", CDBField.FieldTypes.cftInteger, "nights_available - " & vCount)
      mvEnv.Connection.UpdateRecords("room_block_bookings", vUpdateFields, vWhereFields)
    End Sub
    Private Sub RemoveBookingLinks()
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCount As Integer

      vCount = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FromDate), CDate(ToDate)) * NumberOfRooms)
      vWhereFields.Add("block_booking_number", CDBField.FieldTypes.cftLong, BlockBookingNumber)
      vUpdateFields.Add("nights_available", CDBField.FieldTypes.cftInteger, "nights_available + " & vCount) 'Hack
      mvEnv.Connection.UpdateRecords("room_block_bookings", vUpdateFields, vWhereFields)
      vWhereFields.Clear()
      vWhereFields.Add("room_booking_number", CDBField.FieldTypes.cftLong, RoomBookingNumber)
      mvEnv.Connection.DeleteRecords("room_booking_links", vWhereFields)
    End Sub

    Private Sub CancelOrDelete(ByRef pDelete As Boolean, Optional ByRef pCancellationReason As String = "", Optional ByRef pCancellationSource As String = "", Optional ByRef pCancellationAmount As Double = 0.0#)
      'This routine will cancel or delete the event booking but will not perform any financial updates
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCancellationFee As New CancellationFee

      If Not pDelete Then
        If TransactionProcessed Then
          vCancellationFee.InitFromBooking(mvEnv, pCancellationReason, FromDate)
          If vCancellationFee.Existing Then
            vCancellationFee.AddCancellationFeeTransaction(Batch, BatchTransaction, BatchTransactionAnalysis, pCancellationAmount)
          End If
        End If
      End If
      mvEnv.Connection.StartTransaction()
      RemoveBookingLinks()
      If pDelete Then
        mvClassFields.Delete(mvEnv.Connection)
      Else
        BookingStatus = EventBooking.EventBookingStatuses.ebsCancelled
        mvClassFields.Item(EventAccommodationBookingFields.eabCancellationReason).Value = pCancellationReason
        mvClassFields.Item(EventAccommodationBookingFields.eabCancelledOn).Value = TodaysDate()
        mvClassFields.Item(EventAccommodationBookingFields.eabCancelledBy).Value = mvEnv.User.UserID
        If pCancellationSource.Length > 0 Then mvClassFields.Item(EventAccommodationBookingFields.eabCancellationSource).Value = pCancellationSource
        Save()
      End If
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public ReadOnly Property TransactionProcessed() As Boolean
      Get
        TransactionProcessed = BatchNumber > 0 And TransactionNumber > 0 And LineNumber > 0
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Friend ReadOnly Property Batch() As Batch
      Get
        If mvBatch Is Nothing Then
          mvBatch = New Batch(mvEnv)
          mvBatch.Init(BatchNumber)
        End If
        Batch = mvBatch
      End Get
    End Property

    Friend ReadOnly Property BatchTransaction() As BatchTransaction
      Get
        If mvBT Is Nothing Then
          mvBT = New BatchTransaction(mvEnv)
          mvBT.Init(BatchNumber, TransactionNumber)
        End If
        BatchTransaction = mvBT
      End Get
    End Property

    Friend ReadOnly Property BatchTransactionAnalysis() As BatchTransactionAnalysis
      Get
        If mvBTA Is Nothing Then
          mvBTA = New BatchTransactionAnalysis(mvEnv)
          mvBTA.Init(BatchNumber, TransactionNumber, LineNumber)
        End If
        BatchTransactionAnalysis = mvBTA
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(EventAccommodationBookingFields.eabAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventAccommodationBookingFields.eabAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventAccommodationBookingFields.eabAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(EventAccommodationBookingFields.eabAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(EventAccommodationBookingFields.eabBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BlockBookingNumber() As Integer
      Get
        BlockBookingNumber = mvClassFields.Item(EventAccommodationBookingFields.eabBlockBookingNumber).IntegerValue
      End Get
    End Property

    Public Property BookedDate() As String
      Get
        BookedDate = mvClassFields.Item(EventAccommodationBookingFields.eabBookedDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventAccommodationBookingFields.eabBookedDate).Value = Value
      End Set
    End Property

    Public Property BookingStatus() As EventBooking.EventBookingStatuses
      Get
        BookingStatus = EventBooking.GetBookingStatus((mvClassFields.Item(EventAccommodationBookingFields.eabBookingStatus).Value))
      End Get
      Set(ByVal Value As EventBooking.EventBookingStatuses)
        mvClassFields.Item(EventAccommodationBookingFields.eabBookingStatus).Value = EventBooking.GetBookingStatusCode(Value)
      End Set
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(EventAccommodationBookingFields.eabCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(EventAccommodationBookingFields.eabCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(EventAccommodationBookingFields.eabCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(EventAccommodationBookingFields.eabCancelledOn).Value
      End Get
    End Property

    Public Property ConfirmedDate() As String
      Get
        ConfirmedDate = mvClassFields.Item(EventAccommodationBookingFields.eabConfirmedDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventAccommodationBookingFields.eabConfirmedDate).Value = Value
      End Set
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(EventAccommodationBookingFields.eabContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventAccommodationBookingFields.eabEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property FromDate() As String
      Get
        FromDate = mvClassFields.Item(EventAccommodationBookingFields.eabFromDate).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(EventAccommodationBookingFields.eabLineNumber).IntegerValue
      End Get
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(EventAccommodationBookingFields.eabNotes).MultiLineValue
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventAccommodationBookingFields.eabNotes).Value = Value
      End Set
    End Property

    Public ReadOnly Property NumberOfRooms() As Integer
      Get
        NumberOfRooms = mvClassFields.Item(EventAccommodationBookingFields.eabNumberOfRooms).IntegerValue
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(EventAccommodationBookingFields.eabRate).Value
      End Get
    End Property

    Public ReadOnly Property RoomBookingNumber() As Integer
      Get
        RoomBookingNumber = mvClassFields.Item(EventAccommodationBookingFields.eabRoomBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SalesContactNumber() As Integer
      Get
        SalesContactNumber = mvClassFields.Item(EventAccommodationBookingFields.eabSalesContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ToDate() As String
      Get
        ToDate = mvClassFields.Item(EventAccommodationBookingFields.eabToDate).Value
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(EventAccommodationBookingFields.eabTransactionNumber).IntegerValue
      End Get
    End Property

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByVal pEventNumber As Integer, ByVal pBlockBookingNumber As Integer, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pBookedDate As String, ByVal pFromDate As String, ByVal pToDate As String, ByVal pNumberOfRooms As Integer, ByVal pBookingStatus As String, ByVal pRate As String, Optional ByVal pNotes As String = "", Optional ByVal pAmount As Double = 0, Optional ByVal pSalesContactNumber As Integer = 0)
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        .Item(EventAccommodationBookingFields.eabEventNumber).IntegerValue = pEventNumber
        .Item(EventAccommodationBookingFields.eabBlockBookingNumber).IntegerValue = pBlockBookingNumber
        .Item(EventAccommodationBookingFields.eabContactNumber).IntegerValue = pContactNumber
        .Item(EventAccommodationBookingFields.eabAddressNumber).IntegerValue = pAddressNumber
        .Item(EventAccommodationBookingFields.eabBookedDate).Value = pBookedDate
        .Item(EventAccommodationBookingFields.eabFromDate).Value = pFromDate
        .Item(EventAccommodationBookingFields.eabToDate).Value = pToDate
        .Item(EventAccommodationBookingFields.eabNumberOfRooms).IntegerValue = pNumberOfRooms
        .Item(EventAccommodationBookingFields.eabBookingStatus).Value = pBookingStatus
        .Item(EventAccommodationBookingFields.eabRate).Value = pRate
        If Len(pNotes) > 0 Then .Item(EventAccommodationBookingFields.eabNotes).Value = pNotes
        If pAmount > 0 Then .Item(EventAccommodationBookingFields.eabAmount).Value = CStr(pAmount)
        If pSalesContactNumber > 0 Then .Item(EventAccommodationBookingFields.eabSalesContactNumber).IntegerValue = pSalesContactNumber
      End With
    End Sub

    Public Sub SetBatchTransactionLine(ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, ByRef pLineNumber As Integer)
      With mvClassFields
        .Item(EventAccommodationBookingFields.eabBatchNumber).IntegerValue = pBatchNumber
        .Item(EventAccommodationBookingFields.eabTransactionNumber).IntegerValue = pTransactionNumber
        .Item(EventAccommodationBookingFields.eabLineNumber).IntegerValue = pLineNumber
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      With mvClassFields
        If pParams.Exists("ContactNumber") Then .Item(EventAccommodationBookingFields.eabContactNumber).Value = pParams.ParameterExists("ContactNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(EventAccommodationBookingFields.eabAddressNumber).Value = pParams.ParameterExists("AddressNumber").Value
        If pParams.Exists("ConfirmedOn") Then .Item(EventAccommodationBookingFields.eabConfirmedDate).Value = pParams.ParameterExists("ConfirmedOn").Value
        If pParams.Exists("BookingStatus") Then .Item(EventAccommodationBookingFields.eabBookingStatus).Value = pParams.ParameterExists("BookingStatus").Value
        If pParams.Exists("Notes") Then .Item(EventAccommodationBookingFields.eabNotes).Value = pParams.ParameterExists("Notes").Value
        If pParams.HasValue("SalesContactNumber") Then .Item(EventAccommodationBookingFields.eabSalesContactNumber).Value = pParams.ParameterExists("SalesContactNumber").Value
        If pParams.Exists("BookedOn") Then .Item(EventAccommodationBookingFields.eabBookedDate).Value = pParams.ParameterExists("BookedOn").Value
        If pParams.Exists("Rate") Then .Item(EventAccommodationBookingFields.eabRate).Value = pParams.ParameterExists("Rate").Value
      End With
    End Sub
  End Class
End Namespace
