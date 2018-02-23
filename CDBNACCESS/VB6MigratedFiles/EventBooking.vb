

Namespace Access
  Public Class EventBooking

    Public Enum EventBookingRecordSetTypes 'These are bit values
      ebrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventBookingFields
      ebfAll = 0
      ebfEventNumber
      ebfBookingNumber
      ebfBookingDate
      ebfBatchNumber
      ebfTransactionNumber
      ebfLineNumber
      ebfRate
      ebfOptionNumber
      ebfBookingStatus
      ebfContactNumber
      ebfAddressNumber
      ebfQuantity
      ebfAllocated
      ebfCancellationReason
      ebfCancelledBy
      ebfCancelledOn
      ebfAmendedBy
      ebfAmendedOn
      ebfSalesContactNumber
      ebfCancellationSource
      ebfNotes
      ebfAdultQuantity
      ebfChildQuantity
      ebfStartTime
      ebfEndTime
    End Enum

    Public Enum EventBookingStatuses
      ebsBooked = 1 'F
      ebsWaiting 'W
      ebsBookedTransfer 'X
      ebsBookedAndPaid 'B
      ebsWaitingPaid 'P
      ebsBookedAndPaidTransfer 'Y
      ebsBookedCreditSale 'S
      ebsWaitingCreditSale 'A
      ebsBookedCreditSaleTransfer 'R
      ebsBookedInvoiced 'V
      ebsWaitingInvoiced 'O
      ebsBookedInvoicedTransfer 'D
      ebsExternal 'E
      ebsCancelled 'C
      ebsInterested 'I
      ebsAwaitingAcceptance 'T
      ebsAmended 'U
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvDelegates As Collection
    Private mvSessions As Collection
    Private mvBookingOption As EventBookingOption
    Private mvBatch As Batch
    Private mvBT As BatchTransaction
    Private mvBTA As BatchTransactionAnalysis
    Private mvBookingLines As CollectionList(Of BatchTransactionAnalysis)
    Private mvAdjBatchTransColl As CollectionList(Of BatchTransaction)  'Used for Online Authorisation
    Private mvCancellationFeeTrans As BatchTransaction  'Used for Online Authorisation
    Private mvInCancellation As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "event_bookings"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("booking_number", CDBField.FieldTypes.cftLong)
          .Add("booking_date", CDBField.FieldTypes.cftDate)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("rate")
          .Add("option_number", CDBField.FieldTypes.cftLong)
          .Add("booking_status")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("allocated", CDBField.FieldTypes.cftInteger)
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("sales_contact_number", CDBField.FieldTypes.cftLong)
          .Add("cancellation_source")
          .Add("notes")
          .Add("adult_quantity", CDBField.FieldTypes.cftInteger)
          .Add("child_quantity", CDBField.FieldTypes.cftInteger)
          .Add("start_time", CDBField.FieldTypes.cftTime)
          .Add("end_time", CDBField.FieldTypes.cftTime)

          .Item(EventBookingFields.ebfEventNumber).SetPrimaryKeyOnly()
          .Item(EventBookingFields.ebfBookingNumber).SetPrimaryKeyOnly()

          .Item(EventBookingFields.ebfContactNumber).PrefixRequired = True
          .Item(EventBookingFields.ebfAddressNumber).PrefixRequired = True
          .Item(EventBookingFields.ebfAmendedBy).PrefixRequired = True
          .Item(EventBookingFields.ebfAmendedOn).PrefixRequired = True
          .Item(EventBookingFields.ebfAllocated).PrefixRequired = True

          .Item(EventBookingFields.ebfNotes).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventBookingNotes)
          .Item(EventBookingFields.ebfAdultQuantity).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity)
          .Item(EventBookingFields.ebfChildQuantity).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity)
          .Item(EventBookingFields.ebfStartTime).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix)
          .Item(EventBookingFields.ebfEndTime).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvDelegates = Nothing
      mvSessions = Nothing
      mvBookingOption = Nothing
      mvBatch = Nothing
      mvBT = Nothing
      mvBTA = Nothing
      mvBookingLines = New CollectionList(Of BatchTransactionAnalysis)
      mvExisting = False
      mvInCancellation = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventBookingFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventBookingFields.ebfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventBookingFields.ebfAmendedBy).Value = mvEnv.User.UserID
    End Sub


    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventBookingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventBookingRecordSetTypes.ebrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "eb")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pEventNumber > 0 Or pBookingNumber > 0 Then
        If pEventNumber > 0 Then vWhereFields.Add("event_number", CDBField.FieldTypes.cftLong, pEventNumber)
        If pBookingNumber > 0 Then vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, pBookingNumber)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventBookingRecordSetTypes.ebrtAll) & " FROM event_bookings eb WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventBookingRecordSetTypes.ebrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventBookingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventBookingFields.ebfEventNumber, vFields)
        .SetItem(EventBookingFields.ebfBookingNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventBookingRecordSetTypes.ebrtAll) = EventBookingRecordSetTypes.ebrtAll Then
          .SetItem(EventBookingFields.ebfBookingDate, vFields)
          .SetItem(EventBookingFields.ebfBatchNumber, vFields)
          .SetItem(EventBookingFields.ebfTransactionNumber, vFields)
          .SetItem(EventBookingFields.ebfLineNumber, vFields)
          .SetItem(EventBookingFields.ebfRate, vFields)
          .SetItem(EventBookingFields.ebfOptionNumber, vFields)
          .SetItem(EventBookingFields.ebfBookingStatus, vFields)
          .SetItem(EventBookingFields.ebfContactNumber, vFields)
          .SetItem(EventBookingFields.ebfAddressNumber, vFields)
          .SetItem(EventBookingFields.ebfQuantity, vFields)
          .SetItem(EventBookingFields.ebfAllocated, vFields)
          .SetItem(EventBookingFields.ebfCancellationReason, vFields)
          .SetItem(EventBookingFields.ebfCancelledBy, vFields)
          .SetItem(EventBookingFields.ebfCancelledOn, vFields)
          .SetItem(EventBookingFields.ebfAmendedBy, vFields)
          .SetItem(EventBookingFields.ebfAmendedOn, vFields)
          .SetItem(EventBookingFields.ebfSalesContactNumber, vFields)
          .SetOptionalItem(EventBookingFields.ebfCancellationSource, vFields)
          .SetOptionalItem(EventBookingFields.ebfNotes, vFields)
          .SetOptionalItem(EventBookingFields.ebfAdultQuantity, vFields)
          .SetOptionalItem(EventBookingFields.ebfChildQuantity, vFields)
          .SetOptionalItem(EventBookingFields.ebfStartTime, vFields)
          .SetOptionalItem(EventBookingFields.ebfEndTime, vFields)
        End If
      End With
    End Sub

    Public Sub InitNewBooking(ByVal pEnv As CDBEnvironment, ByVal pEvent As CDBEvent)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvClassFields.Item(EventBookingFields.ebfEventNumber).Value = CStr(pEvent.EventNumber)
      mvClassFields.Item(EventBookingFields.ebfBookingNumber).Value = CStr(pEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfBookingNumber))
      mvClassFields.Item(EventBookingFields.ebfBookingDate).Value = TodaysDate()
    End Sub

    Public Sub ModifyBooking(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pQuantity As Integer, ByRef pOptionNumber As Integer, ByRef pStatus As EventBooking.EventBookingStatuses, ByRef pRateCode As String, Optional ByVal pNotes As String = "", Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0, Optional ByRef pAdultQuantity As String = "", Optional ByRef pChildQuantity As String = "", Optional ByVal pBookingDate As String = "", Optional ByVal pStartTime As String = "", Optional ByVal pEndTime As String = "", Optional ByVal pSalesContactNumber As Integer = 0)

      mvClassFields.Item(EventBookingFields.ebfContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(EventBookingFields.ebfAddressNumber).Value = CStr(pAddressNumber)
      mvClassFields.Item(EventBookingFields.ebfQuantity).Value = CStr(pQuantity)
      mvClassFields.Item(EventBookingFields.ebfOptionNumber).Value = CStr(pOptionNumber)
      mvClassFields.Item(EventBookingFields.ebfNotes).Value = pNotes
      BookingStatus = pStatus
      mvClassFields.Item(EventBookingFields.ebfRate).Value = pRateCode
      If pBatchNumber > 0 Then
        mvClassFields.Item(EventBookingFields.ebfBatchNumber).Value = CStr(pBatchNumber)
        mvClassFields.Item(EventBookingFields.ebfTransactionNumber).Value = CStr(pTransNumber)
        mvClassFields.Item(EventBookingFields.ebfLineNumber).Value = CStr(pLineNumber)
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity) Then
        If Len(pAdultQuantity) > 0 Then mvClassFields.Item(EventBookingFields.ebfAdultQuantity).Value = CStr(Val(pAdultQuantity))
        If Len(pChildQuantity) > 0 Then mvClassFields.Item(EventBookingFields.ebfChildQuantity).Value = CStr(Val(pChildQuantity))
      End If
      If Len(pBookingDate) > 0 Then mvClassFields.Item(EventBookingFields.ebfBookingDate).Value = pBookingDate
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
        mvClassFields.Item(EventBookingFields.ebfStartTime).Value = pStartTime
        mvClassFields.Item(EventBookingFields.ebfEndTime).Value = pEndTime
      End If
      If pSalesContactNumber > 0 Then mvClassFields.Item(EventBookingFields.ebfSalesContactNumber).IntegerValue = pSalesContactNumber
    End Sub

    Public Function AddDelegate(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pPosition As String, ByVal pOrganisationName As String, ByVal pSource As String, ByVal pOptionNumber As Integer, Optional ByVal pEventDelegateNumber As Integer = 0, Optional ByRef pAmendedBy As String = "", Optional ByVal pPledgedAmount As Double = 0, Optional ByVal pCreateCalByOptSessions As Boolean = False, Optional ByVal pStatus As EventBooking.EventBookingStatuses = 0, Optional ByVal pSequenceNumber As String = "") As EventDelegate
      'BR11719 - Added optional status parameter and code to deal with it
      Dim vDelegate As New EventDelegate
      Dim vContact As New Contact(mvEnv)

      vContact.Init(pContactNumber)
      If vContact.ContactType = Contact.ContactTypes.ctcJoint Then
        RaiseError(DataAccessErrors.daeCannotBookJointToEvent, vContact.ContactNumber & " : " & vContact.LabelName, CStr(mvClassFields(EventBookingFields.ebfEventNumber).IntegerValue))
      End If

      ' BR11719
      AddDelegateCalendar(pContactNumber, pOptionNumber, pCreateCalByOptSessions, pStatus)
      vDelegate.Init(mvEnv)
      vDelegate.InitFromBooking(Me, pContactNumber, pAddressNumber, pPosition, pOrganisationName, pEventDelegateNumber, pPledgedAmount, pSequenceNumber)
      vDelegate.Save(pAmendedBy)
      If mvDelegates Is Nothing Then mvDelegates = New Collection
      mvDelegates.Add(vDelegate, CStr(pContactNumber))
      mvClassFields.Item(EventBookingFields.ebfAllocated).IntegerValue = Allocated + 1
      SetDelegateSessionsAndActivities(pOptionNumber, pSource, "", (vDelegate.EventDelegateNumber), pContactNumber)
      AddDelegate = vDelegate
    End Function

    Public Sub SetDelegateSessionsAndActivities(ByVal pOptionNumber As Integer, ByVal pSource As String, Optional ByVal pSessionList As String = "", Optional ByRef pDelegateNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0)
      Dim vSession As New EventSession
      Dim vDelegateSession As New DelegateSession
      Dim vSQL As String
      Dim vDelegate As EventDelegate
      Dim vDelegateContacts As String = ""
      Dim vIndex As Integer

      If Len(pSessionList) = 0 Then
        For Each vSession In Sessions
          If vSession.SessionType <> vSession.BaseSessionType Then
            If pSessionList.Length > 0 Then pSessionList = pSessionList & ", "
            pSessionList = pSessionList & "'" & vSession.SessionNumber & "'"
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) And pDelegateNumber > 0 Then 'Used by EventBooking.AddDelegate method
              With vDelegateSession
                .Init(mvEnv)
                .Create(pDelegateNumber, vSession.SessionNumber)
                .Save()
              End With
            End If
          End If
        Next vSession
      End If

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) And pDelegateNumber = 0 Then 'Used by CDBEvent.ChangeSessionBookings method
        vSQL = "INSERT INTO delegate_sessions(event_delegate_number,session_number,attended)"
        vSQL = vSQL & " SELECT d.event_delegate_number, sb.session_number,'N' FROM session_bookings sb,sessions s,delegates d"
        vSQL = vSQL & " WHERE sb.booking_number = " & BookingNumber & " AND sb.session_number = s.session_number AND d.booking_number = sb.booking_number"
        vSQL = vSQL & " AND s.session_type <> '0'" 'Don't include Base Type Session
        If Len(pSessionList) > 0 Then vSQL = vSQL & " AND sb.session_number IN (" & pSessionList & ")"
        vSQL = vSQL & " ORDER by d.event_delegate_number,sb.session_number"
        mvEnv.Connection.ExecuteSQL(vSQL)
      End If

      'Delegate Activities
      If pSource.Length > 0 Then
        Dim vInnerAttrs As String = "activity"
        Dim vInnerWhereFields As New CDBFields(New CDBField("contact_number", CDBField.FieldTypes.cftInteger))
        vInnerWhereFields.AddJoin("cc.activity", "sa.activity")
        vInnerWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)

        Dim vAttrs As String = "DISTINCT sa.activity, sa.activity_value"
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("session_activities sa", "os.session_number", "sa.session_number")

        Dim vWhereFields As New CDBFields(New CDBField("os.option_number", pOptionNumber))
        If pSessionList.Length > 0 Then
          vWhereFields.Add("os.session_number", CDBField.FieldTypes.cftInteger, pSessionList, CDBField.FieldWhereOperators.fwoIn)
        End If
        vWhereFields.Add("sa.activity", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotIn)

        Dim vBulkUpdater As New BulkUpdater(mvEnv)
        Dim vParams As New CDBParameters()
        vParams.Add("ContactNumber", CDBField.FieldTypes.cftInteger)
        vParams.Add("Activity")
        vParams.Add("ActivityValue")
        vParams.Add("Source", pSource)
        vParams.Add("ValidFrom", CDBField.FieldTypes.cftDate, TodaysDate())
        vParams.Add("ValidTo", CDBField.FieldTypes.cftDate, DateSerial(9999, 12, 31).ToString(CAREDateFormat))

        If pContactNumber = 0 Then 'Used by CDBEvent.ChangeSessionBookings method
          For Each vDelegate In Delegates
            If Len(vDelegateContacts) > 0 Then vDelegateContacts = vDelegateContacts & ","
            vDelegateContacts = vDelegateContacts & vDelegate.ContactNumber
          Next vDelegate
        Else
          vDelegateContacts = CStr(pContactNumber) 'Used by EventBooking.AddDelegate method
        End If

        Dim vCC As New ContactCategory(mvEnv)
        Dim vRS As CDBRecordSet
        Dim vSQLStatement As SQLStatement = Nothing
        Dim vInnerSQLStatement As SQLStatement = Nothing
        For vIndex = 0 To UBound(Split(vDelegateContacts, ","))
          vInnerWhereFields("contact_number").Value = vDelegateContacts(vIndex)
          vInnerSQLStatement = New SQLStatement(mvEnv.Connection, vInnerAttrs, "contact_categories cc", vInnerWhereFields)
          vWhereFields("sa.activity").Value = "(" & vInnerSQLStatement.SQL & ")"
          vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "option_sessions os", vWhereFields, "", vAnsiJoins)
          vRS = vSQLStatement.GetRecordSet()
          While vRS.Fetch
            vParams("ContactNumber").Value = Split(vDelegateContacts, ",")(vIndex)
            vParams("Activity").Value = vRS.Fields("activity").Value
            vParams("ActivityValue").Value = vRS.Fields("activity_value").Value
            vCC = New ContactCategory(mvEnv)
            vCC.Create(vParams)
            vBulkUpdater.AddItem(vCC)
          End While
          vRS.CloseRecordSet()
        Next
        vCC = New ContactCategory(mvEnv)
        vAttrs = vCC.FieldNames
        vWhereFields = New CDBFields(New CDBField("contact_number", vParams("ContactNumber").IntegerValue))
        vWhereFields.Add("activity", vParams("Activity").Value)
        vWhereFields.Add("activity_value", vParams("ActivityValue").Value)
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "contact_categories cc", vWhereFields)
        vBulkUpdater.SaveBulkUpdate(vSQLStatement)
      End If
    End Sub

    Public Sub AddDelegateCalendar(ByVal pContactNumber As Integer, Optional ByVal pOptionNumber As Integer = 0, Optional ByVal pCreateCalByOptSessions As Boolean = False, Optional ByVal pStatus As EventBooking.EventBookingStatuses = 0, Optional ByVal pConvertInterestedBooking As Boolean = False)
      'BR11719 - Added optional status parameter and code to deal with it
      Dim vTables As String
      Dim vWhere As String
      Dim vStart As String
      Dim vEnd As String
      Dim vDesc As String
      Dim vRecordSet As CDBRecordSet
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vEventNumber As Integer
      Dim vSessionNumber As Integer

      If mvEnv.GetConfigOption("ev_delegate_calendar", True) Then
        If pCreateCalByOptSessions = True And pOptionNumber > 0 Then
          vTables = "option_sessions os,sessions s, events e"
          vWhere = "os.option_number = " & pOptionNumber & " AND s.session_number = os.session_number AND e.event_number = s.event_number"
        Else
          vTables = "session_bookings sb,sessions s, events e"
          vWhere = "sb.booking_number = " & BookingNumber & " AND s.session_number = sb.session_number AND e.event_number = s.event_number"
        End If
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT s.start_date,start_time,end_date,end_time,s.event_number,s.session_number,session_desc,event_desc FROM " & vTables & " WHERE " & vWhere)
        While vRecordSet.Fetch() = True
          vStart = vRecordSet.Fields(1).Value & " " & vRecordSet.Fields(2).Value
          vEnd = vRecordSet.Fields(3).Value & " " & vRecordSet.Fields(4).Value
          vEventNumber = vRecordSet.Fields(5).IntegerValue
          vSessionNumber = vRecordSet.Fields(6).IntegerValue
          vAppointment.Init()
          Dim vConflict As Boolean = vAppointment.CheckCalendarConflict(pContactNumber, vStart, vEnd, ContactAppointment.ContactAppointmentTypes.catEvent, vSessionNumber, True, pConvertInterestedBooking)
          If vConflict Then
            If vSessionNumber = vEventNumber * 10000 Then
              vDesc = vRecordSet.Fields(8).Value
            Else
              vDesc = vRecordSet.Fields(8).Value & " : " & vRecordSet.Fields(7).Value
            End If
            vAppointment.Init()
            ' BR11719
            vAppointment.Create(pContactNumber, vStart, vEnd, ContactAppointment.ContactAppointmentTypes.catEvent, vDesc, vSessionNumber, ContactAppointment.ContactAppointmentTimeStatuses.catsNone, "", pStatus)
            vAppointment.Save()
          End If
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Public Sub RemoveDelegate(ByRef pDelegate As EventDelegate, Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vSession As EventSession
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vDelegateSession As New DelegateSession

      If mvEnv.GetConfigOption("ev_delegate_calendar", True) Then
        vAppointment.Init()
        For Each vSession In Sessions
          'Remove any appointments for the delegate
          vAppointment.ClearEntries(ContactAppointment.ContactAppointmentTypes.catEvent, vSession.SessionNumber, (pDelegate.ContactNumber))
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
            With vDelegateSession
              .Init(mvEnv, (pDelegate.EventDelegateNumber), (vSession.SessionNumber))
              If .Existing Then .Delete()
            End With
          End If
        Next vSession
      End If
      mvDelegates.Remove(CStr(pDelegate.ContactNumber))
      pDelegate.Delete(pAmendedBy, pAHC)
      mvClassFields.Item(EventBookingFields.ebfAllocated).Value = CStr(Allocated - 1)
    End Sub

    Public Sub CheckCanCancel()
      Dim vDelegate As EventDelegate

      If BookingStatus = EventBooking.EventBookingStatuses.ebsCancelled Then
        RaiseError(DataAccessErrors.daeBookingAlreadyCancelled, CStr(BookingNumber))
      End If
      For Each vDelegate In Delegates
        If vDelegate.Attended = True Then RaiseError(DataAccessErrors.daeDelegatesAlreadyAttended)
      Next vDelegate
      If Not Batch.PostedToNominal And BatchNumber > 0 Then
        RaiseError(DataAccessErrors.daeBookingBatchNotPosted, CStr(BatchNumber))
      End If
    End Sub

    Public Sub ProcessXFerBooking(ByRef pNewEvent As CDBEvent, ByRef pAdjustmentParams As CDBParameters, ByRef pCancellationReason As String, ByRef pCancellationSource As String, ByRef pCancellationFeeAmount As Double)
      Dim vFinHist As New FinancialHistory
      Dim vEvent As New CDBEvent(mvEnv)
      Dim vMsg As String = ""
      Dim vCancellationFee As New CancellationFee

      CheckCanCancel()

      mvInCancellation = True
      vCancellationFee = GetCancellationFee(pCancellationReason)
      If vCancellationFee.Existing Then
        If vCancellationFee.IsCancellationAmountRequired And pCancellationFeeAmount = 0 Then RaiseError(DataAccessErrors.daeCancellationFeeMissing)
      End If
      ProcessAdjustment(pAdjustmentParams)
      CancelOrDelete(False, pCancellationReason, pCancellationSource, True, vCancellationFee, pCancellationFeeAmount)
      'If Not Transferring or Transferring to a different Event then assign from waiting list
      If pNewEvent.EventNumber <> EventNumber Then
        vEvent.Init(EventNumber)
        Select Case vEvent.UserWaitingListMethod
          Case CDBEvent.UserWaitingListMethods.euwlAutomatic
            vEvent.ProcessWaitingList(vMsg)
          Case CDBEvent.UserWaitingListMethods.euwlManual
            '
        End Select
      End If
      mvInCancellation = False
    End Sub

    Public Sub Delete()
      CancelOrDelete(True, "", "", False, Nothing, 0)
    End Sub

    Public Sub Cancel(ByRef pCancellationReason As String, ByRef pCancellationSource As String, Optional ByRef pLeaveTransaction As Boolean = False, Optional ByVal pDoAdjustment As Boolean = False, Optional ByVal pAdjustmentParams As CDBParameters = Nothing, Optional ByRef pCancellationFee As Double = 0.0#, Optional ByVal pCanApplyCancellationFee As Boolean = True)
      Dim vCancellationFee As New CancellationFee

      If Len(CancellationReason) > 0 Then RaiseError(DataAccessErrors.daeBookingAlreadyCancelled, CStr(BookingNumber))
      mvInCancellation = True
      If pCanApplyCancellationFee Then
        vCancellationFee = GetCancellationFee(pCancellationReason)
        If vCancellationFee.Existing Then
          If vCancellationFee.IsCancellationAmountRequired And pCancellationFee = 0 Then RaiseError(DataAccessErrors.daeCancellationFeeMissing)
        End If
      End If
      If pDoAdjustment Then ProcessAdjustment(pAdjustmentParams)
      CancelOrDelete(False, pCancellationReason, pCancellationSource, pLeaveTransaction, vCancellationFee, pCancellationFee)
      mvInCancellation = False
    End Sub

    Private Function GetCancellationFee(ByRef pCancellationReason As String) As CancellationFee
      Dim vCancellationFee As New CancellationFee
      Dim vDate As String
      Dim vSession As EventSession

      If TransactionProcessed Then
        If Not Batch.Provisional Then
          If BookingAmount() > 0 Then
            If Sessions.Count > 0 Then
              vDate = CType(Sessions.Item(1), EventSession).StartDate 'Find the earliest session date
              For Each vSession In mvSessions
                If CDate(vSession.StartDate) < CDate(vDate) Then vDate = vSession.StartDate
              Next vSession
            Else
              Dim vEvent As New CDBEvent(mvEnv)
              vEvent.Init(EventNumber)
              vDate = vEvent.StartDate
            End If
            vCancellationFee.InitFromBooking(mvEnv, pCancellationReason, vDate)
          End If
        End If
      End If
      GetCancellationFee = vCancellationFee
    End Function

    Private Sub CancelOrDelete(ByRef pDelete As Boolean, ByRef pCancellationReason As String, ByRef pCancellationSource As String, ByRef pLeaveTransaction As Boolean, ByRef pCancellationFee As CancellationFee, ByRef pCancellationAmount As Double)
      'This routine will cancel or delete the event booking but will not perform any financial updates
      Dim vCount As Integer
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vDelegate As EventDelegate
      Dim vSession As EventSession
      Dim vWhereFields As New CDBFields
      Dim vCancellationFee As New CancellationFee
      Dim vTransaction As Boolean
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vContactCategory As ContactCategory
      Dim vEvent As CDBEvent
      Dim vRate As ProductRate
      Dim vBaseTypeFound As Boolean
      Dim vSessionCategories As New CollectionList(Of ContactCategory)

      'Get the sessions and the delegates they are booked on and init the booking option
      'This is done here to force the reads outside of any transaction

      If mvEnv.GetConfigOption("ev_delegate_calendar", True) Then vAppointment.Init()
      vCount = Sessions.Count()
      vCount = Delegates.Count()
      vCount = BookingOption.MaximumBookings

      If Not pDelete Then
        If TransactionProcessed Then
          If Batch.Provisional Then
            vBTA.Init(BatchNumber, TransactionNumber, LineNumber)
          Else
            If pCancellationFee Is Nothing Then pCancellationFee = GetCancellationFee(pCancellationReason)
            If pCancellationFee.Existing Then
              pCancellationFee.AddCancellationFeeTransaction(Batch, BatchTransaction, BatchTransactionAnalysis, pCancellationAmount, mvCancellationFeeTrans)
            End If
          End If
        End If
      End If

      'BR12637: Only add a base session type session when mvSessions does not already have it
      vEvent = New CDBEvent(mvEnv)
      vEvent.Init(EventNumber)
      If BookingOption.DeductFromEvent Then
        For Each vSession In mvSessions
          If vSession.SessionType = vSession.BaseSessionType Then
            vBaseTypeFound = True
            Exit For
          End If
        Next vSession
        If vBaseTypeFound = False Then
          vSession = New EventSession
          'vSession.Init(mvEnv, CInt((Int(BookingNumber / 10000) * 10000)))
          vSession.Init(mvEnv, vEvent.LowestSessionNumber)
          mvSessions.Add(vSession)
        End If
      End If

      If vEvent.Source.Length > 0 Then
        vContactCategory = New ContactCategory(mvEnv)
        vContactCategory.Init()
        Dim vAttrs As String = vContactCategory.GetRecordSetFields()

        Dim vConCatWhereFields As New CDBFields(New CDBField("d.booking_number", Me.BookingNumber))
        vConCatWhereFields.Add("cc.source", vEvent.Source)
        vConCatWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan)

        Dim vConCatAnsiJoins As New AnsiJoins
        vConCatAnsiJoins.Add("session_bookings sb", "sb.event_number", "d.event_number", "sb.booking_number", "d.booking_number")
        vConCatAnsiJoins.Add("session_activities sa", "sa.session_number", "sb.session_number")
        vConCatAnsiJoins.Add("contact_categories cc", "cc.contact_number", "d.contact_number", "cc.activity", "sa.activity", "cc.activity_value", "sa.activity_value")

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "delegates d", vConCatWhereFields, "", vConCatAnsiJoins)
        vSQLStatement.Distinct = True

        Dim vContactCatRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        While vContactCatRS.Fetch
          vContactCategory = New ContactCategory(mvEnv)
          vContactCategory.InitFromRecordSet(vContactCatRS)
          vSessionCategories.Add(vContactCategory.ContactCategoryNumber.ToString, vContactCategory)
        End While
        vContactCatRS.CloseRecordSet()
      End If

      If Not mvEnv.Connection.InTransaction Then
        mvEnv.Connection.StartTransaction()
        vTransaction = True
      End If

      For Each vSession In mvSessions
        'Remove any appointments for the delegates
        If mvEnv.GetConfigOption("ev_delegate_calendar", True) Then
          For Each vDelegate In mvDelegates
            vAppointment.ClearEntries(ContactAppointment.ContactAppointmentTypes.catEvent, vSession.SessionNumber, (vDelegate.ContactNumber))
          Next vDelegate
        End If
        'If not deducting then ignore session type 0 (zero)
        If vSession.SessionType <> vSession.BaseSessionType Or BookingOption.DeductFromEvent = True Then
          Select Case BookingStatus
            Case EventBooking.EventBookingStatuses.ebsWaiting, EventBooking.EventBookingStatuses.ebsWaitingPaid, EventBooking.EventBookingStatuses.ebsWaitingCreditSale, EventBooking.EventBookingStatuses.ebsWaitingInvoiced
              vSession.NumberOnWaitingList = vSession.NumberOnWaitingList - Quantity
            Case EventBooking.EventBookingStatuses.ebsInterested, EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
              vSession.NumberInterested = vSession.NumberInterested - Quantity
            Case Else
              vSession.NumberOfAttendees = vSession.NumberOfAttendees - Quantity
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then
                If vSession.SessionType = vSession.BaseSessionType Then
                  vRate = New ProductRate(mvEnv)
                  vRate.Init(BookingOption.ProductCode, RateCode)
                  If vRate.FixedPrice Then
                    vEvent.NumberOfBookings = vEvent.NumberOfBookings - 1
                  Else
                    vEvent.NumberOfBookings = vEvent.NumberOfBookings - Quantity
                  End If
                  vEvent.Save()
                End If
              End If
          End Select
          vSession.Save()
        End If
      Next vSession

      'End any Contact (Delegates) Session Activities
      For Each vContactCategory In vSessionCategories
        If pDelete Then
          vContactCategory.Delete()
        Else
          vContactCategory.Cancel()
          vContactCategory.Save()
        End If
      Next

      For Each vDelegate In mvDelegates
        vDelegate.Delete()
      Next vDelegate

      If pDelete Then
        vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, BookingNumber)
        mvEnv.Connection.DeleteRecords("session_bookings", vWhereFields)
        'If mvEnv.GetDataStructureInfo(cdbDataEventMultipleAnalysis) Then mvEnv.Connection.DeleteRecords "event_booking_transactions", vWhereFields, False
        mvClassFields.Delete(mvEnv.Connection)
      Else
        BookingStatus = EventBooking.EventBookingStatuses.ebsCancelled
        mvClassFields.Item(EventBookingFields.ebfCancellationReason).Value = pCancellationReason
        mvClassFields.Item(EventBookingFields.ebfCancelledOn).Value = TodaysDate()
        mvClassFields.Item(EventBookingFields.ebfCancelledBy).Value = mvEnv.User.UserID
        If pCancellationSource.Length > 0 Then mvClassFields.Item(EventBookingFields.ebfCancellationSource).Value = pCancellationSource
        Save()
      End If
      If vBTA.Existing Then
        vBTA.DeleteFromBatch() 'Remove provisional transaction
        If Len(vEvent.EventPricingMatrix) > 0 Then
          'Delete any EventPricingMatrix lines
          DeleteEventPriceMatrixBTALines()
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
          vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
          mvEnv.Connection.DeleteRecords("event_booking_transactions", vWhereFields, False)
        End If
      End If
      If pDelete Then
        vWhereFields.Clear()
        vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, BookingNumber)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then mvEnv.Connection.DeleteRecords("event_booking_transactions", vWhereFields, False)
      End If

      If vTransaction And Not pLeaveTransaction Then
        mvEnv.Connection.CommitTransaction()
      End If
      '*** BR10976 TEMP BOOKING COUNT CHECK; Pls report to Tracey if this occurs ***
      If mvEnv.ClientCode = "CARE" Then
        System.Diagnostics.Debug.Assert(vEvent.CheckEventCounts = True, "")
      End If
    End Sub

    Public Sub CancelPostProcess()
      'Used to authorise the cancellation fee transaction
      If mvCancellationFeeTrans IsNot Nothing Then  'This will only be set when adding a cancellation fee transaction and the database is in transaction
        Dim vCardSale As New CardSale(mvEnv)
        vCardSale.Init(mvCancellationFeeTrans.BatchNumber, mvCancellationFeeTrans.TransactionNumber)
        Dim vCCA As New CreditCardAuthorisation
        vCCA.InitFromTransaction(mvEnv, mvCancellationFeeTrans.BatchNumber, mvCancellationFeeTrans.TransactionNumber)
        vCCA.ContactNumber = mvCancellationFeeTrans.ContactNumber
        vCCA.AuthoriseTransaction(vCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatNormal, mvCancellationFeeTrans.Amount, mvCancellationFeeTrans.AddressNumber)
        vCardSale.Save()
        mvCancellationFeeTrans = Nothing
      End If
    End Sub

    Public Sub SetTransactionInfo(ByVal pEnv As CDBEnvironment, ByVal pBookingNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pSalesContactNumber As Integer, ByVal pEventNumber As Integer, Optional ByVal pResetStatusCS As Boolean = False, Optional ByVal pInitialiseRequired As Boolean = True)
      'This routine initialises the object and sets the primary key values
      'To act as though it has just read the data
      'This is used so we can update it without reading it first
      If pResetStatusCS = True Then
        If pInitialiseRequired Then Init(pEnv, 0, pBookingNumber)
        If BookingStatus = EventBooking.EventBookingStatuses.ebsWaitingPaid Then
          BookingStatus = EventBooking.EventBookingStatuses.ebsWaitingCreditSale
        Else
          BookingStatus = EventBooking.EventBookingStatuses.ebsBookedCreditSale
        End If
      Else
        If pInitialiseRequired Then Init(pEnv)
      End If
      mvClassFields(EventBookingFields.ebfBookingNumber).SetValue = CStr(pBookingNumber)
      mvClassFields(EventBookingFields.ebfEventNumber).SetValue = CStr(pEventNumber)  'CStr(Int(pBookingNumber / 10000))
      mvExisting = True
      mvClassFields(EventBookingFields.ebfBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(EventBookingFields.ebfTransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields(EventBookingFields.ebfLineNumber).IntegerValue = pLineNumber
      If pSalesContactNumber > 0 Then mvClassFields(EventBookingFields.ebfSalesContactNumber).IntegerValue = pSalesContactNumber
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      With mvClassFields
        If pParams.Exists("ContactNumber") Then .Item(EventBookingFields.ebfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(EventBookingFields.ebfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("BookingStatus") Then .Item(EventBookingFields.ebfBookingStatus).Value = pParams("BookingStatus").Value
        If pParams.Exists("OptionNumber") Then .Item(EventBookingFields.ebfOptionNumber).Value = pParams("OptionNumber").Value
        If pParams.Exists("Quantity") Then .Item(EventBookingFields.ebfQuantity).Value = pParams("Quantity").Value
        If pParams.Exists("Notes") Then .Item(EventBookingFields.ebfNotes).Value = pParams("Notes").Value
        If pParams.HasValue("SalesContactNumber") Then .Item(EventBookingFields.ebfSalesContactNumber).Value = pParams("SalesContactNumber").Value
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity) Then
          If pParams.HasValue("AdultQuantity") Then .Item(EventBookingFields.ebfAdultQuantity).Value = CStr(pParams("AdultQuantity").IntegerValue)
          If pParams.HasValue("ChildQuantity") Then .Item(EventBookingFields.ebfChildQuantity).Value = CStr(pParams("ChildQuantity").IntegerValue)
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
          If pParams.Exists("StartTime") Then .Item(EventBookingFields.ebfStartTime).Value = pParams("StartTime").Value
          If pParams.Exists("EndTime") Then .Item(EventBookingFields.ebfEndTime).Value = pParams("EndTime").Value
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventBookingFields.ebfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Public ReadOnly Property Batch() As Batch
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

    Public ReadOnly Property BatchTransactionAnalysis() As BatchTransactionAnalysis
      Get
        If mvBTA Is Nothing Then
          mvBTA = New BatchTransactionAnalysis(mvEnv)
          mvBTA.Init(BatchNumber, TransactionNumber, LineNumber)
        End If
        BatchTransactionAnalysis = mvBTA
      End Get
    End Property

    Friend ReadOnly Property BookingOption() As EventBookingOption
      Get
        If mvBookingOption Is Nothing Then
          mvBookingOption = New EventBookingOption
          mvBookingOption.Init(mvEnv, OptionNumber)
        End If
        BookingOption = mvBookingOption
      End Get
    End Property

    Public ReadOnly Property Delegates() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vDelegate As New EventDelegate

        If mvDelegates Is Nothing Then
          vDelegate.Init(mvEnv)
          mvDelegates = New Collection
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vDelegate.GetRecordSetFields(EventDelegate.EventDelegateRecordSetTypes.edrtAll) & " FROM delegates ed WHERE booking_number = " & BookingNumber)
          While vRecordSet.Fetch() = True
            vDelegate = New EventDelegate
            vDelegate.InitFromRecordSet(mvEnv, vRecordSet, EventDelegate.EventDelegateRecordSetTypes.edrtAll)
            mvDelegates.Add(vDelegate, CStr(vDelegate.ContactNumber))
          End While
          vRecordSet.CloseRecordSet()
        End If
        Delegates = mvDelegates
      End Get
    End Property

    Public ReadOnly Property Sessions() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vSession As New EventSession

        If mvSessions Is Nothing Then
          vSession.Init(mvEnv)
          mvSessions = New Collection
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vSession.GetRecordSetFields(EventSession.SessionRecordSetTypes.ssrtAll) & " FROM session_bookings sb, sessions s WHERE booking_number = " & BookingNumber & " AND sb.session_number = s.session_number")
          While vRecordSet.Fetch() = True
            vSession = New EventSession
            vSession.InitFromRecordSet(mvEnv, vRecordSet, EventSession.SessionRecordSetTypes.ssrtAll)
            mvSessions.Add(vSession)
          End While
          vRecordSet.CloseRecordSet()
        End If
        Sessions = mvSessions
      End Get
    End Property

    Public ReadOnly Property TransactionProcessed() As Boolean
      Get
        TransactionProcessed = BatchNumber > 0 And TransactionNumber > 0 And LineNumber > 0
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(EventBookingFields.ebfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Allocated() As Integer
      Get
        Allocated = mvClassFields.Item(EventBookingFields.ebfAllocated).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventBookingFields.ebfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventBookingFields.ebfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(EventBookingFields.ebfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BookingDate() As String
      Get
        BookingDate = mvClassFields.Item(EventBookingFields.ebfBookingDate).Value
      End Get
    End Property

    Public ReadOnly Property BookingNumber() As Integer
      Get
        BookingNumber = mvClassFields.Item(EventBookingFields.ebfBookingNumber).IntegerValue
      End Get
    End Property

    Public Property BookingStatus() As EventBooking.EventBookingStatuses
      Get
        BookingStatus = EventBooking.GetBookingStatus((mvClassFields.Item(EventBookingFields.ebfBookingStatus).Value))
      End Get
      Set(ByVal Value As EventBooking.EventBookingStatuses)
        mvClassFields.Item(EventBookingFields.ebfBookingStatus).Value = EventBooking.GetBookingStatusCode(Value)
      End Set
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(EventBookingFields.ebfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(EventBookingFields.ebfCancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(EventBookingFields.ebfNotes).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(EventBookingFields.ebfCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(EventBookingFields.ebfCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(EventBookingFields.ebfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventBookingFields.ebfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(EventBookingFields.ebfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OptionNumber() As Integer
      Get
        OptionNumber = mvClassFields.Item(EventBookingFields.ebfOptionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(EventBookingFields.ebfQuantity).IntegerValue
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(EventBookingFields.ebfRate).Value
      End Get
    End Property

    Public ReadOnly Property SalesContactNumber() As Integer
      Get
        SalesContactNumber = mvClassFields.Item(EventBookingFields.ebfSalesContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(EventBookingFields.ebfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AdultQuantity() As Integer
      Get
        AdultQuantity = mvClassFields.Item(EventBookingFields.ebfAdultQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ChildQuantity() As Integer
      Get
        ChildQuantity = mvClassFields.Item(EventBookingFields.ebfChildQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StartTime() As String
      Get
        StartTime = mvClassFields.Item(EventBookingFields.ebfStartTime).Value
      End Get
    End Property

    Public ReadOnly Property EndTime() As String
      Get
        EndTime = mvClassFields.Item(EventBookingFields.ebfEndTime).Value
      End Get
    End Property

    Public Function ValidStatusChange(ByRef pNewBookingStatus As EventBooking.EventBookingStatuses) As Boolean
      Dim vValid As Boolean
      Dim vEvent As New CDBEvent(mvEnv)

      vEvent.Init(EventNumber)

      If pNewBookingStatus = BookingStatus Then
        'No change - valid
        vValid = True
      Else
        Select Case BookingStatus
          Case EventBooking.EventBookingStatuses.ebsBooked
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsCancelled, EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
                vValid = vEvent.FreeOfCharge = False
            End Select
          Case EventBooking.EventBookingStatuses.ebsWaiting
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsCancelled, EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsWaitingPaid
                vValid = vEvent.FreeOfCharge = False
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedTransfer
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
                vValid = vEvent.FreeOfCharge = False
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBooked
                vValid = BatchNumber = 0
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedTransfer
                vValid = BatchNumber = 0
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedCreditSale
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedInvoiced, EventBooking.EventBookingStatuses.ebsBookedAndPaid
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsWaitingCreditSale
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced, EventBooking.EventBookingStatuses.ebsWaitingPaid
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer, EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedInvoiced
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsWaitingPaid
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
                vValid = Not (mvEnv.GetConfigOption("fp_use_sales_ledger"))
            End Select
          Case EventBooking.EventBookingStatuses.ebsInterested
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsAwaitingAcceptance, EventBooking.EventBookingStatuses.ebsCancelled, EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsExternal
                vValid = vEvent.External
              Case EventBooking.EventBookingStatuses.ebsBooked
                vValid = BatchNumber = 0
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
                vValid = BatchNumber = 0 And vEvent.FreeOfCharge = False
            End Select
          Case EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
            Select Case pNewBookingStatus
              Case EventBooking.EventBookingStatuses.ebsCancelled, EventBooking.EventBookingStatuses.ebsAmended
                vValid = True
              Case EventBooking.EventBookingStatuses.ebsExternal
                vValid = vEvent.External
              Case EventBooking.EventBookingStatuses.ebsBooked
                vValid = BatchNumber = 0
              Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
                vValid = BatchNumber = 0 And vEvent.FreeOfCharge = False
            End Select
        End Select
      End If
      ValidStatusChange = vValid
    End Function

    Private Sub ProcessAdjustment(ByRef pAdjustmentParams As CDBParameters)
      Dim vAdjustmentType As Batch.AdjustmentTypes
      Dim vFinHist As New FinancialHistory

      Dim vSQL As String
      Dim vRS As CDBRecordSet
      Dim vBatch As Batch
      Dim vTransaction As Boolean
      Dim vBookingAmount As Double
      Dim vIndex As Integer
      Dim vAdjBatchNumber As Integer = 0
      Dim vAdjTransNumber As Integer = 0
      Dim vBTA As BatchTransactionAnalysis
      Dim vInvoicesTrans As CDBParameters = Nothing 'Used in ProcessAdjustmentPreProcess to keep a track of processed invoices
      Dim vUnallocatedCreditNoteTrans As New CDBParameters

      If TransactionProcessed = True And Batch.Provisional = False Then
        If Not pAdjustmentParams.ContainsKey("FullAmountAllocation") Then pAdjustmentParams.Add("FullAmountAllocation", "Y")
        If mvEnv.Connection.InTransaction = False Then
          mvEnv.Connection.StartTransaction()
          vTransaction = True
        End If
        vBookingAmount = BookingAmount()
        Dim vCancelAmount As Double
        Dim vCreateUnallocatedCreditNote As Boolean = Batch.BatchType = Batch.BatchTypes.CreditSales AndAlso pAdjustmentParams.ParameterExists("UnallocateCreditNote").Bool
        If vBookingAmount > 0 Then
          '(1) Adjust any linked Event Booking Transaction that is a credit note (must be reversed first to get the invoice created correctly)
          If Me.Batch.BatchType = Batch.BatchTypes.CreditSales AndAlso pAdjustmentParams.HasValue("AllocationsChecked") Then
            Dim vLinkedTrans As CollectionList(Of FinancialHistory) = Me.GetLinkedTransactions()
            If vLinkedTrans IsNot Nothing AndAlso vLinkedTrans.Count > 0 Then
              For Each vFH As FinancialHistory In vLinkedTrans
                If vFH.TransactionSign.Equals("D", StringComparison.InvariantCultureIgnoreCase) Then
                  Dim vInvoice As New Invoice()
                  vInvoice.Init(mvEnv, vFH.BatchNumber, vFH.TransactionNumber)
                  If vInvoice.Existing AndAlso vInvoice.InvoiceType = Invoice.InvoiceRecordType.CreditNote Then
                    'Reverse the credit note - this removes any allocations and creates an invoice
                    If vFH.Status = FinancialHistory.FinancialHistoryStatus.fhsNormal Then
                      vFH.AdjustTransaction(Batch.AdjustmentTypes.atReverse, pAdjustmentParams, vFH.Amount, 0, False, vAdjBatchNumber, vAdjTransNumber, mvAdjBatchTransColl)
                    End If
                  End If
                End If
              Next
            End If
            vAdjBatchNumber = 0
            vAdjTransNumber = 0
          End If

          '(2) Adjust the Event Booking transaction
          vAdjustmentType = Batch.AdjustmentTypes.atReverse
          If Batch.RefundAllowed = True Then
            If pAdjustmentParams.ContainsKey("RunType") AndAlso pAdjustmentParams("RunType").Value = "V" Then
              'We are cancelling a credit/debit card event booking with the reverse option 
              ' - i.e. no money taken before cancellation, so a card sale record is not required. 
              vAdjustmentType = Batch.AdjustmentTypes.atReverse
            Else
              vAdjustmentType = Batch.AdjustmentTypes.atRefund
            End If
          End If
          ProcessAdjustmentPreProcess(pAdjustmentParams, BatchNumber, TransactionNumber, vInvoicesTrans)  'Check if any invoice payment allocation are to be removed
          vFinHist.Init(mvEnv, BatchNumber, TransactionNumber)
          vFinHist.AdjustTransaction(vAdjustmentType, pAdjustmentParams, vBookingAmount, LineNumber, False, vAdjBatchNumber, vAdjTransNumber, mvAdjBatchTransColl)
          If vCreateUnallocatedCreditNote AndAlso Not vUnallocatedCreditNoteTrans.Exists(vAdjBatchNumber & "|" & vAdjTransNumber) Then vUnallocatedCreditNoteTrans.Add(vAdjBatchNumber & "|" & vAdjTransNumber)
          vCancelAmount = BatchTransactionAnalysis.Amount

          '(3) If the Event Booking has any linked lines, adjust them as well
          If mvBookingLines.Count() > 1 Then
            Dim vIgnore As Boolean = False
            For vIndex = 1 To mvBookingLines.Count - 1
              vIgnore = False
              vBTA = mvBookingLines.Item(vIndex)
              If vBTA.BatchNumber <> vFinHist.BatchNumber Or vBTA.TransactionNumber <> vFinHist.TransactionNumber Then
                vFinHist = New FinancialHistory
                vFinHist.Init(mvEnv, (vBTA.BatchNumber), (vBTA.TransactionNumber))
                If Batch.BatchType = Batch.BatchTypes.CreditSales Then
                  'If the Event Booking was paid by Invoice and the Invoice has not been printed
                  'then set it so that it does not get printed if we have cancelled the entire Invoice
                  Dim vInvoice As New Invoice
                  vInvoice.Init(mvEnv, vBTA.BatchNumber, vBTA.TransactionNumber)
                  If vInvoice.Existing Then
                    If vInvoice.InvoicePrinted = False AndAlso vCancelAmount = vInvoice.InvoiceAmount Then
                      vInvoice.SetInvoicePrintingNotRequired()
                      vInvoice.Save()
                    End If
                    If vInvoice.InvoiceType = Invoice.InvoiceRecordType.CreditNote Then
                      If vFinHist.Status <> FinancialHistory.FinancialHistoryStatus.fhsNormal Then
                        'This was reversed, above, so ignore this FH record
                        vIgnore = True
                      End If
                    End If
                  End If
                End If
                vCancelAmount = 0
                ProcessAdjustmentPreProcess(pAdjustmentParams, vBTA.BatchNumber, vBTA.TransactionNumber, vInvoicesTrans)  'Check if any invoice payment allocation are to be removed
              End If
              If vFinHist.TransactionSign = "D" Then vBTA.ChangeSign()
              If vIgnore = False Then vFinHist.AdjustTransaction(vAdjustmentType, pAdjustmentParams, vBookingAmount, (vBTA.LineNumber), False, vAdjBatchNumber, vAdjTransNumber, mvAdjBatchTransColl)
              If vCreateUnallocatedCreditNote AndAlso Not vUnallocatedCreditNoteTrans.Exists(vAdjBatchNumber & "|" & vAdjTransNumber) Then vUnallocatedCreditNoteTrans.Add(vAdjBatchNumber & "|" & vAdjTransNumber)
              vCancelAmount += vBTA.Amount
            Next
            'Now need to reset Batch totals?
            Dim vBatchDetails As ParameterList = vFinHist.GetBatchTotal(vAdjBatchNumber)
            Dim vAmount As Double = 0
            Dim vCurrencyAmount As Double = 0
            If vBatchDetails IsNot Nothing Then
              If vBatchDetails.ContainsKey("Amount") Then vAmount = Convert.ToDouble(vBatchDetails("Amount"))
              If vBatchDetails.ContainsKey("CurrencyAmount") Then vCurrencyAmount = Convert.ToDouble(vBatchDetails("CurrencyAmount"))
            End If
            vBatch = New Batch(mvEnv)
            vBatch.Init(vAdjBatchNumber)
            With vBatch
              .BatchTotal = vAmount 'Val(mvEnv.Connection.GetValue("SELECT SUM(amount) FROM batch_transactions WHERE batch_number = " & vAdjBatchNumber))
              .CurrencyBatchTotal = vCurrencyAmount 'Val(mvEnv.Connection.GetValue("SELECT SUM(currency_amount) FROM batch_transactions WHERE batch_number = " & vAdjBatchNumber))
              .NumberOfEntries = 0
              .SetBatchTotals()
              .Save()
            End With
            vBatch = Nothing
          End If
        End If

        If Batch.BatchType = Batch.BatchTypes.CreditSales Then
            'If the Event Booking was paid by Invoice and the Invoice has not been printed
            'then set it so that it does not get printed if we have cancelled the entire Invoice
            Dim vInvoice As New Invoice
            vInvoice.Init(mvEnv, vFinHist.BatchNumber, vFinHist.TransactionNumber)
            If vInvoice.Existing = True AndAlso vInvoice.InvoicePrinted = False AndAlso vCancelAmount = vInvoice.InvoiceAmount Then
              vInvoice.SetInvoicePrintingNotRequired()
              vInvoice.Save()
            End If
          End If

          '***** NOT SURE THAT WE ACTUALLY NEED THIS HERE ANYMORE *****
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
          vSQL = "SELECT DISTINCT bta.batch_number,bta.transaction_number,bta.line_number,bta.amount FROM event_booking_transactions ebt,"
          vSQL = vSQL & " batch_transaction_analysis bta, financial_history_details fhd WHERE booking_number = " & BookingNumber
          vSQL = vSQL & " AND ebt.batch_number = bta.batch_number AND ebt.transaction_number = bta.transaction_number AND ebt.line_number = bta.line_number"
          vSQL = vSQL & " AND ebt.batch_number = fhd.batch_number AND ebt.transaction_number = fhd.transaction_number AND ebt.line_number = fhd.line_number"
          vSQL = vSQL & " AND ( fhd.status = '' OR fhd.status IS NULL ) ORDER BY bta.batch_number"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          Dim vPreviousBatchNo As Integer = BatchNumber
          Dim vPreviousTransNo As Integer = TransactionNumber
          vBatch = Batch
          vCancelAmount = 0
          With vRS
            While .Fetch() = True
              If Not (.Fields(1).IntegerValue = BatchNumber And .Fields(2).IntegerValue = TransactionNumber And .Fields(3).IntegerValue = LineNumber) Then
                If .Fields(4).DoubleValue > 0 Then
                  vAdjustmentType = Batch.AdjustmentTypes.atReverse
                  If .Fields(1).IntegerValue <> vPreviousBatchNo Then
                    If vBatch.BatchType = Batch.BatchTypes.CreditSales AndAlso .Fields(2).IntegerValue <> vPreviousTransNo Then
                      'If the Event Booking was paid by Invoice and the Invoice has not been printed
                      'then set it so that it does not get printed if we have cancelled the entire Invoice
                      Dim vInvoice As New Invoice()
                      vInvoice.Init(mvEnv, vPreviousBatchNo, vPreviousTransNo)
                      If vInvoice.Existing = True AndAlso vInvoice.InvoicePrinted = False AndAlso vCancelAmount = vInvoice.InvoiceAmount Then
                        vInvoice.SetInvoicePrintingNotRequired()
                        vInvoice.Save()
                      End If
                      vCancelAmount = 0
                      vPreviousTransNo = .Fields(2).IntegerValue
                    End If
                    vBatch = New Batch(mvEnv)
                    vBatch.Init((.Fields(1).IntegerValue))
                    vPreviousBatchNo = vBatch.BatchNumber
                  End If
                  If vBatch.RefundAllowed = True Then vAdjustmentType = Batch.AdjustmentTypes.atRefund
                  vCreateUnallocatedCreditNote = vBatch.BatchType = Batch.BatchTypes.CreditSales AndAlso pAdjustmentParams.ParameterExists("UnallocateCreditNote").Bool
                  vFinHist = New FinancialHistory
                  vFinHist.Init(mvEnv, vPreviousBatchNo, vPreviousTransNo)
                  ProcessAdjustmentPreProcess(pAdjustmentParams, .Fields(1).IntegerValue, .Fields(2).IntegerValue, vInvoicesTrans)  'Check if any invoice payment allocation are to be removed
                  vAdjBatchNumber = 0
                  vAdjTransNumber = 0
                  vFinHist.AdjustTransaction(vAdjustmentType, pAdjustmentParams, .Fields(4).DoubleValue, (.Fields(3).IntegerValue), False, vAdjBatchNumber, vAdjTransNumber, mvAdjBatchTransColl)
                  If vCreateUnallocatedCreditNote AndAlso Not vUnallocatedCreditNoteTrans.Exists(vAdjBatchNumber & "|" & vAdjTransNumber) Then vUnallocatedCreditNoteTrans.Add(vAdjBatchNumber & "|" & vAdjTransNumber)
                  If vFinHist.TransactionSign = "D" Then
                    vCancelAmount -= .Fields(4).DoubleValue
                  Else
                    vCancelAmount += .Fields(4).DoubleValue
                  End If
                End If
              End If
            End While
          End With
          vRS.CloseRecordSet()
          If vCancelAmount > 0 AndAlso vBatch.BatchType = Batch.BatchTypes.CreditSales Then
            Dim vInvoice As New Invoice()
            vInvoice.Init(mvEnv, vPreviousBatchNo, vPreviousTransNo)
            If vInvoice.Existing = True AndAlso vInvoice.InvoicePrinted = False AndAlso vCancelAmount = vInvoice.InvoiceAmount Then
              vInvoice.SetInvoicePrintingNotRequired()
              vInvoice.Save()
            End If
          End If
        End If
        If vUnallocatedCreditNoteTrans.Count > 0 Then
          For vInt As Integer = 1 To vUnallocatedCreditNoteTrans.Count
            Dim vKey() As String = vUnallocatedCreditNoteTrans.ItemKey(vInt).Split("|"c)
            Batch.WriteInvoiceAndDetails(IntegerValue(vKey(0)), IntegerValue(vKey(1)), True, False, True, True)
          Next
        End If
        If vTransaction Then
          mvEnv.Connection.CommitTransaction()
          ProcessAdjustmentPostProcess()  'Most likely for CancelEventBooking web service. For UpdateEventBooking web service, this is called within the web service function
        End If
      End If
    End Sub

    Public Sub ProcessAdjustmentPostProcess()
      'Used to authorise the adjusted transactions
      If mvAdjBatchTransColl IsNot Nothing Then
        For vIndex As Integer = 0 To mvAdjBatchTransColl.Count - 1
          Dim vKey() As String = mvAdjBatchTransColl.ItemKey(vIndex).Split("|"c)
          FinancialHistory.AdjustTransactionPostProcess(mvEnv, IntegerValue(vKey(0)), IntegerValue(vKey(1)), vKey(2), mvAdjBatchTransColl(vIndex), "")
        Next
        mvAdjBatchTransColl = Nothing
      End If
    End Sub

    Private Sub ProcessAdjustmentPreProcess(ByVal pAdjustmentParams As CDBParameters, ByVal pBatchNumber As Integer, ByVal pTransNumber As Integer, ByRef pInvoicesTrans As CDBParameters)
      'This is to remove any allocated payments for an invoice
      'BR17149: Where the event booking cancellation credit note is being left unallocated - do not remove invoice allocations
      If pAdjustmentParams.HasValue("AllocationsChecked") AndAlso Not pAdjustmentParams.ParameterExists("UnallocateCreditNote").Bool Then  'If there are some payments and the user has accepted
        If pInvoicesTrans Is Nothing Then pInvoicesTrans = New CDBParameters
        If Not pInvoicesTrans.Exists(pBatchNumber & "|" & pTransNumber) Then  'Only remove allocation when we have not already done it
          pInvoicesTrans.Add(pBatchNumber & "|" & pTransNumber)
          Dim vInvoice As New Invoice
          vInvoice.Init(mvEnv, pBatchNumber, pTransNumber)
          If vInvoice.Existing AndAlso Invoice.GetRecordType(vInvoice.RecordType) = Invoice.InvoiceRecordType.Invoice Then
            'When AllocationChecked is set we should know the payments exists but this is to make it sure
            If vInvoice.AllocationsAmount(False, False, False) > 0 Then
              vInvoice.RemoveAllocations()
            End If
          End If
        End If
      End If
    End Sub

    Public Sub AddLinkedTransaction(ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pLinkedAnalysis As CollectionList(Of BatchTransactionAnalysis) = Nothing, Optional ByVal pImportedLines As Collection = Nothing, Optional ByVal pLastLineNo As Integer = 0, Optional ByVal pGotPriceMatrixLines As Boolean = False)
      Dim vSQL As String
      Dim vBTA As BatchTransactionAnalysis
      Dim vIndex As Integer
      Dim vTransStarted As Boolean
      Dim vBTALine() As String
      Dim vEB As EventBooking

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTransStarted = True
      End If
      If pImportedLines Is Nothing Then
        If pFinancialAdjustment = Batch.AdjustmentTypes.atNone Or pFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment Then
          'Used in SC Trader for creating new transactions only
          vSQL = "INSERT INTO event_booking_transactions SELECT "
          vSQL = vSQL & EventNumber & ", " & BookingNumber & ", " & BatchNumber & ", " & TransactionNumber & ", "
          vSQL = vSQL & " line_number, '" & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate()))
        vSQL = vSQL & " FROM batch_transaction_analysis WHERE batch_number = " & BatchNumber
          vSQL = vSQL & " AND transaction_number = " & TransactionNumber & " AND line_number " & If(pGotPriceMatrixLines = True, ">= ", "> ") & LineNumber
          vSQL = vSQL & " AND line_number <= " & pLastLineNo
          mvEnv.Connection.ExecuteSQL(vSQL)
        Else
          If Not pLinkedAnalysis Is Nothing Then
            'Used in SC Trader for FinancialAdjustment only
            For Each vBTA In pLinkedAnalysis
              If BookingNumber <> vBTA.LinkedBookingNo Then
                vEB = New EventBooking
                vEB.Init(mvEnv, 0, (vBTA.LinkedBookingNo))
              Else
                vEB = Me
              End If
              vSQL = "INSERT INTO event_booking_transactions VALUES (" & vEB.EventNumber & ", " & vEB.BookingNumber & ", "
              vSQL = vSQL & vBTA.BatchNumber & ", " & vBTA.TransactionNumber & "," & vBTA.LineNumber & ", '"
              vSQL = vSQL & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
              mvEnv.Connection.ExecuteSQL(vSQL)
            Next vBTA
          End If
        End If
      Else
        'Used in DataImport and FinancialHistory.ProcessAdjustment
        For vIndex = 1 To pImportedLines.Count()
          vBTALine = Split(CStr(pImportedLines.Item(vIndex)), ",")
          vSQL = "INSERT INTO event_booking_transactions VALUES ("

          If UBound(vBTALine) > 3 Then
            vSQL = vSQL & vBTALine(0) & ", " & vBTALine(1) & ", " & vBTALine(2) & ", " & vBTALine(3) & "," & vBTALine(4) & ", '"
          Else
            vSQL = vSQL & EventNumber & ", " & BookingNumber & ", " & vBTALine(0) & ", " & vBTALine(1) & "," & vBTALine(2) & ", '"
          End If

          vSQL = vSQL & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
          mvEnv.Connection.ExecuteSQL(vSQL)
        Next
      End If
      If vTransStarted Then mvEnv.Connection.CommitTransaction()
    End Sub

    Friend Sub AddLinkedTransactionForEPM(ByVal pLastLineNo As Integer)
      Dim vTDRLine As New TraderAnalysisLine
      Dim vSQL As String

      vSQL = "INSERT INTO event_booking_transactions SELECT "
      vSQL = vSQL & EventNumber & ", " & BookingNumber & ", " & BatchNumber & ", " & TransactionNumber & ","
      vSQL = vSQL & " line_number, '" & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate()))
      vSQL = vSQL & " FROM batch_transaction_analysis WHERE batch_number = " & BatchNumber
      vSQL = vSQL & " AND transaction_number = " & TransactionNumber & " AND line_number >= " & LineNumber
      If pLastLineNo > 0 Then vSQL = vSQL & " AND line_number <= " & pLastLineNo
      vSQL = vSQL & " AND member_number = '" & BookingNumber & "'"
      mvEnv.Connection.ExecuteSQL(vSQL)

    End Sub

    Private Sub DeleteEventPriceMatrixBTALines()
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vWhereFields As New CDBFields
      Dim vRS As CDBRecordSet
      Dim vTDRLine As New TraderAnalysisLine
      Dim vSQL As String

      vBTA.Init()
      With vWhereFields
        .Add("event_number", CDBField.FieldTypes.cftLong, EventNumber)
        .Add("booking_number", CDBField.FieldTypes.cftLong, BookingNumber)
        .Add("ebt.batch_number", CDBField.FieldTypes.cftLong, "bta.batch_number")
        .Add("ebt.transaction_number", CDBField.FieldTypes.cftLong, "bta.transaction_number")
        .Add("ebt.line_number", CDBField.FieldTypes.cftLong, "bta.line_number")
        .Add("line_type", CDBField.FieldTypes.cftCharacter, vTDRLine.GetAnalysisLineTypeCode(TraderAnalysisLine.TraderAnalysisLineTypes.taltEventPricingMatrixLine))
      End With
      vSQL = "SELECT " & vBTA.GetRecordSetFields() & " FROM event_booking_transactions ebt, batch_transaction_analysis bta"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vBTA = New BatchTransactionAnalysis(mvEnv)
        vBTA.InitFromRecordSet(vRS)
        vBTA.DeleteFromBatch()
      End While
      vRS.CloseRecordSet()

    End Sub

    Friend Sub DeleteEventPriceMatrixLines()
      'Delete any lines linked to the Event Booking as they will get re-created as part of transferring an interest-only Booking
      Dim vWhereFields As New CDBFields

      If mvExisting = True And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) = True Then
        With vWhereFields
          .Add("event_number", CDBField.FieldTypes.cftLong, EventNumber)
          .Add("booking_number", CDBField.FieldTypes.cftLong, BookingNumber)
        End With
        mvEnv.Connection.DeleteRecords("event_booking_transactions", vWhereFields, False)
      End If

    End Sub

    Public Sub SetBookingAmended(ByRef pEvent As CDBEvent, ByVal pUpdateBookingTotals As Boolean)
      'Set this Booking as Amended
      Dim vNewStatus As EventBooking.EventBookingStatuses
      Dim vRate As ProductRate
      Dim vSession As EventSession

      vNewStatus = EventBooking.EventBookingStatuses.ebsAmended
      If ValidStatusChange(vNewStatus) Then
        If pUpdateBookingTotals = True Then
          'Get the sessions
          Dim vSessions As Collection = Sessions
          If vSessions Is Nothing OrElse vSessions.Count = 0 Then
            'If we didn't pick any sessions then get the Base session
            vSessions = New Collection
            vSessions.Add(pEvent.BaseSession)
          End If

          For Each vSession In vSessions
            'If not deducting then ignore session type 0 (zero)
            If Not (BookingOption.DeductFromEvent = False AndAlso vSession.SessionType = vSession.BaseSessionType) Then
              Select Case BookingStatus
                Case EventBooking.EventBookingStatuses.ebsWaiting, EventBooking.EventBookingStatuses.ebsWaitingPaid, EventBooking.EventBookingStatuses.ebsWaitingCreditSale, EventBooking.EventBookingStatuses.ebsWaitingInvoiced
                  vSession.NumberOnWaitingList = vSession.NumberOnWaitingList - Quantity
                Case EventBooking.EventBookingStatuses.ebsInterested, EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
                  vSession.NumberInterested = vSession.NumberInterested - Quantity
                Case Else
                  vSession.NumberOfAttendees = vSession.NumberOfAttendees - Quantity
                  If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then
                    If vSession.SessionType = vSession.BaseSessionType Then
                      vRate = New ProductRate(mvEnv)
                      vRate.Init(BookingOption.ProductCode, RateCode)
                      If vRate.FixedPrice Then
                        pEvent.NumberOfBookings = pEvent.NumberOfBookings - 1
                      Else
                        pEvent.NumberOfBookings = pEvent.NumberOfBookings - Quantity
                      End If
                      pEvent.Save()
                    End If
                  End If
              End Select
              vSession.Save()
            End If
          Next vSession
        End If
        mvClassFields.Item(EventBookingFields.ebfBookingStatus).Value = EventBooking.GetBookingStatusCode(vNewStatus)
      End If

    End Sub

    Public Function GetBookingAnalysisLines() As CollectionList(Of BatchTransactionAnalysis)
      'Return collection of all BTA lines that make up this Booking
      Dim vColl As New CollectionList(Of BatchTransactionAnalysis)
      Dim vRS As CDBRecordSet
      Dim vFound As Boolean
      Dim vEvent As New CDBEvent(mvEnv)
      Dim vPricingMatrix As Boolean

      If mvBookingLines Is Nothing Then mvBookingLines = New CollectionList(Of BatchTransactionAnalysis)
      If mvExisting = True Then
        If mvBookingLines.Count() = 0 Then
          If BatchNumber > 0 Then
            vEvent.Init(EventNumber)
            If vEvent.Existing = True AndAlso vEvent.EventPricingMatrix.Length > 0 Then vPricingMatrix = True
            If mvInCancellation = False OrElse vPricingMatrix = False Then
              'If we're not cancelling the booking OR the event does not have a pricing matrix then add the current BTA to the collection
              vColl.Add(BatchTransactionAnalysis.Key(mvInCancellation), BatchTransactionAnalysis)
            End If
            'Now add an others
            Dim vWhereFields As New CDBFields()
            With vWhereFields
              .Add("ebt.event_number", EventNumber)
              .Add("ebt.booking_number", BookingNumber)
              .Add("r1.batch_number", CDBField.FieldTypes.cftInteger)
              .Add("r2.batch_number", CDBField.FieldTypes.cftInteger)
            End With
            If mvInCancellation Then
              'Do nothing as we want this SQL to include all of the analysis lines linked to the booking
            ElseIf vPricingMatrix Then
              vWhereFields.Add("bta.line_type", "X")
            Else
              'If vPricingMatrix = False AndAlso mvInCancellation = False Then
              vWhereFields.Add("bta.product", CDBField.FieldTypes.cftInteger, "ebo.product")
              vWhereFields.Add("bta.rate", CDBField.FieldTypes.cftInteger, "eb.rate")
            End If

            Dim vBTA As New BatchTransactionAnalysis(mvEnv)
            Dim vAttrs As String = vBTA.GetRecordSetFields.Replace(",line_number,", ",bta.line_number,") & ", tt.transaction_sign "
            If vPricingMatrix = False Then vAttrs = vAttrs.Replace("sales_contact_number", "bta.sales_contact_number")
            Dim vOrderby As String = "bta.batch_number, bta.transaction_number, bta.line_number"

            Dim vAnsiJoins As New AnsiJoins()
            With vAnsiJoins
              .Add("batch_transaction_analysis bta", "ebt.batch_number", "bta.batch_number", "ebt.transaction_number", "bta.transaction_number", "ebt.line_number", "bta.line_number")
              .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
              .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
              If vPricingMatrix = False Then
                .Add("event_bookings eb", "eb.event_number", "ebt.event_number", "eb.booking_number", "ebt.booking_number")
                .Add("event_booking_options ebo", "ebo.event_number", "eb.event_number", "ebo.option_number", "eb.option_number")
              End If
              .AddLeftOuterJoin("reversals r1", "r1.batch_number", "bta.batch_number", "r1.transaction_number", "bta.transaction_number", "r1.line_number", "bta.line_number")
              .AddLeftOuterJoin("reversals r2", "r2.was_batch_number", "ebt.batch_number", "r2.was_transaction_number", "ebt.transaction_number", "r2.was_line_number", "ebt.line_number")
            End With
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_booking_transactions ebt", vWhereFields, vOrderby, vAnsiJoins)
            vRS = vSQLStatement.GetRecordSet()
            While vRS.Fetch() = True
              vBTA = New BatchTransactionAnalysis(mvEnv)
              vBTA.InitFromRecordSet(vRS)
              If vRS.Fields("transaction_sign").Value = "D" Then vBTA.ChangeSign()
              vFound = False
              For Each vBTA2 As BatchTransactionAnalysis In vColl
                If vBTA2.ProductCode = vBTA.ProductCode And vBTA2.RateCode = vBTA.RateCode Then
                  If mvInCancellation = False Then
                    If Not vPricingMatrix Then
                      vFound = True
                    ElseIf Len(vBTA2.Notes) > 0 And Len(vBTA.Notes) > 0 Then
                      If vBTA2.Notes = vBTA.Notes Then vFound = True
                    End If
                  End If
                End If
                If vFound Then
                  vBTA2.Quantity = vBTA2.Quantity + vBTA.Quantity
                  vBTA2.Issued = vBTA2.Quantity
                  vBTA2.CurrencyAmount = FixTwoPlaces(vBTA2.CurrencyAmount + vBTA.CurrencyAmount)
                  vBTA2.CurrencyVatAmount = FixTwoPlaces(vBTA2.CurrencyVatAmount + vBTA.CurrencyVatAmount)
                  vBTA2.Amount = FixTwoPlaces(vBTA2.Amount + vBTA.Amount)
                  vBTA2.VatAmount = FixTwoPlaces(vBTA2.VatAmount + vBTA.VatAmount)
                End If
                If vFound Then Exit For
              Next vBTA2
              If vFound = False Then vColl.Add(vBTA.Key(mvInCancellation), vBTA)
            End While
            vRS.CloseRecordSet()
            mvBookingLines = vColl
            vColl = Nothing
          End If
        End If
      End If
      Return mvBookingLines
    End Function

    Private Function BookingAmount() As Double
      Dim vBTA As BatchTransactionAnalysis
      Dim vAmount As Double

      vAmount = BatchTransactionAnalysis.Amount
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
        GetBookingAnalysisLines()
        If mvBookingLines.Count() > 0 Then
          vAmount = 0
          For Each vBTA In mvBookingLines
            vAmount = FixTwoPlaces(vAmount + vBTA.Amount)
          Next vBTA
        End If
      End If

      BookingAmount = vAmount
    End Function

    Friend Function GetLinkedTransactions() As CollectionList(Of FinancialHistory)
      Dim vColl As New CollectionList(Of FinancialHistory)
      Dim vFH As New FinancialHistory()
      vFH.Init(mvEnv)

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("financial_history fh", "ebt.batch_number", "fh.batch_number", "ebt.transaction_number", "fh.transaction_number")})
      vAnsiJoins.Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")
      'vAnsiJoins.Add("batches b", "fh.batch_number", "b.batch_number")

      Dim vWhereFields As New CDBFields({New CDBField("ebt.event_number", Me.EventNumber), New CDBField("ebt.booking_number", Me.BookingNumber)})
      vWhereFields.Add("ebt.batch_number", CDBField.FieldTypes.cftInteger, Me.BatchNumber.ToString, CDBField.FieldWhereOperators.fwoNotEqual)

      Dim vAttrs As String = vFH.GetRecordSetFields(FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers Or FinancialHistory.FinancialHistoryRecordSetTypes.fhrtDetail) & ", tt.transaction_sign"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_booking_transactions ebt", vWhereFields, "ebt.batch_number, ebt.transaction_number, ebt.line_number", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRS.Fetch
        vFH = New FinancialHistory()
        vFH.InitFromRecordSet(mvEnv, vRS, FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers Or FinancialHistory.FinancialHistoryRecordSetTypes.fhrtDetail)
        vColl.Add(vFH.Key, vFH)
      End While
      vRS.CloseRecordSet()

      Return vColl

    End Function
  End Class
End Namespace
