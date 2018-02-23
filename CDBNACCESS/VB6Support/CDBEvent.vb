Namespace Access

  Public Enum EventDeleteAllowedStatuses
    edasOK
    edasBookings
    edasOrganiserTransactions
    edasBlockBookings
    edasActions
  End Enum

  Public Enum UserWaitingListMethods
    euwlManual
    euwlAutomatic
    euwlNone
  End Enum

  Public Enum CheckEventBookingResults
    cebrCanAddBooking
    cebrCanAddToWaitingList
    cebrCannotAddBooking
  End Enum

  Partial Public Class CDBEvent

    Public Enum EventNumberFields
      enfSessionNumber
      enfBookingNumber
      enfOptionNumber
    End Enum

    Public Enum WaitingListControlMethods
      ewlAlwaysManual
      ewlAlwaysAutomatic
      ewlManualIfAccess
    End Enum

    Public Enum UserWaitingListMethods
      euwlManual
      euwlAutomatic
      euwlNone
    End Enum

    Private mvSessions As CollectionList(Of EventSession)
    Private mvOrganiser As EventOrganiser
    Private mvBookingMessage As String
    Private mvBookingOptions As CollectionList(Of EventBookingOption)
    Private mvOptionSessions As List(Of EventOptionSession)
    Private mvResources As List(Of EventResource)
    Private mvSessionActivities As List(Of EventSessionActivity)
    Private mvVenueBookings As List(Of EventVenueBooking)
    Private mvSessionTests As List(Of SessionTest)
    Private mvEventOwners As CollectionList(Of EventOwner)
    Private mvPersonnel As List(Of EventPersonnel)

    Protected Overrides Sub ClearFields()
      mvOptionSessions = Nothing
      mvResources = Nothing
      mvSessionActivities = Nothing
      mvVenueBookings = Nothing
      mvSessionTests = Nothing
      mvEventOwners = Nothing
      mvPersonnel = Nothing
      mvSessions = Nothing
      mvOrganiser = Nothing
      mvBookingMessage = "2"
    End Sub

    Friend Sub CalculateSponsorshipIncome()
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet

      Dim vIncomeValue As Double = 0
      Dim vMaxValue As Double = 0
      If MaintenanceAttributes.ContainsKey(mvClassFields.Item(EventFields.SponsorshipIncome).Name) Then vMaxValue = DoubleValue(MaintenanceAttributes(mvClassFields.Item(EventFields.SponsorshipIncome).Name).MaximumValue)

      'Sponsorship Income
      mvClassFields(EventFields.SponsorshipIncome).IntegerValue = 0
      vWhereFields.Add("es.event_number", CDBField.FieldTypes.cftLong, EventNumber)
      vWhereFields.Add("fhd.source", CDBField.FieldTypes.cftLong, "es.source")
      vWhereFields.Add("p.product", CDBField.FieldTypes.cftLong, "fhd.product", CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("p.sponsorship_event", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoEqual)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM(fhd.amount) AS fhd_total FROM event_sources es, financial_history_details fhd, products p WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch Then
        vIncomeValue = vRecordSet.Fields(1).DoubleValue
      End If
      vRecordSet.CloseRecordSet()

      vWhereFields.Clear()
      vWhereFields.Add("efl.event_number", CDBField.FieldTypes.cftLong, EventNumber)
      vWhereFields.Add("fhd.batch_number", CDBField.FieldTypes.cftLong, "efl.batch_number")
      vWhereFields.Add("fhd.transaction_number", CDBField.FieldTypes.cftLong, "efl.transaction_number")
      vWhereFields.Add("fhd.line_number", CDBField.FieldTypes.cftLong, "efl.line_number")
      vWhereFields.Add("p.product", CDBField.FieldTypes.cftLong, "fhd.product", CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("p.sponsorship_event", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoEqual)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT SUM(fhd.amount) AS fhd_total FROM event_financial_links efl, financial_history_details fhd, products p WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch Then
        vIncomeValue += vRecordSet.Fields(1).DoubleValue
      End If
      vRecordSet.CloseRecordSet()

      vIncomeValue = FixTwoPlaces(vIncomeValue)
      If vMaxValue > 0 AndAlso vIncomeValue.CompareTo(vMaxValue) > 0 Then
        SetCalculateTotalsError(mvClassFields.Item(EventFields.SponsorshipIncome).Caption)
      Else
        mvClassFields(EventFields.SponsorshipIncome).Value = FixedFormat(vIncomeValue)
      End If

    End Sub

    Public ReadOnly Property BaseSession() As EventSession
      Get
        Dim vSession As EventSession

        If mvSessions Is Nothing Then
          mvSessions = New CollectionList(Of EventSession)(1)
          If mvExisting Then
            InitSessions()
          Else
            vSession = New EventSession
            vSession.Init(mvEnv)
            mvSessions.Add(vSession.SessionNumber.ToString, vSession)
          End If
        End If
        If mvSessions.Count() = 1 Then
          Return mvSessions.Item(1)
        Else
          Return mvSessions.Item(LowestSessionNumber.ToString)  'mvSessions.Item(BaseItemNumber.ToString)
        End If
      End Get
    End Property

    Public ReadOnly Property ItemsMultiplier() As Integer
      Get
        ItemsMultiplier = ITEMS_MULTIPLIER
      End Get
    End Property

    Public Sub InitSessions()

      mvSessions = New CollectionList(Of EventSession)(1)
      Dim vSession As New EventSession
      vSession.Init(mvEnv)
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vSession.GetRecordSetFields(EventSession.SessionRecordSetTypes.ssrtAll) & " FROM sessions s WHERE event_number = " & EventNumber & " ORDER BY start_date, start_time")
      While vRecordSet.Fetch()
        vSession = New EventSession
        vSession.InitFromRecordSet(mvEnv, vRecordSet, EventSession.SessionRecordSetTypes.ssrtAll)
        'Sessions are added in start date order
        mvSessions.Add(vSession.SessionNumber.ToString, vSession)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Function AllocateNextNumber(ByRef pEventNumberField As EventNumberFields) As Integer
      Dim vNumber As Integer
      Dim vWhereFields As New CDBFields


      Select Case pEventNumberField
        Case EventNumberFields.enfBookingNumber
          vNumber = mvEnv.GetControlNumber("EBN") '   EventFields.NextBookingNumber
        Case EventNumberFields.enfSessionNumber
          vNumber = mvEnv.GetControlNumber("ESN")  'EventFields.NextSessionNumber
        Case EventNumberFields.enfOptionNumber
          vNumber = mvEnv.GetControlNumber("EON") 'EventFields.NextOptionNumber
      End Select

      Return vNumber 'BaseItemNumber + vNumber
    End Function

    Public Sub SetVenueFromBooking(ByRef pVenueBooking As EventVenueBooking)
      With pVenueBooking
        mvClassFields.Item(EventFields.Venue).Value = pVenueBooking.Venue
        mvClassFields.Item(EventFields.VenueReference).Value = pVenueBooking.VenueReference
        mvClassFields.Item(EventFields.VenueConfirmedBy).Value = pVenueBooking.ConfirmedBy
        mvClassFields.Item(EventFields.VenueConfirmed).Value = pVenueBooking.ConfirmedOn
      End With
    End Sub

    Public ReadOnly Property UserWaitingListMethod() As UserWaitingListMethods
      Get
        If BaseSession.NumberOnWaitingList > 0 Then
          Select Case WaitingMethodFromCode(WaitingListControlMethod)
            Case WaitingListControlMethods.ewlAlwaysManual
              If mvEnv.User.HasAccessRights("CDEVWL") Then
                Return UserWaitingListMethods.euwlManual
              Else
                Return UserWaitingListMethods.euwlNone
              End If
            Case WaitingListControlMethods.ewlManualIfAccess
              If mvEnv.User.HasAccessRights("CDEVWL") Then
                Return UserWaitingListMethods.euwlManual
              Else
                Return UserWaitingListMethods.euwlAutomatic
              End If
            Case Else
              'ewlAlwaysAutomatic; default:
              Return UserWaitingListMethods.euwlAutomatic
          End Select
        Else
          Return UserWaitingListMethods.euwlNone
        End If
      End Get
    End Property

    Public ReadOnly Property WaitingMethodFromCode(ByVal pWaitingListMethodCode As String) As WaitingListControlMethods
      Get
        Select Case pWaitingListMethodCode
          Case "M"
            WaitingMethodFromCode = WaitingListControlMethods.ewlAlwaysManual
          Case "C"
            WaitingMethodFromCode = WaitingListControlMethods.ewlManualIfAccess
          Case Else
            '"A"/Not set - default:
            WaitingMethodFromCode = WaitingListControlMethods.ewlAlwaysAutomatic
        End Select
      End Get
    End Property

    Public Function ProcessWaitingList(ByRef pMsg As String) As Boolean
      Dim vTransferred As Boolean
      Dim vBookingNumber As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vWhere As New CDBFields
      Dim vWaitingInList As String

      ' Find any bookings on the waiting list ordered by booking_date
      '  AND eb.option_number=ebo.option_number
      vWaitingInList = "'" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaiting) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingCreditSale) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid) & "'"
      vWhere.Add("eb.event_number", EventNumber, CDBField.FieldWhereOperators.fwoEqual)
      vWhere.Add("booking_status", vWaitingInList, CDBField.FieldWhereOperators.fwoIn)
      vWhere.Add("ebo.option_number", CDBField.FieldTypes.cftLong, "eb.option_number", CDBField.FieldWhereOperators.fwoEqual)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT booking_number, quantity FROM event_bookings eb, event_booking_options ebo WHERE " & mvEnv.Connection.WhereClause(vWhere) & " ORDER BY booking_date, booking_number")
      Do While vRecordSet.Fetch()
        'Re-initialise Event each time after the first to ensure we have correct Event Capacity.
        Init(EventNumber)
        vBookingNumber = vRecordSet.Fields(1).LongValue
        If TransferFromWaiting(vBookingNumber) Then
          If Len(pMsg) > 0 Then pMsg = pMsg & vbCrLf
          pMsg = pMsg & String.Format(ProjectText.String25607, vBookingNumber.ToString, vRecordSet.Fields(2).LongValue.ToString) 'Booking Number %s for %s places has been transferred from waiting list
          vTransferred = True
        End If
      Loop
      vRecordSet.CloseRecordSet()
      Return vTransferred
      ''*** BR10976 TEMP BOOKING COUNT CHECK; Pls report to Tracey if this occurs ***
      'If mvEnv.ClientCode = "CARE" Then
      '  System.Diagnostics.Debug.Assert(CheckEventCounts = True, "")
      'End If
    End Function

    Public Function TransferFromWaiting(ByVal pBookingNumber As Integer, Optional ByRef pInvoiceNumber As Integer = 0) As Boolean
      Dim vTransaction As Boolean
      Dim vCanBook As Boolean
      Dim vEB As New EventBooking
      Dim vBOpt As New EventBookingOption
      Dim vRS As CDBRecordSet
      Dim vUpdate As New CDBFields
      Dim vWhere As New CDBFields
      Dim vRows As Integer
      Dim vNewStatusCode As String = ""
      Dim vBatch As New Batch(mvEnv)
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vRate As New ProductRate(mvEnv)
      Dim vVatRate As New VatRate(mvEnv)
      Dim vInvoice As New Invoice
      Dim vOrigCreditSale As New CreditSale(mvEnv)
      Dim vCreditSale As New CreditSale(mvEnv)
      Dim vCreditCustomer As New CreditCustomer
      Dim vOutstanding As Double
      Dim vOnOrder As Double
      Dim vBankAccount As New BankAccount(mvEnv)
      Dim vCSTerms As New CreditSalesTerms
      Dim vDate As Date
      Dim vStatus As String = ""
      Dim vOrigInvoice As New Invoice
      Dim vOrigCardSale As New CardSale(mvEnv)
      Dim vCardSale As New CardSale(mvEnv)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vBatchOpened As Boolean = False

      vEB.Init(mvEnv, (mvClassFields(EventFields.EventNumber).LongValue), pBookingNumber)
      vBOpt.Init(mvEnv, (vEB.OptionNumber))
      vCanBook = True

      If vBOpt.DeductFromEvent Then
        If BaseSession.NumberOfAttendees + vEB.Quantity > BaseSession.MaximumAttendees Then vCanBook = False
      End If

      If vCanBook Then
        vRS = mvEnv.Connection.GetRecordSet("SELECT number_of_attendees, maximum_attendees FROM session_bookings sb, sessions s WHERE sb.booking_number=" & pBookingNumber & " AND sb.session_number=s.session_number AND session_type <> '0'")
        Do While vRS.Fetch() And vCanBook
          If vRS.Fields("number_of_attendees").LongValue + vEB.Quantity > vRS.Fields("maximum_attendees").LongValue Then vCanBook = False
        Loop
        vRS.CloseRecordSet()
      End If

      If vCanBook Then
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vTransaction = True
        End If

        ' Change quantities from the session count
        vUpdate.Add("number_on_waiting_list", CDBField.FieldTypes.cftInteger, "number_on_waiting_list - " & vEB.Quantity)
        vUpdate.Add("number_of_attendees", CDBField.FieldTypes.cftInteger, "number_of_attendees + " & vEB.Quantity)
        vWhere.Add("session_number", CDBField.FieldTypes.cftLong, "SELECT session_number FROM session_bookings WHERE booking_number=" & pBookingNumber, CDBField.FieldWhereOperators.fwoIn)
        vWhere.Add("session_type", CDBField.FieldTypes.cftCharacter, "0", CDBField.FieldWhereOperators.fwoNotEqual)
        vRows = mvEnv.Connection.UpdateRecords("sessions", vUpdate, vWhere, False)
        ' Change quantities from the event if required
        If vBOpt.DeductFromEvent Then
          vWhere.Clear()
          vWhere.Add("session_number", CDBField.FieldTypes.cftLong, LowestSessionNumber)
          vRows = mvEnv.Connection.UpdateRecords("sessions", vUpdate, vWhere)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then
            vRate = New ProductRate(mvEnv)
            vRate.Init(vEB.BookingOption.ProductCode, (vEB.RateCode))
            If vRate.FixedPrice Then
              NumberOfBookings = NumberOfBookings + 1
            Else
              NumberOfBookings = NumberOfBookings + vEB.Quantity
            End If
            Save()
          End If
        End If
        If vRows > 0 Then
          Select Case vEB.BookingStatus
            Case EventBooking.EventBookingStatuses.ebsWaiting
              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedTransfer)
            Case EventBooking.EventBookingStatuses.ebsWaitingPaid
              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer)
            Case EventBooking.EventBookingStatuses.ebsWaitingCreditSale
              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer)
            Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced
              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
          End Select
          vUpdate.Clear()
          vUpdate.AddAmendedOnBy(mvEnv.User.UserID)
          vUpdate.Add("booking_status", CDBField.FieldTypes.cftCharacter, vNewStatusCode)
          vWhere.Clear()
          vWhere.Add("booking_number", CDBField.FieldTypes.cftLong, pBookingNumber)
          ' Update the booking
          vRows = mvEnv.Connection.UpdateRecords("event_bookings", vUpdate, vWhere)

          If ChargeForWaiting = False Then
            If vEB.BatchTransactionAnalysis.Existing Then 'i.e. not FOC
              With vEB.BatchTransactionAnalysis
                vRate = New ProductRate(mvEnv)
                vRate.Init(.ProductCode, .RateCode)
              End With
              If Not vRate.PriceIsZero Then
                vBatch.InitOpenBatch(vEB.Batch)
                vBatch.Save()
                vBatchOpened = True
                With vBT
                  .Init()
                  .InitForUpdate(vBatch.BatchNumber, vBatch.AllocateTransactionNumber, False)
                  .NextLineNumber = 1
                  .CloneForFA(vEB.BatchTransaction)
                  .TransactionDate = TodaysDate()
                End With

                With vBTA
                  .InitFromTransaction(vBT)
                  .CloneFromBTA(vEB.BatchTransactionAnalysis)
                  .Amount = vRate.Price(vEB.ContactNumber) * .Quantity
                  .CurrencyAmount = .Amount
                  vRS = mvEnv.Connection.GetRecordSet("SELECT * FROM vat_rates WHERE vat_rate = '" & .VatRate & "'")
                  If vRS.Fetch() Then
                    vVatRate.InitFromRecordSet(vRS)
                    .SetVATAmounts(vVatRate, "", (vBT.TransactionDate))
                  End If
                  vRS.CloseRecordSet()
                  .Save()
                End With

                vBT.Save()

                vWhereFields.Add("batch_number", vBatch.BatchNumber, CDBField.FieldWhereOperators.fwoEqual)
                vUpdateFields.Add("number_of_transactions", CDBField.FieldTypes.cftLong, "number_of_transactions + 1")
                vUpdateFields.Add("transaction_total", CDBField.FieldTypes.cftLong, "transaction_total + " & vBT.Amount)
                vUpdateFields.Add("contents_amended_by", mvEnv.User.UserID)
                vUpdateFields.Add("contents_amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
                mvEnv.Connection.UpdateRecords("batches", vUpdateFields, vWhereFields)

                If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataLinkToCommunication) Then
                  vWhere.Clear()
                  vUpdate.Clear()
                  vWhere.Add("batch_number", vEB.BatchNumber, CDBField.FieldWhereOperators.fwoEqual)
                  vWhere.Add("transaction_number", vEB.TransactionNumber, CDBField.FieldWhereOperators.fwoEqual)
                  vUpdate.Add("batch_number", vBT.BatchNumber, CDBField.FieldWhereOperators.fwoEqual)
                  vUpdate.Add("transaction_number", vBT.TransactionNumber, CDBField.FieldWhereOperators.fwoEqual)
                  mvEnv.Connection.UpdateRecords("communications_log_trans", vUpdate, vWhere, False)
                End If

                Select Case vBatch.BatchType
                  Case Batch.BatchTypes.CreditSales
                    vOrigInvoice.Init(mvEnv, vEB.BatchNumber, vEB.TransactionNumber)
                    vOrigCreditSale.Init(vEB.BatchNumber, vEB.TransactionNumber)
                    vBankAccount.Init(vBatch.BankAccount)
                    vCreditCustomer.Init(mvEnv, (vBT.ContactNumber), vBankAccount.Company, vOrigCreditSale.SalesLedgerAccount)
                    vOutstanding = vCreditCustomer.Outstanding
                    vOnOrder = vCreditCustomer.OnOrder
                    If mvEnv.GetConfigOption("fp_use_sales_ledger", True) = True Then
                      vCSTerms.Init(mvEnv, (vBT.ContactNumber), (vBankAccount.Company), (vOrigCreditSale.SalesLedgerAccount))
                      'Create Credit Sale
                      vCreditSale.Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
                      vCreditSale.Update(vCreditCustomer.ContactNumber, vCreditCustomer.AddressNumber, vOrigCreditSale.AddressTo, vCreditCustomer.SalesLedgerAccount, False)
                      vCreditSale.Save()
                      'Create Invoice
                      vInvoice.Init(mvEnv)
                      If vInvoice.CalcInvPayDue(vCSTerms.TermsFrom, vCSTerms.TermsPeriod, vCSTerms.TermsNumber, vBTA.BatchNumber, vBTA.TransactionNumber, CDate(vBT.TransactionDate), vDate) = True Then
                        If vInvoice.GetInvPayStatus(Invoice.InvPayStatuses.InvNotPaid, vStatus) = True Then
                          vInvoice.Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
                          vInvoice.Update(0, vCreditCustomer.ContactNumber, vCreditCustomer.AddressNumber, vCreditCustomer.Company, vCreditCustomer.SalesLedgerAccount, 0, -1, vBT.TransactionDate, CStr(vDate), vStatus, "I")
                          vInvoice.Save()
                        End If
                      End If
                      'Invoice will be created later - increment Outstanding
                      vOutstanding = vOutstanding + vBT.Amount
                      With vCreditCustomer
                        .Update(.CreditCategory, .CreditLimit, .CustomerType, .StopCode, .TermsNumber, .TermsPeriod, .TermsFrom, FixedFormat(vOutstanding), .AddressNumber, FixedFormat(vOnOrder))
                        .Save()
                      End With
                    End If 'Use Sales Ledger
                  Case Batch.BatchTypes.CreditCard
                    vOrigCardSale.Init(vEB.BatchNumber, vEB.TransactionNumber)
                    With vCardSale
                      .Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
                      .CloneFromCardSale(vOrigCardSale)
                      .Save()
                    End With
                End Select
                'Reset Event Booking to point to new CS Transaction
                vEB = New EventBooking
                vEB.SetTransactionInfo(mvEnv, pBookingNumber, vBTA.BatchNumber, vBTA.TransactionNumber, vBTA.LineNumber, vBTA.SalesContactNumber, EventNumber)
                vEB.Save()
              End If 'Current Price > 0
            End If 'BTA.existing
          End If 'Charge for Waiting
        End If
        If vTransaction Then mvEnv.Connection.CommitTransaction()

        If vBatchOpened AndAlso vBatch.BatchType = Batch.BatchTypes.CreditCard Then
          If vCardSale.TemplateNumber.Length > 0 Then
            'Used by TransferWaitingListBooking, CancelEventBooking and UpdateEventBooking web services
            Dim vCCA As New CreditCardAuthorisation
            vCCA.InitFromTransaction(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
            vCCA.AuthoriseTransaction(vCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatNormal, vBT.Amount, vBT.AddressNumber)
            vCardSale.Save()
          End If
        End If
      End If
      Return vRows > 0
    End Function

    Public Function CheckEventCounts() As Boolean
      '*** BR10976 TEMP BOOKING COUNT CHECK ***
      Dim vValid As Boolean
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet

      If MultiSession = True Then
        'For now only check single session Events
        vValid = True
      Else
        InitSessions()
        vWhereFields.Add("event_number", EventNumber, CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("booking_status", CDBField.FieldTypes.cftCharacter, CurrentAttendeeBookingStatuses.InList, CDBField.FieldWhereOperators.fwoIn)
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT sum(quantity) FROM event_bookings eb WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch Then
          If vRecordSet.Fields(1).LongValue = DirectCast(Sessions(LowestSessionNumber.ToString), EventSession).NumberOfAttendees Then
            vValid = True
          End If
        Else
          vValid = True
        End If
        vRecordSet.CloseRecordSet()
      End If
      CheckEventCounts = vValid
    End Function

    Public ReadOnly Property CurrentAttendeeBookingStatuses() As CDBParameters
      Get
        Dim vParams As New CDBParameters

        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBooked))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaid))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedTransfer))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSale))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer))
        vParams.Add(EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsExternal))
        Return vParams
      End Get
    End Property

    Public ReadOnly Property Sessions() As CollectionList(Of EventSession)
      Get
        If mvSessions Is Nothing Then InitSessions()
        Return mvSessions
      End Get
    End Property

    Public ReadOnly Property LastBookingMessage() As String
      Get
        Return mvBookingMessage
      End Get
    End Property

    Public ReadOnly Property OptionSessions() As List(Of EventOptionSession)
      Get
        If mvOptionSessions Is Nothing Then InitOptionSessions()
        Return mvOptionSessions
      End Get
    End Property

    Public Function AddEventBooking(ByVal pContact As Contact, ByVal pAddressNumber As Integer, ByVal pQuantity As Integer, ByVal pOptionNumber As Integer, ByVal pStatus As EventBooking.EventBookingStatuses, ByRef pRateCode As String, ByRef pSessionList As String, Optional ByVal pNotes As String = "", Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0, Optional ByVal pConvertInterestedBooking As EventBooking = Nothing, Optional ByVal pPledgedAmount As Double = 0, Optional ByVal pCreateCalByOptSessions As Boolean = False, Optional ByVal pAdultQuantity As String = "", Optional ByVal pChildQuantity As String = "", Optional ByVal pBookingDate As String = "", Optional ByVal pStartTime As String = "", Optional ByVal pEndTime As String = "", Optional ByVal pSalesContactNumber As Integer = 0) As EventBooking
      Dim vEventBooking As New EventBooking
      Dim vDelegateNumber As Integer
      Dim vTransaction As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vCount As Integer
      Dim vPosition As String = ""
      Dim vOrgName As String = ""
      Dim vExistingSessionList As String = ""
      Dim vNewEventBooking As EventBooking = Nothing
      Dim vPrevNumberOfBooking As Integer = 0
      Dim vBookingThreshold As Integer = 0

      If Len(BookingsClose) > 0 Then
        If IsDate(BookingsClose) And CDate(BookingsClose) < Today Then RaiseError(DataAccessErrors.daeCannotBookEvent, EventNumber.ToString)
      Else
        If Booking = False Then RaiseError(DataAccessErrors.daeCannotBookEvent, CStr(EventNumber))
      End If
      vDelegateNumber = mvEnv.GetControlNumber("ED") 'This needs to be outside the Transaction
      If pContact.ContactType = Contact.ContactTypes.ctcJoint Then
        RaiseError(DataAccessErrors.daeCannotBookJointToEvent, pContact.ContactNumber & " : " & pContact.LabelName, CStr(EventNumber))
      End If
      Try
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vTransaction = True
        End If
        vPrevNumberOfBooking = BaseSession.NumberOfAttendees
        vBookingThreshold = IntegerValue(mvEnv.GetConfig("ev_percentage_booked_email", "0"))
        If UpdateEventCounts(pOptionNumber, pQuantity, pSessionList, pStatus) Then
          If Not pConvertInterestedBooking Is Nothing Then
            If pConvertInterestedBooking.Existing Then
              With pConvertInterestedBooking
                'Decrement "Interested" counts
                If .Sessions.Count() > 0 Then
                  For vCount = 1 To .Sessions.Count()
                    If Len(vExistingSessionList) > 0 Then vExistingSessionList = vExistingSessionList & ","
                    vExistingSessionList = vExistingSessionList & DirectCast(.Sessions.Item(vCount), EventSession).SessionNumber
                  Next
                End If
                'Create a caledar entry
                vEventBooking.Init(mvEnv, pConvertInterestedBooking.EventNumber, pConvertInterestedBooking.BookingNumber)
                vEventBooking.AddDelegateCalendar(pContact.ContactNumber, pOptionNumber, pCreateCalByOptSessions, pStatus, True)

                UpdateEventCounts(.OptionNumber, .Quantity * -1, vExistingSessionList, .BookingStatus)
                'Update Booking to actual requirements
                .ModifyBooking((pContact.ContactNumber), pAddressNumber, pQuantity, pOptionNumber, pStatus, pRateCode, pNotes, pBatchNumber, pTransNumber, pLineNumber, pAdultQuantity, pChildQuantity, pBookingDate)
                .Save()
                UpdateNumberOfBookings(pConvertInterestedBooking, pOptionNumber, pQuantity, pRateCode, True)
                mvEnv.AddJournalRecord(JournalTypes.jnlEvent, JournalOperations.jnlUpdate, pContact.ContactNumber, pAddressNumber, EventNumber, .BookingNumber, 0, 0, 0)
                vNewEventBooking = pConvertInterestedBooking
              End With
            End If
          Else
            vEventBooking.InitNewBooking(mvEnv, Me)
            vEventBooking.ModifyBooking((pContact.ContactNumber), pAddressNumber, pQuantity, pOptionNumber, pStatus, pRateCode, pNotes, pBatchNumber, pTransNumber, pLineNumber, pAdultQuantity, pChildQuantity, pBookingDate, pStartTime, pEndTime, pSalesContactNumber)
            AddSessionBookings(pOptionNumber, vEventBooking.BookingNumber, pQuantity, pSessionList)

            'Add the default delegate - namely the booker of the event
            ' BR11719 - Need to pass the status as well
            ' BR11718 - Default Position and Organisation only if the config is set and there is only one current position
            vCount = 0
            If mvEnv.GetConfigOption("default_delegate_position") = True Then
              vRecordSet = mvEnv.Connection.GetRecordSet("SELECT cp.position,o.name FROM contact_positions cp, organisations o WHERE cp.contact_number = " & pContact.ContactNumber & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND cp.organisation_number = o.organisation_number")
              While vRecordSet.Fetch
                vPosition = vRecordSet.Fields(1).Value
                vOrgName = vRecordSet.Fields(2).Value
                vCount = vCount + 1
              End While
              vRecordSet.CloseRecordSet()
            End If
            If vCount <> 1 Then
              vPosition = pContact.Position
              vOrgName = pContact.OrganisationName
            End If
            vEventBooking.AddDelegate(pContact.ContactNumber, pAddressNumber, vPosition, vOrgName, Source, pOptionNumber, vDelegateNumber, "", pPledgedAmount, pCreateCalByOptSessions, pStatus)
            vEventBooking.Save(mvEnv.User.UserID, True)
            UpdateNumberOfBookings(vEventBooking, pOptionNumber, pQuantity, pRateCode)
            mvEnv.AddJournalRecord(JournalTypes.jnlEvent, JournalOperations.jnlInsert, pContact.ContactNumber, pAddressNumber, EventNumber, vEventBooking.BookingNumber, 0, 0, 0)
            vNewEventBooking = vEventBooking
          End If
        Else
          mvEnv.Connection.RollbackTransaction()
        End If
        If vTransaction Then mvEnv.Connection.CommitTransaction()
        If vBookingThreshold > 0 AndAlso Me.AdminEmailAddress.Length > 0 Then
          BaseSession.Init(mvEnv, BaseSession.SessionNumber)        'Reload the session to get the actual number of attendees
          If (vPrevNumberOfBooking * 100 / BaseSession.MaximumAttendees) < vBookingThreshold AndAlso (BaseSession.NumberOfAttendees * 100 / BaseSession.MaximumAttendees) >= vBookingThreshold Then
            Dim vEmailJob As New EmailJob(mvEnv)
            vEmailJob.Init()
            vEmailJob.SendEmail("Event has reached its Capacity ", "Event " & EventNumber & " has reached " & vBookingThreshold & " % Capacity", Me.AdminEmailAddress, EventNumber.ToString)
          End If
        End If
        '*** BR10976 TEMP BOOKING COUNT CHECK; Pls report to Tracey if this occurs ***
        If mvEnv.ClientCode = "CARE" Then
          System.Diagnostics.Debug.Assert(CheckEventCounts() = True, "")
        End If
      Catch vEx As Exception
        mvEnv.Connection.RollbackTransaction()
        Throw vEx
      End Try
      Return vNewEventBooking
    End Function

    Private Function UpdateEventCounts(ByVal pOptionNumber As Integer, ByVal pCount As Integer, ByVal pSessionList As String, ByVal pStatus As EventBooking.EventBookingStatuses) As Boolean
      Dim vSessionSQL As String
      Dim vAddSQL As String
      Dim vTestSQL As String
      Dim vDesc As String
      Dim vRecordSet As CDBRecordSet
      Dim vUpdated As Boolean
      Dim vBookingOption As EventBookingOption

      If pSessionList.Length > 0 Then
        vSessionSQL = pSessionList
      Else
        vSessionSQL = "SELECT session_number FROM option_sessions WHERE option_sessions.option_number = " & pOptionNumber
      End If
      vUpdated = True
      If pStatus = EventBooking.EventBookingStatuses.ebsWaiting Or pStatus = EventBooking.EventBookingStatuses.ebsWaitingPaid Or pStatus = EventBooking.EventBookingStatuses.ebsWaitingCreditSale Or pStatus = EventBooking.EventBookingStatuses.ebsWaitingInvoiced Then
        vAddSQL = " number_on_waiting_list = number_on_waiting_list + "
        vTestSQL = " number_on_waiting_list > maximum_on_waiting_list "
        vDesc = ProjectText.String25622 'on waiting list
      Else
        If pStatus = EventBooking.EventBookingStatuses.ebsInterested Or pStatus = EventBooking.EventBookingStatuses.ebsAwaitingAcceptance Then
          vAddSQL = " number_interested = number_interested + "
          vTestSQL = ""
          vDesc = ""
        Else
          vAddSQL = " number_of_attendees = number_of_attendees + "
          vTestSQL = " number_of_attendees > maximum_attendees "
          vDesc = ProjectText.String25623 'attendees
        End If
      End If
      mvEnv.Connection.ExecuteSQL("UPDATE sessions SET" & vAddSQL & pCount & " WHERE session_number IN (" & vSessionSQL & ") AND session_type <> '0'")
      If vTestSQL.Length > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT session_desc FROM sessions WHERE session_number IN (" & vSessionSQL & ") AND" & vTestSQL & "AND session_type <> '0'")
        If vRecordSet.Fetch() Then
          vUpdated = False
          mvBookingMessage = String.Format(ProjectText.String25610, vDesc, vRecordSet.Fields(1).Value) 'Quantity would exceed maximum %s for session: %s
        End If
        vRecordSet.CloseRecordSet()
      End If
      If vUpdated Then
        For Each vBookingOption In BookingOptions
          If vBookingOption.OptionNumber = pOptionNumber Then
            If vBookingOption.DeductFromEvent Then
              'mvEnv.Connection.ExecuteSQL("UPDATE sessions SET" & vAddSQL & pCount & " WHERE session_number = " & BaseItemNumber)
              mvEnv.Connection.ExecuteSQL("UPDATE sessions SET" & vAddSQL & pCount & " WHERE session_number = " & LowestSessionNumber)
              If vTestSQL.Length > 0 Then
                'vRecordSet = mvEnv.Connection.GetRecordSet("SELECT session_desc FROM sessions WHERE session_number = " & BaseItemNumber & " AND" & vTestSQL)
                vRecordSet = mvEnv.Connection.GetRecordSet("SELECT session_desc FROM sessions WHERE session_number = " & LowestSessionNumber & " AND" & vTestSQL)
                If vRecordSet.Fetch() Then
                  vUpdated = False
                  mvBookingMessage = String.Format(ProjectText.String25611, vDesc) 'Quantity would exceed maximum %s for event
                End If
                vRecordSet.CloseRecordSet()
              End If
            End If
            Exit For
          End If
        Next vBookingOption
      End If
      UpdateEventCounts = vUpdated
    End Function

    Public ReadOnly Property BookingOptions() As CollectionList(Of EventBookingOption)
      Get
        If mvBookingOptions Is Nothing Then InitBookingOptions()
        BookingOptions = mvBookingOptions
      End Get
    End Property

    Public Sub InitBookingOptions(Optional ByRef pCandidateNumberingOrder As Boolean = False)
      Dim vBookingOption As New EventBookingOption
      Dim vRecordSet As CDBRecordSet
      Dim vOrderBy As String

      mvBookingOptions = New CollectionList(Of EventBookingOption)(1)
      vBookingOption.Init(mvEnv)
      If pCandidateNumberingOrder Then
        vOrderBy = " ORDER BY number_of_sessions DESC, option_number"
      Else
        vOrderBy = " ORDER BY option_number"
      End If
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vBookingOption.GetRecordSetFields(EventBookingOption.EventBookingOptionRecordSetTypes.ebortAll) & " FROM event_booking_options WHERE event_number = " & EventNumber & vOrderBy)
      While vRecordSet.Fetch()
        vBookingOption = New EventBookingOption
        vBookingOption.InitFromRecordSet(mvEnv, vRecordSet, EventBookingOption.EventBookingOptionRecordSetTypes.ebortAll)
        mvBookingOptions.Add(vBookingOption.OptionNumber.ToString, vBookingOption)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Function AddSessionBookings(ByVal pOptionNumber As Integer, ByVal pBookingNumber As Integer, ByVal pQty As Integer, ByVal pSessionList As String) As Integer
      Dim vSQL As String

      vSQL = "INSERT INTO session_bookings (event_number, booking_number, session_number, quantity, amended_by, amended_on)"
      vSQL = vSQL & " SELECT " & EventNumber & "," & pBookingNumber & ", session_number, " & pQty & ",'" & mvEnv.User.UserID & "',"
      vSQL = vSQL & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, TodaysDate) & " FROM option_sessions WHERE option_number = " & pOptionNumber
      If pSessionList.Length > 0 Then vSQL = vSQL & " AND session_number IN (" & pSessionList & ")"
      AddSessionBookings = mvEnv.Connection.ExecuteSQL(vSQL)
    End Function

    Public Sub AddOwner(ByRef pOwner As String)
      Dim vOwner As New EventOwner
      With vOwner
        .Init(mvEnv)
        .EventNumber = EventNumber
        .Department = pOwner
        .Save()
      End With
    End Sub

    Public Sub AddStandardResources()
      Dim vAttrList As String
      Dim vSelectList As String

      If Not MultiSession Then
        vAttrList = "session_number, product, rate, copy_to, despatch_to, issue_basis, amended_on, amended_by, allocated"
        vSelectList = "session_number,sr.product, sr.rate, copy_to, despatch_to, issue_basis, " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, TodaysDate) & ", '" & mvEnv.User.Logname & "'"
        vSelectList = vSelectList & ", 'N'"
        mvEnv.Connection.ExecuteSQL("INSERT INTO event_resources (" & vAttrList & ") SELECT " & vSelectList & " FROM sessions s, standard_resources sr WHERE s.event_number = " & EventNumber & " AND sr.subject = s.subject AND sr.skill_level = s.skill_level")
      End If
    End Sub

    Public Function BaseSessionPersonnel(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer) As EventPersonnel
      Dim vEventPersonnel As New EventPersonnel(mvEnv)
      Dim vPersonnelParams As New CDBParameters
      Dim vSession As EventSession = BaseSession
      vEventPersonnel.Init(vSession.SessionNumber, pContactNumber)
      If vEventPersonnel.Existing = False Then
        vPersonnelParams.Add("StartDate", CDBField.FieldTypes.cftCharacter, vSession.StartDate)
        vPersonnelParams.Add("EndDate", CDBField.FieldTypes.cftCharacter, vSession.EndDate)
        vPersonnelParams.Add("StartTime", CDBField.FieldTypes.cftCharacter, vSession.StartTime)
        vPersonnelParams.Add("EndTime", CDBField.FieldTypes.cftCharacter, vSession.EndTime)
        vEventPersonnel.Create(vSession.SessionNumber, pContactNumber, pAddressNumber, vPersonnelParams)
        vEventPersonnel.AppointmentDescription = PersonnelAppointmentDescription(vEventPersonnel)
        vEventPersonnel.CheckCalendarConflict()
        vEventPersonnel.Save(mvEnv.User.UserID)
      End If
      Return vEventPersonnel
    End Function

    Public Function PersonnelAppointmentDescription(ByRef pEventPersonnel As EventPersonnel) As String
      If pEventPersonnel.SessionNumber = LowestSessionNumber Then
        Return EventDesc
      Else
        Return EventDesc & " : " & DirectCast(Sessions(pEventPersonnel.SessionNumber.ToString), EventSession).SessionDesc
      End If
    End Function

    Public Function ResourceAppointmentDescription(ByRef pEventResource As EventResource) As String
      If pEventResource.SessionNumber = LowestSessionNumber Then
        Return EventDesc
      Else
        Return EventDesc & " : " & DirectCast(Sessions(pEventResource.SessionNumber.ToString), EventSession).SessionDesc
      End If
    End Function

    Public ReadOnly Property Organiser() As EventOrganiser
      Get
        If mvOrganiser Is Nothing Then InitOrganiser()
        Return mvOrganiser
      End Get
    End Property

    Private Sub InitOrganiser()
      mvOrganiser = New EventOrganiser
      mvOrganiser.Init(mvEnv, EventNumber)
    End Sub

    Public Sub DeleteOrganiser()
      If Organiser.Existing Then
        mvOrganiser.Delete()
        mvOrganiser = Nothing
      End If
    End Sub

    Public Function DeleteAllowedStatus() As EventDeleteAllowedStatuses
      Dim vWhereFields As CDBFields

      With mvEnv.Connection
        vWhereFields = New CDBFields
        vWhereFields.Add("event_number", EventNumber)
        If .GetCount("event_bookings", vWhereFields) > 0 Then
          Return EventDeleteAllowedStatuses.edasBookings              'Bookings have been made for this Event; Event cannot be deleted
        ElseIf .GetCount("organiser_transactions", vWhereFields) > 0 Then
          Return EventDeleteAllowedStatuses.edasOrganiserTransactions 'Organiser Transactions exist for this Event; Event cannot be deleted
        ElseIf .GetCount("event_room_links", vWhereFields) > 0 Then
          Return EventDeleteAllowedStatuses.edasBlockBookings         'Room Block Bookings exist for this Event; Event cannot be deleted
        Else
          If MasterAction > 0 Then
            vWhereFields = New CDBFields
            vWhereFields.Add("master_action", CDBField.FieldTypes.cftLong, MasterAction)
            If .GetCount("actions", vWhereFields) > 0 Then Return EventDeleteAllowedStatuses.edasActions
          End If
          Return EventDeleteAllowedStatuses.edasOK
        End If
      End With
    End Function

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vWhereFields As CDBFields
      With mvEnv.Connection
        .StartTransaction()
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, ClassFields.AmendmentHistoryCreation.ahcDefault))

        'Get all the session for the current event
        Dim vDataTable As CDBDataTable = GetEventSession()
        If vDataTable.Rows.Count > 0 Then
          Dim vSessionNumber As String = vDataTable.RowsAsCommaSeperated(vDataTable, "session_number")
          vWhereFields = New CDBFields
          'vWhereFields.Add("session_number", BaseItemNumber, CDBField.FieldWhereOperators.fwoBetweenFrom)
          'vWhereFields.Add("session_number2", MaxItemNumber, CDBField.FieldWhereOperators.fwoBetweenTo)
          vWhereFields.Add("session_number", vSessionNumber, CDBField.FieldWhereOperators.fwoIn)

          .DeleteRecords("event_personnel", vWhereFields, False)
          .DeleteRecords("event_resources", vWhereFields, False)
          .DeleteRecords("option_sessions", vWhereFields, False)
          .DeleteRecords("session_tests", vWhereFields, False)
          .DeleteRecords("session_test_results", vWhereFields, False)
          .DeleteRecords("session_candidate_numbers", vWhereFields, False)
        End If


        'Delete Contact Appointments created while creating bookings for event session
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
          Dim vSessionNumber As String = vDataTable.RowsAsCommaSeperated(vDataTable, "session_number")
          DeleteContactAppointment(vSessionNumber)
        End If

        ''Delete contact appointments which are created while creating event personnel 
        'Dim vPersonnelDataTable As CDBDataTable = GetEventPersonnel()
        'If vPersonnelDataTable IsNot Nothing AndAlso vPersonnelDataTable.Rows.Count > 0 Then
        '  Dim vSessionNumber As String = vPersonnelDataTable.ConvertToDataTable().Rows.AsCommaSeperated("session_number")
        '  DeleteContactAppointment(vSessionNumber)
        'End If

        vWhereFields = New CDBFields
        vWhereFields.Add("event_number", EventNumber)
        .DeleteRecords("contact_room_bookings", vWhereFields, False)
        .DeleteRecords("delegates", vWhereFields, False)
        .DeleteRecords("event_booking_options", vWhereFields, False)
        .DeleteRecords("event_bookings", vWhereFields, False)
        .DeleteRecords("event_organisers", vWhereFields, False)
        .DeleteRecords("event_room_links", vWhereFields, False)
        .DeleteRecords("event_submissions", vWhereFields, False)
        .DeleteRecords("loan_items", vWhereFields, False)
        .DeleteRecords("organiser_transactions", vWhereFields, False)
        .DeleteRecords("session_activities", vWhereFields, False)
        .DeleteRecords("session_bookings", vWhereFields, False)
        .DeleteRecords("sessions", vWhereFields, False)
        .DeleteRecords("event_venue_bookings", vWhereFields, False)
        .DeleteRecords("event_topics", vWhereFields, False)
        .DeleteRecords("event_contacts", vWhereFields, False)
        .DeleteRecords("event_owners", vWhereFields, False)
        .DeleteRecords("external_resources", vWhereFields, False)
        .DeleteRecords("event_sources", vWhereFields, False)
        vWhereFields = New CDBFields
        vWhereFields.Add("unique_id", EventNumber)
        vWhereFields.Add("record_type", "E")
        .DeleteRecords("sundry_costs", vWhereFields, False)
        .CommitTransaction()
        InitClassFields()
        SetDefaults()
      End With
    End Sub

    Public Sub UpdateNumberOfBookings(ByVal pEventBooking As EventBooking, ByVal pOptionNumber As Integer, ByVal pQuantity As Integer, ByVal pRateCode As String, Optional ByVal pConvertInterestedBooking As Boolean = False, Optional ByVal pUpdateSessionAttendees As Boolean = False)
      Dim vRate As ProductRate
      Dim vBookingOption As EventBookingOption

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then
        If CurrentAttendeeBookingStatuses.Exists(EventBooking.GetBookingStatusCode(pEventBooking.BookingStatus)) _
        OrElse ((pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsInterested OrElse pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsAwaitingAcceptance) AndAlso pConvertInterestedBooking = True) _
        OrElse pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsAmended Then 'BR14897 : added amended status condition to reset no. of booking to actual value after being set to 0 in Eventbooking.vb --> CancelOrDelete function
          For Each vBookingOption In BookingOptions
            If vBookingOption.OptionNumber = pOptionNumber Then
              If vBookingOption.DeductFromEvent Then
                vRate = New ProductRate(mvEnv)
                vRate.Init(vBookingOption.ProductCode, If(pRateCode.Length > 0, pRateCode, vBookingOption.RateCode))
                If vRate.FixedPrice Then
                  NumberOfBookings = NumberOfBookings + 1
                Else
                  NumberOfBookings = NumberOfBookings + pQuantity
                End If
                Save()
              End If
            End If
          Next vBookingOption
          'BR15803 
          'update numberofattandees in session during amend event booking
          If pUpdateSessionAttendees = True Then
            Dim vSessions As Collection = pEventBooking.Sessions
            If vSessions Is Nothing OrElse vSessions.Count = 0 Then
              vSessions = New Collection
              vSessions.Add(BaseSession)
            End If

            Dim vSessionList As String = ""
            For Each vSession As EventSession In vSessions
              If vSessionList.Length > 0 Then vSessionList &= ","
              vSessionList &= vSession.SessionNumber.ToString
            Next
            UpdateEventCounts(pEventBooking.OptionNumber, pQuantity, vSessionList, EventBooking.EventBookingStatuses.ebsAmended)
          End If
        End If
      End If
    End Sub

    Public Function ChangeSessionBookings(ByVal pEventBooking As EventBooking, ByVal pNewOptionNo As Integer, ByVal pNewQty As Integer, ByVal pNewStatus As EventBooking.EventBookingStatuses, ByVal pSessionList As String) As Boolean
      Dim vSubSQL As String
      Dim vWhereFields As New CDBFields
      Dim vBookingOption As EventBookingOption
      Dim vRecordSet As CDBRecordSet
      Dim vDelegate As EventDelegate

      'First remove the event counts
      If pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsWaiting Or pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsWaitingPaid Or pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsWaitingCreditSale Or pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsWaitingInvoiced Then
        vSubSQL = " number_on_waiting_list = number_on_waiting_list - "
      Else
        If pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsInterested OrElse pEventBooking.BookingStatus = EventBooking.EventBookingStatuses.ebsAwaitingAcceptance Then
          vSubSQL = " number_interested = number_interested - "
        Else
          vSubSQL = " number_of_attendees = number_of_attendees - "
        End If
      End If
      If pSessionList.Length = 0 Then
        If pNewOptionNo = pEventBooking.OptionNumber Then
          'Build up Session List for existing booking
          vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, pEventBooking.BookingNumber)
          vWhereFields.Add("s.session_number", CDBField.FieldTypes.cftLong, "sb.session_number")
          vWhereFields.Add("s.session_type", CDBField.FieldTypes.cftCharacter, "0", CDBField.FieldWhereOperators.fwoNotEqual)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT sb.session_number FROM session_bookings sb, sessions s WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          While vRecordSet.Fetch
            If pSessionList.Length > 0 Then pSessionList = pSessionList & ","
            pSessionList = pSessionList & vRecordSet.Fields(1).Value
          End While
          vRecordSet.CloseRecordSet()
        Else
          'BR 11945: Booking Option has changed; set list as defined by Booking Option
          If pEventBooking.BookingOption.PickSessions = False Then
            vWhereFields.Add("option_number", pNewOptionNo)
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT os.session_number FROM option_sessions os WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
            While vRecordSet.Fetch
              If pSessionList.Length > 0 Then pSessionList = pSessionList & ","
              pSessionList = pSessionList & vRecordSet.Fields(1).Value
            End While
            vRecordSet.CloseRecordSet()
          End If
        End If
      End If

      'Remove the old quantity from the sesssions count - ignore type 0 - then do type 0 if deducting from event
      mvEnv.Connection.ExecuteSQL("UPDATE sessions SET" & vSubSQL & pEventBooking.Quantity & " WHERE session_number IN ( SELECT session_number FROM session_bookings WHERE booking_number = " & pEventBooking.BookingNumber & ") AND session_type <> '0'")
      For Each vBookingOption In BookingOptions
        If vBookingOption.OptionNumber = pEventBooking.OptionNumber Then
          If vBookingOption.DeductFromEvent Then mvEnv.Connection.ExecuteSQL("UPDATE sessions SET" & vSubSQL & pEventBooking.Quantity & " WHERE session_number = " & LowestSessionNumber) 'BaseItemNumber)
        End If
      Next vBookingOption

      'Delete any delegate session activities before deleting the session bookings
      For Each vDelegate In pEventBooking.Delegates
        vDelegate.DeleteSessionActivities()
      Next vDelegate
      'Now delete the session bookings
      vWhereFields = New CDBFields
      vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, pEventBooking.BookingNumber)
      mvEnv.Connection.DeleteRecords("session_bookings", vWhereFields, False)
      If UpdateEventCounts(pNewOptionNo, pNewQty, pSessionList, pNewStatus) Then
        AddSessionBookings(pNewOptionNo, pEventBooking.BookingNumber, pNewQty, pSessionList)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
          vWhereFields.Clear()
          vWhereFields.Add("event_delegate_number", CDBField.FieldTypes.cftLong, "SELECT event_delegate_number FROM delegates WHERE booking_number = " & pEventBooking.BookingNumber, CDBField.FieldWhereOperators.fwoIn)
          mvEnv.Connection.DeleteRecords("delegate_sessions", vWhereFields, False)
        End If
        pEventBooking.SetDelegateSessionsAndActivities(pEventBooking.OptionNumber, Source, pSessionList)
        Return True
      End If
    End Function

    Public Function CheckEventBooking(ByVal pCount As Integer, ByVal pOptionNumber As Integer, ByVal pSessionList As String, Optional ByVal pWaiting As Boolean = False) As CheckEventBookingResults
      Dim vSessionSQL As String
      Dim vCanBook As Boolean
      Dim vCanWait As Boolean
      Dim vBookingOption As New EventBookingOption
      Dim vOptionSessions As New Collection
      Dim vSession As New EventSession
      Dim vRecordSet As CDBRecordSet

      vBookingOption.Init(mvEnv, pOptionNumber)
      If pSessionList.Length > 0 Then
        vSessionSQL = pSessionList
      Else
        vSessionSQL = "SELECT session_number FROM option_sessions WHERE option_sessions.option_number = " & pOptionNumber
      End If
      'First find out if it appears to be possible to add this booking
      vCanBook = True
      vCanWait = True
      If vBookingOption.DeductFromEvent Then
        With BaseSession
          If .NumberOfAttendees + pCount > .MaximumAttendees Then vCanBook = False
          If .NumberOnWaitingList + pCount > .MaximumOnWaitingList Then vCanWait = False
        End With
      End If
      If vCanBook Or vCanWait Then
        With vSession
          .Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & .GetRecordSetFields(EventSession.SessionRecordSetTypes.ssrtAll) & " FROM sessions s WHERE session_number IN (" & vSessionSQL & ") AND session_type <> '0'")
          While vRecordSet.Fetch()
            .InitFromRecordSet(mvEnv, vRecordSet, EventSession.SessionRecordSetTypes.ssrtAll)
            If .NumberOfAttendees + pCount > .MaximumAttendees Then vCanBook = False
            If .NumberOnWaitingList + pCount > .MaximumOnWaitingList Then vCanWait = False
          End While
          vRecordSet.CloseRecordSet()
        End With
      End If
      If vCanBook And pWaiting = False Then
        Return CheckEventBookingResults.cebrCanAddBooking
      Else
        If vCanWait Then
          If pWaiting = False Then
            If pCount = 1 Then
              mvBookingMessage = ProjectText.String25603 'Event or Sessions are fully booked\r\n\r\nWaiting list places are available\r\n\r\nBook places on waiting list?
            Else
              mvBookingMessage = String.Format(ProjectText.String25604, pCount.ToString) 'Requested places (%s) are not available for selected booking option\r\n\r\nWaiting list places are available\r\n\r\nBook places on waiting list?
            End If
          End If
          Return CheckEventBookingResults.cebrCanAddToWaitingList
        Else
          If pCount = 1 Then
            mvBookingMessage = ProjectText.String25605 'Cannot take booking\r\n\r\nEvent or Sessions are fully booked\r\n\r\nWaiting list places are not available
          Else
            mvBookingMessage = String.Format(ProjectText.String25606, pCount.ToString) 'Cannot take booking\r\n\r\nRequested places (%s) are not available for selected booking option\r\n\r\nWaiting list places are not available
          End If
          Return CheckEventBookingResults.cebrCannotAddBooking
        End If
      End If
    End Function

    Public Function Duplicate(ByVal pNewStartDate As String, ByVal pNewDescription As String, ByVal pNewLongDescription As String, ByVal pCopyFromTemplate As Boolean) As CDBEvent
      Dim vBookingOption As EventBookingOption
      Dim vOptionSession As EventOptionSession
      Dim vPersonnel As EventPersonnel
      Dim vResource As EventResource
      Dim vSession As EventSession
      Dim vSessionActivity As EventSessionActivity
      Dim vSessionTest As SessionTest
      Dim vVenueBooking As EventVenueBooking
      Dim vEventTopic As EventTopic
      Dim vEventOwner As EventOwner
      Dim vNewEvent As New CDBEvent(mvEnv)
      Dim vNewBookingOption As EventBookingOption
      Dim vNewOptionSession As EventOptionSession
      Dim vNewOrganiser As EventOrganiser
      Dim vNewPersonnel As EventPersonnel
      Dim vNewResource As EventResource
      Dim vNewExternalResource As ExternalResource
      Dim vNewSession As EventSession
      Dim vNewSessionActivity As EventSessionActivity
      Dim vNewSessionTest As SessionTest
      Dim vNewVenueBooking As EventVenueBooking
      Dim vNewEventTopic As EventTopic
      Dim vNewEventOwner As EventOwner
      Dim vWeekdaysOnly As Boolean
      Dim vBaseSession As EventSession
      Dim vBaseDate As String = ""
      Dim vNewBaseDate As String = ""
      Dim vIndex As Integer
      Dim vNewVenueBookings As List(Of EventVenueBooking)
      Dim vNewInternalResource As InternalResource
      Dim vAppointment As ContactAppointment
      Dim vDays As Integer
      Dim vEndDate As String
      Dim vStartDate As String

      If mvBookingOptions Is Nothing Then InitBookingOptions()
      If mvOptionSessions Is Nothing Then InitOptionSessions()
      If mvOrganiser Is Nothing Then InitOrganiser()
      If mvResources Is Nothing Then InitResources()
      If mvSessions Is Nothing Then InitSessions()
      If mvSessionActivities Is Nothing Then InitSessionActivities()
      If mvEnv.GetConfigOption("ev_session_tests") = True And mvSessionTests Is Nothing Then InitSessionTests()
      If mvVenueBookings Is Nothing Then InitVenueBookings()
      If mvEventTopics Is Nothing Then InitEventTopics()
      If mvEventOwners Is Nothing Then InitEventOwners()

      If Not pCopyFromTemplate Then 'Personnel handled separately if from template
        If mvPersonnel Is Nothing Then InitPersonnel()
      Else
        vWeekdaysOnly = True 'If init from template then only use weekdays
      End If

      vNewEvent.Init()
      If pCopyFromTemplate = False Then
        If mvPersonnel.Count() > 0 And mvEnv.GetConfigOption("ignore_calendar_conflicts") = False Then
          'Check for conflicting appointments before attempting to duplicate the Event
          'Any conflicts will raise an error and prevent the duplication
          For Each vPersonnel In mvPersonnel
            With vPersonnel
              vAppointment = New ContactAppointment(mvEnv)
              vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(StartDate), CDate(.StartDate)))
              vStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDays, CDate(pNewStartDate)))
              vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(StartDate), CDate(.EndDate)))
              vEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDays, CDate(pNewStartDate)))
              vAppointment.Init()
              vAppointment.CheckCalendarConflict(.ContactNumber, vStartDate & " " & .StartTime, vEndDate & " " & .EndTime, ContactAppointment.ContactAppointmentTypes.catPersonnel, .SessionNumber, True)
            End With
          Next vPersonnel
        End If
      End If

      mvEnv.Connection.StartTransaction()

      vNewEvent.InitFromEvent(Me, pNewStartDate, pNewDescription, pNewLongDescription, vWeekdaysOnly, pCopyFromTemplate)
      vNewEvent.Save()
      'Create the new venue bookings first
      vNewVenueBookings = New List(Of EventVenueBooking)
      For Each vVenueBooking In mvVenueBookings
        vNewVenueBooking = New EventVenueBooking
        vNewVenueBooking.Init(mvEnv)
        vNewVenueBooking.InitFromVenueBooking(vVenueBooking, vNewEvent)
        vNewVenueBooking.Save()
        vNewVenueBookings.Add(vNewVenueBooking)
      Next vVenueBooking

      'Sessions are processed in start date order
      For Each vSession In mvSessions
        vNewSession = New EventSession
        vNewSession.Init(mvEnv)
        If Len(vBaseDate) = 0 Then
          vBaseDate = StartDate
          vNewBaseDate = vNewEvent.StartDate
        End If
        vNewSession.InitFromSession(vNewEvent, vBaseDate, vNewBaseDate, vSession, vWeekdaysOnly)
        'Set the venue booking numbers
        If IntegerValue(vSession.VenueBookingNumber) > 0 Then
          For vIndex = 0 To mvVenueBookings.Count - 1
            'TA BC3627 11/12:Added CLng as variant returned wasnt recognised as being the same
            'SDT 18/12  Changed to use VAL - Don't use Clng as it gives a type mismatch on NULLS
            If IntegerValue(vSession.VenueBookingNumber) = mvVenueBookings.Item(vIndex).VenueBookingNumber Then
              vNewSession.VenueBookingNumber = vNewVenueBookings.Item(vIndex).VenueBookingNumber.ToString
              Exit For
            End If
          Next
        End If
        vNewSession.Save()
        If vNewSession.SessionNumber <> vNewEvent.LowestSessionNumber Then  'BaseItemNumber Then
          vBaseDate = vSession.EndDate
          vNewBaseDate = vNewSession.EndDate
        End If
      Next vSession
      If pCopyFromTemplate And mvSessions.Count() > 0 Then
        vBaseSession = mvSessions.Item(LowestSessionNumber.ToString) '(BaseItemNumber.ToString)
        For Each vSession In mvSessions
          If CDate(vSession.EndDate) > CDate(vBaseSession.EndDate) Then
            vBaseSession.NewStartDate = vBaseSession.StartDate
            vBaseSession.NewEndDate = vSession.EndDate
            vBaseSession.SetNewDates()
          End If
        Next vSession
        vBaseSession.Save()
      End If
      For Each vBookingOption In mvBookingOptions
        vNewBookingOption = New EventBookingOption
        vNewBookingOption.Init(mvEnv)
        vNewBookingOption.InitFromBookingOption(vBookingOption, vNewEvent)
        vNewBookingOption.Save()
      Next vBookingOption
      For Each vOptionSession In mvOptionSessions
        vNewOptionSession = New EventOptionSession
        vNewOptionSession.Init(mvEnv)
        vNewOptionSession.InitFromOptionSession(Me, vOptionSession, vNewEvent)
        vNewOptionSession.Save()
      Next vOptionSession
      If mvOrganiser.Existing Then
        vNewOrganiser = New EventOrganiser
        vNewOrganiser.Init(mvEnv)
        vNewOrganiser.InitFromOrganiser(mvOrganiser, vNewEvent)
        vNewOrganiser.Save()
      End If
      For Each vResource In mvResources
        vNewResource = New EventResource(mvEnv)
        vNewResource.Init()
        vNewResource.InitFromResource(Me, vResource, vNewEvent)
        If vResource.ResourceNumber = 0 Then
          'Old format, convert to new Internal Resource
          vNewInternalResource = New InternalResource(mvEnv)
          vNewInternalResource.Init(mvEnv.User.ContactNumber, vResource.ProductCode, vResource.RateCode)
          vNewInternalResource.Save()
          vNewResource.ResourceNumber = vNewInternalResource.ResourceNumber
        ElseIf vResource.ResourceType = EventResource.ResourceTypes.rtExternal Then
          vNewExternalResource = New ExternalResource
          vNewExternalResource.Init(mvEnv)
          vNewExternalResource.InitFromExternalResource((vResource.ExternalResource), vNewEvent)
          vNewExternalResource.Save()
          vNewResource.ResourceNumber = vNewExternalResource.ResourceNumber
        Else
          vNewResource.ResourceNumber = vResource.ResourceNumber
        End If
        vNewResource.Save()
      Next vResource
      For Each vSessionActivity In mvSessionActivities
        vNewSessionActivity = New EventSessionActivity
        vNewSessionActivity.Init(mvEnv)
        vNewSessionActivity.InitFromSessionActivity(Me, vSessionActivity, vNewEvent)
        vNewSessionActivity.Save()
      Next vSessionActivity
      If mvEnv.GetConfigOption("ev_session_tests") = True Then
        For Each vSessionTest In mvSessionTests
          vNewSessionTest = New SessionTest
          vNewSessionTest.Init(mvEnv)
          vNewSessionTest.InitFromSessionTest(Me, vSessionTest, vNewEvent)
          vNewSessionTest.Save()
        Next vSessionTest
      End If
      If Not pCopyFromTemplate Then 'Personnel handled separately if from template
        For Each vPersonnel In mvPersonnel
          vNewPersonnel = New EventPersonnel(mvEnv)
          vNewPersonnel.Init()
          vNewPersonnel.InitFromPersonnel(Me, vPersonnel, vNewEvent)
          If vNewEvent.MultiSession = True Then
            'Do not create appointment for base session
            If vNewPersonnel.SessionNumber <> vNewEvent.LowestSessionNumber Then
              vNewPersonnel.AppointmentDescription = vNewEvent.PersonnelAppointmentDescription(vNewPersonnel)
              vNewPersonnel.Save()
            Else
              vNewPersonnel.SaveWithoutAppointment("")
            End If
          Else
            vNewPersonnel.Save()
          End If
        Next vPersonnel
      End If
      For Each vEventTopic In mvEventTopics
        vNewEventTopic = New EventTopic(mvEnv)
        vNewEventTopic.Init()
        vNewEventTopic.InitFromTopic(vEventTopic, vNewEvent)
        vNewEventTopic.Save()
      Next vEventTopic
      For Each vEventOwner In mvEventOwners
        vNewEventOwner = New EventOwner
        vNewEventOwner.Init(mvEnv)
        vNewEventOwner.InitFromOwner(vEventOwner, vNewEvent)
        vNewEventOwner.Save()
      Next vEventOwner

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) = True And EventPricingMatrix.Length > 0 Then
        'Ensure that Pricing Matrix is valid for the new Event
        vNewEvent.SetValidPricingMatrix()
        vNewEvent.Save()
      End If
      mvEnv.Connection.CommitTransaction()
      Return vNewEvent
    End Function

    Public Sub InitOptionSessions()
      mvOptionSessions = New List(Of EventOptionSession)
      Dim vOptionSession As New EventOptionSession
      vOptionSession.Init(mvEnv)

      'Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOptionSession.GetRecordSetFields(EventOptionSession.EventOptionSessionRecordSetTypes.osrtAll) & " FROM option_sessions WHERE session_number BETWEEN " & BaseItemNumber & " AND " & MaxItemNumber & " ORDER BY session_number, option_number")

      Dim vWhereClause As New CDBFields()
      vWhereClause.Add("s.event_number", EventNumber)

      Dim vAnsiJoin As New AnsiJoins()
      vAnsiJoin.Add("sessions s", "s.session_number", "os.session_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, vOptionSession.GetRecordSetFields(EventOptionSession.EventOptionSessionRecordSetTypes.osrtAll), "option_sessions os", vWhereClause, "os.session_number,os.option_number", vAnsiJoin)
      Dim vRecordSet As CDBRecordSet = vSqlStatement.GetRecordSet

      While vRecordSet.Fetch()
        vOptionSession = New EventOptionSession
        vOptionSession.InitFromRecordSet(mvEnv, vRecordSet, EventOptionSession.EventOptionSessionRecordSetTypes.osrtAll)
        mvOptionSessions.Add(vOptionSession)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitResources()
      mvResources = New List(Of EventResource)
      Dim vResource As New EventResource(mvEnv)
      vResource.Init()
      'Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vResource.GetRecordSetFields() & " FROM event_resources er WHERE session_number BETWEEN " & BaseItemNumber & " AND " & MaxItemNumber & " ORDER BY session_number, product")

      Dim vWhereClause As New CDBFields()
      vWhereClause.Add("s.event_number", EventNumber)

      Dim vAnsiJoin As New AnsiJoins()
      vAnsiJoin.Add("sessions s", "s.session_number", "er.session_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, vResource.GetRecordSetFields(), "event_resources er", vWhereClause, "er.session_number,er.product", vAnsiJoin)
      Dim vRecordSet As CDBRecordSet = vSqlStatement.GetRecordSet

      While vRecordSet.Fetch()
        vResource = New EventResource(mvEnv)
        vResource.InitFromRecordSet(vRecordSet)
        mvResources.Add(vResource)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitSessionActivities()
      Dim vSessionActivity As New EventSessionActivity
      Dim vRecordSet As CDBRecordSet

      mvSessionActivities = New List(Of EventSessionActivity)
      vSessionActivity.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vSessionActivity.GetRecordSetFields(EventSessionActivity.EventSessionActivityRecordSetTypes.esartAll) & " FROM session_activities WHERE event_number = " & EventNumber & " ORDER BY session_number,activity")
      While vRecordSet.Fetch()
        vSessionActivity = New EventSessionActivity
        vSessionActivity.InitFromRecordSet(mvEnv, vRecordSet, EventSessionActivity.EventSessionActivityRecordSetTypes.esartAll)
        mvSessionActivities.Add(vSessionActivity)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitSessionTests()
      mvSessionTests = New List(Of SessionTest)
      Dim vSessionTest As New SessionTest
      vSessionTest.Init(mvEnv)
      'Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vSessionTest.GetRecordSetFields(SessionTest.SessionTestRecordSetTypes.estrtAll) & " FROM session_tests WHERE session_number BETWEEN " & BaseItemNumber & " AND " & MaxItemNumber & " ORDER BY session_number,test_number")

      Dim vWhereClause As New CDBFields()
      vWhereClause.Add("s.event_number", EventNumber)

      Dim vAnsiJoin As New AnsiJoins()
      vAnsiJoin.Add("sessions s", "s.session_number", "st.session_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, vSessionTest.GetRecordSetFields(SessionTest.SessionTestRecordSetTypes.estrtAll), "session_tests st", vWhereClause, "st.session_number,st.test_number", vAnsiJoin)
      Dim vRecordSet As CDBRecordSet = vSqlStatement.GetRecordSet

      While vRecordSet.Fetch()
        vSessionTest = New SessionTest
        vSessionTest.InitFromRecordSet(mvEnv, vRecordSet, SessionTest.SessionTestRecordSetTypes.estrtAll)
        mvSessionTests.Add(vSessionTest)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitVenueBookings()
      mvVenueBookings = New List(Of EventVenueBooking)
      Dim vVenueBooking As New EventVenueBooking
      vVenueBooking.Init(mvEnv)
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vVenueBooking.GetRecordSetFields(EventVenueBooking.EventVenueBookingRecordSetTypes.evbrtAll Or EventVenueBooking.EventVenueBookingRecordSetTypes.evbrtVenueInfo) & " FROM event_venue_bookings evb, venues v WHERE event_number = " & EventNumber & " AND evb.venue = v.venue ORDER BY evb.venue")
      While vRecordSet.Fetch()
        vVenueBooking = New EventVenueBooking
        vVenueBooking.InitFromRecordSet(mvEnv, vRecordSet, EventVenueBooking.EventVenueBookingRecordSetTypes.evbrtAll Or EventVenueBooking.EventVenueBookingRecordSetTypes.evbrtVenueInfo)
        'Sessions are added in start date order
        mvVenueBookings.Add(vVenueBooking)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitEventOwners()
      mvEventOwners = New CollectionList(Of EventOwner)
      Dim vEventOwner As New EventOwner
      vEventOwner.Init(mvEnv)
      Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT department_desc, " & vEventOwner.GetRecordSetFields(EventOwner.EventOwnerRecordSetTypes.eowrtAll) & " FROM event_owners eo, departments d WHERE event_number = " & EventNumber & " AND eo.department = d.department")
      While vRS.Fetch()
        vEventOwner = New EventOwner
        vEventOwner.InitFromRecordSet(mvEnv, vRS, EventOwner.EventOwnerRecordSetTypes.eowrtAll)
        If Not mvEventOwners.ContainsKey(vEventOwner.Department) Then mvEventOwners.Add(vEventOwner.Department, vEventOwner)
      End While
      vRS.CloseRecordSet()
    End Sub

    Public Sub InitPersonnel()
      mvPersonnel = New List(Of EventPersonnel)
      Dim vPersonnel As New EventPersonnel(mvEnv)
      vPersonnel.Init()
      'Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vPersonnel.GetRecordSetFields() & " FROM event_personnel ep WHERE session_number BETWEEN " & BaseItemNumber & " AND " & MaxItemNumber & " ORDER BY session_number")

      Dim vWhereClause As New CDBFields()
      vWhereClause.Add("s.event_number", EventNumber)

      Dim vAnsiJoin As New AnsiJoins()
      vAnsiJoin.Add("sessions s", "s.session_number", "ep.session_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, vPersonnel.GetRecordSetFields(), "event_personnel ep", vWhereClause, "ep.session_number", vAnsiJoin)
      Dim vRecordSet As CDBRecordSet = vSqlStatement.GetRecordSet

      While vRecordSet.Fetch()
        vPersonnel = New EventPersonnel(mvEnv)
        vPersonnel.InitFromRecordSet(vRecordSet)
        mvPersonnel.Add(vPersonnel)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Friend Sub InitFromEvent(ByVal pEvent As CDBEvent, ByRef pNewStartDate As String, ByRef pNewDescription As String, ByRef pNewLongDescription As String, ByRef pWeekdaysOnly As Boolean, ByRef pCopyFromTemplate As Boolean)
      With pEvent
        mvClassFields.Item(EventFields.EventNumber).IntegerValue = mvEnv.GetControlNumber("EV")
        mvClassFields.Item(EventFields.EventDesc).Value = pNewDescription
        mvClassFields.Item(EventFields.LongDescription).Value = pNewLongDescription
        If pWeekdaysOnly Then
          mvClassFields.Item(EventFields.StartDate).Value = NextWeekDay(CDate(pNewStartDate)).ToString(CAREDateFormat)
        Else
          mvClassFields.Item(EventFields.StartDate).Value = pNewStartDate
        End If
        mvClassFields.Item(EventFields.Venue).Value = .Venue
        mvClassFields.Item(EventFields.VenueReference).Value = .VenueReference
        mvClassFields.Item(EventFields.Branch).Value = .Branch
        mvClassFields.Item(EventFields.External).Bool = .External
        mvClassFields.Item(EventFields.MultiSession).Bool = .MultiSession
        mvClassFields.Item(EventFields.QpFormType).Value = .QpFormType
        mvClassFields.Item(EventFields.FreeOfCharge).Bool = .FreeOfCharge
        mvClassFields.Item(EventFields.NextSessionNumber).IntegerValue = .NextSessionNumber
        mvClassFields.Item(EventFields.NextOptionNumber).IntegerValue = .NextOptionNumber
        mvClassFields.Item(EventFields.Source).Value = .Source
        mvClassFields.Item(EventFields.Template).Bool = .Template
        mvClassFields.Item(EventFields.MoveSessionDates).Bool = .MoveSessionDates

        CandidateNumberingMethod = .CandidateNumberingMethod
        mvClassFields.Item(EventFields.FirstCandidateNumber).IntegerValue = .FirstCandidateNumber
        mvClassFields.Item(EventFields.CandidateNumberBlockSize).IntegerValue = .CandidateNumberBlockSize
        mvClassFields.Item(EventFields.EligibilityCheckRequired).Bool = .EligibilityCheckRequired
        mvClassFields.Item(EventFields.EligibilityCheckText).Value = .EligibilityCheckText
        mvClassFields.Item(EventFields.DeferredBookingAct).Value = .DeferredBookingAct
        mvClassFields.Item(EventFields.DeferredBookingActValue).Value = .DeferredBookingActValue
        mvClassFields.Item(EventFields.RejectedBookingAct).Value = .RejectedBookingAct
        mvClassFields.Item(EventFields.RejectedBookingActValue).Value = .RejectedBookingActValue

        mvClassFields.Item(EventFields.Department).Value = .Department
        mvClassFields.Item(EventFields.EventStatus).Value = .EventStatus

        mvClassFields.Item(EventFields.WaitingListControlMethod).Value = .WaitingListControlMethod
        mvClassFields.Item(EventFields.ChargeForWaiting).Bool = .ChargeForWaiting
        mvClassFields.Item(EventFields.EventGroup).Value = .EventGroupCode
        mvClassFields.Item(EventFields.EventClass).Value = .EventClass
        mvClassFields.Item(EventFields.EventPricingMatrix).Value = .EventPricingMatrix

        mvClassFields.Item(EventFields.ActivityGroup).Value = .ActivityGroup
        mvClassFields.Item(EventFields.RelationshipGroup).Value = .RelationshipGroup
        If pCopyFromTemplate Then
          mvClassFields.Item(EventFields.Template).Bool = False 'No longer a template
          mvClassFields.Item(EventFields.Booking).Bool = True 'Can Book (Not sure about this one!)
        End If
        mvClassFields.Item(EventFields.NameAttendees).Bool = .NameAttendees
        'Copying an event should not publish it to the web
        mvClassFields.Item(EventFields.WebPublish).Bool = False

        mvPricingMatrixValid = False
      End With
    End Sub

    Friend Sub SetValidPricingMatrix()
      'When copying an Event or creating from a template, validate the EventPricingMatrix
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
        If EventPricingMatrix.Length > 0 Then
          Dim vEPM As String = ""
          Dim vSQL As String
          vSQL = "SELECT event_pricing_matrix, event_fee_start_date, event_fee_end_date, venue, start_date, end_date FROM sessions s, event_pricing_matrices epm WHERE session_number = " & LowestSessionNumber  'BaseItemNumber
          vSQL = vSQL & " AND (((event_fee_start_date <= s.start_date AND event_fee_end_date >= s.end_date)"
          vSQL = vSQL & " AND (venue = '" & Venue & "' OR venue IS NULL)) OR event_pricing_matrix = '" & EventPricingMatrix & "')"
          vSQL = vSQL & " ORDER BY event_fee_start_date DESC, venue"
          If mvEnv.Connection.NullsSortAtEnd() = False Then vSQL = vSQL & " DESC" 'Ensure the null Venue is at the end
          Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
          If vRS.Fetch Then
            vEPM = vRS.Fields(1).Value 'First record is the Pricing Matrix to use
            Do
              If vRS.Fields(1).Value = EventPricingMatrix Then
                If (DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vRS.Fields(2).Value), CDate(vRS.Fields(5).Value)) >= 0 And DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vRS.Fields(6).Value), CDate(vRS.Fields(3).Value)) >= 0) Then
                  'Current PricingMatrix is valid so keep it
                  vEPM = ""
                End If
              End If
            Loop While vRS.Fetch And vEPM.Length > 0
          End If
          vRS.CloseRecordSet()
          If vEPM.Length > 0 Then
            'If we found a new, valid Pricing Matrix then update it, otherwise leave the original one
            If vEPM <> EventPricingMatrix Then
              mvClassFields.Item(EventFields.EventPricingMatrix).Value = vEPM
              'BR16770 Event pricing matrix does exist for current date range.
              mvPricingMatrixValid = True
            Else
              'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
              mvPricingMatrixValid = False
              mvClassFields.Item(EventFields.Booking).Value = "N"
            End If
          End If
        End If
      End If
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vEventSource As New EventSource
      Dim vPreviousSource As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        If Source.Length > 0 Then
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("event_number", EventNumber)
          vWhereFields.Add("source", Source)
          If mvEnv.Connection.GetCount("event_sources", vWhereFields) = 0 Then
            With vEventSource
              .Init(mvEnv)
              .EventNumber = EventNumber
              .Source = Source
              .Save()
            End With
          End If
        End If
        If mvClassFields.Item(EventFields.Source).ValueChanged And mvClassFields.Item(EventFields.Source).SetValue.Length > 0 Then
          vPreviousSource = mvClassFields.Item(EventFields.Source).SetValue
        End If
      End If
      MyBase.Save(pAmendedBy, pAudit, 0)
      If vPreviousSource.Length > 0 Then
        'Remove old Event Source if necessary
        vEventSource.Init(mvEnv, EventNumber, vPreviousSource)
        If vEventSource.Existing Then vEventSource.Delete()
      End If
    End Sub

    Public Overloads Sub Save(ByVal pAmendedBy As String, ByVal pAHC As ClassFields.AmendmentHistoryCreation)
      Save(pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC), 0)
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      mvClassFields.SetControlNumber(mvEnv)
      If Department.Length = 0 Then mvClassFields.Item(EventFields.Department).Value = mvEnv.User.Department
      If pParameterList.Exists("Product") Then mvClassFields.Item(EventFields.SponsorshipProduct).Value = pParameterList("Product").Value
      If pParameterList.Exists("Rate") Then mvClassFields.Item(EventFields.SponsorshipRate).Value = pParameterList("Rate").Value

      pParameterList.Add("SessionType", "0")
      pParameterList.Add("SessionDesc", String.Format(ProjectText.String26125, TruncateString(EventDesc, 42)))
      pParameterList.Add("EventNumber", EventNumber)
      BaseSession.Create(mvEnv, Me, pParameterList)
      CheckValidity(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.Exists("Product") Then mvClassFields.Item(EventFields.SponsorshipProduct).Value = pParameterList("Product").Value
      If pParameterList.Exists("Rate") Then mvClassFields.Item(EventFields.SponsorshipRate).Value = pParameterList("Rate").Value
      CheckValidity(pParameterList)
    End Sub

    Private Sub CheckValidity(ByRef pParams As CDBParameters)
      If External And FreeOfCharge Then
        RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String26107) 'Cannot set Event as external and free of charge
      End If
      If External And pParams.HasValue("Organiser") = False Then 'Relies on the organiser parameter being supplied
        RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17382) 'Organiser must be selected for external event
      End If
      Dim vBookingOption As EventBookingOption
      If EligibilityCheckRequired Then
        If Len(EligibilityCheckText) = 0 Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17383) 'Eligibility Check Text must be specified if Eligibility Check is set
        ElseIf Len(DeferredBookingAct) = 0 Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17384) 'Deferred Activity must be specified if Eligibility Check is set
        ElseIf Len(DeferredBookingActValue) = 0 Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17385) 'Deferred Activity Value must be specified if Eligibility Check is set
        ElseIf Len(RejectedBookingAct) = 0 Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17386) 'Rejected Activity must be specified if Eligibility Check is set
        ElseIf Len(RejectedBookingActValue) = 0 Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String17387) 'Rejected Activity Value must be specified if Eligibility Check is set
        End If
        If mvClassFields(EventFields.EligibilityCheckRequired).ValueChanged And mvExisting Then
          For Each vBookingOption In BookingOptions
            If vBookingOption.MaximumBookings > 1 Then
              RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String26103) 'Some booking options have the maximum bookings greater then 1, eligibility check can not be set
            End If
          Next vBookingOption
        End If
      End If
      Dim vSQL As String
      If Booking And mvClassFields(EventFields.Booking).ValueChanged Then
        If Not mvExisting Then
          RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String26104) 'Bookings cannot be taken - Booking options must be defined
        Else
          If BookingOptions.Count() = 0 Then
            RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String26105) 'Bookings cannot be taken - No booking options are defined
          ElseIf MultiSession Then
            For Each vBookingOption In BookingOptions
              Dim vOptionSessionCount As Integer = vBookingOption.OptionSessions.Count()
              If vBookingOption.IssueEventResources Then vOptionSessionCount = vOptionSessionCount - 1
              If vOptionSessionCount <= vBookingOption.NumberOfSessions And vBookingOption.PickSessions = True Then
                RaiseError(DataAccessErrors.daeEventParameterError, ProjectText.String26106) 'Bookings cannot be taken - Not all booking options have the correct number of sessions defined
              End If
            Next vBookingOption
          End If
        End If
        If mvEnv.GetConfigOption("ev_check_mandatory_topics") Then
          Dim vDS As New DataSelection(mvEnv, DataSelection.DataSelectionTypes.dstEventSelectionPages, Nothing, DataSelection.DataSelectionListType.dsltUser, DataSelection.DataSelectionUsages.dsuSmartClient, EventGroupCode)
          Dim vItems() As String = Split(vDS.DisplayColumns, ",")
          Dim vMissingTopic As String = ""
          Dim vTopicGroup As String
          For Each vItem As String In vItems
            If vItem.StartsWith("TopicGroup") Then
              vTopicGroup = vItem.Substring(10)
              vSQL = "SELECT /* SQLServerCSC */ topic_desc FROM topic_group_details tgd INNER JOIN topics t ON tgd.topic = t.topic"
              vSQL = vSQL & " LEFT OUTER JOIN (SELECT topic,sub_topic FROM event_topics WHERE event_number = " & EventNumber & ")et ON tgd.topic = et.topic"
              vSQL = vSQL & " WHERE tgd.topic_group = '" & vTopicGroup & "' AND mandatory = 'Y' AND et.sub_topic IS NULL"
              vMissingTopic = mvEnv.Connection.GetValue(vSQL)
              If vMissingTopic.Length > 0 Then RaiseError(DataAccessErrors.daeEventTopicMandatory, vMissingTopic)
            End If
          Next
        End If
      End If
      'cannot change the sponsorship product and rate if sponsorship payments have been received.
      If Existing And (mvClassFields(EventFields.SponsorshipProduct).ValueChanged Or mvClassFields(EventFields.SponsorshipRate).ValueChanged) Then
        CalculateSponsorshipIncome()
        If Val(SponsorshipIncome) > 0 Then RaiseError(DataAccessErrors.daeCannotUpdateSponsorshipAsPayments)
      End If
    End Sub

    Public Sub RenumberCandidates(ByVal pRenumberAll As Boolean)
      Dim vDelegate As New EventDelegate
      Dim vRecordSet As CDBRecordSet
      Dim vSession As EventSession
      Dim vOption As EventBookingOption
      Dim vWhereFields As New CDBFields
      Dim vFields As New CDBFields
      Dim vNumber As Integer
      Dim vBaseNumber As Integer

      If pRenumberAll Then
        'If Renumbering All then clear the numbers first
        vWhereFields.Add("event_number", CDBField.FieldTypes.cftLong, EventNumber)
        vFields.Add("candidate_number", CDBField.FieldTypes.cftLong)
        mvEnv.Connection.UpdateRecords("delegates", vFields, vWhereFields, False)
        vNumber = FirstCandidateNumber
      End If
      vDelegate.Init(mvEnv)

      If CandidateNumberingMethod = EventCandidateNumberingMethods.ecnmFirstSessionSequence Then
        InitSessions()
        If Not pRenumberAll Then
          vNumber = IntegerValue(mvEnv.Connection.GetValue("SELECT MAX (candidate_number) FROM delegates ed WHERE event_number = " & EventNumber)) + 1
        End If
        For Each vSession In mvSessions
          If MultiSession And vSession.SessionType = vSession.BaseSessionType Then
            'Ignore base session for multi session event
          Else
            'Select the correct set of delegates
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vDelegate.GetRecordSetFields(EventDelegate.EventDelegateRecordSetTypes.edrtNumbers) & " FROM session_bookings sb, delegates ed, contacts c WHERE sb.session_number = " & vSession.SessionNumber & " AND sb.booking_number = ed.booking_number AND candidate_number IS NULL AND ed.contact_number = c.contact_number ORDER BY surname")
            While vRecordSet.Fetch()
              vDelegate.InitFromRecordSet(mvEnv, vRecordSet, EventDelegate.EventDelegateRecordSetTypes.edrtNumbers)
              vDelegate.CandidateNumber = vNumber
              vNumber = vNumber + 1
              vDelegate.Save()
            End While
          End If
        Next vSession
      Else
        InitBookingOptions(True)
        vBaseNumber = FirstCandidateNumber
        For Each vOption In mvBookingOptions
          If Not pRenumberAll Then
            vNumber = IntegerValue(mvEnv.Connection.GetValue("SELECT MAX (candidate_number) FROM event_bookings eb, delegates ed WHERE eb.option_number = " & vOption.OptionNumber & " AND eb.booking_number = ed.booking_number")) + 1
            If vNumber = 1 Then vNumber = vBaseNumber 'None found
          Else
            vNumber = vBaseNumber
          End If
          'Select the correct set of delegates
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vDelegate.GetRecordSetFields(EventDelegate.EventDelegateRecordSetTypes.edrtNumbers) & " FROM event_bookings eb, delegates ed, contacts c WHERE eb.option_number = " & vOption.OptionNumber & " AND eb.booking_number = ed.booking_number AND candidate_number IS NULL AND ed.contact_number = c.contact_number ORDER BY surname")
          While vRecordSet.Fetch()
            vDelegate.InitFromRecordSet(mvEnv, vRecordSet, EventDelegate.EventDelegateRecordSetTypes.edrtNumbers)
            vDelegate.CandidateNumber = vNumber
            vNumber = vNumber + 1
            vDelegate.Save()
          End While
          vBaseNumber = vBaseNumber + CandidateNumberBlockSize
        Next vOption
      End If
    End Sub

    Public Sub RenumberSessionBookings(ByVal pRenumberAll As Boolean)
      Dim vSessionCandidateNumber As New SessionCandidateNumber(mvEnv)
      Dim vDelegate As New EventDelegate
      Dim vRecordSet As CDBRecordSet
      Dim vSession As EventSession
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vSCNumber As Integer 'Session Candidate Number

      vDelegate.Init(mvEnv)
      InitSessions()

      If pRenumberAll Then
        'If Renumbering All then clear the numbers first
        vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong)
        For Each vSession In mvSessions
          vWhereFields(1).Value = CStr(vSession.SessionNumber)
          mvEnv.Connection.DeleteRecords("session_candidate_numbers", vWhereFields, False)
        Next vSession
      End If

      For Each vSession In mvSessions
        If pRenumberAll Then
          vSCNumber = 1
        Else
          vSCNumber = IntegerValue(mvEnv.Connection.GetValue("SELECT MAX (session_candidate_number) FROM sessions s, session_candidate_numbers scn WHERE s.event_number = " & EventNumber & " AND s.session_number = " & vSession.SessionNumber & " AND scn.session_number = s.session_number")) + 1
        End If
        If MultiSession And vSession.SessionType = vSession.BaseSessionType Then
          'Ignore base session for multi session event
        Else
          'Select the correct set of delegates
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT event_delegate_number,session_number FROM session_bookings sb, delegates ed, contacts c WHERE sb.session_number = " & vSession.SessionNumber & " AND sb.booking_number = ed.booking_number AND event_delegate_number NOT IN (SELECT event_delegate_number FROM session_candidate_numbers scn WHERE scn.session_number = sb.session_number and scn.event_delegate_number = ed.event_delegate_number) AND ed.contact_number = c.contact_number ORDER BY surname")
          While vRecordSet.Fetch()
            vSessionCandidateNumber.Init()
            vSessionCandidateNumber.SetCandidate(vRecordSet.Fields(1).IntegerValue, vRecordSet.Fields(2).IntegerValue, vSCNumber)
            vSCNumber = vSCNumber + 1
            vSessionCandidateNumber.Save()
          End While
        End If
      Next vSession
    End Sub

    Private Function GetEventSession() As CDBDataTable
      Dim vSQL As New SQLStatement(mvEnv.Connection, "session_number", "sessions s", New CDBFields().Add("event_number", Me.EventNumber))
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQL(mvEnv, vSQL)
      Return vDataTable
    End Function

    'Private Function GetEventPersonnel() As CDBDataTable
    '  Dim vSQL As New SQLStatement(mvEnv.Connection, "session_number", "sessions s", New CDBFields().Add("event_number", Me.EventNumber))
    '  Dim vDataTable As New CDBDataTable
    '  vDataTable.FillFromSQL(mvEnv, vSQL)
    '  Return vDataTable
    'End Function

    Private Sub DeleteContactAppointment(ByVal pSessionNumber As String)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("unique_id", pSessionNumber, CDBField.FieldWhereOperators.fwoIn)
      vWhereFields.Add("record_type", "'E','P'", CDBField.FieldWhereOperators.fwoIn)
      mvEnv.Connection.DeleteRecords("contact_appointments", vWhereFields, False)
    End Sub

  End Class

  
End Namespace

