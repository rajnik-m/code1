

Namespace Access
  Public Class ServiceBooking

    Public Enum ServiceBookingRecordSetTypes 'These are bit values
      sbrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ServiceBookingFields
      sbfAll = 0
      sbfServiceBookingNumber
      sbfBookingContactNumber
      sbfBookingAddressNumber
      sbfServiceContactNumber
      sbfRelatedContactNumber
      sbfStartDate
      sbfEndDate
      sbfBookingStatus
      sbfBatchNumber
      sbfTransactionNumber
      sbfLineNumber
      sbfSalesContactNumber
      sbfOrderNumber
      sbfTransactionDate
      sbfAmount
      sbfVatAmount
      sbfVatRate
      sbfCancellationReason
      sbfCancelledBy
      sbfCancelledOn
      sbfAmendedBy
      sbfAmendedOn
      sbfCancellationSource
    End Enum

    Public Enum ServiceBookingStatuses
      sbsBooked 'B
      sbsCancelled 'C
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvServiceControl As ServiceControl
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "service_bookings"
          .Add("service_booking_number", CDBField.FieldTypes.cftLong)
          .Add("booking_contact_number", CDBField.FieldTypes.cftLong)
          .Add("booking_address_number", CDBField.FieldTypes.cftLong)
          .Add("service_contact_number", CDBField.FieldTypes.cftLong)
          .Add("related_contact_number", CDBField.FieldTypes.cftLong)
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("end_date", CDBField.FieldTypes.cftDate)
          .Add("booking_status")
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("sales_contact_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_date", CDBField.FieldTypes.cftDate)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric)
          .Add("vat_rate")
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("cancellation_source")
        End With
        mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub
    Private Sub SetValid(ByRef pField As ServiceBookingFields)
      'Add code here to ensure all values are valid before saving

      If mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue = 0 Then mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue = mvEnv.GetControlNumber("SB")
      mvClassFields.Item(ServiceBookingFields.sbfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ServiceBookingFields.sbfAmendedBy).Value = mvEnv.User.Logname
    End Sub
    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ServiceBookingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ServiceBookingRecordSetTypes.sbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sb")
      End If
      GetRecordSetFields = vFields
    End Function
    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pServiceBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pServiceBookingNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ServiceBookingRecordSetTypes.sbrtAll) & " FROM service_bookings WHERE service_booking_number = " & pServiceBookingNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ServiceBookingRecordSetTypes.sbrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub
    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ServiceBookingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And ServiceBookingRecordSetTypes.sbrtAll) = ServiceBookingRecordSetTypes.sbrtAll Then
          .SetItem(ServiceBookingFields.sbfServiceBookingNumber, vFields)
          .SetItem(ServiceBookingFields.sbfBookingContactNumber, vFields)
          .SetItem(ServiceBookingFields.sbfBookingAddressNumber, vFields)
          .SetItem(ServiceBookingFields.sbfServiceContactNumber, vFields)
          .SetItem(ServiceBookingFields.sbfRelatedContactNumber, vFields)
          .SetItem(ServiceBookingFields.sbfStartDate, vFields)
          .SetItem(ServiceBookingFields.sbfEndDate, vFields)
          .SetItem(ServiceBookingFields.sbfBookingStatus, vFields)
          .SetItem(ServiceBookingFields.sbfBatchNumber, vFields)
          .SetItem(ServiceBookingFields.sbfTransactionNumber, vFields)
          .SetItem(ServiceBookingFields.sbfLineNumber, vFields)
          .SetItem(ServiceBookingFields.sbfSalesContactNumber, vFields)
          .SetItem(ServiceBookingFields.sbfOrderNumber, vFields)
          .SetItem(ServiceBookingFields.sbfTransactionDate, vFields)
          .SetItem(ServiceBookingFields.sbfAmount, vFields)
          .SetItem(ServiceBookingFields.sbfVatAmount, vFields)
          .SetItem(ServiceBookingFields.sbfVatRate, vFields)
          .SetItem(ServiceBookingFields.sbfCancellationReason, vFields)
          .SetItem(ServiceBookingFields.sbfCancelledBy, vFields)
          .SetItem(ServiceBookingFields.sbfCancelledOn, vFields)
          .SetItem(ServiceBookingFields.sbfAmendedBy, vFields)
          .SetItem(ServiceBookingFields.sbfAmendedOn, vFields)
          .SetOptionalItem(ServiceBookingFields.sbfCancellationSource, vFields)
        End If
      End With
    End Sub

    Public Sub Delete()
      Dim vWhereFields As CDBFields

      mvEnv.Connection.StartTransaction()
      vWhereFields = New CDBFields
      vWhereFields.Add("unique_id", CDBField.FieldTypes.cftLong, ServiceBookingNumber)
      vWhereFields.Add("record_type", CDBField.FieldTypes.cftCharacter, "S")
      mvEnv.Connection.DeleteRecords("contact_appointments", vWhereFields, False)
      vWhereFields = New CDBFields
      vWhereFields.Add("service_booking_number", CDBField.FieldTypes.cftLong, ServiceBookingNumber)
      mvEnv.Connection.DeleteRecords("service_booking_revenue", vWhereFields, False)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then mvEnv.Connection.DeleteRecords("service_booking_transactions", vWhereFields, False)
      mvClassFields.Delete(mvEnv.Connection)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SetTransactionInfo(ByVal pEnv As CDBEnvironment, ByVal pBookingNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pSalesContactNumber As Integer)
      'This routine initialises the object and sets the primary key values
      'To act as though it has just read the data
      'This is used so we can update it without reading it first
      Init(pEnv)
      mvClassFields(ServiceBookingFields.sbfServiceBookingNumber).SetValue = CStr(pBookingNumber)
      mvExisting = True
      mvClassFields(ServiceBookingFields.sbfBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(ServiceBookingFields.sbfTransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields(ServiceBookingFields.sbfLineNumber).IntegerValue = pLineNumber
      If pSalesContactNumber > 0 Then mvClassFields(ServiceBookingFields.sbfSalesContactNumber).IntegerValue = pSalesContactNumber
    End Sub

    Public Sub Save()
      SetValid(ServiceBookingFields.sbfAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Public Sub Cancel(ByVal pCancellationReason As String, ByRef pLogname As String, ByVal pCancelDate As String, ByVal pCancellationSource As String, Optional ByVal pInTransaction As Boolean = False, Optional ByVal pCreateNegativeSB As Boolean = True)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vRowsAffected As Integer
      Dim vInsertFields As New CDBFields
      Dim vAmount As Double
      Dim vVatAmount As Double
      Dim vSB As Integer
      Dim vRS As CDBRecordSet
      Dim vStartDate As String = ""
      Dim vFinYear As Integer
      Dim vFinPeriod As Integer
      Dim vValuePerDay As Double
      Dim vDaysUsed As Integer
      Dim vPercentage As Double
      Dim vSBRDays As Integer

      On Error GoTo CancelError

      If Not pInTransaction Then mvEnv.Connection.StartTransaction()

      vWhereFields.Add("service_booking_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue)
      With vUpdateFields
        .Add("cancellation_reason", CDBField.FieldTypes.cftCharacter, pCancellationReason)
        .Add("cancelled_on", CDBField.FieldTypes.cftDate, pCancelDate)
        .Add("cancelled_by", CDBField.FieldTypes.cftCharacter, pLogname)
        .Add("booking_status", CDBField.FieldTypes.cftCharacter, mvEnv.GetServiceBookingStatusCode(ServiceBookingStatuses.sbsCancelled))
        If Len(pCancellationSource) > 0 Then .Add("cancellation_source", CDBField.FieldTypes.cftCharacter, pCancellationSource)
      End With
      vRowsAffected = mvEnv.Connection.UpdateRecords("service_bookings", vUpdateFields, vWhereFields)

      vWhereFields = New CDBFields
      With vWhereFields
        .Add("contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfServiceContactNumber).IntegerValue)
        .Add("unique_id", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue)
      End With
      vRowsAffected = mvEnv.Connection.DeleteRecords("contact_appointments", vWhereFields)

      If pCreateNegativeSB Then
        'create -ve SB Revenue
        vSB = mvEnv.GetControlNumber("SB")
        '1. Get the details of the financial period in which the SB is being cancelled
        vRS = mvEnv.Connection.GetRecordSet("SELECT * FROM calendar WHERE start_date" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pCancelDate) & " ORDER BY start_date DESC")
        With vRS
          If .Fetch() = True Then
            vStartDate = .Fields("start_date").Value
            vFinYear = .Fields("fin_year").IntegerValue
            vFinPeriod = .Fields("fin_period").IntegerValue
          End If
          .CloseRecordSet()
        End With
        '2. Create -ve SB Revenue records for the financial periods after the financial period in which the SB is being cancelled
        vRS = mvEnv.Connection.GetRecordSet("SELECT " & vSB & " AS  service_booking_number,sbr.fin_year,sbr.fin_period,sbr.fin_period_days,-sbr.fin_period_value AS  fin_period_value FROM service_booking_revenue sbr, calendar c WHERE sbr.service_booking_number = " & mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue & " AND sbr.fin_year = c.fin_year AND sbr.fin_period = c.fin_period AND c.start_date" & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, pCancelDate))
        With vRS
          While .Fetch() = True
            mvEnv.Connection.InsertRecord("service_booking_revenue", .Fields)
          End While
          .CloseRecordSet()
        End With
        '3. Select the original SB Revenue record for the period in which the SB is being cancelled
        vRS = mvEnv.Connection.GetRecordSet("SELECT * FROM service_booking_revenue WHERE service_booking_number = " & mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue & " AND fin_year = " & vFinYear & " AND fin_period = " & vFinPeriod)
        With vRS
          If .Fetch() = True Then
            '4. Calculate the value/day from that record
            vValuePerDay = .Fields("fin_period_value").DoubleValue / .Fields("fin_period_days").IntegerValue
            '5. Calculate the number of days into the period before the orig SB was cancelled
            'if SB start date >= cancellation date...
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pCancelDate), CDate(mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value)) >= 0 Then
              vDaysUsed = 0 'cancelled before or on the day it started, so no days used
            Else
              'if the SB start date > FP start date AND cancellation date >= SB start date...
              If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value)) > 0 And DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value), CDate(pCancelDate)) >= 0 Then
                vStartDate = mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value
              End If
              'if cancellation date = SB end date...
              If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value), CDate(pCancelDate)) = 0 Then
                'vDaysUsed = DateDiff("d", mvClassFields.Item(sbfStartDate).Value, mvClassFields.Item(sbfEndDate).Value) + 1
                vDaysUsed = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value)) + 1)
              Else
                vDaysUsed = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(pCancelDate)))
              End If
            End If
            '6. Calculate the value for the period for the -ve SB
            vSBRDays = .Fields("fin_period_days").IntegerValue - vDaysUsed
            vAmount = -Int(((vSBRDays * vValuePerDay) * 100) + 0.5) / 100
            InsertSBR(.Fields("fin_year").IntegerValue, (.Fields("fin_period").IntegerValue), vSBRDays, vAmount, vSB)
          End If
          .CloseRecordSet()
        End With
        '7. Get the total value of the -ve SB
        vRS = mvEnv.Connection.GetRecordSet("SELECT SUM(fin_period_value)  AS  total_sb_rev FROM service_booking_revenue WHERE service_booking_number = " & vSB)
        With vRS
          If .Fetch() = True Then
            vAmount = -.Fields(1).DoubleValue
          End If
          .CloseRecordSet()
        End With
        '8. create -ve SB
        vInsertFields = New CDBFields
        With vInsertFields
          .Add("service_booking_number", CDBField.FieldTypes.cftLong, vSB)
          .Add("booking_contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfBookingContactNumber).IntegerValue)
          .Add("booking_address_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfBookingAddressNumber).IntegerValue)
          .Add("service_contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfServiceContactNumber).IntegerValue)
          If mvClassFields.Item(ServiceBookingFields.sbfRelatedContactNumber).IntegerValue > 0 Then .Add("related_contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfRelatedContactNumber).IntegerValue)
          If vDaysUsed = 0 Then
            .Add("start_date", CDBField.FieldTypes.cftDate, mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value)
          Else
            .Add("start_date", CDBField.FieldTypes.cftDate, pCancelDate)
          End If
          .Add("end_date", CDBField.FieldTypes.cftDate, mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value)
          If mvClassFields.Item(ServiceBookingFields.sbfSalesContactNumber).IntegerValue > 0 Then .Add("sales_contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfSalesContactNumber).IntegerValue)
          .AddAmendedOnBy(mvEnv.User.Logname)
          vPercentage = CDbl(mvEnv.Connection.GetValue("SELECT percentage FROM vat_rates WHERE vat_rate = '" & mvClassFields.Item(ServiceBookingFields.sbfVatRate).Value & "'"))
          vVatAmount = Int(((vAmount - (vAmount / (1 + vPercentage / 100))) * 100) + 0.5) / 100
          .Add("amount", CDBField.FieldTypes.cftNumeric, -vAmount)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric, -vVatAmount)
          .Add("vat_rate", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(ServiceBookingFields.sbfVatRate).Value)
          .Add("transaction_date", CDBField.FieldTypes.cftDate, pCancelDate)
        End With
        mvEnv.Connection.InsertRecord("service_bookings", vInsertFields)
      End If

      If Not pInTransaction Then mvEnv.Connection.CommitTransaction()

      Exit Sub

CancelError:
      mvEnv.Connection.RollbackTransaction()
      Err.Raise(Err.Number, Err.Source, Err.Description)
    End Sub
    Public Sub CreateRevenue()
      Dim vStartDate As String
      Dim vEndDate As String
      Dim vAmount As Double
      Dim vRecordSet As CDBRecordSet
      Dim vAmountPerDay As Double
      Dim vRevenueAmount As Double
      Dim vRevenueTotal As Double
      Dim vPeriodDays As Integer
      Dim vFirstPeriod As Boolean
      Dim vLastPeriodStart As String = ""
      Dim vDeleteFields As New CDBFields
      Dim vFinYear As Integer
      Dim vFinPeriod As Integer
      Dim vRowsAffected As Integer

      vStartDate = mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value
      vEndDate = mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value
      vAmount = Val(mvClassFields.Item(ServiceBookingFields.sbfAmount).Value)

      vDeleteFields.Add("service_booking_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue)
      vRowsAffected = mvEnv.Connection.DeleteRecords("service_booking_revenue", vDeleteFields, False)

      vAmountPerDay = vAmount / (DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(vEndDate)) + 1)
      vFirstPeriod = True

      'Ensure that all financial periods exist
      If mvEnv.Connection.GetCount("calendar", Nothing, "start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, vEndDate)) < 1 Then
        RaiseError(DataAccessErrors.daeNoFinancialPeriod, vEndDate)
      End If
      'Get the first period
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT * FROM calendar WHERE start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, vStartDate) & " ORDER BY start_date DESC")
      With vRecordSet
        If .Fetch() = True Then
          vFinYear = IntegerValue(.Fields.Item("fin_year").Value)
          vFinPeriod = IntegerValue(.Fields.Item("fin_period").Value)
          vLastPeriodStart = .Fields.Item("start_date").Value
        Else
          RaiseError(DataAccessErrors.daeNoFinancialPeriod, vStartDate)
        End If
        .CloseRecordSet()
      End With
      'Get the rest of the periods
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT * FROM calendar WHERE start_date " & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, vStartDate) & " AND start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, vEndDate) & " ORDER BY start_date")
      With vRecordSet
        While .Fetch() = True
          If vFirstPeriod Then
            vPeriodDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(.Fields.Item("start_date").Value)))
            vFirstPeriod = False
          Else
            vPeriodDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vLastPeriodStart), CDate(.Fields.Item("start_date").Value)))
          End If
          vRevenueAmount = Int(((vAmountPerDay * vPeriodDays) * 100) + 0.5) / 100
          vRevenueTotal = vRevenueTotal + vRevenueAmount
          InsertSBR(vFinYear, vFinPeriod, vPeriodDays, vRevenueAmount)
          vLastPeriodStart = .Fields.Item("start_date").Value
          vFinYear = .Fields.Item("fin_year").IntegerValue
          vFinPeriod = .Fields.Item("fin_period").IntegerValue
        End While
        .CloseRecordSet()
      End With
      'Do the last period
      If vFirstPeriod Then 'service booking begins and ends in the same financial period
        vPeriodDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vStartDate), CDate(vEndDate)) + 1)
      Else
        vPeriodDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vLastPeriodStart), CDate(vEndDate)) + 1)
      End If
      'Because the number of days booked may not divide into the SB amount evenly
      'sum each revenue amount and for the last period make the value vAmount - sum of other period's amounts.
      vRevenueAmount = FixTwoPlaces(vAmount - vRevenueTotal)
      InsertSBR(vFinYear, vFinPeriod, vPeriodDays, vRevenueAmount)
    End Sub

    Public ReadOnly Property TransactionProcessed() As Boolean
      Get
        TransactionProcessed = BatchNumber > 0 And TransactionNumber > 0 And LineNumber > 0
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
    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(ServiceBookingFields.sbfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ServiceBookingFields.sbfAmendedOn).Value
      End Get
    End Property

    Public Property Amount() As Double
      Get
        Amount = mvClassFields.Item(ServiceBookingFields.sbfAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(ServiceBookingFields.sbfAmount).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(ServiceBookingFields.sbfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BookingAddressNumber() As Integer
      Get
        BookingAddressNumber = mvClassFields.Item(ServiceBookingFields.sbfBookingAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BookingContactNumber() As Integer
      Get
        BookingContactNumber = mvClassFields.Item(ServiceBookingFields.sbfBookingContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BookingStatus() As ServiceBookingStatuses
      Get
        Select Case mvClassFields.Item(ServiceBookingFields.sbfBookingStatus).Value
          Case "B"
            Return ServiceBooking.ServiceBookingStatuses.sbsBooked
          Case "C"
            Return ServiceBooking.ServiceBookingStatuses.sbsCancelled
        End Select
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(ServiceBookingFields.sbfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(ServiceBookingFields.sbfCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(ServiceBookingFields.sbfCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(ServiceBookingFields.sbfCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property TransactionDate() As String
      Get
        TransactionDate = mvClassFields.Item(ServiceBookingFields.sbfTransactionDate).Value
      End Get
    End Property

    Public Property EndDate() As String
      Get
        EndDate = mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ServiceBookingFields.sbfEndDate).Value = Value
      End Set
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(ServiceBookingFields.sbfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OrderNumber() As Integer
      Get
        OrderNumber = mvClassFields.Item(ServiceBookingFields.sbfOrderNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RelatedContactNumber() As Integer
      Get
        RelatedContactNumber = mvClassFields.Item(ServiceBookingFields.sbfRelatedContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SalesContactNumber() As Integer
      Get
        SalesContactNumber = mvClassFields.Item(ServiceBookingFields.sbfSalesContactNumber).IntegerValue
      End Get
    End Property

    Public Property ServiceBookingNumber() As Integer
      Get
        ServiceBookingNumber = mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property ServiceContactNumber() As Integer
      Get
        ServiceContactNumber = mvClassFields.Item(ServiceBookingFields.sbfServiceContactNumber).IntegerValue
      End Get
    End Property

    Public Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ServiceBookingFields.sbfStartDate).Value = Value
      End Set
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(ServiceBookingFields.sbfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property VatAmount() As Double
      Get
        VatAmount = mvClassFields.Item(ServiceBookingFields.sbfVatAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property VATRate() As String
      Get
        VATRate = mvClassFields.Item(ServiceBookingFields.sbfVatRate).Value
      End Get
    End Property

    Public ReadOnly Property ServiceControl() As ServiceControl
      Get
        Dim vRS As CDBRecordSet

        If mvServiceControl Is Nothing Then
          mvServiceControl = New ServiceControl
          mvServiceControl.Init(mvEnv)
          If mvExisting Then
            vRS = mvEnv.Connection.GetRecordSet("SELECT " & mvServiceControl.GetRecordSetFields(ServiceControl.ServiceControlRecordSetTypes.svcrtAll) & " FROM contacts c, service_controls sc WHERE c.contact_number = " & mvClassFields.Item(ServiceBookingFields.sbfServiceContactNumber).IntegerValue & " AND c.contact_group = sc.contact_group")
            If vRS.Fetch() = True Then
              mvServiceControl = New ServiceControl
              mvServiceControl.InitFromRecordSet(mvEnv, vRS, ServiceControl.ServiceControlRecordSetTypes.svcrtAll)
            End If
            vRS.CloseRecordSet()
          End If
        End If
        ServiceControl = mvServiceControl
      End Get
    End Property

    Private Sub InsertSBR(ByVal pFinYear As Integer, ByRef pFinPeriod As Integer, ByRef pPeriodDays As Integer, ByRef pAmount As Double, Optional ByRef pServiceBookingNumber As Integer = 0)
      Dim vInsertFields As CDBFields

      vInsertFields = New CDBFields
      With vInsertFields
        .Add("service_booking_number", If(pServiceBookingNumber > 0, pServiceBookingNumber, mvClassFields.Item(ServiceBookingFields.sbfServiceBookingNumber).IntegerValue))
        .Add("fin_year", CDBField.FieldTypes.cftInteger, pFinYear)
        .Add("fin_period", CDBField.FieldTypes.cftInteger, pFinPeriod)
        .Add("fin_period_days", CDBField.FieldTypes.cftInteger, pPeriodDays)
        .Add("fin_period_value", CDBField.FieldTypes.cftNumeric, pAmount)
      End With
      mvEnv.Connection.InsertRecord("service_booking_revenue", vInsertFields)
    End Sub

    Public Sub Create(ByVal pBookingContact As Integer, ByVal pBookingAddress As Integer, ByVal pServiceContact As Integer, ByVal pRelatedContact As Integer, ByVal pStartDate As String, ByVal pEndDate As String, ByVal pBookingStatus As ServiceBookingStatuses, ByVal pSalesContact As Integer, ByVal pAmount As Double, ByVal pVATRate As String, ByVal pPercentage As Double, ByVal pTransactionDate As String, ByVal pServiceCredits As Boolean, ByVal pAppointmentDesc As String, Optional ByVal pOfferActivity As String = "", Optional ByVal pOfferActivityValue As String = "", Optional ByVal pOfferPayeeType As Contact.ContactTypes = Contact.ContactTypes.ctcContact, Optional ByVal pTransSource As String = "", Optional ByVal pNewQuantity As String = "")
      Dim vAmount As Double
      Dim vVatAmount As Double
      Dim vContactAppointment As ContactAppointment

      With mvClassFields
        .Item(ServiceBookingFields.sbfBookingContactNumber).IntegerValue = pBookingContact
        .Item(ServiceBookingFields.sbfBookingAddressNumber).IntegerValue = pBookingAddress
        .Item(ServiceBookingFields.sbfServiceContactNumber).IntegerValue = pServiceContact
        If pRelatedContact > 0 Then .Item(ServiceBookingFields.sbfRelatedContactNumber).IntegerValue = pRelatedContact
        .Item(ServiceBookingFields.sbfStartDate).Value = pStartDate
        .Item(ServiceBookingFields.sbfEndDate).Value = pEndDate
        .Item(ServiceBookingFields.sbfBookingStatus).Value = mvEnv.GetServiceBookingStatusCode(pBookingStatus)
        If pSalesContact > 0 Then .Item(ServiceBookingFields.sbfSalesContactNumber).IntegerValue = pSalesContact
        vAmount = pAmount
        vVatAmount = Int(((vAmount - (vAmount / (1 + pPercentage / 100))) * 100) + 0.5) / 100
        If pServiceCredits Then
          vAmount = -vAmount
          vVatAmount = -vVatAmount
        End If
        .Item(ServiceBookingFields.sbfAmount).Value = CStr(vAmount)
        .Item(ServiceBookingFields.sbfVatAmount).Value = CStr(vVatAmount)
        .Item(ServiceBookingFields.sbfVatRate).Value = pVATRate
        If Len(pTransactionDate) = 0 Then pTransactionDate = TodaysDate()
        .Item(ServiceBookingFields.sbfTransactionDate).Value = pTransactionDate
      End With
      'create the service booking
      Save()
      'create the service booking revenue record(s)
      CreateRevenue()
      'create the appointment for the service contact
      If Not pServiceCredits Then
        vContactAppointment = New ContactAppointment(mvEnv)
        vContactAppointment.Init()
        vContactAppointment.Create(pServiceContact, pStartDate, pEndDate, ContactAppointment.ContactAppointmentTypes.catServiceBooking, pAppointmentDesc, ServiceBookingNumber, ContactAppointment.ContactAppointmentTimeStatuses.catsBusy)
        vContactAppointment.Save()
      End If
      'create the product offer activity
      If Len(pOfferActivity) > 0 And Len(pOfferActivityValue) > 0 Then
        Dim vCC As New ContactCategory(mvEnv)
        vCC.ContactTypeSaveActivity(pOfferPayeeType, BookingContactNumber, pOfferActivity, pOfferActivityValue, pTransSource, "", "", pNewQuantity, ContactCategory.ActivityEntryStyles.aesNormal)
      End If
    End Sub

    Public Function ValidateDateDetails(ByVal pServiceContactNumber As Integer, ByVal pContactGroup As String, ByVal pStartDate As String, ByVal pEndDate As String, ByVal pRequiresStartDays As Boolean) As MsgBoxResult
      Dim vParams As New CDBParameters
      Dim vDS As New VBDataSelection
      Dim vDT As New CDBDataTable
      Dim vDR As CDBDataRow
      Dim vBookingDuration As Integer
      Dim vBookingStartDay As Integer
      Dim vFound As MsgBoxResult
      Dim vStartDay As Integer

      vFound = MsgBoxResult.No
      With vParams
        .Add("ContactNumber", pServiceContactNumber)
        .Add("ContactGroup", CDBField.FieldTypes.cftCharacter, pContactGroup)
        .Add("StartDate", CDBField.FieldTypes.cftDate, pStartDate)
        .Add("EndDate", CDBField.FieldTypes.cftDate, pEndDate)
      End With
      vDS.Init(mvEnv, DataSelection.DataSelectionTypes.dstServiceStartDays, vParams)
      vDT = vDS.DataTable
      If vDT.Rows.Count() > 0 Then
        vBookingDuration = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pStartDate).Date, CDate(pEndDate).Date))
        vBookingStartDay = Weekday(CDate(pStartDate))
        For Each vDR In vDT.Rows
          If vBookingDuration = Val(vDR.Item("DurationDays")) Then
            Select Case vDR.Item("StartDay")
              Case "M"
                vStartDay = FirstDayOfWeek.Monday
              Case "T"
                vStartDay = FirstDayOfWeek.Tuesday
              Case "W"
                vStartDay = FirstDayOfWeek.Wednesday
              Case "H"
                vStartDay = FirstDayOfWeek.Thursday
              Case "F"
                vStartDay = FirstDayOfWeek.Friday
              Case "S"
                vStartDay = FirstDayOfWeek.Saturday
              Case "U"
                vStartDay = FirstDayOfWeek.Sunday
            End Select
            If vStartDay = vBookingStartDay Then
              vFound = MsgBoxResult.Yes
              Exit For
            End If
          End If
        Next vDR
      ElseIf pRequiresStartDays Then
        'No Start Days data has been set up for either the service contact or the service control, and the service control indicates that Start Days are required, so prevent the booking
        vFound = MsgBoxResult.Abort
      Else
        'No Start Days data has been set up for either the service contact or the service control, and the service control doesn't indicate that Start Days are required, so must therefore assume that the start day and duration are valid
        vFound = MsgBoxResult.Yes
      End If
      ValidateDateDetails = vFound
    End Function

    Public Function ValidateServiceControlRestrictions(ByVal pServiceContactNumber As Integer, ByVal pStartDate As String, ByVal pEndDate As String) As Boolean
      ' BR 11756
      Dim vParams As New CDBParameters
      With vParams
        .Add("ContactNumber", pServiceContactNumber)
        .Add("ValidFrom", CDBField.FieldTypes.cftDate, pStartDate)
      End With
      Dim vDS As New DataSelection(mvEnv, DataSelection.DataSelectionTypes.dstServiceControlRestrictions, vParams, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuWEBServices)
      Dim vDT As CDBDataTable = vDS.DataTable
      Dim vValidateRestrictions As Boolean = True 'Do not apply the restriction unless certain conditions are met. See below
      If vDT.Rows.Count > 0 Then
        Dim vBookingDuration As Integer = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pStartDate), CDate(pEndDate)))
        For Each vDR As CDBDataRow In vDT.Rows
          ' First see if the booking start date is in the valid from/to range
          If (CDate(pStartDate) >= CDate(vDR.Item("ValidFrom"))) And (CDate(pStartDate) <= CDate(vDR.Item("ValidTo"))) Then
            ' Dates are OK, lets see if the booking duration value is in range
            If vBookingDuration <= IntegerValue(vDR.Item("ShortStayDuration")) Then
              ' Duration is OK, now check the late booking days
              If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(TodaysDate()), CDate(pStartDate)) <= IntegerValue(vDR.Item("LateBookingDays")) Then
                'Do not apply the restriction and proceed as normal.
                vValidateRestrictions = True
                Exit For
              Else
                'Only apply restriction when
                '1. Booking Start Date is within valid from/to range
                '2. Booking Duration value is within the defined Short Stay Duration range
                '3. The number of days between the booking date and the start date is more than the late booking days
                vValidateRestrictions = False
              End If
            End If
          End If
        Next vDR
      End If
      Return vValidateRestrictions
    End Function

    Public Sub AddLinkedTransaction(ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pLinkedAnalysis As Collection = Nothing, Optional ByVal pSCFinAdjustmentLines As Collection = Nothing, Optional ByVal pLineNumber As Integer = 0, Optional ByVal pLastLineNumber As Integer = 0)
      Dim vSQL As String
      Dim vBTA As BatchTransactionAnalysis
      Dim vIndex As Integer
      Dim vTransStarted As Boolean
      Dim vBTALine() As String
      Dim vSB As ServiceBooking

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTransStarted = True
      End If
      If pSCFinAdjustmentLines Is Nothing Then
        If pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
          'Used in creating new transactions only
          vSQL = "INSERT INTO service_booking_transactions SELECT "
          vSQL = vSQL & ServiceBookingNumber & ", " & BatchNumber & ", " & TransactionNumber & ","
          vSQL = vSQL & " line_number, '" & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate()))
          vSQL = vSQL & " FROM batch_transaction_analysis WHERE batch_number = " & BatchNumber
          vSQL = vSQL & " AND transaction_number = " & TransactionNumber
          If pLineNumber > 0 And pLastLineNumber = 0 Then
            vSQL = vSQL & " AND line_number = " & pLineNumber
          ElseIf pLineNumber > 0 And pLastLineNumber > 0 Then
            vSQL = vSQL & " AND line_number > " & pLineNumber & " AND line_number <= " & pLastLineNumber
          ElseIf pLastLineNumber > 0 Then
            vSQL = vSQL & " AND line_number > " & LineNumber & " AND line_number <= " & pLastLineNumber
          Else
            vSQL = vSQL & " AND line_number = " & LineNumber
          End If
          mvEnv.Connection.ExecuteSQL(vSQL)
        Else
          If Not pLinkedAnalysis Is Nothing Then
            'Used in FinancialAdjustment only
            For Each vBTA In pLinkedAnalysis
              If ServiceBookingNumber <> vBTA.LinkedBookingNo Then
                vSB = New ServiceBooking
                vSB.Init(mvEnv, (vBTA.LinkedBookingNo))
              Else
                vSB = Me
              End If
              vSQL = "INSERT INTO service_booking_transactions VALUES (" & vSB.ServiceBookingNumber & ", "
              vSQL = vSQL & vBTA.BatchNumber & ", " & vBTA.TransactionNumber & "," & vBTA.LineNumber & ", '"
              vSQL = vSQL & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
              mvEnv.Connection.ExecuteSQL(vSQL)
            Next vBTA
          End If
        End If
      Else
        'Coming from Smart Client Trader- used in FinancialHistory.ProcessAdjustment
        For vIndex = 1 To pSCFinAdjustmentLines.Count()
          vBTALine = Split(CStr(pSCFinAdjustmentLines.Item(vIndex)), ",")
          vSQL = "INSERT INTO service_booking_transactions VALUES ("
          If UBound(vBTALine) > 3 Then
            vSQL = vSQL & vBTALine(0) & ", " & vBTALine(1) & ", " & vBTALine(2) & ", " & vBTALine(3) & "," & vBTALine(4) & ", '"
          Else
            vSQL = vSQL & ServiceBookingNumber & ", " & vBTALine(0) & ", " & vBTALine(1) & "," & vBTALine(2) & ", '"
          End If

          vSQL = vSQL & mvEnv.User.Logname & "', " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
          mvEnv.Connection.ExecuteSQL(vSQL)
        Next
      End If
      If vTransStarted Then mvEnv.Connection.CommitTransaction()
    End Sub
  End Class
End Namespace
