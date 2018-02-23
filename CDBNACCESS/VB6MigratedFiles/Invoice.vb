

Namespace Access
  Partial Public Class Invoice

    Public Enum InvoiceRecordSetTypes 'These are bit values
      irtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum InvoiceFields
      ifAll = 0
      ifContactNumber
      ifAddressNumber
      ifCompany
      ifBatchNumber
      ifTransactionNumber
      ifInvoiceNumber
      ifInvoiceDate
      ifPaymentDue
      ifSalesLedgerBatch
      ifSalesLedgerAccount
      ifInvoicePayStatus
      ifInvoiceDisputeCode
      ifAmountPaid
      ifRecordType
      ifReprintCount
      ifDepositAmount
      ifPrintJobNumber
      ifPrintInvoice
      ifProvisionalInvoiceNumber
      ifAdjustmentStatus
    End Enum

    'variables and structures for the production of invoices
    Enum InvPayStatuses
      InvNotPaid = 0
      InvPartPaid
      InvFullyPaid
      InvPendingDDPayment
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Public Structure InvoicePayment
      Dim InvoiceNumberUsed As Integer
      Dim AmountPaid As Double
      Dim RecordType As String
      Dim ContactNumber As Integer
      Dim AddressNumber As Integer
      Dim SalesLedgerAccount As String
    End Structure

    Private mvNowPaid As Double
    Private mvAmountUsed As Double
    Private mvInvoiceAmount As Double
    Private mvInvoicePayments As Collection
    Private mvIsSundryCreditNote As Nullable(Of Boolean)
    Private mvFinancialHistory As FinancialHistory

    Private Property InvoiceAmountCalculated As Boolean = False

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "invoices"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("company")
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("invoice_number", CDBField.FieldTypes.cftLong)
          .Add("invoice_date", CDBField.FieldTypes.cftDate)
          .Add("payment_due", CDBField.FieldTypes.cftDate)
          .Add("sales_ledger_batch")
          .Add("sales_ledger_account")
          .Add("invoice_pay_status")
          .Add("invoice_dispute_code")
          .Add("amount_paid", CDBField.FieldTypes.cftNumeric)
          .Add("record_type")
          .Add("reprint_count", CDBField.FieldTypes.cftInteger)
          .Add("deposit_amount", CDBField.FieldTypes.cftNumeric)
          .Add("print_job_number", CDBField.FieldTypes.cftLong)
          .Add("print_invoice")
          .Add("provisional_invoice_number", CDBField.FieldTypes.cftLong)
          .Add("adjustment_status").InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbInvoiceAdjustmentStatus)
        End With

        mvClassFields.Item(InvoiceFields.ifDepositAmount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets)
        mvClassFields.Item(InvoiceFields.ifPrintJobNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPrintJobNumber)
        mvClassFields.Item(InvoiceFields.ifPrintInvoice).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix)
        mvClassFields.Item(InvoiceFields.ifProvisionalInvoiceNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber)

        mvClassFields.Item(InvoiceFields.ifContactNumber).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifAddressNumber).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifBatchNumber).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifTransactionNumber).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifProvisionalInvoiceNumber).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifCompany).PrefixRequired = True
        mvClassFields.Item(InvoiceFields.ifSalesLedgerAccount).PrefixRequired = True

        mvClassFields.Item(InvoiceFields.ifInvoiceNumber).PrimaryKey = True
        mvClassFields.Item(InvoiceFields.ifBatchNumber).PrimaryKey = True
        mvClassFields.Item(InvoiceFields.ifTransactionNumber).PrimaryKey = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvIsSundryCreditNote = Nothing
      InvoiceAmountCalculated = False
      mvFinancialHistory = Nothing
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(InvoiceFields.ifPrintInvoice).Value = "Y"
    End Sub

    Private Sub SetValid(ByVal pField As InvoiceFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    ''' <summary>Calculate the Invoice Amount using Invoice Details and Batch Transaction Analysis</summary>
    ''' <remarks>This is only run if the Invoice Amount has not previously been set</remarks>
    Private Sub SetInvoiceAmount()
      If InvoiceAmountCalculated = False Then
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("batch_transaction_analysis bta", "id.batch_number", "bta.batch_number", "id.transaction_number", "bta.transaction_number", "id.line_number", "bta.line_number")
        Dim vWhereFields As New CDBFields()
        If mvClassFields.Item(InvoiceFields.ifInvoiceNumber).LongValue > 0 Then
          vWhereFields.Add("id.invoice_number", mvClassFields.Item(InvoiceFields.ifInvoiceNumber).IntegerValue)
        Else
          vWhereFields.Add("id.batch_number", BatchNumber)
          vWhereFields.Add("id.transaction_number", TransactionNumber)
        End If

        Dim vSQL As New SQLStatement(mvEnv.Connection, "SUM(bta.amount) AS invoice_amount", "invoice_details id", vWhereFields, "", vAnsiJoins)
        mvInvoiceAmount = DoubleValue(vSQL.GetValue)

        If mvInvoiceAmount = 0 Then
          'See if this has come from a Financial Adjustment, and if so set InvoiceAmount to be the sum of the positive amounts
          vAnsiJoins.Add("batches b", "bta.batch_number", "b.batch_number")
          vWhereFields.Add("bta.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
          vWhereFields.Add("batch_type", Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment))

          vSQL = New SQLStatement(mvEnv.Connection, "SUM(bta.amount) AS invoice_amount", "invoice_details id", vWhereFields, "", vAnsiJoins)
          mvInvoiceAmount = DoubleValue(vSQL.GetValue)
        End If
        InvoiceAmountCalculated = True
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As InvoiceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = InvoiceRecordSetTypes.irtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "i")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pInvoiceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      InitClassFields()
      If pBatchNumber > 0 Or pInvoiceNumber > 0 Then
        If pInvoiceNumber > 0 Then vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftLong, pInvoiceNumber)
        If pBatchNumber > 0 Then vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
        If pTransactionNumber > 0 Then vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(InvoiceRecordSetTypes.irtAll) & " FROM invoices i WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, InvoiceRecordSetTypes.irtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As InvoiceRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And InvoiceRecordSetTypes.irtAll) = InvoiceRecordSetTypes.irtAll Then
          .SetItem(InvoiceFields.ifContactNumber, vFields)
          .SetItem(InvoiceFields.ifAddressNumber, vFields)
          .SetItem(InvoiceFields.ifCompany, vFields)
          .SetItem(InvoiceFields.ifBatchNumber, vFields)
          .SetItem(InvoiceFields.ifTransactionNumber, vFields)
          .SetItem(InvoiceFields.ifInvoiceNumber, vFields)
          .SetItem(InvoiceFields.ifInvoiceDate, vFields)
          .SetItem(InvoiceFields.ifPaymentDue, vFields)
          .SetItem(InvoiceFields.ifSalesLedgerBatch, vFields)
          .SetItem(InvoiceFields.ifSalesLedgerAccount, vFields)
          .SetItem(InvoiceFields.ifInvoicePayStatus, vFields)
          .SetItem(InvoiceFields.ifInvoiceDisputeCode, vFields)
          .SetItem(InvoiceFields.ifAmountPaid, vFields)
          .SetItem(InvoiceFields.ifRecordType, vFields)
          .SetItem(InvoiceFields.ifReprintCount, vFields)
          .SetOptionalItem(InvoiceFields.ifDepositAmount, vFields)
          .SetOptionalItem(InvoiceFields.ifPrintJobNumber, vFields)
          .SetOptionalItem(InvoiceFields.ifPrintInvoice, vFields)
          .SetOptionalItem(InvoiceFields.ifProvisionalInvoiceNumber, vFields)
          .SetOptionalItem(InvoiceFields.ifAdjustmentStatus, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Init(pEnv)
      mvClassFields(InvoiceFields.ifBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(InvoiceFields.ifTransactionNumber).IntegerValue = pTransactionNumber
    End Sub

    Public Sub Create(ByVal pRecordType As Invoice.InvoiceRecordType, ByVal pParameterList As CDBParameters)
      Dim vCheckProvisionalInvoiceNumber As Boolean = False
      'Set Invoices RecordType
      If pParameterList.ContainsKey("RecordType") = False Then pParameterList.Add("RecordType", CDBField.FieldTypes.cftCharacter)
      pParameterList("RecordType").Value = Invoice.GetRecordTypeCode(pRecordType)
      'If we have InvoiceNumber or ProvisionalInvoiceNumber parametera then only keep them if they are greater than zero
      If pParameterList.ContainsKey("InvoiceNumber") AndAlso pParameterList("InvoiceNumber").IntegerValue <= 0 Then pParameterList.Remove("InvoiceNumber")
      If pParameterList.ContainsKey("ProvisionalInvoiceNumber") AndAlso pParameterList("ProvisionalInvoiceNumber").IntegerValue <= 0 Then pParameterList.Remove("ProvisionalInvoiceNumber")
      If pRecordType = InvoiceRecordType.SalesLedgerCash AndAlso Not pParameterList.ContainsKey("InvoiceNumber") Then
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber) Then
          'Jira 380: For Sales Ledger Cash (record type 'C') Invoices use Provisional Invoice Number Control Number
          pParameterList.Add("InvoiceNumber", mvEnv.GetControlNumber("PR"))
          vCheckProvisionalInvoiceNumber = True
        Else
          pParameterList.Add("InvoiceNumber", mvEnv.GetControlNumber("I"))
        End If
      End If
      'CheckInvoiceNumber(vProvisionalInvoiceNumber)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = True Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next
      'BR15581: For Credit Notes where Invoice Number is not being set i.e. from Batch.WriteInvoice, do not require Invoice Number in SetAmountPaid
      Update(pParameterList, pRecordType = InvoiceRecordType.CreditNote AndAlso Not pParameterList.ContainsKey("InvoiceNumber"))
      CheckInvoiceNumber(vCheckProvisionalInvoiceNumber)
    End Sub

    Public Sub Update(ByVal pInvoiceNumber As Integer, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pAmountPaid As Double, ByVal pReprintCount As Integer, ByVal pInvoiceDate As String, ByVal pPaymentDue As String, ByVal pInvoicePayStatus As String, ByVal pRecordType As String, Optional ByVal pDepositAmount As Double = 0)
      If pInvoiceNumber > 0 Then mvClassFields.Item(InvoiceFields.ifInvoiceNumber).Value = CStr(pInvoiceNumber)
      mvClassFields.Item(InvoiceFields.ifContactNumber).IntegerValue = pContactNumber
      mvClassFields.Item(InvoiceFields.ifAddressNumber).IntegerValue = pAddressNumber
      mvClassFields.Item(InvoiceFields.ifCompany).Value = pCompany
      mvClassFields.Item(InvoiceFields.ifSalesLedgerAccount).Value = pSalesLedgerAccount
      mvClassFields.Item(InvoiceFields.ifAmountPaid).DoubleValue = pAmountPaid
      mvClassFields.Item(InvoiceFields.ifReprintCount).IntegerValue = pReprintCount
      mvClassFields.Item(InvoiceFields.ifInvoiceDate).Value = pInvoiceDate
      mvClassFields.Item(InvoiceFields.ifPaymentDue).Value = pPaymentDue
      mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = pInvoicePayStatus
      mvClassFields.Item(InvoiceFields.ifRecordType).Value = pRecordType
      If pDepositAmount > 0 Then mvClassFields.Item(InvoiceFields.ifDepositAmount).DoubleValue = pDepositAmount
    End Sub

    Public Sub Update(ByVal pParameterList As CDBParameters, ByVal pSetAmountPaidInvNoNotRequired As Boolean)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next

      'Set AmountPaid and InvoicePayStatus
      SetAmountPaid(0, False, pSetAmountPaidInvNoNotRequired)

    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(InvoiceFields.ifAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Function CalcInvPayDue(ByVal pTermsFrom As String, ByVal pTermsPeriod As String, ByVal pTermsNumber As Integer, ByVal pBatchNo As Integer, ByVal pTransNo As Integer, ByVal pInvoiceDate As Date, ByRef pPayDue As Date) As Boolean
      Dim vDate As Date
      Dim vRecordSet As CDBRecordSet

      If pTermsFrom = "I" Then
        pPayDue = DateAdd(LCase(pTermsPeriod), pTermsNumber, pInvoiceDate)
        CalcInvPayDue = True
      Else
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT start_date FROM batch_transactions bt, calendar c WHERE bt.batch_number = " & pBatchNo & " AND bt.transaction_number = " & pTransNo & " AND c.start_date > bt.transaction_date ORDER BY c.start_date")
        If vRecordSet.Fetch() = True Then
          vDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vRecordSet.Fields(1).Value))
          pPayDue = DateAdd(LCase(pTermsPeriod), pTermsNumber, vDate)
          CalcInvPayDue = True
        Else
          CalcInvPayDue = False
          vRecordSet.CloseRecordSet()
          RaiseError(DataAccessErrors.daeInvCalendarNotFound)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Function

    Public Function GetInvPayStatus(ByVal pPayStatus As InvPayStatuses, ByRef pStatus As String) As Boolean
      Dim vAttr As String = ""
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      Select Case pPayStatus
        Case InvPayStatuses.InvNotPaid
          vAttr = "not_paid"
        Case InvPayStatuses.InvPartPaid
          vAttr = "part_paid"
        Case InvPayStatuses.InvFullyPaid
          vAttr = "fully_paid"
        Case InvPayStatuses.InvPendingDDPayment
          vAttr = "pending_dd_payment"
      End Select

      vSQL = "SELECT invoice_pay_status FROM invoice_pay_statuses WHERE " & vAttr & " = 'Y'"
      vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      If vRecordSet.Fetch() = True Then
        pStatus = vRecordSet.Fields(1).Value
        GetInvPayStatus = True
      Else
        GetInvPayStatus = False
        RaiseError(DataAccessErrors.daeInvPayStatusNotFound, CapitaliseWords(ReplaceString(vAttr, "_", " ")))
      End If
      vRecordSet.CloseRecordSet()
    End Function

    Public Sub SetInvoiceNumber(Optional ByVal pUpdateInvoiceDetails As Boolean = False, Optional ByVal pSetProvisionalInvoiceNumber As Boolean = False)
      Dim vWhereFields As CDBFields
      Dim vUpdateFields As CDBFields
      Dim vContinue As Boolean
      If mvExisting = True And mvClassFields.Item(InvoiceFields.ifInvoiceNumber).IntegerValue = 0 Then
        vContinue = True
        If pSetProvisionalInvoiceNumber Then
          If mvClassFields.Item(InvoiceFields.ifProvisionalInvoiceNumber).InDatabase Then
            mvClassFields.Item(InvoiceFields.ifInvoiceNumber).Value = mvEnv.GetControlNumber("PR").ToString
            mvClassFields.Item(InvoiceFields.ifProvisionalInvoiceNumber).Value = InvoiceNumber
          Else
            vContinue = False
          End If
        Else
          mvClassFields.Item(InvoiceFields.ifInvoiceNumber).Value = mvEnv.GetControlNumber("I").ToString
        End If
        If vContinue Then
          CheckInvoiceNumber(pSetProvisionalInvoiceNumber)
          If pUpdateInvoiceDetails Then
            vUpdateFields = New CDBFields
            vWhereFields = New CDBFields
            vUpdateFields.Add("invoice_number", CDBField.FieldTypes.cftLong, InvoiceNumber)
            vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
            mvEnv.Connection.UpdateRecords("invoice_details", vUpdateFields, vWhereFields, False)
          End If
        End If
      End If
    End Sub

    ''' <summary>Remove the allocations from either (1) an invoice / credit Note, or (2) an invoice payment</summary>
    ''' <returns>Batch / Transaction reference of the adjustment transaction</returns>
    Public Function RemoveAllocations() As String
      Dim vCash As Boolean = False
      Dim vIsInvoice As Boolean = False
      Select Case RecordType
        Case "C"
          vCash = True
        Case "I"
          vIsInvoice = True
      End Select

      Dim vWhereFields As New CDBFields()
      Dim vOrderBy As String

      Dim vAttrs As String = "iph.amount, i.invoice_number, i.batch_number, i.transaction_number, i.amount_paid, %1.amount AS %1_amount, %1.transaction_type"
      Dim vTable As String = "invoice_payment_history iph"
      Dim vAnsiJoins As New AnsiJoins()

      If vIsInvoice Then
        'Check no unposted allocations first
        Dim vContainsUnpostedAllocations As Boolean
        AllocationsAmount(False, True, vContainsUnpostedAllocations)
        If vContainsUnpostedAllocations Then RaiseError(DataAccessErrors.daeInvoiceAllocationsUnposted)

        'Check if any of the allocations are from credit notes which do not have a number
        vAnsiJoins.Add("invoices i", "iph.batch_number", "i.batch_number", "iph.transaction_number", "i.transaction_number")
        vWhereFields.Add("iph.invoice_number", IntegerValue(InvoiceNumber))
        vWhereFields.Add("i.invoice_number", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoEqual)
        Dim vCountSQL As New SQLStatement(mvEnv.Connection, "", "invoice_payment_history iph", vWhereFields, "", vAnsiJoins)
        If mvEnv.Connection.GetCountFromStatement(vCountSQL) > 0 Then RaiseError(DataAccessErrors.daeCreditNoteAllocNoNumber)
        vWhereFields.Clear()
        vAnsiJoins.Clear()

        'Remove allocations against this invoice
        vAttrs = vAttrs.Replace("%1", "fh")
        With vAnsiJoins
          .Add("invoices i", "i.invoice_number", "(select MIN(id.invoice_number) from invoice_details id where iph.batch_number = id.batch_number AND iph.transaction_number = id.transaction_number)")
          .Add("financial_history fh", "i.batch_number", "fh.batch_number", "i.transaction_number", "fh.transaction_number")
        End With

        'Add Join to sundry credit note allocations as the amount could have the wrong sign (negative instead of positive)!!!
        Dim vNestedAnsiJoins As New AnsiJoins()
        vNestedAnsiJoins.Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bta.transaction_number")
        vNestedAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")

        Dim vNestedWhereFields As New CDBFields(New CDBField("bta.line_type", "K"))
        vNestedWhereFields.Add("tt.transaction_sign", "C")
        vNestedWhereFields.Add("bta.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)

        Dim vNestedSQL = New SQLStatement(mvEnv.Connection, "bta.batch_number, bta.transaction_number, bta.line_number, bta.amount", "batch_transaction_analysis bta", vNestedWhereFields, "", vNestedAnsiJoins)

        vAnsiJoins.AddLeftOuterJoin("(" & vNestedSQL.SQL & ") bta", "iph.allocation_batch_number", "bta.batch_number", "iph.allocation_transaction_number", "bta.transaction_number", "iph.allocation_line_number", "bta.line_number")

        With vWhereFields
          .Add("iph.invoice_number", IntegerValue(InvoiceNumber))
          .Add("iph.status", CDBField.FieldTypes.cftCharacter)
          .Add("iph.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOpenBracket)
          .Add("iph.amount#2", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
          .Add("bta.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
        End With
      Else
        'Remove allocations made w/ this payment/credit note
        vAttrs = vAttrs.Replace("%1", "bt")
        With vAnsiJoins
          .Add("invoices i", "iph.invoice_number", "i.invoice_number")
          .Add("batch_transactions bt", "i.batch_number", "bt.batch_number", "i.transaction_number", "bt.transaction_number")
          .Add("invoice_details id", "iph.batch_number", "id.batch_number", "iph.transaction_number", "id.transaction_number", "iph.line_number", "id.line_number")
        End With
        With vWhereFields
          .Add("iph.batch_number", BatchNumber)
          .Add("iph.transaction_number", TransactionNumber)
          .Add("iph.status", CDBField.FieldTypes.cftCharacter)
          .Add("iph.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
        End With
      End If

      CheckAllocationsForRemoval(vIsInvoice)

      vOrderBy = "i.batch_number, i.transaction_number"

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, vTable, vWhereFields, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet()

      'Set-up WhereFields & UpdateFields
      vWhereFields.Clear()
      vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftInteger)
      Dim vUpdateFields As New CDBFields()
      With vUpdateFields
        .Add("invoice_pay_status", CDBField.FieldTypes.cftCharacter)
        .Add("amount_paid", CDBField.FieldTypes.cftNumeric)
      End With

      Dim vAmountPaid As Double
      Dim vBookingsRS As CDBRecordSet
      Dim vBookingsSQL As SQLStatement
      Dim vBookingWhereFields As New CDBFields
      Dim vBookingUpdateFields As New CDBFields(New CDBField("booking_status", CDBField.FieldTypes.cftCharacter))
      Dim vEBBatchNumber As Integer
      Dim vEBTransactionNumber As Integer
      Dim vEBInvoiceNumber As Integer
      Dim vInvoicesRS As CDBRecordSet
      Dim vInvoicesSQL As SQLStatement
      Dim vPayStatus As String = ""
      Dim vTransactionType As String = ""
      Dim vInvoiceAmount As Double
      Dim vLastBatchNumber As Integer = -1
      Dim vLastTransactionNumber As Integer = -1
      While vRS.Fetch = True
        With vRS
          'For each record update the invoices record by resetting amount_paid and invoice_pay_status
          vWhereFields("invoice_number").Value = .Fields("invoice_number").Value
          vTransactionType = .Fields("transaction_type").Value
          vInvoiceAmount = If(vIsInvoice, .Fields("fh_amount").DoubleValue, .Fields("bt_amount").DoubleValue)

          If vIsInvoice And vTransactionType = "A" And vInvoiceAmount <= 0 Then
            'For transaction_type 'A' (i.e. Adjustments), vAmountPaid will be 0 if fully paid
            ' To determine if the PayStutus will be part paid or not paid, the initial Adjustment amount is required. This
            ' is held on the first financial_history_details for transaction.
            Dim vPaymentBatchNumber As Integer
            Dim vPaymentTransactionNumber As Integer
            Dim vRecordSet As CDBRecordSet
            Dim vAdjustmentAmount As Double = 0
            vPaymentBatchNumber = .Fields("batch_number").IntegerValue
            vPaymentTransactionNumber = .Fields("transaction_number").IntegerValue
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT fhd.amount FROM financial_history_details fhd WHERE fhd.batch_number = " & vPaymentBatchNumber & " AND fhd.transaction_number = " & vPaymentTransactionNumber & " AND fhd.line_number = (SELECT MIN(fhd2.line_number) FROM financial_history_details fhd2 WHERE fhd2.batch_number = " & vPaymentBatchNumber & " AND fhd2.transaction_number = " & vPaymentTransactionNumber & ")")
            If vRecordSet.Fetch() = True Then
              vAdjustmentAmount = vRecordSet.Fields(1).IntegerValue
            End If
            vRecordSet.CloseRecordSet()
            vAmountPaid = .Fields("amount_paid").DoubleValue - .Fields("amount").DoubleValue
            If vAmountPaid = vAdjustmentAmount Then
              vPayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
            Else
              vPayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)
            End If
          Else
            If vLastBatchNumber = .Fields("batch_number").IntegerValue _
              And vLastTransactionNumber = .Fields("transaction_number").IntegerValue Then
              vAmountPaid = vAmountPaid - Math.Abs(.Fields("amount").DoubleValue)
            Else
              vAmountPaid = .Fields("amount_paid").DoubleValue - Math.Abs(.Fields("amount").DoubleValue)
            End If
            If vAmountPaid < 0 Then vAmountPaid = 0 'just in case...
            If vAmountPaid = 0 Then
              vPayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
            Else
              vPayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)
            End If
          End If
          vUpdateFields("invoice_pay_status").Value = vPayStatus
          vUpdateFields("amount_paid").Value = vAmountPaid.ToString
          If vCash = False Then
            mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields)
          End If
          'BR 8818 Update Booking Status on any associated Event/Accommodation Bookings for original Invoice
          'Pass 1: Update Event Bookings
          'Pass 2: Update Accommodation Bookings
          '- Booked (Paid) => Booked (Invoiced)
          '- Booked (Paid) Transfer => Booked (Invoiced) Transfer
          '- Waiting (Paid) => Waiting (Invoiced)
          vBookingWhereFields.Clear()
          vEBBatchNumber = 0
          vEBTransactionNumber = 0
          vEBInvoiceNumber = IntegerValue(IIf(vIsInvoice, IntegerValue(InvoiceNumber), .Fields("invoice_number").IntegerValue).ToString)
          vBookingWhereFields.Add("invoice_number", vEBInvoiceNumber)
          vInvoicesSQL = New SQLStatement(mvEnv.Connection, "batch_number,transaction_number", "invoices", vBookingWhereFields)
          vInvoicesRS = vInvoicesSQL.GetRecordSet()
          If vInvoicesRS.Fetch Then
            vEBBatchNumber = vInvoicesRS.Fields("batch_number").IntegerValue
            vEBTransactionNumber = vInvoicesRS.Fields("transaction_number").IntegerValue
            vBookingWhereFields.Clear()
            vBookingWhereFields.Add("batch_number", vEBBatchNumber)
            vBookingWhereFields.Add("transaction_number", vEBTransactionNumber)
            For vPass As Integer = 1 To 2
              Select Case vPass
                Case 1
                  vTable = "event_bookings"
                Case 2
                  vTable = "contact_room_bookings"
              End Select
              vBookingsSQL = New SQLStatement(mvEnv.Connection, "line_number,booking_status", vTable, vBookingWhereFields)
              vBookingsRS = vBookingsSQL.GetRecordSet()
              While vBookingsRS.Fetch = True
                Dim vNewStatusCode As String = ""
                Select Case vBookingsRS.Fields("booking_status").Value
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaid)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced)
                End Select
                If vNewStatusCode.Length > 0 Then
                  If Not vBookingWhereFields.ContainsKey("line_number") Then vBookingWhereFields.Add("line_number", CDBField.FieldTypes.cftInteger)
                  vBookingWhereFields("line_number").Value = vBookingsRS.Fields("line_number").Value
                  vBookingUpdateFields("booking_status").Value = vNewStatusCode
                  mvEnv.Connection.UpdateRecords(vTable, vBookingUpdateFields, vBookingWhereFields)
                End If
              End While
              vBookingsRS.CloseRecordSet()
            Next
          End If
          vInvoicesRS.CloseRecordSet()
          vLastBatchNumber = .Fields("batch_number").IntegerValue
          vLastTransactionNumber = .Fields("transaction_number").IntegerValue
        End With
      End While
      vRS.CloseRecordSet()

      Dim vReference As String = ProcessInvoicePaymentHistory(vIsInvoice)

      'Reset the invoice to be unpaid/unallocated
      If ((Invoice.GetRecordType(RecordType) = InvoiceRecordType.Invoice) _
      OrElse (Invoice.GetRecordType(RecordType) = InvoiceRecordType.CreditNote AndAlso IsSundryCreditNote = False)) Then
        vWhereFields.Clear()
        vWhereFields.Add("invoice_number", IntegerValue(InvoiceNumber))
        vUpdateFields("invoice_pay_status").Value = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
        vUpdateFields("amount_paid").Value = "0"
        mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields)
      End If

      Return vReference
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(InvoiceFields.ifAddressNumber).IntegerValue
      End Get
      Set(pValue As Integer)
        mvClassFields.Item(InvoiceFields.ifAddressNumber).IntegerValue = pValue
      End Set
    End Property

    Public ReadOnly Property AmountPaid() As Double
      Get
        AmountPaid = mvClassFields.Item(InvoiceFields.ifAmountPaid).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(InvoiceFields.ifBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Company() As String
      Get
        Company = mvClassFields.Item(InvoiceFields.ifCompany).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(InvoiceFields.ifContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property InvoiceDate() As String
      Get
        InvoiceDate = mvClassFields.Item(InvoiceFields.ifInvoiceDate).Value
      End Get
    End Property

    Public ReadOnly Property InvoiceDisputeCode() As String
      Get
        InvoiceDisputeCode = mvClassFields.Item(InvoiceFields.ifInvoiceDisputeCode).Value
      End Get
    End Property

    Public ReadOnly Property InvoiceNumber() As String
      Get
        InvoiceNumber = mvClassFields.Item(InvoiceFields.ifInvoiceNumber).Value
      End Get
    End Property

    Public ReadOnly Property InvoicePayStatus() As String
      Get
        InvoicePayStatus = mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value
      End Get
    End Property

    Public Property PaymentDue() As String
      Get
        PaymentDue = mvClassFields.Item(InvoiceFields.ifPaymentDue).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(InvoiceFields.ifPaymentDue).Value = Value
      End Set
    End Property

    Public ReadOnly Property PrintInvoice() As Boolean
      Get
        Dim vPrintInvoice As Boolean = True
        If mvClassFields.Item(InvoiceFields.ifPrintInvoice).Value.Length > 0 Then vPrintInvoice = mvClassFields.Item(InvoiceFields.ifPrintInvoice).Bool 'Null = True
        Return vPrintInvoice
      End Get
    End Property

    Public ReadOnly Property PrintJobNumber() As String
      Get
        PrintJobNumber = mvClassFields.Item(InvoiceFields.ifPrintJobNumber).Value
      End Get
    End Property

    Public ReadOnly Property ProvisionalInvoiceNumber() As Integer
      Get
        ProvisionalInvoiceNumber = mvClassFields.Item(InvoiceFields.ifProvisionalInvoiceNumber).IntegerValue
      End Get
    End Property

    ''' <summary>Returns the record type code of the Invoice.</summary>
    ''' <remarks>If the RecordType enum value is required then please see the InvoiceType property.</remarks>
    Public ReadOnly Property RecordType() As String
      Get
        RecordType = mvClassFields.Item(InvoiceFields.ifRecordType).Value
      End Get
    End Property

    Public ReadOnly Property ReprintCount() As Integer
      Get
        ReprintCount = mvClassFields.Item(InvoiceFields.ifReprintCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SalesLedgerAccount() As String
      Get
        SalesLedgerAccount = mvClassFields.Item(InvoiceFields.ifSalesLedgerAccount).Value
      End Get
    End Property

    Public ReadOnly Property SalesLedgerBatch() As String
      Get
        SalesLedgerBatch = mvClassFields.Item(InvoiceFields.ifSalesLedgerBatch).Value
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(InvoiceFields.ifTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DepositAmount() As String
      Get
        DepositAmount = mvClassFields.Item(InvoiceFields.ifDepositAmount).Value
      End Get
    End Property

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        Select Case pAttributeName
          Case "AmountPaid"
            mvNowPaid = Val(Value)
          Case "AmountUsed"
            mvAmountUsed = Val(Value)
          Case "InvoiceAmount"
            mvInvoiceAmount = Val(Value)
          Case "DisputeCode"
            mvClassFields(InvoiceFields.ifInvoiceDisputeCode).Value = Value
          Case Else
            mvClassFields.ItemValue(pAttributeName) = Value
        End Select
      End Set
    End Property

    Public Property NowPaid() As Double
      Get
        NowPaid = mvNowPaid
      End Get
      Set(ByVal Value As Double)
        mvNowPaid = Value
      End Set
    End Property

    Public Property AmountUsed() As Double
      Get
        AmountUsed = mvAmountUsed
      End Get
      Set(ByVal Value As Double)
        mvAmountUsed = Value
      End Set
    End Property

    Public ReadOnly Property InvoicePayments() As Collection
      Get
        If mvInvoicePayments Is Nothing Then mvInvoicePayments = New Collection
        InvoicePayments = mvInvoicePayments
      End Get
    End Property

    Public Property InvoiceAmount() As Double
      Get
        SetInvoiceAmount()
        Return mvInvoiceAmount
      End Get
      Set(ByVal Value As Double)
        mvInvoiceAmount = Value
        InvoiceAmountCalculated = True    'We have set the invoice amount so no need to calculate it
      End Set
    End Property

    Public ReadOnly Property InvoicePrinted() As Boolean
      Get
        Dim vPrinted As Boolean
        If Val(InvoiceNumber) > 0 Then
          If ReprintCount >= 0 Then vPrinted = True
        End If
        InvoicePrinted = vPrinted
      End Get
    End Property

    ''' <summary>Returns an InvoiceRecordType enum value representing the type of Invoice.</summary>
    ''' <remarks>If the actual RecordType code is required then please see the RecordType property.</remarks>
    Public ReadOnly Property InvoiceType As Invoice.InvoiceRecordType
      Get
        Return Invoice.GetRecordType(mvClassFields.Item(InvoiceFields.ifRecordType).Value)
      End Get
    End Property

    Public ReadOnly Property AdjustmentStatusCode As String
      Get
        Return mvClassFields.Item(InvoiceFields.ifAdjustmentStatus).Value
      End Get
    End Property

    Public ReadOnly Property AdjustmentStatus As InvoiceAdjustmentStatus
      Get
        Return GetAdjustmentStatus(mvClassFields.Item(InvoiceFields.ifAdjustmentStatus).Value)
      End Get
    End Property

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      Select Case pAttributeName
        Case "AmountPaid", "AmountUsed", "InvoiceAmount"
          LineDataType = CDBField.FieldTypes.cftNumeric
        Case "DisputeCode"
          LineDataType = CDBField.FieldTypes.cftCharacter
        Case Else
          LineDataType = mvClassFields.ItemDataType(pAttributeName)
      End Select
    End Function

    Public Sub SCSetPaymentValues(ByVal pInvoiceNumberUsed As Integer, ByVal pAmountPaid As Double)
      SCSetPaymentValues(pInvoiceNumberUsed, pAmountPaid, "", 0, 0, "")
    End Sub

    Public Sub SCSetPaymentValues(ByVal pInvoiceNumberUsed As Integer, ByVal pAmountPaid As Double, ByVal pRecordType As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pSalesledgerAccount As String)
      Dim vPayment As New InvoicePayment
      vPayment.InvoiceNumberUsed = pInvoiceNumberUsed
      vPayment.AmountPaid = pAmountPaid
      If pRecordType.Length > 0 Then vPayment.RecordType = pRecordType
      If pContactNumber > 0 AndAlso pAddressNumber > 0 Then
        vPayment.ContactNumber = pContactNumber
        vPayment.AddressNumber = pAddressNumber
      End If
      vPayment.SalesLedgerAccount = pSalesledgerAccount
      If mvInvoicePayments Is Nothing Then mvInvoicePayments = New Collection
      mvInvoicePayments.Add(vPayment)
    End Sub

    Public Sub ProcessPayment(pAmount As Double)
      mvNowPaid = pAmount
      SetInvoiceAmount()
      SCUpdatePayment()
      Dim vCC As New CreditCustomer
      vCC.Init(mvEnv, 0, Company, SalesLedgerAccount)
      If vCC.Existing Then
        vCC.AdjustOutstanding(pAmount * -1)
        vCC.Save()
      Else
        RaiseError(DataAccessErrors.daeCreditCustomerMissing1, Company, SalesLedgerAccount)
      End If
    End Sub

    Public Sub SCUpdatePayment()
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vPass As Integer
      Dim vTable As String = ""
      Dim vAlias As String = ""
      Dim vNewStatusCode As String
      Dim vUpdateFields As New CDBFields

      With mvClassFields
        If RecordType = "I" Then
          If (FixTwoPlaces(AmountPaid) + FixTwoPlaces(NowPaid)) > FixTwoPlaces(InvoiceAmount) Then
            RaiseError(DataAccessErrors.daeCannotOverPayInvoice)
          End If
          .Item(InvoiceFields.ifAmountPaid).DoubleValue = FixTwoPlaces(.Item(InvoiceFields.ifAmountPaid).DoubleValue) + FixTwoPlaces(NowPaid)
        Else
          .Item(InvoiceFields.ifAmountPaid).DoubleValue = FixTwoPlaces(.Item(InvoiceFields.ifAmountPaid).DoubleValue) + FixTwoPlaces(AmountUsed)
        End If
        If Me.AmountPaid.Equals(0) AndAlso IsFinancialAdjustmentInvoice = True Then
          If Me.FinancialHistory IsNot Nothing AndAlso Me.FinancialHistory.Existing Then
            'If the Invoice was created as a result of a re-analysis specifically set the InvoiceAmount to be zero
            If Me.FinancialHistory.Amount.Equals(0) Then Me.InvoiceAmount = Me.FinancialHistory.Amount
          End If
        End If
        If FixTwoPlaces(System.Math.Abs(mvInvoiceAmount)) = FixTwoPlaces(AmountPaid) Then
          .Item(InvoiceFields.ifInvoicePayStatus).Value = "F"
        ElseIf (FixTwoPlaces(System.Math.Abs(mvInvoiceAmount)) <> FixTwoPlaces(AmountPaid)) And AmountPaid <> 0 Then
          .Item(InvoiceFields.ifInvoicePayStatus).Value = "P"
        Else
          .Item(InvoiceFields.ifInvoicePayStatus).Value = "N"
        End If
      End With

      If (BatchNumber > 0 And TransactionNumber > 0) And ((mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = "F" And mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).SetValue <> "F") Or (mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value <> "F" And mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).SetValue = "F")) Then
        'TA BR 8068 Update Status on any associated Event/Accommodation Bookings to indicate invoice has been printed
        'Pass 1: Update Event Bookings
        'Pass 2: Update Accommodation Bookings
        'IF FULLY PAID:
        '- Booked (Invoiced) => Booked (Paid)
        '- Booked (Invoiced) Transfer => Booked (Paid) Transfer
        '- Waiting (Invoiced) => Waiting (Paid)
        'IF WAS FULLY PAID BUT NOW UNPAID/PART PAID:
        '- Booked (Paid) => Booked (Invoiced)
        '- Booked (Paid) Transfer => Booked (Invoiced) Transfer
        '- Waiting (Paid) => Waiting (Invoiced)
        For vPass = 1 To 2
          Select Case vPass
            Case 1
              vTable = "event_bookings"
              vAlias = "eb"
            Case 2
              vTable = "contact_room_bookings"
              vAlias = "crb"
          End Select
          With vWhereFields
            .Clear()
            .Add(vAlias & ".batch_number", BatchNumber, CDBField.FieldWhereOperators.fwoEqual)
            .Add(vAlias & ".transaction_number", TransactionNumber, CDBField.FieldWhereOperators.fwoEqual)
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT line_number,booking_status FROM " & vTable & " " & vAlias & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
            While vRecordSet.Fetch() = True
              vNewStatusCode = ""
              If mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = "F" And mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).SetValue <> "F" Then
                Select Case vRecordSet.Fields(2).Value
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaid)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid)
                End Select
              ElseIf (mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value <> "F" And mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).SetValue = "F") Then
                Select Case vRecordSet.Fields(2).Value
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaid)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
                  Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid)
                    vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced)
                End Select
              End If
              If vNewStatusCode.Length > 0 Then
                vWhereFields.Clear()
                vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, TransactionNumber)
                vWhereFields.Add("line_number", CDBField.FieldTypes.cftInteger, vRecordSet.Fields(1).Value)
                vUpdateFields = New CDBFields
                vUpdateFields.Add("booking_status", CDBField.FieldTypes.cftCharacter, vNewStatusCode)
                mvEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWhereFields, False)
              End If
            End While
            vRecordSet.CloseRecordSet()
          End With
        Next
        'Now handle exam exemptions
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then
          Dim vStudentExemption As New ExamStudentExemption(mvEnv)
          If mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) Then
            'If the invoice is not set to fully paid, then grant any exam exemptions waiting for this invoice to be paid
            vStudentExemption.GrantExemptions(BatchNumber, TransactionNumber)
          Else
            vStudentExemption.RevokeExemptions(BatchNumber, TransactionNumber)
          End If
        End If
      End If
      Save()
    End Sub

    Public Sub SetInvoicePrintingNotRequired()
      If InvoicePrinted = False Then mvClassFields.Item(InvoiceFields.ifPrintInvoice).Value = "N"
    End Sub

    Public Sub SetAmountPaid(ByVal pPaid As Double)
      SetAmountPaid(pPaid, False, False)
    End Sub

    Public Sub SetAmountPaid(ByVal pPaid As Double, ByVal pCheckInvoiceAmount As Boolean)
      SetAmountPaid(pPaid, pCheckInvoiceAmount, False)
    End Sub

    Public Sub SetAmountPaid(ByVal pPaid As Double, ByVal pCheckInvoiceAmount As Boolean, ByVal pInvoiceNumberNotRequired As Boolean)
      'pCheckInvoiceAmount = True will never the set the total amount paid greater than the original invoice amount. Currently used by Batch.WriteInvoice when creating a Credit Note
      If (pInvoiceNumberNotRequired OrElse DoubleValue(InvoiceNumber) > 0) AndAlso ((mvExisting = True) OrElse (mvExisting = False AndAlso mvInvoiceAmount <> 0)) Then
        SetInvoiceAmount()
        If pCheckInvoiceAmount AndAlso FixTwoPlaces(AmountPaid + pPaid) > mvInvoiceAmount Then
          mvClassFields.Item(InvoiceFields.ifAmountPaid).DoubleValue = mvInvoiceAmount
        Else
          mvClassFields.Item(InvoiceFields.ifAmountPaid).DoubleValue = (FixTwoPlaces(AmountPaid + pPaid))
        End If
        If System.Math.Abs(mvInvoiceAmount) = System.Math.Abs(AmountPaid) Then      'Figures could be negative
          mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid)
        ElseIf (System.Math.Abs(mvInvoiceAmount) <> AmountPaid) And AmountPaid <> 0 Then
          mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)
        Else
          mvClassFields.Item(InvoiceFields.ifInvoicePayStatus).Value = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
        End If
      End If
    End Sub

    Private Sub CheckInvoiceNumber(ByVal pCheckProvisionalInvoiceNumber As Boolean)
      If IntegerValue(InvoiceNumber) > 0 Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftLong, InvoiceNumber)
        If mvEnv.Connection.GetCount((mvClassFields.DatabaseTableName), vWhereFields, InvoiceNumber) > 0 Then
          If pCheckProvisionalInvoiceNumber Then
            RaiseError(DataAccessErrors.daeSetInvoiceNumberDuplicateRecord, "Provisional Invoice Number")
          Else
            RaiseError(DataAccessErrors.daeSetInvoiceNumberDuplicateRecord, "Invoice Number")
          End If
        End If
      End If
    End Sub

    ''' <summary>Used by CDBNet to decide whether to display the Remove Allocations menu</summary>
    ''' <returns>Count of Invoice Payment History records</returns>
    Public Function PaymentHistoryCount() As Integer
      Dim vWhereFields As New CDBFields

      Dim vTables As String = "invoice_payment_history iph"
      If Invoice.GetRecordType(RecordType) = Invoice.InvoiceRecordType.Invoice Then
        'Invoice - Find payments for this InvoiceNumber
        vWhereFields.Add("iph.invoice_number", InvoiceNumber)
      Else
        'For everything else, find the payment history that this invoice represents
        vTables = "invoice_details id, " & vTables
        With vWhereFields
          .Add("id.invoice_number", InvoiceNumber)
          .Add("id.batch_number", BatchNumber)
          .Add("id.transaction_number", TransactionNumber)
          .Add("iph.batch_number", CDBField.FieldTypes.cftInteger, "id.batch_number")
          .Add("iph.transaction_number", CDBField.FieldTypes.cftInteger, "id.transaction_number")
          .Add("iph.line_number", CDBField.FieldTypes.cftInteger, "id.line_number")
        End With
      End If
      With vWhereFields
        .Add("status", CDBField.FieldTypes.cftCharacter, "")
        .Add("amount", 0, CDBField.FieldWhereOperators.fwoGreaterThan)
      End With
      Return mvEnv.Connection.GetCount(vTables, Nothing, mvEnv.Connection.WhereClause(vWhereFields))
    End Function

    ''' <summary>When removing allocations for an invoice / invoice payment, perform a financial adjustment of the payment history</summary>
    ''' <param name="pIsInvoice">Boolean flag indicating whether this is an Invoice</param>
    ''' <returns>Batch / Transaction reference for the adjustment(s)</returns>
    ''' <remarks>Only used by RemoveAllocations</remarks>
    Private Function ProcessInvoicePaymentHistory(ByVal pIsInvoice As Boolean) As String
      Dim vCash As Boolean = False
      If RecordType = "C" Then vCash = True

      Dim vRecordType As Invoice.InvoiceRecordType
      If pIsInvoice = False Then vRecordType = Invoice.GetRecordType(RecordType)

      Dim vSLDT As New CDBDataTable
      vSLDT.AddColumnsFromList("BatchNumber,TransactionNumber,LineNumber,SalesLedgerAccount,Amount")

      Dim vTable As String = "invoice_payment_history iph"
      Dim vOrderBy As String = "iph.line_number"
      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields
      If pIsInvoice Then
        vOrderBy = "iph.batch_number, iph.transaction_number, " & vOrderBy
        vAnsiJoins.Add("invoices i", "i.invoice_number", "(select id.invoice_number from invoice_details id where iph.batch_number = id.batch_number AND iph.transaction_number = id.transaction_number and id.line_number = (select MIN(id2.line_number) from invoice_details id2 where iph.batch_number = id2.batch_number AND iph.transaction_number = id2.transaction_number))")

        With vWhereFields
          .Add("iph.invoice_number", IntegerValue(InvoiceNumber))
        End With
      Else
        vTable = "invoice_details id"   ', " & vTable
        vAnsiJoins.Add("invoice_payment_history iph", "id.batch_number", "iph.batch_number", "id.transaction_number", "iph.transaction_number", "id.line_number", "iph.line_number")
        With vWhereFields
          .Add("id.invoice_number", IntegerValue(InvoiceNumber))
          .Add("id.batch_number", BatchNumber)
          .Add("id.transaction_number", TransactionNumber)
        End With
      End If

      vWhereFields.Add("status", CDBField.FieldTypes.cftCharacter)
      If pIsInvoice Then
        'Include credit note allocations where the amount is negative
        vWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("amount#2", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("iph.batch_number", CDBField.FieldTypes.cftInteger, "iph.allocation_batch_number", CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("i.record_type", CDBField.FieldTypes.cftCharacter, "N", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Else
        vWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
      End If

      'Find and reverse all related Invoice Payment History
      Dim vAllocBatchNo As Integer
      Dim vAllocTransNo As Integer
      Dim vAllocLineNo As Integer
      Dim vLastBatch As Integer
      Dim vLastTransaction As Integer
      Dim vOriginalBatch As Batch
      Dim vOriginalBT As New BatchTransaction(mvEnv)
      Dim vOriginalBTA As BatchTransactionAnalysis
      Dim vAdjBatch As New Batch(mvEnv)
      Dim vAdjBT As New BatchTransaction(mvEnv)
      Dim vAdjBTA As BatchTransactionAnalysis
      Dim vAdjReference As String = ""
      Dim vFH As New FinancialHistory
      vFH.Init(mvEnv)
      Dim vIPH As New InvoicePaymentHistory(mvEnv)
      Dim vAttrs As String = vIPH.GetRecordSetFields()
      If pIsInvoice Then vAttrs &= ", i.record_type "
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, vTable, vWhereFields, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet()
      While vRS.Fetch = True
        vIPH = New InvoicePaymentHistory(mvEnv)
        vIPH.InitFromRecordSet(vRS)
        If pIsInvoice Then vRecordType = Invoice.GetRecordType(vRS.Fields("record_type").Value)
        Dim vInvoice As New Invoice()
        vInvoice.Init(mvEnv, vRS.Fields("batch_number").IntegerValue, vRS.Fields("transaction_number").IntegerValue)
        If vRecordType = InvoiceRecordType.CreditNote AndAlso Not vInvoice.IsSundryCreditNote Then
          'Simply reverse the IPH
          With vIPH
            .Reverse("R", .BatchNumber, .TransactionNumber, .LineNumber, Today)
          End With
        Else
          'Sales Ledger Cash or Sundry Credit Note
          vIPH.GetAllocationNumbers(vAllocBatchNo, vAllocTransNo, vAllocLineNo)
          If vAllocBatchNo = 0 OrElse (vAllocBatchNo <> vLastBatch OrElse vAllocTransNo <> vLastTransaction) Then
            '1. Find the transaction that created the allocation
            vFH = New FinancialHistory()
            vFH.Init(mvEnv, vAllocBatchNo, vAllocTransNo)
            If vFH.Existing = False Then RaiseError(DataAccessErrors.daeAllocationsBatchUnposted, vIPH.AllocationBatchNumber.ToString)
            For Each vFHD As FinancialHistoryDetail In vFH.Details
              If vFHD.LineNumber = vAllocLineNo Then
                If vFHD.Status <> FinancialHistory.FinancialHistoryStatus.fhsNormal Then
                  vInvoice = New Invoice()
                  vInvoice.Init(mvEnv, vFHD.BatchNumber, vFHD.TransactionNumber)
                  If vInvoice.IsSundryCreditNote Then
                    'If the invoice allocation is from a sundry credit note reversal then prevent further reversal and raise error
                    RaiseError(DataAccessErrors.daeCannotRemoveSundryCreditNoteReversal)
                  Else
                    RaiseError(DataAccessErrors.daeCannotAdjustPaymentStatus, vFH.StatusDesc(vFH.StatusCode), vFH.StatusDesc("R"))
                  End If
                End If
                Exit For
              End If
            Next

            '2.Setup the original Batch and Transaction
            vOriginalBatch = New Batch(mvEnv)
            vOriginalBatch.Init(vFH.BatchNumber)

            vOriginalBT = New BatchTransaction(mvEnv)
            vOriginalBT.Init(vFH.BatchNumber, vFH.TransactionNumber, True)  'J643: Set IncludeTransactionType to True so that it can be used later

            '3.Setup the reversal Batch and Transaction
            vAdjBatch = New Batch(mvEnv)
            vAdjBatch.InitOpenBatch(Nothing, Batch.ProvisionalOrConfirmed.Confirmed, vOriginalBatch.AdjustmentBatchType(Batch.AdjustmentTypes.atReverse), vOriginalBatch.BankAccount)

            vAdjBT = New BatchTransaction(mvEnv)
            With vAdjBT
              .InitFromBatch(mvEnv, vAdjBatch)
              .CloneForFA(vOriginalBT)
              .TransactionType = vOriginalBT.AdjustmentTransactionType(vOriginalBatch.BatchType, vOriginalBT.TransactionSign, Batch.AdjustmentTypes.atReverse)
              .TransactionDate = TodaysDate()
            End With
          End If
          If vFH.Existing = False Then RaiseError(DataAccessErrors.daeAllocationsBatchUnposted, vIPH.AllocationBatchNumber.ToString)

          '4.Setup the original Analysis
          vOriginalBTA = New BatchTransactionAnalysis(mvEnv)
          vOriginalBTA.Init(vFH.BatchNumber, vFH.TransactionNumber, vAllocLineNo)

          '5. Reverse the analysis line
          Dim vRemoveAlloc As Boolean = True
          If pIsInvoice = True AndAlso (vOriginalBTA.LineType = "L" OrElse vOriginalBTA.LineType = "K") Then vRemoveAlloc = False
          vFH.Reverse(vAdjBatch, vAdjBT, Batch.AdjustmentTypes.atReverse, vAllocLineNo, pRemoveInvoiceAllocations:=vRemoveAlloc, pUpdateCashInvoiceAmountPaid:=pIsInvoice)

          If vOriginalBTA.LineType = "N" Then
            'Invoice Payment
            '6. Store the Original BTA as we will have to add a U-type line to a FA transaction later
            vSLDT.AddRowFromList(vOriginalBTA.BatchNumber & "," & vOriginalBTA.TransactionNumber & "," & vOriginalBTA.LineNumber & "," & vOriginalBTA.MemberNumber & "," & vOriginalBTA.Amount)
          Else  'LineType = "L" (S/L Allocation of Cash-Invoice) or LineType = "K" (Sundry Credit Note Invoice Allocation)
            '7. Find and reverse the other matching L-type/K-type analysis line
            vOriginalBT.InitBatchTransactionAnalysis(vFH.BatchNumber, vFH.TransactionNumber)
            For Each vAllocationBTA As BatchTransactionAnalysis In vOriginalBT.Analysis
              With vAllocationBTA
                If .LineType = vOriginalBTA.LineType AndAlso .MemberNumber = vOriginalBTA.MemberNumber AndAlso .InvoiceNumber = vOriginalBTA.InvoiceNumber AndAlso .LineNumber > vOriginalBTA.LineNumber Then
                  vFH.Reverse(vAdjBatch, vAdjBT, Batch.AdjustmentTypes.atReverse, vAllocationBTA.LineNumber)
                  Exit For
                ElseIf .LineType.Equals("L", StringComparison.InvariantCultureIgnoreCase) AndAlso .LineType = vOriginalBTA.LineType AndAlso .MemberNumber <> vOriginalBTA.MemberNumber _
                       AndAlso .InvoiceNumber = vOriginalBTA.InvoiceNumber AndAlso .LineNumber > vOriginalBTA.LineNumber AndAlso .ContactNumber <> vOriginalBTA.ContactNumber _
                       AndAlso .ContactNumber = vFH.ContactNumber AndAlso System.Math.Abs(.Amount) = System.Math.Abs(vOriginalBTA.Amount) Then
                  vFH.Reverse(vAdjBatch, vAdjBT, Batch.AdjustmentTypes.atReverse, vAllocationBTA.LineNumber)
                  'Payment was made by a different Contact so need to update their Credit Customer record as the cash goes back to them
                  Dim vCC As New CreditCustomer()
                  vCC.InitCompanySalesLedgerAccount(mvEnv, Company, vAllocationBTA.MemberNumber)
                  If vCC.Existing Then
                    vCC.AdjustOutstanding(vAllocationBTA.Amount)
                    vCC.Save(mvEnv.User.UserID, True)
                  End If
                  vCC = New CreditCustomer()
                  vCC.InitCompanySalesLedgerAccount(mvEnv, Company, vOriginalBTA.MemberNumber)
                  If vCC.Existing Then
                    vCC.AdjustOutstanding(Math.Abs(vAllocationBTA.Amount))
                    vCC.Save(mvEnv.User.UserID, True)
                  End If
                  Exit For
                End If
              End With
            Next
          End If
          vLastBatch = vFH.BatchNumber
          vLastTransaction = vFH.TransactionNumber
          vAdjReference = vAdjBT.BatchNumber & "/" & vAdjBT.TransactionNumber
        End If
      End While
      vRS.CloseRecordSet()

      '8. Create another transaction in the FA batch with a U-type line for each Sales Ledger Account adjusted so that +ve unallocated Cash invoices will be created by Batch Processing
      If (pIsInvoice = True OrElse vCash = True) AndAlso vSLDT.Rows.Count > 0 Then
        vAdjBT = New BatchTransaction(mvEnv)
        With vAdjBT
          .InitFromBatch(mvEnv, vAdjBatch)
          .CloneForFA(vOriginalBT) 'This local object should still be available from above
          .TransactionDate = TodaysDate()
        End With

        Dim vSLAmount As Double = 0
        For vIndex As Integer = 0 To vSLDT.Rows.Count - 1
          vSLAmount += Val(vSLDT.Rows(vIndex).Item("Amount"))
        Next

        vOriginalBTA = New BatchTransactionAnalysis(mvEnv)
        vOriginalBTA.Init(CInt(vSLDT.Rows(0).Item("BatchNumber")), CInt(vSLDT.Rows(0).Item("TransactionNumber")), CInt(vSLDT.Rows(0).Item("LineNumber")))
        vAdjBTA = New BatchTransactionAnalysis(mvEnv)
        With vAdjBTA
          .InitFromTransaction(vAdjBT)
          .CloneFromBTA(vOriginalBTA)
          .LineType = "U"
          .Amount = vSLAmount
          .CurrencyAmount = vSLAmount
          .MemberNumber = vSLDT.Rows(0).Item("SalesLedgerAccount")
          .InvoiceNumber = 0
          .Save()
        End With

        vAdjBT.SaveChanges()

        'Update Batch - first set the correct batch total
        Dim vAmount As Double = 0
        Dim vCount As Integer = 0
        Dim vCurrAmount As Double = 0
        vWhereFields.Clear()
        vWhereFields.Add("batch_number", vAdjBatch.BatchNumber)
        vWhereFields.Add("tt.transaction_type", CDBField.FieldTypes.cftInteger, "bt.transaction_type")
        vSQL = New SQLStatement(mvEnv.Connection, "transaction_sign, amount, currency_amount", "batch_transactions bt, transaction_types tt", vWhereFields)
        vRS = vSQL.GetRecordSet()
        While vRS.Fetch = True
          vCount += 1
          If vRS.Fields("transaction_sign").Value = "D" Then
            vAmount -= vRS.Fields("amount").DoubleValue
            vCurrAmount -= vRS.Fields("currency_amount").DoubleValue
          Else
            vAmount += vRS.Fields("amount").DoubleValue
            vCurrAmount += vRS.Fields("currency_amount").DoubleValue
          End If
        End While
        vRS.CloseRecordSet()
        With vAdjBatch
          .BatchTotal = vAmount
          .CurrencyBatchTotal = vCurrAmount
          .NumberOfTransactions = vCount
          If .NumberOfEntries > 0 Then .NumberOfEntries = 0
          .SetBatchTotals()             'Force re-setting of the batch_totals
          .Save()
        End With
      End If

      Return vAdjReference
    End Function

    ''' <summary>Calculates the sum of payment allocations agaimst this invoice</summary>
    ''' <param name="pExcludeCreditNotes">Exclude credit note allocations</param>
    ''' <param name="pIncludeUnpostedTransactions">Include unposted transaction allocations</param>
    ''' <param name="pContainsUnpostedTransactions">When pIncludeUnpostedTransactions is True, indicates whether the allocations amount includes unposted allocations</param>
    ''' <returns>Sum of allocations for this invoice</returns>
    Public Function AllocationsAmount(ByVal pExcludeCreditNotes As Boolean, ByVal pIncludeUnpostedTransactions As Boolean, ByRef pContainsUnpostedTransactions As Boolean) As Double
      Dim vAllocations As Double
      Dim vWhereFields As New CDBFields(New CDBField("i.batch_number", BatchNumber))
      vWhereFields.Add("i.transaction_number", TransactionNumber)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("invoice_payment_history iph", "i.invoice_number", "iph.invoice_number")
      pContainsUnpostedTransactions = False
      If pExcludeCreditNotes Then
        'Event Bookings only at the moment
        'Make sure to not include any credit note payment
        vAnsiJoins.AddLeftOuterJoin("invoices i2", "iph.batch_number", "i2.batch_number", "iph.transaction_number", "i2.transaction_number")
        vWhereFields.Add("i2.record_type", "N", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If
      vAllocations = DoubleValue(New SQLStatement(mvEnv.Connection, "SUM(amount)", "invoices i", vWhereFields, "", vAnsiJoins).GetValue)
      If pIncludeUnpostedTransactions Then
        'Check for un-posted payments
        With vAnsiJoins
          .Clear()
          .AddLeftOuterJoin("batch_transaction_analysis bta", "i.invoice_number", "bta.invoice_number")
          .AddLeftOuterJoin("invoice_payment_history iph", "bta.batch_number", "iph.batch_number", "bta.transaction_number", "iph.transaction_number", "bta.line_number", "iph.line_number")
          .AddLeftOuterJoin("invoices i2", "iph.batch_number", "i2.batch_number", "iph.transaction_number", "i2.transaction_number")
        End With
        vWhereFields.Add("line_type", "N")    'Invoice payment
        vWhereFields.Add("iph.invoice_number", CDBField.FieldTypes.cftInteger, "")
        Dim vUnPostedAllocations As Double = DoubleValue(New SQLStatement(mvEnv.Connection, "SUM(bta.amount)", "invoices i", vWhereFields, "", vAnsiJoins).GetValue)
        If vUnPostedAllocations <> 0 Then pContainsUnpostedTransactions = True
        vAllocations = vAllocations + vUnPostedAllocations
      End If
      Return vAllocations
    End Function

    Public Sub SetAdjustmentStatus(ByVal pAdjustmentStatus As InvoiceAdjustmentStatus)
      If AdjustmentStatus = InvoiceAdjustmentStatus.Normal AndAlso RecordType = GetRecordTypeCode(InvoiceRecordType.SalesLedgerCash) Then
        mvClassFields.Item(InvoiceFields.ifAdjustmentStatus).Value = GetAdjustmentStatusCode(pAdjustmentStatus)
      End If
    End Sub

    ''' <summary>Get the InvoiceAmount of a cash-invoice for a re-analysed sales ledger transaction. Only used by Batch Posting when creating the Sales Ledger Invoice.</summary>
    ''' <param name="pBatchNumber">Batch Number of the adjustment transaction.</param>
    ''' <param name="pTransactionNumber">Transaction Number of the adjustment transaction.</param>
    ''' <param name="pFromBatchPosting">Is the cash-invoice currently being created by Batch Posting?</param>
    ''' <returns>The calculated InvoiceAmount</returns>
    ''' <remarks>Only used by Batch Posting when creating the Sales Ledger Invoice.</remarks>
    Friend Function GetAdjustmentInvoiceAmounts(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pFromBatchPosting As Boolean) As Double
      Dim vInvoicePaid As Double
      Return GetAdjustmentInvoiceAmounts(pBatchNumber, pTransactionNumber, pFromBatchPosting, "", vInvoicePaid)
    End Function
    ''' <summary>Get the InvoiceAmount of a cash-invoice for a re-analysed sales ledger transaction. Only used by Batch Posting when creating the Sales Ledger Invoice.</summary>
    ''' <param name="pBatchNumber">Batch Number of the adjustment transaction.</param>
    ''' <param name="pTransactionNumber">Transaction Number of the adjustment transaction.</param>
    ''' <param name="pFromBatchPosting">Is the cash-invoice currently being created by Batch Posting?</param>
    ''' <param name="pSalesLedgerAccount">Sales Ledger Account of the adjustment transaction. Only used when <paramref name="pFromBatchPosting">pFromBatchPosting</paramref> is False.</param>
    ''' <param name="pAmountPaid">Set to the sum of the AmountPaid</param>
    ''' <returns>The calculated InvoiceAmount</returns>
    ''' <remarks>Only used by Batch Posting when creating the Sales Ledger Invoice.</remarks>
    Friend Function GetAdjustmentInvoiceAmounts(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pFromBatchPosting As Boolean, ByVal pSalesLedgerAccount As String, ByRef pAmountPaid As Double) As Double
      Dim vInvoiceAmount As Double
      Dim vTableName As String = "batch_transaction_analysis t"
      Dim vGroupBy As String = "t.batch_number, t.transaction_number"
      Dim vAnsiJoins As AnsiJoins = Nothing

      pAmountPaid = 0   'Re-set
      Dim vWhereFields As New CDBFields()
      With vWhereFields
        .Add("t.batch_number", pBatchNumber)
        .Add("t.transaction_number", pTransactionNumber)
        If pFromBatchPosting Then .Add("t.line_type", CDBField.FieldTypes.cftCharacter, "'N','U','L'", CDBField.FieldWhereOperators.fwoIn)
      End With

      If pFromBatchPosting = False Then
        vTableName = "invoices i"
        If String.IsNullOrWhiteSpace(pSalesLedgerAccount) = False Then vWhereFields.Add("i.sales_ledger_account", CDBField.FieldTypes.cftCharacter, pSalesLedgerAccount)
        vAnsiJoins = New AnsiJoins
        With vAnsiJoins
          .Add("invoice_details id", "i.batch_number", "id.batch_number", "i.transaction_number", "id.transaction_number", "i.invoice_number", "id.invoice_number")
          .Add("financial_history_details t", "id.batch_number", "t.batch_number", "id.transaction_number", "t.transaction_number", "id.line_number", "t.line_number")
        End With
        vGroupBy &= ", i.amount_paid"
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "SUM(t.amount) AS adj_amount" & If(pFromBatchPosting = True, "", ", i.amount_paid"), vTableName, vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = vGroupBy
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
      If vRS.Fetch Then
        vInvoiceAmount = vRS.Fields("adj_amount").DoubleValue
        If pFromBatchPosting = False Then pAmountPaid = vRS.Fields("amount_paid").DoubleValue
      End If
      vRS.CloseRecordSet()

      If pFromBatchPosting = False Then
        'Now get the original Invoice amount
        vTableName = "reversals t"
        vAnsiJoins = New AnsiJoins
        With vAnsiJoins
          .Add("invoices i", "t.was_batch_number", "i.batch_number", "t.was_transaction_number", "i.transaction_number")
          .Add("invoice_details id", "i.batch_number", "id.batch_number", "i.transaction_number", "id.transaction_number")
          .Add("financial_history_details fhd", "id.batch_number", "fhd.batch_number", "id.transaction_number", "fhd.transaction_number", "id.line_number", "fhd.line_number")
        End With

        Dim vBatchNumber As Integer = pBatchNumber
        Dim vTransNumber As Integer = pTransactionNumber
        While vBatchNumber > 0
          vWhereFields.Item(1).Value = vBatchNumber.ToString
          vWhereFields.Item(2).Value = vTransNumber.ToString
          vGroupBy &= ", i.batch_number, i.transaction_number"
          vSQLStatement = New SQLStatement(mvEnv.Connection, "SUM(fhd.amount) AS orig_amount, i.amount_paid, i.batch_number, i.transaction_number", vTableName, vWhereFields, "", vAnsiJoins)
          vSQLStatement.GroupBy = vGroupBy
          vBatchNumber = 0
          vTransNumber = 0
          vRS = vSQLStatement.GetRecordSet
          If vRS.Fetch Then
            vInvoiceAmount += vRS.Fields("orig_amount").DoubleValue
            pAmountPaid += vRS.Fields("amount_paid").DoubleValue
            vBatchNumber = vRS.Fields("batch_number").IntegerValue
            vTransNumber = vRS.Fields("transaction_number").IntegerValue
          End If
          vRS.CloseRecordSet()
        End While
      End If

      Return vInvoiceAmount

    End Function

    Private Sub CheckAllocationsForRemoval(ByVal pIsInvoice As Boolean)
      Dim vRecordType As Invoice.InvoiceRecordType
      Dim vTable As String = "invoice_payment_history iph"
      Dim vOrderBy As String = "iph.line_number"
      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields

      If Not pIsInvoice Then
        vRecordType = Invoice.GetRecordType(RecordType)
      End If

      If pIsInvoice Then
        vOrderBy = "iph.batch_number, iph.transaction_number, " & vOrderBy
        vAnsiJoins.Add("invoices i", "i.invoice_number", "(select id.invoice_number from invoice_details id where iph.batch_number = id.batch_number AND iph.transaction_number = id.transaction_number and id.line_number = (select MIN(id2.line_number) from invoice_details id2 where iph.batch_number = id2.batch_number AND iph.transaction_number = id2.transaction_number))")
        vWhereFields.Add("iph.invoice_number", IntegerValue(InvoiceNumber))
      Else
        vTable = "invoice_details id"
        vAnsiJoins.Add("invoice_payment_history iph", "id.batch_number", "iph.batch_number", "id.transaction_number", "iph.transaction_number", "id.line_number", "iph.line_number")
        vWhereFields.Add("id.invoice_number", IntegerValue(InvoiceNumber))
        vWhereFields.Add("id.batch_number", BatchNumber)
        vWhereFields.Add("id.transaction_number", TransactionNumber)
      End If
      vWhereFields.Add("status", CDBField.FieldTypes.cftCharacter)
      vWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)

      'Find all related Invoice Payment History
      Dim vAllocBatchNo As Integer
      Dim vAllocTransNo As Integer
      Dim vAllocLineNo As Integer
      Dim vLastBatch As Integer
      Dim vLastTransaction As Integer

      Dim vFH As New FinancialHistory
      vFH.Init(mvEnv)
      Dim vIPH As New InvoicePaymentHistory(mvEnv)
      Dim vAttrs As String = vIPH.GetRecordSetFields()
      If pIsInvoice Then vAttrs &= ", i.record_type "
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, vTable, vWhereFields, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet()
      While vRS.Fetch = True
        vIPH = New InvoicePaymentHistory(mvEnv)
        vIPH.InitFromRecordSet(vRS)
        If pIsInvoice Then vRecordType = Invoice.GetRecordType(vRS.Fields("record_type").Value)
        Dim vInvoice As New Invoice()
        vInvoice.Init(mvEnv, vRS.Fields("batch_number").IntegerValue, vRS.Fields("transaction_number").IntegerValue)
        If vRecordType = InvoiceRecordType.SalesLedgerCash OrElse (vRecordType = InvoiceRecordType.CreditNote AndAlso vInvoice.IsSundryCreditNote) Then
          'Sales Ledger Cash or Sundry Credit Note
          vIPH.GetAllocationNumbers(vAllocBatchNo, vAllocTransNo, vAllocLineNo)
          If vAllocBatchNo = 0 OrElse (vAllocBatchNo <> vLastBatch OrElse vAllocTransNo <> vLastTransaction) Then
            '1. Find the transaction that created the allocation
            vFH = New FinancialHistory()
            vFH.Init(mvEnv, vAllocBatchNo, vAllocTransNo)
            If vFH.Existing = False Then
              RaiseError(DataAccessErrors.daeAllocationsBatchUnposted, vIPH.AllocationBatchNumber.ToString)
            End If
            For Each vFHD As FinancialHistoryDetail In vFH.Details
              If vFHD.LineNumber = vAllocLineNo Then
                If vFHD.Status <> FinancialHistory.FinancialHistoryStatus.fhsNormal Then
                  vInvoice = New Invoice()
                  vInvoice.Init(mvEnv, vFHD.BatchNumber, vFHD.TransactionNumber)
                  If vInvoice.IsSundryCreditNote Then
                    'If the invoice allocation is from a sundry credit note reversal then prevent further reversal and raise error
                    RaiseError(DataAccessErrors.daeCannotRemoveSundryCreditNoteReversal)
                  Else
                    RaiseError(DataAccessErrors.daeCannotAdjustPaymentStatus, vFH.StatusDesc(vFH.StatusCode), vFH.StatusDesc("R"))
                  End If
                End If
                Exit For
              End If
            Next
          End If
          If vFH.Existing = False Then
            RaiseError(DataAccessErrors.daeAllocationsBatchUnposted, vIPH.AllocationBatchNumber.ToString)
          End If
          vLastBatch = vFH.BatchNumber
          vLastTransaction = vFH.TransactionNumber
        End If
      End While
      vRS.CloseRecordSet()
    End Sub

    Public ReadOnly Property IsSundryCreditNote() As Boolean
      Get
        'Sundry Credit Note invoice type record will not have a reversals record
        'therefore return whether the count of linked reversals records = 0
        If Not mvIsSundryCreditNote.HasValue Then
          If Existing Then
            If Not RecordType = "N" Then Return False
            Dim vAnsiJoins As New AnsiJoins()
            vAnsiJoins.Add("batch_transactions bt", "i.batch_number", "bt.batch_number", "i.transaction_number", "bt.transaction_number")
            vAnsiJoins.Add("reversals r", "bt.batch_number", "r.batch_number", "bt.transaction_number", "r.transaction_number")
            Dim vWhereFields As New CDBFields(New CDBField("i.invoice_number", CDBField.FieldTypes.cftLong, InvoiceNumber))
            If InvoiceNumber.Length = 0 Then
              'Credit Note has not been printed so include Batch & Transaction Numbers
              vWhereFields.Add("i.batch_number", BatchNumber)
              vWhereFields.Add("i.transaction_number", TransactionNumber)
            End If
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "", "invoices i", vWhereFields, "", vAnsiJoins)

            mvIsSundryCreditNote = mvEnv.Connection.GetCountFromStatement(vSQLStatement) = 0
          Else
            mvIsSundryCreditNote = False
          End If
        End If
        Return mvIsSundryCreditNote.Value
      End Get
    End Property

    ''' <summary>Gets the <see cref="FinancialHistory">FinancialHistory</see> for this Invoice.</summary>
    Private ReadOnly Property FinancialHistory As FinancialHistory
      Get
        If mvFinancialHistory Is Nothing OrElse mvFinancialHistory.Existing = False Then
          mvFinancialHistory = New FinancialHistory()
          mvFinancialHistory.Init(mvEnv, Me.BatchNumber, Me.TransactionNumber)
        End If
        Return mvFinancialHistory
      End Get
    End Property

    ''' <summary>Gets a boolean indicating whether this Invoice is a result of a Financial Adjustment.</summary>
    ''' <returns></returns>
    Friend ReadOnly Property IsFinancialAdjustmentInvoice As Boolean
      Get
        Dim vWhereFields As New CDBFields(New CDBField(mvClassFields(InvoiceFields.ifBatchNumber).Name, Me.BatchNumber))
        Dim vBatchTypeCode As String = New SQLStatement(mvEnv.Connection, "batch_type", "batches", vWhereFields).GetValue()
        If Not (String.IsNullOrWhiteSpace(vBatchTypeCode)) AndAlso vBatchTypeCode.Equals(Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment), StringComparison.InvariantCultureIgnoreCase) Then
          Return True
        Else
          Return False
        End If
      End Get
    End Property

  End Class
End Namespace
