Namespace Access

  Partial Public Class Invoice

    'RecordType enum
    Public Enum InvoiceRecordType
      Invoice
      CreditNote
      SalesLedgerCash
    End Enum

    Public Enum InvoiceAdjustmentStatus
      Normal      'Null    Normal value
      Adjusted    'A       Transaction has been adjusted
      Moved       'M       Transaction has been moved to another contact
      Reversed    'R       Transaction has been reversed (refunded)
    End Enum

    ''' <summary>Returns the invoice RecordType code for the InvoiceRecordType enum</summary>
    ''' <remarks>Valid values for the RecordType code are (I)nvoice, Credit (N)ote and Sales Ledger (C)ash</remarks>
    Friend Shared Function GetRecordTypeCode(ByVal pRecordType As InvoiceRecordType) As String
      Select Case pRecordType
        Case InvoiceRecordType.CreditNote
          Return "N"
        Case InvoiceRecordType.Invoice
          Return "I"
        Case Else
          Return "C"
      End Select
    End Function

    ''' <summary>Returns the invoice InvoiceRecordType enum for the RecordType code</summary>
    ''' <remarks>Valid values for the RecordType code are (I)nvoice, Credit (N)ote and Sales Ledger (C)ash</remarks>
    Friend Shared Function GetRecordType(ByVal pRecordTypeCode As String) As InvoiceRecordType
      Select Case pRecordTypeCode
        Case "C"
          Return InvoiceRecordType.SalesLedgerCash
        Case "I"
          Return InvoiceRecordType.Invoice
        Case Else
          Return InvoiceRecordType.CreditNote
      End Select
    End Function

    ''' <summary>Gets the invoice AdjustmentStatus code for the InvoiceAdjustmentStatus enum.</summary>
    ''' <remarks>Valid values for the AdjustmentStatus as (A)djusted, (M)oved, (R)eversed and null (normal).</remarks>
    Friend Shared Function GetAdjustmentStatusCode(ByVal pAdjustmentStatus As InvoiceAdjustmentStatus) As String
      Select Case pAdjustmentStatus
        Case InvoiceAdjustmentStatus.Adjusted
          Return "A"
        Case InvoiceAdjustmentStatus.Moved
          Return "M"
        Case InvoiceAdjustmentStatus.Reversed
          Return "R"
        Case Else
          Return ""
      End Select
    End Function

    ''' <summary>Gets the InvoiceAdjustmentStatus enum for the InvoiceAdjustmentStatus code.</summary>
    ''' <remarks>Valid values for the AdjustmentStatus as (A)djusted, (M)oved, (R)eversed and null (normal).</remarks>
    Friend Shared Function GetAdjustmentStatus(ByVal pAdjustmentStatusCode As String) As InvoiceAdjustmentStatus
      Select Case pAdjustmentStatusCode
        Case "A"
          Return InvoiceAdjustmentStatus.Adjusted
        Case "M"
          Return InvoiceAdjustmentStatus.Moved
        Case "R"
          Return InvoiceAdjustmentStatus.Reversed
        Case Else
          Return InvoiceAdjustmentStatus.Normal
      End Select
    End Function

    ''' <summary>
    ''' Returns whether the invoice has been refunded or partially refunded
    ''' </summary>
    ''' <param name="pConnection"></param>
    ''' <param name="pInvoiceNumber"></param>
    ''' <remarks>Applicable for example where the invoice was for multiple event/exam bookings and one of the bookings has been cancelled resulting in a reversal and credit note</remarks>
    Public Shared Function IsPartiallyRefunded(ByVal pConnection As CDBConnection, ByVal pInvoiceNumber As Integer) As Boolean
      Dim vWhereFields As New CDBFields(New CDBField("i.invoice_number", pInvoiceNumber, CDBField.FieldWhereOperators.fwoEqual))
      vWhereFields.Add("fh.amount", CDBField.FieldTypes.cftNumeric, "fh2.amount", CDBField.FieldWhereOperators.fwoGreaterThan)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("financial_history fh", "i.batch_number", "fh.batch_number", "i.transaction_number", "fh.transaction_number")
      vAnsiJoins.Add("reversals r", "fh.batch_number", "r.was_batch_number", "fh.transaction_number", "r.was_transaction_number")
      vAnsiJoins.Add("financial_history fh2", "r.batch_number", "fh2.batch_number", "r.transaction_number", "fh2.transaction_number")

      Dim vSQLStatement As New SQLStatement(pConnection, "", "invoices i", vWhereFields, "", vAnsiJoins)
      Return pConnection.GetCountFromStatement(vSQLStatement) > 0
    End Function

    Public Shared Function RemoveInvoiceAllocations(ByVal pEnv As CDBEnvironment, ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByRef pInvoice As Invoice, ByVal pBatchType As Batch.BatchTypes, ByRef pInvoicesToDelete As String, ByVal pFHAdjustmentType As Batch.AdjustmentTypes) As Boolean
      Dim vRemoved As Boolean
      Dim vUpdateFields As New CDBFields

      '1. Get original payment invoices record -  
      'select was_* attrs, invoice # from reversals, invoices where r batch attrs = pBTA properties, join to invoices via was_batch/trans AND record type in C or N
      Dim vNestedSQL As SQLStatement
      Dim vInvoice As New Invoice()
      vInvoice.Init(pEnv)
      Dim vWherefields As New CDBFields(New CDBField("r.batch_number", pBTA.BatchNumber))
      vWherefields.Add("r.transaction_number", pBTA.TransactionNumber)
      Dim vAnsiJoins As New AnsiJoins()
      If (pBTA.LineType = "L" OrElse pBTA.LineType = "K") AndAlso pBTA.Amount < 0 Then
        vAnsiJoins.Add("invoice_payment_history iph", "was_batch_number", "iph.allocation_batch_number", "was_transaction_number", "iph.allocation_transaction_number", "was_line_number", "iph.allocation_line_number")
        vAnsiJoins.Add("invoice_details id", "iph.batch_number", "id.batch_number", "iph.transaction_number", "id.transaction_number", "iph.line_number", "id.line_number")
      Else
        vAnsiJoins.Add("invoice_details id", "was_batch_number", "id.batch_number", "was_transaction_number", "id.transaction_number", "was_line_number", "id.line_number")
      End If
      vAnsiJoins.Add("batch_transaction_analysis bta", "id.batch_number", "bta.batch_number", "id.transaction_number", "bta.transaction_number", "id.line_number", "bta.line_number")
      Dim vNestedAttrs As String = vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", SUM(bta.amount) AS invoice_amount"

      If pFHAdjustmentType = Batch.AdjustmentTypes.atNone Then
        'If we have come from BatchPosting of a sundry credit note reversal ensure it has a credit note number in order for the allocations to be added below
        vAnsiJoins.Add("invoices i", "id.batch_number", "i.batch_number", "id.transaction_number", "i.transaction_number")
        vNestedSQL = New SQLStatement(pEnv.Connection, vNestedAttrs, "reversals r", vWherefields, "", vAnsiJoins)
        vNestedSQL.GroupBy = vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll)
        Dim vRS As CDBRecordSet = vNestedSQL.GetRecordSet()
        If vRS.Fetch Then vInvoice.InitFromRecordSet(pEnv, vRS, InvoiceRecordSetTypes.irtAll)
        vRS.CloseRecordSet()
        If vInvoice.Existing = True AndAlso vInvoice.InvoiceType = InvoiceRecordType.CreditNote AndAlso vInvoice.IsSundryCreditNote = True Then
          If vInvoice.InvoiceNumber.Length = 0 Then
            vInvoice.SetInvoiceNumber(True, True)
            vInvoice.Save(pEnv.User.UserID, True)
          End If
        End If
        vAnsiJoins.RemoveJoin("invoices i")   'Remove last item
        vInvoice = New Invoice()
        vInvoice.Init(pEnv)
      End If

      vAnsiJoins.Add("invoices i", "id.invoice_number", "i.invoice_number")
      'Build nested SQL
      vNestedSQL = New SQLStatement(pEnv.Connection, vNestedAttrs, "reversals r", vWherefields, "", vAnsiJoins)
      vNestedSQL.GroupBy = vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll)
      Dim vNestedSQLString As String = vNestedSQL.SQL
      'Build main SQL
      Dim vAttrs As String = "was_batch_number, was_transaction_number, was_line_number, bta.amount, " & vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", i.invoice_amount"
      With vWherefields
        .Add("r.line_number", pBTA.LineNumber)
        .Add("record_type", "'C','N'", CDBField.FieldWhereOperators.fwoIn)
      End With
      vAnsiJoins.RemoveAt(vAnsiJoins.Count - 1)   'Remove the last item which is Invoices
      vAnsiJoins.Add("(" & vNestedSQLString & ") i", "id.invoice_number", "i.invoice_number")
      'Handle part-refunds - need to be excluded as creating the part-refund has already done this
      vWherefields.Add("fhd.status", CDBField.FieldTypes.cftCharacter, "A", CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWherefields.Add("ABS(bta.amount)", CDBField.FieldTypes.cftInteger, "ABS(newbta.amount)", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWherefields.Add("fhd.status#2", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWherefields.Add("fhd.status#3", CDBField.FieldTypes.cftCharacter, "A", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      vAnsiJoins.Add("batch_transaction_analysis newbta", "r.batch_number", "newbta.batch_number", "r.transaction_number", "newbta.transaction_number", "r.line_number", "newbta.line_number")
      vAnsiJoins.AddLeftOuterJoin("financial_history_details fhd", "bta.batch_number", "fhd.batch_number", "bta.transaction_number", "fhd.transaction_number", "bta.line_number", "fhd.line_number")
      Dim vSQLStatement As New SQLStatement(pEnv.Connection, vAttrs, "reversals r", vWherefields, "", vAnsiJoins)
      Dim vINVRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      With vINVRS
        Dim vLastSLCompany As String = ""
        Dim vLastSLAccount As String = ""
        Dim vSLAToUpdate As String = String.Empty
        While .Fetch() = True
          vInvoice = New Invoice()
          vInvoice.InitFromRecordSet(pEnv, vINVRS, Invoice.InvoiceRecordSetTypes.irtAll)
          'Set passed in invoice object to the current invoice object
          pInvoice = vInvoice
          'Are we looking at the same Sales Ledger Account?  If so then don't update the Credit Customer again.
          Dim vCustomerUpdated As Boolean = .Fields("company").Value = vLastSLCompany AndAlso .Fields("sales_ledger_account").Value = vLastSLAccount
          vSLAToUpdate = .Fields("sales_ledger_account").Value

          '2.  Get the Invoice we originally paid -  select from IPH, invoices where batch/trans/line = was_* attrs, join to invoices via invoice #
          vAttrs = "DISTINCT i.invoice_number,i.company,i.sales_ledger_account,i.amount_paid,iph.batch_number,iph.transaction_number,iph.line_number,bt.amount,fh.status"
          vAttrs &= ",iph.status AS iph_status, iph.allocation_batch_number, iph.allocation_transaction_number, iph.allocation_line_number"
          vAttrs &= ", cn.batch_number AS cn_batch_number, cniph.batch_number AS cniph_batch_number"
          With vAnsiJoins
            .Clear()
            .Add("invoices i", "iph.invoice_number", "i.invoice_number")
            .Add("batch_transactions bt", "i.batch_number", "bt.batch_number", "i.transaction_number", "bt.transaction_number")
            .AddLeftOuterJoin("financial_history fh", "bt.batch_number", "fh.batch_number", "bt.transaction_number", "fh.transaction_number")
            .AddLeftOuterJoin("reversals r", "i.batch_number", "r.was_batch_number", "i.transaction_number", "r.was_transaction_number")
            .AddLeftOuterJoin("invoices cn", "r.batch_number", "cn.batch_number", "r.transaction_number", "cn.transaction_number")
            .AddLeftOuterJoin("invoice_payment_history cniph", "i.invoice_number", "cniph.invoice_number", "cn.batch_number", "cniph.batch_number", "cn.transaction_number", "cniph.transaction_number")
          End With
          With vWherefields
            .Clear()
            If (pBTA.LineType = "L" OrElse pBTA.LineType = "K") AndAlso pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) = True Then
              .Add("iph.allocation_batch_number", "", CType(CDBField.FieldWhereOperators.fwoNotEqual + CDBField.FieldWhereOperators.fwoOpenBracketTwice, CDBField.FieldWhereOperators))
              .Add("iph.allocation_batch_number#2", vINVRS.Fields(1).IntegerValue)
              .Add("iph.allocation_transaction_number", vINVRS.Fields(2).IntegerValue)
              .Add("iph.allocation_line_number", vINVRS.Fields(3).IntegerValue, CDBField.FieldWhereOperators.fwoCloseBracket)
              .Add("iph.allocation_batch_number#3", "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket + CDBField.FieldWhereOperators.fwoOR, CDBField.FieldWhereOperators))
              .Add("iph.batch_number", vINVRS.Fields(1).IntegerValue)
              .Add("iph.transaction_number", vINVRS.Fields(2).IntegerValue)
              .Add("iph.line_number", vINVRS.Fields(3).IntegerValue, CDBField.FieldWhereOperators.fwoCloseBracketTwice)
            Else
              .Add("iph.batch_number", vINVRS.Fields(1).IntegerValue)
              .Add("iph.transaction_number", vINVRS.Fields(2).IntegerValue)
              .Add("iph.line_number", vINVRS.Fields(3).IntegerValue)
            End If
          End With
          vSQLStatement = New SQLStatement(pEnv.Connection, vAttrs, "invoice_payment_history iph", vWherefields, "", vAnsiJoins)
          Dim vIPHRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          With vIPHRS
            While .Fetch() = True
              vRemoved = True
              Dim vContinue As Boolean = True
              If .Fields("status").Value.Length > 0 Then    'fh.status
                If .Fields(4).DoubleValue = 0 Then    '4=AmountPaid  'All allocations were removed for this invoice (invoice had been already updated) and Credit note batch has not been processed
                  'do nothing, processing Credit note batch will update this correctly.
                  vContinue = False
                  If .Fields("cn_batch_number").Value.Length > 0 AndAlso .Fields("cniph_batch_number").Value.Length > 0 Then
                    'Reversing a credit note that was created to reverse an invoice
                    vCustomerUpdated = True   'Don't update the Credit Customer as it will be done when the invoice is created
                  End If
                ElseIf .Fields(4).DoubleValue = .Fields(8).DoubleValue Then   '4=AmountPaid, 8=BT.Amount  'Credit note batch has been processed
                  If pBatchType = Batch.BatchTypes.FinancialAdjustment AndAlso pBT.Amount = 0 AndAlso pBTA.Amount < 0 AndAlso pBTA.LineType = "N" Then
                    'Appears to be a re-analysis, removing an invoice payment
                    vContinue = True
                  ElseIf .Fields("cn_batch_number").Value.Length > 0 Then
                    'Original invoice has been cancelled and we have a credit note
                    vContinue = (.Fields("cniph_batch_number").Value.Length = 0)    'If credit note is not allocated to the invoice then set vContinue to True
                  Else
                    'No credit note - do nothing
                    vContinue = False
                  End If
                Else
                  'Only the Payment Allocation are removed AND
                  'EITHER Credit Note batch has not been processed
                  'OR Credit Note batch has been processed but there are some invoice details that has not been reversed yet (invoice amount greater than total paid) 

                  'BR17004: Where the Invoice has been partially refunded (e.g. through multi booking invoice and a booking is cancelled) and the reversal credit note
                  'has been posted first, the invoice amount_paid will have already been updated and so should not be updated again.
                  'Where no reversal credit note exists or it is not posted continue with the update
                  vContinue = Not Invoice.IsPartiallyRefunded(pEnv.Connection, .Fields(1).IntegerValue)
                End If
              ElseIf .Fields("iph_status").Value.Length > 0 Then  'iph.status
                If pBTA.LineType = "U" AndAlso (.Fields("batch_number").IntegerValue <> .Fields("allocation_batch_number").IntegerValue) AndAlso .Fields("allocation_batch_number").IntegerValue > 0 Then
                  'Original Un-allocated S/L Cash was allocated in separate batch and since removed from this Invoice
                  'This allocation will have already updated the Invoice so leave as it is, if the removal was not by a Move
                  'For financial history reversal/refund allow update of the invoice
                  If .Fields("iph_status").Value.ToUpper.Equals("M") Then
                    vSLAToUpdate = .Fields("sales_ledger_account").Value  'Need to use the SalesLedgerAccount for the Invoice that was paid
                  Else
                    If pFHAdjustmentType = Batch.AdjustmentTypes.atNone Then vContinue = False
                  End If
                ElseIf (pBTA.LineType.Equals("N", StringComparison.InvariantCultureIgnoreCase) AndAlso .Fields("batch_number").IntegerValue.Equals(.Fields("allocation_batch_number").IntegerValue) = True) Then
                  If pFHAdjustmentType = Batch.AdjustmentTypes.atNone AndAlso .Fields("iph_status").Value.Equals("R", StringComparison.InvariantCultureIgnoreCase) Then
                    'The original reversal of this transaction has already updated the Invoice amount paid
                    vContinue = False
                  End If
                End If
              End If
              Dim vAdjInvoiceAmountPaid As Double
              Dim vInvoiceNumber As Integer
              Dim vPayStatus As String = ""
              If vContinue Then
                '2a. update the invoice being paid
                vInvoiceNumber = .Fields(1).IntegerValue
                vAdjInvoiceAmountPaid = .Fields(4).DoubleValue - System.Math.Abs(pBTA.Amount) 'IS THIS +ve or -ve?
                If vAdjInvoiceAmountPaid < 0 Then vAdjInvoiceAmountPaid = 0
                If FixTwoPlaces(vAdjInvoiceAmountPaid) = FixTwoPlaces(.Fields(8).DoubleValue) Then
                  vPayStatus = pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid)
                Else
                  If (FixTwoPlaces(.Fields(8).DoubleValue) <> FixTwoPlaces(vAdjInvoiceAmountPaid)) And vAdjInvoiceAmountPaid > 0 Then
                    vPayStatus = pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)
                  Else
                    vPayStatus = pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
                  End If
                  'BR 8818 Update Booking Status on any associated Event/Accommodation Bookings for original Invoice
                  'Pass 1: Update Event Bookings
                  'Pass 2: Update Accommodation Bookings
                  '- Booked (Paid) => Booked (Invoiced)
                  '- Booked (Paid) Transfer => Booked (Invoiced) Transfer
                  '- Waiting (Paid) => Waiting (Invoiced)
                  If vInvoiceNumber > 0 Then
                    vWherefields = New CDBFields
                    vWherefields.Add("invoice_number", CDBField.FieldTypes.cftLong, vInvoiceNumber)
                    vSQLStatement = New SQLStatement(pEnv.Connection, "batch_number,transaction_number,invoice_pay_status", "invoices", vWherefields)
                    Dim vRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet()
                    Dim vBatchNumber As Integer
                    Dim vTransactionNumber As Integer
                    Dim vTable As String = ""
                    Dim vAlias As String = ""
                    If vRecordSet.Fetch() = True Then
                      vBatchNumber = vRecordSet.Fields(1).IntegerValue
                      vTransactionNumber = vRecordSet.Fields(2).IntegerValue
                      For vPass As Integer = 1 To 2
                        Select Case vPass
                          Case 1
                            vTable = "event_bookings"
                            vAlias = "eb"
                          Case 2
                            vTable = "contact_room_bookings"
                            vAlias = "crb"
                        End Select
                        vWherefields.Clear()
                        vWherefields.Add(vAlias & ".batch_number", vBatchNumber, CDBField.FieldWhereOperators.fwoEqual)
                        vWherefields.Add(vAlias & ".transaction_number", vTransactionNumber, CDBField.FieldWhereOperators.fwoEqual)
                        vSQLStatement = New SQLStatement(pEnv.Connection, "line_number,booking_status", vTable & " " & vAlias, vWherefields)
                        Dim vRSBookings As CDBRecordSet = vSQLStatement.GetRecordSet()
                        While vRSBookings.Fetch() = True
                          Dim vNewStatusCode As String = ""
                          Select Case vRSBookings.Fields(2).Value
                            Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaid)
                              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced)
                            Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer)
                              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
                            Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid)
                              vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced)
                          End Select
                          If vNewStatusCode.Length > 0 Then
                            vWherefields.Clear()
                            vWherefields.Add("batch_number", CDBField.FieldTypes.cftLong, vBatchNumber)
                            vWherefields.Add("transaction_number", CDBField.FieldTypes.cftInteger, vTransactionNumber)
                            vWherefields.Add("line_number", CDBField.FieldTypes.cftInteger, vRSBookings.Fields(1).Value)
                            vUpdateFields = New CDBFields
                            vUpdateFields.Add("booking_status", CDBField.FieldTypes.cftCharacter, vNewStatusCode)
                            pEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWherefields, False)
                          End If
                        End While
                        vRSBookings.CloseRecordSet()
                      Next
                      'Now handle exam exemptions
                      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then
                        Dim vStudentExemption As New ExamStudentExemption(pEnv)
                        If vRecordSet.Fields(3).Value <> pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) AndAlso vPayStatus = pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) Then
                          vStudentExemption.GrantExemptions(vBatchNumber, vTransactionNumber)
                        ElseIf vRecordSet.Fields(3).Value = pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) AndAlso vPayStatus <> pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) Then
                          vStudentExemption.RevokeExemptions(vBatchNumber, vTransactionNumber)
                        End If
                      End If
                    End If
                    vRecordSet.CloseRecordSet()
                  End If
                End If
              End If
              If vInvoiceNumber > 0 Then
                vUpdateFields = New CDBFields
                vUpdateFields.Add("invoice_pay_status", CDBField.FieldTypes.cftCharacter, vPayStatus)
                vUpdateFields.Add("amount_paid", CDBField.FieldTypes.cftNumeric, vAdjInvoiceAmountPaid)
                vWherefields.Clear()
                vWherefields.Add("invoice_number", CDBField.FieldTypes.cftLong, .Fields(1).IntegerValue)
                pEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWherefields)
                If Invoice.GetRecordType(vINVRS.Fields("record_type").Value) = Invoice.InvoiceRecordType.CreditNote AndAlso Not vInvoice.IsSundryCreditNote Then
                  '2b. delete IPH record
                  vWherefields = New CDBFields
                  vWherefields.Add("invoice_number", CDBField.FieldTypes.cftLong, .Fields(1).IntegerValue)
                  vWherefields.Add("batch_number", CDBField.FieldTypes.cftLong, .Fields(5).IntegerValue)
                  vWherefields.Add("transaction_number", CDBField.FieldTypes.cftInteger, .Fields(6).IntegerValue)
                  vWherefields.Add("line_number", CDBField.FieldTypes.cftInteger, .Fields(7).IntegerValue)
                  pEnv.Connection.DeleteRecords("invoice_payment_history", vWherefields)
                End If
              End If
            End While
            .CloseRecordSet()
          End With
          If pFHAdjustmentType = Batch.AdjustmentTypes.atNone Then
            If Not vCustomerUpdated AndAlso pBTA.LineType.Equals("U", StringComparison.InvariantCultureIgnoreCase) AndAlso pBTA.Amount < 0 Then
              Dim vOriginalPaymentInvoice As Invoice = Nothing
              Dim vOriginalPaymentIPH As InvoicePaymentHistory = BTAReversalOriginalIPHAndInvoice(pEnv, pBTA, vOriginalPaymentInvoice)

              If vInvoice IsNot Nothing AndAlso vOriginalPaymentInvoice IsNot Nothing AndAlso vOriginalPaymentIPH IsNot Nothing AndAlso
              vOriginalPaymentInvoice.SalesLedgerAccount.Equals(vInvoice.SalesLedgerAccount, StringComparison.InvariantCultureIgnoreCase) = False Then
                'BR21355: Where unallocated cash reversal for allocation belonging to a different sales ledger account (e.g. Contact A paid invoice on Contact B)
                'i) update credit customer outstanding for payer account only by remainder unallocated, as allocated amount will have already been added back
                'ii) update credit customer outstanding for recipient account for the full unallocated reversal amount
                Dim vOutstandingAllocationAmount As Double = FixTwoPlaces((pBTA.Amount * -1) - vOriginalPaymentIPH.Amount)
                If vOutstandingAllocationAmount > 0 Then
                  Dim vOriginalPayerCreditCustomer As New CreditCustomer()
                  vOriginalPayerCreditCustomer.InitCompanySalesLedgerAccount(pEnv, vInvoice.Company, vInvoice.SalesLedgerAccount)
                  vOriginalPayerCreditCustomer.AdjustOutstanding(vOutstandingAllocationAmount)
                  vOriginalPayerCreditCustomer.Save()
                End If

                Dim vOriginalPaymentCreditCustomer As New CreditCustomer()
                vOriginalPaymentCreditCustomer.InitCompanySalesLedgerAccount(pEnv, vOriginalPaymentInvoice.Company, vOriginalPaymentInvoice.SalesLedgerAccount)
                vOriginalPaymentCreditCustomer.AdjustOutstanding(vOriginalPaymentIPH.Amount)
                vOriginalPaymentCreditCustomer.Save()

                vCustomerUpdated = True
              End If
            End If
            '3. update credit customer record- only where a) line type not Invoice Allocation (L) or Sundry CN Invoice Allocation (K)
            ' and b) where the analysis line isn't a Sundry Credit Note reversal transaction (as the Credit Customers Outstanding amount will be updated through invoice creation in WriteInvoice())
            If vCustomerUpdated = False AndAlso Not (pBTA.LineType = "L" OrElse pBTA.LineType = "K") AndAlso Not (vInvoice.IsSundryCreditNote AndAlso pBTA.Amount > 0) Then
              vWherefields = New CDBFields
              vWherefields.Add("company", CDBField.FieldTypes.cftCharacter, .Fields("company").Value)
              vWherefields.Add("sales_ledger_account", CDBField.FieldTypes.cftCharacter, vSLAToUpdate)
              vUpdateFields = New CDBFields
              vUpdateFields.Add("outstanding", CDBField.FieldTypes.cftNumeric, "outstanding + " & System.Math.Abs(pBTA.Amount)) 'IS THIS +ve or -ve?
              vUpdateFields.AddAmendedOnBy(pEnv.User.UserID)
              pEnv.Connection.UpdateRecords("credit_customers", vUpdateFields, vWherefields)
            End If
            If Invoice.GetRecordType(vINVRS.Fields("record_type").Value) = Invoice.InvoiceRecordType.CreditNote AndAlso Not vInvoice.IsSundryCreditNote Then
              '4. check to see if any other IPH records exist for the was_batch/trans from #1
              '4a. Remove the Invoice Detail record for this payment.
              If vInvoice.InvoiceNumber.Length = 0 AndAlso .Fields("invoice_number").LongValue = 0 Then
                'Only remove this data if the credit note & invoice do not have a number
                vWherefields = New CDBFields
                vWherefields.Add("batch_number", CDBField.FieldTypes.cftLong, .Fields(1).IntegerValue)
                vWherefields.Add("transaction_number", CDBField.FieldTypes.cftLong, .Fields(2).IntegerValue)
                vWherefields.Add("line_number", CDBField.FieldTypes.cftLong, .Fields(3).IntegerValue)
                vWherefields.Add("invoice_number", CDBField.FieldTypes.cftLong, .Fields("invoice_number").IntegerValue)
                pEnv.Connection.DeleteRecords("invoice_details", vWherefields)
                If pEnv.Connection.GetCount("invoice_payment_history", Nothing, "batch_number = " & .Fields(1).IntegerValue & " AND transaction_number = " & .Fields(2).IntegerValue) = 0 Then
                  '4b. if none exist store the invoice # from #1 in a module-level variable...this is the invoice to delete later
                  If Len(pInvoicesToDelete) = 0 Then pInvoicesToDelete = ","
                  If InStr(pInvoicesToDelete, "," & .Fields("invoice_number").IntegerValue & ",") = 0 Then
                    pInvoicesToDelete = pInvoicesToDelete & .Fields("invoice_number").IntegerValue & ","
                  End If
                End If
              Else
                'Set vRemoved to False so that the Invoices record gets created when posting the Batch
                vRemoved = False
              End If
            ElseIf Not (pBTA.LineType = "L" OrElse pBTA.LineType = "K") AndAlso vINVRS.Fields("invoice_pay_status").Value <> pEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) Then
              '4. If the original C-type invoices record is not yet fully allocated update it to be so.
              'For U-type BTA lines the original C-type invoices record would not be fully allocated.
              'Whereas for N-type BTA lines the original C-type invoices record would already be fully allocated.
              If Not (pBTA.LineType = "N" AndAlso Invoice.GetRecordType(vINVRS.Fields("record_type").Value) = Invoice.InvoiceRecordType.SalesLedgerCash) Then
                'In this scenario, the AmountPaid was set when we created the payment so we do not need to increase it again?
                Dim vLineReversal As Nullable(Of Boolean)
                If pBatchType = Batch.BatchTypes.FinancialAdjustment AndAlso pBTA.LineType = "U" AndAlso vInvoice.InvoiceType = Invoice.InvoiceRecordType.SalesLedgerCash Then
                  'See if we have just reversed a single line, if so then do not update the invoice as it has already been done.
                  Dim vIDAnsiJoins As New AnsiJoins
                  With vIDAnsiJoins
                    .Add("financial_history_details fhd", "id.batch_number", "fhd.batch_number", "id.transaction_number", "fhd.transaction_number", "id.line_number", "fhd.line_number")
                    .Add("financial_history fh", "fhd.batch_number", "fh.batch_number", "fhd.transaction_number", "fh.transaction_number")
                    .AddLeftOuterJoin("reversals r", "fhd.batch_number", "r.was_batch_number", "fhd.transaction_number", "r.was_transaction_number", "fhd.line_number", "r.was_line_number")
                  End With
                  Dim vIDWhereFields As New CDBFields(New CDBField("id.batch_number", vInvoice.BatchNumber))
                  vIDWhereFields.Add("id.transaction_number", vInvoice.TransactionNumber)
                  Dim vIDSQLStatement As New SQLStatement(pEnv.Connection, "id.line_number AS id_line_number, fhd.amount, fhd.status As fhd_status, fh.status AS fh_status, r.batch_number, r.transaction_number, r.line_number", "invoice_details id", vIDWhereFields, "", vIDAnsiJoins)
                  Dim vIDRS As CDBRecordSet = vIDSQLStatement.GetRecordSet()
                  With vIDRS
                    While .Fetch
                      If .Fields("batch_number").Value.Length > 0 Then
                        If .Fields("batch_number").IntegerValue = pBTA.BatchNumber AndAlso .Fields("transaction_number").IntegerValue = pBTA.TransactionNumber Then
                          If (.Fields("line_number").IntegerValue = pBTA.LineNumber AndAlso .Fields("fh_status").Value = "A" AndAlso .Fields("fhd_status").Value = "R") Then
                            If vLineReversal.HasValue = False Then vLineReversal = True
                          Else
                            vLineReversal = False
                          End If
                        End If
                      End If
                    End While
                  End With
                End If
                If vLineReversal.HasValue = False Then vLineReversal = False
                If vLineReversal = False Then
                  vInvoice.InvoiceAmount = .Fields("invoice_amount").DoubleValue
                  vInvoice.SetAmountPaid(Math.Abs(pBTA.Amount))
                  vInvoice.Save(pEnv.User.UserID)
                End If
              End If
            End If
            'Set the Last Sales Ledger Account values
            vLastSLCompany = .Fields("company").Value
            vLastSLAccount = vSLAToUpdate
          End If
        End While
        .CloseRecordSet()
      End With
      Return vRemoved
    End Function

    Private Shared Function BTAReversalOriginalIPHAndInvoice(pEnv As CDBEnvironment, pBTA As BatchTransactionAnalysis, ByRef pInvoice As Invoice) As InvoicePaymentHistory
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("reversals r", "bta.batch_number", "r.batch_number", "bta.transaction_number", "r.transaction_number", "bta.line_number", "r.line_number")
      vAnsiJoins.Add("invoice_payment_history iph", "r.was_batch_number", "iph.batch_number", "r.was_transaction_number", "iph.transaction_number", "r.was_line_number", "iph.line_number")
      vAnsiJoins.Add("invoices i", "iph.invoice_number", "i.invoice_number")

      Dim vWhereFields As New CDBFields(New CDBField("bta.batch_number", pBTA.BatchNumber))
      vWhereFields.Add("bta.transaction_number", pBTA.TransactionNumber)
      vWhereFields.Add("bta.line_number", pBTA.LineNumber)

      Dim vSQLStatement As New SQLStatement(pEnv.Connection, "iph.batch_number,iph.transaction_number,iph.line_number,iph.invoice_number", "batch_transaction_analysis bta", vWhereFields, String.Empty, vAnsiJoins)

      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      Dim vOriginalIPH As InvoicePaymentHistory = Nothing
      If vRS.Fetch() Then
        vOriginalIPH = New InvoicePaymentHistory(pEnv)
        vOriginalIPH.InitFromBatchTransactionLine(vRS.Fields.Item("batch_number").IntegerValue,
                                                    vRS.Fields.Item("transaction_number").IntegerValue,
                                                    vRS.Fields.Item("line_number").IntegerValue)
        Dim vInvoice As New Invoice
        vInvoice.Init(pEnv, pInvoiceNumber:=vRS.Fields.Item("invoice_number").IntegerValue)
        pInvoice = vInvoice
      End If

      Return vOriginalIPH
    End Function
  End Class

End Namespace