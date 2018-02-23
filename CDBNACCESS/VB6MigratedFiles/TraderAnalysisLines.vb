Namespace Access

  Public Class TraderAnalysisLines
    Implements System.Collections.IEnumerable

    'This is a collection of TraderAnalysisLine objects
    Private mvCol As New Collection

    Friend Function Add(ByRef pKey As String) As TraderAnalysisLine
      'create a new object
      Dim vTDRLine As TraderAnalysisLine

      vTDRLine = New TraderAnalysisLine
      mvCol.Add(vTDRLine, pKey)
      Add = vTDRLine
      vTDRLine = Nothing
    End Function

    Public Sub AddItem(ByVal pTDRLine As TraderAnalysisLine, ByVal pIndexKey As String)
      mvCol.Add(pTDRLine, pIndexKey)
    End Sub

    Default Public ReadOnly Property Item(ByVal pIndexKey As String) As TraderAnalysisLine
      Get
        Item = CType(mvCol.Item(pIndexKey), TraderAnalysisLine)
      End Get
    End Property

    Default Public ReadOnly Property Item(ByVal pIndexKey As Integer) As TraderAnalysisLine
      Get
        Item = CType(mvCol.Item(pIndexKey), TraderAnalysisLine)
      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public Function Exists(ByVal pIndexKey As String) As Boolean
      Return mvCol.Contains(pIndexKey)
    End Function

    Friend Sub Remove(ByVal pIndexKey As String)
      mvCol.Remove(pIndexKey)
    End Sub

    Friend Sub Remove(ByVal pIndexKey As Integer)
      mvCol.Remove(pIndexKey)
    End Sub

    Public Sub Clear()
      mvCol = New Collection
    End Sub

    ''' <summary>Initialise collection of <see cref="TraderAnalysisLine">TraderAnalysisLine</see> items from a collection of <see cref="BatchTransactionAnalysis">BatchTransactionAnalysis</see> items for a Financial Adjustment Move.</summary>
    ''' <param name="pAnalysis"><see cref="CollectionList(Of BatchTransactionAnalysis)">CollectionList</see> of BatchTransactionAnalysis objects.</param>
    ''' <param name="pBatchType">Current batch type</param>
    ''' <param name="pPayerContactNumber">Contact number of the payer</param>
    ''' <param name="pPayerAddressNumber">Address number of the payer</param>
    ''' <param name="pBatchCurrencyCode">Batch currency code used to determine whether to use the amount or currency amount</param>
    ''' <param name="pTransDate">Date of the transaction</param>
    ''' <param name="pStockMovements"><see cref="CDBCollection">CDBCollection</see> of <see cref="StockMovement"> StockMovement's</see></param>
    ''' <param name="pPositiveTransOnly">True if the Move is for an adjusted transaction that contained sales ledger payments</param>
    Friend Sub InitAnalysisFromBTForMove(ByVal pAnalysis As CollectionList(Of BatchTransactionAnalysis), ByVal pBatchType As Batch.BatchTypes, ByVal pPayerContactNumber As Integer, ByVal pPayerAddressNumber As Integer, ByVal pBatchCurrencyCode As String, ByVal pTransDate As String, ByVal pStockMovements As CDBCollection, ByVal pPositiveTransOnly As Boolean)
      InitAnalysisFromBT(pAnalysis, pBatchType, Batch.AdjustmentTypes.atMove, "", pPayerContactNumber, pPayerAddressNumber, pBatchCurrencyCode, pTransDate, pStockMovements, False, pPositiveTransOnly)
    End Sub
    ''' <summary>Initialise collection of <see cref="TraderAnalysisLine">TraderAnalysisLine</see> items from a collection of <see cref="BatchTransactionAnalysis">BatchTransactionAnalysis</see> items.</summary>
    ''' <param name="pAnalysis"><see cref="CollectionList(Of BatchTransactionAnalysis)">CollectionList</see> of BatchTransactionAnalysis objects.</param>
    ''' <param name="pBatchType">Current batch type</param>
    ''' <param name="pFinancialAdjustment">Financial adjustment type</param>
    ''' <param name="pSundryCreditProductCode">Sundry credit product code used to set the <see cref="TraderAnalysisLine">TraderAnalysisLine</see> line type to be a sundry credit note</param>
    ''' <param name="pPayerContactNumber">Contact number of the payer</param>
    ''' <param name="pPayerAddressNumber">Address number of the payer</param>
    ''' <param name="pBatchCurrencyCode">Batch currency code used to determine whether to use the amount or currency amount</param>
    ''' <param name="pTransDate">Date of the transaction</param>
    ''' <param name="pStockMovements"><see cref="CDBCollection">CDBCollection</see> of <see cref="StockMovement"> StockMovement's</see></param>
    ''' <param name="pSundryCreditApplication">Is this from a sundry credit application</param>
    Public Sub InitAnalysisFromBT(ByVal pAnalysis As CollectionList(Of BatchTransactionAnalysis), ByVal pBatchType As Batch.BatchTypes, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, ByVal pSundryCreditProductCode As String, ByVal pPayerContactNumber As Integer, ByVal pPayerAddressNumber As Integer, ByVal pBatchCurrencyCode As String, ByVal pTransDate As String, ByVal pStockMovements As CDBCollection, ByVal pSundryCreditApplication As Boolean)
      InitAnalysisFromBT(pAnalysis, pBatchType, pFinancialAdjustment, pSundryCreditProductCode, pPayerContactNumber, pPayerAddressNumber, pBatchCurrencyCode, pTransDate, pStockMovements, pSundryCreditApplication, False)
    End Sub
    ''' <summary>Initialise collection of <see cref="TraderAnalysisLine">TraderAnalysisLine</see> items from a collection of <see cref="BatchTransactionAnalysis">BatchTransactionAnalysis</see> items.</summary>
    ''' <param name="pAnalysis"><see cref="CollectionList(Of BatchTransactionAnalysis)">CollectionList</see> of BatchTransactionAnalysis objects.</param>
    ''' <param name="pBatchType">Current batch type</param>
    ''' <param name="pFinancialAdjustment">Financial adjustment type</param>
    ''' <param name="pSundryCreditProductCode">Sundry credit product code used to set the <see cref="TraderAnalysisLine">TraderAnalysisLine</see> line type to be a sundry credit note</param>
    ''' <param name="pPayerContactNumber">Contact number of the payer</param>
    ''' <param name="pPayerAddressNumber">Address number of the payer</param>
    ''' <param name="pBatchCurrencyCode">Batch currency code used to determine whether to use the amount or currency amount</param>
    ''' <param name="pTransDate">Date of the transaction</param>
    ''' <param name="pStockMovements"><see cref="CDBCollection">CDBCollection</see> of <see cref="StockMovement"> StockMovement's</see></param>
    ''' <param name="pSundryCreditApplication">Is this from a sundry credit application</param>
    ''' <param name="pMovePositiveTransOnly">For AdjustmentType of atMove will only include positive analysis lines</param>
    Private Sub InitAnalysisFromBT(ByVal pAnalysis As CollectionList(Of BatchTransactionAnalysis), ByVal pBatchType As Batch.BatchTypes, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, ByVal pSundryCreditProductCode As String, ByVal pPayerContactNumber As Integer, ByVal pPayerAddressNumber As Integer, ByVal pBatchCurrencyCode As String, ByVal pTransDate As String, ByVal pStockMovements As CDBCollection, ByVal pSundryCreditApplication As Boolean, ByVal pMovePositiveTransOnly As Boolean)
      'Set up the TraderAnalysisLines from pBT
      Dim vBTA As BatchTransactionAnalysis
      Dim vTDRLine As TraderAnalysisLine
      Dim vLineType As TraderAnalysisLine.TraderAnalysisLineTypes
      Dim vLineNumber As Integer
      Dim vStockMovement As New StockMovement
      Dim vSMNumbers As String = ""
      Dim vStockTransID As Integer

      vLineNumber = 1
      mvCol = New Collection
      For Each vBTA In pAnalysis
        If ((pMovePositiveTransOnly = False) OrElse (pMovePositiveTransOnly = True AndAlso vBTA.Amount >= 0)) Then
          vTDRLine = Add(CStr(vLineNumber))

          'Set up the LineType
          vLineType = vTDRLine.GetAnalysisLineTypeFromCode(vBTA.LineType)
          If vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale Then
            If vBTA.ProductCode = pSundryCreditProductCode Or pSundryCreditApplication Then
              vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote
            ElseIf pFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation And pBatchType <> Batch.BatchTypes.SaleOrReturn Then
              vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit
            ElseIf pBatchType = Batch.BatchTypes.PostTaxPayrollGiving Then
              vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment
            ElseIf pBatchType = Batch.BatchTypes.GiveAsYouEarn Then
              vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
            End If
          End If

          'If the product is a stock product then determine which stock movements are linked to this BTA
          If vBTA.Product.StockItem Then
            vSMNumbers = ""
            For Each vStockMovement In pStockMovements
              If vStockMovement.LineNumber = vBTA.LineNumber Then
                If Len(vSMNumbers) > 0 Then vSMNumbers = vSMNumbers & ","
                vSMNumbers = vSMNumbers & CStr(vStockMovement.StockMovementNumber)
                If vStockMovement.TransactionID > 0 Then vStockTransID = vStockMovement.TransactionID
              End If
            Next vStockMovement
          End If

          vTDRLine.InitFromBTA(vBTA, vLineType, pFinancialAdjustment, pPayerContactNumber, pPayerAddressNumber, pBatchCurrencyCode, pTransDate, vSMNumbers, vStockTransID)
          vLineNumber = vLineNumber + 1
        End If
      Next vBTA
    End Sub
    Public Sub SaveAnalysis(ByVal pEnv As CDBEnvironment, ByVal pBT As BatchTransaction, ByRef pInvIssued As Integer, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, ByVal pExisting As Boolean, ByVal pUseSalesLedger As Boolean, ByVal pCreditSales As Boolean, ByVal pPayMethodsAtEnd As Boolean, Optional ByVal pCurruncyCode As String = "", Optional ByVal pExchangeRate As Double = 0, Optional ByVal pOrigBatchNumber As Integer = 0, Optional ByVal pOrigTransNumber As Integer = 0, Optional ByVal pUseStockTransactionID As Boolean = False, Optional ByVal pServiceBookingAnalysis As Boolean = False, Optional ByRef pEventBookingLinks As Boolean = False, Optional ByVal pLinkToFundraisingPayments As Boolean = False, Optional ByVal pProvisionalBatch As Boolean = False)
      'Create the BTA records from the TraderAnalysisLines in the collection
      'pInvIssued = Sum of the Issued values for Invoice Details that have been created - returned back to create the Invoice etc.
      Dim vBTA As BatchTransactionAnalysis
      Dim vOPH As OrderPaymentHistory
      Dim vPP As PaymentPlan = Nothing
      Dim vTDRLine As TraderAnalysisLine
      Dim vEA As EventAccommodationBooking
      Dim vEB As EventBooking = Nothing
      Dim vLB As LegacyBequest
      Dim vSB As ServiceBooking = Nothing
      Dim vCP As CollectionPayment
      Dim vPIS As CollectionPIS
      Dim vInvDetail As InvoiceDetail
      Dim vRS As CDBRecordSet
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vAmount As Double
      Dim vBoxAmounts() As String
      Dim vBoxNumber As Integer
      Dim vCreateLegRec As Boolean
      Dim vHoldContactNo As Integer
      Dim vIndex As Integer
      Dim vIssuedValue As String = ""
      Dim vLineNumber As Integer
      Dim vRows As Integer
      Dim vSMNumbers() As String
      Dim vStockIssued As Boolean
      Dim vStockProducts As Boolean
      Dim vQuantityValue As String
      Dim vSBLine As Integer
      Dim vLinkedAnalysis As Collection = Nothing
      Dim vAddBookingLink As Boolean
      Dim vParams As CDBParameters
      Dim vEBT As EventBookingTransaction

      If pExisting OrElse (pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation AndAlso pOrigBatchNumber > 0) Then
        '-----------------------------------------------------------------------------------------------------
        'BR15547: for CashBatchConfirmation delete any OPH line(s) for the provisional batch transaction
        'Else for an existing transaction, update PP and delete the BTA, OPH, Invoice Details & Collection Payments
        '-----------------------------------------------------------------------------------------------------
        vWhereFields = New CDBFields
        If pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation Then
          'BR15547: select provisional batch_number & transaction_number
          vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pOrigBatchNumber)
          vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pOrigTransNumber)
        Else
          vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
          vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        End If
        vPP = New PaymentPlan
        vOPH = New OrderPaymentHistory
        vPP.Init(pEnv)
        vOPH.Init(pEnv)
        'Update PP
        vRS = pEnv.Connection.GetRecordSet("SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & "  FROM order_payment_history oph WHERE " & pEnv.Connection.WhereClause(vWhereFields) & " ORDER BY line_number DESC")
        While vRS.Fetch() = True
          vOPH.InitFromRecordSet(pEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          If vOPH.OrderNumber <> vPP.PlanNumber Then vPP.Init(pEnv, (vOPH.OrderNumber))
          If vOPH.PaymentNumber = vPP.PaymentNumber Then
            vPP.PaymentNumber = vPP.PaymentNumber - 1
            vPP.SaveChanges()
          End If
          If pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation Then
            'BR15547: Reverse OPS payments for the provisional batch transaction
            Dim vOPS As New OrderPaymentSchedule()
            vOPS.Init(pEnv, IntegerValue(vOPH.ScheduledPaymentNumber))
            If vOPS.Existing Then
              vOPS.Reverse(vPP, vOPH.Amount, True)
              vOPS.ProcessPayment(False)
              vOPS.Save()
            End If
          End If
        End While
        vRS.CloseRecordSet()
        If pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation Then
          'BR15547: Delete OPH for the provisional batch transaction without errors
          pEnv.Connection.DeleteRecords("order_payment_history", vWhereFields, False)
        Else
          'Delete BTA, OPH, Invoice Details & Collection Payment without errors
          pEnv.Connection.DeleteRecordsMultiTable("batch_transaction_analysis,order_payment_history", vWhereFields)
          If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then pEnv.Connection.DeleteRecords("collection_payments", vWhereFields, False)
          If pCreditSales And pUseSalesLedger Then pEnv.Connection.DeleteRecords("invoice_details", vWhereFields, False)
          If pServiceBookingAnalysis And pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then pEnv.Connection.DeleteRecords("service_booking_transactions", vWhereFields, False)
          'If pEnv.GetDataStructureInfo(cdbDataEventMultipleAnalysis) Then pEnv.Connection.DeleteRecords "event_booking_transactions", vWhereFields, False

          '-------------------------------------------------------------------------
          'Delete any linked records
          '-------------------------------------------------------------------------
          If pLinkToFundraisingPayments AndAlso pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
            pEnv.Connection.DeleteRecords("fundraising_payment_history", vWhereFields, False)
          End If

          '-------------------------------------------------------------------------
          'Update Delivery Contact / Address if using the Holding Contact Number
          '-------------------------------------------------------------------------
          vHoldContactNo = IntegerValue(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber))
          If vHoldContactNo > 0 And pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
            vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, vHoldContactNo)
            If pEnv.Connection.GetCount("batch_transactions", vWhereFields) > 0 Then
              For Each vTDRLine In mvCol
                If vTDRLine.DeliveryContactNumber = vHoldContactNo Then vTDRLine.SetDeliveryContactAndAddress(pBT.ContactNumber, pBT.AddressNumber)
              Next vTDRLine
            End If
          End If
        End If
      End If

      Dim vFPHOrigBatchNo As Integer
      Dim vFPHOrigTransNo As Integer
      If pOrigBatchNumber > 0 AndAlso (pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation OrElse pFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation) Then
        vFPHOrigBatchNo = pOrigBatchNumber
        vFPHOrigTransNo = pOrigTransNumber
        If pLinkToFundraisingPayments = False Then
          'Where not linking to fundraising payment history clear the pOrigBatchNumber and pOrigTransNumber
          pOrigBatchNumber = 0
          pOrigTransNumber = 0
        End If
      End If

      '-------------------------------------------------------------------------
      'If no stock issued then re-set issued for P&P
      '-------------------------------------------------------------------------
      For Each vTDRLine In mvCol
        'Check to see if any stock was actually issued
        If vTDRLine.StockSale Then
          vStockProducts = True
          If CDbl(vTDRLine.Issued) > 0 Then vStockIssued = True
        End If
      Next vTDRLine

      If vStockProducts = True And vStockIssued = False Then
        'Set the P&P (if any) to 0 issued
        For Each vTDRLine In mvCol
          If vTDRLine.PostagePacking Then vTDRLine.ResetIssuedForPP()
        Next vTDRLine
      End If

      '-------------------------------------------------------------------------
      'Update the stock_movements record(s)
      '-------------------------------------------------------------------------
      If vStockProducts Then
        vUpdateFields = New CDBFields
        vWhereFields = New CDBFields
        With vUpdateFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
          .Add("line_number", CDBField.FieldTypes.cftLong)
        End With
        '(1) First update any records that have a comma-separated list of StockMovementNumbers (these come from Rich Client)
        vWhereFields.Add("stock_movement_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoInOrEqual)
        vLineNumber = 1
        For Each vTDRLine In mvCol
          If vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliHasStockMovement) Then
            If vTDRLine.StockMovementNumbers.Length > 0 Then
              vUpdateFields(3).Value = CStr(vLineNumber)
              vWhereFields(1).Value = vTDRLine.StockMovementNumbers
              vSMNumbers = Split(vTDRLine.StockMovementNumbers, ",")
              vRows = pEnv.Connection.UpdateRecords("stock_movements", vUpdateFields, vWhereFields)
              If vRows <> UBound(vSMNumbers) + 1 Then
                'If for some reason not all of the stock movements listed in vLines are updated then raise an error
                RaiseError(DataAccessErrors.daeDidNotUpdateAllStockMovements, CStr(vLineNumber), (vTDRLine.StockMovementNumbers))
              End If
            End If
          End If
          vLineNumber = vLineNumber + 1
        Next vTDRLine
        '(2) Second, update any records that have a TransactionID number (these come from Smart Client & Web Services)
        If pUseStockTransactionID Then
          vWhereFields.Clear()
          vWhereFields.Add("transaction_id", CDBField.FieldTypes.cftLong)
          vLineNumber = 1
          For Each vTDRLine In mvCol
            If vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliHasStockMovement) Then
              If vTDRLine.StockTransactionID > 0 Then
                vUpdateFields(3).Value = CStr(vLineNumber)
                vWhereFields(1).Value = CStr(vTDRLine.StockTransactionID)
                vRows = pEnv.Connection.UpdateRecords("stock_movements", vUpdateFields, vWhereFields)
              End If
            End If
            vLineNumber = vLineNumber + 1
          Next vTDRLine
        End If
      End If

      '------------------------------------------------------------------------
      'Create the BTA
      '------------------------------------------------------------------------
      If Not pEnv.Connection.InTransaction Then pEnv.Connection.StartTransaction()

      vLineNumber = 1
      Dim vSBTranExisting As Boolean
      Dim vProduct As New Product(pEnv)
      Dim vLinksDeleted As Boolean
      Dim vCreatesBTA As Boolean
      Dim vSetContactNumber As Boolean
      Dim vExamBooking As New ExamBooking(pEnv)

      For Each vTDRLine In mvCol
        vAddBookingLink = False
        vIssuedValue = ""
        vBTA = New BatchTransactionAnalysis(pEnv)
        vCreatesBTA = vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliCreatesBTA)
        If pFinancialAdjustment = Batch.AdjustmentTypes.atNone OrElse pFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment Then
          vSetContactNumber = True
        Else
          vSetContactNumber = Not (vCreatesBTA AndAlso vTDRLine.DeliveryContactNumber = 0)
        End If
        vBTA.InitFromTransaction(pBT, vLineNumber, vSetContactNumber)
        With vBTA
          If vCreatesBTA Then
            Select Case vTDRLine.TraderLineType
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent, TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingEntitlement, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote, TraderAnalysisLine.TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
                .LineType = "P"
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment
                .LineType = If(vTDRLine.DeceasedContactNumber > 0, "S", "P")
              Case Else
                .LineType = vTDRLine.TraderLineTypeCode
            End Select
            .ProductCode = vTDRLine.ProductCode
            .RateCode = vTDRLine.RateCode
            .DistributionCode = vTDRLine.DistributionCode
            vQuantityValue = ""
            If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBTAQuantityDecimal) Then
              vQuantityValue = FixedFormat(vTDRLine.Quantity)
              If vTDRLine.Issued.Length > 0 Then vIssuedValue = FixedFormat(Val(vTDRLine.Issued))
            Else
              vQuantityValue = CStr(vTDRLine.Quantity)
              If vTDRLine.Issued.Length > 0 Then vIssuedValue = CStr(CInt(Val(vTDRLine.Issued)))
            End If
            .Quantity = CInt(vQuantityValue)
            If vIssuedValue.Length > 0 Then
              .Issued = CInt(vIssuedValue)
            ElseIf vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit Then
              vIssuedValue = CStr(vTDRLine.Quantity)
            End If
            Select Case vTDRLine.TraderLineType
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltDeceased, TraderAnalysisLine.TraderAnalysisLineTypes.taltSoftCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltInMemoriamSoftCredit
                If vTDRLine.PaymentPlanNumber > 0 Then .PaymentPlanNumber = vTDRLine.PaymentPlanNumber
                If vTDRLine.DeceasedContactNumber > 0 Then .DeceasedContactNumber = vTDRLine.DeceasedContactNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltInMemoriamHardCredit
                If vTDRLine.GiverContactNumber > 0 Then
                  .DeceasedContactNumber = vTDRLine.GiverContactNumber
                  If vTDRLine.DeceasedContactNumber > 0 Then .MemberNumber = CStr(vTDRLine.DeceasedContactNumber)
                  If vTDRLine.PaymentPlanNumber > 0 Then .PaymentPlanNumber = vTDRLine.PaymentPlanNumber
                Else
                  If vTDRLine.DeceasedContactNumber > 0 Then .DeceasedContactNumber = vTDRLine.DeceasedContactNumber
                End If
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership
                .MemberNumber = vTDRLine.MemberNumber
                If vTDRLine.PaymentPlanNumber > 0 Then .PaymentPlanNumber = vTDRLine.PaymentPlanNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant
                If vTDRLine.CovenantNumber > 0 Then .CovenantNumber = vTDRLine.CovenantNumber
                If vTDRLine.PaymentPlanNumber > 0 Then .PaymentPlanNumber = vTDRLine.PaymentPlanNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan
                If vTDRLine.PaymentPlanNumber > 0 Then .PaymentPlanNumber = vTDRLine.PaymentPlanNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, _
                   TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
                Select Case vTDRLine.TraderLineType
                  Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
                    If vTDRLine.InvoiceNumber > 0 Then .InvoiceNumber = vTDRLine.InvoiceNumber
                End Select
                'Use Member Number to store the Sales Ledger Account used for invoice payments
                vBTA.MemberNumber = vTDRLine.SalesLedgerAccount
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltIncentive
                If vTDRLine.IncentiveLineNumber > 0 Then .PaymentPlanNumber = vTDRLine.IncentiveLineNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
                'Use Member Number to store Post or Pre Tax Payroll Giving Pledge number
                If vTDRLine.PayrollGivingPledgeNumber > 0 Then .MemberNumber = CStr(vTDRLine.PayrollGivingPledgeNumber)
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment
                If vTDRLine.DeceasedContactNumber > 0 Then .DeceasedContactNumber = vTDRLine.DeceasedContactNumber
              Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEventPricingMatrixLine, TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
                'use Member Number to store the Event Booking Number
                If vTDRLine.EventBookingNumber > 0 Then vBTA.MemberNumber = CStr(vTDRLine.EventBookingNumber)
            End Select
            .Source = vTDRLine.Source
            If vTDRLine.GrossAmount.Length > 0 Then .GrossAmount = CStr(Val(vTDRLine.GrossAmount))
            If vTDRLine.Discount.Length > 0 Then .Discount = CStr(Val(vTDRLine.Discount))
            .CurrencyAmount = vTDRLine.Amount 'Amount in the currency for the Batch
            If Len(vTDRLine.ProductCode) > 0 And vTDRLine.ProductCode = pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRoundingErrorProduct) Then
              .Amount = 0
            Else
              'Amount calculated to the Base Currency
              .Amount = CalculateCurrencyAmount(vTDRLine.Amount, pCurruncyCode, pExchangeRate, True)
            End If
            If vTDRLine.AcceptAsFull Then .AcceptAsFull = True
            .WhenValue = vTDRLine.LineDate
            vBTA.DespatchMethod = vTDRLine.DespatchMethod
            If vTDRLine.DeliveryContactNumber > 0 Then
              .ContactNumber = vTDRLine.DeliveryContactNumber
              .AddressNumber = vTDRLine.DeliveryAddressNumber
            End If
            .VatRate = vTDRLine.VatRate
            If vTDRLine.Quantity > 0 Or (vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking Or vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit) Then
              .CurrencyVatAmount = vTDRLine.VatAmount 'VAT Amount in the currency for the Batch
              .VatAmount = CalculateCurrencyAmount(vTDRLine.VatAmount, pCurruncyCode, pExchangeRate, True) 'VAT Amount calculated to the Base Currency
            End If
            If vTDRLine.SalesContactNumber > 0 Then
              .SalesContactNumber = vTDRLine.SalesContactNumber
            End If
            .Notes = vTDRLine.Notes
            If pProvisionalBatch AndAlso vTDRLine.ScheduledPaymentNumber > 0 Then
              'BR15547: For the provisional BTA set the Scheduled payment number in the notes such that if reversing confirmed, in FinancialHistory.Reverse 
              'we will re-create the OPH against the correct OPS record
              Select Case vTDRLine.TraderLineType
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership, TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant, TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan
                  If .Notes.Length > 0 Then .Notes += ", "
                  .Notes += String.Format("Scheduled Payment Number: {0}", vTDRLine.ScheduledPaymentNumber)
              End Select
            End If
            If vTDRLine.ProductNumber > 0 Then .ProductNumber = vTDRLine.ProductNumber
            If vTDRLine.WarehouseCode.Length > 0 Then .Warehouse = vTDRLine.WarehouseCode
            If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInMemoriamHardCredit OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInMemoriamSoftCredit Then
              'For LineTypes D (InMemoriam HardCredit) & F (InMemoriam SoftCredit) set the BTA ContactNumber to be the CreditedContactNumber
              If vTDRLine.CreditedContactNumber > 0 Then
                .ContactNumber = vTDRLine.CreditedContactNumber
                .AddressNumber = vTDRLine.CreditedContactDefaultAddressNumber(pEnv)
              End If
            End If
            .Save()
          End If
          pBT.Analysis.Add(vLineNumber.ToString & .LineType & .ProductCode, vBTA)
        End With

        '------------------------------------------------------------------------
        'Save any additional items
        '------------------------------------------------------------------------
        Select Case vTDRLine.TraderLineType
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation
            If pExisting = False And pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
              vEA = New EventAccommodationBooking
              vEA.SetTransactionInfo(pEnv, vTDRLine.RoomBookingNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vTDRLine.SalesContactNumber, (pPayMethodsAtEnd = True And pCreditSales = True))
              vEA.Save()
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
            If pExisting = False Then
              Select Case pFinancialAdjustment
                Case Batch.AdjustmentTypes.atCashBatchConfirmation, Batch.AdjustmentTypes.atEventAdjustment,
                     Batch.AdjustmentTypes.atMove, Batch.AdjustmentTypes.atNone
                  vEB = New EventBooking
                  'This should really be the actual event booking object that was created but we don't have it
                  vEB.SetTransactionInfo(pEnv, vTDRLine.EventBookingNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vTDRLine.SalesContactNumber, vTDRLine.EventNumber, (pPayMethodsAtEnd = True And pCreditSales = True))
                  vEB.Save()
              End Select
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltExamBooking
            If pExisting = False And (pFinancialAdjustment = Batch.AdjustmentTypes.atNone Or pFinancialAdjustment = Batch.AdjustmentTypes.atMove Or pFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment Or pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment) Then
              If vTDRLine.ExamUnitProductId > 0 Then
                Dim vExamBookingTransaction As New ExamBookingTransaction(pEnv)
                vExamBookingTransaction.SetTransactionInfo(vTDRLine.ExamBookingNumber, vTDRLine.ExamUnitId, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber)
              Else
                Dim vExamBookingUnit As New ExamBookingUnit(pEnv)
                vExamBookingUnit.SetTransactionInfo(vTDRLine.ExamBookingNumber, vTDRLine.ExamUnitId, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber)
              End If
              If vExamBooking.Existing = False OrElse vExamBooking.ExamBookingId <> vTDRLine.ExamBookingNumber Then vExamBooking.Init(vTDRLine.ExamBookingNumber)
              If vExamBooking.Existing AndAlso vExamBooking.BatchNumber = 0 Then
                vExamBooking.SetTransactionInfo(pBT.BatchNumber, pBT.TransactionNumber)
                vExamBooking.Save()
              End If
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit
            vSB = New ServiceBooking
            If pFinancialAdjustment = Batch.AdjustmentTypes.atNone And Not pExisting Then
              vSB.SetTransactionInfo(pEnv, vTDRLine.ServiceBookingNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vTDRLine.SalesContactNumber)
              vSB.Save()
              vSBLine = vLineNumber
            Else
              vSB.Init(pEnv, (vTDRLine.ServiceBookingNumber))
              If vSB.Existing Then vSBLine = vLineNumber
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan, TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership, TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant, _
               TraderAnalysisLine.TraderAnalysisLineTypes.taltSoftCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit
            If vBTA.PaymentPlanNumber > 0 And vTDRLine.ScheduledPaymentNumber > 0 Then
              If vPP Is Nothing Then
                vPP = New PaymentPlan
                vPP.Init(pEnv)
              End If
              If vPP.PlanNumber <> vBTA.PaymentPlanNumber Then
                vPP.Init(pEnv, (vBTA.PaymentPlanNumber))
              End If
              vPP.PaymentNumber = vPP.PaymentNumber + 1
              vPP.SaveChanges()
              vOPH = New OrderPaymentHistory
              vOPH.Init(pEnv)
              vOPH.SetValues((pBT.BatchNumber), (pBT.TransactionNumber), (vPP.PaymentNumber), (vPP.PlanNumber), (vBTA.Amount), vLineNumber, 0, vTDRLine.ScheduledPaymentNumber)
              vOPH.Save()
              If pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation Then
                'BR15547: For the confirmed order payment line re-create the OPS as this is reversed for the provisional transaction (see above in TraderAnalysisLines.SaveAnalysis)
                Dim vOPS As New OrderPaymentSchedule()
                vOPS.Init(pEnv, IntegerValue(vOPH.ScheduledPaymentNumber))
                If vOPS.Existing Then
                  vOPS.SetUnProcessedPayment(True, vOPH.Amount)
                  vOPS.Save()
                End If
              End If
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltLegacyBequestReceipt
            vCreateLegRec = True
            If pExisting Then
              vWhereFields = New CDBFields
              With vWhereFields
                .Add("legacy_number", CDBField.FieldTypes.cftLong, vTDRLine.LegacyNumber)
                .Add("bequest_number", CDBField.FieldTypes.cftLong, vTDRLine.BequestNumber)
                .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
                .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
                .Add("line_number", CDBField.FieldTypes.cftLong, vLineNumber)
                .Add("amount", CDBField.FieldTypes.cftNumeric, vTDRLine.Amount)
              End With
              If pEnv.Connection.GetCount("legacy_bequest_receipts", vWhereFields) > 0 Then vCreateLegRec = False
            End If

            If pFinancialAdjustment <> Batch.AdjustmentTypes.atNone And pOrigBatchNumber > 0 Then
              '1. We only want to create the new Receipt if this Bequest payment
              '   was not in original FA Batch
              '2. We may not have initialised the Legacy Bequest
              vWhereFields = New CDBFields
              With vWhereFields
                .Add("bta.batch_number", CDBField.FieldTypes.cftLong, pOrigBatchNumber)
                .Add("bta.transaction_number", CDBField.FieldTypes.cftLong, pOrigTransNumber)
                .Add("line_type", CDBField.FieldTypes.cftCharacter, "B")
                .Add("lbr.batch_number", CDBField.FieldTypes.cftLong, "bta.batch_number")
                .Add("lbr.transaction_number", CDBField.FieldTypes.cftLong, "bta.transaction_number")
                .Add("lbr.line_number", CDBField.FieldTypes.cftLong, "bta.line_number")
              End With
              vRS = pEnv.Connection.GetRecordSet("SELECT product, rate, legacy_number, bequest_number, lbr.amount, date_received FROM batch_transaction_analysis bta, legacy_bequest_receipts lbr WHERE " & pEnv.Connection.WhereClause(vWhereFields))
              While vRS.Fetch() = True
                With vTDRLine
                  If vRS.Fields("legacy_number").IntegerValue = .LegacyNumber And vRS.Fields("bequest_number").IntegerValue = .BequestNumber And vRS.Fields("amount").DoubleValue = .Amount And CDate(vRS.Fields("date_received").Value) = CDate(.LineDate) And vRS.Fields("product").Value = .ProductCode And vRS.Fields("rate").Value = .RateCode Then
                    vCreateLegRec = False
                  End If
                End With
              End While
              vRS.CloseRecordSet()
            End If

            If vCreateLegRec Then
              vLB = New LegacyBequest(pEnv)
              vLB.Init((vTDRLine.LegacyNumber), (vTDRLine.BequestNumber))
              vLB.AddReceipt(pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vTDRLine.Amount, CDate(vTDRLine.LineDate).ToString(CAREDateFormat), vTDRLine.Notes)
            End If

          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment
            vCP = New CollectionPayment
            vCP.Init(pEnv)
            If vTDRLine.CollectionBoxNumbers.Length > 0 Then
              vSMNumbers = Split(vTDRLine.CollectionBoxNumbers, ",")
              vBoxAmounts = Split(vTDRLine.CollectionBoxAmounts, ",")
              If UBound(vSMNumbers) > UBound(vBoxAmounts) Then ReDim Preserve vBoxAmounts(UBound(vSMNumbers))

              If UBound(vSMNumbers) = 0 Then
                'Just the one box
                vBoxNumber = IntegerValue(vSMNumbers(0))
                vAmount = vTDRLine.Amount
                vCP.Create(vTDRLine.CollectionNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vAmount, (vTDRLine.CollectionPisNumber), vBoxNumber)
                vCP.Save()
              Else
                'Multiple Boxes
                For vIndex = 0 To UBound(vSMNumbers)
                  vBoxNumber = IntegerValue(vSMNumbers(vIndex))
                  vAmount = Val(vBoxAmounts(vIndex))
                  vCP = New CollectionPayment
                  vCP.Init(pEnv)
                  vCP.Create(vTDRLine.CollectionNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vAmount, (vTDRLine.CollectionPisNumber), vBoxNumber)
                  vCP.Save()
                Next
              End If
            Else
              vCP.Create(vTDRLine.CollectionNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, vTDRLine.Amount, (vTDRLine.CollectionPisNumber))
              vCP.Save()
            End If
            If vCP.CollectionPisNumber > 0 Then
              vPIS = New CollectionPIS
              vPIS.Init(pEnv, (vCP.CollectionPisNumber))
              vPIS.Reconcile(CollectionPIS.CollectionPISReconciledStatus.cpisrsTrader)
              vPIS.Save()
            End If

          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale
            'Handle ServiceBooking links
            If pServiceBookingAnalysis And pFinancialAdjustment <> Batch.AdjustmentTypes.atNone And (vTDRLine.ServiceBookingNumber > 0 Or (vSBLine > 0 And (vSBLine <> vLineNumber))) Then
              vLinkedAnalysis = New Collection
              If vTDRLine.ServiceBookingNumber > 0 Then
                vBTA.LinkedBookingNo = vTDRLine.ServiceBookingNumber
              Else
                vBTA.LinkedBookingNo = vSB.ServiceBookingNumber
              End If
              vLinkedAnalysis.Add(vBTA)
            End If
            If pServiceBookingAnalysis And pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then
              If (vSBLine > 0 And (vSBLine <> vLineNumber)) Or vTDRLine.ServiceBookingNumber > 0 Then
                If vSBLine = 0 Then 'Linked product sale analysis line
                  vSB = New ServiceBooking
                  vSB.Init(pEnv, (vTDRLine.ServiceBookingNumber))
                End If
                If vSB.Existing Then
                  vSB.SetTransactionInfo(pEnv, vSB.ServiceBookingNumber, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, 0)
                  vWhereFields = New CDBFields
                  vWhereFields.Add("service_booking_number", CDBField.FieldTypes.cftLong, vSB.ServiceBookingNumber)
                  vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, vSB.BatchNumber)
                  vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, vSB.TransactionNumber)
                  vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, vSB.LineNumber)
                  vSBTranExisting = pEnv.Connection.GetCount("service_booking_transactions", vWhereFields) > 0
                  If Not vSBTranExisting Then
                    vProduct.Init((vTDRLine.ProductCode))
                    If Not vProduct.Donation Then vSB.AddLinkedTransaction(pFinancialAdjustment, vLinkedAnalysis, Nothing, vLineNumber)
                  End If
                End If
              End If
            End If
            'Handle EventBooking links
            If pEventBookingLinks = True And pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
              If vTDRLine.EventBookingNumber > 0 Then
                vAddBookingLink = True
                If Not (vEB Is Nothing) Then
                  If vEB.LineNumber < vBTA.LineNumber Then vAddBookingLink = False 'Line is already linked to a Booking, so use that one only
                End If
                If vAddBookingLink Then
                  vProduct = New Product(pEnv)
                  vProduct.Init((vTDRLine.ProductCode))
                  vAddBookingLink = Not (vProduct.Donation)
                End If
                vEB = New EventBooking
                vEB.Init(pEnv, 0, (vTDRLine.EventBookingNumber))
                If vEB.Existing Then
                  vParams = New CDBParameters
                  With vParams
                    .Add("EventNumber", vEB.EventNumber)
                    .Add("BookingNumber", vEB.BookingNumber)
                    .Add("BatchNumber", vBTA.BatchNumber)
                    .Add("TransactionNumber", vBTA.TransactionNumber)
                    .Add("LineNumber", vBTA.LineNumber)
                  End With
                  vEBT = New EventBookingTransaction
                  vEBT.Init(pEnv)
                  vEBT.Create(vParams)
                  vEBT.Save(pEnv.User.Logname, False)
                End If
              End If
            End If

          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote
            'Handle EventBooking links
            If pEventBookingLinks = True And pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
              If vTDRLine.EventBookingNumber > 0 Then
                vAddBookingLink = True
                If Not (vEB Is Nothing) Then
                  If vEB.LineNumber < vBTA.LineNumber Then vAddBookingLink = False 'Line is already linked to a Booking, so use that one only
                End If
                If vAddBookingLink Then
                  vProduct = New Product(pEnv)
                  vProduct.Init(vTDRLine.ProductCode)
                  vAddBookingLink = Not (vProduct.Donation)
                End If
                vEB = New EventBooking
                vEB.Init(pEnv, 0, vTDRLine.EventBookingNumber)
                If vEB.Existing Then
                  vParams = New CDBParameters
                  With vParams
                    .Add("EventNumber", vEB.EventNumber)
                    .Add("BookingNumber", vEB.BookingNumber)
                    .Add("BatchNumber", vBTA.BatchNumber)
                    .Add("TransactionNumber", vBTA.TransactionNumber)
                    .Add("LineNumber", vBTA.LineNumber)
                  End With
                  vEBT = New EventBookingTransaction
                  vEBT.Init(pEnv)
                  vEBT.Create(vParams)
                  vEBT.Save(pEnv.User.Logname, False)
                End If
              End If
            End If

          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
            If (pFinancialAdjustment = Batch.AdjustmentTypes.atNone AndAlso vBTA.Amount > 0) OrElse ((pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment OrElse pFinancialAdjustment = Batch.AdjustmentTypes.atReverse) AndAlso vBTA.Amount < 0) Then
              'Now that Trader will generate two L-Type lines on the TAS grid when allocating SL cash to an invoice:
              '(1) When allocating the cash only create IPH for the +ve BTA line
              '(2) When adjusting the allocation only create IPH for the -ve BTA line
              Dim vFHD As New FinancialHistoryDetail(pEnv)
              vFHD.InitFromInvoiceNumber(vTDRLine.InvoiceNumberUsed)
              If vFHD.Existing Then
                vParams = New CDBParameters
                With vParams
                  .Add("InvoiceNumber", vBTA.InvoiceNumber)  'vTDRLine.InvoiceNumber
                  .Add("BatchNumber", vFHD.BatchNumber)
                  .Add("TransactionNumber", vFHD.TransactionNumber)
                  .Add("LineNumber", vFHD.LineNumber)
                  .Add("Amount", vBTA.Amount)
                  If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) Then
                    .Add("AllocationDate", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
                    .Add("AllocationBatchNumber", pBT.BatchNumber)
                    .Add("AllocationTransactionNumber", pBT.TransactionNumber)
                    .Add("AllocationLineNumber", vLineNumber)
                  End If
                End With
                Dim vIPH As New InvoicePaymentHistory(pEnv)
                vIPH.Create(vParams)
                vIPH.Save()
              End If
            End If

        End Select

        If vTDRLine.ScheduledPaymentNumber > 0 AndAlso vTDRLine.PaymentPlanNumber = 0 Then
          If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit _
          OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale _
          OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSoftCredit _
          OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltDeceased Then
            Dim vFPH As New FundraisingPaymentHistory(pEnv)
            vFPH.Init()
            If pLinkToFundraisingPayments Then
              'Delete links to provisional Batch and create link to this payment
              vFPH.CreateNewLink(vTDRLine.ScheduledPaymentNumber, vBTA.BatchNumber, vBTA.TransactionNumber, vBTA.LineNumber, vFPHOrigBatchNo, vFPHOrigTransNo)
            ElseIf vLinksDeleted = False Then
              'Delete links to provisional Batch
              vFPH.DeleteLinks(vFPHOrigBatchNo, vFPHOrigTransNo)
            End If
            vFPHOrigBatchNo = 0
            vFPHOrigTransNo = 0
            vLinksDeleted = True
          End If
        End If

        'Add any Invoice Details
        If pCreditSales And pUseSalesLedger And Len(vIssuedValue) > 0 And Not vStockProducts Then
          If Val(vIssuedValue) > 0 Then
            pInvIssued = pInvIssued + IntegerValue(vIssuedValue)
            vInvDetail = New InvoiceDetail
            vInvDetail.Create(pEnv, pBT.BatchNumber, pBT.TransactionNumber, vLineNumber, 0)
            vInvDetail.Save()
          End If
        End If

        '------------------------------------------------------------------------
        'Update any Incentive lines
        '------------------------------------------------------------------------
        SetPaymentIncentive(vTDRLine.LineNumber, vLineNumber)

        vLineNumber = vLineNumber + 1
      Next

      '------------------------------------------------------------------------
      'Finally update the Batch Transaction
      '------------------------------------------------------------------------
      pBT.NextLineNumber = vLineNumber

    End Sub

    Friend Sub UpdateVATRates(ByVal pEnv As CDBEnvironment, ByVal pTransactionDate As String)
      'TransactionDate has been changed so update the current Analysis lines
      Dim vTDRLine As TraderAnalysisLine
      Dim vVAT As VatRate
      Dim vPercentage As Double

      For Each vTDRLine In mvCol
        Select Case vTDRLine.TraderLineType
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale, TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent, TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation, TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment
            vVAT = pEnv.VATRate((vTDRLine.VatRate), "")
            vPercentage = vVAT.CurrentPercentage(pTransactionDate)
            If CDbl(vTDRLine.VATPercentage) <> vPercentage Then
              vTDRLine.UpdateVAT(vPercentage)
            End If
        End Select
      Next vTDRLine

    End Sub

    Private Sub SetPaymentIncentive(ByVal pOrigLineNumber As Integer, ByVal pBTALineNumber As Integer)
      'Find any Incentives that relate to the current line (pOrigLineNumber)
      'And for each one found, update to link to the new BTA line (pBTALineNumber)
      Dim vTDRLine As TraderAnalysisLine

      For Each vTDRLine In mvCol
        If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltIncentive Then
          If vTDRLine.IncentiveLineNumber = pOrigLineNumber Then
            'Update the IncentiveLineNumber to be pBTALineNumber
            vTDRLine.SetNewIncentiveLineNumber(pBTALineNumber)
          End If
        End If
      Next vTDRLine

    End Sub

    Public Sub SetDepositAllowed(ByVal pEnv As CDBEnvironment)
      'Used to set the TALine Deposit Allowed flag which will be used on the client-side to calculate the CSDepositAmount
      Dim vNoDepositAllowed As Boolean = False

      For Each vTDRLine As TraderAnalysisLine In mvCol
        Select Case vTDRLine.TraderLineType
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking
            If Not vNoDepositAllowed Then
              Dim vSB As New ServiceBooking
              vSB.Init(pEnv, vTDRLine.ServiceBookingNumber)
              If vSB.Existing Then
                'If the transaction contains a service booking where the booking date is within the late notification period, then force payment of full amount.
                Dim vDate As Date = DateAdd(DateInterval.Day, vSB.ServiceControl.LateBookingNotificationDays * -1, DateValue(vSB.StartDate))
                vNoDepositAllowed = DateDiff(DateInterval.Day, DateValue(vSB.TransactionDate), vDate) <= 0
              Else
                'This transaction has a service booking analysis line that is linked to a non-existent service booking
              End If
            End If
            vTDRLine.SetDepositAllowed(Not vNoDepositAllowed)
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale
            'If the transaction contains a stock sale, then force payment of full amount.
            If Not vNoDepositAllowed Then vNoDepositAllowed = vTDRLine.StockSale
            vTDRLine.SetDepositAllowed(Not vNoDepositAllowed)
        End Select
      Next vTDRLine
    End Sub

  End Class
End Namespace
