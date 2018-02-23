Imports Advanced.LanguageExtensions.EnumerableExtensions

Namespace Access
  Public Class FinancialHistory

    Public Enum FinancialHistoryRecordSetTypes 'These are bit values
      fhrtAll = &HFFFFS
      'ADD additional recordset types here
      fhrtNumbers = 1
      fhrtDetail = 2
      fhrtDetailLines = 4
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum FinancialHistoryFields
      fhfAll = 0
      fhfBatchNumber
      fhfTransactionNumber
      fhfContactNumber
      fhfTransactionDate
      fhfTransactionType
      fhfBankDetailsNumber
      fhfAmount
      fhfPaymentMethod
      fhfReference
      fhfPosted
      fhfAddressNumber
      fhfNotes
      fhfStatus
      fhfCurrencyAmount
      fhfTransactionOrigin
    End Enum

    Public Enum FinancialHistoryStatus
      fhsNormal 'Null    Normal value
      fhsAdjusted 'A       Transaction has been adjusted
      fhsOnBackOrder 'B       Transaction has been placed on back order
      fhsMoved 'M       Transaction has been moved to another contact
      fhsReversed 'R       Transaction has been reversed (refunded)
    End Enum

    Public Enum AdjustmentStates
      adjsNone
      adjsIsAnAdjustment
      adjsHasBeenAdjusted
      adjsIsAnAdjustmentAndHasBeenAdjusted
    End Enum

    Public Enum SalesLedgerItems
      None = 0
      InvoicePayments = 1
      UnallocatedSLCash = 2
      SLCashAllocation = 4
      CreditNoteAllocation = 8
    End Enum

    Private Const FA_STATUS_ADJUSTMENT As String = "A"
    Private Const FA_STATUS_MOVE As String = "M"
    Private Const FA_STATUS_REVERSAL As String = "R"
    Private Const FA_STATUS_BACK_ORDER As String = "B"
    Private Const FA_STATUS_IN_ADVANCE As String = "I" 'used only on order payment history records
    Private Const FA_STATUS_IN_ADVANCE_USED As String = "B" 'used only on order payment history records

    'Other Class Variables
    Private mvDetails As Collection
    Private mvDetail As FinancialHistoryDetail
    Private mvContactAccount As ContactAccount
    Private mvNextLineNumber As Integer
    Private mvConfirmedTransChecked As Boolean
    Private mvIsConfirmedTransaction As Boolean
    Private mvTransactionSign As String
    Private mvNegativesAllowed As String

    'Used for deciding which Smart Client menu items should be available
    Private mvBatchType As Batch.BatchTypes
    Private mvGotBatchType As Boolean
    Private mvIsEventBooking As Boolean
    Private mvCheckedForEventBooking As Boolean
    Private mvContainsExamBooking As Boolean
    Private mvCheckedForExamBooking As Boolean
    Private mvContainsInMemoriamPPPayment As Boolean
    Private mvCheckedForInMemoriamPPPayment As Boolean
    Private mvAdjustmentState As FinancialHistory.AdjustmentStates
    Private mvAdjustmentStateSet As Boolean
    Private mvAdjustmentBatchNumber As Integer
    Private mvAdjustmentTransactionNumber As Integer
    Private mvAdjustmentWasBatchNumber As Integer
    Private mvAdjustmentWasTransactionNumber As Integer
    Private mvContainsSalesLedgerCashAllocation As Boolean
    Private mvCheckedForSalesLedgerCashAllocation As Boolean

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    'Private class to handle the part-refund data
    Private Class PartRefundAnalysis
      Public Property RefundQuantity As Integer = 0
      Public Property RefundIssued As Integer = 0
      Public Property SalesLedgerPartRefund As Boolean = False
      Public Property InvoicePaymentAmount As Double = 0
      Public Property UnallocatedAmount As Double = 0
      Public Property RefundAmount As Double = 0
    End Class

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "financial_history"
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_date", CDBField.FieldTypes.cftDate)
          .Add("transaction_type")
          .Add("bank_details_number", CDBField.FieldTypes.cftLong)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("payment_method")
          .Add("reference")
          .Add("posted", CDBField.FieldTypes.cftDate)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("status")
          .Add("currency_amount", CDBField.FieldTypes.cftNumeric)
          .Add("transaction_origin")
        End With

        mvClassFields.Item(FinancialHistoryFields.fhfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(FinancialHistoryFields.fhfPosted).PrefixRequired = True

        mvClassFields.Item(FinancialHistoryFields.fhfCurrencyAmount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode)
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionOrigin).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataTransactionOrigins)
      Else
        mvClassFields.ClearItems()
      End If

      mvDetails = Nothing
      mvDetails = New Collection
      mvDetail = Nothing
      mvDetail = New FinancialHistoryDetail(mvEnv)
      mvDetail.Init(mvEnv)
      mvContactAccount = Nothing
      mvNextLineNumber = 1
      mvTransactionSign = ""
      mvNegativesAllowed = ""
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvConfirmedTransChecked = False
    End Sub

    Private Sub SetValid(ByRef pField As FinancialHistoryFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Sub AddDetail(ByRef pProduct As String, ByRef pRate As String, ByRef pQuantity As Integer, ByRef pAmount As Double, ByRef pSource As String, ByRef pVATRate As String, ByRef pVATAmount As Double, ByRef pStatus As FinancialHistoryStatus, ByRef pSalesContactNumber As Integer, ByRef pInvoicePayment As Boolean, ByRef pCurrencyAmount As Double, ByRef pCurrencyVATAmount As Double, ByRef pDistributionCode As String, Optional ByRef pLineNumber As Integer = 0)
      Dim vDetail As New FinancialHistoryDetail(mvEnv)

      With vDetail
        .Init(mvEnv)
        .BatchNumber = mvClassFields.Item(FinancialHistoryFields.fhfBatchNumber).IntegerValue
        .TransactionNumber = mvClassFields.Item(FinancialHistoryFields.fhfTransactionNumber).IntegerValue
        If pLineNumber > 0 Then
          .LineNumber = pLineNumber
        Else
          .LineNumber = mvNextLineNumber
        End If
        .Amount = pAmount
        .ProductCode = pProduct
        .RateCode = pRate
        .Source = pSource
        .Quantity = pQuantity.ToString
        .VatRate = pVATRate
        .VatAmount = pVATAmount
        .Status = pStatus
        .CurrencyAmount = pCurrencyAmount
        .CurrencyVatAmount = pCurrencyVATAmount
        If pSalesContactNumber > 0 Then .SalesContactNumber = pSalesContactNumber
        .InvoicePayment = pInvoicePayment
        .DistributionCode = pDistributionCode
      End With
      mvDetails.Add(vDetail)
      'Set the Detail property to point to the new detail line added
      mvDetail = vDetail
      If mvNextLineNumber <= vDetail.LineNumber Then mvNextLineNumber = vDetail.LineNumber + 1
      'update the financial history details now
      mvClassFields.Item(FinancialHistoryFields.fhfAmount).Value = CStr(Val(mvClassFields.Item(FinancialHistoryFields.fhfAmount).Value) + pAmount)
      mvClassFields.Item(FinancialHistoryFields.fhfCurrencyAmount).Value = CStr(Val(mvClassFields.Item(FinancialHistoryFields.fhfCurrencyAmount).Value) + pCurrencyAmount)

    End Sub

    Public Sub AddDetailFromRecordSet(ByVal pRecordSet As CDBRecordSet)
      Dim vDetail As New FinancialHistoryDetail(mvEnv)

      'Adds a detail line
      vDetail.InitFromRecordSet(mvEnv, pRecordSet, FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll)
      mvDetails.Add(vDetail)
      mvDetail = vDetail
      If mvNextLineNumber <= vDetail.LineNumber Then mvNextLineNumber = vDetail.LineNumber + 1
    End Sub

    Public Sub GetDetail(ByRef pNumber As Integer)
      'Assume the line number already exists
      mvDetail = CType(mvDetails.Item(pNumber), FinancialHistoryDetail)
    End Sub

    Public Function GetRecordSetFields(ByVal pRSType As FinancialHistoryRecordSetTypes) As String
      Dim vFields As String = ""
      'Always include the primary key attributes
      vFields = "fh.batch_number,fh.transaction_number,"
      If (pRSType And FinancialHistoryRecordSetTypes.fhrtNumbers) > 0 Then
        vFields = vFields & "fh.contact_number,fh.address_number,fh.transaction_date,"
        vFields = vFields & "fh.transaction_type,fh.amount,fh.payment_method,"
        vFields = vFields & "fh.reference,fh.posted,fh.status,fh.bank_details_number,"
        If mvClassFields.Item(FinancialHistoryFields.fhfCurrencyAmount).InDatabase Then vFields = vFields & "fh.currency_amount,"
      End If
      If (pRSType And FinancialHistoryRecordSetTypes.fhrtAll) > 0 Then If mvClassFields.Item(FinancialHistoryFields.fhfTransactionOrigin).InDatabase Then vFields = vFields & "fh.transaction_origin,"
      If (pRSType And FinancialHistoryRecordSetTypes.fhrtDetail) > 0 Then vFields = vFields & "fh.notes,"
      If (pRSType And FinancialHistoryRecordSetTypes.fhrtDetailLines) > 0 Then vFields = vFields & mvDetail.GetRecordSetFields(FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll)

      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      GetRecordSetFields = vFields
    End Function
    Public Sub Init(ByVal pEnv As CDBEnvironment)
      Init(pEnv, 0, 0)
    End Sub
    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pBatchNumber > 0 And pTransactionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(FinancialHistoryRecordSetTypes.fhrtAll) & " FROM financial_history fh, financial_history_details fhd WHERE fh.batch_number = " & pBatchNumber & " AND fh.transaction_number = " & pTransactionNumber & " AND fh.batch_number = fhd.batch_number AND fh.transaction_number = fhd.transaction_number ORDER BY fhd.batch_number, fhd.transaction_number, fhd.line_number")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, FinancialHistoryRecordSetTypes.fhrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As FinancialHistoryRecordSetTypes)
      Dim vFields As CDBFields
      Dim vEndOfDetails As Boolean

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(FinancialHistoryFields.fhfBatchNumber, vFields)
        .SetItem(FinancialHistoryFields.fhfTransactionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And FinancialHistoryRecordSetTypes.fhrtNumbers) > 0 Then
          .SetItem(FinancialHistoryFields.fhfContactNumber, vFields)
          .SetItem(FinancialHistoryFields.fhfAddressNumber, vFields)
          .SetItem(FinancialHistoryFields.fhfTransactionDate, vFields)
          .SetItem(FinancialHistoryFields.fhfTransactionType, vFields)
          .SetItem(FinancialHistoryFields.fhfAmount, vFields)
          .SetItem(FinancialHistoryFields.fhfPaymentMethod, vFields)
          .SetItem(FinancialHistoryFields.fhfReference, vFields)
          .SetItem(FinancialHistoryFields.fhfPosted, vFields)
          .SetItem(FinancialHistoryFields.fhfStatus, vFields)
          .SetItem(FinancialHistoryFields.fhfBankDetailsNumber, vFields)
        End If
        If (pRSType And FinancialHistoryRecordSetTypes.fhrtDetail) > 0 Then
          .SetItem(FinancialHistoryFields.fhfNotes, vFields)
        End If
        If pRecordSet.Fields.ContainsKey("transaction_sign") Then mvTransactionSign = pRecordSet.Fields("transaction_sign").Value
      End With
      If (pRSType And FinancialHistoryRecordSetTypes.fhrtDetailLines) > 0 Then
        While pRecordSet.Status() = True And Not vEndOfDetails
          AddDetailFromRecordSet(pRecordSet)
          pRecordSet.Fetch()
          If (pRecordSet.Fields("batch_number").IntegerValue <> BatchNumber) Or (pRecordSet.Fields("transaction_number").IntegerValue <> TransactionNumber) Then vEndOfDetails = True
        End While
      End If
    End Sub

    Public Sub Save()
      Dim vDetail As FinancialHistoryDetail
      Dim vTransaction As Boolean

      SetValid(FinancialHistoryFields.fhfAll)

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      If mvExisting Then
        'WARNING - WHAT? I don't think you should try and update a FH / FHD record!
        System.Diagnostics.Debug.Assert(False, "")
        'mvEnv.Connection.UpdateRecords "financial_history", mvClassFields.UpdateFields, mvClassFields.WhereFields
      Else
        If Not mvContactAccount Is Nothing Then
          If Len(mvContactAccount.AccountNumber) > 0 OrElse mvContactAccount.IbanNumber.Length > 0 Then
            'Contact Account needs inserting
            mvContactAccount.Save()
          End If
        End If
        For Each vDetail In mvDetails
          vDetail.Save()
        Next vDetail

        mvEnv.Connection.InsertRecord("financial_history", mvClassFields.UpdateFields)
      End If
      If vTransaction Then mvEnv.Connection.CommitTransaction()

    End Sub

    Public ReadOnly Property TransactionSign() As String
      Get
        If Len(mvTransactionSign) = 0 Then
          mvTransactionSign = mvEnv.Connection.GetValue("SELECT transaction_sign FROM transaction_types WHERE transaction_type = '" & TransactionType & "'")
        End If
        TransactionSign = mvTransactionSign
      End Get
    End Property

    Public ReadOnly Property NegativesAllowed() As String
      Get
        If Len(mvNegativesAllowed) = 0 Then
          mvNegativesAllowed = mvEnv.Connection.GetValue("SELECT negatives_allowed FROM transaction_types WHERE transaction_type = '" & TransactionType & "'")
        End If
        NegativesAllowed = mvNegativesAllowed
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

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(FinancialHistoryFields.fhfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(FinancialHistoryFields.fhfAddressNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property Amount() As Double
      Get
        Amount = mvClassFields.Item(FinancialHistoryFields.fhfAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(FinancialHistoryFields.fhfAmount).DoubleValue = Value
      End Set
    End Property

    Public Property BankDetailsNumber() As Integer
      Get
        BankDetailsNumber = mvClassFields.Item(FinancialHistoryFields.fhfBankDetailsNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(FinancialHistoryFields.fhfBankDetailsNumber).IntegerValue = Value
      End Set
    End Property

    Public Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(FinancialHistoryFields.fhfBatchNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(FinancialHistoryFields.fhfBatchNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property ContactAccount() As ContactAccount
      Get
        If mvContactAccount Is Nothing Then
          mvContactAccount = New ContactAccount
          mvContactAccount.Init(mvEnv, (mvClassFields.Item(FinancialHistoryFields.fhfBankDetailsNumber).IntegerValue))
        End If
        ContactAccount = mvContactAccount
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(FinancialHistoryFields.fhfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(FinancialHistoryFields.fhfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property Detail() As FinancialHistoryDetail
      Get
        Detail = mvDetail
      End Get
    End Property

    Public ReadOnly Property Details() As Collection
      Get
        Details = mvDetails
      End Get
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(FinancialHistoryFields.fhfNotes).MultiLineValue
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfNotes).Value = Value
      End Set
    End Property

    Public Property PaymentMethod() As String
      Get
        PaymentMethod = mvClassFields.Item(FinancialHistoryFields.fhfPaymentMethod).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfPaymentMethod).Value = Value
      End Set
    End Property

    Public Property Posted() As String
      Get
        Posted = mvClassFields.Item(FinancialHistoryFields.fhfPosted).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfPosted).Value = Value
      End Set
    End Property

    Public Property Reference() As String
      Get
        Reference = mvClassFields.Item(FinancialHistoryFields.fhfReference).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfReference).Value = Value
      End Set
    End Property

    Public Property Status() As FinancialHistoryStatus
      Get
        Select Case mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value
          Case "A"
            Status = FinancialHistoryStatus.fhsAdjusted
          Case "B"
            Status = FinancialHistoryStatus.fhsOnBackOrder
          Case "M"
            Status = FinancialHistoryStatus.fhsMoved
          Case "R"
            Status = FinancialHistoryStatus.fhsReversed
          Case Else
            Status = FinancialHistoryStatus.fhsNormal
        End Select
      End Get
      Set(ByVal Value As FinancialHistoryStatus)
        Select Case Value
          Case FinancialHistoryStatus.fhsAdjusted
            mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value = "A"
          Case FinancialHistoryStatus.fhsOnBackOrder
            mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value = "B"
          Case FinancialHistoryStatus.fhsMoved
            mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value = "M"
          Case FinancialHistoryStatus.fhsReversed
            mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value = "R"
          Case Else
            mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value = ""
        End Select
      End Set
    End Property

    Public Property TransactionDate() As String
      Get
        TransactionDate = mvClassFields.Item(FinancialHistoryFields.fhfTransactionDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionDate).Value = Value
      End Set
    End Property

    Public Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(FinancialHistoryFields.fhfTransactionNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property TransactionOrigin() As String
      Get
        TransactionOrigin = mvClassFields.Item(FinancialHistoryFields.fhfTransactionOrigin).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionOrigin).Value = Value
      End Set
    End Property

    Public Property TransactionType() As String
      Get
        TransactionType = mvClassFields.Item(FinancialHistoryFields.fhfTransactionType).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(FinancialHistoryFields.fhfTransactionType).Value = Value
      End Set
    End Property

    Public ReadOnly Property StatusCode() As String
      Get
        StatusCode = mvClassFields.Item(FinancialHistoryFields.fhfStatus).Value
      End Get
    End Property
    Public ReadOnly Property AdjustmentBatchNumber() As Integer
      Get
        AdjustmentBatchNumber = mvAdjustmentBatchNumber
      End Get
    End Property
    Public ReadOnly Property AdjustmentTransactionNumber() As Integer
      Get
        AdjustmentTransactionNumber = mvAdjustmentTransactionNumber
      End Get
    End Property
    Public ReadOnly Property AdjustmentWasBatchNumber() As Integer
      Get
        AdjustmentWasBatchNumber = mvAdjustmentWasBatchNumber
      End Get
    End Property
    Public ReadOnly Property AdjustmentWasTransactionNumber() As Integer
      Get
        AdjustmentWasTransactionNumber = mvAdjustmentWasTransactionNumber
      End Get
    End Property

    Public Function AdjustTransaction(ByVal pAdjustType As Batch.AdjustmentTypes, ByVal pAdjustmentParameters As CDBParameters, ByVal pAmount As Double, Optional ByVal pLineNo As Integer = 0, Optional ByVal pFAOnly As Boolean = True, Optional ByRef pNewBatchNo As Integer = 0, Optional ByRef pNewTransNo As Integer = 0,
                                      Optional ByRef pNewBatchTransColl As CollectionList(Of BatchTransaction) = Nothing) As String
      'Run Financial Adjustments
      'pFAOnly = False is for EventBooking/Accommodation adjustments

      'Will raise errors to get the calling routine to ask questions and return parameters for the results
      'AllocationsChecked
      'PostToCashBook
      'PartRefundParameters
      'AdjustOriginalProductCost

      Dim vRS As CDBRecordSet
      Dim vNewBatch As New Batch(mvEnv)
      Dim vNewTrans As New BatchTransaction(mvEnv)
      Dim vOldBatch As New Batch(mvEnv)
      Dim vOldTrans As New BatchTransaction(mvEnv)
      Dim vOldAnal As New BatchTransactionAnalysis(mvEnv)
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vBatchType As Batch.BatchTypes
      Dim vAdjust As Boolean
      Dim vMsg As String = ""
      Dim vNewTransNo As Integer
      Dim vStockSale As Boolean
      Dim vSQL As String
      Dim vTransDate As Date
      Dim vTransType As String
      Dim vTransNotes As String
      Dim vBatchDate As String = ""
      Dim vBankAccount As String
      Dim vDeleteTrans As Boolean
      Dim vTransaction As Boolean
      Dim vQuantity As Integer
      Dim vIssued As Integer
      Dim vErrorNumber As Integer
      Dim vAdjustmentType As Batch.AdjustmentTypes
      Dim vPostToCashBook As Boolean
      Dim vSavedBatchNumber As Integer
      Dim vSavedTransNumber As Integer
      Dim vTRDTransaction As TraderTransaction
      Dim vTraderApp As TraderApplication
      Dim vAppName As String
      Dim vPayMethod As String = ""
      Dim vTDRLine As TraderAnalysisLine
      Dim vEBNumber As Integer
      Dim vEB As EventBooking
      Dim vColl As New Collection
      Dim vEPMLines As Boolean
      Dim vLineNumber As Integer
      Dim vPositiveLinesOnly As Boolean = False

      Try
        vAdjustmentType = pAdjustType
        vOldBatch.Init(BatchNumber)
        vOldTrans.Init((vOldBatch.BatchNumber), TransactionNumber)

        Dim vTransactionAmount As Double = vOldTrans.CurrencyAmount
        Dim vPartRefundAnalysis As PartRefundAnalysis = Nothing

        If vAdjustmentType = Batch.AdjustmentTypes.atMove Then
          If PreProcessAdjustmentChecks(pAmount, 0, pAdjustmentParameters.HasValue("AllocationsChecked"), False, vAdjustmentType) Then
            If pNewBatchNo = 0 Then
              'Create new Batch/Transaction/Analysis (for Change Payer)
              vOldTrans.InitBatchTransactionAnalysis((vOldTrans.BatchNumber), (vOldTrans.TransactionNumber))
              vOldTrans.InitAnalysisAdditionalData()
              If vOldBatch.BatchType = Batch.BatchTypes.FinancialAdjustment AndAlso vOldTrans.Amount = 0 Then
                'Performing a ChangePayer on a transaction after a Re-analysis
                Dim vPosAmount As Double = 0
                For Each vBTA As BatchTransactionAnalysis In vOldTrans.Analysis
                  If vBTA.Amount >= 0 Then
                    vPosAmount += vBTA.Amount
                    Select Case vBTA.LineType
                      Case "L", "N", "U"
                        'If we have some Sales Ledger payments then only do this for the positive lines
                        vPositiveLinesOnly = True
                    End Select
                  End If
                Next
                If vPositiveLinesOnly Then vTransactionAmount = vPosAmount
              End If
              vTRDTransaction = New TraderTransaction
              vTRDTransaction.Init(mvEnv)
              vTRDTransaction.TraderAnalysisLines.InitAnalysisFromBTForMove(vOldTrans.Analysis, vOldBatch.BatchType, pAdjustmentParameters("ContactNumber").IntegerValue, pAdjustmentParameters("AddressNumber").IntegerValue, vOldBatch.CurrencyCode, vOldTrans.TransactionDate, vOldTrans.StockMovements, vPositiveLinesOnly)
              vTRDTransaction.TraderAnalysisLines.SetDepositAllowed(mvEnv)
              vAppName = mvEnv.GetConfig("trader_application_fa")
              If vOldBatch.BatchType = Batch.BatchTypes.GiveAsYouEarn Then
                vAppName = mvEnv.GetConfig("trader_application_fapg")
              ElseIf vOldBatch.BatchType = Batch.BatchTypes.PostTaxPayrollGiving Then
                vAppName = mvEnv.GetConfig("trader_application_fa_postpg")
              End If
              vTraderApp = New TraderApplication
              vTraderApp.Init(vAppName, mvEnv, vOldTrans.BatchNumber, vOldTrans.TransactionNumber, Batch.GetBatchTypeCode((vOldBatch.BatchType)), False, Batch.AdjustmentTypes.atMove)
              If vTraderApp.IsValid = False Then RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vAppName)
              If pAdjustmentParameters.Exists("SmartClient") Then
                vTraderApp.GetPayMethod(vOldTrans.PaymentMethod, "", vPayMethod, vAdjustmentType)
              Else
                vTraderApp.GetPayMethod(If(Len(vOldBatch.PaymentMethod) > 0, vOldBatch.PaymentMethod, vOldTrans.PaymentMethod), "", vPayMethod, vAdjustmentType)
              End If
              With pAdjustmentParameters
                .Add("TransactionPaymentMethod", CDBField.FieldTypes.cftCharacter, vPayMethod)
                .Add("PayerContactNumber", .Item("ContactNumber").IntegerValue)
                .Add("PayerAddressNumber", .Item("AddressNumber").IntegerValue)
                .Add("TRD_Receipt", CDBField.FieldTypes.cftCharacter, vOldTrans.Receipt)
                .Add("TRD_EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, BooleanString(vOldTrans.EligibleForGiftAid))
                .Add("Mailing", CDBField.FieldTypes.cftCharacter, vOldTrans.Mailing)
                .Add("COM_Notes", CDBField.FieldTypes.cftCharacter, vOldTrans.Notes)
                .Add("TRD_Amount", CDBField.FieldTypes.cftNumeric, vTransactionAmount.ToString) 'BR17956 pass the amount entered to Trader. Not the converted amount. Trader will convert.
                If .Exists("FATransactionType") = False Then .Add("FATransactionType", CDBField.FieldTypes.cftCharacter, .Item("TransactionType").Value)
                If vPayMethod = "CARD" Then
                  Dim vCardType As String = "C"
                  If vOldTrans.PaymentMethod <> mvEnv.GetConfig("pm_cc") Then vCardType = "D"
                  If .Exists("CDC_CreditOrDebitCard") = False Then .Add("CDC_CreditOrDebitCard", CDBField.FieldTypes.cftCharacter, vCardType)
                End If
              End With
              If vOldBatch.BatchType <> Batch.BatchTypes.CreditSales Then
                'If batch contains invoice payments then update Invoice being paid and CreditCustomers record
                UpdateInvoiceForMove(vOldTrans.BatchNumber, vOldTrans.TransactionNumber, vPositiveLinesOnly, False)
              End If
              vTraderApp.SaveTransaction(vTRDTransaction, pAdjustmentParameters, Batch.AdjustmentTypes.atMove, False)

              'BR12729: Link multiple lines to single event booking
              If vTraderApp.EventMultipleAnalysis = True And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) = True Then
                vEPMLines = False
                For Each vTDRLine In vTRDTransaction.TraderAnalysisLines
                  If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent Then
                    vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, vTDRLine.EventBookingNumber)
                    If (mvEnv.Connection.DeleteRecords("event_booking_transactions", vWhereFields, False)) > 0 Then
                      vWhereFields = New CDBFields
                      If vEBNumber = 0 Then
                        vEBNumber = vTDRLine.EventBookingNumber
                      Else
                        vEB = New EventBooking
                        vEB.Init(mvEnv, 0, vEBNumber)
                        vEB.AddLinkedTransaction(Batch.AdjustmentTypes.atNone, Nothing, Nothing, vTDRLine.LineNumber - 1, vEPMLines)
                        vEBNumber = vTDRLine.EventBookingNumber
                        vEPMLines = False
                      End If
                    End If
                  ElseIf vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEventPricingMatrixLine Then
                    vEPMLines = True
                  End If
                Next vTDRLine
                If vEBNumber > 0 Then
                  vEB = New EventBooking
                  vEB.Init(mvEnv, 0, vEBNumber)
                  vEB.AddLinkedTransaction(Batch.AdjustmentTypes.atNone, Nothing, Nothing, (vTRDTransaction.TraderAnalysisLines(vTRDTransaction.TraderAnalysisLines.Count).LineNumber), vEPMLines)
                End If
              End If

              pNewBatchNo = vTRDTransaction.BatchNumber
              pNewTransNo = vTRDTransaction.TransactionNumber
            End If
            'A new batch and new transaction already been created
            vNewBatch.Init(pNewBatchNo)
            vNewBatch.LockBatch()

            vSavedBatchNumber = pNewBatchNo
            vSavedTransNumber = pNewTransNo

            'Get next transaction number
            vNewTransNo = vNewBatch.AllocateTransactionNumber
            vNewTrans.InitFromBatch(mvEnv, vNewBatch, vNewTransNo)

            'Do the following to make sure that the -ve tras'n is processed before the +ve tras'n
            If Not mvEnv.Connection.InTransaction Then
              mvEnv.Connection.StartTransaction()
              vTransaction = True
            End If
            With vFields
              If vNewTrans.BatchNumber <> vNewBatch.BatchNumber Then .Add("batch_number", CDBField.FieldTypes.cftLong, vNewTrans.BatchNumber)
              .Add("transaction_number", CDBField.FieldTypes.cftLong, vNewTransNo)
            End With
            With vWhereFields
              .Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNo)
              .Add("transaction_number", CDBField.FieldTypes.cftLong, pNewTransNo)
            End With
            mvEnv.Connection.UpdateRecords("batch_transaction_analysis", vFields, vWhereFields)
            mvEnv.Connection.UpdateRecords("batch_transactions", vFields, vWhereFields)
            mvEnv.Connection.UpdateRecords("card_sales", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("credit_sales", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("invoices", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("invoice_details", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("order_payment_history", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("legacy_bequest_receipts", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("gaye_pledge_payment_history", vFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("event_bookings", vFields, vWhereFields, False)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
              mvEnv.Connection.UpdateRecords("event_booking_transactions", vFields, vWhereFields, False)
            End If
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
              'BR13623: Update any new transactions links (created by vTraderApp.SaveTransaction)
              mvEnv.Connection.UpdateRecords("fundraising_payment_history", vFields, vWhereFields, False)
            End If
            mvEnv.Connection.UpdateRecords("collection_payments", vFields, vWhereFields, False)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then
              mvEnv.Connection.UpdateRecords("service_booking_transactions", vFields, vWhereFields, False)
            End If

            'Now update the back_order_details to show reversed status
            With vFields
              .Clear()
              .Add("status", CDBField.FieldTypes.cftCharacter, FA_STATUS_REVERSAL)
            End With
            With vWhereFields
              .Clear()
              .Add("batch_number", CDBField.FieldTypes.cftLong, vOldTrans.BatchNumber)
              .Add("transaction_number", CDBField.FieldTypes.cftLong, vOldTrans.TransactionNumber)
            End With
            mvEnv.Connection.UpdateRecords("back_order_details", vFields, vWhereFields, False)

            If vTransaction Then mvEnv.Connection.CommitTransaction()
            vNewTrans = New BatchTransaction(mvEnv)
            vNewTrans.InitFromBatch(mvEnv, vNewBatch, pNewTransNo)
            With vNewTrans
              .CloneForFA(vOldTrans, False)
              .TransactionDate = pAdjustmentParameters("TransactionDate").Value
              .TransactionType = pAdjustmentParameters("TransactionType").Value
              .Receipt = "N"
              .Notes = pAdjustmentParameters.ParameterExists("Notes").Value
            End With
            vAdjust = True
            pNewBatchNo = vNewTrans.BatchNumber
            pNewTransNo = vNewTransNo
          End If
        Else
          If PreProcessAdjustmentChecks(pAmount, pLineNo, pAdjustmentParameters.HasValue("AllocationsChecked"), pAdjustmentParameters.ParameterExists("FullAmountAllocation").Bool, vAdjustmentType) Then
            'Now do the reversal
            vAdjust = True
            With vOldBatch
              If .Existing = False Then RaiseError(DataAccessErrors.daeOriginalBatchPurged)
              If .PostedToNominal = False Then
                If .ReadyForBanking = True Or .PayingInSlipPrinted = True Or .PostedToCashBook = True Then
                  RaiseError(DataAccessErrors.daeOriginalPaymentPartProcessed)
                Else
                  vDeleteTrans = (pLineNo = 0)
                  If pLineNo > 0 Then
                    'Delete line only unless only line in transaction
                    With vWhereFields
                      .Clear()
                      .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                      .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
                      .Add("line_number", pLineNo, CDBField.FieldWhereOperators.fwoNotEqual)
                    End With
                    If mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) > 0 Then
                      'Other lines in transaction
                      vOldAnal.Init(BatchNumber, TransactionNumber, pLineNo)
                      If vOldAnal.Existing = False Then RaiseError(DataAccessErrors.daeOriginalTransactionPurged)
                      If Not mvEnv.Connection.InTransaction Then
                        mvEnv.Connection.StartTransaction()
                        vTransaction = True
                      End If
                      'a) Update transaction
                      With vOldTrans
                        .Amount = (.Amount - vOldAnal.Amount)
                        .LineTotal = (.LineTotal - vOldAnal.Amount)
                        .CurrencyAmount = (.CurrencyAmount - vOldAnal.CurrencyAmount)
                        .SaveChanges()
                      End With
                      'b) Set new totals on batch
                      With vFields
                        .Clear()
                        .Add("transaction_total", CDBField.FieldTypes.cftNumeric, "transaction_total - " & vOldAnal.Amount)
                        .Add("currency_batch_total", CDBField.FieldTypes.cftNumeric, "currency_batch_total - " & vOldAnal.CurrencyAmount)
                        .Add("currency_transaction_total", CDBField.FieldTypes.cftNumeric, "currency_transaction_total - " & vOldAnal.CurrencyAmount)
                        If vOldBatch.DetailCompleted Then
                          .Add("detail_completed", CDBField.FieldTypes.cftCharacter, "N")
                        End If
                      End With
                      'c) Delete analysis line
                      vWhereFields.Remove("line_number")
                      vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, pLineNo)
                      mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields)
                      'd) Update batch
                      vWhereFields.Remove("transaction_number")
                      vWhereFields.Remove("line_number")
                      mvEnv.Connection.UpdateRecords("batches", vFields, vWhereFields)
                      If vTransaction Then mvEnv.Connection.CommitTransaction()
                    Else
                      'Only line in transaction so transaction can be deleted
                      vDeleteTrans = True
                    End If
                  End If
                  If vDeleteTrans Then
                    vErrorNumber = vOldBatch.Delete(TransactionNumber, pFAOnly) 'delete transaction/analsysis etc. - pFAOnly = False for Event/Accomodation bookings therefore no checks are required.
                    If vErrorNumber <> 0 Then RaiseError(CType(vErrorNumber, DataAccessErrors))
                    If vOldBatch.NumberOfTransactions = 0 Then vErrorNumber = vOldBatch.Delete(0, pFAOnly) 'Delete the batch header if no other transactions in batch
                    If vErrorNumber <> 0 Then RaiseError(CType(vErrorNumber, DataAccessErrors))
                  End If
                  vAdjust = False
                End If
              Else
                vBatchType = vOldBatch.AdjustmentBatchType(pAdjustType)
                vPostToCashBook = True
                If pFAOnly Then
                  'Financial Adjustments
                  'If Not pAdjustmentParameters.Exists("AdjustmentParameters") Then RaiseError daeMissingAdjustmentParameters
                  If pAdjustmentParameters.Exists("TransactionDate") = False Then RaiseError(DataAccessErrors.daeMissingAdjustmentParameters)
                  If (vOldBatch.BatchType <> Batch.BatchTypes.DirectDebit) And (vOldBatch.BatchType <> Batch.BatchTypes.CreditSales) Then
                    If Not pAdjustmentParameters.Exists("PostToCashBook") Then RaiseError(DataAccessErrors.daeMissingAdjustmentParameters)
                    If pAdjustmentParameters("PostToCashBook").Bool = False Then vPostToCashBook = False
                  End If
                  vTransDate = CDate(pAdjustmentParameters("TransactionDate").Value)
                  vTransType = pAdjustmentParameters("TransactionType").Value
                  vTransNotes = pAdjustmentParameters.ParameterExists("Notes").Value
                  vBatchDate = pAdjustmentParameters.ParameterExists("BatchDate").Value
                Else
                  'Only Events related adjustments come in here
                  vTransDate = CDate(TodaysDate())
                  vTransType = String.Empty
                  vTransNotes = String.Empty
                  If vBatchType = Batch.BatchTypes.CreditSales AndAlso pAdjustType = Batch.AdjustmentTypes.atReverse AndAlso TransactionSign.Equals("D", StringComparison.InvariantCultureIgnoreCase) Then
                    'Reversing a credit note so create an invoice
                    vTransType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSTransType)
                  End If
                  If String.IsNullOrWhiteSpace(vTransType) Then vTransType = vOldTrans.AdjustmentTransactionType(vOldBatch.BatchType, TransactionSign, vAdjustmentType)
                  If String.IsNullOrWhiteSpace(vTransType) Then RaiseError(DataAccessErrors.daeCannotFindAdjTransType)
                End If

                If vAdjustmentType = Batch.AdjustmentTypes.atPartRefund Then
                  vOldAnal = New BatchTransactionAnalysis(mvEnv)
                  vOldAnal.Init((vOldTrans.BatchNumber), (vOldTrans.TransactionNumber), pLineNo)
                  If vOldAnal.Existing Then
                    If pAdjustmentParameters.ParameterExists("SalesLedgerPartRefund").Bool = True Then
                      If pAdjustmentParameters.Exists("InvoicePaymentAmount") = False OrElse
                         pAdjustmentParameters.Exists("UnallocatedAmount") = False OrElse
                         pAdjustmentParameters.Exists("RefundAmount") = False Then
                        RaiseError(DataAccessErrors.daeMissingPartRefundParameters, "InvoicePaymentAmount, UnallocatedAmount & RefundAmount")
                      End If
                      vPartRefundAnalysis = New PartRefundAnalysis
                      vPartRefundAnalysis.SalesLedgerPartRefund = True
                      vPartRefundAnalysis.InvoicePaymentAmount = pAdjustmentParameters("InvoicePaymentAmount").DoubleValue
                      vPartRefundAnalysis.UnallocatedAmount = pAdjustmentParameters("UnallocatedAmount").DoubleValue
                      vPartRefundAnalysis.RefundAmount = pAdjustmentParameters("RefundAmount").DoubleValue
                    Else
                      If Not (pAdjustmentParameters.Exists("Quantity") Or pAdjustmentParameters.Exists("Issued")) Then RaiseError(DataAccessErrors.daeMissingPartRefundParameters, "Quantity & Issued")
                      vPartRefundAnalysis = New PartRefundAnalysis
                      vPartRefundAnalysis.SalesLedgerPartRefund = False
                      vPartRefundAnalysis.RefundIssued = pAdjustmentParameters("Issued").IntegerValue
                      vPartRefundAnalysis.RefundQuantity = pAdjustmentParameters("Quantity").IntegerValue
                      vQuantity = pAdjustmentParameters("Quantity").IntegerValue
                      vIssued = pAdjustmentParameters("Issued").IntegerValue
                    End If
                  Else
                    RaiseError(DataAccessErrors.daeCannotLocateOriginalPayment)
                  End If
                End If

                'ab 7oct02 was this: If mvFinHist.Status <> fhsNormal Then RaiseError daeCannotAdjustPayment, FinancialStatusDesc(mvFinHist.StatusCode), FinancialStatusDesc(FA_STATUS_REVERSAL)
                If ((pLineNo > 0 And Status <> FinancialHistoryStatus.fhsNormal And Status <> FinancialHistoryStatus.fhsAdjusted) Or (pLineNo = 0 And Status <> FinancialHistoryStatus.fhsNormal)) And Status <> FinancialHistoryStatus.fhsOnBackOrder Then
                  RaiseError(DataAccessErrors.daeCannotAdjustPaymentStatus, StatusDesc(StatusCode), StatusDesc(FA_STATUS_REVERSAL))
                End If
                vBankAccount = mvEnv.Connection.GetValue("SELECT default_bank_account FROM batch_types WHERE batch_type = '" & Batch.GetBatchTypeCode(vBatchType) & "'")
                If Len(vBankAccount) = 0 Then vBankAccount = vOldBatch.BankAccount

                If pLineNo = 0 AndAlso (pAdjustType = Batch.AdjustmentTypes.atReverse OrElse pAdjustType = Batch.AdjustmentTypes.atRefund) Then
                  'Reversing a transaction containing Sales Ledger items
                  Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber)
                  If vSLItems > SalesLedgerItems.None AndAlso IsFinancialAdjustment() Then
                    If CanReverseAdjustedTransaction(vSLItems) Then
                      vPositiveLinesOnly = True
                    End If
                  End If
                End If

                If pNewBatchNo > 0 And pFAOnly = False Then
                  'Event Bookings with multiple BTA lines to be reversed
                  vNewBatch = New Batch(mvEnv)
                  vNewBatch.Init(pNewBatchNo)
                  vNewBatch.LockBatch()
                  vNewTrans = New BatchTransaction(mvEnv)
                  vNewTrans.Init((vNewBatch.BatchNumber), pNewTransNo)
                Else
                  With vNewBatch
                    .InitOpenBatch(Nothing, Batch.ProvisionalOrConfirmed.Confirmed, vBatchType, vBankAccount, "", vPostToCashBook, (vOldBatch.BatchType), (vOldBatch.CurrencyCode), (vOldBatch.CurrencyExchangeRate), (vOldBatch.BatchCategory), vBatchDate, False, (vOldTrans.Reference), (vOldBatch.BatchAnalysisCode))
                    .TransactionType = vTransType
                    If (vBatchType = Batch.BatchTypes.GiveAsYouEarn Or .BatchType = Batch.BatchTypes.PostTaxPayrollGiving) Then
                      .ReadyForBanking = True
                      .SetPayingInSlipPrinted(0)
                    End If
                    .LockBatch()
                  End With
                  With vNewTrans
                    vNewTransNo = vNewBatch.AllocateTransactionNumber
                    .InitFromBatch(mvEnv, vNewBatch, vNewTransNo)
                    .CloneForFA(vOldTrans, False)
                    .TransactionDate = CStr(vTransDate)
                    .TransactionType = vTransType
                    .Receipt = "N"
                    .Notes = vTransNotes
                  End With
                End If

                'Check the stock levels
                vSQL = "SELECT fhd.product,fhd.rate,line_number FROM financial_history_details fhd,products p WHERE batch_number = " & BatchNumber & " AND transaction_number = " & TransactionNumber
                If pLineNo > 0 Then vSQL = vSQL & " AND line_number = " & pLineNo
                vSQL = vSQL & " AND fhd.product IS NOT NULL AND p.product = fhd.product AND "
                If pLineNo > 0 Then
                  'Adjusting just the line so stock products only
                  vSQL = vSQL & "p.stock_item = 'Y'"
                Else
                  'Adjusting the transaction so include P&P products
                  vSQL = vSQL & "(p.stock_item = 'Y' OR p.postage_packing = 'Y')"
                End If
                vRS = mvEnv.Connection.GetRecordSet(vSQL)
                While vRS.Fetch() = True
                  vOldAnal = New BatchTransactionAnalysis(mvEnv)
                  vOldAnal.Init(BatchNumber, TransactionNumber, (vRS.Fields("line_number").IntegerValue))
                  If vOldAnal.LineType = "P" Or vOldAnal.LineType = "G" Then
                    vOldTrans.InitBatchTransactionAnalysis(BatchNumber, TransactionNumber)
                    If vAdjustmentType <> Batch.AdjustmentTypes.atPartRefund Then vQuantity = vOldTrans.Analysis.Item(CStr(vRS.Fields("line_number").IntegerValue) & vOldAnal.LineType & vRS.Fields("product").Value).Quantity
                    CheckStockAdjustment(pAdjustmentParameters, vRS.Fields("line_number").IntegerValue, vRS.Fields("product").Value, vRS.Fields("rate").Value, vQuantity, vStockSale, CType(IIf(vAdjustmentType = Batch.AdjustmentTypes.atPartRefund, Batch.AdjustmentTypes.atPartRefund, Batch.AdjustmentTypes.atReverse), Batch.AdjustmentTypes), vNewTrans, vIssued, vOldBatch.BatchType)
                  End If
                End While
                vRS.CloseRecordSet()
              End If
            End With
          End If
        End If

        If vAdjust Then
          If Reverse(vNewBatch, vNewTrans, vAdjustmentType, pLineNo, vStockSale, False, False, True, vPositiveLinesOnly, vPartRefundAnalysis) Then
            If vAdjustmentType = Batch.AdjustmentTypes.atMove Then
              'Need to deal with the ops records here for a Move
              MoveScheduledPayments(BatchNumber, TransactionNumber, vNewTrans.BatchNumber, vNewTransNo)
              'vMsg = (ProjectText.String29516)    'Transaction has been Moved
              vMsg = (ProjectText.String19060) 'Payer has been changed
            ElseIf vAdjustmentType = Batch.AdjustmentTypes.atPartRefund Then
              vMsg = (ProjectText.String19061) 'Analysis Line has been part refunded
              pNewBatchNo = vNewTrans.BatchNumber
              pNewTransNo = vNewTrans.TransactionNumber
            ElseIf vAdjustmentType = Batch.AdjustmentTypes.atRefund Then
              vMsg = (ProjectText.String29515) 'Transaction has been refunded
              pNewBatchNo = vNewTrans.BatchNumber
              pNewTransNo = vNewTrans.TransactionNumber
            Else
              If pLineNo > 0 Then
                vMsg = (ProjectText.String29513) 'Analysis Line as been reversed
              Else
                vMsg = (ProjectText.String29514) 'Transaction has been reversed
              End If
              pNewBatchNo = vNewTrans.BatchNumber
              pNewTransNo = vNewTrans.TransactionNumber
            End If

            If Not mvEnv.Connection.InTransaction Then  'InTransaction = True when cancelling an event booking. (CancelEventBooking and UpdateEventBooking web services)
              'Create Card Sales
              AdjustTransactionPostProcess(mvEnv, BatchNumber, TransactionNumber, vNewBatch.BatchTypeCode, vNewTrans, vMsg)
            Else
              If vNewBatch.BatchType = Batch.BatchTypes.CreditCard OrElse vNewBatch.BatchType = Batch.BatchTypes.DebitCard OrElse vNewBatch.BatchType = Batch.BatchTypes.CreditCardWithInvoice Then 'Refund
                Dim vAllowCardSales As Boolean = True
                If pAdjustmentParameters.ContainsKey("RunType") Then
                  If pAdjustmentParameters("RunType").Value = "V" Then
                    'We are cancelling a credit/debit card event booking with the reverse option 
                    ' - i.e. no money taken before cancellation, so a card sale record is not required. 
                    vAllowCardSales = False
                  End If
                End If
                If vAllowCardSales Then

                  'Set these to create the card sale record after the database transaction
                  If pNewBatchTransColl Is Nothing Then pNewBatchTransColl = New CollectionList(Of BatchTransaction)
                  Dim vNewBatchTransKey As String = BatchNumber & "|" & TransactionNumber & "|" & vNewBatch.BatchTypeCode
                  If Not pNewBatchTransColl.ContainsKey(vNewBatchTransKey) Then pNewBatchTransColl.Add(vNewBatchTransKey, vNewTrans)
                End If
              End If
            End If

            If pFAOnly Then
              If vNewBatch.Existing Then vNewBatch.UnLockBatch()
            End If

            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, vOldTrans.BatchNumber)
              vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, vOldTrans.TransactionNumber)
              vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong)
              If pLineNo > 0 Then
                vWhereFields("line_number").Value = CStr(pLineNo)
                vLineNumber = 1
                If vNewTrans.NextLineNumber > 2 Then vLineNumber = vNewTrans.NextLineNumber - 1
                vSQL = "SELECT event_number, booking_number FROM event_booking_transactions WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
                vRS = mvEnv.Connection.GetRecordSet(vSQL)
                If vRS.Fetch() = True Then
                  vColl.Add(vRS.Fields(1).Value & "," & vRS.Fields(2).Value & "," & vNewTrans.BatchNumber & "," & vNewTrans.TransactionNumber & "," & vLineNumber)
                End If
                vRS.CloseRecordSet()
              Else
                vOldTrans.InitBatchTransactionAnalysis((vOldTrans.BatchNumber), (vOldTrans.TransactionNumber))
                For Each vOldAnal In vOldTrans.Analysis
                  vWhereFields("line_number").Value = CStr(vOldAnal.LineNumber)
                  vSQL = "SELECT event_number, booking_number FROM event_booking_transactions WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
                  vRS = mvEnv.Connection.GetRecordSet(vSQL)
                  If vRS.Fetch() = True Then
                    vColl.Add(vRS.Fields(1).Value & "," & vRS.Fields(2).Value & "," & vNewTrans.BatchNumber & "," & vNewTrans.TransactionNumber & "," & vOldAnal.LineNumber)
                  End If
                  vRS.CloseRecordSet()
                Next vOldAnal
              End If
              vEB = New EventBooking
              vEB.Init(mvEnv)
              vEB.AddLinkedTransaction(Batch.AdjustmentTypes.atNone, Nothing, vColl)
            End If

            'BR13623: Add reversed analysis lines to the linked table - Don't check the flag on TraderApp
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", vOldTrans.BatchNumber)
              vWhereFields.Add("transaction_number", vOldTrans.TransactionNumber)
              If mvEnv.Connection.GetCount("fundraising_payment_history", vWhereFields) > 0 Then
                vOldTrans.InitBatchTransactionAnalysis(vOldTrans.BatchNumber, vOldTrans.TransactionNumber)
                vOldTrans.SetAdditionalData(BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment)
                For Each vBTA As BatchTransactionAnalysis In vOldTrans.Analysis
                  Dim vFPH As New FundraisingPaymentHistory(mvEnv)
                  If pLineNo > 0 Then
                    If pLineNo = vBTA.LineNumber Then
                      vLineNumber = 1
                      If vNewTrans.NextLineNumber > 2 Then vLineNumber = vNewTrans.NextLineNumber - 1
                      vFPH.Init()
                      vFPH.CreateNewLink(vBTA.AdditionalNumber, vNewTrans.BatchNumber, vNewTrans.TransactionNumber, vLineNumber)
                      Exit For
                    End If
                  Else
                    vFPH.Init()
                    vFPH.CreateNewLink(vBTA.AdditionalNumber, vNewTrans.BatchNumber, vNewTrans.TransactionNumber, vBTA.LineNumber)
                  End If
                Next
              End If
            End If
          End If
          If vNewBatch.Existing And vNewBatch.Locked Then vNewBatch.UnLockBatch()
        End If
        AdjustTransaction = vMsg
      Catch vEx As Exception
        If mvEnv.Connection.InTransaction Then mvEnv.Connection.RollbackTransaction()
        If vAdjustmentType = Batch.AdjustmentTypes.atMove Then
          If vSavedBatchNumber > 0 Then DeleteFromBatch(vSavedBatchNumber, vSavedTransNumber, pNewBatchNo, pNewTransNo, True, vPositiveLinesOnly)
        End If
        PreserveStackTrace(vEx)
        Throw vEx
      End Try
    End Function

    Public Shared Sub AdjustTransactionPostProcess(ByVal pEnv As CDBEnvironment, ByVal pOldBatchNo As Integer, ByVal pOldTransNo As Integer, ByVal pNewBatchType As String, ByVal pNewTrans As BatchTransaction, ByRef pMsg As String)
      'This needs to be called outside of database transaction and only to be called if not already called within AdjustTransaction
      Dim vCardSale As New CardSale(pEnv)
      If Batch.GetBatchType(pNewBatchType) = Batch.BatchTypes.CreditCard OrElse Batch.GetBatchType(pNewBatchType) = Batch.BatchTypes.DebitCard OrElse Batch.GetBatchType(pNewBatchType) = Batch.BatchTypes.CreditCardWithInvoice Then 'Refund
        vCardSale.Init(pOldBatchNo, pOldTransNo)
        vCardSale.Clone((pNewTrans.BatchNumber), (pNewTrans.TransactionNumber), False) 'Claim is required
        Dim vCCA As New CreditCardAuthorisation
        vCCA.InitFromTransaction(pEnv, pOldBatchNo, pOldTransNo)
        If vCCA.AuthorisedTransactionNo.Length > 0 Then
          Dim vNewCCA As New CreditCardAuthorisation
          vNewCCA.Init(pEnv)
          vNewCCA.ContactNumber = pNewTrans.ContactNumber
          If vNewCCA.AuthoriseTransaction(vCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatRefund, pNewTrans.Amount, pNewTrans.AddressNumber, vCCA.AuthorisedTransactionNo, vCCA.AuthorisedTextId, vCCA.AuthorisationCode, vCCA.AuthorisationNumber) Then
            pMsg &= vbCrLf & "Online Card Authorisation: " & vNewCCA.AuthorisationResponseMessage
          Else
            pMsg &= vbCrLf & "Online Card Authorisation: Failed- " & vNewCCA.AuthorisationResponseMessage
          End If
        End If
        vCardSale.Save()
      End If
    End Sub

    Private Sub MoveScheduledPayments(ByVal pOldBatchNo As Integer, ByVal pOldTransNo As Integer, ByVal pNewBatchNo As Integer, ByVal pNewTransNo As Integer)
      Dim vRS As CDBRecordSet
      Dim vOPH As OrderPaymentHistory
      Dim vOPS As New OrderPaymentSchedule
      Dim vOPS1 As New OrderPaymentSchedule
      Dim vPP As New PaymentPlan
      Dim vNegOPS As New Collection
      Dim vPosOPS As New Collection
      Dim vPPColl As New CDBCollection
      Dim vSQL As String
      Dim vRemove As Boolean

      vOPS.Init(mvEnv)
      vOPS1.Init(mvEnv)
      vPP.Init(mvEnv)

      '(1) Select all the ops for the negative transaction
      vSQL = "SELECT " & vOPS.GetRecordSetFields(OrderPaymentSchedule.OrderPaymentScheduleRecordSetTypes.opsrtAll) & ",oph.amount FROM reversals r, order_payment_history oph, order_payment_schedule ops"
      vSQL = vSQL & " WHERE r.was_batch_number = " & pOldBatchNo & " AND r.was_transaction_number = " & pOldTransNo
      vSQL = vSQL & " AND oph.batch_number = r.batch_number AND oph.transaction_number = r.transaction_number AND oph.line_number = r.line_number"
      vSQL = vSQL & " AND ops.order_number = oph.order_number AND ops.scheduled_payment_number = oph.scheduled_payment_number"
      vSQL = vSQL & " ORDER BY ops.scheduled_payment_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vOPS = New OrderPaymentSchedule
        vOPS.InitFromRecordSet(mvEnv, vRS, OrderPaymentSchedule.OrderPaymentScheduleRecordSetTypes.opsrtAll)
        vOPS.PaymentAmount = vRS.Fields("amount").DoubleValue
        vNegOPS.Add(vOPS)
      End While
      vRS.CloseRecordSet()

      '(2) Select all the ops for the positive transaction
      vSQL = "SELECT " & vOPS.GetRecordSetFields(OrderPaymentSchedule.OrderPaymentScheduleRecordSetTypes.opsrtAll) & ",oph.amount FROM order_payment_history oph, order_payment_schedule ops"
      vSQL = vSQL & " WHERE batch_number = " & pNewBatchNo & " AND transaction_number = " & pNewTransNo
      vSQL = vSQL & " AND ops.order_number = oph.order_number AND ops.scheduled_payment_number = oph.scheduled_payment_number"
      vSQL = vSQL & " ORDER BY ops.scheduled_payment_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vOPS = New OrderPaymentSchedule
        vOPS.InitFromRecordSet(mvEnv, vRS, OrderPaymentSchedule.OrderPaymentScheduleRecordSetTypes.opsrtAll)
        vOPS.PaymentAmount = vRS.Fields("amount").DoubleValue
        vPosOPS.Add(vOPS)
      End While
      vRS.CloseRecordSet()

      '(3) Update ops records which are in both transactions
      For Each vOPS In vNegOPS
        For Each vOPS1 In vPosOPS
          If vOPS1.ScheduledPaymentNumber = vOPS.ScheduledPaymentNumber Then
            'Same record in both
            vOPS.SetUnProcessedPayment(True, vOPS.PaymentAmount)
            vOPS.SetUnProcessedPayment(True, vOPS1.PaymentAmount)
            vOPS.Save()
            vOPS.PaymentAmount = 0
            vOPS1.PaymentAmount = 0
          End If
        Next vOPS1
      Next vOPS

      '(4) Update ops records that are in negative transaction only - need to Reverse the payment
      For Each vOPS In vNegOPS
        If vOPS.PaymentAmount <> 0 Then
          If vPP.PlanNumber <> vOPS.PlanNumber Then vPP.Init(mvEnv, (vOPS.PlanNumber))
          vOPS.Reverse(vPP, vOPS.PaymentAmount)
          vOPS.Save()
          vOPS.PaymentAmount = 0
        End If
      Next vOPS

      '(5) Check for any negative transactions not allocated to an OPS (moving a pre-v5.x payment)
      '    If both negative & positive sides are the same then leave as they are,
      '    otherwise if just the negative side has no OPS then treat the same as a reversal
      vSQL = "batch_number = " & pOldBatchNo & " AND transaction_number = " & pOldTransNo & " AND scheduled_payment_number IS NULL"
      If mvEnv.Connection.GetCount("order_payment_history", Nothing, vSQL) > 0 Then
        'Only do this if we have original OPH without OPS Number
        vNegOPS = New Collection
        vPPColl = New CDBCollection
        vOPH = New OrderPaymentHistory
        vPP = New PaymentPlan
        vOPH.Init(mvEnv)
        vPP.Init(mvEnv)
        '(a) Get the OPH/PP for the negative transaction
        vSQL = Replace(vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll), "oph.order_number,", "")
        vSQL = "SELECT " & Replace(vPP.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstNumbers Or PaymentPlan.PayPlanRecordSetTypes.pprstPayment Or PaymentPlan.PayPlanRecordSetTypes.pprstType Or PaymentPlan.PayPlanRecordSetTypes.pprstCancel), "balance,", "") & ", " & vSQL
        vSQL = vSQL & ", r.batch_number AS r_batch_number,"
        vSQL = vSQL & " r.transaction_number AS r_transaction_number, r.line_number AS r_line_number"
        vSQL = vSQL & " FROM reversals r, order_payment_history oph, orders o"
        vSQL = vSQL & " WHERE r.was_batch_number = " & pOldBatchNo & " AND r.was_transaction_number = " & pOldTransNo
        vSQL = vSQL & " AND oph.batch_number = r.batch_number AND oph.transaction_number = r.transaction_number"
        vSQL = vSQL & " AND oph.line_number = r.line_number AND oph.scheduled_payment_number IS NULL"
        vSQL = vSQL & " AND o.order_number = oph.order_number ORDER BY r.batch_number, r.transaction_number, r.line_number"
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch() = True
          vOPH = New OrderPaymentHistory
          vOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          vNegOPS.Add(vOPH, CStr(vOPH.PaymentNumber))
          If vPP.PlanNumber <> vOPH.OrderNumber Then
            vPP = New PaymentPlan
            vPP.InitFromRecordSet(mvEnv, vRS, PaymentPlan.PayPlanRecordSetTypes.pprstNumbers Or PaymentPlan.PayPlanRecordSetTypes.pprstPayment Or PaymentPlan.PayPlanRecordSetTypes.pprstType Or PaymentPlan.PayPlanRecordSetTypes.pprstCancel)
            If vPPColl.Exists(vPP.PlanNumber.ToString) = False Then vPPColl.Add(vPP, vPP.PlanNumber.ToString)
          End If
        End While
        vRS.CloseRecordSet()

        If vNegOPS.Count() > 0 Then
          '(b) We found some negative OPH without OPS, now get any positive OPH without OPS
          '    Anything found that matches negative OPH does not need to be updated
          vSQL = "SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph"
          vSQL = vSQL & " WHERE batch_number = " & pNewBatchNo & " AND transaction_number = " & pNewTransNo
          vSQL = vSQL & " ORDER BY line_number"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          While vRS.Fetch() = True
            For Each vOPH In vNegOPS
              vRemove = False
              If (vOPH.LineNumber = vRS.Fields("line_number").IntegerValue) And (vOPH.OrderNumber = vRS.Fields("order_number").IntegerValue) Then
                'Same LineNumber & PPNumber
                If System.Math.Abs(vOPH.Amount) = System.Math.Abs(vRS.Fields("amount").DoubleValue) Then
                  'Same Amount so assume same payment
                  'Remove OPH from collection as no change required
                  vRemove = True
                End If
              ElseIf (vOPH.LineNumber > vRS.Fields("line_number").IntegerValue) And (vOPH.OrderNumber = vRS.Fields("order_number").IntegerValue) Then
                'Same PPNumber but LineNumber is lower
                If System.Math.Abs(vOPH.Amount) = System.Math.Abs(vRS.Fields("amount").DoubleValue) Then
                  'Amount is the same so assume same payment
                  'Remove OPH from collection as no change required
                  vRemove = True
                End If
              End If
              If vRemove Then Exit For
            Next vOPH
            If vRemove Then vNegOPS.Remove(CStr(vOPH.PaymentNumber))
          End While
          vRS.CloseRecordSet()

          'Anything left in vNegOPS collection are negative OPH that do not have positive OPH
          vOPS = New OrderPaymentSchedule
          vOPS.Init(mvEnv)
          For Each vOPH In vNegOPS
            'Update each OPH to have an OPS Number
            vPP = CType(vPPColl(vOPH.OrderNumber.ToString), PaymentPlan)
            With vOPH
              .SetValues(.BatchNumber, .TransactionNumber, .PaymentNumber, .OrderNumber, .Amount, .LineNumber, .Balance, vOPS.ReverseHistoricPayment(vPP, vOPH.Amount))
              .Save()
            End With
          Next vOPH
        End If
      End If
    End Sub

    Private Sub CheckAllocations(ByVal pBatch As Integer, ByVal pTransaction As Integer, ByVal pTransAmount As Double, ByVal pFullAmountAllocation As Boolean, ByVal pEventAdjustment As Boolean)
      Dim vInvoice As New Invoice()
      vInvoice.Init(mvEnv, pBatch, pTransaction)
      If vInvoice.Existing Then
        Dim vText As String = "transaction"
        Dim vType As Invoice.InvoiceRecordType = Invoice.GetRecordType(vInvoice.RecordType)
        Select Case vType
          Case Invoice.InvoiceRecordType.Invoice
            Dim vAllocations As Double
            Dim vContainsUnpostedTrans As Boolean
            vAllocations = vInvoice.AllocationsAmount(pFullAmountAllocation, pFullAmountAllocation, vContainsUnpostedTrans)
            If vContainsUnpostedTrans Then RaiseError(DataAccessErrors.daeInvoiceAllocationsUnpostedFA)
            If pFullAmountAllocation Then
              'Currently used by Event Booking and Exam Booking Cancellation
              'BR17149: Where the new 'cancel_cn_leave_unallocated' config is unset or set to 'A' (allocate) the default behaviour to prompt where allocations exist and automatically allocate the credit note to the invoice will apply.
              'Where the config is set to 'U' (unallocate) the credit note will not be automatically allocated to the invoice and the invoice allocations will remain with the user will being prompted to confirm this
              'Where the config is set to 'S' (ask) the user will be prompted if they want to allocate the credit note or leave it unallocated
              Select Case mvEnv.GetConfig("cancel_cn_leave_unallocated", "A")
                Case "U"  'Unallocate credit note
                  If vAllocations > 0 Then
                    RaiseError(DataAccessErrors.daeUnallocateCreditNoteWithAllocations, FixedFormat(vAllocations), FixedFormat(pTransAmount))
                    'This transaction corresponds to an invoice that has allocations of %1.\r\nThe resulting credit note for %2 will not be allocated against the invoice 
                    'and the invoice allocations will remain.\r\n\r\nDo you wish to continue with this adjustment?
                  Else
                    RaiseError(DataAccessErrors.daeUnallocateCreditNoteWithoutAllocations, FixedFormat(pTransAmount))
                    'Processing this transaction will result in a credit note for %1. The credit note will not be allocated against the invoice linked to the transaction.\r\nDo you wish to continue with this adjustment?
                  End If
                Case "S"  'Ask to allocate or unallocate
                  If vAllocations > 0 Then
                    RaiseError(DataAccessErrors.daeAllocateOrUnallocateCNWithAllocations, FixedFormat(vAllocations), FixedFormat(pTransAmount))
                    'This transaction corresponds to an invoice that has allocations of %1.\r\nThe resulting credit note for %2 can either be allocated to the invoice or left unallocated.\r\nSelect 'Yes' to allocate the credit note to the invoice. The allocations will become unallocated sales ledger cash.\r\nSelect 'No' to leave the credit note unallocated. The allocations will remain on the invoice.\r\nSelect 'Cancel' to cancel this adjustment.
                  Else
                    RaiseError(DataAccessErrors.daeAllocateOrUnallocateCNWithoutAllocations, FixedFormat(pTransAmount))
                    'Processing this transaction will result in a credit note for %1 which can either be allocated to the invoice or left unallocated.\r\nSelect 'Yes' to allocate the credit note to the invoice, 'No' to leave the credit note unallocated or 'Cancel' to cancel this adjustment.
                  End If
                Case Else
                  ''A' - Allocate credit note - default behaviour
                  If vAllocations > 0 Then RaiseError(DataAccessErrors.daeFullInvoiceAllocations, FixedFormat(vAllocations), FixedFormat(pTransAmount)) 'This transaction corresponds to an invoice that has allocations of %1.\r\nThe resulting credit note for %2 will be allocated against the invoice and the allocations will become unallocated sales ledger cash.\r\n\r\nDo you wish to continue with this adjustment?
              End Select
            ElseIf FixTwoPlaces(vAllocations + pTransAmount) > FixTwoPlaces(pTransAmount) Then
              'Dim vUnpaid As Double = pTransAmount - vAllocations
              'Dim vRemaining As Double = pTransAmount - vUnpaid
              'RaiseError(DataAccessErrors.daeinvoiceAllocations, vText, FixedFormat(vAllocations), FixedFormat(vUnpaid), FixedFormat(vRemaining))
              'BR16941: For allocated invoice transaction refunds replaced question option to create unallocated credit note 
              'to instead inform user to manually to remove invoice allocations
              'RaiseError(DataAccessErrors.daeInvoiceAllocationsRemovalRequired, vText, FixedFormat(vAllocations))
              If pEventAdjustment Then
                Dim vUnpaid As Double = pTransAmount - vAllocations
                Dim vRemaining As Double = pTransAmount - vUnpaid
                RaiseError(DataAccessErrors.daeinvoiceAllocations, vText, FixedFormat(vAllocations), FixedFormat(vUnpaid), FixedFormat(vRemaining))
              Else
                RaiseError(DataAccessErrors.daeInvoiceAllocationsRemovalRequired, vText, FixedFormat(vAllocations))
              End If
            End If
          Case Else 'cash or credit note
            Dim vWhereFields As New CDBFields(New CDBField("batch_number", pBatch))
            vWhereFields.Add("transaction_number", pTransaction)
            If mvEnv.Connection.GetCount("invoice_payment_history", vWhereFields) > 0 Then
              If vType = Invoice.InvoiceRecordType.CreditNote Then
                RaiseError(DataAccessErrors.daeAllocated)
              Else 'Invoice.InvoiceRecordType.SalesLedgerCash
                'RaiseError(DataAccessErrors.daeAllocatedSLCash, vText)
              End If
            Else
              If vType = Invoice.InvoiceRecordType.CreditNote Then
                RaiseError(DataAccessErrors.daeUnallocated)
              Else  'Invoice.InvoiceRecordType.SalesLedgerCash
                'RaiseError(DataAccessErrors.daeUnallocatedSLCash, vText)
              End If
            End If
        End Select
      End If
    End Sub

    Private Sub CheckStockAdjustment(ByRef pAdjustmentParameters As CDBParameters, ByVal pLineNumber As Integer, ByVal pProductCode As String, ByVal pRate As String, ByVal pOrdered As Integer, ByRef pStockSale As Boolean, ByVal pAdjustment As Batch.AdjustmentTypes, ByVal pFATrans As BatchTransaction, ByVal pIssued As Integer, ByVal pOrigBatchType As Batch.BatchTypes)
      Dim vBOD As New BackOrderDetail
      Dim vRecordSet As CDBRecordSet
      Dim vIssued As Integer
      Dim vUpdateStock As MsgBoxResult
      Dim vStockMovement As StockMovement
      Dim vSQL As String
      Dim vWarehouse As String
      Dim vBackOrder As Boolean
      Dim vProcessed As Boolean
      Dim vIndex As Integer
      Dim vUseProdCosts As Boolean
      Dim vAdjustOrigPC As Boolean
      Dim vProductCosts As ProductCosts
      Dim vProdCostNo As Integer
      Dim vQuantity As Integer

      vUseProdCosts = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts)
      vBackOrder = (pOrigBatchType = Batch.BatchTypes.BackOrder)

      'Check for stock movements
      'For BackOrders there are no IssuedSock records for our Batch/Transaction/Line so need to link via DespatchNotes
      'But DespatchNotes could link to multiple IssuedStock
      'May be neccessary to perform two selections to find the data - '1' Non-back-orders, '2' Back-orders
      For vIndex = 1 To 2
        If vIndex = 2 Then vBackOrder = True 'Always want 2nd selection to be BackOrders
        vSQL = "SELECT p.product, st.warehouse AS st_warehouse, st.issued"
        If vUseProdCosts Then
          vSQL = vSQL & ", sm.product_cost_number, SUM(sm.movement_quantity) AS sum_movement_quantity"
          If vBackOrder Then vSQL = vSQL & ", stock_movement_reason"
        End If
        vSQL = vSQL & " FROM"
        If vBackOrder Then
          vSQL = vSQL & " batch_transaction_analysis bta"
          vSQL = vSQL & " INNER JOIN despatch_notes dn ON dn.batch_number = bta.batch_number AND dn.transaction_number = bta.transaction_number"
          vSQL = vSQL & " INNER JOIN issued_stock st ON st.picking_list_number = dn.picking_list_number AND st.despatch_note_number = dn.despatch_note_number AND st.allocated = bta.quantity"
        Else
          vSQL = vSQL & " issued_stock st"
        End If
        vSQL = vSQL & " INNER JOIN products p ON p.product = st.product AND p.stock_item = 'Y'"
        If vUseProdCosts Then vSQL = vSQL & " LEFT OUTER JOIN stock_movements sm ON sm.batch_number = st.batch_number AND sm.transaction_number = st.transaction_number AND sm.line_number = st.line_number AND sm.warehouse = st.warehouse"
        vSQL = vSQL & " WHERE"
        If vBackOrder Then
          vSQL = vSQL & " bta.batch_number = " & BatchNumber & " AND bta.transaction_number = " & TransactionNumber & " AND bta.line_number = " & pLineNumber
        Else
          vSQL = vSQL & " st.batch_number = " & BatchNumber & " AND st.transaction_number = " & TransactionNumber & " AND st.line_number = " & pLineNumber
        End If
        If vUseProdCosts Then
          vSQL = vSQL & " GROUP BY p.product, st.warehouse, st.issued, sm.product_cost_number"
          If vBackOrder Then vSQL = vSQL & ", stock_movement_reason"
        End If
        vRecordSet = mvEnv.Connection.GetRecordSetAnsiJoins(vSQL)

        With vRecordSet
          If .Fetch() = True Then
            pStockSale = True
            vProcessed = True
            If pAdjustment = Batch.AdjustmentTypes.atPartRefund Then
              vIssued = pIssued 'May not be the complete stock
            Else
              vIssued = .Fields.Item("issued").IntegerValue
            End If
            vWarehouse = .Fields("st_warehouse").Value
            If pAdjustment = Batch.AdjustmentTypes.atAdjustment Then 'Or pAdjustment = atPartRefund Then
              vUpdateStock = MsgBoxResult.Yes
            Else
              If pAdjustmentParameters.Exists("AdjustStockLevels") = False Then RaiseError(DataAccessErrors.daeMissingAdjustmentParameters)
              If pAdjustmentParameters("AdjustStockLevels").Bool Then vUpdateStock = MsgBoxResult.Yes
            End If
            If vUpdateStock = MsgBoxResult.Yes Then
              If Not mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlStockInterface) And mvEnv.Connection.GetCount("stock_movement_controls", Nothing) > 0 Then
                'Could have selected multiple records due to linking to StockMovements
                vAdjustOrigPC = False
                If vUseProdCosts Then
                  If pAdjustmentParameters.Exists("AdjustOriginalProductCost") = False Then RaiseError(DataAccessErrors.daeMissingAdjustmentParameters)
                  vAdjustOrigPC = pAdjustmentParameters("AdjustOriginalProductCost").Bool
                End If
                Do
                  vProdCostNo = 0
                  If vUseProdCosts Then
                    If vAdjustOrigPC Then vProdCostNo = vRecordSet.Fields("product_cost_number").IntegerValue
                    If vProdCostNo = 0 Then
                      'If original StockMovement did not have ProductCostNumber then allocate stock to latest record
                      vProductCosts = New ProductCosts
                      vProductCosts.InitFromProductAndWarehouse(mvEnv, vRecordSet.Fields("product").Value, vWarehouse, False)
                      vProdCostNo = vProductCosts.GetLatestProductCost.ProductCostNumber
                    End If
                  End If
                  vQuantity = vIssued
                  If vUseProdCosts Then
                    If (vQuantity > System.Math.Abs(vRecordSet.Fields("sum_movement_quantity").IntegerValue)) And vAdjustOrigPC = True Then
                      vQuantity = (vRecordSet.Fields("sum_movement_quantity").IntegerValue * -1)
                    End If
                    If vBackOrder Then
                      'Back Orders will select the original Stock Movement as well so only look at the back order Stock Movement
                      If vRecordSet.Fields("stock_movement_reason").Value <> mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonBackOrder) Then
                        vQuantity = 0
                      End If
                    End If
                  End If
                  If vQuantity <> 0 Then
                    vStockMovement = New StockMovement
                    vStockMovement.Create(mvEnv, (vRecordSet.Fields("product").Value), vQuantity, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonReversal), (pFATrans.BatchNumber), (pFATrans.TransactionNumber), (pFATrans.NextLineNumber), False, vWarehouse, vProdCostNo)
                  End If
                Loop While .Fetch() = True And vAdjustOrigPC = True
              End If
            End If
          End If
          If vBackOrder Then vProcessed = True
          .CloseRecordSet()
        End With
        If vProcessed Then Exit For
      Next
      'Check for back order details
      vBOD.Init(mvEnv, BatchNumber, TransactionNumber, pLineNumber)
      If vBOD.Existing Then
        vBOD.SetAdjustment(FA_STATUS_REVERSAL, (pOrdered * -1), (vIssued * -1))
        vBOD.Save()
      End If
    End Sub

    Public Function Reverse(ByVal pNewBatch As Batch, ByVal pNewTrans As BatchTransaction, ByVal pAdjustType As Batch.AdjustmentTypes, Optional ByRef pLineNo As Integer = 0, Optional ByVal pStockSale As Boolean = False, Optional ByVal pRefundQuantity As Integer = 0, Optional ByVal pRefundIssued As Integer = 0, Optional ByVal pReverseCurrentOPSOnly As Boolean = False, Optional ByVal pRemoveInvoiceAllocations As Boolean = False, Optional ByVal pUpdateCashInvoiceAmountPaid As Boolean = True) As Boolean
      Dim PartRefundAnalysis As PartRefundAnalysis = Nothing
      If pRefundIssued <> 0 OrElse pRefundQuantity <> 0 Then
        PartRefundAnalysis = New PartRefundAnalysis
        PartRefundAnalysis.RefundIssued = pRefundIssued
        PartRefundAnalysis.RefundQuantity = pRefundQuantity
      End If
      Return Reverse(pNewBatch, pNewTrans, pAdjustType, pLineNo, pStockSale, pReverseCurrentOPSOnly, pRemoveInvoiceAllocations, pUpdateCashInvoiceAmountPaid, False, PartRefundAnalysis)
    End Function
    Private Function Reverse(ByVal pNewBatch As Batch, ByVal pNewTrans As BatchTransaction, ByVal pAdjustType As Batch.AdjustmentTypes, ByRef pLineNo As Integer, ByVal pStockSale As Boolean, ByVal pReverseCurrentOPSOnly As Boolean, ByVal pRemoveInvoiceAllocations As Boolean, ByVal pUpdateCashInvoiceAmountPaid As Boolean, ByVal pPosiveLinesOnly As Boolean, ByVal pPartRefundData As PartRefundAnalysis) As Boolean
      'Reverse this transaction
      'Assumed original batch is able to be adjusted
      'pReverseCurrentOPSOnly is used by BACS Messaging to indicate that the current OPS record will always be updated as outstanding
      Dim vRS As CDBRecordSet
      Dim vAllBTA As New Collection
      Dim vAllOPH As New Collection
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vNewBTA As BatchTransactionAnalysis = Nothing
      Dim vOldBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vOldFHD As New FinancialHistoryDetail(mvEnv)
      Dim vOldOPH As New OrderPaymentHistory
      Dim vCreditSale As New CreditSale(mvEnv)
      Dim vCardSale As New CardSale(mvEnv)
      Dim vDLU As New DeclarationLinesUnclaimed(mvEnv)
      Dim vGSLU As New GaSponsorshipLinesUnclaimed
      Dim vOPS As OrderPaymentSchedule
      Dim vPP As New PaymentPlan
      Dim vPGPledge As PreTaxPledge
      Dim vCP As CollectionPayment
      Dim vPIS As CollectionPIS
      Dim vSB As ServiceBooking = Nothing
      Dim vSBTransBTA As BatchTransactionAnalysis
      Dim vLinkAnalysis As Collection

      Dim vAdjust As Boolean
      Dim vAmount As Double
      Dim vCount As Integer
      Dim vDone As Boolean
      Dim vStatus As String = ""
      Dim vSQL As String
      Dim vFound As Boolean
      Dim vTrans As Boolean
      Dim vCannotAdjust As Integer
      Dim vCurrAmount As Double
      Dim vPass As Integer
      Dim vPass1OPH As New CDBParameters
      Dim vNumPass As Integer
      Dim vOPSNumber As Integer
      Dim vPGPledgeNo As Integer
      Dim vConfirmedTransactionChecked As Boolean

      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
      vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
      If pLineNo > 0 Then vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, pLineNo)
      If pPosiveLinesOnly Then vWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)

      'First, find all related financial history details
      vOldFHD.Init(mvEnv)
      vSQL = "SELECT " & vOldFHD.GetRecordSetFields(FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll) & " FROM financial_history_details fhd WHERE "
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY line_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        'All original fhd in mvDetails
        AddDetailFromRecordSet(vRS)
      End While
      vRS.CloseRecordSet()

      'Second, find all related order payment history
      vOldOPH.Init(mvEnv)
      vSQL = "SELECT " & vOldOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE "
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY line_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vOldOPH = New OrderPaymentHistory
        vOldOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
        vAllOPH.Add(vOldOPH)
      End While
      vRS.CloseRecordSet()

      'Third, find all related batch transaction analysis but exclude I type lines for which there is no FHD
      vWhereFields.Add("line_type", CDBField.FieldTypes.cftCharacter, "I", CDBField.FieldWhereOperators.fwoNotEqual)
      vOldBTA.Init()
      vSQL = "SELECT " & vOldBTA.GetRecordSetFields() & " FROM batch_transaction_analysis bta WHERE "
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY line_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vOldBTA = New BatchTransactionAnalysis(mvEnv)
        vOldBTA.InitFromRecordSet(vRS)
        vAllBTA.Add(vOldBTA)
      End While
      vRS.CloseRecordSet()

      'Fourth, find all related InvoicePaymentHistory
      'IPH may exist where Batch/Transaction/Line Numbers match the transaction / line being adjusted.
      'This will be the case when the transaction being adjusted contains an N-type (Invoice Payment) line.
      Dim vOldIPH As New InvoicePaymentHistory(mvEnv)
      Dim vAllIPH As New CollectionList(Of InvoicePaymentHistory)
      Dim vIPHWhereFields As New CDBFields(New CDBField("batch_number", BatchNumber))
      vIPHWhereFields.Add("transaction_number", TransactionNumber)
      If pLineNo > 0 Then vIPHWhereFields.Add("line_number", pLineNo)
      If pPosiveLinesOnly Then vIPHWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vOldIPH.GetRecordSetFields(), "invoice_payment_history iph", vIPHWhereFields, "line_number, status")
      vRS = vSQLStatement.GetRecordSet()
      Dim vKey As String
      While vRS.Fetch
        vOldIPH = New InvoicePaymentHistory(mvEnv)
        vOldIPH.InitFromRecordSet(vRS)
        vKey = (vAllIPH.Count + 1).ToString
        vAllIPH.Add(vKey, vOldIPH)
      End While
      vRS.CloseRecordSet()
      If vAllIPH.Count = 0 AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) = True Then
        'Otherwise IPH might exist where AllocationBatch/Transaction/Line Numbers match the transaction / line being adjusted.
        'This will be the case when the transaction being adjusted contains an L-type (Invoice payment allocation) line.
        With vIPHWhereFields
          .Clear()
          .Add("allocation_batch_number", BatchNumber)
          .Add("allocation_transaction_number", TransactionNumber)
          If pLineNo > 0 Then .Add("allocation_line_number", pLineNo)
          If pPosiveLinesOnly Then vIPHWhereFields.Add("amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End With
        vSQLStatement = New SQLStatement(mvEnv.Connection, vOldIPH.GetRecordSetFields(), "invoice_payment_history iph", vIPHWhereFields, "line_number")
        vRS = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          vOldIPH = New InvoicePaymentHistory(mvEnv)
          vOldIPH.InitFromRecordSet(vRS)
          vKey = (vAllIPH.Count + 1).ToString
          vAllIPH.Add(vKey, vOldIPH)
        End While
        vRS.CloseRecordSet()
      End If
      If pAdjustType = Batch.AdjustmentTypes.atPartRefund AndAlso (pPartRefundData IsNot Nothing AndAlso pPartRefundData.SalesLedgerPartRefund = True) Then
        'If we didn't find any IPH using Batch/Transaction/Line, try without the Line
        Dim vGetMoreIPH As Boolean = (vAllIPH.Count = 0 AndAlso pPartRefundData.InvoicePaymentAmount > 0)
        If vAllIPH.Count > 0 AndAlso pPartRefundData.InvoicePaymentAmount > 0 Then
          Dim vSumIPH As Double = 0
          For Each vOldIPH In vAllIPH
            vSumIPH += vOldIPH.Amount
          Next
          If pPartRefundData.InvoicePaymentAmount > FixTwoPlaces(vSumIPH) Then vGetMoreIPH = True
        End If
        If vGetMoreIPH Then
          vIPHWhereFields = New CDBFields(New CDBField("batch_number", BatchNumber))
          vIPHWhereFields.Add("transaction_number", TransactionNumber)
          vSQLStatement = New SQLStatement(mvEnv.Connection, vOldIPH.GetRecordSetFields(), "invoice_payment_history iph", vIPHWhereFields, "line_number")
          vRS = vSQLStatement.GetRecordSet()
          While vRS.Fetch
            vOldIPH = New InvoicePaymentHistory(mvEnv)
            vOldIPH.InitFromRecordSet(vRS)
            vKey = (vAllIPH.Count + 1).ToString
            vAllIPH.Add(vKey, vOldIPH)
          End While
          vRS.CloseRecordSet()
        End If
      End If

      'Check we have the correct signs & amounts
      If pAdjustType = Batch.AdjustmentTypes.atReverse AndAlso pNewBatch.BatchType = Batch.BatchTypes.CreditSales AndAlso pLineNo > 0 AndAlso TransactionSign = "D" AndAlso pNewTrans.TransactionSign = "D" Then
        'Reversing a reversal where the transaction type's the same (probably just cancelling an event booking that has a credit note attached to refund part of the booking amount)
        For Each vOldBTA In vAllBTA
          vOldBTA.ChangeSign()
        Next
      End If

      vPP.Init(mvEnv)
      'Now process the reversal - batch will have already been locked
      If Not (mvEnv.Connection.InTransaction) Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      If vAllBTA.Count() = 0 Then
        'If we found no batch transaction analysis lines then the batch may have been purged and it's an error
        RaiseError(DataAccessErrors.daeCannotFindBatchTransactionAnalysis, CStr(BatchNumber), CStr(TransactionNumber))
      Else
        'Check each batch transaction analysis line to make sure there is an FHD line for it
        For Each vOldBTA In vAllBTA
          vFound = False
          For Each vOldFHD In mvDetails
            If vOldFHD.LineNumber = vOldBTA.LineNumber Then
              vFound = True
              Exit For
            End If
          Next vOldFHD
          If Not vFound Then RaiseError(DataAccessErrors.daeCannotFindFinancialHistoryDetails, CStr(BatchNumber), CStr(TransactionNumber), CStr(vOldBTA.LineNumber))
        Next vOldBTA
      End If

      'If we found all the FHD records then proceed
      If vFound Then
        'Go through each of the old batch transaction analysis lines in turn
        'TA BR8115: Do 2 passes; 1=OPH status I, 2=all others
        vPass = 1
        vNumPass = If(pLineNo > 0, 1, 2) 'If we have a line-number then only process once, otherwise twice
        For vPass = 1 To vNumPass '2
          For Each vOldBTA In vAllBTA
            vAdjust = True
            'Find the corresponding financial history details line
            For Each vOldFHD In mvDetails
              If vOldFHD.LineNumber = vOldBTA.LineNumber Then
                Exit For
              End If
            Next vOldFHD
            'Now find the first OPH record for the BTA
            vFound = False
            For Each vOldOPH In vAllOPH
              If vOldOPH.LineNumber = vOldBTA.LineNumber Then
                vFound = True
                Exit For
              End If
            Next vOldOPH
            If Not vFound Then vOldOPH.Init(mvEnv)

            If (Len(vOldOPH.Status) > 0 And vOldOPH.Status <> "I") Then vAdjust = False
            If (vOldFHD.Status <> FinancialHistoryStatus.fhsNormal And vOldFHD.Status <> FinancialHistoryStatus.fhsOnBackOrder) Then vAdjust = False
            If pPosiveLinesOnly = True AndAlso vOldBTA.Amount < 0 Then vAdjust = False

            'Now find the first IPH record for the BTA
            vFound = False
            For Each vOldIPH In vAllIPH
              Select Case vOldBTA.LineType
                Case "N", "U"     'S/L Invoice Payment, S/L Unallocated Cash
                  vFound = (vOldIPH.BatchNumber = vOldBTA.BatchNumber AndAlso vOldIPH.TransactionNumber = vOldBTA.TransactionNumber AndAlso vOldIPH.LineNumber = vOldBTA.LineNumber)
                Case Else 'L      'S/L Allocation of Cash-Invoice
                  vFound = (vOldIPH.AllocationBatchNumber = vOldBTA.BatchNumber AndAlso vOldIPH.AllocationTransactionNumber = vOldBTA.TransactionNumber AndAlso vOldIPH.AllocationLineNumber = vOldBTA.LineNumber)
              End Select
              If vFound Then Exit For
            Next

            If vOldIPH IsNot Nothing AndAlso vOldIPH.Status.Length > 0 Then
              If TransactionSign = "D" AndAlso Status = FinancialHistoryStatus.fhsNormal Then
                'If we are reversing a credit note and the IPH allocating it has already been adjusted then allow the adjustment as this means the allocation has been removed from the Invoice
                Dim vInvoice As New Invoice()
                vInvoice.Init(mvEnv, BatchNumber, TransactionNumber)
                If Not (vInvoice.Existing = True AndAlso vInvoice.InvoiceType = Invoice.InvoiceRecordType.CreditNote) Then vAdjust = False
              ElseIf Not (vOldIPH.BatchNumber.Equals(vOldIPH.AllocationBatchNumber) AndAlso vOldIPH.TransactionNumber.Equals(vOldIPH.AllocationTransactionNumber) AndAlso
                            vOldIPH.LineNumber.Equals(vOldIPH.AllocationLineNumber)) Then
                'BR21351: IPH invoice allocation ('L' type line) reversed through remove allocations but original payment not reversed so allow adjustment
              Else
                vAdjust = False
              End If
            End If

            If Not vAdjust Then
              If pLineNo > 0 Then RaiseError(DataAccessErrors.daeCannotAdjustPayment, CStr(BatchNumber), CStr(TransactionNumber))
              vCannotAdjust = vCannotAdjust + 1
            End If

            'Only adjust In Advance OPH on the 1st pass; adjust all other lines on 2nd pass
            'But only if we are adjusting the entire transaction - i.e. do not do this if just adjusting a single analysis line
            If vAdjust = True And pLineNo = 0 Then
              Select Case vPass
                Case 1
                  vAdjust = vOldOPH.Status = "I"
                  If vAdjust = True Then vPass1OPH.Add(CStr(vOldOPH.LineNumber))
                Case 2
                  vAdjust = Not vPass1OPH.Exists(CStr(vOldOPH.LineNumber))
              End Select
            End If

            If vAdjust Then
              'Create BTA
              vNewBTA = New BatchTransactionAnalysis(mvEnv)
              vNewBTA.InitFromTransaction(pNewTrans)
              If pAdjustType = Batch.AdjustmentTypes.atPartRefund Then
                'pPartRefundData hold new values
                If pPartRefundData IsNot Nothing Then
                  If pPartRefundData.SalesLedgerPartRefund Then
                    'Part refund sales ledger payment
                    '(1) Validate the amounts before we start
                    Dim vCashInvoice As New Invoice()
                    vCashInvoice.Init(mvEnv, BatchNumber, TransactionNumber)
                    If vCashInvoice.InvoiceType <> Invoice.InvoiceRecordType.SalesLedgerCash Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "BatchNumber, TransactionNumber")
                    If pPartRefundData.UnallocatedAmount <> 0 Then
                      Dim vCashAmount As Double = vCashInvoice.InvoiceAmount  'Use this to better handle cash invoices created from a re-analysis
                      If Me.Amount.Equals(0) = True AndAlso vCashInvoice.IsFinancialAdjustmentInvoice = True Then vCashAmount = Me.Amount
                      If FixTwoPlaces(vCashAmount - vCashInvoice.AmountPaid) <> pPartRefundData.UnallocatedAmount Then
                        RaiseError(DataAccessErrors.daeParameterValueInvalid, "UnallocatedAmount")
                      End If
                    End If
                    If pPartRefundData.InvoicePaymentAmount <> 0 Then
                      Dim vSum As Double = 0
                      For Each vOldIPH In vAllIPH
                        vSum += vOldIPH.Amount
                      Next
                      If FixTwoPlaces(vSum) < pPartRefundData.InvoicePaymentAmount Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "InvoicePaymentAmount")
                    End If
                    If pPartRefundData.RefundAmount > vCashInvoice.InvoiceAmount Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "RefundAmount")

                    '(2) Refund any unallocated cash
                    Dim vRefundAmount As Double = pPartRefundData.RefundAmount
                    If pPartRefundData.UnallocatedAmount <> 0 Then
                      Dim vPayAmount As Double = If(pPartRefundData.UnallocatedAmount > vRefundAmount, vRefundAmount, pPartRefundData.UnallocatedAmount)
                      vCashInvoice.AmountUsed = vPayAmount
                      vCashInvoice.ProcessPayment((vPayAmount * -1))   'This will also update the Credit Customer
                      vCashInvoice.Save(mvEnv.User.UserID, True)
                      vRefundAmount -= vPayAmount
                    End If

                    '(3) Refund any invoice payments
                    If vRefundAmount > 0 AndAlso pPartRefundData.InvoicePaymentAmount <> 0 Then
                      'Need to reverse out InvoicePaymentHistory up to the amount we need to refund
                      Dim vInvoicePaid As New Invoice()
                      vInvoicePaid.Init(mvEnv)
                      For Each vOldIPH In vAllIPH
                        If vRefundAmount <= 0 Then Exit For
                        If vInvoicePaid.InvoiceNumber.Equals(vOldIPH.InvoiceNumber) = False Then
                          If vInvoicePaid.Existing Then vInvoicePaid.Save(mvEnv.User.UserID, True)
                          vInvoicePaid = New Invoice()
                          vInvoicePaid.Init(mvEnv, 0, 0, vOldIPH.InvoiceNumber)
                        End If
                        'Adjust Invoice AmountPaid and create negative IPH
                        vRefundAmount -= vOldIPH.Amount
                        vInvoicePaid.ProcessPayment((vOldIPH.Amount * -1))
                        vOldIPH.Reverse(FA_STATUS_ADJUSTMENT, vNewBTA.BatchNumber, vNewBTA.TransactionNumber, vNewBTA.LineNumber, CDate(pNewTrans.TransactionDate))
                      Next
                      If vRefundAmount < 0 Then
                        'Refunded too much so create a new positive IPH
                        Dim vIPHParams As New CDBParameters()
                        vIPHParams.Add("InvoiceNumber", vInvoicePaid.InvoiceNumber)
                        vIPHParams.Add("BatchNumber", vNewBTA.BatchNumber)
                        vIPHParams.Add("TransactionNumber", vNewBTA.TransactionNumber)
                        vIPHParams.Add("LineNumber", vNewBTA.LineNumber)
                        vIPHParams.Add("Amount", CDBField.FieldTypes.cftNumeric, Math.Abs(vRefundAmount).ToString)
                        vIPHParams.Add("AllocationDate", CDBField.FieldTypes.cftDate, pNewTrans.TransactionDate)
                        vIPHParams.Add("AllocationBatchNumber", vNewBTA.BatchNumber)
                        vIPHParams.Add("AllocationTransactionNumber", vNewBTA.TransactionNumber)
                        vIPHParams.Add("AllocationLineNumber", vNewBTA.LineNumber)
                        Dim vNewIPH As New InvoicePaymentHistory(mvEnv)
                        vNewIPH.Create(vIPHParams)
                        vNewIPH.Save(mvEnv.User.UserID, True)
                        vInvoicePaid.ProcessPayment(Math.Abs(vRefundAmount))
                        vRefundAmount = 0
                      End If
                      If vInvoicePaid.Existing Then vInvoicePaid.Save(mvEnv.User.UserID, True)
                    End If

                    '(4) Create new BTA for refund amount
                    vNewBTA.CloneForPartRefund(vOldBTA, vOldBTA.Quantity, vOldBTA.Issued, pNewBatch.CurrencyCode, pNewBatch.CurrencyExchangeRate, pPartRefundData.RefundAmount)
                    pRemoveInvoiceAllocations = False   'Already done this above so don't need to do it again
                  Else
                    'Part Refund product sale
                    vNewBTA.CloneForPartRefund(vOldBTA, pPartRefundData.RefundQuantity, pPartRefundData.RefundIssued, pNewBatch.CurrencyCode, pNewBatch.CurrencyExchangeRate, 0)
                  End If
                  vNewBTA.Save(mvEnv.User.UserID)
                End If
              Else
                vNewBTA.CloneFromBTA(vOldBTA)
                If pNewBatch.BatchType = Batch.BatchTypes.GiveAsYouEarn And vPGPledgeNo = 0 Then
                  If Val(vNewBTA.MemberNumber) > 0 Then vPGPledgeNo = IntegerValue(vNewBTA.MemberNumber) 'BTA.MemberNumber contains the PG Pledge Number
                End If
                'Check the accept_as_full
                If vOldOPH.LineNumber = vOldBTA.LineNumber Then
                  If vOldOPH.Balance <> 0 Then vNewBTA.AcceptAsFull = True
                End If
                'Set line type to 'P' if original line type related to pay plan but no history created (i.e. payment went against cancelled order product)
                Select Case vOldBTA.LineType
                  Case "M", "C", "O"
                    If Not vOldOPH.Existing Then
                      'Order line without any order_payment_history
                      vNewBTA.LineType = "P"
                      vNewBTA.ProductCode = vOldFHD.ProductCode
                      vNewBTA.RateCode = vOldFHD.RateCode
                      vNewBTA.Quantity = CInt(vOldFHD.Quantity)
                      vNewBTA.Issued = CInt(vOldFHD.Quantity)
                      vNewBTA.MemberNumber = ""
                    End If
                End Select
                vNewBTA.Save()
              End If
              vDone = True

              'Create Reversal
              vFields.Clear()
              With vFields
                .Add("batch_number", CDBField.FieldTypes.cftLong, vNewBTA.BatchNumber)
                .Add("transaction_number", CDBField.FieldTypes.cftLong, vNewBTA.TransactionNumber)
                .Add("line_number", CDBField.FieldTypes.cftLong, vNewBTA.LineNumber)
                .Add("was_batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                .Add("was_transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
                .Add("was_line_number", CDBField.FieldTypes.cftLong, vOldBTA.LineNumber)
                If vOldOPH.Existing Then
                  If Len(vOldOPH.Status) > 0 And vOldOPH.Status = "I" Then
                    .Add("was_oph_status", CDBField.FieldTypes.cftCharacter, vOldOPH.Status)
                  End If
                End If
              End With
              mvEnv.Connection.InsertRecord("reversals", vFields)

              'Update old FHD, OPH & IPH
              If pAdjustType = Batch.AdjustmentTypes.atMove Then
                Status = FinancialHistoryStatus.fhsMoved
                vStatus = "M"
              ElseIf pAdjustType = Batch.AdjustmentTypes.atAdjustment Or pAdjustType = Batch.AdjustmentTypes.atPartRefund Then
                Status = FinancialHistoryStatus.fhsAdjusted
                vStatus = "A"
              Else
                Status = FinancialHistoryStatus.fhsReversed
                vStatus = "R"
              End If
              With vFields
                .Clear()
                .Add("status", CDBField.FieldTypes.cftCharacter, vStatus)
              End With
              With vWhereFields
                .Clear()
                .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
                .Add("line_number", CDBField.FieldTypes.cftInteger, vOldFHD.LineNumber)
                .Add("status", CDBField.FieldTypes.cftCharacter, If(vOldFHD.Status = FinancialHistoryStatus.fhsOnBackOrder, "B", ""))
              End With
              mvEnv.Connection.UpdateRecords("financial_history_details", vFields, vWhereFields)

              If vOldOPH.Existing Then
                With vOldOPH
                  .Status = vStatus
                  .Save()
                End With
                'Create new oph and update ops
                If vPP.PlanNumber <> vOldOPH.OrderNumber Then vPP.Init(mvEnv, (vOldOPH.OrderNumber))
                If TransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlFirstClaimTransactionType) And vPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
                  If vPP.PaymentNumber = vOldOPH.PaymentNumber Then
                    vPP.DirectDebit.FirstClaim = True
                    vPP.DirectDebit.SetAmended((TodaysDate()), mvEnv.User.Logname)
                    vPP.DirectDebit.SaveChanges()
                  End If
                End If
                vPP.PaymentNumber = vPP.PaymentNumber + 1
                vPP.SaveChanges()

                If vOldOPH.ScheduledPaymentNumber.Length > 0 Then
                  vOPSNumber = IntegerValue(vOldOPH.ScheduledPaymentNumber)
                ElseIf pAdjustType <> Batch.AdjustmentTypes.atMove Then
                  'Move will be dealt with elsewhere
                  vOPS = New OrderPaymentSchedule
                  vOPS.Init(mvEnv)
                  vOPSNumber = vOPS.ReverseHistoricPayment(vPP, vOldOPH.Amount, True, pReverseCurrentOPSOnly)
                Else
                  vOPSNumber = 0
                End If
                If Len(vOldOPH.ScheduledPaymentNumber) > 0 And pAdjustType <> Batch.AdjustmentTypes.atMove Then
                  'Move will be dealt with elsewhere
                  vOPS = New OrderPaymentSchedule
                  vOPS.Init(mvEnv, IntegerValue(vOldOPH.ScheduledPaymentNumber))
                  vOPS.Reverse(vPP, vOldOPH.Amount, pReverseCurrentOPSOnly)
                  vOPS.Save()
                End If

                'Create the new OPH
                vOldOPH.Reverse(vNewBTA.BatchNumber, vNewBTA.TransactionNumber, vNewBTA.LineNumber, vPP.PaymentNumber, vOPSNumber)
                vOldOPH.Save()
              End If

              'Set the Status on the IPH record
              Dim vRemoveInvoiceAllocations As Boolean = False
              If pRemoveInvoiceAllocations = True AndAlso pLineNo > 0 Then vRemoveInvoiceAllocations = True
              If pAdjustType <> Batch.AdjustmentTypes.atPartRefund Then
                'Part refund has made all the required changes above so don't do any of this
                For Each vOldIPH In vAllIPH
                  With vOldIPH
                    If .Existing Then
                      Dim vReverse As Boolean = False
                      If vRemoveInvoiceAllocations Then
                        If (.BatchNumber = vOldBTA.BatchNumber AndAlso .TransactionNumber = vOldBTA.TransactionNumber AndAlso .LineNumber = vOldBTA.LineNumber AndAlso (.AllocationBatchNumber = 0 OrElse .AllocationBatchNumber = .BatchNumber)) _
                        OrElse (.BatchNumber <> .AllocationBatchNumber AndAlso .AllocationBatchNumber = vOldBTA.BatchNumber AndAlso .AllocationTransactionNumber = vOldBTA.TransactionNumber AndAlso .AllocationLineNumber = vOldBTA.LineNumber) Then
                          'Only set the Status if the Batch/Transaction/Line match and either both Batch & allocationBatch match or we only match on Allocation...
                          vReverse = True
                        End If
                      Else
                        If (.BatchNumber = vOldBTA.BatchNumber AndAlso .TransactionNumber = vOldBTA.TransactionNumber AndAlso .LineNumber = vOldBTA.LineNumber) _
                        OrElse (.AllocationBatchNumber = vOldBTA.BatchNumber AndAlso .AllocationTransactionNumber = vOldBTA.TransactionNumber AndAlso .AllocationLineNumber = vOldBTA.LineNumber) Then
                          'Only set the Status if the Batch/Transaction/Line match because there is no unique index on the IPH and so the update could fail on subsequent line numbers
                          If Not (vOldBTA.LineType.ToUpper.Equals("U") = True AndAlso .BatchNumber.Equals(.AllocationBatchNumber) = False) Then
                            'BTA.LineType = 'U' + IPH.BatchNumber <> IPH.AllocationBatchNumber Then do not reverse IPH here as it will be done in ReverseInvoiceCashAllocation
                            vReverse = True
                          End If
                        End If
                      End If
                      If vReverse = True AndAlso .Status.Length > 0 Then
                        'See if the reversal IPH has 0 for the numbers and if so update them to the adjustment numbers (caused by cancelling an event booking?)
                        vReverse = False
                        Dim vIPHAdjWhereFields As New CDBFields(New CDBField("invoice_number", .InvoiceNumber))
                        With vIPHAdjWhereFields
                          .Add("batch_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("transaction_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("line_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("allocation_batch_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("allocation_transaction_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("allocation_line_number", CDBField.FieldTypes.cftInteger, 0)
                          .Add("allocation_date", CDBField.FieldTypes.cftDate, TodaysDate)
                          .Add("amount", CDBField.FieldTypes.cftNumeric, (vOldIPH.Amount * -1))
                        End With
                        Dim vRevIPH As New InvoicePaymentHistory(mvEnv)
                        vRevIPH.Init()
                        Dim vIPHSQL As New SQLStatement(mvEnv.Connection, vRevIPH.GetRecordSetFields(), "invoice_payment_history iph", vIPHAdjWhereFields)
                        Dim vRevIPHRS As CDBRecordSet = vIPHSQL.GetRecordSet()
                        If vRevIPHRS.Fetch Then vRevIPH.InitFromRecordSet(vRevIPHRS)
                        vRevIPHRS.CloseRecordSet()
                        If vRevIPH.Existing Then
                          Dim vIPHParams As New CDBParameters()
                          With vIPHParams
                            .Add("BatchNumber", vNewBTA.BatchNumber)
                            .Add("TransactionNumber", vNewBTA.TransactionNumber)
                            .Add("LineNumber", vNewBTA.LineNumber)
                            .Add("AllocationBatchNumber", vNewBTA.BatchNumber)
                            .Add("AllocationTransactionNumber", vNewBTA.TransactionNumber)
                            .Add("AllocationLineNumber", vNewBTA.LineNumber)
                          End With
                          vRevIPH.Update(vIPHParams)
                          vRevIPH.Save(mvEnv.User.Logname, True)
                        End If
                      End If
                      If vReverse = True AndAlso vOldBTA.CashInvoiceNumber = 0 Then
                        Dim vInvAnsiJoins As New AnsiJoins()
                        vInvAnsiJoins.Add("reversals r", "cn.batch_number", "r.batch_number", "cn.transaction_number", "r.transaction_number")
                        vInvAnsiJoins.Add("invoices i", "r.was_batch_number", "i.batch_number", "r.was_transaction_number", "i.transaction_number")
                        Dim vInvWhereFields As New CDBFields(New CDBField("cn.batch_number", vOldBTA.BatchNumber))
                        vInvWhereFields.Add("cn.transaction_number", vOldBTA.TransactionNumber)
                        vInvWhereFields.Add("cn.record_type", Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.CreditNote))
                        vInvWhereFields.Add("i.record_type", Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.Invoice))
                        Dim vInvSQLStatement As New SQLStatement(mvEnv.Connection, String.Empty, "invoices cn", vInvWhereFields, String.Empty, vInvAnsiJoins)
                        If mvEnv.Connection.GetCountFromStatement(vInvSQLStatement) > 0 Then
                          'We are reversing a credit note that was created as a reversal of an invoice so keep the IPH unchanged (it's not a sundry credit note)
                          vReverse = False
                        End If
                      End If
                      If vReverse Then
                        If vOldIPH.Amount.Equals(vOldBTA.Amount) = False Then
                          If Math.Abs(vOldIPH.Amount).Equals(vOldBTA.Amount) AndAlso vOldBTA.Amount.CompareTo(0) > 0 AndAlso vOldBTA.LineType.Equals("K", StringComparison.InvariantCultureIgnoreCase) Then
                            'IPH is negative and BTA is positive for credit note allocation
                            If Not (Me.Amount.CompareTo(0) >= 0 AndAlso Me.TransactionSign.Equals("C", StringComparison.InvariantCultureIgnoreCase)) Then
                              vReverse = False
                            End If
                          End If
                        End If
                      End If
                      If vReverse Then
                        'vOldIPH is about to be reversed so it's sign will change, as it becomes vNewIPH
                        .Reverse(vStatus, vNewBTA.BatchNumber, vNewBTA.TransactionNumber, vNewBTA.LineNumber, CDate(pNewTrans.TransactionDate))
                      End If
                    End If
                  End With
                Next
                If vOldBTA.LineType.Equals("U") Then
                  Dim vIPHFound As Boolean = False
                  For Each vOldIPH In vAllIPH
                    If vOldIPH.BatchNumber.Equals(BatchNumber) AndAlso vOldIPH.TransactionNumber.Equals(vOldBTA.TransactionNumber) AndAlso vOldIPH.LineNumber.Equals(vOldBTA.LineNumber) Then
                      vIPHFound = True
                    End If
                    If vIPHFound Then Exit For
                  Next
                  If vIPHFound = False OrElse (Math.Abs(vOldIPH.Amount).Equals(Math.Abs(vOldBTA.Amount)) OrElse vOldIPH.BatchNumber.Equals(vOldIPH.AllocationBatchNumber) = False) Then
                    'vOldIPH may have been reversed by now so it's sign will have changed, as it becomes vNewIPH
                    ReverseInvoiceCashAllocation(vOldBTA.LineNumber, vOldBTA.LineType, vOldBTA.CurrencyAmount, vNewBTA, CDate(pNewTrans.TransactionDate), vStatus)
                  End If
                End If
                'Update status on cash invoice
                If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbInvoiceAdjustmentStatus) Then
                  Dim vInvoice As New Invoice()
                  If vOldBTA.CashInvoiceNumber > 0 Then
                    vInvoice.Init(mvEnv, 0, 0, vOldBTA.CashInvoiceNumber)
                  Else
                    vInvoice.Init(mvEnv, vOldBTA.BatchNumber, vOldBTA.TransactionNumber)
                  End If
                  If (vOldBTA.LineType = "L" OrElse vOldBTA.LineType = "K") Then
                    If vRemoveInvoiceAllocations = True AndAlso vInvoice.Existing = False AndAlso (vOldBTA.LineType = "L" OrElse vOldBTA.LineType = "K") Then
                      'Get the invoice
                      vOldIPH = New InvoicePaymentHistory(mvEnv)
                      vOldIPH.InitFromBatchTransactionLine(vOldBTA.BatchNumber, vOldBTA.TransactionNumber, vOldBTA.LineNumber)
                      If vOldIPH.Existing = True AndAlso vOldIPH.BatchNumber <> vOldIPH.AllocationBatchNumber Then
                        vInvoice = New Invoice()
                        vInvoice.Init(mvEnv, vOldIPH.BatchNumber, vOldIPH.TransactionNumber)
                      End If
                    ElseIf vRemoveInvoiceAllocations = False AndAlso vInvoice.Existing = True AndAlso vOldBTA.LineType.Equals("K", StringComparison.InvariantCultureIgnoreCase) Then
                      If Invoice.GetRecordType(vInvoice.RecordType) = Invoice.InvoiceRecordType.SalesLedgerCash Then
                        'Credit note allocation on a Cash Invoice transaction - don't update the Invoice
                        vInvoice = New Invoice()
                        vInvoice.Init(mvEnv)
                      End If
                    End If
                  End If
                  If vInvoice.Existing = True AndAlso (vInvoice.RecordType = Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash) OrElse vInvoice.IsSundryCreditNote) Then
                      If pAdjustType = Batch.AdjustmentTypes.atMove Then
                        vStatus = "M"
                      ElseIf pLineNo > 0 Then
                        vStatus = "A"
                      Else
                        vStatus = "R"
                      End If
                      vInvoice.SetAdjustmentStatus(Invoice.GetAdjustmentStatus(vStatus))
                      If TransactionSign.Equals("C", StringComparison.InvariantCultureIgnoreCase) AndAlso vOldBTA.Amount > 0 Then
                        Dim vPaid As Double = vOldBTA.Amount
                        If vRemoveInvoiceAllocations Then
                          Select Case vOldBTA.LineType
                            Case "N"
                              If pUpdateCashInvoiceAmountPaid = False Then
                                'BR16604: Invoice Payment line- where we do not want to increment amount paid. Currently this only applies when calling Remove Allocations from a (s/l cash) payment line.
                                'In RemoveAllocations the invoice type is cash so we aren't updating amount_paid and invoice_pay_status as we do for non-cash.
                                vPaid = 0
                              End If
                            Case "L", "K"
                              'S/L case Allocation - deduct amount paid
                              vPaid = (vOldBTA.Amount * -1)
                          End Select
                        Else
                          If pAdjustType = Batch.AdjustmentTypes.atReverse OrElse pAdjustType = Batch.AdjustmentTypes.atRefund Then
                            'Do not update the cash-invoice for an invoice payment
                            If vOldBTA.LineType.Equals("N", StringComparison.InvariantCultureIgnoreCase) Then vPaid = 0
                          End If
                        End If
                        vInvoice.SetAmountPaid(vPaid, True)
                        If pAdjustType.Equals(Batch.AdjustmentTypes.atReverse) AndAlso pLineNo.Equals(0) Then
                          'Reversed the whole transaction so vInvoice must be fully paid
                          If mvEnv.GetInvoicePayStatusType(vInvoice.InvoicePayStatus) <> CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid Then
                            vInvoice.SetAmountPaid(vInvoice.InvoiceAmount, True)
                          End If
                        End If
                        vInvoice.Save(mvEnv.User.UserID, True)
                        If pAdjustType = Batch.AdjustmentTypes.atMove AndAlso vInvoice.RecordType = Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash) _
                      AndAlso vOldBTA.LineType.Equals("U", StringComparison.InvariantCultureIgnoreCase) AndAlso vInvoice.InvoiceAmount.CompareTo(vOldBTA.Amount) > 0 Then
                          'Moving the unallocated cash so update the Credit Customer
                          'This is what would have happened when the transaction is first created
                          Dim vCC As New CreditCustomer()
                          vCC.InitCompanySalesLedgerAccount(mvEnv, vInvoice.Company, vInvoice.SalesLedgerAccount)
                          If vCC.Existing Then
                            vCC.AdjustOutstanding((vOldBTA.Amount * -1))
                            vCC.Save(mvEnv.User.UserID, True)
                          End If
                        End If
                      End If
                    End If
                  End If
                End If

              'BR21011: For reversal of invoice payment remove invoice allocations straight away, outside of batch posting
              If pAdjustType = Batch.AdjustmentTypes.atReverse OrElse pAdjustType = Batch.AdjustmentTypes.atRefund Then
                If vOldIPH.Amount = vOldBTA.Amount Then
                  Invoice.RemoveInvoiceAllocations(mvEnv, pNewTrans, vNewBTA, New Invoice, pNewBatch.BatchType, String.Empty, pAdjustType)
                End If
              End If

                'BR15547: Where the old BTA being reversed is within a confirmed transaction (where the provisional transaction will be re-instated in ConfirmedTransaction.ClearConfirmationForReversal)
                're-create the OPH record for the provisional transaction
                If Not vConfirmedTransactionChecked Then
                Dim vConfirmedTransaction As New ConfirmedTransaction(mvEnv)
                vConfirmedTransaction.InitConfirmed(vOldBTA.BatchNumber, vOldBTA.TransactionNumber)
                If vConfirmedTransaction.Existing Then
                  Dim vProvisionalBT As New BatchTransaction(mvEnv)
                  vProvisionalBT.InitBatchTransactionAnalysis(vConfirmedTransaction.ProvisionalBatchNumber, vConfirmedTransaction.ProvisionalTransNumber)
                  For Each vProvisionalBTA As BatchTransactionAnalysis In vProvisionalBT.Analysis
                    With vProvisionalBTA
                      Select Case .LineType
                        Case "M", "C", "O"
                          If vPP.PlanNumber <> .PaymentPlanNumber Then vPP.Init(mvEnv, .PaymentPlanNumber)
                          If vPP.Existing Then
                            vPP.PaymentNumber = vPP.PaymentNumber + 1
                            vPP.SaveChanges()
                            Dim vProvisionalOPH As New OrderPaymentHistory
                            vProvisionalOPH.Init(mvEnv)
                            Dim vScheduledPaymentNumber As Integer = 0
                            vOPS = New OrderPaymentSchedule
                            Dim vOPSFound As Boolean = False
                            If .Notes.Contains("Scheduled Payment Number: ") Then
                              'Get Scheduled Payment Number from provisional BTA notes
                              vScheduledPaymentNumber = IntegerValue(Mid(.Notes, .Notes.LastIndexOf(": ") + 3))
                            Else
                              'If the provisional transaction is historic ands created without the scheduled payment number in the notes field
                              'then simply get next Scheduled Payment record to pay
                              For Each vOPS In vPP.ScheduledPayments
                                If .Amount <= vOPS.AmountOutstanding Then
                                  vScheduledPaymentNumber = vOPS.ScheduledPaymentNumber
                                  vOPSFound = True
                                  Exit For
                                End If
                              Next
                            End If
                            vProvisionalOPH.SetValues(.BatchNumber, .TransactionNumber, vPP.PaymentNumber, .PaymentPlanNumber, .Amount, .LineNumber, 0, vScheduledPaymentNumber, False)
                            vProvisionalOPH.Save()
                            If Not vOPSFound Then vOPS.Init(mvEnv, vScheduledPaymentNumber)
                            If vOPS.Existing Then
                              vOPS.SetUnProcessedPayment(True, .Amount)
                              vOPS.Save()
                            End If
                          End If
                      End Select
                    End With
                  Next
                End If
                vConfirmedTransactionChecked = True
              End If
              'Check for any DLU or DTCL lines that need deleting/changing
              With vFields
                .Clear()
                .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
                .Add("line_number", CDBField.FieldTypes.cftLong, vOldBTA.LineNumber)
              End With

              'Check for any Gift Aid Sponsorship lines that need deleting / changing
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidSponsorship) Then
                mvEnv.Connection.DeleteRecords("ga_sponsorship_lines_unclaimed", vFields, False)
                vGSLU.Init(mvEnv)
                vGSLU.CreateNewNegativeLines(vNewBTA.BatchNumber, vNewBTA.TransactionNumber, vNewBTA.LineNumber, BatchNumber, TransactionNumber, vOldBTA.LineNumber)
              End If

              'If this is a Service Booking Transaction then need to add the reversal transactions to a collection
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then
                vWhereFields.Clear()
                With vWhereFields
                  .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
                  .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
                  .Add("line_number", CDBField.FieldTypes.cftLong, vOldBTA.LineNumber)
                End With
                If mvEnv.Connection.GetCount("service_booking_transactions", vWhereFields) > 0 Then
                  vSQL = "SELECT service_booking_number FROM service_booking_transactions WHERE batch_number = " & BatchNumber
                  vSQL = vSQL & " AND transaction_number = " & TransactionNumber
                  vSQL = vSQL & " AND line_number = " & vOldBTA.LineNumber
                  vRS = mvEnv.Connection.GetRecordSet(vSQL)
                  While vRS.Fetch() = True
                    vSB = New ServiceBooking
                    vSB.Init(mvEnv, (vRS.Fields(1).IntegerValue))
                  End While
                  vRS.CloseRecordSet()

                  vSBTransBTA = New BatchTransactionAnalysis(mvEnv)
                  vSBTransBTA.Init()
                  vSQL = "SELECT " & vSBTransBTA.GetRecordSetFields()
                  vSQL = vSQL & " FROM batch_transaction_analysis bta WHERE batch_number = " & vNewBTA.BatchNumber
                  vSQL = vSQL & " AND transaction_number = " & vNewBTA.TransactionNumber
                  vSQL = vSQL & " AND line_number = " & vNewBTA.LineNumber
                  vRS = mvEnv.Connection.GetRecordSet(vSQL)
                  While vRS.Fetch() = True
                    vSBTransBTA = New BatchTransactionAnalysis(mvEnv)
                    vSBTransBTA.InitFromRecordSet(vRS)
                    vSBTransBTA.LinkedBookingNo = vSB.ServiceBookingNumber
                    vLinkAnalysis = New Collection
                    vLinkAnalysis.Add(vSBTransBTA)
                    vSB.AddLinkedTransaction(pAdjustType, vLinkAnalysis)
                  End While
                  vRS.CloseRecordSet()
                End If
              End If

              If pAdjustType = Batch.AdjustmentTypes.atRefund AndAlso pNewBatch.BatchType = Batch.BatchTypes.CreditSales Then
                'Refunding an Invoice - re-set the status of any Exam Student Exemptions
                Dim vExamStudentExamption As New ExamStudentExemption(mvEnv)
                If (Me.TransactionSign.Equals("C", StringComparison.InvariantCultureIgnoreCase) AndAlso BooleanValue(Me.NegativesAllowed) = False) Then
                  'Invoice is being cancelled
                  vExamStudentExamption.RefundExemptionInvoice(BatchNumber, TransactionNumber, 0, 0)
                Else
                  'Credit note being cancelled
                  vExamStudentExamption.RefundExemptionInvoice(BatchNumber, TransactionNumber, pNewTrans.BatchNumber, pNewTrans.TransactionNumber)
                End If
              End If

            End If
          Next vOldBTA 'Next vOldBTA
        Next
      End If

      If vDone Then
        'Save the BT record
        pNewTrans.SaveChanges()
        'Create Credit Sales
        If pNewBatch.BatchType = Batch.BatchTypes.CreditSales Then
          If pLineNo > 0 And vNewBTA.LineNumber > 1 Then vCreditSale.Init((vNewBTA.BatchNumber), (vNewBTA.TransactionNumber)) 'Could be cancelling an Event Booking
          If vCreditSale.Existing = False Then
            vCreditSale = New CreditSale(mvEnv)
            vCreditSale.Init(BatchNumber, TransactionNumber)
            If vCreditSale.Existing Then
              vCreditSale.Clone((vNewBTA.BatchNumber), (vNewBTA.TransactionNumber))
            Else
              RaiseError(DataAccessErrors.daeCannotFindCreditSale, CStr(BatchNumber), CStr(TransactionNumber))
            End If
          End If
          If pLineNo > 0 Then
            vCreditSale.StockSale = pStockSale
          Else
            vCreditSale.StockSale = False
          End If
          vCreditSale.Save()
        End If

        'Check Card Sales
        If pNewBatch.BatchType = Batch.BatchTypes.CreditCard OrElse pNewBatch.BatchType = Batch.BatchTypes.DebitCard OrElse pNewBatch.BatchType = Batch.BatchTypes.CreditCardWithInvoice Then
          'This is handled in AdjustTransaction
          vCardSale.Init(BatchNumber, TransactionNumber)
          If Not vCardSale.Existing Then RaiseError(DataAccessErrors.daeCannotFindCardSale, CStr(BatchNumber), CStr(TransactionNumber))
        End If

        'Create PG Pledge reversal Payment History
        If pNewBatch.BatchType = Batch.BatchTypes.GiveAsYouEarn Then
          vPGPledge = New PreTaxPledge(mvEnv)
          vPGPledge.Init(vPGPledgeNo)
          'PGPledge.ReversePayment will Init class from PaymentHistory if vPGPledgeNo is 0
          'and also raise any errors if the Pledge or PaymentHistory is missing
          vPGPledge.ReversePayment(BatchNumber, TransactionNumber, vNewBTA.BatchNumber, vNewBTA.TransactionNumber)
          If vPGPledge.Existing Then vPGPledge.Save()
        End If

        'If this is an AppealCollection payment then need to add a reversal CollectionPayment
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
          vCP = New CollectionPayment
          vCP.Init(mvEnv)
          vSQL = "SELECT " & vCP.GetRecordSetFields(CollectionPayment.CollectionPaymentRecordSetTypes.cpyrtAll) & " FROM collection_payments cp WHERE batch_number = " & BatchNumber
          vSQL = vSQL & " AND transaction_number = " & TransactionNumber
          If pLineNo > 0 Then vSQL = vSQL & " AND line_number = " & pLineNo
          vSQL = vSQL & " ORDER BY line_number"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          While vRS.Fetch() = True
            vCP = New CollectionPayment
            vCP.InitFromRecordSet(mvEnv, vRS, CollectionPayment.CollectionPaymentRecordSetTypes.cpyrtAll)
            vCP.Reverse(vNewBTA.BatchNumber, vNewBTA.TransactionNumber, If(pLineNo > 0, vNewBTA.LineNumber, pLineNo))
            vCP.Save()
            If vCP.CollectionPisNumber > 0 Then
              vPIS = New CollectionPIS
              vPIS.Init(mvEnv, (vCP.CollectionPisNumber))
              vPIS.Reconcile(CollectionPIS.CollectionPISReconciledStatus.cpisrsReversed)
              vPIS.Save()
            End If
          End While
          vRS.CloseRecordSet()
        End If

        'Update Batch - first set the correct batch total
        vAmount = 0
        vCount = 0
        vCurrAmount = 0
        Dim vBatchDetails As ParameterList = GetBatchTotal(pNewBatch.BatchNumber)
        If vBatchDetails IsNot Nothing Then
          If vBatchDetails.ContainsKey("Amount") Then vAmount = Convert.ToDouble(vBatchDetails("Amount"))
          If vBatchDetails.ContainsKey("CurrencyAmount") Then vCurrAmount = Convert.ToDouble(vBatchDetails("CurrencyAmount"))
          If vBatchDetails.ContainsKey("Count") Then vCount = CInt(vBatchDetails("Count"))
        End If

        With pNewBatch
          .BatchTotal = vAmount
          .CurrencyBatchTotal = vCurrAmount
          .NumberOfTransactions = vCount
          If .NumberOfEntries > 0 Then .NumberOfEntries = 0
          .SetBatchTotals() 'Force re-setting of the batch_totals
          .Save()
        End With

        'Update this FH class
        If pAdjustType = Batch.AdjustmentTypes.atMove Then
          Status = FinancialHistoryStatus.fhsMoved
        ElseIf pLineNo > 0 Then
          Status = FinancialHistoryStatus.fhsAdjusted
          vStatus = "A"
        Else
          Status = FinancialHistoryStatus.fhsReversed
        End If
        With vFields
          .Clear()
          .Add("status", CDBField.FieldTypes.cftCharacter, vStatus)
        End With
        With vWhereFields
          .Clear()
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End With
        mvEnv.Connection.UpdateRecords("financial_history", vFields, vWhereFields)
      Else
        If vCannotAdjust > 0 Then RaiseError(DataAccessErrors.daeCannotAdjustPayment, CStr(BatchNumber), CStr(TransactionNumber))
      End If
      If vTrans Then mvEnv.Connection.CommitTransaction()
      Return vDone
    End Function

    Public Function StatusDesc(ByRef pStatusCode As String) As String
      'Given a Status Code, return a description
      Dim vStatusDesc As String = ""

      Select Case pStatusCode
        Case FA_STATUS_ADJUSTMENT
          vStatusDesc = (ProjectText.String29504) 'Adjusted
        Case FA_STATUS_MOVE
          vStatusDesc = (ProjectText.String29505) 'Moved to another Contact
        Case FA_STATUS_REVERSAL
          vStatusDesc = (ProjectText.String29506) 'Reversed
        Case FA_STATUS_BACK_ORDER
          vStatusDesc = (ProjectText.String29507) 'Placed on Back Order
        Case FA_STATUS_IN_ADVANCE
          vStatusDesc = (ProjectText.String29508) 'Paid in Advance
      End Select
      StatusDesc = vStatusDesc
    End Function

    Public Function CheckFAInAdvance(Optional ByVal pLineNumber As Integer = 0, Optional ByRef pMsg As String = "", Optional ByVal pReversalAdjustment As Boolean = False) As Boolean
      'Check in-advance for FA - return True (if all OK) of False + pMsg (if can not be adjusted)
      Dim vSQL As String
      Dim vMsg As String = ""
      Dim vFound As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vBatchNumber As Integer
      Dim vTransactionNumber As Integer
      Dim vRaiseError As Boolean

      'First, check whether this is an in-advance batch that has already been reversed.
      vSQL = "oph.batch_number = " & BatchNumber & " AND oph.transaction_number = " & TransactionNumber
      vSQL = vSQL & " AND r.batch_number = oph.batch_number AND r.transaction_number = oph.transaction_number"
      vSQL = vSQL & " AND r.line_number = oph.line_number AND r.was_oph_status = 'I'"
      If mvEnv.Connection.GetCount("order_payment_history oph, reversals r", Nothing, vSQL) > 0 Then
        'Payment already reversed - can not adjust further
        vFound = False
        vMsg = "This transaction contains an in-advance payment that has already been adjusted and therefore can not be further adjusted."
      Else
        'OK
        vSQL = "SELECT oph.order_number,status FROM order_payment_history oph, orders o WHERE oph.batch_number = " & BatchNumber & " AND oph.transaction_number = " & TransactionNumber
        If pLineNumber > 0 Then vSQL = vSQL & " AND oph.line_number = " & pLineNumber
        'vSQL = vSQL & " AND (oph.status IS NULL OR oph.status = 'I' OR oph.status = 'B') AND o.order_number = oph.order_number AND o.in_advance > 0"
        vSQL = vSQL & " AND oph.order_number = o.order_number AND (((oph.status IS NULL OR oph.status = 'I') AND o.in_advance > 0) OR (oph.status = 'B' AND o.in_advance >= 0))"
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        While vRecordSet.Fetch() = True
          'If Status is "I" that means you're about to adjust the actual in-advance payment, so OK
          'If Status is "B" that means the transaction you're trying to adjust was an in-advance payment that has already been allocated against the pay plan
          'If Status is null you need to make sure the payment plan doesn't have any un-adjusted in-advance payments
          If vRecordSet.Fields("status").Value <> "I" Then
            If vRecordSet.Fields("status").Value = "" Then
              vFound = False
              vRecordSet2 = mvEnv.Connection.GetRecordSet("SELECT batch_number,transaction_number,line_number FROM order_payment_history WHERE order_number = " & vRecordSet.Fields("order_number").Value & " AND status = 'I'")
              While vRecordSet2.Fetch() = True
                vBatchNumber = vRecordSet2.Fields("batch_number").IntegerValue
                vTransactionNumber = vRecordSet2.Fields("transaction_number").IntegerValue
                'BR 8115:If reversing the whole transaction, only error if there are In-Advance payments relating to another transaction.
                vRaiseError = True
                If pReversalAdjustment = True And pLineNumber = 0 And vBatchNumber = BatchNumber And vTransactionNumber = TransactionNumber Then
                  vRaiseError = False
                End If
                If vRaiseError = True Then
                  If Not vFound Then
                    vMsg = "Payment Plan " & vRecordSet.Fields("order_number").Value & " has an in-advance payment.  The following payments must be adjusted first:" & vbCrLf & vbCrLf
                  Else
                    vMsg = vMsg & vbCrLf
                  End If
                  vMsg = vMsg & vbTab & "Batch :" & vBatchNumber & ", Transaction :" & vTransactionNumber & ", Line Number :" & vRecordSet2.Fields("line_number").Value
                End If
                vFound = True
              End While
              vRecordSet2.CloseRecordSet()
              If Not vFound Then
                If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf
                vMsg = "Payment Plan " & vRecordSet.Fields("order_number").Value & " has an in-advance payment, but the corresponding payment record cannot be found"
              End If
            Else
              vMsg = "This transaction contains an in-advance payment that has been allocated against the payment plan and therefore cannot be adjusted"
            End If
          End If
        End While
        vRecordSet.CloseRecordSet()
      End If

      pMsg = vMsg
      CheckFAInAdvance = (vMsg = "")
    End Function

    Public Function CanAddGiftAidDeclaration() As Boolean
      Dim vFields As New CDBFields
      Dim vGADeclaration As New GiftAidDeclaration
      Dim vPayPlan As New PaymentPlan
      Dim vRS As CDBRecordSet
      Dim vAdd As Boolean
      Dim vSQL As String

      'First check transaction is correct
      If BatchNumber > 0 Then
        Dim vBT As New BatchTransaction(mvEnv)
        vBT.Init(BatchNumber, TransactionNumber)
        If vBT.EligibleForGiftAid Then
          vGADeclaration.Init(mvEnv, pRaiseNoGAControlError:=False)
          vAdd = vGADeclaration.GADControlsExists
          If vAdd Then
            If CDate(TransactionDate) >= CDate(vGADeclaration.GiftAidEarliestStartDate) Then
              If PaymentMethod <> vGADeclaration.CAFPaymentMethod Then
                vAdd = True
              End If
            End If
          End If
        End If
      End If

      'Second see if an existing Declaration links to this payment
      If vAdd Then
        With vFields
          .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End With
        If mvEnv.Connection.GetCount("gift_aid_declarations", vFields) > 0 Then vAdd = False
      End If

      'Third, Check that transaction could be eligible for a Declaration
      If vAdd Then
        vFields.Remove("contact_number")
        'Select any Order Payment History
        vPayPlan.Init(mvEnv)
        vFields.Add("oph.order_number", CDBField.FieldTypes.cftLong, "o.order_number")
        vSQL = Replace(vPayPlan.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstAll), "payment_number", "o.payment_number")
        vSQL = "SELECT " & vSQL & " FROM order_payment_history oph, orders o WHERE " & mvEnv.Connection.WhereClause(vFields)
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        vRS.Fetch()
        If vRS.Status() = True Then
          'There is a Payment Plan
          'Loop through the records looking for an eligible Payment Plan
          Do
            vPayPlan.InitFromRecordSet(mvEnv, vRS, PaymentPlan.PayPlanRecordSetTypes.pprstAll)
            If vPayPlan.PlanType = CDBEnvironment.ppType.pptMember Then
              vAdd = vPayPlan.MembershipEligibleForGiftAid(TransactionDate)
            End If
          Loop While vRS.Fetch() = True And vAdd = False
        Else
          vAdd = False
        End If
        vRS.CloseRecordSet()

        If vAdd = False Then
          'No (eligible) Payment Plans
          With vFields
            .Remove("oph.order_number")
            .Add("fhd.product", CDBField.FieldTypes.cftLong, "p.product")
            .Add("donation", CDBField.FieldTypes.cftCharacter, "Y")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductEligibleGA) Then
              .Add("eligible_for_gift_aid", CDBField.FieldTypes.cftCharacter, "Y")
            End If
          End With
          If mvEnv.Connection.GetCount("financial_history_details fhd, products p", vFields) > 0 Then 'Some donation Products
            vAdd = True
          End If
        End If
      End If

      CanAddGiftAidDeclaration = vAdd

    End Function

    ''' <summary>Is this transaction a confirmation of a provisional batch?</summary>
    Public Function IsConfirmedTransaction() As Boolean
      Dim vWhereFields As New CDBFields

      If Not mvConfirmedTransChecked Then
        vWhereFields.Add("confirmed_batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
        vWhereFields.Add("confirmed_trans_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        'Changes made to return True only if the Transaction confirmed and is CAF type
        'else retunr false as "Change Payer" menu option needs to be displayed (BR15229)
        If mvEnv.Connection.GetCount("confirmed_transactions", vWhereFields) > 0 AndAlso (BatchType() = Batch.BatchTypes.CAFCards OrElse BatchType() = Batch.BatchTypes.CAFVouchers OrElse BatchType() = Batch.BatchTypes.CAFCommitmentReconciliation) Then
          mvIsConfirmedTransaction = True
        Else
          mvIsConfirmedTransaction = False
        End If
        mvConfirmedTransChecked = True
      End If
      IsConfirmedTransaction = mvIsConfirmedTransaction
    End Function

    Public Function CanMove() As Boolean
      Dim vCan As Boolean

      vCan = TransactionSign <> "D" And Len(mvEnv.GetConfig("trader_application_fa")) > 0
      If vCan Then
        vCan = Not (PaymentMethod.In(mvEnv.GetConfig("pm_dd"),
                                     mvEnv.GetConfig("pm_ccca")))
      End If
      If vCan Then
        vCan = Not (Status.In(FinancialHistoryStatus.fhsMoved,
                              FinancialHistoryStatus.fhsReversed,
                              FinancialHistoryStatus.fhsAdjusted))
      End If
      If vCan Then
        vCan = Not (BatchType.In(Batch.BatchTypes.GiveAsYouEarn,
                                 Batch.BatchTypes.PostTaxPayrollGiving,
                                 Batch.BatchTypes.DirectCredit,
                                 Batch.BatchTypes.CreditSales,
                                 Batch.BatchTypes.BankStatement))
      End If
      If vCan Then
        vCan = Not IsConfirmedTransaction()
      End If
      If vCan Then
        vCan = Not ContainsInMemoriamPPPayment()
      End If
      If vCan = True AndAlso BatchType() <> Batch.BatchTypes.DirectDebit Then
        'Cannot move the transaction when:
        'i)  Transaction contains SL Cash Allocations or Credit Note Allocations
        'ii) Transaction is the result of a financial adjustment and it contains any type of Sales Ledger item
        Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber)
        If (vSLItems.HasFlag(SalesLedgerItems.SLCashAllocation) OrElse vSLItems.HasFlag(SalesLedgerItems.CreditNoteAllocation)) Then
          vCan = False
        End If
        If vCan Then
          If IsFinancialAdjustment = True AndAlso vSLItems > SalesLedgerItems.None Then
            vCan = False
          End If
        End If
      End If

      If vCan Then
        'Check against the financial control 'one reversal only' value
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneReversalOnly) = "Y" _
          And IsFinancialAdjustment = True Then
          'Prevent reversal
          vCan = False
        End If
      End If

      CanMove = vCan
    End Function
    Public Function CanReverse() As Boolean
      Dim vCan As Boolean

      vCan = Not IsEventBooking()
      If vCan Then vCan = Not ContainsExamBooking()
      If vCan Then vCan = Not PaymentMethod = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod)
      If vCan Then vCan = (Status <> FinancialHistoryStatus.fhsMoved And Status <> FinancialHistoryStatus.fhsReversed And Status <> FinancialHistoryStatus.fhsAdjusted)
      If vCan Then vCan = BatchType() <> Batch.BatchTypes.DirectCredit
      If vCan Then vCan = Not ContainsInMemoriamPPPayment()

      If vCan = True AndAlso BatchType() <> Batch.BatchTypes.DirectDebit Then
        'Cannot reverse the transaction when:
        'i)   Transaction contains SL Cash Allocations or Credit Note Allocations
        'ii)  Transaction is the result of a financial adjustment and it contains any type of Sales Ledger item
        'iii) The original transaction was not a re-analysis
        Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber)
        If (vSLItems.HasFlag(SalesLedgerItems.SLCashAllocation) OrElse vSLItems.HasFlag(SalesLedgerItems.CreditNoteAllocation)) Then vCan = False
        If vCan Then
          If IsFinancialAdjustment = True AndAlso vSLItems > SalesLedgerItems.None Then
            If CanReverseAdjustedTransaction(vSLItems) = False Then vCan = False
          End If
        End If
      End If

      If vCan Then
        'Check against the financial control 'one reversal only' value
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneReversalOnly) = "Y" _
          And IsFinancialAdjustment = True Then
          'Prevent reversal only if original is not a reversible Adjusted Transaction
          If CanReverseAdjustedTransaction() = False Then
            vCan = False
          End If
        End If
      End If

      CanReverse = vCan
    End Function
    Public Function CanRefund() As Boolean
      Dim vCan As Boolean

      vCan = Not IsEventBooking()
      If vCan Then vCan = Not ContainsExamBooking()
      If vCan Then vCan = PaymentMethod = mvEnv.GetConfig("pm_dd") Or PaymentMethod = mvEnv.GetConfig("pm_cc") Or PaymentMethod = mvEnv.GetConfig("pm_dc") Or PaymentMethod = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod) Or PaymentMethod = mvEnv.GetConfig("pm_ccca")
      If vCan Then vCan = (Status <> FinancialHistoryStatus.fhsMoved And Status <> FinancialHistoryStatus.fhsReversed And Status <> FinancialHistoryStatus.fhsAdjusted)
      If vCan Then vCan = (BatchType() <> Batch.BatchTypes.GiveAsYouEarn And BatchType() <> Batch.BatchTypes.PostTaxPayrollGiving And BatchType() <> Batch.BatchTypes.DirectCredit And BatchType() <> Batch.BatchTypes.FinancialAdjustment)
      If vCan Then vCan = Not ContainsInMemoriamPPPayment()
      If vCan = True AndAlso BatchType() <> Batch.BatchTypes.DirectDebit Then
        'Cannot refund the transaction when:
        'i)   Transaction contains SL Cash Allocations or Credit Note Allocations
        'ii)  Transaction is the result of a financial adjustment and it contains any type of Sales Ledger item
        'iii) The original transaction was not a re-analysis
        Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber)
        If (vSLItems.HasFlag(SalesLedgerItems.SLCashAllocation) OrElse vSLItems.HasFlag(SalesLedgerItems.CreditNoteAllocation)) Then vCan = False
        If vCan Then
          If IsFinancialAdjustment = True AndAlso vSLItems > SalesLedgerItems.None Then
            If CanReverseAdjustedTransaction(vSLItems) = False Then vCan = False
          End If
        End If
      End If

      If vCan Then
        'Check against the financial control 'one reversal only' value
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneReversalOnly) = "Y" _
          And IsFinancialAdjustment = True Then
          'Prevent reversal only if original is not a reversible Adjusted Transaction
          If CanReverseAdjustedTransaction() = False Then
            vCan = False
          End If
        End If
      End If

      CanRefund = vCan
    End Function
    Public Function CanReanalyse() As Boolean
      Dim vCan As Boolean

      vCan = Len(mvEnv.GetConfig("trader_application_fa")) > 0
      If vCan And mvEnv.GetConfigOption("fp_use_sales_ledger", True) Then vCan = Not PaymentMethod = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod)
      If vCan Then vCan = (Status <> FinancialHistoryStatus.fhsMoved And Status <> FinancialHistoryStatus.fhsReversed And Status <> FinancialHistoryStatus.fhsAdjusted)
      If vCan Then vCan = (BatchType() <> Batch.BatchTypes.GiveAsYouEarn And BatchType() <> Batch.BatchTypes.PostTaxPayrollGiving And BatchType() <> Batch.BatchTypes.DirectCredit)
      If vCan Then vCan = Not ContainsInMemoriamPPPayment()
      If vCan = True AndAlso BatchType() <> Batch.BatchTypes.DirectDebit Then
        'Cannot re-analyse the transaction when:
        'i)  Transaction contains SL Cash Allocations or Credit Note Allocations
        'ii) Transaction is the result of a financial adjustment and it contains any type of Sales Ledger item
        Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber)
        If (vSLItems.HasFlag(SalesLedgerItems.SLCashAllocation) OrElse vSLItems.HasFlag(SalesLedgerItems.CreditNoteAllocation)) Then vCan = False
        If vCan Then
          If IsFinancialAdjustment = True AndAlso vSLItems > SalesLedgerItems.None Then vCan = False
        End If
      End If

      If vCan Then
        'Check against the financial control 'one reversal only' value
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneReversalOnly) = "Y" _
          And IsFinancialAdjustment = True Then
          'Prevent reversal
          vCan = False
        End If
      End If

      CanReanalyse = vCan
    End Function

    Private Function IsEventBooking() As Boolean
      If Not mvCheckedForEventBooking Then
        Dim vWhereFields As New CDBFields({New CDBField("batch_number", BatchNumber), New CDBField("transaction_number", TransactionNumber)})
        mvIsEventBooking = mvEnv.Connection.GetCount("event_bookings", vWhereFields) > 0
        If mvIsEventBooking = False Then mvIsEventBooking = IsEventBookingTransaction()
        mvCheckedForEventBooking = True
      End If
      Return mvIsEventBooking
    End Function

    ''' <summary>Is this a transaction created as a result of amending an Event Booking?</summary>
    ''' <returns>True if the transaction was created by amending an Event Booking, otherwise False.</returns>
    Private Function IsEventBookingTransaction() As Boolean
      Dim vIsEventBookingTransaction As Boolean = False

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("event_bookings eb", "ebt.event_number", "eb.event_number", "ebt.booking_number", "eb.booking_number")})
      vAnsiJoins.Add("financial_history fh", "ebt.batch_number", "fh.batch_number", "ebt.transaction_number", "fh.transaction_number")

      Dim vWhereFields As New CDBFields({New CDBField("ebt.batch_number", BatchNumber), New CDBField("ebt.transaction_number", TransactionNumber)})
      vWhereFields.Add("ebt.batch_number#2", CDBField.FieldTypes.cftInteger, "eb.batch_number", CDBField.FieldWhereOperators.fwoNotEqual)


      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "eb.batch_number, eb.transaction_number, tt.transaction_sign, bta.line_type", "event_booking_transactions ebt", vWhereFields, "ebt.line_number", vAnsiJoins)
      If mvEnv.Connection.GetCountFromStatement(vSQLStatement) > 0 Then vIsEventBookingTransaction = True

      Return vIsEventBookingTransaction
    End Function

    Public Function ContainsExamBooking() As Boolean
      Dim vWhereFields As New CDBFields

      If Not mvCheckedForExamBooking Then
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then
          With vWhereFields
            .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
          End With
          mvContainsExamBooking = mvEnv.Connection.GetCount("exam_bookings", vWhereFields) > 0
        End If
        mvCheckedForExamBooking = True
      End If
      Return mvContainsExamBooking
    End Function

    Public Function ContainsInMemoriamPPPayment() As Boolean
      Dim vWhereFields As New CDBFields

      If Not mvCheckedForInMemoriamPPPayment Then
        With vWhereFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
          .Add("line_type", CDBField.FieldTypes.cftCharacter, "G")
          .Add("order_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
        End With
        mvContainsInMemoriamPPPayment = mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) > 0
        mvCheckedForInMemoriamPPPayment = True
      End If
      ContainsInMemoriamPPPayment = mvContainsInMemoriamPPPayment
    End Function
    Public Function BatchType() As Batch.BatchTypes
      If Not mvGotBatchType Then
        mvBatchType = Batch.GetBatchType(mvEnv.Connection.GetValue("SELECT batch_type FROM batches WHERE batch_number = " & BatchNumber))
        mvGotBatchType = True
      End If
      BatchType = mvBatchType
    End Function
    Public Function AdjustmentState() As FinancialHistory.AdjustmentStates
      Dim vWhereFields As New CDBFields
      Dim vRS As CDBRecordSet

      If Not mvAdjustmentStateSet Then
        mvAdjustmentState = FinancialHistory.AdjustmentStates.adjsNone
        mvAdjustmentBatchNumber = 0
        mvAdjustmentTransactionNumber = 0
        mvAdjustmentWasBatchNumber = 0
        mvAdjustmentWasTransactionNumber = 0

        'Find the transaction that this financial history changed
        With vWhereFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End With
        vRS = mvEnv.Connection.GetRecordSet("SELECT was_batch_number, was_transaction_number FROM reversals WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        With vRS
          If .Fetch() = True Then
            mvAdjustmentState = FinancialHistory.AdjustmentStates.adjsIsAnAdjustment
            mvAdjustmentWasBatchNumber = .Fields.Item(1).IntegerValue
            mvAdjustmentWasTransactionNumber = .Fields.Item(2).IntegerValue
          End If
          .CloseRecordSet()
        End With

        'Find the transaction that changed this financial history  
        With vWhereFields
          .Clear()
          .Add("was_batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("was_transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End With
        vRS = mvEnv.Connection.GetRecordSet("SELECT batch_number, transaction_number FROM reversals WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        With vRS
          If .Fetch() = True Then
            If mvAdjustmentState = FinancialHistory.AdjustmentStates.adjsIsAnAdjustment Then
              mvAdjustmentState = FinancialHistory.AdjustmentStates.adjsIsAnAdjustmentAndHasBeenAdjusted
            Else
              mvAdjustmentState = FinancialHistory.AdjustmentStates.adjsHasBeenAdjusted
            End If
            mvAdjustmentBatchNumber = .Fields.Item(1).IntegerValue
            mvAdjustmentTransactionNumber = .Fields.Item(2).IntegerValue
          End If
          .CloseRecordSet()
        End With

        mvAdjustmentStateSet = True
      End If
      AdjustmentState = mvAdjustmentState
    End Function

    Public Function PreProcessAdjustmentChecks(ByVal pAmount As Double, ByVal pLineNumber As Integer, ByVal pAllocationsChecked As Boolean, ByVal pFullAmountAllocation As Boolean, ByVal pAdjustmentType As Batch.AdjustmentTypes) As Boolean
      'Perform checks to see whether the Financial Adjustment can take place
      'Called by CheckAdjustmentAllowed WebService and Financial Adjustments
      Dim vPP As New PaymentPlan
      Dim vCanAdjust As Boolean
      Dim vCount As Integer
      Dim vMsg As String = ""
      Dim vSQL As String

      'If just reversing an analysis line check if part of an order payment
      'SDT 28/3/2002  Don't check unless the financial history exists (batch may not be posted yet)
      If pLineNumber > 0 And mvExisting Then
        vSQL = "oph.batch_number = " & BatchNumber & " AND oph.transaction_number = " & TransactionNumber & " AND oph.line_number = " & pLineNumber
        vSQL = vSQL & " AND fhd.batch_number = oph.batch_number AND fhd.transaction_number = oph.transaction_number AND fhd.line_number = oph.line_number"
        vCount = mvEnv.Connection.GetCount("order_payment_history oph, financial_history_details fhd", Nothing, vSQL)
        vSQL = "batch_number = " & BatchNumber & " AND transaction_number = " & TransactionNumber
        If vCount = mvEnv.Connection.GetCount("financial_history_details", Nothing, vSQL) Then
          RaiseError(DataAccessErrors.daeCannotReverseOrderAnalysis)
        End If
      End If

      vPP.Init(mvEnv)
      If vPP.ContainsUnprocessedPayments(True, BatchNumber, TransactionNumber, pLineNumber) = True Then
        'This is a Payment Plan payment with a zero balance and unprocessed payments
        RaiseError(DataAccessErrors.daeCannotAdjustZeroBalancePP)
      End If

      vCanAdjust = CheckFAInAdvance(pLineNumber, vMsg, True)
      If vCanAdjust = False And Len(vMsg) > 0 Then RaiseError(DataAccessErrors.daeAdjustmentError, vMsg)

      If vCanAdjust Then
        Dim vSLItems As SalesLedgerItems = ContainsSalesLedgerItems(mvEnv, BatchNumber, TransactionNumber, pLineNumber)
        If (vSLItems.HasFlag(SalesLedgerItems.SLCashAllocation) OrElse vSLItems.HasFlag(SalesLedgerItems.CreditNoteAllocation)) Then vCanAdjust = False
        If vCanAdjust = True AndAlso vSLItems > SalesLedgerItems.None AndAlso IsFinancialAdjustment = True Then
          Select Case pAdjustmentType
            Case Batch.AdjustmentTypes.atReverse, Batch.AdjustmentTypes.atRefund
              'Can adjust if result of a re-analysis
              If CanReverseAdjustedTransaction(vSLItems) = False Then vCanAdjust = False
            Case Batch.AdjustmentTypes.atPartRefund
              'Can adjust if result of a re-analysis and line number > 0
              If Not (pLineNumber > 0 AndAlso CanReverseAdjustedTransaction(vSLItems) = True) Then vCanAdjust = False
            Case Else
              'Cannot adjust
              vCanAdjust = False
          End Select
        End If
        If vCanAdjust = False Then RaiseError(DataAccessErrors.daeCannotAdjustSLAllocation)
      End If

      If vCanAdjust = True And pAllocationsChecked = False Then
        'Check to see if the transaction or line being adjusted forms part of a transaction used to allocate SL cash against an invoice.
        'If it does then do not allow the transaction / line to tbe adjusted 
        If vCanAdjust = False Then RaiseError(DataAccessErrors.daeCannotAdjustSLAllocation)
        If vCanAdjust = True Then CheckAllocations(BatchNumber, TransactionNumber, pAmount, pFullAmountAllocation, (pAdjustmentType = Batch.AdjustmentTypes.atEventAdjustment))
      End If
      Return vCanAdjust
    End Function

    Private Sub DeleteFromBatch(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      DeleteFromBatch(pBatchNumber, pTransactionNumber, 0, 0, False, False)
    End Sub
    Private Sub DeleteFromBatch(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pOtherBatchNumber As Integer, ByVal pOtherTransactionNumber As Integer, ByVal pMoveTrans As Boolean, ByVal pAdjPositiveLinesOnly As Boolean)
      Dim vBatch As Batch
      Dim vOtherBTCount As Integer
      Dim vWhereFields As CDBFields
      Dim vMultipleBatches As Boolean

      vMultipleBatches = pBatchNumber <> pOtherBatchNumber And pOtherBatchNumber > 0

      vWhereFields = New CDBFields
      With vWhereFields
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
        If Not vMultipleBatches And pOtherTransactionNumber > 0 Then
          .Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber & "," & pOtherTransactionNumber, CDBField.FieldWhereOperators.fwoNotIn)
        Else
          .Add("transaction_number", pTransactionNumber, CDBField.FieldWhereOperators.fwoNotEqual)
        End If
      End With
      vOtherBTCount = mvEnv.Connection.GetCount("batch_transactions", vWhereFields, "")
      vBatch = New Batch(mvEnv)
      vBatch.Init(pBatchNumber)
      vBatch.Delete(pTransactionNumber, False)
      If Not vMultipleBatches And pOtherTransactionNumber > 0 Then vBatch.Delete(pOtherTransactionNumber, False)
      If vOtherBTCount = 0 Then vBatch.Delete(0, False)

      If vMultipleBatches Then
        vWhereFields = New CDBFields
        With vWhereFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, pOtherBatchNumber)
          .Add("transaction_number", pOtherTransactionNumber, CDBField.FieldWhereOperators.fwoNotEqual)
        End With
        vOtherBTCount = mvEnv.Connection.GetCount("batch_transactions", vWhereFields, "")
        vBatch = New Batch(mvEnv)
        vBatch.Init(pOtherBatchNumber)
        vBatch.Delete(pOtherTransactionNumber, False)
        If vOtherBTCount = 0 Then vBatch.Delete(0, False)
      End If

      If pMoveTrans Then
        'If original batch contains invoice payments then update Invoice being paid and CreditCustomers record
        UpdateInvoiceForMove(BatchNumber, TransactionNumber, pAdjPositiveLinesOnly, True)
      End If

    End Sub

    Public Sub ReanalyseTransaction(ByVal pNewBatchNumber As Integer, ByVal pNewTransactionNumber As Integer, ByVal pParams As CDBParameters, ByVal pOrigBatch As Batch, Optional ByVal pSelectedTransactions As CDBCollection = Nothing, Optional ByVal pLinkedAnalysis As CollectionList(Of BatchTransactionAnalysis) = Nothing)
      Dim vDLU As New DeclarationLinesUnclaimed(mvEnv)
      Dim vGSLU As New GaSponsorshipLinesUnclaimed
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vNextLineNumber As Integer
      Dim vAmount As Double
      Dim vCurrencyAmount As Double
      Dim vLineNumber As Integer
      Dim vFABTA As CollectionList(Of BatchTransactionAnalysis)
      Dim vOrigBTA As CollectionList(Of BatchTransactionAnalysis)
      Dim vFABTALine As New BatchTransactionAnalysis(mvEnv)
      Dim vOrigBTALine As New BatchTransactionAnalysis(mvEnv)
      Dim vBatchTransaction As New BatchTransaction(mvEnv)
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vLineRemoved As Boolean
      Dim vRemovedLines As String = ""
      Dim vLineIncrement As Integer
      Dim vCC As New CompanyControl
      Dim vRS As CDBRecordSet
      Dim vRemove As Boolean
      Dim vOrigSchNo As Integer
      Dim vNewSchNo As Integer
      Dim vOrigOPH As New OrderPaymentHistory
      Dim vOrigOPS As New OrderPaymentSchedule
      Dim vPP As PaymentPlan
      Dim vPrevLineNumber As Integer
      Dim vCP As CollectionPayment
      Dim vPIS As CollectionPIS
      Dim vSQL As String
      Dim vContinue As Boolean
      Dim vMultipleTransactions As Boolean
      Dim vWasBatchNumber As Integer
      Dim vWasTransactionNumber As Integer
      Dim vBookingExists As Boolean
      Dim vBookingNo As Integer
      Dim vSB As New ServiceBooking
      Dim vSBLinkedAnalysis As New Collection
      Dim vServiceBookingTrans As Boolean

      Try
        If Not (pSelectedTransactions Is Nothing) Then
          If pSelectedTransactions.Count > 1 Then vMultipleTransactions = True
        End If
        vContinue = True
        If vContinue Then
          'create two BTA collections - one for the new FA transaction and one for the original transaction
          vBatchTransaction.InitForUpdate(pNewBatchNumber, pNewTransactionNumber, True)
          vBatchTransaction.InitBatchTransactionAnalysis(pNewBatchNumber, pNewTransactionNumber)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then vBatchTransaction.InitAnalysisAdditionalData()
          vFABTA = vBatchTransaction.Analysis

          'Get original BTA
          If Not (pSelectedTransactions Is Nothing) Then
            'We have multiple transactions to be selected
            vOrigBTA = New CollectionList(Of BatchTransactionAnalysis)
            For Each vBT In pSelectedTransactions
              vBT.InitDetailsFromFinancialHistory(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
              vBT.InitAnalysisAdditionalData()
              For Each vOrigBTALine In vBT.Analysis
                vOrigBTA.Add(vOrigBTALine.Key(True), vOrigBTALine)
              Next vOrigBTALine
            Next vBT
          Else
            vBT = New BatchTransaction(mvEnv)
            vBT.InitDetailsFromFinancialHistory(mvEnv, BatchNumber, TransactionNumber)
            vBT.InitAnalysisAdditionalData()
            vOrigBTA = vBT.Analysis
          End If

          For Each vOrigBTALine In vOrigBTA
            If vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBookingTransaction Then
              vOrigBTALine.LinkedBookingNo = vOrigBTALine.AdditionalNumber2
            End If
          Next vOrigBTALine

          vLineIncrement = vOrigBTA.Count()
          'for each BTA in the FA compare it against every BTA in the original transaction
          'if all values, or at least the important ones, match then flag that BTA to be removed from the FA transaction
          Dim vSameLine As Boolean = False
          Dim vLinkedToFundPayment As Boolean
          Dim vTempFABTA As New CollectionList(Of BatchTransactionAnalysis)
          For Each vFABTALine In vFABTA
            For Each vOrigBTALine In vOrigBTA
              vSameLine = False
              If vFABTALine.LineType = vOrigBTALine.LineType AndAlso vFABTALine.ProductCode = vOrigBTALine.ProductCode AndAlso vFABTALine.RateCode = vOrigBTALine.RateCode _
              AndAlso vFABTALine.DistributionCode = vOrigBTALine.DistributionCode AndAlso vFABTALine.Quantity = vOrigBTALine.Quantity AndAlso vFABTALine.Issued = vOrigBTALine.Issued _
              AndAlso vFABTALine.Amount = vOrigBTALine.Amount AndAlso vFABTALine.MemberNumber = vOrigBTALine.MemberNumber AndAlso vFABTALine.CovenantNumber = vOrigBTALine.CovenantNumber _
              AndAlso vFABTALine.Source = vOrigBTALine.Source AndAlso vFABTALine.SalesContactNumber = vOrigBTALine.SalesContactNumber Then
                'Basic details are the same
                If vOrigBTALine.LineType = "N" Then
                  'For an invoice payment, compare the InvoiceNumber of the invoice being paid
                  If vOrigBTALine.InvoiceNumber = vFABTALine.InvoiceNumber Then vSameLine = True
                ElseIf vFABTALine.PaymentPlanNumber = vOrigBTALine.PaymentPlanNumber Then
                  'Otherwise compare the PaymentPlanNumber
                  vSameLine = True
                ElseIf vOrigBTALine.LineType = "U" Then   'BR20381
                  If vOrigBTALine.MemberNumber = vFABTALine.MemberNumber Then vSameLine = True
                End If
              End If
              If vSameLine Then
                vRemove = True
                If vOrigBTALine.Product.StockItem And mvEnv.GetConfigOption("fp_stock_multiple_warehouses") And Len(vFABTALine.Warehouse) > 0 And vFABTALine.Warehouse <> vOrigBTALine.Warehouse Then
                  'The Warehouse property of the vFABTALine object will be null if nothing on the line had been changed during the reanalysis
                  'The warehouse has changed so treat as a different line
                  vRemove = False
                End If
                If vRemove = True And vOrigBTALine.LineType = "B" Then
                  vWhereFields = New CDBFields
                  vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, vFABTALine.BatchNumber)
                  vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, vFABTALine.TransactionNumber)
                  vWhereFields.Add("line_number", CDBField.FieldTypes.cftInteger, vFABTALine.LineNumber)
                  If mvEnv.Connection.GetCount("legacy_bequest_receipts", vWhereFields) > 0 Then
                    'We made the decision to create a Receipt in Trader SaveTransaction.
                    'Probably because different Legacy Bequest with same Product.
                    'Hence new Legacy Bequest payment is required.
                    vRemove = False
                  End If
                End If
                If vOrigBTALine.PaymentPlanNumber > 0 Then
                  'Could be re-analysing an incorrectly posted payment plan posting, so check fhd product/rate
                  '.. if match CompanyControls product/rate then do not remove the line
                  'Looks like no need to check over-payment product
                  If vCC.Existing = False Then vCC.InitFromBankAccount(mvEnv, (pOrigBatch.BankAccount))
                  vRS = mvEnv.Connection.GetRecordSet("SELECT product,rate FROM financial_history_details WHERE batch_number = " & vOrigBTALine.BatchNumber & " AND transaction_number = " & vOrigBTALine.TransactionNumber & " AND line_number = " & vOrigBTALine.LineNumber)
                  While vRS.Fetch() = True And vRemove = True
                    If vRS.Fields("product").Value = vCC.DetailsProductCode And vRS.Fields("rate").Value = vCC.DetailsRate Then
                      vRemove = False
                    ElseIf vRS.Fields("product").Value = vCC.LockedProductCode And vRS.Fields("rate").Value = vCC.LockedRate Then
                      vRemove = False
                    End If
                  End While
                  vRS.CloseRecordSet()
                End If

                If vRemove And vOrigBTALine.PaymentPlanNumber > 0 Then
                  'Check the scheduled payment being paid - could have changed
                  'First the original one
                  vWhereFields = New CDBFields
                  With vWhereFields
                    .Add("order_number", CDBField.FieldTypes.cftLong, vOrigBTALine.PaymentPlanNumber)
                    .Add("batch_number", CDBField.FieldTypes.cftLong, vOrigBTALine.BatchNumber)
                    .Add("transaction_number", CDBField.FieldTypes.cftLong, vOrigBTALine.TransactionNumber)
                    .Add("line_number", CDBField.FieldTypes.cftLong, vOrigBTALine.LineNumber)
                    .Add("amount", CDBField.FieldTypes.cftLong, vOrigBTALine.Amount)
                    .Add("posted", CDBField.FieldTypes.cftCharacter, "Y")
                  End With
                  vOrigSchNo = IntegerValue(mvEnv.Connection.GetValue("SELECT scheduled_payment_number FROM order_payment_history WHERE " & mvEnv.Connection.WhereClause(vWhereFields)))
                  If vOrigSchNo > 0 Then
                    'Now the new payment
                    With vWhereFields
                      .Item("batch_number").Value = CStr(vFABTALine.BatchNumber)
                      .Item("transaction_number").Value = CStr(vFABTALine.TransactionNumber)
                      .Item("line_number").Value = CStr(vFABTALine.LineNumber)
                      .Item("posted").Value = "N"
                    End With
                    vNewSchNo = IntegerValue(mvEnv.Connection.GetValue("SELECT scheduled_payment_number FROM order_payment_history WHERE " & mvEnv.Connection.WhereClause(vWhereFields)))
                    If vOrigSchNo <> vNewSchNo Then vRemove = False 'Same Pay Plan but different scheduled payment
                  End If
                End If

                If vRemove = True Then
                  If vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatCollectionPayments Then
                    'Original was a CollectionPayment
                    'AdditionalNumber = CollectionNumber
                    'AdditionalNumber2 = CollectionPISNumber
                    'MemberNumber = CollectionBoxNumbers (comma-separated list)
                    If vFABTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatCollectionPayments And (vOrigBTALine.AdditionalNumber = vFABTALine.AdditionalNumber) And (vOrigBTALine.AdditionalNumber2 = vFABTALine.AdditionalNumber2) Then
                      'Paying the same Collection and same PIS Number so OK to remove
                    Else
                      'Paying different Collection/PIS Number or not paying a Collection
                      vRemove = False
                    End If
                  ElseIf vFABTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatCollectionPayments Then
                    'Original not a CollectionPayment but the new line is so do not remove
                    vRemove = False
                  ElseIf vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBooking Then
                    If Not pLinkedAnalysis Is Nothing Then
                      vWhereFields = New CDBFields
                      vWhereFields.Add("booking_number", CDBField.FieldTypes.cftLong, vOrigBTALine.AdditionalNumber)
                      If mvEnv.Connection.GetCount("event_booking_transactions", vWhereFields) > 0 Then
                        vBookingExists = True
                        vBookingNo = vOrigBTALine.AdditionalNumber
                      End If
                    End If
                  ElseIf vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment Then
                    vLinkedToFundPayment = True
                  End If
                End If

                If vRemove And vFABTALine.LineType Like "[GHS]" Then vRemove = vFABTALine.DeceasedContactNumber = vOrigBTALine.DeceasedContactNumber

                If vRemove Then
                  vOrigBTA.Remove(vOrigBTALine.Key(vMultipleTransactions))
                  vLineRemoved = True
                  Exit For
                End If
              End If
            Next
            If vLineRemoved Then
              If vRemovedLines.Length > 0 Then vRemovedLines = vRemovedLines & ", "
              vRemovedLines = vRemovedLines & CStr(vFABTALine.LineNumber)
              vAmount = vAmount + (vFABTALine.Amount * -1)
              If vFABTALine.CurrencyAmount <> 0 Then vCurrencyAmount = vCurrencyAmount + (vFABTALine.CurrencyAmount * -1)
              vTempFABTA.Add(vFABTALine.Key, vFABTALine)
              vLineRemoved = False
              Dim vCompany As String  'BR20381
              vCompany = mvEnv.Connection.GetValue("SELECT company FROM invoices WHERE batch_number = " & vOrigBTALine.BatchNumber & " AND transaction_number = " & vOrigBTALine.TransactionNumber)
              If vOrigBTALine.LineType = "U" AndAlso Not vOrigBTALine.Existing AndAlso Not String.IsNullOrEmpty(vCompany) Then
                Dim vCCU As New CreditCustomer()
                vCCU.InitCompanySalesLedgerAccount(mvEnv, vCompany, vOrigBTALine.MemberNumber)
                If vCCU.Existing Then
                  vCCU.AdjustOutstanding(vOrigBTALine.Amount)
                  vCCU.Save(mvEnv.User.UserID)
                End If
              End If
            Else
              If vBookingExists Then
                vNextLineNumber = vFABTALine.LineNumber
                vFABTALine.LinkedBookingNo = vBookingNo
                pLinkedAnalysis.Add(vFABTALine.Key, vFABTALine)
                vFABTALine.LineNumber = vFABTALine.LineNumber + vLineIncrement
              Else
                vNextLineNumber = vFABTALine.LineNumber
              End If
            End If
          Next
          For Each vFABTALine In vTempFABTA
            vFABTA.Remove(vFABTALine.Key)
          Next
          If vOrigBTA.Count > 0 Then
            mvEnv.Connection.StartTransaction()
            '          vNextLineNumber = vNextLineNumber + 1
            'remove any flagged BTA records from the FA trans'n
            If vRemovedLines.Length > 0 Then
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNumber)
              vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, pNewTransactionNumber)
              vWhereFields.Add("line_number", CDBField.FieldTypes.cftCharacter, vRemovedLines, CDBField.FieldWhereOperators.fwoIn)
              mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields)
              mvEnv.Connection.DeleteRecords("order_payment_history", vWhereFields, False)
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then mvEnv.Connection.DeleteRecords("collection_payments", vWhereFields, False)
              'BR13623: Remove New FA lines links (adde by TraderApp.SaveTransaction)
              'Do not raise error as the new FA may not be linked
              If vLinkedToFundPayment Then mvEnv.Connection.DeleteRecords("fundraising_payment_history", vWhereFields, False)
            End If

            'Next, move the +ve lines forward so the -ve lines are before the +ve lines
            vUpdateFields = New CDBFields
            vWhereFields = New CDBFields
            vUpdateFields.Add("line_number", CDBField.FieldTypes.cftLong, "line_number + " & vLineIncrement)
            With vWhereFields
              .Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNumber)
              .Add("transaction_number", CDBField.FieldTypes.cftLong, pNewTransactionNumber)
            End With
            mvEnv.Connection.UpdateRecords("batch_transaction_analysis", vUpdateFields, vWhereFields)
            mvEnv.Connection.UpdateRecords("stock_movements", vUpdateFields, vWhereFields, False)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then mvEnv.Connection.UpdateRecords("order_payment_history", vUpdateFields, vWhereFields, False)
            mvEnv.Connection.UpdateRecords("legacy_bequest_receipts", vUpdateFields, vWhereFields, False)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then mvEnv.Connection.UpdateRecords("collection_payments", vUpdateFields, vWhereFields, False)
            'BR13623: Update any linked records
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then mvEnv.Connection.UpdateRecords("fundraising_payment_history", vUpdateFields, vWhereFields, False)
            'BR13302/J1224: Update any service booking transaction records
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then mvEnv.Connection.UpdateRecords("service_booking_transactions", vUpdateFields, vWhereFields, False)

            'Add a -ve duplicate of the remaining original BTA to the FA trans'n (line number will start with 1)
            vNextLineNumber = 1
            vLineNumber = 0
            vPrevLineNumber = 0
            Dim vOriginalLineAmount As Double = 0
            For Each vOrigBTALine In vOrigBTA
              With vOrigBTALine
                vWasBatchNumber = .BatchNumber
                vWasTransactionNumber = .TransactionNumber
                vLineNumber = .LineNumber
                vOriginalLineAmount = .CurrencyAmount
                .CloneForFA(pNewBatchNumber, pNewTransactionNumber, vNextLineNumber)
                vAmount = FixTwoPlaces(vAmount + .Amount)
                If .CurrencyAmount <> 0 Then vCurrencyAmount = FixTwoPlaces(vCurrencyAmount + .CurrencyAmount)
              End With
              vOrigBTALine.Save()

              'create Reversals record
              vOrigOPH.Init(mvEnv)
              If vLineNumber <> vPrevLineNumber Then
                vUpdateFields = New CDBFields
                With vUpdateFields
                  .Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNumber)
                  .Add("transaction_number", CDBField.FieldTypes.cftInteger, pNewTransactionNumber)
                  .Add("line_number", CDBField.FieldTypes.cftInteger, vNextLineNumber)
                  .Add("was_batch_number", CDBField.FieldTypes.cftLong, vWasBatchNumber)
                  .Add("was_transaction_number", CDBField.FieldTypes.cftInteger, vWasTransactionNumber)
                  .Add("was_line_number", CDBField.FieldTypes.cftInteger, vLineNumber)
                  If (Len(vOrigOPH.Status) > 0 And vOrigOPH.Status = "I") Then .Add("was_oph_status", CDBField.FieldTypes.cftCharacter, vOrigOPH.Status)
                  mvEnv.Connection.InsertRecord("reversals", vUpdateFields)
                End With
              End If

              If vBookingExists And (vOrigBTALine.LinkedBookingNo > 0) Then pLinkedAnalysis.Add(vOrigBTALine.Key, vOrigBTALine)

              If vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBookingTransaction AndAlso (vOrigBTALine.AdditionalNumber > 0) Then
                vSB = New ServiceBooking
                With vOrigBTALine
                  vSB.SetTransactionInfo(mvEnv, .AdditionalNumber, .BatchNumber, .TransactionNumber, .LineNumber, .SalesContactNumber)
                  .LinkedBookingNo = .AdditionalNumber
                End With
                vSBLinkedAnalysis.Add(vOrigBTALine)
                vServiceBookingTrans = True
              End If

              'BR13623: Add reversed analysis lines to the linked table - Dont check the flag on TraderApp
              If vLineNumber <> vPrevLineNumber And vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment And vOrigBTALine.AdditionalNumber > 0 Then
                Dim vFPH As New FundraisingPaymentHistory(mvEnv)
                vFPH.Init()
                vFPH.CreateNewLink(vOrigBTALine.AdditionalNumber, pNewBatchNumber, pNewTransactionNumber, vNextLineNumber)
              End If

              'Create OrderPaymentHistory
              If vOrigBTALine.PaymentPlanNumber > 0 AndAlso vOrigBTALine.LineType Like "[OMC]" Then
                vRS = mvEnv.Connection.GetRecordSet("SELECT " & vOrigOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE batch_number = " & vWasBatchNumber & " AND transaction_number = " & vWasTransactionNumber & " AND line_number = " & vLineNumber)
                If vRS.Fetch() = True Then vOrigOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
                vRS.CloseRecordSet()

                vPP = New PaymentPlan
                vPP.Init(mvEnv, (vOrigOPH.OrderNumber))
                vPP.PaymentNumber = vPP.PaymentNumber + 1
                vPP.SaveChanges()

                If vOrigOPH.ScheduledPaymentNumber.Length > 0 Then
                  vOrigSchNo = IntegerValue(vOrigOPH.ScheduledPaymentNumber)
                Else
                  'This is a pre-v5.x payment being re-analysed so need to find an OPS to allocate the payment against
                  vOrigOPS = New OrderPaymentSchedule
                  vOrigOPS.Init(mvEnv)
                  vOrigSchNo = vOrigOPS.ReverseHistoricPayment(vPP, vOrigOPH.Amount)
                End If

                If vOrigOPH.ScheduledPaymentNumber.Length > 0 Then
                  'Before reversing the original ops, see if the new payment has been allocated against it
                  vWhereFields = New CDBFields
                  With vWhereFields
                    .Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNumber)
                    .Add("transaction_number", CDBField.FieldTypes.cftLong, pNewTransactionNumber)
                    .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong, vOrigOPH.ScheduledPaymentNumber)
                    .Add("amount", 0, CDBField.FieldWhereOperators.fwoGreaterThan)
                  End With
                  If mvEnv.Connection.GetCount("order_payment_history", vWhereFields) = 0 Then
                    'By now Trader has already updated the AmountOutstanding, so just need to set as an unprocessed payment
                    vOrigOPS.Init(mvEnv, CInt(vOrigOPH.ScheduledPaymentNumber))
                    vOrigOPS.SetUnProcessedPayment(True, vOrigOPH.Amount)
                    vOrigOPS.Reverse(vPP, vOrigOPH.Amount)
                    vOrigOPS.Save()
                  End If
                End If
                vOrigOPH.Reverse(pNewBatchNumber, pNewTransactionNumber, vNextLineNumber, vPP.PaymentNumber, vOrigSchNo)
                vOrigOPH.Save()
              End If
              'update FHD & OPH status
              vUpdateFields = New CDBFields
              vWhereFields = New CDBFields
              With vWhereFields
                .Add("batch_number", CDBField.FieldTypes.cftLong, vWasBatchNumber)
                .Add("transaction_number", CDBField.FieldTypes.cftInteger, vWasTransactionNumber)
                .Add("line_number", CDBField.FieldTypes.cftInteger, vLineNumber)
              End With
              vUpdateFields.Add("status", CDBField.FieldTypes.cftCharacter, FA_STATUS_ADJUSTMENT)
              mvEnv.Connection.UpdateRecords("financial_history_details", vUpdateFields, vWhereFields)
              mvEnv.Connection.UpdateRecords("order_payment_history", vUpdateFields, vWhereFields, False)
              'deal with any gift aid declaration or covenant payments already claimed
              vDLU.Init()
              vDLU.CreateNewNegativeLines(pNewBatchNumber, pNewTransactionNumber, vNextLineNumber, vWasBatchNumber, vWasTransactionNumber, vLineNumber)

              'Deal with GA Sponsorship payments
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidSponsorship) Then
                mvEnv.Connection.DeleteRecords("ga_sponsorship_lines_unclaimed", vWhereFields, False)
                vGSLU.Init(mvEnv)
                vGSLU.CreateNewNegativeLines(pNewBatchNumber, pNewTransactionNumber, vNextLineNumber, vWasBatchNumber, vWasTransactionNumber, vLineNumber)
              End If

              'Deal with CollectionPayments
              If vOrigBTALine.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatCollectionPayments Then
                vCP = New CollectionPayment
                vCP.Init(mvEnv)
                vSQL = "SELECT " & vCP.GetRecordSetFields(CollectionPayment.CollectionPaymentRecordSetTypes.cpyrtAll) & " FROM collection_payments cp"
                vSQL = vSQL & " WHERE batch_number = " & vWasBatchNumber & " AND transaction_number = "
                vSQL = vSQL & vWasTransactionNumber & " AND line_number = " & vLineNumber
                vRS = mvEnv.Connection.GetRecordSet(vSQL)
                While vRS.Fetch() = True
                  vCP = New CollectionPayment
                  vCP.InitFromRecordSet(mvEnv, vRS, CollectionPayment.CollectionPaymentRecordSetTypes.cpyrtAll)
                  vCP.Reverse(pNewBatchNumber, pNewTransactionNumber, vNextLineNumber)
                  vCP.Save()
                  If vCP.CollectionPisNumber > 0 Then
                    vPIS = New CollectionPIS
                    vPIS.Init(mvEnv, (vCP.CollectionPisNumber))
                    vPIS.Reconcile(CollectionPIS.CollectionPISReconciledStatus.cpisrsReversed)
                    vPIS.Save()
                  End If
                End While
                vRS.CloseRecordSet()
              End If

              'Update IPH & Invoice
              Dim vInvoice As Invoice
              Dim vOrigIPH As New InvoicePaymentHistory(mvEnv)
              vOrigIPH.InitFromBatchTransactionLine(vWasBatchNumber, vWasTransactionNumber, vLineNumber)
              If vOrigIPH.Existing = True AndAlso vOrigIPH.Status.Length = 0 Then
                If vOrigIPH.BatchNumber.Equals(vWasBatchNumber) AndAlso vOrigIPH.TransactionNumber.Equals(vWasTransactionNumber) Then
                  If (vOrigBTALine.LineType.Equals("U", StringComparison.InvariantCultureIgnoreCase) _
                  AndAlso (vOrigIPH.AllocationBatchNumber.Equals(vWasBatchNumber) = False OrElse vOrigIPH.AllocationTransactionNumber.Equals(vWasTransactionNumber) = False)) Then
                    vInvoice = New Invoice()
                    vInvoice.Init(mvEnv, 0, 0, vOrigIPH.InvoiceNumber)
                    If vInvoice.Existing Then
                      vInvoice.SetAmountPaid((vOrigIPH.Amount * -1))
                      vInvoice.Save(mvEnv.User.UserID, True)
                    End If
                  ElseIf (vOrigBTALine.LineType.Equals("N", StringComparison.InvariantCultureIgnoreCase) _
                      AndAlso (vOrigIPH.AllocationBatchNumber.Equals(vWasBatchNumber) AndAlso vOrigIPH.AllocationTransactionNumber.Equals(vWasTransactionNumber))) Then
                    Invoice.RemoveInvoiceAllocations(mvEnv, vBT, vOrigBTALine, vInvoice, pOrigBatch.BatchType, "", Batch.AdjustmentTypes.atAdjustment)
                  End If
                End If
                vOrigIPH.Reverse(FA_STATUS_REVERSAL, pNewBatchNumber, pNewTransactionNumber, vNextLineNumber, CDate(pParams("TransactionDate").Value))
              ElseIf vOrigIPH.Existing = False AndAlso vOrigBTALine.LineType.Equals("U") Then
                'Note: vOrigBTALine is not actually the original line - it's become the new line!!!
                Dim vTransDate As Nullable(Of Date) = Nothing
                If pParams.Exists("TransactionDate") Then vTransDate = CDate(pParams("TransactionDate").Value)
                If vTransDate.HasValue = False AndAlso pParams.Exists("TRD_TransactionDate") Then vTransDate = CDate(pParams("TRD_Transaction_Date").Value)
                If vTransDate.HasValue = False Then vTransDate = Today
                ReverseInvoiceCashAllocation(vLineNumber, vOrigBTALine.LineType, vOriginalLineAmount, vOrigBTALine, vTransDate.Value, FA_STATUS_REVERSAL)
              End If
              vInvoice = New Invoice()
              vInvoice.Init(mvEnv, vWasBatchNumber, vWasTransactionNumber)
              If vInvoice.Existing = True AndAlso vInvoice.RecordType = Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash) Then
                vInvoice.SetAdjustmentStatus(Invoice.InvoiceAdjustmentStatus.Adjusted)
                vInvoice.Save(mvEnv.User.UserID, True)
              End If

              vNextLineNumber = vNextLineNumber + 1
              vPrevLineNumber = vLineNumber
            Next vOrigBTALine

            'Add any Service Booking Transaction lines added/ adjusted
            If vServiceBookingTrans Then vSB.AddLinkedTransaction(Batch.AdjustmentTypes.atAdjustment, vSBLinkedAnalysis)

            'We now need to update all the original FH etc. for each Batch/Transaction selected
            'Update original FH
            vUpdateFields = New CDBFields
            vUpdateFields.Add("status", CDBField.FieldTypes.cftCharacter, FA_STATUS_ADJUSTMENT)
            vWhereFields = New CDBFields
            With vWhereFields
              .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
              .Add("transaction_number", CDBField.FieldTypes.cftInteger, TransactionNumber)
            End With
            mvEnv.Connection.UpdateRecords("financial_history", vUpdateFields, vWhereFields)
            'Update original back_order_details
            With vUpdateFields
              .Clear()
              .Add("status", CDBField.FieldTypes.cftCharacter, FA_STATUS_REVERSAL)
            End With
            mvEnv.Connection.UpdateRecords("back_order_details", vUpdateFields, vWhereFields, False)

            If vMultipleTransactions Then
              'Update remaining FH & BOD
              vBatchTransaction = Nothing
              For Each vBatchTransaction In pSelectedTransactions
                If Not (vBatchTransaction.BatchNumber = BatchNumber And vBatchTransaction.TransactionNumber = TransactionNumber) Then
                  vWhereFields(1).Value = CStr(vBatchTransaction.BatchNumber)
                  vWhereFields(2).Value = CStr(vBatchTransaction.TransactionNumber)
                  'Update original FH
                  vUpdateFields(1).Value = FA_STATUS_ADJUSTMENT
                  mvEnv.Connection.UpdateRecords("financial_history", vUpdateFields, vWhereFields)
                  'Update original back_order_details
                  vUpdateFields(1).Value = FA_STATUS_REVERSAL
                  mvEnv.Connection.UpdateRecords("back_order_details", vUpdateFields, vWhereFields, False)
                End If
              Next vBatchTransaction
            End If

            'Update new FA batch_transactions
            vWhereFields = New CDBFields
            With vWhereFields
              .Add("batch_number", CDBField.FieldTypes.cftLong, pNewBatchNumber)
              .Add("transaction_number", CDBField.FieldTypes.cftInteger, pNewTransactionNumber)
            End With
            vUpdateFields = New CDBFields
            With vUpdateFields
              .Add("amount", CDBField.FieldTypes.cftNumeric, "amount + " & CStr(vAmount))
              .Add("line_total", CDBField.FieldTypes.cftNumeric, "line_total + " & CStr(vAmount))
              .Add("next_line_number", CDBField.FieldTypes.cftInteger, vNextLineNumber + vLineIncrement)
              .Add("transaction_date", CDBField.FieldTypes.cftDate, pParams("TransactionDate").Value)
              .Add("transaction_type", CDBField.FieldTypes.cftCharacter, pParams("FATransactionType").Value)
              If pParams.Exists("Notes") Then .Add("notes", CDBField.FieldTypes.cftCharacter, pParams("Notes").Value)
              If vCurrencyAmount.ToString.Length > 0 Then
                .Add("currency_amount", CDBField.FieldTypes.cftNumeric, "currency_amount + " & CStr(vCurrencyAmount))
              End If
              'Clear the Mailing, MailingContactNumber, MailingAddressNumber
              .Add("mailing")
              .Add("mailing_contact_number")
              .Add("mailing_address_number")
            End With
            mvEnv.Connection.UpdateRecords("batch_transactions", vUpdateFields, vWhereFields)
            'Update new FA batch
            vWhereFields.Remove((2))
            vUpdateFields = New CDBFields
            With vUpdateFields
              .Add("batch_total", CDBField.FieldTypes.cftNumeric, "batch_total + " & CStr(vAmount))
              .Add("transaction_total", CDBField.FieldTypes.cftNumeric, "transaction_total + " & CStr(vAmount))
              If vCurrencyAmount.ToString.Length > 0 Then
                .Add("currency_batch_total", CDBField.FieldTypes.cftNumeric, "currency_batch_total + " & CStr(vCurrencyAmount))
                .Add("currency_transaction_total", CDBField.FieldTypes.cftNumeric, "currency_transaction_total + " & CStr(vCurrencyAmount))
              End If
            End With
            mvEnv.Connection.UpdateRecords("batches", vUpdateFields, vWhereFields)
            mvEnv.Connection.CommitTransaction()
            'MsgBox (ProjectText.String29517), vbInformation   'Transaction has been Re-Analysed
          Else
            'just in case the resulting FA trans'n contains exactly the same analysis as the original trans'n
            DeleteFromBatch(pNewBatchNumber, pNewTransactionNumber)
          End If
        End If
      Catch vEx As Exception
        mvEnv.Connection.RollbackTransaction()
        If pNewBatchNumber > 0 Then DeleteFromBatch(pNewBatchNumber, pNewTransactionNumber)
        PreserveStackTrace(vEx)
      End Try
    End Sub

    Public Function GetMultipleTransactions(ByVal pList As String) As CDBCollection
      Dim vList() As String
      Dim vBatchNumber As Integer
      Dim vTransNumber As Integer
      Dim vBT As BatchTransaction
      Dim vColl As New CDBCollection
      Dim vIndex As Integer

      vList = Split(pList, ",")
      For vIndex = 0 To UBound(vList)
        vBatchNumber = CInt(Split(vList(vIndex), "|")(0)) '�
        vTransNumber = CInt(Split(vList(vIndex), "|")(1))
        vBT = New BatchTransaction(mvEnv)
        vBT.Init(vBatchNumber, vTransNumber)
        vColl.Add(vBT)
      Next
      GetMultipleTransactions = vColl
    End Function


    Public Function GetBatchTotal(ByVal pAdjBatchNumber As Integer) As ParameterList
      Dim vBatchAmount As Double
      Dim vCurrencyAmount As Double
      Dim vWhereFields As New CDBFields()
      Dim vAttribute As String
      Dim vResult As New ParameterList()
      Dim vCount As Integer

      vAttribute = "tt.transaction_sign" + ",amount" + ",currency_amount"
      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pAdjBatchNumber)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttribute, "batch_transactions bt", vWhereFields, "", vAnsiJoins)
      Dim vAddItems As String = ""
      Dim vDataTable As CDBDataTable = New CDBDataTable
      vDataTable.FillFromSQL(mvEnv, vSQLStatement)

      If vDataTable IsNot Nothing Then
        For Each vRow As CDBDataRow In vDataTable.Rows
          vCount = vCount + 1
          If vRow.Item("transaction_sign") = "D" Then
            vBatchAmount = vBatchAmount - Val(vRow.Item("amount"))
            vCurrencyAmount = vCurrencyAmount - Val(vRow.Item("currency_amount"))
          Else
            vBatchAmount = vBatchAmount + Val(vRow.Item("amount"))
            vCurrencyAmount = vCurrencyAmount + Val(vRow.Item("currency_amount"))
          End If
        Next
      End If
      vResult.Add("Amount", vBatchAmount)
      vResult.Add("CurrencyAmount", vCurrencyAmount)
      vResult.Add("Count", vCount)
      Return vResult
    End Function

    Friend ReadOnly Property Key() As String
      Get
        Return mvClassFields.Item(FinancialHistoryFields.fhfBatchNumber).Value.PadLeft(9, "0"c) & mvClassFields.Item(FinancialHistoryFields.fhfTransactionNumber).Value.PadLeft(4, "0"c)
      End Get
    End Property

    Friend Sub CreateFromBatchTransaction(ByVal pBT As BatchTransaction)
      With mvClassFields
        .Item(FinancialHistoryFields.fhfBatchNumber).IntegerValue = pBT.BatchNumber
        .Item(FinancialHistoryFields.fhfTransactionNumber).IntegerValue = pBT.TransactionNumber
        .Item(FinancialHistoryFields.fhfContactNumber).IntegerValue = pBT.ContactNumber
        .Item(FinancialHistoryFields.fhfAddressNumber).IntegerValue = pBT.AddressNumber
        .Item(FinancialHistoryFields.fhfTransactionDate).Value = pBT.TransactionDate
        .Item(FinancialHistoryFields.fhfTransactionType).Value = pBT.TransactionType
        .Item(FinancialHistoryFields.fhfAmount).DoubleValue = pBT.Amount
        .Item(FinancialHistoryFields.fhfPaymentMethod).Value = pBT.PaymentMethod
        .Item(FinancialHistoryFields.fhfReference).Value = pBT.Reference
        .Item(FinancialHistoryFields.fhfPosted).Value = TodaysDate()
        .Item(FinancialHistoryFields.fhfNotes).Value = pBT.Notes
        .Item(FinancialHistoryFields.fhfTransactionOrigin).Value = pBT.TransactionOrigin
        If pBT.BankDetailsNumber > 0 Then .Item(FinancialHistoryFields.fhfBankDetailsNumber).IntegerValue = pBT.BankDetailsNumber
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then .Item(FinancialHistoryFields.fhfCurrencyAmount).DoubleValue = pBT.CurrencyAmount
      End With
    End Sub

    ''' <summary>When changing the payer (moving) of a transaction that had invoice allocations, update the paid <see cref="Invoice">Invoice</see> and <see cref="CreditCustomer">CreditCustomer</see>.</summary>
    ''' <param name="pOrigBatchNumber">Batch number of the original transaction</param>
    ''' <param name="pOrigTransNumber">Transaction number of the original transaction</param>
    ''' <param name="pPositiveLinesOnly">Include positive lines only?</param>
    ''' <param name="pReverseData">Roll back the changes due to a failure of the move?</param>
    Private Sub UpdateInvoiceForMove(ByVal pOrigBatchNumber As Integer, ByVal pOrigTransNumber As Integer, ByVal pPositiveLinesOnly As Boolean, ByVal pReverseData As Boolean)
      'If batch contains invoice payments then update Invoice being paid and CreditCustomers record
      Dim vAnsiJoins As New AnsiJoins()
      If pReverseData Then vAnsiJoins.Add("batches b", "iph.batch_number", "b.batch_number")
      vAnsiJoins.Add("invoices i", "iph.invoice_number", "i.invoice_number")

      Dim vWhereFields As New CDBFields(New CDBField("iph.batch_number", pOrigBatchNumber))
      With vWhereFields
        .Add("iph.transaction_number", pOrigTransNumber)
        If pPositiveLinesOnly Then .Add("iph.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        If pReverseData Then .Add("b.batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(Batch.BatchTypes.CreditSales), CDBField.FieldWhereOperators.fwoNotEqual)
      End With

      Dim vCreditCustomer As New CreditCustomer()
      vCreditCustomer.Init(mvEnv)
      Dim vInvoicePaid As New Invoice()
      vInvoicePaid.Init(mvEnv)

      Dim vGotAllocations As Boolean = False
      Dim vPayAmount As Double = 0

      Dim vAttrs As String = vInvoicePaid.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", iph.amount AS iph_amount, iph.allocation_batch_number, iph.allocation_transaction_number"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "invoice_payment_history iph", vWhereFields, "iph.line_number", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRS.Fetch
        If ((vRS.Fields("allocation_batch_number").IntegerValue <> pOrigBatchNumber OrElse vRS.Fields("allocation_transaction_number").IntegerValue <> pOrigTransNumber) AndAlso vRS.Fields("allocation_batch_number").IntegerValue > 0) Then
          'Payment was allocated to the Invoice in a separate batch, therefore leave the Invoice as it is
        Else
          vInvoicePaid = New Invoice
          vInvoicePaid.InitFromRecordSet(mvEnv, vRS, Invoice.InvoiceRecordSetTypes.irtAll)
          vPayAmount = vRS.Fields("iph_amount").DoubleValue
          If pReverseData Then vPayAmount = (vPayAmount * -1)
          If vCreditCustomer.Company <> vInvoicePaid.Company OrElse vCreditCustomer.SalesLedgerAccount <> vInvoicePaid.SalesLedgerAccount Then
            If vCreditCustomer.Existing Then vCreditCustomer.Save()
            vCreditCustomer = New CreditCustomer()
            vCreditCustomer.InitCompanySalesLedgerAccount(mvEnv, vInvoicePaid.Company, vInvoicePaid.SalesLedgerAccount)
          End If
          If vCreditCustomer.Existing Then vCreditCustomer.AdjustOutstanding((vPayAmount * -1))
          If vInvoicePaid.Existing Then
            vInvoicePaid.SetAmountPaid(vPayAmount, False)
            vInvoicePaid.Save()
          End If
          vGotAllocations = True
        End If
      End While
      vRS.CloseRecordSet()
      If vCreditCustomer.Existing Then vCreditCustomer.Save()

      If vGotAllocations = True AndAlso pPositiveLinesOnly = True Then
        'Update CreditCustomers for any Un-allocated S/L Cash if we had some Invoice allocations
        vCreditCustomer = New CreditCustomer()
        vCreditCustomer.Init(mvEnv)
        With vAnsiJoins
          .Clear()
          .Add("invoice_details cid", "ci.invoice_number", "cid.invoice_number", "ci.batch_number", "cid.batch_number", "ci.transaction_number", "cid.transaction_number")
          .Add("batch_transaction_analysis bta", "cid.batch_number", "bta.batch_number", "cid.transaction_number", "bta.transaction_number", "cid.line_number", "bta.line_number")
        End With
        With vWhereFields
          .Clear()
          .Add("bta.batch_number", pOrigBatchNumber)
          .Add("bta.transaction_number", pOrigTransNumber)
          .Add("bta.line_type", "U")
          .Add("bta.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End With
        vSQLStatement = New SQLStatement(mvEnv.Connection, "ci.batch_number, ci.transaction_number, ci.company, ci.sales_ledger_account, bta.amount", "invoices ci", vWhereFields, "ci.company, ci.sales_ledger_account", vAnsiJoins)
        vRS = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          vPayAmount = (vRS.Fields("amount").DoubleValue * -1)
          If pReverseData Then vPayAmount = (vPayAmount * -1)
          If vCreditCustomer.Company <> vRS.Fields("company").Value OrElse vCreditCustomer.SalesLedgerAccount <> vRS.Fields("sales_ledger_account").Value Then
            If vCreditCustomer.Existing Then vCreditCustomer.Save(mvEnv.User.UserID, True)
            vCreditCustomer = New CreditCustomer()
            vCreditCustomer.InitCompanySalesLedgerAccount(mvEnv, vRS.Fields("company").Value, vRS.Fields("sales_ledger_account").Value)
          End If
          If vCreditCustomer.Existing Then vCreditCustomer.AdjustOutstanding(vPayAmount)
        End While
        vRS.CloseRecordSet()
        If vCreditCustomer.Existing Then vCreditCustomer.Save(mvEnv.User.UserID, True)
      End If
    End Sub

    ''' <summary>Is this transaction the result of a financial adjustment?</summary>
    ''' <returns>True if the transaction is the result of a financial adjustment, otherwise False</returns>
    Public ReadOnly Property IsFinancialAdjustment() As Boolean
      Get
        Dim vWhereFields As New CDBFields({New CDBField("batch_number", BatchNumber), New CDBField("transaction_number", TransactionNumber)})
        If mvEnv.Connection.GetCount("reversals", vWhereFields) > 0 Then
          Return True
        Else
          Return False
        End If
      End Get
    End Property

    ''' <summary>Can a financial adjustment transaction be further adjusted?</summary>
    ''' <param name="pSLItems">The Sales Ledger items (if any) it contains.</param>
    ''' <returns>True if the transaction can be adjusted, otherwise False.</returns>
    Private Function CanReverseAdjustedTransaction(ByVal pSLItems As SalesLedgerItems) As Boolean
      Dim vCanAdjust As Boolean = True
      If pSLItems > SalesLedgerItems.None Then
        vCanAdjust = CanReverseAdjustedTransaction()
      End If
      Return vCanAdjust
    End Function
    Private Function CanReverseAdjustedTransaction() As Boolean
      Dim vCanAdjust As Boolean = True
      Dim vWhereFields As New CDBFields({New CDBField("r.batch_number", BatchNumber), New CDBField("r.transaction_number", TransactionNumber)})
      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("financial_history fh", "r.was_batch_number", "fh.batch_number", "r.was_transaction_number", "fh.transaction_number")})
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "fh.status", "reversals r", vWhereFields, "", vAnsiJoins)
      Dim vStatus As String = vSQLStatement.GetValue()
      If (String.IsNullOrWhiteSpace(vStatus) = False AndAlso vStatus.Equals(FA_STATUS_ADJUSTMENT, StringComparison.InvariantCultureIgnoreCase)) Then
        'Original transaction has a status of 'A' (Adjusted)
        'Can only adjust if this amount is zero (it's a re-analysis)
        vCanAdjust = (Me.Amount.Equals(0))
      Else
        'Original transaction has some other status
        'Cannot adjust
        vCanAdjust = False
      End If
      Return vCanAdjust
    End Function
    ''' <summary>Does the transaction contain sales ledger items?</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pBatchNumber">The batch number to check.</param>
    ''' <param name="pTransactionNumber">The transaction number to check.</param>
    ''' <returns><see cref="FinancialHistory.SalesLedgerItems">ContainsSLItems</see> enumeration detailing the type of Sales Ledger items present</returns>
    ''' <remarks>The Sales Ledger items checked are: 'N' Invoice payments, 'U' Unallocated Sales Ledger Cash, 'L' Sales Ledger Cash Allocations and 'K' Credit Note Allocations.</remarks>
    Private Shared Function ContainsSalesLedgerItems(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer) As SalesLedgerItems
      Return ContainsSalesLedgerItems(pEnv, pBatchNumber, pTransactionNumber, 0)
    End Function
    ''' <summary>Does the transaction contain sales ledger items?</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pBatchNumber">The batch number to check.</param>
    ''' <param name="pTransactionNumber">The transaction number to check.</param>
    ''' <param name="pLineNumber">The line number to check when checks are required at line level.</param>
    ''' <returns><see cref="FinancialHistory.SalesLedgerItems">ContainsSLItems</see> enumeration detailing the type of Sales Ledger items present</returns>
    ''' <remarks>The Sales Ledger items checked are: 'N' Invoice payments, 'U' Unallocated Sales Ledger Cash, 'L' Sales Ledger Cash Allocations and 'K' Credit Note Allocations.</remarks>
    Public Shared Function ContainsSalesLedgerItems(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer) As SalesLedgerItems
      Dim vSLItems As SalesLedgerItems = SalesLedgerItems.None

      Dim vWhereFields As New CDBFields(New CDBField("batch_number", pBatchNumber))
      vWhereFields.Add("transaction_number", pTransactionNumber)
      If pLineNumber > 0 Then vWhereFields.Add("line_number", pLineNumber)
      vWhereFields.Add("line_type", CDBField.FieldTypes.cftCharacter, "'U','N','L','K'", CDBField.FieldWhereOperators.fwoIn)

      Dim vSQLStatement As New SQLStatement(pEnv.Connection, "line_type, COUNT(*) AS line_count", "batch_transaction_analysis", vWhereFields)
      vSQLStatement.GroupBy = "line_type"

      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRS.Fetch
        If vRS.Fields(2).IntegerValue > 0 Then
          Select Case vRS.Fields(1).Value.ToUpper
            Case "K"
              vSLItems = (vSLItems Or SalesLedgerItems.CreditNoteAllocation)
            Case "L"
              vSLItems = (vSLItems Or SalesLedgerItems.SLCashAllocation)
            Case "N"
              vSLItems = (vSLItems Or SalesLedgerItems.InvoicePayments)
            Case "U"
              vSLItems = (vSLItems Or SalesLedgerItems.UnallocatedSLCash)
          End Select
        End If
      End While
      vRS.CloseRecordSet()

      Return vSLItems

    End Function

    ''' <summary>Can a Receipt be printed for this transaction?</summary>
    Public Function CanPrintReceipt() As Boolean
      Dim vCanPrint As Boolean = False
      If Existing = True AndAlso Amount >= 0 AndAlso IsFinancialAdjustment() = False Then
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReceiptPrintStdDocument).Length > 0 Then vCanPrint = True
      End If

      Return vCanPrint
    End Function

    ''' <summary>When reversing or re-analysing an un-allocated sales ledger cash line find and reverse any invoice payment history.</summary>
    ''' <param name="pOriginalLineNumber">Line number of the original analysis line.</param>
    ''' <param name="pOriginalLineType">Line type of the original analysis line.</param>
    ''' <param name="pOriginalAmount">Amount of the original analysis line.</param>
    ''' <param name="pNewBTA">BatchTransactionAnalysis for the new line.</param>
    ''' <param name="pTransactionDate">Date of the adjustment transaction.</param>
    ''' <param name="pAdjustmentStatus">The adjustment status.</param>
    Private Sub ReverseInvoiceCashAllocation(ByVal pOriginalLineNumber As Integer, ByVal pOriginalLineType As String, ByVal pOriginalAmount As Double, ByVal pNewBTA As BatchTransactionAnalysis, ByVal pTransactionDate As Date, ByVal pAdjustmentStatus As String)
      If pOriginalLineType.Equals("U") Then
        Dim vCashInvoice As New Invoice()
        vCashInvoice.Init(mvEnv, Me.BatchNumber, Me.TransactionNumber)
        If vCashInvoice.Existing = True AndAlso vCashInvoice.RecordType.Equals(Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash)) Then
          Dim vCashAmount As Double = vCashInvoice.InvoiceAmount
          If Me.Amount.Equals(0) = True AndAlso vCashInvoice.IsFinancialAdjustmentInvoice = True Then vCashAmount = Me.Amount
          Dim vCurrentUnallocated As Double = FixTwoPlaces(vCashAmount - vCashInvoice.AmountPaid)
          If vCurrentUnallocated.Equals(0) OrElse vCurrentUnallocated.CompareTo(pOriginalAmount) < 0 Then
            'Some of this has been allocated
            Dim vIPH As New InvoicePaymentHistory(mvEnv)
            vIPH.Init()
            Dim vInvoice As New Invoice()
            vInvoice.Init(mvEnv)
            Dim vWhereFields As New CDBFields(New CDBField("iph.batch_number", Me.BatchNumber))
            vWhereFields.Add("iph.transaction_number", Me.TransactionNumber)
            vWhereFields.Add("iph.amount", CDBField.FieldTypes.cftNumeric, pOriginalAmount.ToString, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vWhereFields.Add("iph.allocation_batch_number", CDBField.FieldTypes.cftInteger, "iph.batch_number", CDBField.FieldWhereOperators.fwoNotEqual)
            vWhereFields.Add("iph.status", CDBField.FieldTypes.cftCharacter, String.Empty)

            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vIPH.GetRecordSetFields, "invoice_payment_history iph", vWhereFields, "iph.allocation_batch_number, iph.allocation_transaction_number")
            Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
            Dim vInTrans As Boolean = False
            Dim vAllocationsRemoved As Double = 0

            While vRS.Fetch = True AndAlso vAllocationsRemoved <= pOriginalAmount
              vIPH = New InvoicePaymentHistory(mvEnv)
              vIPH.InitFromRecordSet(vRS)
              vAllocationsRemoved += vIPH.Amount
              If vIPH.InvoiceNumber.Equals(DoubleValue(vInvoice.InvoiceNumber)) = False Then
                vInvoice = New Invoice()
                vInvoice.Init(mvEnv, 0, 0, vIPH.InvoiceNumber)
              End If
              vInTrans = mvEnv.Connection.StartTransaction()
              'Update the Invoice we paid to reduce the amount paid
              If vInvoice.Existing Then
                vInvoice.SetAmountPaid((vIPH.Amount * -1), True)
                vInvoice.Save(mvEnv.User.UserID, True)
              End If
              'Add Invoice Payment History for the reversal
              vIPH.Reverse(pAdjustmentStatus, pNewBTA.BatchNumber, pNewBTA.TransactionNumber, pNewBTA.LineNumber, pTransactionDate)
              If vInTrans Then mvEnv.Connection.CommitTransaction()
            End While
            vRS.CloseRecordSet()
          End If
        End If
      End If
    End Sub

  End Class
End Namespace
