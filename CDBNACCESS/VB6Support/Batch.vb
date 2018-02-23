Imports Advanced.LanguageExtensions
Namespace Access

  Public Class Batch

    Public Enum BatchRecordSetTypes 'These are bit values
      brtAll = &HFFFFS
      'ADD additional recordset types here
      brtNumber = 1
      brtType = 2
    End Enum

    Private Enum TransactionStockStatus
      tssNoStock 'No stock items in transaction
      tssStockNothingDespatched 'Stock items in transaction but none despatched
      tssStockSomeDespatched 'Stock items in transaction some despatched
    End Enum

    Public Enum DeleteBatchAllowed
      dbNo
      dbYes
      dbWarn
    End Enum

    Private Enum FHDSourceOrigin
      fhdsoBatchTransactionAnalysis
      fhdsoPaymentPlanDetails
    End Enum

    'FHD Distribution Code Origin for Payment Plan Payments only
    Private Enum FHDDistCodeOriginPPP
      PaymentPlanDetails
      BatchTransactionAnalysis
    End Enum

    Public Enum AdjustmentTypes
      atNone
      atAdjustment
      atMove
      atRefund
      atReverse
      atGIKConfirmation
      atPartRefund
      atCashBatchConfirmation
      atEventAdjustment 'Used by Smart Client only
    End Enum

    'Standard Class Setup
    Private mvJob As JobSchedule 'The batch processing job being run
    Private WithEvents mvCCA As CreditCardAuthorisation

    'The following variables are used by batch processing
    Private mvConn As CDBConnection
    Private mvCreateFinancialHistory As Boolean
    Private mvCreateBackOrder As Boolean
    Private mvStartReceiptNumber As Integer
    Private mvProductList As String

    Private mvInvoicesToDelete As String

    Private mvGAOperationalChangeDate As String
    Private mvGAMembershipTaxReclaim As Boolean
    Private mvCAFPaymentMethod As String
    Private mvGraceDays As Integer
    Private mvImportPayment As Boolean 'Used by Data Import so that zero balance Payment Plans can have Gift Aid unclaimed lines created

    Private mvGAInitialized As Boolean

    Private mvAmendedValid As Boolean

    Private mvFHDSourceOrigin As FHDSourceOrigin
    Private mvFHDDistCodeOriginPPP As FHDDistCodeOriginPPP

    Private mvHoldingContactChecked As Boolean
    Private mvUsesHoldingContact As Boolean

    Public Event AuthorisingCreditCard(ByRef pMaxTime As Integer, ByRef pTime As Integer)

    Private mvBatchLocked As Boolean
    Private mvOpenBatchType As String
    Private mvCompanyControl As CompanyControl
    Private mvBatchType As BatchTypes

    'Payment plan related batch processing items
    Private mvMeFutureChange As String
    Private mvMeFutureChangeTrigger As String
    Private mvOrgRenewalDate As String
    Private mvOrgTerm As Integer
    Private mvOrgBalance As Double
    Private mvMemTypeTerm As MembershipType.MembershipTypeTerms

    'Cancellation info for Group One Year Gift Membership
    Private mvGOYGReason As String
    Private mvGOYGStatus As String
    Private mvGOYGDesc As String

    'Cancellation info for cancelling Non Group One Year Gift Membership Auto Payment Method (where cancel_one_year_gift_apm flag is set) 
    Private mvNonGroupOYGAutoPayReason As String
    Private mvNonGroupOYGAutoPayStatus As String
    Private mvNonGroupOYGAutoPayDesc As String

    'Cancellation info for One Off Payment PaymentPlan
    Private mvOOPReason As String
    Private mvOOPStatus As String
    Private mvOOPDesc As String

    Private mvBranchIncomePeriod As String
    Private mvDeedOrder As Boolean

    Private mvIgnoreDiscountForSkip As Boolean    'When skipping PP payments, this will ignore discount lines

    Protected Overrides Sub ClearFields()
      mvBatchLocked = False
      mvOpenBatchType = ""
      mvCompanyControl = Nothing
      mvBatchType = BatchTypes.None
      mvMeFutureChange = ""
      mvMeFutureChangeTrigger = ""
      mvMemTypeTerm = MembershipType.MembershipTypeTerms.mtfAnnualTerm
      mvOrgBalance = 0
      mvOrgRenewalDate = ""
      mvOrgTerm = 0
      mvGOYGReason = ""
      mvGOYGStatus = ""
      mvGOYGDesc = ""
      mvOOPReason = ""
      mvOOPStatus = ""
      mvOOPDesc = ""
      mvBranchIncomePeriod = ""
      mvDeedOrder = False
      mvNonGroupOYGAutoPayReason = ""
      mvNonGroupOYGAutoPayStatus = ""
      mvNonGroupOYGAutoPayDesc = ""
    End Sub

    Friend Sub SetNextTransactionNumber(ByVal pTransactionNumber As Integer)
      If pTransactionNumber > mvClassFields.Item(BatchFields.NextTransactionNumber).IntegerValue Then
        mvClassFields.Item(BatchFields.NextTransactionNumber).IntegerValue = pTransactionNumber + 1
      End If
    End Sub

    Private Function CheckInAdvance(ByVal pBatchNumber As Long, ByVal pTransactionNumber As Long, Optional ByVal pLineNumber As Long = 0) As Boolean
      Dim vSQL As String

      If mvCompanyControl.InAdvanceProductCode.Length > 0 Then
        vSQL = "SELECT r.batch_number FROM reversals r, order_payment_history oph, financial_history_details fhd WHERE "
        vSQL = vSQL & "r.batch_number = " & pBatchNumber & " AND r.transaction_number = " & pTransactionNumber
        If pLineNumber > 0 Then vSQL = vSQL & " AND r.line_number = " & pLineNumber
        vSQL = vSQL & " AND oph.batch_number = r.was_batch_number AND oph.transaction_number = r.was_transaction_number AND oph.line_number = r.was_line_number"
        vSQL = vSQL & " AND fhd.batch_number = r.was_batch_number AND fhd.transaction_number = r.was_transaction_number AND fhd.line_number = r.was_line_number"
        vSQL = vSQL & " AND fhd.product = '" & mvCompanyControl.InAdvanceProductCode & "' AND fhd.rate = '" & mvCompanyControl.InAdvanceRate & "'"
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() Then CheckInAdvance = True
        vRecordSet.CloseRecordSet()
      End If
    End Function


    'AFTER HERE ADDED ------------------------------------------
    Public Function GetNextTransactionNumber() As Integer
      'This method should only be used when the calling application/routine/method has exclusive access to this batch and no other process(es) could be adding transactions to this batch.
      'If the calling application/routine/method does not have exclusive access to this batch then the AllocateTransactionNumber method should be used.
      Dim vNextTransactionNumber As Integer

      If mvClassFields.Item(BatchFields.BatchNumber).IntegerValue > 0 Then
        vNextTransactionNumber = mvClassFields.Item(BatchFields.NextTransactionNumber).IntegerValue
        If vNextTransactionNumber > mvMaximumTransactions Then
          Save()
          mvClassFields.ClearSetValues()
          mvExisting = False
          mvClassFields.Item(BatchFields.BatchNumber).IntegerValue = mvEnv.GetControlNumber("B")

          mvClassFields.Item(BatchFields.BatchTotal).DoubleValue = 0
          mvClassFields.Item(BatchFields.TransactionTotal).DoubleValue = 0

          mvClassFields.Item(BatchFields.NumberOfTransactions).IntegerValue = 0
          mvClassFields.Item(BatchFields.NumberOfEntries).IntegerValue = 0

          'If mvEnv.GetDataStructureInfo(cdbDataCurrencyCode) Then
          mvClassFields.Item(BatchFields.CurrencyBatchTotal).DoubleValue = 0
          mvClassFields.Item(BatchFields.CurrencyTransactionTotal).DoubleValue = 0
          'End If

          mvClassFields.Item(BatchFields.NextTransactionNumber).IntegerValue = 2
          vNextTransactionNumber = 1 'Return next transaction is number one
        Else
          mvClassFields.Item(BatchFields.NextTransactionNumber).IntegerValue = vNextTransactionNumber + 1
        End If
        GetNextTransactionNumber = vNextTransactionNumber
      Else
        'TODO An Error? "Batch has not been initialised"
      End If
    End Function
    Public Overloads Function GetRecordSetFields(ByVal pEnv As CDBEnvironment, ByVal pRSType As BatchRecordSetTypes) As String
      Dim vFields As String

      mvEnv = pEnv
      'Modify below to add each recordset type as required
      If pRSType = BatchRecordSetTypes.brtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "b")
      Else
        vFields = "b.batch_number" 'Always get the number
        If (pRSType And BatchRecordSetTypes.brtType) > 0 Then vFields = vFields & ",b.batch_date,b.bank_account,b.batch_type"
      End If
      Return vFields
    End Function
    'Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0)
    '  Dim vRecordSet As CDBRecordSet

    '  mvEnv = pEnv
    '  InitClassFields()
    '  If pBatchNumber > 0 Then
    '    vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(pEnv, BatchRecordSetTypes.brtAll) & " FROM batches b WHERE b.batch_number = " & pBatchNumber)
    '    If vRecordSet.Fetch() = True Then InitFromRecordSet(pEnv, vRecordSet, BatchRecordSetTypes.brtAll)
    '    vRecordSet.CloseRecordSet()
    '  Else
    '    SetDefaults()
    '  End If
    'End Sub

    Public Sub InitNewBatch(ByVal pEnv As CDBEnvironment)
      'Create new batch and allocate batch number
      MyBase.Init()
      mvClassFields.Item(BatchFields.BatchNumber).Value = CStr(mvEnv.GetControlNumber("B"))
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BatchRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Modify below to handle each recordset type as required
        .SetItem(BatchFields.BatchNumber, vFields) 'Always get the number
        If (pRSType And BatchRecordSetTypes.brtType) > 0 Then
          .SetItem(BatchFields.BatchType, vFields)
          .SetItem(BatchFields.BatchDate, vFields)
          .SetItem(BatchFields.BankAccount, vFields)
        End If
        If (pRSType And BatchRecordSetTypes.brtAll) = BatchRecordSetTypes.brtAll Then
          .SetItem(BatchFields.CashBookBatch, vFields)
          .SetItem(BatchFields.NumberOfEntries, vFields)
          .SetItem(BatchFields.BatchTotal, vFields)
          .SetItem(BatchFields.TransactionTotal, vFields)
          .SetItem(BatchFields.NumberOfTransactions, vFields)
          .SetItem(BatchFields.NextTransactionNumber, vFields)
          .SetItem(BatchFields.ReadyForBanking, vFields)
          .SetItem(BatchFields.PayingInSlipPrinted, vFields)
          .SetItem(BatchFields.PostedToCashBook, vFields)
          .SetItem(BatchFields.DetailCompleted, vFields)
          .SetItem(BatchFields.PostedToNominal, vFields)
          .SetItem(BatchFields.Picked, vFields)
          .SetItem(BatchFields.Product, vFields)
          .SetItem(BatchFields.Rate, vFields)
          .SetItem(BatchFields.Source, vFields)
          .SetItem(BatchFields.TransactionType, vFields)
          .SetItem(BatchFields.PaymentMethod, vFields)
          .SetItem(BatchFields.PayingInSlipNumber, vFields)
          .SetItem(BatchFields.CurrencyIndicator, vFields)
          .SetItem(BatchFields.CurrencyBatchTotal, vFields)
          .SetItem(BatchFields.CurrencyTransactionTotal, vFields)
          .SetItem(BatchFields.CurrencyExchangeRate, vFields)
          .SetItem(BatchFields.AmendedBy, vFields)
          .SetItem(BatchFields.AmendedOn, vFields)
          .SetItem(BatchFields.JournalNumber, vFields)
          .SetItem(BatchFields.BatchCategory, vFields)
          .SetItem(BatchFields.BalancedBy, vFields)
          .SetItem(BatchFields.BalancedOn, vFields)
          .SetItem(BatchFields.PostedBy, vFields)
          .SetItem(BatchFields.PostedOn, vFields)
          .SetItem(BatchFields.ContentsAmendedBy, vFields)
          .SetItem(BatchFields.ContentsAmendedOn, vFields)
          .SetItem(BatchFields.HeaderAmendedBy, vFields)
          .SetItem(BatchFields.HeaderAmendedOn, vFields)
          .SetItem(BatchFields.BatchCreatedBy, vFields)
          .SetItem(BatchFields.BatchCreatedOn, vFields)
          .SetItem(BatchFields.PostNominal, vFields)
          .SetItem(BatchFields.JobNumber, vFields)
          .SetOptionalItem(BatchFields.CurrencyCode, vFields)
          .SetOptionalItem(BatchFields.Provisional, vFields)
          .SetOptionalItem(BatchFields.AgencyNumber, vFields)
          .SetOptionalItem(BatchFields.ClaimSent, vFields)
          .SetOptionalItem(BatchFields.BatchAnalysisCode, vFields)
          .SetOptionalItem(BatchFields.Campaign, vFields)
          .SetOptionalItem(BatchFields.Appeal, vFields)
          .SetOptionalItem(BatchFields.BankingDate, vFields)
        End If
      End With
    End Sub

    Public Overrides Sub Update(ByVal pParams As CDBParameters)
      'used in SmartClient
      If pParams.Exists("PayingInSlipNumber") Then mvClassFields.Item(BatchFields.PayingInSlipNumber).Value = pParams("PayingInSlipNumber").Value
      If pParams.Exists("PayingInSlipPrinted") Then mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = pParams("PayingInSlipPrinted").Bool
      If pParams.Exists("ReadyForBanking") Then
        If ReadyForBanking = False AndAlso pParams("ReadyForBanking").Bool AndAlso mvEnv.GetConfigOption("fp_no_paying_in_slip_required") Then
          mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = True
        End If
        mvClassFields.Item(BatchFields.ReadyForBanking).Bool = pParams("ReadyForBanking").Bool
      End If
      If pParams.Exists("BankAccount") Then mvClassFields.Item(BatchFields.BankAccount).Value = pParams("BankAccount").Value
      If pParams.Exists("PaymentMethod") Then mvClassFields.Item(BatchFields.PaymentMethod).Value = pParams("PaymentMethod").Value
      If pParams.Exists("TransactionType") Then mvClassFields.Item(BatchFields.TransactionType).Value = pParams("TransactionType").Value
      If pParams.Exists("BatchCategory") Then mvClassFields.Item(BatchFields.BatchCategory).Value = pParams("BatchCategory").Value
      If pParams.Exists("NumberOfEntries") Then mvClassFields.Item(BatchFields.NumberOfEntries).IntegerValue = pParams("NumberOfEntries").IntegerValue
      If pParams.Exists("BatchTotal") Then mvClassFields.Item(BatchFields.BatchTotal).DoubleValue = pParams("BatchTotal").DoubleValue
      If pParams.Exists("CurrencyBatchTotal") Then mvClassFields.Item(BatchFields.CurrencyBatchTotal).DoubleValue = pParams("CurrencyBatchTotal").DoubleValue
      If pParams.Exists("CurrencyTransactionTotal") Then mvClassFields.Item(BatchFields.CurrencyTransactionTotal).DoubleValue = pParams("CurrencyTransactionTotal").DoubleValue
      If pParams.Exists("Product") Then mvClassFields.Item(BatchFields.Product).Value = pParams("Product").Value
      If pParams.Exists("Rate") Then mvClassFields.Item(BatchFields.Rate).Value = pParams("Rate").Value
      If pParams.Exists("Source") Then mvClassFields.Item(BatchFields.Source).Value = pParams("Source").Value
      If pParams.Exists("AgencyNumber") Then mvClassFields.Item(BatchFields.AgencyNumber).IntegerValue = pParams("AgencyNumber").IntegerValue
      If pParams.Exists("CurrencyCode") Then mvClassFields.Item(BatchFields.CurrencyCode).Value = pParams("CurrencyCode").Value
      If pParams.Exists("CurrencyExchangeRate") Then mvClassFields.Item(BatchFields.CurrencyExchangeRate).DoubleValue = pParams("CurrencyExchangeRate").DoubleValue
      If pParams.Exists("BatchAnalysisCode") Then mvClassFields.Item(BatchFields.BatchAnalysisCode).Value = pParams("BatchAnalysisCode").Value
      If pParams.Exists("DetailCompleted") Then mvClassFields.Item(BatchFields.DetailCompleted).Bool = pParams("DetailCompleted").Bool
      If Existing And Not PostedToCashBook Then
        If pParams.Exists("BalancedOn") Then mvClassFields.Item(BatchFields.BalancedOn).Value = pParams("BalancedOn").Value
        If pParams.Exists("PostedOn") Then mvClassFields.Item(BatchFields.PostedOn).Value = pParams("PostedOn").Value
        If pParams.Exists("ContentsAmendedOn") Then mvClassFields.Item(BatchFields.ContentsAmendedOn).Value = pParams("ContentsAmendedOn").Value
        If pParams.Exists("HeaderAmendedOn") Then mvClassFields.Item(BatchFields.HeaderAmendedOn).Value = pParams("HeaderAmendedOn").Value
        If pParams.Exists("BatchCreatedOn") Then mvClassFields.Item(BatchFields.BatchCreatedOn).Value = pParams("BatchCreatedOn").Value
        If pParams.Exists("BankingDate") Then mvClassFields.Item(BatchFields.BankingDate).Value = pParams("BankingDate").Value
        If pParams.Exists("BatchDate") Then mvClassFields.Item(BatchFields.BatchDate).Value = pParams("BatchDate").Value
      End If
    End Sub
    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Public ReadOnly Property Locked() As Boolean
      Get
        Locked = mvBatchLocked
      End Get
    End Property

    Public ReadOnly Property UsesHoldingContact() As Boolean
      Get
        'Check if this batch contains transactions using the holding_contact_number
        Dim vFields As New CDBFields
        Dim vContact As Integer

        If mvHoldingContactChecked = False Then
          vContact = CInt(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber))
          If vContact > 0 Then
            vFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            vFields.Add("contact_number", CDBField.FieldTypes.cftLong, vContact)
            If mvEnv.Connection.GetCount("batch_transactions", vFields) > 0 Then mvUsesHoldingContact = True
          End If
          mvHoldingContactChecked = True
        End If
        UsesHoldingContact = mvUsesHoldingContact
      End Get
    End Property

    Public ReadOnly Property DataTableColumns() As String
      Get
        Dim vColumns As String

        vColumns = "BatchNumber,BatchType,BatchDate,CashBookBatch,BankAccount,NumberOfEntries,BatchTotal,TransactionTotal,"
        vColumns = vColumns & "NumberOfTransactions,NextTransactionNumber,ReadyForBanking,PayingInSlipPrinted,PostedToCashBook,"
        vColumns = vColumns & "DetailCompleted,PostedToNominal,Picked,Product,Rate,Source,TransactionType,PaymentMethod,"
        vColumns = vColumns & "PayingInSlipNumber,CurrencyIndicator,CurrencyBatchTotal,CurrencyTransactionTotal,CurrencyExchangeRate,"
        vColumns = vColumns & "AmendedBy,AmendedOn,JournalNumber,BatchCategory,BalancedBy,BalancedOn,PostedBy,PostedOn,"
        vColumns = vColumns & "ContentsAmendedBy,ContentsAmendedOn,HeaderAmendedBy,HeaderAmendedOn,BatchCreatedBy,"
        vColumns = vColumns & "BatchCreatedOn,PostNominal,JobNumber,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vColumns = vColumns & "CurrencyCode,"
        vColumns = vColumns & "Provisional,AgencyNumber,ClaimSent"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBatchAnalysisCodes) Then vColumns = vColumns & ",BatchAnalysisCode"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBankingDate) Then vColumns = vColumns & ",BankingDate"

        DataTableColumns = vColumns

      End Get
    End Property

    Public Overrides ReadOnly Property DataTable() As CDBDataTable
      Get
        'This function is only used by WEB Services at present
        Dim vTable As New CDBDataTable
        Dim vRow As CDBDataRow
        Dim vLastColumn As Integer

        With vTable
          .AddColumnsFromList(DataTableColumns)
          vRow = .AddRow
        End With
        With vRow
          .Item(1) = CStr(BatchNumber)
          .Item(2) = mvClassFields.Item(BatchFields.BatchType).Value
          .Item(3) = BatchDate
          .Item(4) = CashBookBatch
          .Item(5) = BankAccount
          .Item(6) = CStr(NumberOfEntries)
          .Item(7) = CStr(BatchTotal)
          .Item(8) = CStr(TransactionTotal)
          .Item(9) = CStr(NumberOfTransactions)
          .Item(10) = CStr(NextTransactionNumber)
          .Item(11) = mvClassFields.Item(BatchFields.ReadyForBanking).Value
          .Item(12) = mvClassFields.Item(BatchFields.PayingInSlipPrinted).Value
          .Item(13) = mvClassFields.Item(BatchFields.PostedToCashBook).Value
          .Item(14) = mvClassFields.Item(BatchFields.DetailCompleted).Value
          .Item(15) = mvClassFields.Item(BatchFields.PostedToNominal).Value
          .Item(16) = Picked
          .Item(17) = ProductCode
          .Item(18) = RateCode
          .Item(19) = Source
          .Item(20) = TransactionType
          .Item(21) = PaymentMethod
          .Item(22) = PayingInSlipNumber.ToString
          .Item(23) = CurrencyIndicator
          .Item(24) = CStr(CurrencyBatchTotal)
          .Item(25) = CStr(CurrencyTransactionTotal)
          .Item(26) = mvClassFields.Item(BatchFields.CurrencyExchangeRate).Value
          .Item(27) = AmendedBy
          .Item(28) = AmendedOn
          .Item(29) = mvClassFields.Item(BatchFields.JournalNumber).Value
          .Item(30) = BatchCategory
          .Item(31) = BalancedBy
          .Item(32) = BalancedOn
          .Item(33) = PostedBy
          .Item(34) = PostedOn
          .Item(35) = ContentsAmendedBy
          .Item(36) = ContentsAmendedOn
          .Item(37) = HeaderAmendedBy
          .Item(38) = HeaderAmendedOn
          .Item(39) = BatchCreatedBy
          .Item(40) = BatchCreatedOn
          .Item(41) = PostNominal
          .Item(42) = mvClassFields.Item(BatchFields.JobNumber).Value
          vLastColumn = 42
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
            .Item(43) = CurrencyCode
            vLastColumn = 43
          End If
          .Item(vLastColumn + 1) = BooleanString(Provisional = True) 'Provisional could be null but want to return Y or N
          .Item(vLastColumn + 2) = mvClassFields.Item(BatchFields.AgencyNumber).Value
          .Item(vLastColumn + 3) = ClaimSent
          vLastColumn = vLastColumn + 3
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBatchAnalysisCodes) Then
            .Item(vLastColumn + 1) = BatchAnalysisCode
            vLastColumn = vLastColumn + 1
          End If
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBankingDate) Then
            .Item(vLastColumn + 1) = BankingDate
            vLastColumn = vLastColumn + 1
          End If
        End With

        DataTable = vTable

      End Get
    End Property
    Public Function SetDetailComplete(ByRef pInvalidTransactions As CDBParameters, Optional ByRef pCheckTransactions As Boolean = True, Optional ByVal pClosingOpenBatch As Boolean = False) As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vDetailOK As Boolean
      Dim vTransNumber As Integer
      Dim vLastTransNumber As Integer
      Dim vLastTransAmount As Double
      Dim vAnalysisTotal As Double

      If mvClassFields.Item(BatchFields.BatchTotal).DoubleValue = mvClassFields.Item(BatchFields.TransactionTotal).DoubleValue And mvClassFields.Item(BatchFields.NumberOfEntries).IntegerValue = mvClassFields.Item(BatchFields.NumberOfTransactions).IntegerValue Then
        vDetailOK = True 'Assume OK
        If pCheckTransactions Then
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT bt.transaction_number, bt.amount, bt.line_total, bta.amount  AS  bta_amount FROM batch_transactions bt, batch_transaction_analysis bta WHERE bt.batch_number = " & mvClassFields.Item(BatchFields.BatchNumber).Value & " AND bt.batch_number = bta.batch_number and bt.transaction_number = bta.transaction_number ORDER BY bt.transaction_number")
          While vRecordSet.Fetch() = True
            With vRecordSet
              vTransNumber = .Fields(1).IntegerValue
              If .Fields(2).DoubleValue <> .Fields(3).DoubleValue Then
                If Not pInvalidTransactions.Exists(CStr(vTransNumber)) Then pInvalidTransactions.Add(CStr(vTransNumber))
                vDetailOK = False 'Amount does not match line total
              Else
                If vTransNumber <> vLastTransNumber Then 'New transaction
                  If FixTwoPlaces(vAnalysisTotal) <> FixTwoPlaces(vLastTransAmount) Then
                    If Not pInvalidTransactions.Exists(CStr(vLastTransNumber)) Then pInvalidTransactions.Add(CStr(vLastTransNumber))
                    vDetailOK = False
                  End If
                  vAnalysisTotal = .Fields(4).DoubleValue
                Else
                  vAnalysisTotal = vAnalysisTotal + .Fields(4).DoubleValue
                End If
                vLastTransNumber = vTransNumber
                vLastTransAmount = .Fields(2).DoubleValue
              End If
            End With
          End While
          vRecordSet.CloseRecordSet()
          If FixTwoPlaces(vAnalysisTotal) <> FixTwoPlaces(vLastTransAmount) Then
            If Not pInvalidTransactions.Exists(CStr(vTransNumber)) Then pInvalidTransactions.Add(CStr(vTransNumber))
            vDetailOK = False
          End If
        End If

        If vDetailOK Then
          mvClassFields.Item(BatchFields.DetailCompleted).Bool = True
          mvClassFields.Item(BatchFields.HeaderAmendedOn).Value = TodaysDate()
          mvClassFields.Item(BatchFields.HeaderAmendedBy).Value = mvEnv.User.UserID
          If Len(mvClassFields.Item(BatchFields.BalancedBy).Value) = 0 Then
            mvClassFields.Item(BatchFields.BalancedOn).Value = TodaysDate()
            mvClassFields.Item(BatchFields.BalancedBy).Value = mvEnv.User.UserID
          End If

          'BR17150 new config options to speed up processing batches
          If pClosingOpenBatch AndAlso mvEnv.GetConfigOption("fp_no_cheque_list_required") Then
            Dim vBatchType As New BatchTypeData
            vBatchType.Init(mvEnv, Batch.GetBatchTypeCode((BatchType)))
            If vBatchType.PrintChequeList Then
              mvClassFields.Item(BatchFields.ReadyForBanking).Value = "Y"
            End If
          End If

          Select Case BatchType
            Case BatchTypes.DebitCard, BatchTypes.CreditCard
              'Changes made to BR15227 - Streamline the batch processing of Credit Card Payments
              Dim vAnsiJoins As New AnsiJoins
              Dim vWhereFields As New CDBFields
              vAnsiJoins.AddLeftOuterJoin("card_sales cs", "bt.batch_number", "cs.batch_number", "bt.transaction_number", "cs.transaction_number")
              vWhereFields.Add("bt.batch_number", BatchNumber)
              vWhereFields.Add("cs.no_claim_required", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
              Dim vSql As New SQLStatement(mvEnv.Connection, "", "batch_transactions bt", vWhereFields, Nothing, vAnsiJoins)
              If mvEnv.Connection.GetCountFromStatement(vSql) = 0 Then
                SetPayingInSlipPrinted(0)
                If Not mvEnv.GetConfigOption("fp_card_batches_to_CB") Then PostedToCashBook = True
              End If

            Case BatchTypes.NonFinancial, BatchTypes.SaleOrReturn, BatchTypes.GiftAidClaimAdjustment
              'TA BR 5763: Should only apply to Non-Financial. If Provisional Or BatchType = btNonFinancial Then
              SetPayingInSlipPrinted(0)
              PostedToCashBook = True
              SetBatchPosted(True)
              Picked = "C"

            Case BatchTypes.CAFCards, BatchTypes.CAFVouchers
              If Provisional Then Picked = "C"

            Case BatchTypes.Cash, BatchTypes.CashWithInvoice
              If Provisional Then
                SetPayingInSlipPrinted(0)
                PostedToCashBook = True
                Picked = "C"
                mvClassFields(BatchFields.PostedToNominal).Bool = True
              End If
          End Select

          If pClosingOpenBatch AndAlso Picked <> "C" AndAlso PayingInSlipPrinted = False AndAlso mvEnv.GetConfigOption("fp_auto_cash_book_posting") Then
            If HasStockItem(BatchNumber) = False Then
              mvClassFields(BatchFields.Picked).Value = "C"
            End If
          End If

          If pClosingOpenBatch AndAlso PayingInSlipPrinted = False AndAlso mvEnv.GetConfigOption("fp_no_paying_in_slip_required") Then
            SetPayingInSlipPrinted(0)
          End If
          Save()
          Return True
        End If
      End If
    End Function
    Public Sub SetPickedForCardSales()
      'Check for any stock products in a credit or debit card batch
      'If no stock then set picked to 'C'
      If BatchType = BatchTypes.CreditCard OrElse BatchType = BatchTypes.DebitCard OrElse BatchType = BatchTypes.CreditCardWithInvoice Then
        If mvClassFields.Item(BatchFields.Picked).Value <> "C" Then
          If mvEnv.Connection.GetCount("batch_transaction_analysis bta, products p", Nothing, "batch_number = " & BatchNumber & " AND bta.product = p.product AND p.stock_item = 'Y'") = 0 Then
            mvClassFields.Item(BatchFields.Picked).Value = "C"
            Save()
          End If
        End If
      End If
    End Sub
    Public Sub SetBatchFlagsFromBackOrders(ByRef pCurrencyCode As String, ByRef pExchangeRate As Double)
      mvClassFields.Item(BatchFields.CurrencyExchangeRate).Value = CStr(pExchangeRate)
      If mvClassFields.Item(BatchFields.CurrencyCode).InDatabase Then mvClassFields.Item(BatchFields.CurrencyCode).Value = pCurrencyCode
      mvClassFields.Item(BatchFields.ReadyForBanking).Bool = True
      If Not (BatchType = BatchTypes.CreditCard OrElse BatchType = BatchTypes.DebitCard OrElse BatchType = BatchTypes.CreditCardWithInvoice) Then
        mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = True
        mvClassFields.Item(BatchFields.PostedToCashBook).Bool = True
      End If
      mvClassFields.Item(BatchFields.DetailCompleted).Bool = True
      mvClassFields.Item(BatchFields.Picked).Value = "C"
    End Sub
    Public Sub SetPayingInSlipPrinted(ByRef pSlipNumber As Integer, Optional ByRef pSlipPrinted As Boolean = True)
      mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = pSlipPrinted
      If pSlipNumber > 0 Then mvClassFields.Item(BatchFields.PayingInSlipNumber).Value = CStr(pSlipNumber)
    End Sub
    Sub LockBatch()
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As CDBFields
      Dim vUpdateFields As CDBFields
      Dim vInsertFields As CDBFields
      Dim vUser As String

      'First see if there is a record for this batch in the open batches table
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT in_use_by,batch_type FROM open_batches WHERE batch_number = " & mvClassFields.Item(BatchFields.BatchNumber).IntegerValue)
      If vRecordSet.Fetch() = True Then
        'There is a record so check who is using it
        vUser = vRecordSet.Fields(1).Value
        mvOpenBatchType = vRecordSet.Fields(2).Value
        If vUser = "" Then
          'If not in use by anyone then update it to flag as in_use_by the user
          vWhereFields = New CDBFields
          vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          vWhereFields.Add("in_use_by", CDBField.FieldTypes.cftCharacter)
          vUpdateFields = New CDBFields
          vUpdateFields.Add("in_use_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
          If mvEnv.Connection.UpdateRecords("open_batches", vUpdateFields, vWhereFields, False) = 0 Then
            vRecordSet.CloseRecordSet()
            RaiseError(DataAccessErrors.daeCannotLockBatch, (mvClassFields.Item(BatchFields.BatchNumber).Value))
          End If
        ElseIf vUser <> mvEnv.User.Logname Then
          vRecordSet.CloseRecordSet()
          RaiseError(DataAccessErrors.daeCannotLockBatchInUseBy, (mvClassFields.Item(BatchFields.BatchNumber).Value), vUser)
        End If
      Else
        'Open batch record not found so add it and flag as in_use_by the user
        vInsertFields = New CDBFields
        vInsertFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
        vInsertFields.Add("in_use_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
        vInsertFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
        vInsertFields.Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
        mvEnv.Connection.InsertRecord("open_batches", vInsertFields)
      End If
      vRecordSet.CloseRecordSet()
      mvBatchLocked = True
    End Sub

    Public Sub UnLockBatch()
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As CDBFields
      Dim vUpdateFields As CDBFields
      Dim vDeleteOpenBatch As Boolean

      'Unlock the batch by deleting the record in the open batches table if not allowed to add transactions
      vWhereFields = New CDBFields
      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchFields.BatchNumber).IntegerValue)
      vWhereFields.Add("in_use_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)

      If mvBatchLocked And Len(mvOpenBatchType) = 0 Then
        vDeleteOpenBatch = True
      Else
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT in_use_by FROM open_batches WHERE batch_number = " & mvClassFields.Item(BatchFields.BatchNumber).IntegerValue & " AND batch_type IS NOT NULL")
        If vRecordSet.Fetch() = True Then
          'If we found a batch with a non-null batch type then it is a real open batch so update it
          vUpdateFields = New CDBFields
          vUpdateFields.Add("in_use_by", CDBField.FieldTypes.cftCharacter)
          mvEnv.Connection.UpdateRecords("open_batches", vUpdateFields, vWhereFields)
        Else
          vDeleteOpenBatch = True
        End If
        vRecordSet.CloseRecordSet()
      End If
      If vDeleteOpenBatch Then
        vWhereFields.Add("batch_type")
        vWhereFields.Add("bank_account")
        mvEnv.Connection.DeleteRecords("open_batches", vWhereFields, False)
      End If
      mvBatchLocked = False
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vError As Integer = Delete(0, True)
      If vError > 0 Then RaiseError(CType(vError, DataAccessErrors))
    End Sub

    Public Overloads Function Delete(ByVal pTransactionNumber As Integer) As Integer
      Dim vError As Integer = Delete(pTransactionNumber, True)
      If vError > 0 Then RaiseError(CType(vError, DataAccessErrors))
    End Function

    Public Overloads Function Delete(ByVal pTransactionNumber As Integer, ByVal pCheckDataFirst As Boolean) As Integer
      'When performing the data checks, any errors will be returned instead of actually raising the error (BR10180)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCount As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vCancel As Boolean
      Dim vCSSundryCreditProduct As String = ""
      Dim vCTWhereFields As CDBFields
      Dim vStockMovement As StockMovement
      Dim vSQL As String
      Dim vDeleteProvisionalTrans As Boolean
      Dim vStartTransaction As Boolean
      Dim vCMDWhereFields As New CDBFields
      Dim vOPH As New OrderPaymentHistory
      Dim vCardSale As New CardSale(mvEnv)
      Dim vCCA As New CreditCardAuthorisation
      Dim vTransactionNumber As Integer
      Dim vLineNumber As Integer
      Dim vProductCode As String
      Dim vWarehouse As String
      Dim vKey As String
      Dim vWarehouseProductMovements As New CDBParameters
      Dim vIndex As Integer
      Dim vMovementQty As Integer
      Dim vAdjustments As Boolean
      Dim vErrorNumber As Integer
      Dim vProductCosts As ProductCosts
      Dim vProductCostNumber As Integer
      Dim vBatchTransaction As BatchTransaction = Nothing

      'If transaction number is zero then delete all transactions
      LockBatch() 'Must Unlock if error
      Try


        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
        If pTransactionNumber > 0 Then vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)

        If pTransactionNumber > 0 And Provisional Then
          'Transaction can be deleted if it has not been confirmed
          vCTWhereFields = New CDBFields
          With vCTWhereFields
            .Add("provisional_batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            .Add("provisional_trans_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
            .Add("confirmed_batch_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
          End With
          If mvEnv.Connection.GetCount("confirmed_transactions", vCTWhereFields) = 0 Then vDeleteProvisionalTrans = True
        End If

        If pCheckDataFirst Then
          If pTransactionNumber = 0 Then
            'Check to see if the Batch contains adjustments
            Select Case BatchType
              Case BatchTypes.FinancialAdjustment
                vAdjustments = True
              Case BatchTypes.CreditCard, BatchTypes.CreditCardWithInvoice, BatchTypes.CreditSales, BatchTypes.DebitCard, BatchTypes.DirectCredit, BatchTypes.GiveAsYouEarn, BatchTypes.PostTaxPayrollGiving
                If mvEnv.Connection.GetCount("reversals", vWhereFields) > 0 Then vAdjustments = True
              Case Else
                vAdjustments = False
            End Select

            If vAdjustments Then
              'Deletion of a Financial Adjustment batch is not allowed
              vCancel = True
              vErrorNumber = DataAccessErrors.daeCannotDeleteAdjustmentBatches
            End If
          End If

          If vCancel = False Then
            If pTransactionNumber = 0 Then
              If PostedToCashBook Then
                'if all of the batch header values are non-0, disallow delete
                If CDbl(mvClassFields.Item(BatchFields.NumberOfEntries).Value) > 0 And CDbl(mvClassFields.Item(BatchFields.BatchTotal).Value) > 0 And CDbl(mvClassFields.Item(BatchFields.TransactionTotal).Value) > 0 And CDbl(mvClassFields.Item(BatchFields.NumberOfTransactions).Value) > 0 Then
                  vCancel = True
                  vErrorNumber = DataAccessErrors.daeBatchPostedToCashBook
                End If
              End If
              If mvEnv.User.AccessLevel <> CDBUser.UserAccessLevel.ualSupervisor And mvEnv.User.AccessLevel <> CDBUser.UserAccessLevel.ualDatabaseAdministrator Then RaiseError(DataAccessErrors.daeAccessLevel)
            Else
              If PostedToCashBook And BatchType <> BatchTypes.CreditSales And BatchType <> BatchTypes.BankStatement And Not (Provisional = True And vDeleteProvisionalTrans = True) Then
                If Not mvEnv.GetConfigOption("cb_delete_transactions", False) Then
                  vCancel = True
                  vErrorNumber = DataAccessErrors.daeBatchPostedToCashBook
                End If
              End If
            End If
          End If

          If vCancel = False And Picked <> "N" Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeBatchPicked
          End If

          If vCancel = False And (PostedToNominal And Not (Provisional = True And vDeleteProvisionalTrans = True)) Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeBatchPostedToNominal
          End If

          If vCancel = False And mvEnv.Connection.GetCount("event_bookings", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransEventBooking
          End If

          If vCancel = False And mvEnv.Connection.GetCount("exam_booking_units", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransExamBooking
          End If

          If vCancel = False And mvEnv.Connection.GetCount("exam_booking_transactions", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransExamBooking
          End If

          If vCancel = False And mvEnv.Connection.GetCount("contact_room_bookings", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransAccommBooking
          End If

          If vCancel = False And mvEnv.Connection.GetCount("service_bookings", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransServiceBooking
          End If

          If vCancel = False And mvEnv.Connection.GetCount("legacy_bequest_receipts", vWhereFields) > 0 Then
            vCancel = True
            vErrorNumber = DataAccessErrors.daeTransLegacyReceipt
          End If

          If vCancel = False Then
            vWhereFields.Add("line_type", CDBField.FieldTypes.cftCharacter, "'N','L','U'", CDBField.FieldWhereOperators.fwoIn)
            If mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) > 0 Then
              vCancel = True
              vErrorNumber = DataAccessErrors.daeTransInvoicePayments
            End If
            vWhereFields.Remove((vWhereFields.Count))
          End If

          If vCancel = False And BatchType = BatchTypes.CreditSales Then
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT sundry_credit_product FROM bank_accounts ba, company_controls cc WHERE bank_account = '" & BankAccount & "' AND ba.company = cc.company")
            If vRecordSet.Fetch() = True Then
              vCSSundryCreditProduct = vRecordSet.Fields(1).Value
            End If
            vRecordSet.CloseRecordSet()
            If vCSSundryCreditProduct.Length > 0 Then
              vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, vCSSundryCreditProduct)
              If mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) > 0 Then
                vCancel = True
                vErrorNumber = DataAccessErrors.daeTransCreditNotes
              End If
              vWhereFields.Remove((vWhereFields.Count))
            End If
          End If

          If vCancel = False And (BatchType = BatchTypes.CreditSales And mvEnv.GetConfigOption("fp_use_sales_ledger", True)) Then
            vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
            If mvEnv.Connection.GetCount("invoices", vWhereFields) > 0 Then
              vCancel = True
              vErrorNumber = DataAccessErrors.daeTransInvoicePrinted
            End If
            vWhereFields.Remove((vWhereFields.Count))
          End If
        End If

        If pTransactionNumber > 0 Then
          If Me.HasInvoices(pTransactionNumber) Then
            vErrorNumber = DataAccessErrors.daeCannotDeleteSalesLedgerTransaction
            vCancel = True
          End If
        Else
          If Me.HasInvoices Then
            vErrorNumber = DataAccessErrors.daeCannotDeleteSalesLedgerBatch
            vCancel = True
          End If
        End If

        If Not vCancel Then
          If pTransactionNumber > 0 Then
            vBatchTransaction = New BatchTransaction(mvEnv)
            vBatchTransaction.Init(BatchNumber, pTransactionNumber, True)
            If BatchType = BatchTypes.CreditCard OrElse BatchType = BatchTypes.DebitCard OrElse BatchType = BatchTypes.CreditCardWithInvoice Then
              vCardSale.Init(BatchNumber, pTransactionNumber)
              If vCardSale.NoClaimRequired Then
                'Already authorised this payment so we need to reverse it
                vCCA.InitFromTransaction(mvEnv, BatchNumber, pTransactionNumber)
                If vCCA.Existing Then
                  vCCA.CheckOnlineAuthorisation() 'Check server is running to avoid delay
                  mvCCA = New CreditCardAuthorisation
                  mvCCA.Init(mvEnv)
                  mvCCA.ContactNumber = vBatchTransaction.ContactNumber
                  If Not mvCCA.AuthoriseTransaction(vCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatRefund, (vCCA.AuthorisedAmount), vBatchTransaction.AddressNumber, vCCA.AuthorisedTransactionNo, vCCA.AuthorisedTextId, vCCA.AuthorisationCode, vCCA.AuthorisationNumber) Then
                    RaiseError(DataAccessErrors.daeCCAuthorisationFailed, mvCCA.AuthorisationResponseMessage)
                  Else
                    'Set the no claim required to false so that if the following delete fails
                    'and the batch gets processed then a claim will take place
                    vCardSale.NoClaimRequired = False
                  End If
                End If
              End If
            End If
          End If

          If Not mvEnv.Connection.InTransaction Then
            vStartTransaction = True
            mvEnv.Connection.StartTransaction()
          End If
          If pTransactionNumber > 0 Then
            NumberOfTransactions = NumberOfTransactions - 1
            mvClassFields.Item(BatchFields.TransactionTotal).Value = CStr(TransactionTotal - If(vBatchTransaction.TransactionSign = "C", vBatchTransaction.Amount, (vBatchTransaction.Amount * -1)))
            If mvClassFields.Item(BatchFields.CurrencyCode).InDatabase Then
              If CurrencyCode.Length > 0 Then
                mvClassFields.Item(BatchFields.CurrencyTransactionTotal).Value = CStr(CurrencyTransactionTotal - If(vBatchTransaction.TransactionSign = "C", vBatchTransaction.CurrencyAmount, (vBatchTransaction.CurrencyAmount * -1)))
              End If
            End If
            If Not PostedToCashBook And BatchType <> BatchTypes.CreditCard And BatchType <> BatchTypes.DebitCard And BatchType <> BatchTypes.CreditCardWithInvoice And BatchType <> BatchTypes.BankStatement And BatchType <> BatchTypes.StandingOrder Then
              ReadyForBanking = False
            End If
            If CurrencyIndicator.Length > 0 Then mvClassFields.Item(BatchFields.CurrencyTransactionTotal).Value = CStr(CurrencyTransactionTotal - If(vBatchTransaction.TransactionSign = "C", vBatchTransaction.CurrencyAmount, (vBatchTransaction.CurrencyAmount * -1)))
            Save()
          End If

          'Remove any contact mailing documents that relate only to this batch/transaction
          'If the documents relate to anything else or they have been fulfilled then leave them behind
          'No need to clear the Batch/Transaction nos as the report will ignore the transaction data - I hope...
          vCMDWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          If pTransactionNumber > 0 Then vCMDWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
          vCMDWhereFields.Add("order_number", CDBField.FieldTypes.cftLong)
          vCMDWhereFields.Add("gaye_pledge_number", CDBField.FieldTypes.cftLong)
          vCMDWhereFields.Add("declaration_number", CDBField.FieldTypes.cftLong)
          vCMDWhereFields.Add("fulfillment_number", CDBField.FieldTypes.cftLong)
          vCMDWhereFields.Add("new_contact", CDBField.FieldTypes.cftCharacter, "N")
          mvEnv.Connection.DeleteRecords("contact_mailing_documents", vCMDWhereFields, False)

          'Reverse any stock movements
          If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockInterface) = "Y" Or mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCoreStockControl) = "Y" Then
            'vSQL = "SELECT bt.transaction_number, bta.line_number, bta.quantity, p.product FROM batch_transactions bt, batch_transaction_analysis bta, products p, stock_movements sm WHERE bt.batch_number = " & BatchNumber
            vSQL = "SELECT bt.transaction_number, bta.line_number, sm.movement_quantity,sm.warehouse AS  sm_warehouse, p.product"
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then vSQL = vSQL & ", sm.product_cost_number"
            vSQL = vSQL & " FROM batch_transactions bt, batch_transaction_analysis bta, products p, stock_movements sm"
            vSQL = vSQL & " WHERE bt.batch_number = " & BatchNumber
            If pTransactionNumber > 0 Then vSQL = vSQL & " AND bt.transaction_number = " & pTransactionNumber
            vSQL = vSQL & " AND bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number AND bta.product IS NOT NULL"
            vSQL = vSQL & " AND p.product = bta.product AND p.stock_item = 'Y' AND sm.batch_number = bta.batch_number AND sm.transaction_number = bta.transaction_number"
            vSQL = vSQL & " AND sm.line_number = bta.line_number AND sm.product = bta.product"
            vSQL = vSQL & " ORDER BY bta.transaction_number,bta.line_number,bta.product,sm.warehouse"
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            While vRecordSet.Fetch() = True
              vTransactionNumber = CInt(vRecordSet.Fields("transaction_number").Value)
              vLineNumber = CInt(vRecordSet.Fields("line_number").Value)
              vProductCode = vRecordSet.Fields("product").Value
              vWarehouse = vRecordSet.Fields("sm_warehouse").Value
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then vProductCostNumber = vRecordSet.Fields("product_cost_number").IntegerValue
              vKey = vTransactionNumber & Space(8 - Len(CStr(vTransactionNumber))) & vLineNumber & Space(8 - Len(CStr(vLineNumber))) & vWarehouse & Space(2 - Len(vWarehouse)) & vProductCode & Space(20 - Len(vProductCode)) & vProductCostNumber
              If vWarehouseProductMovements.Exists(vKey) Then
                vWarehouseProductMovements(vKey).Value = (vWarehouseProductMovements(vKey).IntegerValue + vRecordSet.Fields("movement_quantity").IntegerValue).ToString
              Else
                vWarehouseProductMovements.Add(vKey, CDBField.FieldTypes.cftLong, vRecordSet.Fields("movement_quantity").Value)
              End If
            End While
            vRecordSet.CloseRecordSet()
            For vIndex = 1 To vWarehouseProductMovements.Count
              If vWarehouseProductMovements(1).IntegerValue <> 0 Then
                'Create a balancing Stock Movement to reverse final activity
                vTransactionNumber = IntegerValue(Mid(vWarehouseProductMovements(vIndex).Name, 1, 8))
                vLineNumber = IntegerValue(Mid(vWarehouseProductMovements(vIndex).Name, 9, 8))
                vWarehouse = Trim(Mid(vWarehouseProductMovements(vIndex).Name, 17, 2))
                vProductCode = Trim(Mid(vWarehouseProductMovements(vIndex).Name, 19, 20))
                vProductCostNumber = IntegerValue(Mid(vWarehouseProductMovements(vIndex).Name, 39))
                vMovementQty = vWarehouseProductMovements(vIndex).IntegerValue * -1

                If vProductCostNumber = 0 Then
                  If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then
                    'The original payment was created before the upgrade so re-allocate against the earliest ProductCost record
                    vProductCosts = New ProductCosts
                    vProductCosts.InitFromProductAndWarehouse(mvEnv, vProductCode, vWarehouse, False)
                    vProductCostNumber = vProductCosts.GetEarliestProductCost.ProductCostNumber
                  End If
                End If

                vStockMovement = New StockMovement
                vStockMovement.Create(mvEnv, vProductCode, vMovementQty, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonReversal), BatchNumber, vTransactionNumber, vLineNumber, False, vWarehouse, vProductCostNumber)
              End If
            Next
          End If

          If BatchType = BatchTypes.CreditSales Then
            If mvEnv.GetConfigOption("fp_use_sales_ledger", True) Then
              mvEnv.Connection.DeleteRecords("invoice_details", vWhereFields, False)
              mvEnv.Connection.DeleteRecords("invoices", vWhereFields, False)
              'Update the CreditCustomer (could be OnOrder or Outstanding that needs updating)
              'Build nested SQL first
              Dim vNestedWhereFields As New CDBFields(New CDBField("batch_number", BatchNumber))
              With vNestedWhereFields
                If pTransactionNumber > 0 Then .Add("transaction_number", pTransactionNumber)
                .Add("line_type", CDBField.FieldTypes.cftCharacter, "'O','M','C'", CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
                .Add("p.product", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
                .Add("p.product#2", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
                .Add("stock_item", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
              End With
              Dim vNestedAnsiJoins As New AnsiJoins
              vNestedAnsiJoins.AddLeftOuterJoin("products p", "bta.product", "p.product")
              Dim vNestedSQLStatement As New SQLStatement(mvEnv.Connection, "batch_number, transaction_number, COUNT(*) AS pp_or_stock_count", "batch_transaction_analysis bta", vNestedWhereFields, "", vNestedAnsiJoins)
              vNestedSQLStatement.GroupBy = "batch_number, transaction_number"
              'Main SQL
              Dim vAttrs As String = "ccu.company, ccu.contact_number, ccu.sales_ledger_account, transaction_sign, bt.amount AS bt_amount,"
              vAttrs &= mvEnv.Connection.DBIsNull("bta.pp_or_stock_count", "0") & "AS pp_or_stock_count, bt.transaction_number"
              Dim vSLWhereFields As New CDBFields(New CDBField("cs.batch_number", BatchNumber))
              If pTransactionNumber > 0 Then vSLWhereFields.Add("cs.transaction_number", pTransactionNumber)
              Dim vSLAnsiJoins As New AnsiJoins()
              With vSLAnsiJoins
                .Add("batches b", "cs.batch_number", "b.batch_number")
                .Add("batch_transactions bt", "cs.batch_number", "bt.batch_number", "cs.transaction_number", "bt.transaction_number")
                .Add("bank_accounts ba", "b.bank_account", "ba.bank_account")
                .Add("credit_customers ccu", "cs.contact_number", "ccu.contact_number", "cs.sales_ledger_account", "ccu.sales_ledger_account", "ba.company", "ccu.company")
                .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
                .AddLeftOuterJoin("(" & vNestedSQLStatement.SQL & ") bta", "bt.batch_number", "bta.batch_number", "bt.transaction_number", "bta.transaction_number")
              End With
              Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "credit_sales cs", vSLWhereFields, "bt.transaction_number, ccu.contact_number, ccu.company, ccu.sales_ledger_account", vSLAnsiJoins)
              vSQLStatement.GroupBy = "bt.batch_number, bt.transaction_number, ccu.company, ccu.contact_number, ccu.sales_ledger_account, transaction_sign, bt.amount, pp_or_stock_count"
              vRecordSet = vSQLStatement.GetRecordSet()

              Dim vUpdateAmount As Double
              Dim vCreditCustomer As CreditCustomer
              While vRecordSet.Fetch
                With vRecordSet
                  vCreditCustomer = New CreditCustomer
                  vCreditCustomer.Init(mvEnv, .Fields("contact_number").IntegerValue, .Fields("company").Value, .Fields("sales_ledger_account").Value)  'Always re-initialise to get latest values
                  If vCreditCustomer.Existing Then
                    vUpdateAmount = .Fields("bt_amount").DoubleValue
                    If .Fields("pp_or_stock_count").IntegerValue > 0 Then
                      If .Fields("transaction_sign").Value <> "C" Then vUpdateAmount = vUpdateAmount * -1
                      vCreditCustomer.AdjustOnOrder(vUpdateAmount)    'Reduce OnOrder
                    Else
                      If .Fields("transaction_sign").Value = "C" Then vUpdateAmount = vUpdateAmount * -1
                      vCreditCustomer.AdjustOutstanding(vUpdateAmount)  'Increase Outstanding
                    End If
                    vCreditCustomer.Save(mvEnv.User.UserID, True)
                  End If
                End With
              End While
              vRecordSet.CloseRecordSet()
            End If
            mvEnv.Connection.DeleteRecords("credit_sales", vWhereFields, False)
          End If

          'Remove any order payment history records (starting with the last one)
          'Update the payment schedule to show that there is no longer an unprocessed payment
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then
            vOPH.Init(mvEnv)
            vOPH.DeleteFromBatch(vWhereFields, BatchType)
          End If
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataLinkToCommunication) Then
            mvEnv.Connection.DeleteRecords("communications_log_trans", vWhereFields, False)
          End If
          mvEnv.Connection.DeleteRecords("batch_transactions", vWhereFields, False)
          mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields, False)

          If BatchType = BatchTypes.CreditCard OrElse BatchType = BatchTypes.DebitCard OrElse BatchType = BatchTypes.CreditCardWithInvoice Then mvEnv.Connection.DeleteRecords("card_sales", vWhereFields, False)
          If BatchType = BatchTypes.CreditSales Or BatchType = BatchTypes.FinancialAdjustment Then mvEnv.Connection.DeleteRecords("reversals", vWhereFields, False)
          If mvClassFields.Item(BatchFields.Provisional).InDatabase And Provisional Then
            vCTWhereFields = New CDBFields
            vCTWhereFields.Add("provisional_batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            If pTransactionNumber > 0 Then
              vCTWhereFields.Add("provisional_trans_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
            Else
              vCTWhereFields.Add("confirmed_batch_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
              vCount = mvEnv.Connection.GetCount("confirmed_transactions", vCTWhereFields)
              If vCount > 0 And vCount <> NumberOfTransactions Then RaiseError(DataAccessErrors.daeBatchContainsConfirmedTrans)
              vCTWhereFields.Remove((vCTWhereFields.Count))
            End If
            mvEnv.Connection.DeleteRecords("confirmed_transactions", vCTWhereFields, False)
          End If

          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
            mvEnv.Connection.DeleteRecords("collection_payments", vWhereFields, False)
          End If

          If pTransactionNumber = 0 Then
            mvEnv.Connection.DeleteRecords("batches", vWhereFields, False)
          End If
          If vStartTransaction Then mvEnv.Connection.CommitTransaction()
        End If
        UnLockBatch()

        'If the data checks have failed then return the error number
        If (vCancel = True And vErrorNumber <> 0) Then Delete = vErrorNumber

      Catch vEx As Exception
        PreserveStackTrace(vEx)
        UnLockBatch()
        Throw vEx
      End Try
    End Function
    Public Sub SetBatchPosted(ByRef pSuccessful As Boolean, Optional ByRef pPostedOn As String = "", Optional ByRef pPostedBy As String = "")

      If Len(pPostedBy) = 0 Then pPostedBy = mvEnv.User.UserID
      If Not IsDate(pPostedOn) Then pPostedOn = TodaysDate()
      mvClassFields.Item(BatchFields.JobNumber).Value = "" 'Clear the job number
      If pSuccessful Then
        mvClassFields.Item(BatchFields.PostedToNominal).Bool = True
        mvClassFields.Item(BatchFields.PostedBy).Value = pPostedBy
        mvClassFields.Item(BatchFields.PostedOn).Value = pPostedOn
      End If
    End Sub

    '-----------------------------------------------------------------------------
    ' BATCH PROCESSING FUNCTIONS BEYOND HERE
    '-----------------------------------------------------------------------------
    Private Sub CheckDDIncentiveDate(ByRef pPP As PaymentPlan)
      'This script will be called from 'update_next_payment_due' when the
      'balance is the same as the renewal amount and order_term is negative.
      'Hence there has been a reversal against a membership that involved an initial
      'free period. Rather than re-set the next_payment_due to the renewal_date,
      'check for a direct debit record set up within the last year and if it exists,
      'set the next_payment_due on the orders record to the direct debit start date.
      '(This is the logic employed in ME_Fast_Member_Entry)
      If pPP.DirectDebitStatus <> PaymentPlan.ppYesNoCancel.ppCancelled AndAlso CDate(pPP.DirectDebit.StartDate) >= Today.AddYears(-1) Then
        pPP.NextPaymentDue = pPP.DirectDebit.StartDate
      End If
    End Sub
    Private Function CheckInAdvance(ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, Optional ByRef pLineNumber As Integer = 0) As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      If mvCompanyControl.InAdvanceProductCode.Length > 0 Then
        vSQL = "SELECT r.batch_number FROM reversals r, order_payment_history oph, financial_history_details fhd WHERE "
        vSQL = vSQL & "r.batch_number = " & pBatchNumber & " AND r.transaction_number = " & pTransactionNumber
        If pLineNumber > 0 Then vSQL = vSQL & " AND r.line_number = " & pLineNumber
        vSQL = vSQL & " AND oph.batch_number = r.was_batch_number AND oph.transaction_number = r.was_transaction_number AND oph.line_number = r.was_line_number"
        vSQL = vSQL & " AND fhd.batch_number = r.was_batch_number AND fhd.transaction_number = r.was_transaction_number AND fhd.line_number = r.was_line_number"
        vSQL = vSQL & " AND fhd.product = '" & mvCompanyControl.InAdvanceProductCode & "' AND fhd.rate = '" & mvCompanyControl.InAdvanceRate & "'"
        vRecordSet = mvConn.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then CheckInAdvance = True
        vRecordSet.CloseRecordSet()
      End If
    End Function
    Private Function CheckProcessSubscriptions(ByRef pPP As PaymentPlan, ByRef pPPD As PaymentPlanDetail) As String
      Dim vWhereFields As New CDBFields
      'This will be called from ProcessOrderDetailsPayment and ProcessOrderDetailsReversal before
      'the call to ProcessSubscriptions to check for future subs creation
      'Return values
      ' Y: proceed with process_subscriptions
      ' N: do nothing
      ' T: terminate subscription
      ' D: delete subscription
      Dim vProcessSubscriptions As String = "Y"

      Select Case mvMeFutureChange
        Case "N"
          vProcessSubscriptions = "Y"
        Case "Y"
          'a future change is required, see if we are processing a future order_details line
          If pPPD.TimeStatus = "F" Then
            vProcessSubscriptions = "Y"
          Else
            'if there is future line for the same product: do nothing
            'subs will get processed for that line if not, terminate subs
            vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
            vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pPPD.ProductCode)
            vWhereFields.Add("time_status", CDBField.FieldTypes.cftCharacter, "F")

            If mvEnv.Connection.GetCount("order_details", vWhereFields) > 0 Then
              vProcessSubscriptions = "N"
            Else
              vProcessSubscriptions = "T"
            End If
          End If

        Case "R"
          'a future change reversal is required, see if we are processing a future order_details line
          If pPPD.TimeStatus = "F" Then
            'if there is current line for the same product: do nothing
            'subs will get processed for that line if not, delete subs
            vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
            vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pPPD.ProductCode)
            vWhereFields.Add("time_status", CDBField.FieldTypes.cftCharacter, "C")
            If mvEnv.Connection.GetCount("order_details", vWhereFields) > 0 Then
              vProcessSubscriptions = "N"
            Else
              vProcessSubscriptions = "D"
            End If
          Else
            vProcessSubscriptions = "Y" 'time_status is C, process as normal
          End If
      End Select
      Return vProcessSubscriptions
    End Function

    Private Sub CreatedDeclarationLinesUnclaimed(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pNumber As Integer, ByRef pContactNumber As Integer, ByRef pDorC As String, ByRef pAmount As Double)
      Dim vInsertFields As New CDBFields
      Dim vCovenant As New Covenant
      Dim vCreateDLU As Boolean

      If ((pAmount > 0) OrElse (pAmount < 0 AndAlso mvImportPayment = True)) Then 'Do create reversal lines except for import
        vCreateDLU = True
        If pDorC = "C" Then
          vCovenant.Init(mvEnv, pNumber)
          If CDate(pBT.TransactionDate) > DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(mvEnv.GetConfig("cv_no_days_claim_grace")), vCovenant.EndDate) Then
            vCreateDLU = False
          End If
        End If
        If vCreateDLU Then
          With vInsertFields
            .Add("cd_number", CDBField.FieldTypes.cftLong, pNumber)
            .Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
            .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
            .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
            .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
            .Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, pDorC)
            .Add("net_amount", CDBField.FieldTypes.cftNumeric, pAmount)
            mvConn.InsertRecord("declaration_lines_unclaimed", vInsertFields)
          End With
        End If
      End If
    End Sub
    Private Sub DeleteContactSuppressions(ByRef pContactNumber As Integer, ByRef pMailingSuppression As String)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      vWhereFields.Add("mailing_suppression", CDBField.FieldTypes.cftCharacter, pMailingSuppression)
      mvConn.DeleteRecords("contact_suppressions", vWhereFields, False)
    End Sub
    Private Sub DeleteSubscriptions(ByRef pPPNumber As Integer, ByRef pContactNumber As Integer, ByRef pProduct As String)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPPNumber)
      vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pProduct)
      If pContactNumber > 0 Then
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      Else
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong)
      End If
      mvConn.DeleteRecords("subscriptions", vWhereFields)
    End Sub
    Private Function GetOriginalWriteOff(ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, ByRef pLineNumber As Integer, ByRef pBTAAcceptAsFull As Boolean) As Double
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vOPH As New OrderPaymentHistory
      vOPH.Init(mvEnv)
      vSQL = "SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM reversals r, order_payment_history oph WHERE"
      vSQL = vSQL & " r.batch_number = " & pBatchNumber & " AND r.transaction_number = " & pTransactionNumber & " AND r.line_number = " & pLineNumber
      vSQL = vSQL & " AND oph.batch_number = r.was_batch_number AND oph.transaction_number = r.was_transaction_number AND oph.line_number = r.was_line_number"
      vRecordSet = mvConn.GetRecordSet(vSQL)
      If vRecordSet.Fetch() = True Then
        If pBTAAcceptAsFull Then
          Return vRecordSet.Fields("balance").DoubleValue
        Else
          'Check OPH.WriteOffLineAmount
          vOPH.InitFromRecordSet(mvEnv, vRecordSet, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          If vOPH.Existing Then Return vOPH.WriteOffLineAmount
        End If
      End If
      vRecordSet.CloseRecordSet()
      Return 0
    End Function
    Private Function ReverseFinancialDetails(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByVal pPaymentAmount As Double) As Double
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vFHD As New FinancialHistoryDetail(mvEnv)
      Dim vFL As FinancialLink

      vFHD.Init(mvEnv)
      vSQL = "SELECT " & vFHD.GetRecordSetFields(FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll) & " FROM reversals r, financial_history_details fhd WHERE "
      vSQL = vSQL & "r.batch_number = " & pBTA.BatchNumber & " AND r.transaction_number = " & pBTA.TransactionNumber & " AND r.line_number = " & pBTA.LineNumber
      vSQL = vSQL & " AND fhd.batch_number = r.was_batch_number AND fhd.transaction_number = r.was_transaction_number AND fhd.line_number = r.was_line_number"
      vRecordSet = mvConn.GetRecordSet(vSQL)
      While vRecordSet.Fetch() = True
        vFHD.InitFromRecordSet(mvEnv, vRecordSet, FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll)
        If (vFHD.ProductCode = mvCompanyControl.OverPaymentProductCode And vFHD.RateCode = mvCompanyControl.OverPaymentRate) And (pBTA.Amount = pPaymentAmount) Then
          'Do Nothing as Re-analysis has processed this separately
        Else
          vFHD.BatchNumber = pBTA.BatchNumber
          vFHD.TransactionNumber = pBTA.TransactionNumber
          vFHD.LineNumber = pBTA.LineNumber
          'vFHD.SalesContactNumber =
          vFHD.Reverse()
          vFHD.Status = FinancialHistory.FinancialHistoryStatus.fhsNormal
          vFHD.Save()
          If pBTA.LineType = "S" Then
            'create the reverse financial link
            vFL = New FinancialLink
            vFL.InitFromValues(mvEnv, (pBTA.DeceasedContactNumber), (pBT.ContactNumber), (pBTA.BatchNumber), (pBTA.TransactionNumber), (pBTA.LineNumber), (pBTA.LineType))
            vFL.Save()
          End If
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Function
    Private Function ReversePaymentPlanHistoryDetails(ByRef pPP As PaymentPlan, ByRef pBTA As BatchTransactionAnalysis, ByVal pPaymentAmount As Double) As Double
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("order_payment_history oph", "r.was_batch_number", "oph.batch_number", "r.was_transaction_number", "oph.transaction_number", "r.was_line_number", "oph.line_number")
        vAnsiJoins.Add("payment_plan_history_details pphd", "oph.order_number", "pphd.order_number", "oph.payment_number", "pphd.payment_number")

        Dim vWherefields As New CDBFields()
        vWherefields.Add("r.batch_number", CDBField.FieldTypes.cftInteger, pBTA.BatchNumber)
        vWherefields.Add("r.transaction_number", CDBField.FieldTypes.cftInteger, pBTA.TransactionNumber)
        vWherefields.Add("r.line_number", CDBField.FieldTypes.cftInteger, pBTA.LineNumber)

        Dim vPPHD As New PaymentPlanHistoryDetail(mvEnv)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vPPHD.GetRecordSetFields().Replace("pphd.payment_number", pBTA.OrderPaymentHistory.PaymentNumber.ToString).Replace("payment_amount", "payment_amount *-1 As payment_amount"), "reversals r", vWherefields, "", vAnsiJoins)
        Dim vSQL As New StringBuilder
        vSQL.Append("INSERT INTO payment_plan_history_details (" & vPPHD.GetRecordSetFields.Replace("pphd.", "") & ")")
        vSQL.Append(vSQLStatement.SQL)
        mvEnv.Connection.ExecuteSQL(vSQL.ToString)
      End If
    End Function
    Public Function IntegrityCheck(ByVal pConn As CDBConnection, ByRef pErrorList As String, ByRef pErrorMsg As String) As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vErrors As String = ""
      Dim vErrorCount As Integer

      'Check the contact_number exists in contacts table for the each transaction in batch_transaction
      vRecordSet = pConn.GetRecordSetAnsiJoins("SELECT transaction_number FROM batch_transactions bt LEFT OUTER JOIN contacts c ON bt.contact_number = c.contact_number WHERE batch_number = " & BatchNumber & " AND transaction_type IN ('P','I') AND c.contact_number IS NULL ORDER BY transaction_number")
      While vRecordSet.Fetch() = True
        If vErrorCount > 0 Then vErrors = vErrors & ", "
        vErrors = vErrors & vRecordSet.Fields(1).Value 'Only display the Transaction Number
        vErrorCount = vErrorCount + 1
      End While
      vRecordSet.CloseRecordSet()
      If vErrorCount = 0 Then
        'Check the contact_number exists in contacts table for each transaction in batch_transaction_analysis that has a contact_number
        vRecordSet = pConn.GetRecordSetAnsiJoins("SELECT transaction_number,line_number FROM batch_transaction_analysis bta LEFT OUTER JOIN contacts c ON bta.contact_number = c.contact_number WHERE batch_number = " & BatchNumber & " AND line_type IN ('P','I') AND bta.contact_number IS NOT NULL AND c.contact_number IS NULL ORDER BY transaction_number, line_number")
        While vRecordSet.Fetch() = True
          If vErrorCount > 0 Then vErrors = vErrors & ", "
          vErrors = vErrors & vRecordSet.Fields(1).Value & "/" & vRecordSet.Fields(2).Value
          vErrorCount = vErrorCount + 1
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vErrorCount = 0 Then
        'Check that Order Number is set for Line Type O,M  and that product is set for Line type P,G,B,I
        vRecordSet = pConn.GetRecordSet("SELECT transaction_number,line_number FROM batch_transaction_analysis WHERE batch_number = " & BatchNumber & " AND ((line_type IN ('O','M') AND order_number IS NULL) OR (line_type IN ('P','G','B','I') AND product IS NULL)) ORDER BY transaction_number, line_number")
        While vRecordSet.Fetch() = True
          If vErrorCount > 0 Then vErrors = vErrors & ", "
          vErrors = vErrors & vRecordSet.Fields(1).Value & "/" & vRecordSet.Fields(2).Value
          vErrorCount = vErrorCount + 1
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vErrorCount = 0 Then
        'Check that stock sales Line Type P,I have contact and address number set
        vRecordSet = pConn.GetRecordSet("SELECT transaction_number,line_number FROM batch_transaction_analysis bta, products p WHERE batch_number = " & BatchNumber & " AND line_type IN ('P','I') AND bta.product = p.product AND p.stock_item = 'Y' AND (bta.contact_number IS NULL or bta.address_number IS NULL) ORDER BY transaction_number, line_number")
        While vRecordSet.Fetch() = True
          If vErrorCount > 0 Then vErrors = vErrors & ", "
          vErrors = vErrors & vRecordSet.Fields(1).Value & "/" & vRecordSet.Fields(2).Value
          vErrorCount = vErrorCount + 1
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vErrorCount = 0 Then
        'Check the analysis is not an order payment of amount zero (ignore I type lines)
        vRecordSet = pConn.GetRecordSet("SELECT transaction_number,line_number FROM batch_transaction_analysis bta WHERE batch_number = " & BatchNumber & " AND order_number IS NOT NULL AND amount = 0 AND line_type <> 'I' ORDER BY transaction_number, line_number")
        While vRecordSet.Fetch() = True
          If vErrorCount > 0 Then vErrors = vErrors & ", "
          vErrors = vErrors & vRecordSet.Fields(1).Value & "/" & vRecordSet.Fields(2).Value
          vErrorCount = vErrorCount + 1
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vErrorCount = 0 Then
        'Check for batch transactions with no analysis
        vRecordSet = pConn.GetRecordSet("SELECT transaction_number FROM batch_transactions WHERE batch_number = " & BatchNumber & " AND transaction_number NOT IN (SELECT DISTINCT transaction_number FROM batch_transaction_analysis WHERE batch_number = " & BatchNumber & ")")
        While vRecordSet.Fetch() = True
          If vErrorCount > 0 Then vErrors = vErrors & ", "
          vErrors = vErrors & vRecordSet.Fields(1).Value
          vErrorCount = vErrorCount + 1
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vErrorCount = 0 Then
        If UsesHoldingContact Then
          vErrors = "Transactions are using the Holding Contact"
          vErrorCount = vErrorCount + 1
        End If
      End If
      If vErrorCount = 0 And Len(BatchCategory) = 0 And UCase(Left(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAccountsInterface), 4)) = "RCPC" Then
        vErrors = "Batch Category has not been set"
        vErrorCount = vErrorCount + 1
      End If

      If vErrorCount = 0 AndAlso Picked = "N" Then
        'Check the Batch for stock products that are in stock if the batch has not been picked & confirmed
        Dim vAnsiJoins As New AnsiJoins()
        With vAnsiJoins
          .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
          .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
          .Add("products p", "bta.product", "p.product")
          .AddLeftOuterJoin("product_warehouses pw", "bta.product", "pw.product", "bta.warehouse", "pw.warehouse")
        End With
        Dim vWherefields As New CDBFields(New CDBField("bta.batch_number", BatchNumber))
        With vWherefields
          .Add("p.stock_item", "Y")
          .Add("bta.issued", CDBField.FieldTypes.cftInteger, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
          .Add("tt.transaction_sign", "C")
          .Add("tt.negatives_allowed", "N")
        End With
        Dim vAttrs As String = "bta.batch_number, bta.transaction_number, bta.line_number, p.product, p.last_stock_count AS product_stock_count, pw.warehouse, pw.last_stock_count AS warehouse_stock_count"
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "batch_transaction_analysis bta", vWherefields, "", vAnsiJoins)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        If vRS.Fetch Then
          pErrorMsg = ProjectText.String33075   'Batch {0} contains stock products that appear to be in stock.  Batch must be picked & confirmed before it can be posted
          vErrorCount += 1
        End If
        vRS.CloseRecordSet()
      End If

      If vErrorCount = 0 And Not (BatchType = BatchTypes.CreditSales OrElse BatchType = BatchTypes.DirectDebit OrElse BatchType = BatchTypes.StandingOrder) Then
        'For Batch containing Line Types L, N, K 'S/L Allocation of Cash-Invoice, S/L Invoice Payment, Sundry Credit Note Invoice Allocation only
        'Check for unposted reversal of invoice payment for the same invoice number, where current batch contains invoice payment, where the reversal should be posted first, raise error
        If mvEnv.Connection.GetCount("batch_transaction_analysis",
                           New CDBFields({New CDBField("batch_number", BatchNumber),
                                          New CDBField("line_type", "'L','N','K'", CDBField.FieldWhereOperators.fwoIn)})) > 0 Then

          Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("batch_transaction_analysis bta2", "bta.invoice_number", "bta2.invoice_number")})
          vAnsiJoins.Add("reversals r", "bta2.batch_number", "r.batch_number", "bta2.transaction_number", "r.transaction_number", "bta2.line_number", "r.line_number")
          vAnsiJoins.Add("invoice_payment_history iph", "bta2.batch_number", "iph.allocation_batch_number", "bta2.transaction_number", "iph.allocation_transaction_number", "bta2.line_number", "iph.allocation_line_number")
          vAnsiJoins.Add("batches b", "bta2.batch_number", "b.batch_number")
          vAnsiJoins.AddLeftOuterJoin("reversals r2", "bta.batch_number", "r2.batch_number", "bta.transaction_number", "r2.transaction_number", "bta.line_number", "r2.line_number")

          Dim vWhereFields As New CDBFields(New CDBField("bta.batch_number", BatchNumber))
          vWhereFields.Add("bta.invoice_number", "", CDBField.FieldWhereOperators.fwoNotEqual)
          vWhereFields.Add("bta2.batch_number", CDBField.FieldTypes.cftInteger, "bta.batch_number", CDBField.FieldWhereOperators.fwoNotEqual)
          vWhereFields.Add("r2.batch_number", "", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual)
          vWhereFields.Add("r2.transaction_number", "", CDBField.FieldWhereOperators.fwoEqual)
          vWhereFields.Add("r2.line_number", "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
          vWhereFields.Add("posted_to_nominal", "N")

          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bta2.batch_number", "batch_transaction_analysis bta", vWhereFields, "", vAnsiJoins)
          vSQLStatement.Distinct = True
          Dim vErrorBatchNumbers As List(Of String) = vSQLStatement.GetValues()
          If vErrorBatchNumbers.Count > 0 Then
            vErrorCount += 1
            pErrorMsg = String.Format(ProjectText.CannotPostInvoiceAllocationBatch, "{0}", String.Join(",", vErrorBatchNumbers)) 'Batch {0} contains invoice allocations and cannot be posted. Please post Batch(es) {1} which contains reversals against the same invoice first.
          End If
        End If
      End If

      If vErrorCount > 0 Then mvClassFields.Item(BatchFields.DetailCompleted).Bool = False
      pErrorList = vErrors
      Return vErrorCount
    End Function
    Private Sub UpdateNextPaymentDue(ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan)
      Dim vOPSNo As Integer
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vFreqAmount As Double

      With pPP
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then
          vOPSNo = IntegerValue(mvConn.GetValue("SELECT scheduled_payment_number FROM order_payment_history WHERE batch_number = " & pBTA.BatchNumber & " AND transaction_number = " & pBTA.TransactionNumber & " AND line_number = " & pBTA.LineNumber & " AND order_number = " & .PlanNumber))
        End If
        If vOPSNo > 0 Then
          For Each vOPS In .ScheduledPayments
            If vOPS.ScheduledPaymentNumber = vOPSNo Then
              If vOPS.RevisedAmount.Length > 0 Then
                vFreqAmount = Val(vOPS.RevisedAmount)
              Else
                vFreqAmount = vOPS.AmountDue
              End If
            End If
            If vFreqAmount > 0 Then Exit For
          Next vOPS
          If vFreqAmount = 0 Then vFreqAmount = .FrequencyAmount
        Else
          vFreqAmount = .FrequencyAmount
        End If
        .NextPaymentDue = .CalculateNextPaymentDue(.NextPaymentDue, .RenewalDate, pBTA.Amount, vFreqAmount, .Balance, vOPS)

        'if balance is renewal amount and if they have had a free initial period within the last year ,
        'and if there is a direct debit record, set next payment due to dd start date */
        If .Balance = .RenewalAmount And .Term < 0 And .DirectDebitStatus <> PaymentPlan.ppYesNoCancel.ppNo Then
          CheckDDIncentiveDate(pPP)
        End If
      End With

    End Sub
    Public Sub PrintLog(ByRef pError As String)
      Dim vWriter As IO.StreamWriter = Nothing
      Try
        Debug.Print("Printing Log " & pError)
        Dim vFileName As String = mvEnv.GetLogFileName("Batch" & BatchNumber & ".log")
        vWriter = My.Computer.FileSystem.OpenTextFileWriter(vFileName, True)
        vWriter.WriteLine(pError)
      Finally
        If vWriter IsNot Nothing Then vWriter.Close()
      End Try
    End Sub
    Private Function ProcessArrears(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double) As Double
      Dim vAmount As Double
      If pAmount > pPP.Arrears Then
        vAmount = pPP.Arrears
      Else
        vAmount = pAmount
      End If
      pPP.Arrears = pPP.Arrears - vAmount
      pPP.Balance = pPP.Balance - vAmount
      ProcessPaymentPlanDetailsArrears(pBT, pBTA, pPP, vAmount)
      ProcessArrears = pAmount - vAmount
    End Function
    Public Sub ProcessCategories(ByRef pNumber As Integer, ByRef pContactType As Contact.ContactTypes, ByRef pActivity As String, ByRef pActivityValue As String, ByRef pSource As String, ByRef pFrom As String, ByRef pTo As String, ByRef pProductCatExtension As Boolean, Optional ByRef pMemberCategories As Boolean = False, Optional ByVal pMembershipNumber As Integer = 0, Optional ByVal pActivityDurationMonths As Integer = 0)
      Dim vTableName As String = String.Empty
      Dim vAttrName As String = String.Empty
      Dim vCC As ContactCategory = Nothing
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        vTableName = "organisation_categories"
        vAttrName = "organisation_number"
        vCC = New OrganisationCategory(mvEnv)
      Else
        vTableName = "contact_categories"
        vAttrName = "contact_number"
        vCC = New ContactCategory(mvEnv)
      End If

      Dim vSQLStatement As SQLStatement = Nothing
      Dim vRecordSet As CDBRecordSet = Nothing
      If CDate(pTo) >= CDate(pFrom) Then
        Dim vWhereFields As New CDBFields
        Dim vFound As Boolean = False
        If pMemberCategories Then
          'Processing a member payment
          If mvMeFutureChangeTrigger = "RENEWAL_DATE" Then
            With vWhereFields
              .Add(vAttrName, pNumber)
              .Add("activity", pActivity)
              .Add("activity_value", pActivityValue)
              .Add("valid_from", CDBField.FieldTypes.cftDate, mvOrgRenewalDate)
              .Add("valid_to", CDBField.FieldTypes.cftNumeric, "valid_from", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            End With

            vSQLStatement = New SQLStatement(mvEnv.Connection, vCC.FieldNames, vCC.AliasedTableName, vWhereFields)
            vRecordSet = vSQLStatement.GetRecordSet()
            If vRecordSet.Fetch() = True Then
              'Activity found for future member type, so extend it rather than create a new one (below)
              vCC.InitFromRecordSet(vRecordSet)
              vCC.Update(vCC.ValidFrom, pTo)
              If vCC.IsValidForUpdate Then
                vFound = True
                vCC.Save("automatic")
              End If
            End If
            vRecordSet.CloseRecordSet()

            If Not (vFound) And pMembershipNumber > 0 Then
              Dim vFMT As New FutureMembershipType(mvEnv)
              vFMT.Init(pMembershipNumber)
              If vFMT.Existing Then
                Dim vRecordTo As Date = CDate(vFMT.FutureChangeDate).AddDays(-1)
                If CDate(pTo).CompareTo(vRecordTo) > 0 Then
                  'To date needs to be set to the earliest of pTo or the future_change_date
                  pTo = vRecordTo.ToString(CAREDateFormat)
                End If
              End If
            End If
          End If
        End If

        If Not (vFound) Then
          'Either not member categories or we did not find a record to update for a member
          vWhereFields.Clear()
          vWhereFields.Add(vAttrName, CDBField.FieldTypes.cftLong, pNumber)
          vWhereFields.Add("activity", CDBField.FieldTypes.cftCharacter, Trim(pActivity))
          vWhereFields.Add("activity_value", CDBField.FieldTypes.cftCharacter, Trim(pActivityValue))
          If pMemberCategories = False OrElse CDate(pFrom).CompareTo(CDate(pTo)) <> 0 Then
            vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, pFrom)
          End If

          If pContactType = Contact.ContactTypes.ctcOrganisation Then
            vCC = New OrganisationCategory(mvEnv)
          Else
            vCC = New ContactCategory(mvEnv)
          End If
          vSQLStatement = New SQLStatement(mvEnv.Connection, vCC.FieldNames, vCC.AliasedTableName, vWhereFields, "valid_to DESC")
          vRecordSet = vSQLStatement.GetRecordSet()

          Dim vDone As Boolean = False
          Dim vUpdate As Boolean = False
          Dim vTo As Date = CDate(pTo)

          If vRecordSet.Fetch Then
            Do
              vUpdate = False
              If pContactType = Contact.ContactTypes.ctcOrganisation Then
                vCC = New OrganisationCategory(mvEnv)
              Else
                vCC = New ContactCategory(mvEnv)
              End If
              vCC.InitFromRecordSet(vRecordSet)

              'Set ValidFrom date
              If CDate(pFrom).CompareTo(CDate(pTo)) <> 0 Then
                vCC.ValidFrom = pFrom
                vUpdate = True
              End If

              'Set ValidTo date
              If CDate(pFrom).CompareTo(CDate(pTo)) = 0 Then
                'only put extension in place if this is a product related act
                If (pProductCatExtension OrElse pActivityDurationMonths > 0) Then
                  'J1445: If Product Activity Duration Months is set, greater than 0, then set valid_to date to be current pTo date
                  'incremented by the pActivityDurationMonths months value
                  vTo = If(pActivityDurationMonths > 0, CDate(pTo).AddMonths(pActivityDurationMonths), CDate(pTo).AddYears(99))
                End If
                If vTo.CompareTo(CDate(vCC.ValidTo)) > 0 Then
                  vCC.ValidTo = vTo.ToString(CAREDateFormat)
                  vUpdate = True
                End If
              Else
                vCC.ValidTo = pTo
              End If

              If vUpdate Then
                If vCC.IsValidForUpdate Then
                  vCC.Save("automatic")
                Else
                  PrintLog(String.Format("Update of {0} failed - Number: {1} Category: {2} Value: {3} - Possible Duplicate Records?", {vTableName, pNumber, pActivity, pActivityValue}))
                End If
                If pMemberCategories Then vDone = True 'We only need to update one entry
              End If
            Loop While vRecordSet.Fetch = True AndAlso vDone = False
            vRecordSet.CloseRecordSet()
          Else
            vRecordSet.CloseRecordSet()
            If (pProductCatExtension OrElse pActivityDurationMonths > 0) AndAlso CDate(pFrom).CompareTo(CDate(pTo)) = 0 Then
              'J1445: If Product Activity Duration Months is set, greater than 0, then set valid_to date to be current pTo date
              'incremented by the pActivityDurationMonths months value
              vTo = If(pActivityDurationMonths > 0, CDate(pTo).AddMonths(pActivityDurationMonths), CDate(pTo).AddYears(99))
            End If
            vCC.SaveActivity(ContactCategory.ActivityEntryStyles.aesCheckDateRange Or ContactCategory.ActivityEntryStyles.aesSmartClient, pNumber, pActivity, pActivityValue, pSource, pFrom, vTo.ToString(CAREDateFormat), "", "", "", "automatic", "", "")
          End If
        End If
      End If
    End Sub
    Private Sub ProcessContactSuppressions(ByRef pContactNumber As Integer, ByRef pMailingSuppression As String, ByRef pValidFrom As String, ByRef pValidTo As String)
      Dim vFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vRecordSet As CDBRecordSet

      'will only happen when join date = future change date
      'i.e. should not occur in proper use of system, but code here just in case
      If CDate(pValidFrom) <= CDate(pValidTo) Then
        vFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        vFields.Add("mailing_suppression", CDBField.FieldTypes.cftCharacter, pMailingSuppression)
        vRecordSet = mvConn.GetRecordSet("SELECT valid_from FROM contact_suppressions WHERE " & mvConn.WhereClause(vFields))
        If vRecordSet.Fetch() = True Then
          vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidTo)
          vUpdateFields.AddAmendedOnBy("automatic")
          mvConn.UpdateRecords("contact_suppressions", vUpdateFields, vFields, False)
        Else
          vFields.Add("valid_from", CDBField.FieldTypes.cftDate, pValidFrom)
          vFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidTo)
          vFields.AddAmendedOnBy("automatic")
          mvConn.InsertRecord("contact_suppressions", vFields)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    Private Sub ProcessGiftAidDeclarations(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pType As String, ByRef pAmount As Double)
      Dim vRecordSet As CDBRecordSet
      Dim vContact As New Contact(mvEnv)
      Dim vContactLink As New ContactLink(mvEnv)
      Dim vContactNos As String = ""
      Dim vSQL As String
      Dim vSQLRestrict As String
      Dim vDecNumber As Integer
      Dim vContactNumber As Integer

      If ((pAmount > 0) OrElse (pAmount < 0 AndAlso mvImportPayment = True)) Then 'Do create reversal lines except for import
        If pBT.ContactType = Contact.ContactTypes.ctcJoint Then
          vContact.Init((pBT.ContactNumber))
          For Each vContactLink In vContact.GetJointLinks(True)
            vContactNos = vContactNos & "," & vContactLink.ContactNumber2
          Next vContactLink
        End If
        If vContactNos.Length > 0 Then
          vContactNos = " IN (" & pBT.ContactNumber & vContactNos & ")"
        Else
          vContactNos = " = " & pBT.ContactNumber
        End If

        vSQL = "SELECT declaration_number, contact_number FROM gift_aid_declarations WHERE contact_number " & vContactNos & " AND start_date" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (pBT.TransactionDate)) & " AND (end_date" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (pBT.TransactionDate)) & " OR end_date IS NULL )"

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidMergeCancellation) Then
          If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason).Length > 0 Then
            vSQL = vSQL & " AND (cancellation_reason IS NULL or cancellation_reason <> '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason) & "')"
          End If
        End If

        'Exclude any Declarations that are cancelled on or after the transaction date
        vSQL = vSQL & " AND (cancelled_on " & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, (pBT.TransactionDate)) & " OR cancelled_on IS NULL)"
        vSQLRestrict = " (batch_number = " & pBTA.BatchNumber & " AND transaction_number = " & pBTA.TransactionNumber & " AND order_number IS NULL)"
        If pBTA.PaymentPlanNumber > 0 Then
          vSQLRestrict = " AND ((order_number = " & pBTA.PaymentPlanNumber & " AND batch_number IS NULL) OR " & vSQLRestrict
          vSQLRestrict = vSQLRestrict & ")"
        Else
          vSQLRestrict = " AND " & vSQLRestrict
        End If
        vSQL = vSQL & " AND declaration_type IN('" & pType & "','A')"
        'First check for a linked Declaration
        vRecordSet = mvConn.GetRecordSet(vSQL & vSQLRestrict)
        If vRecordSet.Fetch() = True Then
          'There is a declaration so create declaration lines unclaimed
          vDecNumber = vRecordSet.Fields("declaration_number").IntegerValue
          vContactNumber = vRecordSet.Fields("contact_number").IntegerValue
        End If
        vRecordSet.CloseRecordSet()

        If vDecNumber = 0 Then
          'Second, if no linked declaration found then check for an unlinked one
          vSQL = vSQL & " AND order_number IS NULL AND batch_number IS NULL"
          vRecordSet = mvConn.GetRecordSet(vSQL)
          If vRecordSet.Fetch() = True Then
            'There is a declaration so create declaration lines unclaimed
            vDecNumber = vRecordSet.Fields("declaration_number").IntegerValue
            vContactNumber = vRecordSet.Fields("contact_number").IntegerValue
          End If
          vRecordSet.CloseRecordSet()
        End If

        If vDecNumber > 0 Then
          CreatedDeclarationLinesUnclaimed(pBT, pBTA, vDecNumber, vContactNumber, "D", pAmount)
        End If
      End If

    End Sub
    Private Sub ProcessCheckLegacyReceipt(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pAmount As Double)
      Dim vRS As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      'Reversal of Product payment:
      'Check for Legacy Bequest Receipt, if one exists set status on original to Reversed
      'Create new Receipt for Reversal
      vRS = mvConn.GetRecordSet("SELECT receipt_number,lb.legacy_number,lb.bequest_number,amount,estimated_outstanding, date_received FROM reversals r,legacy_bequest_receipts lbr,legacy_bequests lb WHERE r.batch_number = " & pBTA.BatchNumber & " AND r.transaction_number = " & pBTA.TransactionNumber & " AND r.line_number = " & pBTA.LineNumber & " AND lbr.batch_number = r.was_batch_number AND lbr.transaction_number = r.was_transaction_number AND lbr.line_number = r.was_line_number AND lb.legacy_number = lbr.legacy_number AND lb.bequest_number = lbr.bequest_number")
      If vRS.Fetch() = True Then
        vWhereFields.Add("legacy_number", CDBField.FieldTypes.cftLong, vRS.Fields(2).Value)
        vWhereFields.Add("bequest_number", CDBField.FieldTypes.cftLong, vRS.Fields(3).Value)
        vUpdateFields.Add("estimated_outstanding", CDBField.FieldTypes.cftNumeric, vRS.Fields(4).DoubleValue + vRS.Fields(5).DoubleValue)
        mvConn.UpdateRecords("legacy_bequests", vUpdateFields, vWhereFields)
        vWhereFields.Add("receipt_number", CDBField.FieldTypes.cftLong, vRS.Fields(1).Value)
        vUpdateFields.Clear()
        vUpdateFields.Add("status", CDBField.FieldTypes.cftCharacter, "R")
        mvConn.UpdateRecords("legacy_bequest_receipts", vUpdateFields, vWhereFields)
        vUpdateFields.Clear()
        vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
        vUpdateFields.Add("legacy_number", CDBField.FieldTypes.cftLong, vRS.Fields(2).Value)
        vUpdateFields.Add("bequest_number", CDBField.FieldTypes.cftLong, vRS.Fields(3).Value)
        vUpdateFields.Add("receipt_number", CDBField.FieldTypes.cftLong, mvEnv.GetControlNumber("LR"))
        vUpdateFields.Add("amount", CDBField.FieldTypes.cftNumeric, pAmount)
        vUpdateFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
        vUpdateFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, pBTA.TransactionNumber)
        vUpdateFields.Add("line_number", CDBField.FieldTypes.cftInteger, pBTA.LineNumber)
        vUpdateFields.Add("date_received", CDBField.FieldTypes.cftDate, vRS.Fields("date_received").Value) 'Set this to the original date
        vUpdateFields.Add("notes", CDBField.FieldTypes.cftMemo, pBTA.Notes)
        mvConn.InsertRecord("legacy_bequest_receipts", vUpdateFields)
      End If
      vRS.CloseRecordSet()
    End Sub
    Private Sub ProcessInAdvance(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double)
      Dim vRS As CDBRecordSet
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vOPH As New OrderPaymentHistory
      Dim vWhereFields As New CDBFields
      Dim vOPSFound As Boolean

      With vWhereFields
        .Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
      End With

      vOPH.Init(mvEnv)
      vRS = mvConn.GetRecordSet("SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE " & mvConn.WhereClause(vWhereFields))
      If vRS.Fetch() = True Then vOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
      vRS.CloseRecordSet()

      If vOPH.Existing Then
        vOPH.SetPosted(True)
        vOPH.Status = "I"
        vOPH.Save()
        If vOPH.ScheduledPaymentNumber.Length > 0 Then
          For Each vOPS In pPP.ScheduledPayments
            If vOPS.ScheduledPaymentNumber = CDbl(vOPH.ScheduledPaymentNumber) Then vOPSFound = True
            If vOPSFound = True Then Exit For
          Next vOPS
          If vOPSFound = False Then
            'The scheduled payments selected could be just for next year, so just get the payment
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv, CInt(vOPH.ScheduledPaymentNumber))
            If vOPS.PlanNumber = pPP.PlanNumber Then vOPSFound = True
          End If
          If vOPSFound Then
            vOPS.ProcessPayment() 'pAmount
            vOPS.Save()
          End If
        End If
      Else
        pPP.PaymentNumber = pPP.PaymentNumber + 1
        vOPH.SetValues((pBTA.BatchNumber), (pBTA.TransactionNumber), (pPP.PaymentNumber), (pPP.PlanNumber), pAmount, (pBTA.LineNumber), 0, 0, True)
        vOPH.Status = "I"
        vOPH.Save()
      End If
      pPP.InAdvance = pPP.InAdvance + pAmount

      ProcessProduct(pBT, pBTA, (mvCompanyControl.InAdvanceProduct), (mvCompanyControl.InAdvanceRate), "", 1, pAmount, 1, "", 0)
    End Sub

    ''' <summary>Posting invoice payments; creates Invoice Payment History, Invoices and Financial History Details records depending upon the type of invoice payment</summary>
    ''' <param name="pBT">The BatchTransaction record being processed.</param>
    ''' <param name="pBTA">The BatchTransactionAnalysis record being procesed</param>
    ''' <param name="pInvoiceNumber">The number of the invoice being paid</param>
    ''' <remarks>S/L Invoice Payment (line type N) - creates Invoice Payment History, sales ledger Invoice and Financial History Details
    ''' S/L Unallocated Cash (line type U) - creates sales ledger Invoice and Financial History Details
    ''' S/L Allocation of Cash-Invoice (line type L) - creates Financial History Details
    ''' Sundry Credit Note Invoice Allocation (line type K) - creates Financial History Details</remarks>
    Private Sub ProcessInvoice(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pInvoiceNumber As Integer)
      If Not mvConn.InTransaction Then mvConn.StartTransaction() 'TRANSACTION START HERE

      Dim vTransactionSignD As Boolean = (pBT.TransactionSign = "D")
      Dim vCreateSLInvoice As Boolean = Not (pBTA.LineType = "L" OrElse pBTA.LineType = "K")
      If pBTA.LineType = "N" Then WriteInvoicePaymentHistory(pBTA.InvoiceNumber, pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, pBTA.Amount, pBT.TransactionDate, pBTA.IsPartRefund)
      If vCreateSLInvoice Then WriteSalesLedgerInvoice(pBT, pBTA, vTransactionSignD)
      WriteFinancialHistoryAnalysis(pBT, pBTA, "", "", "", 0, pBTA.Amount, "", 0)

    End Sub
    Private Sub ProcessMembershipPayment(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByRef pExpiry As String, ByRef pGraceExpiry As String, ByRef pAmount As Double)
      Dim vMember As Member
      Dim vActivity As String
      Dim vActivityValue As String
      Dim vMailingSuppression As String
      Dim vProcessedOrders As Boolean
      Dim vExpiry As String
      Dim vContact As Contact
      Dim vFutureMailingSuppression As String
      Dim vFutureActivity As String = ""
      Dim vFutureActivityValue As String = ""
      Dim vHasFutureType As Boolean
      Dim vProductCatExtension As Boolean

      vProcessedOrders = False
      For Each vMember In pPP.CurrentMembers
        If vProcessedOrders = False Then
          ProcessPaymentPlanDetails(pBT, pBTA, pPP, pAmount, pGraceExpiry)
          vProcessedOrders = True
        End If
        vHasFutureType = Not (vMember.FutureMembershipType Is Nothing)

        'First deal with suppressions
        vMailingSuppression = vMember.MembershipType.MailingSuppression
        If mvMeFutureChange = "Y" And vHasFutureType Then
          If vMailingSuppression = vMember.FutureMembershipType.MailingSuppression Then
            vFutureMailingSuppression = "" 'no difference
            vExpiry = pExpiry
          Else 'difference so expire current
            vFutureMailingSuppression = vMember.FutureMembershipType.MailingSuppression
            vExpiry = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
          End If
        Else
          vFutureMailingSuppression = ""
          vExpiry = pExpiry 'not a future record
          If pPP.ContinuousRenewals Then vExpiry = pGraceExpiry
        End If
        If vMailingSuppression.Length > 0 Then
          ProcessContactSuppressions((vMember.ContactNumber), vMailingSuppression, (vMember.Joined), vExpiry)
        End If
        If vFutureMailingSuppression.Length > 0 Then
          ProcessContactSuppressions((vMember.ContactNumber), vFutureMailingSuppression, mvOrgRenewalDate, pExpiry)
        End If
        'check reversal
        If mvMeFutureChange = "R" And vHasFutureType Then
          If vMailingSuppression <> vMember.FutureMembershipType.MailingSuppression Then
            DeleteContactSuppressions((vMember.ContactNumber), vMember.FutureMembershipType.MailingSuppression)
          End If
        End If

        'Now deal with activities
        If mvMeFutureChange = "Y" And vHasFutureType Then
          If vMember.MembershipType.Activity = vMember.FutureMembershipType.Activity And vMember.MembershipType.ActivityValue = vMember.FutureMembershipType.ActivityValue Then
            vFutureActivity = ""
            vExpiry = pExpiry
          Else
            vFutureActivity = vMember.FutureMembershipType.Activity
            vFutureActivityValue = vMember.FutureMembershipType.ActivityValue
            vExpiry = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
          End If
        Else
          vFutureActivity = ""
          vExpiry = pExpiry
          If pPP.ContinuousRenewals Then vExpiry = pGraceExpiry
        End If
        vActivity = vMember.MembershipType.Activity
        vActivityValue = vMember.MembershipType.ActivityValue
        vProductCatExtension = mvEnv.GetConfigOption("opt_fp_product_cat_extension")

        If vActivity.Length > 0 Then
          ProcessCategories((vMember.ContactNumber), (vMember.ContactType), vActivity, vActivityValue, (vMember.Source), (vMember.Joined), vExpiry, vProductCatExtension, True, (vMember.MembershipNumber))
        End If
        If vFutureActivity.Length > 0 Then
          ProcessCategories((vMember.ContactNumber), (vMember.ContactType), vFutureActivity, vFutureActivityValue, (vMember.Source), mvOrgRenewalDate, pExpiry, vProductCatExtension, True, (vMember.MembershipNumber))
        End If
        'check reversal
        If mvMeFutureChange = "R" And vHasFutureType Then
          If vMember.MembershipType.Activity <> vMember.FutureMembershipType.Activity Or vMember.MembershipType.ActivityValue <> vMember.FutureMembershipType.ActivityValue Then
            ProcessCategories((vMember.ContactNumber), (vMember.ContactType), vMember.FutureMembershipType.Activity, vMember.FutureMembershipType.ActivityValue, vMember.Source, CDate(mvOrgRenewalDate).AddYears(-1).ToString(CAREDateFormat), CDate(mvOrgRenewalDate).AddYears(-1).ToString(CAREDateFormat), vProductCatExtension)
          End If
        End If

        If mvMeFutureChange = "Y" And vHasFutureType Then
          If (mvMeFutureChangeTrigger = "PAYMENT_DATE" And CDate(mvOrgRenewalDate) <= Today) Or (mvMeFutureChangeTrigger = "FIRST_PAYMENT" And CDate(mvOrgRenewalDate) <= Today And mvOrgBalance > 0) Or (mvMeFutureChangeTrigger = "FIRST_PAYMENT" And CDate(mvOrgRenewalDate) > Today) Then 'And mvOrgBalance = 0) Then
            pPP.ProcessFutureMembership(vMember, mvOrgRenewalDate, mvOrgTerm, False)
          End If
        End If

        If pPP.GiftMembership = True Then
          vContact = New Contact(mvEnv) 'TODO Get this somewhere else
          vContact.Init((pPP.ContactNumber))
          ProcessCategories((pPP.ContactNumber), (vContact.ContactType), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivity), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivityValue), (vMember.Source), (vMember.Joined), pExpiry, vProductCatExtension)
          vContact = Nothing
        End If

        If vMember.MembershipType.MembersPerOrder = 0 And pPP.GiftMembership And pPP.OneYearGift Then
          vMember.CancellationReason = mvGOYGReason
          vMember.CancelledOn = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(pExpiry)))
          vMember.CancelledBy = "automatic"
        End If
      Next vMember
      If vProcessedOrders = False Then ProcessPaymentPlanDetails(pBT, pBTA, pPP, pAmount, pGraceExpiry)
    End Sub
    Private Sub ProcessPaymentPlan(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vRecordSet As CDBRecordSet
      Dim vPP As New PaymentPlan
      Dim vAmount As Double
      Dim vReceived As Double
      Dim vDiscount As Double
      Dim vWriteOff As Double
      Dim vInAdvance As Boolean
      Dim vLastRenewalDate As String
      Dim vError As String = ""
      Dim vMemsPerOrder As Integer
      Dim vLineNumber As Integer
      Dim vMember As New Member
      Dim vPPD As PaymentPlanDetail
      Dim vBranch As String
      Dim vExpiryDate As String
      Dim vReCalc As Boolean
      Dim vMoveDates As Boolean
      Dim vOPS As OrderPaymentSchedule
      Dim vCreateSchedule As Boolean
      Dim vRenewalAmount As Double
      Dim vCNRequired As Integer
      Dim vOverPayment As Double
      Dim vOldRenewDate As String = ""
      Dim vArrears As Double
      Dim vOrigOPSClaimDate As String = ""

      Dim vOverriddenPPRenewalDate As String = String.Empty
      Dim vOverriddenPPNextPaymentDue As String = String.Empty
      Dim vPPDatesOveridden As Boolean = False

      'For the moment go and read the payment plan and details as each analysis line is processed
      If pBTA.PaymentPlanNumber = 0 Then RaiseError(DataAccessErrors.daePaymentPlanNotFound)
      vPP.Init(mvEnv)
      vAmount = pBTA.Amount
      vRecordSet = mvConn.GetRecordSet("SELECT " & vPP.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstAll Or PaymentPlan.PayPlanRecordSetTypes.pprstDetailLines Or PaymentPlan.PayPlanRecordSetTypes.pprstDetailProduct) & " FROM orders o, order_details od, products p, rates r WHERE o.order_number = " & pBTA.PaymentPlanNumber & " AND o.order_number = od.order_number AND p.product = od.product AND r.product = od.product and r.rate = od.rate ORDER BY detail_number")
      If vRecordSet.Fetch() = False Then
        vRecordSet.CloseRecordSet()
        PrintLog("Error : locked orders record for Payment Plan: " & pBTA.PaymentPlanNumber)
        ProcessPaymentPlanError(vPP, pBT, pBTA, (mvCompanyControl.LockedProduct), (mvCompanyControl.LockedRate), vAmount)
        Exit Sub
      End If
      vPP.InitFromRecordSet(mvEnv, vRecordSet, PaymentPlan.PayPlanRecordSetTypes.pprstAll Or PaymentPlan.PayPlanRecordSetTypes.pprstDetailLines Or PaymentPlan.PayPlanRecordSetTypes.pprstDetailProduct)

      If ForceCreationOfRegularProvisionalPayment(pBT, vPP) Then
        vOverriddenPPRenewalDate = vPP.RenewalDate
        vOverriddenPPNextPaymentDue = vPP.NextPaymentDue
        vPPDatesOveridden = True

        Dim vOPSSql As SQLStatement
        Dim vOPSWhereFields As CDBFields
        Dim vOPSDataTable As DataTable
        vOPSWhereFields = New CDBFields(New CDBField("order_number", vPP.OrderNumber))
        vOPSSql = New SQLStatement(Me.Environment.Connection, "scheduled_payment_number,due_date,claim_date", "order_payment_schedule", vOPSWhereFields, "scheduled_payment_number DESC")
        vOPSDataTable = vOPSSql.GetDataTable
        vPP.RenewalDate = vOPSDataTable.Rows(0).Item("due_date").ToString()
        vPP.NextPaymentDue = vOPSDataTable.Rows(0).Item("claim_date").ToString()
      End If
      If vPP.ScheduledPayments(False).Count() = 0 Then
        'RaiseError daePaymentPlanNotFound     'Select all records (not just oustanding records) now before the balance is upated
      End If
      vRecordSet.CloseRecordSet()
      vReceived = pBTA.Amount
      vMoveDates = True
      vCreateSchedule = mvBatchType <> BatchTypes.None
      mvMeFutureChange = "N"
      If Len(vPP.CancellationReason) > 0 And vAmount > 0 Then
        'order cancelled after posting of money - so post to exception product
        PrintLog("Error : cancelled orders record for Payment Plan: " & pBTA.PaymentPlanNumber)
        ProcessPaymentPlanError(vPP, pBT, pBTA, (mvCompanyControl.CancelledProduct), (mvCompanyControl.CancelledRate), vAmount)
      Else
        mvOrgRenewalDate = vPP.RenewalDate
        mvOrgTerm = vPP.Term
        mvOrgBalance = vPP.Balance
        vMember.Init(mvEnv)
        If vPP.PlanType = CDBEnvironment.ppType.pptMember Then
          mvMemTypeTerm = vPP.MembershipType.PaymentTerm
          If mvBatchType <> BatchTypes.None Then
            vPP.LoadMembers()
            vBranch = vPP.Branch
            For Each vMember In vPP.CurrentMembers
              If vMember.ContactNumber = vPP.ContactNumber Then
                vBranch = vMember.Branch
                Exit For
              End If
            Next vMember
            If vPP.CurrentMembers.Count() > 0 Then SetMeFutureChangeFlag(vPP, vAmount)
          End If
        ElseIf vPP.PlanType = CDBEnvironment.ppType.pptLoan Then
          'Calculate the Loan interest but OPS not re-created here
          Dim vIntToDate As Date = CDate(If(CDate(pBT.TransactionDate) > Today, pBT.TransactionDate, TodaysDate()))
          If vIntToDate >= CDate(vPP.Loan.LoanCapitalisationDate) Then
            If CDate(pBT.TransactionDate) < CDate(vPP.Loan.LoanCapitalisationDate) Then vIntToDate = CDate(vPP.Loan.LoanCapitalisationDate).AddDays(-1)
          End If
          vPP.CalculateLoanInterest("", False, vIntToDate.ToString(CAREDateFormat), True)
        End If

        'Before we start the Transaction get some Control Numbers so that payment schedule
        'creation does not have to within the transaction - not required for reversals
        'For DD/CC/SO retrieve 2 journal numbers but always get the 1 OPS number
        '(Only need more than 1 OPS number if renewing installment Payment Plan)
        If vAmount > 0 And Not mvConn.InTransaction Then
          If (mvBatchType = BatchTypes.DirectDebit Or mvBatchType = BatchTypes.CreditCardAuthority Or mvBatchType = BatchTypes.StandingOrder) Then
            'Get 2 journal numbers each time because regular don's will go into OPS creation twice if balance = 0
            'And they will probably mostly be paid by DD/CC/SO
            vCNRequired = 2
          Else
            'For everything else just get 1 number
            vCNRequired = 1
          End If
          mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnJournal, vCNRequired)
          mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnPaymentSchedule, 1)
        End If

        If Not mvConn.InTransaction Then mvConn.StartTransaction() 'TRANSACTION START HERE

        If vAmount > 0 Then

          vMoveDates = Not (mvBatchType = BatchTypes.FinancialAdjustment And CheckInAdvance((pBT.BatchNumber), (pBT.TransactionNumber)) And vPP.PlanType = CDBEnvironment.ppType.pptMember And vPP.FixedRenewalCycle And vPP.PreviousRenewalCycle And (vPP.ProportionalBalanceSetting And PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment) > 0)

          If vPP.Balance = 0 And Len(mvCompanyControl.InAdvanceProductCode) > 0 And ((vPP.PlanType = CDBEnvironment.ppType.pptMember Or vPP.PlanType = CDBEnvironment.ppType.pptOther) And CDate(vPP.RenewalDate) > CDate(pBT.TransactionDate)) Then
            'Renewal is not due yet and they dont owe any money
            ProcessInAdvance(pBT, pBTA, vPP, vAmount)
            vInAdvance = True
            If vPP.PlanType = CDBEnvironment.ppType.pptMember And mvMeFutureChange = "Y" And mvMeFutureChangeTrigger = "FIRST_PAYMENT" And mvOrgRenewalDate = vMember.FutureChangeDate Then
              'Paying future membership in advance
              vMember = DirectCast(vPP.CurrentMembers.Item(1), Member)
              vPP.ProcessFutureMembership(vMember, mvOrgRenewalDate, mvOrgTerm, True)
              mvMemTypeTerm = vPP.MembershipType.PaymentTerm
              mvMeFutureChange = "N"
            End If
          Else
            'Pay off any arrears first
            If vPP.Arrears <> 0 Then
              If vPP.Balance < vPP.Arrears Then
                vError = "Balance is less than Arrears"
              Else
                vAmount = ProcessArrears(pBT, pBTA, vPP, vAmount)
                If vAmount = 0 Then
                  'TA BC 3470: We have paid off Arrears and have no extra to pay, If only Arrears
                  'was owed, ensure Renewal Date still gets rolled forward.
                  If vPP.Balance = 0 And Not vPP.ContinuousRenewals Then UpdateRenewalDate(pBT, pBTA, vPP, True)
                  'End TA BC 3470
                End If
              End If
            End If
            If vAmount > 0 Then
              If vPP.Balance > 0 And (FixTwoPlaces(vAmount) > FixTwoPlaces(vPP.Balance)) Then
                If vPP.PaymentFrequencyFrequency = 1 And (vPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or vPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) And (mvEnv.GetConfig("fp_arrears_claim_method_create") <> "NONE" And FixTwoPlaces(vAmount) = FixTwoPlaces(vPP.Balance + vPP.RenewalAmount)) Then
                  'The next claim has been increased and is the balance + the renewal amount
                  'First pay off the balance before processing the remainder in the normal way
                  ProcessPaymentPlanDetails(pBT, pBTA, vPP, vPP.Balance, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPP.RenewalDate))))
                  If mvMemTypeTerm <> MembershipType.MembershipTypeTerms.mtfMonthlyTerm Then UpdateNextPaymentDue(pBTA, vPP)
                  If vPP.Balance = vPP.RenewalAmount Then UpdateRenewalDate(pBT, pBTA, vPP, True)
                  vAmount = vAmount - vPP.Balance
                  vPP.Balance = 0
                  vCreateSchedule = False 'Do not re-create payment schedule in vPP.CalculateBalance
                End If
              End If
              If vPP.Balance = 0 Then
                'Calculate new balance (for standing orders etc..)
                If vPP.PlanType = CDBEnvironment.ppType.pptMember And mvMeFutureChange = "Y" And mvMeFutureChangeTrigger = "FIRST_PAYMENT" And mvOrgRenewalDate = vMember.FutureChangeDate Then
                  'Paying future membership in advance - new detail lines etc required before processing payment
                  vMember = DirectCast(vPP.CurrentMembers.Item(1), Member)
                  vPP.ProcessFutureMembership(vMember, mvOrgRenewalDate, mvOrgTerm, True)
                  mvMemTypeTerm = vPP.MembershipType.PaymentTerm
                  mvMeFutureChange = "N"
                  'Balances still set to 0 so recalc
                  vPP.CalculateBalance("C", True, vCreateSchedule)
                ElseIf (vPP.PaymentFrequencyFrequency = 1 And vPP.PaymentFrequencyInterval = 1) And vCreateSchedule = True Then
                  'Regular monthly Payment Plan - if the OPS being paid is after the renewal date then temporarily reset the RenewalDate
                  vOPS = New OrderPaymentSchedule
                  vOPS.Init(mvEnv, IntegerValue(pBTA.ScheduledPaymentNumber))
                  If vOPS.Existing Then
                    If CDate(vOPS.DueDate) > CDate(mvOrgRenewalDate) Then
                      vOldRenewDate = vPP.RenewalDate
                      vPP.RenewalDate = vOPS.DueDate
                    ElseIf CDate(vOPS.DueDate) < CDate(mvOrgRenewalDate) Then
                      'Paying a missed payment
                      vOldRenewDate = vPP.RenewalDate
                      vPP.RenewalDate = vPP.CalculateRenewalDate(vPP.RenewalDate, True)
                    End If
                    vOrigOPSClaimDate = vOPS.ClaimDate
                  End If
                  vPP.CalculateBalance("C", True, vCreateSchedule)
                  If IsDate(vOldRenewDate) Then vPP.RenewalDate = vOldRenewDate
                Else
                  If mvBatchType = BatchTypes.DirectDebit Or mvBatchType = BatchTypes.CreditCard Or mvBatchType = BatchTypes.CreditCardWithInvoice Or mvBatchType = BatchTypes.CreditCardAuthority Then
                    'We need to renew the PP so store the original ClaimDate for this payment
                    vOPS = New OrderPaymentSchedule
                    vOPS.Init(mvEnv, CInt(pBTA.ScheduledPaymentNumber))
                    If vOPS.Existing Then vOrigOPSClaimDate = vOPS.ClaimDate
                  ElseIf mvBatchType = BatchTypes.None And vPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
                    'Skipping DD payment so allow OPS to be re-created
                    vCreateSchedule = True
                  End If
                  vPP.CalculateBalance("C", True, vCreateSchedule)
                End If
                vReCalc = True
              End If
              If vError.Length > 0 Then
                vError = "Error : Problem on Payment Plan: " & pBTA.PaymentPlanNumber & " - " & vError
                PrintLog(vError)
                ProcessPaymentPlanError(vPP, pBT, pBTA, (mvCompanyControl.DetailsProduct), (mvCompanyControl.DetailsRate), vAmount)
              ElseIf vPP.Balance = 0 Then
                'They really really really don't owe us any money
                vError = "Error : Zero Balance on Payment Plan: " & pBTA.PaymentPlanNumber
                PrintLog(vError)
                ProcessPaymentPlanError(vPP, pBT, pBTA, (mvCompanyControl.DetailsProduct), (mvCompanyControl.DetailsRate), vAmount, True)
              Else
                'commence order when they actually pay
                If vPP.RenewalDate = vPP.StartDate And Len(mvEnv.GetFixedCycleConfig(vPP.PlanType)) = 0 Then
                  If mvEnv.GetConfigOption("order_dates_to_batch_dates") Then
                    vPP.StartDate = pBT.TransactionDate
                    vPP.RenewalDate = pBT.TransactionDate
                    vPP.NextPaymentDue = pBT.TransactionDate
                  End If
                End If

                If (Not vPP.ContinuousRenewals And vMoveDates) Or (vPP.ContinuousRenewals And vReCalc) Then
                  If vPP.Balance = vPP.RenewalAmount Then
                    UpdateRenewalDate(pBT, pBTA, vPP, True)
                  Else
                    If vPP.Balance > vPP.RenewalAmount Then
                      If (vAmount > FixTwoPlaces(vPP.Balance - vPP.RenewalAmount) And vPP.RenewalPending) Or ((vPP.Balance - vAmount) = 0 And vPP.RenewalAmount = 0) Then
                        UpdateRenewalDate(pBT, pBTA, vPP, True)
                      End If
                    Else
                      If vPP.Balance < vPP.RenewalAmount And vPP.RenewalPending = True Then
                        'TA BC3471: Owe less than Renewal but RP still set: User has overridden
                        'Balance so we still need to roll Renewal on.
                        UpdateRenewalDate(pBT, pBTA, vPP, True)
                      Else
                        'need to roll renewal date forward if CMT has occurred to life mship and balance has been paid off
                        If vPP.Balance - vAmount <= 0 And mvMemTypeTerm = MembershipType.MembershipTypeTerms.mtfLifeTerm Then UpdateRenewalDate(pBT, pBTA, vPP, True)
                      End If
                    End If
                  End If
                End If

                'update branch income so the branch gets some money
                'ab: moved to UpdateRenewalDate If mvBranchIncomePeriod = "FIRST" Then WriteBranchIncome pBT, pBTA, vPP, Nothing, True, 0

                If mvBatchType = BatchTypes.None AndAlso mvIgnoreDiscountForSkip Then
                  vDiscount = 0
                Else
                  vDiscount = ProcessPaymentPlanDetailsDiscounts(pBT, pBTA, vPP)
                End If
                'check to see if payment exceeds balance (due to data entry problem)
                'and if so post excess to over-payment product - ? in advance
                If FixTwoPlaces(vAmount) > FixTwoPlaces(vPP.Balance) Then
                  vLineNumber = pBTA.LineNumber
                  'pBTA.LineNumber = pBT.NextLineNumber
                  'pBT.NextLineNumber = pBT.NextLineNumber + 1
                  PrintLog("Error : Over payment on Payment Plan: " & pBTA.PaymentPlanNumber)
                  ProcessPaymentPlanError(vPP, pBT, pBTA, (mvCompanyControl.OverPaymentProduct), (mvCompanyControl.OverPaymentRate), vAmount, False, True)
                  pBTA.LineNumber = vLineNumber
                  vAmount = vPP.Balance
                  vReceived = vAmount
                  vPP.Balance = 0
                Else
                  vPP.Balance = vPP.Balance - vAmount
                  If pBTA.AcceptAsFull = True Then
                    vWriteOff = vPP.Balance
                    vPP.Balance = 0
                  End If
                End If
              End If
            End If
          End If
        Else
          If vPP.InAdvance > 0 Then
            'TA 9/1/04 used to include pBTA.LineNumber but removed for new Order Payment
            'Schedule code since In Advance Reversals now created with incremental Line Numbers
            'that do not necessarily map back to original payment.
            vInAdvance = CheckInAdvance((pBTA.BatchNumber), (pBTA.TransactionNumber))
          End If
          If Not vInAdvance Then
            vWriteOff = -GetOriginalWriteOff((pBTA.BatchNumber), (pBTA.TransactionNumber), (pBTA.LineNumber), pBTA.AcceptAsFull)

            'Check for over-payments
            vOverPayment = CheckOverpayment(pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, vAmount)
            If vOverPayment <> 0 Then
              'If there was an over-payment then deduct that amount from this payment so that PP/OPH etc. are updated correctly
              vAmount = FixTwoPlaces(vAmount + vOverPayment)
              vReceived = FixTwoPlaces(vReceived + vOverPayment)
            End If

            With vPP
              vRenewalAmount = .RenewalAmount
              If .PlanType = CDBEnvironment.ppType.pptMember And .FixedRenewalCycle And .PreviousRenewalCycle And (.ProportionalBalanceSetting And (PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew)) > 0 And .MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfAnnualTerm Then
                If CDate(.RenewalDate) = CDate(.StartDate).AddYears(.Term) Then
                  'Pro-rated membership within 1st year; RenewalAmount = full annual amount
                  'Set vRenewalAmount to the amount expected to be paid this year
                  vRenewalAmount = .GetProrataBalance(.RenewalAmount, .Member.Joined)
                End If
              End If
            End With

            If (vRenewalAmount <= FixTwoPlaces(vPP.Balance - (vAmount + vWriteOff))) AndAlso vPP.PlanType <> CDBEnvironment.ppType.pptLoan Then
              vArrears = FixTwoPlaces((vPP.Balance - (vAmount + vWriteOff)) - vPP.RenewalAmount)
              Dim vPPArrears As Double = FixTwoPlaces(vPP.Balance - vPP.RenewalAmount)
              If vArrears <> 0 And (FixTwoPlaces(System.Math.Abs(vArrears) - vPPArrears) = System.Math.Abs(vAmount)) Then
                'Amount was only paying Arrears and so RenewalDate & NPD Date does not need to be rolled back
                vMoveDates = False
              Else
                'Amount paid off the Arrears and some of the new Balance so RenewalDate needs to be rolled back
                If vPP.ContinuousRenewals = False Then UpdateRenewalDate(pBT, pBTA, vPP, False)
                If vArrears <> 0 Then
                  If FixTwoPlaces(System.Math.Abs(vAmount) - vArrears) < vPP.FrequencyAmount Then
                    'We paid more than the Arrears but less than the next FrequencyAmount so NPD Date does not need to be rolled back
                    vMoveDates = False
                  End If
                End If
              End If
            End If
            vPP.Balance = vPP.Balance - (vAmount + vWriteOff)
            If vPP.Balance > vPP.RenewalAmount Then
              If vPP.PlanType = CDBEnvironment.ppType.pptLoan Then
                'A Loan cannot go into arrears so increase the RenewalAmount
                vPP.RenewalAmount = vPP.Balance
              Else
                vPP.Arrears = vPP.Balance - vPP.RenewalAmount
              End If
            End If
          Else
            vPP.InAdvance = vPP.InAdvance + vAmount
            ReverseFinancialDetails(pBT, pBTA, vAmount)
            WriteOrderPaymentHistory(pBTA, vPP, vAmount, 0, False, vOrigOPSClaimDate)
            ReversePaymentPlanHistoryDetails(vPP, pBTA, vAmount)
          End If
        End If
        If vError = "" Then
          'Set last payment and date
          If mvBatchType <> BatchTypes.None Then
            If Len(vPP.LastPaymentDate) = 0 Then
              vPP.LastPayment = vReceived.ToString
              vPP.LastPaymentDate = pBT.TransactionDate
            Else
              If CDate(vPP.LastPaymentDate) <= CDate(pBT.TransactionDate) Then
                vPP.LastPayment = vReceived.ToString
                vPP.LastPaymentDate = pBT.TransactionDate
              End If
            End If
          End If
          If Not vInAdvance Then
            If (vPP.CovenantStatus <> PaymentPlan.ppCovenant.ppcDepositedDeed) Then
              If Not ((mvMemTypeTerm = MembershipType.MembershipTypeTerms.mtfMonthlyTerm Or vPP.FixedDDClaimDate) And vAmount < 0) And (vMoveDates Or (Not vMoveDates And vPP.Balance = 0)) Then
                UpdateNextPaymentDue(pBTA, vPP)
              End If
            Else
              If mvEnv.GetConfigOption("cv_deposited_deed_processing") Then
                vPP.Amount = ""
                vPP.Arrears = 0
                vPP.InAdvance = vPP.RenewalAmount
                vPP.LastPaymentDate = TodaysDate()
                vPP.StartDate = vPP.ExpiryDate
                mvDeedOrder = True
              End If
            End If

            If mvBatchType <> BatchTypes.None Then WriteOrderPaymentHistory(pBTA, vPP, vReceived, vWriteOff, vReCalc, vOrigOPSClaimDate)
            'note that process_membership_payments looks at payment_number
            If vReceived > 0 AndAlso pBTA.WriteOffLineAmount <> 0 Then vPP.Balance = vPP.Balance - pBTA.WriteOffLineAmount

            If vPP.PlanType = CDBEnvironment.ppType.pptMember And mvBatchType <> BatchTypes.None Then
              vExpiryDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Month, vPP.MembershipType.SuspensionGrace, CDate(vPP.RenewalDate))))
              vMemsPerOrder = vPP.MembershipType.MembersPerOrder
              ProcessMembershipPayment(pBT, pBTA, vPP, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPP.RenewalDate))), vExpiryDate, vAmount + vDiscount)
              If vPP.GiftMembership AndAlso vPP.OneYearGift Then
                If vMemsPerOrder = 0 Then
                  'group mem'ship
                  If Len(mvGOYGReason) = 0 Then
                    mvGOYGReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneYearGiftedGroupReason)
                    mvEnv.GetCancellationInfo(mvGOYGReason, mvGOYGStatus, mvGOYGDesc)
                  End If
                  vPP.Cancel(PaymentPlan.PaymentPlanCancellationTypes.pctPaymentPlan, mvGOYGReason, mvGOYGStatus, mvGOYGDesc, "automatic", "")
                ElseIf vPP.CancelOneYearGiftApm AndAlso vPP.AutoPaymentStatus AndAlso vPP.Balance <= vPP.RenewalAmount Then
                  'BR16375: For non-group one year gift memberships where the 'cancel_one_year_gift_apm' flag is set cancel the auto payment method
                  If mvNonGroupOYGAutoPayReason.Length = 0 Then
                    mvNonGroupOYGAutoPayReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneYearGiftsAutoReason)
                    mvEnv.GetCancellationInfo(mvNonGroupOYGAutoPayReason, mvNonGroupOYGAutoPayStatus, mvNonGroupOYGAutoPayDesc)
                  End If
                  'Unset Gift Membership and One Year Gift flags now that Auto Payment Method has been cancelled
                  With vPP
                    .SetMember(.MembershipTypeCode, .Branch, False, False)
                  End With
                  vPP.Cancel(PaymentPlan.PaymentPlanCancellationTypes.pctAutoPayment, mvNonGroupOYGAutoPayReason, mvNonGroupOYGAutoPayStatus, mvNonGroupOYGAutoPayDesc, "automatic", "")
                End If
              End If
              vPP.NumberOfReminders = 0
            Else
              ProcessPaymentPlanDetails(pBT, pBTA, vPP, vAmount + vDiscount, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPP.RenewalDate))))
              If vPP.OneOffPayment = True Then
                If Len(mvOOPReason) = 0 Then
                  mvOOPReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneOffPPCancelReason)
                  mvEnv.GetCancellationInfo(mvOOPReason, mvOOPStatus, mvOOPDesc)
                End If
                ' BR12005 - Need to check if the payment plan is already cancelled before trying to do it again
                If Len(vPP.CancellationReason) = 0 Then
                  vPP.Cancel(PaymentPlan.PaymentPlanCancellationTypes.pctPaymentPlan, mvOOPReason, mvOOPStatus, mvOOPDesc, "automatic", "")
                End If
              End If
            End If
          End If
          vLastRenewalDate = vPP.RenewalDate
          'if the order_date has been advanced, the renewal is no longer pending
          'if the order_date has been set back, the renewal is once again pending
          If CDate(vLastRenewalDate) > CDate(mvOrgRenewalDate) Then
            vPP.RenewalPending = False
            'if membership order renewal date is advanced
            'then if members reprint_mship_card flag is set, change to null
            If vPP.PlanType = CDBEnvironment.ppType.pptMember Then
              For Each vMember In vPP.CurrentMembers
                With vMember
                  If .ReprintMshipCard = True Then .ReprintMshipCard = False 'Was set to null
                  If Not mvEnv.GetConfigOption("me_set_card_expiry") Then .SetMembershipCardIssueNumber(Member.SetCardIssueNumberTypes.scintReinitialise) 'Reinitialise the membership card issue number, if necessary
                End With
              Next vMember
            End If
          End If
          If CDate(vLastRenewalDate) < CDate(mvOrgRenewalDate) Then vPP.RenewalPending = True
          For Each vMember In vPP.CurrentMembers
            vMember.SaveChanges("", mvEnv.AuditStyle = CDBEnvironment.AuditStyleTypes.ausExtended)
          Next vMember
          For Each vPPD In vPP.Details
            vPPD.SaveChanges()
          Next vPPD

          'Refresh the payment schedule if required
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then
            If vPP.PlanType = CDBEnvironment.ppType.pptLoan Then
              vPP.RegenerateLoanScheduledPayments(pAmendedBy, pAudit)
            Else
              If vAmount > 0 _
                   Or ForceCreationOfRegularProvisionalPayment(pBT, vPP) Then
                'Only for payments (not reversals) and NOT for Loans
                RecreatePaymentSchedule(pBT, pBTA, vPP, vInAdvance)
              End If
            End If
          End If
          If vPPDatesOveridden Then
            vPP.RenewalDate = vOverriddenPPRenewalDate
            vPP.NextPaymentDue = vOverriddenPPNextPaymentDue
          End If

          'If continuous renewals and pp has regular payments, recalculate renewal date if required. 
          ' (this is the case for skipped payments)
          If vPP.ContinuousRenewals _
            And vPP.PaymentFrequencyFrequency = 1 _
            And vPP.NextPaymentDue > vPP.RenewalDate Then
            vPP.RenewalDate = vPP.CalculateRenewalDate(vPP.RenewalDate, True)
          End If

          vPP.SaveChanges(pAmendedBy, pAudit)
        End If
      End If
    End Sub
    Private Sub ProcessPaymentPlanDetails(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double, ByRef pExpiry As String)
      If pPP.PlanType = CDBEnvironment.ppType.pptLoan Then
        If pAmount >= 0 Then
          ProcessLoanDetailsPayment(pBT, pBTA, pPP, pAmount, pExpiry)
        Else
          ReverseFinancialDetails(pBT, pBTA, pAmount)
          ProcessLoanDetailsReversal(pBT, pBTA, pPP, pAmount, pExpiry)
          ReversePaymentPlanHistoryDetails(pPP, pBTA, pAmount)
        End If
      Else
        If pAmount >= 0 Then
          ProcessPaymentPlanDetailsPayment(pBT, pBTA, pPP, pAmount, pExpiry)
        Else
          ReverseFinancialDetails(pBT, pBTA, pAmount)
          ProcessPaymentPlanDetailsReversal(pBT, pBTA, pPP, pAmount, pExpiry)
          ReversePaymentPlanHistoryDetails(pPP, pBTA, pAmount)
        End If
      End If
    End Sub
    Private Sub ProcessPaymentPlanDetailsArrears(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double)
      Dim vPPD As PaymentPlanDetail
      Dim vAmount As Double
      Dim vVatRate As VatRate
      Dim vFound As Boolean
      Dim vProportion As Integer
      Dim vPaymentPlanShowPaymentDetails As Boolean = mvEnv.GetConfigOption("fp_pp_show_payment_details", False)

      Dim vPayerVATCategory As String = pPP.Payer.VATCategory
      If mvEnv.GetConfigOption("fp_pay_proportional_details", False) Then
        'Pay off the arrears proportionally across all lines
        vProportion = pPP.PaymentFrequencyFrequency
        For Each vPPD In pPP.Details
          With vPPD
            If pAmount <> 0 Then
              If .Arrears <> 0 Then
                vFound = True
                If .Amount <> "" Then
                  vAmount = CDbl(.Amount)
                Else
                  vAmount = .CurrentPrice * .Quantity
                End If
                If vAmount <> 0 Then vAmount = FixTwoPlaces(vAmount / vProportion)
                If vAmount > pAmount Then vAmount = pAmount
                If vAmount > .Arrears Then
                  vAmount = .Arrears
                  .Arrears = 0
                Else
                  .Arrears = FixTwoPlaces(.Arrears - vAmount)
                End If
                .Balance = FixTwoPlaces(.Balance - vAmount)
                If mvBatchType <> BatchTypes.None Then
                  If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
                  vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATCategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, (pPP.GiverContactNumber))
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                End If
                pAmount = FixTwoPlaces(pAmount - vAmount)
                If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                  WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, vAmount)
                End If
              End If
            End If
          End With
        Next vPPD

        'Check that all money has been allocated
        If pAmount <> 0 Then
          For Each vPPD In pPP.Details
            With vPPD
              If pAmount <> 0 Then
                If .Arrears <> 0 Then
                  vFound = True
                  vAmount = pAmount
                  If vAmount > .Arrears Then
                    vAmount = .Arrears
                    .Arrears = 0
                  Else
                    .Arrears = FixTwoPlaces(.Arrears - vAmount)
                  End If
                  .Balance = FixTwoPlaces(.Balance - vAmount)
                  pAmount = FixTwoPlaces(pAmount - vAmount)
                  If mvBatchType <> BatchTypes.None Then
                    If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
                    vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATCategory, pBT.ContactVatCategory).ToString)
                    Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, (pPP.GiverContactNumber))
                    If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                  End If
                  If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                    WriteBranchIncome(pBT, pBTA, pPP, vPPD, True, vAmount)
                  End If
                End If
              End If
            End With
          Next vPPD
        End If
      Else
        'Pay off all the arrears 1 line at a time
        For Each vPPD In pPP.Details
          With vPPD
            If .Arrears <> 0 Then
              vFound = True
              If pAmount > 0 Then
                If pAmount > .Arrears Then
                  vAmount = .Arrears
                  pAmount = pAmount - .Arrears
                  .Arrears = 0
                Else
                  vAmount = pAmount
                  .Arrears = .Arrears - pAmount
                  pAmount = 0
                End If
                .Balance = .Balance - vAmount
                If mvBatchType <> BatchTypes.None Then
                  If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
                  vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATCategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, (pPP.GiverContactNumber))
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                End If
                If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                  WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, vAmount)
                End If
              End If
            End If
          End With
        Next vPPD
      End If

      If Not vFound Then PrintLog("Error : No order_details record with arrears for order: " & pPP.PlanNumber)
      If pAmount <> 0 And mvBatchType <> BatchTypes.None Then
        ProcessProduct(pBT, pBTA, (mvCompanyControl.DetailsProduct), (mvCompanyControl.DetailsRate), "", 1, pAmount, 1, "", 0)
      End If
    End Sub
    Private Function ProcessPaymentPlanDetailsDiscounts(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan) As Double
      Dim vPPD As PaymentPlanDetail
      Dim vTotalDiscount As Double
      Dim vVatRate As VatRate
      Dim vProportion As Integer
      Dim vAmount As Double
      Dim vPaymentPlanShowPaymentDetails As Boolean = mvEnv.GetConfigOption("fp_pp_show_payment_details", False)

      Dim vPayerVATCategory As String = pPP.Payer.VATCategory
      If (Not pPP.Balance.Equals(pBT.Amount)) AndAlso mvEnv.GetConfigOption("fp_pay_proportional_details", False) Then
        'Allocate the discounts proportionally
        If (Not String.IsNullOrEmpty(pPP.MembershipTypeCode)) AndAlso (Not String.IsNullOrEmpty(mvEnv.GetConfig("fixed_cycle_M"))) And pBTA.ScheduledPayment.IsInFirstYearOfSchedule Then
          'Fixed cycle Membership in the first year, PaymentFrequencyFrequency is unlikely to be the same as the actual number of payments
          vProportion = CInt((pPP.FrequencyAmount / pBT.Amount) * pPP.PaymentFrequencyFrequency)
        Else
          vProportion = pPP.PaymentFrequencyFrequency
        End If
        For Each vPPD In pPP.Details
          If vPPD.Balance < 0 Then
            If vPPD.Amount <> "" Then
              vAmount = CDbl(vPPD.Amount)
            Else
              vAmount = vPPD.CurrentPrice * vPPD.Quantity
            End If
            vAmount = FixTwoPlaces(vAmount / vProportion)
            If mvBatchType <> BatchTypes.None Then
              If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
              vVatRate = mvEnv.VATRate(vPPD.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATCategory, pBT.ContactVatCategory).ToString)
              Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, vPPD.DistributionCode, pBTA.DistributionCode)
              WriteFinancialHistoryAnalysis(pBT, pBTA, vPPD.ProductCode, vPPD.RateCode, vDistributionCode, CInt(vPPD.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), vPPD.Source)
              If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
            End If
            vPPD.Balance = FixTwoPlaces(vPPD.Balance - vAmount)
            If vPPD.Arrears <> 0 Then
              If vAmount > vPPD.Arrears Then
                vPPD.Arrears = 0
              Else
                vPPD.Arrears = FixTwoPlaces(vPPD.Arrears - vAmount)
              End If
            End If
            vTotalDiscount = vTotalDiscount - vAmount
            vPPD.SaveChanges()
          End If
        Next vPPD
      Else
        'Remove all discounts
        For Each vPPD In pPP.Details
          If vPPD.Balance < 0 Then
            If mvBatchType <> BatchTypes.None Then
              If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
              vVatRate = mvEnv.VATRate(vPPD.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATCategory, pBT.ContactVatCategory).ToString)
              Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, vPPD.DistributionCode, pBTA.DistributionCode)
              WriteFinancialHistoryAnalysis(pBT, pBTA, (vPPD.ProductCode), vPPD.RateCode, vDistributionCode, CInt(vPPD.Quantity), vPPD.Balance, vVatRate.VatRateCode, CalculateVATAmount((vPPD.Balance), vVatRate.CurrentPercentage(pBT.TransactionDate)), (vPPD.Source))
              If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vPPD.Balance)
            End If
            vTotalDiscount = vTotalDiscount - vPPD.Balance
            vPPD.Balance = 0
            vPPD.Arrears = 0
            vPPD.SaveChanges()
          End If
        Next vPPD
      End If
      ProcessPaymentPlanDetailsDiscounts = vTotalDiscount
    End Function
    Private Sub ProcessPaymentPlanDetailsPayment(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double, ByRef pExpiry As String)
      Dim vPPD As PaymentPlanDetail
      Dim vVatRate As VatRate
      Dim vAmount As Double
      Dim vProportion As Integer
      Dim vBTAWriteOffAmountRemaining As Double = pBTA.WriteOffLineAmount
      Dim vPaymentPlanShowPaymentDetails As Boolean = mvEnv.GetConfigOption("fp_pp_show_payment_details", False)

      If pPP.Details.Count() = 0 Then
        PrintLog("Error : Order_details record missing - payment for order: " & pPP.PlanNumber)
      Else
        Dim vPayerVATcategory As String = pPP.Payer.VATCategory
        ' The payment has already beeen subtracted from the Paymnent Plan Balance, so the last payment will have PP Balance = 0, no prorating for last payment 
        If (mvEnv.GetConfigOption("fp_pay_proportional_details", False)) And pPP.Details.Count() > 1 AndAlso (Not pPP.Balance.Equals(0)) Then
          'Pay the Details lines proportionately
          If (Not String.IsNullOrEmpty(pPP.MembershipTypeCode)) AndAlso (Not String.IsNullOrEmpty(mvEnv.GetConfig("fixed_cycle_M"))) And pBTA.ScheduledPayment.IsInFirstYearOfSchedule Then
            'Fixed cycle Membership in the first year, PaymentFrequencyFrequency is unlikely to be the same as the actual number of payments
            vProportion = CInt((pPP.FrequencyAmount / pBT.Amount) * CDbl(pPP.PaymentFrequencyFrequency))
          Else
            vProportion = pPP.PaymentFrequencyFrequency
          End If
          For Each vPPD In pPP.Details
            If pAmount <> 0 Then
              If vPPD.Balance <> 0 Then
                If vPPD.Amount <> "" Then
                  vAmount = CDbl(vPPD.Amount)
                Else
                  vAmount = vPPD.Quantity * vPPD.CurrentPrice
                End If
                If vAmount <> 0 Then
                  vAmount = FixTwoPlaces(vAmount / CDbl(vProportion)) 'The proportionate amount to be paid each time
                End If
                If Math.Abs(vAmount) >= Math.Abs(vPPD.Balance) Then 'vAmount and vPPD.Balance will have the same sign, and can be -ve, -1 > -10, but we want -10 > -1 so 
                  vAmount = vPPD.Balance
                  vPPD.Balance = 0
                  pAmount = FixTwoPlaces(pAmount - vAmount) 'pAmount now amount o/s after this allocation
                Else
                  vPPD.Balance = FixTwoPlaces(vPPD.Balance - vAmount)
                  pAmount = FixTwoPlaces(pAmount - vAmount) 'pAmount now amount o/s after this allocation
                  If pBTA.AcceptAsFull = True Then
                    vPPD.Balance = 0
                    pAmount = 0
                  ElseIf vBTAWriteOffAmountRemaining > 0 Then
                    'If BTA.WriteOffAmount set then reduce the balance from this write off amount
                    If vBTAWriteOffAmountRemaining >= vPPD.Balance Then
                      vBTAWriteOffAmountRemaining = FixTwoPlaces(vBTAWriteOffAmountRemaining - vPPD.Balance)
                      vPPD.Balance = 0
                    Else
                      vPPD.Balance = FixTwoPlaces(vPPD.Balance - vBTAWriteOffAmountRemaining)
                      vBTAWriteOffAmountRemaining = 0
                    End If
                  End If
                End If
                If mvBatchType <> BatchTypes.None Then
                  If vPPD.ProductRateIsValid = False Then
                    vPPD.SetPrices()
                  End If
                  vVatRate = mvEnv.VATRate(vPPD.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, vPPD.DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, vPPD.ProductCode, vPPD.RateCode, vDistributionCode, CInt(vPPD.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), vPPD.Source, (pPP.GiverContactNumber))
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then
                    WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                  End If
                End If
              End If
            Else
              If pBTA.AcceptAsFull = True And vPPD.Balance > 0 Then
                vPPD.Balance = 0
              ElseIf vBTAWriteOffAmountRemaining > 0 AndAlso vPPD.Balance > 0 Then
                'If BTA.WriteOffAmount set then reduce the balance from this write off amount
                If vBTAWriteOffAmountRemaining >= vPPD.Balance Then
                  vBTAWriteOffAmountRemaining = FixTwoPlaces(vBTAWriteOffAmountRemaining - vPPD.Balance)
                  vPPD.Balance = 0
                Else
                  vPPD.Balance = FixTwoPlaces(vPPD.Balance - vBTAWriteOffAmountRemaining)
                  vBTAWriteOffAmountRemaining = 0
                End If
              End If
            End If
            If vPPD.Product.Subscription = True And mvBatchType <> BatchTypes.None Then
              Select Case CheckProcessSubscriptions(pPP, vPPD)
                Case "Y"
                  ProcessSubscriptions(pPP, vPPD, vPPD.ContactNumber, vPPD.ProductCode, pExpiry)
                Case "T"
                  TerminateSubscriptions((pPP.PlanNumber), vPPD.ContactNumber, vPPD.ProductCode, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
                Case "D"
                  DeleteSubscriptions((pPP.PlanNumber), vPPD.ContactNumber, vPPD.ProductCode)
                Case Else
                  'Do nothing
              End Select
            End If
            If vPPD.ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
              WriteBranchIncome(pBT, pBTA, pPP, vPPD, True, vAmount)
            End If
          Next vPPD

          'Check that all money has been allocated (anything here should just be rounding errors)
          If pAmount <> 0 Then
            For Each vPPD In pPP.Details
              If pAmount <> 0 Then
                If vPPD.Balance <> 0 Then
                  vAmount = pAmount
                  If vAmount >= vPPD.Balance Then
                    vAmount = vPPD.Balance
                    vPPD.Balance = 0
                  Else
                    vPPD.Balance = FixTwoPlaces(vPPD.Balance - vAmount)
                  End If
                  pAmount = FixTwoPlaces(pAmount - vAmount)
                  If mvBatchType <> BatchTypes.None Then
                    If vPPD.ProductRateIsValid = False Then
                      vPPD.SetPrices()
                    End If
                    vVatRate = mvEnv.VATRate(vPPD.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                    Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, vPPD.DistributionCode, pBTA.DistributionCode)
                    If vAmount > 0 Then
                      WriteFinancialHistoryAnalysis(pBT, pBTA, vPPD.ProductCode, vPPD.RateCode, vDistributionCode, CInt(vPPD.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), vPPD.Source, (pPP.GiverContactNumber))
                      If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then
                        WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                      End If
                    End If
                  End If
                  If vPPD.ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                    WriteBranchIncome(pBT, pBTA, pPP, vPPD, True, vAmount)
                  End If
                End If
              End If
            Next vPPD
          End If

        Else
          'Pay off the Detail lines 1 at a time
          For Each vPPD In pPP.Details
            If pAmount <> 0 Then
              If vPPD.Balance <> 0 Then
                If pAmount > vPPD.Balance Then
                  vAmount = vPPD.Balance
                  pAmount = FixTwoPlaces(pAmount - vPPD.Balance)
                  vPPD.Balance = 0
                Else
                  vAmount = pAmount
                  vPPD.Balance = FixTwoPlaces(vPPD.Balance - pAmount)
                  pAmount = 0
                  If pBTA.AcceptAsFull = True Then
                    vPPD.Balance = 0
                  ElseIf vBTAWriteOffAmountRemaining > 0 Then
                    'If BTA.WriteOffAmount set then reduce the balance from this write off amount
                    If vBTAWriteOffAmountRemaining >= vPPD.Balance Then
                      vBTAWriteOffAmountRemaining = FixTwoPlaces(vBTAWriteOffAmountRemaining - vPPD.Balance)
                      vPPD.Balance = 0
                    Else
                      vPPD.Balance = FixTwoPlaces(vPPD.Balance - vBTAWriteOffAmountRemaining)
                      vBTAWriteOffAmountRemaining = 0
                    End If
                  End If
                End If
                If mvBatchType <> BatchTypes.None Then
                  If vPPD.ProductRateIsValid = False Then
                    vPPD.SetPrices()
                  End If
                  vVatRate = mvEnv.VATRate(vPPD.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, vPPD.DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, vPPD.ProductCode, vPPD.RateCode, vDistributionCode, CInt(vPPD.Quantity), vAmount, vVatRate.VatRateCode, CalculateVATAmount(vAmount, vVatRate.CurrentPercentage(pBT.TransactionDate)), vPPD.Source, (pPP.GiverContactNumber))
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then
                    WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vAmount)
                  End If
                End If
              End If
            Else
              If pBTA.AcceptAsFull = True And vPPD.Balance > 0 Then
                vPPD.Balance = 0
              ElseIf vBTAWriteOffAmountRemaining > 0 AndAlso vPPD.Balance > 0 Then
                'If BTA.WriteOffAmount set then reduce the balance from this write off amount
                If vBTAWriteOffAmountRemaining >= vPPD.Balance Then
                  vBTAWriteOffAmountRemaining = FixTwoPlaces(vBTAWriteOffAmountRemaining - vPPD.Balance)
                  vPPD.Balance = 0
                Else
                  vPPD.Balance = FixTwoPlaces(vPPD.Balance - vBTAWriteOffAmountRemaining)
                  vBTAWriteOffAmountRemaining = 0
                End If
              End If
            End If
            If vPPD.Product.Subscription = True And mvBatchType <> BatchTypes.None Then
              Select Case CheckProcessSubscriptions(pPP, vPPD)
                Case "Y"
                  ProcessSubscriptions(pPP, vPPD, vPPD.ContactNumber, vPPD.ProductCode, pExpiry)
                Case "T"
                  TerminateSubscriptions((pPP.PlanNumber), vPPD.ContactNumber, vPPD.ProductCode, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
                Case "D"
                  DeleteSubscriptions((pPP.PlanNumber), vPPD.ContactNumber, vPPD.ProductCode)
                Case Else
                  'Do nothing
              End Select
            End If
            If vPPD.ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
              WriteBranchIncome(pBT, pBTA, pPP, vPPD, True, vAmount)
            End If
          Next vPPD
        End If
      End If
      If pAmount <> 0 Then
        ProcessProduct(pBT, pBTA, (mvCompanyControl.DetailsProduct), (mvCompanyControl.DetailsRate), "", 1, pAmount, 1, "", 0)
      End If
    End Sub
    Private Sub ProcessPaymentPlanDetailsReversal(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double, ByRef pExpiry As String)
      Dim vPPD As PaymentPlanDetail
      Dim vCost As Double
      Dim vStartAmount As Double
      Dim vNumber As Integer
      Dim vProportion As Integer

      pAmount = -pAmount

      pAmount = pAmount + GetOriginalWriteOff((pBTA.BatchNumber), (pBTA.TransactionNumber), (pBTA.LineNumber), pBTA.AcceptAsFull)

      vNumber = pPP.Details.Count()
      If vNumber = 0 Then PrintLog("Error : Order_details record missing - reversal for order: " & pPP.PlanNumber)

      If mvEnv.GetConfigOption("fp_pay_proportional_details", False) Then
        'Reverse the payment proportionally across all detail lines
        vProportion = pPP.PaymentFrequencyFrequency
        For Each vPPD In pPP.Details
          With vPPD
            If pAmount <> 0 Then
              If .Amount <> "" Then
                vCost = CDbl(.Amount)
                If vCost > 0 Then vCost = FixTwoPlaces(vCost / vProportion)
                If vCost > pAmount Then vCost = pAmount 'vCost is what we were expecting to receive
                If vCost >= 0 Then
                  If .Balance < CDbl(.Amount) Then
                    If vCost > (CDbl(.Amount) - .Balance) Then
                      vCost = FixTwoPlaces(CDbl(.Amount) - .Balance)
                      .Balance = FixTwoPlaces(CDbl(.Amount))
                    Else
                      .Balance = FixTwoPlaces(.Balance + vCost)
                    End If
                  End If
                End If
              Else
                vCost = .Quantity * .CurrentPrice
                If vCost <> 0 Then vCost = FixTwoPlaces(vCost / vProportion) 'vCost is what we were expecting to receive
                If vCost > pAmount Then vCost = pAmount
                If vCost >= 0 Then
                  If .CurrentPrice <> 0 Then
                    If vCost > ((.CurrentPrice * .Quantity) - .Balance) Then
                      vCost = FixTwoPlaces((.CurrentPrice * .Quantity) - .Balance)
                      .Balance = FixTwoPlaces(.CurrentPrice * .Quantity)
                    Else
                      .Balance = FixTwoPlaces(.Balance + vCost)
                    End If
                  End If
                Else
                  'discount pricing line - so increase amount by discount
                  'need to do only if balance made >= renewal_amount
                  If vCost < (.CurrentPrice * .Quantity) - .Balance Then
                    vCost = FixTwoPlaces((.CurrentPrice * .Quantity) - .Balance)
                    .Balance = FixTwoPlaces(.CurrentPrice * .Quantity)
                  Else
                    .Balance = FixTwoPlaces(.Balance + vCost)
                  End If
                End If
              End If
              pAmount = FixTwoPlaces(pAmount - vCost)
            End If
            If .Product.Subscription = True Then
              'add check for future membership subscriptions
              Select Case CheckProcessSubscriptions(pPP, vPPD)
                Case "Y"
                  ProcessSubscriptions(pPP, vPPD, .ContactNumber, .ProductCode, pExpiry)
                Case "T"
                  TerminateSubscriptions((pPP.PlanNumber), .ContactNumber, .ProductCode, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
                Case "D"
                  DeleteSubscriptions((pPP.PlanNumber), .ContactNumber, .ProductCode)
                Case Else
                  'do nothing
              End Select
            End If
            If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
              WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
            End If
          End With
        Next vPPD

        'Check that all money has been allocated
        If pAmount <> 0 Then
          For Each vPPD In pPP.Details
            With vPPD
              If pAmount <> 0 Then
                If .Balance <> 0 Then
                  vCost = pAmount
                  If vCost > .Balance Then
                    vCost = .Balance
                    .Balance = 0
                  Else
                    .Balance = FixTwoPlaces(.Balance + vCost)
                  End If
                  pAmount = FixTwoPlaces(pAmount - vCost)
                End If
              End If
              If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
              End If
            End With
          Next vPPD
        End If

      Else
        'select in reverse order because we want to allocate back in reverse of the order we would have allocated
        If vNumber > 0 And pPP.PlanType = CDBEnvironment.ppType.pptMember Then
          If pPP.FixedRenewalCycle And pPP.PreviousRenewalCycle And pPP.MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfAnnualTerm _
          And (pPP.ProportionalBalanceSetting And (PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
          And pPP.DetermineMembershipPeriod = PaymentPlan.MembershipPeriodTypes.mptFirstPeriod Then
            'Pro-rated membership in first period.
            'As Balance will have been proportional we need to allocate the payment back to the original PPD lines (as far as we can)
            'In all other cases we will always use the existing logic
            Dim vFHD As New FinancialHistoryDetail(mvEnv)
            vFHD.Init(mvEnv)
            Dim vWhereFields As New CDBFields()
            With vWhereFields
              .Add("r.batch_number", pBTA.BatchNumber)
              .Add("r.transaction_number", pBTA.TransactionNumber)
              .Add("r.line_number", pBTA.LineNumber)
              .Add("was_batch_number", CDBField.FieldTypes.cftInteger, "fhd.batch_number")
              .Add("was_transaction_number", CDBField.FieldTypes.cftInteger, "fhd.transaction_number")
              .Add("was_line_number", CDBField.FieldTypes.cftInteger, "fhd.line_number")
            End With
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFHD.GetRecordSetFields(FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll), "reversals r, financial_history_details fhd", vWhereFields)
            Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
            While vRS.Fetch = True
              vFHD = New FinancialHistoryDetail(mvEnv)
              vFHD.InitFromRecordSet(mvEnv, vRS, FinancialHistoryDetail.FinancialHistoryDetailRecordSetTypes.fhdrtAll)
              Dim vPPDAllocated As Boolean = False
              Dim vAmount As Double = vFHD.Amount
              For Each vPPD In pPP.Details
                If (vPPD.ProductCode = vFHD.ProductCode) And (vPPD.RateCode = vFHD.RateCode) Then
                  'Found the correct PPD
                  With vPPD
                    If .Amount.Length <> 0 Then
                      If DoubleValue(.Amount) <> 0 Then
                        If .Balance < DoubleValue(.Amount) Then
                          If vAmount > (DoubleValue(.Amount) - .Balance) Then
                            vAmount = vAmount - (DoubleValue(.Amount) - .Balance)
                            pAmount = pAmount - (DoubleValue(.Amount) - .Balance)
                            .Balance = DoubleValue(.Amount)
                          Else
                            .Balance = .Balance + vAmount
                            pAmount = FixTwoPlaces(pAmount - vAmount)
                            vAmount = 0
                          End If
                          vPPDAllocated = True
                        End If
                      End If
                    Else
                      If .CurrentPrice <> 0 Then
                        vCost = FixTwoPlaces(.Quantity * .CurrentPrice)
                        If vCost >= 0 Then
                          If vCost > .Balance Then
                            If vAmount > (vCost - .Balance) Then
                              vAmount = vAmount - (vCost - .Balance)
                              pAmount = pAmount - (vCost - .Balance)
                              .Balance = vCost
                            Else
                              .Balance = .Balance + vAmount
                              pAmount = FixTwoPlaces(pAmount - vAmount)
                              vAmount = 0
                            End If
                            vPPDAllocated = True
                          End If
                        Else
                          'discount pricing line - so increase amount by discount
                          'need to do only if balance made >= renewal_amount
                          If pPP.Balance >= pPP.RenewalAmount And pPP.Balance - vAmount < pPP.RenewalAmount Then
                            If vCost < .Balance Then
                              vAmount = vAmount - (vCost - .Balance)
                              pAmount = pAmount - (vCost - .Balance)
                              .Balance = vCost
                              vPPDAllocated = True
                            End If
                          End If
                        End If
                      End If
                    End If
                    If .Product.Subscription = True Then
                      'add check for future membership subscriptions
                      Select Case CheckProcessSubscriptions(pPP, vPPD)
                        Case "Y"
                          ProcessSubscriptions(pPP, vPPD, .ContactNumber, .ProductCode, pExpiry)
                        Case "T"
                          TerminateSubscriptions(pPP.PlanNumber, .ContactNumber, .ProductCode, DateAdd("d", -1, mvOrgRenewalDate))
                        Case "D"
                          DeleteSubscriptions(pPP.PlanNumber, .ContactNumber, .ProductCode)
                        Case Else
                          'do nothing
                      End Select
                    End If
                    If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                      WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
                    End If
                  End With
                End If
                If vPPDAllocated = True And vAmount = 0 Then Exit For
              Next
            End While
            vRS.CloseRecordSet()
          End If
        End If

        If pAmount <> 0 Then
          While vNumber > 0
            vPPD = DirectCast(pPP.Details.Item(vNumber), PaymentPlanDetail)
            With vPPD
              If pAmount <> 0 Then
                If .Amount <> "" Then
                  If CDbl(.Amount) <> 0 Then
                    If .Balance < CDbl(.Amount) Then
                      If pAmount > (CDbl(.Amount) - .Balance) Then
                        pAmount = pAmount - (CDbl(.Amount) - .Balance)
                        .Balance = CDbl(.Amount)
                      Else
                        .Balance = .Balance + pAmount
                        pAmount = 0
                      End If
                    End If
                  End If
                Else
                  If .CurrentPrice <> 0 Then
                    vCost = FixTwoPlaces(.Quantity * .CurrentPrice)
                    If vCost >= 0 Then
                      If vCost > .Balance Then
                        If pAmount > (vCost - .Balance) Then
                          pAmount = pAmount - (vCost - .Balance)
                          .Balance = vCost
                        Else
                          .Balance = .Balance + pAmount
                          pAmount = 0
                        End If
                      End If
                    Else
                      'discount pricing line - so increase amount by discount
                      'need to do only if balance made >= renewal_amount
                      If pPP.Balance >= pPP.RenewalAmount And pPP.Balance - pAmount < pPP.RenewalAmount Then
                        If vCost < .Balance Then
                          pAmount = pAmount - (vCost - .Balance)
                          .Balance = vCost
                        End If
                      End If
                    End If
                  End If
                End If
              End If
              If .Product.Subscription = True Then
                'add check for future membership subscriptions
                Select Case CheckProcessSubscriptions(pPP, vPPD)
                  Case "Y"
                    ProcessSubscriptions(pPP, vPPD, .ContactNumber, .ProductCode, pExpiry)
                  Case "T"
                    TerminateSubscriptions((pPP.PlanNumber), .ContactNumber, .ProductCode, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvOrgRenewalDate)))
                  Case "D"
                    DeleteSubscriptions((pPP.PlanNumber), .ContactNumber, .ProductCode)
                  Case Else
                    'do nothing
                End Select
              End If
              If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
              End If
              vNumber = vNumber - 1
            End With
          End While
        End If
        'processed through once and everything should have been reset
        'however may have owed more than last renewal amount
        'and above will not cope with that situation
        vStartAmount = 0
        While pAmount <> 0 And vStartAmount <> pAmount
          vStartAmount = pAmount
          For Each vPPD In pPP.Details
            With vPPD
              If pAmount <> 0 Then
                If .Amount <> "" Then
                  If CDbl(.Amount) <> 0 Then
                    If pAmount > CDbl(.Amount) Then
                      pAmount = pAmount - CDbl(.Amount)
                      .Balance = .Balance + CDbl(.Amount)
                      .Arrears = .Arrears + CDbl(.Amount)
                    Else
                      .Balance = .Balance + pAmount
                      .Arrears = .Arrears + pAmount
                      pAmount = 0
                    End If
                  End If
                Else
                  If .CurrentPrice <> 0 Then
                    vCost = FixTwoPlaces(.Quantity * .CurrentPrice)
                    If vCost > 0 Then
                      If pAmount > vCost Then
                        .Balance = .Balance + vCost
                        .Arrears = .Arrears + vCost
                        pAmount = pAmount - vCost
                      Else
                        .Balance = .Balance + pAmount
                        .Arrears = .Arrears + pAmount
                        pAmount = 0
                      End If
                    End If
                  End If
                End If
                If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
                  WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
                End If
              End If
            End With
          Next vPPD
        End While
      End If 'from Config

      If pAmount <> 0 Then
        vPPD = DirectCast(pPP.Details.Item(1), PaymentPlanDetail)
        With vPPD
          .Balance = .Balance + pAmount
          'this will be run if amount = null & current price = 0
          'so changed to only add arrears if orders.arrears > 0
          If pPP.Arrears > 0 Then
            If pPP.Arrears < pAmount Then
              .Arrears = .Arrears + pPP.Arrears
            Else
              .Arrears = .Arrears + pAmount
            End If
          End If
          If .ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchProduct) And mvBranchIncomePeriod = "LAST" Then
            WriteBranchIncome(pBT, pBTA, pPP, vPPD, False, pAmount)
          End If
        End With
      End If
    End Sub
    Private Function ProcessProduct(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pProduct As Product, ByRef pRate As String, ByRef pDistributionCode As String, ByRef pQuantity As Integer, ByVal pAmount As Double, ByRef pIssued As Integer, ByRef pVATRate As String, ByRef pVATAmount As Double) As Double
      Dim vPosted As Double
      If pProduct IsNot Nothing Then
        If Not mvConn.InTransaction Then mvConn.StartTransaction() 'TRANSACTION START HERE

        With pProduct
          'Check for Gift Aid
          If .SponsorshipEvent And pBT.EligibleForGiftAid And pBTA.LineType <> "B" And CDate(pBT.TransactionDate) >= CDate(mvGAOperationalChangeDate) And pBT.PaymentMethod <> mvCAFPaymentMethod Then
            ProcessGiftAidSponsorship(pBT, pBTA, (pBTA.Amount))
          End If
          'Check for Legacy Bequest Receipt
          If pBTA.LineType = "B" And pBTA.Amount < 0 Then
            ProcessCheckLegacyReceipt(pBT, pBTA, (pBTA.Amount))
          End If
          Dim vProductVatRate As VatRate
          vProductVatRate = mvEnv.VATRate(.ProductVatCategory, (pBT.ContactVatCategory))
          If Not .StockItem Then
            If Len(pVATRate) = 0 Then
              WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pQuantity, pAmount, vProductVatRate.VatRateCode, CalculateVATAmount(pAmount, vProductVatRate.CurrentPercentage(pBT.TransactionDate)))
            Else
              WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pQuantity, pAmount, pVATRate, pVATAmount)
            End If
            vPosted = pAmount
            'BR13623: Allocate the amount if the analysis line is linked to a Fundraising Payment Schedule
            If pBTA.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment _
            AndAlso pBTA.AdditionalNumber > 0 Then
              Dim vFPS As New FundraisingPaymentSchedule(mvEnv)
              vFPS.Init(pBTA.AdditionalNumber)
              If vFPS.Existing Then vFPS.AllocateAmount(pAmount, pBT.TransactionDate, pBTA.Source)
            End If
          Else
            Dim vVatRate As String
            Dim vVatAmount As Double
            If Len(pVATRate) = 0 Then
              vVatRate = vProductVatRate.VatRateCode
              vVatAmount = CalculateVATAmount(pAmount, vProductVatRate.CurrentPercentage(pBT.TransactionDate))
            Else
              vVatRate = pVATRate
              vVatAmount = pVATAmount
            End If

            Dim vQuantity As Integer
            If pQuantity > 0 Then
              vQuantity = 0
              vPosted = 0
              If Picked = "N" Then
                'Batch has not been picked & confirmed
                vQuantity = pQuantity
                If pIssued > 0 Then
                  'BR15041:    we shouldn't create issued stock record if we're not issuing the stock
                  'WriteIssuedStock(pBTA)
                  'The stock has already been decreased, but we've not issued the stock, so need to create a stock movement, to increase the stock amount
                  Dim vStockM As New StockMovement
                  vStockM.Create(mvEnv, pBTA.ProductCode, pBTA.Issued, "AJ", pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, , pBTA.Warehouse)
                  'This will also create BackOrders, so need to set BTA.Issued to zero as nothing has been issued
                  pBTA.Issued = 0
                  Dim vUpdateFields As CDBFields = New CDBFields
                  Dim vWhereFields As CDBFields = New CDBFields
                  vUpdateFields.Add("issued", CDBField.FieldTypes.cftLong, 0)
                  With vWhereFields
                    .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
                    .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
                    .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
                  End With
                  mvConn.UpdateRecords("batch_transaction_analysis", vUpdateFields, vWhereFields)
                End If
              Else
                vQuantity = pQuantity - pIssued
              End If
              If BatchType = BatchTypes.Cash OrElse BatchType = BatchTypes.CashWithInvoice Then
                WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pQuantity, pAmount, vVatRate, vVatAmount)
                vPosted = pAmount
              Else
                If Picked = "C" Then
                  If pIssued > 0 Then
                    vPosted = pAmount * pIssued / pQuantity
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pIssued, vPosted, vVatRate, vVatAmount * pIssued / pQuantity)
                  Else
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, 0, 0, vVatRate, 0)
                  End If
                Else
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, 0, 0, vVatRate, 0)
                End If
              End If
              If vQuantity > 0 Then
                mvCreateBackOrder = True
                WriteBackOrderDetails(pBTA, vVatRate, vVatAmount)
              End If
            Else
              'some form of adjustment
              vQuantity = 0
              vPosted = 0
              vQuantity = pQuantity - pIssued
              If BatchType = BatchTypes.Cash OrElse BatchType = BatchTypes.CashWithInvoice Or BatchType = BatchTypes.FinancialAdjustment Then
                WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pQuantity, pAmount, vVatRate, vVatAmount)
                vPosted = pAmount
              Else
                If pIssued <> 0 Then
                  vPosted = pAmount * pIssued / pQuantity
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, pIssued, vPosted, vVatRate, vVatAmount * pIssued / pQuantity)
                Else
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, pRate, pDistributionCode, 0, 0, vVatRate, 0)
                End If
              End If
              If vQuantity > 0 Then
                mvCreateBackOrder = True
                WriteBackOrderDetails(pBTA, vVatRate, vVatAmount)
              End If
            End If
          End If
          If Len(.Activity) > 0 And (pAmount > 0 Or (pAmount = 0 And mvEnv.GetConfigOption("fp_category_from_zero_bta") And Not pBT.IsFinancialAdjustment)) Then
            Dim vProductCatExtension As Boolean = mvEnv.GetConfigOption("opt_fp_product_cat_extension")
            ProcessCategories((pBT.ContactNumber), (pBT.ContactType), .Activity, .ActivityValue, (pBTA.Source), (pBT.TransactionDate), (pBT.TransactionDate), vProductCatExtension, False, 0, .ActivityDurationMonths)
          End If
        End With
      End If
      Return vPosted
    End Function

    Private Function GetStockStatus(ByRef pBT As BatchTransaction) As TransactionStockStatus
      Dim vBTA As BatchTransactionAnalysis
      Dim vTSS As TransactionStockStatus

      vTSS = TransactionStockStatus.tssNoStock
      For Each vBTA In pBT.Analysis
        If vBTA.Quantity > 0 And vBTA.Product.StockItem Then
          If vTSS = TransactionStockStatus.tssNoStock Then vTSS = TransactionStockStatus.tssStockNothingDespatched
          If vBTA.Issued > 0 Then vTSS = TransactionStockStatus.tssStockSomeDespatched
        End If
      Next vBTA
      GetStockStatus = vTSS
    End Function

    Public Function HasStockItem(ByRef pBatchNumber As Integer) As Boolean
      Dim vStockItem As Boolean
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("bta.product", CDBField.FieldTypes.cftLong, "p.product")
      vWhereFields.Add("bta.batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
      vWhereFields.Add("p.stock_item", CDBField.FieldTypes.cftCharacter, "Y")
      If (mvEnv.Connection.GetCount("batch_transaction_analysis bta, products p", vWhereFields)) > 0 Then
        vStockItem = True
      End If
      HasStockItem = vStockItem
    End Function

    Private Sub ProcessProductsToBackOrder(ByVal pConn As CDBConnection, ByRef pBT As BatchTransaction)
      Dim vWhereFields As New CDBFields
      Dim vFields As New CDBFields
      Dim vBTAFields As New CDBFields
      Dim vBTA As BatchTransactionAnalysis

      vFields.Add("quantity", CDBField.FieldTypes.cftLong, 0)
      vFields.Add("amount", CDBField.FieldTypes.cftNumeric, 0)
      vFields.Add("vat_amount", CDBField.FieldTypes.cftNumeric, 0)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vFields.Add("currency_amount", CDBField.FieldTypes.cftNumeric, 0)
        vFields.Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric, 0)
      End If
      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
      vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
      vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong)
      vBTAFields.Add("issued", CDBField.FieldTypes.cftLong, 0)

      For Each vBTA In pBT.Analysis
        'Does this really need to check to see if the p&p product amount is > 0 ?
        If ((vBTA.Product.PostagePacking = True) And (vBTA.Amount > 0)) Or ((vBTA.Product.Existing = True And vBTA.Product.StockItem = False) And (BatchType <> BatchTypes.CreditCard) And (BatchType <> BatchTypes.CreditCardWithInvoice) And (BatchType <> BatchTypes.DebitCard)) Then
          WriteBackOrderDetails(vBTA, (vBTA.VatRate), (vBTA.VatAmount), True, pBT)
          vWhereFields(3).Value = CStr(vBTA.LineNumber)
          If BatchType <> BatchTypes.Cash AndAlso BatchType <> BatchTypes.CashWithInvoice And vBTA.LineType <> "I" Then
            'Since these items are going onto back order take the amount off the BT so that FH will be written correctly
            pBT.Amount = pBT.Amount - vBTA.Amount
            pBT.CurrencyAmount = pBT.CurrencyAmount - vBTA.CurrencyAmount
            pConn.UpdateRecords("financial_history_details", vFields, vWhereFields)
          End If
          'Need to update the issued amount for P&P, but can not use vBTA.Save as other values may have changed
          If vBTA.Product.PostagePacking = True And vBTA.Issued <> 0 Then
            vBTA.Issued = 0
            pConn.UpdateRecords("batch_transaction_analysis", vBTAFields, vWhereFields)
          End If
        End If
      Next vBTA
    End Sub

    Public Sub ProcessSubscriptions(ByRef pPP As PaymentPlan, ByRef pPPD As PaymentPlanDetail, ByRef pContactNumber As Integer, ByRef pProduct As String, ByRef pExpiry As String)
      Dim vWhereFields As New CDBFields
      Dim vFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vSub As New Subscription

      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
      vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pProduct)
      If pContactNumber > 0 Then
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      Else
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong)
      End If

      vSub.Init(mvEnv)
      vRecordSet = mvConn.GetRecordSet("SELECT " & vSub.GetRecordSetFields(Subscription.SubscriptionRecordSetTypes.subrstAll) & " FROM subscriptions WHERE " & mvConn.WhereClause(vWhereFields))
      With vSub
        If vRecordSet.Fetch() = True Then
          .InitFromRecordSet(mvEnv, vRecordSet, Subscription.SubscriptionRecordSetTypes.subrstAll)
          .AddressNumber = pPPD.AddressNumber
          .Quantity = CInt(pPPD.Quantity)
          .SetValidTo((pPP.StartDate), pExpiry)
          'update for current subscription
          If CDate(.ValidTo) > Today Then .ReasonForDespatch = pPP.ReasonForDespatch 'If null will get set to control value
          .UnCancel()
          .Save("automatic")
        Else
          .Init(mvEnv)
          .PaymentPlanNumber = pPP.PlanNumber
          .ContactNumber = pContactNumber
          .AddressNumber = pPPD.AddressNumber
          .Product = pProduct
          .Quantity = CInt(pPPD.Quantity)
          If mvMeFutureChange = "Y" Then
            .ValidFrom = mvOrgRenewalDate
          Else
            .ValidFrom = pPP.StartDate
          End If
          .SetValidTo((pPP.StartDate), pExpiry)
          .DespatchMethod = pPPD.DespatchMethod
          .ReasonForDespatch = pPP.ReasonForDespatch 'If null will get set to control value
          .SubscriptionNumber = mvEnv.GetControlNumber("S")
          .CommunicationNumber = pPPD.CommunicationNumber
          .Save("automatic")
        End If
      End With
      vRecordSet.CloseRecordSet()
    End Sub
    Public Sub ProcessTransactions(ByVal pConn As CDBConnection, ByRef pJob As JobSchedule)
      Dim vBTRS As CDBRecordSet 'Batch Transaction RecordSet
      Dim vBTARS As CDBRecordSet 'Batch Transaction Analysis RecordSet
      Dim vBTARSP As CDBRecordSet 'Batch Transaction Analysis RecordSet Products
      Dim vFHRS As CDBRecordSet 'Financial History RecordSet
      Dim vFHDRS As CDBRecordSet 'Financial History Details RecordSet
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vProduct As New Product(mvEnv)
      Dim vAmount As Double
      Dim vCSWithStockItems As Boolean
      Dim vCSContactNumber As Integer
      Dim vCSAddressNumber As Integer
      Dim vSalesLedgerAccount As String = ""
      Dim vCheckTransaction As Boolean
      Dim vCheckAnalysis As Boolean
      Dim vAlreadyProcessed As Boolean
      Dim vFHDTransactionNo As Integer
      Dim vFHDLineNo As Integer
      Dim vFHTransactionNo As Integer
      Dim vWhereFields As CDBFields
      Dim vGAYEPledge As New PreTaxPledge(mvEnv)
      Dim vTSS As TransactionStockStatus
      Dim vWriteInvoice As Boolean

      mvConn = pConn
      mvJob = pJob
      If mvCompanyControl Is Nothing Then mvCompanyControl = New CompanyControl
      mvCompanyControl.InitFromBankAccount(mvEnv, BankAccount)

      mvBranchIncomePeriod = mvEnv.GetConfig("me_branch_income_period")
      If Len(mvBranchIncomePeriod) = 0 Then mvBranchIncomePeriod = "FIRST"

      mvMeFutureChangeTrigger = mvEnv.GetConfig("me_future_change_trigger")

      mvGAOperationalChangeDate = mvEnv.GetConfig("ga_operational_change_date")
      If Len(mvGAOperationalChangeDate) = 0 Then mvGAOperationalChangeDate = CStr(DateSerial(2000, 4, 6))
      mvGAMembershipTaxReclaim = mvEnv.GetConfigOption("ga_membership_tax_reclaim")
      mvCAFPaymentMethod = mvEnv.GetConfig("pm_caf")
      mvGraceDays = IntegerValue(mvEnv.GetConfig("cv_no_days_claim_grace"))

      If mvEnv.GetConfig("fp_fhd_source_to_use") = "PPD" Then
        mvFHDSourceOrigin = FHDSourceOrigin.fhdsoPaymentPlanDetails
      Else
        mvFHDSourceOrigin = FHDSourceOrigin.fhdsoBatchTransactionAnalysis
      End If

      If mvEnv.GetConfig("fp_fhd_dist_code") = "BTA" Then mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.BatchTransactionAnalysis

      vBT.Init()
      vBTA.Init()

      'Read all transaction analysis lines
      vBTARS = pConn.GetRecordSet("SELECT " & vBTA.GetRecordSetFields() & " FROM batch_transaction_analysis bta WHERE batch_number = " & BatchNumber & " ORDER BY transaction_number, line_number")
      vBTARS.Fetch()

      'Read all the product information for each analysis line that has a product
      vProduct.Init()
      vBTARSP = pConn.GetRecordSet("SELECT bt.transaction_number,bta.line_number," & vProduct.GetRecordSetFields(Product.ProductRecordSetTypes.prstMain) & " FROM batch_transactions bt, batch_transaction_analysis bta, products p WHERE bt.batch_number = " & BatchNumber & " AND bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number AND bta.product = p.product ORDER BY bt.transaction_number, bta.line_number")
      vBTARSP.Fetch()

      'Read all batch transactions
      If BatchType = BatchTypes.CreditSales Then
        vBTRS = pConn.GetRecordSet("SELECT sales_ledger_account,stock_sale,cs.contact_number AS cs_contact_number, cs.address_number AS cs_address_number," & vBT.GetRecordSetFieldsTransactionType & " FROM batch_transactions bt, credit_sales cs, transaction_types tt, contacts c WHERE bt.batch_number = " & BatchNumber & " AND bt.batch_number = cs.batch_number AND bt.transaction_number = cs.transaction_number AND bt.transaction_type = tt.transaction_type AND cs.contact_number = c.contact_number ORDER BY bt.transaction_number")
      Else
        vBTRS = pConn.GetRecordSet("SELECT " & vBT.GetRecordSetFieldsTransactionType & " FROM batch_transactions bt, transaction_types tt, contacts c WHERE bt.batch_number = " & BatchNumber & " AND bt.transaction_type = tt.transaction_type AND bt.contact_number = c.contact_number ORDER BY bt.transaction_number")
      End If

      'Read any existing financial history to see if we have already partially processed the batch
      vFHRS = pConn.GetRecordSet("SELECT transaction_number FROM financial_history WHERE batch_number = " & BatchNumber & " ORDER BY transaction_number")
      If vFHRS.Fetch() = True Then
        vFHTransactionNo = vFHRS.Fields(1).IntegerValue
        vCheckTransaction = True
      End If

      'Read any existing financial history details to see if we have already partially processed the batch
      vFHDRS = pConn.GetRecordSet("SELECT transaction_number, line_number FROM financial_history_details WHERE batch_number = " & BatchNumber & " ORDER BY transaction_number, line_number")
      vCheckAnalysis = True

      While vBTRS.Fetch() = True
        vBT.InitFromRecordSetTransactionType(vBTRS)
        vBT.InitAnalysisFromRecordSets(mvEnv, vBTARS, vBTARSP)
        If BatchType = BatchTypes.CreditSales Then
          vCSWithStockItems = vBTRS.Fields("stock_sale").Bool
          vCSContactNumber = vBTRS.Fields("cs_contact_number").IntegerValue
          vCSAddressNumber = vBTRS.Fields("cs_address_number").IntegerValue
          vSalesLedgerAccount = vBTRS.Fields("sales_ledger_account").Value
        Else
          vCSWithStockItems = False
        End If
        mvCreateFinancialHistory = False
        mvCreateBackOrder = False

        mvProductList = SetUpProductList(vBT)
        mvInvoicesToDelete = ""

        'Here we are going to check for existing financial history details records
        'If we have already processed the batch these will exist and we can ignore
        'each analysis line for which a history details record already exists
        If vCheckAnalysis Then
          Do
            If vFHDTransactionNo = vBT.TransactionNumber Then
              For Each vBTA In vBT.Analysis
                If vFHDLineNo = vBTA.LineNumber Then
                  vBTA.Processed = True 'Mark analysis
                  Exit For
                End If
              Next vBTA
            End If
            'If an fhd line is missing then the xaction no in the FHD could be greater
            'In this case we will not want to move on to the next FHD

            If vFHDTransactionNo <= vBT.TransactionNumber Then 'Added SDT 11/10/2001
              If vFHDRS.Fetch() = True Then
                vFHDTransactionNo = vFHDRS.Fields(1).IntegerValue
                vFHDLineNo = vFHDRS.Fields(2).IntegerValue
              Else
                vCheckAnalysis = False
              End If
            End If
          Loop While vFHDRS.Status() = True And vFHDTransactionNo = vBT.TransactionNumber
        End If

        'Now process the analysis lines which have not already been processed
        vWriteInvoice = False 'Default to not creating an invoice
        Dim vAllocationsInvoice As Invoice = Nothing
        vAmount = ProcessTransactionAnalysis(vBT, vCSWithStockItems, vBT.TransactionSign = "D", vWriteInvoice, vAllocationsInvoice)

        'Here we are going to check for existing financial history records
        'If one exists for this transaction then we can ignore the transaction itself
        vAlreadyProcessed = False
        If vCheckTransaction Then
          If vFHTransactionNo = vBT.TransactionNumber Then
            vAlreadyProcessed = True
          End If
          'If the batch transaction has no analysis lines then the xaction no in the FH could be greater
          'In this case we will not want to move on to the next FH transaction
          If vFHTransactionNo <= vBT.TransactionNumber Then 'Added SDT 18/7/2001
            If vFHRS.Fetch() = True Then
              vFHTransactionNo = vFHRS.Fields(1).IntegerValue 'Was vFHDRS incorrectly SDT 18/7/2001
            Else
              vCheckTransaction = False
            End If
          End If
        End If

        If Not vAlreadyProcessed Then
          vTSS = GetStockStatus(vBT)
          mvConn.StartTransaction()
          Select Case BatchType
            Case BatchTypes.CreditSales
              vBT.Amount = vAmount
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
                vBT.CurrencyAmount = FixTwoPlaces(vBT.Amount * CurrencyExchangeRate)
              End If
              If vBT.TransactionSign = "D" Then
                vBT.Amount = -vAmount
                vBT.CurrencyAmount = -vBT.CurrencyAmount
                If vWriteInvoice Then WriteInvoice(BatchNumber, vBT, vCSContactNumber, vCSAddressNumber, (mvCompanyControl.Company), vSalesLedgerAccount)
              Else
                If vCSWithStockItems = False Then
                  'BR16409: Where the allocation invoice is a sundry credit note and the transaction sign is "C" (normally "D") this transaction represents a sundry credit note reversal- set invoice date
                  Dim vSundryCreditNoteReversal As Boolean = vAllocationsInvoice IsNot Nothing AndAlso vAllocationsInvoice.IsSundryCreditNote
                  Dim vInvoice As Invoice = WriteInvoice(BatchNumber, vBT, vCSContactNumber, vCSAddressNumber, (mvCompanyControl.Company), vSalesLedgerAccount, vSundryCreditNoteReversal)
                  If vInvoice IsNot Nothing AndAlso vSundryCreditNoteReversal Then
                    'BR16409: Where we are posting the reversal of a Sundry Credit Note which creates an Invoice but no invoice details create invoice details and create an IPH record 
                    'to show the Sundry Credit Note paying off the Invoice and set the amount 
                    'Create invoice details
                    For Each vBTA In vBT.Analysis
                      WriteInvoiceDetails(vBTA)
                    Next
                    'Set Invoice Amount Paid to be BT Amount
                    vInvoice.SetAmountPaid(vBT.CurrencyAmount, True, True)
                    'Set Invoice Number to a provisional invoice number (as invoice paid but unprinted) and update invoice details
                    vInvoice.SetInvoiceNumber(True, True)
                    'Save Invoice
                    vInvoice.Save()
                    'Create invoice payment history
                    For Each vBTA In vBT.Analysis
                      Dim vIPHParams As New CDBParameters()
                      With vIPHParams
                        .Add("InvoiceNumber", vInvoice.InvoiceNumber)
                        .Add("BatchNumber", vAllocationsInvoice.BatchNumber)
                        .Add("TransactionNumber", vAllocationsInvoice.TransactionNumber)
                        .Add("LineNumber", vBTA.LineNumber)
                        .Add("Amount", CDBField.FieldTypes.cftNumeric, vBTA.Amount.ToString)
                        .Add("AllocationDate", CDBField.FieldTypes.cftDate, vBT.TransactionDate)
                        .Add("AllocationBatchNumber", vAllocationsInvoice.BatchNumber)
                        .Add("AllocationTransactionNumber", vAllocationsInvoice.TransactionNumber)
                        .Add("AllocationLineNumber", vBTA.LineNumber)
                        If vInvoice.ProvisionalInvoiceNumber > 0 Then .Add("ProvisionalInvoiceNumber", vInvoice.ProvisionalInvoiceNumber)
                      End With
                      Dim vIPH As New InvoicePaymentHistory(mvEnv)
                      vIPH.Create(vIPHParams)
                      vIPH.Save()
                    Next
                  End If
                Else
                  If (vTSS = TransactionStockStatus.tssStockNothingDespatched) Or ((Picked = "N") And mvCreateBackOrder) Then
                    ProcessProductsToBackOrder(pConn, vBT)
                  End If
                End If
              End If

            Case BatchTypes.CreditCard, BatchTypes.DebitCard, BatchTypes.CreditCardWithInvoice
              'If not all stock is issued then the BT.Amount will need updating to reflect this
              vBT.Amount = vAmount
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
                vBT.CurrencyAmount = FixTwoPlaces(vBT.Amount * CurrencyExchangeRate)
              End If
              If vBT.TransactionSign = "D" Then
                vBT.Amount = -vAmount
                vBT.CurrencyAmount = -vBT.CurrencyAmount
              End If
              If vTSS = TransactionStockStatus.tssStockNothingDespatched Then ProcessProductsToBackOrder(pConn, vBT)

            Case BatchTypes.Cash, BatchTypes.CashWithInvoice
              If vTSS = TransactionStockStatus.tssStockNothingDespatched Then ProcessProductsToBackOrder(pConn, vBT)
          End Select

          If mvCreateFinancialHistory Then WriteFinancialHistory(vBT, vCSContactNumber, vCSAddressNumber)
          If mvCreateBackOrder Then WriteBackOrders(vBT)

          If vBT.Amount > 0 Then
            If vBT.Receipt = "V" Or vBT.Receipt = "Y" Then WriteReceipt(vBT, vBT.Receipt, (mvCompanyControl.Company))
            'Receipt type M means there is a contact mailing document so no thank you letter
            If Len(vBT.Mailing) > 0 And vBT.Receipt <> "M" Then WriteThankYouLetter(vBT, (mvCompanyControl.Company))
          End If

          If BatchType = BatchTypes.Cash OrElse BatchType = BatchTypes.CashWithInvoice Then
            If Picked = "N" And vTSS <> TransactionStockStatus.tssNoStock And vBT.Receipt = "N" Then WriteReceipt(vBT, "V", (mvCompanyControl.Company))
            If Picked = "C" And vTSS = TransactionStockStatus.tssStockNothingDespatched And vBT.Receipt = "N" Then WriteReceipt(vBT, "V", (mvCompanyControl.Company))
          Else
            If vTSS = TransactionStockStatus.tssStockNothingDespatched And vBT.Receipt = "N" Then WriteReceipt(vBT, "O", (mvCompanyControl.Company))
          End If

          If mvInvoicesToDelete.Length > 0 Then
            vWhereFields = New CDBFields
            vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftLong, Mid(mvInvoicesToDelete, 2, Len(mvInvoicesToDelete) - 2), CDBField.FieldWhereOperators.fwoIn)
            mvConn.DeleteRecords("invoices", vWhereFields)
          End If

          If BatchType = BatchTypes.FinancialAdjustment And vBT.TransactionSign = "D" Then
            'Reversing a confirmed_transactions record needs to set the confirmed batch / trans numbers to null, but not if the transaction has be moved (change of payer)
            Dim vAnsiJoins As New AnsiJoins()
            vAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")

            vWhereFields = New CDBFields(New CDBField("bt.batch_number", vBT.BatchNumber))
            vWhereFields.Add("transaction_sign", "C")

            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "batch_number, transaction_number, contact_number, amount", "batch_transactions bt", vWhereFields, "", vAnsiJoins)
            Dim vNestedSQL As String = vSQLStatement.SQL

            With vAnsiJoins
              .Clear()
              .Add("reversals r", "bt.batch_number", "r.batch_number", "bt.transaction_number", "r.transaction_number")
              .Add("confirmed_transactions ct", "r.was_batch_number", "ct.confirmed_batch_number", "r.was_transaction_number", "ct.confirmed_trans_number")
              .Add("batches b", "ct.confirmed_batch_number", "b.batch_number")
              .AddLeftOuterJoin("(" & vNestedSQL & ") bt2", "bt.batch_number", "bt2.batch_number", "(bt.transaction_number + 1)", "bt2.transaction_number")
            End With

            With vWhereFields
              .Clear()
              .Add("bt.batch_number", vBT.BatchNumber)
              .Add("bt.transaction_number", vBT.TransactionNumber)
              .Add("bt2.contact_number", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
              .Add("bt2.contact_number#2", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoOR)
              .Add("bt2.contact_number#3", CDBField.FieldTypes.cftInteger, "bt.contact_number", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
              .Add("bt2.amount", CDBField.FieldTypes.cftNumeric, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
              .Add("bt2.amount#2", CDBField.FieldTypes.cftNumeric, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
              .Add("bt2.amount#3", CDBField.FieldTypes.cftNumeric, "bt.amount", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
            End With

            Dim vCT As New ConfirmedTransaction(mvEnv)
            Dim vAttrs As String = vCT.GetRecordSetFields() & ", batch_type, bt2.batch_number AS move_batch_number, bt2.transaction_number AS move_trans_number"
            vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "batch_transactions bt", vWhereFields, "", vAnsiJoins)
            Dim vRevRS As CDBRecordSet = vSQLStatement.GetRecordSet()
            If vRevRS.Fetch Then
              With vRevRS
                If .Fields("move_batch_number").IntegerValue > 0 Then
                  'This may be a Move rather than a Reverse
                  Dim vBatchType As BatchTypes = Batch.GetBatchType(.Fields("batch_type").Value)
                  If vBatchType = BatchTypes.CAFCards OrElse vBatchType = BatchTypes.CAFVouchers OrElse vBatchType = BatchTypes.CAFCommitmentReconciliation Then
                    'These cannot be moved so must be a reversal
                    vCT.InitFromRecordSet(vRevRS)
                  End If
                Else
                  'Definetely a Reverse
                  vCT.InitFromRecordSet(vRevRS)
                End If
                If vCT.Existing Then vCT.ClearConfirmationForReversal()
                .CloseRecordSet()
              End With
            End If
          End If
          'set the Gaye Pledges Last Payment date if this is a GP batch
          'moved to ProcessTransactionAnalysis
          mvConn.CommitTransaction()
        End If
      End While
      If vBTRS.Status And vBTARS.Status() = True Then
        'Check for BTA records with higher TransactionNumber then last BT record
        'I.e. There are BTA records with no corresponding BT record
        If Not vBT Is Nothing Then
          If vBT.TransactionNumber > 0 And (vBT.TransactionNumber < vBTARS.Fields("transaction_number").IntegerValue) Then
            RaiseError(DataAccessErrors.daeBTAndBTADoNotMatch, CStr(BatchNumber))
          End If
        End If
      End If
      vBTRS.CloseRecordSet()
      vBTARS.CloseRecordSet()
      vBTARSP.CloseRecordSet()
      vFHDRS.CloseRecordSet()
      vFHRS.CloseRecordSet()
      mvConn = Nothing
    End Sub
    Private Function ProcessTransactionAnalysis(ByRef pBT As BatchTransaction, ByRef pCSWithStockItems As Boolean, ByRef pTransactionSignD As Boolean, ByRef pWriteInvoice As Boolean, ByRef pAllocationsInvoice As Invoice) As Double
      Dim vBTA As BatchTransactionAnalysis
      Dim vAmount As Double
      Dim vInvoiceDetails As Boolean = False
      Dim vGAYEPledge As New PreTaxPledge(mvEnv)
      Dim vPostTaxPGPledge As PostTaxPledge

      mvBatchType = BatchType
      pBT.SetAdditionalData(BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment)
      For Each vBTA In pBT.Analysis
        If vBTA.Processed Then
          'Already processed
          'If .LineType <> "L" Then mvCreateFinancialHistory = True
          mvCreateFinancialHistory = True
        ElseIf vBTA.LineType = "I" Then
          'I type incentives we may need to add activites for these
          If Len(vBTA.Product.Activity) > 0 And mvEnv.GetConfigOption("fp_payment_incentive_activity") Then
            ProcessCategories((pBT.ContactNumber), (pBT.ContactType), vBTA.Product.Activity, vBTA.Product.ActivityValue, vBTA.Source, (pBT.TransactionDate), (pBT.TransactionDate), mvEnv.GetConfigOption("opt_fp_product_cat_extension"), False, 0, vBTA.Product.ActivityDurationMonths)
          End If
        Else
          'If .LineType <> "L" Then mvCreateFinancialHistory = True
          mvCreateFinancialHistory = True
          mvDeedOrder = False
          If pBT.TransactionSign = "D" Then
            vBTA.ChangeSign()
          End If
          Select Case vBTA.LineType
            Case "P", "G", "S", "H", "B", "X", "D", "F", "Q"  'P Product, G Gift in-memorium, S Soft Credit, H Hard Credit, B Legacy Bequest, X Event Pricing Matrix line, D InMemoriamHardCredit, F InMemoriamSoftCredit, Q Exam booking line
              'Check if this is a product or payplan payment and handle accordingly for S and H line types RH 08/01/2002
              If vBTA.ProductCode.Length > 0 Then
                vAmount = vAmount + ProcessProduct(pBT, vBTA, vBTA.Product, vBTA.RateCode, vBTA.DistributionCode, vBTA.Quantity, vBTA.Amount, vBTA.Issued, vBTA.VatRate, vBTA.VatAmount)
              Else
                If vBTA.Amount <> 0 Then
                  ProcessPaymentPlan(pBT, vBTA)
                  vAmount = vAmount + vBTA.Amount
                End If
              End If
            Case Else
              If vBTA.Amount <> 0 Then
                'N Invoice Payment, U Unallocated Cash, L Allocation of Cash - From unallocated cash, K Sundry Credit Note Allocation to Invoice
                If vBTA.LineType = "N" OrElse vBTA.LineType = "U" OrElse vBTA.LineType = "L" OrElse vBTA.LineType = "K" Then
                  If mvEnv.GetConfigOption("fp_use_sales_ledger") Then
                    If (BatchType = BatchTypes.FinancialAdjustment OrElse (BatchType = BatchTypes.CreditCard OrElse BatchType = BatchTypes.DebitCard OrElse BatchType = BatchTypes.CreditCardWithInvoice)) _
                      AndAlso (pTransactionSignD = True OrElse (pTransactionSignD = False AndAlso vBTA.Amount < 0)) Then
                      If Not mvConn.InTransaction Then
                        mvConn.StartTransaction()
                      End If
                      RemoveInvoiceAllocations(pBT, vBTA, pAllocationsInvoice)
                    End If
                    ProcessInvoice(pBT, vBTA, vBTA.InvoiceNumber)
                  End If
                Else
                  ProcessPaymentPlan(pBT, vBTA)
                End If
                vAmount = vAmount + vBTA.Amount
              End If
          End Select
          If BatchType = BatchTypes.CreditSales Then
            vInvoiceDetails = False
            If pTransactionSignD Then
              'It is potentially a reversal of a stock back order
              Dim vWhereFields As New CDBFields(New CDBField("r.batch_number", vBTA.BatchNumber))
              vWhereFields.Add("r.transaction_number", vBTA.TransactionNumber)
              vWhereFields.Add("r.line_number", vBTA.LineNumber)
              With vWhereFields
                .Add("bod.status", CDBField.FieldTypes.cftCharacter, "R", CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoEqual)
                .Add("bod.ordered", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
                .Add("bod.ordered#2", CDBField.FieldTypes.cftNumeric, "bod.issued", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
                .Add("bod.ordered#3", CDBField.FieldTypes.cftNumeric, "bod.issued", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
              End With
              Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("back_order_details bod", "r.was_batch_number", "bod.batch_number", "r.was_transaction_number", "bod.transaction_number", "r.was_line_number", "bod.line_number", AnsiJoin.AnsiJoinTypes.InnerJoin)})
              Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bod.status, bod.ordered, bod.issued", "reversals r", vWhereFields, "", vAnsiJoins)
              Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
              Dim vIsBackOrder As Boolean = False
              If vRS.Fetch Then
                vIsBackOrder = True
                'Reversing an unfulfilled back order so no invoice details if no stock issued
                vInvoiceDetails = (vRS.Fields("ordered").IntegerValue > 0 AndAlso vRS.Fields("issued").IntegerValue > 0)
              End If
              vRS.CloseRecordSet()
              If vIsBackOrder Then
                ReverseUnfulfilledCSBackOrder(vBTA) 'Update the CreditCustomers
              Else
                vInvoiceDetails = True 'No back order found so create invoice details for credit note
              End If
            Else
              'There was no stock in the transaction so create the invoice details
              If pCSWithStockItems = False Then
                vInvoiceDetails = True
              End If
            End If
            If vInvoiceDetails And Not RemoveInvoiceAllocations(pBT, vBTA, pAllocationsInvoice) Then
              'Write invoice details where remove invoice allocations returns false or where we are 
              WriteInvoiceDetails(vBTA)
              pWriteInvoice = True
            End If
          End If
          'set the Pre Tax Payroll Giving Pledges Last Payment date if this is a GP batch
          If BatchType = BatchTypes.GiveAsYouEarn And Val(vBTA.MemberNumber) <> 0 Then
            vGAYEPledge.Init(IntegerValue(vBTA.MemberNumber))
            If vGAYEPledge.Existing Then
              vGAYEPledge.LastPaymentDate = pBT.TransactionDate
              vGAYEPledge.Save()
            End If
          End If

          'Check for Post Tax Payroll Giving batches
          If BatchType = BatchTypes.PostTaxPayrollGiving And vBTA.LineNumber = 1 And Val(vBTA.MemberNumber) > 0 Then
            vPostTaxPGPledge = New PostTaxPledge(mvEnv)
            vPostTaxPGPledge.Init(IntegerValue(vBTA.MemberNumber))
            If vPostTaxPGPledge.Existing Then
              vPostTaxPGPledge.AddPayment(vBTA.BatchNumber, vBTA.TransactionNumber, pBT.TransactionDate, vBTA.Amount, FixTwoPlaces(If(pBT.TransactionSign = "D", (pBT.Amount * -1), pBT.Amount) - vBTA.Amount))
            End If
          End If

          'We previously changed the amounts so change them back again before we save
          If pBT.TransactionSign.ToUpper.Equals("D") Then
            vBTA.ChangeSign()
          End If

          'Now we have finished with this BTA record, may need to update it
          If BatchType <> BatchTypes.CreditSales AndAlso vBTA.Amount <> 0 AndAlso (vBTA.LineType = "N" OrElse vBTA.LineType = "U" OrElse vBTA.LineType = "L" OrElse vBTA.LineType = "K") AndAlso vBTA.CashInvoiceNumber > 0 Then
            'N Invoice Payment, U Unallocated Cash, L Allocation of Cash - From unallocated cash, K Sundry Credit Note Allocation to Invoice
            vBTA.Save(mvEnv.User.UserID, True)
          End If
        End If
        vBTA.Save()
        mvConn.CommitTransaction()
      Next vBTA
      ProcessTransactionAnalysis = vAmount
    End Function
    Private Sub SetMeFutureChangeFlag(ByRef pPP As PaymentPlan, ByVal pAmount As Double)
      Dim vRecordSet As CDBRecordSet
      Dim vFutureDate As String = ""
      Dim vMember As Member
      Dim vMembershipNumber As Integer
      'This script is be called from start of ProcessPaymentPlan and sets mvMeFutureChange to:
      ' 'Y': this payment is for a future membership change
      ' 'N': this payment is not for a future membership change
      ' 'R': this is a future membership change payment reversal

      If mvMeFutureChangeTrigger = "" Or pAmount = 0 Then
        mvMeFutureChange = "N"
      Else
        If pAmount > 0 Then
          vFutureDate = mvOrgRenewalDate
        End If
        If pAmount < 0 Then 'see if future membership payment being reversed
          vFutureDate = CDate(mvOrgRenewalDate).AddYears(-1).ToString(CAREDateFormat)
        End If
        vRecordSet = mvConn.GetRecordSet("SELECT m.membership_number, mft.future_membership_type, future_change_date FROM members m, member_future_type mft, membership_types mt WHERE m.order_number = " & pPP.PlanNumber & " AND m.membership_number = mft.membership_number AND mft.future_change_date" & mvConn.SQLLiteral("=", CDBField.FieldTypes.cftDate, vFutureDate) & " AND mt.membership_type = mft.future_membership_type")
        If vRecordSet.Fetch() = True Then
          If pAmount < 0 Then
            mvMeFutureChange = "R"
          Else
            mvMeFutureChange = "Y"
          End If
          vMembershipNumber = vRecordSet.Fields(1).IntegerValue
          For Each vMember In pPP.CurrentMembers
            If vMembershipNumber = vMember.MembershipNumber Then
              vMember.FutureMembershipTypeCode = vRecordSet.Fields(2).Value
              vMember.FutureChangeDate = vRecordSet.Fields(3).Value
              Exit For
            End If
          Next vMember
        Else
          mvMeFutureChange = "N"
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    Private Function SetUpProductList(ByRef pBT As BatchTransaction) As String
      Dim vProductList As String = ""
      Dim vBTA As BatchTransactionAnalysis
      Dim vParams As New CDBParameters
      Dim vParam As CDBParameter

      For Each vBTA In pBT.Analysis
        If vBTA.LineType = "I" Then
          If vParams.Exists(vBTA.Product.ProductDesc) Then
            vParams(vBTA.Product.ProductDesc).Value = vParams(vBTA.Product.ProductDesc).Value & vBTA.Quantity.ToString
          Else
            vParams.Add(vBTA.Product.ProductDesc, vBTA.Quantity)
          End If
        End If
      Next vBTA
      For Each vParam In vParams
        If vProductList.Length > 0 Then vProductList = vProductList & vbLf
        vProductList = vProductList & vParam.Name & " " & vParam.IntegerValue
      Next vParam
      SetUpProductList = vProductList
    End Function
    Private Sub TerminateSubscriptions(ByRef pPPNumber As Integer, ByRef pContactNumber As Integer, ByRef pProduct As String, ByRef pExpiryDate As Date)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPPNumber)
      vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pProduct)
      If pContactNumber > 0 Then
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      Else
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong)
      End If
      vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, pExpiryDate.ToString(CAREDateFormat))
      vUpdateFields.AddAmendedOnBy("automatic")
      mvConn.UpdateRecords("subscriptions", vUpdateFields, vWhereFields, False)
    End Sub
    Private Sub UpdateRenewalDate(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByRef pAdd As Boolean)
      pPP.RenewalDate = pPP.CalculateRenewalDate(pPP.RenewalDate, pAdd)
      'update branch income so the branch gets some money
      If mvBranchIncomePeriod = "FIRST" Then WriteBranchIncome(pBT, pBTA, pPP, Nothing, pAdd, 0)

    End Sub
    Private Sub WriteBackOrderDetails(ByRef pBTA As BatchTransactionAnalysis, ByRef pVATRate As String, ByRef pVATAmount As Double, Optional ByRef pZeroIssued As Boolean = False, Optional ByRef pBT As BatchTransaction = Nothing)
      Dim vInsertFields As New CDBFields

      With vInsertFields
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
        .Add("source", CDBField.FieldTypes.cftCharacter, pBTA.Source)
        .Add("despatch_method", CDBField.FieldTypes.cftCharacter, pBTA.DespatchMethod)
        If Val(CStr(pBTA.ContactNumber)) = 0 Then
          If Not pBT Is Nothing Then
            .Add("contact_number", CDBField.FieldTypes.cftCharacter, pBT.ContactNumber)
            .Add("address_number", CDBField.FieldTypes.cftCharacter, pBT.AddressNumber)
          End If
        Else
          .Add("contact_number", CDBField.FieldTypes.cftCharacter, pBTA.ContactNumber)
          .Add("address_number", CDBField.FieldTypes.cftCharacter, pBTA.AddressNumber)
        End If
        .Add("earliest_delivery", CDBField.FieldTypes.cftDate, pBTA.WhenValue)
        .Add("ordered", CDBField.FieldTypes.cftLong, pBTA.Quantity)
        If pZeroIssued Then
          .Add("issued", CDBField.FieldTypes.cftLong, 0) 'Zero issued for Non-Stock and P&P so Confirm Stock Allocation only processes them once
        Else
          .Add("issued", CDBField.FieldTypes.cftLong, pBTA.Issued)
        End If
        .Add("product", CDBField.FieldTypes.cftCharacter, pBTA.ProductCode)
        .Add("rate", CDBField.FieldTypes.cftCharacter, pBTA.RateCode)
        .Add("vat_rate", CDBField.FieldTypes.cftCharacter, pVATRate)
        .Add("unit_price", CDBField.FieldTypes.cftNumeric, pBTA.Amount / pBTA.Quantity)
        .Add("vat_amount", CDBField.FieldTypes.cftNumeric, pVATAmount / pBTA.Quantity)
        If pBTA.GrossAmount <> "" Then .Add("gross_amount", CDBField.FieldTypes.cftNumeric, CDbl(pBTA.GrossAmount) / pBTA.Quantity)
        If pBTA.Discount <> "" Then .Add("discount", CDBField.FieldTypes.cftNumeric, CDbl(pBTA.Discount) / pBTA.Quantity)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
          .Add("currency_unit_price", CDBField.FieldTypes.cftNumeric, pBTA.CurrencyAmount / pBTA.Quantity)
          .Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric, pBTA.CurrencyVatAmount / pBTA.Quantity)
        End If
        .Add("warehouse", CDBField.FieldTypes.cftCharacter, pBTA.Warehouse)
        mvConn.InsertRecord("back_order_details", vInsertFields)
      End With
    End Sub
    Private Sub WriteBackOrders(ByRef pBT As BatchTransaction)
      Dim vInsertFields As New CDBFields

      With vInsertFields
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        .Add("contact_number", CDBField.FieldTypes.cftLong, pBT.ContactNumber)
        .Add("address_number", CDBField.FieldTypes.cftLong, pBT.AddressNumber)
        .Add("transaction_date", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
        .Add("reference", CDBField.FieldTypes.cftCharacter, pBT.Reference)
        If BatchType = BatchTypes.Cash OrElse BatchType = BatchTypes.CashWithInvoice Or BatchType = BatchTypes.FinancialAdjustment Then
          .Add("batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(BatchTypes.BackOrder))
        Else
          .Add("batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(BatchType))
        End If
        .Add("bank_account", CDBField.FieldTypes.cftCharacter, BankAccount)
        .Add("notes", CDBField.FieldTypes.cftMemo, pBT.Notes)
        mvConn.InsertRecord("back_orders", vInsertFields)
      End With
    End Sub
    Sub WriteBranchIncome(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByRef pPPD As PaymentPlanDetail, ByRef pAdd As Boolean, ByVal pAmount As Double)
      Dim vInsertFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vBranchAmount As Double
      'Write or delete a record in the branch_income record
      'ie. the membership is paid this is then used in the branch income report

      vBranchAmount = pAmount
      If pAdd Then
        If mvBranchIncomePeriod = "FIRST" Then
          If pPP.Branch.Length > 0 Then
            If pPP.MembershipType.BranchMembership = True Then
              vInsertFields.Add("branch_code", CDBField.FieldTypes.cftCharacter, pPP.Branch)
              vInsertFields.Add("membership_type", CDBField.FieldTypes.cftCharacter, pPP.MembershipTypeCode)
              vInsertFields.Add("amount", CDBField.FieldTypes.cftNumeric, pBTA.Amount)
              vInsertFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
              vInsertFields.Add("payment_date", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
              mvConn.InsertRecord("branch_income", vInsertFields)
            End If
          End If
        Else
          'code for add for LAST period
          vUpdateFields.Add("amount_outstanding", CDBField.FieldTypes.cftNumeric)
          vUpdateFields.Add("payment_date", CDBField.FieldTypes.cftDate)
          vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
          vWhereFields.Add("renewal_date", CDBField.FieldTypes.cftDate)

          If pBTA.AcceptAsFull = False Then
            'Positive amount to offset against the order_details as normal
            vRecordSet = mvConn.GetRecordSet("SELECT renewal_date,amount_outstanding FROM branch_income WHERE order_number = " & pPP.PlanNumber & " AND amount_outstanding > 0 ORDER BY renewal_date")
            'Select records that the are still owed money against sort by
            'earliest date to latest so that the amount is processed against the earliest record first
            While vRecordSet.Fetch() = True And vBranchAmount > 0
              vBranchAmount = FixTwoPlaces(vBranchAmount - vRecordSet.Fields(2).DoubleValue + pPPD.Balance)
              vWhereFields(2).Value = vRecordSet.Fields(1).Value
              vUpdateFields(1).Value = CStr(pPPD.Balance)
              vUpdateFields(2).Value = pBT.TransactionDate
              mvConn.UpdateRecords("branch_income", vUpdateFields, vWhereFields, True)
            End While
            vRecordSet.CloseRecordSet()
          Else
            'accept as full
            vRecordSet = mvConn.GetRecordSet("SELECT renewal_date,amount_outstanding FROM branch_income WHERE order_number = " & pPP.PlanNumber & " AND amount_due > 0 ORDER BY renewal_date")
            While vRecordSet.Fetch() = True
              vWhereFields(2).Value = vRecordSet.Fields(1).Value
              vUpdateFields(1).Value = CStr(0)
              vUpdateFields(2).Value = pBT.TransactionDate
              mvConn.UpdateRecords("branch_income", vUpdateFields, vWhereFields, True)
            End While
            vRecordSet.CloseRecordSet()
          End If
        End If
      Else
        'negative amount needs to be reversed out against the latest
        'branch income record that has had money processed against it
        vUpdateFields.Add("amount_outstanding", CDBField.FieldTypes.cftNumeric)
        vUpdateFields.Add("payment_date", CDBField.FieldTypes.cftDate)
        vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPP.PlanNumber)
        vWhereFields.Add("renewal_date", CDBField.FieldTypes.cftDate)

        vRecordSet = mvConn.GetRecordSet("SELECT renewal_date,amount_outstanding FROM branch_income WHERE order_number = " & pPP.PlanNumber & " AND amount_due <> amount_outstanding ORDER BY renewal_date DESC")
        While vRecordSet.Fetch() = True And vBranchAmount > 0
          vBranchAmount = FixTwoPlaces(vBranchAmount + vRecordSet.Fields(2).DoubleValue - pPPD.Balance)
          vUpdateFields(1).Value = CStr(pPPD.Balance)
          vUpdateFields(2).Value = pBT.TransactionDate
          vWhereFields(2).Value = vRecordSet.Fields(1).Value
          mvConn.UpdateRecords("branch_income", vUpdateFields, vWhereFields, True)
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    Private Sub WriteFinancialHistory(ByRef pBT As BatchTransaction, ByRef pCSContactNumber As Integer, ByRef pCSAddressNumber As Integer)
      Dim vInsertFields As New CDBFields
      Dim vField As CDBField

      With vInsertFields
        If BatchType = BatchTypes.CreditSales Then
          .Add("contact_number", CDBField.FieldTypes.cftLong, pCSContactNumber)
          .Add("address_number", CDBField.FieldTypes.cftLong, pCSAddressNumber)
        Else
          .Add("contact_number", CDBField.FieldTypes.cftLong, pBT.ContactNumber)
          .Add("address_number", CDBField.FieldTypes.cftLong, pBT.AddressNumber)
        End If
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        .Add("transaction_date", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
        .Add("transaction_type", CDBField.FieldTypes.cftCharacter, pBT.TransactionType)
        If pBT.BankDetailsNumber > 0 Then .Add("bank_details_number", CDBField.FieldTypes.cftLong, pBT.BankDetailsNumber)
        .Add("amount", CDBField.FieldTypes.cftNumeric, pBT.Amount)
        .Add("payment_method", CDBField.FieldTypes.cftCharacter, pBT.PaymentMethod)
        vField = .Add("reference", CDBField.FieldTypes.cftCharacter, pBT.Reference)
        vField.SpecialColumn = True
        .Add("posted", CDBField.FieldTypes.cftDate, TodaysDate)
        .Add("notes", CDBField.FieldTypes.cftMemo, pBT.Notes)
        If BatchType = BatchTypes.BackOrder Then .Add("status", CDBField.FieldTypes.cftCharacter, "B")

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
          .Add("currency_amount", CDBField.FieldTypes.cftNumeric, pBT.CurrencyAmount)
        End If

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataTransactionOrigins) Then
          .Add("transaction_origin", CDBField.FieldTypes.cftCharacter, pBT.TransactionOrigin)
        End If

        mvConn.InsertRecord("financial_history", vInsertFields)
      End With
    End Sub
    Private Sub WriteFinancialHistoryAnalysis(ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pProduct As String, ByRef pRate As String, ByRef pDistributionCode As String, ByRef pQuantity As Integer, ByVal pAmount As Double, ByRef pVATRate As String, ByRef pVATAmount As Double, Optional ByVal pSource As String = "", Optional ByVal pGiverContact As String = "")
      Dim vInsertFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vCurrencyAmount As Double
      Dim vCurrencyVATAmount As Double

      With vInsertFields
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
        If Len(pSource) > 0 And mvFHDSourceOrigin <> FHDSourceOrigin.fhdsoBatchTransactionAnalysis Then
          .Add("source", CDBField.FieldTypes.cftCharacter, pSource) 'Comes from order_details.source
        Else
          .Add("source", CDBField.FieldTypes.cftCharacter, pBTA.Source)
        End If
        If mvEnv.GetConfigOption("fp_use_sales_ledger", True) AndAlso (pBTA.LineType = "N" OrElse pBTA.LineType = "U" OrElse pBTA.LineType = "L" OrElse pBTA.LineType = "K") Then
          .Add("invoice_payment", CDBField.FieldTypes.cftCharacter, "Y")
        Else
          If Not mvDeedOrder Then
            .Add("product", CDBField.FieldTypes.cftCharacter, pProduct)
            .Add("rate", CDBField.FieldTypes.cftCharacter, pRate)
          Else
            .Add("product", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDeedReceivedProduct))
            .Add("rate", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDeedReceivedRate))
          End If
          .Add("distribution_code", CDBField.FieldTypes.cftCharacter, pDistributionCode)
          .Add("quantity", CDBField.FieldTypes.cftLong, pQuantity)
          .Add("invoice_payment", CDBField.FieldTypes.cftCharacter, "N")
          .Add("vat_rate", CDBField.FieldTypes.cftCharacter, pVATRate)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric, pVATAmount)
        End If
        .Add("amount", CDBField.FieldTypes.cftNumeric, pAmount)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
          If CurrencyCode.Length > 0 Then
            'Ensure the currency figures are apportioned the same as the amount
            If System.Math.Abs(pAmount) = System.Math.Abs(pBTA.Amount) And System.Math.Abs(pVATAmount) = System.Math.Abs(pBTA.VatAmount) Then
              If pAmount + pBTA.Amount = 0 Then
                vCurrencyAmount = -(pBTA.CurrencyAmount)
                vCurrencyVATAmount = -(pBTA.CurrencyVatAmount)
              Else
                vCurrencyAmount = pBTA.CurrencyAmount
                vCurrencyVATAmount = pBTA.CurrencyVatAmount
              End If
            Else
              vCurrencyAmount = FixTwoPlaces(pAmount * CurrencyExchangeRate) 'pBTA.CurrencyAmount
              vCurrencyVATAmount = FixTwoPlaces(pVATAmount * CurrencyExchangeRate) 'pBTA.CurrencyVATAmount
            End If
            .Add("currency_amount", CDBField.FieldTypes.cftNumeric, vCurrencyAmount)
            .Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric, vCurrencyVATAmount)
          Else
            .Add("currency_amount", CDBField.FieldTypes.cftNumeric, pAmount)
            .Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric, pVATAmount)
          End If
        End If
        If BatchType = BatchTypes.BackOrder Then .Add("status", CDBField.FieldTypes.cftCharacter, "B")
        If pBTA.SalesContactNumber > 0 Then .Add("sales_contact_number", CDBField.FieldTypes.cftLong, pBTA.SalesContactNumber)
        If pBTA.CashInvoiceNumber > 0 AndAlso (pBTA.LineType = "L" OrElse pBTA.LineType = "N" OrElse pBTA.LineType = "U" OrElse pBTA.LineType = "K") AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbInvoiceAdjustmentStatus) Then
          .Add("cash_invoice_number", pBTA.CashInvoiceNumber)
        End If
      End With
      mvConn.InsertRecord("financial_history_details", vInsertFields)

      If pBTA.LineType = "G" OrElse pBTA.LineType = "S" OrElse pBTA.LineType = "H" Then
        WriteFinancialHistoryLinks(pBT, pBTA, pBTA.LineType, pGiverContact)
      ElseIf pBTA.LineType = "D" Then
        WriteFinancialHistoryLinks(pBT, pBTA, "G")
        WriteFinancialHistoryLinks(pBT, pBTA, "H")
      ElseIf pBTA.LineType = "F" Then
        WriteFinancialHistoryLinks(pBT, pBTA, "G")
        WriteFinancialHistoryLinks(pBT, pBTA, "S")
      End If

    End Sub
    Private Sub WriteFinancialHistoryLinks(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pLineType As String)
      WriteFinancialHistoryLinks(pBT, pBTA, pLineType, "")
    End Sub
    Private Sub WriteFinancialHistoryLinks(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pLineType As String, ByVal pGiverContact As String)
      Dim vDeceasedContact As Integer
      Dim vDonorContact As Integer
      Dim vFLAlreadyExists As Boolean

      If pBTA.LineType = "H" And pGiverContact.Length > 0 Then
        'Create link between the Payer (pBT.ContactNumber) and the Giver (pBTA.DeceasedContactNumber)
        vDonorContact = pBTA.DeceasedContactNumber
        vDeceasedContact = pBT.ContactNumber
      ElseIf (pBTA.LineType = "D" AndAlso pLineType = "H") OrElse (pBTA.LineType = "F" AndAlso pLineType = "S") Then
        vDonorContact = pBT.ContactNumber
        vDeceasedContact = pBTA.ContactNumber
      Else
        vDonorContact = pBT.ContactNumber
        vDeceasedContact = pBTA.DeceasedContactNumber
      End If
      Dim vInsertFields As New CDBFields()
      With vInsertFields
        .Clear()
        .Add("batch_number", pBTA.BatchNumber)
        .Add("transaction_number", pBTA.TransactionNumber)
        .Add("line_number", pBTA.LineNumber)
        .Add("line_type", pLineType)
        .Add("contact_number", vDeceasedContact) 'pBTA.DeceasedContactNumber
        .Add("donor_contact_number", vDonorContact) 'pBT.ContactNumber
      End With
      If Len(pBTA.ProductCode) = 0 And pBTA.Amount <> 0 Then
        'Since this is a payment plan payment ensure that there isn't already an Financial Links record for this BTA
        'This IF is the same logic that's used in ProcessTransactionAnalysis to determine whether a BTA w/ a line type of G, S or H is a payment plan payment
        vFLAlreadyExists = mvConn.GetCount("financial_links", vInsertFields, "") > 0
      End If
      If Not vFLAlreadyExists Then mvConn.InsertRecord("financial_links", vInsertFields)

      If pLineType = "G" Then
        vDonorContact = pBT.ContactNumber
        vDeceasedContact = pBTA.DeceasedContactNumber
        Dim vRelationship As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlInMemoriamRelationship)
        If vRelationship.Length > 0 Then
          Dim vContactLink As New ContactLink(Me.Environment)
          vContactLink.Init(mvEnv, ContactLink.ContactLinkTypes.cltContact, vDeceasedContact, vDonorContact, vRelationship)
          If vContactLink.Existing = False Then
            vContactLink.InitNew(Me.Environment, ContactLink.ContactLinkTypes.cltContact, vDeceasedContact, vDonorContact, vRelationship)
            vContactLink.Save("automatic")
          End If

          vRelationship = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlInMemoriamCompRelationship)
          If vRelationship.Length > 0 Then
            vContactLink = New ContactLink(Me.Environment)
            vContactLink.Init(mvEnv, ContactLink.ContactLinkTypes.cltContact, vDonorContact, vDeceasedContact, vRelationship)
            If vContactLink.Existing = False Then
              vContactLink.InitNew(Me.Environment, ContactLink.ContactLinkTypes.cltContact, vDonorContact, vDeceasedContact, vRelationship)
              vContactLink.Save("automatic")
            End If
          End If
        End If
      End If

    End Sub

    Private Function WriteInvoice(ByVal pBatchNumber As Integer, ByVal pBT As BatchTransaction, ByVal pCSContactNumber As Integer, ByVal pCSAddressNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String) As Invoice
      Return WriteInvoice(pBatchNumber, pBT, pCSContactNumber, pCSAddressNumber, pCompany, pSalesLedgerAccount, False, False)
    End Function

    Private Function WriteInvoice(ByVal pBatchNumber As Integer, ByVal pBT As BatchTransaction, ByVal pCSContactNumber As Integer, ByVal pCSAddressNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pSetInvoiceDate As Boolean) As Invoice
      Return WriteInvoice(pBatchNumber, pBT, pCSContactNumber, pCSAddressNumber, pCompany, pSalesLedgerAccount, pSetInvoiceDate, False)
    End Function

    Private Function WriteInvoice(ByVal pBatchNumber As Integer, ByVal pBT As BatchTransaction, ByVal pCSContactNumber As Integer, ByVal pCSAddressNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pSetInvoiceDate As Boolean, ByVal pUnallocateCreditNote As Boolean) As Invoice
      'pSetInvoiceDate is only set to True when creating new CS batch from batch type CI or AI in normal Batch Posting circumstances it will be False
      Dim vCreateInvoice As Boolean = True
      Dim vUseSalesLedger As Boolean = mvEnv.GetConfigOption("fp_use_sales_ledger", True)
      Dim vWhereFields As New CDBFields(New CDBField("batch_number", pBatchNumber))
      vWhereFields.Add("transaction_number", pBT.TransactionNumber)
      If mvConn Is Nothing Then mvConn = mvEnv.Connection
      If vUseSalesLedger = True AndAlso mvConn.GetCount("invoices", vWhereFields) > 0 Then vCreateInvoice = False
      Dim vNewInvoice As Invoice = Nothing
      If vCreateInvoice Then
        Dim vInvoiceParams As New CDBParameters()
        Dim vInvoiceRecordType As Invoice.InvoiceRecordType = Invoice.InvoiceRecordType.Invoice
        Dim vCCUAmount As Double
        Dim vUpdateCCU As Boolean
        With vInvoiceParams
          .Add("BatchNumber", pBatchNumber)
          .Add("TransactionNumber", pBT.TransactionNumber)
          .Add("ContactNumber", pCSContactNumber)
          .Add("AddressNumber", pCSAddressNumber)
          .Add("Company", pCompany)
          .Add("SalesLedgerAccount", pSalesLedgerAccount)
          .Add("PrintInvoice", "Y")
          .Add("AmountPaid", 0)
          .Add("ReprintCount", CDBField.FieldTypes.cftInteger, "-1")
          .Add("InvoicePayStatus", mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue))
        End With
        Dim vAddInvoiceNumberToCreditNote As Boolean = False
        If pBT.TransactionSign = "C" Then
          vCCUAmount = pBT.Amount
          vUpdateCCU = True
          If vUseSalesLedger Then
            'See if the Transaction contains Payment Plans and if it does then set the InvoiceDate & PaymentDue on the Invoice
            Dim vGotPP As Boolean = False
            For Each vBTA As BatchTransactionAnalysis In pBT.Analysis
              If vBTA.PaymentPlanNumber > 0 Then vGotPP = True
              If vGotPP Then Exit For
            Next
            If vGotPP = True OrElse pSetInvoiceDate = True Then
              Dim vCST As New CreditSalesTerms()
              vCST.Init(mvEnv, pCSContactNumber, pCompany, pSalesLedgerAccount)
              Dim vInvDate As Date = CDate(pBT.TransactionDate)
              Dim vPayDueDate As Date
              Dim vInvoice As New Invoice()
              vInvoice.Init(mvEnv)
              If vInvoice.CalcInvPayDue(vCST.TermsFrom, vCST.TermsPeriod, vCST.TermsNumber, pBT.BatchNumber, pBT.TransactionNumber, vInvDate, vPayDueDate) Then
                vInvoiceParams.Add("InvoiceDate", CDBField.FieldTypes.cftDate, vInvDate.ToString(CAREDateFormat))
                vInvoiceParams.Add("PaymentDue", CDBField.FieldTypes.cftDate, vPayDueDate.ToString(CAREDateFormat))
                If vGotPP = False AndAlso DoubleValue(mvEnv.GetConfig("invoice_date_from_event_start").ToString) > 0 Then
                  'event start date
                  Dim vAnsiJoins As New AnsiJoins
                  vAnsiJoins.Add("events e", "eb.event_number", "e.event_number")
                  Dim vStartDate As String = New SQLStatement(mvEnv.Connection, "e.start_date", "event_bookings eb", vWhereFields, "", vAnsiJoins).GetValue
                  If vStartDate.Length > 0 Then vInvoiceParams("InvoiceDate").Value = vStartDate
                End If
              End If
            End If
          End If
        Else
          vInvoiceRecordType = Invoice.InvoiceRecordType.CreditNote
          If pUnallocateCreditNote Then
            'BR17149- based on config 'cancel_cn_leave_unallocated' do not allocate the credit note being raised against the original invoice
            vUpdateCCU = True
            vCCUAmount = pBT.Amount * -1
          ElseIf vUseSalesLedger Then
            'Need to allocate the credit note being raised against the original invoice
            Dim vAttrs As String = "i.invoice_number, i.amount_paid, bt.amount, r.was_batch_number, r.was_transaction_number, r.line_number, (fhd.amount * -1) AS bta_amount, x.invoice_number AS adjusted_invoice_number, print_invoice"
            Dim vAnsiJoins As New AnsiJoins()
            vAnsiJoins.Add("invoices i", "r.was_batch_number", "i.batch_number", "r.was_transaction_number", "i.transaction_number")
            vAnsiJoins.Add("batch_transactions bt", "i.batch_number", "bt.batch_number", "i.transaction_number", "bt.transaction_number")
            vAnsiJoins.Add("financial_history_details fhd", "r.batch_number", "fhd.batch_number", "r.transaction_number", "fhd.transaction_number", "r.line_number", "fhd.line_number")

            Dim vNestedAttrs As String = "i2.invoice_number, r2.batch_number, r2.transaction_number"
            Dim vNestedAnsiJoins As New AnsiJoins({New AnsiJoin("invoices i2", "r2.was_batch_number", "i2.batch_number", "r2.was_transaction_number", "i2.transaction_number")})
            Dim vNestedSQLStatement As New SQLStatement(mvEnv.Connection, vNestedAttrs, "reversals r2", Nothing, "", vNestedAnsiJoins)
            vAnsiJoins.AddLeftOuterJoin("(" & vNestedSQLStatement.SQL & ") x", "r.was_batch_number", "x.batch_number", "r.was_transaction_number", "x.transaction_number")

            vWhereFields("batch_number").Name = "r.batch_number"
            vWhereFields("transaction_number").Name = "r.transaction_number"
            vWhereFields.Add("fhd.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoNotEqual)
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "reversals r", vWhereFields, "", vAnsiJoins)
            Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
            If vRS.Fetch Then
              Dim vAmount As Double
              Do
                vAmount = vRS.Fields("bta_amount").DoubleValue
                If FixTwoPlaces(vRS.Fields("amount").DoubleValue - vRS.Fields("amount_paid").DoubleValue) = 0 Then
                  'IF the amount of payment is greater than zero
                  'BUT the invoice is fully paid
                  'THEN reduce the outstanding value on the credit_customers record
                  vCCUAmount = vAmount * -1
                  vUpdateCCU = True
                Else
                  'Increase the amount paid on the credit note about to be created
                  vInvoiceParams.Item("AmountPaid").Value = FixTwoPlaces(vInvoiceParams.Item("AmountPaid").DoubleValue + vAmount).ToString

                  Do
                    'Increment the amount_paid attribute for the original invoice
                    Dim vAmountUsed As Double
                    Dim vInvErrMsg As String = ""
                    Dim vOrigInvoice As New Invoice()
                    vOrigInvoice.Init(mvEnv)
                    If vRS.Fields("adjusted_invoice_number").Value.Length > 0 Then
                      'Reversals record may be for an invoice that's been adjusted, e.g. reanalysed.
                      vOrigInvoice.Init(mvEnv, 0, 0, vRS.Fields("adjusted_invoice_number").IntegerValue)
                      vInvErrMsg = "for Invoice Number " & vRS.Fields("adjusted_invoice_number").LongValue
                    ElseIf vRS.Fields("invoice_number").Value.Length > 0 Then
                      'Invoice number may not have been assigned yet...
                      vOrigInvoice.Init(mvEnv, 0, 0, vRS.Fields("invoice_number").IntegerValue)
                      vInvErrMsg = "for Invoice Number " & vRS.Fields("invoice_number").LongValue
                    Else  '...so use the original batch & transation numbers instead
                      vOrigInvoice.Init(mvEnv, vRS.Fields("was_batch_number").IntegerValue, vRS.Fields("was_transaction_number").IntegerValue)
                      vInvErrMsg = "for Batch Number " & vRS.Fields("was_batch_number").IntegerValue & " and Transaction Number " & vRS.Fields("was_transaction_number").IntegerValue
                    End If

                    If Not vOrigInvoice.Existing Then RaiseError(DataAccessErrors.daeInvoiceNotFound, vInvErrMsg)
                    'Set InvoiceNumber (if required)
                    If vOrigInvoice.InvoiceNumber.Length = 0 Then vOrigInvoice.SetInvoiceNumber(True, True)
                    vAddInvoiceNumberToCreditNote = True   'Allocating credit note to invoice so always give it a number

                    If (vAmount <= FixTwoPlaces(vRS.Fields("amount").DoubleValue - vRS.Fields("amount_paid").DoubleValue)) OrElse (vRS.Fields("amount").DoubleValue = 0) Then
                      'IF the amount of the payment is less than or equal to the difference between the total invoice amount and the amount already paid against then invoice...
                      'OR the total invoice amount is zero (which will be the case e.g. when an invoice or event booking paid by invoice is reanalysed and only one of the products is changed to a different product of the same price)
                      'THEN allocate all of vAmount against the invoice
                      vAmountUsed = vAmount
                    Else
                      'Only allocated what's required
                      vAmountUsed = FixTwoPlaces(vRS.Fields("amount").DoubleValue - vRS.Fields("amount_paid").DoubleValue)
                      If vAmountUsed < 0 Then vAmountUsed = vAmount
                    End If
                    vAmount = FixTwoPlaces(vAmount - vAmountUsed)

                    vOrigInvoice.SetAmountPaid(vAmountUsed, True) 'This will not set the amount paid greater than the invoice amount
                    vOrigInvoice.Save()

                    'Reduce the outstanding value on the credit_customers record
                    If vOrigInvoice.InvoiceType = Invoice.InvoiceRecordType.CreditNote AndAlso vRS.Fields("bta_amount").DoubleValue < 0 Then
                      'If the invoice being paid is a credit note and the payment amount was negative then update CCU by the bta amount
                      vAmountUsed = vRS.Fields("bta_amount").DoubleValue
                    End If
                    vCCUAmount += (vAmountUsed * -1)
                    vUpdateCCU = True

                    'Create invoice_payment_history for original invoice
                    Dim vIPHParams As New CDBParameters()
                    With vIPHParams
                      .Add("InvoiceNumber", vOrigInvoice.InvoiceNumber)
                      .Add("BatchNumber", pBatchNumber)
                      .Add("TransactionNumber", pBT.TransactionNumber)
                      .Add("LineNumber", vRS.Fields("line_number").LongValue)
                      .Add("Amount", CDBField.FieldTypes.cftNumeric, vAmountUsed.ToString)
                      .Add("AllocationDate", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
                      .Add("AllocationBatchNumber", pBatchNumber)
                      .Add("AllocationTransactionNumber", pBT.TransactionNumber)
                      .Add("AllocationLineNumber", vRS.Fields("line_number").LongValue)
                      .Add("ProvisionalInvoiceNumber", vOrigInvoice.ProvisionalInvoiceNumber)
                    End With
                    Dim vIPH As New InvoicePaymentHistory(mvEnv)
                    vIPH.Create(vIPHParams)
                    vIPH.Save()
                  Loop While vAmount > 0
                End If

                If vRS.Fields("print_invoice").Value = "N" Then vInvoiceParams.Item("PrintInvoice").Value = "N" 'If Invoice was marked not to print then mark this CreditNote as not to be printed (it could be null!!)
              Loop While vRS.Fetch
            End If
          End If
        End If

        vNewInvoice = New Invoice()
        vNewInvoice.Init(mvEnv)
        vNewInvoice.InvoiceAmount = pBT.Amount
        vNewInvoice.Create(vInvoiceRecordType, vInvoiceParams)
        vNewInvoice.Save(mvEnv.User.UserID)
        If vAddInvoiceNumberToCreditNote Then
          'BR18590 - Add an Invoice number to a Credit Note
          vNewInvoice.SetInvoiceNumber(True, True)
          vNewInvoice.Save(mvEnv.User.UserID)
        End If

        If vUpdateCCU Then
          Dim vCCU As New CreditCustomer
          vCCU.Init(mvEnv, pCSContactNumber, pCompany, pSalesLedgerAccount)
          If vCCU.Existing Then
            If pBT.TransactionSign = "C" Then vCCU.AdjustOnOrder(vCCUAmount) 'Reduce OnOrder
            vCCU.AdjustOutstanding(vCCUAmount)                               'Increase Outstanding
            vCCU.Save()
          End If
        End If
      End If
      Return vNewInvoice
    End Function
    Private Sub WriteInvoiceDetails(ByRef pBTA As BatchTransactionAnalysis)
      Dim vFields As New CDBFields

      With vFields
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
        If mvConn.GetCount("invoice_details", vFields) = 0 Then
          mvConn.InsertRecord("invoice_details", vFields)
        End If
      End With
    End Sub

    ''' <summary>Create an Invoice Payment History record for an Invoice Payment line</summary>
    ''' <param name="pInvoiceNumber">Invoice number of the Invoice being paid</param>
    ''' <param name="pBatchNumber">Batch number of the payment line</param>
    ''' <param name="pTransactionNumber">Transaction number of the payment line</param>
    ''' <param name="pLineNumber">Line number of the payment line</param>
    ''' <param name="pAmount">Amount of invoice payment</param>
    Private Sub WriteInvoicePaymentHistory(ByVal pInvoiceNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pAmount As Double)
      WriteInvoicePaymentHistory(pInvoiceNumber, pBatchNumber, pTransactionNumber, pLineNumber, pAmount, "", False)
    End Sub
    ''' <summary>Create an Invoice Payment History record for an Invoice Payment line</summary>
    ''' <param name="pInvoiceNumber">Invoice number of the Invoice being paid</param>
    ''' <param name="pBatchNumber">Batch number of the payment line</param>
    ''' <param name="pTransactionNumber">Transaction number of the payment line</param>
    ''' <param name="pLineNumber">Line number of the payment line</param>
    ''' <param name="pAmount">Amount of invoice payment</param>
    ''' <param name="pTransactionDate">Date of the invoice payment</param>
    Private Sub WriteInvoicePaymentHistory(ByVal pInvoiceNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pAmount As Double, ByVal pTransactionDate As String, ByVal pIsPartRefund As Boolean)
      'The reason that the BatchTransactionAnalysis object parameter was removed is because the both subs ProcessInvoice
      'and WriteSalesLedgerInvoice know all of the relevant info that this sub requires so it was no longer necessary to
      'pass a BatchTransactionAnalysis object.

      Dim vIPH As New InvoicePaymentHistory(mvEnv)
      'IPH may already exist so select it first (as there is no unique index have to use all fields)
      Dim vWherefields As New CDBFields()
      With vWherefields
        .Add("invoice_number", pInvoiceNumber)
        .Add("batch_number", pBatchNumber)
        .Add("transaction_number", pTransactionNumber)
        .Add("line_number", pLineNumber)
        If pIsPartRefund = False Then .Add("amount", CDBField.FieldTypes.cftNumeric, pAmount) 'Do not include amount for part-refund as it could be different
        If IsDate(pTransactionDate) = False Then pTransactionDate = TodaysDate()
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) Then
          .Add("allocation_date", CDBField.FieldTypes.cftDate, pTransactionDate)
          .Add("allocation_batch_number", pBatchNumber)
          .Add("allocation_transaction_number", pTransactionNumber)
          .Add("allocation_line_number", pLineNumber)
        End If
      End With
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vIPH.GetRecordSetFields, "invoice_payment_history iph", vWherefields)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      Dim vCreateIPH As Boolean = (vRS.Fetch = False)   'If we found a record then don't do anything
      vRS.CloseRecordSet()

      If vCreateIPH Then
        Dim vParams As New CDBParameters()
        With vParams
          .Add("InvoiceNumber", pInvoiceNumber)
          .Add("BatchNumber", pBatchNumber)
          .Add("TransactionNumber", pTransactionNumber)
          .Add("LineNumber", pLineNumber)
          .Add("Amount", pAmount)
          If IsDate(pTransactionDate) = False Then pTransactionDate = TodaysDate()
          .Add("AllocationDate", pTransactionDate)
          .Add("AllocationBatchNumber", pBatchNumber)
          .Add("AllocationTransactionNumber", pTransactionNumber)
          .Add("AllocationLineNumber", pLineNumber)
        End With
        vIPH.Create(vParams)
        vIPH.Save(mvEnv.User.UserID)
      End If

    End Sub
    'Private Sub WriteIssuedStock(ByRef pBTA As BatchTransactionAnalysis)
    '  Dim vIS As New IssuedStock
    '  vIS.Init(mvEnv)
    '  With pBTA
    '    vIS.Create(.BatchNumber, .TransactionNumber, .LineNumber, .ProductCode, .Issued, .Warehouse, mvJob.JobNumber)
    '  End With
    '  vIS.Save()
    'End Sub
    Private Sub WriteOrderPaymentHistory(ByRef pBTA As BatchTransactionAnalysis, ByRef pPP As PaymentPlan, ByVal pAmount As Double, ByRef pBalance As Double, ByVal pPPReCalc As Boolean, ByVal pOrigOPSClaimDate As String)
      Dim vOPH As OrderPaymentHistory = pBTA.OrderPaymentHistory
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vWhereFields As New CDBFields
      Dim vAmount As Double
      Dim vFound As Boolean
      Dim vSum As Double
      Dim vOPSSaved As Boolean
      Dim vWriteOffLineAmount As Double = 0

      If vOPH.Existing Then
        If pBalance <> 0 AndAlso Not pBTA.AcceptAsFull AndAlso vOPH.WriteOffLineAmount <> 0 Then pBalance = 0
        vOPH.SetPosted(True, pBalance)
        vOPH.Save()
        If vOPH.ScheduledPaymentNumber.Length > 0 Then
          For Each vOPS In pPP.ScheduledPayments
            If vOPS.ScheduledPaymentNumber = Val(vOPH.ScheduledPaymentNumber) Then vFound = True
            If vFound Then Exit For
          Next vOPS
          If Not vFound Then
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv, IntegerValue(vOPH.ScheduledPaymentNumber))
          End If
          If (mvBatchType = BatchTypes.DirectDebit Or mvBatchType = BatchTypes.CreditCard Or mvBatchType = BatchTypes.CreditCardWithInvoice Or mvBatchType = BatchTypes.CreditCardAuthority) And pPPReCalc = True And IsDate(pOrigOPSClaimDate) Then
            'May need to "move" the OPH from one OPS to another
            vOPSSaved = ResetOPSForAutoPayment(pPP, vOPH, vOPS, pOrigOPSClaimDate)
          End If
          If vOPSSaved = False Then
            If vOPH.WriteOffLineAmount <> 0 AndAlso pBalance = 0 Then
              Select Case vOPS.ScheduledPaymentStatus
                Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsArrears, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid
                  'Where OPH write off line amount set (either +ve for write off payment or -ve for reversal of write off payment) and BTA accept as full not set (pBalance = 0) then write off OPS line
                  'WriteOff returning the actual amount written off (as the OPS.AmountOutstanding could be less than the Write Off Line Amount due to other Order Payments)
                  Dim vOPHWriteOffLineAmount As Double = vOPH.WriteOffLineAmount
                  vOPH.SetWriteOffLineAmount(vOPS.WriteOff(vOPHWriteOffLineAmount))
                  If vOPH.WriteOffLineAmount <> vOPHWriteOffLineAmount Then
                    'Write Off Line Amount changed so save
                    vOPH.Save()
                  End If
                  'Set BTA.WrittenOffLineAmount such that the PP Balance and Payment Plan Details Balance will be adjusted by the write off amount
                  pBTA.WriteOffLineAmount = vOPH.WriteOffLineAmount
              End Select
            Else
              vOPS.ProcessPayment((pPP.IsCancelled))
            End If
            If CheckConfirmedTransactionReversal(pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber) Then
              'BR15574: If posting an order payment- confirmed transaction reversal then reset the OPS Payment Status to 'Unprocessed Payment'.
              vOPS.SetUnProcessedPayment(True, 0)
            End If
            vOPS.Save()
          End If
        End If
      Else
        pPP.PaymentNumber = pPP.PaymentNumber + 1
        vOPH.Init(mvEnv)
        vOPH.SetValues((pBTA.BatchNumber), (pBTA.TransactionNumber), (pPP.PaymentNumber), (pPP.PlanNumber), pAmount, (pBTA.LineNumber), pBalance, 0, True)
        vOPH.Save()
      End If

      If pBalance <> 0 Then
        If pAmount > 0 Then
          'Amount to be written off
          For Each vOPS In pPP.ScheduledPayments
            Select Case vOPS.ScheduledPaymentStatus
              Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsArrears, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid
                vOPS.WriteOff(vOPS.AmountOutstanding)
                vOPS.Save()
            End Select
          Next vOPS
        Else
          'Amount written off to be reinstated
          'There is no definitive list of how much was written off for each scheduled payment
          'Process the payments from the last to the first
          vOPS = Nothing
          With vWhereFields
            .Clear()
            .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong)
          End With
          Do
            'First deal with all ops with a staus of Written Off
            vOPS = DirectCast(mvEnv.GetPreviousItem(pPP.ScheduledPayments, vOPS), OrderPaymentSchedule)
            If Not (vOPS Is Nothing) Then
              Select Case vOPS.ScheduledPaymentStatus
                Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsWrittenOff
                  'Find value of payments already made
                  vWhereFields("scheduled_payment_number").Value = CStr(vOPS.ScheduledPaymentNumber)
                  vSum = Val(mvEnv.Connection.GetValue("SELECT SUM(amount) FROM order_payment_history WHERE " & mvEnv.Connection.WhereClause(vWhereFields)))
                  vAmount = FixTwoPlaces(vOPS.AmountDue - vSum)
                  If vAmount > 0 Then
                    If pBalance + vAmount > 0 Then vAmount = pBalance * -1
                    vOPS.WriteOff((vAmount * -1))
                    vOPS.Save()
                    pBalance = FixTwoPlaces(pBalance + vAmount)
                  End If
              End Select
            End If
          Loop While pBalance < 0 And Not (vOPS Is Nothing)

          If pBalance < 0 Then
            'Now deal with all other statuses until pBalance = 0
            vOPS = Nothing
            Do
              vOPS = DirectCast(mvEnv.GetPreviousItem(pPP.ScheduledPayments, vOPS), OrderPaymentSchedule)
              If Not (vOPS Is Nothing) Then
                Select Case vOPS.ScheduledPaymentStatus
                  Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsFullyPaid, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
                    'Find value of payments already made less those still expected
                    vWhereFields("scheduled_payment_number").Value = CStr(vOPS.ScheduledPaymentNumber)
                    vSum = Val(mvEnv.Connection.GetValue("SELECT SUM(amount) FROM order_payment_history WHERE " & mvEnv.Connection.WhereClause(vWhereFields)))
                    vAmount = FixTwoPlaces((vOPS.AmountDue - vOPS.AmountOutstanding) - vSum)
                    If vAmount > 0 Then
                      If pBalance + vAmount > 0 Then vAmount = pBalance * -1
                      vOPS.WriteOff((vAmount * -1))
                      vOPS.Save()
                      pBalance = FixTwoPlaces(pBalance + vAmount)
                    End If
                End Select
              End If
            Loop While pBalance < 0 And Not (vOPS Is Nothing)
          End If
        End If
      End If

    End Sub
    Private Sub WritePPBackOrderDetails(ByVal pBT As BatchTransaction, ByVal pRS As CDBRecordSet)
      Dim vInsertFields As New CDBFields
      Dim vQuantity As Integer
      Dim vAmount As Double
      Dim vVatAmount As Double
      Dim vBTA As BatchTransactionAnalysis

      With vInsertFields
        vQuantity = pRS.Fields("quantity").IntegerValue
        vAmount = pRS.Fields("amount").DoubleValue
        vVatAmount = pRS.Fields("vat_amount").DoubleValue

        .Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pRS.Fields("line_number").IntegerValue)
        .Add("source", CDBField.FieldTypes.cftCharacter, pRS.Fields("source").Value)
        .Add("contact_number", CDBField.FieldTypes.cftLong, pBT.ContactNumber)
        .Add("address_number", CDBField.FieldTypes.cftLong, pBT.AddressNumber)
        .Add("ordered", CDBField.FieldTypes.cftLong, vQuantity)
        .Add("issued", CDBField.FieldTypes.cftLong, 0)
        .Add("product", CDBField.FieldTypes.cftCharacter, pRS.Fields("product").Value)
        .Add("rate", CDBField.FieldTypes.cftCharacter, pRS.Fields("rate").Value)
        .Add("vat_rate", CDBField.FieldTypes.cftCharacter, pRS.Fields("vat_rate").Value)
        If vQuantity = 0 Then
          .Add("unit_price", CDBField.FieldTypes.cftNumeric, 0)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric, 0)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
            .Add("currency_unit_price", CDBField.FieldTypes.cftNumeric, 0)
            .Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric, 0)
          End If
        Else
          .Add("unit_price", CDBField.FieldTypes.cftNumeric, vAmount / vQuantity)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric, vVatAmount / vQuantity)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
            .Add("currency_unit_price", CDBField.FieldTypes.cftNumeric, pRS.Fields("currency_amount").DoubleValue / vQuantity)
            .Add("currency_unit_price", CDBField.FieldTypes.cftNumeric, pRS.Fields("currency_vat_amount").DoubleValue / vQuantity)
          End If
        End If
        For Each vBTA In pBT.Analysis
          If vBTA.ProductCode.Length > 0 Then
            If vBTA.Product.PostagePacking Then
              .Add("gross_amount", CDBField.FieldTypes.cftNumeric, vBTA.GrossAmount)
              .Add("discount", CDBField.FieldTypes.cftNumeric, vBTA.Discount)
            End If
          End If
        Next vBTA
        mvConn.InsertRecord("back_order_details", vInsertFields)
      End With
    End Sub
    Private Sub WriteReceipt(ByRef pBT As BatchTransaction, ByRef pReceiptType As String, ByRef pCompany As String)
      Dim vInsertFields As New CDBFields
      Dim vReceiptNumber As Integer

      vReceiptNumber = mvEnv.GetControlNumber("R")
      If mvStartReceiptNumber = 0 Then mvStartReceiptNumber = vReceiptNumber
      With vInsertFields
        .Add("receipt_number", CDBField.FieldTypes.cftLong, vReceiptNumber)
        .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
        .Add("contact_number", CDBField.FieldTypes.cftLong, pBT.ContactNumber)
        .Add("address_number", CDBField.FieldTypes.cftLong, pBT.AddressNumber)
        .Add("vat_receipt", CDBField.FieldTypes.cftCharacter, pReceiptType)
        .Add("company", CDBField.FieldTypes.cftCharacter, pCompany)
        mvConn.InsertRecord("receipts", vInsertFields)
      End With
    End Sub

    Private Sub WriteSalesLedgerInvoice(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pAdjustment As Boolean)
      Dim vInvoice As New Invoice()
      vInvoice.Init(mvEnv)

      'Determine if an existing C-type record already exists for the batch and transaction number
      Dim vWhereFields As New CDBFields()
      With vWhereFields
        .Add("batch_number", pBT.BatchNumber)
        .Add("transaction_number", pBT.TransactionNumber)
        .Add("contact_number", pBTA.ContactNumber)
        .Add("address_number", pBTA.AddressNumber)
        .Add("company", mvCompanyControl.Company)
        .Add("sales_ledger_account", pBTA.MemberNumber) 'Hack??
        .Add("record_type", Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash))
        .Add("invoice_date", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
      End With

      Dim vInvoiceAmount As Double = 0
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll), "invoices i", vWhereFields)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      If vRS.Fetch Then
        vInvoice.InitFromRecordSet(mvEnv, vRS, Invoice.InvoiceRecordSetTypes.irtAll)
        If BatchType = BatchTypes.FinancialAdjustment AndAlso pBT.Amount = 0 Then
          If pAdjustment = False AndAlso pBT.TransactionSign = "C" AndAlso pBTA.Amount > 0 And pBTA.LineNumber > 1 Then
            vInvoiceAmount = FixTwoPlaces(vInvoice.InvoiceAmount + pBTA.Amount)
          Else
            vInvoiceAmount = vInvoice.GetAdjustmentInvoiceAmounts(pBT.BatchNumber, pBT.TransactionNumber, True)
          End If
        Else
          vInvoiceAmount = pBT.Amount
        End If
        vInvoice.InvoiceAmount = vInvoiceAmount

        'Create the invoice details for the found invoice
        Dim vInvoiceDetail As New InvoiceDetail()
        vInvoiceDetail.Create(mvEnv, pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, IntegerValue(vInvoice.InvoiceNumber))
        vInvoiceDetail.Save(mvEnv.User.UserID)

        If pBTA.LineType = "U" Then 'unallocated portion of current cash payment
          If vInvoice.InvoicePayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid) Then 'invoice previously created by line type of N
            'Change the Amountpaid and reset the invoice's pay status if required
            If pBTA.Amount < 0 AndAlso pAdjustment = True Then
              vInvoice.SetAmountPaid(pBTA.Amount, False)
            Else
              vInvoice.SetAmountPaid(0, False)
            End If
          ElseIf (BatchType = BatchTypes.FinancialAdjustment AndAlso pBT.Amount = 0 AndAlso pBTA.Amount < 0) Then
            'Increase the amount paid
            vInvoice.SetAmountPaid(pBTA.Amount, (vInvoice.InvoicePayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)))
          End If
        Else 'assume line type = N - allocated portion of current cash payment
          Select Case vInvoice.InvoicePayStatus
            Case mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue)
              'Reset the invoice's pay status to Partially Paid and increase the amount paid
              vInvoice.SetAmountPaid(pBTA.Amount, False)
            Case mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid), mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid)
              'Increase the amount paid
              vInvoice.SetAmountPaid(pBTA.Amount, (vInvoice.InvoicePayStatus = mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid)))
          End Select
        End If
        vInvoice.Save(mvEnv.User.UserID)
      Else
        'An invoice doesn't exists, so create one
        Dim vParams As New CDBParameters()
        With vParams
          .Add("BatchNumber", pBT.BatchNumber)
          .Add("TransactionNumber", pBT.TransactionNumber)
          .Add("ContactNumber", pBTA.ContactNumber)
          .Add("AddressNumber", pBTA.AddressNumber)
          .Add("Company", mvCompanyControl.Company)
          .Add("SalesLedgerAccount", pBTA.MemberNumber) 'Hack??
          .Add("InvoiceDate", CDBField.FieldTypes.cftDate, pBT.TransactionDate)
          .Add("AmountPaid", 0)
          .Add("ReprintCount", 0)
          vInvoiceAmount = pBTA.Amount
          If BatchType = BatchTypes.FinancialAdjustment AndAlso pBT.Amount = 0 Then
            If Not (pAdjustment = False AndAlso pBT.TransactionSign = "C" AndAlso pBTA.Amount < 0) Then vInvoiceAmount = vInvoice.GetAdjustmentInvoiceAmounts(pBT.BatchNumber, pBT.TransactionNumber, True)
          End If
          If pBTA.LineType = "N" OrElse pAdjustment = True OrElse (pBTA.LineType = "U" AndAlso pBT.IsFinancialAdjustment AndAlso pBTA.Amount < 0) Then
            'When the Line Type is Invoice Payment OR when the TransactionType.TransactionSign = 'D' OR when an Unallocated Cash analysis line exists in a FA transaction 
            'THEN create the invoices records as Fully Paid/Allocated
            .Item("AmountPaid").Value = pBTA.Amount.ToString
          End If
          vInvoice.InvoiceAmount = vInvoiceAmount
          Dim vPayStatus As CDBEnvironment.InvoicePayStatusTypes = CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue
          If .Item("AmountPaid").DoubleValue <> 0 Then
            If .Item("AmountPaid").DoubleValue < vInvoiceAmount Then
              vPayStatus = CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid
            Else
              vPayStatus = CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid
            End If
          End If
          .Add("InvoicePayStatus", mvEnv.GetInvoicePayStatus(vPayStatus))
        End With
        vInvoice.Create(Invoice.InvoiceRecordType.SalesLedgerCash, vParams)
        vInvoice.Save(mvEnv.User.UserID)

        'If the BTA is a U-type line that was created via the use of the Remove Allocations menu item on the Account card...
        If pBTA.LineType = "U" AndAlso BatchType = Batch.BatchTypes.FinancialAdjustment AndAlso pAdjustment = False AndAlso pBT.IsFinancialAdjustment = False Then
          '...then update the credit customer
          Dim vCCU As New CreditCustomer()
          vCCU.InitCompanySalesLedgerAccount(mvEnv, mvCompanyControl.Company, pBTA.MemberNumber)
          If vCCU.Existing Then
            vCCU.AdjustOutstanding((Math.Abs(pBTA.Amount) * -1))
            vCCU.Save(mvEnv.User.UserID)
          End If
        End If

        'Create the invoice details for the created invoice
        Dim vInvoiceDetail As New InvoiceDetail()
        vInvoiceDetail.Create(mvEnv, pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, IntegerValue(vInvoice.InvoiceNumber))
        vInvoiceDetail.Save(mvEnv.User.UserID)
      End If
      vRS.CloseRecordSet()
      If vInvoice.Existing Then pBTA.SetCashInvoiceNumber(IntegerValue(vInvoice.InvoiceNumber))

    End Sub
    Private Sub WriteThankYouLetter(ByRef pBT As BatchTransaction, ByRef pCompany As String)
      Dim vInsertFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vSuppExclusion As String = ""
      Dim vSuppressions() As String
      Dim vIndex As Integer
      Dim vContactNumber As Integer
      Dim vAddressNumber As Integer
      Dim vDone As Boolean

      If pBT.MailingContactNumber > 0 Then
        'SDT Removed the following config check 11/1/2001
        'If mvEnv.GetConfigOption("different_contact_to_mail") And pBT.MailingContactNumber > 0 Then
        vContactNumber = pBT.MailingContactNumber
        vAddressNumber = pBT.MailingAddressNumber
      Else
        vContactNumber = pBT.ContactNumber
        vAddressNumber = pBT.AddressNumber
      End If
      vSuppressions = Split(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlTYLSupressionExclusionList), ",")
      If UBound(vSuppressions) >= 0 Then
        For vIndex = 0 To UBound(vSuppressions)
          If vIndex > 0 Then vSuppExclusion = vSuppExclusion & ","
          vSuppExclusion = vSuppExclusion & "'" & vSuppressions(vIndex) & "'"
        Next
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, vContactNumber)
        vWhereFields.Add("mailing_suppression", CDBField.FieldTypes.cftCharacter, vSuppExclusion, CDBField.FieldWhereOperators.fwoIn)
        If mvConn.GetCount("contact_suppressions", vWhereFields) > 0 Then vDone = True
      End If

      If Not vDone Then
        With vInsertFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
          .Add("contact_number", CDBField.FieldTypes.cftLong, vContactNumber)
          .Add("address_number", CDBField.FieldTypes.cftLong, vAddressNumber)
          .Add("mailing", CDBField.FieldTypes.cftCharacter, pBT.Mailing)
          .Add("company", CDBField.FieldTypes.cftCharacter, pCompany)
          .Add("product_list", CDBField.FieldTypes.cftCharacter, mvProductList)
          mvConn.InsertRecord("thank_you_letters", vInsertFields)
        End With
      End If
    End Sub

    Public Sub SetCurrency(ByRef pCurrencyCode As String, ByRef pExchangeRate As Double)
      mvClassFields.Item(BatchFields.CurrencyCode).Value = pCurrencyCode
      mvClassFields.Item(BatchFields.CurrencyExchangeRate).Value = CStr(pExchangeRate)
    End Sub
    Public Sub ProcessDataImportGiftAid(ByVal pConn As CDBConnection, ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, Optional ByRef pPP As PaymentPlan = Nothing)
      'Read config value if required config
      If Not mvGAInitialized Then
        mvConn = pConn
        If Len(mvGAOperationalChangeDate) = 0 Then mvGAOperationalChangeDate = mvEnv.GetConfig("ga_operational_change_date")
        If Len(mvGAOperationalChangeDate) = 0 Then mvGAOperationalChangeDate = CStr(DateSerial(2000, 4, 6))
        If Not mvGAMembershipTaxReclaim Then mvGAMembershipTaxReclaim = mvEnv.GetConfigOption("ga_membership_tax_reclaim")
        If Len(mvCAFPaymentMethod) = 0 Then mvCAFPaymentMethod = mvEnv.GetConfig("pm_caf")
        If mvGraceDays = 0 Then mvGraceDays = IntegerValue(mvEnv.GetConfig("cv_no_days_claim_grace"))
        mvGAInitialized = True
      End If
      mvImportPayment = True
      If pPP Is Nothing Then
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidSponsorship) = True And pBTA.Product.SponsorshipEvent = True Then
          ProcessGiftAidSponsorship(pBT, pBTA, (pBTA.Amount))
        Else
          If pBTA.Product.Donation = True And pBTA.Product.EligibleForGiftAid = True Then
            ProcessGiftAidDeclarations(pBT, pBTA, "D", (pBTA.Amount))
          End If
        End If
      End If
    End Sub
    Public Sub InitForDataImport(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0)
      If pBatchNumber = 0 Then
        InitNewBatch(pEnv)
      Else
        MyBase.Init()
        mvClassFields.Item(BatchFields.BatchNumber).IntegerValue = pBatchNumber
      End If
    End Sub
    Private Function RemoveInvoiceAllocations(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByRef pInvoice As Invoice) As Boolean
      Return Invoice.RemoveInvoiceAllocations(mvEnv, pBT, pBTA, pInvoice, BatchType, mvInvoicesToDelete, AdjustmentTypes.atNone)
    End Function
    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(BatchFields.AmendedOn).Value = pAmendedOn
      mvClassFields.Item(BatchFields.AmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Function IsOpenBatch() As Boolean
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("b.batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
      vWhereFields.Add("ob.batch_type", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("ob.batch_number", CDBField.FieldTypes.cftInteger, "b.batch_number")
      If mvEnv.Connection.GetCount("open_batches ob, batches b", vWhereFields) = 1 Then IsOpenBatch = True
    End Function

    Public Sub CloseOpenBatch()
      Dim vWhereFields As New CDBFields
      Dim vTransactions As New CDBParameters

      LockBatch()
      mvEnv.Connection.StartTransaction()
      'remove open_batches record
      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchFields.BatchNumber).IntegerValue)
      mvEnv.Connection.DeleteRecords("open_batches", vWhereFields)
      'set the batch totals
      mvClassFields.Item(BatchFields.NumberOfEntries).Value = mvClassFields.Item(BatchFields.NumberOfTransactions).Value
      mvClassFields.Item(BatchFields.BatchTotal).Value = mvClassFields.Item(BatchFields.TransactionTotal).Value
      If BatchType = BatchTypes.FinancialAdjustment Then
        mvClassFields.Item(BatchFields.ReadyForBanking).Bool = True
        mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = True
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        mvClassFields.Item(BatchFields.CurrencyBatchTotal).Value = mvClassFields.Item(BatchFields.CurrencyTransactionTotal).Value
      End If
      If mvClassFields.Item(BatchFields.Provisional).InDatabase And Provisional Then
        mvClassFields.Item(BatchFields.Picked).Value = "C"
        If BatchType <> BatchTypes.CAFCards Then
          mvClassFields.Item(BatchFields.PostedToNominal).Bool = True
          mvClassFields.Item(BatchFields.PayingInSlipPrinted).Bool = True
          mvClassFields.Item(BatchFields.PostedToCashBook).Bool = True
        End If
      End If
      'verify that the batch is complete
      If SetDetailComplete(vTransactions, True, True) Then
        mvEnv.Connection.CommitTransaction()
      Else
        mvEnv.Connection.CommitTransaction()
        RaiseError(DataAccessErrors.daeUnbalancedBatch)
      End If
    End Sub

    Public Function AdjustmentBatchType(ByVal pAdjustmentType As AdjustmentTypes) As BatchTypes
      'If adding new BatchTypes in here please change Delete as well (adjustment batches can not be deleted)

      Select Case BatchType
        Case BatchTypes.DirectDebit, BatchTypes.StandingOrder
          If pAdjustmentType = AdjustmentTypes.atRefund Then
            AdjustmentBatchType = BatchTypes.DirectCredit
          Else
            AdjustmentBatchType = BatchTypes.FinancialAdjustment
          End If
        Case BatchTypes.CreditSales
          AdjustmentBatchType = BatchTypes.CreditSales
        Case BatchTypes.DebitCard
          If pAdjustmentType = AdjustmentTypes.atRefund Or pAdjustmentType = AdjustmentTypes.atPartRefund Then
            AdjustmentBatchType = BatchTypes.DebitCard
          Else
            AdjustmentBatchType = BatchTypes.FinancialAdjustment
          End If
        Case BatchTypes.CreditCard, BatchTypes.CreditCardWithInvoice
          If pAdjustmentType = AdjustmentTypes.atRefund Or pAdjustmentType = AdjustmentTypes.atPartRefund Then
            AdjustmentBatchType = BatchTypes.CreditCard
          Else
            AdjustmentBatchType = BatchTypes.FinancialAdjustment
          End If
        Case BatchTypes.CreditCardAuthority
          If pAdjustmentType = AdjustmentTypes.atRefund Then
            AdjustmentBatchType = BatchTypes.CreditCard
          Else
            AdjustmentBatchType = BatchTypes.FinancialAdjustment
          End If
        Case BatchTypes.GiveAsYouEarn
          AdjustmentBatchType = BatchTypes.GiveAsYouEarn
        Case BatchTypes.PostTaxPayrollGiving
          AdjustmentBatchType = BatchTypes.PostTaxPayrollGiving
        Case Else
          AdjustmentBatchType = BatchTypes.FinancialAdjustment
      End Select

    End Function

    Public Function RefundAllowed() As Boolean
      If BatchType = BatchTypes.CreditCard Or BatchType = BatchTypes.CreditCardWithInvoice Or BatchType = BatchTypes.DebitCard Then
        RefundAllowed = True
      Else
        RefundAllowed = False
      End If
    End Function
    Public Function CheckDeleteAllowed() As DeleteBatchAllowed

      CheckDeleteAllowed = DeleteBatchAllowed.dbYes
      If (NumberOfEntries = 0 And NumberOfTransactions = 0 And BatchTotal = 0 And TransactionTotal = 0) Then
        CheckDeleteAllowed = DeleteBatchAllowed.dbYes
      End If
      If (NumberOfEntries <> 0 And BatchTotal <> 0 And NumberOfTransactions = 0 And TransactionTotal = 0) And PostedToCashBook = False Then
        CheckDeleteAllowed = DeleteBatchAllowed.dbYes
      End If
      If (NumberOfEntries <> 0 And NumberOfTransactions <> 0 And BatchTotal <> 0 And TransactionTotal <> 0) And PostedToCashBook = True Then
        '
      Else
        'Not all are 0
        If (NumberOfEntries <> 0 Or NumberOfTransactions <> 0 Or BatchTotal <> 0 Or TransactionTotal <> 0) And PostedToCashBook = True Then
          CheckDeleteAllowed = DeleteBatchAllowed.dbWarn
        End If
      End If

    End Function

    Public Sub SkipPaymentPlanPayment(ByVal pPlanNumber As Integer, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pAmount As Double, Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False, Optional ByVal pIgnoreDiscounts As Boolean = False)
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)

      With vBT
        .InitFromBatch(mvEnv, Me)
        .ContactNumber = pContactNumber
        .AddressNumber = pAddressNumber
        .TransactionDate = TodaysDate()
        .Amount = pAmount
        .LineTotal = .Amount
        .EligibleForGiftAid = False
      End With
      With vBTA
        .InitFromTransaction(vBT)
        .LineType = "O"
        .Amount = vBT.Amount
        .PaymentPlanNumber = pPlanNumber
        .AcceptAsFull = False
      End With
      mvConn = mvEnv.Connection
      mvCompanyControl = New CompanyControl
      mvCompanyControl.Init(mvEnv)
      mvBatchType = BatchTypes.None
      mvIgnoreDiscountForSkip = pIgnoreDiscounts
      ProcessPaymentPlan(vBT, vBTA, pAmendedBy, pAudit)
      '  mvConn.CommitTransaction
    End Sub
    Public Sub UpdateNumberOfTransactions(ByVal pCount As Integer)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
      vUpdateFields.Add("number_of_transactions", CDBField.FieldTypes.cftLong, "number_of_transactions + " & pCount)
      vUpdateFields.Add("contents_amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
      vUpdateFields.Add("contents_amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      mvEnv.Connection.UpdateRecords("batches", vUpdateFields, vWhereFields)
    End Sub

    Public Sub SetBatchPicked(ByVal pStockSale As Boolean, Optional ByRef pConfirmBatch As Boolean = False)
      'Check to see if all picking lists have been run
      Dim vBTACount As Integer
      Dim vISCount As Integer
      Dim vWhere As String

      If pConfirmBatch = False Then
        'Set from Picking Lists
        If mvEnv.GetConfigOption("fp_stock_multiple_warehouses") = True Then
          'Using multiple warehouses so check all picking lists have been run (different picking list for each warehouse)
          '1. Number of bta stock product lines with stock issued
          vWhere = "bt.batch_number = " & BatchNumber & " AND tt.transaction_type = bt.transaction_type AND tt.transaction_sign <> 'D' AND bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number AND bta.issued > 0 AND bta.product IS NOT NULL AND p.product = bta.product AND p.stock_item = 'Y'"
          vBTACount = mvEnv.Connection.GetCount("batch_transactions bt, transaction_types tt, batch_transaction_analysis bta, products p", Nothing, vWhere)

          '2. Number of bta stock product lines with picking list information
          vWhere = "bta.batch_number = " & BatchNumber & " AND bta.issued <> 0 AND bta.product IS NOT NULL AND p.product = bta.product AND p.stock_item = 'Y'"
          vWhere = vWhere & " AND st.batch_number = bta.batch_number AND st.transaction_number = bta.transaction_number AND st.line_number = bta.line_number"
          vWhere = vWhere & " AND st.product = bta.product AND pld.picking_list_number = st.picking_list_number AND pld.product = st.product AND pld.warehouse = st.warehouse"
          vWhere = vWhere & " AND pl.picking_list_number = pld.picking_list_number AND pl.back_orders = 'N'"
          vISCount = mvEnv.Connection.GetCount("batch_transaction_analysis bta, products p, issued_stock st, picking_list_details pld, picking_lists pl", Nothing, vWhere)

          If vBTACount = vISCount Then
            If vBTACount = 0 Then
              'No issued stock products for this batch
              mvClassFields(BatchFields.Picked).Value = "C" 'Confirmed
            Else
              mvClassFields(BatchFields.Picked).Value = "Y" 'Picked
            End If
          Else
            If vISCount > 0 Then
              'We have actually processed something
              mvClassFields(BatchFields.Picked).Value = "P" 'Part Picked
            End If
          End If
        Else
          '1 Picking List per batch
          If pStockSale Then
            mvClassFields(BatchFields.Picked).Value = "Y" 'Picked
          Else
            mvClassFields(BatchFields.Picked).Value = "C" 'Confirmed
          End If
        End If

      ElseIf pConfirmBatch = True And Picked = "Y" Then
        'Set from Confirm Stock Allocation
        If mvEnv.GetConfigOption("fp_stock_multiple_warehouses") = True Then
          'Using multiple warehouses so check all picking lists are confirmed
          vWhere = "st.batch_number = " & BatchNumber & " AND pld.picking_list_number = st.picking_list_number"
          vWhere = vWhere & " AND pld.product = st.product AND pld.warehouse = st.warehouse"
          vWhere = vWhere & " AND pl.picking_list_number = pld.picking_list_number AND pl.back_orders = 'N'"
          vWhere = vWhere & " AND confirmed_by IS NULL"
          If mvEnv.Connection.GetCount("issued_stock st, picking_list_details pld, picking_lists pl", Nothing, vWhere) = 0 Then
            mvClassFields(BatchFields.Picked).Value = "C" 'Confirmed
          End If
        Else
          '1 Picking List per batch
          mvClassFields(BatchFields.Picked).Value = "C" 'Confirmed
        End If
      End If

    End Sub

    Private Sub ProcessGiftAidSponsorship(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByRef pAmount As Double)
      Dim vFields As New CDBFields
      If pAmount > 0 Then 'Do not create reversal lines
        With vFields
          .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
          .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
          .Add("contact_number", CDBField.FieldTypes.cftLong, pBT.ContactNumber)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric, pAmount)
        End With
        mvConn.InsertRecord("ga_sponsorship_lines_unclaimed", vFields)
      End If

    End Sub

    Private Sub mvCCA_AuthorisingCreditCard(ByRef pMaxTime As Integer, ByRef pTime As Integer) Handles mvCCA.AuthorisingCreditCard
      RaiseEvent AuthorisingCreditCard(pMaxTime, pTime)
    End Sub

    Private Sub ProcessPaymentPlanError(ByVal pPP As PaymentPlan, ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pProduct As Product, ByRef pRate As String, ByVal pAmount As Double, Optional ByVal pFreePaymentPlan As Boolean = False, Optional ByVal pOverpaymentError As Boolean = False)
      'Payment Plan payment unable to be posted entirely against Payment Plan
      Dim vOPH As New OrderPaymentHistory
      Dim vOPS As New OrderPaymentSchedule
      Dim vFields As New CDBFields
      Dim vRS As CDBRecordSet
      Dim vFound As Boolean

      vOPH.Init(mvEnv)
      vOPS.Init(mvEnv)
      With vFields
        .Add("order_number", CDBField.FieldTypes.cftLong, pBTA.PaymentPlanNumber) 'pPP.PlanNumber will be 0 if Payment Plan not found or locked so use pBTA.PaymentPlanNumber
        .Add("batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
        .Add("transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
        .Add("line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
        .Add("amount", CDBField.FieldTypes.cftNumeric, pAmount)
      End With

      vRS = mvConn.GetRecordSet("SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE " & mvConn.WhereClause(vFields))
      If vRS.Fetch() = True Then vOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
      vRS.CloseRecordSet()

      If pOverpaymentError Then
        'Need to update OPH
        pAmount = FixTwoPlaces(pAmount - pPP.Balance)
        If vOPH.Existing Then
          With vOPH
            .SetValues(.BatchNumber, .TransactionNumber, .PaymentNumber, .OrderNumber, (pPP.Balance), .LineNumber, .Balance, CInt(.ScheduledPaymentNumber), .Posted)
          End With

          If Not mvConn.InTransaction Then mvConn.StartTransaction() 'TRANSACTION START HERE

          vOPH.Save()
        End If
      Else
        'Need to update the ops record and delete the oph record
        If vOPH.Existing And vOPH.ScheduledPaymentNumber.Length > 0 Then
          If pPP.PlanNumber > 0 Then
            For Each vOPS In pPP.ScheduledPayments
              If vOPS.ScheduledPaymentNumber = Val(vOPH.ScheduledPaymentNumber) Then vFound = True
              If vFound Then Exit For
            Next vOPS
          End If
          If vFound = False Then
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv, IntegerValue(vOPH.ScheduledPaymentNumber))
          End If
          If vOPS.Existing Then
            'Update OPS so that it no longer shows as unprocessed
            vOPS.SetUnProcessedPayment(False, (vOPH.Amount * -1), (pPP.PlanType = CDBEnvironment.ppType.pptLoan))
            If pPP.CancellationReason.Length > 0 Then
              'If Pay Plan cancelled then set payment as cancelled
              If (vOPS.ScheduledPaymentStatus = OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue Or vOPS.ScheduledPaymentStatus = OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid) And vOPS.AmountOutstanding > 0 Then vOPS.SetCancelled()
            ElseIf pFreePaymentPlan Then
              'If a free Pay Plan then set payment as nothing due or outstanding
              vOPS.Update(vOPS.DueDate, 0, 0, 0, vOPS.ClaimDate, vOPS.RevisedAmount, OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance)
            End If
          End If
        End If
        'Reduce the payment number
        If pPP.PlanNumber > 0 And (pPP.PaymentNumber = vOPH.PaymentNumber) Then pPP.PaymentNumber = pPP.PaymentNumber - 1

        If Not mvConn.InTransaction Then mvConn.StartTransaction() 'TRANSACTION START HERE

        If vOPS.Existing Then vOPS.Save()
        If vOPH.Existing Then vOPH.Delete() 'Delete the OPH as the payment is no longer against the Payment Plan
        If pPP.PlanNumber > 0 Then pPP.SaveChanges()
      End If

      ProcessProduct(pBT, pBTA, pProduct, pRate, "", 1, pAmount, 1, "", 0)

    End Sub

    Private Sub ReverseUnfulfilledCSBackOrder(ByVal pBTA As BatchTransactionAnalysis)
      'If an unfulfilled back order has been reversed then there is no invoice to refund
      'but we need to update the CreditCustomers to show that the amount is no longer on order
      Dim vRS As CDBRecordSet
      Dim vCCU As New CreditCustomer
      Dim vWhereFields As New CDBFields

      If mvEnv.GetConfigOption("fp_use_sales_ledger", True) Then
        With vWhereFields
          .Add("r.batch_number", CDBField.FieldTypes.cftLong, pBTA.BatchNumber)
          .Add("r.transaction_number", CDBField.FieldTypes.cftLong, pBTA.TransactionNumber)
          .Add("r.line_number", CDBField.FieldTypes.cftLong, pBTA.LineNumber)
          .Add("cs.batch_number", CDBField.FieldTypes.cftLong, "r.was_batch_number")
          .Add("cs.transaction_number", CDBField.FieldTypes.cftLong, "r.was_transaction_number")
          .Add("ccu.contact_number", CDBField.FieldTypes.cftLong, "cs.contact_number")
          .Add("ccu.sales_ledger_account", CDBField.FieldTypes.cftLong, "cs.sales_ledger_account")
          .Add("ccu.company", CDBField.FieldTypes.cftCharacter, mvCompanyControl.Company)
        End With

        Dim vBackorderAmount As Double = pBTA.Amount
        If pBTA.Quantity <> 0 AndAlso pBTA.Issued <> 0 AndAlso pBTA.Quantity <> pBTA.Issued Then
          vBackorderAmount = FixTwoPlaces((vBackorderAmount / pBTA.Quantity) * pBTA.Issued)
        End If

        vCCU.Init(mvEnv)
        vRS = mvConn.GetRecordSet("SELECT " & vCCU.GetRecordSetFields(CreditCustomer.CreditCustomerRecordSetTypes.ccurtAll) & " FROM reversals r, credit_sales cs, credit_customers ccu WHERE " & mvConn.WhereClause(vWhereFields))
        If vRS.Fetch() = True Then
          vCCU.InitFromRecordSet(mvEnv, vRS, CreditCustomer.CreditCustomerRecordSetTypes.ccurtAll)
          vCCU.AdjustOnOrder((vBackorderAmount * -1))
          vCCU.Save()
        End If
        vRS.CloseRecordSet()
      End If

    End Sub

    Private Function CheckOverpayment(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pAmount As Double) As Double
      'This will check to see if the Payment Plan payment was an over-payment
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vOverPay As Double

      If mvCompanyControl.OverPaymentProductCode.Length > 0 Then
        vSQL = "SELECT fhd.amount, oph.amount AS oph_amount FROM reversals r, order_payment_history oph, financial_history_details fhd"
        vSQL = vSQL & " WHERE r.batch_number = " & pBatchNumber & " AND r.transaction_number = " & pTransactionNumber
        vSQL = vSQL & " AND r.line_number = " & pLineNumber & " AND oph.batch_number = r.was_batch_number AND oph.transaction_number = r.was_transaction_number"
        vSQL = vSQL & " AND oph.line_number = r.was_line_number AND fhd.batch_number = oph.batch_number AND fhd.transaction_number = oph.transaction_number"
        vSQL = vSQL & " AND fhd.line_number = oph.line_number AND fhd.product = '" & mvCompanyControl.OverPaymentProductCode & "' AND fhd.rate = '" & mvCompanyControl.OverPaymentRate & "'"
        vRS = mvConn.GetRecordSet(vSQL)
        If vRS.Fetch() = True Then
          If FixTwoPlaces(vRS.Fields(1).DoubleValue + vRS.Fields(2).DoubleValue) = FixTwoPlaces(pAmount * -1) Then
            vOverPay = vRS.Fields(1).DoubleValue
          End If
        End If
        vRS.CloseRecordSet()
      End If
      CheckOverpayment = vOverPay

    End Function
    ''' <summary>
    ''' Checks if the transaction is a reversal of a confirmed order payment transaction
    ''' </summary>
    ''' <param name="pBatchNumber"></param>
    ''' <param name="pTransactionNumber"></param>
    ''' <param name="pLineNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckConfirmedTransactionReversal(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer) As Boolean
      Select Case BatchType
        Case BatchTypes.FinancialAdjustment, BatchTypes.CreditCard, BatchTypes.DebitCard
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("oph.batch_number", CDBField.FieldTypes.cftInteger, pBatchNumber)
          vWhereFields.Add("oph.transaction_number", CDBField.FieldTypes.cftInteger, pTransactionNumber)
          vWhereFields.Add("oph.line_number", CDBField.FieldTypes.cftInteger, pLineNumber)

          Dim vAnsiJoins As New AnsiJoins
          vAnsiJoins.Add("reversals r", "r.batch_number", "oph.batch_number", "r.transaction_number", "oph.transaction_number", "r.line_number", "oph.line_number")
          vAnsiJoins.Add("confirmed_transactions ct", "r.was_batch_number", "ct.confirmed_batch_number", "r.was_transaction_number", "ct.confirmed_trans_number")

          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "", "order_payment_history oph", vWhereFields, "", vAnsiJoins)

          Return mvEnv.Connection.GetCountFromStatement(vSQLStatement) > 0
      End Select
      Return False
    End Function

    Private Sub RecreatePaymentSchedule(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pPP As PaymentPlan, ByVal pInAdvance As Boolean)
      'After processing a Payment Plan payment, recreate the payment schedule so there is an OPS for the next payment
      Dim vOPS As OrderPaymentSchedule
      Dim vBalance As Double
      Dim vCreateOPS As Boolean
      Dim vDone As Boolean
      Dim vRenewalDate As String
      Dim vContainsUnprocPayments As Boolean = False

      If ForceCreationOfRegularProvisionalPayment(pBT, pPP) Then
        RecreateRegualrPaymentScheduleWithHistoricReversals(pBT, pBTA, pPP)
      Else
        'Decide whether we need to re-create the payment schedule
        vCreateOPS = (pPP.Balance = 0 And pInAdvance = False)
        If vCreateOPS = False Then
          'Check the existing schedule
          vCreateOPS = True
          For Each vOPS In pPP.ScheduledPayments
            Select Case vOPS.ScheduledPaymentStatus
              Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsProvisional
                If (pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or pPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) And Len(vOPS.ClaimDate) = 0 Then
                  'DD/CCCA claims need new schedule if only payments are unclaimable
                Else
                  vCreateOPS = False
                End If
              Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
                vContainsUnprocPayments = True
              Case Else
                'Do Nothing
            End Select
            If vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then vCreateOPS = False
            If Not vCreateOPS Then Exit For
          Next vOPS
        End If

        'Create provisional payment schedule - requires re-setting some values to force this
        If vCreateOPS Then
          vOPS = New OrderPaymentSchedule
          vOPS.Init(mvEnv, IntegerValue(pBTA.ScheduledPaymentNumber))
          With pPP
            vBalance = .Balance
            If .Balance = 0 And vOPS.Existing = True Then
              If (CDate(vOPS.DueDate) > CDate(mvOrgRenewalDate)) And (.PaymentFrequencyFrequency = 1 And (.PaymentFrequencyInterval >= 1 And .PaymentFrequencyInterval < 12)) Then
                'The OPS was due after the original RenewalDate - Regular monthly Payment Plan only
                'Reset dates to ensure provisional OPS created is dated after vOPS.DueDate
                If mvOrgBalance > 0 And vOPS.ScheduledPaymentStatus = OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsFullyPaid And vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then
                  'This payment is paid and needs to be updated to show CreationReason as BatchPosting
                  With vOPS
                    .Update(.DueDate, .AmountDue, .AmountOutstanding, .ExpectedBalance, .ClaimDate, .RevisedAmount, OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrBatchPosting)
                    .Save()
                  End With
                End If
                vRenewalDate = .RenewalDate
                .RenewalDate = vOPS.DueDate
                vOPS.Init(mvEnv)
                vOPS.CreateInAdvance(mvEnv, pPP, .RenewalAmount, False)
                .RenewalDate = vRenewalDate
                vDone = True
              End If
            ElseIf .Balance = 0 And mvBatchType = BatchTypes.None Then  'BR12656: Always create advance OPS record when balance is 0 and we are skipping a payment
              vOPS.Init(mvEnv)
              vOPS.CreateInAdvance(mvEnv, pPP, .RenewalAmount, False)
              vDone = True
            End If
            If vDone = False AndAlso vContainsUnprocPayments = True Then
              'When fp_record_payment_plan_changes config is set we need to create PaymentPlanChange records as part of regenerating the payment schedule
              'But as we have not processed all the unprocessed payments yet we are not in a position to create the PaymentPlanChange records.
              'So don't do it until all the unprocessed payments have been processed.
              If .Balance > 0 AndAlso mvEnv.GetConfigOption("fp_record_payment_plan_changes", False) = True Then vDone = True
            End If
            If vDone = False Then
              vBalance = .Balance
              .Balance = 0
              .RegenerateScheduledPayments(OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance, .RenewalDate)
              .Balance = vBalance
            End If
          End With
        End If
      End If
    End Sub

    Private Sub RecreateRegualrPaymentScheduleWithHistoricReversals(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pPP As PaymentPlan)
      Dim vOPS As OrderPaymentSchedule
      Dim vNewProvisionalOPS As OrderPaymentSchedule

      vOPS = New OrderPaymentSchedule
      vOPS.Init(mvEnv, IntegerValue(pBTA.ScheduledPaymentNumber))
      vOPS.Update(vOPS.DueDate, vOPS.AmountDue, vOPS.AmountOutstanding, vOPS.ExpectedBalance, vOPS.ClaimDate, vOPS.RevisedAmount, OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrBatchPosting)
      vOPS.Save()
      vNewProvisionalOPS = New OrderPaymentSchedule()
      vNewProvisionalOPS.Init(mvEnv)
      vNewProvisionalOPS.CreateInAdvance(mvEnv, pPP, pPP.RenewalAmount, False)
    End Sub



    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Used by WEB Services only to create a new Batch for a batch-led Trader Application
      'Note: This will NOT create an open_batches record
      Dim vPostToCashBook As Boolean
      Dim vReadyForBanking As Boolean
      InitNewBatch(pEnv) 'This will Init the ClassFields and set the BatchNumber

      vPostToCashBook = False
      vReadyForBanking = False
      Select Case pParams("BatchType").Value
        Case "CC", "DC", "SO", "SP", "CI"
          vReadyForBanking = True
        Case "CS"
          vPostToCashBook = True
      End Select

      With mvClassFields
        .Item(BatchFields.BatchType).Value = pParams("BatchType").Value
        .Item(BatchFields.BatchDate).Value = pParams("BatchDate").Value
        .Item(BatchFields.BankAccount).Value = pParams("BankAccount").Value
        .Item(BatchFields.NumberOfEntries).Value = CStr(pParams("NumberOfEntries").DoubleValue)
        .Item(BatchFields.TransactionTotal).Value = CStr(0)
        .Item(BatchFields.NumberOfTransactions).Value = CStr(0)
        .Item(BatchFields.NextTransactionNumber).Value = CStr(1)
        .Item(BatchFields.ReadyForBanking).Value = BooleanString(vReadyForBanking)
        .Item(BatchFields.PayingInSlipPrinted).Value = "N"
        .Item(BatchFields.PostedToCashBook).Value = BooleanString(vPostToCashBook)
        .Item(BatchFields.DetailCompleted).Value = "N"
        .Item(BatchFields.PostedToNominal).Value = "N"
        .Item(BatchFields.Picked).Value = "N"
        .Item(BatchFields.Product).Value = pParams.ParameterExists("Product").Value
        .Item(BatchFields.Rate).Value = pParams.ParameterExists("Rate").Value
        .Item(BatchFields.Source).Value = pParams.ParameterExists("Source").Value
        .Item(BatchFields.TransactionType).Value = pParams.ParameterExists("TransactionType").Value
        .Item(BatchFields.PaymentMethod).Value = pParams.ParameterExists("PaymentMethod").Value
        .Item(BatchFields.PayingInSlipNumber).Value = pParams.ParameterExists("PayingInSlipNumber").Value
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
          If pParams.Exists("BatchTotal") = True Then
            .Item(BatchFields.CurrencyBatchTotal).Value = CStr(pParams("BatchTotal").DoubleValue)
            'Set BatchTotal to be � (CurrencyBatchTotal will be in the foreign currency)
            pParams("BatchTotal").Value = FixedFormat(System.Math.Round(FixTwoPlaces(pParams("BatchTotal").DoubleValue) / Val(pParams.OptionalValue("CurrencyExchangeRate", "1")), 2))
          End If
          .Item(BatchFields.CurrencyTransactionTotal).Value = CStr(0)
          If pParams.ParameterExists("CurrencyCode").Value.Length > 0 Then
            .Item(BatchFields.CurrencyCode).Value = pParams("CurrencyCode").Value
            .Item(BatchFields.CurrencyExchangeRate).Value = pParams("CurrencyExchangeRate").Value
          Else
            .Item(BatchFields.CurrencyCode).Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode)
            .Item(BatchFields.CurrencyExchangeRate).Value = CStr(1)
          End If
        End If
        .Item(BatchFields.BatchTotal).Value = CStr(pParams("BatchTotal").DoubleValue)
        .Item(BatchFields.BatchCategory).Value = pParams.ParameterExists("BatchCategory").Value
        .Item(BatchFields.BatchCreatedBy).Value = mvEnv.User.UserID
        .Item(BatchFields.BatchCreatedOn).Value = TodaysDate()
        .Item(BatchFields.PostNominal).Value = "N"
        .Item(BatchFields.Provisional).Value = pParams.OptionalValue("Provisional", "N")
        If BatchType = BatchTypes.CAFVouchers Then .Item(BatchFields.AgencyNumber).Value = pParams.ParameterExists("AgencyNumber").Value
        .Item(BatchFields.BatchAnalysisCode).Value = pParams.ParameterExists("BatchAnalysisCode").Value
        .Item(BatchFields.Campaign).Value = pParams.ParameterExists("Campaign").Value
        .Item(BatchFields.Appeal).Value = pParams.ParameterExists("Appeal").Value
        .Item(BatchFields.BankingDate).Value = pParams.ParameterExists("BankingDate").Value
      End With
    End Sub

    Public Sub SetBatchCategory(ByVal pBatchCategory As String)
      mvClassFields.Item(BatchFields.BatchCategory).Value = pBatchCategory
    End Sub

    Private Function ResetOPSForAutoPayment(ByVal pPP As PaymentPlan, ByVal pOPH As OrderPaymentHistory, ByVal pOPS As OrderPaymentSchedule, ByVal pOrigOPSClaimDate As String) As Boolean
      'DD has been added to a Cash PaymentPlan after the RenewalDate and R&R is not run
      'First DD claim is made on zero-balance PP
      'Batch Posting has renewed PP but first OPS (which payment is against) no longer has a ClaimDate
      'see BR12553 for further information
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vFound As Boolean
      Dim vUpdated As Boolean

      If (pPP.DirectDebitStatus <> PaymentPlan.ppYesNoCancel.ppNo Or pPP.CreditCardStatus <> PaymentPlan.ppYesNoCancel.ppNo) And (IsDate(pOPS.ClaimDate) = False And IsDate(pOrigOPSClaimDate) = True) Then
        'This is a DD/CC payment and the OPS no longer has a claim date following the renewal of the PaymentPlan
        For Each vOPS In pPP.ScheduledPayments
          If IsDate(vOPS.ClaimDate) Then
            If CDate(pOrigOPSClaimDate) = CDate(vOPS.ClaimDate) Then vFound = True
          End If
          If vFound Then Exit For
        Next vOPS

        If vFound = False Then
          'If we didn't find an OPS with the original ClaimDate then get first record with ClaimDate after the original date (this should only happen for regular donations)
          For Each vOPS In pPP.ScheduledPayments
            If IsDate(vOPS.ClaimDate) Then
              If CDate(vOPS.ClaimDate) >= CDate(pOrigOPSClaimDate) Then vFound = True
            End If
            If vFound Then Exit For
          Next vOPS
        End If

        If vFound Then
          'Only do this if we found another OPS to allocate the payment against
          vOPS.PaymentAmount = pOPH.Amount
          pOPS.PaymentAmount = FixTwoPlaces(pOPH.Amount * -1)
          vOPS.SetUnProcessedPayment(True, vOPS.PaymentAmount)
          pOPS.SetUnProcessedPayment(False, pOPS.PaymentAmount)
          With pOPH
            .SetValues(.BatchNumber, .TransactionNumber, .PaymentNumber, .OrderNumber, .Amount, .LineNumber, .Balance, vOPS.ScheduledPaymentNumber, .Posted)
            .Save()
          End With
          vOPS.ProcessPayment()
          pOPS.Save()
          vOPS.Save()
          vUpdated = True
        End If
      End If
      ResetOPSForAutoPayment = vUpdated
    End Function

#Region " Loans "

    ''' <summary>Assign a Loan payment between the Capital and Interest products.</summary>
    ''' <param name="pBT">BatchTransaction being processed</param>
    ''' <param name="pBTA">BatchTransactionAnalysis being processed</param>
    ''' <param name="pPP">Loan PaymentPlan being processed</param>
    ''' <param name="pAmount">Amount of payment</param>
    ''' <param name="pExpiry">Subscription expiry date</param>
    Private Sub ProcessLoanDetailsPayment(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pPP As PaymentPlan, ByVal pAmount As Double, ByVal pExpiry As String)
      Dim vPaymentPlanShowPaymentDetails As Boolean = mvEnv.GetConfigOption("fp_pp_show_payment_details", False)

      If pPP.Details.Count() = 0 Then
        PrintLog("Error : Order_details record missing - payment for order: " & pPP.PlanNumber)
      Else
        Dim vLoanCapital As Double = 0
        For Each vPPD As PaymentPlanDetail In pPP.Details
          If vPPD.AccruesInterest Then vLoanCapital += vPPD.Balance
        Next
        Dim vInterestPaid As Nullable(Of Double)
        If BatchType = BatchTypes.FinancialAdjustment AndAlso pBT.TransactionNumber > 1 Then
          'If this is a change of payer then find the original interest allocation otherwise we will just get zero
          Dim vWhereFields As New CDBFields(New CDBField("order_number", pPP.PlanNumber))
          With vWhereFields
            .Add("oph.batch_number", pBT.BatchNumber)
            .Add("oph.transaction_number", pBT.TransactionNumber - 1)
            .Add("oph.amount", CDBField.FieldTypes.cftNumeric, (pBTA.Amount * -1))
          End With
          For Each vPPD As PaymentPlanDetail In pPP.Details
            If vPPD.LoanInterest Then
              vWhereFields.Add("product", vPPD.ProductCode)
              vWhereFields.Add("rate", vPPD.RateCode)
              Exit For
            End If
          Next
          Dim vAnsiJoins As New AnsiJoins
          vAnsiJoins.Add("financial_history_details fhd", "oph.batch_number", "fhd.batch_number", "oph.transaction_number", "fhd.transaction_number", "oph.line_number", "fhd.line_number")
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "SUM(fhd.amount) AS interest_allocated", "order_payment_history oph", vWhereFields, "", vAnsiJoins)
          vSQLStatement.GroupBy = "oph.batch_number, oph.transaction_number, oph.line_number"
          Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          If vRS.Fetch Then vInterestPaid = System.Math.Abs(vRS.Fields(1).DoubleValue)
          vRS.CloseRecordSet()
        End If
        If vInterestPaid.HasValue = False Then
          If pPP.Loan.InterestCapitalisationDate.Length > 0 AndAlso (pPP.LoanLastPaymentDate < CDate(pPP.Loan.LoanCapitalisationDate)) Then
            'Loan has been capitalised and the payment is dated before the Capitalisation date
            Dim vToDate As Date = CDate(pBT.TransactionDate)
            If vToDate >= CDate(pPP.Loan.LoanCapitalisationDate) Then vToDate = pPP.Loan.LoanCapitalisationDate.AddDays(-1)
            vInterestPaid = pPP.CalculateLoanPaymentInterest(pPP.LoanLastPaymentDate, vToDate, FixTwoPlaces(vLoanCapital - pPP.Loan.InterestCapitalisationAmount))
            If CDate(pBT.TransactionDate) >= CDate(pPP.Loan.LoanCapitalisationDate) Then vInterestPaid = FixTwoPlaces(vInterestPaid.Value + pPP.CalculateLoanPaymentInterest(pPP.Loan.LoanCapitalisationDate, CDate(pBT.TransactionDate), vLoanCapital))
          Else
            vInterestPaid = pPP.CalculateLoanPaymentInterest(pPP.LoanLastPaymentDate(), Date.Parse(pBT.TransactionDate), vLoanCapital)
          End If
        End If

        If vInterestPaid > pAmount Then vInterestPaid = pAmount
        Dim vCapitalPaid As Double = FixTwoPlaces(pAmount - vInterestPaid.Value)
        If vCapitalPaid < 0 Then vCapitalPaid = 0

        Dim vPayerVATcategory As String = pPP.Payer.VATCategory
        Dim vVatRate As VatRate
        If vInterestPaid > 0 Then
          Dim vInterestAllocated As Double
          For Each vPPD As PaymentPlanDetail In pPP.Details
            With vPPD
              If .LoanInterest Then
                If .Balance >= vInterestPaid Then
                  vPPD.Balance = FixTwoPlaces(.Balance - vInterestPaid.Value)
                  vInterestAllocated = vInterestPaid.Value
                Else
                  'Should only come here if interest has been capitalised before all payments received
                  vInterestAllocated = .Balance
                  .Balance = 0
                End If
                vInterestPaid = FixTwoPlaces(vInterestPaid.Value - vInterestAllocated)
                If pBTA.AcceptAsFull = True AndAlso .Balance > 0 Then .Balance = 0
                If mvBatchType <> BatchTypes.None Then
                  If .ProductRateIsValid = False Then .SetPrices()
                  vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vInterestAllocated, vVatRate.VatRateCode, CalculateVATAmount(vInterestAllocated, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, pPP.GiverContactNumber)
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vInterestAllocated)
                End If
              End If
            End With
          Next
          If vInterestPaid.Value > 0 Then vCapitalPaid = FixTwoPlaces(vCapitalPaid + vInterestPaid.Value)
        End If

        If vCapitalPaid > 0 Then
          'Add all Loan Capital PPD to separate collection to deal with
          Dim vLoanCapitalLines As New CollectionList(Of PaymentPlanDetail)
          For Each vPPD As PaymentPlanDetail In pPP.Details
            If vPPD.AccruesInterest Then vLoanCapitalLines.Add(vPPD.DetailNumber.ToString, vPPD)
          Next

          Dim vCapitalAllocated As Double = 0
          If mvEnv.GetConfigOption("fp_pay_proportional_details", False) AndAlso vLoanCapitalLines.Count > 1 Then
            'Pay the Loan Capital Details lines proportionately
            Dim vProportion As Integer = pPP.PaymentFrequencyFrequency
            Dim vPrice As Double = 0
            For Each vPPD As PaymentPlanDetail In pPP.Details
              If vCapitalPaid > 0 Then
                With vPPD
                  If .Amount.Length > 0 Then
                    vPrice = DoubleValue(.Amount)
                  Else
                    vPrice = FixTwoPlaces(.CurrentPrice * .Quantity)
                  End If
                  If vPrice > 0 Then
                    vPrice = FixTwoPlaces(vPrice / vProportion) 'The proportionate amount to be paid each time
                    If vPrice > vCapitalPaid Then vPrice = vCapitalPaid 'Account for rounding differences
                  End If
                  If vPrice >= .Balance Then
                    vCapitalAllocated = .Balance
                    .Balance = 0
                  Else
                    vCapitalAllocated = vPrice
                    .Balance = FixTwoPlaces(.Balance - vPrice)
                    If pBTA.AcceptAsFull = True Then
                      .Balance = 0
                      vCapitalAllocated = vCapitalPaid
                    End If
                  End If
                  vCapitalPaid = FixTwoPlaces(vCapitalPaid - vCapitalAllocated) 'vCapitalPaid now amount o/s after this allocation
                  If mvBatchType <> BatchTypes.None Then
                    If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
                    vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                    Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vCapitalAllocated, vVatRate.VatRateCode, CalculateVATAmount(vCapitalAllocated, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, pPP.GiverContactNumber)
                    If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vCapitalAllocated)
                  End If
                End With
              End If
            Next

            'If there is anything left, then ensure that all Capital products have been paid (accounting for rounding errors)
            If vCapitalPaid > 0 Then
              For Each vPPD As PaymentPlanDetail In vLoanCapitalLines
                With vPPD
                  If vCapitalPaid >= .Balance Then
                    vCapitalAllocated = .Balance
                    .Balance = 0
                  Else
                    vCapitalAllocated = vCapitalPaid
                    .Balance = FixTwoPlaces(.Balance - vCapitalAllocated)
                  End If
                  vCapitalPaid = FixTwoPlaces(vCapitalPaid - vCapitalAllocated)
                  If mvBatchType <> BatchTypes.None Then
                    If vPPD.ProductRateIsValid = False Then vPPD.SetPrices()
                    vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(vPPD.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                    Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vCapitalAllocated, vVatRate.VatRateCode, CalculateVATAmount(vCapitalAllocated, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, pPP.GiverContactNumber)
                    If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vCapitalAllocated)
                  End If
                End With
              Next
            End If
          Else
            'Pay the Loan Capital Detail lines 1 at a time
            For Each vPPD As PaymentPlanDetail In vLoanCapitalLines
              vCapitalAllocated = 0
              With vPPD
                If vCapitalPaid >= .Balance Then
                  vCapitalAllocated = .Balance
                  .Balance = 0
                Else
                  vCapitalAllocated = vCapitalPaid
                  .Balance = FixTwoPlaces(.Balance - vCapitalPaid)
                End If
                If pBTA.AcceptAsFull Then .Balance = 0
                vCapitalPaid = FixTwoPlaces(vCapitalPaid - vCapitalAllocated)
                If mvBatchType <> BatchTypes.None Then
                  If .ProductRateIsValid = False Then .SetPrices()
                  vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                  Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                  WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vCapitalAllocated, vVatRate.VatRateCode, CalculateVATAmount(vCapitalAllocated, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, pPP.GiverContactNumber)
                  If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vCapitalAllocated)
                End If
              End With
              If vCapitalPaid = 0 Then Exit For
            Next
          End If
        End If

        If vCapitalPaid > 0 OrElse pBTA.AcceptAsFull = True Then
          'Pay off any Loan fees
          Dim vFeesAllocated As Double
          For Each vPPD As PaymentPlanDetail In pPP.Details
            With vPPD
              If .AccruesInterest = False AndAlso .LoanInterest = False Then
                If vCapitalPaid > 0 Then
                  If vCapitalPaid >= .Balance Then
                    vFeesAllocated = .Balance
                    .Balance = 0
                  Else
                    vFeesAllocated = vCapitalPaid
                    .Balance = FixTwoPlaces(.Balance - vFeesAllocated)
                    If pBTA.AcceptAsFull = True Then
                      .Balance = 0
                      vFeesAllocated = vCapitalPaid
                    End If
                  End If
                  vCapitalPaid = FixTwoPlaces(vCapitalPaid - vFeesAllocated)
                  If .Product.Subscription = True AndAlso mvBatchType <> BatchTypes.None Then
                    Select Case CheckProcessSubscriptions(pPP, vPPD)
                      Case "Y"
                        ProcessSubscriptions(pPP, vPPD, .ContactNumber, .ProductCode, pExpiry)
                      Case "T"
                        TerminateSubscriptions(pPP.PlanNumber, .ContactNumber, .ProductCode, CDate(mvOrgRenewalDate).AddDays(-1))
                      Case "D"
                        DeleteSubscriptions(pPP.PlanNumber, .ContactNumber, .ProductCode)
                      Case Else
                        'Do nothing
                    End Select
                  End If
                  If mvBatchType <> BatchTypes.None Then
                    If .ProductRateIsValid = False Then .SetPrices()
                    vVatRate = mvEnv.VATRate(.Product.ProductVatCategory, IIf(.VATExclusive, vPayerVATcategory, pBT.ContactVatCategory).ToString)
                    Dim vDistributionCode As String = If(mvFHDDistCodeOriginPPP = FHDDistCodeOriginPPP.PaymentPlanDetails, .DistributionCode, pBTA.DistributionCode)
                    WriteFinancialHistoryAnalysis(pBT, pBTA, .ProductCode, .RateCode, vDistributionCode, CInt(.Quantity), vFeesAllocated, vVatRate.VatRateCode, CalculateVATAmount(vFeesAllocated, vVatRate.CurrentPercentage(pBT.TransactionDate)), .Source, pPP.GiverContactNumber)
                    If vPaymentPlanShowPaymentDetails AndAlso vPPD.HasPriceInfo Then WritePaymentPlanHistoryDetails(pBTA.OrderPaymentHistory, vPPD, vFeesAllocated)
                  End If
                End If
              End If
              If pBTA.AcceptAsFull Then .Balance = 0
            End With
          Next
        End If

        If CDate(pBT.TransactionDate) < CDate(pPP.Loan.LoanCapitalisationDate) AndAlso Today >= CDate(pPP.Loan.LoanCapitalisationDate) Then
          pPP.CalculateLoanInterest(mvEnv.User.UserID, True, TodaysDate, True)
        End If
      End If

    End Sub

    ''' <summary>Assign reversal of Loan payment against Capital and Interest products.</summary>
    ''' <param name="pBT">Batchtransaction being processed</param>
    ''' <param name="pBTA">BatchTransactionAnalysis being processed</param>
    ''' <param name="pPP">Loan PaymentPlan being processed</param>
    ''' <param name="pAmount">Reversal payment amount</param>
    ''' <param name="pExpiry">Subscription expiry date</param>
    Private Sub ProcessLoanDetailsReversal(ByVal pBT As BatchTransaction, ByVal pBTA As BatchTransactionAnalysis, ByVal pPP As PaymentPlan, ByVal pAmount As Double, ByVal pExpiry As String)

      If pPP.Details.Count() = 0 Then PrintLog("Error : Order_details record missing - reversal for order: " & pPP.PlanNumber)

      pAmount = -pAmount
      pAmount += GetOriginalWriteOff(pBTA.BatchNumber, pBTA.TransactionNumber, pBTA.LineNumber, pBTA.AcceptAsFull)

      'First reverse any Interest paid (need to find this amount)
      Dim vInterestPaid As Double = 0
      Dim vWhereFields As New CDBFields(New CDBField("r.batch_number", pBTA.BatchNumber))
      vWhereFields.Add("r.transaction_number", pBTA.TransactionNumber)
      vWhereFields.Add("r.line_number", pBTA.LineNumber)
      For Each vPPD As PaymentPlanDetail In pPP.Details
        If vPPD.LoanInterest Then
          vWhereFields.Add("product", vPPD.ProductCode)
          vWhereFields.Add("rate", vPPD.RateCode)
          Exit For
        End If
      Next
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("financial_history_details fhd", "r.was_batch_number", "fhd.batch_number", "r.was_transaction_number", "fhd.transaction_number", "r.was_line_number", "fhd.line_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "SUM(fhd.amount) AS amount", "reversals r", vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = "fhd.batch_number, fhd.transaction_number, fhd.line_number"
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      If vRS.Fetch Then vInterestPaid = vRS.Fields(1).DoubleValue
      vRS.CloseRecordSet()

      If vInterestPaid > 0 Then
        For Each vPPD As PaymentPlanDetail In pPP.Details
          If vPPD.LoanInterest Then
            vPPD.Balance = FixTwoPlaces(vPPD.Balance + vInterestPaid)
            pAmount = FixTwoPlaces(pAmount - vInterestPaid)
          End If
        Next
      End If

      'Deal with any Capital products
      Dim vCapitalPaid As Double = pAmount
      If vCapitalPaid > 0 Then
        'Add all Loan Capital PPD to separate collection to deal with
        Dim vLoanCapitalLines As New CollectionList(Of PaymentPlanDetail)
        For Each vPPD As PaymentPlanDetail In pPP.Details
          If vPPD.AccruesInterest Then vLoanCapitalLines.Add(vPPD.DetailNumber.ToString, vPPD)
        Next

        'Dim vCapitalAllocated As Double = 0
        If mvEnv.GetConfigOption("fp_pay_proportional_details", False) AndAlso vLoanCapitalLines.Count > 1 Then
          'Reverse the payment proportionally across all detail lines
          Dim vProportion As Integer = pPP.PaymentFrequencyFrequency
          Dim vPrice As Double
          For Each vPPD As PaymentPlanDetail In vLoanCapitalLines
            If vCapitalPaid > 0 Then
              With vPPD
                If .Amount.Length > 0 Then
                  vPrice = DoubleValue(.Amount)
                Else
                  vPrice = FixTwoPlaces(.CurrentPrice * .Quantity)
                End If
                If vPrice > 0 Then
                  vPrice = FixTwoPlaces(vPrice / vProportion) 'The proportionate amount to be paid each time
                  If vPrice > vCapitalPaid Then vPrice = vCapitalPaid 'Account for rounding differences
                End If
                If vPrice > 0 Then
                  If vPrice > FixTwoPlaces(DoubleValue(.Amount) - .Balance) Then vPrice = Fix(DoubleValue(.Amount) - .Balance)
                  .Balance = FixTwoPlaces(.Balance + vPrice)
                End If
                vCapitalPaid = FixTwoPlaces(vCapitalPaid - vPrice)
              End With
            End If
          Next
        Else
          'Reverse the Loan Capital Detail lines 1 at a time in reverse order
          Dim vCapitalAllocated As Double = 0
          For vItemNumber As Integer = vLoanCapitalLines.Count - 1 To 0 Step -1
            If vCapitalPaid > 0 Then
              Dim vPPD As PaymentPlanDetail = vLoanCapitalLines(vItemNumber)
              With vPPD
                Dim vPrice As Double
                If .Amount.Length > 0 Then
                  vPrice = DoubleValue(.Amount)
                Else
                  vPrice = FixTwoPlaces(.CurrentPrice * .Quantity)
                End If
                If .Balance < vPrice Then
                  If vCapitalPaid >= FixTwoPlaces(vPrice - .Balance) Then
                    vCapitalAllocated = FixTwoPlaces(vPrice - .Balance)
                  Else
                    vCapitalAllocated = vCapitalPaid
                  End If
                  .Balance = FixTwoPlaces(.Balance + vCapitalAllocated)
                  vCapitalPaid = FixTwoPlaces(vCapitalPaid - vCapitalAllocated)
                End If
              End With
            End If
          Next
        End If
      End If

      If vCapitalPaid > 0 Then
        'Deal with any fees that may have been paid (in reverse order)
        Dim vFeesAllocated As Double
        For vItemNumber As Integer = pPP.Details.Count To 1 Step -1
          If vCapitalPaid > 0 Then
            Dim vPPD As PaymentPlanDetail = DirectCast(pPP.Details(vItemNumber), PaymentPlanDetail)
            With vPPD
              If .AccruesInterest = False AndAlso .LoanInterest = False Then
                Dim vPrice As Double
                If .Amount.Length > 0 Then
                  vPrice = DoubleValue(.Amount)
                Else
                  vPrice = FixTwoPlaces(.CurrentPrice * .Quantity)
                End If
                If vPrice >= 0 Then
                  If vCapitalPaid >= FixTwoPlaces(vPrice - .Balance) Then
                    vFeesAllocated = FixTwoPlaces(vPrice - .Balance)
                  Else
                    vFeesAllocated = vCapitalPaid
                  End If
                  .Balance = FixTwoPlaces(.Balance + vFeesAllocated)
                Else
                  'Discount line - increase vCapitalPaid by discount
                  If pPP.Balance >= pPP.RenewalAmount AndAlso FixTwoPlaces(pPP.Balance - vCapitalPaid) < pPP.RenewalAmount Then
                    If vPrice < .Balance Then
                      vFeesAllocated = FixTwoPlaces(vPrice - .Balance)
                      .Balance = vPrice
                    End If
                  End If
                End If
                vCapitalPaid = FixTwoPlaces(vCapitalPaid - vFeesAllocated)
              End If
            End With
          End If
        Next
      End If

    End Sub

#End Region

    Public Sub WritePaymentPlanHistoryDetails(ByVal pOPH As OrderPaymentHistory, ByVal pPPD As PaymentPlanDetail, ByVal pAmountPaid As Double)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
        Dim vPPHistoryDetail As New PaymentPlanHistoryDetail(mvEnv)
        Dim vParams As New CDBParameters

        vParams.Add("OrderNumber", CDBField.FieldTypes.cftInteger, pOPH.OrderNumber.ToString)
        vParams.Add("PaymentNumber", CDBField.FieldTypes.cftInteger, pOPH.PaymentNumber.ToString)
        vParams.Add("DetailNumber", CDBField.FieldTypes.cftInteger, pPPD.DetailNumber.ToString)
        vPPHistoryDetail.Init(vParams)
        If vPPHistoryDetail.Existing Then
          vPPHistoryDetail.UpdateFromPaymentPlanPayment(pAmountPaid, pPPD.Balance)
        Else
          vPPHistoryDetail.CreateFromPaymentPlanPayment(pOPH, pPPD, pAmountPaid)
        End If
        vPPHistoryDetail.Save(mvEnv.User.UserID, True)
      End If
    End Sub

    Public Function WriteInvoiceAndDetails(ByVal pBT As BatchTransaction, ByVal pCreditSale As CreditSale, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pSetInvoiceNumber As Boolean, ByVal pSetInvoiceDate As Boolean) As Invoice
      Return WriteInvoiceAndDetails(pBT, pCreditSale, pCompany, pSalesLedgerAccount, pSetInvoiceNumber, pSetInvoiceDate, False, False)
    End Function

    Public Function WriteInvoiceAndDetails(ByVal pBatchNumber As Integer, pTransNumber As Integer, ByVal pSetInvoiceNumber As Boolean, ByVal pSetInvoiceDate As Boolean, ByVal pIgnoreStockSales As Boolean, ByVal pUnallocateCreditNote As Boolean) As Invoice
      Dim vBT As New BatchTransaction(mvEnv)
      vBT.Init(pBatchNumber, pTransNumber)
      Dim vCompanyControl As New CompanyControl()
      vCompanyControl.InitFromBankAccount(mvEnv, BankAccount)
      Dim vCreditSale As New CreditSale(mvEnv)
      vCreditSale.Init(pBatchNumber, pTransNumber)
      Return WriteInvoiceAndDetails(vBT, Nothing, vCompanyControl.Company, vCreditSale.SalesLedgerAccount, pSetInvoiceNumber, pSetInvoiceDate, pIgnoreStockSales, pUnallocateCreditNote)
    End Function

    Public Function WriteInvoiceAndDetails(ByVal pBT As BatchTransaction, ByVal pCreditSale As CreditSale, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pSetInvoiceNumber As Boolean, ByVal pSetInvoiceDate As Boolean, ByVal pIgnoreStockSales As Boolean, ByVal pUnallocateCreditNote As Boolean) As Invoice
      'Make sure the analysis is read if it has not been read already
      If pBT.Analysis.Count = 0 Then pBT.InitBatchTransactionAnalysis(pBT.BatchNumber, pBT.TransactionNumber)

      If pIgnoreStockSales = False Then
        If pCreditSale.StockSale Then Return Nothing 'only write invoice for non stock sales
        Dim vIssued As Integer = 0
        For Each vBTA As BatchTransactionAnalysis In pBT.Analysis
          If vBTA.Issued > 0 Then vIssued = vIssued + vBTA.Issued
        Next vBTA
        If Not vIssued > 0 Then Return Nothing
      End If

      'Write invoice
      Dim vInvoice As Invoice = WriteInvoice(pBT.BatchNumber, pBT, pBT.ContactNumber, pBT.AddressNumber, pCompany, pSalesLedgerAccount, pSetInvoiceDate, pUnallocateCreditNote)
      'Write Invoice Details
      For Each vBTA As BatchTransactionAnalysis In pBT.Analysis
        WriteInvoiceDetails(vBTA)
      Next
      If pSetInvoiceNumber AndAlso vInvoice IsNot Nothing AndAlso vInvoice.InvoiceNumber.Length = 0 Then
        vInvoice.SetInvoiceNumber(True)
        vInvoice.Save()
      End If
      Return vInvoice
    End Function

#Region "Count"
    Public ReadOnly Property SelectionCount(ByVal pParams As CDBParameters) As Integer
      Get
        'Return the number of selected records
        Dim vWhereFields As New CDBFields

        vWhereFields.Add("posted_to_nominal", CDBField.FieldTypes.cftCharacter, "N") 'changed post_nominal with posted_to_nominal

        If pParams.Exists("StartBatch") AndAlso pParams.Exists("EndBatch") Then
          vWhereFields.Add("batch_number", pParams("StartBatch").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenFrom)
          vWhereFields.Add("batch_number#2", pParams("EndBatch").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenTo)
        End If

        If pParams.Exists("BatchType") Then
          vWhereFields.Add("batch_type", CDBField.FieldTypes.cftCharacter, pParams("BatchType").Value)
        End If

        Return mvEnv.Connection.GetCount("batches", vWhereFields)
      End Get
    End Property
#End Region
    ''' <summary>
    ''' Returns True if a new provisional payment record is required on tbhe order payments schedule table.
    ''' </summary>
    ''' <param name="pBatchTransaction"></param>
    ''' <param name="pPaymentPlan"></param>
    ''' <returns></returns>
    ''' <remarks>When the last due payment (provisional) is made for a regular DD paymet a new provisional payemnt must be created.</remarks>
    Private Function ForceCreationOfRegularProvisionalPayment(pBatchTransaction As BatchTransaction, pPaymentPlan As PaymentPlan) As Boolean


      Dim vOPHSql As SQLStatement
      Dim vOPHWhereFields As CDBFields
      Dim vOPHDataTable As DataTable
      Dim vOPSSql As SQLStatement
      Dim vOPSWhereFields As CDBFields
      Dim vOPSDataTable As DataTable
      Dim vLastScheduledPaymentNumber As Integer
      Dim vReturn As Boolean

      If pBatchTransaction.TransactionType = "P" _
        AndAlso pPaymentPlan.PaymentFrequencyFrequency * pPaymentPlan.PaymentFrequencyInterval < 12 _
        AndAlso pPaymentPlan.PlanType = CDBEnvironment.ppType.pptDD Then

        ' pPaymentPlan.PaymentFrequencyFrequency * pPaymentPlan.PaymentFrequencyInterval < 12 identifies a regular payment, monthly, quarterly or biannually

        ' Get the order payment schedule records
        vOPSWhereFields = New CDBFields(New CDBField("order_number", pPaymentPlan.OrderNumber))
        vOPSSql = New SQLStatement(Me.Environment.Connection, "scheduled_payment_number,scheduled_payment_status,amount_outstanding,schedule_creation_reason", "order_payment_schedule", vOPSWhereFields, "scheduled_payment_number DESC")
        vOPSDataTable = vOPSSql.GetDataTable
        If vOPSDataTable.Rows.Count > 0 Then
          If CDbl(vOPSDataTable.Compute("SUM(amount_outstanding)", "scheduled_payment_status NOT IN ('V','S')")) > 0 Then ' filter just incase
            ' There is atleast one reversal in the order payment schedule.  

            'the datatable is ordered by scheduled_payment_number descending so the first record in the datatable is the most recent, not interested if it the first payment, or is provisional
            If vOPSDataTable.Rows(0).Item("scheduled_payment_status").ToString <> "V" And vOPSDataTable.Rows(0).Item("schedule_creation_reason").ToString <> "NP" Then
              vLastScheduledPaymentNumber = CInt(vOPSDataTable.Rows(0).Item("scheduled_payment_number"))

              ' Match the Order payment schedule to the batch tranactions via the order payment history table
              vOPHWhereFields = New CDBFields(New CDBField("batch_number", pBatchTransaction.BatchNumber))
              vOPHWhereFields.Add(New CDBField("transaction_number", pBatchTransaction.TransactionNumber))
              vOPHWhereFields.Add(New CDBField("scheduled_payment_number", vLastScheduledPaymentNumber))

              vOPHSql = New SQLStatement(Me.Environment.Connection, "batch_number,transaction_number,line_number,scheduled_payment_number", "order_payment_history", vOPHWhereFields)
              vOPHDataTable = vOPHSql.GetDataTable
              If vOPHDataTable.Rows.Count = 1 Then
                vReturn = True ' The previous provisional payment is being paid
              Else
                vReturn = False ' The previous provisional payment is being paid
              End If
            Else
              vReturn = False 'The Last Order payment schedule record is a provisional or is the first payment, so do nothing
            End If
          Else
            vReturn = False 'There are no reversals
          End If
        Else
          vReturn = False 'There are no order payment schedule records
        End If
      Else
        vReturn = False 'Not a regular DD payment
      End If
      Return vReturn
    End Function
    ''' <summary>
    ''' Determines whether any transactions in this batch are associated with an Invoice, Credit Note or Unallocated Payment in Sales Ledger
    ''' </summary>
    ''' <returns>True if any transaction in this batch are linked to an Invoice, Credit Note or Unallocated Payment otherwise false </returns>
    ''' <remarks>It is transactions that are linked to invoices but there can be upto 9999 transactions in a batch so we don't want to ask each transaction</remarks>
    Public ReadOnly Property HasInvoices As Boolean
      Get
        Dim vHasInvoices As Boolean = False
        Dim vWhereFields As New CDBFields

        If Me.Existing Then
          vWhereFields.Add("batch_number", Me.BatchNumber)
          If mvEnv.Connection.GetCount("invoices", vWhereFields) > 0 Then
            vHasInvoices = True
          Else
            vWhereFields.Add("allocation_batch_number", Me.BatchNumber, CDBField.FieldWhereOperators.fwoOR)
            If mvEnv.Connection.GetCount("invoice_payment_history", vWhereFields) > 0 Then
              vHasInvoices = True
            End If
          End If
        End If
        Return vHasInvoices
      End Get
    End Property

    ''' <summary>
    ''' Determines whether a transactions in this batch is associated with an Invoice, Credit Note or Unallocated Payment in Sales Ledger
    ''' </summary>
    ''' <returns>True if tranaction is linked to an Invoice, Credit Note or Unallocated Payment otherwise false </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HasInvoices(pTransactionNumber As Integer) As Boolean
      Get
        Dim vHasInvoices As Boolean = False
        Dim vWhereFields As New CDBFields

        If Me.Existing Then

          vWhereFields.Add("batch_number", Me.BatchNumber)
          vWhereFields.Add("transaction_number", pTransactionNumber)
          If mvEnv.Connection.GetCount("invoices", vWhereFields) > 0 Then
            vHasInvoices = True
          Else
            vWhereFields = New CDBFields
            vWhereFields.Add("batch_number", Me.BatchNumber, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
            vWhereFields.Add("transaction_number", pTransactionNumber, CDBField.FieldWhereOperators.fwoCloseBracket)
            vWhereFields.Add("allocation_batch_number", Me.BatchNumber, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
            vWhereFields.Add("allocation_transaction_number", pTransactionNumber, CDBField.FieldWhereOperators.fwoCloseBracketTwice)
            If mvEnv.Connection.GetCount("invoice_payment_history", vWhereFields) > 0 Then
              vHasInvoices = True
            End If
          End If
        End If
        Return vHasInvoices
      End Get
    End Property
  End Class

End Namespace
