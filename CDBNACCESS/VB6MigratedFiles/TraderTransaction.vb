Imports Advanced.LanguageExtensions

Namespace Access
  Public Class TraderTransaction

    Private mvEnv As CDBEnvironment
    Private mvBatch As Batch
    Private mvBT As BatchTransaction
    Private mvFinancialAdjustment As Batch.AdjustmentTypes
    Private mvTraderAnalysisLines As TraderAnalysisLines 'Collection of Analysis lines
    Private mvPaymentPlanDetails As TraderPaymentPlanDetails 'Collection of Order Detail Lines
    Private mvPaymentPlan As PaymentPlan 'Used by Smart Client Trader only
    Private mvUpdatedOPS As Collection 'Used by Smart Client returning OPS data that may have been updated by the user
    Private mvOutstandingOPS As Collection 'used to hold the outstanding ops for the payments
    Private mvSummaryMembers As CDBCollection 'Collection of Members, used by Smart Client to hold MembershipMembersSummary grid data
    Private mvPurchaseOrderDetails As Collection
    Private mvPurchaseOrderPayments As Collection
    Private mvPurchaseInvoiceDetails As Collection
    Private mvActivities As Collection
    Private mvSuppressions As Collection
    Private mvIncentivesTable As CDBDataTable
    Private mvBatchInvoices As Collection
    Private mvEventBookingLines As CDBDataTable 'Used by Smart Client Event Bookings when using the Event Pricing Matrix
    Private mvExamBookingLines As CDBDataTable 'Used by Smart Client Exam Bookings
    Private mvRemovedSchPayments As CDBCollection
    Private mvCMTOldPPDLines As Collection
    Private mvActions As CollectionList(Of Action)

    Private mvExisting As Boolean
    Private mvBatchNumber As Integer
    Private mvTransactionNumber As Integer
    Private mvCurrencyCode As String = ""
    Private mvExchangeRate As Double
    Private mvPostCashBook As Boolean
    Private mvPostCashBookSet As Boolean
    Private mvBatchAnalysisCode As String = ""
    Private mvTransactionOrigin As String = ""

    'Credit Sales
    Private mvCreditCustomer As CreditCustomer
    Private mvCreditSale As CreditSale
    Private mvCSCredCustomerChanged As Boolean
    Private mvCSCompany As String = ""
    Private mvCSCreditNote As Boolean
    Private mvCSStockSale As Boolean
    Private mvCSInvoiceCreated As Boolean
    Private mvUseSalesLedger As Boolean
    Private mvServiceBookingCredits As Boolean

    'Provisional Batches
    Private mvProvBatch As Batch
    Private mvProvBTColl As CDBCollection
    Private mvProvBTAColl As CDBCollection
    Private mvConfirmedTrans As New CDBCollection
    Private mvCheckProvTrans As Boolean

    'Card Sales
    Private mvCardSale As CardSale
    Private mvDebitCard As Boolean
    Private WithEvents mvCCA As CreditCardAuthorisation

    Private mvTraderInvoiceLines As Collection

    Private mvOriginalOPS As OrderPaymentSchedule

    Public Event StartCCAuthorisation()
    Public Event CCAuthorisationProgress(ByVal pMessage As String)
    Public Event AuthorisingCreditCard(ByVal pMaxTime As Integer, ByVal pTime As Integer)
    Public Event EndCCAuthorisation()

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub SetUpBatchAndTransaction(ByVal pBatchType As Batch.BatchTypes, ByVal pBankAccount As String, ByRef pPaymentMethod As String, Optional ByVal pBatchDate As String = "", Optional ByVal pBatchCategory As String = "", Optional ByVal pProvisional As Batch.ProvisionalOrConfirmed = Batch.ProvisionalOrConfirmed.Confirmed, Optional ByVal pTransReference As String = "")
      'Get the Batch and BatchTransaction

      mvBatch = New Batch(mvEnv)
      If mvBatchNumber > 0 And mvFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
        'Initialise the Batch using the BatchNumber
        mvBatch.Init(mvBatchNumber)
        If mvTransactionNumber < 1 Then mvTransactionNumber = mvBatch.AllocateTransactionNumber
      Else
        'Initialise an open batch
        If mvPostCashBookSet = False Then
          Select Case mvFinancialAdjustment
            Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              mvPostCashBook = True
            Case Else
              mvPostCashBook = False
          End Select
          mvPostCashBookSet = True
        End If
        '(Ref BR19091) pPostToCashBook 'Y' value passed from Smart Client is understood to mean Post Cash Book, however the following mvBatch.InitOpenBatch routine
        'treats this value to mean the opposite. To rectify this mvPostCashBook value is passed as pPostToCashBook.     
        mvBatch.InitOpenBatch(Nothing, pProvisional, pBatchType, pBankAccount, pPaymentMethod, mvPostCashBook, Batch.BatchTypes.None, mvCurrencyCode, mvExchangeRate, pBatchCategory, pBatchDate, True, pTransReference, mvBatchAnalysisCode)
        mvBatchNumber = mvBatch.BatchNumber
        mvTransactionNumber = mvBatch.AllocateTransactionNumber
      End If
      mvBT = New BatchTransaction(mvEnv)
      mvBT.InitForUpdate(mvBatchNumber, mvTransactionNumber, mvExisting)
    End Sub

    Private Sub AddTransaction(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pTransactionDate As String, ByVal pPayMethod As String, ByVal pReceipt As String, ByRef pEligibleGiftAid As Boolean, Optional ByVal pMailing As String = "", Optional ByVal pMailingContactNumber As String = "", Optional ByVal pMailingAddressNumber As String = "", Optional ByVal pReference As String = "", Optional ByVal pNotes As String = "", Optional ByRef pBankDetailsNo As Integer = 0, Optional ByVal pFATransType As String = "")
      'Populate the BatchTransaction fields
      Dim vTransType As String

      If mvCSCreditNote Or mvServiceBookingCredits Then
        'Credit Sales - Credit Note or Service Booking Credit
        vTransType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSCreditTransType)
      ElseIf mvBatch.BatchType = Batch.BatchTypes.CreditSales Then
        'Credit Sales
        vTransType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSTransType)
      ElseIf mvFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment And pFATransType.Length > 0 Then
        'Financial Adjustment - Re-analysis
        vTransType = pFATransType
      Else
        vTransType = "P"
      End If

      With mvBT
        .ContactNumber = pContactNumber
        .AddressNumber = pAddressNumber
        .TransactionDate = pTransactionDate
        .TransactionType = vTransType
        .PaymentMethod = pPayMethod
        .Receipt = pReceipt
        .EligibleForGiftAid = pEligibleGiftAid
        .Mailing = pMailing
        If Len(pMailingContactNumber) > 0 Then
          .MailingContactNumber = CInt(pMailingContactNumber)
          .MailingAddressNumber = CInt(pMailingAddressNumber)
        End If
        .Reference = pReference
        .Notes = pNotes
        If pBankDetailsNo > 0 Then .BankDetailsNumber = pBankDetailsNo
        If Len(mvTransactionOrigin) > 0 And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataTransactionOrigins) Then .TransactionOrigin = mvTransactionOrigin

      End With

    End Sub

    Private Function BuildNotes(ByVal pAdditionalRef1 As String, ByVal pAdditionalRef2 As String, ByVal pAdditionalRef1Caption As String, ByVal pAdditionalRef2Caption As String) As String
      'Build the Notes field from the Additional References for provisional transactions
      Dim vNotes As String = ""
      If Len(pAdditionalRef1) > 0 Then
        vNotes = pAdditionalRef1Caption & " " & pAdditionalRef1
        If Len(pAdditionalRef2) > 0 Then
          vNotes = vNotes & vbLf & pAdditionalRef2Caption & " " & pAdditionalRef2
        End If
      End If
      BuildNotes = vNotes
    End Function

    Private Sub CreateNewProvisionalBatches(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      'When confirming SaleOrReturn (SR) batches it may be neccessary to create a new provisional SR Batch
      'Check that all analysis lines in the provisional transaction are in the confirmed transaction
      'And if not then create a new analysis line in a new provisional batch
      Dim vProvAnalysis As New Collection 'Collection of provisional analysis lines
      Dim vProvBatch As New Batch(mvEnv)
      Dim vProvBT As New BatchTransaction(mvEnv)
      Dim vProvBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vCT As ConfirmedTransaction 'New Confirmed Transaction
      Dim vNewBT As BatchTransaction = Nothing 'New provisional transaction
      Dim vNewBTA As BatchTransactionAnalysis 'New provisional analysis line
      Dim vFields As New CDBFields
      Dim vSQL As String
      Dim vRS As CDBRecordSet
      Dim vFound As Boolean

      'Retrieve all the provisional transaction data
      vProvBatch.Init()
      vProvBT.Init()
      vProvBTA.Init()
      vSQL = "SELECT " & vProvBatch.GetRecordSetFields(mvEnv, Batch.BatchRecordSetTypes.brtType) & ", " & vProvBT.GetRecordSetFields() & "," & vProvBTA.GetRecordSetFields()
      vSQL = Replace(vSQL, "bt.batch_number", "bt.batch_number AS bt_batch_number")
      vSQL = Replace(vSQL, "bta.batch_number", "bta.batch_number AS bta_batch_number")
      vSQL = Replace(vSQL, "bta.transaction_number", "bta.transaction_number AS bta_transaction_number")
      vSQL = Replace(vSQL, "bt.amount", "bt.amount AS bt_amount")
      vSQL = Replace(vSQL, "bt.currency_amount", "bt.currency_amount AS bt_currency_amount")
      vSQL = Replace(vSQL, "bt.contact_number", "bt.contact_number AS bt_contact_number")
      vSQL = Replace(vSQL, "bt.address_number", "bt.address_number AS bt_address_number")
      vSQL = Replace(vSQL, "bt.notes", "bt.notes AS bt_notes")
      vSQL = Replace(Replace(vSQL, "bt.amended_by,", ""), "bt.amended_on,", "")
      vSQL = vSQL & " FROM batches b, batch_transactions bt, batch_transaction_analysis bta WHERE b.batch_number = " & pBatchNumber & " AND batch_type = 'SR'"
      vSQL = vSQL & " AND bt.batch_number = b.batch_number AND bt.transaction_number = " & pTransactionNumber & " AND bta.batch_number = bt.batch_number"
      vSQL = vSQL & " AND bta.transaction_number = bt.transaction_number ORDER BY line_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      If vRS.Fetch() = True Then
        vProvBatch.InitFromRecordSet(mvEnv, vRS, Batch.BatchRecordSetTypes.brtType)
        vProvBT.InitFromRecordSet(vRS)
        Do
          vProvBTA = New BatchTransactionAnalysis(mvEnv)
          vProvBTA.InitFromRecordSet(vRS)
          vProvAnalysis.Add(vProvBTA)
        Loop While vRS.Fetch() = True
      End If
      vRS.CloseRecordSet()

      If vProvBatch.Existing Then
        'Check for analysis lines that were deleted during confirmation
        For Each vProvBTA In vProvAnalysis
          vFound = False
          For Each vNewBTA In mvBT.Analysis
            If vNewBTA.ProductNumber = vProvBTA.ProductNumber Then vFound = True
            If vFound Then
              If vNewBTA.Quantity = 0 And vProvBTA.ProductNumber > 0 Then
                'Need to allow the product number to be re-used
                If Len(vProvBTA.ProductCode) > 0 And Val(vProvBTA.ProductNumber) > 0 Then
                  vFields = New CDBFields
                  vFields.Add("product", CDBField.FieldTypes.cftCharacter, vProvBTA.ProductCode)
                  vFields.Add("product_number", CDBField.FieldTypes.cftLong, vProvBTA.ProductNumber)
                  mvEnv.Connection.InsertRecord("product_numbers", vFields, True)
                End If
              End If
            End If
            If vFound Then Exit For
          Next vNewBTA

          If Not vFound Then
            'Line has not been confirmed (it was deleted) so create a new provisional line
            If mvProvBatch Is Nothing Then
              mvProvBatch = New Batch(mvEnv)
              mvProvBatch.InitOpenBatch(Nothing, Batch.ProvisionalOrConfirmed.Provisional, Batch.BatchTypes.SaleOrReturn, vProvBatch.BankAccount, mvEnv.GetConfig("pm_sr"), False, "", 0)
            End If
            mvProvBatch.LockBatch()

            If vNewBT Is Nothing Then
              If mvProvBTColl Is Nothing Then
                mvProvBTColl = New CDBCollection
              End If
              vNewBT = New BatchTransaction(mvEnv)
              With vNewBT
                .InitFromBatch(mvEnv, mvProvBatch, mvProvBatch.AllocateTransactionNumber)
                .CloneForFA(vProvBT)
                .TransactionDate = mvBT.TransactionDate
              End With
              mvProvBTColl.Add(vNewBT)
            End If

            If mvProvBTAColl Is Nothing Then
              mvProvBTAColl = New CDBCollection
            End If
            vNewBTA = New BatchTransactionAnalysis(mvEnv)
            With vNewBTA
              .InitFromTransaction(vNewBT)
              .CloneFromBTA(vProvBTA)
              .WhenValue = mvBT.TransactionDate
            End With
            mvProvBTAColl.Add(vNewBTA)
          End If
        Next vProvBTA

        If Not (mvProvBatch Is Nothing) Then
          mvProvBatch.AddTransactionAmount((vNewBT.Amount))
          mvProvBatch.UnLockBatch()
          vCT = New ConfirmedTransaction(mvEnv)
          vCT.Create((vNewBT.BatchNumber), (vNewBT.TransactionNumber))
          mvConfirmedTrans.Add(vCT)
        End If
      End If

    End Sub

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
      mvTraderAnalysisLines = New TraderAnalysisLines
      mvPaymentPlanDetails = New TraderPaymentPlanDetails
    End Sub
    Public Sub New()
      MyBase.New()
      Class_Initialize_Renamed()
    End Sub

    Private Sub mvCCA_AuthorisingCreditCard(ByRef pMaxTime As Integer, ByRef pTime As Integer) Handles mvCCA.AuthorisingCreditCard
      RaiseEvent AuthorisingCreditCard(pMaxTime, pTime)
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Friend Sub ConfirmProvisionalTransaction(ByVal pConfirmedBTList As String, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0)
      'Update the ConfirmedTransactions records
      'For mvFinancialAdjustment = atGIKConfirmation Or atCashBatchConfirmation, pBatch/TransactionNumber = Original Batch/Transaction Numbers
      Dim vCT As ConfirmedTransaction
      Dim vConfirmList() As String
      Dim vIndex As Integer

      If Len(pConfirmedBTList) > 0 Then
        'This will be confirming an SR transaction from Trader
        'List is BatchNumber,TransactionNumber,....
        vConfirmList = Split(pConfirmedBTList, ",")
        For vIndex = 0 To UBound(vConfirmList) - 1 Step 2
          vCT = New ConfirmedTransaction(mvEnv)
          vCT.InitForUpdate(IntegerValue(vConfirmList(vIndex)), IntegerValue(vConfirmList(vIndex + 1)))
          vCT.ConfirmedBatchNumber = mvBatchNumber
          vCT.ConfirmedTransNumber = mvTransactionNumber
          mvConfirmedTrans.Add(vCT)
          mvCheckProvTrans = True
        Next
      ElseIf (mvFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Or mvFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation) Then
        vCT = New ConfirmedTransaction(mvEnv)
        vCT.InitForUpdate(pBatchNumber, pTransactionNumber)
        vCT.ConfirmedBatchNumber = mvBatchNumber
        vCT.ConfirmedTransNumber = mvTransactionNumber
        mvConfirmedTrans.Add(vCT)
        If mvFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Then mvCheckProvTrans = True
      End If

    End Sub

    Friend Sub GetCreditCardAuthorisation(ByVal pAmount As Double, Optional ByVal pIssuedSet As Boolean = False,
                                          Optional ByVal pOnlineCCAuthorisation As Boolean = False, Optional ByVal pSecurityCode As String = "",
                                          Optional ByVal pTnsSession As String = "", Optional pAuthorisationNumber As Integer = 0,
                                          Optional pParamList As ParameterList = Nothing)
      'Online Credit Card Authorisation
      'Must be called before the Transaction is started

      If (pOnlineCCAuthorisation And mvFinancialAdjustment = Batch.AdjustmentTypes.atNone) And Not (mvCardSale Is Nothing) Then
        'Don't authorise zero amount transactions
        mvCardSale.SetSecurityCode(pSecurityCode)
        If pAmount > 0 Then
          mvCCA = New CreditCardAuthorisation
          RaiseEvent StartCCAuthorisation()
          RaiseEvent CCAuthorisationProgress((ProjectText.String15676)) 'Authorising Credit Card
          mvCCA.Init(mvEnv)
          mvCCA.ContactNumber = mvBT.ContactNumber
          If Not mvCCA.AuthoriseTransaction(mvCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatNormal, pAmount, mvBT.AddressNumber, "", "", "", pAuthorisationNumber, pTnsSession, "") Then
            RaiseEvent EndCCAuthorisation()
            RaiseError(DataAccessErrors.daeCCAuthorisationFailed, (mvCCA.AuthorisationResponseMessage))
          Else
            If mvEnv.GetConfig("fp_cc_authorisation_type") = "SAGEPAYHOSTED" AndAlso pParamList IsNot Nothing Then
              mvCCA.StoreCardToken(pParamList)
            End If
          End If
          RaiseEvent EndCCAuthorisation()
        ElseIf pAmount = 0 Then
          'The amount of the transaction is zero
          'Could be that no stock was issued rather than all items were free so check here
          If pIssuedSet AndAlso mvEnv.GetConfig("fp_cc_authorisation_type") <> "SCXLVPCSCP" Then  'Nominal Amount functionality does not apply to SecureCXL
            'This is where we would do the notional amount processing
            pAmount = Val(mvEnv.GetConfig("fp_cc_authorise_nominal_amount"))
            If pAmount > 0 Then
              mvCCA = New CreditCardAuthorisation
              RaiseEvent StartCCAuthorisation()
              RaiseEvent CCAuthorisationProgress((ProjectText.String15676)) 'Authorising Credit Card
              mvCCA.Init(mvEnv)
              mvCCA.ContactNumber = mvBT.ContactNumber
              If Not mvCCA.AuthoriseTransaction(mvCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatNotional, pAmount, mvBT.AddressNumber) Then
                RaiseEvent EndCCAuthorisation()
                RaiseError(DataAccessErrors.daeCCAuthorisationFailed, (mvCCA.AuthorisationResponseMessage))
              End If
              RaiseEvent EndCCAuthorisation()
            End If
            'Set no claim required so the back order allocation does the authorisation
            mvCardSale.NoClaimRequired = True
          Else
            'We have got the card sale record but we have not taken any payment. Save the card details as Template Number if the SecureCXL is in use.
            mvCardSale.SetTemplateNumber()
          End If
        End If
        'UPGRADE_NOTE: Object mvCCA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mvCCA = Nothing
      End If

    End Sub

    Friend Sub InitCash(ByVal pPaymentMethod As String, ByVal pBankAccount As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pTransactionDate As String, ByVal pReceipt As String, ByRef pEligibleGiftAid As Boolean, ByVal pExisting As Boolean, ByVal pProvisional As Boolean, Optional ByVal pCurrencyCode As String = "", Optional ByVal pExchangeRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pMailingContactNumber As String = "", Optional ByVal pMailingAddressNumber As String = "", Optional ByVal pBatchCategory As String = "", Optional ByVal pReference As String = "", Optional ByVal pNotes As String = "", Optional ByVal pBankDetailsNumber As Integer = 0, Optional ByVal pAddRef1 As String = "", Optional ByVal pAddRef2 As String = "", Optional ByVal pAddRef1Caption As String = "", Optional ByVal pAddRef2Caption As String = "", Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pAddressTo As String = "", Optional ByVal pSalesLedgerAcc As String = "", Optional ByVal pStockSale As Boolean = False, Optional ByVal pPayMethodCode As String = "")
      'Cash/Cheque/PostalOrder/GiftInKind/Voucher/SaleOrReturn -- No Financial Adjustments
      'pPaymentMethod = Trader mvPayMethod1
      Dim vCT As ConfirmedTransaction
      Dim vProvisional As Batch.ProvisionalOrConfirmed
      Dim vBatchType As Batch.BatchTypes
      Dim vNotes As String = ""
      Dim vPaymentMethod As String

      mvBatchNumber = pBatchNumber
      mvTransactionNumber = pTransactionNumber
      mvExisting = pExisting
      mvFinancialAdjustment = Batch.AdjustmentTypes.atNone
      If Len(pCurrencyCode) > 0 Then
        mvCurrencyCode = pCurrencyCode
        mvExchangeRate = Val(pExchangeRate)
      End If
      vProvisional = CType(IIf(pProvisional = True, Batch.ProvisionalOrConfirmed.Provisional, Batch.ProvisionalOrConfirmed.Confirmed), Access.Batch.ProvisionalOrConfirmed)

      If vProvisional = Batch.ProvisionalOrConfirmed.Provisional Then
        vNotes = BuildNotes(pAddRef1, pAddRef2, pAddRef1Caption, pAddRef2Caption)
      End If
      If Len(pNotes) > 0 Then
        If Len(vNotes) > 0 Then vNotes = vNotes & vbLf
        vNotes = vNotes & pNotes
      End If

      Select Case pPaymentMethod
        Case "GFIK"
          vPaymentMethod = mvEnv.GetConfig("pm_gift_in_kind")
          vBatchType = Batch.BatchTypes.GiftInKind
        Case "SAOR"
          vPaymentMethod = mvEnv.GetConfig("pm_sr")
          vBatchType = Batch.BatchTypes.SaleOrReturn
        Case "VOUC"
          vPaymentMethod = mvEnv.GetConfig("pm_voucher")
          vBatchType = Batch.BatchTypes.CAFVouchers
        Case "CHEQ"
          vPaymentMethod = mvEnv.GetConfig("pm_cheque")
          vBatchType = Batch.BatchTypes.Cash
        Case "CQIN"
          vPaymentMethod = mvEnv.GetConfig("pm_cheque")
          vBatchType = Batch.BatchTypes.CashWithInvoice
        Case "POST"
          vPaymentMethod = mvEnv.GetConfig("pm_po")
          vBatchType = Batch.BatchTypes.Cash
        Case Else '"CASH"
          If pPaymentMethod = mvEnv.GetConfig("pm_sp") Then
            'Bank Statement Posting
            vPaymentMethod = pPaymentMethod
            vBatchType = Batch.BatchTypes.BankStatement
          ElseIf pPaymentMethod = mvEnv.GetConfig("pm_so") Then
            'Credit List Reconciliation
            vPaymentMethod = pPaymentMethod
            vBatchType = Batch.BatchTypes.StandingOrder
          Else
            'Cash
            vPaymentMethod = mvEnv.GetConfig("pm_cash")
            vBatchType = Batch.BatchTypes.Cash
          End If
      End Select
      If Not String.IsNullOrEmpty(pPayMethodCode) Then
        'Override the Payment Method with pPayMethodCode (in credit list reconciliation this value is the payment method in bank transactions record or the pm_credit_list_reconciliation config value) 
        vPaymentMethod = pPayMethodCode
      End If

      SetUpBatchAndTransaction(vBatchType, pBankAccount, pPaymentMethod, "", pBatchCategory, vProvisional)

      AddTransaction(pContactNumber, pAddressNumber, pTransactionDate, vPaymentMethod, pReceipt, pEligibleGiftAid, pMailing, pMailingContactNumber, pMailingAddressNumber, pReference, vNotes, pBankDetailsNumber)

      If pPaymentMethod = "CQIN" Then
        mvCreditSale = New CreditSale(mvEnv)
        If mvExisting Then mvCreditSale.Init(mvBatchNumber, mvTransactionNumber)
        If mvCreditSale.Existing = False Then mvCreditSale.Create(mvEnv, mvBatchNumber, mvTransactionNumber)
        mvCreditSale.Update(pContactNumber, pAddressNumber, pAddressTo, pSalesLedgerAcc, pStockSale)
      End If

      If mvExisting = True And vProvisional = Batch.ProvisionalOrConfirmed.Provisional Then
        vCT = New ConfirmedTransaction(mvEnv)
        vCT.Init(mvBatchNumber, mvTransactionNumber)
        vCT.AdditionalReference1 = pAddRef1
        vCT.AdditionalReference2 = pAddRef2
        mvConfirmedTrans.Add(vCT)
      ElseIf mvBatch.Provisional Then
        vCT = New ConfirmedTransaction(mvEnv)
        vCT.Create(mvBatchNumber, mvTransactionNumber, pAddRef1, pAddRef2)
        mvConfirmedTrans.Add(vCT)
      End If

    End Sub

    Friend Sub InitCardSale(ByVal pPayMethod As String, ByVal pBankAcc As String, ByVal pContact As Integer, ByVal pAddress As Integer, ByVal pTrnsDate As String, ByVal pRcpt As String, ByVal pGA As Boolean, ByVal pExist As Boolean, ByVal pCardType As String, ByVal pCardNo As String, ByVal pIssueNo As String, ByVal pValidDate As String, ByVal pExpiry As String, ByVal pAuthCode As String, ByVal pProv As Boolean, Optional ByVal pCurr As String = "", Optional ByVal pExRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pMailContact As String = "", Optional ByVal pMailAddress As String = "", Optional ByVal pBatchCat As String = "", Optional ByVal pRef As String = "", Optional ByVal pNotes As String = "", Optional ByVal pRef1 As String = "", Optional ByVal pRef2 As String = "", Optional ByVal pRef1Cap As String = "", Optional ByVal pRef2Cap As String = "", Optional ByVal pBatchDate As String = "", Optional ByVal pBatchNo As Integer = 0, Optional ByVal pTransNo As Integer = 0, Optional ByVal pAddressTo As String = "", Optional ByVal pSalesLedgerAcc As String = "", Optional ByVal pStockSale As Boolean = False)
      'CreditCard/DebitCard/CAFCard -- No Financial Adjustments
      'pPayMethod = Trader mvPayMethod1
      'pRcpt = Receipt
      'pGA = Eligible For Gift Aid
      'pCurr = Currency Code
      Dim vCT As ConfirmedTransaction
      Dim vBatchType As Batch.BatchTypes
      Dim vNotes As String = ""
      Dim vPayMethod As String
      Dim vProvisional As Batch.ProvisionalOrConfirmed

      mvBatchNumber = pBatchNo
      mvTransactionNumber = pTransNo
      mvExisting = pExist
      mvFinancialAdjustment = Batch.AdjustmentTypes.atNone
      If Len(pCurr) > 0 Then
        mvCurrencyCode = pCurr
        mvExchangeRate = Val(pExRate)
      End If
      vProvisional = CType(IIf(pProv = True, Batch.ProvisionalOrConfirmed.Provisional, Batch.ProvisionalOrConfirmed.Confirmed), Access.Batch.ProvisionalOrConfirmed)

      If vProvisional = Batch.ProvisionalOrConfirmed.Provisional Then
        vNotes = BuildNotes(pRef1, pRef2, pRef1Cap, pRef2Cap)
      End If
      If Len(pNotes) > 0 Then
        If Len(vNotes) > 0 Then vNotes = vNotes & vbLf
        vNotes = vNotes & pNotes
      End If

      Select Case pPayMethod
        Case "CAFC"
          vPayMethod = mvEnv.GetConfig("pm_caf_card")
          vBatchType = Batch.BatchTypes.CAFCards
        Case "CCIN"
          vPayMethod = mvEnv.GetConfig("pm_cc")
          vBatchType = Access.Batch.BatchTypes.CreditCardWithInvoice
        Case Else
          If pPayMethod = "CCARD" Then
            vBatchType = Batch.BatchTypes.CreditCard
            vPayMethod = mvEnv.GetConfig("pm_cc")
          Else '"DCARD"
            mvDebitCard = True
            vBatchType = Batch.BatchTypes.DebitCard
            vPayMethod = mvEnv.GetConfig("pm_dc")
          End If
      End Select

      SetUpBatchAndTransaction(vBatchType, pBankAcc, pPayMethod, pBatchDate, pBatchCat, vProvisional)

      AddTransaction(pContact, pAddress, pTrnsDate, vPayMethod, pRcpt, pGA, pMailing, pMailContact, pMailAddress, pRef, vNotes)

      If Not ContainsNumbers(pValidDate) Then
        pValidDate = ""
      Else
        pValidDate = Replace(pValidDate, "/", "")
      End If
      pExpiry = Replace(pExpiry, "/", "")

      If pPayMethod = "CCIN" Then
        mvCreditSale = New CreditSale(mvEnv)
        If mvExisting Then mvCreditSale.Init(mvBatchNumber, mvTransactionNumber)
        If mvCreditSale.Existing = False Then mvCreditSale.Create(mvEnv, mvBatchNumber, mvTransactionNumber)
        mvCreditSale.Update(pContact, pAddress, pAddressTo, pSalesLedgerAcc, pStockSale)
      End If

      mvCardSale = New CardSale(mvEnv)
      If mvExisting Then mvCardSale.Init(mvBatchNumber, mvTransactionNumber)
      If mvCardSale.Existing = False Then mvCardSale.Create(mvEnv, mvBatchNumber, mvTransactionNumber)
      mvCardSale.Update(pCardNo, pIssueNo, pValidDate, pExpiry, pAuthCode, pCardType, False)

      If mvExisting = True And vProvisional = Batch.ProvisionalOrConfirmed.Provisional Then
        vCT = New ConfirmedTransaction(mvEnv)
        vCT.Init(mvBatchNumber, mvTransactionNumber)
        vCT.AdditionalReference1 = pRef1
        vCT.AdditionalReference2 = pRef2
        mvConfirmedTrans.Add(vCT)
      ElseIf mvBatch.Provisional Then
        vCT = New ConfirmedTransaction(mvEnv)
        vCT.Create(mvBatchNumber, mvTransactionNumber, pRef1, pRef2)
        mvConfirmedTrans.Add(vCT)
      End If

    End Sub

    Friend Sub InitCreditSale(ByVal pPayMethod As String, ByVal pBankAccount As String, ByVal pContactNo As Integer, ByVal pAddressNo As Integer, ByVal pTransDate As String, ByVal pReceipt As String, ByVal pEligGiftAid As Boolean, ByVal pExisting As Boolean, ByVal pAddressTo As String, ByVal pSalesLedgerAcc As String, ByVal pCompany As String, ByRef pUseSalesLedger As Boolean, ByVal pStockSale As Boolean, ByVal pServBookCredits As Boolean, Optional ByVal pCurrCode As String = "", Optional ByVal pExRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pMailContactNo As String = "", Optional ByVal pMailAddressNo As String = "", Optional ByVal pBatchCat As String = "", Optional ByVal pRef As String = "", Optional ByVal pNotes As String = "", Optional ByVal pFA As Batch.AdjustmentTypes = Batch.AdjustmentTypes.atNone, Optional ByVal pFATransType As String = "", Optional ByVal pBatchDate As String = "", Optional ByVal pBatchNo As Integer = 0, Optional ByVal pTransNo As Integer = 0)
      'CreditSale
      'pPayMethod = Trader mvPayMethod1

      'BR15226: Make sure any changes in this method are reflected in InitCash and InitCardSale methods for CQIN and CCIN

      mvBatchNumber = pBatchNo
      mvTransactionNumber = pTransNo
      mvExisting = pExisting
      mvFinancialAdjustment = pFA
      If Len(pCurrCode) > 0 Then
        mvCurrencyCode = pCurrCode
        mvExchangeRate = Val(pExRate)
      End If

      If pPayMethod = "CRDN" Then mvCSCreditNote = True
      mvServiceBookingCredits = pServBookCredits
      mvCSCompany = pCompany
      mvCSStockSale = pStockSale
      mvUseSalesLedger = pUseSalesLedger

      SetUpBatchAndTransaction(Batch.BatchTypes.CreditSales, pBankAccount, pPayMethod, pBatchDate, pBatchCat)

      AddTransaction(pContactNo, pAddressNo, pTransDate, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod), pReceipt, pEligGiftAid, pMailing, pMailContactNo, pMailAddressNo, pRef, pNotes, 0, pFATransType)

      mvCreditSale = New CreditSale(mvEnv)
      If mvExisting Then mvCreditSale.Init(mvBatchNumber, mvTransactionNumber)
      If mvCreditSale.Existing = False Then mvCreditSale.Create(mvEnv, mvBatchNumber, mvTransactionNumber)
      mvCreditSale.Update(pContactNo, pAddressNo, pAddressTo, pSalesLedgerAcc, mvCSStockSale)

      '  This is no longer done here but code left in case it is needed again (SAS 3/10/2006)
      '  'Get the Invoice Number
      '  If pPrintInv Then
      '    If mvExisting Then
      '      Set vRS = mvEnv.Connection.GetRecordSet("SELECT invoice_number FROM invoices WHERE batch_number = " & mvBatchNumber & " AND transaction_number = " & mvTransactionNumber)
      '      If vRS.Fetch = rssOK Then
      '        mvInvoiceNumber = vRS.Fields(1).IntegerValue
      '        If mvInvoiceNumber = 0 Then mvInvoiceNumber = mvEnv.GetControlNumber("I")
      '      End If
      '      vRS.CloseRecordSet
      '      If mvInvoiceNumber = 0 Then RaiseError daeInvoiceNotFound, "for this transaction"
      '    Else
      '      mvInvoiceNumber = mvEnv.GetControlNumber("I")
      '    End If
      '  End If

    End Sub

    Friend Sub InitFACash(ByVal pPaymentMethod As String, ByVal pBankAccount As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pBatchDate As String, ByVal pTransDate As String, ByVal pReceipt As String, ByRef pEligibleGiftAid As Boolean, ByVal pExisting As Boolean, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pCurrencyCode As String = "", Optional ByVal pExchangeRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pMailingContactNumber As String = "", Optional ByVal pMailingAddressNumber As String = "", Optional ByVal pBatchCategory As String = "", Optional ByVal pReference As String = "", Optional ByVal pNotes As String = "", Optional ByVal pAdjustTransType As String = "", Optional ByVal pOrigPayMethod As String = "", Optional ByVal pAddRef1 As String = "", Optional ByVal pAddRef2 As String = "", Optional ByVal pBankDetailsNumber As Integer = 0, Optional ByVal pBatchNo As Integer = 0, Optional ByVal pTransNo As Integer = 0, Optional ByVal pAppType As TraderApplication.ApplicationType = 0, Optional ByVal pPostToCashBook As String = "")
      'Financial Adjustment Cash transaction
      'Cash/Cheque/PostalOrder/GiftInKind/Voucher/SaleOrReturn -- Financial Adjustments only
      'pPaymentMethod = Trader mvPayMethod1
      Dim vBatchType As Batch.BatchTypes
      Dim vPayMethod As String

      mvBatchNumber = pBatchNo
      mvTransactionNumber = pTransNo
      mvExisting = pExisting
      mvFinancialAdjustment = pFinancialAdjustment
      If Len(pCurrencyCode) > 0 Then
        mvCurrencyCode = pCurrencyCode
        mvExchangeRate = Val(pExchangeRate)
      End If

      Select Case pPostToCashBook
        Case "Y"
          mvPostCashBook = False
          mvPostCashBookSet = True
        Case "N"
          mvPostCashBook = True
          mvPostCashBookSet = True
      End Select

      Select Case pPaymentMethod
        Case "GFIK"
          vPayMethod = mvEnv.GetConfig("pm_gift_in_kind")
          vBatchType = Batch.BatchTypes.GiftInKind
        Case "SAOR"
          vPayMethod = mvEnv.GetConfig("pm_sr")
          vBatchType = Batch.BatchTypes.SaleOrReturn
        Case "VOUC"
          vPayMethod = mvEnv.GetConfig("pm_voucher")
          vBatchType = Batch.BatchTypes.FinancialAdjustment
        Case Else
          If pPaymentMethod = "CHEQ" Then
            vPayMethod = mvEnv.GetConfig("pm_cheque")
          ElseIf pPaymentMethod = "POST" Then
            vPayMethod = mvEnv.GetConfig("pm_po")
          ElseIf pPaymentMethod = "SO" And pAppType <> TraderApplication.ApplicationType.atCreditListReconciliation Then
            vPayMethod = mvEnv.GetConfig("pm_so")
          Else '"CASH"
            vPayMethod = mvEnv.GetConfig("pm_cash")
          End If
          Select Case mvFinancialAdjustment
            Case Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              vBatchType = Batch.BatchTypes.Cash
            Case Batch.AdjustmentTypes.atAdjustment, Batch.AdjustmentTypes.atMove
              vBatchType = Batch.BatchTypes.FinancialAdjustment
              pPaymentMethod = "" 'Do not want to set pay method on batch header.
              If Len(pOrigPayMethod) > 0 Then vPayMethod = pOrigPayMethod
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
      End Select

      SetUpBatchAndTransaction(vBatchType, pBankAccount, pPaymentMethod, pBatchDate, pBatchCategory)

      AddTransaction(pContactNumber, pAddressNumber, pTransDate, vPayMethod, pReceipt, pEligibleGiftAid, pMailing, pMailingContactNumber, pMailingAddressNumber, pReference, pNotes, pBankDetailsNumber, pAdjustTransType)

    End Sub

    Friend Sub InitFACardSale(ByVal pPaymentMethod As String, ByVal pBankAccount As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pTransDate As String, ByVal pReceipt As String, ByVal pEligibleGiftAid As Boolean, ByVal pExisting As Boolean, ByVal pCardType As String, ByVal pCardNumber As String, ByVal pIssueNumber As String, ByVal pValidDate As String, ByVal pExpiry As String, ByVal pAuthorisationCode As String, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pCurrencyCode As String = "", Optional ByVal pExchangeRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pMailContactNumber As String = "", Optional ByVal pMailAddressNumber As String = "", Optional ByVal pBatchCat As String = "", Optional ByVal pRef As String = "", Optional ByVal pNotes As String = "", Optional ByVal pFATransType As String = "", Optional ByVal pBatchDate As String = "", Optional ByVal pBatchNo As Integer = 0, Optional ByVal pTransNo As Integer = 0)
      'CreditCard/DebitCard/CAFCard -- Financial Adjustment only
      'pPayMethod = Trader mvPayMethod1
      Dim vBatchType As Batch.BatchTypes
      Dim vPayMethod As String

      mvBatchNumber = pBatchNo
      mvTransactionNumber = pTransNo
      mvExisting = pExisting
      mvFinancialAdjustment = pFinancialAdjustment
      If Len(pCurrencyCode) > 0 Then
        mvCurrencyCode = pCurrencyCode
        mvExchangeRate = Val(pExchangeRate)
      End If

      Select Case pPaymentMethod
        Case "CAFC"
          vPayMethod = mvEnv.GetConfig("pm_caf_card")
          vBatchType = Batch.BatchTypes.CAFCards
        Case Else
          Select Case mvFinancialAdjustment
            Case Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              If pPaymentMethod = "CCARD" Then
                vBatchType = Batch.BatchTypes.CreditCard
              Else
                vBatchType = Batch.BatchTypes.DebitCard
              End If
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
          If pPaymentMethod = "CCARD" Then
            vPayMethod = mvEnv.GetConfig("pm_cc")
          Else '"DCARD"
            mvDebitCard = True
            vPayMethod = mvEnv.GetConfig("pm_dc")
          End If
      End Select

      SetUpBatchAndTransaction(vBatchType, pBankAccount, pPaymentMethod, pBatchDate, pBatchCat)

      AddTransaction(pContactNumber, pAddressNumber, pTransDate, vPayMethod, pReceipt, pEligibleGiftAid, pMailing, pMailContactNumber, pMailAddressNumber, pRef, pNotes, 0, pFATransType)

    End Sub

    Public Sub InitPreTaxPayrollGiving(ByVal pEnv As CDBEnvironment, ByVal pPaymentMethod As String, ByVal pBankAccount As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pTransactionDate As String, ByVal pReference As String, ByVal pPostCashBook As Boolean, ByVal pExisting As Boolean, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pCurrencyCode As String = "", Optional ByVal pExchangeRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pFATransType As String = "", Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0)
      'Pre Tax Payroll Giving payments
      'pPaymentMethod = pm_... config

      mvEnv = pEnv
      mvBatchNumber = pBatchNumber
      mvTransactionNumber = pTransactionNumber
      mvExisting = pExisting
      mvFinancialAdjustment = pFinancialAdjustment
      If Len(pCurrencyCode) > 0 Then
        mvCurrencyCode = pCurrencyCode
        mvExchangeRate = Val(pExchangeRate)
      End If
      mvPostCashBook = pPostCashBook
      mvPostCashBookSet = True

      SetUpBatchAndTransaction(Batch.BatchTypes.GiveAsYouEarn, pBankAccount, pPaymentMethod, "", "", Batch.ProvisionalOrConfirmed.Confirmed, pReference)
      If mvTransactionNumber = 1 Then 'Make sure we only change the following flags for new batch
        mvBatch.ReadyForBanking = True
        mvBatch.SetPayingInSlipPrinted(0)
        mvBatch.Save()
      End If
      AddTransaction(pContactNumber, pAddressNumber, pTransactionDate, pPaymentMethod, "N", False, pMailing, CStr(pContactNumber), CStr(pAddressNumber), pReference, "", 0, pFATransType)
    End Sub

    Public Sub InitPostTaxPayrollGiving(ByVal pEnv As CDBEnvironment, ByVal pPaymentMethod As String, ByVal pBankAccount As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pTransactionDate As String, ByVal pReference As String, ByVal pPostCashBook As Boolean, ByVal pExisting As Boolean, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, Optional ByVal pCurrencyCode As String = "", Optional ByVal pExchangeRate As String = "", Optional ByVal pMailing As String = "", Optional ByVal pFATransType As String = "", Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0)
      'Post Tax Payroll Giving payments
      'pPaymentMethod = pm_... config

      mvEnv = pEnv
      mvBatchNumber = pBatchNumber
      mvTransactionNumber = pTransactionNumber
      mvExisting = pExisting
      mvFinancialAdjustment = pFinancialAdjustment
      If Len(pCurrencyCode) > 0 Then
        mvCurrencyCode = pCurrencyCode
        mvExchangeRate = Val(pExchangeRate)
      End If
      mvPostCashBook = pPostCashBook
      mvPostCashBookSet = True

      SetUpBatchAndTransaction(Batch.BatchTypes.PostTaxPayrollGiving, pBankAccount, pPaymentMethod, "", "", Batch.ProvisionalOrConfirmed.Confirmed, pReference)
      AddTransaction(pContactNumber, pAddressNumber, pTransactionDate, pPaymentMethod, "N", True, pMailing, CStr(pContactNumber), CStr(pAddressNumber), pReference, "", 0, pFATransType)
    End Sub

    Public Sub SaveTransaction(ByVal pTransAmount As Double, Optional ByVal pOrigTransAmount As Double = 0, Optional ByVal pOrigTransCurrencyAmount As Double = 0, Optional ByVal pCSTermsFrom As String = "", Optional ByVal pCSTermsPeriod As String = "", Optional ByVal pCSTermsNumber As String = "", Optional ByVal pDepositAmount As Double = 0, Optional ByVal pEventNumber As Integer = 0)
      'pTransAmount is amount entered by user for the Transaction Amount and may be different to sum of BTA lines
      'Save the Batch and BatchTransaction etc.
      Dim vBT As BatchTransaction
      Dim vBTA As BatchTransactionAnalysis
      Dim vCT As ConfirmedTransaction
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vInvoice As Invoice
      Dim vAmount As Double
      Dim vCurrAmount As Double
      Dim vInvDueDate As Date
      Dim vInvStatus As String = ""
      Dim vIssued As Integer
      Dim vPPNumbers As String = ""
      Dim vTrans As Boolean
      Dim vZeroEBCount As Integer
      Dim vOnlyZeroEventBookings As Boolean
      Dim vBTAmount As Double
      Dim vBTCurrencyAmount As Double

      'Transaction should have already been started, but if not then start it now
      If Not (mvEnv.Connection.InTransaction) Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      '-------------------------------------------------------------------------
      'Save Batch Transaction
      '-------------------------------------------------------------------------
      mvBT.CurrencyAmount = pTransAmount 'Set this to the amount entered by the user
      mvBT.Amount = CalculateCurrencyAmount(mvBT.CurrencyAmount, mvCurrencyCode, mvExchangeRate, True)
      mvBT.Save()

      '-------------------------------------------------------------------------
      'Update the Batch information
      '-------------------------------------------------------------------------
      If mvCSCreditNote Or mvServiceBookingCredits Then
        vBTAmount = mvBT.Amount * -1
        vBTCurrencyAmount = mvBT.CurrencyAmount * -1
      Else
        vBTAmount = mvBT.Amount
        vBTCurrencyAmount = mvBT.CurrencyAmount
      End If
      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, mvBatch.BatchNumber)
      If mvExisting Then
        If mvCSCreditNote Or mvServiceBookingCredits Then
          pOrigTransAmount = pOrigTransAmount * -1
          pOrigTransCurrencyAmount = pOrigTransCurrencyAmount * -1
        End If
        If Len(mvCurrencyCode) > 0 Then vCurrAmount = FixTwoPlaces(vBTCurrencyAmount - pOrigTransCurrencyAmount)
        vAmount = FixTwoPlaces(vBTAmount - pOrigTransAmount)
      Else
        vAmount = vBTAmount
        vCurrAmount = vBTCurrencyAmount
        vUpdateFields.Add("number_of_transactions", CDBField.FieldTypes.cftLong, "number_of_transactions + 1")
        If mvBatch.BatchType = Batch.BatchTypes.BankStatement Or mvBatch.BatchType = Batch.BatchTypes.StandingOrder Then
          'Bank Statement Posting  or  Credit List Reconciliation
          vUpdateFields.Add("paying_in_slip_printed", CDBField.FieldTypes.cftCharacter, "Y")
          If Not mvEnv.GetConfigOption("option_post_batches_to_CB") Then vUpdateFields.Add("posted_to_cash_book", CDBField.FieldTypes.cftCharacter, "Y")
        End If
      End If
      vUpdateFields.Add("detail_completed", CDBField.FieldTypes.cftCharacter, "N")
      vUpdateFields.Add("transaction_total", CDBField.FieldTypes.cftLong, "transaction_total + " & vAmount)
      vUpdateFields.Add("contents_amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
      vUpdateFields.Add("contents_amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      If Len(mvCurrencyCode) > 0 Then vUpdateFields.Add("currency_transaction_total", CDBField.FieldTypes.cftLong, "currency_transaction_total + " & vCurrAmount)
      mvEnv.Connection.UpdateRecords("batches", vUpdateFields, vWhereFields)

      '-------------------------------------------------------------------------
      'Add the Card Sales record if required
      '-------------------------------------------------------------------------
      If Not (mvCardSale Is Nothing) Then
        mvCardSale.Save()
        '-------------------------------------------------------------------------
        'If this is a debit card transaction and any payment plans were created by it
        'they may have the wrong payment method on them due to having been created before
        'we knew if it was a debit or credit card.
        '-------------------------------------------------------------------------
        If mvDebitCard Then
          For Each vBTA In mvBT.Analysis
            If vBTA.LineType Like "[CMO]" Then
              If Len(vPPNumbers) > 0 Then vPPNumbers = vPPNumbers & ","
              vPPNumbers = vPPNumbers & vBTA.PaymentPlanNumber
            End If
          Next vBTA

          If Len(vPPNumbers) > 0 Then
            vWhereFields = New CDBFields
            vUpdateFields = New CDBFields
            vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, vPPNumbers, CDBField.FieldWhereOperators.fwoInOrEqual)
            vWhereFields.Add("payment_method", CDBField.FieldTypes.cftCharacter, mvEnv.GetConfig("pm_cc"))
            vUpdateFields.Add("payment_method", CDBField.FieldTypes.cftCharacter, mvEnv.GetConfig("pm_dc"))
            mvEnv.Connection.UpdateRecords("orders", vUpdateFields, vWhereFields, False)
          End If
        End If
      End If

      '-------------------------------------------------------------------------
      'Add the Credit Sales record if required
      '-------------------------------------------------------------------------
      Dim vEvent As New CDBEvent(mvEnv)
      Dim vEventInvoice As Boolean
      If Not (mvCreditSale Is Nothing) Then
        mvCreditSale.Save()
        If mvBatch.BatchType = Access.Batch.BatchTypes.CreditSales Then
          If mvUseSalesLedger = True And mvCSStockSale = False Then
            mvBT.InitAnalysisAdditionalData()
            vZeroEBCount = 0
            For Each vBTA In mvBT.Analysis
              If mvServiceBookingCredits Then
                If vBTA.Quantity > 0 Then vIssued = vIssued + vBTA.Quantity
              Else
                If vBTA.Issued > 0 Then vIssued = vIssued + vBTA.Issued
              End If
              If vBTA.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBooking And vBTA.Amount = 0 Then
                vZeroEBCount = vZeroEBCount + 1
              End If
            Next vBTA
            If vZeroEBCount = mvBT.Analysis.Count() Then
              vOnlyZeroEventBookings = True
            Else
              vOnlyZeroEventBookings = False
            End If
            If vIssued > 0 Then
              'Create Invoice
              vInvoice = New Invoice
              vInvoice.Init(mvEnv)
              vInvoice.CalcInvPayDue(pCSTermsFrom, pCSTermsPeriod, IntegerValue(pCSTermsNumber), mvBatchNumber, mvTransactionNumber, CDate(mvBT.TransactionDate), vInvDueDate)
              vInvoice.GetInvPayStatus(CType(IIf(vOnlyZeroEventBookings = True, Invoice.InvPayStatuses.InvFullyPaid, Invoice.InvPayStatuses.InvNotPaid), Invoice.InvPayStatuses), vInvStatus)
              If mvExisting Then vInvoice.Init(mvEnv, mvBatchNumber, mvTransactionNumber)
              If vInvoice.Existing = False Then vInvoice.Create(mvEnv, mvBatchNumber, mvTransactionNumber)
              If Len(mvEnv.GetConfig("invoice_date_from_event_start")) > 0 Then
                If CDbl(mvEnv.GetConfig("invoice_date_from_event_start")) > 0 Then
                  'event start date
                  vEvent.Init(pEventNumber)
                  vEventInvoice = True
                End If
              End If
              If vEventInvoice = True Then
                vInvoice.Update(0, mvCreditSale.ContactNumber, mvCreditSale.AddressNumber, mvCSCompany, mvCreditSale.SalesLedgerAccount, 0, -1, vEvent.StartDate, CStr(vInvDueDate), vInvStatus, If((mvCSCreditNote Or mvServiceBookingCredits), "N", "I"), pDepositAmount)
              Else
                vInvoice.Update(0, mvCreditSale.ContactNumber, mvCreditSale.AddressNumber, mvCSCompany, mvCreditSale.SalesLedgerAccount, 0, -1, mvBT.TransactionDate, CStr(vInvDueDate), vInvStatus, If((mvCSCreditNote Or mvServiceBookingCredits), "N", "I"), pDepositAmount)
              End If
              vInvoice.Save()
              mvCSInvoiceCreated = True
            End If
          End If
        End If
      End If

      '-------------------------------------------------------------------------
      'Save any Provisional / Confirmed Transactions
      '-------------------------------------------------------------------------
      If mvConfirmedTrans.Count > 0 Then
        If mvCheckProvTrans Then
          For Each vCT In mvConfirmedTrans
            'This may need to create new provisional Batches
            CreateNewProvisionalBatches(vCT.ProvisionalBatchNumber, vCT.ProvisionalTransNumber)
          Next vCT

          If Not (mvProvBatch Is Nothing) Then
            For Each vBTA In mvProvBTAColl
              vBTA.Save()
            Next vBTA
            For Each vBT In mvProvBTColl
              vBT.Save()
            Next vBT
            mvProvBatch.Save()
          End If
        End If

        For Each vCT In mvConfirmedTrans
          vCT.Save()
        Next vCT
      End If

      If vTrans Then mvEnv.Connection.CommitTransaction()

    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pResetAnalysisLines As Boolean = True, Optional ByVal pBatchAnalysisCode As String = "")
      mvEnv = pEnv
      If pResetAnalysisLines Then mvTraderAnalysisLines = New TraderAnalysisLines
      If Len(pBatchAnalysisCode) > 0 Then mvBatchAnalysisCode = pBatchAnalysisCode
    End Sub

    Public Function GetTraderAnalysisLine(ByVal pKey As Integer, Optional ByVal pTransactionType As String = "") As TraderAnalysisLine
      Dim vTDRLine As TraderAnalysisLine

      If mvTraderAnalysisLines.Exists(CStr(pKey)) Then
        'Already exists, so retrieve line
        vTDRLine = mvTraderAnalysisLines(CStr(pKey))
      Else
        'Does not exist, so create a new line and add it to the collection
        vTDRLine = mvTraderAnalysisLines.Add(CStr(pKey))
        vTDRLine.Init(pKey, pTransactionType)
      End If
      GetTraderAnalysisLine = vTDRLine

    End Function

    Public Sub DeleteTraderAnalysisLine(ByVal pKey As Integer)
      Dim vCol As TraderAnalysisLines
      Dim vLine As TraderAnalysisLine
      Dim vIndex As Integer

      'Delete this Key
      If mvTraderAnalysisLines.Exists(CStr(pKey)) Then
        mvTraderAnalysisLines.Remove(CStr(pKey))

        vCol = mvTraderAnalysisLines
        mvTraderAnalysisLines = New TraderAnalysisLines

        'Now, re-number the collection
        vIndex = 1
        For Each vLine In vCol
          mvTraderAnalysisLines.AddItem(vLine, CStr(vIndex))
          vIndex = vIndex + 1
        Next vLine
      End If

    End Sub

    Public Function GetPurchaseOrderDetail(ByVal pKey As Integer, Optional ByRef pPurchaseOrderNumber As Integer = 0) As PurchaseOrderDetail
      Dim vPOD As PurchaseOrderDetail

      If PODExists(pKey) Then
        vPOD = CType(PurchaseOrderDetails.Item(CStr(pKey)), PurchaseOrderDetail)
      Else
        vPOD = New PurchaseOrderDetail
        vPOD.Init(mvEnv, pPurchaseOrderNumber, pKey)
        PurchaseOrderDetails.Add(vPOD, CStr(pKey))
      End If

      GetPurchaseOrderDetail = vPOD

    End Function

    Public Function PODExists(ByVal pKey As Integer) As Boolean
      Dim vPOD As PurchaseOrderDetail
      Dim vFound As Boolean

      If PurchaseOrderDetails.Count() > 0 Then
        For Each vPOD In PurchaseOrderDetails
          If vPOD.LineNumber = pKey Then
            vFound = True
            Exit For
          End If
        Next vPOD
      End If
      PODExists = vFound
    End Function

    Public Function GetPurchaseInvoiceDetail(ByVal pKey As Integer, Optional ByRef pPurchaseInvoiceNumber As Integer = 0, Optional ByRef pPurchaseOrderNumber As Integer = 0) As PurchaseInvoiceDetail
      Dim vPID As PurchaseInvoiceDetail
      Dim vPOD As PurchaseOrderDetail

      If PIDExists(pKey) Then
        vPID = CType(PurchaseInvoiceDetails.Item(CStr(pKey)), PurchaseInvoiceDetail)
      Else
        vPID = New PurchaseInvoiceDetail
        If pPurchaseOrderNumber > 0 And pPurchaseInvoiceNumber = 0 Then
          vPOD = New PurchaseOrderDetail
          vPOD.Init(mvEnv, pPurchaseOrderNumber, pKey)
          vPID.Create(mvEnv, vPOD.GetDataAsParameters)
        Else
          vPID.Init(mvEnv, pPurchaseInvoiceNumber, pKey)
        End If
        PurchaseInvoiceDetails.Add(vPID, CStr(pKey))
      End If

      GetPurchaseInvoiceDetail = vPID

    End Function

    Public Function PIDExists(ByVal pKey As Integer) As Boolean
      Dim vPID As PurchaseInvoiceDetail
      Dim vFound As Boolean

      If PurchaseInvoiceDetails.Count() > 0 Then
        For Each vPID In PurchaseInvoiceDetails
          If vPID.LineNumber = pKey Then
            vFound = True
            Exit For
          End If
        Next vPID
      End If
      PIDExists = vFound
    End Function

    Public Function GetPurchaseOrderPayment(ByVal pKey As Integer, Optional ByRef pPurchaseOrderNumber As Integer = 0) As PurchaseOrderPayment
      Dim vPPA As PurchaseOrderPayment

      If PPAExists(pKey) Then
        vPPA = CType(PurchaseOrderPayments.Item(CStr(pKey)), PurchaseOrderPayment)
      Else
        vPPA = New PurchaseOrderPayment(mvEnv)
        vPPA.Init(pPurchaseOrderNumber, pKey)
        If vPPA.Existing Then vPPA.InitReadyForPayment()
        PurchaseOrderPayments.Add(vPPA, CStr(pKey))
      End If

      GetPurchaseOrderPayment = vPPA

    End Function

    Public Function PPAExists(ByVal pKey As Integer) As Boolean
      Dim vPPA As PurchaseOrderPayment
      Dim vFound As Boolean

      If PurchaseOrderPayments.Count() > 0 Then
        For Each vPPA In PurchaseOrderPayments
          If vPPA.PaymentNumber = pKey Then
            vFound = True
            Exit For
          End If
        Next vPPA
      End If
      PPAExists = vFound
    End Function

    Public Function GetPaymentPlanDetail(ByVal pKey As Integer) As PaymentPlanDetail
      Dim vPPD As PaymentPlanDetail

      If mvPaymentPlanDetails.Exists(CStr(pKey)) Then
        vPPD = mvPaymentPlanDetails(CStr(pKey))
      Else
        vPPD = mvPaymentPlanDetails.Add(CStr(pKey))
        vPPD.Init(mvEnv)
        vPPD.LineNumber = pKey
      End If
      GetPaymentPlanDetail = vPPD

    End Function

    Public Sub SetCreditCustomer(ByVal pContactNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pAddressNumber As Integer, ByVal pTermsNumber As String, ByVal pTermsPeriod As String, ByVal pTermsFrom As String, ByVal pCreditCategory As String, ByVal pCreditLimit As Double, ByVal pCustomerType As String, Optional ByVal pStopCode As String = "", Optional ByVal pCCChanged As Boolean = False, Optional ByVal pCreditCustomer As CreditCustomer = Nothing, Optional ByVal pDefaultTermsNumber As String = "", Optional ByVal pDefaultTermsPeriod As String = "", Optional ByVal pDefaultTermsFrom As String = "")
      'Set CreditCustomer object and determine whether it has changed

      '-----------------------------------------------------
      ' Determine whether new or changed Credit Customer
      '-----------------------------------------------------
      If pCreditCustomer Is Nothing Then
        'Came from Trader (Smart Client) so select the CreditCustomer
        mvCreditCustomer = New CreditCustomer
        mvCreditCustomer.Init(mvEnv, pContactNumber, pCompany, pSalesLedgerAccount)

        'See if anything has changed
        With mvCreditCustomer
          If ((Len(.TermsFrom) = 0 And (pTermsFrom = pDefaultTermsFrom)) And (Len(.TermsNumber) = 0 And (pTermsNumber = pDefaultTermsNumber)) And (Len(.TermsPeriod) = 0 And (pTermsPeriod = pDefaultTermsPeriod))) Then
            'Terms have not changed from the defaults
          Else
            If ((.TermsFrom <> pTermsFrom) Or (.TermsNumber <> pTermsNumber) Or (.TermsPeriod <> pTermsPeriod)) Then
              'Terms have changed
              mvCSCredCustomerChanged = True
            End If
          End If

          If ((.CreditCategory <> pCreditCategory) Or (.CreditLimit <> pCreditLimit) Or (.CustomerType <> pCustomerType) Or (.StopCode <> pStopCode) Or (.AddressNumber <> pAddressNumber)) Then
            mvCSCredCustomerChanged = True
          End If
        End With

      Else
        'Came from Trader (Thick Client) so use the existing class
        'Expect CreditCustomerChanged parameter to have been set
        mvCreditCustomer = pCreditCustomer
        mvCSCredCustomerChanged = pCCChanged
      End If

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvBatchNumber
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvTransactionNumber
      End Get
    End Property

    Public ReadOnly Property CSInvoiceCreated() As Boolean
      Get
        CSInvoiceCreated = mvCSInvoiceCreated
      End Get
    End Property

    Public ReadOnly Property CreditCustomerDetailsChanged() As Boolean
      Get
        CreditCustomerDetailsChanged = mvCSCredCustomerChanged
      End Get
    End Property

    Public ReadOnly Property Batch() As Batch
      Get
        Batch = mvBatch
      End Get
    End Property

    Public ReadOnly Property BatchTransaction() As BatchTransaction
      Get
        BatchTransaction = mvBT
      End Get
    End Property

    Public ReadOnly Property CreditCustomer() As CreditCustomer
      Get
        CreditCustomer = mvCreditCustomer
      End Get
    End Property

    Public ReadOnly Property TraderAnalysisLines() As TraderAnalysisLines
      Get
        TraderAnalysisLines = mvTraderAnalysisLines
      End Get
    End Property

    Public ReadOnly Property TraderInvoiceLines() As Collection
      Get
        If mvTraderInvoiceLines Is Nothing Then mvTraderInvoiceLines = New Collection
        TraderInvoiceLines = mvTraderInvoiceLines
      End Get
    End Property

    Public Property TraderPPDLines() As TraderPaymentPlanDetails
      Get
        TraderPPDLines = mvPaymentPlanDetails
      End Get
      Set(ByVal Value As TraderPaymentPlanDetails)
        mvPaymentPlanDetails = Value
      End Set
    End Property

    Public ReadOnly Property PurchaseOrderDetails() As Collection
      Get
        If mvPurchaseOrderDetails Is Nothing Then mvPurchaseOrderDetails = New Collection
        PurchaseOrderDetails = mvPurchaseOrderDetails
      End Get
    End Property

    Public ReadOnly Property PurchaseOrderPayments() As Collection
      Get
        If mvPurchaseOrderPayments Is Nothing Then mvPurchaseOrderPayments = New Collection
        PurchaseOrderPayments = mvPurchaseOrderPayments
      End Get
    End Property

    Public ReadOnly Property PurchaseInvoiceDetails() As Collection
      Get
        If mvPurchaseInvoiceDetails Is Nothing Then mvPurchaseInvoiceDetails = New Collection
        PurchaseInvoiceDetails = mvPurchaseInvoiceDetails
      End Get
    End Property

    Public ReadOnly Property Activities() As Collection
      Get
        If mvActivities Is Nothing Then mvActivities = New Collection
        Activities = mvActivities
      End Get
    End Property
    Public ReadOnly Property Suppressions() As Collection
      Get
        If mvSuppressions Is Nothing Then mvSuppressions = New Collection
        Suppressions = mvSuppressions
      End Get
    End Property

    Public Property IncentivesTable() As CDBDataTable
      Get
        IncentivesTable = mvIncentivesTable
      End Get
      Set(ByVal Value As CDBDataTable)
        mvIncentivesTable = Value
      End Set
    End Property

    Public Property BatchInvoices() As Collection
      Get
        If mvBatchInvoices Is Nothing Then mvBatchInvoices = New Collection
        BatchInvoices = mvBatchInvoices
      End Get
      Set(ByVal Value As Collection)
        mvBatchInvoices = Value
      End Set
    End Property

    Public ReadOnly Property PaymentPlan() As PaymentPlan
      Get
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan
          mvPaymentPlan.Init(mvEnv)
        End If
        PaymentPlan = mvPaymentPlan
      End Get
    End Property

    Public ReadOnly Property UpdatedOPS() As Collection
      Get
        If mvUpdatedOPS Is Nothing Then mvUpdatedOPS = New Collection
        UpdatedOPS = mvUpdatedOPS
      End Get
    End Property

    Public ReadOnly Property SummaryMembers() As CDBCollection
      Get
        If mvSummaryMembers Is Nothing Then
          mvSummaryMembers = New CDBCollection
        End If
        SummaryMembers = mvSummaryMembers
      End Get
    End Property

    Public ReadOnly Property OutstandingOPS() As Collection
      Get
        If mvOutstandingOPS Is Nothing Then mvOutstandingOPS = New Collection
        OutstandingOPS = mvOutstandingOPS
      End Get
    End Property

    Public ReadOnly Property RemovedSchPayments() As CDBCollection
      Get
        If mvRemovedSchPayments Is Nothing Then mvRemovedSchPayments = New CDBCollection
        Return mvRemovedSchPayments
      End Get
    End Property

    Public ReadOnly Property EventBookingLines() As CDBDataTable
      Get
        If mvEventBookingLines Is Nothing Then mvEventBookingLines = New CDBDataTable
        EventBookingLines = mvEventBookingLines
      End Get
    End Property

    Public ReadOnly Property ExamBookingLines() As CDBDataTable
      Get
        If mvExamBookingLines Is Nothing Then mvExamBookingLines = New CDBDataTable
        Return mvExamBookingLines
      End Get
    End Property
    ''' <summary>
    ''' BR19606 For Transaction History, Analysis followed by Edit or Delete will change the Order Payment Schedule, when Edit or Delete are clicked. This is the original order payment history before the change.
    ''' If the user cancels analysis we need the original (before analyis changed the OPS) OPS to restore the database to a stable conditon. The OPS needs to be to Smart Client so that cancel can use it. 
    ''' </summary>
    ''' <value>The order payment schedule before any changes, or Nothing</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OriginalOPS() As OrderPaymentSchedule
      Get
        Return mvOriginalOPS
      End Get
      Set(ByVal Value As OrderPaymentSchedule)
        mvOriginalOPS = Value
      End Set
    End Property

    Public Function GetMembershipJointContact(ByVal pMembershipTypeCode As String, ByVal pSourceCode As String) As Contact
      'Used by Smart Client to find/create a joint contact during membership creation and CMT
      Dim vJointContact As New Contact(mvEnv)
      Dim vContact1 As Contact = Nothing
      Dim vContact2 As Contact = Nothing
      Dim vMember As Member

      vJointContact.Init()

      If SummaryMembers.Count > 1 Then
        For Each vMember In SummaryMembers
          If vMember.MembershipTypeCode = pMembershipTypeCode Then
            If vContact1 Is Nothing Then
              vContact1 = vMember.Contact
            Else
              vContact2 = vMember.Contact
            End If
          End If
          If (Not (vContact1 Is Nothing)) And (Not (vContact2 Is Nothing)) Then Exit For
        Next vMember

        If (Not (vContact1 Is Nothing)) And (Not (vContact2 Is Nothing)) Then
          vJointContact.Init(vContact1.ProcessJointContact(vContact2, pSourceCode))
        End If
      End If

      GetMembershipJointContact = vJointContact

    End Function

    Public Sub SetTransactionOrigin(ByVal pTransactionOrigin As String)

      mvTransactionOrigin = pTransactionOrigin
    End Sub

    Public Function CMTOldPPDLines() As Collection
      If mvCMTOldPPDLines Is Nothing Then mvCMTOldPPDLines = New Collection
      Return mvCMTOldPPDLines
    End Function

    Public Function Actions() As CollectionList(Of Action)
      If mvActions Is Nothing Then mvActions = New CollectionList(Of Action)
      Return mvActions
    End Function

    Public Function AutoAddGiftAidDeclaration(source As String, method As String) As GiftAidDeclaration
      Dim result As GiftAidDeclaration = Nothing
      If mvBT Is Nothing Then
        Throw New InvalidOperationException("AutoAddGiftAidDeclaration called before batch transaction is initialised")
      End If
      Dim newDeclaration As GiftAidDeclaration = Nothing
      If mvBT.EligibleForGiftAid Then
        newDeclaration = New GiftAidDeclaration()
        newDeclaration.Init(mvEnv)
        newDeclaration.UpdateFields(
                                      mvBT.ContactNumber,
                                      mvBT.TransactionDate,
                                      If(source.IsNullOrWhitespace, mvBT.Analysis(0).Source, source),
                                      String.Empty,
                                      method,
                                      newDeclaration.GiftAidEarliestStartDate,
                                      String.Empty,
                                      String.Empty
                                   )

        Dim existingDeclaration As GiftAidDeclaration = newDeclaration.GetPreviousDeclaration()
        Dim transactionDate As Date = CDate(Me.BatchTransaction.TransactionDate)

        If existingDeclaration Is Nothing Then
          'No existing GAD, so create a new GAD.
          newDeclaration.Save()
          result = newDeclaration
        ElseIf existingDeclaration.CancellationReason.HasValue Then
          'Existing GAD is cancelled.  Create new GAD starting the day after
          Dim newStartDate As Date = CDate(existingDeclaration.EndDate).AddDays(1)
          newDeclaration.Update(newStartDate.ToString(CAREDateFormat), String.Empty, GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll, newDeclaration.Notes) 'There's no method to just change the Start Date.
          If (transactionDate - CDate(newDeclaration.StartDate)).Days >= 0 Then 'does the newly calculated start date cover the transaction date.BR21530
            newDeclaration.Save()
            result = newDeclaration
          End If
        ElseIf (transactionDate - CDate(existingDeclaration.StartDate)).Days <= 0 Then 'the existing declaration starts on or after the transaction date,   No declaration will be created.BR21530
          newDeclaration = Nothing
          result = Nothing
        Else
          existingDeclaration.EndDate = transactionDate.AddDays(-1).ToString(CAREDateFormat)
          existingDeclaration.Save()
          Dim newStartDate As Date = CDate(existingDeclaration.EndDate).AddDays(1) 'End Date has validation so let's take it back.
          newDeclaration.Update(newStartDate.ToString(CAREDateFormat), String.Empty, GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll, newDeclaration.Notes)
          newDeclaration.Save()
          result = newDeclaration
        End If
      End If
      Return result
    End Function
  End Class
End Namespace
