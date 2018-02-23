
Public Enum DefaultMailingTypes
  dmtNone
  dmtLetterBreaks
  dmtSource
  dmtLetterBreaksOrSource
End Enum

Public Enum ApplicationTypes
  atTransaction = 1           'TRANS
  atPurchaseInvoice           'PINVE
  atPurchaseOrder             'PORDE
  atPurchaseOrderCancellation 'PORDC
  atChequeNumberAllocation    'CHQNA
  atChequeReconciliation      'CHQRE
  atCreditStatementGeneration 'CSTAT
  atBatchInvoiceGeneration    'BINVG
  atMaintenance               'MAINT
  atConversion                'CNVRT
  atCreditListReconciliation  'CLREC
  atBankStatementPosting      'BSPOS
  atPurchaseOrderGeneration   'POGEN    does not use trader form
  atPurchaseOrderPrint        'POPRT    does not use trader form
  atChequeProcessing          'POCHQ    does not use trader form
  atGiveAsYouEarnPayments     'GAYEP    (Pre Tax Payroll Giving)
  atPostTaxPGPayments         'POTPG    (Post Tax Payroll Giving)
End Enum

Public Enum PurchaseOrderTypes
  None
  PaymentSchedule
  AdHocPayments
  RegularPayments
End Enum

Public Class TraderApplication
#Region " Declarations "
  Friend Enum TraderApplicationStartPoint
    taspTransaction
    taspRightMouse
    taspApplication
  End Enum

  'Application properties
  Public Description As String
  Public ApplicationNumber As Integer
  Public MainPageType As CareServices.TraderPageType
  Public CreditCard As Boolean
  Public CSCompany As String
  Public CSPaymentMethod As String
  Public CSTermsNumber As String
  Public CSTermsFrom As String
  Public CSTermsPeriod As String
  Public CSDepositPercentage As Double
  Public CSDepositAmount As Double
  Public DebitCard As Boolean
  Public Voucher As Boolean
  Public CAFCard As Boolean
  Public GiftInKind As Boolean
  Public SaleOrReturn As Boolean
  Public GiftAidMinimum As Double
  Public Provisional As Boolean
  Public DefaultMailingType As DefaultMailingTypes
  Public ShowTransactionReference As Boolean
  Public InvoiceDocument As String
  Public ServiceBookingCredits As Boolean
  Public ReceiptDocument As String
  Public ProvisionalCashTransactionDocument As String
  Public PaymentPlanDocument As String
  Public BatchLedApp As Boolean
  Public ProductCode As String
  Public RateCode As String
  Public SourceCode As String
  Public BatchCategory As String
  Public IncludeProvisionalTransactions As Boolean
  Public BypassMailingParagraphs As Boolean
  Public CACompany As String
  Public AutoSetAmount As Boolean
  Public ConfirmAnalysis As Boolean
  Public SourceFromLastMailing As Boolean
  Public PayMethodsAtEnd As Boolean
  Public PackToDonorDefault As Boolean
  Public DefaultCurrencyCode As String
  Public ApplicationType As ApplicationTypes
  Public FinancialAdjustment As BatchInfo.AdjustmentTypes = CDBNETCL.BatchInfo.AdjustmentTypes.None
  Public ConfirmPayPlanDetails As Boolean
  Public PayPlanConvMaintenance As Boolean
  Public ConversionShowPPD As Boolean
  Public MemberContactToAdd As Integer
  Public EditMemberLineNumber As Integer
  Public MailingCodeMandatory As Boolean
  Public NonFinancialBatch As Boolean
  Public BatchAnalysisCode As String
  Public CreditNotes As Boolean
  Public OnlineCCAuthorisation As Boolean
  Public RequireCCAuthorisation As Boolean
  Public DefaultActivityGroup As String
  Public DefaultSuppression As String
  Public MaintenanceOnly As Boolean
  Public DefaultStatus As String
  Public CheckIncentives As Boolean
  Public PreFulfilledIncentives As Boolean
  Public ContactAlerts As Boolean
  Public SalesGroup As String
  Public EventMultipleAnalysis As Boolean
  Public LinkToCommunication As String
  Public LinkToFundraisingPayments As Boolean
  Public PayPlanPayMethod As Boolean
  Public AllocationsChecked As Boolean
  Public Loans As Boolean
  Public LoanAmount As Double
  Public DonationProduct As String
  Public UnallocateCreditNote As Boolean
  Private mvTraderAlerts As Boolean

  Public PONumberOfPayments As Integer
  Public POPercentage As Boolean
  Public PurchaseOrderNumber As Integer
  Public PurchaseOrderScheduleChanged As Boolean
  Public PurchaseInvoiceNumber As Integer
  Public OriginalPayerContactNumber As Integer
  Private mvPurchaseOrderType As PurchaseOrderTypes = PurchaseOrderTypes.None 'To check if PO Type supports Regular Payments, Ad Hoc Payments etc
  Public OldPORegularPaymentAmount As Nullable(Of Double)  'Old Purchase Order Payment Amount for Regular Payments- Used to create purchase order history record

  'These two not actual TraderApplication properties
  Public BatchTypeCode As String
  Public ProvisionalCashBatch As Boolean
  Public BatchCurrencyCode As String
  Public BatchPaymentMethod As String

  'Transaction properties
  Public TransactionPaymentMethod As String         'CASH, CARD etc..
  Public TransactionType As String                  'DONR, MEMB etc..
  Public UnbalancedTransactionChoice As String = String.Empty      'TRAN,PLAN
  Public FATransactionType As String = ""
  Public PPPaymentType As String                    'STO?
  Public CurrentPaymentMethod As Boolean
  Public TransactionLines As Integer
  Public TransactionDate As String
  Public TransactionDateChanged As Boolean
  Public TransactionSource As String = ""
  Public TransactionDistributionCode As String
  Public SalesContactNumber As Integer
  Public EventNumber As String = ""
  Public BookingOptionNumber As String = ""
  Public EventBookingRate As String = ""
  Public TransactionAmount As Double
  Public LinePrice As Double                        'Set from Rate validation
  Public LinePriceVATEx As Boolean                  'Set from Rate validation
  Public FixedPrice As Boolean                      'Set from Rate validation
  Public LineVATAmount As Double                    'Set from SetAmount
  Public CurrentLineTotal As Double
  Public NewBank As Boolean                         'New bank required from Bank Details page
  Public BankDetailsNumber As Integer               'Bank Details number selected from Bank Details page
  Public CreditCardDetailsNumber As Integer         'Bank Details number selected from Bank Details page
  Public NewCreditCustomer As Boolean               'New credit customer from the credit customers page
  Public CreditTermsChanged As Boolean              'Credit customers terms have been changed
  Public SavePaymentTerms As Boolean                'Save changed payment terms
  Public CreditCustomerAddressChanged As Boolean
  Public SaveCreditCustomerAddressChange As Boolean
  Public CheckQuantityBreaks As Boolean
  Public OriginalTransactionAmount As Double          'Used when editing existing transactions to hold original Amount
  Public OriginalTransactionCurrencyAmount As Double  'Used when editing existing transactions to hold original CurrencyAmount
  Public TransactionReference As String = ""
  Public PPRenewalPeriodStartDate As String = String.Empty
  Public BankPaymentMethod As String = ""
  'Variables for Credit List Reconciliation
  Public StatementDate As String = String.Empty
  Public BankTransactionLineNumber As Integer
  Public PayersSortCode As String = String.Empty
  Public PayersAccountNumber As String = String.Empty
  Public PayersName As String = String.Empty
  Public ReferenceNumber As String = String.Empty
  Private mvCLRAdditionalCriteria As ParameterList
  'Variables for Payment Plans
  Public PPBalance As Double
  Public PPDLines As Integer
  Public CurrentPPDLineTotal As Double
  Public CurrentPPDAmount As Double
  Public CurrentPPDArrears As Double
  Public PPAmount As String = ""
  Public MemberCount As Integer
  Public ScheduledPaymentNumber As Integer
  Private mvPPDetailsPricing As PaymentPlanDetailsPricing

  'variable for payments
  Public MemberNumber As String
  Private mvPPNumbers As CollectionList(Of Integer)   ' I want to delete this ASAP, so please avoid using this
  Public CovenantNumber As Integer
  Public CancellationReason As String = ""
  Public CancellationSource As String = ""
  Public DeclarationNumber As Integer                 'Last GAD Number saved in a transaction
  Public TransactionDonationAmount As Double = 0      'Amount of donation to automatically add when transaction amount greater than lines amount

  'variables for service bookings
  Public SBGrossAmount As Double
  Public SBGrossQty As Double
  Public SBEntitlementQty As Double
  Public SBNewQuantity As Integer
  Public ConfirmSBDuration As Boolean
  Public ConfirmSBShortStay As Boolean
  Public ConfirmCalendarConflict As Boolean
  Public ServiceBookingAnalysis As Boolean
  Public BatchLocked As Boolean                       'Use to keep track if the batch has been locked
  Public FixedUnitRate As Boolean

  Private mvTraderPages As New CollectionList(Of TraderPage)
  Public AnalysisDataSet As DataSet
  Public CollectionBoxDataSet As DataSet
  Public PPDDataSet As DataSet
  Public OPSDataSet As DataSet
  Private mvOriginalOPS As DataTable 'BR19606: see property OriginalOPS for details
  Public OSPDataSet As DataSet
  Public MembersDataSet As DataSet
  Public PaymentPlan As PaymentPlanInfo
  Public PaymentPlans As CollectionList(Of PaymentPlanInfo)     'BR12989: Added new variable as mvPPNumbers would be removed at some point
  Public POSDataSet As DataSet
  Public PPADataSet As DataSet
  Public PISDataSet As DataSet
  Public IncentiveDataSet As DataSet
  Public BatchInvoicesDataSet As DataSet
  Public EventBookingDataSet As DataSet
  Public ExamBookingDataSet As DataSet
  Public RemovedSchPaymentsDataSet As DataSet
  Public CMTOldPPDDataSet As DataSet
  Public CMTNewPPDDataSet As DataSet


  Public BatchNumber As Integer                     'Used by BatchLed Applications to store batch selected by user
  Public TransactionNumber As Integer               'Used by BatchLed Applications to store batch selected by user
  Public BatchNumbers As String                     'Used for Multiple Selected Transactions ( its a comma seperated string )
  Public BatchDate As String                        'Used in Analysis of Batch Transaction
  Public PostToCashBook As String = ""              'Used in Analysis of Batch Transaction
  Private mvBatchInfo As BatchInfo
  Public ConfirmTransList As New StringBuilder
  Public FulfilIncentives As Boolean
  Public PPIncentivesCompleted As Boolean
  Public AddIncentivesLinesRequired As Boolean
  Public PPDMemberOrPayer As String = ""
  Public CMTOriginalMemberJoined As String = ""
  Public CMTPrevMembershipTypeCode As String = ""

  Private mvPayerContactNumber As Integer
  Private mvPayerAddressNumber As Integer
  Private mvDiscountActivityChecked As Boolean
  Private mvPayerHasDiscountActivity As Boolean
  Private mvPayerDiscountPercentage As Double
  Private mvSourceDiscountPercentage As Double
  Private mvContactVATCategory As String
  Private mvProductVATCategory As String
  Private mvVATInfoValid As Boolean
  Private mvVATRateInfo As VATRateInfo
  Private mvEditLineNumber As Integer              'When editing transactions, the line number currently being edited.
  Private mvAnalysisLinesDeleted As Integer = 0    'Count of Analysis lines deleted
  Private mvChangeMembershipType As Boolean
  Private mvCMTMemberNumber As String
  'Private mvCMTMembershipNumber As Integer
  Private mvCMTMemberContactNumber As Integer
  Private mvCMTMemberAddressNumber As Integer
  Public CMTPriceDate As String = ""
  Public CMTUpdatePPFixedAmount As Nullable(Of Boolean)
  Private mvPPDProductNumbersCount As Integer = 0    'To hold a count of the number of detail lines so far that have products with product numbers
  Private mvMemberships As Boolean
  Private mvEditPPDetailNumber As Integer = 0         'The PP DetailNumber currently being edited
  Private mvEditPPDSubscriptionNumber As Integer = 0  'The SubscriptionNumber on the PPD line currently being edited
  Private mvDeliveryContactNumber As Integer
  Private mvDeliveryAddressNumber As Integer
  Private mvCABankAccount As String = String.Empty

  'Stock Product sales
  Private mvStockSales As Boolean
  Private mvStockTransactionID As Integer
  Private mvStockIssued As Integer
  Private mvStockProductCode As String = ""
  Private mvStockWarehouseCode As String = ""
  Private mvStockQuantity As Integer
  Private mvNewPaymentFrequency As String = ""
  Private mvTransactionNote As String = ""
  Public WarehouseChanged As Boolean = False

  'These values are used to determine whether we need to re-set membership information
  Private mvLastMembershipType As String
  Private mvLastMembershipRate As String
  Private mvLastMembershipSource As String
  Private mvLastMembershipJoinedDate As String
  Private mvLastMembershipNumberMembers As Integer
  Private mvLastMembershipNumberAssociates As Integer
  Private mvLastMembershipFixedAmount As String
  Private mvMembershipNumber As Integer = 0

  'Values to store the Transaction's Last value that needs to be defaulted for other analysis lines
  Private mvLastDeceasedContactNumber As Integer

  Public DOBChanged As Boolean                          'BR11822: Flag to check if DOB is changed on TRD Page

  'Variables to hold answer of a question presented to the user
  Private mvChangeBranchWithAddress As String = Nothing
  Private mvCreateContactAccount As String = Nothing
  Private mvCreateCommLink As String = Nothing

  'Private mvPPCreated As Boolean
  Private mvAutoPaymentCreated As Boolean                        'BR12868: Flag to check if Auto Payment method created during Trader Transaction

  Private mvAppStartPoint As TraderApplicationStartPoint = TraderApplicationStartPoint.taspTransaction

  Private mvAlbacsBankDetails As String
  Public mvAutoCreateCreditCust As Boolean
  Public mvUnpostedBatchMsgInPrint As Boolean
  Public mvDateRangeMsgInPrint As Boolean
  Public mvCreditCategory As String = String.Empty
  Public mvChangedStartDate As String = String.Empty

#End Region

  Public Sub New(ByVal pApplicationNumber As Integer, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pIsDesign As Boolean = False, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pFinancialAdjustment As CareServices.AdjustmentTypes = 0)

    Dim vDataSet As DataSet = DataHelper.GetTraderApplication(pApplicationNumber, pBatchNumber, pIsDesign, pTransactionNumber, pFinancialAdjustment)
    Dim vTraderTable As DataTable = vDataSet.Tables("TraderApplication")
    With vTraderTable.Rows(0)
      Description = .Item("ApplicationDesc").ToString
      ApplicationNumber = CInt(.Item("Application"))
      MainPageType = CType(.Item("MainPage"), CareServices.TraderPageType)
      CSCompany = .Item("CSCompany").ToString
      CSPaymentMethod = .Item("CSPaymentMethod").ToString
      CSTermsFrom = .Item("CSTermsFrom").ToString
      CSTermsNumber = .Item("CSTermsNumber").ToString
      CSTermsPeriod = .Item("CSTermsPeriod").ToString
      CSDepositPercentage = DoubleValue(.Item("CSDepositPercentage").ToString)

      Voucher = .Item("Voucher").ToString = "Y"
      CAFCard = .Item("CAFCard").ToString = "Y"
      GiftInKind = .Item("GiftInKind").ToString = "Y"
      SaleOrReturn = .Item("SaleOrReturn").ToString = "Y"
      CreditCard = .Item("CreditCard").ToString = "Y"
      DebitCard = .Item("DebitCard").ToString = "Y"
      GiftAidMinimum = DoubleValue(.Item("GiftAidMinimum").ToString)
      DefaultMailingType = CType(.Item("DefaultMailingType"), DefaultMailingTypes)
      ShowTransactionReference = .Item("ShowTransactionReference").ToString = "Y"

      BatchLedApp = .Item("BatchLed").ToString = "Y"
      ProductCode = .Item("Product").ToString
      RateCode = .Item("Rate").ToString
      SourceCode = .Item("Source").ToString
      BatchCategory = .Item("BatchCategory").ToString
      IncludeProvisionalTransactions = .Item("IncludeProvisionalTransactions").ToString = "Y"
      BypassMailingParagraphs = .Item("BypassMailingParagraphs").ToString = "Y"

      BatchTypeCode = .Item("BatchTypeCode").ToString
      ProvisionalCashBatch = .Item("ProvisionalCashBatch").ToString = "Y"
      If .Table.Columns.Contains("BatchPaymentMethod") Then
        BatchPaymentMethod = .Item("BatchPaymentMethod").ToString
      End If

      ServiceBookingCredits = .Item("ServiceBookingCredits").ToString = "Y"
      InvoiceDocument = .Item("InvoiceDocument").ToString
      ReceiptDocument = .Item("ReceiptDocument").ToString
      ProvisionalCashTransactionDocument = .Item("ProvisionalCashTransactionDocument").ToString
      PaymentPlanDocument = .Item("PaymentPlanDocument").ToString
      CACompany = .Item("CACompany").ToString
      AutoSetAmount = .Item("AutoSetAmount").ToString = "Y"
      ConfirmAnalysis = .Item("ConfirmAnalysis").ToString = "Y"
      SourceFromLastMailing = .Item("SourceFromLastMailing").ToString = "Y"
      PayMethodsAtEnd = .Item("PayMethodsAtEnd").ToString = "Y"
      PackToDonorDefault = BooleanValue(.Item("PackToDonorDefault").ToString)
      DefaultCurrencyCode = .Item("DefaultCurrencyCode").ToString
      ApplicationType = GetApplicationType(.Item("ApplicationType").ToString)
      ConfirmPayPlanDetails = BooleanValue(.Item("ConfirmPPDetails").ToString)
      PayPlanConvMaintenance = BooleanValue(.Item("PPConvMaintenance").ToString)
      ConversionShowPPD = BooleanValue(.Item("ConversionShowPPD").ToString)
      mvChangeMembershipType = BooleanValue(.Item("ChangeMembershipType").ToString)
      mvMemberships = BooleanValue(.Item("memberships").ToString)
      MailingCodeMandatory = BooleanValue(.Item("MailingCodeMandatory").ToString)
      NonFinancialBatch = BooleanValue(.Item("NonFinancialBatch").ToString)
      BatchAnalysisCode = .Item("BatchAnalysisCode").ToString
      CreditNotes = BooleanValue(.Item("SundryCreditNotes").ToString)
      OnlineCCAuthorisation = BooleanValue(.Item("OnlineCCAuthorisation").ToString)
      RequireCCAuthorisation = BooleanValue(.Item("RequireCCAuthorisation").ToString)
      DefaultActivityGroup = .Item("DefaultActivityGroup").ToString
      DefaultSuppression = .Item("DefaultSuppression").ToString
      MaintenanceOnly = BooleanValue(.Item("MaintenanceOnly").ToString)
      DefaultStatus = .Item("DefaultStatus").ToString
      CheckIncentives = BooleanValue(.Item("CheckIncentives").ToString)
      PreFulfilledIncentives = BooleanValue(.Item("PreFulfilledIncentives").ToString)
      ContactAlerts = BooleanValue(.Item("ContactAlerts").ToString)
      SalesGroup = .Item("SalesGroup").ToString
      PayPlanPayMethod = BooleanValue(.Item("PayPlanPayMethod").ToString)
      DonationProduct = .Item("DonationProduct").ToString
      If .Table.Columns.Contains("EventMultipleAnalysis") Then EventMultipleAnalysis = BooleanValue(.Item("EventMultipleAnalysis").ToString)
      If .Table.Columns.Contains("LinkToCommunication") Then LinkToCommunication = .Item("LinkToCommunication").ToString
      If .Table.Columns.Contains("AlbacsBankDetails") Then mvAlbacsBankDetails = .Item("AlbacsBankDetails").ToString
      If .Table.Columns.Contains("LinkToFundraisingPayments") Then LinkToFundraisingPayments = BooleanValue(.Item("LinkToFundraisingPayments").ToString)
      If .Table.Columns.Contains("ServiceBookingAnalysis") Then ServiceBookingAnalysis = BooleanValue(.Item("ServiceBookingAnalysis").ToString)
      If .Table.Columns.Contains("Loans") Then Loans = BooleanValue(.Item("Loans").ToString)
      If .Table.Columns.Contains("AutoCreateCreditCustomer") Then mvAutoCreateCreditCust = BooleanValue(.Item("AutoCreateCreditCustomer").ToString)
      If .Table.Columns.Contains("UnpostedBatchMsgInPrint") Then mvUnpostedBatchMsgInPrint = BooleanValue(.Item("UnpostedBatchMsgInPrint").ToString)
      If .Table.Columns.Contains("DateRangeMsgInPrint") Then mvDateRangeMsgInPrint = BooleanValue(.Item("DateRangeMsgInPrint").ToString)
      If .Table.Columns.Contains("CreditCategory") Then mvCreditCategory = .Item("CreditCategory").ToString
      If .Table.Columns.Contains("CABankAccount") Then mvCABankAccount = .Item("CABankAccount").ToString
      If vTraderTable.Rows(0).Table.Columns.Contains("MerchantRetailNumber") Then
        Me.MerchantDetailnumber = vTraderTable.Rows(0).Item("MerchantRetailNumber").ToString
      End If
      Me.TraderAlerts = False
      If vTraderTable.Rows(0).Table.Columns.Contains("TraderAlerts") Then
        Me.TraderAlerts = BooleanValue(vTraderTable.Rows(0).Item("TraderAlerts").ToString)
      End If
    End With

    Dim vControlsTable As DataTable = vDataSet.Tables("TraderControls")
    Dim vPagesTable As DataTable = vDataSet.Tables("TraderPages")
    If vPagesTable IsNot Nothing Then
      For Each vRow As DataRow In vPagesTable.Rows
        Dim vPage As New TraderPage(vRow)
        mvTraderPages.Add(vPage.PageType.ToString, vPage)
        Dim vEPL As New EditPanel
        vPage.EditPanel = vEPL
        vEPL.Name = vPage.PageCode
        vEPL.Dock = DockStyle.Fill
        vEPL.Visible = False
        vControlsTable.DefaultView.RowFilter = "PageType = '" & vPage.PageType & "'"
        vEPL.Init(New EditPanelInfo(vPage, vControlsTable))
        If vPage.PageType = CareServices.TraderPageType.tpCollectionPayments Then
          vEPL.FindTextLookupBox("Source").ActiveOnly = False         'For this page ONLY, allow historic source codes
        End If
      Next
    End If

    If BatchDate Is Nothing Then
      BatchDate = ""
    End If

    BuildAnalysisDataSet()
    BuildCBDataSet()
    BuildPPDDataSet()
    BuildOPSDataSet(OPSDataSet)
    BuildOPSDataSet(RemovedSchPaymentsDataSet)
    BuildOSPDataSet()
    BuildMembersDataSet()
    BuildPOSDataSet()
    BuildPPADataSet()
    BuildPISDataSet()
    BuildBatchInvoicesDataSet()
    BuildEventBookingDataSet()
    BuildExamBookingDataSet()
    BuildCMTDataSets()
    NewTransaction()

  End Sub

  Public ReadOnly Property Pages() As CollectionList(Of TraderPage)
    Get
      Return mvTraderPages
    End Get
  End Property

  Public Property ContactVATCategory() As String
    Get
      Return mvContactVATCategory
    End Get
    Set(ByVal pValue As String)
      mvContactVATCategory = pValue
      mvVATInfoValid = False
    End Set
  End Property
  Public Property ProductVATCategory() As String
    Get
      Return mvProductVATCategory
    End Get
    Set(ByVal pValue As String)
      mvProductVATCategory = pValue
      mvVATInfoValid = False
    End Set
  End Property

  Public ReadOnly Property ShowVATExclusiveAmount() As Boolean
    Get
      Return LinePriceVATEx AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_vat_exclusive_method).ToUpper = "E"
    End Get
  End Property
  Public ReadOnly Property PayerContactNumber() As Integer
    Get
      Return mvPayerContactNumber
    End Get
  End Property
  Public ReadOnly Property PayerAddressNumber() As Integer
    Get
      Return mvPayerAddressNumber
    End Get
  End Property
  Public ReadOnly Property DiscountPercentage() As Double
    Get
      CheckDiscountActivity()
      If mvPayerHasDiscountActivity Then
        Return mvPayerDiscountPercentage
      Else
        Return mvSourceDiscountPercentage
      End If
    End Get
  End Property
  Public ReadOnly Property LineVATPercentage() As Double
    Get
      CheckVATInfo()
      Return mvVATRateInfo.CurrentPercentage(TransactionDate)
    End Get
  End Property
  Public ReadOnly Property LineVATRate() As String
    Get
      CheckVATInfo()
      Return mvVATRateInfo.VATRateCode
    End Get
  End Property

  Public Sub SetPayerContact(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer)
    If pContactNumber <> mvPayerContactNumber Then
      mvDiscountActivityChecked = False
      mvPayerHasDiscountActivity = False
      mvContactVATCategory = ""  'Blank this value out so that it can be repopulated for the new Payer Contact Number 
    End If
    mvPayerContactNumber = pContactNumber
    mvPayerAddressNumber = pAddressNumber
  End Sub

  Public WriteOnly Property SourceDiscountPercentage() As Double
    Set(ByVal pValue As Double)
      mvSourceDiscountPercentage = pValue
    End Set
  End Property

  Public ReadOnly Property PayerHasDiscount() As Boolean
    Get
      CheckDiscountActivity()
      Return mvPayerHasDiscountActivity OrElse mvSourceDiscountPercentage > 0
    End Get
  End Property

  Public ReadOnly Property EditExistingTransaction() As Boolean
    Get
      If BatchNumber > 0 AndAlso TransactionNumber > 0 Then
        Return ExistingAdjustmentTran = False
      End If
    End Get
  End Property
  Public ReadOnly Property ExistingAdjustmentTran() As Boolean
    Get
      If BatchNumber > 0 AndAlso TransactionNumber > 0 Then
        Select Case FinancialAdjustment
          Case CDBNETCL.BatchInfo.AdjustmentTypes.Move, CDBNETCL.BatchInfo.AdjustmentTypes.Adjustment, CDBNETCL.BatchInfo.AdjustmentTypes.EventAdjustment, _
               CDBNETCL.BatchInfo.AdjustmentTypes.GIKConfirmation, CDBNETCL.BatchInfo.AdjustmentTypes.CashBatchConfirmation
            Return True
          Case Else
            Return False
        End Select
      End If
    End Get
  End Property

  Public Property BatchInfo() As BatchInfo
    Get
      If mvBatchInfo Is Nothing Then mvBatchInfo = New BatchInfo(BatchNumber)
      Return mvBatchInfo
    End Get
    Set(ByVal pValue As BatchInfo)
      mvBatchInfo = pValue
    End Set
  End Property

  Public Property EditLineNumber() As Integer
    Get
      Return mvEditLineNumber
    End Get
    Set(ByVal pValue As Integer)
      mvEditLineNumber = pValue
    End Set
  End Property

  Friend Property EditPPDetailNumber() As Integer
    Get
      Return mvEditPPDetailNumber
    End Get
    Set(ByVal pValue As Integer)
      mvEditPPDetailNumber = pValue
    End Set
  End Property

  Friend WriteOnly Property EditPPDSubscriptionNumber() As Integer
    Set(ByVal pValue As Integer)
      mvEditPPDSubscriptionNumber = pValue
    End Set
  End Property

  Public ReadOnly Property LastMembershipType() As String
    Get
      Return mvLastMembershipType
    End Get
  End Property

  Public ReadOnly Property LastMembershipRate() As String
    Get
      Return mvLastMembershipRate
    End Get
  End Property

  Public ReadOnly Property LastMembershipJoinedDate() As String
    Get
      Return mvLastMembershipJoinedDate
    End Get
  End Property

  Public ReadOnly Property LastMembershipSource() As String
    Get
      Return mvLastMembershipSource
    End Get
  End Property

  Public ReadOnly Property LastMembershipNumberMembers() As Integer
    Get
      Return mvLastMembershipNumberMembers
    End Get
  End Property

  Public ReadOnly Property LastMembershipNumberAssociates() As Integer
    Get
      Return mvLastMembershipNumberAssociates
    End Get
  End Property

  Friend ReadOnly Property LastMembershipFixedAmount() As String
    Get
      Return mvLastMembershipFixedAmount
    End Get
  End Property

  Friend ReadOnly Property ChangeMembershipType() As Boolean
    Get
      Return mvChangeMembershipType
    End Get
  End Property

  Friend ReadOnly Property CMTMemberNumber() As String
    Get
      Return mvCMTMemberNumber
    End Get
  End Property

  'Friend ReadOnly Property CMTMembershipNumber() As Integer
  '  Get
  '    Return mvCMTMembershipNumber
  '  End Get
  'End Property

  Friend ReadOnly Property CMTMemberContactNumber() As Integer
    Get
      Return mvCMTMemberContactNumber
    End Get
  End Property

  Friend ReadOnly Property CMTMemberAddressNumber() As Integer
    Get
      Return mvCMTMemberAddressNumber
    End Get
  End Property

  Public ReadOnly Property PPDProductNumbersCount() As Integer
    Get
      Return mvPPDProductNumbersCount
    End Get
  End Property
  Public Property TransactionNote() As String
    Get
      Return mvTransactionNote
    End Get
    Set(ByVal value As String)
      mvTransactionNote = value
    End Set
  End Property
  ''' <summary>
  ''' For Transaction History, Analysis followed by Edit or Delete will change the Order Payment Schedule, when Edit or Delete are clicked. This is the original order payment history before the change.
  ''' It is used to restore the Order Payment Schedule if Analysis is Cancelled.
  ''' </summary>
  ''' <value>The order payment schedule record for the tranascation being analysed, as it was before analysis started.</value>
  ''' <returns></returns>
  ''' <remarks>BR19606 On no account change the contents of this datatable, it presence indicates the database in in an unstable state, and is required to be return to the server unchanged to restore stability</remarks>
  Public Property OriginalOPS() As DataTable
    Get
      Return mvOriginalOPS
    End Get
    Set(ByVal value As DataTable)
      mvOriginalOPS = value
    End Set
  End Property

  Dim mvMerchantDetailnumber As String = String.Empty
  Public Property MerchantDetailnumber() As String
    Get
      Return mvMerchantDetailnumber
    End Get
    Private Set(ByVal value As String)
      mvMerchantDetailnumber = value
    End Set
  End Property

  Private Sub CheckDiscountActivity()
    If Not mvDiscountActivityChecked Then
      If mvPayerContactNumber > 0 Then
        Dim vActivity As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_discount_activity)
        If vActivity.Length > 0 Then
          Dim vList As New ParameterList(True)
          vList("Activity") = vActivity
          vList("Current") = "Y"
          Dim vRow As DataRow = Nothing
          Try
            vRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactCategories, mvPayerContactNumber, vList, True)
          Catch vEx As CareException
            'BR11964: If the discount activity is invalid (has been deleted) then we are going to treat it as if it has not been set on the contact
            If vEx.ErrorNumber <> CareException.ErrorNumbers.enParameterInvalidValue Then Throw vEx
          End Try
          If vRow IsNot Nothing Then
            mvPayerHasDiscountActivity = True
            mvPayerDiscountPercentage = DoubleValue(vRow("Quantity").ToString)
          End If
        End If
        mvDiscountActivityChecked = True
      End If
    End If
  End Sub

  Private Sub CheckVATInfo()
    If mvVATInfoValid = False Then
      If ContactVATCategory.Length = 0 And PayerContactNumber > 0 Then
        Dim vContactInfo As New ContactInfo(PayerContactNumber)
        If vContactInfo IsNot Nothing Then
          ContactVATCategory = vContactInfo.VATCategory
        End If
      End If
      mvVATRateInfo = New VATRateInfo(ContactVATCategory, ProductVATCategory)
    End If
  End Sub

  Private Sub BuildAnalysisDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "TraderLineType", "Type")
    DataHelper.AddDataColumn(vTable, "ProductCode", "Product")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
    DataHelper.AddDataColumn(vTable, "Quantity", "Qty", "Long")
    DataHelper.AddDataColumn(vTable, "Source", "Source")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount", "Number")
    DataHelper.AddDataColumn(vTable, "AcceptAsFull", "Accept")
    DataHelper.AddDataColumn(vTable, "LineDate", "Date", "Date")
    DataHelper.AddDataColumn(vTable, "DespatchMethod", "Method")
    DataHelper.AddDataColumn(vTable, "DeliveryContactNumber", "Contact", "Long")
    DataHelper.AddDataColumn(vTable, "DeliveryAddressNumber", "Address", "Long")
    DataHelper.AddDataColumn(vTable, "VatRate", "Rate")
    DataHelper.AddDataColumn(vTable, "VATPercentage", "VAT%", "Numeric")
    DataHelper.AddDataColumn(vTable, "ProductNumber", "Product Number", "Long")
    DataHelper.AddDataColumn(vTable, "GiverContactNumber", "Giver Contact", "Long")
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentNumber", "Scheduled Payment No", "Long")
    DataHelper.AddDataColumn(vTable, "DeceasedContactNumber", "Deceased Contact", "Long")
    DataHelper.AddDataColumn(vTable, "PaymentPlanNumber", "Payment Plan", "Long")
    DataHelper.AddDataColumn(vTable, "MemberNumber", "Member", "Long")
    DataHelper.AddDataColumn(vTable, "CollectionNumber", "Collection No", "Long")
    DataHelper.AddDataColumn(vTable, "InvoiceNumber", "Invoice Number", "Long")
    DataHelper.AddDataColumn(vTable, "InvoiceNumberUsed", "Invoice Number Used", "Long")
    DataHelper.AddDataColumn(vTable, "Notes", "Notes", "Char")       'Keep this as the last visible column
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "CovenantNumber", "Covenant", "Long", "N")
    DataHelper.AddDataColumn(vTable, "StockSale", "Stock Sale", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Issued", "Stock Issued", "Long", "N")
    DataHelper.AddDataColumn(vTable, "StockTransactionID", "Stock ID", "Long", "N")
    DataHelper.AddDataColumn(vTable, "VATAmount", "VAT", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "PriceVATExclusive", "VAT Excl", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Discount", "Discount", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "EventBookingNumber", "Booking Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "AmendedEventBookingNumber", "Amended Booking Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "CreditedContactNumber", "Credited Contact", "Long", "N")
    DataHelper.AddDataColumn(vTable, "InvoiceTypeUsed", "Invoice Type Used", "Char", "N")
    DataHelper.AddDataColumn(vTable, "DepositAllowed", "Deposit Allowed", "Char", "N")

    AnalysisDataSet = New DataSet
    AnalysisDataSet.Tables.Add(vTable)
  End Sub
  Private Sub BuildCBDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "BoxReference", "Box Reference")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount", "Number")
    DataHelper.AddDataColumn(vTable, "Pay", "Pay?", "Boolean")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "CollectionBoxNumber", "Box No", "Long", "N")
    DataHelper.AddDataColumn(vTable, "ContactNumber", "Collector", "Long", "N")

    CollectionBoxDataSet = New DataSet
    CollectionBoxDataSet.Tables.Add(vTable)
  End Sub
  Private Sub BuildPPDDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "Product", "Product")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
    DataHelper.AddDataColumn(vTable, "Quantity", "Quantity", "Long")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount", "Number")
    DataHelper.AddDataColumn(vTable, "Balance", "Balance", "Number")
    DataHelper.AddDataColumn(vTable, "Arrears", "Arrears", "Number")
    DataHelper.AddDataColumn(vTable, "Source", "Source")
    DataHelper.AddDataColumn(vTable, "DespatchMethod", "Method")
    DataHelper.AddDataColumn(vTable, "ContactNumber", "Contact Number")
    DataHelper.AddDataColumn(vTable, "AddressNumber", "Address Number")
    DataHelper.AddDataColumn(vTable, "CommunicationNumber", "Communication Number", "Long")
    DataHelper.AddDataColumn(vTable, "EffectiveDate", "Effective Date", "Date")
    DataHelper.AddDataColumn(vTable, "NetFixedAmount", "Net Fixed Amount", "Number")
    DataHelper.AddDataColumn(vTable, "ModifierActivity", "Activity")
    DataHelper.AddDataColumn(vTable, "ModifierActivityValue", "Activity Value")
    DataHelper.AddDataColumn(vTable, "ModifierActivityQuantity", "Activity Quantity", "Long")
    DataHelper.AddDataColumn(vTable, "ModifierActivityDate", "Activity Date", "Date")
    DataHelper.AddDataColumn(vTable, "ModifierPrice", "Modifier Price", "Number")
    DataHelper.AddDataColumn(vTable, "ModifierPerItem", "Per Item")
    DataHelper.AddDataColumn(vTable, "UnitPrice", "Unit Price", "Number")
    DataHelper.AddDataColumn(vTable, "ProRated", "Pro-Rated?")
    DataHelper.AddDataColumn(vTable, "NetAmount", "Net", "Number")
    DataHelper.AddDataColumn(vTable, "VatAmount", "VAT", "Number")
    DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross", "Number")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "PaymentPlanNumber", "Plan Number", , "N")
    DataHelper.AddDataColumn(vTable, "DetailNumber", "Detail Number", , "N")
    DataHelper.AddDataColumn(vTable, "TimeStatus", "Time Status", , "N")
    DataHelper.AddDataColumn(vTable, "ProductNumber", "Product Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "AmendedBy", "Amended By", , "N")
    DataHelper.AddDataColumn(vTable, "AmendedOn", "Amended On", "Date", "N")
    DataHelper.AddDataColumn(vTable, "CreatedBy", "Created By", , "N")
    DataHelper.AddDataColumn(vTable, "CreatedOn", "Created On", "Date", "N")
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", , "N")
    DataHelper.AddDataColumn(vTable, "SubscriptionNumber", "Subscription Number", , "N")
    DataHelper.AddDataColumn(vTable, "PPDLineType", "Line Type", "Long", "N")
    DataHelper.AddDataColumn(vTable, "MemberOrPayer", "Member or Payer", , "N")
    DataHelper.AddDataColumn(vTable, "AccruesInterest", "Loan Capital", "Char", "N")
    DataHelper.AddDataColumn(vTable, "LoanInterest", "Loan Interest", "Char", "N")
    DataHelper.AddDataColumn(vTable, "VatRate", "Vat Rate", "Char", "N")
    DataHelper.AddDataColumn(vTable, "VatPercentage", "VAT %", "Number", "N")
    DataHelper.AddDataColumn(vTable, "IncentiveLineType", "Inc Line Type", "Char", "N")
    DataHelper.AddDataColumn(vTable, "IncentiveIgnoreProductAndRate", "Inc Ignore PR", "Char", "N")

    PPDDataSet = New DataSet
    PPDDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildOPSDataSet(ByRef pDataSet As DataSet)
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentNumber", "Number", "Long")
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentStatusDesc", "Status")
    DataHelper.AddDataColumn(vTable, "DueDate", "Due", "Date")
    DataHelper.AddDataColumn(vTable, "ClaimDate", "Claim", "Date")
    DataHelper.AddDataColumn(vTable, "AmountDue", "Amount Due", "Numeric")
    DataHelper.AddDataColumn(vTable, "AmountOutstanding", "Outstanding", "Numeric")
    DataHelper.AddDataColumn(vTable, "ExpectedBalance", "Expected Balance", "Numeric")
    DataHelper.AddDataColumn(vTable, "RevisedAmount", "Revised Amount", "Numeric")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "PaymentPlanNumber", "Plan Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentStatus", "Status Code", , "N")
    DataHelper.AddDataColumn(vTable, "ScheduleCreationReason", "Creation Code", , "N")
    DataHelper.AddDataColumn(vTable, "OrigAmountDue", "Orig Amount Due", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", "Long", "N")

    pDataSet = New DataSet
    pDataSet.Tables.Add(vTable)
  End Sub

  Public Sub BuildOSPDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentNumber", "Payment", "Long")
    DataHelper.AddDataColumn(vTable, "DueDate", "Due", "Date")
    DataHelper.AddDataColumn(vTable, "AmountDue", "Amount Due", "Numeric")
    DataHelper.AddDataColumn(vTable, "AmountOutstanding", "Outstanding", "Numeric")
    DataHelper.AddDataColumn(vTable, "RevisedAmount", "Revised", "Numeric")
    DataHelper.AddDataColumn(vTable, "CheckValue", "Pay?", "Boolean")
    DataHelper.AddDataColumn(vTable, "PaymentAmount", "Allocated", "Numeric")

    'Hidden columns
    DataHelper.AddDataColumn(vTable, "PaymentPlanNumber", "Plan Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentStatus", "Status Code", , "N")
    DataHelper.AddDataColumn(vTable, "ScheduleCreationReason", "Creation Code", , "N")
    DataHelper.AddDataColumn(vTable, "OrigAmountDue", "Orig Amount Due", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "ScheduledPaymentStatusDesc", "Status", "N")
    DataHelper.AddDataColumn(vTable, "ClaimDate", "Claim", "Date", "N")
    DataHelper.AddDataColumn(vTable, "ExpectedBalance", "Expected Balance", "Numeric", "N")
    OSPDataSet = New DataSet
    OSPDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildMembersDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "ContactNumber", "Contact", "Long")
    DataHelper.AddDataColumn(vTable, "MembershipType", "Type")
    DataHelper.AddDataColumn(vTable, "ContactName", "Name")
    DataHelper.AddDataColumn(vTable, "Joined", "Joined", "Date")
    DataHelper.AddDataColumn(vTable, "Branch", "Branch")
    DataHelper.AddDataColumn(vTable, "BranchMember", "Branch Member")
    DataHelper.AddDataColumn(vTable, "Applied", "Applied", "Date")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "AddressNumber", "Address Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "MembershipNumber", "Membership Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "DateOfBirth", "Date Of Birth", "Date", "N")
    DataHelper.AddDataColumn(vTable, "DOBEstimated", "DOB Estimated", , "N")
    DataHelper.AddDataColumn(vTable, "AgeOverride", "Age Override", "Long", "N")
    DataHelper.AddDataColumn(vTable, "AddressLine", "Address", , "N")

    MembersDataSet = New DataSet
    MembersDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildPOSDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "LineItem", "Item")
    DataHelper.AddDataColumn(vTable, "LinePrice", "Price")
    DataHelper.AddDataColumn(vTable, "Quantity", "Quantity")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount")
    DataHelper.AddDataColumn(vTable, "NominalAccount", "NominalAccount")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
    DataHelper.AddDataColumn(vTable, "Product", "Product")
    DataHelper.AddDataColumn(vTable, "Warehouse", "Warehouse")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", "Long", "N")

    POSDataSet = New DataSet
    POSDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildPISDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "LineItem", "Item")
    DataHelper.AddDataColumn(vTable, "LinePrice", "Price")
    DataHelper.AddDataColumn(vTable, "Quantity", "Quantity")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount")
    DataHelper.AddDataColumn(vTable, "NominalAccount", "NominalAccount")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", "Long", "N")

    PISDataSet = New DataSet
    PISDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildPPADataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable

    DataHelper.AddDataColumn(vTable, "DueDate", "Due Date", "Date")
    DataHelper.AddDataColumn(vTable, "LatestExpectedDate", "Expected Date", "Date")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount", "Numeric")
    DataHelper.AddDataColumn(vTable, "Percentage", "Percentage")
    DataHelper.AddDataColumn(vTable, "AuthorisationRequired", "Auth")
    DataHelper.AddDataColumn(vTable, "AuthorisationStatus", "Status")
    DataHelper.AddDataColumn(vTable, "PostedOn", "Posted", "Date")

    DataHelper.AddDataColumn(vTable, "PayeeContactNumber", "Payee Contact Number", "Long")
    DataHelper.AddDataColumn(vTable, "Finder", "Find", "Long")
    DataHelper.AddDataColumn(vTable, "ContactName", "Payee Contact Name")
    DataHelper.AddDataColumn(vTable, "PayeeAddressNumber", "Payee Address")

    'DataHelper.AddDataColumn(vTable, "PayByBacs", "Pay By BACS", "Char")
    DataHelper.AddDataColumn(vTable, "PopPaymentMethod", "Pop Payment Method", "Char")
    DataHelper.AddDataColumn(vTable, "PoPaymentType", "Purchase Order Payment Type", "Char")
    DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code", "Char")
    DataHelper.AddDataColumn(vTable, "NominalAccount", "Nominal Acc", "Char")
    DataHelper.AddDataColumn(vTable, "SeparatePayment", "Separate Payment", "Char")

    'Hidden columns
    DataHelper.AddDataColumn(vTable, "PaymentNumber", "Payment Number", "Long", "N")
    DataHelper.AddDataColumn(vTable, "ReadyForPayment", "Ready For Payment", "Char", "N")
    PPADataSet = New DataSet
    PPADataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildBatchInvoicesDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "Print", "Print", "Boolean")
    DataHelper.AddDataColumn(vTable, "RecordType", "Record Type")
    DataHelper.AddDataColumn(vTable, "InvoiceDate", "Invoice Date")
    DataHelper.AddDataColumn(vTable, "InvoiceNumber", "Invoice Number")
    DataHelper.AddDataColumn(vTable, "ContactNumber", "Contact Number")
    DataHelper.AddDataColumn(vTable, "Name", "Payers Name")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount")
    DataHelper.AddDataColumn(vTable, "VATAmount", "VAT Amount")
    DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross Amount")
    DataHelper.AddDataColumn(vTable, "EventDescription", "Event Description")
    DataHelper.AddDataColumn(vTable, "EventNumber", "Event Number")
    DataHelper.AddDataColumn(vTable, "Company", "Company")
    DataHelper.AddDataColumn(vTable, "BatchNumber", "Batch Number")
    DataHelper.AddDataColumn(vTable, "TransactionNumber", "Transaction Number")
    DataHelper.AddDataColumn(vTable, "SalesLedgerAccount", "Sales Ledger Account")
    'Hidden column
    BatchInvoicesDataSet = New DataSet
    BatchInvoicesDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildEventBookingDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "Product", "Product")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate")
    DataHelper.AddDataColumn(vTable, "Quantity", "Quantity")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount")
    DataHelper.AddDataColumn(vTable, "VATAmount", "VAT Amount")
    DataHelper.AddDataColumn(vTable, "VATRate", "VAT Rate")
    DataHelper.AddDataColumn(vTable, "VATPercentage", "Percentage")
    DataHelper.AddDataColumn(vTable, "Notes", "Notes")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line No.", "Char", "N")

    EventBookingDataSet = New DataSet
    EventBookingDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildExamBookingDataSet()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "Product", "Product")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate")
    DataHelper.AddDataColumn(vTable, "Quantity", "Quantity")
    DataHelper.AddDataColumn(vTable, "Amount", "Amount")
    DataHelper.AddDataColumn(vTable, "VATAmount", "VAT Amount")
    DataHelper.AddDataColumn(vTable, "VATRate", "VAT Rate")
    DataHelper.AddDataColumn(vTable, "VATPercentage", "Percentage")
    DataHelper.AddDataColumn(vTable, "Notes", "Notes")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "LineNumber", "Line No.", "Char", "N")
    DataHelper.AddDataColumn(vTable, "ExamUnitId", "Unit No.", "Char", "N")
    DataHelper.AddDataColumn(vTable, "ExamUnitProductId", "Line No.", "Char", "N")
    ExamBookingDataSet = New DataSet
    ExamBookingDataSet.Tables.Add(vTable)
  End Sub

  Private Sub BuildCMTDataSets()
    Dim vTable As DataTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "ProductDesc", "Product")
    DataHelper.AddDataColumn(vTable, "RateDesc", "Rate")
    DataHelper.AddDataColumn(vTable, "Balance", "Balance", "Numeric")
    DataHelper.AddDataColumn(vTable, "FullPrice", "Full Price", "Numeric")
    DataHelper.AddDataColumn(vTable, "ProratedPrice", "Pro-rated Price", "Numeric")
    DataHelper.AddDataColumn(vTable, "ExcessAmount", "Excess Amount", "Numeric")
    DataHelper.AddDataColumn(vTable, "CMTProrateCost", "Pro-rate?")
    DataHelper.AddDataColumn(vTable, "CMTExcessPaymentType", "Handle Excess")
    DataHelper.AddDataColumn(vTable, "EntitlementSequenceNumber", "Sequence")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "DetailNumber", "Number", "long", "N")
    DataHelper.AddDataColumn(vTable, "CMTProrateCostCode", "ProRate Type", "Char", "N")
    DataHelper.AddDataColumn(vTable, "CMTExcessPaymentTypeCode", "RefundType", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Product", "Product", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate", "Char", "N")
    DataHelper.AddDataColumn(vTable, "PPDLineType", "Line Type", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "CMTRefundProductCode", "Refund Product", "Char", "N")
    DataHelper.AddDataColumn(vTable, "CMTRefundRateCode", "Refund Rate", "Char", "N")
    DataHelper.AddDataColumn(vTable, "UnitPrice", "Unit Price", "Number", "N")
    DataHelper.AddDataColumn(vTable, "ProRated", "Pro-Rated?", "Char", "N")
    DataHelper.AddDataColumn(vTable, "NetAmount", "Net", "Number", "N")
    DataHelper.AddDataColumn(vTable, "VatAmount", "VAT", "Number", "N")
    DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross", "Number", "N")

    CMTOldPPDDataSet = New DataSet
    CMTOldPPDDataSet.Tables.Add(vTable)

    vTable = DataHelper.NewColumnTable
    DataHelper.AddDataColumn(vTable, "ProductDesc", "Product")
    DataHelper.AddDataColumn(vTable, "RateDesc", "Rate")
    DataHelper.AddDataColumn(vTable, "FullPrice", "Full Price", "Numeric")
    DataHelper.AddDataColumn(vTable, "ProratedPrice", "Pro-rated Price", "Numeric")
    DataHelper.AddDataColumn(vTable, "CMTProrateCost", "Pro-rate?")
    DataHelper.AddDataColumn(vTable, "EntitlementSequenceNumber", "Sequence")
    'Hidden columns
    DataHelper.AddDataColumn(vTable, "DetailNumber", "Number", "long", "N")
    DataHelper.AddDataColumn(vTable, "CMTProrateCostCode", "ProRate Type", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Product", "Product", "Char", "N")
    DataHelper.AddDataColumn(vTable, "Rate", "Rate", "Char", "N")
    DataHelper.AddDataColumn(vTable, "PPDLineType", "Line Type", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "Balance", "Balance", "Numeric", "N")
    DataHelper.AddDataColumn(vTable, "Refund", "Handle Excess", "Char", "N")
    DataHelper.AddDataColumn(vTable, "UnitPrice", "Unit Price", "Number", "N")
    DataHelper.AddDataColumn(vTable, "ProRated", "Pro-Rated?", "Char", "N")
    DataHelper.AddDataColumn(vTable, "NetAmount", "Net", "Number", "N")
    DataHelper.AddDataColumn(vTable, "VatAmount", "VAT", "Number", "N")
    DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross", "Number", "N")
    CMTNewPPDDataSet = New DataSet
    CMTNewPPDDataSet.Tables.Add(vTable)

  End Sub

  Public Sub NewTransaction()
    TransactionPaymentMethod = "CASH"     'Default required for PayMethodsAtEnd
    TransactionType = ""
    PPPaymentType = "CASH"
    CurrentPaymentMethod = False
    TransactionLines = 0
    'BR15961: transaction date was set in a previous transaction on this trader app - should be used again to match RC functionality so don't reset it
    If TransactionDate Is Nothing Then TransactionDate = ""
    TransactionDateChanged = False
    If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_retain_source) Then TransactionSource = ""
    If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ev_retain_trader_booking_dets) Then
      EventNumber = ""
      BookingOptionNumber = ""
      EventBookingRate = ""
    End If
    TransactionDistributionCode = ""
    mvPayerContactNumber = 0
    mvPayerAddressNumber = 0
    mvDiscountActivityChecked = False
    mvPayerHasDiscountActivity = False
    mvPayerDiscountPercentage = 0
    mvSourceDiscountPercentage = 0
    mvProductVATCategory = ""
    mvContactVATCategory = ""
    mvVATInfoValid = False
    'SalesContactNumber = 0  Don't clear sales contact number between transactions so we always have the default
    TransactionAmount = 0
    CurrentLineTotal = 0
    NewBank = False
    BankDetailsNumber = 0
    CreditCardDetailsNumber = 0
    NewCreditCustomer = False
    CreditTermsChanged = False
    SavePaymentTerms = False
    CreditCustomerAddressChanged = False
    SaveCreditCustomerAddressChange = False
    CheckQuantityBreaks = False
    TransactionNumber = 0
    OriginalTransactionAmount = 0
    OriginalTransactionCurrencyAmount = 0
    mvEditLineNumber = 0
    mvEditPPDetailNumber = 0
    mvEditPPDSubscriptionNumber = 0
    TransactionReference = ""
    mvAnalysisLinesDeleted = 0
    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_retain_product_rate) = False Then
      mvLastMembershipType = ""
      mvLastMembershipRate = ""
    End If
    If Not BatchLedApp Then
      BatchCurrencyCode = ""
    End If
    DOBChanged = False

    mvLastMembershipJoinedDate = ""   'Always clear the LastMembershipJoinedDate & Source
    mvLastMembershipSource = ""
    mvCMTMemberNumber = ""
    'mvCMTMembershipNumber = 0
    mvCMTMemberContactNumber = 0
    CMTPriceDate = ""
    CMTUpdatePPFixedAmount = Nothing
    MemberCount = 0
    mvMembershipNumber = 0
    mvStockSales = False
    mvStockTransactionID = 0
    mvStockIssued = 0
    mvStockProductCode = ""
    mvStockWarehouseCode = ""
    mvStockQuantity = 0
    WarehouseChanged = False
    If AnalysisDataSet.Tables.Contains("DataRow") Then AnalysisDataSet.Tables.Remove("DataRow")
    If CollectionBoxDataSet.Tables.Contains("DataRow") Then CollectionBoxDataSet.Tables.Remove("DataRow")
    If PPDDataSet.Tables.Contains("DataRow") Then PPDDataSet.Tables.Remove("DataRow")
    If OPSDataSet.Tables.Contains("DataRow") Then OPSDataSet.Tables.Remove("DataRow")
    If RemovedSchPaymentsDataSet.Tables.Contains("DataRow") Then RemovedSchPaymentsDataSet.Tables.Remove("DataRow")
    If OSPDataSet.Tables.Contains("DataRow") Then OSPDataSet.Tables.Remove("DataRow")
    If MembersDataSet.Tables.Contains("DataRow") Then MembersDataSet.Tables.Remove("DataRow")
    If POSDataSet.Tables.Contains("DataRow") Then POSDataSet.Tables.Remove("DataRow")
    If PISDataSet.Tables.Contains("DataRow") Then PISDataSet.Tables.Remove("DataRow")
    If PPADataSet.Tables.Contains("DataRow") Then PPADataSet.Tables.Remove("DataRow")
    If BatchInvoicesDataSet.Tables.Contains("DataRow") Then BatchInvoicesDataSet.Tables.Remove("DataRow")
    If EventBookingDataSet.Tables.Contains("DataRow") Then EventBookingDataSet.Tables.Remove("DataRow")
    If ExamBookingDataSet.Tables.Contains("DataRow") Then ExamBookingDataSet.Tables.Remove("DataRow")
    If CMTOldPPDDataSet.Tables.Contains("DataRow") Then CMTOldPPDDataSet.Tables.Remove("DataRow")
    If CMTNewPPDDataSet.Tables.Contains("DataRow") Then CMTNewPPDDataSet.Tables.Remove("DataRow")
    IncentiveDataSet = Nothing
    FulfilIncentives = False
    'Payment plan resets
    PPDLines = 0
    mvAppStartPoint = TraderApplicationStartPoint.taspTransaction
    If mvChangeMembershipType Then mvAppStartPoint = TraderApplicationStartPoint.taspApplication
    mvChangeBranchWithAddress = Nothing
    mvCreateContactAccount = Nothing
    mvCreateCommLink = Nothing

    mvPPNumbers = New CollectionList(Of Integer)
    PaymentPlans = New CollectionList(Of PaymentPlanInfo)
    Provisional = CAFCard OrElse Voucher OrElse GiftInKind OrElse SaleOrReturn OrElse IncludeProvisionalTransactions
    BatchNumbers = ""
    If MemberNumber Is Nothing Then MemberNumber = ""
    mvAutoPaymentCreated = False
    DeclarationNumber = 0
    PPIncentivesCompleted = False
    mvLastDeceasedContactNumber = 0
    PPDMemberOrPayer = ""
    CSDepositAmount = 0
    CMTOriginalMemberJoined = ""
    mvDeliveryContactNumber = 0
    mvDeliveryAddressNumber = 0
    LoanAmount = 0
    mvPPDetailsPricing = Nothing
    CMTPrevMembershipTypeCode = ""
    AddIncentivesLinesRequired = False
    TransactionDonationAmount = 0
  End Sub

  Public Sub SetLineTotal()
    TransactionLines = 0
    If AnalysisDataSet.Tables.Contains("DataRow") Then
      Dim vAmount As Double
      Dim vNoDeposit As Boolean = False
      Dim vMaxLineNumber As Integer = 0
      Dim vDT As DataTable = AnalysisDataSet.Tables("DataRow")
      For Each vRow As DataRow In vDT.Rows
        Select Case vRow.Item("TraderLineType").ToString
          Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "M", "N", "O", "PG", "R", "S", "U", "VC", "VE", "AP", "X", "Q"
            vAmount += DoubleValue(vRow.Item("Amount").ToString)
          Case "P", "V"
            vAmount += DoubleValue(vRow.Item("Amount").ToString)
            If CSDepositPercentage > 0 AndAlso Not vNoDeposit Then
              'If CSDepositPercentage set and the TraderLineType is either ProductSale or Service Booking then use the
              'TALine Deposit Allowed flag to calculate the CSDepositAmount
              vNoDeposit = vRow.Item("DepositAllowed").ToString = "N"
            End If
        End Select
        TransactionLines += 1
        If vDT.Columns.Contains("LineNumber") Then
          If IntegerValue(vRow.Item("LineNumber").ToString) > vMaxLineNumber Then vMaxLineNumber = IntegerValue(vRow.Item("LineNumber").ToString)
        End If
      Next
      If vMaxLineNumber > TransactionLines Then TransactionLines = vMaxLineNumber
      If TransactionLines > 0 Then TransactionLines += mvAnalysisLinesDeleted
      CurrentLineTotal = FixTwoPlaces(vAmount)
      If CSDepositPercentage > 0 Then
        'calculate and display the invoice deposit amount
        Dim vDepositPercentage As Double = 1
        If Not vNoDeposit Then vDepositPercentage = CSDepositPercentage / 100
        CSDepositAmount = FixTwoPlaces(vAmount * vDepositPercentage)
      End If
    End If
  End Sub

  Public Sub SaveApplicationValues(ByVal pPage As TraderPage, ByVal pList As ParameterList)
    Select Case pPage.PageType
      Case CareServices.TraderPageType.tpActivityEntry, CareServices.TraderPageType.tpSetStatus, CareServices.TraderPageType.tpSuppressionEntry, CareServices.TraderPageType.tpGiftAidDeclaration, _
      CareServices.TraderPageType.tpGoneAway, CareServices.TraderPageType.tpAddressMaintenance, CareServices.TraderPageType.tpGiveAsYouEarnEntry
        If pList.ContainsKey("ContactNumber") AndAlso pList.ContainsKey("AddressNumber") Then
          If PayerContactNumber = 0 OrElse PayerAddressNumber = 0 Then SetPayerContact(pList.IntegerValue("ContactNumber"), pList.IntegerValue("AddressNumber"))
        Else
          Dim vContactInfo As ContactInfo = pPage.EditPanel.FindTextLookupBox("ContactNumber").ContactInfo
          If vContactInfo IsNot Nothing Then
            If PayerContactNumber = 0 OrElse PayerAddressNumber = 0 Then SetPayerContact(vContactInfo.ContactNumber, vContactInfo.AddressNumber)
            pList.IntegerValue("AddressNumber") = vContactInfo.AddressNumber
          End If
        End If
      Case CareServices.TraderPageType.tpCardDetails
        If pList.ContainsKey("Reference") Then TransactionReference = pList("Reference")
      Case CareServices.TraderPageType.tpCollectionPayments
        pList("VatRate") = LineVATRate
        pList("VatPercentage") = LineVATPercentage.ToString
        pList("VatAmount") = LineVATAmount.ToString
      Case CareServices.TraderPageType.tpCreditCustomer
        If NewCreditCustomer = False Then
          If CreditCustomerAddressChanged Then SaveCreditCustomerAddressChange = ShowQuestion(QuestionMessages.QmCreditCustomerAddressChanged, MessageBoxButtons.YesNo) = DialogResult.Yes
        End If
        If pList.ContainsKey("Reference") Then TransactionReference = pList("Reference")
      Case CareServices.TraderPageType.tpEventBooking, CareServices.TraderPageType.tpProductDetails, CareServices.TraderPageType.tpAccommodationBooking, CareNetServices.TraderPageType.tpAmendEventBooking
        pList("VatRate") = LineVATRate    'mvVATRate
        pList("VatPercentage") = LineVATPercentage.ToString   'mvVATPercentage.ToString
        pList("VatAmount") = LineVATAmount.ToString
        pList("ContactDiscount") = CBoolYN(PayerHasDiscount)
        If ShowVATExclusiveAmount AndAlso pPage.PageType = CareNetServices.TraderPageType.tpProductDetails Then      'BR20047 (only applicable to Product Sales)If we were showing the VAT exclusive amount then add the VAT back in
          pList("Amount") = FixTwoPlaces(DoubleValue(pList("Amount")) + LineVATAmount).ToString("0.00")
        End If
        Select Case pPage.PageType
          Case CareServices.TraderPageType.tpEventBooking
            If pList.ContainsKey("EventNumber") Then EventNumber = pList("EventNumber")
            If pList.ContainsKey("OptionNumber") Then BookingOptionNumber = pList("OptionNumber")
            If pList.ContainsKey("Rate") Then EventBookingRate = pList("Rate")
          Case CareServices.TraderPageType.tpProductDetails
            If pList.Contains("DeceasedContactNumber") Then mvLastDeceasedContactNumber = pList.IntegerValue("DeceasedContactNumber")
            If pList.Contains("ContactNumber") Then SetDeliveryContactAndAddress(pList.IntegerValue("ContactNumber"), pList.IntegerValue("AddressNumber"))
        End Select
      Case CareServices.TraderPageType.tpMembership
        If pList.Contains("MembershipType") AndAlso pList.Contains("Rate") AndAlso pList.Contains("Joined") AndAlso pList.Contains("Source") _
          AndAlso pList.Contains("NumberOfMembers") AndAlso pList.Contains("MaxFreeAssociates") AndAlso pList.Contains("Amount") Then
          'J1493: Validate that ParameterList contains the expected items as when coming in from CancelTransaction TraderProcessDataType some parameter items are missing
          SetMembershipValues(pList("MembershipType"), pList("Rate"), pList("Joined"), pList("Source"), IntegerValue(pList("NumberOfMembers")), IntegerValue(pList("MaxFreeAssociates")), pList("Amount"))
        End If
        If pList.ContainsKey("Source") Then TransactionSource = pList("Source")
      Case CareServices.TraderPageType.tpPaymentPlanMaintenance
        If pList.ContainsKey("PaymentFrequency") Then mvNewPaymentFrequency = pList("PaymentFrequency")
      Case CareServices.TraderPageType.tpTransactionDetails, CareServices.TraderPageType.tpPaymentPlanDetails, CareServices.TraderPageType.tpContactSelection
        If pList.ContainsKey("TransactionDate") Then TransactionDate = pList("TransactionDate")
        If pList.ContainsKey("Source") Then TransactionSource = pList("Source")
        If pList.ContainsKey("ContactNumber") AndAlso pList.ContainsKey("AddressNumber") Then SetPayerContact(pList.IntegerValue("ContactNumber"), pList.IntegerValue("AddressNumber"))
        If pList.ContainsKey("SalesContactNumber") Then SalesContactNumber = pList.IntegerValue("SalesContactNumber")
        If pPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails Then
          If pList.Contains("Balance") Then PPBalance = DoubleValue(pList("Balance").ToString) Else PPBalance = 0
          PPAmount = pList("Amount")
        Else
          If pList.ContainsKey("Amount") Then TransactionAmount = DoubleValue(pList("Amount"))
        End If
        If pList.ContainsKey("MemberNumber") Then MemberNumber = pList("MemberNumber")
        If pList.ContainsKey("CovenantNumber") Then CovenantNumber = IntegerValue(pList("CovenantNumber"))
        'If pList.ContainsKey("PaymentPlanNumber") Then mvPPNumber = IntegerValue(pList("PaymentPlanNumber"))
        If pList.ContainsKey("Reference") Then TransactionReference = pList("Reference")
        If pList.ContainsKey("OrderDate") Then
          PPRenewalPeriodStartDate = pList("OrderDate")
        ElseIf PaymentPlan IsNot Nothing Then
          'on amend set start date from paymen plan info
          If ApplicationType = ApplicationTypes.atMaintenance OrElse ApplicationType = ApplicationTypes.atConversion Then
            PPRenewalPeriodStartDate = PaymentPlan.CurrentTermStartDate().ToString(AppValues.DateFormat)
          Else
            PPRenewalPeriodStartDate = PaymentPlan.CalculateRenewalDate(PaymentPlan.RenewalPeriodEnd, False).ToString(AppValues.DateFormat)
          End If
        End If
        If pPage.PageType = CareServices.TraderPageType.tpTransactionDetails Then
          If pList.ContainsKey("Notes") Then TransactionNote = pList("Notes")
        End If
      Case CareServices.TraderPageType.tpPurchaseInvoiceDetails
        If pList.ContainsKey("PayeeContactNumber") Then SetPayerContact(pList.IntegerValue("PayeeContactNumber"), pList.IntegerValue("PayeeAddressNumber"))
        If pList.ContainsKey("Amount") Then PPBalance = DoubleValue(pList("Amount").ToString)
        TransactionDate = pList("PurchaseInvoiceDate")
      Case CareServices.TraderPageType.tpPurchaseOrderDetails
        If pList.ContainsKey("PayeeContactNumber") Then SetPayerContact(pList.IntegerValue("PayeeContactNumber"), pList.IntegerValue("PayeeAddressNumber"))
        If pList.ContainsKey("NumberOfPayments") Then PONumberOfPayments = pList.IntegerValue("NumberOfPayments")
        If pList.ContainsKey("Amount") Then PPBalance = DoubleValue(pList("Amount").ToString)
        POPercentage = pList("PaymentAsPercentage") = "Y"
        TransactionDate = pList("StartDate")
      Case CareServices.TraderPageType.tpPurchaseOrderProducts
        AppValues.LastDistributionCode = pList("DistributionCode")
      Case CareServices.TraderPageType.tpConfirmProvisionalTransactions
        If ConfirmTransList.Length > 0 Then ConfirmTransList.Append(",")
        ConfirmTransList.Append(pPage.EditPanel.GetValue("ProvisionalBatchNumber"))
        ConfirmTransList.Append(",")
        ConfirmTransList.Append(pPage.EditPanel.GetValue("ProvisionalTransNumber"))
      Case CareNetServices.TraderPageType.tpPostageAndPacking
        pList("VatRate") = LineVATRate
        pList("VatPercentage") = LineVATPercentage.ToString
        pList("VatAmount") = LineVATAmount.ToString
        pList("PriceVATExclusive") = LinePriceVATEx.ToString
      Case CareNetServices.TraderPageType.tpStatementList
        If ApplicationType = ApplicationTypes.atCreditListReconciliation Then
          Dim vGRD As DisplayGrid = CType(pPage.EditPanel.FindPanelControl("StatementDisplayGrid"), DisplayGrid)
          If vGRD IsNot Nothing AndAlso vGRD.RowCount > 0 Then
            TransactionAmount = DoubleValue(vGRD.GetValue(vGRD.CurrentRow, "Amount"))
            TransactionDate = StatementDate
            BankPaymentMethod = vGRD.GetValue(vGRD.CurrentRow, "PaymentMethod")
            TransactionNote = vGRD.GetValue(vGRD.CurrentRow, "Notes")
          End If
        End If
      Case CareNetServices.TraderPageType.tpTransactionAnalysisSummary
        If ApplicationType = ApplicationTypes.atTransaction And TransactionNote.Length > 0 Then
          pList("Notes") = TransactionNote
        End If
      Case CareNetServices.TraderPageType.tpComments
        If pList.ContainsKey("Notes") Then TransactionNote = pList("Notes")
    End Select
    If pList.ContainsKey("CheckIncentives") Then CheckIncentives = BooleanValue(pList("CheckIncentives"))
  End Sub

  Public Sub GetApplicationValues(ByVal pList As ParameterList)
    pList.IntegerValue("TransactionLines") = TransactionLines
    If BatchDate.Length > 0 Then pList("BatchDate") = BatchDate
    If PostToCashBook.Length > 0 Then pList("PostToCashBook") = PostToCashBook
    If TransactionDate.Length > 0 Then pList("TransactionDate") = TransactionDate
    If TransactionDateChanged Then pList("TransactionDateChanged") = "Y"
    If TransactionSource.Length > 0 Then pList("TransactionSource") = TransactionSource
    If EventNumber.Length > 0 Then pList("EventNumber") = EventNumber
    If BookingOptionNumber.Length > 0 Then pList("BookingOptionNumber") = BookingOptionNumber
    If EventBookingRate.Length > 0 Then pList("EventBookingRate") = EventBookingRate
    If TransactionDistributionCode.Length > 0 Then pList("TransactionDistributionCode") = TransactionDistributionCode
    If PayerContactNumber > 0 Then pList.IntegerValue("PayerContactNumber") = PayerContactNumber
    If PayerAddressNumber > 0 Then pList.IntegerValue("PayerAddressNumber") = PayerAddressNumber
    If SalesContactNumber > 0 Then pList.IntegerValue("SalesContactNumber") = SalesContactNumber

    If TransactionPaymentMethod.Length > 0 Then pList("TransactionPaymentMethod") = TransactionPaymentMethod 'Set by button clicked on PM1
    If TransactionType.Length > 0 Then pList("TransactionType") = TransactionType 'Set by button clicked on Transaction Analysis
    If UnbalancedTransactionChoice.Length > 0 Then pList("UnbalancedTransactionChoice") = UnbalancedTransactionChoice
    If MultiCurrency() AndAlso TransactionType = "PAYM" Then
      If pList.Contains("Amount") Then pList("Amount") = CalcCurrencyAmount(CDbl(pList("Amount")), True).ToString
    End If
    pList("TransactionAmount") = TransactionAmount.ToString

    pList("DetailLineTotal") = CurrentLineTotal.ToString
    If CSDepositPercentage > 0 AndAlso CSDepositAmount > 0 Then pList("CSDepositAmount") = CSDepositAmount.ToString
    If PPPaymentType.Length > 0 Then pList("PPPaymentType") = PPPaymentType
    If TransactionReference.Length > 0 Then pList("TransactionReference") = TransactionReference
    'If CurrentPaymentMethod.Length > 0 Then 
    pList("CurrentPaymentMethod") = CBoolYN(CurrentPaymentMethod)
    If PPDMemberOrPayer.Length > 0 Then pList("MemberOrPayer") = PPDMemberOrPayer

    If BatchNumber > 0 Then
      pList.IntegerValue("BatchNumber") = BatchNumber
      pList.IntegerValue("TransactionNumber") = TransactionNumber
      If TransactionNumber > 0 Then
        pList("ExistingTransaction") = CBoolYN(EditExistingTransaction)
        pList("OriginalTransactionAmount") = OriginalTransactionAmount.ToString
        pList("OriginalTransactionCurrencyAmount") = OriginalTransactionCurrencyAmount.ToString
        If BatchNumbers.Length > 0 Then pList("BatchNumbers") = BatchNumbers
      End If
      If BatchInfo.BatchNumber > 0 Then pList("Provisional") = CBoolYN(BatchInfo.Provisional)
    End If
    If BatchInfo.BatchNumber = 0 AndAlso Provisional Then pList("Provisional") = "Y"

    'Payment Plan Specific parameters - would like to move them to a more specific place...mabe a separate method or may be into the frmtrader
    pList("PPBalance") = PPBalance.ToString
    pList("PPDLines") = PPDLines.ToString
    pList("FixedAmount") = PPAmount
    If TransactionType = "MEMB" OrElse TransactionType = "MEMC" Then
      If EditMemberLineNumber > 0 Then
        pList("EditMemberLineNumber") = EditMemberLineNumber.ToString
        'If TransactionType = "MEMC" Then pList("MembershipNumber") = MembersDataSet.Tables(0).Rows(EditMemberLineNumber).Item("MembershipNumber").ToString
      End If
      pList.IntegerValue("CurrentMembers") = CurrentMembers
    End If

    If TransactionType = "PAYM" Then
      If Not pList.Contains("MemberNumber") Then pList("MemberNumber") = MemberNumber
      If Not pList.Contains("CovenantNumber") Then pList.IntegerValue("CovenantNumber") = CovenantNumber
      If (Not pList.Contains("PaymentPlanNumber") And PaymentPlan IsNot Nothing) Or (pList.Contains("PaymentPlanNumber") And PaymentPlan IsNot Nothing AndAlso pList.IntegerValue("PaymentPlanNumber") = 0) Then
        pList.IntegerValue("PaymentPlanNumber") = PaymentPlan.PaymentPlanNumber
      End If
    End If

    If TransactionType = "MEMC" Then
      If mvAppStartPoint = TraderApplicationStartPoint.taspTransaction Then pList("CreateTransaction") = "Y"
      If CMTPriceDate.Length > 0 Then pList("CMTPriceDate") = CMTPriceDate
      If CMTUpdatePPFixedAmount.HasValue Then pList("UpdatePPFixedAmount") = CBoolYN(CMTUpdatePPFixedAmount.Value)
    End If

    If mvChangeBranchWithAddress IsNot Nothing Then
      pList("ChangeBranchWithAddress") = mvChangeBranchWithAddress
    End If

    If mvCreateCommLink IsNot Nothing Then
      pList("CreateCommLink") = mvCreateCommLink
    End If

    If mvCreateContactAccount IsNot Nothing Then
      pList("CreateAccount") = mvCreateContactAccount
    End If

    If mvStockSales Then
      pList("StockSale") = "Y"
      pList.IntegerValue("StockIssued") = mvStockIssued
      pList.IntegerValue("StockTransactionID") = mvStockTransactionID
    End If

    If (ApplicationType = ApplicationTypes.atConversion AndAlso PayPlanConvMaintenance = True) OrElse ApplicationType = ApplicationTypes.atMaintenance Then
      If mvNewPaymentFrequency.Length > 0 Then pList("NewPaymentFrequency") = mvNewPaymentFrequency
      If mvEditPPDetailNumber > 0 Then pList("DetailNumber") = mvEditPPDetailNumber.ToString
      If mvEditPPDSubscriptionNumber > 0 Then pList("SubscriptionNumber") = mvEditPPDSubscriptionNumber.ToString
      If CancellationReason.Length > 0 Then pList("CancellationReason") = CancellationReason
      If CancellationSource.Length > 0 Then pList("CancellationSource") = CancellationSource
      If PaymentPlan IsNot Nothing AndAlso PaymentPlan.Existing Then
        If PaymentPlan.PlanType = PaymentPlanInfo.ppType.pptLoan AndAlso BooleanValue(PaymentPlan.LoanStatus) = True Then pList("LoanPaymentPlan") = "Y"
      End If
    End If
    If (PurchaseOrderType = PurchaseOrderTypes.PaymentSchedule OrElse PurchaseOrderType = PurchaseOrderTypes.RegularPayments) AndAlso PONumberOfPayments > 0 Then pList.IntegerValue("NumberOfPayments") = PONumberOfPayments
    If PurchaseOrderType = PurchaseOrderTypes.RegularPayments AndAlso OldPORegularPaymentAmount.HasValue Then pList("OldPORegularPaymentAmount") = OldPORegularPaymentAmount.Value.ToString
    If PurchaseOrderNumber > 0 Then pList.IntegerValue("PurchaseOrderNumber") = PurchaseOrderNumber
    If PurchaseInvoiceNumber > 0 Then pList.IntegerValue("PurchaseInvoiceNumber") = PurchaseInvoiceNumber
    If FinancialAdjustment <> CDBNETCL.BatchInfo.AdjustmentTypes.None Then
      pList.IntegerValue("FinancialAdjustment") = FinancialAdjustment
      If FATransactionType.Length > 0 Then pList("FATransactionType") = FATransactionType
    End If
    If ConfirmTransList.Length > 0 Then pList("ConfirmTransList") = ConfirmTransList.ToString
    If CheckIncentives Then pList("CheckIncentives") = "Y"
    If FulfilIncentives Then pList("FulfilIncentives") = "Y"
    If AllocationsChecked Then pList("AllocationsChecked") = "Y"
    If UnallocateCreditNote Then pList("UnallocateCreditNote") = "Y"
  End Sub

  Public Sub DeleteAnalysisLine(ByVal pRowNumber As Integer)
    DeleteAnalysisLine(pRowNumber, False, 0, 0)
  End Sub
  Public Sub DeleteAnalysisLine(ByVal pRowNumber As Integer, ByVal pIncentiveLinesOnly As Boolean)
    DeleteAnalysisLine(pRowNumber, pIncentiveLinesOnly, 0, 0)
  End Sub
  Public Sub DeleteAnalysisLine(ByVal pRowNumber As Integer, ByVal pEventBookingNumber As Integer, ByVal pExamBookingNumber As Integer)
    DeleteAnalysisLine(pRowNumber, False, pEventBookingNumber, pExamBookingNumber)
  End Sub
  Public Sub DeleteAnalysisLine(ByVal pRowNumber As Integer, ByVal pIncentiveLinesOnly As Boolean, ByVal pEventBookingNumber As Integer, ByVal pExamBookingNumber As Integer)
    'Deleting row and update count of number of lines deleted (this ensures adding new line uses correct line number)
    If AnalysisDataSet.Tables.Contains("DataRow") Then
      Dim vRowDeleted As Boolean = False
      Dim vAnalysisTable As DataTable = AnalysisDataSet.Tables("DataRow")

      'Delete any incentive lines linked to current row
      Dim vIncentivesLines As New StringBuilder
      Dim vIndex As Integer
      For Each vRow As DataRow In vAnalysisTable.Rows
        If vRow("TraderLineType").ToString = "I" Then
          If vRow("IncentiveLineNumber").ToString = vAnalysisTable.Rows(pRowNumber)("LineNumber").ToString Then
            If vIncentivesLines.Length > 0 Then vIncentivesLines.Insert(0, ",")
            vIncentivesLines.Insert(0, vIndex)
          End If
        End If
        vIndex += 1
      Next
      If vIncentivesLines.Length > 0 Then
        For Each vLineNo As String In vIncentivesLines.ToString.Split(","c)
          vAnalysisTable.Rows.RemoveAt(IntegerValue(vLineNo))
          mvAnalysisLinesDeleted += 1
        Next
      End If

      If pEventBookingNumber > 0 AndAlso vAnalysisTable.Rows.Count > 0 Then
        'If Event Booking used the Pricing Matrix then there will be other lines that we need to delete as well
        'So delete these lines from last to first
        Dim vEPMLines As New StringBuilder
        For vRowIndex As Integer = (vAnalysisTable.Rows.Count - 1) To pRowNumber Step -1
          Dim vRow As DataRow = vAnalysisTable.Rows(vRowIndex)
          If vRow("TraderLineType").ToString = "X" AndAlso IntegerValue(vRow("EventBookingNumber").ToString) = pEventBookingNumber Then
            If vEPMLines.ToString.Length > 0 Then vEPMLines.Append(",")
            vEPMLines.Append(vRowIndex.ToString)
          End If
        Next
        If vEPMLines.ToString.Length > 0 Then
          For Each vLineNo As String In vEPMLines.ToString.Split(","c)
            vAnalysisTable.Rows.RemoveAt(IntegerValue(vLineNo))
            mvAnalysisLinesDeleted += 1
          Next
        End If
      End If

      If pExamBookingNumber > 0 AndAlso vAnalysisTable.Rows.Count > 0 Then
        'An Exam Booking may hav other lines that we need to delete as well
        'So delete these lines from last to first
        Dim vEPMLines As New StringBuilder
        For vRowIndex As Integer = (vAnalysisTable.Rows.Count - 1) To 0 Step -1     'Look at all lines since they may have deleted a line after the first
          Dim vRow As DataRow = vAnalysisTable.Rows(vRowIndex)
          If vRow("TraderLineType").ToString = "Q" AndAlso IntegerValue(vRow("ExamBookingId").ToString) = pExamBookingNumber Then
            If vEPMLines.ToString.Length > 0 Then vEPMLines.Append(",")
            vEPMLines.Append(vRowIndex.ToString)
          End If
        Next
        If vEPMLines.ToString.Length > 0 Then
          For Each vLineNo As String In vEPMLines.ToString.Split(","c)
            vAnalysisTable.Rows.RemoveAt(IntegerValue(vLineNo))
            mvAnalysisLinesDeleted += 1
            vRowDeleted = True
          Next
        End If
      ElseIf vAnalysisTable.Rows.Count > 0 Then
        'Delete any Sales Ledger Cash Allocation (L) / Sundry Credit Note Invoice Allocation (RL) lines linked to current L/RL-type row (from last to first)
        Select Case vAnalysisTable.Rows(pRowNumber)("TraderLineType").ToString
          Case "L", "K" 'L Sales Ledger Cash Allocation, K Sundry Credit Note Invoice Allocation
            Dim vTraderLineType As String = vAnalysisTable.Rows(pRowNumber)("TraderLineType").ToString
            Dim vInvoiceNumberString As String = vAnalysisTable.Rows(pRowNumber)("InvoiceNumber").ToString
            Dim vAllocationLines As New StringBuilder
            For vRowIndex As Integer = (vAnalysisTable.Rows.Count - 1) To 0 Step -1
              Dim vRow As DataRow = vAnalysisTable.Rows(vRowIndex)
              If vRow("TraderLineType").ToString = vTraderLineType Then
                If vRow("LineNumber").ToString <> vAnalysisTable.Rows(pRowNumber)("LineNumber").ToString AndAlso vRow("InvoiceNumber").ToString = vInvoiceNumberString Then
                  If vAllocationLines.Length > 0 Then vAllocationLines.Append(",")
                  vAllocationLines.Append(vRowIndex)
                End If
              End If
            Next
            If vAllocationLines.Length > 0 Then
              For Each vLineNo As String In vAllocationLines.ToString.Split(","c)
                If IntegerValue(vLineNo) < pRowNumber AndAlso vRowDeleted = False Then
                  'Need to delete pRowNumber first to prevent errors
                  vAnalysisTable.Rows.RemoveAt(pRowNumber)
                  mvAnalysisLinesDeleted += 1
                  vRowDeleted = True
                End If
                vAnalysisTable.Rows.RemoveAt(IntegerValue(vLineNo))
                mvAnalysisLinesDeleted += 1
              Next
            End If
        End Select
      End If

      If pIncentiveLinesOnly = False AndAlso vRowDeleted = False Then
        vAnalysisTable.Rows.RemoveAt(pRowNumber)
        mvAnalysisLinesDeleted += 1
      End If
      SetLineTotal()
    End If
  End Sub

  Public Sub SetPPDLineTotal()
    PPDLines = 0
    mvPPDProductNumbersCount = 0
    CurrentPPDAmount = 0
    CurrentPPDArrears = 0
    Dim vDiscount As Double
    Dim vNonPerBalance As Double
    Dim vPercentage As Integer
    Dim vValidFrom As String = String.Empty
    Dim vValidTo As String = String.Empty
    Dim vValid As Boolean
    If PPDDataSet.Tables.Contains("DataRow") Then
      Dim vBalance As Double
      For Each vRow As DataRow In PPDDataSet.Tables("DataRow").Rows
        vValid = True
        vValidFrom = vRow("ValidFrom").ToString
        vValidTo = vRow("ValidTo").ToString

        'NFPCARE-559
        'Check if the payment plan start date falls within the valid dates for the line item
        If vValidFrom.Length > 0 AndAlso PPRenewalPeriodStartDate.Length > 0 AndAlso CDate(PPRenewalPeriodStartDate) < CDate(vValidFrom) Then vValid = False
        If vValid AndAlso (vValidTo.Length > 0 AndAlso PPRenewalPeriodStartDate.Length > 0 AndAlso CDate(PPRenewalPeriodStartDate) > CDate(vValidTo)) Then vValid = False

        If vValid Then
          If vRow.Item("PriceIsPercentage").ToString = "N" OrElse vRow.Item("PriceIsPercentage").ToString.Length = 0 Then
            vBalance += DoubleValue(vRow.Item("Balance").ToString)
            vNonPerBalance += DoubleValue(vRow.Item("Balance").ToString)
            CurrentPPDAmount += DoubleValue(vRow.Item("Amount").ToString)
            CurrentPPDArrears += DoubleValue(vRow.Item("Arrears").ToString)
          Else
            If vRow("DiscountPercentage").ToString.Length = 0 Then
              'Fetch the rate from the rates table
              Dim vParams As New ParameterList(True)
              vParams("Product") = vRow("Product").ToString
              vParams("Rate") = vRow("Rate").ToString
              Dim vRate As ParameterList = DataHelper.GetLookupItem(CareNetServices.XMLLookupDataTypes.xldtRates, vParams)
              vRow("DiscountPercentage") = vRate("CurrentPrice")
              'Use the createdOn date for the rate so that we using the correct rate while amending a PP
              If vRate("PriceChangeDate").Length > 0 Then
                If vRow("CreatedOn").ToString.Length > 0 AndAlso CDate(vRow("CreatedOn")) >= CDate(vRate("PriceChangeDate")) Then
                  vRow("DiscountPercentage") = vRate("FuturePrice")
                End If
              End If
            End If
            vPercentage = IntegerValue(vRow("DiscountPercentage"))

            If vPercentage > 0 Then vPercentage = vPercentage * -1
            If vRow("PriceIsPercentage").ToString = "T" Then
              'calculate discount on the previous non-percentage total
              vDiscount = FixTwoPlaces(vNonPerBalance * (vPercentage / 100))
              vBalance += vDiscount
            Else
              'calculate discount on the previous total
              vDiscount = FixTwoPlaces(vBalance * (vPercentage / 100))
              vBalance += vDiscount
            End If
            vRow("Balance") = vDiscount
          End If
        Else
          vRow("Balance") = "0.00"
        End If

        If BooleanValue(vRow("UsesProductNumbers").ToString) Then mvPPDProductNumbersCount += 1
        PPDLines += 1

      Next
      'PPDLines += mvAnalysisLinesDeleted
      CurrentPPDLineTotal = vBalance
    End If
  End Sub

  Public Sub DeleteDataSetLine(ByVal pDataSet As DataSet, ByVal pLineNumber As Integer)
    'Deleting row and update count of number of lines deleted (this ensures adding new line uses correct line number)
    If pDataSet.Tables.Contains("DataRow") Then
      pDataSet.Tables("DataRow").Rows.Remove(GetDataSetLine(pDataSet, pLineNumber))
      Dim vLine As Integer = 1
      For Each vDR As DataRow In pDataSet.Tables("DataRow").Rows
        vDR.Item("LineNumber") = vLine
        vLine = vLine + 1
      Next
      If pDataSet.DataSetName = PPDDataSet.DataSetName AndAlso PPDDataSet.Tables.Contains("DataRow") Then
        SetPPDLineTotal()
      Else
        SetDataSetLineTotal(pDataSet)
      End If
    End If
  End Sub

  Public Sub SetDataSetLineTotal(ByVal pDataSet As DataSet)
    PPDLines = 0
    CurrentPPDLineTotal = 0
    If pDataSet.Tables.Contains("DataRow") Then
      For Each vRow As DataRow In pDataSet.Tables("DataRow").Rows
        CurrentPPDLineTotal += DoubleValue(vRow.Item("Amount").ToString)
        PPDLines += 1
      Next
    End If
  End Sub

  Private Sub SetMembershipValues(ByVal pMembershipType As String, ByVal pMembershipRate As String, ByVal pJoined As String, ByVal pSource As String, ByVal pNumberMembers As Integer, ByVal pNumberAssociates As Integer, ByVal pFixedAmount As String)
    mvLastMembershipType = pMembershipType
    mvLastMembershipRate = pMembershipRate
    mvLastMembershipSource = pSource
    mvLastMembershipJoinedDate = pJoined
    mvLastMembershipNumberMembers = pNumberMembers
    mvLastMembershipNumberAssociates = pNumberAssociates
    mvLastMembershipFixedAmount = pFixedAmount
  End Sub

  Friend Sub SetCMTValues(ByVal pCMTMemberContactNumber As Integer, ByVal pCMTMemberAddressNumber As Integer) ', ByVal pCMTMembershipNumber As Integer)
    mvCMTMemberContactNumber = pCMTMemberContactNumber
    mvCMTMemberAddressNumber = pCMTMemberAddressNumber
    'mvCMTMembershipNumber = pCMTMembershipNumber
  End Sub

  Public Sub DeleteMember(ByVal pRow As Integer)
    'Delete row and update count of number of members
    If MembersDataSet.Tables.Contains("DataRow") Then
      MembersDataSet.Tables("DataRow").Rows.RemoveAt(pRow)
    End If
  End Sub

  Public ReadOnly Property CurrentMembers As Integer
    Get
      Dim vResult As Integer = 0
      If MembersDataSet IsNot Nothing AndAlso
        MembersDataSet.Tables IsNot Nothing AndAlso
        MembersDataSet.Tables.Contains("DataRow") Then
        vResult = MembersDataSet.Tables("DataRow").Rows.Count
      End If
      Return vResult
    End Get
  End Property

  Private Function GetApplicationType(ByVal pCode As String) As ApplicationTypes
    Select Case pCode
      Case "TRANS"
        GetApplicationType = ApplicationTypes.atTransaction
      Case "PINVE"
        GetApplicationType = ApplicationTypes.atPurchaseInvoice
      Case "PORDE"
        GetApplicationType = ApplicationTypes.atPurchaseOrder
      Case "PORDC"
        GetApplicationType = ApplicationTypes.atPurchaseOrderCancellation
      Case "CHQNA"
        GetApplicationType = ApplicationTypes.atChequeNumberAllocation
      Case "CHQRE"
        GetApplicationType = ApplicationTypes.atChequeReconciliation
      Case "CSTAT"
        GetApplicationType = ApplicationTypes.atCreditStatementGeneration
      Case "BINVG"
        GetApplicationType = ApplicationTypes.atBatchInvoiceGeneration
      Case "MAINT"
        GetApplicationType = ApplicationTypes.atMaintenance
      Case "CNVRT"
        GetApplicationType = ApplicationTypes.atConversion
      Case "CLREC"
        GetApplicationType = ApplicationTypes.atCreditListReconciliation
      Case "BSPOS"
        GetApplicationType = ApplicationTypes.atBankStatementPosting
      Case "POGEN"
        GetApplicationType = ApplicationTypes.atPurchaseOrderGeneration
      Case "POPRT"
        GetApplicationType = ApplicationTypes.atPurchaseOrderPrint
      Case "POCHQ"
        GetApplicationType = ApplicationTypes.atChequeProcessing
      Case "GAYEP"
        GetApplicationType = ApplicationTypes.atGiveAsYouEarnPayments
      Case "POTPG"
        GetApplicationType = ApplicationTypes.atPostTaxPGPayments
    End Select
  End Function

  Public Function GetDataSetLine(ByVal pDataSet As DataSet, ByVal pLineNumber As Integer) As DataRow
    Dim vFound As Boolean = False
    Dim vRow As DataRow = Nothing

    Dim vTable As DataTable = pDataSet.Tables("DataRow")
    For Each vRow In vTable.Rows
      If IntegerValue(vRow("LineNumber").ToString) = pLineNumber Then
        vFound = True
        Exit For
      End If
    Next
    If vFound Then Return vRow Else Return Nothing
  End Function

  Friend Sub SetPopupMenuDetails(ByVal pList As ParameterList)
    If ApplicationType = ApplicationTypes.atConversion Or ApplicationType = ApplicationTypes.atMaintenance Then
      PaymentPlan = New PaymentPlanInfo(pList.IntegerValue("PaymentPlanNumber"))
    ElseIf ChangeMembershipType Then
      mvCMTMemberNumber = pList("MemberNumber").ToString
      mvMembershipNumber = pList.IntegerValue("MembershipNumber")
    ElseIf ApplicationType = ApplicationTypes.atPurchaseOrder Then
      PurchaseOrderNumber = pList.IntegerValue("PurchaseOrderNumber")
    End If
    mvAppStartPoint = TraderApplicationStartPoint.taspRightMouse
  End Sub


  Public Property ChangeBranchWithAddress() As String
    Get
      Return mvChangeBranchWithAddress
    End Get
    Set(ByVal pValue As String)
      mvChangeBranchWithAddress = pValue
    End Set
  End Property

  Public Property CreateCommLink() As String
    Get
      Return mvCreateCommLink
    End Get
    Set(ByVal pValue As String)
      mvCreateCommLink = pValue
    End Set
  End Property

  Public Property CreateContactAccount() As String
    Get
      Return mvCreateContactAccount
    End Get
    Set(ByVal pValue As String)
      mvCreateContactAccount = pValue
    End Set
  End Property

  'Public Function CalcCurrencyAmount(ByVal pAmount As Double, ByVal pAsBaseValue As Boolean) As Double
  '  Dim vAmount As Double

  '  If MultiCurrency() And Len(BatchCurrencyCode) > 0 Then
  '    If pAsBaseValue Then
  '      vAmount = Round(FixTwoPlaces(pAmount) / mvTraderApplication.BatchExchangeRate, 2)
  '    Else
  '      vAmount = Round(FixTwoPlaces(pAmount) * mvTraderApplication.BatchExchangeRate, 2)
  '    End If
  '  Else
  '    vAmount = pAmount
  '  End If
  '  CalcCurrencyAmount = vAmount
  'End Function
  Public ReadOnly Property TraderPage(ByVal pEPL As EditPanel) As TraderPage
    Get
      For Each vPage As TraderPage In Pages
        If vPage.EditPanel Is pEPL Then
          Return vPage
          Exit For
        End If
      Next
      Return Nothing
    End Get
  End Property

  Public ReadOnly Property Memberships() As Boolean
    Get
      Return mvMemberships
    End Get
  End Property

  Public Property MembershipNumber() As Integer
    Get
      Return mvMembershipNumber
    End Get
    Set(ByVal pValue As Integer)
      mvMembershipNumber = pValue
    End Set
  End Property

  Friend ReadOnly Property ApplicationStartPoint() As TraderApplicationStartPoint
    Get
      Return mvAppStartPoint
    End Get
  End Property

  Public Function CreatesTransaction() As Boolean
    Dim vPayPlansOnly As Boolean = True
    Dim vInvoiceAllocationsOnly As Boolean = True
    Dim vMaintenanceOnly As Boolean = True
    SetLineTotal()
    'If TransactionLines = 0 Then
    ' vPayPlansOnly = False
    ' vInvoiceAllocationsOnly = False
    ' vMaintenanceOnly = False
    'Else
    If AnalysisDataSet.Tables.Contains("DataRow") Then
      For Each vRow As DataRow In AnalysisDataSet.Tables("DataRow").Rows
        Select Case vRow.Item("TraderLineType").ToString
          Case "P", "G", "S", "H", "M", "C", "O", "E", "A", "I", "V", "N", "U", "VE", "R", "VC", "B", "AP", "D", "F", "Q"  'P payment, G deceased, E event, A accommodation, M membership, C covenant, O order, SO, DD, CC, H hard credit, S soft credit, I incentive, V service, VC -ve service, B Legacy Receipt, D InMemoriamHardCredit, F InMemoriamSoftCredit
            'taltProductSale, taltDeceased, taltSoftCredit, taltHardCredit, taltMembership, taltCovenant, taltPaymentPlan, _
            'taltEvent, taltAccomodation, taltIncentive, taltServiceBooking, taltInvoicePayment, taltUnallocatedSalesLedgerCash, _
            'taltServiceBookingEntitlement, taltSundryCreditNote, taltServiceBookingCredit, taltLegacyBequestReceipt, taltCollectionPayment
            vPayPlansOnly = False
            vInvoiceAllocationsOnly = False
            vMaintenanceOnly = False
            Exit For
          Case "L", "K"     'L taltInvoiceAllocation, K taltSundryCreditNoteInvoiceAllocation
            vPayPlansOnly = False
            'vAllocationsExist = True
            vInvoiceAllocationsOnly = True
            vMaintenanceOnly = False
          Case "AS", "GA", "ST", "CP", "AA", "GD", "GP", "CG", "ADDR" 'AS Add Suppression, GA Gone Away, ST Set Status, CP Cancel Payment Plan, AA Add Activity, GD Gift Aid Declaration, GP GAYE Pledge, CG Cancel Gift Aid Declaration
            'taltAddSuppression, taltGoneAway, taltStatus, taltCancelPaymentPlan, taltActivityEntry, taltGiftAidDeclaration, _
            'taltPayrollGivingPledge, taltCancelGiftAidDeclaration, taltAddressUpdate
            vPayPlansOnly = False
            vInvoiceAllocationsOnly = False
            vMaintenanceOnly = True
        End Select
      Next
    End If
    '    End If
    Return (Not (vPayPlansOnly Or vInvoiceAllocationsOnly Or vMaintenanceOnly))
  End Function

  Public Sub SetPaymentPlanCreated(ByVal pPPNumber As Integer)
    Dim vPP As New PaymentPlanInfo(pPPNumber)
    PaymentPlans.Add(pPPNumber.ToString, vPP)
    'mvPPCreated = True
    mvPPNumbers.Add(pPPNumber.ToString, pPPNumber)
    If IncentiveDataSet IsNot Nothing Then PPIncentivesCompleted = True
  End Sub
  'Public ReadOnly Property PaymentPlanCreated() As Boolean
  '  Get
  '    Return mvPPCreated
  '  End Get
  'End Property

  Public ReadOnly Property PPNumbersCreated() As CollectionList(Of Integer)
    Get
      Return mvPPNumbers
    End Get
  End Property

  Public Sub ClearPPDefinition()
    If PPDDataSet IsNot Nothing AndAlso PPDDataSet.Tables.Contains("DataRow") Then PPDDataSet.Tables.Remove("DataRow")
    SetPPDLineTotal()
  End Sub
  Public ReadOnly Property PPPaymentMethod() As String
    Get
      Select Case PPPaymentType
        Case "STDO"
          PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_so)
        Case "DIRD"
          PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dd)
        Case "CCCA"
          PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_ccca)
        Case Else
          Select Case TransactionPaymentMethod
            Case "CASH"
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash)
            Case "POST"
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_po)
            Case "CHEQ", "CQIN"
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cheque)
            Case "CRED"
              PPPaymentMethod = CSPaymentMethod
            Case "CARD", "CCIN"
              If CreditCard And Not DebitCard Then
                PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cc)       'This app only support credit cards
              ElseIf DebitCard And Not CreditCard Then
                PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dc)       'This app only supports debit cards
              Else
                PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cc)       'We don't know yet, so default to credit card
              End If
            Case "VOUC"
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_voucher)
            Case "CAFC"
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_caf_card)
            Case Else
              PPPaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash)
          End Select
      End Select
    End Get
  End Property

  Friend ReadOnly Property CABankAccount() As String
    Get
      Return mvCABankAccount
    End Get
  End Property

  Friend Sub SetStockTransactionValues(ByVal pStockTransactionID As Integer, ByVal pStockIssued As Integer, ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer)
    mvStockTransactionID = pStockTransactionID
    mvStockIssued = pStockIssued
    'Store last Stock values
    mvStockProductCode = pProductCode
    mvStockWarehouseCode = pWarehouseCode
    mvStockQuantity = pQuantity
    WarehouseChanged = True
  End Sub

  Friend ReadOnly Property StockIssued() As Integer
    Get
      Return mvStockIssued
    End Get
  End Property

  Friend Property StockSales() As Boolean
    Get
      Return mvStockSales
    End Get
    Set(ByVal pValue As Boolean)
      'Only set mvStockSales to True for a Stock Product and CreditNotes = Fales
      If pValue = True AndAlso CreditNotes = False Then
        mvStockSales = True
      Else
        mvStockSales = False
      End If
      If pValue = False Then
        SetStockTransactionValues(0, 0, "", "", 0)
      End If
    End Set
  End Property

  Friend ReadOnly Property StockTransactionID() As Integer
    Get
      Return mvStockTransactionID
    End Get
  End Property

  Friend Function StockValuesChanged(ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer, ByVal pValidatePage As Boolean) As Boolean
    'Check whether the Product/Warehouse/Quantity have changed since last StockMovement created
    'pValidatePage will be True when validating the tpProductDetails page and is used to ensure that we have always created a StockMovement
    Dim vValuesChanged As Boolean = False

    If mvStockSales Then
      If mvStockProductCode.Length = 0 Then
        'Onle set as changed if we are validating the page
        If pValidatePage = True Then vValuesChanged = True
      Else
        If (mvStockProductCode <> pProductCode) OrElse (mvStockWarehouseCode <> pWarehouseCode) OrElse (mvStockQuantity <> pQuantity) Then vValuesChanged = True
      End If
      If WarehouseChanged = True Then vValuesChanged = True
    End If

    Return vValuesChanged

  End Function

  Public Property AutoPaymentCreated() As Boolean
    Get
      Return mvAutoPaymentCreated
    End Get
    Set(ByVal pValue As Boolean)
      mvAutoPaymentCreated = pValue
    End Set
  End Property

  Public ReadOnly Property AlbacsBankDetails() As String
    Get
      Return mvAlbacsBankDetails
    End Get
  End Property

  Friend Sub ClearEventBookingDataSet()
    If EventBookingDataSet.Tables.Contains("DataRow") Then EventBookingDataSet.Tables.Remove("DataRow")
  End Sub

  Friend Property LastDeceasedContactNumber() As Integer
    Get
      Return mvLastDeceasedContactNumber
    End Get
    Set(ByVal pValue As Integer)
      mvLastDeceasedContactNumber = pValue
    End Set
  End Property

  Public ReadOnly Property PurchaseOrderType As PurchaseOrderTypes
    Get
      Return mvPurchaseOrderType
    End Get
  End Property
  Public Sub SetPurchaseOrderType(ByVal pPurchaseOrderType As DataRow)
    If BooleanValue(pPurchaseOrderType("RegularPayments").ToString) Then
      mvPurchaseOrderType = PurchaseOrderTypes.RegularPayments
    ElseIf BooleanValue(pPurchaseOrderType("AdHocPayments").ToString) Then
      mvPurchaseOrderType = PurchaseOrderTypes.AdHocPayments
    ElseIf BooleanValue(pPurchaseOrderType("PaymentSchedule").ToString) Then
      mvPurchaseOrderType = PurchaseOrderTypes.PaymentSchedule
    Else
      mvPurchaseOrderType = PurchaseOrderTypes.None
    End If
  End Sub

  ''' <summary>Sets the Delivery Contact and Address for Product Sale lines in order to default subsequent lines correctly</summary>
  Private Sub SetDeliveryContactAndAddress(ByVal pDeliveryContactNumber As Integer, ByVal pDeliveryAddressNumber As Integer)
    If (pDeliveryContactNumber > 0 AndAlso pDeliveryAddressNumber > 0) Then
      mvDeliveryContactNumber = pDeliveryContactNumber
      mvDeliveryAddressNumber = pDeliveryAddressNumber
    Else
      mvDeliveryContactNumber = 0
      mvDeliveryAddressNumber = 0
    End If
  End Sub
  Friend ReadOnly Property DeliveryContactNumber() As Integer
    Get
      Return mvDeliveryContactNumber
    End Get
  End Property
  Friend ReadOnly Property DeliveryAddressNumber() As Integer
    Get
      Return mvDeliveryAddressNumber
    End Get
  End Property

  Friend Sub SetPaymentPlanDetailsPricing(ByVal pPPDPricing As PaymentPlanDetailsPricing)
    mvPPDetailsPricing = pPPDPricing
  End Sub

  Friend Sub ClearPaymentPlanDetailsPricing()
    mvPPDetailsPricing = Nothing
  End Sub

  Friend ReadOnly Property PaymentPlanDetailsPricing As PaymentPlanDetailsPricing
    Get
      If mvPPDetailsPricing Is Nothing Then mvPPDetailsPricing = New PaymentPlanDetailsPricing
      Return mvPPDetailsPricing
    End Get
  End Property

  Friend Property CreditListRecAdditionalCriteria As ParameterList
    Get
      If mvCLRAdditionalCriteria Is Nothing Then mvCLRAdditionalCriteria = New ParameterList
      Return mvCLRAdditionalCriteria
    End Get
    Set(value As ParameterList)
      mvCLRAdditionalCriteria = value
    End Set
  End Property

  Friend Property TraderAlerts() As Boolean
    Get
      Return mvTraderAlerts
    End Get
    Private Set(value As Boolean)
      mvTraderAlerts = value
    End Set
  End Property

#Region "Multicurrency"
  Public Function MultiCurrency() As Boolean
    Dim vMultiCurrencies As Boolean
    Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCurrencyAndUnicode)
    If BooleanValue(vDT.Rows(0).Item("CurrencyCodes").ToString) Then
      vMultiCurrencies = True
    End If
    Return vMultiCurrencies
  End Function

  Public Function CalcCurrencyAmount(ByVal pAmount As Double, ByVal pAsBaseValue As Boolean) As Double
    Dim vAmount As Double
    If MultiCurrency() Then
      Dim vBatchInfo As New BatchInfo(BatchNumber)
      vAmount = vBatchInfo.CalculateCurrencyAmount(pAmount, pAsBaseValue)
    Else
      vAmount = FixTwoPlaces(pAmount)
    End If
    Return vAmount
  End Function
#End Region

  Public Property TokenDescription As String

 
End Class
