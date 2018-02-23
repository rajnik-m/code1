Namespace Access
  Public Class TraderApplication

    Public LastProduct As String
    Public LastRate As String
    Public LastMembershipType As String
    Public LastMembershipRate As String
    Public LastEventNumber As String
    Public LastBookingOptionNumber As String
    Public LastEventBookingRate As String
    Public StatementDate As String

    Private mvEnv As CDBEnvironment
    Private mvTraderPages As TraderPages
    Private mvTraderControls As TraderControls

    'Application Properties From fp_applications Table
    Private mvApplication As String
    Private mvAppDesc As String
    Private mvAppType As ApplicationType
    Private mvAppTypeCode As String
    Private mvBatchLedApp As Boolean
    Private mvCash As Boolean
    Private mvCheque As Boolean
    Private mvPostalOrder As Boolean
    Private mvCreditCard As Boolean
    Private mvDebitCard As Boolean
    Private mvCreditSales As Boolean
    Private mvChequeWithInvoice As Boolean
    Private mvCCWithInvoice As Boolean
    Private mvBankDetails As Boolean
    Private mvTransactionComments As Boolean
    Private mvMemberships As Boolean
    Private mvChangeMembership As Boolean
    Private mvSubscriptions As Boolean
    Private mvDonationsRegular As Boolean
    Private mvCovMemberships As Boolean
    Private mvCovSubscriptions As Boolean
    Private mvCovDonationsRegular As Boolean
    Private mvProductSales As Boolean
    Private mvOneOffDonations As Boolean
    Private mvPayments As Boolean
    Private mvStandingOrders As Boolean
    Private mvDirectDebits As Boolean
    Private mvCreditCardAuthorities As Boolean
    Private mvEventBooking As Boolean
    Private mvExamBooking As Boolean
    Private mvAccomodationBooking As Boolean
    Private mvSource As String
    Private mvCampaign As String
    Private mvAppeal As String
    Private mvSourceInvalid As Boolean
    Private mvCABankAccount As String
    Private mvCCBankAccount As String
    Private mvDCBankAccount As String
    Private mvCSBankAccount As String
    Private mvSOBankAccount As String
    Private mvDDBankAccount As String
    Private mvCCABankAccount As String
    Private mvShowTransReference As Boolean
    Private mvConfirmDefaultProduct As Boolean
    Private mvConfirmAnalysis As Boolean
    Private mvCarriage As Boolean
    Private mvConfirmCarriage As Boolean
    Private mvCarriageProduct As String
    Private mvCarriageRate As String
    Private mvCarriagePercentage As Double
    Private mvNoPaymentRequired As Boolean
    Private mvServiceBookings As Boolean
    Private mvAnalysisComments As Boolean
    Private mvPayMethodsAtEnd As Boolean
    Private mvDefaultSalesContact As Integer
    Private mvInvoicePayments As Boolean
    Private mvPayPlanPayMethod As Boolean
    Private mvCreditNotes As Boolean
    Private mvForeignCurrency As Boolean
    Private mvInvoiceDocument As String
    Private mvReceiptDocument As String
    Private mvPayPlanDocument As String
    Private mvCreditStatementDocument As String
    Private mvMembersOnly As Boolean
    Private mvSalesGroup As String
    Private mvAutoSetAmount As Boolean
    Private mvServiceBookingCredits As Boolean
    Private mvLegacyReceipt As Boolean
    Private mvGoneAway As Boolean
    Private mvSetStatus As Boolean
    Private mvDefaultStatus As String
    Private mvAddActivity As Boolean
    Private mvDefaultActivityGroup As String
    Private mvAddSuppression As Boolean
    Private mvDefaultSuppression As String
    Private mvCancelPayPlan As Boolean
    Private mvDefaultCancellationReason As String
    Private mvGiftAidDeclaration As Boolean
    Private mvConfirmPayPlanDetails As Boolean
    Private mvDistributionCodeMandatory As Boolean
    Private mvSalesContactMandatory As Boolean
    Private mvIncludeProvisionalTrans As Boolean
    Private mvIncludeConfirmedTrans As Boolean
    Private mvVoucher As Boolean
    Private mvCAFCard As Boolean
    Private mvGiftInKind As Boolean
    Private mvCVBankAccount As String
    Private mvDonationProduct As String
    Private mvDonationRate As String
    Private mvPayrollGiving As Boolean
    Private mvBatchCategory As String
    Private mvAppHeight As Integer
    Private mvAppWidth As Integer
    Private mvSourceFromLastMailing As Boolean
    Private mvMailingCodeMandatory As Boolean
    Private mvBypassMailingParagraphs As Boolean
    Private mvSupportsNonFinancialBatch As Boolean
    Private mvAddressMaintenance As Boolean
    Private mvGiftAidCancellation As Boolean
    Private mvAutoPaymentMaintenance As Boolean
    Private mvSaleOrReturn As Boolean
    Private mvProvisionalPaymentPlan As Boolean
    Private mvDefaultMemberBranch As String
    Private mvProvisionalCashTransDoc As String
    Private mvPreFulfilledIncentives As Boolean
    Private mvContactAlerts As Boolean
    Private mvDisplayScheduledPayments As Boolean
    Private mvConfirmSRTransactions As Boolean
    Private mvOnLineCCAuthorisation As Boolean
    Private mvRequireCCAuthorisation As Boolean
    Private mvPayPlanConvMaintenance As Boolean
    Private mvLinkToCommunication As String
    Private mvCollectionPayments As Boolean
    Private mvBatchAnalysisCode As String
    Private mvEventMultipleAnalysis As Boolean
    Private mvDefaultTransactionOrigin As String
    Private mvServiceBookingAnalysis As Boolean
    Private mvAlbacsBankDetails As String
    Private mvLinkToFundraisingPayments As Boolean
    Private mvInvoicePrintPreviewDefault As Boolean
    Private mvLoans As Boolean
    Private mvAutoCreateCreditCustomer As Boolean
    Private mvUnpostedBatchMsgInPrint As Boolean
    Private mvDateRangeMsgInPrint As Boolean
    Private mvDefaultExamSessionCode As String
    Private mvDefaultExamUnitCode As String
    Private mvInvoicePrintUnpostedBatches As Boolean
    Private mvUseToken As Boolean?

    Private mvCreditCategory As String

    'Default Product and Rate variables
    Private mvProduct As String = ""
    Private mvRate As String = ""
    Private mvProductFromSource As String = ""
    Private mvRateFromSource As String = ""
    Private mvProductNumbersFromSource As Boolean

    'Batch Properties
    Private mvBatchNumber As Integer
    Private mvTransNumber As Integer
    Private mvBatchType As String
    Private mvBatchDate As String = ""
    Private mvPostedToCashBook As Boolean
    Private mvPostedToNominal As Boolean
    Private mvBatchExchangeRate As Double
    Private mvBatchPayMethod As String
    Private mvBatchCurrency As String = ""
    Private mvBatch As Batch
    Private mvNonFinancialBatch As Batch
    Private mvNFBatchNumber As Integer
    Private mvNFTransNumber As Integer

    'Supplementary Application Properties
    Private mvMainPage As TraderPage.TraderPageType
    Private mvSummaryPage As TraderPage.TraderPageType
    Private mvLinePage As TraderPage.TraderPageType
    Private mvAmountTag As String
    Private mvTotalTag As String
    Private mvCurrentPrice As Double
    Private mvSourceDesc As String
    Private mvDistributionCode As String
    Private mvDistributionCodeDesc As String
    Private mvMailing As String
    Private mvMailingDesc As String
    Private mvGiftAidMinimum As Double
    Private mvCovReason As String
    Private mvMinCovPeriod As Integer
    Private mvCovActivity As String
    Private mvCovActivityValue As String
    Private mvMemReason As String
    Private mvSponsorAct As String
    Private mvSponsorActVal As String
    Private mvCMTCancelReason As String
    Private mvCCReason As String
    Private mvDDReason As String
    Private mvSOReason As String
    Private mvOReason As String
    Private mvDistActivity As String
    Private mvStockInterface As Boolean
    Private mvUseSalesLedger As Boolean
    Private mvCSPayMethod As String
    Private mvCSTransType As String
    Private mvSundryCreditTransType As String
    Private mvCACompany As String
    Private mvCSCompany As String
    Private mvCSTermsNumber As Integer
    Private mvCSTermsPeriod As String
    Private mvCSTermsFrom As String
    Private mvCSDepositPercentage As Double
    Private mvCSSundryCreditProduct As String
    Private mvCSSundryCreditRate As String
    Private mvDefaultMailingType As DefaultMailingTypes
    Private mvDiscountActivity As String
    Private mvAmendDiscount As Boolean
    Private mvCoreStockControl As Boolean
    Private mvPostageWarnStock As Boolean
    Private mvBackOrderPrompt As Boolean
    Private mvDefaultCurrencyCode As String
    Private mvBatchPerUser As BatchCreationTypes
    Private mvCurrencyBankAccounts As Collection
    Private mvDefaultStatusDesc As String
    Private mvDefaultSuppresionDesc As String
    Private mvDefaultCancellationReasonDesc As String
    Private mvMaintenanceOnly As Boolean
    Private mvBlankMembershipJoinedDate As Boolean
    Private mvMaxVoucherTransactions As Integer
    Private mvSalesQuantity As String
    Private mvUsesProductNumbers As Boolean
    Private mvShowPPDInConversionApp As Boolean
    Private mvCAFFloorLimit As Double
    Private mvNonCAFFloorLimit As Double
    Private mvCAFFloorLimitRead As Boolean
    Private mvNonCAFFloorLimitRead As Boolean
    Private mvOptionalCreditSalesReference As Boolean
    Private mvPackToDonorDefault As String
    Private mvBankTransaction As BankTransaction
    Private mvPaymentPlanDetails As TraderPaymentPlanDetails 'Collection of Payment Plan Details
    Private mvMultiCurrency As Boolean
    Private mvStockMovementTransactionID As Integer
    Private mvExistingAdjustmentTran As Boolean
    Private mvCheckIncentives As Boolean
    Private mvValid As Boolean 'set to True if .Init successful

    Private mvAutoGiftAidDeclaration As Boolean
    Private mvAutoGiftAidMethod As String
    Private mvAutoGiftAidSource As String
    Private mvMerchantRetailNumber As String
    Private mvTraderAlerts As Boolean

    Public Enum TraderProcessDataTypes
      tpdtFirstPage
      tpdtNextPage
      tpdtPreviousPage
      tpdtFinished
      tpdtEditTransaction
      tpdtEditAnalysisLine
      tpdtCancelTransaction
      tpdtDeleteAnalysisLine
      tpdtAddMemberSummary
      tpdtAmendMemberSummary
      tpdtDeletePaymentPlanLine
    End Enum

    Public Enum ApplicationType
      atTransaction = 1 'TRANS
      atPurchaseInvoice 'PINVE
      atPurchaseOrder 'PORDE
      atPurchaseOrderCancellation 'PORDC
      atChequeNumberAllocation 'CHQNA
      atChequeReconciliation 'CHQRE
      atCreditStatementGeneration 'CSTAT
      atBatchInvoiceGeneration 'BINVG
      atMaintenance 'MAINT
      atConversion 'CNVRT
      atCreditListReconciliation 'CLREC
      atBankStatementPosting 'BSPOS
      atPurchaseOrderGeneration 'POGEN    does not use trader form
      atPurchaseOrderPrint 'POPRT    does not use trader form
      atChequeProcessing 'POCHQ    does not use trader form
      atGiveAsYouEarnPayments 'GAYEP    (Pre Tax Payroll Giving)
      atPostTaxPGPayments 'POTPG    (Post Tax Payroll Giving)
    End Enum

    Public Enum TraderControlTypes 'Constants for the different loaded control types
      tctTXT = 1
      tctCHK = 2
      tctOPT = 4
      tctCMD = 8
      tctFND = 16
      tctDSC = 32
      tctTXN = 64
      tctCPN = 128
      tctMEB = 256
      tctTXM = 512
      tctDTP = 1024
      tctTTP = 2048
      tctHDR = 4096
      tctSSC = 8192
      tctACE = 16384
      tctCBO = 32768
      tctSUP = 65536
      tctCBD = 131072
      tctALL = &H7FFFFFFF
    End Enum

    Public Enum BatchCreationTypes
      bctPerUser
      bctPerDepartment
      bctPerSystem
    End Enum

    Public Enum DefaultMailingTypes
      dmtNone
      dmtLetterBreaks
      dmtSource
      dmtLetterBreaksOrSource
    End Enum

    Public Enum LinkToCommunicationTypes
      ltcYes
      ltcNo
      ltcAsk
    End Enum

    Public Enum SaveTransactionStatus
      stsNone = 0
      stsComplete = 1
      stsPrintInvoice = 2
      stsPrintReceipt = 4
      stsPrintProvisionalCashDoc = 8
      stsCreateMailingDocument = 16
      stsContactWarningSuppressionsPrompt = 32
    End Enum

    Const STD_APP_HEIGHT As Integer = 5955
    Const STD_APP_WIDTH As Integer = 8950
    Const DEFAULT_CAF_CARD_NUMBER As String = "564193"

    Public ReadOnly Property BankTransaction() As BankTransaction
      Get
        If mvBankTransaction Is Nothing Then
          mvBankTransaction = New BankTransaction
          mvBankTransaction.Init(mvEnv)
        End If
        BankTransaction = mvBankTransaction
      End Get
    End Property

    Public Property BatchCurrencyCode() As String
      Get
        BatchCurrencyCode = mvBatchCurrency
      End Get
      Set(ByVal Value As String)
        mvBatchCurrency = Value
        If mvPayPlanPayMethod Then
          If mvBatchCurrency <> DefaultCurrencyCode Then mvPayPlanPayMethod = False
        End If
        If mvBatchCurrency.Length > 0 Then
          If mvBatchExchangeRate = 0 Then mvBatchExchangeRate = mvEnv.GetExchangeRate(mvBatchCurrency)
        Else
          mvBatchExchangeRate = 0
        End If
      End Set
    End Property
    Public ReadOnly Property CAFCard() As Boolean
      Get
        CAFCard = mvCAFCard
      End Get
    End Property
    Public ReadOnly Property CVBankAccount() As String
      Get
        CVBankAccount = mvCVBankAccount
      End Get
    End Property

    Public ReadOnly Property DefaultProductUsesProductNumbers() As Boolean
      Get
        If Len(mvProductFromSource) > 0 Then
          DefaultProductUsesProductNumbers = mvProductNumbersFromSource
        ElseIf Len(mvProduct) > 0 Then
          DefaultProductUsesProductNumbers = mvUsesProductNumbers
        End If
      End Get
    End Property

    Public ReadOnly Property IsDefaultProductAndRate() As Boolean
      Get
        If ((Len(mvProductFromSource) > 0 And Len(mvRateFromSource) > 0) Or (Len(mvProduct) > 0 And Len(mvRate) > 0)) Then IsDefaultProductAndRate = True
      End Get
    End Property

    Public ReadOnly Property AppDesc() As String
      Get
        AppDesc = mvAppDesc
      End Get
    End Property
    Public ReadOnly Property AppType() As ApplicationType
      Get
        AppType = mvAppType
      End Get
    End Property
    Public ReadOnly Property AppTypeCode() As String
      Get
        AppTypeCode = mvAppTypeCode
      End Get
    End Property
    Public ReadOnly Property BatchLedApp() As Boolean
      Get
        BatchLedApp = mvBatchLedApp
      End Get
    End Property
    Public ReadOnly Property Cash() As Boolean
      Get
        Cash = mvCash
      End Get
    End Property
    Public ReadOnly Property Cheque() As Boolean
      Get
        Cheque = mvCheque
      End Get
    End Property
    Public ReadOnly Property GiftInKind() As Boolean
      Get
        GiftInKind = mvGiftInKind
      End Get
    End Property
    Public ReadOnly Property PostalOrder() As Boolean
      Get
        PostalOrder = mvPostalOrder
      End Get
    End Property
    Public ReadOnly Property CreditCard() As Boolean
      Get
        CreditCard = mvCreditCard
      End Get
    End Property
    Public ReadOnly Property DebitCard() As Boolean
      Get
        DebitCard = mvDebitCard
      End Get
    End Property
    Public ReadOnly Property CreditSales() As Boolean
      Get
        CreditSales = mvCreditSales
      End Get
    End Property
    Public ReadOnly Property ChequeWithInvoice() As Boolean
      Get
        Return mvChequeWithInvoice
      End Get
    End Property
    Public ReadOnly Property CCWithInvoice() As Boolean
      Get
        Return mvCCWithInvoice
      End Get
    End Property
    Public ReadOnly Property BankDetails() As Boolean
      Get
        BankDetails = mvBankDetails
      End Get
    End Property
    Public ReadOnly Property TransactionComments() As Boolean
      Get
        TransactionComments = mvTransactionComments
      End Get
    End Property
    Public ReadOnly Property Memberships() As Boolean
      Get
        Memberships = mvMemberships
      End Get
    End Property
    Public ReadOnly Property ChangeMembership() As Boolean
      Get
        ChangeMembership = mvChangeMembership
      End Get
    End Property
    Public ReadOnly Property Subscriptions() As Boolean
      Get
        Subscriptions = mvSubscriptions
      End Get
    End Property
    Public ReadOnly Property DonationsRegular() As Boolean
      Get
        DonationsRegular = mvDonationsRegular
      End Get
    End Property
    Public ReadOnly Property CovMemberships() As Boolean
      Get
        CovMemberships = mvCovMemberships
      End Get
    End Property
    Public ReadOnly Property CovSubscriptions() As Boolean
      Get
        CovSubscriptions = mvCovSubscriptions
      End Get
    End Property
    Public ReadOnly Property CovDonationsRegular() As Boolean
      Get
        CovDonationsRegular = mvCovDonationsRegular
      End Get
    End Property
    Public ReadOnly Property ProductSales() As Boolean
      Get
        ProductSales = mvProductSales
      End Get
    End Property
    Public ReadOnly Property OneOffDonations() As Boolean
      Get
        OneOffDonations = mvOneOffDonations
      End Get
    End Property
    Public ReadOnly Property Payments() As Boolean
      Get
        Payments = mvPayments
      End Get
    End Property
    Public ReadOnly Property PaymentPlans() As Boolean
      Get
        PaymentPlans = Memberships Or ChangeMembership Or Subscriptions Or DonationsRegular Or CovMemberships Or CovSubscriptions Or CovDonationsRegular Or PayPlanPayMethod
      End Get
    End Property
    Public ReadOnly Property StandingOrders() As Boolean
      Get
        StandingOrders = mvStandingOrders
      End Get
    End Property
    Public ReadOnly Property DirectDebits() As Boolean
      Get
        DirectDebits = mvDirectDebits
      End Get
    End Property
    Public ReadOnly Property CreditCardAuthorities() As Boolean
      Get
        CreditCardAuthorities = mvCreditCardAuthorities
      End Get
    End Property
    Public ReadOnly Property SupportsEventBooking() As Boolean
      Get
        SupportsEventBooking = mvEventBooking
      End Get
    End Property
    Public ReadOnly Property SupportsExamBooking() As Boolean
      Get
        Return mvExamBooking
      End Get
    End Property
    Public ReadOnly Property LinkToCommunication() As LinkToCommunicationTypes
      Get
        Select Case mvLinkToCommunication
          Case "Y"
            LinkToCommunication = LinkToCommunicationTypes.ltcYes
          Case "A"
            LinkToCommunication = LinkToCommunicationTypes.ltcAsk
          Case Else
            ' "N" - default
            LinkToCommunication = LinkToCommunicationTypes.ltcNo
        End Select
      End Get
    End Property
    Public ReadOnly Property AccomodationBooking() As Boolean
      Get
        AccomodationBooking = mvAccomodationBooking
      End Get
    End Property
    Public ReadOnly Property Product() As String
      Get
        Product = mvProduct
      End Get
    End Property
    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvRate
      End Get
    End Property
    Public ReadOnly Property DonationProduct() As String
      Get
        DonationProduct = mvDonationProduct
      End Get
    End Property
    Public ReadOnly Property DonationRate() As String
      Get
        DonationRate = mvDonationRate
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Source = mvSource
      End Get
    End Property
    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvCampaign
      End Get
    End Property
    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvAppeal
      End Get
    End Property
    Public ReadOnly Property SourceValid() As Boolean
      Get
        SourceValid = Not mvSourceInvalid
      End Get
    End Property
    Public ReadOnly Property CABankAccount() As String
      Get
        CABankAccount = mvCABankAccount
      End Get
    End Property
    Public ReadOnly Property CCBankAccount() As String
      Get
        CCBankAccount = mvCCBankAccount
      End Get
    End Property
    Public ReadOnly Property DCBankAccount() As String
      Get
        DCBankAccount = mvDCBankAccount
      End Get
    End Property
    Public ReadOnly Property CSBankAccount() As String
      Get
        CSBankAccount = mvCSBankAccount
      End Get
    End Property
    Public ReadOnly Property SOBankAccount() As String
      Get
        SOBankAccount = mvSOBankAccount
      End Get
    End Property
    Public ReadOnly Property DDBankAccount() As String
      Get
        DDBankAccount = mvDDBankAccount
      End Get
    End Property
    Public ReadOnly Property CCABankAccount() As String
      Get
        CCABankAccount = mvCCABankAccount
      End Get
    End Property
    Public ReadOnly Property ShowTransReference() As Boolean
      Get
        ShowTransReference = mvShowTransReference
      End Get
    End Property
    Public ReadOnly Property ConfirmDefaultProduct() As Boolean
      Get
        ConfirmDefaultProduct = mvConfirmDefaultProduct
      End Get
    End Property
    Public ReadOnly Property ConfirmAnalysis() As Boolean
      Get
        ConfirmAnalysis = mvConfirmAnalysis
      End Get
    End Property
    Public ReadOnly Property Carriage() As Boolean
      Get
        Carriage = mvCarriage
      End Get
    End Property
    Public ReadOnly Property ConfirmCarriage() As Boolean
      Get
        ConfirmCarriage = mvConfirmCarriage
      End Get
    End Property
    Public ReadOnly Property CarriageProduct() As String
      Get
        CarriageProduct = mvCarriageProduct
      End Get
    End Property
    Public ReadOnly Property CarriageRate() As String
      Get
        CarriageRate = mvCarriageRate
      End Get
    End Property
    Public ReadOnly Property CarriagePercentage() As Double
      Get
        CarriagePercentage = mvCarriagePercentage
      End Get
    End Property
    Public ReadOnly Property NoPaymentRequired() As Boolean
      Get
        NoPaymentRequired = mvNoPaymentRequired
      End Get
    End Property
    Public ReadOnly Property ServiceBookings() As Boolean
      Get
        ServiceBookings = mvServiceBookings
      End Get
    End Property
    Public ReadOnly Property AnalysisComments() As Boolean
      Get
        AnalysisComments = mvAnalysisComments
      End Get
    End Property
    Public ReadOnly Property PayMethodsAtEnd() As Boolean
      Get
        PayMethodsAtEnd = mvPayMethodsAtEnd
      End Get
    End Property
    Public ReadOnly Property AutoSetAmount() As Boolean
      Get
        AutoSetAmount = mvAutoSetAmount
      End Get
    End Property
    Public ReadOnly Property DefaultSalesContact() As Integer
      Get
        DefaultSalesContact = mvDefaultSalesContact
      End Get
    End Property
    Public Property InvoicePayments() As Boolean
      Get
        InvoicePayments = mvInvoicePayments
      End Get
      Set(ByVal Value As Boolean)
        mvInvoicePayments = Value
      End Set
    End Property
    Public ReadOnly Property PayPlanPayMethod() As Boolean
      Get
        PayPlanPayMethod = mvPayPlanPayMethod
      End Get
    End Property
    Public ReadOnly Property CreditNotes() As Boolean
      Get
        CreditNotes = mvCreditNotes
      End Get
    End Property
    Public ReadOnly Property InvoiceDocument() As String
      Get
        InvoiceDocument = mvInvoiceDocument
      End Get
    End Property
    Public ReadOnly Property ReceiptDocument() As String
      Get
        ReceiptDocument = mvReceiptDocument
      End Get
    End Property
    Public ReadOnly Property PayPlanDocument() As String
      Get
        PayPlanDocument = mvPayPlanDocument
      End Get
    End Property
    Public ReadOnly Property CreditStatementDocument() As String
      Get
        CreditStatementDocument = mvCreditStatementDocument
      End Get
    End Property
    Public ReadOnly Property MembersOnly() As Boolean
      Get
        MembersOnly = mvMembersOnly
      End Get
    End Property
    Public ReadOnly Property SalesGroup() As String
      Get
        SalesGroup = mvSalesGroup
      End Get
    End Property
    Public Property BatchDate() As String
      Get
        BatchDate = mvBatchDate
      End Get
      Set(ByVal Value As String)
        mvBatchDate = Value
      End Set
    End Property
    Public ReadOnly Property BatchType() As String
      Get
        BatchType = mvBatchType
      End Get
    End Property
    Public ReadOnly Property PostedToNominal() As Boolean
      Get
        PostedToNominal = mvPostedToNominal
      End Get
    End Property
    Public ReadOnly Property PostedToCashBook() As Boolean
      Get
        PostedToCashBook = mvPostedToCashBook
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvBatchNumber
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvTransNumber
      End Get
    End Property

    Public ReadOnly Property NonFinancialBatchNumber() As Integer
      Get
        NonFinancialBatchNumber = mvNFBatchNumber
      End Get
    End Property
    Public ReadOnly Property NonFinancialTransactionNumber() As Integer
      Get
        NonFinancialTransactionNumber = mvNFTransNumber
      End Get
    End Property

    Public ReadOnly Property Application() As String
      Get
        Application = mvApplication
      End Get
    End Property
    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvDistributionCode
      End Get
    End Property
    Public ReadOnly Property DistributionCodeDesc() As String
      Get
        DistributionCodeDesc = mvDistributionCodeDesc
      End Get
    End Property
    Public ReadOnly Property DistributionCodeMandatory() As Boolean
      Get
        Return mvDistributionCodeMandatory
      End Get
    End Property
    Public ReadOnly Property SalesContactMandatory() As Boolean
      Get
        Return mvSalesContactMandatory
      End Get
    End Property
    Public ReadOnly Property CurrentPrice() As Double
      Get
        CurrentPrice = mvCurrentPrice
      End Get
    End Property
    Public ReadOnly Property SourceDesc() As String
      Get
        SourceDesc = mvSourceDesc
      End Get
    End Property
    Public ReadOnly Property Mailing() As String
      Get
        Mailing = mvMailing
      End Get
    End Property
    Public ReadOnly Property MailingDesc() As String
      Get
        MailingDesc = mvMailingDesc
      End Get
    End Property
    Public ReadOnly Property GiftAidMinimum() As Double
      Get
        GiftAidMinimum = mvGiftAidMinimum
      End Get
    End Property
    Public ReadOnly Property CovReason() As String
      Get
        CovReason = mvCovReason
      End Get
    End Property
    Public ReadOnly Property MinCovPeriod() As Integer
      Get
        MinCovPeriod = mvMinCovPeriod
      End Get
    End Property
    Public ReadOnly Property CovActivity() As String
      Get
        CovActivity = mvCovActivity
      End Get
    End Property
    Public ReadOnly Property CovActivityValue() As String
      Get
        CovActivityValue = mvCovActivityValue
      End Get
    End Property
    Public ReadOnly Property MemReason() As String
      Get
        MemReason = mvMemReason
      End Get
    End Property
    Public ReadOnly Property SponsorAct() As String
      Get
        SponsorAct = mvSponsorAct
      End Get
    End Property
    Public ReadOnly Property SponsorActVal() As String
      Get
        SponsorActVal = mvSponsorActVal
      End Get
    End Property
    Public ReadOnly Property CMTCancelReason() As String
      Get
        CMTCancelReason = mvCMTCancelReason
      End Get
    End Property
    Public ReadOnly Property CCReason() As String
      Get
        CCReason = mvCCReason
      End Get
    End Property
    Public ReadOnly Property DDReason() As String
      Get
        DDReason = mvDDReason
      End Get
    End Property
    Public ReadOnly Property SOReason() As String
      Get
        SOReason = mvSOReason
      End Get
    End Property
    Public ReadOnly Property OReason() As String
      Get
        OReason = mvOReason
      End Get
    End Property
    Public ReadOnly Property DistributorActivity() As String
      Get
        DistributorActivity = mvDistActivity
      End Get
    End Property
    Public ReadOnly Property StockInterface() As Boolean
      Get
        StockInterface = mvStockInterface
      End Get
    End Property
    Public ReadOnly Property UseSalesLedger() As Boolean
      Get
        UseSalesLedger = mvUseSalesLedger
      End Get
    End Property
    Public ReadOnly Property CSPayMethod() As String
      Get
        CSPayMethod = mvCSPayMethod
      End Get
    End Property
    Public ReadOnly Property CSTransType() As String
      Get
        CSTransType = mvCSTransType
      End Get
    End Property
    Public ReadOnly Property SundryCreditTransType() As String
      Get
        SundryCreditTransType = mvSundryCreditTransType
      End Get
    End Property
    Public ReadOnly Property CACompany() As String
      Get
        CACompany = mvCACompany
      End Get
    End Property
    Public ReadOnly Property CSCompany() As String
      Get
        CSCompany = mvCSCompany
      End Get
    End Property
    Public Property CSTermsNumber() As Integer
      Get
        CSTermsNumber = mvCSTermsNumber
      End Get
      Set(ByVal Value As Integer)
        mvCSTermsNumber = Value
      End Set
    End Property
    Public Property CSTermsPeriod() As String
      Get
        CSTermsPeriod = mvCSTermsPeriod
      End Get
      Set(ByVal Value As String)
        mvCSTermsPeriod = Value
      End Set
    End Property
    Public Property CSTermsFrom() As String
      Get
        CSTermsFrom = mvCSTermsFrom
      End Get
      Set(ByVal Value As String)
        mvCSTermsFrom = Value
      End Set
    End Property
    Public ReadOnly Property SundryCreditProduct() As String
      Get
        SundryCreditProduct = mvCSSundryCreditProduct
      End Get
    End Property
    Public ReadOnly Property SundryCreditRate() As String
      Get
        SundryCreditRate = mvCSSundryCreditRate
      End Get
    End Property
    Public ReadOnly Property DefaultMailingType() As DefaultMailingTypes
      Get
        DefaultMailingType = mvDefaultMailingType
      End Get
    End Property
    Public ReadOnly Property BatchExchangeRate() As Double
      Get
        BatchExchangeRate = mvBatchExchangeRate
      End Get
    End Property

    Public ReadOnly Property DiscountActivity() As String
      Get
        DiscountActivity = mvDiscountActivity
      End Get
    End Property
    Public ReadOnly Property AmendDiscount() As Boolean
      Get
        AmendDiscount = mvAmendDiscount
      End Get
    End Property
    Public ReadOnly Property CoreStockControl() As Boolean
      Get
        CoreStockControl = mvCoreStockControl
      End Get
    End Property
    Public ReadOnly Property PostageWarnStock() As Boolean
      Get
        PostageWarnStock = mvPostageWarnStock
      End Get
    End Property
    Public ReadOnly Property BackOrderPrompt() As Boolean
      Get
        BackOrderPrompt = mvBackOrderPrompt
      End Get
    End Property
    Public ReadOnly Property MainPage() As TraderPage.TraderPageType
      Get
        Return mvMainPage
      End Get
    End Property
    Public ReadOnly Property LinePage() As TraderPage.TraderPageType
      Get
        Return mvLinePage
      End Get
    End Property
    Public ReadOnly Property SummaryPage() As TraderPage.TraderPageType
      Get
        Return mvSummaryPage
      End Get
    End Property
    Public ReadOnly Property AmountTag() As String
      Get
        AmountTag = mvAmountTag
      End Get
    End Property
    Public ReadOnly Property TotalTag() As String
      Get
        TotalTag = mvTotalTag
      End Get
    End Property
    Public ReadOnly Property IsValid() As Boolean
      Get
        IsValid = mvValid
      End Get
    End Property
    Public ReadOnly Property DefaultCurrencyCode() As String
      Get
        DefaultCurrencyCode = mvDefaultCurrencyCode
      End Get
    End Property
    Public ReadOnly Property BatchPerUser() As BatchCreationTypes
      Get
        BatchPerUser = mvBatchPerUser
      End Get
    End Property
    Public ReadOnly Property CurrencyBankAccounts() As Collection
      Get
        CurrencyBankAccounts = mvCurrencyBankAccounts
      End Get
    End Property
    Public ReadOnly Property ServiceBookingCredits() As Boolean
      Get
        ServiceBookingCredits = mvServiceBookingCredits
      End Get
    End Property
    Public ReadOnly Property LegaceyReceipt() As Boolean
      Get
        LegaceyReceipt = mvLegacyReceipt
      End Get
    End Property
    Public ReadOnly Property GoneAway() As Boolean
      Get
        GoneAway = mvGoneAway
      End Get
    End Property
    Public ReadOnly Property SetStatus() As Boolean
      Get
        SetStatus = mvSetStatus
      End Get
    End Property
    Public ReadOnly Property DefaultStatus() As String
      Get
        DefaultStatus = mvDefaultStatus
      End Get
    End Property
    Public ReadOnly Property DefaultStatusDesc() As String
      Get
        DefaultStatusDesc = mvDefaultStatusDesc
      End Get
    End Property
    Public ReadOnly Property AddActivity() As Boolean
      Get
        AddActivity = mvAddActivity
      End Get
    End Property
    Public ReadOnly Property DefaultActivityGroup() As String
      Get
        DefaultActivityGroup = mvDefaultActivityGroup
      End Get
    End Property
    Public ReadOnly Property AddSuppression() As Boolean
      Get
        AddSuppression = mvAddSuppression
      End Get
    End Property
    Public ReadOnly Property DefaultSuppression() As String
      Get
        DefaultSuppression = mvDefaultSuppression
      End Get
    End Property
    Public ReadOnly Property DefaultSuppressionDesc() As String
      Get
        DefaultSuppressionDesc = mvDefaultSuppresionDesc
      End Get
    End Property
    Public ReadOnly Property AutoPaymentMaintenance() As Boolean
      Get
        AutoPaymentMaintenance = mvAutoPaymentMaintenance
      End Get
    End Property
    Public ReadOnly Property AddressMaintenance() As Boolean
      Get
        AddressMaintenance = mvAddressMaintenance
      End Get
    End Property
    Public ReadOnly Property CancelPaymentPlan() As Boolean
      Get
        CancelPaymentPlan = mvCancelPayPlan
      End Get
    End Property
    Public ReadOnly Property DefaultCancellationReason() As String
      Get
        DefaultCancellationReason = mvDefaultCancellationReason
      End Get
    End Property
    Public ReadOnly Property DefaultCancellationReasonDesc() As String
      Get
        DefaultCancellationReasonDesc = mvDefaultCancellationReasonDesc
      End Get
    End Property
    Public ReadOnly Property GiftAidDeclaration() As Boolean
      Get
        GiftAidDeclaration = mvGiftAidDeclaration
      End Get
    End Property

    Public ReadOnly Property MaintenanceOnly() As Boolean
      Get
        MaintenanceOnly = mvMaintenanceOnly
      End Get
    End Property

    Public ReadOnly Property ConfirmPayPlanDetails() As Boolean
      Get
        ConfirmPayPlanDetails = mvConfirmPayPlanDetails
      End Get
    End Property

    Public ReadOnly Property BlankMembershipJoinedDate() As Boolean
      Get
        BlankMembershipJoinedDate = mvBlankMembershipJoinedDate
      End Get
    End Property

    Public ReadOnly Property CollectionPayments() As Boolean
      Get
        CollectionPayments = mvCollectionPayments
      End Get
    End Property

    Public ReadOnly Property IncludeProvisionalTransactions() As Boolean
      Get
        IncludeProvisionalTransactions = mvIncludeProvisionalTrans
      End Get
    End Property

    Public ReadOnly Property IncludeConfirmedTransactions() As Boolean
      Get
        IncludeConfirmedTransactions = mvIncludeConfirmedTrans
      End Get
    End Property

    Public ReadOnly Property Voucher() As Boolean
      Get
        Voucher = mvVoucher
      End Get
    End Property

    Public ReadOnly Property MaximumVoucherTransactions() As Integer
      Get
        MaximumVoucherTransactions = mvMaxVoucherTransactions
      End Get
    End Property

    Public ReadOnly Property PayrollGiving() As Boolean
      Get
        PayrollGiving = mvPayrollGiving
      End Get
    End Property
    Public ReadOnly Property SourceFromLastMailing() As Boolean
      Get
        SourceFromLastMailing = mvSourceFromLastMailing
      End Get
    End Property
    Public ReadOnly Property BatchCategory() As String
      Get
        BatchCategory = mvBatchCategory
      End Get
    End Property
    Public ReadOnly Property BatchAnalysisCode() As String
      Get
        BatchAnalysisCode = mvBatchAnalysisCode
      End Get
    End Property
    Public ReadOnly Property EventMultipleAnalysis() As Boolean
      Get
        EventMultipleAnalysis = mvEventMultipleAnalysis
      End Get
    End Property
    Public ReadOnly Property DefaultTransactionOrigin() As String
      Get
        DefaultTransactionOrigin = mvDefaultTransactionOrigin
      End Get
    End Property
    Public ReadOnly Property LinkToFundraisingPayments() As Boolean
      Get
        Return mvLinkToFundraisingPayments
      End Get
    End Property

    Public ReadOnly Property ApplicationHeight() As Integer
      Get
        If mvAppHeight > 0 Then
          ApplicationHeight = mvAppHeight
        Else
          ApplicationHeight = STD_APP_HEIGHT
        End If
      End Get
    End Property

    Public ReadOnly Property ApplicationWidth() As Integer
      Get
        If mvAppWidth > 0 Then
          ApplicationWidth = mvAppWidth
        Else
          ApplicationWidth = STD_APP_WIDTH
        End If
      End Get
    End Property

    Public ReadOnly Property MailingCodeMandatory() As Boolean
      Get
        MailingCodeMandatory = mvMailingCodeMandatory
      End Get
    End Property

    Public Property Batch() As Batch
      Get
        Batch = mvBatch
      End Get
      Set(ByVal Value As Batch)
        mvBatch = Value
      End Set
    End Property

    Public ReadOnly Property BypassMailingParagraphs() As Boolean
      Get
        BypassMailingParagraphs = mvBypassMailingParagraphs
      End Get
    End Property

    Public ReadOnly Property SupportsNonFinancialBatch() As Boolean
      Get
        SupportsNonFinancialBatch = mvSupportsNonFinancialBatch
      End Get
    End Property

    Public ReadOnly Property SaleOrReturn() As Boolean
      Get
        SaleOrReturn = mvSaleOrReturn
      End Get
    End Property

    Public ReadOnly Property SalesQuantity() As String
      Get
        SalesQuantity = mvSalesQuantity
      End Get
    End Property

    Public ReadOnly Property UsesProductNumbers() As Boolean
      Get
        UsesProductNumbers = mvUsesProductNumbers
      End Get
    End Property

    Public Property NonFinancialBatch() As Batch
      Get
        If mvNonFinancialBatch Is Nothing Then InitNonFinancialTransaction()
        NonFinancialBatch = mvNonFinancialBatch
      End Get
      Set(ByVal Value As Batch)
        mvNonFinancialBatch = Value
      End Set
    End Property

    Public ReadOnly Property Controls() As TraderControls
      Get
        Controls = mvTraderControls
      End Get
    End Property

    Public ReadOnly Property Pages() As TraderPages
      Get
        Pages = mvTraderPages
      End Get
    End Property

    Public ReadOnly Property ConversionShowPPD() As Boolean
      Get
        'Note: If the new pp_conversion_incl_maintenance attribute is set then it takes precedence over this
        ConversionShowPPD = AppType = ApplicationType.atConversion And mvShowPPDInConversionApp
      End Get
    End Property

    Public ReadOnly Property CAFCardFloorLimit() As Double
      Get
        If Not mvCAFFloorLimitRead Then
          mvCAFFloorLimit = mvEnv.GetFloorLimit(CDBEnvironment.MailOrderControlTypes.moctCAF)
          mvCAFFloorLimitRead = True
        End If
        CAFCardFloorLimit = mvCAFFloorLimit
      End Get
    End Property

    Public ReadOnly Property CreditCardFloorLimit() As Double
      Get
        If Not mvNonCAFFloorLimitRead Then
          mvNonCAFFloorLimit = mvEnv.GetFloorLimit(CDBEnvironment.MailOrderControlTypes.moctNonCAF)
          mvNonCAFFloorLimitRead = True
        End If
        CreditCardFloorLimit = mvNonCAFFloorLimit
      End Get
    End Property

    Public ReadOnly Property ProvisionalPaymentPlan() As Boolean
      Get
        ProvisionalPaymentPlan = mvProvisionalPaymentPlan
      End Get
    End Property

    Public ReadOnly Property DefaultMemberBranch() As String
      Get
        DefaultMemberBranch = mvDefaultMemberBranch
      End Get
    End Property

    Public ReadOnly Property CancelGiftAidDeclaration() As Boolean
      Get
        CancelGiftAidDeclaration = mvGiftAidCancellation
      End Get
    End Property

    Public ReadOnly Property ProvisionalCashTransactionDocument() As String
      Get
        ProvisionalCashTransactionDocument = mvProvisionalCashTransDoc
      End Get
    End Property

    Public ReadOnly Property PreFulfilledIncentives() As Boolean
      Get
        PreFulfilledIncentives = mvPreFulfilledIncentives
      End Get
    End Property

    Public ReadOnly Property ContactAlerts() As Boolean
      Get
        Return mvContactAlerts
      End Get
    End Property

    Public ReadOnly Property CreditSalesReferenceOptional() As Boolean
      Get
        CreditSalesReferenceOptional = mvOptionalCreditSalesReference
      End Get
    End Property

    Public ReadOnly Property DisplayScheduledPayments() As Boolean
      Get
        DisplayScheduledPayments = mvDisplayScheduledPayments
      End Get
    End Property

    Public ReadOnly Property ConfirmSaleOrReturnTransactions() As Boolean
      Get
        ConfirmSaleOrReturnTransactions = mvConfirmSRTransactions
      End Get
    End Property

    Public ReadOnly Property OnlineCCAuthorisation() As Boolean
      Get
        OnlineCCAuthorisation = mvOnLineCCAuthorisation
      End Get
    End Property

    Public ReadOnly Property RequireCCAuthorisation() As Boolean
      Get
        Return mvRequireCCAuthorisation
      End Get
    End Property

    Public ReadOnly Property PackToDonorDefault() As Boolean
      Get
        Return BooleanValue(mvPackToDonorDefault)
      End Get
    End Property

    Public ReadOnly Property PayPlanConversionMaintenance() As Boolean
      Get
        'Note: Extends (and takes precedence over) the use of the trader_conversion_app_show_ppd config
        PayPlanConversionMaintenance = (AppType = ApplicationType.atConversion And mvPayPlanConvMaintenance = True)
      End Get
    End Property

    Public ReadOnly Property PaymentPlanDetails() As TraderPaymentPlanDetails
      Get
        PaymentPlanDetails = mvPaymentPlanDetails
      End Get
    End Property

    Public ReadOnly Property ServiceBookingAnalysis() As Boolean
      Get
        ServiceBookingAnalysis = mvServiceBookingAnalysis
      End Get
    End Property

    Public ReadOnly Property AlbacsBankDetails() As String
      Get
        If Len(mvAlbacsBankDetails) = 0 Then mvAlbacsBankDetails = "C"
        AlbacsBankDetails = mvAlbacsBankDetails
      End Get
    End Property

    Public ReadOnly Property CSDepositPercentage() As Double
      Get
        CSDepositPercentage = mvCSDepositPercentage
      End Get
    End Property

    Public Property StockMovementTransactionID() As Integer
      Get
        StockMovementTransactionID = mvStockMovementTransactionID
      End Get
      Set(ByVal Value As Integer)
        mvStockMovementTransactionID = Value
      End Set
    End Property

    Public ReadOnly Property CheckIncentives() As Boolean
      Get
        CheckIncentives = mvCheckIncentives
      End Get
    End Property

    Public ReadOnly Property Loans() As Boolean
      Get
        Return mvLoans
      End Get
    End Property

    Public Property TraderAlerts() As Boolean
      Get
        Return mvTraderAlerts
      End Get
      Private Set(value As Boolean)
        mvTraderAlerts = value
      End Set
    End Property

    Private ReadOnly Property UseToken As Boolean?
      Get
        If mvUseToken Is Nothing Then
          Dim vString As String = String.Empty
          If String.IsNullOrWhiteSpace(mvBatchCategory) Then
            vString = "Select use_tokens from merchant_details where merchant_retail_number = '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMerchantRetailNumber) & "'"
          Else
            vString = "Select use_tokens from merchant_details md join batch_categories bc on md.merchant_retail_number = bc.merchant_retail_number where bc.batch_category = '" & mvBatchCategory & "'"
          End If
          Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet(vString)
          If vRecordSet IsNot Nothing Then
            With vRecordSet
              If .Fetch() = True Then
                mvValid = True
                With .Fields
                  mvUseToken = .Item("use_tokens").Bool
                End With
              Else
                mvUseToken = False
              End If
            End With
          Else
            mvUseToken = False
          End If
          vRecordSet.CloseRecordSet()
        End If
        Return mvUseToken
      End Get
    End Property

    Public ReadOnly Property AutoCreateCreditCustomer() As Boolean
      Get
        Return mvAutoCreateCreditCustomer
      End Get
    End Property

    Public ReadOnly Property UnpostedBatchMsgInPrint() As Boolean
      Get
        Return mvUnpostedBatchMsgInPrint
      End Get
    End Property

    Public ReadOnly Property DateRangeMsgInPrint() As Boolean
      Get
        Return mvDateRangeMsgInPrint
      End Get
    End Property

    Public ReadOnly Property CreditCategory As String
      Get
        Return mvCreditCategory
      End Get
    End Property

    Public ReadOnly Property DefaultExamSessionCode() As String
      Get
        Return mvDefaultExamSessionCode
      End Get
    End Property

    Public ReadOnly Property DefaultExamUnitCode() As String
      Get
        Return mvDefaultExamUnitCode
      End Get
    End Property

    Public ReadOnly Property AutoGiftAidDeclaration() As Boolean
      Get
        Return mvAutoGiftAidDeclaration
      End Get
    End Property

    Public ReadOnly Property AutoGiftAidMethod() As String
      Get
        Return mvAutoGiftAidMethod
      End Get
    End Property
    Public ReadOnly Property AutoGiftAidSource() As String
      Get
        Return mvAutoGiftAidSource
      End Get
    End Property
    Public Property MerchantRetailNumber() As String
      Get
        Dim vDataTable As CDBDataTable = Nothing
        Dim vWhereFields As New CDBFields()
        If String.IsNullOrEmpty(mvMerchantRetailNumber) Then
          If mvBatchCategory.Length > 0 Then
            vWhereFields.Add("batch_category", mvBatchCategory)
            vDataTable = New CDBDataTable(mvEnv, New SQLStatement(mvEnv.Connection, "merchant_retail_number", "batch_categories", vWhereFields))
            If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
              mvMerchantRetailNumber = vDataTable.Rows(0).Item("merchant_retail_number")
            End If
          End If
          If String.IsNullOrEmpty(mvMerchantRetailNumber) Then
            mvMerchantRetailNumber = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMerchantRetailNumber)
            ''If vMerchantRetailNumber.Length = 0 Then IncorrectSetup()
          End If
        End If
        Return mvMerchantRetailNumber
      End Get
      Private Set(ByVal value As String)
        mvMerchantRetailNumber = value
      End Set
    End Property

    Friend ReadOnly Property PrintInvoiceUnpostedBatches() As Boolean
      Get
        Return (AppType = ApplicationType.atBatchInvoiceGeneration AndAlso mvInvoicePrintUnpostedBatches = True)
      End Get
    End Property

    Public Function GetPageCode(ByVal pPageType As TraderPage.TraderPageType) As String
      Dim vPageCode As String = ""

      Select Case pPageType
        Case TraderPage.TraderPageType.tpPaymentMethod1
          vPageCode = "PM1" 'Payment method choice 1
        Case TraderPage.TraderPageType.tpCreditCustomer
          vPageCode = "CCU" 'Credit customer
        Case TraderPage.TraderPageType.tpTransactionDetails
          vPageCode = "TRD" 'Transaction details
        Case TraderPage.TraderPageType.tpComments
          vPageCode = "COM" 'Comments
        Case TraderPage.TraderPageType.tpBankDetails
          vPageCode = "BKD" 'Bank Details
        Case TraderPage.TraderPageType.tpCardDetails
          vPageCode = "CDC" 'Credit or Debit Card details
        Case TraderPage.TraderPageType.tpTransactionAnalysis
          vPageCode = "TRA" 'Transaction Analysis choice
        Case TraderPage.TraderPageType.tpPaymentMethod2
          vPageCode = "PM2" 'Payment method choice 2
        Case TraderPage.TraderPageType.tpPaymentMethod3
          vPageCode = "PM3" 'Payment Method 3; order conversion type
        Case TraderPage.TraderPageType.tpProductDetails
          vPageCode = "PRD" 'Product details
        Case TraderPage.TraderPageType.tpPayments
          vPageCode = "PAY" 'Payments
        Case TraderPage.TraderPageType.tpPaymentPlanDetails
          vPageCode = "PPD" 'Payment Plan details
        Case TraderPage.TraderPageType.tpPaymentPlanProducts
          vPageCode = "PPP" 'Payment Plan products
        Case TraderPage.TraderPageType.tpStandingOrder
          vPageCode = "STO" 'Standing Order
        Case TraderPage.TraderPageType.tpDirectDebit
          vPageCode = "DDR" 'Direct Debit
        Case TraderPage.TraderPageType.tpCreditCardAuthority
          vPageCode = "CCA" 'Continuous Credit Card Authority
        Case TraderPage.TraderPageType.tpMembership
          vPageCode = "MEM" 'Membership
        Case TraderPage.TraderPageType.tpAmendMembership
          vPageCode = "AMD" 'Amend Membership Details
        Case TraderPage.TraderPageType.tpChangeMembershipType
          vPageCode = "CMT" 'Change Membership Type  (CMT)
        Case TraderPage.TraderPageType.tpMembershipPayer
          vPageCode = "MSP" 'Membership Payer  (CMT)
        Case TraderPage.TraderPageType.tpCovenant
          vPageCode = "COV" 'Covenant
        Case TraderPage.TraderPageType.tpContactSelection
          vPageCode = "CSE" 'Contact Selection
        Case TraderPage.TraderPageType.tpEventBooking
          vPageCode = "EVE" 'Event Booking
        Case TraderPage.TraderPageType.tpExamBooking
          vPageCode = "EXA" 'Exam Booking
        Case TraderPage.TraderPageType.tpAccommodationBooking
          vPageCode = "ACO" 'Accomodation
        Case TraderPage.TraderPageType.tpPurchaseInvoiceDetails
          vPageCode = "PID" 'Purchase Invoice Details
        Case TraderPage.TraderPageType.tpPurchaseInvoiceProducts
          vPageCode = "PIP" 'Purchase Invoice Products
        Case TraderPage.TraderPageType.tpPurchaseOrderDetails
          vPageCode = "POD" 'Purchase Order Details
        Case TraderPage.TraderPageType.tpPurchaseOrderProducts
          vPageCode = "POP" 'Purchase Order Products
        Case TraderPage.TraderPageType.tpPurchaseOrderPayments
          vPageCode = "PPA" 'Purchase Order Payments
        Case TraderPage.TraderPageType.tpPurchaseOrderCancellation
          vPageCode = "POC" 'Purchase Order Cancellation
        Case TraderPage.TraderPageType.tpChequeNumberAllocation
          vPageCode = "CNA" 'Cheque Number Allocation
        Case TraderPage.TraderPageType.tpChequeReconciliation
          vPageCode = "CRE" 'Cheque Reconciliation
        Case TraderPage.TraderPageType.tpPostageAndPacking
          vPageCode = "PAP" 'Postage & Packing - Carriage
        Case TraderPage.TraderPageType.tpServiceBooking
          vPageCode = "SVC" 'Service Booking
        Case TraderPage.TraderPageType.tpCreditStatementGeneration
          vPageCode = "CSG" 'Credit Statement Generation
        Case TraderPage.TraderPageType.tpBatchInvoiceProduction
          vPageCode = "ING" 'Batch Invoice Generation
        Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          vPageCode = "TPP" 'Create Payment Plan on Unbalance Transaction
        Case TraderPage.TraderPageType.tpPaymentPlanMaintenance
          vPageCode = "PPM" 'Payment Plan Maintenance
        Case TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
          vPageCode = "PPN" 'Payment Plan Details Maintenance
        Case TraderPage.TraderPageType.tpSetStatus
          vPageCode = "STA" 'Set Status
        Case TraderPage.TraderPageType.tpCancelPaymentPlan
          vPageCode = "CPP" 'Cancel Payment Plan
        Case TraderPage.TraderPageType.tpLegacyBequestReceipt
          vPageCode = "LBR" 'Legacy Bequest Receipt
        Case TraderPage.TraderPageType.tpActivityEntry
          vPageCode = "ACT" 'Activity Entry
        Case TraderPage.TraderPageType.tpGiveAsYouEarn
          vPageCode = "GYP" 'Give As You Earn
        Case TraderPage.TraderPageType.tpGiftAidDeclaration
          vPageCode = "GAD"
        Case TraderPage.TraderPageType.tpGiveAsYouEarnEntry
          vPageCode = "GYE"
        Case TraderPage.TraderPageType.tpGoneAway 'Gone Away processing
          vPageCode = "GAW"
        Case TraderPage.TraderPageType.tpAddressMaintenance 'Address Maintenance
          vPageCode = "ADM"
        Case TraderPage.TraderPageType.tpSuppressionEntry 'Suppression Entry
          vPageCode = "SUP"
        Case TraderPage.TraderPageType.tpCancelGiftAidDeclaration
          vPageCode = "CGA" 'Cancel Gift Aid Declaration
        Case TraderPage.TraderPageType.tpScheduledPayments 'Display Scheduled Payments
          vPageCode = "SCP"
        Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vPageCode = "OSP" 'Choose Scheduled Payment to pay
        Case TraderPage.TraderPageType.tpConfirmProvisionalTransactions
          vPageCode = "CPT"
        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary 'PG_TAS   Summary only No Controls
          vPageCode = "TAS"
        Case TraderPage.TraderPageType.tpPaymentPlanSummary 'PG_PPS   Summary only No Controls
          vPageCode = "PPS"
        Case TraderPage.TraderPageType.tpStatementList 'PG_STL   Summary only No Controls
          vPageCode = "STL"
        Case TraderPage.TraderPageType.tpInvoicePayments 'PG_INV   Summary only No Controls
          vPageCode = "INV"
        Case TraderPage.TraderPageType.tpMembershipMembersSummary 'PG_MMS   Summary only No Controls
          vPageCode = "MMS"
        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary 'PG_PIS   PINVE Summary only No Controls
          vPageCode = "PIS"
        Case TraderPage.TraderPageType.tpPurchaseOrderSummary 'PG_POS   PORDE Summary only No Controls
          vPageCode = "POS"
        Case TraderPage.TraderPageType.tpPostTaxPGPayment 'Post Tax Payroll Giving payments only
          vPageCode = "PGP"
        Case TraderPage.TraderPageType.tpCollectionPayments
          vPageCode = "PCP"
        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          vPageCode = "INS"
        Case TraderPage.TraderPageType.tpAmendEventBooking
          vPageCode = "AEV"
        Case TraderPage.TraderPageType.tpLoans
          vPageCode = "LON"
        Case TraderPage.TraderPageType.tpAdvancedCMT
          vPageCode = "MTC"
        Case TraderPage.TraderPageType.tpTokenSelection
          vPageCode = "TKN"
      End Select
      GetPageCode = vPageCode
    End Function

    Public Function DefaultProductCode(Optional ByVal pGetLast As Boolean = False) As String
      If mvProductFromSource.Length > 0 Then
        Return mvProductFromSource
      ElseIf mvProduct.Length > 0 Then
        Return mvProduct
      ElseIf pGetLast Then
        Return LastProduct
      Else
        Return ""
      End If
    End Function
    Public Function DefaultEventNumber() As String
      DefaultEventNumber = LastEventNumber
    End Function
    Public Function DefaultBookingOption() As String
      DefaultBookingOption = LastBookingOptionNumber
    End Function
    Public Function DefaultBookingRate() As String
      DefaultBookingRate = LastEventBookingRate
    End Function
    Public Function DefaultRateCode(Optional ByVal pGetLast As Boolean = False) As String
      If Len(mvRateFromSource) > 0 Then
        Return mvRateFromSource
      ElseIf Len(mvRate) > 0 Then
        Return mvRate
      ElseIf pGetLast Then
        Return LastRate
      Else
        Return ""
      End If
    End Function

    Public Sub SetProductFromSource(ByRef pProduct As String, ByRef pRate As String, ByRef pUsesProductNumbers As Boolean)
      mvProductFromSource = pProduct
      mvRateFromSource = pRate
      mvProductNumbersFromSource = pUsesProductNumbers
    End Sub
    Public Sub Init(ByVal pApplication As String, ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransNumber As Integer, Optional ByVal pBatchType As String = "", Optional ByVal pDesignMode As Boolean = False, Optional ByVal pFinancialAdjustment As Batch.AdjustmentTypes = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vCount As Integer
      Dim vIndex As Integer

      mvEnv = pEnv
      mvApplication = pApplication
      mvBatchNumber = pBatchNumber
      mvTransNumber = pTransNumber
      mvBatchType = pBatchType

      If Len(mvApplication) > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT * FROM fp_applications WHERE fp_application = '" & mvApplication & "'")
        With vRecordSet
          If .Fetch() = True Then
            mvValid = True
            With .Fields
              mvAppDesc = .Item("fp_application_desc").Value
              mvBatchLedApp = .Item("batch_application").Bool
              mvCash = .Item("pm_cash").Bool
              mvCheque = .Item("pm_cheque").Bool
              mvPostalOrder = .Item("pm_postal_order").Bool
              mvCreditCard = .Item("pm_credit_card").Bool
              mvDebitCard = .Item("pm_debit_card").Bool
              If mvEnv.GetConfig("fp_card_sales_combined_claim") = "A" Or mvEnv.GetConfig("fp_card_sales_combined_claim") = "Y" Then mvDebitCard = False
              mvCreditSales = .Item("pm_credit").Bool
              mvChequeWithInvoice = .FieldExists("cheque_with_invoice").Bool
              mvCCWithInvoice = .FieldExists("cc_with_invoice").Bool
              mvBankDetails = .Item("bank_details").Bool
              mvTransactionComments = .Item("transaction_comments").Bool
              mvMemberships = .Item("memberships").Bool
              mvSubscriptions = .Item("subscriptions").Bool
              mvDonationsRegular = .Item("donations_regular").Bool
              mvProductSales = .Item("product_sales").Bool
              mvOneOffDonations = .Item("donations_one_off").Bool
              mvPayments = .Item("payments").Bool
              mvStandingOrders = .Item("standing_orders").Bool
              mvDirectDebits = .Item("direct_debits").Bool
              mvCreditCardAuthorities = .Item("credit_card_authorities").Bool
              mvCovMemberships = .Item("covenant_membership").Bool
              mvCovSubscriptions = .Item("covenant_subscription").Bool
              mvCovDonationsRegular = .Item("covenant_donation_regular").Bool
              mvEventBooking = .Item("event_booking").Bool
              mvExamBooking = .FieldExists("exam_booking").Bool
              mvAccomodationBooking = .Item("accomodation_booking").Bool
              mvProduct = .Item("product").Value
              mvRate = .Item("rate").Value
              mvSource = .Item("source").Value
              mvCABankAccount = .Item("ca_bank_account").Value
              mvCCBankAccount = .Item("cc_bank_account").Value
              mvDCBankAccount = .Item("dc_bank_account").Value
              mvCSBankAccount = .Item("cs_bank_account").Value
              mvSOBankAccount = .Item("so_bank_account").Value
              mvDDBankAccount = .Item("dd_bank_account").Value
              mvCCABankAccount = .Item("cca_bank_account").Value
              mvShowTransReference = .Item("show_transaction_reference").Bool
              mvConfirmDefaultProduct = .Item("confirm_default_product").Bool
              mvConfirmAnalysis = .Item("confirm_analysis").Bool
              mvAppTypeCode = .Item("fp_application_type").Value
              mvAppType = GetApplicationType(mvAppTypeCode)
              If .Exists("carriage") Then
                mvCarriage = .Item("carriage").Bool
                mvConfirmCarriage = .Item("confirm_carriage").Bool
                mvCarriageProduct = .Item("carriage_product").Value
                mvCarriageRate = .Item("carriage_rate").Value
                mvCarriagePercentage = Val(.Item("carriage_percentage").Value)
                mvAnalysisComments = .Item("analysis_comments").Bool
              End If
              mvNoPaymentRequired = .FieldExists("non_paid_payment_plans").Bool
              mvServiceBookings = .FieldExists("service_bookings").Bool
              mvPayMethodsAtEnd = .FieldExists("pay_methods_at_end").Bool
              mvAutoSetAmount = .FieldExists("auto_set_amount").Bool
              mvDefaultSalesContact = IntegerValue(.FieldExists("default_sales_contact").Value)
              mvInvoicePayments = .FieldExists("invoice_payments").Bool
              mvPayPlanPayMethod = .FieldExists("pay_plan_pay_method").Bool
              mvCreditNotes = .FieldExists("sundry_credit_notes").Bool
              mvForeignCurrency = .FieldExists("foreign_currency").Bool
              mvInvoiceDocument = .FieldExists("invoice_document").Value
              mvReceiptDocument = .FieldExists("receipt_document").Value
              mvPayPlanDocument = .FieldExists("payment_plan_document").Value
              mvMembersOnly = .FieldExists("members_only").Bool
              mvSalesGroup = .FieldExists("sales_group").Value
              mvChangeMembership = .FieldExists("change_membership").Bool
              mvCreditStatementDocument = .FieldExists("credit_statement_document").Value
              mvServiceBookingCredits = .FieldExists("service_booking_credits").Bool
              mvLegacyReceipt = .FieldExists("legacy_receipts").Bool
              mvGoneAway = .FieldExists("set_gone_away").Bool
              mvSetStatus = .FieldExists("set_status").Bool
              mvDefaultStatus = .FieldExists("status").Value
              mvAddActivity = .FieldExists("add_activity").Bool
              mvDefaultActivityGroup = .FieldExists("activity_group").Value
              mvAddSuppression = .FieldExists("add_suppression").Bool
              mvDefaultSuppression = .FieldExists("mailing_suppression").Value
              mvCancelPayPlan = .FieldExists("cancel_payment_plan").Bool
              mvDefaultCancellationReason = .FieldExists("cancellation_reason").Value
              mvGiftAidDeclaration = .FieldExists("gift_aid_declaration").Bool
              mvConfirmPayPlanDetails = .FieldExists("confirm_details").Bool
              mvLinkToCommunication = .FieldExists("link_to_communication").Value
              mvDistributionCodeMandatory = .FieldExists("distribution_code_mandatory").Bool
              mvSalesContactMandatory = .FieldExists("sales_contact_mandatory").Bool
              mvAppHeight = .FieldExists("application_height").IntegerValue
              mvAppWidth = .FieldExists("application_width").IntegerValue
              'Set MaintenanceOnly property
              mvMaintenanceOnly = False
              For vIndex = 1 To .Count
                Select Case .Item(vIndex).Name
                  'maintenance options
                  Case "set_gone_away", "set_status", "add_activity", "add_suppression", "cancel_payment_plan", "gift_aid_declaration", "payroll_giving", "gift_aid_cancellation", "auto_payment_maintenance", "address_maintenance"
                    If .Item(vIndex).Bool And vCount = 0 Then mvMaintenanceOnly = True
                    'other analysis options
                  Case "memberships", "subscriptions", "donations_regular", "product_sales", "donations_one_off", "payments", "covenant_membership", "covenant_subscription", "covenant_donation_regular", "event_booking", "exam_booking", "accomodation_booking", "service_bookings", "invoice_payments", "sundry_credit_notes", "change_membership", "service_booking_credits", "legacy_receipt"
                    If .Item(vIndex).Bool Then
                      vCount = vCount + 1
                      mvMaintenanceOnly = False
                    End If
                End Select
              Next
              mvIncludeConfirmedTrans = .FieldExists("include_confirmed_trans").Bool
              mvIncludeProvisionalTrans = .FieldExists("include_provisional_trans").Bool
              mvVoucher = .FieldExists("pm_voucher").Bool
              mvCAFCard = .FieldExists("pm_caf_card").Bool
              mvGiftInKind = .FieldExists("pm_gift_in_kind").Bool
              mvCVBankAccount = .FieldExists("cv_bank_account").Value
              mvDonationProduct = .FieldExists("donation_product").Value
              mvDonationRate = .FieldExists("donation_rate").Value
              mvPayrollGiving = .FieldExists("payroll_giving").Bool
              mvBatchCategory = .FieldExists("batch_category").Value
              mvSourceFromLastMailing = .FieldExists("source_from_last_mailing").Bool
              mvMailingCodeMandatory = .FieldExists("mailing_code_mandatory").Bool
              mvBypassMailingParagraphs = .FieldExists("bypass_mailing_paragraphs").Bool
              mvSupportsNonFinancialBatch = .FieldExists("non_financial_batch").Bool
              mvAddressMaintenance = .FieldExists("address_maintenance").Bool
              mvAutoPaymentMaintenance = .FieldExists("auto_payment_maintenance").Bool
              mvGiftAidCancellation = .FieldExists("gift_aid_cancellation").Bool
              mvSaleOrReturn = .FieldExists("pm_sale_or_return").Bool
              mvProvisionalPaymentPlan = .FieldExists("provisional_payment_plan").Bool
              mvDefaultMemberBranch = .FieldExists("default_member_branch").Value
              mvProvisionalCashTransDoc = .FieldExists("provisional_cash_document").Value
              mvPreFulfilledIncentives = .FieldExists("prefulfilled_incentives").Bool
              mvContactAlerts = .FieldExists("contact_alerts").Bool
              mvDisplayScheduledPayments = .FieldExists("display_scheduled_payments").Bool
              mvConfirmSRTransactions = .FieldExists("confirm_sr_transactions").Bool
              mvOnLineCCAuthorisation = .FieldExists("online_cc_authorisation").Bool
              mvRequireCCAuthorisation = .FieldExists("require_cc_authorisation").Bool
              mvPayPlanConvMaintenance = .FieldExists("pp_conversion_incl_maintenance").Bool
              mvCollectionPayments = .FieldExists("collection_payments").Bool
              mvBatchAnalysisCode = .FieldExists("batch_analysis_code").Value
              mvCampaign = .FieldExists("campaign").Value
              mvAppeal = .FieldExists("appeal").Value
              mvEventMultipleAnalysis = .FieldExists("event_multiple_analysis").Bool
              mvDefaultTransactionOrigin = .FieldExists("transaction_origin").Value
              mvServiceBookingAnalysis = .FieldExists("service_booking_analysis").Bool
              mvAlbacsBankDetails = .FieldExists("albacs_bank_details").Value
              mvLinkToFundraisingPayments = .FieldExists("link_to_fundraising_payments").Bool
              mvInvoicePrintPreviewDefault = .FieldExists("invoice_print_preview_default").Bool
              mvLoans = .FieldExists("loans").Bool
              mvAutoCreateCreditCustomer = .FieldExists("auto_create_credit_customer").Bool
              mvUnpostedBatchMsgInPrint = .FieldExists("unposted_batch_msg_in_print").Bool
              mvDateRangeMsgInPrint = .FieldExists("date_range_msg_in_print").Bool
              mvCreditCategory = .FieldExists("credit_category").Value
              mvDefaultExamSessionCode = .FieldExists("exam_session_code").Value
              mvDefaultExamUnitCode = .FieldExists("exam_unit_code").Value
              mvInvoicePrintUnpostedBatches = .FieldExists("invoice_print_unposted_batches").Bool
              mvAutoGiftAidDeclaration = .FieldExists("auto_gift_aid_declaration").Bool
              mvAutoGiftAidMethod = .FieldExists("auto_gift_aid_method").Value
              mvAutoGiftAidSource = .FieldExists("auto_gift_aid_source").Value
              MerchantRetailNumber = .FieldExists("merchant_retail_number").Value
              'See if was have any alerts
              TraderAlerts = False
              Dim vTraderAppNumber As Integer
              Integer.TryParse(mvApplication, vTraderAppNumber)
              Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "", "contact_alert_links", New CDBFields(New CDBField("trader_application_number", vTraderAppNumber)), "")
              If mvEnv.Connection.GetCountFromStatement(vSQLStatement) > 0 Then
                TraderAlerts = True
              End If
            End With
          Else
            'we never expect this to happen!
            mvValid = False
          End If
        End With
        vRecordSet.CloseRecordSet()

        If mvValid Then
          'If this is a batch led application then get the batch to use
          If mvBatchLedApp Then
            If mvBatchNumber > 0 Then
              mvBatch = New Batch(mvEnv)
              mvBatch.Init(mvBatchNumber)
              With mvBatch
                If .Existing Then
                  mvProduct = If(Len(.ProductCode) > 0, .ProductCode, mvProduct)
                  mvRate = If(Len(.RateCode) > 0, .RateCode, mvRate)
                  mvSource = If(Len(.Source) > 0, .Source, mvSource)
                  If Len(.Campaign) > 0 Then
                    mvCampaign = .Campaign
                    mvAppeal = .Appeal
                  End If
                  mvCABankAccount = .BankAccount
                  mvCCBankAccount = .BankAccount
                  mvDCBankAccount = .BankAccount
                  mvCSBankAccount = .BankAccount
                  If .BatchType = Batch.BatchTypes.CreditCard OrElse .BatchType = Batch.BatchTypes.CreditCardWithInvoice Then 'For credit or debit card batches make sure the right options are set
                    mvCreditCard = True
                    mvDebitCard = False
                  ElseIf .BatchType = Batch.BatchTypes.DebitCard Then
                    mvCreditCard = False
                    mvDebitCard = True
                  End If
                  mvBatchType = Access.Batch.GetBatchTypeCode(.BatchType)
                  mvBatchDate = .BatchDate
                  mvPostedToCashBook = .PostedToCashBook
                  mvPostedToNominal = .PostedToNominal
                Else
                  mvValid = False 'we never expect this to happen
                End If
              End With
            End If
          ElseIf pFinancialAdjustment = Batch.AdjustmentTypes.atMove And Len(mvBatchType) > 0 Then
            mvCash = False
            mvCheque = False
            mvPostalOrder = False
            mvCreditCard = False
            mvDebitCard = False
            mvCreditSales = False
            Select Case mvBatchType
              Case Access.Batch.GetBatchTypeCode(Batch.BatchTypes.Cash), Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CashWithInvoice)
                mvCash = True
                mvCheque = True
                mvPostalOrder = True
              Case Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditCard), Access.Batch.GetBatchTypeCode(Batch.BatchTypes.DebitCard), Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditCardWithInvoice)
                mvCreditCard = True
                mvDebitCard = True
              Case Access.Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment)
                mvCash = True
                mvCheque = True
                mvPostalOrder = True
                mvCreditCard = True
                mvDebitCard = True
              Case Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditSales)
                mvCreditSales = True
            End Select
            mvBatchType = mvEnv.Connection.GetValue("SELECT batch_type FROM batches WHERE batch_number = " & mvBatchNumber)
          End If

          If BatchNumber > 0 And TransactionNumber > 0 Then
            Select Case pFinancialAdjustment
              Case Batch.AdjustmentTypes.atMove, Batch.AdjustmentTypes.atAdjustment, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
                mvExistingAdjustmentTran = True
              Case Else
                mvExistingAdjustmentTran = False
            End Select
          End If

          If pFinancialAdjustment <> Batch.AdjustmentTypes.atNone And mvBatchNumber > 0 And mvBatch Is Nothing Then
            'Needed for Financial Adjustments
            mvBatch = New Batch(mvEnv)
            mvBatch.Init(mvBatchNumber)
            If (pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or pFinancialAdjustment = Batch.AdjustmentTypes.atMove) And Len(mvBatchType) = 0 Then
              mvBatchType = Access.Batch.GetBatchTypeCode(mvBatch.AdjustmentBatchType(pFinancialAdjustment))
            End If
          End If

          'Set up some page info about the app
          Select Case mvAppType
            Case ApplicationType.atTransaction, ApplicationType.atConversion, ApplicationType.atBankStatementPosting 'TRANS or CNVRT
              mvMainPage = TraderPage.TraderPageType.tpPaymentMethod1 'PG_PM1
              mvSummaryPage = TraderPage.TraderPageType.tpPaymentPlanSummary 'PG_PPS
              mvAmountTag = "Payment Plan Amount"
            Case ApplicationType.atPurchaseInvoice 'PINVE
              mvMainPage = TraderPage.TraderPageType.tpPurchaseInvoiceDetails 'PG_PID
              mvLinePage = TraderPage.TraderPageType.tpPurchaseInvoiceProducts 'PG_PIP
              mvSummaryPage = TraderPage.TraderPageType.tpPurchaseInvoiceSummary 'PG_PIS
              mvAmountTag = "Amount"
            Case ApplicationType.atPurchaseOrder 'PORDE
              mvMainPage = TraderPage.TraderPageType.tpPurchaseOrderDetails 'PG_POD
              mvLinePage = TraderPage.TraderPageType.tpPurchaseOrderProducts 'PG_POP
              mvSummaryPage = TraderPage.TraderPageType.tpPurchaseOrderSummary 'PG_POS
            Case ApplicationType.atPurchaseOrderCancellation 'PORDC
              mvMainPage = TraderPage.TraderPageType.tpPurchaseOrderCancellation 'PG_POC
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atChequeNumberAllocation 'CHQNA
              mvMainPage = TraderPage.TraderPageType.tpChequeNumberAllocation 'PG_CNA
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atChequeReconciliation 'CHQRE
              mvMainPage = TraderPage.TraderPageType.tpChequeReconciliation 'PG_CRE
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atCreditStatementGeneration 'CSTAT
              mvMainPage = TraderPage.TraderPageType.tpCreditStatementGeneration 'PG_CSG
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atBatchInvoiceGeneration 'BINVG
              mvMainPage = TraderPage.TraderPageType.tpBatchInvoiceProduction 'PG_ING
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atMaintenance 'MAINT
              mvMainPage = TraderPage.TraderPageType.tpContactSelection 'PG_CSE  'PG_PPM
              mvSummaryPage = TraderPage.TraderPageType.tpPaymentPlanSummary 'PG_PPS
              mvAmountTag = "Payment Plan Balance"
              mvTotalTag = "Current Line Balance"
            Case ApplicationType.atCreditListReconciliation
              If mvBatchNumber > 0 And mvTransNumber > 0 Then
                mvMainPage = TraderPage.TraderPageType.tpPaymentMethod1 'PG_PM1
              Else
                mvMainPage = TraderPage.TraderPageType.tpStatementList
              End If
              mvSummaryPage = TraderPage.TraderPageType.tpPaymentPlanSummary 'PG_PPS
              mvAmountTag = "Payment Plan Amount"
            Case ApplicationType.atGiveAsYouEarnPayments
              mvMainPage = TraderPage.TraderPageType.tpGiveAsYouEarn
              mvLinePage = 0
              mvSummaryPage = 0
            Case ApplicationType.atPostTaxPGPayments
              mvMainPage = TraderPage.TraderPageType.tpPostTaxPGPayment
              mvLinePage = 0
              mvSummaryPage = 0
          End Select

          If pDesignMode Then
            mvDiscountActivity = mvEnv.GetConfig("fp_discount_activity")
          Else
            GetDefaults()
            If pFinancialAdjustment = Batch.AdjustmentTypes.atMove And Not mvConfirmAnalysis Then mvConfirmAnalysis = True
          End If
        End If

        mvPaymentPlanDetails = New TraderPaymentPlanDetails

      Else
        'No Application number passed
        mvValid = False
      End If
    End Sub

    Public Sub ValidateTraderApplication(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pFinancialAdjustment As Batch.AdjustmentTypes)
      'Validation checks to ensure that Application is valid
      Dim vIndex As Integer
      Dim vSQL As String
      Dim vBatchType As String
      Dim vBankAccount As String = ""
      Dim vValid As Boolean
      Dim vErrorMsg As String

      '---------------------------------------------------------------------------------------------------------
      'Check if transaction payment method is available for the Financial Adjustment Trader Application
      '---------------------------------------------------------------------------------------------------------
      If pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then ValidatePaymentMethod(pBatchNumber, pTransactionNumber)

      '-----------------------------------------------------
      '1.  Check for suspense accounts for all BankAccounts
      '-----------------------------------------------------
      'Check to see if the default bank accounts have the corresponding suspense account
      vValid = True
      vSQL = "batch_type = '%1' AND bank_account = '%2'"
      For vIndex = 1 To 15
        vBatchType = ""
        Select Case vIndex
          Case 1 'CA Batch Type
            If Cash Or Cheque Or PostalOrder Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.Cash)
              vBankAccount = mvCABankAccount
            End If
          Case 2 'CC Batch Type
            If CreditCard Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditCard)
              vBankAccount = mvCCBankAccount
            End If
          Case 3 'DC Batch Type
            If DebitCard Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.DebitCard)
              vBankAccount = mvDCBankAccount
            End If
          Case 4 'CS Batch Type
            If CreditSales Or CreditNotes Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditSales)
              vBankAccount = mvCSBankAccount
            End If
          Case 5 'SO Batch Type
            If StandingOrders Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.StandingOrder)
              vBankAccount = mvSOBankAccount
            End If
          Case 6 'DD Batch Type
            If DirectDebits Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.DirectDebit)
              vBankAccount = mvDDBankAccount
            End If
          Case 7 'CO Batch Type
            If CreditCardAuthorities Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditCardAuthority)
              vBankAccount = mvCCABankAccount
            End If
          Case 8 'CV Batch Type
            If Voucher Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CAFVouchers)
              vBankAccount = mvCVBankAccount
            End If
          Case 9 'CF Batch Type
            If CAFCard Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CAFCards)
              vBankAccount = mvCVBankAccount
            End If
          Case 10 'GK Batch Type
            If GiftInKind Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.GiftInKind)
              vBankAccount = mvCABankAccount
            End If
          Case 11 'GP Batch Type
            If mvEnv.GetConfigOption("option_gaye") Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.GiveAsYouEarn)
              vBankAccount = mvSOBankAccount
            End If
          Case 12 'SR Batch Type
            If SaleOrReturn Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.SaleOrReturn)
              vBankAccount = mvCABankAccount
            End If
          Case 13 'PG Batch Type
            If mvEnv.GetConfigOption("option_gaye") And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPostTaxPayrollGiving) Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.PostTaxPayrollGiving)
              vBankAccount = mvSOBankAccount
            End If
          Case 14 ' CI Batch Type
            If CCWithInvoice Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditCardWithInvoice)
              vBankAccount = mvCCBankAccount
            End If
          Case 15 ' AI Batch Type
            If ChequeWithInvoice Then
              vBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CashWithInvoice)
              vBankAccount = mvCABankAccount
            End If
        End Select
        If Len(vBatchType) > 0 Then
          vValid = mvEnv.Connection.GetCount("bank_suspense_accounts", Nothing, Replace(Replace(vSQL, "%1", vBatchType), "%2", vBankAccount)) > 0
          If Not vValid Then Exit For
        End If
      Next
      If vValid = False Then
        vErrorMsg = (ProjectText.String28999) 'At least one of the default bank accounts for this application does not have a corresponding suspense account.  This will prevent batches created with that bank account being processed by Cash Book Posting
        RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vErrorMsg)
      End If

      '-----------------------------------------------------
      '2. Check Source valid
      '-----------------------------------------------------
      If SourceValid = False Then
        vErrorMsg = (ProjectText.String15564) 'The source code defined for this Application is invalid. It may be historic or have an invalid thank you letter defined
        RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vErrorMsg)
      End If

      '-----------------------------------------------------
      '3. Check Credit Sales
      '-----------------------------------------------------
      If CreditSales Then
        If Len(CSTermsFrom) = 0 Then
          vErrorMsg = String.Format(ProjectText.String29379, CSCompany) 'Credit Controls have not been defined for Company '%s'
          RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vErrorMsg)
        End If
      End If

      '-----------------------------------------------------
      '4.  Check Payroll Giving
      '-----------------------------------------------------
      If (AppType = ApplicationType.atGiveAsYouEarnPayments Or AppType = ApplicationType.atPostTaxPGPayments) Then
        If Len(mvEnv.GetConfig("pm_gaye")) = 0 Then
          vErrorMsg = (ProjectText.String29380) 'The configuration option 'pm_gaye' has not been set
          RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vErrorMsg)
        End If
        If AppType = ApplicationType.atPostTaxPGPayments Then
          If Len(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGEmployerProduct)) = 0 Or Len(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGEmployerRate)) = 0 Then
            vErrorMsg = (ProjectText.String16949) 'The Post Tax Payroll Giving controls have not been set up
            RaiseError(DataAccessErrors.daeTraderApplicationInvalid, vErrorMsg)
          End If
        End If
      End If

    End Sub

    Private Sub ValidatePaymentMethod(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("batch_number", pBatchNumber)
      vWhereFields.Add("transaction_number", pTransactionNumber)
      Dim vPaymentMethod As String = New SQLStatement(mvEnv.Connection, "payment_method", "financial_history", vWhereFields).GetValue

      Dim vValid As Boolean = True
      If vPaymentMethod.Length > 0 Then
        Select Case vPaymentMethod
          Case mvEnv.GetConfig("pm_caf")
            'vValid - Trader App does not hold this value
          Case mvEnv.GetConfig("pm_caf_card")
            vValid = CAFCard
          Case mvEnv.GetConfig("pm_cash")
            vValid = Cash
          Case mvEnv.GetConfig("pm_cc")
            vValid = CreditCard
          Case mvEnv.GetConfig("pm_ccca")
            vValid = CreditCardAuthorities
          Case mvEnv.GetConfig("pm_cheque")
            vValid = Cheque
          Case mvEnv.GetConfig("pm_dc")
            vValid = DebitCard
          Case mvEnv.GetConfig("pm_dd")
            vValid = DirectDebits
          Case mvEnv.GetConfig("pm_dr")
            'vValid - 1. Trader App does not hold this value 2. We do not use this payment method anymore
          Case mvEnv.GetConfig("pm_gaye")
            vValid = PayrollGiving
          Case mvEnv.GetConfig("pm_gift_in_kind")
            vValid = GiftInKind
          Case mvEnv.GetConfig("pm_po")
            vValid = PostalOrder
          Case mvEnv.GetConfig("pm_so")
            vValid = StandingOrders
          Case mvEnv.GetConfig("pm_sp")
            vValid = Cash
          Case mvEnv.GetConfig("pm_sr")
            vValid = SaleOrReturn
          Case mvEnv.GetConfig("pm_voucher")
            vValid = True 'as this is a FA reanalysis, allow voucher transactions
          Case mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlInAdvancePaymentMethod)
            'Valid - Trader App does not hold this value
          Case mvCSPayMethod
            vValid = CreditSales
          Case Else
            RaiseError(DataAccessErrors.daeUnknownPaymentMethod, vPaymentMethod)
        End Select
        If Not vValid Then RaiseError(DataAccessErrors.daeTransactionPaymentMethodInvalid)
      Else
        RaiseError(DataAccessErrors.daeUnknownPaymentMethod, vPaymentMethod)
      End If
    End Sub

    Private Sub GetDefaults()
      Dim vSQL As String
      Dim vCount As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vCompanyControl As CompanyControl
      Dim vWhereFields As CDBFields

      'Get multi currency flag
      mvMultiCurrency = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode)
      'Get the default currency code
      mvDefaultCurrencyCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode)
      'Get the details of the default product, if specified
      If mvProduct.Length > 0 Then
        vSQL = "SELECT product_desc, despatch_method, stock_item,uses_product_numbers,max_numbers_allowed,sales_quantity"
        vSQL = vSQL & " FROM products WHERE product = '" & mvProduct & "'"
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        With vRecordSet
          If .Fetch() = True Then
            With .Fields
              If .Item(3).Bool Then mvConfirmDefaultProduct = True
              mvUsesProductNumbers = .Item("uses_product_numbers").Bool
              mvSalesQuantity = .Item("sales_quantity").Value
            End With
          End If
        End With
        vRecordSet.CloseRecordSet()
        If mvRate.Length > 0 Then
          vSQL = "SELECT rate_desc, current_price,vat_exclusive"
          If mvMultiCurrency Then vSQL = vSQL & ",currency_code"
          vSQL = vSQL & " FROM rates WHERE product = '" & mvProduct & "' AND rate = '" & mvRate & "'"
          vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
          With vRecordSet
            If .Fetch() = True Then
              With .Fields
                mvCurrentPrice = Val(.Item(2).Value)
              End With
            End If
          End With
          vRecordSet.CloseRecordSet()
        End If
      End If
      'Get the details of the default source, if specified
      mvSourceDesc = ""
      mvMailing = ""
      mvMailingDesc = ""
      mvDistributionCode = ""
      If mvSource.Length > 0 Then
        vSQL = "SELECT source_desc, thank_you_letter, mailing_desc, incentive_scheme, distribution_code"
        vSQL = vSQL & " FROM sources s, mailings m WHERE source = '" & mvSource & "' AND s.history_only = 'N' AND s.thank_you_letter = m.mailing"
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        With vRecordSet
          If .Fetch() = True Then
            With .Fields
              mvSourceDesc = .Item(1).Value
              mvMailing = .Item(2).Value
              mvMailingDesc = .Item(3).Value
              mvDistributionCode = .Item(5).Value

              vWhereFields = New CDBFields
              vWhereFields.Add("source", CDBField.FieldTypes.cftCharacter, mvSource)
              vWhereFields.Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
              vWhereFields.Add("incentive_scheme", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
              If mvEnv.Connection.GetCount("sources", vWhereFields) > 0 Then mvCheckIncentives = True
            End With
          Else
            mvSourceInvalid = True
            mvSource = ""
          End If
        End With
        vRecordSet.CloseRecordSet()
        If mvDistributionCode.Length > 0 Then
          Dim vRestrictionFields As New CDBFields
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHistoryOnlyDistribCodes) Then vRestrictionFields.Add("history_only", "N")
          mvDistributionCodeDesc = mvEnv.GetDescription("distribution_codes", "distribution_code", mvDistributionCode, vRestrictionFields)
        End If
      End If
      'Get the Gift Aid Minimum from the Covenant Controls Table
      mvGiftAidMinimum = Val(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVGiftAidMinimum))
      mvCovReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVReasonForDespatch)
      mvMinCovPeriod = IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVMinimumCovenantPeriod))
      mvCovActivity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivity)
      mvCovActivityValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivityValue)
      'Get the Reason for Despatch, Sponsor Activity & Sponsor Activity Value from the Membership Controls Table
      mvMemReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReasonForDespatch)
      mvSponsorAct = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivity)
      mvSponsorActVal = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivityValue)
      mvCMTCancelReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlTypeChangeCancelReason)
      'Get the control values from the Financial Controls Table
      mvCCReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCCReason)
      mvDDReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDDReason)
      mvSOReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSOReason)
      mvOReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOReason)
      mvDistActivity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCCReason)
      mvStockInterface = mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlStockInterface)
      'Get the config option that specifies whether a client is using our sales ledger
      mvUseSalesLedger = mvEnv.GetConfigOption("fp_use_sales_ledger", True)
      'If the app supports CS get some control data
      If mvCreditSales Then
        mvCSPayMethod = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod)
        mvCSTransType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSTransType)
        mvSundryCreditTransType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSCreditTransType)
        'Get the Company from the bank account
        vCompanyControl = New CompanyControl
        vCompanyControl.InitFromBankAccount(mvEnv, mvCSBankAccount, True)
        With vCompanyControl
          If .Existing Then
            mvCSCompany = .Company
            mvCSTermsNumber = .CSTermsNumber
            mvCSTermsPeriod = .CSTermsPeriod
            mvCSTermsFrom = .CSTermsFrom
            mvCSSundryCreditProduct = .SundryCreditProductCode
            mvCSSundryCreditRate = .SundryCreditRate
            mvCSDepositPercentage = .CSDepositPercentage
          End If
        End With
      End If
      'Get the Default Mailing config option
      Select Case LCase(mvEnv.GetConfig("default_mailing"))
        Case "letterbreaks"
          mvDefaultMailingType = DefaultMailingTypes.dmtLetterBreaks
        Case "source"
          mvDefaultMailingType = DefaultMailingTypes.dmtSource
        Case "breaksorsource"
          mvDefaultMailingType = DefaultMailingTypes.dmtLetterBreaksOrSource
        Case Else
          mvDefaultMailingType = DefaultMailingTypes.dmtNone
      End Select
      'Get the config option that specifies whether a warning message is to be displayed when no P&P is specified in a stock sale
      mvPostageWarnStock = mvEnv.GetConfigOption("opt_fp_postage_warn_stock", True)
      If mvMultiCurrency And mvBatchNumber > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT currency_exchange_rate, currency_code FROM batches WHERE batch_number = " & mvBatchNumber)
        With vRecordSet
          If .Fetch() = True Then
            mvBatchExchangeRate = Val(.Fields(1).Value)
            BatchCurrencyCode = .Fields(2).Value
          End If
        End With
        vRecordSet.CloseRecordSet()
      End If
      mvBackOrderPrompt = mvEnv.GetConfigOption("opt_fp_back_order_prompt")
      'Get the config option that specifies that contact discounting is done via an activity
      mvDiscountActivity = mvEnv.GetConfig("fp_discount_activity")
      mvAmendDiscount = mvEnv.GetConfigOption("fp_discount_amend")
      'Determine if client may be using CORE-based stock control
      vCount = mvEnv.Connection.GetCount("stock_movement_controls", Nothing, "")
      If vCount > 0 Then mvCoreStockControl = True
      'Get the set of bank accounts to support multi-currency
      If mvMultiCurrency Then
        mvCurrencyBankAccounts = New Collection
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT currency_code,batch_type,bank_account FROM fp_application_bank_accounts WHERE fp_application = '" & mvApplication & "'")
        With vRecordSet
          While .Fetch() = True
            mvCurrencyBankAccounts.Add(.Fields(3).Value, .Fields(1).Value & "-" & .Fields(2).Value)
          End While
          .CloseRecordSet()
        End With
      End If
      'Get the config option that controls whether a non-batch led Trader application creates a batch per user, per department or system wide
      Select Case mvEnv.GetConfig("opt_batch_per_user")
        Case "DEPARTMENT"
          mvBatchPerUser = BatchCreationTypes.bctPerDepartment
        Case "USER"
          mvBatchPerUser = BatchCreationTypes.bctPerUser
        Case Else
          mvBatchPerUser = BatchCreationTypes.bctPerSystem
      End Select
      'If the app supports the setting of a status then get the description of the default status
      If mvSetStatus And mvDefaultStatus.Length > 0 Then
        mvDefaultStatusDesc = mvEnv.GetDescription("statuses", "status", mvDefaultStatus)
      End If
      'If the app supports the adding of a mailing suppression then get the description of the default suppression
      If mvAddSuppression And mvDefaultSuppression.Length > 0 Then
        mvDefaultSuppresionDesc = mvEnv.GetDescription("mailing_suppressions", "mailing_suppression", mvDefaultSuppression)
      End If
      'If the app supports the cancelling of a payment plan then get the description of the default cancellation reason
      If mvCancelPayPlan And mvDefaultCancellationReason.Length > 0 Then
        mvDefaultCancellationReasonDesc = mvEnv.GetDescription("cancellation_reasons", "cancellation_reason", mvDefaultCancellationReason)
      End If
      'Get the config option that controls whether the default membership joined date should be blank
      mvBlankMembershipJoinedDate = mvEnv.GetConfigOption("me_blank_joined_date")
      'Get the config option that controls the maximum number of transactions allowed in a voucher batch
      mvMaxVoucherTransactions = IntegerValue(mvEnv.GetConfig("cv_max_number_of_vouchers"))
      If mvMaxVoucherTransactions = 0 Then mvMaxVoucherTransactions = 9999 'if no config defined then unlimited number of transactions allowed
      'Get the config option that controls whether the Payment Plan Details page should be shown for a Conversion-type application
      If mvAppType = ApplicationType.atConversion Then mvShowPPDInConversionApp = mvEnv.GetConfigOption("trader_conversion_app_show_ppd")
      'Get the config that indicates whether a reference is required for a CS transaction
      mvOptionalCreditSalesReference = mvEnv.GetConfigOption("fp_cs_reference_optional")
      'Get the config that defines the default setting of the Pack To Donor checkbox
      mvPackToDonorDefault = mvEnv.GetConfig("me_gift_pack_to_donor_default")
      If mvInvoicePayments Then mvCACompany = mvEnv.Connection.GetValue("SELECT company FROM bank_accounts WHERE bank_account = '" & mvCABankAccount & "'")
    End Sub

    Public Function CurrencyBankAccountExists(ByVal pKey As String) As Boolean
      Try
        Return mvCurrencyBankAccounts.Contains(pKey)
      Catch ex As Exception
        Return False
      End Try
    End Function

    Public Function CanDesign(ByRef pAppType As ApplicationType) As Boolean
      Select Case pAppType
        Case ApplicationType.atPurchaseOrderPrint, ApplicationType.atPurchaseOrderGeneration, ApplicationType.atChequeProcessing
          'These application types do not use Trader forms
          CanDesign = False
        Case Else
          CanDesign = True
      End Select
    End Function

    Public Function GetApplicationType(ByRef pCode As String) As ApplicationType
      Select Case pCode
        Case "TRANS"
          GetApplicationType = ApplicationType.atTransaction
        Case "PINVE"
          GetApplicationType = ApplicationType.atPurchaseInvoice
        Case "PORDE"
          GetApplicationType = ApplicationType.atPurchaseOrder
        Case "PORDC"
          GetApplicationType = ApplicationType.atPurchaseOrderCancellation
        Case "CHQNA"
          GetApplicationType = ApplicationType.atChequeNumberAllocation
        Case "CHQRE"
          GetApplicationType = ApplicationType.atChequeReconciliation
        Case "CSTAT"
          GetApplicationType = ApplicationType.atCreditStatementGeneration
        Case "BINVG"
          GetApplicationType = ApplicationType.atBatchInvoiceGeneration
        Case "MAINT"
          GetApplicationType = ApplicationType.atMaintenance
        Case "CNVRT"
          GetApplicationType = ApplicationType.atConversion
        Case "CLREC"
          GetApplicationType = ApplicationType.atCreditListReconciliation
        Case "BSPOS"
          GetApplicationType = ApplicationType.atBankStatementPosting
        Case "POGEN"
          GetApplicationType = ApplicationType.atPurchaseOrderGeneration
        Case "POPRT"
          GetApplicationType = ApplicationType.atPurchaseOrderPrint
        Case "POCHQ"
          GetApplicationType = ApplicationType.atChequeProcessing
        Case "GAYEP"
          GetApplicationType = ApplicationType.atGiveAsYouEarnPayments
        Case "POTPG"
          GetApplicationType = ApplicationType.atPostTaxPGPayments
      End Select
    End Function

    Public Sub SaveApplicationSize(ByRef pAppHeight As Integer, ByRef pAppWidth As Integer)
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields

      vFields.Add("application_height", CDBField.FieldTypes.cftLong, pAppHeight)
      vFields.Add("application_width", CDBField.FieldTypes.cftLong, pAppWidth)

      vWhereFields.Add("fp_application", CDBField.FieldTypes.cftCharacter, mvApplication)
      mvEnv.Connection.UpdateRecords("fp_applications", vFields, vWhereFields, False)
    End Sub

    Public Sub InitNonFinancialTransaction(Optional ByRef pAllocateTransactionNo As Boolean = True)
      If SupportsNonFinancialBatch Then
        If mvNFBatchNumber = 0 Then
          If BatchLedApp Then
            If mvBatch.BatchType = Batch.BatchTypes.NonFinancial Then mvNonFinancialBatch = mvBatch
          End If
          If mvNonFinancialBatch Is Nothing Then
            mvNonFinancialBatch = New Batch(mvEnv)
            mvNonFinancialBatch.InitOpenBatch(Nothing, Batch.ProvisionalOrConfirmed.Confirmed, Batch.BatchTypes.NonFinancial, CABankAccount, "CASH", True, Batch.BatchTypes.None, "", 0, BatchCategory, "", True, "", BatchAnalysisCode)
          End If
          mvNFBatchNumber = mvNonFinancialBatch.BatchNumber
          If pAllocateTransactionNo Then
            mvNFTransNumber = mvNonFinancialBatch.AllocateTransactionNumber
          Else
            mvNFTransNumber = mvNonFinancialBatch.NextTransactionNumber
          End If
        Else
          mvNFTransNumber = mvNonFinancialBatch.AllocateTransactionNumber
        End If
      End If
    End Sub

    Public Sub InitControls()
      Dim vIndex As TraderPage.TraderPageType
      Dim vRecordSet As CDBRecordSet
      Dim vPage As String
      Dim vLastPage As String = ""
      Dim vCurrentPage As Integer
      Dim vType As String
      Dim vIgnorePage As Boolean
      Dim vIgnoreControl As Boolean
      Dim vPayPlans As Boolean
      Dim vPageType As TraderPage.TraderPageType
      Dim vNextTop As Integer
      Dim vNextOffset As Integer
      Dim vAttrs As String
      Dim vMenu As Boolean
      Dim vRestriction As String
      Dim vCursorCount As Integer
      Dim vPages As String
      Dim vStartPage As TraderPage.TraderPageType
      Dim vEndPage As TraderPage.TraderPageType
      Dim vControlIndex As Integer
      Dim vGetSource As TraderPage.GetSourceFromMailingType
      Dim vTabIndex As Integer
      Dim vUpdateAfMember As Boolean
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vUpdateBookingParam As Boolean

      mvTraderControls = New TraderControls
      mvTraderPages = New TraderPages

      vPayPlans = PaymentPlans

      vAttrs = "fp_page_type,control_type,control_top,control_left,control_width,control_height,control_caption,caption_width,help_text,visible,fc.table_name,fc.attribute_name"
      vAttrs = vAttrs & ",type,entry_length,nulls_invalid,pattern,validation_table,validation_attribute,"
      vAttrs = vAttrs & mvEnv.Connection.DBSpecialCol("fa", "case")
      vAttrs = vAttrs & ",minimum_value,maximum_value"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataControlsContactGroup) Then vAttrs = vAttrs & ",contact_group"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataControlParameterName) Then vAttrs = vAttrs & ",parameter_name"
      vAttrs &= ",mandatory_item,default_value"

      For vCursorCount = 1 To 2
        Select Case AppType
          Case ApplicationType.atTransaction, ApplicationType.atConversion, ApplicationType.atCreditListReconciliation, ApplicationType.atBankStatementPosting 'TRANS or CNVRT
            vStartPage = TraderPage.TraderPageType.tpPaymentMethod1
            If mvPayPlanConvMaintenance Then
              vEndPage = TraderPage.TraderPageType.tpLoans
            ElseIf ConversionShowPPD Then
              vEndPage = TraderPage.TraderPageType.tpLoans
            Else
              vEndPage = TraderPage.TraderPageType.tpAmendEventBooking
            End If
          Case ApplicationType.atPurchaseInvoice 'PINVE
            vStartPage = TraderPage.TraderPageType.tpPurchaseInvoiceDetails
            vEndPage = TraderPage.TraderPageType.tpPurchaseInvoiceProducts
          Case ApplicationType.atPurchaseOrder 'PORDE
            vStartPage = TraderPage.TraderPageType.tpPurchaseOrderDetails
            vEndPage = TraderPage.TraderPageType.tpPurchaseOrderPayments
          Case ApplicationType.atPurchaseOrderCancellation 'PORDC
            vStartPage = TraderPage.TraderPageType.tpPurchaseOrderCancellation
            vEndPage = TraderPage.TraderPageType.tpPurchaseOrderCancellation
          Case ApplicationType.atChequeNumberAllocation 'CHQNA
            vStartPage = TraderPage.TraderPageType.tpChequeNumberAllocation
            vEndPage = TraderPage.TraderPageType.tpChequeNumberAllocation
          Case ApplicationType.atChequeReconciliation 'CHQRE
            vStartPage = TraderPage.TraderPageType.tpChequeReconciliation
            vEndPage = TraderPage.TraderPageType.tpChequeReconciliation
          Case ApplicationType.atCreditStatementGeneration 'CSTAT
            vStartPage = TraderPage.TraderPageType.tpCreditStatementGeneration
            vEndPage = TraderPage.TraderPageType.tpCreditStatementGeneration
          Case ApplicationType.atBatchInvoiceGeneration 'BINVG
            vStartPage = TraderPage.TraderPageType.tpBatchInvoiceProduction
            vEndPage = TraderPage.TraderPageType.tpBatchInvoiceSummary
          Case ApplicationType.atMaintenance 'MAINT
            vStartPage = TraderPage.TraderPageType.tpContactSelection
            vEndPage = TraderPage.TraderPageType.tpLoans
          Case ApplicationType.atGiveAsYouEarnPayments 'GAYEP
            vStartPage = TraderPage.TraderPageType.tpGiveAsYouEarn
            vEndPage = TraderPage.TraderPageType.tpGiveAsYouEarn
          Case ApplicationType.atPostTaxPGPayments 'POTPG
            vStartPage = TraderPage.TraderPageType.tpPostTaxPGPayment
            vEndPage = TraderPage.TraderPageType.tpPostTaxPGPayment
          Case ApplicationType.atChequeProcessing 'POCHQ
            Exit Sub
        End Select

        vPages = ""
        For vIndex = vStartPage To vEndPage
          If AppType = ApplicationType.atMaintenance Then '"MAINT"
            Select Case vIndex
              Case TraderPage.TraderPageType.tpContactSelection, TraderPage.TraderPageType.tpPaymentPlanMaintenance, TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance, TraderPage.TraderPageType.tpScheduledPayments, TraderPage.TraderPageType.tpLoans
                'Contact Selection and Maint PP and PP Details, OPS & Loans
                vIgnorePage = False
              Case Else
                vIgnorePage = True
            End Select
          Else
            Select Case vIndex
              Case TraderPage.TraderPageType.tpCreditCustomer 'Credit customer"
                vIgnorePage = Not CreditSales
              Case TraderPage.TraderPageType.tpComments 'Comments
                vIgnorePage = Not TransactionComments
              Case TraderPage.TraderPageType.tpBankDetails 'Bank Details
                vIgnorePage = Not BankDetails
              Case TraderPage.TraderPageType.tpCardDetails 'Credit or Debit Card details
                vIgnorePage = Not (CreditCard Or DebitCard Or CAFCard Or CCWithInvoice)
              Case TraderPage.TraderPageType.tpPaymentMethod2 'Payment method choice 2
                vIgnorePage = Not vPayPlans
              Case TraderPage.TraderPageType.tpProductDetails 'Product details
                vIgnorePage = Not (ProductSales Or OneOffDonations Or CreditNotes)
              Case TraderPage.TraderPageType.tpPayments 'Payments
                vIgnorePage = Not (Payments Or vPayPlans)
              Case TraderPage.TraderPageType.tpPaymentPlanDetails 'Payment Plan details
                vIgnorePage = Not vPayPlans
              Case TraderPage.TraderPageType.tpPaymentPlanProducts 'Payment Plan products
                vIgnorePage = Not vPayPlans
              Case TraderPage.TraderPageType.tpStandingOrder 'Standing Order
                vIgnorePage = Not ((vPayPlans And StandingOrders) Or AutoPaymentMaintenance)
              Case TraderPage.TraderPageType.tpDirectDebit 'Direct Debit
                vIgnorePage = Not ((vPayPlans And DirectDebits) Or AutoPaymentMaintenance)
              Case TraderPage.TraderPageType.tpCreditCardAuthority 'Continuous Credit Card Authority
                vIgnorePage = Not ((vPayPlans And CreditCardAuthorities) Or AutoPaymentMaintenance)
              Case TraderPage.TraderPageType.tpMembership, TraderPage.TraderPageType.tpChangeMembershipType, TraderPage.TraderPageType.tpAmendMembership, TraderPage.TraderPageType.tpMembershipPayer 'Membership, Amend Membership Details & Change Membership Type & Membership Payer (CMT)
                vIgnorePage = Not (vPayPlans And (Memberships Or ChangeMembership Or CovMemberships))
              Case TraderPage.TraderPageType.tpCovenant 'Covenant
                vIgnorePage = Not (vPayPlans And (CovMemberships Or CovSubscriptions Or CovDonationsRegular))
              Case TraderPage.TraderPageType.tpContactSelection
                vIgnorePage = Not AutoPaymentMaintenance And (Cash Or Cheque Or PostalOrder Or CreditCard Or DebitCard Or CreditSales Or Voucher Or GiftInKind Or CAFCard Or SaleOrReturn)
              Case TraderPage.TraderPageType.tpEventBooking, TraderPage.TraderPageType.tpAmendEventBooking 'Event Booking, Amend Event Booking
                vIgnorePage = Not SupportsEventBooking
              Case TraderPage.TraderPageType.tpExamBooking
                vIgnorePage = Not SupportsExamBooking
              Case TraderPage.TraderPageType.tpAccommodationBooking 'Accommodation Booking
                vIgnorePage = Not AccomodationBooking
              Case TraderPage.TraderPageType.tpServiceBooking
                vIgnorePage = Not (ServiceBookings Or ServiceBookingCredits)
              Case TraderPage.TraderPageType.tpPostageAndPacking 'Postage and Packing - Carriage
                vIgnorePage = Not Carriage
              Case TraderPage.TraderPageType.tpInvoicePayments 'Invoice Payments
                vIgnorePage = Not InvoicePayments
                If vIgnorePage Then vIgnorePage = Not CreditNotes
              Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
                vIgnorePage = Not PayPlanPayMethod
              Case TraderPage.TraderPageType.tpSetStatus 'Set Status
                vIgnorePage = Not SetStatus
              Case TraderPage.TraderPageType.tpCancelPaymentPlan 'Cancel Payment Plan
                vIgnorePage = Not CancelPaymentPlan
              Case TraderPage.TraderPageType.tpLegacyBequestReceipt 'Legacy Bequest Receipt
                vIgnorePage = Not LegaceyReceipt
              Case TraderPage.TraderPageType.tpActivityEntry 'Activity Entry
                vIgnorePage = Not AddActivity
              Case TraderPage.TraderPageType.tpGiftAidDeclaration 'Gift Aid Declaration
                vIgnorePage = Not GiftAidDeclaration
              Case TraderPage.TraderPageType.tpGiveAsYouEarnEntry 'Give As You Earn or Payroll Giving Entry
                vIgnorePage = Not PayrollGiving
              Case TraderPage.TraderPageType.tpGoneAway 'Gone Away
                vIgnorePage = Not GoneAway
              Case TraderPage.TraderPageType.tpAddressMaintenance 'Address Maintenance
                vIgnorePage = Not AddressMaintenance
              Case TraderPage.TraderPageType.tpSuppressionEntry 'Suppression Entry
                vIgnorePage = Not AddSuppression
              Case TraderPage.TraderPageType.tpCancelGiftAidDeclaration 'Cancel Gift Aid Declaration
                vIgnorePage = Not CancelGiftAidDeclaration
              Case TraderPage.TraderPageType.tpScheduledPayments 'Scheduled Payments
                vIgnorePage = Not DisplayScheduledPayments
              Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
                vIgnorePage = Not (Payments Or vPayPlans)
              Case TraderPage.TraderPageType.tpConfirmProvisionalTransactions 'Sale or Return Transactions
                vIgnorePage = Not ConfirmSaleOrReturnTransactions
              Case TraderPage.TraderPageType.tpCollectionPayments
                vIgnorePage = Not CollectionPayments 'Public Collection Payments
              Case TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
                vIgnorePage = False
                Select Case AppType
                  Case ApplicationType.atTransaction, ApplicationType.atConversion, ApplicationType.atCreditListReconciliation, ApplicationType.atBankStatementPosting
                    If Not ConversionShowPPD AndAlso Not mvPayPlanConvMaintenance Then vIgnorePage = True
                End Select
              Case TraderPage.TraderPageType.tpLoans
                vIgnorePage = Not (Loans)
              Case TraderPage.TraderPageType.tpAdvancedCMT
                vIgnorePage = Not (ChangeMembership)
              Case Else
                vIgnorePage = False
            End Select
          End If
          If Not vIgnorePage Then
            vPage = GetPageCode(vIndex) 'Get the three character code for the page
            If (vIndex = TraderPage.TraderPageType.tpMembership Or vIndex = TraderPage.TraderPageType.tpChangeMembershipType) Then vUpdateAfMember = True
            If vIndex = TraderPage.TraderPageType.tpEventBooking Then vUpdateBookingParam = True

            Debug.Print("Loading Page " & vPage)

            If vPage.Length > 0 Then
              If GetPageIndex(vIndex) = 0 Then
                If vPages.Length > 0 Then vPages = vPages & ","
                vPages = vPages & "'" & vPage & "'"
              End If
            End If
          End If
        Next
        If vPages.Length > 0 Then
          vRestriction = " AND fp_page_type IN (" & vPages & ")"
        Else
          Exit For
        End If

        If vCursorCount = 1 And vUpdateAfMember = True Then
          'Update "affiliated_member_number" to be "member_number"
          'Problem with Trader Designer incorrectly re-named the attribute name
          vUpdateFields = New CDBFields
          vWhereFields = New CDBFields
          vUpdateFields.Add("attribute_name", CDBField.FieldTypes.cftCharacter, "member_number")
          With vWhereFields
            .Add("fp_application", CDBField.FieldTypes.cftCharacter, Application)
            .Add("fp_page_type", CDBField.FieldTypes.cftCharacter, "'MEM','CMT'", CDBField.FieldWhereOperators.fwoIn)
            .Add("table_name", CDBField.FieldTypes.cftCharacter, "members")
            .Add("attribute_name", CDBField.FieldTypes.cftCharacter, "affiliated_member_number")
          End With
          mvEnv.Connection.UpdateRecords("fp_controls", vUpdateFields, vWhereFields, False)
        End If

        'Change to set 2nd Event Booking page booking_number parameter name
        If vCursorCount = 1 And vUpdateBookingParam Then
          vUpdateFields = New CDBFields
          vWhereFields = New CDBFields
          vUpdateFields.Add("parameter_name", CDBField.FieldTypes.cftCharacter, "InterestBookingNumber")
          With vWhereFields
            .Add("fp_page_type", CDBField.FieldTypes.cftCharacter, "EVE")
            .Add("table_name", CDBField.FieldTypes.cftCharacter, "delegates")
            .Add("attribute_name", CDBField.FieldTypes.cftCharacter, "booking_number")
            .Add("parameter_name", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual)
          End With
          mvEnv.Connection.UpdateRecords("fp_controls", vUpdateFields, vWhereFields, False)
        End If

        If vCursorCount = 1 Then 'Get application specific pages
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAttrs & " FROM fp_controls fc, maintenance_attributes fa WHERE fp_application = '" & Application & "' " & vRestriction & " AND fc.table_name = fa.table_name AND fc.attribute_name = fa.attribute_name ORDER BY fp_page_type, fc.sequence_number")
        Else 'Get default pages
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAttrs & " FROM fp_controls fc, maintenance_attributes fa WHERE fp_application is null " & vRestriction & " AND fc.table_name = fa.table_name AND fc.attribute_name = fa.attribute_name ORDER BY fp_page_type, fc.sequence_number")
        End If
        While vRecordSet.Fetch() = True
          vPage = vRecordSet.Fields("fp_page_type").Value
          If vPage <> vLastPage Then
            vIgnorePage = False
            vMenu = False
            vLastPage = vPage
            vGetSource = TraderPage.GetSourceFromMailingType.gsfmtNever
            Select Case vPage
              Case "PM1" 'Payment method choice 1
                vPageType = TraderPage.TraderPageType.tpPaymentMethod1
                vMenu = True
              Case "CCU" 'Credit customer
                vPageType = TraderPage.TraderPageType.tpCreditCustomer
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "TRD" 'Transaction details
                vPageType = TraderPage.TraderPageType.tpTransactionDetails
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "COM" 'Comments
                vPageType = TraderPage.TraderPageType.tpComments
              Case "BKD" 'Bank Details
                vPageType = TraderPage.TraderPageType.tpBankDetails
              Case "CDC" 'Credit or Debit Card details
                vPageType = TraderPage.TraderPageType.tpCardDetails
              Case "TRA" 'Transaction Analysis choice
                vPageType = TraderPage.TraderPageType.tpTransactionAnalysis
                vMenu = True
              Case "PM2" 'Payment method choice 2
                vPageType = TraderPage.TraderPageType.tpPaymentMethod2
                vMenu = True
              Case "PM3" 'Payment method choice 3
                vPageType = TraderPage.TraderPageType.tpPaymentMethod3
                vMenu = True
              Case "PRD" 'Product details
                vPageType = TraderPage.TraderPageType.tpProductDetails
              Case "PAY" 'Payments
                vPageType = TraderPage.TraderPageType.tpPayments
              Case "PPD" 'Payment Plan details
                vPageType = TraderPage.TraderPageType.tpPaymentPlanDetails
              Case "PPP" 'Payment Plan products
                vPageType = TraderPage.TraderPageType.tpPaymentPlanProducts
              Case "STO" 'Standing Order
                vPageType = TraderPage.TraderPageType.tpStandingOrder
              Case "DDR" 'Direct Debit
                vPageType = TraderPage.TraderPageType.tpDirectDebit
              Case "CCA" 'Continuous Credit Card Authority
                vPageType = TraderPage.TraderPageType.tpCreditCardAuthority
              Case "MEM" 'Membership
                vPageType = TraderPage.TraderPageType.tpMembership
              Case "CMT" 'Change Membership Type (CMT)
                vPageType = TraderPage.TraderPageType.tpChangeMembershipType
              Case "AMD" 'Amend Membership Details
                vPageType = TraderPage.TraderPageType.tpAmendMembership
              Case "MSP" 'Covenant
                vPageType = TraderPage.TraderPageType.tpMembershipPayer
              Case "COV" 'Covenant
                vPageType = TraderPage.TraderPageType.tpCovenant
              Case "CSE" 'Contact selection
                vPageType = TraderPage.TraderPageType.tpContactSelection
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "EVE" 'Event Booking
                vPageType = TraderPage.TraderPageType.tpEventBooking
              Case "EXA" 'Exam Booking
                vPageType = TraderPage.TraderPageType.tpExamBooking
              Case "ACO" 'Accomodation Booking
                vPageType = TraderPage.TraderPageType.tpAccommodationBooking
              Case "PID" 'Purchase Invoice Details
                vPageType = TraderPage.TraderPageType.tpPurchaseInvoiceDetails
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "PIP" 'Purchase Invoice Products
                vPageType = TraderPage.TraderPageType.tpPurchaseInvoiceProducts
              Case "POD" 'Purchase Order Details
                vPageType = TraderPage.TraderPageType.tpPurchaseOrderDetails
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "POP" 'Purchase Order Products
                vPageType = TraderPage.TraderPageType.tpPurchaseOrderProducts
              Case "PPA" 'Purchase Order Payments
                vPageType = TraderPage.TraderPageType.tpPurchaseOrderPayments
              Case "POC" 'Purchase Order Cancellation
                vPageType = TraderPage.TraderPageType.tpPurchaseOrderCancellation
              Case "CNA" 'Cheque Number Allocation
                vPageType = TraderPage.TraderPageType.tpChequeNumberAllocation
              Case "CRE" 'Cheque Reconciliation
                vPageType = TraderPage.TraderPageType.tpChequeReconciliation
              Case "PAP"
                vPageType = TraderPage.TraderPageType.tpPostageAndPacking 'Postage & Packing - Carriage
              Case "SVC"
                vPageType = TraderPage.TraderPageType.tpServiceBooking
              Case "CSG"
                vPageType = TraderPage.TraderPageType.tpCreditStatementGeneration
              Case "ING"
                vPageType = TraderPage.TraderPageType.tpBatchInvoiceProduction
              Case "TPP"
                vPageType = TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
                vMenu = True
              Case "PPM"
                vPageType = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Case "PPN"
                vPageType = TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
              Case "STA"
                vPageType = TraderPage.TraderPageType.tpSetStatus
              Case "CPP"
                vPageType = TraderPage.TraderPageType.tpCancelPaymentPlan
              Case "LBR"
                vPageType = TraderPage.TraderPageType.tpLegacyBequestReceipt
              Case "ACT"
                vPageType = TraderPage.TraderPageType.tpActivityEntry
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtMaybe
              Case "GYP"
                vPageType = TraderPage.TraderPageType.tpGiveAsYouEarn
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtAlways
              Case "GAD"
                vPageType = TraderPage.TraderPageType.tpGiftAidDeclaration
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtMaybe
              Case "GYE"
                vPageType = TraderPage.TraderPageType.tpGiveAsYouEarnEntry
                vGetSource = TraderPage.GetSourceFromMailingType.gsfmtMaybe
              Case "GAW"
                vPageType = TraderPage.TraderPageType.tpGoneAway
              Case "ADM"
                vPageType = TraderPage.TraderPageType.tpAddressMaintenance
              Case "SUP"
                vPageType = TraderPage.TraderPageType.tpSuppressionEntry
              Case "CGA"
                vPageType = TraderPage.TraderPageType.tpCancelGiftAidDeclaration
              Case "SCP"
                vPageType = TraderPage.TraderPageType.tpScheduledPayments
              Case "OSP"
                vPageType = TraderPage.TraderPageType.tpOutstandingScheduledPayments
              Case "CPT"
                vPageType = TraderPage.TraderPageType.tpConfirmProvisionalTransactions
              Case "PGP"
                vPageType = TraderPage.TraderPageType.tpPostTaxPGPayment
              Case "PCP"
                vPageType = TraderPage.TraderPageType.tpCollectionPayments
              Case "INS"
                vPageType = TraderPage.TraderPageType.tpBatchInvoiceSummary
              Case "AEV"
                vPageType = TraderPage.TraderPageType.tpAmendEventBooking
              Case "LON"
                vPageType = TraderPage.TraderPageType.tpLoans
              Case "MTC"
                vPageType = TraderPage.TraderPageType.tpAdvancedCMT
              Case "TKN"
                vPageType = TraderPage.TraderPageType.tpTokenSelection
            End Select
            If GetPageIndex(vPageType) > 0 Then vIgnorePage = True 'Check if page already loaded
            If Not vIgnorePage Then
              vCurrentPage = vCurrentPage + 1 'New page
              vTabIndex = 0
              If vCurrentPage > mvTraderPages.Count Then
                With mvTraderPages.Add
                  .First = vControlIndex
                  .PageType = vPageType
                  .PageCode = vPage
                  .Menu = vMenu
                  .MenuCount = 0
                  .FirstMenuIndex = 0
                  .GetSourceFromMailing = vGetSource
                End With
                If vCurrentPage > 1 Then mvTraderPages(vCurrentPage - 1).Last = vControlIndex - 1
              End If
            End If
            vNextTop = 0 'Reset moving controls
          End If

          If Not vIgnorePage Then
            vType = vRecordSet.Fields("control_type").Value
            If Left(vType, 3) = "cmd" Then
              Select Case vType
                Case "cmd_CASH"
                  vIgnoreControl = Not Cash
                Case "cmd_CHEQ"
                  vIgnoreControl = Not Cheque
                Case "cmd_POST"
                  vIgnoreControl = Not PostalOrder
                Case "cmd_CARD"
                  vIgnoreControl = Not (CreditCard Or DebitCard)
                Case "cmd_CRED"
                  vIgnoreControl = Not CreditSales
                Case "cmd_CQIN"
                  vIgnoreControl = Not ChequeWithInvoice
                Case "cmd_CCIN"
                  vIgnoreControl = Not CCWithInvoice OrElse mvEnv.GetConfig("fp_card_sales_combined_claim", "N") = "N"
                Case "cmd_MEMB"
                  vIgnoreControl = Not Memberships
                Case "cmd_SUBS"
                  vIgnoreControl = Not Subscriptions
                Case "cmd_DONR"
                  vIgnoreControl = Not DonationsRegular
                Case "cmd_CMEM"
                  vIgnoreControl = Not CovMemberships
                Case "cmd_CSUB"
                  vIgnoreControl = Not CovSubscriptions
                Case "cmd_CDON"
                  vIgnoreControl = Not CovDonationsRegular
                Case "cmd_SALE"
                  vIgnoreControl = Not ProductSales
                Case "cmd_DONS"
                  vIgnoreControl = Not OneOffDonations
                Case "cmd_PAYM"
                  vIgnoreControl = Not Payments
                Case "cmd_CURR"
                  vIgnoreControl = Not (Cash Or Cheque Or PostalOrder Or CreditCard Or DebitCard Or CreditSales Or CAFCard Or Voucher)
                Case "cmd_STDO"
                  vIgnoreControl = Not StandingOrders
                Case "cmd_DIRD"
                  vIgnoreControl = Not DirectDebits
                Case "cmd_CCCA"
                  vIgnoreControl = Not CreditCardAuthorities
                Case "cmd_NPAY"
                  vIgnoreControl = Not NoPaymentRequired
                Case "cmd_EVNT"
                  vIgnoreControl = Not SupportsEventBooking
                Case "cmd_EXAM"
                  vIgnoreControl = Not SupportsExamBooking
                Case "cmd_ACOM"
                  vIgnoreControl = Not AccomodationBooking
                Case "cmd_SRVC"
                  vIgnoreControl = Not (ServiceBookings Or ServiceBookingCredits)
                Case "cmd_INVC"
                  vIgnoreControl = Not InvoicePayments
                Case "cmd_CRDN"
                  vIgnoreControl = Not CreditNotes
                Case "cmd_MEMC" 'Change Membership Type (CMT)
                  vIgnoreControl = Not ChangeMembership
                Case "cmd_LEGR"
                  vIgnoreControl = Not LegaceyReceipt
                Case "cmd_STAT"
                  vIgnoreControl = Not SetStatus
                Case "cmd_AWAY"
                  vIgnoreControl = Not GoneAway
                Case "cmd_GIFT"
                  vIgnoreControl = Not GiftAidDeclaration
                Case "cmd_ACTV"
                  vIgnoreControl = Not AddActivity
                Case "cmd_SUPP"
                  vIgnoreControl = Not AddSuppression
                Case "cmd_CANC"
                  vIgnoreControl = Not CancelPaymentPlan
                Case "cmd_APAY"
                  vIgnoreControl = Not AutoPaymentMaintenance
                Case "cmd_ADDR"
                  vIgnoreControl = Not AddressMaintenance
                Case "cmd_VOUC"
                  vIgnoreControl = Not Voucher
                Case "cmd_GFIK"
                  vIgnoreControl = Not GiftInKind
                Case "cmd_CAFC"
                  vIgnoreControl = Not CAFCard
                Case "cmd_COVT"
                  vIgnoreControl = Not (CovDonationsRegular = True Or CovMemberships = True Or CovSubscriptions = True)
                Case "cmd_DIRD"
                  vIgnoreControl = Not DirectDebits
                Case "cmd_STDO"
                  vIgnoreControl = Not StandingOrders
                Case "cmd_CCCA"
                  vIgnoreControl = Not CreditCardAuthorities
                Case "cmd_CVDD"
                  vIgnoreControl = Not (DirectDebits = True And (CovDonationsRegular = True Or CovMemberships = True Or CovSubscriptions = True))
                Case "cmd_CVSO"
                  vIgnoreControl = Not (StandingOrders = True And (CovDonationsRegular = True Or CovMemberships = True Or CovSubscriptions = True))
                Case "cmd_CVCC"
                  vIgnoreControl = Not (CreditCardAuthorities = True And (CovDonationsRegular = True Or CovMemberships = True Or CovSubscriptions = True))
                Case "cmd_GAYE"
                  vIgnoreControl = Not PayrollGiving
                Case "cmd_SAOR"
                  vIgnoreControl = Not SaleOrReturn
                Case "cmd_TRAN"
                  vIgnoreControl = Not (PayPlanPayMethod And (Cash Or Cheque Or PostalOrder Or CreditCard Or DebitCard Or CreditSales))
                Case "cmd_PLAN"
                  vIgnoreControl = Not PayPlanPayMethod
                Case "cmd_CGAD"
                  vIgnoreControl = Not CancelGiftAidDeclaration
                Case "cmd_CSRT"
                  vIgnoreControl = Not ConfirmSaleOrReturnTransactions
                Case "cmd_COLP"
                  vIgnoreControl = Not CollectionPayments
                Case "cmd_LOAN"
                  vIgnoreControl = Not Loans
                Case Else
                  '
              End Select
              If vIgnoreControl And vNextTop = 0 Then
                vNextTop = vRecordSet.Fields("control_top").IntegerValue
              Else
                If vNextTop > 0 And vNextOffset = 0 Then
                  vNextOffset = vRecordSet.Fields("control_top").IntegerValue - vNextTop
                End If
              End If
              If Not vIgnoreControl Then
                With mvTraderPages(vCurrentPage)
                  .MenuCount = .MenuCount + 1
                  If .MenuCount = 1 Then .FirstMenuIndex = vControlIndex
                End With
              End If
            Else
              vIgnoreControl = False
            End If

            If Not vIgnoreControl Then
              mvTraderControls.AddFromRecordSet(mvEnv, vRecordSet, vPageType, vTabIndex, vNextTop, mvTraderPages(vCurrentPage).First, Me)
              If vNextTop > 0 Then
                vNextTop = vNextTop + vNextOffset
              End If
              vControlIndex = vControlIndex + 1
            End If
          End If
        End While
        vRecordSet.CloseRecordSet()
      Next

      mvTraderPages(vCurrentPage).Last = vControlIndex - 1
      'Now add in the summary pages
      AddTraderPage(vControlIndex, vControlIndex, TraderPage.TraderPageType.tpTransactionAnalysisSummary, True)
      AddTraderPage(vControlIndex, vControlIndex, SummaryPage, True)
      AddTraderPage(vControlIndex, vControlIndex, TraderPage.TraderPageType.tpInvoicePayments, True)
      AddTraderPage(vControlIndex, vControlIndex, TraderPage.TraderPageType.tpMembershipMembersSummary, True)
      AddTraderPage(vControlIndex, vControlIndex, TraderPage.TraderPageType.tpDummyPage, True) 'BR9067
      If AppType = ApplicationType.atCreditListReconciliation Then AddTraderPage(vControlIndex, vControlIndex, TraderPage.TraderPageType.tpStatementList, True)
    End Sub

    Private Sub AddTraderPage(ByRef pFirst As Integer, ByRef pLast As Integer, ByRef pType As TraderPage.TraderPageType, ByRef pSummary As Boolean)
      With mvTraderPages.Add
        .First = pFirst
        .Last = pLast
        .PageType = pType
        .PageCode = GetPageCode(pType)
        .Summary = pSummary
      End With
    End Sub

    Public Function GetPageIndex(ByVal pType As TraderPage.TraderPageType) As Integer
      Dim vIndex As Integer

      For vIndex = 1 To mvTraderPages.Count
        If mvTraderPages(vIndex).PageType = pType Then 'Found the required page
          GetPageIndex = vIndex
          Exit Function
        End If
      Next
    End Function

    Public Function GetPaymentPlanDetail(ByVal pKey As Integer) As PaymentPlanDetail
      Dim vPPD As PaymentPlanDetail

      If mvPaymentPlanDetails.Exists(CStr(pKey)) Then
        vPPD = mvPaymentPlanDetails(CStr(pKey))
      Else
        vPPD = mvPaymentPlanDetails.Add(CStr(pKey))
        vPPD.Init(mvEnv)
      End If
      GetPaymentPlanDetail = vPPD

    End Function

    Public Sub DeletePaymentPlanDetail(ByVal pKey As Integer)
      Dim vCol As TraderPaymentPlanDetails
      Dim vPPD As PaymentPlanDetail
      Dim vIndex As Integer

      'Delete this Key
      If mvPaymentPlanDetails.Exists(CStr(pKey)) Then
        mvPaymentPlanDetails.Remove(CStr(pKey))

        vCol = mvPaymentPlanDetails
        mvPaymentPlanDetails = New TraderPaymentPlanDetails

        'Now, re-number the collection
        vIndex = 1
        For Each vPPD In vCol
          mvPaymentPlanDetails.AddItem(vPPD, CStr(vIndex))
          vIndex = vIndex + 1
        Next vPPD
      End If
    End Sub

    Public Function SCProcessTraderData(ByRef pType As TraderProcessDataTypes, ByRef pParams As CDBParameters, ByRef pTransaction As TraderTransaction) As CDBParameters
      'Called from the smart client at each page change (Next, Previous, Finished)
      'The given parameters should include the current page 'CurrentPageType' and any round trip information
      'The general sequence should be
      ' Validate the given details for the current page (if required)
      ' Save any data required for subsequent processing - This data may need to round trip to the client on each call
      ' Generate any analysis lines which need to be returned
      ' Identify the next required page 'NextPageType'
      ' Return any defaults for the page (if required)
      '
      Dim vResults As New CDBParameters
      Dim vCurrentPage As TraderPage.TraderPageType
      Dim vNextPage As TraderPage.TraderPageType
      Dim vFinancialAdjustment As Batch.AdjustmentTypes

      vCurrentPage = CType(pParams("CurrentPageType").IntegerValue, TraderPage.TraderPageType) 'Get the current page type
      Select Case pType
        Case TraderProcessDataTypes.tpdtFirstPage
          vNextPage = Me.MainPage
          If vNextPage = TraderPage.TraderPageType.tpPaymentMethod1 Then
            If PayMethodsAtEnd Then
              vNextPage = TraderPage.TraderPageType.tpTransactionDetails
            Else
              InitControls()
              If mvTraderPages(GetPageIndex(vNextPage)).MenuCount = 0 Then 'No types
                If MaintenanceOnly Then
                  vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis
                Else
                  vNextPage = TraderPage.TraderPageType.tpContactSelection
                End If
              End If
            End If
          End If
        Case TraderProcessDataTypes.tpdtNextPage
          SCSaveAnalysis(vCurrentPage, pParams, pTransaction)
          If pParams.ParameterExists("TransactionDateChanged").Bool Then ProcessVATRateChange(pTransaction, pParams.ParameterExists("TransactionDate").Value)
          vNextPage = SCGetNextPage(vCurrentPage, pParams) 'Get the next page to goto
          Select Case vNextPage
            Case TraderPage.TraderPageType.tpAddressMaintenance, TraderPage.TraderPageType.tpGiftAidDeclaration ', TraderPage.TraderPageType.tpPaymentPlanSummary
              If SupportsNonFinancialBatch Then
                InitNonFinancialTransaction()
                vResults.Add("NonFinancialBatchNumber", NonFinancialBatchNumber)
                vResults.Add("NonFinancialTransactionNumber", NonFinancialTransactionNumber)
                mvNonFinancialBatch.UpdateNumberOfTransactions(1)
              End If
            Case TraderPage.TraderPageType.tpBatchInvoiceSummary
              pTransaction.BatchInvoices = GetBatchInvoices(pParams)
          End Select

        Case TraderProcessDataTypes.tpdtPreviousPage
          vNextPage = SCGetPreviousPage(vCurrentPage, pParams) 'Get the previous page to goto
        Case TraderProcessDataTypes.tpdtFinished
          If pParams.ParameterExists("TransactionDateChanged").Bool Then ProcessVATRateChange(pTransaction, pParams.ParameterExists("TransactionDate").Value)
          vNextPage = SCDoFinished(vCurrentPage, pParams, pTransaction, vResults)
        Case TraderProcessDataTypes.tpdtEditTransaction
          SCGetTransactionData(pParams, pTransaction, vResults)
          vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)
          Select Case vFinancialAdjustment
            Case Batch.AdjustmentTypes.atMove

            Case Batch.AdjustmentTypes.atGIKConfirmation
              If PayMethodsAtEnd Then
                vNextPage = TraderPage.TraderPageType.tpTransactionDetails
              Else
                vNextPage = MainPage
              End If
            Case Batch.AdjustmentTypes.atCashBatchConfirmation
              vNextPage = TraderPage.TraderPageType.tpTransactionDetails 'Do not want to choose payment method
            Case Else
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
          End Select
        Case TraderProcessDataTypes.tpdtEditAnalysisLine
          vNextPage = SCDoEditAnalysis(vCurrentPage, pParams, pTransaction)
        Case TraderProcessDataTypes.tpdtCancelTransaction
          vNextPage = SCDoCancelled(vCurrentPage, pParams, pTransaction, vResults)
        Case TraderProcessDataTypes.tpdtDeleteAnalysisLine
          vNextPage = SCDoDeleteAnalysis(vCurrentPage, pParams, pTransaction, vResults)
        Case TraderProcessDataTypes.tpdtAddMemberSummary
          'We are adding a member to the MembershipMembersSummary grid
          vNextPage = TraderPage.TraderPageType.tpMembershipMembersSummary
          SCAddMemberSummary(pParams, pTransaction, vResults)
        Case TraderProcessDataTypes.tpdtAmendMemberSummary
          vNextPage = TraderPage.TraderPageType.tpAmendMembership
        Case TraderProcessDataTypes.tpdtDeletePaymentPlanLine
          vNextPage = SCDoDeletePaymentPlanLine(vCurrentPage, pParams, pTransaction)
      End Select
      vResults.Add("NextPageType", vNextPage)
      vResults.Add("NextPageCode", CDBField.FieldTypes.cftCharacter, GetPageCode(vNextPage))
      SCGetButtonStates(vNextPage, pParams, vResults)
      If (pType = TraderProcessDataTypes.tpdtNextPage Or pType = TraderProcessDataTypes.tpdtAmendMemberSummary Or _
         (pType = TraderProcessDataTypes.tpdtFinished And vNextPage <> TraderPage.TraderPageType.tpNone) Or _
         (pType = TraderProcessDataTypes.tpdtFirstPage And _
          ((vNextPage = TraderPage.TraderPageType.tpTransactionDetails) Or _
           (vNextPage = TraderPage.TraderPageType.tpGiveAsYouEarn) Or _
           (vNextPage = TraderPage.TraderPageType.tpPostTaxPGPayment))) Or _
         (pType = TraderProcessDataTypes.tpdtEditAnalysisLine And vNextPage = TraderPage.TraderPageType.tpAmendEventBooking)) Then
        SCGetDefaults(vNextPage, pParams, vResults, pTransaction)
      End If
      SCProcessTraderData = vResults
    End Function

    Private Sub SCGetButtonStates(ByRef pPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByRef pResults As CDBParameters)
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vEdit As Boolean
      Dim vNext As Boolean
      Dim vPrevious As Boolean
      Dim vFinished As Boolean
      Dim vPayMethod As Boolean
      Dim vTransPayMethod As String
      Dim vProvisional As Boolean
      Dim vConfirmed As Boolean

      'The following pages need to be handled for the first pass
      'BKD Bank Details,ING Batch Invoice Production,COM Comments,CCU Credit Customer,EVE Event Booking,INV Invoice Payments,CDC Card Details
      'PM1 Payment Method 1,PRD Product Details,TAS Transaction Analysis Summary,TRA Transaction Analysis,TRD Transaction Details

      vNext = True
      vPrevious = True
      vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)

      Select Case pPageType
        Case TraderPage.TraderPageType.tpBankDetails, TraderPage.TraderPageType.tpCreditCustomer
          vPayMethod = True
          If PayMethodsAtEnd Then
            If Not (pPageType = TraderPage.TraderPageType.tpBankDetails AndAlso pParams("TransactionPaymentMethod").Value = "CQIN") AndAlso _
               Not (pPageType = TraderPage.TraderPageType.tpCreditCustomer AndAlso pParams("TransactionPaymentMethod").Value = "CCIN") Then
              vNext = False
              vFinished = True
            End If
          End If

        Case TraderPage.TraderPageType.tpPurchaseOrderCancellation, TraderPage.TraderPageType.tpChequeNumberAllocation, TraderPage.TraderPageType.tpChequeReconciliation, TraderPage.TraderPageType.tpCreditStatementGeneration, TraderPage.TraderPageType.tpGiveAsYouEarn
          vFinished = True
          vPrevious = False
          vNext = False

        Case TraderPage.TraderPageType.tpBatchInvoiceProduction
          vFinished = False
          vPrevious = False
          vNext = True

        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          vFinished = True
          vPrevious = True
          vNext = False

        Case TraderPage.TraderPageType.tpCardDetails
          vNext = False
          vFinished = True

        Case TraderPage.TraderPageType.tpComments
          vTransPayMethod = pParams("TransactionPaymentMethod").Value
          If vTransPayMethod <> "CARD" AndAlso vTransPayMethod <> "CCIN" AndAlso vTransPayMethod <> "CAFC" AndAlso Not PayMethodsAtEnd Then
            vNext = False
            vFinished = True
          End If

        Case TraderPage.TraderPageType.tpPaymentMethod1
          vNext = False
          vPrevious = PayMethodsAtEnd

        Case TraderPage.TraderPageType.tpTransactionAnalysis, TraderPage.TraderPageType.tpPaymentMethod2
          vNext = False
          vPrevious = Not (vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or MaintenanceOnly)

        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary
          vConfirmed = True
          'Note: Enabling/Disabling of Edit & Delete buttons for each line will be handled by the Smart Client.
          '      Here we just want to set whether Edit & Delete are visible or not.
          If (PostedToCashBook And BatchType <> Access.Batch.GetBatchTypeCode(Batch.BatchTypes.CreditSales) And BatchType <> Access.Batch.GetBatchTypeCode(Batch.BatchTypes.StandingOrder) And BatchType <> Access.Batch.GetBatchTypeCode(Batch.BatchTypes.BankStatement) And (mvEnv.GetConfigOption("batch_bypass_cheque_list", False) = False) And vProvisional = False) Or (PostedToNominal And vConfirmed) Then
            vFinished = False
            vEdit = False
            vNext = False
            vPrevious = False
          Else
            vFinished = True
            vEdit = True
            If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
              vPrevious = False
              vFinished = False
            End If
          End If

        Case TraderPage.TraderPageType.tpTransactionDetails
          vTransPayMethod = pParams("TransactionPaymentMethod").Value
          Select Case vTransPayMethod
            Case "CHEQ"
              If BankDetails = False Then vPayMethod = True
            Case "CRED", "CQIN", "CCIN"
              'OK to do previous
            Case Else
              vPayMethod = True
          End Select
          Select Case vTransPayMethod
            Case "CASH", "CHEQ", "POST"
              If pParams.ParameterExists("ExistingTransaction").Bool = False And mvExistingAdjustmentTran = False Then vFinished = True
          End Select
          vPrevious = Not PayMethodsAtEnd

          '  Case  tpAddressMaintenance
        Case TraderPage.TraderPageType.tpSetStatus, TraderPage.TraderPageType.tpSuppressionEntry, TraderPage.TraderPageType.tpActivityEntry, _
             TraderPage.TraderPageType.tpGiveAsYouEarnEntry, TraderPage.TraderPageType.tpGiftAidDeclaration, TraderPage.TraderPageType.tpGoneAway, _
             TraderPage.TraderPageType.tpCancelGiftAidDeclaration, TraderPage.TraderPageType.tpCancelPaymentPlan
          vNext = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) > 1
          vFinished = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) < 2
          vPrevious = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) > 1 'mvTraderApplication.MaintenanceOnly

        Case TraderPage.TraderPageType.tpPostTaxPGPayment, TraderPage.TraderPageType.tpGiveAsYouEarn
          vNext = False
          vPrevious = False
          vFinished = True

        Case TraderPage.TraderPageType.tpContactSelection
          If pParams.ParameterExists("TransactionType").Value = "APAY" Then
            vPrevious = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) > 1
          Else
            vPrevious = False
          End If

        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority
          If AppType = ApplicationType.atConversion Then 'Convert
            vFinished = True
            vNext = False
          ElseIf pParams("TransactionType").Value = "APAY" Then
            vNext = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) > 1
            vFinished = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) < 2
            vPrevious = True
          ElseIf pParams("TransactionType").Value <> "CMEM" Then
            'this bit has been taken from frmtrader...cmdnext clcik on the dd page
            vFinished = True
            vNext = False
          End If

        Case TraderPage.TraderPageType.tpAddressMaintenance
          vNext = True 'AddressMaintenance always needs the Next button
          vFinished = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) < 2
          vPrevious = CountChoices(TraderPage.TraderPageType.tpTransactionAnalysis) > 1 'mvTraderApplication.MaintenanceOnly

          '--------------------------------------------------------------------------------------------------
          ' BELOW HERE NOT YET SUPPORTED
          '--------------------------------------------------------------------------------------------------

        Case TraderPage.TraderPageType.tpPurchaseInvoiceDetails, TraderPage.TraderPageType.tpPurchaseOrderDetails, TraderPage.TraderPageType.tpStatementList
          vPrevious = False

        Case TraderPage.TraderPageType.tpPaymentPlanSummary, TraderPage.TraderPageType.tpPurchaseInvoiceSummary, TraderPage.TraderPageType.tpPurchaseOrderSummary
          vFinished = True
          vEdit = True
        Case TraderPage.TraderPageType.tpPaymentMethod3
          vNext = False

        Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          vNext = False
          vFinished = False
        Case TraderPage.TraderPageType.tpPaymentPlanMaintenance
          vNext = True
          vPrevious = (ConversionShowPPD = True Or PayPlanConversionMaintenance = True)
          vFinished = Not vPrevious
          '
        Case TraderPage.TraderPageType.tpChangeMembershipType
          vNext = True
          vPrevious = False
          vFinished = False

        Case TraderPage.TraderPageType.tpScheduledPayments
          vPrevious = False
          vNext = False
          vFinished = True

        Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vNext = False

        Case TraderPage.TraderPageType.tpMembershipPayer
          vNext = False
          vFinished = True

        Case TraderPage.TraderPageType.tpAmendEventBooking
          vPrevious = False
          vNext = True
          vFinished = False
        Case TraderPage.TraderPageType.tpPostageAndPacking
          vPrevious = True
          vNext = True
          vFinished = False
        Case Else
          '
      End Select

      'If trying to go to the payment method 1 page then disable if editing or only one option
      If vPayMethod Then
        If pParams.ParameterExists("ExistingTransaction").Bool = True Or mvExistingAdjustmentTran = True Then
          vPrevious = False
        End If
      End If

      pResults.Add("NextButton", CDBField.FieldTypes.cftCharacter, If(vNext, "Enabled", "Disabled"))
      pResults.Add("FinishedButton", CDBField.FieldTypes.cftCharacter, If(vFinished, "Enabled", "Disabled"))
      pResults.Add("PreviousButton", CDBField.FieldTypes.cftCharacter, If(vPrevious, If(vPayMethod, "CheckPM1", "Enabled"), "Disabled"))

      pResults.Add("EditButton", CDBField.FieldTypes.cftCharacter, If(vEdit, "Visible", "Invisible"))
      pResults.Add("DeleteButton", CDBField.FieldTypes.cftCharacter, If(vEdit, "Visible", "Invisible"))
    End Sub

    Private Sub SCSaveAnalysis(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByRef pTransaction As TraderTransaction)
      Dim vRate As New ProductRate(mvEnv)
      Dim vDate As String
      Dim vWarehouse As String = ""
      Dim vTDRLine As TraderAnalysisLine = Nothing
      Dim vResult As New CDBParameters
      Dim vSLA As String = ""
      Dim vLineType As String = ""
      Dim vLineCount As Integer
      Dim vStockSales As Boolean
      Dim vStockIssued As Integer
      Dim vStockTransID As Integer
      Dim vInvoice As Invoice
      Dim vInvPayment As Invoice.InvoicePayment
      Dim vIndex As Integer
      Dim vNtypeTotal As Double
      Dim vUnAllocated As Double
      Dim vTransactionType As String
      Dim vOPS As OrderPaymentSchedule
      Dim vOOPS As OrderPaymentSchedule
      Dim vAmount As Double
      Dim vOPSAmount As Double
      Dim vPaymentNumber As String = ""
      Dim vPaymentType As String = ""
      Dim vSource As String = ""
      Dim vDeceasedContact As String
      Dim vTransType As String = ""
      Dim vProduct As Product
      Dim vCount As Integer
      Dim vRecCount As Integer
      Dim vContact As Contact
      Dim vStatus As String
      Dim vVATExclusive As Boolean
      Dim vProvisionalBT As BatchTransaction
      Dim vBTA As BatchTransactionAnalysis
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vFirstLine As Boolean
      Dim vCheckedValueFound As Boolean
      Dim vAddIncentives As Boolean
      Dim vStartLineNo As Integer
      Dim vRow As CDBDataRow

      vTransactionType = pParams.ParameterExists("TransactionType").Value
      vVATExclusive = False
      If pParams.Exists("Product") = True And pParams.Exists("Rate") = True Then
        vRate.Init((pParams("Product").Value), (pParams("Rate").Value))
        vVATExclusive = vRate.VatExclusive
      End If

      vAddIncentives = True

      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpEventBooking, TraderPage.TraderPageType.tpProductDetails, TraderPage.TraderPageType.tpAccommodationBooking,
             TraderPage.TraderPageType.tpCollectionPayments, TraderPage.TraderPageType.tpLegacyBequestReceipt,
             TraderPage.TraderPageType.tpAmendEventBooking, TraderPage.TraderPageType.tpExamBooking
          vLineCount = pParams("TransactionLines").IntegerValue + 1
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vTDRLine.Init(vLineCount, (pParams("TransactionType").Value))
      End Select

      Dim vBequest As New LegacyBequest(mvEnv)
      Dim vPOD As PurchaseOrderDetail
      Dim vPONumber As Integer
      Dim vPID As PurchaseInvoiceDetail
      Dim vPINumber As Integer
      Dim vGAYEPledge As New PreTaxPledge(mvEnv)
      Dim vPPD As PaymentPlanDetail
      Dim vLineNumber As Integer
      Dim vAmtOutstanding As Double
      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpExamBooking
          If pTransaction.ExamBookingLines.Rows.Count() > 0 Then
            vCount = 0
            For Each vRow In pTransaction.ExamBookingLines.Rows
              vCount = vCount + 1
              If vCount = 1 Then
                vTDRLine.AddExamBookingLine(vRow.Item("Product"), vRow.Item("Rate"), IntegerValue(vRow.Item("Quantity")), Val(vRow.Item("Amount")), pParams("TransactionSource").Value, vRow.Item("VATRate"), Val(vRow.Item("VATPercentage")), pParams("ExamBookingId").IntegerValue, vRow.Item("ExamUnitId"), vRow.Item("ExamUnitProductId"), vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value), vRow.Item("Notes"))
              Else
                'Add remainig lines
                vLineCount = vLineCount + 1
                vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
                vTDRLine.Init(vLineCount, (pParams("TransactionType").Value))
                vTDRLine.AddExamBookingLine(vRow.Item("Product"), vRow.Item("Rate"), IntegerValue(vRow.Item("Quantity")), Val(vRow.Item("Amount")), pParams("TransactionSource").Value, vRow.Item("VATRate"), Val(vRow.Item("VATPercentage")), pParams("ExamBookingId").IntegerValue, vRow.Item("ExamUnitId"), vRow.Item("ExamUnitProductId"), vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value), vRow.Item("Notes"))
              End If
            Next vRow
          End If

        Case TraderPage.TraderPageType.tpEventBooking, TraderPage.TraderPageType.tpAmendEventBooking
          vSource = pParams("TransactionSource").Value
          If pTransaction.EventBookingLines.Rows.Count() > 0 Then
            'Smart Client Event Booking on an Event using the Pricing Matrix
            vCount = 0
            For Each vRow In pTransaction.EventBookingLines.Rows
              vCount = vCount + 1
              If vCount = 1 Then
                'Fist line is the Event Booking line
                vTDRLine.AddEventBooking(vRow.Item("Product"), vRow.Item("Rate"), IntegerValue(vRow.Item("Quantity")), Val(vRow.Item("Amount")), pParams("TransactionSource").Value, vRow.Item("VATRate"), Val(vRow.Item("VATPercentage")), pParams("BookingNumber").IntegerValue, vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value), vRow.Item("Notes"))
                vTDRLine.SetEventBookingInfo(pParams("EventNumber").IntegerValue, pParams("OptionNumber").IntegerValue, pParams.ParameterExists("AdultQuantity").Value, pParams.ParameterExists("ChildQuantity").Value, pParams.ParameterExists("StartTime").Value, pParams.ParameterExists("EndTime").Value, pParams.ParameterExists("AmendedBookingNumber").IntegerValue)
              Else
                'Add remainig lines
                vLineCount = vLineCount + 1
                vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
                vTDRLine.Init(vLineCount, (pParams("TransactionType").Value))
                vTDRLine.AddEventBookingPriceMatrixLine(vRow.Item("Product"), vRow.Item("Rate"), IntegerValue(vRow.Item("Quantity")), Val(vRow.Item("Amount")), vRow.Item("VATRate"), Val(vRow.Item("VATAmount")), Val(vRow.Item("VATPercentage")), pParams("TransactionSource").Value, vRow.Item("Notes"), pParams("BookingNumber").IntegerValue, pParams.ParameterExists("DistributionCode").Value)
              End If
            Next vRow
          Else
            vTDRLine.AddEventBooking(pParams("Product").Value, pParams("Rate").Value, pParams("Quantity").IntegerValue, pParams("Amount").DoubleValue, pParams("TransactionSource").Value, pParams("VatRate").Value, pParams("VatPercentage").DoubleValue, pParams("BookingNumber").IntegerValue, vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value))
            vTDRLine.SetEventBookingInfo(pParams("EventNumber").IntegerValue, pParams("OptionNumber").IntegerValue, pParams.ParameterExists("AdultQuantity").Value, pParams.ParameterExists("ChildQuantity").Value, pParams.ParameterExists("StartTime").Value, pParams.ParameterExists("EndTime").Value, pParams.ParameterExists("AmendedBookingNumber").IntegerValue)
          End If

        Case TraderPage.TraderPageType.tpAccommodationBooking
          vTDRLine.AddAccomodationBooking(pParams("Product").Value, pParams("Rate").Value, pParams("Quantity").IntegerValue, pParams("Amount").DoubleValue, pParams("TransactionSource").Value, pParams.ParameterExists("VatRate").Value, pParams.ParameterExists("VatPercentage").DoubleValue, pParams("RoomBookingNumber").IntegerValue, vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value))
          vSource = pParams("TransactionSource").Value

        Case TraderPage.TraderPageType.tpCollectionPayments
          vTDRLine.AddCollectionPayment(pParams("AppealCollectionNumber").IntegerValue, pParams("Product").Value, pParams("Rate").Value, pParams("Source").Value, pParams("Amount").DoubleValue, pParams("VatRate").Value, pParams("VatPercentage").DoubleValue, pParams("BankAccount").Value, vVATExclusive, (pParams.ParameterExists("Notes").Value), (pParams.ParameterExists("CollectionPISNumber").IntegerValue), (pParams.ParameterExists("DeceasedContactNumber").IntegerValue), (pParams.ParameterExists("CollectionBoxNumbers").Value), (pParams.ParameterExists("CollectionBoxAmounts").Value))

        Case TraderPage.TraderPageType.tpLegacyBequestReceipt
          vBequest.Init((pParams("LegacyNumber").IntegerValue), (pParams("BequestNumber").IntegerValue))
          vContact = New Contact(mvEnv)
          vContact.Init((pParams("PayerContactNumber").IntegerValue))
          If vBequest.Existing And vContact.Existing Then
            vTDRLine.AddLegacyBequestReceipt(vBequest.LegacyNumber, vBequest.BequestNumber, vBequest.ProductCode, vBequest.RateCode, pParams("Amount").DoubleValue, pParams("DateReceived").Value, pParams("TransactionSource").Value, vBequest.VatRate(vContact).VatRateCode, vBequest.VatRate(vContact).Percentage, vContact.ContactNumber, pParams("PayerAddressNumber").IntegerValue, (pParams("Notes").Value), (vBequest.DistributionCode))
            'What if it does not exist?
          End If
          vSource = pParams("TransactionSource").Value

        Case TraderPage.TraderPageType.tpPurchaseOrderProducts

          If pParams.Exists("LineNumber") Then
            If pTransaction.PODExists(pParams("LineNumber").IntegerValue) Then vLineCount = pParams("LineNumber").IntegerValue
          Else
            vLineCount = pParams("PPDLines").IntegerValue + 1
            pParams.Add("LineNumber", vLineCount)
          End If
          vPONumber = pParams.ParameterExists("PurchaseOrderNumber").IntegerValue
          vPOD = pTransaction.GetPurchaseOrderDetail(vLineCount, vPONumber)

          vPOD.Create(mvEnv, pParams)

          vAddIncentives = False

        Case TraderPage.TraderPageType.tpPurchaseInvoiceProducts
          If pParams.Exists("LineNumber") Then
            If pTransaction.PIDExists(pParams("LineNumber").IntegerValue) Then vLineCount = pParams("LineNumber").IntegerValue
          Else
            vLineCount = pParams("PPDLines").IntegerValue + 1
            pParams.Add("LineNumber", vLineCount)
          End If
          vPINumber = pParams.ParameterExists("PurchaseInvoiceNumber").IntegerValue
          vPID = pTransaction.GetPurchaseInvoiceDetail(vLineCount, vPINumber)
          vPID.Create(mvEnv, pParams)

          vAddIncentives = False

        Case TraderPage.TraderPageType.tpActivityEntry, TraderPage.TraderPageType.tpSuppressionEntry, TraderPage.TraderPageType.tpSetStatus, _
             TraderPage.TraderPageType.tpGiftAidDeclaration, TraderPage.TraderPageType.tpGiveAsYouEarnEntry, _
             TraderPage.TraderPageType.tpGoneAway, TraderPage.TraderPageType.tpAddressMaintenance, _
             TraderPage.TraderPageType.tpCancelGiftAidDeclaration, TraderPage.TraderPageType.tpCancelPaymentPlan
          If pParams("TransactionLines").IntegerValue > 0 Then
            vLineCount = pParams("TransactionLines").IntegerValue + 1
          Else
            vLineCount = pTransaction.TraderAnalysisLines.Count + 1
          End If
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount, pParams("TransactionType").Value)
          Select Case pCurrentPageType
            Case TraderPage.TraderPageType.tpActivityEntry
              vTDRLine.AddActivity(pParams("ContactNumber").IntegerValue, DefaultActivityGroup)
            Case TraderPage.TraderPageType.tpSuppressionEntry
              vTDRLine.AddSuppression(pParams("ContactNumber").IntegerValue, DefaultSuppression)
            Case TraderPage.TraderPageType.tpSetStatus
              vDate = ""
              vStatus = pParams.ParameterExists("Status2").Value
              If Len(vStatus) > 0 Then
                vDate = pParams.ParameterExists("TransactionDate").Value
                If Not (IsDate(vDate)) Then vDate = TodaysDate()
              End If
              vTDRLine.AddContactStatus(pParams("ContactNumber").IntegerValue, vStatus, vDate)
            Case TraderPage.TraderPageType.tpGiftAidDeclaration
              vTDRLine.AddGiftAidDeclaration(pParams("ContactNumber").IntegerValue, pParams("AddressNumber").IntegerValue, pParams("StartDate").Value, pParams("Source").Value, (pParams("Notes").Value))
            Case TraderPage.TraderPageType.tpGiveAsYouEarnEntry
              vContact = New Contact(mvEnv)
              vContact.Init((pParams("ContactNumber").IntegerValue))
              vGAYEPledge.Init()
              vGAYEPledge.InitNewPledge(pParams("ContactNumber").IntegerValue, pParams("OrganisationNumber").IntegerValue, pParams("AgencyNumber").IntegerValue, pParams("WorkAddressNumber").IntegerValue, pParams("PaybillAddressNumber").IntegerValue, pParams("DonorId").Value, pParams("PfoCode").Value, pParams("PledgeAmount").DoubleValue, pParams("StartDate").Value, pParams("Product").Value, pParams("Rate").Value, pParams("Source").Value, pParams("PayrollOrganisationNumber").IntegerValue, pParams("DistributionCode").Value, pParams("PaymentFrequency").Value, pParams("PayrollNumber").Value, vContact.Address.AddressNumber)
              If SupportsNonFinancialBatch Then InitNonFinancialTransaction()
              vGAYEPledge.Save(mvEnv.User.Logname, False, NonFinancialBatchNumber, NonFinancialTransactionNumber)
              If SupportsNonFinancialBatch Then mvNonFinancialBatch.UpdateNumberOfTransactions(1)
              vTDRLine.AddPayrollGivingPledge(pParams("ContactNumber").IntegerValue, pParams("WorkAddressNumber").IntegerValue, pParams("StartDate").Value, pParams("PledgeAmount").DoubleValue, pParams("DonorId").Value, pParams("Product").Value, pParams("Rate").Value, (pParams("DistributionCode").Value))
            Case TraderPage.TraderPageType.tpGoneAway
              vTDRLine.AddGoneAway(pParams("ContactNumber").IntegerValue)
            Case TraderPage.TraderPageType.tpAddressMaintenance
              vTDRLine.AddAddressUpdate(pParams("ContactNumber").IntegerValue)
            Case TraderPage.TraderPageType.tpCancelPaymentPlan
              If pParams.Exists("TransactionDate") Then 'BR16743 - If Tranaction Date is not present let AddPaymentPlanCancellation use its default Cancellation Date
                vTDRLine.AddPaymentPlanCancellation(pParams("PaymentPlanNumber").IntegerValue, pParams("CancellationReason").Value, pParams("TransactionDate").Value, pParams("CancellationSource").Value)
              Else
                vTDRLine.AddPaymentPlanCancellation(pParams("PaymentPlanNumber").IntegerValue, pParams("CancellationReason").Value, String.Empty, pParams("CancellationSource").Value)
              End If
            Case TraderPage.TraderPageType.tpCancelGiftAidDeclaration
              vDate = pParams.ParameterExists("TransactionDate").Value
              If Not (IsDate(vDate)) Then vDate = TodaysDate()
              vTDRLine.AddGiftAidDeclarationCancellation(pParams("DeclarationNumber").IntegerValue, pParams("ContactNumber").IntegerValue, pParams("CancellationReason").Value, vDate, pParams("CancellationSource").Value)
          End Select
        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority
          If pParams.ParameterExists("TransactionType").Value = "APAY" Then
            SaveAutoPaymentMethodChanges(pCurrentPageType, pParams, pTransaction)
          End If
        Case TraderPage.TraderPageType.tpConfirmProvisionalTransactions
          vProvisionalBT = New BatchTransaction(mvEnv)
          With vProvisionalBT
            .InitBatchTransactionAnalysis((pParams("BatchNumber").IntegerValue), (pParams("TransactionNumber").IntegerValue))
            'add each provisional analysis line to the grid
            vLineCount = pParams("TransactionLines").IntegerValue + 1
            For Each vBTA In .Analysis
              vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
              vAmount = Val(mvEnv.Connection.GetValue("SELECT percentage FROM vat_rates WHERE vat_rate = '" & vBTA.VatRate & "'"))
              vTDRLine.Init(vLineCount, (pParams.ParameterExists("TransactionType").Value))
              vTDRLine.ConfirmProvisionalTransaction(vBTA, vAmount, (pParams.ParameterExists("ContactDiscount").Bool), (pParams.ParameterExists("StockSale").Bool))
              vLineCount = vLineCount + 1
            Next vBTA
          End With
          vSource = pParams.ParameterExists("TransactionSource").Value

        Case TraderPage.TraderPageType.tpProductDetails
          'Set the Date
          vStockSales = pParams.ParameterExists("StockSale").Bool
          vDate = pParams("When").Value
          If vStockSales And Not (IsDate(vDate)) Then
            If vStockIssued = 0 Then
              vDate = pParams("TransactionDate").Value 'should be earliest stock delivery date but at the moment the stock lookup DLL doesn't give us that
            Else
              If pParams("TransactionPaymentMethod").Value = "CARD" OrElse pParams("TransactionPaymentMethod").Value = "CCIN" Then
                vDate = CDate(Me.BatchDate).ToString(CAREDateFormat)
              Else
                vDate = TodaysDate()
              End If
            End If
          End If
          'For product sales, set the Warehouse
          If vStockSales Then
            vStockIssued = pParams("StockIssued").IntegerValue
            vStockTransID = pParams("StockTransactionID").IntegerValue
            If pParams.Exists("Warehouse") Then
              vWarehouse = pParams("Warehouse").Value
            Else
              vWarehouse = mvEnv.Connection.GetValue("SELECT warehouse FROM products WHERE product = '" & pParams("Product").Value & "'")
            End If
          End If

          If pParams("TransactionType").Value = "CRDN" Then vSLA = pParams("SalesLedgerAccount").Value
          If pParams.ParameterExists("DeceasedContactNumber").IntegerValue > 0 Then
            vLineType = pParams.ParameterExists("LineType").Value
            If vLineType.Length = 0 Then
              If pParams.ParameterExists("LineTypeG").Bool Then
                If pParams.ParameterExists("LineTypeH").Bool Then
                  vLineType = "D"     'Gift InMemoriam & HardCredit
                ElseIf pParams.ParameterExists("LineTypeS").Bool Then
                  vLineType = "F"     'Gift InMemoriam & SoftCredit
                Else
                  vLineType = "G"     'Gift InMemoriam
                End If
              ElseIf pParams.ParameterExists("LineTypeH").Bool Then
                vLineType = "H"       'Hard Credit
              ElseIf pParams.ParameterExists("LineTypeS").Bool Then
                vLineType = "S"       'Soft Credit
              End If
            End If
          End If

          If vStockSales Then
            vTDRLine.AddStockProductSale(pParams("Product").Value, pParams("Rate").Value, vWarehouse, pParams("Quantity").IntegerValue, vStockIssued, pParams("Amount").DoubleValue, vDate, pParams("Source").Value, pParams("DespatchMethod").Value, pParams("ContactNumber").IntegerValue, pParams("AddressNumber").IntegerValue, pParams("VatRate").Value, pParams("VatPercentage").DoubleValue, pParams("VatAmount").DoubleValue, "", vVATExclusive, (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value), (pParams.ParameterExists("Notes").Value), (pParams.ParameterExists("ProductNumber").Value), (pParams("ContactDiscount").Bool), (pParams.ParameterExists("GrossAmount").DoubleValue), (pParams.ParameterExists("Discount").DoubleValue), (pParams.ParameterExists("DeceasedContactNumber").Value), vLineType, vSLA, vStockTransID, (pParams.ParameterExists("ServiceBookingNumber").IntegerValue), (pParams.ParameterExists("EventBookingNumber").IntegerValue))
          Else
            vTDRLine.AddNonStockProductSale(pParams("Product").Value, pParams("Rate").Value, pParams("Quantity").IntegerValue, pParams("Amount").DoubleValue, vDate, pParams("Source").Value, pParams("ContactNumber").IntegerValue, pParams("AddressNumber").IntegerValue, pParams("VatRate").Value, pParams("VatPercentage").DoubleValue, pParams("VatAmount").DoubleValue, vVATExclusive, (pParams.ParameterExists("DespatchMethod").Value), (pParams.ParameterExists("DistributionCode").Value), (pParams.ParameterExists("SalesContactNumber").Value), (pParams.ParameterExists("Notes").Value), (pParams.ParameterExists("ProductNumber").Value), (pParams("ContactDiscount").Bool), (pParams.ParameterExists("GrossAmount").DoubleValue), (pParams.ParameterExists("Discount").DoubleValue), (pParams.ParameterExists("DeceasedContactNumber").Value), vLineType, vSLA, (pParams.ParameterExists("ServiceBookingNumber").IntegerValue), (pParams.ParameterExists("EventBookingNumber").IntegerValue), pParams.ParameterExists("ScheduledPaymentNumber").IntegerValue, pParams.ParameterExists("CreditedContactNumber").Value)
          End If

          If CSDepositPercentage > 0 Then pTransaction.TraderAnalysisLines.SetDepositAllowed(mvEnv)

          vSource = vTDRLine.Source

        Case TraderPage.TraderPageType.tpInvoicePayments
          Dim vDiffPayerOfUnallocatedCash As Boolean = False
          vLineCount = pParams("TransactionLines").IntegerValue

          'Instantiate the Credit Customer in order to store the Contact Number and Address Number on the Trader Analysis Line
          Dim vCC As New CreditCustomer()
          vCC.InitCompanySalesLedgerAccount(mvEnv, pParams.ParameterExists("Company").Value, pParams.ParameterExists("SalesLedgerAccount").Value)
          If Not vCC.Existing Then
            RaiseError(DataAccessErrors.daeCreditCustomerMissing1, pParams.ParameterExists("Company").Value, pParams.ParameterExists("SalesLedgerAccount").Value)
          End If
          Dim vSLAccount As String = vCC.SalesLedgerAccount

          For Each vInvoice In pTransaction.TraderInvoiceLines
            If vInvoice.RecordType = "I" Then
              'This is the invoice being paid
              If vInvoice.NowPaid > 0 Then
                vInvoice.SetInvoiceNumber(True, True)
                'loop through each payment creating a line in the TAS grid
                For vIndex = 1 To vInvoice.InvoicePayments.Count
                  vInvPayment = CType(vInvoice.InvoicePayments(vIndex), Invoice.InvoicePayment)
                  Dim vCurrencyAmount As Double
                  If Me.BatchLedApp Then 'BR20811 Invoices are in local currency ONLY but Batches can be in any currency so the exchange rate for the batch needs to be applied.
                    vCurrencyAmount = FixTwoPlaces(vInvPayment.AmountPaid * Me.BatchExchangeRate)
                  Else
                    vCurrencyAmount = vInvPayment.AmountPaid
                  End If
                  vLineCount = vLineCount + 1
                  vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
                  vTDRLine.AddInvoicePayment(IntegerValue(vInvoice.InvoiceNumber), vCurrencyAmount, pParams("TransactionSource").Value, vCC.SalesLedgerAccount, vInvPayment.InvoiceNumberUsed, vInvPayment.InvoiceNumberUsed, pParams.ParameterExists("TransactionDistributionCode").Value, vCC.ContactNumber, vCC.AddressNumber, vInvPayment.RecordType)
                  If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment Then vNtypeTotal = vNtypeTotal + vInvPayment.AmountPaid
                  If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation Then
                    'Create this 2nd -ve line for InvoiceAllocation (L-type line) and for SundryCreditNoteInvoiceAllocation (K-type line) 
                    vSLAccount = vCC.SalesLedgerAccount
                    If Not (String.IsNullOrWhiteSpace(vInvPayment.SalesLedgerAccount)) Then vSLAccount = vInvPayment.SalesLedgerAccount
                    Dim vPayerCC As CreditCustomer = vCC
                    If vCC.SalesLedgerAccount.Equals(vSLAccount, StringComparison.InvariantCultureIgnoreCase) = False Then
                      'Unallocated SL Cash was from a different Contact so negative line needs to debit their account
                      vPayerCC = New CreditCustomer()
                      vPayerCC.InitCompanySalesLedgerAccount(mvEnv, pParams.ParameterExists("Company").Value, vSLAccount)
                      If vPayerCC.Existing = False Then RaiseError(DataAccessErrors.daeCreditCustomerMissing1, pParams.ParameterExists("Company").Value, vSLAccount)
                      vDiffPayerOfUnallocatedCash = True
                    End If
                    vLineCount = vLineCount + 1
                    vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
                    vTDRLine.AddInvoicePayment(IntegerValue(vInvoice.InvoiceNumber), vCurrencyAmount * -1, pParams("TransactionSource").Value, vPayerCC.SalesLedgerAccount, vInvPayment.InvoiceNumberUsed, vInvPayment.InvoiceNumberUsed, pParams.ParameterExists("TransactionDistributionCode").Value, vPayerCC.ContactNumber, vPayerCC.AddressNumber, vInvPayment.RecordType)
                  End If
                Next
              End If
            End If
          Next

          'If any portion of the current payment is unallocated then create a U-type line for that amount here
          vUnAllocated = pParams("CurrentUnAllocated").DoubleValue
          If pParams("InvoicePaymentAmount").DoubleValue <> 0 And vUnAllocated <> 0 Then
            Dim vUnallocatedCurrencyAmount As Double
            If Me.BatchLedApp Then 'BR20811 Invoices are in local currency ONLY but Batches can be in any currency so the exchange rate for the batch needs to be applied.
              vUnallocatedCurrencyAmount = FixTwoPlaces(vUnAllocated * Me.BatchExchangeRate)
            Else
              vUnallocatedCurrencyAmount = vUnAllocated
            End If
            vLineCount = vLineCount + 1
            vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
            vTDRLine.AddUnallocatedSalesledgerPayment(vUnallocatedCurrencyAmount, pParams("TransactionSource").Value, pParams("SalesLedgerAccount").Value, pParams.ParameterExists("DistributionCode").Value, vCC.ContactNumber, vCC.AddressNumber)
          End If

          'Update both the invoices being paid and the unallocated-cash invoices being used
          mvEnv.Connection.StartTransaction()
          '1.update invoices being paid

          For Each vInvoice In pTransaction.TraderInvoiceLines
            With vInvoice
              If .RecordType = "I" Then
                'This is the invoice being paid
                If .NowPaid > 0 Then
                  vInvoice.SCUpdatePayment()
                End If
              Else
                '2.update the cash invoices that were used
                If IntegerValue(vInvoice.InvoiceNumber) > 0 And .AmountUsed > 0 Then
                  vInvoice.SCUpdatePayment()
                End If
              End If
            End With
          Next

          If vNtypeTotal > 0 Or vUnAllocated > 0 Then
            '3.update the Outstanding attribute on the Credit Customers table
            UpdateOutstanding(pParams("Company").Value, pParams("SalesLedgerAccount").Value, (FixTwoPlaces(vNtypeTotal) + Val(vUnAllocated)) * -1, True)
          End If

          If vDiffPayerOfUnallocatedCash Then
            '4. Used unallocated SL cash from different CC so update both records accordingly
            For Each vInvoice In pTransaction.TraderInvoiceLines
              If vInvoice.RecordType.Equals("I", StringComparison.InvariantCultureIgnoreCase) AndAlso vInvoice.NowPaid > 0 Then
                For vIndex = 1 To vInvoice.InvoicePayments.Count
                  vInvPayment = CType(vInvoice.InvoicePayments(vIndex), Invoice.InvoicePayment)
                  If vCC.SalesLedgerAccount.Equals(If(String.IsNullOrWhiteSpace(vInvPayment.SalesLedgerAccount), vCC.SalesLedgerAccount, vInvPayment.SalesLedgerAccount), StringComparison.InvariantCultureIgnoreCase) = False Then
                    UpdateOutstanding(vCC.Company, vCC.SalesLedgerAccount, (vInvPayment.AmountPaid * -1))    'Credit account for invoice we paid
                    UpdateOutstanding(vCC.Company, vInvPayment.SalesLedgerAccount, vInvPayment.AmountPaid)   'Debit account for payer of unallocated cash
                  End If
                Next
              End If
            Next
          End If

          mvEnv.Connection.CommitTransaction()
          vSource = pParams("TransactionSource").Value

        Case TraderPage.TraderPageType.tpPaymentPlanProducts, TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance, TraderPage.TraderPageType.tpPaymentPlanDetails
          'determine if payment plan already has the maximum number of products that use product numbers
          If pCurrentPageType = TraderPage.TraderPageType.tpPaymentPlanProducts Or pCurrentPageType = TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance Then
            vProduct = New Product(mvEnv)
            vProduct.Init((pParams("Product").Value))
            If vProduct.UsesProductNumbers Then
              vCount = pParams("PPDProductNumbersCount").IntegerValue
              If vCount = vProduct.MaxNumbersAllowed Then
                RaiseError(DataAccessErrors.daePPMaxProdNumbers, CStr(vProduct.MaxNumbersAllowed))
              Else
                vContact = New Contact(mvEnv)
                vContact.Init((pParams.Item("PayerContactNumber").IntegerValue))
                vRecCount = mvEnv.Connection.GetCount("order_details od, orders o", Nothing, "od.contact_number = " & pParams("PayerContactNumber").Value & " AND product_number IS NOT NULL AND od.order_number = o.order_number AND cancellation_reason IS NULL")
                If (vCount + vRecCount) >= vProduct.MaxNumbersAllowed Then
                  RaiseError(DataAccessErrors.daeMaxProdNumbers, If(MembersOnly, "Member", If(vContact.ContactType = Contact.ContactTypes.ctcOrganisation, "Organisation", "Contact")), CStr(vProduct.MaxNumbersAllowed)) '%s already has %s product(s) of this type - no more allowed
                End If
              End If
            End If
            vLineNumber = pParams("PPDLines").IntegerValue + 1
            vPPD = pTransaction.GetPaymentPlanDetail(vLineNumber)
            If pParams.Exists("Amount") Then pParams.Add("DetailFixedAmount", CDBField.FieldTypes.cftNumeric, pParams("Amount").Value)
            vPPD.CreateSC(pParams)

          ElseIf pCurrentPageType = TraderPage.TraderPageType.tpPaymentPlanDetails And (vTransactionType = "SALE" Or vTransactionType = "EVNT" Or vTransactionType = "ACOM" Or vTransactionType = "SRVC") And Me.PaymentPlanDetails.Count = 0 Then
            vLineNumber = pParams("PPDLines").IntegerValue + 1
            'Create Payment Plan from unbalanced Transaction
            For Each vTDRLine In pTransaction.TraderAnalysisLines
              vPPD = GetPaymentPlanDetail(vLineNumber)
              vPPD.Init(mvEnv)
              vPPD.Create(0, 0, pParams("PayerContactNumber").IntegerValue, pParams("PayerAddressNumber").IntegerValue, _
              vTDRLine.ProductCode, vTDRLine.RateCode, vTDRLine.Quantity, vTDRLine.Amount, , , , _
                           pParams("Source").Value, vTDRLine.DistributionCode)
              vPPD.LineNumber = vLineNumber
              pTransaction.TraderPPDLines.AddItem(vPPD, vLineNumber.ToString)
              vLineNumber = vLineNumber + 1
            Next
          End If
          '
          'Find any relevant incentives
          If SCGetNextPage(pCurrentPageType, pParams) = TraderPage.TraderPageType.tpPaymentPlanSummary And vTransactionType <> "MEMB" And pParams.ParameterExists("CheckIncentives").Bool And AppType <> ApplicationType.atMaintenance Then
            ProcessIncentives(pParams, pTransaction, pTransaction.PaymentPlan, pParams("PayerContactNumber").IntegerValue)
          End If
          vAddIncentives = False

        Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vAmount = pParams("Amount").DoubleValue
          vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)
          If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
            vLineCount = pParams("LineNumber").IntegerValue
            vFirstLine = True
          Else
            vLineCount = pParams("TransactionLines").IntegerValue + 1
          End If
          'Set payment amount
          vOPSAmount = 0
          If Len(pParams.ParameterExists("MemberNumber").Value) > 0 Then
            vPaymentNumber = pParams("MemberNumber").Value
            vPaymentType = "M"
          ElseIf pParams.ParameterExists("CovenantNumber").IntegerValue > 0 Then
            vPaymentNumber = pParams("CovenantNumber").Value
            vPaymentType = "C"
          ElseIf pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
            vPaymentNumber = pParams("PaymentPlanNumber").Value
            vPaymentType = "O"
          End If

          If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
            pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
          End If
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vTDRLine.Init(vLineCount, vPaymentType) 'IIf((mvCurrPageType = tpPayments Or mvCurrPageType = tpOutstandingScheduledPayments), mvPaymentType, mvTransactionType)

          vSource = pParams("TransactionSource").Value
          vDeceasedContact = pParams.ParameterExists("DeceasedContact").Value
          'updatedOPS is the OPS xoming back from the client

          vAmtOutstanding = pParams("AmountOutstanding").DoubleValue
          While vAmtOutstanding > 0
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv)
            vOPS.CreateInAdvance(mvEnv, pTransaction.PaymentPlan, pParams("AmountOutstanding").DoubleValue, False)
            If vOPS.AmountOutstanding = 0 Then
              'If we have landed up with a zero amount record, Payment Plan may be free
              'So just get a record with the amount we want so that we don't get suck in an endless loop
              vOPS = New OrderPaymentSchedule
              vOPS.Init(mvEnv)
              vOPS.CreateInAdvance(mvEnv, pTransaction.PaymentPlan, pParams("AmountOutstanding").DoubleValue, True)
            End If
            pTransaction.UpdatedOPS.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
            PayScheduledPayment(vOPS, vAmtOutstanding)
          End While

          If pTransaction.UpdatedOPS.Count() > 0 Then
            pTransaction.PaymentPlan.Init(mvEnv, CType(pTransaction.UpdatedOPS.Item(1), OrderPaymentSchedule).PlanNumber)
            vCheckedValueFound = True 'To make following loop run first time
            Dim vLoan As Boolean = (pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan)
            Do While vAmount > 0
              If vCheckedValueFound = True Then
                vCheckedValueFound = False
                For Each vOOPS In pTransaction.UpdatedOPS
                  If vOOPS.SCCheckValue Then
                    vCheckedValueFound = True
                    If vOOPS.PaymentAmount > 0 Then
                      vOPS = New OrderPaymentSchedule
                      vOPS.Init(mvEnv, (vOOPS.ScheduledPaymentNumber))
                      If vOPS.Existing Then
                        If (vOOPS.PaymentAmount > pTransaction.PaymentPlan.Balance) And pTransaction.PaymentPlan.Balance > 0 And (vLineCount = 1 Or vFirstLine) Then
                          'First line must not exceed balance (Probably only RCPCH FirstAmount set-up anyway)
                          vOPSAmount = pTransaction.PaymentPlan.Balance
                        Else
                          vOPSAmount = vOOPS.PaymentAmount
                        End If
                        vOPS.SetUnProcessedPayment(True, vOPSAmount, vLoan)
                        vOPS.Save()
                        vAmount = FixTwoPlaces(vAmount - vOPSAmount)
                        vOOPS.PaymentAmount = FixTwoPlaces(vOOPS.PaymentAmount - vOPSAmount)
                        vTDRLine.AddPaymentPlanPayment(pTransaction.PaymentPlan.PlanNumber, vPaymentNumber, vOPS.ScheduledPaymentNumber, vOPSAmount, vSource, vAmount = 0 And pParams("AcceptAsFull").Value = "Y", If((vTransType Like "[SGH]" And Len(vDeceasedContact) > 0), vDeceasedContact, ""), If((vTransType Like "[SGH]" And Len(vDeceasedContact) > 0), vPaymentType, ""), pTransaction.PaymentPlan.GiverContactNumber, vTransType, pParams.OptionalValue("DistributionCode", mvDistributionCode), (pParams.ParameterExists("SalesContactNumber").Value))
                        If vAmount > 0 Then
                          If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
                            If pParams("TransactionLines").IntegerValue > 0 Then
                              vLineCount = pParams("TransactionLines").IntegerValue + 1
                            Else
                              vLineCount = vLineCount + 1
                            End If
                            vFirstLine = False
                          Else
                            vLineCount = vLineCount + 1
                          End If
                          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
                          vTDRLine.Init(vLineCount, vPaymentType)
                        End If
                      End If
                    End If
                  End If
                Next vOOPS
              Else
                Exit Do
              End If
            Loop
          End If

          'Not all of the payment has been allocated against the payment schedule
          'we have come here so the user must have decided to allocate the remainder to advance

        Case TraderPage.TraderPageType.tpServiceBooking
          Dim vVATRate As VatRate
          Dim vSBValue As Integer
          Dim vStartDate As String
          Dim vCurrentPrice As Double
          Dim vProductOffer As New ProductOffer(mvEnv)

          vProductOffer.Init(pParams("Product").Value, pParams("Rate").Value)
          vContact = New Contact(mvEnv)
          vContact.Init((pParams("PayerContactNumber").IntegerValue))
          vProduct = New Product(mvEnv)
          vProduct.InitWithRate(mvEnv, pParams("Product").Value, pParams("Rate").Value)
          vVATRate = mvEnv.VATRate(vProduct.ProductVatCategory, vContact.VATCategory)
          vSource = pParams.ParameterExists("TransactionSource").Value

          vLineCount = pParams("TransactionLines").IntegerValue + 1
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vTDRLine.Init(vLineCount, vTransactionType)

          'in the following code the ABS function is used because the Service Booking Credits
          'option creates a service booking w/ a -ve value, but the transaction should contain
          '+ve values; the transaction should look just like a sundry credit note transaction.
          If mvServiceBookingCredits Then 'If mvTraderApplication.ServiceBookingCredits Then
            If vProductOffer.EntitlementProduct.Length > 0 AndAlso vProductOffer.EntitlementRate.Length > 0 AndAlso pParams("SBEntitlementQty").Value.Length > 0 Then
              vTDRLine.AddServiceBookingCredit(pParams("Product").Value, pParams("Rate").Value, pParams("SBGrossQty").IntegerValue, pParams("SBGrossAmount").DoubleValue, pParams("Source").Value, vVATRate.VatRateCode, vVATRate.CurrentPercentage(pParams("StartDate").Value), pParams("ServiceBookingNumber").IntegerValue, pParams("VATExclusive").Bool, pParams.ParameterExists("TransactionDistributionCode").Value, pParams.ParameterExists("SalesContactNumber").Value)
            Else
              vTDRLine.AddServiceBookingCredit(pParams("Product").Value, pParams("Rate").Value, pParams("SBGrossQty").IntegerValue, pParams("Amount").DoubleValue, pParams("Source").Value, vVATRate.VatRateCode, vVATRate.CurrentPercentage(pParams("StartDate").Value), pParams("ServiceBookingNumber").IntegerValue, pParams("VATExclusive").Bool, pParams.ParameterExists("TransactionDistributionCode").Value, pParams.ParameterExists("SalesContactNumber").Value, pParams.ParameterExists("ContactDiscount").Bool, pParams.ParameterExists("GrossAmount").DoubleValue, pParams.ParameterExists("Discount").DoubleValue)
            End If
          Else
            If vProductOffer.EntitlementProduct.Length > 0 AndAlso vProductOffer.EntitlementRate.Length > 0 AndAlso pParams("SBEntitlementQty").Value.Length > 0 Then
              vTDRLine.AddServiceBooking(pParams("Product").Value, pParams("Rate").Value, pParams("SBGrossQty").IntegerValue, pParams("SBGrossAmount").DoubleValue, pParams("Source").Value, vVATRate.VatRateCode, vVATRate.CurrentPercentage(pParams("StartDate").Value), pParams("ServiceBookingNumber").IntegerValue, pParams("VATExclusive").Bool, pParams.ParameterExists("TransactionDistributionCode").Value, pParams.ParameterExists("SalesContactNumber").Value)
            Else
              vTDRLine.AddServiceBooking(pParams("Product").Value, pParams("Rate").Value, pParams("SBGrossQty").IntegerValue, pParams("Amount").DoubleValue, pParams("Source").Value, vVATRate.VatRateCode, vVATRate.CurrentPercentage(pParams("StartDate").Value), pParams("ServiceBookingNumber").IntegerValue, pParams("VATExclusive").Bool, pParams.ParameterExists("TransactionDistributionCode").Value, pParams.ParameterExists("SalesContactNumber").Value, pParams.ParameterExists("ContactDiscount").Bool, pParams.ParameterExists("GrossAmount").DoubleValue, pParams.ParameterExists("Discount").DoubleValue)
            End If
          End If

          If CSDepositPercentage > 0 Then pTransaction.TraderAnalysisLines.SetDepositAllowed(mvEnv)

          If vProductOffer.EntitlementProduct.Length > 0 AndAlso vProductOffer.EntitlementRate.Length > 0 AndAlso pParams("SBEntitlementQty").Value.Length > 0 Then
            vProduct = New Product(mvEnv)
            vProduct.InitWithRate(mvEnv, vProductOffer.EntitlementProduct, vProductOffer.EntitlementRate)

            'Get the current price for the entitlement product
            vCurrentPrice = vProduct.ProductRate.Price(vProduct.ProductRate.PriceChangeDate.AddDays(-1), 0)

            If vProduct.Existing Then
              vVATRate = mvEnv.VATRate(vProduct.ProductVatCategory, vContact.VATCategory)
              'Set Quantity and Amount
              If mvServiceBookingCredits Then
                vSBValue = Math.Abs(pParams("SBEntitlementQty").IntegerValue)
              Else
                vSBValue = pParams("SBEntitlementQty").IntegerValue
              End If
              vStartDate = pParams("StartDate").Value
              If vProduct.ProductRate.VatExclusive Then
                vAmount = FixTwoPlaces(Int(((vCurrentPrice * vSBValue) + ((vCurrentPrice * vSBValue) * (vVATRate.CurrentPercentage(vStartDate) / 100))) * 100) / 100)
              Else
                vAmount = FixTwoPlaces(vCurrentPrice * vSBValue)
              End If

              'Add a new line item for the entitlement
              vLineCount = pParams("TransactionLines").IntegerValue + 1
              vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
              vTDRLine.Init(vLineCount, vTransactionType)
              vTDRLine.AddServiceBookingEntitlementProduct(vProductOffer.EntitlementProduct, vProductOffer.EntitlementRate, pParams("SBEntitlementQty").IntegerValue, vAmount, pParams("TransactionSource").Value, vVATRate.VatRateCode, vVATRate.Percentage, vProduct.ProductRate.VatExclusive, pParams("TransactionDistributionCode").Value, pParams("SalesContactNumber").Value)
            End If
          End If

        Case TraderPage.TraderPageType.tpPostageAndPacking
          'Adding postage and packing 
          If pParams.ParameterExists("Product").Value.Length > 0 Then
            vLineCount = pParams("TransactionLines").IntegerValue + 1
            vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
            vTDRLine.Init(vLineCount, (pParams("TransactionType").Value))
            vTDRLine.AddPostageAndPacking(pParams("Product").Value, pParams("Rate").Value, pParams("Amount2").DoubleValue, pParams("TransactionSource").Value, pParams("VatRate").Value, pParams("VatPercentage").DoubleValue, pParams("PriceVATExclusive").Bool, pParams.ParameterExists("TransactionDistributionCode").Value)
          End If
        Case TraderPage.TraderPageType.tpGoneAway, TraderPage.TraderPageType.tpCancelPaymentPlan, _
             TraderPage.TraderPageType.tpGiveAsYouEarnEntry, TraderPage.TraderPageType.tpCancelGiftAidDeclaration
          'do nothing,  vAddIncentives = True
        Case Else
          vAddIncentives = False
      End Select

      If vAddIncentives Then vAddIncentives = Len(vSource) > 0
      If vAddIncentives And pParams.ParameterExists("CheckIncentives").Bool And SCGetNextPage(pCurrentPageType, pParams) = TraderPage.TraderPageType.tpTransactionAnalysisSummary AndAlso vTDRLine IsNot Nothing Then
        vStartLineNo = vTDRLine.LineNumber
        AddPaymentIncentives(pTransaction, vTDRLine, pParams, vSource, vStartLineNo)
      End If
    End Sub
    Private Sub AddPaymentIncentives(ByRef pTransaction As TraderTransaction, ByRef pTDRLine As TraderAnalysisLine, ByVal pParams As CDBParameters, ByVal pSource As String, ByVal pStartLineNo As Integer)
      Dim vDS As New VBDataSelection
      Dim vDT As CDBDataTable
      Dim vContact As Contact

      'Add Basic Incentives
      vDS.Init(mvEnv, DataSelection.DataSelectionTypes.dstIncentives)
      vDS.AddParameter("Source", CDBField.FieldTypes.cftCharacter, pSource)
      vDS.AddParameter("ReasonForDespatch", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPaymentReason))
      vContact = New Contact(mvEnv)
      vContact.Init((pParams("PayerContactNumber").IntegerValue))
      vDS.AddParameter("VatCategory", CDBField.FieldTypes.cftCharacter, (vContact.VATCategory))
      vDS.AddParameter("Basic", CDBField.FieldTypes.cftCharacter, "Y")
      vDT = vDS.DataTable

      AddPaymentIncentivesToLines(vDT, pTransaction, pTDRLine, pParams, pSource, pStartLineNo)

      'Add Optional Incentives
      If Not pTransaction.IncentivesTable Is Nothing Then
        vDT = pTransaction.IncentivesTable
        AddPaymentIncentivesToLines(vDT, pTransaction, pTDRLine, pParams, pSource, pStartLineNo)
        pTransaction.IncentivesTable = New CDBDataTable
      End If
    End Sub
    Private Sub AddPaymentIncentivesToLines(ByVal pIncentivesTable As CDBDataTable, ByRef pTransaction As TraderTransaction, ByRef pTDRLine As TraderAnalysisLine, ByVal pParams As CDBParameters, ByVal pSource As String, ByVal pStartLineNo As Integer)
      Dim vRow As CDBDataRow
      Dim vLineNo As Integer

      vLineNo = pTDRLine.LineNumber
      For Each vRow In pIncentivesTable.Rows
        vLineNo = vLineNo + 1
        pTDRLine = pTransaction.GetTraderAnalysisLine(vLineNo)
        pTDRLine.Init(vLineNo)
        pTDRLine.AddIncentive(vRow.Item("Product"), vRow.Item("Rate"), CInt(vRow.Item("Quantity")), pSource, pParams("TransactionDate").Value, vRow.Item("DespatchMethod"), pParams("PayerContactNumber").IntegerValue, pParams("PayerAddressNumber").IntegerValue, vRow.Item("VatRate"), CDbl(vRow.Item("Percentage")), pStartLineNo)
      Next vRow
    End Sub

    Private Function SCGetNextPage(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters) As TraderPage.TraderPageType
      Dim vNextPage As TraderPage.TraderPageType
      Dim vTransPayMethod As String
      Dim vPPPayMethod As String
      Dim vTransactionType As String

      'Parameters that mey be required
      'TransactionPaymentMethod,TransactionType,PPPaymentType

      vNextPage = pCurrentPageType 'Default to staying on the same page
      vTransactionType = pParams.ParameterExists("TransactionType").Value

      Dim vNumber As Integer
      Dim vSO As New StandingOrder
      Dim vMember As New Member
      Dim vDD As New DirectDebit
      Dim vCovenant As New Covenant
      Dim vCCCA As New CreditCardAuthority
      Dim vPaymentPlan As New PaymentPlan
      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpBankDetails 'Bank Details
          If pParams("TransactionPaymentMethod").Value = "CQIN" Then
            vNextPage = TraderPage.TraderPageType.tpCreditCustomer
          ElseIf PayMethodsAtEnd Then
            vNextPage = TraderPage.TraderPageType.tpNone 'Finished
          Else
            vNextPage = TraderPage.TraderPageType.tpTransactionDetails
          End If

        Case TraderPage.TraderPageType.tpBatchInvoiceProduction 'Batch Invoice Production
          If pParams.ParameterExists("DisplayInvoices").Bool Then
            vNextPage = TraderPage.TraderPageType.tpBatchInvoiceSummary
          Else
            vNextPage = TraderPage.TraderPageType.tpNone
          End If

        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          vNextPage = TraderPage.TraderPageType.tpNone

        Case TraderPage.TraderPageType.tpCollectionPayments
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpComments 'Comments
          vTransPayMethod = pParams("TransactionPaymentMethod").Value
          If PayMethodsAtEnd Then
            vNextPage = TraderPage.TraderPageType.tpPaymentMethod1
          ElseIf vTransPayMethod = "CARD" OrElse vTransPayMethod = "CAFC" OrElse vTransPayMethod = "CCIN" Then
            If vTransPayMethod = "CARD" AndAlso (String.Compare(mvEnv.GetConfig("fp_cc_authorisation_type"), "SAGEPAYHOSTED", True) = 0 AndAlso UseToken) Then
              vNextPage = TraderPage.TraderPageType.tpTokenSelection
            Else
              vNextPage = TraderPage.TraderPageType.tpCardDetails
            End If
          Else
            vNextPage = TraderPage.TraderPageType.tpNone 'Finished
          End If


        Case TraderPage.TraderPageType.tpTokenSelection
          vTransPayMethod = pParams("TransactionPaymentMethod").Value
          If PayMethodsAtEnd Then
            vNextPage = TraderPage.TraderPageType.tpPaymentMethod1
          ElseIf vTransPayMethod = "CARD" OrElse vTransPayMethod = "CAFC" OrElse vTransPayMethod = "CCIN" Then
            vNextPage = TraderPage.TraderPageType.tpTokenSelection
            vNextPage = TraderPage.TraderPageType.tpCardDetails
          Else
            vNextPage = TraderPage.TraderPageType.tpNone 'Finished
          End If

        Case TraderPage.TraderPageType.tpCreditCustomer 'Credit customer
          If PayMethodsAtEnd Then
            If pParams("TransactionPaymentMethod").Value = "CCIN" Then
              vNextPage = TraderPage.TraderPageType.tpCardDetails
            Else
              vNextPage = TraderPage.TraderPageType.tpNone 'Finished
            End If
          Else
            vNextPage = TraderPage.TraderPageType.tpTransactionDetails
          End If

        Case TraderPage.TraderPageType.tpEventBooking, TraderPage.TraderPageType.tpAmendEventBooking 'Event Booking
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
        Case TraderPage.TraderPageType.tpExamBooking
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
        Case TraderPage.TraderPageType.tpInvoicePayments 'Invoice Payments
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
        Case TraderPage.TraderPageType.tpPaymentMethod1 'Payment Method
          'THE FOLLOWING CODE COMES FROM DOCHOICE
          vTransPayMethod = pParams("TransactionPaymentMethod").Value
          Select Case vTransPayMethod
            Case "CRED", "CCIN"
              vNextPage = TraderPage.TraderPageType.tpCreditCustomer
            Case "CHEQ", "CQIN"
              If BankDetails Then
                vNextPage = TraderPage.TraderPageType.tpBankDetails
              ElseIf vTransPayMethod = "CQIN" Then
                vNextPage = TraderPage.TraderPageType.tpCreditCustomer
              ElseIf PayMethodsAtEnd Then
                vNextPage = TraderPage.TraderPageType.tpNone
              Else
                vNextPage = TraderPage.TraderPageType.tpTransactionDetails
              End If
              '      End If
            Case "CARD"
              If PayMethodsAtEnd Then
                vNextPage = TraderPage.TraderPageType.tpCardDetails
              Else
                vNextPage = TraderPage.TraderPageType.tpTransactionDetails
              End If
            Case Else
              If PayMethodsAtEnd Then
                vNextPage = TraderPage.TraderPageType.tpNone
              Else
                vNextPage = TraderPage.TraderPageType.tpTransactionDetails
              End If
          End Select

        Case TraderPage.TraderPageType.tpProductDetails 'Product details
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpTransactionAnalysis
          'THE FOLLOWING CODE COMES FROM DOCHOICE
          Select Case vTransactionType
            Case "MEMB", "SUBS", "DONR", "CMEM", "CSUB", "CDON", "LOAN"
              vNextPage = TraderPage.TraderPageType.tpPaymentMethod2
            Case "SALE", "DONS"
              vNextPage = TraderPage.TraderPageType.tpProductDetails
            Case "PAYM"
              vNextPage = TraderPage.TraderPageType.tpPayments
            Case "EVNT"
              vNextPage = TraderPage.TraderPageType.tpEventBooking
            Case "EXAM"
              vNextPage = TraderPage.TraderPageType.tpExamBooking
            Case "ACOM"
              vNextPage = TraderPage.TraderPageType.tpAccommodationBooking
            Case "COLP"
              vNextPage = TraderPage.TraderPageType.tpCollectionPayments
            Case "SRVC"
              vNextPage = TraderPage.TraderPageType.tpServiceBooking
            Case "INVC"
              vNextPage = TraderPage.TraderPageType.tpInvoicePayments
            Case "CRDN"
              vNextPage = TraderPage.TraderPageType.tpProductDetails
            Case "MEMC"
              vNextPage = TraderPage.TraderPageType.tpChangeMembershipType
            Case "STAT"
              vNextPage = TraderPage.TraderPageType.tpSetStatus
            Case "CANC"
              vNextPage = TraderPage.TraderPageType.tpCancelPaymentPlan
            Case "ACTV"
              vNextPage = TraderPage.TraderPageType.tpActivityEntry
            Case "GIFT"
              vNextPage = TraderPage.TraderPageType.tpGiftAidDeclaration
            Case "GAYE"
              vNextPage = TraderPage.TraderPageType.tpGiveAsYouEarnEntry
            Case "LEGR"
              vNextPage = TraderPage.TraderPageType.tpLegacyBequestReceipt
            Case "AWAY"
              vNextPage = TraderPage.TraderPageType.tpGoneAway
            Case "ADDR"
              vNextPage = TraderPage.TraderPageType.tpAddressMaintenance
            Case "SUPP"
              vNextPage = TraderPage.TraderPageType.tpSuppressionEntry
            Case "APAY"
              vNextPage = TraderPage.TraderPageType.tpContactSelection
            Case "CGAD"
              vNextPage = TraderPage.TraderPageType.tpCancelGiftAidDeclaration
            Case "CSRT"
              '      mvFireTransactionFinder = False
              vNextPage = TraderPage.TraderPageType.tpConfirmProvisionalTransactions
          End Select

        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary 'Transaction Analysis Summary
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis

        Case TraderPage.TraderPageType.tpTransactionDetails 'Transaction details
          'SetDummyContact
          If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Or pParams.ParameterExists("MemberNumber").IntegerValue > 0 Or pParams.ParameterExists("CovenantNumber").IntegerValue > 0 Then
            If Len(pParams.ParameterExists("MemberNumber").Value) > 0 And vTransactionType = "MEMC" Then
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis 'Could be CMT do not go straight for a payment
            Else
              vNextPage = TraderPage.TraderPageType.tpPayments
            End If
          Else
            If pParams.ParameterExists("TransactionLines").IntegerValue > 0 Then
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            Else
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis
            End If
          End If

        Case TraderPage.TraderPageType.tpMembership 'Membership
          'mvCardStatus = GetCurrentValue("opt_N")
          If pParams("NumberOfMembers").IntegerValue > 1 Or pParams("MaxFreeAssociates").IntegerValue > 0 Then
            vNextPage = TraderPage.TraderPageType.tpMembershipMembersSummary
          Else
            'lblCurrentMembers = ""
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetails
            If AppType = ApplicationType.atConversion Then
              'AddPayPlanDetailsToPPS      'Need to add the existing lines to the grid
            End If
          End If

        Case TraderPage.TraderPageType.tpContactSelection 'Contact selection
          Select Case AppType
            Case ApplicationType.atMaintenance
              If pParams.ParameterExists("LoanPaymentPlan").Bool Then
                vNextPage = TraderPage.TraderPageType.tpLoans
              Else
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              End If
            Case ApplicationType.atConversion
              vNextPage = TraderPage.TraderPageType.tpPaymentMethod3
            Case Else
              If vTransactionType = "APAY" Then
                If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
                  vNumber = pParams("PaymentPlanNumber").IntegerValue
                ElseIf pParams.ParameterExists("BankersOrderNumber").IntegerValue > 0 Then
                  vSO.Init(mvEnv, (pParams("BankersOrderNumber").IntegerValue))
                  vNumber = vSO.PaymentPlanNumber
                ElseIf pParams.ParameterExists("MemberNumber").IntegerValue > 0 Then
                  vMember.Init(mvEnv, (pParams("MemberNumber").IntegerValue))
                  vNumber = vMember.PaymentPlanNumber
                ElseIf pParams.ParameterExists("DirectDebitNumber").IntegerValue > 0 Then
                  vDD.Init(mvEnv, (pParams("DirectDebitNumber").IntegerValue))
                  vNumber = vDD.PaymentPlanNumber
                ElseIf pParams.ParameterExists("CovenantNumber").IntegerValue > 0 Then
                  vCovenant.Init(mvEnv, (pParams("CovenantNumber").IntegerValue))
                  vNumber = vCovenant.PaymentPlanNumber
                ElseIf pParams.ParameterExists("CreditCardAuthorityNumber").IntegerValue > 0 Then
                  vCCCA.Init(mvEnv, (pParams("CreditCardAuthorityNumber").IntegerValue))
                  vNumber = vCCCA.PaymentPlanNumber
                End If
                vPaymentPlan.Init(mvEnv, vNumber)
                If Not vPaymentPlan.Existing Then RaiseError(DataAccessErrors.daePaymentPlanNotFound)
                If vPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
                  vNextPage = TraderPage.TraderPageType.tpDirectDebit
                ElseIf vPaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes Then
                  vNextPage = TraderPage.TraderPageType.tpStandingOrder
                ElseIf vPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes Then
                  vNextPage = TraderPage.TraderPageType.tpCreditCardAuthority
                End If
                If Not pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then pParams.Item("PaymentPlanNumber").SetData("PaymentPlanNumber", vNumber.ToString, CDBField.FieldTypes.cftLong)
              Else
                vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis
              End If
          End Select

        Case TraderPage.TraderPageType.tpPaymentPlanDetails 'Payment Plan details
          Select Case pParams("TransactionType").Value
            Case "MEMB", "CMEM", "MEMC"
              vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary
            Case "SALE", "EVNT", "ACOM", "SRVC" 'called from TPP page
              vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary
            Case Else
              If pParams.ParameterExists("PPDLines").IntegerValue > 0 Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary
              Else
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanProducts
              End If
          End Select

        Case TraderPage.TraderPageType.tpMembershipMembersSummary 'Members summary
          If vTransactionType = "MEMC" Then
            Dim vAdvancedCMT As Boolean
            If mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlAdvancedCMT) Then
              Dim vPP As New PaymentPlan()
              vPP.Init(mvEnv, pParams("PaymentPlanNumber").IntegerValue)
              Dim vNewMembershipType As MembershipType = mvEnv.MembershipType(pParams("MembershipType").Value)
              If vPP.Existing = True AndAlso vNewMembershipType.Existing = True AndAlso vPP.CanUseAdvancedCMT(vNewMembershipType) = True Then vAdvancedCMT = True
            End If
            If vAdvancedCMT Then
              vNextPage = TraderPage.TraderPageType.tpAdvancedCMT
            Else
              vNextPage = TraderPage.TraderPageType.tpMembershipPayer 'CMT - Membership Payer details
            End If
          Else
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetails
          End If

        Case TraderPage.TraderPageType.tpPurchaseOrderDetails 'Purchase Order Details
          If pParams.ParameterExists("PPDLines").IntegerValue > 0 Then
            vNextPage = TraderPage.TraderPageType.tpPurchaseOrderSummary
          ElseIf pParams.ParameterExists("PurchaseOrderNumber").IntegerValue > 0 Then
            vNextPage = MainPage
          Else
            vNextPage = LinePage
          End If

        Case TraderPage.TraderPageType.tpPurchaseInvoiceDetails
          If pParams.ParameterExists("PPDLines").IntegerValue > 0 OrElse pParams("PurchaseOrderNumber").IntegerValue > 0 Then
            vNextPage = TraderPage.TraderPageType.tpPurchaseInvoiceSummary
          Else
            vNextPage = LinePage
          End If

        Case TraderPage.TraderPageType.tpPurchaseInvoiceProducts, TraderPage.TraderPageType.tpPurchaseOrderProducts 'Purchase Order/Invoice Products
          vNextPage = SummaryPage

        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary, TraderPage.TraderPageType.tpPurchaseOrderSummary 'Purchase Order/Invoice Summary
          vNextPage = LinePage

        Case TraderPage.TraderPageType.tpPurchaseOrderPayments
          vNextPage = MainPage

        Case TraderPage.TraderPageType.tpStatementList
          vNextPage = TraderPage.TraderPageType.tpBankDetails

        Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          Select Case pParams("UnbalancedTransactionChoice").Value
            Case "TRAN"
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            Case "PLAN"
              vNextPage = TraderPage.TraderPageType.tpPaymentMethod2
          End Select

          '--------------------------------------------------------------------------------------------------
          ' BELOW HERE NOT YET SUPPORTED
          '--------------------------------------------------------------------------------------------------

        Case TraderPage.TraderPageType.tpPaymentMethod2
          'THE FOLLOWING CODE COMES FROM DOCHOICE
          vTransactionType = pParams("TransactionType").Value
          'Get ready for a new payment plan to be set up
          'ClearPageDefaults tpPaymentPlanDetails
          'ClearPageDefaults tpPaymentPlanProducts
          'ClearGrid grdPPS, False, True
          'mvTraderApplication.PaymentPlanDetails.Clear
          'mvPPLine = 1
          'SelectGridRow grdPPS, mvPPLine
          Select Case vTransactionType
            Case "CSUB", "CDON"
              vNextPage = TraderPage.TraderPageType.tpCovenant
            Case "MEMB", "CMEM", "MEMC"
              vNextPage = TraderPage.TraderPageType.tpMembership
            Case "SUBS", "DONR"
              vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetails
            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetails
            Case "LOAN"
              vNextPage = TraderPage.TraderPageType.tpLoans
          End Select

        Case TraderPage.TraderPageType.tpPaymentMethod3
          'THE FOLLOWING CODE COMES FROM DOCHOICE
          vPPPayMethod = pParams("TransactionPaymentMethod").Value
          Select Case vPPPayMethod
            Case "COVT"
              If PayPlanConversionMaintenance Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpCovenant
              End If
            Case "CVDD"
              If PayPlanConversionMaintenance Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpCovenant
              End If
            Case "CVSO"
              If PayPlanConversionMaintenance Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpCovenant
              End If
            Case "CVCC"
              If PayPlanConversionMaintenance Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpCovenant
              End If
            Case "DIRD"
              If (ConversionShowPPD = True Or PayPlanConversionMaintenance = True) Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpDirectDebit
              End If
            Case "STDO"
              If (ConversionShowPPD = True Or PayPlanConversionMaintenance = True) Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpStandingOrder
              End If
            Case "CCCA"
              If (ConversionShowPPD = True Or PayPlanConversionMaintenance = True) Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              Else
                vNextPage = TraderPage.TraderPageType.tpCreditCardAuthority
              End If
            Case "MEMB"
              '      mvTransactionType = "MEMB"
              vNextPage = TraderPage.TraderPageType.tpMembership
            Case "MAINT"
              If pParams.ParameterExists("LoanPaymentPlan").Bool Then
                vNextPage = TraderPage.TraderPageType.tpLoans
              Else
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
              End If
          End Select

        Case TraderPage.TraderPageType.tpPayments 'Payments
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then
            vNextPage = TraderPage.TraderPageType.tpOutstandingScheduledPayments
          Else
            vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
          End If

        Case TraderPage.TraderPageType.tpAccommodationBooking 'Accommodation Booking
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpServiceBooking
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpPaymentPlanProducts 'Payment Plan products
          vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary

        Case TraderPage.TraderPageType.tpPaymentPlanMaintenance, TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance 'Payment Plan Maintenance, Payment Plan Details Maintenance
          vNextPage = SummaryPage

        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority 'Standing Order, Direct Debit, Continuous Credit Card Authority
          If vTransactionType = "APAY" Then
            vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
          End If

        Case TraderPage.TraderPageType.tpAddressMaintenance
          If pParams.ParameterExists("TraderAddressUpdated").Bool = True Then vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpChangeMembershipType 'Change Membership Type (CMT)
          vNextPage = TraderPage.TraderPageType.tpMembershipMembersSummary

        Case TraderPage.TraderPageType.tpAmendMembership 'Amend Membership Details
          vNextPage = TraderPage.TraderPageType.tpMembershipMembersSummary

        Case TraderPage.TraderPageType.tpMembershipPayer

        Case TraderPage.TraderPageType.tpCovenant 'Covenant
          vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetails
        Case TraderPage.TraderPageType.tpPaymentPlanSummary 'Payment Plan Summary
          If (AppType = ApplicationType.atMaintenance Or PayPlanConversionMaintenance) Then '"MAINT"
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
          Else
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanProducts
          End If

        Case TraderPage.TraderPageType.tpPostageAndPacking 'Postage and Packing - Carriage
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpGoneAway, TraderPage.TraderPageType.tpSetStatus, TraderPage.TraderPageType.tpCancelPaymentPlan, _
             TraderPage.TraderPageType.tpLegacyBequestReceipt, TraderPage.TraderPageType.tpGiftAidDeclaration, _
             TraderPage.TraderPageType.tpGiveAsYouEarnEntry, TraderPage.TraderPageType.tpCancelGiftAidDeclaration, _
             TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpActivityEntry, TraderPage.TraderPageType.tpSuppressionEntry
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpConfirmProvisionalTransactions
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpLoans
          vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary

        Case TraderPage.TraderPageType.tpAdvancedCMT
          vNextPage = TraderPage.TraderPageType.tpMembershipPayer
      End Select
      SCGetNextPage = vNextPage
    End Function

    Private Function SCGetPreviousPage(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters) As TraderPage.TraderPageType
      Dim vDT As CDBDataTable
      Dim vPreviousPage As TraderPage.TraderPageType

      vPreviousPage = pCurrentPageType 'Default to staying on the same page

      Select Case pCurrentPageType

        Case TraderPage.TraderPageType.tpBankDetails 'Bank Details
          vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod1

        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          vPreviousPage = TraderPage.TraderPageType.tpBatchInvoiceProduction

        Case TraderPage.TraderPageType.tpCardDetails
          If Me.PayMethodsAtEnd Then
            vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod1
          Else
            If TransactionComments Then
              If (String.Compare(mvEnv.GetConfig("fp_cc_authorisation_type"), "SAGEPAYHOSTED", True) = 0 AndAlso UseToken) Then
                vPreviousPage = TraderPage.TraderPageType.tpTokenSelection
              Else
                vPreviousPage = TraderPage.TraderPageType.tpComments
              End If
            Else
              vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            End If
          End If


        Case TraderPage.TraderPageType.tpTokenSelection
          If Me.PayMethodsAtEnd Then
            vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod1
          Else
            If TransactionComments Then
              vPreviousPage = TraderPage.TraderPageType.tpComments
            Else
              vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            End If
          End If

        Case TraderPage.TraderPageType.tpComments
          vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpCreditCustomer 'Credit customer
          If pParams("TransactionPaymentMethod").Value = "CQIN" AndAlso BankDetails Then
            vPreviousPage = TraderPage.TraderPageType.tpBankDetails
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod1
          End If

        Case TraderPage.TraderPageType.tpEventBooking, TraderPage.TraderPageType.tpAccommodationBooking, TraderPage.TraderPageType.tpServiceBooking, _
             TraderPage.TraderPageType.tpInvoicePayments, TraderPage.TraderPageType.tpChangeMembershipType, TraderPage.TraderPageType.tpSetStatus, _
             TraderPage.TraderPageType.tpCancelPaymentPlan, TraderPage.TraderPageType.tpLegacyBequestReceipt, _
             TraderPage.TraderPageType.tpActivityEntry, TraderPage.TraderPageType.tpGiftAidDeclaration, _
             TraderPage.TraderPageType.tpGiveAsYouEarnEntry, TraderPage.TraderPageType.tpSuppressionEntry, TraderPage.TraderPageType.tpGoneAway, _
             TraderPage.TraderPageType.tpCancelGiftAidDeclaration, TraderPage.TraderPageType.tpAddressMaintenance, _
             TraderPage.TraderPageType.tpConfirmProvisionalTransactions, TraderPage.TraderPageType.tpCollectionPayments, _
             TraderPage.TraderPageType.tpExamBooking
          vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis

        Case TraderPage.TraderPageType.tpPaymentMethod1
          If PayMethodsAtEnd Then
            If TransactionComments Then
              vPreviousPage = TraderPage.TraderPageType.tpComments
            Else
              vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            End If
          End If

        Case TraderPage.TraderPageType.tpProductDetails
          'ClearPageDefaults tpProductDetails
          vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis
          If pParams.ParameterExists("StockSale").Bool = True And pParams.ParameterExists("StockTransactionID").IntegerValue > 0 Then
            'Previous from tpProductDetails page so delete current StockMovement and update the stock levels
            vDT = SCSumStockMovements(pParams("StockTransactionID").IntegerValue, pParams.ParameterExists("ExistingTransaction").Bool Or mvExistingAdjustmentTran = True, True)
            SCRemoveStockMovements(vDT, 0, True, False)
          End If

        Case TraderPage.TraderPageType.tpTransactionAnalysis
          vPreviousPage = TraderPage.TraderPageType.tpTransactionDetails

        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary
          '    EditScheduledPaymentAnalysisLine IIf(mvEditMode = EDIT_NEW, True, False), False
          '    ClearPageDefaults tpOutstandingScheduledPayments
          vPreviousPage = TraderPage.TraderPageType.tpTransactionDetails

        Case TraderPage.TraderPageType.tpTransactionDetails 'Transaction details
          If pParams("TransactionPaymentMethod").Value = "CRED" OrElse pParams("TransactionPaymentMethod").Value = "CCIN" OrElse pParams("TransactionPaymentMethod").Value = "CQIN" Then
            vPreviousPage = TraderPage.TraderPageType.tpCreditCustomer
          ElseIf pParams("TransactionPaymentMethod").Value = "CHEQ" And BankDetails Then
            vPreviousPage = TraderPage.TraderPageType.tpBankDetails
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod1
          End If

        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority
          If Me.AppType = ApplicationType.atConversion Then 'Conversion
            If Me.PayPlanConversionMaintenance = True Then
              'Go back to maintain detail lines
              vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanSummary
            ElseIf Me.ConversionShowPPD = True Then
              'Go back to maintain payment plan header
              vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
            Else
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod3
            End If
          ElseIf pParams("TransactionType").Value = "APAY" Then
            vPreviousPage = TraderPage.TraderPageType.tpContactSelection
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanSummary
          End If

          '--------------------------------------------------------------------------------------------------
          ' BELOW HERE NOT YET SUPPORTED
          '--------------------------------------------------------------------------------------------------

        Case TraderPage.TraderPageType.tpMembershipMembersSummary
          If pParams("TransactionType").Value = "MEMC" Then 'CMT
            vPreviousPage = TraderPage.TraderPageType.tpChangeMembershipType
          Else
            vPreviousPage = TraderPage.TraderPageType.tpMembership
          End If

        Case TraderPage.TraderPageType.tpPaymentPlanSummary
          If pParams("TransactionType").Value = "LOAN" Then
            vPreviousPage = TraderPage.TraderPageType.tpLoans
          ElseIf Me.AppType = ApplicationType.atMaintenance Or (Me.PayPlanConversionMaintenance = True And pParams("TransactionPaymentMethod").Value <> "MEMB") Then '"MAINT"
            'Pay Plan Maintenance, or Pay Plan Conversion with Maintenance not adding a Membership
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanDetails
          End If
          '    CalcRenewalAmountFromPPS

        Case TraderPage.TraderPageType.tpPaymentPlanProducts
          If pParams("TransactionType").Value = "LOAN" Then
            vPreviousPage = TraderPage.TraderPageType.tpLoans
          ElseIf (Me.PayPlanConversionMaintenance = True And pParams("TransactionPaymentMethod").Value <> "MEMB") Then
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanMaintenance
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanDetails
          End If
          '    CalcRenewalAmountFromPPS

        Case TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
          vPreviousPage = Me.SummaryPage

        Case TraderPage.TraderPageType.tpPaymentPlanDetails
          Select Case pParams("TransactionType").Value
            Case "CSUB", "CDON"
              vPreviousPage = TraderPage.TraderPageType.tpCovenant
            Case "MEMB", "CMEM"
              If pParams("NumberOfMembers").IntegerValue > 1 Or pParams("MaxFreeAssociates").IntegerValue > 0 Then
                vPreviousPage = TraderPage.TraderPageType.tpMembershipMembersSummary
              Else
                vPreviousPage = TraderPage.TraderPageType.tpMembership
              End If
            Case "MEMC" 'CMT
              If pParams("NumberOfMembers").IntegerValue > 1 Or pParams("MaxFreeAssociates").IntegerValue > 0 Then
                vPreviousPage = TraderPage.TraderPageType.tpMembershipMembersSummary
              Else
                vPreviousPage = TraderPage.TraderPageType.tpMembership
              End If
            Case "SUBS", "DONR"
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod2
            Case "SALE", "EVNT", "ACOM", "SRVC", "COLP" 'called from TPP page
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod2
          End Select
        Case TraderPage.TraderPageType.tpMembership
          If pParams("TransactionType").Value = "CMEM" Then
            vPreviousPage = TraderPage.TraderPageType.tpCovenant
          ElseIf pParams("TransactionPaymentMethod").Value = "CAFC" Or pParams("TransactionPaymentMethod").Value = "VOUC" Then  'Since CAF transactions only support donations & memberships don't show the PM2 options
            vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis
          Else
            If Me.AppType = ApplicationType.atConversion Then
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod3
            Else
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod2
            End If
          End If

        Case TraderPage.TraderPageType.tpAmendMembership
          vPreviousPage = TraderPage.TraderPageType.tpMembershipMembersSummary

        Case TraderPage.TraderPageType.tpMembershipPayer
          Dim vAdvancedCMT As Boolean
          If mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlAdvancedCMT) Then
            Dim vPP As New PaymentPlan
            vPP.Init(mvEnv, pParams.ParameterExists("PaymentPlanNumber").IntegerValue)
            Dim vMT As MembershipType = mvEnv.MembershipType(pParams.ParameterExists("MembershipType").Value)
            vAdvancedCMT = vPP.CanUseAdvancedCMT(vMT)
          End If
          If vAdvancedCMT Then
            vPreviousPage = TraderPage.TraderPageType.tpAdvancedCMT
          Else
            vPreviousPage = TraderPage.TraderPageType.tpMembershipMembersSummary
          End If

        Case TraderPage.TraderPageType.tpPayments
          If pParams.ParameterExists("PaymentPlanCreated").Bool Then
            'This Payment Plan has just been created so can not change the analysis type
            vPreviousPage = TraderPage.TraderPageType.tpTransactionDetails
          Else
            vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis
          End If

        Case TraderPage.TraderPageType.tpContactSelection
          If pParams("TransactionType").Value = "APAY" Then vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis

        Case TraderPage.TraderPageType.tpCovenant
          If Me.AppType = ApplicationType.atConversion Then 'Conversion
            If Me.PayPlanConversionMaintenance Then
              vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanSummary
            Else
              vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod3
            End If
          Else
            vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod2
          End If

        Case TraderPage.TraderPageType.tpPaymentMethod2
          If pParams.ParameterExists("TPPDone").Bool Then
            vPreviousPage = TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          Else
            vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysis
          End If

        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary, TraderPage.TraderPageType.tpPurchaseInvoiceProducts
          vPreviousPage = TraderPage.TraderPageType.tpPurchaseInvoiceDetails

        Case TraderPage.TraderPageType.tpPurchaseOrderSummary, TraderPage.TraderPageType.tpPurchaseOrderProducts
          vPreviousPage = TraderPage.TraderPageType.tpPurchaseOrderDetails

        Case TraderPage.TraderPageType.tpPurchaseOrderPayments
          vPreviousPage = TraderPage.TraderPageType.tpPurchaseOrderSummary

        Case TraderPage.TraderPageType.tpPostageAndPacking
          '    ClearPageDefaults tpPostageAndPacking
          vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          vPreviousPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary

        Case TraderPage.TraderPageType.tpPaymentMethod3
          vPreviousPage = TraderPage.TraderPageType.tpContactSelection

        Case TraderPage.TraderPageType.tpPaymentPlanMaintenance
          vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod3

        Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vPreviousPage = TraderPage.TraderPageType.tpPayments

        Case TraderPage.TraderPageType.tpLoans
          vPreviousPage = TraderPage.TraderPageType.tpPaymentMethod2

        Case TraderPage.TraderPageType.tpAdvancedCMT
          vPreviousPage = TraderPage.TraderPageType.tpMembershipMembersSummary
      End Select
      SCGetPreviousPage = vPreviousPage
    End Function

    Private Function SCDoFinished(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pResults As CDBParameters) As TraderPage.TraderPageType
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vOPS As OrderPaymentSchedule
      Dim vTransStatus As SaveTransactionStatus
      Dim vTDRLine As TraderAnalysisLine
      Dim vUpdatedOPS As OrderPaymentSchedule
      Dim vPayPlansOnly As Boolean
      Dim vInvoiceAllocationsOnly As Boolean
      Dim vMaintenanceOnly As Boolean
      Dim vDisplayOPS As Boolean
      Dim vDone As Boolean
      Dim vExistingTrans As Boolean
      Dim vMaintenanceExists As Boolean
      Dim vMailingCode As String

      Dim vMailingTemplate As MailingTemplate = Nothing
      Dim vContact As Contact
      Dim vOrganisation As Organisation = Nothing
      Dim vSaveTransaction As Boolean
      Dim vTransactionPayMethod As String
      Dim vTransactionType As String
      Dim vFH As FinancialHistory
      Dim vSelectedTrans As CDBCollection = Nothing
      Dim vEB As EventBooking
      Dim vLinkedAnalysis As CollectionList(Of BatchTransactionAnalysis) = Nothing
      Dim vEBNumber As Integer
      Dim vAppType As ApplicationType
      Dim vCreateCMD As Boolean
      Dim vUpdateSource As Boolean
      Dim vEPMLines As Boolean
      Dim vNextPage As TraderPage.TraderPageType

      vNextPage = TraderPage.TraderPageType.tpNone 'Default
      vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)
      vExistingTrans = pParams.ParameterExists("ExistingTransaction").Bool
      vTransactionType = pParams.ParameterExists("TransactionType").Value
      vTransactionPayMethod = pParams("TransactionPaymentMethod").Value
      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          If ProduceInvoice(pParams, pResults) > 0 Then
            pResults.Add("Company", CDBField.FieldTypes.cftCharacter, pParams("Company").Value)
          Else
            RaiseError(DataAccessErrors.daeNoInvoicesMatchCriteria)
          End If

        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary
          If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or vFinancialAdjustment = Batch.AdjustmentTypes.atMove Then
          Else
            vPayPlansOnly = True
            vInvoiceAllocationsOnly = True
            vMaintenanceOnly = True
            For Each vTDRLine In pTransaction.TraderAnalysisLines
              Select Case vTDRLine.TraderLineType
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale, TraderAnalysisLine.TraderAnalysisLineTypes.taltDeceased, TraderAnalysisLine.TraderAnalysisLineTypes.taltSoftCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership, TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant, TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan, TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent, TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation, TraderAnalysisLine.TraderAnalysisLineTypes.taltIncentive, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingEntitlement, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltLegacyBequestReceipt, TraderAnalysisLine.TraderAnalysisLineTypes.taltExamBooking
                  '"P", "G", "S", "H", "M", "C", "O", "E", "A", "I", "V", "N", "U", "VE", "R", "VC", "B"  'P payment, G deceased, E event, A accommodation, M membership, C covenant, O order, SO, DD, CC, H hard credit, S soft credit, I incentive, V service, VC -ve service, B Legacy Receipt,Q Exam
                  vPayPlansOnly = False
                  vInvoiceAllocationsOnly = False
                  vMaintenanceOnly = False
                  Exit For
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
                  '"L", "K"  'L Invoice Allocation, K Sundry Credit Note Invoice Allocation
                  vPayPlansOnly = False
                  'vAllocationsExist = True 'BR16409: Changed as sundry credit note invoice allocations no longer created in non financial batch
                  vMaintenanceOnly = False
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAddSuppression, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltGoneAway, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltStatus, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltCancelPaymentPlan, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltActivityEntry, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltGiftAidDeclaration, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltPayrollGivingPledge, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltCancelGiftAidDeclaration, _
                     TraderAnalysisLine.TraderAnalysisLineTypes.taltAddressUpdate
                  '"AS", "GA", "ST", "CP", "AA", "GD", "GP", "CG", "ADDR" 'AS Add Suppression, GA Gone Away, ST Set Status, CP Cancel Payment Plan, AA Add Activity, GD Gift Aid Declaration, GP GAYE Pledge, CG Cancel Gift Aid Declaration, ADDR Address Updated
                  vPayPlansOnly = False
                  vInvoiceAllocationsOnly = False
                  vMaintenanceExists = True
              End Select
            Next vTDRLine
            If Not (vPayPlansOnly Or vInvoiceAllocationsOnly Or vMaintenanceOnly) Then
              If PayPlanPayMethod AndAlso pParams("DetailLineTotal").DoubleValue <> pParams("TransactionAmount").DoubleValue AndAlso Not pParams.ParameterExists("TPPDone").Bool Then
                vNextPage = TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
                vDone = True
              ElseIf Carriage AndAlso Not pParams.ParameterExists("PAPDone").Bool Then
                vNextPage = TraderPage.TraderPageType.tpPostageAndPacking
                vDone = True
              ElseIf TransactionComments Then
                If PayMethodsAtEnd AndAlso (pParams.ParameterExists("TransactionPaymentMethod").Value = "CQIN" OrElse pParams.ParameterExists("TransactionPaymentMethod").Value = "CCIN") Then
                  vNextPage = TraderPage.TraderPageType.tpCreditCustomer
                Else
                  vNextPage = TraderPage.TraderPageType.tpComments
                End If
                vDone = True
              ElseIf PayMethodsAtEnd Then
                If PayMethodsAtEnd AndAlso (pParams.ParameterExists("TransactionPaymentMethod").Value = "CQIN" OrElse pParams.ParameterExists("TransactionPaymentMethod").Value = "CCIN") Then
                  vNextPage = TraderPage.TraderPageType.tpCreditCustomer
                Else
                  vNextPage = TraderPage.TraderPageType.tpPaymentMethod1
                End If
                vDone = True
              ElseIf pParams("TransactionPaymentMethod").Value = "CARD" Or pParams("TransactionPaymentMethod").Value = "CAFC" Or pParams("TransactionPaymentMethod").Value = "CCIN" Then
                vNextPage = TraderPage.TraderPageType.tpCardDetails
                vDone = True
              End If
            End If
          End If

        Case TraderPage.TraderPageType.tpComments 'Comments
          If Me.PayMethodsAtEnd Then
            vNextPage = TraderPage.TraderPageType.tpPaymentMethod1
            vDone = True
          ElseIf pParams("TransactionPaymentMethod").Value = "CARD" OrElse pParams("TransactionPaymentMethod").Value = "CAFC" OrElse _
            pParams("TransactionPaymentMethod").Value = "CCIN" Then
            vNextPage = TraderPage.TraderPageType.tpCardDetails
            vDone = True
          End If

        Case TraderPage.TraderPageType.tpBatchInvoiceProduction 'Batch Invoice Generation
          pResults.Add("PrintInvoice", CDBField.FieldTypes.cftCharacter, "Y")
          If Not pParams.ParameterExists("DisplayInvoices").Bool Then
            vNextPage = TraderPage.TraderPageType.tpBatchInvoiceProduction 'Force Smart Client to remain on this page.
          Else
            vNextPage = TraderPage.TraderPageType.tpBatchInvoiceSummary
          End If
          vDone = True

        Case TraderPage.TraderPageType.tpTransactionDetails
          If pTransaction.TraderAnalysisLines.Count = 0 Then
            vNextPage = TraderPage.TraderPageType.tpTransactionAnalysis
            vDone = True
          End If
        Case TraderPage.TraderPageType.tpCreditCustomer 'Credit customer

        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority, TraderPage.TraderPageType.tpPaymentPlanMaintenance 'Standing Order, Direct Debit, Continuous Credit Card Authority, Pay Plan Amend
          If (AppType = ApplicationType.atConversion Or AppType = ApplicationType.atMaintenance) Or vTransactionType = "APAY" Then 'Maint
            If vTransactionType = "APAY" Then
              SaveAutoPaymentMethodChanges(pCurrentPageType, pParams, pTransaction)
              vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
              vDone = True
            Else
              If pCurrentPageType = TraderPage.TraderPageType.tpPaymentPlanMaintenance Then
                If pParams.Exists("Amount") Then
                  If pParams("Amount").DoubleValue > 0 And (pParams("PPBalance").DoubleValue <> pParams("PPDTotal").DoubleValue) Then 'Amount > 0 and totals have changed
                    vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary
                    vDone = True
                  End If
                End If
                If pParams.Exists("Balance") Then
                  Dim vPaymentPlan As PaymentPlan = pTransaction.PaymentPlan
                  vPaymentPlan.Init(mvEnv, pParams("PaymentPlanNumber").IntegerValue)
                  If vPaymentPlan.Balance <> pParams("PPBalance").DoubleValue Then ' User has changed the balance
                    If vPaymentPlan.Details.Count > 0 And pParams("PPDLines").IntegerValue = 0 Then 'User has got to Payment Plan Maintenance so no payment plan lines 
                      RaiseError(DataAccessErrors.daePaymentPlanBalanceChangedButNotDetails) 'Do not proceed with Finish, Details need to be changed. Informational rather than error.
                    End If
                  End If
                End If
              End If
              If vDone = False Then ProcessPaymentPlan(pParams, pTransaction, pResults, vNextPage, vUpdateSource)
            End If
          Else
            ProcessPaymentPlan(pParams, pTransaction, pResults, vNextPage)
            vDone = True
          End If

        Case TraderPage.TraderPageType.tpPaymentPlanSummary
          'If this is Conversion app and maintenance only option chosen then its really a maintenance app
          vAppType = Me.AppType
          If (Me.PayPlanConversionMaintenance = True And pParams.ParameterExists("TransactionPaymentMethod").Value = "MAINT") Then vAppType = ApplicationType.atMaintenance
          If SupportsNonFinancialBatch Then
            InitNonFinancialTransaction()
            pParams.Add("NonFinancialBatchNumber", NonFinancialBatchNumber)
            pParams.Add("NonFinancialTransactionNumber", NonFinancialTransactionNumber)
            mvNonFinancialBatch.UpdateNumberOfTransactions(1)
            If mvApplication = "3656" Then
              'For this trader app "DD's and SO's" the following parameters is needed to be passed to the next page via the smart client for the 
              ' creation of contact_journals records. This is done via pResults as this is returned to the client - not pParams  
              pResults.Add("NonFinancialBatchNumber", NonFinancialBatchNumber)
              pResults.Add("NonFinancialTransactionNumber", NonFinancialTransactionNumber)
            End If
          End If

          'Now go to the right place
          If vAppType <> ApplicationType.atMaintenance Then 'TA 25/4 - pages are not loaded
            Select Case pParams("PPPaymentType").Value
              Case "STDO"
                vNextPage = TraderPage.TraderPageType.tpStandingOrder
              Case "DIRD"
                vNextPage = TraderPage.TraderPageType.tpDirectDebit
              Case "CCCA"
                vNextPage = TraderPage.TraderPageType.tpCreditCardAuthority
              Case Else
                If (Me.ConversionShowPPD = False And Me.PayPlanConversionMaintenance = False) Or ((Me.ConversionShowPPD = True Or Me.PayPlanConversionMaintenance = True) And vTransactionType = "MEMB" And pParams("PPPaymentType").Value <> "NPAY") Then
                  'Conversion w/o PP details or Adding m/ship with PP details
                  ProcessPaymentPlan(pParams, pTransaction, pResults, vNextPage)
                  vDisplayOPS = True
                End If
            End Select
          Else
            ProcessPaymentPlan(pParams, pTransaction, pResults, vNextPage)
            vDisplayOPS = True
          End If

          If vAppType <> ApplicationType.atMaintenance Then vDone = True '"MAINT"

        Case TraderPage.TraderPageType.tpScheduledPayments
          'User may have amended the OPS so save any changes
          If pTransaction.UpdatedOPS.Count() > 0 Then
            vUpdatedOPS = CType(pTransaction.UpdatedOPS.Item(1), OrderPaymentSchedule)
            If vUpdatedOPS.PlanNumber > 0 Then pTransaction.PaymentPlan.Init(mvEnv, (vUpdatedOPS.PlanNumber))
            For Each vUpdatedOPS In pTransaction.UpdatedOPS
              For Each vOPS In pTransaction.PaymentPlan.ScheduledPayments
                If vOPS.ScheduledPaymentNumber = vUpdatedOPS.ScheduledPaymentNumber Then
                  With vOPS
                    .Update(.DueDate, .AmountDue, vUpdatedOPS.AmountOutstanding, vUpdatedOPS.ExpectedBalance, .ClaimDate, vUpdatedOPS.RevisedAmount)
                    .Save()
                  End With
                  Exit For
                End If
              Next vOPS
            Next vUpdatedOPS
          End If
          'We do not need to set vDone = True as it will stop creating CMD, if required, for new DD Payment Plans.

        Case TraderPage.TraderPageType.tpMembershipPayer
          If Not pParams.Exists("TRD_BatchCategory") AndAlso Not String.IsNullOrEmpty(BatchCategory) Then pParams.Add("TRD_BatchCategory", BatchCategory)
          ProcessPaymentPlan(pParams, pTransaction, pResults, TraderPage.TraderPageType.tpNone)

        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary, TraderPage.TraderPageType.tpPurchaseOrderSummary
          If pCurrentPageType = TraderPage.TraderPageType.tpPurchaseOrderSummary And pParams.ParameterExists("NumberOfPayments").IntegerValue > 0 Then
            vDone = True
            vNextPage = TraderPage.TraderPageType.tpPurchaseOrderPayments
          Else
            SCSavePurchaseOI(pCurrentPageType, pParams, pTransaction, pResults)
            vDone = True
          End If

        Case TraderPage.TraderPageType.tpPurchaseOrderPayments
          SCSavePurchaseOI(pCurrentPageType, pParams, pTransaction, pResults)
          vDone = True

        Case TraderPage.TraderPageType.tpPurchaseOrderCancellation 'Purchase order Cancellation
          CancelPOs(pParams)
          vDone = True

        Case TraderPage.TraderPageType.tpChequeNumberAllocation
          ChequeNoAlloc(pParams, pResults)
          vDone = True

        Case TraderPage.TraderPageType.tpChequeReconciliation
          ChequeReconcile(pParams)
          vDone = True

        Case TraderPage.TraderPageType.tpActivityEntry, TraderPage.TraderPageType.tpSuppressionEntry, TraderPage.TraderPageType.tpSetStatus, _
             TraderPage.TraderPageType.tpGiftAidDeclaration, TraderPage.TraderPageType.tpGiveAsYouEarnEntry, TraderPage.TraderPageType.tpGoneAway, _
              TraderPage.TraderPageType.tpCancelGiftAidDeclaration, TraderPage.TraderPageType.tpCancelPaymentPlan
          SCSaveAnalysis(pCurrentPageType, pParams, pTransaction)
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
          vDone = True

        Case TraderPage.TraderPageType.tpPostTaxPGPayment
          CreatePostTaxPGPayment(pParams, pTransaction)
          pResults.Add("BatchNumber", pTransaction.BatchNumber)
          pResults.Add("TransactionNumber", pTransaction.TransactionNumber)
          vDone = True

        Case TraderPage.TraderPageType.tpGiveAsYouEarn
          CreatePreTaxPGPayment(pParams, pTransaction)
          pResults.Add("BatchNumber", pTransaction.BatchNumber)
          pResults.Add("TransactionNumber", pTransaction.TransactionNumber)
          vDone = True

      End Select

      If pParams.ParameterExists("AddDonationToTrans").Value.ToString.Length > 0 Then
        If (pCurrentPageType = TraderPage.TraderPageType.tpTransactionAnalysisSummary AndAlso Not TransactionComments) OrElse _
          (pCurrentPageType = TraderPage.TraderPageType.tpComments AndAlso TransactionComments) Then
          'add the donation to make up the transaction price if the user has requested
          Dim vLine As TraderAnalysisLine = pTransaction.GetTraderAnalysisLine(pTransaction.TraderAnalysisLines.Count + 1, pParams("TransactionType").Value)
          Dim vProduct As New Product(mvEnv)
          vProduct.InitWithRate(mvEnv, DonationProduct, DonationRate)
          Dim vPriceVATExclusive As Boolean = vProduct.ProductRate.VatExclusive
          Dim vPayerContact As New Contact(mvEnv)
          vPayerContact.Init(pParams("PayerContactNumber").IntegerValue)
          Dim vVatRate As VatRate = mvEnv.VATRate(vProduct.ProductVatCategory, vPayerContact.VATCategory)
          Dim vVatPercent As Double = vVatRate.Percentage
          Dim vVatAmount As Double = CalculateVATAmount(pParams("AddDonationToTrans").DoubleValue, vVatPercent)
          vLine.AddNonStockProductSale(DonationProduct, DonationRate, 1, DoubleValue(pParams("AddDonationToTrans").Value), pParams("TransactionDate").Value, pParams("TransactionSource").Value, pParams("PayerContactNumber").IntegerValue, pParams("PayerAddressNumber").IntegerValue, vVatRate.VatRateCode, vVatPercent, vVatAmount, vPriceVATExclusive)
          pParams("TRD_Amount").Value = (pParams("TRD_Amount").DoubleValue + pParams("AddDonationToTrans").DoubleValue).ToString
        End If
      End If

      If vDisplayOPS = True And vDone = False Then
        'Display the Scheduled Payments page if required
        If Me.DisplayScheduledPayments = True And pTransaction.PaymentPlan.StandingOrderStatus <> PaymentPlan.ppYesNoCancel.ppYes Then
          vDone = True
          If pTransaction.PaymentPlan.ScheduledPayments.Count() > 1 Then
            'Only show this page if there is more than 1 ops record to display
            vNextPage = TraderPage.TraderPageType.tpScheduledPayments
          End If
        End If
      End If

      If vDone = False Then
        'No more pages to go to, now need to save transaction etc.
        If vTransactionPayMethod = "CRED" OrElse vTransactionPayMethod = "CQIN" OrElse vTransactionPayMethod = "CCIN" Then
          'Set up CreditCustomers for CreditSales
          With pParams
            pTransaction.SetCreditCustomer(.Item("CCU_ContactNumber").IntegerValue, CSCompany, .Item("CCU_SalesLedgerAccount").Value, .Item("CCU_AddressNumber").IntegerValue, .Item("CCU_TermsNumber").Value, .Item("CCU_TermsPeriod").Value, .Item("CCU_TermsFrom").Value, .Item("CCU_CreditCategory").Value, .Item("CCU_CreditLimit").DoubleValue, .Item("CCU_CustomerType").Value, (.ParameterExists("CCU_StopCode").Value), False, Nothing, CStr(CSTermsNumber), CSTermsPeriod, CSTermsFrom)
          End With

          If vFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment And pTransaction.CreditCustomer.Existing = False Then
            RaiseError(DataAccessErrors.daeCreditCustomerMissing, (pParams.Item("CCU_ContactNumber").Value), CSCompany, (pParams.Item("CCU_SalesLedgerAccount").Value))
          End If
        End If

        If vFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment Then
          'This will be creating a new transaction including Event Booking.  If existing booking was not changed then we need to create it
          If pParams.Exists("EventNumber") = False Then
            EditEventBooking(pParams, pTransaction)
          End If
        End If

        If pTransaction.TraderAnalysisLines.Count > 0 Then
          For Each vTDRLine In pTransaction.TraderAnalysisLines
            If vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliCreatesBTA) Then
              vSaveTransaction = True
            ElseIf vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliIsMaintenanceType) Then
              vMaintenanceExists = True
            End If
          Next vTDRLine
          If vSaveTransaction Then vTransStatus = SaveTransaction(pTransaction, pParams, vFinancialAdjustment, vExistingTrans, "", True)
        End If

        '*** CONTACT MAILING DOCUMENT ***
        If AppType = ApplicationType.atConversion Then
          vCreateCMD = True
        ElseIf vFinancialAdjustment = Batch.AdjustmentTypes.atNone And pParams.OptionalValue("PayerContactNumber", "") <> (mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber)) Then
          If pTransaction.BatchNumber > 0 Or MailingCodeMandatory Then
            vCreateCMD = True
          ElseIf pParams.ParameterExists("CreateMailingDocument").Bool Then
            pResults.Add("CreateMailingDocument", CDBField.FieldTypes.cftCharacter, "Y")
          ElseIf pParams.ParameterExists("PaymentPlanCreated").Bool Then
            vCreateCMD = True
          End If
        End If
        If vCreateCMD Then
          vMailingCode = pParams.ParameterExists("TRD_Mailing").Value
          If Len(vMailingCode) = 0 Then vMailingCode = pParams.ParameterExists("CSE_Mailing").Value
          'Other Mailing / Source code checks are done client-side
          vMailingTemplate = New MailingTemplate(mvEnv)
          vMailingTemplate.InitFromMailing(mvEnv, vMailingCode)
          If vMailingTemplate.Existing Then
            vContact = New Contact(mvEnv)
            vContact.Init(If(pParams.ParameterExists("TRD_MailingContactNumber").IntegerValue > 0, pParams("TRD_MailingContactNumber").IntegerValue, pParams.Item("PayerContactNumber").IntegerValue))
            If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
              vOrganisation = New Organisation(mvEnv)
              vOrganisation.InitNumberOnly((vContact.ContactNumber))
            End If
            If Not vMailingTemplate.ContactHasExclusionSuppressions(vContact, vOrganisation) Then
              vTransStatus = vTransStatus Or SaveTransactionStatus.stsCreateMailingDocument
              If Len(vMailingTemplate.ContactWarningSuppressions(vContact, vOrganisation)) > 0 Then vTransStatus = vTransStatus Or SaveTransactionStatus.stsContactWarningSuppressionsPrompt
            End If
          ElseIf vExistingTrans Then
            'Editing an existing transaction but the new mailing code isn't linked to a mailing template.
            'Need to delete any CMD previously created for this transaction.
            vMailingTemplate.ContactMailingDocument.InitFromTransaction(pTransaction.BatchNumber, pTransaction.TransactionNumber)
            vMailingTemplate.ContactMailingDocument.Delete()
          End If
        End If
      End If

      If vSaveTransaction = True And pTransaction.BatchNumber > 0 And pTransaction.TransactionNumber > 0 Then
        If mvExistingAdjustmentTran = True And vFinancialAdjustment <> Batch.AdjustmentTypes.atGIKConfirmation And vFinancialAdjustment <> Batch.AdjustmentTypes.atCashBatchConfirmation Then
          vFH = New FinancialHistory
          vFH.Init(mvEnv, BatchNumber, TransactionNumber)
          If pParams.Exists("BatchNumbers") Then vSelectedTrans = vFH.GetMultipleTransactions(pParams("BatchNumbers").Value)
          vLinkedAnalysis = Nothing
          If vExistingTrans = False And EventMultipleAnalysis = True And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) = True Then vLinkedAnalysis = New CollectionList(Of BatchTransactionAnalysis)
          vFH.ReanalyseTransaction(pTransaction.BatchNumber, pTransaction.TransactionNumber, pParams, Batch, vSelectedTrans, vLinkedAnalysis)
        ElseIf mvExistingAdjustmentTran = False And vFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment Then
          'We have vcreated a new Transaction for the new details, now need to reverse the existing Transaction
          pParams.Add("OriginalBatchNumber", pParams("BatchNumber").IntegerValue)
          pParams.Add("OriginalTransactionNumber", pParams("TransactionNumber").IntegerValue)
          pParams("BatchNumber").Value = pTransaction.BatchNumber.ToString
          pParams("TransactionNumber").Value = pTransaction.TransactionNumber.ToString
        End If
        'BR12729: Link multiple lines to single event booking
        If vExistingTrans = False And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) = True Then
          If EventMultipleAnalysis = True Then
            vEPMLines = False
            For Each vTDRLine In pTransaction.TraderAnalysisLines
              If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent Then
                If vEBNumber = 0 Then
                  vEBNumber = vTDRLine.EventBookingNumber
                Else
                  If (vFinancialAdjustment = Batch.AdjustmentTypes.atNone Or vFinancialAdjustment = Batch.AdjustmentTypes.atEventAdjustment) Then
                    vEB = New EventBooking
                    vEB.Init(mvEnv, 0, vEBNumber)
                    vEB.AddLinkedTransaction(vFinancialAdjustment, vLinkedAnalysis, Nothing, vTDRLine.LineNumber - 1, vEPMLines)
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
              vEB.AddLinkedTransaction(vFinancialAdjustment, vLinkedAnalysis, Nothing, (pTransaction.TraderAnalysisLines(pTransaction.TraderAnalysisLines.Count).LineNumber), vEPMLines)
            End If
          Else
            'Handle Event Pricing Matrix lines that must be linked to the EventBooking irrespective of the EventMultipleAnalysis flag on the Trader Application
            vEPMLines = False
            For Each vTDRLine In pTransaction.TraderAnalysisLines
              If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent Then
                If vEBNumber = 0 Then
                  vEBNumber = vTDRLine.EventBookingNumber
                ElseIf vFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
                  If vEPMLines = True Then
                    vEB = New EventBooking
                    vEB.Init(mvEnv, 0, vEBNumber)
                    vEB.AddLinkedTransactionForEPM(vTDRLine.LineNumber - 1)
                  End If
                  vEBNumber = vTDRLine.EventBookingNumber
                End If
                vEPMLines = False
              ElseIf vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEventPricingMatrixLine Then
                vEPMLines = True
              End If
            Next vTDRLine
            If vEBNumber > 0 And vEPMLines = True Then
              vEB = New EventBooking
              vEB.Init(mvEnv, 0, vEBNumber)
              vEB.AddLinkedTransactionForEPM(pTransaction.TraderAnalysisLines(pTransaction.TraderAnalysisLines.Count).LineNumber)
            End If
          End If
        End If

        'Handle the linking of exam booking analysis lines
        For Each vTDRLine In pTransaction.TraderAnalysisLines
          If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltExamBooking Then
            If vEBNumber = 0 Then
              vEBNumber = vTDRLine.EventBookingNumber
            ElseIf vFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
              If vEPMLines = True Then
                vEB = New EventBooking
                vEB.Init(mvEnv, 0, vEBNumber)
                vEB.AddLinkedTransactionForEPM(vTDRLine.LineNumber - 1)
              End If
              vEBNumber = vTDRLine.EventBookingNumber
            End If
            vEPMLines = False
          ElseIf vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltEventPricingMatrixLine Then
            vEPMLines = True
          End If
        Next vTDRLine




      End If

      If vTransStatus <> SaveTransactionStatus.stsNone Then
        If pTransaction.BatchNumber > 0 Or pTransaction.TransactionNumber > 0 Then
          pResults.Add("BatchNumber", pTransaction.BatchNumber)
          pResults.Add("TransactionNumber", pTransaction.TransactionNumber)
          If vFinancialAdjustment = Access.Batch.AdjustmentTypes.atAdjustment AndAlso vNextPage = TraderPage.TraderPageType.tpNone Then
            Dim vCountWhereFields As New CDBFields
            vCountWhereFields.Add("gad.batch_number", CDBField.FieldTypes.cftInteger, pParams("BatchNumber").IntegerValue)
            vCountWhereFields.Add("gad.transaction_number", CDBField.FieldTypes.cftInteger, pParams("TransactionNumber").IntegerValue)
            Dim vCountSQL As New SQLStatement(mvEnv.Connection, "", "gift_aid_declarations gad", vCountWhereFields, "", Nothing)
            Dim vGADCount = mvEnv.Connection.GetCountFromStatement(vCountSQL)
            If vGADCount > 0 Then
              pResults.Add("OneOffGADMessage", CDBField.FieldTypes.cftCharacter, "Y")
            End If
          End If
        End If
        If (vTransStatus And SaveTransactionStatus.stsPrintInvoice) = SaveTransactionStatus.stsPrintInvoice Then
          pResults.Add("PrintInvoice", CDBField.FieldTypes.cftCharacter, "Y")
        End If
        If (vTransStatus And SaveTransactionStatus.stsPrintReceipt) = SaveTransactionStatus.stsPrintReceipt Then
          pResults.Add("PrintReceipt", CDBField.FieldTypes.cftCharacter, "Y")
        End If
        If (vTransStatus And SaveTransactionStatus.stsPrintProvisionalCashDoc) = SaveTransactionStatus.stsPrintProvisionalCashDoc Then
          pResults.Add("PrintProvisionalCashDoc", CDBField.FieldTypes.cftCharacter, "Y")
        End If
        If (vTransStatus And SaveTransactionStatus.stsCreateMailingDocument) = SaveTransactionStatus.stsCreateMailingDocument Then
          pResults.Add("CreateMailingDocument", CDBField.FieldTypes.cftCharacter, "Y")
        End If
        If (vTransStatus And SaveTransactionStatus.stsContactWarningSuppressionsPrompt) = SaveTransactionStatus.stsContactWarningSuppressionsPrompt Then
          pResults.Add("ContactWarningSuppressionsPrompt", CDBField.FieldTypes.cftCharacter, "Y")
          pResults.Add("WarningSuppressions", CDBField.FieldTypes.cftCharacter, vMailingTemplate.WarningSuppressions)
        End If
      End If

      If vSaveTransaction AndAlso Me.AutoGiftAidDeclaration Then
        pTransaction.AutoAddGiftAidDeclaration(Me.AutoGiftAidSource, Me.AutoGiftAidMethod)
      End If

      '--------------------------------------------------------------------------------------------------
      ' BELOW HERE NOT YET SUPPORTED
      '--------------------------------------------------------------------------------------------------
      '
      If Not vDone Then
        If vMaintenanceExists Then SaveNonFinancial(pTransaction)
      End If
      SCDoFinished = vNextPage
    End Function

    Private Function SCDoCancelled(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pResults As CDBParameters) As TraderPage.TraderPageType
      Dim vDT As CDBDataTable
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vOPSFound As Boolean
      Dim vUpdate As Boolean
      Dim vTDRLine As TraderAnalysisLine
      Dim vIndex As Integer
      Dim vErrorMsg As String = ""
      Dim vExistingTrans As Boolean
      Dim vEventBooking As New EventBooking
      Dim vPaymentPlan As New PaymentPlan
      Dim vLineNumber As Integer
      Dim vPPNos() As String
      Dim vDeletedPaymentPlan As Boolean

      Dim vFinacialAdjustment As Batch.AdjustmentTypes = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Batch.AdjustmentTypes)
      Dim vBT As BatchTransaction = Nothing

      vExistingTrans = pParams.ParameterExists("ExistingTransaction").Bool
      If vExistingTrans = False Then
        'For Each vTDRLine In mvTDRTransaction.TraderAnalysisLines
        vPaymentPlan.Init(mvEnv)
        'SMART CLIENT: Set the collection for deleted or edited analysis lines
        'Currently supports atAdjustment only
        Dim vSetRemovedSchPayments As Boolean = pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment And pParams.ParameterExists("BatchNumber").IntegerValue > 0 And pParams.ParameterExists("TransactionNumber").IntegerValue > 0
        For Each vTDRLine In pTransaction.TraderAnalysisLines
          Select Case vTDRLine.TraderLineType
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
              'BR13480 Delete event booking but only if it does not have the batch number and transaction number set
              If vTDRLine.EventBookingNumber > 0 Then DeleteEventBooking(vTDRLine.EventBookingNumber, True, 0, vErrorMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltExamBooking
              DeleteExamBooking(vTDRLine.ExamBookingNumber, True, vErrorMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation
              'BR13480 Delete room booking but only if it does not have the batch number and transaction number set
              DeleteRoomBooking(vTDRLine.RoomBookingNumber, True, 0, vErrorMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit
              'BR13480 Delete service booking but only if it does not have the batch number and transaction number set
              DeleteServiceBooking(vTDRLine.ServiceBookingNumber, True, vErrorMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation, _
              TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment  'L, K, U, N
              If vTDRLine.Amount <= 0 AndAlso (vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation _
              OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation) Then
                'Only adjust the invoices for the positive line (as the negative line reverses these updates!)
                Exit Select
              End If
              If vFinacialAdjustment = Access.Batch.AdjustmentTypes.atNone OrElse (vFinacialAdjustment = Access.Batch.AdjustmentTypes.atAdjustment AndAlso IsOriginalAnalsysisLine(vBT, vTDRLine, pParams) = False) Then
                DeleteInvoicePayment(vTDRLine, vErrorMsg, vExistingTrans, pParams.ParameterExists("BatchNumber").LongValue, pParams.ParameterExists("TransactionNumber").LongValue)
              End If
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale
              If Not (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atGIKConfirmation Or pParams.ParameterExists("TransactionType").Value = "CSRT") Then
                'Do not need to re-allocate any product numbers
                If vTDRLine.ProductNumber > 0 Then ReAllocateProductNumber(vTDRLine.ProductCode, vTDRLine.ProductNumber)
              End If
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan, TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership, TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant, TraderAnalysisLine.TraderAnalysisLineTypes.taltHardCredit
              vOPSFound = False

              If vTDRLine.ScheduledPaymentNumber > 0 AndAlso vTDRLine.PaymentPlanNumber > 0 Then
                If vPaymentPlan.PlanNumber <> vTDRLine.ScheduledPaymentNumber Then vPaymentPlan.Init(mvEnv, (vTDRLine.PaymentPlanNumber))
                For Each vOPS In vPaymentPlan.ScheduledPayments
                  If vOPS.ScheduledPaymentNumber = vTDRLine.ScheduledPaymentNumber Then vOPSFound = True ' vSchNumber Then vOPSFound = True
                  If vOPSFound Then Exit For
                Next vOPS
                If Not vOPSFound Then
                  vOPS = New OrderPaymentSchedule
                  vOPS.Init(mvEnv, (vTDRLine.ScheduledPaymentNumber))
                  If vOPS.Existing Then vOPSFound = True
                End If
                If vOPSFound Then
                  If (pParams.ParameterExists("FinancialAdjustment").IntegerValue <> Batch.AdjustmentTypes.atAdjustment) Or (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment And pTransaction.RemovedSchPayments.Exists(CStr(vOPS.ScheduledPaymentNumber)) = False) Then
                    If (pCurrentPageType = TraderPage.TraderPageType.tpTransactionAnalysisSummary Or pCurrentPageType = TraderPage.TraderPageType.tpComments Or pCurrentPageType = TraderPage.TraderPageType.tpPostageAndPacking) Then
                      vUpdate = True 'Always update
                    ElseIf pCurrentPageType = TraderPage.TraderPageType.tpTransactionDetails Then  'And (grdTAS.SelStartRow = vTDRLine.LineNumber) Then
                      'Editing current line, so no update required
                      vUpdate = False
                    ElseIf pParams.ParameterExists("FinancialAdjustment").IntegerValue <> Batch.AdjustmentTypes.atAdjustment Then
                      'Only update if the ops is not in the collection
                      vUpdate = True ' Not mvRemovedSchPayments.Exists(Format$(vOPS.ScheduledPaymentNumber))
                    End If
                    vUpdate = vUpdate AndAlso Not (vSetRemovedSchPayments AndAlso pTransaction.RemovedSchPayments.Count = 0) 'An alternative of mvDataChange
                    If vUpdate Then
                      vOPS.SetUnProcessedPayment(False, vTDRLine.Amount * -1, (vPaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan))
                      vOPS.Save()
                    End If
                    If pTransaction.RemovedSchPayments.Count > 0 Then
                      'Remove from collection - could have been added by deleting analysis line
                      If pTransaction.RemovedSchPayments.Exists(CStr(vOPS.ScheduledPaymentNumber)) Then pTransaction.RemovedSchPayments.Remove(CStr(vOPS.ScheduledPaymentNumber))
                    End If
                  End If
                End If
              End If
            Case Else
              '
          End Select
        Next vTDRLine
        If mvEnv.GetConfigOption("fp_pp_delete_no_payment", False) Then
          If Len(pParams.ParameterExists("PaymentPlansToDelete").Value) > 0 Then
            vPPNos = Split(pParams("PaymentPlansToDelete").Value, ",")
            For vIndex = 0 To UBound(vPPNos)
              vDeletedPaymentPlan = False
              For Each vTDRLine In pTransaction.TraderAnalysisLines
                If vTDRLine.PaymentPlanNumber > 0 Then
                  If vTDRLine.PaymentPlanNumber = Val(vPPNos(vIndex)) Then
                    If Len(vTDRLine.MemberNumber) > 0 Then
                      DeletePaymentPlan(vTDRLine.PaymentPlanNumber, vErrorMsg, IntegerValue(vTDRLine.MemberNumber))
                      vDeletedPaymentPlan = True
                    Else
                      DeletePaymentPlan(vTDRLine.PaymentPlanNumber, vErrorMsg)
                      vDeletedPaymentPlan = True
                    End If
                  End If
                End If
              Next
              If vDeletedPaymentPlan = False Then
                DeletePaymentPlan(IntegerValue(vPPNos(vIndex)), vErrorMsg)
              End If
            Next
          End If
        End If

        If pCurrentPageType = TraderPage.TraderPageType.tpProductDetails Then
          If Not (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atGIKConfirmation Or pParams.ParameterExists("TransactionType").Value = "CSRT") Then
            If Len(pParams.ParameterExists("Product").Value) > 0 And pParams.ParameterExists("ProductNumber").IntegerValue > 0 Then
              ReAllocateProductNumber(pParams("Product").Value, pParams("ProductNumber").IntegerValue)
            End If
          End If
        End If

        If pCurrentPageType = TraderPage.TraderPageType.tpGiftAidDeclaration OrElse pCurrentPageType = TraderPage.TraderPageType.tpAddressMaintenance Then
          If pParams.Exists("NonFinancialBatchNumber") Then AdjustNumberOfTransactions(pParams("NonFinancialBatchNumber").IntegerValue, pParams("NonFinancialTransactionNumber").IntegerValue)
        End If
      End If

      'Put back any stock that has been issued
      vLineNumber = 0
      For Each vTDRLine In pTransaction.TraderAnalysisLines
        vLineNumber = vLineNumber + 1
        If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale Then
          If vTDRLine.StockSale Then
            'Delete StockMovements, so select just those created in this instance (BatchNumber will be null)
            'than as this is cancelling Trader, just update the stock levels
            vDT = SCSumStockMovements(vTDRLine.StockTransactionID, vExistingTrans, True)
            SCRemoveStockMovements(vDT, vLineNumber, True, False)
          End If
        End If
      Next vTDRLine

      If pParams.ParameterExists("StockSale").Bool = True And pParams.ParameterExists("StockTransactionID").IntegerValue > 0 Then
        vDT = SCSumStockMovements(pParams("StockTransactionID").IntegerValue, False, True)
        SCRemoveStockMovements(vDT, 0, True, False)
      End If

      If pTransaction.RemovedSchPayments.Count > 0 Then
        For Each vOPS In pTransaction.RemovedSchPayments
          vOPS.InitFromOPS(mvEnv, vOPS) 'This is to set the SetValue property of each class field and mvExisting = True as it is not being set in XMLTrader.GetParameterList
          If pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment Then
            vOPS.ProcessReanalysis(vOPS.PaymentAmount)
          Else
            vOPS.SetUnProcessedPayment(False, vOPS.PaymentAmount)
          End If
          vOPS.Save()
        Next vOPS
      End If

      'everything went OK. Now delete all the tdr lines, as we dont need to pass them back.
      For vIndex = 0 To pTransaction.TraderAnalysisLines.Count - 1
        pTransaction.DeleteTraderAnalysisLine((1))
      Next

      If pTransaction.OriginalOPS IsNot Nothing Then
        'BR19606 - Restore the OPS to it original state, but time stamp it as we have changed it (and changed it back), two OPS calsses used, pTransaction.OriginalOPS thinks it is new
        Dim vOriginalOps As New OrderPaymentSchedule
        vOriginalOps.Init(mvEnv, pTransaction.OriginalOPS.ScheduledPaymentNumber)
        vOriginalOps.Update(pTransaction.OriginalOPS.DueDate, pTransaction.OriginalOPS.AmountDue, pTransaction.OriginalOPS.AmountOutstanding, pTransaction.OriginalOPS.ExpectedBalance, pTransaction.OriginalOPS.ClaimDate)
        vOriginalOps.Save(mvEnv.User.UserID)
        pTransaction.OriginalOPS = Nothing 'OriginalOPS only exists to allow the OPS to be restored when the user cancels. As this has been done, it should be deleted
      End If

      If Len(vErrorMsg) > 0 Then
        pResults.Add("InformationMessage", CDBField.FieldTypes.cftCharacter, vErrorMsg)
      End If

    End Function

    Public Sub DeleteInvoicePayment(ByVal pTDRLine As TraderAnalysisLine, ByRef pErrorMsg As String, ByVal pExistingTransaction As Boolean, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vInvoicePaid As Integer
      Dim vInvoiceUsed As Integer
      Dim vAmount As Double
      Dim vRS As CDBRecordSet
      Dim vInvoice As New Invoice
      Dim vTrans As Boolean

      vInvoicePaid = pTDRLine.InvoiceNumber
      vAmount = pTDRLine.Amount

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      Select Case pTDRLine.TraderLineType
        Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
          'update unallocated-cash invoice (for taltInvoicePayment and taltInvoiceAllocation) and sundry credit note (for taltSundryCreditNoteInvoiceAllocation)
          If pTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation OrElse pTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation Then
            vInvoiceUsed = pTDRLine.InvoiceNumberUsed
            vInvoice.Init(mvEnv)
            vRS = mvEnv.Connection.GetRecordSet("SELECT " & vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", bt.amount FROM invoices i, batch_transactions bt WHERE i.invoice_number = " & vInvoiceUsed & " AND bt.batch_number = i.batch_number AND bt.transaction_number = i.transaction_number")
            With vRS
              If .Fetch() = True Then
                vInvoice = New Invoice
                vInvoice.InitFromRecordSet(mvEnv, vRS, Invoice.InvoiceRecordSetTypes.irtAll)
                vInvoice.InvoiceAmount = .Fields("amount").DoubleValue 'amount 
                vInvoice.AmountUsed = vAmount * -1
                vInvoice.SCUpdatePayment()

                'Update credit customer's outstanding amount
                If pTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment Then UpdateOutstanding(vInvoice.Company, vInvoice.SalesLedgerAccount, pTDRLine.Amount, True, pErrorMsg)
              Else
                If pErrorMsg <> "" Then pErrorMsg = pErrorMsg & vbCrLf
                pErrorMsg = pErrorMsg & String.Format(ProjectText.String29045, CStr(vInvoiceUsed)) 'Failed to retrieve the details for invoice %s
              End If
              .CloseRecordSet()

              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) Then
                If pExistingTransaction = True And pBatchNumber > 0 And pTransactionNumber > 0 Then
                  Dim vDeleteFields As New CDBFields()
                  With vDeleteFields
                    .Add("allocation_batch_number", pBatchNumber)
                    .Add("allocation_transaction_number", pTransactionNumber)
                    .Add("allocation_line_number", pTDRLine.LineNumber)
                  End With
                  mvEnv.Connection.DeleteRecords("invoice_payment_history", vDeleteFields, False)
                End If
              End If

            End With
          End If
          'update invoice that was paid
          vInvoice = New Invoice
          vInvoice.Init(mvEnv)
          vRS = mvEnv.Connection.GetRecordSet("SELECT amount," & vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & " FROM invoices i, batch_transactions bt WHERE i.invoice_number = " & vInvoicePaid & " AND bt.batch_number = i.batch_number AND bt.transaction_number = i.transaction_number")
          With vRS
            If .Fetch() = True Then
              vInvoice = New Invoice
              vInvoice.InitFromRecordSet(mvEnv, vRS, Invoice.InvoiceRecordSetTypes.irtAll)
              vInvoice.SCSetPaymentValues(0, pTDRLine.Amount * -1)
              vInvoice.InvoiceAmount = .Fields(1).DoubleValue 'amount   
              vInvoice.NowPaid = pTDRLine.Amount * -1

              vInvoice.SCUpdatePayment()
              'Update credit customer's outstanding amount
              If pTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment Then UpdateOutstanding(vInvoice.Company, vInvoice.SalesLedgerAccount, pTDRLine.Amount, True, pErrorMsg)
            Else
              If pErrorMsg <> "" Then pErrorMsg = pErrorMsg & vbCrLf
              pErrorMsg = pErrorMsg & String.Format(ProjectText.String29046, CStr(vInvoicePaid)) 'Failed to retrieve the details for invoice %s
            End If
            .CloseRecordSet()
          End With

        Case TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash, TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNote
          If pTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash Then 'Or (pTDRLine.TraderLineType = taltSundryCreditNote And Not mvNewCredit) Then 'need to add this bit when we implement credit notes in trader
            UpdateOutstanding(CACompany, pTDRLine.SalesLedgerAccount, vAmount, False, pErrorMsg)
          End If
      End Select

      If vTrans Then mvEnv.Connection.CommitTransaction()

    End Sub

    Private Sub SCGetDefaults(ByRef pPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByRef pResults As CDBParameters, ByVal pTransaction As TraderTransaction)
      Dim vDate As String = ""
      Dim vInc As Integer
      Dim vBalance As Double
      Dim vAmount As Double
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vOPSUnpaid As Boolean
      Dim vTDRLine As TraderAnalysisLine
      Dim vRS As CDBRecordSet
      Dim vRate As ProductRate
      Dim vSQL As String
      Dim vSalesContactNumber As Integer
      Dim vPaymentPlan As PaymentPlan
      Dim vRenewalAmount As Double
      Dim vRFD As String = ""
      Dim vTransactionType As String
      Dim vMembershipType As MembershipType
      Dim vAssocMemberType As MembershipType
      Dim vContact As Contact = Nothing
      Dim vCount As Integer
      Dim vDOB As String = ""
      Dim vGotAssociate As Boolean
      Dim vMemberCount As Integer
      Dim vMember As Member = Nothing
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vOPSInclude As Boolean
      Dim vLineNumber As Integer
      Dim vPPD As PaymentPlanDetail
      Dim vAppliedDate As String
      Dim vBranchMember As Boolean
      Dim vJoinedDate As String
      Dim vContact1 As Contact = Nothing
      Dim vContact2 As Contact = Nothing
      Dim vForceGifted As Boolean
      Dim vSCPPDLines As TraderPaymentPlanDetails = Nothing
      Dim vPM As String
      Dim vUpdateType As PaymentPlan.PaymentPlanUpdateTypes
      Dim vPO As PurchaseOrder
      Dim vPONumber As Integer
      Dim vPI As PurchaseInvoice
      Dim vPINumber As Integer
      Dim vMembershipPeriod As PaymentPlan.MembershipPeriodTypes
      Dim vEventBooking As EventBooking

      Select Case pPageType
        Case TraderPage.TraderPageType.tpAmendEventBooking
          If pTransaction.TraderAnalysisLines.Count > 0 Then
            'Should only have the one line
            vTDRLine = pTransaction.TraderAnalysisLines(1)
            vEventBooking = New EventBooking
            vEventBooking.Init(mvEnv, 0, (vTDRLine.EventBookingNumber))
            With vEventBooking
              If .Existing Then pTransaction.TraderAnalysisLines(1).SetEventBookingInfo(.EventNumber, .OptionNumber, CStr(.AdultQuantity), CStr(.ChildQuantity), .StartTime, .EndTime, 0)
            End With
          End If

        Case TraderPage.TraderPageType.tpBankDetails
          If mvEnv.GetConfig("fp_ba_default_sortcode").Length > 0 Then
            pResults.Add("SortCode", CDBField.FieldTypes.cftCharacter, mvEnv.GetConfig("fp_ba_default_sortcode"))
          End If

          'BankDetails False
          If mvAppType = ApplicationType.atCreditListReconciliation Then
            If pResults.ContainsKey("SortCode") Then pResults.Remove("SortCode") 'remove default sort code if set
            pResults.Add("SortCode", pParams("PayersSortCode").Value)
            pResults.Add("AccountNumber", pParams("PayersAccountNumber").Value)
            pResults.Add("AccountName", pParams("PayersName").Value)
            pResults.Add("Reference", pParams("ReferenceNumber").Value)
          End If

        Case TraderPage.TraderPageType.tpCreditCustomer
          If PayMethodsAtEnd Then
            If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
            If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          End If
          pResults.Add("TermsFrom", CDBField.FieldTypes.cftCharacter, "I")

        Case TraderPage.TraderPageType.tpEventBooking
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("Quantity", 1)
          If Len(pParams.ParameterExists("TransactionDistributionCode").Value) > 0 Then pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, pParams("TransactionDistributionCode").Value)
          If pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then pResults.Add("SalesContactNumber", pParams("SalesContactNumber").IntegerValue)

        Case TraderPage.TraderPageType.tpExamBooking
          vContact = New Contact(mvEnv)
          vContact.Init()
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then
            pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
            vContact.Init(pParams("PayerContactNumber").IntegerValue)
          End If
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("ExamSessionCode", Me.DefaultExamSessionCode)
          pResults.Add("ExamUnitCode", Me.DefaultExamUnitCode)
          If vContact.Existing = True AndAlso vContact.ContactReference.Length > 0 Then pResults.Add("ContactReference", CDBField.FieldTypes.cftCharacter, vContact.ContactReference)

        Case TraderPage.TraderPageType.tpAccommodationBooking
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("Quantity", 1)
          If Len(pParams.ParameterExists("TransactionDistributionCode").Value) > 0 Then pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, pParams("TransactionDistributionCode").Value)
          If pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then pResults.Add("SalesContactNumber", pParams("SalesContactNumber").IntegerValue)

        Case TraderPage.TraderPageType.tpCollectionPayments

        Case TraderPage.TraderPageType.tpProductDetails
          GetSegmentProductRate(pParams.ParameterExists("TransactionSource").Value, pParams("TransactionType").Value, pParams.ParameterExists("TransactionPaymentMethod").Value, pPageType, Batch.AdjustmentTypes.atNone)
          If pParams("TransactionLines").IntegerValue = 0 OrElse pParams.ParameterExists("GetDefaultProductAndRate").Bool Then
            pResults.Add("Product", CDBField.FieldTypes.cftCharacter, Me.DefaultProductCode(True))
            pResults.Add("Rate", CDBField.FieldTypes.cftCharacter, Me.DefaultRateCode(True))
            If Len(Me.Product) > 0 And Len(Me.SalesQuantity) > 0 Then
              pResults.Add("Quantity", CDBField.FieldTypes.cftLong, Me.SalesQuantity)
            Else
              pResults.Add("Quantity", 1)
            End If
            If Me.IsDefaultProductAndRate Then
              'If we have set a product and rate and the result is a price of zero (donation)
              'Then set the amount of the donation to the amount of the transaction
              vRate = New ProductRate(mvEnv)
              vRate.Init(Me.DefaultProductCode(True), Me.DefaultRateCode(True))
              If vRate.Existing = True And vRate.PriceIsZero Then pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(pParams.ParameterExists("TransactionAmount").DoubleValue))
            End If
          Else
            pResults.Add("Quantity", 1)
          End If
          If Len(pParams.ParameterExists("TransactionDistributionCode").Value) > 0 Then pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, pParams("TransactionDistributionCode").Value)
          If Len(pParams.ParameterExists("TransactionSource").Value) > 0 Then pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams("TransactionSource").Value)
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("When", CDBField.FieldTypes.cftDate, If(IsDate(pParams.ParameterExists("TransactionDate").Value), pParams("TransactionDate").Value, TodaysDate))
          pResults.Add("LineType", CDBField.FieldTypes.cftCharacter, "G")
          If pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then pResults.Add("SalesContactNumber", pParams("SalesContactNumber").IntegerValue)
        Case TraderPage.TraderPageType.tpTransactionDetails
          If pParams.ParameterExists("ExistingTransaction").Bool = True Or mvExistingAdjustmentTran Then
            'Existing Transaction, need to return current Amount,Mailing,Source,TransactionDate,EligibleForGiftAid
            If pParams.Exists("BatchNumber") = True And pParams.Exists("TransactionNumber") = True Then
              vSQL = "SELECT currency_amount, mailing, transaction_date,eligible_for_gift_aid FROM batch_transactions"
              vSQL = vSQL & " WHERE batch_number = " & pParams("BatchNumber").IntegerValue & " AND transaction_number = " & pParams("TransactionNumber").IntegerValue
              vRS = mvEnv.Connection.GetRecordSet(vSQL)
              If vRS.Fetch() = True Then
                pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(vRS.Fields("currency_amount").DoubleValue))
                pResults.Add("Mailing", CDBField.FieldTypes.cftCharacter, vRS.Fields("mailing").Value)
                pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, vRS.Fields("transaction_date").Value)
                pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, vRS.Fields("eligible_for_gift_aid").Value)
              End If
              vRS.CloseRecordSet()
            End If
            pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams.ParameterExists("TransactionSource").Value)
          Else
            If (Len(Me.Source) = 0 And mvEnv.GetConfigOption("fp_retain_source")) Or (Me.SourceFromLastMailing And Len(pParams.ParameterExists("TransactionSource").Value) > 0) Then
              pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams.ParameterExists("TransactionSource").Value)
            Else
              pResults.Add("Source", CDBField.FieldTypes.cftCharacter, Me.Source)
            End If
            If Len(Me.Campaign) > 0 Then
              pResults.Add("Campaign", CDBField.FieldTypes.cftCharacter, Me.Campaign)
              pResults.Add("Appeal", CDBField.FieldTypes.cftCharacter, Me.Appeal)
            End If
            If Me.Mailing.Length > 0 Then
              pResults.Add("Mailing", CDBField.FieldTypes.cftCharacter, Me.Mailing)
            End If
            If Me.AutoSetAmount And Me.GiftAidMinimum = 0 And Not (Me.Voucher Or Me.CAFCard Or Me.GiftInKind Or Me.SaleOrReturn) Then
              pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, "Y")
            Else
              pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, "N")
            End If
            If Me.BatchDate.Length > 0 Then
              pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, Me.BatchDate)
            ElseIf pParams.Exists("TransactionDate") Then
              'BR15961: transaction date was set in a previous transaction on this trader app - should be used again to match RC functionality
              pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, pParams("TransactionDate").Value.ToString)
            Else
              pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, TodaysDate())
            End If
          End If
          pResults.Add("Receipt", CDBField.FieldTypes.cftCharacter, "N")
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then
            pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
            pResults.Add("MailingContactNumber", pParams("PayerContactNumber").IntegerValue)
          End If
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then
            pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
            pResults.Add("MailingAddressNumber", pParams("PayerAddressNumber").IntegerValue)
          End If
          If pParams.ParameterExists("ExistingTransaction").Bool = False And mvExistingAdjustmentTran = False Then
            'When editing an existing transaction, do not reset these values
            If Me.CurrentPrice > 0 Then
              If Len(Me.Product) > 0 And Len(Me.SalesQuantity) > 0 Then
                pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(Me.CurrentPrice * Val(Me.SalesQuantity)))
              Else
                pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(Me.CurrentPrice))
              End If
            ElseIf Me.AutoSetAmount Then
              pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(0))
            End If
          End If

          If Len(pParams.ParameterExists("Reference").Value) > 0 Then pResults.Add("Reference", CDBField.FieldTypes.cftCharacter, pParams("Reference").Value)

          If mvEnv.GetConfig("fp_sales_contact_type") <> "ORGANISATION" Then
            vSalesContactNumber = mvEnv.User.SalesContactNumber
          End If
          If vSalesContactNumber > 0 Then
            pResults.Add("SalesContactNumber", vSalesContactNumber)
          ElseIf pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then
            pResults.Add("SalesContactNumber", pParams("SalesContactNumber").IntegerValue)
          ElseIf Me.DefaultSalesContact > 0 Then
            pResults.Add("SalesContactNumber", Me.DefaultSalesContact)
          End If

          If Me.DefaultTransactionOrigin.Length > 0 Then
            pResults.Add("TransactionOrigin", CDBField.FieldTypes.cftCharacter, Me.DefaultTransactionOrigin)
          End If

        Case TraderPage.TraderPageType.tpCardDetails
          If Not ((CreditCard Or CAFCard Or CCWithInvoice) And DebitCard) Then
            If CreditCard Or CAFCard Then
              pResults.Add("CreditOrDebitCard", CDBField.FieldTypes.cftCharacter, "C")
            Else
              pResults.Add("CreditOrDebitCard", CDBField.FieldTypes.cftCharacter, "D")
            End If
          End If
          If CAFCard AndAlso pParams.ParameterExists("TransactionPaymentMethod").Value = "CAFC" Then pResults.Add("CardNumber", CDBField.FieldTypes.cftLong, DEFAULT_CAF_CARD_NUMBER)
          If Len(pParams.ParameterExists("TransactionReference").Value) > 0 Then pResults.Add("Reference", CDBField.FieldTypes.cftCharacter, pParams("TransactionReference").Value)
        Case TraderPage.TraderPageType.tpMembership, TraderPage.TraderPageType.tpChangeMembershipType
          'Set up the member class
          vPaymentPlan = New PaymentPlan
          vPaymentPlan.Init(mvEnv, (pParams.ParameterExists("PaymentPlanNumber").IntegerValue))
          'For an existing PaymentPlan, will need to initialise with the PaymentPlan number
          vPaymentPlan.LoadMembers() 'If existing then loads members otherwise inits .Member
          With vPaymentPlan.Member
            If pParams("TransactionType").Value = "MEMC" Then 'CMT
            Else
              'New Membership not CMT, set details from payer
              If Not Me.BlankMembershipJoinedDate Then
                .SetJoinedDate(pParams.OptionalValue("TransactionDate", (TodaysDate())))
              End If
              DefaultPPMemberToPayer(vPaymentPlan, pParams("PayerContactNumber").IntegerValue, pParams("PayerAddressNumber").IntegerValue)
              If Me.AppType = ApplicationType.atConversion Then
                .Source = vPaymentPlan.Source
              Else
                .Source = pParams.ParameterExists("TransactionSource").Value
                If .Source.Length = 0 Then .Source = Source
              End If
            End If

            'Set parameters for default values
            If (pPageType = TraderPage.TraderPageType.tpMembership And Len(.MembershipTypeCode) > 0) Then pResults.Add("MembershipType", CDBField.FieldTypes.cftCharacter, .MembershipTypeCode)
            If Not Me.BlankMembershipJoinedDate Then pResults.Add("Joined", CDBField.FieldTypes.cftDate, .GetNewJoinedDate(pParams.ParameterExists("TransactionDate").Value))
            pResults.Add("NumberOfMembers", .NumberOfMembers)
            If .MembershipTypeCode.Length > 0 Then pResults.Add("MaxFreeAssociates", .MembershipType.MaxFreeAssociates)
            pResults.Add("GiftMembership", CDBField.FieldTypes.cftCharacter, BooleanString(vPaymentPlan.GiftMembership))
            pResults.Add("OneYearGift", CDBField.FieldTypes.cftCharacter, BooleanString(vPaymentPlan.OneYearGift))
            If .Source.Length > 0 Then pResults.Add("Source", CDBField.FieldTypes.cftCharacter, .Source)
            If pParams("TransactionType").Value = "MEMC" Then
              'CMT
              pResults.Add("MembershipTypeDesc", CDBField.FieldTypes.cftCharacter, .MembershipType.MembershipTypeDesc)
              pResults.Add("Branch", CDBField.FieldTypes.cftCharacter, .Branch)
              pResults.Add("OriginalJoined", .Joined)
            ElseIf Me.DefaultMemberBranch.Length > 0 Then
              pResults.Add("Branch", CDBField.FieldTypes.cftCharacter, Me.DefaultMemberBranch)
            End If
            pResults.Add("BranchMember", CDBField.FieldTypes.cftCharacter, .BranchMember)
            If pParams("TransactionType").Value = "MEMC" And pPageType = TraderPage.TraderPageType.tpChangeMembershipType Then
              pResults.Add("Applied", CDBField.FieldTypes.cftDate, .Applied)
            Else
              pResults.Add("Applied", CDBField.FieldTypes.cftDate, .GetNewJoinedDate(pParams.ParameterExists("TransactionDate").Value))
            End If
            If Len(.AgeOverride) > 0 Then pResults.Add("AgeOverride", CDBField.FieldTypes.cftLong, .AgeOverride)
            If vPaymentPlan.Details.Count() > 0 Then pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, CType(vPaymentPlan.Details.Item(1), PaymentPlanDetail).DistributionCode)
            If pPageType = TraderPage.TraderPageType.tpChangeMembershipType Then
              If Me.CMTCancelReason.Length > 0 Then pResults.Add("CancellationReason", CDBField.FieldTypes.cftCharacter, Me.CMTCancelReason)
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, vPaymentPlan.PaymentFrequencyCode)
              pResults.Add("EligibleforGiftAid", CDBField.FieldTypes.cftCharacter, BooleanString(vPaymentPlan.EligibleForGiftAid))
              Dim vWriteOff As Boolean = mvEnv.GetConfigOption("fp_pp_wo_missed_payments", False)
              If vPaymentPlan.CMTProportionBalance = PaymentPlan.CMTProportionBalanceTypes.cmtNone Then
                vWriteOff = False   'Never write-off for this setting.
                pResults.Add("DisableWriteOffMissedPayments", CDBField.FieldTypes.cftCharacter, "Y")
              End If
              pResults.Add("WriteOffMissedPayments", BooleanString(vWriteOff))
              pResults.Add("CMTDate", CDBField.FieldTypes.cftDate, TodaysDate)
              pResults.Add("CMTEarliestDate", CDBField.FieldTypes.cftDate, .GetEarliestCMTDate(vPaymentPlan.TermStartDate))
              pResults.Add("CMTLatestDate", CDBField.FieldTypes.cftDate, vPaymentPlan.TermEndDate.AddDays(-1).ToString(CAREDateFormat))
            ElseIf Me.AppType = ApplicationType.atConversion Then
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, vPaymentPlan.PaymentFrequencyCode)
            Else
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, GetDefaultPayFreq(GetPayMethodCode((pParams("PPPaymentType").Value))))
            End If
            pResults.Add("GiftCardStatus_N", CDBField.FieldTypes.cftCharacter, "N")
            pResults.Add("ContactNumber", .ContactNumber)
            pResults.Add("AddressNumber", .AddressNumber)
            If Len(.ContactDateOfBirth) > 0 Then pResults.Add("DateOfBirth", CDBField.FieldTypes.cftDate, .ContactDateOfBirth)
            pResults.Add("DobEstimated", CDBField.FieldTypes.cftCharacter, BooleanString(.ContactDOBEstimated))
          End With

        Case TraderPage.TraderPageType.tpStandingOrder, TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpCreditCardAuthority
          vTransactionType = pParams.ParameterExists("TransactionType").Value
          If vTransactionType = "APAY" Then
            pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
            With pTransaction
              If Not .PaymentPlan.Existing Then RaiseError(DataAccessErrors.daePaymentPlanNotFound)
              pResults.Add("ContactNumber", pParams("ContactNumber").IntegerValue)
              pResults.Add("AddressNumber", pParams("AddressNumber").IntegerValue)
              pResults.Add("PaymentPlanNumber", .PaymentPlan.PlanNumber)
              Select Case pPageType
                Case TraderPage.TraderPageType.tpStandingOrder
                  With .PaymentPlan.StandingOrder
                    pResults.Add("SortCode", CDBField.FieldTypes.cftInteger, .ContactAccount.FormattedSortCode)
                    pResults.Add("AccountNumber", CDBField.FieldTypes.cftLong, .ContactAccount.AccountNumber)
                    pResults.Add("AccountName", CDBField.FieldTypes.cftCharacter, .ContactAccount.AccountName)
                    pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, .BankAccount)
                    pResults.Add("Reference", CDBField.FieldTypes.cftCharacter, .Reference)
                    pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(.Amount))
                    pResults.Add("StartDate", CDBField.FieldTypes.cftDate, .StartDate)
                    pResults.Add("Source", CDBField.FieldTypes.cftCharacter, .Source)
                    pResults.Add("StandingOrderType", CDBField.FieldTypes.cftCharacter, .StandingOrderTypeCode)
                    pResults.Add("BankDetailsNumber", .BankDetailsNumber)
                    pResults.Add("IbanNumber", CDBField.FieldTypes.cftLong, .ContactAccount.IbanNumber)
                    pResults.Add("BicCode", CDBField.FieldTypes.cftLong, .ContactAccount.BicCode)
                  End With
                Case TraderPage.TraderPageType.tpDirectDebit
                  With .PaymentPlan.DirectDebit
                    pResults.Add("SortCode", CDBField.FieldTypes.cftInteger, .ContactAccount.SortCode)
                    pResults.Add("AccountNumber", CDBField.FieldTypes.cftLong, .ContactAccount.AccountNumber)
                    pResults.Add("AccountName", CDBField.FieldTypes.cftCharacter, .ContactAccount.AccountName)
                    pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, .BankAccount)
                    pResults.Add("Reference", CDBField.FieldTypes.cftCharacter, .Reference)
                    pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(DoubleValue(.Amount)))
                    pResults.Add("StartDate", CDBField.FieldTypes.cftDate, .StartDate)
                    pResults.Add("Source", CDBField.FieldTypes.cftCharacter, .Source)
                    pResults.Add("MandateType", CDBField.FieldTypes.cftCharacter, .MandateType)
                    If pTransaction.PaymentPlan.AutoPaymentClaimDateMethod = PaymentPlan.AutoPaymentClaimDateMethods.apcdmDays And pTransaction.PaymentPlan.ClaimDay.Length > 0 Then pResults.Add("ClaimDay", CDBField.FieldTypes.cftInteger, pTransaction.PaymentPlan.ClaimDay)
                    pResults.Add("BankDetailsNumber", .BankDetailsNumber)
                    pResults.Add("DateSigned", CDBField.FieldTypes.cftDate, .DateSigned)
                    pResults.Add("IbanNumber", CDBField.FieldTypes.cftLong, .ContactAccount.IbanNumber)
                    pResults.Add("BicCode", CDBField.FieldTypes.cftLong, .ContactAccount.BicCode)
                  End With
                Case TraderPage.TraderPageType.tpCreditCardAuthority
                  With .PaymentPlan.CreditCardAuthority
                    pResults.Add("CreditCardType", CDBField.FieldTypes.cftCharacter, .ContactCreditCard.CreditCardType)
                    pResults.Add("CreditCardNumber", CDBField.FieldTypes.cftLong, .ContactCreditCard.CreditCardNumber)
                    pResults.Add("ExpiryDate", CDBField.FieldTypes.cftDate, .ContactCreditCard.ExpiryDate)
                    pResults.Add("Issuer", CDBField.FieldTypes.cftCharacter, .ContactCreditCard.Issuer)
                    pResults.Add("IssueNumber", CDBField.FieldTypes.cftInteger, .ContactCreditCard.IssueNumber)
                    pResults.Add("AccountName", CDBField.FieldTypes.cftCharacter, .ContactCreditCard.AccountName)
                    pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, .BankAccount)
                    pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(DoubleValue(.Amount)))
                    pResults.Add("StartDate", CDBField.FieldTypes.cftDate, .StartDate)
                    pResults.Add("Source", CDBField.FieldTypes.cftCharacter, .Source)
                    pResults.Add("AuthorityType", CDBField.FieldTypes.cftCharacter, .GetAuthorityTypeCode(.AuthorityType))
                    If pTransaction.PaymentPlan.AutoPaymentClaimDateMethod = PaymentPlan.AutoPaymentClaimDateMethods.apcdmDays And pTransaction.PaymentPlan.ClaimDay.Length > 0 Then pResults.Add("ClaimDay", CDBField.FieldTypes.cftInteger, pTransaction.PaymentPlan.ClaimDay)
                  End With
              End Select
            End With
          Else
            'contact_number
            pResults.Add("ContactNumber", CDBField.FieldTypes.cftLong, pParams("PayerContactNumber").Value)
            'address_number
            pResults.Add("AddressNumber", CDBField.FieldTypes.cftLong, pParams("PayerAddressNumber").Value)
            'source
            If Not (AppType = ApplicationType.atConversion And mvEnv.GetConfigOption("trader_conv_app_default_source") = False) Then
              If pParams.Exists("TransactionSource") Then
                pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams("TransactionSource").Value)
              Else 'When AppType=Converstion AND Payment Plan Details are not displayed
                pResults.Add("Source", CDBField.FieldTypes.cftCharacter, Source)
              End If
            End If
            'start_date
            vDate = ""
            Dim vAutoPayMethod As PaymentPlan.ppAutoPayMethods
            Dim vBankAccount As BankAccount = Nothing
            If pPageType = TraderPage.TraderPageType.tpDirectDebit Then
              vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMDD
              vBankAccount = mvEnv.BankAccount(DDBankAccount)
            ElseIf pPageType = TraderPage.TraderPageType.tpCreditCardAuthority Then
              vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMCCCA
              vBankAccount = mvEnv.BankAccount(CCABankAccount)
            Else
              vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMSO
              vBankAccount = mvEnv.BankAccount(SOBankAccount)
            End If
            If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
              vDate = mvEnv.GetPaymentPlanAutoPayDate(Today, vAutoPayMethod, vBankAccount).ToString(CAREDateFormat)
              pResults.Add("StartDate", CDBField.FieldTypes.cftDate, vDate)
            Else
              Dim vBaseDate As Nullable(Of Date) = Nothing
              Dim vMemJoinedDate As Nullable(Of Date) = Nothing
              Dim vStartMonth As Nullable(Of Integer) = Nothing
              Select Case vTransactionType
                Case "MEMB", "CMEM"
                  pTransaction.PaymentPlan.SetMember((pParams("MEM_MembershipType").Value), (pParams("MEM_Branch").Value), (pParams.ParameterExists("MEM_GiftMembership").Bool), (pParams.ParameterExists("MEM_OneYearGift").Bool), (pParams.ParameterExists("MEM_GiverContactNumber").Value))
                  vMemJoinedDate = DateValue(pParams("MEM_Joined").Value)
                Case "SUBS", "DONR"
                  If (pPageType = TraderPage.TraderPageType.tpCreditCardAuthority Or pPageType = TraderPage.TraderPageType.tpDirectDebit) Then
                    'If StartMonth control is visible then default DD/CCCA StartDate to Day(PP StartDate) + StartMonth + Year(PP Start Date)
                    'Assume that if we have a StartMonth then the control must have been visible
                    If pParams.ParameterExists("PPD_StartMonth").IntegerValue > 0 Then
                      vStartMonth = pParams("PPD_StartMonth").IntegerValue
                    End If
                  End If
              End Select

              If vBaseDate.HasValue = False Then
                Dim vPrefix As String = "PPD"
                If vTransactionType = "LOAN" Then vPrefix = "LON"
                vBaseDate = DateValue(pParams(vPrefix & "_OrderDate").Value)
              End If

              vDate = pTransaction.PaymentPlan.GetAutoPaymentDefaultStartDate(vBaseDate.Value, vAutoPayMethod, vBankAccount, pParams.ParameterExists("MEM_MembershipType").Value, vMemJoinedDate, vTransactionType, vStartMonth).ToString(CAREDateFormat)
              If Not (IsDate(vDate)) Then vDate = TodaysDate()

              pResults.Add("StartDate", CDBField.FieldTypes.cftDate, vDate)
            End If
            'amount
            If pPageType = TraderPage.TraderPageType.tpDirectDebit Then
              If Len(pParams.ParameterExists("PPD_Amount").Value) > 0 Then
                pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(GetFrequencyAmount(pParams, (pParams.ParameterExists("PaymentPlanNumber").IntegerValue))))
              End If
            ElseIf pPageType = TraderPage.TraderPageType.tpStandingOrder AndAlso vTransactionType = "LOAN" Then
              Dim vFreqAmount As Double
              If pParams.ParameterExists("LON_FixedMonthlyAmount").DoubleValue > 0 Then
                vFreqAmount = pParams("LON_FixedMonthlyAmount").DoubleValue
              Else
                vFreqAmount = pParams.ParameterExists("LON_LoanAmount").DoubleValue / 12
              End If
              pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, vFreqAmount.ToString)
            Else
              pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(GetFrequencyAmount(pParams, (pParams.ParameterExists("PaymentPlanNumber").IntegerValue))))
            End If
            'bank_account etc.
            Select Case pPageType
              Case TraderPage.TraderPageType.tpStandingOrder
                pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, SOBankAccount)
                pResults.Add("StandingOrderType", CDBField.FieldTypes.cftCharacter, "B")
              Case TraderPage.TraderPageType.tpDirectDebit
                pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, DDBankAccount)
                pResults.Add("MandateType") 'Add this as a null value so that combo defaults to 'Unknown'
                pResults.Add("DateSigned", CDBField.FieldTypes.cftDate)
                If mvEnv.GetConfigOption("fp_dd_signed_date_mandatory", False) Then pResults("DateSigned").Value = TodaysDate()
              Case TraderPage.TraderPageType.tpCreditCardAuthority
                pResults.Add("BankAccount", CDBField.FieldTypes.cftCharacter, CCABankAccount)
                pResults.Add("AuthorityType", CDBField.FieldTypes.cftCharacter, "A")
            End Select
            'sort_code
            If mvEnv.GetConfig("fp_ba_default_sortcode").Length > 0 Then
              Select Case pPageType
                Case TraderPage.TraderPageType.tpDirectDebit, TraderPage.TraderPageType.tpStandingOrder
                  pResults.Add("SortCode", CDBField.FieldTypes.cftCharacter, mvEnv.GetConfig("fp_ba_default_sortcode"))
              End Select
            End If
            'text1
            If pPageType = TraderPage.TraderPageType.tpDirectDebit And (mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL") Then
              pResults.Add("Text1", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultDDText1))
            End If

          End If

        Case TraderPage.TraderPageType.tpPaymentPlanDetails
          vTransactionType = pParams("TransactionType").Value
          Dim vOrderDate As String = String.Empty
          If vTransactionType = "SALE" OrElse vTransactionType = "EVNT" OrElse vTransactionType = "ACOM" OrElse vTransactionType = "SRVC" Then 'When called from TPP page
            vBalance = 0
            Dim vGotInfo As Boolean = False
            Dim vEB As EventBooking
            Dim vSB As ServiceBooking
            Dim vEAB As EventAccommodationBooking
            'Need to peruse the lines in the TAS grid and try to determine some order defaults
            For Each vTDRLine In pTransaction.TraderAnalysisLines
              Select Case vTDRLine.TraderLineType
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
                  If Not vGotInfo Then
                    vEB = New EventBooking
                    With vEB
                      .Init(mvEnv, , vTDRLine.EventBookingNumber)
                      If .Existing Then
                        vOrderDate = CType(.Sessions(1), EventSession).StartDate
                        If .SalesContactNumber > 0 Then vSalesContactNumber = .SalesContactNumber
                        vGotInfo = True
                      Else
                        RaiseError(DataAccessErrors.daeInvalidEventBookingLine) 'This transaction has an event booking analysis line that is linked to a non-existant event booking
                      End If
                    End With
                  End If
                  vBalance = vBalance + vTDRLine.Amount
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit
                  If Not vGotInfo Then
                    vSB = New ServiceBooking
                    With vSB
                      .Init(mvEnv, vTDRLine.ServiceBookingNumber)
                      If .Existing Then
                        vOrderDate = .StartDate
                        If .SalesContactNumber > 0 Then vSalesContactNumber = .SalesContactNumber
                        vGotInfo = True
                      Else
                        RaiseError(DataAccessErrors.daeInvalidServiceBookingLine) 'This transaction has a service booking analysis line that is linked to a non-existant service booking
                      End If
                    End With
                  End If
                  vBalance = vBalance + vTDRLine.Amount
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation
                  If Not vGotInfo Then
                    vEAB = New EventAccommodationBooking
                    With vEAB
                      .Init(mvEnv, vTDRLine.RoomBookingNumber)
                      If .Existing Then
                        vOrderDate = .FromDate
                        If .SalesContactNumber > 0 Then vSalesContactNumber = .SalesContactNumber
                        vGotInfo = True
                      Else
                        RaiseError(DataAccessErrors.daeInvalidAccomodationBookingLine) 'This transaction has an accommodation booking analysis line that is linked to a non-existant accommodation booking
                      End If
                    End With
                  End If
                  vBalance = vBalance + vTDRLine.Amount
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltProductSale
                  If Not vGotInfo Then
                    vOrderDate = If(IsDate(pParams.ParameterExists("TransactionDate").Value), pParams("TransactionDate").Value, TodaysDate)
                  End If
                  vBalance = vBalance + vTDRLine.Amount
              End Select
            Next
          End If

          vPaymentPlan = pTransaction.PaymentPlan
          vPaymentPlan.Init(mvEnv)
          If vTransactionType = "MEMB" Then
            If pParams.Exists("PayerContactNumber") = False Then
              pParams.Add("PayerContactNumber", CDBField.FieldTypes.cftLong, pParams("ContactNumber").Value)
              pParams.Add("PayerAddressNumber", CDBField.FieldTypes.cftLong, pParams("AddressNumber").Value)
            End If
            If pParams.Exists("MemberContactNumber") = False Then
              pParams.Add("MemberContactNumber", CDBField.FieldTypes.cftLong, pParams("ContactNumber").Value)
              pParams.Add("MemberAddressNumber", CDBField.FieldTypes.cftLong, pParams("AddressNumber").Value)
            End If
            If pParams.Exists("Rate") = True And pParams.Exists("MembershipRate") = False Then
              pParams.Add("MembershipRate", CDBField.FieldTypes.cftCharacter, pParams("Rate").Value)
            End If
          End If
          If Me.AppType = ApplicationType.atMaintenance Then
            vUpdateType = PaymentPlan.PaymentPlanUpdateTypes.pputPaymentPlan
          ElseIf Me.AppType = ApplicationType.atConversion Then
            vUpdateType = PaymentPlan.PaymentPlanUpdateTypes.pputConversion
            If Me.PayPlanConversionMaintenance Then vUpdateType = vUpdateType Or PaymentPlan.PaymentPlanUpdateTypes.pputPaymentPlan
          Else
            vUpdateType = PaymentPlan.PaymentPlanUpdateTypes.pputNone
          End If
          'order_date
          Select Case vTransactionType
            Case "CMEM", "CDON", "CSUB"

            Case "MEMB"
              If Me.AppType = ApplicationType.atConversion Then
                vPaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
                vPaymentPlan.SetMember((pParams("MembershipType").Value), (pParams("Branch").Value), (pParams("GiftMembership").Bool), (pParams("OneYearGift").Bool), CStr(pParams("GiverContactNumber").IntegerValue))
                vDate = vPaymentPlan.StartDate
              Else
                'Use the fixed cycle defined on the membership type, if there is one.
                vPaymentPlan.SetMember((pParams("MembershipType").Value), (pParams("Branch").Value), (pParams("GiftMembership").Bool), (pParams("OneYearGift").Bool), CStr(pParams("GiverContactNumber").IntegerValue))
                If vPaymentPlan.FixedRenewalCycle And Len(vPaymentPlan.MembershipType.FixedCycle) > 0 Then vPaymentPlan.SetMembershipTypeFixedCycle(vPaymentPlan.MembershipType.FixedCycle)
                vDate = vPaymentPlan.FixedRenewalDate(pParams("Joined").Value)
              End If
            Case "SUBS"
              vDate = mvEnv.GetStartDate(CDBEnvironment.ppType.pptOther)
            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              vDate = vOrderDate
            Case Else
              vDate = pParams.ParameterExists("TransactionDate").Value
          End Select
          If IsDate(vDate) Then
            pResults.Add("OrderDate", CDBField.FieldTypes.cftDate, vDate)
          Else
            pResults.Add("OrderDate", CDBField.FieldTypes.cftDate, TodaysDate())
          End If
          If pParams.Exists("StartDate") = False Then pParams.Add("StartDate", CDBField.FieldTypes.cftDate, vDate)
          'expiry_date
          Select Case vTransactionType
            Case "CMEM", "CDON", "CSUB"

            Case "MEMB"
              If Me.AppType = ApplicationType.atConversion Then
                vDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, 1, CDate(vPaymentPlan.ExpiryDate)))
                vInc = 0
              Else
                vDate = pParams("Joined").Value
                vInc = 99
              End If
            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              vDate = vOrderDate
              vInc = 1
            Case Else
              vDate = pParams.ParameterExists("TransactionDate").Value
              vInc = 99
          End Select
          If Not (IsDate(vDate)) Then vDate = TodaysDate()
          vDate = CDate(vDate).AddYears(vInc).ToString(CAREDateFormat)
          pResults.Add("ExpiryDate", CDBField.FieldTypes.cftDate, CDate(vDate).AddDays(-1).ToString(CAREDateFormat))
          'order_term
          Select Case vTransactionType
            Case "MEMB", "CMEM"
              pResults.Add("OrderTerm", vPaymentPlan.MembershipType.MembershipTerm)
            Case Else
              pResults.Add("OrderTerm", CDBField.FieldTypes.cftInteger, "1")
          End Select
          'source
          If pParams.Exists("TransactionSource") Then
            pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams.ParameterExists("TransactionSource").Value)
          Else
            pResults.Add("Source", CDBField.FieldTypes.cftCharacter, Me.Source)
          End If
          'payment_frequency
          If AppType = ApplicationType.atConversion Then
            pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, vPaymentPlan.PaymentFrequencyCode)
          Else
            If pParams.Exists("PaymentFrequency") Then
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, pParams("PaymentFrequency").Value)
            Else
              Select Case vTransactionType
                Case "MEMB", "CMEM", "DONR", "SUBS", "CMEM", "CDON" 'if a payment plan
                  vPM = GetPayMethodCode((pParams("PPPaymentType").Value))
                Case Else
                  vPM = GetPayMethodCode((pParams("TransactionPaymentMethod").Value))
              End Select
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, GetDefaultPayFreq(vPM))
            End If
          End If
          'balance
          Select Case vTransactionType
            Case "MEMB", "CMEM"
              'BR12477: get the membership period
              vMembershipPeriod = vPaymentPlan.DetermineMembershipPeriod
              If AppType = ApplicationType.atConversion And pTransaction.TraderPPDLines.Count > 0 Then
                vSCPPDLines = pTransaction.TraderPPDLines
                pTransaction.TraderPPDLines.Clear()
              End If
              vBalance = vPaymentPlan.GetMemberBalance(pParams, pTransaction, pResults("PaymentFrequency").Value, vRenewalAmount)
              ' 'Set up the membership products
              If Me.AppType = ApplicationType.atConversion Then
                vBalance = vBalance + vPaymentPlan.Balance
                If Not vSCPPDLines Is Nothing Then
                  vLineNumber = pTransaction.TraderPPDLines(pTransaction.TraderPPDLines.Count).LineNumber
                  For Each vPPD In vSCPPDLines
                    vLineNumber = vLineNumber + 1
                    vPPD.LineNumber = vLineNumber
                    pTransaction.TraderPPDLines.AddItem(vPPD, CStr(vPPD.LineNumber))
                  Next vPPD
                Else
                  pTransaction.TraderPPDLines.AddDetailLinesFromPaymentPlan(vPaymentPlan, True)
                End If
              End If
              pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))

            Case "CDON", "CSUB"

            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))
          End Select
          Select Case vTransactionType
            Case "MEMB", "CMEM"
              If mvEnv.GetConfigOption("reason_is_grade", True) Then
                vRFD = pParams("MembershipType").Value
              Else
                vRFD = MemReason
              End If
            Case "CSUB", "CDON"

            Case Else
              Select Case pParams("PPPaymentType").Value
                Case "DIRD"
                  vRFD = DDReason
                Case "CCCA"
                  vRFD = CCReason
                Case "STDO"
                  vRFD = SOReason
                Case Else
                  vRFD = OReason
              End Select
          End Select
          pResults.Add("ReasonForDespatch", CDBField.FieldTypes.cftCharacter, vRFD)
          '"sales_contact_number"
          Select Case vTransactionType
            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              If vSalesContactNumber > 0 Then pResults.Add("SalesContactNumber", vSalesContactNumber)
            Case Else
              If mvEnv.GetConfig("fp_sales_contact_type") <> "ORGANISATION" Then
                vSalesContactNumber = mvEnv.User.SalesContactNumber
              End If
              If pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then
                pResults.Add("SalesContactNumber", pParams("SalesContactNumber").IntegerValue)
              ElseIf vSalesContactNumber > 0 Then
                pResults.Add("SalesContactNumber", vSalesContactNumber)
              ElseIf Me.DefaultSalesContact > 0 Then
                pResults.Add("SalesContactNumber", Me.DefaultSalesContact)
              End If
          End Select

          Select Case vTransactionType
            Case "MEMB", "CMEM"
              'vRenewalAmount will have already ben set above along with vBalance
              vRenewalAmount = vPaymentPlan.GetProRataRenewalAmount(pTransaction.TraderPPDLines, If(vPaymentPlan.Existing = True, vPaymentPlan.StartDate, pResults("OrderDate").Value), If(vPaymentPlan.Existing = True, vPaymentPlan.RenewalDate, pResults("OrderDate").Value), vRenewalAmount, pParams("ContactNumber").IntegerValue, pParams.ParameterExists("MembershipType").Value, vUpdateType, vTransactionType, pParams.ParameterExists("TransactionPaymentMethod").Value, False, 0, 0, vMembershipPeriod) '.... transactionpaymentmethod is the PM3 page one
              pResults.Add("RenewalAmount", CDBField.FieldTypes.cftNumeric, FixedFormat(vRenewalAmount))
              'CalcRenewalAmountFromPPS
            Case "CDON", "CSUB"
            Case "SALE", "EVNT", "ACOM", "SRVC" 'When called from TPP page
              pResults.Add("RenewalAmount", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))
          End Select
          'checkbox
          pResults.Add("UseAsFirstAmount", CDBField.FieldTypes.cftCharacter, "I") 'UnChecked and Invisible
          If (vTransactionType = "MEMB" Or vTransactionType = "CMEM") Then
            If (vPaymentPlan.ProportionalBalanceSetting And PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment) = PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment And vPaymentPlan.MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfAnnualTerm Then
              pResults("UseAsFirstAmount").Value = "Y" 'Checked
            End If
          End If
          'first_amount
          pResults.Add("FirstAmountVisible", CDBField.FieldTypes.cftCharacter, "N")
          If (vTransactionType = "MEMB" Or vTransactionType = "CMEM") Then
            If (vPaymentPlan.ProportionalBalanceSetting And PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew) = PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew Then pResults("FirstAmountVisible").Value = "Y"
          End If
          'eligible_for_gift_aid
          pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, GetPayPlanEligibleForGiftAid(pParams, pTransaction))
          'StartMonth
          If (vTransactionType = "DONR" Or vTransactionType = "SUBS") Then
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPaymentPlanStartMonth) Then
              'If database supports StartMonth then return an additional parameter to default the StartDate
              vDate = ""
              vDate = mvEnv.GetStartDate(CDBEnvironment.ppType.pptOther, True)
              pResults.Add("FixedStartDate", CDBField.FieldTypes.cftDate, vDate)
              pResults.Add("StartMonth", Now.Month)
            End If
          End If
        Case TraderPage.TraderPageType.tpScheduledPayments
          Dim vAddBalance As Boolean
          For Each vOPS In pTransaction.PaymentPlan.ScheduledPayments
            vAddBalance = True
            Select Case vOPS.ScheduleCreationReason
              Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance
                vAddBalance = False
              Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments
                If (vOPS.AmountOutstanding > 0 AndAlso pTransaction.PaymentPlan.Balance = 0) AndAlso _
                  (pTransaction.PaymentPlan.PaymentFrequencyFrequency = 1 AndAlso pTransaction.PaymentPlan.PaymentFrequencyInterval = 1) Then vAddBalance = False
            End Select
            If vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then Exit For
            If vOPS.ScheduledPaymentStatus = OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment Then vOPSUnpaid = True
            If vAddBalance Then vBalance += vOPS.AmountOutstanding
          Next vOPS
          With pTransaction.PaymentPlan
            If (.StartDate = .RenewalDate) And Len(.FirstAmount) > 0 Then
              If Val(.FirstAmount) > .Balance Then
                'Initial period so .FirstAmount is expected to be paid
                vAmount = Val(.FirstAmount)
              Else
                vAmount = .Balance
              End If
              If vOPSUnpaid Then vAmount = vBalance
            ElseIf Len(.FirstAmount) > 0 And vBalance > 0 Then
              'Some payments have been made, amount expected is the sum of the amount outstanding
              vAmount = vBalance
            Else
              vAmount = If(vOPSUnpaid = True, vBalance, .Balance)
            End If
            pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(vAmount))
            pResults.Add("AmountOutstanding", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))
          End With
        Case TraderPage.TraderPageType.tpMembershipMembersSummary
          vMembershipType = mvEnv.MembershipType((pParams("MembershipType").Value))
          If vMembershipType.AssociateMembershipType.Length > 0 Then
            vAssocMemberType = mvEnv.MembershipType((vMembershipType.AssociateMembershipType))
          Else
            vAssocMemberType = New MembershipType(mvEnv)
            vAssocMemberType.Init()
          End If

          If pParams("CurrentPageType").IntegerValue = TraderPage.TraderPageType.tpAmendMembership Then
            vContact = New Contact(mvEnv)
            vContact.Init((pParams("ContactNumber").IntegerValue), (pParams("AddressNumber").IntegerValue))
            pTransaction.PaymentPlan.PlanNumber = 0
            pTransaction.PaymentPlan.AddMember((vContact.ContactNumber), vContact.Address.AddressNumber, (vMembershipType.MembershipTypeCode), vMembershipType, (vContact.ContactType))
            pTransaction.PaymentPlan.Member.SCAddMemberSummary(vContact, pParams("Joined").Value, pParams("Branch").Value, pParams("BranchMember").Bool, pParams("Applied").Value, pParams("DistributionCode").Value, pParams("AgeOverride").Value, pParams("DateOfBirth").Value, "", If(pParams("TransactionType").Value = "MEMC" And pParams.Exists("MembershipNumber") = True, pParams.ParameterExists("MembershipNumber").IntegerValue, 0))
            pTransaction.PaymentPlan.Member.ContactDOBEstimated = pParams("DobEstimated").Bool
            pResults.Add("CurrentMembers", pParams("CurrentMembers").IntegerValue)
          ElseIf pParams.ParameterExists("TransactionType").Value = "MEMC" Then
            'CMT
            With pTransaction.PaymentPlan
              .Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
              .LoadMembers()
              vMemberCount = 0
              pResults.Add("CurrentMembers", .CurrentMembers.Count())
              'In order to display the MembersMemberSummary grid, need to "manipulate" the members data
              For Each vMember In .CurrentMembers
                vGotAssociate = False
                If vMembershipType.AssociateMembershipType.Length > 0 Then
                  vDOB = vMember.ContactDateOfBirth
                  'If Len(vDOB) = 0 Then vDOB = pParams("DateOfBirth").Value
                  If IsDate(vDOB) And vAssocMemberType.MaxJuniorAge > 0 Then
                    'Check whether DateOfBirth makes this an associate member
                    If DateAdd(Microsoft.VisualBasic.DateInterval.Year, vAssocMemberType.MaxJuniorAge, CDate(vDOB)) >= CDate(TodaysDate()) Then vGotAssociate = True
                  Else
                    'No DOB so allocate membership type depending on whether all main members have been added
                    'BR19011 changed condition from >0 to =0, so that CMT and New Member creation are the same.
                    If (vMemberCount >= pParams("NumberOfMembers").IntegerValue) And vAssocMemberType.MaxJuniorAge = 0 Then vGotAssociate = True
                  End If
                End If

                vAppliedDate = vMember.Applied
                vBranchMember = BooleanValue(vMember.BranchMember)
                vJoinedDate = vMember.Joined
                If (vMember.VotingRights = False AndAlso ((vGotAssociate = False AndAlso vMembershipType.VotingRights = True) OrElse (vGotAssociate = True AndAlso vAssocMemberType.VotingRights = True))) _
                OrElse ((.ProportionalBalanceSetting And (PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew)) > 0 And DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(TodaysDate()), CDate(.RenewalDate)) >= 0) Then
                  vJoinedDate = pParams("Joined").Value
                End If
                If vMember.VotingRights = False AndAlso ((vGotAssociate = False AndAlso vMembershipType.VotingRights = True) OrElse (vGotAssociate = True AndAlso vAssocMemberType.VotingRights = True)) Then
                  vAppliedDate = pParams("Joined").Value
                  vBranchMember = False
                Else
                  If IsDate(pParams.ParameterExists("Applied").Value) Then
                    vAppliedDate = pParams("Applied").Value
                  End If
                End If

                If vGotAssociate Then
                  vMember.SCAddMemberSummary(vMember.Contact, vJoinedDate, vMember.Branch, vBranchMember, vAppliedDate, pParams("DistributionCode").Value, vMember.AgeOverride, vMember.ContactDateOfBirth, (vAssocMemberType.MembershipTypeCode))
                Else
                  vMember.SCAddMemberSummary(vMember.Contact, vJoinedDate, vMember.Branch, vBranchMember, vAppliedDate, pParams("DistributionCode").Value, vMember.AgeOverride, vMember.ContactDateOfBirth, (vMembershipType.MembershipTypeCode))
                  vMemberCount = vMemberCount + 1
                End If
              Next vMember
            End With
          Else
            'Add current Member
            vContact = New Contact(mvEnv)
            vContact.Init(pParams("ContactNumber").IntegerValue)
            If Not vContact.Existing Then
              RaiseError(DataAccessErrors.daeParameterNotFound, "ContactNumber")
            End If

            vSQL = "SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress) & " FROM addresses a,"
            vSQL += If(vContact.ContactType <> Contact.ContactTypes.ctcOrganisation, " contact_addresses ca", " organisation_addresses ca, organisations o")
            vSQL += ", contacts c"
            vSQL = vSQL & " WHERE a.address_number = " & pParams("AddressNumber").IntegerValue & " AND ca.address_number = a.address_number %3 AND"
            vSQL += If(vContact.ContactType <> Contact.ContactTypes.ctcOrganisation, " ca.contact_number = c.contact_number", " ca.organisation_number = o.organisation_number AND o.organisation_number = c.contact_number")
            vSQL += " AND c.contact_number %1 " & pParams("ContactNumber").IntegerValue
            If pParams("GiftMembership").Bool Then
              vSQL = vSQL & " AND c.contact_number <> " & pParams("PayerContactNumber").IntegerValue
            End If
            vSQL = vSQL & " %2 ORDER BY date_of_birth" & mvEnv.Connection.DBSortByNullsFirst & ",surname, forenames"

            'Contact type and historic addresses are ignored for the first Member
            vRS = mvEnv.Connection.GetRecordSet(Replace(Replace(Replace(vSQL, "%3", ""), "%2", ""), "%1", "="))
            If vRS.Fetch() = True Then vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress)
            vRS.CloseRecordSet()
            If vContact.Existing = False Then RaiseError(DataAccessErrors.daeParameterNotFound, "ContactNumber")
            With pTransaction.PaymentPlan
              .PlanNumber = 0
              vGotAssociate = False
              vMemberCount = 0
              vDOB = vContact.DateOfBirth
              If vDOB.Length = 0 Then vDOB = pParams("DateOfBirth").Value
              If Len(vMembershipType.AssociateMembershipType) > 0 Then
                If IsDate(vDOB) And vAssocMemberType.MaxJuniorAge > 0 Then
                  'Check whether DateOfBirth makes this an associate member
                  If DateAdd(Microsoft.VisualBasic.DateInterval.Year, vAssocMemberType.MaxJuniorAge, CDate(vDOB)) >= CDate(TodaysDate()) Then vGotAssociate = True
                End If
              End If
              If vGotAssociate Then
                .AddMember((vContact.ContactNumber), (vContact.AddressNumber), (vAssocMemberType.MembershipTypeCode), vAssocMemberType, (vContact.ContactType))
                .Member.SCAddMemberSummary(vContact, pParams("Joined").Value, pParams("Branch").Value, vAssocMemberType.BranchMembership, pParams("Applied").Value, pParams("DistributionCode").Value, pParams("AgeOverride").Value, vDOB)
              Else
                vMemberCount = vMemberCount + 1
                .AddMember((vContact.ContactNumber), (vContact.AddressNumber), (vMembershipType.MembershipTypeCode), vMembershipType, (vContact.ContactType)) 'This adds Member to CurrentMembers collection
                .Member.SCAddMemberSummary(vContact, pParams("Joined").Value, pParams("Branch").Value, pParams("BranchMember").Bool, pParams("Applied").Value, pParams("DistributionCode").Value, pParams("AgeOverride").Value, vDOB)
              End If
              vCount = 1
            End With

            'Add all related Members
            'Contacts only and without historical addresses
            vRS = mvEnv.Connection.GetRecordSet(Replace(Replace(Replace(vSQL, "%3", "AND historical = 'N'"), "%2", "AND contact_type = 'C'"), "%1", "<>"))
            While vRS.Fetch() = True
              vCount = vCount + 1
              vGotAssociate = False
              vContact = New Contact(mvEnv)
              vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress)
              If Len(vMembershipType.AssociateMembershipType) > 0 Then
                If IsDate(vContact.DateOfBirth) And vAssocMemberType.MaxJuniorAge > 0 Then
                  'Check whether DateOfBirth makes this an associate member
                  If DateAdd(Microsoft.VisualBasic.DateInterval.Year, vAssocMemberType.MaxJuniorAge, CDate(vContact.DateOfBirth)) >= CDate(TodaysDate()) Then vGotAssociate = True
                Else
                  'No DOB so allocate membership type depending on whether all main members have been added
                  If (vMemberCount >= pParams("NumberOfMembers").IntegerValue) And vAssocMemberType.MaxJuniorAge = 0 Then vGotAssociate = True
                End If
              End If
              With pTransaction.PaymentPlan
                If vGotAssociate Then
                  .AddMember((vContact.ContactNumber), (vContact.AddressNumber), (vAssocMemberType.MembershipTypeCode), vAssocMemberType, (vContact.ContactType))
                  .Member.SCAddMemberSummary(vContact, pParams("Joined").Value, pParams("Branch").Value, vAssocMemberType.BranchMembership, pParams("Applied").Value, pParams("DistributionCode").Value, "", vContact.DateOfBirth)
                Else
                  .AddMember((vContact.ContactNumber), (vContact.AddressNumber), (vMembershipType.MembershipTypeCode), vMembershipType, (vContact.ContactType))
                  .Member.SCAddMemberSummary(vContact, pParams("Joined").Value, pParams("Branch").Value, pParams("BranchMember").Bool, pParams("Applied").Value, pParams("DistributionCode").Value, "", vContact.DateOfBirth)
                  vMemberCount = vMemberCount + 1
                End If
              End With
            End While
            vRS.CloseRecordSet()
            pResults.Add("CurrentMembers", vCount)
          End If

        Case TraderPage.TraderPageType.tpAmendMembership
          If pTransaction.SummaryMembers.Count >= 1 Then
            vMember = CType(pTransaction.SummaryMembers(1), Member)
            With vMember
              If .MembershipNumber > 0 Then pResults.Add("MembershipNumber", .MembershipNumber)
              pResults.Add("MembershipType", CDBField.FieldTypes.cftCharacter, .MembershipTypeCode)
              pResults.Add("Joined", CDBField.FieldTypes.cftDate, .Joined)
              pResults.Add("ContactNumber", .ContactNumber)
              pResults.Add("AddressNumber", .Contact.Address.AddressNumber)
              pResults.Add("Branch", CDBField.FieldTypes.cftLong, .Branch)
              pResults.Add("BranchMember", CDBField.FieldTypes.cftCharacter, .BranchMember)
              pResults.Add("Applied", CDBField.FieldTypes.cftDate, .Applied)
              pResults.Add("DobEstimated", CDBField.FieldTypes.cftCharacter, BooleanString(.ContactDOBEstimated))
              pResults.Add("AgeOverride", CDBField.FieldTypes.cftCharacter, .AgeOverride)
              pResults.Add("DateOfBirth", CDBField.FieldTypes.cftDate, .ContactDateOfBirth)
              pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, .DistributionCode)
            End With
          End If

        Case TraderPage.TraderPageType.tpPayments
          vBalance = pParams("TransactionAmount").DoubleValue
          '"member_number"
          If ((pParams("TransactionLines").DoubleValue = 0 Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atMove)) And Len(pParams.ParameterExists("MemberNumber").Value) > 0 Then pResults.Add("MemberNumber", CDBField.FieldTypes.cftCharacter, pParams("MemberNumber").Value)
          '"order_number"
          If ((pParams("TransactionLines").DoubleValue = 0 Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atMove)) And pParams.ParameterExists("PaymentPlanNumber").DoubleValue > 0 Then pResults.Add("PaymentPlanNumber", CDBField.FieldTypes.cftLong, pParams("PaymentPlanNumber").Value)
          '"covenant_number"
          If ((pParams("TransactionLines").DoubleValue = 0 Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atMove)) And pParams.ParameterExists("CovenantNumber").DoubleValue <> 0 Then pResults.Add("CovenantNumber", CDBField.FieldTypes.cftLong, pParams("CovenantNumber").Value)
          '"amount"
          If ((pParams("TransactionLines").DoubleValue = 0 Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atMove)) Then
            pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))
            'If MultiCurrency() Then mvBaseCurrencyAmount = CalcCurrencyAmount(vBalance, True)
          End If
          '"sales_contact_number"
          If pParams.ParameterExists("SalesContactNumber").IntegerValue > 0 Then
            pResults.Add("SalesContactNumber", CDBField.FieldTypes.cftLong, pParams("SalesContactNumber").Value)
          End If

        Case TraderPage.TraderPageType.tpPaymentPlanProducts, TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
          GetSegmentProductRate(pParams.ParameterExists("TransactionSource").Value, pParams("TransactionType").Value, pParams("TransactionPaymentMethod").Value, pPageType, Batch.AdjustmentTypes.atNone)
          If pParams("PPDLines").IntegerValue = 0 Or Me.DefaultProductUsesProductNumbers Then
            pResults.Add("Product", CDBField.FieldTypes.cftCharacter, DefaultProductCode)
            pResults.Add("Rate", CDBField.FieldTypes.cftCharacter, DefaultRateCode)
            'If we have set a product and rate and the result is a price of zero (donation)
            'Then set the amount of the donation to the amount of the transaction
            If IsDefaultProductAndRate And CurrentPrice = 0 Then
              pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(pParams("PPBalance").DoubleValue))
              pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(pParams("PPBalance").DoubleValue))
            End If
          End If
          pResults.Add("Arrears", CDBField.FieldTypes.cftNumeric, "0.00")
          pResults.Add("Quantity", CDBField.FieldTypes.cftInteger, "1")

          'distribution_code
          pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, pParams.ParameterExists("TransactionDistributionCode").Value)
          pResults.Add("Source", CDBField.FieldTypes.cftCharacter, pParams.ParameterExists("TransactionSource").Value)
          pResults.Add("ContactNumber", CDBField.FieldTypes.cftLong, pParams.ParameterExists("PayerContactNumber").Value)
          pResults.Add("AddressNumber", CDBField.FieldTypes.cftLong, pParams.ParameterExists("PayerAddressNumber").Value)
          pResults.Add("TimeStatus", CDBField.FieldTypes.cftCharacter, "C")
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPPDetailsEffectiveDate) Then
            If pTransaction.PaymentPlan.Existing = False Then pTransaction.PaymentPlan.Init(mvEnv, (pParams.ParameterExists("PaymentPlanNumber").IntegerValue))
            pResults.Add("EffectiveDate", CDBField.FieldTypes.cftDate, Me.SetPPDEffectiveDate(pParams.ParameterExists("PPPaymentType").Value, pParams.ParameterExists("NewPaymentFrequency").Value, pTransaction.PaymentPlan, True))
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then
              pResults.Add("ValidFrom", CDBField.FieldTypes.cftDate, pTransaction.PaymentPlan.StartDate)
              pResults.Add("ValidTo", CDBField.FieldTypes.cftDate, pTransaction.PaymentPlan.ExpiryDate)
            End If
          End If
          '
        Case TraderPage.TraderPageType.tpPaymentPlanMaintenance
          If pParams("PaymentPlanNumber").IntegerValue > 0 Then
            pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
            With pTransaction.PaymentPlan
              '.PlanNumber = mvTRDOrderNumber  Now populated from the tpContactSelection page
              If .Existing Then
                '"their_reference"
                pResults.Add("TheirReference", CDBField.FieldTypes.cftCharacter, .TheirReference)
                '"balance"
                pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(.Balance))
                '"reason_for_despatch"
                pResults.Add("ReasonForDespatch", CDBField.FieldTypes.cftCharacter, .ReasonForDespatch)
                '"source"
                pResults.Add("Source", CDBField.FieldTypes.cftCharacter, .Source)
                '"payment_frequency"
                pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, .PaymentFrequencyCode)
                '"payment_method"
                pResults.Add("PaymentMethod", CDBField.FieldTypes.cftCharacter, .PaymentMethod)
                '"giver_contact_number"
                pResults.Add("GiverContactNumber", CDBField.FieldTypes.cftCharacter, .GiverContactNumber) 'Leave this as cftCharacter as most cases it will be null(?)
                '"amount"
                pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, .Amount)
                '"renewal_date"
                pResults.Add("RenewalDate", CDBField.FieldTypes.cftDate, .RenewalDate)
                '"first_amount"
                pResults.Add("FirstAmount", CDBField.FieldTypes.cftCharacter, .FirstAmount) 'Leave this as cftCharacter as most cases it will ne null(?)
                pResults.Add("FirstAmountVisible", CDBField.FieldTypes.cftCharacter, "N")
                If (.PlanType = CDBEnvironment.ppType.pptMember And .FixedRenewalCycle And .PreviousRenewalCycle And .StartDate = .RenewalDate) Then
                  If (.ProportionalBalanceSetting And PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew) = PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew Then pResults("FirstAmountVisible").Value = "Y"
                End If
                'add the trans type so it ca be used for setting the Eligible for gift aid
                If pParams.Exists("TransactionType") Then
                  pParams("TransactionType").Value = .TransactionType
                Else
                  pParams.Add("TransactionType", CDBField.FieldTypes.cftCharacter, .TransactionType)
                End If

                '"eligible_for_gift_aid"

                If .EligibleForGiftAid Then
                  pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, "Y")
                Else
                  If .PlanType = CDBEnvironment.ppType.pptMember And .MembershipEligibleForGiftAid(.StartDate) = False Then
                    pResults.Add("EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, "D")
                  End If
                End If
                '"pack_to_donor"
                pResults.Add("PackToMember", CDBField.FieldTypes.cftCharacter, BooleanString(.PackToMember))
                pResults.Add("RenewalAmount", CDBField.FieldTypes.cftNumeric, FixedFormat(.RenewalAmount))
                pResults.Add("TransactionType", CDBField.FieldTypes.cftCharacter, .TransactionType)
                pResults.Add("GiftMembership", BooleanString(.GiftMembership))
                pResults.Add("OneYearGift", BooleanString(.OneYearGift))
                If .GiftMembership = False OrElse .PlanType <> CDBEnvironment.ppType.pptMember OrElse .MembershipType.MembersPerOrder = 0 Then
                  pResults.Add("OneYearGiftEnabled", "N")
                End If
              Else
              End If
            End With
            pResults.Add("WriteOffMissedPayments", BooleanString(mvEnv.GetConfigOption("fp_pp_wo_missed_payments", False)))
          End If

        Case TraderPage.TraderPageType.tpPaymentPlanSummary
          If pParams("CurrentPageType").IntegerValue = TraderPage.TraderPageType.tpPaymentPlanMaintenance Then
            If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
              pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
              If pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptMember Then
                'This will set the DetailType for all detail lines
                pTransaction.PaymentPlan.SetDetailLineTypesForSC()
              End If
              For Each vPPD In pTransaction.PaymentPlan.Details
                vPPD.LineNumber = vPPD.DetailNumber
                pTransaction.TraderPPDLines.AddItem(vPPD, CStr(vPPD.DetailNumber))
              Next vPPD
              pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, FixedFormat(pTransaction.PaymentPlan.Balance))
            End If
          ElseIf pParams("CurrentPageType").IntegerValue = TraderPage.TraderPageType.tpLoans Then
            Dim vLoanBalance As Double
            If pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
              'Existing Loan, so return current PPD lines
              pTransaction.PaymentPlan.Init(mvEnv, pParams("PaymentPlanNumber").IntegerValue)
              For Each vPPD In pTransaction.PaymentPlan.Details
                vPPD.LineNumber = vPPD.DetailNumber
                pTransaction.TraderPPDLines.AddItem(vPPD, vPPD.DetailNumber.ToString)
              Next
              Dim vDiff As Double = (pParams("LoanAmount").DoubleValue - pTransaction.PaymentPlan.Loan.LoanAmount)
              vLoanBalance = FixTwoPlaces(pTransaction.PaymentPlan.Balance + vDiff)
            Else
              'New Loan, so create the capital & Interest PPD lines
              Dim vLoanType As New LoanType(mvEnv)
              vLoanType.Init(pParams("LoanType").Value)
              If vLoanType.Existing Then
                With pParams
                  If .Exists("ContactNumber") = False Then
                    .Add("ContactNumber", .Item("PayerContactNumber").IntegerValue)
                    .Add("AddressNumber", .Item("PayerAddressNumber").IntegerValue)
                  End If
                  If .ParameterExists("SalesContactNumber").IntegerValue > 0 Then
                    'If we have the SalesContactNumber, then set this against the Loan Capital
                    vContact = New Contact(mvEnv)
                    vContact.Init(pParams("SalesContactNumber").IntegerValue)
                    .Item("ContactNumber").Value = vContact.ContactNumber.ToString
                    .Item("AddressNumber").Value = vContact.AddressNumber.ToString
                  End If
                  .Add("Product", vLoanType.CapitalProduct.ProductCode)
                  .Add("Rate", vLoanType.CapitalProduct.ProductRate.RateCode)
                  .Add("Balance", pParams("LoanAmount").DoubleValue)
                  .Add("DetailFixedAmount", pParams("LoanAmount").DoubleValue)
                End With
                Dim vPPDLine As PaymentPlanDetail = pTransaction.GetPaymentPlanDetail(1)
                vPPDLine.CreateSC(pParams)
                With pParams
                  If .ParameterExists("SalesContactNumber").IntegerValue > 0 Then
                    'For the Loan Interest, always use the payer Contact
                    .Item("ContactNumber").Value = .Item("PayerContactNumber").Value
                    .Item("AddressNumber").Value = .Item("PayerContactNumber").Value
                  End If
                  .Item("Product").Value = vLoanType.InterestProduct.ProductCode
                  .Item("Rate").Value = vLoanType.InterestProduct.ProductRate.RateCode
                  .Item("Balance").Value = "0"
                  .Item("DetailFixedAmount").Value = "0"
                End With
                vPPDLine = pTransaction.GetPaymentPlanDetail(2)
                vPPDLine.CreateSC(pParams)
                vLoanBalance = pParams("LoanAmount").DoubleValue
              End If
            End If
            pResults.Add("Balance", vLoanBalance)
          End If

        Case TraderPage.TraderPageType.tpMembershipPayer
          'This only comes from CMT
          pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
          vMembershipType = mvEnv.MembershipType((pParams("CMT_MembershipType").Value))

          If pParams("CMT_GiftMembership").Bool = False And pTransaction.PaymentPlan.GiftMembership = True Then
            'User has removed Gift Flag from Membership.
            'Default payer to first Main Member/Joint between Members if Main Type is Joint
            If vMembershipType.MembersPerOrder = 2 And pParams("CMT_NumberOfMembers").IntegerValue = 2 Then
              vContact = pTransaction.GetMembershipJointContact(vMembershipType.MembershipTypeCode, pParams("CMT_Source").Value)
            Else
              'Set to first Main Member
              For Each vMember In pTransaction.SummaryMembers
                If vMember.MembershipTypeCode = vMembershipType.MembershipTypeCode Then
                  If vContact Is Nothing Then
                    vContact = vMember.Contact
                  End If
                End If
                If Not (vContact Is Nothing) Then Exit For
              Next vMember
            End If
          End If

          If vContact Is Nothing Then
            vContact = New Contact(mvEnv)
            vContact.Init()
            If vMembershipType.PayerRequired = "M" Then
              'Set payer to be the AffiliatedMember
              'PayerContactNumber will have already been set to the AffiliatedMembers ContactNumber
              vContact.Init((pParams("PayerContactNumber").IntegerValue), (pParams("PayerAddressNumber").IntegerValue))
            ElseIf (pParams("CMT_GiftMembership").Bool = False And pParams("CMT_NumberOfMembers").IntegerValue = 1 And pParams("CMT_MaxFreeAssociates").IntegerValue = 0) Then
              'Changing to a single membership force the payer to be the member unless a Gifted
              If pTransaction.SummaryMembers.Count > 0 Then
                vMember = CType(pTransaction.SummaryMembers(1), Member)
                vContact = vMember.Contact
              End If
            ElseIf pParams("CMT_GiftMembership").Bool = False Then
              'Set to first Main Member
              For Each vMember In pTransaction.SummaryMembers
                If vMember.MembershipTypeCode = vMembershipType.MembershipTypeCode Then
                  If vContact.Existing = False Then vContact = vMember.Contact
                End If
                If vContact.Existing Then Exit For
              Next vMember
            Else
              'Set payer to be original PaymentPlan payer
              vContact.Init(pTransaction.PaymentPlan.ContactNumber, pTransaction.PaymentPlan.AddressNumber)
            End If
          End If
          If vContact.Existing = False Then
            'Set payer to be original PaymentPlan payer
            vContact.Init(pTransaction.PaymentPlan.ContactNumber, pTransaction.PaymentPlan.AddressNumber)
          End If

          pResults.Add("ContactNumber", vContact.ContactNumber)
          pResults.Add("AddressNumber", vContact.Address.AddressNumber)

          'Check whether the GiftMembership flag needs to be reset
          pResults.Add("ForceGiftMembership", CDBField.FieldTypes.cftCharacter, "N")
          For Each vMember In pTransaction.SummaryMembers
            If vMember.ContactNumber = vContact.ContactNumber Then Exit For
          Next vMember
          If vMember Is Nothing Then
            vForceGifted = Not (pParams("CMT_GiftMembership").Bool)
            If vMembershipType.MembersPerOrder = 2 Then
              For Each vMember In pTransaction.SummaryMembers
                If vMember.MembershipTypeCode = vMembershipType.MembershipTypeCode Then
                  If vContact1 Is Nothing Then
                    vContact1 = vMember.Contact
                  Else
                    vContact2 = vMember.Contact
                  End If
                End If
              Next vMember
              If (Not (vContact1 Is Nothing)) And (Not (vContact2 Is Nothing)) Then
                vContact = vContact1.GetJointContact(vContact2)
                If pResults("ContactNumber").IntegerValue = vContact.ContactNumber Then vForceGifted = False
              End If
            End If
            If vForceGifted Then pResults("ForceGiftMembership").Value = "Y"
          End If

        Case TraderPage.TraderPageType.tpOutstandingScheduledPayments
          vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)

          'We should always have a payment plan number coming back from the client
          pTransaction.PaymentPlan.Init(mvEnv, pParams("PaymentPlanNumber").IntegerValue)

          'First check for available scheduled payments
          vLineNumber = 0
          If pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan Then
            'Just get the OPS we require
            pTransaction.PaymentPlan.GetLoanPaymentOPS(Date.Parse(pParams("TransactionDate").Value))
            vOPSInclude = (pTransaction.PaymentPlan.ScheduledPayments.Count > 0)
          Else
            For Each vOPS In pTransaction.PaymentPlan.ScheduledPayments
              If vOPS.AmountOutstanding > 0 Then vOPSInclude = True
              If vOPSInclude Then Exit For
            Next vOPS
          End If

          'slightly different from the rich client.
          If vOPSInclude = False Then
            'No scheduled payments so add in-advance record
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv)

            vOPS.CreateInAdvance(mvEnv, pTransaction.PaymentPlan, pParams("Amount").DoubleValue, False)
            If vOPS.AmountOutstanding > 0 Then pTransaction.OutstandingOPS.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
          Else
            Dim vLoan As Boolean = (pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan)
            For Each vOPS In pTransaction.PaymentPlan.ScheduledPayments
              vOPSInclude = False
              If vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then
                'Only include the provisional line if this is the first line
                If vLineNumber = 0 And vOPS.AmountOutstanding > 0 Then vOPSInclude = True
              Else
                Select Case vOPS.ScheduledPaymentStatus
                  Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid
                    vOPSInclude = True
                  Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
                    If vOPS.AmountOutstanding > 0 OrElse vLoan = True Then vOPSInclude = True
                  Case Else
                    If vLoan Then vOPSInclude = True
                End Select
              End If
              If vOPSInclude Then
                vLineNumber = vLineNumber + 1
                pTransaction.OutstandingOPS.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
              End If
            Next vOPS
          End If

          If pTransaction.OutstandingOPS.Count() = 0 Then RaiseError(DataAccessErrors.daeNoOPSForPP, (pParams("PaymentPlanNumber").Value)) 'There are no outstanding scheduled payments for Payment Plan %s

          'Allocate any defaults
          vAmount = pParams("Amount").DoubleValue
          vBalance = vAmount
          If vFinancialAdjustment = Batch.AdjustmentTypes.atNone And pTransaction.OutstandingOPS.Count() > 0 Then
            If pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan Then
              'Always allocate to the current months payment
              Dim vGotOPS As Boolean = False
              Dim vOPSEndDate As Date
              Dim vTransDate As Date = Date.Parse(pParams("TransactionDate").Value)
              For Each vOPS In pTransaction.OutstandingOPS
                If pTransaction.PaymentPlan.TermUnits = PaymentPlan.OrderTermUnits.otuWeekly OrElse pTransaction.PaymentPlan.PaymentFrequencyPeriod = PaymentFrequency.PaymentFrequencyPeriods.pfpDays Then
                  vOPSEndDate = CDate(vOPS.DueDate).AddDays(pTransaction.PaymentPlan.PaymentFrequencyInterval).AddDays(-1)
                Else
                  vOPSEndDate = CDate(AddMonths(pTransaction.PaymentPlan.StartDate, vOPS.DueDate, pTransaction.PaymentPlan.PaymentFrequencyInterval)).AddDays(-1)
                End If
                If CDate(vOPS.DueDate) <= vTransDate AndAlso vOPSEndDate >= vTransDate Then vGotOPS = True
                If vGotOPS Then
                  vOPS.AddPayment(vBalance)
                  vOPS.SCCheckValue = True
                  vBalance = 0
                  Exit For
                End If
              Next
              If vGotOPS = False And pTransaction.OutstandingOPS.Count = 1 Then
                vOPS = DirectCast(pTransaction.OutstandingOPS(1), OrderPaymentSchedule)
                If Date.Parse(vOPS.DueDate) <= vTransDate AndAlso Date.Parse(vOPS.DueDate).AddMonths(1) >= vTransDate Then
                  vGotOPS = True
                ElseIf pTransaction.PaymentPlan.Loan.FixedMonthlyAmount = 0 AndAlso pTransaction.PaymentPlan.Loan.LoanTerm = 0 Then
                  vGotOPS = True
                End If
                If vGotOPS Then
                  vOPS.AddPayment(vBalance)
                  vOPS.SCCheckValue = True
                  vBalance = 0
                End If
              End If
              If vGotOPS = False Then RaiseError(DataAccessErrors.daeCannotFindLoanPayment, vTransDate.ToString(CAREDateFormat))
            Else
              If vAmount = pTransaction.PaymentPlan.Balance Then
                'Allocate against all rows
                For Each vOPS In pTransaction.OutstandingOPS
                  PayScheduledPayment(vOPS, vBalance)
                Next vOPS
              Else
                'Allocate against first row only, if possible
                vOPS = CType(pTransaction.OutstandingOPS.Item(1), OrderPaymentSchedule)
                If pTransaction.OutstandingOPS.Count() = 1 And vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then
                  'Only 1 row and this is the provisional record
                  PayScheduledPayment(vOPS, vBalance)
                ElseIf vOPS.AmountOutstanding = vAmount Then
                  PayScheduledPayment(vOPS, vBalance)
                End If
              End If
            End If
            'If only provisional OPS and still an amount unallocated add more provisional OPS
            If vBalance > 0 And pTransaction.OutstandingOPS.Count() = 1 And vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then
              While vBalance > 0
                vOPS = New OrderPaymentSchedule
                vOPS.Init(mvEnv)
                vOPS.CreateInAdvance(mvEnv, pTransaction.PaymentPlan, pParams("Amount").DoubleValue, False)
                pTransaction.OutstandingOPS.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
                PayScheduledPayment(vOPS, vBalance)
              End While
            End If
          End If
          If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
            'Try and find original payment
            If pParams.ParameterExists("ScheduledPaymentNumber").IntegerValue > 0 Then
              For Each vOPS In pTransaction.OutstandingOPS
                If vOPS.ScheduledPaymentNumber = pParams("ScheduledPaymentNumber").IntegerValue Then
                  If pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment Then
                    'The amount outstanding needs to be increased by the payment amount
                    vOPS.ProcessReanalysis(FixTwoPlaces(vAmount * -1))
                  End If
                  PayScheduledPayment(vOPS, vBalance)
                  Exit For
                End If
              Next vOPS
            End If
            If pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan Then
              'Always allocate to the current months payment
              Dim vGotOPS As Boolean = False
              Dim vPrevOPS As OrderPaymentSchedule = Nothing
              Dim vTransDate As Date = Date.Parse(pParams("TransactionDate").Value)
              For Each vOPS In pTransaction.OutstandingOPS
                If Date.Parse(vOPS.DueDate) >= vTransDate AndAlso (vPrevOPS IsNot Nothing AndAlso Date.Parse(vPrevOPS.DueDate) <= vTransDate) Then vGotOPS = True
                If vGotOPS Then
                  vPrevOPS.AddPayment(vBalance)
                  vPrevOPS.SCCheckValue = True
                  vBalance = 0
                  Exit For
                End If
                vPrevOPS = vOPS
              Next
              If vGotOPS = False And pTransaction.OutstandingOPS.Count = 1 Then
                vOPS = DirectCast(pTransaction.OutstandingOPS(1), OrderPaymentSchedule)
                If Date.Parse(vOPS.DueDate) <= vTransDate AndAlso Date.Parse(vOPS.DueDate).AddMonths(1) >= vTransDate Then
                  vGotOPS = True
                ElseIf pTransaction.PaymentPlan.Loan.FixedMonthlyAmount = 0 AndAlso pTransaction.PaymentPlan.Loan.LoanTerm = 0 Then
                  vGotOPS = True
                End If
                If vGotOPS Then
                  vOPS.AddPayment(vBalance)
                  vOPS.SCCheckValue = True
                  vBalance = 0
                End If
              End If
              If vGotOPS = False Then RaiseError(DataAccessErrors.daeCannotFindLoanPayment, vTransDate.ToString(CAREDateFormat))
            End If
          End If

          pResults.Add("PaymentAmount", CDBField.FieldTypes.cftNumeric, FixedFormat(vAmount))
          pResults.Add("AmountOutstanding", CDBField.FieldTypes.cftNumeric, FixedFormat(vBalance))

        Case TraderPage.TraderPageType.tpPurchaseOrderDetails
          If pParams.ParameterExists("PurchaseOrderNumber").IntegerValue > 0 Then
            vPO = New PurchaseOrder(mvEnv)
            vPONumber = pParams("PurchaseOrderNumber").IntegerValue
            vPO.Init(vPONumber)
            vPO.InitDetails()
            For Each vPOD As PurchaseOrderDetail In vPO.Details
              pTransaction.GetPurchaseOrderDetail(vPOD.LineNumber, vPO.PurchaseOrderNumber)
            Next
            vPO.InitPayments()
            For Each vPOP As PurchaseOrderPayment In vPO.Payments
              pTransaction.GetPurchaseOrderPayment(vPOP.PaymentNumber, (vPO.PurchaseOrderNumber))
            Next
            pResults.Add("ContactNumber", vPO.ContactNumber)
            pResults.Add("AddressNumber", vPO.AddressNumber)
            pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, vPO.Amount.ToString)
            pResults.Add("Balance", CDBField.FieldTypes.cftNumeric, vPO.Balance.ToString)
            pResults.Add("OutputGroup", CDBField.FieldTypes.cftCharacter, vPO.OutputGroup)
            pResults.Add("PurchaseOrderType", CDBField.FieldTypes.cftCharacter, vPO.PurchaseOrderTypeCode)
            pResults.Add("PurchaseOrderDesc", CDBField.FieldTypes.cftCharacter, vPO.PurchaseOrderDesc)
            pResults.Add("PayeeContactNumber", vPO.PayeeContactNumber)
            pResults.Add("PayeeAddressNumber", vPO.PayeeAddressNumber)

            pResults.Add("StartDate", CDBField.FieldTypes.cftCharacter, vPO.StartDate)
            pResults.Add("NumberOfPayments", CDBField.FieldTypes.cftCharacter, vPO.NumberOfPayments.ToString)
            pResults.Add("DistributionMethod", CDBField.FieldTypes.cftCharacter, If(vPO.DistributionMethod = PurchaseOrder.PODistributionMethods.podmProportional, "P", "S"))
            pResults.Add("PaymentAsPercentage", CDBField.FieldTypes.cftCharacter, BooleanString(vPO.PaymentAsPercentage))
            pResults.Add("Campaign", CDBField.FieldTypes.cftCharacter, vPO.Campaign)
            pResults.Add("Appeal", CDBField.FieldTypes.cftCharacter, vPO.Appeal)
            pResults.Add("Segment", CDBField.FieldTypes.cftCharacter, vPO.Segment)
            pResults.Add("CurrencyCode", CDBField.FieldTypes.cftCharacter, vPO.CurrencyCode)

            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderManagement) Then
              pResults.Add("PaymentFrequency", CDBField.FieldTypes.cftCharacter, vPO.PaymentFrequency)
            End If

          End If

        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary
          If pTransaction.PurchaseInvoiceDetails.Count() = 0 Then
            If pParams.ParameterExists("PurchaseInvoiceNumber").IntegerValue > 0 Then
              vPI = New PurchaseInvoice(mvEnv)
              vPINumber = pParams("PurchaseInvoiceNumber").IntegerValue
              vPI.Init(vPINumber)
              vPI.InitDetails()
              For Each vPID As PurchaseInvoiceDetail In vPI.Details
                pTransaction.GetPurchaseInvoiceDetail(vPID.LineNumber, (vPI.PurchaseInvoiceNumber))
              Next
            Else
              vPO = New PurchaseOrder(mvEnv)
              vPONumber = pParams("PurchaseOrderNumber").IntegerValue
              vPO.Init(vPONumber)
              vPO.InitDetails()
              For Each vPOD As PurchaseOrderDetail In vPO.Details
                pTransaction.GetPurchaseInvoiceDetail(vPOD.LineNumber, 0, vPONumber)
              Next
            End If
          End If

        Case TraderPage.TraderPageType.tpConfirmProvisionalTransactions
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          If pParams.ParameterExists("ProductNumber").IntegerValue > 0 Then
            pResults.Add("ProductNumber", pParams("ProductNumber").IntegerValue)
          Else

          End If

        Case TraderPage.TraderPageType.tpSetStatus, TraderPage.TraderPageType.tpSuppressionEntry, TraderPage.TraderPageType.tpGiftAidDeclaration, TraderPage.TraderPageType.tpGoneAway, TraderPage.TraderPageType.tpAddressMaintenance
          'used in Smart Client
        Case TraderPage.TraderPageType.tpActivityEntry
          pResults.Add("DefaultSource", CDBField.FieldTypes.cftCharacter, mvEnv.Connection.GetValue("SELECT source FROM activity_groups WHERE activity_group = '" & DefaultActivityGroup & "'"))
        Case TraderPage.TraderPageType.tpLegacyBequestReceipt
          If pParams.HasValue("TransactionDate") Then
            pResults.Add("DateReceived", CDBField.FieldTypes.cftDate, pParams("TransactionDate").Value)
          Else
            pResults.Add("DateReceived", CDBField.FieldTypes.cftDate, TodaysDate)
          End If
          pResults.Add("Amount", CDBField.FieldTypes.cftNumeric, FixedFormat(pParams("TransactionAmount").DoubleValue))

        Case TraderPage.TraderPageType.tpGiveAsYouEarn
          pResults.Add("Source", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYESource))
          pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, TodaysDate)

        Case TraderPage.TraderPageType.tpGiveAsYouEarnEntry
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          pResults.Add("EmployerOrganisationNumber", CDBField.FieldTypes.cftCharacter, "")
          pResults.Add("Source", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYESource))
          pResults.Add("Product", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEDonorProduct))
          pResults.Add("Rate", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEDonorRate))
          pResults.Add("DistributionCode", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEDistributionCode))
          pResults.Add("StartDate", CDBField.FieldTypes.cftDate, TodaysDate)

        Case TraderPage.TraderPageType.tpPostTaxPGPayment
          pResults.Add("Source", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGPledgeSource))
          pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, TodaysDate)
          If mvEnv.GetConfigOption("option_post_batches_to_CB") Then pResults.Add("PostCashBook", CDBField.FieldTypes.cftCharacter, "Y")

        Case TraderPage.TraderPageType.tpCancelPaymentPlan
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("AddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("CancellationReason", DefaultCancellationReason)

        Case TraderPage.TraderPageType.tpCancelGiftAidDeclaration
          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("ContactNumber", pParams("PayerContactNumber").IntegerValue)
          pResults.Add("CancellationReason", DefaultCancellationReason)

        Case TraderPage.TraderPageType.tpPostageAndPacking
          If mvCarriageProduct.Length > 0 Then
            pResults.Add("CarriageProduct", mvCarriageProduct)
          Else
            pResults.Add("CarriageProduct", "")
          End If
          Dim vPrice As Double
          If mvCarriageRate.Length > 0 Then
            pResults.Add("CarriageRate", mvCarriageRate)
            Dim vProductRate As New ProductRate(mvEnv)
            vProductRate.Init(mvCarriageProduct, mvCarriageRate)
            If vProductRate.Existing Then vPrice = vProductRate.Price(0)
          Else
            pResults.Add("CarriageRate", "")
          End If
          pResults.Add("Percentage", FixTwoPlaces(mvCarriagePercentage))
          pResults.Add("CarriagePrice", vPrice)
          pResults.Add("TransactionAmount", FixTwoPlaces(SumLineTypes("SALE", pTransaction) + SumLineTypes("P", pTransaction)).ToString)
          If vPrice = 0 Then
            pResults.Add("PAPAmount", FixTwoPlaces((SumLineTypes("SALE", pTransaction) + SumLineTypes("P", pTransaction)) * (mvCarriagePercentage / 100)).ToString)
          Else
            pResults.Add("PAPAmount", FixTwoPlaces(vPrice))
          End If

        Case TraderPage.TraderPageType.tpServiceBooking
          Dim vStartDate As DateTime = DateTime.Now
          Dim vYear As Integer = vStartDate.Year
          Dim vMonth As Integer = vStartDate.Month
          Dim vDay As Integer = vStartDate.Day
          Dim vHour As Integer = vStartDate.Hour
          If vStartDate.Minute > 29 Then
            'Default to the next hour
            If vHour = 23 Then
              vStartDate = New DateTime(vYear, vMonth, vDay, 0, 0, 0).AddDays(1)
            Else
              vStartDate = New DateTime(vYear, vMonth, vDay, vHour + 1, 0, 0)
            End If
          Else
            vStartDate = New DateTime(vYear, vMonth, vDay, vHour, 0, 0)
          End If

          If pParams.ParameterExists("PayerContactNumber").IntegerValue > 0 Then pResults.Add("BookingContactNumber", pParams("PayerContactNumber").IntegerValue)
          If pParams.ParameterExists("PayerAddressNumber").IntegerValue > 0 Then pResults.Add("BookingAddressNumber", pParams("PayerAddressNumber").IntegerValue)
          pResults.Add("StartDate", vStartDate.ToString(CAREDateTimeFormat))
          pResults.Add("EndDate", vStartDate.AddDays(1).ToString(CAREDateTimeFormat))
          pResults.Add("Amount", String.Empty)
          pResults.Add("Discount", String.Empty)
          pResults.Add("GrossAmount", String.Empty)
          If pParams.Exists("TransactionSource") Then pResults.Add("Source", pParams("TransactionSource").Value)

        Case TraderPage.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          pResults.Add("Amount", pParams("TransactionAmount").DoubleValue)
          pResults.Add("LineTotal", pParams("DetailLineTotal").DoubleValue)

        Case TraderPage.TraderPageType.tpBatchInvoiceSummary
          pResults.Add("PrintPreview", IIf(mvInvoicePrintPreviewDefault = True, "Y", "N").ToString)

        Case TraderPage.TraderPageType.tpLoans
          Dim vLoan As New Loan(mvEnv)
          Dim vPP As New PaymentPlan()
          vPP.Init(mvEnv)
          If pParams.ParameterExists("LoanNumber").IntegerValue > 0 Then
            vLoan.Init(pParams("LoanNumber").IntegerValue)
            vPP.Init(mvEnv, vLoan.PaymentPlanNumber)
          ElseIf pParams.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
            vPP.Init(mvEnv, pParams("PaymentPlanNumber").IntegerValue)
            vLoan.InitFromPaymentPlan(vPP.PlanNumber)
          End If
          With pResults
            .Add("OrderDate", CDBField.FieldTypes.cftDate, If(vLoan.Existing, vPP.StartDate, pParams.OptionalValue("TransactionDate", TodaysDate())))
            .Add("Source", If(vLoan.Existing, vLoan.Source, pParams.ParameterExists("TransactionSource").Value))
            If vLoan.Existing Then
              .Add("LoanType", vLoan.LoanTypeCode)
              .Add("LoanAmount", vLoan.LoanAmount.ToString("F"))
              .Add("PaymentFrequency", vPP.PaymentFrequencyCode)
              .Add("InterestRate", vLoan.InterestRate.ToString("F"))
              If vLoan.LoanTerm > 0 Then
                .Add("LoanTerm", vLoan.LoanTerm)
              Else
                .Add("FixedMonthlyAmount", vLoan.FixedMonthlyAmount.ToString("F"))
              End If
              .Add("TheirReference", vPP.TheirReference)
              If vPP.SalesContact > 0 Then .Add("SalesContactNumber", vPP.SalesContact)
              If .Exists("TransactionType") = False Then .Add("TransactionType", "LOAN")
              .Add("Balance", vPP.Balance)
            End If
          End With

        Case TraderPage.TraderPageType.tpAdvancedCMT
          pTransaction.PaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
          pTransaction.PaymentPlan.SetDetailLineTypesForSC(True)
          pParams = pTransaction.PaymentPlan.CheckPayPlanParameterList(pParams)

          Dim vCMTDate As Date = CDate(pParams.OptionalValue("CMTDate", TodaysDate))
          Dim vJoined As Date = CDate(pParams("Joined").Value)
          pTransaction.PaymentPlan.SetupCMT(pParams, pTransaction, vCMTDate, vJoined, False)
      End Select
    End Sub

    Public Function SaveTransaction(ByVal pTDRTransaction As TraderTransaction, ByVal pParams As CDBParameters, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, ByVal pExistingTrans As Boolean, Optional ByVal pConfirmTransList As String = "", Optional ByVal pUseStockTransactionID As Boolean = False) As SaveTransactionStatus
      'Save Batch/BatchTransaction/BatchTransactionAnalisys
      Dim vBatchType As Batch.BatchTypes
      Dim vContact As Contact
      Dim vInsertFields As CDBFields
      Dim vProvisional As Batch.ProvisionalOrConfirmed
      Dim vTDRBankDetails As TraderBankDetails = Nothing
      Dim vTDRLine As TraderAnalysisLine
      Dim vReturnValue As SaveTransactionStatus
      Dim vAddressNumber As Integer 'Used to update CreditCustomer AddressNumber
      Dim vAmount As Double
      Dim vBankAccount As String = ""
      Dim vInvoiceIssued As Integer
      Dim vIssuedSet As Boolean
      Dim vIssuedValue As String
      Dim vOnOrder As Double
      Dim vOutstanding As Double
      Dim vPaymentMethod As String
      Dim vQuantity As Double
      Dim vQuantityValue As String
      Dim vReference As String = ""
      Dim vStockProducts As Boolean
      Dim vTermsFrom As String
      Dim vTermsNumber As String
      Dim vTermsPeriod As String
      Dim vTransAmount As Double
      Dim vNewDOB As Boolean
      Dim vConfirmTransList As String = ""
      Dim vFATransactionType As String = ""
      Dim vPayMethodCode As String = ""

      '-------------------------------------------------------------------------
      'First determine whether the transaction is provisional or confirmed
      '-------------------------------------------------------------------------
      vProvisional = Batch.ProvisionalOrConfirmed.Confirmed
      If pParams.Exists("Provisional") Then
        If pParams("Provisional").Bool Then vProvisional = Batch.ProvisionalOrConfirmed.Provisional
      End If
      If vProvisional <> Batch.ProvisionalOrConfirmed.Provisional Then
        vConfirmTransList = pParams.ParameterExists("ConfirmTransList").Value
      End If

      '-------------------------------------------------------------------------
      'Set the PaymentMethod, BankAccount, etc.
      '-------------------------------------------------------------------------
      If pParams.Exists("BKD_SortCode") = True And pParams.Exists("BKD_AccountNumber") = True And pParams.Exists("BKD_AccountName") = True Then
        If pParams.Exists("BankDetailsNumber") = False Then pParams.Add("BankDetailsNumber", CDBField.FieldTypes.cftLong)
        vTDRBankDetails = New TraderBankDetails
        vTDRBankDetails.Init(mvEnv, pExistingTrans, (pFinancialAdjustment <> Batch.AdjustmentTypes.atNone), pParams("PayerContactNumber").IntegerValue, pParams("BankDetailsNumber").IntegerValue, pParams("BKD_SortCode").Value, pParams("BKD_AccountNumber").Value, pParams("BKD_AccountName").Value, (pParams.ParameterExists("BKD_BranchName").Value), (pParams.ParameterExists("NewBank").Bool), mvBatchNumber, mvTransNumber)
        pParams("BankDetailsNumber").Value = vTDRBankDetails.BankDetailsNumber.ToString
      End If

      vPaymentMethod = pParams("TransactionPaymentMethod").Value
      Select Case pParams("TransactionPaymentMethod").Value
        Case "CASH", "SO"
          If mvAppType = ApplicationType.atBankStatementPosting Then
            vPaymentMethod = mvEnv.GetConfig("pm_sp")
          ElseIf mvAppType = ApplicationType.atCreditListReconciliation Then
            vPaymentMethod = mvEnv.GetConfig("pm_so")
            If pParams.Exists("BankPaymentMethod") AndAlso pParams("BankPaymentMethod").Value.Length > 0 Then
              vPayMethodCode = pParams("BankPaymentMethod").Value
            Else
              vPayMethodCode = mvEnv.GetConfig("pm_credit_list_reconciliation")
            End If
          End If
          vBankAccount = mvCABankAccount
          If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          Select Case pFinancialAdjustment
            Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              vBatchType = Batch.BatchTypes.Cash
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
        Case "POST"
          vBankAccount = mvCABankAccount
          If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          Select Case pFinancialAdjustment
            Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              vBatchType = Batch.BatchTypes.Cash
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
        Case "CHEQ", "CQIN"
          vBankAccount = mvCABankAccount
          If mvPayMethodsAtEnd = True And Not (vTDRBankDetails Is Nothing) Then
            If pParams.Exists("BKD_Reference") Then vReference = pParams("BKD_Reference").Value
          Else
            If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          End If
          Select Case pFinancialAdjustment
            Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              If pParams("TransactionPaymentMethod").Value = "CQIN" Then
                vBatchType = Batch.BatchTypes.CashWithInvoice
              Else
                vBatchType = Batch.BatchTypes.Cash
              End If
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
        Case "CRED"
          If pParams.Exists("TransactionType") = False Then pParams.Add("TransactionType")
          If pExistingTrans = True And Len(pParams("TransactionType").Value) = 0 Then
            'TransactionType not set - may have just clicked Finished without changing anything so see if it was a CreditNote
            If mvEnv.Connection.GetValue("SELECT record_type FROM invoices WHERE batch_number = " & BatchNumber & " AND transaction_number = " & TransactionNumber) = "N" Then
              pParams("TransactionType").Value = "CRDN"
            End If
          End If
          If pParams("TransactionType").Value = "CRDN" Then vPaymentMethod = pParams("TransactionType").Value
          vBankAccount = mvCSBankAccount
          If pParams.Exists("CCU_Reference") Then vReference = pParams("CCU_Reference").Value
          vBatchType = Batch.BatchTypes.CreditSales

        Case "CARD", "CCIN"
          If pParams("TransactionPaymentMethod").Value = "CCIN" Then
            vBankAccount = mvCCBankAccount
            vBatchType = Batch.BatchTypes.CreditCardWithInvoice
          ElseIf pParams.ParameterExists("CDC_CreditOrDebitCard").Value = "C" Then
            vPaymentMethod = "CCARD"
            vBankAccount = mvCCBankAccount
            vBatchType = Batch.BatchTypes.CreditCard
          ElseIf pFinancialAdjustment = Access.Batch.AdjustmentTypes.atAdjustment AndAlso
            pParams("BatchNumber").DoubleValue <> 0 AndAlso
            pParams("TransactionNumber").DoubleValue <> 0 Then
            If New SQLStatement(mvEnv.Connection,
                                "payment_method",
                                "batch_transactions",
                                New CDBFields({New CDBField("batch_number",
                                                            pParams("BatchNumber").Value),
                                               New CDBField("transaction_number",
                                                            pParams("TransactionNumber").Value)})).GetValue = mvEnv.GetConfig("pm_cc") Then
              vPaymentMethod = "CCARD"
              vBankAccount = mvCCBankAccount
              vBatchType = Batch.BatchTypes.CreditCard
            Else
              vPaymentMethod = "DCARD"
              vBankAccount = mvDCBankAccount
              vBatchType = Batch.BatchTypes.DebitCard
            End If
          Else
            vPaymentMethod = "DCARD"
            vBankAccount = mvDCBankAccount
            vBatchType = Batch.BatchTypes.DebitCard
          End If
          If pParams.Exists("CDC_Reference") Then vReference = pParams("CDC_Reference").Value
          Select Case pFinancialAdjustment
            Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
              'vBatchType already set
            Case Else
              vBatchType = Batch.BatchTypes.FinancialAdjustment
          End Select
        Case "VOUC"
          vBankAccount = mvCVBankAccount
          If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          If pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
            vBatchType = Batch.BatchTypes.CAFVouchers
          Else
            vBatchType = Batch.BatchTypes.FinancialAdjustment
          End If
        Case "GFIK"
          vBankAccount = mvCABankAccount
          If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          vBatchType = Batch.BatchTypes.GiftInKind
        Case "CAFC"
          vBankAccount = mvCVBankAccount
          If pParams.Exists("TRD_Reference") Then vReference = pParams("TRD_Reference").Value
          vBatchType = Batch.BatchTypes.CAFCards
        Case "SAOR"
          vBankAccount = mvCABankAccount
          vBatchType = Batch.BatchTypes.SaleOrReturn
      End Select

      If mvCollectionPayments Then
        For Each vTDRLine In pTDRTransaction.TraderAnalysisLines
          If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment Then
            'BankAccount to be set to the CollectionBankAccount (if CollectionPISNumber set)
            If vTDRLine.CollectionPisNumber > 0 Then vBankAccount = vTDRLine.CollectionBankAccount
          End If
          If (vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment And vTDRLine.CollectionPisNumber > 0) Then Exit For
        Next vTDRLine
      End If

      If mvMultiCurrency Then
        If Len(BatchCurrencyCode) = 0 And Len(DefaultCurrencyCode) > 0 Then
          BatchCurrencyCode = DefaultCurrencyCode
        End If
        If BatchCurrencyCode <> DefaultCurrencyCode Then
          If CurrencyBankAccountExists(BatchCurrencyCode & "-" & Access.Batch.GetBatchTypeCode(vBatchType)) Then
            vBankAccount = CStr(CurrencyBankAccounts.Item(BatchCurrencyCode & "-" & Access.Batch.GetBatchTypeCode(vBatchType)))
          End If
        End If
      End If

      Select Case pFinancialAdjustment
        Case Batch.AdjustmentTypes.atGIKConfirmation
          vBankAccount = Batch.BankAccount
        Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atCashBatchConfirmation
          'vBankAccount already set
        Case Else
          vBankAccount = pParams.ParameterExists("FABankAccount").Value
          If Len(vBankAccount) = 0 Then vBankAccount = mvEnv.Connection.GetValue("SELECT default_bank_account FROM batch_types WHERE batch_type = '" & Access.Batch.GetBatchTypeCode(vBatchType) & "'")
          If Len(vBankAccount) = 0 Then vBankAccount = Batch.BankAccount
      End Select
      If mvExistingAdjustmentTran Then
        If pParams.Exists("BatchCategory") = False Then
          If Batch IsNot Nothing AndAlso Not String.IsNullOrEmpty(Batch.BatchCategory) Then   ' Original Batch Category
            pParams.Add("BatchCategory", Batch.BatchCategory)
          Else
            If Not String.IsNullOrEmpty(BatchCategory) Then pParams.Add("BatchCategory", BatchCategory) 'Trader Batch Category
          End If
        End If
      Else
        If BatchCategory.Length > 0 Then
          'Only default the BatchCategory if it has not already been set
          If pParams.Exists("BatchCategory") = False Then pParams.Add("BatchCategory")
          If Len(pParams("BatchCategory").Value) = 0 Then pParams("BatchCategory").Value = BatchCategory
        End If
      End If
      If pFinancialAdjustment <> Batch.AdjustmentTypes.atNone Then
        vFATransactionType = pParams.ParameterExists("FATransactionType").Value
        If pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or pFinancialAdjustment = Batch.AdjustmentTypes.atMove Then
          If Len(vFATransactionType) = 0 Then vFATransactionType = GetAdjustTransType(Batch.AdjustmentTypes.atAdjustment)

          'Add any additional  required items
          With pParams
            Select Case pParams("TransactionPaymentMethod").Value
              Case "CARD", "CAFC"
                If pParams.Exists("CDC_CreditOrDebitCard") = False Then .Add("CDC_CreditOrDebitCard", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_CreditCardType") = False Then .Add("CDC_CreditCardType", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_CardNumber") = False Then .Add("CDC_CardNumber", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_IssueNumber") = False Then .Add("CDC_IssueNumber", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_ValidDate") = False Then .Add("CDC_ValidDate", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_ExpiryDate") = False Then .Add("CDC_ExpiryDate", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_Reference") = False Then .Add("CDC_Reference", CDBField.FieldTypes.cftCharacter, "")
                If pParams.Exists("CDC_AuthorisationCode") = False Then .Add("CDC_AuthorisationCode", CDBField.FieldTypes.cftCharacter, "")
                If OnlineCCAuthorisation And pParams.Exists("SecurityCode") = False Then .Add("SecurityCode", CDBField.FieldTypes.cftCharacter, "")
            End Select
          End With
        ElseIf pFinancialAdjustment = Access.Batch.AdjustmentTypes.atGIKConfirmation OrElse pFinancialAdjustment = Access.Batch.AdjustmentTypes.atCashBatchConfirmation _
        AndAlso (BatchNumber > 0 AndAlso Not pParams.Exists("OrigBatchNumber")) Then
          pParams.Add("OrigBatchNumber", BatchNumber)
          pParams.Add("OrigTransNumber", TransactionNumber)
        End If
      End If

      '-----------------------------------------------------------------------------
      'Check for any required parameters for questions user needs to answer
      '-----------------------------------------------------------------------------
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataLinkToCommunication) = True And pParams.ParameterExists("LinkCommNumber").IntegerValue > 0 Then
        If pParams.Exists("CreateCommLink") = False Then
          Select Case LinkToCommunication
            Case LinkToCommunicationTypes.ltcYes
              pParams.Add("CreateCommLink", CDBField.FieldTypes.cftCharacter, "Y")
            Case LinkToCommunicationTypes.ltcNo
              pParams.Add("CreateCommLink", CDBField.FieldTypes.cftCharacter, "N")
            Case LinkToCommunicationTypes.ltcAsk
              RaiseError(DataAccessErrors.daeTraderCreateCommLink)
          End Select
        End If
      End If

      If pParams("TransactionPaymentMethod").Value = "CRED" OrElse pParams("TransactionPaymentMethod").Value = "CQIN" OrElse pParams("TransactionPaymentMethod").Value = "CCIN" Then
        If pTDRTransaction.CreditCustomerDetailsChanged Then
          With pTDRTransaction.CreditCustomer
            If (((.TermsNumber <> pParams("CCU_TermsNumber").Value) Or (.TermsFrom <> pParams("CCU_TermsFrom").Value) Or (.TermsPeriod <> pParams("CCU_TermsPeriod").Value)) And pParams.Exists("StorePaymentTerms") = False) Then
              RaiseError(DataAccessErrors.daeTraderStorePaymentTerms)
            End If
          End With

          If ((pTDRTransaction.CreditCustomer.AddressNumber <> pParams("CCU_AddressNumber").IntegerValue) And pParams.Exists("UpdateCreditCustomerAddress") = False) Then
            RaiseError(DataAccessErrors.daeTraderUpdateCreditCustomerAddress)
          End If
        End If
        'If we get here and these two parameters are still not set then they are not required so just set to default values of 'N'
        If pParams.Exists("StorePaymentTerms") = False Then pParams.Add("StorePaymentTerms", CDBField.FieldTypes.cftCharacter, "N")
        If pParams.Exists("UpdateCreditCustomerAddress") = False Then pParams.Add("UpdateCreditCustomerAddress", CDBField.FieldTypes.cftCharacter, "N")
      End If

      '-------------------------------------------------------------------------
      'Add the transaction
      '-------------------------------------------------------------------------
      If pParams.Exists("TRD_TransactionOrigin") And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataTransactionOrigins) Then pTDRTransaction.SetTransactionOrigin((pParams("TRD_TransactionOrigin").Value)) 'Setting TransactionOrigin here as unable to add to the init methods and the inits need for this to be set.

      Select Case pParams("TransactionPaymentMethod").Value
        Case "CASH", "CHEQ", "POST", "VOUC", "GFIK", "SAOR", "CQIN", "SO"
          If pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
            vStockProducts = False
            If pParams("TransactionPaymentMethod").Value = "CQIN" Then
              For Each vTDRLine In pTDRTransaction.TraderAnalysisLines
                If vTDRLine.StockSale Then vStockProducts = True
                If vStockProducts Then Exit For
              Next vTDRLine
            End If
            With pParams
              pTDRTransaction.InitCash(vPaymentMethod, vBankAccount, .Item("PayerContactNumber").IntegerValue, .Item("PayerAddressNumber").IntegerValue, .Item("TransactionDate").Value, .Item("TRD_Receipt").Value, (.Item("TRD_EligibleForGiftAid").Bool), pExistingTrans, (vProvisional = Batch.ProvisionalOrConfirmed.Provisional), BatchCurrencyCode, CStr(BatchExchangeRate), (.ParameterExists("TRD_Mailing").Value), (.ParameterExists("TRD_MailingContactNumber").Value), (.ParameterExists("TRD_MailingAddressNumber").Value), (.ParameterExists("BatchCategory").Value), vReference, (.ParameterExists("COM_Notes").Value), (.ParameterExists("BankDetailsNumber").IntegerValue), (.ParameterExists("TRD_AdditionalReference1").Value), (.ParameterExists("TRD_AdditionalReference2").Value), (.ParameterExists("TRD_AdditionalReference1Caption").Value), (.ParameterExists("TRD_AdditionalReference2Caption").Value), BatchNumber, TransactionNumber, .ParameterExists("CCU_AddressTo").Value, .ParameterExists("CCU_SalesLedgerAccount").Value, vStockProducts, vPayMethodCode)
            End With
          Else
            With pParams
              pTDRTransaction.InitFACash(vPaymentMethod, vBankAccount, .Item("PayerContactNumber").IntegerValue, .Item("PayerAddressNumber").IntegerValue, .ParameterExists("BatchDate").Value, .Item("TransactionDate").Value, .Item("TRD_Receipt").Value, (.Item("TRD_EligibleForGiftAid").Bool), pExistingTrans, pFinancialAdjustment, BatchCurrencyCode, CStr(BatchExchangeRate), (.ParameterExists("Mailing").Value), (.ParameterExists("MailingContactNumber").Value), (.ParameterExists("MailingAddressNumber").Value), (.ParameterExists("BatchCategory").Value), vReference, (.ParameterExists("COM_Notes").Value), vFATransactionType, (.ParameterExists("OriginalPaymentMethod").Value), (.ParameterExists("TRD_AdditionalReference1").Value), (.ParameterExists("TRD_AdditionalReference2").Value), (.ParameterExists("BankDetailsNumber").IntegerValue), BatchNumber, TransactionNumber, Me.AppType, (.ParameterExists("PostToCashBook").Value))
            End With
          End If

        Case "CARD", "CAFC", "CCIN"
          If pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
            vStockProducts = False
            If pParams("TransactionPaymentMethod").Value = "CCIN" Then
              For Each vTDRLine In pTDRTransaction.TraderAnalysisLines
                If vTDRLine.StockSale Then vStockProducts = True
                If vStockProducts Then Exit For
              Next vTDRLine
            End If
            With pParams
              pTDRTransaction.InitCardSale(vPaymentMethod, vBankAccount, .Item("PayerContactNumber").IntegerValue, .Item("PayerAddressNumber").IntegerValue, .Item("TransactionDate").Value, .Item("TRD_Receipt").Value, .Item("TRD_EligibleForGiftAid").Bool, pExistingTrans, .Item("CDC_CreditCardType").Value, .Item("CDC_CardNumber").Value, .Item("CDC_IssueNumber").Value, .ParameterExists("CDC_ValidDate").Value, .Item("CDC_ExpiryDate").Value, .ParameterExists("CDC_AuthorisationCode").Value, (vProvisional = Batch.ProvisionalOrConfirmed.Provisional), BatchCurrencyCode, CStr(BatchExchangeRate), (.ParameterExists("TRD_Mailing").Value), (.ParameterExists("TRD_MailingContactNumber").Value), (.ParameterExists("TRD_MailingAddressNumber").Value), (.ParameterExists("BatchCategory").Value), vReference, (.ParameterExists("COM_Notes").Value), (.ParameterExists("TRD_AdditionalReference1").Value), (.ParameterExists("TRD_AdditionalReference2").Value), (.ParameterExists("TRD_AdditionalReference1Caption").Value), (.ParameterExists("TRD_AdditionalReference2Caption").Value), (.ParameterExists("BatchDate").Value), BatchNumber, TransactionNumber, .ParameterExists("CCU_AddressTo").Value, .ParameterExists("CCU_SalesLedgerAccount").Value, vStockProducts)
            End With

            '------------------------------------------------------------------------
            'Authorise Credit Card transactions
            'Must be done before the Database Transaction is started
            '------------------------------------------------------------------------
            'In order to do the authorisation we must know the amount of the claim
            'We must got thru each analysis line and total the amounts while adjusting for stock issued
            vAmount = 0
            For Each vTDRLine In pTDRTransaction.TraderAnalysisLines
              If vTDRLine.GetTraderLineInfo(TraderAnalysisLine.TraderAnalysisLineInfo.taliCreatesBTA) Then
                vQuantity = vTDRLine.Quantity
                vIssuedValue = ""
                If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBTAQuantityDecimal) Then
                  vQuantityValue = FixedFormat(DoubleValue(vTDRLine.Quantity.ToString))
                  vIssuedValue = FixedFormat(DoubleValue(vTDRLine.Issued.ToString))
                Else
                  vQuantityValue = CStr(vTDRLine.Quantity)
                  vIssuedValue = CStr(vTDRLine.Issued)
                End If
                If Len(vIssuedValue) > 0 And Val(vIssuedValue) <> Val(vQuantityValue) Then
                  'We did not issue all the stock so adjust the amount by the qty actually issued
                  vAmount = vAmount + FixTwoPlaces(vTDRLine.Amount * Val(vIssuedValue) / Val(vQuantityValue))
                  'Only set the flag to process nominal amount if there is an amount > 0
                  If vTDRLine.Amount > 0 Then vIssuedSet = True
                Else
                  vAmount = vAmount + vTDRLine.Amount
                End If
              End If
            Next vTDRLine

            If Not (pParams.Exists("BypassCcAuthorisation") AndAlso
                    pParams("BypassCcAuthorisation").Bool) Then
              If pParams.ContainsKey("TnsSession") AndAlso pParams("TnsSession").Value.Length > 0 Then
                pTDRTransaction.GetCreditCardAuthorisation(vAmount, vIssuedSet, OnlineCCAuthorisation, (pParams.ParameterExists("SecurityCode").Value), pParams("TnsSession").Value)
              ElseIf pParams.ContainsKey("VendorCode") AndAlso mvEnv.GetConfig("fp_cc_authorisation_type") = "SAGEPAYHOSTED" Then
                pTDRTransaction.GetCreditCardAuthorisation(vAmount, vIssuedSet, OnlineCCAuthorisation, "", "", CInt(pParams.ParameterExists("VendorCode").Value), Function() As ParameterList
                                                                                                                                                                    Dim vParams As New ParameterList()
                                                                                                                                                                    vParams.Add("TokenDesc", pParams.ParameterExists("TokenDesc").Value)
                                                                                                                                                                    vParams.Add("TokenId", pParams.ParameterExists("Token").Value)
                                                                                                                                                                    vParams.Add("ContactNumber", pParams.ParameterExists("PayerContactNumber").IntegerValue)
                                                                                                                                                                    vParams.Add("CardDigits", pParams.ParameterExists("CardDigits").Value)
                                                                                                                                                                    vParams.Add("CardExpiryDate", pParams.ParameterExists("CardExpiryDate").Value)
                                                                                                                                                                    Return vParams
                                                                                                                                                                  End Function())
              Else
                pTDRTransaction.GetCreditCardAuthorisation(vAmount, vIssuedSet, OnlineCCAuthorisation, (pParams.ParameterExists("SecurityCode").Value))
              End If
            End If
          Else
            With pParams
              pTDRTransaction.InitFACardSale(vPaymentMethod, vBankAccount, .Item("PayerContactNumber").IntegerValue, .Item("PayerAddressNumber").IntegerValue, .Item("TransactionDate").Value, .Item("TRD_Receipt").Value, .Item("TRD_EligibleForGiftAid").Bool, pExistingTrans, .Item("CDC_CreditCardType").Value, .Item("CDC_CardNumber").Value, .Item("CDC_IssueNumber").Value, .ParameterExists("CDC_ValidDate").Value, .Item("CDC_ExpiryDate").Value, .ParameterExists("CDC_AuthorisationCode").Value, pFinancialAdjustment, BatchCurrencyCode, CStr(BatchExchangeRate), (.ParameterExists("TRD_Mailing").Value), (.ParameterExists("TRD_MailingContactNumber").Value), (.ParameterExists("TRD_MailingAddressNumber").Value), (.ParameterExists("BatchCategory").Value), vReference, (.ParameterExists("COM_Notes").Value), vFATransactionType, (.ParameterExists("BatchDate").Value), BatchNumber, TransactionNumber)
            End With
          End If

        Case "CRED"
          'BR15226: Any changes for CRED should be reflected for CCIN and CQIN
          vStockProducts = False
          For Each vTDRLine In pTDRTransaction.TraderAnalysisLines
            If vTDRLine.StockSale Then vStockProducts = True
            If vStockProducts Then Exit For
          Next vTDRLine

          '2) Add the CreditSale transaction details
          With pParams
            pTDRTransaction.InitCreditSale(vPaymentMethod, vBankAccount, .Item("PayerContactNumber").IntegerValue, .Item("PayerAddressNumber").IntegerValue, .Item("TransactionDate").Value, .Item("TRD_Receipt").Value, .Item("TRD_EligibleForGiftAid").Bool, pExistingTrans, .ParameterExists("CCU_AddressTo").Value, .Item("CCU_SalesLedgerAccount").Value, CSCompany, UseSalesLedger, vStockProducts, ServiceBookingCredits, BatchCurrencyCode, CStr(BatchExchangeRate), (.ParameterExists("TRD_Mailing").Value), (.ParameterExists("TRD_MailingContactNumber").Value), (.ParameterExists("TRD_MailingAddressNumber").Value), (.ParameterExists("BatchCategory").Value), vReference, (.ParameterExists("COM_Notes").Value), pFinancialAdjustment, vFATransactionType, (.ParameterExists("BatchDate").Value), BatchNumber, TransactionNumber)
          End With

      End Select

      'Confirm any provisional transactions for all payment methods
      pTDRTransaction.ConfirmProvisionalTransaction(vConfirmTransList, BatchNumber, TransactionNumber)

      '------------------------------------------------------------------------
      'Start the Transaction here
      '------------------------------------------------------------------------
      mvEnv.Connection.StartTransaction()

      '------------------------------------------------------------------------
      'Save the transaction analysis
      '------------------------------------------------------------------------
      pTDRTransaction.TraderAnalysisLines.SaveAnalysis(mvEnv, pTDRTransaction.BatchTransaction, vInvoiceIssued, pFinancialAdjustment, pExistingTrans, UseSalesLedger, (pParams("TransactionPaymentMethod").Value = "CRED"), PayMethodsAtEnd, BatchCurrencyCode, BatchExchangeRate, (pParams.ParameterExists("OrigBatchNumber").IntegerValue), (pParams.ParameterExists("OrigTransNumber").IntegerValue), pUseStockTransactionID, mvServiceBookingAnalysis, mvEventMultipleAnalysis, mvLinkToFundraisingPayments, pTDRTransaction.Batch.Provisional)
      vTransAmount = pTDRTransaction.BatchTransaction.CurrencyAmount 'Sum of BTA lines
      Dim vInvoiceAmount As Double = pTDRTransaction.BatchTransaction.Amount 'Invoices know nothing about currencies
      '-------------------------------------------------------------------------
      'Add/Update the Credit Customers record as required
      '-------------------------------------------------------------------------
      If pParams("TransactionPaymentMethod").Value = "CRED" OrElse pParams("TransactionPaymentMethod").Value = "CQIN" OrElse pParams("TransactionPaymentMethod").Value = "CCIN" Then
        If pParams("TransactionType").Value = "CRDN" Or ServiceBookingCredits = True Then
          vTransAmount = CDbl(FixedFormat(vTransAmount * -1))
          vInvoiceAmount = CDbl(FixedFormat(vInvoiceAmount * -1))
        End If

        If pParams.Exists("CSDepositAmount") AndAlso (CSDepositPercentage = 0 OrElse CreditNotes) Then pParams.Remove("CSDepositAmount")
        If UseSalesLedger Then
          vOnOrder = pTDRTransaction.CreditCustomer.OnOrder
          vOutstanding = pTDRTransaction.CreditCustomer.Outstanding
        End If

        If ((pTDRTransaction.CreditCustomer.Existing = False Or pTDRTransaction.CreditCustomerDetailsChanged = True) Or (UseSalesLedger = True)) Then
          'Set customer's payment terms
          If ((pParams("CCU_TermsNumber").IntegerValue <> CSTermsNumber) Or (pParams("CCU_TermsFrom").Value <> CSTermsFrom) Or (pParams("CCU_TermsPeriod").Value <> CSTermsPeriod)) And pParams("StorePaymentTerms").Bool = True Then
            'Use customer-specific payment terms
            vTermsNumber = pParams("CCU_TermsNumber").Value
            vTermsPeriod = pParams("CCU_TermsPeriod").Value
            vTermsFrom = pParams("CCU_TermsFrom").Value
          Else
            If pTDRTransaction.CreditCustomer.Existing = False Then
              'Use company default payment terms
              vTermsNumber = ""
              vTermsPeriod = ""
              vTermsFrom = ""
            Else
              'Retain values from Credit Customer record
              With pTDRTransaction.CreditCustomer
                vTermsNumber = .TermsNumber
                vTermsPeriod = .TermsPeriod
                vTermsFrom = .TermsFrom
              End With
            End If
          End If

          If UseSalesLedger Then
            If pExistingTrans Then
              'Need to deduct existing Invoice amount from the appropriate figure
              If mvEnv.Connection.GetCount("invoices", Nothing, "batch_number = " & pTDRTransaction.BatchNumber & " AND transaction_number = " & pTDRTransaction.TransactionNumber) = 0 Then
                vOnOrder = FixTwoPlaces(vOnOrder - pParams("OriginalTransactionAmount").DoubleValue)
              Else
                If pParams("TransactionType").Value = "CRDN" Or ServiceBookingCredits = True Then
                  'When editing an existing sundry credit note, add the original value of the credit note back to the customer's Outstanding amount. The new value of the credit note will be subtracted from the customer's Outstanding later.
                  vOnOrder = FixTwoPlaces(vOnOrder + pParams("OriginalTransactionAmount").DoubleValue)
                Else
                  vOnOrder = FixTwoPlaces(vOnOrder - pParams("OriginalTransactionAmount").DoubleValue)
                End If
              End If
            End If

            If (vInvoiceIssued > 0 And vStockProducts = False) Then
              'Invoice will be created immediately - increment Outstanding
              If pFinancialAdjustment <> Batch.AdjustmentTypes.atAdjustment Then
                vOutstanding = FixTwoPlaces(vOutstanding + vInvoiceAmount)
              End If
            Else
              If pParams("TransactionPaymentMethod").Value = "CQIN" OrElse pParams("TransactionPaymentMethod").Value = "CCIN" Then
                Dim vCreditAmount As Double
                For Each vBTA As BatchTransactionAnalysis In pTDRTransaction.BatchTransaction.Analysis
                  If vBTA.ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSUnderPayProduct) AndAlso _
                    vBTA.RateCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSUnderPayRate) Then
                    'Do nothing
                  ElseIf vBTA.ProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSOverPayProduct) AndAlso _
                    vBTA.RateCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSOverPayRate) Then
                    'Do nothing
                  Else
                    vCreditAmount += vBTA.Amount
                  End If
                Next
                vOnOrder = FixTwoPlaces(vOnOrder + vCreditAmount)
                vOutstanding = FixTwoPlaces(vOutstanding - vInvoiceAmount)
              Else
                'Invoice will be created later - increment OnOrder
                vOnOrder = FixTwoPlaces(vOnOrder + vInvoiceAmount)
              End If
            End If
          End If

          With pParams
            If pTDRTransaction.CreditCustomer.Existing = False Then
              pTDRTransaction.CreditCustomer.Create(mvEnv, .Item("CCU_ContactNumber").IntegerValue, .Item("CCU_AddressNumber").IntegerValue, CSCompany, .Item("CCU_SalesLedgerAccount").Value, .Item("CCU_CreditCategory").Value, .Item("CCU_CreditLimit").DoubleValue, FixedFormat(vOutstanding), .Item("CCU_CustomerType").Value, .ParameterExists("CCU_StopCode").Value, vTermsNumber, vTermsPeriod, vTermsFrom, vOnOrder)
            Else
              vAddressNumber = 0
              If pParams("UpdateCreditCustomerAddress").Bool Then vAddressNumber = pParams("CCU_AddressNumber").IntegerValue
              pTDRTransaction.CreditCustomer.Update(.Item("CCU_CreditCategory").Value, .Item("CCU_CreditLimit").DoubleValue, .Item("CCU_CustomerType").Value, .ParameterExists("CCU_StopCode").Value, vTermsNumber, vTermsPeriod, vTermsFrom, If(UseSalesLedger = True, FixedFormat(vOutstanding), ""), vAddressNumber, If(UseSalesLedger = True, FixedFormat(vOnOrder), ""))
            End If
          End With
          pTDRTransaction.CreditCustomer.Save()
        End If
      End If

      '-------------------------------------------------------------------------
      'Save the Batch Transaction
      '-------------------------------------------------------------------------
      pTDRTransaction.SaveTransaction(pParams("TRD_Amount").DoubleValue, (pParams.ParameterExists("OriginalTransactionAmount").DoubleValue), (pParams.ParameterExists("OriginalTransactionCurrencyAmount").DoubleValue), (pParams.ParameterExists("CCU_TermsFrom").Value), (pParams.ParameterExists("CCU_TermsPeriod").Value), (pParams.ParameterExists("CCU_TermsNumber").Value), (pParams.ParameterExists("CSDepositAmount").DoubleValue), (pParams.ParameterExists("EventNumber").IntegerValue))

      '-------------------------------------------------------------------------
      'Create link to Communications record if relevant
      '-------------------------------------------------------------------------
      If pParams.ParameterExists("CreateCommLink").Bool = True Then
        vInsertFields = New CDBFields
        vInsertFields.Add("batch_number", CDBField.FieldTypes.cftLong, pTDRTransaction.BatchNumber)
        vInsertFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, pTDRTransaction.TransactionNumber)
        vInsertFields.Add("communications_log_number", CDBField.FieldTypes.cftLong, pParams("LinkCommNumber").IntegerValue)
        mvEnv.Connection.InsertRecord("communications_log_trans", vInsertFields, True)
      End If

      '-------------------------------------------------------------------------
      'Update the contact details (dob)
      '-------------------------------------------------------------------------
      Dim vUse As String = ""
      If pParams.Exists("MEM_DateOfBirth") OrElse pParams.Exists("TRD_DateOfBirth") Then
        If pParams.Exists("MEM_DateOfBirth") AndAlso pParams("MEM_DateOfBirth").Value.Length > 0 Then  'use this whether or not TRD value exists
          vUse = "MEM"
        ElseIf pParams.Exists("TRD_DateOfBirth") AndAlso pParams("TRD_DateOfBirth").Value.Length > 0 Then  'only TRD  - use this
          vUse = "TRD"
        Else  'neither have values... if the db has a value for dob this should be cleared.
          vContact = New Contact(mvEnv)
          If pParams.Exists("MEM_DateOfBirth") Then
            vContact.Init((pParams("MEM_ContactNumber").IntegerValue))
            If vContact.DateOfBirth.Length > 0 Then vUse = "MEM"
          Else
            vContact.Init((pParams("TRD_ContactNumber").IntegerValue))
            If vContact.DateOfBirth.Length > 0 Then vUse = "TRD"
          End If
        End If
      End If

      If vUse = "MEM" AndAlso pParams.Exists("MEM_DateOfBirth") Then
        vContact = New Contact(mvEnv)
        vContact.Init((pParams("MEM_ContactNumber").IntegerValue))
        If (vContact.DateOfBirth.Length = 0 AndAlso pParams("MEM_DateOfBirth").Value.Length > 0) OrElse (vContact.DateOfBirth.Length > 0 AndAlso pParams("MEM_DateOfBirth").Value.Length = 0) Then
          vNewDOB = True
        ElseIf vContact.DateOfBirth.Length > 0 AndAlso pParams("MEM_DateOfBirth").Value.Length > 0 Then
          If CDate(vContact.DateOfBirth) <> CDate(pParams("MEM_DateOfBirth").Value) Then vNewDOB = True
        End If
        If vNewDOB Then
          vContact.DateOfBirth = pParams("MEM_DateOfBirth").Value
          vContact.SaveChanges()
        End If
      End If

      If vUse = "TRD" AndAlso pParams.Exists("TRD_DateOfBirth") Then
        vContact = New Contact(mvEnv)
        vContact.Init((pParams("TRD_ContactNumber").IntegerValue))
        If (vContact.DateOfBirth.Length = 0 AndAlso pParams("TRD_DateOfBirth").Value.Length > 0) OrElse (vContact.DateOfBirth.Length > 0 AndAlso pParams("TRD_DateOfBirth").Value.Length = 0) Then
          vNewDOB = True
        ElseIf vContact.DateOfBirth.Length > 0 AndAlso pParams("TRD_DateOfBirth").Value.Length > 0 Then
          If CDate(vContact.DateOfBirth) <> CDate(pParams("TRD_DateOfBirth").Value) Then vNewDOB = True
        End If
        If vNewDOB Then
          vContact.DateOfBirth = pParams("TRD_DateOfBirth").Value
          vContact.SaveChanges()
        End If
      End If

      '-------------------------------------------------------------------------
      'Update the bank transaction if required
      '-------------------------------------------------------------------------
      If AppType = ApplicationType.atCreditListReconciliation Then
        With BankTransaction
          .InitFromValues(mvEnv, pParams("StatementDate").Value, pParams("BankTransactionLineNumber").IntegerValue, pParams("PayersSortCode").Value, pParams("PayersAccountNumber").Value, pParams("PayersName").Value, pParams("ReferenceNumber").Value, pParams("OriginalAmount").DoubleValue)
          If .Existing AndAlso .StatementDate.Length > 0 Then
            .ReconciledStatus = "F"
            .Save()
            .Init(mvEnv)
          End If
        End With
      End If

      '-------------------------------------------------------------------------
      'All inserts/updates completed so commit the Transaction
      '-------------------------------------------------------------------------
      mvEnv.Connection.CommitTransaction()

      '-------------------------------------------------------------------------
      'Finish the transaction creation by setting the return value
      '-------------------------------------------------------------------------
      vReturnValue = SaveTransactionStatus.stsComplete
      If pParams("TransactionPaymentMethod").Value = "CRED" And pTDRTransaction.CSInvoiceCreated = True Then
        If UseSalesLedger = True And (System.Math.Abs(vTransAmount) = System.Math.Abs(pTDRTransaction.BatchTransaction.CurrencyAmount)) And (pTDRTransaction.BatchTransaction.CurrencyAmount > 0 Or ServiceBookingCredits = True) And vStockProducts = False Then
          'May need to print an Invoice
          vReturnValue = vReturnValue Or SaveTransactionStatus.stsPrintInvoice
        End If
      End If
      If ReceiptDocument.Length > 0 Then vReturnValue = vReturnValue Or SaveTransactionStatus.stsPrintReceipt
      If pTDRTransaction.Batch.BatchType = Batch.BatchTypes.Cash And pTDRTransaction.Batch.Provisional = True And Len(ProvisionalCashTransactionDocument) > 0 Then
        vReturnValue = vReturnValue Or SaveTransactionStatus.stsPrintProvisionalCashDoc
      End If
      SaveTransaction = vReturnValue
    End Function

    ''' <summary>Prints Invoices</summary>
    ''' <param name="pParams">Parameters collection from Trader</param>
    ''' <param name="pResults">Parameters collection of results to return to Trader</param>
    ''' <returns>Count of the number of Invoices printed</returns>
    Public Function ProduceInvoice(ByVal pParams As CDBParameters, ByRef pResults As CDBParameters) As Integer
      'Set InvoiceNumber etc. ready for printing
      Dim vInvoice As New Invoice
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vAlias As String = ""
      Dim vContinue As Boolean
      Dim vDate As Date
      Dim vFirstTime As Boolean
      Dim vInvDate As String
      Dim vInvoiceCount As Integer
      Dim vNewStatusCode As String
      Dim vReprint As Boolean
      Dim vTermsFrom As String
      Dim vTermsPeriod As String
      Dim vTermsNumber As Integer
      Dim vTrans As Boolean
      Dim vSupportsProvInv As Boolean
      Dim vUseTransDate As Boolean

      vInvoice.Init(mvEnv)
      vContinue = True
      vFirstTime = True

      Dim vPrintJobNumber As Integer
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPrintJobNumber) Then vPrintJobNumber = mvEnv.GetControlNumber("IP")

      Dim vPartPaidOnly As Boolean = pParams.ParameterExists("PartPaidOnly").Bool AndAlso pParams.ParameterExists("InstantPrint").Bool = False
      Dim vBatchOwnership As Boolean = mvEnv.GetConfigOption("opt_batch_ownership") AndAlso mvEnv.GetConfig("opt_batch_per_user") = "DEPARTMENT"
      vSupportsProvInv = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber)
      If (pParams.ParameterExists("FromDate").Value.Length > 0 OrElse pParams.ParameterExists("ToDate").Value.Length > 0) AndAlso (mvEnv.GetConfig("invoice_date_from_event_start").Length > 0) Then vUseTransDate = True

      If pParams.Exists("BatchNumbers") = False Then pParams.Add("BatchNumbers")
      If pParams.Exists("TransactionNumbers") = False Then pParams.Add("TransactionNumbers")
      If pParams("BatchNumbers").Value.Length > 0 AndAlso pParams("TransactionNumbers").Value.Length > 0 Then
        'We have come from the Invoices grid in Trader where the user has selected the Invoices they wish to print
        vContinue = True
        'Update each invoice with a print job number using Batch and Transaction number
        Dim vBatchNumbers() As String = Split(pParams("BatchNumbers").Value, ",")
        Dim vTranNumbers() As String = Split(pParams("TransactionNumbers").Value, ",")
        vUpdateFields.Add("print_job_number", vPrintJobNumber)
        With vWhereFields
          .Add("batch_number", CDBField.FieldTypes.cftInteger)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
            .Add("print_invoice", CDBField.FieldTypes.cftCharacter, "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
            .Add("print_invoice#2", CDBField.FieldTypes.cftCharacter, "Y", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
          End If
        End With

        If mvEnv.Connection.InTransaction = False Then
          mvEnv.Connection.StartTransaction()
          vTrans = True
        End If

        For vIndex As Integer = 0 To UBound(vBatchNumbers)
          vWhereFields("batch_number").Value = vBatchNumbers(vIndex)
          vWhereFields("transaction_number").Value = vTranNumbers(vIndex)
          vInvoiceCount = mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields, False)
        Next
        pParams.Add("PrintJobNumber", vPrintJobNumber)
        If pParams.Exists("FromInvoiceNumber") = False Then pParams.Add("FromInvoiceNumber", CDBField.FieldTypes.cftInteger, "0")
        If pParams.Exists("ToInvoiceNumber") = False Then pParams.Add("ToInvoiceNumber", CDBField.FieldTypes.cftInteger, "0")
      End If

      If (pParams.ParameterExists("ExistingTransaction").Bool = True And pParams.ParameterExists("InstantPrint").Bool = True And pParams("FromInvoiceNumber").IntegerValue = 0 And pParams.ParameterExists("BatchNumber").IntegerValue > 0) Then
        'Smart Client instant print from Trader (which does not have an InvoiceNumber)
        vInvoice.Init(mvEnv, pParams("BatchNumber").IntegerValue, pParams("TransactionNumber").IntegerValue)
        If DoubleValue(vInvoice.InvoiceNumber) > 0 Then
          pParams("FromInvoiceNumber").Value = vInvoice.InvoiceNumber
          pParams("ToInvoiceNumber").Value = vInvoice.InvoiceNumber
          vUpdateFields.Add("invoice_number", CDBField.FieldTypes.cftInteger, vInvoice.InvoiceNumber)
          vWhereFields.Add("batch_number", pParams("BatchNumber").IntegerValue)
          vWhereFields.Add("transaction_number", pParams("TransactionNumber").IntegerValue)
          mvEnv.Connection.UpdateRecords("invoice_details", vUpdateFields, vWhereFields, False)
          vInvoice = New Invoice
          vInvoice.Init(mvEnv)
        End If
      End If

      Dim vPrintPreview As Boolean = False
      Dim vInvoiceNosAdded As New StringBuilder
      If pParams("ToInvoiceNumber").IntegerValue = 0 Then
        'Mandatory Company, FromInvoiceNumber
        'Select invoices and set invoice numbers
        Dim vSQLStatement As SQLStatement
        If pParams.ParameterExists("PrintJobNumber").IntegerValue > 0 Then
          vPrintPreview = pParams.ParameterExists("PrintPreview").Bool
          vSQLStatement = SelectInvoicesForPrinting(pParams, False, False, False)
        Else
          vSQLStatement = SelectInvoicesForPrinting(pParams, vUseTransDate, vBatchOwnership, vPartPaidOnly)
        End If
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        Dim vInvoiceNumber As Integer
        Dim vProvInvNumber As Integer
        While (vRS.Fetch() = True And vContinue = True)
          vInvoice = New Invoice
          vInvoice.InitFromRecordSet(mvEnv, vRS, Invoice.InvoiceRecordSetTypes.irtAll)
          vInvoiceNumber = 0
          vProvInvNumber = 0
          vUpdateFields.Clear()
          If vPrintPreview Then
            If IntegerValue(vInvoice.InvoiceNumber) = 0 Then
              vInvoiceNumber = mvEnv.GetControlNumber("PR")
              vInvoiceNosAdded.AppendLine(vInvoiceNumber.ToString)
              vUpdateFields.Add("provisional_invoice_number", vInvoiceNumber)
            End If
          Else
            If vSupportsProvInv Then vProvInvNumber = vInvoice.ProvisionalInvoiceNumber
            If IntegerValue(vInvoice.InvoiceNumber) > 0 AndAlso (vProvInvNumber = 0 OrElse IntegerValue(vInvoice.InvoiceNumber) <> vProvInvNumber) Then 'When an advance payment is made
              vInvoiceNumber = IntegerValue(vInvoice.InvoiceNumber)
            Else
              vInvoiceNumber = mvEnv.GetControlNumber("I")
            End If
          End If
          If (vFirstTime OrElse (pParams("FromInvoiceNumber").IntegerValue > vInvoiceNumber)) AndAlso vInvoiceNumber > 0 Then
            pParams("FromInvoiceNumber").Value = vInvoiceNumber.ToString
            vFirstTime = False
          End If
          If IntegerValue(vInvoice.InvoiceNumber) = 0 Or vProvInvNumber > 0 Then
            vUpdateFields.Add("invoice_number", vInvoiceNumber)
          End If
          If vPrintPreview Then
            vInvDate = vInvoice.InvoiceDate
          Else
            'Only update invoice dates if we are NOT doing a print preview
            If vInvoice.InvoiceDate.Length = 0 OrElse _
              (mvEnv.GetConfig("invoice_date_from_event_start").Length > 0 AndAlso vRS.Fields("event_number").Value.Length > 0 AndAlso vInvoice.ReprintCount = -1) _
              OrElse (mvEnv.GetConfig("fp_sl_inv_date_when_printed").ToString = "Y" AndAlso vInvoice.ReprintCount = -1) Then
              vUpdateFields.Add("invoice_date", CDBField.FieldTypes.cftDate, TodaysDate)
              vInvDate = TodaysDate()
            Else
              vInvDate = vInvoice.InvoiceDate
            End If
            If vInvoice.PaymentDue.Length = 0 OrElse (mvEnv.GetConfig("invoice_date_from_event_start").Length > 0 AndAlso vRS.Fields("event_number").Value.Length > 0) Then
              If vRS.Fields("terms_from").Value.Length = 0 Then
                vTermsFrom = vRS.Fields("company_terms_from").Value
                vTermsPeriod = vRS.Fields("company_terms_period").Value
                vTermsNumber = IntegerValue(vRS.Fields("company_terms_number").Value)
              Else
                vTermsFrom = vRS.Fields("terms_from").Value
                vTermsPeriod = vRS.Fields("terms_period").Value
                vTermsNumber = IntegerValue(vRS.Fields("terms_number").Value)
              End If
              vContinue = vInvoice.CalcInvPayDue(vTermsFrom, vTermsPeriod, vTermsNumber, CInt(vRS.Fields("batch_number").Value), CInt(vRS.Fields("transaction_number").Value), CDate(vInvDate), vDate)
              vUpdateFields.Add("payment_due", vDate)
            End If
          End If

          If vContinue Then
            'Update Invoices & InvoiceDetails
            With vWhereFields
              .Clear()
              .Add("batch_number", vRS.Fields("batch_number").IntegerValue)
              .Add("transaction_number", vRS.Fields("transaction_number").IntegerValue)
              If vProvInvNumber > 0 Then
                .Add("provisional_invoice_number", vProvInvNumber)
                .Add("invoice_number", vProvInvNumber)
              End If
            End With
            If vTrans = False Then
              mvEnv.Connection.StartTransaction()
              vTrans = True
            End If
            mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields, False)
            vInvoiceCount += 1  'We have selected an Invoice so print it even if there was nothing to update at this point
            vUpdateFields.Clear()
            If vInvoiceNumber > 0 Then vUpdateFields.Add("invoice_number", vInvoiceNumber)
            If vProvInvNumber > 0 Then
              vWhereFields.Remove("provisional_invoice_number")
            Else
              vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftInteger)
            End If
            mvEnv.Connection.UpdateRecords("invoice_details", vUpdateFields, vWhereFields, False)
            If vProvInvNumber > 0 Then
              vWhereFields.Clear()
              vWhereFields.Add("invoice_number", vProvInvNumber)
              mvEnv.Connection.UpdateRecords("batch_transaction_analysis", vUpdateFields, vWhereFields, False)
              vWhereFields.Add("provisional_invoice_number", vProvInvNumber)
              mvEnv.Connection.UpdateRecords("invoice_payment_history", vUpdateFields, vWhereFields, False)
            End If
          End If
          If pParams("ToInvoiceNumber").IntegerValue < vInvoiceNumber Then
            pParams("ToInvoiceNumber").Value = vInvoiceNumber.ToString
          End If
        End While
        vRS.CloseRecordSet()

      ElseIf Not pParams.ParameterExists("InstantPrint").Bool Then
        vReprint = True
        'Count how many Invoices we have
        Dim vSQLStatement As SQLStatement = SelectInvoicesForPrinting(pParams, True, vBatchOwnership, vPartPaidOnly, True)
        vInvoiceCount = mvEnv.Connection.GetCountFromStatement(vSQLStatement)
      Else
        vInvoiceCount = 1
      End If

      If vInvoiceCount > 0 AndAlso vPrintPreview = False Then
        'Invoice will be printed after it comes out of here
        If vTrans = False Then
          mvEnv.Connection.StartTransaction()
          vTrans = True
        End If
        vWhereFields = New CDBFields
        With vWhereFields
          .Add("invoice_number", pParams("FromInvoiceNumber").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenFrom)
          .Add("invoice_number2", pParams("ToInvoiceNumber").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenTo)
          .Add("company", CDBField.FieldTypes.cftCharacter, pParams("Company").Value)
          .Add("record_type", CDBField.FieldTypes.cftCharacter, "'I','N'", CDBField.FieldWhereOperators.fwoIn)
          If vBatchOwnership Then
            .Add("batch_number", CDBField.FieldTypes.cftLong, "SELECT batch_number FROM batches b, users u, departments d WHERE batch_created_by = u.logname AND u.department = d.department AND d.department = '" & mvEnv.User.Department & "'", CDBField.FieldWhereOperators.fwoIn)
          End If
          If vPartPaidOnly Then .Add("invoice_number#2", CDBField.FieldTypes.cftLong, "SELECT DISTINCT invoice_number FROM invoice_payment_history", CDBField.FieldWhereOperators.fwoIn)
          If pParams.ParameterExists("BatchNumbers").Value.Length = 0 Then
            'We did not display Invoices first
            Dim vRunType As String = pParams.ParameterExists("RunType").Value
            If vRunType.Length = 0 Then vRunType = If(vReprint = True, "R", "N")
            Select Case vRunType
              Case "N"
                .Add("reprint_count", 0, CDBField.FieldWhereOperators.fwoLessThan)
              Case "R"
                .Add("reprint_count", 0, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
              Case Else
                'Do Nothing
            End Select
          End If
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
            .Add("print_invoice", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
            .Add("print_invoice#2", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
          End If
        End With
        vUpdateFields = New CDBFields
        vUpdateFields.Add("reprint_count", CDBField.FieldTypes.cftNumeric, "reprint_count + 1")

        'BR17886, BR18857 - This SQL update does different things depending on what parameters are passed
        'BatchNumbers and Transaction numbers, the number of prints needs to be updated for invoice records that have the print job number, the job number is updated at the top of the method.
        'BatchNumber and Transaction number, the print job number and the number of prints need to be updated 
        'Start and End date the print job number and the number of prints need to be updated
        'Note - if the Batch number and Tranasction number are renamed Batch Numbers and Transaction Numbers this method will still work. 
        If pParams.ParameterExists("BatchNumbers").Value.Length > 0 Then
          vWhereFields.Add("print_job_number", CDBField.FieldTypes.cftLong, vPrintJobNumber)
        Else
          vUpdateFields.Add("print_job_number", CDBField.FieldTypes.cftLong, vPrintJobNumber)
        End If

        mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields)

        'TA BR 8068 Update Status on any associated Event/Accommodation Bookings to indicate invoice has been printed
        'Pass 1: Update Event Bookings
        'Pass 2: Update Accommodation Bookings
        '- Booked (Credit Sale) => Booked (Invoiced)
        '- Booked (Credit Sale) Transfer => Booked (Invoiced) Transfer
        '- Waiting (Credit Sale) => Waiting (Invoiced)
        Dim vTables As String = ""
        For vPass As Integer = 1 To 2
          Select Case vPass
            Case 1
              vTables = "event_bookings"
              vAlias = "eb"
            Case 2
              vTables = "contact_room_bookings"
              vAlias = "crb"
          End Select
          With vWhereFields
            .Clear()
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPrintJobNumber) Then
              .Add("i.print_job_number", CDBField.FieldTypes.cftLong, vPrintJobNumber)
            Else
              .Add("i.invoice_number", pParams("FromInvoiceNumber").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenFrom)
              .Add("i.invoice_number2", pParams("ToInvoiceNumber").IntegerValue, CDBField.FieldWhereOperators.fwoBetweenTo)
              If pParams.ContainsKey("FromDate") Then
                If Len(pParams("FromDate").Value) > 0 Then
                  .Add("i.invoice_date", CDBField.FieldTypes.cftDate, pParams("FromDate").Value, CDBField.FieldWhereOperators.fwoGreaterThan)
                End If
              End If
              If pParams.ContainsKey("ToDate") Then
                If Len(pParams("ToDate").Value) > 0 Then
                  .Add("i.invoice_date2", CDBField.FieldTypes.cftDate, pParams("ToDate").Value, CDBField.FieldWhereOperators.fwoLessThan)
                End If
              End If
              .Add("i.company", CDBField.FieldTypes.cftCharacter, pParams("Company").Value)
              .Add("i.record_type", CDBField.FieldTypes.cftCharacter, "'I','N'", CDBField.FieldWhereOperators.fwoIn)
              If vBatchOwnership Then
                .Add("i.batch_number", CDBField.FieldTypes.cftLong, "SELECT batch_number FROM batches b, users u, departments d WHERE batch_created_by = u.logname AND u.department = d.department AND d.department = '" & mvEnv.User.Department & "'", CDBField.FieldWhereOperators.fwoIn)
              End If
              If vPartPaidOnly Then .Add("invoice_number#2", CDBField.FieldTypes.cftLong, "SELECT DISTINCT invoice_number FROM invoice_payment_history", CDBField.FieldWhereOperators.fwoIn)
              'BR13693: Always look for processed invoices as above WhereFields could select invoices which may not have been printed (print_count=-1)
              'and have invoice number set due to advance payments. Such invoices should not be selected
              .Add("reprint_count", 0, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
                .Add("print_invoice", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
                .Add("print_invoice#2", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
              End If
            End If
            .Add("bta.batch_number", CDBField.FieldTypes.cftLong, "i.batch_number")
            .Add("bta.transaction_number", CDBField.FieldTypes.cftLong, "i.transaction_number")
            .Add(vAlias & ".batch_number", CDBField.FieldTypes.cftLong, "bta.batch_number")
            .Add(vAlias & ".transaction_number", CDBField.FieldTypes.cftInteger, "bta.transaction_number")
            .Add(vAlias & ".line_number", CDBField.FieldTypes.cftLong, "bta.line_number")
            .Add(vAlias & ".booking_status", CDBField.FieldTypes.cftCharacter, "'" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSale) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingCreditSale) & "'", CDBField.FieldWhereOperators.fwoIn)
          End With
          Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT bta.batch_number,bta.transaction_number,bta.line_number,booking_status FROM invoices i,batch_transaction_analysis bta," & vTables & " " & vAlias & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          While vRS.Fetch() = True
            With vWhereFields
              .Clear()
              .Add("batch_number", CDBField.FieldTypes.cftLong, vRS.Fields(1).Value)
              .Add("transaction_number", CDBField.FieldTypes.cftInteger, vRS.Fields(2).Value)
              .Add("line_number", CDBField.FieldTypes.cftInteger, vRS.Fields(3).Value)
              vNewStatusCode = ""
              Select Case vRS.Fields(4).Value
                Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSale)
                  vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoiced)
                Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer)
                  vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer)
                Case EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingCreditSale)
                  vNewStatusCode = EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced)
              End Select
              If vNewStatusCode.Length > 0 Then
                vUpdateFields.Clear()
                vUpdateFields.Add("booking_status", CDBField.FieldTypes.cftCharacter, vNewStatusCode)
                mvEnv.Connection.UpdateRecords(vTables, vUpdateFields, vWhereFields, False)
              End If
            End With
          End While
          vRS.CloseRecordSet()
        Next
      ElseIf vPrintPreview = True Then
        pResults.Add("PrintPreview", pParams("PrintPreview").Value)
        If vInvoiceNosAdded.ToString.Length > 0 Then
          Dim vInvoiceNumbers As String = vInvoiceNosAdded.ToString.Replace(vbCrLf, ",")
          If vInvoiceNumbers.EndsWith(",") Then vInvoiceNumbers = vInvoiceNumbers.Substring(0, vInvoiceNumbers.Length - 1)
          pResults.Add("InvoiceNumbersAdded", vInvoiceNumbers)
        End If
      End If
      If vTrans Then mvEnv.Connection.CommitTransaction()
      If vPrintJobNumber > 0 Then pResults.Add("PrintJobNumber", vPrintJobNumber)
      Return vInvoiceCount

    End Function

    Public Function CheckBatchLedApp(ByRef pBatchType As Batch.BatchTypes, ByRef pProvisionalCash As Boolean) As Boolean
      'Checks whether batch-led Trader Application is valid
      'Sets pBatchType & pProvisionalCash
      Dim vNoBatchTypes As Integer
      Dim vValid As Boolean

      vValid = mvValid
      pBatchType = Batch.BatchTypes.None
      pProvisionalCash = False
      If mvValid = True And mvBatchLedApp = True And mvAppType = ApplicationType.atTransaction Then
        If Cash Or Cheque Or PostalOrder Then
          pBatchType = Batch.BatchTypes.Cash '"CA"
          vNoBatchTypes = vNoBatchTypes + 1
          pProvisionalCash = IncludeProvisionalTransactions
        End If
        If CreditCard Then
          pBatchType = Batch.BatchTypes.CreditCard '"CC"
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If DebitCard Then
          pBatchType = Batch.BatchTypes.DebitCard '"DC"
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If CreditSales Then
          pBatchType = Batch.BatchTypes.CreditSales '"CS"
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If ChequeWithInvoice Then
          pBatchType = Access.Batch.BatchTypes.CashWithInvoice
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If CCWithInvoice Then
          pBatchType = Access.Batch.BatchTypes.CreditCardWithInvoice
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If Voucher Then
          pBatchType = Batch.BatchTypes.CAFVouchers
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If CAFCard Then
          pBatchType = Batch.BatchTypes.CAFCards
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If GiftInKind Then
          pBatchType = Batch.BatchTypes.GiftInKind
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If SupportsNonFinancialBatch Then
          pBatchType = Batch.BatchTypes.NonFinancial
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If SaleOrReturn Then
          pBatchType = Batch.BatchTypes.SaleOrReturn
          vNoBatchTypes = vNoBatchTypes + 1
        End If
        If vNoBatchTypes > 1 Then vValid = False
      End If

      CheckBatchLedApp = vValid

    End Function

    Public Sub SCGetTransactionData(ByVal pParams As CDBParameters, ByRef pTransaction As TraderTransaction, ByVal pResults As CDBParameters)
      'Copied from frmTrader.GetTransactionData
      Dim vAnalysisColl As CollectionList(Of BatchTransactionAnalysis)
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vBT As BatchTransaction
      Dim vOBT As New BatchTransaction(mvEnv)
      Dim vCardSale As New CardSale(mvEnv)
      Dim vContactAccount As New ContactAccount
      Dim vConfirmedTrans As New ConfirmedTransaction(mvEnv)
      Dim vCreditCustomer As CreditCustomer
      Dim vCreditSale As CreditSale
      Dim vCSTerms As CreditSalesTerms
      Dim vCompanyControls As CompanyControl
      Dim vGPPH As PreTaxPGPaymentHistory = Nothing
      Dim vPGPH As PostTaxPgPaymentHistory = Nothing
      Dim vWhereFields As CDBFields
      Dim vAddressNumber As Integer
      Dim vContactNumber As Integer
      Dim vOrigPaymentMethod As String = ""
      Dim vOutstanding As Double
      Dim vPayMethod As String
      Dim vPayMethod1 As String = ""
      Dim vSQL As String
      Dim vTransDate As String
      Dim vBTA As BatchTransactionAnalysis
      Dim vFH As FinancialHistory
      Dim vAmount As Double
      Dim vCurrencyAmount As Double
      Dim vLineNumber As Integer
      Dim vAnalysisLine As TraderAnalysisLine

      vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)

      vOBT.Init((Me.BatchNumber), (Me.TransactionNumber))
      If vOBT.Existing Then
        vPayMethod = vOBT.PaymentMethod
        GetPayMethod(vPayMethod, vOrigPaymentMethod, vPayMethod1, vFinancialAdjustment)

        If Me.BatchLedApp = False AndAlso vFinancialAdjustment = Batch.AdjustmentTypes.atNone AndAlso Me.Batch Is Nothing Then
          'If we are editing a transaction in Smart Client from View Batch Details then trader app may not be batch led so ensure batch is set
          Me.Batch = New Batch(mvEnv)
          Me.Batch.Init(Me.BatchNumber)
        End If

        If vPayMethod1 = "CRED" Then
          vCreditSale = New CreditSale(mvEnv)
          vCreditSale.Init((vOBT.BatchNumber), (vOBT.TransactionNumber))

          vCreditCustomer = New CreditCustomer
          vCreditCustomer.Init(mvEnv, (vCreditSale.ContactNumber), Me.CSCompany)

          vCompanyControls = New CompanyControl
          vCompanyControls.InitFromBankAccount(mvEnv, Me.Batch.BankAccount)

          vCSTerms = New CreditSalesTerms
          vCSTerms.Init(mvEnv, (vCreditSale.ContactNumber), (vCompanyControls.Company), (vCreditSale.SalesLedgerAccount))

          If vFinancialAdjustment <> Batch.AdjustmentTypes.atMove Then
            vOutstanding = 0
            If Me.UseSalesLedger Then
              'Get the contact's o/s payment plan balance
              vWhereFields = New CDBFields
              With vWhereFields
                .Add("contact_number", CDBField.FieldTypes.cftLong, vCreditSale.ContactNumber)
                .Add("order_type", CDBField.FieldTypes.cftCharacter, "'O','M'", CDBField.FieldWhereOperators.fwoIn)
                .Add("order_date", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
                .Add("cancellation_reason")
              End With
              vSQL = "SELECT SUM(balance) AS total_balance FROM orders WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
              vOutstanding = Val(mvEnv.Connection.GetValue(vSQL))
            End If

            vContactNumber = vCreditSale.ContactNumber
            vAddressNumber = vCreditSale.AddressNumber
            pResults.Add("CCU_ContactNumber", vCreditSale.ContactNumber)
            pResults.Add("CCU_TermsNumber")
            pResults.Add("CCU_TermsPeriod")
            pResults.Add("CCU_TermsFrom")
            With vCreditCustomer
              pResults.Add("CCU_CreditCategory", CDBField.FieldTypes.cftCharacter, .CreditCategory)
              pResults.Add("CCU_StopCode", CDBField.FieldTypes.cftCharacter, .StopCode)
              pResults.Add("CCU_CreditLimit", CDBField.FieldTypes.cftNumeric, .CreditLimit.ToString)
              pResults.Add("CCU_CustomerType", CDBField.FieldTypes.cftCharacter, .CustomerType)
              pResults.Add("CCU_OnOrder", CDBField.FieldTypes.cftNumeric, .OnOrder.ToString)
              vOutstanding = vOutstanding + .Outstanding
              pResults.Add("CCU_Outstanding", CDBField.FieldTypes.cftNumeric, vOutstanding.ToString)
              If Me.InvoicePayments Then
                If Val(.TermsNumber) = 0 And (Val(.TermsNumber) <> Me.CSTermsNumber) Then
                  pResults("CCU_TermsNumber").Value = .TermsNumber
                  pResults("CCU_TermsPeriod").Value = .TermsPeriod
                  pResults("CCU_TermsFrom").Value = .TermsFrom
                End If
              End If
              pResults("CCU_TermsNumber").Value = If(Val(.TermsNumber) = 0, Me.CSTermsNumber.ToString, .TermsNumber)
              pResults("CCU_TermsPeriod").Value = If(Len(.TermsPeriod) = 0, Me.CSTermsPeriod, .TermsPeriod)
              pResults("CCU_TermsFrom").Value = If(Len(.TermsFrom) = 0, Me.CSTermsFrom, .TermsFrom)
            End With
            pResults.Add("CCU_AddressNumber", vCreditSale.AddressNumber)
            pResults.Add("CCU_SalesLedgerAccount", CDBField.FieldTypes.cftCharacter, vCreditSale.SalesLedgerAccount)
          End If
          pResults.Add("CCU_StockSale", CDBField.FieldTypes.cftCharacter, BooleanString(vCreditSale.StockSale))
          pResults.Add("CCU_AddressTo", CDBField.FieldTypes.cftCharacter, vCreditSale.AddressTo)
          pResults.Add("CCU_Reference", CDBField.FieldTypes.cftCharacter, vOBT.Reference)
          pResults("CCU_TermsNumber").Value = vCSTerms.TermsNumber.ToString
          pResults("CCU_TermsPeriod").Value = vCSTerms.TermsPeriod
          pResults("CCU_TermsFrom").Value = vCSTerms.TermsFrom

        ElseIf vPayMethod1 = "GAYE" Then
          If mvBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.GiveAsYouEarn) Then
            Dim vPaymentFields As New CDBFields(New CDBField("batch_number", BatchNumber))
            vPaymentFields.Add("transaction_number", TransactionNumber)
            vGPPH = New PreTaxPGPaymentHistory(mvEnv)
            vGPPH.InitWithPrimaryKey(vPaymentFields)
          ElseIf mvBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.PostTaxPayrollGiving) Then
            vPGPH = New PostTaxPgPaymentHistory
            vPGPH.Init(mvEnv, (Me.BatchNumber), (Me.TransactionNumber))
            If vPGPH.Existing = False Then
              'No payment history - see if this is a payment linked to the Employer
              vPGPH.InitOrganisationPayment(mvEnv, Me.BatchNumber, Me.TransactionNumber)
              If vPGPH.OrganisationPayment = False Then
                RaiseError(DataAccessErrors.daeCannotFindPGPledgePayHistory, CStr(Me.BatchNumber), CStr(Me.TransactionNumber))
              End If
            End If
          End If
        End If

        If vOBT.BankDetailsNumber > 0 And Me.BankDetails = True Then
          vContactAccount.Init(mvEnv, (vOBT.BankDetailsNumber))
          If vContactAccount.Existing Then
            If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
              'No need to display account holder
            Else
              pResults.Add("BKD_SortCode", CDBField.FieldTypes.cftCharacter, vContactAccount.SortCode)
              pResults.Add("BKD_AccountNumber", CDBField.FieldTypes.cftCharacter, vContactAccount.AccountNumber)
              pResults.Add("BKD_Reference", CDBField.FieldTypes.cftCharacter, vOBT.Reference)
              pResults.Add("BKD_IbanNumber", CDBField.FieldTypes.cftCharacter, vContactAccount.IbanNumber)
              pResults.Add("BKD_BicCode", CDBField.FieldTypes.cftCharacter, vContactAccount.BicCode)
            End If
          End If
        End If

        Select Case vFinancialAdjustment
          Case Batch.AdjustmentTypes.atMove, Batch.AdjustmentTypes.atGIKConfirmation
            'Clear values
            vContactNumber = 0
            vAddressNumber = 0
          Case Else
            vContactNumber = vOBT.ContactNumber
            vAddressNumber = vOBT.AddressNumber
        End Select

        If vFinancialAdjustment <> Batch.AdjustmentTypes.atAdjustment And vFinancialAdjustment <> Batch.AdjustmentTypes.atMove Then
          vOBT.InitBatchTransactionAnalysis((Me.BatchNumber), (Me.TransactionNumber))
          vOBT.InitAnalysisAdditionalData()
          vOBT.InitAnalysisStockMovements()
          If vOBT.Analysis.Count() > 0 Then
            pResults.Add("TransactionSource", CDBField.FieldTypes.cftCharacter, vOBT.Analysis.Item(0).Source)
            pResults.Add("TRD_Source", CDBField.FieldTypes.cftCharacter, vOBT.Analysis.Item(0).Source)
          End If
        End If

        Select Case vFinancialAdjustment
          Case Batch.AdjustmentTypes.atNone, Batch.AdjustmentTypes.atGIKConfirmation, Batch.AdjustmentTypes.atCashBatchConfirmation
            vTransDate = vOBT.TransactionDate
            If vFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Then
              If Me.Batch.BatchType = Batch.BatchTypes.SaleOrReturn Then
                'For Sale or Return  and Cash batches need to set contact/address to the original values
                vContactNumber = vOBT.ContactNumber
                vAddressNumber = vOBT.AddressNumber
              End If
            End If
          Case Else
            If pParams.Exists("TransactionDate") Then
              vTransDate = pParams("TransactionDate").Value
            Else
              If mvEnv.GetConfig("fp_adjust_transaction_date") = "today" Then
                vTransDate = TodaysDate()
              Else
                vTransDate = vOBT.TransactionDate
              End If
            End If
        End Select

        'Set specific page values for Payroll Giving
        If Me.AppType = ApplicationType.atGiveAsYouEarnPayments Then
          With vGPPH
            pResults.Add("GYE_DonorTotal", CDBField.FieldTypes.cftNumeric, .DonorAmount.ToString)
            pResults.Add("GYE_EmployerTotal", CDBField.FieldTypes.cftNumeric, .EmployerAmount.ToString)
            pResults.Add("GYE_GovernmentTotal", CDBField.FieldTypes.cftNumeric, .GovernmentAmount.ToString)
            pResults.Add("GYE_AdminFeesAmount", CDBField.FieldTypes.cftNumeric, .AdminFeeAmount.ToString)
          End With
        ElseIf Me.AppType = ApplicationType.atPostTaxPGPayments Then
          If vPGPH.Existing = True Or vPGPH.OrganisationPayment = True Then
            pResults.Add("PGP_DonorTotal", CDBField.FieldTypes.cftNumeric, vPGPH.DonorAmount.ToString)
            pResults.Add("PGP_EmployerTotal", CDBField.FieldTypes.cftNumeric, vPGPH.EmployerAmount.ToString)
          End If
        End If

        pResults.Add("TRD_Provisional", CDBField.FieldTypes.cftCharacter, "N")
        If Me.Batch.Provisional = True And (vFinancialAdjustment <> Batch.AdjustmentTypes.atGIKConfirmation And vFinancialAdjustment <> Batch.AdjustmentTypes.atCashBatchConfirmation) Then
          pResults("TRD_Provisional").Value = "Y"
          vConfirmedTrans.Init((Me.BatchNumber), (Me.TransactionNumber))
          If vConfirmedTrans.Existing Then
            pResults.Add("TRD_AdditionalReference1", CDBField.FieldTypes.cftCharacter, vConfirmedTrans.AdditionalReference1)
            pResults.Add("TRD_AdditionalReference2", CDBField.FieldTypes.cftCharacter, vConfirmedTrans.AdditionalReference2)
          End If
        End If

        If Me.Batch.Provisional = False And Me.Batch.BatchType = Batch.BatchTypes.CAFVouchers Then
          vConfirmedTrans.InitConfirmed((Me.BatchNumber), (Me.TransactionNumber))
          If vConfirmedTrans.Existing Then
            pResults.Add("TRD_AdditionalReference1", CDBField.FieldTypes.cftCharacter, vConfirmedTrans.AdditionalReference1)
            pResults.Add("TRD_AdditionalReference2", CDBField.FieldTypes.cftCharacter, vConfirmedTrans.AdditionalReference2)
          End If
        End If

        If vPayMethod1 = "CARD" Or vPayMethod1 = "CAFC" Then
          vCardSale.Init((Me.BatchNumber), (Me.TransactionNumber))
          With vCardSale
            If .Existing Then
              pResults.Add("CDC_CreditOrDebitCard", CDBField.FieldTypes.cftCharacter, If(vPayMethod = mvEnv.GetConfig("pm_cc"), "C", "D"))
              pResults.Add("CDC_CardNumber", CDBField.FieldTypes.cftCharacter, .CardNumber)
              pResults.Add("CDC_IssueNumber", CDBField.FieldTypes.cftCharacter, .IssueNumber)
              pResults.Add("CDC_ValidDate", CDBField.FieldTypes.cftCharacter, .ValidDate)
              pResults.Add("CDC_ExpiryDate", CDBField.FieldTypes.cftCharacter, .ExpiryDate)
              pResults.Add("CDC_AuthorisationCode", CDBField.FieldTypes.cftCharacter, .AuthorisationCode)
              pResults.Add("CDC_Reference", CDBField.FieldTypes.cftCharacter, vOBT.Reference)
              If .NoClaimRequired And (vFinancialAdjustment <> Batch.AdjustmentTypes.atMove) And (vFinancialAdjustment <> Batch.AdjustmentTypes.atAdjustment) Then RaiseError(DataAccessErrors.daeCannotEditAuthorisedTransaction)
            End If
          End With
        End If

        pResults.Add("TransactionAmount", If(Len(Me.BatchCurrencyCode) > 0, vOBT.CurrencyAmount, vOBT.Amount))
        pResults.Add("TransactionCurrencyAmount", If(Len(Me.BatchCurrencyCode) > 0, vOBT.Amount, vOBT.CurrencyAmount))
        pResults.Add("TransactionPaymentMethod", CDBField.FieldTypes.cftCharacter, vPayMethod1)
        pResults.Add("OrigPaymentMethod", CDBField.FieldTypes.cftCharacter, vOrigPaymentMethod)
        pResults.Add("TransactionDate", CDBField.FieldTypes.cftDate, vTransDate)
        pResults.Add("PayerContactNumber", vContactNumber)
        pResults.Add("PayerAddressNumber", vAddressNumber)
        pResults.Add("TRD_ContactNumber", vContactNumber)
        pResults.Add("TRD_AddressNumber", vAddressNumber)
        pResults.Add("TRD_TransactionDate", CDBField.FieldTypes.cftDate, vTransDate)
        pResults.Add("TRD_Amount", If(Len(Me.BatchCurrencyCode) > 0, vOBT.CurrencyAmount, vOBT.Amount))
        pResults.Add("TRD_Reference", CDBField.FieldTypes.cftCharacter, vOBT.Reference)
        If vFinancialAdjustment <> Batch.AdjustmentTypes.atMove Then pResults.Add("TRD_Mailing", CDBField.FieldTypes.cftCharacter, vOBT.Mailing)
        pResults.Add("TRD_Receipt", CDBField.FieldTypes.cftCharacter, vOBT.Receipt)
        pResults.Add("TRD_EligibleForGiftAid", CDBField.FieldTypes.cftCharacter, BooleanString(vOBT.EligibleForGiftAid))
        If vFinancialAdjustment = Batch.AdjustmentTypes.atMove Or vFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Then
          '
        Else
          pResults.Add("TRD_MailingContactNumber", vOBT.MailingContactNumber)
          pResults.Add("TRD_MailingAddressNumber", vOBT.MailingAddressNumber)
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataTransactionOrigins) Then pResults.Add("TRD_TransactionOrigin", CDBField.FieldTypes.cftCharacter, vOBT.TransactionOrigin)
        pResults.Add("COM_Notes", CDBField.FieldTypes.cftCharacter, vOBT.Notes)

        If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or vFinancialAdjustment = Batch.AdjustmentTypes.atMove Then
          If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment And pParams.Exists("BatchNumbers") Then
            'We have multiple transactions to be selected
            vAnalysisColl = New CollectionList(Of BatchTransactionAnalysis)
            vFH = New FinancialHistory
            vFH.Init(mvEnv)
            For Each vBT In vFH.GetMultipleTransactions(pParams("BatchNumbers").Value)
              vBT.InitDetailsFromFinancialHistory(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
              vBT.InitAnalysisAdditionalData()
              vAmount = vAmount + If(Len(Me.BatchCurrencyCode) > 0, vBT.CurrencyAmount, vBT.Amount)
              vCurrencyAmount = vCurrencyAmount + If(Len(Me.BatchCurrencyCode) > 0, vBT.Amount, vBT.CurrencyAmount)
              If Len(vTransDate) = 0 Then pResults("TransactionDate").Value = vBT.TransactionDate
              For Each vBTA In vBT.Analysis
                vAnalysisColl.Add(vBTA.Key(True), vBTA)
              Next vBTA
            Next vBT
            pResults("TransactionAmount").Value = vAmount.ToString
            pResults("TransactionCurrencyAmount").Value = vCurrencyAmount.ToString
            pResults("TRD_Amount").Value = vAmount.ToString
          Else
            vBT = New BatchTransaction(mvEnv)
            vBT.InitDetailsFromFinancialHistory(mvEnv, Me.BatchNumber, Me.TransactionNumber)
            vBT.InitAnalysisAdditionalData()
            If Len(vTransDate) = 0 Then pResults("TransactionDate").Value = vBT.TransactionDate
            vAnalysisColl = vBT.Analysis
          End If
        Else
          vAnalysisColl = vOBT.Analysis
        End If
        pTransaction.TraderAnalysisLines.InitAnalysisFromBT(vAnalysisColl, Me.Batch.BatchType, vFinancialAdjustment, Me.SundryCreditProduct, vOBT.ContactNumber, vOBT.AddressNumber, Me.BatchCurrencyCode, vTransDate, vOBT.StockMovements, Me.CreditNotes)
        pTransaction.TraderAnalysisLines.SetDepositAllowed(mvEnv)

        If vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Then
          vLineNumber = 0
          If pParams.Exists("BatchNumbers") Then
            For Each vAnalysisLine In pTransaction.TraderAnalysisLines
              'Reset line number for each line as Smart Client works with unique line numbers only
              vLineNumber = vLineNumber + 1
              vAnalysisLine.SetLineNumber(vLineNumber, vFinancialAdjustment)
            Next vAnalysisLine
          End If
          If pResults.Exists("TransactionSource") = False And pTransaction.TraderAnalysisLines.Count > 0 Then pResults.Add("TransactionSource", CDBField.FieldTypes.cftCharacter, pTransaction.TraderAnalysisLines(1).Source)
          pResults.Add("TRD_DateOfBirth", CDBField.FieldTypes.cftDate, vOBT.DataTable.Rows(0).Item("DateOfBirth")) 'BR17343 - Get the date of birth so that TRD form can be intialised, rename serverside so no changes client side.
        End If
      Else
        RaiseError(DataAccessErrors.daeCannotFindBatchTransaction, CStr(Me.BatchNumber), CStr(Me.TransactionNumber))
      End If
    End Sub

    Private Function SCDoEditAnalysis(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction) As TraderPage.TraderPageType
      'frmTrader.cmdEdit
      Dim vTDRLine As TraderAnalysisLine
      Dim vLineType As TraderAnalysisLine.TraderAnalysisLineTypes
      Dim vNextPage As TraderPage.TraderPageType
      Dim vTransactionType As String
      Dim vPPD As PaymentPlanDetail
      Dim vSubs As Subscription

      vNextPage = pCurrentPageType
      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary
          'Decide which page to go to
          'Expect pTransaction to only contain the line we are editing
          'So just get first item from collection (we do not know which line number it is)
          vTDRLine = pTransaction.TraderAnalysisLines(1)
          vTransactionType = vTDRLine.TraderTransactionTypeCode
          vLineType = vTDRLine.TraderLineType
          If vTransactionType Like "[SGH]" And vTDRLine.PaymentPlanNumber > 0 Then
            vTransactionType = vTDRLine.PaymentPlanTypeCode
          End If
          If vLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment Then vTransactionType = "COLP"

          Select Case vTransactionType
            Case "EVNT"
              vNextPage = TraderPage.TraderPageType.tpEventBooking

            Case "EXAM"
              vNextPage = TraderPage.TraderPageType.tpExamBooking

            Case "M", "C", "O", "MEMB", "MEMC", "SUBS", "DONR", "CMEM", "CSUB", "CDON"
              If Not (vTransactionType Like "[MCO]") Then
              Else
              End If
              vTransactionType = "PAYM"
              If (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atNone Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment) Then
                'Set this scheduled payment as removed
                EditScheduledPaymentAnalysisLine(pTransaction, vTDRLine, pParams, True)
              End If
              'The OutstandingScheduledPayments page always need to be refreshed as the payments could have changed since last time
              'So change DefaultsSet to False to re-select the Scheduled Paymnets
              vNextPage = TraderPage.TraderPageType.tpPayments
              If vTDRLine.AcceptAsFull Then pParams.Add("AcceptAsFull", CDBField.FieldTypes.cftCharacter, "Y")

            Case "SALE", "DONS", "G", "S", "H", "P", "CSRT", "D", "F"
              Select Case vTDRLine.TraderLineType
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
                  vNextPage = TraderPage.TraderPageType.tpAmendEventBooking
                Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit, TraderAnalysisLine.TraderAnalysisLineTypes.taltLegacyBequestReceipt
                  'Do not allow editing
                Case Else
                  vNextPage = TraderPage.TraderPageType.tpProductDetails
              End Select

            Case "COLP"
              vNextPage = TraderPage.TraderPageType.tpCollectionPayments
            Case "STAT"
              vNextPage = TraderPage.TraderPageType.tpSetStatus
            Case "CGAD"
              vNextPage = TraderPage.TraderPageType.tpCancelGiftAidDeclaration
            Case "CANC"
              vNextPage = TraderPage.TraderPageType.tpCancelPaymentPlan
          End Select

        Case TraderPage.TraderPageType.tpPaymentPlanSummary
          'Decide which page to go to
          'Expect pTransaction to only contain the line we are editing
          'So just get first item from collection (we do not know which line number it is)
          vPPD = pTransaction.TraderPPDLines(1)
          If (AppType = ApplicationType.atMaintenance Or PayPlanConversionMaintenance) Then '"MAINT"
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance
          Else
            vNextPage = TraderPage.TraderPageType.tpPaymentPlanProducts
          End If

          With vPPD
            If .Existing And .Subscription Then
              'For a Subscription Product, do not allow editing of the Product if it has a Subscription record
              If .SubscriptionNumber < 0 Then
                vSubs = New Subscription
                vSubs.Init(mvEnv, vPPD.SubscriptionNumber(False))
                If Not vSubs.Existing Then
                  .SetSubscriptionNumber((0))
                End If
              End If
            End If
          End With
        Case TraderPage.TraderPageType.tpPurchaseInvoiceSummary, TraderPage.TraderPageType.tpPurchaseOrderSummary
          vNextPage = LinePage
      End Select

      SCDoEditAnalysis = vNextPage
    End Function

    Public Function UpdateOutstanding(ByVal pCompany As String, ByVal pSLAccount As String, ByVal pAmount As Double, Optional ByRef pRaiseError As Boolean = True, Optional ByRef pErrorMsg As String = "") As Boolean
      Dim vCC As New CreditCustomer

      vCC.Init(mvEnv, 0, pCompany, pSLAccount)
      If vCC.Existing Then
        vCC.AdjustOutstanding(pAmount)
        vCC.Save()
      Else
        UpdateOutstanding = True
        If pRaiseError Then
          RaiseError(DataAccessErrors.daeCreditCustomerMissing1, pCompany, pSLAccount)
        Else
          If pErrorMsg <> "" Then pErrorMsg = pErrorMsg & vbCrLf
          pErrorMsg = pErrorMsg & "Failed to retrieve the Customer Record for Account %s and Company %s"
        End If
      End If
    End Function

    Public Function SCDoDeleteAnalysis(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pResults As CDBParameters) As TraderPage.TraderPageType
      'frmTrader.cmdDelete
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vLineType As TraderAnalysisLine.TraderAnalysisLineTypes
      Dim vTDRLine As TraderAnalysisLine
      Dim vFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vContinue As Boolean
      Dim vDelBTA As Boolean
      Dim vEventNumber As Integer
      Dim vLineNumbers As String
      Dim vLineTotal As Double
      Dim vNextPage As TraderPage.TraderPageType
      Dim vReturnMsg As String = ""
      Dim vTDRLine2 As TraderAnalysisLine

      vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)
      vContinue = True
      vDelBTA = pParams.ParameterExists("ExistingTransaction").Bool Or mvExistingAdjustmentTran = True

      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpTransactionAnalysisSummary
          vNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
          'Expect pTransaction to only contain the line we are deleting
          'So just get first item from collection (we do not know which line number it is)
          vTDRLine = pTransaction.TraderAnalysisLines(1)

          mvEnv.Connection.StartTransaction()

          Dim vBT As BatchTransaction = Nothing
          vLineType = vTDRLine.TraderLineType
          Select Case vLineType
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltEvent
              DeleteEventBooking(vTDRLine.EventBookingNumber, (Not pParams.ParameterExists("ExistingTransaction").Bool), vEventNumber, vReturnMsg)
              If vEventNumber > 0 Then pResults.Add("EventNumber", vEventNumber)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltExamBooking
              DeleteExamBooking(vTDRLine.ExamBookingNumber, (Not pParams.ParameterExists("ExistingTransaction").Bool), vReturnMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAccomodation
              DeleteRoomBooking(vTDRLine.RoomBookingNumber, False, vEventNumber, vReturnMsg)
              If vEventNumber > 0 Then pResults.Add("EventNumber", vEventNumber)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBooking, TraderAnalysisLine.TraderAnalysisLineTypes.taltServiceBookingCredit
              DeleteServiceBooking(vTDRLine.ServiceBookingNumber, False, vReturnMsg)
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLine.TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash,
                 TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation '"N" Invoice Payment, "L" Invoice Allocation, "U" Unallocated Sales Ledger Cash, "K" Sundry Credit Note Invoice Allocation
              If vFinancialAdjustment <> Batch.AdjustmentTypes.atNone Then
                vDelBTA = False
                vContinue = False
              End If
              If vFinancialAdjustment = Batch.AdjustmentTypes.atNone OrElse (vFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment AndAlso IsOriginalAnalsysisLine(vBT, vTDRLine, pParams) = False) Then
                DeleteInvoicePayment(vTDRLine, vReturnMsg, vDelBTA, pParams.ParameterExists("BatchNumber").LongValue, pParams.ParameterExists("TransactionNumber").LongValue)
              End If
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltLegacyBequestReceipt '"B"
              If pParams.ParameterExists("ExistingTransaction").Bool = True Then
                vDelBTA = (vFinancialAdjustment = Batch.AdjustmentTypes.atNone)
                DeleteLegacyBequestReceipt(vTDRLine.LegacyNumber, (vTDRLine.BequestNumber), (vTDRLine.LegacyReceiptNumber))
              End If
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCollectionPayment
              If pParams.ParameterExists("ExistingTransaction").Bool = True Then
                DeleteCollectionPayment(vTDRLine, vReturnMsg)
              End If
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltGiftAidDeclaration
              'Nothing to do
            Case TraderAnalysisLine.TraderAnalysisLineTypes.taltPayrollGivingPledge
              'Nothing to do
            Case Else
              vDelBTA = False 'Never delete the BTA here otherwise Cancel will not undo
          End Select

          If vContinue = True Then
            'If we've gotten this far then make sure that the line we're about to delete actually exists in the database
            If vDelBTA Then vDelBTA = ExistingAnalysisLine(pParams.ParameterExists("BatchNumber").LongValue, pParams.ParameterExists("TransactionNumber").LongValue, vTDRLine.LineNumber)

            If (vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltPaymentPlan Or vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltMembership Or vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltCovenant) Then
              If (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atNone Or pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment) Then
                EditScheduledPaymentAnalysisLine(pTransaction, vTDRLine, pParams, CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Batch.AdjustmentTypes) = Batch.AdjustmentTypes.atAdjustment)
              End If
            End If

            If vDelBTA = True And (pParams.ParameterExists("BatchNumber").IntegerValue > 0 And pParams.ParameterExists("TransactionNumber").IntegerValue > 0) Then
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pParams("BatchNumber").IntegerValue)
              vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pParams("TransactionNumber").IntegerValue)
              vLineNumbers = CStr(vTDRLine.LineNumber)
              For Each vTDRLine2 In pTransaction.TraderAnalysisLines
                If vTDRLine2.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltIncentive Then
                  If vTDRLine2.IncentiveLineNumber = vTDRLine.LineNumber Then
                    If Len(vLineNumbers) > 0 Then vLineNumbers = vLineNumbers & ","
                    vLineNumbers = vLineNumbers & vTDRLine2.LineNumber
                  End If
                End If
              Next vTDRLine2

              'If deleting an L-type (Invoice Allocation line) or a K-type (Sundry Credit Note Invoice Allocation line), will need to delete the other L-type/K-type line 
              'as there will now be 2 L-type/K-type lines when the Unallocated SL Cash/ Sundry Credit Note is allocated to an invoice
              If vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltInvoiceAllocation OrElse vTDRLine.TraderLineType = TraderAnalysisLine.TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation Then
                vFields = New CDBFields
                With vFields
                  .Add("batch_number", pParams("BatchNumber").LongValue)
                  .Add("transaction_number", pParams("TransactionNumber").LongValue)
                  .Add("line_number", vTDRLine.LineNumber, CDBField.FieldWhereOperators.fwoNotEqual)
                  .Add("line_type", vTDRLine.GetAnalysisLineTypeCode(vTDRLine.TraderLineType))
                  .Add("invoice_number", vTDRLine.InvoiceNumber)
                End With
                Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "batch_number, transaction_number, line_number, amount", "batch_transaction_analysis bta", vFields)
                Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
                While vRS.Fetch = True
                  pParams("DetailLineTotal").Value = FixTwoPlaces(pParams("DetailLineTotal").DoubleValue - vRS.Fields("amount").DoubleValue).ToString
                  If vLineNumbers.Length > 0 Then vLineNumbers &= ","
                  vLineNumbers &= vRS.Fields("line_number").Value
                End While
                vRS.CloseRecordSet()
              End If
              vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, vLineNumbers, CDBField.FieldWhereOperators.fwoIn)
              mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields)

              vLineTotal = FixTwoPlaces(pParams("DetailLineTotal").DoubleValue - vTDRLine.Amount)
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pParams("BatchNumber").IntegerValue)
              vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pParams("TransactionNumber").IntegerValue)
              vFields = New CDBFields
              vFields.AddAmendedOnBy(mvEnv.User.Logname)
              vFields.Add("line_total", CDBField.FieldTypes.cftNumeric, vLineTotal)
              mvEnv.Connection.UpdateRecords("batch_transactions", vFields, vWhereFields)

            End If
          End If

          If Not (pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atGIKConfirmation Or pParams.ParameterExists("TransactionType").Value = "CSRT") Then
            'Do not need to re-allocate any product numbers
            If vTDRLine.ProductNumber > 0 Then ReAllocateProductNumber(vTDRLine.ProductCode, vTDRLine.ProductNumber)
          End If

          'Remove analysis line from collection so that it is not sent back
          pTransaction.TraderAnalysisLines.Remove((1))

          mvEnv.Connection.CommitTransaction()
          If vReturnMsg.Length > 0 Then pResults.Add("InformationMessage", CDBField.FieldTypes.cftCharacter, vReturnMsg)
      End Select

      SCDoDeleteAnalysis = vNextPage

    End Function

    Public Sub DeleteEventBooking(ByVal pBookingNumber As Integer, ByVal pOnlyIfNoTransaction As Boolean, ByRef pEventNumber As Integer, ByRef pMsg As String)
      Dim vEventBooking As New EventBooking

      vEventBooking.Init(mvEnv, 0, pBookingNumber)
      pEventNumber = vEventBooking.EventNumber
      If pOnlyIfNoTransaction And vEventBooking.TransactionProcessed Then Exit Sub
      vEventBooking.Delete()
      If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
      pMsg = pMsg & String.Format(ProjectText.String29316, CStr(pBookingNumber)) 'Deleted Event Booking %s
    End Sub

    Public Sub DeleteExamBooking(ByVal pBookingNumber As Integer, ByVal pOnlyIfNoTransaction As Boolean, ByRef pMsg As String)
      Dim vExamBooking As New ExamBooking(mvEnv)

      vExamBooking.Init(pBookingNumber)
      If vExamBooking.Existing Then             'May have been deleted by another analysis line
        If pOnlyIfNoTransaction And vExamBooking.TransactionProcessed Then Exit Sub
        vExamBooking.Delete()
        If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
        pMsg = pMsg & String.Format(ProjectText.String29319, CStr(pBookingNumber)) 'Deleted Exam Booking %s
      End If
    End Sub

    Public Sub DeletePaymentPlan(ByVal pPaymentPlanNumber As Integer, Optional ByRef pMsg As String = "", Optional ByVal pMemberNumber As Integer = 0)
      'this method is only to be used if the transaction has not been created
      Dim vPaymentPlan As New PaymentPlan
      Dim vOPS As New OrderPaymentSchedule

      vPaymentPlan.Init(mvEnv, pPaymentPlanNumber)
      vPaymentPlan.Delete(pPaymentPlanNumber)
      If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
      pMsg = pMsg & String.Format(ProjectText.String19120, CStr(pPaymentPlanNumber)) 'Payment Plan Deleted:
      If pMemberNumber > 0 Then
        If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
        pMsg = pMsg & String.Format(ProjectText.String19121, CStr(pMemberNumber)) 'Membership Deleted - Member Number:
      End If
    End Sub

    Public Sub DeleteRoomBooking(ByVal pBookingNumber As Integer, ByRef pOnlyIfNoTransaction As Boolean, ByRef pEventNumber As Integer, ByRef pMsg As String)
      Dim vRoomBooking As EventAccommodationBooking

      vRoomBooking = New EventAccommodationBooking
      vRoomBooking.Init(mvEnv, pBookingNumber)
      pEventNumber = vRoomBooking.EventNumber
      If pOnlyIfNoTransaction And vRoomBooking.TransactionProcessed Then Exit Sub
      vRoomBooking.Delete()
      If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
      pMsg = pMsg & String.Format(ProjectText.String29318, CStr(pBookingNumber)) 'Deleted Room Booking %s
    End Sub

    Public Sub DeleteServiceBooking(ByVal pBookingNumber As Integer, ByRef pOnlyIfNoTransaction As Boolean, ByRef pMsg As String)
      Dim vSB As New ServiceBooking

      vSB.Init(mvEnv, pBookingNumber)
      If pOnlyIfNoTransaction And vSB.TransactionProcessed Then Exit Sub
      vSB.Delete()
      If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
      pMsg = pMsg & String.Format(ProjectText.String29317, CStr(pBookingNumber)) 'Deleted Service Booking %s
    End Sub

    Public Sub DeleteLegacyBequestReceipt(ByVal pLegacyNumber As Integer, ByRef pBequestNumber As Integer, ByRef pLegacyReceiptNumber As Integer)
      Dim vLegacyBequest As New LegacyBequest(mvEnv)

      vLegacyBequest.Init(pLegacyNumber, pBequestNumber)
      If vLegacyBequest.Receipts.ContainsKey(CStr(pLegacyReceiptNumber)) Then
        vLegacyBequest.DeleteReceipt(pLegacyReceiptNumber)
      End If
    End Sub

    Private Sub SaveNonFinancial(ByRef pTransaction As TraderTransaction)
      Dim vContact As Contact
      Dim vOrganisation As Organisation = Nothing
      Dim vBatchNumber As Integer
      Dim vTransactionNumber As Integer
      Dim vTransactionAdded As Boolean
      Dim vTDRLine As TraderAnalysisLine
      Dim vCC As ContactCategory
      Dim vCSuppression As ContactSuppression
      Dim vCS As ContactSuppression
      Dim vValidFrom As String
      Dim vValidTo As String
      Dim vSource As String
      Dim vNotes As String
      Dim vAllocateTransNo As Boolean

      For Each vTDRLine In pTransaction.TraderAnalysisLines
        Select Case vTDRLine.TraderLineType
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAddressUpdate, TraderAnalysisLine.TraderAnalysisLineTypes.taltGiftAidDeclaration,
               TraderAnalysisLine.TraderAnalysisLineTypes.taltPayrollGivingPledge
            'Do nothing
          Case Else
            vAllocateTransNo = True
        End Select
      Next vTDRLine

      'Determine whether an open non-financial batch exists or a new one will have to be created
      If SupportsNonFinancialBatch Then
        InitNonFinancialTransaction(vAllocateTransNo)
        vBatchNumber = NonFinancialBatchNumber
        vTransactionNumber = NonFinancialTransactionNumber
      End If

      For Each vTDRLine In pTransaction.TraderAnalysisLines
        Select Case vTDRLine.TraderLineType
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltStatus
            vContact = New Contact(mvEnv)
            vContact.Init((vTDRLine.DeliveryContactNumber))
            If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
              vOrganisation = New Organisation(mvEnv)
              vOrganisation.Init((vTDRLine.DeliveryContactNumber))
            End If
            vContact.Status = vTDRLine.ContactStatus
            vContact.StatusDate = TodaysDate()
            vContact.StatusReason = AppDesc
            vContact.Save("", False, vBatchNumber, vTransactionNumber)
            vTransactionAdded = True
            If Not vOrganisation Is Nothing Then
              vOrganisation.Status = vTDRLine.ContactStatus
              vOrganisation.StatusDate = TodaysDate()
              vOrganisation.StatusReason = AppDesc
              vOrganisation.Save()
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltActivityEntry
            If pTransaction.Activities.Count() > 0 Then
              vContact = New Contact(mvEnv)
              vContact.Init(CType(pTransaction.Activities.Item(1), ContactCategory).ContactNumber)
              For Each vCC In pTransaction.Activities
                With vContact
                  vCC.ContactTypeSaveActivity(.ContactType, .ContactNumber, vCC.Activity, vCC.ActivityValue, vCC.Source, vCC.ValidFrom, vCC.ValidTo, vCC.Quantity, ContactCategory.ActivityEntryStyles.aesAllowMultipleSource, vCC.Notes, "", "", vCC.ActivityDate)
                End With
              Next vCC
              If vBatchNumber > 0 Then mvEnv.AddJournalRecord(JournalTypes.jnlActivity, JournalOperations.jnlInsert, vContact.ContactNumber, 0, 0, 0, 0, vBatchNumber, vTransactionNumber)
              vTransactionAdded = True
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltAddSuppression
            If pTransaction.Suppressions.Count() > 0 Then
              vCSuppression = New ContactSuppression(mvEnv)
              vContact = New Contact(mvEnv)
              vContact.Init(CType(pTransaction.Suppressions.Item(1), ContactSuppression).ContactNumber)
              For Each vCS In pTransaction.Suppressions
                If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
                  vCSuppression.Init(ContactSuppression.ContactSuppressionTypes.cstOrganisation, vCS.ContactNumber, vCS.MailingSuppression, vCS.ValidFrom, vCS.ValidTo)
                Else
                  vCSuppression.Init(ContactSuppression.ContactSuppressionTypes.cstContact, vCS.ContactNumber, vCS.MailingSuppression, vCS.ValidFrom, vCS.ValidTo)
                End If
                vValidFrom = vCS.ValidFrom
                vValidTo = vCS.ValidTo
                vSource = vCS.Source
                vNotes = vCS.Notes
                If vCSuppression.Existing Then
                  If CDate(vValidFrom) > CDate(vCSuppression.ValidFrom) Then vValidFrom = vCSuppression.ValidFrom
                  If CDate(vValidTo) < CDate(vCSuppression.ValidTo) Then vValidTo = vCSuppression.ValidTo

                  If Len(vNotes) > 0 And Len(vCSuppression.Notes) > 0 Then vNotes = vCSuppression.Notes & vbCrLf & vNotes
                  vCSuppression.Update(vValidFrom, vValidTo, vNotes, vSource)
                Else
                  vCSuppression.Create((vCS.ContactNumber), (vCS.MailingSuppression), vValidFrom, vValidTo, vNotes, vSource)
                End If
                vCSuppression.Save()
              Next vCS
              If vBatchNumber > 0 Then mvEnv.AddJournalRecord(JournalTypes.jnlSuppression, JournalOperations.jnlInsert, vContact.ContactNumber, 0, 0, 0, 0, vBatchNumber, vTransactionNumber)
              vTransactionAdded = True
            End If
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCancelPaymentPlan
            Dim vPayPlan As PaymentPlan = New PaymentPlan
            vPayPlan.Init(mvEnv, vTDRLine.PaymentPlanNumber)
            Dim vStatus As String = ""
            Dim vDesc As String = ""
            mvEnv.GetCancellationInfo(vTDRLine.CancellationReason, vStatus, vDesc)
            vPayPlan.Cancel(PaymentPlan.PaymentPlanCancellationTypes.pctCovenant Or PaymentPlan.PaymentPlanCancellationTypes.pctMembership Or PaymentPlan.PaymentPlanCancellationTypes.pctPaymentPlan, vTDRLine.CancellationReason, vStatus, vDesc, mvEnv.User.Logname, vTDRLine.Source, vBatchNumber, vTransactionNumber)
            vTransactionAdded = True
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltGoneAway
            vContact = New Contact(mvEnv)
            vContact.Init((vTDRLine.DeliveryContactNumber))
            vContact.MarkAsGoneAway(vBatchNumber, vTransactionNumber)
            'use the MarkAsGoneAway method of the Contact class
            vTransactionAdded = True
          Case TraderAnalysisLine.TraderAnalysisLineTypes.taltCancelGiftAidDeclaration
            Dim vGAD As GiftAidDeclaration = New GiftAidDeclaration
            vGAD.Init(mvEnv, vTDRLine.DeclarationNumber)
            vGAD.Cancel(vTDRLine.CancellationReason, vTDRLine.Source, , , vBatchNumber, vTransactionNumber)
            vTransactionAdded = True
          Case Else 'Not a Maintenance Option Line
            'do nothing
        End Select
      Next vTDRLine
      If SupportsNonFinancialBatch And vTransactionAdded Then mvNonFinancialBatch.UpdateNumberOfTransactions(1)
    End Sub

    Public Sub DeleteCollectionPayment(ByVal pTDRLine As TraderAnalysisLine, ByRef pMsg As String)
      Dim vCP As New CollectionPayment
      Dim vPayNo As Integer

      vCP.InitFromBatch(mvEnv, pTDRLine.CollectionNumber, Me.BatchNumber, Me.TransactionNumber, pTDRLine.LineNumber)
      If vCP.Existing Then
        vPayNo = vCP.CollectionPaymentNumber
        vCP.Delete()
      End If
      If pMsg.Length > 0 Then pMsg = pMsg & vbCrLf
      pMsg = pMsg & String.Format(ProjectText.String18380, CStr(vPayNo)) 'Deleted Collection Payment %s

    End Sub

    Private Sub GetSegmentProductRate(ByVal pSourceCode As String, ByVal pTransactionType As String, ByVal pPayMethod As String, ByVal pCurrentPageType As TraderPage.TraderPageType, ByVal pFinancialAdjustment As Batch.AdjustmentTypes)
      Dim vRecordSet As CDBRecordSet
      Dim vProductSQL As String

      If pSourceCode.Length > 0 AndAlso mvEnv.GetConfigOption("default_analysis_from_source") Then
        vProductSQL = ProductSQL(pTransactionType, pPayMethod, pCurrentPageType, pFinancialAdjustment)

        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT spa.product, spa.rate, uses_product_numbers FROM segments s, segment_product_allocation spa, products p WHERE s.source = '" & pSourceCode & "' AND s.campaign = spa.campaign AND s.appeal = spa.appeal AND s.segment = spa.segment and spa.product = p.product AND " & vProductSQL & " ORDER BY spa.amount_number")
        With vRecordSet
          If .Fetch() = True Then
            Me.SetProductFromSource((.Fields(1).Value), (.Fields(2).Value), (.Fields(3).Bool))
          Else
            Me.SetProductFromSource("", "", False)
          End If
          .CloseRecordSet()
        End With
      End If
    End Sub

    Private Function ProductSQL(ByVal pTransactionType As String, ByVal pPayMethod As String, ByVal pCurrentPageType As TraderPage.TraderPageType, ByVal pFinancialAdjustment As Batch.AdjustmentTypes) As String
      Dim vDF As New DataFinder
      Dim vCourse As Boolean
      Dim vAccommodation As Boolean
      Dim vPostagePacking As Boolean
      Dim vMembershipProduct As Boolean
      Dim vProvisionalConf As Boolean

      vProvisionalConf = (pFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Or pTransactionType = "CSRT")

      vDF.Init(mvEnv, DataFinder.DataFinderTypes.dftProduct)
      With vDF
        If Not (vProvisionalConf And mvEnv.GetConfigOption("fp_prompt_confirm_hist_product")) Then
          .AddSelectItem("history_only", "N", CDBField.FieldTypes.cftCharacter)
        End If
        If pCurrentPageType = TraderPage.TraderPageType.tpProductDetails Then
          If pTransactionType = "CRDN" Then
            vCourse = True
            vAccommodation = True
            vPostagePacking = True
          Else
            .AddSelectItem("subscription", "N", CDBField.FieldTypes.cftCharacter)
            If pTransactionType = "SALE" Then
              .AddSelectItem("donation", "N", CDBField.FieldTypes.cftCharacter)
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidSponsorship) Then .AddSelectItem("sponsorship_event", "N", CDBField.FieldTypes.cftCharacter, CDBField.FieldWhereOperators.fwoNullOrEqual)
              If (pPayMethod = "GFIK" Or pPayMethod = "SAOR") Then .AddSelectItem("stock_item", "N", CDBField.FieldTypes.cftCharacter)
            ElseIf pTransactionType = "DONS" Then
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidSponsorship) Then
                .AddSelectItem("donation_or_sponsorship_event", "Y", CDBField.FieldTypes.cftCharacter)
              Else
                .AddSelectItem("donation", "Y", CDBField.FieldTypes.cftCharacter)
              End If
            End If
          End If
        ElseIf pCurrentPageType = TraderPage.TraderPageType.tpEventBooking Then
          .AddSelectItem("course", "Y", CDBField.FieldTypes.cftCharacter)
          vCourse = True
        ElseIf pCurrentPageType = TraderPage.TraderPageType.tpAccommodationBooking Then
          .AddSelectItem("accommodation", "Y", CDBField.FieldTypes.cftCharacter)
          vAccommodation = True
        ElseIf pCurrentPageType = TraderPage.TraderPageType.tpPostageAndPacking Then
          .AddSelectItem("postage_packing", "Y", CDBField.FieldTypes.cftCharacter)
          vPostagePacking = True
        End If
        If Not vCourse Then .AddSelectItem("course", "N", CDBField.FieldTypes.cftCharacter)
        If Not vAccommodation Then .AddSelectItem("accommodation", "N", CDBField.FieldTypes.cftCharacter)
        If Not vPostagePacking Then .AddSelectItem("postage_packing", "N", CDBField.FieldTypes.cftCharacter)
        If Not vMembershipProduct Then .AddSelectItem("membership_product", "N", CDBField.FieldTypes.cftCharacter)
        If Len(Me.SalesGroup) > 0 Then .AddSelectItem("sales_group", (Me.SalesGroup), CDBField.FieldTypes.cftCharacter)
      End With
      ProductSQL = vDF.GetProductRestrictionSQL()
    End Function

    Private Sub DefaultPPMemberToPayer(ByVal pPaymentPlan As PaymentPlan, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer)
      Dim vContact As New Contact(mvEnv)

      vContact.Init(pContactNumber, pAddressNumber)

      With pPaymentPlan.Member
        .ContactNumber = vContact.ContactNumber
        .ContactDesc = vContact.Name
        .ContactDOBEstimated = vContact.DobEstimated
        .ContactDateOfBirth = vContact.DateOfBirth
        .AddressNumber = vContact.Address.AddressNumber
        .AddressDesc = vContact.AccessCheckAddressLine
        .Branch = vContact.Address.Branch
      End With
    End Sub

    ''' <summary>Select Invoices for display on tpBatchInvoiceSummary Trader page</summary>
    ''' <param name="pParams">Parameters collection from Trader</param>
    ''' <returns>CDBParameters Collection of all Invoices selected</returns>
    Public Function GetBatchInvoices(ByRef pParams As CDBParameters) As Collection
      Dim vBatchOwnership As Boolean = (mvEnv.GetConfigOption("opt_batch_ownership") = True AndAlso mvEnv.GetConfig("opt_batch_per_user") = "DEPARTMENT")
      Dim vPartPaidOnly As Boolean = pParams.ParameterExists("PartPaidOnly").Bool
      Dim vUseTransDate As Boolean
      If (pParams.ParameterExists("FromDate").Value.Length > 0 OrElse pParams.ParameterExists("ToDate").Value.Length > 0) OrElse (mvEnv.GetConfig("invoice_date_from_event_start").Length > 0) Then vUseTransDate = True

      Dim vSQLStatement As SQLStatement = SelectInvoicesForPrinting(pParams, vUseTransDate, vBatchOwnership, vPartPaidOnly)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      Dim vParams As New CDBParameters
      Dim vColl As New Collection
      While vRS.Fetch = True
        vParams = New CDBParameters
        With vParams
          .Add("Print")
          .Add("RecordType")
          .Add("InvoiceDate")
          .Add("InvoiceNumber")
          .Add("Company")
          .Add("BatchNumber")
          .Add("TransactionNumber")
          .Add("ContactNumber")
          .Add("Name")
          .Add("SalesLedgerAccount")
          .Add("Amount")
          .Add("GrossAmount")
          .Add("VATAmount")
          .Add("EventDescription")
          .Add("EventNumber")
          .Item("Print").Value = "N"
          .Item("RecordType").Value = vRS.Fields(1).Value
          .Item("InvoiceDate").Value = vRS.Fields(2).Value
          .Item("InvoiceNumber").Value = vRS.Fields(3).Value
          .Item("Company").Value = vRS.Fields(4).Value
          .Item("BatchNumber").Value = vRS.Fields(5).Value
          .Item("TransactionNumber").Value = vRS.Fields(6).Value
          .Item("ContactNumber").Value = vRS.Fields(7).Value
          .Item("Name").Value = vRS.Fields(8).Value
          .Item("SalesLedgerAccount").Value = vRS.Fields(9).Value
          .Item("Amount").Value = vRS.Fields(10).Value
          .Item("GrossAmount").Value = vRS.Fields(11).Value
          .Item("VATAmount").Value = vRS.Fields(12).Value
          .Item("EventDescription").Value = vRS.Fields(13).Value
          .Item("EventNumber").Value = vRS.Fields(14).Value
        End With
        vColl.Add(vParams)
      End While
      vRS.CloseRecordSet()

      Return vColl

    End Function

    Public Function GetDefaultPayFreq(ByVal pPaymentMethod As String) As String
      Dim vPaymentFreq As String = String.Empty
      If Len(pPaymentMethod) > 0 Then
        vPaymentFreq = mvEnv.Connection.GetValue("SELECT payment_frequency FROM payment_methods WHERE payment_method = '" & pPaymentMethod & "'")
      End If
      Return vPaymentFreq
    End Function

    Public Function GetPayMethodCode(ByRef pPayMethod As String) As String
      Select Case pPayMethod
        Case "STDO"
          GetPayMethodCode = mvEnv.GetConfig("pm_so")
        Case "DIRD"
          GetPayMethodCode = mvEnv.GetConfig("pm_dd")
        Case "CCCA"
          GetPayMethodCode = mvEnv.GetConfig("pm_ccca")
        Case "CASH"
          GetPayMethodCode = mvEnv.GetConfig("pm_cash")
        Case "POST"
          GetPayMethodCode = mvEnv.GetConfig("pm_po")
        Case "CHEQ"
          GetPayMethodCode = mvEnv.GetConfig("pm_cheque")
        Case "CRED", "CCIN", "CQIN"
          GetPayMethodCode = CSPayMethod
        Case "CARD"
          If (CreditCard And Not DebitCard) Then
            GetPayMethodCode = mvEnv.GetConfig("pm_cc") 'This app only support credit cards
          ElseIf DebitCard And Not CreditCard Then
            GetPayMethodCode = mvEnv.GetConfig("pm_dc") 'This app only supports debit cards
          Else
            GetPayMethodCode = mvEnv.GetConfig("pm_cc") 'We don't know yet, so default to credit card
          End If
        Case "VOUC"
          GetPayMethodCode = mvEnv.GetConfig("pm_voucher")
        Case "CAFC"
          GetPayMethodCode = mvEnv.GetConfig("pm_caf_card")
        Case Else
          GetPayMethodCode = mvEnv.GetConfig("pm_cash")
      End Select
    End Function

    Public Function GetPayPlanEligibleForGiftAid(ByRef pParams As CDBParameters, ByVal pTDRTransaction As TraderTransaction, ByVal pPaymentPlan As PaymentPlan) As String
      'Set the eligible_for_gift_aid control
      Dim vEligible As String
      Dim vTransactionType As String
      'Return Values
      'N- Enabled and Unchecked
      'Y- Enabled and Checked
      'D- Disabled and Unchecked

      vEligible = "Y"
      vTransactionType = pParams.ParameterExists("TransactionType").Value
      If (vTransactionType = "MEMB" Or vTransactionType = "MEMC" Or vTransactionType = "CMEM") Then
        'New Membership:  vEligible will be either 'Y' or 'D'
        'CMT:  If the new MembershipType is eligible for Gift Aid then vEligible will be either 'Y' or 'N' (setting on Payment Plan)
        'CMT:  If the new MembershipType is not eligible for Gift Aid then vEligible will be 'D'
        'For everything else, vEligible will be 'Y'
        If pPaymentPlan.DeterminePaymentPlanGiftAidEligibility(pParams, pTDRTransaction) Then
          If vTransactionType = "MEMC" Then
            'For CMT set the checkbox to the original PayPlan value
            vEligible = BooleanString(pPaymentPlan.EligibleForGiftAid = True)
          Else
            vEligible = "Y"
          End If
        Else
          vEligible = "D"
        End If
      End If

      GetPayPlanEligibleForGiftAid = vEligible

    End Function

    Public Function GetPayPlanEligibleForGiftAid(ByRef pParams As CDBParameters, ByVal pTDRTransaction As TraderTransaction) As String
      Return GetPayPlanEligibleForGiftAid(pParams, pTDRTransaction, pTDRTransaction.PaymentPlan)
    End Function

    Private Sub ProcessPaymentPlan(ByRef pParams As CDBParameters, ByRef pTraderTransaction As TraderTransaction, ByRef pResults As CDBParameters, ByRef pNextPage As TraderPage.TraderPageType, Optional ByVal pUpdatePPDSource As Boolean = False)
      Dim vOPS As OrderPaymentSchedule
      Dim vNumber As String = ""
      Dim vAmount As Double
      Dim vType As String = ""
      Dim vSource As String
      Dim vCount As Integer
      Dim vUseRenewalAmount As Boolean
      Dim vFinished As Boolean
      Dim vOPSFound As Boolean
      Dim vFreqAmount As Double
      Dim vTDRLine As TraderAnalysisLine
      Dim vTALine As Integer
      Dim vTransAmount As Double 'lblTASAmount
      Dim vTransTotal As Double 'lblTASTotal
      Dim vTransactionType As String
      Dim vAdditionalParams As New CDBParameters

      vTALine = 1
      vTransactionType = pParams("TransactionType").Value
      vTransAmount = pParams.ParameterExists("TRD_Amount").DoubleValue 'Amount transaction is to add up to
      vTransTotal = pParams.ParameterExists("DetailLineTotal").DoubleValue 'Amount of current analysis lines

      If pParams.Exists("PaymentMethod") = False Then
        pParams.Add("PaymentMethod")
        If (AppType = ApplicationType.atMaintenance Or AppType = ApplicationType.atConversion) And pParams.Exists("PPM_PaymentMethod") Then
          pParams("PaymentMethod").Value = pParams("PPM_PaymentMethod").Value
        Else
          pParams("PaymentMethod").Value = GetPayMethodCode((pParams("PPPaymentType").Value))
        End If
      End If
      vAdditionalParams.Add("AppType", If(AppType > 0, AppType, 0))
      vAdditionalParams.Add("ConversionShowPPD", CDBField.FieldTypes.cftCharacter, BooleanString(ConversionShowPPD))
      vAdditionalParams.Add("PayPlanConversionMaintenance", CDBField.FieldTypes.cftCharacter, BooleanString(PayPlanConversionMaintenance))

      If pParams.Exists("Provisional") = False And ProvisionalPaymentPlan Then pParams.Add("Provisional", CDBField.FieldTypes.cftCharacter, "Y")
      If pParams.Exists("Provisional") And Not ProvisionalPaymentPlan Then pParams("Provisional").Value = "N"

      If pParams.ParameterExists("CheckIncentives").Bool Then mvCheckIncentives = True

      If (pParams.ParameterExists("CheckIncentives").Bool Or vTransactionType = "MEMB") Or ((AppType = ApplicationType.atConversion Or AppType = ApplicationType.atMaintenance Or vTransactionType = "MEMC") And mvCheckIncentives) Then
        If pParams("CurrentPageType").IntegerValue = TraderPage.TraderPageType.tpMembershipPayer Then 'CMT
          vSource = pParams("CMT_Source").Value
        Else
          vSource = ""
        End If
        If Len(vSource) > 0 Or (pParams.ParameterExists("CheckIncentives").Bool) And AppType <> ApplicationType.atMaintenance Then
          pParams.Add("AddIncentiveToPPD", CDBField.FieldTypes.cftCharacter, "N")
          If AppType = ApplicationType.atConversion Or vTransactionType = "MEMC" Then pParams("AddIncentiveToPPD").Value = "Y"
          ProcessIncentives(pParams, pTraderTransaction, pTraderTransaction.PaymentPlan, pParams("PayerContactNumber").IntegerValue, vSource, vAdditionalParams)
        End If
      End If

      'Check for write-offs (PPMaintenance & CMT only)
      Select Case vTransactionType
        Case "MAINT"
          If pParams.Exists("PPM_WriteOffMissedPayments") = True OrElse pParams.Exists("WriteOffMissedPayments") = True Then
            'OK
          Else
            pParams.Add("WriteOffMissedPayments", BooleanString(mvEnv.GetConfigOption("fp_pp_wo_missed_payments", False)))
          End If
        Case "MEMC"
          If pParams.Exists("CMT_WriteOffMissedPayments") = True OrElse pParams.Exists("WriteOffMissedPayments") = True Then
            'OK
          ElseIf pParams.Exists("CMT_WriteOffOldMembershipCost") Then
            'Original Write-off now superceded so use that value
            pParams.Add("WriteOffMissedPayments", pParams("CMT_WriteOffOldMembershipCost").Value)
          Else
            pParams.Add("WriteOffMissedPayments", BooleanString(mvEnv.GetConfigOption("fp_pp_wo_missed_payments", False)))
          End If
        Case Else
          If AppType = ApplicationType.atMaintenance OrElse PayPlanConversionMaintenance = True Then
            If pParams.Exists("PPM_WriteOffMissedPayments") = False AndAlso pParams.Exists("WriteOffMissedPayments") = False Then
              pParams.Add("WriteOffMissedPayments", BooleanString(mvEnv.GetConfigOption("fp_pp_wo_missed_payments", False)))
            End If
          Else
            If pParams.Exists("WriteOffMissedPayments") = True Then pParams.Remove("WriteOffMissedPayments")
          End If
      End Select

      pTraderTransaction.PaymentPlan.SavePaymentPlan(pParams, pTraderTransaction, pResults, vAdditionalParams)

      'If we have just created the PayPlan or done a CMT as part of a transaction
      If (Not pTraderTransaction.PaymentPlan.Existing) Or (pTraderTransaction.PaymentPlan.Existing And vTransactionType = "MEMC" And pParams.ParameterExists("CreateTransaction").Bool = True) Then ' And mvCMTStartPoint = cmtspTransaction) Then
        With pTraderTransaction.PaymentPlan
          If Not .Existing And .PlanType = CDBEnvironment.ppType.pptMember And .FixedRenewalCycle And .PreviousRenewalCycle And (.ProportionalBalanceSetting And (PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew)) > 0 And .StartDate = .RenewalDate Then
            vUseRenewalAmount = pParams.ParameterExists("PPD_UseAsFirstAmount").Bool
            If .FirstAmount.Length > 0 Then
              vAmount = Val(.FirstAmount)
            Else
              vAmount = pParams("PPBalance").DoubleValue
            End If
          Else
            vAmount = pParams("PPBalance").DoubleValue
          End If
        End With

        'Set the row on the Transaction analysis Summary grid
        'If you are doing a transaction reanalysis and turning a one-off payment into
        'payment plan then some of the columns on the appropriate TAS grid row needs clearing
        vTALine = pTraderTransaction.TraderAnalysisLines.Count + 1
        vTDRLine = pTraderTransaction.GetTraderAnalysisLine(vTALine)
        vTDRLine.Init(vTALine, vTransactionType)

        If pParams.ParameterExists("CurrentPaymentMethod").Bool = True AndAlso pTraderTransaction.PaymentPlan.LoanStatus = PaymentPlan.ppYesNoCancel.ppNo Then
          Select Case pParams.ParameterExists("TransactionPaymentMethod").Value
            Case "CCIN", "CQIN"
              'Always pay the full Payment Plan amount
            Case Else
              'If user entered transaction amount, and this is less than pay plan amount then use the transaction amount
              'Take account of any bta already added.
              If vTransAmount > 0 And (FixTwoPlaces(vTransAmount - vTransTotal) < vAmount) Then
                vAmount = FixTwoPlaces(vTransAmount - vTransTotal)
                If vAmount < 0 Then vAmount = 0
              End If
          End Select

          Select Case vTransactionType
            Case "MEMB", "CMEM", "MEMC"
              vType = "M"
              vNumber = pTraderTransaction.PaymentPlan.Member.MemberNumber 'mvMemberNumber
              If vTransactionType = "MEMB" Then
                If pTraderTransaction.PaymentPlan.GiverContactNumber.Length > 0 Then vType = "H"
              ElseIf vTransactionType = "MEMC" Then
              End If
            Case "CSUB", "CDON"
              vType = "C"
            Case "SUBS", "DONR"
              vType = "O"
            Case "SALE", "EVNT", "ACOM", "SRVC"
              vType = "O"
          End Select

          If vAmount = 0 Then
            vTDRLine.AddNonPaymentLine(pTraderTransaction.PaymentPlan.PlanNumber, pParams("TransactionSource").Value, vType, vNumber)
            'PopulateGridFromTALine vTDRLine
            vType = "NP"
          End If

          vCount = 1
          While vAmount > 0
            vFreqAmount = pTraderTransaction.PaymentPlan.FrequencyAmount
            'Scheduled Payments collection is in reverse order
            vOPSFound = False
            'UPGRADE_NOTE: Object vOPS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            vOPS = Nothing
            Do
              vOPS = CType(mvEnv.GetPreviousItem(pTraderTransaction.PaymentPlan.ScheduledPayments, vOPS), OrderPaymentSchedule)
              If Not (vOPS Is Nothing) Then
                If vOPS.AmountOutstanding > 0 Then
                  vOPSFound = True
                End If
              End If
            Loop While vOPSFound = False And Not (vOPS Is Nothing)
            If vOPSFound = True And vUseRenewalAmount = False Then vFreqAmount = vOPS.AmountOutstanding
            'Set the payment amount
            If vCount > 1 And vUseRenewalAmount = True Then
              If vOPSFound Then vOPS.PaymentAmount = vAmount
              vAmount = 0
            Else
              If vAmount > vFreqAmount Then
                If vOPSFound Then vOPS.PaymentAmount = vFreqAmount
                vAmount = FixTwoPlaces(vAmount - vFreqAmount)
              Else
                If vOPSFound Then vOPS.PaymentAmount = vAmount
                vAmount = 0
              End If
            End If

            vTDRLine.AddPaymentPlanPayment(pTraderTransaction.PaymentPlan.PlanNumber, vNumber, vOPS.ScheduledPaymentNumber, CalculateCurrencyAmount(vOPS.PaymentAmount, Me.BatchCurrencyCode, Me.BatchExchangeRate, False), pParams("TransactionSource").Value, False, "", "", pTraderTransaction.PaymentPlan.GiverContactNumber, vType, "", pTraderTransaction.PaymentPlan.SalesContact.ToString)
            vTransTotal = FixTwoPlaces(vTransTotal + vOPS.PaymentAmount) 'Set the total
            vOPS.SetUnProcessedPayment(True, vOPS.PaymentAmount)
            vOPS.Save()
            'PopulateGridFromTALine vTDRLine

            If vAmount > 0 Then
              vTALine = vTALine + 1
              vTDRLine = pTraderTransaction.GetTraderAnalysisLine(vTALine)
              vTDRLine.Init(vTALine, vTransactionType)
            End If
            vCount = vCount + 1
          End While
        Else
          Select Case pParams("PPPaymentType").Value 'mvPayMethod2
            Case "STDO"
              vType = "SO"
            Case "DIRD"
              vType = "DD"
            Case "CCCA"
              vType = "CC"
            Case "NPAY"
              vType = "NP"
          End Select
          If vType.Length = 0 AndAlso pTraderTransaction.PaymentPlan.LoanStatus = PaymentPlan.ppYesNoCancel.ppYes Then vType = "NP"
          vTDRLine.AddNonPaymentLine(pTraderTransaction.PaymentPlan.PlanNumber, pParams("TransactionSource").Value, vType)
          'PopulateGridFromTALine vTDRLine
        End If
        pNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
        If vTransTotal = vTransAmount Then 'Totals are the same
          vFinished = True
          If Me.AppType = ApplicationType.atTransaction And pTraderTransaction.PaymentPlan.Existing = False Then
            'New PaymentPlan
            If Me.DisplayScheduledPayments = True And pTraderTransaction.PaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppNo And pTraderTransaction.PaymentPlan.ScheduledPayments.Count() > 1 Then
              'Display the Schedule
              pNextPage = TraderPage.TraderPageType.tpScheduledPayments
              vFinished = False
            End If
          End If

          If vFinished Then
            If pParams.ParameterExists("ExistingTransaction").Bool = False And Me.ConfirmAnalysis = False Then
              '
            Else
              pNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
            End If
          End If
        Else
          pNextPage = TraderPage.TraderPageType.tpTransactionAnalysisSummary
        End If
      End If
    End Sub

    Private Function GetFrequencyAmount(ByRef pParams As CDBParameters, ByRef pPPNumber As Integer) As Double
      Dim vOrderBalance As Double
      Dim vAmount As Double
      Dim vFixAmount As Double
      Dim vPF As New PaymentFrequency
      Dim vPP As PaymentPlan = Nothing

      If pPPNumber > 0 And (AppType = ApplicationType.atConversion And Not PayPlanConversionMaintenance And Not ConversionShowPPD) Then
        vPP = New PaymentPlan
        vPP.Init(mvEnv, pPPNumber)
      End If

      If pParams.Exists("PPD_PaymentFrequency") Then
        vPF.Init(mvEnv, (pParams("PPD_PaymentFrequency").Value))
      ElseIf pParams.Exists("PPM_PaymentFrequency") Then
        vPF.Init(mvEnv, (pParams("PPM_PaymentFrequency").Value))
      ElseIf pParams.Exists("LON_PaymentFrequency") Then
        vPF.Init(mvEnv, pParams("LON_PaymentFrequency").Value)
      ElseIf Not vPP Is Nothing Then
        vPF.Init(mvEnv, (vPP.PaymentFrequencyCode))
      End If

      Dim vIsLoan As Boolean = False
      If Not vPP Is Nothing Then
        vOrderBalance = vPP.Balance
        vIsLoan = (vPP.PlanType = CDBEnvironment.ppType.pptLoan)
      Else
        vOrderBalance = pParams("PPBalance").DoubleValue
        vIsLoan = (pParams.ParameterExists("TransactionType").Value.Equals("LOAN", StringComparison.CurrentCulture))
      End If
      vAmount = If(vIsLoan, vPP.FrequencyAmount, vOrderBalance / vPF.Frequency)
      vFixAmount = Int(vAmount * 100) / 100
      If vAmount > vFixAmount Then vFixAmount = vFixAmount + 0.01
      GetFrequencyAmount = vFixAmount
    End Function

    Private Sub SCAddMemberSummary(ByVal pParams As CDBParameters, ByVal pTraderTransaction As TraderTransaction, ByRef pResults As CDBParameters)
      'Adding a member to the MembershipMembersSummary grid
      Dim vContact As New Contact(mvEnv)
      Dim vMember As Member
      Dim vMembershipType As MembershipType
      Dim vMT As MembershipType
      Dim vRS As CDBRecordSet
      Dim vBranch As String
      Dim vNoMembers As Integer
      Dim vSQL As String

      vContact.Init()

      vSQL = "SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress) & " FROM contacts c, contact_addresses ca, addresses a"
      vSQL = vSQL & " WHERE c.contact_number = " & pParams("FinderContactNumber").IntegerValue
      vSQL = vSQL & " AND ca.contact_number = c.contact_number AND ca.address_number = c.address_number AND a.address_number = ca.address_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      If vRS.Fetch() = True Then vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress)
      vRS.CloseRecordSet()

      If vContact.ContactNumber > 0 Then
        vMembershipType = mvEnv.MembershipType((pParams("MembershipType").Value))

        vBranch = vContact.Address.Branch
        If Len(vBranch) = 0 Then
          If pParams.ParameterExists("TransactionType").Value = "MEMC" Then
            'Get Branch from original Membership
            If pTraderTransaction.SummaryMembers.Count > 0 Then
              vBranch = CType(pTraderTransaction.SummaryMembers(1), Member).Branch
            End If
          Else
            'Get Branch from Members page
            vBranch = pParams("Branch").Value
          End If
        End If

        For Each vMember In pTraderTransaction.SummaryMembers
          If vMember.MembershipTypeCode = vMembershipType.MembershipTypeCode Then
            vNoMembers = vNoMembers + 1
          End If
        Next vMember

        If (vNoMembers >= pParams("NumberOfMembers").IntegerValue) And vMembershipType.AssociateMembershipType.Length > 0 Then
          vMT = mvEnv.MembershipType((vMembershipType.AssociateMembershipType))
        Else
          vMT = vMembershipType
        End If

        With pTraderTransaction.PaymentPlan
          .PlanNumber = 0
          .AddMember((vContact.ContactNumber), vContact.Address.AddressNumber, (vMT.MembershipTypeCode), vMT, (vContact.ContactType))
          .Member.SCAddMemberSummary(vContact, pParams("Joined").Value, vBranch, vMT.BranchMembership, pParams("Joined").Value, pParams("DistributionCode").Value, "", vContact.DateOfBirth)
        End With

        pResults.Add("CurrentMembers", pParams("CurrentMembers").IntegerValue + 1)

      End If

    End Sub

    Private Sub PayScheduledPayment(ByRef pOPS As OrderPaymentSchedule, ByRef pBalance As Double)
      Dim vAmount As Double

      If pBalance > 0 Then
        If (pOPS.ScheduledPaymentStatus = OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsProvisional Or pOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance) Then
          'Pay the full amount if possible
          vAmount = pBalance
          If vAmount > pOPS.AmountOutstanding Then vAmount = pOPS.AmountOutstanding
        Else
          vAmount = pOPS.AmountOutstanding
        End If
        If vAmount > pBalance Then
          vAmount = pBalance
        End If

        'Reduce the amount oustanding on this line
        pOPS.AddPayment(vAmount)
        pBalance = pBalance - vAmount
        pOPS.SCCheckValue = True
      Else
        'There is nothing outstanding so uncheck the control
        pOPS.SCCheckValue = False
      End If

    End Sub

    Private Sub EditScheduledPaymentAnalysisLine(ByRef pTransaction As TraderTransaction, ByVal pTDRLine As TraderAnalysisLine, ByRef pParams As CDBParameters, ByVal pAddToCollection As Boolean)
      'Called from cmdDelete/cmdEdit/cmdPrevious clicks
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vOriginalOPS As OrderPaymentSchedule = Nothing
      Dim vFound As Boolean

      If pTDRLine.PaymentPlanNumber > 0 And pTDRLine.ScheduledPaymentNumber > 0 Then
        pTransaction.PaymentPlan.Init(mvEnv, (pTDRLine.PaymentPlanNumber))
        For Each vOPS In pTransaction.PaymentPlan.ScheduledPayments
          If vOPS.ScheduledPaymentNumber = pTDRLine.ScheduledPaymentNumber Then
            vFound = True
            Exit For
          End If
        Next vOPS
        If Not vFound Then
          'If not in collection, then add it as collection is current renewal period only
          vOPS = New OrderPaymentSchedule
          vOPS.Init(mvEnv, (pTDRLine.ScheduledPaymentNumber))
          If vOPS.Existing Then pTransaction.PaymentPlan.ScheduledPayments.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
          vFound = vOPS.Existing
        End If
        If vFound Then
          vOPS.PaymentAmount = pTDRLine.Amount 'CalcCurrencyAmount(vTDRLine.Amount, True)
          If pParams.ParameterExists("FinancialAdjustment").IntegerValue = Batch.AdjustmentTypes.atAdjustment Then
            'BR19606 - vOPS.ProcessReanalysis is going to change the OPS, so take a copy before and send it to Smart Client just incase the user cancels the change
            vOriginalOPS = New OrderPaymentSchedule
            vOriginalOPS.Init(mvEnv, (pTDRLine.ScheduledPaymentNumber))
            pTransaction.OriginalOPS = vOriginalOPS
            vOPS.ProcessReanalysis(vOPS.PaymentAmount * -1)
          Else
            vOPS.SetUnProcessedPayment(False, (vOPS.PaymentAmount * -1), (pTransaction.PaymentPlan.PlanType = CDBEnvironment.ppType.pptLoan))
          End If
          vOPS.Save(mvEnv.User.UserID)

          If pAddToCollection Then
            'Add the payment to a collection so that it can be re-instated if the user clicks cancel
            If pTransaction.RemovedSchPayments.Exists(CStr(vOPS.ScheduledPaymentNumber)) = False Then
              pTransaction.RemovedSchPayments.Add(vOPS, CStr(vOPS.ScheduledPaymentNumber))
            End If
          End If
        End If
      End If
    End Sub

    Public Function IssueStock(ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer, ByVal pLineNumber As Integer, ByVal pIssueStock As Boolean, ByVal pResetStockIssued As Boolean, ByRef pStockMovementNumbers As String, Optional ByVal pPlaceOnBackOrder As Boolean = False, Optional ByVal pProductCostNumber As Integer = 0) As Integer
      'pStockMovementNumbers returns list of StockMovementNumber's created
      'pProductCostNumber is used to specify the ProductCost record to be used when Trader is putting stock back (e.g. because user has changed the Product or Warehouse), in all other circumstances the ProductCostNumber should be zero so that the stock is issued from the earliest Product Cost records
      Dim vRS As CDBRecordSet
      Dim vProductCosts As New ProductCosts 'Collection of ProductCost objects
      Dim vProductCost As ProductCost
      Dim vStockMovement As StockMovement
      Dim vBatchNumber As Integer
      Dim vTransactionNumber As Integer
      Dim vMultiplier As Integer
      Dim vQuantity As Integer
      Dim vRemaining As Integer
      Dim vReasonCode As String
      Dim vStockIssued As Integer 'The quantity of stock that this procedure has issued

      If pPlaceOnBackOrder Then
        vReasonCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonAwaitBackOrder)
      Else
        vReasonCode = If(pIssueStock = True, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonSale), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonReversal))
      End If

      vBatchNumber = BatchNumber
      vTransactionNumber = TransactionNumber
      If vBatchNumber < 0 Then vBatchNumber = 0
      If vTransactionNumber < 0 Then vTransactionNumber = 0
      vMultiplier = 1
      If pQuantity < 0 Then vMultiplier = -1

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) = False Then
        'Database does not support ProductCosts so create one StockMovement
        vStockMovement = New StockMovement
        vStockMovement.Create(mvEnv, pProductCode, pQuantity, vReasonCode, vBatchNumber, vTransactionNumber, pLineNumber, False, pWarehouseCode, 0, mvStockMovementTransactionID)
        If vStockMovement.Existing = True Then pStockMovementNumbers = CStr(vStockMovement.StockMovementNumber)
        vStockIssued = vStockMovement.MovementQuantity
      Else
        'Using ProductCosts so create one StockMovement per ProductCost
        If Len(pWarehouseCode) = 0 Then
          'Must have Warehouse code
          vRS = mvEnv.Connection.GetRecordSet("SELECT warehouse FROM product WHERE product = '" & pProductCode & "'")
          If vRS.Fetch() = True Then pWarehouseCode = vRS.Fields(1).Value
          vRS.CloseRecordSet()
        End If
        vStockIssued = 0
        If pProductCostNumber > 0 And pQuantity > 0 Then
          'User has changed Product or Warehouse so put stock back against original ProductCost
          vStockMovement = New StockMovement
          vStockMovement.Create(mvEnv, pProductCode, pQuantity, vReasonCode, vBatchNumber, vTransactionNumber, pLineNumber, False, pWarehouseCode, pProductCostNumber, mvStockMovementTransactionID)
          If vStockMovement.Existing = True Then pStockMovementNumbers = CStr(vStockMovement.StockMovementNumber)
          vStockIssued = vStockMovement.MovementQuantity
        End If
        If vStockIssued = 0 Then
          vProductCosts.InitFromProductAndWarehouse(mvEnv, pProductCode, pWarehouseCode)
          vRemaining = pQuantity
          'StockMovements will be created for each ProductCost record that has outstanding stock
          For Each vProductCost In vProductCosts
            If (vProductCost.LastStockCount + vRemaining) >= 0 Then
              vQuantity = vRemaining
            Else
              vQuantity = (vProductCost.LastStockCount * vMultiplier)
            End If
            If vQuantity = 0 Then
              'No productCost as going onto back order
              vQuantity = vRemaining
              vStockMovement = New StockMovement
              vStockMovement.Create(mvEnv, pProductCode, vQuantity, vReasonCode, vBatchNumber, vTransactionNumber, pLineNumber, False, pWarehouseCode, 0, mvStockMovementTransactionID)
              If vStockMovement.Existing = True Then pStockMovementNumbers = CStr(vStockMovement.StockMovementNumber)
              vStockIssued = vStockMovement.MovementQuantity
              vQuantity = vRemaining
            Else
              vStockMovement = New StockMovement
              vStockMovement.Create(mvEnv, pProductCode, vQuantity, vReasonCode, vBatchNumber, vTransactionNumber, pLineNumber, False, pWarehouseCode, (vProductCost.ProductCostNumber), mvStockMovementTransactionID)
              vStockIssued = vStockIssued + vStockMovement.MovementQuantity
              If vStockMovement.Existing = True Then
                If pStockMovementNumbers.Length > 0 Then pStockMovementNumbers = pStockMovementNumbers & ","
                pStockMovementNumbers = pStockMovementNumbers & CStr(vStockMovement.StockMovementNumber)
              End If
            End If
            vRemaining = vRemaining - vQuantity
            If vRemaining = 0 Then Exit For
          Next vProductCost
        End If
      End If
      IssueStock = vStockIssued
    End Function

    Public Sub InitForStockMovements(ByVal pEnv As CDBEnvironment, ByVal pStockMovementTransactionID As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      'Used by ProcessStockMovements Web Service only (we do not know, or need to know, the actual TraderApplication)
      Init("", pEnv, pBatchNumber, pTransactionNumber)
      mvStockMovementTransactionID = pStockMovementTransactionID
      mvBackOrderPrompt = mvEnv.GetConfigOption("opt_fp_back_order_prompt")
    End Sub

    Public Function SCSumStockMovements(ByVal pStockMovementTransactionID As Integer, Optional ByVal pExistingTransaction As Boolean = False, Optional ByVal pDeleteStockMovements As Boolean = False) As CDBDataTable ', Optional ByVal pBatchNumber As Long = 0, Optional ByVal pTransactionNumber As Long = 0, Optional ByVal pLineNumber As Long = 0) As CDBDataTable
      'Select all StockMovements for the TransactionID and add each unique combination of Product/Warehouse/ProductCostNumber to the DataTable
      'pExistingTransaction will be passed in when only new StockMovements added to an existing transaction are required
      'BR20529 - Rewritten to remove hard coded 2 character limit on Warehouse and concatenated keys.  
      Dim vRS As CDBRecordSet
      Dim vDT As New CDBDataTable
      Dim vSM As New StockMovement
      Dim vTrans As Boolean
      Dim vSQL As SQLStatement
      Dim vFields As String
      Dim vWhereFields As New CDBFields
      Dim vTableName As String = "stock_movements sm"
      Dim vOrderBy As String = "product, warehouse, product_cost_number"


      vDT.AddColumnsFromList("Product,Warehouse,MovementQuantity,ProductCostNumber")
      Dim vDataTable As DataTable = vDT.ConvertToDataTable()
      vSM.Init(mvEnv)
      vFields = vSM.GetRecordSetFields(StockMovement.StockMovementRecordSetTypes.smrtAll)
      vWhereFields.Add("transaction_id", CDBField.FieldTypes.cftInteger, pStockMovementTransactionID)
      If pExistingTransaction Then
        vWhereFields.Add("batch_number", "") 'Only delete StockMovements that are not linked to an existing transaction
      End If
      vSQL = New SQLStatement(mvEnv.Connection, vFields, vTableName, vWhereFields, vOrderBy)
      vRS = vSQL.GetRecordSet
      While vRS.Fetch() = True
        vSM = New StockMovement
        vSM.InitFromRecordSet(mvEnv, vRS, StockMovement.StockMovementRecordSetTypes.smrtAll)

        Dim vDataRows As DataRow() = vDataTable.Select("Product = '" & vSM.ProductCode & "' AND Warehouse= '" & vSM.Warehouse & "' AND ProductCostNumber = '" & vSM.ProductCostNumber & "'")
        Select Case vDataRows.Length
          Case 0
            Dim vNewRow As DataRow = vDataTable.NewRow()
            vNewRow("Product") = vSM.ProductCode
            vNewRow("Warehouse") = vSM.Warehouse
            vNewRow("ProductCostNumber") = vSM.ProductCostNumber
            vNewRow("MovementQuantity") = (vSM.MovementQuantity * -1).ToString()
            vDataTable.Rows.Add(vNewRow)
          Case 1
            vDataRows(0)("MovementQuantity") = (CInt(vDataRows(0)("MovementQuantity")) + (vSM.MovementQuantity * -1)).ToString()
          Case Else
            ' Error
        End Select
        If pDeleteStockMovements = True Then
          If mvEnv.Connection.InTransaction = False Then
            mvEnv.Connection.StartTransaction()
            vTrans = True
          End If
          vSM.Delete()
        End If
      End While

      vRS.CloseRecordSet()
      If vTrans Then mvEnv.Connection.CommitTransaction()
      vDataTable.AcceptChanges()
      vDT = New CDBDataTable(vDataTable)
      Return vDT

    End Function

    Public Function SCRemoveStockMovements(ByVal pDT As CDBDataTable, ByVal pLineNumber As Integer, ByVal pUpdateStockLevels As Boolean, ByVal pIssueNewStock As Boolean) As Boolean
      'Used by SmartClient Trader and WebServices to remove StockMovements
      Dim vDR As CDBDataRow
      Dim vSM As StockMovement
      Dim vSMNumbers As String = ""
      Dim vStockIssued As Integer
      Dim vStockUpdated As Boolean

      vStockUpdated = Not (pUpdateStockLevels)
      If pUpdateStockLevels Then
        If pIssueNewStock Then
          'Always IssueStock
          For Each vDR In pDT.Rows
            vStockIssued = vStockIssued - Me.IssueStock(vDR.Item("Product"), vDR.Item("Warehouse"), vDR.LongItem("MovementQuantity"), pLineNumber, True, False, vSMNumbers, False, vDR.LongItem("ProductCostNumber"))
            vStockUpdated = True
          Next vDR
        Else
          'Delete StockMovements for a new record, just update the Stock Levels
          vSM = New StockMovement
          vSM.Init(mvEnv)
          For Each vDR In pDT.Rows
            If vDR.LongItem("MovementQuantity") <> 0 Then vSM.UpdateStockLevels(mvEnv, vDR.Item("Product"), vDR.Item("Warehouse"), vDR.LongItem("MovementQuantity"), vDR.LongItem("ProductCostNumber"))
            vStockUpdated = True
          Next vDR
        End If
      End If

      SCRemoveStockMovements = vStockUpdated

    End Function

    Private Sub CreatePreTaxPGPayment(ByVal pParams As CDBParameters, ByRef pTransaction As TraderTransaction)
      'Create Pre Tax Payroll Giving batches
      Dim vPledge As New PreTaxPledge(mvEnv)
      Dim vTDRLine As TraderAnalysisLine
      Dim vEmployerProduct As String
      Dim vEmployerRate As String
      Dim vGovProduct As String
      Dim vGovRate As String
      Dim vAdminProduct As String
      Dim vAdminRate As String
      Dim vAmount As Double
      Dim vLineCount As Integer
      Dim vDonorTotal As Double
      Dim vEmployerTotal As Double
      Dim vGovernmentTotal As Double
      Dim vAdminFeesTotal As Double
      Dim vPostCashBook As Boolean

      vPledge.Init((pParams("GayePledgeNumber").IntegerValue))

      vEmployerProduct = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEEmployerProduct)
      vEmployerRate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEEmployerRate)
      vGovProduct = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEGovernmentProduct)
      vGovRate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEGovernmentRate)
      vAdminProduct = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEAdminFeeProduct)
      vAdminRate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAYEAdminFeeRate)

      vLineCount = 1
      vDonorTotal = pParams("DonorTotal").DoubleValue
      vEmployerTotal = pParams("EmployerTotal").DoubleValue
      vGovernmentTotal = pParams("GovernmentTotal").DoubleValue
      vAdminFeesTotal = pParams("AdminFeesTotal").DoubleValue

      vAmount = vDonorTotal
      If Len(vEmployerProduct) > 0 And ((vPledge.ProductCode <> vEmployerProduct) Or (vPledge.RateCode <> vEmployerRate)) Then
        If vEmployerTotal <> 0 Then
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vLineCount = vLineCount + 1
          vTDRLine.AddPreTaxPGPayment(vEmployerProduct, vEmployerRate, vEmployerTotal, pParams("Source").Value, vPledge.DistributionCode, (vPledge.GayePledgeNumber))
        End If
      Else
        vAmount = vAmount + vEmployerTotal
      End If
      If Len(vGovProduct) > 0 And ((vPledge.ProductCode <> vGovProduct) Or (vPledge.RateCode <> vGovRate)) Then
        If vGovernmentTotal <> 0 Then
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vLineCount = vLineCount + 1
          vTDRLine.AddPreTaxPGPayment(vGovProduct, vGovRate, vGovernmentTotal, pParams("Source").Value, vPledge.DistributionCode, (vPledge.GayePledgeNumber))
        End If
      Else
        vAmount = vAmount + vGovernmentTotal
      End If
      If Len(vAdminProduct) > 0 And ((vPledge.ProductCode <> vAdminProduct) Or (vPledge.RateCode <> vAdminRate)) Then
        If vAdminFeesTotal <> 0 Then
          vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
          vLineCount = vLineCount + 1
          vTDRLine.AddPreTaxPGPayment(vAdminProduct, vAdminRate, vAdminFeesTotal, pParams("Source").Value, vPledge.DistributionCode, (vPledge.GayePledgeNumber))
        End If
      Else
        vAmount = vAmount + vAdminFeesTotal
      End If

      vTDRLine = pTransaction.GetTraderAnalysisLine(vLineCount)
      vTDRLine.AddPreTaxPGPayment(vPledge.ProductCode, vPledge.RateCode, vAmount, pParams("Source").Value, vPledge.DistributionCode, (vPledge.GayePledgeNumber))

      'Add Transaction
      vPostCashBook = pParams("PostCashBook").Bool
      pTransaction.InitPreTaxPayrollGiving(mvEnv, mvEnv.GetConfig("pm_gaye"), SOBankAccount, pParams("ContactNumber").IntegerValue, pParams("AddressNumber").IntegerValue, pParams("TransactionDate").Value, pParams("Reference").Value, vPostCashBook, False, Batch.AdjustmentTypes.atNone, BatchCurrencyCode, CStr(BatchExchangeRate), (pParams("Mailing").Value), "", mvBatchNumber, mvTransNumber)
      'Now Save everything
      mvEnv.Connection.StartTransaction()
      pTransaction.TraderAnalysisLines.SaveAnalysis(mvEnv, pTransaction.BatchTransaction, 0, Batch.AdjustmentTypes.atNone, False, False, False, False, BatchCurrencyCode, BatchExchangeRate)
      'Write payment history
      Dim vPaymentParams As New CDBParameters()
      vPaymentParams.Add("DonorAmount", pParams("DonorTotal").DoubleValue)
      vPaymentParams.Add("EmployerAmount", pParams("EmployerTotal").DoubleValue)
      vPaymentParams.Add("GovernmentAmount", pParams("GovernmentTotal").DoubleValue)
      vPaymentParams.Add("AdminFeeAmount", pParams("AdminFeesTotal").DoubleValue)
      vPaymentParams.Add("OtherMatchedAmount", CDBField.FieldTypes.cftNumeric, "0")
      vPaymentParams.Add("BatchNumber", pTransaction.BatchNumber)
      vPaymentParams.Add("TransactionNumber", pTransaction.TransactionNumber)
      vPledge.AddPayment(vPaymentParams)
      vPledge.Save(mvEnv.User.UserID, True)

      pTransaction.SaveTransaction(pTransaction.BatchTransaction.CurrencyAmount)
      'Note: this will leave an open Batch
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub CreatePostTaxPGPayment(ByVal pParams As CDBParameters, ByRef pTransaction As TraderTransaction)
      'Create Post Tax Payroll Giving batches
      Dim vTDRLine As TraderAnalysisLine
      Dim vDistribCode As String
      Dim vPostCashBook As Boolean
      Dim vProductCode As String
      Dim vRate As String
      Dim vPledge As New PostTaxPledge(mvEnv)

      vPledge.Init((pParams("PledgeNumber").IntegerValue))
      'Set values for the Donor analysis line
      If vPledge.Existing Then
        'Use Product/Rate/DistributionCode from Pledge
        vProductCode = vPledge.ProductCode
        vRate = vPledge.RateCode
        vDistribCode = vPledge.DistributionCode
      Else
        'Use Product/Rate/DistributionCode from ControlTable
        vProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGDonorProduct)
        vRate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGDonorRate)
        vDistribCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGDistributionCode)
      End If
      'Always add Donor analysis line
      vTDRLine = pTransaction.GetTraderAnalysisLine(1)
      vTDRLine.Init(1)
      vTDRLine.AddPostTaxPGPayment(vProductCode, vRate, pParams("DonorTotal").DoubleValue, pParams("Source").Value, vDistribCode, (vPledge.PledgeNumber))
      'Add Employer analysis line if amount set
      If pParams("EmployerTotal").DoubleValue > 0 Then
        'Set values for Employer analysis line
        vProductCode = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGEmployerProduct)
        vRate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGEmployerRate)
        vTDRLine = pTransaction.GetTraderAnalysisLine(2)
        vTDRLine.Init(2)
        vTDRLine.AddPostTaxPGPayment(vProductCode, vRate, pParams("EmployerTotal").DoubleValue, pParams("Source").Value, vDistribCode)
      End If
      'Add Transaction
      vPostCashBook = pParams("PostCashBook").Bool
      pTransaction.InitPostTaxPayrollGiving(mvEnv, mvEnv.GetConfig("pm_gaye"), SOBankAccount, pParams("ContactNumber").IntegerValue, pParams("AddressNumber").IntegerValue, pParams("TransactionDate").Value, pParams("Reference").Value, vPostCashBook, False, Batch.AdjustmentTypes.atNone, BatchCurrencyCode, CStr(BatchExchangeRate), (pParams("Mailing").Value), "", mvBatchNumber, mvTransNumber)
      'Now Save everything
      mvEnv.Connection.StartTransaction()
      pTransaction.TraderAnalysisLines.SaveAnalysis(mvEnv, pTransaction.BatchTransaction, 0, Batch.AdjustmentTypes.atNone, False, False, False, False, BatchCurrencyCode, BatchExchangeRate)
      pTransaction.SaveTransaction(pTransaction.BatchTransaction.CurrencyAmount)
      'Note: this will leave an open Batch
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Function SetPPDEffectiveDate(ByVal pConversionType As String, ByVal pPaymentFrequencyCode As String, ByVal pPaymentPlan As PaymentPlan, Optional ByVal pUseTraderPPDetails As Boolean = False) As String
      'Set the PaymentPlanDetail EffectiveDate (used by Rich & Smart Client)
      'Assume we are on the tpPaymentPlanDetailsMaintenance page
      Dim vPaymentFrequency As PaymentFrequency
      Dim vEffectiveDate As String
      Dim vGotAPM As Boolean
      Dim vDetail As PaymentPlanDetail
      Dim vPPDGracePeriodProportion As String

      vEffectiveDate = ""
      If pPaymentPlan.Existing = True And (pPaymentPlan.ProportionalBalanceSetting And PaymentPlan.ProportionalBalanceConfigSettings.pbcsExisting) = PaymentPlan.ProportionalBalanceConfigSettings.pbcsExisting Then
        vPPDGracePeriodProportion = mvEnv.GetConfig("fp_ppd_grace_period_proportion", "N")
        If vPPDGracePeriodProportion <> "N" And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPPDetailsEffectiveDate) = True Then
          'PaymentFrequency might have changed, so use the new value from the PaymentPlanMaintenance (PPM) page in preference to value on PaymentPlan
          If Len(pPaymentFrequencyCode) = 0 Then pPaymentFrequencyCode = pPaymentPlan.PaymentFrequencyCode
          vGotAPM = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes)
          If vGotAPM = False And Me.AppType = ApplicationType.atConversion Then
            'Might be adding a DD/CCCA
            vGotAPM = (pConversionType = "DIRD" Or pConversionType = "CCCA")
          End If
          vPaymentFrequency = mvEnv.GetPaymentFrequency(pPaymentFrequencyCode)

          If vGotAPM = True And vPaymentFrequency.Frequency > 1 And vPaymentFrequency.Period = PaymentFrequency.PaymentFrequencyPeriods.pfpMonths And (vPaymentFrequency.Frequency * vPaymentFrequency.Interval <= 12) Then
            'Existing PaymentPlan paid by DD/CCCA, paid in instalments and annually renewing
            If vPPDGracePeriodProportion = "Y" Then
              Dim vAutoPayMethod As PaymentPlan.ppAutoPayMethods
              If pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes OrElse pConversionType = "DIRD" Then
                vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMDD
              ElseIf pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes OrElse pConversionType = "CCCA" Then
                vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMCCCA
              End If
              vEffectiveDate = mvEnv.GetPaymentPlanAutoPayDate(Today, vAutoPayMethod, pPaymentPlan.AutoPayBankAccount).ToString(CAREDateFormat)
            ElseIf vPPDGracePeriodProportion = "R" Then
              If CDate(pPaymentPlan.RenewalDate) > CDate(TodaysDate()) Then
                vEffectiveDate = pPaymentPlan.RenewalDate
              Else
                vEffectiveDate = TodaysDate()
              End If
            End If
            Dim vPPDetails As Collection
            If pUseTraderPPDetails Then
              vPPDetails = PaymentPlanDetails.Items
            Else
              vPPDetails = pPaymentPlan.Details
            End If
            For Each vDetail In vPPDetails
              If IsDate(vDetail.EffectiveDate) Then
                If CDate(vEffectiveDate) < CDate(vDetail.EffectiveDate) Then
                  vEffectiveDate = vDetail.EffectiveDate
                End If
              End If
            Next vDetail
          End If
        End If
      End If

      SetPPDEffectiveDate = vEffectiveDate

    End Function

    Public Function SCDoDeletePaymentPlanLine(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction) As TraderPage.TraderPageType
      Dim vSubscriptionNumber As Integer
      Dim vCancellationReason As String
      Dim vCancellationSource As String
      Dim vNextPage As TraderPage.TraderPageType

      Select Case pCurrentPageType
        Case TraderPage.TraderPageType.tpPaymentPlanSummary
          vNextPage = TraderPage.TraderPageType.tpPaymentPlanSummary

          'To Cancel Subsciption
          vSubscriptionNumber = CInt(pParams.OptionalValue("SubscriptionNumber", ""))
          If vSubscriptionNumber > 0 And vSubscriptionNumber = pTransaction.TraderPPDLines(1).SubscriptionNumber Then
            vCancellationReason = pParams("CancellationReason").Value
            vCancellationSource = pParams.OptionalValue("CancellationSource", "")
            pTransaction.PaymentPlan.CancelSubscription(vSubscriptionNumber, pTransaction.TraderPPDLines(1).DetailNumber, vCancellationReason, vCancellationSource)
          End If

      End Select

      SCDoDeletePaymentPlanLine = vNextPage
    End Function

    Private Sub SCSavePurchaseOI(ByRef pCurrentPageType As TraderPage.TraderPageType, ByRef pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pResults As CDBParameters)
      Dim vPI As PurchaseInvoice = Nothing
      Dim vPOD As PurchaseOrderDetail
      Dim vPID As PurchaseInvoiceDetail
      Dim vPPA As PurchaseOrderPayment
      Dim vEditMode As Boolean
      Dim vPODAmount As Double = 0
      Dim vPOBalance As Double = 0

      Dim vPO As New PurchaseOrder(mvEnv)
      If MainPage = TraderPage.TraderPageType.tpPurchaseOrderDetails Then
        vPO.Init(pParams.ParameterExists("PurchaseOrderNumber").IntegerValue)
        vEditMode = vPO.Existing
        'Check for Purchase Order Details
        For Each vPOD In pTransaction.PurchaseOrderDetails
          vPO.AddDetail(vPOD)
          If vEditMode Then vPODAmount += vPOD.Amount()
        Next vPOD
        If vEditMode Then
          'Jira 664: Calculate Balance from previous Balance plus difference between new PO Details Amount and old PO Amount
          vPOBalance = (vPO.Balance + (vPODAmount - vPO.Amount))
          pParams("PPBalance").Value = vPOBalance.ToString
        End If
        'Check for Purchase Order Payments
        If pParams.ParameterExists("NumberOfPayments").IntegerValue > 0 Then
          For Each vPPA In pTransaction.PurchaseOrderPayments
            vPO.AddPayment(vPPA)
          Next vPPA
        End If
      Else
        vPI = New PurchaseInvoice(mvEnv)
        vPI.Init(pParams.ParameterExists("PurchaseInvoiceNumber").IntegerValue)
        vEditMode = vPI.Existing
        'Check for Purchase Invoice Details
        For Each vPID In pTransaction.PurchaseInvoiceDetails
          vPI.AddDetail(vPID)
        Next vPID
      End If

      ' Now set up the purchase order/invoice
      mvEnv.Connection.StartTransaction()
      If MainPage = TraderPage.TraderPageType.tpPurchaseOrderDetails Then
        vPO.CreateFromTrader(pParams)
        vPO.SaveWithDetails(vEditMode, mvEnv.User.Logname, True)
      Else
        vPI.CreateFromTrader(pParams)
        vPI.SaveWithDetails(vEditMode)
      End If
      mvEnv.Connection.CommitTransaction()
      If Not vEditMode And ShowTransReference Then
        If MainPage = TraderPage.TraderPageType.tpPurchaseOrderDetails Then
          pResults.Add("PurchaseOrderNumber", vPO.PurchaseOrderNumber)
        Else
          pResults.Add("PurchaseInvoiceNumber", vPI.PurchaseInvoiceNumber)
        End If
      End If
    End Sub

    Private Sub CancelPOs(ByVal pParams As CDBParameters)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCount As Integer
      Dim vOrderNumberFrom As Integer
      Dim vOrderNumberTo As Integer

      vOrderNumberFrom = pParams("PurchaseOrderNumber").IntegerValue
      vOrderNumberTo = pParams("PurchaseOrderNumber2").IntegerValue
      If vOrderNumberTo > 0 Then
        vWhereFields.Add("purchase_order_number", vOrderNumberFrom, CDBField.FieldWhereOperators.fwoBetweenFrom)
        vWhereFields.Add("purchase_order_number2", vOrderNumberTo, CDBField.FieldWhereOperators.fwoBetweenTo)
      Else
        vWhereFields.Add("purchase_order_number", CDBField.FieldTypes.cftLong, vOrderNumberFrom)
      End If

      If mvEnv.Connection.GetCount("purchase_orders", vWhereFields) = 0 Then
        RaiseError(DataAccessErrors.daeTraderApplicationInvalid, (ProjectText.String15265)) 'No Records found in the specified range
      Else
        vCount = mvEnv.Connection.GetCount("purchase_invoices", vWhereFields)
        If vCount = 0 Then
          vWhereFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
          vCount = mvEnv.Connection.GetCount("purchase_orders", vWhereFields)
          If vCount > 0 Then
            RaiseError(DataAccessErrors.daeTraderApplicationInvalid, String.Format(ProjectText.String29031, CStr(vCount))) '%s of these records have already been cancelled
          Else
            Try
              mvEnv.Connection.StartTransaction()
              For vPoNumber As Integer = vOrderNumberFrom To vOrderNumberTo
                Dim vPurchaseOrders As New PurchaseOrder(mvEnv)
                vPurchaseOrders.Init(vPoNumber)
                If Not vPurchaseOrders.Existing Then Continue For
                If Not pParams.ContainsKey("CancelledBy") Then pParams.Add("CancelledBy", mvEnv.User.Logname)
                vPurchaseOrders.Update(pParams)
                vPurchaseOrders.Save(mvEnv.User.Logname, True, 0, True)
              Next
              mvEnv.Connection.CommitTransaction()
            Catch vEx As Exception
              ' If there is any type of exception then make sure the transactions are rolled back
              If mvEnv.Connection.InTransaction Then mvEnv.Connection.RollbackTransaction()
            End Try
          End If
        Else
          RaiseError(DataAccessErrors.daeTraderApplicationInvalid, String.Format(ProjectText.String29032, CStr(vCount))) '%s of these records have already been processed and cannot be cancelled
        End If
      End If
    End Sub

    Private Sub ChequeNoAlloc(ByVal pParams As CDBParameters, ByVal pResults As CDBParameters)
      Dim vFirstRef As Integer
      Dim vLastRef As Integer
      Dim vFirstNo As Integer
      Dim vLastNo As Integer
      Dim vNumber As Integer
      Dim vChqNo As Integer
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      vFirstRef = pParams("ChequeReferenceNumber").IntegerValue
      vLastRef = pParams("ChequeReferenceNumber2").IntegerValue
      vFirstNo = pParams("ChequeNumber").IntegerValue
      vLastNo = pParams("ChequeNumber2").IntegerValue
      If vLastRef < vFirstRef Or vLastNo < vFirstNo Then
        RaiseError(DataAccessErrors.daeTraderApplicationInvalid, (ProjectText.String29018)) 'Invalid range
      ElseIf vLastRef - vFirstRef <> vLastNo - vFirstNo Then
        RaiseError(DataAccessErrors.daeTraderApplicationInvalid, (ProjectText.String15267)) 'Range of Cheque Numbers does not match range of Reference Numbers
      Else
        vNumber = mvEnv.Connection.GetCount("cheques", Nothing, "cheque_reference_number BETWEEN " & vFirstRef & " AND " & vLastRef & " AND cheque_number IS NULL")
        If vNumber <> (vLastRef - vFirstRef) + 1 Then
          RaiseError(DataAccessErrors.daeTraderApplicationInvalid, (ProjectText.String15268)) 'Some Cheque records in this range are not valid or have already been assigned Cheque Numbers
        Else
          mvEnv.Connection.StartTransaction()
          vWhereFields.Add("cheque_reference_number", CDBField.FieldTypes.cftLong)
          vUpdateFields.AddAmendedOnBy(mvEnv.User.Logname)
          vUpdateFields.Add("cheque_number", CDBField.FieldTypes.cftLong)
          vChqNo = vFirstNo
          For vNumber = vFirstRef To vLastRef
            vUpdateFields(3).Value = CStr(vChqNo)
            vWhereFields(1).Value = CStr(vNumber)
            mvEnv.Connection.UpdateRecords("cheques", vUpdateFields, vWhereFields)
            vChqNo = vChqNo + 1
          Next
          mvEnv.Connection.CommitTransaction()
          pResults.Add("TotalCheques", (vLastRef - vFirstRef) + 1)
        End If
      End If
    End Sub

    Private Sub ChequeReconcile(ByVal pParams As CDBParameters)
      Dim vCheque As New Cheque(mvEnv)
      Dim vChequeNumber As Integer

      vChequeNumber = pParams("ChequeNumber").IntegerValue
      vCheque.InitFromChequeNumber(vChequeNumber)
      If vCheque.Existing Then
        mvEnv.Connection.StartTransaction()
        vCheque.Reconcile(pParams("ReconciledOn").Value, pParams("ChequeStatus").Value)
        vCheque.Save()
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Private Function CountChoices(ByVal pType As TraderPage.TraderPageType) As Integer
      Dim vCount As Integer
      If mvAddActivity Then vCount = vCount + 1
      If mvSetStatus Then vCount = vCount + 1
      If mvAddSuppression Then vCount = vCount + 1
      If mvGiftAidDeclaration Then vCount = vCount + 1
      If mvPayrollGiving Then vCount = vCount + 1
      If mvGoneAway Then vCount = vCount + 1
      If mvAddressMaintenance Then vCount = vCount + 1
      If mvAutoPaymentMaintenance Then vCount = vCount + 1
      CountChoices = vCount
    End Function

    Public Sub ProcessVATRateChange(ByVal pTransaction As TraderTransaction, ByVal pTransactionDate As String)
      'Used by Rich & Smart Client Trader to update VAT rates when the TransactionDate changes
      pTransaction.TraderAnalysisLines.UpdateVATRates(mvEnv, pTransactionDate)
    End Sub

    Private Sub ReAllocateProductNumber(ByVal pProductCode As String, ByVal pProductNumber As Integer)
      'Allow product number to be re-used by adding to product_numbers table
      Dim vFields As New CDBFields

      If pProductCode.Length > 0 And pProductNumber > 0 Then
        vFields.Add("product", CDBField.FieldTypes.cftCharacter, pProductCode)
        vFields.Add("product_number", CDBField.FieldTypes.cftLong, pProductNumber)
        mvEnv.Connection.InsertRecord("product_numbers", vFields, True)
      End If

    End Sub

    Private Function GetAdjustTransType(ByVal pAdjustment As Batch.AdjustmentTypes) As String
      'Get Transaction Type to use for Financial Adjustment
      Dim vWhere As String
      Dim vOrigTransactionSign As String
      Dim vNewTransactionType As String
      Dim vRecordSet As CDBRecordSet

      Select Case Batch.BatchType
        Case Batch.BatchTypes.CreditSales
          vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSCreditTransType)
        Case Else
          If pAdjustment = Batch.AdjustmentTypes.atAdjustment Then
            vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAdjustmentTransType)
          Else
            vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReverseTransType)
          End If
      End Select

      If Len(vNewTransactionType) > 0 Then
        vWhere = "transaction_type = '" & vNewTransactionType & "'"
      Else
        'Use first valid opposite Transaction Type
        vOrigTransactionSign = mvEnv.Connection.GetValue("SELECT transaction_sign from transaction_types WHERE transaction_type = '" & Batch.TransactionType & "'")
        vWhere = "transaction_sign = '"
        If pAdjustment = Batch.AdjustmentTypes.atAdjustment Then
          vWhere = vWhere & vOrigTransactionSign
        Else
          vWhere = vWhere & If(vOrigTransactionSign = "C", "D", "C")
        End If
        vWhere = vWhere & "' AND negatives_allowed = 'Y'"
      End If

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT transaction_type,transaction_sign FROM transaction_types WHERE " & vWhere)
      With vRecordSet
        If .Fetch() = True Then
          vNewTransactionType = .Fields.Item(1).Value
        Else
          RaiseError(DataAccessErrors.daeCannotFindAdjTransType)
        End If
        .CloseRecordSet()
      End With

      GetAdjustTransType = vNewTransactionType
    End Function

    Public Sub GetPayMethod(ByVal pPayMethod As String, ByRef pOrigPaymentMethod As String, ByRef pNewPaymentMethod As String, Optional ByVal pFinancialAdjustment As Batch.AdjustmentTypes = Batch.AdjustmentTypes.atNone)
      Dim vBatchType As Batch.BatchTypes
      If Len(mvBatchType) > 0 Then vBatchType = Access.Batch.GetBatchType(mvBatchType)
      If (pFinancialAdjustment = Batch.AdjustmentTypes.atAdjustment Or pFinancialAdjustment = Batch.AdjustmentTypes.atMove) And vBatchType = Batch.BatchTypes.FinancialAdjustment And (pPayMethod = mvEnv.GetConfig("pm_dd") Or pPayMethod = mvEnv.GetConfig("pm_ccca") Or pPayMethod = mvEnv.GetConfig("pm_so")) Then
        pOrigPaymentMethod = pPayMethod
        pNewPaymentMethod = "CASH"
        If pPayMethod = mvEnv.GetConfig("pm_so") And Me.AppType <> ApplicationType.atCreditListReconciliation Then
          pNewPaymentMethod = pPayMethod
        End If
      ElseIf pFinancialAdjustment = Batch.AdjustmentTypes.atCashBatchConfirmation Then
        pOrigPaymentMethod = pPayMethod
        pNewPaymentMethod = "CASH"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_cash") Then
        pNewPaymentMethod = "CASH"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_cheque") Then
        pNewPaymentMethod = "CHEQ"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_po") Then
        pNewPaymentMethod = "POST"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_so") And Me.AppType = ApplicationType.atCreditListReconciliation Then
        pNewPaymentMethod = "CASH" 'Force Credit List Reconciliation to behave as Cash Batch Maint
      ElseIf pPayMethod = mvEnv.GetConfig("pm_so") Then
        pNewPaymentMethod = pPayMethod
      ElseIf pPayMethod = Me.CSPayMethod Then
        pNewPaymentMethod = "CRED"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_cc") Or pPayMethod = mvEnv.GetConfig("pm_dc") Then
        pNewPaymentMethod = "CARD"
        If pFinancialAdjustment <> Batch.AdjustmentTypes.atNone Then pOrigPaymentMethod = pPayMethod
      ElseIf pPayMethod = mvEnv.GetConfig("pm_voucher") Then
        pNewPaymentMethod = "VOUC"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_caf_card") Then
        pNewPaymentMethod = "CAFC"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_gift_in_kind") Then
        pNewPaymentMethod = "GFIK"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_sp") Then
        pNewPaymentMethod = "CASH"
      ElseIf pPayMethod = mvEnv.GetConfig("pm_sr") Then
        pNewPaymentMethod = "SAOR"
      ElseIf pFinancialAdjustment = Batch.AdjustmentTypes.atMove And mvBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment) Then
        pOrigPaymentMethod = pPayMethod
        pNewPaymentMethod = "CASH"
      ElseIf pFinancialAdjustment = Batch.AdjustmentTypes.atMove And pPayMethod = mvEnv.GetConfig("pm_gaye") And (mvBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.GiveAsYouEarn) Or mvBatchType = Access.Batch.GetBatchTypeCode(Batch.BatchTypes.PostTaxPayrollGiving)) Then
        pOrigPaymentMethod = pPayMethod
        pNewPaymentMethod = "GAYE"
      ElseIf Not pFinancialAdjustment = Batch.AdjustmentTypes.atNone Then
        pOrigPaymentMethod = pPayMethod
        pNewPaymentMethod = "CASH"
      End If
    End Sub


    Private Sub ProcessIncentives(ByVal pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pPaymentPlan As PaymentPlan, ByVal pContactNo As Integer, Optional ByRef pSource As String = "", Optional ByVal pAdditionalParams As CDBParameters = Nothing)

      Dim vTransactionType As String
      Dim vJoined As String = ""
      Dim vContactNumber As Integer
      Dim vPayerContact As New Contact(mvEnv)
      Dim vMemberContact As New Contact(mvEnv)
      Dim vSourceCode As String
      Dim vSource As New Source
      Dim vPayMethodReason As String = ""
      Dim vParams As CDBParameters
      Dim vBalance As Double
      Dim vRenewalAmount As Double
      Dim vReasonForDespatch As String
      Dim vIsMembershipPP As Boolean

      vParams = pTransaction.PaymentPlan.CheckPayPlanParameterList(pParams)

      vTransactionType = pParams("TransactionType").Value

      If pPaymentPlan.Existing = False And pParams.Exists("PaymentPlanNumber") Then pPaymentPlan.Init(mvEnv, (pParams("PaymentPlanNumber").IntegerValue))
      If pPaymentPlan.Existing = True AndAlso pTransaction.TraderPPDLines.Count = 0 Then
        'Add detail lines to the collection
        pTransaction.TraderPPDLines.AddDetailLinesFromPaymentPlan(pPaymentPlan)
      End If

      vIsMembershipPP = (vTransactionType = "MEMB" Or vTransactionType = "CMEM" Or vTransactionType = "MEMC") And ConversionShowPPD = False
      If vIsMembershipPP = True AndAlso AppType = ApplicationType.atConversion AndAlso pPaymentPlan.PlanType = CDBEnvironment.ppType.pptMember Then
        'Conversion and already a member so musy have added an APM
        vIsMembershipPP = False
      End If
      If vIsMembershipPP Then
        If vTransactionType = "MEMC" Then 'TODO
          vJoined = pParams.ParameterExists("CMT_Joined").Value
          If Len(vJoined) = 0 Then
            Exit Sub
          Else
            vContactNumber = pParams("CMT_GiverContactNumber").IntegerValue
            vSourceCode = pParams("CMT_Source").Value
          End If
        Else
          vJoined = pParams.ParameterExists("MEM_Joined").Value
          If Len(vJoined) = 0 Then
            'BR13388: Incentives are not required
            Exit Sub
          Else
            vContactNumber = pParams("MEM_ContactNumber").IntegerValue
            vSourceCode = pParams("MEM_Source").Value
          End If
        End If
      Else
        vSourceCode = pParams.ParameterExists("PPD_Source").Value
        If Len(vSourceCode) = 0 Then vSourceCode = pParams.ParameterExists("Source").Value
        If Len(vSourceCode) = 0 Then vSourceCode = pParams.ParameterExists("TransactionSource").Value
        If AppType = ApplicationType.atConversion Then
          If pPaymentPlan.PlanType = CDBEnvironment.ppType.pptMember Then
            pPaymentPlan.LoadMembers()
            vContactNumber = pTransaction.PaymentPlan.Member.ContactNumber
          Else
            vContactNumber = CType(pPaymentPlan.Details.Item(1), PaymentPlanDetail).ContactNumber
          End If
        End If
        'As it's not a membership, set vJoined to todays date so prices calculated correctly
        vJoined = TodaysDate()
      End If

      If Len(pSource) > 0 Then vSourceCode = pSource
      If vParams.Exists("Source") = False Then vParams.Add("Source", CDBField.FieldTypes.cftCharacter, vSourceCode)

      vReasonForDespatch = pPaymentPlan.IncentiveReason(vTransactionType, AppType, pParams)

      If vReasonForDespatch <> CCReason And vReasonForDespatch <> DDReason And vReasonForDespatch <> SOReason Then
        'M or O-type order incentives, use specified payment method to determine if other
        'reason_for_despatch needs to be used in SQL
        Select Case pParams.ParameterExists("PPPaymentType").Value
          Case "DIRD"
            vPayMethodReason = DDReason
          Case "CCCA"
            vPayMethodReason = CCReason
          Case "STDO"
            vPayMethodReason = SOReason
        End Select
      End If

      vPayerContact.Init(pContactNo)
      vMemberContact.Init(vContactNumber)
      vSource = New Source
      vSource.Init(mvEnv, vSourceCode)

      vBalance = pPaymentPlan.Balance
      vRenewalAmount = pPaymentPlan.RenewalAmount

      pPaymentPlan.WSProcessIncentives(mvEnv, vSource, pTransaction, vPayerContact, vMemberContact, vJoined, vReasonForDespatch, vIsMembershipPP, vPayMethodReason, vBalance, vRenewalAmount, False, vParams, BooleanValue(pParams.OptionalValue("AddIncentiveToPPD", "Y")), True, pAdditionalParams)

      If vTransactionType = "MEMC" Then
        If pParams.Exists("Balance") = False Then pParams.Add("Balance", CDBField.FieldTypes.cftNumeric, vBalance.ToString)
        If pParams.Exists("NewRenewalAmount") = False Then pParams.Add("NewRenewalAmount", CDBField.FieldTypes.cftNumeric, vRenewalAmount.ToString)
      ElseIf AppType = ApplicationType.atConversion Then
        If pParams.Exists("PPBalance") = False Then pParams.Add("PPBalance", CDBField.FieldTypes.cftNumeric)
        pParams("PPBalance").Value = vBalance.ToString()
      End If
    End Sub

    Private Sub SaveAutoPaymentMethodChanges(ByVal pPageType As TraderPage.TraderPageType, ByVal pParams As CDBParameters, ByVal pTransaction As TraderTransaction)
      Dim vNumber As Integer
      Dim vBKDNumber As Integer
      Dim vClaimDay As String
      Dim vAutoStart As String
      Dim vChanged As Boolean
      Dim vTDRLine As TraderAnalysisLine
      Dim vType As TraderAnalysisLine.TraderAnalysisLineTypes
      Dim vLineCount As Integer
      Dim vNewBank As Boolean
      Dim vCreateAccount As String
      Dim vTDRBankDetails As New TraderBankDetails
      Dim vExistingTrans As Boolean
      Dim vFinancialAdjustment As Batch.AdjustmentTypes
      Dim vBatchNumber As Integer
      Dim vTransactionNumber As Integer

      If SupportsNonFinancialBatch Then InitNonFinancialTransaction()
      pTransaction.PaymentPlan.Init(mvEnv, (pParams.ParameterExists("PaymentPlanNumber").IntegerValue))

      If BatchNumber > 0 Then
        vBatchNumber = BatchNumber
        vTransactionNumber = TransactionNumber
      Else
        vBatchNumber = NonFinancialBatchNumber
        vTransactionNumber = NonFinancialTransactionNumber
      End If

      With pTransaction
        If Not .PaymentPlan.Existing Then RaiseError(DataAccessErrors.daePaymentPlanNotFound)
        'Before we update the CCCA/DD/SO, check to see if the StartDate has been changed (for CCCA/DD only)
        vAutoStart = pParams.ParameterExists("StartDate").Value
        vCreateAccount = pParams.ParameterExists("CreateAccount").Value
        vNewBank = pParams.ParameterExists("NewBank").Bool

        Select Case pPageType
          Case TraderPage.TraderPageType.tpCreditCardAuthority
            If Not .PaymentPlan.CreditCardAuthority.Existing Then RaiseError(DataAccessErrors.daeRecordDoesNotExists, "Credit Card Authority")
            vNumber = .PaymentPlan.CreditCardAuthority.CreditCardAuthorityNumber
          Case TraderPage.TraderPageType.tpDirectDebit
            If Not .PaymentPlan.DirectDebit.Existing Then RaiseError(DataAccessErrors.daeRecordDoesNotExists, "Direct Debit")
            vBKDNumber = .PaymentPlan.DirectDebit.BankDetailsNumber
            vNumber = .PaymentPlan.DirectDebit.DirectDebitNumber
          Case TraderPage.TraderPageType.tpStandingOrder
            If Not .PaymentPlan.StandingOrder.Existing Then RaiseError(DataAccessErrors.daeRecordDoesNotExists, "Standing Order")
            vBKDNumber = .PaymentPlan.StandingOrder.BankDetailsNumber
            vNumber = .PaymentPlan.StandingOrder.StandingOrderNumber
        End Select

        If vAutoStart.Length > 0 Then
          If pPageType = TraderPage.TraderPageType.tpDirectDebit Then
            If vAutoStart <> .PaymentPlan.DirectDebit.StartDate Then vChanged = True
          ElseIf pPageType = TraderPage.TraderPageType.tpCreditCardAuthority Then
            If vAutoStart <> .PaymentPlan.CreditCardAuthority.StartDate Then vChanged = True
          End If
        End If

        If pPageType = TraderPage.TraderPageType.tpStandingOrder Or pPageType = TraderPage.TraderPageType.tpDirectDebit Then
          If pParams.ParameterExists("BankDetailsNumber").Value.Length > 0 AndAlso pParams.ParameterExists("BankDetailsNumber").IntegerValue < 1 Then
            If pParams.ParameterExists("IbanNumber").Value.Length > 0 Then vCreateAccount = "Y" 'Force to always create a new ContactAccount
            If vCreateAccount.Length = 0 Then
              RaiseError(DataAccessErrors.daeCreateContactAccount)
            ElseIf BooleanValue(vCreateAccount) Then
              vBKDNumber = pParams("BankDetailsNumber").IntegerValue
            End If
          End If
          vExistingTrans = pParams.ParameterExists("ExistingTransaction").Bool
          vFinancialAdjustment = CType(pParams.ParameterExists("FinancialAdjustment").IntegerValue, Access.Batch.AdjustmentTypes)

          If pPageType = TraderPage.TraderPageType.tpDirectDebit Then
            Dim vOldBankDetailsNumber As Integer = .PaymentPlan.DirectDebit.BankDetailsNumber
            Dim vCA As New ContactAccount()
            vCA.Init(mvEnv, vOldBankDetailsNumber)
            pParams.Add("OldIbanNumber", vCA.IbanNumber)
          End If

          vTDRBankDetails.Init(mvEnv, pParams, vExistingTrans, (vFinancialAdjustment <> Access.Batch.AdjustmentTypes.atNone), vNewBank, vBatchNumber, vTransactionNumber)
        End If

        .PaymentPlan.UpdateAutoPaymentMethod(pParams, vTDRBankDetails.BankDetailsNumber)

        Select Case pPageType
          Case TraderPage.TraderPageType.tpCreditCardAuthority
            If .PaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes Then
              With .PaymentPlan.CreditCardAuthority
                If .ContactCreditCard.Existing Then .ContactCreditCard.Save()
                .Save(, False, vBatchNumber, vTransactionNumber)
              End With
            End If
          Case TraderPage.TraderPageType.tpDirectDebit
            If .PaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
              With .PaymentPlan.DirectDebit
                .Save(, False, vBatchNumber, vTransactionNumber)
              End With
            End If
          Case TraderPage.TraderPageType.tpStandingOrder
            If .PaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes Then
              With .PaymentPlan.StandingOrder
                .Save(, False, vBatchNumber, vTransactionNumber)
              End With
            End If
        End Select

        If (pPageType = TraderPage.TraderPageType.tpDirectDebit Or pPageType = TraderPage.TraderPageType.tpCreditCardAuthority) Then
          If .PaymentPlan.AutoPaymentClaimDateMethod = PaymentPlan.AutoPaymentClaimDateMethods.apcdmDays Then
            'Save any change to the ClaimDay
            vClaimDay = CStr(pParams.ParameterExists("ClaimDay").IntegerValue)
            If vClaimDay <> .PaymentPlan.ClaimDay Then
              .PaymentPlan.UpdateAutoPayMethodClaimDay(vClaimDay)
              If vChanged = False And .PaymentPlan.Balance = 0 And (CDate(.PaymentPlan.RenewalDate) > CDate(TodaysDate())) Then
                'Only the ClaimDay has changed; as Balance = 0 and RenewalDate after Today no need to re-create entire schedule
                vChanged = Not (.PaymentPlan.UpdateScheduledPaymentClaimDates) 'If this fails then we will do a complete re-creation of the payment schedule
                If vChanged = False Then .PaymentPlan.SaveChanges()
              Else
                vChanged = True
              End If
            End If
          End If
          If vChanged Then
            'StartDate and/or ClaimDay has changed
            .PaymentPlan.RegenerateScheduledPayments(OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrPaymentPlanMaintenance, TodaysDate)
            .PaymentPlan.SaveChanges()
          End If
        End If

        If SupportsNonFinancialBatch Then mvNonFinancialBatch.UpdateNumberOfTransactions(1)
        If pParams("TransactionLines").IntegerValue > 0 Then
          vLineCount = pParams("TransactionLines").IntegerValue + 1
        Else
          vLineCount = .TraderAnalysisLines.Count + 1
        End If
        vTDRLine = .GetTraderAnalysisLine(vLineCount)
        Select Case pPageType
          Case TraderPage.TraderPageType.tpStandingOrder
            vType = TraderAnalysisLine.TraderAnalysisLineTypes.taltStandingOrderUpdate
          Case TraderPage.TraderPageType.tpDirectDebit
            vType = TraderAnalysisLine.TraderAnalysisLineTypes.taltDirectDebitUpdate
          Case TraderPage.TraderPageType.tpCreditCardAuthority
            vType = TraderAnalysisLine.TraderAnalysisLineTypes.taltCreditCardAuthorityUpdate
        End Select
        vTDRLine.AddAutoPaymentUpdate(vType, .PaymentPlan.PlanNumber, vNumber)
      End With
    End Sub

    ''' <summary>Build Where clause for Invoice printing</summary>
    ''' <param name="pWhereFields">WhereFields collection to be build</param>
    ''' <param name="pAnsiJoins">AnsiJoins to be added</param>
    ''' <param name="pParams">Parameters collection from Trader</param>
    ''' <param name="pPartPaidOnly">Only include part-paid Invoices</param>
    ''' <param name="pBatchOwnership">Only include owned invoices</param>
    ''' <param name="pUseTransDate">Use transaction date etc. in SQL</param>
    Private Sub GetInvoicesWhereClause(ByRef pWhereFields As CDBFields, ByRef pAnsiJoins As AnsiJoins, ByVal pParams As CDBParameters, ByVal pPartPaidOnly As Boolean, ByVal pBatchOwnership As Boolean, ByVal pUseTransDate As Boolean)
      'Build nested SQL
      Dim vNestedSQLStatement As New SQLStatement(mvEnv.Connection, "invoice_number", "invoice_payment_history", New CDBFields, "")
      vNestedSQLStatement.Distinct = True
      Dim vNestedSQL As String = "(" & vNestedSQLStatement.SQL & ")"

      Dim vRunType As String = ""
      If pParams.ParameterExists("RunType").Value.Length > 0 Then vRunType = pParams("RunType").Value

      With pWhereFields
        If pParams.Exists("PrintJobNumber") Then 'Coming from ProcessSelectedInvoices
          .Add("i.print_job_number", pParams("PrintJobNumber").IntegerValue)
        Else
          Dim vFromInvoiceNumber As Integer = pParams.ParameterExists("FromInvoiceNumber").IntegerValue
          If pParams.Exists("FromInvoiceNumber") = False AndAlso pParams.Exists("InvoiceNumber") = True Then vFromInvoiceNumber = pParams("InvoiceNumber").IntegerValue
          Dim vToInvoiceNumber As Integer = pParams.ParameterExists("ToInvoiceNumber").IntegerValue
          If pParams.Exists("ToInvoiceNumber") = False AndAlso pParams.Exists("InvoiceNumber2") = True Then vToInvoiceNumber = pParams("InvoiceNumber2").IntegerValue

          .Add("i.company", pParams("Company").Value)
          If vFromInvoiceNumber > 0 Then
            If vRunType.Length = 0 Then vRunType = "R" 'Default to reprint Invoices
            .Add("i.invoice_number", vFromInvoiceNumber, CType(CDBField.FieldWhereOperators.fwoBetweenFrom + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
            .Add("i.invoice_number#2", vToInvoiceNumber, CType(CDBField.FieldWhereOperators.fwoBetweenTo + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
          Else
            If vRunType.Length = 0 Then vRunType = "N" 'Default to new invoices
          End If
          If pPartPaidOnly Then .Add("i.invoice_number#3", CDBField.FieldTypes.cftInteger, vNestedSQL, CDBField.FieldWhereOperators.fwoIn)
          Select Case vRunType
            Case "N"
              'New Invoices
              .Add("i.invoice_number#4", CDBField.FieldTypes.cftInteger, "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
              .Add("reprint_count", CDBField.FieldTypes.cftInteger, "0", CType(CDBField.FieldWhereOperators.fwoLessThan + CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
            Case "R"
              'Reprint Invoices
              .Add("i.reprint_count", CDBField.FieldTypes.cftInteger, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            Case Else 'A
              'All Invoices so do nothing
          End Select
          If pParams.ParameterExists("BatchNumber").IntegerValue > 0 Then
            'Update just a single Invoice
            .Add("i.batch_number", pParams("BatchNumber").IntegerValue)
            .Add("i.transaction_number", pParams("TransactionNumber").IntegerValue)
          End If
          .Add("i.record_type", "'I','N'", CDBField.FieldWhereOperators.fwoIn)
          If pParams.ParameterExists("FromDate").Value.Length > 0 OrElse pParams.ParameterExists("ToDate").Value.Length > 0 Then
            Dim vUseTransDate As Boolean = (mvEnv.GetConfig("invoice_date_from_event_start").Length > 0)
            Dim vFromDate As String = pParams.ParameterExists("FromDate").Value
            Dim vToDate As String = pParams.ParameterExists("ToDate").Value
            Dim vFromWhereOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoOpenBracket
            Dim vToWhereOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual
            If vUseTransDate Then vFromWhereOperator = CDBField.FieldWhereOperators.fwoOpenBracketTwice
            If vFromDate.Length > 0 AndAlso vToDate.Length > 0 Then
              vFromWhereOperator = CType(vFromWhereOperator + CDBField.FieldWhereOperators.fwoBetweenFrom, CDBField.FieldWhereOperators)
              .Add("invoice_date", CDBField.FieldTypes.cftDate, vFromDate, vFromWhereOperator)
              .Add("invoice_date#2", CDBField.FieldTypes.cftDate, vToDate, CType(CDBField.FieldWhereOperators.fwoBetweenTo + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
            ElseIf vFromDate.Length > 0 Then
              vFromWhereOperator = CType(vFromWhereOperator + CDBField.FieldWhereOperators.fwoGreaterThanEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators)
              .Add("invoice_date", CDBField.FieldTypes.cftDate, vFromDate, vFromWhereOperator)
            Else
              vFromWhereOperator = CType(vFromWhereOperator + CDBField.FieldWhereOperators.fwoLessThanEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators)
              .Add("invoice_date", CDBField.FieldTypes.cftDate, vToDate, vFromWhereOperator)
            End If
            If vUseTransDate Then
              'OR (record_type = 'N' AND invoice_date IS NULL
              .Add("record_type", "N", CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
              .Add("invoice_date#3", CDBField.FieldTypes.cftDate, "")
              vToWhereOperator = CDBField.FieldWhereOperators.fwoCloseBracketTwice
              If vFromDate.Length > 0 AndAlso vToDate.Length > 0 Then
                'AND transaction_date BETWEEN ... AND ...))
                vToWhereOperator = CType(vToWhereOperator + CDBField.FieldWhereOperators.fwoBetweenTo, CDBField.FieldWhereOperators)
                .Add("transaction_date", CDBField.FieldTypes.cftDate, vFromDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
                .Add("transaction_date#2", CDBField.FieldTypes.cftDate, vToDate, vToWhereOperator)
              ElseIf vFromDate.Length > 0 Then
                'AND transaction_date >= ...))
                vToWhereOperator = CType(vToWhereOperator + CDBField.FieldWhereOperators.fwoGreaterThanEqual, CDBField.FieldWhereOperators)
                .Add("transaction_date", CDBField.FieldTypes.cftDate, vFromDate, vToWhereOperator)
              Else
                'AND transaction_date <= ...))
                vToWhereOperator = CType(vToWhereOperator + CDBField.FieldWhereOperators.fwoLessThanEqual, CDBField.FieldWhereOperators)
                .Add("transaction_date", CDBField.FieldTypes.cftDate, vToDate, vToWhereOperator)
              End If
            End If
          End If

          If pParams.ParameterExists("StartBatch").Value.Length > 0 OrElse pParams.ParameterExists("EndBatch").Value.Length > 0 Then
            Dim vStartBatch As String = pParams.ParameterExists("StartBatch").Value
            Dim vEndBatch As String = pParams.ParameterExists("EndBatch").Value
            .Add("i.batch_number", CDBField.FieldTypes.cftInteger, vStartBatch, CDBField.FieldWhereOperators.fwoBetweenFrom)
            .Add("i.batch_number#2", CDBField.FieldTypes.cftInteger, vEndBatch, CDBField.FieldWhereOperators.fwoBetweenTo)
          End If

          If pBatchOwnership Then .Add("d.department", mvEnv.User.Department)
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
          .Add("print_invoice", CDBField.FieldTypes.cftCharacter, "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
          .Add("print_invoice#2", "Y", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        End If
      End With

    End Sub

    ''' <summary>Add Events selection to Invoice printing SQL</summary>
    ''' <param name="pAnsiJoins">AnsiJoins collection to be added to</param>
    ''' <param name="pParams">Parameters collection from Trader</param>
    ''' <param name="pPartPaidOnly">Only include part-paid Invoices</param>
    ''' <param name="pBatchOwnership">Only include owned invoices</param>
    ''' <param name="pUseTransDate">Include transaction date etc. in SQL</param>
    Private Sub GetInvoicesEvents(ByRef pAnsiJoins As AnsiJoins, ByVal pParams As CDBParameters, ByVal pPartPaidOnly As Boolean, ByVal pBatchOwnership As Boolean, ByVal pUseTransDate As Boolean)
      'First nested SQL
      Dim vAttrs1 As String = "det.batch_number, det.transaction_number, eb.event_number, e.event_desc"
      Dim vFrom1 As String = "({0}) det"
      Dim vAnsiJoins1 As New AnsiJoins
      vAnsiJoins1.Add("event_bookings eb", "det.batch_number", "eb.batch_number", "det.transaction_number", "eb.transaction_number")
      vAnsiJoins1.Add("events e", "eb.event_number", "e.event_number")
      Dim vGroupBy1 As String = "det.batch_number, det.transaction_number, eb.event_number, e.event_desc"

      'Second nested SQL
      Dim vAttrs2 As String = "batch_number, transaction_number"
      Dim vFrom2 As String = "({0}) de"
      Dim vGroupBy2 As String = "batch_number, transaction_number HAVING COUNT (*) = 1"

      'Third nested SQL
      Dim vAttrs3 As String = "i.batch_number, i.transaction_number, eb.event_number"
      Dim vFrom3 As String = "invoices i"
      Dim vAnsiJoins3 As New AnsiJoins
      With vAnsiJoins3
        .Add("event_bookings eb", "i.batch_number", "eb.batch_number", "i.transaction_number", "eb.transaction_number")
        If pUseTransDate Then .Add("batch_transactions bt", "eb.batch_number", "bt.batch_number", "eb.transaction_number", "bt.transaction_number")
        If pBatchOwnership Then
          .Add("batches b", "i.batch_number", "b.batch_number")
          .Add("users u", "b.batch_created_by", "u.logname")
          .Add("departments d", "u.department", "d.department")
        End If
      End With
      Dim vWhereFields3 As New CDBFields
      GetInvoicesWhereClause(vWhereFields3, vAnsiJoins3, pParams, pPartPaidOnly, pBatchOwnership, pUseTransDate)

      'Now build the nested SQL statements
      Dim vSQL3 As New SQLStatement(mvEnv.Connection, vAttrs3, vFrom3, vWhereFields3, "", vAnsiJoins3)
      vSQL3.Distinct = True

      vFrom2 = String.Format(vFrom2, vSQL3.SQL)
      Dim vSQL2 As New SQLStatement(mvEnv.Connection, vAttrs2, vFrom2, New CDBFields, "")
      vSQL2.GroupBy = vGroupBy2

      vFrom1 = String.Format(vFrom1, vSQL2.SQL)
      Dim vSQL1 As New SQLStatement(mvEnv.Connection, vAttrs1, vFrom1, New CDBFields, "", vAnsiJoins1)
      vSQL1.GroupBy = vGroupBy1

      'Add the above SQL as a Left Outer Join
      pAnsiJoins.AddLeftOuterJoin("(" & vSQL1.SQL & ") evb", "i.batch_number", "evb.batch_number", "i.transaction_number", "evb.transaction_number")

    End Sub

    Private Sub EditEventBooking(ByVal pParams As CDBParameters, ByVal pTransaction As TraderTransaction)
      'Trader Application allows Event Booking to be amended and a new transaction created but original Bookinhs has not been amended
      'so need to add a new Booking anyway for the new transaction
      Dim vContact As Contact
      Dim vDelegate As EventDelegate
      Dim vEvent As CDBEvent
      Dim vOriginalEB As New EventBooking
      Dim vNewEB As EventBooking
      Dim vRS As CDBRecordSet
      Dim vSession As EventSession
      Dim vSessionList As String = ""
      Dim vSQL As String
      Dim vTDRLine As TraderAnalysisLine
      Dim vTrans As Boolean
      Dim vEventDelegate As EventDelegate

      If pParams.Exists("EventNumber") = False Then
        'Need to set EventNumber to current Event Number
        'And TrdLine Bookingnumber to new Event Booking Number
        'And TrdLine AmendedBookingNumber to orignal Event Booking Number
        vOriginalEB.Init(mvEnv)
        vSQL = "SELECT " & vOriginalEB.GetRecordSetFields(EventBooking.EventBookingRecordSetTypes.ebrtAll) & " FROM event_bookings eb WHERE batch_number = " & pParams("BatchNumber").IntegerValue
        vSQL = vSQL & " AND transaction_number = " & pParams("TransactionNumber").IntegerValue
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch() = True
          vOriginalEB = New EventBooking
          vEvent = New CDBEvent(mvEnv)
          vOriginalEB.InitFromRecordSet(mvEnv, vRS, EventBooking.EventBookingRecordSetTypes.ebrtAll)
          vEvent.Init((vOriginalEB.EventNumber))
          If vOriginalEB.Existing = True And vEvent.Existing = True Then
            With vOriginalEB
              If .BookingStatus = EventBooking.EventBookingStatuses.ebsCancelled Then
                RaiseError(DataAccessErrors.daeAdjustmentError, "The Event Booking has been cancelled and cannot be amended")
              ElseIf .ValidStatusChange(EventBooking.EventBookingStatuses.ebsAmended) = False Then
                RaiseError(DataAccessErrors.daeAdjustmentError, "The Event booking cannot be amended")
              End If
              If .Sessions.Count() = 1 Then
                vSession = CType(.Sessions.Item(1), EventSession)
                If vSession.SessionType <> vSession.BaseSessionType Then vSessionList = CStr(vSession.SessionNumber)
              ElseIf .Sessions.Count() > 1 Then
                For Each vSession In .Sessions
                  If Len(vSessionList) > 0 Then vSessionList = vSessionList & ","
                  vSessionList = vSessionList & vSession.SessionNumber
                Next vSession
              End If
              vContact = New Contact(mvEnv)
              vContact.Init(.ContactNumber)
              vDelegate = New EventDelegate
              vDelegate.Init(mvEnv)
              If .Delegates.Count() > 0 Then
                For Each vEventDelegate In .Delegates
                  If vEventDelegate.ContactNumber = .ContactNumber Then
                    vDelegate = CType(.Delegates.Item(CStr(.ContactNumber)), EventDelegate)
                  End If
                Next
              End If

              If mvEnv.Connection.InTransaction = False Then
                mvEnv.Connection.StartTransaction()
                vTrans = True
              End If

              If vDelegate.Existing Then .RemoveDelegate(vDelegate, mvEnv.User.Logname)
              .SetBookingAmended(vEvent, True) 'Do not Save
              vNewEB = vEvent.AddEventBooking(vContact, .AddressNumber, .Quantity, .OptionNumber, .BookingStatus, .RateCode, vSessionList, .Notes, 0, 0, 0, Nothing, 0, False, CStr(.AdultQuantity), CStr(.ChildQuantity), .BookingDate, .StartTime, .EndTime)
            End With

            If vNewEB Is Nothing Then
              RaiseError(DataAccessErrors.daeAdjustmentError, (vEvent.LastBookingMessage))
            Else
              If pParams.Exists("EventNumber") = False Then pParams.Add("EventNumber", vEvent.EventNumber)
              For Each vTDRLine In pTransaction.TraderAnalysisLines
                vTDRLine.UpdateEventBookingDetails(vOriginalEB.BookingNumber, vNewEB.BookingNumber, vNewEB.EventNumber)
              Next vTDRLine
            End If
          End If
        End While
        vRS.CloseRecordSet()
        If vTrans Then mvEnv.Connection.CommitTransaction()
      End If

    End Sub

    Private Function SumLineTypes(ByVal pLineType As String, ByVal pTransaction As TraderTransaction) As Double
      Dim vAmount As Double
      Dim vTDRLine As TraderAnalysisLine

      For Each vTDRLine In pTransaction.TraderAnalysisLines
        If vTDRLine.TraderTransactionTypeCode = pLineType Then
          Select Case pLineType
            Case "SALE"
              Dim vSQLStatement As SQLStatement
              Dim vWhereField As New CDBFields
              vWhereField.Add("product", vTDRLine.ProductCode)
              vSQLStatement = New SQLStatement(mvEnv.Connection, "stock_item", "products", vWhereField)
              If vSQLStatement.GetValue() = "Y" Then
                vAmount = vAmount + vTDRLine.Amount
              End If
            Case Else
              vAmount = vAmount + vTDRLine.Amount
          End Select
        End If
      Next
      Return vAmount
    End Function

    Private Function ExistingAnalysisLine(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer) As Boolean
      Dim vWhereFields As New CDBFields
      With vWhereFields
        .Add("batch_number", pBatchNumber)
        .Add("transaction_number", TransactionNumber)
        .Add("line_number", pLineNumber)
      End With
      Return (mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) > 0)
    End Function

    Public Sub SetNoCovenants()
      mvCovDonationsRegular = False
      mvCovMemberships = False
      mvCovSubscriptions = False
    End Sub

    Public Sub AdjustNumberOfTransactions(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vNextTransNumber As Integer       'BR13954 - To adjust number of transactions and next transaction number fields
      Dim vNumberOfTrans As Integer
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("select number_of_transactions, next_transaction_number from batches where batch_number = " & pBatchNumber)
      If vRecordSet.Fetch() = True Then
        With vRecordSet
          vNextTransNumber = .Fields("next_transaction_number").IntegerValue
          vNumberOfTrans = .Fields("number_of_transactions").IntegerValue
        End With

        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
        vUpdateFields.Add("number_of_transactions", CDBField.FieldTypes.cftLong, vNumberOfTrans - 1)
        If vNextTransNumber = (pTransactionNumber + 1) Then vUpdateFields.Add("next_transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
        vUpdateFields.Add("contents_amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
        vUpdateFields.Add("contents_amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
        mvEnv.Connection.UpdateRecords("batches", vUpdateFields, vWhereFields)
      End If
    End Sub

    ''' <summary>Select Invoices for printing and further processing.</summary>
    ''' <param name="pParams">Parameter collection from Trader</param>
    ''' <param name="pUseTransDate">Date range has been entered and configuration option invoice_date_from_event_start is set</param>
    ''' <param name="pBatchOwnership">Configuration option opt_batch_ownershipis set and configuration option opt_batch_per_user = "DEPARTMENT"</param>
    ''' <param name="pPartPaidOnly">Only include Invoices that have received a payment</param>
    ''' <returns>SQLStatement</returns>
    Private Function SelectInvoicesForPrinting(ByVal pParams As CDBParameters, ByVal pUseTransDate As Boolean, ByVal pBatchOwnership As Boolean, ByVal pPartPaidOnly As Boolean) As SQLStatement
      Return SelectInvoicesForPrinting(pParams, pUseTransDate, pBatchOwnership, pPartPaidOnly, False)
    End Function
    ''' <summary>Select Invoices for printing and further processing.</summary>
    ''' <param name="pParams">Parameter collection from Trader</param>
    ''' <param name="pUseTransDate">Date range has been entered and configuration option invoice_date_from_event_start is set</param>
    ''' <param name="pBatchOwnership">Configuration option opt_batch_ownershipis set and configuration option opt_batch_per_user = "DEPARTMENT"</param>
    ''' <param name="pPartPaidOnly">Only include Invoices that have received a payment</param>
    ''' <param name="pCountOnly">Build SQL for a Count</param>
    ''' <returns>SQLStatement</returns>
    Private Function SelectInvoicesForPrinting(ByVal pParams As CDBParameters, ByVal pUseTransDate As Boolean, ByVal pBatchOwnership As Boolean, ByVal pPartPaidOnly As Boolean, ByVal pCountOnly As Boolean) As SQLStatement
      Dim vCheckEvents As Boolean = Not pCountOnly
      Dim vDisplayInvoicesOnly As Boolean = pParams.ParameterExists("DisplayInvoices").Bool

      Dim vInvoice As New Invoice()
      vInvoice.Init(mvEnv)
      Dim vAttrs As String
      If vDisplayInvoicesOnly Then
        vAttrs = "i.record_type, i.invoice_date, i.invoice_number, i.company, i.batch_number, i.transaction_number, i.contact_number, ct.label_name, i.sales_ledger_account,"
        vAttrs &= " SUM(bta.amount) AS amount, SUM(bta.gross_amount) AS gross_amount, SUM(bta.vat_amount) AS vat_amount, evb.event_desc, evb.event_number"
      Else
        vAttrs = vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", ct.label_name"  '"i.record_type, i.invoice_date, i.invoice_number,"
        If vCheckEvents Then vAttrs &= ", evb.event_desc, evb.event_number"
        vAttrs &= ", cc.terms_from, cc.terms_period, cc.terms_number,"
        vAttrs &= " c.terms_from AS company_terms_from, c.terms_period AS company_terms_period, c.terms_number AS company_terms_number"
      End If

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        If pParams.ParameterExists("InstantPrint").Bool = True OrElse PrintInvoiceUnpostedBatches = True Then
          'Come from Trader transaction entry so batch is not posted, OR Trader BatchInvoiceProduction that has been set to print invoices in unposted batches
          .Add("batch_transactions bt", "i.batch_number", "bt.batch_number", "i.transaction_number", "bt.transaction_number")
        Else
          'Should be Trader BatchInvoiceProduction so batch has to be posted
          .Add("financial_history fh", "i.batch_number", "fh.batch_number", "i.transaction_number", "fh.transaction_number")
        End If
        .Add("credit_customers cc", "i.company", "cc.company", "i.sales_ledger_account", "cc.sales_ledger_account")
        .Add("company_credit_controls c", "cc.company", "c.company")
        If vDisplayInvoicesOnly Then .Add("batch_transaction_analysis bta", "i.batch_number", "bta.batch_number", "i.transaction_number", "bta.transaction_number")
        If pBatchOwnership Then
          .Add("batches b", "i.batch_number", "b.batch_number")
          .Add("users u", "b.batch_created_by", "u.logname")
          .Add("departments d", "u.department", "d.department")
        End If
        .Add("contacts ct", "i.contact_number", "ct.contact_number")
      End With

      Dim vWhereFields As New CDBFields()

      If vCheckEvents Then GetInvoicesEvents(vAnsiJoins, pParams, pPartPaidOnly, pBatchOwnership, pUseTransDate)
      GetInvoicesWhereClause(vWhereFields, vAnsiJoins, pParams, pPartPaidOnly, pBatchOwnership, pUseTransDate)

      Dim vGroupBy As String = ""
      If vDisplayInvoicesOnly Then
        vGroupBy = "i.record_type, i.invoice_date, i.invoice_number, i.company, i.batch_number, i.transaction_number, i.contact_number, ct.label_name, i.sales_ledger_account, evb.event_desc, evb.event_number "
      ElseIf pCountOnly Then
        vGroupBy = vInvoice.GetRecordSetFields(Invoice.InvoiceRecordSetTypes.irtAll) & ", ct.label_name"
        If vCheckEvents Then vGroupBy &= ", evb.event_desc, evb.event_number"
        vGroupBy &= ", cc.terms_from, cc.terms_period, cc.terms_number, c.terms_from, c.terms_period, c.terms_number, payment_due"
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "invoices i", vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = vGroupBy

      Return vSQLStatement

    End Function

    ''' <summary>After print-preview of Invoices remove any Invoice numbers that got added</summary>
    ''' <param name="pPrintJobNumber">Print Job Number of the Invoice run</param>
    ''' <param name="pInvoiceNumbers">A Comma-separated list of invoice numbers to be removed</param>
    Public Sub ClearSelectedInvoiceNumbers(ByVal pPrintJobNumber As Integer, ByVal pInvoiceNumbers As String)
      Dim vTrans As Boolean
      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      'Update Invoices
      Dim vWhereFields As New CDBFields(New CDBField("print_job_number", pPrintJobNumber))
      With vWhereFields
        .Add("invoice_number", CDBField.FieldTypes.cftInteger, pInvoiceNumbers, CDBField.FieldWhereOperators.fwoIn)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber) Then .Add("provisional_invoice_number", CDBField.FieldTypes.cftInteger, "invoice_number")
        .Add("amount_paid", CDBField.FieldTypes.cftNumeric, "0")
        .Add("record_type", "I")
      End With
      Dim vUpdateFields As New CDBFields(New CDBField("invoice_number", CDBField.FieldTypes.cftInteger, ""))
      vUpdateFields.Add("provisional_invoice_number", CDBField.FieldTypes.cftInteger, "")
      mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields, False)
      'Update Credit Notes
      vWhereFields.Remove("record_type")
      vWhereFields.Add("record_type", "N")
      vWhereFields.Remove("amount_paid")      'As a CreditNote it may have been allocated without a number so AmountPaid will be > 0
      vUpdateFields.Add("invoice_date", CDBField.FieldTypes.cftDate, "")
      vUpdateFields.Add("payment_due", CDBField.FieldTypes.cftDate, "")
      mvEnv.Connection.UpdateRecords("invoices", vUpdateFields, vWhereFields, False)

      'Update InvoiceDetails
      vWhereFields.Clear()
      vWhereFields.Add("invoice_number", CDBField.FieldTypes.cftInteger, pInvoiceNumbers, CDBField.FieldWhereOperators.fwoIn)
      vUpdateFields.Clear()
      vUpdateFields.Add("invoice_number", CDBField.FieldTypes.cftInteger, "")
      mvEnv.Connection.UpdateRecords("invoice_details", vUpdateFields, vWhereFields)

      If vTrans Then mvEnv.Connection.CommitTransaction()

    End Sub

    ''' <summary>Used when cancelling out of a financial adjustment Trader to work out whether the Salesledger UnallocatedCash line is the original BTA line.</summary>
    Private Function IsOriginalAnalsysisLine(ByVal pOriginalBT As BatchTransaction, ByVal pTDRLine As TraderAnalysisLine, ByVal pParams As CDBParameters) As Boolean
      'To handle other lines / info this will need to be extended to match what FA does
      If pOriginalBT Is Nothing _
      OrElse (pOriginalBT.Existing = True AndAlso (pOriginalBT.BatchNumber <> pParams.ParameterExists("BatchNumber").IntegerValue OrElse pOriginalBT.TransactionNumber <> pParams.ParameterExists("TransactionNumber").IntegerValue)) Then
        pOriginalBT = New BatchTransaction(mvEnv)
        pOriginalBT.Init(pParams.ParameterExists("BatchNumber").IntegerValue, pParams.ParameterExists("TransactionNumber").IntegerValue, True)
        pOriginalBT.InitBatchTransactionAnalysis(pOriginalBT.BatchNumber, pOriginalBT.TransactionNumber)
      End If
      If pOriginalBT.Analysis.Count = 0 Then
        Return True   'Failed to find data so assume it is original line
      Else
        For Each vBTA As BatchTransactionAnalysis In pOriginalBT.Analysis
          If pOriginalBT.TransactionSign = "D" Then vBTA.ChangeSign() 'For FA's, BTA has positive figures but Trader has negative figures (as FHD is negative) so make BTA negative as well
          If vBTA.LineType = pTDRLine.TraderLineTypeCode AndAlso vBTA.Amount = pTDRLine.Amount AndAlso vBTA.LineNumber <= pTDRLine.LineNumber Then
            Return True
          End If
        Next
      End If
    End Function

  End Class
End Namespace
