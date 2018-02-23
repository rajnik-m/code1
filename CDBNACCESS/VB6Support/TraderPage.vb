Namespace Access

  Partial Public Class TraderPage

    Public Enum TraderPageType

      tpNone = 0
      tpPaymentMethod1 'PG_PM1
      tpCreditCustomer 'PG_CCU
      tpTransactionDetails 'PG_TRD
      tpComments 'PG_COM
      tpBankDetails 'PG_BKD
      tpCardDetails 'PG_CDC
      tpTransactionAnalysis 'PG_TRA
      tpPaymentMethod2 'PG_PM2
      tpProductDetails 'PG_PRD
      tpPayments 'PG_PAY
      tpPaymentPlanDetails 'PG_PPD
      tpPaymentPlanProducts 'PG_PPP
      tpStandingOrder 'PG_STO
      tpDirectDebit 'PG_DDR
      tpCreditCardAuthority 'PG_CCA
      tpChangeMembershipType 'PG_CMT
      tpMembership 'PG_MEM
      tpAmendMembership 'PG_AMD
      tpMembershipPayer 'PG_MSP
      tpCovenant 'PG_COV
      tpContactSelection 'PG_CSE
      tpTransactionAnalysisSummary 'PG_TAS     Summary only No Controls
      tpPaymentPlanSummary 'PG_PPS     Summary only No Controls
      tpEventBooking 'PG_EVE
      tpExamBooking 'PG_EXA
      tpAccommodationBooking 'PG_ACO
      tpPostageAndPacking 'PG_PAP
      tpServiceBooking 'PG_SVC
      tpPaymentPlanFromUnbalanceTransaction 'PG_TPP
      tpPaymentMethod3 'PG_PM3
      tpSetStatus 'PG_STA
      tpCancelPaymentPlan 'PG_CPP
      tpLegacyBequestReceipt 'PG_LBR
      tpActivityEntry 'PG_ACT
      tpGiftAidDeclaration 'PG_GAD
      tpGoneAway 'PG_GAW
      tpAddressMaintenance 'PG_ADM
      tpSuppressionEntry 'PG_SUP
      tpCancelGiftAidDeclaration 'PG_CGA
      tpScheduledPayments 'PG_SCP
      tpOutstandingScheduledPayments 'PG_OSP
      tpConfirmProvisionalTransactions 'PG_CPT
      tpGiveAsYouEarnEntry 'PG_GYE
      tpCollectionPayments 'PG_COL
      tpPaymentPlanMaintenance 'PG_PPM
      tpPaymentPlanDetailsMaintenance 'PG_PPN
      tpLoans   'PG_LON
      tpAdvancedCMT     'PG_MTC
      tpTokenSelection  'PG_TKN 
      'Add any new pages for transactions above here
      'The following pages have no controls and are summary pages only
      tpStatementList 'PG_STL   Summary only No Controls
      tpInvoicePayments 'PG_INV   Summary only No Controls
      tpMembershipMembersSummary 'PG_MMS   Summary only No Controls
      'The following pages are only loaded for specific application types (non-transactional)
      tpPurchaseInvoiceDetails 'PG_PID   PINVE Application Type
      tpPurchaseInvoiceProducts 'PG_PIP   PINVE
      tpPurchaseInvoiceSummary 'PG_PIS   PINVE Summary only No Controls
      tpPurchaseOrderDetails 'PG_POD   PORDE Application Type
      tpPurchaseOrderProducts 'PG_POP   PORDE
      tpPurchaseOrderPayments 'PG_PPA   PORDE
      tpPurchaseOrderSummary 'PG_POS   PORDE Summary only No Controls
      tpPurchaseOrderCancellation 'PG_POC   PORDC Application Type
      tpChequeNumberAllocation 'PG_CNA   CHQNA Application Type
      tpChequeReconciliation 'PG_CRE   CHQRE Application Type
      tpCreditStatementGeneration 'PG_CSG   CSTAT Application Type
      tpBatchInvoiceProduction 'PG_ING   BINVG Application Type
      tpGiveAsYouEarn 'PG_GYP   GAYEP Application Type  (Pre Tax Payroll Giving)
      tpPostTaxPGPayment 'PG_PGP   POTPG Application Type  (Post Tax Payroll Giving)
      tpBatchInvoiceSummary 'PG_INS   BINVG Application Type
      tpAmendEventBooking 'PG_AEV   TRANS Application Type (Smart Client only)
      tpDummyPage 'BR9067 Dummy page to force everything to be cleared at end of transaction
      'Do NOT just add items to the end - See comments above
      tpMaximumPageTypePlusOne
    End Enum

  End Class

End Namespace
