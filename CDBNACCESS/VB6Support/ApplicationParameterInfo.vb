Public Class ApplicationParameterInfo

  Public Shared Function GetPageCodeFromFPT(ByRef pType As CDBEnvironment.FunctionParameterTypes) As String
    Select Case pType
      Case CDBEnvironment.FunctionParameterTypes.fptActionActivationDate
        Return "PAAE"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdateStandingOrder
        Return "PSOU"
      Case CDBEnvironment.FunctionParameterTypes.fptPayingInSlipNumber
        Return "PPSN"
      Case CDBEnvironment.FunctionParameterTypes.fptCLRStatementDate
        Return "PCLR"
      Case CDBEnvironment.FunctionParameterTypes.fptInvoicePaymentDue
        Return "PIPD"
      Case CDBEnvironment.FunctionParameterTypes.fptReportNumber
        Return "PDRP"
      Case CDBEnvironment.FunctionParameterTypes.fptChangePayer
        Return "PCHP"
      Case CDBEnvironment.FunctionParameterTypes.fptEventProgrammeReport
        Return "PEPR"
      Case CDBEnvironment.FunctionParameterTypes.fptEventPersonnelReport
        Return "PEER"
      Case CDBEnvironment.FunctionParameterTypes.fptEventAttendeeReport
        Return "PEAR"
      Case CDBEnvironment.FunctionParameterTypes.fptFinancialAdjustment
        Return "PFAD"
      Case CDBEnvironment.FunctionParameterTypes.fptCloseOrganisationSite
        Return "PCOS"
      Case CDBEnvironment.FunctionParameterTypes.fptMoveBranch
        Return "PMBR"
      Case CDBEnvironment.FunctionParameterTypes.fptMoveRegion
        Return "PMGR"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdateContact
        Return "PCUP"
      Case CDBEnvironment.FunctionParameterTypes.fptAppealBudgetPeriod
        Return "ABPD"
      Case CDBEnvironment.FunctionParameterTypes.fptOwnershipGroup
        Return "PSOG"
      Case CDBEnvironment.FunctionParameterTypes.fptConfirmProvisionalTransaction
        Return "PCPT"
      Case CDBEnvironment.FunctionParameterTypes.fptCreateGiftAidDeclaration
        Return "PGAD"
      Case CDBEnvironment.FunctionParameterTypes.fptAuthorisePOPayment
        Return "PAPO"
      Case CDBEnvironment.FunctionParameterTypes.fptAdvanceRenewalDate
        Return "PARD"
      Case CDBEnvironment.FunctionParameterTypes.fptFAPartRefund
        Return "PFPR"
      Case CDBEnvironment.FunctionParameterTypes.fptAttachmentList
        Return "PATT"
      Case CDBEnvironment.FunctionParameterTypes.fptCancReason
        Return "PACR"
      Case CDBEnvironment.FunctionParameterTypes.fptGetMailingCode
        Return "PAMC"
      Case CDBEnvironment.FunctionParameterTypes.fptPayPlanMissedPayments
        Return "PPMP"
      Case CDBEnvironment.FunctionParameterTypes.fptFAReverseRefundOptions
        Return "PFAR"
      Case CDBEnvironment.FunctionParameterTypes.fptSOReconciliationReport
        Return "PSOR"
      Case CDBEnvironment.FunctionParameterTypes.fptReportDataSelection
        Return "PRDS"
      Case CDBEnvironment.FunctionParameterTypes.fptEventCancellationFAType
        Return "PECR"
      Case CDBEnvironment.FunctionParameterTypes.fptChangeSubscriptionCommunication
        Return "PCSC"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdatePaymentPlanDetailSource
        Return "PUDS"
      Case CDBEnvironment.FunctionParameterTypes.fptCMTPriceChange
        Return "PCMP"
      Case CDBEnvironment.FunctionParameterTypes.fptCMTEntitlementPriceChange
        Return "PCEN"
      Case CDBEnvironment.FunctionParameterTypes.fptImportTraderApp
        Return "ITAD"
      Case CDBEnvironment.FunctionParameterTypes.fptScheduleTask
        Return "STJP"
      Case CDBEnvironment.FunctionParameterTypes.fptLMAddressUsage
        Return "LMAU"
      Case CDBEnvironment.FunctionParameterTypes.fptAddCollectionBoxes
        Return "PACB"
      Case CDBEnvironment.FunctionParameterTypes.fptCancellationReasonAndSource
        Return "PARS"
      Case CDBEnvironment.FunctionParameterTypes.fptAllocatePISToEvent
        Return "APIS"
      Case CDBEnvironment.FunctionParameterTypes.fptAllocatePISToDelegates
        Return "APSD"
      Case CDBEnvironment.FunctionParameterTypes.fptDuplicateEvent
        Return "EVDP"
      Case CDBEnvironment.FunctionParameterTypes.fptMembershipReinstatement
        Return "PMRN"
      Case CDBEnvironment.FunctionParameterTypes.fptCancellationReasonSourceAndDate
        Return "PRSD"
      Case CDBEnvironment.FunctionParameterTypes.fptEditAppointment
        Return "PAPP"
      Case CDBEnvironment.FunctionParameterTypes.fptRemoveFutureMembershipType
        Return "PRFM"
      Case CDBEnvironment.FunctionParameterTypes.fptLeavePosition
        Return "PLVP"
      Case CDBEnvironment.FunctionParameterTypes.fptMovePosition
        Return "PMVP"
      Case CDBEnvironment.FunctionParameterTypes.fptCancelPaymentPlan
        Return "PPAC"
      Case CDBEnvironment.FunctionParameterTypes.fptPISPrinting
        Return "PPIS"
      Case CDBEnvironment.FunctionParameterTypes.fptScheduledJobDetails
        Return "PEJD"
      Case CDBEnvironment.FunctionParameterTypes.fptNewMailingCode
        Return "PNMC"
      Case CDBEnvironment.FunctionParameterTypes.fptReAllocateProductNumber
        Return "PRAP"
      Case CDBEnvironment.FunctionParameterTypes.fptAddFastDataEntryPage
        Return "PFDE"
      Case CDBEnvironment.FunctionParameterTypes.fptCopyEventPricingMatrix
        Return "PCPM"
      Case CDBEnvironment.FunctionParameterTypes.fptEnterCancellationFee
        Return "PCEC"
      Case CDBEnvironment.FunctionParameterTypes.fptDuplicateSurvey
        Return "PDSV"
      Case CDBEnvironment.FunctionParameterTypes.fptPaymentPlanDocument
        Return "PPPD"
      Case CDBEnvironment.FunctionParameterTypes.fptCopySegment
        Return "CSAG"
      Case CDBEnvironment.FunctionParameterTypes.fptSetChequeStatus
        Return "PSCS"
      Case CDBEnvironment.FunctionParameterTypes.fptLoadDataUpdates
        Return "PLDU"
      Case CDBEnvironment.FunctionParameterTypes.fptExamResultEntry
        Return "PERE"
      Case CDBEnvironment.FunctionParameterTypes.fptListManagerRandomDataSample
        Return "LMRS"
      Case CDBEnvironment.FunctionParameterTypes.fptReCalculateLoanInterest
        Return "PRLI"
      Case CDBEnvironment.FunctionParameterTypes.fptExamChangeCentre
        Return "EXVC"
      Case CDBEnvironment.FunctionParameterTypes.fptCopyAppeal
        Return "PCAP"
      Case CDBEnvironment.FunctionParameterTypes.fptCLIBrowser
        Return "PCLI"
      Case CDBEnvironment.FunctionParameterTypes.fptDuplicateMeeting
        Return "MTDP"
      Case CDBEnvironment.FunctionParameterTypes.fptShareExamUnit
        Return "PSEU"
      Case CDBEnvironment.FunctionParameterTypes.fptPOPAnalysis
        Return "PRPO"
      Case CDBEnvironment.FunctionParameterTypes.fptExamCertificateReprint
        Return "PXCR"
      Case CDBEnvironment.FunctionParameterTypes.fptCopySegmentCriteria
        Return "CSCR"
      Case CDBEnvironment.FunctionParameterTypes.fptWorkstreamGroupActions
        Return "PWSA"
      Case CDBEnvironment.FunctionParameterTypes.fptActionChangeReasons
        Return "FACR"
      Case CDBEnvironment.FunctionParameterTypes.fptExamScheduleWorkstreams
        Return "ESWG"
      Case CDBEnvironment.FunctionParameterTypes.fptSOCancellation
        Return "MCSO"
      Case CDBEnvironment.FunctionParameterTypes.fptExamCertificates
        Return "MCEC"
      Case CDBEnvironment.FunctionParameterTypes.fptStandardFields
        Return "MCSF"
      Case CDBEnvironment.FunctionParameterTypes.fptMiscFields
        Return "MCMF"
      Case CDBEnvironment.FunctionParameterTypes.fptSelectionTester
        Return "MCST"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysis
        Return "PSSA"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisDetailed
        Return "PSSD"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisSummary
        Return "PSSS"
      Case CDBEnvironment.FunctionParameterTypes.fptFASLPartRefund
        Return "PSPR"
      Case CDBEnvironment.FunctionParameterTypes.fptGAReprintTaxClaim
        Return "PRTC"
      Case Else
        Return ""
    End Select
  End Function

  Public Shared Function GetFPTFromPageCode(ByRef pPageCode As String) As CDBEnvironment.FunctionParameterTypes
    Select Case pPageCode
      Case "PAAE"
        Return CDBEnvironment.FunctionParameterTypes.fptActionActivationDate
      Case "PSOU"
        Return CDBEnvironment.FunctionParameterTypes.fptUpdateStandingOrder
      Case "PPSN"
        Return CDBEnvironment.FunctionParameterTypes.fptPayingInSlipNumber
      Case "PDRP"
        Return CDBEnvironment.FunctionParameterTypes.fptReportNumber
      Case "PCHP"
        Return CDBEnvironment.FunctionParameterTypes.fptChangePayer
      Case "PEPR"
        Return CDBEnvironment.FunctionParameterTypes.fptEventProgrammeReport
      Case "PEER"
        Return CDBEnvironment.FunctionParameterTypes.fptEventPersonnelReport
      Case "PEAR"
        Return CDBEnvironment.FunctionParameterTypes.fptEventAttendeeReport
      Case "PIPD"
        Return CDBEnvironment.FunctionParameterTypes.fptInvoicePaymentDue
      Case "PFAD"
        Return CDBEnvironment.FunctionParameterTypes.fptFinancialAdjustment
      Case "PCOS"
        Return CDBEnvironment.FunctionParameterTypes.fptCloseOrganisationSite
      Case "PMBR"
        Return CDBEnvironment.FunctionParameterTypes.fptMoveBranch
      Case "PMGR"
        Return CDBEnvironment.FunctionParameterTypes.fptMoveRegion
      Case "PCUP"
        Return CDBEnvironment.FunctionParameterTypes.fptUpdateContact
      Case "ABPD"
        Return CDBEnvironment.FunctionParameterTypes.fptAppealBudgetPeriod
      Case "PSOG"
        Return CDBEnvironment.FunctionParameterTypes.fptOwnershipGroup
      Case "PCPT"
        Return CDBEnvironment.FunctionParameterTypes.fptConfirmProvisionalTransaction
      Case "PGAD"
        Return CDBEnvironment.FunctionParameterTypes.fptCreateGiftAidDeclaration
      Case "PAPO"
        Return CDBEnvironment.FunctionParameterTypes.fptAuthorisePOPayment
      Case "PARD"
        Return CDBEnvironment.FunctionParameterTypes.fptAdvanceRenewalDate
      Case "PFPR"
        Return CDBEnvironment.FunctionParameterTypes.fptFAPartRefund
      Case "PATT"
        Return CDBEnvironment.FunctionParameterTypes.fptAttachmentList
      Case "PACR"
        Return CDBEnvironment.FunctionParameterTypes.fptCancReason
      Case "PAMC"
        Return CDBEnvironment.FunctionParameterTypes.fptGetMailingCode
      Case "PPMP"
        Return CDBEnvironment.FunctionParameterTypes.fptPayPlanMissedPayments
      Case "PFAR"
        Return CDBEnvironment.FunctionParameterTypes.fptFAReverseRefundOptions
      Case "PSSA"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysis
      Case "PSOR"
        Return CDBEnvironment.FunctionParameterTypes.fptSOReconciliationReport
      Case "PSSD"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisDetailed
      Case "PSSS"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisSummary
      Case "PRDS"
        Return CDBEnvironment.FunctionParameterTypes.fptReportDataSelection
      Case "PECR"
        Return CDBEnvironment.FunctionParameterTypes.fptEventCancellationFAType
      Case "PCSC"
        Return CDBEnvironment.FunctionParameterTypes.fptChangeSubscriptionCommunication
      Case "PUDS"
        Return CDBEnvironment.FunctionParameterTypes.fptUpdatePaymentPlanDetailSource
      Case "PCMP"
        Return CDBEnvironment.FunctionParameterTypes.fptCMTPriceChange
      Case "PCEN"
        Return CDBEnvironment.FunctionParameterTypes.fptCMTEntitlementPriceChange
      Case "ITAD"
        Return CDBEnvironment.FunctionParameterTypes.fptImportTraderApp
      Case "STJP"
        Return CDBEnvironment.FunctionParameterTypes.fptScheduleTask
      Case "LMAU"
        Return CDBEnvironment.FunctionParameterTypes.fptLMAddressUsage
      Case "PACB"
        Return CDBEnvironment.FunctionParameterTypes.fptAddCollectionBoxes
      Case "PARS"
        Return CDBEnvironment.FunctionParameterTypes.fptCancellationReasonAndSource
      Case "APIS"
        Return CDBEnvironment.FunctionParameterTypes.fptAllocatePISToEvent
      Case "APSD"
        Return CDBEnvironment.FunctionParameterTypes.fptAllocatePISToDelegates
      Case "EVDP"
        Return CDBEnvironment.FunctionParameterTypes.fptDuplicateEvent
      Case "PMRN"
        Return CDBEnvironment.FunctionParameterTypes.fptMembershipReinstatement
      Case "PRSD"
        Return CDBEnvironment.FunctionParameterTypes.fptCancellationReasonSourceAndDate
      Case "PAPP"
        Return CDBEnvironment.FunctionParameterTypes.fptEditAppointment
      Case "PRFM"
        Return CDBEnvironment.FunctionParameterTypes.fptRemoveFutureMembershipType
      Case "PLVP"
        Return CDBEnvironment.FunctionParameterTypes.fptLeavePosition
      Case "PMVP"
        Return CDBEnvironment.FunctionParameterTypes.fptMovePosition
      Case "PPAC"
        Return CDBEnvironment.FunctionParameterTypes.fptCancelPaymentPlan
      Case "PPIS"
        Return CDBEnvironment.FunctionParameterTypes.fptPISPrinting
      Case "PEJD"
        Return CDBEnvironment.FunctionParameterTypes.fptScheduledJobDetails
      Case "PNMC"
        Return CDBEnvironment.FunctionParameterTypes.fptNewMailingCode
      Case "PRAP"
        Return CDBEnvironment.FunctionParameterTypes.fptReAllocateProductNumber
      Case "PFDE"
        Return CDBEnvironment.FunctionParameterTypes.fptAddFastDataEntryPage
      Case "PCPM"
        Return CDBEnvironment.FunctionParameterTypes.fptCopyEventPricingMatrix
      Case "PDSV"
        Return CDBEnvironment.FunctionParameterTypes.fptDuplicateSurvey
      Case "PPPD"
        Return CDBEnvironment.FunctionParameterTypes.fptPaymentPlanDocument
      Case "CSAG"
        Return CDBEnvironment.FunctionParameterTypes.fptCopySegment
      Case "PSCS"
        Return CDBEnvironment.FunctionParameterTypes.fptSetChequeStatus
      Case "PERE"
        Return CDBEnvironment.FunctionParameterTypes.fptExamResultEntry
      Case "LMRS"
        Return CDBEnvironment.FunctionParameterTypes.fptListManagerRandomDataSample
      Case "PRLI"
        Return CDBEnvironment.FunctionParameterTypes.fptReCalculateLoanInterest
      Case "EXVC"
        Return CDBEnvironment.FunctionParameterTypes.fptExamChangeCentre
      Case "PCAP"
        Return CDBEnvironment.FunctionParameterTypes.fptCopyAppeal
      Case "PCLI"
        Return CDBEnvironment.FunctionParameterTypes.fptCLIBrowser
      Case "MTDP"
        Return CDBEnvironment.FunctionParameterTypes.fptDuplicateMeeting
      Case "PRPO"
        Return CDBEnvironment.FunctionParameterTypes.fptPOPAnalysis
      Case "CSCR"
        Return CDBEnvironment.FunctionParameterTypes.fptCopySegmentCriteria
      Case "PWSA"
        Return CDBEnvironment.FunctionParameterTypes.fptWorkstreamGroupActions
      Case "FACR"
        Return CDBEnvironment.FunctionParameterTypes.fptActionChangeReasons
      Case "ESWG"
        Return CDBEnvironment.FunctionParameterTypes.fptExamScheduleWorkstreams
      Case "MCSO"
        Return CDBEnvironment.FunctionParameterTypes.fptSOCancellation
      Case "MCEC"
        Return CDBEnvironment.FunctionParameterTypes.fptExamCertificates
      Case "MCSF"
        Return CDBEnvironment.FunctionParameterTypes.fptStandardFields
      Case "MCMF"
        Return CDBEnvironment.FunctionParameterTypes.fptMiscFields
      Case "MCST"
        Return CDBEnvironment.FunctionParameterTypes.fptSelectionTester
      Case "PSSA"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysis
      Case "PSSD"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisDetailed
      Case "PSSS"
        Return CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisSummary
      Case "PSPR"
        Return CDBEnvironment.FunctionParameterTypes.fptFASLPartRefund
      Case "PRTC"
        Return CDBEnvironment.FunctionParameterTypes.fptGAReprintTaxClaim
    End Select
  End Function

  Public Shared Function GetDescriptionFromFPT(ByVal pFPType As CDBEnvironment.FunctionParameterTypes) As String
    Select Case pFPType
      Case CDBEnvironment.FunctionParameterTypes.fptActionActivationDate
        Return "Action Activation Date"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdateStandingOrder
        Return "Update Standing Order"
      Case CDBEnvironment.FunctionParameterTypes.fptPayingInSlipNumber
        Return "Paying In Slip Number"
      Case CDBEnvironment.FunctionParameterTypes.fptCLRStatementDate
        Return "CLR Statement Date"
      Case CDBEnvironment.FunctionParameterTypes.fptInvoicePaymentDue
        Return "Invoice Payment Due"
      Case CDBEnvironment.FunctionParameterTypes.fptReportNumber
        Return "Report Number"
      Case CDBEnvironment.FunctionParameterTypes.fptChangePayer
        Return "Change Payer"
      Case CDBEnvironment.FunctionParameterTypes.fptEventProgrammeReport
        Return "Event Programme Report"
      Case CDBEnvironment.FunctionParameterTypes.fptEventPersonnelReport
        Return "Event Personnel Report"
      Case CDBEnvironment.FunctionParameterTypes.fptEventAttendeeReport
        Return "Event Attendee Report"
      Case CDBEnvironment.FunctionParameterTypes.fptFinancialAdjustment
        Return "Financial Adjustment"
      Case CDBEnvironment.FunctionParameterTypes.fptCloseOrganisationSite
        Return "Close Organisation Site"
      Case CDBEnvironment.FunctionParameterTypes.fptMoveBranch
        Return "Move Branch"
      Case CDBEnvironment.FunctionParameterTypes.fptMoveRegion
        Return "MoveRegion"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdateContact
        Return "Update Contact"
      Case CDBEnvironment.FunctionParameterTypes.fptAppealBudgetPeriod
        Return "Appeal Budget Period"
      Case CDBEnvironment.FunctionParameterTypes.fptOwnershipGroup
        Return "Ownership Group"
      Case CDBEnvironment.FunctionParameterTypes.fptConfirmProvisionalTransaction
        Return "Confirm Provisional Transaction"
      Case CDBEnvironment.FunctionParameterTypes.fptCreateGiftAidDeclaration
        Return "Create Gift Aid Declaration"
      Case CDBEnvironment.FunctionParameterTypes.fptAuthorisePOPayment
        Return "Authorise PO Payment"
      Case CDBEnvironment.FunctionParameterTypes.fptAdvanceRenewalDate
        Return "Advance Renewal Date"
      Case CDBEnvironment.FunctionParameterTypes.fptFAPartRefund
        Return "FA Part Refund"
      Case CDBEnvironment.FunctionParameterTypes.fptAttachmentList
        Return "Attachment List"
      Case CDBEnvironment.FunctionParameterTypes.fptCancReason
        Return "Canc Reason"
      Case CDBEnvironment.FunctionParameterTypes.fptGetMailingCode
        Return "Get Mailing Code"
      Case CDBEnvironment.FunctionParameterTypes.fptPayPlanMissedPayments
        Return "Pay Plan Missed Payments"
      Case CDBEnvironment.FunctionParameterTypes.fptFAReverseRefundOptions
        Return "FA Reverse Refund Options"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysis
        Return "Stock Sales Analysis"
      Case CDBEnvironment.FunctionParameterTypes.fptSOReconciliationReport
        Return "SO Reconciliation Report"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisDetailed
        Return "Stock Sales Analysis Detailed"
      Case CDBEnvironment.FunctionParameterTypes.fptStockSalesAnalysisSummary
        Return "Stock Sales Analysis Summary"
      Case CDBEnvironment.FunctionParameterTypes.fptReportDataSelection
        Return "Report Data Selection"
      Case CDBEnvironment.FunctionParameterTypes.fptEventCancellationFAType
        Return "Event Cancellation FA Type"
      Case CDBEnvironment.FunctionParameterTypes.fptChangeSubscriptionCommunication
        Return "Change Subscription Communication"
      Case CDBEnvironment.FunctionParameterTypes.fptUpdatePaymentPlanDetailSource
        Return "Update Payment Plan Detail Source"
      Case CDBEnvironment.FunctionParameterTypes.fptCMTPriceChange
        Return "CMT Price Change"
      Case CDBEnvironment.FunctionParameterTypes.fptImportTraderApp
        Return "Import Trader App"
      Case CDBEnvironment.FunctionParameterTypes.fptScheduleTask
        Return "Schedule Task"
      Case CDBEnvironment.FunctionParameterTypes.fptLMAddressUsage
        Return "LM Address Usage"
      Case CDBEnvironment.FunctionParameterTypes.fptAddCollectionBoxes
        Return "Add Collection Boxes"
      Case CDBEnvironment.FunctionParameterTypes.fptCancellationReasonAndSource
        Return "Cancellation Reason and Source"
      Case CDBEnvironment.FunctionParameterTypes.fptAllocatePISToEvent
        Return "Allocate Paying-In-Slips to Event"
      Case CDBEnvironment.FunctionParameterTypes.fptAllocatePISToDelegates
        Return "Allocate Paying-In-Slips to Delegates"
      Case CDBEnvironment.FunctionParameterTypes.fptDuplicateEvent
        Return "Duplicate Event"
      Case CDBEnvironment.FunctionParameterTypes.fptMembershipReinstatement
        Return "Reinstate Membership"
      Case CDBEnvironment.FunctionParameterTypes.fptCancellationReasonSourceAndDate
        Return "Cancellation Reason, Source and Date"
      Case CDBEnvironment.FunctionParameterTypes.fptEditAppointment
        Return "Edit Appointment"
      Case CDBEnvironment.FunctionParameterTypes.fptRemoveFutureMembershipType
        Return "Remove Future Membership Type"
      Case CDBEnvironment.FunctionParameterTypes.fptLeavePosition
        Return "Leave Position"
      Case CDBEnvironment.FunctionParameterTypes.fptMovePosition
        Return "Move Position"
      Case CDBEnvironment.FunctionParameterTypes.fptCancelPaymentPlan
        Return "Cancel Payment Plan"
      Case CDBEnvironment.FunctionParameterTypes.fptPISPrinting
        Return "Paying In Slip Printing"
      Case CDBEnvironment.FunctionParameterTypes.fptScheduledJobDetails
        Return "Enter Job Details"
      Case CDBEnvironment.FunctionParameterTypes.fptExamResultEntry
        Return "Enter Exam Results"
      Case CDBEnvironment.FunctionParameterTypes.fptListManagerRandomDataSample
        Return "Random Data Sample"
      Case CDBEnvironment.FunctionParameterTypes.fptReCalculateLoanInterest
        Return "Re-calculate Loan Interest"
      Case CDBEnvironment.FunctionParameterTypes.fptExamChangeCentre
        Return "Change Exam Centre"
      Case CDBEnvironment.FunctionParameterTypes.fptCLIBrowser
        Return "Select Contact"
      Case CDBEnvironment.FunctionParameterTypes.fptDuplicateMeeting
        Return "Duplicate Meeting"
      Case CDBEnvironment.FunctionParameterTypes.fptPOPAnalysis
        Return "PO Payment Reanalysis"
      Case CDBEnvironment.FunctionParameterTypes.fptWorkstreamGroupActions
        Return "Workstream Group Actions"
      Case CDBEnvironment.FunctionParameterTypes.fptExamScheduleWorkstreams
        Return "Workstreams Linked to Exam Schedule"
      Case CDBEnvironment.FunctionParameterTypes.fptSOCancellation
        Return "Standing Order Cancellation"
      Case CDBEnvironment.FunctionParameterTypes.fptExamCertificates
        Return "Mailings Exam Certificates"
      Case CDBEnvironment.FunctionParameterTypes.fptStandardFields
        Return "Mailings Standard Fields"
      Case CDBEnvironment.FunctionParameterTypes.fptMiscFields
        Return "Mailings Misc Fields"
      Case CDBEnvironment.FunctionParameterTypes.fptSelectionTester
        Return "Mailings Selection Tester"
      Case CDBEnvironment.FunctionParameterTypes.fptFASLPartRefund
        Return "FA Sales Ledger Part Refund"
      Case Else
        Return "Unknown Job"
    End Select

  End Function
End Class