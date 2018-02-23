
Public Class BaseFinancialMenu
  Inherits ContextMenuStrip

  Protected Const MAX_EVENT_GROUP_INDEX As Integer = 4

  Private mvParent As MaintenanceParentForm
  Private mvChangePPDs As Boolean
  Private mvAdjustmentBatchNumber As Integer
  Private mvAdjustmentTransactionNumber As Integer
  Private mvAdjustmentWasBatchNumber As Integer
  Private mvAdjustmentWasTransactionNumber As Integer
  Private mvLineType As String

  Protected mvEventNumber(MAX_EVENT_GROUP_INDEX) As Integer
  Protected mvEventInfo(MAX_EVENT_GROUP_INDEX) As CareEventInfo
  Protected mvReadOnly As Boolean
  Protected mvAnalysisRow As DataRow
  Protected mvTargetBatchNumber As Integer
  Protected mvTargetTransactionNumber As Integer
  Protected mvBatchNumber As Integer
  Protected mvTransactionNumber As Integer
  Protected mvTargetLineNumber As Integer
  Protected mvDataRow As DataRow
  Protected mvContactInfo As ContactInfo
  Protected mvDataType As CareServices.XMLContactDataSelectionTypes
  Protected mvMenuContextType As MenuContextTypes
  Protected mvNumber As Integer
  Protected mvMultiSelect As Boolean
  Protected mvSubMenu As Boolean
  Protected mvDisplayTransactionsAllocationType As String = Nothing
  Protected mvCanPartRefund As Boolean = False

  Public Event MenuSelected(ByVal pItem As FinancialMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu)

  Public Enum MenuContextTypes
    mctMain
    mctAnalysis
    mctFinder
  End Enum

  Public Enum FinancialMenuItems
    fmiNone = -1
    fmiNew = 0
    fmiEdit
    fmiGoToSO
    fmiGoToDD
    fmiGoToCC
    fmiGoToMembership
    fmiGoToCovenant
    fmiGoToPayPlan
    fmiGoToTransaction
    fmiGoToContact
    fmiGoToChanges
    fmiGoToChangedBy
    fmiGoToBackOrders
    fmiGoToLinks
    fmiGoToDespatch
    fmiChangeMembershipType
    fmiCancel
    fmiFutureCancel
    fmiChangeCancel
    fmiReprintNumbers
    fmiReprintMembershipCard
    fmiPaymentPlanConversion
    fmiPaymentPlanMaintenance
    fmiPaymentPlanPrint
    fmiReinstateMembership
    fmiAddMember
    fmiMove
    fmiReverse
    fmiRefund
    fmiAnalysis
    fmiFutureMembershipType
    fmiConfirmTransaction
    fmiChangePayer
    fmiGoToPreTaxPledge
    fmiGoToPostTaxPledge
    fmiReinstateAutoPayMethod
    fmiGoToBankAccount
    fmiSkipPayment
    fmiAddGiftAidDeclaration
    fmiAdvanceRenewalDate
    fmiConfirmPaymentPlan
    fmiReinstatePaymentPlan
    fmiSubChangeCommunication
    fmiReplaceMember
    fmiReinstateProvisionalTrans
    fmiAddEventFinancialLink
    fmiAddEventFinancialLink2
    fmiAddEventFinancialLink3
    fmiAddEventFinancialLink4
    fmiAddEventFinancialLink5
    fmiRemoveEventFinancialLink
    fmiRemoveEventFinancialLink2
    fmiRemoveEventFinancialLink3
    fmiRemoveEventFinancialLink4
    fmiRemoveEventFinancialLink5
    fmiGoToEvent
    fmiGoToEvent2
    fmiGoToEvent3
    fmiGoToEvent4
    fmiGoToEvent5
    fmiAmendDueDate
    fmiRemoveAllocations
    fmiRefundInAdvance
    fmiReverseInAdvance
    fmiAmendMembership
    fmiEditNotes
    fmiAmendPurchaseOrder
    fmiAuthorise
    fmiReinstate
    fmiSupplementaryInformation
    fmiAmendBooking
    fmiViewMailingDocument
    fmiRedoFulfilment
    fmiDeleteMailingDocument
    fmiUnfulfillMailingDocument
    fmiReissueCheque
    fmiUnlockFundraisingRequest
    fmiAddFundraisingPaymentLink
    fmiNewAdHocAction
    fmiNewActionFromTemplate
    fmiGoToActions
    fmiEditUnprocessedTransactionNotes
    fmiChangeChequePayee
    fmiCancelEventBooking
    fmiAuthorisePurchaseOrder
    fmiAddPurchaseOrderPaymentReceipt
    fmiChequeSetStatus
    fmiGoToLoan
    fmiRecalcLoanInterest
    fmiCancelExamBooking
    fmiChangeExamCentre
    fmiEditReference
    fmiProduceMembershipCard
    fmiChangeInvoiceAddress
    fmiDisplayTransactions
    fmiFutureRenewalAmount
    fmiChangeClaimDate
    fmiPrintReceipt
    'Sub menu items
    fmiDisplayTransactionsAll
    fmiDisplayTransactionsUnallocatedOnly
    fmiDisplayTransactionsFullyAllocatedOnly
    fmiPreviewInvoice
    fmiCancelPOP
    fmiGoToPopRevChangedBy
    fmiGoToPopRevGoToChanges
    fmiPOPAnalysis
  End Enum
  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New()
    Try
      pParent.SuspendLayout()
      Me.SuspendLayout()
      mvParent = pParent
      mvDataType = pDST
      mvMenuContextType = MenuContextTypes.mctMain
      mvContactInfo = pContactInfo
      Dim vMenuItems As New CollectionList(Of MenuToolbarCommand)

      With vMenuItems
        Dim vNewImage As Image = Nothing
        Dim vEditImage As Image = Nothing
        vNewImage = AppHelper.ImageProvider.NewOtherImages16.Images("New")
        vEditImage = AppHelper.ImageProvider.NewOtherImages16.Images("Edit")
        .Add(FinancialMenuItems.fmiNew.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiNew.ToString, ControlText.MnuFinancialNew, FinancialMenuItems.fmiNew, "", vNewImage))
        .Add(FinancialMenuItems.fmiEdit.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiEdit.ToString, ControlText.MnuFinancialEdit, FinancialMenuItems.fmiEdit, "", vEditImage))
        .Add(FinancialMenuItems.fmiGoToSO.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToSO.ToString, ControlText.MnuFinancialGoToSO, FinancialMenuItems.fmiGoToSO, ""))
        .Add(FinancialMenuItems.fmiGoToDD.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToDD.ToString, ControlText.MnuFinancialGoToDD, FinancialMenuItems.fmiGoToDD, ""))
        .Add(FinancialMenuItems.fmiGoToCC.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToCC.ToString, ControlText.MnuFinancialGoToCC, FinancialMenuItems.fmiGoToCC, ""))
        .Add(FinancialMenuItems.fmiGoToMembership.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToMembership.ToString, ControlText.MnuFinancialGoToMembership, FinancialMenuItems.fmiGoToMembership, ""))
        .Add(FinancialMenuItems.fmiGoToCovenant.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToCovenant.ToString, ControlText.MnuFinancialGoToCovenant, FinancialMenuItems.fmiGoToCovenant, ""))
        .Add(FinancialMenuItems.fmiGoToPayPlan.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToPayPlan.ToString, ControlText.MnuFinancialGoToPayPlan, FinancialMenuItems.fmiGoToPayPlan, ""))
        .Add(FinancialMenuItems.fmiGoToTransaction.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToTransaction.ToString, ControlText.MnuFinancialGoToTransaction, FinancialMenuItems.fmiGoToTransaction, ""))
        .Add(FinancialMenuItems.fmiGoToContact.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToContact.ToString, ControlText.MnuFinancialGoToContact, FinancialMenuItems.fmiGoToContact, ""))
        .Add(FinancialMenuItems.fmiGoToChanges.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToChanges.ToString, ControlText.MnuFinancialGoToChanges, FinancialMenuItems.fmiGoToChanges, ""))
        .Add(FinancialMenuItems.fmiGoToChangedBy.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToChangedBy.ToString, ControlText.MnuFinancialGoToChangedBy, FinancialMenuItems.fmiGoToChangedBy, ""))
        .Add(FinancialMenuItems.fmiGoToBackOrders.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToBackOrders.ToString, ControlText.MnuFinancialGoToBackOrders, FinancialMenuItems.fmiGoToBackOrders, ""))
        .Add(FinancialMenuItems.fmiGoToLinks.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToLinks.ToString, ControlText.MnuFinancialGoToLinks, FinancialMenuItems.fmiGoToLinks, ""))
        .Add(FinancialMenuItems.fmiGoToDespatch.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToDespatch.ToString, ControlText.MnuFinancialGoToDespatch, FinancialMenuItems.fmiGoToDespatch, ""))
        .Add(FinancialMenuItems.fmiChangeMembershipType.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeMembershipType.ToString, ControlText.MnuFinancialChangeMembershipType, FinancialMenuItems.fmiChangeMembershipType, "CDFPCM"))
        .Add(FinancialMenuItems.fmiCancel.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiCancel.ToString, ControlText.MnuFinancialCancel, FinancialMenuItems.fmiCancel, "CDFPCN"))
        .Add(FinancialMenuItems.fmiFutureCancel.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiFutureCancel.ToString, ControlText.MnuFinancialFutureCancel, FinancialMenuItems.fmiFutureCancel, "CDFPFC"))
        .Add(FinancialMenuItems.fmiChangeCancel.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeCancel.ToString, ControlText.MnuFinancialChangeCancel, FinancialMenuItems.fmiChangeCancel, "CDFPCC"))
        .Add(FinancialMenuItems.fmiReprintNumbers.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReprintNumbers.ToString, ControlText.MnuFinancialReprintNumbers, FinancialMenuItems.fmiReprintNumbers, "CDFPRN"))
        .Add(FinancialMenuItems.fmiReprintMembershipCard.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReprintMembershipCard.ToString, ControlText.MnuFinancialReprintMembershipCard, FinancialMenuItems.fmiReprintMembershipCard, "CDFPRC"))
        .Add(FinancialMenuItems.fmiPaymentPlanConversion.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPaymentPlanConversion.ToString, ControlText.MnuFinancialPaymentPlanConversion, FinancialMenuItems.fmiPaymentPlanConversion, "CDFPPC"))
        .Add(FinancialMenuItems.fmiPaymentPlanMaintenance.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPaymentPlanMaintenance.ToString, ControlText.MnuFinancialPaymentPlanMaintenance, FinancialMenuItems.fmiPaymentPlanMaintenance, "CDFPPM"))
        .Add(FinancialMenuItems.fmiPaymentPlanPrint.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPaymentPlanPrint.ToString, ControlText.MnuFinancialPaymentPlanPrint, FinancialMenuItems.fmiPaymentPlanPrint, "CDFPPP"))
        .Add(FinancialMenuItems.fmiReinstateMembership.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReinstateMembership.ToString, ControlText.MnuFinancialReinstateMembership, FinancialMenuItems.fmiReinstateMembership, "CDFPRM"))
        .Add(FinancialMenuItems.fmiAddMember.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddMember.ToString, ControlText.MnuFinancialAddMember, FinancialMenuItems.fmiAddMember, "CDFPME"))
        .Add(FinancialMenuItems.fmiMove.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiMove.ToString, ControlText.MnuFinancialMove, FinancialMenuItems.fmiMove, "CDFPAM"))
        .Add(FinancialMenuItems.fmiReverse.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReverse.ToString, ControlText.MnuFinancialReverse, FinancialMenuItems.fmiReverse, "CDFPAR"))
        .Add(FinancialMenuItems.fmiRefund.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRefund.ToString, ControlText.MnuFinancialRefund, FinancialMenuItems.fmiRefund, "CDFPAF"))
        .Add(FinancialMenuItems.fmiAnalysis.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAnalysis.ToString, ControlText.MnuFinancialAnalysis, FinancialMenuItems.fmiAnalysis, "CDFPAA"))
        .Add(FinancialMenuItems.fmiFutureMembershipType.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiFutureMembershipType.ToString, ControlText.MnuFinancialFutureMembershipType, FinancialMenuItems.fmiFutureMembershipType, "CDFPFM"))
        .Add(FinancialMenuItems.fmiConfirmTransaction.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiConfirmTransaction.ToString, ControlText.MnuFinancialConfirmTransaction, FinancialMenuItems.fmiConfirmTransaction, "SCFPCT"))
        .Add(FinancialMenuItems.fmiChangePayer.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangePayer.ToString, ControlText.MnuFinancialChangePayer, FinancialMenuItems.fmiChangePayer, "CDFPCP"))
        .Add(FinancialMenuItems.fmiGoToPreTaxPledge.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToPreTaxPledge.ToString, ControlText.MnuFinancialGoToPayrollGivingPledge, FinancialMenuItems.fmiGoToPreTaxPledge, ""))
        .Add(FinancialMenuItems.fmiGoToPostTaxPledge.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToPostTaxPledge.ToString, ControlText.MnuFinancialGoToPayrollGivingPledge, FinancialMenuItems.fmiGoToPostTaxPledge, ""))
        .Add(FinancialMenuItems.fmiReinstateAutoPayMethod.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReinstateAutoPayMethod.ToString, ControlText.MnuFinancialReinstateAutoPayMethod, FinancialMenuItems.fmiReinstateAutoPayMethod, "CDFPRD"))
        .Add(FinancialMenuItems.fmiGoToBankAccount.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToBankAccount.ToString, ControlText.MnuFinancialGoToBankAccount, FinancialMenuItems.fmiGoToBankAccount, ""))
        .Add(FinancialMenuItems.fmiSkipPayment.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiSkipPayment.ToString, ControlText.MnuFinancialSkipPayment, FinancialMenuItems.fmiSkipPayment, "CDFPSP"))
        .Add(FinancialMenuItems.fmiAddGiftAidDeclaration.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddGiftAidDeclaration.ToString, ControlText.MnuFinancialAddGiftAidDeclaration, FinancialMenuItems.fmiAddGiftAidDeclaration, "CDFPAG"))
        .Add(FinancialMenuItems.fmiAdvanceRenewalDate.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAdvanceRenewalDate.ToString, ControlText.MnuFinancialAdvanceRenewalDate, FinancialMenuItems.fmiAdvanceRenewalDate, "CDFPAD"))
        .Add(FinancialMenuItems.fmiConfirmPaymentPlan.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiConfirmPaymentPlan.ToString, ControlText.MnuFinancialConfirmPaymentPlan, FinancialMenuItems.fmiConfirmPaymentPlan, "CDFPNP"))
        .Add(FinancialMenuItems.fmiReinstatePaymentPlan.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReinstatePaymentPlan.ToString, ControlText.MnuFinancialReinstatePaymentPlan, FinancialMenuItems.fmiReinstatePaymentPlan, "CDFPRP"))
        .Add(FinancialMenuItems.fmiSubChangeCommunication.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiSubChangeCommunication.ToString, ControlText.MnuFinancialSubChangeCommunication, FinancialMenuItems.fmiSubChangeCommunication, "CDFPSC"))
        .Add(FinancialMenuItems.fmiReplaceMember.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReplaceMember.ToString, ControlText.MnuFinancialReplaceMember, FinancialMenuItems.fmiReplaceMember, "CDFPRL"))
        .Add(FinancialMenuItems.fmiReinstateProvisionalTrans.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReinstateProvisionalTrans.ToString, ControlText.MnuFinancialReinstateProvisionalTrans, FinancialMenuItems.fmiReinstateProvisionalTrans, ""))

        .Add(FinancialMenuItems.fmiAddEventFinancialLink.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddEventFinancialLink.ToString, GetEventMenuText(ControlText.MnuFinancialAddEventFinancialLink, 0), FinancialMenuItems.fmiAddEventFinancialLink, "CDEVFL"))
        .Add(FinancialMenuItems.fmiAddEventFinancialLink2.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddEventFinancialLink2.ToString, GetEventMenuText(ControlText.MnuFinancialAddEventFinancialLink, 1), FinancialMenuItems.fmiAddEventFinancialLink2, "CDEVFL"))
        .Add(FinancialMenuItems.fmiAddEventFinancialLink3.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddEventFinancialLink3.ToString, GetEventMenuText(ControlText.MnuFinancialAddEventFinancialLink, 2), FinancialMenuItems.fmiAddEventFinancialLink3, "CDEVFL"))
        .Add(FinancialMenuItems.fmiAddEventFinancialLink4.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddEventFinancialLink4.ToString, GetEventMenuText(ControlText.MnuFinancialAddEventFinancialLink, 3), FinancialMenuItems.fmiAddEventFinancialLink4, "CDEVFL"))
        .Add(FinancialMenuItems.fmiAddEventFinancialLink5.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddEventFinancialLink5.ToString, GetEventMenuText(ControlText.MnuFinancialAddEventFinancialLink, 4), FinancialMenuItems.fmiAddEventFinancialLink5, "CDEVFL"))

        .Add(FinancialMenuItems.fmiRemoveEventFinancialLink.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveEventFinancialLink.ToString, GetEventMenuText(ControlText.MnuFinancialRemoveEventFinancialLink, 0), FinancialMenuItems.fmiRemoveEventFinancialLink, "CDEVFL"))
        .Add(FinancialMenuItems.fmiRemoveEventFinancialLink2.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveEventFinancialLink2.ToString, GetEventMenuText(ControlText.MnuFinancialRemoveEventFinancialLink, 1), FinancialMenuItems.fmiRemoveEventFinancialLink2, "CDEVFL"))
        .Add(FinancialMenuItems.fmiRemoveEventFinancialLink3.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveEventFinancialLink3.ToString, GetEventMenuText(ControlText.MnuFinancialRemoveEventFinancialLink, 2), FinancialMenuItems.fmiRemoveEventFinancialLink3, "CDEVFL"))
        .Add(FinancialMenuItems.fmiRemoveEventFinancialLink4.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveEventFinancialLink4.ToString, GetEventMenuText(ControlText.MnuFinancialRemoveEventFinancialLink, 3), FinancialMenuItems.fmiRemoveEventFinancialLink4, "CDEVFL"))
        .Add(FinancialMenuItems.fmiRemoveEventFinancialLink5.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveEventFinancialLink5.ToString, GetEventMenuText(ControlText.MnuFinancialRemoveEventFinancialLink, 4), FinancialMenuItems.fmiRemoveEventFinancialLink5, "CDEVFL"))

        .Add(FinancialMenuItems.fmiGoToEvent.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToEvent.ToString, GetEventMenuText(ControlText.MnuFinancialGoToEvent, 0), FinancialMenuItems.fmiGoToEvent, ""))
        .Add(FinancialMenuItems.fmiGoToEvent2.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToEvent2.ToString, GetEventMenuText(ControlText.MnuFinancialGoToEvent, 1), FinancialMenuItems.fmiGoToEvent2, ""))
        .Add(FinancialMenuItems.fmiGoToEvent3.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToEvent3.ToString, GetEventMenuText(ControlText.MnuFinancialGoToEvent, 2), FinancialMenuItems.fmiGoToEvent3, ""))
        .Add(FinancialMenuItems.fmiGoToEvent4.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToEvent4.ToString, GetEventMenuText(ControlText.MnuFinancialGoToEvent, 3), FinancialMenuItems.fmiGoToEvent4, ""))
        .Add(FinancialMenuItems.fmiGoToEvent5.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToEvent5.ToString, GetEventMenuText(ControlText.MnuFinancialGoToEvent, 4), FinancialMenuItems.fmiGoToEvent5, ""))

        .Add(FinancialMenuItems.fmiAmendDueDate.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAmendDueDate.ToString, ControlText.MnuFinancialAmendDueDate, FinancialMenuItems.fmiAmendDueDate, "CDAPAP"))
        .Add(FinancialMenuItems.fmiRemoveAllocations.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRemoveAllocations.ToString, ControlText.MnuFinancialRemoveAllocations, FinancialMenuItems.fmiRemoveAllocations, "CDAPAL"))
        .Add(FinancialMenuItems.fmiRefundInAdvance.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRefundInAdvance.ToString, ControlText.MnuFinancialRefundInAdvance, FinancialMenuItems.fmiRefundInAdvance, "CDFPIF"))
        .Add(FinancialMenuItems.fmiReverseInAdvance.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReverseInAdvance.ToString, ControlText.MnuFinancialReverseInAdvance, FinancialMenuItems.fmiReverseInAdvance, "CDFPIR"))
        .Add(FinancialMenuItems.fmiAmendMembership.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAmendMembership.ToString, ControlText.MnuFinancialAmendMembership, FinancialMenuItems.fmiAmendMembership, ""))
        .Add(FinancialMenuItems.fmiEditNotes.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiEditNotes.ToString, ControlText.MnuFinancialEditNotes, FinancialMenuItems.fmiEditNotes, "SCFPEN"))
        .Add(FinancialMenuItems.fmiAmendPurchaseOrder.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAmendPurchaseOrder.ToString, ControlText.MnuFinancialAmendPurchaseOrder, FinancialMenuItems.fmiAmendPurchaseOrder, ""))
        .Add(FinancialMenuItems.fmiAuthorise.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAuthorise.ToString, ControlText.MnuFinancialAuthorise, FinancialMenuItems.fmiAuthorise, "SCFPAA"))
        .Add(FinancialMenuItems.fmiReinstate.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReinstate.ToString, ControlText.MnuFinancialReinstate, FinancialMenuItems.fmiReinstate, ""))
        .Add(FinancialMenuItems.fmiSupplementaryInformation.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiSupplementaryInformation.ToString, ControlText.MnuFinancialSupplementaryInformation, FinancialMenuItems.fmiSupplementaryInformation, "CDEVSI"))
        .Add(FinancialMenuItems.fmiAmendBooking.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAmendBooking.ToString, ControlText.MnuFinancialAmendBooking, FinancialMenuItems.fmiAmendBooking, "SCFPEB"))
        .Add(FinancialMenuItems.fmiViewMailingDocument.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiViewMailingDocument.ToString, ControlText.MnuFinancialViewMailingDocument, FinancialMenuItems.fmiViewMailingDocument, ""))
        .Add(FinancialMenuItems.fmiRedoFulfilment.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRedoFulfilment.ToString, ControlText.MnuFinancialRedoFulfilment, FinancialMenuItems.fmiRedoFulfilment, ""))
        .Add(FinancialMenuItems.fmiDeleteMailingDocument.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDeleteMailingDocument.ToString, ControlText.MnuFinancialDeleteMailingDocument, FinancialMenuItems.fmiDeleteMailingDocument, "SCMPDM"))
        .Add(FinancialMenuItems.fmiUnfulfillMailingDocument.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiUnfulfillMailingDocument.ToString, ControlText.MnuFinancialUnfulfillMailingDocument, FinancialMenuItems.fmiUnfulfillMailingDocument, "SCMPUM"))
        .Add(FinancialMenuItems.fmiReissueCheque.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiReissueCheque.ToString, ControlText.MnuFinancialReissueCheque, FinancialMenuItems.fmiReissueCheque, "SCFPRC"))
        .Add(FinancialMenuItems.fmiUnlockFundraisingRequest.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiUnlockFundraisingRequest.ToString, ControlText.MnuFinancialUnlock, FinancialMenuItems.fmiUnlockFundraisingRequest, "SCFPUF"))
        .Add(FinancialMenuItems.fmiAddFundraisingPaymentLink.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddFundraisingPaymentLink.ToString, ControlText.MnuFinancialAddFundPaymentLink, FinancialMenuItems.fmiAddFundraisingPaymentLink, "SCFPAF"))
        .Add(FinancialMenuItems.fmiNewAdHocAction.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiNewAdHocAction.ToString, ControlText.MnuActionNew, FinancialMenuItems.fmiNewAdHocAction, "SCBMAC"))
        .Add(FinancialMenuItems.fmiNewActionFromTemplate.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiNewActionFromTemplate.ToString, ControlText.MnuBrowserActionsNewFromTemplate, FinancialMenuItems.fmiNewActionFromTemplate, "SCBMAC"))
        .Add(FinancialMenuItems.fmiGoToActions.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToActions.ToString, ControlText.MnuFinancialGoToActions, FinancialMenuItems.fmiGoToActions, "SCBMAC"))
        .Add(FinancialMenuItems.fmiEditUnprocessedTransactionNotes.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiEditUnprocessedTransactionNotes.ToString, ControlText.MnuFinancialEditNotes, FinancialMenuItems.fmiEditUnprocessedTransactionNotes, "SCFPEU"))
        .Add(FinancialMenuItems.fmiChangeChequePayee.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeChequePayee.ToString, ControlText.MnuFinancialChangeChequePayee, FinancialMenuItems.fmiChangeChequePayee, "SCFPCQ"))
        .Add(FinancialMenuItems.fmiCancelEventBooking.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiCancelEventBooking.ToString, ControlText.MnuFinancialCancelEventBooking, FinancialMenuItems.fmiCancelEventBooking, "CDEVCA"))
        .Add(FinancialMenuItems.fmiAuthorisePurchaseOrder.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAuthorisePurchaseOrder.ToString, ControlText.MnuFinancialAuthorise, FinancialMenuItems.fmiAuthorisePurchaseOrder, ""))
        .Add(FinancialMenuItems.fmiAddPurchaseOrderPaymentReceipt.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiAddPurchaseOrderPaymentReceipt.ToString, ControlText.MnuAddPurchaseOrderPaymentReceipt, FinancialMenuItems.fmiAddPurchaseOrderPaymentReceipt, "SCFPAR"))
        .Add(FinancialMenuItems.fmiChequeSetStatus.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChequeSetStatus.ToString, ControlText.MnuChequeSetStatus, FinancialMenuItems.fmiChequeSetStatus, "SCFPSS"))
        .Add(FinancialMenuItems.fmiGoToLoan.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToLoan.ToString, ControlText.MnuFinancialGoToLoan, FinancialMenuItems.fmiGoToLoan, ""))
        .Add(FinancialMenuItems.fmiRecalcLoanInterest.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiRecalcLoanInterest.ToString, ControlText.MnuFinancialRecalcLoanInterest, FinancialMenuItems.fmiRecalcLoanInterest, "SCFPCI"))
        .Add(FinancialMenuItems.fmiCancelExamBooking.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiCancelExamBooking.ToString, ControlText.MnuFinancialCancelExamBooking, FinancialMenuItems.fmiCancelExamBooking, "SCFPCX"))
        .Add(FinancialMenuItems.fmiChangeExamCentre.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeExamCentre.ToString, ControlText.MnuFinancialChangeExamCentre, FinancialMenuItems.fmiChangeExamCentre, "SCFPXC"))
        .Add(FinancialMenuItems.fmiEditReference.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiEditReference.ToString, ControlText.MnuFinancialEditReference, FinancialMenuItems.fmiEditReference, "SCFPER"))
        .Add(FinancialMenuItems.fmiProduceMembershipCard.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiProduceMembershipCard.ToString, ControlText.MnuFinancialProduceMembershipCard, FinancialMenuItems.fmiProduceMembershipCard, "SCFPMC"))
        .Add(FinancialMenuItems.fmiChangeInvoiceAddress.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeInvoiceAddress.ToString, ControlText.MnuFinancialChangeInvoiceAddress, FinancialMenuItems.fmiChangeInvoiceAddress, "SCFPCA"))
        .Add(FinancialMenuItems.fmiDisplayTransactions.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactions.ToString, ControlText.MnuDisplayTransactions, FinancialMenuItems.fmiDisplayTransactions, "SCFPDT"))
        .Add(FinancialMenuItems.fmiFutureRenewalAmount.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiFutureRenewalAmount.ToString, ControlText.MnuFinancialFutureMemRenewalAmount, FinancialMenuItems.fmiFutureRenewalAmount, "SCFPFR"))
        .Add(FinancialMenuItems.fmiChangeClaimDate.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiChangeClaimDate.ToString, ControlText.MnuFinancialChangeClaimDate, FinancialMenuItems.fmiChangeClaimDate, "SCFPCD"))
        .Add(FinancialMenuItems.fmiPrintReceipt.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPrintReceipt.ToString, ControlText.MnuFinancialPrintReceipt, FinancialMenuItems.fmiPrintReceipt, "SCFPGR"))

        'Sub-menu items below
        .Add(FinancialMenuItems.fmiDisplayTransactionsAll.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsAll.ToString, ControlText.MnuDisplayTransactionsAll, FinancialMenuItems.fmiDisplayTransactionsAll, ""))
        .Add(FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly.ToString, ControlText.MnuDisplayTransactionsUnallocatedOnly, FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly, ""))
        .Add(FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly.ToString, ControlText.MnuDisplayTransactionsFullyAllocatedOnly, FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly, ""))
        .Add(FinancialMenuItems.fmiPreviewInvoice.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPreviewInvoice.ToString, ControlText.MnuPreviewInvoice, FinancialMenuItems.fmiPreviewInvoice, "SCFPPI"))
        .Add(FinancialMenuItems.fmiCancelPOP.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiCancelPOP.ToString, ControlText.MnuFinancialCancelPOP, FinancialMenuItems.fmiCancelPOP, "SCFPOP"))
        .Add(FinancialMenuItems.fmiGoToPopRevChangedBy.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToPopRevChangedBy.ToString, ControlText.MnuFinancialGoToChangedBy, FinancialMenuItems.fmiGoToPopRevChangedBy, ""))
        .Add(FinancialMenuItems.fmiGoToPopRevGoToChanges.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiGoToPopRevGoToChanges.ToString, ControlText.MnuFinancialGoToChanges, FinancialMenuItems.fmiGoToPopRevGoToChanges, ""))
        .Add(FinancialMenuItems.fmiPOPAnalysis.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiPOPAnalysis.ToString, ControlText.MnuFinancialPOPAnalysis, FinancialMenuItems.fmiPOPAnalysis, "SCFPPA"))
      End With
      For Each vItem As MenuToolbarCommand In vMenuItems
        vItem.OnClick = AddressOf MainMenuHandler
        Me.Items.Add(vItem.MenuStripItem)
      Next
      'Add Display Transactions drop down menu items
      Dim vDisplayTransactionsDropDownItems As New CollectionList(Of MenuToolbarCommand)
      With vDisplayTransactionsDropDownItems
        .Add(FinancialMenuItems.fmiDisplayTransactionsAll.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsAll.ToString, ControlText.MnuDisplayTransactionsAll, FinancialMenuItems.fmiDisplayTransactionsAll))
        .Add(FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly.ToString, ControlText.MnuDisplayTransactionsUnallocatedOnly, FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly))
        .Add(FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly.ToString, New MenuToolbarCommand(FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly.ToString, ControlText.MnuDisplayTransactionsFullyAllocatedOnly, FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly))
      End With
      Dim vDisplayTransactionsMenuItem As ToolStripMenuItem = DirectCast(Me.Items(FinancialMenuItems.fmiDisplayTransactions), ToolStripMenuItem)
      With vDisplayTransactionsMenuItem.DropDownItems
        For Each vNewItem As MenuToolbarCommand In vDisplayTransactionsDropDownItems
          vNewItem.OnClick = AddressOf MainMenuHandler
          .Add(vNewItem.MenuStripItem)
        Next
      End With

      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      MenuToolbarCommand.SetAccessControl(vMenuItems)
    Finally
      Me.ResumeLayout()
      pParent.ResumeLayout()
    End Try
  End Sub

  Private Function GetEventMenuText(ByVal pText As String, ByVal pIndex As Integer) As String
    If DataHelper.EventGroups.Count > pIndex Then
      Return String.Format(pText, DataHelper.EventGroups(pIndex).GroupName)
    Else
      Return ""
    End If
  End Function

  Private Sub MainMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MenuHandler(DirectCast(sender, ToolStripMenuItem), CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, FinancialMenuItems))
  End Sub
  Protected Overridable Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FinancialMenuItems)
    Dim vCursor As New BusyCursor
    Try
      RaiseEvent MenuSelected(pItem, mvDataRow, mvChangePPDs, Me)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Protected Overridable Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    SetVisibleItems(e)
  End Sub

  Protected Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Dim vCursor As New BusyCursor
    Dim vCanEdit As Boolean
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      e.Cancel = False
      vCanEdit = mvContactInfo IsNot Nothing AndAlso mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite AndAlso Not mvReadOnly
      Dim vHasAccessRights As Boolean = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
      Dim vShowItems(Me.Items.Count) As Boolean
      vShowItems(FinancialMenuItems.fmiNew) = False
      Select Case mvDataType

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings
          If mvDataRow IsNot Nothing Then
            If vCanEdit Then
              vShowItems(FinancialMenuItems.fmiCancelEventBooking) = (mvDataRow.Item("CancellationReason").ToString.Length = 0 AndAlso mvDataRow.Item("BookingStatus").ToString <> "U")
              If mvDataRow.Item("BookingStatus").ToString <> "C" AndAlso mvDataRow.Item("BookingStatus").ToString <> "U" Then
                Dim vList As New ParameterList(True, False) 'Do not want SystemColumns
                vList("SmartClient") = "Y"
                Dim vBatchNumber As Integer = IntegerValue(mvDataRow.Item("BatchNumber").ToString)
                Dim vTransactionNumber As Integer = IntegerValue(mvDataRow.Item("TransactionNumber").ToString)
                Dim vRow As DataRow = Nothing
                If vBatchNumber > 0 AndAlso vTransactionNumber > 0 Then vRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vBatchNumber, vTransactionNumber, vList))
                If vRow IsNot Nothing Then
                  vShowItems(FinancialMenuItems.fmiAmendBooking) = BooleanValue(vRow.Item("PostedToNominal").ToString)
                  vShowItems(FinancialMenuItems.fmiAnalysis) = BooleanValue(vRow.Item("PostedToNominal").ToString)
                ElseIf vBatchNumber = 0 Then
                  vShowItems(FinancialMenuItems.fmiAmendBooking) = mvDataRow.Item("BatchNumber").ToString.Length = 0
                End If
              End If
            End If
            If SetGotoEventMenuItem(mvDataRow, FinancialMenuItems.fmiGoToEvent) Then vShowItems(FinancialMenuItems.fmiGoToEvent) = True
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings
          If mvDataRow IsNot Nothing Then
            If vCanEdit Then vShowItems(FinancialMenuItems.fmiCancelEventBooking) = mvDataRow.Item("CancellationReason").ToString.Length = 0
            If SetGotoEventMenuItem(mvDataRow, FinancialMenuItems.fmiGoToEvent) Then vShowItems(FinancialMenuItems.fmiGoToEvent) = True
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants, CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities, _
             CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits, CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
          If mvDataRow IsNot Nothing Then
            If vCanEdit Then
              vShowItems(FinancialMenuItems.fmiCancel) = mvDataRow.Item("CancellationReason").ToString.Length = 0
              If mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants Then
                vShowItems(FinancialMenuItems.fmiFutureCancel) = mvDataRow.Item("CancellationReason").ToString.Length = 0
                vShowItems(FinancialMenuItems.fmiReinstateAutoPayMethod) = mvDataRow.Item("CancellationReason").ToString.Length > 0
                vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
              End If
              vShowItems(FinancialMenuItems.fmiChangeCancel) = mvDataRow.Item("CancellationReason").ToString.Length > 0
            End If
            vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactSubscriptions
          If mvDataRow IsNot Nothing Then
            If vCanEdit Then
              vShowItems(FinancialMenuItems.fmiSubChangeCommunication) = mvDataRow.Item("CancellationReason").ToString.Length = 0
            End If
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations ', CareServices.XMLContactDataSelectionTypes.xcdtContactPledges, CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges
          If vCanEdit Then
            vShowItems(FinancialMenuItems.fmiNew) = vHasAccessRights
            'BR12029: don't allow New for Joint Contacts (Or Organisation)
            'BR12030: don't allow maintenance if contact's default address is membership_control.no_address_number
            'BR13326: don't allow New/Maintenance if contact's status is gone away
            Dim vCanMaintain As Boolean = (mvContactInfo.ContactType = ContactInfo.ContactTypes.ctContact) AndAlso _
                           (mvContactInfo.AddressNumber <> IntegerValue(AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.non_address_number))) AndAlso _
                           mvContactInfo.Status <> AppValues.ControlValue(AppValues.ControlValues.gone_away_status)
            Items(FinancialMenuItems.fmiNew).Enabled = vCanMaintain
            If mvDataRow Is Nothing Then
              vShowItems(FinancialMenuItems.fmiEdit) = False
              vShowItems(FinancialMenuItems.fmiCancel) = False
            Else
              Items(FinancialMenuItems.fmiEdit).Enabled = vCanMaintain AndAlso vHasAccessRights
              vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
              vShowItems(FinancialMenuItems.fmiCancel) = mvDataRow.Item("CancellationReason").ToString.Length = 0
            End If
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates
          vShowItems(FinancialMenuItems.fmiNew) = False
          If vCanEdit Then
            If mvDataRow Is Nothing Then
              vShowItems(FinancialMenuItems.fmiEdit) = False
              vShowItems(FinancialMenuItems.fmiCancel) = False
            Else
              vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
              If mvDataRow.Item("CancellationReason").ToString.Length > 0 Then
                vShowItems(FinancialMenuItems.fmiEdit) = False
                vShowItems(FinancialMenuItems.fmiCancel) = False
              ElseIf (IntegerValue(mvDataRow.Item("ClaimNumber")) > 0 And DoubleValue(mvDataRow.Item("AmountPaid").ToString) > 0) Then
                vShowItems(FinancialMenuItems.fmiCancel) = False
              Else
                vShowItems(FinancialMenuItems.fmiCancel) = True
              End If
            End If
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
          If mvDataRow IsNot Nothing Then
            Dim vDT As DataTable = DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipMenu, CInt(mvDataRow.Item("MembershipNumber"))).Tables("DataRow")
            SetMenuItems(vShowItems, vDT)
            vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
            If vShowItems(FinancialMenuItems.fmiReprintMembershipCard) = True Then
              Dim vMenuItem As ToolStripMenuItem = DirectCast(Me.Items(FinancialMenuItems.fmiReprintMembershipCard), ToolStripMenuItem)
              vMenuItem.Checked = BooleanValue(mvDataRow.Item("ReprintCard").ToString)
              vShowItems(FinancialMenuItems.fmiProduceMembershipCard) = True
            End If
            vShowItems(FinancialMenuItems.fmiFutureRenewalAmount) = (mvDataRow.Item("CancellationReason").ToString.Length = 0 AndAlso mvDataRow.Item("FutureMembershipType").ToString.Length > 0)
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
          If mvDataRow IsNot Nothing Then
            Dim vDT As DataTable
            If mvSubMenu Then
              vDT = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanMemberMenu, CInt(mvDataRow.Item("OrderNumber")), CInt(mvDataRow.Item("ContactNumber")), mvDataRow.Item("MemberNumber").ToString).Tables("DataRow")
              If mvDataRow.Item("CancelledOn").ToString.Length > 0 Then
                For Each vRow As DataRow In vDT.Rows
                  vRow.Item("MenuItemAvailable") = False
                Next
              End If
            Else
              vDT = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanMenu, CInt(mvDataRow.Item("PaymentPlanNumber"))).Tables("DataRow")
            End If
            SetMenuItems(vShowItems, vDT)
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
          CanPartRefund = False
          Dim vDT As DataTable = Nothing
          If mvMenuContextType = MenuContextTypes.mctFinder Then
            vDT = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtFinancialHistoryMenu, mvBatchNumber, mvTransactionNumber).Tables("DataRow")
          ElseIf mvMenuContextType = MenuContextTypes.mctAnalysis Then
            If mvDataRow IsNot Nothing Then
              vDT = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtFinancialHistoryDetailsMenu, IntegerValue(mvDataRow("BatchNumber")), IntegerValue(mvDataRow("TransactionNumber")), 0, IntegerValue(mvAnalysisRow("LineNumber"))).Tables("DataRow")
              Dim vPayPlanType As String = mvAnalysisRow("PaymentPlanType").ToString
              Select Case vPayPlanType
                Case "M"
                  vShowItems(FinancialMenuItems.fmiGoToMembership) = True
                  vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
                Case "C"
                  vShowItems(FinancialMenuItems.fmiGoToCovenant) = True
                  vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
                Case Else
                  If vPayPlanType.Length > 0 Then vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
              End Select
              'Turn these on now they may be turned off later
              Dim vEventGroups As Integer = DataHelper.EventGroups.Count
              vShowItems(FinancialMenuItems.fmiAddEventFinancialLink) = vEventGroups > 0
              vShowItems(FinancialMenuItems.fmiAddEventFinancialLink2) = vEventGroups > 1
              vShowItems(FinancialMenuItems.fmiAddEventFinancialLink3) = vEventGroups > 2
              vShowItems(FinancialMenuItems.fmiAddEventFinancialLink4) = vEventGroups > 3
              vShowItems(FinancialMenuItems.fmiAddEventFinancialLink5) = vEventGroups > 4
              mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
              mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
              mvTargetLineNumber = IntegerValue(mvAnalysisRow("LineNumber"))
            End If
          Else
            If mvDataRow IsNot Nothing Then
              vDT = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtFinancialHistoryMenu, CInt(mvDataRow.Item("BatchNumber")), CInt(mvDataRow.Item("TransactionNumber"))).Tables("DataRow")
              If IntegerValue(mvDataRow("BankDetailsNumber")) > 0 Then vShowItems(FinancialMenuItems.fmiGoToBankAccount) = Not mvMultiSelect
              vShowItems(FinancialMenuItems.fmiEditNotes) = Not mvMultiSelect
              vShowItems(FinancialMenuItems.fmiEditReference) = Not mvMultiSelect And vHasAccessRights
              If mvMultiSelect Then vShowItems(FinancialMenuItems.fmiPrintReceipt) = False    'Cannot print receipt if multiple lines selected
            End If
          End If
          SetMenuItems(vShowItems, vDT)
          If mvMenuContextType = MenuContextTypes.mctAnalysis Then
            For Each vRow As DataRow In vDT.Rows
              If vRow.Item(0).ToString.Equals("CanPartReverse", StringComparison.InvariantCultureIgnoreCase) Then
                If vShowItems(FinancialMenuItems.fmiReverse) = True AndAlso BooleanValue(vRow.Item(1).ToString) = True Then
                  CanPartRefund = True
                End If
              ElseIf vRow.Item(0).ToString.Equals("CanPartRefund", StringComparison.InvariantCultureIgnoreCase) Then
                If vShowItems(FinancialMenuItems.fmiRefund) = True AndAlso BooleanValue(vRow.Item(1).ToString) = True Then
                  CanPartRefund = True
                End If
              End If
            Next
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
          If mvDataRow IsNot Nothing Then
            mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
            mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
            mvTargetLineNumber = 1
            If mvSubMenu = False Then vShowItems(FinancialMenuItems.fmiChangeInvoiceAddress) = True
            If IntegerValue(mvDataRow("StoredInvoiceNumber").ToString) > 0 Then
              If mvSubMenu = False AndAlso (mvDataRow("TransactionType").ToString <> "Payment" AndAlso mvDataRow("TransactionType").ToString <> "Adjustment") Then vShowItems(FinancialMenuItems.fmiAmendDueDate) = True
              Dim vList As New ParameterList(True)
              vList("RemoveAllocations") = "N"
              Dim vReturnList As ParameterList = DataHelper.UpdateInvoice(CInt(mvDataRow("StoredInvoiceNumber")), vList)
              If vReturnList.IntegerValue("PaymentHistoryCount") > 0 Then vShowItems(FinancialMenuItems.fmiRemoveAllocations) = True
              If mvTargetBatchNumber > 0 AndAlso mvTargetTransactionNumber > 0 Then vShowItems(FinancialMenuItems.fmiGoToTransaction) = True
            End If
            If mvSubMenu = False AndAlso AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.preview_invoice_std_document).Length > 0 _
            AndAlso (mvDataRow("TransactionType").ToString = "Invoice" OrElse mvDataRow("TransactionType").ToString = "Credit Note") Then
              vShowItems(FinancialMenuItems.fmiPreviewInvoice) = True
            End If
          End If
          'Always show Display Transactions drop down menu & set check state (on Invoices grid only and not the Invoice Payments grid)
          If mvSubMenu = False Then
            Dim vDisplayTransactions As ToolStripItemCollection = DirectCast(Me.Items(FinancialMenuItems.fmiDisplayTransactions), ToolStripMenuItem).DropDownItems
            DirectCast(vDisplayTransactions(FinancialMenuItems.fmiDisplayTransactionsAll - FinancialMenuItems.fmiDisplayTransactionsAll), ToolStripMenuItem).Checked = DisplayTransactionsAllocationType = "A"
            DirectCast(vDisplayTransactions(FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly - FinancialMenuItems.fmiDisplayTransactionsAll), ToolStripMenuItem).Checked = DisplayTransactionsAllocationType = "F"
            DirectCast(vDisplayTransactions(FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly - FinancialMenuItems.fmiDisplayTransactionsAll), ToolStripMenuItem).Checked = DisplayTransactionsAllocationType = "U"
            vShowItems(FinancialMenuItems.fmiDisplayTransactions) = True
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactDespatchNotes
          If mvDataRow IsNot Nothing Then
            mvTargetBatchNumber = IntegerValue(mvDataRow.Item("BatchNumber").ToString)
            mvTargetTransactionNumber = IntegerValue(mvDataRow.Item("TransactionNumber").ToString)
            mvTargetLineNumber = 1
            If mvTargetBatchNumber > 0 AndAlso mvTargetTransactionNumber > 0 Then vShowItems(FinancialMenuItems.fmiGoToTransaction) = True
          End If

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions, _
         CareNetServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions, _
         CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions
          Dim vDT As DataTable = Nothing
          If mvDataRow IsNot Nothing Then
            If mvDataRow.Table.Columns.Contains("Provisional") AndAlso mvDataRow.Table.Columns.Contains("PaymentMethodCode") Then
              If BooleanValue(mvDataRow("Provisional").ToString) Then
                vShowItems(FinancialMenuItems.fmiConfirmTransaction) = True
                If mvDataRow("PaymentMethodCode").ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_caf_card) OrElse mvDataRow("PaymentMethodCode").ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_voucher) Then
                  vShowItems(FinancialMenuItems.fmiCancel) = True
                End If
              End If
            End If

            If mvDataRow IsNot Nothing Then
              'todo Simon needs to add xtdtFinancialTransactionMenu in CareService
              vDT = DataHelper.GetTransactionData(CType(CareNetServices.XMLTransactionDataSelectionTypes.xtdtFinancialTransactionMenu, CareServices.XMLTransactionDataSelectionTypes), CInt(mvDataRow.Item("BatchNumber")), CInt(mvDataRow.Item("TransactionNumber"))).Tables("DataRow")
              SetMenuItems(vShowItems, vDT)
              If mvDataRow.Table.Columns.Contains("Provisional") Then
                If BooleanValue(mvDataRow("Provisional").ToString) Then
                  vShowItems(FinancialMenuItems.fmiCancel) = True
                End If
              End If
            End If
            vShowItems(FinancialMenuItems.fmiEditUnprocessedTransactionNotes) = Not mvMultiSelect
            If mvMultiSelect = True OrElse mvDataType <> CareNetServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions Then
              'Cannot print receipt if selecting multiple lines or from Sales / Delivery Transactions grids
              vShowItems(FinancialMenuItems.fmiPrintReceipt) = False
            End If
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCancelledProvisionalTrans
          vShowItems(FinancialMenuItems.fmiReinstateProvisionalTrans) = mvDataRow IsNot Nothing
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPledges
          If vCanEdit Then
            If mvDataRow IsNot Nothing Then
              If mvDataRow("CancellationReason").ToString = "" Then
                vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
                vShowItems(FinancialMenuItems.fmiCancel) = True
              End If
            End If
            vShowItems(FinancialMenuItems.fmiNew) = True
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges
          If vCanEdit Then
            If mvDataRow IsNot Nothing Then
              If mvDataRow("CancellationReason").ToString = "" Then
                vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
                vShowItems(FinancialMenuItems.fmiCancel) = True
              End If
            End If
            vShowItems(FinancialMenuItems.fmiNew) = vHasAccessRights
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
          'Contact Fundraising Requests Menu          
          If vCanEdit Then
            Dim vCanMaintainActions As Boolean = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_actions)
            vShowItems(FinancialMenuItems.fmiNew) = vHasAccessRights
            If mvDataRow IsNot Nothing Then
              vShowItems(FinancialMenuItems.fmiNewAdHocAction) = vHasAccessRights AndAlso mvDataRow("LogName").ToString.Length > 0 AndAlso vCanMaintainActions
              vShowItems(FinancialMenuItems.fmiNewActionFromTemplate) = vHasAccessRights AndAlso mvDataRow("LogName").ToString.Length > 0 AndAlso vCanMaintainActions
              vShowItems(FinancialMenuItems.fmiGoToActions) = vHasAccessRights AndAlso mvDataRow("LogName").ToString.Length > 0 AndAlso BooleanValue(mvDataRow("HasAction").ToString) AndAlso vCanMaintainActions
              Dim vDefaultStatus As String = AppValues.ControlValue(AppValues.ControlValues.fundraising_status)
              Dim vShowUnlock As Boolean = BooleanValue(AppValues.ControlValue(AppValues.ControlValues.lock_fundraising_request)) AndAlso vDefaultStatus.Length > 0 _
                             AndAlso mvDataRow("FundraisingStatus").ToString.Length > 0 AndAlso mvDataRow("FundraisingStatus").ToString <> vDefaultStatus
              vShowItems(FinancialMenuItems.fmiUnlockFundraisingRequest) = vShowUnlock
              vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights AndAlso Not vShowUnlock
            End If
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
          If mvDataRow IsNot Nothing Then
            vShowItems(FinancialMenuItems.fmiGoToPayPlan) = True
            If vCanEdit = True AndAlso mvDataRow("CancellationReason").ToString.Length = 0 Then
              vShowItems(FinancialMenuItems.fmiCancel) = True
              vShowItems(FinancialMenuItems.fmiRecalcLoanInterest) = True
            End If
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
          If mvDataRow IsNot Nothing Then
            mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
            mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
            vShowItems(FinancialMenuItems.fmiGoToTransaction) = True
            If vCanEdit Then
              vShowItems(FinancialMenuItems.fmiCancelExamBooking) = (mvDataRow.Item("CancellationReason").ToString.Length = 0)
              vShowItems(FinancialMenuItems.fmiChangeExamCentre) = True
            End If
          End If

        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamExemptions
          If mvDataRow IsNot Nothing Then
            mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
            mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
            mvTargetLineNumber = IntegerValue(mvDataRow("LineNumber"))
            If mvTargetBatchNumber > 0 AndAlso mvTargetTransactionNumber > 0 Then vShowItems(FinancialMenuItems.fmiGoToTransaction) = True
          End If
          If vCanEdit Then
            vShowItems(FinancialMenuItems.fmiNew) = vHasAccessRights
            If mvDataRow Is Nothing Then
              vShowItems(FinancialMenuItems.fmiEdit) = False
            Else
              vShowItems(FinancialMenuItems.fmiEdit) = vHasAccessRights
            End If
          End If

      End Select
      Dim vVisibleCount As Integer
      Dim vShowItem As Boolean
      For vIndex As Integer = 0 To Me.Items.Count - 1
        Dim vItem As MenuToolbarCommand = DirectCast(Me.Items(vIndex).Tag, MenuToolbarCommand)
        vShowItem = vShowItems(vIndex) AndAlso vItem.HideItem = False
        Me.Items(vIndex).Visible = vShowItem
        If vShowItem Then vVisibleCount += 1
      Next
      If vVisibleCount = 0 Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub SetMenuItems(ByRef pShowItems() As Boolean, ByVal pDT As DataTable)
    Dim vIndex As FinancialMenuItems
    Dim vShowItem As Boolean
    Try
      Me.SuspendLayout()
      If pDT IsNot Nothing Then
        For Each vDR As DataRow In pDT.Rows
          vShowItem = mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite And Not mvReadOnly
          vIndex = FinancialMenuItems.fmiNone
          Select Case vDR.Item("MenuItemOption").ToString
            Case "CanChangeCancel"
              vIndex = FinancialMenuItems.fmiChangeCancel
            Case "CanCancel"
              vIndex = FinancialMenuItems.fmiCancel
            Case "CanCancelProvisional"
              vIndex = FinancialMenuItems.fmiCancel
              vShowItem = vShowItem AndAlso pShowItems(FinancialMenuItems.fmiCancel)  'fmiCancel only be set as visible when the transaction is Provisional
            Case "CanCMT"
              vIndex = FinancialMenuItems.fmiChangeMembershipType
            Case "CanReinstate"
              If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails Then
                vIndex = FinancialMenuItems.fmiReinstateMembership
              Else
                vIndex = FinancialMenuItems.fmiReinstatePaymentPlan
              End If
            Case "CanFMT"
              vIndex = FinancialMenuItems.fmiFutureMembershipType
            Case "CanFutureCancel"
              vIndex = FinancialMenuItems.fmiFutureCancel
            Case "CanConvert"
              vIndex = FinancialMenuItems.fmiPaymentPlanConversion
            Case "CanChangePayer"
              vIndex = FinancialMenuItems.fmiChangePayer
            Case "CanSkipPayment"
              vIndex = FinancialMenuItems.fmiSkipPayment
            Case "CanAdvanceRenewalDate"
              vIndex = FinancialMenuItems.fmiAdvanceRenewalDate
            Case "CanMaintain"
              If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails Then
                vIndex = FinancialMenuItems.fmiAmendMembership
                'use current user's access level as no access control for Amend Membership
                If vShowItem Then vShowItem = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
              Else
                vIndex = FinancialMenuItems.fmiPaymentPlanMaintenance
              End If
            Case "CanConfirmPaymentPlan"
              vIndex = FinancialMenuItems.fmiConfirmPaymentPlan
              vShowItem = Not mvReadOnly
            Case "CanMove"
              vIndex = FinancialMenuItems.fmiMove
              vShowItem = Not mvReadOnly
            Case "CanReverse"
              vIndex = FinancialMenuItems.fmiReverse
              vShowItem = Not mvReadOnly
            Case "CanRefund"
              vIndex = FinancialMenuItems.fmiRefund
              vShowItem = Not mvReadOnly
            Case "CanReverseInAdvance"
              vIndex = FinancialMenuItems.fmiReverseInAdvance
              vShowItem = Not mvReadOnly
            Case "CanRefundInAdvance"
              vIndex = FinancialMenuItems.fmiRefundInAdvance
              vShowItem = Not mvReadOnly
            Case "CanReanalyse"
              vIndex = FinancialMenuItems.fmiAnalysis
              vShowItem = Not mvReadOnly
            Case "CanGoToChanges"
              vIndex = FinancialMenuItems.fmiGoToChanges
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
              If BooleanValue(vDR.Item("MenuItemAvailable").ToString) Then
                mvAdjustmentWasBatchNumber = IntegerValue(vDR("AdjustmentBatchNumber"))
                mvAdjustmentWasTransactionNumber = IntegerValue(vDR("AdjustmentTransactionNumber"))
              End If
            Case "CanGoToChangedBy"
              vIndex = FinancialMenuItems.fmiGoToChangedBy
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
              If BooleanValue(vDR.Item("MenuItemAvailable").ToString) Then
                mvAdjustmentBatchNumber = IntegerValue(vDR("AdjustmentBatchNumber"))
                mvAdjustmentTransactionNumber = IntegerValue(vDR("AdjustmentTransactionNumber"))
              End If
            Case "ChangePayerChangePPDs"
              mvChangePPDs = BooleanValue(vDR.Item("MenuItemAvailable").ToString)
            Case "CanGoToBackOrders"
              vIndex = FinancialMenuItems.fmiGoToBackOrders
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
              mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
              mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
              mvTargetLineNumber = IntegerValue(mvAnalysisRow("LineNumber"))
            Case "CanGoToCC"
              vIndex = FinancialMenuItems.fmiGoToCC
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToCovenant"
              vIndex = FinancialMenuItems.fmiGoToCovenant
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToDD"
              vIndex = FinancialMenuItems.fmiGoToDD
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToDespatch"
              vIndex = FinancialMenuItems.fmiGoToDespatch
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
              mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
              mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
              mvTargetLineNumber = IntegerValue(mvAnalysisRow("LineNumber"))
            Case "CanGoToEvent"
              If mvMenuContextType <> MenuContextTypes.mctFinder Then
                Dim vEventNumber As Integer = IntegerValue(vDR("EventNumber"))
                If vEventNumber > 0 Then
                  Dim vEventInfo As CareEventInfo = New CareEventInfo(vEventNumber)
                  For vLoop As Integer = 0 To 4
                    If vLoop < DataHelper.EventGroups.Count Then
                      If DataHelper.EventGroups(vLoop).Code = vEventInfo.EventGroup Then
                        mvEventInfo(vLoop) = vEventInfo
                        mvEventNumber(vLoop) = vEventNumber
                        vIndex = CType(FinancialMenuItems.fmiGoToEvent + vLoop, FinancialMenuItems)
                        Me.Items(vIndex).Text = String.Format(ControlText.MnuFinancialGoToEvent, DataHelper.EventGroups(vLoop).GroupName)
                        vShowItem = True
                      End If
                    End If
                  Next
                Else
                  'If we did not find any Event Financial Links then EventNumber will be zero
                  vIndex = FinancialMenuItems.fmiGoToEvent
                  vShowItem = True
                  For vLoop As Integer = 1 To 4
                    If vLoop < DataHelper.EventGroups.Count Then
                      pShowItems(FinancialMenuItems.fmiGoToEvent + vLoop) = False
                    End If
                  Next
                End If
              End If
            Case "CanGoToLinks"
              vIndex = FinancialMenuItems.fmiGoToLinks
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
              mvLineType = vDR("LineType").ToString
              mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
              mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
              mvTargetLineNumber = IntegerValue(mvAnalysisRow("LineNumber"))
            Case "CanGoToMembership"
              vIndex = FinancialMenuItems.fmiGoToMembership
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToPreTaxPledge"
              vIndex = FinancialMenuItems.fmiGoToPreTaxPledge
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToPostTaxPledge"
              vIndex = FinancialMenuItems.fmiGoToPostTaxPledge
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToSO"
              vIndex = FinancialMenuItems.fmiGoToSO
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanAmendBooking"
              vIndex = FinancialMenuItems.fmiAmendBooking
              vShowItem = mvMenuContextType = MenuContextTypes.mctAnalysis AndAlso mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings
            Case "CanAddFundraisingPaymentLink"
              vIndex = FinancialMenuItems.fmiAddFundraisingPaymentLink
              If vShowItem Then vShowItem = mvMenuContextType = MenuContextTypes.mctAnalysis
              If vShowItem Then vShowItem = Not pShowItems(FinancialMenuItems.fmiGoToPayPlan)
            Case "CanReprintCard"
              vIndex = FinancialMenuItems.fmiReprintMembershipCard
            Case "CanGiftAidDeclaration"
              vIndex = FinancialMenuItems.fmiAddGiftAidDeclaration
            Case "CanAddNewMember"
              vIndex = FinancialMenuItems.fmiAddMember
              vShowItem = Not mvReadOnly
            Case "CanReplaceMember"
              vIndex = FinancialMenuItems.fmiReplaceMember
              vShowItem = Not mvReadOnly
            Case "CanAddGiftAidDecleration"
              vIndex = FinancialMenuItems.fmiAddGiftAidDeclaration
              vShowItem = Not mvReadOnly
            Case "CanPrint"
              vIndex = FinancialMenuItems.fmiPaymentPlanPrint
              vShowItem = True
            Case "CanGoToLoan"
              vIndex = FinancialMenuItems.fmiGoToLoan
              vShowItem = mvMenuContextType <> MenuContextTypes.mctFinder
            Case "CanGoToEventLinks"    'This item is telling you if there is an existing financial link and if so to which event
              If mvMenuContextType <> MenuContextTypes.mctFinder Then
                Dim vEventNumber As Integer = IntegerValue(vDR("EventNumber"))
                Dim vItemVisible As Boolean = BooleanValue(vDR.Item("MenuItemAvailable").ToString)
                If vEventNumber > 0 Then
                  Dim vEventInfo As CareEventInfo = New CareEventInfo(vEventNumber)
                  For vLoop As Integer = 0 To 4
                    If vLoop < DataHelper.EventGroups.Count Then
                      If DataHelper.EventGroups(vLoop).Code.Equals(vEventInfo.EventGroup, StringComparison.InvariantCultureIgnoreCase) Then
                        vShowItem = True
                        vIndex = CType(FinancialMenuItems.fmiAddEventFinancialLink + vLoop, FinancialMenuItems)
                        pShowItems(FinancialMenuItems.fmiRemoveEventFinancialLink + vLoop) = Not (vItemVisible)  'We have either Add or Remove menu
                      Else
                        pShowItems(FinancialMenuItems.fmiAddEventFinancialLink + vLoop) = vItemVisible
                        pShowItems(FinancialMenuItems.fmiRemoveEventFinancialLink + vLoop) = False
                      End If
                    End If
                  Next
                Else
                  'If we did not find any Event Financial Links then EventNumber will be zero
                  vShowItem = True
                  vIndex = FinancialMenuItems.fmiAddEventFinancialLink
                  pShowItems(FinancialMenuItems.fmiRemoveEventFinancialLink) = False
                  For vLoop As Integer = 1 To 4
                    If vLoop < DataHelper.EventGroups.Count Then
                      pShowItems(FinancialMenuItems.fmiAddEventFinancialLink + vLoop) = vItemVisible
                      pShowItems(FinancialMenuItems.fmiRemoveEventFinancialLink + vLoop) = False
                    End If
                  Next
                End If
              End If
            Case "CanPrintReceipt"
              vIndex = FinancialMenuItems.fmiPrintReceipt
              vShowItem = (mvMenuContextType = MenuContextTypes.mctMain)
          End Select
          If vIndex > FinancialMenuItems.fmiNone Then pShowItems(vIndex) = BooleanValue(vDR.Item("MenuItemAvailable").ToString) And vShowItem And (mvMultiSelect = False OrElse vIndex = FinancialMenuItems.fmiAnalysis)
        Next
      End If
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

  Protected Function SetGotoEventMenuItem(ByVal pRow As DataRow, ByVal pIndex As FinancialMenuItems) As Boolean
    Dim vEventNumber As Integer = IntegerValue(pRow("EventNumber"))
    If vEventNumber > 0 Then
      Dim vEventInfo As CareEventInfo = New CareEventInfo(vEventNumber)
      mvEventInfo(pIndex - FinancialMenuItems.fmiGoToEvent) = vEventInfo
      Dim vEventType As String = "Event"
      If DataHelper.EventGroups.ContainsKey(vEventInfo.EventGroup) Then vEventType = DataHelper.EventGroups(vEventInfo.EventGroup).GroupName
      Me.Items(pIndex).Text = String.Format(ControlText.MnuFinancialGoToEvent, vEventType)
      mvEventNumber(pIndex - FinancialMenuItems.fmiGoToEvent) = vEventNumber
      Return True
    End If
  End Function

  Public ReadOnly Property AdjustmentBatchNumber() As Integer
    Get
      Return mvAdjustmentBatchNumber
    End Get
  End Property
  Public ReadOnly Property AdjustmentTransactionNumber() As Integer
    Get
      Return mvAdjustmentTransactionNumber
    End Get
  End Property
  Public ReadOnly Property AdjustmentWasBatchNumber() As Integer
    Get
      Return mvAdjustmentWasBatchNumber
    End Get
  End Property
  Public ReadOnly Property AdjustmentWasTransactionNumber() As Integer
    Get
      Return mvAdjustmentWasTransactionNumber
    End Get
  End Property
  Public ReadOnly Property TargetBatchNumber() As Integer
    Get
      Return mvTargetBatchNumber
    End Get
  End Property
  Public ReadOnly Property TargetTransactionNumber() As Integer
    Get
      Return mvTargetTransactionNumber
    End Get
  End Property
  Public ReadOnly Property TargetLineNumber() As Integer
    Get
      Return mvTargetLineNumber
    End Get
  End Property
  Public ReadOnly Property LineType() As String
    Get
      Return mvLineType
    End Get
  End Property
  Public ReadOnly Property EventNumber(ByVal pIndex As Integer) As Integer
    Get
      If pIndex < 0 OrElse pIndex > MAX_EVENT_GROUP_INDEX Then Throw New ArgumentOutOfRangeException
      Return mvEventNumber(pIndex)
    End Get
  End Property
  Public Property CanPartRefund() As Boolean
    Get
      Return mvCanPartRefund
    End Get
    Private Set(value As Boolean)
      mvCanPartRefund = value
    End Set
  End Property

  Public Property DisplayTransactionsAllocationType() As String
    Get
      If mvDisplayTransactionsAllocationType Is Nothing Then
        mvDisplayTransactionsAllocationType = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_sl_default_invoice_display, "A")
      End If
      Return mvDisplayTransactionsAllocationType
    End Get
    Set(pValue As String)
      mvDisplayTransactionsAllocationType = pValue
    End Set
  End Property

End Class

Public Class FinancialMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow, ByVal pContactInfo As ContactInfo, ByVal pReadOnly As Boolean)
    SetContext(pDST, pDataRow, pContactInfo, pReadOnly, False)
  End Sub
  Public Sub SetContext(ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow, ByVal pContactInfo As ContactInfo, ByVal pReadOnly As Boolean, ByVal pMultiSelect As Boolean)
    mvDataType = pDST
    mvDataRow = pDataRow
    mvContactInfo = pContactInfo
    mvReadOnly = pReadOnly
    mvMultiSelect = pMultiSelect
  End Sub

  Public Overloads Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    MyBase.SetVisibleItems(e)
  End Sub

  Public ReadOnly Property DataRow() As DataRow
    Get
      Return mvDataRow
    End Get
  End Property

End Class

Public Class FinancialFinderMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent, CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, Nothing)
    mvMenuContextType = MenuContextTypes.mctFinder
  End Sub

  Public Sub SetContext(ByVal pContactNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
    mvContactInfo = New ContactInfo(pContactNumber)
    mvBatchNumber = pBatchNumber
    mvTransactionNumber = pTransactionNumber
  End Sub

End Class

Public Class FinancialAnalysisMenu
  Inherits BaseFinancialMenu

  Public Shadows Event MenuSelected(ByVal pItem As FinancialMenuItems, ByVal pDataRow As DataRow)

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
    mvMenuContextType = MenuContextTypes.mctAnalysis
  End Sub

  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow, ByVal pAnalysisRow As DataRow)
    mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
    mvAnalysisRow = pAnalysisRow
  End Sub

  Public ReadOnly Property CareEventInfo(ByVal pIndex As Integer) As CareEventInfo
    Get
      If pIndex < 0 OrElse pIndex > 4 Then Throw New ArgumentOutOfRangeException
      Return mvEventInfo(pIndex)
    End Get
  End Property

  Protected Overrides Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FinancialMenuItems)
    Dim vCursor As New BusyCursor
    Try
      RaiseEvent MenuSelected(pItem, mvAnalysisRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class

Public Class TransactionLinkMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub
  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      If mvDataRow IsNot Nothing Then
        If mvDataRow.Table.Columns.Contains("ClaimDate") AndAlso
          Not IsDBNull(mvDataRow("ClaimDate")) AndAlso
           Not String.IsNullOrWhiteSpace(CStr(mvDataRow("ClaimDate"))) Then
          Me.Items(FinancialMenuItems.fmiChangeClaimDate).Visible = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciChangeClaimDate)
          Me.Items(FinancialMenuItems.fmiChangeClaimDate).Enabled = (CStr(mvDataRow("ScheduledPaymentStatus")).Equals("D", StringComparison.InvariantCultureIgnoreCase) OrElse
                                                                     CStr(mvDataRow("ScheduledPaymentStatus")).Equals("V", StringComparison.InvariantCultureIgnoreCase)) AndAlso
                                                                Not (CStr(mvDataRow("ScheduleCreationReason")).Equals("FA", StringComparison.InvariantCultureIgnoreCase) AndAlso CDbl(mvDataRow("Amount")) < 0)
        End If
        If mvDataRow.Table.Columns.Contains("ClaimDate") AndAlso Not IsDBNull(mvDataRow("ClaimDate")) AndAlso
          CStr(mvDataRow("ScheduleCreationReason")).Equals("FA", StringComparison.InvariantCultureIgnoreCase) AndAlso CDbl(mvDataRow("Amount")) > 0 Then
          Dim vDataTable As DataTable = mvDataRow.Table
          Dim vReverseAmount As Double = CDbl(mvDataRow("Amount")) * -1
          Dim vDataRow As DataRow = vDataTable.Select(String.Format("ScheduleCreationReason='FA' AND Convert(Amount,System.Double)={0} AND ClaimDate='{1}'", vReverseAmount.ToString(), CStr(mvDataRow("ClaimDate")))).FirstOrDefault()
          If vDataRow Is Nothing Then
            Me.Items(FinancialMenuItems.fmiChangeClaimDate).Enabled = False
          End If
        End If
        mvTargetBatchNumber = IntegerValue(mvDataRow("BatchNumber"))
        mvTargetTransactionNumber = IntegerValue(mvDataRow("TransactionNumber"))
        If mvDataRow.Table.Columns.Contains("LineNumber") Then
          mvTargetLineNumber = IntegerValue(mvDataRow("LineNumber"))
        Else
          mvTargetLineNumber = 0
        End If
        If mvTargetBatchNumber > 0 AndAlso mvTargetTransactionNumber > 0 Then
          Me.Items(FinancialMenuItems.fmiGoToTransaction).Visible = True
          e.Cancel = False
        Else
          e.Cancel = True
        End If
      Else
        e.Cancel = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

End Class

Public Class JournalLinkMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub
  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      Dim vCancel As Boolean = True
      If mvDataRow IsNot Nothing Then
        Select Case mvDataRow("JournalType").ToString
          Case "ACT"                              'Activities
          Case "ACTA", "ACTM", "ACTN", "ACTR"     'Action
          Case "ADD"                              'Address
          Case "CCCA", "CCM"                      'Continous Credit Card Authority
            Me.Items(FinancialMenuItems.fmiGoToCC).Visible = True
            vCancel = False
          Case "CMAD"                             'Mailing Document
          Case "CONT"                             'Contact
          Case "COV", "COVM"                      'Covenant
            Me.Items(FinancialMenuItems.fmiGoToCovenant).Visible = True
            vCancel = False
          Case "CPDC"                             'CPD Cycles
          Case "CPDP"                             'CPD Points
          Case "DD", "DDM"                        'Direct Debit
            Me.Items(FinancialMenuItems.fmiGoToDD).Visible = True
            vCancel = False
          Case "DOC"                              'Communication
          Case "EVNT"                             'Event Booking
          Case "GAD"                              'Gift Aid Declaration
          Case "GAWA"                             'Gone Away
          Case "MAIL"                             'Mailing Sent
          Case "MEM", "MEMM"                      'Membership
            Me.Items(FinancialMenuItems.fmiGoToMembership).Visible = True
            vCancel = False
          Case "PLDG"                             'Pledge
          Case "PP", "PPM", "PPPS"                'Payment Plan
            Me.Items(FinancialMenuItems.fmiGoToPayPlan).Visible = True
            vCancel = False
          Case "SO", "SOM"                        'Standing Order
            Me.Items(FinancialMenuItems.fmiGoToSO).Visible = True
            vCancel = False
          Case "SUP"                              'Suppressions
          Case "NUMB"                             'Numbers
        End Select
      End If
      e.Cancel = vCancel
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
    End Try
  End Sub
End Class

Public Class PurchaseOrderMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub

  Public Overloads Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Dim vCursor As New BusyCursor
    Dim vCanEdit As Boolean
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      Try
        e.Cancel = False
        vCanEdit = mvContactInfo IsNot Nothing AndAlso mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite AndAlso Not mvReadOnly
        Dim vHasAccessRights As Boolean = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
        Dim vShowItems(Me.Items.Count) As Boolean

        If mvDataRow IsNot Nothing Then
          Dim vShow As Boolean
          If mvDataRow("CancellationReason").ToString.Length = 0 Then vShow = True
          vShowItems(FinancialMenuItems.fmiAmendPurchaseOrder) = vShow AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_pom).Length > 0 AndAlso
                        (BooleanValue(mvDataRow("HasInvoice").ToString) = False OrElse ((BooleanValue(mvDataRow("PaymentSchedule").ToString) = True OrElse
                        BooleanValue(mvDataRow("AdHocPayments").ToString) = True OrElse BooleanValue(mvDataRow("RegularPayments").ToString) = True)))
          vShowItems(FinancialMenuItems.fmiCancel) = vShow
          vShowItems(FinancialMenuItems.fmiReinstate) = Not vShow
          vShow = False
          If mvDataRow.Table.Columns.Contains("RequiresAuthorisation") Then
            If mvDataRow("RequiresAuthorisation").ToString.StartsWith("Y") AndAlso mvDataRow("CanAuthorise").ToString.StartsWith("Y") AndAlso mvDataRow("AuthorisedBy").ToString.Length = 0 Then
              vShow = True
            End If
          End If
          vShowItems(FinancialMenuItems.fmiAuthorisePurchaseOrder) = vShow
        End If

        'Check Access Control for Amend Purchase Order
        If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciPurchaseOrderAmend) AndAlso vShowItems(FinancialMenuItems.fmiAmendPurchaseOrder) Then
          vShowItems(FinancialMenuItems.fmiAmendPurchaseOrder) = True
        Else
          vShowItems(FinancialMenuItems.fmiAmendPurchaseOrder) = False
        End If

        'Check Access Control for Cancel Purchase Order
        If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciPurchaseOrderCancel) AndAlso vShowItems(FinancialMenuItems.fmiCancel) Then
          vShowItems(FinancialMenuItems.fmiCancel) = True
        Else
          vShowItems(FinancialMenuItems.fmiCancel) = False
        End If

        'Check Access Control for Reinstate Purchase Order
        If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciPurchaseOrderReinstate) AndAlso vShowItems(FinancialMenuItems.fmiReinstate) Then
          vShowItems(FinancialMenuItems.fmiReinstate) = True
        Else
          vShowItems(FinancialMenuItems.fmiReinstate) = False
        End If

        Dim vVisibleCount As Integer
        Dim vItem As MenuToolbarCommand
        Dim vShowItem As Boolean
        For vIndex As Integer = 0 To Me.Items.Count - 1
          vItem = DirectCast(Me.Items(vIndex).Tag, MenuToolbarCommand)
          vShowItem = vShowItems(vIndex) AndAlso vItem.HideItem = False
          Me.Items(vIndex).Visible = vShowItem
          If vShowItem Then vVisibleCount += 1
        Next
        If vVisibleCount = 0 Then e.Cancel = True
      Catch vException As Exception
        DataHelper.HandleException(vException)
      Finally
        vCursor.Dispose()
      End Try
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Try
      SetVisibleItems(e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class

Public Class PurchaseOrderPaymentMenu
  Inherits BaseFinancialMenu

  Private mvParentDataRow As DataRow

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub

  Public Sub SetParentContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvParentDataRow = pDataRow
  End Sub
  Public Overloads Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Dim vCursor As New BusyCursor
    Dim vCanEdit As Boolean
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      Try
        e.Cancel = False
        vCanEdit = mvContactInfo IsNot Nothing AndAlso mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite AndAlso Not mvReadOnly
        Dim vHasAccessRights As Boolean = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
        Dim vShowItems(Me.Items.Count) As Boolean
        If mvParentDataRow IsNot Nothing Then
          If mvParentDataRow.Table.Columns.Contains("AdHocPayments") Then
            If mvParentDataRow("AdHocPayments").ToString.StartsWith("Y") AndAlso _
               mvParentDataRow("CancellationReason").ToString.Length = 0 Then vShowItems(FinancialMenuItems.fmiNew) = True
          End If
        End If
        If mvDataRow IsNot Nothing Then
          Dim vIsReceipt As Boolean
          If mvDataRow.Table.Columns.Contains("PaymentNumber") Then
            If mvDataRow("AuthorisedBy").ToString.Length = 0 AndAlso mvDataRow("PostedOn").ToString.Length = 0 Then vShowItems(FinancialMenuItems.fmiAuthorise) = True
            If IntegerValue(mvDataRow("ReceiptForPaymentNumber").ToString) > 0 Then vIsReceipt = True
            'Check parent purchase order
            If (mvParentDataRow("RequiresAuthorisation").ToString.StartsWith("Y") AndAlso _
                mvParentDataRow("AuthorisedBy").ToString.Length = 0) OrElse _
               mvDataRow("NoPaymentRequired").ToString.StartsWith("Y") Then vShowItems(FinancialMenuItems.fmiAuthorise) = False
          End If
          'Can only edit if an uncancelled ad-hoc payment type purchase order
          vShowItems(FinancialMenuItems.fmiEdit) = vShowItems(FinancialMenuItems.fmiNew)
          'Can only add a receipt if an uncancelled ad-hoc payment type purchase order and not already a receipt
          vShowItems(FinancialMenuItems.fmiAddPurchaseOrderPaymentReceipt) = vShowItems(FinancialMenuItems.fmiNew) And Not vIsReceipt

          'BR17340 
          If mvDataRow.Table.Columns.Contains("ChequeProducedOn") Then 'cheque produced on has a value and not already reversed
            If mvDataRow("ChequeProducedOn").ToString.Length > 0 And mvDataRow("AdjustmentStatus").ToString.Length = 0 Then
              vShowItems(FinancialMenuItems.fmiCancelPOP) = True
              If DoubleValue(mvDataRow("Amount").ToString) > 0 Then vShowItems(FinancialMenuItems.fmiPOPAnalysis) = True
            End If
          End If

          'set go to changed by menu
          Dim mvParams As New ParameterList(True)
          mvParams("PurchaseOrderNumber") = mvDataRow("PurchaseOrderNumber").ToString()
          mvParams("PaymentNumber") = mvDataRow("PaymentNumber").ToString()

          Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)

          'no menus by default
          vShowItems(FinancialMenuItems.fmiGoToPopRevChangedBy) = False
          vShowItems(FinancialMenuItems.fmiGoToPopRevGoToChanges) = False

          If vResults("MenuSelection") = "2" Then 'no menus
          ElseIf vResults("MenuSelection") = "1" Then
            vShowItems(FinancialMenuItems.fmiGoToPopRevGoToChanges) = True
            vShowItems(FinancialMenuItems.fmiCancelPOP) = False
          ElseIf vResults("MenuSelection") = "0" Then
            vShowItems(FinancialMenuItems.fmiGoToPopRevChangedBy) = True
            vShowItems(FinancialMenuItems.fmiCancelPOP) = False
          End If
        End If
        Dim vVisibleCount As Integer
        Dim vItem As MenuToolbarCommand
        Dim vShowItem As Boolean
        For vIndex As Integer = 0 To Me.Items.Count - 1
          vItem = DirectCast(Me.Items(vIndex).Tag, MenuToolbarCommand)
          vShowItem = vShowItems(vIndex) AndAlso vItem.HideItem = False
          Me.Items(vIndex).Visible = vShowItem
          If vShowItem Then vVisibleCount += 1
        Next
        If vVisibleCount = 0 Then e.Cancel = True
      Catch vException As Exception
        DataHelper.HandleException(vException)
      Finally
        vCursor.Dispose()
      End Try
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Try
      SetVisibleItems(e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class

Public Class ContactEventDelegateMenu
  Inherits BaseFinancialMenu

  Private mvEventDelegateInfo As EventDelegateInfo

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub

  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      e.Cancel = False
      For vIndex As Integer = 0 To MAX_EVENT_GROUP_INDEX
        mvEventInfo(vIndex) = Nothing
      Next
      Dim vShow As Boolean
      If mvDataRow IsNot Nothing Then
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates
            If SetGotoEventMenuItem(mvDataRow, FinancialMenuItems.fmiGoToEvent) Then
              Me.Items(FinancialMenuItems.fmiGoToEvent).Visible = True
              If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciEventDelegateSupplementaryInformation) And (mvEventInfo(0).ActivityGroup.Length > 0 OrElse mvEventInfo(0).RelationshipGroup.Length > 0) Then vShow = True
              Me.Items(FinancialMenuItems.fmiSupplementaryInformation).Visible = vShow
              Dim vDelegateNumber As Integer = IntegerValue(mvDataRow("DelegateNumber"))
              mvEventDelegateInfo = New EventDelegateInfo(vDelegateNumber, mvContactInfo.ContactNumber, mvContactInfo.ContactName, IntegerValue(mvDataRow("BatchNumber")) > 0, mvDataRow("TransactionSource").ToString)
            End If
        End Select
      End If
      If Not vShow Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public ReadOnly Property CareEventInfo(ByVal pIndex As Integer) As CareEventInfo
    Get
      If pIndex < 0 OrElse pIndex > 4 Then Throw New ArgumentOutOfRangeException
      Return mvEventInfo(pIndex)
    End Get
  End Property

  Public ReadOnly Property EventDelegateInfo() As EventDelegateInfo
    Get
      Return mvEventDelegateInfo
    End Get
  End Property

End Class

Public Class EventDelegateMenu
  Inherits BaseFinancialMenu

  Private mvCareEventInfo As CareEventInfo

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent, CareServices.XMLContactDataSelectionTypes.xcdtNone, Nothing)
  End Sub

  Public Sub SetContext(ByVal pEventInfo As CareEventInfo)
    mvCareEventInfo = pEventInfo
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      e.Cancel = False
      Dim vShow As Boolean
      If mvCareEventInfo IsNot Nothing Then
        If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciEventDelegateSupplementaryInformation) And (mvCareEventInfo.ActivityGroup.Length > 0 OrElse mvCareEventInfo.RelationshipGroup.Length > 0) Then vShow = True
      End If
      Me.Items(FinancialMenuItems.fmiSupplementaryInformation).Visible = vShow
      If Not vShow Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class

Public Class EventFinancialLinkMenu
  Inherits BaseFinancialMenu

  Public Shadows Event MenuSelected(ByVal pItem As FinancialMenuItems)

  Private mvEventDataType As CareServices.XMLEventDataSelectionTypes = CType(-1, CareServices.XMLEventDataSelectionTypes)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent, CareServices.XMLContactDataSelectionTypes.xcdtNone, Nothing)
  End Sub

  Public Sub SetContext(ByVal pEventInfo As CareEventInfo, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer)
    mvEventInfo(0) = pEventInfo
    mvTargetBatchNumber = pBatchNumber
    mvTargetTransactionNumber = pTransactionNumber
    mvTargetLineNumber = pLineNumber
  End Sub
  Public Sub SetContext(ByVal pEventInfo As CareEventInfo, ByVal pEventDataType As CareServices.XMLEventDataSelectionTypes, ByVal pRow As DataRow)
    mvEventInfo(0) = pEventInfo
    mvEventDataType = pEventDataType
    mvDataRow = pRow
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      e.Cancel = False
      Dim vShow As Boolean
      If mvTargetBatchNumber > 0 Then
        Me.Items(FinancialMenuItems.fmiGoToTransaction).Visible = True
        For vLoop As Integer = 0 To 4
          If vLoop < DataHelper.EventGroups.Count Then
            If DataHelper.EventGroups(vLoop).Code = mvEventInfo(0).EventGroup Then
              Me.Items(FinancialMenuItems.fmiRemoveEventFinancialLink + vLoop).Visible = True
            End If
          End If
        Next
        vShow = True
      ElseIf mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then
        If mvDataRow IsNot Nothing Then
          If (mvDataRow.Item("BookingStatusCode").ToString <> "C" AndAlso mvDataRow.Item("BookingStatusCode").ToString <> "U") Then
            Dim vList As New ParameterList(True, False) 'Do not want SystemColumns
            vList("SmartClient") = "Y"
            Dim vBatchNumber As Integer = IntegerValue(mvDataRow.Item("BatchNumber").ToString)
            Dim vTransactionNumber As Integer = IntegerValue(mvDataRow.Item("TransactionNumber").ToString)
            Dim vRow As DataRow = Nothing
            If vBatchNumber > 0 AndAlso vTransactionNumber > 0 Then vRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vBatchNumber, vTransactionNumber, vList))
            If vRow IsNot Nothing Then
              Me.Items(FinancialMenuItems.fmiAmendBooking).Visible = (BooleanValue(vRow.Item("PostedToNominal").ToString) = True)
              Me.Items(FinancialMenuItems.fmiAnalysis).Visible = (BooleanValue(vRow.Item("PostedToNominal").ToString) = True)
            End If
          End If
        End If
      End If
      If Not vShow Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
      vCursor.Dispose()
    End Try
  End Sub

  Protected Overrides Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FinancialMenuItems)
    Dim vCursor As New BusyCursor
    Try
      RaiseEvent MenuSelected(pItem)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class

Public Class MailingDocumentMenu
  Inherits BaseFinancialMenu
  Private mvMailingNumber As Integer

  Private mvFinderType As CareServices.XMLDataFinderTypes = CareNetServices.XMLDataFinderTypes.xdftNone

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent, CareServices.XMLContactDataSelectionTypes.xcdtNone, Nothing)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub
  Public Sub SetContext(ByVal pFulfillmentNumber As Integer)
    mvNumber = pFulfillmentNumber
  End Sub
  Public Sub SetContext(pFinderType As CareServices.XMLDataFinderTypes, ByVal pEMailJobNumber As Integer, pMailingNumber As Integer)
    mvFinderType = pFinderType
    mvNumber = pEMailJobNumber
    mvMailingNumber = pMailingNumber
  End Sub

  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      Me.SuspendLayout()
      e.Cancel = False
      Dim vShow As Boolean

      If mvFinderType = CareNetServices.XMLDataFinderTypes.xdftMailings Then
        If mvNumber > 0 Then
          vShow = True
        ElseIf mvMailingNumber > 0 Then
          Dim vParamList As New ParameterList(True)
          vParamList("MailingNumber") = mvMailingNumber.ToString
          vShow = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMailingHistoryDocuments, vParamList) > 0
        End If
        Me.Items(FinancialMenuItems.fmiViewMailingDocument).Visible = vShow
      Else
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactMailings
            Me.Items(FinancialMenuItems.fmiViewMailingDocument).Visible = False
            Me.Items(FinancialMenuItems.fmiDeleteMailingDocument).Visible = False
            Me.Items(FinancialMenuItems.fmiUnfulfillMailingDocument).Visible = False
            If mvDataRow IsNot Nothing Then
              If mvDataRow.Table.Columns.Contains("Mailing") Then
                If mvDataRow("Mailing").ToString.Length > 0 Then vShow = True
              End If
              Dim vItem As MenuToolbarCommand
              If mvDataRow.Table.Columns.Contains("Type") AndAlso mvDataRow("Type").ToString.Length > 0 Then
                If mvDataRow("Type").ToString <> "Pending" Then
                  Me.Items(FinancialMenuItems.fmiViewMailingDocument).Visible = vShow
                  If mvDataRow.Table.Columns.Contains("FulfillmentNumber") AndAlso mvDataRow("FulfillmentNumber").ToString.Length > 0 _
                  AndAlso IntegerValue(mvDataRow("FulfillmentNumber").ToString) > 0 Then
                    vItem = DirectCast(Me.Items(FinancialMenuItems.fmiUnfulfillMailingDocument).Tag, MenuToolbarCommand)
                    Me.Items(FinancialMenuItems.fmiUnfulfillMailingDocument).Visible = vShow AndAlso Not vItem.HideItem
                  End If
                ElseIf mvDataRow("Type").ToString = "Pending" Then
                  vItem = DirectCast(Me.Items(FinancialMenuItems.fmiDeleteMailingDocument).Tag, MenuToolbarCommand)
                  Me.Items(FinancialMenuItems.fmiDeleteMailingDocument).Visible = vShow AndAlso Not vItem.HideItem
                End If
              End If
            End If

          Case CareServices.XMLContactDataSelectionTypes.xcdtNone  ' Currently used for RedoFulfilment only 
            If mvNumber > 0 Then vShow = True
            Me.Items(FinancialMenuItems.fmiRedoFulfilment).Visible = vShow
        End Select
      End If
      If Not vShow Then e.Cancel = True

    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
      vCursor.Dispose()
    End Try
  End Sub
End Class

Public Class PurchaseInvoiceChequeMenu
  Inherits BaseFinancialMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub
  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      Me.SuspendLayout()
      e.Cancel = False
      Dim vShow As Boolean
      Dim vItem As MenuToolbarCommand
      If mvDataRow IsNot Nothing Then
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseInvoices
            Me.Items(FinancialMenuItems.fmiReissueCheque).Visible = False
            Me.Items(FinancialMenuItems.fmiChangeChequePayee).Visible = False
            If mvDataRow.Table.Columns.Contains("ChequeReferenceNumber") Then
              If mvDataRow("ChequeReferenceNumber").ToString.Length > 0 AndAlso mvDataRow("ReconciledOn").ToString.Length = 0 Then
                If mvDataRow("ChequeNumber").ToString.Length > 0 AndAlso mvDataRow("PrintedOn").ToString.Length > 0 Then
                  vItem = DirectCast(Me.Items(FinancialMenuItems.fmiReissueCheque).Tag, MenuToolbarCommand)
                  vShow = Not vItem.HideItem
                  'If config set then don't allow reissue unless the status says you can
                  If vShow AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_po_cheque_reissue) AndAlso _
                    mvDataRow("AllowReissue").ToString.Length = 0 Then vShow = False
                  Me.Items(FinancialMenuItems.fmiReissueCheque).Visible = vShow
                  'Let them change the status if they have access rights
                  Me.Items(FinancialMenuItems.fmiChequeSetStatus).Visible = Not vItem.HideItem
                ElseIf mvDataRow("ChequeNumber").ToString.Length = 0 AndAlso mvDataRow("PrintedOn").ToString.Length = 0 Then
                  vItem = DirectCast(Me.Items(FinancialMenuItems.fmiChangeChequePayee).Tag, MenuToolbarCommand)
                  vShow = Not vItem.HideItem
                  Me.Items(FinancialMenuItems.fmiChangeChequePayee).Visible = vShow
                End If
              End If
              'BR17340
              'set go to changed by menu
              Dim mvParams As New ParameterList(True)
              mvParams("ChequeReferenceNumber") = mvDataRow("ChequeReferenceNumber").ToString()
              Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)

              'no menus by default
              Me.Items(FinancialMenuItems.fmiGoToPopRevChangedBy).Visible = False
              Me.Items(FinancialMenuItems.fmiGoToPopRevGoToChanges).Visible = False

              If vResults("MenuSelection") = "2" Then 'no menus
              ElseIf vResults("MenuSelection") = "1" Then
                Me.Items(FinancialMenuItems.fmiGoToPopRevGoToChanges).Visible = True
              ElseIf vResults("MenuSelection") = "0" Then
                Me.Items(FinancialMenuItems.fmiGoToPopRevChangedBy).Visible = True
              End If

            End If
        End Select
      End If
      If Not vShow Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
      vCursor.Dispose()
    End Try
  End Sub
End Class

Public Class FundraisingPaymentMenu
  Inherits BaseFinancialMenu
  Private mvPaymentRow As DataRow

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow, ByVal pPaymentRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
    mvPaymentRow = pPaymentRow
  End Sub
  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      SetVisibleItems(e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Overloads Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Try
      Me.SuspendLayout()
      e.Cancel = False
      Dim vShow As Boolean
      If mvDataRow IsNot Nothing Then
        Dim vDefaultStatus As String = AppValues.ControlValue(AppValues.ControlValues.fundraising_status)
        Dim vCanMaintain As Boolean = (DoubleValue(mvDataRow("PledgedAmount").ToString) > 0 OrElse DoubleValue(mvDataRow("ExpectedAmount").ToString) > 0 _
                                      OrElse DoubleValue(mvDataRow("GikPledgedAmount").ToString) > 0 OrElse DoubleValue(mvDataRow("GikExpectedAmount").ToString) > 0) _
                                      AndAlso vDefaultStatus.Length > 0 AndAlso vDefaultStatus = mvDataRow("FundraisingStatus").ToString
        Me.Items(FinancialMenuItems.fmiNew).Visible = vCanMaintain
        vShow = vCanMaintain
        Dim vHasAccessRights As Boolean
        Dim vHasOwner As Boolean
        Dim vCanMaintainActions As Boolean = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_actions)
        If mvPaymentRow IsNot Nothing Then
          vHasAccessRights = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
          Me.Items(FinancialMenuItems.fmiEdit).Visible = vHasAccessRights AndAlso vCanMaintain AndAlso mvPaymentRow("ReceivedDate").ToString.Length = 0
          vHasOwner = mvDataRow("LogName").ToString.Length > 0
          Me.Items(FinancialMenuItems.fmiGoToActions).Visible = vHasAccessRights AndAlso vHasOwner AndAlso BooleanValue(mvPaymentRow("HasAction").ToString) AndAlso vCanMaintainActions
        Else
          Me.Items(FinancialMenuItems.fmiEdit).Visible = False
          Me.Items(FinancialMenuItems.fmiGoToActions).Visible = False
        End If
        Me.Items(FinancialMenuItems.fmiNewAdHocAction).Visible = vHasAccessRights AndAlso vHasOwner AndAlso vCanMaintainActions
        Me.Items(FinancialMenuItems.fmiNewActionFromTemplate).Visible = vHasAccessRights AndAlso vHasOwner AndAlso vCanMaintainActions
        If Not vShow Then vShow = vHasAccessRights AndAlso vHasOwner
      End If
      If Not vShow Then e.Cancel = True
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

End Class

Public Class PurchaseInvoiceMenu
  Inherits BaseFinancialMenu

  Private mvParentDataRow As DataRow

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub
  Public Sub SetContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvDataRow = pDataRow
  End Sub

  Public Sub SetParentContext(ByVal pContactInfo As ContactInfo, ByVal pDataRow As DataRow)
    mvContactInfo = pContactInfo
    mvParentDataRow = pDataRow
  End Sub
  Protected Overrides Sub FinancialMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Try
      SetVisibleItems(e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Overloads Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Dim vCursor As New BusyCursor
    Dim vCanEdit As Boolean
    Try
      Me.SuspendLayout()
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      e.Cancel = False
      vCanEdit = mvContactInfo IsNot Nothing AndAlso mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite AndAlso Not mvReadOnly
      Dim vHasAccessRights As Boolean = DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly
      Dim vShowItems(Me.Items.Count) As Boolean

      If mvDataRow IsNot Nothing Then

        If mvDataRow.Table.Columns.Contains("PurchaseInvoiceNumber") Then

          'set go to changed by menu
          Dim mvParams As New ParameterList(True)
          mvParams("PurchaseInvoiceNumber") = mvDataRow("PurchaseInvoiceNumber").ToString()
          Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)
          'no menus by default
          vShowItems(FinancialMenuItems.fmiGoToPopRevChangedBy) = False
          vShowItems(FinancialMenuItems.fmiGoToPopRevGoToChanges) = False

          If vResults("MenuSelection") = "2" Then 'no menus
          ElseIf vResults("MenuSelection") = "1" Then
            vShowItems(FinancialMenuItems.fmiGoToPopRevGoToChanges) = True
          ElseIf vResults("MenuSelection") = "0" Then
            vShowItems(FinancialMenuItems.fmiGoToPopRevChangedBy) = True
          End If
        End If
      End If
      Dim vVisibleCount As Integer
      Dim vShowItem As Boolean
      For vIndex As Integer = 0 To Me.Items.Count - 1
        Dim vItem As MenuToolbarCommand = DirectCast(Me.Items(vIndex).Tag, MenuToolbarCommand)
        vShowItem = vShowItems(vIndex) AndAlso vItem.HideItem = False
        Me.Items(vIndex).Visible = vShowItem
        If vShowItem Then vVisibleCount += 1
      Next
      If vVisibleCount = 0 Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      Me.ResumeLayout()
      vCursor.Dispose()
    End Try
  End Sub
End Class
