Imports Advanced.LanguageExtensions

Namespace Access
  Partial Public Class DataSelection
    Public Overridable Function DataTable() As CDBDataTable
      If mvParameters Is Nothing Then mvParameters = New CDBParameters
      Dim vDataTable As New CDBDataTable
      Select Case mvType
        Case DataSelectionTypes.dstQueryByExampleContacts, DataSelectionTypes.dstQueryByExampleOrganisations, DataSelectionTypes.dstQueryByExampleEvents
          vDataTable.AddColumnsFromList(mvSelectColumns)
        Case Else
          If mvResultColumns.Length > 0 Then vDataTable.AddColumnsFromList(mvResultColumns)
      End Select
      Select Case mvType
        Case DataSelectionTypes.dstActionContactLinks
          GetActionContactLinks(vDataTable, False)
        Case DataSelectionTypes.dstActionDocumentLinks
          GetActionDocumentLinks(vDataTable, False, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstActionLinks
          GetActionLinks(vDataTable)
        Case DataSelectionTypes.dstActionLinkEmailAddresses
          GetActionLinkEMailAddresses(vDataTable)
        Case DataSelectionTypes.dstActionOrganisationLinks
          GetActionOrganisationLinks(vDataTable, False)
        Case DataSelectionTypes.dstActionOutline
          GetActionOutline(vDataTable)
        Case DataSelectionTypes.dstPriorActions
          GetPriorActions(vDataTable)
        Case DataSelectionTypes.dstActionSubjects
          GetActionSubjects(vDataTable)
        Case DataSelectionTypes.dstActivitiesDataSheet
          GetActivitiesDataSheet(vDataTable)
        Case DataSelectionTypes.dstAppealBudgetDetails
          GetAppealBudgetDetails(vDataTable)
        Case DataSelectionTypes.dstAppealBudgets
          GetAppealBudgets(vDataTable)
        Case DataSelectionTypes.dstAppealCollections
          GetAppealCollections(vDataTable)
        Case DataSelectionTypes.dstAppealResources
          GetAppealResources(vDataTable)
        Case DataSelectionTypes.dstAppealTypes
          GetAppealTypes(vDataTable)
        Case DataSelectionTypes.dstBACSAmendments
          GetBACSAmendments(vDataTable)
        Case DataSelectionTypes.dstBrowserContactPositions
          GetBrowserContactPositions(vDataTable)
        Case DataSelectionTypes.dstCampaignAppeals
          GetCampaignAppeals(vDataTable)
        Case DataSelectionTypes.dstCampaignCollections
          GetCampaignCollections(vDataTable)
        Case DataSelectionTypes.dstCampaignInfo
          GetCampaignInfo(vDataTable)
        Case DataSelectionTypes.dstCampaigns
          GetCampaigns(vDataTable)
        Case DataSelectionTypes.dstCampaignRoles
          GetCampaignRoles(vDataTable)
        Case DataSelectionTypes.dstCampaignSegments
          GetCampaignSegments(vDataTable)
        Case DataSelectionTypes.dstClaimedPayments
          GetClaimedPayments(vDataTable)
        Case DataSelectionTypes.dstCollectionBoxesForPayment
          GetCollectionBoxesForPayment(vDataTable)
        Case DataSelectionTypes.dstCollectionPayments
          GetCollectionPayments(vDataTable)
        Case DataSelectionTypes.dstCollectionPIS, DataSelectionTypes.dsth2hCollectionPIS
          GetCollectionPIS(vDataTable)
        Case DataSelectionTypes.dstCollectionPoints
          GetCollectionPoints(vDataTable)
        Case DataSelectionTypes.dstCollectionRegions
          GetCollectionRegions(vDataTable)
        Case DataSelectionTypes.dstCollectionResources
          GetCollectionResources(vDataTable)
        Case DataSelectionTypes.dstCollectorShifts
          GetCollectorShifts(vDataTable)
        Case DataSelectionTypes.dstContactAccounts
          GetContactAccounts(vDataTable)
        Case DataSelectionTypes.dstContactActions
          GetContactActions(vDataTable)
        Case DataSelectionTypes.dstContactAddresses, DataSelectionTypes.dstContactAddressesWithUsages
          GetContactAddresses(vDataTable)
        Case DataSelectionTypes.dstContactAddressPositionAndOrg
          GetContactAddressPositionAndOrg(vDataTable)
        Case DataSelectionTypes.dstContactAddressUsages
          GetContactAddressUsages(vDataTable)
        Case DataSelectionTypes.dstContactAppointments
          GetContactAppointments(vDataTable)
        Case DataSelectionTypes.dstContactAppropriateCertificates
          GetContactAppropriateCertificates(vDataTable)
        Case DataSelectionTypes.dstContactBackOrders
          GetContactBackOrders(vDataTable)
        Case DataSelectionTypes.dstContactBankAccounts
          GetContactBankAccounts(vDataTable)
        Case DataSelectionTypes.dstContactCancelledProvisionalTrans
          GetContactCancelledProvisionalTrans(vDataTable)
        Case DataSelectionTypes.dstContactCashInvoices
          GetContactCashInvoices(vDataTable)
        Case DataSelectionTypes.dstContactCategories
          GetContactCategories(vDataTable)
        Case DataSelectionTypes.dstContactCategoryGraphData
          GetContactCategoryGraphData(vDataTable)
        Case DataSelectionTypes.dstContactCollectionPayments
          GetContactCollectionPayments(vDataTable)
        Case DataSelectionTypes.dstContactCommsInformation
          GetContactCommsInformation(vDataTable)
        Case DataSelectionTypes.dstContactCommsNumbers, DataSelectionTypes.dstContactCommsNumbersWithUsages
          GetContactCommsNumbers(vDataTable)
        Case DataSelectionTypes.dstContactCommsNumbersEdit
          GetContactCommsNumbersEdit(vDataTable)
        Case DataSelectionTypes.dstContactCommunicationUsages
          GetContactCommunicationUsages(vDataTable)
        Case DataSelectionTypes.dstContactCovenants
          GetContactCovenants(vDataTable)
        Case DataSelectionTypes.dstContactCPDCyclesEdit
          GetContactCPDCyclesEdit(vDataTable)
        Case DataSelectionTypes.dstContactCreditCardAuthorities
          GetContactCreditCardAuthorities(vDataTable)
        Case DataSelectionTypes.dstContactCreditCards
          GetContactCreditCards(vDataTable)
        Case DataSelectionTypes.dstContactCreditCustomers
          GetContactCreditCustomers(vDataTable)
        Case DataSelectionTypes.dstContactDBANotes
          GetContactDBANotes(vDataTable)
        Case DataSelectionTypes.dstContactDepartmentHistory
          GetContactDepartmentHistory(vDataTable)
        Case DataSelectionTypes.dstContactDepartmentNotes
          GetContactDepartmentNotes(vDataTable)
        Case DataSelectionTypes.dstContactDespatchNotes
          GetContactDespatchNotes(vDataTable)
        Case DataSelectionTypes.dstContactDirectDebits
          GetContactDirectDebits(vDataTable)
        Case DataSelectionTypes.dstContactEMailAddresses
          GetContactEMailAddresses(vDataTable)
        Case DataSelectionTypes.dstContactEventBookings
          GetContactEventBookings(vDataTable)
        Case DataSelectionTypes.dstContactEventDelegates
          GetContactEventDelegates(vDataTable)
        Case DataSelectionTypes.dstContactEventOrganiser
          GetContactEventOrganiser(vDataTable)
        Case DataSelectionTypes.dstContactEventPersonnel
          GetContactEventPersonnel(vDataTable)
        Case DataSelectionTypes.dstContactEventRoomBookings
          GetContactEventRoomBookings(vDataTable)
        Case DataSelectionTypes.dstContactEventRoomsAllocated
          GetContactEventRoomsAllocated(vDataTable)
        Case DataSelectionTypes.dstContactEventSessions
          GetContactEventSessions(vDataTable)
        Case DataSelectionTypes.dstContactExternalReferences
          GetContactExternalReferences(vDataTable)
        Case DataSelectionTypes.dstContactFinder, DataSelectionTypes.dstOrganisationFinder,
             DataSelectionTypes.dstActionFinder, DataSelectionTypes.dstContactMailingDocumentsFinder,
             DataSelectionTypes.dstMailingFinder, DataSelectionTypes.dstEventPersonnelFinder,
             DataSelectionTypes.dstEventPersonnelAppointmentFinder, DataSelectionTypes.dstTextSearch,
             DataSelectionTypes.dstExamPersonnelFinder, DataSelectionTypes.dstEventFinder
          GetFinder(vDataTable)
        Case DataSelectionTypes.dstContactFinLinksDonated
          GetContactFinLinksDonated(vDataTable)
        Case DataSelectionTypes.dstContactFinLinksReceived
          GetContactFinLinksReceived(vDataTable)
        Case DataSelectionTypes.dstContactFundraisingEvents
          GetContactFundraisingEvents(vDataTable)
        Case DataSelectionTypes.dstContactFundraisingEventFinder
          GetContactFundraisingEventFinder(vDataTable)
        Case DataSelectionTypes.dstContactFundRaisingRequests
          GetContactFundraisingRequests(vDataTable)
        Case DataSelectionTypes.dstContactGAYEPledgePayments
          GetContactGAYEPledgePayments(vDataTable)
        Case DataSelectionTypes.dstContactGAYEPledges
          GetContactGAYEPledges(vDataTable)
        Case DataSelectionTypes.dstContactGAYEPostTaxPledgePayments
          GetContactGAYEPostTaxPledgePayments(vDataTable)
        Case DataSelectionTypes.dstContactGAYEPostTaxPledges
          GetContactGAYEPostTaxPledges(vDataTable)
        Case DataSelectionTypes.dstContactGiftAidDeclarations
          GetContactGiftAidDeclarations(vDataTable)
        Case DataSelectionTypes.dstContactH2HCollections
          GetContactH2HCollections(vDataTable)
        Case DataSelectionTypes.dstContactHPCategories
          GetContactCategories(vDataTable, False, True, True)
        Case DataSelectionTypes.dstContactHPCategoryValues
          GetContactCategories(vDataTable, False, True, True)
        Case DataSelectionTypes.dstContactDeptCategories
          GetContactCategories(vDataTable, True, False, True)
        Case DataSelectionTypes.dstContactHeaderHPCategories
          GetContactCategories(vDataTable, False, True, True)
        Case DataSelectionTypes.dstContactHeaderDeptCategories
          GetContactCategories(vDataTable, True, False, True)
        Case DataSelectionTypes.dstContactHPLinks, DataSelectionTypes.dstContactHeaderHPLinks
          GetContactHPLinks(vDataTable)
        Case DataSelectionTypes.dstContactInformation
          GetContactInformation(vDataTable)
        Case DataSelectionTypes.dstContactJournal
          GetContactJournal(vDataTable)
        Case DataSelectionTypes.dstContactLegacy
          GetContactLegacy(vDataTable)
        Case DataSelectionTypes.dstContactLegacyActions
          GetContactLegacyActions(vDataTable)
        Case DataSelectionTypes.dstContactLegacyAssets
          GetContactLegacyAssets(vDataTable)
        Case DataSelectionTypes.dstContactLegacyBequests
          GetContactLegacyBequests(vDataTable)
        Case DataSelectionTypes.dstContactLegacyExpenses
          GetContactLegacyExpenses(vDataTable)
        Case DataSelectionTypes.dstContactLegacyLinks
          GetContactLegacyLinks(vDataTable)
        Case DataSelectionTypes.dstContactLegacyTaxCertificates
          GetContactLegacyTaxCertificates(vDataTable)
        Case DataSelectionTypes.dstContactLinksTo, DataSelectionTypes.dstContactLinksFrom
          GetContactLinksToOrFrom(vDataTable, mvType)
        Case DataSelectionTypes.dstContactMailings
          GetContactMailings(vDataTable)
        Case DataSelectionTypes.dstContactMannedCollections
          GetContactMannedCollections(vDataTable)
        Case DataSelectionTypes.dstContactMembershipDetails
          GetContactMembershipDetails(vDataTable)
        Case DataSelectionTypes.dstContactMemberships
          GetContactMemberships(vDataTable)
        Case DataSelectionTypes.dstContactNotifications
          GetContactNotifications(vDataTable)
        Case DataSelectionTypes.dstContactOutstandingInvoices
          GetContactOutstandingInvoices(vDataTable)
        Case DataSelectionTypes.dstContactOwners
          GetContactOwners(vDataTable)
        Case DataSelectionTypes.dstContactPaymentPlans
          GetContactPaymentPlans(vDataTable)
        Case DataSelectionTypes.dstContactPaymentPlansPayments
          GetContactPaymentPlansPayments(vDataTable)
        Case DataSelectionTypes.dstContactPerformances
          GetContactPerformances(vDataTable)
        Case DataSelectionTypes.dstContactPictureDocuments
          GetContactPictureDocuments(vDataTable)
        Case DataSelectionTypes.dstContactPositionCategories
          GetContactPositionCategories(vDataTable)
        Case DataSelectionTypes.dstContactPositionLinks
          GetContactPositionLinks(vDataTable)
        Case DataSelectionTypes.dstContactPositions
          GetContactPositions(vDataTable)
        Case DataSelectionTypes.dstContactPreviousNames
          GetContactPreviousNames(vDataTable)
        Case DataSelectionTypes.dstContactProcessedTransactions
          GetContactProcessedTransactions(vDataTable)
        Case DataSelectionTypes.dstContactPurchaseInvoices
          GetContactPurchaseInvoices(vDataTable)
        Case DataSelectionTypes.dstContactPurchaseOrders
          GetContactPurchaseOrders(vDataTable)
        Case DataSelectionTypes.dstContactRegisteredUsers
          GetContactRegisteredUsers(vDataTable)
        Case DataSelectionTypes.dstContactRoles
          GetContactRoles(vDataTable)
        Case DataSelectionTypes.dstContactSalesLedgerItems
          GetContactSalesLedgerItems(vDataTable)
        Case DataSelectionTypes.dstContactScores
          GetContactScores(vDataTable)
        Case DataSelectionTypes.dstContactServiceBookings
          GetContactServiceBookings(vDataTable)
        Case DataSelectionTypes.dstContactSourceFromLastMailing
          GetContactSourceFromLastMailing(vDataTable)
        Case DataSelectionTypes.dstContactSponsorshipClaimedPayments
          GetContactSponsorshipClaimedPayments(vDataTable)
        Case DataSelectionTypes.dstContactSponsorshipUnClaimedPayments
          GetContactSponsorshipUnClaimedPayments(vDataTable)
        Case DataSelectionTypes.dstContactStandingOrders
          GetContactStandingOrders(vDataTable)
        Case DataSelectionTypes.dstContactStatusHistory
          GetContactStatusHistory(vDataTable)
        Case DataSelectionTypes.dstContactStickyNotes
          GetContactStickyNotes(vDataTable)
        Case DataSelectionTypes.dstContactSubscriptions
          GetContactSubscriptions(vDataTable)
        Case DataSelectionTypes.dstContactSuppressions
          GetContactSuppressions(vDataTable)
        Case DataSelectionTypes.dstContactUnMannedCollections
          GetContactUnMannedCollections(vDataTable)
        Case DataSelectionTypes.dstContactUnProcessedTransactions
          GetContactUnProcessedTransactions(vDataTable)
        Case DataSelectionTypes.dstCovenantGiftAidClaims
          GetCovenantGiftAidClaims(vDataTable)
        Case DataSelectionTypes.dstCovenentClaims
          GetCovenentClaims(vDataTable)
        Case DataSelectionTypes.dstCovenentPayments
          GetCovenentPayments(vDataTable)
        Case DataSelectionTypes.dstCPDDetails
          GetCPDDetails(vDataTable)
        Case DataSelectionTypes.dstCPDPointsEdit
          GetCPDPointsEdit(vDataTable)
        Case DataSelectionTypes.dstCPDSummary
          GetCPDSummary(vDataTable)
        Case DataSelectionTypes.dstContactCPDPointsWithoutCycle
          GetContactCPDPointsWithoutCycle(vDataTable)
        Case DataSelectionTypes.dstCriteriaSetDetails
          GetCriteriaSetDetails(vDataTable)
        Case DataSelectionTypes.dstGeneralMailingSelectionSets
          GetGeneralMailingSelectionSets(vDataTable)
        Case DataSelectionTypes.dstCriteriaSets
          GetCriteriaSets(vDataTable)
        Case DataSelectionTypes.dstCustomFormData
          GetCustomFormData(vDataTable)
        Case DataSelectionTypes.dstDashboardData
          GetDashboardData(vDataTable)
        Case DataSelectionTypes.dstDelegateActivities
          GetDelegateActivities(vDataTable)
        Case DataSelectionTypes.dstDelegateLinks
          GetDelegateLinks(vDataTable)
        Case DataSelectionTypes.dstDepartmentActivities
          GetDepartmentActivities(vDataTable)
        Case DataSelectionTypes.dstDepartmentActivityValues
          GetDepartmentActivityValues(vDataTable)
        Case DataSelectionTypes.dstDespatchStock
          GetDespatchStock(vDataTable)
        Case DataSelectionTypes.dstDocumentContactLinks
          GetDocumentContactLinks(vDataTable, False)
        Case DataSelectionTypes.dstDocumentDocumentLinks
          GetDocumentDocumentLinks(vDataTable, False, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstDocumentHistory
          GetDocumentHistory(vDataTable)
        Case DataSelectionTypes.dstDocumentLinks
          GetDocumentLinks(vDataTable, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstDocumentOrganisationLinks
          GetDocumentOrganisationLinks(vDataTable, False)
        Case DataSelectionTypes.dstDocumentRelatedDocuments
          GetDocumentRelatedDocuments(vDataTable)
        Case DataSelectionTypes.dstCommunicationsLogDocClass
          GetCommunicationsLogDocClass(vDataTable)
        Case DataSelectionTypes.dstMeetings
          GetMeetings(vDataTable)
        Case DataSelectionTypes.dstContactMeetings
          GetContactMeetings(vDataTable)
        Case DataSelectionTypes.dstDocuments, DataSelectionTypes.dstDistinctDocuments, DataSelectionTypes.dstDistinctExternalDocuments, DataSelectionTypes.dstContactDocuments, DataSelectionTypes.dstDistinctContactDocuments, DataSelectionTypes.dstEventDocuments
          GetDocuments(vDataTable, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstDocumentSubjects
          GetDocumentSubjects(vDataTable)
        Case DataSelectionTypes.dstDocumentTransactionLinks
          GetDocumentTransactionLinks(vDataTable, False)
        Case DataSelectionTypes.dstDuplicateContacts
          GetDuplicateContacts(vDataTable)
        Case DataSelectionTypes.dstDuplicateOrganisations
          GetDuplicateOrganisations(vDataTable)
        Case DataSelectionTypes.dstEMailContacts
          GetEMailContacts(vDataTable)
        Case DataSelectionTypes.dstEMailOrganisations
          GetEMailOrganisations(vDataTable)
        Case DataSelectionTypes.dstEventAccommodation
          GetEventAccommodation(vDataTable)
        Case DataSelectionTypes.dstEventAttendees
          GetEventAttendees(vDataTable)
        Case DataSelectionTypes.dstEventCurrentAttendees
          GetEventCurrentAttendees(vDataTable)
        Case DataSelectionTypes.dstEventAuthoriseExpenses
          GetEventAuthoriseExpenses(vDataTable)
        Case DataSelectionTypes.dstEventBookingDelegates
          GetEventBookingDelegates(vDataTable)
        Case DataSelectionTypes.dstEventBookingOptions
          GetEventBookingOptions(vDataTable)
        Case DataSelectionTypes.dstEventBookingOptionSessions
          GetEventBookingOptionSessions(vDataTable)
        Case DataSelectionTypes.dstEventBookings
          GetEventBookings(vDataTable)
        Case DataSelectionTypes.dstEventBookingSessions
          GetEventBookingSessions(vDataTable)
        Case DataSelectionTypes.dstEventCandidates
          GetEventCandidates(vDataTable)
        Case DataSelectionTypes.dstEventContacts
          GetEventContacts(vDataTable)
        Case DataSelectionTypes.dstEventCosts
          GetEventCosts(vDataTable)
        Case DataSelectionTypes.dstEventDelegateIncome
          GetEventDelegateIncome(vDataTable)
        Case DataSelectionTypes.dstEventFinancialHistory
          GetEventFinancialHistory(vDataTable)
        Case DataSelectionTypes.dstEventFinancialLinks
          GetEventFinancialLinks(vDataTable)
        Case DataSelectionTypes.dstEventLoanItems
          GetEventLoanItems(vDataTable)
        Case DataSelectionTypes.dstEventMailings
          GetEventMailings(vDataTable)
        Case DataSelectionTypes.dstEventOrganiserData
          GetEventOrganiserData(vDataTable)
        Case DataSelectionTypes.dstEventOwners
          GetEventOwners(vDataTable)
        Case DataSelectionTypes.dstEventPersonnel
          GetEventPersonnel(vDataTable)
        Case DataSelectionTypes.dstEventPersonnelTasks
          GetEventPersonnelTasks(vDataTable)
        Case DataSelectionTypes.dstEventPIS
          GetEventPIS(vDataTable)
        Case DataSelectionTypes.dstEventResources
          GetEventResources(vDataTable)
        Case DataSelectionTypes.dstEventResults
          GetEventResults(vDataTable)
        Case DataSelectionTypes.dstEventRoomBookingAllocations
          GetEventRoomBookingAllocations(vDataTable)
        Case DataSelectionTypes.dstEventRoomBookings
          GetEventRoomBookings(vDataTable)
        Case DataSelectionTypes.dstEventSessionActivities
          GetEventSessionActivities(vDataTable)
        Case DataSelectionTypes.dstEventSessions
          GetEventSessions(vDataTable)
        Case DataSelectionTypes.dstEventSessionTests
          GetEventSessionTests(vDataTable)
        Case DataSelectionTypes.dstEventSources
          GetEventSources(vDataTable)
        Case DataSelectionTypes.dstEventSubmissions
          GetEventSubmissions(vDataTable)
        Case DataSelectionTypes.dstEventTopics
          GetEventTopics(vDataTable)
        Case DataSelectionTypes.dstEventVenueBookings
          GetEventVenueBookings(vDataTable)
        Case DataSelectionTypes.dstEventVenueData
          GetEventVenueData(vDataTable)
        Case DataSelectionTypes.dstEVWaitingBookings
          GetEVWaitingBookings(vDataTable)
        Case DataSelectionTypes.dstEVWaitingDelegates
          GetEVWaitingDelegates(vDataTable)
        Case DataSelectionTypes.dstFinancialHistoryDetails
          GetFinancialHistoryDetails(vDataTable)
        Case DataSelectionTypes.dstFundraisingEventAnalysis
          GetFundraisingEventAnalysis(vDataTable)
        Case DataSelectionTypes.dstFundraisingRequestTargets
          GetFundraisingRequestTargets(vDataTable)
        Case DataSelectionTypes.dstFundRequestExpectedAmountHistory
          GetFundRequestExpectedAmountHistory(vDataTable)
        Case DataSelectionTypes.dstFundRequestStatusHistory
          GetFundRequestStatusHistory(vDataTable)
        Case DataSelectionTypes.dstFundraisingPaymentHistory
          GetFundraisingPaymentHistory(vDataTable)
        Case DataSelectionTypes.dstFundraisingPaymentSchedule
          GetFundraisingPaymentSchedule(vDataTable)
        Case DataSelectionTypes.dstFundraisingActions
          GetFundraisingActions(vDataTable)
        Case DataSelectionTypes.dstGeographicalRegions
          GetGeographicalRegions(vDataTable)
        Case DataSelectionTypes.dstH2HCollectors
          GetH2HCollectors(vDataTable)
        Case DataSelectionTypes.dstIncentives
          GetIncentives(vDataTable)
        Case DataSelectionTypes.dstFulFilledContactIncentives
          GetFulFilledContactIncentives(vDataTable)
        Case DataSelectionTypes.dstUnFulFilledContactIncentives
          GetUnFulFilledContactIncentives(vDataTable)
        Case DataSelectionTypes.dstFulFilledPayPlanIncentives
          GetFulFilledPayPlanIncentives(vDataTable)
        Case DataSelectionTypes.dstUnFulFilledPayPlanIncentives
          GetUnFulFilledPayPlanIncentives(vDataTable)
        Case DataSelectionTypes.dstLegacyBequestForecasts
          GetLegacyBequestForecasts(vDataTable)
        Case DataSelectionTypes.dstLegacyBequestReceipts
          GetLegacyBequestReceipts(vDataTable)
        Case DataSelectionTypes.dstMannedCollectionBoxes, DataSelectionTypes.dstUnMannedCollectionBoxes, DataSelectionTypes.dstContactCollectionBoxes
          GetMannedCollectionBoxes(vDataTable)
        Case DataSelectionTypes.dstMannedCollectors
          GetMannedCollectors(vDataTable)
        Case DataSelectionTypes.dstMeetingContactLinks
          GetMeetingContactLinks(vDataTable, False)
        Case DataSelectionTypes.dstMeetingDocumentLinks
          GetMeetingDocumentLinks(vDataTable, False, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstMeetingLinks
          GetMeetingLinks(vDataTable, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
        Case DataSelectionTypes.dstMeetingOrganisationLinks
          GetMeetingOrganisationLinks(vDataTable, False)
        Case DataSelectionTypes.dstMembershipChanges
          GetMembershipChanges(vDataTable)
        Case DataSelectionTypes.dstMembershipGroupHistory
          GetMembershipGroupHistory(vDataTable)
        Case DataSelectionTypes.dstMembershipGroups
          GetMembershipGroups(vDataTable)
        Case DataSelectionTypes.dstMembershipOtherMembers
          GetMembershipOtherMembers(vDataTable)
        Case DataSelectionTypes.dstMembershipPaymentPlanDetails
          GetMembershipPaymentPlanDetails(vDataTable)
        Case DataSelectionTypes.dstOrganisationContactCommsNumbers, DataSelectionTypes.dstContactHeaderCommsNumbers
          GetOrganisationContactCommsNumbers(vDataTable)
        Case DataSelectionTypes.dstPackProductDataSheet
          GetPackProductDataSheet(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanAmendmentHistory
          GetPaymentPlanAmendmentHistory(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanDetails
          GetPaymentPlanDetails(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanMembers
          GetPaymentPlanMembers(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanOutstandingOPS
          GetPaymentPlanOutstandingOPS(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanPaymentDetails
          GetPaymentPlanPaymentDetails(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanPayments
          GetPaymentPlanPayments(vDataTable)
        Case DataSelectionTypes.dstPaymentPlanSubscriptions
          GetPaymentPlanSubscriptions(vDataTable)
        Case DataSelectionTypes.dstPersonnelContacts
          GetPersonnelContacts(vDataTable)
        Case DataSelectionTypes.dstPostPointRecipients
          GetPostPointRecipients(vDataTable)
        Case DataSelectionTypes.dstProductWarehouses
          GetProductWarehouses(vDataTable)
        Case DataSelectionTypes.dstPurchaseInvoiceDetails
          GetPurchaseInvoiceDetails(vDataTable)
        Case DataSelectionTypes.dstPurchaseOrderDetails
          GetPurchaseOrderDetails(vDataTable)
        Case DataSelectionTypes.dstPurchaseOrderPayments
          GetPurchaseOrderPayments(vDataTable)
        Case DataSelectionTypes.dstPurchaseOrderInformation
          GetPurchaseOrderInformation(vDataTable)
        Case DataSelectionTypes.dstChequeInformation
          GetChequeInformation(vDataTable)
        Case DataSelectionTypes.dstPurchaseInvoiceInformation
          GetPurchaseInvoiceInformation(vDataTable)
        Case DataSelectionTypes.dstRates
          GetRates(vDataTable)
        Case DataSelectionTypes.dstRelationshipsDataSheet
          GetRelationshipsDataSheet(vDataTable)
        Case DataSelectionTypes.dstSalesContacts
          GetSalesContacts(vDataTable)
        Case DataSelectionTypes.dstSegmentCostCentres
          GetSegmentCostCentres(vDataTable)
        Case DataSelectionTypes.dstSegmentProducts
          GetSegmentProducts(vDataTable)
        Case DataSelectionTypes.dstSelectionSetAppointments
          GetSelectionSetAppointments(vDataTable)
        Case DataSelectionTypes.dstSelectionSetCommsNumbers
          GetSelectionSetCommsNumbers(vDataTable)
        Case DataSelectionTypes.dstSelectionSetContacts
          GetSelectionSetContacts(vDataTable)
        Case DataSelectionTypes.dstSelectionSteps
          GetSelectionSteps(vDataTable)
        Case DataSelectionTypes.dstSelectItemAddresses
          GetSelectItemAddresses(vDataTable)
        Case DataSelectionTypes.dstSelectItemCreditAccount
          GetSelectItemCreditAccount(vDataTable)
        Case DataSelectionTypes.dstSelectItemSelectionSets
          GetSelectItemSelectionSets(vDataTable)
        Case DataSelectionTypes.dstServiceControlRestrictions
          GetServiceControlRestrictions(vDataTable)
        Case DataSelectionTypes.dstServiceStartDays
          GetServiceStartDays(vDataTable)
        Case DataSelectionTypes.dstSuppliers
          GetSuppliers(vDataTable)
        Case DataSelectionTypes.dstSuppressionDataSheet
          GetSuppressionDataSheet(vDataTable)
        Case DataSelectionTypes.dstTickBoxes
          GetTickBoxes(vDataTable)
        Case DataSelectionTypes.dstTopicDataSheet
          GetTopicsDataSheet(vDataTable)
        Case DataSelectionTypes.dstTransactionAnalysis
          GetTransactionAnalysis(vDataTable)
        Case DataSelectionTypes.dstTransactionDetails
          GetTransactionDetails(vDataTable)
        Case DataSelectionTypes.dstUnauthorisedPOPayments
          GetUnauthorisedPOPayments(vDataTable)
        Case DataSelectionTypes.dstUnClaimedPayments
          GetUnClaimedPayments(vDataTable)
        Case DataSelectionTypes.dstVariableParameters
          GetVariableParameters(vDataTable)
        Case DataSelectionTypes.dstBatchProcessingInformation
          GetBatchProcessingInformation(vDataTable)
        Case DataSelectionTypes.dstPickingListDetails
          GetPickingListDetails(vDataTable)
        Case DataSelectionTypes.dstCampaignCosts
          GetCampaignCosts(vDataTable)
        Case DataSelectionTypes.dstEventBookingTransactions
          GetEventBookingTransactions(vDataTable)
        Case DataSelectionTypes.dstMembershipSummaryMembers
          GetMembershipSummaryMembers(vDataTable)
        Case DataSelectionTypes.dstContactEmailingsLinks
          GetContactEmailingsLinks(vDataTable)
        Case DataSelectionTypes.dstPurchaseInvoiceChequeInformation
          GetPurchaseInvoiceChequeInformation(vDataTable)
        Case DataSelectionTypes.dstContactAlerts
          GetContactAlerts(vDataTable)
        Case DataSelectionTypes.dstContactCommunicationHistory
          GetContactCommunicationHistory(vDataTable)
        Case DataSelectionTypes.dstQueryByExampleContacts
          GetQueryByExampleContacts(vDataTable)
        Case DataSelectionTypes.dstQueryByExampleOrganisations
          GetQueryByExampleOrganisations(vDataTable)
        Case DataSelectionTypes.dstQueryByExampleEvents
          GetQueryByExampleEvents(vDataTable)
        Case DataSelectionTypes.dstDespatchNotes
          GetDespatchNotes(vDataTable)
        Case DataSelectionTypes.dstDuplicateContactRecords
          GetDuplicateContactRecords(vDataTable)
        Case DataSelectionTypes.dstFindMeeting
          GetMeetingRecords(vDataTable)
        Case DataSelectionTypes.dstSelectAwaitListConfirmation
          GetAwaitListConfirmation(vDataTable)
        Case DataSelectionTypes.dstContactAddressAndUsage
          GetContactAddressAndUsage(vDataTable)
          'Case DataSelectionTypes.dstGeneralMailingSelectionSets
          '  GetGeneralMailingSelectionSets(vDataTable)
        Case DataSelectionTypes.dstPackedProductDataSheet
          GetPackedProductDataSheet(vDataTable)
        Case DataSelectionTypes.dstJobSchedules
          GetJobSchedule(vDataTable)
        Case DataSelectionTypes.dstJobProcessors
          GetJobProcessors(vDataTable)
        Case DataSelectionTypes.dstConfig
          GetConfig(vDataTable)
        Case DataSelectionTypes.dstSystemModuleUsers
          GetSystemModuleUsers(vDataTable)
        Case DataSelectionTypes.dstReportData
          GetReportData(vDataTable)
        Case DataSelectionTypes.dstReportSectionData
          GetReportSectionData(vDataTable)
        Case DataSelectionTypes.dstReportParameters
          GetReportParameters(vDataTable)
        Case DataSelectionTypes.dstReportSectionDetail
          GetReportSectionDetail(vDataTable)
        Case DataSelectionTypes.dstReportVersion
          GetReportVersion(vDataTable)
        Case DataSelectionTypes.dstReportControl
          GetReportControl(vDataTable)
        Case DataSelectionTypes.dstOwnershipData
          GetOwnershipGroupInformation(vDataTable)
        Case DataSelectionTypes.dstOwnershipUsers
          GetOwnershipUsers(vDataTable)
        Case DataSelectionTypes.dstOwnershipDepartments
          GetOwnershipDepartmentInformation(vDataTable)
        Case DataSelectionTypes.dstOwnershipUserInformation
          GetOwnershipUserInformation(vDataTable)
        Case DataSelectionTypes.dstServiceProductContacts
          GetServiceProductContacts(vDataTable)
        Case DataSelectionTypes.dstEmailAutoReplyText
          GetEmailAutoReplyText(vDataTable)
        Case DataSelectionTypes.dstEntityAlerts
          GetEntityAlerts(vDataTable)
        Case DataSelectionTypes.dstSalesTransactions
          GetContactUnProcessedTransactions(vDataTable)
        Case DataSelectionTypes.dstDeliveryTransactions
          GetContactUnProcessedTransactions(vDataTable)
        Case DataSelectionTypes.dstSalesTransactionAnalysis
          GetTransactionAnalysis(vDataTable)
        Case DataSelectionTypes.dstDeliveryTransactionAnalysis
          GetTransactionAnalysis(vDataTable)
        Case DataSelectionTypes.dstEventDelegateMailing
          GetContactMailings(vDataTable)
        Case DataSelectionTypes.dstCPDObjectives, DataSelectionTypes.dstCPDObjectivesEdit
          GetCPDObjectives(vDataTable)
        Case DataSelectionTypes.dstEventActions
          GetEventActions(vDataTable)
        Case DataSelectionTypes.dstAppealActions
          GetAppealActions(vDataTable)
        Case DataSelectionTypes.dstContactSurveys
          GetContactSurveys(vDataTable)
        Case DataSelectionTypes.dstContactSurveyResponses
          GetContactSurveyResponses(vDataTable)
        Case DataSelectionTypes.dstContactSalesLedgerReceipts
          GetContactSalesLedgerReceipts(vDataTable)
        Case DataSelectionTypes.dstWebProducts
          GetWebProducts(vDataTable)
        Case DataSelectionTypes.dstWebEvents
          GetWebEvents(vDataTable)
        Case DataSelectionTypes.dstWebBookingOptions
          GetWebBookingOptions(vDataTable)
        Case DataSelectionTypes.dstContactEventBookingDelegates
          GetContactEventBookingDelegates(vDataTable)
        Case DataSelectionTypes.dstDocumentActions
          GetDocumentActions(vDataTable)
        Case DataSelectionTypes.dstWebMembershipType
          GetWebMembershipTypes(vDataTable)
        Case DataSelectionTypes.dstWebEventBookings
          GetWebEventBookings(vDataTable)
        Case DataSelectionTypes.dstServiceBookingTransactions
          GetServiceBookingTransactionData(vDataTable)
        Case DataSelectionTypes.dstActivityFromActivityGroup
          GetActivityFromActivityGroup(vDataTable)
        Case DataSelectionTypes.dstEventCategories
          GetEventCategories(vDataTable)
        Case DataSelectionTypes.dstWebSurveys
          GetWebSurveys(vDataTable)
        Case DataSelectionTypes.dstWebDirectoryEntries
          GetWebDirectoryEntries(vDataTable)
        Case DataSelectionTypes.dstBankTransactions
          GetBankTransactions(vDataTable)
        Case DataSelectionTypes.dstWebDocuments
          GetWebDocuments(vDataTable)
        Case DataSelectionTypes.dstWebRelatedOrganisations
          GetWebRelatedOrganisations(vDataTable)
        Case DataSelectionTypes.dstWebRelatedContacts
          GetWebRelatedContacts(vDataTable)
        Case DataSelectionTypes.dstWebContacts
          GetWebContacts(vDataTable)
        Case DataSelectionTypes.dstContactAppointmentDetails
          GetContactAppointmentDetails(vDataTable)
        Case DataSelectionTypes.dstPurchaseOrderHistory
          GetPurchaseOrderHistory(vDataTable)
        Case DataSelectionTypes.dstContactLoans
          GetContactLoans(vDataTable)
        Case DataSelectionTypes.dstContactDirectoryUsage
          GetDirectoryUsage(vDataTable)
        Case DataSelectionTypes.dstContactExamSummary
          GetContactExamSummary(vDataTable)
        Case DataSelectionTypes.dstContactExamSummaryItems
          GetContactExamSummaryItems(vDataTable)
        Case DataSelectionTypes.dstContactExamSummaryList
          GetContactExamSummaryList(vDataTable)
        Case DataSelectionTypes.dstContactExamDetails
          GetContactExamDetails(vDataTable)
        Case DataSelectionTypes.dstContactExamDetailItems
          GetContactExamDetailItems(vDataTable)
        Case DataSelectionTypes.dstContactExamDetailList
          GetContactExamDetailList(vDataTable)
        Case DataSelectionTypes.dstContactExamExemptions
          GetContactExamExemptions(vDataTable)
        Case DataSelectionTypes.dstDataUpdates
          GetDataUpdatesData(vDataTable)
        Case DataSelectionTypes.dstContactAddressesAndPositions
          GetContactAddressesAndPositions(vDataTable)
        Case DataSelectionTypes.dstDuplicateOrganisationsForRegistration
          GetDuplicateOrganisationsForRegistration(vDataTable)
        Case DataSelectionTypes.dstLoanInterestRates
          GetLoanInterestRates(vDataTable)
        Case DataSelectionTypes.dstWebExams
          GetWebExams(vDataTable)
        Case DataSelectionTypes.dstWebMemberOrganisations
          GetWebMemberOrganisations(vDataTable)
        Case DataSelectionTypes.dstWebExamBookings
          GetWebExamBookings(vDataTable)
        Case DataSelectionTypes.dstWebExamHistory
          GetWebExamHistory(vDataTable)
        Case DataSelectionTypes.dstContactAmendments
          GetContactAmendments(vDataTable)
        Case DataSelectionTypes.dstContactAmendmentDetails
          GetContactAmendmentDetails(vDataTable)
        Case DataSelectionTypes.dstMeetingActions
          GetMeetingActions(vDataTable)
        Case DataSelectionTypes.dstContactExamCertificates
          GetContactExamCertificates(vDataTable)
        Case DataSelectionTypes.dstContactExamCertificateItems
          GetContactExamCertificateItems(vDataTable)
        Case DataSelectionTypes.dstContactExamCertificateReprints
          GetContactExamCertificateReprints(vDataTable)
        Case DataSelectionTypes.dstFundraisingRequests
          GetFundraisingRequests(vDataTable)
        Case DataSelectionTypes.dstFundraisingDocuments
          GetDocuments(vDataTable)
        Case DataSelectionTypes.dstExamScheduleFinder
          GetExamSchedules(vDataTable)
        Case DataSelectionTypes.dstContactTokens
          GetContactTokens(vDataTable)
        Case DataSelectionTypes.dstEventSessionCPD
          GetEventSessionCPD(vDataTable)
        Case DataSelectionTypes.dstContactCPDCycleDocuments
          GetDocuments(vDataTable, False)
        Case DataSelectionTypes.dstContactCPDPointDocuments
          GetDocuments(vDataTable, False)
        Case DataSelectionTypes.dstFindCPDCyclePeriods
          GetFindCPDCyclePeriods(vDataTable)
        Case DataSelectionTypes.dstFindCPDPoints
          GetFindCPDPoints(vDataTable)
        Case DataSelectionTypes.dstContactViewOrganisations
          GetContactViewOrganisations(vDataTable)
        Case DataSelectionTypes.dstContactPositionActions
          GetPositionActions(vDataTable)
        Case DataSelectionTypes.dstContactPositionDocuments
          GetPositionDocuments(vDataTable)
        Case DataSelectionTypes.dstContactPositionTimesheets
          GetPositionTimesheets(vDataTable)
        Case DataSelectionTypes.dstSalesLedgerAnalysis
          GetTransactionSalesLedgerAnalysis(vDataTable)
        Case DataSelectionTypes.dstTraderAlerts
          GetTraderAlerts(vDataTable)
        Case DataSelectionTypes.dstContactFinanceAlerts
          GetContactFinanceAlerts(vDataTable)
      End Select
      Return vDataTable
    End Function

    Private Sub GetWebContacts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String
      Dim vFields As String

      If mvParameters.Exists("EmailAddress") Then
        vAttrs = "c.address_number,c.date_of_birth,cm.number AS email_address,c.status,st.status_desc,member_number,m.membership_status,ms.membership_status_desc,cg.contact_group_desc," & vContact.GetRecordSetFieldsNamePhoneGroup
      Else
        vAttrs = "c.address_number,c.date_of_birth,'' AS email_address,c.status,st.status_desc,member_number,m.membership_status,ms.membership_status_desc,cg.contact_group_desc," & vContact.GetRecordSetFieldsNamePhoneGroup
      End If
      vFields = "address_number,date_of_birth,email_address,status,status_desc,member_number,membership_status,membership_status_desc,ni_number,contact_group_desc,contact_number,title,forenames,initials,surname"
      Dim vEmailDeviceCode As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEmailDevice)

      vAnsiJoins.Add("addresses a", "c.address_number", "a.address_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      If mvParameters.Exists("EmailAddress") Then
        vAnsiJoins.Add("communications cm", "c.contact_number", "cm.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If

      Dim vMemWhereFields As New CDBFields
      vMemWhereFields.Add("cancellation_reason", "")
      Dim vMemSQL As New SQLStatement(mvEnv.Connection, "m.contact_number, m.member_number,m.membership_status", "members m", vMemWhereFields)
      Dim vActiveMemJoin As String = String.Format("({0}) m", vMemSQL.SQL)

      vAnsiJoins.Add(vActiveMemJoin, "c.contact_number", "m.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("membership_statuses ms", "m.membership_status", "ms.membership_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("statuses st", "c.status", "st.status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("contact_groups cg", "c.contact_group", "cg.contact_group", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      If mvParameters.Exists("EmailAddress") Then
        vWhereFields.Add("cm.device", vEmailDeviceCode, CDBField.FieldWhereOperators.fwoOpenBracket) 'email address
        vWhereFields.Add("cm.device#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      If mvParameters.Exists("Title") Then vWhereFields.Add("title", mvParameters("Title").Value)
      If mvParameters.Exists("Forenames") Then vWhereFields.Add("forenames", mvParameters("Forenames").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Surname") Then vWhereFields.Add("surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("DateOfBirth") Then vWhereFields.Add("date_of_birth", CDBField.FieldTypes.cftDate, mvParameters("DateOfBirth").Value)
      If mvParameters.Exists("Sex") Then vWhereFields.Add("sex", mvParameters("Sex").Value)
      If mvParameters.Exists("MemberNumber") Then vWhereFields.Add("member_number", mvParameters("MemberNumber").Value)
      If mvParameters.Exists("EmailAddress") Then vWhereFields.Add("cm.number", mvParameters("EmailAddress").Value)
      If mvParameters.Exists("NiNumber") Then vWhereFields.Add("ni_number", mvParameters("NiNumber").Value)
      If mvParameters.Exists("Town") Then vWhereFields.Add("town", mvParameters("Town").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Postcode") Then vWhereFields.Add("postcode", mvParameters("Postcode").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contacts c", vWhereFields, "c.contact_number", vAnsiJoins)
      vSQLStatement.Distinct = True
      If mvParameters.HasValue("NumberOfRows") Then vSQLStatement.MaxRows = mvParameters("NumberOfRows").IntegerValue + 1
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields, "CONTACT_NAME,CONTACT_TELEPHONE")

      If pDataTable.Rows.Count > 0 Then SetEmailAddressOfContacts(vSQLStatement, pDataTable)
    End Sub

    Private Sub GetBankTransactions(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "line_number,payers_sort_code,payers_account_number,payers_name,reference_number,amount,bt.payment_method,payment_method_desc,notes"
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("BatchDate") Then vWhereFields.Add("statement_date", CDBField.FieldTypes.cftDate, mvParameters("BatchDate").Value)
      If mvParameters.Exists("BankAccount") Then vWhereFields.Add("bank_account", mvParameters("BankAccount").Value)
      If mvParameters.Exists("TransactionCode") Then vWhereFields.Add("transaction_code", mvParameters("TransactionCode").Value)
      If mvParameters.Exists("ReconciledStatus") Then vWhereFields.Add("reconciled_status", mvParameters("ReconciledStatus").Value)
      If mvParameters.Exists("ReferenceNumber") Then
        Dim vValue As String = mvParameters("ReferenceNumber").Value
        vValue = vValue.Replace("*", "%")
        If mvParameters.Exists("RefNumberMatchMethod") AndAlso Not vValue.Contains("%") Then
          Dim vMatchMethod As String = mvParameters("RefNumberMatchMethod").Value
          If vMatchMethod.Contains("B") AndAlso Not vValue.EndsWith("%") Then
            vValue += "%"
          End If
          If vMatchMethod.Contains("E") AndAlso Not vValue.StartsWith("%") Then
            vValue = "%" + vValue
          End If
        End If
        vWhereFields.Add("reference_number", vValue, CDBField.FieldWhereOperators.fwoLike)
      End If
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("payment_methods pm", "bt.payment_method", "pm.payment_method")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "bank_transactions bt", vWhereFields, "transaction_date,payers_sort_code,payers_account_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetWebMembershipTypes(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "membership_type,membership_type_desc,long_description"
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("web_publish", "Y")
      vWhereFields.Add("allow_as_first_type", "Y")
      vWhereFields.Add("members_per_order", CDBField.FieldTypes.cftInteger, 1)
      Dim vAnsiJoins As New AnsiJoins()
      If mvParameters.ContainsKey("LookupGroup") Then
        vAnsiJoins.Add("lookup_group_details lgd", "mt.membership_type", "lgd.lookup_item")
        vAnsiJoins.Add("lookup_groups lg", "lg.lookup_group", "lgd.lookup_group")
        vWhereFields.Add("lg.lookup_group", mvParameters("LookupGroup").Value)
      End If
      Dim vContactNumber As Integer
      If mvParameters.Exists("UserContactNumber") AndAlso mvParameters.HasValue("UserContactNumber") Then
        'get contact number of the logged in user
        vContactNumber = mvParameters("UserContactNumber").IntegerValue
      End If

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipTypeCategories) AndAlso vContactNumber > 0 Then
        'changes in here should also be made in DataImportCMT(initCMTrow) and DataImportPaymentPlans(ProcessRow) as these both use the same sql as this for membershipTypeCategories restrictions.
        'we have the correct parameters to check which member types to limit from membership_type_categories table
        Dim vContact As New Contact(mvEnv)
        vContact.Init(vContactNumber)
        'part one of the three part union
        Dim vWhereFields1 As New CDBFields
        Dim vAnsiJoins1 As New AnsiJoins
        vAnsiJoins1.AddLeftOuterJoin("membership_type_categories mtc", "mt.membership_type", "mtc.membership_type")
        vWhereFields1.Add("mtc.membership_type", "")
        Dim vMemberSqlStatement As New SQLStatement(mvEnv.Connection, "mt.membership_type", "membership_types mt", vWhereFields1, "", vAnsiJoins1)
        'part two of the three part union
        vAnsiJoins1 = New AnsiJoins
        vWhereFields1 = New CDBFields
        Dim vEntity As String
        If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vWhereFields1.Add("cc.organisation_number", CDBField.FieldTypes.cftInteger, vContact.OrganisationNumber.ToString)
          vEntity = "organisation"
        Else
          vWhereFields1.Add("cc.contact_number", CDBField.FieldTypes.cftInteger, vContact.ContactNumber.ToString)
          vEntity = "contact"
        End If
        vWhereFields1.Add("cc.valid_from", CDBField.FieldTypes.cftDate, Date.Today.ToShortDateString, CDBField.FieldWhereOperators.fwoLessThanEqual)
        vWhereFields1.Add("cc.valid_to", CDBField.FieldTypes.cftDate, Date.Today.ToShortDateString, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        vAnsiJoins1.Add(New AnsiJoin(vEntity & "_categories cc", "mtc.activity", "cc.activity", "mtc.activity_value", "cc.activity_value", AnsiJoin.AnsiJoinTypes.InnerJoin))
        Dim vMemberSqlStatement2 As New SQLStatement(mvEnv.Connection, "mtc.membership_type", "membership_type_categories mtc", vWhereFields1, "", vAnsiJoins1)
        vMemberSqlStatement2.AddUnion(vMemberSqlStatement)
        'part three of the three part union
        vAnsiJoins1 = New AnsiJoins
        vAnsiJoins1.Add(New AnsiJoin(vEntity & "_categories cc", "mtc.activity", "cc.activity", AnsiJoin.AnsiJoinTypes.InnerJoin))
        Dim vWhereFields2 As New CDBFields
        vWhereFields2.Clone(vWhereFields1)
        vWhereFields2.Add("mtc.activity_value", "")
        Dim vMemberSqlStatement1 As New SQLStatement(mvEnv.Connection, "mtc.membership_type", "membership_type_categories mtc", vWhereFields2, "", vAnsiJoins1)
        vMemberSqlStatement1.AddUnion(vMemberSqlStatement2)
        'add these sql statements to the original sql statement in the where clause
        vWhereFields.Add("mt.membership_type", CDBField.FieldTypes.cftInteger, vMemberSqlStatement1.SQL, CDBField.FieldWhereOperators.fwoIn)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "membership_types mt", vWhereFields, "mt.membership_type", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetEntityAlerts(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEntityAlerts) Then
        Dim vFields As String = "entity_alert_number,entity_alert_desc,entity_alert_message,sequence_number,show_as_dialog,rgb_value,email_address,created_by,created_on,amended_by,amended_on"
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("entity_item_number", CDBField.FieldTypes.cftInteger, mvContact.ContactNumber)
        If mvParameters.Exists("EntityAlertNumber") Then vWhereFields.Add("entity_alert_number", CDBField.FieldTypes.cftInteger, mvParameters("EntityAlertNumber").IntegerValue)
        If mvEnv.User.AccessLevel <> CDBUser.UserAccessLevel.ualDatabaseAdministrator Then
          vWhereFields.Add("created_by", mvEnv.User.Logname)
        End If
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "entity_alerts", vWhereFields, "sequence_number")
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      End If
    End Sub

    Private Sub GetSystemModuleUsers(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "start_time,logname,named_user,active,last_updated_on,access_count,refused_access,build_number"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("smu.module ", mvParameters("Module").Value, CDBField.FieldWhereOperators.fwoInOrEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "sys_module_users smu ", vWhereFields, "named_user DESC,active DESC,logname,last_updated_on DESC")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetConfig(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vConfigName = mvParameters("ConfigName").Value
      vWhereFields.Add("config_name", vConfigName)
      Dim vFields As String = "config_name,config_value,client"
      Dim vScope As Config.ConfigNameScope = mvEnv.GetConfigScopeLevel(vConfigName)
      If (vScope And Config.ConfigNameScope.Department) > 0 Then
        vFields += ",department"
      End If
      If (vScope And Config.ConfigNameScope.User) > 0 Then
        vFields += ",logname"
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "config", vWhereFields, "client,department,logname")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetJobSchedule(ByVal pDataTable As CDBDataTable)
      Dim vIndex As Integer
      Dim vFields As String = "job_number,job_desc,due_date,job_processor,run_date,end_date,job_frequency,job_status,submitted_by,submitted_on,notify_status,update_job_parameter_dates,error_status,command_line"
      Dim vWhereFields As New CDBFields()
      Dim vOrderBy As String = "job_status DESC,run_date DESC, job_number DESC"
      If mvParameters.ContainsKey("SubmittedBy") Then vWhereFields.Add("submitted_by", CDBField.FieldTypes.cftCharacter, mvParameters("SubmittedBy").Value, CDBField.FieldWhereOperators.fwoEqual)
      If mvParameters.ContainsKey("FromDate") Then vWhereFields.Add("due_date", CDBField.FieldTypes.cftDate, mvParameters("FromDate").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If mvParameters.ContainsKey("ToDate") Then vWhereFields.Add("js.due_Date", CDBField.FieldTypes.cftDate, mvParameters("ToDate").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
      If mvParameters.ContainsKey("JobStatus") Then
        Dim vJobStatus() As String = mvParameters("JobStatus").Value.Split(","c)
        vWhereFields.Add("job_status", CDBField.FieldTypes.cftCharacter, vJobStatus(0), CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("job_status#2", CDBField.FieldTypes.cftCharacter, vJobStatus(1), CDBField.FieldWhereOperators.fwoNotEqual)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "job_schedule js", vWhereFields, vOrderBy)
      If mvParameters.ContainsKey("MaxRows") Then vSQLStatement.MaxRows = IntegerValue(mvParameters("MaxRows").Value)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      With pDataTable
        For vIndex = 0 To .Rows.Count - 1
          With .Rows(vIndex)
            Select Case .Item("Frequency")
              Case "O"
                .Item("Frequency") = "Once"
              Case "D"
                .Item("Frequency") = "Daily"
              Case "W"
                .Item("Frequency") = "Weekly"
              Case "M"
                .Item("Frequency") = "Monthly"
              Case "H"
                .Item("Frequency") = "Hourly"
            End Select

            Select Case .Item("Status")
              Case "S"
                .Item("Status") = "Started"
              Case "C"
                .Item("Status") = "Completed"
              Case "W"
                .Item("Status") = "Waiting"
              Case "R"
                .Item("Status") = "Running"
              Case "H"
                .Item("Status") = "On Hold"
            End Select

            Select Case .Item("Notify")
              Case "E"
                .Item("Notify") = "Started"
              Case "C"
                .Item("Notify") = "Completed"
              Case "B"
                .Item("Notify") = "Both"
              Case "N"
                .Item("Notify") = "None"
            End Select

            Select Case .Item("UpdateDates")
              Case "A"
                .Item("UpdateDates") = "Update All Dates"
              Case Else   '"N"
                .Item("UpdateDates") = "None"
            End Select
          End With

        Next
      End With
    End Sub

    Private Sub GetJobProcessors(ByVal pDataTable As CDBDataTable)
      Dim vIndex As Integer
      Dim vFields As String = "job_processor,started,status,max_concurrent_jobs,polling_interval,last_polled"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "job_processors", New CDBFields(), "status, last_polled")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      'Add new columns to hold the current status of the processor 
      pDataTable.AddColumn("Polling", CDBField.FieldTypes.cftCharacter)
      pDataTable.AddColumn("Active", CDBField.FieldTypes.cftCharacter)
      'Changing the field type as we are replacing the values with a description
      pDataTable.Columns("PollingInterval").FieldType = CDBField.FieldTypes.cftCharacter
      pDataTable.Columns("MaxConcurrentJobs").FieldType = CDBField.FieldTypes.cftCharacter

      Dim vCheck As Boolean
      Dim vInActive As Boolean
      Dim vNotPolling As Boolean
      Dim vInterval As Integer
      Dim vLastPolled As Date

      With pDataTable
        For vIndex = 0 To .Rows.Count - 1
          vCheck = False
          vInActive = False
          vNotPolling = False
          vInterval = 0
          vLastPolled = Date.MinValue

          With .Rows(vIndex)
            Select Case .Item("Status")
              Case "C"
                .Item("Status") = "Public jobs"
                vCheck = True
              Case "E"
                .Item("Status") = "Private jobs"
                vCheck = True
              Case "N"
                .Item("Status") = "Not Processing"
                vInActive = True
              Case "X"
                .Item("Status") = "Not Active"
                vInActive = True
            End Select

            vInterval = IntegerValue(.Item("PollingInterval"))
            vLastPolled = CDate(.Item("LastPolled"))

            If vCheck Then
              'Check for Job Processors that appear to be dead and flag to the user...
              If vLastPolled.AddMinutes(vInterval * 3) < Now Then
                'Looks like this Job Processor has crashed and burned
                vNotPolling = True
                .Item("Status") = "NO LONGER POLLING"
              End If
            End If

            .Item("Polling") = IIf(vNotPolling, "N", "Y").ToString
            .Item("Active") = IIf(vInActive, "N", "Y").ToString

            .Item("PollingInterval") = .Item("PollingInterval") & " minute" & (IIf(DoubleValue(.Item("PollingInterval")) > 1, "s", "")).ToString

            If .Item("MaxConcurrentJobs") = "0" OrElse .Item("MaxConcurrentJobs") = "NULL" Then
              .Item("MaxConcurrentJobs") = "No Limit"
            Else
              .Item("MaxConcurrentJobs") = .Item("MaxConcurrentJobs") & " job" & (IIf(DoubleValue(.Item("MaxConcurrentJobs")) > 1, "s", "")).ToString
            End If
          End With
        Next
      End With
    End Sub

    Private Sub GetContactAlerts(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbContactAlerts) Then
        Dim vContactAlert As New ContactAlert(mvEnv)
        vContactAlert.GetContactAlerts(pDataTable, mvParameters("ContactNumber").LongValue)
      End If
    End Sub

    Private Sub GetActionContactLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("C")
      Dim vContact As New Contact(mvEnv)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "t1.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("contact_groups cg", "c.contact_group", "cg.contact_group")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("t1.action_number", mvParameters("ActionNumber").LongValue)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("cg.client", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("cg.client#2", CDBField.FieldTypes.cftCharacter, mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "type," & vContact.GetRecordSetFieldsName & ",notified,'C' AS entity_type, " & mvEnv.Connection.DBIsNull("cg.name", "'" & vEntityDesc & "'") & " AS name", "contact_actions t1", vWhereFields, "surname, initials", vAnsiJoins)
      Dim vAddItems As String = ""
      If pAddType Then vAddItems = "CONTACT_TYPE_1,ACTION_LINK_TYPE"
      vAddItems += ",entity_type,name"
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "type,contact_number,CONTACT_NAME,notified", vAddItems)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified", False, True)
      Next
    End Sub

    Private Sub GetActionDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      GetActionDocumentLinks(pDataTable, pAddType, False)
    End Sub

    Private Sub GetActionDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean, pIncludeEmail As Boolean)
      Dim vCommsLog As New CommunicationsLog(mvEnv)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("t1.action_number", mvParameters("ActionNumber").LongValue)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("communications_log cl", "t1.document_number", "cl.communications_log_number")
      If Not pIncludeEmail Then
        vAnsiJoins.Add("packages pk", "cl.package", "pk.package", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("pk.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
        vAnsiJoins.Add("document_types dt", "cl.document_type", "dt.document_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("dt.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If
      Dim vSQL As SQLStatement
      Dim vAddAttrs As String = String.Format(",'D' AS entity_type, '{0}' AS entity_type_desc", GetLinkEntityTypeDescription("D"))
      If pAddType Then
        vSQL = New SQLStatement(mvEnv.Connection, "'R' AS type," & vCommsLog.GetRecordSetFields & vAddAttrs, "document_actions t1", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, "type,communications_log_number,DOCUMENT_NAME,,,ACTION_LINK_TYPE,entity_type,entity_type_desc")
      Else
        vSQL = New SQLStatement(mvEnv.Connection, "cl.communications_log_number, our_reference" & vAddAttrs, "document_actions t1", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub

    Private Sub GetActionLinkEMailAddresses(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "t1.contact_number", "c.contact_number")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("t1.action_number", mvParameters("ActionNumber").LongValue)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "type," & vContact.GetRecordSetFieldsName & ",notified", "contact_actions t1", vWhereFields, "surname, initials", vAnsiJoins)
      Dim vAddItems As String = ""
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "type,contact_number,CONTACT_NAME,notified,CONTACT_TYPE_1,ACTION_LINK_TYPE,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified", False, True)
        vRow.Item("EMailAddress") = Contact.NotificationEMailAddress(mvEnv, IntegerValue(vRow.Item("ContactNumber")))
      Next
    End Sub

    Private Sub GetActionFundraisingLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vFA As New FundraisingAction(mvEnv)
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("fa.action_number", mvParameters("ActionNumber").LongValue)
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("fundraising_requests fr", "fa.fundraising_request_number", "fr.fundraising_request_number")
        Dim vSQL As SQLStatement
        vSQL = New SQLStatement(mvEnv.Connection, "'R' AS type," & vFA.GetRecordSetFields & ", 'F' AS entity_type, '" & GetLinkEntityTypeDescription("F") & "' AS entity_type_desc", "fundraising_actions fa", vWhereFields, "", vAnsiJoins)
        'BR19023 Added FUNDRAISING_REQUEST as parameter to make hyperlink work
        pDataTable.FillFromSQL(mvEnv, vSQL, "type,fundraising_request_number,FUND_LINK_NAME,,FUNDRAISING_REQUEST,ACTION_LINK_TYPE,entity_type,entity_type_desc")
      End If
    End Sub

    Private Sub GetActionExamLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      pDataTable.FillFromSQL(mvEnv,
                             New SQLStatement(mvEnv.Connection,
                                              ("'N' AS contact_type," & (New ExamCentreAction(mvEnv)).GetRecordSetFields) & ", ec.exam_centre_description,'N' AS entity_type, '" & GetLinkEntityTypeDescription("N") & "' AS entity_type_desc",
                                              "action_links eca",
                                               New CDBFields({New CDBField("eca.action_number", mvParameters("ActionNumber").LongValue)}),
                                               "",
                                               New AnsiJoins({New AnsiJoin("exam_centres ec",
                                                                           "eca.exam_centre_id",
                                                                           "ec.exam_centre_id")})),
                             "type,exam_centre_id,exam_centre_description,,contact_type,ACTION_LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetActionWorkstreamLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vFields As String = "'R' as type,al.action_number,ws.workstream_id,ws.workstream_desc"
      vFields &= ",'W' AS entity_type, wg.workstream_group_desc"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("workstreams ws", "al.workstream_id", "ws.workstream_id")
      vAnsiJoin.Add("workstream_groups wg", "ws.workstream_group", "wg.workstream_group")

      vWhereFields.Add("al.action_number", mvParameters("ActionNumber").LongValue)
      Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "action_links al", vWhereFields, "ws.workstream_desc", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSqlQuery, "type,workstream_id,workstream_desc,,WORKSTREAM,ACTION_LINK_TYPE,entity_type,workstream_group_desc")
    End Sub

    Private Sub GetActionEventLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vFields As String = "'R' as type, a.action_number, e.event_number, e.event_desc"
      vFields += ",'E' AS entity_type," & mvEnv.Connection.DBIsNull("eg.name", "'" & GetLinkEntityTypeDescription("E") & "'") & " AS entity_type_desc"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("events e", "e.master_action", "a.master_action")
      vAnsiJoin.AddLeftOuterJoin("event_groups eg", "e.event_group", "eg.event_group")

      vWhereFields.Add("a.action_number", mvParameters("ActionNumber").LongValue)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "actions a", vWhereFields, "e.event_desc", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSQL, "type,event_number,event_desc,,EVENT_TYPE,ACTION_LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetActionLegacyLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vFields As String = "'R' as type, actions.action_number, contact_legacies.legacy_number, '" + GetLinkEntityTypeDescription("L") + "  - ' + legator.label_name legacy_desc"
      vFields += ",'L' AS entity_type, '" & GetLinkEntityTypeDescription("L") & "' AS entity_type_desc"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("contact_legacies", "contact_legacies.master_action", "actions.master_action")
      vAnsiJoin.Add("contacts legator", "legator.contact_number", "contact_legacies.contact_number")

      vWhereFields.Add("actions.action_number", mvParameters("ActionNumber").LongValue)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "actions", vWhereFields, "legator.label_name", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSQL, "type,legacy_number,legacy_desc,,LEGACY_TYPE,ACTION_LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetActionMeetingLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vFields As String = "'R' as type, actions.action_number, meetings.meeting_number, meetings.meeting_desc"
      vFields += ",'M' AS entity_type, '" & GetLinkEntityTypeDescription("M") & "' AS entity_type_desc"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("meetings", "meetings.master_action", "actions.master_action")

      vWhereFields.Add("actions.action_number", mvParameters("ActionNumber").LongValue)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "actions", vWhereFields, "meetings.meeting_date DESC", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSQL, "type,meeting_number,meeting_desc,,MEETING_TYPE,ACTION_LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetActionPositionLinks(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "al.type, cp.contact_position_number, cp.position"
      vFields &= ", 'P' AS entity_type, '" & GetLinkEntityTypeDescription("P") & "' AS entity_type_desc,"
      Dim vContact As New Contact(mvEnv)
      vFields &= vContact.GetRecordSetFieldsName

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contact_positions cp", "al.contact_position_number", "cp.contact_position_number")
      vAnsiJoins.Add("contacts c", "cp.contact_number", "c.contact_number")

      Dim vWhereFields As New CDBFields(New CDBField("al.action_number", mvParameters("ActionNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "action_links al", vWhereFields, "cp.position", vAnsiJoins)

      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "type,contact_position_number,POSITION_LINK_DESC,,POSITION_LINK_TYPE,ACTION_LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetActionLinks(ByVal pDataTable As CDBDataTable)
      GetActionContactLinks(pDataTable, True)
      GetActionOrganisationLinks(pDataTable, True)
      GetActionDocumentLinks(pDataTable, True, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
      GetActionFundraisingLinks(pDataTable, True)
      GetActionExamLinks(pDataTable, True)
      GetActionWorkstreamLinks(pDataTable, True)
      GetActionEventLinks(pDataTable, True)
      GetActionLegacyLinks(pDataTable, True)
      GetActionMeetingLinks(pDataTable, True)
      GetActionPositionLinks(pDataTable)
      pDataTable.ReOrderRowsByColumn(("LinkType"))
    End Sub

    Private Sub GetActionOrganisationLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("O")
      Dim vAttrs As String = "type,o.organisation_number,o.name,'O' AS entity_type, " & mvEnv.Connection.DBIsNull("og.name", "'" & vEntityDesc & "'") & " AS entity_type_desc"
      Dim vCols As String = "type,organisation_number,name,"

      If pAddType Then vCols &= ",ORGANISATION_TYPE_1,ACTION_LINK_TYPE"
      vCols &= ",entity_type,entity_type_desc"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("organisations o", "t1.organisation_number", " o.organisation_number")
      vAnsiJoins.AddLeftOuterJoin("organisation_groups og", "o.organisation_group", "og.organisation_group")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("t1.action_number", mvParameters("ActionNumber").LongValue)
      vWhereFields.Add("og.client", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("og.client#2", CDBField.FieldTypes.cftCharacter, mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "organisation_actions t1", vWhereFields, "o.name", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
    End Sub
    Private Sub GetActionOutline(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vSQL As String = "SELECT action_desc,action_number,master_action,action_level,sequence_number,created_by,department,a.action_priority,a.action_status,action_priority_desc,action_status_desc"
      vSQL = vSQL & " FROM actions a, action_priorities ap, action_statuses acs, document_classes dc, users u WHERE master_action = " & mvParameters("MasterAction").LongValue
      vSQL = vSQL & " AND a.action_priority = ap.action_priority AND a.action_status = acs.action_status AND a.document_class = dc.document_class AND a.created_by = u.logname ORDER BY sequence_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetPriorActions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vSQL As String = "SELECT a.action_number,a.action_desc,sequence_number"
      vSQL = vSQL & " FROM action_dependencies ad, actions a WHERE ad.action_number = " & mvParameters("ActionNumber").LongValue
      vSQL = vSQL & " AND a.action_number = ad.prior_action"
      vSQL = vSQL & " ORDER BY sequence_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetActionSubjects(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "acs.topic,topic_desc,acs.sub_topic,sub_topic_desc,acs.notes,acs.amended_on,acs.amended_by"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("topics t", "acs.topic", "t.topic")
      vAnsiJoins.Add("sub_topics st", "t.topic", "st.topic", "acs.sub_topic", "st.sub_topic")
      Dim vWhereFields = New CDBFields(New CDBField("acs.action_number", mvParameters("ActionNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "action_subjects acs", vWhereFields, "topic_desc, sub_topic_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "topic,topic_desc,sub_topic,sub_topic_desc,notes,,amended_on,amended_by")
    End Sub
    Private Sub GetActivitiesDataSheet(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vSQL As String
      If mvParameters("UsageCode").Value = "L" Then   'Activity Entry Type is aetLegacy
        vSQL = "SELECT a.activity,activity_desc,av.activity_value,activity_value_desc,'N' AS mandatory,'Y' AS quantity_required,'Y' AS multiple_values,a.contact_group"
        vSQL = vSQL & " FROM activities a,activity_values av,activity_users au,activity_value_users avu "
        vSQL = vSQL & " WHERE a.activity = '" & mvParameters("Activity").Value & "' AND a.activity = av.activity AND a.activity = au.activity AND au.department = '" & mvEnv.User.Department & "'"
        vSQL = vSQL & " AND a.activity = avu.activity AND av.activity_value = avu.activity_value AND avu.department = '" & mvEnv.User.Department & "'"
        vSQL = vSQL & " ORDER BY high_profile,profile_rating,activity_desc,activity_value_desc"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "a.activity,activity_desc,av.activity_value,activity_value_desc,mandatory,quantity_required,multiple_values,a.contact_group,")
      Else
        Dim vBaseSQL As String = "SELECT a.activity,activity_desc,av.activity_value,activity_value_desc,mandatory,quantity_required,multiple_values,a.contact_group,a.duration_months AS aduration_months,a.duration_days AS aduration_days,av.duration_months,av.duration_days,a.is_historic As is_activity_historic,av.is_historic As is_activity_value_historic,%4source,sequence_number FROM %1 WHERE %2"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbActivityDurationDays) = False Then vBaseSQL = vBaseSQL.Replace("a.duration_months AS aduration_months,a.duration_days AS aduration_days,av.duration_months,av.duration_days", "'' AS aduration_months,'' AS aduration_days,'' AS avduration_months,'' AS avduration_days")
        If mvParameters.Exists("ActivityGroupCode") Then vBaseSQL = vBaseSQL & "ag.activity_group = '" & mvParameters("ActivityGroupCode").Value & "' AND "
        vBaseSQL = vBaseSQL & "ag.usage_code = '" & mvParameters("UsageCode").Value & "'%3"  '& vSourceSQL
        vBaseSQL = vBaseSQL & " AND ag.activity_group = agd.activity_group"
        vBaseSQL = vBaseSQL & " AND agd.activity = a.activity"
        vBaseSQL = vBaseSQL & " AND au.activity = a.activity"
        vBaseSQL = vBaseSQL & " AND au.department = '" & mvEnv.User.Department & "'"
        vBaseSQL = vBaseSQL & " AND (a.contact_group IS NULL OR a.contact_group IN (" '& vGroups & "))"
        'Assumption is that either the Contact Group or the Org Group parameter will be provided, maybe both
        If mvParameters.Exists("ContactGroupCode") Then vBaseSQL = vBaseSQL & "'" & mvParameters("ContactGroupCode").Value & "'"
        If mvParameters.Exists("OrganisationGroupCode") Then
          If Right$(vBaseSQL, 1) <> "(" Then vBaseSQL = vBaseSQL & ","
          vBaseSQL = vBaseSQL & "'" & mvParameters("OrganisationGroupCode").Value & "'"
        End If
        vBaseSQL = vBaseSQL & "))"
        vBaseSQL = vBaseSQL & " AND a.activity = av.activity"
        vBaseSQL = vBaseSQL & " AND a.activity = avu.activity"
        vBaseSQL = vBaseSQL & " AND av.activity_value = avu.activity_value"
        vBaseSQL = vBaseSQL & " AND avu.department = '" & mvEnv.User.Department & "'"
        If mvParameters("UsageCode").Value = "E" Then 'for now we will only filter by department for 'E' (Contact Entry)
          vBaseSQL = vBaseSQL & " AND (ag.department = '" & mvEnv.User.Department & "' OR ag.department IS NULL)"
        End If
        If mvParameters.Exists("ExcludeHistoricActivities") Then
          vBaseSQL = vBaseSQL & " AND a.is_historic <> 'Y'"
        End If
        If mvParameters.Exists("ExcludeHistoricActivityValues") Then
          vBaseSQL = vBaseSQL & " AND av.is_historic <> 'Y'"
        End If
        vBaseSQL = vBaseSQL & " ORDER BY sequence_number,activity_desc,activity_value_desc"
        Dim vSourceSQL As String = ""
        If mvParameters.Exists("SourceCode") Then
          vSourceSQL = " AND (ag.source = '" & mvParameters("SourceCode").Value & "' OR ag.source IS NULL)"
          If mvParameters("UsageCode").Value = "E" Then vSourceSQL = vSourceSQL & " AND ag.campaign IS NULL"
        End If
        vSQL = Replace(vBaseSQL, "%1", "activity_groups ag,activity_group_details agd,activities a,activity_users au,activity_values av,activity_value_users avu")
        vSQL = Replace(vSQL, "%2", "")
        vSQL = Replace(vSQL, "%3", vSourceSQL)
        vSQL = Replace(vSQL, "%4", "")
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
        If pDataTable.Rows.Count = 0 And mvParameters.Exists("SourceCode") Then
          vSQL = Replace(vBaseSQL, "SELECT", "SELECT DISTINCT")  'Could have multiple segments so make it distinct
          vSQL = Replace(vSQL, "%1", "segments s,activity_groups ag,activity_group_details agd,activities a,activity_users au,activity_values av,activity_value_users avu")
          vSQL = Replace(vSQL, "%2", "s.source = '" & mvParameters("SourceCode").Value & "' AND s.campaign = ag.campaign AND s.appeal = ag.appeal AND ")
          vSQL = Replace(vSQL, "%3", "")
          vSQL = Replace(vSQL, "%4", "s.")
          pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
        End If
      End If
    End Sub
    Private Sub GetAppealBudgetDetails(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetAppealBudgetDetails")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "appeal_budget_details_number,apd.appeal_budget_number,segment,reason_for_despatch,forecast_units,budgeted_costs,budgeted_income,budget_period,campaign,appeal"
      With vWhereFields
        If mvParameters.Exists("AppealBudgetDetailsNumber") Then .Add("appeal_budget_details_number", mvParameters.Item("AppealBudgetDetailsNumber").LongValue)
        If mvParameters.Exists("AppealBudgetNumber") Then .Add("apd.appeal_budget_number", mvParameters.Item("AppealBudgetNumber").LongValue)
        If mvParameters.Exists("Segment") Then .Add("segment", mvParameters.Item("Segment").Value)
        .AddJoin("ab.appeal_budget_number", "apd.appeal_budget_number")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM appeal_budget_details apd, appeal_budgets ab WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY segment, reason_for_despatch"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",")
      GetDescriptions(pDataTable, "ReasonForDespatch")
    End Sub
    Private Sub GetAppealBudgets(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "appeal_budget_number,campaign,appeal,budget_period,period_start_date,period_end_date,period_percentage"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("AppealBudgetNumber") Then .Add("appeal_budget_number", mvParameters.Item("AppealBudgetNumber").LongValue)
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM appeal_budgets WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetAppealResources(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "appeal_resource_number,cp.campaign,cp.campaign_desc,ap.appeal,ap.appeal_desc,p.product,p.product_desc,total_quantity,quantity_remaining"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("ar.campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("ar.appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("AppealResourceNumber") Then .Add("appeal_resource_number", mvParameters.Item("AppealResourceNumber").LongValue)
        .AddJoin("cp.campaign", "ar.campaign")
        .AddJoin("ap.campaign", "ar.campaign")
        .AddJoin("ap.appeal", "ar.appeal")
        .AddJoin("p.product", "ar.product")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM appeal_resources ar, campaigns cp, appeals ap, products p WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub

    Private Sub GetAppealCollections(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "collection_number,collection,collection_desc,product,rate,source,bank_account,collection_type"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("CollectionNumber") Then .Add("collection_number", mvParameters.Item("CollectionNumber").LongValue)
        If mvParameters.Exists("AppealCollectionNumber") Then .Add("collection_number", mvParameters.Item("AppealCollectionNumber").LongValue)
        If mvParameters.Exists("BankAccount") Then .Add("bank_account", mvParameters.Item("BankAccount").Value)
      End With
      If vWhereFields.Count > 0 Then
        Dim vSQL As String = "SELECT " & vAttrs & " FROM appeal_collections WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      End If
    End Sub

    Private Sub GetAppealTypes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "lookup_code,lookup_desc"
      With vWhereFields
        .Add("table_name", "appeals")
        .Add("attribute_name", "appeal_type")
      End With
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "maintenance_lookup", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",")
      Dim vCreateRights As String = mvEnv.GetConfig("ma_create_appeal_types")
      If vCreateRights.Length = 0 Then vCreateRights = "S"
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vCreateRights.Contains(vRow.Item("AppealType")) Then vRow.Item("Access") = "C"
      Next
    End Sub
    Private Sub GetBACSAmendments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vAttrs As String = "bacs_amendment_number,direct_debit_number,old_bank_details_number,new_bank_details_number,bacs_record_type_desc,bacs_advice_reason_desc,effective_date,advice_reference,payers_name,payers_sort_code,payers_account_number,advice_due_date,bacs_payment_frequency,amount_of_payment,payers_new_name,payers_new_sort_code,payers_new_account_number,new_due_date,new_bacs_payment_frequency,new_amount_of_payment,last_payment_date,building_society_roll_no_1,building_society_roll_no_2,building_society_roll_no_3,bacs_transaction_code,originators_sequence_no,originators_sort_code,originators_account_number,user_number,bacs_notes,notes,ba.amended_by,ba.amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then
        vAttrs = vAttrs & ",payers_iban_number,payers_bic_code,originators_iban_number,originators_bic_Code,end_to_end_id"
      End If
      Dim vSQL As String = "SELECT " & vAttrs & " FROM bacs_amendments ba, bacs_record_types bat, bacs_advice_reasons bar"
      If mvParameters.Exists("DirectDebitNumber") Then
        vSQL = vSQL & " WHERE ba.direct_debit_number = " & mvParameters("DirectDebitNumber").LongValue
      Else
        vSQL = vSQL & " WHERE ((ba.old_bank_details_number = " & mvParameters("BankDetailsNumber").LongValue & ")"
        vSQL = vSQL & " OR (ba.new_bank_details_number = " & mvParameters("BankDetailsNumber").LongValue & "))"
      End If
      vSQL = vSQL & " AND bat.bacs_record_type = ba.bacs_record_type AND bar.bacs_advice_reason = ba.bacs_advice_reason ORDER BY bacs_amendment_number"

      vAttrs = Replace$(vAttrs, "payers_sort_code", "PAYERS_SORT_CODE")
      vAttrs = Replace$(vAttrs, "payers_new_sort_code", "PAYERS_NEW_SORT_CODE")
      vAttrs = Replace$(vAttrs, "originators_sort_code", "ORIGINATORS_SORT_CODE")
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetBrowserContactPositions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vLimitByRoles As Boolean = mvEnv.GetConfigOption("cd_org_browser_limit_by_roles", False)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName & ",contact_group"
      Dim vSQL As String = "SELECT DISTINCT " & vConAttrs & ",position FROM "
      If vLimitByRoles Then vSQL = vSQL & "contact_roles cr, "
      vSQL = vSQL & "contact_positions cp,contacts c WHERE "
      If vLimitByRoles Then
        vSQL = vSQL & "cr.organisation_number = " & mvContact.ContactNumber & " AND cr.organisation_number = cp.organisation_number AND cr.contact_number = cp.contact_number"
      Else
        vSQL = vSQL & "cp.organisation_number = " & mvContact.ContactNumber
      End If
      vSQL = vSQL & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND c.contact_number = cp.contact_number AND contact_type <> 'O'"
      vSQL = vSQL & " ORDER BY surname, initials"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "contact_number,contact_group,CONTACT_NAME,position")
    End Sub
    Private Sub GetCampaignAppeals(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As CDBFields = GetCampaignWhereFields()
      If mvParameters.Exists("Appeal") Then vWhereFields.Add("a.appeal", mvParameters("Appeal").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.AddJoin("a.campaign", "c.campaign")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "a.campaign,a.appeal,a.appeal_desc,a.thank_you_letter", "campaigns c,appeals a", vWhereFields, "a.campaign, a.appeal")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetCampaignCollections(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As CDBFields = GetCampaignWhereFields()
      If mvParameters.Exists("Appeal") Then vWhereFields.Add("a.appeal", mvParameters("Appeal").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("AppealType") Then vWhereFields.Add("a.appeal_type", mvParameters("AppealType").Value)
      If mvParameters.Exists("Collection") Then vWhereFields.Add("ac.collection", mvParameters("Collection").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add("ac.collection_number", mvParameters("CollectionNumber").LongValue)
      vWhereFields.AddJoin("a.campaign", "c.campaign")
      vWhereFields.AddJoin("ac.campaign", "a.campaign")
      vWhereFields.AddJoin("ac.appeal", "a.appeal")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "ac.campaign,campaign_desc,start_date,ac.appeal,appeal_desc,ac.collection_number,ac.collection,ac.collection_desc", "campaigns c, appeals a, appeal_collections ac", vWhereFields, "ac.campaign, ac.appeal, ac.collection")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetCampaignInfo(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As CDBFields = GetCampaignWhereFields()
      Dim vInClause As String = GetCampaignRestriction(GetAppealWhereFields)
      Dim vAppealsOnly As Boolean
      If mvParameters.Exists("Appeal") Then
        Dim vAppeal As String = mvParameters("Appeal").Value
        If vAppeal.Length > 0 Then
          vWhereFields.Add("a.appeal", vAppeal, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        Else
          vAppealsOnly = True
        End If
      End If
      'Note: If extra Appeals attributes are required then please add them further up within this procedure as in some situations it will use an in-clause (see comment above)
      Dim vFields As String = "c.campaign,campaign_desc,a.appeal,appeal_desc,s.segment_sequence,s.segment,s.segment_desc,segment_date,c.end_date as campaign_end_date,c.start_date as campaign_start_date,appeal_date"
      vFields &= ",a.end_date as appeal_end_date,hc.start_date as h2h_start_date,uc.start_date as unmanned_start_date,hc.end_date as h2h_end_date,uc.end_date as unmanned_end_date,collection_date,collection,collection_desc,ac.collection_number,appeal_type"
      Dim vAttrsList As String = "c.campaign,campaign_desc,a.appeal,appeal_desc,s.segment_sequence,s.segment,s.segment_desc,segment_date,campaign_end_date,campaign_start_date,appeal_date"
      vAttrsList &= ",appeal_end_date,h2h_start_date,unmanned_start_date,h2h_end_date,unmanned_end_date,collection_date,collection,collection_desc,ac.collection_number,appeal_type"
      Dim vAnsiJoins As New AnsiJoins()       'Join Appeals
      If vInClause.Length > 0 Then            'Appeals needs to join to multiple tables so use a separate SQL statement
        vAnsiJoins.AddLeftOuterJoin(" (" & vInClause & ") a", "c.campaign", "a.campaign")
      Else
        vAnsiJoins.AddLeftOuterJoin("appeals a", "c.campaign", "a.campaign")
      End If
      'Join Segments 
      Dim vAnsiJoin As AnsiJoin
      If vAppealsOnly Then
        'Force it to not find any segments
        vAnsiJoin = New AnsiJoin("segments s", "a.campaign", "s.campaign", "a.appeal", "s.amended_by", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Else
        vAnsiJoin = New AnsiJoin("segments s", "a.campaign", "s.campaign", "a.appeal", "s.appeal", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If
      vAnsiJoins.Add(vAnsiJoin)
      'Join Appeal Collections
      If vAppealsOnly Then
        'Force it to not find any collections
        vAnsiJoin = New AnsiJoin("appeal_collections ac", "a.campaign", "ac.campaign", "a.appeal", "ac.amended_by", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Else
        vAnsiJoin = New AnsiJoin("appeal_collections ac", "a.campaign", "ac.campaign", "a.appeal", "ac.appeal", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If
      vAnsiJoins.Add(vAnsiJoin)
      vAnsiJoins.AddLeftOuterJoin("manned_collections mc", "ac.collection_number", "mc.collection_number")
      vAnsiJoins.AddLeftOuterJoin("unmanned_collections uc", "ac.collection_number", "uc.collection_number")
      vAnsiJoins.AddLeftOuterJoin("h2h_collections hc", "ac.collection_number", "hc.collection_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "campaigns c", vWhereFields, "c.campaign, a.appeal, s.segment_sequence, ac.collection_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrsList)
    End Sub
    Private Sub GetCampaigns(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As CDBFields = GetCampaignWhereFields()
      Dim vInClause As String = GetCampaignRestriction(GetAppealWhereFields)
      If vInClause.Length > 0 Then vWhereFields.Add("c.campaign#2", vInClause, CDBField.FieldWhereOperators.fwoIn)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("campaign_statuses cs", "c.campaign_status", "cs.campaign_status")
      Dim vAttrs As String = "campaign,campaign_desc,start_date,end_date,manager,campaign_business_type,c.campaign_status,campaign_status_date,campaign_status_reason,notes,actual_income,actual_income_date,mark_historical,total_itemised_cost,topic"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignItemisedCosts) Then vAttrs = vAttrs.Replace(",total_itemised_cost", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing) Then vAttrs = vAttrs.Replace(",topic", ",")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "campaigns c", vWhereFields, "campaign", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub
    Private Sub GetCampaignRoles(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignRoles) Then
        Dim vWhereFields As New CDBFields
        If mvParameters.Exists("ContactCampaignRoleNumber") Then vWhereFields.Add("ccr.contact_campaign_role_number", mvParameters("ContactCampaignRoleNumber").Value)
        If mvParameters.Exists("Campaign") Then vWhereFields.Add("ccr.campaign", mvParameters("Campaign").Value)
        If mvParameters.Exists("Appeal") Then
          vWhereFields.Add("ccr.appeal", mvParameters("Appeal").Value)
        Else
          vWhereFields.Add("ccr.appeal", "")
        End If
        If mvParameters.Exists("Segment") OrElse mvParameters.Exists("CollectionNumber") Then
          If mvParameters.Exists("Segment") Then
            vWhereFields.Add("ccr.segment", mvParameters("Segment").Value)
          Else
            vWhereFields.Add("ccr.collection_number", mvParameters("CollectionNumber").Value)
          End If
        Else
          vWhereFields.Add("ccr.segment", "")
          vWhereFields.Add("ccr.collection_number", "")
        End If
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("contacts c", "ccr.contact_number", "c.contact_number")
        vAnsiJoins.Add("campaign_roles cr", "ccr.campaign_role", "cr.campaign_role")
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "ccr.contact_campaign_role_number,ccr.contact_number,c.label_name,ccr.campaign_role,cr.campaign_role_desc,ccr.amended_by,ccr.amended_on", "contact_campaign_roles ccr", vWhereFields, "c.surname, c.forenames, cr.campaign_role, ccr.amended_on", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      End If
    End Sub
    Private Sub GetCampaignSegments(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As CDBFields = GetCampaignWhereFields()
      If mvParameters.Exists("Appeal") Then vWhereFields.Add("a.appeal", mvParameters("Appeal").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Segment") Then vWhereFields.Add("s.segment", mvParameters("Segment").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.AddJoin("a.campaign", "c.campaign")
      vWhereFields.AddJoin("s.campaign", "a.campaign")
      vWhereFields.AddJoin("s.appeal", "a.appeal")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "s.campaign,campaign_desc,start_date,s.appeal,appeal_desc,s.segment_sequence,s.segment,s.segment_desc", "campaigns c, appeals a, segments s", vWhereFields, "s.campaign, s.appeal, s.segment_sequence")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetClaimedPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetClaimedPayments")
      Dim vAttrs As String = "fhd.batch_number,fhd.transaction_number,fhd.line_number,transaction_date,product_desc,dtcl.amount_claimed,net_amount,dtcl.claim_number,claim_generated_date"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vAttrs
      vSQL = vSQL & " FROM declaration_tax_claim_lines dtcl, financial_history_details fhd, products p, financial_history fh,declaration_tax_claims dtc"
      vSQL = vSQL & " WHERE dtcl.cd_number = " & mvParameters("DeclarationNumber").LongValue & " AND declaration_or_covenant_number = 'D' AND fhd.batch_number = dtcl.batch_number AND"
      vSQL = vSQL & " fhd.transaction_number = dtcl.transaction_number AND fhd.line_number = dtcl.line_number AND p.product = fhd.product"
      vSQL = vSQL & " AND fh.batch_number = fhd.batch_number AND fh.transaction_number = fhd.transaction_number AND dtc.claim_number = dtcl.claim_number"
      vSQL = vSQL & " ORDER BY dtcl.claim_number DESC, dtcl.batch_number DESC, dtcl.transaction_number DESC, dtcl.line_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, Replace$(vAttrs, "product_desc", "DISTINCT_PRODUCT_LINE"))
      'Select any reversal transactions created by the changing of the Declaration end date
      vAttrs = "dtcl.batch_number,dtcl.transaction_number,dtcl.line_number,transaction_date,product_desc,dtcl.amount_claimed,net_amount,dtcl.claim_number,claim_generated_date"
      vSQL = "SELECT /* SQLServerCSC */ " & vAttrs
      vSQL = vSQL & " FROM declaration_tax_claim_lines dtcl, batches b, financial_history_details fhd, products p, batch_transactions bt, declaration_tax_claims dtc"
      vSQL = vSQL & " WHERE dtcl.cd_number = " & mvParameters("DeclarationNumber").LongValue & " AND declaration_or_covenant_number = 'D' AND b.batch_number = dtcl.batch_number"
      vSQL = vSQL & " AND batch_type = '" & Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment) & "' AND fhd.batch_number = dtcl.batch_number"
      vSQL = vSQL & " AND fhd.transaction_number = dtcl.transaction_number AND fhd.line_number = dtcl.line_number AND p.product = fhd.product"
      vSQL = vSQL & " AND bt.batch_number = dtcl.batch_number AND bt.transaction_number = dtcl.transaction_number AND dtc.claim_number = dtcl.claim_number"
      vSQL = vSQL & " ORDER BY dtcl.claim_number DESC, dtcl.batch_number DESC, dtcl.transaction_number DESC, dtcl.line_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, Replace$(vAttrs, "product_desc", "DISTINCT_PRODUCT_LINE"))
    End Sub
    Private Sub GetCollectionBoxesForPayment(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vCollPisNumber As Integer
      If mvParameters.Exists("CollectionPisNumber") Then
        vCollPisNumber = mvParameters.Item("CollectionPisNumber").LongValue
      End If
      Dim vSQL As String = "SELECT collection_box_number, mc.contact_number, box_reference, cb.amount"
      vSQL = vSQL & " FROM appeal_collections ap INNER JOIN collection_boxes cb ON ap.collection_number = cb.collection_number"
      vSQL = vSQL & " LEFT OUTER JOIN manned_collectors mc ON cb.collection_number = mc.collection_number AND cb.collector_number = mc.collector_number"
      vSQL = vSQL & " WHERE ap.collection_number = " & mvParameters.Item("CollectionNumber").LongValue
      vSQL = vSQL & " AND ((collection_type = 'U') OR (collection_type <> 'U' AND collection_pis_number"
      vSQL = vSQL & If(vCollPisNumber > 0, " = " & vCollPisNumber, " IS NULL") & ")) ORDER BY box_reference"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), "", ",")
    End Sub
    Private Sub GetCollectionPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectionPayments")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "box_reference,cp.collection_box_number,transaction_date,cp.amount,cp.collection_payment_number,cp.collection_number,cp.collection_pis_number,pis_number,cp.batch_number,cp.transaction_number,cp.line_number,fl.contact_number,fh.contact_number  AS  fh_contact_number"
      Dim vAttrsList As String = "box_reference,cp.collection_box_number,transaction_date,cp.amount,cp.collection_payment_number,cp.collection_number,cp.collection_pis_number,pis_number,cp.batch_number,cp.transaction_number,cp.line_number,fl.contact_number,fh_contact_number"
      With vWhereFields
        If mvParameters.Exists("CollectionNumber") Then .Add("cp.collection_number", mvParameters("CollectionNumber").LongValue)
        If mvParameters.Exists("CollectionPaymentNumber") Then .Add("collection_payment_number", mvParameters("CollectionPaymentNumber").LongValue)
        If mvParameters.Exists("CollectionPISNumber") Then .Add("cp.collection_pis_number", CDBField.FieldTypes.cftCharacter, mvParameters("CollectionPISNumber").LongValue)
        If mvParameters.Exists("ContactNumber") Then .Add("fh.contact_number", CDBField.FieldTypes.cftCharacter, mvParameters("ContactNumber").LongValue)
      End With
      Dim vSQL As String = "SELECT " & vAttrs & "," & vConAttrs & " FROM financial_history fh"
      vSQL = vSQL & " INNER JOIN contacts c ON c.contact_number = fh.contact_number"
      vSQL = vSQL & " INNER JOIN collection_payments cp ON cp.batch_number = fh.batch_number AND cp.transaction_number = fh.transaction_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_pis pis ON cp.collection_pis_number = pis.collection_pis_number"
      vSQL = vSQL & " LEFT OUTER JOIN financial_links fl ON cp.batch_number = fl.batch_number AND cp.transaction_number = fl.transaction_number AND cp.line_number = fl.line_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_boxes cb ON cp.collection_box_number = cb.collection_box_number"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrsList & "," & ContactNameItems() & ",")
      GetContactNames(pDataTable, "SentOnBehalfOfContactNumber", "SentOnBehalfOfContactName")
    End Sub
    Private Sub GetCollectionPIS(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectionPIS")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "pis.collection_number,collection_pis_number,pis_number,pis.collector_number,x.contact_number,issue_date,amount,banked_by,banked_on,reconciled_on,o.name"
      If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add("pis.collection_number", mvParameters("CollectionNumber").LongValue)
      If mvParameters.Exists("CollectionPISNumber") Then vWhereFields.Add("collection_pis_number", mvParameters("CollectionPISNumber").LongValue)
      If mvParameters.Exists("CollectorNumber") Then vWhereFields.Add("collector_number", mvParameters("CollectorNumber").LongValue)
      Dim vSQL As String = "SELECT " & vConAttrs & "," & vAttrs & " FROM collection_pis pis LEFT OUTER JOIN (SELECT collector_number, hc.contact_number, " & vConAttrs & " FROM h2h_collectors hc INNER JOIN contacts c ON hc.contact_number = c.contact_number) x ON pis.collector_number = x.collector_number LEFT OUTER JOIN organisations o ON pis.banked_by = o.organisation_number WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), ContactNameItems() & "," & vAttrs)
      GetContactNames(pDataTable, "BankedBy", "BankedByContactName", "", True)
    End Sub
    Private Sub GetCollectionPoints(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectionPoints")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "cp.collection_point_number,cr.geographical_region,gr.geographical_region_desc,cp.collection_region_number,cp.collection_point,cp.collection_point_type,cp.organisation_number,o.name,cp.address_number,cp.no_of_collectors,cp.notes"
      With mvParameters
        If .Exists("CollectionPointNumber") Then vWhereFields.Add("cp.collection_point_number", .Item("CollectionPointNumber").LongValue)
        If .Exists("CollectionRegionNumber") Then vWhereFields.Add("cr.collection_region_number", .Item("CollectionRegionNumber").LongValue)
        If .Exists("CollectionNumber") Then vWhereFields.Add("cr.collection_number", .Item("CollectionNumber").LongValue)
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM collection_regions cr"
      vSQL = vSQL & " INNER JOIN geographical_regions gr ON cr.geographical_region = gr.geographical_region"
      vSQL = vSQL & " INNER JOIN collection_points cp ON cr.collection_region_number = cp.collection_region_number"
      vSQL = vSQL & " LEFT OUTER JOIN organisations o ON cp.organisation_number = o.organisation_number"
      If vWhereFields.Count > 0 Then vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs)
    End Sub
    Private Sub GetCollectionRegions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectionRegions")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "cr.collection_region_number,cr.collection_number,cr.geographical_region," & mvEnv.Connection.DBIsNull("name", "geographical_region_desc") & " AS geographical_region_desc "
      Dim vAttrsList As String = "cr.collection_region_number,cr.collection_number,cr.geographical_region,geographical_region_desc"
      If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add("cr.collection_number", mvParameters("CollectionNumber").LongValue)
      If mvParameters.Exists("CollectionRegionNumber") Then vWhereFields.Add("cr.collection_region_number", mvParameters("CollectionRegionNumber").LongValue)
      vWhereFields.Add("geographical_region_type", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCollectionsRegionType))
      'vWhereFields.AddJoin( "cr.geographical_region",  "gr.geographical_region")
      Dim vSQL As String = "SELECT " & vAttrs & " FROM collection_regions cr INNER JOIN geographical_regions gr ON cr.geographical_region = gr.geographical_region LEFT OUTER JOIN organisations o ON gr.organisation_number = o.organisation_number WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrsList)
    End Sub
    Private Sub GetCollectionResources(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectionResources")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "cr.collection_resource_number,cr.collection_number,cr.appeal_resource_number,p.product,p.product_desc,r.rate,r.rate_desc,cr.quantity,cr.despatch_on,dm.despatch_method,dm.despatch_method_desc,cr.amended_by,cr.amended_on"
      With vWhereFields
        If mvParameters.Exists("CollectionResourceNumber") Then .Add("cr.collection_resource_number", mvParameters("CollectionResourceNumber").LongValue)
        If mvParameters.Exists("CollectionNumber") Then .Add("cr.collection_number", mvParameters("CollectionNumber").LongValue)
        .AddJoin("ar.appeal_resource_number", "cr.appeal_resource_number")
        .AddJoin("p.product", "ar.product")
        .AddJoin("r.rate", "cr.rate")
        .AddJoin("r.product", "ar.product")
        .AddJoin("dm.despatch_method", "cr.despatch_method")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM collection_resources cr, appeal_resources ar, products p, rates r, despatch_methods dm WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetCollectorShifts(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetCollectorShifts")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "cs.collector_shift_number,cs.collector_number,cr.collection_region_number,cr.geographical_region," & mvEnv.Connection.DBIsNull("name", "geographical_region_desc") & " AS  geographical_region_desc,cs.collection_point_number,cp.collection_point,cs.start_time,cs.end_time,cs.Notes,cs.amended_by,cs.amended_on"
      Dim vAttrsList As String = "cs.collector_shift_number,cs.collector_number,cr.collection_region_number,cr.geographical_region,geographical_region_desc,cs.collection_point_number,cp.collection_point,cs.start_time,cs.end_time,cs.Notes,cs.amended_by,cs.amended_on"
      If mvParameters.Exists("CollectionPointNumber") Then vWhereFields.Add("cs.collection_point_number", mvParameters("CollectionPointNumber").LongValue)
      If mvParameters.Exists("CollectorNumber") Then vWhereFields.Add("cs.collector_number", mvParameters("CollectorNumber").LongValue)
      If mvParameters.Exists("CollectorShiftNumber") Then vWhereFields.Add("cs.collector_shift_number", mvParameters("CollectorShiftNumber").LongValue)
      If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add("mc.collection_number", mvParameters("CollectionNumber").LongValue)
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("mc.contact_number", mvParameters("ContactNumber").LongValue)
      Dim vSQL As String = "SELECT " & vConAttrs & "," & vAttrs & ",c.contact_number FROM collector_shifts cs INNER JOIN collection_points cp ON cp.collection_point_number = cs.collection_point_number INNER JOIN collection_regions cr ON cr.collection_region_number = cp.collection_region_number INNER JOIN geographical_regions gr ON gr.geographical_region = cr.geographical_region INNER JOIN manned_collectors mc ON mc.collector_number = cs.collector_number INNER JOIN contacts c ON c.contact_number = mc.contact_number LEFT OUTER JOIN organisations o ON gr.organisation_number = o.organisation_number WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), ContactNameItems() & "," & vAttrsList)
    End Sub
    Private Sub GetContactAccounts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      If mvParameters.HasValue("BankDetailsNumber") Then
        vWhereFields.Add("bank_details_number", mvParameters("BankDetailsNumber").LongValue)
      ElseIf mvParameters.HasValue("IbanNumber") Then
        vWhereFields.Add("iban_number", mvParameters("IbanNumber").Value)
      Else
        vWhereFields.Add("sort_code", mvParameters("SortCode").Value)
        vWhereFields.Add("account_number", mvParameters("AccountNumber").Value)
      End If
      vWhereFields.AddJoin("ca.contact_number", "c.contact_number")
      Dim vAttrs As String = "bank_details_number,address_number,sort_code,account_number,account_name,bank_payer_name,ca.notes,"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then vAttrs &= "iban_number"
      Dim vSQLAttrs As String = RemoveBlankItems(vAttrs & ",")
      If vSQLAttrs.EndsWith(",") = False Then vSQLAttrs += ","
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vSQLAttrs & vContact.GetRecordSetFieldsName, "contact_accounts ca, contacts c", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, "contact_number,CONTACT_NAME")
    End Sub
    Private Sub GetContactActions(ByVal pDataTable As CDBDataTable)
      GetActions(pDataTable)
    End Sub
    Private Sub GetContactAddresses(ByVal pDataTable As CDBDataTable)
      Dim vTable As String
      Dim vAttr As String
      Dim vAddressUsageTale As String
      Dim vAddressUsageAttr As String
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vTable = "organisation_addresses ca"
        vAttr = "ca.organisation_number"
        vAddressUsageTale = "organisation_address_usages cau"
        vAddressUsageAttr = "cau.organisation_number"
      Else
        vTable = "contact_addresses ca"
        vAttr = "ca.contact_number"
        vAddressUsageTale = "contact_address_usages cau"
        vAddressUsageAttr = "cau.contact_number"
      End If
      Dim vRegions As Boolean = mvEnv.GetConfigOption("cd_use_government_regions")
      Dim vAttrs As New StringBuilder(vAttr + ",ca.address_number,address_type,house_name,address,town,county,postcode,a.country,ca.valid_from,ca.valid_to,ca.historical,branch,paf,ca.amended_by,ca.amended_on,sortcode,uk,country_desc")
      Dim vFields As New StringBuilder(vAttrs.ToString)
      vFields.Append(",a.amended_by AS address_amended_by,a.amended_on AS address_amended_on")
      If vRegions Then
        vFields.Append(",government_region_desc")
      Else
        vFields.Append(",mosaic_code AS government_region_desc")
      End If
      vAttrs.Append(",address_amended_by,address_amended_on,government_region_desc")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then
        vFields.Append(",building_number")
        vAttrs.Append(",building_number")
      Else
        vAttrs.Append(",")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAddressDPS) Then
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vFields.Append(",delivery_point_suffix,lea_code,lea_name,o.organisation_number AS org_number2,o.name")
          vAttrs.Append(",delivery_point_suffix,lea_code,lea_name,org_number2,o.name")
        Else
          vFields.Append(",delivery_point_suffix,lea_code,lea_name,o.organisation_number,o.name")
          vAttrs.Append(",delivery_point_suffix,lea_code,lea_name,o.organisation_number,o.name")
        End If
      Else
        vAttrs.Append(",,,,,")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCountryAddressFormat) Then
        vFields.Append(",address_format")
        vAttrs.Append(",address_format")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAddressConfirmed) Then
        vFields.Append(",address_line1,address_line2,address_line3,address_confirmed")
        vAttrs.Append(",address_line1,address_line2,address_line3,address_confirmed")
      Else
        vFields.Append(",address_line1,address_line2,address_line3")
        vAttrs.Append(",address_line1,address_line2,address_line3")
      End If

      Dim vWithAddressUsages As Boolean = mvResultColumns.Contains(",AddressUsage,")
      If vWithAddressUsages Then
        vFields.Append(",address_usage")
        vAttrs.Append(",address_usage")
      End If

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("addresses a", "a.address_number", "ca.address_number")
      vAnsiJoins.Add("countries co", "co.country", "a.country")
      If vRegions Then vAnsiJoins.Add("government_regions gr", "a.mosaic_code", "gr.government_region")
      If vWithAddressUsages Then vAnsiJoins.AddLeftOuterJoin(vAddressUsageTale, "ca.address_number", "cau.address_number", vAttr, vAddressUsageAttr)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAddressDPS) Then vAnsiJoins.Add("address_data ad", "a.address_number", "ad.address_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vAnsiJoins.Add("organisations o", "ca.organisation_number", "o.organisation_number")
      Else
        vAnsiJoins.AddLeftOuterJoin("organisation_addresses oa", "a.address_number", "oa.address_number")
        vAnsiJoins.AddLeftOuterJoin("organisations o", "oa.organisation_number", "o.organisation_number")
      End If


      Dim vWhereFields As New CDBFields()
      vWhereFields.Add(vAttr, mvContact.ContactNumber)
      If mvParameters.Exists("AddressNumber") Then vWhereFields.Add("a.address_number", mvParameters("AddressNumber").LongValue)
      If vWithAddressUsages Then vWhereFields.Add("ca.historical", "N")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields.ToString, vTable, vWhereFields, "ca.historical, " & mvEnv.Connection.DBOrderByNullsFirstDesc("ca.valid_from") & ", town", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.ToString, "TOWN_ADDRESS_LINE,,ADDRESS_LINE" & If(vWithAddressUsages, ",ADDRESS_MULTI_LINE", ""))
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Historical")
        If vRow.Item("AddressNumber") = mvContact.AddressNumber.ToString Then
          vRow.Item("Default") = "Y"
          vRow.SetYNValue("Default")
        End If
      Next
    End Sub

    Private Sub GetContactAddressPositionAndOrg(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetContactAddressPositionAndOrg")
      Dim vSQL As String = "SELECT position,name,cp.organisation_number,cp.position_location,cp.started,cp.finished, " & mvEnv.Connection.DBSpecialCol("cp", "current") & " FROM contact_positions cp, organisations o WHERE cp.contact_number = " & mvContact.ContactNumber & " AND cp.address_number = " & mvParameters("AddressNumber").Value & " AND cp.organisation_number = o.organisation_number ORDER BY " & mvEnv.Connection.DBSpecialCol("cp", "current") & " DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactAddressUsages(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetContactAddressUsages")
      Dim vSQL As String = "SELECT address_number,cau.address_usage,address_usage_desc,notes,cau.amended_by,cau.amended_on FROM "
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = vSQL & "organisation_address_usages cau, address_usages au WHERE cau.organisation_number = " & mvContact.ContactNumber
      Else
        vSQL = vSQL & "contact_address_usages cau, address_usages au WHERE cau.contact_number = " & mvContact.ContactNumber
      End If
      If mvParameters.Exists("AddressNumber") Then
        vSQL = vSQL & " AND address_number = " & mvParameters("AddressNumber").LongValue
      End If
      If mvParameters.Exists("AddressUsage") Then
        vSQL = vSQL & " AND cau.address_usage = '" & mvParameters("AddressUsage").Value & "'"
      End If
      vSQL = vSQL & " AND cau.address_usage = au.address_usage ORDER BY address_usage_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactAppointments(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "contact_number,start_date,end_date,record_type,unique_id,description,time_status,amended_by,amended_on"
      vWhereFields.Add("contact_number", mvContact.ContactNumber)
      If mvParameters.HasValue("RecordType") Then vWhereFields.Add("record_type", mvParameters("RecordType").Value)
      If mvParameters.HasValue("UniqueId") Then vWhereFields.Add("unique_id", mvParameters("UniqueId").LongValue)
      If mvParameters.Exists("EndDate") Then vWhereFields.Add("start_date", CDBField.FieldTypes.cftTime, mvParameters("EndDate").Value, CDBField.FieldWhereOperators.fwoLessThan)
      If mvParameters.Exists("StartDate") Then vWhereFields.Add("end_date", CDBField.FieldTypes.cftTime, mvParameters("StartDate").Value, CDBField.FieldWhereOperators.fwoGreaterThan)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_appointments", vWhereFields, "start_date")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactAppointmentDetails(ByVal pDataTable As CDBDataTable)
      'Only supporting ServiceBookings at the moment
      'Using the ContactAppointment, retrieve the details of it, such as the ServiceBooking details (used so that Smart Client can navigate to it)
      Dim vAttrs As String = "ca.contact_number AS ca_contact_number,record_type,unique_id,c.contact_number,a.address_number,"
      Dim vContactAttr As String = "contact_number"
      Dim vAddressAttr As String = "address_number"
      Dim vAnsiJoins As New AnsiJoins()
      Select Case mvParameters("RecordType").Value
        Case "S"
          'ServiceBookings
          vAttrs &= "batch_number,transaction_number,line_number"
          vContactAttr = "booking_" & vContactAttr
          vAddressAttr = "booking_" & vAddressAttr
          With vAnsiJoins
            .Add("service_bookings x", "unique_id", "service_booking_number")
            .Add("contacts c", "x." & vContactAttr, "c.contact_number")
            .Add("addresses a", "x." & vAddressAttr, "a.address_number")
          End With
        Case Else
          vAttrs &= ",,"
      End Select
      Dim vWhereFields As New CDBFields(New CDBField("ca.contact_number", mvContact.ContactNumber))
      With vWhereFields
        .Add("record_type", mvParameters("RecordType").Value)
        .Add("unique_id", mvParameters("UniqueId").IntegerValue)
        .Add("ca.start_date", CDBField.FieldTypes.cftTime, mvParameters("StartDate").Value)
        .Add("ca.end_date", CDBField.FieldTypes.cftTime, mvParameters("EndDate").Value)
      End With
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_appointments ca", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactAppropriateCertificates(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsi As New CDBFields
      Dim vAttrs As String = "contact_number,certificate_number,claim_number,start_date,end_date,certificate_amount,tax_status,signature_date,amount_claimed,amount_paid,gac.cancellation_reason,cancelled_by,cancelled_on,gac.cancellation_source,created_by,created_on,gac.amended_by,gac.amended_on,cancellation_reason_desc,source_desc,tax_status AS TaxStatusCode"
      vWhereFields.Add("contact_number", mvContact.ContactNumber)
      If mvParameters.Exists("CertificateNumber") Then vWhereFields.Add("certificate_number", CDBField.FieldTypes.cftLong, mvParameters("CertificateNumber").Value, CDBField.FieldWhereOperators.fwoEqual)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "gac.cancellation_reason", "cr.cancellation_reason")
      vAnsiJoins.AddLeftOuterJoin("sources s", "gac.cancellation_source", "s.source")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "ga_appropriate_certificates gac", vWhereFields, "start_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      For Each vRow As CDBDataRow In pDataTable.Rows
        Select Case vRow.Item("TaxStatus")
          Case "S"
            vRow.Item("TaxStatus") = "Standard"
          Case "H"
            vRow.Item("TaxStatus") = "High"
        End Select
      Next
    End Sub
    Private Sub GetContactBackOrders(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "bod.batch_number,bod.transaction_number,bod.line_number,transaction_date,bo.bank_account,bod.product,product_desc,bod.rate,rate_desc,ordered,issued,bo.batch_type,earliest_delivery,status,reference,batch_type_desc,bank_account_desc,r.currency_code"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs = vAttrs.Replace(",r.currency_code", "")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("bo.contact_number", mvContact.ContactNumber)
      vWhereFields.AddJoin("bo.batch_number", "bod.batch_number")
      vWhereFields.AddJoin("bo.transaction_number", "bod.transaction_number")
      vWhereFields.AddJoin("bod.product", "p.product")
      vWhereFields.AddJoin("p.product", "r.product")
      vWhereFields.AddJoin("bod.rate", "r.rate")
      vWhereFields.AddJoin("bo.batch_type", "bt.batch_type")
      vWhereFields.AddJoin("bo.bank_account", "ba.bank_account")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "back_orders bo,back_order_details bod, products p,rates r, batch_types bt, bank_accounts ba", vWhereFields, "transaction_date DESC")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub
    Private Sub GetContactBankAccounts(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As New StringBuilder
      vAttrs.Append("bank_details_number,sort_code,account_number,account_name,bank_payer_name,amended_by,amended_on,notes")
      Dim vOrderBy As New StringBuilder
      Dim vAdditionalItems As String = ","
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDefaultBankAccount) Then
        vAttrs.Append(",")
        vAttrs.Append(mvEnv.Connection.DBIsNull("default_account", "'N'"))
        vAttrs.Append("AS default_account")
        vOrderBy.Append("default_account DESC,")
      Else
        vAdditionalItems &= ","
      End If
      vOrderBy.Append("amended_on DESC")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHistoryOnlyAccount) Then
        vAttrs.Append(",history_only")
        vOrderBy.Append(",history_only")
      Else
        vAdditionalItems &= ","
      End If

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then
        vAttrs.Append(",iban_number,bic_code")
      End If

      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("BankDetailsNumber") Then vWhereFields.Add("ca.bank_details_number", CDBField.FieldTypes.cftLong, mvParameters("BankDetailsNumber").LongValue)
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, mvContact.ContactNumber)
      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vAttrs.ToString, "contact_accounts ca", vWhereFields, vOrderBy.ToString), "", vAdditionalItems)
      GetBankInfo(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("DefaultAccount")
        vRow.SetYNValue("HistoryOnly")
      Next
    End Sub
    Private Sub GetBankInfo(ByVal pDataTable As CDBDataTable)
      Dim vList As New CDBParameters
      Dim vCode As String

      For Each vRow As CDBDataRow In pDataTable.Rows
        vCode = vRow.Item("SortCode").Replace("-", "")
        If vCode.Length > 0 Then
          If Not vList.Exists(vCode) Then vList.Add(vCode, CDBField.FieldTypes.cftCharacter, vCode)
        End If
      Next
      If vList.Count > 0 Then
        Dim vUK As Boolean = mvEnv.IsDefaultCountryUK
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("sort_code", CDBField.FieldTypes.cftLong, vList.InList, CDBField.FieldWhereOperators.fwoIn)
        Dim vRS As CDBRecordSet = New SQLStatement(mvEnv.Connection, "sort_code,bank,branch_name", "banks", vWhereFields).GetRecordSet
        While vRS.Fetch()
          vCode = vRS.Fields(1).Value
          If vCode.Length > 5 AndAlso vUK = True Then vCode = String.Concat(vCode.Substring(0, 2), "-", vCode.Substring(2, 2), "-", vCode.Substring(4, 2))
          For Each vRow As CDBDataRow In pDataTable.Rows
            If vCode = vRow.Item("SortCode") Then
              vRow.Item("BankName") = vRS.Fields(2).Value
              vRow.Item("BranchName") = vRS.Fields(3).Value
            End If
          Next
        End While
        vRS.CloseRecordSet()
      End If
    End Sub
    Private Sub GetContactCancelledProvisionalTrans(ByVal pDataTable As CDBDataTable)
      Dim vConfirmedTrans As New ConfirmedTransaction(mvEnv)
      Dim vAttrs As String = "bt.batch_number,bt.transaction_number,transaction_type_desc,transaction_date,amount,payment_method_desc,reference,mailing,receipt,eligible_for_gift_aid,currency_amount,bt.notes,b.provisional,bt.transaction_type,bt.payment_method,currency_code"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs = vAttrs.Replace(",currency_code", "")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("bt.contact_number", mvContact.ContactNumber)
      vWhereFields.Add("ct.status", vConfirmedTrans.GetStatusCode(ConfirmedTransaction.ConfirmedTransactionStatus.Cancelled))
      vWhereFields.Add("b.provisional", "Y")
      Dim vIncludeBC As String = mvEnv.GetConfig("fp_batch_categories_show")
      Dim vExcludeBC As String = mvEnv.GetConfig("fp_batch_categories_hide")
      If vIncludeBC.Length > 0 Then vWhereFields.Add("batch_category", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vIncludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoIn)
      If vExcludeBC.Length > 0 Then
        vWhereFields.Add("batch_category#2", "", CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("batch_category#3", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vExcludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotIn Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("confirmed_transactions ct", "ct.provisional_batch_number", "bt.batch_number", "ct.provisional_trans_number", "bt.transaction_number")
      vAnsiJoins.AddLeftOuterJoin("batches b", "bt.batch_number", "b.batch_number")
      vAnsiJoins.AddLeftOuterJoin("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
      vAnsiJoins.AddLeftOuterJoin("payment_methods pm", "bt.payment_method", "pm.payment_method")

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "batch_transactions bt", vWhereFields, "bt.batch_number DESC, bt.transaction_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Receipt")
        vRow.SetYNValue("EligibleForGiftAid")
      Next
    End Sub
    Private Sub GetContactCashInvoices(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields

      Dim vCC As New CreditCustomer()
      If mvParameters.Exists("SalesLedgerAccount") Then
        'For Trader, this is the number entered on the Invoice Payments page
        vCC.InitCompanySalesLedgerAccount(mvEnv, mvParameters("Company").Value, mvParameters("SalesLedgerAccount").Value)
        'If vCC.Existing = False Then - Don't raise error just return no data
      Else
        vCC.Init(mvEnv, mvParameters("ContactNumber").IntegerValue, mvParameters("Company").Value)
        'If vCC.Existing = False Then  - Don't raise error just return no data
      End If

      '(1) Build LeftOuterjoin SQL for Cash Invoices
      Dim vOJSQLStatement As New SQLStatement(mvEnv.Connection, "batch_number, transaction_number, line_number, SUM(fhd.amount) AS amount", "financial_history_details fhd", New CDBFields(New CDBField("status", CDBField.FieldTypes.cftCharacter, "")), "")
      vOJSQLStatement.GroupBy = "batch_number, transaction_number, line_number"

      '(2) Now build main SQL for cash Invoices
      Dim vAttrs As New StringBuilder
      vAttrs.Append("MAX(i.batch_number) AS batch_number, MAX(i.transaction_number) AS transaction_number, i.invoice_number, invoice_date, MAX(payment_due) AS payment_due, MAX(amount_paid) AS amount_paid")
      vAttrs.Append(", MAX(i.invoice_pay_status) AS invoice_pay_status, MAX(invoice_dispute_code) AS invoice_dispute_code, MAX(record_type) AS record_type, ABS(SUM(x.amount)) AS amount, ABS(SUM(fhd.amount)) AS invoice_amount")
      vAttrs.Append(", MAX(batch_type) AS batch_type, MAX(fh.amount) AS fh_amount")
      vAttrs.Append(", MAX(i.contact_number) AS contact_number, MAX(i.address_number) AS address_number, i.sales_ledger_account")
      Dim vAttrNames As New StringBuilder
      vAttrNames.Append("batch_number,transaction_number,i.invoice_number,invoice_date,payment_due,amount_paid,invoice_pay_status,invoice_dispute_code")
      vAttrNames.Append(",record_type,amount,invoice_amount,batch_type,fh_amount,contact_number,address_number,sales_ledger_account,")

      With vWhereFields
        .Clear()
        .Add("i.sales_ledger_account", vCC.SalesLedgerAccount)
        .Add("i.company", vCC.Company)
        .Add("i.record_type", "C")
        .Add("ips.fully_paid", "N")
        .Add("ips.pending_dd_payment", "N")
        .Add("bta.line_type", CDBField.FieldTypes.cftInteger, "'U','N'", CDBField.FieldWhereOperators.fwoIn)
      End With

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("invoice_details id", "i.batch_number", "id.batch_number", "i.transaction_number", "id.transaction_number", "i.invoice_number", "id.invoice_number ")
        .Add("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
        .Add("batch_transaction_analysis bta", "id.batch_number", "bta.batch_number", "id.transaction_number", "bta.transaction_number", "id.line_number", "bta.line_number") ', "i.sales_ledger_account", "bta.member_number") 'BR17359: Removed join restriction on i.sales_ledger_account=bta.member_number, as i.sales_ledger_account in Where clause, to allow event booking credit notes to be available
        .Add("financial_history_details fhd", "bta.batch_number", "fhd.batch_number", "bta.transaction_number", "fhd.transaction_number", "bta.line_number", "fhd.line_number")
        .Add("financial_history fh", "fhd.batch_number", "fh.batch_number", "fhd.transaction_number", "fh.transaction_number")
        .Add("batches b", "i.batch_number", "b.batch_number")
        .AddLeftOuterJoin("(" & vOJSQLStatement.SQL & ") x", "bta.batch_number", "x.batch_number", "bta.transaction_number", "x.transaction_number", "bta.line_number", "x.line_number")
      End With

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs.ToString, "invoices i", vWhereFields, "", vAnsiJoins)
      Dim vGroupBy As String = "i.invoice_date, i.invoice_number, i.sales_ledger_account"
      vSQL.GroupBy = vGroupBy
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrNames.ToString)

      '(3) See if payer has cash (when payer is different to credit customer)
      If vCC.ContactNumber > 0 AndAlso vCC.ContactNumber.Equals(mvParameters("ContactNumber").IntegerValue) = False Then
        'For Trader, the Contact Number is the payer Contact and we got here because user entered a Sales Ledger Account for a different Contact
        Dim vPayerCC As New CreditCustomer()
        vPayerCC.Init(mvEnv, mvParameters("ContactNumber").IntegerValue, mvParameters("Company").Value)
        If vPayerCC.Existing Then
          vWhereFields.Item(1).Value = vPayerCC.SalesLedgerAccount
          vSQL = New SQLStatement(mvEnv.Connection, vAttrs.ToString, "invoices i", vWhereFields, "", vAnsiJoins)
          vSQL.GroupBy = vGroupBy
          pDataTable.FillFromSQL(mvEnv, vSQL, vAttrNames.ToString)
        Else
          'Don't raise error just return no data
        End If
      End If

      '(4) Manipulate AmountPaid etc. for financial adjustments
      Dim vInvoice As New Invoice
      vInvoice.Init(mvEnv)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vInvoice.Existing Then
          vInvoice = New Invoice
          vInvoice.Init(mvEnv)
        End If
        If vRow.Item("BatchType") = Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment) Then
          If DoubleValue(vRow.Item("TransactionAmount")) = 0 Then
            'Re-analysed transaction so retrieve invoice details of this transaction and the adjusted transaction
            Dim vAmountPaid As Double
            vRow.Item("InvoiceAmount") = vInvoice.GetAdjustmentInvoiceAmounts(IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")), False, vRow.Item("SalesLedgerAccount"), vAmountPaid).ToString("F")
            If DoubleValue(vRow.Item("InvoiceAmount")) = 0 Then
              vInvoice.Init(mvEnv, IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")), IntegerValue(vRow.Item("InvoiceNumber")))
              vRow.Item("InvoiceAmount") = vInvoice.InvoiceAmount.ToString("F")
              If vInvoice.InvoiceAmount > 0 AndAlso vAmountPaid < 0 Then
                vAmountPaid = Math.Abs(FixTwoPlaces((vInvoice.InvoiceAmount * -1) - vInvoice.AmountPaid))
              End If
            End If
            vRow.Item("AmountPaid") = vAmountPaid.ToString("F")
            If FixTwoPlaces(DoubleValue(vRow.Item("InvoiceAmount")) - vAmountPaid) < 0 AndAlso DoubleValue(vRow.Item("InvoiceAmount")) > 0 Then
              vInvoice.Init(mvEnv, IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")))
              If vInvoice.InvoiceAmount = DoubleValue(vRow.Item("InvoiceAmount")) Then
                vRow.Item("AmountPaid") = vInvoice.AmountPaid.ToString("F")
              End If
            End If
          End If
        End If
      Next

      '(5) Remove any redundant rows
      Dim vDeleteRow As CDBDataRow
      For vIndex As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
        vDeleteRow = pDataTable.Rows(vIndex)
        If (mvParameters.ParameterExists("BatchNumber").IntegerValue > 0 _
        AndAlso mvParameters("BatchNumber").Value = vDeleteRow.Item("BatchNumber") AndAlso mvParameters("TransactionNumber").Value = vDeleteRow.Item("TransactionNumber")) _
        OrElse (FixTwoPlaces(DoubleValue(vDeleteRow.Item("InvoiceAmount")) - DoubleValue(vDeleteRow.Item("AmountPaid"))) <= 0) Then
          pDataTable.Rows.RemoveAt(vIndex)
        End If
      Next

      '(6) Load any usable sundry credit notes (restrict to chosen credit customer only)
      vWhereFields.Item(1).Value = vCC.SalesLedgerAccount
      vWhereFields.Add("i.invoice_number", "", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields("i.record_type").Value = "N"
      vWhereFields.Remove("bta.line_type")
      vWhereFields.Add("bta.line_type", "'D','F','G','H','I','K','L','N','S','U'", CDBField.FieldWhereOperators.fwoNotIn)
      vSQL = New SQLStatement(mvEnv.Connection, vAttrs.ToString, "invoices i", vWhereFields, "", vAnsiJoins)
      vSQL.GroupBy = "i.invoice_date, i.invoice_number, i.sales_ledger_account"
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrNames.ToString)
    End Sub

    Private Sub GetContactCategories(ByVal pDataTable As CDBDataTable)
      Dim vDepartmentalOnly As Boolean = mvEnv.GetConfigOption("cd_display_owning_activities")
      Dim vHighProfile As Boolean = mvParameters.ParameterExists("HighProfile").Value = "Y" 'WS: I can't see any of the code passing this parameter. Also this is not in XMLHelper class and hence can't even be used by Web Services.
      Dim vCurrentOnly As Boolean = mvParameters.ParameterExists("Current").Bool
      GetContactCategories(pDataTable, vDepartmentalOnly, vHighProfile, vCurrentOnly)
    End Sub

    Private Sub GetContactCategories(ByVal pDataTable As CDBDataTable, ByVal pDepartmentalOnly As Boolean, ByVal pHighProfile As Boolean, ByVal pCurrentOnly As Boolean)
      'pDepartmentalOnly is only used to display the owning activities. This could be set due to the config cd_display_owning_activities
      'OR when reading categories for dstContactDeptCategories and dstContactHeaderDeptCategories
      Dim vIdentifier As String
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vIdentifier = "cc.organisation_number,cc.organisation_category_number"
      Else
        vIdentifier = "cc.contact_number,cc.contact_category_number"
      End If
      Dim vAttrs As String
      Dim vDistinct As Boolean
      If mvType = DataSelectionTypes.dstContactHPCategories Then
        vAttrs = "profile_rating,a.activity,activity_desc"
        vDistinct = True
      Else
        vAttrs = vIdentifier & ",cc.activity,cc.activity_value,quantity,activity_date,cc.source,valid_from,valid_to,cc.amended_by,cc.amended_on,cc.notes,activity_desc,activity_value_desc,source_desc,rgb_value,cc.response_channel,response_channel_desc"
      End If
      If mvType = DataSelectionTypes.dstContactCategories Then
        vAttrs += ",a.is_historic As is_activity_historic,av.is_historic As is_activity_value_historic"
      End If
      Dim vOrderBy As String
      Dim vHighProfileSelection As HighProfileActivitySelection = HighProfileActivitySelection.BothHighProfileAndNonHighProfile
      If mvType = DataSelectionTypes.dstContactDeptCategories OrElse mvType = DataSelectionTypes.dstContactHeaderDeptCategories Then
        'Only exclude the HP Categories when displaying the categories in Contact Header Departmental Categories grid or Contact Departmental Categories grid in General tab
        vHighProfileSelection = HighProfileActivitySelection.ExcludeHighProfile
        pHighProfile = False  'In theory this will not be passed as True when dealing with above two Departmental Categories types but just in case, set it to False.
      End If
      If pHighProfile Then
        vOrderBy = "profile_rating, activity_desc"
        If vDistinct = False Then vOrderBy &= ", activity_value_desc"
        vHighProfileSelection = HighProfileActivitySelection.HighProfileOnly
      Else
        vOrderBy = "activity_desc, activity_value_desc"
      End If
      Dim vSQLStatement As SQLStatement = GetCategorySQL(pDepartmentalOnly, vHighProfileSelection, RemoveBlankItems(vAttrs), vOrderBy, vDistinct, pCurrentOnly)
      If mvType = DataSelectionTypes.dstContactCategories Then
        vAttrs = vAttrs.Replace(",a.is_historic As is_activity_historic,av.is_historic As is_activity_value_historic", ",is_activity_historic,is_activity_value_historic")
      End If
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",,,,")
      Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
      Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")
      Dim vAccess As Boolean = pDataTable.Columns.ContainsKey("Access")

      If vStatus Then pDataTable.Columns("Status").AttributeName = "status" 'Why
      If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag" 'Why

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
        If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        If pDepartmentalOnly AndAlso vAccess Then vRow.Item("Access") = "Y"
      Next
      If Not pDepartmentalOnly AndAlso vAccess Then
        Dim vDeptSQLStatement As SQLStatement = GetCategorySQL(True, HighProfileActivitySelection.BothHighProfileAndNonHighProfile, "cc.activity, cc.activity_value", "", True, pCurrentOnly)
        Dim vRecordSet As CDBRecordSet = vDeptSQLStatement.GetRecordSet
        While vRecordSet.Fetch
          Dim vActivity As String = vRecordSet.Fields(1).Value
          Dim vValue As String = vRecordSet.Fields(2).Value
          For Each vRow As CDBDataRow In pDataTable.Rows
            If vRow.Item("ActivityCode") = vActivity And vRow.Item("ActivityValueCode") = vValue Then vRow.Item("Access") = "Y"
          Next
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vStatus Then pDataTable.ReOrderRowsByColumn("StatusOrder")

    End Sub

    Private Enum HighProfileActivitySelection As Integer
      BothHighProfileAndNonHighProfile
      HighProfileOnly
      ExcludeHighProfile
    End Enum

    Private Function GetCategorySQL(ByVal pDepartmentalOnly As Boolean, ByVal pHighProfileSelection As HighProfileActivitySelection, ByVal pFields As String, ByVal pOrderBy As String, ByVal pDistinct As Boolean, ByVal pCurrentOnly As Boolean) As SQLStatement
      'Only gets called from GetContactCategories & GetContactPositionCategories
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("activities a", "cc.activity", "a.activity")
      If mvParameters.Exists("ActivityGroup") AndAlso mvParameters("ActivityGroup").Value.Length > 0 Then
        vAnsiJoins.Add("activity_group_details agd", "cc.activity", "agd.activity")
      End If
      vAnsiJoins.Add("activity_values av", "cc.activity", "av.activity", "cc.activity_value", "av.activity_value")
      vAnsiJoins.Add("sources s", "cc.source", "s.source")
      If pDepartmentalOnly Then
        vAnsiJoins.Add("activity_users au", "a.activity", "au.activity")
        vAnsiJoins.Add("activity_value_users avu", "av.activity", "avu.activity", "av.activity_value", "avu.activity_value")
      End If
      vAnsiJoins.AddLeftOuterJoin("response_channels rc", "cc.response_channel", "rc.response_channel")
      Dim vWhereFields As New CDBFields()
      Dim vTableName As String
      If mvParameters.Exists("ContactPositionNumber") Then
        vTableName = "contact_position_activities cc"
        vWhereFields.Add("contact_position_number", mvParameters("ContactPositionNumber").IntegerValue)
      ElseIf mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vTableName = "organisation_categories cc"
        vWhereFields.Add("organisation_number", mvContact.ContactNumber)
      Else
        vTableName = "contact_categories cc"
        vWhereFields.Add("contact_number", mvContact.ContactNumber)
      End If
      If mvParameters.Exists("ContactCategoryNumber") Then
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vWhereFields.Add("organisation_category_number", mvParameters("ContactCategoryNumber").IntegerValue)
        Else
          vWhereFields.Add("contact_category_number", mvParameters("ContactCategoryNumber").IntegerValue)
        End If
      End If
      If mvParameters.Exists("ContactPositionActivityId") Then
        vWhereFields.Add("contact_position_activity_id", mvParameters("ContactPositionActivityId").IntegerValue)
      End If
      If mvParameters.Exists("Activity") Then vWhereFields.Add("cc.activity", mvParameters("Activity").Value)
      If mvParameters.Exists("ActivityValue") Then vWhereFields.Add("cc.activity_value", mvParameters("ActivityValue").Value)
      If mvParameters.Exists("Activities") Then vWhereFields.Add("cc.activity", mvParameters("Activities").Value, CDBField.FieldWhereOperators.fwoIn)
      If mvParameters.Exists("Source") Then vWhereFields.Add("cc.source", mvParameters("Source").Value)
      If mvParameters.Exists("ValidFrom") Then
        vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value)
      ElseIf pCurrentOnly Then
        vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If
      If mvParameters.Exists("ValidTo") Then
        vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ValidTo").Value)
      ElseIf pCurrentOnly Then
        vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      If mvParameters.Exists("ActivityGroup") AndAlso mvParameters("ActivityGroup").Value.Length > 0 Then vWhereFields.Add("agd.activity_group", mvParameters("ActivityGroup").Value)
      If mvParameters.Exists("AmendedOn") Then vWhereFields.Add("cc.amended_on", CDBField.FieldTypes.cftDate, mvParameters("AmendedOn").Value)
      If pDepartmentalOnly Then
        vWhereFields.Add("au.department", mvEnv.User.Department)
        vWhereFields.Add("avu.department", mvEnv.User.Department)
      End If
      Select Case pHighProfileSelection
        Case HighProfileActivitySelection.HighProfileOnly 'Should be set for dstContactHPCategories, dstContactHPCategoryValues and dstContactHeaderHPCategories
          vWhereFields.Add("a.high_profile", "Y")
        Case HighProfileActivitySelection.ExcludeHighProfile  'Should only be set for dstContactHeaderDeptCategories or dstContactDeptCategories (used in Contact General Tab)
          vWhereFields.Add("a.high_profile", "Y", CDBField.FieldWhereOperators.fwoNotEqual)
      End Select
      Dim vSQL As New SQLStatement(mvEnv.Connection, pFields, vTableName, vWhereFields, pOrderBy, vAnsiJoins)
      vSQL.Distinct = pDistinct
      Return vSQL
    End Function

    Private Sub GetContactCategoryGraphData(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetContactCategoryGraphData")
      Dim vAttrs As String = "cc.activity_value,quantity,valid_from,valid_to,activity_value_desc"
      Dim vSQL As String = "SELECT " & vAttrs
      vSQL = vSQL & " FROM contact_categories cc, activity_values av"
      vSQL = vSQL & " WHERE contact_number = " & mvContact.ContactNumber
      If mvParameters.Exists("Activity") Then vSQL = vSQL & " AND cc.activity = '" & mvParameters("Activity").Value & "'"
      If mvParameters.Exists("ValidFrom") Then vSQL = vSQL & " AND valid_from " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value) & " AND valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDate)
      vSQL = vSQL & " AND cc.activity = av.activity AND cc.activity_value = av.activity_value"
      vSQL = vSQL & " ORDER BY activity_value_desc, valid_from"
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vSQL = Replace(vSQL, "contact", "organisation")
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetContactCollectionPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "box_reference,cp.collection_box_number,transaction_date,cp.amount,cp.collection_payment_number,cp.collection_number,cp.collection_pis_number,pis_number,cp.batch_number,cp.transaction_number,cp.line_number,fl.contact_number,fh.contact_number  AS  fh_contact_number"
      Dim vAttrsList As String = "box_reference,cp.collection_box_number,transaction_date,cp.amount,cp.collection_payment_number,cp.collection_number,cp.collection_pis_number,pis_number,cp.batch_number,cp.transaction_number,cp.line_number,fl.contact_number,fh_contact_number"
      With vWhereFields
        If mvParameters.Exists("CollectionNumber") Then .Add("cp.collection_number", mvParameters("CollectionNumber").LongValue)
        If mvParameters.Exists("CollectionPaymentNumber") Then .Add("collection_payment_number", mvParameters("CollectionPaymentNumber").LongValue)
        If mvParameters.Exists("CollectionPISNumber") Then .Add("cp.collection_pis_number", mvParameters("CollectionPISNumber").LongValue)
        If mvParameters.Exists("ContactNumber") Then .Add("fh.contact_number", mvParameters("ContactNumber").LongValue)
        .Add("fl.donor_contact_number", CDBField.FieldTypes.cftLong, "")  'to make sure that we only pick records that have paid for themselves
      End With
      Dim vSQL As String = "SELECT " & vAttrs & "," & vConAttrs & " FROM financial_history fh"
      vSQL = vSQL & " INNER JOIN contacts c ON fh.contact_number = c.contact_number"
      vSQL = vSQL & " INNER JOIN collection_payments cp ON fh.batch_number = cp.batch_number AND  fh.transaction_number = cp.transaction_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_pis pis ON cp.collection_pis_number = pis.collection_pis_number"
      vSQL = vSQL & " LEFT OUTER JOIN financial_links fl ON cp.batch_number = fl.batch_number AND cp.transaction_number = fl.transaction_number AND cp.line_number = fl.line_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_boxes cb ON cp.collection_box_number = cb.collection_box_number"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrsList & "," & ContactNameItems() & ",")
      vWhereFields.Remove("fh.contact_number")
      vWhereFields.Add("fl.contact_number", mvParameters("ContactNumber").LongValue)
      vWhereFields.Remove("fl.donor_contact_number")
      vSQL = "SELECT " & vAttrs & "," & vConAttrs & " FROM financial_history fh"
      vSQL = vSQL & " INNER JOIN contacts c ON fh.contact_number = c.contact_number "
      vSQL = vSQL & " INNER JOIN collection_payments cp ON fh.batch_number = cp.batch_number AND  fh.transaction_number = cp.transaction_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_pis pis ON cp.collection_pis_number = pis.collection_pis_number "
      vSQL = vSQL & " LEFT OUTER JOIN financial_links fl ON cp.batch_number = fl.batch_number AND cp.transaction_number = fl.transaction_number AND cp.line_number = fl.line_number "
      vSQL = vSQL & " LEFT OUTER JOIN collection_boxes cb ON cp.collection_box_number = cb.collection_box_number"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrsList & "," & ContactNameItems() & ",")
      GetContactNames(pDataTable, "SentOnBehalfOfContactNumber", "SentOnBehalfOfContactName")
    End Sub
    Private Sub GetContactCommsNumbers(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "co.contact_number,co.address_number,co.device,device_desc,co.dialling_code,co.std_code,extension,co.ex_directory,co.notes,co.amended_by,co.amended_on,valid_from,valid_to,is_active,co.mail,device_default,preferred_method,cu.communication_usage,communication_usage_desc,isorganisation,co.communication_number,telephone,subscription_count,email,www_address,archive"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbArchiveCommunications) Then vAttrs = vAttrs.Replace("archive", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then vAttrs = vAttrs.Replace("valid_from,valid_to,is_active,co.mail,device_default,preferred_method", ",,,,,")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCommunicationsUsage) AndAlso mvType <> DataSelectionTypes.dstContactCommsNumbersWithUsages Then vAttrs = vAttrs.Replace("cu.communication_usage,communication_usage_desc", ",")
      Dim vFields As String = RemoveBlankItems(vAttrs).Replace("telephone", mvEnv.Connection.DBSpecialCol("", "number") & " AS telephone")
      vFields = vFields.Replace("isorganisation", "'' AS isorganisation")

      Dim vWithCommunicationUsages As Boolean = mvType = DataSelectionTypes.dstContactCommsNumbersWithUsages

      Dim vOrderBy As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then vOrderBy = "preferred_method DESC, device_default DESC, is_active DESC, "
      vOrderBy = vOrderBy & "d.sequence_number, device_desc"

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("devices d", "d.device", "co.device")
      vAnsiJoins.AddLeftOuterJoin("(SELECT communication_number, COUNT(*) as subscription_count FROM subscriptions GROUP BY communication_number) x", "co.communication_number", "x.communication_number")
      If vWithCommunicationUsages Then
        vAnsiJoins.AddLeftOuterJoin("contact_communication_usages ccu", "co.communication_number", "ccu.communication_number")
        vAnsiJoins.AddLeftOuterJoin("communication_usages cu", "ccu.communication_usage", "cu.communication_usage")
      ElseIf mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCommunicationsUsage) Then
        vAnsiJoins.AddLeftOuterJoin("communication_usages cu", "co.communication_usage", "cu.communication_usage")
      End If

      Dim vWhereFields As New CDBFields()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbArchiveCommunications) Then vWhereFields.Add("Archive", "N")
      If mvContact.ContactType <> Contact.ContactTypes.ctcOrganisation Then vWhereFields.Add("co.contact_number", mvContact.ContactNumber)
      If mvParameters.Exists("CommunicationNumber") Then vWhereFields.Add("co.communication_number", mvParameters("CommunicationNumber").LongValue)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then
        If mvParameters.Exists("Active") Then
          vWhereFields.Add("is_active ", mvParameters("Active").Value.Substring(0, 1))
        ElseIf vWithCommunicationUsages Then
          vWhereFields.Add("is_active ", "Y")
        End If
      End If
      If mvParameters.Exists("AddressNumber") Then vWhereFields.Add("co.address_number ", mvParameters("AddressNumber").LongValue)

      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        If Not mvParameters.Exists("LinkedContactsRequired") Then
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, "SELECT address_number FROM organisation_addresses WHERE organisation_number = " & mvContact.ContactNumber, CType(CDBField.FieldWhereOperators.fwoIn + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
          vWhereFields.Add("co.contact_number", "", CDBField.FieldWhereOperators.fwoCloseBracket)
        Else
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, "SELECT address_number FROM organisation_addresses WHERE organisation_number = " & mvContact.ContactNumber, CType(CDBField.FieldWhereOperators.fwoIn + CDBField.FieldWhereOperators.fwoOpenBracket + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        End If
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "communications co", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "CONTACT_TELEPHONE,,")

      If mvContact.ContactType <> Contact.ContactTypes.ctcOrganisation Then
        'Now add the comms records that relate to current positions for the contact
        vAnsiJoins.Clear()
        vAnsiJoins.Add("communications co", "cp.address_number", "co.address_number")
        vAnsiJoins.Add("devices d", "d.device", "co.device")
        vAnsiJoins.AddLeftOuterJoin("(SELECT communication_number, COUNT(*) as subscription_count FROM subscriptions GROUP BY communication_number) x", "co.communication_number", "x.communication_number")
        If vWithCommunicationUsages Then
          vAnsiJoins.AddLeftOuterJoin("contact_communication_usages ccu", "co.communication_number", "ccu.communication_number")
          vAnsiJoins.AddLeftOuterJoin("communication_usages cu", "ccu.communication_usage", "cu.communication_usage")
        ElseIf mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCommunicationsUsage) Then
          vAnsiJoins.AddLeftOuterJoin("communication_usages cu", "co.communication_usage", "cu.communication_usage")
        End If

        vWhereFields.Remove("co.contact_number")
        vWhereFields.Add("cp.contact_number", mvContact.ContactNumber)
        vWhereFields.Add("co.contact_number")
        vWhereFields.Add("co.is_active", "Y")
        vWhereFields.Add("current", "Y").SpecialColumn = True

        Dim vDevices As String = mvEnv.GetConfig("cd_org_numbers_device_codes")
        If vDevices.Length > 0 Then
          Dim vItems As New StringList(vDevices)
          vWhereFields.Add("co.device", CDBField.FieldTypes.cftCharacter, vItems.InList, CDBField.FieldWhereOperators.fwoIn)
        End If

        vSQL = New SQLStatement(mvEnv.Connection, vFields, "contact_positions cp", vWhereFields, vOrderBy, vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "CONTACT_TELEPHONE,,")
      End If

      If mvType = DataSelectionTypes.dstContactCommsNumbers OrElse mvType = DataSelectionTypes.dstContactCommsNumbersWithUsages Then FillAddressData(pDataTable)

      Dim vOrg As New Organisation(mvEnv)
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vOrg.Init(mvContact.ContactNumber)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("Email") = "" Then vRow.Item("Email") = "N"
        If vRow.Item("ExDirectory") = "Y" And mvContact.Department <> mvEnv.User.Department Then
          vRow.Item("PhoneNumber") = DataSelectionText.String23335    'Ex-Directory
        Else
          If Len(vRow.Item("Extension")) > 0 Then
            vRow.Item("PhoneNumber") = String.Format("{0} {1} {2}", vRow.Item("PhoneNumber"), DataSelectionText.String23336, vRow.Item("Extension"))    ' Ext
          End If
        End If
        vRow.SetYNValue("ExDirectory")
        vRow.SetYNValue("IsActive")
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("DeviceDefault")
        vRow.SetYNValue("PreferredMethod")
        If mvType = DataSelectionTypes.dstContactCommsNumbers OrElse mvType = DataSelectionTypes.dstContactCommsNumbersWithUsages Then
          If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            If vRow.Item("Number") = vOrg.Telephone And vRow.Item("STDCode") = vOrg.STDCode And vRow.Item("DiallingCode") = vOrg.DiallingCode Then vRow.Item("Default") = "Y"
          Else
            If vRow.Item("Number") = mvContact.Telephone And vRow.Item("STDCode") = mvContact.StdCode And vRow.Item("DiallingCode") = mvContact.DiallingCode Then vRow.Item("Default") = "Y"
            If vRow.Item("ContactNumber").Length = 0 Then vRow.Item("IsOrganisation") = "Yes"
          End If
          vRow.SetYNValue("Default")
        End If
      Next
    End Sub
    Private Sub GetContactCommsNumbersEdit(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      ''NYI("GetContactCommsNumbersEdit")
      Dim vAttrs As String = "co.contact_number,co.address_number,co.device,device_desc,co.dialling_code,co.std_code,extension,co.ex_directory,co.notes,co.amended_by,co.amended_on,communication_number,telephone,is_active,mail,preferred_method,device_default,valid_from,valid_to"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then vAttrs = Replace$(vAttrs, "communication_number", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then
        vAttrs = Replace(Replace(vAttrs, "is_active", ""), "mail", "")
        vAttrs = Replace(Replace(vAttrs, "preferred_method", ""), "device_default", "")
      End If
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace$(RemoveBlankItems(vAttrs), "telephone", mvEnv.Connection.DBSpecialCol("", "number") & " AS telephone")
      vSQL = vSQL & " FROM communications co, devices d WHERE "
      Dim vAttr As String = ",,"
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = vSQL & "address_number IN (select address_number FROM organisation_addresses WHERE organisation_number = " & mvContact.ContactNumber & ") AND contact_number IS NULL"
        vSQL = vSQL & " AND co.device = d.device"
      Else
        vSQL = vSQL & "contact_number = " & mvContact.ContactNumber
        vSQL = vSQL & " AND co.device = d.device"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDevicesSequenceNumber) Then
        vSQL = vSQL & " ORDER BY d.sequence_number, device_desc"
      Else
        vSQL = vSQL & " ORDER BY device_desc"
      End If
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "CONTACT_TELEPHONE" & vAttr)
      GetAddressData(pDataTable)
      Dim vOrg As New Organisation(mvEnv)
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vOrg.Init(mvContact.ContactNumber)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ExDirectory") = "Y" And mvContact.Department <> mvEnv.User.Department Then
          vRow.Item("PhoneNumber") = DataSelectionText.String23335    'Ex-Directory
        Else
          If Len(vRow.Item("Extension")) > 0 Then
            vRow.Item("PhoneNumber") = String.Format("{0} {1} {2}", vRow.Item("PhoneNumber"), DataSelectionText.String23336, vRow.Item("Extension"))    ' Ext
          End If
        End If
        vRow.SetYNValue("ExDirectory")
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          If vRow.Item("Number") = vOrg.Telephone And vRow.Item("STDCode") = vOrg.STDCode And vRow.Item("DiallingCode") = vOrg.DiallingCode Then vRow.Item("Default") = "Y"
        Else
          If vRow.Item("Number") = mvContact.Telephone And vRow.Item("STDCode") = mvContact.StdCode And vRow.Item("DiallingCode") = mvContact.DiallingCode Then vRow.Item("Default") = "Y"
        End If
        vRow.SetYNValue("Default")
        vRow.SetYNValue("IsActive")
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("PreferredMethod")
        vRow.SetYNValue("DeviceDefault")
      Next
    End Sub
    Private Sub GetContactCommunicationUsages(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("ccu.communication_number", mvParameters("CommunicationNumber").IntegerValue)
      If mvParameters.Exists("CommunicationUsage") Then vWhereFields.Add("ccu.communication_usage", mvParameters("CommunicationUsage").Value)
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("ccu.contact_number", mvParameters("ContactNumber").IntegerValue)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("communication_usages cu", "ccu.communication_usage", "cu.communication_usage")
      vAnsiJoins.Add("communications co", "ccu.communication_number", "co.communication_number")
      Dim vFields As String = "ccu.communication_number,ccu.communication_usage,communication_usage_desc,ccu.notes,co.communication_usage as primary_usage,ccu.amended_by,ccu.amended_on"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCommunicationsUsage) Then vFields = vFields.Replace("co.communication_usage", "''")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "contact_communication_usages ccu", vWhereFields, "communication_usage_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("PrimaryUsage") = vRow.Item("CommunicationUsage") Then
          vRow.Item("PrimaryUsage") = ProjectText.String15904
        Else
          vRow.Item("PrimaryUsage") = ""
        End If
      Next
    End Sub
    Private Sub GetContactCovenants(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactCovenants")
      Dim vAttrs As String = "c.order_number,covenant_number,start_date,signature_date,covenant_term,c.created_by,c.created_on,c.amended_by,c.amended_on,c.cancelled_by,c.cancelled_on,last_tax_claim,tax_claimed_to,r185_return,r185_sent,covenanted_amount,c.cancellation_reason,o.payment_frequency,o.payment_method,payment_frequency_desc,payment_method_desc,deposited_deed,net,annual_claim,fixed,covenant_status,c.cancellation_source"
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs) & " FROM covenants c, orders o, payment_frequencies pf, payment_methods pm WHERE c.contact_number = " & mvContact.ContactNumber & " AND c.order_number = o.order_number AND o.payment_frequency = pf.payment_frequency AND o.payment_method = pm.payment_method"
      vSQL = vSQL & " ORDER BY c.cancelled_on " & mvEnv.Connection.DBSortByNullsFirst & ", c.start_date DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, , ",,,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetDescriptions(pDataTable, "CovenantStatus")
    End Sub
    Private Sub GetContactCPDCyclesEdit(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vSelectCols As String = ""
      If mvParameters.Exists("ContactNumber") Then
        vWhereFields.Add("ccc.contact_number", CDBField.FieldTypes.cftLong, mvParameters("ContactNumber").Value)
      Else
        vWhereFields.Add("'1'", CDBField.FieldTypes.cftLong, "'2'", CDBField.FieldWhereOperators.fwoEqual)
      End If
      If mvParameters.Exists("ContactCpdCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", CDBField.FieldTypes.cftLong, mvParameters("ContactCpdCycleNumber").Value)

      Dim vAttrs As String = "ccc.contact_cpd_cycle_number,ccc.cpd_cycle_type,cpd_cycle_type_desc,start_month,end_month,start_date,end_date,ccc.amended_on,ccc.amended_by"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDCycleStatus) Then
        vAttrs = vAttrs & ",ccs.cpd_cycle_status,ccs.cpd_cycle_status_desc,ccs.rgb_value"
        vSelectCols = vAttrs
      Else
        vSelectCols = vAttrs & ",Null as cpd_cycle_status,Null as cpd_cycle_status_desc,Null as rgb_value"
        vAttrs = vAttrs & ",cpd_cycle_status,cpd_cycle_status_desc,rgb_value"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDObjective) Then
        vAttrs = vAttrs & ",ccty.cpd_type,ccc.contact_number"
        vSelectCols = vAttrs
      Else
        vSelectCols = vAttrs & ",Null as cpd_type,ccc.contact_number"
        vAttrs = vAttrs & ",cpd_type"
      End If
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("cpd_cycle_types ccty", "ccc.cpd_cycle_type", "ccty.cpd_cycle_type", AnsiJoin.AnsiJoinTypes.InnerJoin)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDCycleStatus) Then
        vAnsiJoins.Add("cpd_cycle_statuses ccs", "ccc.cpd_cycle_status", "ccs.cpd_cycle_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If
      If mvParameters.ParameterExists("ForPortal").Bool = True Then
        vWhereFields.Add("ccty.web_publish", CDBField.FieldTypes.cftCharacter, "Y")
        vWhereFields.Add("ccty.start_month", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("ccty.end_month", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vSelectCols, "contact_cpd_cycles ccc", vWhereFields, "start_date DESC,end_date DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("CycleStart") = CDate(vRow.Item("StartDate")).ToString("MMM yyyy")
        vRow.Item("CycleEnd") = CDate(vRow.Item("EndDate")).ToString("MMM yyyy")
      Next
    End Sub
    Private Sub GetContactCreditCardAuthorities(ByVal pDataTable As CDBDataTable)
      Dim vSelAttrs As String = "credit_card_authority_number,authority_type,credit_card_number,start_date,amount,bank_account,order_number,created_by,created_on,cca.amended_by,cca.amended_on,cca.credit_card_details_number,expiry_date,issuer,account_name,ccc.credit_card_type,credit_card_type_desc,cancellation_reason,cancellation_source,cancelled_by,cancelled_on,cca.source,source_desc"
      Dim vAttrs As String = vSelAttrs & If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCCIssueNumber), ",issue_number", ",")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCCIssueNumber) Then vSelAttrs = vSelAttrs & ",issue_number"

      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("CreditCardAuthorityNumber") Then
        vWhereFields.Add("cca.credit_card_authority_number", mvParameters("CreditCardAuthorityNumber").IntegerValue)
      End If
      vWhereFields.Add("cca.contact_number", mvContact.ContactNumber)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contact_credit_cards ccc", "cca.credit_card_details_number", "ccc.credit_card_details_number")
      vAnsiJoins.AddLeftOuterJoin("sources s", "cca.source", "s.source")
      vAnsiJoins.AddLeftOuterJoin("credit_card_types cct", "ccc.credit_card_type", "cct.credit_card_type")
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vSelAttrs), "credit_card_authorities cca", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, ",,,ENCRYPTED_CC_NUMBER")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("AuthorityTypeCode") = vRow.Item("AuthorityType")
        If vRow.Item("AuthorityType") = "C" Then
          vRow.Item("AuthorityType") = "CAF"
        Else
          vRow.Item("AuthorityType") = ""
        End If
      Next
    End Sub
    Private Sub GetContactCreditCards(ByVal pDataTable As CDBDataTable)
      Dim vSelAttrs As String = "credit_card_details_number,credit_card_number,expiry_date,issuer,account_name,credit_card_type_desc,amended_by,amended_on,cct.credit_card_type"
      Dim vAttrs As String = vSelAttrs & If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCCIssueNumber), ",issue_number", ",")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCCIssueNumber) Then vSelAttrs = vSelAttrs & ",issue_number"
      vSelAttrs = vSelAttrs & ",token_id,token_desc"
      vAttrs = vAttrs & ",token_id,token_desc"
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("CreditCardDetailsNumber") Then
        vWhereFields.Add("ccc.credit_card_details_number", mvParameters("CreditCardDetailsNumber").IntegerValue)
      End If
      vWhereFields.Add("contact_number", mvContact.ContactNumber)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("credit_card_types cct", "ccc.credit_card_type", "cct.credit_card_type")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vSelAttrs, "contact_credit_cards ccc", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetContactCreditCustomers(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "cc.contact_number,cc.address_number,cc.company,company_desc,sales_ledger_account,cc.credit_category,credit_category_desc,stop_code,cc.credit_limit,outstanding,on_order,customer_type,terms_number,terms_period,terms_from,last_statement_date,last_statement_closing_balance,last_statement_number,statement_period,cc.amended_by,cc.amended_on,label_name"
      If mvParameters.HasValue("SalesLedgerAccount") Then vWhereFields.Add("sales_ledger_account", mvParameters("SalesLedgerAccount").Value)
      If mvParameters.HasValue("Company") Then vWhereFields.Add("cc.company", mvParameters("Company").Value)
      If mvContact IsNot Nothing Then
        vWhereFields.Add("cc.contact_number", mvContact.ContactNumber)
      Else
        If vWhereFields.Count = 0 Or mvParameters.HasValue("ContactNumber") Then vWhereFields.Add("cc.contact_number", mvParameters("ContactNumber").LongValue)
      End If

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("contacts c", "cc.contact_number", "c.contact_number")
        .Add("companies co", "cc.company", "co.company")
        .Add("credit_categories crc", "cc.credit_category", "crc.credit_category")
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "credit_customers cc", vWhereFields, "cc.company, sales_ledger_account", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactDBANotes(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactDBANotes")
      Dim vAttrs As String = "master,duplicate,merged_on,notes"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM dba_notes WHERE master = " & mvContact.ContactNumber
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactDepartmentHistory(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactDepartmentHistory")
      Dim vRS As CDBRecordSet
      Dim vIndex As Integer
      Dim vSQL As String = "SELECT operation_date, data_values, logname FROM amendment_history WHERE table_name" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, "contacts") & " AND select_1 = " & mvContact.ContactNumber & " AND operation" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, "update") & " ORDER BY operation_date DESC"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      With vRS
        While .Fetch()
          Dim vPos As Integer = InStr(.Fields.Item(2).Value, "NEW")
          If vPos > 0 Then
            Dim vOldValues() As String = Split(Mid$(.Fields.Item(2).Value, 5, (vPos - 7)), Chr(22))
            For vIndex = 0 To UBound(vOldValues)
              Dim vValues() As String = Split(vOldValues(vIndex), ":")
              If vValues(0) = "department" Then
                pDataTable.AddRowFromList(mvContact.ContactNumber.ToString & "," & vValues(1) & "," & mvEnv.GetDescription("departments", "department", vValues(1)) & ", " & CDate(.Fields.Item(1).Value).ToString(CAREDateFormat) & ", " & .Fields.Item(3).Value)
              End If
            Next
          End If
        End While
        .CloseRecordSet()
      End With
    End Sub
    Private Sub GetContactDepartmentNotes(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactDepartmentNotes")
      pDataTable.FillFromSQLDONOTUSE(mvEnv, "SELECT notes,amended_by,amended_on FROM department_notes WHERE unique_id = " & mvContact.ContactNumber & " AND record_type = '" & If(mvContact.ContactType = Contact.ContactTypes.ctcOrganisation, "O", "C") & "' AND department = '" & mvEnv.User.Department & "'")
    End Sub
    Private Sub GetContactDespatchNotes(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactDespatchNotes")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ dn.batch_number,dn.transaction_number,dn.despatch_note_number,dn.picking_list_number,despatch_date,despatch_method_desc," & mvContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry) & ",carrier_reference FROM despatch_notes dn, despatch_methods dm, contacts c, addresses a, countries co WHERE dn.contact_number = " & mvContact.ContactNumber & " AND dn.despatch_method = dm.despatch_method AND dn.contact_number = c.contact_number AND dn.address_number = a.address_number AND a.country = co.country ORDER BY dn.batch_number DESC, dn.transaction_number DESC, despatch_date DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "batch_number,DISTINCT_DESPATCH_TRANSACTION,despatch_note_number,s.picking_list_number,despatch_date,despatch_method_desc,CONTACT_NAME,ADDRESS_LINE,carrier_reference")
    End Sub
    Private Sub GetContactDirectDebits(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "dd.direct_debit_number,start_date,ca.sort_code,ca.account_number,amount,reference,dd.bank_account,bank_account_desc,order_number,emandate_created,auddis_cancel_notified,mandate_type,dd.source,source_desc,created_by,created_on,dd.amended_by,dd.amended_on,account_name,cancellation_reason,cancellation_source,cancelled_by,cancelled_on,dd.bank_details_number"
      vAttrs &= If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers), ",ca.iban_number,ca.bic_code,date_signed,bank_details_changed,previous_bank_details_number", ",,,,,")
      vAttrs &= ",,,,,," & If(mvEnv.DefaultCountry = "CH" OrElse mvEnv.DefaultCountry = "NL", "Text1,Text2,Text3,Text4,Text5", ",,,,")

      Dim vOrderBy As String = "cancelled_on " & mvEnv.Connection.DBSortByNullsFirst & ", start_date DESC"

      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("bank_accounts ba", "dd.bank_account", "ba.bank_account")
        .Add("contact_accounts ca", "dd.bank_details_number", "ca.bank_details_number")
        .Add("sources s", "dd.source", "s.source")
        If mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL" Then
          .AddLeftOuterJoin("direct_debit_references ddr", "dd.direct_debit_number", "ddr.direct_debit_number")
        End If
      End With

      Dim vWhereFields As New CDBFields()
      If mvParameters.Exists("DirectDebitNumber") Then vWhereFields.Add("dd.direct_debit_number", mvParameters("DirectDebitNumber").IntegerValue)
      vWhereFields.Add("dd.contact_number", mvContact.ContactNumber)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs).Replace("reference", mvEnv.Connection.DBSpecialCol("", "reference")), "direct_debits dd", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)

      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetBankInfo(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("MandateTypeCode") = vRow.Item("MandateType")
        If vRow.Item("MandateType") = "P" Then                'Update the mandate_type heading
          vRow.Item("MandateType") = "Paperless"
        ElseIf vRow.Item("MandateType") = "W" Then
          vRow.Item("MandateType") = "Written"
        End If
      Next
    End Sub
    Private Sub GetContactEMailAddresses(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEMailAddresses")
      Dim vAttrs As String = "co.contact_number,co.address_number,co.device,device_desc,co.notes,co.amended_by,co.amended_on,communication_number,email_address,auto_email,valid_from,valid_to,is_active,mail,device_default,preferred_method"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace$(RemoveBlankItems(vAttrs), "email_address", mvEnv.Connection.DBSpecialCol("", "number") & " AS email_address")
      vSQL = vSQL & " FROM communications co, devices d WHERE "
      'If mvType = DataSelectionTypes.dstContactCommsNumbers Then vAttr = ",,"
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = vSQL & "address_number IN (select address_number FROM organisation_addresses WHERE organisation_number = " & mvContact.ContactNumber & ") AND contact_number IS NULL"
      Else
        vSQL = vSQL & "contact_number = " & mvContact.ContactNumber
      End If
      vSQL = vSQL & " AND co.device = d.device AND d.email = 'Y'"
      vSQL = vSQL & " AND is_active = 'Y'"
      vSQL = vSQL & " ORDER BY preferred_method DESC, device_default DESC, is_active DESC, d.sequence_number, device_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("IsActive")
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("DeviceDefault")
        vRow.SetYNValue("PreferredMethod")
      Next
    End Sub
    Private Sub GetContactEventBookings(ByVal pDataTable As CDBDataTable)
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vWhereFields As CDBFields = vEventInfo.WhereFields
      Dim vTransAttrs As String = ",t.batch_number,t.transaction_number,t.line_number"
      Dim vInvAttr As String = ",i.invoice_pay_status,ips.invoice_pay_status_desc,iph.invoice_allocation_amount"
      If mvParameters.ContainsKey("BatchNumber") Then vWhereFields.Add("eb2.batch_number", CARE.Data.CDBField.FieldTypes.cftInteger, mvParameters("BatchNumber").Value)
      If mvParameters.ContainsKey("TransactionNumber") Then vWhereFields.Add("eb2.transaction_number", CARE.Data.CDBField.FieldTypes.cftInteger, mvParameters("TransactionNumber").Value)

      Dim vAttrs As String = "event_reference,e.event_number,event_desc,e.start_date,t.booking_number,t.booking_date,t.quantity,t.booking_status,ebo.option_number,option_desc,s.subject,subject_desc,s.skill_level,skill_level_desc,e.venue,venue_desc,s.location,e.event_class,t.cancellation_reason,t.cancelled_by,t.cancelled_on"
      vAttrs &= ",t.adult_quantity,t.child_quantity,s.start_time,s.end_date,s.end_time,i.invoice_number,reprint_count,batch_type,{0},ebo.product,t.rate,eb2.notes"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity) = False Then vAttrs = vAttrs.Replace("t.adult_quantity", "").Replace("t.child_quantity", "")
      Dim vRSAttrs As String = String.Format(vAttrs, "bt.amount")
      vAttrs &= ",pc.contact_number AS payer_contact_number,pc.label_name AS payer_label_name"
      vRSAttrs &= ",payer_contact_number,payer_label_name"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then vAttrs = vAttrs.Replace("s.start_time", mvEnv.Connection.DBIsNull("t.start_time", "s.start_time") & " AS start_time").Replace("s.end_time", mvEnv.Connection.DBIsNull("t.end_time", "s.end_time") & " AS end_time")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
        vAttrs = String.Format(vAttrs, "(" & mvEnv.Connection.DBIsNull("bta.amount", "0") & " + " & mvEnv.Connection.DBIsNull("bta2.amount", "0") & " + " & mvEnv.Connection.DBIsNull("ebtfac.amount", "0") & " + " & mvEnv.Connection.DBIsNull("ebtfad.amount", "0") & ") AS amount")
      Else
        vAttrs = String.Format(vAttrs, mvEnv.Connection.DBIsNull("bta.amount", "0"))
      End If
      AddOwnerRestrictionToFields(vWhereFields)

      'BTA line-types 'X' (exclude reversals)
      Dim vAnsiJoinsX As New AnsiJoins
      With vAnsiJoinsX
        .Add("event_booking_transactions ebt", "eb.event_number", "ebt.event_number", "eb.booking_number", "ebt.booking_number")
        .Add("batch_transaction_analysis bta", "ebt.batch_number", "bta.batch_number", "ebt.transaction_number", "bta.transaction_number", "ebt.line_number", "bta.line_number")
        .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
        .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
      End With
      Dim vWhereFieldsX As New CDBFields(New CDBField("bta.line_type", "X"))
      With vWhereFieldsX
        .AddJoin("eb.batch_number", "bta.batch_number")
        .AddJoin("eb.transaction_number", "bta.transaction_number")
      End With
      Dim vSQLX As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, SUM(bta.amount) AS amount", "event_bookings eb", vWhereFieldsX, "", vAnsiJoinsX)
      vSQLX.GroupBy = "eb.event_number, eb.booking_number"

      'Financial Adjustments - Credits & Debits
      Dim vFASQLC As String = ""
      Dim vFASQLD As String = ""

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
        Dim vAnsiJoinsFA As New AnsiJoins
        With vAnsiJoinsFA
          .Add("event_booking_transactions ebt", "eb.event_number", "ebt.event_number", "eb.booking_number", "ebt.booking_number")
          .Add("batch_transaction_analysis bta", "ebt.batch_number", "bta.batch_number", "ebt.transaction_number", "bta.transaction_number", "ebt.line_number", "bta.line_number")
          .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
          .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
        End With
        Dim vWhereFieldsFAC As New CDBFields(New CDBField("bta.member_number", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual))
        With vWhereFieldsFAC
          .Add("bta.member_number#2", CDBField.FieldTypes.cftInteger, mvEnv.Connection.DBToString("ebt.booking_number"))
          .Add("eb.batch_number", CDBField.FieldTypes.cftInteger, "bta.batch_number", CDBField.FieldWhereOperators.fwoNotEqual)
        End With
        Dim vWhereFieldsFAD As New CDBFields
        vWhereFieldsFAD.Clone(vWhereFieldsFAC)
        vWhereFieldsFAC.Add("transaction_sign", "C")
        vWhereFieldsFAD.Add("transaction_sign", "D")
        'Credits
        Dim vSQLFAC As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, transaction_sign, SUM(bta.amount) AS amount", "event_bookings eb", vWhereFieldsFAC, "", vAnsiJoinsFA)
        vSQLFAC.GroupBy = "eb.event_number, eb.booking_number, transaction_sign"
        vFASQLC = vSQLFAC.SQL
        'Debits
        Dim vAnsiJoinsFAD As New AnsiJoins
        For Each vJoin As AnsiJoin In vAnsiJoinsFA
          vAnsiJoinsFAD.Add(vJoin)
        Next
        vAnsiJoinsFAD.AddLeftOuterJoin("reversals r", "bta.batch_number", "r.batch_number", "bta.transaction_number", "r.transaction_number", "bta.line_number", "r.line_number")
        vWhereFieldsFAD.Add("r.was_batch_number", "")
        Dim vSQLFAD As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, transaction_sign, SUM(bta.amount * -1) AS amount", "event_bookings eb", vWhereFieldsFAD, "", vAnsiJoinsFAD)
        vSQLFAD.GroupBy = "eb.event_number, eb.booking_number, transaction_sign"
        vFASQLD = vSQLFAD.SQL
      End If

      'BR17149/17420: Display Invoice Allocations - required for Amend Event Booking user dialogue
      Dim vAnsiJoinsIPH As New AnsiJoins
      With vAnsiJoinsIPH
        .Add("invoice_payment_history iph", "i.invoice_number", "iph.invoice_number")
        .AddLeftOuterJoin("invoices i2", "iph.batch_number", "i2.batch_number", "iph.transaction_number", "i2.transaction_number")
      End With
      Dim vSQLIPH As New SQLStatement(mvEnv.Connection, "SUM(iph.amount) AS invoice_allocation_amount,i.batch_number,i.transaction_number", "invoices i", New CDBFields(New CDBField("i2.record_type", "N", CDBField.FieldWhereOperators.fwoNullOrNotEqual)), "", vAnsiJoinsIPH)
      vSQLIPH.GroupBy = "i.batch_number,i.transaction_number"

      'Main SQL
      Dim vSQLAttrs As String = RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & vTransAttrs
      Dim vTable As String = vEventInfo.Table1.Replace(",", "")
      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        If vTable.Length > 0 Then
          .Add("event_bookings t", "cp.contact_number", "t.contact_number") 'BR19676
        Else
          vTable = "event_bookings t"
        End If
        .Add("event_bookings eb2", "t.event_number", "eb2.event_number", "t.booking_number", "eb2.booking_number")
        .Add("contacts c", "t.contact_number", "c.contact_number")
        .Add("events e", "t.event_number", "e.event_number")
        .Add("event_booking_options ebo", "t.option_number", "ebo.option_number")
        .Add("sessions s", "e.event_number", "s.event_number")
        .Add("subjects su", "s.subject", "su.subject")
        .Add("skill_levels sl", "s.skill_level", "sl.skill_level")
        .Add("venues v", "e.venue", "v.venue")
        .AddLeftOuterJoin("batch_transaction_analysis bta", "eb2.batch_number", "bta.batch_number", "eb2.transaction_number", "bta.transaction_number", "eb2.line_number", "bta.line_number")
        .AddLeftOuterJoin("invoices i", "t.batch_number", "i.batch_number", "t.transaction_number", "i.transaction_number")
        .AddLeftOuterJoin("(" & vSQLIPH.SQL & ") iph", "i.batch_number", "iph.batch_number", "i.transaction_number", "iph.transaction_number")
        .AddLeftOuterJoin("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
        .AddLeftOuterJoin("batches b", "t.batch_number", "b.batch_number")
        .AddLeftOuterJoin("batch_transactions bt", "t.batch_number", "bt.batch_number", "t.transaction_number", "bt.transaction_number")
        .AddLeftOuterJoin("contacts pc", "bt.contact_number", "pc.contact_number")

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
          .AddLeftOuterJoin("(" & vSQLX.SQL & ") bta2", "eb2.event_number", "bta2.event_number", "eb2.booking_number", "bta2.booking_number")
          If vFASQLC.Length > 0 Then .AddLeftOuterJoin("(" & vFASQLC & ") ebtfac", "eb2.event_number", "ebtfac.event_number", "eb2.booking_number", "ebtfac.booking_number")
          If vFASQLD.Length > 0 Then .AddLeftOuterJoin("(" & vFASQLD & ") ebtfad", "eb2.event_number", "ebtfad.event_number", "eb2.booking_number", "ebtfad.booking_number")
        End If
      End With
      vWhereFields.Add("s.session_type", "0")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & vTransAttrs & vInvAttr, vTable, vWhereFields, "e.start_date DESC,e.event_number" & vEventInfo.ContactSort, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vRSAttrs & vEventInfo.ContactCols & vTransAttrs & vInvAttr, "BOOKING_STATUS_DESC,,")

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("CreditSale") = If(vRow.Item("BatchType") = "CS", "Y", "N")
        If Len(vRow.Item("InvoiceRePrintCount")) > 0 Then
          If Val(vRow.Item("InvoiceNumber")) = 0 Or (Val(vRow.Item("InvoiceNumber")) > 0 And Val(vRow.Item("InvoiceRePrintCount")) < 0) Then
            vRow.Item("InvoicePrinted") = "N"     'Invoice not yet printed
          Else
            vRow.Item("InvoicePrinted") = "Y"     'Invoice printed
          End If
        Else
          vRow.Item("InvoicePrinted") = "N"
        End If
        vRow.SetYNValue("CreditSale", False, True)
        vRow.SetYNValue("InvoicePrinted", False, True)
      Next
      If mvParameters.OptionalValue("SmartClient", "N") = "Y" Then
        pDataTable.SuppressDuplicateColumnData("EventDesc", "EventNumber")
      End If
    End Sub
    Private Sub GetContactEventDelegates(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As New StringBuilder
      vAttrs.Append("event_reference,e.event_number,event_desc,e.start_date,t.booking_number,attended,booking_status,candidate_number")
      vAttrs.Append(",eb.option_number,option_desc,s.subject,subject_desc,s.skill_level,skill_level_desc,e.venue,venue_desc,s.location,e.event_class")
      vAttrs.Append(",pledged_amount,donation_total,sponsorship_total,booking_payment_amount,other_payments_total,sequence_number")
      Dim vFieldNames As New StringBuilder
      vFieldNames.Append(vAttrs.ToString)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDelegateSequenceNumber) = False Then vFieldNames = vFieldNames.Replace("sequence_number", "")
      vFieldNames.Append(",ebc.label_name AS ebc_label_name,ebc.contact_number AS ebc_contact_number,eb.batch_number,eb.transaction_number,eb.line_number,bta.source,event_delegate_number")
      vFieldNames.Append(",pc.contact_number AS payer_contact_number,pc.label_name AS payer_label_name")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact, "delegates t", "contacts c")
      vFieldNames.Append(vEventInfo.ContactAttrs)
      Dim vTableName As String = "delegates t"
      If vEventInfo.Table1.Length > 0 Then vTableName = vEventInfo.Table1
      Dim vAnsiJoins As AnsiJoins = vEventInfo.AnsiJoins
      vAnsiJoins.Add("events e", "t.event_number", "e.event_number")
      vAnsiJoins.Add("event_bookings eb", "t.booking_number", "eb.booking_number")
      vAnsiJoins.Add("contacts ebc", "eb.contact_number", "ebc.contact_number")
      vAnsiJoins.Add("event_booking_options ebo", "eb.option_number", "ebo.option_number")
      vAnsiJoins.Add("sessions s", "e.event_number", "s.event_number")
      vAnsiJoins.Add("subjects su", "s.subject", "su.subject")
      vAnsiJoins.Add("skill_levels sl", "s.skill_level", "sl.skill_level")
      vAnsiJoins.Add("venues v", "e.venue", "v.venue")
      vAnsiJoins.AddLeftOuterJoin("batch_transaction_analysis bta", "eb.batch_number", "bta.batch_number", "eb.transaction_number", "bta.transaction_number", "eb.line_number", "bta.line_number")
      vAnsiJoins.AddLeftOuterJoin("batch_transactions bt", "eb.batch_number", "bt.batch_number", "eb.transaction_number", "bt.transaction_number")
      vAnsiJoins.AddLeftOuterJoin("contacts pc", "bt.contact_number", "pc.contact_number")
      Dim vWhereFields As CDBFields = vEventInfo.WhereFields
      vWhereFields.Add("s.session_type", "0")
      If mvParameters.HasValue("EventDelegateNumber") Then vWhereFields.Add("t.event_delegate_number", mvParameters("EventDelegateNumber").IntegerValue)
      AddOwnerRestrictionToFields(vWhereFields)
      Dim vOrderBy As String = "e.start_date DESC, e.event_number" & vEventInfo.ContactSort
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFieldNames.ToString), vTableName, vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs.ToString & ",ebc_label_name,ebc_contact_number,batch_number,transaction_number,line_number,source,payer_contact_number,payer_label_name,event_delegate_number" & vEventInfo.ContactCols, "BOOKING_STATUS_DESC")
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Attended")
      Next
      pDataTable.SuppressDuplicateColumnData("EventDesc")
    End Sub
    Private Sub GetContactEventOrganiser(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEventOrganiser")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vAttrs As String = "event_reference,e.event_number,event_desc,e.start_date,t.organiser,organiser_desc,eo.notes,e.venue,venue_desc,eo.reference,s.location,e.event_class"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then vAttrs = Replace$(vAttrs, "e.event_class", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & " FROM " & vEventInfo.Table1
      vSQL = vSQL & "organisers t, contacts c, event_organisers eo, events e, venues v, sessions s"
      vSQL = vSQL & " WHERE " & vEventInfo.Where & " AND t.organiser = eo.organiser AND eo.event_number = e.event_number AND e.venue = v.venue AND s.event_number = e.event_number AND s.session_type = '0' "
      AddOwnerRestriction(vSQL)
      vSQL = vSQL & " ORDER BY e.start_date DESC, e.event_number" & vEventInfo.ContactSort
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & vEventInfo.ContactCols)
      pDataTable.SuppressDuplicateColumnData("EventDesc")
    End Sub
    Private Sub GetContactEventPersonnel(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEventPersonnel")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vAttrs As String = "event_reference,e.event_number,t.session_number,event_desc,session_desc,t.start_date,task,t.end_date,t.start_time,t.end_time,e.venue,venue_desc,s.location,e.event_class"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then vAttrs = Replace$(vAttrs, "e.event_class", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & " FROM " & vEventInfo.Table1
      vSQL = vSQL & "event_personnel t, contacts c, sessions s, events e, venues v"
      vSQL = vSQL & " WHERE " & vEventInfo.Where & " AND t.session_number = s.session_number AND s.event_number = e.event_number AND e.venue = v.venue"
      AddOwnerRestriction(vSQL)
      vSQL = vSQL & " ORDER BY t.start_date DESC, e.event_number" & vEventInfo.ContactSort
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & vEventInfo.ContactCols)
      pDataTable.SuppressDuplicateColumnData("EventDesc")
    End Sub
    Private Sub GetContactEventRoomBookings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEventRoomBookings")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vTransAttrs As String = ",batch_number,transaction_number,line_number"
      Dim vAttrs As String = "event_reference,e.event_number,event_desc,e.start_date,t.room_booking_number,t.booked_date,t.number_of_rooms,room_type_desc,t.from_date,t.to_date,booking_status,name,t.confirmed_date,t.notes,t.cancellation_reason,e.event_class"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then vAttrs = Replace$(vAttrs, "e.event_class", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & vTransAttrs & " FROM " & vEventInfo.Table1
      vSQL = vSQL & "contact_room_bookings t, contacts c, events e, room_block_bookings rbb, room_types rt, organisations o"
      vSQL = vSQL & " WHERE " & vEventInfo.Where & " AND t.event_number = e.event_number "
      AddOwnerRestriction(vSQL)
      vSQL = vSQL & " AND t.block_booking_number = rbb.block_booking_number AND rbb.room_type = rt.room_type AND rbb.organisation_number = o.organisation_number"
      vSQL = vSQL & " ORDER BY e.start_date DESC, e.event_number" & vEventInfo.ContactSort
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & vEventInfo.ContactCols & vTransAttrs, "BOOKING_STATUS_DESC")
      pDataTable.SuppressDuplicateColumnData("EventDesc")
    End Sub
    Private Sub GetContactEventRoomsAllocated(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEventRoomsAllocated")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vAttrs As String = "event_reference,e.event_number,t.room_id,event_desc,e.start_date,t.room_booking_number,room_type_desc,room_date,name,t.notes,e.event_class"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then vAttrs = Replace$(vAttrs, "e.event_class", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & " FROM " & vEventInfo.Table1
      vSQL = vSQL & "room_booking_links t, contacts c, contact_room_bookings crb, events e, room_block_bookings rbb, room_types rt, organisations o"
      vSQL = vSQL & " WHERE " & vEventInfo.Where & " AND t.room_booking_number = crb.room_booking_number AND crb.event_number = e.event_number "
      AddOwnerRestriction(vSQL)
      vSQL = vSQL & " AND crb.block_booking_number = rbb.block_booking_number AND rbb.room_type = rt.room_type AND rbb.organisation_number = o.organisation_number"
      vSQL = vSQL & " ORDER BY e.start_date DESC, e.event_number,room_date,rbb.room_type" & vEventInfo.ContactSort
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & vEventInfo.ContactCols)
      pDataTable.SuppressDuplicateColumnData("EventDesc")
    End Sub
    Private Sub GetContactEventSessions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactEventSessions")
      Dim vEventInfo As New EventSelectionInfo(mvEnv, mvContact)
      Dim vAttrs As String = "event_reference,e.event_number,t.booking_number,sb.session_number,event_desc,session_desc,s.start_date,s.start_time,s.end_date,s.end_time,s.subject,subject_desc,s.skill_level,skill_level_desc,v.venue,venue_desc,s.location,e.event_class,ds.attended"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then vAttrs = Replace$(vAttrs, "e.event_class", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then vAttrs = Replace$(vAttrs, "ds.attended", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & vEventInfo.ContactAttrs & " FROM " & vEventInfo.Table1
      vSQL = vSQL & "delegates t, contacts c, events e, session_bookings sb, "
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then vSQL = vSQL & "delegate_sessions ds, "
      vSQL = vSQL & " sessions s, subjects su, skill_levels sl, venues v"
      vSQL = vSQL & " WHERE " & vEventInfo.Where & " AND t.event_number = e.event_number "
      AddOwnerRestriction(vSQL)
      vSQL = vSQL & " AND t.booking_number = sb.booking_number "
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
        vSQL = vSQL & " AND ds.event_delegate_number = t.event_delegate_number AND ds.session_number = sb.session_number "
      End If
      vSQL = vSQL & " AND sb.session_number = s.session_number"
      vSQL = vSQL & " AND s.subject = su.subject AND s.skill_level = sl.skill_level AND e.venue = v.venue"
      vSQL = vSQL & " ORDER BY e.start_date DESC, e.event_number, s.start_date, s.start_time" & vEventInfo.ContactSort
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & vEventInfo.ContactCols)
      pDataTable.SuppressDuplicateColumnData("EventDesc")
      pDataTable.SuppressDuplicateColumnData("SessionDesc")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          vRow.SetYNValue("Attended")
        Next
      End If
    End Sub
    Private Sub GetContactExternalReferences(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "cel.data_source,ds.data_source_desc,cel.external_reference,cel.amended_on,cel.amended_by"
      Dim vWhereFields As New CDBFields()
      If Not mvParameters.Exists("UseContactRestriction") OrElse mvParameters("UseContactRestriction").Bool Then vWhereFields.Add("contact_number", mvContact.ContactNumber)
      If mvParameters.HasValue("DataSource") Then
        vWhereFields.Add("cel.data_source", mvParameters("DataSource").Value)
      Else
        vWhereFields.Add("ds.maintenance_available", "Y")
      End If
      If mvParameters.HasValue("ExternalReference") Then vWhereFields.Add("cel.external_reference", mvParameters("ExternalReference").Value)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("data_sources ds", "cel.data_source", "ds.data_source")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "contact_external_links cel", vWhereFields, "data_source_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetContactFinLinksDonated(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "fl.batch_number,fl.transaction_number,fl.line_number,c.contact_number,CONTACT_NAME,amount,product_desc,rate_desc,quantity,source_desc,fhd.status,vat_rate,vat_amount"
      Dim vFields As String = vAttrs & ",payee_contact_number,PAYEE_CONTACT_NAME,fl2.line_type"
      vAttrs = vAttrs.Replace("c.contact_number,CONTACT_NAME", mvContact.GetRecordSetFieldsName("c"))
      vAttrs &= "," & mvContact.GetRecordSetFieldsName("c2", "payee") & ",fl2.line_type"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vAttrs = vAttrs & ",currency_code"
        vFields &= ",currency_code"
      End If
      Dim vInnerSQL As String = "(SELECT batch_number, transaction_number, line_number, line_type, contact_number FROM financial_links fl WHERE line_type <> '" & mvParameters("LineType").Value & "') fl2"

      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("contacts c", "fl.contact_number", "c.contact_number")
        .Add("financial_history_details fhd", "fl.batch_number", "fhd.batch_number", "fl.transaction_number", "fhd.transaction_number", "fl.line_number", "fhd.line_number")
        .Add("products p", "fhd.product", "p.product")
        .Add("rates r", "fhd.product", "r.product", "fhd.rate", "r.rate")
        .Add("sources s", "fhd.source", "s.source")
        .AddLeftOuterJoin(vInnerSQL, "fl.batch_number", "fl2.batch_number", "fl.transaction_number", "fl2.transaction_number", "fl.line_number", "fl2.line_number")
        .AddLeftOuterJoin("contacts c2", "fl2.contact_number", "c2.contact_number")
      End With

      Dim vWhereFields As New CDBFields()
      With vWhereFields
        .Add("fl.donor_contact_number", mvContact.ContactNumber)
        .Add("fl.line_type", mvParameters("LineType").Value)
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "financial_links fl", vWhereFields, "c.surname, c.title, c.initials, fl.batch_number, fl.transaction_number, fl.line_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)

      With vWhereFields
        .Clear()
        .Add("table_name", "financial_links")
        .Add("attribute_name", "line_type")
      End With
      Dim vLinkNameSQL As New SQLStatement(mvEnv.Connection, "lookup_code,lookup_desc", "maintenance_lookup", vWhereFields)
      Dim vLinkNameRS As CDBRecordSet = vLinkNameSQL.GetRecordSet()
      Dim vLinkNames As New CollectionList(Of LookupItem)
      While vLinkNameRS.Fetch
        vLinkNames.Add(vLinkNameRS.Fields(1).Value, New LookupItem(vLinkNameRS.Fields(1).Value, vLinkNameRS.Fields(2).Value))
      End While
      vLinkNameRS.CloseRecordSet()

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetDescriptionFromCode("CreditType", "CreditType", vLinkNames)
      Next

    End Sub
    Private Sub GetContactFinLinksReceived(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "fl.batch_number,fl.transaction_number,fl.line_number,c.contact_number,CONTACT_NAME,amount,product_desc,rate_desc,quantity,source_desc,fhd.status,vat_rate,vat_amount"
      Dim vFields As String = vAttrs & ",payee_contact_number,PAYEE_CONTACT_NAME,fl2.line_type"
      vAttrs = vAttrs.Replace("c.contact_number,CONTACT_NAME", mvContact.GetRecordSetFieldsName("c"))
      vAttrs &= "," & mvContact.GetRecordSetFieldsName("c2", "payee") & ",fl2.line_type"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vAttrs = vAttrs & ",currency_code"
        vFields &= ",currency_code"
      End If
      Dim vInnerSQL As String = "(SELECT batch_number, transaction_number, line_number, line_type, contact_number FROM financial_links fl WHERE line_type <> '" & mvParameters("LineType").Value & "') fl2"

      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("contacts c", "fl.donor_contact_number", "c.contact_number")
        .Add("financial_history_details fhd", "fl.batch_number", "fhd.batch_number", "fl.transaction_number", "fhd.transaction_number", "fl.line_number", "fhd.line_number")
        .Add("products p", "fhd.product", "p.product")
        .Add("rates r", "fhd.product", "r.product", "fhd.rate", "r.rate")
        .Add("sources s", "fhd.source", "s.source")
        .AddLeftOuterJoin(vInnerSQL, "fl.batch_number", "fl2.batch_number", "fl.transaction_number", "fl2.transaction_number", "fl.line_number", "fl2.line_number")
        .AddLeftOuterJoin("contacts c2", "fl2.contact_number", "c2.contact_number")
      End With

      Dim vWhereFields As New CDBFields()
      With vWhereFields
        .Add("fl.contact_number", mvContact.ContactNumber)
        .Add("fl.line_type", mvParameters("LineType").Value)
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "financial_links fl", vWhereFields, "c.surname, c.title, c.initials, fl.batch_number, fl.transaction_number, fl.line_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)

      With vWhereFields
        .Clear()
        .Add("table_name", "financial_links")
        .Add("attribute_name", "line_type")
      End With
      Dim vLinkNameSQL As New SQLStatement(mvEnv.Connection, "lookup_code,lookup_desc", "maintenance_lookup", vWhereFields)
      Dim vLinkNameRS As CDBRecordSet = vLinkNameSQL.GetRecordSet()
      Dim vLinkNames As New CollectionList(Of LookupItem)
      While vLinkNameRS.Fetch
        vLinkNames.Add(vLinkNameRS.Fields(1).Value, New LookupItem(vLinkNameRS.Fields(1).Value, vLinkNameRS.Fields(2).Value))
      End While
      vLinkNameRS.CloseRecordSet()

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetDescriptionFromCode("CreditType", "CreditType", vLinkNames)
      Next

    End Sub
    Private Sub GetContactFundraisingEvents(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "cfe.contact_fundraising_number,cfe.contact_number,fundraising_description,cfe.source,source_desc,target_amount,target_date,cfe.event_number,event_desc,web_page_number,cfe.amended_by,cfe.amended_on,donation_total,gift_aid_total,thank_you_message"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataThankYouMessage) = False Then vFields = vFields.Replace("thank_you_message", "")
      Dim vAttrs As String = RemoveBlankItems(vFields)

      Dim vWhereFields As New CDBFields()
      If mvContact.ContactNumber > 0 Then
        vWhereFields.Add("cfe.contact_number", mvContact.ContactNumber)
      Else
        AddWhereFieldFromIntegerParameter(vWhereFields, "WebPageNumber", "web_page_number")
      End If
      AddWhereFieldFromIntegerParameter(vWhereFields, "ContactFundraisingNumber", "cfe.contact_fundraising_number")
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("sources s", "cfe.source", "s.source")
      vAnsiJoins.AddLeftOuterJoin("events e", "cfe.event_number", "e.event_number")
      Dim vSub1AnsiJoins As New AnsiJoins()
      vSub1AnsiJoins.AddLeftOuterJoin("fundraising_event_analysis fea", "cfe.contact_fundraising_number", "fea.contact_fundraising_number")
      vSub1AnsiJoins.AddLeftOuterJoin("batch_transaction_analysis bta", "fea.batch_number", "bta.batch_number", "fea.transaction_number", "bta.transaction_number", "fea.line_number", "bta.line_number")
      Dim vSub1Statement As New SQLStatement(mvEnv.Connection, "cfe.contact_fundraising_number, SUM(amount) AS donation_total", "contact_fundraising_events cfe", vWhereFields, "", vSub1AnsiJoins)
      vSub1Statement.GroupBy = "cfe.contact_fundraising_number"
      Dim vSub2AnsiJoins As New AnsiJoins()
      vSub2AnsiJoins.Add("fundraising_event_analysis fea", "cfe.contact_fundraising_number", "fea.contact_fundraising_number")
      vSub2AnsiJoins.Add("declaration_lines_unclaimed dlu", "fea.batch_number", "dlu.batch_number", "fea.transaction_number", "dlu.transaction_number", "fea.line_number", "dlu.line_number")
      vSub2AnsiJoins.Add("batch_transactions bt", "fea.batch_number", "bt.batch_number", "fea.transaction_number", "bt.transaction_number")
      vSub2AnsiJoins.Add("tax_rates tr ", "tr.date_from <", "bt.transaction_date", "tr.date_to >", "bt.transaction_date")
      Dim vSub2Statement As New SQLStatement(mvEnv.Connection, "cfe.contact_fundraising_number, SUM((net_amount / (100 - tax_percent) * 100) - net_amount) AS gift_aid_total", "contact_fundraising_events cfe", vWhereFields, "", vSub2AnsiJoins)
      vSub2Statement.GroupBy = "cfe.contact_fundraising_number"

      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) dt", vSub1Statement.SQL), "cfe.contact_fundraising_number", "dt.contact_fundraising_number")
      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) gat", vSub2Statement.SQL), "cfe.contact_fundraising_number", "gat.contact_fundraising_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_fundraising_events cfe", vWhereFields, "target_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetContactFundraisingEventFinder(ByVal pDataTable As CDBDataTable)
      Dim vEventsOnly As Boolean = False
      Dim vFields As String = "contact_fundraising_number,fundraising_description,cfe.source,s.source_desc,target_amount,target_date,cfe.event_number,event_desc,event_reference,cfe.web_page_number,cfe.amended_by,cfe.amended_on,"
      Dim vWhereFields As New CDBFields()
      AddWhereFieldFromParameter(vWhereFields, "Surname", "surname")
      AddWhereFieldFromParameter(vWhereFields, "Forenames", "forenames")
      AddWhereFieldFromParameter(vWhereFields, "Town", "town")
      AddWhereFieldFromIntegerParameter(vWhereFields, "EventNumber", "cfe.event_number")
      AddWhereFieldFromParameter(vWhereFields, "FundraisingDescription", "fundraising_description")
      AddWhereFieldFromParameter(vWhereFields, "EventDesc", "event_desc")
      AddWhereFieldFromParameter(vWhereFields, "EventReference", "event_reference")
      AddWhereFieldFromDateParameter(vWhereFields, "TargetDate", "target_date")
      AddWhereFieldFromParameter(vWhereFields, "Venue", "evb.venue")
      AddWhereFieldFromParameter(vWhereFields, "Organiser", "eo.organiser")
      AddWhereFieldFromParameter(vWhereFields, "SkillLevel", "skill_level")
      AddWhereFieldFromParameter(vWhereFields, "Topic", "et.topic")
      AddWhereFieldFromParameter(vWhereFields, "EventGroup", "event_group")
      AddWhereFieldFromParameter(vWhereFields, "Branch", "e.branch")
      AddWhereFieldFromParameter(vWhereFields, "DistributionCode", "so.distribution_code")
      vWhereFields.Add("cfe.web_page_number", 0, CDBField.FieldWhereOperators.fwoGreaterThan)
      vWhereFields.Add("page_published", "Y")

      If mvParameters.ContainsKey("EventDesc") OrElse mvParameters.ContainsKey("EventReference") OrElse mvParameters.ContainsKey("Venue") OrElse mvParameters.ContainsKey("Organiser") OrElse mvParameters.ContainsKey("SkillLevel") OrElse mvParameters.ContainsKey("Topic") OrElse mvParameters.ContainsKey("EventGroup") OrElse mvParameters.ContainsKey("Branch") OrElse mvParameters.ContainsKey("DistributionCode") Then
        'For these parameters we must only select Events
        vEventsOnly = True
      End If

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "cfe.contact_number", "c.contact_number")
      'vAnsiJoins.Add("contact_addresses ca", "c.contact_number", "ca.contact_number")
      vAnsiJoins.Add("addresses a", "c.address_number", "a.address_number")
      vAnsiJoins.Add("web_pages wp", "cfe.web_page_number", "wp.web_page_number")
      If vEventsOnly Then
        'We have specified events data so we only want Events that match the criteria
        vAnsiJoins.Add("events e", "cfe.event_number", "e.event_number")
        If mvParameters.ContainsKey("Venue") Then vAnsiJoins.Add("event_venue_bookings evb", "e.event_number", "evb.event_number")
        If mvParameters.ContainsKey("Organiser") Then vAnsiJoins.Add("event_organisers eo", "e.event_number", "eo.event_number")
        If mvParameters.ContainsKey("Topic") Then vAnsiJoins.Add("event_topics et", "e.event_number", "et.event_number")
        If mvParameters.ContainsKey("DistributionCode") Then
          vAnsiJoins.Add("event_sources es", "e.event_number", "es.event_number")
          vAnsiJoins.Add("sources so", "es.source", "so.source")
        End If
      End If
      vAnsiJoins.AddLeftOuterJoin("sources s", "cfe.source", "s.source")
      If vEventsOnly = False Then vAnsiJoins.AddLeftOuterJoin("events e", "cfe.event_number", "e.event_number")
      Dim vContact As New Contact(mvEnv)
      vContact.Init()
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields & vContact.GetRecordSetFieldsName, "contact_fundraising_events cfe", vWhereFields, "target_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields & ",,contact_number", ContactNameItems())
      Dim vWC As New WebControl(mvEnv)
      For Each vRow As CDBDataRow In pDataTable.Rows
        Dim vPageNumber As Integer = IntegerValue(vRow.Item("WebPageNumber"))
        If vPageNumber > 0 Then
          Dim vWebNumber As Integer = vPageNumber \ WebControl.ItemsMultiplier
          If vWC.Existing = False OrElse vWC.WebNumber <> vWebNumber Then vWC.Init(vWebNumber)
          vRow.Item("WebURL") = vWC.PageURL(vPageNumber)
          vRow.Item("DescriptionLink") = String.Format("<a href=""{0}"">{1}</a>", vWC.PageURL(vPageNumber), vRow.Item("FundraisingDescription"))
        End If
      Next
    End Sub
    Private Sub GetContactFundraisingRequests(ByVal pDataTable As CDBDataTable)
      Dim vFields As New StringBuilder
      vFields.Append("fr.fundraising_request_number,fr.contact_number,request_date,request_description,fr.fundraising_request_stage,fundraising_request_stage_desc,fr.fundraising_status,fundraising_status_desc,fr.fundraising_request_type,fundraising_request_type_desc,fr.source,source_desc,target_amount,pledged_amount,pledged_date,")
      vFields.Append(mvEnv.Connection.DBIsNull("received_amount", "0"))
      vFields.Append(" AS received_amount,received_date,notes,fr.amended_by,fr.amended_on")
      Dim vNewFields As New StringBuilder
      With vNewFields
        .Append(",fr.request_end_date,fr.expected_amount,fr.gik_expected_amount,fr.gik_pledged_amount,fr.gik_pledged_date,")
        .Append(mvEnv.Connection.DBIsNull("fr.total_gik_received_amount", "0"))
        .Append(" AS total_gik_received_amount, fr.latest_gik_received_date, fr.number_of_payments, fr.logname, fr.created_by, fr.created_on, fr.target_date, ")
        .Append(mvEnv.Connection.DBIsNull("fr.expected_amount", "0"))
        .Append(" + ")
        .Append(mvEnv.Connection.DBIsNull("fr.gik_expected_amount", "0"))
        .Append(" AS total_expected_amount,")
        .Append(mvEnv.Connection.DBIsNull("fr.pledged_amount", "0"))
        .Append(" + ")
        .Append(mvEnv.Connection.DBIsNull("fr.gik_pledged_amount", "0"))
        .Append(" AS total_pledged_amount,")
        .Append(mvEnv.Connection.DBIsNull("fr.received_amount", "0"))
        .Append(" + ")
        .Append(mvEnv.Connection.DBIsNull("fr.total_gik_received_amount", "0"))
        .Append(" AS total_recived_amount,")
        .Append(mvEnv.Connection.DBIsNull("fr.pledged_amount", "0"))
        .Append(" - ")
        .Append(mvEnv.Connection.DBIsNull("fr.received_amount", "0"))
        .Append(" AS out_income_pledged_amount, ")
        .Append(mvEnv.Connection.DBIsNull("fr.gik_pledged_amount", "0"))
        .Append(" - ")
        .Append(mvEnv.Connection.DBIsNull("fr.total_gik_received_amount", "0"))
        .Append(" AS outstanding_gik_pledged_amount,(")
        .Append(mvEnv.Connection.DBIsNull("fr.pledged_amount", "0"))
        .Append(" - ")
        .Append(mvEnv.Connection.DBIsNull("fr.received_amount", "0"))
        .Append(") + (")
        .Append(mvEnv.Connection.DBIsNull("fr.gik_pledged_amount", "0"))
        .Append(" - ")
        .Append(mvEnv.Connection.DBIsNull("fr.total_gik_received_amount", "0"))
        .Append(") AS total_out_pledged_amount,fa.action_number,fps.total_gik_scheduled_amount")
        .Append(",fr.fundraising_business_type,fbt.fundraising_business_type_desc")
      End With
      vFields.Append(vNewFields.ToString)
      Dim vAdditionalItems As String = ""
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        vFields = vFields.Replace(vNewFields.ToString, "")
        vAdditionalItems = ",,,,,,,,,,,,,,,,,,,"
      End If
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbFundraisingBusinessType) Then
        vFields = vFields.Replace(",fr.fundraising_business_type,fbt.fundraising_business_type_desc", "")
        vAdditionalItems &= ",,"
      End If
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("contact_number", mvContact.ContactNumber)
      AddWhereFieldFromIntegerParameter(vWhereFields, "FundraisingRequestNumber", "fr.fundraising_request_number")

      Dim vSubSQL As String = ""
      Dim vSubSQL2 As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        vWhereFields.Add("fa.scheduled_payment_number", 0, CDBField.FieldWhereOperators.fwoNullOrEqual)
        vWhereFields.Add("a.completed_on", CDBField.FieldTypes.cftDate, "")
        Dim vSubAnsiJoins As New AnsiJoins
        vSubAnsiJoins.Add("fundraising_requests fr", "fa.fundraising_request_number", "fr.fundraising_request_number")
        vSubAnsiJoins.Add("actions a", "fa.action_number", "a.action_number")
        Dim vSubSQLStatement As SQLStatement = New SQLStatement(mvEnv.Connection, "fa.fundraising_request_number,MAX(fa.action_number) AS action_number", "fundraising_actions fa", vWhereFields, "", vSubAnsiJoins)
        vSubSQLStatement.GroupBy = "fa.fundraising_request_number"
        vSubSQL = vSubSQLStatement.SQL
        vWhereFields.Remove("fa.scheduled_payment_number")
        vWhereFields.Remove("a.completed_on")

        'Total GIK Scheduled Amount
        vSubAnsiJoins = New AnsiJoins
        Dim vWFields As New CDBFields(New CDBField("fundraising_payment_type", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType), CDBField.FieldWhereOperators.fwoNotEqual))
        vSubSQLStatement = New SQLStatement(mvEnv.Connection, "fundraising_request_number, SUM(payment_amount) AS total_gik_scheduled_amount", "fundraising_payment_schedule", vWFields)
        vSubSQLStatement.GroupBy = "fundraising_request_number"
        vSubSQL2 = vSubSQLStatement.SQL
      End If

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("fundraising_request_stages frs", "fr.fundraising_request_stage", "frs.fundraising_request_stage")
      vAnsiJoins.Add("fundraising_request_types frt", "fr.fundraising_request_type", "frt.fundraising_request_type")
      vAnsiJoins.AddLeftOuterJoin("fundraising_statuses frst", "fr.fundraising_status", "frst.fundraising_status")
      vAnsiJoins.AddLeftOuterJoin("sources s", "fr.source", "s.source")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        vAnsiJoins.AddLeftOuterJoin("(" & vSubSQL & ") fa", "fr.fundraising_request_number", "fa.fundraising_request_number")
        vAnsiJoins.AddLeftOuterJoin("(" & vSubSQL2 & ") fps", "fr.fundraising_request_number", "fps.fundraising_request_number")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbFundraisingBusinessType) Then
        vAnsiJoins.AddLeftOuterJoin("fundraising_business_types fbt", "fr.fundraising_business_type", "fbt.fundraising_business_type")
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields.ToString), "fundraising_requests fr", vWhereFields, "request_date", vAnsiJoins)
      vSQLStatement.SetDecimalPlaces = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "", vAdditionalItems)

      pDataTable.Columns("TotalGikReceivedAmount").FieldType = CDBField.FieldTypes.cftNumeric

      pDataTable.Columns("HasAction").FieldType = CDBField.FieldTypes.cftCharacter

      Dim vDefaultStatus As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundStatus)
      Dim vOrderByStatus As Boolean = vDefaultStatus.Length > 0 AndAlso mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlLockFundRequest)
      If vOrderByStatus Then pDataTable.AddColumn("SortColumn", CDBField.FieldTypes.cftInteger)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("HasAction").Length > 0 Then
          vRow.Item("HasAction") = ProjectText.String15904
        End If
        If vOrderByStatus Then
          If vRow.Item("FundraisingStatus") = vDefaultStatus Then
            vRow.Item("SortColumn") = "1"
          Else
            vRow.Item("SortColumn") = "2"  'Show locked at the end
          End If
        End If
      Next
      If vOrderByStatus Then
        Dim vTableSort(1) As CDBDataTable.SortSpecification
        vTableSort(0).Column = "SortColumn"
        vTableSort(1).Column = "RequestDate"
        vTableSort(1).Descending = True
        pDataTable.ReOrderRowsByMultipleColumns(vTableSort)
      End If
    End Sub
    Private Sub GetContactGAYEPledgePayments(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "transaction_sign, gaye_pledge_number, transaction_date, gph.batch_number, gph.transaction_number, fh.amount"
      vAttrs = vAttrs & ", donor_amount, employer_amount, government_amount, admin_fee_amount, other_matched_amount, payment_number"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("financial_history fh", "gph.batch_number", "fh.batch_number", "gph.transaction_number", "fh.transaction_number")
      vAnsiJoins.Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")

      Dim vWhereFields As New CDBFields(New CDBField("gph.gaye_pledge_number", mvParameters("PledgeNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "gaye_pledge_payment_history gph", vWhereFields, "gph.batch_number DESC,gph.transaction_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("TransactionSign") = "D" Then
          vRow.ChangeSign("Amount")
        End If
      Next
    End Sub
    Private Sub GetContactGAYEPledges(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "gaye_pledge_number, donor_id, org.name, ag.name As agency_name, payroll_organisation_number, pledge_amount, net_total, gp.start_date, pfo_code, gp.source, "
      vAttrs &= "s.source_desc, donor_total, employer_total, government_total, admin_fees_total, gp.product, product_desc, gp.rate, rate_desc, gp.cancellation_reason, cr.cancellation_reason_desc, "
      vAttrs &= "gp.cancelled_on, gp.cancelled_by, cancellation_source, cs.source_desc, gp.distribution_code, payroll_number, gp.payment_frequency, payment_frequency_desc, gp.amended_on, "
      vAttrs &= "gp.amended_by, gp.organisation_number, gp.agency_number, gp.created_by, gp.created_on, other_matched_total, charity_donor_reference"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("organisations org", "gp.organisation_number", "org.organisation_number")
      vAnsiJoins.Add("gaye_agencies ga", "gp.agency_number", "ga.organisation_number")
      vAnsiJoins.Add("organisations ag", "ga.organisation_number", "ag.organisation_number")
      vAnsiJoins.Add("sources s", "gp.source", "s.source")
      vAnsiJoins.Add("products p", "gp.product", "p.product")
      vAnsiJoins.Add("rates r", "gp.product", "r.product", "gp.rate", "r.rate")
      vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "gp.cancellation_reason", "cr.cancellation_reason")
      vAnsiJoins.AddLeftOuterJoin("sources cs", "gp.cancellation_source", "cs.source")
      vAnsiJoins.AddLeftOuterJoin("payment_frequencies pf", "gp.payment_frequency", "pf.payment_frequency")

      Dim vWhereFields As New CDBFields(New CDBField("gp.contact_number", mvContact.ContactNumber))
      If mvParameters.Exists("GayePledgeNumber") Then vWhereFields.Add("gp.gaye_pledge_number", mvParameters("GayePledgeNumber").IntegerValue)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "gaye_pledges gp", vWhereFields, "gp.gaye_pledge_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactGAYEPostTaxPledgePayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactGAYEPostTaxPledgePayments")
      Dim vAttrs As String = "transaction_sign,transaction_date,pth.batch_number,pth.transaction_number,fh.amount"
      vAttrs = vAttrs & ",donor_amount,employer_amount,fh.reference"
      Dim vSQL As String = "Select " & vAttrs & " FROM post_tax_pg_payment_history pth, financial_history fh, transaction_types tt"
      vSQL = vSQL & " WHERE pth.pledge_number = " & mvParameters("PledgeNumber").Value & " And fh.batch_number = pth.batch_number"
      vSQL = vSQL & " And fh.transaction_number = pth.transaction_number And tt.transaction_type = fh.transaction_type"
      vSQL = vSQL & " ORDER BY pth.batch_number DESC, pth.transaction_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("TransactionSign") = "D" Then
          vRow.ChangeSign("PayAmount")
        End If
      Next
    End Sub
    Private Sub GetContactGAYEPostTaxPledges(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactGAYEPostTaxPledges")
      Dim vAttrs As String = "pledge_number,org.name,label_name,payroll_number,pt.start_date,pledge_amount,total_pledged,donor_total,employer_total,pt.source,source_desc,pt.product,product_desc,pt.rate,rate_desc,pt.distribution_code,,last_payment_date,payment_number,cancellation_reason,,cancelled_on,cancelled_by,cancellation_source,,pt.amended_on,pt.amended_by,pt.organisation_number,pt.contact_number,,,pt.created_by,pt.created_on"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayrollGivingCreatedByOn) Then vAttrs = Replace$(vAttrs, ",pt.created_by,pt.created_on", ",,")
      Dim vSQL As String = "Select " & RemoveBlankItems(vAttrs) & " FROM post_tax_pg_pledges pt, Contacts c, organisations org, sources s, products p, rates r WHERE"
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = vSQL & " pt.organisation_number = " & mvContact.ContactNumber
      Else
        vSQL = vSQL & " pt.contact_number = " & mvContact.ContactNumber
      End If
      If mvParameters.Exists("PledgeNumber") Then vSQL = vSQL & "And pledge_number =  " & mvParameters("PledgeNumber").LongValue
      vSQL = vSQL & " And pt.contact_number = c.contact_number"
      vSQL = vSQL & " And pt.organisation_number = org.organisation_number"
      vSQL = vSQL & " And s.source = pt.source "
      vSQL = vSQL & " And p.product = pt.product "
      vSQL = vSQL & " And r.product = pt.product And r.rate = pt.rate"
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = vSQL & " ORDER BY surname, initials"
      Else
        vSQL = vSQL & " ORDER BY pledge_number"
      End If
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetDescriptions(pDataTable, "DistributionCode")
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        GetContactNames(pDataTable, "EmployeeContactNumber", "Employee")
        HideColumn("Employer")
        HideColumn("EmployerPayrollOrganisationName")
      Else
        GetPayrollInfo(pDataTable)
        HideColumn("Employee")
      End If
    End Sub
    Private Sub GetContactGiftAidDeclarations(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactGiftAidDeclarations")
      Dim vContactNoList As String = ""
      If mvContact.ContactType = Contact.ContactTypes.ctcJoint AndAlso mvParameters.ParameterExists("FastDataEntry").Bool Then
        vContactNoList = mvContact.ContactNumber.ToString
        For Each vContactLink As ContactLink In mvContact.GetJointLinks(True)
          vContactNoList = vContactNoList & ", " & vContactLink.ContactNumber2
        Next
      End If
      Dim vAttrs As String = "gad.contact_number, declaration_number, declaration_date, 'A' declaration_type, gad.source, confirmed_on, method, start_date, end_date, source_desc"
      vAttrs = vAttrs & ", batch_number, transaction_number, order_number, cancellation_reason, cancelled_by, cancelled_on, cancellation_source"
      vAttrs = vAttrs & ", gad.amended_by, gad.amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidDecCreatedBy) Then
        vAttrs = vAttrs & ", created_by, created_on"
      Else
        vAttrs = vAttrs & ",, "
      End If
      vAttrs = vAttrs & ", gad.notes, c.label_name"
      vAttrs &= ", c.title, c.forenames, c.initials, c.surname, c.honorifics, c.salutation, c.preferred_forename, c.sex, c.prefix_honorifics, c.surname_prefix, c.informal_salutation, a.address, a.building_number, a.house_name, a.town, a.county, a.postcode, co.country, a.branch"
      Dim vSQL As String = "Select " & RemoveBlankItems(vAttrs) & " FROM gift_aid_declarations gad, sources s, Contacts c, addresses a, countries co WHERE "
      If mvParameters.Exists("DeclarationNumber") Then
        vSQL = vSQL & "gad.declaration_number =  " & mvParameters("DeclarationNumber").LongValue & " And "
      End If
      If vContactNoList.Length > 0 Then
        vSQL = vSQL & " gad.contact_number IN (" & vContactNoList & ")"
      Else
        vSQL = vSQL & " gad.contact_number = " & mvContact.ContactNumber
      End If
      vSQL = vSQL & " AND gad.source = s.source AND gad.contact_number = c.contact_number AND a.address_number = c.address_number AND co.country = a.country ORDER BY start_date DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",,,,,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        Dim vGADSummary As String = ""
        Select Case vRow.Item("DeclarationType")
          Case "D"
            vRow.Item("Donations") = "Y"
            vRow.Item("DeclarationType") = "Donations Only"
            vGADSummary = "(Donations Only)"
          Case "M"
            vRow.Item("Members") = "Y"
            vRow.Item("DeclarationType") = "Memberships Only"
            vGADSummary = "(Memberships Only)"
          Case "A"
            vRow.Item("Donations") = "Y"
            vRow.Item("Members") = "Y"
            vRow.Item("DeclarationType") = "All"
            vGADSummary = "(All Payments)"
        End Select
        'BR19026 BR19437 Removed hard coding of DeclarationMethod
        GetLookupData(pDataTable, "DeclarationMethod", "gift_aid_declarations", "method")

        If Len(vRow.Item("StartDate")) > 0 And Len(vRow.Item("EndDate")) > 0 Then
          vGADSummary = vRow.Item("StartDate") & " to " & vRow.Item("EndDate") & " " & vGADSummary
        ElseIf Len(vRow.Item("StartDate")) > 0 And Len(vRow.Item("EndDate")) = 0 Then
          vGADSummary = vRow.Item("StartDate") & " onwards " & vGADSummary
        End If

        If mvContact.ContactType = Contact.ContactTypes.ctcJoint Then vGADSummary = vRow.Item("LabelName") & " - " & vGADSummary
        vRow.Item("Summary") = vGADSummary
      Next

      'BR14239: A new row having DeclarationNumber=0 and Summary=Organisations will be added on client side in FDEGiftAidDisplay

      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetContactH2HCollections(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactH2HCollections")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "hco.contact_number,hc.collection_number,c.campaign_desc,a.appeal_desc,ac.collection_desc,hc.start_date,hc.end_date,hco.route,rt.route_type,rt.route_type_desc,cs.collector_status,cs.collector_status_desc,hco.operator_contact_number"
      With vWhereFields
        .Add("hco.contact_number", mvParameters("ContactNumber").LongValue)
        .AddJoin("ac.collection_number", "hco.collection_number")
        .AddJoin("ac.campaign", "a.campaign")
        .AddJoin("ac.appeal", "a.appeal")
        .AddJoin("a.campaign", "c.campaign")
        .AddJoin("hc.collection_number", "ac.collection_number")
        .AddJoin("hco.operator_contact_number", "co.contact_number")
        .AddJoin("hco.route_type", "rt.route_type")
        .AddJoin("hco.collector_status", "cs.collector_status")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & "," & vConAttrs & " FROM h2h_collectors hco, appeal_collections ac, appeals a, campaigns c, h2h_collections hc, contacts co,route_types rt, collector_statuses cs WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & "," & ContactNameItems())
    End Sub
    Private Sub GetContactHPLinks(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactHPLinks")
      Dim vSQL As String
      'mvResultColumns = "ParentLink,ContactType,RelationshipCode,RelationshipDesc,ContactName,ContactNumber,Phone"
      Dim vConLinkOrder As String = mvEnv.GetConfig("link_order_con")
      If vConLinkOrder = "" Then vConLinkOrder = "surname, initials"
      Dim vOrgLinkOrder As String = mvEnv.GetConfig("link_order_org")
      If vOrgLinkOrder = "" Then vOrgLinkOrder = "name"
      'Get the Branch Relationship
      Dim vTable As String
      Dim vAttr As String
      Dim vOrg As New Organisation(mvEnv)
      If mvEnv.GetConfigOption("me_show_branch_relationship", False) Then
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vTable = "organisations"
          vAttr = "organisation_number"
        Else
          vTable = "contacts"
          vAttr = "contact_number"
        End If
        vOrg.Init()
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT o.name,o.organisation_number,o.contact_number,o.sort_name,o.abbreviation,o.dialling_code,o.std_code,o.telephone FROM " & vTable & " c, addresses a, branches b, organisations o WHERE c." & vAttr & " = " & mvContact.ContactNumber & " AND c.address_number = a.address_number AND a.branch = b.branch AND b.organisation_number = o.organisation_number")
        If vRecordSet.Fetch() Then
          vOrg.InitFromRecordSet(mvEnv, vRecordSet, OrganisationRecordSetTypes.ortName Or OrganisationRecordSetTypes.ortPhone)
          Dim vRow As CDBDataRow = pDataTable.AddRow
          With vRow
            .Item("ParentLink") = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBranchParent)
            .Item("ContactType") = "O"
            .Item("RelationshipCode") = "AUTO"
            .Item("RelationshipDesc") = Replace$(mvEnv.GetBranchText(DataSelectionText.String14600), ":", "")    'Branch
            .Item("ContactName") = vOrg.Name
            .Item("ContactNumber") = vOrg.OrganisationNumber.ToString
            .Item("Phone") = vOrg.PhoneNumber
          End With
        End If
        vRecordSet.CloseRecordSet()
      End If
      Dim vSelAttrs As String = ",r.relationship,relationship_desc,parent_relationship"
      Dim vRestrict As String = " AND r.high_profile = 'Y' AND (historical IS NULL OR historical = 'N')"
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone) & vSelAttrs
      vOrg.Init()
      Dim vOrgAttrs As String = "name," & vOrg.GetRecordSetFieldsPhone & vSelAttrs
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = "SELECT /* SQLServerCSC */ organisation_number_2," & vOrgAttrs & ",ol.organisation_link_number FROM organisation_links ol, organisations o, relationships r WHERE ol.organisation_number_1 = " & mvContact.ContactNumber & " AND o.organisation_number = ol.organisation_number_2 AND ol.relationship = r.relationship" & vRestrict & " ORDER BY " & vOrgLinkOrder
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,ORGANISATION_TYPE_1,relationship,relationship_desc," & OrgNameItems() & ",organisation_number_2,ORGANISATION_TELEPHONE,organisation_link_number")
        vSQL = "SELECT /* SQLServerCSC */ organisation_number_2," & vConAttrs & ",ol.organisation_link_number FROM organisation_links ol, contacts c, relationships r WHERE ol.organisation_number_1 = " & mvContact.ContactNumber & " AND c.contact_number = ol.organisation_number_2 AND contact_type <> 'O' AND ol.relationship = r.relationship" & vRestrict & " ORDER BY " & vConLinkOrder
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,CONTACT_TYPE_1,relationship,relationship_desc," & ContactNameItems() & ",organisation_number_2,CONTACT_TELEPHONE,organisation_link_number")
      Else
        vSQL = "SELECT /* SQLServerCSC */ contact_number_2," & vConAttrs & ",cl.contact_link_number FROM contact_links cl, contacts c, relationships r WHERE contact_number_1 = " & mvContact.ContactNumber & " AND c.contact_number = cl.contact_number_2 AND contact_type <> 'O' AND cl.relationship = r.relationship" & vRestrict & " ORDER BY " & vConLinkOrder
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,CONTACT_TYPE_1,relationship,relationship_desc," & ContactNameItems() & ",contact_number_2,CONTACT_TELEPHONE,contact_link_number")
        vSQL = "SELECT /* SQLServerCSC */ contact_number_2," & vOrgAttrs & ",cl.contact_link_number FROM contact_links cl, organisations o, relationships r WHERE contact_number_1 = " & mvContact.ContactNumber & " AND o.organisation_number = cl.contact_number_2 AND cl.relationship = r.relationship" & vRestrict & " ORDER BY " & vOrgLinkOrder
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,ORGANISATION_TYPE_1,relationship,relationship_desc," & OrgNameItems() & ",contact_number_2,ORGANISATION_TELEPHONE,contact_link_number")
      End If
      If mvContact.ContactType <> Contact.ContactTypes.ctcOrganisation Then
        Dim vCParentCode As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlParentRelationship)
        If vCParentCode.Length > 0 Then
          'See if a link is already shown from the contact to the parent (committee)
          Dim vFound As Boolean
          For Each vRow As CDBDataRow In pDataTable.Rows
            If vRow.Item("RelationshipCode") = vCParentCode Then
              vFound = True
              Exit For
            End If
          Next
          If Not vFound Then        'Get any link thru a current position
            vSQL = "SELECT organisation_number_2," & vOrgAttrs & " FROM contact_positions cp, organisation_links ol, organisations o, relationships r WHERE cp.contact_number = " & mvContact.ContactNumber & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND cp.organisation_number = ol.organisation_number_1 AND ol.relationship = '" & vCParentCode & "' AND o.organisation_number = ol.organisation_number_2 AND ol.relationship = r.relationship" & vRestrict
            pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,ORGANISATION_TYPE_1,relationship,relationship_desc," & OrgNameItems() & ",organisation_number_2,ORGANISATION_TELEPHONE")
          End If
        End If
      End If
      Dim vIndex As Integer = 0
      While vIndex < pDataTable.Rows.Count
        Dim vRow As CDBDataRow = pDataTable.Rows(vIndex)
        If Len(vRow.Item("ParentLink")) > 0 Then
          Dim vParentLink As String = vRow.Item("ParentLink")
          Dim vParentNo As Integer = IntegerValue(vRow.Item("ContactNumber"))
          If vRow.Item("ContactType") = "C" Then
            vSQL = "SELECT contact_number_2," & vConAttrs & " FROM contact_links cl, contacts c, relationships r WHERE contact_number_1 = " & vParentNo & " AND c.contact_number = cl.contact_number_2 AND contact_type <> 'O' AND cl.relationship" & mvEnv.Connection.DBLikeOrEqual(vParentLink) & " AND cl.relationship = r.relationship" & vRestrict
            pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,CONTACT_TYPE_1,relationship,relationship_desc," & ContactNameItems() & ",contact_number_2,CONTACT_TELEPHONE")
          Else
            vSQL = "SELECT organisation_number_2," & vOrgAttrs & " FROM organisation_links ol, organisations o, relationships r WHERE ol.organisation_number_1 = " & vParentNo & " AND o.organisation_number = ol.organisation_number_2 AND ol.relationship" & mvEnv.Connection.DBLikeOrEqual(vParentLink) & " AND ol.relationship = r.relationship" & vRestrict
            pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "parent_relationship,ORGANISATION_TYPE_1,relationship,relationship_desc," & OrgNameItems() & ",organisation_number_2,ORGANISATION_TELEPHONE")
          End If
        End If
        vIndex = vIndex + 1
      End While
    End Sub
    Private Sub GetContactCommsInformation(ByRef pDataTable As CDBDataTable)
      pDataTable = mvContact.CommsDataTable
    End Sub
    Private Sub GetContactInformation(ByRef pDataTable As CDBDataTable)
      pDataTable = mvContact.DataTable
    End Sub
    Private Sub GetContactJournal(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "contact_journal_number,journal_type,journal_time,batch_number,transaction_number,operation,journal_by,select_1,select_2,select_3"
      Dim vSQLStatement As SQLStatement
      If mvContact.Existing = False Then
        vAttrs = vAttrs & ",c.contact_number,label_name"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbJournalSelectName) Then
          vAttrs = vAttrs & ",cj.select_name"
        Else
          vAttrs = vAttrs & ",null as select_name"
        End If
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("cj.journal_by", mvEnv.User.Logname)
        vWhereFields.AddJoin("c.contact_number", "cj.contact_number")
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "contact_journals cj, contacts c", vWhereFields, "journal_time DESC")
      ElseIf mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vAttrs = vAttrs & ",c.contact_number,label_name"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbJournalSelectName) Then
          vAttrs = vAttrs & ",cj.select_name"
        Else
          vAttrs = vAttrs & ",null as select_name"
        End If
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("oa.organisation_number", mvContact.ContactNumber)
        vWhereFields.AddJoin("cj.address_number", "oa.address_number")
        vWhereFields.AddJoin("c.contact_number", "cj.contact_number")
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "organisation_addresses oa, contact_journals cj, contacts c", vWhereFields, "journal_time DESC")
      Else
        Dim vItem As String
        vItem = vAttrs
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbJournalSelectName) Then
          vAttrs = vAttrs & ",cj.select_name"
        Else
          vAttrs = vAttrs & ",null as select_name"
        End If
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "contact_journals cj", New CDBField("contact_number", mvContact.ContactNumber), "journal_time DESC")
        vAttrs = vItem & ",,,select_name"
      End If
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)

      Dim vJO As JournalOperation = mvEnv.GetJournalOperation("MAIL", "insert")
      If Not vJO Is Nothing Then
        vAttrs = "mailing_number,journal_type,mailing_date,,,operation,journal_by,,,"
        Dim vDate As Date = Today.AddDays(vJO.ActiveDays * -1)
        Dim vWhereFields As New CDBFields()
        Dim vFields As String = "mh.mailing_number, mh.mailing_date,'MAIL' AS  journal_type, '" & vJO.Description & "'  AS  operation,mailing_by AS journal_by"
        If mvContact.Existing = False Then
          vWhereFields.Add("mailing_by", mvEnv.User.Logname)
          vWhereFields.Add("mailing_date", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "mailing_history mh", vWhereFields, "mh.mailing_date DESC")
          vAttrs &= ",,"
        ElseIf mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vAttrs &= ",contact_number,label_name"
          vFields &= ",c.contact_number,label_name"
          vWhereFields.Add("oa.organisation_number", mvContact.ContactNumber)
          vWhereFields.AddJoin("cm.address_number", "oa.address_number")
          vWhereFields.AddJoin("cm.mailing_number", "mh.mailing_number")
          vWhereFields.Add("mh.mailing_date", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vWhereFields.AddJoin("c.contact_number", "cm.contact_number")
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "organisation_addresses oa, contact_mailings cm, mailing_history mh, contacts c", vWhereFields, "mh.mailing_date DESC")
        Else
          vWhereFields.Add("cm.contact_number", mvContact.ContactNumber)
          vWhereFields.AddJoin("cm.mailing_number", "mh.mailing_number")
          vWhereFields.Add("mh.mailing_date", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "contact_mailings cm, mailing_history mh", vWhereFields, "mh.mailing_date DESC")
          vAttrs &= ",,"
        End If
        vAttrs &= ","
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
      End If

      vJO = mvEnv.GetJournalOperation("CMAD", "insert")
      If Not vJO Is Nothing Then
        vAttrs = "mailing_document_number,journal_type,created_on,batch_number,transaction_number,operation,journal_by,,,"
        Dim vDate As Date = Today.AddDays(vJO.ActiveDays * -1)
        Dim vWhereFields As New CDBFields()
        Dim vFields As String = "mailing_document_number,'CMAD' AS  journal_type, created_on, batch_number, transaction_number, '" & vJO.Description & "'  AS  operation,created_by AS journal_by"
        If mvContact.Existing = False Then
          vAttrs &= ",contact_number,label_name"
          vFields &= ",c.contact_number,label_name"
          vWhereFields.Add("created_by", mvEnv.User.Logname)
          vWhereFields.Add("cmd.created_on", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vWhereFields.AddJoin("c.contact_number", "cmd.contact_number")
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "contact_mailing_documents cmd, contacts c", vWhereFields, "cmd.created_on DESC")
        ElseIf mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vAttrs &= ",contact_number,label_name"
          vFields &= ",c.contact_number,label_name"
          vWhereFields.Add("oa.organisation_number", mvContact.ContactNumber)
          vWhereFields.AddJoin("cmd.address_number", "oa.address_number")
          vWhereFields.Add("cmd.created_on", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vWhereFields.AddJoin("c.contact_number", "cmd.contact_number")
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "organisation_addresses oa, contact_mailing_documents cmd, contacts c", vWhereFields, "cmd.created_on DESC")
        Else
          vWhereFields.Add("cmd.contact_number", mvContact.ContactNumber)
          vWhereFields.Add("cmd.created_on", vDate).WhereOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
          vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, "contact_mailing_documents cmd", vWhereFields, "cmd.created_on DESC")
          vAttrs &= ",,"
        End If
        vAttrs &= ","
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
      End If

      'Remove all the journal rows that are outside the active days for that journal type
      Dim vInvalidRows As New List(Of CDBDataRow)
      For Each vDataRow As CDBDataRow In pDataTable.Rows
        Dim vValue As String = vDataRow.Item("JournalType")
        If vValue <> "MAIL" And vValue <> "CMAD" Then
          'Debug.Print(vValue)
          vJO = mvEnv.GetJournalOperation(vValue, vDataRow.Item("JournalEntry"))
          If vJO Is Nothing Then
            vInvalidRows.Add(vDataRow)
          Else
            Dim vDate As Date = Today.AddDays(vJO.ActiveDays * -1)
            If CDate(vDataRow.Item("Time")) < vDate Then
              vInvalidRows.Add(vDataRow)
            Else
              vDataRow.Item("JournalEntry") = vJO.Description
              If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
                vDataRow.Item("JournalEntry") = vDataRow.Item("JournalEntry") & " (" & vDataRow.Item("OrganisationContact") & ")"
              End If
            End If
          End If
        Else
          If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            vDataRow.Item("JournalEntry") = vDataRow.Item("JournalEntry") & " (" & vDataRow.Item("OrganisationContact") & ")"
          End If
        End If
      Next
      For Each vRow As CDBDataRow In vInvalidRows
        pDataTable.Rows.Remove(vRow)
      Next

      pDataTable.Columns("Time").FieldType = CDBField.FieldTypes.cftTime
      pDataTable.ReOrderRowsByColumn("Time", True)
      'If the first row only has a date portion with no time part (mailing) then we need to change it
      'This is to ensure that if it gets put in a spreadsheet it will not format as a date and loose the time portion from other rows
      If pDataTable.Rows.Count > 0 Then
        Dim vRow As CDBDataRow = pDataTable.Rows(0)
        If Len(vRow.Item("Time")) = 10 Then
          vRow.Item("Time") = vRow.Item("Time") & " "
        End If
      End If
    End Sub


    Private Sub GetContactLegacy(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "l.legacy_number,legacy_id,l.legacy_status,legacy_status_desc,l.source,source_desc,l.source_date," &
      "will_date,last_codicil_date,gross_estate_value,net_estate_value,total_estimated_value,admin_expenses_value,tax_value," &
      "other_bequests_value,net_for_probate,liabilities_value,date_of_death,death_notification_source,death_notification_date," &
      "date_of_probate,review_date,l.legacy_review_reason,legacy_review_reason_desc,agency_notification_date,accounts_received," &
      "accounts_approved,age_at_death,lead_charity,in_dispute,l.legacy_dispute_reason,legacy_dispute_reason_desc," &
      "net_estate_value - other_bequests_value AS residue,l.amended_by,l.amended_on,total_received,total_estimated_value - total_received AS outstanding_value,c.label_name"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("l.contact_number", mvParameters("ContactNumber").LongValue)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("legacy_statuses ls", "l.legacy_status", "ls.legacy_status")
      vAnsiJoins.Add("sources s", "l.source", "s.source")
      vAnsiJoins.Add("legacy_review_reasons lrr", "l.legacy_review_reason", "lrr.legacy_review_reason", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("legacy_dispute_reasons ldr", "l.legacy_dispute_reason", "ldr.legacy_dispute_reason", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("contacts c", "l.death_notification_source", "c.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vSubAnsiJoins As New AnsiJoins()
      vSubAnsiJoins.Add("legacy_bequest_receipts lbr", "l.legacy_number", "lbr.legacy_number")
      Dim vSubSQL As New SQLStatement(mvEnv.Connection, "l.legacy_number, SUM(amount) AS total_received", "contact_legacies l", vWhereFields, "", vSubAnsiJoins)
      vSubSQL.GroupBy = "l.legacy_number"

      vAnsiJoins.Add("(" & vSubSQL.SQL & ") x", "l.legacy_number", "x.legacy_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_legacies l", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      If pDataTable.Rows.Count > 0 Then
        If pDataTable.Rows(0).Item("TotalReceivedValue").ToString = "" Then
          pDataTable.Rows(0).Item("TotalReceivedValue") = "0.00"
          pDataTable.Rows(0).Item("OutstandingValue") = pDataTable.Rows(0).Item("TotalEstimatedValue")
        End If
        If pDataTable.Rows(0).Item("OtherBequestsValue").ToString = "" Then
          pDataTable.Rows(0).Item("Residue") = pDataTable.Rows(0).Item("NetEstateValue")
        End If
      End If
    End Sub
    Private Sub GetContactLegacyActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cl.master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,,,,,,,,,,,duration_days,duration_hours,duration_minutes,a.document_class,action_text"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cl.contact_number", mvContact.ContactNumber)
      vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department", mvEnv.User.Department)
      vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("actions a", "cl.master_action", "a.master_action")
      vAnsiJoins.Add("users u", "a.created_by", "u.logname")
      vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
      vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
      vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_legacies cl", vWhereFields, "sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
      'For Each vRow As CDBDataRow In pDataTable.Rows
      '  'Order by Status (Overdue, Defined, Scheduled)
      '  If vRow.Item("SortColumn") = Action.GetActionStatusCode(astScheduled) Then vRow.Item("SortColumn") = ""
      'Next
      'pDataTable.ReOrderRowsByColumn("SortColumn", True)
      'GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
      GetActionersAndSubjects(pDataTable)
    End Sub
    Private Sub GetContactLegacyAssets(ByVal pDataTable As CDBDataTable)
      If Not mvParameters.Exists("Activity") Then mvParameters.Add("Activity")
      mvParameters("Activity").Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGAssetActivity)
      GetContactCategories(pDataTable, False, False, False)
    End Sub
    Private Sub GetContactLegacyBequests(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "lb.legacy_number,lb.bequest_number,bequest_description,lb.bequest_type,bequest_type_desc,lb.bequest_sub_type,bequest_sub_type_desc,lb.bequest_status,bequest_status_desc,expected_value,estimated_outstanding,estimate,expected_fraction_quantity,expected_fraction_division,lb.product,product_desc,lb.rate,rate_desc,lb.distribution_code,distribution_code_desc,notes,total_received,condition_met_date"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("legacy_bequests lb", "cl.legacy_number", "lb.legacy_number")
      vAnsiJoins.Add("products p", "lb.product", "p.product")
      vAnsiJoins.Add("rates r", "lb.product", "r.product", "lb.rate", "r.rate")
      vAnsiJoins.Add("legacy_bequest_types lbt", "lb.bequest_type", "lbt.bequest_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("legacy_bequest_sub_types lbst", "lb.bequest_sub_type", "lbst.bequest_sub_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("legacy_bequest_statuses lbs", "lb.bequest_status", "lbs.bequest_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("distribution_codes dc", "lb.distribution_code", "dc.distribution_code", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cl.contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.HasValue("BequestNumber") Then vWhereFields.Add("lb.bequest_number", mvParameters("BequestNumber").LongValue)
      Dim vSubAnsiJoins As New AnsiJoins()
      vSubAnsiJoins.Add("legacy_bequests lb", "cl.legacy_number", "lb.legacy_number")
      vSubAnsiJoins.Add("legacy_bequest_receipts lbr", "cl.legacy_number", "lbr.legacy_number", "lb.bequest_number", "lbr.bequest_number")
      Dim vSubSQL As New SQLStatement(mvEnv.Connection, "lb.bequest_number, SUM(amount) AS total_received", "contact_legacies cl", vWhereFields, "", vSubAnsiJoins)
      vSubSQL.GroupBy = "lb.bequest_number"

      vAnsiJoins.Add("(" & vSubSQL.SQL & ") x", "lb.bequest_number", "x.bequest_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_legacies cl", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "", "EXPECTED_FRACTION")
    End Sub
    Private Sub GetContactLegacyExpenses(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.HasValue("BequestNumber") Then vWhereFields.Add("bequest_number", mvParameters("BequestNumber").LongValue)
      If mvParameters.HasValue("DateReceived") Then vWhereFields.Add("date_received", CDBField.FieldTypes.cftDate, mvParameters("DateReceived").Value)
      If mvParameters.HasValue("Amount") Then vWhereFields.Add("value", CDBField.FieldTypes.cftNumeric, mvParameters("Amount").Value)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("legacy_expenses le", "cl.legacy_number", "le.legacy_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bequest_number,date_received,value,notes", "contact_legacies cl", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactLegacyLinks(ByVal pDataTable As CDBDataTable)
      Dim vList As New ArrayListEx(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGRelationshipList))
      mvParameters.Add("Relationships", vList.CSStringList)
      GetContactLinksToOrFrom(pDataTable, DataSelectionTypes.dstContactLinksTo)
    End Sub
    Private Sub GetContactLegacyTaxCertificates(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.HasValue("TaxCertificateNumber") Then vWhereFields.Add("tax_certificate_number", mvParameters("TaxCertificateNumber").Value)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("legacy_tax_certificates ltc", "cl.legacy_number", "ltc.legacy_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "tax_certificate_number,tax_year,tax_percent,reference,date_received,gross_amount,net_amount,(gross_amount - net_amount) AS tax_amount,tax_claimed,tax_received", "contact_legacies cl", vWhereFields, "tax_year DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetContactLinksToOrFrom(ByVal pDataTable As CDBDataTable, ByVal pType As DataSelectionTypes)
      If pType = DataSelectionTypes.dstContactLinksTo Then
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          GetLinks(pDataTable, LinkSourceTable.OrganisationLinks, LinkDestinationTable.Organisations, pType)   'Get organisation_links to organisations
          GetLinks(pDataTable, LinkSourceTable.OrganisationLinks, LinkDestinationTable.Contacts, pType)        'Get organisation_links to contacts
        Else
          GetLinks(pDataTable, LinkSourceTable.ContactLinks, LinkDestinationTable.Contacts, pType)             'Get contact_links to contacts
          GetLinks(pDataTable, LinkSourceTable.ContactLinks, LinkDestinationTable.Organisations, pType)        'Get contact_links to organisations
        End If
      Else
        If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          GetLinks(pDataTable, LinkSourceTable.OrganisationLinks, LinkDestinationTable.Organisations, pType)   'Get organisation_links from organisations
          GetLinks(pDataTable, LinkSourceTable.ContactLinks, LinkDestinationTable.Contacts, pType)             'Get contact_links from contacts
        Else
          GetLinks(pDataTable, LinkSourceTable.ContactLinks, LinkDestinationTable.Contacts, pType)             'Get contact_links from contacts
          GetLinks(pDataTable, LinkSourceTable.OrganisationLinks, LinkDestinationTable.Organisations, pType)   'Get organisation_links from organisations
        End If
      End If
      If pType = DataSelectionTypes.dstContactLinksFrom AndAlso mvParameters.Exists("Maintenance") = False Then GetDelegateLinks(pDataTable)
      Dim vRow As CDBDataRow
      For vIndex As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
        vRow = pDataTable.Rows(vIndex)
        vRow.SetYNValue("Historical")
        If vRow.Item("ContactGroup") = "" Then
          If vRow.Item("Type2") = "O" Then
            vRow.Item("ContactGroup") = OrganisationGroup.DefaultGroupCode
          Else
            vRow.Item("ContactGroup") = ContactGroup.DefaultGroupCode
          End If
        End If
        If mvParameters.Exists("ContactNumber2") Then
          If vRow.Item("ContactNumber") <> mvParameters("ContactNumber2").Value Then pDataTable.RemoveRow(vRow)
        End If
        If mvParameters.Exists("ValidFrom") AndAlso mvParameters("ValidFrom").Value.Length > 0 Then
          If vRow.Item("ValidFrom") <> mvParameters("ValidFrom").Value Then pDataTable.RemoveRow(vRow)
        End If
        If mvParameters.Exists("ValidTo") AndAlso mvParameters("ValidTo").Value.Length > 0 Then
          If vRow.Item("ValidTo") <> mvParameters("ValidTo").Value Then pDataTable.RemoveRow(vRow)
        End If
        If mvParameters.Exists("RestrictNonHistoricLinks") AndAlso String.Compare(mvParameters("RestrictNonHistoricLinks").Value, "Y", True) = 0 AndAlso String.Compare(vRow.Item("Historical"), "yes", True) = 0 Then
          pDataTable.RemoveRow(vRow)
        End If
      Next
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        For Each vRow In pDataTable.Rows
          If mvEnv.User.AccessLevelFromOwnershipGroup(vRow.Item("OwnershipGroup")) < CDBEnvironment.OwnershipAccessLevelTypes.oaltRead Then
            vRow.Item("Phone") = ""
          End If
        Next
      End If
    End Sub

    Private Sub GetLegacyBequestForecasts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("legacy_number", mvParameters("LegacyNumber").LongValue)
      vWhereFields.Add("bequest_number", mvParameters("BequestNumber").LongValue)
      If mvParameters.HasValue("StageMonthsDelay") Then vWhereFields.Add("stage_months_delay", mvParameters("StageMonthsDelay").LongValue)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bequest_number,stage_months_delay,stage_percentage", "legacy_bequest_forecasts", vWhereFields, "stage_months_delay")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetLegacyBequestReceipts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      If mvParameters.HasValue("ReceiptNumber") Then
        vWhereFields.Add("receipt_number", mvParameters("ReceiptNumber").LongValue)
      Else
        vWhereFields.Add("legacy_number", mvParameters("LegacyNumber").LongValue)
        vWhereFields.Add("bequest_number", mvParameters("BequestNumber").LongValue)
      End If
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add(New AnsiJoin("batch_transactions bt", "lbr.batch_number", "bt.batch_number", "lbr.transaction_number", "bt.transaction_number"))
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bt.contact_number,receipt_number,lbr.batch_number,lbr.transaction_number,line_number,date_received,lbr.amount,lbr.status,lbr.notes", "legacy_bequest_receipts lbr", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Enum LinkSourceTable As Integer
      ContactLinks
      OrganisationLinks
    End Enum

    Private Enum LinkDestinationTable As Integer
      Contacts
      Organisations
    End Enum

    Private Sub GetLinks(ByVal pDataTable As CDBDataTable, ByVal pSourceLinkTable As LinkSourceTable, ByVal pLinkDestinationTable As LinkDestinationTable, ByVal pType As DataSelectionTypes)

      Dim vAttrs As String = ",valid_from,valid_to,historical,notes,amended_by,amended_on"
      Dim vSelAttrs As String = ",cl.relationship,relationship_desc,valid_from,valid_to,historical,cl.notes,cl.amended_by,cl.amended_on "

      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsNamePhoneGroup & vSelAttrs
      Dim vOrg As New Organisation(mvEnv)
      Dim vOrgAttrs As String = "name," & vOrg.GetRecordSetFieldsPhoneGroup & vSelAttrs

      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vConAttrs = vConAttrs & ",ownership_group "
        vOrgAttrs = vOrgAttrs & ",ownership_group "
        vAttrs = vAttrs & ",ownership_group"
      Else
        vConAttrs = vConAttrs & ",department "
        vOrgAttrs = vOrgAttrs & ",department "
        vAttrs = vAttrs & ",department"
      End If

      Dim vAnsiJoins As New AnsiJoins()
      Dim vTableName As String
      Dim vLinkAttr As String
      Dim vType1Attr As String
      Dim vPKAttribute As String
      Dim vPKParameter As String
      Select Case pSourceLinkTable
        Case LinkSourceTable.ContactLinks
          vTableName = "contact_links cl"
          vLinkAttr = "contact_number"
          vType1Attr = "CONTACT_TYPE"
          vPKAttribute = "cl.contact_link_number"
          vPKParameter = "ContactLinkNumber"
        Case Else
          vTableName = "organisation_links cl"
          vLinkAttr = "organisation_number"
          vType1Attr = "ORGANISATION_TYPE"
          vPKAttribute = "cl.organisation_link_number"
          vPKParameter = "ContactLinkNumber" 'should be OrganisationLinkNumber - but it isn't because there is only 1 Add/Update/Delete web service for the 2 tables, hence only 1 parameter name.  Also, this code is called from 4 places, Contact Rel.To + From + Org Rel. To + From
      End Select
      Dim vSelectSuffix As String
      Dim vLinkSuffix As String
      If pType = DataSelectionTypes.dstContactLinksTo Then
        vSelectSuffix = "_1"
        vLinkSuffix = "_2"
      Else
        vSelectSuffix = "_2"
        vLinkSuffix = "_1"
      End If
      Dim vSelectAttrs As String
      Dim vOrderBy As String
      Dim vItems As New StringBuilder
      Dim vTelephone As String
      Dim vGroup As String
      Select Case pLinkDestinationTable
        Case LinkDestinationTable.Contacts
          vAnsiJoins.Add("contacts c", "c.contact_number", "cl." & vLinkAttr & vLinkSuffix)
          vSelectAttrs = vLinkAttr & vLinkSuffix & "," & vConAttrs
          vOrderBy = mvEnv.GetConfig("link_order_con")
          If vOrderBy.Length = 0 Then
            If mvEnv.GetConfigOption("option_contact_groups", False) Then vOrderBy = "contact_group,"
            vOrderBy &= "surname,initials"
          End If
          vTelephone = "CONTACT_TELEPHONE"
          vGroup = ",contact_group"
          vItems.Append("relationship,")
          vItems.Append(vType1Attr)
          vItems.Append("_1,CONTACT_TYPE_2,")
          If pType = DataSelectionTypes.dstContactLinksTo Then
            vItems.Append("relationship_desc,")
          End If
          vItems.Append(ContactNameItems())
        Case Else
          vAnsiJoins.Add("organisations o", "o.organisation_number", "cl." & vLinkAttr & vLinkSuffix)
          vSelectAttrs = vLinkAttr & vLinkSuffix & "," & vOrgAttrs
          vOrderBy = mvEnv.GetConfig("link_order_org")
          If vOrderBy.Length = 0 Then
            If mvEnv.GetConfigOption("option_organisation_groups", False) Then vOrderBy = "organisation_group,"
            vOrderBy &= "name"
          End If
          vTelephone = "ORGANISATION_TELEPHONE"
          vGroup = ",organisation_group"
          vItems.Append("relationship," & vType1Attr & "_1,ORGANISATION_TYPE_2,")
          If pType = DataSelectionTypes.dstContactLinksTo Then
            vItems.Append("relationship_desc,")
          End If
          vItems.Append(OrgNameItems)
      End Select
      vItems.Append(",")
      vItems.Append(vLinkAttr)
      vItems.Append(vLinkSuffix)
      If pType = DataSelectionTypes.dstContactLinksFrom Then
        vItems.Append(",relationship_desc")
      End If
      vItems.Append(",")
      vItems.Append(vTelephone)
      vItems.Append(vAttrs)
      vItems.Append(vGroup)
      If mvType <> DataSelectionTypes.dstContactLegacyLinks Then
        If mvType = DataSelectionTypes.dstContactLinksFrom Then
          vItems.Append(",,,")
        Else
          vItems.Append(",status,status_date,status_reason")
        End If
        vSelectAttrs &= ",status,status_date,status_reason"
      End If

      If mvType <> DataSelectionTypes.dstContactLegacyLinks Then 'for some reason the Legacy links functionality doesn't support statuses
        vItems.Append(",cl.relationship_status,relationship_status_desc,rgb_value").Append(vPKAttribute)
        vSelectAttrs &= ",cl.relationship_status,relationship_status_desc,rgb_value"
      End If
      vItems.Append(",").Append(vPKAttribute)
      vSelectAttrs &= "," & vPKAttribute

      vAnsiJoins.Add("relationships r", "cl.relationship", "r.relationship")

      vAnsiJoins.AddLeftOuterJoin("relationship_statuses rs", "cl.relationship", "rs.relationship", "cl.relationship_status", "rs.relationship_status")

      Dim vWhereFields As New CDBFields()
      If vPKParameter.HasValue AndAlso mvParameters.Exists(vPKParameter) Then
        vWhereFields.Add(vPKAttribute, CDBField.FieldTypes.cftInteger, mvParameters(vPKParameter).IntegerValue) 'Called when the user is trying to edit 1 row.
      Else
        vWhereFields.Add("cl." & vLinkAttr & vSelectSuffix, mvContact.ContactNumber)
        If mvParameters.Exists("Relationships") Then
          vWhereFields.Add("r.relationship", mvParameters("Relationships").Value, CDBField.FieldWhereOperators.fwoIn)
        ElseIf mvParameters.Exists("Relationship") Then
          vWhereFields.Add("r.relationship", mvParameters("Relationship").Value)
        End If
        If mvParameters.Exists("ContactGroup") Then
          If pLinkDestinationTable = LinkDestinationTable.Contacts Then
            vWhereFields.Add("c.contact_group", mvParameters("ContactGroup").Value)
          Else
            vWhereFields.Add("o.organisation_group", mvParameters("ContactGroup").Value)
          End If
        End If
      End If
      If (mvType = DataSelectionTypes.dstContactLinksTo OrElse mvType = DataSelectionTypes.dstContactLegacyLinks) AndAlso pLinkDestinationTable = LinkDestinationTable.Contacts Then
        vWhereFields.Add("c.contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vSelectAttrs, vTableName, vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vItems.ToString)
    End Sub

    Private Sub GetContactMailings(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cm.mailing_number,cm.contact_number,cm.address_number,mh.mailing_date,direction,mh.mailing,mailing_desc,mailing_template,,m.notes,mailing_history_notes,mailing_by,mh.mailing_filename,mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,,,,"
      If mvType <> DataSelectionTypes.dstEventDelegateMailing AndAlso mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("oa.organisation_number", mvContact.ContactNumber)
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("contact_mailings cm", "oa.address_number", "cm.address_number")
        vAnsiJoins.Add("contacts c", "cm.contact_number", "c.contact_number")
        vAnsiJoins.Add("mailing_history mh", "cm.mailing_number", "mh.mailing_number")
        vAnsiJoins.Add("mailings m", "mh.mailing", "m.mailing")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMailingHistoryTopic) Then
          vAnsiJoins.AddLeftOuterJoin("topics t", "mh.topic", "t.topic")
          vAnsiJoins.AddLeftOuterJoin("sub_topics st", "mh.topic", "st.topic", "mh.sub_topic", "st.sub_topic")
        End If
        Dim vSQLS As New SQLStatement(mvEnv.Connection,
                                      RemoveBlankItems(vAttrs.Replace("mailing_history_notes", "mh.notes AS mailing_history_notes")) & "," & Replace(mvContact.GetRecordSetFieldsName, "c.contact_number,", "") & CheetahMailAttrs(True),
                                      "organisation_addresses oa", vWhereFields, "mh.mailing_date DESC, c.surname, c.initials DESC ", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLS, vAttrs, "," & ContactNameItems() & "," & CheetahMailItems(False) & ",,,")

        'Inclusion of contact emailings for organisation addressees 
        vAnsiJoins.RemoveJoin("contact_mailings cm")
        Dim vAnsiJoinArray() As AnsiJoin = {New AnsiJoin("contact_emailings cm", "oa.address_number", "cm.address_number")}
        ReDim Preserve vAnsiJoinArray(vAnsiJoins.Count)
        vAnsiJoins.CopyTo(vAnsiJoinArray, 1)
        vAnsiJoins = New AnsiJoins(vAnsiJoinArray)
        vSQLS = New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs.Replace("mailing_history_notes", "mh.notes AS mailing_history_notes")) & "," & Replace(mvContact.GetRecordSetFieldsName, "c.contact_number,", "") & CheetahMailAttrs(True) & ",opened_datetime AS opened_on,processed_on,processed_status,cm.error_number,communication_number", "organisation_addresses oa", vWhereFields, "mh.mailing_date DESC", vAnsiJoins)
        Dim vSelAttrs As String = vAttrs.Replace(",,,,", ",,processed_on,processed_status,cm.error_number")
        pDataTable.FillFromSQL(mvEnv, vSQLS, vSelAttrs, "communication_number," & ContactNameItems() & "," & CheetahMailItems(False, True) & ",,,")

        vAttrs = "mailing_document_number, cmd.contact_number,cmd.address_number,created_on,fulfillment_number,cmd.mailing,mailing_desc,cmd.mailing_template,,m.notes,,created_by,,,,,,,cmd.fulfillment_number,"
        vAnsiJoins.Clear()
        vAnsiJoins.Add("contact_mailing_documents cmd", "oa.address_number", "cmd.address_number")
        vAnsiJoins.Add("contacts c", "cmd.contact_number", "c.contact_number")
        vAnsiJoins.Add("mailings m", "cmd.mailing", "m.mailing")
        vSQLS = New SQLStatement(mvEnv.Connection,
                                 RemoveBlankItems(vAttrs.Replace("cmd.fulfillment_number", "cmd.fulfillment_number AS FNumber")) & "," & Replace(mvContact.GetRecordSetFieldsName, "c.contact_number,", ""),
                                 "organisation_addresses oa", vWhereFields, "cmd.created_on DESC, c.surname, c.initials DESC ", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLS, vAttrs, ",,," & ContactNameItems() & ",,,," & CheetahMailItems(True) & ",,,")
      Else
        Dim vWhereFields As New CDBFields
        If mvType = DataSelectionTypes.dstEventDelegateMailing Then
          vWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
          vWhereFields.Add("em.event_number", mvParameters("EventNumber").IntegerValue)
        Else
          vWhereFields.Add("contact_number", mvContact.ContactNumber)
        End If
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("mailing_history mh", "cm.mailing_number", "mh.mailing_number")
        vAnsiJoins.Add("mailings m", "mh.mailing", "m.mailing")
        If mvType = DataSelectionTypes.dstEventDelegateMailing Then
          vAnsiJoins.Add("event_mailings em", "em.mailing", "mh.mailing")
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMailingHistoryTopic) Then
          vAnsiJoins.AddLeftOuterJoin("topics t", "mh.topic", "t.topic")
          vAnsiJoins.AddLeftOuterJoin("sub_topics st", "mh.topic", "st.topic", "mh.sub_topic", "st.sub_topic")
        Else
          vAttrs = vAttrs.Replace("mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,", ",,,,,")
        End If
        Dim vSQLS As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs.Replace("mailing_history_notes", "mh.notes AS mailing_history_notes")) & CheetahMailAttrs(True), "contact_mailings cm", vWhereFields, "mh.mailing_date DESC", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLS, vAttrs, ",," & ContactNameItems(True) & ",,,," & CheetahMailItems(False) & ",,,")

        vSQLS = New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs.Replace("mailing_history_notes", "mh.notes AS mailing_history_notes")) & CheetahMailAttrs(True) & ",opened_datetime AS opened_on,processed_on,processed_status,cm.error_number,communication_number", "contact_emailings cm", vWhereFields, "mh.mailing_date DESC", vAnsiJoins)
        Dim vSelAttrs As String = vAttrs.Replace(",,,,", ",,processed_on,processed_status,cm.error_number")
        pDataTable.FillFromSQL(mvEnv, vSQLS, vSelAttrs, "communication_number," & ContactNameItems(True) & "," & CheetahMailItems(False, True) & ",,,")

        If mvType <> DataSelectionTypes.dstEventDelegateMailing Then
          vAttrs = "cmd.mailing_document_number,cmd.contact_number,cmd.address_number,cmd.created_on,{DIRECTION CLAUSE},cmd.mailing,m.mailing_desc,cmd.mailing_template,,m.notes,,cmd.created_by,fulfillment_history.fulfilment_filename,,,,,,fulfillment_history.fulfillment_number,,"
          vAnsiJoins.Clear()
          vAnsiJoins.Add("mailings m", "cmd.mailing", "m.mailing")
          vAnsiJoins.AddLeftOuterJoin("fulfillment_history", "cmd.fulfillment_number", "fulfillment_history.fulfillment_number")
          Dim vDirectionClause As String = "CASE WHEN cmd.fulfillment_number IS NULL THEN NULL ELSE 'P' END Direction" 'Null is turned into 'Pending' below, and anything other than O or I becomes 'Printed'
          Dim vSQLFieldNames As String = RemoveBlankItems(vAttrs)
          vSQLFieldNames = vSQLFieldNames.Replace("{DIRECTION CLAUSE}", vDirectionClause) 'crappy string manipulation because the CDBDataTable needs to have column names but doesn't understand CASE statements - or anything other than a field name for that matter
          vAttrs = vAttrs.Replace("{DIRECTION CLAUSE}", "Direction")
          vSQLS = New SQLStatement(mvEnv.Connection, vSQLFieldNames, "contact_mailing_documents cmd", vWhereFields, "created_on DESC", vAnsiJoins)
          pDataTable.FillFromSQL(mvEnv, vSQLS, vAttrs, ",,," & ContactNameItems(True) & ",,,," & CheetahMailItems(True) & ",,,")
        End If
      End If
      'End If
      Dim vCol As CDBDataColumn = pDataTable.Columns("Date")
      vCol.FieldType = CDBField.FieldTypes.cftDate
      pDataTable.ReOrderRowsByColumn("Date", True)
      GetDescriptions(pDataTable, "MailingTemplate")
      GetAddressData(pDataTable, True)
      GetCommunicationsData(pDataTable)
      vCol = pDataTable.Columns("Type")
      vCol.FieldType = CDBField.FieldTypes.cftCharacter
      vCol = pDataTable.Columns("AddressLine")
      vCol.FieldType = CDBField.FieldTypes.cftCharacter
      vCol = pDataTable.Columns("Direction")
      vCol.FieldType = CDBField.FieldTypes.cftCharacter
      For Each vRow As CDBDataRow In pDataTable.Rows
        Dim vAttr As String = vRow.Item("Direction")
        If vAttr = "O" Then
          vAttr = DataSelectionText.String23321    'Outgoing
        ElseIf vAttr = "I" Then
          vAttr = DataSelectionText.String23322    'Incoming
        ElseIf Len(vAttr) = 0 Then
          vAttr = DataSelectionText.String23338    'Pending
        Else
          vAttr = DataSelectionText.String23339    'Printed
        End If
        vRow.Item("Type") = vAttr
      Next
    End Sub
    Private Sub GetContactMannedCollections(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactMannedCollections")
      Dim vWhereFields As New CDBFields
      'ContactNumber , CollectionNumber, CampaignDesc, AppealDesc, CollectionDesc, CollectionDate, StartTime, EndTime
      Dim vAttrs As String = "mco.contact_number,mc.collection_number,c.campaign_desc,a.appeal_desc,ac.collection_desc,mc.collection_date,mc.start_time,mc.end_time"
      With vWhereFields
        .Add("mco.contact_number", mvParameters("ContactNumber").LongValue)
        .AddJoin("ac.collection_number", "mco.collection_number")
        .AddJoin("ac.campaign", "a.campaign")
        .AddJoin("ac.appeal", "a.appeal")
        .AddJoin("a.campaign", "c.campaign")
        .AddJoin("mc.collection_number", "ac.collection_number")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM manned_collectors mco, appeal_collections ac, appeals a, campaigns c, manned_collections mc WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactMembershipDetails(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactMembershipDetails")
      Dim vAddress As New Address(mvEnv)
      Dim vWhereFields As New CDBFields
      vAddress.Init()
      Dim vPaymentPlan As New PaymentPlan
      vPaymentPlan.Init(mvEnv)
      Dim vAttrs As String = "membership_number,o.contact_number,member_number,membership_type_desc,m.cancelled_on,renewal_date,renewal_amount,payment_method_desc,payment_frequency_desc,label_name,balance,joined,m.order_number,m.amended_by,m.amended_on,mt.members_per_order,membership_card,m.address_number"
      vAttrs = vAttrs & ",age_override,membership_card_expires,reprint_mship_card,number_of_members,branch_member,m.branch,name,applied,accepted,m.source,m.voting_rights,m.cancelled_by,m.cancellation_reason,m.cancellation_source"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataRgbValueForMemberType) Then vAttrs = vAttrs & ",mt.rgb_value"
      vAttrs = vAttrs & ",order_date,o.payment_method,o.payment_frequency,o.created_by,o.created_on,giver_contact_number,provisional,next_payment_due,last_payment_date,arrears,last_payment,order_term,in_advance,order_term_units,frequency_amount,expiry_date,renewal_pending,number_of_reminders,renewal_change_reason,renewal_changed_by,renewal_changed_on,renewal_change_value,future_cancellation_reason,future_cancellation_date,future_cancellation_source,amount,direct_debit,credit_card,bankers_order,covenant,their_reference,eligible_for_gift_aid,frequency"

      Dim vSelectCols As String = vAttrs & ",approval_membership,mt.membership_type,mt.subsequent_membership_type,o.membership_type AS o_membership_type,payer_required,annual,membership_term,m.contact_number AS member_contact_number,mt.branch_membership," & Replace$(vAddress.GetRecordSetFieldsDetailCountrySortCode, ",a.branch", "")
      vAttrs &= ",approval_membership,membership_type,subsequent_membership_type,o_membership_type,payer_required,,annual,membership_term,member_contact_number,branch_membership"
      Dim vAdditionalAttrs As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipCardIssueNumber) Then
        vAttrs &= ",membership_card_issue_number"
        vSelectCols &= ",membership_card_issue_number"
      Else
        vAdditionalAttrs = ","
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipStatus) Then
        vAttrs &= ",ms.membership_status,ms.membership_status_desc,rgb_membership_status"
        vSelectCols &= ",ms.membership_status,ms.membership_status_desc,ms.rgb_value AS rgb_membership_status"
      Else
        vAdditionalAttrs &= ",,,"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLockBranch) Then
        vAttrs &= ",lock_branch"
        vSelectCols &= ",lock_branch"
      Else
        vAdditionalAttrs &= ","
      End If

      Dim vAnsiJoines As New AnsiJoins
      vAnsiJoines.Add("orders o", "m.order_number", "o.order_number")
      vAnsiJoines.Add("membership_types mt", "m.membership_type", "mt.membership_type")
      vAnsiJoines.Add("payment_methods pm", "o.payment_method", "pm.payment_method ")
      vAnsiJoines.Add("payment_frequencies pf", "o.payment_frequency", "pf.payment_frequency")
      vAnsiJoines.Add("contacts c", "o.contact_number", "c.contact_number ")
      vAnsiJoines.Add("addresses a", "m.address_number", "a.address_number ")
      vAnsiJoines.Add("countries co", "a.country ", "co.country")
      vAnsiJoines.Add("branches b", "m.branch", "b.branch")
      vAnsiJoines.Add("organisations og", "b.organisation_number", "og.organisation_number")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipStatus) Then
        vAnsiJoines.Add("membership_statuses ms", "m.membership_status", "ms.membership_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If
      vWhereFields.Add("m.contact_number", CDBField.FieldTypes.cftInteger, mvContact.ContactNumber)
      If mvParameters.Exists("MembershipNumber") Then
        vWhereFields.Add("m.membership_number", CDBField.FieldTypes.cftInteger, mvParameters("MembershipNumber").IntegerValue)
      End If
      Dim vOrderBy As String = mvEnv.Connection.DBOrderByNullsFirstDesc("m.cancelled_on") & ", joined DESC"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vSelectCols, "members m", vWhereFields, vOrderBy, vAnsiJoines)
      pDataTable.FillFromSQL(mvEnv, vSQL, Replace$(vAttrs, "m.order_number", "DISTINCT_PAYMENT_PLAN_NUMBER"), vAdditionalAttrs & ",,,,,,,ADDRESS_LINE,,,,,,,,,")


      vPaymentPlan = New PaymentPlan
      vPaymentPlan.Init(mvEnv)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If mvParameters.Exists("MembershipNumber") Then
          Dim vMember As New Member
          vMember.Init(mvEnv, IntegerValue(vRow.Item("MembershipNumber")))
          vRow.Item("SubsequentMembershipTypeChangeDate") = vMember.FutureTypeChangeDate(vRow.Item("RenewalDate"), vRow.IntegerItem("Term"), vMember.Contact.DateOfBirth, BooleanValue(vRow.Item("RenewalPending")), vPaymentPlan.GetTermUnits(vRow.Item("TermUnits")))
        End If
        If Len(vRow.Item("CancelledOn")) > 0 Then vRow.Item("RenewalAmount") = ""
        vRow.SetYNValue("Provisional")
        vRow.SetYNValue("RenewalPending")
        vRow.SetYNValue("BranchMember")
        vRow.SetYNValue("VotingRights")
        vRow.SetYNValue("EligibleForGiftAid", True)
        Select Case vRow.Item("TermUnits")
          Case "M"
            vRow.Item("TermDesc") = DataSelectionText.String22822  '"Months"
          Case "W"
            vRow.Item("TermDesc") = DataSelectionText.String22823  '"Weeks"
          Case Else
            vRow.Item("TermDesc") = DataSelectionText.String22821  '"Years"
        End Select
        If Val(vRow.Item("Term")) < 0 Then
          vRow.Item("Term") = Math.Abs(vRow.IntegerItem("Term")).ToString
          vRow.Item("TermDesc") = DataSelectionText.String22822  '"Months"
        End If
        Dim vNextPayment As Double = DoubleValue(vRow.Item("FrequencyAmount"))
        Dim vNextPayDue As String = vRow.Item("NextPaymentDue")
        vPaymentPlan.GetNextSchedulePaymentInfo(vNextPayment, vNextPayDue, IntegerValue(vRow.Item("PaymentPlanNumber")), DoubleValue(vRow.Item("Balance")), Left$(vRow.Item("DirectDebitStatus"), 1) = "Y" Or Left$(vRow.Item("CreditCardStatus"), 1) = "Y", IntegerValue(vRow.Item("PaymentFrequencyFrequency")), False)
        vRow.Item("NextPaymentAmount") = Format$(vNextPayment, "Fixed")
        vRow.Item("NextPaymentDue") = vNextPayDue
        If mvEnv.GetConfigOption("me_display_expiry_minus_one", False) And IsDate(vRow.Item("MembershipCardExpires")) Then
          vRow.Item("MembershipCardExpires") = Format$(DateAdd("d", -1, vRow.Item("MembershipCardExpires")), CAREDateFormat)
        End If
      Next
      GetDescriptions(pDataTable, "SourceCode")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetDescriptions(pDataTable, "FutureCancellationReason")
      GetDescriptions(pDataTable, "FutureCancellationSource")
      GetMemberFutureTypes(pDataTable)
      GetNewOrderData(pDataTable, "PaymentPlanNumber", "", "GiftFrom", "GiftTo", "GiftMessage")
    End Sub
    Private Sub GetContactMemberships(ByVal pDataTable As CDBDataTable)
      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      Dim vAttrs As String = "membership_number,o.contact_number,member_number,membership_type_desc,m.cancelled_on,renewal_date,renewal_amount,payment_method_desc,payment_frequency_desc,label_name,balance,joined,m.order_number,m.amended_by,m.amended_on,mt.members_per_order,membership_card,m.address_number"
      vAttrs = vAttrs & ",age_override,membership_card_expires,reprint_mship_card,number_of_members,branch_member,m.branch,name,applied,accepted,m.source,m.voting_rights,m.cancelled_by,m.cancellation_reason,m.cancellation_source"
      Dim vFields As String = vAttrs & "," & Replace$(vAddress.GetRecordSetFieldsDetailCountrySortCode, ",a.branch", "") & ", single_membership"
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("m.contact_number", mvContact.ContactNumber)
      If mvParameters.Exists("CancellationReason") Then vWhereFields.Add("m.cancellation_reason", mvParameters("CancellationReason").Value)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("orders o", "m.order_number", "o.order_number")
      vAnsiJoins.Add("membership_types mt", "m.membership_type", "mt.membership_type")
      vAnsiJoins.Add("payment_methods pm", "o.payment_method", "pm.payment_method")
      vAnsiJoins.Add("payment_frequencies pf", "o.payment_frequency", "pf.payment_frequency")
      vAnsiJoins.Add("contacts c", "o.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "m.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      vAnsiJoins.Add("branches b", "m.branch", "b.branch")
      vAnsiJoins.Add("organisations og", "b.organisation_number", "og.organisation_number")
      Dim vOrderBy As String = mvEnv.Connection.DBOrderByNullsFirstDesc("m.cancelled_on") & ", joined DESC"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "members m", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, ",,,,,,,ADDRESS_LINE,single_membership")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If Len(vRow.Item("CancelledOn")) > 0 Then vRow.Item("RenewalAmount") = ""
        vRow.SetYNValue("VotingRights")
        If mvEnv.GetConfigOption("me_display_expiry_minus_one", False) And IsDate(vRow.Item("MembershipCardExpires")) Then
          vRow.Item("MembershipCardExpires") = Format$(DateAdd("d", -1, vRow.Item("MembershipCardExpires")), CAREDateFormat)
        End If
      Next
      GetDescriptions(pDataTable, "SourceCode")
      GetMemberFutureTypes(pDataTable)
    End Sub
    Private Sub GetContactNotifications(ByVal pDataTable As CDBDataTable)
      GetContactNotifications(pDataTable, False)
    End Sub
    Private Sub GetContactNotifications(ByVal pDataTable As CDBDataTable, pIncludeEmail As Boolean)
      Dim vWhereFields As CDBFields = Nothing
      Dim vAnsiJoins As AnsiJoins = Nothing

      If mvParameters.Exists("NotifyDeadlines") Then
        mvParameters.Add("IgnoreStatus", "N")   'Restrict on statuses
        mvParameters.Add("DeadlinesOnly", "Y")  'Overdue Actions
        GetActions(pDataTable)
        For Each vRow As CDBDataRow In pDataTable.Rows
          If Len(vRow.Item("ItemCode")) = 0 Then
            vRow.Item("ItemCode") = "O"
            vRow.Item("ItemDesc") = DataSelectionText.String18160    'Action (Overdue)
          End If
        Next
      End If
      If mvParameters.Exists("NotifyActions") Then
        If mvParameters.Exists("IgnoreStatus") = False Then mvParameters.Add("IgnoreStatus")
        If mvParameters.Exists("DeadlinesOnly") = False Then mvParameters.Add("DeadlinesOnly")
        mvParameters("IgnoreStatus").Value = "Y"    'Do not restrict on statuses
        mvParameters("DeadlinesOnly").Value = "N"   'All linked Actions
        GetActions(pDataTable)
        For Each vRow As CDBDataRow In pDataTable.Rows
          If Len(vRow.Item("ItemCode")) = 0 Then
            vRow.Item("ItemCode") = "A"
            If vRow.Item("LinkType") = "M" Then
              vRow.Item("ItemDesc") = DataSelectionText.String18161    'Action (Manager)
            Else
              vRow.Item("ItemDesc") = DataSelectionText.String18162    'Action (Actioner)
            End If
          End If
        Next
      End If
      If mvParameters.Exists("NotifyMeetings") Then
        Dim vAttrs As String = "m.meeting_number,link_type,meeting_date,meeting_desc,meeting_type_desc"
        vWhereFields = New CDBFields(New CDBField("ml.contact_number", mvEnv.User.ContactNumber))
        vWhereFields.Add("ml.notified", "N")
        vWhereFields.Add("ml.link_type", "C")
        vAnsiJoins = New AnsiJoins({New AnsiJoin("meetings m", "ml.meeting_number", "m.meeting_number"),
                                    New AnsiJoin("meeting_types mt", "m.meeting_type", "mt.meeting_type")})
        Dim vMeetingSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "meeting_links ml", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vMeetingSQLStatement, vAttrs, ",,,,")
        For Each vRow As CDBDataRow In pDataTable.Rows
          If Len(vRow.Item("ItemCode")) = 0 Then
            vRow.Item("ItemCode") = "M"
            vRow.Item("ItemDesc") = DataSelectionText.String18163    'Meeting
          End If
        Next
      End If
      If mvParameters.Exists("NotifyDocuments") Then
        If mvParameters.ContainsKey("Notified") = False Then mvParameters.Add("Notified", "N")
        GetDocuments(pDataTable, pIncludeEmail)
        For Each vRow As CDBDataRow In pDataTable.Rows
          If Len(vRow.Item("ItemCode")) = 0 Then
            vRow.Item("ItemCode") = "D"
            Dim vValue As String = ""
            Select Case vRow.Item("LinkType")
              Case "A"
                vValue = DataSelectionText.String18164    'Document (Addressee)
              Case "S"
                vValue = DataSelectionText.String18165    'Document (Sender)
              Case "C"
                vValue = DataSelectionText.String18166    'Document (Copied)
              Case "D"
                vValue = DataSelectionText.String18167    'Document (Distributed)
              Case Else
                vValue = DataSelectionText.String18168    'Document
            End Select
            vRow.Item("ItemDesc") = vValue
          End If
          If Len(vRow.Item("Subject")) > 0 Then
            vRow.Item("ItemDescription") = vRow.Item("Subject")
          End If
        Next
      End If

      'Get entity alert items that have not yet been notified
      vWhereFields = New CDBFields
      vAnsiJoins = New AnsiJoins

      vWhereFields.Add("eai.alert_notified", "N")
      vWhereFields.Add("ea.created_by", mvEnv.User.Logname)
      vAnsiJoins.Add("entity_alerts ea", "eai.entity_alert_number", "ea.entity_alert_number")
      vAnsiJoins.Add("contacts c", "eai.entity_item_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "eai.entity_alert_item_number,eai.created_on,eai.created_by,eai.entity_alert_message,ea.entity_item_number", "entity_alert_items eai", vWhereFields, "eai.entity_alert_item_number", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
      Dim vCDBRow As CDBDataRow = Nothing
      Dim vContact As New Contact(mvEnv)
      vContact.Init()
      While vRS.Fetch
        If vContact.ContactNumber <> vRS.Fields.Item("entity_item_number").IntegerValue Then vContact.Init(vRS.Fields.Item("entity_item_number").IntegerValue)
        vCDBRow = pDataTable.AddRow()
        vCDBRow.Item("ItemNumber") = vRS.Fields("entity_alert_item_number").Value
        vCDBRow.Item("ItemDesc") = String.Format("Alert ({0})", vContact.Name)
        vCDBRow.Item("ItemDate") = vRS.Fields("created_on").Value
        vCDBRow.Item("ItemDescription") = vRS.Fields("entity_alert_message").Value
        vCDBRow.Item("ItemType") = vRS.Fields("created_by").Value
        vCDBRow.Item("ItemCode") = "I"
        vCDBRow.Item("Access") = String.Empty
      End While
      vRS.CloseRecordSet()

      If pDataTable.Rows.Count > 0 Then
        For vRowIndex As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
          vCDBRow = pDataTable.Rows(vRowIndex)
          If vCDBRow.Item("Access") IsNot Nothing AndAlso vCDBRow.Item("Access").Length > 0 AndAlso BooleanValue(vCDBRow.Item("Access")) = False Then
            pDataTable.Rows.Remove(vCDBRow)
          End If
        Next
      End If

    End Sub
    Private Sub GetContactOutstandingInvoices(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vValue As String
      If mvParameters.Exists("SalesLedgerAccount") Then
        vValue = mvParameters("SalesLedgerAccount").Value
      Else
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactNumber").LongValue)
        vValue = New SQLStatement(mvEnv.Connection, "sales_ledger_account", "credit_customers", vWhereFields).GetValue
      End If
      Dim vAttrs As String = ",,i.batch_number,i.transaction_number,invoice_number,invoice_date,payment_due,amount_paid,i.invoice_pay_status,invoice_dispute_code,record_type,fh.amount,i.deposit_amount,outstanding"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets) Then vAttrs = Replace(vAttrs, ",i.deposit_amount", "")
      Dim vAnsiJoines As New AnsiJoins
      vAnsiJoines.Add("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
      vAnsiJoines.Add("financial_history fh", "i.batch_number", "fh.batch_number", "i.transaction_number", "fh.transaction_number")
      With vWhereFields
        .Clear()
        .Add("i.sales_ledger_account", vValue)
        .Add("i.company", mvParameters("Company").Value)
        .Add("i.record_type", "I")
        .Add("ips.fully_paid", "N")
        .Add("ips.pending_dd_payment", "N")
      End With
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
        vWhereFields.Add("i.print_invoice", "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
        vWhereFields.Add("i.print_invoice#2", "Y", CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
      End If
      'BR13747 display invoices owned by user department
      If mvEnv.GetConfigOption("opt_batch_ownership") And mvEnv.GetConfig("opt_batch_per_user") = "DEPARTMENT" Then
        vAnsiJoines.Add("batches b", "i.batch_number", "b.batch_number")
        vAnsiJoines.Add("users u", "b.batch_created_by", "u.logname")
        vAnsiJoines.Add("departments d", "u.department", "d.department")
        vWhereFields.Add("d.department", mvEnv.User.Department)
      End If
      Dim vFields As String = vAttrs.Substring(2)
      vFields = vFields.Replace("outstanding", "fh.amount - amount_paid AS outstanding")
      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vFields, "invoices i", vWhereFields, "i.payment_due, i.invoice_number", vAnsiJoines), vAttrs)
      If mvParameters.Exists("InvoiceNumbersAdded") Then
        Dim vInvoice As New Invoice
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.IntegerItem("InvoiceNumber") = 0 Then
            vInvoice.Init(mvEnv, vRow.IntegerItem("BatchNumber"), vRow.IntegerItem("TransactionNumber"))
            If vInvoice.Existing Then
              mvEnv.Connection.StartTransaction()
              vInvoice.SetInvoiceNumber(True, True)
              vInvoice.Save()
              mvEnv.Connection.CommitTransaction()
              vRow.Item("InvoiceNumber") = vInvoice.InvoiceNumber
            End If
          End If
        Next
      End If
    End Sub
    Private Sub GetContactOwners(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactOwners")
      Dim vSQL As String
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vSQL = "SELECT d.department,department_desc,ou.amended_by,ou.amended_on FROM organisation_users ou, departments d WHERE organisation_number = " & mvContact.ContactNumber & " AND ou.department = d.department"
      Else
        vSQL = "SELECT d.department,department_desc,cu.amended_by,cu.amended_on FROM contact_users cu, departments d WHERE contact_number = " & mvContact.ContactNumber & " AND d.department = cu.department"
      End If
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactPaymentPlans(ByRef pDataTable As CDBDataTable)
      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      Dim vPaymentPlan As New PaymentPlan
      vPaymentPlan.Init(mvEnv)
      Dim vAttrs As String = "order_number,o.contact_number,o.address_number,order_type,order_date,o.payment_method,payment_method_desc,o.payment_frequency,payment_frequency_desc,renewal_date,renewal_amount,o.source,cancelled_on,created_by,created_on,giver_contact_number,provisional,"
      vAttrs = vAttrs & "frequency_amount,next_payment_due,balance,arrears,last_payment,last_payment_date,amount,in_advance,order_term,order_term_units,cancelled_by,cancellation_reason,cancellation_source,future_cancellation_reason,future_cancellation_date,future_cancellation_source,"
      vAttrs = vAttrs & "renewal_change_reason,renewal_changed_by,renewal_changed_on,renewal_change_value,direct_debit,credit_card,bankers_order,covenant,sales_contact_number,renewal_pending,expiry_date,their_reference,o.amended_by,o.amended_on,reason_for_despatch,eligible_for_gift_aid,frequency,number_of_reminders,"
      vAttrs = vAttrs & "gift_membership,one_year_gift,payment_schedule_amended_on,first_amount,claim_day"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanPackToMember) Then
        vAttrs = vAttrs & ",pack_to_member"
      Else
        vAttrs = vAttrs & ",pack_to_donor"
      End If
      vAttrs = vAttrs & ",membership_type"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vAttrs = vAttrs & ",one_off_payment"
      vAttrs = vAttrs & If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoans) = True, ",loan", ",").ToString
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanEligibleForGiftAid) Then vAttrs = Replace$(vAttrs, "eligible_for_gift_aid", "")

      Dim vPaymentsOnly As Boolean = mvParameters.HasValue("ContactPaymentPlansPayments") 'Used in GetContactPaymentPlansPayments

      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("PaymentPlanNumber") Then
        vWhereFields.Add("order_number", mvParameters("PaymentPlanNumber").IntegerValue)
      Else
        vWhereFields.Add("o.contact_number", mvContact.ContactNumber)
      End If
      Dim vAnsiJoins As New AnsiJoins
      Dim vPaymentsData As CDBDataTable = Nothing
      Dim vAdditionalItems As New StringBuilder("ADDRESS_LINE,,,,,,,,,,,,,,,,,,,,LabelName")

      vAnsiJoins.Add("payment_methods pm", "o.payment_method", "pm.payment_method")
      vAnsiJoins.Add("payment_frequencies pf", "o.payment_frequency", "pf.payment_frequency")
      vAnsiJoins.Add("addresses a", "o.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      If vPaymentsOnly Then
        vAttrs = vAttrs & ",label_name"
        vWhereFields.Add("o.balance", 0, CDBField.FieldWhereOperators.fwoGreaterThan)
        vWhereFields.Add("o.cancelled_on", "")
        vWhereFields.Add("o.direct_debit", "Y", CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("o.credit_card", "Y", CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("o.bankers_order", "Y", CDBField.FieldWhereOperators.fwoNotEqual)
        vAnsiJoins.Add("contacts c", "o.contact_number", "c.contact_number")
        vPaymentsData = New CDBDataTable
        vPaymentsData.AddColumnsFromList(mvResultColumns)
        vAdditionalItems.Append(",,")
      Else
        vAttrs = vAttrs & ","
        If mvParameters.Exists("CancellationReason") Then vWhereFields.Add("cancellation_reason", mvParameters("CancellationReason").Value)
        If mvParameters.Exists("DirectDebit") Then vWhereFields.Add("direct_debit", mvParameters("DirectDebit").Value)
        If mvParameters.Exists("StandingOrder") Then vWhereFields.Add("bankers_order", mvParameters("StandingOrder").Value)
        If mvParameters.Exists("CCCA") Then vWhereFields.Add("credit_card", mvParameters("CCCA").Value)
      End If
      Dim vOrderBy As String = "cancelled_on" & mvEnv.Connection.DBSortByNullsFirst & ", order_date DESC, o.order_number DESC"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs) & "," & vAddress.GetRecordSetFieldsDetailCountrySortCode, "orders o", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "PAYMENT_PLAN_TYPE," & RemoveBlankItems(vAttrs), vAdditionalItems.ToString)

      Dim vContinue As Boolean = True
      For Each vRow As CDBDataRow In pDataTable.Rows
        vPaymentPlan.Init(mvEnv, IntegerValue(vRow.Item("PaymentPlanNumber")))
        If vPaymentPlan.Existing Then
          If vPaymentsOnly Then vContinue = Not vPaymentPlan.ContainsUnprocessedPayments
          If vContinue Then
            vRow.SetYNValue("Provisional")
            vRow.SetYNValue("RenewalPending")
            vRow.SetYNValue("EligibleForGiftAid", True)
            vRow.SetYNValue("GiftMembership")
            vRow.SetYNValue("OneYearGift")
            vRow.SetYNValue(If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanPackToMember) = True, "PackToMember", "PackToDonor"))
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vRow.SetYNValue("OneOffPayment")
            Select Case vRow.Item("TermUnits")
              Case "M"
                vRow.Item("TermDesc") = DataSelectionText.String22822  '"Months"
              Case "W"
                vRow.Item("TermDesc") = DataSelectionText.String22823  '"Weeks"
              Case Else
                vRow.Item("TermDesc") = DataSelectionText.String22821  '"Years"
            End Select
            If Val(vRow.Item("Term")) < 0 Then
              vRow.Item("Term") = Math.Abs(vRow.IntegerItem("Term")).ToString
              vRow.Item("TermDesc") = DataSelectionText.String22822  '"Months"
            End If
            vWhereFields.Clear()
            vWhereFields.Add("order_number", IntegerValue(vRow.Item("PaymentPlanNumber")))
            vRow.Item("NumberOfPayments") = mvEnv.Connection.GetCount("order_payment_history", vWhereFields).ToString
            Dim vNextPayment As Double = DoubleValue(vRow.Item("FrequencyAmount"))
            Dim vNextPayDue As String = vRow.Item("NextPaymentDue")
            vPaymentPlan.GetNextSchedulePaymentInfo(vNextPayment, vNextPayDue, IntegerValue(vRow.Item("PaymentPlanNumber")), DoubleValue(vRow.Item("Balance")), Left$(vRow.Item("DirectDebitStatus"), 1) = "Y" Or Left$(vRow.Item("CreditCardStatus"), 1) = "Y", vRow.IntegerItem("PaymentFrequencyFrequency"), False)
            vRow.Item("NextPaymentAmount") = Format$(vNextPayment, "Fixed")
            vRow.Item("NextPaymentDue") = vNextPayDue
            Dim vIndex As Integer = 0
            For Each vOPS As OrderPaymentSchedule In vPaymentPlan.ScheduledPayments(False)
              If vOPS.ScheduledPaymentStatusCode = "U" Then vIndex = vIndex + 1
            Next
            vRow.Item("UnprocessedPayments") = vIndex.ToString
            vRow.Item("NonDonationProducts") = "N"
            For Each vPPD As PaymentPlanDetail In vPaymentPlan.Details
              If Not vPPD.Product.Donation Then
                vRow.Item("NonDonationProducts") = "Y"
                Exit For
              End If
            Next
            Dim vPPDFixedAmount As String = ""
            For Each vPPD As PaymentPlanDetail In vPaymentPlan.Details
              If vPPD.Amount.Length > 0 Then
                vPPDFixedAmount = FixTwoPlaces(DoubleValue(vPPDFixedAmount) + DoubleValue(vPPD.Amount)).ToString("F")
              End If
            Next
            If vPPDFixedAmount.Length > 0 Then vRow.Item("OriginalPPDFixedAmount") = vPPDFixedAmount
            If vPaymentPlan.DetermineMembershipPeriod = PaymentPlan.MembershipPeriodTypes.mptFirstPeriod Then
              vRow.Item("FirstYearMembership") = "Y"
            Else
              vRow.Item("FirstYearMembership") = "N"
            End If
            If vPaymentPlan.PlanType = CDBEnvironment.ppType.pptMember Then
              vRow.Item("MembershipRateCode") = DirectCast(vPaymentPlan.Details(vPaymentPlan.GetDetailKeyFromLineNo(1)), PaymentPlanDetail).RateCode
              If vPaymentsOnly Then
                vPaymentPlan.LoadMembers()
                vPaymentPlan.GetMember(mvContact.ContactNumber)
                vRow.Item("MemberNumber") = vPaymentPlan.Member.MemberNumber
              End If
            Else
              vRow.Item("MembershipRateCode") = ""
            End If

            If vPaymentsOnly Then vPaymentsData.Rows.Add(vRow)
          End If
        End If
      Next
      If vPaymentsOnly Then pDataTable = vPaymentsData
      GetDescriptions(pDataTable, "SourceCode")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetDescriptions(pDataTable, "FutureCancellationReason")
      GetDescriptions(pDataTable, "FutureCancellationSource")
      GetDescriptions(pDataTable, "RenewalDateChangeReason")
      GetDescriptions(pDataTable, "ReasonForDespatch")
      If vPaymentsOnly Then GetDescriptions(pDataTable, "PayPlanMembershipTypeCode")
      If vPaymentsOnly Then GetContactNames(pDataTable, "ContactNumber", "LabelName")
      GetSalesContactNames(pDataTable)
      GetContactNames(pDataTable, "GiverContactNumber", "GiverContactName", "GiverContactAddressLine")
      GetNewOrderData(pDataTable, "PaymentPlanNumber", "NewOrderPackToDonor")
    End Sub

    Private Sub GetContactPaymentPlansPayments(ByRef pDataTable As CDBDataTable)
      'Used in Portal for PayMultiplePaymentPlans
      mvParameters.Add("ContactPaymentPlansPayments", "Y")
      GetContactPaymentPlans(pDataTable)
    End Sub

    Private Sub GetContactPerformances(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactPerformances")
      Dim vSQL As String = "SELECT p.performance,performance_desc,sequence,notes"
      vSQL = vSQL & ",number_of_payments,value_of_payments,number_above,value_above,number_between,value_between,number_below,value_below,number_of_mailings,no_response,first_payment_date"
      vSQL = vSQL & ",first_payment,last_payment_date,last_payment,maximum_payment_date,maximum_payment,rolling_value,preceding_rolling_value,average_value,average_per_mailing,response_rate"
      vSQL = vSQL & ",upper_level,lower_level,rolling_boundary,number_of_payments / std_number_of_payments AS std_number_of_payments,value_above / std_value_above AS std_value_above_,value_between / std_value_between AS std_value_between_,value_below / std_value_below AS std_value_below_,number_of_mailings / std_number_of_mailings AS std_number_of_mailings_,no_response / std_no_response AS std_no_response_,first_payment / std_first_payment AS std_first_payment_"
      vSQL = vSQL & ",last_payment / std_last_payment AS std_last_payment_,maximum_payment / std_maximum_payment AS std_maximum_payment_,rolling_value / std_rolling_value AS std_rolling_value_,preceding_rolling_value / std_preceding_rolling_value AS std_preceding_rolling_value_ ,average_value / std_average_value AS std_average_value_,average_per_mailing / std_average_per_mailing AS std_average_per_mailing_,response_rate / std_response_rate AS std_response_rate_"
      vSQL = vSQL & " FROM performances p LEFT OUTER JOIN "
      vSQL = vSQL & "(SELECT performance,number_of_payments,value_of_payments,number_above,value_above,number_between,value_between,number_below,value_below,number_of_mailings,no_response,first_payment_date"
      vSQL = vSQL & ",first_payment,last_payment_date,last_payment,maximum_payment_date,maximum_payment,rolling_value,preceding_rolling_value,average_value,average_per_mailing,response_rate"
      vSQL = vSQL & " FROM contact_performances WHERE contact_number = " & mvContact.ContactNumber & ") cp "
      vSQL = vSQL & "ON p.performance = cp.performance ORDER BY sequence"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL))
    End Sub
    Private Sub GetContactPictureDocuments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactPictureDocuments")
      Dim vSQL As String = "SELECT communications_log_number,docfile_extension FROM communications_log cl, packages p WHERE contact_number = " & mvContact.ContactNumber & " AND document_type = '" & mvEnv.GetConfig("cd_contact_image_document_type") & "' AND cl.package = p.package"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactPositionCategories(ByVal pDataTable As CDBDataTable)
      Dim vDepartmental As Boolean = mvEnv.GetConfigOption("cd_display_owning_activities")
      Dim vAttrs As String = "cc.contact_position_activity_id,contact_position_number,cc.activity,cc.activity_value,quantity,activity_date,cc.source,valid_from,valid_to,cc.amended_by,cc.amended_on,cc.notes,activity_desc,activity_value_desc,source_desc"

      Dim vSQLStatement As SQLStatement = GetCategorySQL(vDepartmental, HighProfileActivitySelection.BothHighProfileAndNonHighProfile, vAttrs, "activity_desc, activity_value_desc", False, False)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",,,")

      pDataTable.Columns("Status").AttributeName = "status"
      pDataTable.Columns("NoteFlag").AttributeName = "note_flag"
      For Each vRow As CDBDataRow In pDataTable.Rows
        If Len(vRow.Item("Notes")) > 0 Then vRow.Item("NoteFlag") = "Y"
        vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
      Next
      pDataTable.ReOrderRowsByColumn("StatusOrder")
    End Sub
    Private Sub GetContactPositionLinks(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cpl.contact_position_number,cpl.relationship,relationship_desc," & mvContact.GetRecordSetFieldsName() & ",cpl.history_only,cpl.valid_from,cpl.valid_to,cpl.notes,to_contact_group," & mvEnv.Connection.DBIsNull("o.organisation_group", "c.contact_group") & " AS contact_group" 'ContactNameResults()
      Dim vFieldNames As String = "contact_position_number,contact_number,relationship,relationship_desc," & ContactNameItems() & ",history_only,valid_from,valid_to,notes,to_contact_group,contact_group"
      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("contact_position_links cpl", "cp.contact_position_number", "cpl.contact_position_number")
        .Add("contacts c", "cpl.linked_contact_number", "c.contact_number")
        .AddLeftOuterJoin("organisations o", "cpl.linked_contact_number", "o.organisation_number")
        .Add("relationships r", "cpl.relationship", "r.relationship")
      End With
      Dim vWhereFields As New CDBFields(New CDBField("cp.contact_position_number", mvParameters("ContactPositionNumber").IntegerValue))
      If mvParameters.Exists("Relationship") Then vWhereFields.Add("cpl.relationship", mvParameters("Relationship").Value)
      If mvParameters.Exists("LinkedContactNumber") Then vWhereFields.Add("cpl.linked_contact_number", mvParameters("LinkedContactNumber").IntegerValue)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_positions cp", vWhereFields, "cpl.valid_from,cpl.valid_to", vAnsiJoins, True)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFieldNames, ",,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("Notes").Length > 0 Then vRow.Item("NotesFlag") = "Y"
        vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        If vRow.Item("Type2").Length = 0 Then vRow.Item("Type2") = "C"
      Next
      pDataTable.ReOrderRowsByColumn("StatusOrder")
    End Sub
    Private Sub GetContactPositions(ByVal pDataTable As CDBDataTable)
      Dim vItems As New StringBuilder
      vItems.Append("cp.organisation_number,cp.address_number,position,started,finished,cp.mail,")
      vItems.Append(mvEnv.Connection.DBSpecialCol("cp", "current"))
      vItems.Append(",position_location,position_function,position_seniority,contact_position_number,cp.amended_by,cp.amended_on,single_position,o.organisation_group,c.contact_group,o.name,")
      vItems.Append(",address_line1, address_line2, address_line3,") 'BR18024 First line of address missing from Contact > Positions > Address in table
      vItems.Append(mvContact.GetRecordSetFieldsName)
      vItems.Append(",")
      Dim vAddress As New Address(mvEnv)
      vItems.Append(vAddress.GetRecordSetFieldsDetailCountrySortCode)
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then vItems = vItems.Replace("position_function,position_seniority,", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks) Then vItems = vItems.Replace("single_position", "")
      vItems.Append(",oa.valid_from,oa.valid_to,")
      vItems.Append(mvEnv.Connection.DBToString("o.contact_number - cp.contact_number"))
      vItems.Append(" as default_contact")
      vItems.Append(",ps.position_status,ps.position_status_desc")

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("organisations o", "cp.organisation_number", "o.organisation_number")
      vAnsiJoins.Add("contacts c", "cp.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "cp.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      vAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number", "a.address_number", "oa.address_number")
      vAnsiJoins.Add("position_statuses ps", "cp.position_status", "ps.position_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("organisation_groups og", "o.organisation_group", "og.organisation_group", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vWhereFields As New CDBFields
      Dim vAttrs As New StringBuilder
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vWhereFields.Add("cp.organisation_number", mvContact.ContactNumber)
        vAttrs.Append("contact_number,address_number,position," & ContactNameItems())
        vItems.Append(",c.dialling_code,c.std_code,c.telephone,c.ex_directory")
      Else
        vWhereFields.Add("cp.contact_number", mvContact.ContactNumber)
        vAttrs.Append("organisation_number,address_number,position," & OrgNameItems())
        vItems.Append(",o.dialling_code,o.std_code,o.telephone")
        If mvParameters.Exists("OrganisationGroup") Then vWhereFields.Add("o.organisation_group", mvParameters("OrganisationGroup").Value)
        vWhereFields.Add("og.view_in_contact_card", "Y", CDBField.FieldWhereOperators.fwoNOT)
      End If
      If mvParameters.HasValue("Current") Then vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), mvParameters("Current").Value)
      If mvParameters.Exists("AddressNumber") Then vWhereFields.Add("cp.address_number", mvParameters("AddressNumber").LongValue)
      If mvParameters.Exists("ContactPositionNumber") Then vWhereFields.Add("cp.contact_position_number", mvParameters("ContactPositionNumber").LongValue)
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)

      vAttrs.Append(",started,finished,mail,current,position_location,position_function,position_seniority,amended_by,amended_on,contact_position_number,single_position,organisation_group,contact_group,ADDRESS_LINE,,,valid_from,valid_to,dialling_code,std_code,telephone,")
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vAttrs.Append("ex_directory,CONTACT_TELEPHONE")
      Else
        vAttrs.Append(",ORGANISATION_TELEPHONE")
      End If
      vAttrs.Append(",default_contact,position_status,position_status_desc")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then vAttrs = vAttrs.Replace("position_function,position_seniority,", ",,")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks) Then vAttrs = vAttrs.Replace("single_position", "")

      Dim vOrderBy As New StringBuilder
      vOrderBy.Append(mvEnv.Connection.DBSpecialCol("cp", "current"))
      vOrderBy.Append(" DESC, finished DESC, surname, forenames")

      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vItems.ToString), "contact_positions cp", vWhereFields, vOrderBy.ToString, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs.ToString)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("Current") = "Y" AndAlso IntegerValue(vRow.Item("OrganisationDefault")) = 0 Then
          vRow.Item("OrganisationDefault") = ProjectText.String15904
        Else
          vRow.Item("OrganisationDefault") = ""
        End If
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("Current")
        vRow.SetYNValue("SinglePosition")
        If vRow.Item("ExDirectory") = "Y" And mvContact.Department <> mvEnv.User.Department Then vRow.Item("PhoneNumber") = DataSelectionText.String23335 'Ex-Directory
        vRow.SetYNValue("ExDirectory")
      Next
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then
        GetDescriptions(pDataTable, "PositionFunction")
        GetDescriptions(pDataTable, "PositionSeniority")
      End If
    End Sub
    Private Sub GetContactPreviousNames(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactPreviousNames")
      Dim vAttrs As String = "search_name,created_on,amended_by,amended_on"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM contact_search_names WHERE contact_number = " & mvContact.ContactNumber & " AND is_active = 'N'"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactProcessedTransactions(ByVal pDataTable As CDBDataTable)
      Dim vIncludeBC As String = mvEnv.GetConfig("fp_batch_categories_show")
      Dim vExcludeBC As String = mvEnv.GetConfig("fp_batch_categories_hide")

      Dim vAttrs As New StringBuilder()
      vAttrs.Append("fh.batch_number,fh.transaction_number,transaction_type_desc,fh.transaction_date,fh.amount,payment_method_desc,fh.reference,posted,")
      vAttrs.Append("status,transaction_sign,pm.payment_method,")
      vAttrs.Append(If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode), "fh.currency_amount,", "fh.amount AS currency_amount,"))
      vAttrs.Append("fh.bank_details_number,fh.notes,fh.transaction_type,x.stock_item, y.postage_packing,")
      vAttrs.Append(If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode), "b.currency_code,", ","))
      vAttrs.Append("fh.transaction_origin, ts.transaction_origin_desc,bt.eligible_for_gift_aid AS can_add_gift_aid, sl.sl_line_count,")

      Dim vCols As String = vAttrs.ToString()
      vAttrs.Append("ba.bank_account, ba.bank_account_desc,ba.rgb_value AS rgb_bank_account, ba.rgb_value AS rgb_amount, ba.rgb_value AS rgb_currency_amount,")
      vCols &= "ba.bank_account, ba.bank_account_desc, rgb_bank_account, rgb_amount, rgb_currency_amount,"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) = False Then vCols = vCols.Replace("fh.amount AS currency_amount,", "currency_amount,")

      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      vAttrs.Append(vAddress.GetRecordSetFieldsCountry)

      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields As New CDBFields(New CDBField("fh.contact_number", mvContact.ContactNumber))

      If mvParameters.Exists("BatchNumber") Then
        vWhereFields.Add("fh.batch_number", mvParameters("BatchNumber").LongValue)
        vWhereFields.Add("fh.transaction_number", mvParameters("TransactionNumber").LongValue)
      End If

      If mvParameters.Exists("BatchCategory") Then
        vWhereFields.Add("b.batch_category", mvParameters("BatchCategory").Value)
      End If

      If mvParameters.Exists("PaymentMethod") Then
        vWhereFields.Add("pm.payment_method", mvParameters("PaymentMethod").Value)
      End If

      If vIncludeBC.Length > 0 OrElse vExcludeBC.Length > 0 Then
        vAnsiJoins.Add("batches b", "b.batch_number", "fh.batch_number")
        vAnsiJoins.Add("bank_accounts ba", "b.bank_account", "ba.bank_account")
        If vIncludeBC.Length > 0 Then
          vWhereFields.Add("b.batch_category", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vIncludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoIn)
        End If
        If vExcludeBC.Length > 0 Then
          vWhereFields.Add("b.batch_category#2", CDBField.FieldTypes.cftCharacter, String.Empty, CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("b.batch_category#3", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vExcludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotIn Or CDBField.FieldWhereOperators.fwoCloseBracket)
        End If
      End If

      vAnsiJoins.Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")
      vAnsiJoins.Add("payment_methods pm", "fh.payment_method", "pm.payment_method")
      vAnsiJoins.Add("addresses a", "fh.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      vAnsiJoins.AddLeftOuterJoin("batch_transactions bt", "fh.batch_number", "bt.batch_number", "fh.transaction_number", "bt.transaction_number")
      If vIncludeBC.Length = 0 AndAlso vExcludeBC.Length = 0 Then
        vAnsiJoins.AddLeftOuterJoin("batches b", "fh.batch_number", "b.batch_number")
        vAnsiJoins.AddLeftOuterJoin("bank_accounts ba", "b.bank_account", "ba.bank_account")
      End If
      vAnsiJoins.AddLeftOuterJoin("transaction_origins ts", "fh.transaction_origin", "ts.transaction_origin")

      'Nested SQL for stock products
      Dim vNestedAttrs As String = "fhd1.batch_number, fhd1.transaction_number, COUNT(*) AS stock_item"
      Dim vNestedAnsiJoins As New AnsiJoins({New AnsiJoin("products p1", "fhd1.product", "p1.product")})
      Dim vNestedWhereFields As New CDBFields(New CDBField("p1.stock_item", "Y"))
      Dim vNestedSQLStatement As New SQLStatement(mvEnv.Connection, vNestedAttrs, "financial_history_details fhd1", vNestedWhereFields, String.Empty, vNestedAnsiJoins)
      vNestedSQLStatement.Distinct = True
      vNestedSQLStatement.GroupBy = "fhd1.batch_number, fhd1.transaction_number"

      vAnsiJoins.AddLeftOuterJoin("( " & vNestedSQLStatement.SQL & " ) x", "fh.batch_number", "x.batch_number", "fh.transaction_number", "x.transaction_number")

      'Nested SQL for postage & packing
      vNestedAttrs = "fhd2.batch_number, fhd2.transaction_number, COUNT(*) AS postage_packing"
      vNestedAnsiJoins = New AnsiJoins({New AnsiJoin("products p2", "fhd2.product", "p2.product")})
      vNestedWhereFields = New CDBFields(New CDBField("p2.postage_packing", "Y"))
      vNestedSQLStatement = New SQLStatement(mvEnv.Connection, vNestedAttrs, "financial_history_details fhd2", vNestedWhereFields, String.Empty, vNestedAnsiJoins)
      vNestedSQLStatement.Distinct = True
      vNestedSQLStatement.GroupBy = "fhd2.batch_number, fhd2.transaction_number"

      vAnsiJoins.AddLeftOuterJoin("( " & vNestedSQLStatement.SQL & " ) y", "fh.batch_number", "y.batch_number", "fh.transaction_number", "y.transaction_number")

      'Nested SQL for sales ledger items
      vNestedAttrs = "bta.batch_number, bta.transaction_number, COUNT(*) AS sl_line_count"
      vNestedAnsiJoins = Nothing
      vNestedWhereFields = New CDBFields(New CDBField("bta.line_type", CDBField.FieldTypes.cftCharacter, "'K','L','N','U'", CDBField.FieldWhereOperators.fwoIn))
      vNestedSQLStatement = New SQLStatement(mvEnv.Connection, vNestedAttrs, "batch_transaction_analysis bta", vNestedWhereFields)
      vNestedSQLStatement.GroupBy = "bta.batch_number, bta.transaction_number"

      vAnsiJoins.AddLeftOuterJoin("( " & vNestedSQLStatement.SQL & " ) sl", "fh.batch_number", "sl.batch_number", "fh.transaction_number", "sl.transaction_number")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs.ToString), "financial_history fh", vWhereFields, "fh.batch_number DESC, fh.transaction_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols, "ADDRESS_LINE")

      pDataTable.Columns("ContainsStock").FieldType = CDBField.FieldTypes.cftCharacter
      pDataTable.Columns("ContainsPostage").FieldType = CDBField.FieldTypes.cftCharacter
      pDataTable.Columns("ContainsSalesLedgerItems").FieldType = CDBField.FieldTypes.cftCharacter

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("TransactionSign") = "D" Then
          vRow.ChangeSign("Amount")
          vRow.ChangeSign("CurrencyAmount")
        End If
        If vRow.Item("ContainsStock").Length > 0 Then
          vRow.Item("ContainsStock") = "Y"
          vRow.SetYNValue(("ContainsStock"))
        End If
        If vRow.Item("ContainsPostage").Length > 0 Then
          vRow.Item("ContainsPostage") = "Y"
          vRow.SetYNValue(("ContainsPostage"))
        End If
        vRow.Item("ContainsSalesLedgerItems") = If(IntegerValue(vRow.Item("ContainsSalesLedgerItems")) > 0, "Y", "N")
        vRow.SetYNValue("ContainsSalesLedgerItems")
        CheckAmountRGBValue(vRow)
      Next
    End Sub
    Private Sub GetContactPurchaseInvoices(ByVal pDataTable As CDBDataTable)
      'NYI("GetContactPurchaseInvoices")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "payee_contact_number,purchase_invoice_number,purchase_order_number,amount,purchase_invoice_date,payee_reference,cheque_reference_number,sort_code,account_number,bacs_processed,,pi.currency_code,currency_code_desc,ca.iban_number,ca.bic_code"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS) Then vAttrs = vAttrs.Replace(",sort_code,account_number,bacs_processed", ",,,")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vAttrs = vAttrs.Replace("pi.currency_code,currency_code_desc", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then vAttrs = vAttrs.Replace("ca.iban_number,ca.bic_code", ",")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contacts c", "pi.payee_contact_number", "c.contact_number")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vAnsiJoins.Add("currency_codes cc", "pi.currency_code", "cc.currency_code")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS) Then vAnsiJoins.AddLeftOuterJoin("contact_accounts ca", "pi.bank_details_number", "ca.bank_details_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vConAttrs & "," & RemoveBlankItems(vAttrs), "purchase_invoices pi", New CDBFields(New CDBField("pi.contact_number", mvContact.ContactNumber)), "purchase_invoice_date DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "CONTACT_NAME," & vAttrs)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS) Then GetLookupData(pDataTable, "BacsProcessed", "purchase_invoices", "bacs_processed")
    End Sub
    Private Sub GetContactPurchaseOrders(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "payee_contact_number,po.purchase_order_number,purchase_order_type_desc,purchase_order_desc,amount,balance,start_date,number_of_payments,output_group,po.cancellation_reason,cancellation_source,"
      vAttrs &= "cancelled_by,cancelled_on,ad_hoc_payments,payment_schedule,regular_payments,requires_authorisation,po.po_authorisation_level,po_authorisation_level_desc,authorised_by,authorised_on,cancellation_reason_desc,"
      vAttrs &= "source_desc,logname,po.currency_code,currency_code_desc,pi.pi_number,pot.requires_po_payment_type"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments) Then vAttrs = vAttrs.Replace("ad_hoc_payments", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRegularPurchaseOrderPayments) Then vAttrs = vAttrs.Replace("regular_payments", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) Then
        vAttrs = vAttrs.Replace("requires_authorisation,po.po_authorisation_level,po_authorisation_level_desc,authorised_by,authorised_on", ",,,,")
        vAttrs = vAttrs.Replace(",logname", ",")
      End If
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vAttrs = vAttrs.Replace("po.currency_code,currency_code_desc", ",")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("purchase_order_types pot", "po.purchase_order_type", "pot.purchase_order_type")
      vAnsiJoins.Add("contacts c", "po.payee_contact_number", "c.contact_number")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then
        vAnsiJoins.Add("currency_codes cc", "po.currency_code", "cc.currency_code")
      End If
      vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "po.cancellation_reason", "cr.cancellation_reason")
      vAnsiJoins.AddLeftOuterJoin("sources s", "po.cancellation_source", "s.source")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("po.contact_number", mvContact.ContactNumber)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) = True Then
        'Get a list of authorisation levels and logname
        Dim vAuthorisationUsersSQL As New SQLStatement(mvEnv.Connection, "po_authorisation_level, logname", "po_authorisation_users", New CDBField("logname", mvEnv.User.Logname))
        vAnsiJoins.AddLeftOuterJoin("po_authorisation_levels pal", "po.po_authorisation_level", "pal.po_authorisation_level")
        vAnsiJoins.AddLeftOuterJoin("(" & vAuthorisationUsersSQL.SQL & ") pau", "po.po_authorisation_level", "pau.po_authorisation_level")
      End If
      Dim vSubSQL As New SQLStatement(mvEnv.Connection, "pi.purchase_order_number,MAX(pi.purchase_invoice_number) AS pi_number", "purchase_invoices pi")
      vSubSQL.GroupBy = "pi.purchase_order_number"
      vAnsiJoins.AddLeftOuterJoin("(" & vSubSQL.SQL & ") pi", "po.purchase_order_number", "pi.purchase_order_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vConAttrs & "," & RemoveBlankItems(vAttrs), "purchase_orders po", vWhereFields, "start_date DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "CONTACT_NAME," & vAttrs)
      pDataTable.Columns("HasInvoice").FieldType = CDBField.FieldTypes.cftCharacter
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("RequiresAuthorisation")
        If vRow.Item("CanAuthorise").Length > 0 Then vRow.Item("CanAuthorise") = ProjectText.String15904
        If vRow.Item("HasInvoice").Length > 0 Then vRow.Item("HasInvoice") = ProjectText.String15904
      Next
    End Sub
    Private Sub GetContactRegisteredUsers(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("ru.contact_number", mvParameters.Item("ContactNumber").LongValue)
      If mvParameters.Exists("UserName") Then vWhereFields.Add("ru.user_name", mvParameters.Item("UserName").Value)
      If vWhereFields.Count > 0 Then
        Dim vAttrs As String = "ru.user_name,'"
        vAttrs += ProjectText.EncryptedPasswordString
        vAttrs += "' AS password,ru.email_address,ru.contact_number,ru.log_on_count,ru.last_logged_on,ru.created_on,ru.registration_data,ru.security_question,ru.security_answer,ru.amended_by,ru.amended_on,ru.last_updated_on,ru.valid_from,ru.valid_to,ru.login_attempts,ru.locked_out,password_expiry_date"
        Dim vList As String = "ru.user_name,ru.password,ru.email_address,ru.contact_number,ru.log_on_count,ru.last_logged_on,ru.created_on,ru.registration_data,ru.security_question,ru.security_answer,ru.amended_by,ru.amended_on,ru.last_updated_on,ru.valid_from,ru.valid_to,ru.login_attempts,ru.locked_out,password_expiry_date"
        'Removed some 'If Not mvEnv.GetDataStructureInfo' code
        If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPasswordExpiry) Then vAttrs = Replace(vAttrs, ",password_expiry_date", ",")
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "registered_users ru", vWhereFields)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vList)
        If Not mvEnv.GetConfigOption("view_registered_users_password", False) Then
          For Each vRow As CDBDataRow In pDataTable.Rows
            vRow.Item("Password") = "*********"
          Next
        End If
      End If
    End Sub
    Private Sub GetContactRoles(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "{0}{1},cr.organisation_number,name,cr.role,role_desc,valid_from,valid_to,is_active,cr.amended_by,cr.amended_on,contact_role_number"

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("roles r", "cr.role", "r.role")})
      vAnsiJoins.Add("organisations o", "cr.organisation_number", "o.organisation_number")
      vAnsiJoins.Add("contacts c", "cr.contact_number", "c.contact_number")
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactRoleNumber") Then vWhereFields.Add("cr.contact_role_number", mvParameters("ContactRoleNumber").IntegerValue)

      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        If mvParameters.Exists("AddressNumber") Then
          vAnsiJoins.Add("contact_positions cp", "cr.organisation_number", "cp.organisation_number", "cr.contact_number", "cp.contact_number")
          vWhereFields.Add("cp.organisation_number", mvContact.ContactNumber)
          vWhereFields.Add("cp.address_number", mvParameters.Item("AddressNumber").IntegerValue)
        Else
          vWhereFields.Add("cr.organisation_number", mvContact.ContactNumber)
        End If
      Else
        vWhereFields.Add("cr.contact_number", mvContact.ContactNumber)
      End If

      If mvParameters.Exists("ContactPositionNumber") Then
        If vAnsiJoins.ContainsAnyJoinToTable("contact_positions") = False Then vAnsiJoins.Add("contact_positions cp", "cr.contact_number", "cp.contact_number", "cr.organisation_number", "cp.organisation_number")
        vWhereFields.Add("cp.contact_position_number", mvParameters("ContactPositionNumber").IntegerValue)
      End If

      If mvParameters.Exists("ContactNumber2") Then vWhereFields.Add("c.contact_number", mvParameters.Item("ContactNumber2").IntegerValue)
      If mvParameters.Exists("OrganisationNumber") Then vWhereFields.Add("o.organisation_number", mvParameters.Item("OrganisationNumber").IntegerValue)
      If mvParameters.Exists("RoleStatus") Then vWhereFields.Add("cr.is_active", CDBField.FieldTypes.cftCharacter, "Y")

      If mvParameters.Exists("RoleValidOnDate") AndAlso IsDate(mvParameters("RoleValidOnDate").Value) Then
        'Only select roles that are valid on the given date
        vWhereFields.Add("cr.valid_from", CDBField.FieldTypes.cftDate, String.Empty, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
        vWhereFields.Add("cr.valid_from#2", CDBField.FieldTypes.cftDate, mvParameters("RoleValidOnDate").Value, CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("cr.valid_to", CDBField.FieldTypes.cftDate, String.Empty, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("cr.valid_to#2", CDBField.FieldTypes.cftDate, mvParameters("RoleValidOnDate").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice Or CDBField.FieldWhereOperators.fwoOR)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, String.Format(vAttrs, mvContact.GetRecordSetFieldsName, ""), "contact_roles cr", vWhereFields, "cr.is_active DESC, r.role_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, String.Format(vAttrs, "contact_number,", ContactNameItems()))

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Current")
      Next
    End Sub
    Private Sub GetContactSalesLedgerItems(ByVal pDataTable As CDBDataTable)
      'Add invoices and credit notes to Data Table
      'vAttrs = Attributes used in SQLStatement
      'vAttrNames = Attributes used when buiding the DataTable
      Dim vAttrs As String = "invoice_date,record_type,i.invoice_number,reference,invoice_pay_status,payment_due,deposit_amount,i.batch_number,i.transaction_number,tt.transaction_sign,reprint_count"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets) Then vAttrs = vAttrs.Replace("deposit_amount", "")
      Dim vAttrNames As String = vAttrs & ",bta_amount,other_amount,amount_outstanding,i.invoice_number,fh.contact_number"
      vAttrs = RemoveBlankItems(vAttrs) & ",fh.amount AS bta_amount,fh.amount AS other_amount,(fh.amount - amount_paid) AS amount_outstanding"

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("financial_history fh", "i.batch_number", "fh.batch_number", "i.transaction_number", "fh.transaction_number")
        .Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")
      End With

      Dim vInvoiceRecordTypes As String = "'" & Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.Invoice) & "','" & Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.CreditNote) & "'"
      Dim vWhereFields As New CDBFields()
      With vWhereFields
        If mvParameters.Exists("Company") Then .Add("i.company", mvParameters("Company").Value)
        If mvParameters.Exists("SalesLedgerAccount") Then .Add("i.sales_ledger_account", CDBField.FieldTypes.cftCharacter, mvParameters.Item("SalesLedgerAccount").Value)
        .Add("i.contact_number", CDBField.FieldTypes.cftInteger, mvParameters.OptionalValue("ContactNumber", "0"))
        .Add("i.record_type", vInvoiceRecordTypes, CDBField.FieldWhereOperators.fwoIn)
      End With

      Select Case mvParameters.OptionalValue("AllocationType", "A")
        Case "A"
          'All.
          'Do nothing
        Case "F"
          'Fully allocated.
          vWhereFields.Add("invoice_pay_status", mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid))
        Case "U"
          'Unallocated.
          Dim vPayStatus As String = "'" & mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid) & "','" & mvEnv.GetInvoicePayStatus(CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue) & "'"
          vWhereFields.Add("invoice_pay_status", vPayStatus, CDBField.FieldWhereOperators.fwoIn)
      End Select
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs & ",i.address_number", "invoices i", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrNames, "address_number,")

      'Add Un-Allocated SL cash to Data Table
      '**Any changes to the SQL below may need to be applied to the SQL in GetContactSalesLedgerReceipts**

      'Sub-select to calculate Cash-invoice amount
      Dim vSubAttrs As String = "fh2.batch_number, fh2.transaction_number, transaction_type, reference, SUM(fhd2.amount) AS bta_amount, SUM(fhd2.amount) AS other_amount, (SUM(fhd2.amount) - amount_paid) AS amount_outstanding, fh2.contact_number"
      Dim vSubAnsiJoins As New AnsiJoins
      With vSubAnsiJoins
        .Add("invoice_details id2", "i2.invoice_number", "id2.invoice_number")
        .Add("financial_history_details fhd2", "id2.batch_number", "fhd2.batch_number", "id2.transaction_number", "fhd2.transaction_number", "id2.line_number", "fhd2.line_number")
        .Add("financial_history fh2", "fhd2.batch_number", "fh2.batch_number", "fhd2.transaction_number", "fh2.transaction_number")
      End With
      Dim vSubWhereFields As New CDBFields(New CDBField("record_type", Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash)))
      If mvParameters.Exists("SalesLedgerAccount") Then vSubWhereFields.Add("i2.sales_ledger_account", CDBField.FieldTypes.cftCharacter, mvParameters.Item("SalesLedgerAccount").Value)
      Dim vSubSQLStatement As New SQLStatement(mvEnv.Connection, vSubAttrs, "invoices i2", vSubWhereFields, "", vSubAnsiJoins)
      vSubSQLStatement.GroupBy = "fh2.batch_number, fh2.transaction_number, fh2.contact_number, transaction_type, reference, amount_paid"

      'Main SQL for payment ('C'-type) Invoices
      vAttrs = "invoice_date,record_type,i.invoice_number,reference,invoice_pay_status,payment_due,deposit_amount,i.batch_number,i.transaction_number,tt.transaction_sign,reprint_count"
      vAttrs &= ",bta_amount,other_amount,amount_outstanding,{0}fh.contact_number"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets) Then vAttrs = vAttrs.Replace("deposit_amount", "")
      vAttrNames = String.Format(vAttrs, "i.invoice_number,")
      vAttrs = String.Format(vAttrs, "")
      Dim vGroupBy As String = vAttrs
      vAttrs = RemoveBlankItems(vAttrs)

      With vAnsiJoins
        .Clear()
        .Add("invoice_details id", "i.invoice_number", "id.invoice_number")
        .Add("(" & vSubSQLStatement.SQL & ") fh", "id.batch_number", "fh.batch_number", "id.transaction_number", "fh.transaction_number")
        .Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")
        .Add("batches b", "fh.batch_number", "b.batch_number")
      End With

      With vWhereFields
        .Remove("i.record_type")
        .Add("i.record_type", Invoice.GetRecordTypeCode(Invoice.InvoiceRecordType.SalesLedgerCash))
        .Add("fh.amount_outstanding", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("fh.amount_outstanding#2", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoOR)
        .Add("b.batch_type", Batch.GetBatchTypeCode(Batch.BatchTypes.FinancialAdjustment), CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With

      vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs & ",i.address_number", "invoices i", vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = vGroupBy & ",i.address_number"
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrNames, "address_number,")
      GetAddressData(pDataTable)

      Dim vTableSort(1) As CDBDataTable.SortSpecification
      vTableSort(0).Column = "Date"
      vTableSort(0).Descending = True
      vTableSort(1).Column = "InvoiceNumber"
      vTableSort(1).Descending = False
      pDataTable.ReOrderRowsByMultipleColumns(vTableSort)
      Dim vInvoice As New Invoice
      vInvoice.Init(mvEnv)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vInvoice.Existing Then
          vInvoice = New Invoice()
          vInvoice.Init(mvEnv)
        End If
        If DoubleValue(vRow.Item("Outstanding")) < 0 AndAlso vRow.Item("TransactionType") = "C" Then
          'Un-allocated cash
          vInvoice.Init(mvEnv, IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")))
          If vInvoice.InvoiceAmount > 0 Then
            vRow.Item("Credit") = vInvoice.InvoiceAmount.ToString("F")
            vRow.Item("Outstanding") = FixTwoPlaces(vInvoice.InvoiceAmount - Math.Abs(DoubleValue(vRow.Item("Outstanding")))).ToString("F")
            If vInvoice.AdjustmentStatus <> Invoice.InvoiceAdjustmentStatus.Normal AndAlso mvEnv.GetInvoicePayStatusType(vInvoice.InvoicePayStatus) = CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid Then
              If vInvoice.AdjustmentStatus = Invoice.InvoiceAdjustmentStatus.Moved _
              OrElse (vInvoice.AdjustmentStatus = Invoice.InvoiceAdjustmentStatus.Reversed AndAlso DoubleValue(vRow.Item("Debit")) = 0) Then
                'Invoice has been moved / reversed and is fully paid
                vRow.Item("Outstanding") = "-0.01"    'Set to negative so that the row gets removed below
              End If
            End If
          End If
        End If
        Select Case mvEnv.GetInvoicePayStatusType(vRow.Item("Status"))
          Case CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue
            vRow.Item("Status") = ""
          Case CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid
            'don't need to do anything
          Case CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid
            vRow.Item("Status") = "A"
          Case CDBEnvironment.InvoicePayStatusTypes.ipsDDCollectionPending
            'not yet supported
        End Select
        Select Case vRow.Item("TransactionType")
          Case "C"
            vRow.Item("InvoiceNumber") = ""
            vRow.Item("Debit") = ""
            vRow.Item("TransactionType") = DataSelectionText.String40020    'Un-Allocated Payment
            If vRow.Item("TransactionSign") = "D" OrElse DoubleValue(vRow.Item("Credit")) < 0 Then
              If DoubleValue(vRow.Item("Credit")) < 0 Then
                'vRow.Item("Debit") = System.Math.Abs(DoubleValue(vRow.Item("Credit"))).ToString
                vRow.Item("Debit") = ""
                vRow.Item("Credit") = vInvoice.GetAdjustmentInvoiceAmounts(IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")), False).ToString("F")
              End If
            Else
              If DoubleValue(vRow.Item("Credit")) = 0 Then
                vInvoice.Init(mvEnv, IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")))
                Dim vIAmount As Double = vInvoice.InvoiceAmount
                If vIAmount = 0 Then vIAmount = vInvoice.GetAdjustmentInvoiceAmounts(IntegerValue(vRow.Item("BatchNumber")), IntegerValue(vRow.Item("TransactionNumber")), False)
                If vIAmount > 0 Then vRow.Item("Credit") = vIAmount.ToString("F")
              End If
            End If
          Case "N"
            vRow.Item("TransactionType") = DataSelectionText.String18681    'Credit Note
            vRow.Item("Debit") = ""
          Case "I"
            vRow.Item("TransactionType") = DataSelectionText.String18682    'Invoice
            vRow.Item("Credit") = ""
        End Select
      Next
      'Finally remove any rows where the outstanding amount is negative
      Dim vAdjRow As CDBDataRow
      For vIndex As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
        vAdjRow = pDataTable.Rows(vIndex)
        If DoubleValue(vAdjRow.Item("Outstanding")) < 0 Then pDataTable.RemoveRow(vAdjRow)
      Next
    End Sub
    Private Sub GetContactScores(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactScores")
      Dim vSQL As String = "SELECT s.score,s.score_desc,cs.points,s.sequence,s.notes FROM scores s LEFT OUTER JOIN "
      vSQL = vSQL & "(SELECT score, points FROM contact_scores WHERE contact_number = " & mvContact.ContactNumber & ") cs "
      vSQL = vSQL & "ON s.score = cs.score ORDER BY sequence"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL))
    End Sub
    Private Sub GetContactServiceBookings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactServiceBookings")
      Dim vAttrs As String = "service_booking_number,label_name,start_date,end_date,transaction_date,amount,vat_amount,vat_rate_desc,cancelled_by,cancelled_on,cancellation_source,"
      vAttrs = vAttrs & "sb.amended_by,sb.amended_on,status,booking_contact_number,booking_address_number,service_contact_number,related_contact_number,batch_number,transaction_number,line_number,order_number,sales_contact_number,cancellation_reason"
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs) & ",amount - vat_amount AS net_amount FROM service_bookings sb, vat_rates vr, contacts c WHERE booking_contact_number = " & mvContact.ContactNumber & " AND sb.booking_status IS NOT NULL AND sb.amount >= 0 AND sb.vat_rate = vr.vat_rate AND service_contact_number = contact_number ORDER BY start_date DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "net_amount,,,,")
      GetSalesContactNames(pDataTable)
      GetContactNames(pDataTable, "RelatedContactNumber", "RelatedContactName")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetContactSourceFromLastMailing(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactSourceFromLastMailing")
      Dim vWhereFields As New CDBFields
      Dim vMonths As Integer = IntegerValue(mvEnv.GetConfig("fp_source_mailing_months"))
      If vMonths > 0 Then
        With vWhereFields
          .Add("cm.contact_number", CDBField.FieldTypes.cftLong, mvContact.ContactNumber)
          .Add("mh.mailing_date", CDBField.FieldTypes.cftDate, DateAdd("m", (vMonths * -1), Date.Today).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          .Add("mh.mailing_date#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
        End With
        If mvParameters.HasValue("PartSource") Then
          Dim vValue As String = mvParameters("PartSource").Value
          Dim vSplitValue() As String = vValue.Split(","c)
          Dim vIndex As Integer
          For Each vValue In vSplitValue
            If Not ((InStr(vValue, "*") > 0) Or (InStr(vValue, "?") > 0)) Then vValue = vValue & "*"
            Dim vOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoLike
            Dim vAttrNames As String = "s.source"
            If vSplitValue.Length > 1 Then
              If vIndex = 0 Then
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoOpenBracket
              Else
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoOR
                If vIndex = UBound(vSplitValue) Then vOperator = vOperator Or CDBField.FieldWhereOperators.fwoCloseBracket
              End If
              vAttrNames = vAttrNames & "#" & vIndex
            End If
            vWhereFields.Add(vAttrNames, CDBField.FieldTypes.cftCharacter, vValue, vOperator)
            vIndex += 1
          Next
        End If
        Dim vSQL As String = mvEnv.Connection.GetSelectSQLCSC & "s.source, source_desc,mh.mailing_date,m.mailing"
        vSQL = vSQL & " FROM contact_mailings cm INNER JOIN mailing_history mh ON cm.mailing_number = mh.mailing_number "
        vSQL = vSQL & " INNER JOIN segments sg ON mh.mailing = sg.mailing INNER JOIN sources s ON sg.source = s.source "
        vSQL = vSQL & " LEFT OUTER JOIN (Select mailing, mailing_desc FROM mailings WHERE history_only ='N') m ON s.thank_you_letter = m.mailing "
        vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY mh.mailing_date DESC"
        'vSQL = mvEnv.Connection.GetSelectSQLCSC & "s.source, source_desc, mh.mailing_date FROM contact_mailings cm, mailing_history mh, segments sg, sources s WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY mh.mailing_date DESC"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
      End If
    End Sub
    Private Sub GetContactSponsorshipClaimedPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactSponsorshipClaimedPayments")
      Dim vSQL As String = "SELECT gscl.contact_number, gsc.claim_number, gscl.batch_number, gscl.transaction_number, gscl.line_number, product_desc, gscl.amount_claimed, gscl.net_amount, gsc.claim_generated_date, fh.reference"
      vSQL = vSQL & " FROM ga_sponsorship_tax_claim_lines gscl, ga_sponsorship_tax_claims gsc, financial_history_details fhd, products p, financial_history fh"
      vSQL = vSQL & " WHERE gscl.contact_number = " & mvContact.ContactNumber & " AND gsc.claim_number = gscl.claim_number"
      vSQL = vSQL & " AND fhd.batch_number = gscl.batch_number AND fhd.transaction_number = gscl.transaction_number"
      vSQL = vSQL & " AND fhd.line_number = gscl.line_number AND p.product = fhd.product AND fh.batch_number = fhd.batch_number AND fh.transaction_number = fhd.transaction_number"
      vSQL = vSQL & " ORDER BY gscl.claim_number, gscl.batch_number DESC, gscl.transaction_number DESC, gscl.line_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactSponsorshipUnClaimedPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactSponsorshipUnClaimedPayments")
      Dim vSQL As String = "SELECT gsu.contact_number, gsu.batch_number, gsu.transaction_number, gsu.line_number, fh.transaction_date, product_desc, gsu.net_amount, fh.reference"
      vSQL = vSQL & " FROM ga_sponsorship_lines_unclaimed gsu, financial_history_details fhd, products p, financial_history fh"
      vSQL = vSQL & " WHERE gsu.contact_number = " & mvContact.ContactNumber & " AND fhd.batch_number = gsu.batch_number"
      vSQL = vSQL & " AND fhd.transaction_number = gsu.transaction_number AND fhd.line_number = gsu.line_number"
      vSQL = vSQL & " AND p.product = fhd.product AND fh.batch_number = fhd.batch_number AND fh.transaction_number = fhd.transaction_number"
      vSQL = vSQL & " ORDER BY gsu.batch_number DESC, gsu.transaction_number DESC, gsu.line_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetContactStandingOrders(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactStandingOrders")
      Dim vAttrs As String = "bo.bankers_order_number,standing_order_type,start_date,ca.sort_code,ca.account_number,amount,reference,bo.bank_account,bank_account_desc,order_number,created_by,created_on,bo.amended_by,bo.amended_on,bo.source,source_desc,account_name,cancellation_reason,cancellation_source,cancelled_by,cancelled_on,bo.bank_details_number"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then vAttrs += ",ca.iban_number,ca.bic_code"
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs)
      vSQL = Replace$(vSQL, "reference", mvEnv.Connection.DBSpecialCol("", "reference"))
      vSQL = vSQL & " FROM bankers_orders bo, contact_accounts ca, bank_accounts ba, sources s WHERE "
      If mvParameters.Exists("StandingOrderNumber") Then
        vSQL = vSQL & "bo.bankers_order_number =  " & mvParameters("StandingOrderNumber").LongValue & " AND "
      End If
      vSQL = vSQL & " bo.contact_number = " & mvContact.ContactNumber & " AND ca.bank_details_number = bo.bank_details_number AND bo.bank_account = ba.bank_account AND bo.source = s.source"
      vSQL = vSQL & " ORDER BY cancelled_on " & mvEnv.Connection.DBSortByNullsFirst & ", start_date DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",,,,,,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      GetBankInfo(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("StandingOrderTypeCode") = vRow.Item("StandingOrderType")
        If vRow.Item("StandingOrderType") = "C" Then
          vRow.Item("StandingOrderType") = "CAF"
        Else
          vRow.Item("StandingOrderType") = ""
        End If
      Next
    End Sub
    Private Sub GetContactStatusHistory(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "contact_number,sh.status,status_reason,valid_to,sh.response_channel,response_channel_desc,sh.amended_by,sh.amended_on,status_desc,rgb_value"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataRgbValueForStatus) = False Then vAttrs = vAttrs.Replace("rgb_value", "")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) = False Then vAttrs = vAttrs.Replace("sh.response_channel,response_channel_desc", ",")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("statuses s", "sh.status", "s.status")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) Then
        vAnsiJoins.AddLeftOuterJoin("response_channels rc", "sh.response_channel", "rc.response_channel")
      End If
      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "status_history sh", New CDBFields(New CDBField("contact_number", mvContact.ContactNumber)), "valid_to DESC", vAnsiJoins), vAttrs)
    End Sub
    Private Sub GetContactStickyNotes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("unique_id", mvParameters("ContactNumber").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("NoteNumber") Then vWhereFields.Add("note_number", mvParameters("NoteNumber").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "note_number,notes,created_on," & mvEnv.Connection.DBSpecialCol("", "permanent") & ",amended_by,amended_on,record_type", "sticky_notes", vWhereFields, "note_number DESC")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetContactSubscriptions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactSubscriptions")
      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      Dim vAttrs As String = "subscription_number,s.address_number,order_number,product_desc,quantity,despatch_method_desc,reason_for_despatch_desc,valid_from,valid_to,cancellation_reason,cancelled_by,cancelled_on,cancellation_source,s.product,s.despatch_method,s.reason_for_despatch,communication_number,"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then vAttrs = Replace$(vAttrs, "communication_number", "")
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs) & "," & vAddress.GetRecordSetFieldsDetailCountrySortCode & " FROM subscriptions s, products p, despatch_methods dm, reasons_for_despatch rfd, addresses a, countries co WHERE s.contact_number = " & mvContact.ContactNumber & " AND s.product = p.product AND s.despatch_method = dm.despatch_method AND s.reason_for_despatch = rfd.reason_for_despatch AND s.address_number = a.address_number AND a.country = co.country ORDER BY cancelled_on " & mvEnv.Connection.DBSortByNullsFirst & ", order_number DESC, s.product"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",ADDRESS_LINE,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then GetCommunicationInfo(pDataTable, "CommunicationNumber", "DeliverTo")
    End Sub

    Private Sub GetContactSuppressions(ByVal pDataTable As CDBDataTable)
      Dim vTable As String
      Dim vAttr As String
      Dim vCurrentOnly As Boolean = mvParameters.ParameterExists("Current").Bool

      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vTable = "organisation_suppressions st"
        vAttr = "organisation_number"
      Else
        vTable = "contact_suppressions st"
        vAttr = "contact_number"
      End If
      Dim vAttrs As String = "ms.mailing_suppression,mailing_suppression_desc,valid_from,valid_to,st.notes,ms.notes as m_notes,st.source,source_desc,st.response_channel,response_channel_desc,st.amended_by,st.amended_on"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailingSuppressionsNotes) Then vAttrs = Replace(vAttrs, "ms.notes as m_notes", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbSuppressionSource) Then vAttrs = Replace(vAttrs, "st.source,source_desc", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) Then vAttrs = Replace(vAttrs, "st.response_channel,response_channel_desc", ",")

      Dim vWhereFields As New CDBFields
      Dim vAnisJoins As New AnsiJoins
      vAnisJoins.Add("mailing_suppressions ms", "st.mailing_suppression", "ms.mailing_suppression")
      If mvParameters.Exists("LookupGroup") Then
        vAnisJoins.Add("lookup_group_details lgd", "st.mailing_suppression", "lgd.lookup_item")
        vWhereFields.Add("lgd.lookup_group", mvParameters("LookupGroup").Value)
      ElseIf mvParameters.Exists("SuppressionGroup") Then
        vAnisJoins.Add("suppression_group_details sgd", "st.mailing_suppression", "sgd.mailing_suppression")
        vWhereFields.Add("sgd.suppression_group", mvParameters("SuppressionGroup").Value)
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbSuppressionSource) Then
        vAnisJoins.AddLeftOuterJoin("sources s", "st.source", "s.source")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) Then
        vAnisJoins.AddLeftOuterJoin("response_channels rc", "st.response_channel", "rc.response_channel")
      End If
      vWhereFields.Add(vAttr, mvContact.ContactNumber)
      If mvParameters.Exists("Suppression") Then vWhereFields.Add("ms.mailing_suppression", mvParameters("Suppression").Value)
      If mvParameters.Exists("ValidFrom") Then
        vWhereFields.Add("st.valid_from", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value)
      ElseIf vCurrentOnly Then
        vWhereFields.Add("st.valid_from", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If
      If mvParameters.Exists("ValidTo") Then
        vWhereFields.Add("st.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ValidTo").Value)
      ElseIf vCurrentOnly Then
        vWhereFields.Add("st.valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      If mvParameters.Exists("AmendedOn") Then vWhereFields.Add("st.amended_on", CDBField.FieldTypes.cftDate, mvParameters("AmendedOn").Value)
      Dim vOrderBy As String = "valid_from DESC"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vTable, vWhereFields, vOrderBy, vAnisJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs.Replace("ms.notes as m_notes", "m_notes"))
    End Sub
    Private Sub GetContactUnMannedCollections(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetContactUnMannedCollections")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      '"OrganisationNumber,       CollectionNumber,CampaignDesc,     AppealDesc,     CollectionDesc,    StartDate,   EndDate,ContactNumber," & ContactNameResults
      Dim vAttrs As String = "uc.organisation_number,uc.collection_number,c.campaign_desc,a.appeal_desc,ac.collection_desc,uc.start_date,uc.end_date,uc.contact_number"
      With vWhereFields
        .Add("uc.organisation_number", mvParameters("ContactNumber").LongValue)
        .AddJoin("ac.collection_number", "uc.collection_number")
        .AddJoin("ac.campaign", "a.campaign")
        .AddJoin("ac.appeal", "a.appeal")
        .AddJoin("a.campaign", "c.campaign")
        .AddJoin("uc.contact_number", "cm.contact_number")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & "," & vConAttrs & " FROM unmanned_collections uc, appeal_collections ac, appeals a, campaigns c, contacts cm WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & "," & ContactNameItems())
    End Sub
    Private Sub GetContactUnProcessedTransactions(ByVal pDataTable As CDBDataTable)
      Dim vIncludeBC As String = mvEnv.GetConfig("fp_batch_categories_show")
      Dim vExcludeBC As String = mvEnv.GetConfig("fp_batch_categories_hide")

      Dim vAttrs As String = "bt.batch_number,bt.transaction_number,provisional,transaction_type_desc,transaction_date,amount,payment_method_desc,reference,mailing,receipt,"
      vAttrs &= "eligible_for_gift_aid,currency_amount,bt.notes,bt.transaction_type,bt.payment_method"
      vAttrs &= If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) = True, ",b.currency_code", ",")
      vAttrs &= ",bt.transaction_origin, ts.transaction_origin_desc, ba.bank_account, ba.bank_account_desc, "

      Dim vCols As String = vAttrs
      vAttrs &= "ba.rgb_value AS rgb_bank_account, ba.rgb_value AS rgb_amount, ba.rgb_value AS rgb_currency_amount,"
      vAttrs = vAttrs.Replace("reference", mvEnv.Connection.DBSpecialCol("", "reference"))
      vCols &= "rgb_bank_account, rgb_amount, rgb_currency_amount,"

      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      vAttrs &= vAddress.GetRecordSetFieldsCountry

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("batches b", "bt.batch_number", "b.batch_number")
      vAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
      vAnsiJoins.Add("bank_accounts ba", "b.bank_account", "ba.bank_account")
      vAnsiJoins.Add("payment_methods pm", "bt.payment_method", "pm.payment_method")
      vAnsiJoins.Add("addresses a", "bt.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")

      Dim vNestedSQLStatement As SQLStatement = Nothing
      Dim vNestedWhere As CDBFields
      If mvType = DataSelectionTypes.dstSalesTransactions OrElse mvType = DataSelectionTypes.dstDeliveryTransactions Then
        Dim vNestedAttrs As String = "bta.batch_number, bta.transaction_number"
        vNestedWhere = New CDBFields()
        Select Case mvType
          Case DataSelectionTypes.dstSalesTransactions
            vNestedWhere.Add("sales_contact_number", mvContact.ContactNumber)
          Case DataSelectionTypes.dstDeliveryTransactions
            vNestedWhere.Add("contact_number", mvContact.ContactNumber)
        End Select
        vNestedSQLStatement = New SQLStatement(mvEnv.Connection, vNestedAttrs, "batch_transaction_analysis bta", vNestedWhere)
        vNestedSQLStatement.Distinct = True
      End If

      Dim vWhereFields As New CDBFields()
      Select Case mvType
        Case DataSelectionTypes.dstSalesTransactions, DataSelectionTypes.dstDeliveryTransactions
          If vNestedSQLStatement IsNot Nothing Then vAnsiJoins.Add("(" & vNestedSQLStatement.SQL & ") bta", "bt.batch_number", "bta.batch_number", "bt.transaction_number", "bta.transaction_number")
        Case DataSelectionTypes.dstContactUnProcessedTransactions
          vWhereFields.Add("bt.contact_number", mvContact.ContactNumber)
      End Select

      vAnsiJoins.AddLeftOuterJoin("transaction_origins ts", "bt.transaction_origin", "ts.transaction_origin")

      If vIncludeBC.Length > 0 Then
        vWhereFields.Add("b.batch_category", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vIncludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoIn)
      End If
      If vExcludeBC.Length > 0 Then
        vWhereFields.Add("b.batch_category#2", CDBField.FieldTypes.cftCharacter, String.Empty, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("b.batch_category#3", CDBField.FieldTypes.cftCharacter, New ArrayListEx(vExcludeBC, "|".ToCharArray).CSStringList, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoNotIn)
      End If

      vWhereFields.Add("b.posted_to_nominal", CDBField.FieldTypes.cftCharacter, "N", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("b.posted_to_nominal#2", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("b.provisional", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)

      Dim vConfirmedTransaction As New ConfirmedTransaction(mvEnv)
      vConfirmedTransaction.Init()
      vNestedSQLStatement = Nothing
      vNestedWhere = New CDBFields(New CDBField("ct.provisional_batch_number", CDBField.FieldTypes.cftInteger, "b.batch_number"))
      vNestedWhere.Add("ct.confirmed_batch_number", CDBField.FieldTypes.cftCharacter, String.Empty, CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vNestedWhere.Add("ct.status", CDBField.FieldTypes.cftCharacter, vConfirmedTransaction.GetStatusCode(ConfirmedTransaction.ConfirmedTransactionStatus.Cancelled), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vNestedSQLStatement = New SQLStatement(mvEnv.Connection, "ct.provisional_trans_number", "confirmed_transactions ct", vNestedWhere)

      vWhereFields.Add("bt.transaction_number", CDBField.FieldTypes.cftInteger, vNestedSQLStatement.SQL, CDBField.FieldWhereOperators.fwoNotIn)

      If mvParameters.Exists("BatchNumber") Then
        vWhereFields.Add("bt.batch_number", mvParameters("BatchNumber").LongValue)
        vWhereFields.Add("bt.transaction_number#2", mvParameters("TransactionNumber").LongValue)
      End If

      Dim vOrderBy As String = "bt.batch_number DESC, bt.transaction_number DESC"

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "batch_transactions bt", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols, "ADDRESS_LINE")

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Provisional")
        vRow.SetYNValue("Receipt")
        vRow.SetYNValue("EligibleForGiftAid")
        CheckAmountRGBValue(vRow)
      Next
    End Sub
    Private Sub GetContactViewOrganisations(ByVal pDataTable As CDBDataTable)

      Dim vItems As New StringBuilder
      vItems.Append("cp.organisation_number,cp.address_number,position,started,finished,cp.mail,")
      vItems.Append(mvEnv.Connection.DBSpecialCol("cp", "current"))
      vItems.Append(",position_location,position_function,position_seniority,contact_position_number,cp.amended_by,cp.amended_on,single_position,o.organisation_group,c.contact_group,o.name,")
      vItems.Append(",address_line1, address_line2, address_line3,") 'BR18024 First line of address missing from Contact > Positions > Address in table
      vItems.Append(mvContact.GetRecordSetFieldsName)
      vItems.Append(",")
      Dim vAddress As New Address(mvEnv)
      vItems.Append(vAddress.GetRecordSetFieldsDetailCountrySortCode)
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then vItems = vItems.Replace("position_function,position_seniority,", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks) Then vItems = vItems.Replace("single_position", "")

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("organisations o", "cp.organisation_number", "o.organisation_number")
      vAnsiJoins.Add("organisation_groups og", "o.organisation_group", "og.organisation_group")
      vAnsiJoins.Add("contacts c", "cp.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "cp.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      vAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number", "a.address_number", "oa.address_number")
      vAnsiJoins.Add("position_statuses ps", "cp.position_status", "ps.position_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vWhereFields As New CDBFields
      Dim vAttrs As New StringBuilder
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vWhereFields.Add("cp.organisation_number", mvContact.ContactNumber)
        vAttrs.Append("contact_number,address_number,position," & ContactNameItems())
        vItems.Append(",c.dialling_code,c.std_code,c.telephone,c.ex_directory")
      Else
        vWhereFields.Add("cp.contact_number", mvContact.ContactNumber)
        vWhereFields.Add("og.view_in_contact_card", "Y")
        vAttrs.Append("contact_Position_number,organisation_number,address_number,position," & OrgNameItems())
        vItems.Append(",o.dialling_code,o.std_code,o.telephone")
        If mvParameters.Exists("OrganisationGroup") Then vWhereFields.Add("o.organisation_group", mvParameters("OrganisationGroup").Value)
      End If

      If mvParameters.HasValue("Current") Then vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), mvParameters("Current").Value)
      If mvParameters.Exists("AddressNumber") Then vWhereFields.Add("cp.address_number", mvParameters("AddressNumber").LongValue)
      If mvParameters.Exists("ContactPositionNumber") Then vWhereFields.Add("cp.contact_position_number", mvParameters("ContactPositionNumber").LongValue)
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)

      vAttrs.Append(",started,finished,mail,current,position_location,position_status,position_function,position_seniority,single_position,organisation_group,contact_group,ADDRESS_LINE")
      vAttrs.Append(",position_status_desc,PositionFunctionDesc,PositionSeniorityDesc")

      vItems.Append(",ps.position_status,ps.position_status_desc")


      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then vAttrs = vAttrs.Replace("position_function,position_seniority,", ",,")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionLinks) Then vAttrs = vAttrs.Replace("single_position", "")

      Dim vOrderBy As New StringBuilder
      vOrderBy.Append(mvEnv.Connection.DBSpecialCol("cp", "current"))
      vOrderBy.Append(" DESC, finished DESC, surname, forenames")

      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vItems.ToString), "contact_positions cp", vWhereFields, vOrderBy.ToString, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs.ToString)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("Current")
        vRow.SetYNValue("SinglePosition")
      Next
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPositionFuntionSeniority) Then
        GetDescriptions(pDataTable, "PositionFunction")
        GetDescriptions(pDataTable, "PositionSeniority")
      End If
    End Sub
    Private Sub GetCovenantGiftAidClaims(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetCovenantGiftAidClaims")
      Dim vAttrs As String = "claim_generated_date,dtcl.claim_number,net_amount,dtcl.amount_claimed,payment_number,transaction_date,oph.amount,balance,fh.batch_number,fh.transaction_number"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM declaration_tax_claim_lines dtcl, declaration_tax_claims dtc, order_payment_history oph,financial_history fh WHERE cd_number = " & mvParameters("CovenantNumber").Value & " AND declaration_or_covenant_number = 'C' AND dtcl.claim_number = dtc.claim_number AND oph.batch_number = dtcl.batch_number AND oph.transaction_number = dtcl.transaction_number AND oph.line_number = dtcl.line_number AND fh.batch_number = oph.batch_number AND fh.transaction_number = oph.transaction_number ORDER BY claim_generated_date DESC,payment_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetCovenentClaims(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetCovenentClaims")
      Dim vAttrs As String = "claim_generated,tcl.claim_number,start_payment_number,end_payment_number,net_amount,amount_calculated,tcl.amount_claimed,tcl.amended_on,tcl.amended_by"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM tax_claim_lines tcl, tax_claims tc WHERE covenant_number = " & mvParameters("CovenantNumber").Value & " AND tcl.claim_number = tc.claim_number ORDER BY claim_generated DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetCovenentPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetCovenentPayments")
      Dim vAttrs As String = "payment_number,transaction_date,oph.amount,balance,fh.batch_number,fh.transaction_number"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM order_payment_history oph,financial_history fh WHERE order_number = " & mvParameters("PaymentPlanNumber").Value & " AND payment_number >= " & mvParameters("StartPaymentNumber").Value & " AND payment_number <= " & mvParameters("EndPaymentNumber").Value & " AND fh.batch_number = oph.batch_number AND fh.transaction_number = oph.transaction_number ORDER BY payment_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetCPDDetails(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactCPDPeriodNumber") Then vWhereFields.Add("ccp.contact_cpd_period_number", mvParameters.Item("ContactCPDPeriodNumber").LongValue)
      If mvParameters.Exists("CategoryType") Then vWhereFields.Add("ccpo.cpd_category_type", mvParameters.Item("CategoryType").Value)
      If mvParameters.Exists("ContactCPDCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", mvParameters.Item("ContactCPDCycleNumber").LongValue)
      If vWhereFields.Count > 0 Then
        Dim vAttrs As String = ""
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number", CARE.Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
        If mvParameters.Exists("CPDType") AndAlso mvParameters("CPDType").Value = "O" Then
          vAttrs = "ccc.contact_cpd_cycle_number,ccp.contact_cpd_period_number,ccb.cpd_objective_number,ccb.cpd_category_type,cpd_category_type_desc,ccb.cpd_category,cpd_category_desc,ccb.completion_date,ccb.cpd_objective_desc,ccb.supervisor_accepted,ccb.amended_on,ccb.amended_by,,,,,"
          vAnsiJoins.Add("contact_cpd_objectives ccb", "ccp.contact_cpd_period_number", "ccb.contact_cpd_period_number", CARE.Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
          vAnsiJoins.Add("cpd_category_types cct ", "ccb.cpd_category_type", "cct.cpd_category_type")
          vAnsiJoins.Add("cpd_categories cc", "ccb.cpd_category_type", "cc.cpd_category_type", "ccb.cpd_category", "cc.cpd_category")
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_cpd_cycles ccc", vWhereFields, "ccb.cpd_category_type , ccb.cpd_category, ccb.amended_on", vAnsiJoins)
          pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
        Else
          vAttrs = "ccc.contact_cpd_cycle_number,ccp.contact_cpd_period_number,ccpo.contact_cpd_point_number,ccpo.cpd_category_type,cpd_category_type_desc,ccpo.cpd_category,cpd_category_desc,ccpo.points_date,ccpo.cpd_points,ccpo.evidence_seen,ccpo.amended_on,ccpo.amended_by,ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome"
          If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAttrs = vAttrs.Replace(",ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome", ",,,,,")
          vAnsiJoins.Add("contact_cpd_points ccpo", "ccp.contact_cpd_period_number", "ccpo.contact_cpd_period_number", CARE.Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
          vAnsiJoins.Add("cpd_category_types cct ", "ccpo.cpd_category_type", "cct.cpd_category_type")
          vAnsiJoins.Add("cpd_categories cc", "ccpo.cpd_category_type", "cc.cpd_category_type", "ccpo.cpd_category", "cc.cpd_category")
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAnsiJoins.AddLeftOuterJoin("cpd_item_types cit", "ccpo.cpd_item_type", "cit.cpd_item_type")
          Dim vColumnAttrs As String = vAttrs
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber) AndAlso mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
            vAttrs = vAttrs.Replace("ccpo.cpd_points,", "CAST(ccpo.cpd_points AS int) AS cpd_points,")
            vAttrs = vAttrs.Replace("ccpo.cpd_points_2,", "CAST(ccpo.cpd_points_2 AS int) AS cpd_points_2,")
          End If
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_cpd_cycles ccc", vWhereFields, "ccpo.cpd_category_type, ccpo.cpd_category," & mvEnv.Connection.DBIsNull("points_date", mvEnv.Connection.SQLLiteral("", DateSerial(1900, 1, 1))) & " DESC, ccpo.amended_on", vAnsiJoins)
          pDataTable.FillFromSQL(mvEnv, vSQLStatement, vColumnAttrs)
        End If
      End If
    End Sub
    Private Sub GetCPDPointsEdit(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactCpdPeriodNumber") Then vWhereFields.Add("ccp.contact_cpd_period_number", mvParameters.Item("ContactCpdPeriodNumber").LongValue)
      If mvParameters.Exists("CategoryType") Then vWhereFields.Add("ccpo.cpd_category_type", mvParameters.Item("CategoryType").Value)
      If mvParameters.Exists("ContactCpdCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", mvParameters.Item("ContactCpdCycleNumber").LongValue)
      If mvParameters.Exists("ContactCpdPointNumber") Then vWhereFields.Add("ccpo.contact_cpd_point_number", mvParameters.Item("ContactCpdPointNumber").LongValue)
      If mvParameters.Exists("FromWPD") Then vWhereFields.Add("'1'", CDBField.FieldTypes.cftLong, "'2'", CDBField.FieldWhereOperators.fwoEqual)
      If mvParameters.Exists("ForPortal") AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then
        vWhereFields.Add("ccpo.web_publish", "N", CDBField.FieldWhereOperators.fwoNotEqual)
      End If
      Dim vAttrs As String = "ccc.contact_cpd_cycle_number,ccc.cpd_cycle_type,ccp.contact_cpd_period_number,ccp.start_date,ccp.end_date,ccpo.contact_cpd_point_number,ccpo.cpd_category_type,cpd_category_type_desc,ccpo.cpd_category,cpd_category_desc,ccpo.cpd_points,ccpo.evidence_seen,ccpo.amended_on,ccpo.amended_by,ccpo.points_date,contact_cpd_period_number_desc,ccpo.notes,ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAttrs = vAttrs.Replace(",ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome", ",,,,,")
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number")
      vAnsiJoins.Add("contact_cpd_points ccpo", "ccp.contact_cpd_period_number", "ccpo.contact_cpd_period_number")
      vAnsiJoins.Add("cpd_category_types cct", "ccpo.cpd_category_type", "cct.cpd_category_type")
      vAnsiJoins.Add("cpd_categories cc", "ccpo.cpd_category_type", "cc.cpd_category_type", "ccpo.cpd_category", "cc.cpd_category")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAnsiJoins.AddLeftOuterJoin("cpd_item_types cit", "ccpo.cpd_item_type", "cit.cpd_item_type")
      Dim vColumnAttrs As String = vAttrs
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber) AndAlso mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
        vAttrs = vAttrs.Replace("ccpo.cpd_points,", "CAST(ccpo.cpd_points AS int) AS cpd_points,")
        vAttrs = vAttrs.Replace("ccpo.cpd_points_2,", "CAST(ccpo.cpd_points_2 AS int) AS cpd_points_2,")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_cpd_cycles ccc", vWhereFields, "ccp.start_date DESC,ccp.end_date,ccpo.cpd_category_type,ccpo.cpd_category," & mvEnv.Connection.DBIsNull("points_date", mvEnv.Connection.SQLLiteral("", DateSerial(1900, 1, 1))) & "DESC, ccpo.amended_on", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vColumnAttrs)
      'SetCPDDurations(pDataTable)
    End Sub
    Private Sub GetCPDObjectives(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactCpdPeriodNumber") Then vWhereFields.Add("ccp.contact_cpd_period_number", mvParameters.Item("ContactCpdPeriodNumber").LongValue)
      If mvParameters.Exists("CategoryType") Then vWhereFields.Add("cco.cpd_category_type", mvParameters.Item("CategoryType").Value)
      If mvParameters.Exists("ContactCpdCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", mvParameters.Item("ContactCpdCycleNumber").LongValue)
      If mvParameters.Exists("CpdObjectiveNumber") Then vWhereFields.Add("cco.cpd_objective_number", mvParameters.Item("CpdObjectiveNumber").LongValue)
      Dim vAttrs As String = "ccc.contact_cpd_cycle_number,ccc.cpd_cycle_type,ccp.contact_cpd_period_number,ccp.start_date,ccp.end_date,cctp.default_duration,cco.cpd_objective_number,cco.cpd_objective_desc,cco.long_description,cco.cpd_category_type,cpd_category_type_desc,cco.cpd_category,cpd_category_desc,cco.completion_date,cco.target_date,cco.supervisor_contact_number,c.surname,cco.supervisor_accepted,cco.notes,cco.created_by,cco.created_on,cco.amended_on,cco.amended_by"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("cpd_cycle_types cctp", "ccc.cpd_cycle_type", "cctp.cpd_cycle_type", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("contact_cpd_objectives cco", "ccp.contact_cpd_period_number", "cco.contact_cpd_period_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("cpd_category_types cct", "cco.cpd_category_type", "cct.cpd_category_type")
      vAnsiJoins.Add("cpd_categories cc", "cco.cpd_category_type", "cc.cpd_category_type", "cco.cpd_category", "cc.cpd_category")
      vAnsiJoins.Add("contacts c", "cco.supervisor_contact_number", "c.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_cycles ccc", vWhereFields, "ccp.start_date,ccp.end_date,cco.cpd_category_type , cco.cpd_category, cco.amended_on", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetCPDSummary(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("ccc.contact_number", mvParameters.Item("ContactNumber").LongValue)
      If mvParameters.Exists("ContactCpdCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", mvParameters.Item("ContactCpdCycleNumber").LongValue)
      If mvParameters.Exists("ContactCpdPeriodNumber") Then vWhereFields.Add("ccp.contact_cpd_period_number", mvParameters.Item("ContactCpdPeriodNumber").LongValue)
      Dim vAttrs As String = "ccc.contact_cpd_cycle_number, MAX(ccc.cpd_cycle_type) AS cpd_cycle_type, MAX(cpd_cycle_type_desc) AS cpd_cycle_type_desc, MAX(start_month) AS start_month, MAX(end_month) AS end_month, MAX(ccc.start_date) AS start_date, MAX(ccc.end_date) AS end_date, ccp.contact_cpd_period_number, MAX(ccp.start_date) AS period_start_date, MAX(ccp.end_date) AS period_end_date, MAX(ccpo.cpd_category_type) AS cpd_category_type, MAX(cpd_category_type_desc) AS cpd_category_type_desc, SUM(ccpo.cpd_points) AS cpd_points, MAX(cycle_desc) AS cycle_desc, MAX(ccc.amended_on) AS amended_on, MAX(ccc.amended_by) AS amended_by, MAX(contact_cpd_period_number_desc) AS contact_cpd_period_number_desc,ccc.contact_cpd_cycle_number AS copy_contact_cpd_cycle_number, MAX(ccpo.cpd_category_type) AS copy_cpd_category_type, SUM(ccpo.cpd_points_2) AS cpd_points_2, SUM(ccpo.cpd_points) + SUM(ccpo.cpd_points_2) AS total_points"
      Dim vSelAttrs As String = "ccc.contact_cpd_cycle_number,cpd_cycle_type,cpd_cycle_type_desc,start_month,end_month,start_date,end_date,ccp.contact_cpd_period_number,period_start_date,period_end_date,cpd_category_type,cpd_category_type_desc,cpd_points,cycle_desc,amended_on,amended_by,contact_cpd_period_number_desc,copy_contact_cpd_cycle_number,copy_cpd_category_type,cpd_points_2,total_points"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then
        vAttrs = vAttrs.Replace(", SUM(ccpo.cpd_points_2) AS cpd_points_2, SUM(ccpo.cpd_points) + SUM(ccpo.cpd_points_2) AS total_points", "")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDCycleStatus) Then
        vAttrs = vAttrs & ",ccs.cpd_cycle_status,ccs.cpd_cycle_status_desc,ccs.rgb_value"
        vSelAttrs = vSelAttrs & ",ccs.cpd_cycle_status,ccs.cpd_cycle_status_desc,ccs.rgb_value"
      Else
        vAttrs = vAttrs & ",Null as cpd_cycle_status,Null as cpd_cycle_status_desc,Null as rgb_value"
        vSelAttrs = vSelAttrs & ",cpd_cycle_status,cpd_cycle_status_desc,rgb_value"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDObjective) Then
        vAttrs = vAttrs & ",ccty.cpd_type,MAX(ccty.cpd_type) as cpd_type_desc"
        vSelAttrs = vSelAttrs & ",ccty.cpd_type,cpd_type_desc"
      Else
        vAttrs = vAttrs & ",Null as cpd_type,Null as cpd_type_desc"
        vSelAttrs = vSelAttrs & ",cpd_type,cpd_type_desc"
      End If
      vAttrs &= ", ccpo.contact_cpd_point_number"
      vSelAttrs &= ",contact_cpd_point_number"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("cpd_cycle_types ccty", "ccc.cpd_cycle_type", "ccty.cpd_cycle_type")
      vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number")
      vAnsiJoins.AddLeftOuterJoin("contact_cpd_points ccpo", "ccp.contact_cpd_period_number", "ccpo.contact_cpd_period_number")
      vAnsiJoins.AddLeftOuterJoin("cpd_category_types cct", "ccpo.cpd_category_type", "cct.cpd_category_type")
      vAnsiJoins.AddLeftOuterJoin("cpd_categories cc", "ccpo.cpd_category_type", "cc.cpd_category_type", "ccpo.cpd_category", "cc.cpd_category")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDCycleStatus) Then
        vAnsiJoins.Add("cpd_cycle_statuses ccs", "ccc.cpd_cycle_status", "ccs.cpd_cycle_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber) AndAlso mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
        vAttrs = vAttrs.Replace("SUM(ccpo.cpd_points) + SUM(ccpo.cpd_points_2)", "CAST(SUM(ccpo.cpd_points) + SUM(ccpo.cpd_points_2) AS int)")
        vAttrs = vAttrs.Replace("SUM(ccpo.cpd_points)", "CAST(SUM(ccpo.cpd_points) AS int)")
        vAttrs = vAttrs.Replace("SUM(ccpo.cpd_points_2)", "CAST(SUM(ccpo.cpd_points_2) AS int)")
      End If
      Dim vSQLStatement As SQLStatement
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDObjective) Then
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_cycles ccc", vWhereFields, "ccty.cpd_type, start_date DESC,end_date DESC,period_start_date DESC, ccpo.cpd_category_type", vAnsiJoins)
      Else
        vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_cycles ccc", vWhereFields, "start_date DESC,end_date DESC,period_start_date DESC, ccpo.cpd_category_type", vAnsiJoins)
      End If
      vSQLStatement.GroupBy = "ccc.contact_cpd_cycle_number,ccp.contact_cpd_period_number,ccpo.cpd_category_type"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDCycleStatus) Then
        vSQLStatement.GroupBy = vSQLStatement.GroupBy & ",ccs.cpd_cycle_status,ccs.cpd_cycle_status_desc,ccs.rgb_value"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCPDObjective) Then
        vSQLStatement.GroupBy = vSQLStatement.GroupBy & ",ccty.cpd_type"
      End If
      vSQLStatement.GroupBy &= ", ccpo.contact_cpd_point_number"
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vSelAttrs)
      For vCtr As Integer = 0 To pDataTable.Rows.Count - 1
        If pDataTable.Rows(vCtr).Item("CPDType") = "O" Then
          pDataTable.Rows(vCtr).Item("CPDTypeDesc") = ProjectText.CPDTypeObjectives
        Else
          pDataTable.Rows(vCtr).Item("CPDTypeDesc") = ProjectText.CPDTypePoints
        End If
      Next
      'SetCPDDurations(pDataTable)
    End Sub
    Private Sub GetContactCPDPointsWithoutCycle(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("ccpo.contact_number", mvParameters("ContactNumber").IntegerValue)
        vWhereFields.Add("ccpo.contact_cpd_period_number", 0)
        If mvParameters.Exists("CategoryType") Then vWhereFields.Add("ccpo.cpd_category_type", mvParameters.Item("CategoryType").Value)
        If mvParameters.Exists("ContactCpdPointNumber") Then vWhereFields.Add("ccpo.contact_cpd_point_number", mvParameters.Item("ContactCpdPointNumber").LongValue)
        Dim vAttrs As String = "ccpo.contact_cpd_point_number,ccpo.contact_number,ccpo.cpd_category_type,cpd_category_type_desc,ccpo.cpd_category,cpd_category_desc,ccpo.cpd_points,ccpo.evidence_seen,ccpo.amended_on,ccpo.amended_by,ccpo.points_date,ccpo.notes,ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome"
        If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAttrs = vAttrs.Replace(",ccpo.cpd_points_2,ccpo.web_publish,ccpo.cpd_item_type,cit.cpd_item_type_desc,ccpo.cpd_outcome", ",,,,,")
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("cpd_category_types cct", "ccpo.cpd_category_type", "cct.cpd_category_type")
        vAnsiJoins.Add("cpd_categories cc", "ccpo.cpd_category_type", "cc.cpd_category_type", "ccpo.cpd_category", "cc.cpd_category")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2) Then vAnsiJoins.AddLeftOuterJoin("cpd_item_types cit", "ccpo.cpd_item_type", "cit.cpd_item_type")
        Dim vColumnAttrs As String = vAttrs
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber) AndAlso mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
          vAttrs = vAttrs.Replace("ccpo.cpd_points,", "CAST(ccpo.cpd_points AS int) AS cpd_points,")
          vAttrs = vAttrs.Replace("ccpo.cpd_points_2,", "CAST(ccpo.cpd_points_2 AS int) AS cpd_points_2,")
        End If
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_cpd_points ccpo", vWhereFields, mvEnv.Connection.DBIsNull("ccpo.points_date", mvEnv.Connection.SQLLiteral("", DateSerial(1900, 1, 1))) & " DESC,ccpo.cpd_category_type , ccpo.cpd_category, ccpo.amended_on", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vColumnAttrs)
      End If
    End Sub
    'Private Sub GetCriteriaSetDetails(ByVal pDataTable As CDBDataTable)
    '  Dim vWhereFields As New CDBFields
    '  Dim vAnsiJoins As New AnsiJoins()
    '  Dim vAttrs As String = "csd.criteria_set,sequence_number,search_area,i_e,c_o,main_value,subsidiary_value,period,counted,and_or,left_parenthesis,right_parenthesis"
    '  If mvParameters.Exists("MarketingControls") AndAlso mvParameters("MarketingControls").Bool Then
    '    vAnsiJoins.Add("marketing_controls mc", "csd.criteria_set", "mc.criteria_set", CARE.Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
    '  End If

    '  If mvParameters.Exists("CriteriaSet") Then vWhereFields.Add("csd.criteria_set", mvParameters.Item("CriteriaSet").LongValue)
    '  Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "criteria_set_details csd", vWhereFields, "sequence_number", vAnsiJoins)
    '  pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    'End Sub

    'Private Sub GetGeneralMailingSelectionSets(ByVal pDataTable As CDBDataTable)
    '  'SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source
    '  Dim vFields As String = "selection_set,selection_set_desc,user_name,department,number_in_set,source"
    '  Dim vWhereFields As New CDBFields()
    '  If mvParameters.Exists("SelectionSetNumber") Then vWhereFields.Add("selection_set", mvParameters("SelectionSetNumber").IntegerValue, CDBField.FieldWhereOperators.fwoEqual)
    '  If mvParameters.Exists("SelectionGroup") Then vWhereFields.Add("selection_group", mvParameters("MailingType").Value, CDBField.FieldWhereOperators.fwoEqual)
    '  If mvParameters.Exists("Department") Then vWhereFields.Add("department", mvParameters("Department").Value, CDBField.FieldWhereOperators.fwoEqual)
    '  Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "selection_sets", vWhereFields, "selection_set_desc")
    '  pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    'End Sub

    Private Sub GetCriteriaSets(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vAttrs As String = "cs.criteria_set,criteria_set_desc,user_name,department"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports) Then vAttrs = vAttrs & ",report_code,standard_document"
      Dim vSQL As String = "SELECT " & vAttrs
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports) Then
        vSQL = vSQL & "," & mvEnv.Connection.DBIsNull("c_criteria_count", "0") & "AS c_criteria_count," & mvEnv.Connection.DBIsNull("o_criteria_count", "0") & "AS o_criteria_count,mailing"
        vAttrs = vAttrs & ",c_criteria_count,o_criteria_count,mailing"
      End If
      vSQL = vSQL & " FROM criteria_sets cs"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports) Then
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT criteria_set, COUNT(*) AS c_criteria_count FROM criteria_set_details WHERE c_o = 'C' GROUP BY criteria_set) x ON cs.criteria_set = x.criteria_set"
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT criteria_set, COUNT(*) AS o_criteria_count FROM criteria_set_details WHERE c_o = 'O' GROUP BY criteria_set) y ON cs.criteria_set = y.criteria_set"
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT mailing FROM mailings) z ON cs.standard_document = z.mailing"
      End If
      If mvParameters.HasValue("CriteriaSet") Then
        vSQL = vSQL & " WHERE cs.criteria_set = " & mvParameters("CriteriaSet").LongValue
      Else
        If mvParameters.Exists("ApplicationName") Then
          If mvParameters("ApplicationName").Value = "CA" Then
            vSQL = vSQL & " WHERE criteria_group IN ('CA','ST')"
          Else
            vSQL = vSQL & " WHERE criteria_group = '" & mvParameters("ApplicationName").Value & "'"
          End If
        Else
          vSQL = vSQL & " WHERE criteria_group = 'GM'"
        End If
        vSQL = vSQL & " AND ((department = '" & mvEnv.User.Department & "') OR (user_name = '" & mvEnv.User.Logname & "'))"
        If mvParameters.Exists("Owner") Then
          vSQL = vSQL & " AND (user_name" & mvEnv.Connection.DBLikeOrEqual(mvParameters("Owner").Value) & ")"
        End If
        If mvParameters.Exists("CriteriaSetDesc") Then
          vSQL = vSQL & " AND (criteria_set_desc" & mvEnv.Connection.DBLikeOrEqual(mvParameters("CriteriaSetDesc").Value) & ")"
        End If
        If mvParameters.Exists("Department") Then
          vSQL = vSQL & " AND (department" & mvEnv.Connection.DBLikeOrEqual(mvParameters("Department").Value) & ")"
        End If
        If mvParameters.Exists("ListManager") Then
          If mvParameters("ListManager").Bool Then
            vSQL = vSQL & " AND cs.criteria_set IN (SELECT DISTINCT criteria_set FROM selection_steps)"
          Else
            vSQL = vSQL & " AND cs.criteria_set NOT IN (SELECT DISTINCT criteria_set FROM selection_steps)"
          End If
        End If
        If mvParameters.Exists("FieldName") And mvParameters.Exists("ApplicationName") Then
          vSQL = vSQL & " AND cs.criteria_set IN (SELECT DISTINCT criteria_set FROM selection_control sc, criteria_set_details csd WHERE application_name = '" & mvParameters("ApplicationName").Value & "'"
          vSQL = vSQL & " AND main_attribute = '" & mvParameters("FieldName").Value & "' AND sc.search_area = csd.search_area AND main_value " & mvEnv.Connection.DBLike("$*") & ")"
        End If
        If mvParameters.Exists("CriteriaSetNumber") Then
          vSQL = vSQL & " AND cs.criteria_set = " & mvParameters("CriteriaSetNumber").LongValue
        End If
        vSQL = vSQL & " ORDER BY criteria_set_desc"
      End If
      vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetCustomFormData(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String = ""
      Dim vSequenceNo As Integer
      Dim vSelectColumns As String
      Dim vDefaultValue As String
      Dim vValue As String
      Dim vRow As CDBDataRow
      Dim vPos As Integer

      vSelectColumns = mvCustomFieldNames
      If mvCustomForm.CustomFormUrl.Length > 0 Then
        'J1414: When CustomFormUrl is set return ShowBrowserToolbar flag & CustomFormUrl value (with parameters {Database}, {UserID} & {ContactNumber} replaced with values from the database)
        Dim vCustomFormUrl As String = mvCustomForm.CustomFormUrl
        If mvParameters.ContainsKey("MasterRecordID") Then
          vCustomFormUrl = mvCustomForm.CustomFormUrl.Replace("{Database}", mvEnv.INISection).Replace("{UserID}", mvEnv.User.UserID).Replace("{Number}", mvParameters("MasterRecordID").Value)
        Else
          vCustomFormUrl = mvCustomForm.CustomFormUrl.Replace("{Database}", mvEnv.INISection).Replace("{UserID}", mvEnv.User.UserID).Replace("{ContactNumber}", mvParameters("ContactNumber").Value)
        End If
        If vCustomFormUrl.Contains("{Postcode}") Then
          If mvParameters.ContainsKey("MasterRecordID") Then
            If mvParameters.ContainsKey("Postcode") Then
              vCustomFormUrl = vCustomFormUrl.Replace("{Postcode}", mvParameters("Postcode").ToString)
            End If
          Else
            Dim vContact As New Contact(mvEnv)
            vContact.Init(mvParameters("ContactNumber").IntegerValue)
            vCustomFormUrl = vCustomFormUrl.Replace("{Postcode}", vContact.Address.Postcode)
          End If
        End If
        If vCustomFormUrl.StartsWith("Device") Then
          Dim vDeviceValue As String = ""
          Dim vIndex As Integer = vCustomFormUrl.IndexOf("Device")
          Dim vDevice As String = TruncateString(vCustomFormUrl.Substring(vIndex + 6), 2)
          If vDevice.Length > 0 Then
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
            vWhereFields.Add("co.is_active", "Y")
            vWhereFields.Add("co.device", vDevice)
            Dim vCommsSQL As New SQLStatement(mvEnv.Connection, "number", "communications co", vWhereFields)
            vDeviceValue = vCommsSQL.GetValue
            If vDeviceValue.Length > 0 Then
              vCustomFormUrl = vDeviceValue
            Else
              vCustomFormUrl = vCustomFormUrl.Replace("Device" & vDevice, "")
            End If
          End If
        End If

        vRow = pDataTable.AddRow
        vRow.Item("CustomFormUrl") = vCustomFormUrl
        vRow.Item("ShowBrowserToolbar") = mvCustomForm.ShowBrowserToolbar
      ElseIf mvParameters.HasValue("Default") Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT attribute_name,default_value FROM custom_form_controls WHERE custom_form = " & mvCustomForm.CustomFormNumber & " AND default_value IS NOT NULL ORDER BY sequence_number")
        While vRecordSet.Fetch
          vRow = pDataTable.AddRow
          vRow.Item("ParameterName") = ProperName(vRecordSet.Fields(1).Value)
          vDefaultValue = vRecordSet.Fields(2).Value
          vValue = ""
          If Len(vDefaultValue) > 0 Then
            If vDefaultValue = "#" Or vDefaultValue = "?" Then 'Current Contact/Org number
              If mvParameters.ContainsKey("MasterRecordID") Then
                vValue = mvParameters("MasterRecordID").Value.ToString()
              Else
                vValue = CStr(mvContact.ContactNumber)
              End If
            ElseIf UCase(Left(vDefaultValue, 5)) = "#DATE" Then
              If UCase(vDefaultValue) = "#DATE" Then 'Todays Date
                vValue = TodaysDate()
              Else
                vDefaultValue = Trim(Mid(vDefaultValue, 6))
                If IsDate(vDefaultValue) Then vValue = CStr(CDate(vDefaultValue))
              End If
            ElseIf vDefaultValue = "#USER" Then  'User code
              vValue = mvEnv.User.Logname
            ElseIf UCase(Left(vDefaultValue, 7)) = "#SELECT" Then  'SQL Statement to get the Value
              If mvParameters.ContainsKey("MasterRecordID") Then
                vSQL = mvCustomForm.ParseNumberIntoSQLStatement(mvParameters("MasterRecordID").IntegerValue, Mid(vDefaultValue, 2))
              Else
                vSQL = mvCustomForm.ParseNumberIntoSQLStatement((mvContact.ContactNumber), Mid(vDefaultValue, 2))
              End If
              If Len(vSQL) > 0 Then vValue = mvCustomForm.GetFirstAttributeValue(vSQL)
            ElseIf UCase(Left(vDefaultValue, 8)) = "#CONTROL" Then  'SQL Statement to get the Value
              If mvParameters.ContainsKey("MasterRecordID") Then
                vSQL = mvCustomForm.ParseNumberIntoSQLStatement(mvParameters("MasterRecordID").IntegerValue, Mid(vDefaultValue, 10))
              Else
                vSQL = mvCustomForm.ParseNumberIntoSQLStatement((mvContact.ContactNumber), Mid(vDefaultValue, 10))
              End If
              If vSQL.Length > 0 Then
                vPos = InStr(10, vSQL, "UPDATE ", CompareMethod.Text)
                If vPos > 0 Then
                  vValue = mvCustomForm.GetFirstAttributeValue(Left(vSQL, vPos - 1))
                  mvCustomForm.ExecuteSQL(Mid(vSQL, vPos))
                End If
              End If
            ElseIf Left(vDefaultValue, 1) = "^" Then  'External Reference
              vDefaultValue = Mid(vDefaultValue, 2, Len(vDefaultValue) - 2)
              vValue = mvEnv.Connection.GetValue("SELECT external_reference FROM contact_external_links WHERE contact_number = " & mvContact.ContactNumber & " AND data_source = '" & vDefaultValue & "'")
            ElseIf Left(vDefaultValue, 1) = "#" Then
              'Grid value from Parent Display Form
              'This is unlikely to be called as the grid line will be blank to get Insert mode anyway
            Else 'Just a Hard Coded Value
              vValue = vDefaultValue
            End If
          End If
          vRow.Item("DefaultValue") = vValue
        End While
        vRecordSet.CloseRecordSet()
      ElseIf mvParameters.HasValue("Detail") Then
        If Left(mvParameters("Detail").Value, 3) = "vas" Then
          vSequenceNo = CInt(Mid(mvParameters("Detail").Value, 4))
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT desc_sql FROM custom_form_controls WHERE custom_form = " & mvCustomForm.CustomFormNumber & " AND sequence_number = " & vSequenceNo)
          If vRecordSet.Fetch Then
            vSQL = vRecordSet.Fields(1).Value
            vSelectColumns = ""
          End If
          vRecordSet.CloseRecordSet()
          If mvParameters.Exists("Values") Then
            vSQL = mvCustomForm.ParseValuesIntoSQLStatement((mvParameters("Values").Value), vSQL)
          End If
        Else
          If Len(mvCustomForm.GridSelectSql) > 0 Then
            vSQL = mvCustomForm.ParseValuesIntoSQLStatement((mvParameters("Values").Value), (mvCustomForm.GridSelectSql))
          Else
            vSQL = mvCustomForm.SelectSql
          End If
        End If
        If mvParameters.HasValue("MasterRecordID") Then
          vSQL = mvCustomForm.ParseNumberIntoSQLStatement(mvParameters("MasterRecordID").IntegerValue, vSQL)
        Else
          vSQL = mvCustomForm.ParseNumberIntoSQLStatement((mvContact.ContactNumber), vSQL)
        End If
        If Len(vSQL) > 0 Then pDataTable.FillFromSQLDB(mvEnv, (mvCustomForm.DbName), vSQL, vSelectColumns)
      ElseIf mvParameters.HasValue("MasterRecordID") Then
        vSQL = mvCustomForm.ParseNumberIntoSQLStatement(mvParameters("MasterRecordID").IntegerValue, (mvCustomForm.SelectSql))
        If Len(vSQL) > 0 Then pDataTable.FillFromSQLDB(mvEnv, (mvCustomForm.DbName), vSQL, vSelectColumns)
      Else
        vSQL = mvCustomForm.ParseNumberIntoSQLStatement((mvContact.ContactNumber), (mvCustomForm.SelectSql))
        If Len(vSQL) > 0 Then pDataTable.FillFromSQLDB(mvEnv, (mvCustomForm.DbName), vSQL, vSelectColumns)
      End If
    End Sub
    Private Sub GetDashboardData(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDashboardData")
      pDataTable = New CDBDataTable
      GetDashboardData(pDataTable)
    End Sub
    Private Sub GetDelegateActivities(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDelegateActivities")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateActivities) Then
        Dim vAttrs As String = "da.delegate_activity_number,da.event_delegate_number,da.activity,da.activity_value,quantity,activity_date,da.source,valid_from,valid_to,da.amended_by,da.amended_on,da.notes,activity_desc,activity_value_desc,source_desc"
        Dim vSQL As String = "SELECT " & vAttrs & " FROM delegate_activities da, activities a, activity_values av, sources s"
        vSQL = vSQL & " WHERE event_delegate_number = " & mvParameters("EventDelegateNumber").LongValue
        If mvParameters.Exists("Activities") Then vSQL = vSQL & " AND da.activity IN (" & mvParameters("Activities").Value & ")"
        vSQL = vSQL & " AND da.activity = a.activity AND da.activity = av.activity AND da.activity_value = av.activity_value"
        vSQL = vSQL & " AND da.source = s.source"
        vSQL = vSQL & " ORDER BY activity_desc, activity_value_desc"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & ",,,")
      End If
      pDataTable.Columns("Status").AttributeName = "status"
      pDataTable.Columns("NoteFlag").AttributeName = "note_flag"
      For Each vRow As CDBDataRow In pDataTable.Rows
        If Len(vRow.Item("Notes")) > 0 Then vRow.Item("NoteFlag") = "Y"
        vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
      Next
      pDataTable.ReOrderRowsByColumn("StatusOrder")
    End Sub
    Private Sub GetDelegateLinks(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDelegateLinks")
      Dim vOrg As New Organisation(mvEnv)
      Dim vRestriction As String = ""
      Dim vConLinkOrder As String = mvEnv.GetConfig("link_order_con")
      If Len(vConLinkOrder) = 0 Then
        If mvEnv.GetConfigOption("option_contact_groups", False) Then vConLinkOrder = "contact_group,"
        vConLinkOrder = vConLinkOrder & "surname,initials"
      End If
      Dim vOrgLinkOrder As String = mvEnv.GetConfig("link_order_org")
      If Len(vOrgLinkOrder) = 0 Then
        If mvEnv.GetConfigOption("option_organisation_groups", False) Then vOrgLinkOrder = "organisation_group,"
        vOrgLinkOrder = vOrgLinkOrder & "name"
      End If
      Dim vAttrs As String = ",valid_from,valid_to,historical,notes,amended_by,amended_on"
      Dim vSelAttrs As String = ",dl.relationship,relationship_desc,valid_from,valid_to,historical,dl.notes,dl.amended_by,dl.amended_on "
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone Or Contact.ContactRecordSetTypes.crtGroup) & vSelAttrs
      vOrg.Init()
      Dim vOrgAttrs As String = "name," & vOrg.GetRecordSetFieldsPhoneGroup & vSelAttrs
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vConAttrs = vConAttrs & ",ownership_group "
        vOrgAttrs = vOrgAttrs & ",ownership_group "
        vAttrs = vAttrs & ",ownership_group"
      Else
        vConAttrs = vConAttrs & ",c.department "
        vOrgAttrs = vOrgAttrs & ",o.department "
        vAttrs = vAttrs & ",department"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateActivities) Then
        If mvParameters.Exists("Relationships") Then vRestriction = " AND r.relationship IN (" & mvParameters("Relationships").Value & ")"
        Dim vSQL As String
        If mvParameters.Exists("EventDelegateNumber") Then
          'Get delegate_links to contacts
          vSQL = "SELECT delegate_link_number,event_delegate_number," & vConAttrs & "FROM delegate_links dl, contacts c, relationships r WHERE event_delegate_number = " & mvParameters("EventDelegateNumber").Value & " AND c.contact_number = dl.contact_number AND c.contact_type <> 'O' AND dl.relationship = r.relationship" & vRestriction & " ORDER BY " & vConLinkOrder
          pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "delegate_link_number,event_delegate_number,relationship,CONTACT_TYPE_1,CONTACT_TYPE_2,relationship_desc," & ContactNameItems() & ",contact_number,CONTACT_TELEPHONE" & vAttrs & ",contact_group")
          'Get delegate_links to organisations
          vSQL = "SELECT delegate_link_number,event_delegate_number,dl.contact_number," & vOrgAttrs & "FROM delegate_links dl, organisations o, relationships r WHERE event_delegate_number = " & mvParameters("EventDelegateNumber").Value & " AND o.organisation_number = dl.contact_number AND dl.relationship = r.relationship" & vRestriction & " ORDER BY " & vOrgLinkOrder
          pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "delegate_link_number,event_delegate_number,relationship,CONTACT_TYPE_1,ORGANISATION_TYPE_2,relationship_desc," & OrgNameItems() & ",contact_number,ORGANISATION_TELEPHONE" & vAttrs & ",organisation_group")
          For Each vRow As CDBDataRow In pDataTable.Rows
            vRow.SetYNValue("Historical")
            If vRow.Item("ContactGroup") = "" Then
              If vRow.Item("Type2") = "O" Then
                vRow.Item("ContactGroup") = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtOrganisation).EntityGroupCode
              Else
                vRow.Item("ContactGroup") = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode
              End If
            End If
          Next
          If mvParameters.Exists("ContactNumber2") Then
            Dim vRemoved As Boolean
            Do
              vRemoved = False
              For Each vRow As CDBDataRow In pDataTable.Rows
                If vRow.Item("ContactNumber") <> mvParameters("ContactNumber2").Value Then
                  pDataTable.RemoveRow(vRow)
                  vRemoved = True
                End If
              Next
            Loop While vRemoved
          End If

          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
            For Each vRow As CDBDataRow In pDataTable.Rows
              If mvEnv.User.AccessLevelFromOwnershipGroup(vRow.Item("OwnershipGroup")) < CDBEnvironment.OwnershipAccessLevelTypes.oaltRead Then
                vRow.Item("Phone") = ""
              End If
            Next
          End If
        Else
          'This is coming from the Contact Links From data selection so assume that mvContact has been instantiated
          vSQL = "SELECT " & vConAttrs & ", dl.event_delegate_number, event_desc, option_desc, '' as relationshipstatus , '' as relationshipstatusdesc, '' as rgbrelationshipstatus FROM delegate_links dl, delegates d, event_bookings eb, event_booking_options ebo, events e, contacts c, relationships r WHERE dl.contact_number = " & mvContact.ContactNumber & " AND dl.event_delegate_number = d.event_delegate_number AND d.event_number = eb.event_number AND d.booking_number = eb.booking_number AND eb.option_number = ebo.option_number AND ebo.event_number = e.event_number AND d.contact_number = c.contact_number AND dl.relationship = r.relationship " & vRestriction & "ORDER BY " & vConLinkOrder
          pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "relationship,CONTACT_TYPE_1,CONTACT_TYPE_2," & ContactNameItems() & ",contact_number,relationship_desc,CONTACT_TELEPHONE" & vAttrs & ",contact_group,event_delegate_number,event_desc,option_desc,relationshipstatus,relationshipstatusdesc,rgbrelationshipstatus")
        End If
      End If
    End Sub
    Private Sub GetDepartmentActivities(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("au.department", mvEnv.User.Department)
      vWhereFields.AddJoin("a.activity", "au.activity")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "a.activity,activity_desc,high_profile,profile_rating,contact_group", "activities a, activity_users au", vWhereFields, "a.activity")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetDepartmentActivityValues(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("av.activity", mvParameters("Activity").Value)
      vWhereFields.AddJoin("av.activity#1", "avu.activity")
      vWhereFields.AddJoin("av.activity_value", "avu.activity_value")
      vWhereFields.Add("avu.department", mvEnv.User.Department)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "av.activity_value,activity_value_desc", "activity_values av, activity_value_users avu", vWhereFields, "av.activity_value")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetDespatchStock(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDespatchStock")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "batch_number,transaction_number,line_number,s.product,product_desc,allocated,s.warehouse,warehouse_desc"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM issued_stock s INNER JOIN products p ON s.product = p.product "
      vSQL = vSQL & "LEFT OUTER JOIN warehouses w ON s.warehouse = w.warehouse WHERE "
      If mvParameters.Exists("DespatchNoteNumber") Then
        vWhereFields.Add("picking_list_number", mvParameters("PickingListNumber").LongValue)
        vWhereFields.Add("despatch_note_number", mvParameters("DespatchNoteNumber").LongValue)
      Else
        vWhereFields.Add("batch_number", mvParameters("BatchNumber").LongValue)
        vWhereFields.Add("transaction_number", mvParameters("TransactionNumber").LongValue)
        If mvParameters.Exists("PickingListNumber") Then
          If mvParameters("PickingListNumber").LongValue > 0 Then vWhereFields.Add("picking_list_number", mvParameters("PickingListNumber").LongValue)
        End If
        If mvParameters.Exists("LineNumber") Then
          If mvParameters("LineNumber").LongValue > 0 Then vWhereFields.Add("line_number", mvParameters("LineNumber").LongValue)
        End If
      End If
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY line_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs)
    End Sub
    Private Sub GetDocumentActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,,,,,,,,,,,duration_days,duration_hours,duration_minutes,a.document_class,action_text"
      Dim vWhereFields As New CDBFields()
      If mvParameters.HasValue("ActionNumber") Then vWhereFields.Add("a.action_number", mvParameters("ActionNumber").IntegerValue)
      vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department", mvEnv.User.Department)
      vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      vWhereFields.Add("da.document_number", mvParameters("DocumentNumber").Value)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("users u", "a.created_by", "u.logname")
      vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
      vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
      vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")
      vAnsiJoins.Add("document_actions da", "a.action_number", "da.action_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "actions a", vWhereFields, "sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
    End Sub
    Private Sub GetDocumentContactLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("C")
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = "link_type," & vContact.GetRecordSetFieldsName & ",t1.address_number,notified,processed,'C' AS entity_type, " & mvEnv.Connection.DBIsNull("cg.name", "'" & vEntityDesc & "'") & " AS name"
      Dim vCols As String = "link_type,contact_number,address_number,CONTACT_NAME,notified,processed"

      If pAddType Then vCols &= ",CONTACT_TYPE_1,LINK_TYPE"
      vCols &= ",entity_type,name"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "t1.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("contact_groups cg", "c.contact_group", "cg.contact_group")

      Dim vWhereFields As New CDBFields(New CDBField("t1.communications_log_number", mvParameters("DocumentNumber").LongValue))
      vWhereFields.Add("c.contact_type", CDBField.FieldTypes.cftCharacter, "O", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("cg.client", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("cg.client#2", CDBField.FieldTypes.cftCharacter, mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log_links t1", vWhereFields, "surname, initials", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified")
        vRow.SetYNValue("Processed")
      Next
    End Sub
    Private Sub GetDocumentDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      GetDocumentDocumentLinks(pDataTable, pAddType, False)
    End Sub
    Private Sub GetDocumentDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean, pIncludeEmail As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("D")
      Dim vAttrs As String = "communications_log_number, our_reference"
      Dim vCols As String = "communications_log_number, our_reference"

      Dim vCommsLog As New CommunicationsLog(mvEnv)
      If pAddType Then
        vAttrs = "'R' AS link_type," & vCommsLog.GetRecordSetFieldsDetail
        vCols = "link_type,communications_log_number,,DOCUMENT_NAME,,,,LINK_TYPE"
      End If
      vAttrs &= ",'D' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
      vCols &= ",entity_type,entity_type_desc"

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("communications_log cl", "cl.communications_log_number", "cldl.communications_log_number_2")})
      Dim vWhereFields As New CDBFields({New CDBField("communications_log_number_1", mvParameters("DocumentNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)})
      If Not pIncludeEmail Then
        vAnsiJoins.Add("packages pk", "cl.package", "pk.package", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("pk.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
        vAnsiJoins.Add("document_types dt", "cl.document_type", "dt.document_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("dt.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log_doc_links cldl", vWhereFields, "communications_log_number_2", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vCols)

    End Sub
    Private Sub GetDocumentHistory(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDocumentHistory")
      Dim vSQL As String = mvEnv.Connection.GetSelectSQLCSC & "action_date, action_time, action, user_name, notes FROM communications_log_history WHERE communications_log_number = " & mvParameters("DocumentNumber").LongValue & " ORDER BY action_date DESC, action_time DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
      For Each vRow As CDBDataRow In pDataTable.Rows
        Dim vAttr As String = vRow.Item("ActionTime")
        vRow.Item("ActionTime") = Left$(vAttr, 2) & ":" & Mid$(vAttr, 3, 2) & ":" & Right$(vAttr, 2)
      Next
    End Sub
    Private Sub GetDocumentLinks(ByVal pDataTable As CDBDataTable, pIncludeEmail As Boolean)
      GetDocumentContactLinks(pDataTable, True)
      GetDocumentOrganisationLinks(pDataTable, True)
      GetDocumentDocumentLinks(pDataTable, True, pIncludeEmail)
      GetDocumentTransactionLinks(pDataTable, True)
      GetDocumentEventLinks(pDataTable, True)
      GetDocumentLinksForExamUnits(pDataTable, True)
      GetDocumentLinksForExamCentre(pDataTable, True)
      GetDocumentLinksForExamCentreUnit(pDataTable, True)
      GetDocumentLinksForFundraisingRequest(pDataTable, True) 'BR19023
      GetDocumentLinksForWorkstreams(pDataTable, True)
      GetDocumentLinksForCPDCyclePeriods(pDataTable, True)
      GetDocumentLinksForCPDPoints(pDataTable, True)
      GetDocumentLinksForPositions(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("LinkType") = "R" Then vRow.Item("LinkType") = "Z"
      Next
      pDataTable.ReOrderRowsByColumn(("LinkType"))
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("LinkType") = "Z" Then vRow.Item("LinkType") = "R"
      Next
    End Sub
    Private Sub GetDocumentOrganisationLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("O")
      Dim vAttrs As String = "link_type,o.organisation_number,t1.address_number,o.name,notified,processed,'O' AS entity_type, " & mvEnv.Connection.DBIsNull("og.name", "'" & vEntityDesc & "'") & " AS entity_type_desc"
      Dim vCols As String = "link_type,organisation_number,address_number,name,notified,processed"

      If pAddType Then vCols &= ",ORGANISATION_TYPE_1,LINK_TYPE"
      vCols &= ",entity_type,entity_type_desc"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("organisations o", "t1.contact_number", "o.organisation_number")
      vAnsiJoins.AddLeftOuterJoin("organisation_groups og", "o.organisation_group", "og.organisation_group")

      Dim vWhereFields As New CDBFields(New CDBField("t1.communications_log_number", mvParameters("DocumentNumber").LongValue))
      vWhereFields.Add("og.client", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("og.client#2", CDBField.FieldTypes.cftCharacter, mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log_links t1", vWhereFields, "o.name", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified")
        vRow.SetYNValue("Processed")
      Next
    End Sub
    Private Sub GetDocumentRelatedDocuments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDocumentRelatedDocuments")
      Dim vSQL As String = "SELECT 'Y' AS attach,cl.communications_log_number,our_reference,subject,dt.document_source,dt.word_processor_document,precis,docfile_extension"
      vSQL = vSQL & " FROM communications_log_doc_links cldl INNER JOIN communications_log cl ON communications_log_number_2 = communications_log_number INNER JOIN document_types dt ON cl.document_type = dt.document_type LEFT OUTER JOIN packages p ON cl.package = p.package "
      vSQL = vSQL & " WHERE communications_log_number_1 = " & mvParameters("DocumentNumber").LongValue
      vSQL = vSQL & " AND dt.document_source IS NOT NULL AND NOT ( dt.document_source = 'W' AND dt.word_processor_document = 'N' AND precis IS NULL ) ORDER BY communications_log_number_2"
      vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub

    Private Sub GetDocuments(ByVal pDataTable As CDBDataTable)
      GetDocuments(pDataTable, False)
    End Sub
    Private Sub GetDocuments(ByVal pDataTable As CDBDataTable, pIncludeEmail As Boolean)
      Dim vEventDocs As Boolean = mvParameters.Exists("EventNumber")
      Dim vMeetingDocs As Boolean = mvParameters.Exists("MeetingNumber")
      Dim vDocLinks As Boolean = mvParameters.Exists("CommunicationsLogNumber1")
      Dim vActionDocs As Boolean = mvParameters.Exists("ActionNumber")
      Dim vFundraisingDocs As Boolean = mvParameters.Exists("FundraisingRequestNumber")
      Dim vContactPositionDocs As Boolean = mvParameters.Exists("ContactPositionNumber")

      'Exams Changes
      Dim vExamUnitLinks As Boolean = mvParameters.Exists("ExamUnitLinkId")
      Dim vExamCentre As Boolean = mvParameters.Exists("ExamCentreId")
      Dim vExamCentreUnit As Boolean = mvParameters.Exists("ExamCenterUnitId")

      Dim vCPDDocuments As Boolean = False
      Select Case mvType
        Case DataSelectionTypes.dstContactCPDCycleDocuments
          If mvParameters.Exists("ContactCPDPeriodNumber") OrElse mvParameters.Exists("ContactCPDCycleNumber") Then vCPDDocuments = True
        Case DataSelectionTypes.dstContactCPDPointDocuments
          vCPDDocuments = True
      End Select

      Dim vOutstandingDocs As Boolean
      If mvParameters.Exists("Notified") Or mvParameters.Exists("Processed") Then vOutstandingDocs = True
      Dim vHistoryDocs As Boolean = mvParameters.Exists("HistoryItems")
      Dim vContactDocs As Boolean
      If mvParameters.Exists("LinkType") Or mvParameters.Exists("ContactNumber") Then vContactDocs = True

      Dim vAttrs As String = "dated,cl.communications_log_number,cl.package,label_name,c.contact_number,document_type_desc,created_by,department_desc,our_reference,direction,their_reference,cl.document_type,cl.document_class,document_class_desc,standard_document,cl.source,recipient,forwarded,archiver,completed,cls.topic,topic_desc,cls.sub_topic,sub_topic_desc,creator_header,department_header,public_header,d.department,creator_header AS access_level,standard_document AS standard_document_desc"
      Select Case mvType
        Case DataSelectionTypes.dstContactDocuments, DataSelectionTypes.dstDistinctDocuments, DataSelectionTypes.dstDistinctExternalDocuments,
             DataSelectionTypes.dstFundraisingDocuments
          vAttrs &= ",precis"
      End Select
      vAttrs &= ",subject,call_duration,total_duration,selection_set,original_uri"

      Select Case mvType
        Case DataSelectionTypes.dstContactCPDCycleDocuments
          vAttrs &= ",cpp.contact_cpd_cycle_number, cpp.contact_cpd_period_number, cpp.contact_cpd_period_number_desc"
        Case DataSelectionTypes.dstContactCPDPointDocuments
          vAttrs &= ",cpo.contact_cpd_point_number"
        Case DataSelectionTypes.dstContactNotifications
          'Remove topic fields as they are not required for notifications and can cause duplicates
          vAttrs = vAttrs.Replace("cls.sub_topic,", "").Replace("sub_topic_desc,", "")
          vAttrs = vAttrs.Replace("cls.topic,", "").Replace("topic_desc,", "")
      End Select

      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields()

      Select Case DataSelectionType
        Case DataSelectionTypes.dstDocuments, DataSelectionTypes.dstDistinctDocuments, DataSelectionTypes.dstDistinctExternalDocuments,
             DataSelectionTypes.dstEventDocuments, DataSelectionTypes.dstFundraisingDocuments, DataSelectionTypes.dstContactCPDCycleDocuments,
             DataSelectionTypes.dstContactCPDPointDocuments, DataSelectionTypes.dstContactNotifications, DataSelectionTypes.dstContactPositionDocuments
          Dim vUseFindCriteria As Boolean = False
          If vEventDocs Then
            vAnsiJoins.Add("event_documents ed", "ed.communications_log_number", "cl.communications_log_number")
            vWhereFields.Add("ed.event_number", mvParameters("EventNumber").LongValue)
          ElseIf vMeetingDocs Then
            vAnsiJoins.Add("meeting_documents md", "md.communications_log_number", "cl.communications_log_number")
            vWhereFields.Add("md.meeting_number", mvParameters("MeetingNumber").LongValue)
          ElseIf vDocLinks Then
            vAnsiJoins.Add("communications_log_doc_links cldl", "cldl.communications_log_number_2", "cl.communications_log_number")
            vWhereFields.Add("cldl.communications_log_number_1", mvParameters("CommunicationsLogNumber1").LongValue)
            If (mvType = DataSelectionTypes.dstDistinctDocuments Or mvType = DataSelectionTypes.dstDistinctExternalDocuments) AndAlso mvParameters.Exists("FindRelatedDocuments") Then vUseFindCriteria = mvParameters("FindRelatedDocuments").Bool
          ElseIf vActionDocs Then
            vAnsiJoins.Add("document_actions da", "da.document_number", "cl.communications_log_number")
            vWhereFields.Add("da.action_number", mvParameters("ActionNumber").LongValue)
          ElseIf vHistoryDocs Then
            vWhereFields.Add("cl.communications_log_number", CDBField.FieldTypes.cftInteger, mvParameters("HistoryItems").Value, CDBField.FieldWhereOperators.fwoIn)
          ElseIf vExamUnitLinks Then
            vWhereFields.Add("dl.exam_unit_link_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamUnitLinkId").Value)
          ElseIf vExamCentreUnit Then
            vWhereFields.Add("dl.exam_centre_unit_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamCentreUnitId").Value)
          ElseIf vExamCentre Then
            vWhereFields.Add("dl.exam_centre_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamCentreId").Value)
          ElseIf vFundraisingDocs Then
            vWhereFields.Add("dl.fundraising_request_number", CDBField.FieldTypes.cftInteger, mvParameters("FundraisingRequestNumber").Value)
            vUseFindCriteria = True
          ElseIf vCPDDocuments Then
            If mvParameters.Exists("ContactCPDPeriodNumber") Then
              vWhereFields.Add("dl.contact_cpd_period_number", mvParameters("ContactCPDPeriodNumber").IntegerValue)
            ElseIf mvParameters.Exists("ContactCPDPointNumber") Then
              vWhereFields.Add("dl.contact_cpd_point_number", mvParameters("ContactCPDPointNumber").IntegerValue)
            End If
          ElseIf vContactPositionDocs Then
            vWhereFields.Add("dl.contact_position_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactPositionNumber").Value)
          Else
            vUseFindCriteria = True
            If vOutstandingDocs Then
              vAnsiJoins.Add("communications_log_links cll", "cll.communications_log_number", "cl.communications_log_number")
              vWhereFields.Add("cll.contact_number", mvEnv.User.ContactNumber)
              If mvParameters.Exists("Notified") Then
                vWhereFields.Add("notified", mvParameters("Notified").Value)
              End If
              If mvParameters.Exists("Processed") Then
                vWhereFields.Add("processed", mvParameters("Processed").Value)
              End If
            ElseIf vContactDocs Then
              vAnsiJoins.Add("communications_log_links cll", "cll.communications_log_number", "cl.communications_log_number")
              vWhereFields.Add("cll.contact_number", mvParameters("ContactNumber").LongValue)
              If mvParameters.Exists("DocumentLinkType") Then
                vWhereFields.Add("link_type", mvParameters("DocumentLinkType").Value, CDBField.FieldWhereOperators.fwoInOrEqual)
              Else
                vWhereFields.Add("link_type", "'A','S','R'", CDBField.FieldWhereOperators.fwoIn)
              End If
            End If
          End If
          If vUseFindCriteria Then
            If mvParameters.Exists("DocumentNumbers") Then vWhereFields.Add("cl.communications_log_number", mvParameters("DocumentNumbers").Value, CDBField.FieldWhereOperators.fwoIn)
            If mvParameters.Exists("DocumentNumber") Then vWhereFields.Add("cl.communications_log_number", mvParameters("DocumentNumber").LongValue)
            If mvParameters.Exists("OurReference") Then vWhereFields.Add("our_reference", mvParameters("OurReference").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
            If mvParameters.Exists("TheirReference") Then vWhereFields.Add("their_reference", mvParameters("TheirReference").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
            If mvParameters.Exists("Source") Then vWhereFields.Add("cl.source", mvParameters("Source").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
            If mvParameters.Exists("DocumentSubject") Then vWhereFields.Add("subject", mvParameters("DocumentSubject").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
            If mvParameters.Exists("Precis") Then vWhereFields.Add("precis", mvParameters("Precis").Value & "*", CDBField.FieldWhereOperators.fwoLike)

            AddWhereFieldFromParameter(vWhereFields, "Package", "cl.package")
            AddWhereFieldFromParameter(vWhereFields, "StandardDocument", "cl.standard_document")
            AddWhereFieldFromParameter(vWhereFields, "DocumentType", "cl.document_type")
            AddWhereFieldFromParameter(vWhereFields, "DocumentClass", "cl.document_class")
            AddWhereFieldFromParameter(vWhereFields, "Department", "cl.department")

            If mvParameters.Exists("DatedOnOrAfter") Then vWhereFields.Add("dated", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrAfter").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            If mvParameters.Exists("DatedOnOrBefore") Then vWhereFields.Add("cl.dated", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrBefore").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
            If mvParameters.Exists("Topic") Then
              Dim vSubWhereFields As New CDBFields()
              vSubWhereFields.Add("clss.topic", mvParameters("Topic").Value)
              If mvParameters.Exists("SubTopic") Then vSubWhereFields.Add("clss.sub_topic", mvParameters("SubTopic").Value)
              Dim vSubSelect As New SQLStatement(mvEnv.Connection, "communications_log_number", "communications_log_subjects clss", vSubWhereFields)
              vWhereFields.Add("cl.communications_log_number#1", CDBField.FieldTypes.cftInteger, vSubSelect.SQL, CDBField.FieldWhereOperators.fwoIn)
            Else
              If vWhereFields.Count = 0 Then RaiseError(DataAccessErrors.daeNoSelectionData)
            End If
          End If
          If mvType = DataSelectionTypes.dstDistinctExternalDocuments Then
            vAnsiJoins.Add("packages pa", "cl.package", "pa.package")
          End If
          If mvType = DataSelectionTypes.dstDistinctExternalDocuments Then
            vWhereFields.Add(New CDBField("pa.storage_type", CDBField.FieldTypes.cftCharacter, "E", CDBField.FieldWhereOperators.fwoEqual))
          End If

        Case Else
          vAnsiJoins.Add("communications_log_links cll", "cll.communications_log_number", "cl.communications_log_number")
          If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            vAnsiJoins.Add("organisation_addresses oa", "oa.address_number", "cll.address_number")
            vWhereFields.Add("oa.organisation_number", mvContact.ContactNumber)
            vWhereFields.Add("link_type", "'A', 'S', 'R'", CDBField.FieldWhereOperators.fwoIn)
          Else
            vWhereFields.Add("cll.contact_number", mvContact.ContactNumber)
            vWhereFields.Add("link_type", "'A', 'S', 'R'", CDBField.FieldWhereOperators.fwoIn)
          End If
          If mvParameters.Exists("DocumentNumber") Then vWhereFields.Add("cl.communications_log_number", mvParameters("DocumentNumber").LongValue)

      End Select

      If mvType <> DataSelectionTypes.dstContactNotifications Then
        vWhereFields.Add("primary", "Y").SpecialColumn = True
        vAnsiJoins.Add("communications_log_subjects cls", "cl.communications_log_number", "cls.communications_log_number")
      End If
      vAnsiJoins.Add("contacts c", "c.contact_number", "cl.contact_number")
      vAnsiJoins.Add("document_types dt", "dt.document_type", "cl.document_type")
      vAnsiJoins.Add("document_classes dc", "dc.document_class", "cl.document_class")
      vAnsiJoins.Add("departments d", "d.department", "cl.department")
      If mvType <> DataSelectionTypes.dstContactNotifications Then
        vAnsiJoins.Add("topics t", "t.topic", "cls.topic")
        vAnsiJoins.Add("sub_topics st", "st.topic", "t.topic", "st.sub_topic", "cls.sub_topic")
      End If
      If vExamCentre OrElse vExamCentreUnit OrElse vExamUnitLinks OrElse vFundraisingDocs OrElse vCPDDocuments OrElse vContactPositionDocs Then
        vAnsiJoins.Add("document_log_links dl", "cl.communications_log_number", "dl.communications_log_number")
      End If

      Select Case mvType
        Case DataSelectionTypes.dstContactCPDCycleDocuments
          vAnsiJoins.Add("contact_cpd_periods cpp", "dl.contact_cpd_period_number", "cpp.contact_cpd_period_number")
          If mvParameters.Exists("ContactCPDPeriodNumber") = False Then vWhereFields.Add("cpp.contact_cpd_cycle_number", mvParameters("ContactCPDCycleNumber").IntegerValue) 'Need to select all documents for the CPD Cycle rather than for the individual CPD Cycle Period

        Case DataSelectionTypes.dstContactCPDPointDocuments
          vAnsiJoins.Add("contact_cpd_points cpo", "dl.contact_cpd_point_number", "cpo.contact_cpd_point_number")

      End Select

      If Not pIncludeEmail Then
        If Not vAnsiJoins.ContainsJoinToTable("packages pa") Then
          vAnsiJoins.Add("packages pa", "cl.package", "pa.package", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        End If
        vWhereFields.Add("pa.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
        vWhereFields.Add("dt.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "communications_log cl", vWhereFields, "dated DESC, cl.communications_log_number DESC", vAnsiJoins)
      If Not (mvType = DataSelectionTypes.dstContactDocuments Or mvType = DataSelectionTypes.dstDistinctDocuments Or mvType = DataSelectionTypes.dstDistinctExternalDocuments Or mvType = DataSelectionTypes.dstFundraisingDocuments) Then vSQLStatement.Distinct = True
      If mvParameters.Exists("NumberOfRows") Then
        vSQLStatement.RecordSetOptions = CDBConnection.RecordSetOptions.NoDataTable
        pDataTable.MaximumRows = mvParameters("NumberOfRows").IntegerValue + 1
        pDataTable.CheckAccess = True
      End If
      If mvType = DataSelectionTypes.dstContactDocuments Or mvType = DataSelectionTypes.dstDistinctDocuments Or mvType = DataSelectionTypes.dstDistinctExternalDocuments Or mvType = DataSelectionTypes.dstFundraisingDocuments Then vAttrs = Replace(vAttrs, "cl.communications_log_number", "DISTINCT_DOCUMENT_NUMBER")
      vAttrs = vAttrs.Replace("creator_header AS access_level", "ACCESS")
      vAttrs = vAttrs.Replace("standard_document AS ", "")
      If mvType = DataSelectionTypes.dstContactNotifications Then
        'Contact Notifications only has these columns in the DataTable
        vAttrs = "communications_log_number,link_type,dated,our_reference,document_type_desc,subject,,,ACCESS"
      End If
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
      'Set document access is now done in the DataTable pDataTable.SetDocumentAccess()
      If DataSelectionType <> DataSelectionTypes.dstContactNotifications Then
        GetDescriptions(pDataTable, "StandardDocument")
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("Direction") = "I" Then
            vRow.Item("Direction") = DataSelectionText.String18677    'In
          Else
            vRow.Item("Direction") = DataSelectionText.String18678    'Out
          End If
          If vRow.Item("CallDuration").Length > 0 Then
            vRow.Item("CallDuration") = vRow.Item("CallDuration").Insert(2, ":").Insert(5, ":")
            vRow.Item("TotalDuration") = vRow.Item("TotalDuration").Insert(2, ":").Insert(5, ":")
          End If
        Next
      End If

    End Sub
    Private Sub GetDocumentSubjects(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cls.topic,topic_desc,cls.sub_topic,sub_topic_desc,quantity,primary,activity,activity_value,activity_duration,cls.amended_on,cls.amended_by"
      vAttrs = vAttrs.Replace("primary", mvEnv.Connection.DBSpecialCol("", "primary"))
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("topics t", "cls.topic", "t.topic")
      vAnsiJoins.Add("sub_topics st", "cls.topic", "st.topic", "cls.sub_topic", "st.sub_topic")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cls.communications_log_number", mvParameters("DocumentNumber").LongValue)
      AddWhereFieldFromParameter(vWhereFields, "Topic", "cls.topic")
      AddWhereFieldFromParameter(vWhereFields, "SubTopic", "cls.sub_topic")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log_subjects cls", vWhereFields, mvEnv.Connection.DBSpecialCol("", "primary") & " DESC, topic_desc, sub_topic_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetMeetings(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "meeting_number,meeting_desc,meeting_date,meeting_type,meeting_location,duration_days,duration_hours,duration_minutes,preamble,notes,agenda,communications_log_number,master_action,amended_by,amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataOwnerContactNumber) Then
        vAttrs = vAttrs & ",owner_contact_number"
      End If

      'Dim vAnsiJoins As New AnsiJoins()
      'vAnsiJoins.Add("document_classes dc", "cl.document_class", "dc.document_class")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("meeting_number", mvParameters("MeetingNumber").IntegerValue)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "meetings", vWhereFields, "")
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetContactMeetings(ByVal pDataTable As CDBDataTable)

      'BR19438 Re-written for new functionality. Meetings will be visible for contacts that are invited to meetings, not just meeting owners. 
      Dim vSubSelectAttrs As String = "DISTINCT m.meeting_number"
      Dim vSubSelectAnsiJoins As New AnsiJoins()
      vSubSelectAnsiJoins.AddLeftOuterJoin("meeting_links mk", "mk.meeting_number", "m.meeting_number")
      Dim vSubSelectWhereFields As New CDBFields()
      vSubSelectWhereFields.Add("owner_contact_number", mvContact.ContactNumber, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vSubSelectWhereFields.Add("mk.contact_number", mvContact.ContactNumber, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vSubSelectWhereFields.Add("mk.link_type", "R", CDBField.FieldWhereOperators.fwoOpenBracket)
      vSubSelectWhereFields.Add("mk.link_type#1", "W", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Dim vSubSelectSQL As New SQLStatement(mvEnv.Connection, vSubSelectAttrs, "meetings m", vSubSelectWhereFields, "", vSubSelectAnsiJoins)

      Dim vAttrs As String = "m.meeting_number,meeting_desc,meeting_date,m.meeting_type,meeting_type_desc,m.meeting_location,meeting_location_desc,duration_days,duration_hours,duration_minutes,preamble,m.notes,m.agenda,communications_log_number,master_action,m.amended_by,m.amended_on,m.owner_contact_number"
      Dim vItems As String = vAttrs & ",CONTACT_NAME"
      vAttrs = vAttrs & ", " & mvContact.GetRecordSetFieldsName("c")

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("meeting_types mt", "mt.meeting_type", "m.meeting_type")
      vAnsiJoins.Add("meeting_locations ml", "ml.meeting_location", "m.meeting_location")
      vAnsiJoins.Add("contacts c", "m.owner_contact_number", "c.contact_number")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("m.meeting_number", CDBField.FieldTypes.cftInteger, vSubSelectSQL.SQL, CDBField.FieldWhereOperators.fwoIn)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "meetings m", vWhereFields, "", vAnsiJoins)

      pDataTable.FillFromSQL(mvEnv, vSQL, vItems)

    End Sub
    Private Sub GetCommunicationsLogDocClass(ByVal pDataTable As CDBDataTable)
      Dim vAccessRights As New AccessRights()
      vAccessRights.Init(mvEnv)

      Dim vAttrs As String = "cl.document_class,standard_document,in_use_by,created_by,dated,package,document_type,department," & vAccessRights.AccessAttributes()
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("document_classes dc", "cl.document_class", "dc.document_class")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("communications_log_number", mvParameters("DocumentNumber").IntegerValue)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log cl", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetDocumentTransactionLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("T")
      Dim vAttrs As String = "communications_log_number,batch_number,transaction_number,'T' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
      Dim vCols As String = "batch_number,TRANSACTION_REFERENCE"

      If pAddType Then
        vAttrs = "'R' AS link_type," & vAttrs
        vCols = "link_type,batch_number,transaction_number,TRANSACTION_REFERENCE,,,TRANSACTION_TYPE_1,LINK_TYPE"
      End If
      vCols &= ",entity_type,entity_type_desc"

      Dim vWhereFields As New CDBFields(New CDBField("communications_log_number", mvParameters("DocumentNumber").LongValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "communications_log_trans", vWhereFields, "batch_number,transaction_number")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
    End Sub
    Private Sub GetDocumentEventLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("E")
      Dim vAttrs As String = "communications_log_number,ed.event_number,event_reference,event_desc,'E' AS entity_type," & mvEnv.Connection.DBIsNull("eg.name", "'" & vEntityDesc & "'") & " AS name"
      Dim vCols As String = "event_number,EVENT_REFERENCE"

      If pAddType Then
        vAttrs = "'R' AS link_type," & vAttrs
        vCols = "link_type,event_number,,EVENT_REFERENCE,,,EVENT_TYPE,LINK_TYPE"
      End If
      vCols &= ",entity_type,name"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("events e", "ed.event_number", "e.event_number")
      vAnsiJoins.AddLeftOuterJoin("event_groups eg", "e.event_group", "eg.event_group")

      Dim vWhereFields As New CDBFields(New CDBField("communications_log_number", mvParameters("DocumentNumber").LongValue))
      vWhereFields.Add("eg.client", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("eg.client#2", CDBField.FieldTypes.cftCharacter, mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_documents ed", vWhereFields, "ed.event_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
    End Sub
    Private Sub GetDocumentLinksForExamUnits(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)

      'LinkType,ContactNumber,AddressNumber,ContactName,Notified,Processed,ContactType,LinkTypeDesc
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDocumentLogLinks) Then
        Dim vEntityDesc As String = GetLinkEntityTypeDescription("U")
        Dim vFields As String = "'R' as link_type,dl.communications_log_number,eul.exam_unit_link_id,eu.exam_unit_code,eu.exam_unit_description"
        vFields &= ",'U' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
        Dim vCols As String = "link_type,exam_unit_link_id,,exam_unit_description,,,EXAM_UNIT_TYPE,LINK_TYPE,entity_type,entity_type_desc"
        Dim vAnsiJoin As New AnsiJoins
        Dim vWhereFields As New CDBFields

        vAnsiJoin.Add("exam_unit_links eul", "dl.exam_unit_link_id", "eul.exam_unit_link_id")
        vAnsiJoin.Add("exam_units eu", "eul.exam_unit_id_2", "eu.exam_unit_id")

        vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
        vWhereFields.Add("dl.exam_unit_link_id", "0", CDBField.FieldWhereOperators.fwoNotEqual)

        Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "eu.exam_unit_description", vAnsiJoin)

        pDataTable.FillFromSQL(mvEnv, vSqlQuery, vCols)
      End If
    End Sub

    Private Sub GetDocumentLinksForExamCentre(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDocumentLogLinks) Then
        Dim vEntityDesc As String = GetLinkEntityTypeDescription("N")
        Dim vFields As String = "'R' as link_type,communications_log_number,ec.exam_centre_id,ec.exam_centre_code,ec.exam_centre_description"
        vFields &= ",'N' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
        Dim vCols As String = "link_type,exam_centre_id,,exam_centre_description,,,EXAM_CENTRE_TYPE,LINK_TYPE,entity_type,entity_type_desc"
        Dim vAnsiJoin As New AnsiJoins
        Dim vWhereFields As New CDBFields

        vAnsiJoin.Add("exam_centres ec", "dl.exam_centre_id", "ec.exam_centre_id")

        vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
        vWhereFields.Add("dl.exam_centre_id", "0", CDBField.FieldWhereOperators.fwoNotEqual)

        Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "ec.exam_centre_description", vAnsiJoin)

        pDataTable.FillFromSQL(mvEnv, vSqlQuery, vCols)
      End If
    End Sub

    Private Sub GetDocumentLinksForExamCentreUnit(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDocumentLogLinks) Then
        Dim vEntityDesc As String = GetLinkEntityTypeDescription("X")
        Dim vFields As String = "'R' as link_type,dl.communications_log_number,ecu.exam_centre_unit_id,eu.exam_unit_code,eu.exam_unit_description"
        vFields &= ",'X' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
        Dim vCols As String = "link_type,exam_centre_unit_id,,exam_unit_description,,,EXAM_CENTRE_UNIT,LINK_TYPE,entity_type,entity_type_desc"
        Dim vAnsiJoin As New AnsiJoins
        Dim vWhereFields As New CDBFields

        vAnsiJoin.Add("exam_centre_units ecu", "dl.exam_centre_unit_id", "ecu.exam_centre_unit_id")
        vAnsiJoin.Add("exam_units eu ", "ecu.exam_unit_id", "eu.exam_unit_id")

        vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
        vWhereFields.Add("dl.exam_centre_unit_id", "0", CDBField.FieldWhereOperators.fwoNotEqual)

        Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "eu.exam_unit_description", vAnsiJoin)

        pDataTable.FillFromSQL(mvEnv, vSqlQuery, vCols)
      End If
    End Sub

    Private Sub GetDocumentLinksForFundraisingRequest(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      'BR19023 From Document Links
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDocumentLogLinks) Then
        Dim vEntityDesc As String = GetLinkEntityTypeDescription("F")
        Dim vFields As String = "'R' as link_type,dl.communications_log_number,frr.fundraising_request_number,frr.request_description"
        vFields &= ", 'F' AS entity_type, '" & vEntityDesc & "' AS entity_type_desc"
        Dim vCols As String = "link_type,fundraising_request_number,,request_description,,,FUNDRAISING_REQUEST,LINK_TYPE,entity_type,entity_type_desc"
        Dim vAnsiJoin As New AnsiJoins
        Dim vWhereFields As New CDBFields

        vAnsiJoin.Add("fundraising_requests frr", "dl.fundraising_request_number", "frr.fundraising_request_number")

        vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
        vWhereFields.Add("dl.fundraising_request_number", "0", CDBField.FieldWhereOperators.fwoNotEqual)

        Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "frr.request_description", vAnsiJoin)

        pDataTable.FillFromSQL(mvEnv, vSqlQuery, vCols)
      End If
    End Sub

    Private Sub GetDocumentLinksForWorkstreams(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vFields As String = "'R' as link_type,dl.communications_log_number,ws.workstream_id,ws.workstream_group,ws.workstream_desc"
      vFields &= ",'W' AS entity_type, wg.workstream_group_desc"
      Dim vCols As String = "link_type,workstream_id,workstream_group,workstream_desc,,,WORKSTREAM,LINK_TYPE,entity_type,workstream_group_desc"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("workstreams ws", "dl.workstream_id", "ws.workstream_id")
      vAnsiJoin.Add("workstream_groups wg", "ws.workstream_group", "wg.workstream_group")

      vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
      Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "ws.workstream_desc", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSqlQuery, vCols)
    End Sub


    Private Sub GetFundraisingRequestDocuments(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      'BR19023 From Fundraising Request Links
      Dim vFields As String = "clg.dated,clg.communications_log_number,clg.package,con.label_name,clg.contact_number,dtp.document_type_desc,clg.created_by,dps.department_desc,clg.our_reference,clg.direction,clg.their_reference,clg.document_type,clg.document_class,dcd.document_class_desc,clg.standard_document,clg.source,clg.recipient,clg.forwarded,clg.archiver,clg.completed,cls.topic,tps.topic_desc,cls.sub_topic,sts.sub_topic_desc,dcc.creator_header,dcc.department_header,dcc.public_header,clg.department,std.standard_document_desc,clg.precis, clg.subject, clg.call_duration, clg.total_duration, clg.selection_set"
      Dim vAnsiJoin As New AnsiJoins
      Dim vWhereFields As New CDBFields

      vAnsiJoin.Add("fundraising_requests frr", "dl.fundraising_request_number", "frr.fundraising_request_number")

      vWhereFields.Add("dl.communications_log_number", mvParameters("DocumentNumber").LongValue)
      vWhereFields.Add("dl.fundraising_request_number", "0", CDBField.FieldWhereOperators.fwoNotEqual)

      Dim vSqlQuery As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dl", vWhereFields, "frr.request_description", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSqlQuery, "link_type,fundraising_request_number,,request_description,,,FUNDRAISING_REQUEST,LINK_TYPE")
    End Sub

    Private Sub GetDuplicateContacts(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetDuplicateContacts")
      Dim vWhereFields As New CDBFields

      Dim vAttrs As String = "title,initials,surname,forenames,preferred_forename,honorifics,salutation,label_name,sex,date_of_birth,department,contact_type,c.status,a.address_number,address,house_name,town,county,postcode,country,ownership_access_level,c.ownership_group"
      vAttrs = CheckContactNameAttrs(vAttrs)
      If mvEnv.OwnershipMethod <> CDBEnvironment.OwnershipMethods.omOwnershipGroups Then vAttrs = Replace(vAttrs, ",ownership_access_level,c.ownership_group", ",,")
      Dim vSQL As String = "SELECT c.contact_number," & RemoveBlankItems(vAttrs)

      If Len(mvEnv.GetConfig("uniserv_mail")) > 0 Then
        'BR15650
        'get contact records from local db
        Dim vSQLNew As String = vSQL
        vSQLNew = vSQLNew & " FROM addresses a, contact_addresses ca, contacts c" & mvEnv.User.OwnershipSelect("c", True, , True) & " WHERE "
        If mvParameters.Exists("Postcode") Then vWhereFields.Add("Postcode", mvParameters("Postcode").Value)
        'Replacing any line feed characters with wildcard characters within the multi-line Address parameter
        If mvParameters.Exists("Address") Then vWhereFields.Add("Address", mvParameters("Address").Value.Replace(vbLf, "%"), CDBField.FieldWhereOperators.fwoLike)
        If mvParameters.Exists("AddressLine1") Then vWhereFields.Add("address_line1", mvParameters("AddressLine1").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vWhereFields.AddJoin("a.address_number", "ca.address_number")
        vWhereFields.AddJoin("ca.contact_number", "c.contact_number")
        vSQLNew = vSQLNew & mvEnv.Connection.WhereClause(vWhereFields)
        vSQLNew = vSQLNew & mvEnv.User.OwnershipSelect("c", False, , True)
        vSQLNew = vSQLNew & " ORDER BY c.contact_number"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQLNew, "DISTINCT_CONTACT_NUMBER,CONTACT_NAME," & vAttrs)

        Dim vExternalNumbers As String = ""
        Dim vStreetNo As String = mvEnv.UniservInterface.GetStreetNo(mvParameters.OptionalValue("Address", ""))
        Dim vErrorNumber As Integer = 0
        vErrorNumber = mvEnv.UniservInterface.FindContact(mvParameters.OptionalValue("Forenames", ""), mvParameters.OptionalValue("Surname", ""), "", vStreetNo, mvParameters.OptionalValue("Address", mvParameters.OptionalValue("AddressLine1", "")), mvParameters.OptionalValue("Town", ""), mvParameters.OptionalValue("Postcode", ""), mvParameters.OptionalValue("Country", ""), "", vExternalNumbers)
        If vErrorNumber = 0 Then
          If vExternalNumbers.Length > 0 Then
            'remove duplicate contact numbers from external numbers list
            If pDataTable.Rows.Count > 0 Then
              Dim vContactNumbers As New ArrayList(vExternalNumbers.Split(CChar(",")))
              If vContactNumbers.Count > 0 Then
                For cnt1 As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
                  For cnt2 As Integer = 0 To vContactNumbers.Count - 1
                    If pDataTable.Rows(cnt1).Item("ContactNumber") = vContactNumbers(cnt2).ToString Then
                      vContactNumbers.RemoveAt(cnt2)
                      Exit For
                    End If
                  Next
                Next
                vExternalNumbers = String.Join(",", vContactNumbers.ToArray())
              End If
            End If
            If vExternalNumbers.Length > 0 Then
              'append uniserv contact record to local db result
              vWhereFields.Clear()

              vSQLNew = vSQL
              vSQLNew = vSQLNew & " FROM addresses a, contact_addresses ca, contacts c" & mvEnv.User.OwnershipSelect("c", True, , True) & " WHERE "
              vSQLNew &= "c.contact_number IN(" & vExternalNumbers & ") AND "
              vWhereFields.AddJoin("a.address_number", "ca.address_number")
              vWhereFields.AddJoin("ca.contact_number", "c.contact_number")
              vSQLNew = vSQLNew & mvEnv.Connection.WhereClause(vWhereFields)
              vSQLNew = vSQLNew & mvEnv.User.OwnershipSelect("c", False, , True)
              vSQLNew = vSQLNew & " ORDER BY contact_type DESC, surname, c.contact_number, historical"

              pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQLNew, "DISTINCT_CONTACT_NUMBER,CONTACT_NAME," & vAttrs)
            End If
          End If
        End If

      Else
        Dim vMaxRows As Integer
        If mvParameters.HasValue("NumberOfRows") Then
          vMaxRows = mvParameters("NumberOfRows").IntegerValue + 1
          If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseTopN Then vSQL = vSQL.Replace("SELECT ", "SELECT /* SQLServerCSC */ TOP " & vMaxRows & " ")
        End If
        Dim vAllAddresses As Boolean
        If (mvParameters.Exists("Surname") Or mvParameters.Exists("DateOfBirth")) And mvParameters.Exists("PostcodeOnly") = False Then      'Contacts First
          If mvParameters.HasValue("Postcode") Or mvParameters.HasValue("Country") Or mvParameters.HasValue("BuildingNumber") Then
            vAllAddresses = True
            vSQL = vSQL & " FROM contacts c, contact_addresses ca, addresses a" & mvEnv.User.OwnershipSelect("c", True, , True) & " WHERE "
          Else
            vSQL = vSQL & " FROM contacts c, addresses a" & mvEnv.User.OwnershipSelect("c", True, , True) & " WHERE "
          End If
          If mvParameters.Exists("Surname") Then vWhereFields.Add("surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          If mvParameters.Exists("Title") Then vWhereFields.Add("title", mvParameters("Title").Value)
          If mvParameters.Exists("Forenames") Then
            If mvEnv.GetConfigOption("cd_dedup_forenames_wildcards") = True Then      'BR12427: Use wildcard on both ends
              vWhereFields.Add("forenames", "*" & mvParameters("Forenames").Value & "*", CDBField.FieldWhereOperators.fwoLike)
            Else
              vWhereFields.Add("forenames", mvParameters("Forenames").Value & "*", CDBField.FieldWhereOperators.fwoLike)
            End If
          End If
          If mvParameters.Exists("Initials") Then vWhereFields.Add("initials", mvParameters("Initials").Value & "*", CDBField.FieldWhereOperators.fwoLike)
          If mvParameters.Exists("ContactGroup") Then
            If mvParameters("ContactGroup").Value = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
              vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
            Else
              vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value)
            End If
          End If
          If vAllAddresses Then
            vWhereFields.AddJoin("c.contact_number", "ca.contact_number")
            vWhereFields.AddJoin("ca.address_number", "a.address_number")
          Else
            vWhereFields.AddJoin("c.address_number", "a.address_number")
          End If
          If mvParameters.Exists("Postcode") Then vWhereFields.Add("postcode", mvParameters("Postcode").Value)
          If mvParameters.Exists("BuildingNumber") And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vWhereFields.Add("building_number", mvParameters("BuildingNumber").Value)
          If mvParameters.Exists("Country") Then vWhereFields.Add("country", mvParameters("Country").Value)
          If mvParameters.Exists("DateOfBirth") Then vWhereFields.Add("date_of_birth", CDBField.FieldTypes.cftDate, mvParameters("DateOfBirth").Value)
          vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields)
        Else
          vSQL = vSQL & " FROM addresses a, contact_addresses ca, contacts c" & mvEnv.User.OwnershipSelect("c", True, , True) & " WHERE "
          If mvParameters.Exists("Country") Then vWhereFields.Add("country", mvParameters("Country").Value)
          If mvParameters.Exists("Postcode") Then vWhereFields.Add("postcode", mvParameters("Postcode").Value)
          If mvParameters.Exists("BuildingNumber") And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vWhereFields.Add("building_number", mvParameters("BuildingNumber").Value)
          vWhereFields.AddJoin("a.address_number", "ca.address_number")
          vWhereFields.AddJoin("ca.contact_number", "c.contact_number")
          vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields)
        End If
        vSQL = vSQL & mvEnv.User.OwnershipSelect("c", False, , True)
        If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseRownum And mvParameters.HasValue("NumberOfRows") Then vSQL = vSQL & " AND rownum < " & vMaxRows + 1
        vSQL = vSQL & " ORDER BY contact_type DESC, surname, c.contact_number"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "DISTINCT_CONTACT_NUMBER,CONTACT_NAME," & vAttrs)
      End If
    End Sub
    Private Sub GetDuplicateOrganisations(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "name,abbreviation,o.department,o.status,a.address_number,address,house_name,town,county,postcode,country,ogu.ownership_access_level,o.ownership_group"
      If mvEnv.OwnershipMethod <> CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vAttrs = vAttrs.Replace(",ogu.ownership_access_level,o.ownership_group", ",,")
      End If
      Dim vCols As String = vAttrs
      vAttrs = "o.organisation_number," & RemoveBlankItems(vAttrs)

      If Len(mvEnv.GetConfig("uniserv_mail")) > 0 Then
        Dim vExternalNumbers As String = ""
        Dim vStreetNo As String = mvEnv.UniservInterface.GetStreetNo(mvParameters.OptionalValue("Address", ""))
        Dim vErrorNumber As Integer = mvEnv.UniservInterface.FindContact(mvParameters.OptionalValue("Forenames", ""), mvParameters.OptionalValue("Surname", ""), mvParameters.OptionalValue("Name", ""), vStreetNo, mvParameters.OptionalValue("Address", ""), mvParameters.OptionalValue("Town", ""), mvParameters.OptionalValue("Postcode", ""), mvParameters.OptionalValue("Country", ""), "", vExternalNumbers)
        If vErrorNumber = 0 Then
          If vExternalNumbers.Length > 0 Then
            Dim vSQL As String = "SELECT " & vAttrs
            vSQL = vSQL & ",oa.historical FROM organisations o, organisation_addresses oa, addresses a"
            vSQL &= mvEnv.User.OwnershipSelect("o", True, , True)
            vSQL &= " WHERE o.organisation_number IN(" & vExternalNumbers & ")"
            vSQL = vSQL & " AND o.organisation_number = oa.organisation_number AND oa.address_number = a.address_number"
            vSQL = vSQL & mvEnv.User.OwnershipSelect("o", False, , True)
            vSQL = vSQL & " ORDER BY name, o.organisation_number, historical"
            pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "DISTINCT_ORGANISATION_NUMBER," & vCols)
          End If
        End If
      Else
        Dim vMaxRows As Integer
        If mvParameters.HasValue("NumberOfRows") Then
          vMaxRows = mvParameters("NumberOfRows").IntegerValue + 1
        End If

        Dim vAnsiJoins As New AnsiJoins()
        Dim vWhereFields As New CDBFields()
        Dim vTableName As String = "organisations o"
        Dim vOrderBy As String = "name, o.organisation_number"

        If mvParameters.Exists("Name") And mvParameters.Exists("PostcodeOnly") = False Then      'Contacts First
          If mvParameters.HasValue("Postcode") Then
            Dim vOrgName As String = mvParameters("Name").Value
            Dim vPostcode As String = mvParameters("Postcode").Value
            vOrderBy = "oa.sequence_number, " & vOrderBy

            'Name and Postcode
            Dim vNamePostcodeAttrs As String = "oa1.organisation_number, oa1.address_number, '1' AS sequence_number"
            Dim vNamePostcodeAJ As New AnsiJoins()
            vNamePostcodeAJ.Add("organisation_addresses oa1", "o1.organisation_number", "oa1.organisation_number")
            vNamePostcodeAJ.Add("addresses a1", "oa1.address_number", "a1.address_number")
            Dim vNamePostcodeWF As New CDBFields(New CDBField("name", vOrgName, CDBField.FieldWhereOperators.fwoLikeOrEqual))
            vNamePostcodeWF.Add("a1.postcode", vPostcode)
            Dim vNamePostcodeSQL As New SQLStatement(mvEnv.Connection, vNamePostcodeAttrs, "organisations o1", vNamePostcodeWF, String.Empty, vNamePostcodeAJ)

            'Name only
            Dim vNameAttrs As String = "o2.organisation_number, a2.address_number, '2' AS sequence_number"
            Dim vNameAJ As New AnsiJoins({New AnsiJoin("addresses a2", "o2.address_number", "a2.address_number")})
            Dim vNameWF As New CDBFields(New CDBField("o2.name", vOrgName, CDBField.FieldWhereOperators.fwoLike))
            Dim vNameSQL As New SQLStatement(mvEnv.Connection, vNameAttrs, "organisations o2", vNameWF, String.Empty, vNameAJ)

            'Postcode only
            Dim vPostcodeAttrs As String = "oa3.organisation_number, oa3.address_number, '3' AS sequence_number"
            Dim vPostcodeAJ As New AnsiJoins({New AnsiJoin("organisation_addresses oa3", "o3.organisation_number", "oa3.organisation_number")})
            vPostcodeAJ.Add("addresses a3", "oa3.address_number", "a3.address_number")
            Dim vPostcodeLike As String = vPostcode
            If vPostcodeLike.Contains("*") = False OrElse vPostcodeLike.Contains("%") = False Then
              vPostcodeLike &= "%"
            End If
            Dim vPostcodeWF As New CDBFields(New CDBField("a3.postcode", vPostcodeLike, CDBField.FieldWhereOperators.fwoLike))
            vPostcodeWF.Add("o3.name", vOrgName, CDBField.FieldWhereOperators.fwoNotLike)
            Dim vPostcodeSQL As New SQLStatement(mvEnv.Connection, vPostcodeAttrs, "organisations o3", vPostcodeWF, String.Empty, vPostcodeAJ)

            'Wrap the SQL
            Dim vInnerAttrs As String = mvEnv.Connection.DBIsNull("ta1.organisation_number", mvEnv.Connection.DBIsNull("ta2.organisation_number", "ta3.organisation_number")) & " AS organisation_number"
            vInnerAttrs &= ", " & mvEnv.Connection.DBIsNull("ta1.address_number", mvEnv.Connection.DBIsNull("ta2.address_number", "ta3.address_number")) & " AS address_number"
            vInnerAttrs &= ", " & mvEnv.Connection.DBIsNull("ta1.sequence_number", mvEnv.Connection.DBIsNull("ta2.sequence_number", "ta3.sequence_number")) & " AS sequence_number"

            Dim vInnerAJ As New AnsiJoins()
            vInnerAJ.AddLeftOuterJoin("(" & vNamePostcodeSQL.SQL & ") ta1", "org.organisation_number", "ta1.organisation_number")
            vInnerAJ.AddLeftOuterJoin("(" & vNameSQL.SQL & ") ta2", "org.organisation_number", "ta2.organisation_number")
            vInnerAJ.AddLeftOuterJoin("(" & vPostcodeSQL.SQL & ") ta3", "org.organisation_number", "ta3.organisation_number")

            Dim vInnerWF As New CDBFields(New CDBField("ta1.organisation_number", String.Empty, CDBField.FieldWhereOperators.fwoNotEqual))
            vInnerWF.Add("ta2.organisation_number", String.Empty, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual)
            vInnerWF.Add("ta3.organisation_number", String.Empty, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual)

            Dim vInnerSQL As New SQLStatement(mvEnv.Connection, vInnerAttrs, "organisations org", vInnerWF, String.Empty, vInnerAJ)

            'Main SQL
            vAnsiJoins.Add("(" & vInnerSQL.SQL & ") oa", "o.organisation_number", "oa.organisation_number")
            vAnsiJoins.Add("addresses a", "oa.address_number", "a.address_number")

            'vWhereFields.Add("name", vOrgName, CDBField.FieldWhereOperators.fwoLikeOrEqual)

          ElseIf (mvParameters.HasValue("Country") Or mvParameters.HasValue("BuildingNumber")) Then
            'vAllAddresses = True
            vAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number")
            vAnsiJoins.Add("addresses a", "oa.address_number", "a.address_number")
            vWhereFields.Add("name", mvParameters("Name").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          Else
            vAnsiJoins.Add("addresses a", "o.address_number", "a.address_number")
            vWhereFields.Add("name", mvParameters("Name").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          End If

          If mvParameters.Exists("OrganisationGroup") Then
            If mvParameters("OrganisationGroup").Value = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtOrganisation).EntityGroupCode Then
              vWhereFields.Add("o.organisation_group", mvParameters("OrganisationGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
            Else
              vWhereFields.Add("o.organisation_group", mvParameters("OrganisationGroup").Value)
            End If
          End If
        Else
          vTableName = "addresses a"
          vAnsiJoins.Add("organisation_addresses oa", "a.address_number", "oa.address_number")
          vAnsiJoins.Add("organisations o", "oa.organisation_number", "o.organisation_number")

          If mvParameters.Exists("Postcode") Then vWhereFields.Add("a.postcode", mvParameters("Postcode").Value)
        End If

        If mvParameters.Exists("BuildingNumber") Then vWhereFields.Add("a.building_number", mvParameters("BuildingNumber").Value)
        If mvParameters.Exists("Country") Then vWhereFields.Add("a.country", mvParameters("Country").Value)

        mvEnv.User.OwnershipSelect("o", String.Empty, True, vAnsiJoins, vWhereFields)

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTableName, vWhereFields, vOrderBy, vAnsiJoins)
        vSQLStatement.MaxRows = vMaxRows

        pDataTable.FillFromSQL(mvEnv, vSQLStatement, "DISTINCT_ORGANISATION_NUMBER," & vCols)
      End If
    End Sub
    Private Sub GetEMailContacts(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vSQL As String = "Select "
      Dim vMaxRows As Integer
      Dim vAllAddresses As Boolean
      Dim vPreferred As Boolean = mvParameters.OptionalValue("PreferredMethod", "N") = "Y"

      If mvParameters.HasValue("NumberOfRows") Then
        vMaxRows = mvParameters("NumberOfRows").IntegerValue + 1
        If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseTopN Then vSQL = "Select TOP " & vMaxRows & " "
      End If
      If mvParameters.Exists("ViewAllRows") Then vAllAddresses = mvParameters("ViewAllRows").Bool
      If vPreferred OrElse mvParameters.Exists("Corporate") Then vAllAddresses = True

      Dim vAttrs As String = "title,initials,surname,forenames,preferred_forename,honorifics,salutation,label_name,sex,date_of_birth,c.department,contact_type,c.status,a.address_number,address,house_name,town,county,postcode,country,number,d.device,device_desc,position,name"
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then vAttrs = vAttrs & ",c.ownership_group"
      vAttrs = CheckContactNameAttrs(vAttrs)
      vSQL = vSQL & "c.contact_number," & vAttrs
      Dim vAlias As String
      Dim vSelectionSetTable As String = "selected_contacts"
      If mvParameters.HasValue("SelectionSetNumber") Then
        If mvParameters.HasValue("FromTemporaryTable") Then vSelectionSetTable = "smcam_smapp_" & mvParameters("FromTemporaryTable").Value
        vSQL = vSQL & " FROM " & vSelectionSetTable & " sc INNER JOIN contacts c On sc.contact_number = c.contact_number"
        vSQL = vSQL & " INNER JOIN addresses a On sc.address_number = a.address_number"
        vAlias = "sc."
      Else
        vSQL = vSQL & " FROM contacts c INNER JOIN addresses a On c.address_number = a.address_number"
        vAlias = "c."
      End If
      If vAllAddresses Then
        vSQL = vSQL & " INNER JOIN communications co On c.contact_number = co.contact_number"
      Else
        vSQL = vSQL & " INNER JOIN communications co On c.contact_number = co.contact_number And " & vAlias & "address_number = co.address_number"
      End If
      vSQL = vSQL & " INNER JOIN devices d On co.device = d.device"
      If mvParameters.Exists("Corporate") Then
        vSQL = vSQL & " LEFT OUTER JOIN contact_positions cp On " & vAlias & "contact_number = cp.contact_number"
      Else
        vSQL = vSQL & " LEFT OUTER JOIN contact_positions cp On " & vAlias & "contact_number = cp.contact_number And " & vAlias & "address_number = cp.address_number"
      End If
      vSQL = vSQL & " LEFT OUTER JOIN organisations o On cp.organisation_number = o.organisation_number WHERE "
      If mvParameters.Exists("SelectionSetNumber") Then vWhereFields.Add("sc.selection_set", mvParameters("SelectionSetNumber").LongValue)
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("c.contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.Exists("Surname") Then vWhereFields.Add("surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Title") Then vWhereFields.Add("title", mvParameters("Title").Value)
      If mvParameters.Exists("Forenames") Then vWhereFields.Add("forenames", mvParameters("Forenames").Value & "*", CDBField.FieldWhereOperators.fwoLike)
      If mvParameters.Exists("Initials") Then vWhereFields.Add("initials", mvParameters("Initials").Value & "*", CDBField.FieldWhereOperators.fwoLike)
      If mvParameters.Exists("ContactGroup") Then
        If mvParameters("ContactGroup").Value = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
          vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
        Else
          vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value)
        End If
      End If
      If mvParameters.Exists("OrganisationNumber") Then vWhereFields.Add("o.organisation_number", mvParameters("OrganisationNumber").Value)
      If mvParameters.Exists("EMailAddress") Then vWhereFields.Add(mvEnv.Connection.DBSpecialCol("co", "number"), mvParameters("EMailAddress").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.Add("email", "Y")
      'vWhereFields.Add(("current",  "Y", CDBField.FieldWhereOperators.fwoNullOrEqual).SpecialColumn = True)
      vWhereFields.Add("cp.mail", "Y", CDBField.FieldWhereOperators.fwoNullOrEqual)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("co.is_active", "Y")
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields)
      If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseRownum And vMaxRows > 1 Then vSQL = vSQL & " And rownum < " & vMaxRows + 1
      vSQL = vSQL & " ORDER BY surname, c.contact_number, preferred_method DESC, device_default DESC"
      vSQL = Replace$(vSQL, ",number,", "," & mvEnv.Connection.DBSpecialCol("", "number") & ",")
      Dim vContactField As String = "contact_number"
      If vPreferred Then vContactField = "DISTINCT_CONTACT_NUMBER"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vContactField & ",CONTACT_NAME," & vAttrs)

      'If using a selection set there may have been organisations in the selection set in which case we would not have retrieved any records
      If mvParameters.HasValue("SelectionSetNumber") Then
        Dim vSQLAttrs As String = "c.contact_number As contact_number, " & vAttrs.Replace(",house_name,", ",'' AS house_name,").Replace(",address,", ",'' AS address,") & ",preferred_method,device_default"
        vSQLAttrs = Replace$(vSQLAttrs, ",number,", "," & mvEnv.Connection.DBSpecialCol("", "number") & ",")
        'Get the email addresses of any organisation in the selection set 
        Dim vOrgAnsiJoins As New AnsiJoins
        vOrgAnsiJoins.Add("organisations o", "sc.contact_number", "o.organisation_number")
        vOrgAnsiJoins.Add("contacts c", "sc.contact_number", "c.contact_number")
        vOrgAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number")
        vOrgAnsiJoins.Add("addresses a", "oa.address_number", "a.address_number")
        vOrgAnsiJoins.Add("communications co", "oa.address_number", "co.address_number")
        vOrgAnsiJoins.Add("devices d", "co.device", "d.device")
        Dim vOrgWhere As New CDBFields
        vOrgWhere.Add("sc.selection_set", mvParameters("SelectionSetNumber").IntegerValue)
        vOrgWhere.Add("co.contact_number")
        vOrgWhere.Add("email", "Y")
        vOrgWhere.Add("co.is_active", "Y")
        Dim vOrgSQL As New SQLStatement(mvEnv.Connection, vSQLAttrs.Replace(",position,", ",'' AS position,"), vSelectionSetTable & " sc", vOrgWhere, "name, contact_type desc, contact_number, preferred_method DESC, device_default DESC", vOrgAnsiJoins)

        'Get the email addresses of the default contact for any organisation in the selection set 
        Dim vConAnsiJoins As New AnsiJoins
        vConAnsiJoins.Add("organisations o", "sc.contact_number", "o.organisation_number")
        vConAnsiJoins.Add("contacts c", "o.contact_number", "c.contact_number")
        vConAnsiJoins.Add("contact_positions cp", "c.contact_number", "cp.contact_number", "o.organisation_number", "cp.organisation_number")
        vConAnsiJoins.Add("addresses a", "cp.address_number", "a.address_number")
        vConAnsiJoins.Add("communications co", "c.contact_number", "co.contact_number")
        vConAnsiJoins.Add("devices d", "co.device", "d.device")
        Dim vConWhere As New CDBFields
        vConWhere.Add("sc.selection_set", mvParameters("SelectionSetNumber").IntegerValue)
        vConWhere.Add("email", "Y")
        vConWhere.Add("co.is_active", "Y")
        vConWhere.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
        vConWhere.Add("cp.mail", "Y", CDBField.FieldWhereOperators.fwoNullOrEqual)
        vConWhere.Add("current", "Y", CDBField.FieldWhereOperators.fwoNullOrEqual).SpecialColumn = True
        Dim vConSQL As New SQLStatement(mvEnv.Connection, vSQLAttrs, vSelectionSetTable & " sc", vConWhere, "", vConAnsiJoins)
        vConSQL.UseAnsiSQL = True
        vOrgSQL.AddUnion(vConSQL)
        vOrgSQL.UseAnsiSQL = True
        pDataTable.FillFromSQL(mvEnv, vOrgSQL, vContactField & ",CONTACT_NAME," & vAttrs)
      End If
      pDataTable.SuppressData()
      pDataTable.RemoveRowsWithBlankColumn("EMailAddress")
    End Sub
    Private Sub GetEMailOrganisations(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vWhereFields As New CDBFields
      Dim vSQL As String = "SELECT "
      Dim vMaxRows As Integer
      Dim vPreferred As Boolean = mvParameters.OptionalValue("PreferredMethod", "N") = "Y"

      If mvParameters.HasValue("NumberOfRows") Then
        vMaxRows = mvParameters("NumberOfRows").LongValue + 1
        If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseTopN Then vSQL = "SELECT TOP " & vMaxRows & " "
      End If
      'BR11685/11700 - Using GetRecordsetFields as is the required way!
      'SDT Restored to previous list
      Dim vAttrs As String = "title,initials,surname,forenames,preferred_forename,honorifics,salutation,label_name,sex,date_of_birth,c.department,contact_type,c.status,a.address_number,address,house_name,town,county,postcode,country,number,d.device,device_desc,position,name"
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then vAttrs = vAttrs & ",c.ownership_group"
      vAttrs = CheckContactNameAttrs(vAttrs)
      vSQL = vSQL & "c.contact_number," & vAttrs
      vSQL = vSQL & " FROM organisations o INNER JOIN organisation_addresses oa ON o.organisation_number = oa.organisation_number"
      vSQL = vSQL & " INNER JOIN contact_positions cp ON oa.organisation_number = cp.organisation_number AND oa.address_number = cp.address_number"
      vSQL = vSQL & " INNER JOIN addresses a ON cp.address_number = a.address_number"
      vSQL = vSQL & " INNER JOIN contacts c ON cp.contact_number = c.contact_number"
      vSQL = vSQL & " INNER JOIN communications co ON %1"
      vSQL = vSQL & " INNER JOIN devices d ON co.device = d.device WHERE "
      If mvParameters.Exists("OrganisationNumber") Then vWhereFields.Add("o.organisation_number", mvParameters("OrganisationNumber").LongValue)
      If mvParameters.Exists("Name") Then vWhereFields.Add("name", mvParameters("Name").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Abbreviation") Then vWhereFields.Add("abbreviation", mvParameters("Abbreviation").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Title") Then vWhereFields.Add("title", mvParameters("Title").Value)
      If mvParameters.Exists("Forenames") Then vWhereFields.Add("forenames", mvParameters("Forenames").Value & "*", CDBField.FieldWhereOperators.fwoLike)
      If mvParameters.Exists("Initials") Then vWhereFields.Add("initials", mvParameters("Initials").Value & "*", CDBField.FieldWhereOperators.fwoLike)
      If mvParameters.Exists("Surname") Then vWhereFields.Add("surname", mvParameters("Surname").Value)
      If mvParameters.Exists("ContactGroup") Then
        If mvParameters("ContactGroup").Value = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode Then
          vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
        Else
          vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value)
        End If
      End If
      vWhereFields.Add("cp.mail", "Y")
      vWhereFields.Add("email", "Y")
      vWhereFields.Add("co.is_active", "Y")
      vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields)
      vSQL = vSQL & " %2 "
      If mvEnv.Connection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseRownum And vMaxRows > 1 Then vSQL = vSQL & " AND rownum < " & vMaxRows + 1
      Dim vBaseSQL As String = vSQL & " ORDER BY contact_type DESC, surname, c.contact_number, preferred_method DESC, device_default DESC"
      vBaseSQL = Replace$(vBaseSQL, ",number,", "," & mvEnv.Connection.DBSpecialCol("", "number") & ",")
      'First get the email addresses for the organisation
      vSQL = Replace$(vBaseSQL, "%1", "oa.address_number = co.address_number")
      vSQL = Replace$(vSQL, "%2", " AND co.contact_number IS NULL AND c.contact_type = 'O'")
      'BR11685/11700 - Ensure that we have the right values in vAttrs
      Dim vContactField As String = "c.contact_number"
      If vPreferred Then vContactField = "DISTINCT_CONTACT_NUMBER"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vContactField & ",CONTACT_NAME," & vAttrs)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("Position") = ""
      Next

      'Now get the email addresses for the contacts at the organisation
      'By default restrict to the ones which have the organisations address
      Dim vAllAddresses As Boolean
      If mvParameters.Exists("ViewAllRows") Then vAllAddresses = mvParameters("ViewAllRows").Bool
      If vPreferred Then vAllAddresses = True
      If vAllAddresses Then
        vSQL = Replace$(vBaseSQL, "%1", "c.contact_number = co.contact_number")
      Else
        vSQL = Replace$(vBaseSQL, "%1", "c.contact_number = co.contact_number AND cp.address_number = co.address_number")
      End If
      vSQL = Replace$(vSQL, "%2", "")
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vContactField & ",CONTACT_NAME," & vAttrs)
      pDataTable.SuppressData()
      pDataTable.RemoveRowsWithBlankColumn("EMailAddress")
    End Sub
    Private Sub GetEventAccommodation(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventAccommodation")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "rbb.block_booking_number,room_type_desc,from_date,to_date,number_of_rooms,nights_available,o.name,a.address,rack_rate,agreed_rate,booked_date,release_date,p.product,p.product_desc,confirmed_date,r.rate,r.rate_desc,rbb.amended_by,rbb.amended_on,o.organisation_number,a.address_number,rbb.notes,rbb.room_type"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vAttrs & "," & vConAttrs
      vSQL = vSQL & " FROM event_room_links erl, room_block_bookings rbb, room_types rt, organisations o, contacts c, addresses a, products p, rates r WHERE "
      vSQL = vSQL & "erl.event_number = " & mvParameters("EventNumber").LongValue & " AND erl.block_booking_number = rbb.block_booking_number"
      If mvParameters.Exists("BlockBookingNumber") Then vSQL = vSQL & " AND rbb.block_booking_number = " & mvParameters("BlockBookingNumber").LongValue
      vSQL = vSQL & " AND rbb.contact_number = c.contact_number AND rbb.organisation_number = o.organisation_number" & " AND rbb.room_type = rt.room_type AND rbb.address_number = a.address_number" & " AND rbb.product = p.product AND rbb.rate = r.rate AND p.product = r.product"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & ",contact_number," & ContactNameItems())
      vSQL = Replace(vSQL, "contacts c,", "")
      vSQL = Replace(vSQL, "rbb.contact_number = c.contact_number", "rbb.contact_number IS NULL")
      vSQL = Replace(vSQL, "," & vConAttrs, "")
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",," & ContactNameItems(True))
      Dim vTableSort(2) As CDBDataTable.SortSpecification
      vTableSort(2).Column = "RoomTypeDesc"
      vTableSort(2).Descending = False
      vTableSort(1).Column = "FromDate"
      vTableSort(1).Descending = False
      vTableSort(0).Column = "BookedOn"
      vTableSort(0).Descending = False
      pDataTable.ReOrderRowsByMultipleColumns(vTableSort)
    End Sub
    Private Sub GetEventAttendees(ByVal pDataTable As CDBDataTable)
      Dim vSelectAttrs As String = "d.booking_number,d.contact_number,d.address_number,t1.label_name AS  delegate_name,t2.label_name AS  booker_name,d.attended,d.position,d.organisation_name,d.candidate_number,eb.booking_status,eb.quantity,d.event_delegate_number"
      vSelectAttrs = vSelectAttrs & ",pledged_amount,donation_total,sponsorship_total,booking_payment_amount,other_payments_total,sequence_number,pc.contact_number AS payer_contact_number,pc.label_name AS payer_label_name"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDelegateSequenceNumber) = False Then vSelectAttrs = vSelectAttrs.Replace("sequence_number", "")
      Dim vAttrs As String = "d.booking_number,d.contact_number,d.address_number,delegate_name,booker_name,d.attended,d.position,d.organisation_name,d.candidate_number,eb.booking_status,eb.quantity,d.event_delegate_number"
      vAttrs = vAttrs & ",pledged_amount,donation_total,sponsorship_total,booking_payment_amount,other_payments_total,sequence_number,payer_contact_number,payer_label_name"

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vTable As String
      Dim vOrderBy As String

      If mvParameters.Exists("SessionNumber") Then
        vTable = "session_bookings sb"
        vAnsiJoins.Add("delegates d", "sb.booking_number", "d.booking_number")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
          vAnsiJoins.Add("delegate_sessions ds", "ds.event_delegate_number", "d.event_delegate_number", "ds.session_number", "sb.session_number")
          'Supply attended from delegate_sessions rather than delegates table.
          vSelectAttrs = vSelectAttrs.Replace("d.attended", "ds.attended")
          If mvParameters.Exists("Attended") Then vWhereFields.Add("ds.attended", mvParameters("Attended").Value)
        End If
        vWhereFields.Add("sb.session_number", mvParameters("SessionNumber").LongValue)
      Else
        vTable = "delegates d"
        vWhereFields.Add("d.event_number", mvParameters("EventNumber").LongValue)
      End If
      vAnsiJoins.Add("event_bookings eb", "d.booking_number", "eb.booking_number")
      vAnsiJoins.Add("contacts t1", "d.contact_number", "t1.contact_number")
      vAnsiJoins.Add("contacts t2", "eb.contact_number", "t2.contact_number")
      vAnsiJoins.AddLeftOuterJoin("batch_transactions bt", "eb.batch_number", "bt.batch_number", "eb.transaction_number", "bt.transaction_number")
      vAnsiJoins.AddLeftOuterJoin("contacts pc", "bt.contact_number", "pc.contact_number")

      If mvParameters.Exists("EventDelegateNumber") Then vWhereFields.Add("d.event_delegate_number", mvParameters("EventDelegateNumber").LongValue)

      vWhereFields.Add("eb.booking_status", CDBField.FieldTypes.cftCharacter, "'" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsInterested) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsAwaitingAcceptance) & "'", CDBField.FieldWhereOperators.fwoNotIn)
      vOrderBy = "t2.surname,t2.contact_number,d.booking_number"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vSelectAttrs), vTable, vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs)

      Dim vValue As String = ""
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Attended")
        vRow.Item("BookingStatus") = EventBooking.GetBookingStatusDescription(vRow.Item("BookingStatus"))
        If vValue = vRow.Item("BookerName") Then
          vRow.Item("BookerName") = ""
        Else
          vValue = vRow.Item("BookerName")
        End If
      Next
    End Sub
    Private Sub GetEventCurrentAttendees(ByVal pDataTable As CDBDataTable)
      Dim vEvent As New CDBEvent(mvEnv)
      vEvent.Init()
      '(1) Event booking details
      Dim vSelectAttrs As String = "d.booking_number,d.contact_number,d.address_number,t1.label_name AS delegate_name,t2.label_name AS booker_name,d.attended,d.position,d.organisation_name,d.candidate_number,eb.booking_status,eb.quantity,d.event_delegate_number,s.session_number,s.session_type"
      Dim vAttrs As String = "d.booking_number,d.contact_number,d.address_number,delegate_name,booker_name,d.attended,d.position,d.organisation_name,d.candidate_number,eb.booking_status,eb.quantity,d.event_delegate_number,s.session_number,s.session_type"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        vAttrs = vAttrs & ",pledged_amount,donation_total,sponsorship_total,booking_payment_amount,other_payments_total"
        vSelectAttrs = vSelectAttrs & ",pledged_amount,donation_total,sponsorship_total,booking_payment_amount,other_payments_total"
      Else
        vAttrs = vAttrs & ",,,,,"
      End If
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vSelectAttrs)
      vSQL = vSQL & " FROM delegates d,event_bookings eb,contacts t1,contacts t2,sessions s WHERE d.event_number=" & mvParameters("EventNumber").LongValue
      If mvParameters.Exists("EventDelegateNumber") Then vSQL = vSQL & " AND d.event_delegate_number = " & mvParameters("EventDelegateNumber").Value
      vSQL = vSQL & " AND d.contact_number=t1.contact_number AND d.booking_number = eb.booking_number"
      vSQL = vSQL & " AND eb.booking_status IN (" & vEvent.CurrentAttendeeBookingStatuses.InList & ")"
      vSQL = vSQL & " AND eb.contact_number = t2.contact_number"
      vSQL = vSQL & " AND s.event_number = eb.event_number AND s.session_type = '0'"
      vSQL = vSQL & " ORDER BY t1.surname, t1.forenames"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, RemoveBlankItems(vSQL), vAttrs)

      If mvUsage = DataSelectionUsages.dsuSmartClient Then
        '(2) Second pass: Add any session booking details if relevant
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then
          vSQL = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vSelectAttrs)
          vSQL = Replace(vSQL, "d.attended", "ds.attended")
          vSQL = vSQL & " FROM session_bookings sb, delegates d, delegate_sessions ds, event_bookings eb, contacts t1, contacts t2, sessions s"
          vSQL = vSQL & " WHERE sb.event_number = " & mvParameters("EventNumber").LongValue
          If mvParameters.Exists("SessionNumber") Then vSQL = vSQL & " AND sb.session_number = " & mvParameters("SessionNumber").LongValue
          vSQL = vSQL & " AND sb.booking_number = d.booking_number"
          If mvParameters.Exists("EventDelegateNumber") Then vSQL = vSQL & " AND d.event_delegate_number = " & mvParameters("EventDelegateNumber").Value
          vSQL = vSQL & " AND ds.event_delegate_number = d.event_delegate_number AND ds.session_number = sb.session_number "
          If mvParameters.Exists("Attended") Then
            vSQL = vSQL & " AND ds.attended = '" & mvParameters("Attended").Value & "'"
          End If
          vSQL = vSQL & " AND d.contact_number=t1.contact_number AND d.booking_number = eb.booking_number"
          vSQL = vSQL & " AND eb.booking_status IN (" & vEvent.CurrentAttendeeBookingStatuses.InList & ")"
          vSQL = vSQL & " AND eb.contact_number = t2.contact_number"
          vSQL = vSQL & " AND s.session_number = sb.session_number"
          vSQL = vSQL & " AND s.session_type <> '0' "
          vSQL = vSQL & " ORDER BY t1.surname, t1.forenames"
          pDataTable.FillFromSQLDONOTUSE(mvEnv, RemoveBlankItems(vSQL), vAttrs)
        End If
      End If

      Dim vValue As String = ""
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Attended")
        vRow.Item("BookingStatus") = EventBooking.GetBookingStatusDescription(vRow.Item("BookingStatus"))
        If vValue = vRow.Item("BookerName") Then
          vRow.Item("BookerName") = ""
        Else
          vValue = vRow.Item("BookerName")
        End If
      Next vRow
    End Sub
    Private Sub GetEventAuthoriseExpenses(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "ep.session_number,ep.contact_number,c.label_name,ep.expenses,e.event_desc,ep.authorised_on,ep.authorised_by,ep.event_personnel_number,ep.address_number"
      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("sessions s", "ep.session_number", "s.session_number")
        .Add("events e", "s.event_number", "e.event_number")
        .Add("contacts c", "ep.contact_number", "c.contact_number")
      End With

      Dim vWhereFields As New CDBFields(New CDBField("ep.authorised_on", CDBField.FieldTypes.cftCharacter, ""))
      vWhereFields.Add("ep.expenses", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_personnel ep", vWhereFields, "c.surname, c.forenames, e.event_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetEventBookingDelegates(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "d.address_number,position,organisation_name,event_delegate_number,sequence_number"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDelegateSequenceNumber) = False Then vAttrs = vAttrs.Replace("sequence_number", "")
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromIntegerParameter(vWhereFields, "BookingNumber", "d.booking_number")
      AddWhereFieldFromIntegerParameter(vWhereFields, "EventDelegateNumber", "event_delegate_number")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contacts c", "d.contact_number", "c.contact_number")
      Dim vOrderBy As String = "sequence_number, c.surname"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDelegateSequenceNumber) = False Then vOrderBy = "c.surname"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs) & "," & vConAttrs, "delegates d", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "title,forenames,surname,contact_number,CONTACT_NAME")
    End Sub
    Private Sub GetEventBookingOptions(ByVal pDataTable As CDBDataTable)

      Dim vFields As String = "ebo.option_number,option_desc,pick_sessions,number_of_sessions,deduct_from_event,minimum_bookings,maximum_bookings,ebo.product,product_desc,ebo.rate,rate_desc,ebo.amended_by,ebo.amended_on,os.session_number AS issue_event_resources, booking_count,ebo.long_description,ebo.free_of_charge"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLongDescription) = False Then vFields = vFields.Replace("ebo.long_description", "")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEventMinimumBookings) = False Then vFields = vFields.Replace("minimum_bookings", "")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbFreeOfChangeBookingOption) = False Then vFields = vFields.Replace("ebo.free_of_charge", "")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("products p", "ebo.product", "p.product")
      vAnsiJoins.Add("rates r", "ebo.product", "r.product", "ebo.rate", "r.rate")

      Dim vSubAnsiJoins As New AnsiJoins
      vSubAnsiJoins.Add("sessions s", "os.session_number", "s.session_number")

      Dim vSub1SQL As New SQLStatement(mvEnv.Connection, "option_number, session_number", "option_sessions", New CDBField("session_number", GetBaseSessionNumber(mvParameters("EventNumber").LongValue)))


      'Dim vSub1SQL As New SQLStatement(mvEnv.Connection, "option_number, os.session_number", "option_sessions os", New CDBFields(New CDBField("s.event_number", mvParameters("EventNumber").LongValue)), "", vSubAnsiJoins)
      vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) os", vSub1SQL.SQL), "ebo.option_number", "os.option_number")

      Dim vSub2WhereFields As New CDBFields
      vSub2WhereFields.Add("event_number", mvParameters("EventNumber").IntegerValue)
      vSub2WhereFields.Add("booking_status", CDBField.FieldTypes.cftCharacter, "'F', 'X', 'B', 'Y', 'S', 'R', 'V', 'D'", CDBField.FieldWhereOperators.fwoIn)
      Dim vSub2SQL As New SQLStatement(mvEnv.Connection, "option_number, count(*) AS booking_count", "event_bookings", vSub2WhereFields)
      vSub2SQL.GroupBy = "option_number"
      vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) eb", vSub2SQL.SQL), "ebo.option_number", "eb.option_number")

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("event_number", mvParameters("EventNumber").IntegerValue)
      AddWhereFieldFromIntegerParameter(vWhereFields, "OptionNumber", "ebo.option_number")
      If mvParameters.Exists("NoPickSessions") Then vWhereFields.Add("ebo.pick_sessions", "Y", CDBField.FieldWhereOperators.fwoNotEqual)
      AddWhereFieldFromParameter(vWhereFields, "PickSessions", "ebo.pick_sessions")
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "event_booking_options ebo", vWhereFields, "option_desc", vAnsiJoins)
      vSQL.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields.Replace("session_number AS issue_event_resources", "issue_event_resources"))
      For Each vRow As CDBDataRow In pDataTable.Rows
        If Len(vRow.Item("IssueEventResources")) > 0 Then vRow.Item("IssueEventResources") = DataSelectionText.String15904 'Yes
      Next
    End Sub
    Private Sub GetEventBookingOptionSessions(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "os.session_number,session_desc,s.subject,subject_desc,s.skill_level,skill_level_desc,start_date,end_date,start_time,end_time,allocation,s.location,s.long_description,maximum_attendees - number_of_attendees AS places_available,os.amended_by,os.amended_on"
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("sessions s", "os.session_number", " s.session_number")
      vAnsiJoins.Add("subjects su", "s.subject", "su.subject")
      vAnsiJoins.Add("skill_levels sl", "s.skill_level", "sl.skill_level")
      vWhereFields.Add("option_number", CDBField.FieldTypes.cftInteger, mvParameters.ParameterExists("OptionNumber").IntegerValue)
      vWhereFields.Add("s.session_type", "0", CDBField.FieldWhereOperators.fwoNotEqual)
      If mvParameters.Exists("SessionNumber") Then
        vWhereFields.Add("os.session_number", CDBField.FieldTypes.cftInteger, mvParameters("SessionNumber").IntegerValue)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "option_sessions os", vWhereFields, "start_date, start_time", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetEventBookings(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName("c")
      Dim vInvAttr As String = ",i.invoice_pay_status,ips.invoice_pay_status_desc"
      Dim vAttrs As String = "eb.booking_date,eb.booking_status,eb.quantity,option_desc,eb.address_number,eb.rate,eb.sales_contact_number,eb.notes,eb.cancelled_on,eb.cancelled_by,eb.cancellation_reason,eb.cancellation_source,eb.option_number,ebo.product,organisation_number,eb.batch_number,eb.transaction_number,eb.line_number"
      vAttrs &= ",eb.adult_quantity,eb.child_quantity,{0},{1},i.invoice_number,reprint_count,batch_type,{2}"

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventAdultChildQuantity) = False Then vAttrs = vAttrs.Replace("eb.adult_quantity", "").Replace("eb.child_quantity", "")
      Dim vAmountString As String = mvEnv.Connection.DBIsNull("bta.amount", "0")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then vAmountString = "(" & mvEnv.Connection.DBIsNull("bta.amount", "0") & " + " & mvEnv.Connection.DBIsNull("bta2.amount", "0") & " + " & mvEnv.Connection.DBIsNull("ebtfac.amount", "0") & " + " & mvEnv.Connection.DBIsNull("ebtfad.amount", "0") & ") AS amount"

      Dim vSelectAttrs As String = vAttrs
      vAttrs &= ",pc.contact_number AS payer_contact_number,pc.label_name AS payer_label_name"
      vSelectAttrs &= ",payer_contact_number,payer_label_name"

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventPricingMatrix) Then
        vSelectAttrs = String.Format(vSelectAttrs, "start_time", "end_time", "bta.amount")
        vAttrs = String.Format(vAttrs, mvEnv.Connection.DBIsNull("eb.start_time", "s.start_time") & " AS start_time", mvEnv.Connection.DBIsNull("eb.end_time", "s.end_time") & " AS end_time", vAmountString)
      Else
        vAttrs = String.Format(vAttrs, "", "", vAmountString)
        vSelectAttrs = String.Format(vSelectAttrs, "", "", "bta.amount")
      End If

      'BTA line-types 'X' (exclude reversals)
      Dim vAnsiJoinsX As New AnsiJoins
      With vAnsiJoinsX
        .Add("event_booking_transactions ebt", "eb.event_number", "ebt.event_number", "eb.booking_number", "ebt.booking_number")
        .Add("batch_transaction_analysis bta", "ebt.batch_number", "bta.batch_number", "ebt.transaction_number", "bta.transaction_number", "ebt.line_number", "bta.line_number")
        .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
        .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
      End With
      Dim vWhereFieldsX As New CDBFields(New CDBField("bta.line_type", CDBField.FieldTypes.cftCharacter, "X"))
      With vWhereFieldsX
        .AddJoin("eb.batch_number", "bta.batch_number")
        .AddJoin("eb.transaction_number", "bta.transaction_number")
      End With
      Dim vSQLX As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, SUM(bta.amount) AS amount", "event_bookings eb", vWhereFieldsX, "", vAnsiJoinsX)
      vSQLX.GroupBy = "eb.event_number, eb.booking_number"

      'Financial Adjustments - Credits & Debits
      Dim vFASQLC As String = ""
      Dim vFASQLD As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
        Dim vAnsiJoinsFA As New AnsiJoins
        With vAnsiJoinsFA
          .Add("event_booking_transactions ebt", "ebt.event_number", "eb.event_number", "ebt.booking_number", "eb.booking_number")
          .Add("batch_transaction_analysis bta", "bta.batch_number", "ebt.batch_number", "bta.transaction_number", "ebt.transaction_number", "bta.line_number", "ebt.line_number")
          .Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
          .Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
        End With

        Dim vWhereFieldsFAC As New CDBFields(New CDBField("bta.member_number", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual))
        With vWhereFieldsFAC
          .Add("bta.member_number#2", CDBField.FieldTypes.cftInteger, mvEnv.Connection.DBToString("ebt.booking_number"))
          .Add("eb.batch_number", CDBField.FieldTypes.cftInteger, "bta.batch_number", CDBField.FieldWhereOperators.fwoNotEqual)
        End With
        Dim vWhereFieldsFAD As New CDBFields
        vWhereFieldsFAD.Clone(vWhereFieldsFAC)
        vWhereFieldsFAC.Add("transaction_sign", "C")
        vWhereFieldsFAD.Add("transaction_sign", "D")
        'Credits
        Dim vSQLFAC As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, transaction_sign, SUM(bta.amount) AS amount", "event_bookings eb", vWhereFieldsFAC, "", vAnsiJoinsFA)
        vSQLFAC.GroupBy = "eb.event_number, eb.booking_number, transaction_sign"
        vFASQLC = vSQLFAC.SQL
        'Debits
        Dim vAnsiJoinsFAD As New AnsiJoins
        For Each vJoin As AnsiJoin In vAnsiJoinsFA
          vAnsiJoinsFAD.Add(vJoin)
        Next
        vAnsiJoinsFAD.AddLeftOuterJoin("reversals r", "bta.batch_number", "r.batch_number", "bta.transaction_number", "r.transaction_number", "bta.line_number", "r.line_number")
        vWhereFieldsFAD.Add("r.was_batch_number", "")
        Dim vSQLFAD As New SQLStatement(mvEnv.Connection, "eb.event_number, eb.booking_number, transaction_sign, SUM(bta.amount * -1) AS amount", "event_bookings eb", vWhereFieldsFAD, "", vAnsiJoinsFAD)
        vSQLFAD.GroupBy = "eb.event_number, eb.booking_number, transaction_sign"
        vFASQLD = vSQLFAD.SQL
      End If

      'Main SQL
      Dim vAnsiJoins As New AnsiJoins
      Dim vTable As String
      If Val(mvParameters.OptionalValue("SessionNumber", "0")) > 0 Then
        vTable = "session_bookings sb"
        vAnsiJoins.Add("event_bookings eb", "sb.booking_number", "eb.booking_number")
        vWhereFields.Add("sb.session_number", mvParameters("SessionNumber").IntegerValue)
        vWhereFields.Add("eb.event_number", mvParameters("EventNumber").IntegerValue)
      Else
        vTable = "event_bookings eb"
        If mvParameters.Exists("BookingNumber") Then
          vWhereFields.Add("eb.booking_number", mvParameters("BookingNumber").IntegerValue)
        Else
          vWhereFields.Add("eb.event_number", mvParameters("EventNumber").IntegerValue)
        End If
      End If
      vWhereFields.Add("s.session_type", "0")
      With vAnsiJoins
        .Add("event_bookings eb2", "eb.event_number", "eb2.event_number", "eb.booking_number", "eb2.booking_number")
        .Add("sessions s", "eb.event_number", "s.event_number")
        .Add("event_booking_options ebo", "eb.option_number", "ebo.option_number")
        .Add("contacts c", "eb.contact_number", "c.contact_number")
        .Add("addresses a", "eb.address_number", "a.address_number")
        .AddLeftOuterJoin("batch_transaction_analysis bta", "eb2.batch_number", "bta.batch_number", "eb2.transaction_number", "bta.transaction_number", "eb2.line_number", "bta.line_number")
        .AddLeftOuterJoin("organisations o", "a.address_number", "o.address_number")
        .AddLeftOuterJoin("invoices i", "eb.batch_number", "i.batch_number", "eb.transaction_number", "i.transaction_number")
        .AddLeftOuterJoin("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
        .AddLeftOuterJoin("batches b", "eb.batch_number", "b.batch_number")
        .AddLeftOuterJoin("batch_transactions bt", "eb.batch_number", "bt.batch_number", "eb.transaction_number", "bt.transaction_number")
        .AddLeftOuterJoin("contacts pc", "bt.contact_number", "pc.contact_number")

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
          .AddLeftOuterJoin("(" & vSQLX.SQL & ") bta2", "eb2.event_number", "bta2.event_number", "eb2.booking_number", "bta2.booking_number")
          If vFASQLC.Length > 0 Then .AddLeftOuterJoin("(" & vFASQLC & ") ebtfac", "eb2.event_number", "ebtfac.event_number", "eb2.booking_number", "ebtfac.booking_number")
          If vFASQLD.Length > 0 Then .AddLeftOuterJoin("(" & vFASQLD & ") ebtfad", "eb2.event_number", "ebtfad.event_number", "eb2.booking_number", "ebtfad.booking_number")
        End If
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "eb.booking_number," & vConAttrs & "," & RemoveBlankItems(vAttrs) & vInvAttr, vTable, vWhereFields, "c.surname, c.forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "booking_number,CONTACT_NAME," & vSelectAttrs & vInvAttr, "contact_number,,,,,")

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("BookingStatus").Length > 0 Then
          vRow.Item("BookingStatusCode") = vRow.Item("BookingStatus")
          vRow.Item("BookingStatus") = EventBooking.GetBookingStatusDescription(vRow.Item("BookingStatus"))
        End If
        vRow.Item("CreditSale") = If(vRow.Item("BatchType") = "CS", "Y", "N")
        If Len(vRow.Item("InvoiceRePrintCount")) > 0 Then
          If Val(vRow.Item("InvoiceNumber")) = 0 Or (Val(vRow.Item("InvoiceNumber")) > 0 And Val(vRow.Item("InvoiceRePrintCount")) < 0) Then
            vRow.Item("InvoicePrinted") = "N"     'Invoice not yet printed
          Else
            vRow.Item("InvoicePrinted") = "Y"     'Invoice printed
          End If
        Else
          vRow.Item("InvoicePrinted") = "N"
        End If
        'If Len(vRow.Item("BookingAmount")) = 0 Then vRow.Item("BookingAmount") = "0"
        vRow.SetYNValue("CreditSale", False, True)
        vRow.SetYNValue("InvoicePrinted", False, True)
      Next
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetEventBookingSessions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventBookingSessions")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ s.session_number,session_desc,s.subject,subject_desc,s.skill_level,skill_level_desc,start_date,end_date,start_time,end_time"
      If mvParameters.Exists("SessionNumbers") Then
        vSQL = vSQL & " FROM sessions s, subjects su, skill_levels sl WHERE session_number IN ( " & mvParameters("SessionNumbers").Value & ") AND s.session_type <> '0' "
      Else
        vSQL = vSQL & " FROM session_bookings sb, sessions s, subjects su, skill_levels sl WHERE booking_number = " & mvParameters("BookingNumber").LongValue & " AND sb.session_number = s.session_number AND s.session_type <> '0' "
        If mvParameters.Exists("SessionNumber") Then vSQL = vSQL & " AND sb.session_number = " & mvParameters("SessionNumber").Value
      End If
      vSQL = vSQL & " AND s.subject = su.subject AND s.skill_level = sl.skill_level ORDER BY start_date, start_time"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetEventCandidates(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventCandidates")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "c.contact_number," & ContactNameItems()
      Dim vSQL As String = "SELECT /* SQLServerCSC */ "
      If mvParameters("SessionNumber").LongValue = mvParameters("BaseItemNumber").LongValue Then
        vSQL = vSQL & "DISTINCT "
        vWhereFields.Add("sb.event_number", mvParameters("EventNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)
      Else
        vWhereFields.Add("sb.session_number", mvParameters("SessionNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)
      End If
      vSQL = vSQL & vConAttrs & " FROM session_bookings sb,delegates d,contacts c WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      vSQL = vSQL & " AND d.booking_number = sb.booking_number AND c.contact_number = d.contact_number ORDER BY surname,forenames"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "contact_number,,")
    End Sub
    Private Sub GetEventContacts(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventContacts")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = ContactNameItems() & ",ec.address_number,event_contact_relationship,ec.amended_by,ec.amended_on,contact_number,notes"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vConAttrs & ",ec.address_number,event_contact_relationship,ec.amended_by,ec.amended_on,ec.notes"
      vSQL = vSQL & " FROM event_contacts ec, contacts c WHERE event_number = " & mvParameters("EventNumber").LongValue & " AND c.contact_number = ec.contact_number"
      If mvParameters.Exists("ContactNumber") Then vSQL = vSQL & " AND ec.contact_number = " & mvParameters("ContactNumber").LongValue
      vSQL = vSQL & " ORDER BY surname,forenames"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "contact_number,,")
    End Sub
    Private Sub GetEventCosts(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventCosts")
      Dim vAttrs As String = "sc.sundry_cost_type,full_amount,full_payment_date,deposit_amount,deposit_date,amount,payment_date,sc.amended_by,sc.amended_on,sundry_cost_number,sundry_cost_type_desc"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        vAttrs = vAttrs & ",sponsorship_value,contact_number,address_number,item_received,reserve_amount,sold_amount,supplier_contact_number,supplier_address_number,notes"
      Else
        vAttrs = vAttrs & ",,,,,,,,,"
      End If
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & " FROM sundry_costs sc, sundry_cost_types sct WHERE "
      If mvParameters.Exists("SundryCostNumber") Then
        vSQL = vSQL & "sundry_cost_number = " & mvParameters("SundryCostNumber").LongValue
      Else
        vSQL = vSQL & "unique_id = " & mvParameters("EventNumber").LongValue
      End If
      vSQL = vSQL & " AND record_type = 'E' AND sc.sundry_cost_type = sct.sundry_cost_type"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetEventDelegateIncome(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "fh.batch_number,fh.transaction_number,fhd.line_number,fh.transaction_date,fhd.amount,fhd.source,s.source_desc,fhd.product,p.product_desc,fhd.status,fh.notes"
      vWhereFields.Add("fh.contact_number", mvParameters("ContactNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("es.event_number", mvParameters("EventNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("financial_history_details fhd", "fh.batch_number", "fhd.batch_number", "fh.transaction_number", "fhd.transaction_number")
      vAnsiJoins.Add("event_sources es", "fhd.source", "es.source")
      vAnsiJoins.Add("sources s", "es.source", "s.source")
      vAnsiJoins.AddLeftOuterJoin("products p", "fhd.product", "p.product")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "financial_history fh", vWhereFields, "fh.transaction_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub
    Private Sub GetEventFinancialHistory(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventFinancialHistory")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = ContactNameItems() & ",fh.contact_number,fh.address_number,fh.batch_number,fh.transaction_number,fhd.line_number,fh.transaction_date,fhd.amount,fhd.source,s.source_desc,fhd.product,p.product_desc,fhd.status,fh.notes"
      vWhereFields.Clear()
      vWhereFields.Add("es.event_number", CDBField.FieldTypes.cftLong, mvParameters("EventNumber").Value)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vConAttrs & ",fh.address_number,fh.batch_number,fh.transaction_number,fhd.line_number,fh.transaction_date,fhd.amount,fhd.source,s.source_desc,fhd.product,p.product_desc,fhd.status,fh.notes"
      vSQL = vSQL & " FROM event_sources es "
      vSQL = vSQL & " INNER JOIN financial_history_details fhd ON es.source = fhd.source "
      vSQL = vSQL & " INNER JOIN financial_history fh ON fhd.batch_number = fh.batch_number AND fhd.transaction_number = fh.transaction_number "
      vSQL = vSQL & " INNER JOIN sources s ON fhd.source = s.source "
      vSQL = vSQL & " INNER JOIN contacts c ON fh.contact_number = c.contact_number "
      vSQL = vSQL & " LEFT OUTER JOIN products p ON fhd.product = p.product "
      If vWhereFields.Count > 0 Then vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      vSQL = vSQL & " ORDER BY fh.transaction_date "
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs)
    End Sub
    Private Sub GetEventFinancialLinks(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventFinancialLinks")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = ContactNameItems() & ",fh.contact_number,fh.address_number,fh.batch_number,fh.transaction_number,fhd.line_number,fh.transaction_date,fhd.amount,fhd.source,s.source_desc,fhd.product,p.product_desc,fhd.status,fh.notes"
      vWhereFields.Clear()
      vWhereFields.Add("efl.event_number", CDBField.FieldTypes.cftLong, mvParameters("EventNumber").Value)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vConAttrs & ",fh.address_number,fh.batch_number,fh.transaction_number,fhd.line_number,fh.transaction_date,fhd.amount,fhd.source,s.source_desc,fhd.product,p.product_desc,fhd.status,fh.notes"
      vSQL = vSQL & " FROM event_financial_links efl "
      vSQL = vSQL & " INNER JOIN financial_history_details fhd ON efl.batch_number = fhd.batch_number AND efl.transaction_number = fhd.transaction_number AND efl.line_number = fhd.line_number "
      vSQL = vSQL & " INNER JOIN financial_history fh ON efl.batch_number = fh.batch_number AND efl.transaction_number = fh.transaction_number "
      vSQL = vSQL & " INNER JOIN sources s ON fhd.source = s.source "
      vSQL = vSQL & " INNER JOIN contacts c ON fh.contact_number = c.contact_number "
      vSQL = vSQL & " LEFT OUTER JOIN products p ON fhd.product = p.product"
      If vWhereFields.Count > 0 Then vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      vSQL = vSQL & " ORDER BY fh.transaction_date "
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs)
    End Sub
    Private Sub GetEventLoanItems(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventLoanItems")
      Dim vAttrs As String = "contact_number,address_number,li.product,product_desc,quantity_issued,issued,due,returned,quantity_returned,accept_as_complete,reference"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace(vAttrs, "reference", mvEnv.Connection.DBSpecialCol("li", "reference")) & " FROM loan_items li,products p WHERE due" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDateAndTime) & " AND (accept_as_complete = 'N' or accept_as_complete IS NULL) AND li.product = p.product ORDER by due"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetEventMailings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventMailings")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ em.mailing,mailing_desc,m.department,department_desc,history_only,marketing,direction,mailing_due,mailing_date,number_in_mailing,m.notes,em.amended_by,em.amended_on FROM event_mailings em "
      vSQL = vSQL & " INNER JOIN mailings m ON em.mailing = m.mailing "
      vSQL = vSQL & " INNER JOIN departments d ON d.department = m.department "
      vSQL = vSQL & " LEFT OUTER JOIN mailing_history mh ON em.mailing = mh.mailing "
      vSQL = vSQL & " WHERE em.event_number = " & mvParameters("EventNumber").LongValue
      If mvParameters.Exists("Mailing") Then vSQL = vSQL & " AND em.mailing = '" & mvParameters("Mailing").Value & "'"
      vSQL = vSQL & " ORDER BY em.mailing, mh.mailing_date"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL))
    End Sub
    Private Sub GetEventOrganiserData(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = "organiser,organiser_desc," & vContact.GetRecordSetFieldsNameAddressCountryPhone
      Dim vWhereFields As New CDBFields()
      vWhereFields.AddJoin("o.contact_number", "c.contact_number")
      vWhereFields.AddJoin("o.address_number", "a.address_number")
      vWhereFields.AddJoin("a.country", "co.country")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "organisers o, contacts c, addresses a, countries co", vWhereFields, "organiser_desc")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "organiser,organiser_desc,contact_number,CONTACT_NAME,address_number,ADDRESS_LINE,,,,")
      vAttrs = "organiser, " & vContact.GetRecordSetFieldsNameAddressCountryPhone
      vWhereFields.Clear()
      vWhereFields.AddJoin("o.invoice_contact", "c.contact_number")
      vWhereFields.AddJoin("o.invoice_address", "a.address_number")
      vWhereFields.AddJoin("a.country", "co.country")
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, vAttrs, "organisers o, contacts c, addresses a, countries co", vWhereFields, "organiser_desc").GetRecordSet
      While vRecordSet.Fetch()
        Dim vValue As String = vRecordSet.Fields(1).Value
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vValue = vRow.Item("Organiser") Then
            vContact.InitFromRecordSetNameAddressCountryPhone(vRecordSet)
            vRow.Item("InvoiceContactNumber") = vContact.ContactNumber.ToString
            vRow.Item("InvoiceContactName") = vContact.Name
            vRow.Item("InvoiceContactAddressNumber") = vContact.Address.AddressNumber.ToString
            vRow.Item("InvoiceContactAddressLine") = vContact.Address.AddressLine
          End If
        Next
      End While
      vRecordSet.CloseRecordSet()
    End Sub
    Private Sub GetEventOwners(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventOwners")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("event_number", mvParameters("EventNumber").LongValue)
      If mvParameters.Exists("Department") Then vWhereFields.Add("department", mvParameters("Department").Value)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ event_number,eo.department,department_desc,eo.amended_by,eo.amended_on FROM event_owners eo, departments d WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND eo.department = d.department ORDER BY department_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetEventPersonnel(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventPersonnel")
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = ",ep.session_number,ep.confirmed,ep.confirmed_by,ep.expenses,ep.expenses_received,ep.authorised_by,ep.authorised_on,ep.amended_by,ep.amended_on,session_desc,ep.task,ep.position,ep.organisation_name,ep.start_date,ep.start_time,ep.end_date,ep.end_time,ep.address_number,"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vContact.GetRecordSetFieldsName & vAttrs & mvEnv.Connection.DBSpecialCol("ep", "function") & ",ep.event_personnel_number"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        vSQL = vSQL & ",ep.issue_resources  AS issue_event_resources"
      Else
        vSQL = vSQL & ",epi.session_number  AS issue_event_resources"
      End If
      vSQL = vSQL & " FROM event_personnel ep INNER JOIN contacts c ON c.contact_number = ep.contact_number"
      vSQL = vSQL & " INNER JOIN sessions s ON s.session_number = ep.session_number"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        'vSQL = vSQL & " LEFT OUTER JOIN (SELECT event_personnel_number, session_number FROM event_personnel WHERE session_number = (session_number / 10000) * 10000) epi ON ep.event_personnel_number = epi.event_personnel_number "
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT event_personnel_number, session_number FROM event_personnel WHERE session_number = session_number epi ON ep.event_personnel_number = epi.event_personnel_number "
      End If
      vAttrs = vAttrs & "function,event_personnel_number,issue_event_resources"
      If mvParameters.Exists("EventPersonnelNumber") Then
        vSQL = vSQL & " WHERE ep.event_personnel_number = " & mvParameters("EventPersonnelNumber").LongValue
      Else
        If mvParameters.Exists("SessionNumber") Then
          vSQL = vSQL & " WHERE ep.session_number = " & mvParameters("SessionNumber").LongValue
        Else
          Dim vEvent As New CDBEvent(mvEnv)
          vEvent.Init(mvParameters("EventNumber").LongValue)
          Dim vSessionNumber As Integer = vEvent.LowestSessionNumber
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) = False Then
            If vEvent.MultiSession Then vSessionNumber = vSessionNumber + 1
          End If
          'vSQL = vSQL & " WHERE ep.session_number BETWEEN " & vSessionNumber & " AND " & vEvent.MaxItemNumber
          vSQL = vSQL & " WHERE s.event_number = " & vEvent.EventNumber
        End If
      End If
      vSQL = vSQL & " ORDER BY c.surname, c.initials, session_desc"
      vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "contact_number," & ContactNameItems() & vAttrs)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetChairPerson("ChairPerson")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
          vRow.SetYNValue("IssueEventResources")
        Else
          If Len(vRow.Item("IssueEventResources")) > 0 Then vRow.Item("IssueEventResources") = DataSelectionText.String15904 'Yes
        End If
      Next
    End Sub
    Private Sub GetEventPersonnelTasks(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventPersonnelTasks")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String
      Dim vConAttrs As String = Replace(vContact.GetRecordSetFieldsName, "c.contact_number", "")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then
        vAttrs = "ept.event_personnel_number,ept.event_personnel_task_number,ept.personnel_task,personnel_task_desc,ept.start_date,ept.start_time,ept.end_date,ept.end_time,ept.notes,ept.amended_by,ept.amended_on,ep.contact_number,ep.address_number"
      Else
        vAttrs = "ept.event_personnel_number,,ept.personnel_task,personnel_task_desc,,,,,,ept.amended_by,ept.amended_on,ep.contact_number,ep.address_number"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbOutlookIntegration) Then
        vAttrs = vAttrs & ",ept.external_task_id"
      Else
        vAttrs = vAttrs & ", NULL AS ept.external_task_id"
      End If
      vWhereFields.Clear()
      If mvParameters.Exists("EventNumber") Then vWhereFields.Add("ept.event_number", mvParameters("EventNumber").LongValue)
      If mvParameters.Exists("EventPersonnelNumber") Then vWhereFields.Add("ept.event_personnel_number", mvParameters("EventPersonnelNumber").LongValue)
      If mvParameters.Exists("EventPersonnelTaskNumber") Then vWhereFields.Add("ept.event_personnel_task_number", mvParameters("EventPersonnelTaskNumber").LongValue)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs & "," & vConAttrs) & " FROM event_personnel_tasks ept INNER JOIN personnel_tasks pt ON ept.personnel_task = pt.personnel_task"
      vSQL = vSQL & " LEFT OUTER JOIN event_personnel ep ON ept.event_personnel_number = ep.event_personnel_number"
      vSQL = vSQL & " LEFT OUTER JOIN contacts c ON ep.contact_number = c.contact_number"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis) Then vSQL = vSQL & " ORDER BY ept.start_date,ept.start_time,ept.end_date,ept.end_time"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, "CONTACT_NAME")
    End Sub
    Private Sub GetEventPIS(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventPIS")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "pis.event_number,event_pis_number,pis_number,pis.event_delegate_number,x.contact_number,issue_date,amount,banked_by,banked_on,reconciled_on,o.name"
      If mvParameters.Exists("EventNumber") Then vWhereFields.Add("pis.event_number", mvParameters("EventNumber").LongValue)
      If mvParameters.Exists("EventPisNumber") Then vWhereFields.Add("event_pis_number", mvParameters("EventPisNumber").LongValue)
      If mvParameters.Exists("EventDelegateNumber") Then vWhereFields.Add("event_delegate_number", mvParameters("EventDelegateNumber").LongValue)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vConAttrs & "," & vAttrs & " FROM event_pis pis LEFT OUTER JOIN (SELECT event_delegate_number, d.contact_number, " & vConAttrs & " FROM delegates d INNER JOIN contacts c ON d.contact_number = c.contact_number) x ON pis.event_delegate_number = x.event_delegate_number LEFT OUTER JOIN organisations o ON pis.banked_by = o.organisation_number WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), ContactNameItems() & "," & vAttrs)
      GetContactNames(pDataTable, "BankedBy", "BankedByContactName", "", True)
    End Sub
    Private Sub GetEventResources(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventResources")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "er.resource_number,xr.resource_desc,,quantity_required,quantity_issued,allocated,er.amended_by,er.amended_on,s.session_desc,s.session_number,er.product,copy_to,issue_basis,despatch_to,issued,er.notes,resource_type"
      vAttrs = vAttrs & ",organisation_number,xr.address_number,xr.contact_number,external_resource_type,obtained_on,return_by,returned_on,full_amount,full_payment_date,deposit_amount,deposit_date"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & ",p.product_desc,p2.product_desc AS standard_product," & vConAttrs & " FROM event_resources er INNER JOIN sessions s ON s.session_number = er.session_number"
      vSQL = vSQL & " LEFT OUTER JOIN external_resources xr ON er.resource_number = xr.resource_number"
      vSQL = vSQL & " LEFT OUTER JOIN internal_resources ir ON er.resource_number = ir.resource_number"
      vSQL = vSQL & " LEFT OUTER JOIN products p ON ir.product = p.product"
      vSQL = vSQL & " LEFT OUTER JOIN products p2 ON er.product = p2.product"
      vSQL = vSQL & " LEFT OUTER JOIN contacts c ON ir.resource_contact_number = c.contact_number WHERE "
      If mvParameters.Exists("SessionNumber") Then
        vSQL = vSQL & "er.session_number = " & mvParameters("SessionNumber").LongValue
      Else
        'Should at least have one session but there are some events in the database where there were 
        'no sessions
        Dim vSessionList As String = GetEventSessionList(mvParameters("EventNumber").LongValue)
        If vSessionList.Length = 0 Then vSessionList = "0"
        vSQL = vSQL & "s.event_number = " & mvParameters("EventNumber").LongValue  '"er.session_number in (" & vSessionList & ")"      ' BETWEEN " & mvParameters("LowestSessionNumber").LongValue & " AND " & mvParameters("MaxItemNumber").LongValue
      End If
      If mvParameters.HasValue("ResourceNumber") Then vSQL = vSQL & " AND er.resource_number = " & mvParameters("ResourceNumber").LongValue
      If mvParameters.HasValue("Product") Then vSQL = vSQL & " AND er.product = '" & mvParameters("Product").Value & "'"
      If mvParameters.Exists("CopyTo") Then vSQL = vSQL & " AND copy_to = '" & mvParameters("CopyTo").Value & "'"
      vSQL = vSQL & " ORDER BY er.resource_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, "product_desc,standard_product,CONTACT_NAME")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ResourceTypeCode") = "I" Then
          If Len(vRow.Item("InternalResourceName")) > 0 Then
            vRow.Item("ResourceDesc") = vRow.Item("InternalResourceName")
          Else
            vRow.Item("ResourceDesc") = vRow.Item("InternalProductDesc")
          End If
          vRow.Item("ResourceType") = "Internal"
        ElseIf vRow.Item("ResourceTypeCode") = "E" Then
          vRow.Item("ResourceType") = "External"
        Else
          vRow.Item("ResourceType") = "Standard"
          vRow.Item("ResourceDesc") = vRow.Item("StandardProductDesc")
        End If
      Next
    End Sub
    Private Sub GetEventResults(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventResults")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "str.session_number,str.test_number,test_desc,test_result,certificate_number,str.notes,str.amended_by,str.amended_on"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vConAttrs & "," & vAttrs & " FROM session_test_results str,session_tests st,contacts c WHERE "
      If mvParameters.Exists("SessionNumber") Then
        vSQL = vSQL & " str.session_number = " & mvParameters("SessionNumber").LongValue
      Else
        vSQL = vSQL & " str.session_number BETWEEN " & mvParameters("LowestSessionNumber").LongValue & " AND " & mvParameters("MaxItemNumber").LongValue
      End If
      If mvParameters.Exists("ContactNumber") Then vSQL = vSQL & " AND str.contact_number = " & mvParameters("ContactNumber").LongValue
      If mvParameters.Exists("TestNumber") Then vSQL = vSQL & " AND str.test_number = " & mvParameters("TestNumber").LongValue
      vSQL = vSQL & " AND st.session_number = str.session_number AND st.test_number = str.test_number AND str.contact_number = c.contact_number ORDER BY surname,forenames,str.test_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, ContactNameItems() & ",contact_number," & vAttrs)
    End Sub
    Private Sub GetEventRoomBookingAllocations(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventRoomBookingAllocations")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "rbl.address_number,room_date,room_id,room_booking_link_number"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vAttrs & "," & vConAttrs & ",position,name FROM room_booking_links rbl INNER JOIN contacts c ON rbl.contact_number = c.contact_number"
      vSQL = vSQL & " LEFT OUTER JOIN contact_positions cp ON rbl.contact_number = cp.contact_number AND rbl.address_number = cp.address_number"
      vSQL = vSQL & " LEFT OUTER JOIN organisations o ON cp.organisation_number = o.organisation_number"
      vSQL = vSQL & " WHERE rbl.room_booking_number = " & mvParameters("RoomBookingNumber").LongValue
      If mvParameters.Exists("RoomBookingLinkNumber") Then vSQL = vSQL & " AND room_booking_link_number = " & mvParameters("RoomBookingLinkNumber").LongValue
      vSQL = vSQL & " ORDER BY room_date, surname, room_id"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, "contact_number,position,name,CONTACT_NAME")
    End Sub
    Private Sub GetEventRoomBookings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventRoomBookings")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "room_type_desc,crb.number_of_rooms,crb.from_date,crb.to_date,crb.booked_date,crb.cancelled_on,crb.cancelled_by,booking_status,crb.rate,crb.notes,sales_contact_number,crb.confirmed_date,crb.cancellation_reason,crb.cancellation_source,crb.amended_by,crb.amended_on,crb.address_number,rbb.room_type,rbb.product,o.organisation_number,crb.batch_number,crb.transaction_number,crb.line_number,capacity,enforce_allocation"
      Dim vSQL As String = "SELECT  /* SQLServerCSC */ room_booking_number," & vConAttrs & "," & vAttrs & " FROM contact_room_bookings crb"
      vSQL = vSQL & " INNER JOIN contacts c ON crb.contact_number = c.contact_number"
      vSQL = vSQL & " INNER JOIN room_block_bookings rbb ON crb.block_booking_number = rbb.block_booking_number"
      vSQL = vSQL & " INNER JOIN room_types rt ON rbb.room_type = rt.room_type"
      vSQL = vSQL & " LEFT OUTER JOIN organisations o ON crb.address_number = o.address_number"
      vSQL = vSQL & " WHERE "
      If mvParameters.Exists("RoomBookingNumber") Then
        vSQL = vSQL & " crb.room_booking_number = " & mvParameters("RoomBookingNumber").LongValue
      Else
        vSQL = vSQL & " crb.event_number = " & mvParameters("EventNumber").LongValue
      End If
      vSQL = vSQL & "  ORDER BY surname, forenames"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), "room_booking_number,CONTACT_NAME," & vAttrs, "contact_number,,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetEventSessionActivities(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventSessionActivities")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("event_number", mvParameters("EventNumber").LongValue)
      If mvParameters.Exists("SessionNumber") Then vWhereFields.Add("session_number", mvParameters("SessionNumber").LongValue)
      If mvParameters.Exists("Activity") Then vWhereFields.Add("sa.activity", mvParameters("Activity").Value)
      If mvParameters.Exists("ActivityValue") Then vWhereFields.Add("sa.activity_value", mvParameters("ActivityValue").Value)
      Dim vSQL As String = "SELECT /* SQLServerCSC */ event_number,session_number,sa.activity,sa.activity_value,activity_desc,activity_value_desc,sa.amended_by,sa.amended_on FROM session_activities sa, activities a, activity_values av WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND sa.activity = a.activity AND sa.activity = av.activity AND sa.activity_value = av.activity_value ORDER BY session_number, sa.activity"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetEventSessions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventSessions")
      If mvParameters.Exists("BaseSessionOnly") Then
        Dim vSQL As String = "SELECT /* SQLServerCSC */ session_number,session_desc,s.session_type, NULL AS session_type_desc,s.subject,NULL as subject_desc,s.skill_level,NULL as kill_level_desc,start_date,end_date,start_time,end_time,location,minimum_attendees,maximum_attendees,target_attendees,number_interested,number_of_attendees,number_on_waiting_list,maximum_on_waiting_list,notes,venue_booking_number,s.amended_by,s.amended_on,maximum_attendees - number_of_attendees AS available,cpd_approval_status,cpd_awarding_body,cpd_category,cpd_date_approved,cpd_notes,cpd_points,cpd_year,maximum_on_waiting_list - number_on_waiting_list AS wavailable"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventLongDescription) Then vSQL = vSQL & ",long_description"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbOutlookIntegration) Then
          vSQL = vSQL & ",s.external_appointment_id"
        Else
          vSQL = vSQL & ", NULL AS s.external_appointment_id"
        End If
        vSQL = vSQL & " FROM sessions s WHERE event_number = " & mvParameters("EventNumber").LongValue
        vSQL = vSQL & " AND session_number = " & GetBaseSessionNumber(mvParameters("EventNumber").LongValue)
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, , If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventLongDescription), "", ","))
        pDataTable.Columns("Available").FieldType = CDBField.FieldTypes.cftLong
        pDataTable.Columns("WaitingAvailable").FieldType = CDBField.FieldTypes.cftLong
      Else
        'TODO Convert to new SQL Syntax
        'NYI("GetEventSessions")
        Dim vSQL As String = "SELECT /* SQLServerCSC */ session_number,session_desc,s.session_type,session_type_desc,s.subject,subject_desc,s.skill_level,skill_level_desc,start_date,end_date,start_time,end_time,location,minimum_attendees,maximum_attendees,target_attendees,number_interested,number_of_attendees,number_on_waiting_list,maximum_on_waiting_list,notes,venue_booking_number,s.amended_by,s.amended_on,maximum_attendees - number_of_attendees AS available,cpd_approval_status,cpd_awarding_body,cpd_category,cpd_date_approved,cpd_notes,cpd_points,cpd_year,maximum_on_waiting_list - number_on_waiting_list AS wavailable"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventLongDescription) Then vSQL = vSQL & ",long_description"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbOutlookIntegration) Then
          vSQL = vSQL & ",s.external_appointment_id"
        Else
          vSQL = vSQL & ", NULL AS s.external_appointment_id"
        End If
        vSQL = vSQL & " FROM sessions s,session_types st, subjects su, skill_levels sl WHERE event_number = " & mvParameters("EventNumber").LongValue
        If mvParameters.Exists("SessionNumber") Then vSQL = vSQL & " AND session_number = " & mvParameters("SessionNumber").LongValue
        If mvParameters.Exists("BaseItemNumber") Then vSQL = vSQL & " AND session_number > " & GetBaseSessionNumber(mvParameters("EventNumber").LongValue) 'mvParameters("BaseItemNumber").LongValue
        vSQL = vSQL & " AND s.session_type = st.session_type AND s.subject = su.subject AND s.skill_level = sl.skill_level ORDER BY start_date, start_time"
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, , If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventLongDescription), "", ","))
        pDataTable.Columns("Available").FieldType = CDBField.FieldTypes.cftLong
        pDataTable.Columns("WaitingAvailable").FieldType = CDBField.FieldTypes.cftLong
      End If

    End Sub
    Private Sub GetEventSessionTests(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventSessionTests")

      Dim vFields As String = "st.session_number, test_number, test_desc, grade_data_type, minimum_value, maximum_value, pattern, st.amended_by, st.amended_on, grade_data_type As grade_data_type_desc"
      Dim vAnsiJoin As New AnsiJoins({New AnsiJoin("sessions s", "st.session_number", "s.session_number")})
      Dim vWhereField As New CDBFields

      If mvParameters.Exists("SessionNumber") Then
        vWhereField.Add("s.session_number", mvParameters("SessionNumber").LongValue)
      Else
        vWhereField.Add("s.event_number", mvParameters("EventNumber").LongValue)
      End If

      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vFields, "session_tests st", vWhereField, "", vAnsiJoin))
      GetLookupData(pDataTable, "GradeDataType", "session_tests", "grade_data_type")


    End Sub
    Private Sub GetEventSources(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventSources")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ es.source,source_desc,es.amended_by,es.amended_on,history_only FROM event_sources es, sources s WHERE event_number = " & mvParameters("EventNumber").LongValue
      If mvParameters.Exists("Source") Then vSQL = vSQL & " AND es.source = '" & mvParameters("Source").Value & "'"
      vSQL = vSQL & " AND es.source = s.source ORDER BY es.source"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetEventSubmissions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventSubmissions")
      'First get all records for which there is an Assessor.
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = ",es.paper_title,es.submitted,es.amended_by,es.amended_on,es.submission_number,a.address,es.address_number,es.subject,su.subject_desc,es.skill_level,sk.skill_level_desc,es.forwarded,es.assessor,es.returned,es.result,assessor_name"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace(vContact.GetRecordSetFieldsName, ",", ",c.") & Replace(vAttrs, "assessor_name", "ca.label_name AS assessor_name")
      vSQL = vSQL & " FROM event_submissions es, addresses a, subjects su, skill_levels sk, contacts c, contacts ca"
      If mvParameters.Exists("SubmissionNumber") Then
        vSQL = vSQL & " WHERE es.submission_number=" & mvParameters("SubmissionNumber").LongValue
      Else
        vSQL = vSQL & " WHERE es.event_number=" & mvParameters("EventNumber").LongValue
      End If
      vSQL = vSQL & " AND ca.contact_number = es.assessor AND c.contact_number = es.contact_number AND es.address_number = a.address_number AND es.skill_level = sk.skill_level AND es.subject = su.subject"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "contact_number," & ContactNameItems() & vAttrs)
      'Now get all records for which there is no Assesor.
      vAttrs = ",es.paper_title,es.submitted,es.amended_by,es.amended_on,es.submission_number,a.address,es.address_number,es.subject,su.subject_desc,es.skill_level,sk.skill_level_desc,es.forwarded,es.assessor,es.returned,es.result"
      vSQL = "SELECT " & vContact.GetRecordSetFieldsName & vAttrs
      vSQL = vSQL & " FROM event_submissions es, addresses a, subjects su, skill_levels sk, contacts c"
      If mvParameters.Exists("SubmissionNumber") Then
        vSQL = vSQL & " WHERE es.submission_number=" & mvParameters("SubmissionNumber").LongValue
      Else
        vSQL = vSQL & " WHERE es.event_number=" & mvParameters("EventNumber").LongValue
      End If
      vSQL = vSQL & " AND c.contact_number = es.contact_number AND es.address_number = a.address_number AND es.skill_level = sk.skill_level AND es.subject = su.subject AND es.assessor IS NULL"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, "contact_number," & ContactNameItems() & vAttrs, ",")
      'Order the table by Submitted.
      pDataTable.ReOrderRowsByColumn("Submitted")
      'vDescriptions.Add( DataSelectionText.String18158), "A"    'Accepted)
      'vDescriptions.Add( DataSelectionText.String18159), "F"    'Failed)
      For Each vRow As CDBDataRow In pDataTable.Rows
        'vRow.SetDescriptionFromCode "Result", vDescriptions
        If Len(vRow.Item("Forwarded")) = 0 Then vRow.Item("Returned") = ""
        If Len(vRow.Item("Returned")) = 0 Then vRow.Item("SubmissionResult") = ""
      Next
    End Sub
    Private Sub GetEventTopics(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventTopics")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("event_number", mvParameters("EventNumber").LongValue)
      If mvParameters.Exists("Topic") Then vWhereFields.Add("et.topic", mvParameters("Topic").Value)
      If mvParameters.Exists("SubTopic") Then vWhereFields.Add("et.sub_topic", mvParameters("SubTopic").Value)
      Dim vAttrs As String = "et.topic,topic_desc,et.sub_topic,sub_topic_desc,quantity,notes,et.amended_by,et.amended_on"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventTopicNotes) Then vAttrs = Replace(vAttrs, "notes", "")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & RemoveBlankItems(vAttrs) & " FROM event_topics et,topics t,sub_topics st WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND et.topic = t.topic AND et.topic = st.topic AND et.sub_topic = st.sub_topic ORDER BY et.topic,et.sub_topic"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetEventVenueBookings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEventVenueBookings")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ venue_booking_number,evb.venue,venue_desc,venue_reference,full_amount,full_payment_date,deposit_amount,deposit_date,amount,payment_date,confirmed_by,confirmed_on,evb.amended_by,evb.amended_on,v.organisation_number,v.contact_number FROM event_venue_bookings evb, venues v WHERE "
      vSQL = vSQL & "evb.event_number = " & mvParameters("EventNumber").LongValue & " AND evb.venue = v.venue"
      If mvParameters.Exists("VenueBookingNumber") Then vSQL = vSQL & " AND venue_booking_number= " & mvParameters("VenueBookingNumber").LongValue
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, , ",,,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If Val(vRow.Item("OrganisationNumber")) > 0 Then
          Dim vOrg As New Organisation(mvEnv)
          vOrg.Init()
          Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOrg.GetRecordSetFields(OrganisationRecordSetTypes.ortAll Or OrganisationRecordSetTypes.ortAddress) & " FROM organisations o, addresses a WHERE o.organisation_number = " & vRow.Item("OrganisationNumber") & " AND a.address_number = o.address_number")
          If vRecordSet.Fetch Then
            vOrg.InitFromRecordSet(mvEnv, vRecordSet, OrganisationRecordSetTypes.ortAll Or OrganisationRecordSetTypes.ortAddress)
            vRow.Item("OrganisationName") = vOrg.Name
            vRow.Item("Address") = vOrg.Address.AddressLine
            vRow.Item("Telephone") = vOrg.PhoneNumber
          End If
          vRecordSet.CloseRecordSet()
        End If
        If Val(vRow.Item("ContactNumber")) > 0 Then
          Dim vContact As New Contact(mvEnv)
          Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail) & " FROM contacts c WHERE contact_number = " & vRow.Item("ContactNumber"))
          If vRecordSet.Fetch Then
            vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail)
            vRow.Item("ContactName") = vContact.LabelName
            If vRow.Item("Telephone") = "" Then vRow.Item("Telephone") = vContact.PhoneNumber
          End If
          vRecordSet.CloseRecordSet()
        End If
      Next
    End Sub
    Private Sub GetEventVenueData(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "venue,venue_desc"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEventVenueCapacity) Then
        vFields = vFields & ", venue_capacity"
      Else
        vFields = vFields & ", NULL AS venue_capacity"
      End If
      vFields = vFields & ",location,v.organisation_number,name," & New Contact(mvEnv).GetRecordSetFieldsNameAddressCountryPhone
      Dim vWhereFields As New CDBFields()
      If mvParameters.ContainsKey("Venue") Then vWhereFields.Add("v.venue", mvParameters("Venue").Value)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("organisations o", "v.organisation_number", "o.organisation_number")
      vAnsiJoins.AddLeftOuterJoin("addresses a", "v.address_number", "a.address_number")
      vAnsiJoins.AddLeftOuterJoin("countries co", "a.country", "co.country")
      vAnsiJoins.AddLeftOuterJoin("contacts c", "v.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "venues v", vWhereFields, "venue_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "venue,venue_desc,venue_capacity,location,organisation_number,name,address_number,ADDRESS_LINE,contact_number,CONTACT_NAME,CONTACT_TELEPHONE")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("VenueTelephone").ToString.Length = 0 Then
          If IntegerValue(vRow.Item("VenueOrganisationNumber").ToString) > 0 Then
            Dim vOrganisation As New Organisation(mvEnv)
            vOrganisation.Init(IntegerValue(vRow.Item("VenueOrganisationNumber").ToString))
            If vOrganisation.Existing Then
              vRow.Item("VenueTelephone") = vOrganisation.PhoneNumber
            End If
          End If
        End If
      Next
    End Sub
    Private Sub GetEVWaitingBookings(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEVWaitingBookings")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = ContactNameItems() & ",booking_number,booking_date,quantity,option_desc,booking_status,contact_number"
      Dim vSQL As String = "SELECT  /* SQLServerCSC */ eb.booking_number," & vConAttrs & ",booking_date,eb.quantity,option_desc,booking_status"
      vSQL = vSQL & " FROM event_bookings eb, event_booking_options ebo, contacts c "
      vSQL = vSQL & " WHERE eb.event_number = " & mvParameters("EventNumber").Value
      vSQL = vSQL & " AND eb.booking_status IN ('" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaiting) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingCreditSale) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid) & "')"
      vSQL = vSQL & " AND cancellation_reason IS NULL"
      vSQL = vSQL & " AND eb.option_number = ebo.option_number AND eb.contact_number = c.contact_number ORDER by booking_date,eb.booking_number,surname,initials"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "contact_number,,")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("BookingStatus").Length > 0 Then vRow.Item("BookingStatus") = EventBooking.GetBookingStatusDescription(vRow.Item("BookingStatus"))
      Next
    End Sub
    Private Sub GetEVWaitingDelegates(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetEVWaitingDelegates")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vAttrs As String = "booking_number," & ContactNameItems() & ",contact_number"
      Dim vSQL As String = "SELECT  /* SQLServerCSC */ eb.booking_number," & vConAttrs
      vSQL = vSQL & " FROM event_bookings eb, delegates d, contacts c "
      vSQL = vSQL & " WHERE eb.event_number = " & mvParameters("EventNumber").Value
      vSQL = vSQL & " AND eb.booking_status IN ('" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaiting) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingCreditSale) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingInvoiced) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsWaitingPaid) & "')"
      vSQL = vSQL & " AND cancellation_reason IS NULL"
      vSQL = vSQL & " AND d.booking_number = eb.booking_number "
      vSQL = vSQL & " AND d.contact_number = c.contact_number ORDER by booking_date,eb.booking_number,surname,initials"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "contact_number,,")
    End Sub
    Private Sub GetFinancialHistoryDetails(ByVal pDataTable As CDBDataTable)
      Dim vBaseAttrs As String = "x.line_number,x.product,{0},x.rate,{1},x.distribution_code,x.quantity,x.amount,x.vat_amount,status,x.source,source_desc,{2},"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vBaseAttrs &= "x.currency_amount,x.currency_vat_amount,{3}"
      Else
        vBaseAttrs &= ",,"
      End If
      vBaseAttrs &= ",x.sales_contact_number, bta.notes,bta.issued,bta.warehouse,x.invoice_payment,{4}"

      Dim vBaseCols As String = vBaseAttrs
      vBaseAttrs &= ",ba.rgb_value AS rgb_amount, ba.rgb_value AS rgb_currency_amount"
      vBaseCols &= ",rgb_amount, rgb_currency_amount"

      'Nested SQL
      Dim vSubTableAttrs As String = "batch_number,transaction_number,line_number,fhd.product,product_desc,fhd.rate,rate_desc,fhd.distribution_code,quantity,amount,vat_amount,status,fhd.source,source_desc,vat_rate_desc,"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vSubTableAttrs &= "fhd.currency_amount,fhd.currency_vat_amount,r.currency_code"
      Else
        vSubTableAttrs &= ",,"
      End If
      vSubTableAttrs &= ",sales_contact_number,fhd.invoice_payment,p.stock_item"

      Dim vNestedWhereFields As New CDBFields(New CDBField("fhd.batch_number", mvParameters("BatchNumber").LongValue))
      vNestedWhereFields.Add("fhd.transaction_number", mvParameters("TransactionNumber").LongValue)

      Dim vNestedAnsiJoins As New AnsiJoins({New AnsiJoin("products p", "fhd.product", "p.product")})
      vNestedAnsiJoins.Add("rates r", "fhd.product", "r.product", "fhd.rate", "r.rate")
      vNestedAnsiJoins.Add("vat_rates vr", "fhd.vat_rate", "vr.vat_rate")
      vNestedAnsiJoins.Add("sources s", "fhd.source", "s.source")

      Dim vNestedSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vSubTableAttrs), "financial_history_details fhd", vNestedWhereFields, String.Empty, vNestedAnsiJoins)

      'Main SQL
      Dim vMainTable As String = "(" & vNestedSQL.SQL & ") x"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("batch_transaction_analysis bta", "x.batch_number", "bta.batch_number", "x.transaction_number", "bta.transaction_number", "x.line_number", "bta.line_number")
      vAnsiJoins.AddLeftOuterJoin("batches b", "x.batch_number", "b.batch_number")
      vAnsiJoins.AddLeftOuterJoin("bank_accounts ba", "b.bank_account", "ba.bank_account")

      vBaseCols &= ",,,,,,"
      Dim vAttrs As String = String.Format(vBaseAttrs, {"product_desc", "rate_desc", "vat_rate_desc", "x.currency_code", "x.stock_item"})
      Dim vCols As String = String.Format(vBaseCols, {"product_desc", "rate_desc", "vat_rate_desc", "x.currency_code", "x.stock_item"})
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vMainTable, Nothing, String.Empty, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)


      'Process data where FHD.Product is null
      'Nested SQL
      vSubTableAttrs = "batch_number,transaction_number,line_number,fhd.product,fhd.rate,fhd.distribution_code,quantity,amount,vat_amount,status,fhd.source,source_desc,"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vSubTableAttrs &= "fhd.currency_amount,fhd.currency_vat_amount"
      Else
        vSubTableAttrs &= ","
      End If
      vSubTableAttrs &= ",sales_contact_number,vat_rate,fhd.invoice_payment"

      vNestedAnsiJoins = New AnsiJoins()
      vNestedAnsiJoins.Add("sources s", "fhd.source", "s.source")

      vNestedWhereFields.Add("fhd.product", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual)

      vNestedSQL = New SQLStatement(mvEnv.Connection, RemoveBlankItems(vSubTableAttrs), "financial_history_details fhd", vNestedWhereFields, String.Empty, vNestedAnsiJoins)

      'Main SQL
      vMainTable = "(" & vNestedSQL.SQL & ") x"
      vAttrs = String.Format(vBaseAttrs, {"x.product AS product_desc", "x.rate AS rate_desc", "x.vat_rate AS vat_rate_desc", "x.product AS currency_code", ""})
      vCols = String.Format(vBaseCols, {"product_desc", "rate_desc", "vat_rate_desc", "currency_code", ""})

      vSQLStatement = New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vMainTable, Nothing, String.Empty, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)

      GetSalesContactNames(pDataTable)
      GetDescriptions(pDataTable, "DistributionCode")
      GetPaymentPlanInfo(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue(("StockProduct"))
        vRow.SetYNValue(("InvoicePayment"))
        Me.CheckAmountRGBValue(vRow)
      Next
    End Sub

    Private Sub GetFundraisingEventAnalysis(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      vContact.Init()
      Dim vFields As String = "fea.batch_number,fea.transaction_number,fea.line_number,bt.transaction_date,bta.amount,bta.notes,"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("contact_fundraising_number", mvParameters("ContactFundraisingNumber").LongValue)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("batch_transactions bt", "fea.batch_number", "bt.batch_number", "fea.transaction_number", "bt.transaction_number")
      vAnsiJoins.AddLeftOuterJoin("contacts c", "bt.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("batch_transaction_analysis bta", "fea.batch_number", "bta.batch_number", "fea.transaction_number", "bta.transaction_number", "fea.line_number", "bta.line_number")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields & vContact.GetRecordSetFieldsName, "fundraising_event_analysis fea", vWhereFields, "transaction_date desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields & "c.contact_number", ContactNameItems())
    End Sub

    Private Sub GetFundraisingRequestTargets(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "target_amount,previous_target_amount,change_reason,changed_by,changed_on", "fundraising_request_targets", vWhereFields, "changed_on desc")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetFundRequestExpectedAmountHistory(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "fundraising_request_number,expected_amount,previous_expected_amount,change_reason,is_income_amount,changed_by,changed_on", "fund_expected_amount_history", vWhereFields, "changed_on desc")
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
        For Each vRow As CDBDataRow In pDataTable.Rows
          vRow.SetYNValue("IsIncomeAmount")
        Next
      End If
    End Sub
    Private Sub GetFundRequestStatusHistory(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("fundraising_statuses fs1", "frsh.fundraising_status", "fs1.fundraising_status")
        vAnsiJoins.Add("fundraising_statuses fs2", "previous_fundraising_status", "fs2.fundraising_status")
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "fundraising_request_number,frsh.fundraising_status,fs1.fundraising_status_desc,previous_fundraising_status,fs2.fundraising_status_desc AS previous_status_desc,change_reason,changed_by,changed_on", "fund_request_status_history frsh", vWhereFields, "changed_on desc", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      End If
    End Sub
    Private Sub GetFundraisingPaymentSchedule(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vAttrs As String = "fps.fundraising_request_number,fps.scheduled_payment_number,scheduled_payment_desc,payment_amount,due_date,fps.fundraising_payment_type,fpt.fundraising_payment_type_desc,fps.fund_income_payment_type,fipt.fund_income_payment_type_desc,fps.received_amount,fps.received_date,fps.source,s.source_desc,fps.notes,fps.created_by,fps.created_on,fps.amended_by,fps.amended_on,fa.action_number"
        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("FundraisingRequestNumber") Then vWhereFields.Add("fps.fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)
        If mvParameters.Exists("ScheduledPaymentNumber") Then vWhereFields.Add("fps.scheduled_payment_number", mvParameters("ScheduledPaymentNumber").IntegerValue)

        Dim vUseRestriction As Boolean = mvParameters.ParameterExists("TraderApplication").LongValue > 0 OrElse mvParameters.ParameterExists("BatchNumber").LongValue > 0
        If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("fr.contact_number", mvParameters("ContactNumber").LongValue)
        If mvParameters.ParameterExists("BatchNumber").LongValue > 0 Then
          Dim vBatch As New Batch(mvEnv)
          vBatch.Init(mvParameters("BatchNumber").LongValue)
          If vBatch.Existing AndAlso Not (vBatch.BatchType = Batch.BatchTypes.GiftInKind) Then
            If Not mvParameters.Exists("FundraisingPaymentType") Then mvParameters.Add("FundraisingPaymentType")
            mvParameters("FundraisingPaymentType").Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType)
          ElseIf mvParameters.Exists("FundraisingPaymentType") Then
            mvParameters.Remove("FundraisingPaymentType")
          End If
        End If
        If mvParameters.Exists("FundraisingPaymentType") Then
          vWhereFields.Add("fps.fundraising_payment_type", mvParameters("FundraisingPaymentType").Value)
        ElseIf vUseRestriction Then
          'Show all payment schedule records having payment type other than the one defined in the control value
          vWhereFields.Add("fps.fundraising_payment_type", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType), CDBField.FieldWhereOperators.fwoNotEqual)
        End If
        If vUseRestriction Then
          vWhereFields.Add("fr.fundraising_status", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundStatus))
        End If

        Dim vSubAnsiJoins As New AnsiJoins
        vWhereFields.Add("a.completed_on", CDBField.FieldTypes.cftDate, "")
        vSubAnsiJoins.Add("fundraising_requests fr", "fa.fundraising_request_number", "fr.fundraising_request_number")
        vSubAnsiJoins.Add("fundraising_payment_schedule fps", "fa.scheduled_payment_number", "fps.scheduled_payment_number", "fa.fundraising_request_number", "fps.fundraising_request_number", "fr.fundraising_request_number", "fps.fundraising_request_number")
        vSubAnsiJoins.Add("actions a", "fa.action_number", "a.action_number")
        Dim vSubSQLStatement As New SQLStatement(mvEnv.Connection, "fa.scheduled_payment_number,MAX(fa.action_number) AS action_number", "fundraising_actions fa", vWhereFields, "", vSubAnsiJoins)
        vSubSQLStatement.GroupBy = "fa.scheduled_payment_number"
        Dim vSubSQL As String = vSubSQLStatement.SQL
        vWhereFields.Remove("a.completed_on")

        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("fundraising_requests fr", "fps.fundraising_request_number", "fr.fundraising_request_number")
        vAnsiJoins.Add("fundraising_payment_types fpt", "fps.fundraising_payment_type", "fpt.fundraising_payment_type")
        vAnsiJoins.AddLeftOuterJoin("fund_income_payment_types fipt", "fps.fund_income_payment_type", "fipt.fund_income_payment_type")
        vAnsiJoins.AddLeftOuterJoin("sources s", "fps.source", "s.source")
        vAnsiJoins.AddLeftOuterJoin("(" & vSubSQL & ") fa", "fps.scheduled_payment_number", "fa.scheduled_payment_number")

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "fundraising_payment_schedule fps", vWhereFields, "fps.due_date,fps.scheduled_payment_number", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, "")
        pDataTable.Columns("HasAction").FieldType = CDBField.FieldTypes.cftCharacter
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("HasAction").Length > 0 Then
            vRow.Item("HasAction") = ProjectText.String15904
          End If
        Next
      End If
    End Sub

    Private Sub GetFundraisingRequests(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "frs.fundraising_request_number, frs.contact_number, frs.request_date, frs.request_description, frs.fundraising_request_stage, frs.fundraising_status, frs.fundraising_request_type,frs.target_amount,frs.pledged_amount,frs.pledged_date,frs.received_amount,frs.received_date,frs.expected_amount,frs.gik_expected_amount,frs.gik_pledged_amount,frs.gik_pledged_date,frs.total_gik_received_amount,frs.latest_gik_received_date,frs.number_of_payments,frs.request_end_date,frs.fundraising_business_type"

      Dim vWhereFields As New CDBFields()
      AddWhereFieldFromParameter(vWhereFields, "FundraisingRequestNumber", "frs.fundraising_request_number")
      AddWhereFieldFromParameter(vWhereFields, "ContactNumber", "frs.contact_number")
      AddWhereFieldFromParameter(vWhereFields, "RequestDescription", "frs.request_description")
      AddWhereFieldFromParameter(vWhereFields, "FundraisingRequestStage", "frs.fundraising_request_stage")
      AddWhereFieldFromParameter(vWhereFields, "FundraisingStatus", "frs.fundraising_status")
      AddWhereFieldFromParameter(vWhereFields, "FundraisingRequestType", "frs.fundraising_request_type")
      AddWhereFieldFromParameter(vWhereFields, "FundraisingBusinessType", "frs.fundraising_business_type")
      If mvParameters.Exists("RequestDate") Then vWhereFields.Add("frs.request_date", CDBField.FieldTypes.cftDate, mvParameters("RequestDate").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If mvParameters.Exists("RequestEndDate") Then vWhereFields.Add("frs.request_date#2", CDBField.FieldTypes.cftDate, mvParameters("RequestEndDate").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "fundraising_requests frs", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, "")

    End Sub

    Private Sub GetFundraisingPaymentHistory(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vAttrs As String = "fps.scheduled_payment_number,bta.batch_number,bta.transaction_number,bta.line_number,bt.transaction_date,bta.product,product_desc,bta.rate,rate_desc,bt.transaction_type,transaction_type_desc,payment_method,bta.distribution_code,quantity,bta.amount,vat_amount,bta.source,bta.currency_amount,currency_vat_amount,bta.notes,tt.transaction_sign"
        Dim vItems As String = vAttrs
        Dim vContact As New Contact(mvEnv)
        vAttrs &= "," & vContact.GetRecordSetFieldsName
        vItems &= ",c.contact_number,CONTACT_NAME"
        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("ScheduledPaymentNumber") Then vWhereFields.Add("fps.scheduled_payment_number", mvParameters("ScheduledPaymentNumber").LongValue)
        If mvParameters.Exists("FundraisingRequestNumber") Then vWhereFields.Add("fps.fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)

        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("fundraising_payment_schedule fps", "fps.scheduled_payment_number", "fph.scheduled_payment_number")
        vAnsiJoins.Add("fundraising_requests fr", "fr.fundraising_request_number", "fps.fundraising_request_number")
        If Not mvParameters.Exists("ContactNumber") Then
          vAnsiJoins.Add("batch_transaction_analysis bta", "bta.batch_number", "fph.batch_number", "bta.transaction_number", "fph.transaction_number", "bta.line_number", "fph.line_number")
          vAnsiJoins.Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
          vAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
          vAnsiJoins.Add("products p", "bta.product", "p.product")
          vAnsiJoins.Add("rates r", "bta.rate", "r.rate", "p.product", "r.product")
          vAnsiJoins.Add("contacts c", "bt.contact_number", "c.contact_number")
        End If
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "fundraising_payment_history fph", vWhereFields, "fps.scheduled_payment_number", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vItems, "")
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("TransactionSign") = "D" Then
            If vRow.Item("Amount").Length > 0 Then vRow.Item("Amount") = (CDbl(vRow.Item("Amount")) * -1).ToString
            vRow.Item("Quantity") = (IntegerValue(vRow.Item("Quantity")) * -1).ToString
            If vRow.Item("CurrencyAmount").Length > 0 Then vRow.Item("CurrencyAmount") = (CDbl(vRow.Item("CurrencyAmount")) * -1).ToString
          End If
        Next
      End If
    End Sub
    Private Sub GetFundraisingActions(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
        Dim vAttrs As String = "fa.fundraising_request_number,fa.scheduled_payment_number,a.master_action,action_level,sequence_number,fa.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,duration_days,duration_hours,duration_minutes,a.document_class,action_text"
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("fa.fundraising_request_number", mvParameters("FundraisingRequestNumber").LongValue)
        If mvParameters.Exists("ScheduledPaymentNumber") Then vWhereFields.Add("fa.scheduled_payment_number", mvParameters("ScheduledPaymentNumber").LongValue)
        vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
        vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)

        vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("department", mvEnv.User.Department)
        vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)

        vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)

        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("actions a", "fa.action_number", "a.action_number")
        vAnsiJoins.Add("users u", "a.created_by", "u.logname")
        vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
        vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
        vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")

        Dim vOrderBy As New StringBuilder
        With vOrderBy
          .Append(mvEnv.Connection.DBOrderByNullsFirstDesc("a.completed_on"))
          .Append(",")
          .Append(mvEnv.Connection.DBIsNull("fa.scheduled_payment_number - fa.scheduled_payment_number", "1")) 'This is a workaround to sort payment schedule number with values of 0 and 1 to get the correct scheduled_on order
          .Append(" DESC,")
          .Append(mvEnv.Connection.DBOrderByNullsFirstAsc("a.scheduled_on"))
          .Append(",a.action_number")
        End With
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "fundraising_actions fa", vWhereFields, vOrderBy.ToString, vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
      End If
    End Sub
    Private Sub GetGeographicalRegions(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins()
      Dim vTables As String
      Dim vSelAttrs As String = "gr.geographical_region,geographical_region_desc"
      'OrganisationNumber could be null in which case GeographicalRegionDesc will be set
      Dim vAttrs As String = String.Format("gr.geographical_region,{0}AS geographical_region_desc", mvEnv.Connection.DBIsNull("name", "geographical_region_desc"))
      If mvParameters.Exists("CollectionNumber") Then
        vTables = "collection_regions cr"
        vAnsiJoins.Add("geographical_regions gr", "cr.geographical_region", "gr.geographical_region")
        vWhereFields.Add("cr.collection_number", mvParameters.Item("CollectionNumber").LongValue)
        vAttrs = vAttrs & ",collection_region_number,collection_number"
        vSelAttrs = vSelAttrs & ",collection_region_number,collection_number"
      Else
        vTables = "geographical_regions gr"
        vSelAttrs = vSelAttrs & ",,"
      End If
      vAnsiJoins.AddLeftOuterJoin("organisations o", "gr.organisation_number", "o.organisation_number")
      If mvParameters.Exists("GeographicalRegionType") Then vWhereFields.Add("geographical_region_type", mvParameters.Item("GeographicalRegionType").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTables, vWhereFields, "geographical_region_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vSelAttrs)
    End Sub
    Private Sub GetH2HCollectors(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetH2HCollectors")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry Or Contact.ContactRecordSetTypes.crtPhone), "c.contact_number,", "")
      Dim vAttrs As String = "ready_for_confirmation,collection_number,collector_number,hc.contact_number,route,hc.route_type,route_type_desc,no_of_premises,operator_contact_number,hc.collector_status,collector_status_desc,hc.notes,hc.confirmation_produced_on,hc.reminder_produced_on,hc.amended_by,hc.amended_on"
      With vWhereFields
        If mvParameters.Exists("CollectorNumber") Then .Add("collector_number", mvParameters("CollectorNumber").LongValue)
        If mvParameters.Exists("ContactNumber") Then .Add("contact_number", mvParameters("ContactNumber").LongValue)
        If mvParameters.Exists("CollectionNumber") Then .Add("collection_number", mvParameters("CollectionNumber").LongValue)
        .AddJoin("hc.contact_number", "c.contact_number")
        .AddJoin("c.address_number", "a.address_number")
        .AddJoin("a.country", "co.country")
        .AddJoin("hc.route_type", "rt.route_type")
        .AddJoin("hc.collector_status", "cs.collector_status")
      End With
      Dim vSQL As String = "SELECT " & vConAttrs & "," & vAttrs & " FROM h2h_collectors hc, contacts c, addresses a, countries co, route_types rt, collector_statuses cs WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, ContactNameItems() & ",ADDRESS_LINE,CONTACT_TELEPHONE," & vAttrs & ",")
      GetContactNames(pDataTable, "OperatorContactNumber", "OperatorContactName")
    End Sub
    Private Sub GetIncentives(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetIncentives")
      Dim vAttrs As String = "isp.incentive_scheme,isp.reason_for_despatch,isp.sequence_number,isp.product,isp.rate,isp.for_whom,isp.incentive_type,isp.incentive_desc,isp.quantity,isp.basic,despatch_method,current_price,future_price,price_change_date,vri.vat_rate,percentage,product_desc,subscription,thank_you_letter,vat_exclusive"
      Dim vSQL As String = "SELECT " & vAttrs
      vSQL = vSQL & ",ignore_product_and_rate"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIncentiveProductMinMax) Then vSQL = vSQL & ",isp.minimum_quantity,isp.maximum_quantity"
      vSQL = vSQL & " FROM sources s, incentive_scheme_products isp, products p, vat_rate_identification vri, vat_rates vr, rates r "
      vSQL = vSQL & " WHERE s.source = '" & mvParameters("Source").Value & "'"
      If mvParameters.Exists("Amount") Then vSQL = vSQL & " AND s.incentive_trigger_level <= " & mvParameters("Amount").Value
      vSQL = vSQL & " AND s.incentive_scheme = isp.incentive_scheme "
      If Not mvParameters.Exists("PayMethodReason") Then
        vSQL = vSQL & "AND isp.reason_for_despatch = '" & mvParameters("ReasonForDespatch").Value & "'"
      Else
        vSQL = vSQL & "AND isp.reason_for_despatch IN( '" & mvParameters("ReasonForDespatch").Value & "','" & mvParameters("PayMethodReason").Value & "')"
      End If
      If mvParameters.Exists("Type") Then vSQL = vSQL & " AND isp.incentive_type = '" & mvParameters("Type").Value & "'"
      If mvParameters.Exists("Basic") Then vSQL = vSQL & " AND isp.basic = '" & mvParameters("Basic").Value & "' "
      vSQL = vSQL & " AND isp.product = p.product"
      vSQL = vSQL & " AND p.product_vat_category = vri.product_vat_category"
      vSQL = vSQL & " AND vri.contact_vat_category = '" & mvParameters("VatCategory").Value & "'"
      vSQL = vSQL & " AND vri.vat_rate = vr.vat_rate AND p.product = r.product AND r.rate = isp.rate"
      vSQL = vSQL & " ORDER BY isp.reason_for_despatch, basic DESC, sequence_number"
      vAttrs = vAttrs & ",quantity,ignore_product_and_rate"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIncentiveProductMinMax) Then
        vAttrs = vAttrs & ",minimum_quantity,maximum_quantity"
      Else
        vAttrs = vAttrs & ",,"
      End If
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      If mvParameters.Exists("ExtraSource") Then
        vSQL = Replace$(vSQL, "s.source = '" & mvParameters("Source").Value & "'", "s.source = '" & mvParameters("ExtraSource").Value & "'")
        If mvParameters.Exists("ExtraReason") Then
          vSQL = Replace$(vSQL, "isp.reason_for_despatch = '" & mvParameters("ReasonForDespatch").Value & "'", "isp.reason_for_despatch = '" & mvParameters("ExtraReason").Value & "'")
        End If
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      End If
    End Sub
    Private Sub GetFulFilledContactIncentives(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "c.contact_number,c.label_name,product_desc,quantity,source_desc,cir.date_responded,cir.date_fulfilled"
      Dim vTableName As String = "contacts c, contact_incentive_responses cir, contact_incentives ci, sources s, products p"
      Dim vOrderBy As String = "product_desc, cir.date_responded, cir.date_fulfilled"
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "ContactNumber", "c.contact_number")

      vWhereFields.Add("cir.date_fulfilled", "", CDBField.FieldWhereOperators.fwoNOT) ' TODO: need testing
      vWhereFields.AddJoin("cir.contact_number", "c.contact_number")
      vWhereFields.AddJoin("ci.contact_number", "cir.contact_number")
      vWhereFields.AddJoin("ci.source", "cir.source")
      vWhereFields.AddJoin("p.product", "ci.product")
      vWhereFields.AddJoin("s.source", "ci.source")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFieldNames, vTableName, vWhereFields, vOrderBy)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub

    Private Sub GetUnFulFilledContactIncentives(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "c.contact_number,c.label_name,product_desc,quantity,source_desc,cir.date_responded"
      Dim vTableName As String = "contacts c, contact_incentive_responses cir, contact_incentives ci, products p, sources s"
      Dim vOrderBy As String = "product_desc, date_responded"
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "ContactNumber", "c.contact_number")
      AddWhereFieldFromParameter(vWhereFields, "Source", "cir.source")

      vWhereFields.Add("cir.date_fulfilled", "") ' TODO: need testing
      vWhereFields.AddJoin("cir.contact_number", "c.contact_number")
      vWhereFields.AddJoin("ci.contact_number", "cir.contact_number")
      vWhereFields.AddJoin("ci.source", "cir.source")
      vWhereFields.AddJoin("p.product", "ci.product")
      vWhereFields.AddJoin("s.source", "ci.source")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFieldNames, vTableName, vWhereFields, vOrderBy)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetUnFulFilledPayPlanIncentives(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "e.contact_number,c.label_name,product_desc,quantity,source_desc,reason_for_despatch_desc,date_created"
      Dim vTableName As String = "orders o, enclosures e, contacts c, products p, sources s, reasons_for_despatch rfd"
      Dim vOrderBy As String = "product_desc, date_created"
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "PaymentPlanNumber", "o.order_number")
      AddWhereFieldFromParameter(vWhereFields, "Source", "s.source")
      AddWhereFieldFromParameter(vWhereFields, "ReasonForDespatch", "rfd.reason_for_despatch")

      vWhereFields.Add("e.date_fulfilled", "")
      vWhereFields.AddJoin("e.order_number", "o.order_number")
      vWhereFields.AddJoin("c.contact_number", "e.contact_number")
      vWhereFields.AddJoin("p.product", "e.product")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFieldNames, vTableName, vWhereFields, vOrderBy)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetFulFilledPayPlanIncentives(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "e.contact_number,c.label_name,product_desc,quantity,source_desc,reason_for_despatch_desc,date_created, date_fulfilled"
      Dim vTableName As String = "orders o, sources s, reasons_for_despatch rfd, enclosures e, contacts c, products p"
      Dim vOrderBy As String = "product_desc, e.date_created, e.date_fulfilled"
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "PaymentPlanNumber", "o.order_number")
      'AddWhereFieldFromParameter(vWhereFields, "Source", "s.source")
      'AddWhereFieldFromParameter(vWhereFields, "ReasonForDispatch", "rfd.reason_for_despatch")
      vWhereFields.Add("e.date_fulfilled", "", CDBField.FieldWhereOperators.fwoNOT)
      vWhereFields.AddJoin("s.source", "o.source")
      vWhereFields.AddJoin("rfd.reason_for_despatch", "o.reason_for_despatch")
      vWhereFields.AddJoin("e.order_number", "o.order_number")
      vWhereFields.AddJoin("c.contact_number", "e.contact_number")
      vWhereFields.AddJoin("p.product", "e.product")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFieldNames, vTableName, vWhereFields, vOrderBy)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
    End Sub
    Private Sub GetMannedCollectionBoxes(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetMannedCollectionBoxes")
      Dim vWhereFields As New CDBFields
      Dim vWhereFields2 As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry Or Contact.ContactRecordSetTypes.crtPhone), "c.contact_number,", "")
      Dim vAttrs As String = "cb.collection_box_number,cb.collection_number,cb.box_reference,cb.collector_number,x.contact_number,cb.amount,cb.collection_pis_number,cp.pis_number,cb.amended_by,cb.amended_on,sum_pay_amount"
      With vWhereFields
        If mvParameters.Exists("CollectionNumber") Then
          .Add("cb.collection_number", mvParameters("CollectionNumber").LongValue)
          vWhereFields2.Add("colb.collection_number", mvParameters("CollectionNumber").LongValue)
        End If
        If mvParameters.Exists("CollectionBoxNumber") Then
          .Add("cb.collection_box_number", mvParameters("CollectionBoxNumber").LongValue)
          vWhereFields2.Add("colb.collection_box_number", mvParameters("CollectionBoxNumber").LongValue)
        End If
        If mvParameters.Exists("CollectionPISNumber") Then .Add("cb.collection_pis_number", mvParameters("CollectionPISNumber").LongValue)
        If mvType <> DataSelectionTypes.dstMannedCollectionBoxes Then
          If mvParameters.Exists("CollectorNumber") Then
            .Add("cb.collector_number", mvParameters("CollectorNumber").LongValue)
          ElseIf mvParameters.Exists("ContactNumber") Then
            'contact number should have only been passed for manned collectors
            .Add("cb.collector_number", CDBField.FieldTypes.cftInteger, "SELECT collector_number FROM manned_collectors WHERE contact_number = " & mvParameters("ContactNumber").Value, CDBField.FieldWhereOperators.fwoIn)
          End If
        End If
      End With
      Dim vSQL As String = "SELECT " & Replace$(Replace$(vConAttrs, ",c.", ",x."), ",a.", ",x.") & "," & vAttrs
      vSQL = vSQL & " FROM collection_boxes cb LEFT OUTER JOIN (SELECT collector_number,mc.contact_number," & vConAttrs & " FROM manned_collectors mc INNER JOIN contacts c ON mc.contact_number = c.contact_number INNER JOIN addresses a ON c.address_number = a.address_number INNER JOIN countries co ON a.country = co. country) x ON cb.collector_number = x.collector_number"
      vSQL = vSQL & " LEFT OUTER JOIN collection_pis cp ON cb.collection_pis_number = cp.collection_pis_number "
      vSQL = vSQL & " LEFT OUTER JOIN (SELECT SUM(colp.amount) AS sum_pay_amount, colp.collection_number, colp.collection_box_number FROM collection_payments colp INNER JOIN collection_boxes colb ON colp.collection_box_number = colb.collection_box_number"
      If vWhereFields2.Count > 0 Then vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields2)
      vSQL = vSQL & " GROUP BY colp.collection_number, colp.collection_box_number) z ON cb.collection_number = z.collection_number AND cb.collection_box_number = z.collection_box_number"
      If vWhereFields.Count > 0 Then vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), ContactNameItems() & ",ADDRESS_LINE,CONTACT_TELEPHONE," & vAttrs)
      '  Case dstPayingInSlips
      '    vAttrs = "pis.issue_date,pis.collection_number,pis.paying_in_slip_number,pis.amount,pis.banked_by,pis.banked_on,pis.reconciled_on"
      '    If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add( "pis.collection_number",  mvParameters("CollectionNumber").LongValue)
      '    If mvParameters.Exists("PayingInSlipNumber") Then vWhereFields.Add( "pis.paying_in_slip_number",  mvParameters("PayingInSlipNumber").Value)
      '    vSQL = "SELECT " & vAttrs & " FROM paying_in_slips pis WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      '    pDataTable.FillFromSQLDONOTUSE( mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetMannedCollectors(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetMannedCollectors")
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "mc.collector_number,mc.collection_number,mc.contact_number,mc.total_time,mc.attended,mc.ready_for_confirmation,ready_for_acknowledgement,confirmation_produced_on,acknowledgement_produced_on,mc.notes,mc.amended_by,mc.amended_on"
      If mvParameters.Exists("CollectionNumber") Then vWhereFields.Add("mc.collection_number", mvParameters("CollectionNumber").LongValue)
      If mvParameters.Exists("CollectorNumber") Then vWhereFields.Add("mc.collector_number", mvParameters("CollectorNumber").LongValue)
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("mc.contact_number", mvParameters("ContactNumber").LongValue)
      vWhereFields.AddJoin("c.contact_number", "mc.contact_number")
      Dim vSQL As String = "SELECT " & vConAttrs & "," & vAttrs & " FROM manned_collectors mc, contacts c WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, ContactNameItems() & "," & vAttrs)
    End Sub
    Private Sub GetMeetingContactLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      'TODO Convert to new SQL Syntax
      'NYI("GetMeetingContactLinks")
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = "link_type,contact_number,CONTACT_NAME,notified,attended,meeting_role,meeting_role"
      Dim vSQL As String = "SELECT link_type," & vContact.GetRecordSetFieldsName & ",notified,attended,meeting_role"
      vSQL = vSQL & " FROM meeting_links t1, contacts c "
      vSQL = vSQL & " WHERE t1.meeting_number = " & mvParameters("MeetingNumber").LongValue & " AND t1.contact_number = c.contact_number AND c.contact_type <> 'O' ORDER BY surname, initials"
      Dim vAddItems As String = ""
      If pAddType Then vAddItems = "CONTACT_TYPE_1,MEETING_LINK_TYPE"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, vAddItems)
      SetMeetingRoleDesc(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified")
        vRow.SetAttended("Attended")
      Next
    End Sub
    Private Sub GetMeetingDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      GetMeetingDocumentLinks(pDataTable, pAddType, False)
    End Sub
    Private Sub GetMeetingDocumentLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean, pIncludeEmail As Boolean)
      'TODO Convert to new SQL Syntax
      'NYI("GetMeetingDocumentLinks")
      Dim vCommsLog As New CommunicationsLog(mvEnv)
      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("communications_log cl", "cl.communications_log_number", "t1.communications_log_number")})
      Dim vWhereFields As New CDBFields({New CDBField("t1.meeting_number", mvParameters("MeetingNumber").LongValue, CDBField.FieldWhereOperators.fwoEqual)})
      If Not pIncludeEmail Then
        vAnsiJoins.Add("packages pk", "cl.package", "pk.package", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("pk.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
        vAnsiJoins.Add("document_types dt", "cl.document_type", "dt.document_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("dt.document_source", "E", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If
      If pAddType Then
        Dim vSQL As New SQLStatement(mvEnv.Connection, "'R' AS link_type," & vCommsLog.GetRecordSetFieldsDetail, "meeting_documents t1", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, "link_type,communications_log_number,DOCUMENT_NAME,,,,,,MEETING_LINK_TYPE")
      Else
        Dim vSQL As New SQLStatement(mvEnv.Connection, "cl.communications_log_number, our_reference", "meeting_documents t1", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub
    Private Sub GetMeetingLinks(ByVal pDataTable As CDBDataTable)
      GetMeetingLinks(pDataTable, False)
    End Sub
    Private Sub GetMeetingLinks(ByVal pDataTable As CDBDataTable, pIncludeEmail As Boolean)
      'TODO Convert to new SQL Syntax
      'NYI("GetMeetingLinks")
      GetMeetingContactLinks(pDataTable, True)
      GetMeetingOrganisationLinks(pDataTable, True)
      GetMeetingDocumentLinks(pDataTable, True, pIncludeEmail)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("LinkType") = "W" Then vRow.Item("LinkType") = "A"
      Next
      pDataTable.ReOrderRowsByColumn(("LinkType"))
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("LinkType") = "A" Then vRow.Item("LinkType") = "W"
      Next
    End Sub
    Private Sub GetMeetingOrganisationLinks(ByVal pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      'TODO Convert to new SQL Syntax
      'NYI("GetMeetingOrganisationLinks")
      Dim vAttrs As String = "link_type,o.organisation_number,name,notified,attended,meeting_role"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM meeting_links t1, organisations o "
      vSQL = vSQL & " WHERE t1.meeting_number = " & mvParameters("MeetingNumber").LongValue & " AND t1.contact_number = o.organisation_number ORDER BY name"
      Dim vAddItems As String = ""
      If pAddType Then vAddItems = "ORGANISATION_TYPE_1,MEETING_LINK_TYPE"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs & ",meeting_role", vAddItems)
      SetMeetingRoleDesc(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Notified")
        vRow.SetAttended("Attended")
      Next
    End Sub
    Private Sub GetMembershipChanges(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetMembershipChanges")
      Dim vAttrs As String = "m.contact_number,m.address_number,m.member_number,m.membership_type,membership_type_desc,m.joined,name,m.cancelled_on,m.cancelled_by,m.cancellation_reason,m.cancellation_source,m.amended_on,m.amended_by,m.membership_number,m.source,source_desc"
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs) & " FROM members m1, members m, membership_types mt, branches b, organisations o, sources s WHERE m1.membership_number = " & mvParameters("MembershipNumber").Value & " AND m1.order_number = m.order_number AND m.membership_number <> m1.membership_number AND m.contact_number = m1.contact_number AND m.membership_type = mt.membership_type AND m.branch = b.branch AND b.organisation_number = o.organisation_number AND m.source = s.source ORDER BY m.cancelled_on DESC, m.joined DESC "
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetMembershipGroupHistory(ByVal pDataTable As CDBDataTable)
      '"HistoryNumber,MembershipNumber,OldGroupName,NewGroupName,ChangeDate"
      Dim vAttrs As String = "member_group_history_number,mgh.membership_number,org1.name AS old_group_name,org2.name AS new_group_name,change_date"
      Dim vWhereFields As New CDBFields

      With vWhereFields
        .Add("m.membership_number", CDBField.FieldTypes.cftLong, mvParameters("MembershipNumber").Value)
        .AddJoin("m.membership_number#2", "mgh.membership_number")
        .AddJoin("mgh.old_organisation_number", "org1.organisation_number")
        .AddJoin("mgh.new_organisation_number", "org2.organisation_number")
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "members m, membership_group_history mgh, organisations org1, organisations org2", vWhereFields, "change_date DESC,member_group_history_number")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)

    End Sub
    Private Sub GetMembershipGroups(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "mg.membership_group_number,mg.membership_number,mg.organisation_number,name,default_group,valid_from,valid_to,is_current"
      Dim vOrderby As String = String.Format("valid_from DESC, valid_to {0}", IIf(mvEnv.Connection.NullsSortAtEnd = True, "DESC", ""))
      Dim vWhereFields As New CDBFields

      With vWhereFields
        If mvParameters.Exists("MembershipGroupNumber") Then .Add("mg.membership_group_number", CDBField.FieldTypes.cftLong, mvParameters("MembershipGroupNumber").Value)
        .Add("m.membership_number", CDBField.FieldTypes.cftLong, mvParameters("MembershipNumber").Value)
        .AddJoin("m.membership_number#2", "mg.membership_number")
        'If mvParameters.Exists("OrganisationNumber") Then .Add("mg.organisation_number", mvParameters("OrganisationNumber").Value)
        .AddJoin("mg.organisation_number", "org.organisation_number")
        '.Add("", "", CType(CDBField.FieldWhereOperators.fwoBetweenFrom + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "members m, membership_groups mg, organisations org", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("DefaultGroup")
        vRow.SetYNValue("IsCurrent")
      Next

    End Sub
    Private Sub GetMembershipOtherMembers(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetMembershipOtherMembers")
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "m.contact_number,m.address_number,m.member_number,m.membership_type,membership_type_desc,m.joined,name,m.cancelled_on,m.cancelled_by,m.cancellation_reason,m.cancellation_source,m.membership_card_expires,m.amended_on,m.amended_by,m.membership_number"
      Dim vSQL As String = "SELECT " & vAttrs & "," & vConAttrs & " FROM members m1, members m, contacts c, membership_types mt, branches b, organisations o WHERE m1.membership_number = " & mvParameters("MembershipNumber").Value & " AND m1.order_number =  m.order_number AND m.contact_number <> m1.contact_number AND m.contact_number = c.contact_number AND m.membership_type = mt.membership_type AND m.branch = b.branch AND b.organisation_number = o.organisation_number ORDER BY m.cancelled_on" & mvEnv.Connection.DBSortByNullsFirst & ", c.surname"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ContactNameItems() & ",,")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetMembershipPaymentPlanDetails(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "od.contact_number,od.address_number,p.product,product_desc,od.rate,rate_desc,od.distribution_code,quantity,amount,balance,arrears,od.despatch_method,od.source,product_number,od.amended_on,od.amended_by,currency_code,distribution_code_desc,vat_exclusive"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
        vAttrs &= ",valid_from,valid_to,modifier_activity,activity_desc,modifier_activity_value,activity_value_desc,modifier_activity_quantity,modifier_activity_date,modifier_price,modifier_per_item,unit_price,pro_rated,net_amount,vat_amount,gross_amount,vat_rate,vat_percentage"
      Else
        vAttrs &= ",,,,,,,,,,,,,,,,,"
      End If
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs = vAttrs.Replace("currency_code", "")

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("order_details od", "m.order_number", "od.order_number")
        .Add("products p", "od.product", "p.product")
        .Add("rates r", "od.product", "r.product", "od.rate", "r.rate")
        .Add("contacts c", "od.contact_number", "c.contact_number")
        .AddLeftOuterJoin("distribution_codes dc", "od.distribution_code", "dc.distribution_code")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
          .AddLeftOuterJoin("activities a", "od.modifier_activity", "a.activity")
          .AddLeftOuterJoin("activity_values av", "od.modifier_activity", "av.activity", "od.modifier_activity_value", "av.activity_value")
        End If
      End With
      Dim vWhereFields As New CDBFields(New CDBField("membership_number", mvParameters("MembershipNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "members m", vWhereFields, "detail_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",,,,,,,,,,")

      'If multiple RateModifiers were used then set the data to "Multiple"
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ModifierActivity").ToUpper = "MULTI" Then
          vRow.Item("ModifierActivity") = "Multiple"
          vRow.Item("ModifierActivityValue") = "Multiple"
          vRow.Item("ModifierActivityQuantity") = ""
          vRow.Item("ModifierActivityDate") = ""
          vRow.Item("ModifierPrice") = ""
          vRow.Item("ModifierPerItem") = "Multiple"
          vRow.Item("ModifierActivityDesc") = "Multiple"
          vRow.Item("ModifierActivityValueDesc") = "Multiple"
        End If
      Next

      'Now get the subs data
      vAttrs = "s.contact_number,s.address_number,s.product,s.quantity,s.despatch_method,valid_from,valid_to,s.cancelled_by,s.cancelled_on,s.cancellation_reason,s.cancellation_source,s.reason_for_despatch,cancellation_reason_desc,source_desc"
      With vAnsiJoins
        .Clear()
        .Add("subscriptions s", "m.order_number", "s.order_number")
        .AddLeftOuterJoin("cancellation_reasons cr", "s.cancellation_reason", "cr.cancellation_reason")
        .AddLeftOuterJoin("sources crs", "s.cancellation_source", "crs.source")
      End With

      vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "members m", vWhereFields, "", vAnsiJoins)
      Dim vRecordset As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRecordset.Fetch()
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRecordset.Fields(1).Value = vRow.Item("ContactNumber") And vRecordset.Fields(2).Value = vRow.Item("AddressNumber") And vRecordset.Fields(3).Value = vRow.Item("Product") And vRecordset.Fields(4).Value = vRow.Item("Quantity") And vRecordset.Fields(5).Value = vRow.Item("DespatchMethod") Then
            vRow.Item("SubsValidFrom") = vRecordset.Fields(6).Value
            vRow.Item("SubsValidTo") = vRecordset.Fields(7).Value
            vRow.Item("CancelledBy") = vRecordset.Fields(8).Value
            vRow.Item("CancelledOn") = vRecordset.Fields(9).Value
            vRow.Item("CancellationReason") = vRecordset.Fields(10).Value
            vRow.Item("CancellationSource") = vRecordset.Fields(11).Value
            vRow.Item("CancellationReasonDesc") = vRecordset.Fields(13).Value
            vRow.Item("CancellationSourceDesc") = vRecordset.Fields(14).Value
            Exit For
          End If
        Next
      End While
      vRecordset.CloseRecordSet()
    End Sub
    Private Sub GetMembershipSummaryMembers(ByVal pDataTable As CDBDataTable)
      'Build DataTable containing members when creating Joint membership etc.
      'Does NOT currently support Associate MembershipTypes
      Dim vMembershipType As String = mvParameters("MembershipType").Value
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = "a.address_number," & vContact.GetRecordSetFieldsName & ",date_of_birth"

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contact_addresses ca", "a.address_number", "ca.address_number")
      vAnsiJoins.Add("contacts c", "ca.contact_number", "c.contact_number")

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("a.address_number", mvParameters("AddressNumber").LongValue)
      vWhereFields.Add("c.contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters("GiftMembership").Bool Then vWhereFields.Add("c.contact_number#2", mvParameters("PayerContactNumber").LongValue, CDBField.FieldWhereOperators.fwoNotEqual)

      Dim vOrderBy As String = "date_of_birth"
      If mvEnv.Connection.NullsSortAtEnd Then vOrderBy &= " DESC"
      vOrderBy &= ",surname,forenames"

      'First add current member
      'Contact type and historic addresses are ignored for the first Member
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "addresses a", vWhereFields, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet()
      If vRS.Fetch = False Then RaiseError(DataAccessErrors.daeParameterNotFound, "ContactNumber")
      ' If Len(vMembershipType.AssociateMembershipType) > 0 Then
      '    vDOB = vContact.DateOfBirth
      '    If Len(vDOB) = 0 Then vDOB = pParams("DateOfBirth").Value
      '    If IsDate(vDOB) And vAssocMemberType.MaxJuniorAge > 0 Then
      '      'Check whether DateOfBirth makes this an associate member
      '      If DateAdd("yyyy", vAssocMemberType.MaxJuniorAge, vDOB) >= CDate(gvSystem.TodaysDate) Then vGotAssociate = True
      '    End If
      '  End If
      pDataTable.AddRowFromItems("1", vRS.Fields("address_number").Value, vRS.Fields("surname").Value, vRS.Fields("contact_number").Value, vMembershipType, vRS.Fields("label_name").Value, mvParameters("Joined").Value, mvParameters("Branch").Value, mvParameters("BranchMember").Value, mvParameters("Applied").Value, mvParameters("DistributionCode").Value, vRS.Fields("date_of_birth").Value)
      vRS.CloseRecordSet()

      'Second, add all related members
      vWhereFields(2).WhereOperator = CDBField.FieldWhereOperators.fwoNotEqual
      vWhereFields.Add("historical", "N")
      vWhereFields.Add("contact_type", "C")
      vSQL = New SQLStatement(mvEnv.Connection, vAttrs, "addresses a", vWhereFields, vOrderBy, vAnsiJoins)
      vRS = vSQL.GetRecordSet()
      Dim vSequence As Integer = 2
      While vRS.Fetch
        '  If Len(vMembershipType.AssociateMembershipType) > 0 Then
        '    If IsDate(vContact.DateOfBirth) And vAssocMemberType.MaxJuniorAge > 0 Then
        '      'Check whether DateOfBirth makes this an associate member
        '      If DateAdd("yyyy", vAssocMemberType.MaxJuniorAge, vContact.DateOfBirth) >= CDate(gvSystem.TodaysDate) Then vGotAssociate = True
        '    Else
        '      'No DOB so allocate membership type depending on whether all main members have been added
        '      If (vMemberCount >= pParams("NumberOfMembers").LongValue) And vAssocMemberType.MaxJuniorAge > 0 Then vGotAssociate = True
        '    End If
        '  End If
        pDataTable.AddRowFromItems(vSequence.ToString, vRS.Fields("address_number").Value, vRS.Fields("surname").Value, vRS.Fields("contact_number").Value, vMembershipType, vRS.Fields("label_name").Value, mvParameters("Joined").Value, mvParameters("Branch").Value, mvParameters("BranchMember").Value, mvParameters("Applied").Value, mvParameters("DistributionCode").Value, vRS.Fields("date_of_birth").Value)
        vSequence += 1
      End While
      vRS.CloseRecordSet()

    End Sub
    Private Sub GetOrganisationContactCommsNumbers(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetOrganisationContactCommsNumbers")
      Dim vAttrs As String = "co.contact_number,co.address_number,co.device,device_desc,co.dialling_code,co.std_code,extension,co.ex_directory,co.notes,co.amended_by,co.amended_on,valid_from,valid_to,is_active,co.mail,device_default,preferred_method,communication_number,telephone"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then vAttrs = Replace$(vAttrs, "communication_number", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then vAttrs = Replace$(vAttrs, "valid_from,valid_to,is_active,co.mail,device_default,preferred_method", ",,,,,")
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace$(RemoveBlankItems(vAttrs), "telephone", mvEnv.Connection.DBSpecialCol("", "number") & " AS telephone")
      Dim vAttr As String = ""
      If mvType = DataSelectionTypes.dstOrganisationContactCommsNumbers Then
        vSQL = vSQL & "," & Replace$(mvContact.GetRecordSetFieldsName, "c.contact_number,", "") & " FROM contact_positions cp, contacts c, communications co, devices d WHERE "
        vAttr = "," & ContactNameItems()
      Else
        vSQL = vSQL & " FROM communications co, devices d WHERE "
        If mvType = DataSelectionTypes.dstContactCommsNumbers Then vAttr = ",,"
      End If
      If mvParameters.Exists("CommunicationNumber") Then vSQL = vSQL & "co.communication_number" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftLong, mvParameters("CommunicationNumber").Value) & " AND "
      If mvParameters.Exists("Active") And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then vSQL = vSQL & "is_active " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, Left(mvParameters("Active").Value, 1)) & " AND "
      If mvParameters.Exists("AddressNumber") AndAlso mvType <> DataSelectionTypes.dstContactHeaderCommsNumbers Then vSQL = vSQL & "address_number " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftLong, mvParameters("AddressNumber").Value) & " AND "

      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        If mvType = DataSelectionTypes.dstOrganisationContactCommsNumbers Then
          vSQL = vSQL & "cp.organisation_number = " & mvContact.ContactNumber & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND c.contact_number = cp.contact_number AND contact_type <> 'O' AND co.contact_number = c.contact_number AND co.address_number = cp.address_number AND d.device = co.device"
        Else
          vSQL = vSQL & "address_number IN (select address_number FROM organisation_addresses WHERE organisation_number = " & mvContact.ContactNumber & ") AND contact_number IS NULL"
          vSQL = vSQL & " AND co.device = d.device"
        End If
      Else
        vSQL = vSQL & "contact_number = " & mvContact.ContactNumber
        vSQL = vSQL & " AND co.device = d.device"
      End If
      vSQL = vSQL & " ORDER BY "
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then vSQL = vSQL & "preferred_method DESC, device_default DESC, is_active DESC, "
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDevicesSequenceNumber) Then
        If mvType = DataSelectionTypes.dstOrganisationContactCommsNumbers Then
          vSQL = vSQL & "surname, d.sequence_number, device_desc"
        Else
          vSQL = vSQL & "d.sequence_number, device_desc"
        End If
      Else
        If mvType = DataSelectionTypes.dstOrganisationContactCommsNumbers Then
          vSQL = vSQL & "surname, device_desc"
        Else
          vSQL = vSQL & "device_desc"
        End If
      End If
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "CONTACT_TELEPHONE" & vAttr)
      If mvType = DataSelectionTypes.dstContactCommsNumbers Then GetAddressData(pDataTable)
      Dim vOrg As New Organisation(mvEnv)
      If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vOrg.Init(mvContact.ContactNumber)
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ExDirectory") = "Y" And mvContact.Department <> mvEnv.User.Department Then
          vRow.Item("PhoneNumber") = DataSelectionText.String23335    'Ex-Directory
        Else
          If Len(vRow.Item("Extension")) > 0 Then
            vRow.Item("PhoneNumber") = String.Format("{0} {1} {2}", vRow.Item("PhoneNumber"), DataSelectionText.String23336, vRow.Item("Extension"))    ' Ext
          End If
        End If
        vRow.SetYNValue("ExDirectory")
        vRow.SetYNValue("IsActive")
        vRow.SetYNValue("Mail")
        vRow.SetYNValue("DeviceDefault")
        vRow.SetYNValue("PreferredMethod")
        If mvType = DataSelectionTypes.dstContactCommsNumbers Then
          If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            If vRow.Item("Number") = vOrg.Telephone And vRow.Item("STDCode") = vOrg.STDCode And vRow.Item("DiallingCode") = vOrg.DiallingCode Then vRow.Item("Default") = "Y"
          Else
            If vRow.Item("Number") = mvContact.Telephone And vRow.Item("STDCode") = mvContact.StdCode And vRow.Item("DiallingCode") = mvContact.DiallingCode Then vRow.Item("Default") = "Y"
          End If
          vRow.SetYNValue("Default")
        End If
      Next
    End Sub
    Private Sub GetPackProductDataSheet(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPackProductDataSheet")
      Dim vSQL As String = "SELECT pp.link_product, p.product_desc,pw.warehouse, w.warehouse_desc, pw.last_stock_count, p.cost_of_sale, w2.warehouse_desc  AS default_warehouse_desc"
      vSQL = vSQL & " FROM packed_products pp, products p, product_warehouses pw, warehouses w, warehouses w2 "
      vSQL = vSQL & " WHERE pp.product = '" & mvParameters("Product").Value & "' AND pp.rate = '" & mvParameters("Rate").Value & "'"
      vSQL = vSQL & " AND p.product = pp.link_product AND pw.product = pp.link_product AND w.warehouse = pw.warehouse AND w2.warehouse = p.warehouse"
      vSQL = vSQL & " ORDER BY product_desc, w.warehouse_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetPaymentPlanAmendmentHistory(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPaymentPlanAmendmentHistory")
      Dim vRS As CDBRecordSet
      Dim vIndex As Integer
      Dim vIndex2 As Integer
      Dim vAttr As String = ""
      Dim vDesc As String
      Dim vLookupRS As CDBRecordSet
      Dim vOldMembership As String = ""
      Dim vNewMembership As String = ""

      Dim vSQL As String = "SELECT operation_date, data_values FROM amendment_history"
      vSQL = vSQL & " WHERE table_name" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, "orders")
      vSQL = vSQL & " AND select_1 = " & mvParameters.Item("PaymentPlanNumber").LongValue
      vSQL = vSQL & " AND operation" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, "update")
      '  vSQL = vSQL & " AND contact_journal_number = 9081"
      vSQL = vSQL & " ORDER BY operation_date DESC"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      With vRS
        While .Fetch()
          Dim vColumns(6) As String
          Dim vCreatedRow As Boolean = False
          vColumns(0) = .Fields.Item(1).Value
          'Find the new values in the data_values column
          Dim vPos As Integer = InStr(.Fields.Item(2).Value, "NEW")
          If vPos > 1 Then  'BR14214: If NEW is not the first word then make sure it is really the NEW string saved by the system and not a user code or some other value
            Dim vPos2 As Integer = InStr(.Fields.Item(2).Value, vbLf & "NEW" & Chr(22)) 'vbLf and Chr$(22) will always be there if NEW is not the first word and the system saved the NEW value
            If vPos2 = 0 Then vPos = 0 'If above is not found then NEW is a user code or some other value
          End If
          If vPos > 0 Then
            'Split the old and new values into arrays
            Dim vOldValues() As String = Split(Mid$(.Fields.Item(2).Value, 5, (vPos - 7)), Chr(22))
            Dim vNewValues() As String = Split(Mid$(.Fields.Item(2).Value, vPos + 4, Len(.Fields.Item(2).Value) - (vPos + 5)), Chr(22))
            'Find the old and new balance values
            For vIndex = 0 To UBound(vOldValues)
              Dim vOldValue() As String = Split(vOldValues(vIndex), ":")
              If vOldValue(0) = "balance" Then
                vColumns(4) = Format$(vOldValue(1), "Fixed")
                Exit For
              End If
            Next
            For vIndex = 0 To UBound(vNewValues)
              Dim vNewValue() As String = Split(vNewValues(vIndex), ":")
              If vNewValue(0) = "balance" Then
                vColumns(5) = Format$(vNewValue(1), "Fixed")
                Exit For
              End If
            Next
            'Loop through the old values
            For vIndex = 0 To UBound(vOldValues)
              Dim vOldValue() As String = Split(vOldValues(vIndex), ":")
              'Loop through the new values
              For vIndex2 = 0 To UBound(vNewValues)
                Dim vCreate As Boolean = False
                Dim vTable As String = ""
                Dim vNewValue() As String = Split(vNewValues(vIndex2), ":")
                If vOldValue(0) = vNewValue(0) Then
                  'When the old and new column names match decide whether the column is one that is to be output
                  Select Case vOldValue(0)
                    Case "membership_type", "payment_method"
                      vTable = vOldValue(0) & "s"
                      vAttr = vOldValue(0)
                      vCreate = True
                    Case "payment_frequency"
                      vTable = "payment_frequencies"
                      vAttr = vOldValue(0)
                      vCreate = True
                    Case "membership_rate"
                      vTable = "rates"
                      vAttr = "rate"
                      vCreate = True
                    Case "membership_product"
                      vOldMembership = vOldValue(1)
                      vNewMembership = vNewValue(1)
                    Case "renewal", "amount", "arrears", "cancellation_reason", "cancellation_source", "cancelled_by", "cancelled_on", "frequency_amount", "next_detail_line", "next_payment_due", "renewal_amount", "renewal_date"
                      vCreate = True
                  End Select
                  If vCreate Then
                    vColumns(1) = vOldValue(0)
                    If Len(vTable) > 0 Then
                      vDesc = vAttr & "_desc"
                      If vTable = "rates" Then
                        vLookupRS = mvEnv.Connection.GetRecordSet("SELECT " & vAttr & "," & vDesc & " FROM " & vTable & " WHERE product = '" & vOldMembership & "' AND " & vAttr & " = '" & vOldValue(1) & "' OR product = '" & vNewMembership & "' AND " & vAttr & " = '" & vNewValue(1) & "'")
                      Else
                        vLookupRS = mvEnv.Connection.GetRecordSet("SELECT " & vAttr & "," & vDesc & " FROM " & vTable & " WHERE " & vAttr & " IN ('" & vOldValue(1) & "','" & vNewValue(1) & "')")
                      End If
                      With vLookupRS
                        While .Fetch()
                          If .Fields.Item(1).Value = vOldValue(1) Then
                            vColumns(2) = .Fields.Item(2).Value
                          Else
                            If String.IsNullOrEmpty(vOldValue(1)) Then vColumns(2) = String.Empty   'Clear the original value
                            vColumns(3) = .Fields.Item(2).Value
                          End If
                        End While
                        .CloseRecordSet()
                      End With
                    Else
                      vColumns(2) = vOldValue(1)
                      vColumns(3) = vNewValue(1)
                    End If
                  ElseIf Not vCreate Then
                    If UBound(vOldValues) = 0 Then  'assume that the only column in the old values is "balance"
                      vColumns(1) = vOldValue(0)
                      vCreate = True
                    End If
                  End If
                  If vCreate Then
                    pDataTable.AddRowFromListWithQuotes(vColumns(0) & "," & StrConv(Replace$(vColumns(1), "_", " "), vbProperCase) & ",'" & vColumns(2) & "','" & vColumns(3) & "'," & vColumns(4) & "," & vColumns(5))
                    vCreatedRow = True
                  End If
                  Exit For
                End If
              Next
            Next
            If vCreatedRow = False AndAlso Not String.IsNullOrEmpty(vColumns(4)) AndAlso Not String.IsNullOrEmpty(vColumns(5)) Then
              If DoubleValue(vColumns(4)) <> DoubleValue(vColumns(5)) Then
                pDataTable.AddRowFromListWithQuotes(vColumns(0) & ",,,," & vColumns(4) & "," & vColumns(5))
              End If
            End If
          End If
        End While
        .CloseRecordSet()
      End With
      If pDataTable.Rows.Count > 0 Then
        pDataTable.Columns("OldBalance").FieldType = CDBField.FieldTypes.cftNumeric
        pDataTable.Columns("NewBalance").FieldType = CDBField.FieldTypes.cftNumeric
      End If
    End Sub
    Private Sub GetPaymentPlanDetails(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")
      Dim vAttrs As String = "od.contact_number,od.address_number,p.product,product_desc,od.rate,rate_desc,od.distribution_code,quantity,amount,balance,arrears,od.despatch_method,od.source,product_number,od.amended_on,od.amended_by,created_by,created_on,currency_code,distribution_code_desc,vat_exclusive"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
        vAttrs &= ",valid_from,valid_to,modifier_activity,activity_desc,modifier_activity_value,activity_value_desc,modifier_activity_quantity,modifier_activity_date,modifier_price,modifier_per_item,unit_price,pro_rated,net_amount,vat_amount,gross_amount,vat_rate,vat_percentage"
      Else
        vAttrs &= ",,,,,,,,,,,,,,,,,"
      End If
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs = vAttrs.Replace("currency_code", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanDetailCreatedBy) Then vAttrs = vAttrs.Replace("created_by,created_on", ",")

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("products p", "od.product", "p.product")
        .Add("rates r", "od.product", "r.product", "od.rate", "r.rate")
        .Add("contacts c", "od.contact_number", "c.contact_number")
        .AddLeftOuterJoin("distribution_codes dc", "od.distribution_code", "dc.distribution_code")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
          .AddLeftOuterJoin("activities a", "od.modifier_activity", "a.activity")
          .AddLeftOuterJoin("activity_values av", "od.modifier_activity", "av.activity", "od.modifier_activity_value", "av.activity_value")
        End If
      End With
      Dim vWhereFields As New CDBFields(New CDBField("order_number", mvParameters("PaymentPlanNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs) & "," & vConAttrs, "order_details od", vWhereFields, "detail_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ContactNameItems())

      'If multiple RateModifiers were used then set the data to "Multiple"
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ModifierActivity").ToUpper = "MULTI" Then
          vRow.Item("ModifierActivity") = "Multiple"
          vRow.Item("ModifierActivityValue") = "Multiple"
          vRow.Item("ModifierActivityQuantity") = ""
          vRow.Item("ModifierActivityDate") = ""
          vRow.Item("ModifierPrice") = ""
          vRow.Item("ModifierPerItem") = "Multiple"
          vRow.Item("ModifierActivityDesc") = "Multiple"
          vRow.Item("ModifierActivityValueDesc") = "Multiple"
        End If
      Next

    End Sub
    Private Sub GetCriteriaSetDetails(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins()
      Dim vAttrs As String = "csd.criteria_set,sequence_number,search_area,i_e,c_o,main_value,subsidiary_value,period,counted,and_or,left_parenthesis,right_parenthesis"
      If mvParameters.Exists("MarketingControls") AndAlso mvParameters("MarketingControls").Bool Then
        vAnsiJoins.Add("marketing_controls mc", "csd.criteria_set", "mc.criteria_set", CARE.Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      End If

      If mvParameters.Exists("CriteriaSet") Then vWhereFields.Add("csd.criteria_set", mvParameters.Item("CriteriaSet").LongValue)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "criteria_set_details csd", vWhereFields, "sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, String.Empty, String.Empty, True)
    End Sub
    Private Sub GetPaymentPlanMembers(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPaymentPlanMembers")
      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = Replace$(vContact.GetRecordSetFieldsName, "c.contact_number,", "")
      Dim vAttrs As String = "order_number,m.contact_number,m.address_number,m.source,m.age_override,member_number,m.membership_type,membership_type_desc,joined,m.branch,name,cancelled_on,cancelled_by"
      Dim vSQL As String = "SELECT " & vAttrs & "," & vAddress.GetRecordSetFieldsDetailCountrySortCode.Replace(",a.branch", "") & "," & vConAttrs & " FROM members m, contacts c, addresses a, countries co, membership_types mt, branches b, organisations o WHERE order_number = " & mvParameters("PaymentPlanNumber").Value & " AND m.contact_number = c.contact_number AND m.address_number = a.address_number AND a.country = co.country AND m.membership_type = mt.membership_type AND m.branch = b.branch AND b.organisation_number = o.organisation_number ORDER BY m.cancelled_on" & mvEnv.Connection.DBSortByNullsFirst
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "ADDRESS_LINE," & ContactNameItems())
    End Sub
    Private Sub GetPaymentPlanOutstandingOPS(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPaymentPlanOutstandingOPS")
      Dim vPaymentPlan As New PaymentPlan
      vPaymentPlan.Init(mvEnv, mvParameters("PaymentPlanNumber").LongValue)
      vPaymentPlan.GetOutstandingOPSDataTable(pDataTable)
      GetDescriptions(pDataTable, "ScheduledPaymentStatus")
      GetDescriptions(pDataTable, "ScheduleCreationReason")
    End Sub
    Private Sub GetPaymentPlanPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPaymentPlanPayments")
      SelectScheduledPaymentData(pDataTable)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Posted")
      Next
    End Sub
    Private Sub GetPaymentPlanSubscriptions(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPaymentPlanSubscriptions")
      Dim vAddress As New Address(mvEnv)
      vAddress.Init()
      Dim vAttrs As String = "subscription_number,contact_number,s.address_number,s.product,product_desc,quantity,s.despatch_method,despatch_method_desc,s.reason_for_despatch,reason_for_despatch_desc,valid_from,valid_to,cancellation_reason,cancelled_by,communication_number,cancellation_source,cancelled_on"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then vAttrs = Replace$(vAttrs, "communication_number", "")
      Dim vSQL As String = "SELECT " & RemoveBlankItems(vAttrs) & "," & vAddress.GetRecordSetFieldsDetailCountrySortCode & " FROM subscriptions s, products p, despatch_methods dm, reasons_for_despatch rfd, addresses a, countries co WHERE s.order_number = " & mvParameters("PaymentPlanNumber").Value & " AND s.product = p.product AND s.despatch_method = dm.despatch_method AND s.reason_for_despatch = rfd.reason_for_despatch AND s.address_number = a.address_number AND a.country = co.country ORDER BY s.cancellation_reason, s.product"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, "ADDRESS_LINE,,,")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then GetCommunicationInfo(pDataTable, "CommunicationNumber", "DeliverTo")
      GetDescriptions(pDataTable, "CancellationReason")
      GetDescriptions(pDataTable, "CancellationSource")
    End Sub
    Private Sub GetPersonnelContacts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = ",working_hours,evenings,weekends,p.notes"
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("c.contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.Exists("ContactName") Then vWhereFields.Add("surname", mvParameters("ContactName").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.AddJoin("p.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vContact.GetRecordSetFieldsName & vAttrs, "personnel p, contacts c", vWhereFields, "surname, forenames")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, String.Format("contact_number,{0}{1}", ContactNameItems(), vAttrs))
    End Sub

    Private Sub GetServiceProductContacts(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vFields As String = ",sp.product,sp.rate,sp.fixed_unit_rate"
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("c.contact_number", mvParameters("ContactNumber").LongValue)
      If mvParameters.Exists("ContactName") Then vWhereFields.Add("surname", mvParameters("ContactName").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.AddJoin("sp.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vContact.GetRecordSetFieldsName & vFields, "service_products sp, contacts c", vWhereFields, "surname, forenames")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "contact_number," & ContactNameItems() & vFields)
    End Sub
    Private Sub GetPositionActions(ByVal pDataTable As CDBDataTable)
      If mvParameters.ContainsKey("ContactPositionNumber") = False Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterNotFound, "ContactPositionNumber")
      ElseIf mvParameters.ParameterExists("ContactPositionNumber").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "ContactPositionNumber")
      End If
      GetActions(pDataTable)
    End Sub
    Private Sub GetPositionDocuments(ByVal pDataTable As CDBDataTable)
      If mvParameters.ContainsKey("ContactPositionNumber") = False Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterNotFound, "ContactPositionNumber")
      ElseIf mvParameters.ParameterExists("ContactPositionNumber").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "ContactPositionNumber")
      End If
      GetDocuments(pDataTable)
    End Sub
    Private Sub GetPositionTimesheets(ByVal pDataTable As CDBDataTable)
      Dim vPositionTimesheet As New ContactPositionTimesheet(mvEnv)
      Dim vAttrs As String = "contact_position_number,timesheet_number,timesheet_date,duration_hours,duration_minutes,timesheet_desc,r.role,r.role_desc"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins
      vWhereFields.Add("contact_position_number", mvParameters("ContactPositionNumber").Value)
      If mvParameters.Exists("TimesheetNumber") Then vWhereFields.Add("timesheet_number", mvParameters("TimesheetNumber").LongValue)
      vAnsiJoins.AddLeftOuterJoin("contact_roles cr", "cpt.contact_role_number", "cr.contact_role_number")
      vAnsiJoins.AddLeftOuterJoin("roles r", "cr.role", "r.role")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_position_timesheet cpt", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetPostPointRecipients(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vAttrs As String = vContact.GetRecordSetFieldsNamePhone
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("post_point", mvParameters("PostPoint").Value)
      vWhereFields.AddJoin("ppr.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "post_point_recipients ppr, contacts c", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "contact_number,CONTACT_NAME")
    End Sub
    Private Sub GetProductWarehouses(ByVal pDataTable As CDBDataTable)
      Dim vMultiWarehouses As Boolean = mvEnv.GetConfigOption("fp_stock_multiple_warehouses")
      Dim vWhereFields As New CDBFields()
      Dim vAttrs As String = "p.product,product_desc,pw.warehouse,pw.bin_number,warehouse_desc,pw.re_order_level,pw.last_stock_count,p.cost_of_sale,"
      Dim vOrderBy As String = "p.product, pw.warehouse"
      If Not vMultiWarehouses Then
        vAttrs = Replace$(vAttrs, "warehouse_desc", "")
        vAttrs = Replace$(vAttrs, "pw.re_order_level", "")
        vAttrs = Replace$(vAttrs, "pw.", "p.")
        vOrderBy = vOrderBy.Replace("pw.", "p.")
      End If
      Dim vTables As String = "products p"
      If vMultiWarehouses Then vTables &= ", product_warehouses pw, warehouses w"
      If mvParameters.Exists("Product") Then
        vWhereFields.Add("p.product", mvParameters("Product").Value)
        If mvParameters.Exists("Warehouse") Then
          If vMultiWarehouses Then
            vWhereFields.Add("pw.warehouse", mvParameters("Warehouse").Value)
          Else
            vWhereFields.Add("p.warehouse", mvParameters("Warehouse").Value)
          End If
        End If
      End If
      If vMultiWarehouses Then
        vWhereFields.AddJoin("p.product#1", "pw.product")
        vWhereFields.AddJoin("pw.warehouse#1", "w.warehouse")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vTables, vWhereFields, vOrderBy)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
      'SmartClient Trader needs to display a composite of a number of the fields
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("WarehouseStock") = vRow.Item("Warehouse") & " - " & vRow.Item("WarehouseDesc") & " (" & vRow.Item("LastStockCount") & ")"
      Next
    End Sub
    Private Sub GetPurchaseInvoiceDetails(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPurchaseInvoiceDetails")
      Dim vAttrs As String = "purchase_invoice_number,line_number,line_item,line_price,quantity,amount,nominal_account,distribution_code,Adjustment_Status,Cancellation_Reason,Cancellation_Source,Cancelled_By,Cancelled_On"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM purchase_invoice_details pod WHERE purchase_invoice_number = " & mvParameters("PurchaseInvoiceNumber").LongValue & " ORDER BY line_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub
    Private Sub GetPurchaseOrderDetails(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetPurchaseOrderDetails")
      Dim vAttrs As String = "purchase_order_number,line_number,line_item,line_price,quantity,amount,balance,nominal_account,distribution_code"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        vAttrs = vAttrs & ",product,warehouse"
      End If
      Dim vSQL As String = "SELECT " & vAttrs & " FROM purchase_order_details pod WHERE purchase_order_number = " & mvParameters("PurchaseOrderNumber").LongValue & " ORDER BY line_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",,")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        GetDescriptions(pDataTable, "Product")
        GetDescriptions(pDataTable, "Warehouse")
      End If
    End Sub
    Private Sub GetPurchaseOrderHistory(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) Then
        Dim vAttrs As String = "purchase_order_number,amount,previous_amount,previous_authorisation_level,po_authorisation_level_desc,previous_authorised_by,previous_authorised_on,changed_by,changed_on"
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("purchase_order_number", mvParameters("PurchaseOrderNumber").IntegerValue)
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.AddLeftOuterJoin("po_authorisation_levels pal", "poh.previous_authorisation_level", "pal.po_authorisation_level")
        Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "purchase_order_history poh", vWhereFields, "changed_on DESC", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub
    Private Sub GetPurchaseOrderPayments(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vAddress As New Address(mvEnv)
      Dim vAnsiJoins As New AnsiJoins
      Dim vAttrs As String = "pop.purchase_order_number,payment_number,due_date,latest_expected_date,pop.amount,percentage,authorisation_required,authorisation_status,authorised_by,authorised_on,posted_on,pop.payee_contact_number,pop.payee_address_number"
      vAttrs = vAttrs & ",cheque_produced_on,pop.cheque_reference_number,pay_by_bacs,pop.payee_reference,no_payment_required,pop.po_payment_type,po_payment_type_desc,pop.nominal_account,pop.distribution_code,pop.separate_payment,expected_receipt,expected_receipt_amount,expected_receipt_reason,receipt_for_payment_number"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPMultiplePayees) Then vAttrs = Replace$(vAttrs, "pop.payee_contact_number,pop.payee_address_number", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderManagement) Then vAttrs = Replace$(vAttrs, ",cheque_produced_on,pop.cheque_reference_number", ",,")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS) Then vAttrs = Replace$(vAttrs, ",pay_by_bacs", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments) Then vAttrs = vAttrs.Replace(",pop.payee_reference,no_payment_required,pop.po_payment_type,po_payment_type_desc,pop.nominal_account,pop.distribution_code,pop.separate_payment,expected_receipt,expected_receipt_amount,expected_receipt_reason,receipt_for_payment_number", ",,,,,,,,,,,")
      Dim vItems As String = vAttrs & "," 'Additional seperator for AuthorisationStatusDesc field populated by GetDescriptions routine
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPMultiplePayees) Then
        vAttrs = vAttrs & ",c.label_name AS payee_contact_label_name," & vContact.GetRecordSetFieldsName & "," & vAddress.GetRecordSetFieldsCountry
        vItems = vItems & ",payee_contact_label_name,CONTACT_NAME,ADDRESS_LINE"
        vAnsiJoins.Add("contacts c", "pop.payee_contact_number", "c.contact_number")
        vAnsiJoins.Add("addresses a", "pop.payee_address_number", "a.address_number")
        vAnsiJoins.Add("countries co", "a.country", "co.country")
        vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "pop.cancellation_reason", "cr.cancellation_reason")
        vAnsiJoins.AddLeftOuterJoin("sources cs", "pop.cancellation_source", "cs.source")
      Else
        vItems = vItems & ",,,"
      End If
      'BR17340
      vAttrs = vAttrs & ",pop.adjustment_status,pop.cancellation_reason,pop.cancellation_source,pop.cancelled_by,pop.cancelled_on,cs.source_desc AS cancellation_source_desc,cancellation_reason_desc,pop.pop_payment_method,ppm.pop_payment_method_desc,ppd.amount AS ppd_amount,ch.amount AS ch_amount,ppd.pop_production_number,distribution_code_desc"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPopPaymentMethod) Then vAttrs.Replace(",pop.pop_payment_method,ppm.pop_payment_method_desc", ",,,,,")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPopPaymentMethod) Then
        vAnsiJoins.Add("pop_payment_methods ppm", "ppm.pop_payment_method", "pop.pop_payment_method")
        vAnsiJoins.AddLeftOuterJoin("purchase_invoices pi", "pop.purchase_invoice_number", "pi.purchase_invoice_number")
        vAnsiJoins.AddLeftOuterJoin("cheques ch", "pi.cheque_reference_number", "ch.cheque_reference_number")
        vAnsiJoins.AddLeftOuterJoin("pop_production_details ppd", "ch.pop_production_number", "ppd.pop_production_number")
      End If
      vItems = vItems & ",adjustment_status,pop.cancellation_reason,cancellation_source,cancelled_by,cancelled_on,cancellation_source_desc,cancellation_reason_desc,pop_payment_method,pop_payment_method_desc,ppd_amount,ch_amount,pop_production_number,distribution_code_desc"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments) Then
        vAnsiJoins.AddLeftOuterJoin("po_payment_types ppt", "pop.po_payment_type", "ppt.po_payment_type")
      End If
      vAnsiJoins.AddLeftOuterJoin("distribution_codes dc", "pop.distribution_code", "dc.distribution_code")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("pop.purchase_order_number", mvParameters("PurchaseOrderNumber").IntegerValue)
      If mvParameters.ContainsKey("PaymentNumber") Then vWhereFields.Add("payment_number", mvParameters("PaymentNumber").IntegerValue)

      Dim vOrderBy As String = "payment_number"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments) Then
        vOrderBy = mvEnv.Connection.DBIsNull("receipt_for_payment_number", "payment_number") & ",payment_number"
      End If
      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "purchase_order_payments pop", vWhereFields, vOrderBy, vAnsiJoins), vItems)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("Authorisation")
        vRow.SetYNValue("ExpectedReceipt")
        vRow.SetYNValue("PayByBacs")
        vRow.SetYNValue("NoPaymentRequired")
      Next
      GetDescriptions(pDataTable, "AuthorisationStatus")
    End Sub
    Private Sub GetPurchaseOrderInformation(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "po.purchase_order_number,purchase_invoice_number,po.contact_number,po.address_number,po.amount,output_group,po.purchase_order_type,purchase_order_desc,po.payee_contact_number,po.payee_address_number,start_date,payment_frequency,number_of_payments,distribution_method,payment_as_percentage,po.campaign,po.appeal,po.segment,po.cancellation_reason,po.cancellation_source,po.cancelled_by,po.cancelled_on,cheque_reference_number,bacs_processed,po.currency_code,payment_schedule,ad_hoc_payments,regular_payments"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS) Then vFieldNames = vFieldNames.Replace("bacs_processed", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments) Then vFieldNames = vFieldNames.Replace("ad_hoc_payments", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRegularPurchaseOrderPayments) Then vFieldNames = vFieldNames.Replace("regular_payments", "")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vFieldNames = vFieldNames.Replace("po.currency_code", "")

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "PurchaseOrderNumber", "po.purchase_order_number")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("purchase_order_types pot", "po.purchase_order_type", "pot.purchase_order_type")
      vAnsiJoins.AddLeftOuterJoin("purchase_invoices pi", "po.purchase_order_number", "pi.purchase_order_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFieldNames), "purchase_orders po", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFieldNames)
    End Sub
    Private Sub GetChequeInformation(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "c.cheque_reference_number,c.cheque_number,c.contact_number,c.address_number,c.amount,c.printed_on,c.reconciled_on,c.cheque_status,cs.cheque_status_desc,cs.allow_reissue"
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vSupportsChequeHistory As Boolean = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataChequeReissue) AndAlso Not mvParameters.ParameterExists("ChequeReconciliation").Bool
      If mvParameters.Exists("ChequeNumber") Then
        If vSupportsChequeHistory Then
          vWhereFields.Add("c.cheque_number", CDBField.FieldTypes.cftLong, mvParameters("ChequeNumber").Value, CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("ch.cheque_number", CDBField.FieldTypes.cftLong, mvParameters("ChequeNumber").Value, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
          vAnsiJoins.AddLeftOuterJoin("cheque_history ch", "c.cheque_reference_number", "ch.cheque_reference_number")
        Else
          vWhereFields.Add("c.cheque_number", CDBField.FieldTypes.cftLong, mvParameters("ChequeNumber").Value)
        End If
      End If
      vAnsiJoins.AddLeftOuterJoin("cheque_statuses cs", "c.cheque_status", "cs.cheque_status")
      If mvParameters.Exists("ChequeReferenceNumber") Then vWhereFields.Add("c.cheque_reference_number", CDBField.FieldTypes.cftLong, mvParameters("ChequeReferenceNumber").Value)
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFieldNames, "cheques c", vWhereFields, "", vAnsiJoins)
      vSQL.Distinct = vSupportsChequeHistory
      pDataTable.FillFromSQL(mvEnv, vSQL, False)
    End Sub
    Private Sub GetPurchaseInvoiceInformation(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "purchase_invoice_number,purchase_order_number,contact_number,address_number,amount,payee_contact_number,payee_address_number,payee_reference,purchase_invoice_date,cheque_reference_number,source,campaign,appeal,segment,currency_code"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vFieldNames = vFieldNames.Replace("currency_code", "")
      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "PurchaseInvoiceNumber", "purchase_invoice_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFieldNames), "purchase_invoices", vWhereFields)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFieldNames)
    End Sub
    Private Sub GetRates(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As String = "r.rate,rate_desc,p.product,product_desc,current_price,future_price,price_change_date,vat_exclusive,current_price_lower_limit,current_price_upper_limit"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then vFieldNames = vFieldNames & ",fixed_price"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMinPriceMandatory) Then vFieldNames = vFieldNames & ",upper_lower_price_mandatory"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then vFieldNames = vFieldNames & ",r.price_is_percentage"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDaysPrior) Then
        vFieldNames = vFieldNames & "," & mvEnv.Connection.DBIsNull("r.days_prior_to", "0") & "As days_prior_to," & mvEnv.Connection.DBIsNull("r.days_prior_From", "99999") & "As days_prior_From"
      Else
        vFieldNames = vFieldNames & ",NULL As days_prior_to,NULL As days_prior_From"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipLookupGroup) Then
        vFieldNames = vFieldNames & ",membership_lookup_group"
      Else
        vFieldNames = vFieldNames & ",Null as membership_lookup_group"
      End If
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("rates r", "p.product", "r.product")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDaysPrior) And mvParameters.Exists("EventNumber") And mvParameters.Exists("BookingDate") Then
        vFieldNames = vFieldNames & ",e.start_date "
        vAnsiJoins.Add("event_booking_options ebo", "ebo.product", "p.product")
        vAnsiJoins.Add("events e", "ebo.event_number", "e.event_number")
      Else
        vFieldNames = vFieldNames & " ,Null as start_date "
      End If
      AddWhereFieldFromParameter(vWhereFields, "Product", "p.product")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then AddWhereFieldFromParameter(vWhereFields, "CurrencyCode", "r.currency_code")
      AddWhereFieldFromParameter(vWhereFields, "Rate", "r.rate")
      vWhereFields.Add("r.history_only", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
      If mvParameters.ContainsKey("CurrentPrice") Then vWhereFields.Add("r.current_price", CDBField.FieldTypes.cftInteger, mvParameters("CurrentPrice").Value)
      If mvParameters.ContainsKey("VatExclusive") Then vWhereFields.Add("r.vat_exclusive", mvParameters("VatExclusive").Value)
      If mvParameters.ParameterExists("ExtraSessionFeeRate").Value = "Y" Then
        vWhereFields.Add("r.current_price", CDBField.FieldTypes.cftInteger, "0", CDBField.FieldWhereOperators.fwoNotEqual)
        vWhereFields.Add("r.vat_exclusive", "Y")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDaysPrior) And mvParameters.Exists("EventNumber") And mvParameters.Exists("BookingDate") Then
        vWhereFields.Add("e.event_number", CDBField.FieldTypes.cftInteger, mvParameters("EventNumber").IntegerValue)
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifier) Then
        vFieldNames = vFieldNames & ",r.use_modifiers "
      Else
        vFieldNames = vFieldNames & " ,Null as use_modifiers "
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vFieldNames = vFieldNames & ",currency_code"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoans) Then
        vFieldNames &= ", loan_interest"
      Else
        vFieldNames &= ", Null AS loan_interest"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipLookupGroup) And mvParameters.Exists("ContactNumber") Then
        GetMembershipLookupGroupSQL(vAnsiJoins, vWhereFields)
      End If
      If mvParameters.Exists("LoanInterest") = True AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoans) = True Then
        vWhereFields.Add("loan_interest", mvParameters("LoanInterest").Value)
        If mvParameters("LoanInterest").Bool Then vWhereFields.Add("r.current_price", CDBField.FieldTypes.cftInteger, "0")
      End If
      If mvParameters.Exists("ZeroPrice") = True Then
        vWhereFields.Add("r.current_price", CDBField.FieldTypes.cftInteger, "0")
        vWhereFields.Add("r.future_price", CDBField.FieldTypes.cftInteger, "0")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFieldNames, "products p", vWhereFields, "", vAnsiJoins)
      vSQLStatement.Distinct = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDaysPrior) And mvParameters.Exists("EventNumber") And mvParameters.Exists("BookingDate") Then
        Dim vIndex As Integer
        While vIndex <= pDataTable.Rows.Count - 1
          Dim vDiff As Integer
          vDiff = IntegerValue(DateDiff(DateInterval.Day, CDate(mvParameters("BookingDate").Value), CDate(pDataTable.Rows(vIndex).Item("StartDate").ToString)).ToString)
          If vDiff >= 0 Then
            If vDiff >= IntegerValue(pDataTable.Rows(vIndex).Item("DaysPriorTo").ToString) And vDiff <= IntegerValue(pDataTable.Rows(vIndex).Item("DaysPriorFrom").ToString) Then
              'Don't remove record
            Else
              pDataTable.RemoveRow(pDataTable.Rows(vIndex))
              vIndex = vIndex - 1
            End If
          End If
          vIndex += 1
        End While
      End If
    End Sub
    Private Sub GetRelationshipsDataSheet(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vAttrs As String = "r.relationship,relationship_desc,mandatory,to_contact_group,multiple_entries,contact_selection_type,complimentary_relationship,post_point"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM relationship_groups rg,relationship_group_details rgd, relationships r WHERE "
      If mvParameters.Exists("RelationshipGroupCode") Then vSQL = vSQL & " rg.relationship_group = '" & mvParameters("RelationshipGroupCode").Value & "' AND "
      vSQL = vSQL & "rg.usage_code = '" & mvParameters("UsageCode").Value & "' AND rg.relationship_group = rgd.relationship_group AND rgd.relationship = r.relationship AND (r.from_contact_group IS NULL OR r.from_contact_group = '" & mvParameters("FromContactGroupCode").Value & "')"
      vSQL = vSQL & " ORDER BY sequence_number,relationship_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetSalesContacts(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vWhereFields As New CDBFields()
      vWhereFields.AddJoin("sp.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vContact.GetRecordSetFieldsName, "sales_persons sp,contacts c", vWhereFields, "surname, forenames")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "contact_number," & ContactNameItems())
    End Sub
    Private Sub GetSegmentCostCentres(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSegmentCostCentres")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "campaign,appeal,segment,cost_centre,appeal_cost_centre_desc,cost_centre_percentage"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("Segment") Then .Add("segment", mvParameters.Item("Segment").Value)
        If mvParameters.Exists("CostCentre") Then .Add("cs.cost_centre", mvParameters.Item("CostCentre").Value)
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM segment_cost_centres cs, appeal_cost_centres acs WHERE "
      If vWhereFields.Count > 0 Then vSQL = vSQL & mvEnv.Connection.WhereClause(vWhereFields) & " AND "
      vSQL = vSQL & "cs.cost_centre = acs.appeal_cost_centre"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetSegmentProducts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "spa.campaign,spa.appeal,spa.segment,spa.amount_number,spa.product,p.product_desc,spa.rate,r.rate_desc"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("spa.campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("spa.appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("Segment") Then .Add("spa.segment", mvParameters.Item("Segment").Value)
        If mvParameters.Exists("AmountNumber") Then .Add("spa.amount_number", mvParameters.Item("AmountNumber").LongValue)
      End With
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("products p", " spa.product", "p.product", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.AddLeftOuterJoin("rates r", "spa.product", "r.product", "spa.rate", "r.rate")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "segment_product_allocation spa", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub
    Private Sub GetSelectionSetAppointments(ByVal pDataTable As CDBDataTable)
      Dim vTable As String = GetSelectionSetTableName()
      Dim vContact As New Contact(mvEnv)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("sc.selection_set", mvParameters("SelectionSetNumber").LongValue)
      vWhereFields.Add("revision", 1)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "sc.contact_number", "c.contact_number")

      Dim vSubWhere As New CDBFields
      vSubWhere.Add("sc.selection_set", mvParameters("SelectionSetNumber").LongValue)
      vSubWhere.Add("revision", 1)
      If mvParameters.Exists("EndDate") Then vSubWhere.Add("start_date", CDBField.FieldTypes.cftTime, mvParameters("EndDate").Value, CDBField.FieldWhereOperators.fwoLessThan)
      If mvParameters.Exists("StartDate") Then vSubWhere.Add("end_date", CDBField.FieldTypes.cftTime, mvParameters("StartDate").Value, CDBField.FieldWhereOperators.fwoGreaterThan)
      Dim vAttrs As String = ",start_date,end_date,record_type,unique_id,description,time_status,ca.amended_by,ca.amended_on"
      Dim vSubAnsi As New AnsiJoins
      vSubAnsi.Add("contact_appointments ca", "sc.contact_number", "ca.contact_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, "sc.contact_number" & vAttrs, "selected_contacts sc", vSubWhere, "", vSubAnsi)
      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) ca", vSQL.SQL), "sc.contact_number", "ca.contact_number")

      vAnsiJoins.AddLeftOuterJoin("fundraising_requests fr", "sc.contact_number", "fr.contact_number")
      vAnsiJoins.AddLeftOuterJoin("fundraising_request_stages frs", "fr.fundraising_request_stage", "frs.fundraising_request_stage")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vContact.GetRecordSetFieldsName() & ",fundraising_request_stage_desc" & vAttrs, vTable, vWhereFields, "sc.contact_number,start_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "contact_number,CONTACT_NAME,fundraising_request_stage_desc" & vAttrs)
    End Sub
    Private Sub GetSelectionSetCommsNumbers(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSelectionSetCommsNumbers")
      Dim vTable As String
      If mvEnv.Connection.GetValue("SELECT selection_group FROM selection_sets WHERE selection_set = " & mvParameters("SelectionSetNumber").LongValue) = "AU" Then
        vTable = "selected_contacts_temp"
      Else
        vTable = "selected_contacts"
      End If
      Dim vAttrs As String = "sc.contact_number,co.address_number,co.device,device_desc,co.dialling_code,co.std_code,extension,co.ex_directory,co.notes,co.amended_by,co.amended_on,communication_number,telephone"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & Replace$(RemoveBlankItems(vAttrs), "telephone", mvEnv.Connection.DBSpecialCol("", "number") & " AS telephone")
      vSQL = vSQL & " FROM " & vTable & " sc INNER JOIN contacts c ON sc.contact_number = c.contact_number"
      vSQL = vSQL & " INNER JOIN communications co ON c.contact_number = co.contact_number"
      vSQL = vSQL & " INNER JOIN devices d ON co.device = d.device"
      vSQL = vSQL & " WHERE sc.selection_set = " & mvParameters("SelectionSetNumber").LongValue & " AND revision = 1 AND c.contact_type = 'C' ORDER BY sc.contact_number, d.sequence_number, device_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, "CONTACT_TELEPHONE")
      vAttrs = "sc.contact_number,sc.address_number,co.device,device_desc,co.dialling_code,co.std_code,extension,co.ex_directory,co.notes,co.amended_by,co.amended_on,communication_number,telephone"
      vSQL = "SELECT /* SQLServerCSC */ " & Replace$(RemoveBlankItems(vAttrs), "telephone", mvEnv.Connection.DBSpecialCol("", "number") & " AS telephone")
      vSQL = vSQL & " FROM " & vTable & " sc INNER JOIN organisation_addresses oa ON sc.address_number = oa.address_number"
      vSQL = vSQL & " INNER JOIN communications co ON oa.address_number = co.address_number"
      vSQL = vSQL & " INNER JOIN devices d ON co.device = d.device"
      vSQL = vSQL & " WHERE sc.selection_set = " & mvParameters("SelectionSetNumber").LongValue & " AND revision = 1 AND co.contact_number IS NULL ORDER BY sc.contact_number, d.sequence_number, device_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, "CONTACT_TELEPHONE")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ExDirectory") = "Y" Then
          vRow.Item("PhoneNumber") = DataSelectionText.String23335    'Ex-Directory
        End If
        vRow.SetYNValue("ExDirectory")
      Next
    End Sub
    Private Sub GetSelectionSetContacts(ByVal pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vAddress As New Address(mvEnv)
      Dim vTable As String = GetSelectionSetTableName()

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "sc.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "sc.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      mvEnv.User.AddOwnershipJoins(vAnsiJoins, "c")
      vAnsiJoins.AddLeftOuterJoin("contact_positions cp", "sc.contact_number", "cp.contact_number", "sc.address_number", "cp.address_number")
      vAnsiJoins.AddLeftOuterJoin("organisations o", "cp.organisation_number", "o.organisation_number")

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("sc.selection_set", mvParameters("SelectionSetNumber").LongValue)
      vWhereFields.Add("revision", 1)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      mvEnv.User.AddOwnershipWhere(vWhereFields, "c")

      Dim vAttrs As String
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vAttrs = ",ownership_group_desc,department_desc,ownership_access_level_desc,ogu.ownership_access_level"
      Else
        vAttrs = ""
      End If
      Dim vFields As String = vContact.GetRecordSetFieldsNamePhoneGroup & ",date_of_birth,c.department,c.status,contact_position_number,position,o.organisation_number,name," & vAddress.GetRecordSetFieldsCountry & vAttrs
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, vTable, vWhereFields, "surname,forenames", vAnsiJoins)
      Dim vSelAttrs As String = "contact_number,CONTACT_NAME,contact_type,CONTACT_GROUP,CONTACT_TELEPHONE,address_number,address_type,house_name,address,town,county,postcode,country,ADDRESS_LINE,status,title,forenames,preferred_forename,surname,contact_position_number,position,organisation_number,name,date_of_birth" & vAttrs
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vSelAttrs)

      vAnsiJoins.Clear()
      vAnsiJoins.Add("organisations o", "sc.contact_number", "o.organisation_number")
      vAnsiJoins.Add("addresses a", "sc.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      mvEnv.User.AddOwnershipJoins(vAnsiJoins, "o")
      vWhereFields.Clear()
      vWhereFields.Add("sc.selection_set", mvParameters("SelectionSetNumber").LongValue)
      vWhereFields.Add("revision", 1)
      mvEnv.User.AddOwnershipWhere(vWhereFields, "o")
      vFields = "sc.contact_number,o.organisation_number,name,dialling_code,std_code,telephone,o.department,o.status,organisation_group," & vAddress.GetRecordSetFieldsCountry & vAttrs
      vSQLStatement = New SQLStatement(mvEnv.Connection, vFields, vTable, vWhereFields, "name", vAnsiJoins)
      vSelAttrs = "contact_number,name,ORGANISATION_TYPE_1,ORGANISATION_GROUP,ORGANISATION_TELEPHONE,address_number,address_type,house_name,address,town,county,postcode,country,ADDRESS_LINE,status,,,,,,,organisation_number,name," & vAttrs
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vSelAttrs)

      pDataTable.RemoveDuplicateRows("ContactNumber")
    End Sub
    Private Sub GetSelectionSteps(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSelectionSteps")
      'Dim vSQL As String = "SELECT sequence_number,select_action AS select_action_desc,view_name_desc,filter_sql,record_count,select_action,sh.view_name FROM selection_steps sh, view_names vn WHERE criteria_set = " & mvParameters("CriteriaSet").LongValue & " AND sh.view_name = vn.view_name ORDER BY sequence_number"
      Dim vAttr As String = "sequence_number,select_action AS select_action_desc,view_name_desc,filter_sql,record_count,select_action,sh.view_name,cs.criteria_set_desc"
      Dim vAnsiJoin As New AnsiJoins()
      vAnsiJoin.Add("view_names vn", "sh.view_name", "vn.view_name", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoin.Add("criteria_sets cs", "sh.criteria_set", "cs.criteria_set")

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("sh.criteria_set", CDBField.FieldTypes.cftNumeric, mvParameters("CriteriaSet").LongValue)

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttr, "selection_steps sh", vWhereFields, "sequence_number", vAnsiJoin)

      pDataTable.FillFromSQL(mvEnv, vSQL) '    .FillFromSQLDONOTUSE(mvEnv, vSQL)
      Dim vSelectionStep As New SelectionStep
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("SelectActionDesc") = vSelectionStep.GetActionDesc(vRow.Item("SelectAction"))
      Next
    End Sub
    Private Sub GetSelectItemAddresses(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSelectItemAddresses")
      Dim vAttrs As String = "address_usage_desc,notes"
      Dim vTable As String
      Dim vAttr As String
      If Contact.GetContactType(mvParameters.Item("ContactType").Value) = Contact.ContactTypes.ctcContact Then
        vTable = "contact_address_usages"
        vAttr = "contact_number"
      Else
        vTable = "organisation_address_usages"
        vAttr = "organisation_number"
      End If
      Dim vSQL As String = "SELECT " & vAttrs & " FROM " & vTable & " cau, address_usages au WHERE " & vAttr & " = " & mvParameters("ContactNumber").LongValue & " AND address_number = " & mvParameters("AddressNumber").LongValue & " AND cau.address_usage = au.address_usage"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetSelectItemCreditAccount(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSelectItemCreditAccount")
      Dim vAttrs As String = "sales_ledger_account,company,credit_category,stop_code,customer_type"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM credit_customers WHERE contact_number = " & mvParameters.Item("ContactNumber").LongValue & " AND company = '" & mvParameters.Item("Company").Value & "'"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
    End Sub
    Private Sub GetSelectItemSelectionSets(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      Dim vAttrs As String = "selection_set,selection_set_desc,department_desc,user_name,number_in_set,selection_group,custom_data"
      Dim vSQL As String = "SELECT " & vAttrs & " FROM selection_sets ss, departments d WHERE selection_group IN ('GM','AU') AND ((ss.department = '" & mvParameters.Item("Department").Value & "') OR (user_name = '" & mvEnv.User.Logname & "'))"
      If mvParameters.Exists("SelectionSetNumber") Then
        vSQL = vSQL & " AND selection_set = " & mvParameters("SelectionSetNumber").LongValue
      End If
      If mvParameters.Exists("SelectionSetDesc") Then
        vSQL = vSQL & " AND selection_set_desc " & mvEnv.Connection.DBLikeOrEqual(mvParameters("SelectionSetDesc").Value)
      End If
      If mvParameters.Item("UserOnly").Value = "Y" Then vSQL = vSQL & " AND user_name = '" & mvEnv.User.Logname & "'"
      vSQL = vSQL & " AND ss.department = d.department ORDER BY selection_set_desc"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("Auto") = If(vRow.Item("Auto") = "AU", ProjectText.String15904, "") 'Yes
        vRow.SetYNValue("Custom")
      Next
    End Sub
    Private Sub GetServiceStartDays(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetServiceStartDays")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceStartDays) Then
        'Construct a base SQL statement
        Dim vBaseAttrs As String = "start_day_number,%1,start_day,duration_days,valid_from,valid_to"
        Dim vBaseSQL As String = "SELECT %1 FROM %2 WHERE %3 = %4 AND ((valid_from IS NULL AND valid_to IS NULL) OR (valid_from IS NOT NULL AND valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, mvParameters.Item("StartDate").Value) & " AND valid_to IS NOT NULL AND valid_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvParameters.Item("StartDate").Value) & "))"
        'First see if any data exists for the service contact
        Dim vAttrs As String = Replace(vBaseAttrs, "%1", "contact_number")
        Dim vSQL As String = Replace(vBaseSQL, "%1", vAttrs)
        vSQL = Replace(vSQL, "%2", "service_start_days")
        vSQL = Replace(vSQL, "%3", "contact_number")
        vSQL = Replace(vSQL, "%4", mvParameters("ContactNumber").Value)
        pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
        If pDataTable.Rows.Count = 0 Then
          'None exists so see if any data exists for the service control
          vAttrs = Replace(vBaseAttrs, "%1", "contact_group")
          vSQL = Replace(vBaseSQL, "%1", vAttrs)
          vSQL = Replace(vSQL, "%2", "service_control_start_days")
          vSQL = Replace(vSQL, "%3", "contact_group")
          vSQL = Replace(vSQL, "%4", "'" & mvParameters("ContactGroup").Value & "'")
          pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs)
        End If
      End If
    End Sub
    Private Sub GetServiceControlRestrictions(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceControlRestrictions) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
        vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
        vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        Dim vSQL As New SQLStatement(mvEnv.Connection, "service_restriction_number,contact_number,short_stay_duration,late_booking_days,valid_from,valid_to", "service_control_restrictions", vWhereFields)
        pDataTable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub
    Private Sub GetSuppliers(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetSuppliers")
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "campaign,appeal,segment,cs.contact_number,name,supplier_role,cs.notes,cs.amended_by,cs.amended_on"
      With vWhereFields
        'Always add Campaign/Appeal/Segment even if they are not in mvParameters
        .Add("cs.campaign", mvParameters.ParameterExists("Campaign").Value)
        .Add("cs.appeal", mvParameters.ParameterExists("Appeal").Value)
        .Add("cs.segment", mvParameters.ParameterExists("Segment").Value)
        If mvParameters.Exists("OrganisationNumber") Then .Add("cs.contact_number", mvParameters.Item("OrganisationNumber").LongValue)
        .AddJoin("o.organisation_number", "cs.contact_number")
      End With
      Dim vSQL As String = "SELECT " & vAttrs & " FROM campaign_suppliers cs, organisations o WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vAttrs, ",,,")
    End Sub
    Private Sub GetSuppressionDataSheet(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.ParameterExists("SuppressionGroup").Value.Length > 0 Then vWhereFields.Add("sg.suppression_group", mvParameters("SuppressionGroup").Value)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("suppression_group_details sgd", "sgd.suppression_group", "sg.suppression_group")
      vAnsiJoins.Add(" mailing_suppressions ms", "sgd.mailing_suppression", "ms.mailing_suppression")
      Dim vSQL As New SQLStatement(mvEnv.Connection, "ms.mailing_suppression,mailing_suppression_desc", "suppression_groups sg", vWhereFields, "sequence_number,mailing_suppression_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetTickBoxes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As String = "campaign,appeal,segment,tick_box_number,av.activity,av.activity_value,ms.mailing_suppression,a.activity_desc,av.activity_value_desc,ms.mailing_suppression_desc"
      With vWhereFields
        If mvParameters.Exists("Campaign") Then .Add("tb.campaign", mvParameters.Item("Campaign").Value)
        If mvParameters.Exists("Appeal") Then .Add("tb.appeal", mvParameters.Item("Appeal").Value)
        If mvParameters.Exists("Segment") Then .Add("tb.segment", mvParameters.Item("Segment").Value)
        If mvParameters.Exists("TickBoxNumber") Then .Add("tb.tick_box_number", mvParameters.Item("TickBoxNumber").LongValue)
      End With
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.AddLeftOuterJoin("activity_values av", "tb.activity", "av.activity", "tb.activity_value", "av.activity_value")
      vAnsiJoins.AddLeftOuterJoin("activities a", "tb.activity", "a.activity")
      vAnsiJoins.AddLeftOuterJoin("mailing_suppressions ms", "tb.mailing_suppression", "ms.mailing_suppression")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "tick_boxes tb", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetTopicsDataSheet(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      If mvParameters.Exists("TopicGroupCode") Then vWhereFields.Add("tg.topic_group", mvParameters("TopicGroupCode").Value)
      vWhereFields.Add("tg.usage_code", mvParameters("UsageCode").Value)
      vWhereFields.Add("t.history_only", "N")
      vWhereFields.Add("st.history_only", "N")
      vWhereFields.AddJoin("tg.topic_group#2", "tgd.topic_group")
      vWhereFields.AddJoin("tgd.topic", "t.topic")
      vWhereFields.AddJoin("t.topic", "st.topic")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "tgd.topic,topic_desc,st.sub_topic,sub_topic_desc,mandatory,quantity_required,multiple_values,primary_topic,sequence_number", "topic_groups tg, topic_group_details tgd, topics t, sub_topics st", vWhereFields, "sequence_number,topic_desc,sub_topic_desc")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub
    Private Sub GetTransactionAnalysis(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = ""
      Dim vAttrs As String = "bta.line_number,bta.product,bta.rate,bta.distribution_code,bta.quantity,bta.amount,bta.vat_amount,bta.vat_rate,bta.source,s.source_desc,bta.currency_amount,bta.currency_vat_amount,bta.sales_contact_number,,bta.product_number,bta.notes,bta.line_type"
      vAttrs &= ",{0},{1}"    'rgb amount & rgb currency amount
      If mvType = DataSelectionTypes.dstTransactionAnalysis Then
        vAttrs &= ",bta.invoice_number,,,,p.product_desc,r.rate_desc,e.event_desc,e.start_date,o.name,c.label_name,exam_unit_description,,bta.member_number,bta.order_number,bta.covenant_number"
        vFields = vAttrs & ",NULL AS ItemType, Null as Description"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then
          vFields = vFields.Replace("exam_unit_description", mvEnv.Connection.DBIsNull("eu1.exam_unit_description", "eu2.exam_unit_description") & " AS exam_unit_description")
        Else
          vFields = vFields.Replace("exam_unit_description", "NULL AS exam_unit_description")
        End If
        vAttrs = vAttrs & ",ItemType,Description"
      Else
        vFields = vAttrs & ",member_number,order_number,covenant_number"
      End If
      vFields = String.Format(vFields, "ba.rgb_value AS rgb_amount", "ba.rgb_value AS rgb_currency_amount")
      vAttrs = String.Format(vAttrs, "rgb_amount", "rgb_currency_amount")

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("bta.batch_number", mvParameters("BatchNumber").LongValue)
      vWhereFields.Add("bta.transaction_number", mvParameters("TransactionNumber").LongValue)
      If mvParameters.HasValue("LineNumber") Then vWhereFields.Add("bta.line_number ", mvParameters("LineNumber").LongValue)
      Select Case mvType
        Case DataSelectionTypes.dstSalesTransactionAnalysis
          vWhereFields.Add("bta.sales_contact_number", mvParameters("ContactNumber").IntegerValue)
        Case DataSelectionTypes.dstDeliveryTransactionAnalysis
          vWhereFields.Add("bta.contact_number", mvParameters("ContactNumber").IntegerValue)
      End Select
      vAnsiJoins.Add("sources s", "bta.source", "s.source")
      vAnsiJoins.Add("batches b", "bta.batch_number", "b.batch_number")
      vAnsiJoins.Add("bank_accounts ba", "b.bank_account", "ba.bank_account")
      If mvType = DataSelectionTypes.dstTransactionAnalysis Then
        vAnsiJoins.Add("products p", " bta.product", "p.product", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vAnsiJoins.AddLeftOuterJoin("rates r", "bta.product", "r.product", "bta.rate", "r.rate ")
        vAnsiJoins.AddLeftOuterJoin("event_bookings eb ", "bta.batch_number", "eb.batch_number", "bta.transaction_number", "eb.transaction_number", "bta.line_number", "eb.line_number")
        vAnsiJoins.AddLeftOuterJoin("events e", "eb.event_number", "e.event_number")

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then
          vAnsiJoins.AddLeftOuterJoin("exam_booking_transactions ebt ", "bta.batch_number", "ebt.batch_number", "bta.transaction_number", "ebt.transaction_number", "bta.line_number", "ebt.line_number")
          vAnsiJoins.AddLeftOuterJoin("exam_booking_units ebu1 ", "ebt.exam_booking_unit_id", "ebu1.exam_booking_unit_id")
          vAnsiJoins.AddLeftOuterJoin("exam_units eu1", "ebu1.exam_unit_id", "eu1.exam_unit_id")

          vAnsiJoins.AddLeftOuterJoin("exam_booking_units ebu2", "bta.batch_number", "ebu2.batch_number", "bta.transaction_number", "ebu2.transaction_number", "bta.line_number", "ebu2.line_number")
          vAnsiJoins.AddLeftOuterJoin("exam_units eu2", "ebu2.exam_unit_id", "eu2.exam_unit_id")
        End If
        vAnsiJoins.AddLeftOuterJoin("contact_room_bookings crb", "bta.batch_number", "crb.batch_number", "bta.transaction_number", "crb.transaction_number", "bta.line_number", "crb.line_number")
        vAnsiJoins.AddLeftOuterJoin("room_block_bookings rbb", "crb.block_booking_number", "rbb.block_booking_number")
        vAnsiJoins.AddLeftOuterJoin("organisations o", "rbb.organisation_number", "o.organisation_number")
        vAnsiJoins.AddLeftOuterJoin("service_bookings sb", "bta.batch_number ", "sb.batch_number", "bta.transaction_number", "sb.transaction_number", "bta.line_number", "sb.line_number")
        vAnsiJoins.AddLeftOuterJoin("contacts c", "sb.booking_contact_number", "c.contact_number ")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "batch_transaction_analysis bta", vWhereFields, "bta.line_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",LINE_TYPE_NUMBER")

      For Each vRow As CDBDataRow In pDataTable.Rows
        If mvType = DataSelectionTypes.dstTransactionAnalysis Then
          If vRow.Item("ExamUnitDescription").Length > 0 Then
            vRow.Item("ItemType") = "Q"
            vRow.Item("Description") = vRow.Item("ExamUnitDescription")
          ElseIf vRow.Item("EventDesc").Length > 0 Then
            vRow.Item("ItemType") = "E"
            vRow.Item("Description") = vRow.Item("EventDesc")
          ElseIf vRow.Item("RoomDesc").Length > 0 Then
            vRow.Item("ItemType") = "A"
            vRow.Item("Description") = vRow.Item("RoomDesc")
          ElseIf vRow.Item("ServiceDesc").Length > 0 Then
            vRow.Item("ItemType") = "V"
            vRow.Item("Description") = vRow.Item("ServiceDesc")
          ElseIf vRow.Item("ProductDesc").Length > 0 Then
            vRow.Item("Description") = vRow.Item("ProductDesc")
            vRow.Item("ItemType") = "P"
          Else
            vRow.Item("ItemType") = vRow.Item("LineType")
          End If
          vRow.Item("LineType") = vRow.Item("ItemType")

          'BR16443 - Populate Number with payment plan(order) number
          Select Case (vRow.Item("LineType").ToUpper)
            Case "C" 'Covenant
              If vRow.Item("CovenantNumber").Length > 0 Then
                vRow.Item("Number") = vRow.Item("CovenantNumber")
              Else
                vRow.Item("Number") = vRow.Item("OrderNumber")
              End If
            Case "M" 'Membership
              If vRow.Item("MemberNumber").Length > 0 Then
                vRow.Item("Number") = vRow.Item("MemberNumber")
              Else
                vRow.Item("Number") = vRow.Item("OrderNumber")
              End If
            Case "O" 'Order number
              vRow.Item("Number") = vRow.Item("OrderNumber")
            Case Else
          End Select
        End If

        CheckAmountRGBValue(vRow)
      Next

      GetSalesContactNames(pDataTable)
      GetLookupData(pDataTable, "LineType", "batch_transaction_analysis", "line_type")
    End Sub
    Private Sub GetTransactionDetails(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetTransactionDetails")
      DataTableTransactionDetails(pDataTable)
      ' BR11756
    End Sub
    Private Sub GetUnauthorisedPOPayments(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As New StringBuilder
      vAttrs.Append("due_date,latest_expected_date,pop.amount,percentage,pop.purchase_order_number,payment_number,pop.authorisation_status,authorisation_status_desc,contact_number,CONTACT_NAME")
      Dim vSupportsMultiplePayees As Boolean = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPMultiplePayees)
      Dim vPOPPayByBacs As Boolean = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPOPPayByBACS)
      If vSupportsMultiplePayees Then
        vAttrs.Append(",payee_contact_number,PAYEE_CONTACT_NAME")
      Else
        vAttrs.Append(",,")
      End If
      Dim vSQLAttrs As New StringBuilder
      vSQLAttrs.Append(New Contact(mvEnv).GetRecordSetFieldsName("c"))
      vSQLAttrs.Append(",due_date,latest_expected_date,pop.amount,percentage,pop.purchase_order_number,payment_number,pop.authorisation_status,authorisation_status_desc")
      If vSupportsMultiplePayees Then
        vSQLAttrs.Append(",")
        vSQLAttrs.Append(New Contact(mvEnv).GetRecordSetFieldsName("pc", "payee"))
      End If
      If vPOPPayByBacs Then
        vSQLAttrs.Append(",pay_by_bacs")
        vAttrs.Append(",pay_by_bacs")
      Else
        vAttrs.Append(",")
      End If
      vSQLAttrs.Append(",po.purchase_order_type,purchase_order_type_desc")
      vAttrs.Append(",purchase_order_type,purchase_order_type_desc")
      If vSupportsMultiplePayees Then
        Dim vAddress As New Address(mvEnv)
        vSQLAttrs.Append(",c.label_name AS payee_contact_label_name,c.salutation AS payee_contact_salutation," & vAddress.GetRecordSetFieldsCountry)
        vAttrs.Append(",payee_contact_label_name,salutation,ADDRESS_LINE")
      Else
        vAttrs.Append(",,,")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPopPaymentMethod) Then
        vSQLAttrs.Append(",pop.pop_payment_method,ppm.pop_payment_method_desc")
        vAttrs.Append(",pop.pop_payment_method,ppm.pop_payment_method_desc")
      Else
        vAttrs.Append(",,")
      End If

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPopBankAccount) Then
        vSQLAttrs.Append(",ppm.bank_account")
        vAttrs.Append(",ppm.bank_account")
      Else
        vAttrs.Append(",")
      End If

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("pop.authorisation_required", "Y")
      vWhereFields.Add("pop.authorised_by")
      vWhereFields.Add("po.cancelled_on")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) Then
        vWhereFields.Add("pot.requires_authorisation", "Y", CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("po.authorised_by", "", CDBField.FieldWhereOperators.fwoCloseBracket)
        vWhereFields.Add("pop.no_payment_required", "N")
      End If

      If mvParameters.Exists("PoPaymentDueDateFrom") AndAlso IsDate(mvParameters("PoPaymentDueDateFrom").Value) And
        mvParameters.Exists("PoPaymentDueDateTo") AndAlso IsDate(mvParameters("PoPaymentDueDateTo").Value) Then
        If Not ValidateFromToDates(CDate(mvParameters("PoPaymentDueDateFrom").Value), CDate(mvParameters("PoPaymentDueDateTo").Value)) Then
          RaiseError(DataAccessErrors.daeInvalidDateRange)
        End If
        vWhereFields.Add("pop.due_date", CDate(mvParameters("PoPaymentDueDateFrom").Value).Date, CDBField.FieldWhereOperators.fwoBetweenFrom)
        vWhereFields.Add("pop.due_date_2", CDate(mvParameters("PoPaymentDueDateTo").Value).Date, CDBField.FieldWhereOperators.fwoBetweenTo)
      ElseIf mvParameters.Exists("PoPaymentDueDateTo") AndAlso Not mvParameters.Exists("PoPaymentDueDateFrom") Then
        vWhereFields.Add("pop.due_date", CDate(mvParameters("PoPaymentDueDateTo").Value).Date, CDBField.FieldWhereOperators.fwoLessThanEqual)
      ElseIf mvParameters.Exists("PoPaymentDueDateFrom") AndAlso Not mvParameters.Exists("PoPaymentDueDateTo") Then
        vWhereFields.Add("pop.due_date", CDate(mvParameters("PoPaymentDueDateFrom").Value).Date, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If

      For vContactType As Contact.ContactTypes = Contact.ContactTypes.ctcContact To Contact.ContactTypes.ctcOrganisation
        If vContactType = Contact.ContactTypes.ctcContact Then
          vWhereFields.Add("pc.contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
        Else
          vWhereFields.Remove("pc.contact_type")
          vWhereFields.Add("pc.contact_type", "O")
        End If
        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vSQLAttrs.ToString, "purchase_order_payments pop", vWhereFields, "due_date", GetUnauthorisedPOPaymentJoins(vContactType)), vAttrs.ToString, ",")
      Next

      Dim vTableSort(0) As CDBDataTable.SortSpecification
      vTableSort(0).Column = "DueDate"
      pDataTable.ReOrderRowsByMultipleColumns(vTableSort)

      If vPOPPayByBacs Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          vRow.SetYNValue("PayByBacs")
        Next
      End If
    End Sub

    Public Function GetUnauthorisedPOPaymentJoins(ByVal pContactType As Contact.ContactTypes) As AnsiJoins
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("purchase_orders po", "pop.purchase_order_number", "po.purchase_order_number")
      vAnsiJoins.Add("contacts c", "c.contact_number", "po.contact_number")
      vAnsiJoins.Add("purchase_order_types pot", "po.purchase_order_type", "pot.purchase_order_type")
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        vAnsiJoins.Add("organisation_addresses ca", "pop.payee_contact_number", "ca.organisation_number", "pop.payee_address_number", "ca.address_number")
      Else
        vAnsiJoins.Add("contact_addresses ca", "pop.payee_contact_number", "ca.contact_number", "pop.payee_address_number", "ca.address_number")
      End If
      vAnsiJoins.Add("contacts pc", If(pContactType = Contact.ContactTypes.ctcOrganisation, "ca.organisation_number", "ca.contact_number"), "pc.contact_number")
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        vAnsiJoins.Add("organisations o", "pc.contact_number", "o.organisation_number")
      End If
      vAnsiJoins.Add("addresses a", "ca.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      vAnsiJoins.Add("pop_payment_methods ppm", "ppm.pop_payment_method", "pop.pop_payment_method")
      vAnsiJoins.AddLeftOuterJoin("authorisation_statuses ast", "pop.authorisation_status", "ast.authorisation_status")
      Return vAnsiJoins
    End Function

    Private Sub GetUnClaimedPayments(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetUnClaimedPayments")
      Dim vAttrs As String = "fhd.batch_number,fhd.transaction_number,fhd.line_number,transaction_date,product_desc,net_amount"
      Dim vSQL As String = "SELECT /* SQLServerCSC */ " & vAttrs
      vSQL = vSQL & " FROM declaration_lines_unclaimed dlu, financial_history_details fhd, products p, financial_history fh"
      vSQL = vSQL & " WHERE dlu.cd_number = " & mvParameters("DeclarationNumber").LongValue & " AND declaration_or_covenant_number = 'D' AND fhd.batch_number = dlu.batch_number"
      vSQL = vSQL & " AND fhd.transaction_number = dlu.transaction_number AND fhd.line_number = dlu.line_number AND p.product = fhd.product"
      vSQL = vSQL & " AND fh.batch_number = fhd.batch_number AND fh.transaction_number = fhd.transaction_number"
      vSQL = vSQL & " ORDER BY dlu.batch_number DESC, dlu.transaction_number DESC, dlu.line_number DESC"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, Replace$(vAttrs, "product_desc", "DISTINCT_PRODUCT_LINE"))

      'Select any reversal transactions created by the changing of the Declaration end date
      vAttrs = "bta.batch_number,bta.transaction_number,bta.line_number,transaction_date,product_desc,bta.amount"
      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("batch_transactions bt", "b.batch_number", "bt.batch_number")
        .Add("batch_transaction_analysis bta", "bt.batch_number", "bta.batch_number", "bt.transaction_number", "bta.transaction_number")
        .Add("financial_history_details fhd", "bta.batch_number", "fhd.batch_number", "bta.transaction_number", "fhd.transaction_number", "bta.line_number", "fhd.line_number")
        .Add("products p", "fhd.product", "p.product")
      End With

      Dim vInnerSQL As New StringBuilder
      vInnerSQL.Append("(SELECT dtcl.batch_number, dtcl.transaction_number, dtcl.line_number, dtcl.cd_number FROM declaration_tax_claim_lines dtcl")
      vInnerSQL.Append(" WHERE cd_number = " & mvParameters("DeclarationNumber").LongValue & " AND declaration_or_covenant_number = 'D') cl")
      vAnsiJoins.AddLeftOuterJoin(vInnerSQL.ToString, "fhd.batch_number", "cl.batch_number", "fhd.transaction_number", "cl.transaction_number", "fhd.line_number", "cl.line_number")

      Dim vWhereFields As New CDBFields(New CDBField("batch_type", Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment)))
      With vWhereFields
        .Add("bta.member_number", CDBField.FieldTypes.cftCharacter, "", CType(CDBField.FieldWhereOperators.fwoNotEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
        .Add("bta.member_number#2", CDBField.FieldTypes.cftCharacter, mvParameters("DeclarationNumber").Value, CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        .Add("cl.cd_number", "")
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "batches b", vWhereFields, "bta.batch_number DESC, bta.transaction_number DESC, bta.line_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("product_desc", "DISTINCT_PRODUCT_LINE"))

    End Sub
    Private Sub GetVariableParameters(ByVal pDataTable As CDBDataTable)
      'TODO Convert to new SQL Syntax
      'NYI("GetVariableParameters")
      Dim vList() As String = Nothing
      Dim vRecordSet As CDBRecordSet
      Dim vItems() As String
      Dim vIndex As Integer
      If Not mvParameters.Exists("Variables") Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT campaign, appeal, variable_parameters FROM appeals WHERE campaign = '" & mvParameters.Item("Campaign").Value & "' AND appeal = '" & mvParameters.Item("Appeal").Value & "'")
        With vRecordSet
          If .Fetch() Then
            vList = Split(.Fields("variable_parameters").Value, "|")
          End If
          .CloseRecordSet()
        End With
      Else
        vList = Split(mvParameters.Item("Variables").Value, "|")
      End If
      For vIndex = 0 To UBound(vList)
        If Len(vList(vIndex)) > 0 Then
          'Determine which character separates the variable name from the variable value.
          'The old format was: variable name+variable value.  The new format is: variable name=variable value.
          'If the variable name is not $TODAY and "=" doesn't appears in either the name or the value, then assume the old format.
          'Other assume the new format.
          Dim vSeparator As String
          If UCase$(Mid(vList(vIndex), 6)) <> "$TODAY" And InStr(vList(vIndex), "=") = 0 Then
            vSeparator = "+"
          Else
            vSeparator = "="
          End If
          vItems = Split(vList(vIndex), vSeparator)
          pDataTable.AddRowFromList(mvParameters.Item("Campaign").Value & "," & mvParameters.Item("Appeal").Value & "," & vItems(0) & "," & vItems(1))
        End If
      Next
    End Sub
    Private Sub GetDuplicateContactRecords(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "match_or_potential,contact_number_2,c2.label_name,address_number_2,contact_number_1,c1.label_name as label_name1,address_number_1"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contacts c1", "dc.contact_number_1", "c1.contact_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("contacts c2", "dc.contact_number_2", "c2.contact_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      If mvParameters.Exists("ApplicationCode") Then
        vWhereFields.Add("application_code", mvParameters("ApplicationCode").Value)
      End If
      If mvParameters.Exists("FromContactNumber") And Not mvParameters.Exists("ToContactNumber") Then
        vWhereFields.Add("contact_number_2", mvParameters("FromContactNumber").LongValue, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      If mvParameters.Exists("FromContactNumber") And mvParameters.Exists("ToContactNumber") Then
        vWhereFields.Add("contact_number_2", mvParameters("FromContactNumber").LongValue, CDBField.FieldWhereOperators.fwoBetweenFrom)
        vWhereFields.Add("contact_number_2#2", mvParameters("ToContactNumber").LongValue, CDBField.FieldWhereOperators.fwoBetweenTo)
      End If
      If mvParameters.Exists("RecordType") Then
        If mvParameters("RecordType").Value <> "-" Then
          vWhereFields.Add("match_or_potential", mvParameters("RecordType").Value)
        End If
      End If
      If mvParameters.Exists("RunDate") Then
        vWhereFields.Add("run_date", mvParameters("RunDate").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "duplicate_contacts dc", vWhereFields, "match_or_potential,c2.surname,c2.forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields.Replace("c1.label_name as ", ""))
    End Sub
    Private Sub GetMeetingRecords(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "m.meeting_number,meeting_date,meeting_date as meeting_time,meeting_desc,meeting_type_desc,meeting_location_desc"
      'vSQL = Replace$(vSQL, "MEETING_DATE", "meeting_date")
      'vSQL = Replace$(vSQL, "MEETING_TIME", "meeting_date" & mvEnv.Connection.DBAs & "meeting_time")

      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields
      Dim vDate As Date
      Dim vOrderBy As String

      If mvParameters.Exists("Checkbox") OrElse mvParameters.Exists("Checkbox2") OrElse mvParameters.Exists("ContactNumber") OrElse mvParameters.Exists("ContactNumber2") Then
        vAnsiJoins.Add("meeting_links ml", "ml.meeting_number", "m.meeting_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
        If mvParameters.Exists("Checkbox") OrElse mvParameters.Exists("Checkbox2") Then
          vWhereFields.Add("ml.contact_number", CDBField.FieldTypes.cftInteger, mvEnv.User.ContactNumber)
          vWhereFields.Add("link_type", CDBField.FieldTypes.cftCharacter, "W")

          If mvParameters.Exists("Checkbox2") Then vWhereFields.Add("notified", CDBField.FieldTypes.cftCharacter, "N") '  vWhere = vWhere & " AND notified = 'N'"
          If mvParameters.Exists("Checkbox") Then vWhereFields.Add("attended", "") '
        ElseIf mvParameters.Exists("ContactNumber") Then
          vWhereFields.Add("ml.contact_number", mvParameters("ContactNumber").IntegerValue)
          vWhereFields.Add("link_type", "W")
        ElseIf mvParameters.Exists("ContactNumber2") Then
          vWhereFields.Add("ml.contact_number", mvParameters("ContactNumber2").IntegerValue)
          vWhereFields.Add("link_type", "R")
        End If
      Else
        If mvParameters.Exists("MeetingNumber") Then vWhereFields.Add("meeting_number", mvParameters("MeetingNumber").IntegerValue)
        If mvParameters.Exists("MeetingDesc") Then vWhereFields.Add("meeting_desc", mvParameters("MeetingDesc").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("MeetingType") Then vWhereFields.Add("m.meeting_type", mvParameters("MeetingType").Value)
        If mvParameters.Exists("MeetingLocation") Then vWhereFields.Add("m.meeting_location", mvParameters("MeetingLocation").Value)

        If mvParameters.Exists("MeetingDate") Then      'Meeting Date retained in case database not yet updated
          vDate = CDate(mvParameters("MeetingDate").Value)
          vWhereFields.Add("meeting_dat e", CDBField.FieldTypes.cftTime, vDate.AddDays(-1) & " 23:59", CDBField.FieldWhereOperators.fwoGreaterThan)
          vWhereFields.Add("meeting_date#2", CDBField.FieldTypes.cftTime, vDate.AddDays(1) & " 00:00", CDBField.FieldWhereOperators.fwoLessThan)
        ElseIf mvParameters.Exists("Date") Then
          vDate = CDate(mvParameters("Date").Value)
          vWhereFields.Add("meeting_date", CDBField.FieldTypes.cftTime, vDate & " 00:00", CDBField.FieldWhereOperators.fwoGreaterThan)
          If mvParameters.Exists("Date2") Then
            vDate = CDate(mvParameters("Date2").Value)
            vWhereFields.Add("meeting_date#2", CDBField.FieldTypes.cftTime, vDate & " 23:59", CDBField.FieldWhereOperators.fwoLessThan)
          End If
        ElseIf mvParameters.Exists("Date2") Then
          vDate = CDate(mvParameters("Date2").Value)
          vWhereFields.Add("meeting_date", CDBField.FieldTypes.cftTime, vDate.AddDays(1) & " 00:00", CDBField.FieldWhereOperators.fwoLessThan)
        ElseIf mvParameters.Exists("Checkbox3") Then
          vWhereFields.Add("meeting_date", CDBField.FieldTypes.cftTime, Date.Today.AddMonths(-1).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThan)
          vWhereFields.Add("meeting_date#2", CDBField.FieldTypes.cftTime, Date.Today.AddMonths(1).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThan)
        End If
      End If
      vAnsiJoins.Add("meeting_types mt", "mt.meeting_type", "m.meeting_type", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("meeting_locations mlo", "mlo.meeting_location", "m.meeting_location", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vOrderBy = "meeting_date DESC"
      Dim vSQLStatement As SQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, "meetings m", vWhereFields, vOrderBy, vAnsiJoins)
      vAttrs = vAttrs.Replace("meeting_date as ", "")
      vAttrs = vAttrs.Replace("m.meeting_number", "meeting_number")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs) ' vAttrs, "", True)
    End Sub

    Private Sub SetCPDDurations(ByVal pDataTable As CDBDataTable)
      Dim vRow As CDBDataRow
      Dim vStart As Date
      Dim vEnd As Date

      For Each vRow In pDataTable.Rows
        'Set the cycle Duration
        If pDataTable.Columns.ContainsKey("CycleDuration") _
          Or pDataTable.Columns.ContainsKey("StartDate") _
          Or pDataTable.Columns.ContainsKey("EndDate") Then
          vStart = Date.Parse(vRow.Item("StartDate").ToString)
          vEnd = Date.Parse(vRow.Item("EndDate").ToString)
          If pDataTable.Columns.ContainsKey("StartDate") Then vRow.Item("StartDate") = vStart.Year.ToString
          If pDataTable.Columns.ContainsKey("EndDate") Then vRow.Item("EndDate") = vEnd.Year.ToString
        End If
      Next
    End Sub
    Private Sub GetBatchProcessingInformation(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As New StringBuilder
      vFieldNames.Append("b.batch_number,b.batch_type,b.bank_account,batch_date,number_of_entries,batch_total,number_of_transactions,transaction_total,detail_completed,")
      vFieldNames.Append("ready_for_banking,paying_in_slip_printed,paying_in_slip_number,picked,posted_to_cash_book,posted_to_nominal,currency_batch_total,")
      vFieldNames.Append("currency_transaction_total,currency_exchange_rate,")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vFieldNames.Append("b.currency_code,") Else vFieldNames.Append("'',")
      vFieldNames.Append("b.job_number, payment_method, batch_category, provisional, claim_sent,")
      vFieldNames.Append("transaction_type,b.product,rate,source,cash_book_batch,journal_number,b.amended_by,b.amended_on,balanced_by,balanced_on,posted_by,")
      vFieldNames.Append("posted_on,contents_amended_by,contents_amended_on,header_amended_by,header_amended_on,batch_created_by,batch_created_on,")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBankingDate) Then vFieldNames.Append("banking_date,") Else vFieldNames.Append("'',")
      vFieldNames.Append(" batch_analysis_code,print_cheque_list,company,picking_list_number,u.department,d.department_desc")

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("posted_to_nominal", "N")
      vWhereFields.Add("b.batch_number", "SELECT batch_number FROM open_batches ob WHERE ob.batch_type IS NOT NULL", CDBField.FieldWhereOperators.fwoNotIn)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("batch_types bt", "bt.batch_type", "b.batch_type")
      vAnsiJoins.Add("users u", "b.batch_created_by", "u.logname")
      vAnsiJoins.Add("departments d", "u.department", "d.department")
      If mvEnv.GetConfig("opt_batch_per_user") = "DEPARTMENT" AndAlso mvEnv.GetConfigOption("opt_batch_ownership") Then
        vWhereFields.Add("d.department", mvEnv.User.Department)
      End If
      vAnsiJoins.AddLeftOuterJoin("bank_accounts ba", "b.bank_account", "ba.bank_account")
      vAnsiJoins.AddLeftOuterJoin("(SELECT batch_number, MAX(picking_list_number) AS picking_list_number FROM issued_stock GROUP BY batch_number) si", "b.batch_number", "si.batch_number")

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFieldNames.ToString, "batches b", vWhereFields, "batch_date DESC, b.batch_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "", ",,")
      If pDataTable.Rows.Count > 0 Then
        vWhereFields.Clear()
        vWhereFields.Add("cash_batch_number", CStr(If(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCashBookBatchLimit) = String.Empty, 999, CDbl(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCashBookBatchLimit))) - 50), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        Dim vRS As CDBRecordSet = New SQLStatement(mvEnv.Connection, "bank_account", "bank_accounts", vWhereFields).GetRecordSet
        Dim vBankAccounts As New StringBuilder
        While vRS.Fetch
          If vBankAccounts.Length > 0 Then vBankAccounts.Append(",")
          vBankAccounts.Append(vRS.Fields(1).Value)
        End While
        pDataTable.Rows(0).Item("BankAccounts") = vBankAccounts.ToString
      End If
    End Sub

    Private Sub GetPickingListDetails(ByVal pDataTable As CDBDataTable)
      Dim vFieldNames As New StringBuilder
      vFieldNames.Append("pl.picking_list_number,product,quantity,shortfall,shortfall AS original_shortfall,confirmed_on,pld.warehouse,warehouse_desc")
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("pl.picking_list_number", mvParameters("PickingListNumber").LongValue)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("picking_list_details pld", "pl.picking_list_number", "pld.picking_list_number")
      vAnsiJoins.Add("warehouses w", "pld.warehouse", "w.warehouse")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFieldNames.ToString, "picking_lists pl", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "", ",")
    End Sub

    Private Sub GetCampaignCosts(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignItemisedCosts) Then
        Dim vWhereFields As New CDBFields
        With vWhereFields
          .Add("campaign", mvParameters.ParameterExists("Campaign").Value)
          .Add("appeal", mvParameters.ParameterExists("Appeal").Value)
          If mvParameters.Exists("CollectionNumber") Then
            .Add("collection_number", mvParameters("CollectionNumber").LongValue)
          Else
            .Add("segment_collection", mvParameters.ParameterExists("Segment").Value)
          End If
          If mvParameters.Exists("CampaignCostNumber") Then .Add("campaign_cost_number", mvParameters("CampaignCostNumber").LongValue)
        End With
        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, New CampaignItemisedCost(mvEnv).GetRecordSetFields, "campaign_itemised_costs cic", vWhereFields))
      End If
    End Sub

    Private Sub GetEventBookingTransactions(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventMultipleAnalysis) Then
        Dim vAttrs As String = "bta.batch_number,bta.transaction_number,bta.line_number,bt.transaction_date,product,rate,transaction_type_desc,payment_method,bta.distribution_code,quantity,bta.amount,vat_amount,vat_rate,source,bta.currency_amount,currency_vat_amount,bta.notes,tt.transaction_sign"
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("batch_transaction_analysis bta", "bta.batch_number", "ebt.batch_number", "bta.transaction_number", "ebt.transaction_number", "bta.line_number", "ebt.line_number")
        vAnsiJoins.Add("batch_transactions bt", "bta.batch_number", "bt.batch_number", "bta.transaction_number", "bt.transaction_number")
        vAnsiJoins.AddLeftOuterJoin("transaction_types tt", "bt.transaction_type", "tt.transaction_type")
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("ebt.booking_number", CDBField.FieldTypes.cftLong, mvParameters("BookingNumber").Value)
        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vAttrs, "event_booking_transactions ebt", vWhereFields, "", vAnsiJoins))

        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("TransactionSign") = "D" Then
            If vRow.Item("Amount").Length > 0 Then vRow.Item("Amount") = (CDbl(vRow.Item("Amount")) * -1).ToString
            vRow.Item("Quantity") = (IntegerValue(vRow.Item("Quantity")) * -1).ToString
            If vRow.Item("CurrencyAmount").Length > 0 Then vRow.Item("CurrencyAmount") = (CDbl(vRow.Item("CurrencyAmount")) * -1).ToString
          End If
        Next

        Dim vSS(3) As CDBDataTable.SortSpecification
        vSS(0).Column = "TransactionDate"
        vSS(0).Descending = True
        vSS(1).Column = "BatchNumber"
        vSS(1).Descending = True
        vSS(2).Column = "TransactionNumber"
        vSS(2).Descending = True
        vSS(3).Column = "LineNumber"
        vSS(3).Descending = True
        pDataTable.ReOrderRowsByMultipleColumns(vSS)
      End If
    End Sub

    Private Sub GetContactEmailingsLinks(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCheetahMail) Then
        Dim vAttrs As String = mvEnv.Connection.DBSpecialCol(Nothing, "c.number") & ",c.valid_from,c.valid_to,c.is_active,cel.email_link,cel.email_link_name,cel.clicked_datetime,cel.communication_number"
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("mailing_history mh", "ce.mailing_number", "mh.mailing_number")
        vAnsiJoins.Add("mailings m", "mh.mailing", "m.mailing")
        vAnsiJoins.Add("contact_emailings_links cel", "ce.contact_number", "cel.contact_number", "ce.address_number", "cel.address_number", "ce.mailing_number", "cel.mailing_number")
        vAnsiJoins.Add("communications c", "cel.communication_number", "c.communication_number")
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("ce.contact_number", CDBField.FieldTypes.cftLong, mvParameters("ContactNumber").LongValue)
        vWhereFields.Add("ce.address_number", CDBField.FieldTypes.cftLong, mvParameters("AddressNumber").LongValue)
        vWhereFields.Add("ce.communication_number", CDBField.FieldTypes.cftLong, mvParameters("CommunicationNumber").LongValue)
        vWhereFields.Add("ce.mailing_number", CDBField.FieldTypes.cftLong, mvParameters("MailingNumber").LongValue)
        vWhereFields.Add("cel.communication_number", CDBField.FieldTypes.cftLong, "ce.communication_number")
        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vAttrs, "contact_emailings ce", vWhereFields, "mh.mailing_date DESC", vAnsiJoins))
      End If
    End Sub

    Private Sub GetPurchaseInvoiceChequeInformation(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataChequeReissue) Then
        Dim vContact As New Contact(mvEnv)
        Dim vAddress As New Address(mvEnv)
        Dim vAttrs As String = "cheque_reference_number,cheque_number,amount,printed_on,reconciled_on,ch.cheque_status,cheque_status_desc,allow_reissue,reprint_count,ch.contact_number AS payee_contact_number,ch.address_number AS payee_address_number,c.label_name AS payee_contact_label_name," & vContact.GetRecordSetFieldsName & "," & vAddress.GetRecordSetFieldsCountry & ",ch.currency_code,currency_code_desc"

        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("contacts c", "ch.contact_number", "c.contact_number")
        vAnsiJoins.Add("addresses a", "ch.address_number", "a.address_number")
        vAnsiJoins.Add("countries co", "a.country", "co.country")
        Dim vCurrencyJoin As New AnsiJoin("currency_codes cc", "ch.currency_code", "cc.currency_code")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vAnsiJoins.Add(vCurrencyJoin)
        vAnsiJoins.AddLeftOuterJoin("cheque_statuses cs", "ch.cheque_status", "cs.cheque_status")
        vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "ch.cancellation_reason", "cr.cancellation_reason")
        vAnsiJoins.AddLeftOuterJoin("sources ss", "ch.cancellation_source", "ss.source")
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("cheque_reference_number", CDBField.FieldTypes.cftLong, mvParameters("ChequeReferenceNumber").LongValue)
        If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) Then vAttrs = vAttrs.Replace("allow_reissue", "")
        If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode) Then vAttrs = vAttrs.Replace("ch.currency_code,currency_code_desc", ",")
        Dim vItems As String = ",cheque_number,amount,printed_on,reconciled_on,cheque_status,cheque_status_desc,allow_reissue,"

        'BR17340
        vAttrs = vAttrs & ",ch.adjustment_status,ch.cancellation_reason,ch.cancellation_source,ch.cancelled_by,ch.cancelled_on,ss.source_desc AS cancellation_source_desc,cancellation_reason_desc"
        ' vItems = vItems & ",adjustment_status,cancellation_reason,cancellation_source,cancelled_by,cancelled_on"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPopPaymentMethod) Then
          vAttrs = vAttrs & ",ch.pop_payment_method"
        Else
          vAttrs = vAttrs & ","
        End If

        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "cheques ch", vWhereFields, "", vAnsiJoins), "cheque_reference_number" & vItems & "reprint_count,payee_contact_number,payee_address_number,payee_contact_label_name,CONTACT_NAME,ADDRESS_LINE,currency_code,currency_code_desc,adjustment_status,cancellation_reason,cancellation_source,cancelled_by,cancelled_on,cancellation_source_desc,cancellation_reason_desc,pop_payment_method")
        vAttrs = vAttrs.Replace("ch.currency_code,currency_code_desc", ",")
        vAttrs = vAttrs.Replace("ch.pop_payment_method", ",")
        If vAnsiJoins.Contains(vCurrencyJoin) Then vAnsiJoins.Remove(vCurrencyJoin)
        pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "cheque_history ch", vWhereFields, "reprint_count DESC", vAnsiJoins), vItems & ",payee_contact_number,payee_address_number,payee_contact_label_name,CONTACT_NAME,ADDRESS_LINE,,,,,,,,,,")
        For Each vRow As CDBDataRow In pDataTable.Rows
          vRow.SetYNValue("AllowReissue")
        Next
      End If
    End Sub
    Private Sub GetContactCommunicationHistory(ByVal pDataTable As CDBDataTable)
      Dim vMailingDataTable As CDBDataTable = New CDBDataTable()
      Dim vDocumentTable As CDBDataTable = New CDBDataTable()

      Dim vMailingDataSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactMailings, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)
      Dim vDocumentDataSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactDocuments, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)

      vMailingDataTable.AddColumnsFromList(vMailingDataSelection.mvResultColumns)
      vDocumentTable.AddColumnsFromList(vDocumentDataSelection.mvResultColumns)

      GetContactMailings(vMailingDataTable)
      mvType = DataSelectionTypes.dstContactDocuments
      GetDocuments(vDocumentTable, mvParameters.OptionalValue("IncludeEmailDocSource", "N") = "Y")
      mvType = DataSelectionTypes.dstContactCommunicationHistory

      If vDocumentTable.Rows.Count > 0 Then
        For Each vRow As CDBDataRow In vDocumentTable.Rows
          Dim vCommHistory As CDBDataRow = New CDBDataRow(pDataTable.Columns, 0)
          For Each vColumn As CDBDataColumn In vDocumentTable.Columns
            Select Case vColumn.Name
              Case "Dated"
                vCommHistory.Item("Date") = vRow.Item(vColumn.Name)
                pDataTable.Columns("Date").FieldType = vColumn.FieldType
              Case "DocumentNumber"
                vCommHistory.Item("MailingNumber") = vRow.Item(vColumn.Name)
                pDataTable.Columns("MailingNumber").FieldType = vColumn.FieldType
              Case "LabelName"
                vCommHistory.Item("ContactName") = vRow.Item(vColumn.Name)
                pDataTable.Columns("ContactName").FieldType = vColumn.FieldType
              Case "CreatedBy"
                vCommHistory.Item("MailedBy") = vRow.Item(vColumn.Name)
                pDataTable.Columns("MailedBy").FieldType = vColumn.FieldType
              Case "Subject"
                vCommHistory.Item("Description") = vRow.Item(vColumn.Name)
                pDataTable.Columns("Description").FieldType = vColumn.FieldType
              Case "Direction"
                vCommHistory.Item("Type") = vRow.Item(vColumn.Name)
                pDataTable.Columns("Description").FieldType = vColumn.FieldType
              Case Else
                vCommHistory.Item(vColumn.Name) = vRow.Item(vColumn.Name)
                pDataTable.Columns(vColumn.Name).FieldType = vColumn.FieldType
            End Select
          Next
          pDataTable.Rows.Add(vCommHistory)
        Next
      End If

      If vMailingDataTable.Rows.Count > 0 Then
        For Each vRow As CDBDataRow In vMailingDataTable.Rows
          Dim vCommHistory As CDBDataRow = New CDBDataRow(pDataTable.Columns, 0)
          For Each vColumn As CDBDataColumn In vMailingDataTable.Columns
            vCommHistory.Item(vColumn.Name) = vRow.Item(vColumn.Name)
            pDataTable.Columns(vColumn.Name).FieldType = vColumn.FieldType
          Next
          pDataTable.Rows.Add(vCommHistory)
        Next
      End If

      Dim vTableSort(0) As CDBDataTable.SortSpecification
      vTableSort(0).Column = "Date"
      vTableSort(0).Descending = True
      pDataTable.ReOrderRowsByMultipleColumns(vTableSort)
    End Sub
    Private Sub GetContactAddressAndUsage(ByVal pDataTable As CDBDataTable)
      Dim vAddressDataTable As CDBDataTable = New CDBDataTable()
      Dim vUsageDataTable As CDBDataTable = New CDBDataTable()

      Dim vAddressSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactAddresses, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)

      vAddressDataTable.AddColumnsFromList(vAddressSelection.mvResultColumns)

      GetContactAddresses(vAddressDataTable)

      If vAddressDataTable.Rows.Count > 0 Then
        For Each vRow As CDBDataRow In vAddressDataTable.Rows
          Dim vAddressUsage As CDBDataRow = New CDBDataRow(pDataTable.Columns, 0)
          Dim vParams As New CDBParameters
          mvParameters("AddressNumber").Value = vRow.Item("AddressNumber")
          Dim vUsageSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactAddressUsages, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)
          mvType = DataSelectionTypes.dstContactAddressUsages
          vUsageDataTable.AddColumnsFromList(vUsageSelection.mvResultColumns)
          GetContactAddressUsages(vUsageDataTable)
          For Each vColumn As CDBDataColumn In vUsageDataTable.Columns
            For Each vUsage As CDBDataRow In vUsageDataTable.Rows
              Select Case vColumn.Name
                Case "AddressUsage"
                  vAddressUsage.Item("AddressUsage") = vUsage.Item(vColumn.Name)
                Case "AddressUsageDesc"
                  vAddressUsage.Item("AddressUsageDesc") = vUsage.Item(vColumn.Name)
              End Select
              pDataTable.Rows.Add(vAddressUsage)
            Next
          Next
        Next
      End If

      mvType = DataSelectionTypes.dstContactAddressAndUsage

      If vAddressDataTable.Rows.Count > 0 Then
        For Each vRow As CDBDataRow In vAddressDataTable.Rows
          Dim vAddressAndUsage As CDBDataRow = New CDBDataRow(pDataTable.Columns, 0)
          For Each vColumn As CDBDataColumn In vAddressDataTable.Columns
            Select Case vColumn.Name
              Case "AddressNumber", "AddressLine", "Historical", "AddressUsage", "AddressUsageDesc"
                vAddressAndUsage.Item(vColumn.Name) = vRow.Item(vColumn.Name)
            End Select
          Next
          pDataTable.Rows.Add(vAddressAndUsage)
        Next
      End If
    End Sub

    Private Sub GetGeneralMailingSelectionSets(ByVal pDataTable As CDBDataTable)
      'SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source
      Dim vFields As String = "selection_set,selection_set_desc,user_name,department,number_in_set,source,custom_data,attribute_captions"
      Dim vWhereFields As New CDBFields()
      If mvParameters.Exists("SelectionSetNumber") AndAlso mvParameters("SelectionSetNumber").Value <> "" Then
        vWhereFields.Add("selection_set", mvParameters("SelectionSetNumber").Value, CDBField.FieldWhereOperators.fwoEqual)
      Else
        If mvParameters.Exists("MailingType") AndAlso mvParameters("MailingType").Value <> "" Then vWhereFields.Add("selection_group", mvParameters("MailingType").Value, CDBField.FieldWhereOperators.fwoEqual)

        vWhereFields.Add("department#1", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoCloseBracket)
        vWhereFields.Add("user_name#1", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)

        If mvParameters.Exists("UserName") AndAlso mvParameters("UserName").Value <> "" Then vWhereFields.Add("user_name#2", mvParameters("UserName").Value, CDBField.FieldWhereOperators.fwoEqual)
        If mvParameters.Exists("SelectionSetDesc") AndAlso mvParameters("SelectionSetDesc").Value <> "" Then vWhereFields.Add("selection_set_desc", mvParameters("SelectionSetDesc").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Department") AndAlso mvParameters("Department").Value <> "" Then vWhereFields.Add("department#2", mvParameters("Department").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Source") AndAlso mvParameters("Source").Value <> "" Then vWhereFields.Add("source", mvParameters("Source").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "selection_sets", vWhereFields, "selection_set_desc")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetDespatchNotes(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "dn.despatch_note_number,despatch_method_desc,despatch_date,delivery_charge,invoice_number,order_date,carrier_reference,dn.despatch_method,a.address_number,address,house_name,town,county,postcode,branch,a.country,address_type,building_number,sortcode,uk,country_desc,c.contact_number,title,forenames,initials,surname,honorifics,salutation,label_name,preferred_forename,contact_type,ni_number,prefix_honorifics,surname_prefix,informal_salutation"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("despatch_methods dm", "dn.despatch_method", "dm.despatch_method", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("contacts c", "dn.contact_number", "c.contact_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("addresses a", "dn.address_number", "a.address_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("countries co", "a.country", "co.country", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      AddWhereFieldFromIntegerParameter(vWhereFields, "PickingListNumber", "picking_list_number")
      If mvParameters.HasValue("BatchNumber") Then
        'Here we need to find the despatch note that relates to a transaction but the transaction info is on the issued stock record not the despatch note (back orders?)
        vAnsiJoins.Add("issued_stock st", "dn.picking_list_number", "st.picking_list_number", "dn.despatch_note_number", "st.despatch_note_number")
        AddWhereFieldFromIntegerParameter(vWhereFields, "BatchNumber", "st.batch_number")
        AddWhereFieldFromIntegerParameter(vWhereFields, "TransactionNumber", "st.transaction_number")
        AddWhereFieldFromIntegerParameter(vWhereFields, "LineNumber", "st.line_number")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "despatch_notes dn", vWhereFields, "dn.despatch_note_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields, "ADDRESS_LINE,CONTACT_NAME")
    End Sub

    Private Sub GetContactTokens(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "contact_number,card_token_number,card_token_desc,card_number"

      Dim vSqlServer As New SQLStatement(mvEnv.Connection, vFields, "contact_card_tokens", New CDBField("contact_number", mvParameters("ContactNumber").Value))
      pDataTable.FillFromSQL(mvEnv, vSqlServer, vFields)
    End Sub


    'Private Sub AddOwnerRestriction(ByRef pSQL As String)
    'BR12727 Added during part conversion of GetContactEventBookings
    '  'NOTE: Use of this sub assumes Event Table Instance is "e"
    '  If mvEnv.GetConfigOption("ev_display_owned_bookings_only") Then
    '    pSQL = pSQL & " AND (e.department IS NULL OR e.department = '" & mvEnv.User.Department & "' OR e.event_number IN (SELECT event_number FROM event_owners ewn WHERE ewn.event_number = e.event_number AND ewn.department = '" & mvEnv.User.Department & "'))"
    '  End If
    'End Sub
    Private Sub FillAddressData(ByVal pDataTable As CDBDataTable)
      FillAddressData(pDataTable, False)
    End Sub
    Private Sub FillAddressData(ByVal pDataTable As CDBDataTable, ByVal pGetOrganisation As Boolean)
      Dim vAddressNumber As Integer
      Dim vList As New CDBParameters
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("AddressLine").Length = 0 Then
          vAddressNumber = IntegerValue(vRow.Item("AddressNumber"))
          If vAddressNumber > 0 AndAlso Not vList.Exists(vAddressNumber.ToString) Then vList.Add(vAddressNumber.ToString)
        End If
      Next
      If vList.Count > 0 Then
        Dim vAddress As New Address(mvEnv)
        vAddress.Init()
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("address_number", vList.ItemList, CDBField.FieldWhereOperators.fwoIn)
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("countries co", "a.country", "co.country")
        Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, vAddress.GetRecordSetFieldsCountry, "addresses a", vWhereFields, "", vAnsiJoins).GetRecordSet
        While vRecordSet.Fetch
          vAddress.InitFromRecordSetCountry(vRecordSet)
          For Each vRow As CDBDataRow In pDataTable.Rows
            vAddressNumber = IntegerValue(vRow.Item("AddressNumber"))
            If vRow.Item("AddressLine").Length = 0 AndAlso vAddressNumber = vAddress.AddressNumber Then
              vRow.Item("AddressLine") = vAddress.AddressLine
              If pGetOrganisation Then vRow.Item("OrganisationName") = "" 'vAddress.OrganisationName
            End If
          Next
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    Private Sub GetAwaitListConfirmation(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "pl.picking_list_number,company_desc,pld.warehouse,w.warehouse_desc,created_on"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("batches b", "st.batch_number", "b.batch_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("picking_list_details pld", "pld.picking_list_number", "st.picking_list_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("picking_lists pl", "pl.picking_list_number", "pld.picking_list_number", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("companies cm", "cm.company", "pl.company", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("warehouses w", "w.warehouse", "pld.warehouse", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vWhereFields.Add("pl.confirmed_by", "", CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("b.picked ", "'Y','P'", CDBField.FieldWhereOperators.fwoIn)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "issued_stock st", vWhereFields, "pld.warehouse,company_desc,pl.picking_list_number", vAnsiJoins)
      vSQLStatement.Distinct = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetPackedProductDataSheet(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "pp.link_product, p.product_desc,pw.warehouse, w.warehouse_desc, pw.last_stock_count, p.cost_of_sale, p.warehouse"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("products p", "p.product", "pp.link_product", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("product_warehouses pw", "pw.product", "pp.link_product", Data.AnsiJoin.AnsiJoinTypes.InnerJoin)
      vAnsiJoins.Add("warehouses w", "w.warehouse", "pw.warehouse")
      vWhereFields.Add("pp.product", mvParameters("Product").Value)
      vWhereFields.Add("pp.rate", mvParameters("Rate").Value)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields & " AS default_warehouse", "packed_products pp", vWhereFields, "product_desc, w.warehouse_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetReportData(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      vWhereFields.Clear()
      Dim vFields As String = "report_number,report_name,report_code,client,header,footer,mail_merge_output,file_output,detail_exclusive,landscape,use_ssrs,mailmerge_header,application_name,amended_on,amended_by"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbReportUseSsrs) = False Then vFields = vFields.Replace("use_ssrs,", ",")
      vSQL = New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "reports", vWhereFields, "report_number")
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields)
    End Sub
    Private Sub GetReportParameters(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("report_number", mvParameters("ReportNumber").LongValue)
      vSQL = New SQLStatement(mvEnv.Connection, "parameter_number,parameter_desc,parameter_name,field_type,parameter_value,rp.expression,rp.amended_on,rp.amended_by", "report_parameters rp", vWhereFields, "parameter_number")
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetReportVersion(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("report_number", mvParameters("ReportNumber").LongValue)
      vSQL = New SQLStatement(mvEnv.Connection, "version_number,change_description,change_date,logname,rvh.amended_on,rvh.amended_by", "report_version_history rvh", vWhereFields, "version_number")
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetReportControl(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("fp_application", mvParameters("ApplicationNumber").LongValue)
      vWhereFields.Add("fp_page_type", mvParameters("PageType").Value)
      vSQL = New SQLStatement(mvEnv.Connection, "sequence_number,control_type,table_name,attribute_name,control_top,control_left,control_width,control_height,visible,resource_id,control_caption,caption_width,help_text,contact_group,parameter_name,mandatory_item,readonly_item,default_value", "fp_controls fc", vWhereFields, "fc.sequence_number")
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetReportSectionDetail(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("report_item_types rit", "ri.report_item_type", "rit.report_item_type")
      vWhereFields.Add("report_number", mvParameters("ReportNumber").LongValue)
      vWhereFields.Add("section_number", mvParameters("SectionNumber").LongValue)
      vSQL = New SQLStatement(mvEnv.Connection, "item_number,report_item_type_desc,caption,attribute_name,parameter_name,item_format,item_alignment,item_width,item_newline,suppress_blanks,ri.amended_on,ri.amended_by,ri.report_item_type", "report_items ri", vWhereFields, "item_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    Private Sub GetReportSectionData(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vSQL As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("report_sections rs", "r.report_number", "rs.report_number")
      vAnsiJoins.Add("section_types st", "st.section_type", "rs.section_type")
      vWhereFields.Clear()
      vSQL = New SQLStatement(mvEnv.Connection, "rs.report_number,section_number,section_name,section_type_desc,table_flags,suppress_output,control_attributes,exclusive_attributes,section_sql,rs.amended_on,rs.amended_by,rs.section_type", "reports r", vWhereFields, "r.report_number,section_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetOwnershipUserInformation(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "ogu.logname,al.ownership_access_level,al.ownership_access_level_desc,ogu.valid_from,ogu.valid_to,ogu.amended_on,ogu.amended_by"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vOrder As String

      vAnsiJoins.Add("ownership_access_levels al", "ogu.ownership_access_level", "al.ownership_access_level")
      vWhereFields.Add("ogu.ownership_group", mvParameters("OwnershipGroup").Value, CDBField.FieldWhereOperators.fwoEqual)
      If mvParameters.Exists("Logname") Then
        vWhereFields.Add("ogu.logname", mvParameters("Logname").Value, CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("ogu.valid_from", CDBField.FieldTypes.cftDate, String.Empty, CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("ogu.valid_from#2", CDBField.FieldTypes.cftDate, TodaysDate, CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoLessThanEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        vWhereFields.Add("ogu.valid_to", CDBField.FieldTypes.cftDate, String.Empty, CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("ogu.valid_to#2", CDBField.FieldTypes.cftDate, TodaysDate, CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoGreaterThanEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        vOrder = "ogu.valid_from DESC"
      Else
        vWhereFields.Add("ogu.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
        vWhereFields.Add("ogu.valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("ogu.valid_to#2", CDBField.FieldTypes.cftDate, TodaysDate, CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoGreaterThanEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
        vOrder = "ogu.ownership_group,ogu.logname"
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "ownership_group_users ogu", vWhereFields, vOrder, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetOwnershipUsers(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "ogu.ownership_group,ogu.logname,u.full_name"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("users u", "ogu.logname", "u.logname")
      vWhereFields.Add("ogu.ownership_group", mvParameters("OwnershipGroup").Value, CDBField.FieldWhereOperators.fwoEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "ownership_group_users ogu", vWhereFields, "ogu.ownership_group,ogu.logname", vAnsiJoins)
      vSQLStatement.Distinct = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetOwnershipGroupInformation(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "og.ownership_group,og.ownership_group_desc,d.department,d.department_desc,og.principal_department_logname,c.label_name,og.read_access_text,og.view_access_text,og.notes,og.amended_on,og.amended_by"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("departments d", "og.principal_department", "d.department")
      vAnsiJoins.AddLeftOuterJoin("users u", "og.principal_department_logname", "u.logname")
      vAnsiJoins.AddLeftOuterJoin("contacts c", "u.contact_number", "c.contact_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "ownership_groups og", New CDBFields(), "ownership_group_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetOwnershipDepartmentInformation(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "dod.department,d.department_desc,al.ownership_access_level_desc,al.ownership_access_level"
      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields As New CDBFields()
      vAnsiJoins.Add("ownership_access_levels al", "dod.ownership_access_level", "al.ownership_access_level")
      vAnsiJoins.Add("departments d", "dod.department", "d.department")
      vWhereFields.Add("dod.ownership_group ", mvParameters("OwnershipGroup").Value, CDBField.FieldWhereOperators.fwoEqual)
      If mvParameters.Exists("Department") Then
        vWhereFields.Add("dod.department", mvParameters("Department").Value, CDBField.FieldWhereOperators.fwoEqual)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "department_ownership_defaults dod", vWhereFields, "d.department_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub
    Private Sub GetEmailAutoReplyText(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "stp.paragraph_text"
      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields As New CDBFields()
      Dim vParagraphText As String = String.Empty
      Dim vDataTable As New CDBDataTable
      Dim vSubTopic As SubTopic
      Dim vResultRow As CDBDataRow = Nothing
      vAnsiJoins.Add("communications_log_subjects cls", "stp.sub_topic", "cls.sub_topic")
      vWhereFields.Add("cls.communications_log_number", CDBField.FieldTypes.cftInteger, mvParameters("DocumentNumber").Value)
      vDataTable.AddColumnsFromList(mvResultColumns)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "sub_topic_paragraphs stp", vWhereFields, "", vAnsiJoins)
      vDataTable.FillFromSQL(mvEnv, vSQLStatement, True)
      'Appending all the Paragraph Texts
      For Each vRow As CDBDataRow In vDataTable.Rows
        If vRow.Item("ParagraphText").Length > 0 Then
          If vParagraphText.Length > 0 Then vParagraphText = vParagraphText & "\r\n\r\n"
          vParagraphText = vParagraphText & vRow.Item("ParagraphText")
        End If
      Next
      If vParagraphText.Length > 0 Then
        vSubTopic = New SubTopic
        With vSubTopic
          'Adding Email Reply Topic
          .Init(mvEnv, mvEnv.GetConfig("email_reply_topic"), mvEnv.GetConfig("email_reply_sub_topic"))
          If .Existing Then
            If .ParagraphText IsNot Nothing AndAlso .ParagraphText.Length > 0 Then vParagraphText = .ParagraphText & "\r\n\r\n" & vParagraphText
          Else
            RaiseError(DataAccessErrors.daeUndefinedEmailReplyTopicSubTopic)
          End If
          'Adding Email Signoff Topic
          .Init(mvEnv, mvEnv.GetConfig("email_signoff_topic"), mvEnv.GetConfig("email_signoff_sub_topic"))
          If .Existing Then
            If .ParagraphText IsNot Nothing AndAlso .ParagraphText.Length > 0 Then vParagraphText = vParagraphText & "\r\n\r\n" & .ParagraphText
          Else
            RaiseError(DataAccessErrors.daeUndefinedEmailSignOffTopicSubTopic)
          End If
        End With
      End If
      vResultRow = pDataTable.AddRow()
      vResultRow.Item("ParagraphText") = vParagraphText
    End Sub

    Private Sub GetEventActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "e.master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,,,,,,,,,,,duration_days,duration_hours,duration_minutes,a.document_class,action_text"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("e.event_number", mvParameters("EventNumber").IntegerValue)
      If mvParameters.HasValue("ActionNumber") Then vWhereFields.Add("a.action_number", mvParameters("ActionNumber").IntegerValue)
      vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("u.department", mvEnv.User.Department)
      vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("u.department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("actions a", "e.master_action", "a.master_action")
      vAnsiJoins.Add("users u", "a.created_by", "u.logname")
      vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
      vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
      vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "events e", vWhereFields, "sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
      'For Each vRow As CDBDataRow In pDataTable.Rows
      '  'Order by Status (Overdue, Defined, Scheduled)
      '  If vRow.Item("SortColumn") = Action.GetActionStatusCode(astScheduled) Then vRow.Item("SortColumn") = ""
      'Next
      'pDataTable.ReOrderRowsByColumn("SortColumn", True)
      'GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
      GetActionersAndSubjects(pDataTable)
    End Sub

    Private Sub GetAppealActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cap.master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,,,,,,,,,,,duration_days,duration_hours,duration_minutes,a.document_class,action_text"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cap.campaign", mvParameters("Campaign").Value)
      vWhereFields.Add("cap.appeal", mvParameters("Appeal").Value)
      If mvParameters.HasValue("ActionNumber") Then vWhereFields.Add("a.action_number", mvParameters("ActionNumber").IntegerValue)
      vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department", mvEnv.User.Department)
      vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("actions a", "cap.master_action", "a.master_action")
      vAnsiJoins.Add("users u", "a.created_by", "u.logname")
      vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
      vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
      vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "appeals cap", vWhereFields, "sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
      'For Each vRow As CDBDataRow In pDataTable.Rows
      '  'Order by Status (Overdue, Defined, Scheduled)
      '  If vRow.Item("SortColumn") = Action.GetActionStatusCode(astScheduled) Then vRow.Item("SortColumn") = ""
      'Next
      'pDataTable.ReOrderRowsByColumn("SortColumn", True)
      'GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
      GetActionersAndSubjects(pDataTable)
    End Sub

    Private Sub GetContactSurveys(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "cs.contact_survey_number,cs.survey_number,s.survey_name,sv.survey_version,cs.sent_on,cs.completed_on,sv.valid_from,sv.valid_to,sv.closing_date,s.long_description,cs.notes,sv.survey_version_number,cs.created_by,cs.created_on,cs.amended_by,cs.amended_on"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      If mvParameters.Exists("ContactSurveyNumber") Then vWhereFields.Add("cs.contact_survey_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactSurveyNumber").IntegerValue)
      vWhereFields.Add("cs.contact_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactNumber").IntegerValue)
      vAnsiJoins.Add("contact_surveys cs", "cs.survey_number", "s.survey_number")
      vAnsiJoins.Add("survey_versions sv", "cs.survey_version_number", "sv.survey_version_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "surveys s", vWhereFields, "s.survey_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetContactSurveyResponses(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "csr.survey_response_number,cs.contact_survey_number,sq.survey_question_number,sq.question_text,sq.mandatory,sq.answer_type,sq.help_text,sa.survey_answer_number,sa.answer_text,sa.answer_data_type,sa.minimum_value,sa.maximum_value,sa.list_values,sa.next_question_number,csr.response_answer_text,csr.created_by,csr.created_on,csr.amended_by,csr.amended_on,csr.response_answer_text AS display_answer_text,sq.new_page"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("cs.contact_survey_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactSurveyNumber").IntegerValue)
      vAnsiJoins.Add("survey_questions sq", "cs.survey_number", "sq.survey_number")
      vAnsiJoins.Add("survey_answers sa", "sa.survey_question_number", "sq.survey_question_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("contact_survey_responses csr", "cs.contact_survey_number", "csr.contact_survey_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins(2).AddJoinFields("sq.survey_question_number", "csr.survey_question_number")
      vAnsiJoins(2).AddJoinFields("sa.survey_answer_number", "csr.survey_answer_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "contact_surveys cs", vWhereFields, "sq.question_sequence,sa.answer_sequence", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      For vCtr As Integer = 0 To pDataTable.Rows.Count - 1
        If pDataTable.Rows(vCtr).Item("AnswerType") = "S" AndAlso pDataTable.Rows(vCtr).Item("AnswerDataType") = "Y" AndAlso pDataTable.Rows(vCtr).Item("ResponseAnswerText") <> "" Then
          pDataTable.Rows(vCtr).Item("DisplayAnswerText") = pDataTable.Rows(vCtr).Item("AnswerText")
        Else
          pDataTable.Rows(vCtr).Item("DisplayAnswerText") = pDataTable.Rows(vCtr).Item("ResponseAnswerText")
        End If
      Next
    End Sub

    Private Sub GetContactSalesLedgerReceipts(ByVal pDataTable As CDBDataTable)
      Dim vInvoice As New Invoice()
      vInvoice.Init(mvEnv, mvParameters("BatchNumber").IntegerValue, mvParameters("TransactionNumber").IntegerValue)

      Dim vAttrs As String = "ci.invoice_date, ci.record_type, ci.invoice_number, fh.reference, fh.status, ci.batch_number, ci.transaction_number, iph.line_number, "
      vAttrs &= "tt.transaction_sign, iph.amount AS debit_amount, iph.amount AS credit_amount, ci.outstanding, ci.invoice_number AS stored_invoice_number, "
      vAttrs &= "fh.contact_number, iph.status AS iph_status"
      Dim vTableName As String = "invoices"
      Dim vTableAlias As String = "i"
      Dim vPayRecordType As String = "C"
      Select Case vInvoice.InvoiceType
        Case Invoice.InvoiceRecordType.Invoice
        Case Invoice.InvoiceRecordType.CreditNote
          vTableAlias = "cn"
          vPayRecordType = "I"
        Case Else
          'Invoice.InvoiceRecordType.SalesLedgerCash
          vTableAlias = "ca"
          vPayRecordType = "I"
      End Select

      'Nested SQL for payments invoice
      Dim vSubAttrs As String = "{0}.batch_number, {0}.transaction_number, pi.invoice_number, pi.invoice_pay_status, pi.invoice_date, pi.record_type, pi.sales_ledger_account,"
      vSubAttrs &= " SUM(bta.amount) As invoice_amount, pi.amount_paid, (SUM(bta.amount) - pi.amount_paid) AS outstanding"
      Dim vSubAnsiJoins As New AnsiJoins()
      vSubAnsiJoins.Add("invoice_details pid", "pi.batch_number", "pid.batch_number", "pi.transaction_number", "pid.transaction_number", "pi.invoice_number", "pid.invoice_number")
      vSubAnsiJoins.Add("batch_transaction_analysis bta", "pid.batch_number", "bta.batch_number", "pid.transaction_number", "bta.transaction_number", "pid.line_number", "bta.line_number")
      If vInvoice.InvoiceType = Invoice.InvoiceRecordType.Invoice Then
        vSubAnsiJoins.AddLeftOuterJoin("reversals r", "bta.batch_number", "r.batch_number", "bta.transaction_number", "r.transaction_number", "bta.line_number", "r.line_number")
        vSubAnsiJoins.AddLeftOuterJoin("financial_history fh", "r.was_batch_number", "fh.batch_number", "r.was_transaction_number", "fh.transaction_number")
        vSubAttrs &= ", fh.status AS fh_status"
      End If
      Dim vSubWhereFields As New CDBFields()
      If vInvoice.InvoiceType <> Invoice.InvoiceRecordType.Invoice Then vSubWhereFields.Add("pi.record_type", vPayRecordType)
      Dim vSubGroupBy As String = "{0}.batch_number, {0}.transaction_number, pi.amount_paid, pi.invoice_pay_status, pi.invoice_date, pi.record_type, pi.invoice_number, pi.sales_ledger_account"
      If vInvoice.InvoiceType = Invoice.InvoiceRecordType.Invoice Then vSubGroupBy &= ", fh.status"
      Dim vSubSQLStatement As New SQLStatement(mvEnv.Connection, String.Format(vSubAttrs, "pi"), "invoices pi", vSubWhereFields, "", vSubAnsiJoins)
      vSubSQLStatement.GroupBy = String.Format(vSubGroupBy, "pi")

      'Main SQL tables
      Dim vAnsiJoins As New AnsiJoins()
      If vInvoice.InvoiceType = Invoice.InvoiceRecordType.Invoice Then
        vAnsiJoins.Add("invoice_payment_history iph", "i.invoice_number", "iph.invoice_number")
        vAnsiJoins.Add("(" & vSubSQLStatement.SQL & ") ci", "iph.batch_number", "ci.batch_number", "iph.transaction_number", "ci.transaction_number")    'Payment invoice
      Else
        vAnsiJoins.Add("invoice_payment_history iph", vTableAlias & ".batch_number", "iph.batch_number", vTableAlias & ".transaction_number", "iph.transaction_number")
        vAnsiJoins.Add("(" & vSubSQLStatement.SQL & ") ci", "iph.invoice_number", "ci.invoice_number")    'Paid invoice
      End If
      vAnsiJoins.Add("financial_history fh", "ci.batch_number", "fh.batch_number", "ci.transaction_number", "fh.transaction_number")
      vAnsiJoins.Add("transaction_types tt", "fh.transaction_type", "tt.transaction_type")

      'Main SQL Where clause
      Dim vWhereFields As New CDBFields(New CDBField(vTableAlias & ".batch_number", mvParameters("BatchNumber").IntegerValue))
      vWhereFields.Add(vTableAlias & ".transaction_number", mvParameters("TransactionNumber").IntegerValue)
      Select Case vInvoice.InvoiceType
        Case Invoice.InvoiceRecordType.Invoice
          If mvParameters.ParameterExists("InvoiceNumber").IntegerValue > 0 Then vWhereFields.Add(vTableAlias & ".invoice_number", mvParameters("InvoiceNumber").IntegerValue)
          vWhereFields.Add("ci.invoice_amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoCloseBracket)
          vWhereFields.Add("ci.invoice_amount#2", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLessThan)
          vWhereFields.Add("ci.fh_status", CDBField.FieldTypes.cftCharacter, "A", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
          vWhereFields.Add("fh.status", CDBField.FieldTypes.cftCharacter, "A", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoNOT)
          vWhereFields.Add("ci.outstanding", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoCloseBracket)
        Case Invoice.InvoiceRecordType.SalesLedgerCash
          If mvContact.ContactNumber > 0 Then vWhereFields.Add("ca.contact_number", mvContact.ContactNumber)    'Restrict to only show payments for this Contact
      End Select

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTableName & " " & vTableAlias, vWhereFields, "ci.invoice_date DESC, ci.invoice_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      'Handle reversals of Sundry Credit Notes & Sales Ledger Cash from Remove Allocations
      'Invoice Payment History has been created with (probably) incorrect Batch / Transaction / Line and no invoices record for those numbers
      If vInvoice.InvoiceType = Invoice.InvoiceRecordType.Invoice Then
        'Sub SQL
        vSubAttrs = vSubAttrs.Substring(0, vSubAttrs.Length - 24)    'Remove ", fh.status AS fh_status"
        vSubGroupBy = vSubGroupBy.Substring(0, vSubGroupBy.Length - 11)  'Remove ", fh.status"
        vSubAnsiJoins.RemoveAt(vSubAnsiJoins.Count - 1)   'Remove last two entries
        vSubAnsiJoins.RemoveAt(vSubAnsiJoins.Count - 1)
        vSubAnsiJoins.Insert(0, New AnsiJoin("invoice_payment_history iphr", "r.was_batch_number", "iphr.allocation_batch_number", "r.was_transaction_number", "iphr.allocation_transaction_number", "r.was_line_number", "iphr.allocation_line_number", AnsiJoin.AnsiJoinTypes.InnerJoin))
        vSubAnsiJoins.Insert(1, New AnsiJoin("invoices pi", "iphr.batch_number", "pi.batch_number", "iphr.transaction_number", "pi.transaction_number", AnsiJoin.AnsiJoinTypes.InnerJoin))
        vSubAnsiJoins.AddLeftOuterJoin("invoices oi", "r.batch_number", "oi.batch_number", "r.transaction_number", "oi.transaction_number")
        vSubWhereFields.Add("pi.record_type", CDBField.FieldTypes.cftCharacter, "'N','C'", CDBField.FieldWhereOperators.fwoIn)
        vSubWhereFields.Add("oi.batch_number", CDBField.FieldTypes.cftInteger, "")
        vSubSQLStatement = New SQLStatement(mvEnv.Connection, String.Format(vSubAttrs, "r"), "reversals r", vSubWhereFields, "", vSubAnsiJoins)
        vSubSQLStatement.GroupBy = String.Format(vSubGroupBy, "r")
        'Main SQL
        vAttrs = vAttrs.Replace("ci.invoice_date,", "fh.transaction_date,")
        vAnsiJoins.RemoveAt(1)
        vAnsiJoins.Insert(1, New AnsiJoin("(" & vSubSQLStatement.SQL & ") ci", "iph.batch_number", "ci.batch_number", "iph.transaction_number", "ci.transaction_number", AnsiJoin.AnsiJoinTypes.InnerJoin))   'Payment invoice
        vWhereFields.Remove(vWhereFields.Count)   'Remove last 5 items
        vWhereFields.Remove(vWhereFields.Count)
        vWhereFields.Remove(vWhereFields.Count)
        vWhereFields.Remove(vWhereFields.Count)
        vWhereFields.Remove(vWhereFields.Count)
        vWhereFields.Add("ci.invoice_amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      Else
        vAttrs = vAttrs.Replace("iph.amount", "iphr.amount").Replace(", iph.status", ", iphr.status")
        vAnsiJoins.Insert(1, New AnsiJoin("reversals r", "iph.allocation_batch_number", "r.was_batch_number", "iph.allocation_transaction_number", "r.was_transaction_number", "iph.allocation_line_number", "r.was_line_number", AnsiJoin.AnsiJoinTypes.InnerJoin))
        vAnsiJoins.Insert(2, New AnsiJoin("invoice_payment_history iphr", "r.batch_number", "iphr.batch_number", "r.transaction_number", "iphr.transaction_number", "r.line_number", "iphr.line_number", AnsiJoin.AnsiJoinTypes.InnerJoin))
      End If
      vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrs, vTableName & " " & vTableAlias, vWhereFields, "ci.invoice_date DESC, ci.invoice_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      'Now re-sort the data
      Dim vTableSort(1) As CDBDataTable.SortSpecification
      vTableSort(0).Column = "Date"
      vTableSort(0).Descending = True
      vTableSort(1).Column = "InvoiceNumber"
      vTableSort(1).Descending = False
      pDataTable.ReOrderRowsByMultipleColumns(vTableSort)

      'Remove any duplicate that can occur when money moved from Contact to Contact
      If vInvoice.InvoiceType = Invoice.InvoiceRecordType.SalesLedgerCash Then pDataTable.RemoveFullyDuplicatedRows()

      'Manipulate the data
      For Each vRow As CDBDataRow In pDataTable.Rows
        Select Case mvEnv.GetInvoicePayStatusType(vRow.Item("InvoicePayStatus"))    'Invoice Pay Status
          Case CDBEnvironment.InvoicePayStatusTypes.ipsPaymentDue
            vRow.Item("InvoicePayStatus") = ""
          Case CDBEnvironment.InvoicePayStatusTypes.ipsPartPaid
            'don't need to do anything
          Case CDBEnvironment.InvoicePayStatusTypes.ipsFullyPaid
            vRow.Item("InvoicePayStatus") = "A"
          Case CDBEnvironment.InvoicePayStatusTypes.ipsDDCollectionPending
            'not yet supported
        End Select
        Select Case Invoice.GetRecordType(vRow.Item("TransactionType"))    'Invoice Record Type
          Case Invoice.InvoiceRecordType.SalesLedgerCash
            vRow.Item("InvoiceNumber") = ""
            vRow.Item("Debit") = ""
            If vRow.Item("TransactionSign") = "D" OrElse DoubleValue(vRow.Item("Credit")) < 0 Then
              vRow.Item("TransactionType") = DataSelectionText.String18679    'Adjustment
              If DoubleValue(vRow.Item("Credit")) < 0 Then
                vRow.Item("Debit") = System.Math.Abs(DoubleValue(vRow.Item("Credit"))).ToString
                vRow.Item("Credit") = ""
              End If
            Else
              vRow.Item("TransactionType") = DataSelectionText.String18680    'Payment
            End If
          Case Invoice.InvoiceRecordType.CreditNote
            vRow.Item("TransactionType") = DataSelectionText.String18681    'Credit Note
            If vRow.Item("TransactionSign").Equals("C", StringComparison.InvariantCultureIgnoreCase) OrElse DoubleValue(vRow.Item("Debit")) < 0 Then
              vRow.Item("Debit") = Math.Abs(DoubleValue(vRow.Item("Debit"))).ToString("0.00")
              vRow.Item("Credit") = ""
            Else
              vRow.Item("Debit") = ""
            End If
          Case Invoice.InvoiceRecordType.Invoice
            vRow.Item("TransactionType") = DataSelectionText.String18682    'Invoice
            If (vRow.Item("TransactionSign").Equals("D", StringComparison.InvariantCultureIgnoreCase) OrElse DoubleValue(vRow.Item("Credit")) < 0) Then
              Dim vAmount As Double = Math.Abs(DoubleValue(vRow.Item("Credit")))
              If vInvoice.InvoiceType = Invoice.InvoiceRecordType.Invoice Then
                vRow.Item("Debit") = vAmount.ToString("0.00")
                vRow.Item("Credit") = ""
              Else
                vRow.Item("Credit") = vAmount.ToString("0.00")
                vRow.Item("Debit") = ""
              End If
            Else
              vRow.Item("Credit") = ""
            End If
        End Select
      Next
    End Sub

    Private Sub GetWebProducts(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "p.product,p.product_desc,p.extra_key,p.sales_group,sg.sales_group_desc,p.secondary_group,scg.secondary_group_desc,p.product_category,pc.product_category_desc,p.sales_description,r.rate,r.rate_desc,r.current_price,r.future_price,r.vat_exclusive,r.price_change_date,vr.percentage,0 AS GrossPrice,0 AS NetPrice,'' as ProductImage,Null As VatAmount"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields1 As New CDBFields()
      Dim vAnsiJoins1 As New AnsiJoins()
      Dim vVatCategory As String
      Dim vSQLCategory As SQLStatement
      Dim vWhereCategory As New CDBFields()
      If mvParameters.Exists("ContactNumber") Then
        vWhereCategory.Add("contact_number", CDBField.FieldTypes.cftCharacter, mvParameters("ContactNumber").Value)
        vSQLCategory = New SQLStatement(mvEnv.Connection, "contact_vat_category", "contacts", vWhereFields)
        vVatCategory = vSQLCategory.GetValue()
      Else
        vSQLCategory = New SQLStatement(mvEnv.Connection, "default_contact_vat_cat", "financial_controls")
        vVatCategory = vSQLCategory.GetValue()
      End If
      If mvParameters.Exists("ContactNumber") Then
        GetMembershipLookupGroupSQL(vAnsiJoins, vWhereFields)
      End If
      vWhereFields.Add("r.web_publish", "Y")
      vWhereFields.Add("r.history_only", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "product,MIN(current_price) as current_price", "rates r", vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = "product"

      vAnsiJoins1.Add("(" & vSQLStatement.SQL & ") ir", "ir.product", "r.product", "ir.current_price", "r.current_price")
      vAnsiJoins1.Add("products p", "r.product", "p.product")
      vAnsiJoins1.AddLeftOuterJoin("sales_groups sg", "p.sales_group", "sg.sales_group")
      vAnsiJoins1.AddLeftOuterJoin("secondary_groups scg", "p.secondary_group", "scg.secondary_group")
      vAnsiJoins1.AddLeftOuterJoin("product_categories pc", "p.product_category", "pc.product_category")
      vAnsiJoins1.AddLeftOuterJoin("vat_rate_identification vri", "p.product_vat_category", "vri.product_vat_category")
      vAnsiJoins1.AddLeftOuterJoin("vat_rates vr", "vri.vat_rate", "vr.vat_rate")
      vWhereFields1.Add("p.subscription", "N")
      vWhereFields1.Add("p.donation", "N")
      vWhereFields1.Add("p.course", "N")
      vWhereFields1.Add("p.accommodation", "N")
      vWhereFields1.Add("p.postage_packing", "N")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExams) Then vWhereFields1.Add("exam", "N")
      vWhereFields1.Add("p.history_only", "N")
      vWhereFields1.Add("r.history_only", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
      vWhereFields1.Add("p.uses_product_numbers", "N")
      vWhereFields1.Add("p.web_publish", "Y")
      vWhereFields1.Add("r.web_publish", "Y")
      vWhereFields1.Add("vri.contact_vat_category", vVatCategory, CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields1.Add("vri.contact_vat_category#2", "", CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
      If mvParameters.Exists("SalesGroup") Then vWhereFields1.Add("p.sales_group", mvParameters("SalesGroup").Value)
      If mvParameters.Exists("SecondaryGroup") Then vWhereFields1.Add("p.secondary_group", mvParameters("SecondaryGroup").Value)
      If mvParameters.Exists("ProductCategory") Then vWhereFields1.Add("p.product_category", mvParameters("ProductCategory").Value)
      'Adding where fields if Search criteria is passed as a parameter
      If mvParameters.Exists("SearchProduct") Then
        vWhereFields1.Add("p.product_desc", mvParameters("SearchProduct").Value & "*", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vWhereFields1.Add("p.extra_key", mvParameters("SearchProduct").Value & "*", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vWhereFields1.Add("p.sales_description", mvParameters("SearchProduct").Value & "*", CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
      End If
      If mvParameters.Exists("Product") Then
        vWhereFields1.Add("p.product", mvParameters("Product").Value)
      End If
      Dim vNewSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "rates r", vWhereFields1, "r.product", vAnsiJoins1)
      pDataTable.FillFromSQL(mvEnv, vNewSQLStatement)
      Dim vIndex As Integer
      Dim vPrevProduct As String = ""
      Dim vGrossPrice As Double
      Dim vNetPrice As Double
      Dim vProducts As New SortedList()
      Dim vVRate As New VatRate(mvEnv)
      Dim vVatAmount As Double
      While vIndex <= pDataTable.Rows.Count - 1
        If vProducts.ContainsKey(pDataTable.Rows(vIndex).Item("Product").ToString) Then
          pDataTable.RemoveRow(pDataTable.Rows(vIndex))
          vIndex = vIndex - 1
        Else
          vProducts.Add(pDataTable.Rows(vIndex).Item("Product").ToString, pDataTable.Rows(vIndex).Item("Product").ToString)
          If CDate(pDataTable.Rows(vIndex).Item("PriceChangeDate").ToString) <= Today Then
            vNetPrice = CDbl(pDataTable.Rows(vIndex).Item("FuturePrice").ToString)
          Else
            vNetPrice = CDbl(pDataTable.Rows(vIndex).Item("CurrentPrice").ToString)
          End If

          vVatAmount = vVRate.CalculateVATAmount(vNetPrice, BooleanValue(pDataTable.Rows(vIndex).Item("VatExclusive").ToString), DoubleValue(pDataTable.Rows(vIndex).Item("Percentage").ToString))
          If pDataTable.Rows(vIndex).Item("VatExclusive").ToString = "Y" Then
            vGrossPrice = vNetPrice + vVatAmount
          Else
            vGrossPrice = vNetPrice
            vNetPrice = vGrossPrice - vVatAmount
          End If

          pDataTable.Rows(vIndex).Item("NetPrice") = vNetPrice.ToString("0.00")
          pDataTable.Rows(vIndex).Item("GrossPrice") = vGrossPrice.ToString("0.00")
          vPrevProduct = pDataTable.Rows(vIndex).Item("Product").ToString
        End If
        pDataTable.Rows(vIndex).Item("VatAmount") = vVatAmount.ToString
        vIndex += 1
      End While
    End Sub

    Private Sub GetWebEvents(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "e.event_number,e.event_desc,e.event_reference,Null As EventImage,sb.subject,sb.subject_desc,sl.skill_level,sl.skill_level_desc,s.start_date,s.start_time,s.end_date,s.end_time,o.name,e.venue,v.venue_desc,e.venue_reference,v.location,s.number_of_attendees,s.maximum_attendees,e.event_status,es.event_status_desc,e.bookings_close,e.free_of_charge,e.long_description,e.long_description AS short_long_description,e.event_class,ec.event_class_desc,s.notes,e.branch,e.balance_booking_fee,e.balance_booking_due,e.minimum_sponsorship_amount,e.sponsorship_due,e.pledged_amount_due,e.multi_session," & mvEnv.Connection.DBSpecialCol("e", "external") & " ,eo.organiser,eo.price_to_attendees,e.eligibility_check_required,e.eligibility_check_text,eo.reference,e.department,e.name_attendees,e.number_of_bookings,Null As BookingStatusDesc,Null as AvailablePlaces"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("e.web_publish", "Y")
      vWhereFields.Add("e.booking", "Y")
      vWhereFields.Add("e.template", "N")
      If mvEnv.GetConfigOption("portal_display_closed_events") = False Then
        vWhereFields.Add("e.bookings_close", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      Else
        vWhereFields.Add("e.bookings_close", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoNotEqual)
      End If
      vWhereFields.Add("s.session_type", "0")
      If mvParameters.Exists("EventNumber") Then vWhereFields.Add("e.event_number", mvParameters("EventNumber").IntegerValue)
      If mvParameters.Exists("SearchEvent") Then
        vWhereFields.Add("e.event_reference", mvParameters("SearchEvent").Value, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vWhereFields.Add("e.event_desc", mvParameters("SearchEvent").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("v.venue_desc", mvParameters("SearchEvent").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("e.long_description", mvParameters("SearchEvent").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      If mvParameters.Exists("StartDate") AndAlso mvParameters.Exists("EndDate") Then
        vWhereFields.Add("e.start_date", CDBField.FieldTypes.cftDate, mvParameters("StartDate").Value, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        vWhereFields.Add("e.start_date#1", CDBField.FieldTypes.cftDate, mvParameters("EndDate").Value, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Else
        vWhereFields.Add("e.start_date", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      vAnsiJoins.Add("sessions s", "e.event_number", "s.event_number")
      vAnsiJoins.Add("subjects sb", "s.subject", "sb.subject")
      vAnsiJoins.Add("skill_levels sl", "s.skill_level", "sl.skill_level")
      vAnsiJoins.Add("venues v", "v.venue", "e.venue")
      If mvParameters.Exists("Topic") Then
        vAnsiJoins.Add("event_topics et", "e.event_number", "et.event_number")
        vWhereFields.Add("et.topic", mvParameters("Topic").Value)
      End If
      vAnsiJoins.Add("event_statuses es", "e.event_status", "es.event_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("event_classes ec", "e.event_class", "ec.event_class", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("event_organisers eo", "e.event_number", "eo.event_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("branches b", "e.branch", "b.branch", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("organisations o", "b.organisation_number", "o.organisation_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "events e", vWhereFields, "s.start_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, False)
      ' For setting the Booking Status Description.
      Dim vTable As New CDBDataTable
      vWhereFields.Clear()
      vSQLStatement = New SQLStatement(mvEnv.Connection, "event_number,session_number,maximum_attendees,number_of_attendees", "sessions", vWhereFields, "event_number")
      vTable.FillFromSQL(mvEnv, vSQLStatement, False)
      Dim vList As New SortedList(Of String, String)
      For Each vDataRow As CDBDataRow In vTable.Rows
        If IntegerValue(vDataRow.Item("number_of_attendees")) >= IntegerValue(vDataRow.Item("maximum_attendees")) Then
          If IntegerValue(vDataRow.Item("session_number")) = IntegerValue(vDataRow.Item("event_number").ToString) * 10000 Then
            If vList.ContainsKey(vDataRow.Item("event_number").ToString) Then
              vList.Remove(vDataRow.Item("event_number").ToString)
            End If
            vList.Add(vDataRow.Item("event_number").ToString, ProjectText.String17129) 'FULLY BOOKED
          Else
            If vList.ContainsKey(vDataRow.Item("event_number").ToString) Then
              vList.Remove(vDataRow.Item("event_number").ToString)
            End If
            vList.Add(vDataRow.Item("event_number").ToString, ProjectText.String17130) 'SOME SESSIONS FULLY BOOKED
          End If
        End If
      Next
      Dim vConfigValue As String = mvEnv.GetConfig("web_event_image_name")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vList.ContainsKey(vRow.Item("EventNumber").ToString) Then
          vRow.Item("BookingStatusDesc") = vList(vRow.Item("EventNumber").ToString).ToString
        Else
          vRow.Item("BookingStatusDesc") = "" 'NOT FULLY BOOKED
        End If
        If vConfigValue.Length > 0 Then
          vRow.Item("EventImage") = String.Format(vConfigValue, vRow.Item("EventNumber"))
        Else
          vRow.Item("EventImage") = String.Format("Event{0}.png", vRow.Item("EventNumber"))
        End If
        'Calculate Available Places depending on Maximum Attendees and Number of Bookings
        vRow.Item("AvailablePlaces") = (IntegerValue(vRow.Item("MaximumAttendees")) - IntegerValue(vRow.Item("NumberOfAttendees"))).ToString
        vRow.Item("LongDescription") = vRow.Item("LongDescription").Replace(vbLf, "<br>")
        If vRow.Item("ShortLongDescription").Length > 120 Then vRow.Item("ShortLongDescription") = vRow.Item("ShortLongDescription").Substring(0, 120) & "..."
      Next
    End Sub

    Private Sub GetWebBookingOptions(ByVal pDataTable As CDBDataTable)
      ''''''''''
      Dim vAttrs As New StringBuilder("ebo.option_number,ebo.option_desc,ebo.pick_sessions,ebo.minimum_bookings,ebo.maximum_bookings,ebo.product,p.product_desc,p.minimum_quantity,p.maximum_quantity,r.rate,r.rate_desc,0 as GrossPrice,0 as NetPrice,e.start_date,r.price_change_date,r.future_price,r.current_price,vr.percentage")

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEventMinimumBookings) = False Then vAttrs = vAttrs.Replace("ebo.minimum_bookings", "1 AS minimum_bookings")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLongDescription) Then
        vAttrs.Append(",ebo.long_description,")
      Else
        vAttrs.Append(",NULL AS long_description,")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDaysPrior) Then
        vAttrs.Append(mvEnv.Connection.DBIsNull("r.days_prior_to", "0"))
        vAttrs.Append("AS days_prior_to,")
        vAttrs.Append(mvEnv.Connection.DBIsNull("r.days_prior_From", "99999"))
        vAttrs.Append("AS days_prior_From")
      Else
        vAttrs.Append("NULL AS days_prior_to,NULL AS days_prior_From")
      End If
      vAttrs.Append(",ebo.number_of_sessions,r.vat_exclusive")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifier) Then
        vAttrs.Append(",r.use_modifiers ")
      Else
        vAttrs.Append(" ,NULL AS use_modifiers ")
      End If

      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("event_booking_options ebo", "r.product", "ebo.product")
      vAnsiJoins.Add("products p", "r.product", "p.product")
      vAnsiJoins.Add("events e", "ebo.event_number", "e.event_number")
      vAnsiJoins.AddLeftOuterJoin("product_categories pc", "p.product_category", "pc.product_category")
      vAnsiJoins.AddLeftOuterJoin("vat_rate_identification vri", "p.product_vat_category", "vri.product_vat_category")
      vAnsiJoins.Add("vat_rates vr", "vri.vat_rate", "vr.vat_rate", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      If mvParameters.Exists("ContactNumber") Then
        GetMembershipLookupGroupSQL(vAnsiJoins, vWhereFields)
      End If

      If mvParameters.Exists("EventNumber") Then vWhereFields.Add("e.event_number", mvParameters("EventNumber").IntegerValue)
      vWhereFields.Add("e.web_publish", "Y")
      vWhereFields.Add("e.bookings_close", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      Dim vVatCategory As String
      If mvParameters.Exists("ContactNumber") Then
        Dim vWhereCategory As New CDBFields()
        vWhereCategory.Add("contact_number", CDBField.FieldTypes.cftCharacter, mvParameters("ContactNumber").Value)
        vVatCategory = New SQLStatement(mvEnv.Connection, "contact_vat_category", "contacts", vWhereCategory).GetValue()
      Else
        vVatCategory = New SQLStatement(mvEnv.Connection, "default_contact_vat_cat", "financial_controls").GetValue()
      End If
      vWhereFields.Add("vri.contact_vat_category", vVatCategory, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("vri.contact_vat_category#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      If mvParameters.Exists("OptionNumber") Then vWhereFields.Add("ebo.option_number", mvParameters("OptionNumber").Value)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWebPublish) Then
        vWhereFields.Add("p.web_publish", "Y", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("r.web_publish", "Y", CDBField.FieldWhereOperators.fwoEqual)
      End If
      vWhereFields.Add("r.history_only", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
      'Adding where fields if Search criteria is passed as a parameter

      'Add this clause to exclude booking options for which there are no sessions with any places (pick sessions or not)
      Dim vSubWhereFields As New CDBFields
      If mvParameters.Exists("EventNumber") Then vSubWhereFields.Add("ebo.event_number", mvParameters("EventNumber").IntegerValue)
      vSubWhereFields.Add("maximum_attendees", CDBField.FieldTypes.cftInteger, "number_of_attendees", CDBField.FieldWhereOperators.fwoGreaterThan)
      'Don't include session type zero if this is a multi session event
      vSubWhereFields.Add("session_type", "0", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoNotEqual)
      vSubWhereFields.Add("e.multi_session", "N", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)

      Dim vSubAnsiJoins As New AnsiJoins
      vSubAnsiJoins.Add("events e", "ebo.event_number", "e.event_number")
      vSubAnsiJoins.AddLeftOuterJoin("option_sessions os", "ebo.option_number", "os.option_number")
      vSubAnsiJoins.AddLeftOuterJoin("sessions s", "os.session_number", "s.session_number")
      Dim vSubSelect As New SQLStatement(mvEnv.Connection, "ebo.option_number", "event_booking_options ebo", vSubWhereFields, "", vSubAnsiJoins)
      vSubSelect.GroupBy = "ebo.option_number"
      vWhereFields.Add("ebo.option_number#2", CDBField.FieldTypes.cftInteger, String.Format("( {0} )", vSubSelect.SQL), CDBField.FieldWhereOperators.fwoIn)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbFreeOfChangeBookingOption) Then
        vWhereFields.Add("ebo.free_of_charge", "N", CDBField.FieldWhereOperators.fwoEqual)
      End If

      Dim vNewSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs.ToString, "rates r", vWhereFields, "ebo.option_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vNewSQLStatement)

      'Check for Days Prior From and To validation before trying to find the lowest price rate
      Dim vIndex As Integer
      While vIndex <= pDataTable.Rows.Count - 1
        'Applying days prior from and days prior to validation
        Dim vDiff As Integer
        vDiff = IntegerValue(DateDiff(DateInterval.Day, CDate(TodaysDate()), CDate(pDataTable.Rows(vIndex).Item("StartDate").ToString)).ToString)
        If vDiff >= 0 Then
          If vDiff >= IntegerValue(pDataTable.Rows(vIndex).Item("DaysPriorTo").ToString) And vDiff <= IntegerValue(pDataTable.Rows(vIndex).Item("DaysPriorFrom").ToString) Then
            'Don't remove record
          Else
            pDataTable.RemoveRow(pDataTable.Rows(vIndex))
            vIndex = vIndex - 1
          End If
        End If
        vIndex += 1
      End While

      vIndex = 0
      Dim vGrossPrice As Double
      Dim vNetPrice As Double
      Dim vVatAmount As Double
      Dim vVRate As New VatRate(mvEnv)
      Dim vLines As New SortedList(Of String, ProductRateOption)
      Dim vRate As New ProductRate(mvEnv)
      Dim vContact As Contact = Nothing
      Dim vCalculateGrossAmount As Boolean
      'removing duplicate product row and finding the lowest price
      While vIndex <= pDataTable.Rows.Count - 1
        vCalculateGrossAmount = True
        If BooleanValue(pDataTable.Rows(vIndex).Item("UseModifiers").ToString) Then
          'Rate is setup to use Rate Modifiers
          If mvParameters.Exists("ContactNumber") AndAlso vContact Is Nothing Then
            vContact = New Contact(mvEnv)
            vContact.Init(mvParameters("ContactNumber").IntegerValue)
          End If
          vRate.Init(pDataTable.Rows(vIndex).Item("Product").ToString, pDataTable.Rows(vIndex).Item("Rate").ToString)
          vNetPrice = vRate.Price(Today, vContact)
          If vRate.VatExclusive AndAlso (vRate.PriceIsPercentage.Length = 0 OrElse vRate.PriceIsPercentage = "N") Then
            'When this conditin is true, vRate.Price adds the VatAmount to the price. This price is GrossAmount but we also need to calculate the NetAmount
            vVatAmount = vVRate.CalculateVATAmount(vNetPrice, False, CDbl(pDataTable.Rows(vIndex).Item("Percentage").ToString))
            vGrossPrice = vNetPrice
            vNetPrice = vGrossPrice - vVatAmount
            vCalculateGrossAmount = False
          End If
        Else
          If CDate(pDataTable.Rows(vIndex).Item("PriceChangeDate").ToString) <= Today Then
            vNetPrice = CDbl(pDataTable.Rows(vIndex).Item("FuturePrice").ToString)
          Else
            vNetPrice = CDbl(pDataTable.Rows(vIndex).Item("CurrentPrice").ToString)
          End If
        End If
        If vCalculateGrossAmount Then
          vVatAmount = vVRate.CalculateVATAmount(vNetPrice, BooleanValue(pDataTable.Rows(vIndex).Item("VatExclusive").ToString), CDbl(pDataTable.Rows(vIndex).Item("Percentage").ToString))
          If pDataTable.Rows(vIndex).Item("VatExclusive").ToString = "Y" Then
            vGrossPrice = vNetPrice + vVatAmount
          Else
            vGrossPrice = vNetPrice
            vNetPrice = vGrossPrice - vVatAmount
          End If
        End If
        pDataTable.Rows(vIndex).Item("NetPrice") = vNetPrice.ToString("0.00")
        pDataTable.Rows(vIndex).Item("GrossPrice") = vGrossPrice.ToString("0.00")

        Dim vOptionNumber As String = pDataTable.Rows(vIndex).Item("OptionNumber").ToString
        Dim vProductRateOption As New ProductRateOption(vGrossPrice, pDataTable.Rows(vIndex))

        If vLines.ContainsKey(vOptionNumber) Then
          If vLines(vOptionNumber).Price > vGrossPrice Then
            'The price of the current row is lower than the price of the previous row. Remove the previous row.
            Dim vRequiredRow As CDBDataRow = pDataTable.Rows(vIndex)
            pDataTable.Rows.Remove(vLines(vOptionNumber).Row)
            vLines(vOptionNumber).Price = vGrossPrice 'Set the new price to be compared with the rest of the rows
            vLines(vOptionNumber).Row = vRequiredRow
          Else
            'The price of the current row is not lower than the price of the previous row. Remove the current row.
            pDataTable.RemoveRow(pDataTable.Rows(vIndex))
          End If
          vIndex = vIndex - 1
        Else
          'A new product/rate row.
          vLines.Add(vOptionNumber, vProductRateOption)
        End If
        vIndex += 1
      End While
    End Sub

    Private Class ProductRateOption
      Private mvPrice As Double
      Private mvRow As CDBDataRow

      Sub New(pPrice As Double, pRow As CDBDataRow)
        mvPrice = pPrice
        mvRow = pRow
      End Sub

      Public Property Price As Double
        Get
          Return mvPrice
        End Get
        Set(value As Double)
          mvPrice = value
        End Set
      End Property

      Public Property Row As CDBDataRow
        Get
          Return mvRow
        End Get
        Set(value As CDBDataRow)
          mvRow = value
        End Set
      End Property
    End Class

    Private Sub GetContactEventBookingDelegates(ByVal pDataTable As CDBDataTable)
      ''''''''''
      Dim vAttrs As String = "'Y' as select_rec,d.address_number,d.contact_number,c.label_name,d.position,d.organisation_name,d.event_delegate_number"
      Dim vWhereFields As New CDBFields()
      Dim vWhereFieldsNew As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vAnsiJoinsNew As New AnsiJoins()
      Dim vContactList As New SortedList()
      vAnsiJoins.Add("contacts c", "d.contact_number", "c.contact_number")
      If mvParameters.Exists("BookingNumber") Then
        vWhereFields.Add("d.booking_number ", CDBField.FieldTypes.cftCharacter, mvParameters("BookingNumber").Value)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "delegates d", vWhereFields, "", vAnsiJoins)
      If mvParameters.ContainsKey("ContactNumber") Then
        vAttrs = "'N' as select_rec,d.address_number,d.contact_number,c.label_name,d.position,d.organisation_name,d.event_delegate_number"
        vAnsiJoinsNew.Add("delegates d", "eb.booking_number", "d.booking_number")
        vAnsiJoinsNew.Add("contacts c", "d.contact_number", "c.contact_number")
        vWhereFieldsNew.Add("eb.contact_number", CDBField.FieldTypes.cftCharacter, mvParameters("ContactNumber").Value)
        Dim vNewSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_bookings eb", vWhereFieldsNew, "select_rec desc", vAnsiJoinsNew)
        vNewSQLStatement.Distinct = True
        vSQLStatement.AddUnion(vNewSQLStatement)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
        Dim vindex As Integer = 0
        While vindex <= pDataTable.Rows.Count - 1
          If Not vContactList.Contains(pDataTable.Rows(vindex).Item("ContactNumber")) Then
            vContactList.Add(pDataTable.Rows(vindex).Item("ContactNumber"), pDataTable.Rows(vindex).Item("ContactNumber"))
          Else
            pDataTable.RemoveRow(pDataTable.Rows(vindex))
            vindex = vindex - 1
          End If
          vindex = vindex + 1
        End While
      Else
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      End If
    End Sub

    Private Sub GetWebEventBookings(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "eb.booking_number,c1.label_name,c1.contact_number,eb.booking_date,c2.label_name AS delegate_label_name,c2.contact_number AS delegate_contact_number,eb.quantity,ebo.option_desc,e.event_number,e.event_desc,e.event_reference,s.start_date,s.start_time,s.end_date,s.end_time,v.venue,v.venue_desc,v.location,e.long_description,bta.gross_amount,bta.amount,bta.vat_amount,bookings_close"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("eb.contact_number", mvParameters("ContactNumber").IntegerValue, CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("d.contact_number", mvParameters("ContactNumber").IntegerValue, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("e.web_publish", "Y")
      vWhereFields.Add("e.cancellation_reason", "")
      vWhereFields.Add("e.start_date", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      vWhereFields.Add("b.provisional", "N")
      vWhereFields.Add("s.session_type", "0")

      vAnsiJoins.Add("contacts c1", "c1.contact_number", "eb.contact_number")
      vAnsiJoins.Add("delegates d", "eb.booking_number", "d.booking_number")
      vAnsiJoins.Add("contacts c2", "d.contact_number", "c2.contact_number")
      vAnsiJoins.Add("batch_transaction_analysis bta", "bta.batch_number", "eb.batch_number", "bta.transaction_number", "eb.transaction_number", "eb.line_number", "bta.line_number")
      vAnsiJoins.Add("batches b", "b.batch_number", "eb.batch_number")
      vAnsiJoins.Add("events e", "e.event_number", "eb.event_number")
      vAnsiJoins.Add("event_booking_options ebo", "e.event_number", "ebo.event_number", "ebo.option_number", "eb.option_number")
      'vAnsiJoins.Add("option_sessions os", "os.option_number", "ebo.option_number")
      vAnsiJoins.Add("sessions s", "s.event_number", "eb.event_number") ' , "os.session_number", "s.session_number")
      'vAnsiJoins.Add("session_bookings sb", "sb.booking_number", "eb.booking_number", "sb.session_number", "s.session_number")
      vAnsiJoins.Add("venues v", "e.venue", "v.venue")
      Dim vOrderBy As String = "e.start_date desc,eb.booking_number desc"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "event_bookings eb", vWhereFields, vOrderBy, vAnsiJoins)
      vAttrs = vAttrs.Replace("c2.label_name AS ", "")
      vAttrs = vAttrs.Replace("c2.contact_number AS ", "")
      vAttrs = vAttrs.Replace("eb.booking_number", "DISTINCT_EVENT_BOOKING_NUMBER")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub

    Private Sub GetServiceBookingTransactionData(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAttrs As String = Nothing
      Dim vAnsiJoins As New AnsiJoins()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceBookingAnalysis) Then
        vAttrs = "bta.batch_number,bta.transaction_number,bta.line_number,bt.transaction_date,product,rate,transaction_type_desc,payment_method,bta.distribution_code,quantity,bta.amount,vat_amount,vat_rate,source,bta.currency_amount,currency_vat_amount,bta.notes,tt.transaction_sign"
        vWhereFields.Add("sbt.service_booking_number", CDBField.FieldTypes.cftLong, mvParameters("ServiceBookingNumber").Value)
        vAnsiJoins.Add("batch_transaction_analysis bta", "bta.batch_number", "sbt.batch_number")
        vAnsiJoins(0).AddJoinFields("bta.transaction_number", "sbt.transaction_number")
        vAnsiJoins(0).AddJoinFields("bta.line_number", "sbt.line_number")
        vAnsiJoins.Add("batch_transactions bt", "bta.batch_number", "bt.batch_number")
        vAnsiJoins(1).AddJoinFields("bta.transaction_number", "bt.transaction_number")
        vAnsiJoins.Add("transaction_types tt", "bt.transaction_type", "tt.transaction_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "service_booking_transactions sbt", vWhereFields, "bt.transaction_date DESC, bta.batch_number DESC, bta.transaction_number DESC, bta.line_number DESC", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("TransactionSign") = "D" Then
            If Len(vRow.Item("Amount")) > 0 Then vRow.Item("Amount") = CStr(CDbl(vRow.Item("Amount")) * -1)
            vRow.Item("Quantity") = CStr(CInt(vRow.Item("Quantity")) * -1)
            If Len(vRow.Item("CurrencyAmount")) > 0 Then vRow.Item("CurrencyAmount") = CStr(CDbl(vRow.Item("CurrencyAmount")) * -1)
          End If
        Next
      End If
    End Sub

    Private Sub GetActivityFromActivityGroup(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "activity_group,activity,quantity_required,sequence_number,multiple_values"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("activity_group", mvParameters("ActivityGroup").Value, CDBField.FieldWhereOperators.fwoEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "activity_group_details", vWhereFields, "")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetEventCategories(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "delegate_activity_number,event_delegate_number,da.activity,da.activity_value,quantity,activity_date,da.source,valid_from,valid_to,da.amended_by,da.amended_on,da.notes,activity_desc,activity_value_desc,source_desc"
      Dim vOrderBy As String
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("activities a", "da.activity", "a.activity")
      vAnsiJoins.Add("activity_values av", "da.activity", "av.activity", "da.activity_value", "av.activity_value")
      vAnsiJoins.Add("sources s", "da.source", "s.source")
      vWhereFields.Add("event_delegate_number", mvParameters("EventDelegateNumber").Value)

      If mvParameters.Exists("Activity") Then vWhereFields.Add("da.activity", mvParameters("Activity").Value)
      If mvParameters.Exists("ActivityValue") Then vWhereFields.Add("da.activity_value", mvParameters("ActivityValue").Value)
      If mvParameters.Exists("Activities") Then vWhereFields.Add("da.activity", mvParameters("Activities").Value, CDBField.FieldWhereOperators.fwoIn)
      If mvParameters.Exists("Source") Then vWhereFields.Add("da.source", mvParameters("Source").Value)
      If mvParameters.Exists("AmendedOn") Then vWhereFields.Add("da.amended_on", CDBField.FieldTypes.cftDate, mvParameters("AmendedOn").Value)

      'If mvParameters.Exists("") Then vWhereFields.Add("da.activity", mvParameters("EventDelegateNumber").Value, CDBField.FieldWhereOperators.fwoEqual)
      vOrderBy = "activity_desc, activity_value_desc"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "delegate_activities da", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, ",,")

      Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
      Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")

      If vStatus Then pDataTable.Columns("Status").AttributeName = "status"
      If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag"

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
        If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
      Next
      If vStatus Then pDataTable.ReOrderRowsByColumn("StatusOrder")
    End Sub

    Private Sub GetWebSurveys(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "s.survey_number,s.survey_name,s.long_description,s.notes,sv.survey_version_number,sv.valid_from,sv.valid_to,sv.closing_date,sv.source"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("sv.web_publish", "Y")
      vWhereFields.Add("sv.valid_from", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("sv.valid_from#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("sv.valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("sv.valid_to#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("scg.contact_group", mvParameters.ParameterExists("ContactGroup").Value)
      vAnsiJoins.Add("survey_versions sv", "s.survey_number", "sv.survey_number")
      vAnsiJoins.Add("survey_contact_groups scg", "s.survey_number", "scg.survey_number")
      If mvParameters.ParameterExists("SurveyType").Value = "RS" OrElse mvParameters.ParameterExists("SurveyType").Value = "US" Then
        vAnsiJoins.AddLeftOuterJoin("contact_surveys cs", "s.survey_number", "cs.survey_number", "sv.survey_version_number", "cs.survey_version_number", "cs.contact_number", mvParameters("ContactNumber").Value)
        If mvParameters.ParameterExists("SurveyType").Value = "RS" Then
          vWhereFields.Add("cs.contact_survey_number", "", CDBField.FieldWhereOperators.fwoNotEqual)
        Else
          vWhereFields.Add("cs.contact_survey_number", "")
        End If
        If mvParameters.ParameterExists("RegisteredSurveyType").Value = "CRS" Then
          vWhereFields.Add("cs.completed_on", "", CDBField.FieldWhereOperators.fwoNotEqual)
        ElseIf mvParameters.ParameterExists("RegisteredSurveyType").Value = "URS" Then
          vWhereFields.Add("cs.completed_on", "")
        End If
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "surveys s", vWhereFields, "sv.closing_date", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetWebDirectoryEntries(ByVal pDataTable As CDBDataTable)
      'Following code will genrate the query for searcing contacts from directory
      'this code will execute only if "Contact Type" Module parameter is set to Contact or Both
      If mvParameters.ParameterExists("ContactType").Value = "C" OrElse mvParameters.ParameterExists("ContactType").Value = "B" Then
        Dim vAttrs As String = "v.contact_number,v.address_number,c.label_name,cp.position,m.member_number," & mvEnv.Connection.DBSpecialCol("a", "address") & ",a.town,a.country,a.county,a.sortcode,a.postcode,cu.uk,cu.country_desc,c.notes,a.house_name,a.branch,a.address_type,a.building_number,c.title,c.forenames,c.surname,c.initials,c.honorifics,c.salutation,c.preferred_forename,c.contact_type,c.ni_number,c.prefix_honorifics,c.surname_prefix,c.informal_salutation"
        Dim vFillFromSQLAttrs As String = "v.contact_number,v.address_number,c.label_name,cp.position,m.member_number," & mvEnv.Connection.DBSpecialCol("a", "address") & ",a.town,a.country,a.county,a.sortcode,a.postcode,cu.uk,cu.country_desc,c.notes,a.house_name,a.branch,a.address_type,a.building_number,c.title,c.forenames,c.surname,c.initials,c.honorifics,c.salutation,c.preferred_forename,c.contact_type,c.ni_number,c.prefix_honorifics,c.surname_prefix,c.informal_salutation"
        Dim vWhereFields As New CDBFields()
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("contacts c", "v.contact_number", "c.contact_number")
        vAnsiJoins.Add("addresses a", "v.address_number ", "a.address_number")

        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) Then
            vAttrs = vAttrs & ",cc" & vIndex & ".activity as Category" & vIndex & ", av" & vIndex & ".activity_value_desc as Activity" & vIndex
            vFillFromSQLAttrs = vFillFromSQLAttrs & ",Category" & vIndex & ",Activity" & vIndex
          Else
            vAttrs = vAttrs & ",null as Category" & vIndex & ",null as Activity" & vIndex
            vFillFromSQLAttrs = vFillFromSQLAttrs & ",null as Category" & vIndex & ",null as Activity" & vIndex
          End If
        Next vIndex

        'If Activity values are selected for configured Activity
        'add inner join for every activity which contains activity value
        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) AndAlso mvParameters.Exists("Activity" & vIndex) Then
            Dim vJoinWhereFileds As New CDBFields()
            Dim vInActivity As New StringBuilder
            vJoinWhereFileds.Add("activity", mvParameters("Category" & vIndex).Value)
            Dim vActvities As String() = mvParameters("Activity" & vIndex).Value.Split(","c)
            For Each vActivity As String In vActvities
              If vInActivity.ToString.Trim.Length > 0 Then
                vInActivity.Append(",")
                vInActivity.Append("'" & vActivity & "'")
              Else
                vInActivity.Append("'" & vActivity & "'")
              End If
            Next
            vJoinWhereFileds.Add("activity_value", vInActivity.ToString, CDBField.FieldWhereOperators.fwoIn)
            If mvParameters.Exists("ItemSelectType" & vIndex) AndAlso mvParameters("ItemSelectType" & vIndex).Value = "ALL" Then vWhereFields.Add("act" & vIndex & ".activity_count", vActvities.Length)
            vJoinWhereFileds.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vJoinWhereFileds.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            Dim vSql As New SQLStatement(mvEnv.Connection, "contact_number,activity,count(*) as activity_count", "contact_categories", vJoinWhereFileds)
            vSql.GroupBy = "contact_number,activity"
            Dim vJion As String = String.Format("({0}) Act" & vIndex, vSql.SQL)
            vAnsiJoins.Add(vJion, "act" & vIndex & ".contact_number", "v.contact_number")
            vAnsiJoins.Add("contact_categories cc" & vIndex, "act" & vIndex & ".contact_number", "cc" & vIndex & ".contact_number", "act" & vIndex & ".activity", "cc" & vIndex & ".activity")
            vAnsiJoins.Add("activity_values av" & vIndex, "cc" & vIndex & ".activity", "av" & vIndex & ".activity", "cc" & vIndex & ".activity_value", "av" & vIndex & ".activity_value")
          End If
        Next vIndex

        'If Activity values are not selected for configured Activity
        'add Left Outer Join for every activity which dose not contains activity value
        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) AndAlso Not mvParameters.Exists("Activity" & vIndex) Then
            Dim vJoinWhereFileds As New CDBFields()
            vJoinWhereFileds.Add("activity", mvParameters("Category" & vIndex).Value)
            vJoinWhereFileds.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vJoinWhereFileds.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            Dim vSql As New SQLStatement(mvEnv.Connection, "contact_number,activity,count(*) as activity_count", "contact_categories", vJoinWhereFileds)
            vSql.GroupBy = "contact_number,activity"
            Dim vJion As String = String.Format("({0}) Act" & vIndex, vSql.SQL)
            vAnsiJoins.AddLeftOuterJoin(vJion, "act" & vIndex & ".contact_number", "v.contact_number")
            vAnsiJoins.AddLeftOuterJoin("contact_categories cc" & vIndex, "act" & vIndex & ".contact_number", "cc" & vIndex & ".contact_number", "act" & vIndex & ".activity", "cc" & vIndex & ".activity")
            vAnsiJoins.AddLeftOuterJoin("activity_values av" & vIndex, "cc" & vIndex & ".activity", "av" & vIndex & ".activity", "cc" & vIndex & ".activity_value", "av" & vIndex & ".activity_value")
          End If
        Next vIndex
        vAnsiJoins.Add("members m", "v.contact_number", "m.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vAnsiJoins.AddLeftOuterJoin("contact_positions cp", "v.contact_number", "cp.contact_number", "v.address_number", "cp.address_number")
        vAnsiJoins.Add("countries cu", "a.country", "cu.country", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("c.contact_type", "C")
        vWhereFields.Add("m.cancellation_reason", "")
        vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), "Y", CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current") & "#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        If mvParameters.Exists("MemberNumber") Then vWhereFields.Add("m.member_number", mvParameters("MemberNumber").Value)
        If mvParameters.Exists("Forenames") Then vWhereFields.Add("c.forenames", mvParameters("Forenames").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Surname") Then vWhereFields.Add("c.surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("NiNumber") Then vWhereFields.Add("c.ni_number", mvParameters("NiNumber").Value)
        If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("v.contact_number", mvParameters("ContactNumber").Value)
        If mvParameters.Exists("DateOfBirth") Then vWhereFields.Add("c.date_of_birth", mvParameters("DateOfBirth").Value)
        If mvParameters.Exists("Address") Then vWhereFields.Add("a.address", mvParameters("Address").Value, CDBField.FieldWhereOperators.fwoLike)
        If mvParameters.Exists("Town") Then vWhereFields.Add("a.town", mvParameters("Town").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Postcode") Then vWhereFields.Add("a.postcode", mvParameters("Postcode").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Country") Then vWhereFields.Add("a.country", mvParameters("Country").Value)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, mvParameters("ViewName").Value & " v", vWhereFields, "c.surname", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFillFromSQLAttrs.Replace("null as ", ""), "ADDRESS_LINE,CONTACT_NAME")
      End If

      'Following code will generate the query for searching Organisations from directory
      'this code will execute only if "Contact Type" Module parameter is set to Organisation or Both
      If mvParameters.ParameterExists("ContactType").Value = "O" OrElse mvParameters.ParameterExists("ContactType").Value = "B" Then
        Dim vAttrs As String = "v.contact_number,v.address_number,c.label_name,cp.position,m.member_number," & mvEnv.Connection.DBSpecialCol("a", "address") & ",a.town,a.country,a.county,a.sortcode,a.postcode,cu.uk,cu.country_desc,o.notes,a.house_name,a.branch,a.address_type,a.building_number,c.title,c.forenames,c.surname,c.initials,c.honorifics,c.salutation,c.preferred_forename,c.contact_type,c.ni_number,c.prefix_honorifics,c.surname_prefix,c.informal_salutation"
        Dim vFillFromSQLAttrs As String = "v.contact_number,v.address_number,c.label_name,cp.position,m.member_number," & mvEnv.Connection.DBSpecialCol("a", "address") & ",a.town,a.country,a.county,a.sortcode,a.postcode,cu.uk,cu.country_desc,o.notes,a.house_name,a.branch,a.address_type,a.building_number,c.title,c.forenames,c.surname,c.initials,c.honorifics,c.salutation,c.preferred_forename,c.contact_type,c.ni_number,c.prefix_honorifics,c.surname_prefix,c.informal_salutation"
        Dim vWhereFields As New CDBFields()
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("contacts c", "v.contact_number", "c.contact_number")
        vAnsiJoins.Add("organisations o", "v.contact_number", "o.organisation_number", "v.address_number", "o.address_number")
        vAnsiJoins.Add("addresses a", "v.address_number ", "a.address_number")

        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) Then
            vAttrs = vAttrs & ",oc" & vIndex & ".activity as Category" & vIndex & ", av" & vIndex & ".activity_value_desc as Activity" & vIndex
            vFillFromSQLAttrs = vFillFromSQLAttrs & ",Category" & vIndex & ",Activity" & vIndex
          Else
            vAttrs = vAttrs & ",null as Category" & vIndex & ",null as Activity" & vIndex
            vFillFromSQLAttrs = vFillFromSQLAttrs & ",null as Category" & vIndex & ",null as Activity" & vIndex
          End If
        Next vIndex

        'If Activity values are selected for Activity
        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) And mvParameters.Exists("Activity" & vIndex) Then
            Dim vJoinWhereFileds As New CDBFields()
            Dim vInActivity As New StringBuilder
            vJoinWhereFileds.Add("activity", mvParameters("Category" & vIndex).Value)
            Dim vActvities As String() = mvParameters("Activity" & vIndex).Value.Split(","c)
            For Each vActivity As String In vActvities
              If vInActivity.ToString.Trim.Length > 0 Then
                vInActivity.Append(",")
                vInActivity.Append("'" & vActivity & "'")
              Else
                vInActivity.Append("'" & vActivity & "'")
              End If
            Next
            vJoinWhereFileds.Add("activity_value", vInActivity.ToString, CDBField.FieldWhereOperators.fwoIn)
            If mvParameters.Exists("ItemSelectType" & vIndex) AndAlso mvParameters("ItemSelectType" & vIndex).Value = "ALL" Then vWhereFields.Add("act" & vIndex & ".activity_count", vActvities.Length)
            vJoinWhereFileds.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vJoinWhereFileds.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            Dim vSql As New SQLStatement(mvEnv.Connection, "organisation_number,activity,count(*) as activity_count", "organisation_categories", vJoinWhereFileds)
            vSql.GroupBy = "organisation_number,activity"
            Dim vJion As String = String.Format("({0}) Act" & vIndex, vSql.SQL)
            vAnsiJoins.Add(vJion, "act" & vIndex & ".organisation_number", "v.contact_number")
            vAnsiJoins.Add("organisation_categories oc" & vIndex, "act" & vIndex & ".organisation_number", "oc" & vIndex & ".organisation_number", "act" & vIndex & ".activity", "oc" & vIndex & ".activity")
            vAnsiJoins.Add("activity_values av" & vIndex, "oc" & vIndex & ".activity", "av" & vIndex & ".activity", "oc" & vIndex & ".activity_value", "av" & vIndex & ".activity_value")
          End If
        Next vIndex

        'If Activity values are not selected for Activity
        For vIndex As Integer = 1 To 6
          If mvParameters.Exists("Category" & vIndex) AndAlso Not mvParameters.Exists("Activity" & vIndex) Then
            Dim vJoinWhereFileds As New CDBFields()
            vJoinWhereFileds.Add("activity", mvParameters("Category" & vIndex).Value)
            vJoinWhereFileds.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vJoinWhereFileds.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            Dim vSql As New SQLStatement(mvEnv.Connection, "organisation_number,activity,count(*) as activity_count", "organisation_categories", vJoinWhereFileds)
            vSql.GroupBy = "organisation_number,activity"
            Dim vJion As String = String.Format("({0}) Act" & vIndex, vSql.SQL)
            vAnsiJoins.AddLeftOuterJoin(vJion, "act" & vIndex & ".organisation_number", "v.contact_number")
            vAnsiJoins.AddLeftOuterJoin("organisation_categories oc" & vIndex, "act" & vIndex & ".organisation_number", "oc" & vIndex & ".organisation_number", "act" & vIndex & ".activity", "oc" & vIndex & ".activity")
            vAnsiJoins.AddLeftOuterJoin("activity_values av" & vIndex, "oc" & vIndex & ".activity", "av" & vIndex & ".activity", "oc" & vIndex & ".activity_value", "av" & vIndex & ".activity_value")
          End If
        Next vIndex
        'If mvParameters.ParameterExists("ContactType").Value = "O" Then mvResultColumns = mvResultColumns & ",AddressLine,ContactName"
        vAnsiJoins.Add("members m", "v.contact_number", "m.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vAnsiJoins.AddLeftOuterJoin("contact_positions cp", "v.contact_number", "cp.contact_number", "v.address_number", "cp.address_number")
        vAnsiJoins.Add("countries cu", "a.country", "cu.country", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        vWhereFields.Add("c.contact_type", "O")
        vWhereFields.Add("m.cancellation_reason", "")
        vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), "Y", CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current") & "#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        If mvParameters.Exists("MemberNumber") Then vWhereFields.Add("m.member_number", mvParameters("MemberNumber").Value)
        If mvParameters.Exists("Forenames") Then vWhereFields.Add("c.forenames", mvParameters("Forenames").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Surname") Then vWhereFields.Add("c.surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("NiNumber") Then vWhereFields.Add("c.ni_number", mvParameters("NiNumber").Value)
        If mvParameters.Exists("ContactNumber") Then vWhereFields.Add("v.contact_number", mvParameters("ContactNumber").Value)
        If mvParameters.Exists("DateOfBirth") Then vWhereFields.Add("c.date_of_birth", mvParameters("DateOfBirth").Value)
        If mvParameters.Exists("Address") Then vWhereFields.Add("a.address", mvParameters("Address").Value, CDBField.FieldWhereOperators.fwoLike)
        If mvParameters.Exists("Town") Then vWhereFields.Add("a.town", mvParameters("Town").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Postcode") Then vWhereFields.Add("a.postcode", mvParameters("Postcode").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If mvParameters.Exists("Country") Then vWhereFields.Add("a.country", mvParameters("Country").Value)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, mvParameters("ViewName").Value & " v", vWhereFields, "c.label_name", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFillFromSQLAttrs.Replace("null as ", ""), "ADDRESS_LINE,CONTACT_NAME")
      End If

      Dim vTempDataTable As New CDBDataTable
      For Each vDataRow As CDBDataRow In pDataTable.Rows
        vTempDataTable.Rows.Add(vDataRow)
      Next

      'added column to check for recodrd is duplicate or not                       
      pDataTable.AddColumn("IsDuplicate", CDBField.FieldTypes.cftCharacter)


      'Following code creats list of activity values for each activity and adding it to specific column.
      Dim vContactList As New List(Of String)
      For Each vRecord As CDBDataRow In pDataTable.Rows
        Dim vActivities1 As New List(Of String)
        Dim vActivities2 As New List(Of String)
        Dim vActivities3 As New List(Of String)
        Dim vActivities4 As New List(Of String)
        Dim vActivities5 As New List(Of String)
        Dim vActivities6 As New List(Of String)
        If Not vContactList.Contains(vRecord.Item("ContactNumber").ToString) Then
          vContactList.Add(vRecord.Item("ContactNumber").ToString)
          For Each vActivityRecord As CDBDataRow In vTempDataTable.Rows
            If vRecord.Item("ContactNumber").ToString = vActivityRecord.Item("ContactNumber").ToString Then
              If mvParameters.Exists("Category1") AndAlso mvParameters("Category1").Value = vActivityRecord.Item("Category1").ToString Then
                If Not vActivities1.Contains(vActivityRecord.Item("Activity1").ToString) Then vActivities1.Add(vActivityRecord.Item("Activity1").ToString)
              End If
              If mvParameters.Exists("Category2") AndAlso mvParameters("Category2").Value = vActivityRecord.Item("Category2").ToString Then
                If Not vActivities2.Contains(vActivityRecord.Item("Activity2").ToString) Then vActivities2.Add(vActivityRecord.Item("Activity2").ToString)
              End If
              If mvParameters.Exists("Category3") AndAlso mvParameters("Category3").Value = vActivityRecord.Item("Category3").ToString Then
                If Not vActivities3.Contains(vActivityRecord.Item("Activity3").ToString) Then vActivities3.Add(vActivityRecord.Item("Activity3").ToString)
              End If
              If mvParameters.Exists("Category4") AndAlso mvParameters("Category4").Value = vActivityRecord.Item("Category4").ToString Then
                If Not vActivities4.Contains(vActivityRecord.Item("Activity4").ToString) Then vActivities4.Add(vActivityRecord.Item("Activity4").ToString)
              End If
              If mvParameters.Exists("Category5") AndAlso mvParameters("Category5").Value = vActivityRecord.Item("Category5").ToString Then
                If Not vActivities5.Contains(vActivityRecord.Item("Activity5").ToString) Then vActivities5.Add(vActivityRecord.Item("Activity5").ToString)
              End If
              If mvParameters.Exists("Category6") AndAlso mvParameters("Category6").Value = vActivityRecord.Item("Category6").ToString Then
                If Not vActivities6.Contains(vActivityRecord.Item("Activity6").ToString) Then vActivities6.Add(vActivityRecord.Item("Activity6").ToString)
              End If
            End If
          Next
          Dim vDisplayActivity As StringBuilder
          If vActivities1.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities1
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity1") = vDisplayActivity.ToString
          End If
          If vActivities2.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities2
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity2") = vDisplayActivity.ToString
          End If
          If vActivities3.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities3
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity3") = vDisplayActivity.ToString
          End If
          If vActivities4.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities4
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity4") = vDisplayActivity.ToString
          End If
          If vActivities5.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities5
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity5") = vDisplayActivity.ToString
          End If
          If vActivities6.Count > 0 Then
            vDisplayActivity = New StringBuilder
            For Each vActivity As String In vActivities6
              If vDisplayActivity.ToString.Trim.Length > 0 Then
                vDisplayActivity.Append(", ")
                vDisplayActivity.Append(vActivity)
              Else
                vDisplayActivity.Append(vActivity)
              End If
            Next
            vRecord.Item("Activity6") = vDisplayActivity.ToString
          End If
        Else
          vRecord.Item("IsDuplicate") = "Y"
        End If
      Next

      'this loop is to remove duplicate recodrds from the datatable
      'it is revesed to avoid index chage
      For vRowIndex As Integer = pDataTable.Rows.Count - 1 To 0 Step -1
        Dim vDataRow As CDBDataRow
        vDataRow = pDataTable.Rows(vRowIndex)
        If vDataRow.Item("IsDuplicate") IsNot Nothing AndAlso vDataRow.Item("IsDuplicate").ToString = "Y" Then
          pDataTable.Rows.RemoveAt(vRowIndex)
        End If
      Next
      pDataTable.ReOrderRowsByColumn("ContactName")
    End Sub

    Private Sub GetWebDocuments(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "d.web_document_number,d.web_document_title,d.file_name,d.web_document_extension,e.image_name,d.description,d.mime_type,d.view_name,d.valid_from,d.valid_to,d.download_count,d.last_downloaded_on,d.created_by,d.created_on,d.amended_by,d.amended_on"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("d.valid_from", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("d.valid_from#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("d.valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("d.valid_to#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      If mvParameters.Exists("Views") AndAlso mvParameters("Views").Value.Length > 0 Then
        vWhereFields.Add("d.view_name", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("d.view_name#1", mvParameters("Views").Value, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Else
        vWhereFields.Add("d.view_name")
      End If
      If mvParameters.Exists("Topic") AndAlso mvParameters("Topic").Value.Length > 0 Then
        vAnsiJoins.Add("web_document_topics t", "d.web_document_number", "t.web_document_number")
        vWhereFields.Add("t.topic", CDBField.FieldTypes.cftCharacter, mvParameters("Topic").Value)
      End If
      If mvParameters.Exists("SearchDocument") Then
        vWhereFields.Add("d.web_document_title", mvParameters("SearchDocument").Value & "*", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vWhereFields.Add("d.description", mvParameters("SearchDocument").Value & "*", CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
      End If
      vAnsiJoins.AddLeftOuterJoin("web_document_extensions e", "d.web_document_extension", "e.web_document_extension")
      If mvParameters.Exists("WebDocumentNumber") Then
        vWhereFields.Add("d.web_document_number", CDBField.FieldTypes.cftLong, mvParameters("WebDocumentNumber").Value)
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "web_documents d", vWhereFields, "d.web_document_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)
    End Sub

    Private Sub GetWebRelatedOrganisations(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "v.organisation_number,v.address_number,v.organisation_name," & mvEnv.Connection.DBSpecialCol("v", "status") & ",v.status_desc,v.organisation_group,v.organisation_group_desc," & mvEnv.Connection.DBSpecialCol("a", "address") & ",a.town,a.county,a.sortcode,a.postcode,a.country,v.country_desc,v.uk,a.house_name,a.branch,a.address_type,a.building_number,v.organisation_number as parent_organisation,null as child_count"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vChildOrgDataTable As New CDBDataTable
      vAnsiJoins.Add("addresses a", "v.address_number", "a.address_number")
      vWhereFields.Add("v.contact_number", mvParameters("ContactNumber").Value, CDBField.FieldWhereOperators.fwoEqual)
      If mvParameters.Exists("OrganisationNumber") Then vWhereFields.Add("v.organisation_number", mvParameters("OrganisationNumber").Value, CDBField.FieldWhereOperators.fwoEqual)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "VPortalRelatedOrgs v", vWhereFields, "v.organisation_name", vAnsiJoins)
      'Get all the parent organisations
      Dim vSupppressChildLinksOnLine1 As Boolean = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("null as ", ""), "ADDRESS_LINE")
      If pDataTable.Rows.Count = 0 Then vSupppressChildLinksOnLine1 = False
      If (mvParameters.Exists("ShowChildren") AndAlso mvParameters("ShowChildren").Value <> "N") OrElse mvParameters.Exists("OrganisationNumber") Then
        'Find any child organisations for the parents
        vChildOrgDataTable = GetWebChildOrganisations(pDataTable)
      End If

      If mvParameters.Exists("ShowChildren") AndAlso (mvParameters("ShowChildren").Value = "C" OrElse mvParameters("ShowChildren").Value = "D") Then
        For vLoop As Integer = 1 To 2
          For vRowNumber As Integer = 0 To pDataTable.Rows.Count - 1
            Dim vParentDr As CDBDataRow = pDataTable.Rows(vRowNumber)
            Dim vOrgNumber As String = vParentDr.Item("OrganisationNumber")
            Dim vChildCount As Integer = 0
            If vChildOrgDataTable IsNot Nothing AndAlso vChildOrgDataTable.Rows.Count > 0 Then
              For Each vChildDr As CDBDataRow In vChildOrgDataTable.Rows
                Dim vParent As String = vChildDr.Item("ParentOrganisation")
                Dim vChild As String = vChildDr.Item("OrganisationNumber")
                If vOrgNumber <> vChild AndAlso vOrgNumber = vParent Then
                  vChildCount = vChildCount + 1
                End If
              Next
              If vRowNumber = 0 AndAlso vSupppressChildLinksOnLine1 = True AndAlso mvParameters.Exists("OrganisationNumber") Then
                'First row is parent row we have just selected so return a child count of zero for this row
                vParentDr.Item("ChildCount") = "0"
              Else
                vParentDr.Item("ChildCount") = vChildCount.ToString
              End If
            End If
          Next
          'This is just so that the ChildCount column can be set.
          If vLoop = 1 AndAlso mvParameters.Exists("OrganisationNumber") Then
            'If we do not have an Organisation Number then we are on the top level, so do not need to get additional child organisations
            'If this is the first time through this loop with an Organisation Number, then we are not on the top level so want to get additional child organisations
            vChildOrgDataTable = GetWebChildOrganisations(vChildOrgDataTable)
          Else
            vChildOrgDataTable = Nothing
            vLoop += 1
          End If
        Next
      End If
    End Sub

    Private Function GetWebChildOrganisations(ByVal pDataTable As CDBDataTable) As CDBDataTable
      Dim vOrganisationNumbers As New StringBuilder
      For Each vDr As CDBDataRow In pDataTable.Rows
        Dim vParentOrgNum As String = vDr.Item("OrganisationNumber")
        If vOrganisationNumbers.ToString.Trim.Length > 0 Then
          vOrganisationNumbers.Append(",'" & vParentOrgNum & "'")
        Else
          vOrganisationNumbers.Append("'" & vParentOrgNum & "'")
        End If
      Next
      Dim vChildOrgDataTable As CDBDataTable = Nothing
      If vOrganisationNumbers.ToString.Length = 0 AndAlso mvParameters.Exists("OrganisationNumber") Then
        'We don't have a position at the parent organisation so just use it's number
        vOrganisationNumbers.Append("'" & mvParameters("OrganisationNumber").Value & "'")
      End If
      If vOrganisationNumbers.ToString.Length > 0 Then
        vChildOrgDataTable = GetChildOrg(vOrganisationNumbers.ToString)
      End If
      If vChildOrgDataTable IsNot Nothing _
      AndAlso (mvParameters.Exists("ShowChildren") AndAlso mvParameters("ShowChildren").Value = "D") OrElse mvParameters.Exists("OrganisationNumber") Then
        For Each vChildDr As CDBDataRow In vChildOrgDataTable.Rows
          pDataTable.Rows.Add(vChildDr)
        Next
      End If
      Return vChildOrgDataTable
    End Function

    Private Function GetChildOrg(ByVal pOrgNum As String) As CDBDataTable
      Dim vChildOrgTable As New CDBDataTable
      If mvResultColumns.Length > 0 Then vChildOrgTable.AddColumnsFromList(mvResultColumns)
      Dim vAttrs As String = "oc.organisation_number,oc.address_number,oc.name,oc.status,s.status_desc,oc.organisation_group,og.organisation_group_desc,a.address,a.town,a.county,a.sortcode,a.postcode,a.country,c.country_desc,c.uk,a.house_name,a.branch,a.address_type,a.building_number,o.organisation_number as parent_organisation,null as child_count"
      Dim vColList As String = "oc.organisation_number,oc.address_number,oc.name,oc.status,s.status_desc,oc.organisation_group,og.organisation_group_desc,a.address,a.town,a.county,a.sortcode,a.postcode,a.country,c.country_desc,c.uk,a.house_name,a.branch,a.address_type,a.building_number,parent_organisation,null as child_count"
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("organisation_links ol ", "o.organisation_number", "ol.organisation_number_2")
      vAnsiJoins.Add("relationships r", "ol.relationship", "r.relationship")
      vAnsiJoins.Add("organisations oc", "ol.organisation_number_1", "oc.organisation_number")
      vAnsiJoins.Add("addresses a", "oc.address_number", "a.address_number")
      vAnsiJoins.Add("countries c", "a.country", "c.country", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("statuses s", "oc.status", "s.status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("organisation_groups og", "oc.organisation_group", "og.organisation_group", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vWhereFields.Add("o.organisation_number", pOrgNum, CDBField.FieldWhereOperators.fwoIn)
      vWhereFields.Add("r.relationship", "SELECT DISTINCT parent_relationship FROM relationships WHERE parent_relationship IS NOT NULL", CDBField.FieldWhereOperators.fwoIn)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "organisations o", vWhereFields, "oc.name", vAnsiJoins)
      vChildOrgTable.FillFromSQL(mvEnv, vSQLStatement, vColList.Replace("null as ", ""), "ADDRESS_LINE")
      Return vChildOrgTable
    End Function

    Private Sub GetWebRelatedContacts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vContact As New Contact(mvEnv)

      Dim vAttrs As String = "vn.organisation_number,vn.contact_number,vn.address_number,cp.contact_position_number,vn.title,vn.initials,vn.forenames,vn.surname,vn.label_name,vn.date_of_birth, Null as EmailAddress,cp.started,cp.finished,cp.mail," & mvEnv.Connection.DBSpecialCol("cp", "current") & ",cp.position_location,cp.position,cp.position_function,pf.position_function_desc,cp.position_seniority,ps.position_seniority_desc,vn.status,vn.status_desc,m.member_number,m.membership_status,ms.membership_status_desc,c.ni_number,vn.contact_group,vn.contact_group_desc,o.contact_number AS DefaultContactNumber,Null as DefaultNo," & vContact.GetRecordSetFieldsPhone.Replace("c.contact_number,", "")
      Dim vFields As String = "vn.organisation_number,vn.contact_number,vn.address_number,cp.contact_position_number,vn.title,vn.initials,vn.forenames,vn.surname,vn.label_name,vn.date_of_birth, Null as EmailAddress,cp.started,cp.finished,cp.mail," & mvEnv.Connection.DBSpecialCol("cp", "current") & ",cp.position_location,cp.position,cp.position_function,pf.position_function_desc,cp.position_seniority,ps.position_seniority_desc,vn.status,vn.status_desc,m.member_number,m.membership_status,ms.membership_status_desc,c.ni_number,vn.contact_group,vn.contact_group_desc,DefaultContactNumber,Null as [Default]"
      vWhereFields.Add("vn.organisation_number", mvParameters.ParameterExists("OrganisationNumber").IntegerValue)
      vWhereFields.Add("m.cancellation_reason", "")
      vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), "Y")
      vAnsiJoins.Add("contact_positions cp", "vn.organisation_number", "cp.organisation_number", "vn.contact_number", "cp.contact_number", "vn.address_number", "cp.address_number")
      vAnsiJoins.Add("contacts c", "vn.contact_number", "c.contact_number")
      vAnsiJoins.Add("organisations o ", "vn.organisation_number", "o.organisation_number")
      vAnsiJoins.Add("position_functions pf", "cp.position_function", "pf.position_function", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("position_seniorities ps", "cp.position_seniority", "ps.position_seniority", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("members m", "vn.contact_number", "m.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("membership_statuses ms", "m.membership_status", "ms.membership_status", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, mvParameters("ViewName").Value & " vn", vWhereFields, "vn.contact_number", vAnsiJoins)
      vSQLStatement.Distinct = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields, "CONTACT_TELEPHONE")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("DefaultContactNumber") = vRow.Item("ContactNumber") Then
          vRow.Item("Default") = "Yes"
        End If
      Next
      If pDataTable.Rows.Count > 0 Then SetEmailAddressOfContacts(vSQLStatement, pDataTable)
    End Sub

    Private Sub SetEmailAddressOfContacts(ByVal pContactSQL As SQLStatement, ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      Dim vTable As New CDBDataTable
      Dim vPreffered As Boolean
      Dim vDefault As Boolean
      Dim vOrderBy As String = ""
      Dim vDevice As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEmailDevice)

      pContactSQL.FieldNames = "c.contact_number"
      pContactSQL.Distinct = True
      pContactSQL.OrderBy = ""
      vWhereFields.Add("c.contact_number", String.Format("( {0} )", pContactSQL.SQL), CDBField.FieldWhereOperators.fwoIn)

      vWhereFields.Add("cm.device", vDevice)
      vWhereFields.Add("cm.valid_from", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("cm.valid_from#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("cm.valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("cm.valid_to#1", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vAnsiJoins.Add("communications cm", "c.contact_number", " cm.contact_number ", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vOrderBy = "cm.contact_number,cm.preferred_method desc,cm.device_default desc"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "cm.contact_number,cm.device," & mvEnv.Connection.DBSpecialCol("cm", "number") & ",cm.valid_from, cm.preferred_method, cm.device_default", "contacts c", vWhereFields, vOrderBy, vAnsiJoins)
      vTable.FillFromSQL(mvEnv, vSQLStatement, False)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vPreffered = False
        vDefault = False
        For Each vCommRow As CDBDataRow In vTable.Rows
          If vRow.Item("ContactNumber").ToString = vCommRow.Item("contact_number") Then
            If vCommRow.Item("preferred_method") = "Y" Then
              vRow.Item("EmailAddress") = vCommRow.Item("number")
              If vPreffered Then Exit For
              vPreffered = True
            ElseIf vCommRow.Item("device_default") = "Y" Then
              If Not vPreffered Then
                vRow.Item("EmailAddress") = vCommRow.Item("number")
              End If
              If vDefault Or vPreffered Then Exit For
              vDefault = True
            ElseIf vCommRow.Item("preferred_method") = "N" AndAlso vCommRow.Item("device_default") = "N" Then
              If vPreffered Or vDefault Then
                Exit For
              Else
                vRow.Item("EmailAddress") = vCommRow.Item("number")
              End If
            End If
          End If
        Next
      Next
    End Sub

    Private Sub GetContactLoans(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "l.order_number,o.contact_number,loan_number,lt.loan_type,lt.loan_type_desc,loan_amount,balance,l.created_by,l.created_on,"
      vAttrs &= "order_date,interest_rate,interest_calculated_date,next_payment_due,fixed_monthly_amount,loan_term,frequency_amount,l.source,s.source_desc,"
      vAttrs &= "l.cancellation_reason,cancellation_reason_desc,l.cancellation_source,cs.source_desc AS cancellation_source_desc,l.cancelled_by,l.cancelled_on,"
      vAttrs &= "direct_debit,credit_card,pf.frequency,o.payment_method,payment_method_desc,expiry_date"
      vAttrs &= If(mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoanInterestRates) = True, ",interest_capitalisation_date,interest_capitalisation_amount", ",,")
      Dim vOrderBy As String = mvEnv.Connection.DBOrderByNullsFirstDesc("l.cancelled_on") & ", order_date"

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("orders o", "l.order_number", "o.order_number")
        .Add("loan_types lt", "l.loan_type", "lt.loan_type")
        .Add("sources s", "l.source", "s.source")
        .Add("contacts c", "o.contact_number", "c.contact_number")
        .Add("payment_frequencies pf", "o.payment_frequency", "pf.payment_frequency")
        .Add("payment_methods pm", "o.payment_method", "pm.payment_method")
        .AddLeftOuterJoin("cancellation_reasons cr", "l.cancellation_reason", "cr.cancellation_reason")
        .AddLeftOuterJoin("sources cs", "l.cancellation_source", "cs.source")
      End With
      Dim vWherefields As New CDBFields(New CDBField("o.contact_number", mvParameters("ContactNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "loans l", vWherefields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs)

      Dim vPP As New PaymentPlan
      vPP.Init(mvEnv)
      For Each vRow As CDBDataRow In pDataTable.Rows
        Dim vNextPayment As Double = DoubleValue(vRow.Item("FrequencyAmount"))
        Dim vNextPayDue As String = vRow.Item("NextPaymentDue")
        Dim vGotAutoPayment As Boolean = (vRow.Item("DirectDebitStatus") = "Y" OrElse vRow.Item("CreditCardStatus") = "Y")
        vPP.GetNextSchedulePaymentInfo(vNextPayment, vNextPayDue, IntegerValue(vRow.Item("PaymentPlanNumber")), DoubleValue(vRow.Item("Balance")), vGotAutoPayment, IntegerValue(vRow.Item("PaymentFrequencyFrequency")), True)
        vRow.Item("FrequencyAmount") = vNextPayment.ToString("F")
        vRow.Item("NextPaymentDue") = vNextPayDue
        If vRow.Item("LoanTerm").Length > 0 Then
          vRow.Item("ExpiryDate") = CDate(vRow.Item("StartDate")).AddYears(IntegerValue(vRow.Item("LoanTerm"))).ToString(CAREDateFormat)
        ElseIf DoubleValue(vRow.Item("MonthlyPaymentAmount")) = 0 Then
          vRow.Item("ExpiryDate") = CDate(vRow.Item("StartDate")).AddYears(101).ToString(CAREDateFormat)
        End If
      Next

    End Sub

    Private Sub GetDirectoryUsage(ByVal pDataTable As CDBDataTable)
      ' Need to write SQL here
      ' need columns  Data,Usage
      Dim vList As New ArrayListEx
      If mvParameters.Exists("CommunicationUsage1") Then vList.Add(mvParameters("CommunicationUsage1").Value)
      If mvParameters.Exists("CommunicationUsage2") Then vList.Add(mvParameters("CommunicationUsage2").Value)
      If mvParameters.Exists("CommunicationUsage3") Then vList.Add(mvParameters("CommunicationUsage3").Value)
      If mvParameters.Exists("CommunicationUsage4") Then vList.Add(mvParameters("CommunicationUsage4").Value)

      If vList.Count > 0 Then ' Communication - Data
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("com.communication_usage", vList.CSStringList, CDBField.FieldWhereOperators.fwoIn)
        vWhereFields.Add("com.contact_number", mvParameters("ContactNumber").IntegerValue)
        vWhereFields.Add("com.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("com.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("contact_communication_usages ccu", "ccu.communication_number", "com.communication_number")

        Dim vCommunication As New Communication(mvEnv)
        Dim vFields As String = vCommunication.GetRecordSetFieldsPhone

        ' Communication - Data (person)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "device,com.communication_usage," & vFields, "communications com", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, ",device,,communication_usage,PHONE_NUMBER")

        vWhereFields.Clear()
        vWhereFields.Add("com.communication_usage", vList.CSStringList, CDBField.FieldWhereOperators.fwoIn)
        vWhereFields.Add("ca.contact_number", mvParameters("ContactNumber").IntegerValue)
        vWhereFields.Add("com.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("com.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
        vWhereFields.Add("com.contact_number")

        vAnsiJoins.Add("contact_addresses ca", "ca.address_number", "com.address_number")

        ' Communication - Data (employment)
        vSQLStatement = New SQLStatement(mvEnv.Connection, "device,com.communication_usage," & vFields, "communications com", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, ",device,,communication_usage,PHONE_NUMBER")
      End If

      vList.Clear()
      If mvParameters.Exists("AddressUsage") Then vList.Add(mvParameters("AddressUsage").Value)

      If vList.Count > 0 Then ' Address - Data
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("cau.address_usage", vList.CSStringList, CDBField.FieldWhereOperators.fwoIn)
        vWhereFields.Add("ca.contact_number", mvParameters("ContactNumber").IntegerValue)
        vWhereFields.Add("ca.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("ca.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("contact_addresses ca", "ca.address_number", "a.address_number")
        vAnsiJoins.Add("contact_address_usages cau", "cau.address_number", "a.address_number", "cau.contact_number", "ca.contact_number") 'J1521: Added join for contact_number
        vAnsiJoins.Add("countries co", "a.country", "co.country")

        Dim vAddress As New Address(mvEnv)
        Dim vFields As String = vAddress.GetRecordSetFieldsCountry

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "address_usage," & vFields, "addresses a", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, "address_type,,address_usage,,ADDRESS_MULTI_LINE")
      End If
    End Sub

    Private Sub GetContactExamSummary(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "exam_student_header_id,esh.exam_unit_id,exam_base_unit_id,exam_unit_code,exam_unit_description,first_session_id,fs.exam_session_code,fs.exam_session_description,last_session_id,ls.exam_session_code AS last_exam_session_code,ls.exam_session_description AS last_exam_session_desc,last_marked_date,last_graded_date,eul.exam_unit_link_id,eul.parent_unit_link_id,esh.created_by,esh.created_on,esh.amended_by,esh.amended_on"
      Dim vWhereFields As New CDBFields(New CDBField("contact_number", mvParameters("ContactNumber").IntegerValue))
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_unit_links eul", "esh.exam_unit_id", "eul.exam_unit_id_2", "esh.exam_unit_link_id", "eul.exam_unit_link_id")
      vAnsiJoins.Add("exam_units eu", "eul.exam_unit_id_2", "eu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_sessions fs", "esh.first_session_id", "fs.exam_session_id")
      vAnsiJoins.AddLeftOuterJoin("exam_sessions ls", "esh.last_session_id", "ls.exam_session_id")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "exam_student_header esh", vWhereFields, mvEnv.Connection.DBIsNull("eu.sequence_number", "10000"), vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetContactExamSummaryItems(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "esuh.exam_student_unit_header_id,esuh.exam_student_header_id,esuh.exam_unit_id,exam_unit_code,exam_unit_description,attempts,current_mark,current_grade,current_result,eg.exam_grade_desc,"
      vAttrs &= "{0},esuh.first_passed,esuh.expires,eul.exam_unit_link_id,eul.parent_unit_link_id,esuh.created_by,esuh.created_on,esuh.amended_by,esuh.amended_on,{1}"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAttrs &= ",esuh.results_release_date,esuh.previous_mark,esuh.previous_grade,esuh.previous_result,{2}"
      Else
        vAttrs &= ",,,,"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLastExamDate) Then
        vAttrs += ",esuh.last_exam_date,esuh.last_exam_date"
      End If
      Dim vFields As String = String.Format(vAttrs, "current_result_desc", "can_edit_results", "previous_grade_desc,previous_result_desc")
      vAttrs = RemoveBlankItems(String.Format(vAttrs, "current_result AS current_result_desc", "'Y' AS can_edit_results", "peg.exam_grade_desc AS previous_grade_desc,previous_result AS previous_result_desc"))

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamStudentUnitHeaderId", "exam_student_unit_header_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamStudentHeaderId", "exam_student_header_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamUnitId", "exam_unit_id")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_unit_links eul", "esuh.exam_unit_link_id", "eul.exam_unit_link_id")
      vAnsiJoins.Add("exam_units eu", "esuh.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "esuh.current_grade", "eg.exam_grade")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAnsiJoins.AddLeftOuterJoin("exam_grades peg", "esuh.previous_grade", "peg.exam_grade")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "exam_student_unit_header esuh", vWhereFields, "", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then RestrictExamResults(pDataTable, "Current")
      GetLookupData(pDataTable, "CurrentResult", "exam_booking_units", "original_result")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLastExamDate) Then GetChildLastExamDate(pDataTable)
    End Sub

    Private Sub GetContactExamSummaryList(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "esuh.exam_student_unit_header_id,esuh.exam_student_header_id,esuh.exam_unit_id,exam_unit_code,exam_unit_description,attempts,current_mark,current_grade,current_result,eg.exam_grade_desc,"
      vAttrs &= "{0},esuh.first_passed,esuh.expires,eul.exam_unit_link_id,eul.parent_unit_link_id,esuh.created_by,esuh.created_on,esuh.amended_by,esuh.amended_on,{1}"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAttrs &= ",esuh.results_release_date,esuh.previous_mark,esuh.previous_grade,esuh.previous_result,{2}"
      Else
        vAttrs &= ",,,,"
      End If
      Dim vFields As String = String.Format(vAttrs, "current_result_desc", "can_edit_results", "previous_grade_desc,previous_result_desc")
      vAttrs = RemoveBlankItems(String.Format(vAttrs, "current_result AS current_result_desc", "'Y' AS can_edit_results", "peg.exam_grade_desc AS previous_grade_desc,previous_result AS previous_result_desc"))

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamStudentUnitHeaderId", "exam_student_unit_header_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamStudentHeaderId", "exam_student_header_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamUnitId", "exam_unit_id_1")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ContactNumber", "contact_number")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_student_header esh", "esuh.exam_student_header_id", "esh.exam_student_header_id")
      vAnsiJoins.Add("exam_units eu", "esuh.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "esuh.current_grade", "eg.exam_grade")
      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "esuh.exam_unit_id", "eul.exam_unit_id_2", "esuh.exam_unit_link_id", "eul.exam_unit_link_id")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAnsiJoins.AddLeftOuterJoin("exam_grades peg", "esuh.previous_grade", "peg.exam_grade")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "exam_student_unit_header esuh", vWhereFields, "", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then RestrictExamResults(pDataTable, "Current")
      GetLookupData(pDataTable, "CurrentResult", "exam_booking_units", "original_result")
    End Sub

    Private Sub GetContactExamDetails(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "eb.exam_booking_id,eb.exam_session_id,eb.exam_centre_id,eb.exam_unit_id,exam_session_code,exam_session_description"
      vFields += ",exam_unit_code,exam_unit_description,exam_centre_code,exam_centre_description,eb.amount,eb.batch_number,eb.transaction_number"
      vFields += ",bt.contact_number,bt.transaction_date,eb.cancellation_reason,eb.cancellation_source,cancellation_reason_desc,source_desc,eb.cancelled_on,eb.cancelled_by,eb.special_requirements"
      vFields += ",eul.exam_unit_link_id,eul.parent_unit_link_id"
      vFields += ",eb.created_by,eb.created_on,eb.amended_by,eb.amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading) Then
        vFields &= ",ebu.course_start_date,ebu.exam_assessment_language,eal.exam_assessment_language_desc"
      Else
        vFields &= ",,,"
      End If
      vFields &= ","
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then vFields &= "COALESCE(ecu.local_name, ecbu.local_name) local_name"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamStudyModes) Then
        vFields &= ",eb.study_mode,sm.study_mode_desc"
      Else
        vFields &= ",,"
      End If

      'Build nested SQL for ExamBookingUnits data
      Dim vNestedAttrs As String = "ebu.exam_booking_id,MIN(ebu.exam_assessment_language) exam_assessment_language,MIN(ebu.course_start_date) course_start_date"
      Dim vNestedSQL As SQLStatement = Nothing
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading) Then
        Dim vNestedAnsiJoins As New AnsiJoins
        With vNestedAnsiJoins
          .Add("exam_booking_units ebu", "eb.exam_booking_id", "ebu.exam_booking_id")
        End With
        Dim vNestedWhereFields As New CDBFields(New CDBField("eb.contact_number", mvParameters("ContactNumber").IntegerValue))
        vNestedSQL = New SQLStatement(mvEnv.Connection, vNestedAttrs, "exam_bookings eb", vNestedWhereFields, "", vNestedAnsiJoins)
        vNestedSQL.GroupBy = "ebu.exam_booking_id"
      End If

      Dim vWhereFields As New CDBFields(New CDBField("eb.contact_number", mvParameters("ContactNumber").IntegerValue))
      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("exam_unit_links eul", "eb.exam_unit_link_id", "eul.exam_unit_link_id")
        .Add("exam_units eu", "eul.exam_unit_id_2", "eu.exam_unit_id")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading) = True AndAlso vNestedSQL IsNot Nothing Then
          .Add("(" & vNestedSQL.SQL & ") ebu", "eb.exam_booking_id", "ebu.exam_booking_id")
        End If
        .AddLeftOuterJoin("exam_sessions es", "eb.exam_session_id", "es.exam_session_id")
        .AddLeftOuterJoin("exam_centres ec", "eb.exam_centre_id", "ec.exam_centre_id")
        .AddLeftOuterJoin("exam_centre_units ecu", "ec.exam_centre_id", "ecu.exam_centre_id", "eul.exam_unit_link_id", "ecu.exam_unit_link_id") 'non-session based centre unit
        .AddLeftOuterJoin("exam_centre_units ecbu", "ec.exam_centre_id", "ecbu.exam_centre_id", "eul.base_unit_link_id", "ecbu.exam_unit_link_id") 'sesison based centre unit
        .AddLeftOuterJoin("cancellation_reasons cr", "eb.cancellation_reason", "cr.cancellation_reason")
        .AddLeftOuterJoin("sources s", "eb.cancellation_source", "s.source")
        .AddLeftOuterJoin("batch_transactions bt", "eb.batch_number", "bt.batch_number", "eb.transaction_number", "bt.transaction_number")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading) Then .AddLeftOuterJoin("exam_assessment_languages eal", "ebu.exam_assessment_language", "eal.exam_assessment_language")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamStudyModes) Then .AddLeftOuterJoin("study_modes sm", "eb.study_mode", "sm.study_mode")
      End With

      Dim vOrderByClause As String = String.Format("COALESCE({0}, {1}, {2}) DESC", "es.exam_session_year", mvEnv.Connection.DBYear("ebu.course_start_date"), mvEnv.Connection.DBYear("eb.created_on"))
      vOrderByClause += String.Format(", COALESCE({0}, {1}, {2}) DESC", "es.exam_session_month", mvEnv.Connection.DBMonth("ebu.course_start_date"), mvEnv.Connection.DBMonth("eb.created_on"))
      vOrderByClause += String.Format(", COALESCE(CASE WHEN es.exam_session_id IS NOT NULL THEN 1 ELSE NULL END, {0}, {1}) DESC", mvEnv.Connection.DBDay("ebu.course_start_date"), mvEnv.Connection.DBDay("eb.created_on"))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "exam_bookings eb", vWhereFields, vOrderByClause, vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("SpecialRequirements")
      Next
    End Sub

    Private Sub GetContactExamDetailItems(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "exam_booking_unit_id,ebu.exam_booking_id,ebu.exam_unit_id,eu.exam_unit_code,eu.exam_unit_description,exam_candidate_number,"
      vAttrs &= "attempt_number,ebu.exam_student_unit_status,exam_student_unit_status_desc,original_mark,moderated_mark,total_mark,original_grade,"
      vAttrs &= "moderated_grade,total_grade,eg.exam_grade_desc,original_result,moderated_result,total_result,{0},"
      vAttrs &= "done_date,start_date,start_time,end_time,bta.source,eu.activity_group,ebu.cancellation_reason,ebu.cancellation_source"
      vAttrs &= ",cancellation_reason_desc,source_desc,ebu.cancelled_on,ebu.cancelled_by,eul.exam_unit_link_id,eul.parent_unit_link_id,"
      vAttrs &= "ebu.created_by,ebu.created_on,ebu.amended_by,ebu.amended_on,eb.exam_centre_id,{1}"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamsQualsRegistrationGrading) Then
        vAttrs &= ",ebu.course_start_date,ebu.exam_assessment_language,eu.exam_session_id"
      Else
        vAttrs &= ",,,"
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAttrs &= ",eses.results_release_date,esuh.previous_mark,esuh.previous_grade,esuh.previous_result,{2}"
      Else
        vAttrs &= ",,,,"
      End If
      vAttrs &= ","
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamStudyModes) Then vAttrs &= "eb.study_mode"
      Dim vFields As String = String.Format(vAttrs, "total_result_desc", "can_edit_results", "previous_grade_desc,previous_result_desc")
      vAttrs = RemoveBlankItems(String.Format(vAttrs, "total_result AS total_result_desc", "'Y' AS can_edit_results", "'' AS previous_grade_desc,'' AS previous_result_desc"))

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromIntegerParameter(vWhereFields, "ContactNumber", "ebu.contact_number")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingId", "ebu.exam_booking_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingUnitId", "ebu.exam_booking_unit_id")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_bookings eb", "ebu.exam_booking_id", "eb.exam_booking_id")
      vAnsiJoins.Add("exam_unit_links eul", "ebu.exam_unit_link_id", "eul.exam_unit_link_id")
      vAnsiJoins.Add("exam_units eu", "eul.exam_unit_id_2", "eu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_student_unit_statuses esus", "ebu.exam_student_unit_status", "esus.exam_student_unit_status")
      vAnsiJoins.AddLeftOuterJoin("exam_schedule es", "ebu.exam_schedule_id", "es.exam_schedule_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "ebu.total_grade", "eg.exam_grade")
      vAnsiJoins.AddLeftOuterJoin("batch_transaction_analysis bta", "bta.batch_number", "ebu.batch_number", "bta.transaction_number", "ebu.transaction_number", "bta.line_number", "ebu.line_number")
      vAnsiJoins.AddLeftOuterJoin("exam_centre_units ecu", "ecu.exam_unit_link_id", "ebu.exam_unit_link_id", "ecu.exam_centre_id", "eb.exam_centre_id") 'non-session centre unit
      vAnsiJoins.AddLeftOuterJoin("exam_centre_units ecbu", "ecbu.exam_unit_link_id", "eul.base_unit_link_id", "ecbu.exam_centre_id", "eb.exam_centre_id") 'session centre unit
      vAnsiJoins.AddLeftOuterJoin("exam_sessions eses", "eb.exam_session_id", "eses.exam_session_id")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation) Then
        vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "ebu.cancellation_reason", "cr.cancellation_reason")
        vAnsiJoins.AddLeftOuterJoin("sources s", "ebu.cancellation_source", "s.source")
      Else
        vAttrs = vAttrs.Replace("ebu.cancellation_reason,ebu.cancellation_source,cancellation_reason_desc,source_desc,ebu.cancelled_on,ebu.cancelled_by,", ",,,,,,")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        Dim vESUHAttrs As String = "esuh.exam_student_header_id, esuh.exam_unit_link_id, esh.contact_number, current_mark, current_grade, current_result,"
        vESUHAttrs &= "results_release_date, '' AS previous_mark, '' AS previous_grade, '' AS previous_result"
        Dim vESUHAnsiJoins As New AnsiJoins()
        vESUHAnsiJoins.Add("exam_student_header esh", "esuh.exam_student_header_id", "esh.exam_student_header_id")
        Dim vESUHSqlStatment As New SQLStatement(mvEnv.Connection, vESUHAttrs, "exam_student_unit_header esuh", Nothing, "", vESUHAnsiJoins)
        vAnsiJoins.AddLeftOuterJoin("(" & vESUHSqlStatment.SQL & ") esuh", "eul.base_unit_link_id", "esuh.exam_unit_link_id", "ebu.contact_number", "esuh.contact_number")
        vAnsiJoins.AddLeftOuterJoin("exam_grades peg", "esuh.previous_grade", "peg.exam_grade")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        'Add Local Name field
        vAttrs &= ",COALESCE(ecu.local_name, ecbu.local_name) AS local_name"
        vFields &= ",local_name"
        'Add achieved units fields
        vAttrs &= ",achieved_qual.exam_unit_code AS achieved_unit_code, achieved_qual.exam_unit_description AS achieved_unit_description"
        vFields &= ",achieved_unit_code,achieved_unit_description"

        'Add the SQL to get the achieved unit.  Warning: this is horrible.  Expand below for an explanation (warning: this is boring)
        'Booking units are linked to their achieved units by assuming that the achieved unit will get the same grading run number as the booking.
        'When a booking is graded, every graded unit gets the same grading run number, whether it's a parent unit (provided it has grading rules) or a child unit.
        'However, when looking for the achieved unit, the user wants to see the highest achievement, and that's where the complication lies.  So we do the following:
        'first create 1 SQL statement that gets us the highest parent unit.  This will get us the parent of the highest unit, not the highest unit.
        'From the highest parent, we join back to the units that have the grading run number that we're looking for.
        'Of course we must also check that the highest parent has a pass grade too.
        'I hope this makes sense.  MRP 2014.01.14
        'Inner sub-select
        Dim vAchievedParentFields As String = "contact_number, grading_run_number, min(parent_unit_link_id) parent_unit_link_id"
        Dim vAchievedParentQual As New SQLStatement(mvEnv.Connection, vAchievedParentFields, "VExamSummaryAchievements achieved_qual_parent")
        vAchievedParentQual.GroupBy = "contact_number, grading_run_number"

        'Outer sub-select (I know, I know but this is horrible, see notes above)
        Dim vAchievedJoins As New AnsiJoins
        Dim vAchievedWhere As New CDBFields
        Dim vAchievedQualFields As String = "contact_number, grading_run_number, min(exam_unit_link_id) exam_unit_link_id"
        Dim vAchievedQual As New SQLStatement(mvEnv.Connection, vAchievedQualFields, "VExamSummaryAchievements achieved_qual")
        vAchievedQual.GroupBy = "contact_number, grading_run_number, parent_unit_link_id"
        vAchievedJoins.Add("(" + vAchievedQual.SQL + ") achieved_unit_header", "achieved_unit_header.grading_run_number", "achieved_qual_parent.grading_run_number", "achieved_unit_header.contact_number", "achieved_qual_parent.contact_number")
        vAchievedJoins.Add("exam_unit_links achieved_link", "achieved_link.exam_unit_link_id", "achieved_unit_header.exam_unit_link_id", "achieved_link.parent_unit_link_id", "achieved_qual_parent.parent_unit_link_id")
        vAchievedJoins.Add("exam_units achieved_unit", "achieved_unit.exam_unit_id", "achieved_link.exam_unit_id_2")

        Dim vAchievedField As String = "achieved_qual_parent.contact_number, achieved_qual_parent.grading_run_number, achieved_unit.exam_unit_code, achieved_unit.exam_unit_description"
        Dim vAchievedSQL As New SQLStatement(mvEnv.Connection, vAchievedField, "(" + vAchievedParentQual.SQL + ") achieved_qual_parent", vAchievedWhere, "", vAchievedJoins)

        vAnsiJoins.AddLeftOuterJoin("(" + vAchievedSQL.SQL + ") achieved_qual", "achieved_qual.contact_number", "eb.contact_number", "achieved_qual.grading_run_number", "ebu.grading_run_number")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "exam_booking_units ebu", vWhereFields, "exam_booking_unit_id", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then RestrictExamResults(pDataTable, "Total")
      GetLookupData(pDataTable, "TotalResult", "exam_booking_units", "original_result")
    End Sub

    Private Sub GetContactExamDetailList(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "exam_booking_unit_id,ebu.exam_booking_id,ebu.exam_unit_id,exam_unit_code,exam_unit_description,exam_candidate_number,attempt_number,ebu.exam_student_unit_status,exam_student_unit_status_desc,"
      vAttrs &= "original_mark,moderated_mark,total_mark,original_grade,moderated_grade,total_grade,eg.exam_grade_desc,original_result,moderated_result,total_result,{0},done_date,start_date,"
      vAttrs &= "start_time,end_time"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation) Then
        vAttrs &= ",ebu.cancellation_reason,ebu.cancellation_source,cancellation_reason_desc,source_desc,ebu.cancelled_on,ebu.cancelled_by"
      Else
        vAttrs &= ",,,,,,"
      End If
      vAttrs &= ",eul.exam_unit_link_id,eul.parent_unit_link_id,ebu.created_by,ebu.created_on,ebu.amended_by,ebu.amended_on,{1}"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAttrs &= ",eses.results_release_date,esuh.previous_mark,esuh.previous_grade,esuh.previous_result,{2}"
      Else
        vAttrs &= ",,,,"
      End If
      Dim vFields As String = String.Format(vAttrs, "total_result_desc", "can_edit_results", "previous_grade_desc,previous_result_desc")
      vAttrs = RemoveBlankItems(String.Format(vAttrs, "total_result AS total_result_desc", "'Y' AS can_edit_results", "'' AS previous_grade_desc,'' AS previous_result_desc"))

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingId", "ebu.exam_booking_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingUnitId", "exam_booking_unit_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamUnitId", "exam_unit_id_1")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamUnitLinkId", "eul.parent_unit_link_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ContactNumber", "ebu.contact_number")
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_units eu", "ebu.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_bookings eb", "ebu.exam_booking_id", "eb.exam_booking_id")
      vAnsiJoins.AddLeftOuterJoin("exam_student_unit_statuses esus", "ebu.exam_student_unit_status", "esus.exam_student_unit_status")
      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "ebu.exam_unit_link_id", "eul.exam_unit_link_id")
      vAnsiJoins.AddLeftOuterJoin("exam_schedule es", "ebu.exam_schedule_id", "es.exam_schedule_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "ebu.total_grade", "eg.exam_grade")
      vAnsiJoins.AddLeftOuterJoin("exam_sessions eses", "eb.exam_session_id", "eses.exam_session_id")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation) Then
        vAnsiJoins.AddLeftOuterJoin("cancellation_reasons cr", "ebu.cancellation_reason", "cr.cancellation_reason")
        vAnsiJoins.AddLeftOuterJoin("sources s", "ebu.cancellation_source", "s.source")
      End If
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        Dim vESUHAttrs As String = "esuh.exam_student_header_id, esuh.exam_unit_link_id, esh.contact_number, current_mark, current_grade, current_result,"
        vESUHAttrs &= "results_release_date, '' AS previous_mark, '' AS previous_grade, '' As previous_result"
        Dim vESUHAnsiJoins As New AnsiJoins()
        vESUHAnsiJoins.Add("exam_student_header esh", "esuh.exam_student_header_id", "esh.exam_student_header_id")
        Dim vESUHSqlStatment As New SQLStatement(mvEnv.Connection, vESUHAttrs, "exam_student_unit_header esuh", Nothing, "", vESUHAnsiJoins)
        vAnsiJoins.AddLeftOuterJoin("(" & vESUHSqlStatment.SQL & ") esuh", "eul.base_unit_link_id", "esuh.exam_unit_link_id", "ebu.contact_number", "esuh.contact_number")
        vAnsiJoins.AddLeftOuterJoin("exam_grades peg", "esuh.previous_grade", "peg.exam_grade")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "exam_booking_units ebu", vWhereFields, "exam_booking_unit_id", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
      RestrictExamResults(pDataTable, "Total")
      GetLookupData(pDataTable, "TotalResult", "exam_booking_units", "original_result")
    End Sub

    Private Sub GetContactExamExemptions(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "exam_student_exemption_id,ese.exam_exemption_id,exam_exemption_code,exam_exemption_description,ese.exam_exemption_status,exam_exemption_status_desc,allow_exemption_entry,exam_exemption_status_type,ese.status_date,ee.product,product_desc,ee.rate,rate_desc,batch_number,transaction_number,line_number,ese.organisation_number,name,exemption_module,ese.created_by,ese.created_on,ese.amended_by,ese.amended_on"
      Dim vWhereFields As New CDBFields
      If mvParameters.HasValue("ExamStudentExemptionId") Then vWhereFields.Add("exam_student_exemption_id", mvParameters("ExamStudentExemptionId").IntegerValue)
      If mvParameters.HasValue("ContactNumber") Then vWhereFields.Add("ese.contact_number", mvParameters("ContactNumber").IntegerValue)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("exam_exemptions ee", "ese.exam_exemption_id", "ee.exam_exemption_id")
      vAnsiJoins.Add("exam_exemption_statuses ees", "ese.exam_exemption_status", "ees.exam_exemption_status")
      vAnsiJoins.Add("products p", "ee.product", "p.product")
      vAnsiJoins.Add("rates r", "ee.product", "r.product", "ee.rate", "r.rate")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamExemptionModule) Then
        vAnsiJoins.AddLeftOuterJoin("organisations o", "ese.organisation_number", "o.organisation_number")
      Else
        vFields = vFields.Replace("ese.organisation_number,name,exemption_module", ",,")
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "exam_student_exemptions ese", vWhereFields, "", vAnsiJoins)
      vSQLStatement.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vFields)
    End Sub

    Private Sub GetDataUpdatesData(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "'N' AS check_value, data_update_number, brief_desc, detailed_desc, single_application, amended_by, amended_on"
      Dim vWhereFields As New CDBFields(New CDBField("data_update_number", 28, CDBField.FieldWhereOperators.fwoGreaterThanEqual))   'This is the first Data Update available in the Smart Client
      With vWhereFields
        .Add("single_application", "N", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("single_application#2", "Y", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual)
        .Add("applied_by", "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "data_updates du", vWhereFields, "data_update_number")
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      'Add any special restrictions below that mean an Update does not need to be available (none at the moment)
    End Sub

    Private Sub GetContactAddressesAndPositions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "ca.address_number,address_type,house_name,address,town,county,postcode,a.country,country_desc,branch,paf,sortcode,uk,mosaic_code,building_number,delivery_point_suffix,a.amended_by,a.amended_on,address_format,address_line1,address_line2,address_line3,o.organisation_number,o.name"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vAttrs = vAttrs.Replace(",building_number", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAddressDPS) Then vAttrs = vAttrs.Replace(",delivery_point_suffix", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCountryAddressFormat) Then vAttrs = vAttrs.Replace(",address_format", ",")

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("ca.contact_number", mvParameters("ContactNumber").IntegerValue)
      vWhereFields.Add("ca.historical", "N")
      vWhereFields.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), "Y", CDBField.FieldWhereOperators.fwoNullOrEqual)
      vWhereFields.Add("organisation_group", mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtOrganisation).EntityGroupCode, CDBField.FieldWhereOperators.fwoNullOrEqual)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("addresses a", "a.address_number", "ca.address_number")
      vAnsiJoins.Add("countries co", "co.country", "a.country")
      vAnsiJoins.AddLeftOuterJoin("contact_positions cp", "ca.address_number", "cp.address_number", "ca.contact_number", "cp.contact_number")
      vAnsiJoins.AddLeftOuterJoin("organisations o", "cp.organisation_number", "o.organisation_number")

      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_addresses ca", vWhereFields, "a.address_type, o.name", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "ADDRESS_LINE")
      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.Item("OrganisationNumber") = ""
        vRow.Item("OrganisationName") = ""
      Next
      vWhereFields.Add("a.address_type", "O")
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "ADDRESS_LINE")
    End Sub

    Private Sub GetDuplicateOrganisationsForRegistration(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "view_name"
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("view_type", "O")
      Dim vMainTable As String = New SQLStatement(mvEnv.Connection, vAttrs, "view_names", vWhereFields, "view_name").GetValue
      If vMainTable.Length = 0 Then vMainTable = "organisations"

      vAttrs = "oa.address_number,address_type,house_name,address,town,county,postcode,a.country,country_desc,branch,paf,sortcode,uk,mosaic_code,building_number,delivery_point_suffix,a.amended_by,a.amended_on,address_format,address_line1,address_line2,address_line3,o.organisation_number,o.name"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vAttrs = vAttrs.Replace(",building_number", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAddressDPS) Then vAttrs = vAttrs.Replace(",delivery_point_suffix", ",")
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCountryAddressFormat) Then vAttrs = vAttrs.Replace(",address_format", ",")

      vWhereFields.Clear()
      Dim vAnsiJoins As New AnsiJoins
      If mvEnv.GetConfigOption("cd_advanced_name_searching", False) AndAlso mvEnv.Connection.IsCaseSensitive Then
        vAnsiJoins.Add("contact_search_names csn", "o.organisation_number", "csn.contact_number")
        vWhereFields.Add("search_name", "*" & mvParameters("Name").Value.ToLower & "*", CDBField.FieldWhereOperators.fwoLike)
        vWhereFields.Add("csn.is_active", "Y")
      Else
        vWhereFields.Add("o.name", "*" & mvParameters("Name").Value & "*", CDBField.FieldWhereOperators.fwoLike Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("o.abbreviation", "*" & mvParameters("Name").Value & "*", CDBField.FieldWhereOperators.fwoLike Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      vWhereFields.Add("oa.historical", "N")

      vAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number")
      vAnsiJoins.Add("addresses a", "a.address_number", "oa.address_number")
      vAnsiJoins.Add("countries co", "co.country", "a.country")

      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vMainTable & " o", vWhereFields, "o.name", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vAttrs, "ADDRESS_LINE")
    End Sub

    Private Sub GetLoanInterestRates(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoanInterestRates) Then
        Dim vAttrs As String = "lir.loan_number,lir.interest_rate,rate_changed"
        Dim vWhereFields As New CDBFields()
        Dim vAnsiJoins As AnsiJoins = Nothing
        Dim vTableName As String
        If mvParameters.ParameterExists("PaymentPlanNumber").IntegerValue > 0 Then
          vTableName = "loans l"
          vAnsiJoins = New AnsiJoins
          vAnsiJoins.Add("loan_interest_rates lir", "l.loan_number", "lir.loan_number")
          vWhereFields.Add("l.order_number", mvParameters("PaymentPlanNumber").IntegerValue)
        Else
          vTableName = "loan_interest_rates lir"
          vWhereFields.Add("loan_number", mvParameters("LoanNumber").IntegerValue)
        End If

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTableName, vWhereFields, "rate_changed DESC", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement)
      End If
    End Sub

    Private Sub GetWebExams(ByVal pDataTable As CDBDataTable)
      'mvResultColumns = "ExamUnitId,ExamUnitDescription,ExamImage,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,StartDate,StartTime"

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins

      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamSessionId", "es.exam_session_id")
      AddWhereFieldFromParameter(vWhereFields, "ExamSessionCode", "es.exam_session_code")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamCentreId", "ec.exam_centre_id")
      AddWhereFieldFromParameter(vWhereFields, "ExamCentreCode", "ec.exam_centre_code")
      AddWhereFieldFromParameter(vWhereFields, "Subject", "eu.subject")
      AddWhereFieldFromParameter(vWhereFields, "SkillLevel", "eu.skill_level")

      If mvParameters.Exists("SearchExam") Then
        vWhereFields.Add("es.exam_session_description", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If Not mvParameters.ContainsKey("ExamSessionCode") Then vWhereFields.Add("es.exam_session_code", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("eu.exam_unit_description", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("eu.exam_unit_code", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        If Not mvParameters.ContainsKey("ExamCentreCode") Then vWhereFields.Add("ec.exam_centre_code", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR)
        vWhereFields.Add("ec.exam_centre_description", mvParameters("SearchExam").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If

      vAnsiJoins.Add("exam_schedule esc", "eu.exam_unit_id", "esc.exam_unit_id")
      vAnsiJoins.Add("exam_sessions es", "esc.exam_session_id", "es.exam_session_id")
      vAnsiJoins.Add("exam_centres ec", "esc.exam_centre_id", "ec.exam_centre_id")
      vAnsiJoins.AddLeftOuterJoin("subjects s", "eu.subject", "s.subject")
      vAnsiJoins.AddLeftOuterJoin("skill_levels sl", "eu.skill_level", "sl.skill_level")

      vWhereFields.Add("allow_bookings", "Y")
      vWhereFields.Add("es.web_publish", "Y")
      vWhereFields.Add("ec.web_publish", "Y")
      vWhereFields.Add("eu.web_publish", "Y")
      vWhereFields.Add("esc.start_date", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan)

      ' Do not return schedules where the corresponding Exam Session Home/Overseas (dependent on Centre location) date has exipred
      vWhereFields.Add("ec.overseas", "Y", CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("es.overseas_closing_date", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("ec.overseas ", "N", CDBField.FieldWhereOperators.fwoNullOrEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("es.home_closing_date", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)

      mvResultColumns = "ExamUnitId,ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitCode,ExamUnitDescription,ExamSessionCode,ExamSessionDescription,ExamCentreCode,ExamCentreDescription,ExamImage,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,StartDate,StartTime"

      Dim vFields As String = "eu.exam_unit_id,esc.exam_schedule_id,es.exam_session_id,ec.exam_centre_id,exam_unit_code,exam_unit_description,exam_session_code,exam_session_description,exam_centre_code,exam_centre_description,Null as exam_image,eu.subject,subject_desc,eu.skill_level,skill_level_desc,start_date,start_time"


      'If (AddExtras) Then
      '  Dim vBookingWhereFields As New CDBFields
      '  Dim vBookingAnsiJoins As New AnsiJoins

      '  vBookingAnsiJoins.AddLeftOuterJoin("exam_booking_units ebu", "eb.exam_booking_id", "ebu.exam_booking_id")
      '  vBookingAnsiJoins.AddLeftOuterJoin("exam_units eu", "ebu.exam_unit_id", "eu.exam_unit_id")

      '  If mvParameters.HasValue("ExamSessionId") Then
      '    vBookingWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      '  ElseIf mvParameters.HasValue("ExamSessionCode") Then
      '    vBookingWhereFields.Add("es.exam_session_code", mvParameters("ExamSessionCode").Value)
      '    vBookingAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "eu.exam_session_id")
      '  Else
      '    vBookingWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
      '  End If

      '  ' if contact is supplied, limit to this contact
      '  If mvParameters.HasValue("ContactNumber") Then
      '    vBookingWhereFields.Add("eb.contact_number", mvParameters("ContactNumber").IntegerValue)
      '  Else
      '    vExtraSort = "contact_number,"
      '  End If


      '  Dim vInnerFields As String = "eb.contact_number,ebu.exam_unit_id,eb.exam_booking_id,ebu.exam_booking_unit_id, ebu.total_mark, ebu.total_grade, ebu.total_result, ebu.exam_student_unit_status, ebu.done_date, eb.exam_session_id, eb.special_requirements"
      '  Dim vBookingSQL As New SQLStatement(mvEnv.Connection, vInnerFields, "exam_bookings eb", vBookingWhereFields, "", vBookingAnsiJoins)
      '  vBookingSQL.GroupBy = vInnerFields ' added as multi summary header lines caused cartesian join.

      '  vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) eb", vBookingSQL.SQL), "eb.exam_unit_id", "eu.exam_unit_id", "eb.exam_session_id", "eu.exam_session_id")
      '  vFields = vFields.Replace("0,0,0,0,'','','','',''", "exam_booking_unit_id,eb.exam_booking_id,eb.total_mark,total_grade,total_result,eb.exam_student_unit_status," + mvEnv.Connection.DBIsNull("eb.contact_number", "0") + " contact_number,eb.done_date,eb.special_requirements")

      If mvParameters.HasValue("ContactNumber") Then
        Dim vPassesWhereFields As New CDBFields
        vPassesWhereFields.Add("eg.grade_is_pass", "Y")
        vPassesWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
        Dim vPassesAnsiJoins As New AnsiJoins
        vPassesAnsiJoins.Add("exam_student_unit_header esuh", "esh.exam_student_header_id", "esuh.exam_student_header_id")
        vPassesAnsiJoins.Add("exam_grades eg", "esuh.current_grade", "eg.exam_grade")
        Dim vPassesSQL As New SQLStatement(mvEnv.Connection, "esuh.exam_unit_id, grade_is_pass", "exam_student_header esh ", vPassesWhereFields, "", vPassesAnsiJoins)
        vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) esg", vPassesSQL.SQL), "esg.exam_unit_id", "eu.exam_base_unit_id")
        vWhereFields.Add("grade_is_pass")
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_units eu", vWhereFields, "start_date, exam_unit_description, exam_centre_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
      Dim vConfigValue As String = mvEnv.GetConfig("web_exam_image_name")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vConfigValue.Length > 0 Then
          vRow.Item("ExamImage") = String.Format(vConfigValue, vRow.Item("ExamUnitCode"))
        Else
          vRow.Item("ExamImage") = String.Format("Exam{0}.png", vRow.Item("ExamUnitCode"))
        End If
      Next
    End Sub

    Private Sub GetWebExamBookings(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      ' Mandatory parameters
      vWhereFields.Add("eb.contact_number", mvParameters("ContactNumber").IntegerValue)

      ' Optional parameters
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingId", "eb.exam_booking_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamBookingUnitId", "ebu.exam_booking_unit_id")
      AddWhereFieldFromIntegerParameter(vWhereFields, "ExamUnitId", "eu.exam_unit_id")

      vAnsiJoins.Add("exam_schedule esc", "ebu.exam_schedule_id", "esc.exam_schedule_id")
      'vAnsiJoins.Add("exam_sessions es", "esc.exam_session_id", "es.exam_session_id")
      vAnsiJoins.Add("exam_sessions es", "esc.exam_session_id", "es.exam_session_id")
      vAnsiJoins.Add("exam_centres ec", "esc.exam_centre_id", "ec.exam_centre_id")
      vAnsiJoins.Add("exam_units eu", "ebu.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")
      vAnsiJoins.AddLeftOuterJoin("exam_student_unit_statuses esus", "ebu.exam_student_unit_status", "esus.exam_student_unit_status")

      vWhereFields.Add("es.web_publish", "Y")
      vWhereFields.Add("eu.web_publish", "Y")
      vWhereFields.Add("ec.web_publish", "Y")
      vWhereFields.Add("esc.start_date + 1", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan)
      vWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
      vWhereFields.Add("ebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings

      Dim vFields As String = "ebu.exam_booking_unit_id,eu.exam_unit_id,esc.exam_schedule_id,es.exam_session_id,ec.exam_centre_id,eu.exam_unit_code,eu.exam_unit_description,es.exam_session_code,es.exam_session_description,ec.exam_centre_code,ec.exam_centre_description,esc.start_date,esc.start_time,esc.end_time,ebu.exam_student_unit_status,esus.exam_student_unit_status_desc"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_booking_units ebu", vWhereFields, "esc.start_date, esc.start_time, eu.exam_unit_description, ec.exam_centre_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetWebExamHistory(ByVal pDataTable As CDBDataTable)
      ' Mandatory parameters - ContactNumber

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vFields As String = "esh.exam_student_header_id,esuh.exam_student_unit_header_id,eu.exam_unit_id,es.exam_session_id,es.exam_session_code,es.exam_session_description,es.exam_session_month,es.exam_session_year,esc.start_date,eu.exam_unit_code,eu.exam_unit_description,ebu.total_mark AS current_mark,ebu.total_grade AS current_grade,eg.exam_grade_desc,ebu.total_result AS Current_result,ebu.total_result AS CurrentResultDesc"

      ' Get all session based exams 
      vAnsiJoins.Add("exam_units eu", "esuh.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_student_header esh", "esh.exam_student_header_id", "esuh.exam_student_header_id")
      vAnsiJoins.Add("exam_units bookeu", "bookeu.exam_base_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_booking_units ebu", "ebu.exam_unit_id", "bookeu.exam_unit_id")
      vAnsiJoins.Add("exam_grades eg", "eg.exam_grade", "ebu.total_grade")
      vAnsiJoins.Add("exam_bookings eb", "ebu.exam_booking_id", "eb.exam_booking_id", "eb.contact_number", "esh.contact_number")
      vAnsiJoins.Add("exam_schedule esc", "esc.exam_schedule_id", "ebu.exam_schedule_id")
      vAnsiJoins.Add("exam_sessions es", "esc.exam_session_id", "es.exam_session_id")
      vWhereFields.Add("eb.cancelled_by")
      vWhereFields.Add("ebu.cancelled_by")
      vWhereFields.Add("esh.contact_number", mvParameters("ContactNumber").IntegerValue)
      vWhereFields.Add("es.web_publish", "Y")
      vWhereFields.Add("eu.web_publish", "Y")

      Dim vSubWhereFields As New CDBFields
      Dim vSubAnsiJoins As New AnsiJoins
      Dim vSubFields As String = "max(subebu.exam_booking_unit_id)"
      vSubAnsiJoins.Add("exam_bookings subeb", "subeb.exam_booking_id", "subebu.exam_booking_id", "subeb.contact_number", "eb.contact_number")
      vSubAnsiJoins.Add("exam_units subeu", "subebu.exam_unit_id", "subeu.exam_unit_id", "subeu.exam_base_unit_id", "eu.exam_unit_id")
      vSubAnsiJoins.Add("exam_schedule subesc", "subesc.exam_schedule_id", "subebu.exam_schedule_id")
      vSubAnsiJoins.Add("exam_sessions subes", "subesc.exam_session_id", "subes.exam_session_id")
      vSubWhereFields.Add("subes.web_publish", "Y")
      vSubWhereFields.Add("subeu.web_publish", "Y")
      Dim vSubSQL As New SQLStatement(mvEnv.Connection, vSubFields, "exam_booking_units subebu", vSubWhereFields, "", vSubAnsiJoins)
      vWhereFields.Add("ebu.exam_booking_unit_id", CDBField.FieldTypes.cftLong, "(" + vSubSQL.SQL + ")")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_student_unit_header esuh", vWhereFields, "eu.exam_unit_description, exam_session_month,exam_session_year", vAnsiJoins)

      ' Get all non-session based exams
      Dim vNonSessWhereFields As New CDBFields
      Dim vNonSessAnsiJoins As New AnsiJoins
      Dim vNonSessFields As String = "esh.exam_student_header_id,esuh.exam_student_unit_header_id,eu.exam_unit_id,null exam_session_id,'' exam_session_code,'' exam_session_description,'' exam_session_month,'' exam_session_year,null AS start_date,eu.exam_unit_code,eu.exam_unit_description,esuh.current_mark,esuh.current_grade,eg.exam_grade_desc,esuh.current_result,esuh.current_result AS CurrentResultDesc"
      vNonSessAnsiJoins.Add("exam_units eu", "esuh.exam_unit_id", "eu.exam_unit_id")
      vNonSessAnsiJoins.Add("exam_student_header esh", "esh.exam_student_header_id", "esuh.exam_student_header_id")
      vNonSessAnsiJoins.Add("exam_grades eg", "eg.exam_grade", "esuh.current_grade")
      vNonSessWhereFields.Add("eu.web_publish", "Y")
      vNonSessWhereFields.Add("esh.contact_number", mvParameters("ContactNumber").IntegerValue)
      Dim vNonSessSubAnsiJoins As New AnsiJoins
      Dim vNonSessSubFields As String = "subebu.exam_booking_id"
      vNonSessSubAnsiJoins.Add("exam_bookings subeb", "subeb.exam_booking_id", "subebu.exam_booking_id", "subeb.contact_number", "esh.contact_number")
      vNonSessSubAnsiJoins.Add("exam_units subeu", "subebu.exam_unit_id", "subebu.exam_unit_id", "subeu.exam_base_unit_id", "eu.exam_unit_id")
      Dim vNonSessSubSQL As New SQLStatement(mvEnv.Connection, vNonSessSubFields, "exam_booking_units subebu", Nothing, "", vNonSessSubAnsiJoins)
      vNonSessWhereFields.Add("Exclude", vNonSessSubSQL.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
      Dim vNonSessSQL As New SQLStatement(mvEnv.Connection, vNonSessFields, "exam_student_unit_header esuh", vNonSessWhereFields, "", vNonSessAnsiJoins)

      vSQL.AddUnion(vNonSessSQL)
      pDataTable.FillFromSQL(mvEnv, vSQL)
      GetLookupData(pDataTable, "CurrentResult", "exam_booking_units", "original_result") ' fill CurrentResultDesc with result description of current_result
    End Sub

    Private Sub GetContactAmendments(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAmendmentContactNumber) Then
        ' Mandatory parameters - ContactNumber
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("contact_number", mvContact.ContactNumber)
        Dim vSQL As New SQLStatement(mvEnv.Connection, "operation_date,logname,operation,table_name,contact_journal_number", "amendment_history", vWhereFields, "operation_date DESC")
        pDataTable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub

    Private Sub GetContactAmendmentDetails(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAmendmentContactNumber) Then
        ' Mandatory parameters - ContactNumber
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("contact_number", mvContact.ContactNumber)
        vWhereFields.Add("operation_date", CDBField.FieldTypes.cftTime, mvParameters("OperationDate").Value)
        vWhereFields.Add("operation", mvParameters("Operation").Value)
        vWhereFields.Add("table_name", mvParameters("TableName").Value)
        AddWhereFieldFromIntegerParameter(vWhereFields, "JournalNumber", "contact_journal_number")
        Dim vSQL As New SQLStatement(mvEnv.Connection, "data_values", "amendment_history", vWhereFields, "operation_date DESC")
        Dim vRS As CDBRecordSet = vSQL.GetRecordSet
        If vRS.Fetch = True Then
          Dim vItems() As String
          vItems = Split(vRS.Fields(1).MultiLine, Chr(22))
          Dim vRow As CDBDataRow = pDataTable.AddRow
          Dim vRowNumber As Integer = 0
          Dim vColumn As String = ""        'Add fix for compiler warning
          For vIndex As Integer = 0 To vItems.Length - 1
            If vItems(vIndex).Trim = "OLD" Then
              vColumn = "OldValues"
            ElseIf vItems(vIndex).Trim = "NEW" Then
              vColumn = "NewValues"
              vRowNumber = 0
            ElseIf Mid(vItems(vIndex), 3) = "NEW" Then
              vColumn = "NewValues"
              vRowNumber = 0
            Else
              Dim vValues() As String = Split(vItems(vIndex), ":")
              If vValues.Length > 1 Then
                If vRowNumber > pDataTable.Rows.Count - 1 Then
                  vRow = pDataTable.AddRow
                Else
                  vRow = pDataTable.Rows.Item(vRowNumber)
                End If
                vRow.Item("Item") = StrConv(Replace(vValues(0), "_", " "), VbStrConv.ProperCase)
                vRow.Item(vColumn) = Replace(vValues(1), vbCrLf, ", ")
                vRowNumber = vRowNumber + 1
              End If
            End If
          Next
        End If
      End If
    End Sub

    Private Sub GetMeetingActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "m.master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,a.action_status AS sort_column,,,,,,,,,,,a.duration_days,a.duration_hours,a.duration_minutes,a.document_class,action_text"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("m.meeting_number", mvParameters("MeetingNumber").IntegerValue)
      If mvParameters.HasValue("ActionNumber") Then vWhereFields.Add("a.action_number", mvParameters("ActionNumber").IntegerValue)
      vWhereFields.Add("a.created_by", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      vWhereFields.Add("creator_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#2", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("u.department", mvEnv.User.Department)
      vWhereFields.Add("department_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("a.created_by#3", mvEnv.User.Logname, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("u.department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("public_header", "Y", CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("actions a", "m.master_action", "a.master_action")
      vAnsiJoins.Add("users u", "a.created_by", "u.logname")
      vAnsiJoins.Add("document_classes dc", "a.document_class", "dc.document_class")
      vAnsiJoins.Add("action_priorities ap", "a.action_priority", "ap.action_priority")
      vAnsiJoins.Add("action_statuses acs", "a.action_status", "acs.action_status")

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "meetings m", vWhereFields, "sequence_number, action_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs.Replace("a.action_status AS sort_column", "action_status"))
      'For Each vRow As CDBDataRow In pDataTable.Rows
      '  'Order by Status (Overdue, Defined, Scheduled)
      '  If vRow.Item("SortColumn") = Action.GetActionStatusCode(astScheduled) Then vRow.Item("SortColumn") = ""
      'Next
      'pDataTable.ReOrderRowsByColumn("SortColumn", True)
      'GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
      GetActionersAndSubjects(pDataTable)
    End Sub


    Private Sub GetWebMemberOrganisations(pDataTable As CDBDataTable)
      'MemberNumber,Name,EmailAddress,Postcode

      Dim vAddress As New Address(mvEnv)
      Dim vFields As String = "o.organisation_number,name,abbreviation,member_number," & vAddress.GetRecordSetFieldsCountry

      Dim vWhereFields As New CDBFields
      AddWhereFieldFromParameter(vWhereFields, "MemberNumber", "member_number")
      AddWhereFieldFromParameter(vWhereFields, "Name", "name")
      AddWhereFieldFromParameter(vWhereFields, "Postcode", "postcode")
      vWhereFields.Add("cancellation_reason")

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("members m", "m.contact_number", "o.organisation_number")
      vAnsiJoins.Add("addresses a", "m.address_number", "a.address_number")
      vAnsiJoins.Add("countries co", "a.country", "co.country")
      If mvParameters.HasValue("EmailAddress") Then
        Dim vSubAnsiJoins As New AnsiJoins
        Dim vSubWhereFields As New CDBFields
        vSubAnsiJoins.Add("organisation_addresses oa", "o.organisation_number", "oa.organisation_number")
        vSubAnsiJoins.Add("communications cm", "oa.address_number", "cm.address_number")
        vSubWhereFields.Add(mvEnv.Connection.DBSpecialCol("cm", "number"), CDBField.FieldTypes.cftCharacter, mvParameters("EmailAddress").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        vSubWhereFields.Add("cm.valid_from", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
        vSubWhereFields.Add("cm.valid_from#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        vSubWhereFields.Add("cm.valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
        vSubWhereFields.Add("cm.valid_to#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        Dim vSubSQL As New SQLStatement(mvEnv.Connection, "o.organisation_number", "organisations o", vSubWhereFields, "", vSubAnsiJoins)
        vSubSQL.Distinct = True
        vAnsiJoins.Add(String.Format("({0}) cm", vSubSQL.SQL), "o.organisation_number", "cm.organisation_number")
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "organisations o", vWhereFields, "o.name", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "organisation_number,address_number,name,abbreviation,member_number,house_name,address,town,county,postcode,country,countrydesc,ADDRESS_LINE")
    End Sub

    Private Sub GetPaymentPlanPaymentDetails(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
        Dim vPPHDetails As New PaymentPlanHistoryDetail(mvEnv)
        Dim vAttrs As String = vPPHDetails.GetRecordSetFields.Replace(",pphd.amended_by,pphd.amended_on", "") & ",product_desc,rate_desc,vat_exclusive,activity_desc,activity_value_desc"

        Dim vAnsiJoins As New AnsiJoins
        With vAnsiJoins
          .Add("products p", "pphd.product", "p.product")
          .Add("rates r", "pphd.product", "r.product", "pphd.rate", "r.rate")
          .AddLeftOuterJoin("activities a", "pphd.modifier_activity", "a.activity")
          .AddLeftOuterJoin("activity_values av", "pphd.modifier_activity", "av.activity", "pphd.modifier_activity_value", "av.activity_value")
        End With

        Dim vWhereFields As New CDBFields(New CDBField("order_number", mvParameters("PaymentPlanNumber").IntegerValue))
        vWhereFields.Add("payment_number", mvParameters("PaymentNumber").IntegerValue)

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "payment_plan_history_details pphd", vWhereFields, "detail_number", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQLStatement, vAttrs, "," & ContactNameItems())

        'If multiple RateModifiers were used then set the data to "Multiple"
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("ModifierActivity").ToUpper = "MULTI" Then
            vRow.Item("ModifierActivity") = "Multiple"
            vRow.Item("ModifierActivityValue") = "Multiple"
            vRow.Item("ModifierActivityQuantity") = ""
            vRow.Item("ModifierActivityDate") = ""
            vRow.Item("ModifierPrice") = ""
            vRow.Item("ModifierPerItem") = "Multiple"
            vRow.Item("ModifierActivityDesc") = "Multiple"
            vRow.Item("ModifierActivityValueDesc") = "Multiple"
          End If
        Next
      End If

    End Sub

    Private Sub GetContactExamCertificates(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "contact_exam_cert_id, cec.contact_number,exam_cert_number_prefix,exam_cert_number,exam_cert_number_suffix,exam_unit_code,exam_unit_description,is_certificate_recalled"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cec.contact_number", mvParameters("ContactNumber").IntegerValue)
      If mvParameters.Exists("ContactExamCertId") Then
        vWhereFields.Add("cec.contact_exam_cert_id", mvParameters("ContactExamCertId").IntegerValue)
      End If

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("exam_student_unit_header esuh", "cec.exam_student_unit_header_id", "esuh.exam_student_unit_header_id")
      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "esuh.exam_unit_id")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_exam_certs cec", vWhereFields, "cec.contact_exam_cert_id", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSqlStatement)
    End Sub

    Private Sub GetContactExamCertificateItems(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cec_items.contact_exam_cert_item_id,ce_certs.contact_number,cec_items.contact_exam_cert_id,cec_items.exam_cert_attribute,cec_items.exam_cert_attribute_value,cec_items.amended_on,cec_items.amended_by"
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("cec_items.contact_exam_cert_id", mvParameters("ContactExamCertId").IntegerValue)

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contact_exam_certs ce_certs", "ce_certs.contact_exam_cert_id", "cec_items.contact_exam_cert_id")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), "contact_exam_cert_items cec_items", vWhereFields, "cec_items.contact_exam_cert_item_id", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSqlStatement)
    End Sub

    Private Sub GetContactExamCertificateReprints(ByVal pDataTable As CDBDataTable)
      Dim vSql As New SQLStatement(mvEnv.Connection,
                                     "cxcr.contact_exam_cert_reprint_id,cxcr.contact_exam_cert_id,cxcr.exam_cert_reprint_type,cxcr.amended_by,cxcr.amended_on",
                                     "contact_exam_cert_reprints cxcr",
                                     New CDBFields({New CDBField("contact_exam_cert_id", mvParameters("ContactExamCertId").Value)}))
      pDataTable.FillFromSQL(mvEnv, vSql)
    End Sub

    Private Function GetEventSessionList(pEventNumber As Integer) As String
      Dim vSessionNumberList As String = String.Empty
      Dim vWhereClause As New CDBFields
      vWhereClause.Add("event_number", pEventNumber)

      Dim vSQL As New SQLStatement(mvEnv.Connection, "session_number", "sessions s", vWhereClause)
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQL(mvEnv, vSQL)
      If vDataTable.Rows.Count > 0 Then vSessionNumberList = vDataTable.RowsAsCommaSeperated(vDataTable, "session_number")
      Return vSessionNumberList
    End Function

    Private Function GetBaseSessionNumber(pEventNumber As Integer) As Integer
      Dim vBaseSessionNumber As Integer
      Dim vEvent As New CDBEvent(mvEnv)
      vEvent.Init(pEventNumber)
      If vEvent.Existing Then vBaseSessionNumber = vEvent.LowestSessionNumber
      Return vBaseSessionNumber
    End Function

    Friend Sub GetActions(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.amended_by,a.amended_on,a.created_by,a.created_on,deadline,scheduled_on,"
      vAttrs &= "completed_on,a.action_priority,a.action_status,alk.type,{0},{1},{2}duration_days,duration_hours,duration_minutes,a.document_class,action_text,outlook_id"

      Dim vCols As String = String.Format(vAttrs, "alk.link_type", "a.action_status", ",,,,,,,,,,")   'Blank columns are ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,
      vAttrs = String.Format(vAttrs, "alk.type AS link_type_description", "a.action_status AS sort_column", "")

      'If type is dstContactPositionActions then we don't display Contact columns so remove them
      'If they get added back in then will also need to call GetActionersAndSubjects (below) to populate them
      If mvType = DataSelectionTypes.dstContactPositionActions Then vCols = RemoveBlankItems(vCols)

      Dim vTableName As String = "contact_actions alk"
      Dim vWhereFields As New CDBFields()
      If mvParameters.Exists("ActionNumber") Then vWhereFields.Add("alk.action_number", mvParameters("ActionNumber").IntegerValue)

      If mvType = DataSelectionTypes.dstContactNotifications Then
        vTableName = "contact_actions alk"
        vWhereFields.Add("alk.contact_number", mvEnv.User.ContactNumber)
        If mvParameters.ParameterExists("DeadlinesOnly").Bool = True Then
          vWhereFields.Add("alk.type", "A")
          vWhereFields.Add("deadline", CDBField.FieldTypes.cftTime, TodaysDateAndTime, CDBField.FieldWhereOperators.fwoLessThanEqual)
        Else
          vWhereFields.Add("alk.type", CDBField.FieldTypes.cftCharacter, "'A', 'M'", CDBField.FieldWhereOperators.fwoIn)
          vWhereFields.Add("alk.notified", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("alk.notified#2", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("alk.notified#3", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
        End If
      ElseIf mvType = DataSelectionTypes.dstContactPositionActions And mvParameters.ContainsKey("ContactPositionNumber") Then
        vTableName = "action_links alk"
        vWhereFields.Add("alk.contact_position_number", mvParameters("ContactPositionNumber").IntegerValue)
      Else
        If mvContact IsNot Nothing AndAlso mvContact.Existing Then
          If mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            vTableName = "organisation_actions alk"
            vWhereFields.Add("alk.organisation_number", mvContact.ContactNumber)
          Else
            vTableName = "contact_actions alk"
            vWhereFields.Add("alk.contact_number", mvContact.ContactNumber)
          End If
        ElseIf mvParameters IsNot Nothing Then
          If mvParameters.ContainsKey("WorkstreamId") Then
            vTableName = "action_links alk"
            vWhereFields.Add("alk.workstream_id", mvParameters("WorkstreamId").IntegerValue)
          End If
        End If
      End If

      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("actions a", "alk.action_number", "a.action_number")
        .Add("users u", "a.created_by", "u.logname")
        .Add("document_classes dc", "a.document_class", "dc.document_class")
        .Add("action_priorities ap", "a.action_priority", "ap.action_priority")
        .Add("action_statuses acs", "a.action_status", "acs.action_status")
      End With

      'Add Document Class joins
      'AND ((a.created_by = '" & mvEnv.User.Logname & "' AND creator_header = 'Y')" & " OR (a.created_by <> '" & mvEnv.User.Logname & "' AND department = '" & mvEnv.User.Department & "' AND department_header = 'Y')" & " OR (a.created_by <> '" & mvEnv.User.Logname & "' AND department <> '" & mvEnv.User.Department & "' AND public_header = 'Y' ))" 
      With vWhereFields
        .Add("a.created_by", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
        .Add("dc.creator_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("a.created_by#2", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("u.department", mvEnv.User.Department)
        .Add("dc.department_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("a.created_by#3", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("u.department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
        .Add("dc.public_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With

      'Add other criteria
      If mvParameters.ParameterExists("IgnoreStatus").Bool = False Then
        vWhereFields.Add("a.action_status", "'" & Action.GetActionStatusCode(Action.ActionStatuses.astDefined) & "','" & Action.GetActionStatusCode(Action.ActionStatuses.astScheduled) & "','" & Action.GetActionStatusCode(Action.ActionStatuses.astOverdue) & "'", CDBField.FieldWhereOperators.fwoIn)
      End If

      Dim vOrderBy As String = String.Empty
      If mvType = DataSelectionTypes.dstContactNotifications Then
        vCols = "a.action_number,alk.type,a.created_on,action_desc,action_priority_desc,,,,"  'ItemNumber,LinkType,ItemDate,ItemDescription,ItemType,Subject,ItemCode,ItemDesc,Access
      Else
        vOrderBy = mvEnv.GetConfig("actions_order", "").Trim
      End If
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTableName, vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)

      If mvType <> DataSelectionTypes.dstContactNotifications Then
        If vOrderBy.Length = 0 Then
          For Each vRow As CDBDataRow In pDataTable.Rows
            'Order by Status (Overdue, Defined, Scheduled)
            If vRow.Item("SortColumn") = Action.GetActionStatusCode(Action.ActionStatuses.astScheduled) Then vRow.Item("SortColumn") = ""
          Next
          pDataTable.ReOrderRowsByColumn("SortColumn", True)
        End If

        GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
        If mvType <> DataSelectionTypes.dstContactPositionActions Then GetActionersAndSubjects(pDataTable)
      End If

    End Sub

    Private Sub GetExamSchedules(pDataTable As CDBDataTable)
      Dim vFields As String = {"exam_schedule.exam_schedule_id", "exam_sessions.exam_session_id", "exam_centres.exam_centre_id",
                                 "exam_sessions.exam_session_code", "exam_sessions.exam_session_description", "exam_sessions.exam_session_year", "exam_sessions.exam_session_month", "exam_sessions.sequence_number",
                                 "exam_centres.exam_centre_code", "exam_centres.exam_centre_description",
                                 "exam_units.exam_unit_id", "exam_units.exam_unit_code", "exam_units.exam_unit_description",
                                 "exam_schedule.start_date", "exam_schedule.start_time", "exam_schedule.end_time"
                                }.AsCommaSeperated

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("exam_units", "exam_units.exam_unit_id", "exam_schedule.exam_unit_id")
      vAnsiJoins.Add("exam_sessions", "exam_sessions.exam_session_id", "exam_units.exam_session_id")
      vAnsiJoins.Add("exam_centres", "exam_centres.exam_centre_id", "exam_schedule.exam_centre_id")

      Dim vWhereFields As New CDBFields()
      If mvParameters.ContainsKey("ExamSessionCode") Then
        vWhereFields.Add("exam_sessions.exam_session_code", mvParameters("ExamSessionCode").Value)
      End If
      If mvParameters.ContainsKey("ExamCentreCode") Then
        vWhereFields.Add("exam_centres.exam_centre_code", mvParameters("ExamCentreCode").Value)
      End If
      If mvParameters.ContainsKey("ExamUnitCode") Then
        vWhereFields.Add("exam_units.exam_unit_code", mvParameters("ExamUnitCode").Value)
      End If
      If mvParameters.ContainsKey("FromDate") Then
        vWhereFields.Add("exam_schedule.start_date#1", CDBField.FieldTypes.cftDate, mvParameters("FromDate").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      If mvParameters.ContainsKey("ToDate") Then
        vWhereFields.Add("exam_schedule.start_date#2", CDBField.FieldTypes.cftDate, mvParameters("ToDate").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If
      If mvParameters.ContainsKey("WorkstreamId") Then
        Dim vSubWhere As New CDBFields()
        vSubWhere.Add("workstream_links.exam_schedule_id", "", CDBField.FieldWhereOperators.fwoNotEqual)
        vSubWhere.Add("workstream_links.exam_schedule_id#2", CDBField.FieldTypes.cftInteger, "exam_schedule.exam_schedule_id")
        vSubWhere.Add("workstream_links.workstream_id", CDBField.FieldTypes.cftInteger, mvParameters("WorkstreamId").IntegerValue)
        Dim vSubSql As New SQLStatement(mvEnv.Connection, "exam_schedule_id", "workstream_links", vSubWhere)
        vWhereFields.Add("Exclude", vSubSql.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_schedule", vWhereFields, "exam_schedule.start_date desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ",,")
    End Sub

    Private Sub GetEventSessionCPD(ByRef pDataTable As CDBDataTable)
      Dim vAttrs As String = "sc.event_session_cpd_number, sc.event_number, sc.session_number, sc.cpd_category_type, cct.cpd_category_type_desc, sc.cpd_category, cc.cpd_category_desc,"
      vAttrs &= " sc.cpd_year, sc.cpd_points, sc.cpd_points_2, sc.cpd_item_type, cit.cpd_item_type_desc, sc.cpd_outcome, sc.cpd_approval_status, cas.cpd_approval_status_desc, sc.cpd_date_approved,"
      vAttrs &= " sc.cpd_awarding_body, sc.web_publish, sc.cpd_notes"

      Dim vColumns As String = vAttrs
      If mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
        vAttrs = vAttrs.Replace("sc.cpd_points,", "CAST(sc.cpd_points AS int) AS cpd_points,")
        vAttrs = vAttrs.Replace("sc.cpd_points_2,", "CAST(sc.cpd_points_2 AS int) AS cpd_points_2,")
      End If

      Dim vOrderBy As String = "sc.session_number, cct.cpd_category_type_desc, cc.cpd_category_desc, sc.event_session_cpd_number"

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("sessions s", "sc.session_number", "s.session_number")
      vAnsiJoins.Add("cpd_category_types cct", "sc.cpd_category_type", "cct.cpd_category_type")
      vAnsiJoins.Add("cpd_categories cc", "sc.cpd_category", "cc.cpd_category", "sc.cpd_category_type", "cc.cpd_category_type")
      vAnsiJoins.Add("cpd_approval_statuses cas", "sc.cpd_approval_status", "cas.cpd_approval_status")
      vAnsiJoins.AddLeftOuterJoin("cpd_item_types cit", "sc.cpd_item_type", "cit.cpd_item_type")

      Dim vWhereFields As New CDBFields(New CDBField("sc.event_number", mvParameters("EventNumber").IntegerValue))
      If mvParameters.ContainsKey("SessionNumber") Then vWhereFields.Add("sc.session_number", mvParameters("SessionNumber").IntegerValue)
      If mvParameters.ContainsKey("EventSessionCpdNumber") Then vWhereFields.Add("sc.event_session_cpd_number", mvParameters("EventSessionCpdNumber").IntegerValue)

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "session_cpd sc", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vColumns)

    End Sub

    Private Sub GetDocumentLinksForCPDCyclePeriods(ByRef pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntityDesc As String = GetLinkEntityTypeDescription("CE")
      Dim vAttrs As String = "cpp.contact_cpd_period_number,dll.communications_log_number,cpp.contact_cpd_period_number_desc"
      vAttrs &= ", 'CE' AS entity_type,'" & vEntityDesc & "' AS entity_type_desc"
      Dim vCols As String = "contact_cpd_period_number,communications_log_number,contact_cpd_period_number_desc"

      If pAddType Then
        vAttrs = "link_type," & vAttrs
        vCols = "link_type," & vCols & ",,,CPD_PERIOD_DOC,LINK_TYPE"
      End If
      vCols &= ",entity_type,entity_type_desc"

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("contact_cpd_periods cpp", "dll.contact_cpd_period_number", "cpp.contact_cpd_period_number")})
      Dim vWhereFields As New CDBFields(New CDBField("dll.communications_log_number", mvParameters("DocumentNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "document_log_links dll", vWhereFields, "cpp.contact_cpd_period_number_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
    End Sub

    Private Sub GetDocumentLinksForCPDPoints(ByRef pDataTable As CDBDataTable, ByVal pAddType As Boolean)
      Dim vEntity_desc As String = GetLinkEntityTypeDescription("CP")
      Dim vAttrs As String = "cpp.contact_cpd_point_number,dll.communications_log_number,cct.cpd_category_type_desc,cc.cpd_category_desc"
      vAttrs &= ",'CP' AS entity_type,'" & vEntity_desc & "' AS entity_type_desc"
      Dim vCols As String = "contact_cpd_point_number,communications_log_number,CPD_POINT_DESC"

      If pAddType Then
        vAttrs = "link_type," & vAttrs
        vCols = "link_type," & vCols & ",,,CPD_POINT_DOC,LINK_TYPE"
      End If
      vCols &= ",entity_type,entity_type_desc"

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("contact_cpd_points cpp", "dll.contact_cpd_point_number", "cpp.contact_cpd_point_number")})
      vAnsiJoins.Add("cpd_categories cc", "cpp.cpd_category", "cc.cpd_category", "cpp.cpd_category_type", "cc.cpd_category_type")
      vAnsiJoins.Add("cpd_category_types cct", "cc.cpd_category_type", "cct.cpd_category_type")
      Dim vWhereFields As New CDBFields(New CDBField("dll.communications_log_number", mvParameters("DocumentNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "document_log_links dll", vWhereFields, "cct.cpd_category_type_desc,cc.cpd_category_desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)
    End Sub

    Private Sub GetDocumentLinksForPositions(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "dll.link_type, cp.contact_position_number, dll.communications_log_number, cp.position"
      vFields &= ", 'P' AS entity_type, '" & GetLinkEntityTypeDescription("P") & "' AS entity_type_desc,"
      Dim vContact As New Contact(mvEnv)
      vFields &= vContact.GetRecordSetFieldsName

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contact_positions cp", "dll.contact_position_number", "cp.contact_position_number")
      vAnsiJoins.Add("contacts c", "cp.contact_number", "c.contact_number")

      Dim vWhereFields As New CDBFields(New CDBField("dll.communications_log_number", mvParameters("DocumentNumber").IntegerValue))

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "document_log_links dll", vWhereFields, "cp.position", vAnsiJoins)

      pDataTable.FillFromSQL(mvEnv, vSQLStatement, "link_type,contact_position_number,communications_log_number,POSITION_LINK_DESC,,,POSITION_LINK_TYPE,LINK_TYPE,entity_type,entity_type_desc")
    End Sub

    Private Sub GetFindCPDCyclePeriods(ByRef pDataTable As CDBDataTable)
      Dim vAttrs As String = "ccc.contact_cpd_cycle_number, ccc.cpd_cycle_type, cct.cpd_cycle_type_desc, cct.start_month, cct.end_month, ccc.start_date, ccc.end_date, "
      vAttrs &= "ccp.contact_cpd_period_number, ccp.contact_cpd_period_number_desc, ccp.start_date, ccp.end_date, ccc.cpd_cycle_status, ccs.cpd_cycle_status_desc, cct.cpd_type, "

      Dim vContact As New Contact(mvEnv)
      Dim vCols As String = vAttrs.Replace(" ", "") & "cpd_type_desc,contact_number,CONTACT_NAME"
      vAttrs &= "'' AS cpd_type_desc, " & vContact.GetRecordSetFieldsName()

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("cpd_cycle_types cct", "ccc.cpd_cycle_type", "cct.cpd_cycle_type")
      vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number")
      vAnsiJoins.Add("contacts c", "ccc.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("cpd_cycle_statuses ccs", "ccc.cpd_cycle_status", "ccs.cpd_cycle_status")

      Dim vWhereFields As New CDBFields()
      If mvParameters.ContainsKey("ContactCpdCycleNumber") Then vWhereFields.Add("ccc.contact_cpd_cycle_number", mvParameters("ContactCpdCycleNumber").IntegerValue)
      If mvParameters.ContainsKey("ContactCpdPeriodNumber") Then vWhereFields.Add("ccp.contact_cpd_period_number", mvParameters("ContactCpdPeriodNumber").IntegerValue)
      If mvParameters.ContainsKey("ContactNumber") Then vWhereFields.Add("ccc.contact_number", mvParameters("ContactNumber").IntegerValue)
      If mvParameters.ContainsKey("CpdCycleType") Then vWhereFields.Add("ccc.cpd_cycle_type", mvParameters("CpdCycleType").Value)
      If mvParameters.ContainsKey("CpdCycleStatus") Then vWhereFields.Add("ccc.cpd_cycle_status", mvParameters("CpdCycleStatus").Value)

      Dim vCycleStartDate As Nullable(Of Date) = Nothing
      Dim vCycleEndDate As Nullable(Of Date) = Nothing
      Dim vDate As Date
      If mvParameters.ContainsKey("CycleStartDate") AndAlso Date.TryParse(mvParameters("CycleStartDate").Value, vDate) Then
        vCycleStartDate = vDate
      End If
      If mvParameters.ContainsKey("CycleEndDate") AndAlso Date.TryParse(mvParameters("CycleEndDate").Value, vDate) Then
        vCycleEndDate = vDate
      End If
      If vCycleStartDate.HasValue Then vWhereFields.Add("ccc.start_date", CDBField.FieldTypes.cftDate, vCycleStartDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If vCycleEndDate.HasValue Then vWhereFields.Add("ccc.end_date", CDBField.FieldTypes.cftDate, vCycleEndDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThanEqual)

      Dim vPeriodStartDate As Nullable(Of Date) = Nothing
      Dim vPeriodEndDate As Nullable(Of Date) = Nothing
      If mvParameters.ContainsKey("PeriodStartDate") AndAlso Date.TryParse(mvParameters("PeriodStartDate").Value, vDate) Then
        vPeriodStartDate = vDate
      End If
      If mvParameters.ContainsKey("PeriodEndDate") AndAlso Date.TryParse(mvParameters("PeriodEndDate").Value, vDate) Then
        vPeriodEndDate = vDate
      End If
      If vPeriodStartDate.HasValue Then vWhereFields.Add("ccp.start_date", CDBField.FieldTypes.cftDate, vPeriodStartDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If vPeriodEndDate.HasValue Then vWhereFields.Add("ccp.end_date", CDBField.FieldTypes.cftDate, vPeriodEndDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThanEqual)

      Dim vOrderBy As String = "ccc.start_date DESC, ccp.start_date DESC"

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_cycles ccc", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)

      For vIndex As Integer = 0 To pDataTable.Rows.Count - 1
        If pDataTable.Rows(vIndex).Item("CPDType") = "O" Then
          pDataTable.Rows(vIndex).Item("CPDTypeDesc") = ProjectText.CPDTypeObjectives
        Else
          pDataTable.Rows(vIndex).Item("CPDTypeDesc") = ProjectText.CPDTypePoints
        End If
      Next

    End Sub

    Private Sub GetFindCPDPoints(ByRef pDataTable As CDBDataTable)
      Dim vAttrs As String = "ccp.contact_cpd_point_number, cco.contact_cpd_period_number, ccp.cpd_category_type, ccp.cpd_category, ccp.cpd_points, ccp.cpd_points_2, ccp.points_date, "
      vAttrs &= "ccp.activity, ccp.activity_value, ccp.cpd_item_type, ct.cpd_category_type_desc, cc.cpd_category_desc, a.activity_desc, av.activity_value_desc, cit.cpd_item_type_desc, "
      vAttrs &= "cco.contact_cpd_period_number_desc, ccy.cpd_cycle_type, cy.cpd_cycle_type_desc,"

      Dim vContact As New Contact(mvEnv)
      Dim vCols As String = vAttrs.Replace(" ", "") & "contact_number,CONTACT_NAME"
      vAttrs &= vContact.GetRecordSetFieldsName()

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("cpd_category_types ct", "ccp.cpd_category_type", "ct.cpd_category_type")
      vAnsiJoins.Add("cpd_categories cc", "ccp.cpd_category_type", "cc.cpd_category_type", "ccp.cpd_category", "cc.cpd_category")
      vAnsiJoins.Add("contacts c", "ccp.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("cpd_item_types cit", "ccp.cpd_item_type", "cit.cpd_item_type")
      vAnsiJoins.AddLeftOuterJoin("activities a", "ccp.activity", "a.activity")
      vAnsiJoins.AddLeftOuterJoin("activity_values av", "ccp.activity", "av.activity", "ccp.activity_value", "av.activity_value")
      vAnsiJoins.AddLeftOuterJoin("contact_cpd_periods cco", "ccp.contact_cpd_period_number", "cco.contact_cpd_period_number")
      vAnsiJoins.AddLeftOuterJoin("contact_cpd_cycles ccy", "cco.contact_cpd_cycle_number", "ccy.contact_cpd_cycle_number")
      vAnsiJoins.AddLeftOuterJoin("cpd_cycle_types cy", "ccy.cpd_cycle_type", "cy.cpd_cycle_type")

      Dim vWhereFields As New CDBFields()
      If mvParameters.ContainsKey("ContactNumber") Then vWhereFields.Add("ccp.contact_number", mvParameters("ContactNumber").IntegerValue)
      If mvParameters.ContainsKey("ContactCpdPointNumber") Then vWhereFields.Add("ccp.contact_cpd_point_number", mvParameters("ContactCpdPointNumber").IntegerValue)
      If mvParameters.ContainsKey("CpdCategoryType") Then vWhereFields.Add("ccp.cpd_category_type", mvParameters("CpdCategoryType").Value)
      If mvParameters.ContainsKey("CpdCategory") Then vWhereFields.Add("ccp.cpd_category", mvParameters("CpdCategory").Value)

      Dim vDate As Date
      If mvParameters.ContainsKey("PointsDate") AndAlso Date.TryParse(mvParameters("PointsDate").Value, vDate) Then
        vWhereFields.Add("ccp.points_date", CDBField.FieldTypes.cftDate, vDate.ToString(CAREDateFormat))
      End If

      If mvParameters.ContainsKey("ContactCpdPeriodNumber") Then vWhereFields.Add("cco.contact_cpd_period_number", mvParameters("ContactCpdPeriodNumber").IntegerValue)
      If mvParameters.ContainsKey("ContactCpdCycleNumber") Then vWhereFields.Add("ccy.contact_cpd_cycle_number", mvParameters("ContactCpdCycleNumber").IntegerValue)
      If mvParameters.ContainsKey("CpdCycleType") Then vWhereFields.Add("ccy.cpd_cycle_type", mvParameters("CpdCycleType").Value)
      If mvParameters.ContainsKey("CpdItemType") Then vWhereFields.Add("ccp.cpd_item_type", mvParameters("CpdItemType").Value)

      Dim vCycleStartDate As Nullable(Of Date) = Nothing
      Dim vCycleEndDate As Nullable(Of Date) = Nothing
      If mvParameters.ContainsKey("CycleStartDate") AndAlso Date.TryParse(mvParameters("CycleStartDate").Value, vDate) Then
        vCycleStartDate = vDate
      End If
      If mvParameters.ContainsKey("CycleEndDate") AndAlso Date.TryParse(mvParameters("CycleEndDate").Value, vDate) Then
        vCycleEndDate = vDate
      End If
      If vCycleStartDate.HasValue Then vWhereFields.Add("ccy.start_date", CDBField.FieldTypes.cftDate, vCycleStartDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If vCycleEndDate.HasValue Then vWhereFields.Add("ccy.end_date", CDBField.FieldTypes.cftDate, vCycleEndDate.Value.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThanEqual)

      Dim vOrderBy As String = "ccp.points_date DESC, ct.cpd_category_type_desc DESC, cc.cpd_category_desc DESC"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_points ccp", vWhereFields, vOrderBy, vAnsiJoins)

      pDataTable.FillFromSQL(mvEnv, vSQLStatement, vCols)

    End Sub

    Private Sub GetTransactionSalesLedgerAnalysis(ByVal pDataTable As CDBDataTable)
      'Retrieve the breakdown of a sales ledger line
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      vBTA.Init(mvParameters("BatchNumber").IntegerValue, mvParameters("TransactionNumber").IntegerValue, mvParameters("LineNumber").IntegerValue)
      If vBTA.Existing = False Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "BatchNumber,TransactionNumber,LineNumber")

      'Get the total amount of sales ledger cash received
      Dim vTotalCashAmount As Double = 0
      Dim vWhereFields As New CDBFields({New CDBField("i.batch_number", vBTA.BatchNumber), New CDBField("i.transaction_number", vBTA.TransactionNumber)})
      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("invoice_details id", "i.batch_number", "id.batch_number", "i.transaction_number", "id.transaction_number")})
      vAnsiJoins.Add("batch_transaction_analysis bta", "id.batch_number", "bta.batch_number", "id.transaction_number", "bta.transaction_number", "id.line_number", "bta.line_number")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "SUM(bta.amount) AS invoice_amount", "invoices i", vWhereFields, "", vAnsiJoins)
      vTotalCashAmount = DoubleValue(vSQLStatement.GetValue)

      'Get the total amount allocated to invoices
      Dim vTotalInvoicePayAmount As Double = 0
      vSQLStatement = New SQLStatement(mvEnv.Connection, "SUM(amount) AS amount_paid", "invoice_payment_history i", vWhereFields)
      vTotalInvoicePayAmount = DoubleValue(vSQLStatement.GetValue)

      'Get the total amount unallocated
      Dim vTotalUnallocatedAmount As Double = FixTwoPlaces((vTotalCashAmount - vTotalInvoicePayAmount))
      If vTotalUnallocatedAmount < 0 Then vTotalUnallocatedAmount = 0    'Just in case!!

      Dim vLineInvoicePayAmount As Double = 0
      Dim vLineUnallocatedAmount As Double = 0
      If vTotalInvoicePayAmount > 0 Then
        'First see if we can find an invoice payment for this batch/transaction/line
        Dim vIPH As New InvoicePaymentHistory(mvEnv)
        vIPH.InitFromBatchTransactionLine(vBTA.BatchNumber, vBTA.TransactionNumber, vBTA.LineNumber)
        If vIPH.Existing = True AndAlso vIPH.Amount.Equals(vBTA.Amount) Then
          If String.IsNullOrEmpty(vIPH.Status) Then
            'Invoice payment
            vLineInvoicePayAmount = vIPH.Amount
          End If
        Else
          'Payment has since been adjusted OR no record found
          'As allocations does not always use the correct line number we may have allocated it
        End If
        If vLineInvoicePayAmount > 0 Then
          'Something was allocated
          vLineUnallocatedAmount = FixTwoPlaces(vBTA.Amount - vLineInvoicePayAmount)
          If vLineUnallocatedAmount < 0 Then
            'Full amount allocated
            vLineInvoicePayAmount = vBTA.Amount
            vLineUnallocatedAmount = 0
          End If
        Else
          'Could not find a specific allocation for this line
          vLineUnallocatedAmount = FixTwoPlaces((vTotalCashAmount - vTotalInvoicePayAmount))
          If vLineUnallocatedAmount >= 0 Then
            If vLineUnallocatedAmount > vBTA.Amount Then vLineUnallocatedAmount = vBTA.Amount
            vLineInvoicePayAmount = FixTwoPlaces((vBTA.Amount - vLineUnallocatedAmount))
          Else
            vLineUnallocatedAmount = 0
            vLineInvoicePayAmount = vBTA.Amount
          End If
        End If
      Else
        'No invoice payments so all must be unallocated
        vLineInvoicePayAmount = 0
        vLineUnallocatedAmount = vBTA.Amount
      End If

      'Now add a row to the data table
      'BatchNumber,TransactionNumber,LineNumber,TransactionSLAmount,LineAmount,InvoicePaymentAmount,UnallocatedAmount
      Dim vRow As New CDBDataRow(pDataTable.Columns, 0)
      vRow.Item(1) = vBTA.BatchNumber.ToString
      vRow.Item(2) = vBTA.TransactionNumber.ToString
      vRow.Item(3) = vBTA.LineNumber.ToString
      vRow.Item(4) = vTotalCashAmount.ToString("0.00")
      vRow.Item(5) = vBTA.Amount.ToString("0.00")
      vRow.Item(6) = vLineInvoicePayAmount.ToString("0.00")
      vRow.Item(7) = vLineUnallocatedAmount.ToString("0.00")
      pDataTable.Rows.Add(vRow)

    End Sub

    Private Sub GetTraderAlerts(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "cal.trader_application_number,cal.contact_alert_link_number,ca.contact_alert,ca.contact_alert_desc,ca.contact_alert_sql,ca.contact_alert_message,ca.show_as_dialog,ca.contact_alert_message_type,ca.sequence_number"

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("contact_alerts ca", "cal.contact_alert", "ca.contact_alert")})
      Dim vWhereFields As New CDBFields(New CDBField("cal.trader_application_number", mvParameters("FpApplicationNumber").IntegerValue))
      vWhereFields.Add("ca.contact_alert_type", "F")
      If mvParameters.HasValue("ContactAlert") Then
        vWhereFields.Add("ca.contact_alert", mvParameters("ContactAlert").Value)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_alert_links cal", vWhereFields, "ca.sequence_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      For Each vRow As CDBDataRow In pDataTable.Rows
        vRow.SetYNValue("ShowAsDialog")
        Select Case vRow.Item("AlertMessageType").ToString.ToUpper
          Case "E"
            vRow.Item("AlertMessageType") = ProjectText.ContactAlertErrorMessageType
          Case "W"
            vRow.Item("AlertMessageType") = ProjectText.ContactAlertWarningMessageType
        End Select
      Next

    End Sub

    Private Sub GetContactFinanceAlerts(ByVal pDataTable As CDBDataTable)
      Dim vContactAlert As New ContactAlert(mvEnv)
      vContactAlert.GetContactFinanceAlerts(pDataTable, mvParameters("ContactNumber").IntegerValue, mvParameters("FpApplicationNumber").IntegerValue)
    End Sub

  End Class
End Namespace
