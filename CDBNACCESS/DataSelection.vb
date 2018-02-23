Namespace Access

  Partial Public Class DataSelection

    Public Enum DataSelectionTypes
      dstNone = 0
      dstContactActions
      dstContactAddresses
      dstContactAddressUsages
      dstContactCategories
      dstContactPositionCategories
      dstContactRoles
      dstContactSuppressions
      dstContactPositions
      dstContactProcessedTransactions
      dstContactUnProcessedTransactions
      dstContactCancelledProvisionalTrans
      dstContactLinksTo
      dstContactLinksFrom
      dstContactStatusHistory
      dstContactEventBookings
      dstContactEventDelegates
      dstContactEventSessions
      dstContactEventRoomBookings
      dstContactEventRoomsAllocated
      dstContactEventOrganiser
      dstContactEventPersonnel
      dstContactPurchaseOrders
      dstContactPurchaseInvoices
      dstContactCommsNumbers
      dstContactCommsNumbersEdit
      dstOrganisationContactCommsNumbers
      dstContactDocuments
      dstDistinctContactDocuments
      dstContactMailings
      dstContactExternalReferences
      dstContactDepartmentNotes
      dstContactMemberships
      dstContactJournal
      dstContactCategoryGraphData
      dstContactStickyNotes
      dstContactOwners
      dstContactHPLinks
      dstContactHPCategories
      dstContactHPCategoryValues
      dstContactDeptCategories
      dstContactBankAccounts
      dstContactCreditCards
      dstContactCreditCardAuthorities
      dstContactStandingOrders
      dstContactDirectDebits
      dstContactBackOrders
      dstContactDespatchNotes
      dstContactFinLinksReceived
      dstContactFinLinksDonated
      dstContactPaymentPlans
      dstContactServiceBookings
      dstContactSubscriptions
      dstContactDBANotes
      dstContactCovenants
      dstContactMembershipDetails
      dstContactAddressPositionAndOrg
      dstContactPreviousNames

      dstContactSponsorshipClaimedPayments
      dstContactSponsorshipUnClaimedPayments
      dstContactGiftAidDeclarations
      dstContactCreditCustomers
      dstContactOutstandingInvoices
      dstContactCashInvoices                    'UnAllocated Cash and Sundry Credit Notes

      dstContactCPDCyclesEdit

      dstClaimedPayments
      dstUnClaimedPayments

      dstDespatchStock
      dstFinancialHistoryDetails
      dstBACSAmendments
      dstPostPointRecipients
      dstRates
      dstProductWarehouses
      dstTransactionAnalysis
      dstDocuments
      dstDistinctDocuments
      dstDistinctExternalDocuments
      dstDocumentSubjects
      dstDocumentHistory
      dstDocumentLinks
      dstDocumentRelatedDocuments
      dstActionSubjects
      dstDuplicateContacts
      dstDuplicateOrganisations

      dstPurchaseOrderDetails
      dstPurchaseOrderPayments
      dstPurchaseInvoiceDetails
      dstUnauthorisedPOPayments

      dstDocumentContactLinks
      dstDocumentOrganisationLinks
      dstDocumentDocumentLinks
      dstDocumentTransactionLinks
      dstActionContactLinks
      dstActionOrganisationLinks
      dstActionDocumentLinks
      dstMeetingContactLinks
      dstMeetingOrganisationLinks
      dstMeetingDocumentLinks

      dstActionLinks
      dstMeetingLinks

      dstEventAttendees
      dstEventBookingOptions
      dstEventBookingOptionSessions
      dstEventBookingSessions
      dstEventCurrentAttendees
      dstEventPersonnel
      dstEventSessions
      dstEventSessionTests
      dstEventSubmissions
      dstEventAccommodation
      dstEventResources
      dstEventVenueBookings
      dstEventPIS

      dstIncentives
      dstFulFilledContactIncentives
      dstFulFilledPayPlanIncentives
      dstUnFulFilledContactIncentives
      dstUnFulFilledPayPlanIncentives
      dstPaymentPlanPayments
      dstPaymentPlanDetails
      dstPaymentPlanSubscriptions
      dstPaymentPlanMembers
      dstPaymentPlanOutstandingOPS
      dstPaymentPlanPaymentDetails

      dstMembershipPaymentPlanDetails
      dstMembershipOtherMembers
      dstMembershipChanges

      dstCovenantGiftAidClaims
      dstCovenentClaims
      dstCovenentPayments

      dstContactHeaderInfo
      dstContactHeaderCommsNumbers
      dstContactHeaderHPCategories
      dstContactHeaderDeptCategories
      dstContactHeaderHPLinks

      dstContactGAYEPledges
      dstContactGAYEPledgePayments

      dstContactGAYEPostTaxPledges
      dstContactGAYEPostTaxPledgePayments

      dstContactEMailAddresses

      dstDepartmentActivities
      dstDepartmentActivityValues

      dstBrowserContactPositions

      dstCPDSummary
      dstCPDDetails
      dstCPDPointsEdit

      dstContactSalesLedgerItems

      dstSelectItemAddresses
      dstSelectItemSelectionSets
      dstSelectItemCreditAccount

      dstEventLoanItems
      dstEventAuthoriseExpenses
      dstEventRoomBookings
      dstEventBookings
      dstEventCosts
      dstEventContacts
      dstEventDelegateIncome
      dstEventFinancialHistory
      dstEventFinancialLinks
      dstEventPersonnelTasks
      dstEventResults
      dstEventCandidates
      dstEventTopics
      dstEventSessionActivities
      dstEventOwners

      dstEVWaitingBookings
      dstEVWaitingDelegates

      dstContactNotifications
      dstSelectionSteps
      dstSelectionSetContacts
      dstSelectionSetCommsNumbers
      dstCriteriaSets

      dstSelectionPages
      dstEventSelectionPages
      dstContactInformation
      dstCustomFormData
      dstActivitiesDataSheet
      dstRelationshipsDataSheet
      dstPackProductDataSheet

      dstAppealCollections
      dstCampaigns
      dstCampaignAppeals
      dstCampaignSegments
      dstCampaignCollections
      dstCampaignInfo

      dstActionFinder

      dstDelegateActivities
      dstDelegateLinks

      dstCollectionRegions
      dstCollectionPoints
      dstCollectionResources
      dstMannedCollectors
      dstCollectorShifts
      dstMannedCollectionBoxes
      dstUnMannedCollectionBoxes
      dstContactCollectionBoxes
      'dstPayingInSlips
      dstAppealResources
      dstCollectionBoxesForPayment
      dstCollectionPayments
      dstCollectionPIS
      dsth2hCollectionPIS
      dstH2HCollectors
      dstSegmentProducts
      dstAppealBudgets
      dstAppealBudgetDetails
      dstTickBoxes
      dstSegmentCostCentres
      dstVariableParameters
      dstSuppliers
      dstAppealTypes

      dstContactMannedCollections
      dstContactUnMannedCollections
      dstContactH2HCollections
      dstContactCollectionPayments

      dstContactDepartmentHistory
      dstEventSources
      dstEventMailings
      dstEventVenueData
      dstEventOrganiserData
      dstEventHeaderInfo
      dstEventBookingDelegates
      dstEventRoomBookingAllocations

      dstContactFinder
      dstOrganisationFinder

      dstServiceStartDays
      dstGeographicalRegions
      dstSalesContacts
      dstPersonnelContacts
      dstContactAppointments
      dstContactAccounts
      dstTransactionDetails

      dstEMailContacts
      dstEMailOrganisations
      dstContactScores
      dstContactPerformances
      dstContactCommunicationUsages
      dstContactPictureDocuments

      dstContactSourceFromLastMailing
      dstActionOutline
      dstPaymentPlanAmendmentHistory
      dstDashboardData
      dstEventDocuments
      dstContactFundRaisingRequests         'Used by the .NET data access only. Please retain this position in the enum list
      dstSelectionSetAppointments           'Used by the .NET data access only. Please retain this position in the enum list
      dstFundraisingRequestTargets          'Used by the .NET data access only. Please retain this position in the enum list
      dstContactAppropriateCertificates
      dstServiceControlRestrictions
      dstTopicDataSheet                     'Used by the .NET data access only. Please retain this position in the enum list
      dstCampaignMailOrCount                'Used by the .NET data access only. Please retain this position in the enum list
      dstContactFundraisingEvents           'Used by the .NET data access only. Please retain this position in the enum list
      dstContactFundraisingEventFinder      'Used by the .NET data access only. Please retain this position in the enum list
      dstPriorActions
      dstMembershipGroups                   'Used by the .NET data access only. Please retain this position in the enum list
      dstMembershipGroupHistory             'Used by the .NET data access only. Please retain this position in the enum list
      dstFundraisingEventAnalysis           'Used by the .NET data access only. Please retain this position in the enum list
      dstPurchaseOrderInformation           'Used by the .NET data access only. Please retain this position in the enum list
      dstPurchaseInvoiceInformation         'Used by the .NET data access only. Please retain this position in the enum list
      dstChequeInformation                  'Used by the .NET data access only. Please retain this position in the enum list
      dstBatchProcessingInformation         'Used by the .NET data access only. Please retain this position in the enum list
      dstPickingListDetails                 'Used by the .NET data access only. Please retain this position in the enum list
      dstSuppressionDataSheet               'Used by the .NET data access only. Please retain this position in the enum list
      dstContactMailingDocumentsFinder      'Used by the .NET data access only. Please retain this position in the enum list
      dstMailingFinder                      'Used by the .NET data access only. Please retain this position in the enum list
      dstActionLinkEmailAddresses           'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacy                      'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyBequests              'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyAssets                'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyLinks                 'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyTaxCertificates       'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyExpenses              'Used by the .NET data access only. Please retain this position in the enum list
      dstContactLegacyActions               'Used by the .NET data access only. Please retain this position in the enum list
      dstLegacyBequestForecasts             'Used by the .NET data access only. Please retain this position in the enum list
      dstLegacyBequestReceipts              'Used by the .NET data access only. Please retain this position in the enum list
      dstEventPersonnelFinder               'Used by the .NET data access only. Please retain this position in the enum list
      dstEventPersonnelAppointmentFinder    'Used by the .NET data access only. Please retain this position in the enum list
      dstCampaignCosts                      'Used by the .NET data access only. Please retain this position in the enum list
      dstCampaignRoles                      'Used by the .NET data access only. Please retain this position in the enum list
      dstTextSearch                         'Used by the .NET data access only. Please retain this position in the enum list
      dstEventBookingTransactions           'Used by the .NET data access only. Please retain this position in the enum list
      dstContactRegisteredUsers             'Used by the .NET data access only. Please retain this position in the enum list
      dstMembershipSummaryMembers           'Used by the .NET data access only. Please retain this position in the enum list
      dstServiceBookingTransactions
      dstContactEmailingsLinks              'Used by the .NET data access only. Please retain this position in the enum list
      dstPurchaseInvoiceChequeInformation   'Used by the .NET data access only. Please retain this position in the enum list
      dstContactCommsInformation            'Used by the .NET data access only. Please retain this position in the enum list
      dstContactPositionLinks               'Used by the .NET data access only. Please retain this position in the enum list
      dstFundraisingActions
      dstFundraisingPaymentHistory
      dstFundraisingPaymentSchedule
      dstFundRequestExpectedAmountHistory
      dstFundRequestStatusHistory
      dstContactCommunicationHistory        'Used by the .NET data access only.
      dstContactAddressAndUsage
      dstGeneralMailingSelectionSets
      dstCriteriaSetDetails
      dstContactTokens

      dstFindBatch
      dstFindCCCA
      dstFindCovenant
      dstFindDirectDebit
      dstFindEvent
      dstFindEventBooking
      dstFindGAD
      dstFindLegacy
      dstFindMeeting
      dstFindMember
      dstFindOpenBatch
      dstFindPaymentPlan
      dstFindProduct
      dstFindStandingOrder
      dstFindTransaction
      dstFindInvoice
      dstFindManualSOReconciliation
      dstFindVenue
      dstFindGiveAsYouEarn                'Pre Tax Payroll Giving
      dstFindPurchaseOrder
      dstFindInternalResource
      dstFindServiceProduct
      dstFindCommunication
      dstFindUniservPhoneBook
      dstFindCampaign
      dstFindPostTaxPayrollGiving
      dstFindStandardDocuments
      dstQueryByExampleContacts
      dstQueryByExampleOrganisations
      dstQueryByExampleEvents
      dstMaximumDataSelectionTypePlusOne
      dstContactAlerts
      dstContactNetwork
      dstDespatchNotes
      dstDuplicateContactRecords
      dstSelectAwaitListConfirmation
      dstPackedProductDataSheet
      dstJobSchedules
      dstJobProcessors
      dstConfig
      dstSystemModuleUsers
      dstReportData
      dstReportSectionData
      dstReportParameters
      dstReportSectionDetail
      dstReportVersion
      dstReportControl
      dstOwnershipData
      dstOwnershipUsers
      dstOwnershipDepartments
      dstOwnershipUserInformation
      dstPostcodeProximity
      dstCommunicationsLogDocClass
      dstMeetings
      dstContactMeetings
      dstServiceProductContacts
      dstEmailAutoReplyText
      dstEntityAlerts
      dstSalesTransactions
      dstDeliveryTransactions
      dstSalesTransactionAnalysis
      dstDeliveryTransactionAnalysis
      dstEventDelegateMailing
      dstCPDObjectives
      dstCPDObjectivesEdit
      dstEventActions
      dstAppealActions
      dstContactSurveys
      dstContactSurveyResponses
      dstContactSalesLedgerReceipts
      dstWebProducts
      dstWebEvents
      dstWebBookingOptions
      dstContactEventBookingDelegates
      dstDocumentActions ' Enum for action related to Documents
      dstWebMembershipType
      dstActivityFromActivityGroup
      dstEventCategories
      dstWebEventBookings
      dstWebSurveys
      dstWebDirectoryEntries
      dstBankTransactions
      dstWebDocuments
      dstWebRelatedContacts
      dstWebRelatedOrganisations
      dstWebContacts
      dstContactAppointmentDetails
      dstDefaultName
      dstPurchaseOrderHistory
      dstContactLoans
      dstContactDirectoryUsage

      dstContactExamSummary
      dstContactExamSummaryItems
      dstContactExamSummaryList
      dstContactExamDetails
      dstContactExamDetailItems
      dstContactExamDetailList
      dstContactExamExemptions
      dstExamPersonnelFinder
      dstDataUpdates
      dstContactAddressesAndPositions
      dstDuplicateOrganisationsForRegistration
      dstContactAddressesWithUsages
      dstContactCommsNumbersWithUsages
      dstContactCPDPointsWithoutCycle
      dstContactPaymentPlansPayments
      dstLoanInterestRates
      dstWebExams
      dstWebMemberOrganisations
      dstWebExamBookings
      dstWebExamHistory
      dstContactAmendments
      dstContactAmendmentDetails
      dstMeetingActions
      dstEventFinder
      dstContactExamCertificates
      dstContactExamCertificateItems
      dstContactExamCertificateReprints
      dstFundraisingRequests
      dstFundraisingDocuments
      dstExamScheduleFinder
      dstEventSessionCPD
      dstContactCPDCycleDocuments
      dstContactCPDPointDocuments
      dstFindCPDCyclePeriods
      dstFindCPDPoints
      dstContactViewOrganisations
      dstContactPositionActions
      dstContactPositionDocuments
      dstContactPositionTimesheets
      dstSalesLedgerAnalysis
      dstTraderAlerts
      dstContactFinanceAlerts
    End Enum

    Public Enum DataSelectionUsages
      dsuNone = 0
      dsuCare = 1                 'Unsupported in this class
      dsuWEBServices = 2
      dsuSmartClient = 4
    End Enum

    Public Enum DataSelectionListType
      dsltUser = 0
      dsltDefault
      dsltEditing
    End Enum

    Protected mvEnv As CDBEnvironment
    Private mvType As DataSelectionTypes
    Protected mvUsage As DataSelectionUsages
    Private mvDepartment As String = ""
    Private mvLogname As String = ""
    Private mvContactGroup As String = ""
    Private mvHeaderLines As Integer = 1
    Private mvDefaults As Boolean = True
    Private mvCustomFieldNames As String
    Protected mvAvailableUsages As DataSelectionUsages
    Private mvCustomForm As CustomForm
    Protected mvContact As Contact
    Protected mvGroupCode As String

    Protected mvRequiredItems As String = ""
    Protected mvDescription As String = ""
    Protected mvResultColumns As String = ""
    Protected mvSelectColumns As String = ""      'The default list of columns normally displayed
    Protected mvDisplayTitle As String = ""
    Protected mvMaintenanceDesc As String = ""    'The maintenance name to be used in 'New {0}' and 'Edit {1}'
    Protected mvHeadings As String = ""           'The default list of headings
    Protected mvWidths As String = ""             'The default list of widths
    Protected mvCode As String = ""
    Protected mvParameters As CDBParameters
    Protected mvDisplayListItems As CollectionList(Of DisplayListItem)
    Private mvDataFinder As DataFinder
    Protected mvDataSelectionListType As DataSelectionListType
    Private mvLinkEntityTypes As DataTable

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pListType As DataSelectionListType)
      Init(pEnv, pDataSelectionType, Nothing, pListType, DataSelectionUsages.dsuCare, "")
    End Sub
    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, "")
    End Sub
    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages, ByVal pContact As Contact)
      mvContact = pContact
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, "")
    End Sub
    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages, ByVal pGroup As String)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, pGroup)
    End Sub

    Public Sub AddParameter(ByVal pName As String, ByVal pFieldType As CDBField.FieldTypes, ByVal pValue As String)
      If mvParameters Is Nothing Then mvParameters = New CDBParameters
      mvParameters.Add(pName, pFieldType, pValue)
    End Sub


    Private Sub Init(ByVal pEnv As CDBEnvironment, ByVal pType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsage As DataSelectionUsages, ByVal pGroup As String)
      mvEnv = pEnv
      mvType = pType
      mvParameters = pParams
      mvUsage = pUsage
      mvGroupCode = pGroup
      mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
      mvDisplayListItems = Nothing
      mvDataSelectionListType = pListType

      Dim vPrimaryList As Boolean

      Select Case mvType
        Case DataSelectionTypes.dstEventLoanItems
          mvResultColumns = "ContactNumber,AddressNumber,ProductCode,Product,Quantity,Issued,Due,Returned,QuantityReturned,Complete,Reference"
          mvSelectColumns = "ContactNumber,AddressNumber,ProductCode,Product,Quantity,Issued,Due,Returned,QuantityReturned,Complete,Reference"
          mvHeadings = DataSelectionText.String26402 'Contact Number,Address Number,Product Code,Product,Qty,Issued,Due,Returned,Qty,Complete,Reference
          mvWidths = "1,1,1,2500,600,1150,1150,1150,600,600,2000"
          mvRequiredItems = "ProductCode,Issued,Returned,QuantityReturned,Complete,Reference"

        Case DataSelectionTypes.dstEventAuthoriseExpenses
          mvResultColumns = "SessionNumber,ContactNumber,PersonnelContactName,Expenses,EventName,AuthorisedOn,AuthorisedBy,EventPersonnelNumber,AddressNumber"
          mvSelectColumns = "SessionNumber,ContactNumber,PersonnelContactName,Expenses,EventName,AuthorisedOn,AuthorisedBy,EventPersonnelNumber,AddressNumber"
          mvHeadings = DataSelectionText.String25403 'Session Number,Contact Number,Name,Expenses,Event,Authorised On,Authorised By,Event Personnel Number
          mvWidths = "1,1,2000,1150,2400,1300,1300,1,1"
          mvRequiredItems = "PersonnelContactName,AuthorisedOn,AuthorisedBy"

        Case DataSelectionTypes.dstEventRoomBookings
          mvResultColumns = "BookingNumber,ContactName,RoomTypeDesc,Quantity,FromDate,ToDate,BookedOn,CancelledOn,CancelledBy,BookingStatusCode,Rate,Notes,SalesContactNumber,ConfirmedOn,CancellationReasonDesc,CancellationSourceDesc,AmendedBy,AmendedOn,AddressNumber,RoomType,Product,OrganisationNumber,BatchNumber,TransactionNumber,LineNumber,Capacity,EnforceAllocation,ContactNumber,CancellationReason,CancellationSource"
          mvSelectColumns = "BookingNumber,ContactName,RoomTypeDesc,Quantity,FromDate,ToDate,BookedOn,CancelledOn,CancelledBy,CancellationReasonDesc,CancellationSourceDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String25640 'Booking Number,Booked By,Type,Qty,From,To,Booked On,Cancelled On,Cancelled By,Cancellation Reason,Amended By,Amended On
          mvWidths = "1,3000,1400,600,1200,1200,1200,1200,1500,3500,3500,1200,1200"
          mvDescription = DataSelectionText.String17760    'Event Room Bookings
          mvCode = "ERB"
          mvRequiredItems = "ContactName,RoomTypeDesc,Quantity,FromDate,ToDate,BookedOn,CancellationReason,CancellationSource,AmendedBy,AmendedOn,EnforceAllocation"

        Case DataSelectionTypes.dstEventBookings
          mvResultColumns = "BookingNumber,ContactName,BookingDate,BookingStatus,Quantity,OptionDesc,AddressNumber,Rate,SalesContactNumber,Notes,CancelledOn,CancelledBy,CancellationReason,CancellationSource,OptionNumber,Product,OrganisationNumber,BatchNumber,TransactionNumber,LineNumber,AdultQuantity,ChildQuantity,StartTime,EndTime,InvoiceNumber,InvoiceRePrintCount,BatchType,BookingAmount,PayerContactNumber,PayerName,InvoicePayStatus,InvoicePayStatusDesc,ContactNumber,CancellationReasonDesc,CancellationSourceDesc,BookingStatusCode,CreditSale,InvoicePrinted"
          mvSelectColumns = "BookingNumber,ContactName,BookingDate,BookingStatus,Quantity,OptionDesc,CancelledOn,CancelledBy,CancellationReasonDesc,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String25621 'Booking Number,Booked By,Booked On,Status,Quantity,Option,Cancelled On,Cancelled By,Cancellation Reason
          mvWidths = "850,3000,1250,2100,800,3000,1200,1500,3500,3500"
          mvDescription = DataSelectionText.String17761    'Event Bookings
          mvCode = "EBK"
          mvRequiredItems = "ContactName,BookingDate,BookingStatus,Quantity,OptionDesc,AdultQuantity,ChildQuantity,StartTime,EndTime,Product,Rate,CreditSale,InvoicePrinted,BookingStatusCode,BookingAmount"

        Case DataSelectionTypes.dstEventCosts
          mvResultColumns = "CostType,TotalAmount,DueDate,Deposit,DepositPaidDate,Balance,BalancePaidDate,AmendedBy,AmendedOn,CostNumber,CostTypeDesc,SponsorshipValue,ContactNumber,AddressNumber,ItemReceived,ReserveAmount,SoldAmount,SupplierContactNumber,SupplierAddressNumber,Notes"
          mvSelectColumns = "CostTypeDesc,TotalAmount,DueDate,Deposit,DepositPaidDate,Balance,BalancePaidDate,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17762    'Cost Type,Total,Due Date,Deposit,Paid Date,Balance,Paid Date,Amended By,Amended On
          mvWidths = "2000,1000,1200,1000,1200,1000,1200,1200,1200"
          mvDescription = DataSelectionText.String17764    'Event Costs
          mvCode = "ECO"
          mvRequiredItems = "CostType"

        Case DataSelectionTypes.dstEventDelegateIncome
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,Amount,Source,SourceDesc,Product,ProductDesc,Status,Notes"
          mvSelectColumns = "TransactionDate,ProductDesc,Amount,SourceDesc,BatchNumber,TransactionNumber,LineNumber,Status,Notes"
          mvHeadings = DataSelectionText.String18173    'Transaction Date,Product,Amount,Source,Batch Number,Transaction Number,LineNumber,Status,Notes
          mvWidths = "1600,1800,1200,1800,1600,1600,1600,1200,3000"
          mvDescription = DataSelectionText.String18175    'Event Delegate Income
          mvCode = "EDI"

        Case DataSelectionTypes.dstEventContacts
          mvResultColumns = ContactNameResults() & ",AddressNumber,Relationship,AmendedBy,AmendedOn,ContactNumber,Notes"
          mvSelectColumns = "ContactName,Relationship,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17765    'Name,Relationship with Event,Amended By,Amended On
          mvWidths = "2000,2000,1300,1300"
          mvDescription = DataSelectionText.String17767    'Event Contacts
          mvCode = "ECT"
          mvRequiredItems = "ContactName"

        Case DataSelectionTypes.dstEventFinancialHistory
          mvResultColumns = ContactNameResults() & ",ContactNumber,AddressNumber,BatchNumber,TransactionNumber,LineNumber,TransactionDate,Amount,Source,SourceDesc,Product,ProductDesc,Status,Notes"
          mvSelectColumns = "TransactionDate,ContactName,ProductDesc,Amount,SourceDesc,BatchNumber,TransactionNumber,LineNumber,Status,Notes"
          mvHeadings = DataSelectionText.String18176    'Transaction Date,Name,Product,Amount,Source,Batch Number,Transaction Number,Line Number,Status,Notes
          mvWidths = "1600,2000,1800,1200,1800,1600,1600,1600,1200,3000"
          mvDescription = DataSelectionText.String18178    'Event Financial History
          mvCode = "EFH"

        Case DataSelectionTypes.dstEventFinancialLinks
          mvResultColumns = ContactNameResults() & ",ContactNumber,AddressNumber,BatchNumber,TransactionNumber,LineNumber,TransactionDate,Amount,Source,SourceDesc,Product,ProductDesc,Status,Notes"
          mvSelectColumns = "TransactionDate,ContactName,ProductDesc,Amount,SourceDesc,BatchNumber,TransactionNumber,LineNumber,Status,Notes"
          mvHeadings = DataSelectionText.String18176    'Transaction Date,Name,Product,Amount,Source,Batch Number,Transaction Number,Line Number,Status,Notes
          mvWidths = "1600,2000,1800,1200,1800,1600,1600,1600,1200,3000"
          mvDescription = DataSelectionText.String18179    'Event Financial Links
          mvCode = "EFL"

        Case DataSelectionTypes.dstEventResults
          mvResultColumns = ContactNameResults() & ",ContactNumber,SessionNumber,TestNumber,Description,TestResult,CertificateNumber,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactName,TestNumber,Description,TestResult,CertificateNumber,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18263    'Candidate,Test,Description,Result,Certificate Number,Amended By,Amended On
          mvWidths = "2000,1000,2000,1200,1400,1300,1300"
          mvDescription = DataSelectionText.String17770    'Event Results
          mvCode = "EVR"

        Case DataSelectionTypes.dstEventCandidates
          mvResultColumns = "ContactNumber," & ContactNameResults()
          mvSelectColumns = "ContactName"
          mvHeadings = DataSelectionText.String17771    'Candidate Name
          mvWidths = "1500"
          mvDescription = DataSelectionText.String17773    'Event Candidates
          mvCode = "EVC"
          mvRequiredItems = "ContactName"

        Case DataSelectionTypes.dstEventPersonnelTasks
          mvResultColumns = "EventPersonnelNumber,EventPersonnelTaskNumber,PersonnelTask,PersonnelTaskDesc,StartDate,StartTime,EndDate,EndTime,Notes,AmendedBy,AmendedOn,ContactNumber,AddressNumber,ExternalTaskID,ContactName"
          mvSelectColumns = "PersonnelTaskDesc,StartDate,StartTime,EndDate,EndTime,AmendedBy,AmendedOn,ExternalTaskID"
          mvHeadings = DataSelectionText.String18260    'Personnel Task,Start Date,Start Time,End Date,End Time,Amended By,Amended On
          mvWidths = "2500,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String18262    'Event Personnel Tasks
          mvCode = "EPT"
          mvRequiredItems = "PersonnelTask,ExternalTaskID"

        Case DataSelectionTypes.dstEventTopics
          mvResultColumns = "Topic,TopicDesc,SubTopic,SubTopicDesc,Quantity,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "TopicDesc,SubTopicDesc,Quantity,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18169     'Topic,SubTopic,Quantity,Notes,Amended By,Amended On
          mvWidths = "2000,2000,1200,2000,1200,1200"
          mvDescription = DataSelectionText.String18170    'Event Topics
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvCode = "EVT"
          mvRequiredItems = "Topic,SubTopic"

        Case DataSelectionTypes.dstEventSessionActivities
          mvResultColumns = "EventNumber,SessionNumber,Activity,ActivityValue,ActivityDesc,ActivityValueDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc"
          mvHeadings = DataSelectionText.String18264     'Category,Value
          mvWidths = "2000,2000"
          mvDescription = DataSelectionText.String22897    'Session Activities
          mvCode = "EVSA"
          mvRequiredItems = "Activity,ActivityValue"

        Case DataSelectionTypes.dstEventOwners
          mvResultColumns = "EventNumber,Department,DepartmentDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "DepartmentDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18265     'Department,Amended By,Amended On
          mvWidths = "2000,1200,1200"
          mvDescription = DataSelectionText.String22898    'Event Owners
          mvCode = "EVOW"
          mvRequiredItems = "Department"

        Case DataSelectionTypes.dstEVWaitingBookings
          mvResultColumns = ContactNameResults() & ",BookingNumber,BookingDate,Quantity,OptionDesc,BookingStatus,ContactNumber"
          mvSelectColumns = "ContactName,BookingNumber,BookingDate,Quantity,OptionDesc,BookingStatus,ContactNumber"
          mvHeadings = DataSelectionText.String17774     'Booked By,Booking Number,Booked On,Quantity,Option,Booking Status,Contact Number
          mvWidths = "3000,1000,1250,800,3000,2100,1"
          mvDescription = DataSelectionText.String17776    'Event Waiting List Bookings
          mvCode = "EWB"
          mvRequiredItems = "ContactName,BookingDate,Quantity"

        Case DataSelectionTypes.dstEVWaitingDelegates
          mvResultColumns = "BookingNumber," & ContactNameResults() & ",ContactNumber"
          mvSelectColumns = "BookingNumber,ContactName,ContactNumber"
          mvHeadings = DataSelectionText.String17777     'Booking Number,Delegate,ContactNumber
          mvWidths = "1,3000,1"
          mvDescription = DataSelectionText.String17779     'Event Waiting List Delegates
          mvCode = "EWD"
          mvRequiredItems = "ContactName"

        Case DataSelectionTypes.dstEventSources
          mvResultColumns = "Source,SourceDesc,AmendedBy,AmendedOn,HistoryOnly"
          mvSelectColumns = "Source,SourceDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18266     'Source,Description,Amended By,Amended On
          mvWidths = "800,3000,1200,1200"
          mvDescription = DataSelectionText.String22899     'Event Sources
          mvCode = "EVSO"
          mvRequiredItems = "Source"

        Case DataSelectionTypes.dstEventMailings
          mvResultColumns = "Mailing,MailingDesc,Department,DepartmentDesc,HistoryOnly,Marketing,Direction,MailingDue,MailingDate,NumberInMailing,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "Mailing,MailingDesc,Department,DepartmentDesc,HistoryOnly,Marketing,Direction,MailingDue,MailingDate,NumberInMailing,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18268     'Mailing,Description,Dept,Department,History Only,Marketing,Direction,Mailing Due,Mailing Date,Number In Mailing,Notes,Amended By,Amended On
          mvWidths = "800,3000,1,2000,800,800,800,1200,1200,1200,2000,1200,1200"
          mvDescription = DataSelectionText.String22900    'Event Mailings
          mvCode = "EVMA"
          mvRequiredItems = "Mailing"

        Case DataSelectionTypes.dstEventVenueData
          mvResultColumns = "Venue,VenueDesc,VenueCapacity,Location,VenueOrganisationNumber,VenueOrganisationName,VenueAddressNumber,VenueAddress,VenueContactNumber,VenueContactName,VenueTelephone"
          mvSelectColumns = "Venue,VenueDesc,VenueCapacity,Location,VenueOrganisationNumber,VenueOrganisationName,VenueAddressNumber,VenueAddress,VenueContactNumber,VenueContactName,VenueTelephone"

        Case DataSelectionTypes.dstEventOrganiserData
          mvResultColumns = "Organiser,OrganiserDesc,OrganiserContactNumber,OrganiserContactName,OrganiserContactAddressNumber,OrganiserContactAddressLine,InvoiceContactNumber,InvoiceContactName,InvoiceContactAddressNumber,InvoiceContactAddressLine"
          mvSelectColumns = "Organiser,OrganiserDesc,OrganiserContactNumber,OrganiserContactName,OrganiserContactAddressNumber,OrganiserContactAddressLine,InvoiceContactNumber,InvoiceContactName,InvoiceContactAddressNumber,InvoiceContactAddressLine"

        Case DataSelectionTypes.dstEventHeaderInfo
          Dim vEvent As New CDBEvent(mvEnv)
          mvResultColumns = vEvent.DataTableColumns
          mvSelectColumns = "EventDesc,StartDate,NewColumn,EventNumber,VenueDesc"
          mvHeadings = ",,,No:,Venue:"
          mvWidths = "300,300,300,300,300"    'For the header these are really heights with 300 per line
          mvHeaderLines = 0
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18270    'Event Header Info    
          mvCode = "EHI"
          mvRequiredItems = "EventNumber,GroupRGBValue"
          vPrimaryList = True

        Case DataSelectionTypes.dstEventBookingDelegates
          mvResultColumns = "AddressNumber,Position,OrganisationName,EventDelegateNumber,SequenceNumber,Title,Forenames,Surname,ContactNumber,ContactName"
          mvSelectColumns = "ContactNumber,AddressNumber,ContactName,Position,OrganisationName"
          mvWidths = "1,1,3250,2000,2000"
          mvHeadings = DataSelectionText.String26506     'Contact Number,Address Number,Name,Position,Organisation
          mvDescription = "Event Booking Delegates"
          mvCode = "EBD"

        Case DataSelectionTypes.dstEventRoomBookingAllocations
          mvResultColumns = "AddressNumber,RoomDate,RoomID,RoomBookingLinkNumber,ContactNumber,Position,OrganisationName,ContactName"
          mvSelectColumns = "ContactNumber,AddressNumber,ContactName,RoomDate,RoomID,RoomBookingLinkNumber"
          mvWidths = "1,1,4500,1500,1,1"
          mvHeadings = DataSelectionText.String26510     'Contact Number,Address Number,Name,Date,Room ID,Link Number

        Case DataSelectionTypes.dstContactNotifications
          mvResultColumns = "ItemNumber,LinkType,ItemDate,ItemDescription,ItemType,Subject,ItemCode,ItemDesc,Access"
          mvSelectColumns = "ItemNumber,ItemDesc,ItemDate,ItemDescription,ItemType,ItemCode,LinkType,Subject"
          mvHeadings = DataSelectionText.String17780     'No,Item,Date,Description,Type,Code,Link,Subject
          mvWidths = "1000,2025,1200,2500,2000,1,1,1"

        Case DataSelectionTypes.dstSelectionSteps
          mvResultColumns = "SequenceNumber,SelectActionDesc,ViewDesc,FilterSQL,RecordCount,SelectAction,ViewName,CriteriaSetDesc"
          mvSelectColumns = "SequenceNumber,SelectActionDesc,ViewDesc,FilterSQL,RecordCount,SelectAction,ViewName,CriteriaSetDesc"
          mvHeadings = DataSelectionText.String17781     'Step,Action,Using,Where,Count,SelectAction,ViewName,Description
          mvWidths = "1200,1200,2000,4000,1200,1,1,1"

        Case DataSelectionTypes.dstSelectItemAddresses
          mvResultColumns = "AddressUsage,Notes"
          mvSelectColumns = "AddressUsage,Notes"
          mvHeadings = DataSelectionText.String19019 ' Address Usage,Notes
          mvWidths = "2000,5350,1,1,1,1,1"
          mvRequiredItems = "AddressUsage,Notes"

        Case DataSelectionTypes.dstSelectItemSelectionSets
          mvResultColumns = "SetNumber,Description,Department,Owner,Count,Auto,Custom"
          mvSelectColumns = "SetNumber,Description,Department,Owner,Count,Auto,Custom"
          mvHeadings = DataSelectionText.String17782     'Set No,Description,Dept,Owner,Count,Auto,Custom
          mvWidths = "1,4100,1000,900,700,700,700"
          mvRequiredItems = "Description,Owner,Count,Auto,Custom"

        Case DataSelectionTypes.dstSelectItemCreditAccount
          mvResultColumns = "SalesLedgerAccount,Company,CreditCategory,StopCode,CustomerType"
          mvSelectColumns = "SalesLedgerAccount,Company,CreditCategory,StopCode,CustomerType"
          mvHeadings = DataSelectionText.String19014 'Sales Ledger Account,Company,Credit Category,Stop Code,Customer Type
          mvWidths = "2050,1100,1600,1100,1800"

        Case DataSelectionTypes.dstContactHeaderInfo
          mvResultColumns = "ContactName,Position,AddressLine,NewColumn,ContactNumber,PhoneNumber,StatusDesc,NewColumn2,SourceDesc,SourceDate,Sex,DateOfBirth,DOBEstimated,MarketingChart,Picture,CommunicationsList,HighProfileActivitiesList,DepartmentActivitiesList,HighProfileLinksList,OwnershipGroupDesc,Notes,AddressMultiLine,OwnershipAccessLevel,StickyNoteCount,GroupCode,ContactType,AddressNumber,GroupRGBValue,ActionCount,Department,DepartmentDesc,Status,BranchCode,PreferredCommunication,WebAddress,AddressType,DefaultContactName,CurrentPosition,CurrentAddressLine,CurrentAddressMultiLine,RgbStatus"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIrishGiftAid) Then mvResultColumns = mvResultColumns & ",NINumber"
          If pUsage = DataSelectionUsages.dsuWEBServices Then mvResultColumns = Replace$(mvResultColumns, "Picture,", "")
          If pUsage = DataSelectionUsages.dsuWEBServices Then mvResultColumns = Replace$(mvResultColumns, "MarketingChart,", "")
          mvResultColumns = mvResultColumns & ",StatusGroup,StatusGroupDesc,LabelName"
          mvSelectColumns = "ContactName,Position,AddressLine,NewColumn,ContactNumber,StatusDesc,SourceDesc,NewColumn2,CommunicationsList"
          mvHeadings = ",,,,Contact No:,Status:,Source:,,"
          mvWidths = "300,300,300,300,300,300,1200,1200,900"    'For the header these are really heights with 300 per line
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17784    'Contact Header Info
          mvCode = "CHI"
          mvRequiredItems = "ContactNumber,ContactName,OwnershipAccessLevel,StickyNoteCount,GroupCode,ContactType,AddressNumber,GroupRGBValue,ActionCount,Status,StatusDesc,BranchCode,DateOfBirth,DOBEstimated,AddressType,Source,NameGatheringSource,RgbStatus,ContactReference,AddressLine,StatusReason"

        Case DataSelectionTypes.dstContactActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,AmendedBy,AmendedOn,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,LinkType,LinkTypeDesc,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText,OutlookId"
          mvSelectColumns = "LinkTypeDesc,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvWidths = "1200,800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600"
          mvHeadings = DataSelectionText.String17785     'Type,Number,Description,Priority,Status,Action Text,,Deadline,Created By,,Scheduled On,Created On,,Completed On
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String17786   'Contact Actions
          mvDisplayTitle = DataSelectionText.String17786     'Contact Actions
          mvMaintenanceDesc = "Action"
          mvCode = "CAC"
          mvRequiredItems = "ActionStatus,MasterAction"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactAddresses
          mvResultColumns = "ParentNumber,AddressNumber,AddressType,HouseName,Address,Town,County,Postcode,CountryCode,ValidFrom,ValidTo,Historical,Branch,PAF,AmendedBy,AmendedOn,SortCode,UK,CountryDesc,AddressAmendedBy,AddressAmendedOn,GovernmentRegionDesc,BuildingNumber,DeliveryPointSuffix,LeaCode,LeaName,OrganisationNumber,Name"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCountryAddressFormat) Then mvResultColumns = mvResultColumns & ",address_format"
          mvResultColumns = mvResultColumns & ",AddressLine1,AddressLine2,AddressLine3"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAddressConfirmed) Then mvResultColumns = mvResultColumns & ",AddressConfirmed"
          mvResultColumns = mvResultColumns & ",TownAddressLine,Default,AddressLine"  'ALways keep these at end as these will not be read from the main sql
          mvSelectColumns = "ParentNumber,Town,AddressLine,Default,Historical,PAF,DetailItems,ValidFrom,ValidTo,NewColumn,Branch,GovernmentRegionDesc,NewColumn2,AmendedBy,AddressAmendedBy,NewColumn3,AmendedOn,AddressAmendedOn"
          mvHeadings = DataSelectionText.String17787     'Town,Address,Default,Historical,PAF Status,,Known from,Known to,,Branch,Government Region,,Address Link amended by,Address Record amended by,,on,on
          mvWidths = "1,1500,1200,900,900,900,1200,1150,1150,1200,900,1500,1200,1200,1600,1200,1150,1600"
          mvDescription = DataSelectionText.String17788    'Contact Addresses
          mvMaintenanceDesc = "Address"
          mvCode = "CA"
          mvRequiredItems = "AddressLine,Default,Historical"        'Default and Historical required for the portal
          vPrimaryList = True

        Case DataSelectionTypes.dstContactAddressUsages
          mvResultColumns = "AddressNumber,AddressUsage,AddressUsageDesc,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "AddressUsageDesc,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17789     'Usage,Notes,Amended By,Amended On
          mvWidths = "2000,3000,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String17790    'Contact Address Usages
          mvDisplayTitle = DataSelectionText.String17791     'Address Usages
          mvMaintenanceDesc = "Address Usage"
          mvCode = "CAU"
          mvRequiredItems = "AddressUsage"

        Case DataSelectionTypes.dstContactAddressesWithUsages
          Dim vDS As New DataSelection(mvEnv, DataSelectionTypes.dstContactAddresses, pParams, pListType, pUsage, pGroup)
          mvResultColumns = vDS.mvResultColumns.Replace(",TownAddressLine,", ",AddressUsage,TownAddressLine,") & ",AddressMultiLine"
          mvSelectColumns = "AddressNumber,AddressMultiLine,AddressUsage"
          mvHeadings = "Address Number,Address,Address Usage"
          mvWidths = "1200,1200,1200"

        Case DataSelectionTypes.dstContactAddressAndUsage
          mvResultColumns = "AddressNumber,AddressLine,Historical,AddressUsage,AddressUsageDesc"
          mvSelectColumns = "AddressNumber,AddressLine,Historical,AddressUsage,AddressUsageDesc"
          mvHeadings = DataSelectionText.String40011     'Number,Address,Hist,Usage
          mvWidths = "200,4100,200,400,1500"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String40012    'Contact Address And Usages
          mvCode = "CUA"
          vPrimaryList = True
          mvRequiredItems = "AddressLine,Historical,AddressUsage,AddressUsageDesc"

          'selection_set,selection_set_desc,user_name,department,number_in_set,source
        Case DataSelectionTypes.dstGeneralMailingSelectionSets
          mvResultColumns = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source,CustomData,AttributeCaptions"
          mvSelectColumns = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source,CustomData,AttributeCaptions"
          mvHeadings = DataSelectionText.String40013
          mvWidths = "200,800,4500,400,1500,1,1"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String40014 'General Mailing Selection Set
          mvCode = "GMS"
          vPrimaryList = True
          mvRequiredItems = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source,CustomData,AttributeCaptions"

        Case DataSelectionTypes.dstCriteriaSetDetails
          mvResultColumns = "CriteriaSet,SequenceNumber,SearchArea,IE,CO,MainValue,SubsidiaryValue,Period,Counted,AndOr,LeftParenthesis,RightParenthesis"
          mvSelectColumns = "AndOr,LeftParenthesis,IE,CO,SearchArea,MainValue,SubsidiaryValue,Period,RightParenthesis,Counted" '     CriteriaSet,SequenceNumber,SearchArea,IE,CO,MainValue,SubsidiaryValue,Period,Counted,AndOr,LeftParenthesis,RightParenthesis"
          mvHeadings = "And/Or,(,I/E,C/O,Area,Value,Sub Value,Period,),Counted" ' TODO: move to dataselectiontext      'And/Or,(,I/E,C/O,Area,Value,Sub Value,Period,),Counted
          mvWidths = "1200,1200,1200,2200,2200,2200,2200,2200,1200,2200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstContactCommunicationUsages
          mvResultColumns = "CommunicationNumber,CommunicationUsage,CommunicationUsageDesc,Notes,PrimaryUsage,AmendedBy,AmendedOn"
          mvSelectColumns = "CommunicationUsageDesc,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17789     'Usage,Notes,Amended By,Amended On
          mvWidths = "2000,3000,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18271    'Contact Communication Usages
          mvDisplayTitle = DataSelectionText.String18272     'Usages
          mvMaintenanceDesc = "Number Usage"
          mvCode = "CCU"
          vPrimaryList = True
          mvRequiredItems = "CommunicationUsage,PrimaryUsage"

        Case DataSelectionTypes.dstContactPictureDocuments
          mvResultColumns = "DocumentNumber,Extension"
          mvSelectColumns = "DocumentNumber,Extension"
          mvHeadings = "Document Number,Extension"
          mvWidths = "1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CPC"

        Case DataSelectionTypes.dstContactCategories
          'Same result columns as legacy assets - keep in synch
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,ResponseChannel,ResponseChannelDesc,IsActivityHistoric,IsActivityValueHistoric,NoteFlag,Status,Access,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity,Status,SourceDesc,ValidFrom,ValidTo,NoteFlag,DetailItems,Notes,AmendedBy,NewColumn,Spacer1,AmendedOn,NewColumn2,Spacer"
          mvHeadings = DataSelectionText.String17792     'Category,Value,Quantity,Status,Source,Valid from,Valid to,Notes?,,Notes,Amended by,,,on,,
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,Access,RgbActivityValue,Status,ActivityDesc,IsActivityHistoric,IsActivityValueHistoric,ContactCategoryNumber"
          mvDescription = DataSelectionText.String17793    'Contact Categories
          mvMaintenanceDesc = "Category"
          mvCode = "CC"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactPositionCategories
          mvResultColumns = "ContactPositionActivityId,ContactPositionNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityCode,ActivityDesc,ActivityValueCode,ActivityValueDesc,SourceCode,SourceDesc,Quantity,Status,NoteFlag,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes"
          mvHeadings = DataSelectionText.String23809 'Activity Code,Activity,Activity Value Code,Activity Value,Source Code,Source,Quantity,Status,Notes,Valid From,Valid To,Amended By,Amended On,Notes
          mvWidths = "1,1000,1,1000,1,1000,1000,1000,1000,1200,1200,1200,1200,1"
          mvDescription = DataSelectionText.String17794    'Contact Position Categories
          mvDisplayTitle = DataSelectionText.String17795    'Activities
          mvMaintenanceDesc = "Position Category"
          mvCode = "CPA"
          mvRequiredItems = "ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,Notes,ContactPositionActivityId"

        Case DataSelectionTypes.dstContactRoles
          mvResultColumns = "ContactNumber," & ContactNameResults() & ",OrganisationNumber,OrganisationName,RoleCode,RoleDesc,ValidFrom,ValidTo,Current,AmendedBy,AmendedOn,ContactRoleNumber"
          mvSelectColumns = "RoleDesc,Current,ValidFrom,ValidTo,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17797     'Description,Current,Valid from,Valid to,Amended by,Amended on
          mvWidths = "1800,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17798    'Contact Roles
          mvDisplayTitle = DataSelectionText.String17799     'Roles
          mvMaintenanceDesc = "Role"
          mvCode = "CR"
          vPrimaryList = True
          mvRequiredItems = "RoleCode,ContactName,OrganisationName"

        Case DataSelectionTypes.dstContactLinksTo
          'Also see dstContactLegacyLinks -keep Resultcolumns in synch
          mvResultColumns = "RelationshipCode,Type1,Type2,RelationshipDesc," & ContactNameResults() & ",ContactNumber,Phone,ValidFrom,ValidTo,Historical,Notes,AmendedBy,AmendedOn,OwnershipGroup,ContactGroup,Status,StatusDate,StatusReason,RelationshipStatus,RelationshipStatusDesc,RgbRelationshipStatus,ContactLinkNumber"
          mvSelectColumns = "RelationshipDesc,ContactName,Phone,Historical,ValidFrom,ValidTo,Status,StatusDate,StatusReason,DetailItems,Notes,AmendedBy,NewColumn,Spacer,AmendedOn"
          mvHeadings = DataSelectionText.String17800     'Relationship,With,Phone,Historical,Valid from,Valid to,Status,Status Date,Status Reason,,Notes,Amended by,,,on
          mvWidths = "1500,1500,1200,1000,1200,1200,1000,1200,2700,1200,900,600,1200,1200,1200"
          mvDescription = DataSelectionText.String17801    'Contact Links To
          mvMaintenanceDesc = "Relationship"
          mvCode = "CLT"
          mvRequiredItems = "RelationshipCode,Type1,Type2,ContactName,ValidFrom,ValidTo,Notes,ContactGroup,RgbRelationshipStatus,Historical,ContactLinkNumber" 'ContactNumber will be added by default.
          vPrimaryList = True

        Case DataSelectionTypes.dstContactLinksFrom
          mvResultColumns = "RelationshipCode,Type1,Type2," & ContactNameResults() & ",ContactNumber,RelationshipDesc,Phone,ValidFrom,ValidTo,Historical,Notes,AmendedBy,AmendedOn,OwnershipGroup,ContactGroup,EventDelegateNumber,EventDesc,BookingOptionDesc,RelationshipStatus,RelationshipStatusDesc,RgbRelationshipStatus,ContactLinkNumber"
          mvSelectColumns = "ContactName,ContactNumber,RelationshipDesc,Phone,Historical,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String23208     'From,No,Relationship,Phone,Historical,Amended By,Amended On
          mvWidths = "3605,1000,2160,1540,900,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuCare
          mvDescription = DataSelectionText.String17802    'Contact Links From
          mvMaintenanceDesc = "Relationship"
          mvCode = "CLF"
          vPrimaryList = True
          mvRequiredItems = "RelationshipCode,Type1,Type2,ContactName,ValidFrom,ValidTo,Notes,ContactGroup,RgbRelationshipStatus,Historical,ContactLinkNumber" 'ContactNumber will be added by default.

        Case DataSelectionTypes.dstContactSuppressions
          mvResultColumns = "SuppressionCode,SuppressionDesc,ValidFrom,ValidTo,Notes,SuppressionInformation,Source,SourceDesc,ResponseChannel,ResponseChannelDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "SuppressionCode,SuppressionDesc,ValidFrom,ValidTo,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17803     'Suppression,Description,Valid from,Valid to,Notes,Amended by,Amended on
          mvWidths = "1200,3000,1125,1200,2000,1200,1200"
          mvDescription = DataSelectionText.String17804    'Contact Suppressions
          mvMaintenanceDesc = "Suppression"
          mvCode = "CSP"
          mvRequiredItems = "SuppressionCode,Notes,AmendedOn,AmendedBy,SuppressionDesc"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactPositions
          mvResultColumns = "ContactNumber,AddressNumber,Position," & ContactNameResults() & ",ValidFrom,ValidTo,Mail,Current,Location,PositionFunction,PositionSeniority,AmendedBy,AmendedOn,ContactPositionNumber,SinglePosition,OrganisationGroup,ContactGroup,AddressLine,PositionFunctionDesc,PositionSeniorityDesc,AddressValidFrom,AddressValidTo,DiallingCode,STDCode,Number,ExDirectory,PhoneNumber,OrganisationDefault,PositionStatus,PositionStatusDesc"
          mvSelectColumns = "Position,ContactName,AddressLine,Current,Mail,DetailItems,ValidFrom,Location,NewColumn,ValidTo,NewColumn2,AmendedBy,NewColumn3,AmendedOn"
          mvHeadings = DataSelectionText.String17806     'Position,Name,Site Address,Current,Mail,,Valid from,Location,,to,,Amended by,,on
          mvWidths = "1800,2500,3600,900,900,300,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17807    'Contact Positions
          mvMaintenanceDesc = "Position"
          mvCode = "CP"
          vPrimaryList = True
          mvRequiredItems = "ContactName,Location,Position,ValidFrom,ValidTo,AddressValidFrom,AddressValidTo,SinglePosition,OrganisationGroup,ContactGroup,Current"

        Case DataSelectionTypes.dstContactAddressPositionAndOrg
          mvResultColumns = "Position,Name,OrganisationNumber,PositionLocation,Started,Finished,Current"
          mvSelectColumns = "Position,Name,OrganisationNumber"
          mvHeadings = DataSelectionText.String17808     'Position,Organisation,Number
          mvWidths = "1800,3000,1"
          mvCode = "CAP"
          vPrimaryList = True
          mvRequiredItems = "Position,Name,Current"

        Case DataSelectionTypes.dstContactPositionActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,AmendedBy,AmendedOn,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,LinkType,LinkTypeDesc,SortColumn,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText,OutlookId"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,Deadline,CreatedBy,ScheduledOn,CreatedOn,CompletedOn"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvHeadings = DataSelectionText.DstContactPositionActionsHeader  'Number,Description,Priority,Status,Action Text,Deadline,Created By,Scheduled On,Created On,Completed On
          mvDescription = DataSelectionText.String18701    'Position Roles
          mvDisplayTitle = DataSelectionText.String18702   'Actions
          mvMaintenanceDesc = DataSelectionText.String18703  'Action
          mvCode = "POSA"

        Case DataSelectionTypes.dstContactPositionDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,"
          mvResultColumns &= "StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc"
          mvResultColumns &= ",Subject,CallDuration,TotalDuration,SelectionSet"
          mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc"
          mvHeadings = DataSelectionText.String18709  'Document Number, Dated, In/Out, Subject, Reference, Document Type, Topic, Sub Topic
          mvWidths = "2000,2000,2000,2000,2000,2000,2000,2000"
          mvDescription = DataSelectionText.String18706    'Position Documents
          mvDisplayTitle = DataSelectionText.String18707   'Documents
          mvMaintenanceDesc = DataSelectionText.String18708  'Document
          mvCode = "POSD"

        Case DataSelectionTypes.dstContactPositionTimesheets
          mvResultColumns = "ContactPositionNumber,ContactTimesheetNumber,TimesheetDate,DurationHours,DurationMinutes,TimesheetDesc,Role,RoleDesc"
          mvSelectColumns = "TimesheetDate,DurationHours,DurationMinutes,TimesheetDesc,RoleDesc"
          mvHeadings = DataSelectionText.String18714  'Date,Hours,Minutes,Description,Role
          mvWidths = "1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String18711    'Position Timesheet
          mvDisplayTitle = DataSelectionText.DstContactPositionTimesheetDisplayTitle   'Timesheets
          mvMaintenanceDesc = DataSelectionText.String18712   'Timesheet
          mvCode = "POST"

        Case DataSelectionTypes.dstContactProcessedTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,TransactionTypeDesc,TransactionDate,Amount,PaymentMethodDesc,Reference,Posted,Status,TransactionSign,PaymentMethodCode,CurrencyAmount,BankDetailsNumber,Notes,TransactionTypeCode,ContainsStock,ContainsPostage,CurrencyCode,TransactionOrigin,TransactionOriginDesc,CanAddGiftAid,ContainsSalesLedgerItems,BankAccount,BankAccountDesc,RgbBankAccount,RgbAmount,RgbCurrencyAmount,AddressLine"
          mvSelectColumns = "TransactionDate,Amount,TransactionTypeDesc,PaymentMethodCode,Reference,BatchNumber,TransactionNumber,Posted,Status,DetailItems,Notes"
          mvHeadings = DataSelectionText.String17809     'Transaction Date,Amount,Type,Payment Method,Reference,Batch,Transaction,Posted Date,Status,,Notes
          mvWidths = "1488,900,1020,1200,1600,740,975,1530,600,1200,1200,1200"
          mvDescription = DataSelectionText.String17810    'Contact Processed Transactions
          mvCode = "CPT"
          vPrimaryList = True
          mvRequiredItems = "TransactionSign,PaymentMethodCode,ContainsStock,ContainsPostage,ContainsSalesLedgerItems,RgbBankAccount,RgbAmount,RgbCurrencyAmount"

        Case DataSelectionTypes.dstContactUnProcessedTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,Provisional,TransactionTypeDesc,TransactionDate,Amount,PaymentMethodDesc,Reference,Mailing,Receipt,EligibleForGiftAid,CurrencyAmount,Notes,TransactionTypeCode,PaymentMethodCode,CurrencyCode,TransactionOrigin,TransactionOriginDesc,BankAccount,BankAccountDesc,RgbBankAccount,RgbAmount,RgbCurrencyAmount,AddressLine"
          mvSelectColumns = "TransactionDate,Amount,TransactionTypeDesc,PaymentMethodCode,Reference,BatchNumber,TransactionNumber,Provisional,DetailItems,Receipt,Notes,NewColumn,Mailing,NewColumn2,EligibleForGiftAid,NewColumn3"
          mvHeadings = DataSelectionText.String17811     'Date,Amount,Type,Payment Method,Reference,Batch,Transaction,Provisional?,,Receipt?,Notes,,Mailing,,Eligible for Gift Aid?,
          mvWidths = "1100,900,1020,1200,1200,900,1095,1005,1200,600,600,1200,600,1200,600,1200"
          mvDescription = DataSelectionText.String17812     'Contact UnProcessed Transactions
          mvCode = "CUT"
          vPrimaryList = True
          mvRequiredItems = "TransactionDate,Amount,Provisional,PaymentMethodCode,RgbBankAccount,RgbAmount,RgbCurrencyAmount"

        Case DataSelectionTypes.dstContactCancelledProvisionalTrans
          mvResultColumns = "BatchNumber,TransactionNumber,TransactionTypeDesc,TransactionDate,Amount,PaymentMethodDesc,Reference,Mailing,Receipt,EligibleForGiftAid,CurrencyAmount,Notes,Provisional,TransactionTypeCode,PaymentMethodCode,CurrencyCode"
          mvSelectColumns = "TransactionDate,Amount,TransactionTypeDesc,PaymentMethodCode,Reference,BatchNumber,TransactionNumber,Provisional,DetailItems,Receipt,Notes,NewColumn,Mailing,NewColumn2,EligibleForGiftAid,NewColumn3"
          mvHeadings = DataSelectionText.String17811     'Date,Amount,Type,Payment Method,Reference,Batch,Transaction,Provisional?,,Receipt?,Notes,,Mailing,,Eligible for Gift Aid?,
          mvWidths = "1100,900,1020,1200,1200,900,1095,1005,1200,600,600,1200,600,1200,600,1200"
          mvDescription = DataSelectionText.String17816     'Contact Cancelled Provisional Transactions
          mvCode = "CCP"
          vPrimaryList = True
          mvRequiredItems = "TransactionDate,Amount"

        Case DataSelectionTypes.dstContactStatusHistory
          mvResultColumns = "ContactNumber,StatusCode,StatusReason,ValidTo,ResponseChannel,ResponseChannelDesc,AmendedBy,AmendedOn,StatusDesc,RgbStatus"
          mvSelectColumns = "StatusDesc,StatusReason,ValidTo,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String22606        'Status,Reason,Valid To,Amended By,Amended On  '2000,2000,1200,1200,1200
          mvWidths = "2000,2000,1200,1200,1200"
          mvDescription = DataSelectionText.String17817     'Contact Status History
          mvCode = "CSH"
          mvRequiredItems = "RgbStatus"

        Case DataSelectionTypes.dstContactEventBookings
          mvResultColumns = "EventReference,EventNumber,EventDesc,StartDate,BookingNumber,BookingDate,Quantity,BookingStatus,BookingOptionNumber,BookingOptionDesc,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,Venue,VenueDesc,EventLocation,EventClass,CancellationReason,CancelledBy,CancelledOn,AdultQuantity,ChildQuantity,StartTime,EndDate,EndTime,InvoiceNumber,InvoiceRePrintCount,BatchType,BookingAmount,Product,Rate,Notes,PayerContactNumber,PayerName,ContactNumber," & ContactNameResults() & ",BatchNumber,TransactionNumber,LineNumber,InvoicePayStatus,InvoicePayStatusDesc,InvoiceAllocationAmount,BookingStatusDesc,CreditSale,InvoicePrinted"
          mvSelectColumns = "EventNumber,EventReference,EventDesc,StartDate,BookingOptionDesc,BookingDate,BookingNumber,Quantity,BookingStatusDesc,DetailItems,SubjectDesc,VenueDesc,ContactName,BatchNumber,NewColumn,SkillLevelDesc,EventLocation,Spacer,TransactionNumber,NewColumn2,Spacer1,Spacer2,Spacer3,LineNumber"
          mvHeadings = DataSelectionText.String17819     'Event Number,Event Reference,Event,Start Date,Booking Option,Booked on,Booking Number,Quantity,Status,,Subject,Venue,Booked By,Batch,,Skill Level,Location,,Transaction,,,,,Line
          mvWidths = "1000,2000,4000,1200,3000,1200,1000,600,1500,1200,1200,1200,3000,800,1200,1200,1200,1200,800,1200,1200,1200,1200,800"
          mvDescription = DataSelectionText.String17820     'Contact Event Bookings
          mvCode = "CEB"
          vPrimaryList = True
          mvRequiredItems = "CancellationReason,AdultQuantity,ChildQuantity,StartTime,EndTime,CreditSale,InvoicePrinted,BookingStatus,BookingAmount,Product,Rate,InvoiceAllocationAmount"

        Case DataSelectionTypes.dstContactEventDelegates
          mvResultColumns = "EventReference,EventNumber,EventDesc,StartDate,BookingNumber,Attended,BookingStatus,CandidateNumber,BookingOptionNumber,BookingOptionDesc,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,Venue,VenueDesc,EventLocation,EventClass,PledgedAmount,DonationTotal,SponsorshipTotal,BookingPaymentAmount,OtherPaymentsTotal,SequenceNumber,ContactName,ContactNumber,BatchNumber,TransactionNumber,LineNumber,TransactionSource,PayerContactNumber,PayerName,DelegateNumber,DelegateContactNumber,DelegateName,Surname,Forenames,Title,Initials,BookingStatusDesc"
          mvSelectColumns = "EventNumber,EventReference,EventDesc,StartDate,BookingOptionDesc,BookingNumber,BookingStatusDesc,Attended,CandidateNumber,DetailItems,SubjectDesc,VenueDesc,ContactName,NewColumn,SkillLevelDesc,EventLocation,DelegateName"
          mvHeadings = DataSelectionText.String17822     'Event Number,Event Reference,Event,Start Date,Booking Option,Booking Number,Status,Attended,Candidate,,Subject,Venue,Booked by,,Skill Level,Location,Delegate
          mvWidths = "1000,2000,2385,1200,1965,1000,1500,900,1200,1200,1200,1200,3000,1200,1200,1200,3000"
          mvDescription = DataSelectionText.String17823     'Contact Event Delegates
          mvCode = "CED"
          vPrimaryList = True
          mvRequiredItems = "TransactionSource"

        Case DataSelectionTypes.dstContactEventSessions
          mvResultColumns = "EventReference,EventNumber,BookingNumber,SessionNumber,EventDesc,SessionDesc,StartDate,StartTime,EndDate,EndTime,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,Venue,VenueDesc,EventLocation,EventClass,Attended,DelegateNumber,DelegateName"
          If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDelegateSessions) Then mvResultColumns = Replace(mvResultColumns, "Attended,", "")
          mvSelectColumns = "EventNumber,EventReference,EventDesc,SessionDesc,StartDate,StartTime,DelegateName,DetailItems,SubjectDesc,VenueDesc,NewColumn,SkillLevelDesc,EventLocation"
          mvHeadings = DataSelectionText.String17825     'Event Number,Event Reference,Event,Session,Start Date,Start Time,Delegate,,Subject,Venue,,Skill Level,Location
          mvWidths = "1000,2000,3500,3500,1200,1200,3000,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17826     'Contact Event Sessions
          mvCode = "CES"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactEventRoomBookings
          mvResultColumns = "EventReference,EventNumber,EventDesc,StartDate,RoomBookingNumber,BookedDate,Quantity,RoomTypeDesc,FromDate,ToDate,BookingStatus,LocationName,ConfirmedOn,Notes,CancellationReason,EventClass,ContactNumber," & ContactNameResults() & ",BatchNumber,TransactionNumber,LineNumber,BookingStatusDesc"
          mvSelectColumns = "EventNumber,EventReference,EventDesc,StartDate,RoomBookingNumber,Quantity,RoomTypeDesc,FromDate,ToDate,BookingStatusDesc,DetailItems,ContactName,LocationName,BatchNumber,Notes,NewColumn,BookedDate,Spacer,TransactionNumber,NewColumn2,ConfirmedOn,Spacer1,LineNumber"
          mvHeadings = DataSelectionText.String17828     'Event Number,Event Reference,Event,Start Date,Booking Number,Quantity,Room Type,Booked+,from/to,Status,,Booked by,Location Name,Batch,Notes,,Booked on,,Transaction,,Confirmed on,,Line
          mvWidths = "1000,2000,3000,1200,1000,600,1500,1200,1200,1500,1200,3000,1200,800,1200,1200,1200,1200,800,1200,1200,1200,800"
          mvDescription = DataSelectionText.String17829     'Contact Room Bookings
          mvCode = "CRB"
          vPrimaryList = True
          mvRequiredItems = "CancellationReason"

        Case DataSelectionTypes.dstContactEventRoomsAllocated
          mvResultColumns = "EventReference,EventNumber,RoomID,EventDesc,StartDate,RoomBookingNumber,RoomTypeDesc,RoomDate,LocationName,Notes,EventClass,ContactNumber," & ContactNameResults()
          mvSelectColumns = "EventNumber,EventReference,EventDesc,StartDate,RoomBookingNumber,RoomTypeDesc,RoomDate,DetailItems,ContactName,LocationName,Notes"
          mvHeadings = DataSelectionText.String17831     'Event Number,Event Reference,Event,Start Date,Booking Number,Room Type,Booked From,,Occupier,Location Name,Notes
          mvWidths = "1000,2000,3000,1200,1000,1500,1200,1200,3000,1200,1200"
          mvDescription = DataSelectionText.String17832     'Contact Rooms Allocated
          mvCode = "CRA"
          vPrimaryList = True
          mvRequiredItems = "RoomID"

        Case DataSelectionTypes.dstContactEventOrganiser
          mvResultColumns = "EventReference,EventNumber,EventDesc,StartDate,Organiser,OrganiserDesc,Notes,Venue,VenueDesc,OrganiserReference,Location,EventClass,ContactNumber," & ContactNameResults()
          mvSelectColumns = "EventNumber,EventReference,EventDesc,StartDate,OrganiserDesc,DetailItems,ContactName,VenueDesc,Notes,NewColumn,OrganiserReference,Location"
          mvHeadings = DataSelectionText.String17834     'Event Number,Event Reference,Event,Start Date,Organiser,,Organiser Contact,Venue,Notes,,Organiser Reference,Location
          mvWidths = "1000,2000,3690,1200,1200,1200,3000,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17835     'Contact Event Organiser
          mvCode = "CEO"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactEventPersonnel
          mvResultColumns = "EventReference,EventNumber,SessionNumber,EventDesc,SessionDesc,StartDate,Task,EndDate,StartTime,EndTime,Venue,VenueDesc,Location,EventClass,ContactNumber," & ContactNameResults()
          mvSelectColumns = "EventNumber,EventReference,EventDesc,SessionDesc,StartDate,StartTime,Task,DetailItems,VenueDesc,EndDate,NewColumn,Location,EndTime"
          mvHeadings = DataSelectionText.String17837     'Event Number,Event Reference,Event,Session,Start Date+,& Time,Task,,Venue,End Date,,Location,End Time
          mvWidths = "1000,2000,3000,3000,1200,1200,1200,1200,2400,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17838     'Contact Event Personnel
          mvCode = "CEP"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactPurchaseOrders
          mvResultColumns = "PayeeContactName,PayeeContactNumber,PurchaseOrderNumber,PurchaseOrderTypeDesc,PurchaseOrderDesc,Amount,Balance,StartDate,NumberOfPayments,OutputGroup,CancellationReason,CancellationSource,CancelledBy,CancelledOn,AdHocPayments,RegularPayments,PaymentSchedule,RequiresAuthorisation,AuthorisationLevel,AuthorisationLevelDesc,AuthorisedBy,AuthorisedOn,CancellationReasonDesc,CancellationSourceDesc,CanAuthorise,CurrencyCode,CurrencyCodeDesc,HasInvoice,POPaymentTypeRequired"
          mvSelectColumns = "PurchaseOrderNumber,PurchaseOrderTypeDesc,PurchaseOrderDesc,Amount,Balance,StartDate,NumberOfPayments,OutputGroup,CancelledOn,DetailItems,PayeeContactName,CancelledBy,NewColumn,Spacer,CancellationReason"
          mvHeadings = DataSelectionText.String17841     'PO Number,Type,Description,Amount,Balance,Start Date,Number of Payments,Output Group,Cancelled on,,Payee,Cancelled by,,,Reason
          mvWidths = "900,1200,2000,900,1200,1200,1000,900,1200,1200,2000,1200,1200,1200,1000"
          mvDescription = DataSelectionText.String17842     'Contact Purchase Orders
          mvCode = "CPO"
          vPrimaryList = True
          mvRequiredItems = "CancelledBy,CancellationReason,AdHocPayments,PaymentSchedule,RegularPayments,RequiresAuthorisation,AuthorisedBy,CanAuthorise,CurrencyCode,HasInvoice,POPaymentTypeRequired"

        Case DataSelectionTypes.dstContactPurchaseInvoices
          mvResultColumns = "PayeeContactName,PayeeContactNumber,PurchaseInvoiceNumber,PurchaseOrderNumber,Amount,PurchaseInvoiceDate,PayeeReference,ChequeReferenceNumber,SortCode,AccountNumber,BacsProcessed,BacsProcessedDesc,CurrencyCode,CurrencyCodeDesc,IbanNumber,BicCode"
          mvSelectColumns = "PurchaseInvoiceNumber,PurchaseOrderNumber,PayeeContactNumber,PayeeContactName,Amount,PurchaseInvoiceDate,PayeeReference,ChequeReferenceNumber"
          mvHeadings = DataSelectionText.String17843     'Number,PO Number,Payee No,Payee,Amount,Date,Reference,Cheque Reference
          mvWidths = "1000,1000,1,2000,900,1200,1500,1500"
          mvDescription = DataSelectionText.String17845     'Contact Purchase Invoices
          vPrimaryList = True
          mvCode = "CPI"

        Case DataSelectionTypes.dstContactCommsNumbers
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,DiallingCode,STDCode,Extension,ExDirectory,Notes,AmendedBy,AmendedOn,ValidFrom,ValidTo,IsActive,Mail,DeviceDefault,PreferredMethod,CommunicationUsage,CommunicationUsageDesc,IsOrganisation,CommunicationNumber,Number,SubscriptionCount,Email,WwwAddress,Archive,PhoneNumber,AddressLine,Default"
          mvSelectColumns = "DeviceDesc,PhoneNumber,Extension,ExDirectory,Default,IsActive,Mail,PreferredMethod,DeviceDefault,IsOrganisation,DetailItems,ValidFrom,Notes,AmendedBy,NewColumn,ValidTo,Spacer,AmendedOn"
          mvHeadings = DataSelectionText.String17847      'Device,Number,Extension,Ex Directory?,Default,Current,Mail,Preferred,Device default,Is Organisation?,,Valid from,Notes,Amended by,,Valid to,,on
          mvWidths = "2000,2500,900,1200,900,900,900,900,900,900,1200,1200,2000,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17848     'Contact Numbers
          mvMaintenanceDesc = "Number"
          mvCode = "CN"
          vPrimaryList = True
          mvRequiredItems = "DiallingCode,STDCode,DeviceCode,Default,SubscriptionCount,Email,WwwAddress,IsOrganisation,IsActive"

        Case DataSelectionTypes.dstContactCommsNumbersWithUsages
          Dim vDS As New DataSelection(mvEnv, DataSelectionTypes.dstContactCommsNumbers, pParams, pListType, pUsage, pGroup)
          mvResultColumns = vDS.mvResultColumns
          If Not vDS.mvResultColumns.Contains("CommunicationUsage") Then mvResultColumns = vDS.mvResultColumns.Replace(",PhoneNumber,", ",CommunicationUsage,PhoneNumber,")
          mvSelectColumns = "PhoneNumber,CommunicationUsage"
          mvHeadings = "Phone Number,Communication Usage"
          mvWidths = "1200,1200"

        Case DataSelectionTypes.dstContactEMailAddresses
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,Notes,AmendedBy,AmendedOn,CommunicationNumber,EMailAddress,AutoEmail,ValidFrom,ValidTo,IsActive,Mail,DeviceDefault,PreferredMethod"
          mvSelectColumns = "DeviceDesc,EMailAddress,Notes"
          mvHeadings = DataSelectionText.String17849     'Device,EMail Address,Notes
          mvWidths = "900,2000,3000"
          mvDescription = DataSelectionText.String17850     'Contact EMail Addresses
          mvCode = "CE"
          mvRequiredItems = "DeviceCode"

        Case DataSelectionTypes.dstContactCommsNumbersEdit
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,DiallingCode,STDCode,Extension,ExDirectory,Notes,AmendedBy,AmendedOn,CommunicationNumber,Number,IsActive,Mail,PreferredMethod,DeviceDefault,ValidFrom,ValidTo,PhoneNumber,AddressLine,Default"
          mvSelectColumns = "Default,DeviceDesc,PhoneNumber,Extension,ExDirectory,Notes,IsActive,Mail,PreferredMethod,DeviceDefault"
          mvHeadings = DataSelectionText.String17851     'Default,Device Type,Details,Extension,Ex-Directory,Notes,Current,Mail,Preferred,Device Default
          mvWidths = "700,2600,2600,1000,1100,1100,700,700,800,1200"
          mvRequiredItems = "DiallingCode,STDCode,DeviceCode,IsActive,Mail,PreferredMethod,DeviceDefault,ValidFrom,ValidTo"

        Case DataSelectionTypes.dstContactHeaderCommsNumbers
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,DiallingCode,STDCode,Extension,ExDirectory,Notes,AmendedBy,AmendedOn,ValidFrom,ValidTo,IsActive,Mail,DeviceDefault,PreferredMethod,CommunicationNumber,Number,PhoneNumber"
          mvSelectColumns = "DeviceDesc,PhoneNumber"
          mvHeadings = DataSelectionText.String17852     'Device,Details
          mvWidths = "2000,3000"
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17854     'Contact Header Numbers
          mvCode = "CHN"

        Case DataSelectionTypes.dstOrganisationContactCommsNumbers
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,DiallingCode,STDCode,Extension,ExDirectory,Notes,AmendedBy,AmendedOn,ValidFrom,ValidTo,IsActive,Mail,DeviceDefault,PreferredMethod,CommunicationNumber,Number,PhoneNumber," & ContactNameResults()
          mvSelectColumns = "ContactName,PhoneNumber,DeviceDesc,Notes"
          mvHeadings = DataSelectionText.String17855     'Contact Name,Phone Number,Device,Notes
          mvWidths = "2000,2000,2000,3000"
          mvDescription = DataSelectionText.String17856    'Organisation Numbers
          mvHeaderLines = 0
          mvCode = "ON"
          mvRequiredItems = "DiallingCode,STDCode,DeviceCode"

        Case DataSelectionTypes.dstContactMailings
          mvResultColumns = "MailingNumber,ContactNumber,AddressNumber,Date,Direction,Mailing,Description,MailingTemplate,MailingTemplateDesc,Notes,MailingHistoryNotes,MailedBy,MailingFilename,Topic,TopicDesc,SubTopic,SubTopicDesc,Subject,FulfillmentNumber,ProcessedOn,ProcessedStatus,ErrorNumber,CommunicationNumber," & ContactNameResults() & ",CheetahMailId,NumberEmailsBounced,NumberEmailsOpened,NumberEmailsClicked,OpenedOn,Type,AddressLine,OrganisationName"
          mvSelectColumns = "MailingNumber,Date,Type,Mailing,Description,MailedBy,ContactName,ContactNumber"
          mvHeadings = DataSelectionText.String17857     'Number,Date,Type,Mailing,Description,Mailed By,Contact,Contact No.
          mvWidths = "1200,1200,1000,1500,2000,1200,1500,1"
          mvDescription = DataSelectionText.String17858    'Contact Mailings
          mvCode = "CM"
          vPrimaryList = True
          mvRequiredItems = "Type,MailingFilename"

        Case DataSelectionTypes.dstContactExternalReferences
          mvResultColumns = "DataSource,DataSourceDesc,ExternalReference,AmendedOn,AmendedBy"
          mvSelectColumns = "DataSource,DataSourceDesc,ExternalReference,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17860     'Data Source,Description,External Reference,Amended by,Amended on
          mvWidths = "1200,3000,2000,1200,1200"
          mvDescription = DataSelectionText.String17861    'Contact External References
          mvMaintenanceDesc = "External Reference"
          mvRequiredItems = "DataSource,ExternalReference"
          mvCode = "CER"

        Case DataSelectionTypes.dstContactDepartmentNotes
          mvResultColumns = "Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "AmendedBy,AmendedOn,DetailItems,Notes"
          mvHeadings = DataSelectionText.String17862     'Notes,Amended by,Amended on
          mvWidths = "1200,1200,1,1200"
          mvDescription = DataSelectionText.String17863    'Contact Department Notes
          mvMaintenanceDesc = "Notes"
          mvCode = "CDN"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactMemberships
          mvResultColumns = "MembershipNumber,PayerContactNumber,MemberNumber,MembershipTypeDesc,CancelledOn,RenewalDate,RenewalAmount,PaymentMethodDesc,PaymentFrequencyDesc,PayerLabelName,Balance,Joined,PaymentPlanNumber,AmendedBy,AmendedOn,MembersPerOrder,MembershipCard,AddressNumber,AgeOverride,MembershipCardExpires,ReprintCard,NumberOfMembers,BranchMember,Branch,BranchDesc,Applied,Accepted,SourceCode,VotingRights,CancelledBy,CancellationReason,CancellationSource,SourceDesc,FutureMembershipType,FutureMembershipTypeDesc,FutureChangeDate,FutureProduct,FutureRate,FutureAmount,AddressLine,SingleMembership"
          mvSelectColumns = "MembershipNumber,PayerContactNumber,MemberNumber,MembershipTypeDesc,CancelledOn,RenewalDate,RenewalAmount,PaymentMethodDesc,PaymentFrequencyDesc,PayerLabelName,Balance,Joined,PaymentPlanNumber,AmendedBy,AmendedOn,MembersPerOrder,MembershipCard,AddressNumber"
          mvHeadings = DataSelectionText.String23703     'Membership No,Pay Contact No,Number,Membership Type,Cancelled,Renewal,Amount,Method,Frequency,Payer,Balance,Joined,Payment Plan,Amended By,Amended On,Members Per Order,Membership Card,Address Number
          mvWidths = "1,1,900,3000,1100,1100,900,1500,1500,2000,900,1200,1400,1200,1200,1,1,1000"
          mvAvailableUsages = DataSelectionUsages.dsuCare
          mvDescription = DataSelectionText.String17864     'Contact Memberships
          mvCode = "CMEM"
          vPrimaryList = True
          mvRequiredItems = "MembersPerOrder,MembershipCard,CancelledOn,SingleMembership,CancellationReason,CancellationSource,FutureMembershipType"

        Case DataSelectionTypes.dstContactMembershipDetails
          mvResultColumns = "MembershipNumber,PayerContactNumber,MemberNumber,MembershipTypeDesc,CancelledOn,RenewalDate,RenewalAmount,PaymentMethodDesc,PaymentFrequencyDesc,PayerLabelName,Balance,Joined,PaymentPlanNumber,AmendedBy,AmendedOn,MembersPerOrder,MembershipCard,AddressNumber,AgeOverride,MembershipCardExpires,ReprintCard,NumberOfMembers,BranchMember,Branch,BranchDesc,Applied,Accepted,SourceCode,VotingRights,CancelledBy,CancellationReason,CancellationSource"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataRgbValueForMemberType) Then mvResultColumns = mvResultColumns & ",RgbMembershipType"
          mvResultColumns = mvResultColumns & ",StartDate,PaymentMethod,PaymentFrequency,CreatedBy,CreatedOn,GiverContactNumber,Provisional,NextPaymentDue,LastPaymentDate,Arrears,LastPaymentAmount,Term,InAdvance,TermUnits,FrequencyAmount,ExpiryDate,RenewalPending,NumberOfReminders,RenewalDateChangeReason,RenewalDateChangedBy,RenewalDateChangedOn,RenewalDateChangeValue,FutureCancellationReason,FutureCancellationDate,FutureCancellationSource,Amount,DirectDebitStatus,CreditCardStatus,StandingOrderStatus,CovenantStatus,TheirReference,EligibleForGiftAid,PaymentFrequencyFrequency"
          mvResultColumns &= ",ApprovalMembership,MembershipType,SubsequentMembershipType,PaymentPlanMembershipType,PayerRequired,SubsequentMembershipTypeChangeDate,Annual,MembershipTerm,ContactNumber,BranchMembership"
          mvResultColumns = mvResultColumns & ",MembershipCardIssueNumber,MembershipStatus,MembershipStatusDesc,RgbMembershipStatus,LockBranch"
          mvResultColumns = mvResultColumns & ",SourceDesc,FutureMembershipType,FutureMembershipTypeDesc,FutureChangeDate,FutureProduct,FutureRate,FutureAmount,"
          mvResultColumns = mvResultColumns & "AddressLine,TermDesc,NextPaymentAmount,CancellationReasonDesc,CancellationSourceDesc,FutureCancellationReasonDesc,FutureCancellationSourceDesc,GiftFrom,GiftTo,GiftMessage"

          mvSelectColumns = "MemberNumber,MembershipTypeDesc,RenewalDate,CancelledOn,Term,TermUnits,SourceDesc,AmendedBy,AmendedOn,DetailItems,Spacer,Spacer1,Spacer5,Spacer8,Spacer9,Spacer17,Spacer10,Spacer13,Spacer14,Spacer22,Spacer23,NewColumn,Joined,MembershipCardExpires,RenewalPending,Balance,PaymentMethodDesc,PaymentFrequencyDesc,CancellationReasonDesc,BranchMember,BranchDesc,FutureMembershipTypeDesc,FutureCancellationReasonDesc,NewColumn2,StartDate,ReprintCard,RenewalAmount,Arrears,LastPaymentAmount,FrequencyAmount,CancellationSourceDesc,Applied,Spacer16,Spacer20,FutureCancellationSourceDesc,NewColumn3,VotingRights,NumberOfMembers,NumberOfReminders,InAdvance,LastPaymentDate,NextPaymentDue,CancelledBy,CancellationReason,CancellationSource,Accepted,Spacer19,FutureChangeDate,FutureCancellationDate"
          mvHeadings = DataSelectionText.String18689     'Member Number,Membership Type,Renewal Date,Cancellation Date,Term,Units,Source,Amended by,on,,GENERAL,.,RENEWAL,FINANCIAL,.,.,CANCELLATION,BRANCH,.,FUTURE CHANGES,.,,Joined,Card Expires,Pending?,
          mvHeadings = mvHeadings & DataSelectionText.String18687    'Balance,Payment Method,Payment Frequency,Reason,Branch Member?,Branch,Membership Type,Cancellation Reason,,Start Date,Reprint Card?,Amount,Arrears,Last Payment,Frequency Amount,Source,
          mvHeadings = mvHeadings & DataSelectionText.String18688    'Applied,,,Source,,Voting Rights?,No. of Members,No. of Reminders,In Advance,Made on,Next Payment Due,Cancelled by,Accepted,,Effective on,Effective on
          mvWidths = "1005,2010,1100,1005,636,660,1605,1200,960,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,900,1500,1500,1200,1200,2400,2400,1200,1200,1200,1200,900,1200,1200,1200,1200,1200,1200,1200,1200,1200,885,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17865     'Contact Membership Details
          mvCode = "CMD"
          vPrimaryList = True
          mvRequiredItems = "MembersPerOrder,MembershipCard,CancelledOn,ApprovalMembership,MembershipType,FutureMembershipType,SourceCode,MembershipCardExpires,AgeOverride,Branch,BranchMember,Applied,Accepted,BranchMembership,ReprintCard,CancellationReason,CancellationSource,DirectDebitStatus,RgbMembershipType,MembershipTypeDesc,RenewalDate,NextPaymentDue,PaymentMethodDesc"

        Case DataSelectionTypes.dstContactJournal
          mvResultColumns = "ID,JournalType,Time,BatchNumber,TransactionNumber,JournalEntry,ChangedBy,Select1,Select2,Select3,ContactNumber,OrganisationContact,SelectName"
          mvSelectColumns = "Time,JournalEntry,BatchNumber,TransactionNumber,OrganisationContact"
          mvHeadings = DataSelectionText.String17866     'Time,Entry,Batch,Transaction,Organisation Contact
          mvWidths = "1500,2500,1200,1200,1200"
          mvDescription = DataSelectionText.String17867     'Contact Journal
          mvCode = "CJ"
          vPrimaryList = True
          mvRequiredItems = "ID,JournalType,Select1,Select2,Select3,ContactNumber,SelectName"

        Case DataSelectionTypes.dstContactCategoryGraphData
          mvResultColumns = "ActivityValueCode,Quantity,ValidFrom,ValidTo,ActivityValueDesc"

        Case DataSelectionTypes.dstContactStickyNotes
          mvResultColumns = "NoteNumber,Notes,CreatedOn,Permanent,AmendedBy,AmendedOn,RecordType"
          mvSelectColumns = "AmendedBy,CreatedOn,AmendedOn,Permanent,Notes,NoteNumber"
          mvHeadings = DataSelectionText.String17868     'Owner,Created On,Amended On,Permanent,Notes,Note No.
          mvWidths = "1200,1200,1200,1200,3000,1"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String17869    'Contact Sticky Notes
          mvMaintenanceDesc = "Sticky Note"
          mvCode = "SN"
          vPrimaryList = True
          mvRequiredItems = "RecordType,CreatedOn,Notes"

        Case DataSelectionTypes.dstContactOwners
          mvResultColumns = "Department,DepartmentDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "Department,DepartmentDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String22608
          mvWidths = "1,3800,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuCare
          mvDescription = DataSelectionText.String17870    'Contact Owners
          mvCode = "CO"
          mvRequiredItems = "Department,DepartmentDesc"

        Case DataSelectionTypes.dstContactHPLinks
          mvResultColumns = "ParentLink,ContactType,RelationshipCode,RelationshipDesc," & ContactNameResults() & ",ContactNumber,Phone,ContactLinkNumber"
          mvSelectColumns = "RelationshipCode,RelationshipDesc,ContactName,ContactNumber,Phone"
          mvHeadings = DataSelectionText.String22604          ',Relationship,To,No,Phone
          mvWidths = "1,2160,3605,1000,1540"
          mvDisplayTitle = DataSelectionText.String17883    'High Profile Links
          mvDescription = DataSelectionText.String17871    'Contact High Profile Links
          mvCode = "CHPL"
          mvRequiredItems = "RelationshipCode,ContactLinkNumber"

        Case DataSelectionTypes.dstContactHPCategories
          mvResultColumns = "ProfileRating,ActivityCode,ActivityDesc"
          mvSelectColumns = "ActivityDesc"
          mvHeadings = DataSelectionText.String17872     'Category
          mvWidths = "4000"
          mvDisplayTitle = DataSelectionText.String17873     'High Profile Activities
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17874    'Contact High Profile Categories
          mvCode = "CHPC"

        Case DataSelectionTypes.dstContactHPCategoryValues
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc"
          If pEnv.GetConfigOption("option_activity_quantity", False) Then mvSelectColumns = mvSelectColumns & ",Quantity"
          mvHeadings = DataSelectionText.String17875     'Category,Value,Quantity
          mvWidths = "2000,2000,1000"
          mvRequiredItems = "RgbActivityValue,ContactCategoryNumber"
          mvDisplayTitle = DataSelectionText.String17873     'High Profile Activities
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17876     'Contact High Profile Category Values
          mvCode = "CHPV"

        Case DataSelectionTypes.dstContactHeaderHPCategories
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity"
          mvHeadings = DataSelectionText.String17875     'Category,Value,Quantity
          mvWidths = "2000,2000,1000"
          mvRequiredItems = "RgbActivityValue,ContactCategoryNumber"
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17877     'Contact Header High Profile Category Values
          mvCode = "CHHP"

        Case DataSelectionTypes.dstContactDeptCategories
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc"
          If pEnv.GetConfigOption("option_activity_quantity", False) Then mvSelectColumns = mvSelectColumns & ",Quantity"
          mvHeadings = DataSelectionText.String17875     'Category,Value,Quantity
          mvWidths = "2000,2000,1000"
          mvRequiredItems = "RgbActivityValue,ContactCategoryNumber"
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17878     'Contact Department Categories
          mvDisplayTitle = DataSelectionText.String17879     'Department Activities   
          mvCode = "CDC"

        Case DataSelectionTypes.dstContactHeaderDeptCategories
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity"
          mvHeadings = DataSelectionText.String17875     'Category,Value,Quantity
          mvWidths = "2000,2000,1000"
          mvRequiredItems = "RgbActivityValue,ContactCategoryNumber"
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17880     'Contact Header Department Categories
          mvCode = "CHDC"

        Case DataSelectionTypes.dstContactHeaderHPLinks
          mvResultColumns = "ParentLink,ContactType,RelationshipCode,RelationshipDesc," & ContactNameResults() & ",ContactNumber,Phone,ContactLinkNumber"
          mvSelectColumns = "RelationshipDesc,ContactName"
          mvHeadings = DataSelectionText.String17881     'Relationship,To
          mvWidths = "1800,2000"
          mvDisplayTitle = DataSelectionText.String17883     'High Profile Links
          mvHeaderLines = 0
          mvDescription = DataSelectionText.String17884     'Contact Header High Profile Links
          mvCode = "CHHL"
          mvRequiredItems = "RelationshipCode,ContactLinkNumber"

        Case DataSelectionTypes.dstContactBankAccounts
          mvResultColumns = "BankDetailsNumber,SortCode,AccountNumber,AccountName,BankPayerName,AmendedBy,AmendedOn,Notes,DefaultAccount,HistoryOnly,IbanNumber,BicCode,BankName,BranchName"
          mvSelectColumns = "SortCode,BankName,BranchName,AccountNumber,AccountName,BankPayerName,DetailItems,Notes,AmendedBy,NewColumn,Spacer,AmendedOn"
          mvHeadings = DataSelectionText.String17885     'Sort Code,Bank Name,Branch Name,Account Number,Account Name,Payer Name,,Notes,Amended by,,,on
          mvWidths = "1500,1200,1200,2000,2500,2500,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17886     'Contact Bank Accounts
          mvMaintenanceDesc = "Bank Account"
          mvCode = "CBA"
          vPrimaryList = True
          mvRequiredItems = "AccountNumber,AccountName,SortCode,BranchName,BankName,BankPayerName,DefaultAccount,HistoryOnly"

        Case DataSelectionTypes.dstContactCreditCards
          mvResultColumns = "CreditCardDetailsNumber,CreditCardNumber,ExpiryDate,Issuer,AccountName,CreditCardTypeDesc,AmendedBy,AmendedOn,CreditCardType,IssueNumber,TokenId,TokenDesc"
          mvSelectColumns = "CreditCardDetailsNumber,CreditCardNumber,ExpiryDate,Issuer,AccountName,CreditCardTypeDesc,AmendedBy,AmendedOn,CreditCardType,IssueNumber,TokenId,TokenDesc"
          mvHeadings = DataSelectionText.String23606     'Number,Card Number,Expiry,Issuer,Account Name,Card Type,Amended By,Amended On,Card Type Code,Issue Number
          mvWidths = "1,2500,700,900,2500,1200,1200,1200,1,1200"
          mvAvailableUsages = DataSelectionUsages.dsuCare
          mvDescription = DataSelectionText.String17887     'Contact Credit Cards
          mvDisplayTitle = DataSelectionText.String17888     'Credit Card Details
          mvCode = "CCC"
          mvRequiredItems = "TokenId,TokenDesc"

        Case DataSelectionTypes.dstContactCreditCardAuthorities
          mvResultColumns = "CreditCardAuthorityNumber,AuthorityType,CreditCardNumber,StartDate,Amount,BankAccount,PaymentPlanNumber,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CreditCardDetailsNumber,ExpiryDate,Issuer,AccountName,CreditCardType,CreditCardTypeDesc,CancellationReason,CancellationSource,CancelledBy,CancelledOn,Source,SourceDesc,IssueNumber,CancellationReasonDesc,CancellationSourceDesc,AuthorityTypeCode,EncryptedCreditCardNumber"
          mvSelectColumns = "StartDate,CancelledOn,CreditCardTypeDesc,CreditCardNumber,ExpiryDate,Amount,PaymentPlanNumber,DetailItems,AuthorityType,AccountName,SourceDesc,CreatedBy,AmendedBy,CancelledBy,NewColumn,Issuer,CreditCardAuthorityNumber,BankAccount,CreatedOn,AmendedOn,CancellationReasonDesc,NewColumn2,Spacer,Spacer1,Spacer2,Spacer3,Spacer4,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String17890     'Start Date,Cancellation Date,Card Type,Card Number,Expiry Date,Amount,Payment Plan,,Authority Type,Account Name,Source,Created by,Amended by,Cancelled by,,Issuer,Authority Number,Internal Bank Account,on,on,Reason,,,,,,,Source
          mvWidths = "1200,1200,1200,2500,1200,1100,1300,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,800,1200,1200,1200,1200,1200,1200,1200,1200,1200,3600"
          mvDescription = DataSelectionText.String17891     'Contact Credit Card Authorities
          mvCode = "CCCA"
          vPrimaryList = True
          mvRequiredItems = "AuthorityType,AuthorityTypeCode,CancellationReason,CancellationSource"

        Case DataSelectionTypes.dstContactStandingOrders
          mvResultColumns = "StandingOrderNumber,StandingOrderType,StartDate,SortCode,AccountNumber,Amount,Reference,BankAccount,BankAccountDesc,PaymentPlanNumber,CreatedBy,CreatedOn,AmendedBy,AmendedOn,Source,SourceDesc,AccountName,CancellationReason,CancellationSource,CancelledBy,CancelledOn,BankDetailsNumber,IbanNumber,BicCode,CancellationReasonDesc,CancellationSourceDesc,BankName,BranchName,StandingOrderTypeCode"
          mvSelectColumns = "StartDate,CancelledOn,SortCode,AccountNumber,Amount,Reference,PaymentPlanNumber,StandingOrderType,DetailItems,BankName,AccountName,SourceDesc,CreatedBy,AmendedBy,CancelledBy,NewColumn,BranchName,StandingOrderNumber,BankAccountDesc,CreatedOn,AmendedOn,CancellationReasonDesc,NewColumn2,Spacer1,Spacer2,Spacer3,Spacer4,Spacer5,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String17892     'Start Date,Cancellation Date,Sort Code,Account Number,Amount,Reference,Payment Plan,Bankers Order Type,,Bank Name,Account Name,Source,Created by,Amended by,Cancelled by,,Branch Name,Bankers Order No.,Internal Bank Account,on,on,Reason,,,,,,,Source
          mvWidths = "1200,1200,1200,1800,1100,1800,1300,1500,1200,1200,1200,1200,1200,1200,1200,1200,1200,1500,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,600"
          mvDescription = DataSelectionText.String17893     'Contact Standing Orders
          mvCode = "CSO"
          vPrimaryList = True
          mvRequiredItems = "StandingOrderType,StandingOrderTypeCode,CancellationReason,CancellationSource"

        Case DataSelectionTypes.dstContactDirectDebits
          mvResultColumns = "DirectDebitNumber,StartDate,SortCode,AccountNumber,Amount,Reference,BankAccount,BankAccountDesc,PaymentPlanNumber,CreationNotified,CancellationNotified,MandateType,Source,SourceDesc,CreatedBy,CreatedOn,AmendedBy,AmendedOn,AccountName,CancellationReason,CancellationSource,CancelledBy,CancelledOn,BankDetailsNumber,IbanNumber,BicCode,DateSigned,BankDetailsChanged,PreviousBankDetailsNumber,CancellationReasonDesc,CancellationSourceDesc,BankName,BranchName,MandateTypeCode"
          mvResultColumns = mvResultColumns & ",Text1,Text2,Text3,Text4,Text5"
          mvSelectColumns = "StartDate,CancelledOn,SortCode,AccountNumber,Amount,Reference,PaymentPlanNumber,MandateType,DetailItems,BankName,AccountName,Source,CreatedBy,AmendedBy,CancelledBy,CancellationNotified,NewColumn,BranchName,DirectDebitNumber,BankAccountDesc,CreatedOn,AmendedOn,CancellationReasonDesc,NewColumn2,Spacer1,Spacer2,Spacer3,CreationNotified,Spacer4,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String18690     'Start Date,Cancellation Date,Sort Code,Account Number,Amount,Reference,Payment Plan,Mandate Type,,Bank Name,Account Name,Source,Created by,Amended by,Cancelled by,Cancellation Notified on,,Branch Name,
          mvHeadings = mvHeadings & DataSelectionText.String18691    'Direct Debit No.,Internal Bank Account,on,on,Reason,,,,,Notified on,,Source
          mvWidths = "1200,1200,1200,1800,1100,1800,1300,1200,1200,1200,1200,1200,1200,1200,1200,1600,1200,1200,1200,1600,1200,1200,2400,1200,1200,1200,1200,1600,1200,600"
          mvDescription = DataSelectionText.String17894     'Contact Direct Debits
          mvCode = "CDD"
          vPrimaryList = True
          mvRequiredItems = "MandateType,MandateTypeCode,CancellationReason,CancellationSource,BankDetailsNumber,SortCode,AccountNumber"

        Case DataSelectionTypes.dstContactBackOrders
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,BankAccount,Product,ProductDesc,Rate,RateDesc,Ordered,Issued,BatchType,EarliestDelivery,Status,Reference,BatchTypeDesc,BankAccountDesc"
          mvSelectColumns = "TransactionDate,ProductDesc,RateDesc,Ordered,Issued,EarliestDelivery,Status,BatchNumber,TransactionNumber,DetailItems,BatchTypeDesc,NewColumn,BankAccountDesc"
          mvHeadings = DataSelectionText.String17895     'Transaction Date,Product,Rate,Quantity Ordered,Quantity Issued,Earliest Delivery Date,Status,Batch,Transaction,,Batch Type,,Bank Account
          mvWidths = "1100,2000,2000,800,700,1100,700,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17896     'Contact Back Orders
          mvCode = "CBO"
          vPrimaryList = True
          mvRequiredItems = "Product,Rate"

        Case DataSelectionTypes.dstContactDespatchNotes
          mvResultColumns = "BatchNumber,TransactionNumber,DespatchNoteNumber,PickingListNumber,DespatchDate,DespatchMethod,DespatchTo,AddressLine,CarrierReference"
          mvSelectColumns = "DespatchNoteNumber,DespatchDate,DespatchMethod,DespatchTo,AddressLine,DetailItems,BatchNumber,NewColumn,TransactionNumber,NewColumn2,PickingListNumber,NewColumn3,CarrierReference"
          mvHeadings = DataSelectionText.String17897     'Despatch Note,Despatched on,By,To,Of,,Batch,,Transaction,,Picking List,,Carrier Reference
          mvWidths = "1716,1100,1500,2500,4000,1200,740,1200,450,1200,1000,1200,4000"
          mvDescription = DataSelectionText.String17898     'Contact Despatch Notes
          mvCode = "CDP"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactFinLinksReceived
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,ContactNumber,DonatedBy,Amount,Product,Rate,Quantity,Source,Status,VATRate,VATAmount,CreditedContactNumber,CreditedContact,CreditType"
          mvSelectColumns = "DonatedBy,Amount,Product,Rate,Quantity,Status,VATRate,VATAmount,DetailItems,BatchNumber,Source,NewColumn,TransactionNumber,NewColumn2,LineNumber"
          mvHeadings = DataSelectionText.String17899     ' ,Amount,Product,Rate,Qty,Status,VAT Rate,VAT Amount,,Batch,Source,,Transaction,,Line
          mvWidths = "2000,900,2000,1100,400,600,500,900,1200,1200,2000,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17900     'Contact Financial Links Received
          mvCode = "CFLR"
          vPrimaryList = True
          mvRequiredItems = "DonatedBy,Amount"

        Case DataSelectionTypes.dstContactFinLinksDonated
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,ContactNumber,DonatedTo,Amount,Product,Rate,Quantity,Source,Status,VATRate,VATAmount,CreditedContactNumber,CreditedContact,CreditType"
          mvSelectColumns = "DonatedTo,Amount,Product,Rate,Quantity,Status,VATRate,VATAmount,DetailItems,BatchNumber,Source,NewColumn,TransactionNumber,NewColumn2,LineNumber"
          mvHeadings = DataSelectionText.String17899     ' ,Amount,Product,Rate,Qty,Status,VAT Rate,VAT Amount,,Batch,Source,,Transaction,,Line
          mvWidths = "2000,900,2000,1100,400,600,500,900,1200,1200,2000,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17940     'Contact Financial Links Donated
          mvCode = "CFLD"
          vPrimaryList = True
          mvRequiredItems = "DonatedTo"

        Case DataSelectionTypes.dstContactPaymentPlans
          mvResultColumns = "PaymentPlanTypeDesc,PaymentPlanNumber,ContactNumber,AddressNumber,PaymentPlanType,StartDate,PaymentMethod,PaymentMethodDesc,PaymentFrequency,PaymentFrequencyDesc,RenewalDate,RenewalAmount,SourceCode,CancelledOn,CreatedBy,CreatedOn,GiverContactNumber,Provisional,"
          mvResultColumns = mvResultColumns & "FrequencyAmount,NextPaymentDue,Balance,Arrears,LastPaymentAmount,LastPaymentDate,Amount,InAdvance,Term,TermUnits,CancelledBy,CancellationReason,CancellationSource,FutureCancellationReason,FutureCancellationDate,FutureCancellationSource,"
          mvResultColumns = mvResultColumns & "RenewalDateChangeReason,RenewalDateChangedBy,RenewalDateChangedOn,RenewalDateChangeValue,DirectDebitStatus,CreditCardStatus,StandingOrderStatus,CovenantStatus,SalesContactNumber,RenewalPending,ExpiryDate,TheirReference,AmendedBy,AmendedOn,ReasonForDespatch,EligibleForGiftAid,PaymentFrequencyFrequency,NumberOfReminders,GiftMembership,OneYearGift,PaymentScheduleAmendedOn,FirstAmount,ClaimDay,"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanPackToMember) Then
            mvResultColumns = mvResultColumns & "PackToMember,"
          Else
            mvResultColumns = mvResultColumns & "PackToDonor,"
          End If
          mvResultColumns = mvResultColumns & "PayPlanMembershipTypeCode,"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & "OneOffPayment,"
          mvResultColumns = mvResultColumns & "LoanStatus,AddressLine,NextPaymentAmount,NumberOfPayments,SourceDesc,CancellationReasonDesc,CancellationSourceDesc,FutureCancellationReasonDesc,FutureCancellationSourceDesc,ReasonForDespatchDesc,RenewalDateChangeReasonDesc,SalesContactName,GiverContactName,GiverContactAddressLine,TermDesc,UnprocessedPayments,NonDonationProducts,FirstYearMembership,NewOrderPackToDonor,MembershipRateCode,OriginalPPDFixedAmount"
          mvSelectColumns = "PaymentPlanNumber,Provisional,PaymentPlanTypeDesc,RenewalDate,CancelledOn,Term,TermUnits,SourceDesc,AmendedBy,AmendedOn,DetailItems,Spacer,Spacer1,Spacer2,Spacer3,Spacer4,Spacer5,Spacer6,Spacer10,Spacer11,Spacer13,Spacer12,Spacer15,Spacer16,NewColumn,StartDate,Amount,RenewalPending,RenewalDateChangeReasonDesc,RenewalDateChangeValue,Balance,PaymentMethodDesc,PaymentFrequencyDesc,CancellationReasonDesc,FutureCancellationReasonDesc,ReasonForDespatchDesc,AddressLine,GiverContactName,NewColumn2,CreatedBy,CovenantStatus,RenewalAmount,RenewalDateChangedBy,Spacer7,Arrears,LastPaymentAmount,FrequencyAmount,CancellationSourceDesc,FutureCancellationSourceDesc,TheirReference,Spacer14,GiverContactAddressLine,NewColumn3,CreatedOn,Spacer9,NumberOfReminders,RenewalDateChangedOn,Spacer8,InAdvance,LastPaymentDate,NextPaymentDue,CancelledBy,FutureCancellationDate,EligibleForGiftAid,SourceCode"
          mvHeadings = DataSelectionText.String18692     'Payment Plan,Provisional,Type,Renewal Date,Cancellation Date,Term,Term Units,Source Desc,Amended By,Amended On,,GENERAL,.,RENEWAL,RENEWAL DATE,.,FINANCIAL,.,.,CANCELLATION,FUTURE CHANGE,MISCELLANEOUS,.,.,,
          mvHeadings = mvHeadings & DataSelectionText.String18693    'Start Date,Fixed Amount,Pending?,Change Reason,Change Value,Balance,Payment Method,Payment Frequency,Reason,Cancellation Reason,Reason For Despatch,Payer Address,Giver Contact,,Created by,Covenant Status,
          mvHeadings = mvHeadings & DataSelectionText.String18694    'Amount,Changed by,,Arrears,Last Payment,Frequency Amount,Source,Source,Their Reference,,Giver Address,,Created on,,No. of Reminders,Changed on,,In Advance,Made on,Next Payment Due,Cancelled by,Effective on,Eligible For Gift Aid?
          mvWidths = "1260,1000,1065,1100,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1100,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1000,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,2400,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17941     'Contact Payment Plans
          mvCode = "CPP"
          vPrimaryList = True
          mvRequiredItems = "PaymentPlanType,CancellationReason,CancellationSource,LoanStatus,UnprocessedPayments"

        Case DataSelectionTypes.dstContactPaymentPlansPayments
          Dim vDS As New DataSelection(mvEnv, DataSelectionTypes.dstContactPaymentPlans, pParams, pListType, pUsage, pGroup)
          mvResultColumns = vDS.mvResultColumns & ",PayPlanMembershipTypeDesc,MemberNumber,LabelName"
          mvSelectColumns = "PaymentPlanNumber,NextPaymentDue,PayPlanMembershipTypeDesc,MemberNumber,NextPaymentAmount"
          mvHeadings = "Plan Number,Due on,Membership Type,Member Number,Amount"
          mvWidths = "1200,1200,1200,1200,1200"
          mvCode = "CPN"
          mvDescription = "Contact Payment Plans Payments"

        Case DataSelectionTypes.dstContactServiceBookings
          mvResultColumns = "ServiceBookingNumber,ServiceContact,StartDate,EndDate,TransactionDate,GrossAmount,VATAmount,VATRateDesc,CancelledBy,CancelledOn,CancellationSource,AmendedBy,AmendedOn,Status,BookingContactNumber,BookingAddressNumber,ServiceContactNumber,RelatedContactNumber,BatchNumber,TransactionNumber,LineNumber,PaymentPlanNumber,SalesContactNumber,CancellationReason,NetAmount,RelatedContactName,SalesContactName,CancellationReasonDesc,CancellationSourceDesc"
          mvSelectColumns = "TransactionDate,ServiceBookingNumber,ServiceContact,RelatedContactName,StartDate,EndDate,GrossAmount,VATAmount,NetAmount,CancelledOn,DetailItems,BatchNumber,VATRateDesc,AmendedBy,CancelledBy,NewColumn,TransactionNumber,SalesContactName,AmendedOn,CancellationReasonDesc,NewColumn2,LineNumber,Spacer,Spacer1,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String18697     'Transaction Date,Service Booking,Service Contact,Related Contact,Start Date,End Date,Gross Amount,VAT Amount,Net Amount,Cancelled On,,Batch,VAT Rate,Amended by,Cancelled by,,Transaction,Sales Contact,on,Reason,,Line,,,Source
          mvWidths = "1600,1500,2000,2000,1100,1100,1250,1250,1250,1300,1200,1200,600,1300,1300,1200,1200,2000,1300,2400,1200,1200,1200,1200,2400"
          mvDescription = DataSelectionText.String17942    'Contact Service Bookings
          mvCode = "CSB"
          vPrimaryList = True
          mvRequiredItems = "EndDate,Status,CancellationReason"

        Case DataSelectionTypes.dstContactSubscriptions
          mvResultColumns = "SubscriptionNumber,AddressNumber,PaymentPlanNumber,ProductDesc,Quantity,DespatchMethodDesc,ReasonForDespatchDesc,ValidFrom,ValidTo,CancellationReason,CancelledBy,CancelledOn,CancellationSource,ProductCode,DespatchMethod,ReasonForDespatch,CommunicationNumber,CancellationReasonDesc,CancellationSourceDesc,AddressLine,DeliverTo"
          mvSelectColumns = "PaymentPlanNumber,ProductDesc,Quantity,DespatchMethodDesc,ReasonForDespatchDesc,ValidFrom,ValidTo,CancelledOn,DetailItems,DeliverTo,CancelledBy,NewColumn,Spacer,CancellationReason,NewColumn2,Spacer1,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String17943     'Payment Plan,Product,Qty,Despatch Method,Despatch Reason,Valid From,Valid To,Cancelled On,,Deliver to,Cancelled by,,,Reason,,,Source
          mvWidths = "1200,3000,500,1680,1680,1200,1200,1200,1200,2000,1200,1200,1200,1600,1200,1200,1200"
          mvDescription = DataSelectionText.String17944    'Contact Subscriptions
          mvCode = "CS"
          vPrimaryList = True
          mvRequiredItems = "CancellationReason"

        Case DataSelectionTypes.dstContactDBANotes
          mvResultColumns = "Master,Duplicate,MergedOn,Notes"
          mvSelectColumns = "Duplicate,MergedOn,Notes"
          mvHeadings = DataSelectionText.String17945     'Duplicate,Merged On,Notes
          mvWidths = "1200,1200,4000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String17946    'Contact DBA Notes
          mvCode = "CDBN"
          mvRequiredItems = "Notes"

        Case DataSelectionTypes.dstContactCovenants
          mvResultColumns = "PaymentPlanNumber,CovenantNumber,StartDate,SignatureDate,CovenantTerm,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CancelledBy,CancelledOn,LastTaxClaim,TaxClaimedTo,R185Return,R185Sent,CovenantedAmount,CancellationReason,PaymentFrequency,PaymentMethod,PaymentFrequencyDesc,PaymentMethodDesc,DepositedDeed,Net,AnnualClaim,Fixed,CovenantStatus,CancellationSource,CancellationReasonDesc,CancellationSourceDesc,CovenantStatusDesc"
          mvSelectColumns = "CovenantNumber,StartDate,CovenantTerm,CovenantedAmount,SignatureDate,CancelledOn,PaymentPlanNumber,DetailItems,R185Sent,R185Return,LastTaxClaim,CovenantStatusDesc,AmendedBy,CancelledBy,NewColumn,PaymentFrequencyDesc,PaymentMethodDesc,TaxClaimedTo,Spacer,AmendedOn,CancellationReasonDesc,NewColumn2,DepositedDeed,Net,Fixed,AnnualClaim,Spacer1,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String18698     'Covenant Number,Start Date,Term,Covenanted Amount,Signature Date,Cancellation Date,Payment Plan Number,,R185 Sent,R185 Returned,Last Tax Claim,Covenant Status,Amended by,Cancelled by,,Payment Frequency,Payment Method,
          mvHeadings = mvHeadings & DataSelectionText.String18699    'Tax Claimed To,,on,Reason,,Deposited Deed?,Net?,Fixed Amount?,Annual Claim?,,Source
          mvWidths = "1380,1000,1200,1200,1200,1200,1740,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,2400,1200,1200,1200,1200,1200,1200,2400"
          mvDescription = DataSelectionText.String17949    'Contact Covenants
          mvCode = "CCOV"
          vPrimaryList = True
          mvRequiredItems = "CancellationReason,CancellationSource"

        Case DataSelectionTypes.dstContactPreviousNames
          mvResultColumns = "PreviousName,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "PreviousName,CreatedOn"
          mvHeadings = DataSelectionText.String17951     'Previous Name,Created On
          mvWidths = "2950,1600"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuCare
          mvDescription = DataSelectionText.String17952    'Contact Previous Names
          mvDisplayTitle = DataSelectionText.String17953     'Previous Names
          mvCode = "CPRN"
          'mvRequiredItems = "Notes"

        Case DataSelectionTypes.dstContactSponsorshipClaimedPayments
          mvResultColumns = "ContactNumber,ClaimNumber,BatchNumber,TransactionNumber,LineNumber,ProductDesc,AmountClaimed,NetAmount,ClaimDate,Reference"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,ProductDesc,AmountClaimed,NetAmount,ClaimNumber,ClaimDate,Reference"
          mvHeadings = DataSelectionText.String17954     'Batch,Transaction,Line,Product,Amount Claimed,Net Amount,Claim Number,Claim Date,Reference
          mvWidths = "800,800,800,3000,1000,1000,800,1100,2000"
          mvDescription = DataSelectionText.String17955    'Contact Sponsorship Claimed
          mvCode = "CSC"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactSponsorshipUnClaimedPayments
          mvResultColumns = "ContactNumber,BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,Amount,Reference"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,Amount,Reference"
          mvHeadings = DataSelectionText.String17956     'Batch,Transaction,Line,Date,Product,Amount,Reference
          mvWidths = "800,800,800,1100,3000,1000,2000"
          mvDescription = DataSelectionText.String17957    'Contact Sponsorship UnClaimed
          mvCode = "CSU"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactGiftAidDeclarations
          mvResultColumns = "ContactNumber,DeclarationNumber,Date,DeclarationType,Source,ConfirmedOn,DeclarationMethod,StartDate,EndDate,SourceDesc,BatchNumber,TransactionNumber,PaymentPlanNumber,CancellationReason,CancelledBy,CancelledOn,CancellationSource,AmendedBy,AmendedOn,CreatedBy,CreatedOn,Notes,LabelName,Title,Forenames,Initials,Surname,Honorifics,Salutation,PreferredForename,Sex,PrefixHonorifics,SurnamePrefix,InformalSalutation,Address,BuildingNumber,HouseName,Town,County,Postcode,Country,Branch,CancellationReasonDesc,CancellationSourceDesc,DeclarationMethodDesc,Donations,Members,Summary"
          mvSelectColumns = "DeclarationNumber,Date,DeclarationType,DeclarationMethodDesc,ConfirmedOn,StartDate,EndDate,CancelledOn,DetailItems,BatchNumber,CreatedBy,AmendedBy,CancelledBy,NewColumn,TransactionNumber,CreatedOn,AmendedOn,CancellationReasonDesc,NewColumn2,PaymentPlanNumber,SourceDesc,Notes,CancellationSourceDesc"
          mvHeadings = DataSelectionText.String17958     'Declaration Number,Date,Type,Method,Confirmed On,tart Date,End Date,Cancelled On,,Batch,Created by,Amended by,Cancelled by,,Transaction,on,on,Reason,,Payment Plan,Source,Notes,Source
          mvWidths = "1200,1200,1600,1000,1300,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,3000"
          mvDescription = DataSelectionText.String17959    'Contact Gift Aid Declarations
          mvMaintenanceDesc = "Gift Aid Declaration"
          mvCode = "CGAD"
          vPrimaryList = True
          mvRequiredItems = "DeclarationMethod,DeclarationType,DeclarationMethodDesc,Donations,Members,CancellationReason,Summary,LabelName"

        Case DataSelectionTypes.dstContactCreditCustomers
          mvResultColumns = "ContactNumber,AddressNumber,Company,CompanyDesc,SalesLedgerAccount,CreditCategory,CreditCategoryDesc,StopCode,CreditLimit,Outstanding,OnOrder,CustomerType,TermsNumber,TermsPeriod,TermsFrom,LastStatementDate,LastStatementClosingBalance,LastStatementNumber,StatementPeriod,AmendedBy,AmendedOn,LabelName"
          mvSelectColumns = "CompanyDesc,SalesLedgerAccount,CreditCategoryDesc,CreditLimit,Outstanding,StopCode"
          mvHeadings = DataSelectionText.String22850     'Company,Sales Ledger Account,Credit Category,Credit Limit,Outstanding,Stop Code
          mvWidths = "2000,1200,2000,1000,1000,900"
          mvRequiredItems = "Company,CreditCategory,CreditLimit,Outstanding,OnOrder,TermsPeriod,TermsFrom"
          mvDescription = DataSelectionText.String40028
          mvMaintenanceDesc = "Sales Ledger Account"
          mvCode = "CCCU"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactOutstandingInvoices
          mvResultColumns = "PayButton,PayCheck,BatchNumber,TransactionNumber,InvoiceNumber,InvoiceDate,PaymentDue,AmountPaid,InvoicePayStatus,InvoiceDisputeCode,RecordType,Amount,DepositAmount,Outstanding"
          mvSelectColumns = "PayButton,PayCheck,InvoiceNumber,InvoiceDate,Amount,AmountPaid,PaymentDue,InvoiceDisputeCode,BatchNumber,TransactionNumber,InvoicePayStatus,RecordType,DepositAmount"
          mvHeadings = ",,Invoice Number,Invoice Date,Invoice Amount,Amount Paid,Payment Due,Invoice Dispute Code,Batch Number,Transaction Number,Invoice Pay Status,Record Type,Deposit Amount"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1,1,1,1,1"
          mvRequiredItems = "PayButton,PayCheck,InvoiceDate,Amount,AmountPaid,PaymentDue,InvoiceDisputeCode,InvoicePayStatus,RecordType,DepositAmount,Outstanding"
          mvDescription = "Outstanding Invoices"
          mvCode = "COIN"

        Case DataSelectionTypes.dstContactCashInvoices
          mvResultColumns = "BatchNumber,TransactionNumber,InvoiceNumber,InvoiceDate,PaymentDue,AmountPaid,InvoicePayStatus,InvoiceDisputeCode,RecordType,Amount,InvoiceAmount,BatchType,TransactionAmount,ContactNumber,AddressNumber,SalesLedgerAccount,AmountUsed"
          mvSelectColumns = "BatchNumber,TransactionNumber,InvoiceNumber,InvoiceDate,PaymentDue,AmountPaid,InvoicePayStatus,InvoiceDisputeCode,RecordType,Amount,InvoiceAmount,AmountUsed,ContactNumber,AddressNumber,SalesLedgerAccount"
          mvHeadings = mvResultColumns
          mvRequiredItems = "BatchNumber,TransactionNumber,BatchType,TransactionAmount,ContactNumber,AddressNumber,SalesLedgerAccount"

        Case DataSelectionTypes.dstContactGAYEPledges
          mvResultColumns = "PledgeNumber,DonorID,Employer,Agency,PayrollCompany,Amount,NetTotal,StartDate,PFOCode,SourceCode,Source,DonorTotal,EmployerTotal,GovernmentTotal,AdminFees,ProductCode,Product,RateCode,Rate,CancellationReason,CancellationReasonDesc,CancelledOn,CancelledBy,CancellationSource,CancellationSourceDesc,DistributionCode,PayrollNumber,PaymentFrequency,PaymentFrequencyDesc,AmendedOn,AmendedBy,EmployerOrganisationNumber,AgencyOrganisationNumber,CreatedBy,CreatedOn,OtherMatchedTotal,CharityDonorReference"
          mvSelectColumns = "PledgeNumber,StartDate,CancelledOn,DonorID,Employer,Agency,Amount,PaymentFrequencyDesc,NetTotal,DetailItems,DonorTotal,EmployerTotal,GovernmentTotal,AdminFees,CancellationReasonDesc,AmendedBy,NewColumn,Source,DistributionCode,PayrollCompany,Product,CancellationSourceDesc,AmendedOn,NewColumn2,PayrollNumber,PFOCode,Spacer,Rate,CancelledBy"
          mvHeadings = DataSelectionText.String22854     'Pledge No,Start Date,Cancellation Date,Donor ID,Employer,Agency,Amount,Payment Frequency,Net Total,,Donor Total,Employer Total,Government Total,Admin Fees,Cancellation Reason,Amended by,,Source,Distribution Code,
          mvHeadings = mvHeadings & DataSelectionText.String22855     'Payroll Company,Product,Cancellation Source,Amended on,,Payroll Number,PFO Code,,Rate,Cancelled by
          mvWidths = "800,1200,1200,1000,2400,2400,1000,1800,1000,1200,1000,1000,1000,1000,1200,1200,1200,1200,1500,2400,1200,1200,1200,1200,1800,1000,1200,1200,1200"
          mvDescription = DataSelectionText.String17960    'Contact Pre Tax PG Pledges
          mvMaintenanceDesc = "Pre-Tax Pledge"
          mvCode = "CGAY"
          mvRequiredItems = "CancellationReason"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactGAYEPledgePayments
          mvResultColumns = "TransactionSign,PledgeNumber,TransactionDate,BatchNumber,TransactionNumber,Amount,DonorAmount,EmployerAmount,GovernmentAmount,AdminFeeAmount,OtherMatchedAmount,PaymentNumber"
          mvSelectColumns = "TransactionDate,BatchNumber,TransactionNumber,Amount,DonorAmount,EmployerAmount,GovernmentAmount,AdminFeeAmount"
          mvHeadings = DataSelectionText.String17963     'Date,Batch,Transaction,Net Amount,Donor Amount,Employer Amount,Government Amount,Admin Fee
          mvWidths = "1200,1000,1000,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String17964    'Contact Pre Tax PG Pledge Payments
          mvDisplayTitle = DataSelectionText.String17965     'Payments
          mvCode = "CGYP"

        Case DataSelectionTypes.dstContactGAYEPostTaxPledges
          mvResultColumns = "PledgeNumber,Employer,Employee,PayrollNumber,StartDate,PledgeAmount,PledgeTotal,DonorTotal,EmployerTotal,Source,SourceDesc,Product,ProductDesc,Rate,RateDesc,DistributionCode,DistributionCodeDesc,LastPaymentDate,PaymentNumber,CancellationReason,CancellationReasonDesc,CancelledOn,CancelledBy,CancellationSource,CancellationSourceDesc,AmendedOn,AmendedBy,EmployerOrganisationNumber,EmployeeContactNumber,EmployerPayrollOrganisationNumber,EmployerPayrollOrganisationName,CreatedBy,CreatedOn"
          mvSelectColumns = "PledgeNumber,StartDate,CancelledOn,Employer,Employee,PledgeAmount,PledgeTotal,EmployerPayrollOrganisationName,DetailItems,DonorTotal,EmployerTotal,Product,CancellationReason,AmendedBy,NewColumn,Source,DistributionCode,Rate,CancellationSource,AmendedOn,NewColumn2,PayrollNumber,Spacer,Spacer2,CancelledBy"
          mvHeadings = DataSelectionText.String22858     'Pledge No,Start Date,Cancellation Date,Employer,Employee,Pledge Amount,Pledge Total,Payroll Organisation,,Donor Total,Employer Total,Product,Cancellation Reason,Amended By,,Source,Distribution Code,Rate,
          mvHeadings = mvHeadings & DataSelectionText.String22859    'Cancellation Source,Amended On,,Payroll Number,,,Cancelled By
          mvWidths = "800,1200,1200,2400,2400,1000,1000,2400,1200,1000,1000,1200,1200,1200,1200,1200,1500,1200,1200,1200,1200,1800,1200,1200,1200"
          mvDescription = DataSelectionText.String17966    'Contact Post Tax PG Pledges
          mvMaintenanceDesc = "Post-Tax Pledge"
          mvCode = "CGPT"
          mvRequiredItems = "CancellationReason"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactGAYEPostTaxPledgePayments
          mvResultColumns = "TransactionSign,TransactionDate,BatchNumber,TransactionNumber,PayAmount,DonorAmount,EmployerAmount,Reference"
          mvSelectColumns = "TransactionSign,TransactionDate,BatchNumber,TransactionNumber,PayAmount,DonorAmount,EmployerAmount,Reference"
          mvHeadings = DataSelectionText.String17969     'Sign,Date,Batch,Transaction,Pay Amount,Donor Amount,Employer Amount,Reference
          mvWidths = "1,1200,1000,1000,1200,1200,1200,1800"
          mvRequiredItems = "TransactionSign"
          mvDescription = DataSelectionText.String17970    'Contact Post Tax PG Pledge Payments
          mvDisplayTitle = DataSelectionText.String17965    'Payments
          mvCode = "CGPP"

        Case DataSelectionTypes.dstClaimedPayments
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,AmountClaimed,NetAmount,ClaimNumber,ClaimDate"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,NetAmount,AmountClaimed,ClaimNumber,ClaimDate"
          mvHeadings = DataSelectionText.String17971     'Batch,Transaction,Line,Date,Product,Net Amount,Amount Claimed,Claim Number,Claim Date
          mvWidths = "800,800,800,1100,3000,1000,1000,800,1000"
          mvDescription = DataSelectionText.String17972    'Declaration Claimed Payments
          mvDisplayTitle = DataSelectionText.String17973     'Claimed Payments
          mvCode = "DCP"

        Case DataSelectionTypes.dstUnClaimedPayments
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,Amount"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,ProductDesc,Amount"
          mvHeadings = DataSelectionText.String17974     'Batch,Transaction,Line,Date,Product,Amount
          mvWidths = "800,800,800,1100,3000,900"
          mvDescription = DataSelectionText.String17975    'Declaration Unclaimed Payments
          mvDisplayTitle = DataSelectionText.String17976     'Unclaimed Payments
          mvCode = "DUP"

        Case DataSelectionTypes.dstDespatchStock
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,Product,ProductDesc,Quantity,Warehouse,WarehouseDesc"
          mvSelectColumns = "LineNumber,Product,ProductDesc,Quantity"
          mvHeadings = DataSelectionText.String17977     'Line No.,Product,Description,Quantity
          mvWidths = "1200,3060,4000,1000"
          mvDescription = DataSelectionText.String17978    'Despatched Stock
          mvDisplayTitle = DataSelectionText.String17978     'Despatched Stock
          mvCode = "DS"

        Case DataSelectionTypes.dstFinancialHistoryDetails
          mvResultColumns = "LineNumber,Product,ProductDesc,Rate,RateDesc,DistributionCode,Quantity,Amount,VATAmount,Status,Source,SourceDesc,VATRateDesc,CurrencyAmount,CurrencyVATAmount,CurrencyCode,SalesContactNumber,Notes,Issued,Warehouse,InvoicePayment,StockProduct,RgbAmount,RgbCurrencyAmount,PaymentPlanType,Number,PaymentPlanNumber,SalesContactName,DistributionCodeDesc,PaymentPlanPayNumber"
          mvSelectColumns = "Product,Rate,Quantity,PaymentPlanType,Number,Amount,VATAmount,VATRateDesc,Status,Source,SourceDesc,DistributionCode,SalesContactName,StockProduct"
          mvHeadings = DataSelectionText.String17979     'Product,Rate,Qty,PayPlan Type,Number,Amount,VAT,VAT Rate,Status,Source,Description,Distribution Code,Sales Contact
          mvWidths = "2000,500,400,500,900,900,900,1000,600,1500,3000,1704,3000,1"
          mvDescription = DataSelectionText.String17980    'Financial History Details
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvCode = "FHD"
          mvRequiredItems = "PaymentPlanType,Status,Product,Rate,Quantity,Issued,StockProduct,Warehouse,PaymentPlanNumber,PaymentPlanPayNumber,InvoicePayment,Amount,RgbAmount,RgbCurrencyAmount"

        Case DataSelectionTypes.dstBACSAmendments
          mvResultColumns = "AmendmentNumber,DirectDebitNumber,OldBankDetailsNumber,NewBankDetailsNumber,BACSRecordType,BACSAdviceReason,EffectiveDate,AdviceReference,PayersName,PayersSortCode,PayersAccounNumber,AdviceDueDate,BACSPaymentFrequency,AmountPayment,PayersNewName,PayersNewSortCode,PayersNewAccountNumber,NewDueDate,NewBACSPaymentFrequency,NewAmountPayment,LastPaymentDate,BuildingSocietyRollNumber1,BuildingSocietyRollNumber2,BuildingSocietyRollNumber3,BACSTransactionCode,OriginatorsSequenceNumber,OriginatorsSortCode,OriginatorsAccountNumber,UserNumber,BACSNotes,Notes,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then mvResultColumns = mvResultColumns & ",PayersIbanNumber,PayersBicCode,OriginatorsIbanNumber,OriginatorsBicCode,EndToEndId"
          mvSelectColumns = "AmendmentNumber,DirectDebitNumber,OldBankDetailsNumber,NewBankDetailsNumber,BACSRecordType,BACSAdviceReason,EffectiveDate,AdviceReference,PayersName,PayersSortCode,PayersAccounNumber,AdviceDueDate,BACSPaymentFrequency,AmountPayment,PayersNewName,PayersNewSortCode,PayersNewAccountNumber,NewDueDate,NewBACSPaymentFrequency,NewAmountPayment,LastPaymentDate,BuildingSocietyRollNumber1,BuildingSocietyRollNumber2,BuildingSocietyRollNumber3,BACSTransactionCode,OriginatorsSequenceNumber,OriginatorsSortCode,OriginatorsAccountNumber,UserNumber,BACSNotes,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String22860     'Amendment No,Direct Debit No,Old Bank Details No,New Bank Details No,BACS Record Type,BACS Advice Reason,Effective Date,Advice Reference,Payers Name,Payers Sort Code,Payers Accnt No,Advice Due Date,BACS Pay Freq,Amount Payment,
          mvHeadings = mvHeadings & DataSelectionText.String22861    'Payers New Name,Payers New Sort Code,Payers New Accnt No,New Due Date,New BACS Pay Freq,New Amount Payment,Last Payment Date,Bdg Soc Roll No1,Bdg Soc Roll No2,Bdg Soc Roll No3,BACS Trans Code,Originators Seq No,Originators Sort Code,
          mvHeadings = mvHeadings & DataSelectionText.String22862    'Originators Account Number,User Number,BACS Notes,Notes,Amended By,Amended On
          mvWidths = "10,1200,1600,1600,2000,2000,1200,1800,1800,1200,1800,1200,400,1100,1800,1200,1800,1200,400,1100,1200,1000,400,1600,400,1200,1200,1800,1000,400,3000,1200,1200"
          mvDescription = DataSelectionText.String17982    'BACS Amendments
          mvDisplayTitle = mvDescription
          mvCode = "BA"

        Case DataSelectionTypes.dstPostPointRecipients
          mvResultColumns = "ContactNumber,ContactName"

        Case DataSelectionTypes.dstRates
          mvResultColumns = "Rate,RateDesc,Product,ProductDesc,CurrentPrice,FuturePrice,PriceChangeDate,VATExclusive,CurrentPriceLowerLimit,CurrentPriceUpperLimit"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFixedPrice) Then mvResultColumns = mvResultColumns & ",FixedPrice"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMinPriceMandatory) Then mvResultColumns = mvResultColumns & ",UpperLowerPriceMandatory"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then mvResultColumns = mvResultColumns & ",PriceIsPercentage"
          mvResultColumns = mvResultColumns & ",DaysPriorTo,DaysPriorFrom,MembershipLookupGroup,StartDate,UseModifiers"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then mvResultColumns = mvResultColumns & ",CurrencyCode"
          mvResultColumns &= ",LoanInterest"

        Case DataSelectionTypes.dstDepartmentActivities
          mvResultColumns = "Activity,ActivityDesc,HighProfile,ProfileRating,ContactGroup"

        Case DataSelectionTypes.dstDepartmentActivityValues
          mvResultColumns = "ActivityValue,ActivityValueDesc"

        Case DataSelectionTypes.dstBrowserContactPositions
          mvResultColumns = "ContactNumber,ContactGroup,ContactName,Position"
          mvSelectColumns = "ContactNumber,ContactName,Position"
          mvHeadings = DataSelectionText.String17983     'Contact Number,Name,Position
          mvWidths = "900,2000,2000"
          mvRequiredItems = "ContactGroup"
          mvHeaderLines = 0

        Case DataSelectionTypes.dstProductWarehouses
          mvResultColumns = "Product,ProductDesc,Warehouse,BinNumber,WarehouseDesc,ReOrderLevel,LastStockCount,CostOfSale,WarehouseStock"

        Case DataSelectionTypes.dstTransactionAnalysis
          mvResultColumns = "LineNumber,Product,Rate,DistributionCode,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,CurrencyAmount,CurrencyVatAmount,SalesContactNumber,SalesContactName,ProductNumber,Notes,LineType,RgbAmount,RgbCurrencyAmount,InvoiceNumber,LineTypeDesc,Number,ItemType,ProductDesc,RateDesc,EventDesc,StartDate,RoomDesc,ServiceDesc,ExamUnitDescription,Description,MemberNumber,OrderNumber,CovenantNumber"
          mvSelectColumns = "LineTypeDesc,Number,Product,Rate,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,DistributionCode,SalesContactName,Notes" ',ItemType,ProductDesc,RateDesc,EventDesc,RoomDesc,ServiceDesc,Description"
          mvHeadings = DataSelectionText.String17984     'Line Type,Number,Product,Rate,Qty,Amount,Vat Amount,Vat Rate,Source Code,Source,Distribution Code,Sales Contact,Notes
          mvWidths = "1200,1200,2000,1000,400,1000,1200,1200,1200,3000,1000,3000,3000"
          mvDescription = DataSelectionText.String17985    'Batch Transaction Analysis
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvRequiredItems = "Amount,VatAmount,ItemType,LineNumber,Product,RgbAmount,RgbCurrencyAmount"
          mvCode = "BTA"

        Case DataSelectionTypes.dstDocuments, DataSelectionTypes.dstDistinctDocuments, DataSelectionTypes.dstDistinctExternalDocuments, DataSelectionTypes.dstContactDocuments, DataSelectionTypes.dstDistinctContactDocuments, DataSelectionTypes.dstEventDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc"
          If pType = DataSelectionTypes.dstContactDocuments Or pType = DataSelectionTypes.dstDistinctDocuments Or pType = DataSelectionTypes.dstDistinctExternalDocuments Then mvResultColumns = mvResultColumns & ",Precis"
          mvResultColumns = mvResultColumns & ",Subject,CallDuration,TotalDuration,SelectionSet,OriginalUri"      'Optional Attributes
          If pUsage = DataSelectionUsages.dsuWEBServices Then
            mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc,DetailItems,Precis,CreatedBy,Source,PackageCode,NewColumn,Spacer,DepartmentDesc,StandardDocument,NewColumn2,Spacer1,DocumentClassDesc,TheirReference"
            mvHeadings = DataSelectionText.String17986     'Document Number,Dated,In/ Out,Subject,Reference+,Document Type,Topic+,Sub Topic,,Precis,Creator,Source,Package,,,Department,Standard Document,,,Document Class,Their Reference
            mvWidths = "1200,1200,1200,1500,1500,1500,1400,1400,1200,3600,1400,1200,600,1200,1200,1400,1200,1200,1200,1200,1200"
          ElseIf pUsage = DataSelectionUsages.dsuSmartClient Then
            mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc,DetailItems,CreatedBy,Source,PackageCode,NewColumn,DepartmentDesc,StandardDocument,NewColumn2,DocumentClassDesc,TheirReference"
            mvHeadings = DataSelectionText.String17987     'Document Number,Dated,In/ Out,Subject,Reference+,Document Type,Topic+,Sub Topic,,Creator,Source,Package,,Department,Standard Document,,Document Class,Their Reference
            mvWidths = "1200,1200,1200,1500,1500,1500,1400,1400,1200,1400,1200,600,1200,1400,1200,1200,1200,1200"
            mvRequiredItems = "Access"
          End If
          If mvType = DataSelectionTypes.dstContactDocuments Then
            mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          Else
            mvAvailableUsages = DataSelectionUsages.dsuNone
          End If
          If mvType = DataSelectionTypes.dstEventDocuments Then
            mvDescription = "Event Documents"
            mvCode = "ED"
          Else
            mvDescription = DataSelectionText.String17988    'Contact Documents
            mvCode = "CD"
          End If
          mvMaintenanceDesc = "Document"
          vPrimaryList = True

        Case DataSelectionTypes.dstActionSubjects
          mvResultColumns = "TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,Notes,Quantity,AmendedOn,AmendedBy"
          mvSelectColumns = "TopicDesc,SubTopicDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String17989     'Topic,Sub Topic,Amended by,Amended on
          mvWidths = "2000,2000,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String17990    'Action Analysis
          mvDisplayTitle = DataSelectionText.String17991     'Analysis
          mvMaintenanceDesc = "Analysis"
          mvCode = "ACS"
          mvRequiredItems = "TopicCode,SubTopicCode"

        Case DataSelectionTypes.dstActionOutline
          mvResultColumns = "ActionDesc,ActionNumber,MasterAction,ActionLevel,SequenceNumber,CreatedBy,Department,ActionPriority,ActionStatus,ActionPriorityDesc,ActionStatusDesc"
          mvSelectColumns = mvResultColumns

        Case DataSelectionTypes.dstPriorActions
          mvResultColumns = "ActionNumber,ActionDesc,SequenceNumber"
          mvSelectColumns = mvResultColumns

        Case DataSelectionTypes.dstDuplicateContacts
          mvResultColumns = "ContactNumber,ContactName,Title,Initials,Surname,Forenames,PreferredForename,Honorifics,Salutation,LabelName,Sex,DateOfBirth,Department,ContactType,StatusCode,AddressNumber,Address,HouseName,Town,County,Postcode,Country,OwnershipAccessLevel,OwnershipGroup"

        Case DataSelectionTypes.dstDuplicateOrganisations
          mvResultColumns = "OrganisationNumber,Name,Abbreviation,Department,StatusCode,AddressNumber,Address,HouseName,Town,County,Postcode,Country,OwnershipAccessLevel,OwnershipGroup"

        Case DataSelectionTypes.dstEMailContacts
          mvResultColumns = "ContactNumber,ContactName,Title,Initials,Surname,Forenames,PreferredForename,Honorifics,Salutation,LabelName,Sex,DateOfBirth,Department,ContactType,StatusCode,AddressNumber,Address,HouseName,Town,County,Postcode,Country,EMailAddress,DeviceCode,DeviceDesc,Position,Name,OwnershipGroup"
          mvSelectColumns = "ContactName,EMailAddress,DeviceDesc,Position,Name,ContactNumber"
          mvHeadings = DataSelectionText.String22863     'Name,EMail Address,Type,Position,Organisation,Number
          mvWidths = "2000,2000,1200,2000,2000,1"

        Case DataSelectionTypes.dstEMailOrganisations
          'BR11685/11700 - Changes result columns to match required data
          'SDT Restored to previous list
          mvResultColumns = "ContactNumber,ContactName,Title,Initials,Surname,Forenames,PreferredForename,Honorifics,Salutation,LabelName,Sex,DateOfBirth,Department,ContactType,StatusCode,AddressNumber,Address,HouseName,Town,County,Postcode,Country,EMailAddress,DeviceCode,DeviceDesc,Position,Name,OwnershipGroup"
          mvSelectColumns = "ContactName,EMailAddress,DeviceDesc,Position,Name,ContactNumber"
          mvHeadings = DataSelectionText.String22863     'Name,EMail Address,Type,Position,Organisation,Number
          mvWidths = "2000,2000,1200,2000,2000,1"

        Case DataSelectionTypes.dstPurchaseOrderDetails
          mvResultColumns = "PurchaseOrderNumber,LineNumber,Item,Price,Quantity,Amount,Balance,NominalAccount,DistributionCode"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then mvResultColumns = mvResultColumns & ",Product,Warehouse,ProductDesc,WarehouseDesc"
          mvSelectColumns = "LineNumber,Item,Price,Quantity,Amount,Balance,NominalAccount,DistributionCode"
          mvHeadings = DataSelectionText.String17994     'Line Number,Item,Price,Qty,Amount,Balance,Account,Distribution Code
          mvWidths = "700,2500,900,600,1000,1000,1500,900"
          mvDescription = DataSelectionText.String17995    'Purchase Order Details
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvCode = "POD"

        Case DataSelectionTypes.dstPurchaseOrderPayments
          mvResultColumns = "PurchaseOrderNumber,PaymentNumber,DueDate,LatestExpected,Amount,Percentage,Authorisation,AuthorisationStatus,AuthorisedBy,AuthorisedOn,PostedOn,PayeeContactNumber,PayeeAddressNumber,ChequeProducedOn,ChequeReferenceNumber,PayByBacs,PayeeReference,NoPaymentRequired,PaymentType,PaymentTypeDesc,NominalAccount,DistributionCode,SeparatePayment,ExpectedReceipt,ExpectedReceiptAmount,ExpectedReceiptReason,ReceiptForPaymentNumber,AuthorisationStatusDesc,PayeeContactLabelName,PayeeContactName,PayeeContactAddressLine,AdjustmentStatus,CancellationReason,CancellationSource,CancelledBy,CancelledOn,CancellationSourceDesc,CancellationReasonDesc,PopPaymentMethod,PopPaymentMethodDesc,DetailAmount,ChequeAmount,PopProductionNumber,DistributionCodeDesc"
          mvSelectColumns = "PaymentNumber,DueDate,LatestExpected,Amount,Percentage,Authorisation,AuthorisationStatus,AuthorisedBy,AuthorisedOn,PostedOn,ChequeProducedOn"
          mvHeadings = DataSelectionText.String17998     'Payment,Due,Expected,Amount,Percentage,Authorisation Required,Status,Authorised by,on,Posted,Payment Produced On
          mvWidths = "800,1100,1100,1000,900,1275,1785,1200,1200,1200,2000"
          mvDescription = DataSelectionText.String17999    'Purchase Order Payments
          mvDisplayTitle = DataSelectionText.String17965     'Payments
          mvCode = "POP"
          mvMaintenanceDesc = DataSelectionText.String17967   'Payment
          mvRequiredItems = "AuthorisedBy,PostedOn,ReceiptForPaymentNumber,Amount,ExpectedReceiptAmount,NoPaymentRequired,ExpectedReceipt,AdjustmentStatus,ChequeProducedOn,PaymentType,NominalAccount,DistributionCode,DueDate"

        Case DataSelectionTypes.dstPurchaseOrderInformation
          mvResultColumns = "PurchaseOrderNumber,PurchaseInvoiceNumber,ContactNumber,AddressNumber,Amount,OutputGroup,PurchaseOrderType,PurchaseOrderDesc,PayeeContactNumber,PayeeAddressNumber,StartDate,PaymentFrequency,NumberOfPayments,DistributionMethod,PaymentAsPercentage,Campaign,Appeal,Segment,CancellationReason,CancellationSource,CancelledBy,CancelledOn,ChequeReferenceNumber,BacsProcessed,CurrencyCode,PaymentSchedule,AdHocPayments,RegularPayments"
          mvSelectColumns = "PurchaseOrderNumber,ContactNumber,AddressNumber,Amount,OutputGroup,PurchaseOrderType,PurchaseOrderDesc,PayeeContactNumber,PayeeAddressNumber,StartDate,PaymentFrequency,NumberOfPayments,DistributionMethod,PaymentAsPercentage,Campaign,Appeal,Segment"
          mvHeadings = "PurchaseOrderNumber,ContactNumber,AddressNumber,Amount,OutputGroup,PurchaseOrderType,PurchaseOrderDesc,PayeeContactNumber,PayeeAddressNumber,StartDate,PaymentFrequency,NumberOfPayments,DistributionMethod,PaymentAsPercentage,Campaign,Appeal,Segment"
          mvWidths = "900,1200,2000,900,1200,1200,1000,900,1200,1200,2000,1200,1200,1200,1000,1000,1000"
          mvDescription = "Purchase Order Information"
          mvCode = "POI"
          mvRequiredItems = "CancelledBy,CancellationReason,CancelledOn,BacsProcessed,CurrencyCode,PaymentSchedule,AdHocPayments,RegularPayments"

        Case DataSelectionTypes.dstChequeInformation
          mvResultColumns = "ChequeReferenceNumber,ChequeNumber,ContactNumber,AddressNumber,Amount,PrintedOn,ReconciledOn,ChequeStatus,ChequeStatusDesc,AllowReissue"
          mvSelectColumns = "ChequeReferenceNumber,ChequeNumber,ContactNumber,AddressNumber,Amount,PrintedOn,ReconciledOn,ChequeStatus"
          mvHeadings = "Reference Number,Cheque Number,Contact Number,Address Number,Amount,Printed On,Reconciled On,Cheque Status"
          mvWidths = "1000,1000,1200,1200,900,900,900,900"
          mvDescription = "Cheque Information"
          mvCode = "POC"

        Case DataSelectionTypes.dstPurchaseInvoiceInformation
          mvResultColumns = "PurchaseInvoiceNumber,PurchaseOrderNumber,ContactNumber,AddressNumber,Amount,PayeeContactNumber,PayeeAddressNumber,PayeeReference,PurchaseInvoiceDate,ChequeReferenceNumber,Source,Campaign,Appeal,Segment,CurrencyCode"
          mvSelectColumns = "PurchaseInvoiceNumber,PurchaseOrderNumber,ContactNumber,AddressNumber,Amount,PayeeContactNumber,PayeeAddressNumber,PayeeReference,PurchaseInvoiceDate,ChequeReferenceNumber,Source,Campaign,Appeal,Segment,CurrencyCode"
          mvHeadings = "PurchaseInvoiceNumber,PurchaseOrderNumber,ContactNumber,AddressNumber,Amount,PayeeContactNumber,PayeeAddressNumber,PayeeReference,PurchaseInvoiceDate,ChequeReferenceNumber,Source,Campaign,Appeal,Segment,Currency Code"
          mvWidths = "900,1200,2000,900,1200,1200,1000,900,1200,1200,2000,1200,1200,1200,1200"
          mvDescription = "Purchase Invoice Information"
          mvCode = "PII"

        Case DataSelectionTypes.dstUnauthorisedPOPayments
          mvResultColumns = "DueDate,LatestExpected,Amount,Percentage,PurchaseOrderNumber,PaymentNumber,AuthorisationStatus,AuthorisationStatusDesc,ContactNumber,ContactName,PayeeContactNumber,PayeeContactName,PayByBacs,PurchaseOrderType,PurchaseOrderTypeDesc,PayeeContactLabelName,PayeeContactSalutation,PayeeContactAddressLine,PopPaymentMethod,PopPaymentMethodDesc,BankAccount"
          mvSelectColumns = "ContactNumber,DueDate,LatestExpected,Amount,Percentage,ContactName,AuthorisationStatusDesc,PurchaseOrderNumber,PaymentNumber,PayeeContactName"
          mvHeadings = DataSelectionText.String40002
          mvWidths = "1,1100,1100,1000,900,2000,3000,1200,1200,1200"
          mvDescription = "Unauthorised Purchase Order Payments"
          mvDisplayTitle = DataSelectionText.String40001
          mvCode = "UPOP"
          mvRequiredItems = "PayeeContactName"

        Case DataSelectionTypes.dstPurchaseInvoiceDetails
          mvResultColumns = "PurchaseInvoiceNumber,LineNumber,Item,Price,Quantity,Amount,NominalAccount,DistributionCode,AdjustmentStatus,CancellationReason,CancellationSource,CancelledBy,CancelledOn"
          mvSelectColumns = "LineNumber,Item,Price,Quantity,Amount,NominalAccount,DistributionCode"
          mvHeadings = DataSelectionText.String18022     'Line Number,Item,Price,Qty,Amount,Account,Distribution Code
          mvWidths = "700,2500,900,600,1000,1500,900"
          mvDescription = DataSelectionText.String18023    'Purchase Invoice Details
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvCode = "PID"

        Case DataSelectionTypes.dstDocumentSubjects
          mvResultColumns = "TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,Quantity,Primary,ActivityCode,ActivityValueCode,ActivityDuration,AmendedOn,AmendedBy"
          mvSelectColumns = "TopicDesc,SubTopicDesc,Quantity,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18024     'Topic,Sub Topic,Quantity,Amended by,Amended on
          mvWidths = "2000,2000,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18025    'Document Analysis
          mvDisplayTitle = DataSelectionText.String17991     'Analysis
          mvMaintenanceDesc = "Analysis"
          mvCode = "DOS"
          mvRequiredItems = "TopicCode,SubTopicCode,Primary"

        Case DataSelectionTypes.dstDocumentHistory
          mvResultColumns = "ActionDate,ActionTime,Action,UserName,Notes"
          mvSelectColumns = "Action,UserName,ActionDate,ActionTime,Notes"
          mvHeadings = DataSelectionText.String18026     'Action taken,By,On,At,Notes
          mvWidths = "1500,1200,1200,1200,3000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18027    'Document History
          mvDisplayTitle = DataSelectionText.String18028     'History
          mvCode = "DH"

        Case DataSelectionTypes.dstDocumentLinks
          mvResultColumns = "LinkType,ContactNumber,AddressNumber,ContactName,Notified,Processed,ContactType,LinkTypeDesc,EntityType,EntityTypeDesc"
          mvSelectColumns = "LinkTypeDesc,ContactName,Notified,Processed"
          mvHeadings = DataSelectionText.String18029     'Link,To,Notified,Processed
          mvWidths = "1500,2000,1000,1000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18030    'Document Links
          mvDisplayTitle = DataSelectionText.String18031     'Links
          mvMaintenanceDesc = "Document Link"
          mvRequiredItems = "LinkType,ContactType,EntityType"
          mvCode = "DL"

        Case DataSelectionTypes.dstCommunicationsLogDocClass
          mvResultColumns = "DocumentClass,StandardDocument,InUseby,CreatedBy,Dated,Package"
          mvSelectColumns = "DocumentClass,StandardDocument,InUseby,CreatedBy,Dated,Package"
          mvHeadings = "DocumentClass,StandardDocument,InUseby,CreatedBy,Dated,Package"   ' TODO: Move to DataSelectionText   'Select,Document,Our Reference,Subject
          mvWidths = "1200,1200,1800,1200,1200,1"
          mvRequiredItems = "StandardDocument"

        Case DataSelectionTypes.dstMeetings
          mvResultColumns = "MeetingNumber,MeetingDesc,MeetingDate,MeetingType,MeetingLocation,DurationDays,DurationHours,DurationMinutes,Preamble,Notes,Agenda,CommunicationsLogNumber,MasterAction,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataOwnerContactNumber) Then mvResultColumns = mvResultColumns & ",OwnerContactNumber"
          mvSelectColumns = "MeetingNumber,MeetingDesc,MeetingDate,MeetingType,MeetingLocation,DurationDays,DurationHours,DurationMinutes,Preamble,Notes,Agenda,CommunicationsLogNumber,MasterAction,AmendedBy,AmendedOn"
          mvHeadings = "MeetingNumber,MeetingDesc,MeetingDate,MeetingType,MeetingLocation,DurationDays,DurationHours,DurationMinutes,Preamble,Notes,Agenda,CommunicationsLogNumber,MasterAction,AmendedBy,AmendedOn"   ' TODO: Move to DataSelectionText   'Select,Document,Our Reference,Subject
          mvWidths = "1200,1200,1800,1200,1200,1,1,1,1200,1200,1,1,1,1,1"
          mvRequiredItems = "MeetingNumber"

        Case DataSelectionTypes.dstContactMeetings
          mvResultColumns = "MeetingNumber,MeetingDesc,MeetingDate,MeetingType,MeetingTypeDesc,MeetingLocation,MeetingLocationDesc,DurationDays,DurationHours,DurationMinutes,Preamble,Notes,Agenda,CommunicationsLogNumber,MasterAction,AmendedBy,AmendedOn,OwnerContactNumber,OwnerContactName"
          mvSelectColumns = "MeetingNumber,MeetingDesc,MeetingDate,MeetingTypeDesc,MeetingLocationDesc,DurationDays,DurationHours,DurationMinutes,Preamble,Notes,Agenda,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.ContactMeetingsHeadings   'Select,Document,Our Reference,Subject
          mvWidths = "1200,1200,1800,1200,1200,1,1,1,1200,1200,1,1,1,1,1"
          mvRequiredItems = "MeetingNumber"
          mvMaintenanceDesc = "Meeting"
          mvDescription = DataSelectionText.ContactMeetingsDescription
          mvCode = "CME"

        Case DataSelectionTypes.dstDocumentRelatedDocuments
          mvResultColumns = "Select,DocumentNumber,OurReference,Subject,DocumentSource,WordProcessorDocument,Precis,Extension"
          mvSelectColumns = "Select,DocumentNumber,OurReference,Subject"
          mvHeadings = DataSelectionText.String18032     'Select,Document,Our Reference,Subject
          mvWidths = "1200,1200,1800,2000"
          mvRequiredItems = "DocumentSource,WordProcessorDocument,Precis,Extension"

        Case DataSelectionTypes.dstDocumentContactLinks
          mvResultColumns = "LinkType,ContactNumber,AddressNumber,ContactName,Notified,Processed"

        Case DataSelectionTypes.dstDocumentOrganisationLinks
          mvResultColumns = "LinkType,ContactNumber,AddressNumber,ContactName,Notified,Processed"

        Case DataSelectionTypes.dstDocumentDocumentLinks
          mvResultColumns = "DocumentNumber,DocumentReference"

        Case DataSelectionTypes.dstDocumentTransactionLinks
          mvResultColumns = "DocumentNumber,TransactionReference"

        Case DataSelectionTypes.dstActionLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified,ContactType,LinkTypeDesc,EntityType,EntityTypeDesc"
          mvSelectColumns = "LinkTypeDesc,ContactName,Notified"
          mvHeadings = DataSelectionText.String18033     'Link,To,Notified
          mvWidths = "1500,2000,1000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18034    'Action Links
          mvDisplayTitle = DataSelectionText.String18031     'Links
          mvMaintenanceDesc = "Action Link"
          mvRequiredItems = "LinkType,ContactName,ContactType,EntityType"
          mvCode = "ACL"

        Case DataSelectionTypes.dstActionContactLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified"

        Case DataSelectionTypes.dstActionOrganisationLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName"

        Case DataSelectionTypes.dstActionDocumentLinks
          mvResultColumns = "DocumentNumber,DocumentReference"

        Case DataSelectionTypes.dstMeetingLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified,Attended,MeetingRoleCode,MeetingRoleDesc,ContactType,LinkTypeDesc"
          mvSelectColumns = "LinkTypeDesc,ContactName,Notified,Attended,MeetingRoleCode"
          mvHeadings = "Link,To,Notified,Attended,Role"     'Link,To,Notified,Processed
          mvWidths = "1500,2000,1000,1000,1000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Meeting Links"
          mvDisplayTitle = DataSelectionText.String18031     'Links
          mvMaintenanceDesc = "Meeting Link"
          mvRequiredItems = "LinkType,ContactType"
          mvCode = "CML"

        Case DataSelectionTypes.dstMeetingContactLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified,Attended,MeetingRoleCode,MeetingRoleDesc"

        Case DataSelectionTypes.dstMeetingOrganisationLinks
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified,Attended,MeetingRoleCode,MeetingRoleDesc"

        Case DataSelectionTypes.dstMeetingDocumentLinks
          mvResultColumns = "DocumentNumber,DocumentReference"

        Case DataSelectionTypes.dstEventSessions
          mvResultColumns = "SessionNumber,SessionDesc,SessionTypeCode,SessionTypeDesc,SubjectCode,SubjectDesc,SkillLevelCode,SkillLevelDesc,StartDate,EndDate,StartTime,EndTime,Location,MinimumAttendees,MaximumAttendees,TargetAttendees,NumberInterested,NumberOfAttendees,NumberOnWaitingList,MaximumOnWaitingList,Notes,VenueBookingNumber,AmendedBy,AmendedOn,Available,CPDApprovalStatus,CPDAwardingBody,CPDCategory,CPDDateApproved,CPDNotes,CPDPoints,CPDYear,WaitingAvailable,LongDescription,ExternalAppointmentId"
          mvSelectColumns = "SessionDesc,NumberInterested,NumberOfAttendees,NumberOnWaitingList,Available,WaitingAvailable,StartDate,StartTime,EndDate,EndTime,ExternalAppointmentId,SubjectDesc,Location,MaximumAttendees,SkillLevelDesc,LongDescription,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18035     'Session,Interested,Booked,Waiting,Available,Waiting Available,Start Date,Start Time,End Date,End Time,External Appointment Id,Amended By,Amended On
          mvWidths = "2700,1100,1100,1100,1000,1550,1100,1000,1100,1000,1000,1000,1300,1300"
          mvDescription = DataSelectionText.String18037    'Event Sessions
          mvCode = "ESS"
          mvRequiredItems = "NumberInterested,NumberOfAttendees,NumberOnWaitingList,SubjectDesc,Location,MaximumAttendees,SkillLevelDesc,LongDescription,ExternalAppointmentId"

        Case DataSelectionTypes.dstEventBookingOptions
          mvResultColumns = "OptionNumber,OptionDesc,PickSessions,NumberOfSessions,DeductFromEvent,MinimumBookings,MaximumBookings,ProductCode,ProductDesc,RateCode,RateDesc,AmendedBy,AmendedOn,IssueEventResources,BookingCount,LongDescription,FreeOfCharge"
          mvSelectColumns = "OptionNumber,OptionDesc,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String25511 '"Option Number,Option,Amended By,Amended On"
          mvWidths = "1,4800,1250,1250"
          mvDescription = DataSelectionText.String18038    'Event Booking Options
          mvCode = "EBO"
          mvRequiredItems = "OptionNumber,BookingCount,ProductCode,PickSessions,NumberOfSessions,MinimumBookings,MaximumBookings,FreeOfCharge,RateCode"

        Case DataSelectionTypes.dstEventSessionTests
          mvResultColumns = "SessionNumber,TestNumber,TestDesc,GradeDataType,MinimumValue,MaximumValue,Pattern,AmendedBy,AmendedOn,GradeDataTypeDesc"
          mvSelectColumns = "TestNumber,TestDesc,GradeDataType,MinimumValue,MaximumValue,Pattern,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18039     'Number,Description,Grade Data Type,Minimum Value,Maximum Value,Pattern,Amended By,Amended On
          mvWidths = "1000,2000,1400,1200,1200,1500,1300,1300"
          mvDescription = DataSelectionText.String18041    'Event Session Tests
          mvCode = "EST"

        Case DataSelectionTypes.dstEventSubmissions
          mvResultColumns = "ContactNumber," & ContactNameResults() & ",PaperTitle,Submitted,AmendedBy,AmendedOn,SubmissionNumber,Address,AddressNumber,SubjectCode,SubjectDesc,SkillLevelCode,SkillLevelDesc,Forwarded,Assessor,Returned,SubmissionResult,AssessorName"
          mvSelectColumns = "PaperTitle,Submitted,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18042     'Title,Submitted,Amended By,Amended On
          mvWidths = "4000,1400,1200,1200"
          mvDescription = DataSelectionText.String18044    'Event Submissions
          mvCode = "EVS"

        Case DataSelectionTypes.dstEventAccommodation
          mvResultColumns = "BlockBookingNumber,RoomTypeDesc,FromDate,ToDate,NumberOfRooms,NightsAvailable,Organisation,Address,RackRate,AgreedRate,BookedOn,ReleaseDate,ProductCode,ProductDesc,ConfirmedOn,RateCode,RateDesc,AmendedBy,AmendedOn,OrganisationNumber,AddressNumber,Notes,RoomType,ContactNumber," & ContactNameResults()
          mvSelectColumns = "RoomTypeDesc,FromDate,ToDate,NumberOfRooms,NightsAvailable,Organisation,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18045     'Type,From,To,Booked,Nights,Name,Amended By,Amended On
          mvWidths = "1200,1200,1200,900,900,3000,1200,1200"
          mvDescription = DataSelectionText.String18047    'Event Accommodation
          mvCode = "EVA"

        Case DataSelectionTypes.dstEventResources
          mvResultColumns = "ResourceNumber,ResourceDesc,ResourceType,QuantityRequired,QuantityIssued,Allocated,AmendedBy,AmendedOn,Session,SessionNumber,Product,CopyTo,IssueBasis,DespatchTo,IssueDate,Notes,ResourceTypeCode,OrganisationNumber,AddressNumber,ContactNumber,ExternalResourceType,ObtainedOn,ReturnBy,ReturnedOn,TotalAmount,DueDate,Deposit,DepositPaidDate,InternalProductDesc,StandardProductDesc,InternalResourceName"
          mvSelectColumns = "ResourceDesc,ResourceType,QuantityRequired,QuantityIssued,Allocated,AmendedBy,AmendedOn,Session"
          mvHeadings = DataSelectionText.String18048     'Resource Description,Resource Type,Qty,Qty Issued,Allocated,Amended By,Amended On,Session Description
          mvWidths = "2000,1200,1000,1000,1000,1000,1000,3000"
          mvDescription = DataSelectionText.String18050    'Event Resources
          mvCode = "ERS"
          mvRequiredItems = "Product,CopyTo,ResourceTypeCode,StandardProductDesc"

        Case DataSelectionTypes.dstEventVenueBookings
          mvResultColumns = "VenueNumber,Venue,Description,Reference,TotalAmount,DueDate,Deposit,DepositPaidDate,Balance,BalancePaidDate,ConfirmedBy,ConfirmedOn,AmendedBy,AmendedOn,OrganisationNumber,ContactNumber,ContactName,Telephone,OrganisationName,Address"
          mvSelectColumns = "Venue,Description,Reference,TotalAmount,DueDate,Deposit,DepositPaidDate,Balance,BalancePaidDate,ConfirmedBy,ConfirmedOn,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18051     'Venue,Description,Reference,Total,Due Date,Deposit,Paid Date,Balance,Paid Date,Confirmed By,Confirmed On,Amended By,Amended On
          mvWidths = "900,2000,2000,1000,1200,1000,1200,1000,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String18053    'Event Venue Bookings
          mvCode = "EVB"

        Case DataSelectionTypes.dstEventPersonnel
          mvResultColumns = "ContactNumber," & ContactNameResults() & ",SessionNumber,ConfirmedOn,ConfirmedBy,Expenses,ExpensesReceived,AuthorisedBy,AuthorisedOn,AmendedBy,AmendedOn,SessionDesc,Task,Position,OrganisationName,StartDate,StartTime,EndDate,EndTime,AddressNumber,ChairPerson,EventPersonnelNumber,IssueEventResources"
          mvSelectColumns = "ContactName,ConfirmedBy,ConfirmedOn,Expenses,ExpensesReceived,AuthorisedBy,AuthorisedOn,AmendedBy,AmendedOn,SessionDesc,ContactNumber,SessionNumber"
          mvHeadings = DataSelectionText.String26711 ' "Name,Confirmed By,Confirmed On,Claimed,Received,Authorised By,Authorised On,Amended By,Amended On,Session,Contact Number,Session Number"
          mvWidths = "2500,1250,1250,1250,1250,1250,1250,1250,1250,3500,1,1"
          mvDescription = DataSelectionText.String18054    'Event Personnel
          mvCode = "EPL"
          mvRequiredItems = "ContactName"

        Case DataSelectionTypes.dstEventPIS
          mvResultColumns = "ContactName,DelegateSurname,DelegateForenames,DelegateTitle,DelegateInitials,EventNumber,EventPisNumber,PisNumber,EventDelegateNumber,ContactNumber,IssueDate,Amount,BankedBy,BankedOn,ReconciledOn,BankedByContactName"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvSelectColumns = "PisNumber,IssueDate,ContactName,Amount,BankedByContactName,BankedOn,ReconciledOn"
          mvHeadings = DataSelectionText.String22864     'Pis Number,Issue Date,Delegate,Amount,Banked By,Banked On,Reconciled On
          mvWidths = "1200,1200,1200,1200,2000,1200,1200"
          mvDescription = DataSelectionText.String18273    'Event Paying In Slips
          mvDisplayTitle = DataSelectionText.String18273     'Event Paying In Slips
          mvCode = "EPIS"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstEventBookingOptionSessions
          mvResultColumns = "SessionNumber,SessionDesc,SubjectCode,SubjectDesc,SkillLevelCode,SkillLevelDesc,StartDate,EndDate,StartTime,EndTime,Allocation,Location,LongDescription,PlacesAvailable,AmendedBy,AmendedOn"
          mvSelectColumns = "SessionDesc,Allocation"
          mvHeadings = DataSelectionText.String22865     'Session,Allocation
          mvDescription = "Event Booking Option Session"
          mvWidths = "2000,900"
          mvRequiredItems = "SessionNumber,Allocation,PlacesAvailable"
          mvCode = "EOPS"

        Case DataSelectionTypes.dstEventBookingSessions
          mvResultColumns = "SessionNumber,SessionDesc,SubjectCode,SubjectDesc,SkillLevelCode,SkillLevelDesc,StartDate,EndDate,StartTime,EndTime"
          mvSelectColumns = "SessionDesc,StartDate,StartTime,EndDate,EndTime"
          mvHeadings = DataSelectionText.String22866     'Session,Start Date,Time,End Date,Time
          mvWidths = "2000,900,900,900,900"

        Case DataSelectionTypes.dstEventAttendees
          mvResultColumns = "BookingNumber,ContactNumber,AddressNumber,DelegateName,BookerName,Attended,Position,OrganisationName,CandidateNumber,BookingStatus,Quantity,DelegateNumber,PledgedAmount,DonationTotal,SponsorshipTotal,BookingPaymentAmount,OtherPaymentsTotal,SequenceNumber,PayerContactNumber,PayerName"
          mvSelectColumns = "BookerName,DelegateName,BookingStatus,Attended,CandidateNumber,Position,OrganisationName,DelegateNumber"
          mvHeadings = DataSelectionText.String18055     'Booked By,Name,Booking Status,Attended,Candidate,Position,Organisation,DelegateNumber
          mvWidths = "3250,3360,2300,840,1000,2000,3000,0"
          mvDescription = DataSelectionText.String18059    'Event Attendees
          mvCode = "EAT"
          mvRequiredItems = "BookingStatus,Quantity,BookerName,DelegateName,OrganisationName"

        Case DataSelectionTypes.dstEventCurrentAttendees
          mvResultColumns = "BookingNumber,ContactNumber,AddressNumber,DelegateName,BookerName,Attended,Position,OrganisationName,CandidateNumber,BookingStatus,Quantity,DelegateNumber,SessionNumber,SessionType,PledgedAmount,DonationTotal,SponsorshipTotal,BookingPaymentAmount,OtherPaymentsTotal"
          mvSelectColumns = "BookingNumber,AddressNumber,ContactNumber,DelegateName,Attended,BookingStatus"
          mvHeadings = DataSelectionText.String25703   'Booking Number,Address Number,Contact_number,Name,Attended,Att
          mvWidths = "1,1,1,4500,900,1"
          mvRequiredItems = "BookingStatus,Quantity,BookerName,DelegateName,OrganisationName,SessionNumber,SessionType,Attended"

        Case DataSelectionTypes.dstIncentives
          mvResultColumns = "IncentiveScheme,ReasonForDespatch,SequenceNumber,Product,Rate,ForWhom,IncentiveType,IncentiveDesc,Quantity,Basic,DespatchMethod,CurrentPrice,FuturePrice,PriceChangeDate,VatRate,Percentage,ProductDesc,Subscription,ThankYouLetter,VatExclusive,OriginalQuantity,IgnoreProductAndRate,MinimumQuantity,MaximumQuantity"
          mvSelectColumns = "IncentiveDesc,Quantity"
          mvHeadings = "Incentive,Quantity"
          mvWidths = "1200,1200"
          mvRequiredItems = mvResultColumns

        Case DataSelectionTypes.dstFulFilledContactIncentives
          mvResultColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,DateResponded,DateFulfilled"
          mvSelectColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,DateResponded,DateFulfilled"
          mvHeadings = "Contact Number,Recipient,Product Desc,Quantity,Source,Date Created,Date Fulfilled"
          mvWidths = "1,2000,2000,900,1500,1200,1200"
          mvRequiredItems = mvResultColumns

        Case DataSelectionTypes.dstUnFulFilledContactIncentives
          mvResultColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,DateResponded"
          mvSelectColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,DateResponded"
          mvHeadings = "Contact Number,Recipient,Product Desc,Quantity,Source,Date Created"
          mvWidths = "1,2000,2000,900,1500,1200"
          mvRequiredItems = mvResultColumns

        Case DataSelectionTypes.dstFulFilledPayPlanIncentives
          mvResultColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,ReasonForDispatchDesc,DateResponded,DateFulfilled"
          mvSelectColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,ReasonForDispatchDesc,DateResponded,DateFulfilled"
          mvHeadings = "Contact Number,Recipient,Product Desc,Quantity,Source,Despatch Reason,Date Created,Date Fulfilled"
          mvWidths = "1,2000,2000,900,1500,1600,1200,1200"
          mvRequiredItems = mvResultColumns

        Case DataSelectionTypes.dstUnFulFilledPayPlanIncentives
          mvResultColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,ReasonForDispatchDesc,DateResponded"
          mvSelectColumns = "ContactNumber,LabelName,ProductDesc,Quantity,SourceDesc,ReasonForDispatchDesc,DateResponded"
          mvHeadings = "Contact Number,Recipient,Product Desc,Quantity,Source,Despatch Reason,Date Created"
          mvWidths = "1,2000,2000,900,1500,1600,1200"
          mvRequiredItems = mvResultColumns

        Case DataSelectionTypes.dstPaymentPlanAmendmentHistory
          mvResultColumns = "OperationDate,ChangeType,OldValue,NewValue,OldBalance,NewBalance"
          mvSelectColumns = mvResultColumns
          mvHeadings = DataSelectionText.String18274       'Date,Type,Old Value,New Value,Old Balance,New Balance
          mvWidths = "1100,2000,2000,2000,1500,1500"
          mvDescription = DataSelectionText.String18683    'Payment Plan Amendment History
          mvDisplayTitle = DataSelectionText.String18684   'Amendments
          mvCode = "PAH"

        Case DataSelectionTypes.dstPaymentPlanPayments
          mvResultColumns = "ScheduledPaymentNumber,AmountDue,AmountOutstanding,RevisedAmount,DueDate,ClaimDate,ScheduledPaymentStatusDesc,ExpectedBalance,ScheduleCreationReasonDesc,PaymentNumber,Amount,Balance,Posted,Status,BatchNumber,TransactionNumber,LineNumber,PaymentMethodDesc,TransactionDate,PostedDate,ScheduledPaymentStatus,WriteOffLineAmount,InvoicePayStatus,InvoicePayStatusDesc,PayerContactNumber,ScheduleCreationReason"
          mvSelectColumns = "DueDate,ClaimDate,AmountDue,RevisedAmount,Amount,AmountOutstanding,ScheduledPaymentStatusDesc,TransactionDate,PaymentMethodDesc,PostedDate,BatchNumber,TransactionNumber,LineNumber,Balance,Status,ScheduleCreationReasonDesc,ScheduledPaymentStatus,ScheduleCreationReason"
          mvHeadings = DataSelectionText.String18060     'Due Date,Claim Date,Due Amount,Revised Amount,Amount Paid,Amount Outstanding,Payment Status,Transaction Date,Payment Method,Posted Date,Batch,Trans-action,Line,Payplan Balance,Status,Creation Reason
          mvWidths = "1100,1100,1000,1000,1000,1000,1500,1100,1300,1100,900,600,600,1000,600,2000"
          mvHeaderLines = 2
          mvDescription = DataSelectionText.String18061    'Payment Plan Payments
          mvDisplayTitle = DataSelectionText.String17965     'Payments
          mvCode = "PPP"
          mvRequiredItems = "ScheduledPaymentNumber,ScheduledPaymentStatus,Status,PayerContactNumber,ScheduleCreationReason"

        Case DataSelectionTypes.dstPaymentPlanDetails
          mvResultColumns = "ContactNumber,AddressNumber,Product,ProductDesc,Rate,RateDesc,DistributionCode,Quantity,Amount,Balance,Arrears,DespatchMethod,Source,ProductNumber,AmendedOn,AmendedBy,CreatedBy,CreatedOn,CurrencyCode,DistributionCodeDesc,VATExclusive,"
          mvResultColumns &= "ValidFrom,ValidTo,ModifierActivity,ModifierActivityDesc,ModifierActivityValue,ModifierActivityValueDesc,ModifierActivityQuantity,ModifierActivityDate,ModifierPrice,ModifierPerItem,UnitPrice,ProRated,NetAmount,VATAmount,GrossAmount,VATRate,VATPercentage," & ContactNameResults()
          mvSelectColumns = "Product,ProductDesc,RateDesc,Quantity,Amount,Balance,Arrears,Source,DistributionCode,DistributionCodeDesc,DespatchMethod,ProductNumber,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18062     'Product Code,Product,Rate,Qty,Amount,Balance,Arrears,Source,Distribution Code,Despatch Method,Product Number,Amended by,on
          mvWidths = "1500,3000,2000,500,1200,1200,1200,1200,1200,2000,2000,1200,1200,1200"
          mvDescription = DataSelectionText.String18063    'Payment Plan Details
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvRequiredItems = "DistributionCode"
          mvCode = "PPD"

        Case DataSelectionTypes.dstMembershipPaymentPlanDetails
          mvResultColumns = "ContactNumber,AddressNumber,Product,ProductDesc,Rate,RateDesc,DistributionCode,Quantity,Amount,Balance,Arrears,DespatchMethod,Source,ProductNumber,AmendedOn,AmendedBy,CurrencyCode,DistributionCodeDesc,VATExclusive,"
          mvResultColumns &= "ValidFrom,ValidTo,ModifierActivity,ModifierActivityDesc,ModifierActivityValue,ModifierActivityValueDesc,ModifierActivityQuantity,ModifierActivityDate,ModifierPrice,ModifierPerItem,UnitPrice,ProRated,NetAmount,VATAmount,GrossAmount,VATRate,VATPercentage,"
          mvResultColumns &= "SubsValidFrom,SubsValidTo,CancelledBy,CancelledOn,CancellationReason,CancellationSource,ReasonForDespatch,CancellationReasonDesc,CancellationSourceDesc,ReasonForDespatchDesc"    'These columns must be at the end
          mvSelectColumns = "Product,ProductDesc,RateDesc,Quantity,Amount,Balance,Arrears,Source,DistributionCode,DespatchMethod,ProductNumber,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18700     'Product Code,Product,Rate,Qty,Amount,Balance,Arrears,Source,Distribution Code,Despatch Method,Product Number,Amended by,on
          mvWidths = "1500,3000,2000,500,1200,1200,1200,1200,2000,2000,1200,1200,1200"
          mvDescription = DataSelectionText.String18064    'Membership Payment Plan Details
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvCode = "MPD"

        Case DataSelectionTypes.dstPaymentPlanSubscriptions
          mvResultColumns = "SubscriptionNumber,ContactNumber,AddressNumber,Product,ProductDesc,Quantity,DespatchMethod,DespatchMethodDesc,ReasonForDespatch,ReasonForDespatchDesc,ValidFrom,ValidTo,CancellationReason,CancelledBy,CommunicationNumber,CancellationSource,CancelledOn,AddressLine,DeliverTo,CancellationReasonDesc,CancellationSourceDesc"
          mvSelectColumns = "ProductDesc,Quantity,DespatchMethodDesc,ReasonForDespatchDesc,ValidFrom,ValidTo,CancelledBy,CancelledOn,CancellationReason,DeliverTo"
          mvHeadings = DataSelectionText.String18065     'Product,Qty,Despatch Method,Despatch Reason,Valid From,Valid To,Cancelled by,On,Reason,Deliver To
          mvWidths = "3000,500,1680,1680,1200,1200,1200,1200,1600,2000"
          mvDescription = DataSelectionText.String18066    'Payment Plan Subscriptions
          mvDisplayTitle = DataSelectionText.String18067     'Subscriptions
          mvCode = "PPS"

        Case DataSelectionTypes.dstPaymentPlanMembers
          mvResultColumns = "OrderNumber,ContactNumber,AddressNumber,Source,AgeOverride,MemberNumber,MembershipType,MembershipTypeDesc,Joined,Branch,BranchName,CancelledOn,CancelledBy,AddressLine," & ContactNameResults()
          mvSelectColumns = "OrderNumber,ContactNumber,AddressNumber,Source,AgeOverride,MemberNumber,MembershipType,ContactName,MembershipTypeDesc,Joined,Branch,BranchName,CancelledOn,CancelledBy"
          mvRequiredItems = "OrderNumber,MembershipType,AddressNumber,ContactNumber,Branch,Source,Joined,AgeOverride"
          mvHeadings = mvEnv.GetBranchText(DataSelectionText.String22827)      'Payment Plan,Contact No,Address No,Member Number,Member Name,Type,Joined,Branch Code,Branch,Cancelled On,Cancelled By   '1,1,1200,2100,1200,2100,1200,1200
          mvWidths = "1,1,1,1,1,1200,1,1,2100,2100,1,2100,1200,1200"
          mvDescription = DataSelectionText.String18068    'Payment Plan Members
          mvDisplayTitle = DataSelectionText.String18069    'Members
          mvCode = "PPM"

        Case DataSelectionTypes.dstPaymentPlanOutstandingOPS
          mvResultColumns = "ScheduledPaymentNumber,DueDate,ClaimDate,AmountDue,AmountOutstanding,RevisedAmount,ExpectedBalance,ScheduledPaymentStatusDesc,ScheduleCreationReasonDesc,ScheduledPaymentStatus,ScheduleCreationReason"
          mvSelectColumns = "ScheduledPaymentNumber,DueDate,ClaimDate,AmountDue,AmountOutstanding,RevisedAmount,ExpectedBalance,ScheduledPaymentStatusDesc,ScheduleCreationReasonDesc,ScheduledPaymentStatus,ScheduleCreationReason"
          mvHeadings = DataSelectionText.String22832      'Payment Number,Due Date,Claim Date,Amount Due,Amount Outstanding,Revised Amount,Expected Balance,Payment Status,Creation Reason,Payment Status Code,Creation Reason Code
          mvWidths = "1,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String22834      'Payment Plan Outstanding Scheduled Payments
          mvDisplayTitle = DataSelectionText.String22835      'Outstanding Scheduled Payments
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices
          mvCode = "PPO"

        Case DataSelectionTypes.dstMembershipOtherMembers
          mvResultColumns = "ContactNumber,AddressNumber,MemberNumber,MembershipType,MembershipTypeDesc,Joined,BranchName,CancelledOn,CancelledBy,CancellationReason,CancellationSource,MembershipCardExpires,AmendedOn,AmendedBy,MembershipNumber," & ContactNameResults() & ",CancellationReasonDesc,CancellationSourceDesc"
          mvSelectColumns = "ContactName,MemberNumber,MembershipTypeDesc,Joined,MembershipCardExpires,CancelledBy,CancelledOn,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18070     'Name,Member Number,Type,Joined,Card Expires,Cancelled By,On,Amended By,On
          mvWidths = "1800,1200,2100,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String18071    'Membership Other Members
          mvDisplayTitle = DataSelectionText.String18072     'Other Members
          mvCode = "MOM"

        Case DataSelectionTypes.dstMembershipChanges
          mvResultColumns = "ContactNumber,AddressNumber,MemberNumber,MembershipType,MembershipTypeDesc,Joined,BranchName,CancelledOn,CancelledBy,CancellationReason,CancellationSource,AmendedOn,AmendedBy,MembershipNumber,Source,SourceDesc,CancellationReasonDesc,CancellationSourceDesc"
          mvSelectColumns = "MembershipTypeDesc,Joined,CancelledOn,CancelledBy,CancellationReasonDesc,MemberNumber"
          mvHeadings = DataSelectionText.String18073     'Type,Joined,Cancelled on,Cancelled by,Reason,Member Number
          mvWidths = "2100,1200,1200,1200,1800,1200"
          mvDescription = DataSelectionText.String18074    'Membership Changes
          mvDisplayTitle = DataSelectionText.String18075     'Previous Membership Types
          mvCode = "MCM"

        Case DataSelectionTypes.dstCovenantGiftAidClaims
          mvResultColumns = "ClaimGeneratedDate,ClaimNumber,NetAmount,AmountClaimed,PaymentNumber,TransactionDate,Amount,Balance,BatchNumber,TransactionNumber"
          mvSelectColumns = "ClaimGeneratedDate,ClaimNumber,NetAmount,AmountClaimed,PaymentNumber,TransactionDate,Amount,Balance,BatchNumber,TransactionNumber"
          mvHeadings = DataSelectionText.String18078     'Claim Date,Claim No.,Net Amount,Amount Claimed,Payment No.,Transaction Date,Amount,Balance,Batch No.,Transaction No.
          mvWidths = "1300,1200,1500,1500,1000,1400,1000,1000,1000,1050"
          mvDescription = DataSelectionText.String18079    'Covenant Gift Aid Claims
          mvDisplayTitle = DataSelectionText.String18080     'Gift Aid Claims
          mvCode = "CGAC"

        Case DataSelectionTypes.dstCovenentClaims
          mvResultColumns = "ClaimGenerated,ClaimNumber,StartPaymentNumber,EndPaymentNumber,NetAmount,AmountCalculated,AmountClaimed,AmendedOn,AmendedBy"
          mvSelectColumns = "ClaimGenerated,ClaimNumber,StartPaymentNumber,EndPaymentNumber,NetAmount,AmountCalculated,AmountClaimed,AmendedOn,AmendedBy"
          mvHeadings = DataSelectionText.String18083     'Date,Claim,First Payment,Last Payment,Net,Tax Calculated,Tax Claimed,Amended on,Amended by
          mvWidths = "1200,1100,900,900,1100,1200,1100,1200,1200"
          mvDescription = DataSelectionText.String18084    'Covenant Claims
          mvDisplayTitle = DataSelectionText.String18085     'Tax Claims
          mvCode = "COVC"

        Case DataSelectionTypes.dstCovenentPayments
          mvResultColumns = "PaymentNumber,TransactionDate,Amount,Balance,BatchNumber,TransactionNumber"
          mvSelectColumns = "PaymentNumber,TransactionDate,Amount,Balance,BatchNumber,TransactionNumber"
          mvHeadings = DataSelectionText.String18086     'Number,Date,Amount,Balance,Batch,Transaction
          mvWidths = "1000,1200,1000,1000,1000,1050"
          mvDescription = DataSelectionText.String18087     'Covenant Payments
          mvDisplayTitle = DataSelectionText.String17965     'Payments
          mvCode = "CVP"

        Case DataSelectionTypes.dstCPDDetails
          If mvParameters IsNot Nothing AndAlso mvParameters.Exists("CPDType") AndAlso mvParameters("CPDType").Value = "O" Then
            'We do not support Objective2,WebPublish,ItemType and Outcome for Objectives yet but we still need to add these to keep the items in synch.
            mvResultColumns = "ContactCPDCycleNumber,ContactCPDPeriodNumber,ContactCPDPointNumber,CategoryType,CategoryTypeDesc,Category,CategoryDesc,ObjectiveDate,Objective,SupervisorAccepted,AmendedOn,AmendedBy,Objective2,WebPublish,ItemType,ItemTypeDesc,Outcome"
            mvSelectColumns = "CategoryTypeDesc,CategoryDesc,Objective,SupervisorAccepted,AmendedOn,AmendedBy"
            mvHeadings = DataSelectionText.String40007.Replace("Points", "Objective").Replace("Evidence Seen", "Supervisor Accepted")
            mvRequiredItems = "CategoryTypeDesc,CategoryDesc,ObjectiveDate,Objective,SupervisorAccepted"
          Else
            mvResultColumns = "ContactCPDCycleNumber,ContactCPDPeriodNumber,ContactCPDPointNumber,CategoryType,CategoryTypeDesc,Category,CategoryDesc,PointsDate,Points,EvidenceSeen,AmendedOn,AmendedBy,Points2,WebPublish,ItemType,ItemTypeDesc,Outcome"
            mvSelectColumns = "CategoryTypeDesc,CategoryDesc,Points,EvidenceSeen,AmendedOn,AmendedBy"
            mvHeadings = DataSelectionText.String40007        'Category Type,Description,Points,Evidence Seen,Amended On,Amended By
            mvRequiredItems = "CategoryTypeDesc,CategoryDesc,PointsDate,Points,EvidenceSeen"
          End If

          mvWidths = "2000,2000,1200,1200,1200,1200"
          If Len(pEnv.GetConfig("cpd_points_item_name")) <> 0 Then mvHeadings = Replace(mvHeadings, "Points", pEnv.GetConfig("cpd_points_item_name"))
          mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18089     'CPD Details
          mvDisplayTitle = DataSelectionText.String17981    'Details

          mvCode = "CPDD"

        Case DataSelectionTypes.dstCPDPointsEdit
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,ContactCPDPeriodNumber,PeriodStartDate,PeriodEndDate,ContactCPDPointNumber,CategoryType,CategoryTypeDesc,Category,CategoryDesc,Points,EvidenceSeen,AmendedOn,AmendedBy,PointsDate,PeriodDuration,Notes,Points2,WebPublish,ItemType,ItemTypeDesc,Outcome"
          mvSelectColumns = "PeriodDuration,CategoryTypeDesc,CategoryDesc,Points,EvidenceSeen,AmendedOn,AmendedBy"
          mvHeadings = DataSelectionText.String18088        'Period Duration,Category Type,Description,Points,Evidence Seen,Amended On,Amended By
          mvWidths = "2500,2000,2000,1200,1200,1200,1200"
          If Len(pEnv.GetConfig("cpd_points_item_name")) <> 0 Then mvHeadings = Replace(mvHeadings, "Points", pEnv.GetConfig("cpd_points_item_name"))
          mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.CPDPointsMaintenanceForm 'CPD Points Maintenance Form
          mvDisplayTitle = DataSelectionText.CPDPointsEdit           'CPD Points Edit
          mvRequiredItems = "PeriodDuration,CategoryTypeDesc,CategoryDesc,PointsDate,Points,EvidenceSeen,CategoryType,CPDCycleType,Category,Notes,Points2,WebPublish,ItemType,Outcome"
          mvCode = "CPPE"

        Case DataSelectionTypes.dstContactCPDPointsWithoutCycle
          mvResultColumns = "ContactCPDPointNumber,ContactNumber,CategoryType,CategoryTypeDesc,Category,CategoryDesc,Points,EvidenceSeen,AmendedOn,AmendedBy,PointsDate,Notes,Points2,WebPublish,ItemType,ItemTypeDesc,Outcome"
          mvSelectColumns = "PointsDate,CategoryTypeDesc,CategoryDesc,Points,Points2,ItemTypeDesc,Outcome"
          mvHeadings = "Points Date,Category Type,Description,Points,Points 2,Item Desc,Outcome"
          mvWidths = "2500,2000,2000,1200,1200,1200,1200"

          Dim vCpdPointsItemName As String = pEnv.GetConfig("cpd_points_item_name")
          If Not String.IsNullOrEmpty(vCpdPointsItemName) Then
            mvHeadings = Replace(mvHeadings, "Points", pEnv.GetConfig("cpd_points_item_name"))

            mvDescription = String.Format(DataSelectionText.CPDConfigOverrideMaintenanceForm, vCpdPointsItemName) 'CPD {0} Maintenance Form
            mvDisplayTitle = String.Format(DataSelectionText.CPDConfigOverrideEdit, vCpdPointsItemName)           'CPD {0} Points Edit
            mvMaintenanceDesc = String.Format(DataSelectionText.CPDConfigOverrideMaintenance, vCpdPointsItemName) 'CPD {0} Points
          Else
            mvDescription = DataSelectionText.CPDPointsMaintenanceForm 'CPD Points Maintenance Form
            mvDisplayTitle = DataSelectionText.CPDPointsEdit           'CPD Points Edit
            mvMaintenanceDesc = DataSelectionText.CPDPointsMaintenance 'CPD Points
          End If

          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvRequiredItems = "CategoryTypeDesc,CategoryDesc,PointsDate,Points,EvidenceSeen,CategoryType,Category,Notes,Points2,WebPublish,ItemType,Outcome"
          mvCode = "CPPC"

        Case DataSelectionTypes.dstCPDSummary
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,CPDCycleTypeDesc,StartMonth,EndMonth,StartDate,EndDate,ContactCPDPeriodNumber,PeriodStartDate,PeriodEndDate,CategoryType,CategoryTypeDesc,Points,CycleDuration,AmendedOn,AmendedBy,PeriodDuration,CopyContactCPDCycleNumber,CopyCPDCategoryType,Points2,TotalPoints,CPDCycleStatus,CPDCycleStatusDesc,RgbCPDCycleStatus,CPDType,CPDTypeDesc,ContactCPDPointNumber"
          mvSelectColumns = "ContactCPDCycleNumber,CPDCycleTypeDesc,CycleDuration,PeriodDuration,CategoryTypeDesc,Points,CPDTypeDesc"
          mvHeadings = DataSelectionText.String40006        'Cycle Number,Cycle Type,Cycle Duration,Period Duration,Category Type,Points,CPD Type
          If Len(pEnv.GetConfig("cpd_points_item_name")) <> 0 Then mvHeadings = Replace(mvHeadings, "Points", pEnv.GetConfig("cpd_points_item_name"))
          mvWidths = "900,900,1600,1600,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18091     'CPD Summary
          mvDisplayTitle = DataSelectionText.String18092    'Summary
          mvMaintenanceDesc = "CPD Cycle"
          mvCode = "CPDS"
          mvRequiredItems = "PeriodDuration,CategoryType,Points,CPDCycleType,CPDCycleTypeDesc,StartDate,EndDate,ContactCPDPeriodNumber,CopyContactCPDCycleNumber,CopyCPDCategoryType,CycleDuration,ContactCPDCycleNumber,CategoryTypeDesc,CPDType,RgbCPDCycleStatus,CPDCycleStatus,Points2,TotalPoints"

        Case DataSelectionTypes.dstContactCPDCyclesEdit
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,CPDCycleTypeDesc,StartMonth,EndMonth,StartDate,EndDate,AmendedOn,AmendedBy,CPDCycleStatus,CPDCycleStatusDesc,RgbCPDCycleStatus,CPDType,ContactNumber,CycleStart,CycleEnd"
          mvSelectColumns = "ContactCPDCycleNumber,CPDCycleTypeDesc,CycleStart,CycleEnd"
          mvHeadings = "Cycle Number,Cycle Type,Cycle Start,Cycle End"

          'mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,CPDCycleTypeDesc,CycleDuration,StartMonth,EndMonth,StartDate,EndDate,AmendedOn,AmendedBy,DefaultDuration"
          'mvSelectColumns = "ContactCPDCycleNumber,CPDCycleTypeDesc,CycleDuration"
          'mvHeadings = DataSelectionText.String40008        'Cycle Number,Cycle Type,Cycle Duration
          mvWidths = "1600,2500,2000,2000"
          mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18278     'CPD Cycles Maintenance Form
          mvDisplayTitle = DataSelectionText.String18279    'CPD Cycles Edit
          mvCode = "CPCE"
          'mvRequiredItems = "ContactCPDCycleNumber,StartDate,EndDate,CPDCycleType,DefaultDuration"
          mvRequiredItems = "ContactCPDCycleNumber,CycleStart,CycleEnd,CPDCycleType,CPDType,CPDCycleStatus,ContactNumber,StartMonth,EndMonth"

        Case DataSelectionTypes.dstContactSalesLedgerItems
          mvResultColumns = "Date,TransactionType,InvoiceNumber,Reference,Status,DueDate,DepositAmount,BatchNumber,TransactionNumber,TransactionSign,Reprinted,Debit,Credit,Outstanding,StoredInvoiceNumber,PayerContactNumber,AddressNumber,AddressLine"
          mvSelectColumns = "Date,TransactionType,InvoiceNumber,Reference,Status,Debit,Credit,DueDate,DepositAmount,Outstanding,BatchNumber,TransactionNumber,TransactionSign,Reprinted,StoredInvoiceNumber"
          mvHeadings = DataSelectionText.String29954 'Date,Type,Inv/Crn Number,Reference,Status,Debit,Credit,Due Date,Deposit Amount,Outstanding,Batch,Transaction,Sign,Reprinted,Stored Invoice Number
          mvWidths = "1200,1000,1300,1200,500,1200,1200,1200,1200,1200,1200,1200,1,1000,1"
          mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18093     'Sales Ledger Items
          mvDisplayTitle = DataSelectionText.String18092     'Items
          mvCode = "CBAS"
          mvRequiredItems = "TransactionType,Status,DueDate,TransactionSign,BatchNumber,TransactionNumber,Outstanding"

        Case DataSelectionTypes.dstSelectionSetContacts
          mvResultColumns = "ContactNumber,ContactName,ContactType,GroupCode,PhoneNumber,AddressNumber,AddressType,HouseName,Address,Town,County,Postcode,Country,AddressLine,Status,Title,Forenames,PreferredForename,Surname,ContactPositionNumber,Position,OrganisationNumber,OrganisationName,DateOfBirth,OwnershipGroup,PrincipalDepartment,OwnershipAccessLevelDesc,OwnershipAccessLevel"
          mvSelectColumns = "ContactNumber,ContactName,PhoneNumber,Town,Postcode,Status,DateOfBirth,OwnershipAccessLevelDesc"
          mvHeadings = DataSelectionText.String18094     'Contact Number,Name,Phone,Town,Postcode,Status,Date of birth,Access Level
          mvWidths = "900,2000,900,900,1200,1200,1200,1200"
          mvRequiredItems = "ContactType,GroupCode,OwnershipAccessLevel"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstSelectionSetCommsNumbers
          mvResultColumns = "ContactNumber,AddressNumber,DeviceCode,DeviceDesc,DiallingCode,STDCode,Extension,ExDirectory,Notes,AmendedBy,AmendedOn,CommunicationNumber,Number,PhoneNumber"
          mvSelectColumns = "ContactNumber,AddressNumber,DeviceCode,PhoneNumber"
          mvHeadings = DataSelectionText.String22867     'Contact Number,Address Number,Device,Number
          mvWidths = "900,900,900,2000"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCriteriaSetDetails
          mvResultColumns = "CriteriaSet,SequenceNumber,SearchArea,IE,CO,MainValue,SubsidiaryValue,Period,Counted,AndOr,LeftParenthesis,RightParenthesis"
          mvSelectColumns = "AndOr,LeftParenthesis,IE,CO,SearchArea,MainValue,SubsidiaryValue,Period,RightParenthesis,Counted" '     CriteriaSet,SequenceNumber,SearchArea,IE,CO,MainValue,SubsidiaryValue,Period,Counted,AndOr,LeftParenthesis,RightParenthesis"
          mvHeadings = DataSelectionText.String40011     'And/Or,(,I/E,C/O,Area,Value,Sub Value,Period,),Counted
          mvWidths = "1200,1200,1200,2200,2200,2200,2200,2200,1200,2200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstGeneralMailingSelectionSets
          mvResultColumns = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source"
          mvSelectColumns = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source"
          mvHeadings = DataSelectionText.String40013 ' Set,Description,Owner,Department,Records
          mvWidths = "200,800,4100,400,1500"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String40012   ' General Mailing Selection Set
          mvCode = "GMS"
          vPrimaryList = True
          mvRequiredItems = "SelectionSet,SelectionSetDesc,UserName,Department,NumberInSet,Source"

        Case DataSelectionTypes.dstCriteriaSets
          mvResultColumns = "CriteriaSetNumber,CriteriaSetDesc,Owner,Department,Report,StandardDocument,ContactCriteriaCount,OrganisationCriteriaCount,Mailing"
          If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports) Then mvResultColumns = Replace$(mvResultColumns, ",Report,StandardDocument,ContactCriteriaCount,OrganisationCriteriaCount,Mailing", "")
          mvSelectColumns = "CriteriaSetNumber,CriteriaSetDesc,Owner,Department,Report,StandardDocument"
          mvHeadings = DataSelectionText.String18095     'Set,Description,Owner,Department,Report,StandardDocument
          mvWidths = "1,6500,4200,4270,5200,5200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvCode = "CSDT"

        Case DataSelectionTypes.dstSelectionPages
          mvResultColumns = "SelectionHeading,General,Addresses,Numbers,Categories,Positions,Relationships,RelationshipsFrom,References,Journal,StatusHistory,Notes,StickyNotes,Actions,Calendar,DBANotes,DepartmentHistory,Dashboard,CPD,CPDPoints,RegisteredUser,NetworkLink,Meetings,Alerts,Amendments,CreditCards," &
                            "SelectionHeading1,Documents,Mailings,Suppressions,History,Surveys," &
                            "SelectionHeading2,Transactions,FinancialHistory,BackOrders,DespatchNotes,CancelledProvisional,ServiceBookings,BankAccounts,PurchaseOrders,PurchaseInvoices,Account,SalesTransactions,DeliveryTransactions," &
                            "SelectionHeading3,Memberships,Covenants,Subscriptions,StandingOrders,DirectDebits,CreditCardAuthorities,Loans,PaymentPlans,PreTaxPledges,PostTaxPledges," &
                            "SelectionHeading10,Requests,EventSponsorship," &
                            "SelectionHeading12,ExamSummary,ExamDetails,ExamExemptions,ExamCertificates," &
                            "SelectionHeading4,InMemoriamReceived,InMemoriamDonated,ContributedTo,PaidInBy,SentOnBehalfOf,HandledBy," &
                            "SelectionHeading5,Declarations,UnclaimedSponsorship,ClaimedSponsorship,AppropriateCertificates," &
                            "SelectionHeading6,Booked,Delegate,Sessions,Accommodation,Rooms,Organiser,Personnel," &
                            "SelectionHeading7,MannedCollections,H2HCollections,UnmannedCollections," &
                            "SelectionHeading11,Legacy,LegacyBequests,LegacyAssets,LegacyLinks,LegacyTaxCertificates,LegacyExpenses,LegacyActions," &
                            "SelectionHeading8,Performances,Scores"
          mvSelectColumns = mvResultColumns.Replace(",CPDPoints,", ",")
          mvResultColumns = mvResultColumns & ",SelectionHeading9"
          Dim vHeadings As New StringBuilder
          vHeadings.Append(DataSelectionText.String22836) 'View Details,General,Addresses,Numbers,Categories,Positions,Relationships To,Relationships From,References,Journal,Status History,Notes,Sticky Notes,Actions,Calendar,DBANotes,Department History,Dashboard,CPD,Registered User,Network Link,Meetings,Alerts,Amendments,Credit Cards
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22837)  'View Communications,Documents,Mailings,Suppressions,History,Surveys
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22838)  'View Financial,Transactions,History,Back Orders,Despatch Notes,Cancelled Provisional,Service Bookings,Bank Accounts,Purchase Orders,Purchase Invoices,Account,Sales Transactions,Delivery Transactions
          vHeadings.Append(",")

          If Not mvEnv.GetConfigOption("option_gaye", True) Then
            vHeadings.Append(DataSelectionText.String22839.Replace(",Pre Tax Pledges", String.Empty))  'View Commitments,Memberships,Covenants,Subscriptions,Standing Orders,Direct Debits,Credit Card Authorities,Loans,All Payment Plans,Pre Tax Pledges,Post Tax Pledges
            mvResultColumns = mvResultColumns.Replace(",PreTaxPledges", String.Empty)
            mvSelectColumns = mvSelectColumns.Replace(",PreTaxPledges", String.Empty)
          Else
            vHeadings.Append(DataSelectionText.String22839)  'View Commitments,Memberships,Covenants,Subscriptions,Standing Orders,Direct Debits,Credit Card Authorities,Loans,All Payment Plans,Pre Tax Pledges,Post Tax Pledges
          End If

          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String40003)  'View Fundraising,Requests,Event Sponsorship
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String40027)  'View Exams,Summary,Bookings,Exemptions,Certificates
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22840)  'View Financial Links,In Memoriam Received,In Memoriam Donated,Contributed To,Paid In By,Sent On Behalf Of,Handled By
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22841)  'View Gift Aid,Declarations,Unclaimed Sponsorship,Claimed Sponsorship,Appropriate Certificates
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22842)  'View Events,Booked,Delegate,Sessions,Accommodation,Rooms,Organiser,Personnel
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22843)  'View Collections,Manned,House to House,Un-Manned
          vHeadings.Append(",")
          vHeadings.Append("View Legacies,Legacy,Bequests,Assets,Links,Tax Certificates,Expenses,Actions")
          vHeadings.Append(",")
          vHeadings.Append(DataSelectionText.String22844)  'View Profiles,Performances,Scores,View Other
          mvHeadings = vHeadings.ToString
          mvWidths = "300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300," &
                     "300,300,300,300,300,300," &
                     "300,300,300,300,300,300,300,300,300,300,300,300," &
                     "300,300,300,300,300,300,300,300,300,300,300," &
                     "300,300,300," &
                     "300,300,300,300,300," &
                     "300,300,300,300,300,300,300," &
                     "300,300,300,300,300," &
                     "300,300,300,300,300,300,300,300," &
                     "300,300,300,300," &
                     "300,300,300,300,300,300,300,300," &
                     "300,300,300,300"
          If Not mvEnv.GetConfigOption("option_gaye", True) Then
            mvWidths = mvWidths.Remove(mvWidths.Length - ",300".Length())
          End If
          CheckCustomForms()
          CheckActivityDataSheets()
          If (Not pParams Is Nothing) AndAlso pParams.ContainsKey("IncludeGroupsFromContactCard") AndAlso pParams("IncludeGroupsFromContactCard").Value = "Y" Then
            CheckOrganisationGroups()
          End If
          CheckRelationshipDataSheets()
          CheckDashboardItem()
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18096     'Selection Pages
          mvCode = "SCPG"

        Case DataSelectionTypes.dstEventSelectionPages
          mvResultColumns = "SelectionHeading,General,Sessions,Venues,BookingOptions,Organiser,Personnel,Tasks,Resources,Accommodation,Submissions,Contacts,Owners,Documents,Dashboard,Actions,CPD," &
                            "SelectionHeading1,Topics,Sources,Mailings,Activities," &
                            "SelectionHeading2,EventBookings,AccommodationBookings,Delegates," &
                            "SelectionHeading3,Costs,FinancialHistory,FinancialLinks,PayingInSlips," &
                            "SelectionHeading4,SessionTests,TestResults"
          mvSelectColumns = mvResultColumns.Replace(",CPD,", ",")
          mvHeadings = DataSelectionText.String22845                  'Details,General,Sessions,Venues,Booking Options,Organiser,Personnel,Tasks,Resources,Accommodation,Submissions,Contacts,Owners,Documents,Dashboard,Actions
          mvHeadings = mvHeadings & "," & DataSelectionText.String22846  'Analysis,Topics,Sources,Mailings,Activities
          mvHeadings = mvHeadings & "," & DataSelectionText.String22847  'Bookings,Event Bookings,Accommodation Bookings,Delegates
          mvHeadings = mvHeadings & "," & DataSelectionText.String22848  'Financial,Costs,History,Links,Paying-In-Slips
          mvHeadings = mvHeadings & "," & DataSelectionText.String22849  'Tests,Session Tests,Test Results
          mvWidths = "300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300"
          CheckDashboardItem()
          CheckTopicDataSheets()
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18171     'Event Selection Pages
          mvCode = "SCEP"

        Case DataSelectionTypes.dstContactInformation
          Dim vContact As New Contact(mvEnv)
          mvResultColumns = vContact.DataTableColumns
          mvSelectColumns = "ContactNumber,AddressNumber,ContactName"
          mvRequiredItems = "OwnershipGroup,RgbStatus"
          mvHeadings = DataSelectionText.String18097     'Contact Number,Address Number,Name
          mvWidths = "300,300,300"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18098     'Contact Information
          mvCode = "CONI"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactCommsInformation
          Dim vContact As New Contact(mvEnv)
          mvResultColumns = vContact.CommsDataTableColumns
          mvSelectColumns = mvResultColumns
          mvHeadings = "Contact Number,Address Number,Direct Number,Mobile Number,EMail Address,Switchboard Number,Fax Number,Web Address"
          mvWidths = "300,300,300,300,300,300,300,300"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCustomFormData
          InitCustomFormData()

        Case DataSelectionTypes.dstDashboardData
          InitDashboardData()

        Case DataSelectionTypes.dstActivitiesDataSheet
          mvResultColumns = "ActivityCode,ActivityDesc,ActivityValueCode,ActivityValueDesc,Mandatory,QuantityRequired,MultipleValues,ContactGroup,ActivityDurationMonths,ActivityDurationDays,ActivityValueDurationMonths,ActivityValueDurationDays,IsActivityHistoric,IsActivityValueHistoric,SourceCode"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1800,1,1800,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuCare
          mvCode = "ACDS"

        Case DataSelectionTypes.dstRelationshipsDataSheet
          mvResultColumns = "RelationshipCode,RelationshipDesc,Mandatory,ToContactGroup,MultipleValues,ContactSelectionType,ComplimentaryRelationship,PostPoint"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1800,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuCare
          mvCode = "RSDS"
        Case DataSelectionTypes.dstSuppressionDataSheet
          mvResultColumns = "MailingSuppression,MailingSuppressionDesc"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1800"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstPackProductDataSheet
          mvResultColumns = "LinkProductCode,LinkProductDesc,WarehouseCode,WarehouseDesc,LastStockCount,OriginalCost,DefaultWarehouseDesc"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1800,1200,1800,1200,1200,1"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuCare
          mvCode = "PPDS"

        Case DataSelectionTypes.dstAppealCollections
          mvResultColumns = "CollectionNumber,Collection,CollectionDesc,ProductCode,RateCode,SourceCode,BankAccountCode,CollectionType"
          mvSelectColumns = "CollectionNumber,Collection,CollectionDesc,ProductCode,RateCode,SourceCode,BankAccountCode"
          mvHeadings = "Collection Number,Collection,Description,Product,Rate,Source,Bank Account"
          mvRequiredItems = "ProductCode,RateCode,SourceCode,BankAccountCode,CollectionType"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCampaigns
          mvResultColumns = "Campaign,CampaignDesc,StartDate,EndDate,Manager,BusinessType,Status,StatusDate,StatusReason,Notes,ActualIncome,LastUpdated,MarkHistorical,TotalItemisedCost,Topic"
          mvSelectColumns = "Campaign,CampaignDesc,StartDate,EndDate,Manager,BusinessType,Status,StatusDate,StatusReason,Notes,ActualIncome,LastUpdated"
          mvHeadings = DataSelectionText.String18099     'Campaign,Description,Start Date,End Date,Manager,Business Type,Status,Status Date,Status Reason,Notes,Actual Income,Last Updated
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18100     'Campaigns
          mvDisplayTitle = "Campaigns"
          mvCode = "CPCP"
          mvRequiredItems = "MarkHistorical,TotalItemisedCost,Topic"

        Case DataSelectionTypes.dstCampaignAppeals
          mvResultColumns = "Campaign,Appeal,AppealDesc,ThankYouLetter"
          mvSelectColumns = "Campaign,Appeal,AppealDesc,ThankYouLetter"
          mvHeadings = DataSelectionText.String18140     'Campaign,Appeal,Appeal Description,Thank You Letter
          mvWidths = "1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18141     'Campaign Appeals
          mvDisplayTitle = "Campaign Appeals"
          mvCode = "CPAP"

        Case DataSelectionTypes.dstFindMeeting
          mvResultColumns = "MeetingNumber,MeetingDate,MeetingTime,MeetingDesc,MeetingType,MeetingLocationDesc"
          mvSelectColumns = "MeetingNumber,MeetingDate,MeetingTime,MeetingDesc,MeetingType,MeetingLocationDesc"
          mvHeadings = "Number,Date,Time,Description,Type,Location" 'DataSelectionText.String18140     'Campaign,Appeal,Appeal Description,Thank You Letter
          mvWidths = "1,1150,600,3500,2400,2400"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCampaignCollections
          mvResultColumns = "Campaign,CampaignDesc,StartDate,Appeal,AppealDesc,CollectionNumber,Collection,CollectionDesc"
          mvSelectColumns = "Campaign,CampaignDesc,StartDate,Appeal,AppealDesc,CollectionNumber,Collection,CollectionDesc"
          mvHeadings = DataSelectionText.String18142     'Campaign,Description,Start Date,Appeal,Description,Segment Sequence,Segment,Segment Description
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCampaignSegments
          mvResultColumns = "Campaign,CampaignDesc,StartDate,Appeal,AppealDesc,SegmentSequence,Segment,SegmentDesc"
          mvSelectColumns = "Campaign,CampaignDesc,StartDate,Appeal,AppealDesc,SegmentSequence,Segment,SegmentDesc"
          mvHeadings = DataSelectionText.String18142     'Campaign,Description,Start Date,Appeal,Description,Segment Sequence,Segment,Segment Description
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          'mvDescription = DataSelectionText.String18143     'Campaign Segments
          'mvDisplayTitle = "Campaign Segments"
          'mvCode = "CPSG"

        Case DataSelectionTypes.dstCampaignInfo
          mvResultColumns = "Campaign,CampaignDesc,Appeal,AppealDesc,SegmentSequence,Segment,SegmentDesc,SegmentDate,CampaignEndDate,CampaignStartDate,AppealDate,AppealEndDate,H2hStartDate,UnmannedStartDate,H2hEndDate,UnmannedEndDate,CollectionDate,Collection,CollectionDesc,CollectionNumber,AppealType"
          mvSelectColumns = "Campaign,CampaignDesc,StartDate,Appeal,AppealDesc,SegmentSequence,Segment,SegmentDesc,Collection,CollectionDesc,CollectionNumber"
          mvHeadings = DataSelectionText.String18144     'Campaign,Description,Start Date,Appeal,Description,Segment Sequence,Segment,Segment Description,Collection,Collection Desc,Collection Number
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstDelegateActivities
          mvResultColumns = "DelegateActivityNumber,EventDelegateNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity,Status,SourceDesc,ValidFrom,ValidTo,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18147     'Category,Value,Quantity,Status,Source,Valid from,Valid to,Notes,Amended by,on
          mvWidths = "1800,1800,1200,1200,1800,1200,1200,3600,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDisplayTitle = DataSelectionText.String17795     'Activities
          mvDescription = DataSelectionText.String18148     'Delegate Activities
          mvCode = "EDAC"

        Case DataSelectionTypes.dstDelegateLinks
          mvResultColumns = "DelegateLinkNumber,EventDelegateNumber,RelationshipCode,Type1,Type2,RelationshipDesc," & ContactNameResults() & ",ContactNumber,Phone,ValidFrom,ValidTo,Historical,Notes,AmendedBy,AmendedOn,OwnershipGroup,ContactGroup"
          mvSelectColumns = "RelationshipDesc,ContactName,Phone,Historical,ValidFrom,ValidTo,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18149     'Relationship,With,Phone,Historical,Valid from,Valid to,Notes,Amended by,on
          mvWidths = "1500,1500,1200,1000,1200,1200,3600,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDisplayTitle = DataSelectionText.String18150     'Relationships
          mvDescription = DataSelectionText.String18151     'Delegate Relationships
          mvCode = "EDRL"

        Case DataSelectionTypes.dstCollectionRegions
          mvResultColumns = "CollectionRegionNumber,CollectionNumber,GeographicalRegion,GeographicalRegionDesc"
          mvSelectColumns = "GeographicalRegion,GeographicalRegionDesc"
          mvHeadings = DataSelectionText.String22868     'Collection Region,Collection Region Desc
          mvWidths = "1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18640     'Manned and House-To-House Collection Regions
          mvDisplayTitle = DataSelectionText.String18641     'Collection Regions
          mvCode = "ACRE"
          mvRequiredItems = "GeographicalRegion"

        Case DataSelectionTypes.dstCollectionPoints
          mvResultColumns = "CollectionPointNumber,GeographicalRegion,GeographicalRegionDesc,CollectionRegionNumber,CollectionPoint,CollectionPointType,OrganisationNumber,Organisation,AddressLine,NoOfCollectors,Notes"
          mvSelectColumns = "CollectionPointNumber,CollectionPoint,CollectionPointType,Organisation,AddressLine,NoOfCollectors,Notes"
          mvHeadings = DataSelectionText.String22869     'Collection Point Number,Collection Point,Collection Point Type,Organisation,Address,No Of Collectors,Notes
          mvWidths = "1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18642     'Manned Collection Collection Points
          mvDisplayTitle = DataSelectionText.String18643     'Collection Points
          mvCode = "MCCP"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstCollectionResources
          mvResultColumns = "CollectionResourceNumber,CollectionNumber,AppealResourceNumber,Product,ProductDesc,Rate,RateDesc,Quantity,DespatchOn,DespatchMethod,DespatchMethodDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "ProductDesc,RateDesc,Quantity,DespatchOn,DespatchMethodDesc"
          mvHeadings = DataSelectionText.String22870     'Product,Rate,Quantity,Despatch On,Despatch Method
          mvWidths = "1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18644     'Manned and Un-Manned Collection Resources
          mvDisplayTitle = DataSelectionText.String18645     'Collection Resources
          mvCode = "ACRC"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstMannedCollectors
          mvResultColumns = ContactNameResults() & ",CollectorNumber,CollectionNumber,ContactNumber,TotalTime,Attended,ReadyForConfirmation,ReadyForAcknowledgement,ConfirmationProducedOn,AcknowledgementProducedOn,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactName,TotalTime,Attended"
          mvHeadings = DataSelectionText.String22871     'Collector,Total Time,Attended
          mvWidths = "2000,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18646     'Manned Collection Collectors
          mvDisplayTitle = DataSelectionText.String18647     'Manned Collectors
          mvCode = "MCCL"
          mvRequiredItems = "TotalTime"

        Case DataSelectionTypes.dstCollectorShifts
          mvResultColumns = ContactNameResults() & ",CollectorShiftNumber,CollectorNumber,CollectionRegionNumber,GeographicalRegion,GeographicalRegionDesc,CollectionPointNumber,CollectionPoint,StartTime,EndTime,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactName,CollectionPoint,StartTime,EndTime,Notes"
          mvHeadings = DataSelectionText.String22872     'Collector,CollectionPoint,StartTime,EndTime,Notes
          mvWidths = "1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18648     'Manned Collection Collector Shifts
          mvDisplayTitle = DataSelectionText.String18649     'Collector Shifts
          mvCode = "MCCS"
          mvRequiredItems = "StartTime,EndTime"

        Case DataSelectionTypes.dstMannedCollectionBoxes, DataSelectionTypes.dstUnMannedCollectionBoxes, DataSelectionTypes.dstContactCollectionBoxes
          mvResultColumns = ContactNameResults() & ",AddressLine,ContactTelephone,CollectionBoxNumber,CollectionNumber,BoxReference,CollectorNumber,ContactNumber,Amount,CollectionPisNumber,PisNumber,AmendedBy,AmendedOn,SumPayments"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          Select Case mvType
            Case DataSelectionTypes.dstMannedCollectionBoxes
              mvSelectColumns = "BoxReference,ContactName,AddressLine,ContactTelephone,Amount,PisNumber"
              mvHeadings = DataSelectionText.String22873     'BoxReference,ContactName,Address,Communication,Amount,PisNumber
              mvWidths = "1200,1200,1200,1200,1200,1200"
              mvDescription = DataSelectionText.String18650     'Manned Collection Boxes
              mvDisplayTitle = DataSelectionText.String18651     'Collection Boxes
              mvCode = "MCCB"
              mvRequiredItems = "SumPayments"

            Case DataSelectionTypes.dstUnMannedCollectionBoxes
              mvSelectColumns = "BoxReference,Amount"
              mvHeadings = DataSelectionText.String22874     'BoxReference,Amount
              mvWidths = "1200,1200"
              mvDescription = DataSelectionText.String18652     'Un-Manned Collection Boxes
              mvDisplayTitle = DataSelectionText.String18651     'Collection Boxes
              mvCode = "UCCB"
              mvRequiredItems = "BoxReference"

            Case DataSelectionTypes.dstContactCollectionBoxes
              mvSelectColumns = "BoxReference,Amount,PisNumber"
              mvHeadings = DataSelectionText.String22875     'BoxReference,Amount,PisNumber
              mvWidths = "1200,1200,1200"
              mvDescription = DataSelectionText.String18653     'Contact/Organisation Collection Boxes
              mvDisplayTitle = DataSelectionText.String18651     'Collection Boxes
              mvCode = "CCB"
              mvRequiredItems = ""
          End Select

        Case DataSelectionTypes.dstAppealResources
          mvResultColumns = "AppealResourceNumber,Campaign,CampaignDesc,Appeal,AppealDesc,Product,ProductDesc,TotalQuantity,QuantityRemaining"
          mvSelectColumns = "ProductDesc,TotalQuantity,QuantityRemaining"
          mvHeadings = DataSelectionText.String22876     'Product,Total Quantity,Quantity Remaining
          mvWidths = "1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18654     'Appeal Resources
          mvDisplayTitle = DataSelectionText.String18654     'Appeal Resources
          mvCode = "APRC"
          mvRequiredItems = "Product"

        Case DataSelectionTypes.dstCollectionBoxesForPayment
          mvResultColumns = "CollectionBoxNumber,ContactNumber,BoxReference,Amount,Pay"
          mvSelectColumns = "CollectionBoxNumber,ContactNumber,BoxReference,Amount,Pay"
          mvHeadings = DataSelectionText.String19044        'Collection Box Number,Contact Number,Box Reference,Amount,Pay
          mvRequiredItems = ""
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCollectionPayments, DataSelectionTypes.dstContactCollectionPayments
          mvResultColumns = "BoxReference,CollectionBoxNumber,TransactionDate,Amount,CollectionPaymentNumber,CollectionNumber,CollectionPISNumber,PISNumber,BatchNumber,TransactionNumber,LineNumber,SentOnBehalfOfContactNumber,ContactNumber," & ContactNameResults() & ",SentOnBehalfOfContactName"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

          If mvType = DataSelectionTypes.dstCollectionPayments Then
            mvSelectColumns = "TransactionDate,PISNumber,ContactName,SentOnBehalfOfContactName,Amount,BatchNumber,TransactionNumber,LineNumber,BoxReference"
            mvHeadings = DataSelectionText.String22877     'Transaction Date,Paying In Slip,ContactName,SentOnBehalfOfContactName,Amount,Batch Number,Transaction Number,Line Number,Box Reference
            mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200"
            mvDescription = DataSelectionText.String18655     'Payments for the Collection
            mvDisplayTitle = DataSelectionText.String18656     'Collection Payments
            mvCode = "CACP"
            mvRequiredItems = ""
          Else
            mvSelectColumns = "TransactionDate,PISNumber,Amount,BatchNumber,TransactionNumber,LineNumber,ContactName,BoxReference"
            mvHeadings = DataSelectionText.String22878     'Transaction Date,Paying In Slip,Amount,Batch Number,Transaction Number,Line Number,Payer,Box Reference
            mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
            mvDescription = DataSelectionText.String18657     'Collection Payments for the Contact
            mvDisplayTitle = DataSelectionText.String18656     'Collection Payments
            mvCode = "COCP"
            mvRequiredItems = ""
          End If

        Case DataSelectionTypes.dstCollectionPIS, DataSelectionTypes.dsth2hCollectionPIS
          mvResultColumns = "ContactName,CollectorSurname,CollectorForenames,CollectorTitle,CollectorInitials,CollectionNumber,CollectionPISNumber,PisNumber,CollectorNumber,ContactNumber,IssueDate,Amount,BankedBy,BankedOn,ReconciledOn,BankedByContactName"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

          If mvType = DataSelectionTypes.dstCollectionPIS Then
            mvSelectColumns = "PisNumber,IssueDate,Amount,BankedByContactName,BankedOn,ReconciledOn"
            mvHeadings = DataSelectionText.String22879     'Pis Number,Issue Date,Amount,Banked By,Banked On,Reconciled On
            mvWidths = "1200,1200,1200,2000,1200,1200"
            mvDescription = DataSelectionText.String18658     'Manned and Un-Manned Collection Paying In Slips
            mvDisplayTitle = DataSelectionText.String18659     'Collection Paying In Slips
            mvCode = "MCPS"
            mvRequiredItems = ""
          Else
            mvSelectColumns = "PisNumber,IssueDate,ContactName"
            mvHeadings = DataSelectionText.String22880     'Pis Number,Issue Date,Collector
            mvWidths = "1200,1200,2000"
            mvDescription = DataSelectionText.String18660     'House-To-House Collection Paying In Slips
            mvDisplayTitle = DataSelectionText.String18659     'Collection Paying In Slips
            mvCode = "HCPS"
            mvRequiredItems = ""
          End If

        Case DataSelectionTypes.dstH2HCollectors
          mvResultColumns = "ContactName,Surname,Forenames,Title,Initials,AddressLine,Communication,ReadyForConfirmation,CollectionNumber,CollectorNumber,ContactNumber,Route,RouteType,RouteTypeDesc,NoOfPremises,OperatorContactNumber,CollectorStatus,CollectorStatusDesc,Notes,ConfirmationProducedOn,ReminderProducedOn,AmendedBy,AmendedOn,OperatorContactName"
          mvSelectColumns = "ContactName,AddressLine,Communication,Route,RouteTypeDesc,NoOfPremises,OperatorContactName,CollectorStatusDesc,ReadyForConfirmation,Notes,ConfirmationProducedOn,ReminderProducedOn"
          mvHeadings = DataSelectionText.String22881     'Collector,Address,Communication,Route,Route Type,No Of Premises,Operator,Collector Status,Ready For Confirmation,Notes, Confirmation Produced On,Reminder Produced On
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18661     'House-To-House Collection Collectors
          mvDisplayTitle = DataSelectionText.String18662     'Collection Collectors
          mvCode = "HCCL"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstSegmentProducts
          mvResultColumns = "Campaign,Appeal,Segment,AmountNumber,Product,ProductDesc,Rate,RateDesc"
          mvSelectColumns = "AmountNumber,ProductDesc,RateDesc"
          mvHeadings = DataSelectionText.String22882     'Amount Number,Product,Rate
          mvWidths = "1200,2000,2000"
          mvCode = "CSPA"
          mvDescription = "Segment Products"

        Case DataSelectionTypes.dstAppealBudgets
          mvResultColumns = "AppealBudgetNumber,Campaign,Appeal,BudgetPeriod,PeriodStartDate,PeriodEndDate,PeriodPercentage"
          mvSelectColumns = "BudgetPeriod,PeriodStartDate,PeriodEndDate,PeriodPercentage"
          mvHeadings = DataSelectionText.String18152     'Period,Start Date,End Date,Percentage
          mvWidths = "1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18663     'Appeal Budgets
          mvDisplayTitle = DataSelectionText.String18663     'Appeal Budgets
          mvCode = "ABGT"
          mvRequiredItems = "Campaign,Appeal,BudgetPeriod"

        Case DataSelectionTypes.dstAppealBudgetDetails
          mvResultColumns = "AppealBudgetDetailsNumber,AppealBudgetNumber,Segment,ReasonForDespatch,ForecastUnits,BudgetedCosts,BudgetedIncome,BudgetPeriod,Campaign,Appeal,ReasonForDespatchDesc"
          mvSelectColumns = "AppealBudgetDetailsNumber,AppealBudgetNumber,Segment,ReasonForDespatch,ForecastUnits,BudgetedCosts,BudgetedIncome,ReasonForDespatchDesc"
          mvHeadings = DataSelectionText.String22883     'Budget Details Number,Budget Number,Segment,Reason For Despatch,Forecast Units,Budgeted Costs,Budgeted Income,Reason For Despatch Desc
          mvWidths = Replace(CStr(Space(UBound(Split(mvSelectColumns, ",")) + 1)), " ", "300,")
          mvWidths = Left$(mvWidths, Len(mvWidths) - 1)
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18664     'Appeal Budget Details
          mvDisplayTitle = DataSelectionText.String18664     'Appeal Budget Details
          mvCode = "ABGD"
          mvRequiredItems = "Campaign,Appeal"

        Case DataSelectionTypes.dstTickBoxes
          mvResultColumns = "Campaign,Appeal,Segment,TickBoxNumber,Activity,ActivityValue,Suppression,ActivityDesc,ActivityValueDesc,SuppressionDesc"
          mvSelectColumns = "TickBoxNumber,ActivityDesc,ActivityValueDesc,SuppressionDesc"
          mvHeadings = DataSelectionText.String18153     'Tick Box Number,Activity,Activity Value,Suppression
          mvWidths = "1200,2000,2000,2000"
          mvDescription = "Tick Boxes"
          mvCode = "CSTB"

        Case DataSelectionTypes.dstSegmentCostCentres
          mvResultColumns = "Campaign,Appeal,Segment,CostCentre,CostCentreDesc,CostCentrePercentage"
          mvSelectColumns = "CostCentreDesc,CostCentrePercentage"
          mvHeadings = DataSelectionText.String18154     'Cost Centre,Percentage
          mvWidths = "2000,1200"
          mvRequiredItems = "CostCentre"
          mvDescription = "Segment Cost Centres"
          mvCode = "CSCC"

        Case DataSelectionTypes.dstVariableParameters
          mvResultColumns = "Campaign,Appeal,VariableName,VariableValue"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1,2000,3000"

        Case DataSelectionTypes.dstSuppliers
          mvResultColumns = "Campaign,Appeal,Segment,ContactNumber,OrganisationName,SupplierRole,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "OrganisationName,SupplierRole,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String18155     'Supplier,Role,Notes,Amended By,Amended On
          mvWidths = "1200,2000,2000,2000,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18665     'Suppliers
          mvDisplayTitle = DataSelectionText.String18665     'Suppliers
          mvCode = "CASS"
          mvRequiredItems = "Campaign,Appeal,Segment"

        Case DataSelectionTypes.dstAppealTypes
          mvResultColumns = "AppealType,AppealTypeDesc,Access"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1200,1200,1200"

        Case DataSelectionTypes.dstContactDepartmentHistory
          mvResultColumns = "ContactNumber,Department,DepartmentDesc,ValidTo,AmendedBy"
          mvSelectColumns = "DepartmentDesc,ValidTo,AmendedBy"
          mvHeadings = DataSelectionText.String18172     'Department,Valid To,Amended By
          mvWidths = "2000,1200,1200"

        Case DataSelectionTypes.dstContactMannedCollections
          mvResultColumns = "ContactNumber,CollectionNumber,CampaignDesc,AppealDesc,CollectionDesc,CollectionDate,StartTime,EndTime"
          mvSelectColumns = "CampaignDesc,AppealDesc,CollectionDesc,CollectionDate,StartTime,EndTime"
          mvHeadings = DataSelectionText.String22884     'Campaign,Appeal,Collection,Collection Date,Start Time,End Time
          mvWidths = "2000,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18666    'Contact Manned Collections
          mvDisplayTitle = DataSelectionText.String18667    'Manned Collections
          mvCode = "CMC"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstContactUnMannedCollections
          mvResultColumns = "OrganisationNumber,CollectionNumber,CampaignDesc,AppealDesc,CollectionDesc,StartDate,EndDate,ContactNumber," & ContactNameResults()
          mvSelectColumns = "CampaignDesc,AppealDesc,CollectionDesc,StartDate,EndDate,ContactName"
          mvHeadings = DataSelectionText.String22885     'Campaign,Appeal,Collection,Start Date,End Date,Contact Name
          mvWidths = "2000,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18668     'Contact Un-Manned Collections
          mvDisplayTitle = DataSelectionText.String18669     'Un-Manned Collections
          mvCode = "CUC"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstContactH2HCollections
          mvResultColumns = "ContactNumber,CollectionNumber,CampaignDesc,AppealDesc,CollectionDesc,StartDate,EndDate,Route,RouteType,RouteTypeDesc,CollectorStatus,CollectorStatusDesc,OperatorContactNumber,OperatorContactName,OperatorSurname,OperatorForenames,OperatorTitle,OperatorInitials"
          mvSelectColumns = "CampaignDesc,AppealDesc,CollectionDesc,StartDate,EndDate,Route,RouteTypeDesc,CollectorStatusDesc,OperatorContactName"
          mvHeadings = DataSelectionText.String22886     'Campaign,Appeal,Collection,Start Date,End Date,Route,Route Type,Collector Status,Operator
          mvWidths = "2000,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18670     'Contact House-To-House Collections
          mvDisplayTitle = DataSelectionText.String18671     'House-To-House Collections
          mvCode = "CHC"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstActionFinder
          mvDataFinder = New DataFinder
          With mvDataFinder
            .Init(mvEnv, DataFinder.DataFinderTypes.dftActions)
            mvResultColumns = .AvailableColumns
            mvSelectColumns = .DisplayColumns
            mvHeadings = .SSHeadings
            mvWidths = .DisplayWidths
          End With
          mvCode = "ACTF"
          mvDescription = DataSelectionText.String18146     'Action Finder
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDisplayTitle = "Action Finder"
          mvRequiredItems = "ActionStatus,Access,MasterAction"

        Case DataSelectionTypes.dstEventFinder
          mvDataFinder = New DataFinder
          With mvDataFinder
            .Init(mvEnv, DataFinder.DataFinderTypes.dftEvent)
            mvResultColumns = .AvailableColumns
            mvSelectColumns = .DisplayColumns
            mvHeadings = .SSHeadings
            mvWidths = .DisplayWidths
          End With
          mvCode = "EVEF"
          mvDescription = "Event Finder"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDisplayTitle = "Event Finder"
          mvRequiredItems = "StartDate,EventDesc,NumberOfAttendees,MaximumAttendees"

        Case DataSelectionTypes.dstContactFinder
          mvResultColumns = "ContactNumber,Title,Forenames,Initials,Surname,Honorifics,Salutation,LabelName,PreferredForename,ContactType"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIrishGiftAid) Then mvResultColumns = mvResultColumns & ",NiNumber"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & ",PrefixHonorifics,SurnamePrefix,InformalSalutation"
          mvResultColumns = mvResultColumns & ",DiallingCode,StdCode,Telephone,ExDirectory,DateOfBirth,Department,Status,OwnershipGroup,AddressNumber,Address,HouseName,Town,County,Postcode,Branch,Country,AddressType"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & ",BuildingNumber"
          mvResultColumns = mvResultColumns & ",ContactName,ContactTelephone,Position,Location,Name,OrganisationNumber,StatusDesc,RgbStatus"
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then mvResultColumns = mvResultColumns & ",OwnershipGroupDesc,PrincipalDepartmentDesc,OwnershipAccessLevel,OwnershipAccessLevelDesc"
          mvSelectColumns = "ContactNumber,Title,Forenames,Surname,Position,Name,Town,Postcode"
          'Must have ContactName for the WEB application at present - need to hide it for the Smart Client
          mvHeadings = "Number,Title,Forenames,Surname,Position,Name,Town,Postcode"         ',ContactName"
          mvWidths = "800,800,1200,2000,2000,2000,1200,1000"
          mvCode = "CFND"
          mvDescription = "Contact Finder"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvRequiredItems = "Status,RgbStatus"

        Case DataSelectionTypes.dstOrganisationFinder
          mvResultColumns = "OrganisationNumber,ContactNumber,Name,SortName,Abbreviation,DiallingCode,StdCode,Telephone,Status,OwnershipGroup,AddressNumber,Address,HouseName,Town,County,Postcode,Branch,Country,AddressType"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & ",BuildingNumber"
          mvResultColumns = mvResultColumns & ",StatusDesc"
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then mvResultColumns = mvResultColumns & ",OwnershipGroupDesc,PrincipalDepartmentDesc,OwnershipAccessLevel,OwnershipAccessLevelDesc"
          mvSelectColumns = "OrganisationNumber,Name,Town,Postcode"
          mvWidths = "800,2000,2000,1000"
          mvHeadings = "Number,Name,Town,Postcode"
          mvCode = "OFND"
          mvDescription = "Organisation Finder"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvRequiredItems = "Status"

        Case DataSelectionTypes.dstServiceStartDays
          mvResultColumns = "StartDayNumber,UniqueID,StartDay,DurationDays,ValidFrom,ValidTo"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1000,1000,1000,1200,1200"

        Case DataSelectionTypes.dstServiceControlRestrictions
          mvResultColumns = "ServiceRestrictionNumber,ContactNumber,ShortStayDuration,LateBookingDays,ValidFrom,ValidTo"

        Case DataSelectionTypes.dstGeographicalRegions
          mvResultColumns = "GeographicalRegion,GeographicalRegionDesc,CollectionRegionNumber,CollectionNumber"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1000,1000,1000,1000"

        Case DataSelectionTypes.dstSalesContacts
          mvResultColumns = "ContactNumber," & ContactNameResults()
          mvSelectColumns = "ContactNumber,ContactName"
          mvHeadings = DataSelectionText.String22887     'Sales Contact,Name
          mvWidths = "1000,2000"

        Case DataSelectionTypes.dstPersonnelContacts
          mvResultColumns = String.Format("ContactNumber,{0},WorkingHours,Evenings,Weekends,Notes", ContactNameResults())
          mvSelectColumns = "ContactNumber,ContactName"
          mvHeadings = DataSelectionText.String22888     'Contact Number,Name
          mvWidths = "1000,2000"

        Case DataSelectionTypes.dstContactAppointments
          mvResultColumns = "ContactNumber,StartDate,EndDate,RecordType,UniqueId,Description,TimeStatus,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactNumber,StartDate,EndDate,Description,TimeStatus"
          mvHeadings = DataSelectionText.String22889     'Contact Number,Start Date,End Date,Description,TimeStatus
          mvWidths = "1,1200,1200,2000,900"

        Case DataSelectionTypes.dstContactAccounts
          mvResultColumns = "BankDetailsNumber,AddressNumber,SortCode,AccountNumber,AccountName,BankPayerName,Notes,IbanNumber,ContactNumber,ContactName"
          mvSelectColumns = "AccountName,ContactName,Notes"
          mvHeadings = DataSelectionText.String22890     'Account Name,Contact Name,Notes
          mvWidths = "2000,2000,2000"

        Case DataSelectionTypes.dstTransactionDetails
          mvResultColumns = "BatchType,PostedToNominal,BatchNumber,TransactionNumber,ContactNumber,AddressNumber,TransactionDate,TransactionType,BankDetailsNumber,Amount,CurrencyAmount,PaymentMethod,Reference,NextLineNumber,LineTotal,Mailing,Receipt,Notes,MailingContactNumber,MailingAddressNumber,EligibleForGiftAid,PaymentMethodDesc,AmendedBy,AmendedOn,"
          mvResultColumns = mvResultColumns & "Title,Forenames,Initials,Surname,Honorifics,Salutation,LabelName,PreferredForename,ContactType,"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIrishGiftAid) Then mvResultColumns = mvResultColumns & "NINumber,"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & "PrefixHonorifics,SurnamePrefix,InformalSalutation,"
          mvResultColumns = mvResultColumns & "Address,HouseName,Town,County,Postcode,Branch,CountryCode,AddressType,"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then mvResultColumns = mvResultColumns & "BuildingNumber,"
          mvResultColumns = mvResultColumns & "Sortcode,Uk,CountryDesc,OrgName,Adjustment,RecordType,AccessLevel,TransactionSign,Printed,ContactName,AddressLine,Allocated"
          mvSelectColumns = "BatchNumber,Adjustment,Allocated,RecordType,ContactNumber,AddressNumber,TransactionNumber,Reference,CurrencyAmount,Amount,TransactionType,TransactionDate,PaymentMethodDesc,Mailing,Receipt,EligibleForGiftAid,LineTotal,AmendedBy,AmendedOn,Surname,AccessLevel,DetailItems,ContactName,OrgName,AddressLine"
          mvHeadings = DataSelectionText.String22891     'Batch,Adjustment,Allocated,RecordType,Contact,Address,No,Reference,Currency Amount,Amount,Type,Date,Method,Mailing,Receipt,Eligible,Total,Amended By,Amended On,Surname,AccessLevel,DetailItems,Contact,Organisation,Address
          mvWidths = "1,1,1,1,1,1,600,1900,900,900,500,1200,700,700,800,800,900,1200,1200,1900,1,1200,1200,1200,1200"
          mvRequiredItems = "Adjustment,Allocated,RecordType,ContactName,AddressLine,OrgName,AccessLevel,TransactionSign,Printed"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18672     'Transaction Details
          mvDisplayTitle = DataSelectionText.String18673     'Batch Transactions
          mvCode = "BT"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactScores
          mvResultColumns = "Score,ScoreDesc,Points,Sequence,Notes"
          mvSelectColumns = mvResultColumns
          mvHeadings = DataSelectionText.String22892     'Score,Description,Points,Sequence No,Notes
          mvWidths = "900,1800,1000,900,2000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CSCO"
          mvDescription = DataSelectionText.String18674     'Contact Scores
          vPrimaryList = True

        Case DataSelectionTypes.dstContactPerformances
          mvResultColumns = "Performance,PerformanceDesc,Sequence,Notes"
          mvResultColumns = mvResultColumns & ",NumberOfPayments,ValueOfPayments,NumberAbove,ValueAbove,NumberBetween,ValueBetween,NumberBelow,ValueBelow,NumberOfMailings,NoResponse,FirstPaymentDate"
          mvResultColumns = mvResultColumns & ",FirstPayment,LastPaymentDate,LastPayment,MaximumPaymentDate,MaximumPayment,RollingValue,PrecedingRollingValue,AverageValue,AveragePerMailing,ResponseRate"
          mvResultColumns = mvResultColumns & ",UpperLevel,LowerLevel,RollingBoundary,StdNumberOfPayments_,StdValueAbove_,StdValueBetween_,StdValueBelow_,StdNumberOfMailings_,StdNoResponse_,StdFirstPayment_"
          mvResultColumns = mvResultColumns & ",StdLastPayment_,StdMaximumPayment_,StdRollingValue_,StdPrecedingRollingValue_,StdAverageValue_,StdAveragePerMailing_,StdResponseRate_"

          mvSelectColumns = "Performance,PerformanceDesc,Sequence,Notes,DetailItems"
          mvSelectColumns = mvSelectColumns & ",NumberOfPayments,NumberAbove,NumberBetween,NumberBelow,NumberOfMailings,NoResponse,FirstPaymentDate,LastPaymentDate,MaximumPaymentDate,RollingValue,PrecedingRollingValue,AverageValue,AveragePerMailing,ResponseRate,NewColumn"
          mvSelectColumns = mvSelectColumns & ",ValueOfPayments,ValueAbove,ValueBetween,ValueBelow,Spacer,Spacer1,FirstPayment,LastPayment,MaximumPayment,NewColumn2"
          mvSelectColumns = mvSelectColumns & ",StdNumberOfPayments_,StdValueAbove_,StdValueBetween_,StdValueBelow_,StdNumberOfMailings_,StdNoResponse_,StdFirstPayment_,StdLastPayment_,StdMaximumPayment_,StdRollingValue_,StdPrecedingRollingValue_,StdAverageValue_,StdAveragePerMailing_,StdResponseRate_"

          mvHeadings = DataSelectionText.String22893     'Performance,Description,Sequence,Notes,
          mvHeadings = mvHeadings & DataSelectionText.String22894     ',Total payments,Number above,Number between,Number below,Number of mailings,No response,First payment date,Last payment date,Maximum payment date,Rolling value,Preceding rolling value,Average value,Average per mailing,Response rate,
          mvHeadings = mvHeadings & DataSelectionText.String22895    ',Total value,Value above,Value between,Value Below,, ,First payment,Last payment,Maximum payment,
          mvHeadings = mvHeadings & ",,,,,,,,,,,,,,"

          mvWidths = "900,1800,900,2000,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvWidths = mvWidths & ",300,300,300,300,300,300,300,300,300,300,300,300,300,300"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CPER"
          mvDescription = DataSelectionText.String18675     'Contact Performances
          vPrimaryList = True

        Case DataSelectionTypes.dstContactSourceFromLastMailing
          mvResultColumns = "Source,SourceDesc,MailingDate,Mailing"
          mvSelectColumns = "Source,SourceDesc,MailingDate,Mailing"
          mvHeadings = DataSelectionText.String22896     'Source,Desc,MailingDate
          mvRequiredItems = "Source"
          mvWidths = "1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String18676     'Source From Last Mailing
          mvDisplayTitle = DataSelectionText.String18676     'Source From Last Mailing
          mvCode = "SLM"

        Case DataSelectionTypes.dstContactAppropriateCertificates
          mvResultColumns = "ContactNumber,CertificateNumber,ClaimNumber,StartDate,EndDate,CertificateAmount,TaxStatus,SignatureDate,AmountClaimed,AmountPaid,CancellationReason,CancelledBy,CancelledOn,CancellationSource,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CancellationReasonDesc,CancellationSourceDesc,TaxStatusCode"
          mvSelectColumns = "ContactNumber,CertificateNumber,ClaimNumber,StartDate,EndDate,CertificateAmount,TaxStatus,SignatureDate,AmountClaimed,AmountPaid,CancelledOn,DetailItems,CreatedBy,AmendedBy,CancelledBy,CancellationSource,NewColumn,CreatedOn,AmendedOn,CancellationReason"
          mvHeadings = "ContactNumber,Number,Claim Number,Start Date,End Date,Amount,Tax Status,Sign Date,Amount Claimed,Amount Paid,Cancelled On,DetailItems,Created By,Amended By,Cancelled By,Cancellation Source,NewColumn,On,On,Reason"
          mvWidths = "1,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300,300"
          mvDescription = "Contact Irish Gift Aid Certificates"
          mvCode = "CAPC"
          'vPrimaryList = True

        Case DataSelectionTypes.dstContactFundraisingEvents
          mvResultColumns = "ContactFundraisingNumber,ContactNumber,FundraisingDescription,Source,SourceDesc,TargetAmount,TargetDate,EventNumber,EventDesc,WebPageNumber,AmendedBy,AmendedOn,DonationTotal,GiftAidTotal,ThankYouMessage"
          mvSelectColumns = "FundraisingDescription,TargetDate,TargetAmount,SourceDesc,EventDesc,WebPageNumber,AmendedBy,AmendedOn"
          mvHeadings = "Description,Date,Target Amount,Source,Event,Web Page,Amended By,Amended On"
          mvWidths = "2000,1200,1200,1800,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Fundraising Events"
          mvMaintenanceDesc = "Fundraising Event"
          mvCode = "FRE"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactFundraisingEventFinder
          mvResultColumns = String.Format("ContactFundraisingNumber,FundraisingDescription,Source,SourceDesc,TargetAmount,TargetDate,EventNumber,EventDesc,EventReference,WebPageNumber,AmendedBy,AmendedOn,WebURL,DescriptionLink,ContactNumber,{0}", ContactNameResults)
          mvSelectColumns = "ContactName,DescriptionLink,TargetDate,TargetAmount"
          mvHeadings = "Name,Description,Date,Target Amount"
          mvWidths = "1800,2000,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Fundraising Event Results"
          mvCode = "FREF"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactFundRaisingRequests
          mvResultColumns = "FundraisingRequestNumber,ContactNumber,RequestDate,RequestDescription,FundraisingRequestStage,FundraisingRequestStageDesc,FundraisingStatus,FundraisingStatusDesc,FundraisingRequestType,FundraisingRequestTypeDesc,Source,SourceDesc,TargetAmount,PledgedAmount,PledgedDate,ReceivedAmount,ReceivedDate,Notes,AmendedBy,AmendedOn,RequestEndDate,ExpectedAmount,GikExpectedAmount,GikPledgedAmount,GikPledgedDate,TotalGikReceivedAmount,LatestGikReceivedDate,NumberOfPayments,Logname,CreatedBy,CreatedOn,TargetDate,TotalExpectedAmount,TotalPledgedAmount,TotalReceivedAmount,OutstandingIncomePledgedAmount,OutstandingGikPledgedAmount,TotalOutstandingPledgedAmount,HasAction,TotalGikScheduledAmount,FundraisingBusinessType,FundraisingBusinessTypeDesc"
          mvSelectColumns = "RequestDate,TargetAmount,RequestDescription,FundraisingRequestStageDesc,FundraisingStatusDesc,FundraisingRequestTypeDesc,SourceDesc,DetailItems,PledgedAmount,ReceivedAmount,Notes,NewColumn,PledgedDate,ReceivedDate,HasAction"
          mvHeadings = "Date,Target Amount,Description,Stage,Status,Type,Source,,Pledged Amount,Received Amount,Notes,,Pledged Date,Received Date,Action?"
          mvWidths = "1000,1200,2000,1200,1200,1200,1200,300,1200,1200,1200,300,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Fundraising Requests"
          mvMaintenanceDesc = "Fundraising Request"
          mvCode = "FRR"
          vPrimaryList = True
          mvRequiredItems = "FundraisingStatus,PledgedAmount,ExpectedAmount,GikPledgedAmount,GikExpectedAmount,Logname,HasAction,TotalGikScheduledAmount"

        Case DataSelectionTypes.dstSelectionSetAppointments
          mvResultColumns = "ContactNumber,ContactName,RequestStageDesc,StartDate,EndDate,RecordType,UniqueId,Description,TimeStatus,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactNumber,ContactName,RequestStageDesc,StartDate,EndDate,Description,TimeStatus"
          mvHeadings = DataSelectionText.String40004     'Contact Number,Contact Name,Stage,Start Date,End Date,Description,TimeStatus
          mvWidths = "1,2000,1200,1200,1200,2000,900"

        Case DataSelectionTypes.dstFundraisingRequestTargets
          mvResultColumns = "TargetAmount,PreviousTargetAmount,ChangeReason,ChangedOn,ChangeBy"
          mvSelectColumns = "PreviousTargetAmount,ChangeReason,ChangedOn,ChangeBy"
          mvHeadings = DataSelectionText.String40005     'Previous Amount,Change Reason,Changed By,Changed On
          mvWidths = "1200,3000,1200,1200"
          mvDisplayTitle = "Previous Targets"

        Case DataSelectionTypes.dstFundRequestExpectedAmountHistory
          mvResultColumns = "FundraisingRequestNumber,ExpectedAmount,PreviousExpectedAmount,ChangeReason,IsIncomeAmount,ChangedOn,ChangeBy"
          mvSelectColumns = "PreviousExpectedAmount,ChangeReason,IsIncomeAmount,ChangedOn,ChangeBy"
          mvHeadings = "Previous Amount,Change Reason,Income Amount?,Changed By,Changed On"
          mvWidths = "1200,1200,1200,1200,1200"
          mvDisplayTitle = "Previous Expected Amounts"

        Case DataSelectionTypes.dstFundRequestStatusHistory
          mvResultColumns = "FundraisingRequestNumber,FundraisingStatus,FundraisingStatusDesc,PreviousFundraisingStatus,PreviousFundraisingStatusDesc,ChangeReason,ChangedOn,ChangeBy"
          mvSelectColumns = "PreviousFundraisingStatus,ChangeReason,ChangedOn,ChangeBy"
          mvHeadings = "Previous Status,Change Reason,Changed By,Changed On"
          mvWidths = "1200,1200,1200,1200"
          mvDisplayTitle = "Previous Statuses"

        Case DataSelectionTypes.dstContactCommunicationHistory
          Dim vMailingDataSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactMailings, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)
          mvResultColumns = vMailingDataSelection.mvResultColumns
          Dim vDocumentDataSelection As DataSelection = New DataSelection(mvEnv, DataSelectionTypes.dstContactDocuments, Nothing, DataSelectionListType.dsltDefault, DataSelectionUsages.dsuSmartClient)
          Dim vMailingResultColumn As List(Of String) = New List(Of String)(mvResultColumns.Split(","c))
          Dim vDocumentResultColumn As List(Of String) = New List(Of String)(vDocumentDataSelection.mvResultColumns.Split(","c))

          Dim vNewDocumentResultColumn As String = String.Empty
          For Each vMailingItem As String In vMailingResultColumn
            If vDocumentResultColumn.Contains(vMailingItem) Then vDocumentResultColumn.Remove(vMailingItem)
          Next

          For Each vDocumentItem As String In vDocumentResultColumn
            vNewDocumentResultColumn += "," + vDocumentItem
          Next
          mvResultColumns += vNewDocumentResultColumn

          mvSelectColumns = "MailingNumber,Date,Type,Mailing,Description,MailedBy,ContactName,ContactNumber,"
          mvSelectColumns += "Direction,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc,DetailItems,Source,PackageCode,DepartmentDesc,StandardDocument,DocumentClassDesc,TheirReference"
          mvHeadings = DataSelectionText.String40009 'NumberID,Date,Type,Mailing,Description,Mailed By,Contact,Contact No.,Direction,Reference+,Document Type,Topic+,Sub Topic,Details Document,Source,Package Code,Department,Document, Document Class,Their Reference
          mvWidths = "1200,1200,1000,1500,2000,1200,1500,1,1200,1200,1500,1500,1500,1400,1400,1200,1400,1200,600,1200,1400,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String40010    'Contact Communication History
          mvCode = "CMH"
          vPrimaryList = True
          mvRequiredItems = "Type,MailingFilename,Access"

        Case DataSelectionTypes.dstFundraisingPaymentSchedule
          mvResultColumns = "FundraisingRequestNumber,ScheduledPaymentNumber,ScheduledPaymentDesc,PaymentAmount,DueDate,FundraisingPaymentType,FundraisingPaymentTypeDesc,FundIncomePaymentType,FundIncomePaymentTypeDesc,ReceivedAmount,ReceivedDate,Source,SourceDesc,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn,HasAction"
          mvSelectColumns = "ScheduledPaymentNumber,ScheduledPaymentDesc,PaymentAmount,DueDate,FundraisingPaymentTypeDesc,FundIncomePaymentTypeDesc,ReceivedAmount,ReceivedDate,Source,Notes,HasAction"
          mvHeadings = "Payment Number,Payment Desc,Amount,Due Date,Payment Type Desc,Income Type Desc,Received Amount,Received Date,Source,Notes,Action?"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Payment Schedule"
          mvMaintenanceDesc = "Scheduled Payment"
          mvDescription = DataSelectionText.String40024 'Fundraising Payment Schedule
          mvCode = "FPS"
          mvRequiredItems = "FundraisingPaymentType,FundIncomePaymentType,ReceivedDate,HasAction"

        Case DataSelectionTypes.dstFundraisingPaymentHistory
          mvResultColumns = "ScheduledPaymentNumber,BatchNumber,TransactionNumber,LineNumber,TransactionDate,Product,ProductDesc,Rate,RateDesc,TransactionType,TransactionTypeDesc,PaymentMethod,DistributionCode,Quantity,Amount,VatAmount,Source,CurrencyAmount,CurrencyVatAmount,Notes,TransactionSign,ContactNumber,ContactName"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,Product,Rate,Amount,VatAmount,Quantity,TransactionType,PaymentMethod,DistributionCode,Source,Notes"
          mvHeadings = "Batch Number,Transaction Number,Line Number,Transaction Date,Product,Rate,Amount,Vat Amount,Quantity,Transaction Type,Payment Method,Distribution Code,Source,Notes"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String40023 'Fundraising Payment History
          mvDisplayTitle = "Payment History"
          mvCode = "FPH"

        Case DataSelectionTypes.dstFundraisingActions
          mvResultColumns = "FundraisingRequestNumber,ScheduledPaymentNumber,MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ScheduledPaymentNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,Deadline,CreatedBy,ScheduledOn,CreatedOn,CompletedOn"
          mvWidths = "1200,1200,2000,1200,1200,1200,1200,1200,1200,1200,1200"
          mvHeadings = "Payment Number,Number,Description,Priority,Status,Action Text,Deadline,Created By,Scheduled On,Created On,Completed On"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Actions"
          mvDescription = DataSelectionText.String40022  'Fundraising Actions
          mvCode = "FPA"
          mvRequiredItems = "ActionStatus,MasterAction"

        Case DataSelectionTypes.dstFundraisingDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc,Precis,Subject,CallDuration,TotalDuration,SelectionSet"
          mvSelectColumns = "DocumentNumber,DocumentTypeDesc,DocumentType,DocumentClass,DocumentClassDesc,ContactNumber"
          mvWidths = "1200,1200,2000,1200,1200,1200,1200"
          mvHeadings = "Document Number,Document Type Description,Document Type,Document Class,Document Class Description,Contact Number"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Documents"
          mvDescription = "Fundraising Documents"
          mvCode = "FRD"
          mvRequiredItems = "DocumentNumber"

        Case DataSelectionTypes.dstTopicDataSheet
          mvResultColumns = "TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,Mandatory,QuantityRequired,MultipleValues,PrimaryTopic"
          mvSelectColumns = mvResultColumns
          mvHeadings = mvResultColumns
          mvWidths = "1,1800,1,1800,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuCare
          mvCode = "TODS"

        Case DataSelectionTypes.dstMembershipGroups
          mvResultColumns = "MembershipGroupNumber,MembershipNumber,OrganisationNumber,Name,DefaultGroup,ValidFrom,ValidTo,IsCurrent"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Number,Membership Number,Organisation Number,Name,Default,Valid From,Valid To,Current"
          mvWidths = "1,1,1,1800,900,900,900,900"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Groups"
          mvCode = "MGP"

        Case DataSelectionTypes.dstMembershipGroupHistory
          mvResultColumns = "HistoryNumber,MembershipNumber,OldGroupName,NewGroupName,ChangeDate"
          mvSelectColumns = mvResultColumns
          mvHeadings = "History Number,Membership Number,Old Group Name,New Group Name,Change Date"
          mvWidths = "1,1,1800,1800,900"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Group History"
          mvCode = "MGH"

        Case DataSelectionTypes.dstFundraisingEventAnalysis
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,Amount,Notes,ContactNumber,ContactName"
          mvSelectColumns = "TransactionDate,ContactName,Amount,Notes,BatchNumber,TransactionNumber,LineNumber"
          mvHeadings = "Date,Name,Amount,Notes,Batch No.,Trans No.,Line No."
          mvWidths = "1200,1200,900,3000,900,900,900"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Analysis"
          mvDescription = "Fundraising Event Analysis"
          mvCode = "FEA"

        Case DataSelectionTypes.dstBatchProcessingInformation
          mvResultColumns = "BatchNumber,BatchType,BankAccount,BatchDate,NumberOfEntries,BatchTotal,NumberOfTransactions,TransactionTotal,DetailCompleted,ReadyforBanking,PayingInSlipPrinted,PayingInSlipNumber,Picked,PostedToCashBook,PostedToNominal,CurrencyBatchTotal,CurrencyTransactionTotal,CurrencyExchangeRate,CurrencyCode,JobNumber,PaymentMethod,BatchCategory,Provisional,ClaimSent,TransactionType,Product,Rate,Source,CashBookBatch,JournalNumber,AmendedBy,AmendedOn,BalancedBy,BalancedOn,PostedBy,PostedOn,ContentsAmendedBy,ContentsAmendedOn,HeaderAmendedBy,HeaderAmendedOn,BatchCreatedBy,BatchCreatedOn,BankingDate,BatchAnalysisCode,PrintChequeList,Company,PickingListNumber,Department,DepartmentDesc,BankAccounts"
          mvSelectColumns = "BatchNumber,BatchType,BankAccount,BatchDate,NumberOfEntries,BatchTotal,NumberOfTransactions,TransactionTotal,DetailCompleted,ReadyforBanking,PayingInSlipPrinted,PayingInSlipNumber,Picked,PostedToCashBook,PostedToNominal,CurrencyBatchTotal,CurrencyTransactionTotal,CurrencyExchangeRate,CurrencyCode,JobNumber,PaymentMethod,BatchCategory,Provisional,ClaimSent"
          mvRequiredItems = mvResultColumns
          mvHeadings = "Batch Number,Type,Bank Account,Batch Date,Number of Entries,Batch Total,Number of Transactions,Transaction Total,Detail Completed,Ready for Banking,Paying in Slip Printed,Paying in Slip Number,Picked,Cash Book,Posted,Currency Batch Total,Currency Transaction Total,Currency Exchange Rate,Currency Code,Job Number,Payment Method,Category,Provisional,Claim Sent"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Batch Processing Information"
          mvDisplayTitle = "Batch Processing Information"
          mvCode = "BAPI"

        Case DataSelectionTypes.dstPickingListDetails

          mvResultColumns = "PickingListNumber,Product,Quantity,Shortfall,OriginalShortfall,ConfirmedOn,Warehouse,WarehouseDesc"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Picking List Number,Product,Quantity,Shortfall,Original Shortfall,Confirmed On,Warehouse,Warehouse Desc"
          mvWidths = "1,1200,1200,1200,1200,1,1,1"

        Case DataSelectionTypes.dstContactMailingDocumentsFinder
          mvResultColumns = "MailingDocumentNumber,MailingTemplate,MailingTemplateDesc,LabelName,CreatedBy,CreatedOn,Mailing,MailingDesc,EarliestFulfilmentDate,FulfillmentNumber,FulfilledBy,FulfilledOn,NumberOfDocuments"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Mailing Document Number,Mailing Template,Mailing Template Desc,Label Name,Created By,Created On,Mailing,Mailing Desc,Earliest Fulfilment Date,Fulfilment Number,Fulfilled By,Fulfilled On,NumberOfDocuments"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1"
          mvDescription = "Contact Mailing Documents Finder"
          mvRequiredItems = mvResultColumns
          mvCode = "MDFI"

        Case DataSelectionTypes.dstMailingFinder
          mvResultColumns = "Mailing,MailingDesc,MailingDate,MailingNumber,MailingBy,NumberInMailing,NumberOfEmails,NumberProcessed,NumberFailed,IssueId,NumberBounced,NumberClicked,NumberOpened,Topic,TopicDesc,SubTopic,SubTopicDesc,Subject,EmailJobNumber"
          mvSelectColumns = "Mailing,MailingDesc,MailingDate,MailingNumber,MailingBy,NumberInMailing,NumberOfEmails,NumberProcessed,NumberFailed,IssueId,NumberBounced,NumberClicked,NumberOpened"
          mvHeadings = "Mailing,Description,Mailing Date,Mailing Number,Mailing By,Mailing Count,Email Count,Emails Processed,Emails Failed,Issue ID,Number Bounced,Number Clicked,Number Opened"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Mailing Finder"
          mvCode = "MAFI"

        Case DataSelectionTypes.dstActionLinkEmailAddresses
          mvResultColumns = "LinkType,ContactNumber,ContactName,Notified,ContactType,LinkTypeDesc,EMailAddress"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Link Type,Contact Number,Contact Name,Notified,Contact Type,Link Description,EMail Address"
          mvWidths = "800,1200,1200,800,800,1200,1200"

        Case DataSelectionTypes.dstContactLegacy
          mvResultColumns = "LegacyNumber,LegacyId,LegacyStatus,LegacyStatusDesc,Source,SourceDesc,SourceDate," &
          "WillDate,LastCodicilDate,GrossEstateValue,NetEstateValue,TotalEstimatedValue,AdminExpensesValue,TaxValue," &
          "OtherBequestsValue,NetForProbate,LiabilitiesValue,DateOfDeath,DeathNotificationSource,DeathNotificationDate," &
          "DateOfProbate,ReviewDate,LegacyReviewReason,LegacyReviewReasonDesc,AgencyNotificationDate,AccountsReceived," &
          "AccountsApproved,AgeAtDeath,LeadCharity,InDispute,LegacyDisputeReason,LegacyDisputeReasonDesc,Residue,AmendedBy,AmendedOn," &
          "TotalReceivedValue,OutstandingValue,NotificationContactName"
          mvSelectColumns = "DetailItems,LegacyNumber,LegacyId,ReviewDate,LeadCharity,WillDate,DateOfDeath,AgencyNotificationDate,DeathNotificationDate,InDispute,AccountsReceived,AmendedBy," &
                            "NewColumn,LegacyStatusDesc,SourceDate,LegacyReviewReasonDesc,SourceDesc,LastCodicilDate,AgeAtDeath,DateOfProbate,NotificationContactName,LegacyDisputeReasonDesc,AccountsApproved,AmendedOn," &
                            "NewColumn2,TotalEstimatedValue,TotalReceivedValue,OutstandingValue,GrossEstateValue,LiabilitiesValue,NetForProbate,AdminExpensesValue,TaxValue,NetEstateValue,OtherBequestsValue,Residue"
          mvHeadings = ",Legacy Number,Legacy ID,Review Date,Lead Charity,Will Date,Date of Death,Notified by Agency,Executor Notification,In Dispute,Accounts Received,Amended by," &
                       ",Status,Creation Date,Review Reason,Source,Last Codicil,Age at Death,Probate,By,Dispute Reason,Accounts Approved,Amended on," &
                       ",Total Expected,Total Received,Outstanding,Gross Estate,Liabilities,Net for Probate,Admin Expenses,Tax Payable,Net Estate,Other Bequests,Residue"
          mvWidths = "300,300,300,300,300,300,300,300,300,300,300,300," &
                     "300,300,300,300,300,300,300,300,300,300,300,300," &
                     "300,300,300,300,300,300,300,300,300,300,300,300"
          mvRequiredItems = "Residue,Source"
          vPrimaryList = True
          mvDescription = "Legacy"
          mvCode = "CLEG"

        Case DataSelectionTypes.dstContactLegacyBequests
          mvResultColumns = "LegacyNumber,BequestNumber,BequestDescription,BequestType,BequestTypeDesc,BequestSubType,BequestSubTypeDesc,BequestStatus,BequestStatusDesc,ExpectedValue,EstimatedOutstanding,Estimate,ExpectedFractionQuantity,ExpectedFractionDivision,Product,ProductDesc,Rate,RateDesc,DistributionCode,DistributionCodeDesc,Notes,TotalReceivedValue,ConditionMetDate,ExpectedFraction"
          mvSelectColumns = "BequestNumber,BequestDescription,ExpectedValue,EstimatedOutstanding,Estimate,ExpectedFraction,Product,Rate,DistributionCode,DetailItems,BequestTypeDesc,BequestSubTypeDesc,Notes,NewColumn,BequestStatusDesc,TotalReceivedValue,Spacer"
          mvHeadings = "No,Description,Expected Value,Estimated Outstanding,Estimated,Proportion Expected,Product,Rate,Distribution Code,,Type,Sub Type,Notes,,Status,Received,"
          mvWidths = "600,2000,1600,1700,800,900,1000,1000,1500,300,300,300,300,300,300,300,300"
          vPrimaryList = True
          mvDescription = "Legacy Bequests"
          mvMaintenanceDesc = "Legacy Bequest"
          mvCode = "CLBQ"

        Case DataSelectionTypes.dstContactLegacyAssets
          'Same result columns as contact categories - keep in synch
          mvResultColumns = "ContactNumber,ContactCategoryNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,NoteFlag,Status,Access,StatusOrder"
          mvSelectColumns = "ActivityValueCode,ActivityValueDesc,Quantity,Notes"
          mvHeadings = "Asset Value,Asset,Value,Notes"
          mvWidths = "1,3500,1200,4000"
          mvRequiredItems = "ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,Access,ContactCategoryNumber"
          vPrimaryList = True
          mvDescription = "Legacy Assets"
          mvMaintenanceDesc = "Legacy Asset"
          mvCode = "CLAS"

        Case DataSelectionTypes.dstContactLegacyLinks
          'Same as contact links to
          mvResultColumns = "RelationshipCode,Type1,Type2,RelationshipDesc," & ContactNameResults() & ",ContactNumber,Phone,ValidFrom,ValidTo,Historical,Notes,AmendedBy,AmendedOn,OwnershipGroup,ContactGroup,ContactLinkNumber"
          mvSelectColumns = "Type2,RelationshipCode,RelationshipDesc,ContactName,ContactNumber,Phone"
          mvHeadings = "Contact Type,Relationship Code,Relationship,To,No,Phone"
          mvWidths = "1,1,2000,4000,1200,2000"
          vPrimaryList = True
          mvDescription = "Legacy Links"
          mvMaintenanceDesc = "Legacy Relationship"
          mvRequiredItems = "RelationshipCode,ValidFrom,ValidTo,ContactLinkNumber"
          mvCode = "CLLK"

        Case DataSelectionTypes.dstContactLegacyTaxCertificates
          mvResultColumns = "TaxCertificateNumber,TaxYear,TaxPercent,Reference,DateReceived,GrossAmount,NetAmount,TaxAmount,TaxClaimed,TaxReceived"
          mvSelectColumns = "TaxCertificateNumber,TaxYear,TaxPercent,Reference,DateReceived,GrossAmount,NetAmount,TaxAmount,TaxClaimed,TaxReceived"
          mvHeadings = "Tax Certificate No,Tax Year,Tax Percentage,Reference,Tax Certificate Date,Gross Amount,Net Amount,Tax Amount,Tax Claimed,Tax Received"
          mvWidths = "1200,1200,1200,2000,1800,1500,1500,1500,1500,1500"
          vPrimaryList = True
          mvDescription = "Legacy Tax Certificates"
          mvMaintenanceDesc = "Legacy Tax Certificate"
          mvCode = "CLTC"

        Case DataSelectionTypes.dstContactLegacyExpenses
          mvResultColumns = "BequestNumber,DateReceived,Amount,Notes"
          mvSelectColumns = "BequestNumber,DateReceived,Amount,Notes"
          mvHeadings = "Bequest No,Date Paid,Amount,Notes"
          mvWidths = "900,1800,1800,7000"
          vPrimaryList = True
          mvDescription = "Legacy Expenses"
          mvMaintenanceDesc = "Legacy Expense"
          mvCode = "CLEX"

        Case DataSelectionTypes.dstLegacyBequestForecasts
          mvResultColumns = "BequestNumber,StageMonthsDelay,StagePercentage"
          mvSelectColumns = "StageMonthsDelay,StagePercentage"
          mvHeadings = "Months Delay,Percentage"
          mvWidths = "1200,1100"
          mvDisplayTitle = "Forecasts"
          mvDescription = "Legacy Bequest Forecasts"
          mvCode = "CLBF"
          mvRequiredItems = "StageMonthsDelay,StagePercentage"

        Case DataSelectionTypes.dstLegacyBequestReceipts
          mvResultColumns = "ContactNumber,ReceiptNumber,BatchNumber,TransactionNumber,LineNumber,Date,Amount,Status,Notes"
          mvSelectColumns = "ReceiptNumber,BatchNumber,TransactionNumber,LineNumber,Date,Amount,Status,Notes"
          mvHeadings = "Receipt No,Batch,Transaction,Line,Date,Amount,Status,Notes"
          mvWidths = "900,900,900,900,900,1000,1000,4000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Receipts"
          mvDescription = "Legacy Bequest Receipts"
          mvCode = "CLBR"

        Case DataSelectionTypes.dstContactLegacyActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600"
          mvHeadings = "Number,Description,Priority,Status,Action Text,,Deadline,Created By,,Scheduled On,Created On,,Completed On"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Legacy Actions"
          mvCode = "CLAC"
          mvRequiredItems = "ActionStatus,MasterAction"
          vPrimaryList = True

        Case DataSelectionTypes.dstCampaignCosts
          mvResultColumns = "CampaignCostNumber,Campaign,Appeal,SegmentCollection,CollectionNumber,CampaignCostType,Amount,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "CampaignCostNumber,CampaignCostType,Amount,Notes,AmendedBy,AmendedOn"
          mvHeadings = "Cost Number,Cost Type,Amount,Notes,Amended By,Amended On"
          mvWidths = "1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Campaign Itemised Costs"
          mvDisplayTitle = "Costs"
          mvCode = "CASC"
          mvRequiredItems = "Campaign,Appeal,SegmentCollection"

        Case DataSelectionTypes.dstCampaignRoles
          mvResultColumns = "ContactCampaignRoleNumber,ContactNumber,ContactName,CampaignRole,CampaignRoleDesc,AmendedBy,AmendedOn"
          mvSelectColumns = "ContactCampaignRoleNumber,ContactNumber,ContactName,CampaignRole,CampaignRoleDesc"
          mvHeadings = DataSelectionText.String19045     'Role Number,Contact Number,Name,Campaign Role,Campaign Role Desc
          mvWidths = "1200,1200,2000,2000,2000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.String19046 'Campaign Roles
          mvDisplayTitle = DataSelectionText.String19046 'Campaign Roles
          mvCode = "CASR"
          mvRequiredItems = "ContactNumber,CampaignRole"

        Case DataSelectionTypes.dstEventPersonnelFinder
          mvResultColumns = "ContactNumber,AddressNumber,Surname,Initials"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials"
          mvHeadings = "ContactNumber,AddressNumber,Surname,Initials"
          mvWidths = "1,1,1200,1200"
          mvDescription = "Event Personnel Finder"

        Case DataSelectionTypes.dstEventPersonnelAppointmentFinder
          mvResultColumns = "ContactNumber,StartDate,EndDate"
          mvSelectColumns = "ContactNumber,StartDate,EndDate"
          mvHeadings = "ContactNumber,StartDate,EndDate"
          mvWidths = "1,1200,1200"
          mvDescription = "Event Personnel Appointment Finder"

        Case DataSelectionTypes.dstTextSearch
          mvResultColumns = "RankNumber,ItemNumber,OwnershipGroup,Description,ItemType,ItemSource,ItemText"
          mvSelectColumns = "ItemType,ItemNumber,Description,ItemSource,ItemText"
          mvHeadings = "Type,No,Description,Source,Text"
          mvWidths = "1200,2000,1200,2000,4000"
          mvDescription = "Text Search Results"


        Case DataSelectionTypes.dstEventBookingTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,Product,Rate,TransactionType,PaymentMethod,DistributionCode,Quantity,Amount,VatAmount,VatRate,Source,CurrencyAmount,CurrencyVatAmount,Notes,TransactionSign"
          mvSelectColumns = "TransactionDate,Product,Rate,Quantity,Amount,TransactionType,PaymentMethod,Source,DistributionCode,Notes"
          mvHeadings = "Date,Product,Rate,Qty,Amount,Transaction Type,Payment Method,Source,Distribution Code,Notes"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Event Booking Transactions"
          mvDisplayTitle = "Related Financial Data"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "EBT"
          mvRequiredItems = "TransactionDate,TransactionSign"

        Case DataSelectionTypes.dstContactRegisteredUsers
          mvResultColumns = "UserName,Password,EmailAddress,ContactNumber,LogOnCount,LastLoggedOn,CreatedOn,RegistrationData,SecurityQuestion,SecurityAnswer,AmendedBy,AmendedOn,LastUpdatedOn,ValidFrom,ValidTo,LoginAttempts,LockedOut,PasswordExpiryDate"
          mvSelectColumns = "UserName,EmailAddress,SecurityQuestion,SecurityAnswer,DetailItems,ContactNumber,LogOnCount,LastLoggedOn,CreatedOn,RegistrationData,ValidFrom,ValidTo,AmendedBy,AmendedOn"
          mvHeadings = "Username,Email Address,Security Question,Security Answer,,Contact Number,Log On Count,Last Logged On,Created On,Registration Data,ValidFrom,ValidTo,Amended By,Amended On"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          vPrimaryList = True
          mvCode = "CRU"
          mvDescription = "Registered Users"
          mvMaintenanceDesc = "Registered User"
          mvRequiredItems = "UserName,Password"

        Case DataSelectionTypes.dstMembershipSummaryMembers
          mvResultColumns = "SequenceNumber,AddressNumber,Surname,ContactNumber,MembershipType,Name,Joined,Branch,BranchMember,Applied,DistributionCode,DateOfBirth"
          mvSelectColumns = "SequenceNumber,AddressNumber,Surname,ContactNumber,MembershipType,Name,Joined,Branch,BranchMember,Applied,DistributionCode,DateOfBirth"
          mvHeadings = ",,,Contact Number,Type,Name,Joined,Branch,Branch Member,Applied,Distribution Code,Date Of Birth"
          mvWidths = "1,1,1,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Membership Summary Members"
          mvRequiredItems = "Surname,MembershipType"

        Case DataSelectionTypes.dstContactEmailingsLinks
          mvResultColumns = "RecipientEmailAddress,EmailValidFrom,EmailValidTo,EmailActive,EmailLink,EmailLinkName,ClickedOn,CommunicationNumber"
          mvSelectColumns = "RecipientEmailAddress,EmailValidFrom,EmailValidTo,EmailActive,EmailLink,EmailLinkName,ClickedOn"
          mvHeadings = "Recipient Email Address,Email Valid From,Email Valid To,Email Active,Email Link,Email Link Name,Clicked On"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Contact Email Links"
          mvDisplayTitle = "Email Details"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CEL"

        Case DataSelectionTypes.dstPurchaseInvoiceChequeInformation
          mvResultColumns = "ChequeReferenceNumber,ChequeNumber,Amount,PrintedOn,ReconciledOn,ChequeStatus,ChequeStatusDesc,AllowReissue,ReprintCount,PayeeContactNumber,PayeeAddressNumber,PayeeContactLabelName,PayeeContactName,PayeeAddressLine,CurrencyCode,CurrencyCodeDesc,AdjustmentStatus,CancellationReason,CancellationSource,CancelledBy,CancelledOn,CancellationSourceDesc,CancellationReasonDesc,PopPaymentMethod"
          mvSelectColumns = "ChequeReferenceNumber,ChequeNumber,Amount,PrintedOn,ReconciledOn,ChequeStatus,ChequeStatusDesc,ReprintCount"
          mvHeadings = "Cheque Reference Number,Cheque Number,Amount,Printed On,Reconciled On,Cheque Status,Cheque Status Desc,Reprint Count"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Purchase Invoice Cheque Information"
          mvDisplayTitle = "Payment Details"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "PICQ"
          mvRequiredItems = "PrintedOn,ReconciledOn,PayeeContactNumber,PayeeAddressNumber,AllowReissue"

        Case DataSelectionTypes.dstContactPositionLinks
          mvResultColumns = "ContactPositionNumber,ContactNumber,RelationshipCode,RelationshipDesc," & ContactNameResults() & ",Status,ValidFrom,ValidTo,Notes,Type2,ContactGroup,NotesFlag,StatusOrder"
          mvSelectColumns = "RelationshipDesc,ContactName,Status,ValidFrom,ValidTo,NotesFlag"
          mvHeadings = "Relationship Desc,With,Status,Valid From,Valid To,Notes Flag"
          mvWidths = "1200,1200,1200,1200,1200,1200"
          mvDescription = "Contact Position Links"
          mvDisplayTitle = "Links"
          mvRequiredItems = "RelationshipCode,Type2,ContactGroup"
          mvMaintenanceDesc = "Position Link"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CPL"

        Case DataSelectionTypes.dstContactAlerts
          mvResultColumns = "ContactNumber,ContactAlert,AlertMessageDesc,RgbAlertMessage,ShowAsDialog,AlertMessageType"
          mvSelectColumns = "AlertMessageDesc"
          mvRequiredItems = "RgbAlertMessage,ShowAsDialog,AlertMessageType"

        Case DataSelectionTypes.dstContactNetwork
          mvResultColumns = "NotSupported"
          mvSelectColumns = "NotSupported"
          mvAvailableUsages = DataSelectionUsages.dsuNone

        Case DataSelectionTypes.dstQueryByExampleContacts
          InitQueryByExampleContacts()
        Case DataSelectionTypes.dstQueryByExampleOrganisations
          InitQueryByExampleOrganisations()
        Case DataSelectionTypes.dstQueryByExampleEvents
          InitQueryByExampleEvents()
        Case DataSelectionTypes.dstDespatchNotes
          mvResultColumns = "DespatchNoteNumber,DespatchMethodDesc,DespatchDate,DeliveryCharge,InvoiceNumber,OrderDate,CarrierReference,DespatchMethod,AddressNumber,Address,HouseName,Town,County,Postcode,Branch,Country,AddressType,BuildingNumber,Sortcode,Uk,CountryDesc,ContactNumber,Title,Forenames,Initials,Surname,Honorifics,Salutation,LabelName,PreferredForename,ContactType,NiNumber,PrefixHonorifics,SurnamePrefix,InformalSalutation,AddressLine,ContactName"
          mvSelectColumns = "DespatchNoteNumber,DespatchMethodDesc,DespatchDate,DeliveryCharge,AddressLine,ContactName,DespatchMethod,CarrierReference"
          mvRequiredItems = "DespatchNoteNumber,DespatchMethodDesc,DespatchDate,DeliveryCharge,AddressLine,ContactName,DespatchMethod,CarrierReference"
          mvHeadings = "Despatch Note Number,Despatch Method,Despatch Date,Delivery"
          mvWidths = "1000,1000,1000,1000,1,1,1,1"
        Case DataSelectionTypes.dstDuplicateContactRecords
          mvResultColumns = "MatchOrPotential,ContactNumber2,LabelName2,AddressNumber2,ContactNumber1,LabelName1,AddressNumber1"
          mvSelectColumns = "MatchOrPotential,ContactNumber2,LabelName2,AddressNumber2,ContactNumber1,LabelName1,AddressNumber1"
          mvRequiredItems = "MatchOrPotential,ContactNumber2,LabelName2,AddressNumber2,ContactNumber1,LabelName1,AddressNumber1"
          mvHeadings = "Status,Original,Name,Address,Duplicate,Name,Address"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000"
        Case DataSelectionTypes.dstSelectAwaitListConfirmation
          mvResultColumns = "PickingListNumber,CompanyDesc,Warehouse,WarehouseDesc,CreatedOn"
          mvSelectColumns = "PickingListNumber,CompanyDesc,Warehouse,WarehouseDesc,CreatedOn"
          mvRequiredItems = "PickingListNumber,CompanyDesc,Warehouse,WarehouseDesc,CreatedOn"
          mvHeadings = "Picking List,Company,Warehouse,Warehouse Desc,Created On"
          mvWidths = "1000,1000,1000,1000,1000"
        Case DataSelectionTypes.dstActionContactLinks
        Case DataSelectionTypes.dstPackedProductDataSheet
          mvResultColumns = "LinkProduct,ProductDesc,Warehouse,WarehouseDesc,LastStockCount,CostOfSale,DefaultWarehouse"
          mvSelectColumns = "LinkProduct,ProductDesc,Warehouse,WarehouseDesc,LastStockCount,CostOfSale,DefaultWarehouse"
          mvRequiredItems = "LinkProduct,ProductDesc,Warehouse,WarehouseDesc,LastStockCount,CostOfSale,DefaultWarehouse"
          mvHeadings = "LinkProduct,ProductDesc,Warehouse,WarehouseDesc,LastStockCount,CostOfSale,DefaultWarehouse"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000"

        Case DataSelectionTypes.dstJobSchedules
          mvResultColumns = "JobNo,Description,DueDate,JobProcessor,RunDate,EndDate,Frequency,Status,SubmittedBy,SubmittedOn,Notify,UpdateDates,Information,Command"
          mvSelectColumns = "JobNo,Description,DueDate,JobProcessor,RunDate,EndDate,Frequency,Status,SubmittedBy,SubmittedOn,Notify,UpdateDates,Information,Command"
          mvHeadings = "Job No,Description,Due Date,Job Processor,Run Date,End Date,Frequency,Status,Submitted By,Submitted On,Notify,Update Dates,Information,Command"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000"

        Case DataSelectionTypes.dstJobProcessors
          mvResultColumns = "JobProcessor,Started,Status,MaxConcurrentJobs,PollingInterval,LastPolled"
          mvSelectColumns = "JobProcessor,Started,Status,MaxConcurrentJobs,PollingInterval,LastPolled,Polling,Active"
          mvHeadings = "Job Processor,Started,Status,Max Concurrent Jobs,Polling Interval,Last Polled,Polling,Active"
          mvWidths = "1000,1000,1000,1000,1000,1000,1,1"

        Case DataSelectionTypes.dstConfig
          mvResultColumns = "ConfigName,ConfigValue,Client"
          mvHeadings = "Config Name,Value,Client"
          mvWidths = "1,1000,1000"
          If mvParameters.ContainsKey("ConfigName") Then
            Dim vConfigName As String = mvParameters("ConfigName").Value
            If Not String.IsNullOrWhiteSpace(vConfigName) Then
              If (mvEnv.GetConfigScopeLevel(vConfigName) And Config.ConfigNameScope.Department) > 0 Then
                mvResultColumns += ",Department"
                mvHeadings += ",Department"
                mvWidths += ",1000"
              End If
              If (mvEnv.GetConfigScopeLevel(vConfigName) And Config.ConfigNameScope.User) > 0 Then
                mvResultColumns += ",Logname"
                mvHeadings += ",Logname"
                mvWidths += ",1000"
              End If
            End If
          End If
          mvSelectColumns = mvResultColumns
        Case DataSelectionTypes.dstSystemModuleUsers
          mvResultColumns = "StartTime,Logname,NamedUser,Active,LastUpdatedOn,AccessCount,RefusedAccess,BuildNumber"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Start Time,User,Named,Active,Last Used,Accesses,Refused,Build No"
          mvWidths = "1,2000,1000,1000,2000,1000,1000,1000"

        Case DataSelectionTypes.dstReportData
          mvResultColumns = "ReportNumber,ReportName,ReportCode,Client,Header,Footer,MailMergeOutput,FileOutput,DetailExclusive,Landscape,UseSsrs,MailmergeHeader,ApplicationName,AmendedOn,AmendedBy"
          mvSelectColumns = "ReportNumber,ReportName,ReportCode,Client,Header,Footer,MailMergeOutput,FileOutput,DetailExclusive,Landscape,UseSsrs,MailmergeHeader,ApplicationName,AmendedOn,AmendedBy"
          mvHeadings = "No,Name,Code,Client,Header,Footer,Mail Merge Output,File Output,Detail Exclusive,Landscape,Use Reporting Services,Mail Merge Header,Application Name,Amended On,Amended By"
          mvWidths = "1000,1000,1000,1000,1,1,1,1,1,1,1,1,1,1000,1000"

        Case DataSelectionTypes.dstReportSectionData
          mvResultColumns = "ReportNumber,SectionNumber,SectionName,SectionTypeDesc,TableFlags,SuppressOutput,ControlAttributes,ExclusiveAttributes,SectionSql,AmendedOn,AmendedBy,SectionType"
          mvSelectColumns = "ReportNumber,SectionNumber,SectionName,SectionTypeDesc,TableFlags,SuppressOutput,ControlAttributes,ExclusiveAttributes,SectionSql,AmendedOn,AmendedBy,SectionType"
          mvHeadings = "No,No,Name,Type,Table,Suppress,Control Attrs,Exclusive Attrs,SQL,Amended On,Amended By,Section Type"
          mvWidths = "1,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1"

        Case DataSelectionTypes.dstReportParameters
          mvResultColumns = "ParameterNumber,ParameterDesc,ParameterName,FieldType,ParameterValue,Expression,AmendedOn,AmendedBy"
          mvSelectColumns = "ParameterNumber,ParameterDesc,ParameterName,FieldType,ParameterValue,Expression,AmendedOn,AmendedBy"
          mvHeadings = "No,Description,Name,Type,Value,Expression,Amended On,Amended By"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000"

        Case DataSelectionTypes.dstReportSectionDetail
          mvResultColumns = "ItemNumber,ReportItemTypeDesc,Caption,AttributeName,ParameterName,ItemFormat,ItemAlignment,ItemWidth,ItemNewline,SuppressBlanks,AmendedOn,AmendedBy,ReportItemType"
          mvSelectColumns = "ItemNumber,ReportItemTypeDesc,Caption,AttributeName,ParameterName,ItemFormat,ItemAlignment,ItemWidth,ItemNewline,SuppressBlanks,AmendedOn,AmendedBy,ReportItemType"
          mvHeadings = "No,Type,Caption,Attribute Name,Parameter Name,Format,Align,Width,N,S,Amended On,Amended By,ReportItem Type Code"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1"

        Case DataSelectionTypes.dstReportVersion
          mvResultColumns = "VersionNumber,ChangeDescription,ChangeDate,Logname,AmendedOn,AmendedBy"
          mvSelectColumns = "VersionNumber,ChangeDescription,ChangeDate,Logname,AmendedOn,AmendedBy"
          mvHeadings = "VersionNumber,ChangeDescription,ChangeDate,Logname,Amended On,Amended By"
          mvWidths = "1000,1000,1000,1000,1000,1000"

        Case DataSelectionTypes.dstReportControl
          mvResultColumns = "SequenceNumber,ControlType,TableName,AttributeName,ControlTop,ControlLeft,ControlWidth,ControlHeight,Visible,ResourceId,ControlCaption,CaptionWidth,HelpText,ContactGroup,ParameterName,MandatoryItem,ReadonlyItem,DefaultValue"
          mvSelectColumns = "SequenceNumber,ControlType,TableName,AttributeName,ControlTop,ControlLeft,ControlWidth,ControlHeight,Visible,ResourceId,ControlCaption,CaptionWidth,HelpText,ContactGroup,ParameterName,MandatoryItem,ReadonlyItem,DefaultValue"
          mvHeadings = "No,Type,Name,Attribute,Top,Left,Width,Height,Visible,Id,Caption,Caption Width,Help Text,Contact Group,Parameter Name,Mandatory Item,Readonly Item,Default Value"
          mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000,1,1000,1000,1000,1000,1,1,1,1,1"

        Case DataSelectionTypes.dstOwnershipData
          mvResultColumns = "OwnershipGroup,OwnershipGroupDesc,PrincipalDepartment,DepartmentDesc,PrincipalDepartmentLogname,LabelName,ReadAccessText,ViewAccessText,Notes,AmendedOn,AmendedBy"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Group,Description,Principal Dept,Principal Dept,Principal Logname,Label Name,Read Text,View Text,Notes,Amended On,Amended By"
          mvWidths = "1000,1000,1,1000,1000,1000,1,1,1,1000,1000"

        Case DataSelectionTypes.dstOwnershipUsers
          mvResultColumns = "OwnershipGroup,Logname,FullName"
          mvSelectColumns = "OwnershipGroup,Logname,FullName"
          mvHeadings = "Ownership Group,Logname,Full Name"
          mvWidths = "1,1,1"

        Case DataSelectionTypes.dstOwnershipDepartments
          mvResultColumns = "Department,DepartmentDesc,OwnershipAccessLevelDesc,OwnershipAccessLevel"
          mvSelectColumns = "Department,DepartmentDesc,OwnershipAccessLevelDesc,OwnershipAccessLevel"
          mvHeadings = "Department,Department,Access Level,OwnershipAccessLevel"
          mvWidths = "1,1000,1000,1"

        Case DataSelectionTypes.dstOwnershipUserInformation
          mvResultColumns = "Logname,OwnershipAccessLevel,OwnershipAccessLevelDesc,ValidFrom,ValidTo,AmendedOn,AmendedBy"
          mvSelectColumns = "Logname,OwnershipAccessLevel,OwnershipAccessLevelDesc,ValidFrom,ValidTo,AmendedOn,AmendedBy"
          mvHeadings = "User,Level,Access Level,Valid From,Valid To,Amended On,Amended By"
          mvWidths = "1000,1,1000,1000,1000,1000,1000"
          mvRequiredItems = "OwnershipAccessLevel"

        Case DataSelectionTypes.dstServiceProductContacts
          mvResultColumns = String.Format("ContactNumber,{0},Product,Rate,FixedUnitRate", ContactNameResults())
          mvSelectColumns = "ContactNumber,ContactName"
          mvHeadings = DataSelectionText.String22888     'Contact Number,Name
          mvWidths = "1000,2000"

        Case DataSelectionTypes.dstEmailAutoReplyText
          mvResultColumns = "ParagraphText"
          mvSelectColumns = "ParagraphText"
          mvHeadings = "Paragraph Text"
          mvWidths = "1000"

        Case DataSelectionTypes.dstEntityAlerts
          mvSelectColumns = "EntityAlertNumber,EntityAlertDesc,EntityAlertMessage,SequenceNumber,ShowAsDialog,RgbValue,EmailAddress,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvResultColumns = "EntityAlertNumber,EntityAlertDesc,EntityAlertMessage,SequenceNumber,ShowAsDialog,RgbValue,EmailAddress,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvHeadings = "Entity Alert Number,Description,Alert Message,Sequence Number,Show As Dialog,Rgb Value,Email Address,Created By,Created On,Amended By,Amended On"
          mvWidths = "1,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000"
          mvDescription = "Alerts"
          mvMaintenanceDesc = "Alert"
          mvCode = "MALT"

        Case DataSelectionTypes.dstSalesTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,Provisional,TransactionTypeDesc,TransactionDate,Amount,PaymentMethodDesc,Reference,Mailing,Receipt,EligibleForGiftAid,CurrencyAmount,Notes,TransactionTypeCode,PaymentMethodCode,CurrencyCode,TransactionOrigin,TransactionOriginDesc,BankAccount,BankAccountDesc,RgbBankAccount,RgbAmount,RgbCurrencyAmount,AddressLine"
          mvSelectColumns = "TransactionDate,Amount,TransactionTypeDesc,PaymentMethodCode,Reference,BatchNumber,TransactionNumber,Provisional,DetailItems,Receipt,Notes,NewColumn,Mailing,NewColumn2,EligibleForGiftAid,NewColumn3"
          mvHeadings = DataSelectionText.String17811     'Date,Amount,Type,Payment Method,Reference,Batch,Transaction,Provisional?,,Receipt?,Notes,,Mailing,,Eligible for Gift Aid?,
          mvWidths = "1100,900,1020,1200,1200,900,1095,1005,1200,600,600,1200,600,1200,600,1200"
          mvDescription = DataSelectionText.String17813     'Contact Sales Transactions
          mvCode = "STR"
          vPrimaryList = True
          mvRequiredItems = "TransactionDate,Amount,Provisional,PaymentMethodCode,RgbBankAccount,RgbAmount,RgbCurrencyAmount"

        Case DataSelectionTypes.dstDeliveryTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,Provisional,TransactionTypeDesc,TransactionDate,Amount,PaymentMethodDesc,Reference,Mailing,Receipt,EligibleForGiftAid,CurrencyAmount,Notes,TransactionTypeCode,PaymentMethodCode,CurrencyCode,TransactionOrigin,TransactionOriginDesc,BankAccount,BankAccountDesc,RgbBankAccount,RgbAmount,RgbCurrencyAmount,AddressLine"
          mvSelectColumns = "TransactionDate,Amount,TransactionTypeDesc,PaymentMethodCode,Reference,BatchNumber,TransactionNumber,Provisional,DetailItems,Receipt,Notes,NewColumn,Mailing,NewColumn2,EligibleForGiftAid,NewColumn3"
          mvHeadings = DataSelectionText.String17811     'Date,Amount,Type,Payment Method,Reference,Batch,Transaction,Provisional?,,Receipt?,Notes,,Mailing,,Eligible for Gift Aid?,
          mvWidths = "1100,900,1020,1200,1200,900,1095,1005,1200,600,600,1200,600,1200,600,1200"
          mvDescription = DataSelectionText.String17814    'Contact Delivery Transactions
          mvCode = "STR"
          vPrimaryList = True
          mvRequiredItems = "TransactionDate,Amount,Provisional,PaymentMethodCode,RgbBankAccount,RgbAmount,RgbCurrencyAmount"

        Case DataSelectionTypes.dstSalesTransactionAnalysis
          mvResultColumns = "LineNumber,Product,Rate,DistributionCode,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,CurrencyAmount,CurrencyVatAmount,SalesContactNumber,SalesContactName,ProductNumber,Notes,LineType,RgbAmount,RgbCurrencyAmount,LineTypeDesc,Number"
          mvSelectColumns = "LineTypeDesc,Number,Product,Rate,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,DistributionCode,SalesContactName,Notes"
          mvHeadings = DataSelectionText.String17984     'Line Type,Number,Product,Rate,Qty,Amount,Vat Amount,Vat Rate,Source Code,Source,Distribution Code,Sales Contact,Notes
          mvWidths = "1200,1200,2000,1000,400,1000,1200,1200,1200,3000,1000,3000,3000"
          mvDescription = DataSelectionText.String18000 'Sales Transaction Analysis
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvRequiredItems = "Amount,VatAmount,RgbAmount,RgbCurrencyAmount"
          mvCode = "STA"

        Case DataSelectionTypes.dstDeliveryTransactionAnalysis
          mvResultColumns = "LineNumber,Product,Rate,DistributionCode,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,CurrencyAmount,CurrencyVatAmount,SalesContactNumber,SalesContactName,ProductNumber,Notes,LineType,RgbAmount,RgbCurrencyAmount,LineTypeDesc,Number"
          mvSelectColumns = "LineTypeDesc,Number,Product,Rate,Quantity,Amount,VatAmount,VatRate,Source,SourceDesc,DistributionCode,SalesContactName,Notes"
          mvHeadings = DataSelectionText.String17984     'Line Type,Number,Product,Rate,Qty,Amount,Vat Amount,Vat Rate,Source Code,Source,Distribution Code,Sales Contact,Notes
          mvWidths = "1200,1200,2000,1000,400,1000,1200,1200,1200,3000,1000,3000,3000"
          mvDescription = DataSelectionText.String18001  'Delivery Transaction Analysis
          mvDisplayTitle = DataSelectionText.String17981     'Details
          mvRequiredItems = "Amount,RgbAmount,RgbCurrencyAmount"
          mvCode = "DTA"

        Case DataSelectionTypes.dstEventDelegateMailing
          mvResultColumns = "MailingNumber,ContactNumber,AddressNumber,Date,Direction,Mailing,Description,MailingTemplate,MailingTemplateDesc,Notes,MailingHistoryNotes,MailedBy,MailingFilename,Topic,TopicDesc,SubTopic,SubTopicDesc,Subject,FulfillmentNumber,ProcessedOn,ProcessedStatus,ErrorNumber,CommunicationNumber," & ContactNameResults() & ",CheetahMailId,NumberEmailsBounced,NumberEmailsOpened,NumberEmailsClicked,OpenedOn,Type,AddressLine,OrganisationName"
          mvSelectColumns = "MailingNumber,Date,Type,Mailing,Description,MailedBy,ContactName,ContactNumber"
          mvHeadings = DataSelectionText.String17857     'Number,Date,Type,Mailing,Description,Mailed By,Contact,Contact No.
          mvWidths = "1200,1200,1000,1500,2000,1200,1500,1"
          mvDescription = DataSelectionText.String18180    'Event Delegate Mailing
          mvCode = "EDM"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstCPDObjectives
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,ContactCPDPeriodNumber,PeriodStartDate,PeriodEndDate,PeriodDuration,CpdObjectiveNumber,CPDObjectiveDesc,LongDescription,CategoryType,CategoryTypeDesc,Category,CategoryDesc,CompletionDate,TargetDate,SupervisorContactNumber,SupervisorContactName,SupervisorAccepted,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "CategoryTypeDesc,CategoryDesc,CPDObjectiveDesc,CompletionDate,TargetDate,SupervisorContactName,SupervisorAccepted,CreatedBy,CreatedOn,AmendedOn,AmendedBy"
          mvHeadings = "Category Type,Description,Objective,Completion,Target,Supervisor,Accepted,Created By,Created On,Amended By,Amended On"
          mvDescription = "Contact CPD Objectives"
          mvMaintenanceDesc = "Objective"
          mvCode = "CPOB"

        Case DataSelectionTypes.dstCPDObjectivesEdit
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,ContactCPDPeriodNumber,PeriodStartDate,PeriodEndDate,PeriodDuration,CpdObjectiveNumber,CPDObjectiveDesc,LongDescription,CategoryType,CategoryTypeDesc,Category,CategoryDesc,CompletionDate,TargetDate,SupervisorContactNumber,SupervisorContactName,SupervisorAccepted,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "PeriodDuration,CategoryTypeDesc,CategoryDesc,CPDObjectiveDesc,CompletionDate,TargetDate,SupervisorContactName,SupervisorAccepted,CreatedBy,CreatedOn,AmendedOn,AmendedBy"
          mvHeadings = "Period Duration,Category Type,Description,Objective,Completion,Target,Supervisor,Accepted,Created By,Created On,Amended By,Amended On"
          mvDescription = "Contact CPD Objectives Maintenance"
          mvMaintenanceDesc = "Objective"
          mvCode = "CPOE"

        Case DataSelectionTypes.dstEventActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600"
          mvHeadings = "Number,Description,Priority,Status,Action Text,,Deadline,Created By,,Scheduled On,Created On,,Completed On"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Event Actions"
          mvCode = "EVAC"
          mvRequiredItems = "ActionStatus,MasterAction"
          vPrimaryList = True

        Case DataSelectionTypes.dstAppealActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600"
          mvHeadings = "Number,Description,Priority,Status,Action Text,,Deadline,Created By,,Scheduled On,Created On,,Completed On"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Campaign Actions"
          mvCode = "CAAC"
          mvRequiredItems = "ActionStatus,MasterAction"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactSurveys
          mvResultColumns = "ContactSurveyNumber,SurveyNumber,SurveyShortDescription,SurveyVersion,SentOn,CompletedOn,ValidFrom,ValidTo,ClosingDate,SurveyLongDescription,Notes,SurveyVersionNumber,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "SurveyNumber,SurveyShortDescription,SurveyVersion,SentOn,CompletedOn,Notes,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40018 'Survey Number,Description,Version,Sent On,Completed On,Notes,Amended By,Amended On
          mvDescription = "Contact Surveys"
          mvWidths = "1200,1200,1000,1500,2000,1200,1200,1200,1,1,1,1,1,1,1,1" ' "1200,1200,1000,1500,2000,1200,1,1,1,1,1200,1,1,1200,1500"
          mvMaintenanceDesc = "Survey"
          mvCode = "SURV"

        Case DataSelectionTypes.dstContactSurveyResponses
          mvResultColumns = "SurveyResponseNumber,ContactSurveyNumber,SurveyQuestionNumber,QuestionText,Mandatory,AnswerType,HelpText,SurveyAnswerNumber,AnswerText,AnswerDataType,MinimumValue,MaximumValue,ListValues,NextQuestionNumber,ResponseAnswerText,CreatedBy,CreatedOn,AmendedBy,AmendedOn,DisplayAnswerText,NewPage"
          mvSelectColumns = "QuestionText,AnswerText,DisplayAnswerText,AmendedBy,AmendedOn,AnswerType,AnswerDataType,MinimumValue,MaximumValue,Mandatory,ListValues,ResponseAnswerText,NewPage"
          mvHeadings = DataSelectionText.String40019 'Question,Answer,Answer Text,Amended By,Amended On
          mvDescription = "Contact Survey Responses"
          mvWidths = "1200,1200,1000,1500,2000,1,1,1,1,1,1,1,1,1,1,1,1,1000,1000"
          mvCode = "SURR"
          mvRequiredItems = "SurveyResponseNumber,ContactSurveyNumber,SurveyQuestionNumber,QuestionText,Mandatory,AnswerType,SurveyAnswerNumber,AnswerText,AnswerDataType,MinimumValue,MaximumValue,ListValues,NextQuestionNumber,ResponseAnswerText,NewPage"

        Case DataSelectionTypes.dstContactSalesLedgerReceipts
          mvResultColumns = "Date,TransactionType,InvoiceNumber,Reference,InvoicePayStatus,BatchNumber,TransactionNumber,LineNumber,TransactionSign,Debit,Credit,Outstanding,StoredInvoiceNumber,PayerContactNumber,AdjustmentStatus"
          mvSelectColumns = "Date,TransactionType,Reference,Debit,Credit,BatchNumber,TransactionNumber,AdjustmentStatus"
          mvHeadings = DataSelectionText.String40021    'Date,Type,Reference,Debit,Credit,Batch,Transaction,Status
          mvWidths = "1200,1300,500,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Payments"
          mvRequiredItems = "TransactionType,InvoicePayStatus,TransactionSign"
          mvDescription = "Sales Ledger Receipts"
          mvCode = "CSLR"

        Case DataSelectionTypes.dstWebProducts
          mvResultColumns = "Product,ProductDesc,ExtraKey,SalesGroup,SalesGroupDesc,SecondaryGroup,SecondaryGroupDesc,ProductCategory,ProductCategoryDesc,SalesDescription,Rate,RateDesc,CurrentPrice,FuturePrice,VatExclusive,PriceChangeDate,Percentage,GrossPrice,NetPrice,ProductImage,VatAmount"
          mvSelectColumns = "Product,ProductDesc,RateDesc,GrossPrice,SalesDescription,ProductImage"
          mvHeadings = "Product,Description,Rate,Gross Price,Sales Description,ProductImage"
          mvDescription = "Web Products"
          mvWidths = "1200,1200,1000,1500,2000,1000"
          mvCode = "WPRO"
          mvRequiredItems = "Product,Rate,CurrentPrice,VatExclusive,Percentage"

        Case DataSelectionTypes.dstWebEvents
          mvResultColumns = "EventNumber,EventDesc,Reference,EventImage,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,StartDate,StartTime,EndDate,EndTime,Branch,Venue,VenueDesc,VenueReference,Location,NumberOfAttendees,MaximumAttendees,Status,StatusDesc,BookingsClose,FreeOfCharge,ShortLongDescription,LongDescription,EventClass,EventClassDesc,Notes,BranchCode,BalanceBookingFee,BalanceBookingDue,MinimumSponsorshipAmount,SponsorshipDue,PledgedAmountDue,MultiSession,External,Organiser,PriceToAttendees,EligibilityCheckRequired,EligibilityCheckText,OrganiserReference,Department,NameAttendees,NumberOfBookings,BookingStatusDesc,AvailablePlaces"
          mvSelectColumns = "EventNumber,EventDesc,Reference,StartDate,StartTime,EndDate,EndTime,VenueDesc,LongDescription,BookingStatusDesc,AvailablePlaces"
          mvHeadings = "ID,Name,Reference,Start Date,Start Time,End Date,End Time,Venue,Description,Booking Status,Available Places"
          mvDescription = "Web Events"
          mvCode = "WEVN"
          mvRequiredItems = "EventNumber"
        Case DataSelectionTypes.dstWebBookingOptions
          mvResultColumns = "OptionNumber,OptionDesc,PickSessions,MinimumBookings,MaximumBookings,Product,ProductDesc,ProductMinQuantity,ProductMaxQuantity,Rate,RateDesc,GrossPrice,NetPrice,StartDate,PriceChangeDate,FuturePrice,CurrentPrice,Percentage,LongDescription,DaysPriorTo,DaysPriorFrom,NumberOfSession,VatExclusive,UseModifiers"
          mvSelectColumns = "OptionNumber,OptionDesc,LongDescription,RateDesc,GrossPrice"
          mvHeadings = "Option,Name,Description,Rate,Price"
          mvDescription = "Web Booking Options"
          mvWidths = "1200,1200,1200,1200,1200"
          mvCode = "WEBO"
          mvRequiredItems = "OptionNumber,PickSessions,MaximumBookings,Product,Rate,ProductMinQuantity,ProductMaxQuantity,NetPrice,GrossPrice,Percentage,VatExclusive"
        Case DataSelectionTypes.dstContactEventBookingDelegates
          mvResultColumns = "Select,AddressNumber,ContactNumber,ContactName,Position,OrganisationName,EventDelegateNumber"
          mvSelectColumns = "ContactName,Position,OrganisationName"
          mvHeadings = "Name,Position,Organisation"
          mvDescription = "Contact Event Booking Delegates"
          mvWidths = "1200,1200,1200"
          mvCode = "CBD"
          mvRequiredItems = "Select,ContactNumber"
        Case DataSelectionTypes.dstDocumentActions
          mvResultColumns = "ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText"
          mvWidths = "800,2000,1500,1500,2000"
          mvHeadings = "Number,Description,Priority,Status,Action Text"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Document Actions"
          mvRequiredItems = "ActionStatus"
          mvCode = "CDA"
        Case DataSelectionTypes.dstWebMembershipType
          mvResultColumns = "MembershipType,MembershipTypeDesc,LongDescription"
          mvSelectColumns = "MembershipType,MembershipTypeDesc,LongDescription"
          mvHeadings = "Membership Type,Description,Sales Description"
          mvDescription = "Web Membership Types"
          mvCode = "WMTY"
          mvRequiredItems = "MembershipType"
        Case DataSelectionTypes.dstActivityFromActivityGroup
          mvResultColumns = "ActivityGroup,Activity,QuantityRequired,SequenceNumber,MultipleValues"
          mvSelectColumns = "ActivityGroup,Activity,QuantityRequired,SequenceNumber,MultipleValues"
          mvRequiredItems = "ActivityGroup,Activity"
        Case DataSelectionTypes.dstEventCategories
          mvResultColumns = "DelegateActivityNumber,EventDelegateNumber,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity,Status,SourceDesc,ValidFrom,ValidTo,NoteFlag,DetailItems,Notes,AmendedBy,AmendedOn"
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "DelegateActivityNumber,EventDelegateNumber,ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,Access"
          mvDescription = "Event Categories"
          mvMaintenanceDesc = "Category"
          mvCode = "CC"
          vPrimaryList = True
        Case DataSelectionTypes.dstWebEventBookings
          mvResultColumns = "BookingNumber,BookerName,BookerContactNumber,BookingDate,DelegateName,DelegateContactNumber,Quantity,OptionDesc,EventNumber,EventDesc,Reference,StartDate,StartTime,EndDate,EndTime,Venue,VenueDesc,Location,LongDescription,GrossAmount,Amount,VatAmount,BookingsClose"
          mvSelectColumns = "BookerName,DelegateName,EventDesc,Reference"
          mvHeadings = "Booked By,Delegate,Event,Reference"
          mvWidths = "1000,1000,1000,1000"
          mvDescription = "Web Event Bookings"
          mvCode = "WEVB"
          mvRequiredItems = "BookingNumber,BookerContactNumber,BookingsClose"
        Case DataSelectionTypes.dstServiceBookingTransactions
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionDate,Product,Rate,TransactionTypeDesc,PaymentMethod,DistributionCode,Quantity,Amount,VatAmount,VatRate,Source,CurrencyAmount,CurrencyVatAmount,Notes,TransactionSign"
          mvSelectColumns = "TransactionDate,Product,Rate,Quantity,Amount,VatAmount,VatRate,Source,DistributionCode,Notes"
          mvHeadings = "Date,Product,Rate,Qty,Amount,Vat Amount,Vat Rate,Source,Distribution Code,Notes"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Service Booking Transactions"
          mvDisplayTitle = "Related Financial Data"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "SBT"
          mvRequiredItems = "TransactionDate,TransactionSign"
        Case DataSelectionTypes.dstWebSurveys
          mvResultColumns = "SurveyNumber,SurveyName,LongDescription,Notes,SurveyVersionNumber,ValidFrom,ValidTo,ClosingDate,Source"
          mvSelectColumns = "SurveyName,LongDescription,ClosingDate"
          mvHeadings = "Survey,Description,Closing Date"
          mvWidths = "1000,1000,1000"
          mvDescription = "Web Surveys"
          mvCode = "WSSV"
          mvRequiredItems = "SurveyVersionNumber"
        Case DataSelectionTypes.dstWebDirectoryEntries
          mvResultColumns = "ContactNumber,AddressNumber,LabelName,Position,MemberNumber,Address,Town,Country,County,Sortcode,Postcode,Uk,CountryDesc,Notes,HouseName,Branch,AddressType,BuildingNumber,Title,Forenames,Surname,Initials,Honorifics,Salutation,PreferredForename,ContactType,NiNumber,PrefixHonorifics,SurnamePrefix,InformalSalutation,Category1,Activity1,Category2,Activity2,Category3,Activity3,Category4,Activity4,Category5,Activity5,Category6,Activity6,AddressLine,ContactName"
          mvSelectColumns = "ContactName,AddressLine,Notes,Activity1,Activity2,Activity3,Activity4,Activity5,Activity6"
          mvHeadings = "Name,Address,Description,Activity1,Activity2,Activity3,Activity4,Activity5,Activity6"
          mvDescription = "Web Directory Entries"
          mvWidths = "1000,2000,1000,1000,1000,1000,1000,1000,1000"
          mvCode = "WDEN"
          mvRequiredItems = "ContactNumber,Category1,Activity1,Category2,Activity2,Category3,Activity3,Category4,Activity4,Category5,Activity5,Category6,Activity6"
        Case DataSelectionTypes.dstWebDocuments
          mvResultColumns = "WebDocumentNumber,Title,FileName,WebDocumentExtension,ImageName,Description,MimeType,ViewName,ValidFrom,ValidTo,DownloadCount,LastDownloadedOn,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "Title,Description"
          mvHeadings = "Title,Description"
          mvDescription = "Web Documents"
          mvWidths = "1000,2000"
          mvCode = "WDOC"
          mvRequiredItems = "FileName,ImageName,MimeType"
        Case DataSelectionTypes.dstBankTransactions
          mvResultColumns = "LineNumber,PayersSortCode,PayersAccountNumber,PayersName,ReferenceNumber,Amount,PaymentMethod,PaymentMethodDesc,Notes"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Line No,Sort Code,Account Number,Name,Reference,Amount,Payment Method Code,Payment Method,Notes"
          mvWidths = "1,900,1700,2700,2300,800,1,1200,2000"
        Case DataSelectionTypes.dstWebRelatedContacts
          mvResultColumns = "OrganisationNumber,ContactNumber,AddressNumber,ContactPositionNumber,Title,Initials,Forenames,Surname,ContactName,DateOfBirth,EmailAddress,ValidFrom,ValidTo,Mail,Current,Location,Position,PositionFunction,PositionFunctionDesc,PostionSeniority,PostionSeniorityDesc,Status,StatusDesc,MemberNumber,MembershipStatus,MembershipStatusDesc,NINumber,ContactGroup,ContactGroupDesc,DefaultContactNumber,Default,PhoneNumber"
          mvSelectColumns = "ContactName,Position,StatusDesc"
          mvHeadings = "Name,Position,Status"
          mvWidths = "1000,1000,1000"
          mvDescription = "Web Related Contacts"
          mvCode = "WRCN"
          mvRequiredItems = "OrganisationNumber,ContactNumber,EmailAddress,ContactPositionNumber"
        Case DataSelectionTypes.dstWebRelatedOrganisations
          mvResultColumns = "OrganisationNumber,AddressNumber,OrganisationName,Status,StatusDesc,OrganisationGroup,OrganisationGroupDesc,Address,Town,County,Sortcode,Postcode,Country,CountryDesc,Uk,HouseName,Branch,AddressType,BuildingNumber,ParentOrganisation,ChildCount,AddressLine"
          mvSelectColumns = "OrganisationName,AddressLine"
          mvHeadings = "Organisation,Address"
          mvDescription = "Web Related Organisations"
          mvWidths = "1000,2000"
          mvCode = "WROG"
          mvRequiredItems = "ChildCount,OrganisationNumber"
        Case DataSelectionTypes.dstWebContacts
          mvResultColumns = "AddressNumber,DateOfBirth,EmailAddress,Status,StatusDesc,MemberNumber,MembershipStatus,MembershipStatusDesc,NINumber,ContactGroupDesc,ContactNumber,Title,Forenames,Initials,Surname,ContactName,PhoneNumber"
          mvSelectColumns = "Forenames,Surname"
          mvRequiredItems = "ContactNumber"
          mvHeadings = "Forenames,Surname"
          mvWidths = "1000,1000"
          mvDescription = "Web Contacts"
          mvCode = "WCON"

        Case DataSelectionTypes.dstContactAppointmentDetails
          mvResultColumns = "AppointmentContactNumber,AppointmentRecordType,AppointmentUniqueId,ContactNumber,AddressNumber,BatchNumber,TransactionNumber,LineNumber"
          mvSelectColumns = "ContactNumber,AddressNumber,BatchNumber,TransactionNumber,LineNumber"
          mvHeadings = "ContactNumber,AddressNumber,BatchNumber,TransactionNumber,LineNumber"   'DataSelectionText.String22889     'Contact Number,Start Date,End Date,Description,TimeStatus
          mvWidths = "1200,1200,1200,1200,900"

        Case DataSelectionTypes.dstPurchaseOrderHistory
          mvResultColumns = "PurchaseOrderNumber,Amount,PreviousAmount,PreviousAuthorisationLevel,PreviousAuthorisationLevelDesc,PreviousAuthorisedBy,PreviousAuthorisedOn,ChangedBy,ChangedOn"
          mvSelectColumns = "Amount,PreviousAmount,PreviousAuthorisationLevelDesc,PreviousAuthorisedBy,PreviousAuthorisedOn,ChangedBy,ChangedOn"
          mvHeadings = "Amount,Previous Amount,Previous Authorisation Level,Previous Authorised By,Previous Authorised On,Changed By,Changed On"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String40025
          mvDisplayTitle = "History"
          mvCode = "POH"

        Case DataSelectionTypes.dstContactLoans
          mvResultColumns = "PaymentPlanNumber,ContactNumber,LoanNumber,LoanType,LoanTypeDesc,LoanAmount,Balance,CreatedBy,CreatedOn,StartDate,InterestRate,InterestCalculatedDate,NextPaymentDue,MonthlyPaymentAmount,LoanTerm,FrequencyAmount,Source,SourceDesc,CancellationReason,CancellationReasonDesc,CancellationSource,CancellationSourceDesc,CancelledBy,CancelledOn,DirectDebitStatus,CreditCardStatus,PaymentFrequencyFrequency,PaymentMethod,PaymentMethodDesc,ExpiryDate,InterestCapitalisationDate,InterestCapitalisationAmount"
          mvSelectColumns = "LoanNumber,LoanTypeDesc,LoanAmount,Balance,PaymentMethodDesc,CancellationReasonDesc,CreatedBy,CreatedOn,DetailItems,StartDate,InterestRate,InterestCalculatedDate,CancelledBy,NewColumn,FrequencyAmount,MonthlyPaymentAmount,Spacer,CancelledOn,NewColumn2,NextPaymentDue,LoanTerm,Spacer2,CancellationSourceDesc"
          mvHeadings = "Number,Type,Amount,Balance,Payment Method,Cancellation Reason,Created By,Created On,,Start Date,Interest Rate %,Interest Calculated,Cancelled By,,Next Payment,Monthly Payment,,On,,Due,Loan Term,,Source"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = DataSelectionText.String40026
          mvDisplayTitle = "Loans"
          mvCode = "CLN"
          mvRequiredItems = "CancellationReason,ExpiryDate,InterestCalculatedDate"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstContactDirectoryUsage
          mvResultColumns = "AddressType,Device,AddressUsage,CommunicationUsage,Value"
          mvSelectColumns = "AddressType,Device,AddressUsage,CommunicationUsage,Value"
          mvHeadings = "AddressType,Device,AddressUsage,CommunicationUsage,Value"
          mvWidths = "1200,1200,1200,1200,1200"

        Case DataSelectionTypes.dstDataUpdates
          mvResultColumns = "CheckValue,DataUpdateNumber,BriefDesc,DetailedDesc,SingleApplication,AmendedBy,AmendedOn"
          mvSelectColumns = "CheckValue,DataUpdateNumber,BriefDesc,DetailItems,DetailedDesc"
          mvHeadings = ",Number,Description,,Details"
          mvRequiredItems = "CheckValue"
          mvWidths = "1200,1200,1200,1200"

        Case DataSelectionTypes.dstContactAddressesAndPositions
          'All of the address fields are required to get the AddressLine
          mvResultColumns = "AddressNumber,AddressType,HouseName,Address,Town,County,Postcode,CountryCode,CountryDesc,Branch,PAF,SortCode,UK,GovernmentRegionDesc,BuildingNumber,DeliveryPointSuffix,AmendedBy,AmendedOn,AddressFormat,AddressLine1,AddressLine2,AddressLine3,OrganisationNumber,OrganisationName,AddressLine"
          mvSelectColumns = "AddressLine,OrganisationName"
          mvDescription = "Contact Addresses And Positions"
          mvHeadings = "Address,Organisation Name"
          mvWidths = "1200,1200"
          mvCode = "CAAP"

        Case DataSelectionTypes.dstDuplicateOrganisationsForRegistration
          mvResultColumns = "AddressNumber,AddressType,HouseName,Address,Town,County,Postcode,CountryCode,CountryDesc,Branch,PAF,SortCode,UK,GovernmentRegionDesc,BuildingNumber,DeliveryPointSuffix,AmendedBy,AmendedOn,AddressFormat,AddressLine1,AddressLine2,AddressLine3,OrganisationNumber,OrganisationName,AddressLine"
          mvSelectColumns = "OrganisationName,AddressLine"
          mvHeadings = "Name,Address"
          mvWidths = "1200,1200"
          mvDescription = "Duplicate Organisations"
          mvCode = "DOFR"

        Case DataSelectionTypes.dstContactExamSummary
          mvResultColumns = "ExamStudentHeaderId,ExamUnitId,ExamBaseUnitId,ExamUnitCode,ExamUnitDescription,FirstSessionId,FirstSessionCode,FirstSessionDescription,LastSessionId,LastSessionCode,LastSessionDescription,LastMarkedDate,LastGradedDate,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,FirstSessionCode,FirstSessionDescription,LastSessionCode,LastSessionDescription,LastMarkedDate,LastGradedDate"
          mvHeadings = "Code,Description,First Session,Description,Last Session,Description,Last Marked,Last Graded"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamStudentHeaderId,ExamUnitId,ExamBaseUnitId"
          mvDescription = "Contact Exam Summary"
          mvCode = "CXES"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactExamSummaryItems
          mvResultColumns = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId,ExamUnitCode,ExamUnitDescription,Attempts,CurrentMark,CurrentGrade,CurrentResult,CurrentGradeDesc,CurrentResultDesc,FirstPassed,Expires,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CanEditResults,ResultsReleaseDate,PreviousMark,PreviousGrade,PreviousResult,PreviousGradeDesc,PreviousResultDesc"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLastExamDate) Then mvResultColumns += ",LastExamDate,ChildLastExamDate"
          mvSelectColumns = "DetailItems,ExamUnitCode,ExamUnitDescription,Attempts,CurrentMark,CurrentGradeDesc,CurrentResultDesc,FirstPassed,Expires"
          mvHeadings = ",Code,Description,Attempts,Mark,Grade,Result,First Granted,Expiry Date"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId,CanEditResults"
          mvDisplayTitle = "Items"
          mvDescription = "Contact Exam Summary Items"
          mvCode = "CXSI"

        Case DataSelectionTypes.dstContactExamSummaryList
          mvResultColumns = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId,ExamUnitCode,ExamUnitDescription,Attempts,CurrentMark,CurrentGrade,CurrentResult,CurrentGradeDesc,CurrentResultDesc,FirstPassed,Expires,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CanEditResults,ResultsReleaseDate,PreviousMark,PreviousGrade,PreviousResult,PreviousGradeDesc,PreviousResultDesc"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,Attempts,CurrentMark,CurrentGradeDesc,CurrentResultDesc,FirstPassed,Expires"
          mvHeadings = "Code,Description,Attempts,Mark,Grade,Result,First Granted,Expiry Date"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId,CanEditResults"
          mvDisplayTitle = "Item List"
          mvDescription = "Contact Exam Summary List"
          mvCode = "CXSL"

        Case DataSelectionTypes.dstContactExamDetails
          mvResultColumns = "ExamBookingId,ExamSessionId,ExamCentreId,ExamUnitId,ExamSessionCode,ExamSessionDescription,ExamUnitCode,ExamUnitDescription,ExamCentreCode,ExamCentreDescription,Amount,BatchNumber,TransactionNumber,PayerContactNumber,TransactionDate,CancellationReason,CancellationSource,CancellationReasonDesc,CancellationSourceDesc,CancelledOn,CancelledBy,SpecialRequirements,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CourseStartDate,AssessmentLanguage,AssessmentLanguageDesc,CentreUnitLocalName,StudyMode,StudyModeDesc"
          mvSelectColumns = "ExamSessionCode,ExamSessionDescription,ExamUnitCode,ExamUnitDescription,ExamCentreCode,ExamCentreDescription,Amount,BatchNumber,TransactionNumber,SpecialRequirements,CancellationReasonDesc,CancelledOn,CancelledBy"
          mvHeadings = "Session,Description,Unit,Description,Centre,Description,Amount,Batch,Transaction,Special Requirements,Cancellation Reason,Cancelled On,Cancelled By"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamBookingId,ExamSessionId,ExamCentreId,ExamUnitId,CancellationReason"
          mvDisplayTitle = "Details"
          mvDescription = "Contact Exam Detail Items"
          mvCode = "CXED"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactExamDetailItems
          mvResultColumns = "ExamBookingUnitId,ExamBookingId,ExamUnitId,ExamUnitCode,ExamUnitDescription,ExamCandidateNumber,AttemptNumber,ExamStudentUnitStatus,ExamStudentUnitStatusDesc,OriginalMark,ModeratedMark,TotalMark,OriginalGrade,ModeratedGrade,TotalGrade,TotalGradeDesc,OriginalResult,ModeratedResult,TotalResult,TotalResultDesc,DoneDate,StartDate,StartTime,EndTime,Source,ActivityGroup,CancellationReason,CancellationSource,CancellationReasonDesc,CancellationSourceDesc,CancelledOn,CancelledBy,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,ExamCentreId,CanEditResults,CourseStartDate,ExamAssessmentLanguage,ExamSessionId,ResultsReleaseDate,PreviousMark,PreviousGrade,PreviousResult,PreviousGradeDesc,PreviousResultDesc,StudyMode,CentreUnitLocalName,AchievedUnitCode,AchievedUnitDescription"
          mvSelectColumns = "DetailItems,ExamUnitCode,ExamUnitDescription,ExamCandidateNumber,AttemptNumber,ExamStudentUnitStatusDesc,TotalMark,TotalGradeDesc,TotalResultDesc,CentreUnitLocalName,StartDate,DoneDate"
          mvHeadings = ",Code,Description,Candidate No,Attempt,Status,Mark,Grade,Result,Local Name,Exam Date,Result Date"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamBookingUnitId,ExamBookingId,ExamUnitId,CancellationReason,Source,ActivityGroup,ExamSessionId,CourseStartDate,ExamAssessmentLanguage,CanEditResults,StudyMode"
          mvDisplayTitle = "Items"
          mvDescription = "Contact Exam Detail Items"
          mvCode = "CXDI"
          vPrimaryList = True

        Case DataSelectionTypes.dstContactExamDetailList
          mvResultColumns = "ExamBookingUnitId,ExamBookingId,ExamUnitId,ExamUnitCode,ExamUnitDescription,ExamCandidateNumber,AttemptNumber,ExamStudentUnitStatus,ExamStudentUnitStatusDesc,OriginalMark,ModeratedMark,TotalMark,OriginalGrade,ModeratedGrade,TotalGrade,TotalGradeDesc,OriginalResult,ModeratedResult,TotalResult,TotalResultDesc,DoneDate,StartDate,StartTime,EndTime,CancellationReason,CancellationSource,CancellationReasonDesc,CancellationSourceDesc,CancelledOn,CancelledBy,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CanEditResults,ResultsReleaseDate,PreviousMark,PreviousGrade,PreviousResult,PreviousGradeDesc,PreviousResultDesc"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamCandidateNumber,AttemptNumber,ExamStudentUnitStatusDesc,TotalMark,TotalGradeDesc,TotalResultDesc,StartDate,DoneDate"
          mvHeadings = "Code,Description,Candidate No,Attempt,Status,Mark,Grade,Result,Exam Date,Result Date"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamBookingUnitId,ExamBookingId,ExamUnitId,CanEditResults"
          mvDisplayTitle = "Detail List"
          mvDescription = "Contact Exam Detail List"
          mvCode = "CXDL"
          vPrimaryList = True


        Case DataSelectionTypes.dstContactExamExemptions
          mvResultColumns = "ExamStudentExemptionId,ExamExemptionId,ExamExemptionCode,ExamExemptionDescription,ExamExemptionStatus,ExamExemptionStatusDesc,AllowExemptionEntry,ExamExemptionStatusType,StatusDate,Product,ProductDesc,Rate,RateDesc,BatchNumber,TransactionNumber,LineNumber,OrganisationNumber,Name,ExemptionModule,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamExemptionCode,ExamExemptionDescription,ExamExemptionStatus,ExamExemptionStatusDesc,StatusDate,ProductDesc,RateDesc,BatchNumber,TransactionNumber,LineNumber"
          mvHeadings = "Code,Description,Status,Description,Date,Product,Rate,Batch,Transaction,Line"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamStudentExemptionId,ExamExemptionId,AllowExemptionEntry,ExamExemptionStatusType"
          mvMaintenanceDesc = "Exemption"
          mvDescription = "Exam Exemptions"
          mvCode = "CXEX"
          vPrimaryList = True

        Case DataSelectionTypes.dstExamPersonnelFinder
          mvResultColumns = "ExamPersonnelId,ContactNumber,AddressNumber,Surname,Forenames,Initials,ValidFrom,Validto,ExamPersonnelType,ExamPersonnelTypeDesc"
          mvSelectColumns = "ExamPersonnelId,ContactNumber,Surname,Forenames,Initials,ValidFrom,Validto,ExamPersonnelTypeDesc"
          mvHeadings = "ID,Contact No,Surname,Forenames,Initials,ValidFrom,ValidTo,Type"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Exam Personnel Finder"
          mvCode = "EXPF"

        Case DataSelectionTypes.dstLoanInterestRates
          mvResultColumns = "LoanNumber,InterestRate,RateChanged"
          mvSelectColumns = "LoanNumber,InterestRate,RateChanged"
          mvHeadings = "Loan Number,Interest Rate,Change Date"
          mvWidths = "1,1200,1200"
          mvDescription = "Interest Rates"
          mvDisplayTitle = "Interest Rates"
          mvCode = "PLIR"

        Case DataSelectionTypes.dstWebExams
          mvResultColumns = "ExamUnitId,ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitCode,ExamUnitDescription,ExamSessionCode,ExamSessionDescription,ExamCentreCode,ExamCentreDescription,ExamImage,Subject,SubjectDesc,SkillLevel,SkillLevelDesc,StartDate,StartTime"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamSessionCode,ExamCentreCode,StartDate,StartTime"
          mvHeadings = "Exam Code,Description,Session,Centre,Start Date,Start Time"
          mvWidths = "1200,1200,1200,1200,1200,1200"
          mvDescription = "Web Exams"
          mvCode = "WEXN"
          mvRequiredItems = "ExamUnitId,ExamUnitCode,ExamSessionId,ExamCentreId"

        Case DataSelectionTypes.dstPaymentPlanPaymentDetails
          mvResultColumns = "PaymentPlanNumber,PaymentNumber,DetailNumber,Product,Rate,DetailBalance,PaymentAmount,ModifierActivity,ModifierActivityValue,ModifierActivityQuantity,ModifierActivityDate,ModifierPrice,ModifierPerItem,"
          mvResultColumns &= "UnitPrice,ProRated,NetAmount,VATAmount,GrossAmount,VATRate,VATPercentage,ProductDesc,RateDesc,VATExclusive,ModifierActivityDesc,ModifierActivityValueDesc"
          mvSelectColumns = "PaymentPlanNumber,PaymentNumber,DetailNumber,Product,Rate,DetailBalance,PaymentAmount,ModifierActivity,ModifierActivityValue,ModifierActivityQuantity,ModifierActivityDate,ModifierPrice,ModifierPerItem,"
          mvSelectColumns &= "UnitPrice,ProRated,NetAmount,VATAmount,GrossAmount,VATRate,VATPercentage,VATExclusive"
          mvHeadings = "Plan No,Payment No,Detail No,Product,Rate,Detail Balance,Pay Amount,Activity,Activity Value,Activity Quantity,Activity Date,Modifier Price,Per Item,Unit Price,Pro-Rated,Net Amount,VAT Amount,Gross Amount,VAT Rate,VAT Percentage,VAT Exclusive"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvDescription = "Payment Plan History Details"
          mvDisplayTitle = "Pay Plan Details"
          mvCode = "PPHD"
          mvRequiredItems = ""

        Case DataSelectionTypes.dstWebMemberOrganisations
          mvResultColumns = "ContactNumber,AddressNumber,Name,Abbreviation,MemberNumber,HouseName,Address,Town,County,Postcode,Country,CountryDesc,AddressLine"
          mvSelectColumns = "Name,Abbreviation,AddressLine"
          mvHeadings = "Organisation,Abbreviation,Address"
          mvWidths = "1200,1200,1200"
          mvDescription = "Web Member Organisations"
          mvCode = "WMOG"

        Case DataSelectionTypes.dstWebExamBookings
          mvResultColumns = "ExamBookingUnitId,ExamUnitId,ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitCode,ExamUnitDescription,ExamSessionCode,ExamSessionDescription,ExamCentreCode,ExamCentreDescription,StartDate,StartTime,EndTime,ExamStudentUnitStatus,ExamStudentUnitStatusDesc"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamStudentUnitStatusDesc,ExamCentreDescription,StartDate"
          mvHeadings = "Exam Code,Description,Status,Centre,Date"
          mvWidths = "1000,1000,1000,1000,1000"
          mvDescription = "Web Exam Bookings"
          mvCode = "WEXB"
          mvRequiredItems = "ExamBookingUnitId,ExamUnitId,ExamScheduleId,ExamSessionId,ExamCentreId"

        Case DataSelectionTypes.dstWebExamHistory
          mvResultColumns = "ExamStudentHeaderId,ExamStudentHeaderUnitId,ExamUnitId,ExamSessionId,ExamSessionCode,ExamSessionDescription,ExamSessionMonth,ExamSessionYear,StartDate,ExamUnitCode,ExamUnitDescription,CurrentMark,CurrentGrade,CurrentGradeDescription,CurrentResult,CurrentResultDesc"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamSessionDescription,StartDate,CurrentResultDesc"
          mvHeadings = "Exam Code,Description,Session,Date,Result"
          mvWidths = "1000,1000,1000,1000,1000"
          mvDescription = "Web Exam History"
          mvCode = "WEXH"
          mvRequiredItems = "ExamStudentHeaderId,ExamStudentHeaderUnitId,ExamUnitId,ExamSessionId"

        Case DataSelectionTypes.dstContactAmendments
          mvResultColumns = "OperationDate,Logname,Operation,TableName,JournalNumber"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Date,User,Operation,Table,Journal Number"
          mvWidths = "1200,1200,1200,1200,1200"
          mvDescription = "Amendment History"
          mvDisplayTitle = DataSelectionText.String18684   'Amendments
          mvCode = "COAH"
          mvRequiredItems = "OperationDate,Operation,TableName"

        Case DataSelectionTypes.dstContactAmendmentDetails
          mvResultColumns = "Item,OldValues,NewValues"
          mvSelectColumns = mvResultColumns
          mvHeadings = "Item,Old Values,New Values"
          mvWidths = "1200,1200,1200"
          mvDescription = "Amendment History Details"
          mvDisplayTitle = DataSelectionText.String18684   'Amendments
          mvCode = "CAHD"

        Case DataSelectionTypes.dstMeetingActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText"
          mvWidths = "800,2000,1500,1500,2000"
          mvHeadings = "Number,Description,Priority,Status,Action Text"
          mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
          mvDescription = "Meeting Actions"
          mvCode = "MEAC"
          mvRequiredItems = "ActionStatus,MasterAction"

        Case DataSelectionTypes.dstContactExamCertificates
          mvResultColumns = "ContactExamCertId,ContactNumber,CertNumberPrefix,CertNumber,CertNumberSuffix,ExamUnitCode,ExamUnitDescription,IsCertificateRecalled"
          mvSelectColumns = "ContactNumber,CertNumberPrefix,CertNumber,CertNumberSuffix,ExamUnitCode,ExamUnitDescription,IsCertificateRecalled"
          mvHeadings = "Contact Number,Certificate Prefix,Certificate Number, Certificate Suffix,Unit Code,Unit Description,Recalled?"
          mvWidths = "1000,1000,2000,1000,1000,2000"
          mvDescription = "Contact Exam Certificates"
          mvCode = "CEXC"
          mvRequiredItems = "ContactExamCertId,ContactNumber,IsCertificateRecalled"

        Case DataSelectionTypes.dstContactExamCertificateItems
          mvResultColumns = "ContactExamCertItemId,ContactNumber,ContactExamCertId,CertAttribute,CertAttributeValue,AmendeOn,AmendedBy"
          mvSelectColumns = "ContactNumber,ContactExamCertId,CertAttribute,CertAttributeValue,AmendeOn,AmendedBy"
          mvHeadings = "Contact Number,Certificate Id,Attribute,Attribute Value,Amended On,Amended By"
          mvWidths = "2000,2000,2000,1000,1000"
          mvDescription = "Contact Exam Certificates"
          mvCode = "CEXI"
          mvRequiredItems = "ContactNumber,ContactExamCertId,ContactExamCertItemId"

        Case DataSelectionTypes.dstContactExamCertificateReprints
          mvResultColumns = "ContactExamCertReprintId,ContactExamCertId,ExamCertReprintType,AmendeOn,AmendedBy"
          mvSelectColumns = "ContactExamCertReprintId,ContactExamCertId,ExamCertReprintType,AmendeOn,AmendedBy"
          mvHeadings = "ContactExamCertReprintId,ContactExamCertId,ExamCertReprintType,AmendeOn,AmendedBy"
          mvWidths = "2000,2000,2000,1000,1000"
          mvDescription = "Contact Exam Certificate Reprints"
          mvCode = "CEXR"
          mvRequiredItems = "ContactExamCertReprintId"

        Case DataSelectionTypes.dstFundraisingRequests
          mvResultColumns = "RequestNumber,ContactNumber,RequestDate,RequestDesc,RequestStage,Status,FundraisingRequestType,TargetAmount,PledgedAmount,PledgedDate,ReceivedAmount,ReceivedDate,ExpectedAmount,GikExpectedAmount,GikPledgedAmount,GikPledgedDate,TotalGikReceivedAmount,LatestGikReceivedDate,NumberOfPayments,RequestEndDate,FundraisingBusinessType"
          mvSelectColumns = "RequestNumber,ContactNumber,RequestDate,RequestDesc,RequestStage,Status,FundraisingRequestType,TargetAmount,PledgedAmount,PledgedDate,ReceivedAmount,ReceivedDate,ExpectedAmount,GikExpectedAmount,GikPledgedAmount,GikPledgedDate,TotalGikReceivedAmount,LatestGikReceivedDate,NumberOfPayments,RequestEndDate,FundraisingBusinessType"
          mvHeadings = "Request Number,Contact Number,Request Date,Request Desc,Request Stage,Status,Request Type,Target Amount,Pledged Amount,Pledged Date,Received Amount,Received Date,Expected Amount,Gik Expected Amount,Gik Pledged Amount,Gik Pledged Date,Total Gik Received Amount,Latest Gik Received Date,Number Of Payments,Request End Date,Fundraising Business Type"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Fundraising Request"
          mvMaintenanceDesc = "Fundraising Request"
          mvDescription = DataSelectionText.FundraisingRequest
          mvCode = "FFR"
          mvRequiredItems = "RequestNumber"
        Case DataSelectionTypes.dstExamScheduleFinder

          mvResultColumns = "ExamScheduleId,ExamSessionId,ExamCentreId,ExamSessionCode,ExamSessionDescription,ExamSessionYear,ExamSessionMonth,SessionSequenceNumber,ExamCentreCode,ExamCentreDescription,ExamUnitId,ExamUnitCode,ExamUnitDescription,ExamScheduleStartDate,ExamScheduleStartTime,ExamScheduleEndTime"
          mvSelectColumns = "ExamSessionCode,ExamSessionDescription,ExamSessionYear,ExamSessionMonth,SessionSequenceNumber,ExamCentreCode,ExamCentreDescription,ExamUnitCode,ExamUnitDescription,ExamScheduleStartDate,ExamScheduleStartTime,ExamScheduleEndTime"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600,1400,1400,1400"
          mvHeadings = "Session Code,Session,Session Year,Session Month,Session Sequence,Centre Code,Centre,Exam Unit Code,Exam Unit,Schedule Start Date,Schedule Start Time,Schedule End Time"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Workstream Exam Schedule"
          mvDisplayTitle = "Workstream Actions"
          mvMaintenanceDesc = "Action"
          mvRequiredItems = "ExamSessionId,ExamCentreId,ExamScheduleId"
          mvCode = "XSCH"
        Case DataSelectionTypes.dstContactTokens
          mvResultColumns = "ContactNumber,CardTokenNumber,CardTokenDesc,CardNumber"
          mvSelectColumns = "ContactNumber,CardTokenNumber,CardTokenDesc,CardNumber"
          mvHeadings = "Contact Number,Card Token Number,Card Token Desc,Card Number"
          mvWidths = "1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDisplayTitle = "Contact Tokens"
          mvMaintenanceDesc = "Contact Tokens"
          mvDescription = DataSelectionText.FundraisingRequest
          mvCode = "CNTT"
          mvRequiredItems = "ContactNumber"

        Case DataSelectionTypes.dstEventSessionCPD
          mvResultColumns = "EventSessionCpdNumber,EventNumber,SessionNumber,CategoryType,CategoryTypeDesc,Category,CategoryDesc,Year,Points,Points2,ItemType,ItemTypeDesc,Outcome,ApprovalStatus,ApprovalStatusDesc,DateApproved,AwardingBody,WebPublish,Notes"
          mvSelectColumns = "CategoryTypeDesc,CategoryDesc,Points,Points2,ItemTypeDesc,Outcome,ApprovalStatusDesc,DateApproved,AwardingBody"
          mvHeadings = "Category Type,Category,Points,Points 2,Item Type,Outcome,Approval Status,Approval Date,Awarding Body"
          mvWidths = "2000,2000,2000,2000,2000,2000,2000,2000,2000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDescription = DataSelectionText.DstEventSessionCPD
          mvCode = "EVSC"
          mvRequiredItems = "CategoryType,Category"

        Case DataSelectionTypes.dstContactCPDCycleDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,"
          mvResultColumns &= "StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc"
          mvResultColumns &= ",Subject,CallDuration,TotalDuration,SelectionSet,CPDCycleNumber,CPDPeriodNumber,CPDPeriodDesc"
          mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc"
          mvHeadings = "Document Number, Dated, In/Out, Subject, Reference, Document Type, Topic, Sub Topic"
          mvWidths = "2000,2000,2000,2000,2000,2000,2000,2000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDescription = DataSelectionText.DstContactCPDCycleDocumentsDescription
          mvMaintenanceDesc = DataSelectionText.DstContactCPDCycleDocumentsMaintDesc
          mvDisplayTitle = DataSelectionText.DstContactCPDCycleDocumentsDisplayTitle
          mvCode = "CPED"
          mvRequiredItems = "Access"

        Case DataSelectionTypes.dstContactCPDPointDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,"
          mvResultColumns &= "StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc"
          mvResultColumns &= ",Subject,CallDuration,TotalDuration,SelectionSet,CPDPointNumber"
          mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc"
          mvHeadings = "Document Number, Dated, In/Out, Subject, Reference, Document Type, Topic, Sub Topic"
          mvWidths = "2000,2000,2000,2000,2000,2000,2000,2000"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient Or DataSelectionUsages.dsuWEBServices
          mvDescription = DataSelectionText.DstContactCPDPointDocumentsDescription
          mvMaintenanceDesc = DataSelectionText.DstContactCPDPointDocumentsMaintDesc
          mvDisplayTitle = DataSelectionText.DstContactCPDPointDocumentsDisplayTitle
          mvCode = "CPOD"
          mvRequiredItems = "Access"

        Case DataSelectionTypes.dstFindCPDCyclePeriods
          mvResultColumns = "ContactCPDCycleNumber,CPDCycleType,CPDCycleTypeDesc,StartMonth,EndMonth,CycleStartDate,CycleEndDate,ContactCPDPeriodNumber,ContactCPDPeriodNumberDesc,PeriodStartDate,PeriodEndDate,CPDCycleStatus,CPDCycleStatusDesc,CPDType,CPDTypeDesc,ContactNumber,ContactName"
          mvSelectColumns = "ContactCPDPeriodNumber,ContactCPDCycleNumber,ContactCPDPeriodNumberDesc,PeriodStartDate,PeriodEndDate,CPDCycleTypeDesc,CPDTypeDesc,CPDCycleStatusDesc,ContactNumber,ContactName"
          mvHeadings = "Period Number,Cycle Number,Period,Period Start,Period End,Cycle Type,Type,Cycle Status,Contact Number,Contact Name"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.DstFindCPDCyclePeriodsDescription
          mvRequiredItems = "ContactCPDPeriodNumberDesc"
          mvCode = "FCPE"

        Case DataSelectionTypes.dstFindCPDPoints
          mvResultColumns = "ContactCPDPointNumber,ContactCPDPeriodNumber,CategoryType,Category,Points,Points2,PointsDate,Activity,ActivityValue,ItemType,"
          mvResultColumns &= "CategoryTypeDesc,CategoryDesc,ActivityDesc,ActivityValueDesc,ItemTypeDesc,CyclePeriodNumberDesc,CycleType,CycleTypeDesc,ContactNumber,ContactName"
          mvSelectColumns = "ContactCPDPointNumber,PointsDate,CategoryTypeDesc,CategoryDesc,Points,Points2,ItemTypeDesc,CycleTypeDesc,CyclePeriodNumberDesc,ContactNumber,ContactName"
          mvHeadings = "Point Number,Points Date,Category Type,Category,Points,Points 2,Item Desc,Cycle Type,Cycle Period,Contact Number,Contact Name"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          Dim vCpdPointsItemName As String = pEnv.GetConfig("cpd_points_item_name")
          If Not String.IsNullOrEmpty(vCpdPointsItemName) Then
            mvHeadings = Replace(mvHeadings, "Points", pEnv.GetConfig("cpd_points_item_name"))
          End If
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = DataSelectionText.DstFindCPDPointsDescription
          mvRequiredItems = "CategoryTypeDesc,CategoryDesc"
          mvCode = "FCPP"

        Case DataSelectionTypes.dstContactViewOrganisations
          mvResultColumns = "ContactPositionNumber,ContactNumber,AddressNumber,Position," & ContactNameResults() & ",ValidFrom,ValidTo,Mail,Current,Location,PositionStatus,PositionFunction,PositionSeniority,SinglePosition,OrganisationGroup,ContactGroup,AddressLine,PositionStatusDesc,PositionFunctionDesc,PositionSeniorityDesc"
          mvSelectColumns = "Position,ContactName,PositionStatusDesc,AddressLine,Current,Mail,DetailItems,ValidFrom,Location,NewColumn,ValidTo"
          mvHeadings = "Position,Name,Status,Address,Current,Mail,,Valid From,Location,,Valid To"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ContactName,Location,Position,ValidFrom,ValidTo,SinglePosition,OrganisationGroup,ContactGroup,Current"
          mvDescription = "Positions"
          mvMaintenanceDesc = "Position"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "CVOP"
          vPrimaryList = True

        Case DataSelectionTypes.dstSalesLedgerAnalysis
          mvResultColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionSLAmount,LineAmount,InvoicePaymentAmount,UnallocatedAmount"
          mvSelectColumns = "BatchNumber,TransactionNumber,LineNumber,TransactionSLAmount,LineAmount,InvoicePaymentAmount,UnallocatedAmount"
          mvHeadings = "Batch,Transaction,Line,Transaction Amount,Line Amount,Invoice Amount,Unallocated Amount"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "BatchNumber,TransactionNumber,LineNumber,TransactionSLAmount,LineAmount,InvoicePaymentAmount,UnallocatedAmount"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient

        Case DataSelectionTypes.dstTraderAlerts
          mvResultColumns = "TraderApplicationNumber,ContactAlertLinkNumber,ContactAlert,ContactAlertDesc,ContactAlertSql,AlertMessageDesc,ShowAsDialog,AlertMessageType,SequenceNumber"
          mvSelectColumns = "ContactAlertDesc,AlertMessageDesc,ShowAsDialog,AlertMessageType,SequenceNumber"
          mvHeadings = "Description,Message,Show Dialog,Dialog Type,Sequence"
          mvWidths = "1200,1200,1200,1200,1200"
          mvRequiredItems = "ContactAlertLinkNumber"
          mvDescription = "Trader Alerts"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "TRDA"

        Case DataSelectionTypes.dstContactFinanceAlerts
          mvResultColumns = "ContactNumber,ContactAlert,AlertMessageDesc,RgbAlertMessage,ShowAsDialog,AlertMessageType"
          mvSelectColumns = "AlertMessageDesc"
          mvRequiredItems = "RgbAlertMessage,ShowAsDialog,AlertMessageType"

      End Select

      If vPrimaryList Then
        If mvType = DataSelectionTypes.dstSelectionPages Or mvType = DataSelectionTypes.dstEventSelectionPages Then
          '
        ElseIf ((pUsage = DataSelectionUsages.dsuWEBServices) Or (pUsage = DataSelectionUsages.dsuSmartClient) Or (mvType = DataSelectionTypes.dstContactMembershipDetails) Or (mvType = DataSelectionTypes.dstContactMailings) Or (mvType = DataSelectionTypes.dstContactCategories)) And pListType = DataSelectionListType.dsltEditing Then
          mvResultColumns = mvResultColumns & ",DetailItems,NewColumn,NewColumn2,NewColumn3,Spacer"
        End If
      End If

      Select Case pListType
        Case DataSelectionListType.dsltDefault
          'Do nothing
        Case DataSelectionListType.dsltUser
          ReadUserDisplayListItems(mvEnv.User.Department, mvEnv.User.Logname, pGroup, pUsage)
        Case DataSelectionListType.dsltEditing
          'GetDefaultDisplayListItems()
      End Select

    End Sub

    Private Sub CheckCustomForms()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("client", mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("custom_form", mvEnv.FirstCustomFormNumber, CDBField.FieldWhereOperators.fwoBetweenFrom)
      vWhereFields.Add("custom_form#2", mvEnv.LastCustomFormNumber, CDBField.FieldWhereOperators.fwoBetweenTo)
      vWhereFields.Add("custom_form#3", 1000, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("form_usage_code", "CO") 'Only get Contact and Organisation custom formns
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "custom_form,form_caption", "custom_forms", vWhereFields, "custom_form").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",CustomForm" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Sub CheckActivityDataSheets()
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "activity_group,activity_group_desc", "activity_groups", New CDBField("usage_code", "B"), "activity_group").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",ActivityGroup" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Sub CheckTopicDataSheets()
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "topic_group, topic_group_desc", "topic_groups", New CDBField("usage_code", "E"), "topic_group").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",TopicGroup" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Sub CheckRelationshipDataSheets()
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "relationship_group,relationship_group_desc", "relationship_groups", New CDBField("usage_code", "B"), "relationship_group").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",RelationshipGroup" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub
    Private Sub CheckOrganisationGroups()
      Dim vWhereFields = New CDBFields
      vWhereFields.Add("og.view_in_contact_card", "Y")
      vWhereFields.Add("og.client", mvEnv.ClientCode)
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "og.organisation_group, og.name", "organisation_groups og", vWhereFields, "og.organisation_group").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",OrganisationGroup" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property AvailableColumns() As String
      Get
        Return mvResultColumns
      End Get
    End Property

    Public ReadOnly Property DisplayColumns() As String
      Get
        Return mvSelectColumns
      End Get
    End Property

    Public ReadOnly Property DisplayHeadings() As String
      Get
        Return mvHeadings
      End Get
    End Property

    Public ReadOnly Property DisplayTitle() As String
      Get
        Return mvDisplayTitle
      End Get
    End Property

    Public ReadOnly Property MaintenanceDesc() As String
      Get
        Return mvMaintenanceDesc
      End Get
    End Property

    Public ReadOnly Property DisplayWidths() As String
      Get
        Return mvWidths
      End Get
    End Property

    Public ReadOnly Property AvailableUsage() As DataSelectionUsages
      Get
        Return mvAvailableUsages
      End Get
    End Property

    Public ReadOnly Property DataSelectionType() As DataSelectionTypes
      Get
        Return mvType
      End Get
    End Property

    Public ReadOnly Property Description() As String
      Get
        Return mvDescription
      End Get
    End Property
    Public ReadOnly Property HeaderLines As Integer
      Get
        Return mvHeaderLines
      End Get
    End Property
    Public ReadOnly Property Department As String
      Get
        Return mvDepartment
      End Get
    End Property
    Public ReadOnly Property Logname As String
      Get
        Return mvLogname
      End Get
    End Property

    Private Sub InitCustomFormData()
      Dim vIndex As Integer
      Dim vItems As String = ""
      Dim vHeadings As String = ""
      Dim vWidths As String = ""
      Dim vPos As Integer

      Dim vFields() As String
      If mvParameters IsNot Nothing Then
        mvCustomForm = New CustomForm(mvEnv)
        mvCustomForm.Init(mvParameters("CustomForm").LongValue)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCustomFormWebPage) AndAlso mvCustomForm.CustomFormUrl.Length > 0 Then
          vItems = "custom_form_url,show_browser_toolbar"
          vHeadings = "CustomFormUrl,ShowBrowserToolbar"
          vWidths = "1200,1200"
        ElseIf mvParameters.HasValue("Default") Then
          vItems = "parameter_name,default_value"
          vHeadings = "Parameter Name,Default Value"
          vWidths = "1200,1200"
        ElseIf mvParameters.HasValue("Detail") Then
          Dim vGetVASDetail As Boolean
          Dim vSequenceNo As Integer
          If mvParameters("Detail").Value.StartsWith("vas") Then
            vGetVASDetail = True
            vSequenceNo = IntegerValue(mvParameters("Detail").Value.Substring(3))
          End If
          Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "control_type,attribute_name,control_caption,control_width,sequence_number,desc_sql", "custom_form_controls", New CDBField("custom_form", CDBField.FieldTypes.cftInteger, mvCustomForm.CustomFormNumber.ToString), "sequence_number").GetRecordSet
          Do While vRecordSet.Fetch
            Dim vParams As New CDBParameters
            Select Case vRecordSet.Fields(1).Value.Substring(0, 3)
              Case "vas"
                If vSequenceNo = vRecordSet.Fields(5).LongValue Then
                  vHeadings = vRecordSet.Fields(2).Value.Replace("|", ",")
                  vParams.InitFromSQLAttributes(vRecordSet.Fields(6).Value)
                  vItems = vParams.ItemList
                  vWidths = vRecordSet.Fields(3).Value.Replace("|", ",")
                End If
              Case "grd"
                If vRecordSet.Fields(2).Value = mvParameters("Detail").Value Then
                  vParams.InitFromSQLAttributes(mvCustomForm.GridSelectSql)
                  vItems = vParams.ItemList
                  vFields = vRecordSet.Fields(6).Value.Split(","c)
                  vHeadings = ""
                  vWidths = ""
                  vPos = (vFields.GetUpperBound(0) + 1) \ 2
                  For vIndex = 0 To vPos - 1
                    If vIndex > 0 Then
                      vHeadings = vHeadings & ","
                      vWidths = vWidths & ","
                    End If
                    vHeadings = vHeadings & vFields(vIndex)
                    vWidths = vWidths & vFields(vIndex + vPos)
                  Next
                End If
              Case Else
                If Not vGetVASDetail Then
                  If vItems.Length > 0 Then
                    vItems = vItems & ","
                    vHeadings = vHeadings & ","
                    vWidths = vWidths & ","
                  End If
                  vItems = vItems & vRecordSet.Fields(2).Value
                  vHeadings = vHeadings & vRecordSet.Fields(3).Value
                  vWidths = vWidths & vRecordSet.Fields(4).Value
                End If
            End Select
          Loop
          vRecordSet.CloseRecordSet()
        Else
          vItems = mvCustomForm.GridAttributeNames
          vHeadings = mvCustomForm.GridHeadings
          vWidths = mvCustomForm.GridWidths
        End If
      End If
      mvCustomFieldNames = vItems
      vFields = vItems.Split(","c)
      For vIndex = 0 To vFields.GetUpperBound(0)
        Dim vFieldName As String = vFields(vIndex)
        vPos = (vFieldName.IndexOf(".", 0) + 1)
        If vPos >= 0 Then vFieldName = vFieldName.Substring(vPos)
        vFields(vIndex) = StrConv(vFieldName.Replace("_", " "), vbProperCase).Replace(" ", "")
      Next
      mvResultColumns = Join(vFields, ",")
      mvSelectColumns = mvResultColumns
      mvHeadings = vHeadings
      mvWidths = vWidths
    End Sub

    Private Sub InitDashboardData()
      mvResultColumns = ""
      mvSelectColumns = ""
      mvHeadings = ""
      mvWidths = ""
      mvCode = "DSH"
    End Sub

    Public Sub ReadUserDisplayListItems(ByVal pDepartment As String, ByVal pUser As String, ByVal pGroup As String, ByVal pUsage As DataSelectionUsages)
      Dim vSelectionParameters As Boolean
      Dim vContinue As Boolean = True
      'Check if parameters have been passed to setup results
      If Not mvParameters Is Nothing Then
        If mvParameters.HasValue("SelectedColumns") AndAlso mvParameters.HasValue("SelectedHeadings") Then vSelectionParameters = True
        vContinue = Not (mvParameters.ParameterExists("WPD").Bool AndAlso Not mvParameters.ParameterExists("WebPageItemNumber").LongValue > 0)
      End If

      If vContinue Then
        Dim vResultColumns As String() = mvResultColumns.Split(","c)
        If vSelectionParameters Then
          mvSelectColumns = mvParameters("SelectedColumns").Value
          mvHeadings = mvParameters("SelectedHeadings").Value
          Dim vSelectParams As New StringList(mvSelectColumns, StringSplitOptions.None)
          Dim vWidth As New StringBuilder
          Dim vAddSeparator As Boolean
          For Each vParam As String In vSelectParams
            If vAddSeparator Then vWidth.Append(",")
            vWidth.Append("300")
            vAddSeparator = True
          Next
          mvWidths = vWidth.ToString
        Else
          If Len(mvCode) > 0 Then
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("display_list", mvCode)
            If pUsage = DataSelectionUsages.dsuCare Then
              vWhereFields.Add("access_method", "A")
            ElseIf pUsage = DataSelectionUsages.dsuWEBServices Then
              vWhereFields.Add("access_method", "W")
              If mvParameters IsNot Nothing AndAlso mvParameters.ParameterExists("WebPageItemNumber").IntegerValue > 0 Then
                vWhereFields.Add("web_page_item_number", mvParameters("WebPageItemNumber").IntegerValue, CDBField.FieldWhereOperators.fwoEqual)
              End If
            ElseIf pUsage = DataSelectionUsages.dsuSmartClient Then
              vWhereFields.Add("access_method", "S")
            Else
              Debug.Assert(False)
            End If
            If mvParameters IsNot Nothing AndAlso (Not mvParameters.ParameterExists("WebPageItemNumber").IntegerValue > 0) Then
              Dim vDepartment As String = pDepartment
              Dim vUser As String = pUser
              If mvParameters.ParameterExists("CustomiseData").Bool Then
                vDepartment = mvEnv.User.Department
                vUser = mvEnv.User.Logname
              ElseIf vUser.Length > 0 AndAlso vDepartment.Length = 0 Then
                'Where User parameter supplied but Department not supplied (Department Any) include for User Deparment
                Dim vDisplayUser As New CDBUser(mvEnv)
                vDisplayUser.Init(vUser)
                vDepartment = vDisplayUser.Department
              End If
              vWhereFields.AddClientDeptLogname(mvEnv.ClientCode, vDepartment, vUser)
            End If
            If Len(pGroup) > 0 Then
              vWhereFields.Add("contact_group", pGroup, CDBField.FieldWhereOperators.fwoOpenBracket)
              vWhereFields.Add("contact_group#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
            Else
              vWhereFields.Add("contact_group")
            End If
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "display_items,display_headings,display_sizes,heading_lines,department,logname,display_title,contact_group,maintenance_desc", "display_list_items", vWhereFields)
            If mvEnv.Connection.NullsSortAtEnd Then
              vSQLStatement.OrderBy = "contact_group,logname,department"
            Else
              vSQLStatement.OrderBy = "contact_group DESC,logname DESC,department DESC"
            End If
            Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
            If vRS.Fetch() Then
              mvDefaults = False
              mvSelectColumns = vRS.Fields("display_items").Value
              mvHeadings = vRS.Fields("display_headings").Value
              mvWidths = vRS.Fields("display_sizes").Value
              mvHeaderLines = vRS.Fields("heading_lines").IntegerValue
              mvDepartment = vRS.Fields("department").Value
              mvLogname = vRS.Fields("logname").Value
              mvContactGroup = vRS.Fields("contact_group").Value
              mvDisplayTitle = vRS.Fields("display_title").Value
              If vRS.Fields("maintenance_desc").Value.Length > 0 Then mvMaintenanceDesc = vRS.Fields("maintenance_desc").Value
              Select Case mvType
                Case DataSelectionTypes.dstSelectionPages
                  ReplaceItem("Pledges", "PreTaxPledges")
                Case DataSelectionTypes.dstContactHeaderInfo, DataSelectionTypes.dstContactInformation
                  ReplaceItem("Preferred", "PreferredCommunication")
                Case DataSelectionTypes.dstCPDDetails
                  If mvParameters IsNot Nothing AndAlso mvParameters.ParameterExists("CPDType").Value = "O" Then
                    'Special case for CPDDetails as we use one data selection for both Points and Objectives
                    'Display List is provided for Points only hence on loading items for Objectives we need to replace the standard words
                    mvSelectColumns = mvSelectColumns.Replace("Points", "Objective").Replace("EvidenceSeen", "SupervisorAccepted")
                    mvHeadings = mvHeadings.Replace("Points", "Objective").Replace("Evidence Seen", "Supervisor Accepted")
                  End If
                Case DataSelectionTypes.dstContactGiftAidDeclarations 'BR19026 BR19437
                  ReplaceItem("DeclarationMethodCode", "DeclarationMethodDesc")
              End Select
            Else
              If mvType = DataSelectionTypes.dstContactInformation Then
                Dim vEntityGroup As EntityGroup = mvEnv.EntityGroups(pGroup)
                If vEntityGroup.EntityGroupType = EntityGroup.EntityGroupTypes.egtOrganisation Then
                  mvSelectColumns = "DetailItems,OrganisationName,Abbreviation,SourceDesc,StatusReason,OwnershipGroupDesc,DefaultContactName,AmendedBy,Notes,NewColumn,Spacer1,Spacer2,SourceDate,StatusDate,PrincipalDepartmentDesc,OwnershipAccessLevelDesc,AmendedOn,Spacer3"
                  mvHeadings = DataSelectionText.String18156    ',Name,Abbreviation,Source,Status Reason,Ownership Group,Default Contact,Amended By,Notes,,,,Source Date,Status Date,Principal Department,Access Level,Amended On,
                  mvWidths = "1200,1000,1000,1000,1000,1000,1000,1000,1000,1200,1200,1200,1000,1000,1000,1000,1000,1200"
                Else
                  mvSelectColumns = "DetailItems,Title,Forenames,Salutation,PreferredForename,DateOfBirth,SourceDesc,StatusReason,OwnershipGroupDesc,OwnershipAccessLevelDesc,AmendedBy,Notes,NewColumn,Initials,Surname,Honorifics,Sex,DOBEstimated,SourceDate,StatusDate,PrincipalDepartmentDesc,Spacer1,AmendedOn,Spacer2"
                  mvHeadings = DataSelectionText.String18157    ',Title,Forenames,Salutation,Preferred,Date of Birth,Source,Status Reason,Ownership Group,Access Level,Amended By,Notes,,Initials,Surname,Honorifics,Sex,Estimated,Source Date,Status Date,Principal Department,,Amended On,
                  mvWidths = "1200,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1200,1000,1000,1000,1000,1000,1000,1000,1000,1200,1000,1200"
                End If
                mvDisplayTitle = ""
              End If
            End If
            vRS.CloseRecordSet()
          End If
        End If

        'Now check for required items
        Dim vIndex As Integer
        Select Case mvType
          Case DataSelectionTypes.dstContactHeaderInfo, DataSelectionTypes.dstEventHeaderInfo
            'No additional fields
          Case DataSelectionTypes.dstQueryByExampleContacts, DataSelectionTypes.dstQueryByExampleOrganisations, DataSelectionTypes.dstQueryByExampleEvents
            'No additional fields
          Case Else
            'Add to the list of required items the ones that end in Number
            For vIndex = 0 To vResultColumns.GetUpperBound(0)
              If vResultColumns(vIndex).EndsWith("Number") OrElse vResultColumns(vIndex).EndsWith("Id") Then
                If mvRequiredItems.Length > 0 Then mvRequiredItems = mvRequiredItems & ","
                mvRequiredItems = mvRequiredItems & vResultColumns(vIndex)
              End If
            Next
        End Select

        'We now have a list of required items so need to check if they are in the list anywhere
        'If  it is called from display list maintenance than just show selected items do not add required items
        If mvRequiredItems.Length > 0 AndAlso Not mvDataSelectionListType = DataSelectionListType.dsltEditing Then
          Dim vDetailIndex As Integer = -1

          If mvSelectColumns.Length > 0 Then
            Dim vSelectParams As New StringList(mvSelectColumns, StringSplitOptions.None)
            Dim vHeadingParams As New StringList(mvHeadings, StringSplitOptions.None)
            Dim vWidthParams As New StringList(mvWidths, StringSplitOptions.None)
            For vIndex = 0 To vSelectParams.Count - 1
              If vSelectParams(vIndex) = "DetailItems" Then
                vDetailIndex = vIndex
                Exit For
              End If
            Next
            Dim vRequired As String() = mvRequiredItems.Split(","c)
            For Each vItem As String In vRequired
              If Not vSelectParams.Contains(vItem) Then
                If vDetailIndex >= 0 Then
                  vSelectParams.Insert(vDetailIndex, vItem)
                  vHeadingParams.Insert(vDetailIndex, vItem)
                  vWidthParams.Insert(vDetailIndex, "1")
                Else
                  vSelectParams.Add(vItem)
                  vHeadingParams.Add(vItem)
                  vWidthParams.Add("1")
                End If
              End If
            Next
            mvSelectColumns = vSelectParams.ItemList
            mvHeadings = vHeadingParams.ItemList
            mvWidths = vWidthParams.ItemList
          End If
        End If

        If mvType = DataSelectionTypes.dstContactCommsNumbers AndAlso mvContact IsNot Nothing AndAlso mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then RemoveItem("IsOrganisation")
        If (mvType = DataSelectionTypes.dstSelectionPages Or mvType = DataSelectionTypes.dstEventSelectionPages) Then CheckDashboardItem()

        'Ensure that any spacers in the selection are in the collection
        If Not mvDisplayListItems Is Nothing Then
          Dim vSelectColumns As String() = mvSelectColumns.Split(","c)
          Dim vHeadings As String() = mvHeadings.Split(","c)
          Dim vWidths As String() = mvWidths.Split(","c)
          For vIndex = 0 To vSelectColumns.GetUpperBound(0)
            Dim vDLI As DisplayListItem
            If vSelectColumns(vIndex).StartsWith("Spacer") Or vSelectColumns(vIndex).StartsWith("SelectionHeading") Then
              If Not mvDisplayListItems.ContainsKey(vSelectColumns(vIndex)) Then
                vDLI = New DisplayListItem(mvEnv, vSelectColumns(vIndex), "", 1200, False)
                mvDisplayListItems.Add(vDLI.Name, vDLI)
              End If
            End If
            Dim vItemName As String = vSelectColumns(vIndex)
            Dim vReadOnly As Boolean = False
            If vItemName.EndsWith("_RO") Then
              vReadOnly = True
              vItemName = vItemName.Substring(0, vItemName.Length - 3)
            End If
            vDLI = mvDisplayListItems(vItemName)
            vDLI.UserHeading = vHeadings(vIndex)
            vDLI.UserWidth = IntegerValue(vWidths(vIndex))
            vDLI.IsReadOnly = vReadOnly
          Next
          If mvType = DataSelectionTypes.dstSelectionPages Then mvSelectColumns = mvSelectColumns.Replace("_RO", "")
        End If
      End If
    End Sub

    Private Sub AddWhereFieldFromParameter(ByVal pWhereFields As CDBFields, ByVal pParameterName As String, ByVal pFieldName As String)
      If mvParameters.Exists(pParameterName) Then pWhereFields.Add(pFieldName, mvParameters(pParameterName).Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
    End Sub

    Private Sub AddWhereFieldFromParameter(ByVal pWhereFields As CDBFields, ByVal pParameterName As String, ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes)
      If mvParameters.Exists(pParameterName) Then
        If pFieldType = CDBField.FieldTypes.cftMemo Then
          pWhereFields.Add(pFieldName, pFieldType, mvParameters(pParameterName).Value, CDBField.FieldWhereOperators.fwoLike)
        Else
          pWhereFields.Add(pFieldName, pFieldType, mvParameters(pParameterName).Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        End If
      End If
    End Sub

    Private Sub AddWhereFieldFromIntegerParameter(ByVal pWhereFields As CDBFields, ByVal pParameterName As String, ByVal pFieldName As String)
      If mvParameters.Exists(pParameterName) Then pWhereFields.Add(pFieldName, mvParameters(pParameterName).LongValue)
    End Sub

    Private Sub AddWhereFieldFromDateParameter(ByVal pWhereFields As CDBFields, ByVal pParameterName As String, ByVal pFieldName As String)
      If mvParameters.Exists(pParameterName) Then pWhereFields.Add(pFieldName, CDBField.FieldTypes.cftDate, mvParameters(pParameterName).Value)
    End Sub

    Private Function RemoveBlankItems(ByVal pItems As String) As String
      While pItems.Contains(",,")
        pItems = pItems.Replace(",,", ",")
      End While
      'BR 8771: Remove any last comma from end of line in case blank item(s) were at end:
      If pItems.EndsWith(",") Then pItems = pItems.Substring(0, pItems.Length - 1)
      Return pItems
    End Function

    Protected Shared Function ContactNameItems() As String
      'See ContactNameResults before changing
      Return "CONTACT_NAME,surname,forenames,title,initials"
    End Function

    Private Function ContactNameItems(ByVal pJustCommas As Boolean) As String
      'See ContactNameResults before changing
      Return ",,,,"
    End Function

    Private Function OrgNameItems() As String
      'See ContactNameResults before changing
      Return "name,,,,"
    End Function

    Protected Function ContactNameResults() As String
      'Used to specify additional contact fields - If changed ContactEventDelegates requires special handling
      ContactNameResults = "ContactName,Surname,Forenames,Title,Initials"
    End Function

    Private Function CheetahMailItems(ByVal pJustCommas As Boolean, Optional ByVal pContactEmailings As Boolean = False) As String
      'See CheetahMailResults before changing
      Dim vCheetahMailItems As String = "cheetah_mail_id,number_emails_bounced,number_emails_opened,number_emails_clicked,opened_on"
      If pJustCommas Then
        Return ",,,,"
      Else
        If Not pContactEmailings Then Return vCheetahMailItems.Replace("opened_on", "")
      End If
      Return vCheetahMailItems
    End Function

    Private Function CheetahMailAttrs(ByVal pInDatabase As Boolean) As String
      'See CheetahMailResults before changing
      Dim vCheetahMailAttrs As String
      vCheetahMailAttrs = ",cheetah_mail_id,number_emails_bounced,number_emails_opened,number_emails_clicked"
      If pInDatabase Then
        Return vCheetahMailAttrs.Replace("cheetah_mail_id", "mh.issue_id AS cheetah_mail_id")
      Else
        Return ""
      End If
    End Function

    Private Class EventSelectionInfo
      Public WhereFields As CDBFields
      Public ContactSort As String
      Public Where As String
      Public Table1 As String
      Public ContactAttrs As String
      Public ContactCols As String
      Public AnsiJoins As AnsiJoins
      Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pContact As Contact)
        Me.New(pEnv, pContact, "", "")
      End Sub
      Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pContact As Contact, ByVal pTableName1 As String, ByVal pTableName2 As String)
        Dim vContact As New Contact(pEnv)
        vContact.Init()
        WhereFields = New CDBFields
        AnsiJoins = New AnsiJoins
        If pContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          If pTableName1.Length > 0 Then
            'When organisation closes site address, not all address_numbers are updated to the client positions new address_number. 
            ' Therefore there may not be a contact_positions record for pTableName1 records (e.g. for delegates of historical events). For this reason 
            ' Organisation Addresses are used instead of Contact Positons to retrieve pTableName1 records for the organisation.
            Where = "oa.organisation_number = " & pContact.ContactNumber & " AND c.contact_number = t.contact_number"
            ContactSort = ", c.surname, c.forenames"
            AnsiJoins.Add(pTableName1, "oa.address_number", "t.address_number")
            AnsiJoins.Add(pTableName2, "t.contact_number", "c.contact_number")
            Table1 = "organisation_addresses oa"
            WhereFields.Add("oa.organisation_number", pContact.ContactNumber)
            'Now to tie this to current contacts at the organisations using contact positions table and ignoring address numbers
            Dim vSubSelectAttrs As String = "contact_number"
            Dim vSubSelectWhereFields As New CDBFields()
            vSubSelectWhereFields.Add("cp.contact_number", CDBField.FieldTypes.cftInteger, "t.contact_number", CDBField.FieldWhereOperators.fwoEqual)
            vSubSelectWhereFields.Add("cp.organisation_number", CDBField.FieldTypes.cftInteger, "oa.organisation_number", CDBField.FieldWhereOperators.fwoEqual)
            vSubSelectWhereFields.Add("current", "Y")
            vSubSelectWhereFields(3).SpecialColumn = True
            Dim vSubSelectSQL As New SQLStatement(pEnv.Connection, vSubSelectAttrs, "contact_positions cp", vSubSelectWhereFields, "")
            WhereFields.Add("t.contact_number", CDBField.FieldTypes.cftInteger, vSubSelectSQL.SQL, CDBField.FieldWhereOperators.fwoIn)
          Else
            Where = "cp.organisation_number = " & pContact.ContactNumber & " AND " & pEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND cp.contact_number = t.contact_number AND cp.contact_number = c.contact_number"
            ContactSort = ", c.surname, c.forenames"
            Table1 = "contact_positions cp, "
            WhereFields.Add("cp.organisation_number", pContact.ContactNumber)
            WhereFields.Add("current", "Y")
            'BR19676 - Should resolve duplicate rows when contact has multiple positions and also display booking where address is not related to contact position.
            WhereFields.Add("c.contact_type", "O", CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoCloseBracket)
            WhereFields.Add("c.contact_type#2", "O", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
            WhereFields.Add("cp.address_number", CDBField.FieldTypes.cftInteger, "t.address_number", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
            WhereFields.TableAlias = "cp"
            WhereFields(2).SpecialColumn = True
          End If
        Else
          Where = "t.contact_number = " & pContact.ContactNumber & " AND t.contact_number = c.contact_number"
          If pTableName1.Length > 0 Then AnsiJoins.Add(pTableName2, "t.contact_number", "c.contact_number")
          Table1 = ""
          ContactSort = ""
          WhereFields.Add("t.contact_number", pContact.ContactNumber)
        End If
        ContactAttrs = "," & Replace(vContact.GetRecordSetFieldsName, ",", ",c.")
        ContactCols = ",c.contact_number," & ContactNameItems()
      End Sub
    End Class

    Private Function GetCampaignWhereFields() As CDBFields
      Dim vWhereFields As New CDBFields
      With mvParameters
        'Campaign Fields
        If .Exists("Campaign") Then vWhereFields.Add("c.campaign", mvParameters("Campaign").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If .Exists("CampaignDesc") Then vWhereFields.Add("c.campaign_desc", mvParameters("CampaignDesc").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        If .Exists("DatedOnOrAfter") Then
          If .Exists("DatedOnOrBefore") Then
            vWhereFields.Add("c.start_date", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrAfter").Value, CDBField.FieldWhereOperators.fwoBetweenFrom)
            vWhereFields.Add("c.start_date2", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrBefore").Value, CDBField.FieldWhereOperators.fwoBetweenTo)
          Else
            vWhereFields.Add("c.start_date", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrAfter").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          End If
        ElseIf .Exists("DatedOnOrBefore") Then
          vWhereFields.Add("c.start_date", CDBField.FieldTypes.cftDate, mvParameters("DatedOnOrBefore").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
        End If
        If .Exists("Manager") Then vWhereFields.Add("c.manager", mvParameters("Manager").Value, CDBField.FieldWhereOperators.fwoEqual)
        If .Exists("BusinessType") Then vWhereFields.Add("c.campaign_business_type", mvParameters("BusinessType").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      End With
      Return vWhereFields
    End Function

    Private Function GetAppealWhereFields() As CDBFields
      'Appeal Fields
      Dim vWhereFields As New CDBFields
      With mvParameters
        If (.Exists("AppealType") = True Or .Exists("AppealManager") = True) Then
          If .Exists("AppealType") Then vWhereFields.Add("a.appeal_type", mvParameters("AppealType").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          If .Exists("AppealManager") Then vWhereFields.Add("a.manager", mvParameters("AppealManager").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        End If
        If (.Exists("Collection") = True Or .Exists("CollectionDesc") = True Or .Exists("GeographicalRegion") = True) Then
          vWhereFields.AddJoin("ac.campaign", "a.campaign")
          vWhereFields.AddJoin("ac.appeal", "a.appeal")
          If .Exists("Collection") Then vWhereFields.Add("ac.collection", mvParameters("Collection").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          If .Exists("CollectionDesc") Then vWhereFields.Add("ac.collection_desc", mvParameters("CollectionDesc").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          If .Exists("GeographicalRegion") Then
            vWhereFields.AddJoin("cr.collection_number", "ac.collection_number")
            vWhereFields.Add("cr.geographical_region", mvParameters("GeographicalRegion").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
            vWhereFields.AddJoin("gr.geographical_region", "cr.geographical_region")
            vWhereFields.Add("gr.geographical_region_type", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCollectionsRegionType))
          End If
        End If
      End With
      Return vWhereFields
    End Function

    Private Function GetCampaignRestriction(ByVal pAppealWhereFields As CDBFields) As String
      Dim vUseConfig As Boolean
      Dim vInClause As String = ""
      If mvEnv.GetConfig("maintain_owned_appeals") = "O" Then vUseConfig = True 'User can only view and maintain appeals where they have valid position at appeal manager's organisation!
      If (pAppealWhereFields.Count > 0) Or vUseConfig = True Then
        If vUseConfig Then vInClause = "users u1, contact_positions cp, users u2, "
        vInClause &= "appeals a"
        With mvParameters
          If (.Exists("Collection") = True Or .Exists("CollectionDesc") = True Or .Exists("GeographicalRegion") = True) Then
            vInClause &= ", appeal_collections ac"
            If .Exists("GeographicalRegion") Then
              vInClause &= ", collection_regions cr, geographical_regions gr"
            End If
          End If
        End With
        If mvType = DataSelectionTypes.dstCampaignInfo Then
          vInClause = "SELECT DISTINCT a.campaign,a.appeal,appeal_desc,appeal_type,appeal_date,end_date FROM " & vInClause & " WHERE "
        Else
          vInClause = "SELECT DISTINCT a.campaign FROM " & vInClause & " WHERE "
        End If
        If vUseConfig Then
          vInClause = vInClause & "u1.logname = '" & mvEnv.User.Logname & "'"
          vInClause = vInClause & " AND u1.contact_number = cp.contact_number AND u2.organisation_number = cp.organisation_number"
          vInClause = vInClause & " AND a.manager = u2.logname"
          If pAppealWhereFields.Count > 0 Then vInClause = vInClause & " AND "
        End If
        If pAppealWhereFields.Count > 0 Then vInClause = vInClause & mvEnv.Connection.WhereClause(pAppealWhereFields)
        'Note: If adding new Appeals attributes then may need to add them here as well as further down within this procedure (see comment below)
      End If
      Return vInClause
    End Function

    Private Function GetSelectionSetTableName() As String
      Dim vTable As String = "selected_contacts sc"
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "selection_group", "selection_sets", New CDBField("selection_set", mvParameters("SelectionSetNumber").LongValue))
      If vSQLStatement.GetValue = "AU" Then vTable = "selected_contacts_temp sc"
      Return vTable
    End Function

    Protected Sub GetDescriptions(ByVal pDataTable As CDBDataTable, ByVal pCode As String)
      Dim vCodeAttr As String = ""
      Dim vTable As String = ""
      Dim vDesc As String = ""
      Select Case pCode
        Case "Activity"
          vCodeAttr = "activity"
          vTable = "activities"
          vDesc = "ActivityDesc"
        Case "ActivityValue"
          vCodeAttr = "activity_value"
          vTable = "activity_values"
          vDesc = "ActivityValueDesc"
        Case "AuthorisationStatus"
          vCodeAttr = "authorisation_status"
          vTable = "authorisation_statuses"
          vDesc = "AuthorisationStatusDesc"
        Case "CancellationReason"
          vCodeAttr = "cancellation_reason"
          vTable = "cancellation_reasons"
          vDesc = "CancellationReasonDesc"
        Case "CancellationSource"
          vCodeAttr = "source"
          vTable = "sources"
          vDesc = "CancellationSourceDesc"
        Case "ChequeStatus"
          vCodeAttr = "cheque_status"
          vTable = "cheque_statuses"
          vDesc = "ChequeStatusDesc"
        Case "CovenantStatus"
          vCodeAttr = "covenant_status"
          vTable = "covenant_statuses"
          vDesc = "CovenantStatusDesc"
        Case "DistributionCode"
          vCodeAttr = "distribution_code"
          vTable = "distribution_codes"
          vDesc = "DistributionCodeDesc"
        Case "FutureCancellationReason"
          vCodeAttr = "cancellation_reason"
          vTable = "cancellation_reasons"
          vDesc = "FutureCancellationReasonDesc"
        Case "FutureCancellationSource"
          vCodeAttr = "source"
          vTable = "sources"
          vDesc = "FutureCancellationSourceDesc"
        Case "MailingTemplate"
          vCodeAttr = "mailing_template"
          vTable = "mailing_templates"
          vDesc = "MailingTemplateDesc"
        Case "PaymentFrequency"
          vCodeAttr = "payment_frequency"
          vTable = "payment_frequencies"
          vDesc = "PaymentFrequencyDesc"
        Case "PayPlanMembershipTypeCode"
          vCodeAttr = "membership_type"
          vTable = "membership_types"
          vDesc = "PayPlanMembershipTypeDesc"
        Case "PositionFunction"
          vCodeAttr = "position_function"
          vTable = "position_functions"
          vDesc = "PositionFunctionDesc"
        Case "PositionSeniority"
          vCodeAttr = "position_seniority"
          vTable = "position_seniorities"
          vDesc = "PositionSeniorityDesc"
        Case "Product"
          vCodeAttr = "product"
          vTable = "products"
          vDesc = "ProductDesc"
        Case "Rate"
          vCodeAttr = "rate"
          vTable = "rates"
          vDesc = "RateDesc"
        Case "ReasonForDespatch"
          vCodeAttr = "reason_for_despatch"
          vTable = "reasons_for_despatch"
          vDesc = "ReasonForDespatchDesc"
        Case "RenewalDateChangeReason"
          vCodeAttr = "renewal_change_reason"
          vTable = "renewal_change_reasons"
          vDesc = "RenewalDateChangeReasonDesc"
        Case "ScheduleCreationReason"
          vCodeAttr = "schedule_creation_reason"
          vTable = "schedule_creation_reasons"
          vDesc = "ScheduleCreationReasonDesc"
        Case "ScheduledPaymentStatus"
          vCodeAttr = "scheduled_payment_status"
          vTable = "scheduled_payment_statuses"
          vDesc = "ScheduledPaymentStatusDesc"
        Case "SourceCode"
          vCodeAttr = "source"
          vTable = "sources"
          vDesc = "SourceDesc"
        Case "StandardDocument"
          vCodeAttr = "standard_document"
          vTable = "standard_documents"
          vDesc = "StandardDocumentDesc"
        Case "Suppression"
          vCodeAttr = "mailing_suppression"
          vTable = "mailing_suppressions"
          vDesc = "SuppressionDesc"
        Case "Warehouse"
          vCodeAttr = "warehouse"
          vTable = "warehouses"
          vDesc = "WarehouseDesc"
      End Select

      Dim vCode As String
      Dim vList As New CDBParameters
      For Each vRow As CDBDataRow In pDataTable.Rows
        vCode = vRow.Item(pCode)
        If Len(vCode) > 0 Then
          If Not vList.Exists(vCode) Then vList.Add(vCode, vCode)
        End If
      Next
      If vList.Count > 0 Then
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add(vCodeAttr, vList.InList, CDBField.FieldWhereOperators.fwoIn)
        Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, String.Format("{0},{0}{1}", vCodeAttr, "_desc"), vTable, vWhereFields).GetRecordSet
        While vRecordSet.Fetch
          vCode = vRecordSet.Fields(1).Value
          For Each vRow As CDBDataRow In pDataTable.Rows
            If vCode = vRow.Item(pCode) Then
              vRow.Item(vDesc) = vRecordSet.Fields(2).Value
            End If
          Next
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GetSalesContactNames(ByVal pDataTable As CDBDataTable)
      GetContactNames(pDataTable, "SalesContactNumber", "SalesContactName")
    End Sub

    Private Sub GetContactNames(ByRef pDataTable As CDBDataTable, ByRef pNumberColumn As String, ByRef pNameColumn As String, Optional ByRef pAddressColumn As String = "", Optional ByRef pContactsOnly As Boolean = False)
      Dim vContactNumber As Integer
      Dim vList As New CDBParameters
      For Each vRow As CDBDataRow In pDataTable.Rows
        vContactNumber = IntegerValue(vRow.Item(pNumberColumn))
        If vContactNumber > 0 Then
          If Not vList.Exists(vContactNumber.ToString) Then vList.Add(vContactNumber.ToString, vContactNumber)
        End If
      Next vRow
      If vList.Count > 0 Then
        Dim vContact As New Contact(mvEnv)
        vContact.Init()
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry) & " FROM contacts c, addresses a, countries co WHERE contact_number IN(" & vList.ItemList & ") AND " & If(pContactsOnly, "c.contact_type <> 'O' AND ", "") & " c.address_number = a.address_number AND a.country = co.country")
        While vRecordSet.Fetch()
          vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry)
          For Each vRow As CDBDataRow In pDataTable.Rows
            vContactNumber = IntegerValue(vRow.Item(pNumberColumn))
            If vContactNumber = vContact.ContactNumber Then
              vRow.Item(pNameColumn) = vContact.Name
              If pAddressColumn.Length > 0 Then vRow.Item(pAddressColumn) = vContact.Address.AddressLine
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub CheckDashboardItem()
      If mvEnv.GetConfigOption("option_dashboard", False) = False Then
        mvResultColumns = mvResultColumns.Replace(",Dashboard", "")
        RemoveItem("Dashboard")
      End If
    End Sub

    Public Sub RemoveItem(ByVal pName As String)
      Dim vSelectParams As New StringList(mvSelectColumns, StringSplitOptions.None)
      Dim vHeadingParams As New StringList(mvHeadings, StringSplitOptions.None)
      Dim vWidthParams As New StringList(mvWidths, StringSplitOptions.None)
      For vIndex As Integer = 0 To vSelectParams.Count - 1
        If vSelectParams(vIndex) = pName Then
          vSelectParams.RemoveAt(vIndex)
          vHeadingParams.RemoveAt(vIndex)
          vWidthParams.RemoveAt(vIndex)
          mvSelectColumns = vSelectParams.ItemList
          mvHeadings = vHeadingParams.ItemList
          mvWidths = vWidthParams.ItemList
          Exit For
        End If
      Next
    End Sub

    Public Sub ReplaceItem(ByVal pName As String, ByVal pReplace As String)
      If mvSelectColumns.Contains(pName) Then
        Dim vSelectParams As New StringList(mvSelectColumns, StringSplitOptions.None)
        For vIndex As Integer = 0 To vSelectParams.Count - 1
          If vSelectParams(vIndex) = pName Then
            vSelectParams(vIndex) = pReplace
            mvSelectColumns = vSelectParams.ItemList
            Exit For
          End If
        Next
      End If
    End Sub

    Private Sub AddOwnerRestriction(ByRef pSQL As String)
      'NOTE: Use of this sub assumes Event Table Instance is "e"
      If mvEnv.GetConfigOption("ev_display_owned_bookings_only") Then
        pSQL = pSQL & " AND (e.department IS NULL OR e.department = '" & mvEnv.User.Department & "' OR e.event_number IN (SELECT event_number FROM event_owners ewn WHERE ewn.event_number = e.event_number AND ewn.department = '" & mvEnv.User.Department & "'))"
      End If
    End Sub

    Private Sub AddOwnerRestrictionToFields(ByVal pFields As CDBFields)
      'NOTE: Use of this sub assumes Event Table Instance is "e"
      If mvEnv.GetConfigOption("ev_display_owned_bookings_only") Then
        pFields.Add("e.department", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        pFields.Add("e.department#2", CDBField.FieldTypes.cftCharacter, mvEnv.User.Department, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR)
        pFields.Add("e.event_number#2", CDBField.FieldTypes.cftLong, "SELECT event_number FROM event_owners ewn WHERE ewn.event_number = e.event_number AND ewn.department = '" & mvEnv.User.Department & "'", CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
    End Sub

    Public ReadOnly Property DisplayListCode() As String
      Get
        Return mvCode
      End Get
    End Property

    Private Sub GetMemberFutureTypes(ByVal pDataTable As CDBDataTable)
      Dim vRow As CDBDataRow
      Dim vList As New CDBParameters
      Dim vRecordSet As CDBRecordSet
      Dim vCode As String

      For Each vRow In pDataTable.Rows
        vCode = vRow.Item("MembershipNumber")
        If Len(vCode) > 0 Then
          If Not vList.Exists(vCode) Then vList.Add(vCode, CDBField.FieldTypes.cftLong, vCode)
        End If
      Next vRow
      If vList.Count > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT membership_number, future_membership_type, future_change_date, membership_type_desc, mft.product, mft.rate, mft.amount FROM member_future_type mft, membership_types mt WHERE membership_number IN (" & vList.InList & ") AND mft.future_membership_type = mt.membership_type")
        While vRecordSet.Fetch()
          vCode = vRecordSet.Fields(1).Value
          For Each vRow In pDataTable.Rows
            If vCode = vRow.Item("MembershipNumber") Then
              vRow.Item("FutureMembershipType") = vRecordSet.Fields(2).Value
              vRow.Item("FutureChangeDate") = vRecordSet.Fields(3).Value
              vRow.Item("FutureMembershipTypeDesc") = vRecordSet.Fields(4).Value
              vRow.Item("FutureProduct") = vRecordSet.Fields(5).Value
              vRow.Item("FutureRate") = vRecordSet.Fields(6).Value
              vRow.Item("FutureAmount") = vRecordSet.Fields(7).Value
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GetCommunicationInfo(ByVal pDataTable As CDBDataTable, ByVal pNumberColumn As String, ByVal pNameColumn As String)
      Dim vRow As CDBDataRow
      Dim vList As New CDBParameters
      Dim vRecordSet As CDBRecordSet
      Dim vCommunication As New Communication(mvEnv)
      Dim vCommunicationNumber As Integer
      For Each vRow In pDataTable.Rows
        vCommunicationNumber = vRow.IntegerItem(pNumberColumn)
        If vCommunicationNumber > 0 Then
          If Not vList.Exists(vCommunicationNumber.ToString) Then vList.Add(vCommunicationNumber.ToString, vCommunicationNumber)
        End If
        vRow.Item(pNameColumn) = vRow.Item("AddressLine")
      Next vRow
      If vList.Count > 0 Then
        vCommunication.Init()
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vCommunication.GetRecordSetFields() & " FROM communications com WHERE communication_number IN (" & vList.ItemList & ")")
        While vRecordSet.Fetch()
          vCommunication.InitFromRecordSet(vRecordSet)
          For Each vRow In pDataTable.Rows
            vCommunicationNumber = vRow.IntegerItem(pNumberColumn)
            If vCommunicationNumber = vCommunication.CommunicationNumber Then
              vRow.Item(pNameColumn) = vCommunication.PhoneNumber
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GetCommunicationsData(ByRef pDataTable As CDBDataTable)
      Dim vRow As CDBDataRow
      Dim vCommsNumber As Integer
      Dim vList As New CDBParameters
      Dim vCommunication As New Communication(mvEnv)

      For Each vRow In pDataTable.Rows
        vCommsNumber = vRow.IntegerItem("CommunicationNumber")
        If vCommsNumber > 0 Then
          If Not vList.Exists(vCommsNumber.ToString) Then vList.Add(vCommsNumber.ToString, vCommsNumber)
        End If
      Next vRow
      If vList.Count > 0 Then
        vCommunication.Init()
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vCommunication.GetRecordSetFields() & " FROM communications com, devices d WHERE communication_number IN(" & vList.ItemList & ") AND com.device = d.device")
        While vRecordSet.Fetch()
          vCommunication.InitFromRecordSet(vRecordSet)
          For Each vRow In pDataTable.Rows
            vCommsNumber = vRow.IntegerItem("CommunicationNumber")
            If vCommsNumber = vCommunication.CommunicationNumber Then
              vRow.Item("AddressLine") = vCommunication.Number
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Protected Sub GetLookupData(ByVal pDataTable As CDBDataTable, ByVal pCode As String, ByVal pTableName As String, ByVal pAttributeName As String)
      Dim vRow As CDBDataRow
      Dim vCode As String
      Dim vList As New CDBParameters
      Dim vRecordSet As CDBRecordSet

      Dim vDesc As String = pCode & "Desc"
      For Each vRow In pDataTable.Rows
        vCode = vRow.Item(pCode)
        If vCode.Length > 0 Then
          If Not vList.Exists(vCode) Then vList.Add(vCode, vCode)
        End If
      Next vRow
      If vList.Count > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT lookup_code,lookup_desc FROM maintenance_lookup WHERE table_name = '" & pTableName & "' AND attribute_name = '" & pAttributeName & "' AND lookup_code IN(" & vList.InList & ")")
        While vRecordSet.Fetch()
          vCode = vRecordSet.Fields(1).Value
          For Each vRow In pDataTable.Rows
            If vCode = vRow.Item(pCode) Then
              vRow.Item(vDesc) = vRecordSet.Fields(2).Value
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub
    ''' <summary>
    ''' Get the Max (Last Exam Date) form the child exam units. This looks for all the child units for the current exam unit
    ''' and extracts the latest Last Exam Date for child exam. 
    ''' </summary>
    ''' <param name="pDataTable">Data Table</param>
    ''' <remarks></remarks>
    Private Sub GetChildLastExamDate(ByVal pDataTable As CDBDataTable)
      If pDataTable.Rows.Count > 0 Then
        Dim vRow As CDBDataRow
        Dim vFields As String = "max( last_exam_date) as LastExamDate"
        Dim vWhereFields As New CDBFields(New CDBField("esuh.exam_student_header_id", pDataTable.Rows(0).Item("ExamStudentHeaderId")))

        Dim vLinksWhereClause As New CDBFields(New CDBField("exam_unit_id_1", pDataTable.Rows(0).Item("ExamUnitId")))
        Dim vSubSQL2 As New SQLStatement(mvEnv.Connection, "exam_unit_id", "exam_student_unit_header", New CDBField("exam_student_header_id", pDataTable.Rows(0).Item("ExamStudentHeaderId")))
        vLinksWhereClause.Add("exam_unit_id_2", String.Format("({0})", vSubSQL2.SQL), CDBField.FieldWhereOperators.fwoIn)
        Dim vGetExamUnitLink As New SQLStatement(mvEnv.Connection, "exam_unit_id_2", "exam_unit_links", vLinksWhereClause)

        vWhereFields.Add("exam_unit_id", String.Format("( {0} )", vGetExamUnitLink.SQL), CDBField.FieldWhereOperators.fwoIn)

        Dim vAnsiJoins As New AnsiJoins
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "exam_student_unit_header esuh", vWhereFields)
        vSQLStatement.UseAnsiSQL = True
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQLStatement, 10)
        While vRecordSet.Fetch
          For Each vRow In pDataTable.Rows
            vRow.Item("ChildLastExamDate") = vRecordSet.Fields(1).Value
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub


    Private Sub GetNewOrderData(ByVal pDataTable As CDBDataTable, ByVal pNumberColumn As String, ByVal pPackToDonorColumn As String, Optional ByVal pGiftFromColumn As String = "", Optional ByVal pGiftToColumn As String = "", Optional ByVal pGiftMessageColumn As String = "")
      Dim vRow As CDBDataRow
      Dim vOrderNumber As Integer
      Dim vList As New CDBParameters
      Dim vAttrs As String

      For Each vRow In pDataTable.Rows
        vOrderNumber = vRow.IntegerItem(pNumberColumn)
        If vOrderNumber > 0 Then
          If Not vList.Exists(vOrderNumber.ToString) Then vList.Add(vOrderNumber.ToString, vOrderNumber)
        End If
      Next vRow
      If vList.Count > 0 Then
        vAttrs = "order_number, pack_to_donor"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftMessage) Then vAttrs = vAttrs & ", gift_from, gift_to, gift_message"
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAttrs & " FROM new_orders no WHERE order_number IN(" & vList.ItemList & ")")
        While vRecordSet.Fetch()
          For Each vRow In pDataTable.Rows
            If vRow.IntegerItem(pNumberColumn) = vRecordSet.Fields(1).LongValue Then
              If pPackToDonorColumn.Length > 0 Then
                vRow.Item(pPackToDonorColumn) = vRecordSet.Fields(2).Value
                vRow.SetYNValue(pPackToDonorColumn)
              End If
              If pGiftFromColumn.Length > 0 AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftMessage) Then
                vRow.Item(pGiftFromColumn) = vRecordSet.Fields(3).Value
                vRow.Item(pGiftToColumn) = vRecordSet.Fields(4).Value
                vRow.Item(pGiftMessageColumn) = Replace(vRecordSet.Fields(5).Value, vbCrLf, " ")
              End If
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Function CheckContactNameAttrs(ByRef pAttrs As String) As String
      Dim vContact As New Contact(mvEnv)
      Dim vItems As New CDBParameters
      Dim vContactItems As New CDBParameters
      Dim vParam As CDBParameter

      vContact.Init()
      vContactItems.InitFromUniqueList(Replace(vContact.GetRecordSetFieldsName, "c.contact_number,", ""))
      vItems.InitFromUniqueList(pAttrs)
      For Each vParam In vContactItems
        If Not vItems.Exists((vParam.Name)) Then vItems.Add((vParam.Name), CDBField.FieldTypes.cftCharacter, vParam.Value)
      Next vParam
      Return vItems.ItemList
    End Function

    Protected Sub GetActionersAndSubjects(ByRef pDataTable As CDBDataTable)
      Dim vActions As CDBParameters
      Dim vParams As CDBParameters
      Dim vDR As CDBDataRow
      Dim vContact As Contact
      Dim vAttrs As String
      Dim vSQL As String
      Dim vRS As CDBRecordSet

      vActions = New CDBParameters
      For Each vDR In pDataTable.Rows
        If Not vActions.Exists(vDR.Item("ActionNumber")) Then vActions.Add(vDR.Item("ActionNumber"))
      Next vDR
      'Continue as long as at least one unique action number exists
      If vActions.Count > 0 Then
        'Select records from contact_actions where type = R
        'get the contact's name and default phone number
        'if the contact's default address is an organisation address then get the org. number, org. name and position
        vContact = New Contact(mvEnv)
        vContact.Init()
        vAttrs = vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone) & ", x.name, x.position, x.organisation_number, ca.action_number"
        vSQL = "SELECT " & vAttrs
        vSQL = vSQL & " FROM contact_actions ca"
        vSQL = vSQL & " INNER JOIN contacts c ON ca.contact_number = c.contact_number"
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT oa.address_number, cp.contact_number, o.organisation_number, o.name, cp.position"
        vSQL = vSQL & " FROM organisation_addresses oa"
        vSQL = vSQL & " INNER JOIN organisations o ON oa.organisation_number = o.organisation_number"
        vSQL = vSQL & " INNER JOIN contact_positions cp ON oa.address_number = cp.address_number AND o.organisation_number = cp.organisation_number) x"
        vSQL = vSQL & " ON c.address_number = x.address_number AND c.contact_number = x.contact_number"
        vSQL = vSQL & " WHERE ca.action_number IN (" & vActions.ItemList & ") AND type = 'A'"
        'Now that we've used the ItemList property reinitialise the local Parameters collection
        vParams = New CDBParameters
        'Ensure that the SELECT can be executed on Oracle databases
        vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
        'Build the record set
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        With vRS
          While .Fetch()
            For Each vDR In pDataTable.Rows
              If vDR.IntegerItem("ActionNumber") = .Fields.Item("action_number").IntegerValue Then
                If Not vParams.Exists(vDR.Item("ActionNumber")) Then
                  'This is the first time this action has been encountered in this loop
                  vParams.Add(vDR.Item("ActionNumber"))
                  vContact = New Contact(mvEnv)
                  vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone)
                  vDR.Item("ContactNumber") = vContact.ContactNumber.ToString
                  vDR.Item("ContactName") = vContact.Name
                  vDR.Item("Name") = .Fields.Item("name").Value
                  vDR.Item("Position") = .Fields.Item("position").Value
                  vDR.Item("PhoneNumber") = vContact.PhoneNumber
                  vDR.Item("OrganisationNumber") = .Fields.Item("organisation_number").Value
                Else
                  'This is not the first time this action has been encountered in this loop
                  'This action is related to multiple contacts, so we can't display an individual contact's details
                  vDR.Item("ContactNumber") = ""
                  vDR.Item("ContactName") = "<Multiple Actioners>"
                  vDR.Item("Name") = ""
                  vDR.Item("Position") = ""
                  vDR.Item("PhoneNumber") = ""
                  vDR.Item("OrganisationNumber") = ""
                End If
              End If
            Next vDR
          End While
          .CloseRecordSet()
        End With

        vAttrs = "cas.action_number,cas.topic,cas.sub_topic,topic_desc,sub_topic_desc"
        vSQL = "SELECT " & vAttrs
        vSQL = vSQL & " FROM action_subjects cas"
        vSQL = vSQL & " INNER JOIN topics t ON cas.topic = t.topic"
        vSQL = vSQL & " INNER JOIN sub_topics st ON cas.topic = st.topic AND cas.sub_topic = st.sub_topic"
        vSQL = vSQL & " WHERE cas.action_number IN (" & vActions.ItemList & ")"
        'Now that we've used the ItemList property reinitialise the local Parameters collection
        vParams = New CDBParameters
        'Ensure that the SELECT can be executed on Oracle databases
        vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
        'Build the record set
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        With vRS
          While .Fetch()
            For Each vDR In pDataTable.Rows
              If vDR.IntegerItem("ActionNumber") = .Fields.Item("action_number").IntegerValue Then
                If Not vParams.Exists(vDR.Item("ActionNumber")) Then
                  'This is the first time this action has been encountered in this loop
                  vParams.Add(vDR.Item("ActionNumber"))
                  vDR.Item("Topic") = .Fields.Item("topic").Value
                  vDR.Item("SubTopic") = .Fields.Item("sub_topic").Value
                  vDR.Item("TopicDesc") = .Fields.Item("topic_desc").Value
                  vDR.Item("SubTopicDesc") = .Fields.Item("sub_topic_desc").Value
                Else
                  'This is not the first time this action has been encountered in this loop
                  'This action is related to multiple topics se we cannot display an individual subject details
                  vDR.Item("Topic") = ""
                  vDR.Item("SubTopic") = ""
                  vDR.Item("TopicDesc") = "<Multiple Topics>"
                  vDR.Item("SubTopicDesc") = "<Multiple Sub Topics>"
                End If
              End If
            Next vDR
          End While
          .CloseRecordSet()
        End With
      End If
    End Sub

    Private Sub GetAddressData(ByVal pDataTable As CDBDataTable, Optional ByVal pGetOrganisation As Boolean = False)
      Dim vRow As CDBDataRow
      Dim vAddressNumber As Integer
      Dim vList As New CDBParameters
      Dim vRecordSet As CDBRecordSet
      Dim vAddress As New Address(mvEnv)

      For Each vRow In pDataTable.Rows
        If vRow.Item("AddressLine").Length = 0 Then
          vAddressNumber = vRow.IntegerItem("AddressNumber")
          If vAddressNumber > 0 Then
            If Not vList.Exists(vAddressNumber.ToString) Then vList.Add(vAddressNumber.ToString, vAddressNumber)
          End If
        End If
      Next vRow
      If vList.Count > 0 Then
        vAddress.Init()
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAddress.GetRecordSetFields(Address.AddressRecordSetTypes.artAll) & " FROM addresses a, countries co WHERE address_number IN(" & vList.ItemList & ") AND a.country = co.country")
        While vRecordSet.Fetch()
          vAddress.InitFromRecordSet(mvEnv, vRecordSet, Address.AddressRecordSetTypes.artAll)
          For Each vRow In pDataTable.Rows
            If vRow.Item("AddressLine").Length = 0 Then
              vAddressNumber = vRow.IntegerItem("AddressNumber")
              If vAddressNumber = vAddress.AddressNumber Then
                vRow.Item("AddressLine") = vAddress.AddressLine
                If pGetOrganisation Then
                  vRow.Item("OrganisationName") = vAddress.OrganisationName
                End If
              End If
            End If
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GetPaymentPlanInfo(ByRef pDataTable As CDBDataTable)
      Dim vLineNo As Integer
      Dim vRow As CDBDataRow
      Dim vMembership As Boolean
      Dim vCovenant As Boolean

      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT line_number,order_type,oph.order_number,oph.payment_number FROM order_payment_history oph, orders o WHERE batch_number = " & mvParameters("BatchNumber").LongValue & " AND transaction_number = " & mvParameters("TransactionNumber").LongValue & " AND oph.order_number = o.order_number ORDER BY line_number")
      While vRecordSet.Fetch
        vLineNo = vRecordSet.Fields(1).LongValue
        For Each vRow In pDataTable.Rows
          If vLineNo = vRow.IntegerItem("LineNumber") Then
            vRow.Item("PaymentPlanType") = vRecordSet.Fields(2).Value
            Select Case vRecordSet.Fields(2).Value
              Case "M"
                vMembership = True
              Case "C"
                vCovenant = True
            End Select
            vRow.Item("PaymentPlanNumber") = vRecordSet.Fields(3).Value
            vRow.Item("Number") = vRecordSet.Fields(3).Value
            vRow.Item("PaymentPlanPayNumber") = vRecordSet.Fields(4).Value
          End If
        Next vRow
      End While
      vRecordSet.CloseRecordSet()
      If vMembership Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT line_number,member_number FROM order_payment_history oph, orders o, members m WHERE batch_number = " & mvParameters("BatchNumber").LongValue & " AND transaction_number = " & mvParameters("TransactionNumber").LongValue & " AND oph.order_number = o.order_number AND o.order_number = m.order_number ORDER BY line_number")
        While vRecordSet.Fetch
          vLineNo = vRecordSet.Fields(1).LongValue
          For Each vRow In pDataTable.Rows
            If vLineNo = vRow.IntegerItem("LineNumber") Then vRow.Item("Number") = vRecordSet.Fields(2).Value
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
      If vCovenant Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT line_number,covenant_number FROM order_payment_history oph, orders o, covenants c WHERE batch_number = " & mvParameters("BatchNumber").LongValue & " AND transaction_number = " & mvParameters("TransactionNumber").LongValue & " AND oph.order_number = o.order_number AND o.order_number = c.order_number ORDER BY line_number")
        While vRecordSet.Fetch
          vLineNo = vRecordSet.Fields(1).LongValue
          For Each vRow In pDataTable.Rows
            If vLineNo = vRow.IntegerItem("LineNumber") Then vRow.Item("Number") = vRecordSet.Fields(2).Value
          Next vRow
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GetPayrollInfo(ByRef pDataTable As CDBDataTable)
      Dim vRow As CDBDataRow
      Dim vList As New CDBParameters
      Dim vOrganisation As Integer

      Dim vEmpRelationship As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlPostTaxPGEmplrPayrollRelationship)
      If vEmpRelationship.Length > 0 Then
        For Each vRow In pDataTable.Rows
          vOrganisation = vRow.IntegerItem("EmployerOrganisationNumber")
          If Not vList.Exists(vOrganisation.ToString) Then vList.Add(vOrganisation.ToString, vOrganisation)
        Next vRow
        If vList.Count > 0 Then
          Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT organisation_number_1, organisation_number_2, name from organisation_links ol, organisations o WHERE ol.organisation_number_1 IN(" & vList.InList & ") AND relationship = '" & vEmpRelationship & "' AND o.organisation_number = ol.organisation_number_2")
          While vRecordSet.Fetch()
            For Each vRow In pDataTable.Rows
              If vRecordSet.Fields(1).LongValue = vRow.IntegerItem("EmployerOrganisationNumber") Then
                vRow.Item("EmployerPayrollOrganisationNumber") = vRecordSet.Fields(2).Value
                vRow.Item("EmployerPayrollOrganisationName") = vRecordSet.Fields(3).Value
              End If
            Next vRow
          End While
          vRecordSet.CloseRecordSet()
        End If
      End If
    End Sub

    Private Sub HideColumn(ByVal pColumn As String)
      Dim vWidths() As String = mvWidths.Split(","c)
      Dim vHeadings() As String = mvSelectColumns.Split(","c)
      For vIndex As Integer = 0 To vHeadings.Length - 1
        If vHeadings(vIndex) = pColumn Then vWidths(vIndex) = "1"
      Next
      mvWidths = String.Join(",", vWidths)
    End Sub

    Private Sub SetMeetingRoleDesc(ByRef pDataTable As CDBDataTable)
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT DISTINCT ml.meeting_role,meeting_role_desc FROM meeting_links ml, meeting_roles mr WHERE meeting_number = " & mvParameters("MeetingNumber").LongValue & " AND ml.meeting_role = mr.meeting_role")
      While vRecordSet.Fetch()
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("MeetingRoleCode") = vRecordSet.Fields(1).Value Then vRow.Item("MeetingRoleDesc") = vRecordSet.Fields(2).Value
        Next vRow
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Sub SelectScheduledPaymentData(ByRef pDataTable As CDBDataTable)
      'Populate DataTable for scheduled payment data - (1) all OPS data, (2) Payment History only (pre-OPS data)
      Dim vAttrs As String
      Dim vItems As String
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vSqlStatement As SQLStatement

      '(1) Add all ops data
      vAttrs = "ops.scheduled_payment_number,amount_due,amount_outstanding,revised_amount,due_date,claim_date,scheduled_payment_status_desc,expected_balance,schedule_creation_reason_desc,payment_number,oph.amount,oph.balance,oph.posted,oph.status,oph.batch_number,oph.transaction_number,oph.line_number,payment_method_desc,fh.transaction_date,fh.posted AS posted_date,ops.scheduled_payment_status,oph.write_off_line_amount,i.invoice_pay_status,ips.invoice_pay_status_desc," & mvEnv.Connection.DBIsNull("fh.contact_number", "bt.contact_number") & " AS payer_contact_number,ops.schedule_creation_reason"
      vItems = "scheduled_payment_number,amount_due,amount_outstanding,revised_amount,due_date,claim_date,scheduled_payment_status_desc,expected_balance,schedule_creation_reason_desc,payment_number,amount,balance,posted,status,batch_number,transaction_number,line_number,payment_method_desc,transaction_date,posted_date,scheduled_payment_status,write_off_line_amount,invoice_pay_status,invoice_pay_status_desc,payer_contact_number,schedule_creation_reason"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbBankTransactionsImport) Then
        vAttrs = vAttrs.Replace(",oph.write_off_line_amount", "")
        vItems = vItems.Replace("write_off_line_amount", "")
      End If
      With vAnsiJoins
        .Add("scheduled_payment_statuses sps", "ops.scheduled_payment_status", "sps.scheduled_payment_status")
        .Add("schedule_creation_reasons scr", "ops.schedule_creation_reason", "scr.schedule_creation_reason")
        .AddLeftOuterJoin("order_payment_history oph", "ops.scheduled_payment_number", "oph.scheduled_payment_number", "ops.order_number", "oph.order_number")
        .AddLeftOuterJoin("financial_history fh", "oph.batch_number", "fh.batch_number", "oph.transaction_number", "fh.transaction_number")
        .AddLeftOuterJoin("payment_methods pm", "fh.payment_method", "pm.payment_method")
        .AddLeftOuterJoin("invoices i", "oph.batch_number", "i.batch_number", "oph.transaction_number", "i.transaction_number")
        .AddLeftOuterJoin("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
        .AddLeftOuterJoin("batch_transactions bt", "oph.batch_number", "bt.batch_number", "oph.transaction_number", "bt.transaction_number")
      End With
      With vWhereFields
        .Add("ops.order_number", mvParameters("PaymentPlanNumber").LongValue)
        .Add("ops.scheduled_payment_status", CDBField.FieldTypes.cftCharacter, "V", CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("ops.scheduled_payment_status#2", CDBField.FieldTypes.cftCharacter, "V", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual)
        .Add("oph.batch_number", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With
      Dim vOrderBy As String = "ops.due_date DESC, ops.claim_date " & If(mvEnv.Connection.NullsSortAtEnd, "", "DESC") & ", oph.payment_number DESC"
      vSqlStatement = New SQLStatement(mvEnv.Connection, vAttrs, "order_payment_schedule ops", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSqlStatement, vItems)

      '(2) Add oph that does not link to ops (old data)
      Dim vRecordCount As Integer = pDataTable.Rows.Count
      Dim vReOrderTableColumns As Boolean = False
      vAttrs = "payment_number,oph.amount,balance,oph.posted,oph.status,oph.batch_number,oph.transaction_number,oph.line_number,payment_method_desc,transaction_date,fh.posted AS posted_date,oph.write_off_line_amount,i.invoice_pay_status,ips.invoice_pay_status_desc,fh.contact_number"
      vItems = ",,,,,,,,,payment_number,amount,balance,posted,status,batch_number,transaction_number,line_number,payment_method_desc,transaction_date,posted_date,,write_off_line_amount,,,,invoice_pay_status,invoice_pay_status_desc,contact_number"
      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbBankTransactionsImport) Then
        vAttrs = vAttrs.Replace(",oph.write_off_line_amount", "")
        vItems = vItems.Replace("write_off_line_amount", "")
      End If
      vWhereFields = New CDBFields
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, mvParameters("PaymentPlanNumber").LongValue)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) Then vWhereFields.Add("scheduled_payment_number", CDBField.FieldTypes.cftLong, "")
      vAnsiJoins = New AnsiJoins
      vAnsiJoins.Add("financial_history fh", "oph.batch_number", "fh.batch_number", "oph.transaction_number", "fh.transaction_number")
      vAnsiJoins.Add("payment_methods pm", "fh.payment_method", "pm.payment_method")
      vAnsiJoins.AddLeftOuterJoin("invoices i", "oph.batch_number", "i.batch_number", "oph.transaction_number", "i.transaction_number")
      vAnsiJoins.AddLeftOuterJoin("invoice_pay_statuses ips", "i.invoice_pay_status", "ips.invoice_pay_status")
      vSqlStatement = New SQLStatement(mvEnv.Connection, vAttrs, "order_payment_history oph", vWhereFields, "transaction_date DESC, payment_number DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSqlStatement, vItems)
      If pDataTable.Rows.Count > vRecordCount Then vReOrderTableColumns = True 'Only need to re-order the data if this has added some more records

      '(3) Finally ensure that all posted records show the posted flag and paid provisional lines are hidden
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("ScheduledPaymentNumber").Length = 0 AndAlso vRow.Item("Posted").Length = 0 Then
          vRow.Item("Posted") = "Y"
        End If
        If vRow.Item("ScheduledPaymentStatus") = "V" Then
          With vRow
            .Item("ScheduledPaymentNumber") = ""
            .Item("AmountDue") = ""
            .Item("AmountOutstanding") = ""
            .Item("RevisedAmount") = ""
            .Item("DueDate") = ""
            .Item("ClaimDate") = ""
            .Item("ScheduledPaymentStatusDesc") = ""
            .Item("ExpectedBalance") = ""
            .Item("ScheduleCreationReasonDesc") = ""
          End With
        End If
      Next vRow
      If vReOrderTableColumns Then
        Dim vSS(0) As CDBDataTable.SortSpecification
        vSS(0).Column = "DueDate"
        vSS(0).Descending = True
        Dim vPP As New PaymentPlan
        vPP.Init(mvEnv, (mvParameters("PaymentPlanNumber").LongValue))
        If vPP.Existing Then
          If (vPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or vPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) Then
            ReDim Preserve vSS(1)
            vSS(1).Column = "ClaimDate"
            vSS(1).Descending = True
          End If
        End If
        pDataTable.ReOrderRowsByMultipleColumns(vSS)
      End If
    End Sub

    Private Sub DataTableTransactionDetails(ByRef pDataTable As CDBDataTable)
      Dim vContact As New Contact(mvEnv)
      Dim vDataRow As CDBDataRow
      Dim vRecordSet As CDBRecordSet
      Dim vAttrs As String
      Dim vSQL As String

      'This SQL needs to be run as a client-side cursor
      vContact.Init()
      vAttrs = Replace(Replace(vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry), "c.contact_number,", ""), "a.address_number,", "") & ",o.name"
      vAttrs = "batch_type,posted_to_nominal,bt.batch_number,bt.transaction_number,bt.contact_number,bt.address_number,transaction_date,bt.transaction_type,bank_details_number,amount,currency_amount,bt.payment_method,reference,next_line_number,line_total,mailing,receipt,bt.notes,mailing_contact_number,mailing_address_number,eligible_for_gift_aid,payment_method_desc,bt.amended_by,bt.amended_on," & vAttrs & ", r.was_transaction_number,i.record_type,ownership_access_level AS AccessLevel,transaction_sign,i.reprint_count"
      vSQL = "SELECT /* SQLServerCSC */ " & vAttrs & " FROM batches b INNER JOIN batch_transactions bt ON b.batch_number = bt.batch_number"
      vSQL = vSQL & " INNER JOIN payment_methods pm ON bt.payment_method = pm.payment_method"
      vSQL = vSQL & " INNER JOIN contacts c ON bt.contact_number = c.contact_number"
      'BR12995: We dont need to to join contact_addresses table
      'vSQL = vSQL & " INNER JOIN contact_addresses ca ON c.contact_number = ca.contact_number AND bt.address_number = ca.address_number"
      'vSQL = vSQL & " INNER JOIN addresses a ON ca.address_number = a.address_number"
      vSQL = vSQL & " INNER JOIN addresses a ON bt.address_number = a.address_number"
      vSQL = vSQL & " INNER JOIN countries co ON a.country = co.country"
      vSQL = vSQL & " INNER JOIN ownership_group_users ogu ON c.ownership_group = ogu.ownership_group"
      vSQL = vSQL & " INNER JOIN transaction_types tt ON bt.transaction_type = tt.transaction_type"
      vSQL = vSQL & " LEFT OUTER JOIN organisation_addresses oa ON a.address_number = oa.address_number"
      vSQL = vSQL & " LEFT OUTER JOIN organisations o ON oa.organisation_number = o.organisation_number"
      vSQL = vSQL & " LEFT OUTER JOIN (select r.batch_number,r.transaction_number,r.was_transaction_number from reversals r group by batch_number,transaction_number,was_transaction_number) r ON bt.batch_number = r.batch_number AND bt.transaction_number = r.transaction_number"
      vSQL = vSQL & " LEFT OUTER JOIN invoices i ON bt.batch_number = i.batch_number AND bt.transaction_number = i.transaction_number"
      vSQL = vSQL & " WHERE b.batch_number = " & mvParameters("BatchNumber").LongValue
      If mvParameters.Exists("TransactionNumber") Then vSQL = vSQL & " AND bt.transaction_number = " & mvParameters("TransactionNumber").LongValue
      vSQL = vSQL & mvEnv.User.OwnershipSelect("c", False, , True)
      vSQL = vSQL & " ORDER BY bt.transaction_number DESC"
      vAttrs = vAttrs & ",CONTACT_NAME,ADDRESS_LINE"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQL), vAttrs, ",")
      pDataTable.Columns.Item("Printed").FieldType = CDBField.FieldTypes.cftCharacter
      Dim vCSBatch As Boolean = False
      If pDataTable.Rows.Count() > 0 Then
        If Batch.GetBatchType(pDataTable.Rows.Item(0).Item("BatchType")) = Batch.BatchTypes.CreditSales Then
          vCSBatch = True
          'Mark any invoices that have been paid as being allocated
          If (mvEnv.GetConfigOption("fp_use_sales_ledger", True) = True And pDataTable.Rows.Item(0).Item("PostedToNominal") = "N") Then
            vSQL = "SELECT bt.transaction_number, iph.invoice_number"
            vSQL = vSQL & " FROM batch_transactions bt, invoices i, invoice_payment_history iph"
            vSQL = vSQL & " WHERE bt.batch_number = " & mvParameters("BatchNumber").LongValue
            If mvParameters.Exists("TransactionNumber") Then vSQL = vSQL & " AND bt.transaction_number = " & mvParameters("TransactionNumber").LongValue
            vSQL = vSQL & " AND i.batch_number = bt.batch_number AND i.transaction_number = bt.transaction_number"
            vSQL = vSQL & " AND iph.invoice_number = i.invoice_number ORDER BY bt.transaction_number DESC"
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            While vRecordSet.Fetch
              For Each vDataRow In pDataTable.Rows
                If Val(vDataRow.Item("TransactionNumber")) = vRecordSet.Fields("transaction_number").LongValue Then
                  vDataRow.Item("Allocated") = "Y"
                End If
              Next vDataRow
            End While
            vRecordSet.CloseRecordSet()
          End If
        End If
      End If
      For Each vDataRow In pDataTable.Rows
        If Len(vDataRow.Item("Adjustment")) > 0 Then vDataRow.Item("Adjustment") = "Y"
        vDataRow.SetYNValue("EligibleForGiftAid", True)
        If vCSBatch Then
          vDataRow.Item("Printed") = If(vDataRow.Item("Printed").Length > 0 AndAlso IntegerValue(vDataRow.Item("Printed")) >= 0, "Y", String.Empty)
          vDataRow.SetYNValue("Printed")
        End If
      Next vDataRow
    End Sub

    ''' <summary>Restrict display of exam Results.</summary>
    ''' <remarks>For each row in the CDBDataTable, set the Current Mark / Grade / Result to the Previous Mark / Grade / Result when the ResultsReleaseDate is in the future</remarks>
    Friend Sub RestrictExamResults(ByVal pDataTable As CDBDataTable, ByVal pColumnPrefix As String)
      Dim vRestrictResults As Boolean = False
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        Dim vWhereFields As New CDBFields(New CDBField("department", mvEnv.User.Department))
        If mvEnv.Connection.GetCount("exam_result_unrestricted_depts", vWhereFields) = 0 Then vRestrictResults = True
      End If

      If vRestrictResults = True AndAlso pDataTable.Columns.ContainsKey("ResultsReleaseDate") = True Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          If IsDate(vRow.Item("ResultsReleaseDate")) AndAlso CDate(vRow.Item("ResultsReleaseDate")) > Today Then
            vRow.Item(pColumnPrefix & "Mark") = vRow.Item("PreviousMark")
            vRow.Item(pColumnPrefix & "Grade") = vRow.Item("PreviousGrade")
            vRow.Item(pColumnPrefix & "Result") = vRow.Item("PreviousResult")
            If pDataTable.Columns.ContainsKey(pColumnPrefix & "GradeDesc") Then
              vRow.Item(pColumnPrefix & "GradeDesc") = vRow.Item("PreviousGradeDesc")
              vRow.Item(pColumnPrefix & "ResultDesc") = vRow.Item("PreviousResultDesc")
            End If
            vRow.Item("CanEditResults") = "N"
            vRow.Item("PreviousMark") = ""
            vRow.Item("PreviousGrade") = ""
            vRow.Item("PreviousResult") = ""
            If pDataTable.Columns.ContainsKey("PreviousGradeDesc") Then
              vRow.Item("PreviousGradeDesc") = ""
              vRow.Item("PreviousResultDesc") = ""
            End If
            If pDataTable.Columns.ContainsKey("AchievedUnitCode") Then
              vRow.Item("AchievedUnitCode") = ""
              vRow.Item("AchievedUnitDescription") = ""
            End If
            If pDataTable.Columns.ContainsKey("OriginalMark") Then
              vRow.Item("OriginalMark") = vRow.Item("PreviousMark")
              vRow.Item("OriginalGrade") = vRow.Item("PreviousGrade")
              vRow.Item("OriginalResult") = vRow.Item("PreviousResult")
            End If
            If pDataTable.Columns.ContainsKey("ModeratedMark") Then 'BR20535
              vRow.Item("ModeratedMark") = String.Empty
              vRow.Item("ModeratedGrade") = String.Empty
              vRow.Item("ModeratedResult") = String.Empty
            End If
            If pDataTable.Columns.ContainsKey("FirstPassed") Then vRow.Item("FirstPassed") = String.Empty
            If pDataTable.Columns.ContainsKey("Expires") Then vRow.Item("Expires") = String.Empty
          End If
        Next
      End If
    End Sub

    Private Sub GetMembershipLookupGroupSQL(ByVal pAnsiJoins As AnsiJoins, ByVal pWhereFields As CDBFields)
      pAnsiJoins.Add("lookup_groups lg", "r.membership_lookup_group", "lg.lookup_group", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      pAnsiJoins.Add("lookup_group_details lgd", "lg.lookup_group", "lgd.lookup_group", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      pAnsiJoins.Add("members m", "lgd.lookup_item", "m.membership_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)

      Dim vPositionWhere As New CDBFields
      vPositionWhere.Add("cp.contact_number", mvParameters("ContactNumber").IntegerValue)
      vPositionWhere.Add(mvEnv.Connection.DBSpecialCol("cp", "current"), "Y")
      Dim vPositionSelect As New SQLStatement(mvEnv.Connection, "cp.organisation_number", "contact_positions cp", vPositionWhere)
      Dim vInClause As String = String.Format("({0})", vPositionSelect.SQL)

      pWhereFields.Add("m.cancelled_on", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoEqual)
      pWhereFields.Add("lg.table_name", "", CDBField.FieldWhereOperators.fwoOpenBracket)
      pWhereFields.Add("lg.table_name#2", "membership_types", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      pWhereFields.Add("m.contact_number", "", CDBField.FieldWhereOperators.fwoOpenBracketTwice)
      pWhereFields.Add("r.membership_lookup_group", "", CDBField.FieldWhereOperators.fwoCloseBracket)
      pWhereFields.Add("m.contact_number#2", mvParameters("ContactNumber").IntegerValue, CDBField.FieldWhereOperators.fwoOR)
      pWhereFields.Add("m.contact_number#3", CDBField.FieldTypes.cftInteger, vInClause, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoCloseBracket)
    End Sub

    Private Function GetLinkEntityTypeDescription(ByVal pEntityCode As String) As String
      If mvLinkEntityTypes Is Nothing Then
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "link_entity_type, link_entity_type_desc", "link_entity_types lt", New CDBFields(), "link_entity_type")
        mvLinkEntityTypes = vSQLStatement.GetDataTable()
      End If

      Dim vRows() As DataRow = mvLinkEntityTypes.Select("link_entity_type = '" & pEntityCode & "'")
      If vRows.Length > 0 Then
        Return vRows(0).Item("link_entity_type_desc").ToString
      Else
        Return String.Empty
      End If
    End Function

    ''' <summary>Check if the Amount and CurrencyAmount columns are the same value. If they are not then clear the RGB value for the Amount.</summary>
    Protected Sub CheckAmountRGBValue(ByVal pRow As CDBDataRow)
      If String.IsNullOrEmpty(pRow.Item("RgbAmount")) = False Then
        'If Amount & CurrencyAmount are equal, keep their RGB values
        'If Amount & CurrencyAmount are NOT equal, only the CurrencyAmount should have an RGB value
        If DoubleValue(pRow.Item("Amount")).Equals(DoubleValue(pRow.Item("CurrencyAmount"))) = False Then pRow.Item("RgbAmount") = String.Empty
      End If
    End Sub

  End Class

End Namespace
