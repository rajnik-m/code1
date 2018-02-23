Imports Microsoft.Win32
Imports System.IO
Imports System.Configuration
Imports System.Web.Configuration
Imports CARE.Config

Namespace Access

  Public Enum JournalTypes As Integer 'If adding new types you need to update GetJournalType() function too
    jnlContact = 1
    jnlOrganisation
    jnlAddress
    jnlRelationship
    jnlActivity
    jnlSuppression
    jnlPosition
    jnlRole
    jnlDocument
    jnlAction
    jnlEvent
    jnlAccomodation
    jnlMeeting
    jnlAppointment
    jnlDonation
    jnlStockItem
    jnlEventBooking
    jnlAccomodationBooking
    jnlServiceBooking
    jnlOtherProduct
    jnlMember
    jnlCovenant
    jnlStandingOrder
    jnlDirectDebit
    jnlCreditCard
    jnlPayPlan
    jnlMailing
    jnlActionActioner
    jnlActionManager
    jnlActionRelated
    jnlGiftAidDeclaration
    jnlPledge
    jnlGoneAway
    jnlDirectDebitMaintenance
    jnlCreditCardMaintenance
    jnlStandingOrderMaintenance
    jnlMemberMaintenance
    jnlCovenantMaintenance
    jnlPayPlanMaintenance
    jnlPayPlanPaymentSchedule
    jnlNumber
    jnlCPDCycles
    jnlCPDPoints
    jnlIrishAppropriateCertificate
    jnlLogin
    jnlRegisteredUser
    jnlWebProductSale
    jnlWebPaymentPlanPayment
    jnlDirectory
    jnlDownload
  End Enum

  Public Enum JournalOperations As Integer
    jnlInsert = 1
    jnlUpdate
    jnlDelete
    jnlActive
    jnlComplete
    jnlCancel
    jnlReinstate
    jnlSearch
    jnlView
    jnlDowloaded
  End Enum

  Public Enum OutputDirectoryTypes As Integer
    scodtMailing
    scodtOutput
    scodtLogFiles
    scodtAuditFiles
  End Enum

  Partial Public Class CDBEnvironment
    Implements IDisposable

    Private Const MAX_RETRIES As Integer = 10
    Private Const FIRST_CUSTOM_FORM As Integer = 100
    Private Const LAST_CUSTOM_FORM As Integer = 499

    Public Enum AuditTypes
      audInsert = 1
      audUpdate = 2
      audDelete = 3
    End Enum

    Public Enum AuditStyleTypes
      ausNone = 0
      ausAuditMultipleRecords
      ausAuditOneRecord
      ausAmendmentHistory
      ausExtended
    End Enum

    Public Enum CachedControlNumberTypes
      ccnNone
      ccnJournal
      ccnPaymentSchedule
      ccnTimesheet
      ccnContact
      ccnAddress
      ccnAddressLink
      ccnPosition
      ccnExamMarkingBatchDetail
      ccnExamCentreUnit
    End Enum

    Public Enum OwnershipMethods
      omOwnershipGroups
      omOwnershipDepartments
    End Enum

    Public Enum OwnershipAccessLevelTypes
      oaltNone
      oaltBrowse
      oaltRead
      oaltWrite
    End Enum

    'The following enum is used by GetDataStructureInfo
    Public Enum cdbDataStructureConstants
      cdbDataBTAQuantityDecimal         'Batch Transaction Analysis Quantity Decimal
      cdbDataCurrencyCode               'Table currency_codes exists

      cdbDataMailJointsMethod           'mail_joints_method on the appeals table
      cdbDataPrefulfilledIncentives     'prefulfilled_incentives attr on the fp_applications table
      cdbDataIncentiveProductMinMax     'minimum and maximum values on incentive scheme products
      cdbDataGiftAidSponsorship         'Gift Aid Sponsorship tables
      cdbDataCampaignBudgets            'DUK:Budgets and actuals, appeal.budgeted_count,cost,fixed_cost| segments.budgeted_count | purchase_orders.campaign,appeal,segment | purchase_invoices.campaign,appeal,segment
      cdbDataMembershipProRating        'Membership entitlement rate_to_use etc...
      cdbDataReportItemFormatting       'Report Items item_format
      cdbDataWarehouses                 'New warehouses table
      cdbDataScheduledPayments          'New order_payment_schedule table
      cdbDataWarehouseTransferReason    'stock_warehouse_xfer_reason on stock_movement_controls
      cdbDataCampaignActualIncome       'Campaign, Appeal & Segment Actual Income & Actual Income Date
      cdbDataPayrollGivingPayFrequency  'Payment Frequency attr on gaye_pledges table
      cdbDataAutoEmailDevice            'auto_email attribute on the devices table
      cdbDataGiftAidDecCreatedBy        'Gift Aid Declarations new created_by attr
      cdbDataConfirmSRTransactions      'confirm_sr_transactions on fp_applications
      cdbDataExplorerLinks              'explorer_links table
      cdbDataHeaderPostalSector         'postal sector in contact header table
      cdbDataPayrollGivingPaymentNumber 'payment_number of gaye_pledges table
      cdbDataDevicesWWWAddress          'www_address on devices table
      cdbDataAUNotesMandatory           'notes_mandatory on address_usages table
      cdbDataGiftAidMergeCancellation   'Merge Cancellation Reason of Gift Aid Controls
      cdbDataOnLineCCAuthorisation      'Credit Card Authorisations table
      cdbDataCPDCycle                   'cpd_cycle on contact_cpd_items
      cdbDataPayPlanConvMaintenance     'Pay Plan Conversion Trader App includes Maintenance
      cdbDataLinkToCommunication        'FP Application Link To Communication flag
      cdbCalendarCompany                'Different financial periods for different companies
      cdbDataPrintChequeList            'Print Cheque List flag on Batch Types
      cdbDataIssuedStockJobNumber       'Job number on issued stock table
      cdbDataGiftAidMaxJuniorAge        'Gift Aid - Gift Member Max Junior Age on Membership Controls
      cdbDataDataUpdates                'data_updates table
      cdbDataPayPlanEligibleForGiftAid  'Payment Plan - Eligible For Gift Aid flag
      cdbDataDisplayListItems           'display_list_items table
      cdbDataBankStatementNotes         'Notes attribute on bank_statements table
      cdbDataNominalAccountValidation   'Product / Rate Nominal Accounts tables
      cdbDataSubTopicActivityDuration   'Activity Duration on Sub Topics tables
      cdbDataEventBookingNotes          'Event Bookings Notes
      cdbDataPayPlanDetailCreatedBy     'Payment Plan Detail - Created By
      cdbDataCommunicationNumber        'communication_number on communications
      cdbDataDevicesSequenceNumber      'sequence_number on devices table
      cdbDataBTAWarehouse               'warehouse on batch_transaction_analysis table
      cdbDataUnknownAddresses           'unknown address attributes on contact and organisation groups tables
      cdbDataRelationshipGroups         'relationship groups table
      cdbDataStandardDocumentPrecis     'Précis on standard documents table
      cdbDataCheckPayPlans              'Check Payment Plans Problems
      cdbDataEmailControls              'Email controls table
      cdbDataMailingNotes               'Notes on the mailings table
      cdbDataUserHistory                'History Only on Users table
      cdbDataControlsContactGroup       'contact_group on fp_controls
      cdbDataAgencyAdminFee             '3 attributes for Agency Fee on the GAYE_agencies table.
      cdbDataContactRoleNumber          'Contact role number on the contact roles table
      cdbDataAdultGiftMemEligibleGA     'Adult Gift Memberships Eligible for Gift Aid flag on Membership Types table
      cdbDataEventWaitingListControl    '2 Event flags:waiting_list_control_method & charge_for_waiting
      cdbDataUserHistoryItems           'user_history_items Table
      cdbDataEventGroups                'event_groups table
      CDBDataViewNames                  'View names table
      cdbDataExplorerLinksToolbar       'Show Toolbar on the Explorer Links table
      cdbDataPositionFuntionSeniority   'position_function and position_seniority on contact_positions
      cdbDataEntityGroupSpecificStatus  'contact_group added to statuses table to limit application of statuses to contacts/organisations
      cdbDataCustomFormDeletion         'delete button added to custom forms
      cdbDataMailingHistoryNotes        'per-instance notes for mailings (not the mailing history CODE notes).
      cdbDataControlParameterName       'parameter_name on fp_controls table
      cdbDataActivityGroupsCampaign     'Campaign attribute on Activity Groups table
      cdbDataNumberofCCCAs              'Number of CCCAs attribute on the Contact Header table
      cdbDataPayPlanPackToDonor         'Pack to Donor flag on the Orders table
      cdbDataEventLongDescription       'New long_description attributes on events and sessions tables.
      cdbDataEventClass                 'New event_class attribute on events
      cdbDataBacsSkipRejectedPayment    'BACS Messaging improvements including skip rejected payments
      cdbDataLegaciesReviewReason       'Extra fields on Contact Legacies
      cdbDataPostTaxPayrollGiving       'Post Tax Payroll Giving
      cdbDataMembershipTypeFixedCycle   'Fixed Cycle on Membership Types
      cdbDataMembershipCardIssueNumber  'Membership Card Issue Number on Members
      cdbDataActivityDate               'Activity date on Contact and Organisation categories
      cdbDataDelegateActivities         'Delegate Activities table
      cdbDataProductEligibleGA          'Eligible For Gift Aid Flag on Products
      cdbDataLookupGroups               'Lookup groups table
      cdbDataConfirmedTransStatus       'Status on Confirmed Transactions
      cdbDataAppealType                 'Appeal type on Appeals
      cdbDataCollections                'Collection number on Appeal Collections
      cdbDataSegmentOrgSelectionOptions 'segment_org_selection_options on the appeals table
      cdbDataGiftAidTimeLimits          'Gift Aid tax claim time limits (new attrs on Gift Aid Controls)
      cdbDataEventFinancialAnalysis     'Extra Financial Analysis on Events, Delegates etc.
      cdbDataCreditCardAVSCVV2          'Credit Card AVS/CVV2 security checking
      cdbDataHolidayLets                'Holiday Lets extension of Service Bookings
      cdbDataCustomFinderTab            'Custom finder tab number
      cdbDataJobScheduleUpdateDates     'Update Job Parameter Dates on the Job Schedule table
      cdbDataServiceStartDays           'Service Start Days
      cdbDataPrimaryRelationship        'Primary Relationship on Contact Groups
      cdbDataCustomFormAllowUpdate      'Allow update on custom forms
      cdbDataStandardDocumentClass      'Document class on standard documents
      cdbDataMailmergeHeaderOnReports   'Mail Merge Header and Standard Document on Reports
      cdbDataCustomFormRestrictions     'Restriction attributes on custom form controls
      cdbDataEventTopicNotes            'Notes on Event Topics
      cdbDataEventPIS                   'Event PIS
      cdbDataBranchOwnershipGroup       'Ownership Group the Branches table
      cdbDataContactControlsDevices     'Various devices on the contact controls table
      cdbDataPayPlanPackToMember        'PackToMember flag on Payment Plans
      cdbDataMultipleMerchantRetailNos  'Merchant Retail Number on the Batch Categories table
      cdbDataAddressLines               'Address line 1-4 on the addresses table
      cdbDataCommunicationsUsages       'Communications usages and new attributes on communications table
      cdbDataBatchAnalysisCodes         'Batch Analysis Codes table
      cdbDataMembershipPrices           'Membership Prices table
      cdbDataRegistrationData           'registration data on registered users table
      cdbDataBatchCampaign              'Campaign on batches table
      cdbDataProductCosts               'Product Costs table
      cdbDataPurchaseOrderLink          'Product & Warehouse on purchase_order_details table
      cdbDataGiftMessage                'new_orders.gift_message
      cdbDataPackProducts               'Products - Pack Product attribute
      cdbDataDutchSupport               'all Dutch data structure changes, name formats, Dutch addresses(building number)
      cdbDataContactGroupUsers          'Contact Group Users table
      cdbDataDefaultBankAccount         'DefaultAccount flag on ContactAccounts table
      cdbDataPurchaseOrderManagement    'Frequency on Purchase Orders,Cheque Produced & Reference on P.O. Payments etc.
      cdbDataCCIssueNumber              ' Issue number on contact credit cards table
      cdbDataDashboardItems             'Dashboard data
      cdbDataSecurityQuestion           'Security question and answer on registered users table
      cdbDataBankingDate                'Banking date on batches table
      cdbDataStandardPositions          'Standard positions on standard_positions table
      cdbDataPaymentPlanStartMonth      'StartMonth on Orders table
      cdbDataStockMovementTransactionID 'TransactioID on StockMovements table
      cdbDataIrishGiftAid               'Irish Gift Aid ga_approprite_certificates,ga_certificate_tax_claims etc.
      cdbDataDelegateSessions           'Recording attendance for Event Delegates at Session level.
      cdbDataEventDocumentLinks         'Event Documents table
      cdbDataServiceControlRestrictions 'Service Control Restrictions Table
      cdbDataPPDetailsEffectiveDate     'Payment Plan Details Effective Date
      cdbDataEventFixedPrice            'Fixed Price Events/Family Tickets
      cdbDataHistoryOnlyAccount         'History Only flag on ContactAccounts table
      cdbDataMembershipGroups           'Membership Groups table
      cdbDataBoxAccounting              'Distribution Boxes table.
      cdbDataHistoryOnlyCancReasons     'History Only flag on cancellation_reasons table
      cdbDataMailingError               'error_number attribute in contact_emailings table
      cdbDataPurchaseOrderLineItems     'Lookup table for purchase order line items
      cdbDataThankYouMessage            'ThankYouMessage on ContactFundraisingEvents table
      cdbDataUserHistoryFavourites      'Favourite Item flag on user_history_items table
      cdbDataWebFriendlyUrl             'Friendly URL on WebPages table
      cdbDataJobScheduleSmartClientJob  'Smart Client Job flag on job_schedule table
      cdbDataTransactionOrigins         'Transaction Origin on transaction_origins table
      cdbDataBacsDDCancellationRange    'DD Cancellation Range on bacs_operations table
      cdbDataEmailJobs                  'EMail jobs table
      cdbDataCheetahMail                'Cheetah Mail contact_emailings_links table
      cdbDataMinPriceMandatory          'UpperLowerPriceMandatory on Rates table
      cdbDataAppealExpenditureROI       'total_expenditure and return_on_investment attributes in appeals table
      cdbDataCampaignItemisedCosts      'campaign_itemised_costs, campaign_cost_types tables and total_itemised_cost attribute
      cdbDataAddressDPS                 'Delivery point suffix on address and address data table
      cdbDataBankAccountDepartments     'bank account department lookup table

      cdbDataNINumber                   'NI Number on Contacts table (demo)
      cdbDataWebPageSuppressions        'Suppress header and footer on Web pages
      cdbDataWebPageLoginRequired       'Login Type on Web Pages
      cdbDataWebPagePublished           'Page published on web pages table
      cdbDataMembershipTypeTransitions  'Membership Type Transitions table
      cdbDataAllowAsFirstType           'MembershipType Table
      cdbDataCampaignRoles              'campaign roles lookup table
      cdbDataControlNumberLinks         'Control number links table
      cdbDataDashboardViewNames         'DashboardGeneralView on ViewNames table
      cdbDataMailingSuppressionsNotes   'notes attribute in mailing_suppressions table
      cdbDataEventMultipleAnalysis      'new table event_booking_transactions and new attribute event_multiple_analysis in fp_applications table
      cdbDataEventAdultChildQuantity    'adult_quantity and child_quantity attributes on event_bookings table 
      cdbDataDistributionBoxProcess     'New table DistributionBoxProcesses
      cdbDataHistoryOnlyDistribCodes    'New HistoryOnly attribute on DistributionCodes table
      cdbDataPayrollGivingCreatedByOn   'created by and on for payroll giving pledges
      cdbDataFastDataEntry              'Fast Data Entry support in Smart Client
      cdbDataRegisteredUsersAmendedOn   'amended_on attribute in registered_users table
      cdbDataFPTransactionOrigin        'transaction_origin attribute in fp_applications table
      cdbDataCustomFinderWildcards      'wildcard attributes on custom_finder_controls table
      cdbDataEventPricingMatrix         'Event Pricing Matrix mod
      cdbDataPOPMultiplePayees          'Multiple Payees for Purchase Orders - payee_contact_number and payee_address_number attributes in purchase_order_payments table
      cdbDataEditPanelPages             'EPL Pages for customisation of finders etc.
      cdbDataPrintJobNumber             'print_job_number in invoices table
      cdbDataServiceBookingAnalysis     'new table service_booking_transactions and new attribute service_booking_analysis in fp_applications table
      cdbDataProvisionalInvoiceNumber   'new attribute provisional_invoice_number in the invoices table
      cdbDataAlbacsBankDetails          'new attribute albacs_bank_details in fp_applications table
      cdbDataBoxProcessStatus           'new attribute status in distribution_box_processes table
      cdbDataChequeReissue              'Cheque Reissue support in Smart Client
      cdbDataPositionLinks              'New ContactPositionLinks table
      cdbDataBackClaimYears             'new attribute in gift_aid_controls table
      cdbDataPOPPayByBACS               'Pay by BACS Purchase Order Payment support
      cdbDataControlReadonlyAndPanels        'readonly_item on fp_controls table, Left Panel Item Number & Right Panel Item Number on web_controls table, Suppress left panel & Suppress right panel on Web pages 
      cdbDataDefaultValue               'New Default value column in fp_controls
      cdbDataRgbValueForStatus          'New RGB Value column added in the statuses table  
      cdbDataRgbValueForMemberType      'New RG Value column added in membership_types
      cdbDataRgbValueForActivityValue   'New RG Value column added in activity_values

      cdbDataFundraisingPayments        'Fundraising Payments support in Smart Client
      cdbContactAlerts                  'Display alerts for contacts
      cdbDataOwnerContactNumber         'Owner Contact Number column added in meetings
      cdbDataRelationshipStatus         'new attribute relationship in relationship_statuses table
      cdbEventVenueCapacity             'Event Venue Capacity
      cdbStatusMessage                  'Status Message
      cdbDaysPrior                      'Days Prior To And Days Prior From
      cdbMembershipLookupGroup          'Membership Lookup Group
      cdbEntityAlerts                   'New table for entity alerts
      cdbMembershipStatus               'Membership status
      cdbCPDCycleStatus                 'CPD cycle status
      cdbJobScheduleJobId               'New unique identifier used to abort the job
      cdbCPDPointsNotes                 'New field notes added in contact_cpd_points table and New valid_from ,valid_to ,cpd_poins,points_override,date_mandatory,approved added in cpd_categories table.
      cdbCPDObjective                   'CPD Objective based
      cdbActivityCPDPoints              'New table activity_cpd_points added and two new attribute activity and activity_value added in contact_cpd_points table.      
      cdbBulkEmailAttachments           'New table email job attachments used to save the list of attachments for a bulk email
      cdbDataLabelName                  'new format-label_name attribute in Titles table
      cdbDataAllocationsOnIPH           'New attribute allocation_date on invoice_payment_history table
      cdbRateModifier                   'New attribute Rate_Modifier in product_rates table
      cdbEventStatusColor               'New Attribute rgb_value on event_statuses table
      cdbDataVatRateHistory             'New VatRateHistory table
      cdbPriceIsPercentage              'New field in the rates table
      cdbOutlookIntegration             'New Attribute external_appointment_id on sessions table and New Attribute external_task_id on event_personnel_tasks table
      cdbSalesContactMandatory          'New Attribute sales_contact_mandatory on fp_applications table
      cdbDataMerchantDetails            'New table merchant_details
      cdbDataPayPlansVatExcl            'Support for vat-exclusive Payment Plans
      cdbArchiveCommunications          'New attribute archive on communications table
      cdbProductWebPublish              'New attribute for Product Web Publish
      cdbGrabDetailsOption              'New Attribute last_updated_on on registered_user table
      cdbLongDescription                'New attribute Long Description for event booking option
      cdbEventWebPublish                'New attribute web_publish on events table
      cdbValidFromToForRegisteredUsers  'New attributes valid_from and valid_to on registered_users table
      cdbLoginLockout                   'New attributes to support lockout after a configured number of retries
      cdbWebPublish                     'New attributes to support Web Publish for product
      cdbMembershipTypesWebPublish      'New attributes to support web publish and long description
      cdbCreatedBy                      'New attributes created_by and created_on added on confirmed_transactions
      cdbAdminEmailAddress              'New attributes to support admin email address
      cdbAccessViewName                 'New attribute to support online security
      cdbAddErrorLog                    'New Table 'error_logs' for Error Logging
      cdbCPDWebPublish                  'New attribute to cpd_cycle_types
      cdbFundraisingBusinessType        'New attribute fundraising_business_type on fundraising_requests table
      cdbJournalSelectName              'New attribute select_name on contact_journals table
      cdbBankHolidayDays                'New Table 'bank_holiday_days'
      cdbTelemarketing                  'New Table 'telemarketing_contacts' and 'html_scripts', New attributes in segments, campaigns, sub_topics, communications_log
      cdbAuthorisedTextID               'New attribute credit_card_authorisations to store TextID returned from ProtX server
      cdbLabelNameFormatCode            'New attribute label_name_format_code on contacts table
      cdbTraderInvoicePrintPreview      'New attribute invoice_print_preview_default on fp_applications
      cbdProtxCardType                  'New attribute protx_card_type to add new PROTX column value
      cbdGroupDefaultNameFormat         'New attributes name_format and last_used_id in contact_group, event_groups and organisation_groups 
      cdbStatusReasons                  'New Table status_reasons      
      cdbAdHocPurchaseOrderPayments     'New ad_hoc_payments on purchase order types
      cdbPurchaseOrderAuthorisation     'New authorisation level purchase orders
      cdbPurchaseOrderHistory           'New table purchase_order_history
      cdbBACSePay                       'New attribute bacs_msg_file_format on financial_controls
      cdbDataOrgGroupCustomTables       'New attribute custom_table_names to be used in Clone Organisation process
      cdbLoans                          'New loans table
      cdbCustomFormWebPage              'New attributes custom_form_url and display_browser_toolbar on custom_forms table to support displaying a custom form web page
      cdbRegularPurchaseOrderPayments   'New attribute regular_payments in purchase order types
      cdbRateModifiersUseActivityDate   'New attribute use_activity_date on rate_modifiers
      cdbMembershipPricesOverseas       'New attributes on membership prices table
      cdbCountryAddressFormat           'New field to define address format to format address based on country
      cdbBACSErrorCode                  'New field added to bacs_amendments table and a new table bacs_error_codes is added  
      cdbProductActivityDurationMonths  'New attribute activity_duration_months on products
      cdbPurchaseOrderCurrencyCode      'New attribute currency_code in purchase_orders table
      cdbDataCPDPoints2                 'New attributes in contact_cpd_points table
      cdbDataPaymentPlanSurcharges      'New table payment_plan_surcharges
      cdbDataAppExplorerLinks           'New explorer_location attribute to explorer_links table and new table explorer_link_access_levels is added
      cdbExams                          'New exams tables
      cdbDataCPDPointsContactNumber     'New attribute in contact_cpd_points table
      cdbRateModifiersSequence          'New attribute sequence_number in rate_modifiers table
      cdbMembershipTypeCategories       'New Table 'membership_type_categories'
      cdbAutoCreateCreditCustomer       'New attributes auto_create_credit_customer and credit_category in fp_applications table 
      cdbDataInvoiceWithPayment         'New attributes in credit_sales_controls and fp_applications tables
      cdbLoanInterestRates              'New Loan Interest Rates table
      cdbBankTransactionsImport         'New import_number attribute on bank_transactions table and new data_import_files table
      cbdMembershipEntitlementPriority  'New Priority attribute added to membership_entitlement and cmt_   
      cdbReportUseSsrs                  'New use SSRS attribute on reports table
      cdbUnpostedBatchMsgInPrint        'New attribute unposted Batch message in print in fp_applications table       
      cdbPaymentPlanHistoryDetails      'New table payment_plan_history_details
      cdbExamTraderDefaults             'Exam defaults in Trader
      cdbExamUnitCancellation           'Cancellation fields in exam booking units
      cdbInvoiceAdjustmentStatus        'New attribute adjustment_status on invoices table
      cdbWriteOffMissedPayments         'Payment w/o overdue days in financial controls
      cdbSystemMaintenance              'Support for system maintenance group
      cdbBulkMailer                     'Bulk Mailer interface added
      cdbSuppressionSource              'Source code on suppressions tables
      cdbModifierNextSequence           'Next sequence on rate modifiers added
      cdbResponseChannel                'Response channels table
      cdbEventMinimumBookings           'Minimum Bookings on event booking option
      cdbDelegateSequenceNumber         'Sequence number on delegates table
      cdbActivityDurationDays           'Duration days (and months) on activity (and activity value)
      cdbPasswordExpiry                 'Password expiry date on registered users
      cdbAdvanceCMT                     'Advanced CMT on Membership Controls table.
      cdbAmendmentContactNumber         'Contact Number on amendment history
      cdbTnsHostedPayment               'Added new attributes to support TnsHostedPayment
      cdbExamExemptionModule            'Added new attributes to exam student exemptions
      cdbDataContactAlerts              'Added new attribute to trader application to control the display of contact alerts
      cdbCommunicationsUsage            'Communication usage on the communications table
      cdbConfigVersionInformation       'Versioning information in the config_names table
      cdbMailingHistoryTopic            'Topic SubTopic and Subject on mailing history table
      cdbIbanBicNumbers                 'Added new attributes to store IBAN and BIC in bank_transactions,bacs_amendments,bank_accounts,contact_accounts 
      cdbJobFailureFlag                 'Added new attribute to store the fact that a job failed
      cdbPopPaymentMethod               'Added new tables for Purchase Order Payment Method Changes
      cdbBacsEndToEndId                 'BAcs End to end id required to store the transaction/ batch number read from camt053
      cdbPopBankAccount                 'Added BANK ACCOUNT to pop_payment_methods table   
      cdbPayPlanChangesTermStartDate    'Term Start Date attribute on Payment plan Changes table
      cdbFreeOfChangeBookingOption      'Added free_of_charge to event_booking_options
      cdbLockBranch                     'Lock branch on the members table
      cdbUsePaymentProducedOn           'Use Payment Produced On for Transfer Payments
      cdbPOPaymentReversals             'Added new po_payments_reversals table
      cdbRequiresPoPaymentType          'New field added to purchase_order_types 
      cdbCancelOneYearGiftApm           'Added new attribute cancel_one_year_gift_apm to orders table
      cdbLastExamDate                   'Added new attribute to record the last exam date when the student took the exam
      cdbExamsQualsRegistrationGrading  'New table Exam Assessment Languages
      cdbExamUnitLinkLongDescription    'Added new attribute to record the Long Description for an Exam Unit links(exam_unit_links table)
      cdbWithholdExamResults            'New table exam_results_unrestricted_depts
      cdbDocumentLogLinks               'Added new table 'document_log_links' to record  documents against exam units, centres and centre units     
      cdbAccreditationMakeHistoric      'New field to make accreditation historic
      cdbExamStudyModes                 'New study modes table
      cdbExamLoadResult                 'New field to specify where the Result should be loaded against, for non-session bookings     
      cdbAddressConfirmed
      cdbEmailTemplates                 'Added email header fields for communications log
      cdbHistoricPoPaymentType          'Added new attribute to mark Po Payment type as historic
      cdbTraderInvoicePrintUnpostedBatches    'New attribute invoice_print_unposted_batches on fp_applications
      cdbUserSchemes                    'New table added for User Schemes  
      cdbCountriesIso3166CountryCodes   'New fields on countries table to hold ISO 3166 2 and 3 letter country codes
      cdbOrganisationGroupsViewInContactCard 'New ViewInContactCard field on organisation groups table
      cdbIso3166NumericCountryCode      'New field on countries table to hold ISO 3166 numeric country codes
      cdbBankAccountRGBValue            'New field rgb_value on Bank Accounts table
      '--------------------------------------------------------------------------------------------
      'NOTE INSERT ANY NEW ITEMS BEFORE THIS LINE
      '--------------------------------------------------------------------------------------------
      cdbDataMaxDSInfo
    End Enum


    'Get Data Structure Info what to check
    Private Enum cdbDSConstants
      cdbDSExists = 1
      cdbDSType
    End Enum

    Private Enum cdbDSInfoConstants
      cdbDSInfoTrue = -1
      cdbDSInfoFalse = 99
      cdbDSInfoUnknown = 0
    End Enum

    Private mvDSInfo() As cdbDSInfoConstants

    'The following enum is used by GetControlValue
    Public Enum cdbControlConstants
      ' Financial Controls
      cdbControlDefConVatCat
      cdbControlDefOrgVatCat
      cdbControlCCReason
      cdbControlDDReason
      cdbControlSOReason
      cdbControlOReason
      cdbControlSOActivity
      cdbControlSOActivityValue
      cdbControlDDActivity
      cdbControlDDActivityValue
      cdbControlCCCAActivity
      cdbControlCCCAActivityValue
      cdbControlReverseTransType
      cdbControlExpiredOrdersCancellationReason
      cdbControlDistributorActivity
      cdbControlStockInterface
      cdbControlFirstClaimTransactionType
      cdbControlAutoSODefaultDays
      cdbControlAccountsInterface
      cdbControlMaximumOnPickingList
      cdbControlInMemoriamRelationship
      cdbControlInMemoriamCompRelationship
      cdbControlSODefaultProduct
      cdbControlBankAccount
      cdbControlInAdvanceTransType
      cdbControlInAdvancePaymentMethod
      cdbControlCurrencyCode
      cdbControlDespatchTransactionType
      cdbControlDespatchPaymentMethod
      cdbControlPaymentReason
      cdbControlAdjustmentTransType
      cdbControlBACSNewDDSource
      cdbControlAnonymousContactNumber
      cdbControlAnonCardBankAccount
      cdbControlAnonCardProduct
      cdbControlAnonCardRate
      cdbControlAnonCardSource
      cdbControlAnonCardDistributionCode
      cdbControlAutomaticRenewalDateChangeReason
      cdbControlHoldingContactNumber
      cdbControlRoundingErrorProduct
      cdbControlMerchantRetailNumber
      cdbControlAutoPayClaimDateMethod
      cdbControlApplyIncentiveFreePeriod
      cdbControlDeceasedStatus
      cdbControlAdjustOriginalProductCost
      cdbControlDefaultDDText1
      cdbControlOneOffPPCancelReason
      cdbControlDefaultDDBatchCategory
      cdbControlMailingHistorySearchDays
      cdbControlMailingHistorySearchSegmentType
      cdbControlNewContactSource
      cdbControlExistingContactSource
      cdbControlDefaultProduct
      cdbControlDefaultRate
      cdbControlDefaultFundPayType
      cdbControlDefaultFundStatus
      cdbControlLockFundRequest
      cdbControlSCPURL
      cdbControlSCPAPIVersion
      cdbControlVPCURL
      cdbControlVPCAPIVersion
      cdbControlBacsMsgFileFormat
      cdbLoanCapitalisationDate
      cdbAutoSOAcceptAsFull
      cdbControlPaymentWOOverDueDays
      cdbControlCashBookBatchLimit
      cdbControlBacsUserNumber
      cdbControlOneOffClaimTransactionType
      cdbControlUseRenewalDateForRateMod
      cdbControlAccountValidationType
      cdbControlAccountValidationURL
      cdbControlReceiptPrintStdDocument
      cdbControlOneReversalOnly

      ' Contact Controls
      cdbControlGAStatus
      cdbControlGAMailingSupp
      cdbControlDefMailingSupp
      cdbControlCLIDevice
      cdbControlTYLSupressionExclusionList
      cdbControlParentRelationship
      cdbControlDespatchReason
      cdbControlMaxPermittedDaysActivity
      cdbControlMaxPermittedDaysActivityVal
      cdbControlDaysRemainingActivity
      cdbControlDaysRemainingActivityVal
      cdbControlQualifyingPositionActivity
      cdbControlQualifyingPositionActivityVal
      cdbControlAnonymousContactStatus
      cdbControlStartOfDay
      cdbControlEndOfDay
      cdbControlStartOfLunch
      cdbControlEndOfLunch
      cdbControlDirectDevice
      cdbControlSwitchboardDevice
      cdbControlFaxDevice
      cdbControlMobileDevice
      cdbControlEmailDevice
      cdbControlWebDevice
      cdbControlPositionActivityGroup
      cdbControlPositionRelationshipGroup
      cdbControlClosedOrganisationStatus
      cdbControlMergeUseOldestSource
      cdbControlEmailCaseSensitive
      cdbControlRetainRegUserPasswords

      ' Membership Controls
      cdbControlSponsorActivity
      cdbControlSponsorActivityValue
      cdbControlRealToJointLink
      cdbControlRealToRealLink
      cdbControlBranchParent
      cdbControlBranchProduct
      cdbControlReasonForDespatch
      cdbControlTypeChangeCancelReason
      cdbControlNonAddress
      cdbControlOverAgeCancelReason
      cdbControlNonPaymentCancelReason
      cdbControlOneYearGiftedGroupReason
      cdbControlMembershipSalesGroup
      cdbControlOneYearGiftsAutoReason
      cdbControlGiftMemberMaxJuniorAge
      cdbControlCMTProportionBalance
      cdbControlAutoPayAdvancePeriod
      cdbControlMemberOrganisationGroup
      cdbControlCMTMakeRefund
      cdbControlFMTEffectiveDays
      cdbControlAdvancedCMT
      cdbControlRemoveZeroBalancePpdLines
      cdbControlAddMemberCurrentAddress
      cdbControlMoveDDMemberCancelReason

      ' Covenant Controls
      cdbControlCVActivity
      cdbControlCVActivityValue
      cdbControlCVGiftAidMinimum
      cdbControlCVReasonForDespatch
      cdbControlDeedReceivedProduct
      cdbControlDeedReceivedRate
      cdbControlCVMinimumCovenantPeriod

      'Credit Sales Controls
      cdbControlCSTransType
      cdbControlCSPayMethod
      cdbControlCSBankAccount
      cdbControlCSCreditTransType
      cdbControlCSSource
      cdbControlCSOverPayProduct
      cdbControlCSOverPayRate
      cdbControlCSUnderPayProduct
      cdbControlCSUnderPayRate

      ' Marketing Controls
      cdbControlsDerivedSuppression
      cdbControlJointSuppression
      cdbControlDerivedToJointLink
      cdbControlDerivedToDerivedLink
      cdbControlEarliestDonation
      cdbControlMinimumAge
      cdbControlDepartment
      cdbControlLastContactNumber
      cdbControlLastMembershipNumber
      cdbControlLastBatchNumber
      cdbControlTransactionType
      cdbControlVATRate
      cdbControlIncludeHistoric
      cdbControlCriteriaSet
      cdbControlCollectionsRegionType
      cdbControlDefaultCollectorStatus
      cdbControlNegativeAdjustmentProduct
      cdbControlNegativeAdjustmentRate
      cdbControlPositiveAdjustmentProduct
      cdbControlPositiveAdjustmentRate

      'Mail sort
      cdbControlMaxBagWeight

      'Legacy
      cdbControlLGActivity
      cdbControlLGActivityValue
      cdbControlLGJointRelationship
      cdbControlLGAssetActivity
      cdbControlLGConditionalBequestSubType
      cdbControlLGResidualBequestType
      cdbControlLGSpecificBequestType
      cdbControlLGRelationshipList

      'Payroll Giving (Give as You Earn)
      cdbControlGAYEDonorProduct
      cdbControlGAYEDonorRate
      cdbControlGAYEEmployerProduct
      cdbControlGAYEEmployerRate
      cdbControlGAYEGovernmentProduct
      cdbControlGAYEGovernmentRate
      cdbControlGAYEAdminFeeProduct
      cdbControlGAYEAdminFeeRate
      cdbControlGAYEActivity
      cdbControlGAYEActivityValue
      cdbControlGAYEGovernmentPercentage
      cdbControlGAYESource
      cdbControlGAYENiDataSource
      cdbControlGAYEDistributionCode
      cdbControlGAYEAgencyRelationship
      cdbControlGAYEPayrollRelationship
      cdbControlPostTaxPGDonorProduct
      cdbControlPostTaxPGDonorRate
      cdbControlPostTaxPGEmployerProduct
      cdbControlPostTaxPGEmployerRate
      cdbControlPostTaxPGPledgeSource
      cdbControlPostTaxPGDistributionCode
      cdbControlPostTaxPGEmployerActivity
      cdbControlPostTaxPGEmployerActivityValue
      cdbControlPostTaxPGEmplrPayrollRelationship
      cdbControlPostTaxPGPledgeActivity
      cdbControlPostTaxPGPledgeActivityValue
      cdbControlPreTaxOtherMatchedProduct
      cdbControlPreTaxOtherMatchedRate

      'Stock Movement
      cdbControlStockReasonInitial
      cdbControlStockReasonSale
      cdbControlStockReasonReversal
      cdbControlStockReasonShortFall
      cdbControlStockReasonBackOrder
      cdbControlStockReasonAwaitBackOrder
      cdbControlStockProcessLock
      cdbControlStockImportReason
      cdbControlStockPackTransferReason
      cdbControlStockWarehouseTransferReason

      'Gift Aid
      cdbControlGAMergeCancellationReason
      cdbControlGAClaimFileFormat
      cdbControlGACharityTaxStatus
      cdbControlGAAccountingPeriodStart
      cdbControlGAMinimumAnnualDonation
      cdbControlGATaxYearStart
      cdbControlGASubmitterContact
      cdbControlGAAdjustmentText

      'Email
      cdbControlEmailUseHeaderTemplate
      cdbControlEmailHeaderTemplate
      cdbControlEmailForceSmtpAddress

      'Other
      cdbControlCoreStockControl

      'EventControls
      cdbControlEventpisBankAccount
      cdbControlEventPISPerDelegate

      'ExamControls
      cdbControlExamGeographicalRegionType
      cdbControlExamExemptionCompany
      cdbControlExamExemptionSource
      cdbControlExamExemptionCreditCategory
      cdbControlExamExemptionGrade
      cdbControlExamExemptionResult
      cdbControlExamExemptionOrgActivity
      cdbControlExamExemptionOrgActivityValue
      cdbControlExamRecordGradeChangeHistory
      cdbControlExamCentreAccreditation
      cdbControlExamUnitAccreditation
      cdbControlExamCentreUnitAccreditation
      cdbControlExamGradingMethod
      cdbControlExamLoadResult
      cdbControlExamCertNumberPrefix
      cdbControlExamCertNumber
      cdbControlExamCertNumberSuffix

      'Bulk Mailer Controls
      cdbControlBulkMailerLoginId
      cdbControlBulkMailerPassword

      'Purchase Order Controls
      cdbControlPopDefPaymentMethod
      cdbControlUsePaymentProducedOn

      '--------------------------------------------------------------------------------------------
      'NOTE INSERT ANY NEW ITEMS BEFORE THIS LINE
      '--------------------------------------------------------------------------------------------
      cdbControlMaxItem
    End Enum

    Public Enum SetCurrentOutputDirectoryTypes
      scodtMailing
      scodtOutput
    End Enum


    'Journal information
    Private mvJournalInitialised As Boolean
    Private mvOptionJournal As Boolean
    Private mvJournalList As CollectionList(Of JournalOperation)

    'Get Control number
    Private mvCheckDigitsRead As Boolean
    Private mvCheckDigitsOnNumbers As String
    Private mvCheckDigitMethod As String

    'Get Control Value
    Private mvFinancialControls As Boolean
    Private mvContactControls As Boolean
    Private mvMembershipControls As Boolean
    Private mvCovenantControls As Boolean
    Private mvCreditSalesControls As Boolean
    Private mvMarketingControls As Boolean
    Private mvMailsortControls As Boolean
    Private mvLegacyControls As Boolean
    Private mvGAYEControls As Boolean
    Private mvStockMovementControls As Boolean
    Private mvGiftAidControls As Boolean
    Private mvEmailControls As Boolean
    Private mvEventControls As Boolean
    Private mvExamControls As Boolean
    Private mvBulkMailerControls As Boolean
    Private mvPurchaseOrderControls As Boolean
    Private mvControls() As String

    Private mvNameStyle As String = ""
    Private mvDefaultCountry As String = ""
    Private mvUKCountries As CDBParameters
    Private mvConfigs As SortedList

    Private mvEntityGroups As EntityGroups
    Private mvPostcoder As Postcoder
    Private mvJuniorAgeLimit As Integer

    Private mvINISection As String
    Private mvConnectionString As String
    Private mvClientCode As String
    Private mvDescription As String
    Private mvOptimise As String
    Private mvSmartClient As Boolean
    Private mvDBUpdated As Boolean
    Private mvLoginsDB As String
    Private mvSQLLogMode As CDBConnection.SQLLoggingModes = CDBConnection.SQLLoggingModes.None
    Private mvRDBMSType As CDBConnection.RDBMSTypes = CDBConnection.RDBMSTypes.rdbmsUnknown
    Private mvSQLLogQueueName As String
    Private mvInitialiseLocation As String
    Private mvInitialisingDatabase As Boolean

    Private mvAuditStyleChecked As Boolean
    Private mvAuditStyle As AuditStyleTypes
    Private mvConnection As CDBConnection
    Private mvConnections As CollectionList(Of CDBConnection)
    Private mvUser As CDBUser
    Private mvCountryIbanNumbers As CollectionList(Of CountryIbanNumber)
    Private mvCountries As CollectionList(Of Country)
    Private mvCachedConfigScope As New Dictionary(Of String, Config.ConfigNameScope)

    Public Sub New(ByVal pINISection As String, ByVal pLogname As String, ByVal pPassWord As String, ByVal pDatabaseLogname As String, ByVal pUserID As String, ByVal pNetworkLogname As String)
      Init(pINISection, pLogname, pPassWord, pDatabaseLogname, pUserID, pNetworkLogname)
    End Sub
    Public Sub New(ByVal pINISection As String, ByVal pLogname As String, ByVal pPassWord As String, ByVal pDatabaseLogname As String, ByVal pUserID As String)
      Init(pINISection, pLogname, pPassWord, pDatabaseLogname, pUserID, "")
    End Sub

    Public Sub New(ByVal pINISection As String, ByVal pLogname As String, ByVal pPassword As String)
      Init(pINISection, pLogname, pPassword, "", "", "")
    End Sub

    Private Sub Init(ByVal pINISection As String, ByVal pLogname As String, ByVal pPassword As String, ByVal pDatabaseLogname As String, ByVal pUserID As String, ByVal pNetworkLogname As String)

      ReDim mvControls(cdbControlConstants.cdbControlMaxItem)
      For vIndex As Integer = 0 To cdbControlConstants.cdbControlMaxItem - 1
        mvControls(vIndex) = ""
      Next
      ReDim mvDSInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMaxDSInfo)
      If NfpConfigrationManager.Databases IsNot Nothing AndAlso NfpConfigrationManager.Databases(pINISection) IsNot Nothing Then
        mvClientCode = NfpConfigrationManager.Databases(pINISection).ClientCode
        mvDescription = NfpConfigrationManager.Databases(pINISection).Description
        mvConnectionString = ConfigurationManager.ConnectionStrings(NfpConfigrationManager.Databases(pINISection).ConnectionStringName).ConnectionString
        If String.IsNullOrWhiteSpace(mvConnectionString) Then
          Throw New ApplicationException(String.Format("NFP database {0} does not have a valid connection string defined.", pINISection))
        End If
        If String.IsNullOrWhiteSpace(ConfigurationManager.ConnectionStrings(NfpConfigrationManager.Databases(pINISection).ConnectionStringName).ProviderName) OrElse
           ConfigurationManager.ConnectionStrings(NfpConfigrationManager.Databases(pINISection).ConnectionStringName).ProviderName = "System.Data.SqlClient" Then
          mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer
        ElseIf ConfigurationManager.ConnectionStrings(NfpConfigrationManager.Databases(pINISection).ConnectionStringName).ProviderName = "System.Data.OracleClient" Then
          mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle
        End If
        mvInitialiseLocation = NfpConfigrationManager.Databases(pINISection).InitialiseDatabaseFrom
        mvSQLLogQueueName = NfpConfigrationManager.Databases(pINISection).SqlLogQueueName
        mvSQLLogMode = NfpConfigrationManager.Databases(pINISection).SqlLogging
        mvOptimise = String.Empty
        mvDBUpdated = False
        mvLoginsDB = String.Empty
      Else
        Dim vReader As New INIReader
        vReader.Section = pINISection
        Dim vConnectString As String = ""
        mvSQLLogQueueName = vReader.ReadString("SQLLOGQUEUENAME")
        mvInitialiseLocation = vReader.ReadString("INITIALISEDATABASE")
        If pINISection.Length > 0 Then
          If pINISection.StartsWith("DSN=") Then
            mvConnectionString = GetConnectionStringFromDSN(pINISection)
            If mvConnectionString.Length = 0 Then RaiseError(DataAccessErrors.daeEntryInINIFileMissing, String.Format("{0} - {1}", pINISection, "RDBMSTYPE")) 'Entries in the '%1' section of the INI file could not be found\r\n\r\nCannot continue!
            mvClientCode = "CARE"
            mvDescription = ""
            mvOptimise = ""
            mvDBUpdated = True
            mvLoginsDB = ""
            mvConnection = CDBConnection.GetCDBConnection(mvRDBMSType, mvSQLLogMode, mvSQLLogQueueName)
            mvConnection.OpenConnection(mvConnectionString, pLogname, pPassword, False, False)
            mvConnection.PopulateUnicode()
          Else
            Dim vRDBMSType As String = vReader.ReadString("rdbmstype").ToUpper
            Select Case vRDBMSType
              Case "SQLSERVER"
                mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer
                mvConnectionString = vReader.ReadString("connectionstring")
              Case "ORACLE"
                mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle
                mvConnectionString = vReader.ReadString("connectionstring")
              Case Else
                Dim vDBType As String = vReader.ReadString("type").ToUpper
                Select Case vDBType
                  Case "ODBC"
                    vConnectString = vReader.ReadString("connect")
                  Case ""
                    RaiseError(DataAccessErrors.daeEntryInINIFileMissing, pINISection) 'Entries in the '%1' section of the INI file could not be found\r\n\r\nCannot continue!
                End Select
                mvConnectionString = GetConnectionStringFromDSN(vConnectString)
            End Select
            If mvConnectionString.Length = 0 Then RaiseError(DataAccessErrors.daeEntryInINIFileMissing, String.Format("{0} - {1}", pINISection, "RDBMSTYPE")) 'Entries in the '%1' section of the INI file could not be found\r\n\r\nCannot continue!
            mvClientCode = vReader.ReadString("client", "CARE").ToUpper
            mvDescription = vReader.ReadString("description", "Unknown")
            mvOptimise = vReader.ReadString("optimise")
            mvDBUpdated = mvOptimise.Contains("DBUPDATED")
            mvLoginsDB = vReader.ReadString("logins")
          End If
        End If
        mvSQLLogMode = CType(vReader.ReadInteger("SQLLOGGING"), CDBConnection.SQLLoggingModes)
      End If
      mvINISection = pINISection
      mvUser = New CDBUser(Me, pLogname, pPassword, pDatabaseLogname, pUserID, pNetworkLogname)
      If NfpConfigrationManager.ExtensionAssemblies IsNot Nothing Then
        Advanced.Extensibility.ExtensionManager.Initialise(NfpConfigrationManager.ExtensionAssemblies)
      End If
      If mvConnectionString.Length = 0 Then RaiseError(DataAccessErrors.daeNoConnectString, pINISection)
    End Sub

    Public Property InitialisingDatabase As Boolean
      Get
        Return mvInitialisingDatabase
      End Get
      Set(ByVal pValue As Boolean)
        If pValue And mvInitialiseLocation.Length > 0 Then
          mvInitialisingDatabase = True
        Else
          mvInitialisingDatabase = False
        End If
      End Set
    End Property

    Public ReadOnly Property InitialiseLocation As String
      Get
        Return mvInitialiseLocation
      End Get
    End Property

    Public Sub CreateDatabase()
      mvConnection = Nothing
      Dim vConnectionString As String = mvConnectionString
      Dim vItems() As String = vConnectionString.Split(";"c)
      Dim vDBName As String = ""
      Dim vUserID As String = ""
      Dim vPassword As String = ""
      For Each vItem As String In vItems
        If vItem.ToLower.StartsWith("initial catalog") Then
          Dim vPos As Integer = vItem.IndexOf("=")
          vDBName = vItem.Substring(vPos + 1)
        ElseIf vItem.ToLower.StartsWith("database") Then
          Dim vPos As Integer = vItem.IndexOf("=")
          vDBName = vItem.Substring(vPos + 1)
        ElseIf vItem.ToLower.StartsWith("user id") Then
          Dim vPos As Integer = vItem.IndexOf("=")
          vUserID = vItem.Substring(vPos + 1)
        ElseIf vItem.ToLower.StartsWith("password") Then
          Dim vPos As Integer = vItem.IndexOf("=")
          vPassword = vItem.Substring(vPos + 1)
        End If
      Next
      If mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        'First find the existing database connection string and replace the database with master then try to make a connection
        vConnectionString = vConnectionString.Replace(vDBName, "master")
        Dim vConnection As CDBConnection = CDBConnection.GetCDBConnection(mvRDBMSType, mvSQLLogMode, mvSQLLogQueueName)
        vConnection.OpenConnection(vConnectionString, mvUser.DatabaseLogname, mvUser.Password, False, False)
        Dim vDBExists As Boolean
        Dim vRecordSet As CDBRecordSet = vConnection.GetRecordSet("sp_helpdb", CDBConnection.RecordSetOptions.NoDataTable)
        While vRecordSet.Fetch
          If vRecordSet.Fields("name").Value.ToLower = vDBName.ToLower Then
            vDBExists = True
            Exit While
          End If
        End While
        vRecordSet.CloseRecordSet()
        If Not vDBExists Then vConnection.ExecuteSQL("CREATE DATABASE " & vDBName)
        If Not String.IsNullOrWhiteSpace(vUserID) Then
          Dim vExists As Boolean = False
          vRecordSet = vConnection.GetRecordSet("sp_helplogins @LoginNamePattern = '" & vUserID & "'", CDBConnection.RecordSetOptions.NoDataTable)
          If vRecordSet.Fetch Then vExists = True
          vRecordSet.CloseRecordSet()
          If Not vExists Then vConnection.ExecuteSQL("sp_addlogin @loginame = '" & vUserID & "', @passwd = '" & vPassword & "'")
          vConnection.ExecuteSQL("sp_addsrvrolemember @loginame = '" & vUserID & "', @rolename = 'sysadmin'")
        End If
        vConnection.CloseConnection()
      End If
      mvConnection = CDBConnection.GetCDBConnection(mvRDBMSType, mvSQLLogMode, mvSQLLogQueueName)
      mvConnection.OpenConnection(mvConnectionString, vUserID, vPassword, False, False)
      If mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        mvConnection.ExecuteSQL("sp_grantdbaccess @loginame = '" & vUserID & "'", CDBConnection.cdbExecuteConstants.sqlIgnoreError)
      End If
      mvConnection.PopulateUnicode()
      mvUser.InitForNewDatabase(vUserID)
    End Sub

    Public Function CheckControlNumbers(ByVal pEnv As CDBEnvironment, Optional ByVal pInList As String = "") As String
      Dim vControlNumberError As String
      Dim vWhereFields As New CDBFields
      Dim vCheckDigitMethod As String
      Dim vCheckDigitsOnNumbers As String
      Dim vAnsiJoins As New AnsiJoins()
      Dim vDataTable As DataTable
      Dim vNumber As Integer = 0
      Dim vType As String = String.Empty
      Dim vDesc As String = String.Empty
      Dim vTable As String = String.Empty
      Dim vAttr As String = String.Empty
      Dim vCheckNumber As Integer = 0
      Dim vMessage As New StringBuilder

      vWhereFields.Add("config_name", "check_digit_method")
      vCheckDigitMethod = New SQLStatement(pEnv.Connection, "config_value", "config", vWhereFields).GetValue
      vWhereFields("config_name").Value = "check_digits_on_numbers"
      vCheckDigitsOnNumbers = New SQLStatement(pEnv.Connection, "config_value", "config", vWhereFields).GetValue
      vWhereFields.Clear()

      vAnsiJoins.Add("control_number_checks cnc", "cn.control_number_type", "cnc.control_number_type", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      If pInList.Length > 0 Then vWhereFields.Add("cn.control_number_type", pInList, CDBField.FieldWhereOperators.fwoIn)
      vWhereFields.Add("cnc.table_name", "_none", CDBField.FieldWhereOperators.fwoNotEqual)
      vDataTable = New SQLStatement(pEnv.Connection, "control_number,cn.control_number_type,control_number_type_desc,table_name,attribute_name", "control_numbers cn", vWhereFields, "cn.control_number_type", vAnsiJoins).GetDataTable

      Dim vInvoiceControls As New List(Of String) From {"'I'", "'PR'"}
      Dim vHasInvoiceControls As Boolean = False

      If vDataTable IsNot Nothing Then
        For Each vRow As DataRow In vDataTable.Rows
          vNumber = IntegerValue(vRow.Item(0).ToString)
          vType = vRow.Item("control_number_type").ToString
          vDesc = vRow.Item("control_number_type_desc").ToString
          vTable = vRow.Item("table_name").ToString
          vAttr = vRow.Item("attribute_name").ToString
          vCheckNumber = vNumber
          If vInvoiceControls.Contains(vType) Then vHasInvoiceControls = True
          If vCheckDigitsOnNumbers.Contains("|" & vType & "|") Then
            vCheckNumber = GenerateCheckDigit(vCheckDigitMethod, vNumber)
          End If
          If Len(vTable) = 0 Or Len(vAttr) = 0 Then
            Select Case vType
              Case "AR", "LG", "E", "O", "ZZ"
                'Ignore
              Case Else
                vMessage.AppendLine(String.Format("WARNING Unknown Control Number {0} ({1}): ", vType, vDesc))
            End Select
          End If
          If vTable.Length > 0 Then
            vControlNumberError = CheckOneControlNumber(vType, vCheckNumber, vDesc, vTable, vAttr)
            If vControlNumberError.Length > 0 Then
              vMessage.AppendLine(String.Format("WARNING {0}: ", vControlNumberError))
            End If
          End If
        Next
      End If
      If vHasInvoiceControls Then vMessage.AppendLine(CheckInvoiceControlNumbers(vDataTable))
      Return vMessage.ToString
    End Function

    Public Function CheckOneControlNumber(ByVal pControlNumberType As String, ByVal pControlNumber As Integer, ByVal pDesc As String, ByVal pTable As String, ByVal pAttr As String, Optional ByVal pUpdate As Boolean = False) As String
      Dim vMaxNumber As Integer
      Dim vMaxLength As Integer
      Select Case pTable
        Case "fp_applications"      'Special case because the attribute is a string
          vMaxLength = New SQLStatement(Connection, String.Format("MAX({0})", Connection.DBLength(pAttr)), pTable).GetIntegerValue
          vMaxNumber = New SQLStatement(Connection, "MAX(" & pAttr & ") AS maxval", pTable, New CDBField(Connection.DBLength(pAttr), vMaxLength)).GetIntegerValue
        Case "invoices"
          Dim vWherefields As New CDBFields
          If GetDataStructureInfo(cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber) AndAlso Not pAttr = "provisional_invoice_number" Then
            vWherefields.Add("provisional_invoice_number", CDBField.FieldTypes.cftLong, "")
            'Jira 380: Exclude Sales Ledger Cash 'C' type invoices where the invoice number could be set 
            'from the provisional invoice number ('PR') control number range
            vWherefields.Add("record_type", CDBField.FieldTypes.cftCharacter, "C", CDBField.FieldWhereOperators.fwoNotEqual)
          End If
          vMaxNumber = New SQLStatement(Connection, String.Format("MAX({0}) AS maxval", pAttr), pTable, vWherefields).GetIntegerValue
        Case Else
          vMaxNumber = New SQLStatement(Connection, String.Format("MAX({0}) AS maxval", pAttr), pTable).GetIntegerValue
      End Select
      If vMaxNumber >= pControlNumber And vMaxNumber <> 0 Then
        If pUpdate Then
          Dim vWhereFields As New CDBFields
          Dim vUpdateFields As New CDBFields
          vWhereFields.Add("control_number_type", pControlNumberType)
          vUpdateFields.Add("control_number", vMaxNumber + 1)
          Connection.UpdateRecords("control_numbers", vUpdateFields, vWhereFields, False)
          Return ""
        Else
          Dim vControlNumberError As New StringBuilder("Control Number for '" + pDesc + "' is Currently " + pControlNumber.ToString() + " but should be ")
          'Jira 380: Error for Provisional Invoice Number and Invoices control number clash
          If pControlNumberType = "PR" AndAlso pAttr = "invoice_number" Then vControlNumberError.Append("greater than the Invoices Control Number ")
          vControlNumberError.Append((vMaxNumber + 1).ToString)
          Return vControlNumberError.ToString
        End If
      Else
        Return ""
      End If
    End Function
    ''' <summary>
    ''' Check to see if Provisional Invoice control number is more than 10,000 greater than the invoice control number. Warn if not.
    ''' </summary>
    ''' <param name="pDataTable">datatable of table control_numbers</param>
    ''' <returns>Appropriate Warning message or an empty string</returns>
    ''' <remarks></remarks>
    Private Function CheckInvoiceControlNumbers(pDataTable As DataTable) As String
      Dim vControlNumberError As New StringBuilder
      Dim vInvoiceControlRecord As DataRow() = pDataTable.Select("control_number_type='I'")
      Dim vProvisionalInvoiceControlRecord As DataRow() = pDataTable.Select("control_number_type='PR'")
      If vInvoiceControlRecord.Length > 0 And vProvisionalInvoiceControlRecord.Length > 0 Then
        Dim vInvoiceControlValue As Integer = IntegerValue(vInvoiceControlRecord(0)("control_number").ToString())
        Dim vProvisionalInvoiceControlValue As Integer = IntegerValue(vProvisionalInvoiceControlRecord(0)("control_number").ToString())
        If vProvisionalInvoiceControlValue < vInvoiceControlValue + 10000 Then
          vControlNumberError = New StringBuilder("WARNING Control Number for 'Provisional Invoice Number' is currently " & vProvisionalInvoiceControlValue.ToString() & " but should be greater than ")
          vControlNumberError.Append(vInvoiceControlValue + 10000)
        End If
      Else
        If vInvoiceControlRecord.Length = 0 Then
          vControlNumberError = New StringBuilder("WARNING Control Number Invoices not present")
        End If
        If vProvisionalInvoiceControlRecord.Length = 0 Then
          If vControlNumberError.Length > 0 Then
            vControlNumberError.Append(", and Control Number Provisional Invoice not present.")
          Else
            vControlNumberError = New StringBuilder("WARNING Control Number Provisional Invoice not present")
          End If
        End If
      End If
      Return vControlNumberError.ToString()
    End Function


    Public Function GetCopyEnvironment() As CDBEnvironment
      Dim vEnv As New CDBEnvironment(mvINISection, mvUser.Logname, mvUser.Password, mvUser.DatabaseLogname, mvUser.UserID)
      Dim vConn As CDBConnection = vEnv.Connection
      Return vEnv
    End Function

    Private Function GenerateCheckDigit(ByVal pCheckDigitMethod As String, ByVal pNumber As Integer) As Integer
      'Will receive a number as a parameter and will use a config option to determine if the method to use to calculate the check digit.
      'Data Import assumes check digit will always be an ADDITIONAL digit at the end of the control number. (ajh 6/7/99)
      Dim vNumber As String = pNumber.ToString
      Select Case pCheckDigitMethod
        Case "M11+4"
          Dim vLen As Integer = vNumber.Length
          Dim vCounter As Integer = 1
          Dim vAccumulator As Integer = 0
          Do While (vLen > 0 AndAlso vCounter < 6)
            vAccumulator += (CInt(vNumber.Substring(vLen - 1, 1)) * (vCounter + 1))
            vLen = vLen - 1
            vCounter = vCounter + 1
          Loop
          If vLen > 0 Then vAccumulator += (CInt(vNumber.Substring(0, vLen)) * 7)
          Dim vCheckDigit As Integer = CInt(11 - (vAccumulator Mod 11))
          If vCheckDigit = 10 Then vCheckDigit = 4
          If vCheckDigit = 11 Then vCheckDigit = 0
          Return pNumber * 10 + vCheckDigit
        Case Else
          Return pNumber 'Leave the number as it was
      End Select
    End Function


    Private Function GetConnectionStringFromDSN(ByVal pDSNString As String) As String
      Dim vConnectionString As New StringBuilder
      Dim vDSN As String = ""
      Dim vUID As String = ""
      Dim vPWD As String = ""
      Dim vItems() As String = pDSNString.Split(";"c)
      For vIndex As Integer = 0 To vItems.Length - 1
        If vItems(vIndex).StartsWith("DSN=", StringComparison.CurrentCultureIgnoreCase) Then
          vDSN = vItems(vIndex).Substring(4)
        ElseIf vItems(vIndex).StartsWith("UID=", StringComparison.CurrentCultureIgnoreCase) Then
          vUID = vItems(vIndex).Substring(4)
        ElseIf vItems(vIndex).StartsWith("PWD=", StringComparison.CurrentCultureIgnoreCase) Then
          vPWD = vItems(vIndex).Substring(4)
        End If
      Next
      If vDSN.Length > 0 Then
        Dim vServer As String = ""
        Dim vServerName As String = ""
        Dim vDatabase As String = ""
        Dim vTrusted As Boolean
        Dim vKey As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(String.Format("Software\\ODBC\\ODBC.INI\\{0}", vDSN))
        If vKey IsNot Nothing Then
          For Each vName As String In vKey.GetValueNames
            Select Case vName.ToLower
              Case "server"                 'SQL Server
                vServer = vKey.GetValue(vName).ToString
              Case "servername"             'ORACLE
                vServer = vKey.GetValue(vName).ToString
              Case "database"               'SQL Server
                vDatabase = vKey.GetValue(vName).ToString
              Case "trusted_connection"     'SQL Server
                vTrusted = vKey.GetValue(vName).ToString.ToLower = "yes"
              Case "driver"
                Dim vDriver As String = vKey.GetValue(vName).ToString.ToLower
                If vDriver.Contains("sqlsrv32.dll") OrElse vDriver.Contains("sqlncli") Then mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer
            End Select
          Next
        End If
        If mvRDBMSType <> CDBConnection.RDBMSTypes.rdbmsSqlServer Then
          If vServer.Length > 0 AndAlso vDatabase.Length > 0 Then
            mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer
          Else
            If vServerName.Length > 0 AndAlso vServer.Length = 0 Then vServer = vServerName
            mvRDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle
          End If
        Else
          If vServer.Length = 0 Then vServer = vDSN
        End If
        With vConnectionString
          If vServer.Length > 0 Then
            .Append("Data Source=")
            .Append(vServer)
            If vDatabase.Length > 0 Then
              .Append(";")
              .Append("Initial Catalog=")
              .Append(vDatabase)
            End If
            If vTrusted Then
              .Append(";Integrated Security=True")
            End If
            If vUID.Length > 0 AndAlso vPWD.Length > 0 Then
              .Append(";")
              .Append("user id=")
              .Append(vUID)
              .Append(";")
              .Append("password=")
              .Append(vPWD)
            End If
          End If
        End With
      End If
      Return vConnectionString.ToString
    End Function

    Public ReadOnly Property SQLLoggingMode() As CDBConnection.SQLLoggingModes
      Get
        Return mvSQLLogMode
      End Get
    End Property

    Public Property SmartClient() As Boolean
      Get
        Return mvSmartClient
      End Get
      Set(ByVal pValue As Boolean)
        mvSmartClient = pValue
      End Set
    End Property

    Public Shared ReadOnly Property VersionNumber() As String
      Get
        Return String.Format("{0:0}.{1:0}", System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major, System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor)
      End Get
    End Property

    Public Shared ReadOnly Property BuildNumber() As Integer
      Get
        Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build
      End Get
    End Property

    Public ReadOnly Property ClientCode() As String
      Get
        Return mvClientCode
      End Get
    End Property

    Public ReadOnly Property Description() As String
      Get
        Return mvDescription
      End Get
    End Property

    Public ReadOnly Property INISection() As String
      Get
        Return mvINISection
      End Get
    End Property

    Public ReadOnly Property Parameters() As String
      Get
        Return mvINISection & " " & mvClientCode & " " & mvUser.Logname & " " & New String("*"c, mvUser.Password.Length)
      End Get
    End Property

    Public Function GetLoginConnection() As CDBConnection
      Return GetDefaultConnection(True)
    End Function

    Public Function Connection() As CDBConnection
      Return GetDefaultConnection(False)
    End Function

    Public Sub ResetConnection()
      If mvConnection IsNot Nothing Then
        Try
          mvConnection.CloseConnection()
        Finally
          mvConnection = Nothing
        End Try
      End If
    End Sub

    Private Function GetDefaultConnection(ByVal pOverrideConnectionString As Boolean) As CDBConnection
      If mvConnection Is Nothing Then
        mvConnection = CDBConnection.GetCDBConnection(mvRDBMSType, mvSQLLogMode, mvSQLLogQueueName)
        mvConnection.OpenConnection(mvConnectionString, mvUser.DatabaseLogname, mvUser.Password, False, pOverrideConnectionString)
        mvConnection.PopulateUnicode()
        If Not InitialisingDatabase Then
          mvUser.InitWithLogname()
          If mvUser.Existing = False Then RaiseError(DataAccessErrors.daeNotValidAccountOrPassword)
        End If
      End If
      Return mvConnection
    End Function

    Public Function GetConnection(ByRef pName As String, Optional ByVal pConnectString As String = "") As CDBConnection
      Dim vPersonal As Boolean
      Dim vConnection As CDBConnection

      Select Case pName
        Case "DATA" 'Used to set lookup to Default but cannot now due to db upgrade process
          pName = "Default"
        Case "PERSONAL"
          vPersonal = True
      End Select
      'TODO VB6 CONVERSION - Need to sort this out for additional connections
      If mvConnections Is Nothing Then
        mvConnections = New CollectionList(Of CDBConnection)
        If mvConnection IsNot Nothing Then mvConnections.Add("Default", mvConnection)
      End If
      If Not mvConnections.ContainsKey(pName) Then
        If pName = "CONTROLNUMBER" Then pConnectString = mvConnectionString
        If pConnectString.Length = 0 Then
          If vPersonal Then
            Dim vSections() As String
            vSections = Split(mvConnectionString, ";")
            For Each vSection As String In vSections
              If vSection.StartsWith("DSN=") Then
                vSection = "DSN=" & mvUser.PersonalDb
              End If
            Next
            pConnectString = Join(vSections, ";")
          Else
            'Connection information not found in INI file - so try db_info
            If mvConnection.ConnectionOpen Then
              Dim vRecordSet As CDBRecordSet = mvConnection.GetRecordSet("SELECT db_type, connect_string FROM db_info WHERE db_name = '" & pName & "'")
              If vRecordSet.Fetch() Then
                pConnectString = vRecordSet.Fields(2).Value
              End If
              vRecordSet.CloseRecordSet()
            End If
          End If
        End If
        If pConnectString.Length = 0 Then RaiseError(DataAccessErrors.daeNoConnectString, pName)
        vConnection = CDBConnection.GetCDBConnection(mvRDBMSType, mvSQLLogMode, mvSQLLogQueueName)
        mvConnections.Add(pName, vConnection)
        vConnection.OpenConnection(mvConnectionString, mvUser.DatabaseLogname, mvUser.Password, False, False)
      Else
        'Named Connection already exists - If connect string given then attempt reconnection
        If pConnectString.Length > 0 Then
          vConnection = mvConnections(pName)
          vConnection.CloseConnection()
          vConnection.OpenConnection(pConnectString, mvUser.DatabaseLogname, mvUser.Password, False, False)
        End If
      End If
      Return mvConnections(pName)
    End Function

    Public Property AuditStyle() As AuditStyleTypes
      Get
        If Not mvAuditStyleChecked Then ResetAuditStyle()
        Return mvAuditStyle
      End Get
      Set(ByVal pValue As AuditStyleTypes)
        mvAuditStyle = pValue
      End Set
    End Property

    Public Function InsertWithAmendmentHistory(ByVal pConn As CDBConnection, ByVal pTableName As String, ByVal pFields As CDBFields) As Integer
      pConn.InsertRecord(pTableName, pFields)
      Dim vClassFields As New ClassFields
      For Each vField As CDBField In pFields
        Dim vClassField As ClassField = New ClassField(vField.Name, vField.FieldType)
        vClassField.Value = vField.Value
        vClassFields.Add(vClassField)
      Next
      AddAmendmentHistory(AuditTypes.audInsert, pTableName, 0, 0, Me.User.UserID, vClassFields)
    End Function

    Public Function InsertWithExtendedAmendmentHistory(ByVal pConn As CDBConnection, ByVal pTableName As String, ByVal pClassFields As ClassFields, ByVal pSelect1 As Integer) As Integer
      pConn.InsertRecord(pTableName, pClassFields.UpdateFields)
      If AuditStyle = AuditStyleTypes.ausExtended Then
        AddAmendmentHistory(AuditTypes.audInsert, pTableName, pSelect1, 0, Me.User.UserID, pClassFields)
      End If
    End Function

    Public Sub AddAmendmentHistory(ByVal pType As AuditTypes, ByVal pTable As String, ByVal pSelect1 As Integer, ByVal pSelect2 As Integer, ByVal pLogname As String, ByVal pClassFields As ClassFields)
      AddAmendmentHistory(pType, pTable, pSelect1, pSelect2, pLogname, pClassFields, 0, False)
    End Sub

    Public Sub AddAmendmentHistory(ByVal pType As AuditTypes, ByVal pTable As String, ByVal pSelect1 As Integer, ByVal pSelect2 As Integer, ByVal pLogname As String, ByVal pClassFields As ClassFields, ByVal pJournalNumber As Integer)
      AddAmendmentHistory(pType, pTable, pSelect1, pSelect2, pLogname, pClassFields, pJournalNumber, False)
    End Sub

    Public Sub AddAmendmentHistory(ByVal pType As AuditTypes, ByVal pTable As String, ByVal pSelect1 As Integer, ByVal pSelect2 As Integer, ByVal pLogname As String, ByVal pClassFields As ClassFields, ByVal pJournalNumber As Integer, ByVal pForceEntry As Boolean)
      Dim vOperation As String
      Dim vFields As New CDBFields
      Dim vSelect1 As String
      Dim vSelect2 As String
      Dim vValues As New StringBuilder
      Dim vNewFields As New CDBFields
      Dim vOldFields As New CDBFields

      If pLogname.Length = 0 Then pLogname = mvUser.Logname
      If AuditStyle = AuditStyleTypes.ausAmendmentHistory OrElse AuditStyle = AuditStyleTypes.ausExtended OrElse pForceEntry Then
        Select Case pType
          Case AuditTypes.audInsert
            vOperation = "insert"
          Case AuditTypes.audUpdate
            vOperation = "update"
          Case AuditTypes.audDelete
            vOperation = "delete"
          Case Else
            vOperation = "unknown"
        End Select
        If pSelect1 > 0 Then
          vSelect1 = pSelect1.ToString
        Else
          vSelect1 = ""
        End If
        If pSelect2 > 0 Then
          vSelect2 = pSelect2.ToString
        Else
          vSelect2 = ""
        End If

        With vFields
          .Add("operation", vOperation)
          .Add("operation_date", CDBField.FieldTypes.cftTime, TodaysDateAndTime)
          .Add("table_name", pTable)
          .Add("logname", pLogname)
          .Add("select_1", CDBField.FieldTypes.cftLong, vSelect1)
          .Add("select_2", CDBField.FieldTypes.cftLong, vSelect2)
          If pJournalNumber > 0 Then .Add("contact_journal_number", CDBField.FieldTypes.cftLong, pJournalNumber.ToString)
          If pClassFields.TableMaintenance AndAlso GetDataStructureInfo(cdbDataStructureConstants.cdbSystemMaintenance) Then .Add("table_maintenance", "Y")
          vValues = GetDataValues(pType, pSelect1, pClassFields)
          If vValues.Length > 0 Then
            .Add("data_values", CDBField.FieldTypes.cftMemo, vValues.ToString)
            If Me.GetDataStructureInfo(cdbDataStructureConstants.cdbAmendmentContactNumber) Then
              If pClassFields.ContainsKey("organisation_number") AndAlso pTable.StartsWith("organisation") Then
                .Add("contact_number", pClassFields("organisation_number").IntegerValue)
              ElseIf pClassFields.ContainsKey("organisation_number_1") AndAlso pTable.StartsWith("organisation") Then
                .Add("contact_number", pClassFields("organisation_number_1").IntegerValue)
              ElseIf pClassFields.ContainsKey("contact_number") AndAlso pClassFields("contact_number").IntegerValue > 0 Then
                .Add("contact_number", pClassFields("contact_number").IntegerValue)
                'This next is to update any address which may already have been created but does not have the contact number set
                If (pTable = "organisation_addresses" OrElse pTable = "contact_addresses") AndAlso pType = AuditTypes.audInsert Then
                  Dim vWhereFields As New CDBFields
                  vWhereFields.Add("table_name", "addresses")
                  vWhereFields.Add("select_1", pClassFields("address_number").IntegerValue)
                  vWhereFields.Add("contact_number")
                  Dim vUpdateFields As New CDBFields
                  vUpdateFields.Add("contact_number", pClassFields("contact_number").IntegerValue)
                  Connection.UpdateRecords("amendment_history", vUpdateFields, vWhereFields, False)
                End If
              ElseIf pClassFields.ContainsKey("contact_number_1") Then
                .Add("contact_number", pClassFields("contact_number_1").IntegerValue)
              ElseIf AuditStyle = AuditStyleTypes.ausExtended Then
                If pJournalNumber > 0 Then
                  Dim vSQL As New SQLStatement(Connection, "contact_number", "contact_journals", New CDBField("contact_journal_number", pJournalNumber))
                  .Add("contact_number", CDBField.FieldTypes.cftInteger, vSQL.GetValue)
                ElseIf pTable = "addresses" Then
                  Dim vSQL As New SQLStatement(Connection, "contact_number", "contact_addresses", New CDBField("address_number", pClassFields("address_number").IntegerValue))
                  Dim vContactNumber As String = vSQL.GetValue
                  If vContactNumber.Length = 0 Then
                    vSQL = New SQLStatement(Connection, "organisation_number", "organisation_addresses", New CDBField("address_number", pClassFields("address_number").IntegerValue))
                    vContactNumber = vSQL.GetValue
                  End If
                  If IntegerValue(vContactNumber) > 0 Then .Add("contact_number", CDBField.FieldTypes.cftInteger, vContactNumber)
                ElseIf pTable = "communications" Then       'Numbers for an organisation have a null contact number - get the org number from the organisation addresses table
                  Dim vSQL As New SQLStatement(Connection, "organisation_number", "organisation_addresses", New CDBField("address_number", pClassFields("address_number").IntegerValue))
                  Dim vContactNumber As String = vSQL.GetValue
                  If IntegerValue(vContactNumber) > 0 Then .Add("contact_number", CDBField.FieldTypes.cftInteger, vContactNumber)
                ElseIf pTable = "communications_log_subjects" Then
                  Dim vSQL As New SQLStatement(Connection, "contact_number", "communications_log", New CDBField("communications_log_number", pClassFields("communications_log_number").IntegerValue))
                  Dim vContactNumber As String = vSQL.GetValue
                  If IntegerValue(vContactNumber) > 0 Then .Add("contact_number", CDBField.FieldTypes.cftInteger, vContactNumber)
                End If
              End If
            End If
            Connection.InsertRecord("amendment_history", vFields)

            Dim vMessageQueueName As String = GetConfig("amendments_message_queue")       '.\Private$\NGAmendments
            If vMessageQueueName.Length > 0 Then
              vFields.Add("message_queue_name", vMessageQueueName)
              Dim vThreadStart As System.Threading.ThreadStart = New Threading.ThreadStart(AddressOf vFields.WriteToMessageQueue)
              Dim vThread As New Threading.Thread(vThreadStart)
              vThread.Start()
            End If
          End If
        End With
      End If
    End Sub
    Private Function GetDataValues(ByVal pType As AuditTypes, ByVal pSelect1 As Integer, ByVal pClassFields As ClassFields) As StringBuilder
      Dim vValues = New StringBuilder
      Dim vClassField As ClassField
      If pType <> AuditTypes.audInsert Then
        vValues.Append("OLD")
        vValues.Append(Chr(22))
        For Each vClassField In pClassFields
          If pType = AuditTypes.audUpdate Then
            If vClassField.ValueChanged OrElse vClassField.ForceAmendmentHistory OrElse (pSelect1 = 0 AndAlso vClassField.PrimaryKey) OrElse AuditStyle = AuditStyleTypes.ausExtended Then
              vValues.Append(vClassField.Name)
              vValues.Append(":")
              vValues.Append(vClassField.SetValue)
              vValues.Append(Chr(22))
            End If
          Else
            vValues.Append(vClassField.Name)
            vValues.Append(":")
            vValues.Append(vClassField.Value)
            vValues.Append(Chr(22))
          End If
        Next
        vValues.Append(vbCrLf)
      End If
      If pType <> AuditTypes.audDelete Then
        vValues.Append("NEW")
        vValues.Append(Chr(22))
        For Each vClassField In pClassFields
          If pType = AuditTypes.audUpdate Then
            If vClassField.ValueChanged OrElse vClassField.ForceAmendmentHistory OrElse (pSelect1 = 0 AndAlso vClassField.PrimaryKey) OrElse AuditStyle = AuditStyleTypes.ausExtended Then
              vValues.Append(vClassField.Name)
              vValues.Append(":")
              vValues.Append(vClassField.FormattedValue)
              vValues.Append(Chr(22))
            End If
          Else
            vValues.Append(vClassField.Name)
            vValues.Append(":")
            vValues.Append(vClassField.FormattedValue)
            vValues.Append(Chr(22))
          End If
        Next
        vValues.Append(vbCrLf)
      End If
      Return vValues
    End Function
    Public Sub BulkInsertAmendmentHistory(ByVal pType As AuditTypes, ByVal pTable As String, ByVal pSelectIndex1 As Integer, ByVal pSelectIndex2 As Integer, ByVal pLogname As String, ByVal pClassFieldsList As List(Of ClassFields), ByVal pJournalNumber As Integer)

      Dim vDeferSavedList As New DataTable()
      Dim vOperation As String
      Dim vSelect1 As Integer
      Dim vSelect2 As Integer

      Select Case pType
        Case AuditTypes.audInsert
          vOperation = "insert"
        Case AuditTypes.audUpdate
          vOperation = "update"
        Case AuditTypes.audDelete
          vOperation = "delete"
        Case Else
          vOperation = "unknown"
      End Select

      vDeferSavedList.TableName = "amendment_history"
      vDeferSavedList.Columns.Add("operation", GetType(String))
      vDeferSavedList.Columns.Add("operation_date", GetType(DateTime))
      vDeferSavedList.Columns.Add("table_name", GetType(String))
      vDeferSavedList.Columns.Add("logname", GetType(String))
      vDeferSavedList.Columns.Add("data_values", GetType(String))
      vDeferSavedList.Columns.Add("select_1", GetType(Integer))
      vDeferSavedList.Columns.Add("select_2", GetType(Integer))
      vDeferSavedList.Columns.Add("contact_journal_number", GetType(Integer))
      vDeferSavedList.Columns.Add("table_maintenance", GetType(String))
      vDeferSavedList.Columns.Add("contact_number", GetType(Integer))

      For Each vClassFields As ClassFields In pClassFieldsList
        If pSelectIndex1 > 0 AndAlso _
          (vClassFields(pSelectIndex1).FieldType = CDBField.FieldTypes.cftInteger OrElse vClassFields(pSelectIndex1).FieldType = CDBField.FieldTypes.cftLong) Then
          vSelect1 = vClassFields(pSelectIndex1).LongValue
        End If
        If pSelectIndex2 > 0 AndAlso _
          (vClassFields(pSelectIndex2).FieldType = CDBField.FieldTypes.cftInteger OrElse vClassFields(pSelectIndex2).FieldType = CDBField.FieldTypes.cftLong) Then
          vSelect2 = vClassFields(pSelectIndex2).LongValue
        End If

        Dim vFields As New StringBuilder()
        vFields = GetDataValues(pType, vSelect1, vClassFields)
        Dim vAmendmentHistoryRow As DataRow = vDeferSavedList.NewRow()
        vAmendmentHistoryRow("operation") = vOperation
        vAmendmentHistoryRow("operation_date") = TodaysDateAndTime()
        vAmendmentHistoryRow("table_name") = pTable
        vAmendmentHistoryRow("logname") = User.UserID.ToString
        vAmendmentHistoryRow("data_values") = vFields.ToString
        If vSelect1 > 0 Then
          vAmendmentHistoryRow("select_1") = vSelect1.ToString
        Else
          vAmendmentHistoryRow("select_1") = DBNull.Value
        End If
        If vSelect2 > 0 Then
          vAmendmentHistoryRow("select_2") = vSelect2.ToString
        Else
          vAmendmentHistoryRow("select_2") = DBNull.Value
        End If
        vAmendmentHistoryRow("contact_journal_number") = DBNull.Value
        vAmendmentHistoryRow("table_maintenance") = DBNull.Value
        vAmendmentHistoryRow("contact_number") = DBNull.Value
        vDeferSavedList.Rows.Add(vAmendmentHistoryRow)
      Next vClassFields
      Connection.BulkCopyData(Nothing, "amendment_history", vDeferSavedList)

    End Sub
    Public Function CommitOutstandingTransaction() As Boolean
      If mvConnection IsNot Nothing Then
        If mvConnection.ConnectionOpen Then
          If mvConnection.InTransaction Then
            mvConnection.CommitTransaction()
            Return True
          End If
        End If
      End If
      Return False
    End Function

    Public Function ContactNameStyle() As String
      If mvNameStyle.Length = 0 Then
        Dim vNameStyle As New StringBuilder
        Dim vFormatConfig As String = GetConfig("contact_name_format")
        Dim vLen As Integer = vFormatConfig.Length
        Dim vPos As Integer = 0
        Dim vFormat As String
        While vPos < vLen
          vFormat = vFormatConfig.Substring(vPos)
          If vFormat.StartsWith("title") Then
            vPos += 5
            vNameStyle.Append("T")
          ElseIf vFormat.StartsWith("initials") Then
            vPos += 8
            vNameStyle.Append("I")
          ElseIf vFormat.StartsWith("surname") Then
            vPos += 7
            vNameStyle.Append("S")
          ElseIf vFormat.StartsWith("forenames") Then
            vPos += 9
            vNameStyle.Append("F")
          ElseIf vFormat.StartsWith("preferred_forename") Then
            vPos += 18
            vNameStyle.Append("P")
          ElseIf vFormat.StartsWith("honorifics") Then
            vPos += 10
            vNameStyle.Append("H")
          ElseIf vFormat.StartsWith("(") Then
            vPos += 1
            vNameStyle.Append("(")
          ElseIf vFormat.StartsWith(")") Then
            vPos += 1
            vNameStyle.Append(")")
          ElseIf vFormat.StartsWith(",") Then
            vPos += 1
          ElseIf vFormat.StartsWith(" ") Then
            vPos += 1
          Else
            RaiseError(DataAccessErrors.daeContactNameFormat, GetConfig("contact_name_format"))
          End If
        End While
        If vNameStyle.Length = 0 Then
          mvNameStyle = "TIS"
        Else
          mvNameStyle = vNameStyle.ToString
        End If
      End If
      Return mvNameStyle
    End Function

    Public ReadOnly Property DefaultCountry() As String
      Get
        Dim vCountrySetting As String = ""
        If mvDefaultCountry.Length = 0 Then
          mvDefaultCountry = GetConfig("option_country")
          If mvDefaultCountry.Length = 0 Then
            Dim vKey As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Control Panel\International")
            If vKey IsNot Nothing Then vCountrySetting = vKey.GetValue("iCountry").ToString
            Select Case vCountrySetting
              Case "31"
                mvDefaultCountry = "NL" 'Netherlands
              Case "41"
                mvDefaultCountry = "CH"
              Case Else
                mvDefaultCountry = "UK"
            End Select
          End If
        End If
        Return mvDefaultCountry
      End Get
    End Property

    Public ReadOnly Property EntityGroups() As EntityGroups
      Get
        If mvEntityGroups Is Nothing Then mvEntityGroups = New EntityGroups(Me)
        Return mvEntityGroups
      End Get
    End Property

    Public Function GetBranchFromPostcode(ByVal pPostCode As String) As String
      'Does iterative searches on database to broaden postcode branch link
      'eg. if the postcode entered is XX99 4ZZ the searches will be on
      'XX994ZZ, XX994, XX99, XX in that sequence, unless or until a valid record is found.
      'Modified from XX994, XX99, XX
      Dim vBranch As String = ""
      If pPostCode.Length > 0 Then
        Dim vPostcode As Postcode = New Postcode(pPostCode)
        Dim vParams As CDBParameters = vPostcode.Components
        If vParams.Count > 0 Then
          Dim vRS As CDBRecordSet = New SQLStatement(Connection, "branch,outward_postcode", "branch_postcodes", New CDBField("outward_postcode", vParams.InList, CDBField.FieldWhereOperators.fwoIn)).GetRecordSet
          While vRS.Fetch
            For Each vParam As CDBParameter In vParams
              If vParam.Name = vRS.Fields(2).Value Then vParam.Value = vRS.Fields(1).Value
            Next
          End While
          vRS.CloseRecordSet()
          For Each vParam As CDBParameter In vParams
            If vParam.Value.Length > 0 Then
              vBranch = vParam.Value
              Exit For
            End If
          Next
        End If
      End If
      Return vBranch
    End Function

    Public Function BankAccountEntryLength(ByVal pDirectDebit As Boolean) As Integer
      Select Case GetConfig("fp_bank_account_format").ToUpper
        Case "EURO"
          BankAccountEntryLength = 34       'Long enough for IBAN?
        Case "SWISS"
          BankAccountEntryLength = 16
        Case Else
          Select Case DefaultCountry
            Case "CH"
              BankAccountEntryLength = 16
            Case "NL"
              BankAccountEntryLength = 9
            Case Else
              If pDirectDebit Then
                BankAccountEntryLength = 8
              Else
                BankAccountEntryLength = 13
              End If
          End Select
      End Select
    End Function

    Public ReadOnly Property BranchResourceID() As Integer
      Get
        Return 14600        'Default resource ID for 'Branch'
      End Get
    End Property

    Public Function GetBranchText(ByVal pResourceText As String) As String
      Dim vConfig As String = GetConfig("me_branch_description")
      If vConfig.Length > 0 Then
        Dim vDesc() As String = Split(vConfig & ",", ",")
        If vDesc(0).Length > 0 Then
          Dim vText As String = pResourceText
          If vDesc.Length > 0 AndAlso vDesc(1).Length > 0 Then vText = pResourceText.Replace("Branches", vDesc(1))
          vText = vText.Replace("Branch", vDesc(0))
          Return vText
        Else
          Return pResourceText
        End If
      Else
        Return pResourceText
      End If
    End Function

    Public Function GetDataStructureInfo(ByVal pType As cdbDataStructureConstants) As Boolean
      Dim vTable As String = ""
      Dim vAttr As String = ""
      Dim vDataType As String = ""
      Dim vCheck As cdbDSConstants

      vCheck = cdbDSConstants.cdbDSExists
      Select Case pType
        Case cdbDataStructureConstants.cdbDataBTAQuantityDecimal                 'KEEP Not Version Specific but Client Specific
          vTable = "batch_transaction_analysis"
          vAttr = "quantity"
          vCheck = cdbDSConstants.cdbDSType
          vDataType = "N"
        Case cdbDataStructureConstants.cdbDataCurrencyCode                        'KEEP Added in 3.9 But an optional table
          vTable = "currency_codes"
          vAttr = "currency_code"
          '-------------------------------------------------------------------------------
          ' 4.51 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataMailJointsMethod, _
             cdbDataStructureConstants.cdbDataPrefulfilledIncentives, _
             cdbDataStructureConstants.cdbDataIncentiveProductMinMax, _
             cdbDataStructureConstants.cdbDataGiftAidSponsorship, _
             cdbDataStructureConstants.cdbDataCampaignBudgets, _
             cdbDataStructureConstants.cdbDataMembershipProRating
          Return True
          '-------------------------------------------------------------------------------
          ' 5.0 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataReportItemFormatting, _
             cdbDataStructureConstants.cdbDataWarehouses, _
             cdbDataStructureConstants.cdbDataScheduledPayments, _
             cdbDataStructureConstants.cdbDataWarehouseTransferReason, _
             cdbDataStructureConstants.cdbDataCampaignActualIncome, _
             cdbDataStructureConstants.cdbDataPayrollGivingPayFrequency, _
             cdbDataStructureConstants.cdbDataAutoEmailDevice, _
             cdbDataStructureConstants.cdbDataGiftAidDecCreatedBy, _
             cdbDataStructureConstants.cdbDataConfirmSRTransactions
          Return True
          '-------------------------------------------------------------------------------
          ' 5.1 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataExplorerLinks, _
             cdbDataStructureConstants.cdbDataHeaderPostalSector, _
             cdbDataStructureConstants.cdbDataPayrollGivingPaymentNumber, _
             cdbDataStructureConstants.cdbDataDevicesWWWAddress, _
             cdbDataStructureConstants.cdbDataAUNotesMandatory, _
             cdbDataStructureConstants.cdbDataGiftAidMergeCancellation, _
             cdbDataStructureConstants.cdbDataOnLineCCAuthorisation, _
             cdbDataStructureConstants.cdbDataCPDCycle, _
             cdbDataStructureConstants.cdbDataPayPlanConvMaintenance, _
             cdbDataStructureConstants.cdbCalendarCompany, _
             cdbDataStructureConstants.cdbDataPrintChequeList, _
             cdbDataStructureConstants.cdbDataIssuedStockJobNumber, _
             cdbDataStructureConstants.cdbDataGiftAidMaxJuniorAge, _
             cdbDataStructureConstants.cdbDataDataUpdates, _
             cdbDataStructureConstants.cdbDataPayPlanEligibleForGiftAid
          Return True
          '-------------------------------------------------------------------------------
          ' 5.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataDisplayListItems, _
             cdbDataStructureConstants.cdbDataBankStatementNotes, _
             cdbDataStructureConstants.cdbDataNominalAccountValidation, _
             cdbDataStructureConstants.cdbDataSubTopicActivityDuration, _
             cdbDataStructureConstants.cdbDataEventBookingNotes, _
             cdbDataStructureConstants.cdbDataPayPlanDetailCreatedBy, _
             cdbDataStructureConstants.cdbDataCommunicationNumber, _
             cdbDataStructureConstants.cdbDataDevicesSequenceNumber, _
             cdbDataStructureConstants.cdbDataBTAWarehouse
          Return True
          '-------------------------------------------------------------------------------
          ' 5.3 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataUnknownAddresses, _
             cdbDataStructureConstants.cdbDataRelationshipGroups, _
             cdbDataStructureConstants.cdbDataStandardDocumentPrecis, _
             cdbDataStructureConstants.cdbDataCheckPayPlans, _
             cdbDataStructureConstants.cdbDataEmailControls, _
             cdbDataStructureConstants.cdbDataMailingNotes, _
             cdbDataStructureConstants.cdbDataUserHistory, _
             cdbDataStructureConstants.cdbDataControlsContactGroup, _
             cdbDataStructureConstants.cdbDataAgencyAdminFee
          Return True
          '-------------------------------------------------------------------------------
          ' 5.4 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataContactRoleNumber, _
             cdbDataStructureConstants.cdbDataAdultGiftMemEligibleGA, _
             cdbDataStructureConstants.cdbDataEventWaitingListControl, _
             cdbDataStructureConstants.cdbDataUserHistoryItems
          Return True
          '-------------------------------------------------------------------------------
          ' 5.5 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataEventGroups, _
             cdbDataStructureConstants.CDBDataViewNames, _
             cdbDataStructureConstants.cdbDataExplorerLinksToolbar, _
             cdbDataStructureConstants.cdbDataPositionFuntionSeniority, _
             cdbDataStructureConstants.cdbDataEntityGroupSpecificStatus, _
             cdbDataStructureConstants.cdbDataCustomFormDeletion, _
             cdbDataStructureConstants.cdbDataMailingHistoryNotes
          Return True
          '-------------------------------------------------------------------------------
          ' 5.6 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataControlParameterName, _
             cdbDataStructureConstants.cdbDataLinkToCommunication, _
             cdbDataStructureConstants.cdbDataActivityGroupsCampaign, _
             cdbDataStructureConstants.cdbDataNumberofCCCAs, _
             cdbDataStructureConstants.cdbDataPayPlanPackToDonor, _
             cdbDataStructureConstants.cdbDataEventLongDescription, _
             cdbDataStructureConstants.cdbDataEventClass, _
             cdbDataStructureConstants.cdbDataBacsSkipRejectedPayment, _
             cdbDataStructureConstants.cdbDataLegaciesReviewReason
          Return True
          '-------------------------------------------------------------------------------
          ' 5.7 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataPostTaxPayrollGiving, _
             cdbDataStructureConstants.cdbDataMembershipTypeFixedCycle, _
             cdbDataStructureConstants.cdbDataMembershipCardIssueNumber, _
             cdbDataStructureConstants.cdbDataActivityDate, _
             cdbDataStructureConstants.cdbDataDelegateActivities, _
             cdbDataStructureConstants.cdbDataProductEligibleGA, _
             cdbDataStructureConstants.cdbDataLookupGroups, _
             cdbDataStructureConstants.cdbDataConfirmedTransStatus, _
             cdbDataStructureConstants.cdbDataAppealType, _
             cdbDataStructureConstants.cdbDataCollections, _
             cdbDataStructureConstants.cdbDataSegmentOrgSelectionOptions
          Return True
          '-------------------------------------------------------------------------------
          ' 5.8 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataGiftAidTimeLimits, _
             cdbDataStructureConstants.cdbDataEventFinancialAnalysis, _
             cdbDataStructureConstants.cdbDataCreditCardAVSCVV2, _
             cdbDataStructureConstants.cdbDataHolidayLets, _
             cdbDataStructureConstants.cdbDataCustomFinderTab, _
             cdbDataStructureConstants.cdbDataJobScheduleUpdateDates, _
             cdbDataStructureConstants.cdbDataServiceStartDays, _
             cdbDataStructureConstants.cdbDataPrimaryRelationship, _
             cdbDataStructureConstants.cdbDataCustomFormAllowUpdate, _
             cdbDataStructureConstants.cdbDataStandardDocumentClass, _
             cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports, _
             cdbDataStructureConstants.cdbDataCustomFormRestrictions, _
             cdbDataStructureConstants.cdbDataEventTopicNotes, _
             cdbDataStructureConstants.cdbDataEventPIS, _
             cdbDataStructureConstants.cdbDataBranchOwnershipGroup
          Return True
          '-------------------------------------------------------------------------------
          ' 5.9 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataContactControlsDevices, _
             cdbDataStructureConstants.cdbDataPayPlanPackToMember
          Return True
          '-------------------------------------------------------------------------------
          ' 6.0 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataNINumber, _
             cdbDataStructureConstants.cdbDataMultipleMerchantRetailNos, _
             cdbDataStructureConstants.cdbDataAddressLines, _
             cdbDataStructureConstants.cdbDataCommunicationsUsages, _
             cdbDataStructureConstants.cdbDataBatchAnalysisCodes, _
             cdbDataStructureConstants.cdbDataMembershipPrices, _
             cdbDataStructureConstants.cdbDataRegistrationData
          Return True
          '-------------------------------------------------------------------------------
          ' 6.1 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataBatchCampaign, _
             cdbDataStructureConstants.cdbDataProductCosts, _
             cdbDataStructureConstants.cdbDataPurchaseOrderLink, _
             cdbDataStructureConstants.cdbDataGiftMessage, _
             cdbDataStructureConstants.cdbDataPackProducts, _
             cdbDataStructureConstants.cdbDataDutchSupport, _
             cdbDataStructureConstants.cdbDataContactGroupUsers, _
             cdbDataStructureConstants.cdbDataDefaultBankAccount, _
             cdbDataStructureConstants.cdbDataPurchaseOrderManagement, _
             cdbDataStructureConstants.cdbDataDashboardItems, _
             cdbDataStructureConstants.cdbDataSecurityQuestion, _
             cdbDataStructureConstants.cdbDataCCIssueNumber, _
             cdbDataStructureConstants.cdbDataBankingDate, _
             cdbDataStructureConstants.cdbDataPaymentPlanStartMonth
          Return True
          '-------------------------------------------------------------------------------
          ' 6.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataStandardPositions, _
             cdbDataStructureConstants.cdbDataStockMovementTransactionID, _
             cdbDataStructureConstants.cdbDataIrishGiftAid, _
             cdbDataStructureConstants.cdbDataDelegateSessions, _
             cdbDataStructureConstants.cdbDataEventDocumentLinks, _
             cdbDataStructureConstants.cdbDataServiceControlRestrictions, _
             cdbDataStructureConstants.cdbDataPPDetailsEffectiveDate, _
             cdbDataStructureConstants.cdbDataWebPageSuppressions, _
             cdbDataStructureConstants.cdbDataEventFixedPrice, _
             cdbDataStructureConstants.cdbDataWebPageLoginRequired, _
             cdbDataStructureConstants.cdbDataHistoryOnlyAccount
          Return True
          '-------------------------------------------------------------------------------
          ' 6.3 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataHistoryOnlyCancReasons, _
             cdbDataStructureConstants.cdbDataWebPagePublished, _
             cdbDataStructureConstants.cdbDataMembershipGroups
          Return True
          '-------------------------------------------------------------------------------
          ' 6.4 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataMembershipTypeTransitions, _
             cdbDataStructureConstants.cdbDataPurchaseOrderLineItems, _
             cdbDataStructureConstants.cdbDataThankYouMessage, _
             cdbDataStructureConstants.cdbDataUserHistoryFavourites, _
             cdbDataStructureConstants.cdbDataWebFriendlyUrl, _
             cdbDataStructureConstants.cdbDataJobScheduleSmartClientJob, _
             cdbDataStructureConstants.cdbDataAllowAsFirstType, _
             cdbDataStructureConstants.cdbDataTransactionOrigins, _
             cdbDataStructureConstants.cdbDataEmailJobs, _
             cdbDataStructureConstants.cdbDataCheetahMail
          Return True
          '-------------------------------------------------------------------------------
          ' 6.5 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataMinPriceMandatory, _
             cdbDataStructureConstants.cdbDataAppealExpenditureROI, _
             cdbDataStructureConstants.cdbDataCampaignItemisedCosts, _
             cdbDataStructureConstants.cdbDataCampaignRoles, _
             cdbDataStructureConstants.cdbDataAddressDPS, _
             cdbDataStructureConstants.cdbDataBankAccountDepartments, _
             cdbDataStructureConstants.cdbDataControlNumberLinks, _
             cdbDataStructureConstants.cdbDataMailingSuppressionsNotes, _
             cdbDataStructureConstants.cdbDataEventMultipleAnalysis, _
             cdbDataStructureConstants.cdbDataEventAdultChildQuantity, _
             cdbDataStructureConstants.cdbDataDashboardViewNames, _
             cdbDataStructureConstants.cdbDataDistributionBoxProcess
          Return True
          '-------------------------------------------------------------------------------
          ' 6.6 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataHistoryOnlyDistribCodes, _
             cdbDataStructureConstants.cdbDataPayrollGivingCreatedByOn, _
             cdbDataStructureConstants.cdbDataFastDataEntry, _
             cdbDataStructureConstants.cdbDataRegisteredUsersAmendedOn
          Return True
          '-------------------------------------------------------------------------------
          ' 6.7 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataCustomFinderWildcards, _
             cdbDataStructureConstants.cdbDataEventPricingMatrix, _
             cdbDataStructureConstants.cdbDataEditPanelPages, _
             cdbDataStructureConstants.cdbDataPOPMultiplePayees, _
             cdbDataStructureConstants.cdbDataPrintJobNumber, _
             cdbDataStructureConstants.cdbDataServiceBookingAnalysis, _
             cdbDataStructureConstants.cdbDataBoxProcessStatus
          Return True
          '-------------------------------------------------------------------------------
          ' 6.8 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataProvisionalInvoiceNumber, _
             cdbDataStructureConstants.cdbDataAlbacsBankDetails, _
             cdbDataStructureConstants.cdbDataChequeReissue, _
             cdbDataStructureConstants.cdbDataPositionLinks, _
             cdbDataStructureConstants.cdbDataBackClaimYears, _
             cdbDataStructureConstants.cdbDataPOPPayByBACS, _
             cdbDataStructureConstants.cdbDataControlReadonlyAndPanels, _
             cdbDataStructureConstants.cdbDataDefaultValue, _
             cdbDataStructureConstants.cdbDataRgbValueForStatus, _
             cdbDataStructureConstants.cdbDataRgbValueForMemberType, _
             cdbDataStructureConstants.cdbDataRgbValueForActivityValue, _
             cdbDataStructureConstants.cdbDataFundraisingPayments, _
             cdbDataStructureConstants.cdbContactAlerts, _
             cdbDataStructureConstants.cdbDataOwnerContactNumber
          Return True
          '-------------------------------------------------------------------------------
          ' 7.1 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbDataRelationshipStatus, _
             cdbDataStructureConstants.cdbEventVenueCapacity, _
             cdbDataStructureConstants.cdbStatusMessage, _
             cdbDataStructureConstants.cdbDaysPrior, _
             cdbDataStructureConstants.cdbDaysPrior, _
             cdbDataStructureConstants.cdbMembershipLookupGroup, _
             cdbDataStructureConstants.cdbMembershipStatus, _
             cdbDataStructureConstants.cdbCPDCycleStatus, _
             cdbDataStructureConstants.cdbCPDPointsNotes, _
             cdbDataStructureConstants.cdbCPDObjective, _
             cdbDataStructureConstants.cdbActivityCPDPoints, _
             cdbDataStructureConstants.cdbBulkEmailAttachments, _
             cdbDataStructureConstants.cdbDataLabelName, _
             cdbDataStructureConstants.cdbDataAllocationsOnIPH, _
             cdbDataStructureConstants.cdbRateModifier, _
             cdbDataStructureConstants.cdbEventStatusColor, _
             cdbDataStructureConstants.cdbDataVatRateHistory, _
             cdbDataStructureConstants.cdbPriceIsPercentage, _
             cdbDataStructureConstants.cdbOutlookIntegration, _
             cdbDataStructureConstants.cdbSalesContactMandatory, _
             cdbDataStructureConstants.cdbDataMerchantDetails, _
             cdbDataStructureConstants.cdbDataPayPlansVatExcl
          Return True
        Case cdbDataStructureConstants.cdbEntityAlerts
          vTable = "entity_alerts"
          vAttr = "entity_alert_number"
        Case cdbDataStructureConstants.cdbJobScheduleJobId
          vTable = "job_schedule"
          vAttr = "job_id"
        Case cdbDataStructureConstants.cdbArchiveCommunications
          vTable = "communications"
          vAttr = "archive"
          '-------------------------------------------------------------------------------
          ' 7.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbProductWebPublish, _
             cdbDataStructureConstants.cdbGrabDetailsOption, _
             cdbDataStructureConstants.cdbLongDescription, _
             cdbDataStructureConstants.cdbEventWebPublish, _
             cdbDataStructureConstants.cdbValidFromToForRegisteredUsers, _
             cdbDataStructureConstants.cdbLoginLockout, _
             cdbDataStructureConstants.cdbWebPublish, _
             cdbDataStructureConstants.cdbMembershipTypesWebPublish, _
             cdbDataStructureConstants.cdbCreatedBy, _
             cdbDataStructureConstants.cdbAdminEmailAddress, _
             cdbDataStructureConstants.cdbAccessViewName, _
             cdbDataStructureConstants.cdbAddErrorLog, _
             cdbDataStructureConstants.cdbCPDWebPublish, _
             cdbDataStructureConstants.cdbFundraisingBusinessType, _
             cdbDataStructureConstants.cdbJournalSelectName, _
             cdbDataStructureConstants.cdbBankHolidayDays, _
             cdbDataStructureConstants.cdbTelemarketing, _
             cdbDataStructureConstants.cdbAuthorisedTextID, _
             cdbDataStructureConstants.cdbTraderInvoicePrintPreview, _
             cdbDataStructureConstants.cbdProtxCardType
          Return True
        Case cdbDataStructureConstants.cdbLabelNameFormatCode
          vTable = "contacts"
          vAttr = "label_name_format_code"
          '-------------------------------------------------------------------------------
          ' 11.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cbdGroupDefaultNameFormat
          vTable = "contact_groups"
          vAttr = "name_format"
        Case cdbDataStructureConstants.cdbStatusReasons
          vTable = "status_reasons"
          vAttr = "status_reason"
        Case cdbDataStructureConstants.cdbAdHocPurchaseOrderPayments
          vTable = "purchase_order_types"
          vAttr = "ad_hoc_payments"
        Case cdbDataStructureConstants.cdbPurchaseOrderAuthorisation
          vTable = "purchase_orders"
          vAttr = "authorised_by"
        Case cdbDataStructureConstants.cdbPurchaseOrderHistory
          vTable = "purchase_order_history"
          vAttr = "purchase_order_number"
        Case cdbDataStructureConstants.cdbBACSePay
          vTable = "financial_controls"
          vAttr = "bacs_msg_file_format"
        Case cdbDataStructureConstants.cdbDataOrgGroupCustomTables
          vTable = "organisation_groups"
          vAttr = "custom_table_names"
        Case cdbDataStructureConstants.cdbLoans
          vTable = "loans"
          vAttr = "loan_number"
        Case cdbDataStructureConstants.cdbCustomFormWebPage
          vTable = "custom_forms"
          vAttr = "custom_form_url"
        Case cdbDataStructureConstants.cdbRegularPurchaseOrderPayments
          vTable = "purchase_order_types"
          vAttr = "regular_payments"
        Case cdbDataStructureConstants.cdbRateModifiersUseActivityDate
          vTable = "rate_modifiers"
          vAttr = "use_activity_date"
        Case cdbDataStructureConstants.cdbMembershipPricesOverseas
          vTable = "membership_prices"
          vAttr = "overseas"
        Case cdbDataStructureConstants.cdbCountryAddressFormat
          vTable = "countries"
          vAttr = "address_format"
        Case cdbDataStructureConstants.cdbBACSErrorCode
          vTable = "bacs_amendments"
          vAttr = "bacs_msg_processed"
        Case cdbDataStructureConstants.cdbProductActivityDurationMonths
          vTable = "products"
          vAttr = "activity_duration_months"
        Case cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode
          vTable = "purchase_orders"
          vAttr = "currency_code"
        Case cdbDataStructureConstants.cdbDataCPDPoints2
          vTable = "contact_cpd_points"
          vAttr = "cpd_points_2"
        Case cdbDataStructureConstants.cdbDataPaymentPlanSurcharges
          vTable = "payment_plan_surcharges"
          vAttr = "liable_product"
        Case cdbDataStructureConstants.cdbDataAppExplorerLinks
          vTable = "explorer_links"
          vAttr = "explorer_location"
        Case cdbDataStructureConstants.cbdMembershipEntitlementPriority
          vTable = "membership_entitlement"
          vAttr = "priority"

          '-------------------------------------------------------------------------------
          ' 11.3 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbExams
          vTable = "exam_units"
          vAttr = "exam_unit_id"
        Case cdbDataStructureConstants.cdbDataCPDPointsContactNumber
          vTable = "contact_cpd_points"
          vAttr = "contact_number"
        Case cdbDataStructureConstants.cdbRateModifiersSequence
          vTable = "rate_modifiers"
          vAttr = "sequence_number"
        Case cdbDataStructureConstants.cdbMembershipTypeCategories
          vTable = "membership_type_categories"
          vAttr = "membership_type"
        Case cdbDataStructureConstants.cdbAutoCreateCreditCustomer
          vTable = "fp_applications"
          vAttr = "auto_create_credit_customer"
        Case cdbDataStructureConstants.cdbDataInvoiceWithPayment
          vTable = "fp_applications"
          vAttr = "cheque_with_invoice"
        Case cdbDataStructureConstants.cdbLoanInterestRates
          vTable = "loan_interest_rates"
          vAttr = "loan_number"
        Case cdbDataStructureConstants.cdbBankTransactionsImport
          vTable = "bank_transactions"
          vAttr = "import_number"
        Case cdbDataStructureConstants.cdbReportUseSsrs
          vTable = "reports"
          vAttr = "use_ssrs"
        Case cdbDataStructureConstants.cdbUnpostedBatchMsgInPrint
          vTable = "fp_applications"
          vAttr = "unposted_batch_msg_in_print"
        Case cdbDataStructureConstants.cdbPaymentPlanHistoryDetails
          vTable = "payment_plan_history_details"
          vAttr = "order_number"
        Case cdbDataStructureConstants.cdbExamTraderDefaults
          vTable = "fp_applications"
          vAttr = "exam_session_code"
        Case cdbDataStructureConstants.cdbExamUnitCancellation
          vTable = "exam_booking_units"
          vAttr = "cancellation_reason"
        Case cdbDataStructureConstants.cdbInvoiceAdjustmentStatus
          vTable = "invoices"
          vAttr = "adjustment_status"

          '-------------------------------------------------------------------------------
          ' 12.1 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbWriteOffMissedPayments
          vTable = "financial_controls"
          vAttr = "payment_wo_overdue_days"
        Case cdbDataStructureConstants.cdbSystemMaintenance
          vTable = "amendment_history"
          vAttr = "table_maintenance"
        Case cdbDataStructureConstants.cdbBulkMailer
          vTable = "mailings"
          vAttr = "bulk_mailer_mailing"

          '-------------------------------------------------------------------------------
          ' 12.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbSuppressionSource
          vTable = "contact_suppressions"
          vAttr = "source"
        Case cdbDataStructureConstants.cdbModifierNextSequence
          vTable = "rate_modifiers"
          vAttr = "next_sequence_number"
        Case cdbDataStructureConstants.cdbResponseChannel
          vTable = "response_channels"
          vAttr = "response_channel"
        Case cdbDataStructureConstants.cdbEventMinimumBookings
          vTable = "event_booking_options"
          vAttr = "minimum_bookings"
        Case cdbDataStructureConstants.cdbDelegateSequenceNumber
          vTable = "delegates"
          vAttr = "sequence_number"
        Case cdbDataStructureConstants.cdbActivityDurationDays
          vTable = "activities"
          vAttr = "duration_days"
        Case cdbDataStructureConstants.cdbPasswordExpiry
          vTable = "registered_users"
          vAttr = "password_expiry_date"
        Case cdbDataStructureConstants.cdbAdvanceCMT
          vTable = "membership_controls"
          vAttr = "advanced_cmt"
        Case cdbDataStructureConstants.cdbAmendmentContactNumber
          vTable = "amendment_history"
          vAttr = "contact_number"
        Case cdbDataStructureConstants.cdbTnsHostedPayment
          vTable = "merchant_details"
          vAttr = "gateway_password"
        Case cdbDataStructureConstants.cdbExamExemptionModule
          vTable = "exam_student_exemptions"
          vAttr = "organisation_number"
        Case cdbDataStructureConstants.cdbDataContactAlerts
          vTable = "fp_applications"
          vAttr = "contact_alerts"

          '-------------------------------------------------------------------------------
          ' 13.1 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbCommunicationsUsage
          vTable = "communications"
          vAttr = "communication_usage"
        Case cdbDataStructureConstants.cdbConfigVersionInformation
          vTable = "config_names"
          vAttr = "amended_version"
        Case cdbDataStructureConstants.cdbMailingHistoryTopic
          vTable = "mailing_history"
          vAttr = "topic"
        Case cdbDataStructureConstants.cdbIbanBicNumbers
          vTable = "bank_transactions_netherlands"
          vAttr = "payers_iban_number"
        Case cdbDataStructureConstants.cdbJobFailureFlag
          vTable = "job_schedule"
          vAttr = "job_failed"
        Case cdbDataStructureConstants.cdbPopPaymentMethod
          vTable = "pop_payment_methods"
          vAttr = "pop_payment_method"
        Case cdbDataStructureConstants.cdbBacsEndToEndId
          vTable = "bacs_amendments"
          vAttr = "end_to_end_id"
        Case cdbDataStructureConstants.cdbPopBankAccount
          vTable = "pop_payment_methods"
          vAttr = "bank_account"
        Case cdbDataStructureConstants.cdbPayPlanChangesTermStartDate
          vTable = "payment_plan_changes"
          vAttr = "term_start_date"
        Case cdbDataStructureConstants.cdbFreeOfChangeBookingOption
          vTable = "event_booking_options"
          vAttr = "free_of_charge"
        Case cdbDataStructureConstants.cdbLockBranch
          vTable = "members"
          vAttr = "lock_branch"
        Case cdbDataStructureConstants.cdbUsePaymentProducedOn
          vTable = "purchase_order_controls"
          vAttr = "use_payment_produced_on_date"
        Case cdbDataStructureConstants.cdbPOPaymentReversals
          vTable = "po_payments_reversals"
          vAttr = "purchase_order_number"
        Case cdbDataStructureConstants.cdbRequiresPoPaymentType
          vTable = "purchase_order_types"
          vAttr = "requires_po_payment_type"
        Case cdbDataStructureConstants.cdbCancelOneYearGiftApm
          vTable = "orders"
          vAttr = "cancel_one_year_gift_apm"
        Case cdbDataStructureConstants.cdbLastExamDate
          vTable = "exam_student_unit_header"
          vAttr = "last_exam_date"
          '-------------------------------------------------------------------------------
          ' 13.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbExamsQualsRegistrationGrading
          vTable = "exam_assessment_languages"
          vAttr = "exam_assessment_language"
        Case cdbDataStructureConstants.cdbExamUnitLinkLongDescription
          vTable = "exam_unit_links"
          vAttr = "long_description"
        Case cdbDataStructureConstants.cdbWithholdExamResults
          vTable = "exam_result_unrestricted_depts"
          vAttr = "department"
        Case cdbDataStructureConstants.cdbDocumentLogLinks
          vTable = "document_log_links"
          vAttr = "document_link_id"
        Case cdbDataStructureConstants.cdbAccreditationMakeHistoric
          vTable = "exam_accreditation_statuses"
          vAttr = "make_historical"
        Case cdbDataStructureConstants.cdbExamStudyModes
          vTable = "study_modes"
          vAttr = "study_mode"
        Case cdbDataStructureConstants.cdbExamLoadResult
          vTable = "exam_controls"
          vAttr = "load_result"
        Case cdbDataStructureConstants.cdbAddressConfirmed
          vTable = "addresses"
          vAttr = "address_confirmed"
          '-------------------------------------------------------------------------------
          ' 14.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbEmailTemplates
          vTable = "communications_log"
          vAttr = "communications_number"
        Case cdbDataStructureConstants.cdbHistoricPoPaymentType
          vTable = "po_payment_types"
          vAttr = "is_historic"
        Case cdbDataStructureConstants.cdbTraderInvoicePrintUnpostedBatches
          vTable = "fp_applications"
          vAttr = "invoice_print_unposted_batches"
        Case cdbDataStructureConstants.cdbUserSchemes
          vTable = "user_schemes"
          vAttr = "user_scheme_id"
          '-------------------------------------------------------------------------------
          ' 14.3 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbCountriesIso3166CountryCodes
          vTable = "countries"
          vAttr = "iso_3166_alpha2_country_code"
          '-------------------------------------------------------------------------------
          ' 15.3 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbOrganisationGroupsViewInContactCard
          vTable = "organisation_groups"
          vAttr = "View_In_Contact_Card"
        Case cdbDataStructureConstants.cdbIso3166NumericCountryCode
          vTable = "countries"
          vAttr = "iso_3166_numeric_country_code"
          '-------------------------------------------------------------------------------
          ' 16.2 After here
          '-------------------------------------------------------------------------------
        Case cdbDataStructureConstants.cdbBankAccountRGBValue
          vTable = "bank_accounts"
          vAttr = "rgb_value"
        Case Else
          'Unknown parameter
      End Select

      If vTable.Length > 0 Then
        If mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoUnknown Then
          Select Case vCheck
            Case cdbDSConstants.cdbDSExists
              If IsAttributePresent(vTable, vAttr) Then
                mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoTrue
              Else
                mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoFalse
              End If
            Case cdbDSConstants.cdbDSType
              If IsAttributeType(vTable, vAttr, vDataType) Then
                mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoTrue
              Else
                mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoFalse
              End If
          End Select
        End If
        Return mvDSInfo(pType) = cdbDSInfoConstants.cdbDSInfoTrue
      End If
    End Function
    Public Function IsAttributePresent(ByVal pTable As String, ByVal pAttr As String) As Boolean
      If mvDBUpdated Then
        Return True
      Else
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("table_name", pTable)
        vWhereFields.Add("attribute_name", pAttr)
        Return New SQLStatement(Connection, "attribute_name", "maintenance_attributes", vWhereFields).GetValue.Length > 0
      End If
    End Function
    Public Function IsAttributeType(ByVal pTable As String, ByVal pAttr As String, ByVal pType As String) As Boolean
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("table_name", pTable)
      vWhereFields.Add("attribute_name", pAttr)
      Return New SQLStatement(Connection, "type", "maintenance_attributes", vWhereFields).GetValue = pType
    End Function

    Public ReadOnly Property FirstCustomFormNumber() As Integer
      Get
        Return FIRST_CUSTOM_FORM
      End Get
    End Property

    Public ReadOnly Property LastCustomFormNumber() As Integer
      Get
        Return LAST_CUSTOM_FORM
      End Get
    End Property

    Public ReadOnly Property JuniorAgeLimit() As Integer
      Get
        Dim vAge As String
        If mvJuniorAgeLimit = 0 Then
          vAge = GetConfig("jnr_age_limit")
          If vAge.Length > 0 Then
            mvJuniorAgeLimit = CInt(vAge)
          Else
            mvJuniorAgeLimit = 16
          End If
        End If
        JuniorAgeLimit = mvJuniorAgeLimit
      End Get
    End Property

    Public Function GetOutputDirectory(ByVal pDirectoryType As OutputDirectoryTypes) As String
      Dim vDirName As String = ""
      Select Case pDirectoryType
        Case OutputDirectoryTypes.scodtMailing
          vDirName = GetConfig("default_mailing_directory", "C:\contacts\mailings")
        Case OutputDirectoryTypes.scodtOutput
          vDirName = GetConfig("default_output_directory", "C:\contacts\output")
        Case OutputDirectoryTypes.scodtLogFiles
          vDirName = GetConfig("default_logfile_directory", "C:\contacts\logfiles")
        Case OutputDirectoryTypes.scodtAuditFiles
          vDirName = GetConfig("default_audit_directory", "C:\contacts\auditfiles")
      End Select
      Try
        If vDirName.Length > 0 Then
          If Not My.Computer.FileSystem.DirectoryExists(vDirName) Then RaiseError(DataAccessErrors.daeInvalidDirectory, vDirName)
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeInvalidDirectory, vDirName)
      End Try
      Return vDirName
    End Function

    Public Function GetDescription(ByVal pTableName As String, ByVal pAttributeName As String, ByVal pValue As String) As String
      Return GetDescription(pTableName, pAttributeName, pValue, New CDBFields)
    End Function
    Public Function GetDescription(ByVal pTableName As String, ByVal pAttributeName As String, ByVal pValue As String, ByVal pWhereFields As CDBFields) As String
      Dim vDescAttr As String
      Dim vAnsiJoins As New AnsiJoins()
      If pValue.Length > 0 Then
        If pAttributeName = "branch" Then
          vDescAttr = "name"
          pTableName = "branches b,organisations o"
          pWhereFields.AddJoin("b.organisation_number", "o.organisation_number")
        ElseIf pAttributeName = "geographical_region" Then
          vDescAttr = "name"
          pTableName = "geographical_regions gr"
          'pWhereFields.AddJoin("gr.organisation_number", "o.organisation_number")
          'Organisation is not mandatory for a geo. region
          vAnsiJoins.AddLeftOuterJoin("organisations o", "gr.organisation_number", "o.organisation_number")
        ElseIf pAttributeName = "logname" Then
          vDescAttr = "full_name"
          If pTableName = "ownership_group_users" Then
            pTableName = "ownership_group_users ogu, users u"
            pAttributeName = "ogu." & pAttributeName
          End If
        ElseIf pAttributeName = "organisation_number" Then
          vDescAttr = "name"
          If pTableName = "gaye_agencies" Then
            pTableName = "gaye_agencies ga, organisations o"
            pAttributeName = "ga." & pAttributeName
          End If
        ElseIf pAttributeName = "session_number" Then
          vDescAttr = "session_desc"
        ElseIf pAttributeName = "event_number" Then
          vDescAttr = "event_desc"
        ElseIf pAttributeName = "pis_number" Then
          vDescAttr = pAttributeName
        ElseIf pAttributeName = "web_page_user_control" Then
          vDescAttr = "control_title"
        ElseIf pAttributeName = "nominal_account" Then
          vDescAttr = "nominal_account"
        ElseIf pAttributeName = "db_name" Then
          vDescAttr = "db_name"
        ElseIf pAttributeName = "title" Then
          vDescAttr = "title"
        ElseIf pAttributeName = "surname_prefix" Then
          vDescAttr = "surname_prefix"
        ElseIf pAttributeName = "report_code" Then
          vDescAttr = "report_name"
        ElseIf pAttributeName = "honorific" Then
          vDescAttr = "honorific"
        ElseIf pAttributeName = "merchant_retail_number" Then
          vDescAttr = "merchant_details_desc"
        Else
          vDescAttr = pAttributeName & "_desc"
        End If
        If pWhereFields.ContainsKey(pAttributeName) Then
          pWhereFields(pAttributeName).Value = pValue
        Else
          pWhereFields.Add(pAttributeName, pValue)
        End If
        Dim vSQLStatement As New SQLStatement(Connection, vDescAttr, pTableName, pWhereFields, "", vAnsiJoins)
        Dim vDesc As String = vSQLStatement.GetValue
        If vDesc.Length = 0 Then
          'special cases where the description attribute may have null values
          'eg. logname/full name ; geographical_region/organisation_number
          Select Case pAttributeName
            Case "logname"
              vDesc = pValue
            Case "geographical_region"
              'Organisation is not mandatory for a geo. region
              vDesc = " " 'blank space
            Case "table_name"
              'none is a valid value while importing trader applications.
              If pValue = "none" Then vDesc = " " 'blank space
          End Select
        End If
        If vDescAttr = "rate_desc" AndAlso GetDataStructureInfo(cdbDataStructureConstants.cdbDataCurrencyCode) AndAlso vDesc.Length > 0 Then
          vDescAttr = "currency_code"
          Dim vTempStr As String = New SQLStatement(Connection, vDescAttr, pTableName, pWhereFields).GetValue
          If vTempStr <> GetControlValue(cdbControlConstants.cdbControlCurrencyCode) Then
            vDesc = vDesc & " (" & vTempStr & ")"
          End If
        End If
        Return vDesc
      Else
        Return ""
      End If
    End Function

    Public Function GetConfig(ByVal pOption As String) As String
      If mvConfigs Is Nothing Then InitConfigs()
      Dim vValue As String = ""
      If mvConfigs.ContainsKey(pOption) Then
        Return mvConfigs(pOption).ToString
      Else
        Return ""
      End If
    End Function
    Public Function GetConfig(ByVal pOption As String, ByVal pDefault As String) As String
      Dim vValue As String = GetConfig(pOption)
      If vValue.Length = 0 Then
        Return pDefault
      Else
        Return vValue
      End If
    End Function

    Public Function GetConfigOption(ByVal pOption As String) As Boolean
      If GetConfig(pOption).ToUpper.StartsWith("Y") Then
        Return True
      Else
        Return False
      End If
    End Function
    Public Function GetConfigOption(ByVal pOption As String, ByVal pDefault As Boolean) As Boolean
      Dim vValue As String = GetConfig(pOption)
      If vValue.Length = 0 Then
        Return pDefault
      Else
        If vValue.ToUpper.StartsWith("Y") Then
          Return True
        Else
          Return False
        End If
      End If
    End Function

    Public Function GetCachedControlNumber(ByVal pNumberType As CachedControlNumberTypes) As Integer
      Dim vType As String = ""
      Select Case pNumberType
        Case CachedControlNumberTypes.ccnJournal
          vType = "CJ"          'Ensure 2 characters
        Case CachedControlNumberTypes.ccnPaymentSchedule
          vType = "SP"
        Case CachedControlNumberTypes.ccnTimesheet
          vType = "TK"
        Case CachedControlNumberTypes.ccnAddress
          vType = "A"
        Case CachedControlNumberTypes.ccnAddressLink
          vType = "AL"
        Case CachedControlNumberTypes.ccnContact
          vType = "C"
        Case CachedControlNumberTypes.ccnPosition
          vType = "PN"
        Case CachedControlNumberTypes.ccnExamMarkingBatchDetail
          vType = "XMD"
      End Select
      Return GetCachedControlNumber(vType)
    End Function

    Public Function GetCachedControlNumber(ByVal pType As String) As Integer
#If DEBUG Then
      Debug.Print("GetCachedControlNumber request for {0} type numbers made by {1}.", pType, GetCallingMethodDetails)
#End If
      If mvCachedControlNumbers IsNot Nothing AndAlso mvCachedControlNumbers.ContainsKey(pType) Then
        Dim vCCN As CachedControlNumber = mvCachedControlNumbers(pType)
        Return vCCN.NextControlNumber
      End If
      Return GetControlNumber(pType)
    End Function

    Public Function GetControlNumber(ByVal pNumberType As String) As Integer
      Return GetControlNumber(pNumberType, 0, False)
    End Function
    Public Function GetControlNumber(ByVal pNumberType As String, ByVal pNumber As Integer) As Integer
      Return GetControlNumber(pNumberType, pNumber, False)
    End Function
    Public Function GetControlNumber(ByVal pNumberType As String, ByVal pNumber As Integer, ByVal pRemote As Boolean) As Integer
      'Return the next control number of a particular type from the control numbers table
      'This routine will generate an error if the control number is not found
      'If the control number record is locked then an error will be returned
#If DEBUG Then
      Debug.Print("GetControlNumber request for {0} type '{1}' numbers made by {2}.", pNumber, pNumberType, GetCallingMethodDetails)
#End If
      Dim vControlNumber As Integer
      'First check if we have read the check_digits_on_numbers config option
      If Not mvCheckDigitsRead Then
        mvCheckDigitsOnNumbers = GetConfig("check_digits_on_numbers")
        mvCheckDigitMethod = GetConfig("check_digit_method")
        mvCheckDigitsRead = True
      End If
      Dim vRemoteUser As Boolean = mvUser.RemoteUser
      Dim vLogname As String = If(String.IsNullOrWhiteSpace(mvUser.Logname), "dbinit", mvUser.Logname)

      If pNumberType.Substring(0, 1) = "?" Then
        'Just add the check digit, we already have a block of control numbers
        pNumberType = pNumberType.Substring(1)

        If (mvCheckDigitsOnNumbers.IndexOf("|" & pNumberType & "|", 0) + 1) > 0 Then
          vControlNumber = GenerateCheckDigit(pNumber)
        Else
          vControlNumber = pNumber
        End If
      Else
        Dim vNumberIncrement As Integer
        Dim vGetBlock As Boolean
        If pNumberType.Substring(0, 1) = "+" Then
          pNumberType = pNumberType.Substring(1)
          vGetBlock = True
          If pNumber > 0 Then
            vNumberIncrement = pNumber
          Else
            vNumberIncrement = 100
          End If
        Else
          If pRemote Then
            vGetBlock = True
            vNumberIncrement = pNumber
          Else
            vNumberIncrement = 1
          End If
        End If
        If vRemoteUser Then
          vControlNumber = GetRemoteControlNumber(pNumberType, vLogname, vNumberIncrement)
        Else
          vControlNumber = GetStandardControlNumber(pNumberType, vLogname, vNumberIncrement)
        End If
        If Not vGetBlock AndAlso (mvCheckDigitsOnNumbers.IndexOf("|" & pNumberType & "|", 0) + 1) > 0 Then
          'if vGetBlock, return the next control number NO check digit use the ? option to get the check digit
          vControlNumber = GenerateCheckDigit(vControlNumber)
        End If
      End If
      Return vControlNumber
    End Function

    Private Function GetStandardControlNumber(ByVal pNumberType As String, ByVal vLogname As String, ByVal vNumberIncrement As Integer) As Integer
      Dim vRecordSet As CDBRecordSet
      'Loop until we sucessfully update the control number
      Dim vRetries As Integer
      Dim vGotControlNumber As Boolean
      Dim vControlNumber As Integer

      Dim vConnection As CDBConnection = Connection()
      If vConnection.InTransaction Then vConnection = GetConnection("CONTROLNUMBER")
      Dim vTableName As String = "control_numbers"
      Dim vSelectWhereFields As New CDBFields
      vSelectWhereFields.Add("control_number_type", pNumberType)
      Dim vUpdateWhereFields As New CDBFields
      vUpdateWhereFields.Add("control_number_type", pNumberType)
      vUpdateWhereFields.Add("control_number")
      Dim vUpdateFields As New CDBFields
      vUpdateFields.Add("amended_by", vLogname)
      vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      vUpdateFields.Add("control_number")
      Dim vIncrementCount As Integer
      Try
        Do
          vIncrementCount = 0
          Dim vSqlStatement As New SQLStatement(vConnection, "control_number", vTableName, vSelectWhereFields)
          vRecordSet = vSqlStatement.GetRecordSet
          If vRecordSet.Fetch() Then
            vControlNumber = vRecordSet.Fields(1).LongValue
            vRecordSet.CloseRecordSet()
            If Not vGotControlNumber Then
              vUpdateWhereFields("control_number").Value = CStr(vControlNumber)
              vUpdateFields("control_number").Value = CStr(vControlNumber + vNumberIncrement)
              'If someone else has nicked the control number the update will fail and it will loop
              If vConnection.UpdateRecords(vTableName, vUpdateFields, vUpdateWhereFields, False) > 0 Then
                vGotControlNumber = True
              Else
                If vRetries > 0 And vNumberIncrement = 1 Then
                  For vIncrementCount = 1 To MAX_RETRIES
                    vControlNumber += 1
                    vUpdateWhereFields("control_number").Value = CStr(vControlNumber)
                    vUpdateFields("control_number").Value = CStr(vControlNumber + vNumberIncrement)
                    If vConnection.UpdateRecords(vTableName, vUpdateFields, vUpdateWhereFields, False) > 0 Then
                      vGotControlNumber = True
                      Exit For
                    End If
                  Next
                End If
              End If
            End If
          Else
            vRecordSet.CloseRecordSet()
            RaiseError(DataAccessErrors.daeMissingControlNumber, pNumberType)
          End If
          vRetries += 1
        Loop While vGotControlNumber = False And vRetries < MAX_RETRIES
      Catch ex As Exception
        Debug.Print("Control number get failed: '{0}'.", ex.Message)
      Finally
        If vGotControlNumber = False Then
          Debug.Print(String.Format("Get Control Number Type {0} Failed {1} Retries {2} Increments", pNumberType, vRetries, vIncrementCount))
          RaiseError(DataAccessErrors.daeControlNumber, pNumberType)
        End If
        If vRetries > 1 OrElse vIncrementCount > 1 Then
          Debug.Print(String.Format("Get Control Number Type {0} - {1} Retries {2} Increments", pNumberType, vRetries, vIncrementCount))
        End If
      End Try
      Return vControlNumber
    End Function

    Private Function GetRemoteControlNumber(ByVal pNumberType As String, ByVal vLogname As String, ByVal vNumberIncrement As Integer) As Integer
      Dim vRecordSet As CDBRecordSet
      'Loop until we sucessfully update the control number
      Dim vRetries As Integer
      Dim vGotControlNumber As Boolean
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vControlNumber As Integer

      Dim vConnection As CDBConnection = Connection()
      If Connection.InTransaction AndAlso (pNumberType = "CJ" OrElse pNumberType = "SP") Then vConnection = GetConnection("CONTROLNUMBER")
      Do
        Dim vTableName As String
        vWhereFields.Clear()
        vWhereFields.Add("control_number_type", pNumberType)
        vTableName = "remote_control_numbers"
        vWhereFields.Add("logname", vLogname)
        vWhereFields.Add("active_block", "Y")
        vRecordSet = New SQLStatement(vConnection, "control_number,maximum_control_number", vTableName, vWhereFields).GetRecordSet
        If vRecordSet.Fetch() Then
          vControlNumber = vRecordSet.Fields(1).LongValue
          vRecordSet.CloseRecordSet()
          If vControlNumber = vRecordSet.Fields(2).LongValue Then
            'Run out of active numbers for this user
            'Set the current one inactive
            vUpdateFields.Clear()
            vUpdateFields.Add("active_block", CDBField.FieldTypes.cftCharacter, "N")
            vUpdateFields.Add("amended_by", CDBField.FieldTypes.cftCharacter, vLogname)
            vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)

            vWhereFields.Clear()
            vWhereFields.Add("control_number_type", CDBField.FieldTypes.cftCharacter, pNumberType)
            vWhereFields.Add("logname", CDBField.FieldTypes.cftCharacter, vLogname)
            vWhereFields.Add("active_block", "Y")
            vWhereFields.Add("control_number", vControlNumber)

            If vConnection.UpdateRecords(vTableName, vUpdateFields, vWhereFields, False) > 0 Then
              vGotControlNumber = True
            End If

            'Set any available one active
            vUpdateFields(1).Value = "Y"

            vWhereFields.Clear()
            vWhereFields.Add("control_number_type", pNumberType)
            vWhereFields.Add("logname", vLogname)
            vWhereFields.Add("active_block", "N")
            vWhereFields.Add("control_number", CDBField.FieldTypes.cftLong, "maximum_control_number", CDBField.FieldWhereOperators.fwoLessThan) 'The long here is a kludge
            If vConnection.UpdateRecords(vTableName, vUpdateFields, vWhereFields, False) < 1 Then
              'Ignore an error at this point - we have a number to use
              'Next time a number is required an error will be generated
            End If
          End If
          If Not vGotControlNumber Then
            vWhereFields.Clear()
            vWhereFields.Add("control_number_type", pNumberType)
            vWhereFields.Add("logname", vLogname)
            vWhereFields.Add("active_block", "Y")
            vWhereFields.Add("control_number", vControlNumber)

            vUpdateFields.Clear()
            vUpdateFields.Add("control_number", vControlNumber + vNumberIncrement)
            vUpdateFields.Add("amended_by", vLogname)
            vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)

            'If someone else has nicked the control number the update will fail and it will loop
            If vConnection.UpdateRecords(vTableName, vUpdateFields, vWhereFields, False) > 0 Then
              vGotControlNumber = True
            End If
          End If
        Else
          vRecordSet.CloseRecordSet()
          RaiseError(DataAccessErrors.daeMissingControlNumber, pNumberType)
        End If
        vRetries += 1
      Loop While vGotControlNumber = False And vRetries < MAX_RETRIES
      If vGotControlNumber = False Then RaiseError(DataAccessErrors.daeControlNumber, pNumberType)
      Return vControlNumber
    End Function

    Private Function GenerateCheckDigit(ByVal pNumber As Integer) As Integer
      'Will receive a number as a parameter and will use a config option to
      'determine if the method to use to calculate the check digit.
      'Data Import assumes check digit will always be an ADDITIONAL digit at the end
      'of the control number. (ajh 6/7/99)
      Dim vNumber As String = pNumber.ToString
      Select Case mvCheckDigitMethod
        Case "M11+4"
          Dim vLen As Integer = vNumber.Length
          Dim vCounter As Integer = 1
          Dim vAccumulator As Integer = 0
          Do While (vLen > 0 AndAlso vCounter < 6)
            vAccumulator += (CInt(vNumber.Substring(vLen - 1, 1)) * (vCounter + 1))
            vLen = vLen - 1
            vCounter = vCounter + 1
          Loop
          If vLen > 0 Then vAccumulator += (CInt(vNumber.Substring(0, vLen)) * 7)
          Dim vCheckDigit As Integer = CInt(11 - (vAccumulator Mod 11))
          If vCheckDigit = 10 Then vCheckDigit = 4
          If vCheckDigit = 11 Then vCheckDigit = 0
          Return pNumber * 10 + vCheckDigit
        Case Else
          Return pNumber 'Leave the number as it was
      End Select
    End Function

    Public Function GetControlBool(ByVal pType As cdbControlConstants) As Boolean
      GetControlBool = GetControlValue(pType) = "Y"
    End Function

    Public Function GetControlValue(ByVal pType As cdbControlConstants) As String
      Dim vRecordSet As CDBRecordSet
      Dim vTable As String = ""

      Select Case pType
        Case cdbControlConstants.cdbControlDefConVatCat To cdbControlConstants.cdbControlOneReversalOnly
          If mvFinancialControls = False Then
            vTable = "financial_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            mvControls(cdbControlConstants.cdbControlDefConVatCat) = "I"
            mvControls(cdbControlConstants.cdbControlDefOrgVatCat) = "I"
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlDefConVatCat) = .Item("default_contact_vat_cat").Value
                mvControls(cdbControlConstants.cdbControlDefOrgVatCat) = .Item("default_organisation_vat_cat").Value
                mvControls(cdbControlConstants.cdbControlCCReason) = .Item("cc_reason").Value
                mvControls(cdbControlConstants.cdbControlDDReason) = .Item("dd_reason").Value
                mvControls(cdbControlConstants.cdbControlSOReason) = .Item("so_reason").Value
                mvControls(cdbControlConstants.cdbControlOReason) = .Item("o_reason").Value
                mvControls(cdbControlConstants.cdbControlExpiredOrdersCancellationReason) = .Item("expired_orders_canc_reason").Value
                mvControls(cdbControlConstants.cdbControlFirstClaimTransactionType) = .Item("first_claim_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlSOActivity) = .FieldExists("standing_order_activity").Value
                mvControls(cdbControlConstants.cdbControlSOActivityValue) = .FieldExists("standing_order_activity_value").Value
                mvControls(cdbControlConstants.cdbControlDDActivity) = .FieldExists("direct_debit_activity").Value
                mvControls(cdbControlConstants.cdbControlDDActivityValue) = .FieldExists("direct_debit_activity_value").Value
                mvControls(cdbControlConstants.cdbControlCCCAActivity) = .FieldExists("ccca_activity").Value
                mvControls(cdbControlConstants.cdbControlCCCAActivityValue) = .FieldExists("ccca_activity_value").Value
                mvControls(cdbControlConstants.cdbControlReverseTransType) = .FieldExists("reversal_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlDistributorActivity) = .Item("distributor_activity").Value
                mvControls(cdbControlConstants.cdbControlStockInterface) = .FieldExists("stock_interface").Value
                mvControls(cdbControlConstants.cdbControlAutoSODefaultDays) = .FieldExists("auto_so_default_days").Value
                mvControls(cdbControlConstants.cdbControlAccountsInterface) = .Item("accounts_interface").Value
                mvControls(cdbControlConstants.cdbControlMaximumOnPickingList) = .Item("maximum_on_picking_list").Value
                mvControls(cdbControlConstants.cdbControlSODefaultProduct) = .FieldExists("default_so_product").Value
                mvControls(cdbControlConstants.cdbControlBankAccount) = .FieldExists("bank_account").Value
                mvControls(cdbControlConstants.cdbControlInAdvanceTransType) = .FieldExists("in_advance_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlInAdvancePaymentMethod) = .FieldExists("in_advance_payment_method").Value
                mvControls(cdbControlConstants.cdbControlCurrencyCode) = .FieldExists("currency_code").Value
                mvControls(cdbControlConstants.cdbControlDespatchTransactionType) = .FieldExists("despatch_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlDespatchPaymentMethod) = .FieldExists("despatch_payment_method").Value
                mvControls(cdbControlConstants.cdbControlPaymentReason) = .Item("payment_reason").Value
                mvControls(cdbControlConstants.cdbControlAdjustmentTransType) = .FieldExists("adjustment_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlInMemoriamRelationship) = .Item("inmemoriam_relationship").Value
                mvControls(cdbControlConstants.cdbControlAnonymousContactNumber) = .FieldExists("anonymous_contact_number").LongValue.ToString
                mvControls(cdbControlConstants.cdbControlAnonCardBankAccount) = .FieldExists("anonymous_card_bank_account").Value
                mvControls(cdbControlConstants.cdbControlAnonCardProduct) = .FieldExists("anonymous_card_product").Value
                mvControls(cdbControlConstants.cdbControlAnonCardRate) = .FieldExists("anonymous_card_rate").Value
                mvControls(cdbControlConstants.cdbControlAnonCardSource) = .FieldExists("anonymous_card_source").Value
                mvControls(cdbControlConstants.cdbControlAnonCardDistributionCode) = .FieldExists("anonymous_card_dist_code").Value
                mvControls(cdbControlConstants.cdbControlBACSNewDDSource) = .FieldExists("bacs_new_dd_source").Value
                mvControls(cdbControlConstants.cdbControlAutomaticRenewalDateChangeReason) = .FieldExists("auto_renewal_change_reason").Value
                mvControls(cdbControlConstants.cdbControlHoldingContactNumber) = .FieldExists("holding_contact_number").LongValue.ToString
                mvControls(cdbControlConstants.cdbControlRoundingErrorProduct) = .FieldExists("rounding_error_product").Value
                mvControls(cdbControlConstants.cdbControlMerchantRetailNumber) = .FieldExists("merchant_retail_number").Value
                mvControls(cdbControlConstants.cdbControlAutoPayClaimDateMethod) = .FieldExists("auto_pay_claim_date_method").Value
                mvControls(cdbControlConstants.cdbControlApplyIncentiveFreePeriod) = .FieldExists("apply_incentive_free_period").Value
                mvControls(cdbControlConstants.cdbControlDeceasedStatus) = .Item("deceased_status").Value
                mvControls(cdbControlConstants.cdbControlAdjustOriginalProductCost) = .FieldExists("adjust_original_product_cost").Value
                mvControls(cdbControlConstants.cdbControlDefaultDDText1) = .FieldExists("default_dd_text1").Value
                mvControls(cdbControlConstants.cdbControlOneOffPPCancelReason) = .FieldExists("one_off_pp_cancel_reason").Value
                mvControls(cdbControlConstants.cdbControlDefaultDDBatchCategory) = .FieldExists("default_dd_batch_category").Value
                mvControls(cdbControlConstants.cdbControlMailingHistorySearchDays) = .FieldExists("mailing_history_search_days").Value
                mvControls(cdbControlConstants.cdbControlMailingHistorySearchSegmentType) = .FieldExists("mailing_history_segment_type").Value
                mvControls(cdbControlConstants.cdbControlNewContactSource) = .FieldExists("new_contact_source").Value
                mvControls(cdbControlConstants.cdbControlExistingContactSource) = .FieldExists("existing_contact_source").Value
                mvControls(cdbControlConstants.cdbControlDefaultProduct) = .FieldExists("default_product").Value
                mvControls(cdbControlConstants.cdbControlDefaultRate) = .FieldExists("default_rate").Value
                mvControls(cdbControlConstants.cdbControlDefaultFundPayType) = .FieldExists("fundraising_payment_type").Value
                mvControls(cdbControlConstants.cdbControlDefaultFundStatus) = .FieldExists("fundraising_status").Value
                mvControls(cdbControlConstants.cdbControlLockFundRequest) = .FieldExists("lock_fundraising_request").Value
                mvControls(cdbControlConstants.cdbControlSCPURL) = .FieldExists("scp_url").Value
                mvControls(cdbControlConstants.cdbControlSCPAPIVersion) = .FieldExists("scp_api_version").Value
                mvControls(cdbControlConstants.cdbControlVPCURL) = .FieldExists("vpc_url").Value
                mvControls(cdbControlConstants.cdbControlVPCAPIVersion) = .FieldExists("vpc_api_version").Value
                mvControls(cdbControlConstants.cdbControlBacsMsgFileFormat) = .FieldExists("bacs_msg_file_format").Value
                mvControls(cdbControlConstants.cdbLoanCapitalisationDate) = .FieldExists("loan_capitalisation_date").Value
                mvControls(cdbControlConstants.cdbAutoSOAcceptAsFull) = .FieldExists("auto_so_accept_as_full").Value
                mvControls(cdbControlConstants.cdbControlPaymentWOOverDueDays) = .FieldExists("payment_wo_overdue_days").Value
                mvControls(cdbControlConstants.cdbControlCashBookBatchLimit) = .FieldExists("cash_book_batch_limit").Value
                mvControls(cdbControlConstants.cdbControlBacsUserNumber) = .FieldExists("bacs_user_number").Value
                mvControls(cdbControlConstants.cdbControlOneOffClaimTransactionType) = .FieldExists("one_off_claim_transaction_type").Value
                mvControls(cdbControlConstants.cdbControlUseRenewalDateForRateMod) = .FieldExists("use_renewal_date_for_rate_mod").Value
                If .FieldExists("account_validation_type").Value.Length > 0 Then
                  mvControls(cdbControlConstants.cdbControlAccountValidationType) = .Item("account_validation_type").Value
                  mvControls(cdbControlConstants.cdbControlAccountValidationURL) = .Item("account_validation_url").Value
                Else
                  'Columns do not exist so use defaults
                  mvControls(cdbControlConstants.cdbControlAccountValidationType) = "UC"
                  mvControls(cdbControlConstants.cdbControlAccountValidationURL) = String.Empty
                End If
                mvControls(cdbControlConstants.cdbControlReceiptPrintStdDocument) = .FieldExists("receipt_print_std_document").Value
                mvControls(cdbControlConstants.cdbControlOneReversalOnly) = .FieldExists("one_reversal_only").Value
              End With
              mvFinancialControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
            If mvControls(cdbControlConstants.cdbControlInMemoriamRelationship).Length > 0 Then
              Dim vRecordSet2 As CDBRecordSet = New SQLStatement(Connection, "complimentary_relationship", "relationships", New CDBField("relationship", mvControls(cdbControlConstants.cdbControlInMemoriamRelationship))).GetRecordSet
              If vRecordSet2.Fetch() Then
                mvControls(cdbControlConstants.cdbControlInMemoriamCompRelationship) = vRecordSet2.Fields(1).Value
              Else
                RaiseError(DataAccessErrors.daeInMemoriamError)
              End If
              vRecordSet2.CloseRecordSet()
            Else
              mvControls(cdbControlConstants.cdbControlInMemoriamCompRelationship) = ""
            End If
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlGAStatus To cdbControlConstants.cdbControlRetainRegUserPasswords
          If mvContactControls = False Then
            vTable = "contact_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlGAStatus) = .Item("gone_away_status").Value
                mvControls(cdbControlConstants.cdbControlGAMailingSupp) = .Item("gone_away_mailing_suppression").Value
                mvControls(cdbControlConstants.cdbControlDefMailingSupp) = .FieldExists("default_mailing_suppression").Value
                mvControls(cdbControlConstants.cdbControlCLIDevice) = .FieldExists("cli_device").Value
                mvControls(cdbControlConstants.cdbControlTYLSupressionExclusionList) = .Item("tyl_suppression_exclusion_list").Value.Replace(" ", "") 'Remove spaces
                mvControls(cdbControlConstants.cdbControlParentRelationship) = .FieldExists("contact_parent_relationship").Value
                mvControls(cdbControlConstants.cdbControlDespatchReason) = .FieldExists("reason_for_despatch").Value
                mvControls(cdbControlConstants.cdbControlMaxPermittedDaysActivity) = .FieldExists("max_permitted_days_activity").Value
                mvControls(cdbControlConstants.cdbControlMaxPermittedDaysActivityVal) = .FieldExists("max_permitted_days_act_val").Value
                mvControls(cdbControlConstants.cdbControlDaysRemainingActivity) = .FieldExists("curr_days_remaining_activity").Value
                mvControls(cdbControlConstants.cdbControlDaysRemainingActivityVal) = .FieldExists("curr_days_remaining_act_val").Value
                mvControls(cdbControlConstants.cdbControlQualifyingPositionActivity) = .FieldExists("qualifying_position_activity").Value
                mvControls(cdbControlConstants.cdbControlQualifyingPositionActivityVal) = .FieldExists("qualifying_position_act_val").Value
                mvControls(cdbControlConstants.cdbControlAnonymousContactStatus) = .FieldExists("anonymous_contact_status").Value
                If .Exists("start_of_day") Then
                  mvControls(cdbControlConstants.cdbControlStartOfDay) = IIf(.Item("start_of_day").Value.Length > 0, .Item("start_of_day").Value, "09:00").ToString
                  mvControls(cdbControlConstants.cdbControlEndOfDay) = IIf(.Item("end_of_day").Value.Length > 0, .Item("end_of_day").Value, "17:30").ToString
                  mvControls(cdbControlConstants.cdbControlStartOfLunch) = IIf(.Item("start_of_lunch").Value.Length > 0, .Item("start_of_lunch").Value, "13:00").ToString
                  mvControls(cdbControlConstants.cdbControlEndOfLunch) = IIf(.Item("end_of_lunch").Value.Length > 0, .Item("end_of_lunch").Value, "14:00").ToString
                Else
                  'Use default values
                  mvControls(cdbControlConstants.cdbControlStartOfDay) = "09:00"
                  mvControls(cdbControlConstants.cdbControlEndOfDay) = "17:30"
                  mvControls(cdbControlConstants.cdbControlStartOfLunch) = "13:00"
                  mvControls(cdbControlConstants.cdbControlEndOfLunch) = "14:00"
                End If
                mvControls(cdbControlConstants.cdbControlDirectDevice) = .FieldExists("direct_device").Value
                mvControls(cdbControlConstants.cdbControlSwitchboardDevice) = .FieldExists("switchboard_device").Value
                mvControls(cdbControlConstants.cdbControlFaxDevice) = .FieldExists("fax_device").Value
                mvControls(cdbControlConstants.cdbControlMobileDevice) = .FieldExists("mobile_device").Value
                mvControls(cdbControlConstants.cdbControlEmailDevice) = .FieldExists("email_device").Value
                mvControls(cdbControlConstants.cdbControlWebDevice) = .FieldExists("web_device").Value
                mvControls(cdbControlConstants.cdbControlPositionActivityGroup) = .FieldExists("position_activity_group").Value
                mvControls(cdbControlConstants.cdbControlPositionRelationshipGroup) = .FieldExists("position_relationship_group").Value
                mvControls(cdbControlConstants.cdbControlClosedOrganisationStatus) = .FieldExists("closed_organisation_status").Value
                If Not .Exists("merge_use_oldest_source") Then
                  mvControls(cdbControlConstants.cdbControlMergeUseOldestSource) = "N"
                Else
                  mvControls(cdbControlConstants.cdbControlMergeUseOldestSource) = .FieldExists("merge_use_oldest_source").Value
                End If
                mvControls(cdbControlConstants.cdbControlEmailCaseSensitive) = .FieldExists("email_case_sensitive").Value
                If .Exists("retain_reg_user_passwords") Then
                  mvControls(cdbControlConstants.cdbControlRetainRegUserPasswords) = .FieldExists("retain_reg_user_passwords").Value
                Else
                  mvControls(cdbControlConstants.cdbControlRetainRegUserPasswords) = "1"
                End If
              End With
              mvContactControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlSponsorActivity To cdbControlConstants.cdbControlMoveDDMemberCancelReason
          If mvMembershipControls = False Then
            vTable = "membership_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlBranchProduct) = .Item("product").Value
                mvControls(cdbControlConstants.cdbControlSponsorActivity) = .Item("sponsor_activity").Value
                mvControls(cdbControlConstants.cdbControlSponsorActivityValue) = .Item("sponsor_activity_value").Value
                mvControls(cdbControlConstants.cdbControlRealToJointLink) = .Item("real_to_joint_relationship").Value
                mvControls(cdbControlConstants.cdbControlRealToRealLink) = .Item("real_to_real_relationship").Value
                mvControls(cdbControlConstants.cdbControlBranchParent) = .FieldExists("branch_parent_relationship").Value
                mvControls(cdbControlConstants.cdbControlReasonForDespatch) = .Item("reason_for_despatch").Value
                mvControls(cdbControlConstants.cdbControlTypeChangeCancelReason) = .Item("type_change_cancel_reason").Value
                mvControls(cdbControlConstants.cdbControlNonAddress) = .Item("non_address_number").Value
                mvControls(cdbControlConstants.cdbControlOverAgeCancelReason) = .Item("over_age_canc_reason").Value
                mvControls(cdbControlConstants.cdbControlNonPaymentCancelReason) = .Item("non_payment_canc_reason").Value
                mvControls(cdbControlConstants.cdbControlOneYearGiftedGroupReason) = .Item("one_year_gifted_group_reason").Value
                mvControls(cdbControlConstants.cdbControlMembershipSalesGroup) = .Item("membership_sales_group").Value
                mvControls(cdbControlConstants.cdbControlOneYearGiftsAutoReason) = .Item("one_year_gifts_auto_reason").Value
                mvControls(cdbControlConstants.cdbControlGiftMemberMaxJuniorAge) = .FieldExists("gift_member_max_junior_age").Value
                mvControls(cdbControlConstants.cdbControlCMTProportionBalance) = .FieldExists("cmt_calc_proportional_balance").Value
                mvControls(cdbControlConstants.cdbControlAutoPayAdvancePeriod) = .FieldExists("auto_pay_advance_period").Value
                mvControls(cdbControlConstants.cdbControlMemberOrganisationGroup) = .FieldExists("organisation_group").Value
                mvControls(cdbControlConstants.cdbControlCMTMakeRefund) = .FieldExists("cmt_make_refund").Value
                mvControls(cdbControlConstants.cdbControlFMTEffectiveDays) = If(.FieldExists("fmt_effective_days").Value.Length > 0, .Item("fmt_effective_days").Value, "14")
                mvControls(cdbControlConstants.cdbControlAdvancedCMT) = .FieldExists("advanced_cmt").Value
                mvControls(cdbControlConstants.cdbControlRemoveZeroBalancePpdLines) = .FieldExists("remove_zero_balance_ppd_lines").Value
                mvControls(cdbControlConstants.cdbControlAddMemberCurrentAddress) = .FieldExists("add_member_current_address").Value
                mvControls(cdbControlConstants.cdbControlMoveDDMemberCancelReason) = .FieldExists("move_dd_member_cancel_reason").Value
              End With
              mvMembershipControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlCVActivity To cdbControlConstants.cdbControlCVMinimumCovenantPeriod
          If mvCovenantControls = False Then
            vTable = "covenant_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlCVActivity) = .Item("activity").Value
                mvControls(cdbControlConstants.cdbControlCVActivityValue) = .Item("activity_value").Value
                mvControls(cdbControlConstants.cdbControlCVGiftAidMinimum) = .Item("gift_aid_minimum").Value
                mvControls(cdbControlConstants.cdbControlCVReasonForDespatch) = .Item("reason_for_despatch").Value
                mvControls(cdbControlConstants.cdbControlDeedReceivedProduct) = .Item("deed_received_product").Value
                mvControls(cdbControlConstants.cdbControlDeedReceivedRate) = .Item("deed_received_rate").Value
                mvControls(cdbControlConstants.cdbControlCVMinimumCovenantPeriod) = .Item("minimum_covenant_period").Value
              End With
              mvCovenantControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlLGActivity To cdbControlConstants.cdbControlLGRelationshipList
          If mvLegacyControls = False Then
            vTable = "legacy_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlLGActivity) = .Item("activity").Value
                mvControls(cdbControlConstants.cdbControlLGActivityValue) = .Item("activity_value").Value
                mvControls(cdbControlConstants.cdbControlLGJointRelationship) = .Item("joint_legator_relationship").Value
                mvControls(cdbControlConstants.cdbControlLGAssetActivity) = .Item("asset_activity").Value
                mvControls(cdbControlConstants.cdbControlLGResidualBequestType) = .Item("residual_bequest_type").Value
                mvControls(cdbControlConstants.cdbControlLGConditionalBequestSubType) = .Item("conditional_bequest_sub_type").Value
                mvControls(cdbControlConstants.cdbControlLGSpecificBequestType) = .Item("specific_bequest_type").Value
                mvControls(cdbControlConstants.cdbControlLGRelationshipList) = .Item("legacy_relationships_list").Value
              End With
              mvLegacyControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlCSTransType To cdbControlConstants.cdbControlCSUnderPayRate
          If mvCreditSalesControls = False Then
            vTable = "credit_sales_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlCSTransType) = .Item("transaction_type").Value
                mvControls(cdbControlConstants.cdbControlCSPayMethod) = .Item("payment_method").Value
                mvControls(cdbControlConstants.cdbControlCSBankAccount) = .Item("bank_account").Value
                mvControls(cdbControlConstants.cdbControlCSCreditTransType) = .FieldExists("sundry_credit_trans_type").Value
                mvControls(cdbControlConstants.cdbControlCSSource) = .FieldExists("source").Value
                mvControls(cdbControlConstants.cdbControlCSOverPayProduct) = .FieldExists("invoice_over_payment_product").Value
                mvControls(cdbControlConstants.cdbControlCSOverPayRate) = .FieldExists("invoice_over_payment_rate").Value
                mvControls(cdbControlConstants.cdbControlCSUnderPayProduct) = .FieldExists("invoice_under_payment_product").Value
                mvControls(cdbControlConstants.cdbControlCSUnderPayRate) = .FieldExists("invoice_under_payment_rate").Value
              End With
              mvCreditSalesControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlsDerivedSuppression To cdbControlConstants.cdbControlPositiveAdjustmentRate
          If mvMarketingControls = False Then
            vTable = "marketing_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlsDerivedSuppression) = .Item("derived_contact_mailing_supp").Value
                mvControls(cdbControlConstants.cdbControlJointSuppression) = .Item("joint_contact_mailing_supp").Value
                mvControls(cdbControlConstants.cdbControlDerivedToJointLink) = .Item("derived_to_joint_relationship").Value
                mvControls(cdbControlConstants.cdbControlDerivedToDerivedLink) = .Item(Connection.DBAttrName("derived_to_derived_relationship")).Value
                mvControls(cdbControlConstants.cdbControlEarliestDonation) = .Item("earliest_donation").Value
                mvControls(cdbControlConstants.cdbControlMinimumAge) = .Item("minimum_age").Value
                mvControls(cdbControlConstants.cdbControlDepartment) = .Item("department").Value
                mvControls(cdbControlConstants.cdbControlLastContactNumber) = .FieldExists("last_contact_number").Value
                mvControls(cdbControlConstants.cdbControlLastMembershipNumber) = .FieldExists("last_membership_number").Value
                mvControls(cdbControlConstants.cdbControlLastBatchNumber) = .FieldExists("last_batch_number").Value
                mvControls(cdbControlConstants.cdbControlTransactionType) = .Item("transaction_type").Value
                mvControls(cdbControlConstants.cdbControlVATRate) = .Item("vat_rate").Value
                mvControls(cdbControlConstants.cdbControlIncludeHistoric) = .Item("include_historical_addresses").Value
                mvControls(cdbControlConstants.cdbControlCriteriaSet) = .Item("criteria_set").Value
                mvControls(cdbControlConstants.cdbControlCollectionsRegionType) = .FieldExists("collections_region_type").Value
                mvControls(cdbControlConstants.cdbControlDefaultCollectorStatus) = .FieldExists("default_collector_status").Value
                mvControls(cdbControlConstants.cdbControlNegativeAdjustmentProduct) = .FieldExists("negative_adjustment_product").Value
                mvControls(cdbControlConstants.cdbControlNegativeAdjustmentRate) = .FieldExists("negative_adjustment_rate").Value
                mvControls(cdbControlConstants.cdbControlPositiveAdjustmentProduct) = .FieldExists("positive_adjustment_product").Value
                mvControls(cdbControlConstants.cdbControlPositiveAdjustmentRate) = .FieldExists("positive_adjustment_rate").Value
              End With
              mvMarketingControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlGAYEDonorProduct To cdbControlConstants.cdbControlPreTaxOtherMatchedRate
          If mvGAYEControls = False Then
            vTable = "gaye_controls"
            If GetConfigOption("option_gaye") Then
              vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
              If vRecordSet.Fetch() Then
                With vRecordSet.Fields
                  mvControls(cdbControlConstants.cdbControlGAYEDonorProduct) = .Item("donor_product").Value
                  mvControls(cdbControlConstants.cdbControlGAYEDonorRate) = .Item("donor_rate").Value
                  mvControls(cdbControlConstants.cdbControlGAYEEmployerProduct) = .FieldExists("employer_product").Value
                  mvControls(cdbControlConstants.cdbControlGAYEEmployerRate) = .FieldExists("employer_rate").Value
                  mvControls(cdbControlConstants.cdbControlGAYEGovernmentProduct) = .FieldExists("government_product").Value
                  mvControls(cdbControlConstants.cdbControlGAYEGovernmentRate) = .FieldExists("government_rate").Value
                  mvControls(cdbControlConstants.cdbControlGAYEAdminFeeProduct) = .FieldExists("admin_fee_product").Value
                  mvControls(cdbControlConstants.cdbControlGAYEAdminFeeRate) = .FieldExists("admin_fee_rate").Value
                  mvControls(cdbControlConstants.cdbControlGAYEActivity) = .FieldExists("gaye_pledge_activity").Value
                  mvControls(cdbControlConstants.cdbControlGAYEActivityValue) = .FieldExists("gaye_pledge_activity_value").Value
                  mvControls(cdbControlConstants.cdbControlGAYEGovernmentPercentage) = .Item("government_percentage").Value
                  mvControls(cdbControlConstants.cdbControlGAYESource) = .Item("source").Value
                  mvControls(cdbControlConstants.cdbControlGAYENiDataSource) = .Item("ni_data_source").Value
                  mvControls(cdbControlConstants.cdbControlGAYEDistributionCode) = .FieldExists("distribution_code").Value
                  mvControls(cdbControlConstants.cdbControlGAYEAgencyRelationship) = .FieldExists("employer_agency_relationship").Value
                  mvControls(cdbControlConstants.cdbControlGAYEPayrollRelationship) = .FieldExists("employer_payroll_relationship").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGDonorProduct) = .FieldExists("post_tax_donor_product").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGDonorRate) = .FieldExists("post_tax_donor_rate").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGEmployerProduct) = .FieldExists("post_tax_employer_product").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGEmployerRate) = .FieldExists("post_tax_employer_rate").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGPledgeSource) = .FieldExists("post_tax_pledge_source").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGDistributionCode) = .FieldExists("post_tax_distribution_code").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGEmployerActivity) = .FieldExists("post_tax_employer_activity").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGEmployerActivityValue) = .FieldExists("post_tax_employer_act_value").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGEmplrPayrollRelationship) = .FieldExists("post_tax_empr_payroll_relation").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGPledgeActivity) = .FieldExists("post_tax_pledge_activity").Value
                  mvControls(cdbControlConstants.cdbControlPostTaxPGPledgeActivityValue) = .FieldExists("post_tax_pledge_activity_value").Value
                  mvControls(cdbControlConstants.cdbControlPreTaxOtherMatchedProduct) = .FieldExists("other_matched_product").Value
                  mvControls(cdbControlConstants.cdbControlPreTaxOtherMatchedRate) = .FieldExists("other_matched_rate").Value
                End With
                mvGAYEControls = True
              Else
                RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
              End If
              vRecordSet.CloseRecordSet()
            Else
              mvGAYEControls = True
            End If
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlStockReasonInitial To cdbControlConstants.cdbControlStockWarehouseTransferReason 'cdbControlStockImportReason
          If mvStockMovementControls = False OrElse pType = cdbControlConstants.cdbControlStockReasonInitial Then
            vTable = "stock_movement_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlStockReasonInitial) = .FieldExists("stock_initial_value_reason").Value
                mvControls(cdbControlConstants.cdbControlStockReasonSale) = .FieldExists("stock_product_sale_reason").Value
                mvControls(cdbControlConstants.cdbControlStockReasonReversal) = .FieldExists("stock_reversal_reason").Value
                mvControls(cdbControlConstants.cdbControlStockReasonShortFall) = .FieldExists("stock_picking_shortfall_reason").Value
                mvControls(cdbControlConstants.cdbControlStockReasonBackOrder) = .FieldExists("stock_bo_allocation_reason").Value
                mvControls(cdbControlConstants.cdbControlStockReasonAwaitBackOrder) = .FieldExists("stock_awaiting_bo_reason").Value
                mvControls(cdbControlConstants.cdbControlStockProcessLock) = .FieldExists("process_lock").Value
                mvControls(cdbControlConstants.cdbControlStockImportReason) = .FieldExists("stock_import_reason").Value
                mvControls(cdbControlConstants.cdbControlStockPackTransferReason) = .FieldExists("stock_pack_transfer_reason").Value
                mvControls(cdbControlConstants.cdbControlStockWarehouseTransferReason) = .FieldExists("stock_warehouse_xfer_reason").Value
              End With
              mvStockMovementControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlGAMergeCancellationReason To cdbControlConstants.cdbControlGAAdjustmentText 'Gift Aid
          If mvGiftAidControls = False Then
            vTable = "gift_aid_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlGAMergeCancellationReason) = .FieldExists("merge_cancellation_reason").Value
                mvControls(cdbControlConstants.cdbControlGAClaimFileFormat) = .FieldExists("claim_file_format").Value
                mvControls(cdbControlConstants.cdbControlGACharityTaxStatus) = .FieldExists("charity_tax_status").Value
                mvControls(cdbControlConstants.cdbControlGAAccountingPeriodStart) = .FieldExists("accounting_period_start").Value
                mvControls(cdbControlConstants.cdbControlGAMinimumAnnualDonation) = .FieldExists("minimum_annual_donation").Value
                mvControls(cdbControlConstants.cdbControlGATaxYearStart) = .FieldExists("tax_year_start").Value
                mvControls(cdbControlConstants.cdbControlGASubmitterContact) = .FieldExists("submitter_contact").Value
                mvControls(cdbControlConstants.cdbControlGAAdjustmentText) = .FieldExists("adjustment_text").Value
              End With
              mvGiftAidControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlEmailUseHeaderTemplate To cdbControlConstants.cdbControlEmailForceSmtpAddress

          If mvEmailControls = False Then
            vTable = "email_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlEmailUseHeaderTemplate) = .Item("use_header_template").Value
                mvControls(cdbControlConstants.cdbControlEmailHeaderTemplate) = .Item("header_template").Value
                mvControls(cdbControlConstants.cdbControlEmailForceSmtpAddress) = .Item("force_smtp_address").Value
              End With
              mvEmailControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
            mvEmailControls = True
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlCoreStockControl
          If mvControls(cdbControlConstants.cdbControlCoreStockControl).Length = 0 Then
            mvControls(cdbControlConstants.cdbControlCoreStockControl) = "N"
            If (GetControlValue(cdbControlConstants.cdbControlStockInterface) <> "Y") Then
              If Connection.GetCount("stock_movement_controls", Nothing) > 0 Then mvControls(cdbControlConstants.cdbControlCoreStockControl) = "Y"
            End If
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlMaxBagWeight
          If mvMailsortControls = False Then
            vTable = "mailsort_controls"
            vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
            If vRecordSet.Fetch() Then
              With vRecordSet.Fields
                mvControls(cdbControlConstants.cdbControlMaxBagWeight) = .Item("maximum_bag_weight").Value
              End With
              mvMailsortControls = True
            Else
              RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
            End If
            vRecordSet.CloseRecordSet()
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlEventpisBankAccount To cdbControlConstants.cdbControlEventPISPerDelegate
          If Not mvEventControls Then
            If GetDataStructureInfo(cdbDataStructureConstants.cdbDataEventPIS) Then
              vTable = "event_controls"
              vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
              If vRecordSet.Fetch() Then
                With vRecordSet.Fields
                  mvControls(cdbControlConstants.cdbControlEventpisBankAccount) = .Item("pis_bank_account").Value
                  mvControls(cdbControlConstants.cdbControlEventPISPerDelegate) = .Item("pis_per_delegate").Value
                End With
                mvEventControls = True
              Else
                RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
              End If
              vRecordSet.CloseRecordSet()
            End If
            mvEventControls = True
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlExamGeographicalRegionType To cdbControlConstants.cdbControlExamCertNumberSuffix
          If Not mvExamControls Or pType = cdbControlConstants.cdbControlExamCertNumber Then
            If GetDataStructureInfo(cdbDataStructureConstants.cdbExams) Then
              vTable = "exam_controls"
              vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
              If vRecordSet.Fetch() Then
                With vRecordSet.Fields
                  mvControls(cdbControlConstants.cdbControlExamGeographicalRegionType) = .Item("geographical_region_type").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionCompany) = .Item("exemption_company").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionSource) = .Item("exemption_source").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionCreditCategory) = .Item("exemption_credit_category").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionGrade) = .FieldExists("exemption_grade").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionResult) = .FieldExists("exemption_result").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionOrgActivity) = .FieldExists("exemption_org_activity").Value
                  mvControls(cdbControlConstants.cdbControlExamExemptionOrgActivityValue) = .FieldExists("exemption_org_activity_value").Value
                  mvControls(cdbControlConstants.cdbControlExamRecordGradeChangeHistory) = .FieldExists("record_grade_change_history").Value
                  If GetDataStructureInfo(cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
                    mvControls(cdbControlConstants.cdbControlExamCentreAccreditation) = .FieldExists("centre_accreditation").Value
                    mvControls(cdbControlConstants.cdbControlExamUnitAccreditation) = .FieldExists("unit_accreditation").Value
                    mvControls(cdbControlConstants.cdbControlExamCentreUnitAccreditation) = .FieldExists("centre_unit_accreditation").Value
                    mvControls(cdbControlConstants.cdbControlExamGradingMethod) = .FieldExists("grading_method").Value
                    mvControls(cdbControlConstants.cdbControlExamCertNumberPrefix) = .FieldExists("exam_cert_number_prefix").Value
                    mvControls(cdbControlConstants.cdbControlExamCertNumber) = .FieldExists("exam_cert_number").Value
                    mvControls(cdbControlConstants.cdbControlExamCertNumberSuffix) = .FieldExists("exam_cert_number_suffix").Value
                  End If
                  If GetDataStructureInfo(cdbDataStructureConstants.cdbExamLoadResult) Then mvControls(cdbControlConstants.cdbControlExamLoadResult) = .FieldExists("load_result").Value
                End With
                mvExamControls = True
              Else
                RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
              End If
              vRecordSet.CloseRecordSet()
            End If
            mvEventControls = True
          End If
          Return mvControls(pType)

        Case cdbControlConstants.cdbControlBulkMailerLoginId To cdbControlConstants.cdbControlBulkMailerPassword
          If Not mvBulkMailerControls Then
            If GetDataStructureInfo(cdbDataStructureConstants.cdbBulkMailer) Then
              vTable = "bulk_mailer_controls"
              vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
              If vRecordSet.Fetch() Then
                With vRecordSet.Fields
                  mvControls(cdbControlConstants.cdbControlBulkMailerLoginId) = .Item("login_id").Value
                  mvControls(cdbControlConstants.cdbControlBulkMailerPassword) = .Item("password").Value
                End With
                mvBulkMailerControls = True
              Else
                RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
              End If
              vRecordSet.CloseRecordSet()
            End If
            mvEventControls = True
          End If
          Return mvControls(pType)
        Case cdbControlConstants.cdbControlPopDefPaymentMethod To cdbControlConstants.cdbControlUsePaymentProducedOn
          If Not mvPurchaseOrderControls Then
            If GetDataStructureInfo(cdbDataStructureConstants.cdbBulkMailer) Then
              vTable = "purchase_order_controls"
              vRecordSet = New SQLStatement(Connection, "*", vTable).GetRecordSet
              If vRecordSet.Fetch() Then
                With vRecordSet.Fields
                  mvControls(cdbControlConstants.cdbControlPopDefPaymentMethod) = .Item("pop_def_payment_method").Value
                  mvControls(cdbControlConstants.cdbControlUsePaymentProducedOn) = .Item("use_payment_produced_on_date").Value
                End With
                mvPurchaseOrderControls = True
              Else
                RaiseError(DataAccessErrors.daeMissingControlRecord, vTable)
              End If
              vRecordSet.CloseRecordSet()
            End If
            mvEventControls = True
          End If
          Return mvControls(pType)
      End Select
      Return ""
    End Function

    Public Function GetUserConfigs() As CDBDataTable
      If mvConfigs Is Nothing Then InitConfigs()
      Dim vDT As New CDBDataTable
      vDT.AddColumnsFromList("ConfigName,ConfigValue")
      For Each vItem As String In mvConfigs.Keys
        vDT.AddRowFromItems(vItem, mvConfigs(vItem).ToString)
      Next
      Return vDT
    End Function

    Private Sub InitConfigs()
      mvConfigs = New SortedList
      If mvInitialisingDatabase Then
        If Not Connection.TableExists("config") Then Exit Sub
      End If
      Dim vWhereFields As New CDBFields
      vWhereFields.AddJoin("c.config_name", "cn.config_name")
      vWhereFields.AddClientDeptLogname(mvClientCode, mvUser.Department, mvUser.Logname)
      If System.Diagnostics.Debugger.IsAttached OrElse ((mvClientCode IsNot Nothing) AndAlso ((String.Compare(mvClientCode, "care", True) = 0 OrElse String.Compare(mvClientCode, "aqa", True) = 0))) Then
        vWhereFields.Add("config_group", "'UNCH'", CDBField.FieldWhereOperators.fwoNotIn)
      Else
        vWhereFields.Add("config_group", "'UNCH','UNUN'", CDBField.FieldWhereOperators.fwoNotIn)
      End If
      Dim vSQL As New SQLStatement(Connection, "c.config_name,config_value", "config c,config_names cn", vWhereFields)
      vSQL.SetOrderByClientDeptLogname("c.config_name")
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      With vRecordSet
        Dim vName As String
        Dim vValue As String
        Dim vSkip As String = ""
        Dim vIgnore As Boolean
        Do While .Fetch()
          vName = .Fields(1).Value
          vValue = .Fields(2).Value
          If vSkip.Length > 0 Then
            If vName.StartsWith(vSkip) Then
              vIgnore = True
            Else
              vIgnore = False
              vSkip = ""
            End If
          End If
          If Not vIgnore Then
            'Add an item to the list only if not already there
            If Not mvConfigs.ContainsKey(vName) Then
              If vName.StartsWith("profile") Then
                'convert any commas to pipes - used because jet sometimes ignores pipes on insert?
                vValue = vValue.Replace(",", "|")
              End If
              mvConfigs.Add(vName, vValue)
            Else
              If vName.StartsWith("tabs") Then vSkip = vName.Substring(0, vName.Length - 2)
            End If
          End If
        Loop
        .CloseRecordSet()
      End With
      ResetAuditStyle() 'Setup Auditing
    End Sub

    Private Function GetJournalType(ByVal pType As JournalTypes) As String
      Select Case pType
        Case JournalTypes.jnlContact
          Return "CONT"
        Case JournalTypes.jnlOrganisation
          Return "ORG"
        Case JournalTypes.jnlAddress
          Return "ADD"
        Case JournalTypes.jnlRelationship
          Return "REL"
        Case JournalTypes.jnlActivity
          Return "ACT"
        Case JournalTypes.jnlSuppression
          Return "SUP"
        Case JournalTypes.jnlPosition
          Return "POS"
        Case JournalTypes.jnlRole
          Return "ROLE"
        Case JournalTypes.jnlDocument
          Return "DOC"
        Case JournalTypes.jnlAction
          Return "ACTN"
        Case JournalTypes.jnlEvent
          Return "EVNT"
        Case JournalTypes.jnlAccomodation
          Return "ACCO"
        Case JournalTypes.jnlMeeting
          Return "MEET"
        Case JournalTypes.jnlAppointment
          Return "APPT"
        Case JournalTypes.jnlDonation
          Return "DON"
        Case JournalTypes.jnlStockItem
          Return "STCK"
        Case JournalTypes.jnlEventBooking
          Return "EVPY"
        Case JournalTypes.jnlAccomodationBooking
          Return "ACPY"
        Case JournalTypes.jnlServiceBooking
          Return "SVPY"
        Case JournalTypes.jnlOtherProduct
          Return "PAY"
        Case JournalTypes.jnlMember
          Return "MEM"
        Case JournalTypes.jnlCovenant
          Return "COV"
        Case JournalTypes.jnlStandingOrder
          Return "SO"
        Case JournalTypes.jnlDirectDebit
          Return "DD"
        Case JournalTypes.jnlCreditCard
          Return "CCCA"
        Case JournalTypes.jnlPayPlan
          Return "PP"
        Case JournalTypes.jnlMailing
          Return "MAIL"
        Case JournalTypes.jnlActionActioner
          Return "ACTA"
        Case JournalTypes.jnlActionManager
          Return "ACTM"
        Case JournalTypes.jnlActionRelated
          Return "ACTR"
        Case JournalTypes.jnlGiftAidDeclaration
          Return "GAD"
        Case JournalTypes.jnlPledge
          Return "PLDG"
        Case JournalTypes.jnlGoneAway
          Return "GAWA"
        Case JournalTypes.jnlDirectDebitMaintenance
          Return "DDM"
        Case JournalTypes.jnlCreditCardMaintenance
          Return "CCM"
        Case JournalTypes.jnlStandingOrderMaintenance
          Return "SOM"
        Case JournalTypes.jnlMemberMaintenance
          Return "MEMM"
        Case JournalTypes.jnlCovenantMaintenance
          Return "COVM"
        Case JournalTypes.jnlPayPlanMaintenance
          Return "PPM"
        Case JournalTypes.jnlPayPlanPaymentSchedule
          Return "PPPS"
        Case JournalTypes.jnlNumber
          Return "NUMB"
        Case JournalTypes.jnlCPDCycles
          Return "CPDC"
        Case JournalTypes.jnlCPDPoints
          Return "CPDP"
        Case JournalTypes.jnlIrishAppropriateCertificate
          Return "IAC"
        Case JournalTypes.jnlLogin
          Return "LOG"
        Case JournalTypes.jnlRegisteredUser
          Return "REG"
        Case JournalTypes.jnlWebProductSale
          Return "WPS"
        Case JournalTypes.jnlWebPaymentPlanPayment
          Return "WPPP"
        Case JournalTypes.jnlDirectory
          Return "DIR"
        Case JournalTypes.jnlDownload
          Return "DNL"
        Case Else
          Return ""
      End Select
    End Function

    Public ReadOnly Property JournalActive() As Boolean
      Get
        Return GetConfigOption("option_journal", True)
      End Get
    End Property

    Public Function GetJournalOperation(ByVal pJournalType As String, ByVal pOperation As String) As JournalOperation
      If Not mvJournalInitialised Then InitJournal()
      If mvOptionJournal AndAlso mvJournalList.ContainsKey(JournalOperation.GetKey(pJournalType, pOperation)) Then
        Return mvJournalList(JournalOperation.GetKey(pJournalType, pOperation))
      Else
        Return Nothing
      End If
    End Function

    Private Sub InitJournal()
      Dim vRecordSet As CDBRecordSet
      Dim vType As String
      Dim vDesc As String
      Dim vDays As Integer
      Dim vJO As JournalOperation

      mvOptionJournal = GetConfigOption("option_journal", True)
      If mvOptionJournal Then
        mvJournalList = New CollectionList(Of JournalOperation)
        Dim vFields As String = "journal_type, journal_type_desc, active_days, operation1, operation1_log, operation1_text, operation2, operation2_log, operation2_text, operation3, operation3_log, operation3_text"
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("operation1_log", "Y").WhereOperator = CDBField.FieldWhereOperators.fwoOpenBracket
        vWhereFields.Add("operation2_log", "Y").WhereOperator = CDBField.FieldWhereOperators.fwoOR
        vWhereFields.Add("operation3_log", "Y").WhereOperator = CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket
        Dim vSQLStatement As New SQLStatement(Connection, vFields, "journal_types", vWhereFields, "journal_type")
        vRecordSet = vSQLStatement.GetRecordSet
        With vRecordSet
          While .Fetch()
            vType = .Fields(1).Value
            vDesc = .Fields(2).Value
            vDays = .Fields(3).IntegerValue
            If .Fields("operation1_log").Bool Then
              vJO = New JournalOperation(vType, vDesc, vDays, .Fields("operation1").Value, .Fields("operation1_text").Value)
              mvJournalList.Add(vJO.Key, vJO)
            End If
            If .Fields("operation2_log").Bool Then
              vJO = New JournalOperation(vType, vDesc, vDays, .Fields("operation2").Value, .Fields("operation2_text").Value)
              mvJournalList.Add(vJO.Key, vJO)
            End If
            If .Fields("operation3_log").Bool Then
              vJO = New JournalOperation(vType, vDesc, vDays, .Fields("operation3").Value, .Fields("operation3_text").Value)
              mvJournalList.Add(vJO.Key, vJO)
            End If
          End While
          .CloseRecordSet()
        End With
      End If
      mvJournalInitialised = True
    End Sub

    Public Function AddJournalRecord(ByVal pType As JournalTypes, ByVal pOperation As JournalOperations, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer) As Integer
      Return AddJournalRecord(pType, pOperation, pContactNumber, pAddressNumber, 0, 0, 0, 0, 0)
    End Function

    Public Function AddJournalRecord(ByVal pType As JournalTypes, ByVal pOperation As JournalOperations, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pSelect1 As Integer) As Integer
      Return AddJournalRecord(pType, pOperation, pContactNumber, pAddressNumber, pSelect1, 0, 0, 0, 0)
    End Function
    Public Function AddJournalRecord(ByVal pType As JournalTypes, ByVal pOperation As JournalOperations, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pSelectName As String) As Integer
      Return AddJournalRecord(pType, pOperation, pContactNumber, pAddressNumber, 0, 0, 0, 0, 0, pSelectName)
    End Function

    Public Function AddJournalRecord(ByVal pType As JournalTypes, ByVal pOperation As JournalOperations, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pSelect1 As Integer, ByVal pSelect2 As Integer, ByVal pSelect3 As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, Optional ByVal pSelectName As String = "") As Integer
      'Checks if option journal is set and if we are logging this type of Journal
      'if so then adds the journal record
#If DEBUG Then
      Debug.Print("AddJournalRecord request for journal type '{0}', operation '{1}', contact '{2}', address '{3}' Select1 '{4}', Select2 '{5}', Select3 '{6}', Batch Number '{7}', Transaction Number '{8}', Select Name '{9}' made by {10}.", pType, [Enum].GetName(GetType(JournalOperations), pOperation), pContactNumber, pAddressNumber, pSelect1, pSelect2, pSelect3, pBatchNumber, pTransactionNumber, pSelectName, GetCallingMethodDetails)
#End If
      If Not mvJournalInitialised Then InitJournal()
      If mvOptionJournal Then
        Dim vOperation As String
        Select Case pOperation
          Case JournalOperations.jnlActive
            vOperation = "active"
          Case JournalOperations.jnlCancel
            vOperation = "cancel"
          Case JournalOperations.jnlComplete
            vOperation = "complete"
          Case JournalOperations.jnlDelete
            vOperation = "delete"
          Case JournalOperations.jnlInsert
            vOperation = "insert"
          Case JournalOperations.jnlUpdate
            vOperation = "update"
          Case JournalOperations.jnlReinstate
            vOperation = "reinstate"
          Case JournalOperations.jnlSearch
            vOperation = "search"
          Case JournalOperations.jnlView
            vOperation = "view"
          Case JournalOperations.jnlDowloaded
            vOperation = "downloaded"
          Case Else
            vOperation = ""
        End Select

        If mvJournalList.ContainsKey(GetJournalType(pType) & "-" & vOperation) Then
          Dim vJournalNumber As Integer = GetCachedControlNumber(CachedControlNumberTypes.ccnJournal)
          Dim vInsertFields As New CDBFields
          With vInsertFields
            .Add("contact_journal_number", vJournalNumber)
            .Add("journal_type", GetJournalType(pType))
            .Add("journal_time", CDBField.FieldTypes.cftTime, TodaysDateAndTime)
            .Add("journal_by", mvUser.UserID)
            .Add("operation", vOperation)
            .Add("contact_number", pContactNumber)
            .Add("address_number", pAddressNumber)
            If pSelect1 > 0 Then .Add("select_1", pSelect1)
            If pSelect2 > 0 Then .Add("select_2", pSelect2)
            If pSelect3 > 0 Then .Add("select_3", pSelect3)
            If GetDataStructureInfo(cdbDataStructureConstants.cdbJournalSelectName) Then
              If Not String.IsNullOrEmpty(pSelectName) Then .Add("select_name", pSelectName)
            End If
            If pBatchNumber > 0 Then .Add("batch_number", pBatchNumber)
            If pTransactionNumber > 0 Then .Add("transaction_number", pTransactionNumber)
          End With
          Connection.InsertRecord("contact_journals", vInsertFields)
          Return vJournalNumber
        End If
      End If
    End Function

    Public Sub SetConfig(ByVal pOption As String, ByRef pValue As String, Optional ByVal pSystem As Boolean = False, Optional ByVal pUpdateDB As Boolean = True)
      'Sets a value into the configuration table for the current user
      'if a record exists for this config item and this user it is updated
      'otherwise a new record is added

      If mvConfigs Is Nothing OrElse mvConfigs.Count = 0 Then InitConfigs()
      If mvConfigs.ContainsKey(pOption) Then mvConfigs.Remove(pOption)
      mvConfigs.Add(pOption, pValue)

      Dim vClient As String = ""
      Dim vLogname As String = ""
      If Not pSystem Then 'Is this a system wide config
        vClient = mvClientCode
        vLogname = mvUser.Logname
      End If

      If pUpdateDB Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("config_name", pOption)
        vWhereFields.Add("logname", vLogname)
        vWhereFields.Add("client", vClient)
        Dim vSQL As String = "SELECT config_value, amended_by, amended_on FROM config WHERE " & Connection.WhereClause(vWhereFields)
        Dim vRecordSet As CDBRecordSet = Connection.GetRecordSet(vSQL)
        With vRecordSet
          Dim vFields As New CDBFields
          vFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
          vFields.Add("amended_by", If(InitialisingDatabase, "dbinit", mvUser.Logname))
          If .Fetch() Then
            'Do update
            vFields.Add("config_value", pValue.Replace("|", ","))
            Connection.UpdateRecords("config", vFields, vWhereFields)
          Else
            'Do Insert
            vFields.Add("config_name", pOption)
            vFields.Add("config_value", pValue.Replace("|", ","))
            vFields.Add("logname", vLogname)
            vFields.Add("department", "")
            vFields.Add("client", vClient)
            Connection.InsertRecord("config", vFields)
          End If
          .CloseRecordSet()
        End With
      End If
    End Sub

    Public Sub SetupRemoteUsers()
      Dim vUsers As CDBRecordSet
      Dim vMissing As CDBRecordSet
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vInsertFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vLogname As String
      Dim vType As String
      Dim vLastType As String = String.Empty
      Dim vLastLogname As String = String.Empty
      Dim vControlNumber As Long
      Dim vBlockSize As Integer

      If mvUser.RemoteUser Then Exit Sub

      vBlockSize = IntegerValue(GetConfig("control_number_block_size"))
      If vBlockSize < 1 Then vBlockSize = 100

      'First insert control numbers for any new users or new control numbers
      vInsertFields.Add("control_number_type")
      vInsertFields.Add("logname")
      vInsertFields.Add("active_block")
      vInsertFields.Add("control_number", CDBField.FieldTypes.cftLong)
      vInsertFields.Add("maximum_control_number", CDBField.FieldTypes.cftLong)
      vInsertFields.Add("amended_by", mvUser.Logname)
      vInsertFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)

      'Check for each remote user
      vUsers = New SQLStatement(Connection, "logname", "users", New CDBField("remote_user", "Y")).GetRecordSet
      While vUsers.Fetch()
        vLogname = vUsers.Fields("logname").Value
        vMissing = Connection.GetRecordSet("SELECT control_number_type FROM control_numbers WHERE control_number_type NOT IN (SELECT DISTINCT control_number_type FROM remote_control_numbers WHERE logname = '" & vLogname & "')")
        While vMissing.Fetch()
          vType = vMissing.Fields("control_number_type").Value
          vInsertFields("control_number_type").Value = vType
          vInsertFields("logname").Value = vLogname
          'First add the primary set
          vInsertFields("active_block").Value = "Y"
          vControlNumber = GetControlNumber(vType, vBlockSize, True)
          vInsertFields("control_number").Value = vControlNumber.ToString
          vInsertFields("maximum_control_number").Value = (vControlNumber + (vBlockSize - 1)).ToString
          Connection.InsertRecord("remote_control_numbers", vInsertFields)
          'Now add the backup set
          vInsertFields("active_block").Value = "N"
          vControlNumber = GetControlNumber(vType, vBlockSize, True)
          vInsertFields("control_number").Value = vControlNumber.ToString
          vInsertFields("maximum_control_number").Value = (vControlNumber + (vBlockSize - 1)).ToString
          Connection.InsertRecord("remote_control_numbers", vInsertFields)
        End While
        vMissing.CloseRecordSet()
      End While
      vUsers.CloseRecordSet()

      'Now check for any non active blocks that have run out of numbers
      vUpdateFields.Add("control_number", CDBField.FieldTypes.cftLong)
      vUpdateFields.Add("maximum_control_number", CDBField.FieldTypes.cftLong)
      vUpdateFields.Add("active_block")
      vUpdateFields.Add("amended_by", mvUser.Logname)
      vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)

      vWhereFields.Add("control_number_type")
      vWhereFields.Add("logname")
      vWhereFields.Add("active_block", "N")
      vWhereFields.Add("control_number", CDBField.FieldTypes.cftLong)

      Dim vCDBFields As New CDBFields
      vCDBFields.AddJoin("control_number", "maximum_control_number")
      vCDBFields.Add("active_block", "N")
      vRecordSet = New SQLStatement(Connection, "control_number_type,logname,control_number", "remote_control_numbers", vCDBFields, "control_number_type,logname,control_number DESC").GetRecordSet
      While vRecordSet.Fetch()
        vType = vRecordSet.Fields("control_number_type").Value
        vLogname = vRecordSet.Fields("logname").Value
        vWhereFields("control_number_type").Value = vType
        vWhereFields("logname").Value = vLogname
        vWhereFields("control_number").Value = vRecordSet.Fields("control_number").Value

        vControlNumber = GetControlNumber(vType, vBlockSize, True)
        vUpdateFields("control_number").Value = vControlNumber.ToString
        vUpdateFields("maximum_control_number").Value = (vControlNumber + (vBlockSize - 1)).ToString
        If vType = vLastType AndAlso vLogname = vLastLogname Then
          vUpdateFields("active_block").Value = "Y"
        Else
          vUpdateFields("active_block").Value = "N"
          vLastType = vType
          vLastLogname = vLogname
        End If
        Connection.UpdateRecords("remote_control_numbers", vUpdateFields, vWhereFields)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Function IsCountryUK(ByVal pCountry As String) As Boolean
      Select Case pCountry
        Case "UK", "BFPO", "NI", "GBA", "GBG", "GBJ", "GBM"
          Return True
        Case ""
          Return True
        Case Else
          If mvUKCountries Is Nothing Then
            mvUKCountries = New CDBParameters
            Dim vRS As CDBRecordSet = New SQLStatement(Connection, "country,country_desc", "countries", New CDBField("uk", "Y")).GetRecordSet
            Do While vRS.Fetch
              mvUKCountries.Add(vRS.Fields(1).Value, CDBField.FieldTypes.cftCharacter, vRS.Fields(2).Value)
            Loop
            vRS.CloseRecordSet()
          End If
          If mvUKCountries.Exists(pCountry) Then
            Return True
          Else
            Return False
          End If
      End Select
    End Function

    Public Function IsDefaultCountryUK() As Boolean
      Return IsCountryUK(DefaultCountry)
    End Function

    Public ReadOnly Property MaxRetries() As Integer
      Get
        Return MAX_RETRIES
      End Get
    End Property

    Public Function OptionEnabled(ByVal pItemID As String) As Boolean
      If pItemID.StartsWith("CDAE") Then    'All events sub-menus
        Return GetConfigOption("option_events", False)
      Else
        Select Case pItemID
          '----------------------------------------------------------------------------------
          ' ADMIN MANAGER ITEMS HERE
          '----------------------------------------------------------------------------------
          Case "AMAMTA"                       'Admin manager - Trader Applications
            Return GetConfigOption("option_trader")
          Case "AMSMGU"                       'Admin manager - Update Government Regions
            Return GetConfigOption("cd_use_government_regions")
          Case "AMSMCL"
            Return False
            '----------------------------------------------------------------------------------
            ' SYSTEM MANAGER ITEMS HERE
            '----------------------------------------------------------------------------------
          Case "SMAMDO"
            Return False             'Deferred to 5.6
            'Case "SMFMPL", "SMFPCL", "SMFPCR"   'Started in 5.7 but may be deferred to 5.8
            '  Return False 'GetDataStructureInfo(cdbDataCollections)
          Case "SMFSSV"                       'Stock Valuation Report (v6.1)
            Return GetDataStructureInfo(cdbDataStructureConstants.cdbDataProductCosts)
            '----------------------------------------------------------------------------------
            ' CDB AND SMART CLIENT ITEMS HERE
            '----------------------------------------------------------------------------------
          Case "CDTMSM"                       'Selection Manager
            Return GetConfigOption("option_general_mailing", True)
          Case "CDAMSD"                       'Standard Document Maintenance
            Return GetConfigOption("option_standard_documents", True)
          Case "CDTMUC"                       'User Comments
            Return GetConfigOption("option_user_comments", True)
          Case "CDTMDD"                       'Document Distributor
            Return GetConfigOption("option_document_distribution", False)
          Case "CDAMPV"                       'Postcode Validation
            Return GetConfigOption("option_postcode_validation", False) And Postcoder.PostcoderType <> Postcoder.PostcoderTypes.pctNone
          Case "CDAMQP", "CDAEPE"             'Questionnaire Processing, Process Eval Forms
            Return GetConfigOption("option_qp", False)
          Case "CDAMTM"                       'Table Maintenance
            Return GetConfigOption("option_maintenance", True)
          Case "CDAMCO"                       'Close Open Batch
            Return GetConfigOption("option_trader", False)
          Case "CDFMAF", "SCFMAF", "SCFLNA", "SCFLNB", "CDDPND", "SCBMAC", "SCCPAC" 'Action Finder - New action, Document Actions, Browser Menu- Actions, Campaign Actions       
            Return GetConfigOption("option_actions", False)
          Case "CDFMEF", "SCFMEF"             'Event Finder
            Return GetConfigOption("option_events", False)
          Case "CDFME2", "SCFME2"             'Event Finder
            Return GetConfigOption("option_events", False) And EntityGroups.EventGroupCount > 1
          Case "CDFME3", "SCFME3"             'Event Finder
            Return GetConfigOption("option_events", False) And EntityGroups.EventGroupCount > 2
          Case "CDFME4", "SCFME4"             'Event Finder
            Return GetConfigOption("option_events", False) And EntityGroups.EventGroupCount > 3
          Case "CDFME5", "SCFME5"             'Event Finder
            Return GetConfigOption("option_events", False) And EntityGroups.EventGroupCount > 4
          Case "CDAMUL"                       'Update Local Data
            Return GetConfigOption("option_cache_manual_update", False)
          Case "CDFMMF", "CDAMDM", "SCFMFF", "SCAMDM" 'Meeting Finder, Duplicate Meeting
            Return GetConfigOption("option_meetings", False)
          Case "CDFMLG", "SCFMLG"             'Legacy Finder
            Return GetConfigOption("option_legacies", False)
          Case "CDFLIO"                       'Intranet Output
            Return GetConfigOption("option_intranet", False)
          Case "CDTMDN"                       'Dial Number
            Return GetConfigOption("option_cti", False)
          Case "CDFMPP", "CDFMSO", "CDFMDD", "CDFMCC", "SCFMPP", "SCFMSO", "SCFMDD", "SCFMCC"  'Payment Plan Finder, SO, DD, CCCA Finders
            Return GetConfigOption("option_payment_plans", True)
          Case "CDFMCF", "SCFMCF"             'Covenant Finder
            Return GetConfigOption("option_covenants", True)
          Case "CDFMBF", "SCFMBF"             'Member Finder
            Return GetConfigOption("option_membership", True)
          Case "CDFMTF", "CDFMIF", "SCFMTF", "SCFMIF"     'Transaction Finder, Invoice/Credit Note Finder
            Return GetConfigOption("option_financial", True)
          Case "CDFMGY", "SMFYPL", "SMFYRC", "SCFMGY"  'Pre Tax Payroll Giving (Give As You Earn) - Finder and System menu access
            Return GetConfigOption("option_gaye", False)
          Case "CDFMPG", "SMFYPT", "SCFMPG"            'Post Tax Payroll Giving - Finder and System menu access
            Return GetConfigOption("option_gaye", False)
          Case "CDAMPP"                       'Postcode Proximity
            Return GetConfigOption("opt_cd_create_grid_references")
          Case "CDVMMT", "CDVMMC", "CDVMMI", "CDVMMD", "SCVMMT", "SCVMMC", "SCVMMI", "SCVMMD"  'My Details, Calendar, InBox, Documents
            Return (User.ContactNumber > 0)
          Case "CDVMMO", "SCVMMO"             'My Organisation
            Return (User.OrganisationNumber > 0)
          Case "CDVMMM"                       'My meetings
            Return GetConfigOption("option_meetings", False) And (User.ContactNumber > 0)
          Case "CDVMMA", "SCVMMA"             'My Actions
            Return GetConfigOption("option_actions", False) And (User.ContactNumber > 0)
          Case Else
            Return True
        End Select
      End If
    End Function

    Public ReadOnly Property OwnershipMethod() As OwnershipMethods
      Get
        If GetConfig("ownership_method") = "G" Then
          OwnershipMethod = OwnershipMethods.omOwnershipGroups
        Else
          OwnershipMethod = OwnershipMethods.omOwnershipDepartments
        End If
      End Get
    End Property

    Public ReadOnly Property Postcoder() As Postcoder
      Get
        If mvPostcoder Is Nothing Then
          mvPostcoder = New Postcoder
          mvPostcoder.Init(Me)
        End If
        Return mvPostcoder
      End Get
    End Property

    Public Sub ResetAuditStyle()
      'Setup Auditing
      Select Case GetConfig("option_audit", "N").Substring(0, 1).ToUpper 'BR13764: Change to default config value to 'N' to handle no config option set for client/department/logname
        Case "Y"
          mvAuditStyle = AuditStyleTypes.ausAuditMultipleRecords
        Case "O"
          mvAuditStyle = AuditStyleTypes.ausAuditOneRecord
        Case "A"
          mvAuditStyle = AuditStyleTypes.ausAmendmentHistory
        Case "X"
          mvAuditStyle = AuditStyleTypes.ausExtended
        Case Else
          mvAuditStyle = AuditStyleTypes.ausNone
      End Select
    End Sub

    Public Sub RollbackConnection()
      If mvConnection IsNot Nothing Then
        mvConnection.RollbackTransaction(False)
      End If
    End Sub

    Public Property User() As CDBUser
      Get
        Return mvUser
      End Get
      Set(ByVal pValue As CDBUser)
        mvUser = pValue
      End Set
    End Property

    Private mvSurnamePrefixes As List(Of String)

    Private Sub InitSurnamePrefixes()
      If mvSurnamePrefixes Is Nothing Then
        mvSurnamePrefixes = New List(Of String)
        If GetDataStructureInfo(cdbDataStructureConstants.cdbDataDutchSupport) Then
          Dim vRecordSet As CDBRecordSet = New SQLStatement(Connection, "surname_prefix", "surname_prefixes").GetRecordSet
          While vRecordSet.Fetch()
            mvSurnamePrefixes.Add(vRecordSet.Fields.Item(1).Value)
          End While
          vRecordSet.CloseRecordSet()
        End If
      End If
    End Sub

    Public Function ValidSurnamePrefix(ByVal pCode As String) As Boolean
      InitSurnamePrefixes()
      ValidSurnamePrefix = mvSurnamePrefixes.Contains(pCode)
    End Function

    Public Function HasSurnamePrefixes() As Boolean
      InitSurnamePrefixes()
      HasSurnamePrefixes = mvSurnamePrefixes.Count > 0
    End Function

    Public Shared Function GetOwnershipAccessLevel(ByVal pType As String) As OwnershipAccessLevelTypes
      Select Case pType
        Case "W"
          Return OwnershipAccessLevelTypes.oaltWrite
        Case "R"
          Return OwnershipAccessLevelTypes.oaltRead
        Case "B"
          Return OwnershipAccessLevelTypes.oaltBrowse
        Case Else
          Return OwnershipAccessLevelTypes.oaltNone
      End Select
    End Function

    Public Shared Function GetOwnershipAccessLevelCode(ByVal pTypeCode As OwnershipAccessLevelTypes) As String
      Select Case pTypeCode
        Case OwnershipAccessLevelTypes.oaltWrite
          Return "W"
        Case OwnershipAccessLevelTypes.oaltRead
          Return "R"
        Case OwnershipAccessLevelTypes.oaltBrowse
          Return "B"
        Case Else
          Return ""
      End Select
    End Function

    Public Sub SwitchCurrentUser(ByVal pLogname As String)
      'This routine is used by the Job Processor to switch the current user from the Job Processor user
      'to the submitter of the job - It does however need to retain the ability to send email
      'as though it was the Job Processor user
      Dim vEMailLogin As String = mvUser.EmailLogin
      Dim vLogname As String = mvUser.Logname
      Dim vPassword As String = mvUser.Password
      Dim vDBLogname As String = mvUser.DatabaseLogname

      mvUser = New CDBUser(Me, pLogname, vPassword, vDBLogname, pLogname)
      mvUser.InitWithLogname()
      If mvUser.Logname <> pLogname Then RaiseError(DataAccessErrors.daeNotValidAccountOrPassword)
      Dim vParams As New CDBParameters
      If Len(vEMailLogin) > 0 Then
        vParams.Add("EmailLogin", vEMailLogin)
        'mvUser.EmailLogin = vEMailLogin
      Else
        vParams.Add("EmailLogin", vLogname)
        'mvUser.EmailLogin = vLogname
      End If
      mvUser.Update(vParams)
      mvConfigs = Nothing
    End Sub

    Public Sub SetCurrentOutputDirectory(Optional ByVal pDirectoryType As SetCurrentOutputDirectoryTypes = SetCurrentOutputDirectoryTypes.scodtOutput, Optional ByRef pDirName As String = "")
      Dim vDirName As String = Nothing
      Select Case pDirectoryType
        Case SetCurrentOutputDirectoryTypes.scodtOutput
          vDirName = GetOutputDirectory(OutputDirectoryTypes.scodtOutput)
        Case SetCurrentOutputDirectoryTypes.scodtMailing
          vDirName = GetOutputDirectory(OutputDirectoryTypes.scodtMailing)
      End Select
      pDirName = vDirName
    End Sub

    Public Function GetLogFileName(Optional ByVal pFileName As String = "") As String
      Dim vDirName As String = GetOutputDirectory(OutputDirectoryTypes.scodtLogFiles)
      If pFileName.Length > 0 Then vDirName = Path.Combine(vDirName, pFileName)
      Return vDirName
    End Function

    Public Function GetAuditFileName(Optional ByVal pFileName As String = "") As String
      Dim vDirName As String = GetOutputDirectory(OutputDirectoryTypes.scodtAuditFiles)
      If pFileName.Length > 0 Then vDirName = Path.Combine(vDirName, pFileName)
      Return vDirName
    End Function

    Public Function GetMailingFileName(ByVal pMailing As Boolean, ByVal pNumber As Long) As String
      Dim vFileName As New StringBuilder
      Dim vDirName As String = GetOutputDirectory(OutputDirectoryTypes.scodtMailing)
      With vFileName
        If pMailing Then
          .Append("Mailing_")
          .Append(mvUser.Logname.Substring(0, 2).ToUpper)
          .Append(Format$(Date.Now, "ddmmyyyy"))
          .Append(Format$(Now, "hhmmss"))
          .Append(".csv")
          If GetConfig("ma_allocate_mailing_number") = "mailing" And GetConfigOption("ma_auto_name_mailing_files") Then RaiseError(DataAccessErrors.daeMailingNumberConfig)
        Else
          .Append("Fulfillment_")
          .Append(pNumber)
          .Append(".csv")
        End If
        If vDirName.Length > 0 And vDirName.EndsWith("\") = False Then vDirName = vDirName & "\"
        vDirName = vDirName & vFileName.ToString
      End With
      Return vDirName
    End Function

    Public Function GetProductNumber(ByVal pProduct As String) As Integer
      Return GetProductNumber(pProduct, True)
    End Function
    Public Function GetProductNumber(ByVal pProduct As String, ByVal pCheckProductNumbers As Boolean) As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vProductNumber As Integer
      Dim vGotProductNumber As Boolean
      Dim vRetries As Integer
      Dim vExitDo As Boolean

      vWhereFields.Add("product", pProduct)
      If pCheckProductNumbers Then
        'If we get a record loop until we can successfully delete it
        Do
          vRecordSet = New SQLStatement(Connection, "product_number", "product_numbers", vWhereFields, "product_number").GetRecordSet
          If vRecordSet.Fetch() Then
            vProductNumber = vRecordSet.Fields(1).IntegerValue
            vWhereFields.Add("product_number", vProductNumber)
            If Connection.DeleteRecords("product_numbers", vWhereFields, False) > 0 Then vGotProductNumber = True
          Else
            vExitDo = True
          End If
          vRecordSet.CloseRecordSet()
          vRetries += 1
        Loop While vGotProductNumber = False AndAlso vExitDo = False AndAlso vRetries < MAX_RETRIES
      End If

      Dim vUpdateFields As New CDBFields
      Dim vQuantity As Integer

      vRetries = 0
      'Loop until we successfully update the product number
      While vGotProductNumber = False AndAlso vRetries < MAX_RETRIES
        vRecordSet = New SQLStatement(Connection, "next_product_number,sales_quantity", "products", vWhereFields).GetRecordSet
        If vRecordSet.Fetch() Then
          vProductNumber = vRecordSet.Fields("next_product_number").IntegerValue
          vQuantity = vRecordSet.Fields("sales_quantity").IntegerValue
          If vQuantity < 1 Then vQuantity = 1
          vUpdateFields.Add("next_product_number", vProductNumber + vQuantity)  '1
          vUpdateFields.Add("amended_by", mvUser.Logname)
          vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
          vWhereFields.Add("next_product_number", vRecordSet.Fields("next_product_number").IntegerValue)
          'If someone else has nicked the product number the update will fail and it will loop
          If Connection.UpdateRecords("products", vUpdateFields, vWhereFields, False) > 0 Then vGotProductNumber = True
        Else
          RaiseError(DataAccessErrors.daeMissingProduct, pProduct)
        End If
        vRecordSet.CloseRecordSet()
        vRetries += 1
      End While
      Return vProductNumber
    End Function
    Public Function GetTableSpaceInfo(ByVal pTable As String, ByVal pIndex As Boolean) As String
      Dim vRecordSet As CDBRecordSet
      Dim vTableSpace As String
      Dim vReturnValue As String = ""

      If mvConnection.UseTableSpaces() Then
        vRecordSet = New SQLStatement(Connection, "SELECT tablespace_name FROM all_tables WHERE table_name = '" & pTable.ToUpper & "'").GetRecordSet
        If vRecordSet.Fetch() Then
          vTableSpace = vRecordSet.Fields(1).Value
          Select Case vTableSpace
            Case "LARGE_TABLES", "MEDIUM_TABLES", "SMALL_TABLES"
              If pIndex Then
                vReturnValue = " TABLESPACE " & vTableSpace.Replace("TABLES", "INDEXES")
              Else
                vReturnValue = " TABLESPACE " & vTableSpace
              End If
          End Select
        Else
          Select Case pTable.ToLower()
            Case "addresses", "batch_transactions", "batch_transaction_analysis",
                 "contacts", "contact_addresses", "contact_categories", "contact_expenditure",
                 "contact_header", "contact_mailings", "contact_performances",
                 "financial_history", "financial_history_details", "orders", "order_payment_history"
              If pIndex Then
                vReturnValue = " TABLESPACE LARGE_INDEXES"
              Else
                vReturnValue = " TABLESPACE LARGE_TABLES"
              End If
            Case "amendment_history", "bankers_orders", "bank_transactions", "batches",
                 "communications", "communications_log", "communications_log_history",
                 "communications_log_links", "communications_log_subjects", "contact_accounts",
                 "contact_journals", "contact_links", "contact_positions", "contact_suppressions",
                 "contact_users", "covenants", "declaration_lines_unclaimed", "declaration_tax_claim_lines",
                 "gift_aid_declarations", "order_details", "order_external_links", "organisations",
                 "organisation_addresses", "organisation_categories", "tax_claim_lines"
              If pIndex Then
                vReturnValue = " TABLESPACE MEDIUM_INDEXES"
              Else
                vReturnValue = " TABLESPACE MEDIUM_TABLES"
              End If
            Case Else
              If pIndex Then
                vReturnValue = " TABLESPACE SMALL_INDEXES"
              Else
                vReturnValue = " TABLESPACE SMALL_TABLES"
              End If
          End Select
        End If
        vRecordSet.CloseRecordSet()
      End If
      Return vReturnValue
    End Function

    ''' <summary>Calculates and returns the auto payment date.</summary>
    ''' <param name="pDate">The base date used for the calculations.</param>
    ''' <param name="pAutoPayMethod">The type of auto payment method.</param>
    ''' <param name="pBankAccount">The BankAccount for the auto payment method.</param>
    ''' <returns>The calculated date based upon the pAutoPayMethod and the number of working days.</returns>
    Public Function GetPaymentPlanAutoPayDate(ByVal pDate As Date, ByVal pAutoPayMethod As PaymentPlan.ppAutoPayMethods, ByVal pBankAccount As BankAccount) As Date
      Return GetPaymentPlanAutoPayDate(pDate, pAutoPayMethod, pBankAccount, False)
    End Function

    ''' <summary>Calculates and returns the auto payment date.</summary>
    ''' <param name="pDate">The base date used for the calculations.</param>
    ''' <param name="pAutoPayMethod">The type of auto payment method.</param>
    ''' <param name="pBankAccount">The BankAccount for the auto payment method.</param>
    ''' <param name="pGetAutoPayAdvancePeriod">Set to True to use the Membership Controls 'Auto Pay Advance Notice Period', otherwise False to use the current configuration for the auto pay delay days.</param>
    ''' <returns>The calculated date based upon the pAutoPayMethod and the number of working days.</returns>
    Public Function GetPaymentPlanAutoPayDate(ByVal pDate As Date, ByVal pAutoPayMethod As PaymentPlan.ppAutoPayMethods, ByVal pBankAccount As BankAccount, ByVal pGetAutoPayAdvancePeriod As Boolean) As Date
      Dim vDays As Integer = 0

      If pBankAccount Is Nothing Then
        pBankAccount = New BankAccount(Me)
        pBankAccount.Init()
      End If

      If pBankAccount.Existing = False Then
        Debug.WriteLine("Bank Account does not exist")
      End If

      If pGetAutoPayAdvancePeriod Then
        'Get Membership Controls Auto Pay Advance Period value
        vDays = IntegerValue(GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAutoPayAdvancePeriod))
      Else
        'Get Auto Pay Delay from Bank Account / Configuration Option
        vDays = pBankAccount.GetAutoPayDelayDays(pAutoPayMethod)
      End If
      If vDays = 0 Then Return pDate

      Dim vBankHolidays As New CollectionList(Of Date)
      Dim vCheckBankHolidays As Boolean = False
      If pAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMDD AndAlso IsDefaultCountryUK() Then
        Dim vWherefields As New CDBFields
        vWherefields.Add("bank_holiday_day", pDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
        vWherefields.Add("bank_holiday_day#2", pDate.AddDays((vDays * 5) + 4), CDBField.FieldWhereOperators.fwoBetweenTo)
        Dim vSQLStatement As New SQLStatement(Me.Connection, "bank_holiday_day,bank_holiday_day_desc", "bank_holiday_days", vWherefields, "bank_holiday_day")
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          vBankHolidays.Add(CDate(vRS.Fields(1).Value).ToString(CAREDateFormat), CDate(vRS.Fields(1).Value))
        End While
        vRS.CloseRecordSet()
        vCheckBankHolidays = True
      End If
      Dim vNewDate As Date = AddWeekdays(pDate, vDays, vCheckBankHolidays, vBankHolidays)

      Return vNewDate
    End Function

    Friend Function GetCountryIban(ByVal pIbanNumber As String) As CountryIbanNumber
      If mvCountryIbanNumbers Is Nothing Then mvCountryIbanNumbers = New CollectionList(Of CountryIbanNumber)
      Dim vCountryCode As String = Substring(pIbanNumber.Trim, 0, 2).ToUpper.Trim   'Just extract the country code (first 2 characters)
      If vCountryCode.Length > 0 AndAlso mvCountryIbanNumbers.ContainsKey(vCountryCode) Then
        Return mvCountryIbanNumbers(vCountryCode)
      Else
        Dim vCountryIbanNumber As New CountryIbanNumber(Me)
        vCountryIbanNumber.Init(vCountryCode)
        If vCountryCode.Length > 0 Then mvCountryIbanNumbers.Add(vCountryCode, vCountryIbanNumber)
        Return vCountryIbanNumber
      End If
    End Function

    Public Function Country(ByRef pCode As String) As Country
      If mvCountries Is Nothing Then mvCountries = New CollectionList(Of Country)
      If mvCountries.ContainsKey(pCode) Then
        Return mvCountries(pCode)
      Else
        Dim vCountry As Country = New Country(Me)
        vCountry.Init(pCode)
        mvCountries.Add(pCode, vCountry)
        Return vCountry
      End If
    End Function

#Region "VAT Rates"

    Private mvVATRates As CollectionList(Of VatRate)
    Public Function VATRate(ByVal pVATRateCode As String) As VatRate
      Return VATRate(pVATRateCode, "")
    End Function
    Public Function VATRate(ByVal pProductCategory As String, ByVal pContactCategory As String) As VatRate
      If mvVATRates Is Nothing Then InitVATRates()
      Dim vVATRate As New VatRate(Me)
      If mvVATRates.ContainsKey(pProductCategory & pContactCategory) Then
        vVATRate = mvVATRates.Item(pProductCategory & pContactCategory)
      Else
        RaiseError(DataAccessErrors.daeVATRateInvalid, pProductCategory & pContactCategory)
      End If
      Return vVATRate
    End Function
    Private Sub InitVATRates()
      mvVATRates = New CollectionList(Of VatRate)
      Dim vVATRate As New VatRate(Me)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("vat_rate_identification vri", "vr.vat_rate", "vri.vat_rate")
      Dim vAttrs As String = "product_vat_category, contact_vat_category, " & vVATRate.GetRecordSetFields()
      Dim vOrderBy As String = "product_vat_category, contact_vat_category"
      If GetDataStructureInfo(cdbDataStructureConstants.cdbDataVatRateHistory) Then
        vAttrs &= ", vrh.percentage AS vrh_percentage, vrh.rate_changed AS vrh_rate_changed"
        vOrderBy &= ", vr.vat_rate, vrh.rate_changed DESC"
        vAnsiJoins.AddLeftOuterJoin("vat_rate_history vrh", "vr.vat_rate", "vrh.vat_rate")
      End If
      Dim vSQL As New SQLStatement(Me.Connection, vAttrs, "vat_rates vr", Nothing, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      If vRS.Fetch Then
        Dim vCCategory As String
        Dim vPCategory As String
        While vRS.Status = True
          vCCategory = vRS.Fields("contact_vat_category").Value
          vPCategory = vRS.Fields("product_vat_category").Value
          vVATRate = New VatRate(Me)
          vVATRate.InitFromRecordSetWithHistory(vRS)
          mvVATRates.Add(vPCategory & vCCategory, vVATRate)
          'This next will add the VAT Rate with a key of just the VAT Rate Code
          'The check is because the select may give multiple records with the same VAT Rate Code
          If Not mvVATRates.ContainsKey(vVATRate.VatRateCode) Then mvVATRates.Add(vVATRate.VatRateCode, vVATRate)
        End While
      End If
      vRS.CloseRecordSet()
    End Sub

    Public Function GetFixedCycleConfig(ByVal pPlanType As ppType) As String
      Dim vConfig As String

      Select Case pPlanType
        Case ppType.pptMember
          vConfig = "fixed_cycle_M"
        Case ppType.pptCovenant
          vConfig = "fixed_cycle_C"
        Case ppType.pptSO
          vConfig = "fixed_cycle_B"
        Case ppType.pptDD
          vConfig = "fixed_cycle_D"
        Case ppType.pptCCCA
          vConfig = "fixed_cycle_A"
        Case Else
          vConfig = "fixed_cycle_O"
      End Select
      Return GetConfig(vConfig)
    End Function

    Public Function GetStartDate(ByVal pPlanType As ppType, Optional ByVal pUseStartmonth As Boolean = False) As String
      Dim vFixedCycle As String
      Dim vSubsequentMonth As Boolean
      Dim vStartDate As String = ""
      Dim vThisMonth As Integer
      Dim vThisYear As Integer
      Dim vStartDay As Integer
      Dim vStartMonth As Integer
      Dim vLastDateThisMonth As Date
      Dim vLastDayThisMonth As Integer
      Dim vFixedCycleDay As Integer
      Dim vPrevious As Boolean

      If pUseStartmonth Then
        vFixedCycle = GetConfig("fixed_pay_plan_start_date")
        vPrevious = True
      Else
        vFixedCycle = GetFixedCycleConfig(pPlanType)
      End If
      vSubsequentMonth = GetConfigOption("me_next_month_start_date")
      vThisMonth = Today.Month
      vThisYear = Today.Year
      'The use of a P at the end of the config value indicates that the s/w should use the last start period, rather than the next start period.
      'It should work like this:
      '1. Assume the config is set to '0110P' and that the current date is 04/07/2003.
      '2. Since P is the fifth character vPrevious will be set to True and the config value will be set to '0110'
      '3. If the current date < the date constructed from the config value & the current year then subtract one from vThisYear.
      '3a.Is 04/07/2003 < 01/10/2003?  Yes, so vThisYear will be reduced from 2003 to 2002.
      '4. So, the constructed start date will then be the config value & '2002', i.e. 01/10/2002.
      If vFixedCycle.Length = 5 Then 'Format = DDMM + P
        vPrevious = UCase(Right(vFixedCycle, 1)) = "P"
        vFixedCycle = Left(vFixedCycle, Len(vFixedCycle) - 1)
      End If

      If IsNumeric(vFixedCycle) Then
        Select Case Len(vFixedCycle)
          Case 2
            vFixedCycleDay = IntegerValue(vFixedCycle)
            If vFixedCycleDay >= 1 And vFixedCycleDay <= 31 Then
              If vThisMonth = 12 Then
                vLastDateThisMonth = DateSerial(vThisYear + 1, 1, 1)
              Else
                vLastDateThisMonth = DateSerial(vThisYear, vThisMonth + 1, 1)
              End If
              vLastDateThisMonth = vLastDateThisMonth.AddDays(-1)
              vLastDayThisMonth = vLastDateThisMonth.Day
              If vFixedCycleDay > vLastDayThisMonth Then vFixedCycleDay = vLastDayThisMonth
              vStartDate = DateSerial(vThisYear, vThisMonth, vFixedCycleDay).ToString(CAREDateFormat)
              If vSubsequentMonth And pPlanType = ppType.pptMember Then vStartDate = CDate(vStartDate).AddMonths(1).ToString(CAREDateFormat)
            End If
          Case 4 ' Day & Month of current year
            vStartDay = IntegerValue(Mid(vFixedCycle, 1, 2))
            vStartMonth = IntegerValue(Mid(vFixedCycle, 3, 2))
            If vPrevious Then
              If DateDiff(Microsoft.VisualBasic.DateInterval.Day, DateSerial(vThisYear, vThisMonth, Today.Day), DateSerial(vThisYear, vStartMonth, vStartDay)) > 0 Then vThisYear -= 1
            End If
            If vStartDay >= 1 And vStartDay <= 31 And vStartMonth >= 1 And vStartMonth <= 12 Then
              vStartDate = DateSerial(vThisYear, vStartMonth, vStartDay).ToString(CAREDateFormat)
            End If
        End Select
      End If
      Return vStartDate
    End Function

#End Region

    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          If mvConnection IsNot Nothing Then mvConnection.CloseConnection()
          ' TODO: free managed resources when explicitly called
        End If

        ' TODO: free shared unmanaged resources
      End If
      Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

    Function GetConfigScopeLevel(pConfigName As String) As Config.ConfigNameScope
      'Scope always defaults to System + Department + User because that's what it was before we introduced the concept of scope.
      Dim vScope As Config.ConfigNameScope = ConfigNameScope.SystemAndDepartmentAndUser
      If Not mvCachedConfigScope.ContainsKey(pConfigName) Then
        Dim vStmt As New SQLStatement(Me.Connection, "config_scope", "config_names", New CDBField("config_name", pConfigName))
        Dim vValue As String = (0 + vScope).ToString 'simple way to convert enum to int then to string
        If Me.Connection.AttributeExists("config_names", "config_scope") Then
          vValue = Me.Connection.GetValue(vStmt.SQL)
        End If
        [Enum].TryParse(Of ConfigNameScope)(vValue, vScope)
        mvCachedConfigScope.Add(pConfigName, vScope)
      Else
        vScope = mvCachedConfigScope(pConfigName)
      End If
      Return vScope
    End Function

  End Class
End Namespace
