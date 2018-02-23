Namespace Access

  Public Class DataFinder

    Enum DataFinderTypes
      dftBatch
      dftCCCA
      dftCovenant
      dftDirectDebit
      dftEvent
      dftEventBooking
      dftEventPersonnel
      dftGAD
      dftLegacy
      dftMeeting
      dftMember
      dftOpenBatch
      dftPaymentPlan
      dftProduct
      dftStandingOrder
      dftTransaction
      dftInvoice
      dftManualSOReconciliation
      dftVenue
      dftGiveAsYouEarn 'Pre Tax Payroll Giving
      dftPurchaseOrder
      dftInternalResource
      dftServiceProduct
      dftCommunication
      dftFindContacts
      dftFindOrganisations
      dftUniservPhoneBook
      dftFindDocuments 'Only used by the Smart client to get the controls for the finder
      dftFindDuplicateContacts 'Only used by the Smart client to get the controls for the finder
      dftFindDuplicateOrganisations 'Only used by the Smart client to get the controls for the finder
      dftFindEMailContacts 'Only used by the Smart client to get the controls for the finder
      dftFindEMailOrganisations 'Only used by the Smart client to get the controls for the finder
      dftFindMailings 'Only used by the Smart client to get the controls for the finder
      dftTextSearch 'Only used by the Smart client to get the controls for the finder
      dftContactMailingDocuments
      dftActions
      dftSelectionSets 'Only used by the Smart client to get the controls for the finder
      dftCampaign
      dftPostTaxPayrollGiving
      dftAppealCollections
      dftStandardDocuments
      dftMaxFinderType 'Marker for maximum finder type - subtract 1 for last one
      dftFindFundraisingRequest
      dftExternalReference
    End Enum

    Public Enum TransactionFinderTypes
      tftProcessed
      tftUnprocessed
      tftProvisional
      tftCancelledProvisional
    End Enum

    Public Enum SOFinderType
      softNormal
      softCAF
      softBoth
    End Enum

    Private mvAddActionRelatedContactData As Boolean
    Private mvSOFinderType As SOFinderType
    Private mvTransactionFinderType As TransactionFinderTypes
    Private mvSelectItems As CDBFields
    Private mvWhereFlag As Boolean
    Private mvEnv As CDBEnvironment
    Private mvConn As CDBConnection
    Private mvControls As CDBControls
    Private mvType As DataFinderTypes
    Private mvCustomColumns As String
    Private mvCustomHeadings As String
    Private mvResultColumns As String 'The full list of columns which could be displayed
    Private mvSelectColumns As String 'The default list of columns normally displayed
    Private mvSelectWidths As String 'The widths for the default list of columns (WEB only)
    Private mvSelectHeadings As String
    Private mvSelectAttrs As String

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pType As DataFinderTypes)
      mvEnv = pEnv
      mvConn = mvEnv.Connection
      mvSelectItems = New CDBFields
      mvControls = Nothing
      mvType = pType
      mvCustomColumns = ""
      mvCustomHeadings = ""
      mvSelectColumns = ""

      Select Case mvType
        Case DataFinderTypes.dftActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,ActionerSetting,ManagerSetting"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,Deadline,ScheduledOn,CompletedOn,CreatedBy,CreatedOn,ActionPriority,ActionStatus,Access"
          mvSelectWidths = "800,2000,1500,1500,1600,1600,1600,1200,1200,1,1,1"
          'DocumentClass
          mvSelectHeadings = (DataSelectionText.String18145) 'Number,Description,Priority,Status,Deadline,Scheduled On,Completed On,Created By,Created On,Priority,Status
        Case DataFinderTypes.dftUniservPhoneBook
          mvResultColumns = "ContactNumber,Surname,Forenames,Postcode,Town,Address1,PHONEBOOK_PHONE"
          mvSelectColumns = "ContactNumber,Surname,Forenames,Postcode,Town,Address1,PHONEBOOK_PHONE"
          mvSelectWidths = "1,2000,2000,1200,2000,2000,1200"
          mvSelectHeadings = "Number,Surname,Forenames,Postcode,Town,Address,Phone Number"
        Case DataFinderTypes.dftStandardDocuments
          mvResultColumns = "StandardDocument,StandardDocumentDesc,DocumentType,DocumentTypeDesc,Topic,TopicDesc,SubTopic,SubTopicDesc,InstantPrint,Historical,PackageCode,MailmergeHeader"
          mvSelectColumns = "StandardDocument,StandardDocumentDesc,DocumentTypeDesc,InstantPrint,Historical,MailmergeHeader,TopicDesc,SubTopicDesc,PackageCode"
          mvSelectWidths = "800,2000,2000,800,800,900,2000,2000,1200"
          mvSelectHeadings = "Code,Description,Type,Instant Print,Historical,Merge Type,Topic,SubTopic,Package"
        Case DataFinderTypes.dftMember
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,MemberNumber,MembershipType,Branch,Joined,CancellationReason,CancelledOn,NumberOfMembers,LastPaymentDate,Balance,RenewalDate,MembershipNumber,PaymentPlanNumber,Postcode,Forenames"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,MemberNumber,MembershipType,Branch,Joined,CancellationReason,CancelledOn,NumberOfMembers,LastPaymentDate,Balance,RenewalDate,MembershipNumber,PaymentPlanNumber,Postcode,Forenames"
        Case DataFinderTypes.dftCCCA
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,CreditCardAuthorityNumber,CreditCardNumber,StartDate,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,PaymentPlanNumber"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,CreditCardAuthorityNumber,CreditCardNumber,StartDate,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,PaymentPlanNumber"
        Case DataFinderTypes.dftCovenant
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,CovenantNumber,CovenantTerm,CovenantedAmount,StartDate,DepositedDeed,Net,CancellationReason,PaymentPlanNumber"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,CovenantNumber,CovenantTerm,CovenantedAmount,StartDate,DepositedDeed,Net,CancellationReason,PaymentPlanNumber"
        Case DataFinderTypes.dftDirectDebit
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,DirectDebitNumber,Reference,StartDate,SortCode,AccountNumber,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,PaymentPlanNumber,IbanNumber,BicCode"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,DirectDebitNumber,Reference,StartDate,SortCode,AccountNumber,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,PaymentPlanNumber,IbanNumber,BicCode"
        Case DataFinderTypes.dftEventPersonnel
          mvResultColumns = "ContactNumber,AddressNumber,Surname,Initials"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials"
          mvSelectWidths = "1,1,2000,2000"
        Case DataFinderTypes.dftGAD
          mvResultColumns = "ContactNumber,OwnershipGroup,Surname,Initials,DeclarationNumber,DeclarationDate,DeclarationType,Method,ConfirmedOn,StartDate,EndDate"
          mvSelectColumns = "ContactNumber,Surname,Initials,DeclarationNumber,DeclarationDate,DeclarationType,Method,ConfirmedOn,StartDate,EndDate"
        Case DataFinderTypes.dftInvoice
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,InvoiceNumber,Company,SalesLedgerAccount,PaymentDue"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,InvoiceNumber,Company,SalesLedgerAccount,PaymentDue"
        Case DataFinderTypes.dftLegacy
          mvResultColumns = "ContactNumber,OwnershipGroup,Surname,Initials,LegacyNumber,LegacyID,LegacyStatusDesc,DateOfDeath,ReviewDate,LabelName"
          mvSelectColumns = "ContactNumber,Surname,Initials,LegacyNumber,LegacyID,LegacyStatusDesc,DateOfDeath,ReviewDate,LabelName"
        Case DataFinderTypes.dftPaymentPlan
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,OrderNumber,OrderType,Provisional,FrequencyAmount,NextPaymentDue,LastPaymentDate,Balance,OrderDate,RenewalDate,CancellationReason,PaymentMethod,ReasonForDespatch"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,OrderNumber,OrderType,Provisional,FrequencyAmount,NextPaymentDue,LastPaymentDate,Balance,OrderDate,RenewalDate,CancellationReason,PaymentMethod,ReasonForDespatch"
        Case DataFinderTypes.dftStandingOrder
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,BankersOrderNumber,StandingOrderType,StartDate,SortCode,AccountNumber,StandingOrderAmount,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,Reference,PaymentPlanNumber,IbanNumber,BicCode"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,BankersOrderNumber,StandingOrderType,StartDate,SortCode,AccountNumber,StandingOrderAmount,FrequencyAmount,BankAccount,CancellationReason,NextPaymentDue,Reference,PaymentPlanNumber,IbanNumber,BicCode"
        Case DataFinderTypes.dftTransaction
          mvResultColumns = "BatchNumber,TransactionNumber,Status,ContactNumber,OwnershipGroup,LabelName,TransactionDate,Amount,Reference,Posted,Notes,ContactType"
          mvSelectColumns = "BatchNumber,TransactionNumber,Status,ContactNumber,LabelName,TransactionDate,Amount,Reference,Posted,Notes,ContactType"
        Case DataFinderTypes.dftGiveAsYouEarn
          'Pre Tax Payroll Giving
          mvResultColumns = "ContactNumber,OwnershipGroup,GayePledgeNumber,DonorID,Surname,Initials,Name,OrganisationNumber,AgencyName,StartDate,CancelledOn,Source,PostBatchesToCashBook"
          mvSelectColumns = "ContactNumber,GayePledgeNumber,DonorID,Surname,Initials,Name,OrganisationNumber,AgencyName,StartDate,CancelledOn,Source,PostBatchesToCashBook"
        Case DataFinderTypes.dftPurchaseOrder
          mvResultColumns = "ContactNumber,AddressNumber,OwnershipGroup,Surname,Initials,PurchaseOrderNumber,PurchaseOrderType,PurchaseOrderDesc,StartDate,NumberOfPayments,CancellationReason,CancelledOn"
          mvSelectColumns = "ContactNumber,AddressNumber,Surname,Initials,PurchaseOrderNumber,PurchaseOrderType,PurchaseOrderDesc,StartDate,NumberOfPayments,CancellationReason,CancelledOn"
        Case DataFinderTypes.dftPostTaxPayrollGiving
          mvResultColumns = "ContactNumber,OwnershipGroup,PledgeNumber,Surname,Initials,Name,OrganisationNumber,PayrollNumber,StartDate,CancelledOn,Source"
          mvSelectColumns = "ContactNumber,PledgeNumber,Surname,Initials,Name,OrganisationNumber,PayrollNumber,StartDate,CancelledOn,Source"
        Case DataFinderTypes.dftServiceProduct
          mvResultColumns = "ContactNumber,AddressNumber,ServiceName,Town,Postode,GeographicalRegion"
          mvSelectColumns = mvResultColumns
          mvSelectHeadings = "Contact Number,Address Number,Service Name,Town,Postode,Region"
          mvSelectWidths = "1,1,3000,1700,1300,2000"
        Case DataFinderTypes.dftAppealCollections
          mvResultColumns = "CollectionNumber,Campaign,Appeal,Collection,CollectionDesc"
          mvSelectColumns = "CollectionNumber,Campaign,Appeal,Collection,CollectionDesc"
        Case DataFinderTypes.dftInternalResource
          mvResultColumns = "ResourceNumber,ResourceContactNumber,Name,Product,ProductDesc,Rate,RateDesc"
          mvSelectColumns = "ResourceNumber,ResourceContactNumber,Name,Product,ProductDesc,Rate,RateDesc"
          mvSelectAttrs = "resource_number,resource_contact_number,CONTACT_NAME,product,product_desc,rate,rate_desc"
          mvSelectWidths = "1200,1200,2000,1000,2000,1000,2000"
        Case DataFinderTypes.dftProduct
          mvResultColumns = "Product,ProductDesc,ExtraKey,SalesGroup,Donation,Subscription,StockItem,MinimumQuantity,MaximumQuantity,ProductVATCategory,DespatchMethod,SalesDescription,HistoryOnly,Warehouse,LastStockCount,UsesProductNumbers,SalesQuantity,AmendedBy,AmendedOn"
          mvSelectColumns = "Product,ProductDesc,ExtraKey,SalesGroup,Donation,Subscription,StockItem"
          mvSelectWidths = "2130,5265,1980,1125,800,800,800,1,1,1,1,1,1,1"
        Case DataFinderTypes.dftEvent
          mvResultColumns = "EventNumber,StartDate,EventDesc,EventReference,NumberOfAttendees,MaximumAttendees,NumberOnWaitingList,NumberInterested,SubjectDesc,SkillLevelDesc,EventStatus,VenueDesc,Template,MultiSession,EventClass,Booking,BookingsClose,Branch"
          mvSelectColumns = mvResultColumns
          mvSelectWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
        Case DataFinderTypes.dftFindFundraisingRequest
          mvResultColumns = "RequestNumber,ContactNumber,RequestDate,RequestDesc,RequestStage,Status,RequestType,TargetAmount,PledgedAmount,PledgedDate,ReceivedAmount,ReceivedDate,ExpectedAmount,GikExpectedAmount,GikPledgedAmount,GikPledgedDate,TotalGikReceivedAmount,LatestGikReceivedDate,NumberOfPayments,RequestEndDate,FundraisingBusinessType"
          mvSelectColumns = mvResultColumns
          mvSelectWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
      End Select
    End Sub

    Public ReadOnly Property AvailableColumns() As String
      Get
        Return mvResultColumns
      End Get
    End Property

    Public ReadOnly Property DisplayColumns() As String
      Get
        If mvType = DataFinderTypes.dftTransaction Then SetTransactionFinderType()
        Return mvSelectColumns
      End Get
    End Property

    Public ReadOnly Property DisplayWidths() As String
      Get
        Return mvSelectWidths
      End Get
    End Property

    Public ReadOnly Property SelectedAttributes() As String
      Get
        'This is not the best approach to be followed and definitely needs to be changed to use the mvSelectColumns instead of the selected attributes here.
        'For now this has been fixed by removing ownership group from the list
        'when adding any new columns: it would be a good idea to try and follow the above approach and change this to use the mvSelectColumns
        Return Replace(Replace(GetSelectedAttributes(), ",c.ownership_group", ""), ",ownership_group", "")
      End Get
    End Property

    Public ReadOnly Property CustomColumns() As String
      Get
        Return mvCustomColumns
      End Get
    End Property

    Public ReadOnly Property CustomHeadings() As String
      Get
        Return mvCustomHeadings
      End Get
    End Property

    Private Function GetSelectedAttributes() As String
      Dim vAttrs As String = ""
      Select Case mvType
        Case DataFinderTypes.dftBatch, DataFinderTypes.dftOpenBatch
          vAttrs = "b.batch_number,provisional,"
          vAttrs = vAttrs & "batch_date,b.batch_type,b.bank_account,payment_method,source,number_of_entries,batch_total,number_of_transactions,transaction_total"
          vAttrs = vAttrs & ",detail_completed,ready_for_banking,paying_in_slip_printed,picked,posted_to_cash_book,posted_to_nominal,paying_in_slip_number"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs = vAttrs & ",currency_code"
          vAttrs = vAttrs & ",batch_created_by,batch_created_on"
          If mvType = DataFinderTypes.dftBatch Then vAttrs = vAttrs & ",u.department,d.department_desc"
        Case DataFinderTypes.dftCCCA
          vAttrs = "t.contact_number,t.address_number,ownership_group,surname,initials,credit_card_authority_number,credit_card_number,start_date,frequency_amount,bank_account,t.cancellation_reason,next_payment_due,t.order_number"
        Case DataFinderTypes.dftCovenant
          vAttrs = "t.contact_number,t.address_number,ownership_group,surname,initials,covenant_number,covenant_term,covenanted_amount,start_date,deposited_deed,net,t.cancellation_reason,t.order_number"
        Case DataFinderTypes.dftDirectDebit
          vAttrs = "t.contact_number,t.address_number,ownership_group,surname,initials,t.direct_debit_number,reference,start_date,sort_code,account_number,frequency_amount,bank_account,t.cancellation_reason,next_payment_due,t.order_number"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then
            vAttrs = vAttrs & ",iban_number,bic_code"
          Else
            vAttrs = vAttrs & ",NULL AS iban_number,null as bic_code"
          End If
        Case DataFinderTypes.dftEvent
          vAttrs = "ev.event_number,ev.start_date,event_desc,event_reference"
          vAttrs = vAttrs & ",number_of_attendees,maximum_attendees,number_on_waiting_list,number_interested,subject_desc,skill_level_desc,event_status,venue_desc,template,multi_session,ev.event_class,booking,bookings_close,ev.branch"
        Case DataFinderTypes.dftEventBooking
          vAttrs = "eb.booking_number,booking_date,event_desc,option_desc,eb.quantity,label_name,e.event_number,eb.option_number,eb.sales_contact_number,eb.notes"
        Case DataFinderTypes.dftEventPersonnel
          vAttrs = "p.contact_number,address_number,surname,initials"
        Case DataFinderTypes.dftServiceProduct
          vAttrs = "c.contact_number,a.address_number,surname,town,a.postcode,geographical_region"
        Case DataFinderTypes.dftGAD
          vAttrs = "t.contact_number,ownership_group,surname,initials,declaration_number,declaration_date,declaration_type,method,confirmed_on,start_date,end_date"
        Case DataFinderTypes.dftInvoice
          vAttrs = "t.contact_number,t.address_number,ownership_group,surname,initials,invoice_number,company,sales_ledger_account,payment_due"
        Case DataFinderTypes.dftMeeting
          vAttrs = "m.meeting_number,MEETING_DATE,MEETING_TIME,meeting_desc,meeting_type_desc,meeting_location_desc"
        Case DataFinderTypes.dftMember
          vAttrs = "t.contact_number,t.address_number,c.ownership_group,surname,initials,member_number,t.membership_type,t.branch,joined,t.cancellation_reason,t.cancelled_on,number_of_members,last_payment_date,balance,renewal_date,membership_number,t.order_number,a.postcode,c.forenames"
        Case DataFinderTypes.dftLegacy
          vAttrs = "t.contact_number,ownership_group,surname,initials,legacy_number,legacy_id,legacy_status_desc,date_of_death,review_date,label_name"
        Case DataFinderTypes.dftPaymentPlan
          vAttrs = "t.contact_number,t.address_number,ownership_group,surname,initials,order_number,order_type,provisional,"
          vAttrs = vAttrs & "frequency_amount,next_payment_due,last_payment_date,balance,order_date,renewal_date,t.cancellation_reason,payment_method,reason_for_despatch"
        Case DataFinderTypes.dftProduct
          vAttrs = "p.product,product_desc,extra_key,sales_group,donation,subscription,stock_item,minimum_quantity,maximum_quantity,product_vat_category,despatch_method,sales_description,history_only,p.warehouse,p.last_stock_count,uses_product_numbers,sales_quantity,p.amended_by,p.amended_on"
        Case DataFinderTypes.dftStandingOrder, DataFinderTypes.dftManualSOReconciliation
          vAttrs = "t.contact_number,t.address_number"
          If mvType = DataFinderTypes.dftStandingOrder Then vAttrs = vAttrs & ",ownership_group"
          vAttrs = vAttrs & ",surname,initials,bankers_order_number,standing_order_type,"
          vAttrs = vAttrs & "start_date,sort_code,account_number,t.amount,frequency_amount,bank_account,t.cancellation_reason,next_payment_due,reference,t.order_number"

          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then
            vAttrs = vAttrs & ",iban_number,bic_code"
          Else
            vAttrs = vAttrs & ",NULL AS iban_number,NULL as bic_code"
          End If
        Case DataFinderTypes.dftTransaction
          If mvTransactionFinderType = TransactionFinderTypes.tftProcessed Then
            vAttrs = "fh.batch_number,fh.transaction_number,fh.status,fh.contact_number,ownership_group,label_name,transaction_date,fh.amount,reference,fh.posted,fh.notes,contact_type"
          Else
            vAttrs = "bt.batch_number,bt.transaction_number,b.provisional,bt.contact_number,ownership_group,label_name,transaction_date,bt.amount,reference,bt.amended_on,bt.notes,contact_type"
          End If

        Case DataFinderTypes.dftVenue
          vAttrs = "venue,venue_desc"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEventVenueCapacity) Then
            vAttrs = vAttrs & ",venue_capacity"
          Else
            vAttrs = vAttrs & ",NULL AS venue_capacity"
          End If
          vAttrs = vAttrs & ",town,postcode"
        Case DataFinderTypes.dftGiveAsYouEarn
          vAttrs = "c.contact_number,c.ownership_group,t.gaye_pledge_number,donor_id,surname,initials,o.name,o.organisation_number,agency_name,start_date,cancelled_on,t.source,post_batches_to_cb"
        Case DataFinderTypes.dftPurchaseOrder
          vAttrs = "t.contact_number,t.address_number,c.ownership_group,surname,initials,purchase_order_number,purchase_order_type,purchase_order_desc,start_date,number_of_payments,cancellation_reason,cancelled_on"
        Case DataFinderTypes.dftInternalResource
          Dim vContact As New Contact(mvEnv)
          vAttrs = "resource_number,resource_contact_number,t.product,product_desc,t.rate,rate_desc," & vContact.GetRecordSetFieldsName
        Case DataFinderTypes.dftContactMailingDocuments
          vAttrs = "mailing_document_number,cmd.mailing_template,mailing_template_desc,label_name,created_by,created_on,cmd.mailing,mailing_desc,earliest_fulfilment_date"
          If mvSelectItems.Exists("none3") Then vAttrs = vAttrs & ",cmd.fulfillment_number,fulfilled_by,fulfilled_on"
        Case DataFinderTypes.dftCommunication
          vAttrs = "communication_number,com.contact_number,com.address_number,surname,initials,device_desc,com.dialling_code,com.std_code,number,extension,com.ex_directory"
        Case DataFinderTypes.dftPostTaxPayrollGiving
          vAttrs = "c.contact_number,c.ownership_group,t.pledge_number,surname,initials,o.name,o.organisation_number,payroll_number,start_date,cancelled_on,t.source"
        Case DataFinderTypes.dftAppealCollections
          vAttrs = "t.contact_number,t.collection_number,campaign,appeal,collection,collection_desc"
        Case DataFinderTypes.dftStandardDocuments
          vAttrs = "standard_document,standard_document_desc,sd.document_type,document_type_desc,sd.topic,topic_desc,sd.sub_topic,sub_topic_desc,instant_print,sd.history_only,sd.package,mailmerge_header"
        Case DataFinderTypes.dftFindFundraisingRequest
          vAttrs = "fundraising_request_number,contact_number,request_date,request_description,fundraising_request_stage,fundraising_status,fundraising_request_type,target_amount,pledged_amount,pledged_date,received_amount,received_date,expected_amount,gik_expected_amount,gik_pledged_amount,gik_pledged_date,total_gik_received_amount,latest_gik_received_date,number_of_payments,request_end_date,fundraising_business_type"
        Case DataFinderTypes.dftExternalReference
          vAttrs = "contact_number,external_reference,data_source"
      End Select
      Return vAttrs
    End Function

    Private Sub SetTransactionFinderType()
      If mvSelectItems.Exists("posted_to_nominal") Then
        mvTransactionFinderType = TransactionFinderTypes.tftProcessed
      Else
        If mvSelectItems.Exists("posted_to_nominal2") Then
          mvTransactionFinderType = TransactionFinderTypes.tftUnprocessed
        Else
          If mvSelectItems.Exists("posted_to_nominal3") Then
            mvTransactionFinderType = TransactionFinderTypes.tftProvisional
          Else
            If mvSelectItems.Exists("posted_to_nominal4") Then
              mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional
            End If
          End If
        End If
      End If
      If mvTransactionFinderType <> TransactionFinderTypes.tftProcessed Then
        mvResultColumns = Replace(mvResultColumns, "Status", "Provisional")
        mvResultColumns = Replace(mvResultColumns, "Posted", "AmendedOn")
        mvSelectColumns = Replace(mvSelectColumns, "Status", "Provisional")
        mvSelectColumns = Replace(mvSelectColumns, "Posted", "AmendedOn")
      End If
    End Sub

    Public ReadOnly Property TransactionFinderType() As TransactionFinderTypes
      Get
        Return mvTransactionFinderType
      End Get
    End Property

    Public ReadOnly Property SSHeadings() As String
      Get
        Select Case mvType
          Case DataFinderTypes.dftActions, DataFinderTypes.dftFindContacts, DataFinderTypes.dftFindOrganisations, DataFinderTypes.dftServiceProduct, DataFinderTypes.dftUniservPhoneBook, DataFinderTypes.dftStandardDocuments
            Return mvSelectHeadings
          Case DataFinderTypes.dftBatch, DataFinderTypes.dftOpenBatch
            Dim vString As String = DataSelectionText.String29110 'Batch,Provisional,Date,Type,Account,Meth,Source,Entries,Total,Trans,Trans Total,Detail Complete,Ready for Banking,Paying in Slip Printed,Picked,Cash Book,Posted,Paying In Slip Number,Currency,Created By,Created On,Department,Department Desc
            If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vString = vString.Replace(",Currency", "")
            Return vString
          Case DataFinderTypes.dftCCCA
            Return DataSelectionText.String29144 'Contact,Address,Surname,Initials,Number,Card Number,Start Date,Amount,Bank,C/Reason,Next Pay,Payment Plan Number"
          Case DataFinderTypes.dftCovenant
            Return DataSelectionText.String29106 'Contact,Address,Surname,Initials,Number,Term,Amount,Start Date,Deed,Net,C/Reason,Payment Plan Number
          Case DataFinderTypes.dftDirectDebit
            Return DataSelectionText.String29145 'Contact,Address,Surname,Initials,Number,Reference,Start Date,Sort Code,Account No,Amount,Bank,C/Reason,Next Pay,Payment Plan Number"
          Case DataFinderTypes.dftEvent
            Return DataSelectionText.String29120 'Number,Start Date,Event,Reference,No,Max,Waiting,Interest,Subject,Skill Level,Event Status,Venue,Template,Multi-Session,Class,Booking,Bookings Close,Branch,Chairperson"
          Case DataFinderTypes.dftEventBooking
            Return DataSelectionText.String29146 'Booking Number,Booking Date,Event,Booking Option,Quantity,Delegate,Event Number,Option Number,Sales Contact Number,Booking Notes
          Case DataFinderTypes.dftEventPersonnel
            Return DataSelectionText.String29147 'Contact Number,Address Number,Surname,Initials"
          Case DataFinderTypes.dftGAD
            Return DataSelectionText.String29148 'Contact,Surname,Initials,Number,Date,Type,Method,Confirmed,Start Date,End Date"
          Case DataFinderTypes.dftInvoice
            Return DataSelectionText.String29159 'Contact,Address,Surname,Initials,Number,Company,Account,Due Date
          Case DataFinderTypes.dftMeeting
            Return DataSelectionText.String22103 'Number,Date,Time,Description,Type,Location
          Case DataFinderTypes.dftMember
            Return mvEnv.GetBranchText(DataSelectionText.String29104)  'Contact,Address,Surname,Initials,Number,Type,Branch,Joined,No Members,Last Pay,Balance,Renewal Date,Membership Number,Payment Plan Number
          Case DataFinderTypes.dftLegacy
            Return DataSelectionText.String29164 'Contact,Surname,Initials,Number,ID,Status,Date of Death,Review Date,Label Name
          Case DataFinderTypes.dftPaymentPlan
            Return DataSelectionText.String29102 'Contact,Address,Surname,Initials,Number,Type,Provisional,Freq Amount,Next Due,Last Pay,Balance,Date,Renewal,C/Reason,Pay Meth,Reason
          Case DataFinderTypes.dftProduct
            Return DataSelectionText.String29108 'Product,Description,Extra Key,Sales Group,Don,Sub,Stock
          Case DataFinderTypes.dftStandingOrder, DataFinderTypes.dftManualSOReconciliation
            Return DataSelectionText.String29149 'Contact,Address,Surname,Initials,Number,Type,Start Date,Sort Code,Account No,Amount,Bank,C/Reason,Next Pay,Reference,Payment Plan Number"
          Case DataFinderTypes.dftTransaction
            Dim vString As String = DataSelectionText.String29150 'Batch,Trans,Provisional,No,Name,Date,Amount,Reference,Posted,Notes"
            If TransactionFinderType = TransactionFinderTypes.tftProcessed Then vString = Replace(vString, "Provisional", "Status")
            Return vString
          Case DataFinderTypes.dftVenue
            Return DataSelectionText.String29165 'Venue,Description,Venue Capacity,Town,Postcode
          Case DataFinderTypes.dftGiveAsYouEarn
            Dim vString As String = DataSelectionText.String29169 'Contact,Pledge,Donor ID,Surname,Initials,Employer,Employer No.,Agency
            Return vString & ",Start Date,Cancelled On,Source,Post To Cash Book"
          Case DataFinderTypes.dftInternalResource
            Return DataSelectionText.String29172 'Resource,Contact Number,Resource Contact,Product,Rate
            'vSSHeadings = "Resource,Contact Number,Name,Product,Description,Rate,Description"
          Case DataFinderTypes.dftPurchaseOrder
            Return DataSelectionText.String29175 'Contact,Address,Surname,Initials,Number,Type,Description,Start Date,No Payments,C/Reason,C/Date
          Case DataFinderTypes.dftContactMailingDocuments
            Return DataSelectionText.String29178 'Mailing Document Number,Mailing Template Code,Mailing Template Description,Contact Name,Created By,Created On,Mailing Code,Mailing Description,Earliest Fulfilment Date,Fulfilment Number,Fulfilled By,Fulfilled On
          Case DataFinderTypes.dftCommunication
            Return DataSelectionText.String22105 'Communication Number,Contact Number,Address Number,Surname,Initials,Device,Dialling Code,STD Code,Number,Extension,Ex Directory
          Case DataFinderTypes.dftPostTaxPayrollGiving
            Dim vString As String = DataSelectionText.String22108 'Contact,Pledge,Surname,Initials,Employer,Employer No.,Payroll Number
            Return vString & ",Start Date,Cancelled On,Source"
          Case DataFinderTypes.dftAppealCollections
            Return DataSelectionText.String22111 'Contact,Collection No,Campaign,Appeal,Collection,Collection Desc
          Case Else
            Return ""
        End Select
      End Get
    End Property

    Public Sub AddSelectItem(ByRef pName As String, ByRef pValue As String, ByVal pFieldType As CDBField.FieldTypes, Optional ByVal pFWO As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual)
      Dim vName As String
      Dim vCount As Integer

      If mvSelectItems.Exists(pName) Then
        vCount = 1 'Start with 2
        Do
          vCount = vCount + 1
          vName = pName & vCount
        Loop While mvSelectItems.Exists(vName)
      Else
        vName = pName
      End If
      mvSelectItems.Add(vName, pFieldType, pValue, pFWO)
    End Sub

    Public Function GetProductRestrictionSQL() As String
      Dim vWhere As String = ""
      Dim vMT As String
      Dim vMem As String

      AddClause(vWhere, "p.history_only")
      AddClause(vWhere, "subscription")
      AddClause(vWhere, "donation")
      AddClause(vWhere, "sponsorship_event")
      AddClause(vWhere, "stock_item")
      AddClause(vWhere, "course")
      AddClause(vWhere, "accommodation")
      AddClause(vWhere, "postage_packing")
      AddClause(vWhere, "sales_group")
      AddClause(vWhere, "pack_product")
      AddClause(vWhere, "uses_product_numbers") 'NFPCARE-100: uses_product_numbers was not passed in as a param
      AddClause(vWhere, "accrues_interest")
      AddClause(vWhere, "exam")
      If mvSelectItems.Exists("donation_or_sponsorship_event") Then
        If mvSelectItems.Item("donation_or_sponsorship_event").Bool Then
          If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
          vWhere = vWhere & "(donation = 'Y' OR sponsorship_event = 'Y')"
        End If
      End If
      If mvSelectItems.Exists("membership_product") Then
        vMT = "SELECT DISTINCT %s FROM membership_types"
        vMem = " (p.product %1 (" & Replace(vMT, "%s", "first_periods_product") & ") %2 p.product %1 (" & Replace(vMT, "%s", "subsequent_periods_product") & "))"
        If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
        If mvSelectItems.Item("membership_product").Bool Then
          vWhere = vWhere & Replace(Replace(vMem, "%2", "OR"), "%1", "IN")
        Else
          vWhere = vWhere & Replace(Replace(vMem, "%2", "AND"), "%1", "NOT IN")
        End If
      End If
      If mvSelectItems.Exists("campaign") And mvSelectItems.Exists("appeal") Then
        AddClause(vWhere, "campaign")
        AddClause(vWhere, "appeal")
      End If
      If mvSelectItems.Exists("FindProductType") Then
        'BR19642 - Called from Trader so do not want products that appear on the company_controls table. 
        vWhere = vWhere & String.Format(" AND product NOT IN ({0})", ProductsOnCompanyControlsSQL())
      End If
      Return vWhere
    End Function

    Private Sub AddClause(ByRef pWhere As String, ByVal pAttr As String, Optional ByRef pForceWildCard As Boolean = False)
      Dim vAttrName As String
      Dim vPos As Integer
      Dim vCDBField As CDBField
      Dim vValue As String = ""
      Dim vOperator As String = ""
      Dim vParams As CDBParameters

      vPos = InStr(pAttr, ".")
      If vPos > 0 Then
        vAttrName = Mid(pAttr, vPos + 1)
      Else
        vAttrName = pAttr
      End If
      If mvSelectItems.Exists(vAttrName) Then
        vCDBField = mvSelectItems(vAttrName)
        With vCDBField
          If .FieldType = CDBField.FieldTypes.cftCharacter Then
            If Len(.Value) = 0 Then
              If .WhereOperator = CDBField.FieldWhereOperators.fwoEqual Then
                vValue = " IS NULL"
              Else
                vValue = " IS NOT NULL"
              End If
            Else
              If .WhereOperator = CDBField.FieldWhereOperators.fwoNullOrEqual Then
                pAttr = "(" & vCDBField.Name & " "
                pAttr = pAttr & mvEnv.Connection.SQLLiteral("=", (vCDBField.FieldType), "")
                vValue = " OR " & vCDBField.Name & mvEnv.Connection.SQLLiteral("=", (vCDBField.FieldType), (vCDBField.Value)) & ")"
              ElseIf .WhereOperator = CDBField.FieldWhereOperators.fwoIn Then
                vParams = New CDBParameters
                vParams.InitFromUniqueList(.Value)
                vValue = " IN (" & vParams.InList & ")"
              Else
                If pForceWildCard Then
                  If InStr(.Value, "*") = 0 Then .Value = .Value & "*"
                End If
                If mvEnv.Connection.IsUnicodeField(vAttrName) Then
                  vValue = mvEnv.Connection.DBLikeOrEqual(.Value, CDBField.FieldTypes.cftUnicode)
                Else
                  vValue = mvEnv.Connection.DBLikeOrEqual(.Value)
                End If
              End If
            End If
          ElseIf .FieldType = CDBField.FieldTypes.cftMemo Then
            vValue = .Value
            If Not vValue.Contains("*") Then vValue = vValue & "*" 'Used to only do this for SQL Server but may as well always do it
            If mvEnv.Connection.IsUnicodeField(vAttrName) Then
              vValue = mvEnv.Connection.DBLikeOrEqual(vValue, CDBField.FieldTypes.cftUnicode)
            Else
              vValue = mvEnv.Connection.DBLikeOrEqual(vValue)
            End If
          Else
            'TA 10/3/04, Ref BR7230 Added following because CC Batch Find added a Restriction
            'of >= 0 on Transaction Total. However, not sure we should open this routine up to support
            'all operators; if so should be centralised w/Connection.WhereClause operation
            Select Case .WhereOperator
              Case CDBField.FieldWhereOperators.fwoIn
                vOperator = ""
                vValue = " IN (" & .Value & ")"
              Case CDBField.FieldWhereOperators.fwoGreaterThanEqual
                vOperator = ">="
              Case Else
                vOperator = "="
            End Select
            If Len(vOperator) > 0 Then vValue = mvEnv.Connection.SQLLiteral(vOperator, .FieldType, .Value)
          End If
        End With
      End If
      If vValue.Length > 0 Then
        If pWhere.Length > 0 And Trim(Right(pWhere, 6)) <> "WHERE" Then
          pWhere = pWhere & " AND " & pAttr & vValue
        Else
          pWhere = pWhere & pAttr & vValue
        End If
        mvWhereFlag = True
      End If
    End Sub
    ''' <summary>
    ''' Wrapper for ProductsOnCompanyControlsSQL() to allow it to be used standalone, all it needs is a connection, the rest is self contained.
    ''' </summary>
    ''' <param name="pEnv"></param>
    ''' <returns></returns>
    ''' <remarks>Required because DataFinder does not have a constructor than initialises mvEnv</remarks>
    Public Overloads Function ProductsOnCompanyControlsSQL(ByVal pEnv As CDBEnvironment) As String
      If mvEnv Is Nothing Then
        mvEnv = pEnv
      End If
      Return ProductsOnCompanyControlsSQL()
    End Function
    ''' <summary>
    ''' SQL to find all the DISTINCT products on table company_controls
    ''' </summary>
    ''' <returns>A string containing SQL to find all the DISTINCT products on table company_controls</returns>
    ''' <remarks>Uses maintenance _attributes to determine the product columns, intended for use in a NOT IN clause when interrogating the product table.</remarks>
    Public Overloads Function ProductsOnCompanyControlsSQL() As String
      Dim vStringBuilder As New StringBuilder
      Dim vField As String = "attribute_name"
      Dim vWhereFields As New CDBFields(New CDBField("table_Name", "company_controls"))
      'product attributes are those that are validated against the products table
      vWhereFields.Add("validation_table", "products")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vField, "maintenance_attributes", vWhereFields)
      Dim vDataTable As DataTable = vSQLStatement.GetDataTable
      ' Generate a select statement for each product column in company controls and use union to get a table of unique values  
      For vIndex = 0 To vDataTable.Rows.Count - 2
        ' All select statements need a UNION apart from the last so -2 not -1
        vStringBuilder.Append("SELECT ")
        vStringBuilder.Append(vDataTable.Rows(vIndex).Item(vField).ToString())
        vStringBuilder.Append(" As company_product")
        vStringBuilder.Append(" FROM company_controls")
        vStringBuilder.Append(" WHERE ")
        vStringBuilder.Append(vDataTable.Rows(vIndex).Item(vField).ToString())
        vStringBuilder.Append(" IS NOT NULL") 'This is essenatial if SQL is to be used in an IN or NOT IN clause on the products table
        vStringBuilder.AppendLine(" UNION")
      Next
      ' Add the last select, it doesn't need a UNION
      vStringBuilder.Append("SELECT ")
      vStringBuilder.Append(vDataTable.Rows(vDataTable.Rows.Count - 1).Item(vField).ToString())
      vStringBuilder.Append(" As company_product")
      vStringBuilder.Append(" FROM company_controls")
      vStringBuilder.Append(" WHERE ")
      vStringBuilder.Append(vDataTable.Rows(vDataTable.Rows.Count - 1).Item(vField).ToString())
      vStringBuilder.AppendLine(" IS NOT NULL") 'This is essenatial if SQL is to be used in an IN or NOT IN clause on the products table
      Return vStringBuilder.ToString()
    End Function
    Private Sub SetSOFinderType()
      If mvSelectItems.Exists("standing_order_type") Then
        mvSOFinderType = SOFinderType.softNormal
      Else
        If mvSelectItems.Exists("standing_order_type2") Then
          mvSOFinderType = SOFinderType.softCAF
        Else
          If mvSelectItems.Exists("standing_order_type3") Then
            mvSOFinderType = SOFinderType.softBoth
          End If
        End If
      End If
    End Sub

    Public ReadOnly Property SelectionSQL() As String
      Get
        Dim vSQL As String
        Dim vWhere As String = ""
        Dim vWhereFields As New CDBFields
        Dim vAccountLinked As Boolean
        Dim vError As Boolean
        Dim vContactFirst As Boolean
        Dim vSessionsFirst As Boolean
        Dim vOrganiserFirst As Boolean
        Dim vBookingOptionFirst As Boolean
        Dim vTopicsFirst As Boolean
        Dim vGayeAgencyFirst As Boolean
        Dim vAccountFirst As Boolean
        Dim vPayTableFirst As Boolean
        Dim vGAYEPledge As Boolean
        Dim vGayeBatch As Boolean
        Dim vIndex As Integer
        Dim vDate As Date
        Dim vActivities As Boolean
        Dim vJoinedToBatch As Boolean
        Dim vLinkBACS As Boolean
        Dim vPos As Integer
        Dim vSOTypeWhere As String = ""
        Dim vBACSFirst As Boolean
        Dim vTableList As String = ""
        Dim vRestriction As String
        Dim vConfirmedTransaction As New ConfirmedTransaction(mvEnv)
        Dim vCardSalesLinked As Boolean
        Dim vTemplate As Boolean

        Dim pCountOnly As Boolean       'Retain in case needed later
        Select Case mvType
          Case DataFinderTypes.dftTransaction
            SetTransactionFinderType()
          Case DataFinderTypes.dftStandingOrder, DataFinderTypes.dftManualSOReconciliation
            SetSOFinderType()
        End Select

        If (mvSelectItems.Count < 1 And (mvType <> DataFinderTypes.dftBatch And mvType <> DataFinderTypes.dftOpenBatch)) Or (mvType = DataFinderTypes.dftTransaction And mvSelectItems.Count = 1) Then
          RaiseError(DataAccessErrors.daeNoSelectionData)
        End If

        'Must select on the encrypted card numbers so we need to encrypt the number entered by the user.
        If mvSelectItems.Exists("card_number") Then
          mvSelectItems("card_number").Value = mvEnv.EncryptCreditCardNumber(mvSelectItems("card_number").Value)
        End If

        If pCountOnly Then
          vSQL = "SELECT count(*)  AS  record_count FROM "
        Else
          vSQL = "SELECT " & GetSelectedAttributes() & " FROM "
        End If
        Select Case mvType
          Case DataFinderTypes.dftBatch, DataFinderTypes.dftOpenBatch
            If mvType = DataFinderTypes.dftOpenBatch Then
              vSQL = vSQL & "open_batches ob, batches b "
              If mvSelectItems.Exists("department") Then
                vSQL = vSQL & ",users u, departments d WHERE "
              Else
                vSQL = vSQL & "WHERE "
              End If
              vWhere = "ob.batch_type IS NOT NULL AND ob.batch_number = b.batch_number"
            Else
              vSQL = vSQL & "batches b INNER JOIN users u ON b.batch_created_by = u.logname INNER JOIN departments d ON u.department = d.department WHERE "
            End If
            If mvSelectItems.Exists("department") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND"
              If mvType = DataFinderTypes.dftOpenBatch Then vWhere = vWhere & " b.batch_created_by = u.logname AND u.department = d.department AND"
              vWhere = vWhere & " d.department = '" & mvSelectItems("department").Value & "'"
            End If
            AddClause(vWhere, "b.batch_number")
            AddClause(vWhere, "b.batch_date")
            AddClause(vWhere, "b.bank_account")
            AddClause(vWhere, "b.batch_type")
            AddClause(vWhere, "b.currency_indicator")
            AddClause(vWhere, "b.paying_in_slip_number")
            AddClause(vWhere, "b.journal_number")
            AddClause(vWhere, "b.payment_method")
            AddClause(vWhere, "b.transaction_type")
            AddClause(vWhere, "b.batch_category")
            AddClause(vWhere, "b.number_of_entries")
            AddClause(vWhere, "b.batch_total")
            AddClause(vWhere, "b.transaction_total")
            AddClause(vWhere, "b.source")
            AddClause(vWhere, "b.provisional")
            AddClause(vWhere, "b.currency_code")
            AddClause(vWhere, "b.batch_created_by")
            AddClause(vWhere, "b.batch_created_on")
            AddClause(vWhere, "b.batch_analysis_code")
            If vWhere.Length > 0 Or mvType = DataFinderTypes.dftOpenBatch Then
              vSQL = vSQL & vWhere & " ORDER BY b.batch_number DESC"
              vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
            Else
              vError = True
            End If
          Case DataFinderTypes.dftCCCA
            If mvSelectItems.Exists("surname") Then vContactFirst = True
            If vContactFirst Then
              vSQL = vSQL & "contacts c, credit_card_authorities t"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
            Else
              vSQL = vSQL & "credit_card_authorities t, contacts c"
            End If
            vSQL = vSQL & ", contact_credit_cards ccc, orders o WHERE "
            AddClause(vWhere, "credit_card_authority_number")
            AddClause(vWhere, "bank_account")
            AddCancellationClause(vWhere)
            If Not vContactFirst Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "t.contact_number = c.contact_number"
              AddClause(vWhere, "initials")
            End If
            vWhere = vWhere & " AND t.credit_card_details_number = ccc.credit_card_details_number"
            AddClause(vWhere, "credit_card_type")
            AddClause(vWhere, "credit_card_number")
            AddClause(vWhere, "issuer")
            vWhere = vWhere & " AND t.order_number = o.order_number"
            AddClause(vWhere, "o.provisional")
            vSQL = vSQL & vWhere & " ORDER BY surname, initials"

          Case DataFinderTypes.dftCovenant
            AddClause(vWhere, "covenant_number")
            AddClause(vWhere, "t.contact_number")
            AddClause(vWhere, "covenanted_amount")
            AddClause(vWhere, "covenant_term")
            AddClause(vWhere, "start_date")
            AddClause(vWhere, "signature_date")
            AddClause(vWhere, "covenant_status")
            AddClause(vWhere, "last_tax_claim")
            AddClause(vWhere, "r185_return")
            If Not mvSelectItems.Exists("surname") Then
              If Len(vWhere) = 0 Then
                vError = True
              Else
                vSQL = vSQL & "covenants t, contacts c WHERE " & vWhere
                'vSQL = vSQL & " AND t.cancellation_reason IS NULL" TA 17/10 BC 3393 Removed Cancellation Restriction
                vSQL = vSQL & " AND c.contact_number = t.contact_number"
              End If
            Else
              vSQL = vSQL & "contacts c, covenants t WHERE "
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND c.contact_number = t.contact_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              'vSQL = vSQL & " AND t.cancellation_reason IS NULL" TA 17/10 BC 3393 Removed Cancellation Restriction
            End If
            If mvSelectItems.Exists("provisional") Then
              vPos = InStr(1, vSQL, "WHERE")
              vSQL = Left(vSQL, vPos - 2) & ", orders o " & Mid(vSQL, vPos - 1)
              vSQL = vSQL & " AND o.order_number = t.order_number"
              AddClause(vSQL, "o.provisional")
            End If
            vSQL = vSQL & " ORDER BY surname, initials"

          Case DataFinderTypes.dftDirectDebit
            If mvSelectItems.Exists("iban_number") AndAlso Not mvEnv.CheckIbanNumber(mvSelectItems("iban_number").Value) Then RaiseError(DataAccessErrors.daeInvalidIbanNumber) 'Validate IBAN Number

            vSQL = Replace(vSQL, ",reference,", "," & mvConn.DBSpecialCol("", "reference") & ",")
            'Only need to select from the bacs_amendments table if those attrs selected
            If mvSelectItems.Exists("bacs_advice_reason") Or mvSelectItems.Exists("amended_by") Or mvSelectItems.Exists("amended_on") Then
              vLinkBACS = True
            End If
            vSQL = "SELECT DISTINCT" & Mid(vSQL, 7)
            If mvSelectItems.Exists("surname") Then
              vContactFirst = True
            ElseIf mvSelectItems.Exists("account_number") = True Or mvSelectItems.Exists("sort_code") = True Or mvSelectItems.Exists("iban_number") = True Or mvSelectItems.Exists("bic_code") = True Then
              vAccountFirst = True
            ElseIf vLinkBACS = True Then
              vBACSFirst = True
            Else
              vPayTableFirst = True
            End If

            If vContactFirst Then
              vSQL = vSQL & "contacts c, direct_debits t, contact_accounts ca, orders o "
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "direct_debit_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND t.bank_details_number = ca.bank_details_number"
              AddBankDetailsClause(vWhere)
            End If
            If vPayTableFirst Then
              vSQL = vSQL & "direct_debits t, contacts c, contact_accounts ca, orders o "
              AddClause(vWhere, "direct_debit_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "t.contact_number = c.contact_number"
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND t.bank_details_number = ca.bank_details_number"
              AddBankDetailsClause(vWhere)
            End If
            If vAccountFirst Then
              vSQL = vSQL & "contact_accounts ca, direct_debits t, contacts c, orders o "
              AddBankDetailsClause(vWhere)
              vWhere = vWhere & " AND t.bank_details_number = ca.bank_details_number"
              AddClause(vWhere, "direct_debit_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND t.contact_number = c.contact_number "
              AddClause(vWhere, "initials")
            End If
            If vBACSFirst Then
              vSQL = vSQL & "bacs_amendments ba, direct_debits t, contacts c, contact_accounts ca, orders o "
              AddClause(vWhere, "bacs_advice_reason")
              AddClause(vWhere, "ba.amended_by")
              AddClause(vWhere, "ba.amended_on")
              vWhere = vWhere & " AND t.direct_debit_number = ba.direct_debit_number"
              AddClause(vWhere, "direct_debit_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND t.contact_number = c.contact_number "
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND t.bank_details_number = ca.bank_details_number"
            Else
              If vLinkBACS Then
                vSQL = vSQL & ", bacs_amendments ba"
                vWhere = vWhere & " AND t.direct_debit_number = ba.direct_debit_number"
                AddClause(vWhere, "bacs_advice_reason")
                AddClause(vWhere, "ba.amended_by")
                AddClause(vWhere, "ba.amended_on")
              End If
            End If

            vWhere = vWhere & " AND t.order_number = o.order_number"
            AddClause(vWhere, "o.provisional")
            vSQL = vSQL & " WHERE " & vWhere & " ORDER BY c.surname, c.initials"

          Case DataFinderTypes.dftEvent
            vSQL = "SELECT DISTINCT" & Mid(vSQL, 7)
            AddClause(vWhere, "ev.event_number")
            AddClause(vWhere, "ev.event_desc")
            AddClause(vWhere, "ev.venue")
            AddClause(vWhere, "ev.event_status")
            AddClause(vWhere, "ev.branch")
            AddClause(vWhere, "ev.event_reference")
            AddClause(vWhere, "ev.free_of_charge")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventGroups) Then AddClause(vWhere, "event_group")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventClass) Then AddClause(vWhere, "ev.event_class")
            If mvSelectItems.Exists("event_restriction") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & mvSelectItems("event_restriction").Value
            End If
            If mvSelectItems.Exists("checkbox") Or mvSelectItems.Exists("template") Then
              vTemplate = True
            End If
            If vTemplate Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              If mvSelectItems.Exists("template") AndAlso mvSelectItems("template").Bool = False Then
                vWhere = vWhere & "template = 'N'"
              Else
                vWhere = vWhere & "template = 'Y'"
              End If
            Else
              CheckDateRange()
              If mvSelectItems.Exists("date") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "ev.start_date" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
              End If
              If mvSelectItems.Exists("date2") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "ev.start_date" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
              End If
            End If
            If mvSelectItems.Exists("booking") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "booking = 'Y' AND ( bookings_close IS NULL or bookings_close " & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, TodaysDate) & ")"
            End If
            If (mvSelectItems.Exists("subject") Or mvSelectItems.Exists("skill_level")) And vWhere = "" Then
              vSQL = vSQL & "sessions s, events ev"
              vSessionsFirst = True
              AddClause(vWhere, "s.subject")
              AddClause(vWhere, "s.skill_level")
              vWhere = vWhere & " AND ev.event_number = s.event_number AND s.session_type = '0'"
            ElseIf mvSelectItems.Exists("organiser") And vWhere = "" Then
              vSQL = vSQL & "event_organisers eo, events ev"
              AddClause(vWhere, "eo.organiser")
              vWhere = vWhere & " AND eo.event_number = ev.event_number"
              vOrganiserFirst = True
            ElseIf mvSelectItems.Exists("product") And vWhere = "" Then
              vSQL = vSQL & "event_booking_options ebo, events ev"
              AddClause(vWhere, "ebo.product")
              vWhere = vWhere & " AND ebo.event_number = ev.event_number"
              vBookingOptionFirst = True
            ElseIf mvSelectItems.Exists("topic") And vWhere = "" Then
              vSQL = vSQL & "event_topics et, events ev"
              AddClause(vWhere, "et.topic")
              AddClause(vWhere, "et.sub_topic")
              vWhere = vWhere & " AND et.event_number = ev.event_number"
              vTopicsFirst = True
            Else
              vSQL = vSQL & "events ev"
            End If

            If Not vSessionsFirst Then
              vSQL = vSQL & ", sessions s"
              'Always join to session with session_type = '0' as it is an extension of the event record
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "ev.event_number = s.event_number AND s.session_type = '0'"
              AddClause(vWhere, "s.subject")
              AddClause(vWhere, "s.skill_level")
            End If

            If Not vOrganiserFirst Then
              If mvSelectItems.Exists("organiser") Then
                vSQL = vSQL & ", event_organisers eo"
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "ev.event_number = eo.event_number"
                AddClause(vWhere, "eo.organiser")
              End If
            End If

            If Not vBookingOptionFirst Then
              If mvSelectItems.Exists("product") Then
                vSQL = vSQL & ", event_booking_options ebo"
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "ev.event_number = ebo.event_number"
                AddClause(vWhere, "ebo.product")
              End If
            End If

            If Not vTopicsFirst Then
              If mvSelectItems.Exists("topic") Then
                vSQL = vSQL & ", event_topics et"
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "ev.event_number = et.event_number"
                AddClause(vWhere, "et.topic")
                AddClause(vWhere, "et.sub_topic")
              End If
            End If

            If mvSelectItems.Exists("department") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "ev.department = '" & mvSelectItems("department").Value & "'"
            End If

            vSQL = vSQL & ", subjects sb, skill_levels sl, venues v"
            vWhere = vWhere & " AND sb.subject = s.subject "
            vWhere = vWhere & " AND sl.skill_level = s.skill_level AND v.venue = ev.venue "
            vSQL = vSQL & " WHERE " & vWhere & " ORDER BY ev.start_date, event_desc"

          Case DataFinderTypes.dftEventBooking
            If mvSelectItems.Exists("booking_date") And mvSelectItems.Exists("booking_date2") Then
              If CDate(mvSelectItems("booking_date2").Value) < CDate(mvSelectItems("booking_date").Value) Then RaiseError(DataAccessErrors.daeInvalidDateRange)
            End If
            AddClause(vWhere, "eb.contact_number")
            AddClause(vWhere, "eb.booking_number")
            If mvSelectItems.Exists("booking_date") Then vWhere = vWhere & " AND eb.booking_date " & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvSelectItems("booking_date").Value)
            If mvSelectItems.Exists("booking_date2") Then vWhere = vWhere & " AND eb.booking_date " & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, mvSelectItems("booking_date2").Value)
            If mvSelectItems.Exists("booking_status") Then
              AddClause(vWhere, "booking_status")
            Else
              vWhere = vWhere & " AND booking_status not in ('" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsExternal) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsCancelled) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsInterested) & "','" & EventBooking.GetBookingStatusCode(EventBooking.EventBookingStatuses.ebsAwaitingAcceptance) & "')"
            End If
            vWhere = vWhere & " AND e.event_number = eb.event_number AND ebo.option_number = eb.option_number AND d.booking_number = eb.booking_number AND c.contact_number = d.contact_number"
            vSQL = vSQL & "event_bookings eb, events e, event_booking_options ebo, delegates d,contacts c WHERE " & vWhere & " ORDER BY e.event_desc, ebo.option_desc"

          Case DataFinderTypes.dftEventPersonnel
            vSQL = vSQL & "personnel p, contacts c WHERE "
            If mvSelectItems.Exists("activity1") Then
              vIndex = 1
              vWhere = vWhere & "p.contact_number in (SELECT contact_number from contact_categories WHERE "
              Do
                If vIndex > 1 Then vWhere = vWhere & " OR "
                vWhere = vWhere & "(activity = '" & mvSelectItems("activity" & vIndex).Value & "'"
                If mvSelectItems.Exists("activity_value" & vIndex) Then
                  vWhere = vWhere & " AND activity_value = '" & mvSelectItems("activity_value" & vIndex).Value & "'"
                End If
                vWhere = vWhere & ")"
                vIndex = vIndex + 1
              Loop While mvSelectItems.Exists("activity" & vIndex)
              vWhere = vWhere & ") AND "
            End If
            vWhere = vWhere & " c.contact_number = p.contact_number"
            AddClause(vWhere, "surname")
            AddClause(vWhere, "initials")
            vSQL = vSQL & vWhere & " ORDER BY c.surname, c.initials"

          Case DataFinderTypes.dftServiceProduct
            If mvSelectItems.Exists("postcode") Then
              vSQL = vSQL & "addresses a INNER JOIN contacts c ON a.address_number = c.address_number"
              AddClause(vWhere, "a.postcode")
              AddClause(vWhere, "c.contact_number")
              AddClause(vWhere, "surname")
              AddClause(vWhere, "contact_group")
              AddClause(vWhere, "geographical_region_type")
              AddClause(vWhere, "geographical_region")
            Else
              vSQL = vSQL & "contacts c INNER JOIN addresses a ON c.address_number = a.address_number"
              AddClause(vWhere, "c.contact_number")
              AddClause(vWhere, "surname")
              AddClause(vWhere, "contact_group")
              AddClause(vWhere, "a.postcode")
              AddClause(vWhere, "geographical_region_type")
              AddClause(vWhere, "geographical_region")
            End If
            vSQL = vSQL & " INNER JOIN service_products sp ON c.contact_number = sp.contact_number"
            vSQL = vSQL & " LEFT OUTER JOIN address_geographical_regions agr ON a.postcode = agr.postcode"
            vSQL = vSQL & " WHERE " & vWhere
            vSQL = vSQL & " ORDER BY a.postcode,surname" & mvEnv.Connection.DBForceOrder
            vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)

          Case DataFinderTypes.dftGAD
            If mvSelectItems.Exists("surname") Then vContactFirst = True
            If vContactFirst Then
              vSQL = vSQL & "contacts c, gift_aid_declarations t"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
            Else
              vSQL = vSQL & "gift_aid_declarations t, contacts c"
            End If
            vSQL = vSQL & " WHERE "
            AddClause(vWhere, "declaration_number")
            AddClause(vWhere, "declaration_date")
            AddClause(vWhere, "declaration_type")
            AddClause(vWhere, "method")
            AddClause(vWhere, "start_date")
            AddClause(vWhere, "end_date")
            If Not vContactFirst Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "t.contact_number = c.contact_number"
              AddClause(vWhere, "initials")
            End If
            vSQL = vSQL & vWhere & " ORDER BY surname, initials"

          Case DataFinderTypes.dftInvoice
            If mvSelectItems.Exists("surname") Then vContactFirst = True
            If vContactFirst Then
              vSQL = vSQL & "contacts c, invoices t"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
            Else
              vSQL = vSQL & "invoices t, contacts c"
            End If
            vSQL = vSQL & " WHERE "
            AddClause(vWhere, "invoice_number")
            AddClause(vWhere, "invoice_date")
            AddClause(vWhere, "company")
            AddClause(vWhere, "sales_ledger_account")

            If Not vContactFirst Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "t.contact_number = c.contact_number"
              AddClause(vWhere, "initials")
            End If
            vSQL = vSQL & vWhere & " ORDER BY surname, initials"

          Case DataFinderTypes.dftMeeting
            vSQL = Replace(vSQL, "MEETING_DATE", "meeting_date")
            vSQL = Replace(vSQL, "MEETING_TIME", "meeting_date AS meeting_time")
            If mvSelectItems.Exists("checkbox") Or mvSelectItems.Exists("checkbox2") Or mvSelectItems.Exists("contact_number") Or mvSelectItems.Exists("contact_number2") Then
              'Meetings to be attended
              vSQL = vSQL & "meeting_links ml, "
              If mvSelectItems.Exists("checkbox") Or mvSelectItems.Exists("checkbox2") Then
                vWhere = "contact_number = " & mvEnv.User.ContactNumber & " AND link_type = 'W'"
                If mvSelectItems.Exists("checkbox2") Then vWhere = vWhere & " AND notified = 'N'"
                If mvSelectItems.Exists("checkbox") Then vWhere = vWhere & " AND attended IS NULL"
              ElseIf mvSelectItems.Exists("contact_number") Then
                vWhere = vWhere & "contact_number = " & mvSelectItems("contact_number").IntegerValue & " AND link_type = 'W'"
              ElseIf mvSelectItems.Exists("contact_number2") Then
                vWhere = vWhere & "contact_number = " & mvSelectItems("contact_number2").IntegerValue & " AND link_type = 'R'"
              End If
              vWhere = vWhere & " AND ml.meeting_number = m.meeting_number"
            Else
              AddClause(vWhere, "meeting_number")
              AddClause(vWhere, "meeting_desc")
              AddClause(vWhere, "m.meeting_type")
              AddClause(vWhere, "m.meeting_location")
              CheckDateRange()
              If mvSelectItems.Exists("meeting_date") Then 'Meeting Date retained in case database not yet updated
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vDate = CDate(mvSelectItems("meeting_date").Value)
                vWhere = vWhere & "meeting_date" & mvConn.SQLLiteral(">", CDBField.FieldTypes.cftTime, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, vDate) & " 23:59") & " AND meeting_date" & mvConn.SQLLiteral("<", CDBField.FieldTypes.cftTime, DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, vDate) & " 00:00")
              ElseIf mvSelectItems.Exists("date") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vDate = CDate(mvSelectItems("date").Value)
                vWhere = vWhere & "meeting_date" & mvConn.SQLLiteral(">", CDBField.FieldTypes.cftTime, vDate & " 00:00")
                If mvSelectItems.Exists("date2") Then
                  vDate = CDate(mvSelectItems("date2").Value)
                  vWhere = vWhere & " AND meeting_date" & mvConn.SQLLiteral("<", CDBField.FieldTypes.cftTime, vDate & " 23:59")
                End If
              ElseIf mvSelectItems.Exists("date2") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vDate = CDate(mvSelectItems("date2").Value)
                vWhere = vWhere & "meeting_date" & mvConn.SQLLiteral("<", CDBField.FieldTypes.cftTime, vDate.AddDays(1).ToString(CAREDateFormat) & " 00:00")
              ElseIf mvSelectItems.Exists("checkbox3") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "meeting_date" & mvConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, Today.AddMonths(-1).ToString(CAREDateFormat)) & " AND meeting_date" & mvConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, Today.AddMonths(1).ToString(CAREDateFormat))
              End If
            End If
            vWhere = vWhere & " AND m.meeting_type = mt.meeting_type AND m.meeting_location = mlo.meeting_location"
            vSQL = vSQL & "meetings m, meeting_types mt, meeting_locations mlo WHERE " & vWhere & " ORDER BY meeting_date DESC"

          Case DataFinderTypes.dftMember
            AddClause(vWhere, "t.contact_number")
            AddClause(vWhere, "member_number")
            AddClause(vWhere, "t.membership_type")
            AddClause(vWhere, "t.branch")
            AddClause(vWhere, "number_of_members")
            AddClause(vWhere, "joined")
            AddClause(vWhere, "forenames")
            AddClause(vWhere, "preferred_forename")
            'BR 19411 To find just using pay method
            AddClause(vWhere, "o.payment_method")
            AddCancellationClause(vWhere)
            If mvSelectItems.Exists("current") AndAlso mvSelectItems("current").Bool Then
              If vWhere.Length > 0 Then vWhere &= " AND "
              vWhere &= " o.cancellation_reason IS NULL"
            End If
            'Adding the tables
            If mvSelectItems.Exists("ContactType") Then
              If mvSelectItems("ContactType").Value = "O" And Not mvSelectItems.Exists("name") Then
                vSQL = vSQL & "organisations og,"
              End If
            End If
            If mvSelectItems.Exists("email_address2") Then
              If mvSelectItems("ContactType").Value = "O" Then
                vSQL = vSQL & "organisation_addresses oa,communications cm,"
              Else
                vSQL = vSQL & "communications cm,"
              End If
            End If
            If mvSelectItems.Exists("name") Then
              vSQL = vSQL & "organisations og, members t, contacts c, orders o, addresses a"
              vSQL = vSQL & " WHERE name" & mvConn.DBLikeOrEqual(mvSelectItems("name").Value, CDBField.FieldTypes.cftUnicode)
              vSQL = vSQL & " AND og.address_number = t.address_number AND a.address_number = t.address_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              vSQL = vSQL & " AND c.contact_number = t.contact_number"
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND t.order_number = o.order_number"
            ElseIf mvSelectItems.Exists("postcode") Or mvSelectItems.Exists("town") Then
              vSQL = vSQL & "addresses a, members t, contacts c, orders o WHERE "
              AddClause(vSQL, "postcode")
              AddClause(vSQL, "town")
              vSQL = vSQL & " AND a.address_number = t.address_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              vSQL = vSQL & " AND c.contact_number = t.contact_number"
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND t.order_number = o.order_number"
            ElseIf Not mvSelectItems.Exists("surname") Then
              If Len(vWhere) = 0 Then
                vError = True
              Else
                vSQL = vSQL & "members t, contacts c, orders o, addresses a WHERE " & vWhere
                vSQL = vSQL & " AND t.order_number = o.order_number AND a.address_number = t.address_number"
                vSQL = vSQL & " AND c.contact_number = t.contact_number"
              End If
            Else
              vSQL = vSQL & "contacts c, members t, orders o, addresses a WHERE "
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND c.contact_number = t.contact_number AND a.address_number = t.address_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              vSQL = vSQL & " AND t.order_number = o.order_number"
              'If vNotCancelled Then vSQL = vSQL & " AND t.cancellation_reason IS NULL" TA 17/10 BC 3393 Removed Cancellation Restriction
            End If
            AddClause(vSQL, "o.payment_method")
            AddClause(vSQL, "o.provisional")
            AddClause(vSQL, "o.bankers_order")
            If mvSelectItems.Exists("membership_number") Then
              'This field will only come from Smart Client
              vSQL = vSQL & " AND t.membership_number = " & mvSelectItems("membership_number").IntegerValue
            End If
            'Joins and Where Condition
            If mvSelectItems.Exists("email_address2") Then
              vSQL = vSQL & " And (cm.valid_from IS NULL or cm.valid_from " & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDate) & ")  And (cm.valid_to " & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, TodaysDate) & " or cm.valid_to IS NULL) "
              If mvSelectItems("ContactType").Value = "O" Then
                vSQL = vSQL & " And oa.address_number= cm.address_number"
                vSQL = vSQL & " And cm.contact_number IS NULL"
              Else
                vSQL = vSQL & " And cm.contact_number= t.contact_number"
              End If
              vSQL = vSQL & " AND " & mvConn.DBSpecialCol("cm", "number") & mvConn.DBLike(mvSelectItems("email_address2").Value)
            End If
            If mvSelectItems.Exists("ContactType") Then
              If mvSelectItems("ContactType").Value = "C" Then
                vSQL = vSQL & " And c.contact_type <> 'O'"
              ElseIf mvSelectItems("ContactType").Value = "O" Then
                vSQL = vSQL & " And c.contact_type = 'O'"
                vSQL = vSQL & " And og.organisation_number= t.contact_number And oa.organisation_number= og.organisation_number"
                If mvSelectItems.Exists("postcode") Then
                  AddClause(vSQL, "postcode")
                End If
              End If
            End If

            vSQL = vSQL & " ORDER BY surname, initials"

          Case DataFinderTypes.dftLegacy
            AddClause(vWhere, "t.contact_number")
            AddClause(vWhere, "legacy_number")
            AddClause(vWhere, "legacy_id")
            AddClause(vWhere, "t.legacy_status")
            AddClause(vWhere, "date_of_death")
            If mvSelectItems.Exists("date") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "review_date" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
            End If
            If mvSelectItems.Exists("date2") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "review_date" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
            End If
            If mvSelectItems.Exists("postcode") Or mvSelectItems.Exists("town") Then
              vSQL = vSQL & "addresses a, contact_legacies t, contacts c, legacy_statuses ls WHERE "
              AddClause(vSQL, "postcode")
              AddClause(vSQL, "town")
              vSQL = vSQL & " AND a.address_number = c.address_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              vSQL = vSQL & " AND c.contact_number = t.contact_number"
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
            ElseIf (Not mvSelectItems.Exists("surname") AndAlso Not mvSelectItems.Exists("initials")) Then
              If Len(vWhere) = 0 Then
                vError = True
              Else
                vSQL = vSQL & "contact_legacies t, contacts c, legacy_statuses ls WHERE " & vWhere
                vSQL = vSQL & " AND c.contact_number = t.contact_number"
              End If
            Else
              vSQL = vSQL & "contacts c, contact_legacies t, legacy_statuses ls WHERE "
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND c.contact_number = t.contact_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
            End If
            vSQL = vSQL & " AND ls.legacy_status = t.legacy_status"
            vSQL = vSQL & " ORDER BY surname, initials"

          Case DataFinderTypes.dftPaymentPlan
            AddClause(vWhere, "t.order_number")
            AddClause(vWhere, "t.contact_number")
            AddClause(vWhere, "order_type")
            AddClause(vWhere, "frequency_amount")
            AddClause(vWhere, "ops.next_payment_due")
            AddClause(vWhere, "provisional")
            AddClause(vWhere, "their_reference", True)
            AddClause(vWhere, "bankers_order")
            vSQL = vSQL.Replace(",order_number,", ",t.order_number,")
            vSQL = vSQL.Replace(",next_payment_due,", ",COALESCE(ops.next_payment_due, t.next_payment_due) AS next_payment_due,")
            If (Not mvSelectItems.Exists("surname") AndAlso Not mvSelectItems.Exists("initials")) Then
              If Len(vWhere) = 0 Then
                vError = True
              Else
                vSQL &= "orders t "
                vSQL &= "INNER JOIN contacts c "
                vSQL &= "ON c.contact_number = t.contact_number "
                vSQL &= "LEFT OUTER JOIN (SELECT order_number, MIN(due_date) AS next_payment_due FROM order_payment_schedule WHERE scheduled_payment_status IN ('D', 'P', 'V') GROUP BY order_number) ops "
                vSQL &= "ON ops.order_number = t.order_number "
                vSQL &= " WHERE " & vWhere
              End If
            Else
              vSQL &= "orders t "
              vSQL &= "INNER JOIN contacts c "
              vSQL &= "ON c.contact_number = t.contact_number "
              vSQL &= "LEFT OUTER JOIN (SELECT order_number, MIN(due_date) AS next_payment_due FROM order_payment_schedule WHERE scheduled_payment_status IN ('D', 'P', 'V') GROUP BY order_number) ops "
              vSQL &= "ON ops.order_number = t.order_number "
              vSQL &= " WHERE " & vWhere
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
            End If
            vSQL = vSQL & " ORDER BY surname, initials"

          Case DataFinderTypes.dftProduct
            AddClause(vWhere, "product")
            AddClause(vWhere, "product_desc")
            AddClause(vWhere, "extra_key")
            AddClause(vWhere, "secondary_group")
            AddClause(vWhere, "product_category")
            If mvSelectItems.Exists("product_restriction") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & mvSelectItems("product_restriction").Value
              AddClause(vWhere, "sales_group")
            Else
              vRestriction = GetProductRestrictionSQL()
              If Len(vWhere) > 0 And Len(vRestriction) > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & vRestriction
            End If
            If mvSelectItems.Exists("source") Then
              If vWhere.Length > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "p.product IN (SELECT DISTINCT spa.product from segments s, segment_product_allocation spa WHERE s.source = '" & mvSelectItems("source").Value & "' AND s.campaign = spa.campaign AND s.appeal = spa.appeal AND s.segment = spa.segment)"
            End If
            If mvSelectItems.Exists("campaign") And mvSelectItems.Exists("appeal") Then
              vSQL = vSQL & "products p,appeal_resources ar WHERE " & vWhere & If(vWhere.Length > 0, " AND p.product = ar.product", " p.product = ar.product") & " ORDER BY p.product"
              vSQL = Replace(vSQL, " product ", " p.product ")
            Else
              vSQL = vSQL & "products p WHERE " & vWhere & " ORDER BY product"
            End If

          Case DataFinderTypes.dftPurchaseOrder
            AddClause(vWhere, "purchase_order_number")
            AddClause(vWhere, "purchase_order_type")
            AddClause(vWhere, "start_date")
            AddCancellationClause(vWhere)
            If mvSelectItems.Exists("contact_number") Then AddSelectItem("payee_contact_number", (mvSelectItems("contact_number").Value), CDBField.FieldTypes.cftLong)

            If mvSelectItems.Exists("name") Then
              vSQL = vSQL & "organisations og, purchase_orders t, contacts c"
              vSQL = vSQL & " WHERE name" & mvConn.DBLikeOrEqual(mvSelectItems("name").Value, CDBField.FieldTypes.cftUnicode)
              vSQL = vSQL & " AND og.address_number = t.payee_address_number"
              AddClause(vSQL, "t.payee_contact_number")
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
              vSQL = vSQL & " AND c.contact_number = t.payee_contact_number"
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
            ElseIf Not mvSelectItems.Exists("surname") And Not mvSelectItems.Exists("payee_contact_number") Then
              If Len(vWhere) = 0 Then
                vError = True
              Else
                vSQL = vSQL & "purchase_orders t, contacts c WHERE " & vWhere
                vSQL = vSQL & " AND c.contact_number = t.payee_contact_number"
              End If
            Else
              vSQL = vSQL & "contacts c, purchase_orders t WHERE "
              AddClause(vSQL, "c.contact_number")
              AddClause(vSQL, "surname")
              AddClause(vSQL, "initials")
              vSQL = vSQL & " AND c.contact_number = t.payee_contact_number"
              If vWhere.Length > 0 Then vSQL = vSQL & " AND " & vWhere
            End If

          Case DataFinderTypes.dftStandingOrder, DataFinderTypes.dftManualSOReconciliation
            If mvType = DataFinderTypes.dftManualSOReconciliation Then
              CheckRequiredSORecFields()
            Else
              If mvSelectItems.Count = 1 Then RaiseError(DataAccessErrors.daeNoSelectionData)
            End If
            If mvSelectItems.Exists("surname") Then
              vContactFirst = True
            ElseIf mvSelectItems.Exists("account_number") = True Or mvSelectItems.Exists("sort_code") = True Or mvSelectItems.Exists("iban_number") Or mvSelectItems.Exists("bic_code") Then
              vAccountFirst = True
            Else
              vPayTableFirst = True
            End If

            If mvSelectItems.Exists("iban_number") AndAlso Not mvEnv.CheckIbanNumber(mvSelectItems("iban_number").Value) Then RaiseError(DataAccessErrors.daeInvalidIbanNumber)

            vSQL = vSQL.Replace(",next_payment_due,", ",COALESCE(ops.next_payment_due, o.next_payment_due) AS next_payment_due,")

            If vContactFirst Then
              vSQL &= "contacts c "
              vSQL &= "INNER JOIN bankers_orders t "
              vSQL &= "ON c.contact_number = t.contact_number "
              vSQL &= "INNER JOIN contact_accounts ca "
              vSQL &= "ON t.bank_details_number = ca.bank_details_number "
              vSQL &= "INNER JOIN orders o "
              vSQL &= "ON t.order_number = o.order_number "
              vSQL &= "LEFT OUTER JOIN (SELECT order_number, MIN(due_date) AS next_payment_due FROM order_payment_schedule WHERE scheduled_payment_status IN ('D', 'P', 'V') GROUP BY order_number) ops "
              vSQL &= "ON ops.order_number = t.order_number "

              vSQL = vSQL & "WHERE "
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              AddClause(vWhere, "bankers_order_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              AddBankDetailsClause(vWhere)
            End If

            If vPayTableFirst Then
              vSQL &= "bankers_orders t "
              vSQL &= "INNER JOIN contacts c "
              vSQL &= "ON c.contact_number = t.contact_number "
              vSQL &= "INNER JOIN contact_accounts ca "
              vSQL &= "ON t.bank_details_number = ca.bank_details_number "
              vSQL &= "INNER JOIN orders o "
              vSQL &= "ON t.order_number = o.order_number "
              vSQL &= "LEFT OUTER JOIN (SELECT order_number, MIN(due_date) AS next_payment_due FROM order_payment_schedule WHERE scheduled_payment_status IN ('D', 'P', 'V') GROUP BY order_number) ops "
              vSQL &= "ON ops.order_number = t.order_number "
              vSQL &= "WHERE "
              AddClause(vWhere, "bankers_order_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              AddClause(vWhere, "initials")
              AddBankDetailsClause(vWhere)
            End If

            If vAccountFirst Then
              vSQL &= "contact_accounts ca "
              vSQL &= "INNER JOIN bankers_orders t "
              vSQL &= "ON t.bank_details_number = ca.bank_details_number "
              vSQL &= "INNER JOIN contacts c "
              vSQL &= "ON c.contact_number = t.contact_number "
              vSQL &= "INNER JOIN orders o "
              vSQL &= "ON t.order_number = o.order_number "
              vSQL &= "LEFT OUTER JOIN (SELECT order_number, MIN(due_date) AS next_payment_due FROM order_payment_schedule WHERE scheduled_payment_status IN ('D', 'P', 'V') GROUP BY order_number) ops "
              vSQL &= "ON ops.order_number = t.order_number "
              vSQL &= "WHERE "
              AddBankDetailsClause(vWhere)
              AddClause(vWhere, "bankers_order_number")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "bank_account")
              AddCancellationClause(vWhere)
              AddClause(vWhere, "initials")
            End If
            ' End of changes
            AddClause(vWhere, "o.provisional")

            If mvType = DataFinderTypes.dftManualSOReconciliation And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments) = True Then
              If vWhere.Length > 0 And Trim(Right(vWhere, 6)) <> "WHERE" Then vWhere = vWhere & " AND "
              vWhere = vWhere & "o.payment_schedule_amended_on IS NOT NULL "
            End If

            AddPaymentsDueClause(vWhere)
            Select Case mvSOFinderType
              Case SOFinderType.softCAF
                vSOTypeWhere = "(standing_order_type IS NOT NULL AND standing_order_type = 'C')"
              Case SOFinderType.softNormal
                vSOTypeWhere = "(standing_order_type IS NULL OR (standing_order_type IS NOT NULL AND standing_order_type = 'B'))"
              Case SOFinderType.softBoth
                vSOTypeWhere = "(standing_order_type IS NULL OR standing_order_type in ('B', 'C'))"
              Case Else
                'And what should we put here then.  
            End Select
            If vSOTypeWhere.Length > 0 Then
              If vWhere.Length > 0 And Trim(Right(vWhere, 6)) <> "WHERE" And Trim(Right(vWhere, 4)) <> "AND" Then
                vWhere = vWhere & " AND " & vSOTypeWhere
              Else
                vWhere = vWhere & vSOTypeWhere
              End If
            End If

            vSQL = vSQL & vWhere & " ORDER BY surname, initials"

          Case DataFinderTypes.dftTransaction
            If mvSelectItems.Exists("iban_number") AndAlso Not mvEnv.CheckIbanNumber(mvSelectItems("iban_number").Value) Then RaiseError(DataAccessErrors.daeInvalidIbanNumber)
            If mvTransactionFinderType = TransactionFinderTypes.tftProcessed Then
              If mvSelectItems.Exists("member_number") Then
                vSQL = vSQL & "members m, orders o, order_payment_history oph, financial_history fh "
                vWhere = "m.member_number = '" & mvSelectItems("member_number").Value & "' AND m.order_number = o.order_number "
                vWhere = vWhere & "AND o.order_number = oph.order_number "
                vWhere = vWhere & "AND oph.batch_number = fh.batch_number AND oph.transaction_number = fh.transaction_number"
                If mvSelectItems.Exists("paying_in_slip_number") = True Or mvSelectItems.Exists("batch_type") = True Then
                  vSQL = vSQL & ", batches b "
                  vWhere = vWhere & " AND fh.batch_number = b.batch_number "
                End If
              ElseIf mvSelectItems.Exists("order_number") Then
                vSQL = vSQL & "orders o, order_payment_history oph, financial_history fh"
                vWhere = "o.order_number = " & mvSelectItems("order_number").IntegerValue & " AND o.order_number = oph.order_number "
                vWhere = vWhere & "AND oph.batch_number = fh.batch_number AND oph.transaction_number = fh.transaction_number"
                If mvSelectItems.Exists("paying_in_slip_number") = True Or mvSelectItems.Exists("batch_type") = True Then
                  vSQL = vSQL & ", batches b "
                  vWhere = vWhere & " AND fh.batch_number = b.batch_number "
                End If
              Else
                If Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And mvSelectItems.Exists("sort_code") And mvSelectItems.Exists("account_number") Then
                  vSQL = vSQL & "contact_accounts ca, financial_history fh"
                  AddClause(vWhere, "ca.sort_code")
                  AddClause(vWhere, "account_number")
                  AddClause(vWhere, "account_name")
                  vWhere = vWhere & " AND ca.bank_details_number = fh.bank_details_number "
                  vAccountLinked = True
                  If mvSelectItems.Exists("paying_in_slip_number") = True Or mvSelectItems.Exists("batch_type") = True Then
                    vSQL = vSQL & ", batches b "
                    vWhere = vWhere & " AND fh.batch_number = b.batch_number "
                  End If
                ElseIf Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And mvSelectItems.Exists("card_number") = True Then
                  vSQL = vSQL & "card_sales cs, financial_history fh"
                  AddClause(vWhere, "card_number")
                  vWhere = vWhere & " AND cs.batch_number = fh.batch_number AND cs.transaction_number = fh.transaction_number"
                  vCardSalesLinked = True
                ElseIf Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And (mvSelectItems.Exists("iban_number") Or mvSelectItems.Exists("bic_code")) Then
                  vSQL = vSQL & "contact_accounts ca, financial_history fh"
                  If mvSelectItems.Exists("iban_number") Then AddClause(vWhere, "ca.iban_number")
                  If mvSelectItems.Exists("bic_code") Then AddClause(vWhere, "ca.bic_code")
                  AddClause(vWhere, "account_name")
                  vWhere = vWhere & " AND ca.bank_details_number = fh.bank_details_number "
                  vAccountLinked = True
                  If mvSelectItems.Exists("paying_in_slip_number") = True Or mvSelectItems.Exists("batch_type") = True Then
                    vSQL = vSQL & ", batches b "
                    vWhere = vWhere & " AND fh.batch_number = b.batch_number "
                  End If
                Else
                  If mvSelectItems.Exists("paying_in_slip_number") = True Or mvSelectItems.Exists("batch_type") = True Then
                    vSQL = vSQL & "batches b, financial_history fh"
                    vWhere = vWhere & " b.batch_number = fh.batch_number"
                  Else
                    vSQL = vSQL & "financial_history fh"
                  End If
                End If
              End If
              AddClause(vWhere, "fh.contact_number")
              AddClause(vWhere, "fh.batch_number")
              AddClause(vWhere, "fh.transaction_number")
              AddClause(vWhere, "fh.transaction_type")
              AddClause(vWhere, "fh.amount")
              AddClause(vWhere, "fh.payment_method")
              AddClause(vWhere, "reference")
              AddClause(vWhere, "b.paying_in_slip_number")
              AddClause(vWhere, "b.batch_type")
              CheckDateRange()
              If mvSelectItems.Exists("date") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "fh.transaction_date" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
              End If
              If mvSelectItems.Exists("date2") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "fh.transaction_date" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
              End If
              If mvSelectItems.Exists("product_number") Then
                vSQL = vSQL & ",batch_transaction_analysis bta"
                If vWhere.Length > 0 Then vWhere = vWhere & " AND"
                vWhere = vWhere & " bta.batch_number = fh.batch_number AND bta.transaction_number = fh.transaction_number"
                AddClause(vWhere, "product_number")
              End If

              If vAccountLinked = False AndAlso (mvSelectItems.Exists("account_number") OrElse mvSelectItems.Exists("account_name") OrElse mvSelectItems.Exists("sort_code") OrElse mvSelectItems.Exists("iban_number") OrElse mvSelectItems.Exists("bic_code")) Then
                vSQL &= ", contact_accounts ca"
                If vWhere.Length > 0 Then vWhere &= " AND "
                vWhere &= "fh.bank_details_number = ca.bank_details_number "
                If mvSelectItems.Exists("sort_code") Then AddClause(vWhere, "ca.sort_code")
                If mvSelectItems.Exists("account_number") Then AddClause(vWhere, "account_number")
                AddClause(vWhere, "account_name")
                If mvSelectItems.Exists("bic_code") Then AddClause(vWhere, "bic_code")
                If mvSelectItems.Exists("iban_number") Then AddClause(vWhere, "iban_number")
                vAccountLinked = True
              End If

              If vWhere.Length > 0 Then
                If vCardSalesLinked = False And mvSelectItems.Exists("card_number") = True Then
                  vSQL = vSQL & ",card_sales cs, contacts c WHERE " & vWhere
                  vSQL = vSQL & " AND cs.batch_number = fh.batch_number AND cs.transaction_number = fh.transaction_number"
                  AddClause(vSQL, "card_number")
                Else
                  vSQL = vSQL & ", contacts c WHERE " & vWhere
                End If
                vSQL = vSQL & " AND fh.contact_number = c.contact_number "
                vSQL = vSQL & "ORDER BY transaction_date DESC"
              Else
                vError = True
              End If
            Else
              'Unprocessed/Unprocessed Provisional Transactions/Cancelled Provisional
              If mvSelectItems.Exists("member_number") Then
                vSQL = vSQL & "members m, batch_transaction_analysis bta,batch_transactions bt "
                vWhere = "m.member_number = '" & mvSelectItems("member_number").Value & "' "
                vWhere = vWhere & "AND bta.member_number = m.member_number "
                vWhere = vWhere & "AND bt.batch_number = bta.batch_number AND bt.transaction_number = bta.transaction_number"
              ElseIf mvSelectItems.Exists("order_number") Then
                vSQL = vSQL & "orders o, batch_transaction_analysis bta,batch_transactions bt "
                vWhere = "o.order_number = " & mvSelectItems("order_number").Value & " "
                vWhere = vWhere & "AND bta.order_number = o.order_number "
                vWhere = vWhere & "AND bt.batch_number = bta.batch_number AND bt.transaction_number = bta.transaction_number"
              Else
                If Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And mvSelectItems.Exists("sort_code") And mvSelectItems.Exists("account_number") Then
                  vSQL = vSQL & "contact_accounts ca, batch_transactions bt"
                  AddClause(vWhere, "ca.sort_code")
                  AddClause(vWhere, "account_number")
                  AddClause(vWhere, "account_name")
                  vWhere = vWhere & " AND ca.bank_details_number = bt.bank_details_number "
                  vAccountLinked = True
                ElseIf Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And ((mvSelectItems.Exists("bic_code") And mvSelectItems.Exists("iban_number")) Or (mvSelectItems.Exists("bic_code") Or mvSelectItems.Exists("iban_number"))) Then
                  vSQL = vSQL & "contact_accounts ca, batch_transactions bt"
                  AddClause(vWhere, "iban_number")
                  AddClause(vWhere, "bic_code")
                  AddClause(vWhere, "account_name")
                  vWhere = vWhere & " AND ca.bank_details_number = bt.bank_details_number "
                  vAccountLinked = True
                Else
                  If Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And Not mvSelectItems.Exists("transaction_date") And Not mvSelectItems.Exists("reference") And (mvSelectItems.Exists("paying_in_slip_number") Or mvSelectItems.Exists("batch_type")) Then
                    vSQL = vSQL & "batches b, batch_transactions bt"
                    If mvTransactionFinderType = TransactionFinderTypes.tftProvisional Or mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                      vSQL = vSQL & ", confirmed_transactions ct"
                      vWhere = vWhere & " b.provisional = 'Y'"
                    End If
                    AddClause(vWhere, "b.batch_type")
                    AddClause(vWhere, "b.paying_in_slip_number")
                    vWhere = vWhere & " AND b.batch_number = bt.batch_number"
                    vJoinedToBatch = True
                    If mvTransactionFinderType = TransactionFinderTypes.tftProvisional Or mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                      vWhere = vWhere & " AND ct.provisional_batch_number = bt.batch_number AND ct.provisional_trans_number = bt.transaction_number AND ct.confirmed_batch_number IS NULL"
                      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataConfirmedTransStatus) Then
                        If mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                          vConfirmedTransaction.Init()
                          vWhere = vWhere & " AND ct.status = '" & vConfirmedTransaction.GetStatusCode(ConfirmedTransaction.ConfirmedTransactionStatus.Cancelled) & "'"
                        Else
                          vWhere = vWhere & " AND ct.status IS NULL "
                        End If
                      End If
                    End If
                  ElseIf Not mvSelectItems.Exists("contact_number") And Not mvSelectItems.Exists("batch_number") And mvSelectItems.Exists("card_number") = True Then
                    vSQL = vSQL & "card_sales cs, batch_transactions bt"
                    AddClause(vWhere, "card_number")
                    vWhere = vWhere & " AND bt.batch_number = cs.batch_number AND bt.transaction_number = cs.transaction_number"
                    vCardSalesLinked = True
                  Else
                    vSQL = vSQL & "batch_transactions bt"
                  End If
                End If
              End If

              AddClause(vWhere, "bt.contact_number")
              AddClause(vWhere, "bt.batch_number")
              AddClause(vWhere, "bt.transaction_number")
              AddClause(vWhere, "bt.transaction_type")
              AddClause(vWhere, "bt.amount")
              AddClause(vWhere, "bt.payment_method")
              AddClause(vWhere, "reference")
              CheckDateRange()
              If mvSelectItems.Exists("date") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "bt.transaction_date" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
              End If
              If mvSelectItems.Exists("date2") Then
                If vWhere.Length > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "bt.transaction_date" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
              End If
              If mvSelectItems.Exists("product_number") Then
                If Not (mvSelectItems.Exists("order_number")) And Not (mvSelectItems.Exists("member_number")) Then
                  vSQL = vSQL & ", batch_transaction_analysis bta"
                  If vWhere.Length > 0 Then vWhere = vWhere & " AND"
                  vWhere = vWhere & " bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number"
                End If
                AddClause(vWhere, "product_number")
              End If

              If vAccountLinked = False AndAlso (mvSelectItems.Exists("account_number") OrElse mvSelectItems.Exists("account_name") OrElse mvSelectItems.Exists("sort_code") OrElse mvSelectItems.Exists("bic_code") OrElse mvSelectItems.Exists("iban_number")) Then
                vSQL &= ", contact_accounts ca "
                If vWhere.Length > 0 Then vWhere &= " AND "
                vWhere &= " bt.bank_details_number = ca.bank_details_number "
                If mvSelectItems.Exists("sort_code") Then AddClause(vWhere, "ca.sort_code")
                If mvSelectItems.Exists("account_number") Then AddClause(vWhere, "account_number")
                AddClause(vWhere, "account_name")
                If mvSelectItems.Exists("bic_code") Then AddClause(vWhere, "bic_code")
                If mvSelectItems.Exists("iban_number") Then AddClause(vWhere, "iban_number")
                vAccountLinked = True
              End If

              If vWhere.Length > 0 Then
                If Not vJoinedToBatch Then
                  vSQL = vSQL & ", batches b "
                  vWhere = vWhere & " AND b.batch_number = bt.batch_number "
                  If mvTransactionFinderType = TransactionFinderTypes.tftProvisional Or mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                    vSQL = vSQL & ", confirmed_transactions ct"
                    vWhere = vWhere & " AND b.provisional = 'Y'"
                  End If
                  AddClause(vWhere, "b.batch_type")
                  AddClause(vWhere, "b.paying_in_slip_number")
                  If mvTransactionFinderType = TransactionFinderTypes.tftProvisional Or mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                    vWhere = vWhere & " AND ct.provisional_batch_number = bt.batch_number AND ct.provisional_trans_number = bt.transaction_number AND ct.confirmed_batch_number IS NULL"
                    If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataConfirmedTransStatus) Then
                      If mvTransactionFinderType = TransactionFinderTypes.tftCancelledProvisional Then
                        vConfirmedTransaction.Init()
                        vWhere = vWhere & " AND ct.status = '" & vConfirmedTransaction.GetStatusCode(ConfirmedTransaction.ConfirmedTransactionStatus.Cancelled) & "'"
                      Else
                        vWhere = vWhere & " AND ct.status IS NULL "
                      End If
                    End If
                  End If
                End If

                If mvTransactionFinderType = TransactionFinderTypes.tftUnprocessed Then vWhere = vWhere & " AND b.posted_to_nominal = 'N'"
                If vWhere.Length > 0 Then
                  If vCardSalesLinked = False And mvSelectItems.Exists("card_number") = True Then
                    vSQL = vSQL & ",card_sales cs, contacts c WHERE " & vWhere
                    vSQL = vSQL & " AND cs.batch_number = bt.batch_number AND cs.transaction_number = bt.transaction_number"
                    AddClause(vSQL, "card_number")
                  Else
                    vSQL = vSQL & ", contacts c WHERE " & vWhere
                  End If
                  vSQL = vSQL & " AND bt.contact_number = c.contact_number "
                  vSQL = vSQL & "ORDER BY transaction_date DESC"
                Else
                  vError = True
                End If
              Else
                vError = True
              End If
            End If

          Case DataFinderTypes.dftVenue
            If mvSelectItems.Count = 0 Then
              vError = True
            Else
              If mvSelectItems.Exists("activity") Or mvSelectItems.Exists("activity_value") Or mvSelectItems.Exists("quantity") Or mvSelectItems.Exists("quantity2") Then vActivities = True
              If vActivities Then
                If Not mvSelectItems.Exists("activity") Or ((mvSelectItems.Exists("quantity") Or mvSelectItems.Exists("quantity2")) And Not mvSelectItems.Exists("activity_value")) Then RaiseError(DataAccessErrors.daeActivityRequired)
              End If
              AddClause(vWhere, "venue")
              AddClause(vWhere, "venue_desc")
              vSQL = vSQL & "venues v LEFT OUTER JOIN organisations o ON v.organisation_number = o.organisation_number "
              vSQL = vSQL & "LEFT OUTER JOIN addresses a ON v.address_number = a.address_number"
              If vActivities Then vSQL = vSQL & " , organisation_categories oc"
              vSQL = vSQL & " WHERE " & vWhere
              If vWhere.Length > 0 AndAlso vActivities Then vSQL = vSQL & " AND "
              If vActivities Then vSQL = vSQL & "o.organisation_number = oc.organisation_number"
              AddClause(vSQL, "activity")
              AddClause(vSQL, "activity_value")
              If mvSelectItems.Exists("quantity") Then
                vSQL = vSQL & " AND quantity >= " & mvSelectItems("quantity").Value
              End If
              If mvSelectItems.Exists("quantity2") Then
                vSQL = vSQL & " AND quantity <= " & mvSelectItems("quantity2").Value
              End If
              AddClause(vSQL, "town")
              AddClause(vSQL, "postcode")
            End If

          Case DataFinderTypes.dftGiveAsYouEarn
            vSQL = Replace(vSQL, "agency_name", "ao.name AS agency_name")
            If mvSelectItems.Exists("gaye_pledge_number") Then
              'Select the gaye_pledge first
              vSQL = vSQL & "gaye_pledges t,contacts c,organisations o,gaye_agencies ga,organisations ao WHERE "
              vGAYEPledge = True
            ElseIf mvSelectItems.Exists("agency_number") Then
              'Select from gaye_agencies first
              vGayeAgencyFirst = True
              vSQL = vSQL & "gaye_agencies ga,organisations ao,gaye_pledges t,contacts c,organisations o WHERE "
              With mvSelectItems.Item("agency_number")
                vWhere = " ga.organisation_number " & mvEnv.Connection.SQLLiteral("=", .FieldType, .Value)
              End With
              vWhere = vWhere & " AND ao.organisation_number = ga.organisation_number AND t.agency_number = ga.organisation_number"
              AddGAYEPledgeClause(vWhere, True)
              AddClause(vWhere, "t.organisation_number")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
              AddClause(vWhere, "o.name")
            ElseIf mvSelectItems.Exists("organisation_number") Then
              'Select from organisations (employer) first
              vSQL = vSQL & "organisations o,gaye_pledges t,contacts c,gaye_agencies ga,organisations ao WHERE "
              AddClause(vWhere, "o.organisation_number")
              AddClause(vWhere, "o.name")
              vWhere = vWhere & " AND t.organisation_number = o.organisation_number"
              AddGAYEPledgeClause(vWhere, True)
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
            ElseIf mvSelectItems.Exists("batch_number") Or mvSelectItems.Exists("transaction_number") Then
              'Select from payment history first
              vSQL = vSQL & "gaye_pledge_payment_history gph,gaye_pledges t,contacts c,organisations o,gaye_agencies ga,organisations ao WHERE "
              AddClause(vWhere, "batch_number")
              AddClause(vWhere, "transaction_number")
              vWhere = vWhere & " AND t.gaye_pledge_number = gph.gaye_pledge_number"
              vGAYEPledge = True
              vGayeBatch = True
            ElseIf mvSelectItems.Exists("surname") Or mvSelectItems.Exists("initials") Then
              'Select from contacts first
              vSQL = vSQL & "contacts c,gaye_pledges t,organisations o,gaye_agencies ga,organisations ao WHERE "
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND t.contact_number = c.contact_number"
              AddGAYEPledgeClause(vWhere, True)
              AddClause(vWhere, "t.organisation_number")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
              AddClause(vWhere, "o.name")
              vWhere = vWhere & " AND ga.organisation_number = t.agency_number AND ao.organisation_number = ga.organisation_number"
            ElseIf mvSelectItems.Exists("name") Then
              'Select from organisations (employer) first
              vSQL = vSQL & "organisations o,gaye_pledges t,contacts c,gaye_agencies ga,organisations ao WHERE "
              AddClause(vWhere, "o.name")
              vWhere = vWhere & " AND t.organisation_number = o.organisation_number"
              AddGAYEPledgeClause(vWhere, True)
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
            Else
              'Select from gaye_pledges first
              vSQL = vSQL & "gaye_pledges t,contacts c,organisations o,gaye_agencies ga,organisations ao WHERE "
              vGAYEPledge = True
            End If

            If vGAYEPledge Then
              AddGAYEPledgeClause(vWhere, True)
              AddClause(vWhere, "t.organisation_number")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
              AddClause(vWhere, "o.name")
            End If

            If Not vGayeAgencyFirst Then
              vWhere = vWhere & " AND ga.organisation_number = t.agency_number AND ao.organisation_number = ga.organisation_number"
              If mvSelectItems.Exists("agency_number") Then
                With mvSelectItems.Item("agency_number")
                  vWhere = " ga.organisation_number " & mvEnv.Connection.SQLLiteral("=", .FieldType, .Value)
                End With
              End If
            End If

            If (mvSelectItems.Exists("batch_number") Or mvSelectItems.Exists("transaction_number")) And Not (vGayeBatch) Then
              'Also need to link to gaye_pledge_payment_history
              vSQL = Left(vSQL, Len(vSQL) - 7)
              vSQL = vSQL & ",gaye_pledge_payment_history gph WHERE "
              vWhere = vWhere & " AND gph.gaye_pledge_number = t.gaye_pledge_number"
              AddClause(vWhere, "batch_number")
              AddClause(vWhere, "transaction_number")
            End If
            vSQL = vSQL & vWhere & " ORDER BY surname, initials, t.gaye_pledge_number"

          Case DataFinderTypes.dftInternalResource
            AddClause(vWhere, "resource_number")
            AddClause(vWhere, "t.product")
            AddClause(vWhere, "t.rate")
            AddClause(vWhere, "surname")
            vSQL = vSQL & " internal_resources t LEFT OUTER JOIN contacts c ON t.resource_contact_number = c.contact_number"
            vSQL = vSQL & " LEFT OUTER JOIN products p ON t.product = p.product"
            vSQL = vSQL & " LEFT OUTER JOIN rates r ON t.product = r.product AND t.rate = r.rate"
            vSQL = vSQL & " WHERE " & vWhere & " ORDER BY t.resource_number"
            vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)

          Case DataFinderTypes.dftContactMailingDocuments
            AddClause(vWhere, "cmd.mailing_template")
            AddClause(vWhere, "mailing_document_number")
            AddClause(vWhere, "created_by")
            AddClause(vWhere, "cmd.mailing")

            'Check Created on dates
            CheckDateRange()
            If mvSelectItems.Exists("date") Then
              If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "created_on" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
            End If
            If mvSelectItems.Exists("date2") Then
              If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & "created_on" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
            End If
            If mvSelectItems.Exists("none2") Then
              If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & " fulfillment_number is NULL "
              vTableList = "contact_mailing_documents cmd"
            ElseIf mvSelectItems.Exists("none3") Then
              If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
              vWhere = vWhere & " cmd.fulfillment_number is NOT NULL"
              AddClause(vWhere, "cmd.fulfillment_number")
              vWhere = vWhere & " AND cmd.fulfillment_number = fh.fulfillment_number "
              AddClause(vWhere, "fulfilled_by")
              'Check Fulfilled on dates
              If mvSelectItems.Exists("date3") And mvSelectItems.Exists("date4") Then
                If CDate(mvSelectItems("date4").Value) < CDate(mvSelectItems("date3").Value) Then RaiseError(DataAccessErrors.daeInvalidDateRange)
              End If
              If mvSelectItems.Exists("date3") Then
                If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
                vWhere = vWhere & "fulfilled_on" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (mvSelectItems("date3").Value))
              End If
              If mvSelectItems.Exists("date4") Then
                If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
                'Since fulfillment_history.fulfilled_on will contain time data we have to do some trickery and add one day to the end of the date range.
                'We then retrieve those records where fulfilled_on < that calculated date.
                'We do this rather than finding those where fulfilled_on <= the entered end date because that doesn't work due to the attribute containing time data.
                vWhere = vWhere & "fulfilled_on" & mvConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, CDate(mvSelectItems("date4").Value).AddDays(1).ToString(CAREDateFormat))
              End If
              vTableList = "contact_mailing_documents cmd, fulfillment_history fh"
            End If
            vTableList = vTableList & ", mailing_templates mt, contacts c, mailings m"
            If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
            vWhere = vWhere & "cmd.mailing_template = mt.mailing_template AND cmd.contact_number = c.contact_number AND cmd. mailing = m.mailing"
            vSQL = vSQL & vTableList & " WHERE " & vWhere & " ORDER BY mailing_document_number"
          Case DataFinderTypes.dftCommunication
            If mvSelectItems.Exists("contact_number") Then vIndex = vIndex + 1
            If mvSelectItems.Exists("address_number") Then vIndex = vIndex + 1
            If mvSelectItems.Exists("number") Then vIndex = vIndex + 1
            If mvSelectItems.Exists("communication_number") Then vIndex = vIndex + 1
            AddClause(vWhere, "com.contact_number")
            AddClause(vWhere, "com.address_number")
            AddClause(vWhere, "com.device")
            AddClause(vWhere, "com.dialling_code")
            AddClause(vWhere, "com.std_code")
            AddClause(vWhere, "number")
            AddClause(vWhere, "extension")
            AddClause(vWhere, "com.ex_directory")
            AddClause(vWhere, "communication_number")
            If mvSelectItems.Exists("name") Then
              If vIndex > 0 Then
                vSQL = vSQL & "communications com,organisation_addresses oa,organisations o,contacts c, devices d WHERE " & vWhere
                vSQL = vSQL & " AND com.address_number = oa.address_number"
                vSQL = vSQL & " AND oa.organisation_number = o.organisation_number"
                AddClause(vSQL, "name")
                vSQL = vSQL & " AND o.organisation_number = c.contact_number"
              Else
                vSQL = vSQL & "organisations o,contacts c,organisation_addresses oa,communications com, devices d WHERE "
                AddClause(vSQL, "name")
                vSQL = vSQL & " AND o.organisation_number = c.contact_number"
                vSQL = vSQL & " AND o.organisation_number = oa.organisation_number"
                vSQL = vSQL & " AND oa.address_number = com.address_number"
                If Len(vWhere) > 0 Then vSQL = vSQL & " AND " & vWhere
              End If
            ElseIf mvSelectItems.Exists("surname") Then
              If vIndex > 0 Then
                vSQL = vSQL & "communications com,contacts c, devices d WHERE " & vWhere
                vSQL = vSQL & " AND com.contact_number = c.contact_number"
                AddClause(vSQL, "surname")
                AddClause(vSQL, "initials")
              Else
                vSQL = vSQL & "contacts c,communications com, devices d WHERE "
                AddClause(vSQL, "surname")
                AddClause(vSQL, "initials")
                vSQL = vSQL & " AND c.contact_number = com.contact_number"
                If Len(vWhere) > 0 Then vSQL = vSQL & " AND " & vWhere
              End If
            Else
              vSQL = vSQL & "communications com,contacts c, devices d WHERE " & vWhere
              vSQL = vSQL & " AND com.contact_number = c.contact_number"
              AddClause(vSQL, "initials")
            End If
            vSQL = vSQL & " AND com.device = d.device"
            vSQL = vSQL & " ORDER BY c.surname, c.initials, com.device, communication_number"
            vSQL = Replace(vSQL, ",number,", "," & mvConn.DBSpecialCol("com", "number") & ",")
            vSQL = Replace(vSQL, " number ", " " & mvConn.DBSpecialCol("com", "number") & " ")

          Case DataFinderTypes.dftPostTaxPayrollGiving
            If mvSelectItems.Exists("pledge_number") Then
              'Select Pledge data first
              vSQL = vSQL & "post_tax_pg_pledges t,contacts c,organisations o WHERE "
              vGAYEPledge = True
            ElseIf mvSelectItems.Exists("organisation_number") Then
              If mvSelectItems.Exists("payroll_number") Then
                'Select Pledge data first (using Organisation Number & Payroll Number)
                vSQL = vSQL & "post_tax_pg_pledges t,contacts c,organisations o WHERE "
                AddClause(vWhere, "t.organisation_number")
                AddGAYEPledgeClause(vWhere, False)
                AddCancellationClause(vWhere)
                vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
                AddClause(vWhere, "name")
                vWhere = vWhere & " AND c.contact_number = t.contact_number"
                AddClause(vWhere, "surname")
                AddClause(vWhere, "initials")
              Else
                'Select Organisation (employer) data first (using Organisation Number)
                vSQL = vSQL & "organisations o,post_tax_pg_pledges t,contacts c WHERE "
                AddClause(vWhere, "o.organisation_number")
                AddClause(vWhere, "o.name")
                vWhere = vWhere & " AND t.organisation_number = o.organisation_number"
                AddGAYEPledgeClause(vWhere, False)
                AddCancellationClause(vWhere)
                vWhere = vWhere & " AND c.contact_number = t.contact_number"
                AddClause(vWhere, "surname")
                AddClause(vWhere, "initials")
              End If
            ElseIf mvSelectItems.Exists("name") Then
              'Select Organisation (Employer) data first (using Name)
              vSQL = vSQL & "organisations o,post_tax_pg_pledges t,contacts c WHERE "
              AddClause(vWhere, "o.organisation_number")
              AddClause(vWhere, "o.name")
              vWhere = vWhere & " AND t.organisation_number = o.organisation_number"
              AddGAYEPledgeClause(vWhere, False)
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
            ElseIf mvSelectItems.Exists("surname") Or mvSelectItems.Exists("initials") Then
              'Select Contact data first
              vSQL = vSQL & "contacts c,post_tax_pg_pledges t,organisations o WHERE "
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND t.contact_number = c.contact_number"
              AddGAYEPledgeClause(vWhere, False)
              'AddClause vWhere, "t.organisation_number"
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
              AddClause(vWhere, "name")
            ElseIf mvSelectItems.Exists("batch_number") Or mvSelectItems.Exists("transaction_number") Then
              'Select payment data first
              vSQL = vSQL & "post_tax_pg_payment_history pgph,post_tax_pg_pledges t,contacts c,organisations o WHERE "
              AddClause(vWhere, "batch_number")
              AddClause(vWhere, "transaction_number")
              vWhere = vWhere & " AND t.pledge_number = pgph.pledge_number"
              vGayeBatch = True
              vGAYEPledge = True
            Else
              'Select Pledge data first
              vSQL = vSQL & "post_tax_pg_pledges t,contacts c,organisations o WHERE "
              vGAYEPledge = True
            End If

            If vGAYEPledge Then
              AddGAYEPledgeClause(vWhere, False)
              AddClause(vWhere, "t.organisation_number")
              AddCancellationClause(vWhere)
              vWhere = vWhere & " AND c.contact_number = t.contact_number"
              AddClause(vWhere, "surname")
              AddClause(vWhere, "initials")
              vWhere = vWhere & " AND o.organisation_number = t.organisation_number"
              AddClause(vWhere, "name")
            End If

            If (mvSelectItems.Exists("batch_number") Or mvSelectItems.Exists("transaction_number")) And vGayeBatch = False Then
              vSQL = Replace(vSQL, " WHERE", ",post_tax_pg_payment_history pgph WHERE")
              vWhere = vWhere & " AND pgph.pledge_number = t.pledge_number"
              AddClause(vWhere, "batch_number")
              AddClause(vWhere, "transaction_number")
            End If
            vSQL = vSQL & vWhere & " ORDER BY surname, initials, t.pledge_number"

          Case DataFinderTypes.dftAppealCollections
            If mvSelectItems.Exists("pis_number") = True And mvSelectItems.Count = 1 Then
              'Select collection_pis first
              vPayTableFirst = True
              vSQL = vSQL & "collection_pis cp, appeal_collections t WHERE "
              AddClause(vWhere, "pis_number")
              vWhere = vWhere & " AND t.collection_number = cp.collection_number"
            Else
              vSQL = vSQL & "appeal_collections t WHERE "
            End If
            AddClause(vWhere, "t.collection_number")
            AddClause(vWhere, "campaign")
            AddClause(vWhere, "appeal")
            AddClause(vWhere, "collection")
            AddClause(vWhere, "t.bank_account")
            If vPayTableFirst = False And mvSelectItems.Exists("pis_number") = True Then
              'Join to collection_pis table
              vSQL = Replace(vSQL, "WHERE", ", collection_pis cp WHERE")
              vWhere = vWhere & " AND cp.collection_number = t.collection_number"
              AddClause(vWhere, "pis_number")
            End If
            vSQL = vSQL & vWhere & " ORDER BY campaign, appeal, collection"

          Case DataFinderTypes.dftStandardDocuments
            AddClause(vWhere, "standard_document")
            AddClause(vWhere, "standard_document_desc")
            AddClause(vWhere, "sd.document_type")
            AddClause(vWhere, "sd.topic")
            AddClause(vWhere, "sd.sub_topic")
            AddClause(vWhere, "sd.history_only")
            AddClause(vWhere, "sd.mailmerge_header")
            AddClause(vWhere, "instant_print")
            vSQL = vSQL & " standard_documents sd INNER JOIN document_types dt ON sd.document_type = dt.document_type"
            vSQL = vSQL & " INNER JOIN topics t ON sd.topic = t.topic INNER JOIN sub_topics st ON sd.topic = st.topic AND sd.sub_topic = st.sub_topic"
            vSQL = vSQL & " LEFT OUTER JOIN (SELECT p.package, docfile_extension,document_source, communication_type FROM packages p WHERE communication_type IN (1,3) OR document_source = 'E') p ON sd.package = p.package"
            vSQL = vSQL & " WHERE " & If(String.IsNullOrWhiteSpace(vWhere), String.Empty, vWhere & " AND ") & "(sd.package IS NULL OR (sd.package IS NOT NULL AND (docfile_extension IS NOT NULL OR p.document_source = 'E' OR p.communication_type = '3'))) ORDER BY standard_document_desc"
          Case DataFinderTypes.dftExternalReference
            AddClause(vWhere, "external_reference")
            AddClause(vWhere, "data_source")
            vSQL = vSQL & " contact_external_links"
            vSQL = vSQL & " WHERE " & vWhere
        End Select
        If vError Then RaiseError(DataAccessErrors.daeNoSelectionData)
        If pCountOnly Then
          vPos = InStr(vSQL, "ORDER BY")
          If vPos > 0 Then vSQL = Left(vSQL, vPos - 1)
        End If
        Return vSQL
      End Get
    End Property

    Private Sub AddCancellationClause(ByRef pWhere As String)
      Dim vCancelled As Boolean

      AddClause(pWhere, "t.cancellation_reason")
      If mvSelectItems.Exists("date2") Then
        'First check that "From" is before "To" date
        If mvSelectItems.Exists("date") Then
          If CDate(mvSelectItems("date").Value) > CDate(mvSelectItems("date2").Value) Then
            RaiseError(DataAccessErrors.daeInvalidDateRange)
          End If
        End If

        If pWhere.Length > 0 Then pWhere = pWhere & " AND "
        If mvSelectItems.Exists("date") Then
          pWhere = pWhere & "t.cancelled_on" & mvConn.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value)) & mvConn.SQLLiteral("AND", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
        Else
          pWhere = pWhere & "t.cancelled_on" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date2").Value))
        End If
        vCancelled = True
      Else
        If mvSelectItems.Exists("date") Then
          If pWhere.Length > 0 Then pWhere = pWhere & " AND "
          pWhere = pWhere & "t.cancelled_on" & mvConn.SQLLiteral("=", CDBField.FieldTypes.cftDate, (mvSelectItems("date").Value))
          vCancelled = True
        End If
      End If

      'BR14231: Only Get non cancelled records when current is set
      If mvSelectItems.Exists("current") AndAlso mvSelectItems("current").Bool Then
        If Not vCancelled AndAlso Not mvSelectItems.Exists("cancellation_reason") Then
          'cancellation_reason, date and date2 have not been applied
          If pWhere.Length > 0 Then pWhere &= " AND "
          pWhere &= " t.cancellation_reason IS NULL"
        Else
          'We have added the restriction for cancelled records. Un-set current so that any other restriction for CURRENT records would not be applied
          mvSelectItems("current").Bool = False
        End If
      End If
    End Sub
    ''' <summary>
    ''' Add bank details to the where clause
    ''' </summary>
    ''' <param name="pWhere">Where clause string by ref</param>
    ''' <remarks></remarks>
    Private Sub AddBankDetailsClause(ByRef pWhere As String)
      If mvSelectItems.Exists("sort_code") Then AddClause(pWhere, "sort_code")
      If mvSelectItems.Exists("account_number") Then AddClause(pWhere, "account_number")
      If mvSelectItems.Exists("iban_number") Then AddClause(pWhere, "iban_number")
      If mvSelectItems.Exists("bic_code") Then AddClause(pWhere, "bic_code")
    End Sub

    Private Sub AddGAYEPledgeClause(ByRef pWhere As String, ByVal pPreTaxPG As Boolean)
      If pPreTaxPG Then
        'Pre Tax Payroll Giving
        AddClause(pWhere, "t.gaye_pledge_number")
        AddClause(pWhere, "donor_id")
      Else
        'Post Tax Payroll Giving
        AddClause(pWhere, "t.pledge_number")
        AddClause(pWhere, "payroll_number")
      End If
      'Both
      AddClause(pWhere, "start_date")
      AddClause(pWhere, "t.source")
    End Sub

    Private Sub AddPaymentsDueClause(ByRef pWhere As String)
      If mvSelectItems.Exists("date4") Then
        'First check that "From" is before "To" date
        If mvSelectItems.Exists("date3") Then
          If CDate(mvSelectItems("date3").Value) > CDate(mvSelectItems("date4").Value) Then
            RaiseError(DataAccessErrors.daeInvalidDateRange)
          End If
        End If

        If pWhere.Length > 0 And Trim(Right(pWhere, 4)) <> "AND" Then
          pWhere = pWhere & " AND "
        End If
        If mvSelectItems.Exists("date3") Then
          pWhere = pWhere & "COALESCE(ops.next_payment_due, o.next_payment_due)" & mvConn.SQLLiteral("BETWEEN ", CDBField.FieldTypes.cftDate, (mvSelectItems("date3").Value)) & mvConn.SQLLiteral("AND ", CDBField.FieldTypes.cftDate, (mvSelectItems("date4").Value))
        Else
          pWhere = pWhere & "COALESCE(ops.next_payment_due, o.next_payment_due)" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (mvSelectItems("date4").Value))
        End If
      Else
        If mvSelectItems.Exists("date3") Then
          If pWhere.Length > 0 Then pWhere = pWhere & " AND "
          pWhere = pWhere & "COALESCE(ops.next_payment_due, o.next_payment_due)" & mvConn.SQLLiteral("=", CDBField.FieldTypes.cftDate, (mvSelectItems("date3").Value))
        End If
      End If
    End Sub

    Public Sub CheckDateRange()
      If mvSelectItems.Exists("date") And mvSelectItems.Exists("date2") Then
        If CDate(mvSelectItems("date2").Value) < CDate(mvSelectItems("date").Value) Then RaiseError(DataAccessErrors.daeInvalidDateRange)
      End If
    End Sub

    Private Sub CheckRequiredSORecFields()
      'Check that BankAccount has been selected and is valid, and some other field also selected
      Dim vFields As New CDBFields

      If mvSelectItems.Exists("bank_account") Then
        vFields.Add("bank_account", CDBField.FieldTypes.cftCharacter, mvSelectItems("bank_account").Value)
        vFields.Add("batch_type", CDBField.FieldTypes.cftCharacter, "SO")
        If mvConn.GetCount("bank_suspense_accounts", vFields) = 0 Then
          RaiseError(DataAccessErrors.daeNoSuspenseAccount)
        Else
          If mvSelectItems.Count = 1 Then RaiseError(DataAccessErrors.daeNoOtherSelectionData)
        End If
      Else
        RaiseError(DataAccessErrors.daeNoBankAccount)
      End If
    End Sub

    Private Function SelectItemValue(ByRef pItem As String) As String
      If mvSelectItems.Exists(pItem) Then
        SelectItemValue = mvSelectItems(pItem).Value
      Else
        SelectItemValue = ""
      End If
    End Function

    Friend Sub GetActionsData(ByVal pDataTable As CDBDataTable)
      GetActionsForFinder(pDataTable)
      GetCustomMergeData(pDataTable)
    End Sub

    Public ReadOnly Property DataTable() As CDBDataTable
      Get
        Dim vDataTable As New CDBDataTable
        Dim vHideFields As Boolean
        Dim vItems() As String
        Dim vIndex As Integer
        Dim vContactNumbers As String = ""
        Dim vErrorNumber As Integer
        Dim vTimeout As Integer

        If mvSelectItems.Exists("Timeout") Then
          vTimeout = mvSelectItems("Timeout").IntegerValue
          If vTimeout > 0 Then vDataTable.Timeout = vTimeout
        End If
        If mvType = DataFinderTypes.dftTransaction Then SetTransactionFinderType()
        Select Case mvType
          Case DataFinderTypes.dftActions
            vDataTable.AddColumnsFromList(mvResultColumns)
            GetActionsForFinder(vDataTable)
            GetCustomMergeData(vDataTable)

          Case DataFinderTypes.dftUniservPhoneBook
            vDataTable.AddColumnsFromList(mvResultColumns)
            vErrorNumber = mvEnv.UniservInterface.FindInPhoneBook(SelectItemValue("preferred_forename"), SelectItemValue("surname"), "", SelectItemValue("building_number"), SelectItemValue("address"), SelectItemValue("town"), SelectItemValue("postcode"), "", SelectItemValue("date_of_birth"), vContactNumbers)
            If vErrorNumber = 0 Then
              If vContactNumbers.Length > 0 Then
                Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "gefeco_nr,surname,forenames,postcode,town,address1,address2,std_code,telephone", "phone_book", New CDBField("gefeco_nr", vContactNumbers, CDBField.FieldWhereOperators.fwoIn))
                vDataTable.FillFromSQL(mvEnv, vSQLStatement, "gefeco_nr,surname,forenames,postcode,town,address1,PHONEBOOK_PHONE")
              End If
            End If

          Case DataFinderTypes.dftMember, DataFinderTypes.dftCCCA, DataFinderTypes.dftCovenant, DataFinderTypes.dftDirectDebit, DataFinderTypes.dftGAD, DataFinderTypes.dftInvoice, DataFinderTypes.dftLegacy, DataFinderTypes.dftStandingOrder, DataFinderTypes.dftGiveAsYouEarn, DataFinderTypes.dftPurchaseOrder, DataFinderTypes.dftTransaction, DataFinderTypes.dftPostTaxPayrollGiving, DataFinderTypes.dftPaymentPlan
            vDataTable.AddColumnsFromList(mvResultColumns)
            vDataTable.FillFromSQLDONOTUSE(mvEnv, SelectionSQL)
            vHideFields = True

          Case DataFinderTypes.dftProduct, DataFinderTypes.dftStandardDocuments, DataFinderTypes.dftEventPersonnel
            vDataTable.AddColumnsFromList(mvResultColumns)
            vDataTable.FillFromSQLDONOTUSE(mvEnv, SelectionSQL)

          Case DataFinderTypes.dftInternalResource
            vDataTable.AddColumnsFromList(mvResultColumns)
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, SelectionSQL)
            vDataTable.FillFromSQL(mvEnv, vSQLStatement, mvSelectAttrs)

          Case DataFinderTypes.dftServiceProduct
            If CustomColumns.Length > 0 Then Init(mvEnv, mvType) 'We've used the finder once already and custom data was added to the result set, so call the Init method again to reset the module-level variables mvCustomColumns, mvCustomHeadings, mvSelectColumns, mvSelectHeadings & mvSelectWidths
            vDataTable.AddColumnsFromList(mvResultColumns)
            Dim vSQLStatement As New SQLStatement(mvEnv.Connection, SelectionSQL)
            vDataTable.FillFromSQL(mvEnv, vSQLStatement)
            GetCustomMergeData(vDataTable)
            If CustomColumns.Length > 0 Then
              mvSelectColumns = mvSelectColumns & "," & CustomColumns
              mvSelectHeadings = mvSelectHeadings & "," & CustomHeadings
              vItems = Split(CustomColumns, ",")
              For vIndex = 0 To UBound(vItems)
                mvSelectWidths = mvSelectWidths & ",1000"
              Next
            End If
        End Select
        If vHideFields Then vDataTable.SuppressData()
        Return vDataTable
      End Get
    End Property

    Private Sub GetCustomMergeData(ByVal pDT As CDBDataTable)
      'Copied from DataSelectionFinder and slightly modified
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCustomFinderTab) AndAlso mvEnv.GetConfigOption("option_custom_data") Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("client", mvEnv.ClientCode)
        Dim vCustomField As String = ""
        Dim vAttrName As String = ""
        Dim vGroupParam As String = ""
        Select Case mvType
          Case DataFinderTypes.dftActions
            vWhereFields.Add("usage_code", "A")
            vGroupParam = "contact_group"
            vAttrName = "ActionNumber"
            vCustomField = "action_number"
          Case DataFinderTypes.dftServiceProduct
            vWhereFields.Add("usage_code", "S")
            vGroupParam = "contact_group"
            vAttrName = "ContactNumber"
            vCustomField = "contact_number"
        End Select
        If mvSelectItems.Exists(vGroupParam) Then
          vWhereFields.Add("contact_group", mvSelectItems(vGroupParam).Value)
        End If
        Dim vCMD As New CustomMergeData(mvEnv)
        Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, vCMD.GetRecordSetFields, "custom_merge_data cmd", vWhereFields, "sequence_number").GetRecordSet
        Dim vCMDS As New CollectionList(Of CustomMergeData)
        While vRecordSet.Fetch
          vCMD.InitFromRecordSet(vRecordSet)
          vCMDS.Add(vCMDS.Count.ToString, vCMD)
          vCMD = New CustomMergeData(mvEnv)
        End While
        vRecordSet.CloseRecordSet()

        Dim vCount As Integer
        Dim vItemNo As String
        Dim vCheckRow As CDBDataRow
        Dim vIncludeNo As String
        Dim vCustomColumns As New StringBuilder
        Dim vCustomHeadings As New StringBuilder

        For Each vCMD In vCMDS
          pDT.AddColumnsFromList(vCMD.StandardAttributeNames)
          If mvCustomColumns.Length > 0 Then vCustomColumns.Append(",")
          vCustomColumns.Append(vCMD.StandardAttributeNames)
          If vCustomHeadings.Length > 0 Then vCustomHeadings.Append(",")
          vCustomHeadings.Append(vCMD.AttributeCaptions)
          Dim vInclude As New StringBuilder
          For Each vRow As CDBDataRow In pDT.Rows
            vIncludeNo = vRow.Item(vAttrName)
            If vIncludeNo.Length > 0 Then
              If vInclude.Length > 0 Then vInclude.Append(",")
              vInclude.Append(vIncludeNo)
              vCount += 1
              If vCount >= 250 OrElse vRow Is pDT.Rows(pDT.Rows.Count - 1) Then
                If vInclude.Length > 0 Then
                  'now retrieve custom information
                  Dim vRecordSet2 As CDBRecordSet = vCMD.GetRecordSet(vInclude.ToString)
                  While vRecordSet2.Fetch()
                    vItemNo = vRecordSet2.Fields(vCustomField).Value
                    For Each vCheckRow In pDT.Rows
                      If vCheckRow.Item(vAttrName) = vItemNo Then vCMD.SetDataRow(vCheckRow, vRecordSet2, pDT.Columns)
                    Next
                  End While
                  vRecordSet2.CloseRecordSet()
                End If
                vInclude = New StringBuilder
                vCount = 0
              End If
            End If
          Next
        Next
        If vCustomColumns.Length > 0 Then
          mvCustomColumns = vCustomColumns.ToString
          mvCustomHeadings = vCustomHeadings.ToString
        End If
      End If
    End Sub

    Public Sub SetSelectItems(ByVal pParams As CDBParameters)
      Dim vParam As CDBParameter

      For Each vParam In pParams
        AddSelectItem(vParam.Name, vParam.Value, vParam.DataType)
      Next
    End Sub

    Public WriteOnly Property AddActionRelatedContactData() As Boolean
      Set(ByVal value As Boolean)
        mvAddActionRelatedContactData = value
      End Set
    End Property


    Private Sub GetActionsForFinder(ByRef pDataTable As CDBDataTable)
      Dim vSQL As String
      Dim vWhere As String
      Dim vAttrs As String
      Dim vItems As String
      Dim vDR As CDBDataRow
      Dim vRS As CDBRecordSet
      Dim vContact As Contact
      Dim vParams As CDBParameters
      Dim vUseTopic As Boolean
      Dim vActionIn As String

      'Add conditions to the WHERE clause
      vWhere = ActionLinkSQL()

      If mvSelectItems.Exists("action_status") Then
        mvSelectItems("action_status").WhereOperator = CDBField.FieldWhereOperators.fwoIn
        AddClause(vWhere, "a.action_status")
      Else
        If mvSelectItems.Exists("active_status") Then
          If mvSelectItems("active_status").Bool Then AddClauseValue(vWhere, "a.action_status", Action.GetActionStatusCode(Action.ActionStatuses.astDefined) & "," & Action.GetActionStatusCode(Action.ActionStatuses.astScheduled) & "," & Action.GetActionStatusCode(Action.ActionStatuses.astOverdue), , CDBField.FieldWhereOperators.fwoIn)
        End If
      End If

      If mvSelectItems.Exists("action_number") Then
        If InStr(mvSelectItems("action_number").Value, ",") > 0 Then mvSelectItems("action_number").WhereOperator = CDBField.FieldWhereOperators.fwoIn
        AddClause(vWhere, "a.action_number")
      End If
      AddClause(vWhere, "master_action")
      AddClause(vWhere, "actioner_setting")
      AddClause(vWhere, "manager_setting")
      AddClause(vWhere, "a.action_priority")
      AddClause(vWhere, "action_desc")
      AddClause(vWhere, "action_text")

      If mvSelectItems.Exists("my_created_actions") Then
        If mvSelectItems("my_created_actions").Bool Then AddClauseValue(vWhere, "created_by", mvEnv.User.Logname)
      End If

      If mvSelectItems.Exists("scheduled_on_from") Then 'to do
        If mvSelectItems.Exists("scheduled_on_to") Then
          If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
          vWhere = vWhere & "scheduled_on BETWEEN " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (mvSelectItems("scheduled_on_from").Value))
          vWhere = vWhere & " AND " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (mvSelectItems("scheduled_on_to").Value))
        Else
          AddClauseValue(vWhere, "scheduled_on", (mvSelectItems("scheduled_on_from").Value), CDBField.FieldTypes.cftDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
      ElseIf mvSelectItems.Exists("scheduled_on_to") Then
        AddClauseValue(vWhere, "scheduled_on", (mvSelectItems("scheduled_on_to").Value), CDBField.FieldTypes.cftDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If

      If mvSelectItems.Exists("deadline_from") Then 'to do
        If mvSelectItems.Exists("deadline_to") Then
          If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
          vWhere = vWhere & "deadline BETWEEN " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (mvSelectItems("deadline_from").Value))
          vWhere = vWhere & " AND " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (mvSelectItems("deadline_to").Value))
        Else
          AddClauseValue(vWhere, "deadline", (mvSelectItems("deadline_from").Value), CDBField.FieldTypes.cftDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
      ElseIf mvSelectItems.Exists("deadline_to") Then
        AddClauseValue(vWhere, "deadline", (mvSelectItems("deadline_to").Value), CDBField.FieldTypes.cftDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
      End If

      If mvSelectItems.Exists("topic") Then
        vUseTopic = True
        AddClause(vWhere, "topic")
        AddClause(vWhere, "sub_topic")
      End If
      If mvSelectItems.Exists("document_class") Then 'BR20273 - Search on document_class if parameter passed
        AddClause(vWhere, "a.document_class")
      End If

      'Construct the SELECT statement and fill the supplied DataTable
      vAttrs = "master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status"
      vItems = vAttrs
      vAttrs = vAttrs & ",creator_header,department_header,public_header,department,creator_header AS access_level,actioner_setting,manager_setting"
      vItems = vItems & ",,,,,,,creator_header,department_header,public_header,department,access_level,actioner_setting,manager_setting,document_class"
      vSQL = "SELECT DISTINCT " & vAttrs & " FROM actions a,users u,document_classes dc,action_priorities ap,action_statuses acs "
      If vUseTopic Then vSQL = vSQL & ",action_subjects asu "
      vSQL = vSQL & "WHERE "
      If Len(vWhere) > 0 Then vSQL = vSQL & vWhere & " AND "
      vSQL = vSQL & "a.document_class = dc.document_class AND "
      vSQL = vSQL & "a.created_by = u.logname AND ((a.created_by = '" & mvEnv.User.Logname & "' AND creator_header = 'Y') OR (a.created_by <> '" & mvEnv.User.Logname & "' AND department = '" & mvEnv.User.Department & "' AND department_header = 'Y') OR (a.created_by <> '" & mvEnv.User.Logname & "' AND department <> '" & mvEnv.User.Department & "' AND public_header = 'Y' )) AND a.action_priority = ap.action_priority AND a.action_status = acs.action_status "
      If vUseTopic Then vSQL = vSQL & "AND asu.action_number = a.action_number "
      vSQL = vSQL & "ORDER BY master_action, action_level, sequence_number"
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL, vItems)
      pDataTable.SetDocumentAccess()

      'If the associated Display List has any of the Related Contact columns as selected items then populate those columns of the supplied DataTable
      'Only ContactName, Name & Position are checked because ContactNumber, OrganisationNumber & PhoneNumber will have automatically been added to the Selected Items list because they contain the word "number"
      If mvAddActionRelatedContactData Then
        'Build a Parameters collection of unique action numbers
        vParams = New CDBParameters
        For Each vDR In pDataTable.Rows
          If Not vParams.Exists(vDR.Item("ActionNumber")) Then vParams.Add(vDR.Item("ActionNumber"))
        Next vDR
        'Continue as long as at least one unique action number exists
        If vParams.Count > 0 Then
          'Select records from contact_actions where type = R
          'get the contact's name and default phone number
          'if the contact's default address is an organisation address then get the org. number, org. name and position
          vContact = New Contact(mvEnv)
          vContact.Init()
          vAttrs = vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone) & ", o.name, cp.position, o.organisation_number, ca.action_number"
          vSQL = mvConn.GetSelectSQLCSC & vAttrs
          vSQL = vSQL & " FROM (SELECT contact_number,action_number"
          vSQL = vSQL & " FROM contact_actions ca"
          vActionIn = " WHERE action_number IN (" & vParams.ItemList & ") AND type = 'R'"
          vSQL = vSQL & vActionIn
          vSQL = vSQL & " UNION SELECT organisation_number, action_number"
          vSQL = vSQL & " FROM organisation_actions oa"
          vSQL = vSQL & vActionIn & " ) ca"
          vSQL = vSQL & " INNER JOIN contacts c ON ca.contact_number = c.contact_number"
          vSQL = vSQL & " LEFT OUTER JOIN contact_positions cp ON c.contact_number = cp.contact_number AND c.address_number = cp.address_number"
          vSQL = vSQL & " LEFT OUTER JOIN organisations o ON cp.organisation_number = o.organisation_number"
          'Now that we've used the ItemList property reinitialise the local Parameters collection
          vParams = New CDBParameters
          'Ensure that the SELECT can be executed on Oracle databases
          vSQL = mvEnv.Connection.ProcessAnsiJoins(vSQL)
          'Build the record set
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          With vRS
            While .Fetch() = True
              For Each vDR In pDataTable.Rows
                If Val(vDR.Item("ActionNumber")) = .Fields.Item("action_number").IntegerValue Then
                  If Not vParams.Exists(vDR.Item("ActionNumber")) Then
                    'This is the first time this action has been encountered in this loop
                    vParams.Add(vDR.Item("ActionNumber"))
                    vContact = New Contact(mvEnv)
                    vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtPhone)
                    vDR.Item("ContactNumber") = CStr(vContact.ContactNumber)
                    vDR.Item("ContactName") = vContact.Name
                    vDR.Item("Name") = .Fields.Item("name").Value
                    vDR.Item("Position") = .Fields.Item("position").Value
                    vDR.Item("PhoneNumber") = vContact.PhoneNumber
                    vDR.Item("OrganisationNumber") = .Fields.Item("organisation_number").Value
                  Else
                    'This is not the first time this action has been encountered in this loop
                    'This action is related to multiple contacts, so we can't display an individual contact's details
                    vDR.Item("ContactNumber") = ""
                    vDR.Item("ContactName") = "<Multiple Related Contacts>"
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
        End If
      End If
    End Sub

    Private Sub AddClauseValue(ByRef pWhere As String, ByVal pAttr As String, ByRef pValue As String, Optional ByVal pFieldType As CDBField.FieldTypes = CDBField.FieldTypes.cftCharacter, Optional ByVal pWhereOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual)
      Dim vPos As Integer
      Dim vAttrName As String

      vPos = InStr(pAttr, ".")
      If vPos > 0 Then
        vAttrName = Mid(pAttr, vPos + 1)
      Else
        vAttrName = pAttr
      End If
      mvSelectItems.Add(vAttrName, pFieldType, pValue, pWhereOperator)
      AddClause(pWhere, pAttr)
    End Sub

    Private Function ActionLinkSQL() As String
      Dim vMyActions As Boolean
      Dim vMyResponsibilities As Boolean
      Dim vMyRelatedActions As Boolean
      Dim vSQL As String = ""
      Dim vNumber As Integer
      Dim vWhereFields As New CDBFields
      Dim vContact As New Contact(mvEnv)
      Dim vContactType As String

      If mvSelectItems.Exists("my_actions") Then vMyActions = mvSelectItems("my_actions").Bool
      If mvSelectItems.Exists("my_responsibilities") Then vMyResponsibilities = mvSelectItems("my_responsibilities").Bool
      If mvSelectItems.Exists("my_related_actions") Then vMyRelatedActions = mvSelectItems("my_related_actions").Bool

      If vMyActions Then GetActionLinkSQL(vSQL, IActionLink.ActionLinkTypes.altActioner, "C", "contact_actions", "contact_number", mvEnv.User.ContactNumber) 'Actioner contacts
      If vMyResponsibilities Then GetActionLinkSQL(vSQL, IActionLink.ActionLinkTypes.altManager, "C", "contact_actions", "contact_number", mvEnv.User.ContactNumber) 'Manager contacts
      If vMyRelatedActions Then GetActionLinkSQL(vSQL, IActionLink.ActionLinkTypes.altRelated, "C", "contact_actions", "contact_number", mvEnv.User.ContactNumber) 'Get Related contacts

      If mvSelectItems.Exists("contact_number") Then
        vNumber = mvSelectItems("contact_number").IntegerValue
        vContact.Init(vNumber)
        Select Case vContact.ContactType
          Case Contact.ContactTypes.ctcOrganisation
            vContactType = "organisation"
          Case Else
            vContactType = "contact"
        End Select
        If vSQL.Length > 0 Then vSQL = vSQL & " AND "
        vWhereFields.Add(vContactType & "_number", CDBField.FieldTypes.cftLong, vNumber)
        If mvSelectItems.Exists("type") Then
          vWhereFields.Add("type", CDBField.FieldTypes.cftCharacter, mvSelectItems("type").Value, CDBField.FieldWhereOperators.fwoInOrEqual)
        Else
          vWhereFields.Add("type", CDBField.FieldTypes.cftCharacter, "'A','M','R'", CDBField.FieldWhereOperators.fwoIn)
        End If
        vSQL = vSQL & "a.action_number IN (SELECT DISTINCT alnkcn.action_number FROM " & vContactType & "_actions alnkcn WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & ")"
      End If
      If vSQL.Length > 0 Then vSQL = "(" & vSQL & ")"
      ActionLinkSQL = vSQL
    End Function

    Private Sub GetActionLinkSQL(ByRef pSQL As String, ByRef pTypeCode As IActionLink.ActionLinkTypes, ByRef pType As String, ByRef pTable As String, ByRef pAttr As String, ByRef pNumber As Integer)
      Dim vAlias As String
      Dim vTypeCode As String

      Select Case pTypeCode
        Case IActionLink.ActionLinkTypes.altActioner
          vTypeCode = "A"
        Case IActionLink.ActionLinkTypes.altManager
          vTypeCode = "M"
        Case Else
          vTypeCode = "R"
      End Select
      vAlias = "alnk" & pTypeCode & pType
      If pSQL.Length > 0 Then pSQL = pSQL & " AND "
      pSQL = pSQL & "a.action_number IN (SELECT DISTINCT " & vAlias & ".action_number FROM " & pTable & " " & vAlias & " WHERE " & pAttr & "= " & pNumber & " AND type = '" & vTypeCode & "')"
    End Sub

  End Class


End Namespace
