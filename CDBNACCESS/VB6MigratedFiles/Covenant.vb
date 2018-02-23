

Namespace Access
  Public Class Covenant

    Public Enum CovenantRecordSetTypes 'These are bit values
      crtAll = &HFFFFS
      'ADD additional recordset types here
      crtMain = 1
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CovenantFields
      cfAll = 0
      cfContactNumber
      cfAddressNumber
      cfCovenantNumber
      cfCovenantType
      cfCovenantTerm
      cfFixed
      cfCovenantStatus
      cfCovenantedAmount
      cfStartDate
      cfSignatureDate
      cfDepositedDeed
      cfNet
      cfAnnualClaim
      cfPaymentPlanNumber
      cfLastTaxClaim
      cfTaxClaimedTo
      cfR185Return
      cfR185Sent
      cfCancellationReason
      cfCancelledOn
      cfCancelledBy
      cfSource
      cfAmendedBy
      cfAmendedOn
      cfCreatedBy
      cfCreatedOn
      cfPaymentNumber
      cfR185PaymentNumber
      cfInitialPaymentNumber
      cfCancellationSource
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvExistingPaymentPlan As Boolean
    Private mvPaymentFrequencyInterval As Integer
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "covenants"
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("covenant_number", CDBField.FieldTypes.cftLong)
          .Add("covenant_type")
          .Add("covenant_term", CDBField.FieldTypes.cftInteger)
          .Add("fixed")
          .Add("covenant_status")
          .Add("covenanted_amount", CDBField.FieldTypes.cftNumeric)
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("signature_date", CDBField.FieldTypes.cftDate)
          .Add("deposited_deed")
          .Add("net")
          .Add("annual_claim")
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("last_tax_claim", CDBField.FieldTypes.cftDate)
          .Add("tax_claimed_to", CDBField.FieldTypes.cftDate)
          .Add("r185_return", CDBField.FieldTypes.cftDate)
          .Add("r185_sent", CDBField.FieldTypes.cftDate)
          .Add("cancellation_reason")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("cancelled_by")
          .Add("source")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
          .Add("payment_number", CDBField.FieldTypes.cftInteger)
          .Add("r185_payment_number", CDBField.FieldTypes.cftInteger)
          .Add("initial_payment_number", CDBField.FieldTypes.cftInteger)
          .Add("cancellation_source")
        End With

        mvClassFields.Item(CovenantFields.cfCovenantNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(CovenantFields.cfContactNumber).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfAddressNumber).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfPaymentPlanNumber).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCancellationReason).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCancelledOn).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCancelledBy).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfPaymentNumber).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfPaymentPlanNumber).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfAmendedBy).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfAmendedOn).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfSource).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfStartDate).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCancellationSource).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCreatedBy).PrefixRequired = True
        mvClassFields.Item(CovenantFields.cfCreatedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Public Sub SetCancelled(ByRef pCancelledReason As String, ByRef pCancelledOn As String, ByRef pCancelledBy As String, Optional ByRef pCancellationSource As String = "")
      mvClassFields.Item(CovenantFields.cfCancellationReason).Value = pCancelledReason
      mvClassFields.Item(CovenantFields.cfCancelledOn).Value = pCancelledOn
      mvClassFields.Item(CovenantFields.cfCancelledBy).Value = pCancelledBy
      If Len(pCancellationSource) > 0 Then mvClassFields.Item(CovenantFields.cfCancellationSource).Value = pCancellationSource
    End Sub

    Public Sub SetUnCancelled()
      With mvClassFields
        .Item(CovenantFields.cfCancellationReason).Value = ""
        .Item(CovenantFields.cfCancelledOn).Value = ""
        .Item(CovenantFields.cfCancelledBy).Value = ""
        .Item(CovenantFields.cfCancellationSource).Value = ""
      End With
    End Sub

    Public Sub SetContact(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer)
      mvClassFields.Item(CovenantFields.cfContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(CovenantFields.cfAddressNumber).Value = CStr(pAddressNumber)
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Public Sub SetImportDetails(ByRef pCovTerm As Integer, ByRef pCovFixed As Boolean, ByRef pCovStatus As String, ByRef pCovSigDate As String, ByRef pCovDepDeed As Boolean, ByRef pCovNet As Boolean, ByRef pCovAnnual As Boolean, ByRef pvCovLastTaxClaim As String, ByRef pCovTaxClaimTo As String, ByRef pCovCreatedOn As String)

      mvClassFields.Item(CovenantFields.cfCovenantType).Value = "I" 'Assume individual
      mvClassFields.Item(CovenantFields.cfCovenantTerm).Value = CStr(pCovTerm)
      mvClassFields.Item(CovenantFields.cfFixed).Bool = pCovFixed
      mvClassFields.Item(CovenantFields.cfCovenantStatus).Value = pCovStatus
      mvClassFields.Item(CovenantFields.cfSignatureDate).Value = pCovSigDate
      mvClassFields.Item(CovenantFields.cfDepositedDeed).Bool = pCovDepDeed
      mvClassFields.Item(CovenantFields.cfNet).Bool = pCovNet
      mvClassFields.Item(CovenantFields.cfAnnualClaim).Bool = pCovAnnual
      mvClassFields.Item(CovenantFields.cfLastTaxClaim).Value = pvCovLastTaxClaim
      mvClassFields.Item(CovenantFields.cfTaxClaimedTo).Value = pCovTaxClaimTo
      mvClassFields.Item(CovenantFields.cfCreatedOn).Value = pCovCreatedOn
    End Sub

    Public Sub SetFromPayPlan(ByRef pCovAmount As Double, ByRef pStartDate As String, ByRef pPlanNumber As Integer, ByRef pSource As String)

      mvClassFields.Item(CovenantFields.cfCovenantedAmount).Value = CStr(pCovAmount)
      mvClassFields.Item(CovenantFields.cfStartDate).Value = pStartDate
      mvClassFields.Item(CovenantFields.cfPaymentPlanNumber).Value = CStr(pPlanNumber)
      mvClassFields.Item(CovenantFields.cfSource).Value = pSource
    End Sub

    Private Sub SetValid(ByRef pField As CovenantFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CovenantFields.cfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CovenantFields.cfAmendedBy).Value = mvEnv.User.Logname

      If Len(mvClassFields.Item(CovenantFields.cfCreatedBy).Value) = 0 Then mvClassFields.Item(CovenantFields.cfCreatedBy).Value = mvEnv.User.Logname
      If Len(mvClassFields.Item(CovenantFields.cfCreatedOn).Value) = 0 Then mvClassFields.Item(CovenantFields.cfCreatedOn).Value = TodaysDate() 'mvClassFields.Item(cfStartDate).Value

      If Len(mvClassFields.Item(CovenantFields.cfPaymentNumber).Value) = 0 Then mvClassFields.Item(CovenantFields.cfPaymentNumber).IntegerValue = 0
      If Len(mvClassFields.Item(CovenantFields.cfInitialPaymentNumber).Value) = 0 Then mvClassFields.Item(CovenantFields.cfInitialPaymentNumber).IntegerValue = 0
      If Len(mvClassFields.Item(CovenantFields.cfR185PaymentNumber).Value) = 0 Then mvClassFields.Item(CovenantFields.cfR185PaymentNumber).IntegerValue = 0
      If Len(mvClassFields.Item(CovenantFields.cfCovenantNumber).Value) = 0 Then mvClassFields.Item(CovenantFields.cfCovenantNumber).IntegerValue = mvEnv.GetControlNumber("CD")
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CovenantRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CovenantRecordSetTypes.crtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cv")
      Else
        If (pRSType And CovenantRecordSetTypes.crtMain) > 0 Then vFields = "covenant_number,contact_number,address_number,covenant_term,start_date,created_on,created_by"
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCovenantNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      If pCovenantNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CovenantRecordSetTypes.crtAll) & " FROM covenants cv WHERE covenant_number = " & pCovenantNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CovenantRecordSetTypes.crtAll)
        Else
          System.Diagnostics.Debug.Assert(True, "") 'Contact Class Init - Record not found / Locked
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Sub InitFromPaymentPlan(ByVal pEnv As CDBEnvironment, ByRef pPaymentPlanNumber As Integer, Optional ByRef pCancelled As Boolean = False)
      Dim vAndCancellation As String
      Dim vRecordSet As CDBRecordSet

      If pPaymentPlanNumber > 0 Then
        mvEnv = pEnv
        InitClassFields()

        If Not pCancelled Then
          vAndCancellation = " AND cancellation_reason IS NULL"
        Else
          vAndCancellation = " AND cancellation_reason IS NOT NULL ORDER BY cancelled_on DESC, covenant_number DESC"
        End If
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CovenantRecordSetTypes.crtAll) & " FROM covenants cv WHERE order_number = " & pPaymentPlanNumber & vAndCancellation)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CovenantRecordSetTypes.crtAll)
        Else
          Init(pEnv)
        End If
        vRecordSet.CloseRecordSet()
      Else
        Init(pEnv)
      End If
    End Sub

    Sub InitFromPaymentPlanGA(ByVal pEnv As CDBEnvironment, ByRef pPaymentPlanNumber As Integer, ByRef pTransactionDate As String)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv

      If pPaymentPlanNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CovenantRecordSetTypes.crtAll) & " FROM covenants cv WHERE order_number = " & pPaymentPlanNumber & " AND ((cancellation_reason IS NULL) OR (cancellation_reason IS NOT NULL AND cancelled_on" & pEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pTransactionDate) & ")) AND deposited_deed = 'N'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CovenantRecordSetTypes.crtAll)
        Else
          Init(pEnv)
        End If
        vRecordSet.CloseRecordSet()
      Else
        Init(pEnv)
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CovenantRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Modify below to handle each recordset type as required
        If (pRSType And CovenantRecordSetTypes.crtMain) > 0 Then
          .SetItem(CovenantFields.cfContactNumber, vFields)
          .SetItem(CovenantFields.cfAddressNumber, vFields)
          .SetItem(CovenantFields.cfCovenantNumber, vFields)
          .SetItem(CovenantFields.cfStartDate, vFields)
          .SetItem(CovenantFields.cfCovenantTerm, vFields)
          .SetItem(CovenantFields.cfCreatedBy, vFields)
          .SetItem(CovenantFields.cfCreatedOn, vFields)
        End If
        If (pRSType And CovenantRecordSetTypes.crtAll) = CovenantRecordSetTypes.crtAll Then
          .SetItem(CovenantFields.cfCovenantType, vFields)
          .SetItem(CovenantFields.cfFixed, vFields)
          .SetItem(CovenantFields.cfCovenantStatus, vFields)
          .SetItem(CovenantFields.cfCovenantedAmount, vFields)
          .SetItem(CovenantFields.cfSignatureDate, vFields)
          .SetItem(CovenantFields.cfDepositedDeed, vFields)
          .SetItem(CovenantFields.cfNet, vFields)
          .SetItem(CovenantFields.cfAnnualClaim, vFields)
          .SetItem(CovenantFields.cfPaymentPlanNumber, vFields)
          .SetItem(CovenantFields.cfLastTaxClaim, vFields)
          .SetItem(CovenantFields.cfTaxClaimedTo, vFields)
          .SetItem(CovenantFields.cfR185Return, vFields)
          .SetItem(CovenantFields.cfR185Sent, vFields)
          .SetItem(CovenantFields.cfCancellationReason, vFields)
          .SetItem(CovenantFields.cfCancelledOn, vFields)
          .SetItem(CovenantFields.cfCancelledBy, vFields)
          .SetItem(CovenantFields.cfSource, vFields)
          .SetItem(CovenantFields.cfAmendedBy, vFields)
          .SetItem(CovenantFields.cfAmendedOn, vFields)
          .SetItem(CovenantFields.cfPaymentNumber, vFields)
          .SetItem(CovenantFields.cfR185PaymentNumber, vFields)
          .SetItem(CovenantFields.cfInitialPaymentNumber, vFields)
          .SetOptionalItem(CovenantFields.cfCancellationSource, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0)
      Dim vToDate As String

      SetValid(CovenantFields.cfAll)
      If Not Existing Then
        'Create the journal record
        mvEnv.AddJournalRecord(JournalTypes.jnlCovenant, JournalOperations.jnlInsert, ContactNumber, AddressNumber, CovenantNumber, 0, 0, pBatchNumber, pTransactionNumber)
        'Check for any qualifying pre-existing payments
        If mvExistingPaymentPlan And mvPaymentFrequencyInterval > 0 And Not DepositedDeed Then ClaimProcessedPayments()
        'Create the activity
        If Len(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivity)) > 0 And Len(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivityValue)) > 0 Then
          vToDate = CDate(StartDate).AddYears(CovenantTerm).AddDays(-1).ToString(CAREDateFormat)
          Dim vCC As New ContactCategory(mvEnv)
          vCC.ContactTypeSaveActivity(Contact.ContactTypes.ctcContact, ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivity), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVActivityValue), Source, StartDate, vToDate, "", ContactCategory.ActivityEntryStyles.aesCheckDateRange)
        End If
      End If
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub
    Public Sub SaveChanges()
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property EndDate() As Date
      Get
        Dim vTerm As Integer

        vTerm = CInt(mvClassFields.Item(CovenantFields.cfCovenantTerm).Value)
        EndDate = CDate(StartDate).AddYears(vTerm)
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(CovenantFields.cfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CovenantFields.cfAddressNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CovenantFields.cfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CovenantFields.cfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AnnualClaim() As Boolean
      Get
        AnnualClaim = mvClassFields.Item(CovenantFields.cfAnnualClaim).Bool
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(CovenantFields.cfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(CovenantFields.cfCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(CovenantFields.cfCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(CovenantFields.cfCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(CovenantFields.cfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CovenantNumber() As Integer
      Get
        CovenantNumber = mvClassFields.Item(CovenantFields.cfCovenantNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CovenantStatus() As String
      Get
        CovenantStatus = mvClassFields.Item(CovenantFields.cfCovenantStatus).Value
      End Get
    End Property

    Public ReadOnly Property CovenantTerm() As Integer
      Get
        CovenantTerm = mvClassFields.Item(CovenantFields.cfCovenantTerm).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CovenantType() As String
      Get
        CovenantType = mvClassFields.Item(CovenantFields.cfCovenantType).Value
      End Get
    End Property

    Public ReadOnly Property CovenantedAmount() As Double
      Get
        CovenantedAmount = mvClassFields.Item(CovenantFields.cfCovenantedAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property CreatedBy() As String
      Get
        CreatedBy = mvClassFields.Item(CovenantFields.cfCreatedBy).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(CovenantFields.cfCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property DepositedDeed() As Boolean
      Get
        DepositedDeed = mvClassFields.Item(CovenantFields.cfDepositedDeed).Bool
      End Get
    End Property

    Public ReadOnly Property Fixed() As Boolean
      Get
        Fixed = mvClassFields.Item(CovenantFields.cfFixed).Bool
      End Get
    End Property

    Public ReadOnly Property InitialPaymentNumber() As Integer
      Get
        InitialPaymentNumber = mvClassFields.Item(CovenantFields.cfInitialPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LastTaxClaim() As String
      Get
        LastTaxClaim = mvClassFields.Item(CovenantFields.cfLastTaxClaim).Value
      End Get
    End Property

    Public ReadOnly Property Net() As Boolean
      Get
        Net = mvClassFields.Item(CovenantFields.cfNet).Bool
      End Get
    End Property

    Public Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(CovenantFields.cfPaymentPlanNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CovenantFields.cfPaymentPlanNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property PaymentNumber() As Integer
      Get
        PaymentNumber = mvClassFields.Item(CovenantFields.cfPaymentNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CovenantFields.cfPaymentNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property R185PaymentNumber() As Integer
      Get
        R185PaymentNumber = mvClassFields.Item(CovenantFields.cfR185PaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property R185Return() As String
      Get
        R185Return = mvClassFields.Item(CovenantFields.cfR185Return).Value
      End Get
    End Property

    Public ReadOnly Property R185Sent() As String
      Get
        R185Sent = mvClassFields.Item(CovenantFields.cfR185Sent).Value
      End Get
    End Property

    Public ReadOnly Property SignatureDate() As String
      Get
        SignatureDate = mvClassFields.Item(CovenantFields.cfSignatureDate).Value
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(CovenantFields.cfSource).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(CovenantFields.cfStartDate).Value
      End Get
    End Property

    Public Property TaxClaimedTo() As String
      Get
        TaxClaimedTo = mvClassFields.Item(CovenantFields.cfTaxClaimedTo).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CovenantFields.cfTaxClaimedTo).Value = Value
      End Set
    End Property

    Public ReadOnly Property IsValid() As Boolean
      Get
        IsValid = True

        If IsDate(StartDate) Then
          If CDate(StartDate) > CDate(TodaysDate()) Then IsValid = False
          If CovenantTerm > 0 Then
            If CDate(StartDate).AddYears(CovenantTerm) < CDate(TodaysDate()) Then IsValid = False
          End If
        End If
        If IsDate(CancelledOn) Then
          If CDate(CancelledOn) <= CDate(TodaysDate()) Then IsValid = False
        End If

      End Get
    End Property

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pCovenantType As String, ByVal pCovenantTerm As Integer, ByVal pCovenantStatus As String, ByVal pFixed As Boolean, ByVal pCovenantedAmount As Double, ByVal pStartDate As String, ByVal pSignatureDate As String, ByVal pDepositedDeed As Boolean, ByVal pNet As Boolean, ByVal pAnnualClaim As Boolean, ByVal pPaymentPlanNumber As Integer, ByVal pR185Return As String, ByVal pSource As String, Optional ByVal pExistingPaymentPlan As Boolean = False, Optional ByVal pPaymentFrequencyInterval As Integer = 0)
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        .Item(CovenantFields.cfContactNumber).IntegerValue = pContactNumber
        .Item(CovenantFields.cfAddressNumber).IntegerValue = pAddressNumber
        .Item(CovenantFields.cfCovenantType).Value = pCovenantType
        .Item(CovenantFields.cfCovenantTerm).IntegerValue = pCovenantTerm
        .Item(CovenantFields.cfFixed).Bool = pFixed
        .Item(CovenantFields.cfCovenantStatus).Value = pCovenantStatus
        .Item(CovenantFields.cfCovenantedAmount).Value = CStr(pCovenantedAmount)
        .Item(CovenantFields.cfStartDate).Value = pStartDate
        .Item(CovenantFields.cfSignatureDate).Value = pSignatureDate
        .Item(CovenantFields.cfDepositedDeed).Bool = pDepositedDeed
        .Item(CovenantFields.cfNet).Bool = pNet
        .Item(CovenantFields.cfAnnualClaim).Bool = pAnnualClaim
        .Item(CovenantFields.cfPaymentPlanNumber).IntegerValue = pPaymentPlanNumber
        .Item(CovenantFields.cfR185Return).Value = pR185Return
        .Item(CovenantFields.cfSource).Value = pSource
      End With
      mvExistingPaymentPlan = pExistingPaymentPlan
      mvPaymentFrequencyInterval = pPaymentFrequencyInterval
    End Sub
    Private Sub ClaimProcessedPayments()
      Dim vSQL As String
      Dim vInsertFields As New CDBFields
      Dim vRS As CDBRecordSet
      Dim vOpDate As String = ""
      Dim vNextDueDate As Date
      Dim vClaimPayment As Boolean
      Dim vRollDueDate As Boolean
      Dim vAmountPaidForPeriod As Double
      Dim vGrace As Integer

      'Create declaration_lines_unclaimed records

      If IsDate(mvEnv.GetConfig("ga_operational_claim_date")) Then vOpDate = CDate(mvEnv.GetConfig("ga_operational_claim_date")).ToString(CAREDateFormat)
      If vOpDate.Length = 0 Then vOpDate = DateSerial(2000, 4, 6).ToString(CAREDateFormat)
      vGrace = IntegerValue(mvEnv.GetConfig("cv_no_days_claim_grace"))
      If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(StartDate), CDate(vOpDate)) > 0 Then
        If AnnualClaim Or mvPaymentFrequencyInterval = 12 Then
          vNextDueDate = CDate(StartDate).AddYears(1)
        Else
          vNextDueDate = DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvPaymentFrequencyInterval, CDate(StartDate))
        End If
      End If

      vAmountPaidForPeriod = 0

      With vInsertFields
        .Add("cd_number", CDBField.FieldTypes.cftLong, CovenantNumber)
        .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftLong)
        .Add("line_number", CDBField.FieldTypes.cftInteger)
        .Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "C")
        .Add("net_amount", CDBField.FieldTypes.cftNumeric)
      End With

      vSQL = "SELECT oph.batch_number,oph.transaction_number,oph.line_number,oph.amount,fh.transaction_date"
      vSQL = vSQL & " FROM order_payment_history oph,financial_history fh"
      vSQL = vSQL & " WHERE oph.order_number = " & PaymentPlanNumber & " AND oph.status IS NULL AND oph.amount > 0"
      vSQL = vSQL & " AND fh.batch_number = oph.batch_number AND fh.transaction_number = oph.transaction_number AND fh.transaction_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartDate)
      'Check not already claimed under Gift Aid scheme
      vSQL = vSQL & " AND oph.batch_number NOT IN (SELECT batch_number FROM declaration_lines_unclaimed dlu WHERE dlu.batch_number = oph.batch_number AND dlu.transaction_number = oph.transaction_number AND dlu.line_number = oph.line_number)"
      'Check not already claimed under R68 scheme
      vSQL = vSQL & " AND oph.order_number NOT IN (SELECT order_number FROM covenants c,tax_claim_lines tcl WHERE c.covenant_number = " & CovenantNumber & " AND tcl.covenant_number = c.covenant_number AND tcl.start_payment_number <=  oph.payment_number AND tcl.end_payment_number >= oph.payment_number)"
      vSQL = vSQL & " ORDER BY oph.payment_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      With vRS
        While .Fetch() = True
          vClaimPayment = False
          vRollDueDate = False
          If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(StartDate), CDate(vOpDate)) <= 0 Then
            vClaimPayment = True
            vAmountPaidForPeriod = CDbl(.Fields(4).Value) 'amount
          ElseIf DateDiff(Microsoft.VisualBasic.DateInterval.Day, vNextDueDate, CDate(vOpDate)) <= 0 Then
            vClaimPayment = True
            vAmountPaidForPeriod = CDbl(.Fields(4).Value) 'amount
          Else
            vAmountPaidForPeriod = vAmountPaidForPeriod + CDbl(.Fields(4).Value) 'amount
            If vAmountPaidForPeriod >= CovenantedAmount Then
              vRollDueDate = True
            Else
              If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(.Fields(5).Value), DateAdd(Microsoft.VisualBasic.DateInterval.Day, vGrace, vNextDueDate)) < 0 Then
                'Shortfall received for that period
                vRollDueDate = True
              End If
            End If
          End If
          If vClaimPayment Then
            vInsertFields(3).Value = .Fields(1).Value
            vInsertFields(4).Value = .Fields(2).Value
            vInsertFields(5).Value = .Fields(3).Value
            vInsertFields(7).Value = .Fields(4).Value
            mvEnv.Connection.InsertRecord("declaration_lines_unclaimed", vInsertFields)
          End If
          If vRollDueDate Then
            If AnnualClaim Or mvPaymentFrequencyInterval = 12 Then
              vNextDueDate = vNextDueDate.AddYears(1)
            Else
              vNextDueDate = DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvPaymentFrequencyInterval, vNextDueDate)
            End If
            vAmountPaidForPeriod = 0
          End If
        End While
        .CloseRecordSet()
      End With
    End Sub
  End Class
End Namespace
