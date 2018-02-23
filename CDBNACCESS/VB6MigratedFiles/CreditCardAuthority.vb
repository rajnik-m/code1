

Namespace Access
  Public Class CreditCardAuthority
    Implements IAutoPaymentMethod

    Public Enum CreditCardAuthorityRecordSetTypes 'These are bit values
      ccartAll = &HFFFFS
      'ADD additional recordset types here
      ccartNumber = 1
      ccartDetail = 2
      ccartBankInfo = 4
      ccartCancelled = 8
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CreditCardAuthorityFields
      ccafAll = 0
      ccafCreditCardAuthorityNumber
      ccafPaymentPlanNumber
      ccafContactNumber
      ccafAddressNumber
      ccafCreditCardDetailsNumber
      ccafBankAccount
      ccafStartDate
      ccafSource
      ccafCancellationReason
      ccafCancelledBy
      ccafCancelledOn
      ccafAmount
      ccafAmendedBy
      ccafAmendedOn
      ccafAuthorityType
      ccafCancellationSource
      ccafCreatedBy
      ccafCreatedOn
      ccafFutureCancellationReason
      ccafFutureCancellationDate
      ccafFuturecancellationSource
    End Enum

    Public Enum ccaAuthorityType
      catCreditCard
      catCAFCard
    End Enum

    Private mvContactCreditCard As ContactCreditCard

    Private mvActivity As String
    Private mvActivityValue As String
    Private mvContactType As Contact.ContactTypes

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAmendedValid As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "credit_card_authorities"
          .Add("credit_card_authority_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("credit_card_details_number", CDBField.FieldTypes.cftLong)
          .Add("bank_account")
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("source")
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("authority_type")
          .Add("cancellation_source")
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_reason")
          .Add("future_cancellation_date", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_source")

          .Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).SetPrimaryKeyOnly()

          .Item(CreditCardAuthorityFields.ccafPaymentPlanNumber).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafContactNumber).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafAddressNumber).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafSource).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCancellationReason).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCancellationSource).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCancelledBy).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCancelledOn).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafAmount).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCreatedBy).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafCreatedOn).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafAmendedBy).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafAmendedOn).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafFutureCancellationDate).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafFutureCancellationReason).PrefixRequired = True
          .Item(CreditCardAuthorityFields.ccafFuturecancellationSource).PrefixRequired = True
        End With

        mvActivity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCCCAActivity)
        mvActivityValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCCCAActivityValue)
      Else
        mvClassFields.ClearItems()
      End If
      mvContactCreditCard = Nothing
      mvContactType = Contact.ContactTypes.ctcContact 'Default
      mvAmendedValid = False
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CreditCardAuthorityFields)
      'Add code here to ensure all values are valid before saving
      If Not mvAmendedValid Then
        mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedOn).Value = TodaysDate()
        mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedBy).Value = mvEnv.User.UserID
      End If

      If pField = CreditCardAuthorityFields.ccafAll And Len(mvClassFields.Item(CreditCardAuthorityFields.ccafCreatedBy).Value) = 0 And mvExisting = False Then
        mvClassFields.Item(CreditCardAuthorityFields.ccafCreatedBy).Value = mvEnv.User.UserID
        mvClassFields.Item(CreditCardAuthorityFields.ccafCreatedOn).Value = TodaysDate()
      End If
    End Sub

    Public Sub SetContact(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer)
      mvClassFields.Item(CreditCardAuthorityFields.ccafContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(CreditCardAuthorityFields.ccafAddressNumber).Value = CStr(pAddressNumber)
    End Sub

    Private Function GetAuthorityType(ByVal pAuthorityTypeCode As String) As ccaAuthorityType
      Select Case pAuthorityTypeCode
        Case "C"
          GetAuthorityType = ccaAuthorityType.catCAFCard
        Case Else 'A (or null?)
          GetAuthorityType = ccaAuthorityType.catCreditCard
      End Select
    End Function

    Public Function GetAuthorityTypeCode(ByVal pAuthorityType As ccaAuthorityType) As String
      Select Case pAuthorityType
        Case ccaAuthorityType.catCreditCard
          GetAuthorityTypeCode = "A"
        Case Else 'catCAFCard
          GetAuthorityTypeCode = "C"
      End Select
    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CreditCardAuthorityRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CreditCardAuthorityRecordSetTypes.ccartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ccca")
      Else
        vFields = "credit_card_authority_number"
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartNumber) > 0 Then vFields = vFields & ",ccca.contact_number,ccca.address_number,ccca.credit_card_details_number"
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartDetail) > 0 Then vFields = vFields & ",ccca.source"
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartBankInfo) > 0 Then vFields = vFields & ",ccca.bank_account,ccca.amount"
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartCancelled) > 0 Then vFields = vFields & ",ccca.cancellation_reason,ccca.future_cancellation_reason,ccca.future_cancellation_date,ccca.future_cancellation_source"
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pCreditCardAuthorityNumber As Integer = 0) Implements IAutoPaymentMethod.Init
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCreditCardAuthorityNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CreditCardAuthorityRecordSetTypes.ccartAll) & " FROM credit_card_authorities ccca WHERE credit_card_authority_number = " & pCreditCardAuthorityNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CreditCardAuthorityRecordSetTypes.ccartAll)
        Else
          InitClassFields()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Friend Sub InitFromPaymentPlan(ByVal pEnv As CDBEnvironment, ByRef pPaymentPlanNumber As Integer, Optional ByRef pCancelled As Boolean = False)
      Dim vRecordSet As CDBRecordSet
      Dim vAndCancellation As String
      Dim vCCAType As CreditCardAuthorityRecordSetTypes

      If pPaymentPlanNumber > 0 Then
        mvEnv = pEnv
        InitClassFields()
        If Not pCancelled Then
          vAndCancellation = " AND cancellation_reason IS NULL"
          vCCAType = CreditCardAuthorityRecordSetTypes.ccartAll
        Else
          vAndCancellation = " AND cancellation_reason IS NOT NULL ORDER BY cancelled_on DESC"
          vCCAType = CreditCardAuthorityRecordSetTypes.ccartAll
        End If
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(vCCAType) & " FROM credit_card_authorities ccca WHERE order_number = " & pPaymentPlanNumber & vAndCancellation)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, vCCAType)
        Else
          Init(pEnv)
        End If
        vRecordSet.CloseRecordSet()
      Else
        Init(pEnv)
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CreditCardAuthorityRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Modify below to handle each recordset type as required
        .SetItem(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber, vFields)
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartNumber) > 0 Then
          .SetItem(CreditCardAuthorityFields.ccafContactNumber, vFields)
          .SetItem(CreditCardAuthorityFields.ccafAddressNumber, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCreditCardDetailsNumber, vFields)
        End If
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartDetail) > 0 Then
          .SetItem(CreditCardAuthorityFields.ccafSource, vFields)
        End If
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartBankInfo) > 0 Then
          .SetItem(CreditCardAuthorityFields.ccafBankAccount, vFields)
          .SetItem(CreditCardAuthorityFields.ccafAmount, vFields)
        End If
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartCancelled) > 0 Then
          .SetItem(CreditCardAuthorityFields.ccafCancellationReason, vFields)
          .SetItem(CreditCardAuthorityFields.ccafFutureCancellationReason, vFields)
          .SetItem(CreditCardAuthorityFields.ccafFutureCancellationDate, vFields)
          .SetItem(CreditCardAuthorityFields.ccafFuturecancellationSource, vFields)
        End If
        If (pRSType And CreditCardAuthorityRecordSetTypes.ccartAll) = CreditCardAuthorityRecordSetTypes.ccartAll Then
          .SetItem(CreditCardAuthorityFields.ccafPaymentPlanNumber, vFields)
          .SetItem(CreditCardAuthorityFields.ccafBankAccount, vFields)
          .SetItem(CreditCardAuthorityFields.ccafStartDate, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCancelledBy, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCancelledOn, vFields)
          .SetItem(CreditCardAuthorityFields.ccafAmount, vFields)
          .SetItem(CreditCardAuthorityFields.ccafAmendedBy, vFields)
          .SetItem(CreditCardAuthorityFields.ccafAmendedOn, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCancellationSource, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCreatedBy, vFields)
          .SetItem(CreditCardAuthorityFields.ccafCreatedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0)
      Dim vTransaction As Boolean

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      If Len(mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).Value) = 0 Then
        mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).Value = CStr(mvEnv.GetControlNumber("CA"))
      End If
      SetValid(CreditCardAuthorityFields.ccafAll)

      If Not mvContactCreditCard Is Nothing Then 'Save Contact Credit Card record
        mvContactCreditCard.Save()
        mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardDetailsNumber).Value = CStr(mvContactCreditCard.CreditCardDetailsNumber)
      End If

      If mvExisting Then
        mvEnv.Connection.UpdateRecords("credit_card_authorities", mvClassFields.UpdateFields, mvClassFields.WhereFields)
        mvEnv.AddJournalRecord(JournalTypes.jnlCreditCard, JournalOperations.jnlUpdate, mvClassFields.Item(CreditCardAuthorityFields.ccafContactNumber).IntegerValue, mvClassFields.Item(CreditCardAuthorityFields.ccafAddressNumber).IntegerValue, (mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
      Else
        mvEnv.Connection.InsertRecord("credit_card_authorities", mvClassFields.UpdateFields)
        mvEnv.AddJournalRecord(JournalTypes.jnlCreditCard, JournalOperations.jnlInsert, mvClassFields.Item(CreditCardAuthorityFields.ccafContactNumber).IntegerValue, mvClassFields.Item(CreditCardAuthorityFields.ccafAddressNumber).IntegerValue, (mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
        If Len(mvClassFields(CreditCardAuthorityFields.ccafCancellationReason).Value) > 0 Then
          CreateCancelledAutoPMActivity()
        Else
          CreateAutoPMActivity()
        End If
      End If
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub
    Public Sub CreateAutoPMActivity() Implements IAutoPaymentMethod.CreateAutoPMActivity
      If Len(mvActivity) > 0 Then
        Dim vCC As New ContactCategory(mvEnv)
        vCC.ContactTypeSaveActivity(mvContactType, ContactNumber, mvActivity, mvActivityValue, Source, StartDate, CDate(StartDate).AddYears(99).ToString(CAREDateFormat), "", ContactCategory.ActivityEntryStyles.aesCheckDateRange, "", AmendedOn, AmendedBy)
      End If
    End Sub

    Public Sub CreateCancelledAutoPMActivity()
      If Len(mvActivity) > 0 Then
        Dim vCC As New ContactCategory(mvEnv)
        vCC.ContactTypeSaveActivity(mvContactType, ContactNumber, mvActivity, mvActivityValue, Source, StartDate, CancelledOn, "", ContactCategory.ActivityEntryStyles.aesCheckDateRange, "", AmendedOn, AmendedBy)
      End If
    End Sub

    Public Sub SetUnCancelled() Implements IAutoPaymentMethod.SetUnCancelled
      mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationReason).Value = ""
      mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledOn).Value = ""
      mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledBy).Value = ""
      mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationSource).Value = ""
    End Sub

    Public Sub SaveChanges(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False) Implements IAutoPaymentMethod.SaveChanges
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pCreditCardDetailsNumber As Integer, ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pPPNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pType As String, Optional ByRef pContactType As Contact.ContactTypes = Contact.ContactTypes.ctcContact, Optional ByRef pCreatedOn As String = "", Optional ByRef pCreatedBy As String = "")
      With mvClassFields
        .Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).IntegerValue = mvEnv.GetControlNumber("CA")
        .Item(CreditCardAuthorityFields.ccafCreditCardDetailsNumber).IntegerValue = pCreditCardDetailsNumber
        .Item(CreditCardAuthorityFields.ccafContactNumber).IntegerValue = pContactNumber
        .Item(CreditCardAuthorityFields.ccafAddressNumber).IntegerValue = pAddressNumber
        .Item(CreditCardAuthorityFields.ccafPaymentPlanNumber).IntegerValue = pPPNumber
        .Item(CreditCardAuthorityFields.ccafBankAccount).Value = pBankAccount
        .Item(CreditCardAuthorityFields.ccafSource).Value = pSource
        If pAmount > 0 Then .Item(CreditCardAuthorityFields.ccafAmount).DoubleValue = pAmount
        .Item(CreditCardAuthorityFields.ccafStartDate).Value = pStartDate
        If Len(pType) = 0 Then pType = GetAuthorityTypeCode(ccaAuthorityType.catCreditCard)
        .Item(CreditCardAuthorityFields.ccafAuthorityType).Value = pType
        .Item(CreditCardAuthorityFields.ccafCreatedOn).Value = pCreatedOn
        .Item(CreditCardAuthorityFields.ccafCreatedBy).Value = pCreatedBy
        mvContactType = pContactType
      End With
      SetValid(CreditCardAuthorityFields.ccafAll)
    End Sub

    Public Sub Update(ByRef pCreditCardDetailsNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pType As String)
      With mvClassFields
        .Item(CreditCardAuthorityFields.ccafCreditCardDetailsNumber).IntegerValue = pCreditCardDetailsNumber
        .Item(CreditCardAuthorityFields.ccafBankAccount).Value = pBankAccount
        .Item(CreditCardAuthorityFields.ccafSource).Value = pSource
        If pAmount > 0 Then
          .Item(CreditCardAuthorityFields.ccafAmount).DoubleValue = pAmount
        Else
          .Item(CreditCardAuthorityFields.ccafAmount).Value = "" 'Could be changing from a figure to null
        End If
        .Item(CreditCardAuthorityFields.ccafStartDate).Value = pStartDate
        If Len(pType) > 0 Then .Item(CreditCardAuthorityFields.ccafAuthorityType).Value = pType
      End With
    End Sub

    Public Sub SetPlanNumberAndAmount(ByRef pPlanNumber As Integer, ByRef pAmount As Double)
      With mvClassFields
        .Item(CreditCardAuthorityFields.ccafPaymentPlanNumber).IntegerValue = pPlanNumber
        .Item(CreditCardAuthorityFields.ccafAmount).DoubleValue = pAmount
      End With
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean Implements IAutoPaymentMethod.Existing
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer Implements IAutoPaymentMethod.AddressNumber
      Get
        AddressNumber = mvClassFields.Item(CreditCardAuthorityFields.ccafAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(CreditCardAuthorityFields.ccafAmount).Value
      End Get
    End Property

    Public Property BankAccount() As String Implements IAutoPaymentMethod.BankAccount
      Get
        BankAccount = mvClassFields.Item(CreditCardAuthorityFields.ccafBankAccount).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CreditCardAuthorityFields.ccafBankAccount).Value = Value
      End Set
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactCreditCard() As ContactCreditCard
      Get
        If mvContactCreditCard Is Nothing Then
          mvContactCreditCard = New ContactCreditCard
          mvContactCreditCard.Init(mvEnv, (mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardDetailsNumber).IntegerValue))
        End If
        ContactCreditCard = mvContactCreditCard
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer Implements IAutoPaymentMethod.ContactNumber
      Get
        ContactNumber = mvClassFields.Item(CreditCardAuthorityFields.ccafContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreditCardAuthorityNumber() As Integer Implements IAutoPaymentMethod.AutoPaymentNumber
      Get
        CreditCardAuthorityNumber = mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreditCardDetailsNumber() As Integer
      Get
        CreditCardDetailsNumber = mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardDetailsNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(CreditCardAuthorityFields.ccafPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(CreditCardAuthorityFields.ccafSource).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(CreditCardAuthorityFields.ccafStartDate).Value
      End Get
    End Property

    Public ReadOnly Property AuthorityType() As ccaAuthorityType
      Get
        AuthorityType = GetAuthorityType(mvClassFields.Item(CreditCardAuthorityFields.ccafAuthorityType).Value)
      End Get
    End Property

    Public ReadOnly Property CreatedBy() As String
      Get
        CreatedBy = mvClassFields.Item(CreditCardAuthorityFields.ccafCreatedBy).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(CreditCardAuthorityFields.ccafCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationReason() As String
      Get
        FutureCancellationReason = mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationDate() As String
      Get
        FutureCancellationDate = mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationDate).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationSource() As String
      Get
        FutureCancellationSource = mvClassFields.Item(CreditCardAuthorityFields.ccafFuturecancellationSource).Value
      End Get
    End Property

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedOn).Value = pAmendedOn
      mvClassFields.Item(CreditCardAuthorityFields.ccafAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Sub SetCancelled(ByRef pCancellationReason As String, Optional ByRef pCancelledOn As String = "", Optional ByRef pCancelledBy As String = "", Optional ByRef pCancellationSource As String = "")
      If Len(pCancellationReason) > 0 Then
        mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationReason).Value = pCancellationReason
        If Len(pCancelledBy) > 0 Then
          mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledBy).Value = pCancelledBy
        Else
          mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledBy).Value = mvEnv.User.UserID
        End If
        If Len(pCancelledOn) > 0 Then
          mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledOn).Value = pCancelledOn
        Else
          mvClassFields.Item(CreditCardAuthorityFields.ccafCancelledOn).Value = TodaysDate()
        End If
        If Len(pCancellationSource) > 0 Then mvClassFields.Item(CreditCardAuthorityFields.ccafCancellationSource).Value = pCancellationSource
      End If
    End Sub

    Public Sub InitForDataImport(ByVal pEnv As CDBEnvironment, Optional ByRef pCreditCardAuthorityNumber As Integer = 0)
      Init(pEnv)
      If pCreditCardAuthorityNumber > 0 Then mvClassFields.Item(CreditCardAuthorityFields.ccafCreditCardAuthorityNumber).Value = CStr(pCreditCardAuthorityNumber)
    End Sub

    Friend Sub SetFutureCancellation(ByVal pCancelReason As String, ByVal pCancelDate As String, Optional ByVal pCancelSource As String = "")
      mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationReason).Value = pCancelReason
      mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationDate).Value = pCancelDate
      mvClassFields.Item(CreditCardAuthorityFields.ccafFuturecancellationSource).Value = pCancelSource
    End Sub

    Friend Sub UnsetFutureCancellation()
      mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationReason).Value = ""
      mvClassFields.Item(CreditCardAuthorityFields.ccafFutureCancellationDate).Value = ""
      mvClassFields.Item(CreditCardAuthorityFields.ccafFuturecancellationSource).Value = ""
    End Sub
  End Class
End Namespace
