

Namespace Access
  Public Class StandingOrder
    Implements IAutoPaymentMethod

    Public Enum StandingOrderRecordSetTypes 'These are bit values
      sortAll = &HFFS
      'ADD additional recordset types here
      sortNumber = 1
      sortMain = 2
      sortBankInfo = 4
      sortCancelled = 8
      sortDetails = 16
      sortAmendedAlias = &H400S
      sortCancelledAlias = &H800S
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum StandingOrderFields
      sofAll = 0
      sofStandingOrderNumber
      sofContactNumber
      sofAddressNumber
      sofPaymentPlanNumber
      sofReference
      sofStartDate
      sofAmount
      sofBankAccount
      sofBankDetailsNumber
      sofCancellationReason
      sofCancelledOn
      sofCancelledBy
      sofAmendedBy
      sofAmendedOn
      sofSource
      sofStandingOrderType
      sofCancellationSource
      sofCreatedBy
      sofCreatedOn
      sofFutureCancellationReason
      sofFutureCancellationDate
      sofFutureCancellationSource
    End Enum

    Public Enum SOType
      sotBankSO 'Bank Standing Order
      sotCAFSO 'CAF Standing Order
    End Enum

    Private mvContactAccount As ContactAccount

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
          .DatabaseTableName = "bankers_orders"
          .Add("bankers_order_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("reference")
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("bank_account")
          .Add("bank_details_number", CDBField.FieldTypes.cftLong)
          .Add("cancellation_reason")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("cancelled_by")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("source")
          .Add("standing_order_type")
          .Add("cancellation_source")
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_reason")
          .Add("future_cancellation_date", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_source")

          .Item(StandingOrderFields.sofStandingOrderNumber).SetPrimaryKeyOnly()

          .Item(StandingOrderFields.sofReference).SpecialColumn = True

          .Item(StandingOrderFields.sofContactNumber).PrefixRequired = True
          .Item(StandingOrderFields.sofAddressNumber).PrefixRequired = True
          .Item(StandingOrderFields.sofPaymentPlanNumber).PrefixRequired = True
          .Item(StandingOrderFields.sofAmount).PrefixRequired = True
          .Item(StandingOrderFields.sofSource).PrefixRequired = True
          .Item(StandingOrderFields.sofCreatedBy).PrefixRequired = True
          .Item(StandingOrderFields.sofCreatedOn).PrefixRequired = True
          .Item(StandingOrderFields.sofCancellationReason).PrefixRequired = True
          .Item(StandingOrderFields.sofCancellationSource).PrefixRequired = True
          .Item(StandingOrderFields.sofCancelledBy).PrefixRequired = True
          .Item(StandingOrderFields.sofCancelledOn).PrefixRequired = True
          .Item(StandingOrderFields.sofFutureCancellationDate).PrefixRequired = True
          .Item(StandingOrderFields.sofFutureCancellationReason).PrefixRequired = True
          .Item(StandingOrderFields.sofFutureCancellationSource).PrefixRequired = True
          .Item(StandingOrderFields.sofAmendedBy).PrefixRequired = True
          .Item(StandingOrderFields.sofAmendedOn).PrefixRequired = True
        End With
        mvActivity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSOActivity)
        mvActivityValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSOActivityValue)
      Else
        mvClassFields.ClearItems()
      End If
      'UPGRADE_NOTE: Object mvContactAccount may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvContactAccount = Nothing
      mvContactType = Contact.ContactTypes.ctcContact 'Default
      mvAmendedValid = False
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As StandingOrderFields)

      'Add code here to ensure all values are valid before saving
      If pField = StandingOrderFields.sofAll And Not mvAmendedValid Then
        mvClassFields.Item(StandingOrderFields.sofAmendedOn).Value = TodaysDate()
        mvClassFields.Item(StandingOrderFields.sofAmendedBy).Value = mvEnv.User.UserID
      End If

      'Assumes the SO number has already been set
      If ((pField = StandingOrderFields.sofAll) Or ((pField And StandingOrderFields.sofReference) = StandingOrderFields.sofReference)) And Len(mvClassFields.Item(StandingOrderFields.sofReference).Value) = 0 Then
        mvClassFields.Item(StandingOrderFields.sofReference).Value = mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).Value
      End If

      If pField = StandingOrderFields.sofAll And Len(mvClassFields.Item(StandingOrderFields.sofCreatedBy).Value) = 0 And mvExisting = False Then
        mvClassFields.Item(StandingOrderFields.sofCreatedBy).Value = mvEnv.User.UserID
        mvClassFields.Item(StandingOrderFields.sofCreatedOn).Value = TodaysDate()
      End If

    End Sub

    Private Function GetSOType(ByRef pSOType As String) As SOType
      Select Case pSOType
        Case "C"
          GetSOType = SOType.sotCAFSO
        Case Else 'B (or null?)
          GetSOType = SOType.sotBankSO
      End Select
    End Function

    Public Function GetSOTypeCode(ByRef pSOType As SOType) As String
      Select Case pSOType
        Case SOType.sotBankSO
          GetSOTypeCode = "B"
        Case Else 'sotCAFSO
          GetSOTypeCode = "C"
      End Select
    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As StandingOrderRecordSetTypes) As String
      Dim vFields As String

      'Modify below to add each recordset type as required
      If pRSType = StandingOrderRecordSetTypes.sortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bo")
      Else
        vFields = "bo.bankers_order_number"
        If (pRSType And StandingOrderRecordSetTypes.sortCancelled) > 0 Then vFields = vFields & ",bo.cancellation_reason,bo.future_cancellation_reason,bo.future_cancellation_date,bo.future_cancellation_source"
        If (pRSType And StandingOrderRecordSetTypes.sortMain) > 0 Then vFields = vFields & ",contact_number,address_number,order_number"
        If (pRSType And StandingOrderRecordSetTypes.sortBankInfo) > 0 Then vFields = vFields & "," & mvEnv.Connection.DBSpecialCol("bo", "reference") & ",bank_account,bo.amount"
        If (pRSType And StandingOrderRecordSetTypes.sortAmendedAlias) > 0 Then vFields = vFields & ",bo.amended_on AS bo_amended_on, bo.amended_by AS bo_amended_by"
        If (pRSType And StandingOrderRecordSetTypes.sortCancelledAlias) > 0 Then vFields = vFields & ",bo.cancellation_reason AS bo_cancellation_reason, bo.cancelled_by AS bo_cancelled_by, bo.cancelled_on AS bo_cancelled_on"
        If (pRSType And StandingOrderRecordSetTypes.sortDetails) > 0 Then vFields = vFields & ",bo.source,bo.bank_details_number"
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pStandingOrderNumber As Integer = 0) Implements IAutoPaymentMethod.Init
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pStandingOrderNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(StandingOrderRecordSetTypes.sortAll) & " FROM bankers_orders bo WHERE bankers_order_number = " & pStandingOrderNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, StandingOrderRecordSetTypes.sortAll)
        Else
          System.Diagnostics.Debug.Assert(False, "") 'Standing Order not found!!!
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Friend Sub InitFromPaymentPlan(ByVal pEnv As CDBEnvironment, ByRef pPaymentPlanNumber As Integer, Optional ByRef pCancelled As Boolean = False)
      Dim vRecordSet As CDBRecordSet
      Dim vAndCancellation As String
      Dim vSOType As StandingOrderRecordSetTypes

      If pPaymentPlanNumber > 0 Then
        mvEnv = pEnv
        InitClassFields()
        If Not pCancelled Then
          vAndCancellation = " AND cancellation_reason IS NULL"
          vSOType = StandingOrderRecordSetTypes.sortAll
        Else
          vAndCancellation = " AND cancellation_reason IS NOT NULL ORDER BY cancelled_on DESC"
          vSOType = StandingOrderRecordSetTypes.sortAll
        End If
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(vSOType) & " FROM bankers_orders bo WHERE order_number = " & pPaymentPlanNumber & vAndCancellation)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, vSOType)
        Else
          Init(pEnv)
        End If
        vRecordSet.CloseRecordSet()
      Else
        Init(pEnv)
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As StandingOrderRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Modify below to handle each recordset type as required
        .SetItem(StandingOrderFields.sofStandingOrderNumber, vFields)
        If (pRSType And StandingOrderRecordSetTypes.sortMain) = StandingOrderRecordSetTypes.sortMain Then
          .SetItem(StandingOrderFields.sofContactNumber, vFields)
          .SetItem(StandingOrderFields.sofAddressNumber, vFields)
          .SetItem(StandingOrderFields.sofPaymentPlanNumber, vFields)
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortBankInfo) = StandingOrderRecordSetTypes.sortBankInfo Then
          .SetItem(StandingOrderFields.sofReference, vFields)
          .SetItem(StandingOrderFields.sofAmount, vFields)
          .SetItem(StandingOrderFields.sofBankAccount, vFields)
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortCancelled) = StandingOrderRecordSetTypes.sortCancelled Then
          .SetItem(StandingOrderFields.sofCancellationReason, vFields)
          .SetItem(StandingOrderFields.sofFutureCancellationReason, vFields)
          .SetItem(StandingOrderFields.sofFutureCancellationDate, vFields)
          .SetItem(StandingOrderFields.sofFutureCancellationSource, vFields)
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortAll) = StandingOrderRecordSetTypes.sortAll Then
          .SetItem(StandingOrderFields.sofStartDate, vFields)
          .SetItem(StandingOrderFields.sofBankDetailsNumber, vFields)
          '      .SetItem sofCancellationReason, vFields
          .SetItem(StandingOrderFields.sofCancelledOn, vFields)
          .SetItem(StandingOrderFields.sofCancelledBy, vFields)
          .SetItem(StandingOrderFields.sofAmendedBy, vFields)
          .SetItem(StandingOrderFields.sofAmendedOn, vFields)
          .SetItem(StandingOrderFields.sofSource, vFields)
          .SetItem(StandingOrderFields.sofCancellationSource, vFields)
          .SetItem(StandingOrderFields.sofCreatedBy, vFields)
          .SetItem(StandingOrderFields.sofCreatedOn, vFields)
          .SetItem(StandingOrderFields.sofStandingOrderType, vFields)
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortAmendedAlias) > 0 Then
          mvClassFields.Item(StandingOrderFields.sofAmendedBy).SetValue = vFields("bo_amended_by").Value
          mvClassFields.Item(StandingOrderFields.sofAmendedOn).SetValue = vFields("bo_amended_on").Value
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortCancelledAlias) > 0 Then
          mvClassFields.Item(StandingOrderFields.sofCancellationReason).SetValue = vFields("bo_cancellation_reason").Value
          mvClassFields.Item(StandingOrderFields.sofCancelledBy).SetValue = vFields("bo_cancelled_by").Value
          mvClassFields.Item(StandingOrderFields.sofCancelledOn).SetValue = vFields("bo_cancelled_on").Value
        End If
        If (pRSType And StandingOrderRecordSetTypes.sortDetails) = StandingOrderRecordSetTypes.sortDetails Then
          .SetItem(StandingOrderFields.sofSource, vFields)
          .SetItem(StandingOrderFields.sofBankDetailsNumber, vFields)
        End If
      End With
    End Sub
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0)
      Dim vTransaction As Boolean

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      'Must have the SO Number before calling SetValid as the Referance could be generated from the SO Number
      If Len(mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).Value) = 0 Then
        mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).Value = CStr(mvEnv.GetControlNumber("BO"))
      End If
      SetValid(StandingOrderFields.sofAll)

      If Not mvContactAccount Is Nothing Then
        'BR 8634: If user has changed Contact Account, reinitialise it.
        If mvContactAccount.BankDetailsNumber > 0 And BankDetailsNumber > 0 And mvContactAccount.BankDetailsNumber <> BankDetailsNumber Then
          mvContactAccount.Init(mvEnv, BankDetailsNumber)
        Else
          'Save Contact Accounts record
          mvContactAccount.Save()
        End If
        mvClassFields.Item(StandingOrderFields.sofBankDetailsNumber).Value = CStr(mvContactAccount.BankDetailsNumber)
      End If

      If mvExisting Then
        mvEnv.Connection.UpdateRecords("bankers_orders", mvClassFields.UpdateFields, mvClassFields.WhereFields)
        mvEnv.AddJournalRecord(JournalTypes.jnlStandingOrder, JournalOperations.jnlUpdate, mvClassFields.Item(StandingOrderFields.sofContactNumber).IntegerValue, mvClassFields.Item(StandingOrderFields.sofAddressNumber).IntegerValue, (mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
      Else
        mvEnv.Connection.InsertRecord("bankers_orders", mvClassFields.UpdateFields)
        mvEnv.AddJournalRecord(JournalTypes.jnlStandingOrder, JournalOperations.jnlInsert, mvClassFields.Item(StandingOrderFields.sofContactNumber).IntegerValue, mvClassFields.Item(StandingOrderFields.sofAddressNumber).IntegerValue, (mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
        If Len(mvClassFields(StandingOrderFields.sofCancellationReason).Value) > 0 Then
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
    Public Sub SaveChanges(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False) Implements IAutoPaymentMethod.SaveChanges
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub SetCancelled(ByRef pCancellationReason As String, Optional ByRef pCancelledOn As String = "", Optional ByRef pCancelledBy As String = "", Optional ByRef pCancellationSource As String = "")
      If Len(pCancellationReason) > 0 Then
        mvClassFields.Item(StandingOrderFields.sofCancellationReason).Value = pCancellationReason
        If Len(pCancelledBy) > 0 Then
          mvClassFields.Item(StandingOrderFields.sofCancelledBy).Value = pCancelledBy
        Else
          mvClassFields.Item(StandingOrderFields.sofCancelledBy).Value = mvEnv.User.UserID
        End If
        If Len(pCancelledOn) > 0 Then
          mvClassFields.Item(StandingOrderFields.sofCancelledOn).Value = pCancelledOn
        Else
          mvClassFields.Item(StandingOrderFields.sofCancelledOn).Value = TodaysDate()
        End If
        If Len(pCancellationSource) > 0 Then mvClassFields.Item(StandingOrderFields.sofCancellationSource).Value = pCancellationSource
      End If
    End Sub

    Public Sub SetUnCancelled() Implements IAutoPaymentMethod.SetUnCancelled
      mvClassFields.Item(StandingOrderFields.sofCancellationReason).Value = ""
      mvClassFields.Item(StandingOrderFields.sofCancelledOn).Value = ""
      mvClassFields.Item(StandingOrderFields.sofCancelledBy).Value = ""
      mvClassFields.Item(StandingOrderFields.sofCancellationSource).Value = ""
    End Sub

    Public Sub SetContact(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer)
      mvClassFields.Item(StandingOrderFields.sofContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(StandingOrderFields.sofAddressNumber).Value = CStr(pAddressNumber)
    End Sub

    Public Sub Create(ByRef pBankDetailsNumber As Integer, ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pPPNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pReference As String, ByRef pType As String, Optional ByRef pContactType As Contact.ContactTypes = Contact.ContactTypes.ctcContact, Optional ByRef pCreatedOn As String = "", Optional ByRef pCreatedBy As String = "")
      With mvClassFields
        .Item(StandingOrderFields.sofStandingOrderNumber).IntegerValue = mvEnv.GetControlNumber("BO")
        .Item(StandingOrderFields.sofBankDetailsNumber).IntegerValue = pBankDetailsNumber
        .Item(StandingOrderFields.sofContactNumber).IntegerValue = pContactNumber
        .Item(StandingOrderFields.sofAddressNumber).IntegerValue = pAddressNumber
        .Item(StandingOrderFields.sofPaymentPlanNumber).IntegerValue = pPPNumber
        .Item(StandingOrderFields.sofBankAccount).Value = pBankAccount
        .Item(StandingOrderFields.sofSource).Value = pSource
        .Item(StandingOrderFields.sofAmount).DoubleValue = pAmount
        .Item(StandingOrderFields.sofStartDate).Value = pStartDate
        If Len(pReference) = 0 Then pReference = "N" & StandingOrderNumber
        .Item(StandingOrderFields.sofReference).Value = pReference
        If Len(pType) = 0 Then pType = GetSOTypeCode(SOType.sotBankSO)
        .Item(StandingOrderFields.sofStandingOrderType).Value = pType
        .Item(StandingOrderFields.sofCreatedOn).Value = pCreatedOn
        .Item(StandingOrderFields.sofCreatedBy).Value = pCreatedBy
        mvContactType = pContactType
      End With
      SetValid(StandingOrderFields.sofAll)
    End Sub

    Public Sub Update(ByRef pBankDetailsNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pReference As String, ByRef pType As String)
      With mvClassFields
        .Item(StandingOrderFields.sofBankDetailsNumber).IntegerValue = pBankDetailsNumber
        .Item(StandingOrderFields.sofBankAccount).Value = pBankAccount
        .Item(StandingOrderFields.sofSource).Value = pSource
        .Item(StandingOrderFields.sofAmount).DoubleValue = pAmount
        .Item(StandingOrderFields.sofStartDate).Value = pStartDate
        .Item(StandingOrderFields.sofReference).Value = pReference
        If Len(pType) > 0 Then .Item(StandingOrderFields.sofStandingOrderType).Value = pType
      End With
    End Sub

    Public Sub SetPlanNumberAndAmount(ByRef pPlanNumber As Integer, ByRef pAmount As Double)
      With mvClassFields
        .Item(StandingOrderFields.sofPaymentPlanNumber).IntegerValue = pPlanNumber
        .Item(StandingOrderFields.sofAmount).DoubleValue = pAmount
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
        AddressNumber = mvClassFields.Item(StandingOrderFields.sofAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(StandingOrderFields.sofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(StandingOrderFields.sofAmendedOn).Value
      End Get
    End Property

    Public Property Amount() As Double
      Get
        Amount = mvClassFields.Item(StandingOrderFields.sofAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(StandingOrderFields.sofAmount).DoubleValue = Value
      End Set
    End Property

    Public Property BankAccount() As String Implements IAutoPaymentMethod.BankAccount
      Get
        BankAccount = mvClassFields.Item(StandingOrderFields.sofBankAccount).Value
      End Get
      Set(ByVal Value As String)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(StandingOrderFields.sofBankAccount).Value = Value
      End Set
    End Property

    Public ReadOnly Property BankDetailsNumber() As Integer
      Get
        BankDetailsNumber = mvClassFields.Item(StandingOrderFields.sofBankDetailsNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StandingOrderNumber() As Integer Implements IAutoPaymentMethod.AutoPaymentNumber
      Get
        StandingOrderNumber = mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(StandingOrderFields.sofCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(StandingOrderFields.sofCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(StandingOrderFields.sofCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(StandingOrderFields.sofCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactAccount() As ContactAccount
      Get
        If mvContactAccount Is Nothing Then
          mvContactAccount = New ContactAccount
          mvContactAccount.Init(mvEnv, (mvClassFields.Item(StandingOrderFields.sofBankDetailsNumber).IntegerValue))
        End If
        ContactAccount = mvContactAccount
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer Implements IAutoPaymentMethod.ContactNumber
      Get
        ContactNumber = mvClassFields.Item(StandingOrderFields.sofContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(StandingOrderFields.sofPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public Property Reference() As String
      Get
        Reference = mvClassFields.Item(StandingOrderFields.sofReference).Value
      End Get
      Set(ByVal Value As String)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(StandingOrderFields.sofReference).Value = Value
      End Set
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(StandingOrderFields.sofSource).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(StandingOrderFields.sofStartDate).Value
      End Get
    End Property

    Public ReadOnly Property StandingOrderType() As SOType
      Get
        StandingOrderType = GetSOType((mvClassFields.Item(StandingOrderFields.sofStandingOrderType).Value))
      End Get
    End Property

    Public ReadOnly Property StandingOrderTypeCode() As String
      Get
        StandingOrderTypeCode = mvClassFields.Item(StandingOrderFields.sofStandingOrderType).Value
      End Get
    End Property

    Public ReadOnly Property CreatedBy() As String
      Get
        CreatedBy = mvClassFields.Item(StandingOrderFields.sofCreatedBy).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(StandingOrderFields.sofCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationReason() As String
      Get
        FutureCancellationReason = mvClassFields.Item(StandingOrderFields.sofFutureCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationDate() As String
      Get
        FutureCancellationDate = mvClassFields.Item(StandingOrderFields.sofFutureCancellationDate).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationSource() As String
      Get
        FutureCancellationSource = mvClassFields.Item(StandingOrderFields.sofFutureCancellationSource).Value
      End Get
    End Property

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(StandingOrderFields.sofAmendedOn).Value = pAmendedOn
      mvClassFields.Item(StandingOrderFields.sofAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Sub InitForDataImport(ByVal pEnv As CDBEnvironment, Optional ByRef pSONumber As Integer = 0)
      Init(pEnv)
      If pSONumber > 0 Then mvClassFields.Item(StandingOrderFields.sofStandingOrderNumber).Value = CStr(pSONumber)
    End Sub

    Friend Sub SetFutureCancellation(ByVal pCancelReason As String, ByVal pCancelDate As String, Optional ByVal pCancelSource As String = "")
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationReason).Value = pCancelReason
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationDate).Value = pCancelDate
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationSource).Value = pCancelSource
    End Sub

    Friend Sub UnsetFutureCancellation()
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationReason).Value = ""
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationDate).Value = ""
      mvClassFields.Item(StandingOrderFields.sofFutureCancellationSource).Value = ""
    End Sub

  End Class
End Namespace
