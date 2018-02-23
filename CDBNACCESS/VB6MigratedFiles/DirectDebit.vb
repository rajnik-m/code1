

Namespace Access
  Public Class DirectDebit
    Implements IAutoPaymentMethod

    Public Enum DirectDebitRecordSetTypes 'These are bit values
      ddrtAll = &HFFFFS
      'ADD additional recordset types here
      ddrtNumber = 1
      ddrtDetail = 2
      ddrtCancelled = 4
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DirectDebitFields
      ddfAll = 0
      ddfDirectDebitNumber
      ddfContactNumber
      ddfAddressNumber
      ddfPaymentPlanNumber
      ddfStartDate
      ddfBankDetailsNumber
      ddfBankAccount
      ddfCancellationReason
      ddfCancelledOn
      ddfCancelledBy
      ddfAmendedOn
      ddfAmendedBy
      ddfAmount
      ddfSource
      ddfReference
      ddfFirstClaim
      ddfEmandateCreated
      ddfAuddisCancelNotified
      ddfCancellationSource
      ddfMandateType
      ddfCreatedBy
      ddfCreatedOn
      ddfFutureCancellationReason
      ddfFutureCancellationDate
      ddfFutureCancellationSource
      ddfDateSigned
      ddfBankDetailsChanged
      ddfPreviousBankDetailsNumber
    End Enum

    Private mvContactAccount As ContactAccount
    Private mvText(5) As String

    Private mvActivity As String
    Private mvActivityValue As String
    Private mvContactType As Contact.ContactTypes
    Private mvPayerContact As Contact

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAmendedValid As Boolean
    Private mvDDRefUpdated As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      Dim vIndex As Integer

      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "direct_debits"
          .Add("direct_debit_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("bank_details_number", CDBField.FieldTypes.cftLong)
          .Add("bank_account")
          .Add("cancellation_reason")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("cancelled_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("source")
          .Add("reference")
          .Add("first_claim")
          .Add("emandate_created", CDBField.FieldTypes.cftDate)
          .Add("auddis_cancel_notified", CDBField.FieldTypes.cftDate)
          .Add("cancellation_source")
          .Add("mandate_type")
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_reason")
          .Add("future_cancellation_date", CDBField.FieldTypes.cftDate)
          .Add("future_cancellation_source")
          .Add("date_signed", CDBField.FieldTypes.cftDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers)
          .Add("bank_details_changed").InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers)
          .Add("previous_bank_details_number", CDBField.FieldTypes.cftInteger).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers)

          .Item(DirectDebitFields.ddfDirectDebitNumber).SetPrimaryKeyOnly()

          .Item(DirectDebitFields.ddfContactNumber).PrefixRequired = True
          .Item(DirectDebitFields.ddfAddressNumber).PrefixRequired = True
          .Item(DirectDebitFields.ddfPaymentPlanNumber).PrefixRequired = True
          .Item(DirectDebitFields.ddfBankDetailsNumber).PrefixRequired = True
          .Item(DirectDebitFields.ddfBankAccount).PrefixRequired = True
          .Item(DirectDebitFields.ddfAmendedBy).PrefixRequired = True
          .Item(DirectDebitFields.ddfAmendedOn).PrefixRequired = True
          .Item(DirectDebitFields.ddfCancellationReason).PrefixRequired = True
          .Item(DirectDebitFields.ddfCancelledOn).PrefixRequired = True
          .Item(DirectDebitFields.ddfCancelledBy).PrefixRequired = True
          .Item(DirectDebitFields.ddfCancellationSource).PrefixRequired = True
          .Item(DirectDebitFields.ddfAmount).PrefixRequired = True
          .Item(DirectDebitFields.ddfSource).PrefixRequired = True
          .Item(DirectDebitFields.ddfCreatedBy).PrefixRequired = True
          .Item(DirectDebitFields.ddfCreatedOn).PrefixRequired = True
          .Item(DirectDebitFields.ddfFutureCancellationDate).PrefixRequired = True
          .Item(DirectDebitFields.ddfFutureCancellationReason).PrefixRequired = True
          .Item(DirectDebitFields.ddfFutureCancellationSource).PrefixRequired = True
        End With
        mvActivity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDDActivity)
        mvActivityValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDDActivityValue)
      Else
        mvClassFields.ClearItems()
      End If
      For vIndex = 1 To 5
        mvText(vIndex) = ""
      Next
      mvContactAccount = Nothing
      mvContactType = Contact.ContactTypes.ctcContact 'Default
      mvAmendedValid = False
      mvExisting = False
      mvDDRefUpdated = False
      mvPayerContact = New Contact(mvEnv)
      mvPayerContact.Init()
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(DirectDebitFields.ddfBankDetailsChanged).Bool = False
    End Sub

    Private Sub SetValid(ByRef pField As DirectDebitFields)
      'Add code here to ensure all values are valid before saving
      If pField = DirectDebitFields.ddfAll And Not mvAmendedValid Then
        mvClassFields.Item(DirectDebitFields.ddfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(DirectDebitFields.ddfAmendedBy).Value = mvEnv.User.UserID
      End If

      If pField = DirectDebitFields.ddfAll And Len(mvClassFields.Item(DirectDebitFields.ddfCreatedBy).Value) = 0 And mvExisting = False Then
        mvClassFields.Item(DirectDebitFields.ddfCreatedBy).Value = mvEnv.User.UserID
        mvClassFields.Item(DirectDebitFields.ddfCreatedOn).Value = TodaysDate()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Public Function GetClaimTransactionType(ByRef pDefaultTransType As String, ByVal pOneOffClaim As Boolean) As String
      If pOneOffClaim Then
        Return mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlOneOffClaimTransactionType)
      ElseIf FirstClaim Then
        Return mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlFirstClaimTransactionType)
      Else
        Return pDefaultTransType
      End If
    End Function

    Public Function GetRecordSetFields(ByVal pRSType As DirectDebitRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DirectDebitRecordSetTypes.ddrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dd")
      Else
        vFields = "direct_debit_number"
        If (pRSType And DirectDebitRecordSetTypes.ddrtNumber) > 0 Then vFields = vFields & ",dd.contact_number,dd.address_number,dd.bank_details_number"
        If (pRSType And DirectDebitRecordSetTypes.ddrtDetail) > 0 Then vFields = vFields & ",dd.source,dd.reference,first_claim,bank_details_changed,previous_bank_details_number"
        If (pRSType And DirectDebitRecordSetTypes.ddrtCancelled) > 0 Then vFields = vFields & ",dd.cancellation_reason,dd.future_cancellation_reason,dd.future_cancellation_date,dd.future_cancellation_source"
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pDirectDebitNumber As Integer = 0) Implements IAutoPaymentMethod.Init
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pDirectDebitNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DirectDebitRecordSetTypes.ddrtAll) & " FROM direct_debits dd WHERE direct_debit_number = " & pDirectDebitNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DirectDebitRecordSetTypes.ddrtAll)
        Else
          System.Diagnostics.Debug.Assert(False, "") 'Direct Debit not found!!!
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Friend Sub InitFromPaymentPlan(ByVal pEnv As CDBEnvironment, ByRef pPaymentPlanNumber As Integer, Optional ByRef pCancelled As Boolean = False)
      If pPaymentPlanNumber > 0 Then
        mvEnv = pEnv
        InitClassFields()

        mvPayerContact = New Contact(mvEnv)
        mvPayerContact.Init()

        Dim vOrderBy As String = String.Empty
        Dim vWhereFields As New CDBFields(New CDBField("order_number", pPaymentPlanNumber))
        If Not pCancelled Then
          vWhereFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter, "")
        Else
          vWhereFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
          vOrderBy = "cancelled_on DESC, direct_debit_number DESC"
        End If
        vWhereFields.AddJoin("dd.address_number", "a.address_number")

        Dim vAttrs As String = GetRecordSetFields(DirectDebitRecordSetTypes.ddrtAll)
        vAttrs &= "," & mvPayerContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtVAT Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber)
        vAttrs = vAttrs.Replace("c.contact_number,", "").Replace("a.address_number,", "").Replace("c.source,", "")
        vAttrs = vAttrs.Replace("c.amended_by,", "").Replace("c.amended_on,", "")

        Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("contacts c", "dd.contact_number", "c.contact_number")})
        vAnsiJoins.Add("contact_addresses ca", "c.contact_number", "ca.contact_number")
        vAnsiJoins.Add("addresses a", "ca.address_number", "a.address_number")
        vAnsiJoins.Add("countries co", "a.country", "co.country")

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "direct_debits dd", vWhereFields, vOrderBy, vAnsiJoins)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        If vRS.Fetch Then
          InitFromRecordSet(pEnv, vRS, DirectDebitRecordSetTypes.ddrtAll)
          mvPayerContact.InitFromRecordSet(pEnv, vRS, Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtVAT Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber)
        End If
        vRS.CloseRecordSet()
      Else
        Init(pEnv)
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DirectDebitRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always grab the unique key, 'cos you need it for saving
        .SetItem(DirectDebitFields.ddfDirectDebitNumber, vFields)
        If (pRSType And DirectDebitRecordSetTypes.ddrtNumber) > 0 Then
          .SetItem(DirectDebitFields.ddfContactNumber, vFields)
          .SetItem(DirectDebitFields.ddfAddressNumber, vFields)
          .SetItem(DirectDebitFields.ddfBankDetailsNumber, vFields)
        End If
        If (pRSType And DirectDebitRecordSetTypes.ddrtDetail) > 0 Then
          .SetItem(DirectDebitFields.ddfSource, vFields)
          .SetItem(DirectDebitFields.ddfReference, vFields)
          .SetItem(DirectDebitFields.ddfFirstClaim, vFields)
          .SetOptionalItem(DirectDebitFields.ddfBankDetailsChanged, vFields)
          .SetOptionalItem(DirectDebitFields.ddfPreviousBankDetailsNumber, vFields)
        End If
        If (pRSType And DirectDebitRecordSetTypes.ddrtCancelled) > 0 Then
          .SetItem(DirectDebitFields.ddfCancellationReason, vFields)
          .SetOptionalItem(DirectDebitFields.ddfFutureCancellationReason, vFields)
          .SetOptionalItem(DirectDebitFields.ddfFutureCancellationDate, vFields)
          .SetOptionalItem(DirectDebitFields.ddfFutureCancellationSource, vFields)
        End If
        If (pRSType And DirectDebitRecordSetTypes.ddrtAll) = DirectDebitRecordSetTypes.ddrtAll Then
          .SetItem(DirectDebitFields.ddfPaymentPlanNumber, vFields)
          .SetItem(DirectDebitFields.ddfStartDate, vFields)
          .SetItem(DirectDebitFields.ddfBankAccount, vFields)
          .SetItem(DirectDebitFields.ddfCancelledOn, vFields)
          .SetItem(DirectDebitFields.ddfCancelledBy, vFields)
          .SetItem(DirectDebitFields.ddfAmendedOn, vFields)
          .SetItem(DirectDebitFields.ddfAmendedBy, vFields)
          .SetItem(DirectDebitFields.ddfAmount, vFields)
          .SetItem(DirectDebitFields.ddfEmandateCreated, vFields)
          .SetOptionalItem(DirectDebitFields.ddfAuddisCancelNotified, vFields)
          .SetOptionalItem(DirectDebitFields.ddfCancellationSource, vFields)
          .SetOptionalItem(DirectDebitFields.ddfMandateType, vFields)
          .SetOptionalItem(DirectDebitFields.ddfCreatedBy, vFields)
          .SetOptionalItem(DirectDebitFields.ddfCreatedOn, vFields)
          .SetOptionalItem(DirectDebitFields.ddfDateSigned, vFields)
        End If
      End With
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

    Public Sub Save(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pNoTransaction As Boolean = False)
      Dim vDDReferencesFields As New CDBFields
      Dim vDDRefUpdateFields As New CDBFields
      Dim vTransaction As Boolean

      If Not pNoTransaction AndAlso Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      'Must have the DD Number before calling SetValid as the Referance could be generated from the DD Number
      If Len(mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).Value) = 0 Then
        mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).Value = CStr(mvEnv.GetControlNumber("DD"))
      End If
      SetValid(DirectDebitFields.ddfAll)

      If Not mvContactAccount Is Nothing Then 'Save Contact Accounts record
        'BR 8634: If user has changed Contact Account, reinitialise it.
        If mvContactAccount.BankDetailsNumber > 0 And mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue > 0 And mvContactAccount.BankDetailsNumber <> mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue Then
          mvContactAccount.Init(mvEnv, (mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue))
        End If
        mvContactAccount.Save()
        mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).Value = CStr(mvContactAccount.BankDetailsNumber)
      End If

      If mvExisting Then
        mvEnv.Connection.UpdateRecords("direct_debits", mvClassFields.UpdateFields, mvClassFields.WhereFields)
        Dim vJournalNumber As Integer = mvEnv.AddJournalRecord(JournalTypes.jnlDirectDebit, JournalOperations.jnlUpdate, mvClassFields.Item(DirectDebitFields.ddfContactNumber).IntegerValue, mvClassFields.Item(DirectDebitFields.ddfAddressNumber).IntegerValue, (mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
        If pAudit Then mvEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audUpdate, "direct_debits", DirectDebitNumber, 0, pAmendedBy, mvClassFields, vJournalNumber)

        If mvDDRefUpdated = True And (mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL") Then
          vDDRefUpdateFields.Add("direct_debit_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue)
          vDDReferencesFields.Add("text1", CDBField.FieldTypes.cftCharacter, mvText(1))
          vDDReferencesFields.Add("text2", CDBField.FieldTypes.cftCharacter, mvText(2))
          vDDReferencesFields.Add("text3", CDBField.FieldTypes.cftCharacter, mvText(3))
          vDDReferencesFields.Add("text4", CDBField.FieldTypes.cftCharacter, mvText(4))
          vDDReferencesFields.Add("text5", CDBField.FieldTypes.cftCharacter, mvText(5))
          mvEnv.Connection.UpdateRecords("direct_debit_references", vDDReferencesFields, vDDRefUpdateFields)
        End If
      Else
        mvEnv.Connection.InsertRecord("direct_debits", mvClassFields.UpdateFields)
        Dim vJournalNumber As Integer = mvEnv.AddJournalRecord(JournalTypes.jnlDirectDebit, JournalOperations.jnlInsert, mvClassFields.Item(DirectDebitFields.ddfContactNumber).IntegerValue, mvClassFields.Item(DirectDebitFields.ddfAddressNumber).IntegerValue, (mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
        If pAudit Then mvEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audInsert, "direct_debits", DirectDebitNumber, 0, pAmendedBy, mvClassFields, vJournalNumber)
        If Len(mvClassFields(DirectDebitFields.ddfCancellationReason).Value) > 0 Then
          CreateCancelledAutoPMActivity()
        Else
          CreateAutoPMActivity()
        End If

        If (mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL") Then
          vDDReferencesFields.Add("direct_debit_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue)
          vDDReferencesFields.Add("text1", CDBField.FieldTypes.cftCharacter, mvText(1))
          vDDReferencesFields.Add("text2", CDBField.FieldTypes.cftCharacter, mvText(2))
          vDDReferencesFields.Add("text3", CDBField.FieldTypes.cftCharacter, mvText(3))
          vDDReferencesFields.Add("text4", CDBField.FieldTypes.cftCharacter, mvText(4))
          vDDReferencesFields.Add("text5", CDBField.FieldTypes.cftCharacter, mvText(5))
          mvEnv.Connection.InsertRecord("direct_debit_references", vDDReferencesFields)
        End If
      End If
      mvDDRefUpdated = False
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SetCancelled(ByRef pCancellationReason As String, Optional ByRef pCancelledOn As String = "", Optional ByRef pCancelledBy As String = "", Optional ByRef pCancellationSource As String = "")
      If Len(pCancellationReason) > 0 Then
        mvClassFields.Item(DirectDebitFields.ddfCancellationReason).Value = pCancellationReason
        If Len(pCancelledBy) > 0 Then
          mvClassFields.Item(DirectDebitFields.ddfCancelledBy).Value = pCancelledBy
        Else
          mvClassFields.Item(DirectDebitFields.ddfCancelledBy).Value = mvEnv.User.UserID
        End If
        If Len(pCancelledOn) > 0 Then
          mvClassFields.Item(DirectDebitFields.ddfCancelledOn).Value = pCancelledOn
        Else
          mvClassFields.Item(DirectDebitFields.ddfCancelledOn).Value = TodaysDate()
        End If
        If Len(pCancellationSource) > 0 Then mvClassFields.Item(DirectDebitFields.ddfCancellationSource).Value = pCancellationSource
      End If
    End Sub

    Public Sub SetPlanNumberAndAmount(ByVal pPlanNumber As Integer, ByVal pAmount As Nullable(Of Double))
      mvClassFields.Item(DirectDebitFields.ddfPaymentPlanNumber).IntegerValue = pPlanNumber
      If pAmount.HasValue Then
        mvClassFields.Item(DirectDebitFields.ddfAmount).DoubleValue = pAmount.Value
      End If
    End Sub

    Public Sub SetUnCancelled() Implements IAutoPaymentMethod.SetUnCancelled
      With mvClassFields
        .Item(DirectDebitFields.ddfCancellationReason).Value = ""
        .Item(DirectDebitFields.ddfCancelledOn).Value = ""
        .Item(DirectDebitFields.ddfCancelledBy).Value = ""
        .Item(DirectDebitFields.ddfCancellationSource).Value = ""
        If .Item(DirectDebitFields.ddfAuddisCancelNotified).Value.Length > 0 Then
          'Only need to re-notify AUDDIS if cancel notification sent
          .Item(DirectDebitFields.ddfEmandateCreated).Value = ""
        End If
        .Item(DirectDebitFields.ddfAuddisCancelNotified).Value = ""
      End With
      ValidateReference(Reference)
    End Sub

    Public Sub SaveChanges(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False) Implements IAutoPaymentMethod.SaveChanges
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pBankDetailsNumber As Integer, ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pPPNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pReference As String, ByRef pMandateType As String, ByRef pFirstClaim As Boolean, ByRef pEmandateDate As String, Optional ByRef pContactType As Contact.ContactTypes = Contact.ContactTypes.ctcContact, Optional ByRef pCreatedOn As String = "", Optional ByRef pCreatedBy As String = "", Optional ByVal pDateSigned As String = "")
      Dim vNumber As String
      With mvClassFields
        .Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue = mvEnv.GetControlNumber("DD")
        vNumber = .Item(DirectDebitFields.ddfDirectDebitNumber).Value
        .Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue = pBankDetailsNumber
        .Item(DirectDebitFields.ddfContactNumber).IntegerValue = pContactNumber
        .Item(DirectDebitFields.ddfAddressNumber).IntegerValue = pAddressNumber
        .Item(DirectDebitFields.ddfPaymentPlanNumber).IntegerValue = pPPNumber
        .Item(DirectDebitFields.ddfBankAccount).Value = pBankAccount
        .Item(DirectDebitFields.ddfSource).Value = pSource
        If pAmount > 0 Then .Item(DirectDebitFields.ddfAmount).DoubleValue = pAmount
        .Item(DirectDebitFields.ddfStartDate).Value = pStartDate
        .Item(DirectDebitFields.ddfMandateType).Value = pMandateType
        .Item(DirectDebitFields.ddfEmandateCreated).Value = pEmandateDate
        .Item(DirectDebitFields.ddfFirstClaim).Bool = pFirstClaim
        If pReference.Length = 0 And vNumber = New String(vNumber(0), vNumber.Length) Then
          'BACS will not allow a DD number with all digits the same!
          .Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue = mvEnv.GetControlNumber("DD")
          vNumber = .Item(DirectDebitFields.ddfDirectDebitNumber).Value
        End If
        .Item(DirectDebitFields.ddfReference).Value = FormatReference(pReference)
        .Item(DirectDebitFields.ddfCreatedOn).Value = pCreatedOn
        .Item(DirectDebitFields.ddfCreatedBy).Value = pCreatedBy
        If Not String.IsNullOrWhiteSpace(pDateSigned) AndAlso IsDate(pDateSigned) Then .Item(DirectDebitFields.ddfDateSigned).Value = pDateSigned
        mvContactType = pContactType
      End With
      SetValid(DirectDebitFields.ddfAll)
    End Sub
    Private Sub CheckReferenceValidity(ByVal pReference As String)
      Dim vAlphas As String = ""
      Dim vIndex As Integer
      If Len(pReference) > 0 Then
        If Len(pReference) <= 3 And (mvEnv.GetConfig("fp_dd_reference_format") = "PREFIX" Or mvEnv.GetConfig("fp_dd_reference_format") = "SUFFIX") Then
          'Valid if contains at least one alpha
          pReference = Trim(UCase(pReference))
          For vIndex = 1 To Len(pReference)
            Select Case Mid(pReference, vIndex, 1)
              Case "A" To "Z"
                vAlphas = vAlphas & Mid(pReference, vIndex, 1)
            End Select
          Next
          If Len(vAlphas) = 0 Then
            If mvEnv.GetConfig("fp_dd_reference_format") = "PREFIX" Then
              RaiseError(DataAccessErrors.daeDDPrefixNoAlpha)
            Else
              RaiseError(DataAccessErrors.daeDDSuffixNoAlpha)
            End If
          End If
        Else
          pReference = Trim(UCase(pReference))
          '1.strip out non-alpha chars
          For vIndex = 1 To Len(pReference)
            Select Case Mid(pReference, vIndex, 1)
              Case "0" To "9", "A" To "Z"
                vAlphas = vAlphas & Mid(pReference, vIndex, 1)
              Case Else
            End Select
          Next
          '2.check that you have at least 6 alpha chars
          If Len(vAlphas) < 6 Then
            RaiseError(DataAccessErrors.daeDDReferenceSixAlphas)
          Else
            '3.ensure that not all the alpha chars are the same char
            If vAlphas = New String(vAlphas(0), vAlphas.Length) Then
              RaiseError(DataAccessErrors.daeDDReferenceSameCharacters)
            End If
          End If
        End If
      End If
    End Sub

    Public Function FormatReference(ByRef pReference As String) As String
      Dim vNumber As String
      Dim vLen As Integer

      CheckReferenceValidity((pReference))
      vNumber = mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).Value
      If Len(pReference) > 0 And Len(pReference) <= 3 Then
        'Contains Prefix/Suffix
        vLen = Len(pReference & vNumber)
        If mvEnv.GetConfig("fp_dd_reference_format") = "PREFIX" Then
          If vLen >= 6 Then
            pReference = pReference & vNumber
          Else
            pReference = pReference & New String("0"c, 6 - vLen) & vNumber
          End If
        ElseIf mvEnv.GetConfig("fp_dd_reference_format") = "SUFFIX" Then
          If vLen >= 6 Then
            pReference = vNumber & pReference
          Else
            pReference = New String("0"c, 6 - vLen) & vNumber & pReference
          End If
        End If
      ElseIf Len(pReference) = 0 Then
        pReference = mvEnv.GetConfig("fp_dd_reference_prefix")
        If InStr(pReference, "CONTACTNUMBER") > 0 Then pReference = Replace(pReference, "CONTACTNUMBER", CStr(ContactNumber))
        If Len(pReference & vNumber) >= 6 Then
          pReference = pReference & vNumber
        Else
          pReference = pReference & New String("0"c, 6 - Len(pReference & vNumber)) & vNumber
        End If
      End If
      If Len(pReference) > 18 Then RaiseError(DataAccessErrors.daeDDReferenceTooLong)
      ValidateReference(pReference)

      Return pReference

    End Function

    Public Sub Update(ByVal pParams As CDBParameters)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False AndAlso vClassField.InDatabase Then
          If pParams.Exists(vClassField.ProperName) Then
            Select Case vClassField.ProperName
              Case "Amount"
                vClassField.Value = If(pParams("Amount").DoubleValue > 0, pParams("Amount").Value, "")    'Could be changing from a figure to null
              Case "BankDetailsChanged", "PreviousBankDetailsNumber", "ContactNumber", "AddressNumber"
                'Do nothing
              Case "Reference"
                vClassField.Value = FormatReference(pParams("Reference").Value)
              Case Else
                vClassField.Value = pParams(vClassField.ProperName).Value
            End Select
          End If
        End If
      Next
      With mvClassFields
        If .Item(DirectDebitFields.ddfBankDetailsNumber).ValueChanged Then mvContactAccount = Nothing 'Force new data to be selected if required
        If pParams IsNot Nothing AndAlso pParams.ContainsKey("CreationNotified") Then
          .Item(DirectDebitFields.ddfEmandateCreated).Value = pParams("CreationNotified").Value
        End If
        Dim vOldIBANNumber As String = pParams.ParameterExists("OldIbanNumber").Value
        Dim vPreviousBankDetailsNumber As Integer = IntegerValue(.Item(DirectDebitFields.ddfBankDetailsNumber).SetValue)
        Dim vContactAccount As ContactAccount = New ContactAccount()
        vContactAccount.Init(mvEnv)
        If PreviousBankDetailsNumber > 0 Then
          'Bank Details have previously been changed and no claim done just yet for the new details
          If BankDetailsNumber = PreviousBankDetailsNumber Then
            'Bank details have been reverted back to original value, so clear the old details
            .Item(DirectDebitFields.ddfFirstClaim).Bool = False
            .Item(DirectDebitFields.ddfBankDetailsChanged).Bool = False
            .Item(DirectDebitFields.ddfPreviousBankDetailsNumber).Value = String.Empty
            vOldIBANNumber = String.Empty
            vContactAccount.Init(mvEnv, BankDetailsNumber)
          Else
            'Bank details have been changed again, but they are different to the previous details
            'So compare current IBAN with the original one (before the previous change) rather than the current one
            vContactAccount.Init(mvEnv, PreviousBankDetailsNumber)
            vOldIBANNumber = vContactAccount.IbanNumber
            vPreviousBankDetailsNumber = vContactAccount.BankDetailsNumber
            .Item(DirectDebitFields.ddfFirstClaim).Bool = False   'Reset back to False as we now don't know what it should be
            'Re-set vContactAccount to point to new record
            vContactAccount = New ContactAccount
            vContactAccount.Init(mvEnv)
          End If
        End If

        If FirstClaim = False Then
          If vContactAccount.Existing = False Then vContactAccount.Init(mvEnv, BankDetailsNumber)
          Dim vBankChanged As ContactAccount.IbanBankChanged = vContactAccount.HasIbanBankChanged(vOldIBANNumber)
          Select Case vBankChanged
            Case ContactAccount.IbanBankChanged.BankChanged, ContactAccount.IbanBankChanged.SameIbanBank
              'The IBAN number has changed
              If vBankChanged = ContactAccount.IbanBankChanged.BankChanged Then .Item(DirectDebitFields.ddfFirstClaim).Bool = True
              'As the IBAN number has changed, always set BankDetailsChanged and PreviousBankDetailsNumber
              .Item(DirectDebitFields.ddfBankDetailsChanged).Bool = True
              .Item(DirectDebitFields.ddfPreviousBankDetailsNumber).IntegerValue = vPreviousBankDetailsNumber

            Case Else
              'ContactAccount.IbanBankChanged.NotChecked, ContactAccount.IbanBankChanged.NoIbanNumber, ContactAccount.IbanBankChanged.SameIbanNumber
              'Ensure BankDetailsChanged and PreviousBankDetailsNumber are not set
              .Item(DirectDebitFields.ddfBankDetailsChanged).Bool = False
              .Item(DirectDebitFields.ddfPreviousBankDetailsNumber).Value = String.Empty
          End Select
        End If
      End With
    End Sub
    Public Sub Update(ByRef pBankDetailsNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pReference As String, ByRef pMandateType As String)
      Update(pBankDetailsNumber, pBankAccount, pSource, pAmount, pStartDate, pReference, pMandateType, EmandateCreated)
    End Sub

    Public Sub Update(ByRef pBankDetailsNumber As Integer, ByRef pBankAccount As String, ByRef pSource As String, ByRef pAmount As Double, ByRef pStartDate As String, ByRef pReference As String, ByRef pMandateType As String, ByRef pCreationNotified As String)
      With mvClassFields
        .Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue = pBankDetailsNumber
        .Item(DirectDebitFields.ddfBankAccount).Value = pBankAccount
        .Item(DirectDebitFields.ddfSource).Value = pSource
        If pAmount > 0 Then
          .Item(DirectDebitFields.ddfAmount).DoubleValue = pAmount
        Else
          .Item(DirectDebitFields.ddfAmount).Value = "" 'Could be changing from a figure to null
        End If
        .Item(DirectDebitFields.ddfStartDate).Value = pStartDate
        .Item(DirectDebitFields.ddfMandateType).Value = pMandateType
        .Item(DirectDebitFields.ddfReference).Value = FormatReference(pReference)
        If EmandateCreated.Length > 0 Then .Item(DirectDebitFields.ddfEmandateCreated).Value = pCreationNotified
      End With
    End Sub

    Public Sub SetDirectDebitReferences(ByVal pParams As CDBParameters)
      If (mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL") Then
        mvText(1) = pParams("Text1").Value
        mvText(2) = pParams.ParameterExists("Text2").Value
        mvText(3) = pParams.ParameterExists("Text3").Value
        mvText(4) = pParams.ParameterExists("Text4").Value
        mvText(5) = pParams.ParameterExists("Text5").Value
        mvDDRefUpdated = True
      End If
    End Sub

    ''' <summary>Validate the Direct Debit Reference to ensure it is not being used by another live Direct Debit.</summary>
    Private Sub ValidateReference(ByVal pReference As String)
      If mvEnv.GetConfigOption("fp_validate_dd_reference", False) Then
        If CancellationReason.Length = 0 AndAlso pReference.Length > 0 Then
          Dim vWherefields As New CDBFields(New CDBField("reference", pReference))
          vWherefields.Add("cancellation_reason", "")
          If mvExisting Then vWherefields.Add("direct_debit_number", DirectDebitNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          Dim vCount As Integer = mvEnv.Connection.GetCount("direct_debits", vWherefields)
          If vCount > 0 Then
            If mvClassFields.Item(DirectDebitFields.ddfCancellationReason).ValueChanged Then
              'Re-instatement
              RaiseError(DataAccessErrors.daeDirectDebitRefNotUniqueCannotReinstate, pReference)   'Direct Debit cannot be reinstated as reference '%1' is being used by another live Direct Debit.
            Else
              RaiseError(DataAccessErrors.daeDirectDebitReferenceNotUnique, pReference)   'Direct Debit reference '%1' is already being used by another live Direct Debit.
            End If
          End If
        End If
      End If
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
        AddressNumber = mvClassFields.Item(DirectDebitFields.ddfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(DirectDebitFields.ddfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(DirectDebitFields.ddfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(DirectDebitFields.ddfAmount).Value
      End Get
    End Property

    Public Property BankAccount() As String Implements IAutoPaymentMethod.BankAccount
      Get
        BankAccount = mvClassFields.Item(DirectDebitFields.ddfBankAccount).Value
      End Get
      Set(ByVal value As String)
        mvClassFields.Item(DirectDebitFields.ddfBankAccount).Value = value
      End Set
    End Property

    Public Property BankDetailsNumber() As Integer
      Get
        BankDetailsNumber = mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        'Used by BACS Rejections
        mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(DirectDebitFields.ddfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(DirectDebitFields.ddfCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(DirectDebitFields.ddfCancelledBy).Value
      End Get
    End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(DirectDebitFields.ddfCancelledOn).Value
      End Get
    End Property

    Public Property ContactAccount() As ContactAccount
      Get
        If mvContactAccount Is Nothing Then
          mvContactAccount = New ContactAccount
          mvContactAccount.Init(mvEnv, (mvClassFields.Item(DirectDebitFields.ddfBankDetailsNumber).IntegerValue))
        End If
        ContactAccount = mvContactAccount
      End Get
      Set(ByVal Value As ContactAccount)
        mvContactAccount = Value
      End Set
    End Property

    Public ReadOnly Property ContactNumber() As Integer Implements IAutoPaymentMethod.ContactNumber
      Get
        ContactNumber = mvClassFields.Item(DirectDebitFields.ddfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DirectDebitNumber() As Integer Implements IAutoPaymentMethod.AutoPaymentNumber
      Get
        DirectDebitNumber = mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).IntegerValue
      End Get
    End Property

    Public Property EmandateCreated() As String
      Get
        EmandateCreated = mvClassFields.Item(DirectDebitFields.ddfEmandateCreated).Value
      End Get
      Set(ByVal Value As String)
        'Cleared by BACS Rejections
        mvClassFields.Item(DirectDebitFields.ddfEmandateCreated).Value = Value
      End Set
    End Property

    Public Property FirstClaim() As Boolean
      Get
        FirstClaim = mvClassFields.Item(DirectDebitFields.ddfFirstClaim).Bool
      End Get
      Set(ByVal Value As Boolean)
        'Used by Direct Debit Run
        mvClassFields.Item(DirectDebitFields.ddfFirstClaim).Bool = Value
      End Set
    End Property

    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(DirectDebitFields.ddfPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Reference() As String
      Get
        Reference = mvClassFields.Item(DirectDebitFields.ddfReference).Value
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(DirectDebitFields.ddfSource).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(DirectDebitFields.ddfStartDate).Value
      End Get
    End Property

    Public Property Text(ByVal pIndex As Integer) As String
      Get
        Text = mvText(pIndex)
      End Get
      Set(ByVal Value As String)
        If mvEnv.DefaultCountry = "CH" Then mvText(pIndex) = Value
      End Set
    End Property

    Public Property AuddisCancelNotified() As String
      Get
        AuddisCancelNotified = mvClassFields.Item(DirectDebitFields.ddfAuddisCancelNotified).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(DirectDebitFields.ddfAuddisCancelNotified).Value = Value
      End Set
    End Property

    Public ReadOnly Property MandateType() As String
      Get
        MandateType = mvClassFields.Item(DirectDebitFields.ddfMandateType).Value
      End Get
    End Property

    Public ReadOnly Property CreatedBy() As String
      Get
        CreatedBy = mvClassFields.Item(DirectDebitFields.ddfCreatedBy).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(DirectDebitFields.ddfCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationReason() As String
      Get
        FutureCancellationReason = mvClassFields.Item(DirectDebitFields.ddfFutureCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationDate() As String
      Get
        FutureCancellationDate = mvClassFields.Item(DirectDebitFields.ddfFutureCancellationDate).Value
      End Get
    End Property

    Public ReadOnly Property FutureCancellationSource() As String
      Get
        FutureCancellationSource = mvClassFields.Item(DirectDebitFields.ddfFutureCancellationSource).Value
      End Get
    End Property

    ''' <summary>Gets the date the Direct Debit mandate was signed.</summary>
    Public ReadOnly Property DateSigned() As String
      Get
        Return mvClassFields.Item(DirectDebitFields.ddfDateSigned).Value
      End Get
    End Property

    ''' <summary>Gets a boolean flag indicating whether the bank details have changed (within the same Bank) since the last Direct Debit claim.</summary>
    ''' <remarks>Used in combination with <see cref="PreviousBankDetailsNumber">PreviousBankDetailsNumber</see>.</remarks>
    Public ReadOnly Property BankDetailsChanged() As Boolean
      Get
        Return mvClassFields.Item(DirectDebitFields.ddfBankDetailsChanged).Bool
      End Get
    End Property

    ''' <summary>Gets the bank details number of the previous bank account when the bank details have changed (within the same bank) since the last Direct Debit claim.</summary>
    ''' <remarks>Used in combination with <see cref="BankDetailsChanged">BankDetailsChanged</see>.</remarks>
    Public ReadOnly Property PreviousBankDetailsNumber() As Integer
      Get
        Return mvClassFields.Item(DirectDebitFields.ddfPreviousBankDetailsNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Payer() As Contact
      Get
        If mvPayerContact Is Nothing OrElse mvPayerContact.Existing = False Then
          mvPayerContact = New Contact(mvEnv)
          mvPayerContact.Init(ContactNumber, AddressNumber)
        End If
        Return mvPayerContact
      End Get
    End Property

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(DirectDebitFields.ddfAmendedOn).Value = pAmendedOn
      mvClassFields.Item(DirectDebitFields.ddfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Sub InitForDataImport(ByVal pEnv As CDBEnvironment, Optional ByRef pDirectDebitNumber As Integer = 0)
      Init(pEnv)
      If pDirectDebitNumber > 0 Then mvClassFields.Item(DirectDebitFields.ddfDirectDebitNumber).Value = CStr(pDirectDebitNumber)
    End Sub

    Friend Sub SetFutureCancellation(ByVal pCancelReason As String, ByVal pCancelDate As String, Optional ByVal pCancelSource As String = "")
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationReason).Value = pCancelReason
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationDate).Value = pCancelDate
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationSource).Value = pCancelSource
    End Sub

    Friend Sub UnsetFutureCancellation()
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationReason).Value = ""
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationDate).Value = ""
      mvClassFields.Item(DirectDebitFields.ddfFutureCancellationSource).Value = ""
    End Sub

    ''' <summary>Change the payer contact and address of this Direct Debit.</summary>
    ''' <param name="pNewPayer"><see cref="Contact">Contact</see> that will become the new payer.  Muast have been initialised to the required contact and address numbers.</param>
    ''' <remarks>This will only change the payer contact and address numbers, it will NOT change the bank details.</remarks>
    Friend Sub ChangePayer(ByVal pNewPayer As Contact)
      If pNewPayer IsNot Nothing AndAlso pNewPayer.Existing = True AndAlso pNewPayer.Address.AddressNumber > 0 Then
        mvClassFields.Item(DirectDebitFields.ddfContactNumber).IntegerValue = pNewPayer.ContactNumber
        mvClassFields.Item(DirectDebitFields.ddfAddressNumber).IntegerValue = pNewPayer.Address.AddressNumber
        mvPayerContact = pNewPayer
      End If
    End Sub
  End Class
End Namespace
