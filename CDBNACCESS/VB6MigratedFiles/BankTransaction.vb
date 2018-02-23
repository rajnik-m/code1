Namespace Access
  Public Class BankTransaction

    Public Enum BankTransactionRecordSetTypes 'These are bit values
      bktrtAll = &HFFFFS
      bktrtReconcile = 1
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BankTransactionFields
      btfAll = 0
      btfStatementDate
      btfReconciledStatus
      btfUnreconciledReason
      btfBankAccount
      btfSortCode
      btfAccountNumber
      btfAccountType
      btfTransactionCode
      btfPayersSortCode
      btfPayersAccountNumber
      btfPayersReference
      btfAmount
      btfPayersName
      btfReferenceNumber
      btfJulianDate
      btfTransactionDate
      btfJournalNumber
      btfLineNumber
      btfNINumber
      btfImportNumber
      btfExternalReference
      btfDataSource
      btfPayersIbanNumber
      btfPayersBicCode
      btfIbanNumber
      btfBicCode
      btfPaymentMethod
      btfNotes
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "bank_transactions"
          .Add("statement_date", CDBField.FieldTypes.cftDate)
          .Add("reconciled_status")
          .Add("unreconciled_reason")
          .Add("bank_account")
          .Add("sort_code")
          .Add("account_number")
          .Add("account_type")
          .Add("transaction_code")
          .Add("payers_sort_code")
          .Add("payers_account_number")
          .Add("payers_reference")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("payers_name")
          .Add("reference_number")
          .Add("julian_date")
          .Add("transaction_date", CDBField.FieldTypes.cftDate)
          .Add("journal_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("ni_number")
          .Add("import_number", CDBField.FieldTypes.cftLong)
          .Add("external_reference")
          .Add("data_source")
          .Add("payers_iban_number")
          .Add("payers_bic_code")
          .Add("iban_number")
          .Add("bic_code")
          .Add("payment_method")
          .Add("notes", CDBField.FieldTypes.cftMemo)
        End With

        mvClassFields.Item(BankTransactionFields.btfStatementDate).SetPrimaryKeyOnly()
        mvClassFields.Item(BankTransactionFields.btfLineNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As BankTransactionFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BankTransactionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BankTransactionRecordSetTypes.bktrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bt")
      Else
        If (pRSType And BankTransactionRecordSetTypes.bktrtReconcile) = BankTransactionRecordSetTypes.bktrtReconcile Then
          vFields = "statement_date,line_number,bt.amount,bt.bank_account,reference_number,transaction_date"
          If mvClassFields Is Nothing Then InitClassFields()
          If mvClassFields(BankTransactionFields.btfImportNumber).InDatabase Then
            vFields += ",bt.ni_number,bt.import_number,bt.external_reference,bt.data_source"
          End If
        End If
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BankTransactionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BankTransactionFields.btfStatementDate, vFields)
        .SetItem(BankTransactionFields.btfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BankTransactionRecordSetTypes.bktrtReconcile) = BankTransactionRecordSetTypes.bktrtReconcile Then
          .SetItem(BankTransactionFields.btfAmount, vFields)
          .SetItem(BankTransactionFields.btfBankAccount, vFields)
          .SetItem(BankTransactionFields.btfReferenceNumber, vFields)
          .SetItem(BankTransactionFields.btfTransactionDate, vFields)
          .SetOptionalItem(BankTransactionFields.btfNINumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfImportNumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfExternalReference, vFields)
          .SetOptionalItem(BankTransactionFields.btfDataSource, vFields)
        End If
        If (pRSType And BankTransactionRecordSetTypes.bktrtAll) = BankTransactionRecordSetTypes.bktrtAll Then
          .SetItem(BankTransactionFields.btfReconciledStatus, vFields)
          .SetItem(BankTransactionFields.btfUnreconciledReason, vFields)
          .SetItem(BankTransactionFields.btfSortCode, vFields)
          .SetItem(BankTransactionFields.btfAccountNumber, vFields)
          .SetItem(BankTransactionFields.btfAccountType, vFields)
          .SetItem(BankTransactionFields.btfTransactionCode, vFields)
          .SetItem(BankTransactionFields.btfPayersSortCode, vFields)
          .SetItem(BankTransactionFields.btfPayersAccountNumber, vFields)
          .SetItem(BankTransactionFields.btfPayersName, vFields)
          .SetItem(BankTransactionFields.btfReferenceNumber, vFields)
          .SetItem(BankTransactionFields.btfJulianDate, vFields)
          .SetItem(BankTransactionFields.btfJournalNumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfNINumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfImportNumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfExternalReference, vFields)
          .SetOptionalItem(BankTransactionFields.btfDataSource, vFields)
          .SetOptionalItem(BankTransactionFields.btfPayersIbanNumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfPayersBicCode, vFields)
          .SetOptionalItem(BankTransactionFields.btfIbanNumber, vFields)
          .SetOptionalItem(BankTransactionFields.btfBicCode, vFields)
          .SetOptionalItem(BankTransactionFields.btfPaymentMethod, vFields)
          .SetOptionalItem(BankTransactionFields.btfNotes, vFields)
        End If
      End With
    End Sub

    Public Sub InitFromValues(ByVal pEnv As CDBEnvironment, ByRef pStatementDate As String, ByRef pLineNumber As Integer, ByRef pPayersSortCode As String, ByRef pPayersAccountNumber As String, ByRef pPayersName As String, ByRef pReferenceNumber As String, ByRef pAmount As Double)
      mvEnv = pEnv
      InitClassFields()

      mvExisting = True
      With mvClassFields
        .Item(BankTransactionFields.btfStatementDate).SetValue = pStatementDate
        .Item(BankTransactionFields.btfLineNumber).SetValue = CStr(pLineNumber)
        .Item(BankTransactionFields.btfPayersAccountNumber).SetValue = pPayersAccountNumber
        .Item(BankTransactionFields.btfPayersSortCode).SetValue = pPayersSortCode
        .Item(BankTransactionFields.btfPayersName).SetValue = pPayersName
        .Item(BankTransactionFields.btfReferenceNumber).SetValue = pReferenceNumber
        .Item(BankTransactionFields.btfAmount).SetValue = CStr(pAmount)
      End With
    End Sub

    Public Sub Save()
      SetValid(BankTransactionFields.btfAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pStatementDate As String, ByVal pReconciledStatus As String, ByVal pPayersAccountNumber As String, ByVal pPayersSortCode As String, ByVal pPayersName As String, ByVal pReferenceNumber As String, ByVal pAmount As Double,
                      ByVal pBankAccount As String, ByVal pSortCode As String, ByVal pAccountNumber As String, ByVal pAccountType As String, ByVal pTransactionCode As String, ByVal pPayersReference As String, ByVal pJulianDate As String, ByVal pTransactionDate As String,
                      ByVal pNINumber As String, ByVal pExternalRef As String, ByVal pExternalRefSource As String, ByVal pImportNumber As Integer, ByVal pIbanNumber As String, ByVal pBicCode As String, ByVal pPayersIbanNumber As String, ByVal pPayersBicCode As String,
                      ByVal pPaymentMethod As String, ByVal pNotes As String)
      mvEnv = pEnv
      InitClassFields()

      With mvClassFields
        .Item(BankTransactionFields.btfStatementDate).Value = pStatementDate
        .Item(BankTransactionFields.btfReconciledStatus).Value = pReconciledStatus
        .Item(BankTransactionFields.btfLineNumber).IntegerValue = IntegerValue(mvEnv.Connection.GetValue("SELECT MAX(line_number) FROM bank_transactions WHERE statement_date " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftDate, StatementDate))) + 1
        If Len(pPayersAccountNumber) > 0 Then .Item(BankTransactionFields.btfPayersAccountNumber).Value = pPayersAccountNumber
        If Len(pPayersSortCode) > 0 Then .Item(BankTransactionFields.btfPayersSortCode).Value = pPayersSortCode
        If Len(pPayersName) > 0 Then .Item(BankTransactionFields.btfPayersName).Value = pPayersName
        If Len(pReferenceNumber) > 0 Then .Item(BankTransactionFields.btfReferenceNumber).Value = pReferenceNumber
        .Item(BankTransactionFields.btfAmount).DoubleValue = pAmount
        If Len(pBankAccount) > 0 Then .Item(BankTransactionFields.btfBankAccount).Value = pBankAccount
        .Item(BankTransactionFields.btfSortCode).Value = pSortCode
        .Item(BankTransactionFields.btfAccountNumber).Value = pAccountNumber
        .Item(BankTransactionFields.btfAccountType).Value = pAccountType
        If Len(pTransactionCode) > 0 Then .Item(BankTransactionFields.btfTransactionCode).Value = pTransactionCode
        If Len(pPayersReference) > 0 Then .Item(BankTransactionFields.btfPayersReference).Value = pPayersReference
        .Item(BankTransactionFields.btfJulianDate).Value = pJulianDate
        .Item(BankTransactionFields.btfTransactionDate).Value = pTransactionDate
        If Len(pNINumber) > 0 Then .Item(BankTransactionFields.btfNINumber).Value = pNINumber
        If Len(pExternalRef) > 0 Then .Item(BankTransactionFields.btfExternalReference).Value = pExternalRef
        If Len(pExternalRefSource) > 0 Then .Item(BankTransactionFields.btfDataSource).Value = pExternalRefSource
        If pImportNumber > 0 Then .Item(BankTransactionFields.btfImportNumber).IntegerValue = pImportNumber
        If pIbanNumber.Length > 0 Then .Item(BankTransactionFields.btfIbanNumber).Value = pIbanNumber
        If pBicCode.Length > 0 Then .Item(BankTransactionFields.btfBicCode).Value = pBicCode
        If pPayersIbanNumber.Length > 0 Then .Item(BankTransactionFields.btfPayersIbanNumber).Value = pPayersIbanNumber
        If pPayersBicCode.Length > 0 Then .Item(BankTransactionFields.btfPayersBicCode).Value = pPayersBicCode
        If String.IsNullOrWhiteSpace(pPaymentMethod) = False Then .Item(BankTransactionFields.btfPaymentMethod).Value = pPaymentMethod
        If String.IsNullOrWhiteSpace(pNotes) = False Then .Item(BankTransactionFields.btfNotes).Value = pNotes
        'Setting JournalNumber not currently supported
      End With
      Save()
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AccountNumber() As String
      Get
        AccountNumber = mvClassFields.Item(BankTransactionFields.btfAccountNumber).Value
      End Get
    End Property

    Public ReadOnly Property AccountType() As String
      Get
        AccountType = mvClassFields.Item(BankTransactionFields.btfAccountType).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(BankTransactionFields.btfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(BankTransactionFields.btfBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property DataSource() As String
      Get
        DataSource = mvClassFields.Item(BankTransactionFields.btfDataSource).Value
      End Get
    End Property

    Public ReadOnly Property ExternalReference() As String
      Get
        ExternalReference = mvClassFields.Item(BankTransactionFields.btfExternalReference).Value
      End Get
    End Property

    Public ReadOnly Property ImportNumber() As Integer
      Get
        ImportNumber = mvClassFields.Item(BankTransactionFields.btfImportNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property JournalNumber() As Integer
      Get
        JournalNumber = mvClassFields.Item(BankTransactionFields.btfJournalNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property JulianDate() As String
      Get
        JulianDate = mvClassFields.Item(BankTransactionFields.btfJulianDate).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(BankTransactionFields.btfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property NINumber() As String
      Get
        NINumber = mvClassFields.Item(BankTransactionFields.btfNINumber).Value
      End Get
    End Property

    Public ReadOnly Property PayersAccountNumber() As String
      Get
        PayersAccountNumber = mvClassFields.Item(BankTransactionFields.btfPayersAccountNumber).Value
      End Get
    End Property

    Public ReadOnly Property PayersName() As String
      Get
        PayersName = mvClassFields.Item(BankTransactionFields.btfPayersName).Value
      End Get
    End Property

    Public ReadOnly Property PayersReference() As String
      Get
        PayersReference = mvClassFields.Item(BankTransactionFields.btfPayersReference).Value
      End Get
    End Property

    Public ReadOnly Property PayersSortCode() As String
      Get
        PayersSortCode = mvClassFields.Item(BankTransactionFields.btfPayersSortCode).Value
      End Get
    End Property

    Public Property ReconciledStatus() As String
      Get
        ReconciledStatus = mvClassFields.Item(BankTransactionFields.btfReconciledStatus).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankTransactionFields.btfReconciledStatus).Value = Value
      End Set
    End Property

    Public ReadOnly Property ReferenceNumber() As String
      Get
        ReferenceNumber = mvClassFields.Item(BankTransactionFields.btfReferenceNumber).Value
      End Get
    End Property

    Public ReadOnly Property SortCode() As String
      Get
        SortCode = mvClassFields.Item(BankTransactionFields.btfSortCode).Value
      End Get
    End Property

    Public ReadOnly Property StatementDate() As String
      Get
        StatementDate = mvClassFields.Item(BankTransactionFields.btfStatementDate).Value
      End Get
    End Property

    Public ReadOnly Property TransactionCode() As String
      Get
        TransactionCode = mvClassFields.Item(BankTransactionFields.btfTransactionCode).Value
      End Get
    End Property

    Public ReadOnly Property TransactionDate() As String
      Get
        TransactionDate = mvClassFields.Item(BankTransactionFields.btfTransactionDate).Value
      End Get
    End Property

    Public ReadOnly Property UnreconciledReason() As String
      Get
        UnreconciledReason = mvClassFields.Item(BankTransactionFields.btfUnreconciledReason).Value
      End Get
    End Property

    Public ReadOnly Property PayersIbanNumber() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfPayersIbanNumber).Value
      End Get
    End Property

    Public ReadOnly Property PayersBicCode() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfPayersBicCode).Value
      End Get
    End Property

    Public ReadOnly Property IbanNumber() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfIbanNumber).Value
      End Get
    End Property

    Public ReadOnly Property BicCode() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfBicCode).Value
      End Get
    End Property

    Public ReadOnly Property PaymentMethod() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfPaymentMethod).Value
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields.Item(BankTransactionFields.btfNotes).MultiLineValue
      End Get
    End Property
  End Class
End Namespace
