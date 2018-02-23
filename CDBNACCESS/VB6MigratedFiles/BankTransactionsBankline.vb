Namespace Access
  Public Class BankTransactionsBankline

    Public Enum BankTransactionsBanklineRecordSetTypes 'These are bit values
      btbrtAll = &HFFFFS
      btbrtReconcile = 1
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BankTransactionsBanklineFields
      btbfAll = 0
      btbfStatementDate
      btbfReconciledStatus
      btbfReference
      btbfTransactionDate
      btbfValueDate
      btbfInToday
      btbfAmount
      btbfDetails
      btbfTransactionCode
      btbfRfr
      btbfChequeNumber
      btbfLineNumber
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
          .DatabaseTableName = "bank_transactions_bankline"
          .Add("statement_date", CDBField.FieldTypes.cftDate)
          .Add("reconciled_status")
          .Add("reference")
          .Add("transaction_date", CDBField.FieldTypes.cftDate)
          .Add("value_date", CDBField.FieldTypes.cftDate)
          .Add("in_today")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("details")
          .Add("transaction_code")
          .Add("rfr")
          .Add("cheque_number")
          .Add("line_number", CDBField.FieldTypes.cftLong)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As BankTransactionsBanklineFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BankTransactionsBanklineRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BankTransactionsBanklineRecordSetTypes.btbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "btb")
      Else
        If (pRSType And BankTransactionsBanklineRecordSetTypes.btbrtReconcile) = BankTransactionsBanklineRecordSetTypes.btbrtReconcile Then
          vFields = "statement_date,line_number,bt.amount,bt.reference,transaction_date"
        End If
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BankTransactionsBanklineRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BankTransactionsBanklineFields.btbfStatementDate, vFields)
        .SetItem(BankTransactionsBanklineFields.btbfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BankTransactionsBanklineRecordSetTypes.btbrtReconcile) = BankTransactionsBanklineRecordSetTypes.btbrtReconcile Then
          .SetItem(BankTransactionsBanklineFields.btbfReference, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfAmount, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfTransactionDate, vFields)
        End If
        If (pRSType And BankTransactionsBanklineRecordSetTypes.btbrtAll) = BankTransactionsBanklineRecordSetTypes.btbrtAll Then
          .SetItem(BankTransactionsBanklineFields.btbfReconciledStatus, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfValueDate, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfInToday, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfDetails, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfTransactionCode, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfRfr, vFields)
          .SetItem(BankTransactionsBanklineFields.btbfChequeNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(BankTransactionsBanklineFields.btbfAll)
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

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(BankTransactionsBanklineFields.btbfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property ChequeNumber() As String
      Get
        ChequeNumber = mvClassFields.Item(BankTransactionsBanklineFields.btbfChequeNumber).Value
      End Get
    End Property

    Public ReadOnly Property Details() As String
      Get
        Details = mvClassFields.Item(BankTransactionsBanklineFields.btbfDetails).Value
      End Get
    End Property

    Public ReadOnly Property InToday() As String
      Get
        InToday = mvClassFields.Item(BankTransactionsBanklineFields.btbfInToday).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(BankTransactionsBanklineFields.btbfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ReconciledStatus() As String
      Get
        ReconciledStatus = mvClassFields.Item(BankTransactionsBanklineFields.btbfReconciledStatus).Value
      End Get
    End Property

    Public ReadOnly Property Reference() As String
      Get
        Reference = mvClassFields.Item(BankTransactionsBanklineFields.btbfReference).Value
      End Get
    End Property

    Public ReadOnly Property Rfr() As String
      Get
        Rfr = mvClassFields.Item(BankTransactionsBanklineFields.btbfRfr).Value
      End Get
    End Property

    Public ReadOnly Property StatementDate() As String
      Get
        StatementDate = mvClassFields.Item(BankTransactionsBanklineFields.btbfStatementDate).Value
      End Get
    End Property

    Public ReadOnly Property TransactionCode() As String
      Get
        TransactionCode = mvClassFields.Item(BankTransactionsBanklineFields.btbfTransactionCode).Value
      End Get
    End Property

    Public ReadOnly Property TransactionDate() As String
      Get
        TransactionDate = mvClassFields.Item(BankTransactionsBanklineFields.btbfTransactionDate).Value
      End Get
    End Property

    Public ReadOnly Property ValueDate() As String
      Get
        ValueDate = mvClassFields.Item(BankTransactionsBanklineFields.btbfValueDate).Value
      End Get
    End Property
  End Class
End Namespace
