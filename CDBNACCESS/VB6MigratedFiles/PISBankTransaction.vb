Namespace Access
  Public Class PISBankTransaction

    Public Enum PisBankTransactionRecordSetTypes 'These are bit values
      pbtrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PisBankTransactionFields
      pbtfAll = 0
      pbtfPisBankTransactionNumber
      pbtfPisBankStatementNumber
      pbtfSortCode
      pbtfBankAccount
      pbtfPayersReference
      pbtfPisNumber
      pbtfAmount
      pbtfJulianDate
      pbtfTransactionDate
      pbtfReconciledStatus
      pbtfReconciledOn
      pbtfUnreconciledReason
      pbtfAmendedBy
      pbtfAmendedOn
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
        With mvClassFields
          .DatabaseTableName = "pis_bank_transactions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("pis_bank_transaction_number", CDBField.FieldTypes.cftLong)
          .Add("pis_bank_statement_number", CDBField.FieldTypes.cftLong)
          .Add("sort_code")
          .Add("bank_account")
          .Add("payers_reference")
          .Add("pis_number", CDBField.FieldTypes.cftCharacter)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("julian_date")
          .Add("transaction_date", CDBField.FieldTypes.cftDate)
          .Add("reconciled_status")
          .Add("reconciled_on", CDBField.FieldTypes.cftDate)
          .Add("unreconciled_reason")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(PisBankTransactionFields.pbtfPisBankTransactionNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(PisBankTransactionFields.pbtfPisBankStatementNumber).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfPisNumber).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfAmount).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfReconciledStatus).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfReconciledOn).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfBankAccount).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfAmendedOn).PrefixRequired = True
        mvClassFields.Item(PisBankTransactionFields.pbtfAmendedBy).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As PisBankTransactionFields)
      'Add code here to ensure all values are valid before saving
      With mvClassFields
        If .Item(PisBankTransactionFields.pbtfPisBankTransactionNumber).IntegerValue = 0 Then .Item(PisBankTransactionFields.pbtfPisBankTransactionNumber).IntegerValue = mvEnv.GetControlNumber("PT")
        .Item(PisBankTransactionFields.pbtfAmendedOn).Value = TodaysDate()
        .Item(PisBankTransactionFields.pbtfAmendedBy).Value = mvEnv.User.Logname
      End With

    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PisBankTransactionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PisBankTransactionRecordSetTypes.pbtrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pbt")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pPisBankTransactionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pPisBankTransactionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PisBankTransactionRecordSetTypes.pbtrtAll) & " FROM pis_bank_transactions pbt WHERE pis_bank_transaction_number = " & pPisBankTransactionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PisBankTransactionRecordSetTypes.pbtrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PisBankTransactionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PisBankTransactionFields.pbtfPisBankTransactionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PisBankTransactionRecordSetTypes.pbtrtAll) = PisBankTransactionRecordSetTypes.pbtrtAll Then
          .SetItem(PisBankTransactionFields.pbtfPisBankStatementNumber, vFields)
          .SetItem(PisBankTransactionFields.pbtfSortCode, vFields)
          .SetItem(PisBankTransactionFields.pbtfBankAccount, vFields)
          .SetItem(PisBankTransactionFields.pbtfPayersReference, vFields)
          .SetItem(PisBankTransactionFields.pbtfPisNumber, vFields)
          .SetItem(PisBankTransactionFields.pbtfAmount, vFields)
          .SetItem(PisBankTransactionFields.pbtfJulianDate, vFields)
          .SetItem(PisBankTransactionFields.pbtfTransactionDate, vFields)
          .SetItem(PisBankTransactionFields.pbtfReconciledStatus, vFields)
          .SetItem(PisBankTransactionFields.pbtfReconciledOn, vFields)
          .SetItem(PisBankTransactionFields.pbtfUnreconciledReason, vFields)
          .SetItem(PisBankTransactionFields.pbtfAmendedBy, vFields)
          .SetItem(PisBankTransactionFields.pbtfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PisBankTransactionFields.pbtfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(PisBankTransactionFields.pbtfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PisBankTransactionFields.pbtfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = CDbl(mvClassFields.Item(PisBankTransactionFields.pbtfAmount).Value)
      End Get
    End Property

    Public ReadOnly Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(PisBankTransactionFields.pbtfBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property JulianDate() As String
      Get
        JulianDate = mvClassFields.Item(PisBankTransactionFields.pbtfJulianDate).Value
      End Get
    End Property

    Public ReadOnly Property PayersReference() As String
      Get
        PayersReference = mvClassFields.Item(PisBankTransactionFields.pbtfPayersReference).Value
      End Get
    End Property

    Public ReadOnly Property PisBankStatementNumber() As Integer
      Get
        PisBankStatementNumber = CInt(mvClassFields.Item(PisBankTransactionFields.pbtfPisBankStatementNumber).Value)
      End Get
    End Property

    Public ReadOnly Property PisBankTransactionNumber() As Integer
      Get
        PisBankTransactionNumber = CInt(mvClassFields.Item(PisBankTransactionFields.pbtfPisBankTransactionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property PisNumber() As Integer
      Get
        PisNumber = CInt(mvClassFields.Item(PisBankTransactionFields.pbtfPisNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ReconciledOn() As String
      Get
        ReconciledOn = mvClassFields.Item(PisBankTransactionFields.pbtfReconciledOn).Value
      End Get
    End Property

    Public ReadOnly Property ReconciledStatus() As String
      Get
        ReconciledStatus = mvClassFields.Item(PisBankTransactionFields.pbtfReconciledStatus).Value
      End Get
    End Property

    Public ReadOnly Property SortCode() As String
      Get
        SortCode = mvClassFields.Item(PisBankTransactionFields.pbtfSortCode).Value
      End Get
    End Property

    Public ReadOnly Property TransactionDate() As String
      Get
        TransactionDate = mvClassFields.Item(PisBankTransactionFields.pbtfTransactionDate).Value
      End Get
    End Property

    Public ReadOnly Property UnreconciledReason() As String
      Get
        UnreconciledReason = mvClassFields.Item(PisBankTransactionFields.pbtfUnreconciledReason).Value
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      Init(pEnv)
      With mvClassFields
        .Item(PisBankTransactionFields.pbtfPisBankStatementNumber).Value = pParams("StatementNumber").Value
        .Item(PisBankTransactionFields.pbtfReconciledStatus).Value = pParams("ReconciledStatus").Value
        .Item(PisBankTransactionFields.pbtfSortCode).Value = pParams("SortCode").Value
        .Item(PisBankTransactionFields.pbtfBankAccount).Value = pParams("BankAccount").Value
        .Item(PisBankTransactionFields.pbtfPayersReference).Value = pParams("PayersReference").Value
        .Item(PisBankTransactionFields.pbtfPisNumber).Value = pParams("PISNumber").Value
        .Item(PisBankTransactionFields.pbtfAmount).DoubleValue = pParams("Amount").DoubleValue
        .Item(PisBankTransactionFields.pbtfJulianDate).Value = pParams("JulianDate").Value
        .Item(PisBankTransactionFields.pbtfTransactionDate).Value = pParams("TransactionDate").Value
      End With
    End Sub

    Public Sub MarkUnReconciled(ByRef pUnReconciledReasonCode As String)
      mvClassFields(PisBankTransactionFields.pbtfReconciledStatus).Value = "U"
      mvClassFields(PisBankTransactionFields.pbtfUnreconciledReason).Value = pUnReconciledReasonCode
    End Sub

    Public Sub MarkReconciled()
      mvClassFields(PisBankTransactionFields.pbtfReconciledStatus).Value = "F"
      mvClassFields(PisBankTransactionFields.pbtfReconciledOn).Value = TodaysDate()

    End Sub
  End Class
End Namespace
