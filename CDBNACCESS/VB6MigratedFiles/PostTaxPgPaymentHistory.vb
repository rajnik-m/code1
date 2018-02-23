

Namespace Access
  Public Class PostTaxPgPaymentHistory

    Public Enum PostTaxPgPaymentHistoryRecordSetTypes 'These are bit values
      ptpphrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PostTaxPgPaymentHistoryFields
      ptpphfAll = 0
      ptpphfPledgeNumber
      ptpphfBatchNumber
      ptpphfTransactionNumber
      ptpphfPaymentNumber
      ptpphfDonorAmount
      ptpphfEmployerAmount
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvOrgPayment As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "post_tax_pg_payment_history"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("pledge_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("payment_number", CDBField.FieldTypes.cftInteger)
          .Add("donor_amount", CDBField.FieldTypes.cftNumeric)
          .Add("employer_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfPledgeNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfTransactionNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvOrgPayment = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfDonorAmount).DoubleValue = 0
      mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfEmployerAmount).DoubleValue = 0
    End Sub

    Private Sub SetValid(ByVal pField As PostTaxPgPaymentHistoryFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PostTaxPgPaymentHistoryRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PostTaxPgPaymentHistoryRecordSetTypes.ptpphrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ptpph")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pPledgeNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If (pBatchNumber > 0 And pTransactionNumber > 0) Then
        vSQL = "SELECT " & GetRecordSetFields(PostTaxPgPaymentHistoryRecordSetTypes.ptpphrtAll) & " FROM post_tax_pg_payment_history ptpph WHERE"
        If pPledgeNumber > 0 Then vSQL = vSQL & " pledge_number = " & pPledgeNumber & " AND"
        vSQL = vSQL & " batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PostTaxPgPaymentHistoryRecordSetTypes.ptpphrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PostTaxPgPaymentHistoryRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PostTaxPgPaymentHistoryFields.ptpphfPledgeNumber, vFields)
        .SetItem(PostTaxPgPaymentHistoryFields.ptpphfBatchNumber, vFields)
        .SetItem(PostTaxPgPaymentHistoryFields.ptpphfTransactionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PostTaxPgPaymentHistoryRecordSetTypes.ptpphrtAll) = PostTaxPgPaymentHistoryRecordSetTypes.ptpphrtAll Then
          .SetItem(PostTaxPgPaymentHistoryFields.ptpphfPaymentNumber, vFields)
          .SetItem(PostTaxPgPaymentHistoryFields.ptpphfDonorAmount, vFields)
          .SetItem(PostTaxPgPaymentHistoryFields.ptpphfEmployerAmount, vFields)
        End If
      End With
    End Sub

    Public Sub InitOrganisationPayment(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      'Used by Trader financial adjustments to retrieve payments allocated against the Employer
      'This is for information only and is not saved
      Dim vRS As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()

      vSQL = "SELECT bta.line_number, bta.amount FROM batch_transactions bt, organisations o, batch_transaction_analysis bta"
      vSQL = vSQL & " WHERE bt.batch_number = " & pBatchNumber & " AND bt.transaction_number = " & pTransactionNumber
      vSQL = vSQL & " AND o.organisation_number = bt.contact_number AND bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number"
      vSQL = vSQL & " ORDER BY line_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        mvOrgPayment = True
        mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfBatchNumber).IntegerValue = pBatchNumber
        mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfTransactionNumber).IntegerValue = pTransactionNumber
        If vRS.Fields("line_number").IntegerValue = 1 Then
          mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfDonorAmount).DoubleValue = vRS.Fields("amount").DoubleValue
        Else
          mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfEmployerAmount).DoubleValue = vRS.Fields("amount").DoubleValue
        End If
      End While
      vRS.CloseRecordSet()

    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pPledgeNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pPaymentNumber As Integer, ByVal pDonorAmount As Double, ByVal pEmployerAmount As Double)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()

      With mvClassFields
        .Item(PostTaxPgPaymentHistoryFields.ptpphfPledgeNumber).Value = CStr(pPledgeNumber)
        .Item(PostTaxPgPaymentHistoryFields.ptpphfBatchNumber).Value = CStr(pBatchNumber)
        .Item(PostTaxPgPaymentHistoryFields.ptpphfTransactionNumber).Value = CStr(pTransactionNumber)
        .Item(PostTaxPgPaymentHistoryFields.ptpphfPaymentNumber).Value = CStr(pPaymentNumber)
        .Item(PostTaxPgPaymentHistoryFields.ptpphfDonorAmount).Value = CStr(pDonorAmount)
        .Item(PostTaxPgPaymentHistoryFields.ptpphfEmployerAmount).Value = CStr(pEmployerAmount)
      End With

    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PostTaxPgPaymentHistoryFields.ptpphfAll)
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

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DonorAmount() As Double
      Get
        DonorAmount = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfDonorAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property EmployerAmount() As Double
      Get
        EmployerAmount = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfEmployerAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property PaymentNumber() As Integer
      Get
        PaymentNumber = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PledgeNumber() As Integer
      Get
        PledgeNumber = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfPledgeNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(PostTaxPgPaymentHistoryFields.ptpphfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OrganisationPayment() As Boolean
      Get
        OrganisationPayment = mvOrgPayment
      End Get
    End Property
  End Class
End Namespace
