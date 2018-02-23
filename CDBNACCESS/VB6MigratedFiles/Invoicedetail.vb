

Namespace Access
  Public Class InvoiceDetail

    Public Enum InvoiceDetailRecordSetTypes 'These are bit values
      idrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum InvoiceDetailFields
      idfAll = 0
      idfBatchNumber
      idfTransactionNumber
      idfLineNumber
      idfInvoiceNumber
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
          .DatabaseTableName = "invoice_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("invoice_number", CDBField.FieldTypes.cftLong)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As InvoiceDetailFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As InvoiceDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = InvoiceDetailRecordSetTypes.idrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "id")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As InvoiceDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And InvoiceDetailRecordSetTypes.idrtAll) = InvoiceDetailRecordSetTypes.idrtAll Then
          .SetItem(InvoiceDetailFields.idfBatchNumber, vFields)
          .SetItem(InvoiceDetailFields.idfTransactionNumber, vFields)
          .SetItem(InvoiceDetailFields.idfLineNumber, vFields)
          .SetItem(InvoiceDetailFields.idfInvoiceNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransNo As Integer, ByVal pLineNo As Integer, ByVal pInvoiceNumber As Integer)
      Init(pEnv)
      mvClassFields.Item(InvoiceDetailFields.idfBatchNumber).IntegerValue = pBatchNumber
      mvClassFields.Item(InvoiceDetailFields.idfTransactionNumber).IntegerValue = pTransNo
      mvClassFields.Item(InvoiceDetailFields.idfLineNumber).IntegerValue = pLineNo
      If pInvoiceNumber > 0 Then mvClassFields.Item(InvoiceDetailFields.idfInvoiceNumber).IntegerValue = pInvoiceNumber
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(InvoiceDetailFields.idfAll)
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
        BatchNumber = mvClassFields.Item(InvoiceDetailFields.idfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property InvoiceNumber() As String
      Get
        InvoiceNumber = mvClassFields.Item(InvoiceDetailFields.idfInvoiceNumber).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(InvoiceDetailFields.idfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(InvoiceDetailFields.idfTransactionNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
