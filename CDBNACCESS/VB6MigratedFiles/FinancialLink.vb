

Namespace Access
  Public Class FinancialLink

    Public Enum FinancialLinkRecordSetTypes 'These are bit values
      flrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum FinancialLinkFields
      flfAll = 0
      flfContactNumber
      flfDonorContactNumber
      flfBatchNumber
      flfTransactionNumber
      flfLineNumber
      flfLineType
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
          .DatabaseTableName = "financial_links"
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("donor_contact_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("line_type")
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As FinancialLinkFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As FinancialLinkRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = FinancialLinkRecordSetTypes.flrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "fl")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As FinancialLinkRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And FinancialLinkRecordSetTypes.flrtAll) = FinancialLinkRecordSetTypes.flrtAll Then
          .SetItem(FinancialLinkFields.flfContactNumber, vFields)
          .SetItem(FinancialLinkFields.flfDonorContactNumber, vFields)
          .SetItem(FinancialLinkFields.flfBatchNumber, vFields)
          .SetItem(FinancialLinkFields.flfTransactionNumber, vFields)
          .SetItem(FinancialLinkFields.flfLineNumber, vFields)
          .SetItem(FinancialLinkFields.flfLineType, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(FinancialLinkFields.flfAll)
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

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(FinancialLinkFields.flfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(FinancialLinkFields.flfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DonorContactNumber() As Integer
      Get
        DonorContactNumber = mvClassFields.Item(FinancialLinkFields.flfDonorContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(FinancialLinkFields.flfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineType() As String
      Get
        LineType = mvClassFields.Item(FinancialLinkFields.flfLineType).Value
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(FinancialLinkFields.flfTransactionNumber).IntegerValue
      End Get
    End Property

    Public Sub InitFromValues(ByRef pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pDonorContactNumber As Integer, ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, ByRef pLineNumber As Integer, ByRef pLineType As String)
      Init(pEnv)
      mvClassFields.Item(FinancialLinkFields.flfContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(FinancialLinkFields.flfDonorContactNumber).Value = CStr(pDonorContactNumber)
      mvClassFields.Item(FinancialLinkFields.flfBatchNumber).Value = CStr(pBatchNumber)
      mvClassFields.Item(FinancialLinkFields.flfTransactionNumber).Value = CStr(pTransactionNumber)
      mvClassFields.Item(FinancialLinkFields.flfLineNumber).Value = CStr(pLineNumber)
      mvClassFields.Item(FinancialLinkFields.flfLineType).Value = pLineType
    End Sub
  End Class
End Namespace
