

Namespace Access
  Public Class IssuedStock

    Public Enum IssuedStockRecordSetTypes 'These are bit values
      isrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum IssuedStockFields
      isfAll = 0
      isfBatchNumber
      isfTransactionNumber
      isfLineNumber
      isfProduct
      isfIssued
      isfAllocated
      isfPickingListNumber
      isfDespatchNoteNumber
      isfExported
      isfWarehouse
      isfJobNumber
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
          .DatabaseTableName = "issued_stock"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("product")
          .Add("issued", CDBField.FieldTypes.cftLong)
          .Add("allocated", CDBField.FieldTypes.cftLong)
          .Add("picking_list_number", CDBField.FieldTypes.cftLong)
          .Add("despatch_note_number", CDBField.FieldTypes.cftLong)
          .Add("exported")
          .Add("warehouse")
          .Add("job_number", CDBField.FieldTypes.cftLong)

          .Item(IssuedStockFields.isfBatchNumber).SetPrimaryKeyOnly()
          .Item(IssuedStockFields.isfTransactionNumber).SetPrimaryKeyOnly()
          .Item(IssuedStockFields.isfLineNumber).SetPrimaryKeyOnly()
          .Item(IssuedStockFields.isfJobNumber).SetPrimaryKeyOnly()

          .Item(IssuedStockFields.isfJobNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIssuedStockJobNumber)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As IssuedStockFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As IssuedStockRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = IssuedStockRecordSetTypes.isrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "iss")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pLineNumber As Integer = 0, Optional ByVal pJobNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      InitClassFields()
      If pJobNumber > 0 Then
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
        vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, pLineNumber)
        vWhereFields.Add("job_number", CDBField.FieldTypes.cftLong, pJobNumber)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(IssuedStockRecordSetTypes.isrtAll) & " FROM issued_stock iss WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, IssuedStockRecordSetTypes.isrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As IssuedStockRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And IssuedStockRecordSetTypes.isrtAll) = IssuedStockRecordSetTypes.isrtAll Then
          .SetItem(IssuedStockFields.isfBatchNumber, vFields)
          .SetItem(IssuedStockFields.isfTransactionNumber, vFields)
          .SetItem(IssuedStockFields.isfLineNumber, vFields)
          .SetItem(IssuedStockFields.isfProduct, vFields)
          .SetItem(IssuedStockFields.isfIssued, vFields)
          .SetItem(IssuedStockFields.isfAllocated, vFields)
          .SetItem(IssuedStockFields.isfPickingListNumber, vFields)
          .SetItem(IssuedStockFields.isfDespatchNoteNumber, vFields)
          .SetItem(IssuedStockFields.isfExported, vFields)
          .SetOptionalItem(IssuedStockFields.isfWarehouse, vFields)
          .SetOptionalItem(IssuedStockFields.isfJobNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pProductCode As String, ByVal pIssued As Integer, ByVal pWarehouse As String, ByVal pJobNumber As Integer, Optional ByVal pPickingListNumber As Integer = 0)
      With mvClassFields
        .Item(IssuedStockFields.isfBatchNumber).IntegerValue = pBatchNumber
        .Item(IssuedStockFields.isfTransactionNumber).IntegerValue = pTransactionNumber
        .Item(IssuedStockFields.isfLineNumber).IntegerValue = pLineNumber
        .Item(IssuedStockFields.isfProduct).Value = pProductCode
        .Item(IssuedStockFields.isfIssued).IntegerValue = pIssued
        .Item(IssuedStockFields.isfAllocated).IntegerValue = 0
        .Item(IssuedStockFields.isfWarehouse).Value = pWarehouse
        .Item(IssuedStockFields.isfJobNumber).Value = CStr(pJobNumber)
        If pPickingListNumber > 0 Then .Item(IssuedStockFields.isfPickingListNumber).IntegerValue = pPickingListNumber
      End With
    End Sub

    Public Sub Delete()
      mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(IssuedStockFields.isfAll)
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

    Public ReadOnly Property Allocated() As Integer
      Get
        Allocated = mvClassFields.Item(IssuedStockFields.isfAllocated).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(IssuedStockFields.isfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DespatchNoteNumber() As Integer
      Get
        DespatchNoteNumber = mvClassFields.Item(IssuedStockFields.isfDespatchNoteNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Exported() As Boolean
      Get
        Exported = mvClassFields.Item(IssuedStockFields.isfExported).Bool
      End Get
    End Property

    Public ReadOnly Property Issued() As Integer
      Get
        Issued = mvClassFields.Item(IssuedStockFields.isfIssued).IntegerValue
      End Get
    End Property

    Public ReadOnly Property JobNumber() As Integer
      Get
        JobNumber = mvClassFields.Item(IssuedStockFields.isfJobNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(IssuedStockFields.isfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PickingListNumber() As Integer
      Get
        PickingListNumber = mvClassFields.Item(IssuedStockFields.isfPickingListNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(IssuedStockFields.isfProduct).Value
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(IssuedStockFields.isfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Warehouse() As String
      Get
        Warehouse = mvClassFields.Item(IssuedStockFields.isfWarehouse).Value
      End Get
    End Property
  End Class
End Namespace
