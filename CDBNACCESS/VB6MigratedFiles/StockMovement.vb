

Namespace Access
  Public Class StockMovement

    Public Enum StockMovementRecordSetTypes 'These are bit values
      smrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum StockMovementFields
      smfAll = 0
      smfStockMovementNumber
      smfProduct
      smfWarehouse
      smfStockMovementReason
      smfMovementQuantity
      smfResultingStockCount
      smfBatchNumber
      smfTransactionNumber
      smfLineNumber
      smfAmendedBy
      smfAmendedOn
      smfProductCostNumber
      smfTransactionID
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    'Private mvProduct      As Product

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "stock_movements"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("stock_movement_number", CDBField.FieldTypes.cftLong)
          .Add("product")
          .Add("warehouse")
          .Add("stock_movement_reason")
          .Add("movement_quantity", CDBField.FieldTypes.cftLong)
          .Add("resulting_stock_count", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("product_cost_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_id", CDBField.FieldTypes.cftLong)
        End With

        mvClassFields.Item(StockMovementFields.smfStockMovementNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(StockMovementFields.smfProductCostNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts)
        mvClassFields.Item(StockMovementFields.smfProductCostNumber).PrefixRequired = True
        mvClassFields.Item(StockMovementFields.smfTransactionID).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataStockMovementTransactionID)
      Else
        mvClassFields.ClearItems()
      End If
      'Set mvProduct = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As StockMovementFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(StockMovementFields.smfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(StockMovementFields.smfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As StockMovementRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = StockMovementRecordSetTypes.smrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sm")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pStockMovementNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pStockMovementNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(StockMovementRecordSetTypes.smrtAll) & " FROM stock_movements sm WHERE stock_movement_number = " & pStockMovementNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, StockMovementRecordSetTypes.smrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As StockMovementRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(StockMovementFields.smfStockMovementNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And StockMovementRecordSetTypes.smrtAll) = StockMovementRecordSetTypes.smrtAll Then
          .SetItem(StockMovementFields.smfProduct, vFields)
          .SetItem(StockMovementFields.smfWarehouse, vFields)
          .SetItem(StockMovementFields.smfStockMovementReason, vFields)
          .SetItem(StockMovementFields.smfMovementQuantity, vFields)
          .SetItem(StockMovementFields.smfResultingStockCount, vFields)
          .SetItem(StockMovementFields.smfBatchNumber, vFields)
          .SetItem(StockMovementFields.smfTransactionNumber, vFields)
          .SetItem(StockMovementFields.smfLineNumber, vFields)
          .SetItem(StockMovementFields.smfAmendedBy, vFields)
          .SetItem(StockMovementFields.smfAmendedOn, vFields)
          .SetOptionalItem(StockMovementFields.smfProductCostNumber, vFields)
          .SetOptionalItem(StockMovementFields.smfTransactionID, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByRef pProductCode As String, ByRef pQuantity As Integer, ByRef pReason As String, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pLineNumber As Integer = 0, Optional ByVal pQuantityShortfallError As Boolean = True, Optional ByRef pWarehouse As String = "", Optional ByVal pProductCostNumber As Integer = 0, Optional ByRef pTransactionID As Integer = 0)
      Dim vProduct As New Product(pEnv)
      Dim vProductCost As New ProductCost
      Dim vWarehouse As New ProductWarehouse
      Dim vTransaction As Boolean
      Dim vSave As Boolean
      Dim vInitialValue As Boolean
      Dim vShortFall As Boolean
      Dim vLastStockCount As Integer
      Dim vUpdateQuantityOnOrder As Boolean

      mvEnv = pEnv
      InitClassFields()
      vInitialValue = pReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonInitial)
      If Len(pReason) > 0 Then
        vShortFall = pReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonShortFall)
        mvClassFields.Item(StockMovementFields.smfStockMovementNumber).Value = CStr(mvEnv.GetControlNumber("SM"))

        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
          vUpdateQuantityOnOrder = pEnv.Connection.GetValue("SELECT update_quantity_on_order FROM stock_movement_reasons WHERE stock_movement_reason = '" & pReason & "'") = "Y"
        End If
      End If
      mvClassFields.Item(StockMovementFields.smfProduct).Value = pProductCode
      mvClassFields.Item(StockMovementFields.smfWarehouse).Value = pWarehouse
      mvClassFields.Item(StockMovementFields.smfStockMovementReason).Value = pReason

      If pBatchNumber > 0 Then mvClassFields.Item(StockMovementFields.smfBatchNumber).Value = CStr(pBatchNumber)
      If pTransactionNumber > 0 Then mvClassFields.Item(StockMovementFields.smfTransactionNumber).Value = CStr(pTransactionNumber)
      If pLineNumber > 0 Then mvClassFields.Item(StockMovementFields.smfLineNumber).Value = CStr(pLineNumber)

      If Not mvEnv.Connection.InTransaction Then
        mvEnv.Connection.StartTransaction()
        vTransaction = True
      End If

      vProduct.Init(pProductCode) 'Moved outside of with for .NET Conversion
      With vProduct
        If .Existing = False Then RaiseError(DataAccessErrors.daeMissingProduct, pProductCode)
        'If we did not ask for a specific warehouse then use the default one from the product
        If Len(pWarehouse) = 0 Then mvClassFields.Item(StockMovementFields.smfWarehouse).Value = .Warehouse
        vWarehouse.Init(mvEnv, .ProductCode, Warehouse) 'Get the warehouse for this stock movement
        vLastStockCount = vWarehouse.LastStockCount 'Last count is from the warehouse
        If Not (vWarehouse.Existing) Then vLastStockCount = .LastStockCount 'Last stock count is from the product
        vProductCost.Init(mvEnv)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then vProductCost.Init(mvEnv, pProductCostNumber)

        If vShortFall Then pQuantity = -vLastStockCount 'Adjust the quantity to say we are removing this number

        If vLastStockCount + pQuantity < 0 Then 'Check if we have enough stock
          If pQuantityShortfallError Then
            RaiseError(DataAccessErrors.daeInvalidStockMovement, (vProduct.ProductCode)) 'Raise an error
          Else
            pQuantity = -vLastStockCount 'Adjust the quantity to say we are removing this number
          End If
        End If

        'If using ProductCosts then ensure that we still have sufficient stock for this ProductCost record
        If vProductCost.Existing = True And pQuantity <> 0 Then
          If vProductCost.LastStockCount + pQuantity < 0 Then
            pQuantity = -vProductCost.LastStockCount 'Adjust the quantity to say we are removing this number
          End If
        End If

        If pQuantity <> 0 Then 'If we are moving any stock
          If vWarehouse.Existing Then
            vWarehouse.LastStockCount = vLastStockCount + pQuantity 'Adjust the warehouse quantity
            If vUpdateQuantityOnOrder Then
              If pQuantity > vWarehouse.QuantityOnOrder Then
                vWarehouse.QuantityOnOrder = 0
              Else
                vWarehouse.QuantityOnOrder = vWarehouse.QuantityOnOrder - pQuantity
              End If
            End If
            vWarehouse.Save()
          End If
          .LastStockCount = .LastStockCount + pQuantity
          .Save()
        End If
        mvClassFields.Item(StockMovementFields.smfResultingStockCount).Value = CStr(vLastStockCount + pQuantity) 'The resultant stock count (either for the product or for the warehouse)
        mvClassFields.Item(StockMovementFields.smfMovementQuantity).Value = CStr(pQuantity)

        If vProductCost.Existing Then
          mvClassFields.Item(StockMovementFields.smfProductCostNumber).Value = CStr(vProductCost.ProductCostNumber)
          vProductCost.SellStock(pQuantity)
          vProductCost.Save()
        End If

        If Len(pReason) > 0 Then
          'Control Values will be zero if database not updated - so just update product
          If vInitialValue Or vShortFall Then
            vSave = True
          ElseIf mvClassFields.Item(StockMovementFields.smfMovementQuantity).IntegerValue <> 0 Then
            vSave = True
          ElseIf mvClassFields.Item(StockMovementFields.smfMovementQuantity).IntegerValue = 0 Then
            'Only save 0 quantity if going on to back order
            If pReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonAwaitBackOrder) Then vSave = True
          End If
          If vSave Then
            'Only set pTransactionID if we are saving the StockMovement
            If pTransactionID = 0 Then pTransactionID = StockMovementNumber
            mvClassFields.Item(StockMovementFields.smfTransactionID).Value = CStr(pTransactionID)
            Save()
          End If
        End If
      End With
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(StockMovementFields.smfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub SetBatch(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer)
      mvClassFields(StockMovementFields.smfBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(StockMovementFields.smfTransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields(StockMovementFields.smfLineNumber).IntegerValue = pLineNumber
    End Sub

    Public Sub UpdateStockLevels(ByVal pEnv As CDBEnvironment, ByVal pProductCode As String, ByVal pWarehouse As String, ByVal pQuantity As Integer, ByVal pProductCostNumber As Integer)
      'This will add the stock quantity back to the product/warehouse
      Dim vProduct As New Product(mvEnv)
      Dim vPrWarehouse As New ProductWarehouse
      Dim vPrCost As New ProductCost
      Dim vTrans As Boolean

      vPrWarehouse.Init(pEnv)
      vPrCost.Init(pEnv)

      'Always start a Transaction to ensure that all stock counts get updated together, just like when we originally sold the stock
      If pEnv.Connection.InTransaction = False Then
        pEnv.Connection.StartTransaction()
        vTrans = True
      End If

      vProduct.Init(pProductCode)
      If Len(pWarehouse) = 0 Then pWarehouse = vProduct.Warehouse
      vPrWarehouse.Init(pEnv, pProductCode, pWarehouse)
      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then vPrCost.Init(pEnv, pProductCostNumber)

      'Update stock counts
      vProduct.LastStockCount = vProduct.LastStockCount + pQuantity
      If vPrWarehouse.Existing Then vPrWarehouse.LastStockCount = vPrWarehouse.LastStockCount + pQuantity
      If vPrCost.Existing Then vPrCost.SellStock(pQuantity)

      'Save the stock counts
      vProduct.Save()
      If vPrWarehouse.Existing Then vPrWarehouse.Save()
      If vPrCost.Existing Then vPrCost.Save()

      'Commit the Transaction
      If vTrans Then pEnv.Connection.CommitTransaction()
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(StockMovementFields.smfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(StockMovementFields.smfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(StockMovementFields.smfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(StockMovementFields.smfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MovementQuantity() As Integer
      Get
        MovementQuantity = mvClassFields.Item(StockMovementFields.smfMovementQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(StockMovementFields.smfProduct).Value
      End Get
    End Property

    Public ReadOnly Property ProductCostNumber() As Integer
      Get
        ProductCostNumber = mvClassFields.Item(StockMovementFields.smfProductCostNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ResultingStockCount() As Integer
      Get
        ResultingStockCount = mvClassFields.Item(StockMovementFields.smfResultingStockCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StockMovementNumber() As Integer
      Get
        StockMovementNumber = mvClassFields.Item(StockMovementFields.smfStockMovementNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StockMovementReason() As String
      Get
        StockMovementReason = mvClassFields.Item(StockMovementFields.smfStockMovementReason).Value
      End Get
    End Property

    Public ReadOnly Property TransactionID() As Integer
      Get
        TransactionID = mvClassFields.Item(StockMovementFields.smfTransactionID).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(StockMovementFields.smfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Warehouse() As String
      Get
        Warehouse = mvClassFields.Item(StockMovementFields.smfWarehouse).Value
      End Get
    End Property

  End Class
End Namespace
