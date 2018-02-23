Namespace Access
  Public Class BackOrderDetail

    Public Enum BackOrderDetailRecordSetTypes 'These are bit values
      bodrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BackOrderDetailFields
      bodfAll = 0
      bodfBatchNumber
      bodfTransactionNumber
      bodfLineNumber
      bodfProduct
      bodfRate
      bodfOrdered
      bodfIssued
      bodfEarliestDelivery
      bodfDespatchMethod
      bodfContactNumber
      bodfAddressNumber
      bodfSource
      bodfUnitPrice
      bodfGrossAmount
      bodfDiscount
      bodfVatRate
      bodfVatAmount
      bodfStatus
      bodfCurrencyUnitPrice
      bodfCurrencyVatAmount
      bodfWarehouse
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
          .DatabaseTableName = "back_order_details"
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("product")
          .Add("rate")
          .Add("ordered", CDBField.FieldTypes.cftInteger)
          .Add("issued", CDBField.FieldTypes.cftInteger)
          .Add("earliest_delivery", CDBField.FieldTypes.cftDate)
          .Add("despatch_method")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("source")
          .Add("unit_price", CDBField.FieldTypes.cftNumeric)
          .Add("gross_amount", CDBField.FieldTypes.cftNumeric)
          .Add("discount", CDBField.FieldTypes.cftNumeric)
          .Add("vat_rate")
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric)
          .Add("status")
          .Add("currency_unit_price", CDBField.FieldTypes.cftNumeric)
          .Add("currency_vat_amount", CDBField.FieldTypes.cftNumeric)
          .Add("warehouse")
        End With

        mvClassFields.Item(BackOrderDetailFields.bodfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(BackOrderDetailFields.bodfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(BackOrderDetailFields.bodfLineNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(BackOrderDetailFields.bodfBatchNumber).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfTransactionNumber).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfProduct).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfRate).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfDespatchMethod).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfSource).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfWarehouse).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfContactNumber).PrefixRequired = True
        mvClassFields.Item(BackOrderDetailFields.bodfAddressNumber).PrefixRequired = True

        mvClassFields.Item(BackOrderDetailFields.bodfCurrencyUnitPrice).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode)
        mvClassFields.Item(BackOrderDetailFields.bodfCurrencyVatAmount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As BackOrderDetailFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BackOrderDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BackOrderDetailRecordSetTypes.bodrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bod")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(BackOrderDetailRecordSetTypes.bodrtAll) & " FROM back_order_details bod WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BackOrderDetailRecordSetTypes.bodrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BackOrderDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BackOrderDetailFields.bodfBatchNumber, vFields)
        .SetItem(BackOrderDetailFields.bodfTransactionNumber, vFields)
        .SetItem(BackOrderDetailFields.bodfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BackOrderDetailRecordSetTypes.bodrtAll) = BackOrderDetailRecordSetTypes.bodrtAll Then
          .SetItem(BackOrderDetailFields.bodfProduct, vFields)
          .SetItem(BackOrderDetailFields.bodfRate, vFields)
          .SetItem(BackOrderDetailFields.bodfOrdered, vFields)
          .SetItem(BackOrderDetailFields.bodfIssued, vFields)
          .SetItem(BackOrderDetailFields.bodfEarliestDelivery, vFields)
          .SetItem(BackOrderDetailFields.bodfDespatchMethod, vFields)
          .SetItem(BackOrderDetailFields.bodfContactNumber, vFields)
          .SetItem(BackOrderDetailFields.bodfAddressNumber, vFields)
          .SetItem(BackOrderDetailFields.bodfSource, vFields)
          .SetItem(BackOrderDetailFields.bodfUnitPrice, vFields)
          .SetItem(BackOrderDetailFields.bodfGrossAmount, vFields)
          .SetItem(BackOrderDetailFields.bodfDiscount, vFields)
          .SetItem(BackOrderDetailFields.bodfVatRate, vFields)
          .SetItem(BackOrderDetailFields.bodfVatAmount, vFields)
          .SetItem(BackOrderDetailFields.bodfStatus, vFields)
          .SetOptionalItem(BackOrderDetailFields.bodfCurrencyUnitPrice, vFields)
          .SetOptionalItem(BackOrderDetailFields.bodfCurrencyVatAmount, vFields)
          .SetOptionalItem(BackOrderDetailFields.bodfWarehouse, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(BackOrderDetailFields.bodfAll)
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

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(BackOrderDetailFields.bodfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(BackOrderDetailFields.bodfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(BackOrderDetailFields.bodfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DespatchMethod() As String
      Get
        DespatchMethod = mvClassFields.Item(BackOrderDetailFields.bodfDespatchMethod).Value
      End Get
    End Property

    Public ReadOnly Property Discount() As Double
      Get
        Discount = mvClassFields.Item(BackOrderDetailFields.bodfDiscount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property EarliestDelivery() As String
      Get
        EarliestDelivery = mvClassFields.Item(BackOrderDetailFields.bodfEarliestDelivery).Value
      End Get
    End Property

    Public ReadOnly Property GrossAmount() As Double
      Get
        GrossAmount = mvClassFields.Item(BackOrderDetailFields.bodfGrossAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property Issued() As Integer
      Get
        Issued = mvClassFields.Item(BackOrderDetailFields.bodfIssued).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(BackOrderDetailFields.bodfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Ordered() As Integer
      Get
        Ordered = mvClassFields.Item(BackOrderDetailFields.bodfOrdered).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(BackOrderDetailFields.bodfProduct).Value
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(BackOrderDetailFields.bodfRate).Value
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(BackOrderDetailFields.bodfSource).Value
      End Get
    End Property

    Public ReadOnly Property Status() As String
      Get
        Status = mvClassFields.Item(BackOrderDetailFields.bodfStatus).Value
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(BackOrderDetailFields.bodfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property UnitPrice() As Double
      Get
        UnitPrice = mvClassFields.Item(BackOrderDetailFields.bodfUnitPrice).DoubleValue
      End Get
    End Property

    Public ReadOnly Property VatAmount() As Double
      Get
        VatAmount = mvClassFields.Item(BackOrderDetailFields.bodfVatAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property VatRate() As String
      Get
        VatRate = mvClassFields.Item(BackOrderDetailFields.bodfVatRate).Value
      End Get
    End Property

    Public ReadOnly Property CurrencyUnitPrice() As Double
      Get
        CurrencyUnitPrice = mvClassFields.Item(BackOrderDetailFields.bodfCurrencyUnitPrice).DoubleValue
      End Get
    End Property

    Public ReadOnly Property CurrencyVATAmount() As Double
      Get
        CurrencyVATAmount = mvClassFields.Item(BackOrderDetailFields.bodfCurrencyVatAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property Warehouse() As String
      Get
        Warehouse = mvClassFields.Item(BackOrderDetailFields.bodfWarehouse).Value
      End Get
    End Property

    Public Sub AllocateStock(ByVal pJobNumber As Integer)
      Dim vStockMovement As New StockMovement
      Dim vIssued As Integer
      Dim vIS As New IssuedStock
      Dim vProductCosts As ProductCosts
      Dim vPCostNumber As Integer

      vPCostNumber = 0
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then
        vProductCosts = New ProductCosts
        vProductCosts.InitFromProductAndWarehouse(mvEnv, Product, Warehouse)
        vPCostNumber = vProductCosts.GetEarliestProductCost.ProductCostNumber
      End If

      mvEnv.Connection.StartTransaction()
      'create a stock_movement record
      vStockMovement.Create(mvEnv, Product, -(Ordered - Issued), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonBackOrder), BatchNumber, TransactionNumber, LineNumber, False, Warehouse, vPCostNumber)
      vIssued = System.Math.Abs(vStockMovement.MovementQuantity)
      If vIssued > 0 Then
        'create an issued_stock record
        vIS.Init(mvEnv)
        vIS.Create(BatchNumber, TransactionNumber, LineNumber, Product, vIssued, Warehouse, pJobNumber)
        vIS.Save()
        'update this back_order_details record
        mvClassFields.Item(BackOrderDetailFields.bodfIssued).Value = CStr(mvClassFields.Item(BackOrderDetailFields.bodfIssued).IntegerValue + vIssued)
        Save()
      End If
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Function ReverseStockAllocation(ByVal pJobNumber As Integer) As Boolean
      Dim vIS As New IssuedStock
      Dim vStockMovement As New StockMovement
      Dim vRS As CDBRecordSet
      Dim vPCostNumber As Integer
      Dim vSQL As String

      'Reverse the allocation of stock - in one instance due to credit card authorisation failing
      vIS.Init(mvEnv, BatchNumber, TransactionNumber, LineNumber, pJobNumber)

      vPCostNumber = 0
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) = True And vIS.Existing = True Then
        'We need to put the stock back against the same ProductCost record as the original stock movement
        vSQL = "SELECT sm.product_cost_number FROM stock_movements sm WHERE batch_number = " & BatchNumber & " AND transaction_number = " & TransactionNumber & " AND line_number = "
        vSQL = vSQL & LineNumber & " AND warehouse = '" & Warehouse & "' AND stock_movement_reason = '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonBackOrder) & "' AND movement_quantity = " & (vIS.Issued * -1)
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        If vRS.Fetch() = True Then vPCostNumber = vRS.Fields(1).IntegerValue
        vRS.CloseRecordSet()
      End If

      If vIS.Existing Then
        mvEnv.Connection.StartTransaction()
        vStockMovement.Create(mvEnv, Product, (vIS.Issued), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonBackOrder), BatchNumber, TransactionNumber, LineNumber, False, Warehouse, vPCostNumber)
        'update this back_order_details record
        mvClassFields.Item(BackOrderDetailFields.bodfIssued).Value = CStr(mvClassFields.Item(BackOrderDetailFields.bodfIssued).IntegerValue - vIS.Issued)
        Save()
        vIS.Delete()
        mvEnv.Connection.CommitTransaction()
        ReverseStockAllocation = True
      End If
    End Function

    Public Sub SetAdjustment(ByVal pStatus As String, ByVal pOrdered As Integer, ByVal pIssued As Integer)

      If System.Math.Abs(Ordered) = System.Math.Abs(pOrdered) Then
        mvClassFields.Item(BackOrderDetailFields.bodfStatus).Value = pStatus
      Else
        mvClassFields.Item(BackOrderDetailFields.bodfOrdered).Value = CStr(Ordered + pOrdered)
        mvClassFields.Item(BackOrderDetailFields.bodfIssued).Value = CStr(Issued + pIssued)
      End If

    End Sub
  End Class
End Namespace
