

Namespace Access
  Public Class ProductCost

    Public Enum ProductCostRecordSetTypes 'These are bit values
      pcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ProductCostFields
      pcfAll = 0
      pcfProductCostNumber
      pcfProduct
      pcfWarehouse
      pcfCostOfSale
      pcfOriginalQuantity
      pcfLastStockCount
      pcfAmendedBy
      pcfAmendedOn
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
          .DatabaseTableName = "product_costs"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("product_cost_number", CDBField.FieldTypes.cftLong)
          .Add("product")
          .Add("warehouse")
          .Add("cost_of_sale", CDBField.FieldTypes.cftNumeric)
          .Add("original_quantity", CDBField.FieldTypes.cftLong)
          .Add("last_stock_count", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(ProductCostFields.pcfProductCostNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ProductCostFields)
      'Add code here to ensure all values are valid before saving
      If mvExisting = False And mvClassFields.Item(ProductCostFields.pcfProductCostNumber).IntegerValue = 0 Then mvClassFields.Item(ProductCostFields.pcfProductCostNumber).Value = CStr(mvEnv.GetControlNumber("PC"))
      mvClassFields.Item(ProductCostFields.pcfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ProductCostFields.pcfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ProductCostRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ProductCostRecordSetTypes.pcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pProductCostNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pProductCostNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ProductCostRecordSetTypes.pcrtAll) & " FROM product_costs pc WHERE product_cost_number = " & pProductCostNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ProductCostRecordSetTypes.pcrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ProductCostRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ProductCostFields.pcfProductCostNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ProductCostRecordSetTypes.pcrtAll) = ProductCostRecordSetTypes.pcrtAll Then
          .SetItem(ProductCostFields.pcfProduct, vFields)
          .SetItem(ProductCostFields.pcfWarehouse, vFields)
          .SetItem(ProductCostFields.pcfCostOfSale, vFields)
          .SetItem(ProductCostFields.pcfOriginalQuantity, vFields)
          .SetItem(ProductCostFields.pcfLastStockCount, vFields)
          .SetItem(ProductCostFields.pcfAmendedBy, vFields)
          .SetItem(ProductCostFields.pcfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ProductCostFields.pcfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub AddNewStock(ByVal pQuantity As Integer)
      'Just update OriginalQuantity as LastStockCount will be set when adding the StockMovement
      mvClassFields.Item(ProductCostFields.pcfOriginalQuantity).Value = CStr(mvClassFields.Item(ProductCostFields.pcfOriginalQuantity).IntegerValue + pQuantity)
      If mvClassFields.Item(ProductCostFields.pcfOriginalQuantity).LongValue < 0 Then mvClassFields.Item(ProductCostFields.pcfOriginalQuantity).Value = CStr(0)
    End Sub

    Public Sub SellStock(ByVal pQuantity As Integer)
      'Update LastStockCount
      mvClassFields.Item(ProductCostFields.pcfLastStockCount).Value = CStr(mvClassFields.Item(ProductCostFields.pcfLastStockCount).IntegerValue + pQuantity)
    End Sub

    Friend Sub SetShortfall()
      'Set LastStockCount to zero
      mvClassFields.Item(ProductCostFields.pcfLastStockCount).Value = CStr(0)
    End Sub

    Friend Sub Create(ByVal pEnv As CDBEnvironment, ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pCostOfSale As Double, ByVal pOriginalQuantity As Integer)
      Init(pEnv)
      With mvClassFields
        .Item(ProductCostFields.pcfProduct).Value = pProductCode
        .Item(ProductCostFields.pcfWarehouse).Value = pWarehouseCode
        .Item(ProductCostFields.pcfCostOfSale).Value = CStr(pCostOfSale)
        .Item(ProductCostFields.pcfOriginalQuantity).Value = CStr(pOriginalQuantity)
        .Item(ProductCostFields.pcfLastStockCount).Value = CStr(0) 'This is set to the correct figure by the initial StockMovement
      End With
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
        AmendedBy = mvClassFields.Item(ProductCostFields.pcfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ProductCostFields.pcfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CostOfSale() As Double
      Get
        CostOfSale = mvClassFields.Item(ProductCostFields.pcfCostOfSale).DoubleValue
      End Get
    End Property

    Public ReadOnly Property LastStockCount() As Integer
      Get
        LastStockCount = mvClassFields.Item(ProductCostFields.pcfLastStockCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OriginalQuantity() As Integer
      Get
        OriginalQuantity = mvClassFields.Item(ProductCostFields.pcfOriginalQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(ProductCostFields.pcfProduct).Value
      End Get
    End Property

    Public ReadOnly Property ProductCostNumber() As Integer
      Get
        ProductCostNumber = mvClassFields.Item(ProductCostFields.pcfProductCostNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property WarehouseCode() As String
      Get
        WarehouseCode = mvClassFields.Item(ProductCostFields.pcfWarehouse).Value
      End Get
    End Property
  End Class
End Namespace
