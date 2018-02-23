

Namespace Access
  Public Class ProductWarehouse

    Public Enum ProductWarehouseRecordSetTypes 'These are bit values
      pwrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ProductWarehouseFields
      pwfAll = 0
      pwfProduct
      pwfWarehouse
      pwfBinNumber
      pwfLastStockCount
      pwfAmendedBy
      pwfAmendedOn
      pwfQuantityOnOrder
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
          .DatabaseTableName = "product_warehouses"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("product")
          .Add("warehouse")
          .Add("bin_number")
          .Add("last_stock_count", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("quantity_on_order", CDBField.FieldTypes.cftLong)
        End With

        mvClassFields.Item(ProductWarehouseFields.pwfProduct).SetPrimaryKeyOnly()
        mvClassFields.Item(ProductWarehouseFields.pwfWarehouse).SetPrimaryKeyOnly()
        mvClassFields.Item(ProductWarehouseFields.pwfProduct).PrefixRequired = True
        mvClassFields.Item(ProductWarehouseFields.pwfWarehouse).PrefixRequired = True
        mvClassFields.Item(ProductWarehouseFields.pwfBinNumber).PrefixRequired = True
        mvClassFields.Item(ProductWarehouseFields.pwfLastStockCount).PrefixRequired = True
        mvClassFields.Item(ProductWarehouseFields.pwfQuantityOnOrder).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ProductWarehouseFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ProductWarehouseFields.pwfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ProductWarehouseFields.pwfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ProductWarehouseRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ProductWarehouseRecordSetTypes.pwrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pw")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pProduct As String = "", Optional ByRef pWarehouse As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pProduct) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ProductWarehouseRecordSetTypes.pwrtAll) & " FROM product_warehouses pw WHERE product = '" & pProduct & "' AND warehouse = '" & pWarehouse & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ProductWarehouseRecordSetTypes.pwrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ProductWarehouseRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ProductWarehouseFields.pwfProduct, vFields)
        .SetItem(ProductWarehouseFields.pwfWarehouse, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ProductWarehouseRecordSetTypes.pwrtAll) = ProductWarehouseRecordSetTypes.pwrtAll Then
          .SetItem(ProductWarehouseFields.pwfBinNumber, vFields)
          .SetItem(ProductWarehouseFields.pwfLastStockCount, vFields)
          .SetItem(ProductWarehouseFields.pwfAmendedBy, vFields)
          .SetItem(ProductWarehouseFields.pwfAmendedOn, vFields)
          .SetOptionalItem(ProductWarehouseFields.pwfQuantityOnOrder, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ProductWarehouseFields.pwfAll)
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
        AmendedBy = mvClassFields.Item(ProductWarehouseFields.pwfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ProductWarehouseFields.pwfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BinNumber() As String
      Get
        BinNumber = mvClassFields.Item(ProductWarehouseFields.pwfBinNumber).Value
      End Get
    End Property

    Public Property LastStockCount() As Integer
      Get
        LastStockCount = mvClassFields.Item(ProductWarehouseFields.pwfLastStockCount).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(ProductWarehouseFields.pwfLastStockCount).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(ProductWarehouseFields.pwfProduct).Value
      End Get
    End Property
    Public Property QuantityOnOrder() As Integer
      Get
        QuantityOnOrder = mvClassFields.Item(ProductWarehouseFields.pwfQuantityOnOrder).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(ProductWarehouseFields.pwfQuantityOnOrder).IntegerValue = Value
      End Set
    End Property
    Public ReadOnly Property Warehouse() As String
      Get
        Warehouse = mvClassFields.Item(ProductWarehouseFields.pwfWarehouse).Value
      End Get
    End Property
  End Class
End Namespace
