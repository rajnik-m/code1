

Namespace Access
  Public Class PurchaseOrderDetail

    Public Enum PurchaseOrderDetailRecordSetTypes 'These are bit values
      podrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PurchaseOrderDetailFields
      podfAll = 0
      podfPurchaseOrderNumber
      podfLineNumber
      podfLineItem
      podfLinePrice
      podfQuantity
      podfAmount
      podfAmendedBy
      podfAmendedOn
      podfNominalAccount
      podfDistributionCode
      podfBalance
      podfProductCode
      podfWarehouseCode
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvProductWarehouse As New ProductWarehouse

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "purchase_order_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("purchase_order_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("line_item")
          .Add("line_price", CDBField.FieldTypes.cftNumeric)
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("nominal_account")
          .Add("distribution_code")
          .Add("balance", CDBField.FieldTypes.cftNumeric)
          .Add("product", CDBField.FieldTypes.cftCharacter)
          .Add("warehouse", CDBField.FieldTypes.cftCharacter)

          .Item(PurchaseOrderDetailFields.podfPurchaseOrderNumber).SetPrimaryKeyOnly()
          .Item(PurchaseOrderDetailFields.podfLineNumber).SetPrimaryKeyOnly()

          .Item(PurchaseOrderDetailFields.podfPurchaseOrderNumber).PrefixRequired = True
          .Item(PurchaseOrderDetailFields.podfProductCode).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink)
          .Item(PurchaseOrderDetailFields.podfWarehouseCode).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink)
        End With
      Else
        mvClassFields.ClearItems()
        mvProductWarehouse = New ProductWarehouse
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      'UPGRADE_NOTE: Object mvProductWarehouse may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvProductWarehouse = Nothing
      mvProductWarehouse = New ProductWarehouse
      mvProductWarehouse.Init(mvEnv)
    End Sub

    Private Sub SetValid(ByVal pField As PurchaseOrderDetailFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(PurchaseOrderDetailFields.podfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PurchaseOrderDetailFields.podfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PurchaseOrderDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PurchaseOrderDetailRecordSetTypes.podrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pod")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pPurchaseOrderNumber As Integer = 0, Optional ByVal pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pPurchaseOrderNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PurchaseOrderDetailRecordSetTypes.podrtAll) & " FROM purchase_order_details pod WHERE purchase_order_number = " & pPurchaseOrderNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PurchaseOrderDetailRecordSetTypes.podrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PurchaseOrderDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PurchaseOrderDetailFields.podfPurchaseOrderNumber, vFields)
        .SetItem(PurchaseOrderDetailFields.podfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PurchaseOrderDetailRecordSetTypes.podrtAll) = PurchaseOrderDetailRecordSetTypes.podrtAll Then
          .SetItem(PurchaseOrderDetailFields.podfLineItem, vFields)
          .SetItem(PurchaseOrderDetailFields.podfLinePrice, vFields)
          .SetItem(PurchaseOrderDetailFields.podfQuantity, vFields)
          .SetItem(PurchaseOrderDetailFields.podfAmount, vFields)
          .SetItem(PurchaseOrderDetailFields.podfAmendedBy, vFields)
          .SetItem(PurchaseOrderDetailFields.podfAmendedOn, vFields)
          .SetOptionalItem(PurchaseOrderDetailFields.podfNominalAccount, vFields)
          .SetOptionalItem(PurchaseOrderDetailFields.podfDistributionCode, vFields)
          .SetOptionalItem(PurchaseOrderDetailFields.podfBalance, vFields)
          .SetOptionalItem(PurchaseOrderDetailFields.podfProductCode, vFields)
          .SetOptionalItem(PurchaseOrderDetailFields.podfWarehouseCode, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vTransactionStarted As Boolean

      SetValid(PurchaseOrderDetailFields.podfAll)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) = True And Len(ProductCode) > 0 And Len(WarehouseCode) > 0 Then
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vTransactionStarted = True
        End If
        With ProductWarehouse
          If .Existing Then
            .QuantityOnOrder = .QuantityOnOrder + Quantity
            .Save()
          End If
        End With
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      If vTransactionStarted Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Sub InitFromPO(ByVal pEnv As CDBEnvironment, ByVal pPO As PurchaseOrder, ByVal pLineItem As String, ByVal pLinePrice As Double, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pBalance As Double, ByVal pAccount As String, ByVal pDistributionCode As String, Optional ByVal pProductCode As String = "", Optional ByVal pWarehouseCode As String = "")
      Init(pEnv)
      mvClassFields(PurchaseOrderDetailFields.podfPurchaseOrderNumber).Value = CStr(pPO.PurchaseOrderNumber)
      mvClassFields(PurchaseOrderDetailFields.podfLineNumber).Value = CStr(pPO.Details.Count() + 1)
      mvClassFields(PurchaseOrderDetailFields.podfLineItem).Value = pLineItem
      mvClassFields(PurchaseOrderDetailFields.podfLinePrice).Value = CStr(pLinePrice)
      mvClassFields(PurchaseOrderDetailFields.podfQuantity).Value = CStr(pQuantity)
      mvClassFields(PurchaseOrderDetailFields.podfAmount).Value = CStr(pAmount)
      mvClassFields(PurchaseOrderDetailFields.podfBalance).Value = CStr(pBalance)
      mvClassFields(PurchaseOrderDetailFields.podfNominalAccount).Value = pAccount
      mvClassFields(PurchaseOrderDetailFields.podfDistributionCode).Value = pDistributionCode
      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        mvClassFields(PurchaseOrderDetailFields.podfProductCode).Value = pProductCode
        mvClassFields(PurchaseOrderDetailFields.podfWarehouseCode).Value = pWarehouseCode
      End If
    End Sub
    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pParams As CDBParameters)
      Dim vBalance As Double
      If Amount > 0 Then
        vBalance = Balance + pParams("Amount").DoubleValue - Amount 'Make sure the balance is updated correctly for existing detail lines or when amending a new detail line
      Else
        vBalance = pParams("Amount").DoubleValue  'For new detail lines
      End If
      Init(pEnv)
      mvClassFields(PurchaseOrderDetailFields.podfPurchaseOrderNumber).Value = pParams.OptionalValue("PurchaseOrderNumber", "0")
      mvClassFields(PurchaseOrderDetailFields.podfLineNumber).IntegerValue = pParams("LineNumber").IntegerValue
      mvClassFields(PurchaseOrderDetailFields.podfLineItem).Value = pParams("LineItem").Value
      mvClassFields(PurchaseOrderDetailFields.podfLinePrice).DoubleValue = pParams("LinePrice").DoubleValue
      mvClassFields(PurchaseOrderDetailFields.podfQuantity).IntegerValue = pParams("Quantity").IntegerValue
      mvClassFields(PurchaseOrderDetailFields.podfAmount).DoubleValue = pParams("Amount").DoubleValue
      mvClassFields(PurchaseOrderDetailFields.podfBalance).DoubleValue = vBalance
      mvClassFields(PurchaseOrderDetailFields.podfNominalAccount).Value = pParams.ParameterExists("NominalAccount").Value
      mvClassFields(PurchaseOrderDetailFields.podfDistributionCode).Value = pParams.ParameterExists("DistributionCode").Value
      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        mvClassFields(PurchaseOrderDetailFields.podfProductCode).Value = pParams.ParameterExists("Product").Value
        mvClassFields(PurchaseOrderDetailFields.podfWarehouseCode).Value = pParams.ParameterExists("Warehouse").Value
      End If
    End Sub

    Public Function LineDataType(ByVal pAttributeName As String) As CDBField.FieldTypes
      LineDataType = mvClassFields.ItemDataType(pAttributeName)
    End Function
    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        mvClassFields.ItemValue(pAttributeName) = Value
      End Set
    End Property

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
        AmendedBy = mvClassFields.Item(PurchaseOrderDetailFields.podfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PurchaseOrderDetailFields.podfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(PurchaseOrderDetailFields.podfAmount).DoubleValue
      End Get
    End Property

    Public Property Balance() As Double
      Get
        Balance = mvClassFields.Item(PurchaseOrderDetailFields.podfBalance).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(PurchaseOrderDetailFields.podfBalance).DoubleValue = Value
      End Set
    End Property
    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvClassFields.Item(PurchaseOrderDetailFields.podfDistributionCode).Value
      End Get
    End Property
    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(PurchaseOrderDetailFields.podfProductCode).Value
      End Get
    End Property
    Public ReadOnly Property ProductWarehouse() As ProductWarehouse
      Get
        If ProductCode.Length > 0 And WarehouseCode.Length > 0 Then
          If mvProductWarehouse.Existing = False Then
            mvProductWarehouse.Init(mvEnv, ProductCode, WarehouseCode)
          End If
        End If
        Return mvProductWarehouse
      End Get
    End Property
    Public ReadOnly Property WarehouseCode() As String
      Get
        WarehouseCode = mvClassFields.Item(PurchaseOrderDetailFields.podfWarehouseCode).Value
      End Get
    End Property
    Public ReadOnly Property LineItem() As String
      Get
        LineItem = mvClassFields.Item(PurchaseOrderDetailFields.podfLineItem).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(PurchaseOrderDetailFields.podfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LinePrice() As Double
      Get
        LinePrice = mvClassFields.Item(PurchaseOrderDetailFields.podfLinePrice).DoubleValue
      End Get
    End Property

    Public ReadOnly Property NominalAccount() As String
      Get
        NominalAccount = mvClassFields.Item(PurchaseOrderDetailFields.podfNominalAccount).Value
      End Get
    End Property

    Public ReadOnly Property PurchaseOrderNumber() As Integer
      Get
        PurchaseOrderNumber = mvClassFields.Item(PurchaseOrderDetailFields.podfPurchaseOrderNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(PurchaseOrderDetailFields.podfQuantity).IntegerValue
      End Get
    End Property

    Public Function GetDataAsParameters() As CDBParameters
      Dim vParams As New CDBParameters
      Dim vField As ClassField

      For Each vField In mvClassFields
        If vField.Name <> "amended_by" And vField.Name <> "amended_on" Then vParams.Add(ProperName((vField.Name)), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(DoubleValue(vField.Value)), vField.Value))
      Next vField
      GetDataAsParameters = vParams
    End Function
    Public Sub Delete()
      Dim vTransactionStarted As Boolean

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) = True And Len(ProductCode) > 0 And Len(WarehouseCode) > 0 Then
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vTransactionStarted = True
        End If
        With ProductWarehouse
          If .Existing Then
            If Quantity > .QuantityOnOrder Then
              .QuantityOnOrder = 0
            Else
              .QuantityOnOrder = .QuantityOnOrder - Quantity
            End If
            .Save(mvEnv.User.UserID, True)
          End If
        End With
      End If
      mvClassFields.Delete(mvEnv.Connection, mvEnv, mvEnv.User.UserID, True)
      If vTransactionStarted Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Friend Sub UpdateForRegularPayments(ByVal pAmount As Double)
      'Used when authorising a regular payment
      mvClassFields(PurchaseOrderDetailFields.podfAmount).DoubleValue += pAmount
      mvClassFields(PurchaseOrderDetailFields.podfBalance).DoubleValue = pAmount
    End Sub
    Public Sub Update(ByVal pParameterList As CDBParameters)
      Update(pParameterList, False)
    End Sub

    Private Sub Update(ByVal pParameterList As CDBParameters, ByVal pValidate As Boolean)

      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False AndAlso (vClassField.NonUpdatable = False OrElse pValidate = False) Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next

    End Sub
  End Class
End Namespace
