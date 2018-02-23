

Namespace Access
  Public Class ProductCosts
    Implements System.Collections.IEnumerable

    Private mvCol As New Collection

    Private mvEnv As CDBEnvironment
    Private mvProductCode As String
    Private mvWarehouseCode As String

    Public Function Add(ByVal pCostOfSale As Double, ByVal pOriginalQuantity As Integer) As ProductCost
      'create a new object
      Dim vProductCost As New ProductCost

      vProductCost.Create(mvEnv, mvProductCode, mvWarehouseCode, pCostOfSale, pOriginalQuantity)
      vProductCost.Save()

      mvCol.Add(vProductCost, CStr(vProductCost.ProductCostNumber))
      Add = vProductCost
      'UPGRADE_NOTE: Object vProductCost may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      vProductCost = Nothing
    End Function

    Public Sub InitFromProductAndWarehouse(ByVal pEnv As CDBEnvironment, ByVal pProductCode As String, ByVal pWarehouseCode As String, Optional ByVal pWithStockOnly As Boolean = True)
      Dim vRS As CDBRecordSet
      Dim vProductCost As New ProductCost
      Dim vSQL As String

      mvEnv = pEnv
      mvProductCode = pProductCode
      mvWarehouseCode = pWarehouseCode

      vProductCost.Init(mvEnv)
      vSQL = "SELECT " & vProductCost.GetRecordSetFields(ProductCost.ProductCostRecordSetTypes.pcrtAll) & " FROM product_costs pc WHERE product = '" & mvProductCode & "' AND warehouse = '" & mvWarehouseCode & "'"
      If pWithStockOnly Then vSQL = vSQL & " AND last_stock_count > 0"
      vSQL = vSQL & " ORDER BY product_cost_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vProductCost = New ProductCost
        vProductCost.InitFromRecordSet(mvEnv, vRS, ProductCost.ProductCostRecordSetTypes.pcrtAll)
        mvCol.Add(vProductCost, CStr(vProductCost.ProductCostNumber))
      End While
      vRS.CloseRecordSet()

      If mvCol.Count() = 0 Then
        'No records found so add a dummy record
        vProductCost = New ProductCost
        vProductCost.Init(mvEnv)
        mvCol.Add(vProductCost)
      End If

    End Sub

    Public Sub SetShortfall()
      'There is s shortfall so set LastStockCount to 0 on all items
      Dim vProductCost As ProductCost

      For Each vProductCost In mvCol
        vProductCost.SetShortfall()
        vProductCost.Save()
      Next vProductCost

    End Sub

    Public ReadOnly Property Item(ByVal pIndexKey As Integer) As ProductCost
      Get
        Return CType(mvCol.Item(pIndexKey), ProductCost)
      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public Function GetEarliestProductCost() As ProductCost
      'Always return the first item in the collection
      GetEarliestProductCost = CType(mvCol.Item(1), ProductCost)
    End Function

    Public Function GetLatestProductCost() As ProductCost
      'Always return last item in the collection
      GetLatestProductCost = CType(mvCol.Item(mvCol.Count()), ProductCost)
    End Function
  End Class
End Namespace
