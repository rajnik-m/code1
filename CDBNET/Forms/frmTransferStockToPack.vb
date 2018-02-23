Public Class frmTransferStockToPack
  Private mvWareHouseTable As DataTable
  Private mvDataSetHead As DataSet
  Private mvDataSet As DataSet
  Public Sub New()

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Private Sub InitialiseControls()
    SetControlTheme()
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optTransferStockToPack))
    ProcessProducts()
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    Select Case pParameterName
      Case "Product"
        pList("PackProduct") = "Y"
        pList("StockItem") = "Y"
    End Select
  End Sub


  Private Sub epl_ValueChanged(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    Dim vList As New ParameterList(True)
    Dim vDataTable As DataTable
    Dim vCostTable As New DataTable

    Select Case pParameterName
      Case "Product"
        vList("Product") = pValue
        mvWareHouseTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductWarehouses, vList)
        If mvWareHouseTable IsNot Nothing Then
          epl.SetComboDataSource("TargetWarehouse", "Warehouse", "WarehouseDesc", mvWareHouseTable, False)
          If mvDataSet Is Nothing OrElse mvDataSet.Tables("DataRow") Is Nothing Then
            epl.SetValue("ProductCostOfSale", mvWareHouseTable.Rows(0).Item("CostOfSale").ToString)
          End If
        End If
        If epl.GetValue("Rate") = "" AndAlso dgr.RowCount > 0 Then
          dgr.ClearDataRows()
          dgr.DeleteRow(0)
        End If

      Case "TargetWarehouse"
          Dim vRow As DataRow() = mvWareHouseTable.Select("Warehouse = '" & pValue & "'")
          If vRow.Length > 0 Then
            epl.SetValue("StockCount", vRow(0).Item("LastStockCount").ToString())
          End If

      Case "Rate"
          vList("Rate") = epl.GetValue("Rate")
          vList("Product") = epl.GetValue("Product")
          mvDataSet = DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstPackedProductDataSheet, vList)
          vDataTable = DataHelper.GetTableFromDataSet(mvDataSet)

          ProcessProducts()

          If Not vDataTable Is Nothing Then
          dgr.Populate(mvDataSetHead)
        Else
          epl.SetValue("ProductCostOfSale", mvWareHouseTable.Rows(0).Item("CostOfSale").ToString)
          If dgr.RowCount > 0 Then
            dgr.ClearDataRows()
            dgr.DeleteRow(0)
          End If

        End If
          dgr.SetColumnVisible("DefaultWarehouse", False)
          dgr.SetColumnVisible("LinkProduct", False)
          PopulateWarehouse()
    End Select
  End Sub


  ''' <summary>
  ''' This function will create dataset of distinct product and display it on grid.
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ProcessProducts()
    Dim vTableHead As DataTable = DataHelper.NewColumnTable
    Dim vNewRow As DataRow
    Dim vTotalCost As Double

    DataHelper.AddDataColumn(vTableHead, "Product", "Product")
    DataHelper.AddDataColumn(vTableHead, "WarehouseDesc", "WarehouseDesc")
    DataHelper.AddDataColumn(vTableHead, "LastStockCount", "Last Stock Count")
    DataHelper.AddDataColumn(vTableHead, "OriginalCost", "Original Cost")
    DataHelper.AddDataColumn(vTableHead, "LinkProduct", "link_product")
    DataHelper.AddDataColumn(vTableHead, "DefaultWarehouse", "DefaultWarehouse")
    mvDataSetHead = New DataSet()
    mvDataSetHead.Tables.Add(vTableHead)
    Dim vNewTable As DataTable = mvDataSetHead.Tables.Add("DataRow")
    vNewTable.Columns.AddRange(New DataColumn() {New DataColumn("Product"), New DataColumn("Warehouse"), New DataColumn("LastStockCount"), New DataColumn("OriginalCost"), New DataColumn("LinkProduct"), New DataColumn("DefaultWarehouse")})

    vTotalCost = 0
    If mvDataSet IsNot Nothing Then
      'getting single row for a product
      If mvDataSet.Tables("DataRow") IsNot Nothing Then
        For Each vRow As DataRow In mvDataSet.Tables("DataRow").Rows
          If vNewTable.Select("LinkProduct ='" & vRow("linkproduct").ToString() & "'").Length = 0 Then
            vNewRow = vNewTable.NewRow()
            vNewRow("Product") = vRow("ProductDesc")
            vNewRow("Warehouse") = vRow("Warehouse")
            vNewRow("LinkProduct") = vRow("LinkProduct")
            vNewRow("DefaultWarehouse") = vRow("DefaultWarehouse")
            vNewRow("OriginalCost") = DoubleValue(vRow("CostOfSale").ToString()) + 0
            vTotalCost = vTotalCost + DoubleValue(vNewRow("OriginalCost").ToString())
            vNewTable.Rows.Add(vNewRow)
          End If
        Next
      End If
    Else
      'Populate data grid for initial form load
      dgr.Populate(mvDataSetHead)
      dgr.SetColumnVisible("DefaultWarehouse", False)
      dgr.SetColumnVisible("LinkProduct", False)
    End If
    epl.SetValue("ProductCostOfSale", vTotalCost.ToString())
  End Sub
  ''' <summary>
  ''' Create combo box of warehouse
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PopulateWarehouse()
    Dim vRow As DataRow()
    Dim vValues() As String = Nothing
    Dim vDataValues() As String = Nothing
    If mvDataSet.Tables("DataRow") IsNot Nothing Then
      For vIncr As Integer = 0 To dgr.RowCount - 1
        vRow = mvDataSet.Tables("DataRow").Select("ProductDesc ='" & dgr.GetValue(vIncr, 0) & "'")
        If vRow.Length > 0 Then
          For vIncrRow As Integer = 0 To vRow.Length - 1
            ReDim Preserve vValues(vIncrRow)
            ReDim Preserve vDataValues(vIncrRow)
            vValues(vIncrRow) = vRow(vIncrRow).Item("Warehouse").ToString()
            vDataValues(vIncrRow) = vRow(vIncrRow).Item("WarehouseDesc").ToString()
          Next
          dgr.SetComboBoxCell(vIncr, 1, vDataValues, vValues)
          dgr.SetCellsEditable()

        End If
      Next
    End If
  End Sub


  ''' <summary>
  ''' Displaying corrosponding LastStockCount based on Warehouse
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="pRow"></param>
  ''' <param name="pCol"></param>
  ''' <param name="pValue"></param>
  ''' <param name="pOldValue"></param>
  ''' <remarks></remarks>
  Private Sub dgr_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String) Handles dgr.ValueChanged
    Dim vRow As DataRow()
    If mvDataSet.Tables("DataRow") IsNot Nothing Then
      vRow = mvDataSet.Tables("DataRow").Select("LinkProduct='" & dgr.GetValue(pRow, 4) & "' AND Warehouse='" & pValue & "'")
      If vRow.Length > 0 Then
        For Each vDataRow As DataRow In mvDataSetHead.Tables("DataRow").Rows
          If vDataRow("LinkProduct").ToString() = vRow(0).Item("LinkProduct").ToString() Then
            vDataRow("LastStockCount") = vRow(0).Item("LastStockCount").ToString()
            vDataRow("Warehouse") = vRow(0).Item("Warehouse").ToString()
            dgr.SetValue(pRow, 2, vDataRow("LastStockCount").ToString())
            'Grid not refreshing when selected index of combobox in dgr is changed therefore removing focus from current control
            dgr.Focus()
            Exit For
          End If
        Next
      End If
    End If
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vList As New ParameterList(True)
    If epl.AddValuesToList(vList, True) Then
      For vIncr As Integer = 0 To dgr.RowCount - 1
        If dgr.GetValue(vIncr, 1).Length = 0 Then
          ShowErrorMessage(InformationMessages.ImWarehouseNotSelected, dgr.GetValue(vIncr, 0))
          Exit Sub
        End If
        If IntegerValue(dgr.GetValue(vIncr, 2)) < IntegerValue(epl.GetValue("MovementQuantity")) Then
          ShowErrorMessage(InformationMessages.ImNotEnoughStock, dgr.GetValue(vIncr, 0))
          Exit Sub
        End If
      Next
      'calling Actual save method 
      If ShowQuestion(QuestionMessages.QmConfirmInsert, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        DataHelper.AddStockToPack(vList)
        If dgr.RowCount > 0 Then
          For vIncr As Integer = 0 To dgr.RowCount - 1
            vList("Product") = dgr.GetValue(vIncr, 4)
            vList("Warehouse") = dgr.GetValue(vIncr, 1)
            DataHelper.AddStockToPack(vList)
          Next
        End If
        If ShowQuestion(QuestionMessages.QmAddStockMovement, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          epl.Clear()
          epl.FindComboBox("TargetWarehouse").DataSource = Nothing
          If dgr.RowCount > 0 Then dgr.ClearDataRows()

          epl.FindTextLookupBox("Product").Focus()
        Else
          Me.Close()
        End If
      End If
    End If
  End Sub
    
End Class