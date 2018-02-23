Public Class frmStockMovement
    Dim mvTableTarget As DataTable
  Dim mvTableSource As DataTable
  Dim mvCloseForm As Boolean = False
Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Private Sub InitialiseControls()
      Try
      SetControlTheme()
      Dim vPanelItems As New PanelItems("epl")
      Dim vTable As DataTable
      AddHandler epl.GetCodeRestrictions, AddressOf epl_GetCodeRestrictions
      epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optStockMovement))
      vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStockMovementReasons)
      epl.FindComboBox("SourceWarehouse").Enabled = False
      epl.FindComboBox("SourceCostOfSale").Enabled = False
      epl.FindTextBox("SourceStockCount").Enabled=False

    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub


Private Sub cmdCancel_Click( ByVal sender As System.Object,  ByVal e As System.EventArgs) Handles cmdCancel.Click  
 me.Close()
End Sub


Private Sub cmdOk_Click( ByVal sender As System.Object,  ByVal e As System.EventArgs) Handles cmdOk.Click
    SaveStockMovement()
End Sub

Private Sub epl_ValueChanged( ByVal sender As System.Object,  ByVal pParameterName As System.String,  ByVal pValue As System.String) Handles epl.ValueChanged
  Select Case pParameterName
      Case "MovementReason"
        epl.SetValue("MovementReason", pValue)
        Dim vStockMovementReason as String = AppValues.ControlValue(AppValues.ControlTables.stock_movement_controls,AppValues.ControlValues.stock_warehouse_xfer_reason).ToString()
        If pValue = vStockMovementReason Then
          epl.FindComboBox("SourceWarehouse").Enabled = True
          epl.FindComboBox("SourceCostOfSale").Enabled = True
          epl.FindTextBox("SourceStockCount").Enabled=True
        Else
          epl.FindComboBox("SourceWarehouse").Enabled = False
          epl.FindComboBox("SourceCostOfSale").Enabled = False
          epl.FindTextBox("SourceStockCount").Enabled= False
        End If
        epl.SetErrorField("SourceWarehouse","")
        epl.SetErrorField("SourceCostOfSale","")
        epl.SetErrorField("SourceStockCount","")

      Case "Product"
        If pValue <> String.Empty Then
          Dim vTableCost As DataTable

          Dim vList As New ParameterList(True)
          epl.SetValue("CostOfSale", "")
          epl.SetValue("StockCount", "")
          epl.SetValue("TargetWarehouse", "")
          epl.SetValue("SourceWarehouse", "")
          vList("Product") = pValue
          vTableCost = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductCosts, vList)
          If Not vTableCost Is Nothing Then epl.SetValue("CostOfSale", vTableCost.Rows(0).Item("CostOfSale").ToString())
          mvTableTarget = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductWarehouses, vList)
          If Not mvTableTarget Is Nothing Then
            epl.SetComboDataSource("TargetWarehouse", "Warehouse", "WarehouseDesc", mvTableTarget, False)
            epl.SetValue("StockCount", mvTableTarget.Rows(0).Item("LastStockCount").ToString())
            mvTableSource = mvTableTarget.Copy()
            For Each vRow As DataRow In mvTableSource.Rows
              If vRow("Warehouse").ToString() = epl.GetValue("TargetWarehouse") Then
                vRow.Delete()
                mvTableSource.AcceptChanges()
                Exit For
              End If
            Next
            epl.SetComboDataSource("SourceWarehouse", "Warehouse", "WarehouseDesc", mvTableSource, False)
          End If
        End If
      Case "TargetWarehouse"
        If Not mvTableTarget Is Nothing Then
          mvTableSource = mvTableTarget.Copy()
          epl.SetValue("StockCount", mvTableTarget.Rows(epl.FindComboBox("TargetWarehouse").SelectedIndex).Item("LastStockCount").ToString())
          For Each vRow As DataRow In mvTableSource.Rows
            If vRow("Warehouse").ToString() = epl.GetValue("TargetWarehouse") Then
              vRow.Delete()
              mvTableSource.AcceptChanges()
              Exit For
            End If
          Next
          epl.SetComboDataSource("SourceWarehouse", "Warehouse", "WarehouseDesc", mvTableSource, False)
          epl.FindComboBox("SourceWarehouse").SelectedValue = 0
          epl.FindComboBox("SourceCostOfSale").SelectedValue = 0
          epl.SetValue("SourceStockCount", "")
        End If
      Case "SourceWarehouse"
        If Not mvTableSource Is Nothing Then
          Dim vTable As DataTable
          Dim vList As New ParameterList(True)
          vList("Product") = epl.GetValue("Product")
          vList("Warehouse") = epl.GetValue("SourceWarehouse")
          vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductCosts, vList)
          If Not vTable Is Nothing Then
            epl.SetComboDataSource("SourceCostOfSale", Nothing)
            If vTable.Columns.Contains("ProductCostNumber") Then
              epl.SetComboDataSource("SourceCostOfSale", "ProductCostNumber", "CostOfSale", vTable, True)
            Else
              epl.SetComboDataSource("SourceCostOfSale", "CostOfSale", "CostOfSale", vTable, True)
            End If
          Else
            epl.FindComboBox("SourceCostOfSale").DataSource = Nothing
          End If
          epl.SetErrorField("SourceWarehouse", "")
          epl.SetErrorField("SourceCostOfSale", "")
          epl.SetErrorField("SourceStockCount", "")
          epl.SetValue("SourceStockCount", mvTableSource.Rows(epl.FindComboBox("SourceWarehouse").SelectedIndex).Item("LastStockCount").ToString())
        End If
      Case "SourceCostOfSale"
        If Not mvTableSource Is Nothing Then
          Dim vSourceCostofSale As ComboBox = epl.FindComboBox(pParameterName)
          Dim vTable As DataTable
          Dim vList As New ParameterList(True)
          vList("Product") = epl.GetValue("Product")
          vList("Warehouse") = epl.GetValue("SourceWarehouse")
          vList("ProductCostNumber") = pValue
          vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductCosts, vList)
          If vTable IsNot Nothing Then
            epl.SetValue("SourceStockCount", vTable.Rows(0).Item("LastStockCount").ToString())
          End If
          If vSourceCostofSale.Text.Length > 0 Then epl.FindTextBox("CostOfSale").Text = vSourceCostofSale.Text
        End If
    End Select
End Sub

Private Sub epl_GetCodeRestrictions( ByVal sender As System.Object,  ByVal pParameterName As System.String,  ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
  Select Case pParameterName
      Case "Product"
        pList("PackProduct")="N"
        pList("StockItem")="Y"
  End Select
End Sub
  Private Sub SaveStockMovement()
    Dim vList As New ParameterList(True)
    If epl.AddValuesToList(vList, True) = False Then Exit Sub
    If ValidateControl() = False Then Exit Sub

    If ConfirmInsert() Then
      Try
        DataHelper.AddStockMovement(vList)
        epl.SetValue("StockCount", (IntegerValue(epl.GetValue("StockCount")) + IntegerValue(epl.GetValue("MovementQuantity"))).ToString)
        If ShowQuestion(QuestionMessages.QmAddStockMovement, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          epl.Clear()
          epl.FindComboBox("TargetWarehouse").DataSource = Nothing
          epl.FindTextLookupBox("Product").Focus()
          If epl.PanelInfo.PanelItems.Exists("SourceWarehouse") Then epl.FindComboBox("SourceWarehouse").DataSource = Nothing
          If epl.PanelInfo.PanelItems.Exists("SourceCostOfSale") Then epl.FindComboBox("SourceCostOfSale").DataSource = Nothing
          epl.EnableControlList("SourceWarehouse,SourceCostOfSale,SourceStockCount", False)
        Else
          mvCloseForm = True
          Me.Close()
        End If
      Catch vCareException As CareException
        Select Case vCareException.ErrorNumber
          Case CareException.ErrorNumbers.enInvalidStockMovement
            ShowWarningMessage(vCareException.Message)
          Case Else
            DataHelper.HandleException(vCareException)
        End Select
      End Try

    End If
  End Sub

  Private Function ValidateControl() As Boolean
    Dim vResult As Boolean = True
    If epl.GetValue("MovementReason") = "WX" Then
      If epl.GetValue("SourceWarehouse").Trim() = "" Then
        epl.SetErrorField("SourceWarehouse", GetInformationMessage(InformationMessages.ImFieldMandatory))
        vResult = False
      End If
      If epl.GetValue("SourceCostOfSale").Trim() = "" Then
        epl.SetErrorField("SourceCostOfSale", GetInformationMessage(InformationMessages.ImFieldMandatory))
        vResult = False
      End If
      If IntegerValue(epl.GetValue("SourceStockCount")) < IntegerValue(epl.GetValue("MovementQuantity")) Then
        epl.SetErrorField("SourceStockCount", GetInformationMessage(InformationMessages.ImInsufficientWarehouseStock))
        vResult = False
      End If
    End If
    Return vResult
  End Function

Private Sub frmStockMovement_FormClosing( ByVal sender As System.Object,  ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    If epl.DataChanged And Not mvCloseForm Then
      If ConfirmCancel() = False Then e.Cancel = True
    End If
End Sub
End Class