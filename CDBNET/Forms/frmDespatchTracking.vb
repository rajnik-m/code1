Public Class frmDespatchTracking

  Dim mvDataTable As DataTable = Nothing
  Dim mvDespatchNoteNumber As Integer = 0

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()
    AddHandler txtPickingListNumber.KeyPress, AddressOf IntegerKeyPressHandler
    AddHandler txtPickingListNumber.TextChanged, AddressOf IntegerTextChangedHandler

    Dim vPanelItems As New PanelItems("epl")
    Dim mvTmpDataSet As New DataSet
    Dim mvTmpDataTable As DataTable = DataHelper.NewColumnTable

    Try

      SetControlTheme()      
      epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optDespatchTracking))
      'setting default columns for display grid
      DataHelper.AddDataColumn(mvTmpDataTable, "DespatchNoteNumber","Despatch Note Number")
      DataHelper.AddDataColumn(mvTmpDataTable, "DespatchMethod","Despatch Method")
      DataHelper.AddDataColumn(mvTmpDataTable, "DespatchDate", "Despatch Date")
      DataHelper.AddDataColumn(mvTmpDataTable, "Delivery", "Delivery")
      mvTmpDataSet.Tables.Add(mvTmpDataTable)
      dgr.Populate(mvTmpDataSet)
      txtPickingListNumber.Focus()
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try


  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    txtPickingListNumber.Text = String.Empty
    txtWarehouse.Text = String.Empty
    txtWarehouseDesc.Text = String.Empty
    dgr.ClearDataRows()
    epl.Clear()
    cmdSave.Enabled = False
    txtPickingListNumber.Focus()
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub txtPickingListNumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPickingListNumber.Leave
    mvDespatchNoteNumber = 0
    If txtPickingListNumber.Text.Length > 0 Then
      Dim vList As ParameterList = New ParameterList(True)
      vList("PickingListNumber") = txtPickingListNumber.Text
      GetWarehouseInfo(vList)
      GetDespachNote()
    Else
      txtWarehouse.Text = String.Empty
      txtWarehouseDesc.Text = String.Empty
      epl.Clear()
      dgr.ClearDataRows()
      cmdSave.Enabled = False
    End If
  End Sub
  
  Private Sub dgr_RowSelected(ByVal sender As System.Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    If Not mvDataTable Is Nothing Then
      SetDespatchNoteDetails(mvDataTable.Rows(pDataRow))
      mvDespatchNoteNumber = IntegerValue(mvDataTable.Rows(pDataRow).Item("DespatchNoteNumber").ToString)
    End If
  End Sub

  Private Sub GetWarehouseInfo(ByVal pList As ParameterList)
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPickingListWarehouses, pList)
    If Not vDataTable Is Nothing Then
      txtWarehouse.Text = vDataTable.Rows(0).Item("Warehouse").ToString()
      txtWarehouseDesc.Text = vDataTable.Rows(0).Item("WarehouseDesc").ToString()
    Else
      txtWarehouse.Text = String.Empty
      txtWarehouseDesc.Text = String.Empty
    End If
  End Sub

  ''' <summary>
  ''' Setting Details for Each Despatch Note for edit
  ''' </summary>
  ''' <param name="vRow"></param>
  ''' <remarks></remarks>
  Private Sub SetDespatchNoteDetails(ByVal vRow As DataRow)
      cmdSave.Enabled = True
      epl.Populate(vRow)       
  End Sub
  ''' <summary>
  ''' Get Despatch details based on Picking List value.
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetDespachNote()
    Try
      Dim vList As New ParameterList(True)
      vList("PickingListNumber") = txtPickingListNumber.Text
      Dim mvDataSet As DataSet = Nothing
      mvDataSet = DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstDespatchData, vList)
      mvDataTable = DataHelper.GetTableFromDataSet(mvDataSet)
      If Not mvDataTable Is Nothing Then
        dgr.Populate(mvDataSet)
        If (dgr.RowCount > 0) Then
          Dim vSelectedRow As Integer = 0
          If mvDespatchNoteNumber > 0 Then vSelectedRow = dgr.FindRow("DespatchNoteNumber", mvDespatchNoteNumber.ToString)
          If vSelectedRow < 0 Then vSelectedRow = 0
          dgr.SelectRow(vSelectedRow)
        End If
      Else
        dgr.ClearDataRows()
        epl.Clear()
        cmdSave.Enabled = False
        txtPickingListNumber.Text = ""
      End If

    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  ''' <summary>
  ''' Saving Despatch details
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SaveDespatchNote()
    Try
      Dim vList As New ParameterList(True)
      If txtPickingListNumber.Text.Trim().Length = 0 Then
        MessageBox.Show("Please enter picking list number")
        txtPickingListNumber.Focus()
        Exit Sub
      End If
      If epl.AddValuesToList(vList,True,EditPanel.AddNullValueTypes.anvtAll,True ) Then
        vList("PickingListNumber") = txtPickingListNumber.Text.Trim()
        vList("DespatchNoteNumber") = epl.GetValue("DespatchNoteNumber")
        DataHelper.UpdateDespatchNote(vList)
        GetDespachNote()
        MessageBox.Show("Despatch Note Updated Successfully")
      end if
     Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      ElseIf vException.ErrorNumber = CareException.ErrorNumbers.enParameterInvalidValue Then
        ShowInformationMessage(vException.Message)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub


  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    SaveDespatchNote()
  End Sub

End Class