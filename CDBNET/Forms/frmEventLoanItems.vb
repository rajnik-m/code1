Public Class frmEventLoanItems
  Private mvProductDataTable As New DataTable

  Public Sub New(ByVal pEventNumber As Integer)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pEventNumber)
  End Sub

  Private Sub InitialiseControls(ByVal pEventNumber As Integer)
    SetControlTheme()
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optEventLoanItems))

    Dim vProductDataSet As DataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventLoanItems, pEventNumber)
    If Not vProductDataSet Is Nothing Then
      mvProductDataTable = DataHelper.GetTableFromDataSet(vProductDataSet)
      dgr.Populate(vProductDataSet)
      If dgr.RowCount > 0 Then dgr.SelectRow(0)
      epl.DataChanged = False 'Data has not yet changed
    End If
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pDataRow As System.Int32) Handles dgr.RowSelected
    Dim vContactNumber As Integer = IntegerValue(mvProductDataTable.Rows(pDataRow).Item("ContactNumber"))
    Dim vAddressNumber As Integer = IntegerValue(mvProductDataTable.Rows(pDataRow).Item("AddressNumber"))
    Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber)
    If Not vRow Is Nothing Then
      epl.SetValue("ContactName", vRow.Item("ContactName").ToString)
      epl.SetValue("AddressLine", vRow.Item("AddressLine").ToString)
    End If

    Dim vDateReturned As String = mvProductDataTable.Rows(pDataRow).Item("Returned").ToString
    If vDateReturned = String.Empty Then
      epl.FindDateTimePicker("Returned").Checked = False
    Else
      epl.FindDateTimePicker("Returned").Checked = True
      epl.SetValue("Returned", vDateReturned)
    End If

    Dim vQuantityReturned As String = mvProductDataTable.Rows(pDataRow).Item("QuantityReturned").ToString
    epl.SetValue("Quantity", vQuantityReturned)

    Dim vReference As String = mvProductDataTable.Rows(pDataRow).Item("Reference").ToString
    epl.SetValue("Reference", vReference)

    Dim vCompleted As String = mvProductDataTable.Rows(pDataRow).Item("Complete").ToString
    epl.SetValue("AcceptAsComplete", vCompleted)
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pValue As System.String) Handles epl.ValueChanged
    If dgr.RowCount > 0 Then
      Dim vRow As DataRow = mvProductDataTable.Rows(dgr.CurrentDataRow)
      Select Case pParameterName
        Case "Returned"
          If epl.FindDateTimePicker(pParameterName).Checked = True Then
            vRow.Item("Returned") = pValue
          Else
            vRow.Item("Returned") = String.Empty
          End If
        Case "Quantity"
          vRow.Item("QuantityReturned") = pValue
          If IntegerValue(pValue) >= IntegerValue(vRow.Item("Quantity").ToString) Then
            epl.SetValue("AcceptAsComplete", "Y")
            vRow.Item("Complete") = "Y"
          Else
            epl.SetValue("AcceptAsComplete", "N")
            vRow.Item("Complete") = "N"
          End If
        Case "Reference"
          vRow.Item("Reference") = pValue
        Case "AcceptAsComplete"
          If epl.FindCheckBox(pParameterName).Checked Then
            vRow.Item("Complete") = "Y"
          Else
            vRow.Item("Complete") = "N"
          End If
      End Select
    End If
  End Sub

  Private Sub cmdReturned_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReturned.Click
    If dgr.RowCount > 0 Then
      Dim vRow As DataRow = mvProductDataTable.Rows(dgr.CurrentDataRow)
      Dim vQuantity As String = vRow.Item("Quantity").ToString
      epl.SetValue("Quantity", vQuantity)
      vRow.Item("QuantityReturned") = vQuantity
      epl.SetValue("AcceptAsComplete", "Y")
      vRow.Item("Complete") = "Y"
      epl.SetValue("Returned", Today.ToString(AppValues.DateFormat))
      vRow.Item("Returned") = Today.ToString(AppValues.DateFormat)
    End If
  End Sub

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      If dgr.RowCount > 0 AndAlso epl.DataChanged Then
        If ProcessSave() Then
          epl.DataChanged = False 'allow the for to close
          Me.Close()
        End If
      Else
        Me.Close()
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function ProcessSave() As Boolean
    Dim vParamList As New ParameterList(True)
    Dim vList As New ParameterList(True)
    Dim vSuccessful As Boolean
    If epl.AddValuesToList(vList, True) Then
      For vDataRowNumber As Integer = 0 To mvProductDataTable.Rows.Count - 1
        Dim vDataRow As DataRow = mvProductDataTable.Rows(vDataRowNumber)
        vList("AddressNumber") = vDataRow.Item("AddressNumber").ToString
        vList("ContactNumber") = vDataRow.Item("ContactNumber").ToString
        vList("ProductCode") = vDataRow.Item("ProductCode").ToString
        vList("Issued") = vDataRow.Item("Issued").ToString
        vList("Quantity") = vDataRow.Item("Quantity").ToString
        vList("Returned") = vDataRow.Item("Returned").ToString
        vList("AcceptAsComplete") = vDataRow.Item("Complete").ToString
        vList("Reference") = vDataRow.Item("Reference").ToString
        DataHelper.UpdateEventLoanItem(vList)
      Next
      vSuccessful = True
    End If
    Return vSuccessful
  End Function

  Private Sub frmEventLoanItems_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    If epl.DataChanged Then
      If ConfirmSave() Then
        ProcessSave()
      End If
    End If
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub
End Class