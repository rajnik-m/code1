Public Class frmStockShortfall

  Private mvPickingListNumber As Integer
  Private mvList As ParameterList

  Public Sub New(ByVal pList As ParameterList)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pList)
  End Sub

  Private Sub InitialiseControls(ByVal pList As ParameterList)
    SetControlTheme()
    mvList = pList
    mvPickingListNumber = mvList.IntegerValue("PickingListNumber")
    dgr.HeaderLines = 2
    Populate()
  End Sub
  Private Sub Populate()
    Dim vList As New ParameterList(True)
    vList.IntegerValue("PickingListNumber") = mvPickingListNumber
    Dim vDataSet As DataSet = DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstPickingList, vList)
    dgr.Populate(vDataSet)
    dgr.SetCellsEditable()
    Dim vReadOnly As Boolean
    For vIndex As Integer = 0 To dgr.ColumnCount - 1
      If dgr.ColumnName(vIndex) = "Shortfall" Then vReadOnly = False Else vReadOnly = True
      dgr.SetCellsReadOnly(, vIndex, vReadOnly)
    Next
    For vIndex As Integer = 0 To dgr.RowCount - 1
      dgr.SetCellMaxMinValue(vIndex, "Shortfall", IntegerValue(dgr.GetValue(vIndex, "Quantity")), 0, False)
    Next
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vShortfall As Integer
    Dim vOriginalShortfall As Integer
    
    Dim vList As New ParameterList(True)
    For vIndex As Integer = 0 To dgr.RowCount - 1
      vShortfall = IntegerValue(dgr.GetValue(vIndex, "Shortfall"))
      vOriginalShortfall = IntegerValue(dgr.GetValue(vIndex, "OriginalShortfall"))
      If vShortfall <> vOriginalShortfall Then
        vList("Product") = dgr.GetValue(vIndex, "Product")
        vList.IntegerValue("Shortfall") = vShortfall
        DataHelper.UpdatePickingList(mvPickingListNumber, vList)
      End If
    Next
    Me.DialogResult = System.Windows.Forms.DialogResult.OK

  End Sub

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub
End Class
