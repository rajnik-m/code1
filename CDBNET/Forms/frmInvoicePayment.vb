Public Class frmInvoicePayment

  Private mvCashInvoices As CollectionList(Of InvoiceInfo)
  Private mvInvoice As InvoiceInfo

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Public Sub New(ByRef pInvoice As InvoiceInfo, ByRef pCashInvoices As CollectionList(Of InvoiceInfo), ByVal pSalesLedgerAccount As String)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pInvoice, pCashInvoices, pSalesLedgerAccount)
  End Sub

  Private Sub InitialiseControls(ByRef pInvoice As InvoiceInfo, ByRef pCashInvoices As CollectionList(Of InvoiceInfo), ByVal pSalesLedgerAccount As String)
    SetControlTheme()
    Me.Text = ControlText.FrmInvoicePayment
    Dim vPanelItems As New PanelItems("eplTop")
    mvCashInvoices = New CollectionList(Of InvoiceInfo)
    For Each vInvoice As InvoiceInfo In pCashInvoices
      mvCashInvoices.Add(vInvoice.Key, DirectCast(vInvoice.Clone, InvoiceInfo))
    Next
    mvInvoice = DirectCast(pInvoice.Clone, InvoiceInfo)
    'save the originals into module level variables as we need to get back to these if the user cancels it
    'add the controls to the top epl for the invoice status as before any changes occurred/from the database
    vPanelItems.Add(New PanelItem("InvoiceAmount", PanelItem.ControlTypes.ctReadOnly, New Rectangle(5, 8, 100, 24), ControlText.LblInvoiceAmount, 80))
    vPanelItems.Add(New PanelItem("Paid", PanelItem.ControlTypes.ctReadOnly, New Rectangle(215, 8, 100, 24), ControlText.LblAmountPaid, 80))
    vPanelItems.Add(New PanelItem("AmountOutstanding", PanelItem.ControlTypes.ctReadOnly, New Rectangle(400, 8, 100, 24), ControlText.LblOutstanding, 80))

    eplTop.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optNone, vPanelItems))

    'add the controls to the bottom epl for the invoice status as with the current allocations
    vPanelItems = New PanelItems("eplBottom")
    vPanelItems.Add(New PanelItem("TotalUsed", PanelItem.ControlTypes.ctReadOnly, New Rectangle(195, 8, 100, 24), ControlText.LblTotalAmountUsed, 100))
    vPanelItems.Add(New PanelItem("TotalPaid", PanelItem.ControlTypes.ctReadOnly, New Rectangle(195, 33, 100, 24), ControlText.LblTotalAmountPaid, 100))
    vPanelItems.Add(New PanelItem("NewOutstanding", PanelItem.ControlTypes.ctReadOnly, New Rectangle(400, 33, 100, 24), ControlText.LblOutstanding, 80))
    eplBottom.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optNone, vPanelItems))

    Me.Size = New Size(600, 300)

    'add handler for the value change on the grid
    AddHandler dgrCashInvoices.ValueChanged, AddressOf mvCashInvoicesDGR_ValueChanged
    'fill the grid with the current payment and any unallocated cash
    FillGrid(dgrCashInvoices, pCashInvoices, pInvoice, pSalesLedgerAccount)
  End Sub

  Private Sub FillGrid(ByVal pDGR As DisplayGrid, ByVal pCashInvoices As CollectionList(Of InvoiceInfo), ByVal pInvoice As InvoiceInfo, ByVal pSalesLedgerAccount As String)
    Dim vDataSet As New DataSet
    Dim vTable As DataTable = DataHelper.NewColumnTable
    Dim vTotalUsed As Double = 0

    'add columns for the column table
    DataHelper.AddDataColumn(vTable, "PaymentType", "Type")
    DataHelper.AddDataColumn(vTable, "DateReceived", "Date Received", "Date")
    DataHelper.AddDataColumn(vTable, "AmountAvailable", "Amount Available", "Numeric")
    DataHelper.AddDataColumn(vTable, "AmountUsed", "Amount Used", "Numeric")
    DataHelper.AddDataColumn(vTable, "InvoiceNumber", "InvoiceNumber", , "N")
    DataHelper.AddDataColumn(vTable, "BatchNumber", "BatchNumber", , "N")
    DataHelper.AddDataColumn(vTable, "TransactionNumber", "TransactionNumber", , "N")
    vDataSet.Tables.Add(vTable)

    Dim vDataTable As New DataTable("DataRow")

    'define the columns for the datarow table
    vDataTable.Columns.Add("PaymentType", Type.GetType("System.String"))
    vDataTable.Columns.Add("DateReceived", Type.GetType("System.String"))
    vDataTable.Columns.Add("AmountAvailable", Type.GetType("System.Double"))
    vDataTable.Columns.Add("AmountUsed", Type.GetType("System.Double"))
    vDataTable.Columns.Add("InvoiceNumber", Type.GetType("System.String"))
    vDataTable.Columns.Add("BatchNumber", Type.GetType("System.String"))
    vDataTable.Columns.Add("TransactionNumber", Type.GetType("System.String"))

    'add data to the columns from the pCashInvoices passed in from the trader page.
    Dim vPaymentType As String = String.Empty
    For vIndex As Integer = 0 To pCashInvoices.Count - 1
      Dim vDataRow As DataRow = vDataTable.NewRow
      If pCashInvoices(vIndex).InvoiceNumber = 0 Then
        vPaymentType = ControlText.FrmCurrentPayment
      ElseIf pCashInvoices(vIndex).RecordType.Equals("C", StringComparison.InvariantCultureIgnoreCase) Then
        vPaymentType = ControlText.FrmUnallocatedCash
        If (String.IsNullOrEmpty(pSalesLedgerAccount) = False AndAlso pCashInvoices(vIndex).SalesLedgerAccount.Equals(pSalesLedgerAccount, StringComparison.InvariantCultureIgnoreCase) = False) Then vPaymentType = ControlText.FrmPayerUnallocatedCash
      Else
        vPaymentType = ControlText.FrmSundryCreditNote
      End If
      vDataRow.Item("PaymentType") = vPaymentType
      vDataRow.Item("DateReceived") = pCashInvoices(vIndex).InvoiceDate
      vDataRow.Item("AmountAvailable") = FixTwoPlaces(pCashInvoices(vIndex).InvoiceAmount - (pCashInvoices(vIndex).AmountPaid + pCashInvoices(vIndex).AmountUsed))
      vDataRow.Item("AmountUsed") = FixTwoPlaces(pInvoice.GetAmountPaid(pCashInvoices(vIndex).InvoiceNumber))
      vDataRow.Item("InvoiceNumber") = pCashInvoices(vIndex).InvoiceNumber
      vDataRow.Item("BatchNumber") = pCashInvoices(vIndex).BatchNumber
      vDataRow.Item("TransactionNumber") = pCashInvoices(vIndex).TransactionNumber
      vTotalUsed = vTotalUsed + CDbl(vDataRow.Item("AmountUsed"))
      vDataTable.Rows.Add(vDataRow)
    Next
    vDataSet.Tables.Add(vDataTable)
    With pDGR
      .AutoSetHeight = True
      .Populate(vDataSet)
      If .RowCount > 0 Then
        'to set the operation mode to normal
        .SetCellsEditable()
        'set the grid readonly
        .SetCellsReadOnly()
        'make the amountused writable
        .SetCellsReadOnly(, .GetColumn("AmountUsed"), False)
        'select the first amount used column
        .SetActiveCell(0, "AmountUsed")
      End If
    End With
    With pInvoice
      'set the various controls with the appropriate values as passed by the trader page
      eplTop.SetValue("InvoiceAmount", .InvoiceAmount.ToString("N"))
      eplTop.SetValue("Paid", .AmountPaid.ToString("N"))
      eplTop.SetValue("AmountOutstanding", (.InvoiceAmount - .OriginalAmountPaid).ToString("N"))
      eplBottom.SetValue("TotalUsed", vTotalUsed.ToString("N"))
      eplBottom.SetValue("TotalPaid", .AmountPaid.ToString("N"))
      eplBottom.SetValue("NewOutstanding", (CDbl(eplTop.GetValue("AmountOutstanding")) - CDbl(eplBottom.GetValue("TotalUsed"))).ToString("N"))
    End With
  End Sub

  Private Sub mvCashInvoicesDGR_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    Dim vUsed As Double
    Dim vAvailable As Double
    Dim vNowUsed As Double

    With dgrCashInvoices
      'get the key for the invoices collection
      Dim vKey As String = InvoiceInfo.KeyValue(.GetValue(pRow, "BatchNumber"), .GetValue(pRow, "TransactionNumber"))
      If pCol = .GetColumn("AmountUsed") Then
        vUsed = DoubleValue(pValue)
        vNowUsed = vUsed - DoubleValue(pOldValue)
        If vNowUsed <> 0 Then
          'if the user changed the amount used
          vAvailable = CDbl(.GetValue(pRow, "AmountAvailable"))
          If FixTwoPlaces(vUsed) < 0 Then
            'amount used cannot be less than zero
            .SetValue(pRow, pCol, pOldValue)
            ShowInformationMessage(InformationMessages.ImAmoundUsedZero)
          ElseIf FixTwoPlaces(vNowUsed) > FixTwoPlaces(vAvailable) Then
            .SetValue(pRow, pCol, pOldValue)
            ShowInformationMessage(InformationMessages.ImAmountUsedMoreThanAvailable, (CDbl(pOldValue) + vAvailable).ToString)
          ElseIf FixTwoPlaces(vNowUsed) > FixTwoPlaces(CDbl(eplBottom.GetValue("NewOutstanding"))) Then
            .SetValue(pRow, pCol, pOldValue)
            ShowInformationMessage(InformationMessages.ImTotalPaidMoreThanInvAmt)
          Else
            vAvailable = FixTwoPlaces(vAvailable - vNowUsed)
            eplBottom.SetValue("TotalUsed", (CDbl(eplBottom.GetValue("TotalUsed")) + vNowUsed).ToString("N"))
            eplBottom.SetValue("TotalPaid", (CDbl(eplBottom.GetValue("TotalUsed")) + DoubleValue(eplTop.GetValue("Paid"))).ToString("N"))
            eplBottom.SetValue("NewOutstanding", (CDbl(eplBottom.GetValue("NewOutstanding")) - vNowUsed).ToString("N"))
            .SetValue(pRow, "AmountAvailable", vAvailable.ToString)
            'add payment to the current invoice
            mvInvoice.AddPayment(mvCashInvoices(vKey), vNowUsed)
            'update the cash invoice's amount used
            mvCashInvoices(vKey).AmountUsed = mvCashInvoices(vKey).AmountUsed + vNowUsed
          End If
        End If
      End If
    End With
  End Sub

  Public ReadOnly Property CashInvoices() As CollectionList(Of InvoiceInfo)
    Get
      Return mvCashInvoices
    End Get
  End Property

  Public ReadOnly Property Invoice() As InvoiceInfo
    Get
      Return mvInvoice
    End Get
  End Property
End Class