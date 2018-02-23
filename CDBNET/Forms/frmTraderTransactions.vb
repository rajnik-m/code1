Public Class frmTraderTransactions
  Inherits ThemedForm

#Region " Windows Form Designer generated code "

  Public Sub New(ByVal pTraderApplication As TraderApplication, ByVal pBatchInfo As BatchInfo)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pTraderApplication, pBatchInfo)
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdEdit As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdSetTotals As System.Windows.Forms.Button
  Friend WithEvents cmdPrint As System.Windows.Forms.Button
  Friend WithEvents cmdConfirm As System.Windows.Forms.Button
  Friend WithEvents cmdNewBatch As System.Windows.Forms.Button
  Friend WithEvents spl As System.Windows.Forms.SplitContainer
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents tbp4 As System.Windows.Forms.TabPage
  Friend WithEvents dpl As CDBNETCL.DisplayPanel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTraderTransactions))
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdNewBatch = New System.Windows.Forms.Button()
    Me.cmdPrint = New System.Windows.Forms.Button()
    Me.cmdConfirm = New System.Windows.Forms.Button()
    Me.cmdSetTotals = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdAdd = New System.Windows.Forms.Button()
    Me.cmdEdit = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.spl = New System.Windows.Forms.SplitContainer()
    Me.epl = New CDBNETCL.EditPanel()
    Me.dpl = New CDBNETCL.DisplayPanel()
    Me.tbp4 = New System.Windows.Forms.TabPage()
    Me.bpl.SuspendLayout()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spl.Panel1.SuspendLayout()
    Me.spl.Panel2.SuspendLayout()
    Me.spl.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 62)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = True
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(595, 131)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdNewBatch)
    Me.bpl.Controls.Add(Me.cmdPrint)
    Me.bpl.Controls.Add(Me.cmdConfirm)
    Me.bpl.Controls.Add(Me.cmdSetTotals)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdAdd)
    Me.bpl.Controls.Add(Me.cmdEdit)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 287)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(595, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdNewBatch
    '
    Me.cmdNewBatch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNewBatch.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdNewBatch.Location = New System.Drawing.Point(2, 6)
    Me.cmdNewBatch.Name = "cmdNewBatch"
    Me.cmdNewBatch.Size = New System.Drawing.Size(73, 27)
    Me.cmdNewBatch.TabIndex = 12
    Me.cmdNewBatch.Text = "&New Batch"
    Me.cmdNewBatch.Visible = False
    '
    'cmdPrint
    '
    Me.cmdPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrint.Location = New System.Drawing.Point(76, 6)
    Me.cmdPrint.Name = "cmdPrint"
    Me.cmdPrint.Size = New System.Drawing.Size(73, 27)
    Me.cmdPrint.TabIndex = 11
    Me.cmdPrint.Text = "&Print"
    Me.cmdPrint.Visible = False
    '
    'cmdConfirm
    '
    Me.cmdConfirm.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdConfirm.Location = New System.Drawing.Point(150, 6)
    Me.cmdConfirm.Name = "cmdConfirm"
    Me.cmdConfirm.Size = New System.Drawing.Size(73, 27)
    Me.cmdConfirm.TabIndex = 10
    Me.cmdConfirm.Text = "Confirm"
    Me.cmdConfirm.Visible = False
    '
    'cmdSetTotals
    '
    Me.cmdSetTotals.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSetTotals.Location = New System.Drawing.Point(224, 6)
    Me.cmdSetTotals.Name = "cmdSetTotals"
    Me.cmdSetTotals.Size = New System.Drawing.Size(73, 27)
    Me.cmdSetTotals.TabIndex = 9
    Me.cmdSetTotals.Text = "&Set Totals"
    Me.cmdSetTotals.Visible = False
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.Location = New System.Drawing.Point(298, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(73, 27)
    Me.cmdDelete.TabIndex = 8
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.Visible = False
    '
    'cmdAdd
    '
    Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAdd.Location = New System.Drawing.Point(372, 6)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(73, 27)
    Me.cmdAdd.TabIndex = 5
    Me.cmdAdd.Text = "&Add"
    Me.cmdAdd.Visible = False
    '
    'cmdEdit
    '
    Me.cmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdEdit.Location = New System.Drawing.Point(446, 6)
    Me.cmdEdit.Name = "cmdEdit"
    Me.cmdEdit.Size = New System.Drawing.Size(73, 27)
    Me.cmdEdit.TabIndex = 7
    Me.cmdEdit.Text = "&Edit"
    '
    'cmdClose
    '
    Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(520, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(73, 27)
    Me.cmdClose.TabIndex = 6
    Me.cmdClose.Text = "Close"
    '
    'spl
    '
    Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spl.Location = New System.Drawing.Point(0, 0)
    Me.spl.Name = "spl"
    Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'spl.Panel1
    '
    Me.spl.Panel1.Controls.Add(Me.dgr)
    Me.spl.Panel1.Controls.Add(Me.epl)
    '
    'spl.Panel2
    '
    Me.spl.Panel2.Controls.Add(Me.dpl)
    Me.spl.Size = New System.Drawing.Size(595, 287)
    Me.spl.SplitterDistance = 193
    Me.spl.TabIndex = 2
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Top
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(595, 62)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 1
    Me.epl.TabSelectedIndex = 0
    '
    'dpl
    '
    Me.dpl.AutoSetHeight = False
    Me.dpl.BackColor = System.Drawing.Color.Transparent
    Me.dpl.ColumnSizingType = CDBNETCL.DisplayPanel.ColumnSizingTypes.Automatic
    Me.dpl.DataSelectionType = CDBNETCL.CareNetServices.XMLContactDataSelectionTypes.xcdtNone
    Me.dpl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dpl.Location = New System.Drawing.Point(0, 0)
    Me.dpl.Name = "dpl"
    Me.dpl.ProcessResize = False
    Me.dpl.ShowAllText = False
    Me.dpl.Size = New System.Drawing.Size(595, 90)
    Me.dpl.TabIndex = 0
    '
    'tbp4
    '
    Me.tbp4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbp4.Location = New System.Drawing.Point(4, 22)
    Me.tbp4.Name = "tbp4"
    Me.tbp4.Size = New System.Drawing.Size(580, 70)
    Me.tbp4.TabIndex = 2
    Me.tbp4.Visible = False
    '
    'frmTraderTransactions
    '
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(595, 326)
    Me.Controls.Add(Me.spl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTraderTransactions"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.spl.Panel1.ResumeLayout(False)
    Me.spl.Panel2.ResumeLayout(False)
    CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private mvTraderApplicationNumber As Integer
  Private mvTraderApplication As TraderApplication
  Private mvBatchInfo As BatchInfo
  Private mvDataSet As DataSet

  Private Sub InitialiseControls(ByVal pTraderApplication As TraderApplication, ByVal pBatchInfo As BatchInfo)
    SetControlTheme()
    Me.cmdNewBatch.Text = ControlText.CmdNewBatch
    Me.cmdPrint.Text = ControlText.CmdPrint
    Me.cmdConfirm.Text = ControlText.CmdNConfirm
    Me.cmdSetTotals.Text = ControlText.CmdSetTotals
    Me.cmdDelete.Text = ControlText.CmdnDelete
    Me.cmdAdd.Text = ControlText.CmdAdd
    Me.cmdEdit.Text = ControlText.CmdEdit
    Me.cmdClose.Text = ControlText.CmdClose

    Me.Text = ControlText.FrmTransactions
    mvTraderApplication = pTraderApplication
    mvTraderApplicationNumber = pTraderApplication.ApplicationNumber
    mvBatchInfo = pBatchInfo
    MainHelper.SetMDIParent(Me)
    With epl
      .Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optTraderTransactions))
      .SetValue("BatchNumber", mvBatchInfo.BatchNumber.ToString)
      .SetValue("NumberOfEntries", mvBatchInfo.NumberOfEntries.ToString())
      .SetValue("BatchTotal", mvBatchInfo.BatchTotal.ToString("0.00"))
      .Height = epl.RequiredHeight
    End With
    GetTransactionData(mvBatchInfo)
    'Column headings are set by DisplayListMaintenance
    dgr.AutoSetHeight = True
  End Sub

  Public Sub GetTransactionData(ByVal pBatchInfo As BatchInfo)
    mvDataSet = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, mvBatchInfo.BatchNumber)
    If mvDataSet.Tables.Contains("DataRow") = False Then
      'If there is no data, do this again so that we get the table with an empty row otherwise data not displayed correctly whilst in Trader
      Dim vList As New ParameterList(True, True)
      vList("DocumentColumns") = "Y"
      mvDataSet = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, mvBatchInfo.BatchNumber, 0, vList)
    End If
    dgr.Populate(mvDataSet)
    dpl.Init(mvDataSet, False, False)
    dpl.Populate(mvDataSet, 0)
    mvBatchInfo = pBatchInfo
    ResetControls()
  End Sub

  Private Sub ResetControls()
    Dim vZeroAmounts As Boolean
    With mvBatchInfo
      vZeroAmounts = (.BatchTotal = 0 And .NumberOfEntries = 0)
      cmdSetTotals.Visible = vZeroAmounts
      'cmdConfirm.Visible = (.Provisional = True And .PostedToNominal = True)   NOT Supported Yet
      cmdAdd.Visible = True
      cmdDelete.Visible = True
      cmdNewBatch.Visible = True
      cmdPrint.Visible = Not vZeroAmounts
      SetControlsEnabled()
      epl.SetValue("NumberOfTransactions", .NumberOfTransactions.ToString())
      epl.SetValue("TransactionTotal", .TransactionTotal.ToString("0.00"))
    End With
  End Sub

  Private Sub SetControlsEnabled()
    If mvBatchInfo.BatchTypeCode = "NF" Then
      'Non-Financial Batch
      cmdEdit.Enabled = False     'Can not edit an existing transaction
      cmdPrint.Visible = False
    Else
      cmdEdit.Enabled = (dgr.DataRowCount > 0)
      If dgr.DataRowCount > 0 Then
        If dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("Allocated")) = "Y" OrElse
          dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("Adjustment")) = "Y" OrElse
          (dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("RecordType")) = "N" AndAlso
          BooleanValue(dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("Printed")))) Then
          cmdEdit.Enabled = False
        End If
      End If
    End If
    cmdDelete.Enabled = cmdEdit.Enabled
    If cmdConfirm.Visible Then
      cmdConfirm.Enabled = cmdEdit.Enabled
    End If
  End Sub

  Private Sub frmTraderTransactions_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    MainHelper.EnableTraderApplications(True)
  End Sub

  Private Sub frmSelectListItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Me.Size = New Size(680, 422)
    Me.Location = MDILocation(Me.Width, Me.Height)
    bpl.RepositionButtons()
  End Sub

  Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
    Dim vCursor As New BusyCursor
    Try
      'Edit existing transaction
      If mvBatchInfo.BatchTypeCode = "CS" Then
        'This will need amending once CreditNotes are supported
        If dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("RecordType")) = "N" AndAlso dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("Printed")) = "Y" Then
          Throw New CareException(CareException.ErrorNumbers.enTransactionIsCreditNote)
        End If
      End If
      RunTrader(IntegerValue(dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("TransactionNumber"))))
    Catch vCareException As CareException
      If vCareException.ErrorNumber = CareException.ErrorNumbers.enTransactionIsCreditNote Then
        ShowInformationMessage(vCareException.Message)
      Else
        DataHelper.HandleException(vCareException)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    Dim vCursor As New BusyCursor()
    Try
      Dim vMaxVouchers As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.cv_max_number_of_vouchers)
      With mvBatchInfo
        If .PostedToCashBook = True And .BatchTypeCode <> "CS" And .BatchTypeCode <> "SP" And AppValues.ConfigurationOption(AppValues.ConfigurationOptions.batch_bypass_cheque_list) = False Then
          Throw New CareException(CareException.ErrorNumbers.enBatchPostedToCashBook)
        ElseIf .PostedToNominal = True Then
          Throw New CareException(CareException.ErrorNumbers.enBatchPostedToNominal)
        ElseIf .Picked <> "N" Then
          Throw New CareException(CareException.ErrorNumbers.enBatchStockPicked)
        ElseIf (.BatchTypeCode = "CV" And vMaxVouchers.Length > 0) AndAlso (.NumberOfTransactions = CInt(vMaxVouchers)) Then
          Throw New CareException(CareException.ErrorNumbers.enBatchMaxTransactions)
        Else
          RunTrader(0)
        End If
      End With

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enCannotLockBatchInUseBy, CareException.ErrorNumbers.enCannotLockBatch
          ShowErrorMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enBatchPostedToCashBook, CareException.ErrorNumbers.enBatchPostedToNominal, CareException.ErrorNumbers.enBatchMaxTransactions
          ShowWarningMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enBatchStockPicked
          ShowWarningMessage(vCareException.Message, IIf(mvBatchInfo.Picked = "Y", "picked", "picked and confirmed").ToString)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub RunTrader(ByVal pTransactionNumber As Integer)
    'Lock the batch then set the trader application parameters and run trader
    Dim vList As New ParameterList(True)
    vList.IntegerValue("BatchNumber") = mvBatchInfo.BatchNumber
    DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoLockBatch, vList)
    If mvTraderApplication Is Nothing Then mvTraderApplication = New TraderApplication(mvTraderApplicationNumber)
    mvTraderApplication.BatchNumber = mvBatchInfo.BatchNumber
    mvTraderApplication.TransactionNumber = pTransactionNumber
    mvTraderApplication.BatchInfo = mvBatchInfo
    mvTraderApplication.BatchLocked = True
    FormHelper.RunTraderApplication(mvTraderApplication, Nothing, Me)
    mvTraderApplication = Nothing
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Dim vCursor As New BusyCursor
    Try
      If ConfirmDelete() Then
        Dim vRow As Integer = dgr.CurrentRow
        If vRow >= 0 Then
          Dim vTransNo As Integer = IntegerValue(dgr.GetValue(vRow, dgr.GetColumn("TransactionNumber")))
          If vTransNo > 0 Then
            Dim vList As New ParameterList(True)
            vList.IntegerValue("BatchNumber") = mvBatchInfo.BatchNumber
            vList.IntegerValue("TransactionNumber") = vTransNo
            Dim vDS As DataSet = DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoDeleteBatchTransaction, vList)
            mvBatchInfo = New BatchInfo(DataHelper.GetRowFromDataSet(vDS))
            'The table still contains the deleted row so re-populate the grid
            GetTransactionData(mvBatchInfo)
            'select another row
            If vRow > dgr.DataRowCount - 1 Then
              If vRow - 1 >= 0 Then
                dgr.SelectRow(vRow - 1)
              End If
            Else
              dgr.SelectRow(vRow)
            End If
            ResetControls()
          End If
        End If
      End If
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enTransactionDeletionError, CareException.ErrorNumbers.enAccessLevel
          ShowWarningMessage(InformationMessages.ImDeleteTransFailed, vCareException.Message)
        Case CareException.ErrorNumbers.enCannotDeleteSalesLedgerTransaction
          ShowInformationMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdNewBatch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNewBatch.Click
    Dim vCursor As New BusyCursor()

    Try
      If mvTraderApplication Is Nothing Then mvTraderApplication = New TraderApplication(mvTraderApplicationNumber)
      Dim vBatchNumber As Integer = FormHelper.CreateNewTraderBatch(mvTraderApplication, Me)
      If vBatchNumber > 0 Then
        mvBatchInfo = New BatchInfo(vBatchNumber) 'OK initialise the batch
        epl.SetValue("BatchNumber", mvBatchInfo.BatchNumber.ToString)
        epl.SetValue("NumberOfEntries", mvBatchInfo.NumberOfEntries.ToString())
        epl.SetValue("BatchTotal", mvBatchInfo.BatchTotal.ToString("0.00"))
        GetTransactionData(mvBatchInfo)
        RunTrader(0)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdSetTotals_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSetTotals.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      vList.IntegerValue("BatchNumber") = mvBatchInfo.BatchNumber
      Dim vDS As DataSet = DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoSetTotals, vList)
      mvBatchInfo = New BatchInfo(DataHelper.GetRowFromDataSet(vDS))
      epl.SetValue("NumberOfEntries", mvBatchInfo.NumberOfEntries.ToString())
      epl.SetValue("BatchTotal", mvBatchInfo.BatchTotal.ToString("0.00"))
      epl.SetValue("NumberOfTransactions", mvBatchInfo.NumberOfTransactions.ToString())
      epl.SetValue("TransactionTotal", mvBatchInfo.TransactionTotal.ToString("0.00"))
      cmdSetTotals.Visible = False
      cmdPrint.Visible = True
      cmdEdit.Focus()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      vList("ReportCode") = "FPBSUM"
      vList.IntegerValue("RPbatch_number") = mvBatchInfo.BatchNumber
      Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub dgr_ContactSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    dpl.Populate(mvDataSet, pRow)
    SetControlsEnabled()
  End Sub

  ''' <summary>Display Trader form if a new Batch was created.</summary>
  Friend Sub RunTraderForNewBatch()
    If mvBatchInfo.IsNewBatch Then RunTrader(0)
  End Sub

  Friend Sub NewTransactionAdded(ByVal pContactInfo As ContactInfo, ByVal pTransactionAmount As Double, ByVal pCurrencyAmount As Double, ByVal pTransactionDate As String, ByVal pPayerAddressLine As String)
    Dim vList As New ParameterList()
    Dim vNumberOfTrans As Integer = (IntegerValue(epl.GetValue("NumberOfTransactions")) + 1)
    Dim vTransAmount As Double = pTransactionAmount
    Dim vTransCurrencyAmount As Double = pCurrencyAmount

    vList("Amount") = vTransAmount.ToString("0.00")
    vList("CurrencyAmount") = vTransCurrencyAmount.ToString("0.00")
    vList("TransactionDate") = pTransactionDate
    vList("ContactName") = pContactInfo.ContactName
    vList("Surname") = pContactInfo.Surname
    vList.IntegerValue("ContactNumber") = pContactInfo.ContactNumber
    vList.IntegerValue("AddressNumber") = pContactInfo.AddressNumber
    vList("AddressLine") = pPayerAddressLine
    vList.IntegerValue("BatchNumber") = mvBatchInfo.BatchNumber
    vList.IntegerValue("TransactionNumber") = mvBatchInfo.NextTransactionNumber
    mvBatchInfo.NextTransactionNumber += 1

    vTransAmount = FixTwoPlaces(vTransAmount + DoubleValue(epl.GetValue("TransactionTotal")))
    epl.SetValue("NumberOfTransactions", vNumberOfTrans.ToString())
    epl.SetValue("TransactionTotal", vTransAmount.ToString("0.00"))
    dgr.AddDataRow(vList, 0)

  End Sub
End Class
