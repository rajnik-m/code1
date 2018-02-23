Public Class frmBatchProcessing
  Inherits MaintenanceParentForm

  Private mvBatchState As BatchStates = BatchStates.Outstanding
  Private mvPreviousTab As TabPage
  Private mvRowCount As Long
  Private mvDataSet As DataSet = Nothing
  Private mvDataTable As DataTable = Nothing
  Private mvOptionPostToCashBook As Boolean
  Private mvManualPayingInSlips As Boolean
  Private mvManualPISLocation As String
  Private mvPostSingleBatchToCB As Boolean
  Private mvBatchNumber As Integer
  Private mvMainForm As frmMain
  Private mvSelectedRow As Integer

  Private Enum BatchStates
    Outstanding
    Incomplete
    DetailComplete
    PickingList
    ConfirmStock
    ChequeList
    CreateClaim
    PrintPayingInSlips
    PostToCashBook
    PostBatch
  End Enum

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()
    SettingsName = "BatchProcessing"
    MainHelper.SetMDIParent(Me)
    mvManualPayingInSlips = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_manual_paying_in_slips)
    mvManualPISLocation = AppValues.ConfigurationValue(AppValues.ConfigurationValues.manual_pis_location, "CLOSEBATCH")
    mvOptionPostToCashBook = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_post_batches_to_CB)
    mvPostSingleBatchToCB = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_cbp_post_single_batch)
    cmdProcess.Enabled = False
    For vIndex As Integer = 0 To tab.TabPages.Count - 1
      tab.TabPages(vIndex).Tag = vIndex
    Next
    If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_stock_processing) Then
      tab.TabPages.Remove(tbpPickingList)
      tab.TabPages.Remove(tbpConfirmStock)
    End If
    dgr.MultipleSelect = False
    dgr.HeaderLines = 2
    tab.SelectedIndex = BatchStates.Outstanding
    Populate()
    If DataHelper.UserInfo.AccessLevel <> UserInfo.UserAccessLevel.ualDatabaseAdministrator AndAlso DataHelper.UserInfo.AccessLevel <> UserInfo.UserAccessLevel.ualSupervisor Then
      cmdDelete.Enabled = False 'BR20545 moved from top of method to the bottom so that it is not overidden.
    End If
  End Sub

  Public Overrides Sub RefreshData()
    mvBatchNumber = 0
    mvSelectedRow = 0
    Populate()
  End Sub

  Private Sub Populate()
    Dim vList As New ParameterList(True)
    mvDataSet = DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstBatchProcessing)
    mvDataTable = DataHelper.GetTableFromDataSet(mvDataSet)
    If Not mvDataTable Is Nothing Then mvDataTable.Columns.Add("Hide")
    ShowBatches(mvBatchState)
  End Sub

  Private Sub ShowBatches(ByVal pBatchState As BatchStates)
    Try
      Dim vHide As Boolean
      Dim vCount As Long
      Dim vBalanced As Boolean
      Dim vBatch As BatchInfo

      mvSelectedRow = 0
      If Not mvDataTable Is Nothing Then
        For Each vRow As DataRow In mvDataTable.Rows
          '"BatchNo,BatchType,BankAccount,Date,NumberOfEntries,BatchTotal,NumberEntered,TransactionTotal,DetailComplete,ReadyforBanking,
          'PayingInSlipPrinted,PayingInSlipNumber,Picked,CashBook,Posted,JobNumber,PaymentMethod,Category,Provisional,ClaimSent,PrintChequeList,Hide"
          vBatch = New BatchInfo(vRow)
          'Debug.Print(vRow("BatchNumber").ToString)
          With vBatch

            vBalanced = (.BatchTotal = .TransactionTotal) And (.NumberOfEntries = .NumberOfTransactions) And (.NumberOfTransactions > 0) And (.NumberOfEntries > 0)
            vHide = True
            Select Case pBatchState
              Case BatchStates.Outstanding
                vHide = False
              Case BatchStates.Incomplete
                If (.DetailCompleted = False AndAlso vBalanced = False) OrElse (.BatchType = CareServices.BatchTypes.StandingOrder AndAlso .ReadyForBanking = False) Then vHide = False
              Case BatchStates.DetailComplete
                If .DetailCompleted = False AndAlso vBalanced Then vHide = False
              Case BatchStates.ChequeList
                If vBalanced AndAlso .ReadyForBanking = False AndAlso .PrintChequeList = True Then vHide = False
              Case BatchStates.CreateClaim
                If vBalanced Then
                  If .Provisional Then
                    If .DetailCompleted AndAlso .BatchType <> CareServices.BatchTypes.Cash AndAlso .BatchType <> CareServices.BatchTypes.CashWithInvoice AndAlso vRow("ClaimSent").ToString.Length = 0 Then vHide = False
                  Else
                    Select Case .BatchType
                      Case CareServices.BatchTypes.CreditCard, CareServices.BatchTypes.DebitCard, CareNetServices.BatchTypes.CreditCardWithInvoice
                        If .DetailCompleted AndAlso .PayingInSlipPrinted = False AndAlso .PostedToCashBook = False Then vHide = False
                      Case CareServices.BatchTypes.CreditCardAuthority
                        If .DetailCompleted AndAlso .ReadyForBanking = False AndAlso .PayingInSlipPrinted = False Then vHide = False
                      Case CareServices.BatchTypes.DirectDebit, CareServices.BatchTypes.DirectCredit
                        If .ReadyForBanking = False AndAlso .PayingInSlipPrinted = False Then vHide = False
                    End Select
                  End If
                End If
              Case BatchStates.PrintPayingInSlips
                If .ReadyForBanking Or (vBalanced = True And .PrintChequeList = False) Then
                  If .PayingInSlipPrinted = False AndAlso .PostedToCashBook = False AndAlso (.BatchType = CareServices.BatchTypes.Cash OrElse .BatchType = CareServices.BatchTypes.CashWithInvoice) Then vHide = False
                End If
              Case BatchStates.PickingList
                If .DetailCompleted AndAlso (.Picked = "N" OrElse .Picked = "P") Then vHide = False

              Case BatchStates.ConfirmStock
                If .Picked = "Y" OrElse .Picked = "P" Then vHide = False

              Case BatchStates.PostToCashBook
                If vBalanced Or (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.batch_bypass_cheque_list) And .NumberOfTransactions = 0) Then

                  If .ReadyForBanking OrElse .PrintChequeList = False Then

                    If .PayingInSlipPrinted Then
                      If .PostedToCashBook = False Then
                        Select Case .BatchType
                          Case CareServices.BatchTypes.Cash, CareServices.BatchTypes.CashWithInvoice, CareServices.BatchTypes.DirectDebit, CareServices.BatchTypes.FinancialAdjustment, _
                                CareServices.BatchTypes.CreditCard, CareNetServices.BatchTypes.CreditCardWithInvoice, CareServices.BatchTypes.DebitCard, CareServices.BatchTypes.CreditCardAuthority, _
                                CareServices.BatchTypes.DirectCredit, CareServices.BatchTypes.GiveAsYouEarn, CareServices.BatchTypes.CAFCards, _
                                CareServices.BatchTypes.PostTaxPayrollGiving
                            vHide = False
                          Case CareServices.BatchTypes.StandingOrder, CareServices.BatchTypes.CAFVouchers, CareServices.BatchTypes.BankStatement, _
                               CareServices.BatchTypes.CAFCommitmentReconciliation
                            If mvOptionPostToCashBook Then vHide = False
                        End Select
                      End If
                    End If
                  End If
                End If
              Case BatchStates.PostBatch
                'BR 7507: RFB must be set for Direct Refund Batch otherwise poss to Post prior to Claim
                If .ReadyForBanking OrElse .BatchType <> CareServices.BatchTypes.DirectCredit Then
                  If .PostedToCashBook AndAlso .DetailCompleted Then
                    If .Picked = "N" OrElse .Picked = "C" Then
                      If .Provisional = False Then vHide = False
                    End If
                  End If
                End If
            End Select
            If vHide Then vRow("Hide") = "Y" Else vRow("Hide") = "N"
            If Not vHide Then vCount = vCount + 1
          End With
        Next
      End If
      mvRowCount = vCount
      cmdDelete.Enabled = mvRowCount > 0
      If DataHelper.UserInfo.AccessLevel <> UserInfo.UserAccessLevel.ualDatabaseAdministrator AndAlso DataHelper.UserInfo.AccessLevel <> UserInfo.UserAccessLevel.ualSupervisor Then
        cmdDelete.Enabled = False
      End If
      cmdDetails.Enabled = mvRowCount > 0
      Dim vProcess As Boolean = True
      Select Case pBatchState
        Case BatchStates.Outstanding
          vProcess = False
        Case BatchStates.Incomplete
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.batch_bypass_cheque_list) = False Then vProcess = False
      End Select
      cmdProcess.Enabled = mvRowCount > 0 AndAlso vProcess
      dgr.Populate(mvDataSet, "Hide='N'")
      SetProcessHeading(pBatchState)
      If mvSelectedRow >= 0 AndAlso dgr.RowCount > 0 Then dgr.SelectRow(mvSelectedRow)
    Catch vEx As Exception
      DataHelper.HandleJobException(vEx)
    End Try
  End Sub

  Private Sub SetProcessHeading(ByVal pBatchState As BatchStates)
    Dim vMsg As String = ""
    Select Case pBatchState
      Case BatchStates.Outstanding
        vMsg = InformationMessages.ImBatchStateOutstanding
      Case BatchStates.Incomplete
        vMsg = InformationMessages.ImBatchStateIncomplete
      Case BatchStates.DetailComplete
        vMsg = InformationMessages.ImBatchStateDetailComplete
      Case BatchStates.ChequeList
        vMsg = InformationMessages.ImBatchStateChequeList
      Case BatchStates.CreateClaim
        vMsg = InformationMessages.ImBatchStateCreateClaim
      Case BatchStates.PrintPayingInSlips
        vMsg = InformationMessages.ImBatchStatePrintPayingInSlips
      Case BatchStates.PickingList
        vMsg = InformationMessages.ImBatchStatePickingList
      Case BatchStates.ConfirmStock
        vMsg = InformationMessages.ImBatchStateConfirmStock
      Case BatchStates.PostToCashBook
        vMsg = InformationMessages.ImBatchStatePostToCashBook
      Case BatchStates.PostBatch
        vMsg = InformationMessages.ImBatchStatePostBatch
    End Select
    tssl.Text = mvRowCount & " " & vMsg
  End Sub

  Private Sub tab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
    If tab.SelectedIndex >= 0 Then
      mvBatchState = CType(tab.TabPages(tab.SelectedIndex).Tag, BatchStates)
      Dim vProcess As Boolean = True
      Select Case mvBatchState
        Case BatchStates.Outstanding
          vProcess = False
          dgr.MultipleSelect = False
        Case BatchStates.Incomplete
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.batch_bypass_cheque_list) = False Then vProcess = False
          dgr.MultipleSelect = False
        Case BatchStates.PickingList, BatchStates.ConfirmStock
          dgr.MultipleSelect = False
        Case Else
          dgr.MultipleSelect = True
      End Select
      ShowBatches(mvBatchState)
      cmdProcess.Enabled = mvRowCount > 0 AndAlso vProcess
    End If
  End Sub

  Private Sub cmdDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDetails.Click
    Try
      Dim vRows() As DataRow = mvDataTable.Select("BatchNumber = '" & mvBatchNumber & "'")
      If vRows.Length = 1 AndAlso mvBatchNumber > 0 Then ProcessBatch(BatchStates.Outstanding, vRows(0))
    Catch vEx As Exception
      DataHelper.HandleJobException(vEx)
    End Try
  End Sub
  Private Sub ProcessBatch(ByVal pBatchState As BatchStates, ByVal pRow As DataRow)
    ProcessBatch(pBatchState, pRow, 0)
  End Sub
  Private Sub ProcessBatch(ByVal pBatchState As BatchStates, ByVal pBatchNumber As Integer, Optional ByVal pLastBatchNumber As Integer = 0)
    ProcessBatch(pBatchState, Nothing, pBatchNumber, pLastBatchNumber)
  End Sub
  Private Sub ProcessBatch(ByVal pBatchState As BatchStates, ByVal pRow As DataRow, ByVal pBatchNumber As Integer, Optional ByVal pLastBatchNumber As Integer = 0)
    With dgr
      Select Case pBatchState
        Case BatchStates.Outstanding          'Batch Details
          ViewBatchDetails(pRow)

        Case BatchStates.Incomplete
          'Should only get here when batch_bypass_check_list config is set
          Dim vList As New ParameterList(True)
          vList("BatchNumber") = pBatchNumber.ToString
          vList("ReadyForBanking") = "Y"
          DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoBatchDetails, vList)
          RefreshData()

        Case BatchStates.ChequeList
          If PrintChequeList(pBatchNumber) Then RefreshData()

        Case BatchStates.CreateClaim
          Dim vCurrentRow As Integer = pBatchNumber
          Dim vBatchNumber As Integer = IntegerValue(dgr.GetValue(vCurrentRow, "BatchNumber"))
          Dim vLastBatchNo As Integer = 0
          Dim vBatchType As String = dgr.GetValue(vCurrentRow, "BatchType")
          ''Handle contiguous batches
          If vBatchType = "DD" Or vBatchType = "DR" Then
            vLastBatchNo = IntegerValue(dgr.GetValue(pLastBatchNumber, "BatchNumber"))
            Dim vBatchRange As Boolean = True
            Dim vRow As Integer = vCurrentRow
            If vBatchNumber = vLastBatchNo Then
              While vBatchRange
                vBatchRange = False
                vRow += 1
                If dgr.RowCount - 1 >= vRow AndAlso dgr.GetValue(vRow, "BatchType") = vBatchType Then
                  If IntegerValue(dgr.GetValue(vRow, "BatchNumber")) = vLastBatchNo - 1 Then
                    vLastBatchNo = vLastBatchNo - 1
                    vBatchRange = True
                  ElseIf IntegerValue(dgr.GetValue(vRow, "BatchNumber")) = vLastBatchNo + 1 Then
                    vLastBatchNo = vLastBatchNo + 1
                    vBatchRange = True
                  End If
                End If
              End While
            End If
          End If
          If vBatchNumber > vLastBatchNo Then
            ProcessClaimFile(vBatchNumber, vLastBatchNo, vBatchType, dgr.GetValue(vCurrentRow, "Provisional") = "Y")
          Else
            ProcessClaimFile(vLastBatchNo, vBatchNumber, vBatchType, dgr.GetValue(vCurrentRow, "Provisional") = "Y")
          End If
        Case Else
          ShowInformationMessage(InformationMessages.ImNotImplemented)
      End Select
    End With
  End Sub

  Private Sub cmdProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdProcess.Click
    Try
      'Process the selected batch (or batches?)
      Select Case mvBatchState
        Case BatchStates.PrintPayingInSlips
          If ShowQuestion(QuestionMessages.QmProducePayinInSlipsForAllBatches, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then ProcessPIS(False)
        Case BatchStates.PostToCashBook
          Dim vParams As New ParameterList(True)
          If mvPostSingleBatchToCB Then
            'Only the highlighted batch(es) will be posted to the cash book
            If dgr.GetSelectedRowNumbers.Count > 0 Then
              Dim vBatchNumbers As String = SelectedBatchNumbers()
              If vBatchNumbers.Length > 0 Then
                If CheckCashBookBatchNumbers() Then
                  vParams("BatchNumbers") = vBatchNumbers
                  FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCashBookPosting, vParams, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
                End If
              End If
            End If
          Else
            If ShowQuestion(QuestionMessages.QmCompleteAllBatches, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then
              If CheckCashBookBatchNumbers() Then FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCashBookPosting, vParams, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
            End If
          End If
        Case BatchStates.PickingList
          Dim vParams As New ParameterList(True)
          Dim vBankAccount As String = dgr.GetValue(mvSelectedRow, "BankAccount")
          If vBankAccount.Length > 0 Then
            Dim vCompany As String = dgr.GetValue(mvSelectedRow, "Company")
            If vCompany.Length > 0 Then
              vParams("Company") = vCompany
            End If
          End If
          If ShowQuestion(QuestionMessages.QmProducePickingLists, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPickingList, vParams, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case BatchStates.ConfirmStock
          Dim vParams As New ParameterList(True)
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_stock_multiple_warehouses) = False Then
            'Only do this if not using multiple warehouses
            If mvBatchNumber > 0 Then
              Dim vPickingListNumber As String = dgr.GetValue(mvSelectedRow, "PickingListNumber")
              If vPickingListNumber.Length > 0 Then vParams("PickingListNumber") = vPickingListNumber
            End If
          End If
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtConfirmStockAllocation, vParams, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case BatchStates.PostBatch

          If mvSelectedRow >= 0 Then
            Dim vParams As New ParameterList(True)
            vParams("JobName") = FormHelper.GetTaskJobTypeName(CareServices.TaskJobTypes.tjtBatchUpdate)
            Dim vResult As DialogResult = FormHelper.ScheduleTask(vParams)
            If vResult <> System.Windows.Forms.DialogResult.Cancel Then
              Dim vBatchNumbers As String = SelectedBatchNumbers()
              If vBatchNumbers.Length > 0 Then ProcessBatchUpdate(vBatchNumbers, vParams, vResult)
            End If
          End If

        Case BatchStates.DetailComplete
          If dgr.GetSelectedRowNumbers.Count > 0 Then
            Dim vBatchNumber, vPayingInSlipNo, vNewPayingInSlipNo As Integer
            Dim vBatchType As String
            Dim vParams As New ParameterList(True)
            Dim vList As ArrayListEx = dgr.GetSelectedRowNumbers
            Dim vPISList, vDetailCompList As New StringBuilder

            For Each vIndex As Integer In vList
              vBatchNumber = IntegerValue(dgr.GetValue(vIndex, "BatchNumber"))
              vBatchType = dgr.GetValue(vIndex, "BatchType")
              vPayingInSlipNo = IntegerValue(dgr.GetValue(vIndex, "PayingInSlipNumber"))
              If vPayingInSlipNo = 0 And mvManualPayingInSlips And (vBatchType = "CA" OrElse vBatchType = "AI") Then
                If vNewPayingInSlipNo = 0 Then
                  vParams = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptPayingInSlipNumber, vParams)
                End If
                If vParams.Count > 0 AndAlso vBatchNumber > 0 Then
                  vNewPayingInSlipNo = vParams.IntegerValue("PayingInSlipNumber")
                  If vPISList.Length > 0 Then vPISList.Append(",")
                  vPISList.Append(vBatchNumber.ToString)
                  dgr.SetValue(vIndex, "PayingInSlipNumber", vNewPayingInSlipNo.ToString)
                  If vDetailCompList.Length > 0 Then vDetailCompList.Append(",")
                  vDetailCompList.Append(vBatchNumber.ToString)
                End If
              Else
                If vBatchNumber > 0 Then
                  If vDetailCompList.Length > 0 Then vDetailCompList.Append(",")
                  vDetailCompList.Append(vBatchNumber.ToString)
                End If
              End If
            Next
            If vPISList.Length > 0 AndAlso vNewPayingInSlipNo > 0 Then
              vParams("BatchNumber") = vPISList(0)
              vParams("BatchNumbers") = vPISList.ToString
              DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoBatchDetails, vParams)
            End If
            Try
              If vDetailCompList.Length > 0 Then
                vParams("BatchNumber") = vDetailCompList.ToString.Split(","c)(0)
                vParams("BatchNumbers") = vDetailCompList.ToString
                DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoDetailCompleted, vParams)
                RefreshData()
              End If
            Catch vEx As Exception
              Throw vEx
            End Try
          End If
        Case BatchStates.CreateClaim
          If dgr.GetSelectedRowNumbers.Count > 0 Then
            'First deal with batch types to be combined into one claim file.
            If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_claim_method) = "BRIF" Then
              Dim vBatchNumbers As String = SelectedBatchNumbers("CC")
              If vBatchNumbers.Length > 0 Then
                ProcessClaimFile(0, 0, "CC", False, vBatchNumbers)
              End If
              vBatchNumbers = SelectedBatchNumbers("CI")
              If vBatchNumbers.Length > 0 Then
                ProcessClaimFile(0, 0, "CI", False, vBatchNumbers)
              End If
            End If

            Dim vSelectedBatchNumbers As New Dictionary(Of Integer, Integer)
            For Each vIndex As Integer In dgr.GetSelectedRowNumbers
              If (dgr.GetValue(vIndex, "BatchType") <> "CC" AndAlso
                  dgr.GetValue(vIndex, "BatchType") <> "CI") OrElse
                AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_claim_method) <> "BRIF" Then
                vSelectedBatchNumbers.Add(IntegerValue(dgr.GetValue(vIndex, "BatchNumber")), vIndex)
              End If
            Next vIndex
            If vSelectedBatchNumbers.Count > 0 Then
              ProcessBatch(mvBatchState, vSelectedBatchNumbers(vSelectedBatchNumbers.Keys.Min), vSelectedBatchNumbers(vSelectedBatchNumbers.Keys.Max))
            End If
          End If
        Case Else
          If dgr.MultipleSelect = False Then
            If mvBatchNumber > 0 Then ProcessBatch(mvBatchState, mvBatchNumber)
          ElseIf dgr.GetSelectedRowNumbers.Count > 0 Then
            For Each vBatchNumber As String In SelectedBatchNumbers.Split(","c)
              If vBatchNumber.Length > 0 Then ProcessBatch(mvBatchState, IntegerValue(vBatchNumber))
            Next
          End If

      End Select
    Catch vEx As Exception
      DataHelper.HandleJobException(vEx)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      Dim vIndex As Integer
      With dgr
        If .MultipleSelect = False Then
          DeleteBatch(.CurrentRow)
        Else
          If .GetSelectedRowNumbers.Count > 0 Then
            Dim vList As ArrayListEx = .GetSelectedRowNumbers
            For Each vIndex In vList
              DeleteBatch(vIndex)
            Next
          End If
        End If
        RefreshData()
      End With
    Catch vEx As Exception
      DataHelper.HandleJobException(vEx)
    End Try
  End Sub

  Private Sub DeleteBatch(ByVal pRow As Integer)
    With dgr
      Dim vBatchNumber As Integer = IntegerValue(.GetValue(pRow, "BatchNumber"))
      If vBatchNumber > 0 Then
        If ShowQuestion(QuestionMessages.QmDeleteBatch, MessageBoxButtons.YesNo, vBatchNumber.ToString) = System.Windows.Forms.DialogResult.Yes Then
          Dim vList As New ParameterList(True)
          vList("BatchNumber") = vBatchNumber.ToString

          Dim vResult As ParameterList = DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctBatches, vList)
          If vResult.Contains("Confirm") Then
            If ShowQuestion(QuestionMessages.QmDeletePostedBatch, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vList("Confirm") = "Y"
              DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctBatches, vList)
            End If
          End If
        End If
      End If
    End With
  End Sub

  Private Sub cmdRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    RefreshData()
  End Sub

  Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub
  Public Sub ProcessPIS()
    ProcessPIS(True)
  End Sub
  Private Sub ProcessPIS(ByVal pPrintOnly As Boolean)
    Dim vList As New ParameterList(True)
    vList("JobName") = FormHelper.GetTaskJobTypeName(CareServices.TaskJobTypes.tjtPayingInSlips)
    Dim vResult As DialogResult = System.Windows.Forms.DialogResult.No
    If pPrintOnly = False Then
      vResult = FormHelper.ScheduleTask(vList)
      If vResult = System.Windows.Forms.DialogResult.Yes Then
        If mvManualPayingInSlips And mvManualPISLocation = "PROCESSPIS" Then
          'user chose to schedule the task, let the know that the PIS number will still be allocated now.
          vResult = ShowQuestion(QuestionMessages.QmGeneratePayingInSlipNumbers, MessageBoxButtons.OKCancel)
        End If
      End If
    End If
    If vResult <> System.Windows.Forms.DialogResult.Cancel Then
      'Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptPISPrinting)
      Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.TaskJobTypes.tjtPayingInSlips)
      If vParams.Count > 0 Then
        If mvManualPayingInSlips And mvManualPISLocation = "PROCESSPIS" Then
          'a pis number must be provided when the above two configs are set
          If Not vParams.Contains("PayingInSlipNumber") Then
            ShowErrorMessage(QuestionMessages.QmMissingPayingInSlipNumber) ' daePISMissing
            Exit Sub
          End If

          Dim vCountList As New ParameterList(True)
          vCountList("PayingInSlipNumber") = vParams("PayingInSlipNumber")
          If vParams.Contains("BankAccount") Then vCountList("BankAccount") = vParams("BankAccount")
          If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctBPPayingInSlip, vCountList) = 0 Then
            ' this is a new PIS number, assign this to selected batches
            'get a list of selected batch numbers
            If dgr.GetSelectedRowNumbers.Count > 0 Then
              Dim vBatchNumbers As String = SelectedBatchNumbers()
              If vBatchNumbers.Length > 0 Then
                Dim vTemp As New StringBuilder
                For Each vBatchNumber As String In vBatchNumbers.Split(","c)
                  'check that all of the batches are ready for PIS. if some are not then error
                  Dim vBatchInfo As New BatchInfo(IntegerValue(vBatchNumber))
                  With vBatchInfo
                    If (.BatchType <> CareServices.BatchTypes.Cash AndAlso .BatchType <> CareServices.BatchTypes.CashWithInvoice) OrElse .PostedToCashBook OrElse .PayingInSlipPrinted OrElse .PayingInSlipNumber <> "" OrElse (.PrintChequeList AndAlso .ReadyForBanking = False) _
                    OrElse (.PrintChequeList = False AndAlso (.ReadyForBanking = False OrElse .NumberOfEntries = 0 OrElse .NumberOfEntries <> .NumberOfTransactions OrElse .BatchTotal <> .TransactionTotal)) Then
                      If vTemp.Length > 0 Then vTemp.Append(",")
                      vTemp.Append(vBatchNumber)
                    End If
                  End With
                Next
                If vTemp.Length > 0 Then
                  ShowErrorMessage(InformationMessages.ImBatchedNotReadyForPrinting, vTemp.ToString)
                  Exit Sub
                End If
                Dim vNewList As New ParameterList(True)
                vNewList("PayingInSlipNumber") = vParams("PayingInSlipNumber")
                If vParams.Contains("BankAccount") Then vNewList("BankAccount") = vParams("BankAccount")
                vNewList("PayingInSlipPrinted") = "Y"
                vNewList("BatchNumber") = vBatchNumbers.ToString.Split(","c)(0)
                vNewList("BatchNumbers") = vBatchNumbers
                DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoBatchDetails, vNewList)
              End If
            End If
          End If
        End If
        If vParams.Contains("PayingInSlipNumber") Then vList("PayingInSlipNumber") = vParams("PayingInSlipNumber")
        If vParams.Contains("BankAccount") Then vList("BankAccount") = vParams("BankAccount")
        vList("ReportDestination") = vParams("ReportDestination")
        Dim vScheduleType As FormHelper.ProcessTaskScheduleType = FormHelper.ProcessTaskScheduleType.ptsAlwaysRun
        If vResult <> System.Windows.Forms.DialogResult.No Then vResult = System.Windows.Forms.DialogResult.None
        FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPayingInSlips, vList, False, vScheduleType, True)
      End If
    End If
  End Sub

  Function CheckCashBookBatchNumbers() As Boolean
    Dim vResult As DialogResult
    Dim vBankAccounts As String = mvDataTable.Rows(0)("BankAccounts").ToString
    If vBankAccounts.Length > 0 Then
      For Each vBankAccount As String In vBankAccounts.Split(","c)
        vResult = ShowQuestion(QuestionMessages.QmCashBookNumberIsGreater, MessageBoxButtons.OKCancel, vBankAccount)
        If vResult = System.Windows.Forms.DialogResult.Cancel Then Exit For
      Next
    End If
    Return vResult <> System.Windows.Forms.DialogResult.Cancel
  End Function
  Private Function SelectedBatchNumbers() As String
    Return SelectedBatchNumbers("")
  End Function
  Private Function SelectedBatchNumbers(ByVal pFilter As String) As String
    Dim vBatchNumbers As New StringBuilder
    If dgr.GetSelectedRowNumbers.Count > 0 Then
      Dim vList As ArrayListEx = dgr.GetSelectedRowNumbers
      Dim vAppend As Boolean = True
      For Each vIndex As Integer In vList
        If pFilter.Length > 0 Then
          Dim vColumnName As String = ""
          Select Case mvBatchState
            Case BatchStates.CreateClaim
              vColumnName = "BatchType"
          End Select
          If dgr.GetValue(vIndex, vColumnName) = pFilter Then vAppend = True Else vAppend = False
        End If
        If vAppend Then
          If vBatchNumbers.Length > 0 Then vBatchNumbers.Append(",")
          vBatchNumbers.Append(IntegerValue(dgr.GetValue(vIndex, "BatchNumber")))
        End If
      Next
    End If
    Return vBatchNumbers.ToString
  End Function

  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    mvSelectedRow = pRow
    If mvSelectedRow >= 0 AndAlso dgr.RowCount > mvSelectedRow Then mvBatchNumber = IntegerValue(dgr.GetValue(mvSelectedRow, "BatchNumber"))
  End Sub


  Private Sub ProcessBatchUpdate(ByVal pBatchNumbers As String, ByVal pParams As ParameterList, ByVal pResult As DialogResult)

    Dim vBatchNumbers() As String = pBatchNumbers.Split(","c)
    If pResult = System.Windows.Forms.DialogResult.Yes Then
      pParams("BatchNumber") = vBatchNumbers(0)
      pParams("BatchNumbers") = pBatchNumbers
      FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBatchUpdate, pParams, False, FormHelper.ProcessTaskScheduleType.ptsNone, False) 'Dont run asynchronously as scheduling
    Else
      For vIndex As Integer = 0 To UBound(vBatchNumbers)
        Dim vParams As New ParameterList(True)
        vParams("BatchNumber") = vBatchNumbers(vIndex)
        FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBatchUpdate, vParams, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
      Next
    End If
  End Sub

  Public Function PrintChequeList(ByVal pBatchNumber As Integer) As Boolean
    Dim vList As New ParameterList(True)
    Dim vCancelled As Boolean
    Dim vBatch As New BatchInfo(pBatchNumber)
    Dim vReturn As Boolean

    vList("BatchNumber") = pBatchNumber.ToString
    DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoLockBatch, vList)
    vList("RPbatch_number") = pBatchNumber.ToString
    vList("RPpm_cash") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash, "CASH")
    vList("ReportCode") = "CHQLST"
    Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSaveAndNoOutput, vCancelled)

    vList = New ParameterList(True)
    vList("BatchNumber") = pBatchNumber.ToString
    If vCancelled = False Then
      If vBatch.NumberOfEntries = vBatch.NumberOfTransactions AndAlso vBatch.BatchTotal = vBatch.TransactionTotal Then
        vList("ReadyForBanking") = "Y"
        DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoBatchDetails, vList)
        vReturn = True
      End If
    End If
    DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoUnlockBatch, vList)
    Return vReturn
  End Function

  Private Sub ProcessClaimFile(ByVal pBatchNumber As Integer, ByVal pLastBatchNumber As Integer, ByVal pBatchType As String, Optional ByVal pProvisionalBatch As Boolean = False, Optional ByVal pBatchNumbersList As String = "")
    Try
      Dim vList As New ParameterList(True)

      If pBatchNumbersList.Length > 0 Then
        vList("BatchNumbersList") = pBatchNumbersList
      ElseIf pLastBatchNumber > 0 Then
        vList.IntegerValue("BatchNumber") = pLastBatchNumber
        vList.IntegerValue("BatchNumber2") = pBatchNumber
      Else
        vList.IntegerValue("BatchNumber") = pBatchNumber
        pLastBatchNumber = pBatchNumber
      End If

      If pProvisionalBatch Then
        Select Case pBatchType
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.CAFVouchers)
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFProvisionalBatchClaim, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.CAFCards)
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFCardSalesReport, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        End Select
      Else
        Select Case pBatchType
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.CreditCard), AppValues.GetBatchTypeCode(CareServices.BatchTypes.DebitCard), AppValues.GetBatchTypeCode(CareServices.BatchTypes.CreditCardWithInvoice)
            If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_claim_method) = "BRIF" And pBatchNumbersList.Length > 0 Then
              FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCardSalesFile, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
            Else
              If ShowQuestion(QuestionMessages.QmGenerateManualClaim, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCardSalesReport, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
              Else
                FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCardSalesFile, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
              End If
            End If
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.CreditCardAuthority)
            If ShowQuestion(QuestionMessages.QmGenerateManualClaim, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCCClaimReport, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
            Else
              FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCCClaimFile, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
            End If
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.DirectDebit)
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDDClaimFile, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          Case AppValues.GetBatchTypeCode(CareServices.BatchTypes.DirectCredit)
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDDCreditFile, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        End Select
      End If
    Catch vEx As Exception
      DataHelper.HandleJobException(vEx)
    End Try
  End Sub

  Private Sub ViewBatchDetails(pRow As DataRow)
    Dim vList As New ParameterList(True)
    vList("BatchNumber") = mvBatchNumber.ToString
    Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctBatchDetails, pRow)
    RemoveHandler vForm.EditBatchDetails, AddressOf FormHelper.EditBatchDetails
    RemoveHandler vForm.BatchEditComplete, AddressOf BatchEditComplete
    AddHandler vForm.EditBatchDetails, AddressOf FormHelper.EditBatchDetails
    AddHandler vForm.BatchEditComplete, AddressOf BatchEditComplete
    vForm.TopMost = False
    vForm.Show(Me)
  End Sub

  Private Sub BatchEditComplete(ByVal sender As Object, ByVal pBatchNumber As Integer)
    'Unlock the batch then refresh the screen
    FormHelper.BatchEditComplete(sender, pBatchNumber)
    RefreshData()
  End Sub

End Class