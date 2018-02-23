Imports System.Linq
Imports System.ComponentModel

Public Class frmTrader
  Inherits MaintenanceParentForm

  Public Overrides ReadOnly Property SizeMaintenanceForm() As Boolean
    Get
      Return True
    End Get
  End Property

#Region " Windows Form Designer generated code "

  Public Sub New(ByVal pTraderApplication As TraderApplication)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pTraderApplication)
  End Sub

  Public Sub New(ByVal pTraderApplication As TraderApplication, ByVal pTransactionsForm As frmTraderTransactions)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pTraderApplication, pTransactionsForm)
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
  Friend WithEvents bpl As ButtonPanel
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdEdit As System.Windows.Forms.Button
  Friend WithEvents cmdPrevious As System.Windows.Forms.Button
  Friend WithEvents cmdNext As System.Windows.Forms.Button
  Friend WithEvents cmdFinished As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents sbr As System.Windows.Forms.StatusBar
  Friend WithEvents prgBar As System.Windows.Forms.ProgressBar
  Friend WithEvents sbp As System.Windows.Forms.StatusBarPanel

  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTrader))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdEdit = New System.Windows.Forms.Button()
    Me.cmdPrevious = New System.Windows.Forms.Button()
    Me.cmdNext = New System.Windows.Forms.Button()
    Me.cmdFinished = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.pnl = New System.Windows.Forms.Panel()
    Me.prgBar = New System.Windows.Forms.ProgressBar()
    Me.sbr = New System.Windows.Forms.StatusBar()
    Me.sbp = New System.Windows.Forms.StatusBarPanel()
    Me.bpl.SuspendLayout()
    CType(Me.sbp, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdEdit)
    Me.bpl.Controls.Add(Me.cmdPrevious)
    Me.bpl.Controls.Add(Me.cmdNext)
    Me.bpl.Controls.Add(Me.cmdFinished)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 433)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(800, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.Location = New System.Drawing.Point(74, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 0
    Me.cmdDelete.Text = "&Delete"
    '
    'cmdEdit
    '
    Me.cmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdEdit.Location = New System.Drawing.Point(185, 6)
    Me.cmdEdit.Name = "cmdEdit"
    Me.cmdEdit.Size = New System.Drawing.Size(96, 27)
    Me.cmdEdit.TabIndex = 1
    Me.cmdEdit.Text = "&Edit"
    '
    'cmdPrevious
    '
    Me.cmdPrevious.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrevious.Location = New System.Drawing.Point(296, 6)
    Me.cmdPrevious.Name = "cmdPrevious"
    Me.cmdPrevious.Size = New System.Drawing.Size(96, 27)
    Me.cmdPrevious.TabIndex = 2
    Me.cmdPrevious.Text = "<< &Previous"
    '
    'cmdNext
    '
    Me.cmdNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNext.Location = New System.Drawing.Point(407, 6)
    Me.cmdNext.Name = "cmdNext"
    Me.cmdNext.Size = New System.Drawing.Size(96, 27)
    Me.cmdNext.TabIndex = 3
    Me.cmdNext.Text = "&Next >>"
    '
    'cmdFinished
    '
    Me.cmdFinished.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFinished.Location = New System.Drawing.Point(518, 6)
    Me.cmdFinished.Name = "cmdFinished"
    Me.cmdFinished.Size = New System.Drawing.Size(96, 27)
    Me.cmdFinished.TabIndex = 4
    Me.cmdFinished.Text = "&Finished"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(629, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "Cancel"
    '
    'pnl
    '
    Me.pnl.BackColor = System.Drawing.SystemColors.Control
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(800, 394)
    Me.pnl.TabIndex = 0
    '
    'prgBar
    '
    Me.prgBar.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.prgBar.Location = New System.Drawing.Point(0, 394)
    Me.prgBar.Name = "prgBar"
    Me.prgBar.Size = New System.Drawing.Size(800, 14)
    Me.prgBar.Step = 1
    Me.prgBar.TabIndex = 0
    Me.prgBar.Visible = False
    '
    'sbr
    '
    Me.sbr.Location = New System.Drawing.Point(0, 408)
    Me.sbr.Name = "sbr"
    Me.sbr.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbp})
    Me.sbr.ShowPanels = True
    Me.sbr.Size = New System.Drawing.Size(800, 25)
    Me.sbr.SizingGrip = False
    Me.sbr.TabIndex = 1
    '
    'sbp
    '
    Me.sbp.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbp.Name = "sbp"
    Me.sbp.Width = 800
    '
    'frmTrader
    '
    Me.AcceptButton = Me.cmdNext
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(800, 472)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.prgBar)
    Me.Controls.Add(Me.sbr)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTrader"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    CType(Me.sbp, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub

#End Region

#Region " Declarations "
  Private mvTA As TraderApplication
  Private mvTraderPages As New CollectionList(Of TraderPage)
  Private mvCurrentPage As TraderPage
  Private mvLastPage As TraderPage
  Private mvInvoicesDGR As DisplayGrid        'Invoice Payment Grid
  Private mvCBXDGR As DisplayGrid             'DisplayGrid on PCP page
  Private mvTASDGR As DisplayGrid             'DisplayGrid on TAS page
  Private mvPPSDGR As DisplayGrid             'DisplayGrid on PPS page
  Private mvOPSDGR As DisplayGrid             'DisplayGrid on SCP page
  Private mvMembersDGR As DisplayGrid         'DisplayGrid on MMS page
  Private mvPOSDGR As DisplayGrid             'DisplayGrid on POS page
  Private mvPISDGR As DisplayGrid             'DisplayGrid on PIS page
  Private mvPPADGR As DisplayGrid             'DisplayGrid on PPA page
  Private mvOSPDGR As DisplayGrid             'Display Grid On OSP Page
  Private mvInvoiceGrid As DisplayGrid        'Invoice Production Grid
  Private mvStatementGrid As DisplayGrid      'StatementGrid on STL Page
  Private mvCMTOldPPD As DisplayGrid          'Display Grid for old PPD on MTC page
  Private mvCMTNewPPD As DisplayGrid          'Display Grid for new PPD on MTC page
  Private mvAddBtn As Button                  'Add button on MMS page
  Private mvFindBtn As Button                 'Find button on MMS page
  Private mvRemoveBtn As Button               'Remove button on MMS page
  Private mvAmendBtn As Button                'Amend button on MMS page
  Private mvOSInvoices As CollectionList(Of InvoiceInfo)
  Private mvCashInvoices As CollectionList(Of InvoiceInfo)
  Private mvNewContacts As CollectionList(Of ContactInfo)
  Private mvValidDisputeCodes As CollectionList(Of String)
  Private mvCurrentRow As Integer
  Private mvCurrentPPDLine As Integer
  Private mvCMDFileName As String = ""
  Private mvTransactionsForm As frmTraderTransactions
  Private mvEventWLPriceZeroed As Boolean
  Private mvExtApplication As ExternalApplication

  Private mvMembConv As Boolean = False
  Private mvConfirmCancel As Boolean = True
  Private mvNonFinancialBatchNumber As Integer = 0
  Private mvNonFinancialTransactionNumber As Integer = 0
  Private mvPPNumber As Integer = 0
  Private mvSuppressEvents As Boolean
  Private mvNewPageType As Nullable(Of CareServices.TraderPageType)
  Private mvNextPageCode As String
  Private mvFirstTimeOnPM1 As Boolean
  Private mvTPPDone As Boolean
  Private mvPM1Caption As String
  Private mvTotalStock As Integer
  Private mvOldSourceCode As String
  Private mvOldProductCode As String
  Private mvOldRate As String
  Private mvCarraigePercentage As Double
  Private mvPAPDone As Boolean
  Private mvOrgAmountText As String
  Private mvAccountSelected As Boolean
  Private mvIsInvalidCMT As Boolean
  Private mvServiceBookingNumber As Integer
  Private mvCardDetails As ParameterList
  Private mvCardAuthorised As Boolean
  Private mvProcessingExams As Boolean = False
  Private mvCloseMe As Boolean = False
  Private Const SAGEPAYHOSTED = "SAGEPAYHOSTED"
  Private mvCancelSagepay As Boolean = False

  Private Enum RefreshTypes
    rtEventBooking
  End Enum
#End Region
#Region " Initialization "

  Private Sub frmTrader_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If mvTransactionsForm IsNot Nothing Then mvTransactionsForm.Enabled = False
  End Sub

  Private Sub InitialiseControls(ByVal pTraderApplication As TraderApplication)
    InitialiseControls(pTraderApplication, Nothing)
  End Sub

  Private Sub InitialiseControls(ByVal pTraderApplication As TraderApplication, ByVal pTransactionForm As frmTraderTransactions)
    SetControlTheme()
    Me.cmdDelete.Text = ControlText.CmdnDelete
    Me.cmdEdit.Text = ControlText.CmdDocumentEdit
    Me.cmdPrevious.Text = ControlText.CmdPreviousStep
    Me.cmdNext.Text = ControlText.CmdNextStep
    Me.cmdFinished.Text = ControlText.CmdFinished
    Me.cmdCancel.Text = ControlText.CmdCancel

    MainHelper.SetMDIParent(Me)
    mvTA = pTraderApplication
    mvTransactionsForm = pTransactionForm
    mvTraderPages = pTraderApplication.Pages

    Me.Text = mvTA.Description

    For Each vPage As TraderPage In mvTA.Pages
      Dim vEPL As EditPanel = vPage.EditPanel
      If vPage.Menu Then AddHandler vEPL.ButtonClicked, AddressOf ButtonClicked
      AddHandler vEPL.ShowStatusMessage, AddressOf EPL_ShowMessage
      AddHandler vEPL.ValidateItem, AddressOf EPL_ValidateItem
      AddHandler vEPL.ValueChanged, AddressOf EPL_ValueChanged
      AddHandler vEPL.GetCodeRestrictions, AddressOf EPL_GetCodeRestrictions
      AddHandler vEPL.GetInitialCodeRestrictions, AddressOf EPL_GetInitialCodeRestrictions
      AddHandler vEPL.ValidateAllItems, AddressOf EPL_ValidateAllItems
      AddHandler vEPL.ContactSelected, AddressOf EPL_ContactSelected
      Select Case (vPage.PageType)
        Case CareServices.TraderPageType.tpContactSelection, CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpChangeMembershipType
          AddHandler vEPL.MembershipSelected, AddressOf EPL_MembershipSelected
        Case CareNetServices.TraderPageType.tpTransactionDetails
          'Dim externalReference As TextBox = vEPL.FindPanelControl(Of TextBox)("ExternalReference", False)
          'If externalReference IsNot Nothing Then
          '  AddHandler externalReference.Validating, AddressOf ExternalReference_Validating
          'End If
        Case CareServices.TraderPageType.tpConfirmProvisionalTransactions
          AddHandler vEPL.ProductNumberSelected, AddressOf EPL_ProductNumberSelected
        Case CareNetServices.TraderPageType.tpExamBooking
          AddHandler vEPL.CheckedItemsChanged, AddressOf EPL_CheckedItemsChanged
        Case CareServices.TraderPageType.tpCardDetails
          If WebBasedCardAuthoriser.IsAvailable Then
            Dim browserControl As WebBrowser = vEPL.FindPanelControl(Of WebBrowser)("None2", False)
            If browserControl IsNot Nothing Then
              Me.CardAuthoriser = WebBasedCardAuthoriser.GetInstance(browserControl)
              AddHandler CardAuthoriser.ProcessingComplete, AddressOf CardAuthorisationComplete
            End If
          End If
      End Select
      Select Case (vPage.PageType)
        Case CareServices.TraderPageType.tpContactSelection, CareServices.TraderPageType.tpTransactionDetails
          'set the mailing code to me mandatory as the force mailing code to be mandatory is set in trader app maint
          If mvTA.MailingCodeMandatory AndAlso vEPL.PanelInfo.PanelItems("Mailing").Visible Then vEPL.PanelInfo.PanelItems("Mailing").Mandatory = True
      End Select

      pnl.Controls.Add(vEPL)
    Next

    If mvTA.BatchNumber > 0 AndAlso mvTA.TransactionNumber > 0 Then
      SetPage(CareServices.TraderPageType.tpTransactionAnalysisSummary)
      ProcessData(CareServices.TraderProcessDataTypes.tpdtEditTransaction, False)
    ElseIf mvTA.ApplicationType = ApplicationTypes.atPurchaseOrder AndAlso mvTA.PurchaseOrderNumber > 0 Then
      SetPage(CareServices.TraderPageType.tpPurchaseOrderDetails)
      ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, False)
    Else
      mvTA.TransactionPaymentMethod = "CASH"       'Default required for PayMethodsAtEnd
      SetPage(mvTA.MainPageType)
      ProcessData(CareServices.TraderProcessDataTypes.tpdtFirstPage, False)
    End If

    If mvTA.ChangeMembershipType = True AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then
      If mvTA.CMTMemberNumber.Length > 0 Then ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
    End If

    If mvCurrentPage.Menu AndAlso mvCurrentPage.MenuCount = 1 Then mvCurrentPage.EditPanel.DoTraderChoice() 'Only one button so process it
  End Sub

  'Private Property IsProcessingExternalReferenceChange As Boolean = False
  ' ''' <summary>
  ' ''' Handles the Validating event of the ExternalReference control.
  ' ''' </summary>
  ' ''' <param name="sender">The source of the event.</param>
  ' ''' <param name="e">The <see cref="CancelEventArgs"/> instance containing the event data.</param>
  ' ''' <remarks>
  ' ''' If the sending control is an instance of <see cref="TextBox" /> and the 
  ' ''' parent of the sending control is an instance of <see cref="EditPanel" />
  ' ''' and the parent contains an instance of <see cref="TextLookupBox" /> 
  ' ''' named ContactNumber then set the text of the Contact Number control to
  ' ''' the contact number associated with the external reference in the 
  ' ''' sender's Text property.
  ' ''' </remarks>
  'Private Sub ExternalReference_Validating(sender As Object, e As CancelEventArgs)
  '  Try
  '    Me.IsProcessingExternalReferenceChange = True
  '    Dim extRefCtl As TextBox = TryCast(sender, TextBox)
  '    If extRefCtl IsNot Nothing AndAlso
  '       Not String.IsNullOrWhiteSpace(extRefCtl.Text) Then
  '      Dim epl As EditPanel = TryCast(extRefCtl.Parent, EditPanel)
  '      If epl IsNot Nothing Then
  '        Dim contactCtl As TextLookupBox = epl.FindPanelControl(Of TextLookupBox)("ContactNumber", False)
  '        If contactCtl IsNot Nothing Then
  '          epl.ClearControlList("ContactNumber,MailingContactNumber")
  '          epl.SetErrorField(contactCtl.Name, String.Empty)
  '          Dim params As New ParameterList(True, True)
  '          params("ExternalReference") = extRefCtl.Text
  '          params("DataSource") = "URN"
  '          Using contactDataset As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftContacts, params)
  '            If contactDataset.Tables.Contains("DataRow") Then
  '              If contactDataset.Tables("DataRow").Rows.Count = 1 Then
  '                contactCtl.Text = contactDataset.Tables("DataRow").Rows(0).Field(Of String)("ContactNumber")
  '                EPL_ValueChanged(epl, contactCtl.Name, contactCtl.Text)
  '              Else
  '                epl.SetErrorField(extRefCtl.Name, If(contactDataset.Tables("DataRow").Rows.Count > 1,
  '                                                     InformationMessages.ImAmbiguousValue,
  '                                                     InformationMessages.ImInvalidValue))
  '              End If
  '            Else
  '              epl.SetErrorField(extRefCtl.Name, InformationMessages.ImInvalidValue)
  '            End If
  '          End Using
  '        End If
  '      End If
  '    End If
  '  Finally
  '    Me.IsProcessingExternalReferenceChange = False
  '  End Try
  'End Sub
#End Region
#Region " Form Termination "
  Private Sub frmTrader_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Try
      If mvTransactionsForm IsNot Nothing Then
        mvTransactionsForm.Enabled = True
        mvTransactionsForm.GetTransactionData(mvTA.BatchInfo)
      End If
      If mvTA.BatchLedApp AndAlso mvTA.BatchLocked Then 'Only need to unlock if the batch has been locked
        Dim vList As New ParameterList(True)
        vList.IntegerValue("BatchNumber") = mvTA.BatchNumber
        DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoUnlockBatch, vList)
        mvTA.BatchLocked = False
      Else
        MainHelper.EnableTraderApplications(True)
      End If
      If mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance Then
        'Refresh all contact cards
        MainHelper.RefreshData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, 0)
        MainHelper.RefreshData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, 0)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Private Sub frmTrader_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    If mvCloseMe Then
      ProcessData(CareServices.TraderProcessDataTypes.tpdtCancelTransaction, True)
    ElseIf Not DataChanged() OrElse Not mvConfirmCancel OrElse (DataChanged() AndAlso ConfirmCancel()) Then
      Try
        If mvCurrentPage IsNot Nothing Then
          If mvCurrentPage.EditPanel IsNot Nothing Then mvCurrentPage.EditPanel.Visible = False
        End If
        ProcessData(CareServices.TraderProcessDataTypes.tpdtCancelTransaction, True)
      Catch vException As Exception
        DataHelper.HandleException(vException)
      End Try
      For Each vPage As TraderPage In mvTraderPages
        vPage.EditPanel.ClearDataSources(vPage.EditPanel)
      Next
    Else
      mvCancelSagepay = True
      e.Cancel = True
    End If

  End Sub
#End Region

#Region " Buttons "

  Private Sub ButtonClicked(ByVal pSender As Object, ByVal pParameterName As String)
    Try
      Dim vPageType As CareServices.TraderPageType = mvCurrentPage.PageType
      Dim vValid As Boolean = True
      If mvNewPageType.HasValue Then vPageType = mvNewPageType.Value
      Select Case vPageType
        Case CareServices.TraderPageType.tpPaymentMethod1
          vValid = ValidateTransactionPaymentMethod(pParameterName)
          If vValid Then
            mvTA.TransactionPaymentMethod = pParameterName
            mvTA.PPPaymentType = pParameterName
          End If
        Case CareServices.TraderPageType.tpTransactionAnalysis
          mvTA.TransactionType = pParameterName
          If mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "DONR" OrElse mvTA.TransactionType = "SUBS" Then
            'Create Membership/RegularDonation/Subscription
            'Always set mvTA.PaymentPlan to be Nothing (it may have been set from a previous entry)
            mvTA.PaymentPlan = Nothing
          End If
          If mvTA.TransactionType = "MEMC" Then
            If mvTA.PaymentPlan IsNot Nothing Then
              vValid = ValidateCMT()
            Else
              EPL_ShowMessage(mvCurrentPage.EditPanel, InformationMessages.ImPPNorMNNotSpecifed)
              vValid = False
            End If
          End If
        Case CareServices.TraderPageType.tpPaymentMethod2
          ClearPageDefaults(CareServices.TraderPageType.tpPaymentPlanProducts)
          ClearPageDefaults(CareServices.TraderPageType.tpPaymentPlanDetails)
          If pParameterName = "CURR" Then
            mvTA.CurrentPaymentMethod = True
            mvTA.PPPaymentType = mvTA.TransactionPaymentMethod
          Else
            mvTA.CurrentPaymentMethod = False
            mvTA.PPPaymentType = pParameterName
          End If
          mvTA.ClearPPDefinition()
        Case CareServices.TraderPageType.tpPaymentMethod3
          mvTA.TransactionPaymentMethod = pParameterName
          mvTA.PPPaymentType = pParameterName
          mvTA.TransactionType = pParameterName
          If pParameterName = "MEMB" Then mvMembConv = True
        Case CareNetServices.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          mvTA.UnbalancedTransactionChoice = pParameterName
          If pParameterName = "TRAN" Then
            mvTA.TransactionAmount = mvTA.CurrentLineTotal
            Dim vPage As TraderPage = mvTraderPages(CareServices.TraderPageType.tpTransactionDetails.ToString)
            vPage.EditPanel.SetValue("Amount", mvTA.TransactionAmount.ToString("0.00"))
          End If
      End Select
      If vValid AndAlso vPageType = mvCurrentPage.PageType Then
        ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, False)
        If mvCurrentPage.Menu AndAlso mvCurrentPage.MenuCount = 1 Then
          mvCurrentPage.EditPanel.DoTraderChoice() 'Only one button so process it
        ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentMethod2 AndAlso mvTA.TransactionType = "MEMB" AndAlso (mvTA.TransactionPaymentMethod = "CAFC" OrElse mvTA.TransactionPaymentMethod = "VOUC") Then
          'Since CAF transactions only support donations & memberships don't show the PM2 options
          mvCurrentPage.EditPanel.DoTraderChoice()
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'Add button on tpMembershipMembersSummary page
    Try
      Dim vParams As New ParameterList(True)
      vParams("SystemColumns") = "N"
      Dim vCurrentContactNumber As Integer = IntegerValue(mvMembersDGR.GetValue(0, "ContactNumber"))
      Dim vAddressNumber As Integer = IntegerValue(mvMembersDGR.GetValue(0, "AddressNumber"))
      Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, vCurrentContactNumber, vParams))
      Dim vList As New ParameterList()
      vList("Surname") = vRow.Item("Surname").ToString
      vList("Source") = mvTA.TransactionSource
      Dim vResult As DialogResult = System.Windows.Forms.DialogResult.OK
      If BooleanValue(AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.add_member_current_address)) Then
        'Always use the current address of the existing member
      Else
        'Ask the user which address they which to use and allow them to create a new address
        Dim vDS As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vCurrentContactNumber)
        Dim vForm As New frmSelectAddress(vDS, frmSelectAddress.SelectAddressTypes.satTraderCreateMemberContact, vAddressNumber)
        If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
          'If user chose 'New' then vForm.AddressNumber will be zero
          vAddressNumber = vForm.AddressNumber
        Else
          vResult = System.Windows.Forms.DialogResult.Cancel
        End If
      End If
      If vResult = System.Windows.Forms.DialogResult.OK Then
        If vAddressNumber > 0 Then vList("CreateAtAddressNumber") = vAddressNumber.ToString
        Dim vContactNumber As Integer = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctContact, vList, Me, True)
        If vContactNumber > 0 Then
          'Add the Contact to the grid
          mvTA.MemberContactToAdd = vContactNumber
          ProcessData(CareServices.TraderProcessDataTypes.tpdtAddMemberSummary, False)
          mvTA.MemberContactToAdd = 0
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAmend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'Amend button on tpMembershipMembersSummary page
    Try
      mvCurrentRow = mvMembersDGR.CurrentRow
      If mvCurrentRow >= 0 Then
        ClearPageDefaults(CareServices.TraderPageType.tpAmendMembership)
        ProcessData(CareServices.TraderProcessDataTypes.tpdtAmendMemberSummary, False)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try

  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    CancelTransaction()
  End Sub

  Private Sub CancelTransaction()
    If mvTA.OriginalPayerContactNumber > 0 AndAlso mvTA.OriginalPayerContactNumber <> mvTA.PayerContactNumber Then
      Dim vContactInfo As New ContactInfo(mvTA.OriginalPayerContactNumber)
      UserHistory.AddContactHistoryNode(vContactInfo.ContactNumber, vContactInfo.ContactName, vContactInfo.ContactGroup)
    End If
    Me.Close()
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Dim vCursor As New BusyCursor()
    Try
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          mvCurrentPPDLine = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "LineNumber"))
          Dim vPPDLine As Integer = mvCurrentPPDLine
          Dim vPPDDetailNumber As Integer = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "DetailNumber"))
          mvTA.EditPPDetailNumber = vPPDDetailNumber
          Dim vPPDSubscriptionNumber As Integer = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "SubscriptionNumber"))
          mvTA.EditPPDSubscriptionNumber = vPPDSubscriptionNumber
          Dim vRemoveLine As Boolean = True
          If mvTA.PaymentPlan IsNot Nothing AndAlso (mvTA.PaymentPlan.Existing And vPPDSubscriptionNumber > 0) Then
            vRemoveLine = False
            If ShowQuestion(InformationMessages.ImCannotReverseSubCancellation, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource)
              If vParams.Count > 0 Then
                mvTA.CancellationReason = vParams("CancellationReason")
                If vParams.Contains("CancellationSource") Then mvTA.CancellationSource = vParams("CancellationSource")
                ProcessData(CareServices.TraderProcessDataTypes.tpdtDeletePaymentPlanLine, False)
                mvTA.CancellationReason = ""
                mvTA.CancellationSource = ""
                vRemoveLine = True
              End If
            End If
          End If
          If vRemoveLine Then
            mvTA.DeleteDataSetLine(mvTA.PPDDataSet, vPPDLine)
            SetPage(CareServices.TraderPageType.tpPaymentPlanSummary)
            If mvPPSDGR.RowCount > 0 Then
              SetPPDEditable(vPPDLine)
            Else
              SetPPDEditable(0)
            End If
          End If
        Case CareServices.TraderPageType.tpTransactionAnalysisSummary
          DeleteTASLine(mvTASDGR.CurrentRow)
          SetPage(mvCurrentPage.PageType)
        Case CareServices.TraderPageType.tpPurchaseOrderSummary
          mvTA.DeleteDataSetLine(mvTA.POSDataSet, IntegerValue(mvPOSDGR.GetValue(mvPOSDGR.CurrentRow, "LineNumber")))
          SetPage(CareServices.TraderPageType.tpPurchaseOrderSummary)
        Case CareServices.TraderPageType.tpPurchaseInvoiceSummary
          mvTA.DeleteDataSetLine(mvTA.PISDataSet, IntegerValue(mvPISDGR.GetValue(mvPISDGR.CurrentRow, "LineNumber")))
          SetPage(CareServices.TraderPageType.tpPurchaseInvoiceSummary)
      End Select

    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub

  Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
    Try
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          mvCurrentPPDLine = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "LineNumber"))
          mvTA.EditPPDetailNumber = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "DetailNumber"))
          mvTA.EditPPDSubscriptionNumber = IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "SubscriptionNumber"))
          mvCurrentRow = mvPPSDGR.CurrentRow
          If mvCurrentRow >= 0 Then ProcessData(CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine, True)

        Case CareServices.TraderPageType.tpPurchaseInvoiceSummary
          mvCurrentPPDLine = IntegerValue(mvPISDGR.GetValue(mvPISDGR.CurrentRow, "LineNumber"))
          mvCurrentRow = mvPISDGR.CurrentRow
          If mvCurrentRow >= 0 Then ProcessData(CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine, True)

        Case CareServices.TraderPageType.tpPurchaseOrderSummary
          mvCurrentPPDLine = IntegerValue(mvPOSDGR.GetValue(mvPOSDGR.CurrentRow, "LineNumber"))
          mvCurrentRow = mvPOSDGR.CurrentRow
          If mvCurrentRow >= 0 Then ProcessData(CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine, True)
        Case Else
          mvCurrentRow = mvTASDGR.CurrentRow
          If mvCurrentRow >= 0 Then
            Dim vDataTable As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
            Dim vDataRow As DataRow = vDataTable.Rows(mvCurrentRow)
            Dim vEventBookingNumber As Integer
            Dim vExamBookingNumber As Integer
            If vDataRow.Item("TraderLineType").ToString = "E" Then
              vEventBookingNumber = IntegerValue(vDataRow.Item("EventBookingNumber").ToString)
            ElseIf vDataRow.Item("TraderLineType").ToString = "Q" Then
              vExamBookingNumber = IntegerValue(vDataRow.Item("ExamBookingNumber").ToString)
            End If
            ClearPageDefaults(CareServices.TraderPageType.tpPayments)
            ClearPageDefaults(CareServices.TraderPageType.tpOutstandingScheduledPayments)
            ClearPageDefaults(CareServices.TraderPageType.tpCollectionPayments)
            ClearPageDefaults(CareServices.TraderPageType.tpProductDetails)
            ClearPageDefaults(CareServices.TraderPageType.tpAmendEventBooking)
            mvTA.DeleteAnalysisLine(mvCurrentRow, True)
            If vDataTable.Columns.Contains("TraderTransactionType") Then
              mvTA.TransactionType = vDataRow.Item("TraderTransactionType").ToString
            End If
            ProcessData(CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine, True)
            Dim vEditLineNumber As Integer = mvCurrentRow
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.StockSales Then
              If mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then
                Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
                vEditLineNumber = (IntegerValue(vRow("LineNumber").ToString) - 1)
              End If
            End If
            mvTA.DeleteAnalysisLine(mvCurrentRow, vEventBookingNumber, vExamBookingNumber)
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.StockSales Then mvTA.TransactionLines = vEditLineNumber
          End If

      End Select

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      Dim vList As New ParameterList
      Dim vContactNumber As Integer
      Dim vNewContactAtOrg As Boolean
      Dim vRow As Integer
      vList("Source") = mvTA.TransactionSource

      vContactNumber = (ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, vList, Me, True, vNewContactAtOrg))
      If vContactNumber > 0 Then
        vRow = mvMembersDGR.FindRow("ContactNumber", vContactNumber.ToString)
        If vRow >= 0 Then
          'Contact already exists in grid
          mvMembersDGR.SelectRow(vRow)
          My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        Else
          'Add the Contact to the grid
          mvTA.MemberContactToAdd = vContactNumber
          ProcessData(CareServices.TraderProcessDataTypes.tpdtAddMemberSummary, False)
          mvTA.MemberContactToAdd = 0
        End If

      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try

  End Sub

  Private Sub cmdFinished_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFinished.Click
    ProcessFinish()
  End Sub

  Private Sub ProcessFinish()
    Try
      ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
      If mvCurrentPage.Menu AndAlso mvCurrentPage.MenuCount = 1 Then mvCurrentPage.EditPanel.DoTraderChoice() 'Only one button so process it
      If mvTA.OriginalPayerContactNumber > 0 AndAlso mvTA.OriginalPayerContactNumber <> mvTA.PayerContactNumber Then
        Dim vContactInfo As New ContactInfo(mvTA.OriginalPayerContactNumber)
        UserHistory.AddContactHistoryNode(vContactInfo.ContactNumber, vContactInfo.ContactName, vContactInfo.ContactGroup)
        MainHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders, vContactInfo.ContactNumber)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Private Sub cmdNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNext.Click
    Try
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpTransactionAnalysisSummary
          ClearPageDefaults(CareServices.TraderPageType.tpProductDetails)
          ClearPageDefaults(CareServices.TraderPageType.tpEventBooking)
          ClearPageDefaults(CareServices.TraderPageType.tpExamBooking)
          ClearPageDefaults(CareServices.TraderPageType.tpPayments)
          ClearPageDefaults(CareServices.TraderPageType.tpMembership) ', True)          'Clear all fields ready for next membership
          ClearPageDefaults(CareServices.TraderPageType.tpChangeMembershipType)      'CMT
          ClearPageDefaults(CareServices.TraderPageType.tpAmendMembership)
          ClearPageDefaults(CareServices.TraderPageType.tpMembershipPayer)
          ClearPageDefaults(CareServices.TraderPageType.tpStandingOrder)
          ClearPageDefaults(CareServices.TraderPageType.tpDirectDebit)
          ClearPageDefaults(CareServices.TraderPageType.tpCreditCardAuthority)
          ClearPageDefaults(CareServices.TraderPageType.tpLegacyBequestReceipt)
          ClearPageDefaults(CareServices.TraderPageType.tpScheduledPayments)
          ClearPageDefaults(CareServices.TraderPageType.tpOutstandingScheduledPayments)
          ClearPageDefaults(CareServices.TraderPageType.tpAccommodationBooking)
          ClearPageDefaults(CareServices.TraderPageType.tpCollectionPayments)
          ClearPageDefaults(CareServices.TraderPageType.tpSetStatus)
          ClearPageDefaults(CareServices.TraderPageType.tpGiftAidDeclaration)
          ClearPageDefaults(CareServices.TraderPageType.tpGoneAway)
          ClearPageDefaults(CareServices.TraderPageType.tpAddressMaintenance)
          ClearPageDefaults(CareNetServices.TraderPageType.tpServiceBooking) ', True)      'BR9827:Passed True as pClearFields parameter because SetDefaults clears Amount but leaves Rate set
          ClearPageDefaults(CareServices.TraderPageType.tpCancelPaymentPlan)
          ClearPageDefaults(CareServices.TraderPageType.tpCancelGiftAidDeclaration)
        Case CareServices.TraderPageType.tpOutstandingScheduledPayments
          ClearPageDefaults(CareServices.TraderPageType.tpOutstandingScheduledPayments)
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          ClearPageDefaults(CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance)
          mvCurrentRow = -1
        Case CareServices.TraderPageType.tpPurchaseOrderSummary
          ClearPageDefaults(CareServices.TraderPageType.tpPurchaseOrderProducts)
        Case CareNetServices.TraderPageType.tpPostageAndPacking
          ClearPageDefaults(CareNetServices.TraderPageType.tpPostageAndPacking)
      End Select

      If mvTA.MultiCurrency() And mvTA.BatchCurrencyCode IsNot Nothing Then
        If mvCurrentPage.EditPanel.Controls("Rate") IsNot Nothing AndAlso mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow() IsNot Nothing Then
          mvTA.BatchCurrencyCode = mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow().Item("CurrencyCode").ToString
        End If
      End If

      ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
      If mvCurrentPage.Menu AndAlso mvCurrentPage.MenuCount = 1 Then mvCurrentPage.EditPanel.DoTraderChoice() 'Only one button so process it
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdPrevious_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrevious.Click
    Try
      Dim vAddValues As Boolean = False
      Select Case mvCurrentPage.PageType
        'Going back from product page clear the defaults so it is set correctly next time (product sales <-> donation)
        Case CareNetServices.TraderPageType.tpProductDetails, CareNetServices.TraderPageType.tpOutstandingScheduledPayments, CareNetServices.TraderPageType.tpCardDetails
          ClearPageDefaults(mvCurrentPage.PageType)
      End Select
      If mvTASDGR IsNot Nothing AndAlso mvTASDGR.RowCount > 0 AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionAnalysisSummary _
         AndAlso mvTASDGR.GetValue(mvTASDGR.RowCount - 1, "PaymentPlanNumber") IsNot Nothing AndAlso IntegerValue(mvTASDGR.GetValue(mvTASDGR.RowCount - 1, "PaymentPlanNumber")) > 0 _
         AndAlso mvTA.PPNumbersCreated IsNot Nothing AndAlso mvTA.PPNumbersCreated.ContainsKey(mvTASDGR.GetValue(mvTASDGR.RowCount - 1, "PaymentPlanNumber")) = False _
         AndAlso Not mvTA.TransactionType = "APAY" Then
        DeleteTASLine(mvTASDGR.RowCount - 1)
        ClearPageDefaults(CareServices.TraderPageType.tpOutstandingScheduledPayments)
      End If
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpStandingOrder, CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpCreditCardAuthority
          If mvTA.TransactionType = "APAY" Then mvCurrentPage.DefaultsSet = False
        Case CareNetServices.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          mvTPPDone = False
        Case CareNetServices.TraderPageType.tpExamBooking
          mvProcessingExams = False   'Just in case!!
        Case CareNetServices.TraderPageType.tpComments
          vAddValues = True  'Notes from this page is needed to set transaction note
      End Select
      ProcessData(CareServices.TraderProcessDataTypes.tpdtPreviousPage, vAddValues)
      If mvCurrentPage.Menu AndAlso mvCurrentPage.MenuCount = 1 Then          'Only one button so process it
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentMethod2 Then
          SetPage(CareServices.TraderPageType.tpTransactionAnalysis)
        End If
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionAnalysis Then
          If mvCurrentPage.MenuCount < 2 Then
            SetPage(CareServices.TraderPageType.tpTransactionDetails)
            cmdNext.Enabled = True
          End If
        End If
      End If
      If mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionDetails Then
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpPaymentMethod1.ToString) Then
          If mvTraderPages(CareServices.TraderPageType.tpPaymentMethod1.ToString).MenuCount < 1 Then
            SetPage(CareServices.TraderPageType.tpContactSelection)
          End If
        End If
      End If
      PreProcessPage()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'Remove button on tpMembershipMembersSummary page
    Dim vCursor As New BusyCursor()
    Try
      If mvCurrentRow >= 0 Then
        If mvMembersDGR.GetValue(mvCurrentRow, "ContactNumber").Length > 0 Then
          mvTA.DeleteMember(mvCurrentRow)
          mvCurrentPage.EditPanel.SetValue("CurrentMembers", mvTA.CurrentMembers.ToString)
          mvCurrentPage.EditPanel.SetErrorField("CurrentMembers", "")   'Clear any error
          If mvCurrentRow = mvMembersDGR.RowCount Then
            'Just removed last Row so there is now no row highlighted 
            SetMembersGridButtons(0)
          Else
            SetMembersGridButtons(mvMembersDGR.CurrentRow)
          End If
        End If
      End If

      cmdNext.Enabled = MembersSummaryNextEnabled()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub

  Private Sub DeleteTASLine(ByVal pRow As Integer)
    mvCurrentRow = pRow
    Dim vDataRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
    Dim vLineType As String = vDataRow.Item("TraderLineType").ToString
    Dim vDelBTA As Boolean = False
    Dim vMsg As String = ""
    Dim vEventBookingNumber As Integer
    Dim vExamBookingNumber As Integer

    Select Case vLineType
      Case "A"    'Accomodation
        vMsg = QuestionMessages.QmConfirmDeleteAccomodation
      Case "AP"             'CollectionPayment
        If mvTA.EditExistingTransaction Then
          vMsg = QuestionMessages.QmConfirmDeleteCollectionPayment
        Else
          vDelBTA = True    'No need to display a message
        End If
      Case "B"          'LegacyBequestReceipt
        If mvTA.EditExistingTransaction Then
          vMsg = QuestionMessages.QmConfirmDeleteLegacyReceipt
        Else
          vDelBTA = True    'No need to display a message
        End If
      Case "E"    'EventBooking
        vMsg = QuestionMessages.QmConfirmDeleteEventBooking
        vEventBookingNumber = IntegerValue(vDataRow.Item("EventBookingNumber").ToString)
      Case "Q"
        vMsg = QuestionMessages.QmConfirmDeleteExamBooking
        vExamBookingNumber = IntegerValue(vDataRow.Item("ExamBookingId").ToString)

      Case "GD", "GP"       'GiftAidDeclaration, PayrollGivingPledge
        'Do Nothing
      Case "N", "L", "U", "K"    'InvoicePayment, InvoiceAllocation, UnallocatedSalesLedgerCash, SundryCreditNoteInvoiceAllocation
        vMsg = QuestionMessages.QmConfirmDeleteInvoicePayment
      Case "V", "VC"        'ServiceBooking, ServiceBookingCredit
        vMsg = QuestionMessages.QmConfirmDeleteServiceBooking
      Case "VE"             'ServiceBookingEntitlement
        ShowInformationMessage(InformationMessages.ImCannotDeleteServiceBookingEntitlement)
      Case Else
        If vDataRow.Item("StockSale").ToString = "Y" Then
          mvTA.StockSales = True
          mvTA.SetStockTransactionValues(IntegerValue(vDataRow.Item("StockTransactionID").ToString), IntegerValue(vDataRow.Item("Issued").ToString), vDataRow("ProductCode").ToString, vDataRow("WarehouseCode").ToString, IntegerValue(vDataRow("Quantity").ToString))
          vMsg = QuestionMessages.QmConfirmDeleteStockSale
          If mvTA.EditExistingTransaction Then vMsg &= "  " & InformationMessages.ImCannotCancelDelete
          If ShowQuestion(vMsg, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            vDelBTA = mvTA.EditExistingTransaction
            Dim vIncrementStock As Boolean = False
            If mvTA.StockIssued > 0 Then
              vIncrementStock = True
            End If
            If mvTA.EditExistingTransaction Then
              'Update Stock Levels if required
              If (vIncrementStock = True AndAlso mvTA.StockIssued > 0) Then AddStockMovement(vDataRow("ProductCode").ToString, vDataRow("WarehouseCode").ToString, mvTA.StockIssued, True, False)
            Else
              'Delete StockMovements and increment Stock Levels
              DeleteStockMovement(vIncrementStock)
              vDelBTA = True ' BR18006
            End If
          End If
          vMsg = ""
        Else
          vDelBTA = True
        End If
    End Select
    If vMsg.Length > 0 Then
      'Prompt user and delete as required
      If mvTA.EditExistingTransaction Then vMsg &= "  " & InformationMessages.ImCannotCancelDelete
      If ShowQuestion(vMsg, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then vDelBTA = True
    End If
    If vDelBTA Then
      ProcessData(CareServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine, False)
      mvTA.DeleteAnalysisLine(mvCurrentRow, vEventBookingNumber, vExamBookingNumber)
    End If
    SetAnalysisEditable(CInt(IIf(mvTASDGR.RowCount > 0, 0, -1)))
    mvCurrentPage.EditPanel.SetValue("CurrentLineTotal", mvTA.CurrentLineTotal.ToString("N")) 'mvTA.CurrentLineTotal is updated after ProcessData
  End Sub

  Private Function MembersSummaryNextEnabled() As Boolean
    Dim vEPL As EditPanel = Nothing
    Dim vMemberTypeCode As String
    Dim vNext As Boolean

    If mvTA.TransactionType = "MEMB" Then
      vEPL = mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).EditPanel
    ElseIf mvTA.TransactionType = "MEMC" Then
      vEPL = mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel
    End If
    If vEPL IsNot Nothing Then
      Dim vNumberMembers As Integer
      Dim vNumberAssociates As Integer
      vMemberTypeCode = vEPL.GetValue("MembershipType")
      If mvMembersDGR.RowCount > 0 Then
        For vRow As Integer = 0 To mvMembersDGR.RowCount - 1
          If mvMembersDGR.GetValue(vRow, "MembershipType") = vMemberTypeCode Then
            vNumberMembers += 1
          Else
            vNumberAssociates += 1
          End If
        Next
      End If
      If (vNumberMembers = IntegerValue(vEPL.GetValue("NumberOfMembers"))) AndAlso (vNumberAssociates = IntegerValue(vEPL.GetValue("MaxFreeAssociates"))) Then
        'We have the correct number of members & associates so enable the Next button
        vNext = True
      End If
    End If

    Return vNext

  End Function

  Private Sub SetMembersGridButtons(ByVal pRow As Integer)
    Dim vRemove As Boolean = (mvTA.TransactionType = "MEMC" Or pRow > 0)
    Dim vAdd As Boolean = (mvTA.MemberCount > mvTA.CurrentMembers)
    mvAddBtn.Enabled = vAdd
    mvFindBtn.Enabled = vAdd
    mvAmendBtn.Enabled = vRemove
    mvRemoveBtn.Enabled = vRemove
    cmdNext.Enabled = MembersSummaryNextEnabled()
  End Sub

#End Region

#Region " DataSets "

  Private Sub SetCBLines(ByVal pDataSet As DataSet)
    'Set CollectionBox lines
    If pDataSet.Tables.Contains("DataRow") Then
      'Always clear and re-populate the DisplayGrid
      Dim vCBTable As DataTable = pDataSet.Tables("DataRow")
      Dim vBox As DataRow
      If mvTA.CollectionBoxDataSet.Tables.Contains("DataRow") Then mvTA.CollectionBoxDataSet.Tables.Remove("DataRow")
      Dim vNewTable As DataTable = vCBTable.Copy()
      vNewTable.TableName = "DataRow"
      mvTA.CollectionBoxDataSet.Tables.Add(vNewTable)
      vBox = vNewTable.Rows(0)
    End If
  End Sub

  Private Sub SetMemberLines(ByVal pType As CareServices.TraderProcessDataTypes, ByVal pDataSet As DataSet)
    'Called from ProcessData
    If pDataSet.Tables.Contains("MemberLine") Then
      Dim vMemberTable As DataTable = pDataSet.Tables("MemberLine")
      Dim vMember As DataRow
      If Not mvTA.MembersDataSet.Tables.Contains("DataRow") Then
        'First time in so create the table
        Dim vNewTable As DataTable = vMemberTable.Copy()
        vNewTable.TableName = "DataRow"
        mvTA.MembersDataSet.Tables.Add(vNewTable)
        vMember = vNewTable.Rows(0)
      Else
        'Table already exists so either update an existing row or add a new row
        Dim vTable As DataTable = mvTA.MembersDataSet.Tables("DataRow")
        Dim vFound As Boolean
        For Each vExistingRow As DataRow In vTable.Rows
          If vExistingRow("LineNumber").ToString = vMemberTable.Rows(0)("LineNumber").ToString Then
            vFound = True
            vExistingRow.ItemArray = vMemberTable.Rows(0).ItemArray
            Exit For
          End If
        Next
        If vFound = False Then
          vTable.Rows.Add(vMemberTable.Rows(0).ItemArray)
          vMember = vMemberTable.Rows(0)
        End If
      End If

      Dim vRow As DataRow = pDataSet.Tables("Result").Rows(0)
      Dim vEPL As EditPanel = Nothing
      Select Case mvTA.TransactionType
        Case "MEMB"
          vEPL = mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).EditPanel
        Case "MEMC"
          vEPL = mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel
      End Select
      If vEPL IsNot Nothing Then mvTA.MemberCount = IntegerValue(vEPL.GetValue("NumberOfMembers")) + IntegerValue(vEPL.GetValue("MaxFreeAssociates"))
    End If

  End Sub

  Private Sub SetRemovedSchPayments(ByVal pDataSet As DataSet)
    'Called from ProcessData
    If pDataSet.Tables.Contains("RemovedSchPaymentLine") Then
      'Always clear and re-populate the DisplayGrid
      Dim vOPSTable As DataTable = pDataSet.Tables("RemovedSchPaymentLine")
      Dim vOPS As DataRow
      If mvTA.RemovedSchPaymentsDataSet.Tables.Contains("DataRow") Then mvTA.RemovedSchPaymentsDataSet.Tables.Remove("DataRow")
      Dim vNewTable As DataTable = vOPSTable.Copy()
      vNewTable.TableName = "DataRow"
      mvTA.RemovedSchPaymentsDataSet.Tables.Add(vNewTable)
      vOPS = vNewTable.Rows(0)
    End If
  End Sub

  Private Sub SetOPSLines(ByVal pDataSet As DataSet)
    'Called from ProcessData
    If pDataSet.Tables.Contains("OPSLine") Then
      'Always clear and re-populate the DisplayGrid
      Dim vOPSTable As DataTable = pDataSet.Tables("OPSLine")
      Dim vOPS As DataRow
      If mvTA.OPSDataSet.Tables.Contains("DataRow") Then mvTA.OPSDataSet.Tables.Remove("DataRow")
      Dim vNewTable As DataTable = vOPSTable.Copy()
      vNewTable.TableName = "DataRow"
      mvTA.OPSDataSet.Tables.Add(vNewTable)
      vOPS = vNewTable.Rows(0)
    End If
  End Sub
  ''' <summary>
  ''' BR19606 For Transaction History, Analysis followed by Edit or Delete will change the Order Payment Schedule, when Edit or Delete are clicked. This is the original order payment history before the change.
  ''' It is used to restore the Order Payment Schedule if Analysis is Cancelled. Using Samrt Client as a temporary store for ther stateless web services.
  ''' This receives the OPSS from the server and saves it. Inly saves on the first call, as Edit Analysis can be executed within Edit Analysis.
  ''' </summary>
  ''' <param name="pDataSet">A dataset that may or may not contain a datatable called OriginalOPSLine</param>
  ''' <remarks>Called from ProcessData. Datatable is only passed to Smart Client so that it can be passed back to the server.</remarks>
  Private Sub SetOriginalOPSLine(ByVal pDataSet As DataSet)

    If pDataSet.Tables.Contains("OriginalOPSLine") AndAlso mvTA.OriginalOPS Is Nothing Then
      mvTA.OriginalOPS = pDataSet.Tables("OriginalOPSLine").Copy()
    End If
  End Sub

  Private Sub SetOSPLines(ByVal pDataSet As DataSet)
    'Called from ProcessData
    If pDataSet.Tables.Contains("OPSLine") Then
      'Always clear and re-populate the DisplayGrid
      Dim vOPSTable As DataTable = pDataSet.Tables("OPSLine")
      Dim vOPS As DataRow
      If mvTA.OSPDataSet.Tables.Contains("DataRow") Then
        mvTA.OSPDataSet = New DataSet
        mvTA.BuildOSPDataSet()
      End If

      Dim vNewTable As DataTable = vOPSTable.Copy()
      vNewTable.TableName = "DataRow"
      If (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment) Then vNewTable.DefaultView.Sort = "ScheduledPaymentNumber DESC"
      'just to change the checkvalue column to a boolean column: otherwise the checkbox click does not work properly.
      vNewTable.Columns.Add(New DataColumn("CheckValue1", GetType(System.Boolean)))
      For Each vRow As DataRow In vNewTable.Rows
        vRow("CheckValue1") = (BooleanValue(vRow("CheckValue").ToString)).ToString
      Next
      vNewTable.Columns.Remove("CheckValue")
      vNewTable.Columns("CheckValue1").ColumnName = "CheckValue"
      mvTA.OSPDataSet.Tables.Add(vNewTable.DefaultView.ToTable)
      vOPS = vNewTable.DefaultView.ToTable.Rows(0)
    End If
  End Sub

  Private Sub SetPOSLines(ByVal pDataSet As DataSet)
    'Set Purchase Order Summary lines
    If pDataSet.Tables.Contains("PurchaseOrderLine") Then
      Dim vPOSTable As DataTable = pDataSet.Tables("PurchaseOrderLine")
      If mvTA.POSDataSet.Tables.Contains("DataRow") Then
        Dim vTable As DataTable = mvTA.POSDataSet.Tables("DataRow")
        Dim vFound As Boolean
        For Each vNewRow As DataRow In vPOSTable.Rows
          For vIndex As Integer = 0 To vTable.Rows.Count - 1
            vFound = False
            If vNewRow("LineNumber").ToString = vTable.Rows(vIndex)("LineNumber").ToString Then
              'Row already exists, update it
              vTable.Rows(vIndex).ItemArray = vNewRow.ItemArray
              vFound = True
              Exit For
            End If
          Next
          If Not vFound Then
            vTable.Rows.Add(vNewRow.ItemArray)
          End If
        Next
      Else
        Dim vNewTable As DataTable = vPOSTable.Copy()
        vNewTable.TableName = "DataRow"
        mvTA.POSDataSet.Tables.Add(vNewTable)
      End If
      mvTA.SetDataSetLineTotal(mvTA.POSDataSet)
    End If
  End Sub

  Private Sub SetBatchInvoices(ByVal pDataSet As DataSet)
    If pDataSet.Tables.Contains("BatchInvoice") Then
      Dim vBatchInvoicesTable As DataTable = pDataSet.Tables("BatchInvoice")

      If vBatchInvoicesTable IsNot Nothing Then
        If mvTA.BatchInvoicesDataSet.Tables.Contains("DataRow") Then
          mvTA.BatchInvoicesDataSet.Tables.Remove("DataRow")
        End If
        Dim vNewTable As New DataTable
        vNewTable = vBatchInvoicesTable.Copy()
        'just to change the Print column to a boolean column: otherwise the checkbox click does not work properly.
        vNewTable.Columns.Add(New DataColumn("Print1", GetType(System.Boolean)))
        For Each vRow As DataRow In vNewTable.Rows
          vRow("Print1") = (BooleanValue(vRow("Print").ToString)).ToString
        Next
        vNewTable.Columns.Remove("Print")
        vNewTable.Columns("Print1").ColumnName = "Print"
        vNewTable.TableName = "DataRow"
        mvTA.BatchInvoicesDataSet.Tables.Add(vNewTable)
      End If
    End If
  End Sub
  Private Sub SetPISLines(ByVal pDataSet As DataSet)
    'Set Purchase Invoice Summary lines

    If pDataSet.Tables.Contains("PurchaseInvoiceLine") Then
      Dim vPISTable As DataTable = pDataSet.Tables("PurchaseInvoiceLine")
      If vPISTable IsNot Nothing Then
        If mvTA.PISDataSet.Tables.Contains("DataRow") Then
          Dim vTable As DataTable = mvTA.PISDataSet.Tables("DataRow")
          Dim vFound As Boolean
          For Each vNewRow As DataRow In vPISTable.Rows
            For vIndex As Integer = 0 To vTable.Rows.Count - 1
              vFound = False
              If vNewRow("LineNumber").ToString = vTable.Rows(vIndex)("LineNumber").ToString Then
                'Row already exists, update it
                vTable.Rows(vIndex).ItemArray = vNewRow.ItemArray
                vFound = True
                Exit For
              End If
            Next
            If Not vFound Then
              vTable.Rows.Add(vNewRow.ItemArray)
            End If
          Next
        Else
          Dim vNewTable As DataTable = vPISTable.Copy()
          vNewTable.TableName = "DataRow"
          mvTA.PISDataSet.Tables.Add(vNewTable)
        End If
        mvTA.SetDataSetLineTotal(mvTA.PISDataSet)
      End If
    End If
  End Sub

  Private Sub SetPPALines(ByVal pDataSet As DataSet)
    'Set Purchase Order Payment lines
    Dim vSetCheckBoxColumn As Boolean = False
    If pDataSet.Tables.Contains("PPALine") Then
      If mvTA.PPADataSet.Tables.Contains("DataRow") Then mvTA.PPADataSet.Tables.Remove("DataRow")
      Dim vNewTable As DataTable = pDataSet.Tables("PPALine").Copy
      vNewTable.TableName = "DataRow"
      vNewTable.Columns.AddRange(New DataColumn() {New DataColumn("ContactName"), New DataColumn("Finder")})
      mvTA.PPADataSet.Tables.Add(vNewTable)
      vSetCheckBoxColumn = True
    ElseIf mvTA.PurchaseOrderScheduleChanged Then
      If mvTA.PPADataSet.Tables.Contains("DataRow") Then mvTA.PPADataSet.Tables.Remove("DataRow")
      Dim vTable As New DataTable
      vTable.TableName = "DataRow"
      Dim vIndex As Integer
      For vIndex = 0 To mvTA.PPADataSet.Tables(0).Rows.Count - 1
        vTable.Columns.Add(mvTA.PPADataSet.Tables(0).Rows(vIndex)("Name").ToString)
      Next
      Dim vPFRow As DataRow = mvTraderPages(CareServices.TraderPageType.tpPurchaseOrderDetails.ToString).EditPanel.FindTextLookupBox("PaymentFrequency").GetDataRow
      If vPFRow IsNot Nothing AndAlso vPFRow("PaymentFrequency").ToString.Length > 0 Then
        Dim vNextPaymentDate As Date = Date.Parse(mvTA.TransactionDate)
        Dim vFrequencyAmount As Double = FixTwoPlaces(mvTA.PPBalance / mvTA.PONumberOfPayments)

        Dim vRow As DataRow
        For vIndex = 0 To mvTA.PONumberOfPayments - 1
          vRow = vTable.NewRow
          vRow("DueDate") = vNextPaymentDate
          vRow("LatestExpectedDate") = vNextPaymentDate
          If mvTA.POPercentage Then vRow("Percentage") = "0.00" Else vRow("Amount") = vFrequencyAmount
          vRow("AuthorisationRequired") = If(mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments, "N", "Y")
          vRow("PaymentNumber") = vIndex + 1
          vTable.Rows.InsertAt(vRow, vIndex)
          If vPFRow("Period").ToString = "M" Then vNextPaymentDate = vNextPaymentDate.AddMonths(IntegerValue(vPFRow("Interval"))) Else vNextPaymentDate = vNextPaymentDate.AddDays(IntegerValue(vPFRow("Interval")))
        Next
      Else
        Dim vNewRow As DataRow
        For vIndex = 0 To mvTA.PONumberOfPayments - 1
          vNewRow = vTable.NewRow
          vNewRow("PaymentNumber") = vIndex + 1
          vTable.Rows.InsertAt(vNewRow, vIndex)
        Next
      End If
      Dim vContactNumber As String = GetPageValue(CareServices.TraderPageType.tpPurchaseOrderDetails, "PayeeContactNumber")
      Dim vCOntactName As String = mvCurrentPage.EditPanel.FindTextLookupBox("PayeeContactNumber").Description
      Dim vAddress As String = GetPageValue(CareServices.TraderPageType.tpPurchaseOrderDetails, "PayeeAddressNumber")
      For Each vRow As DataRow In vTable.Rows
        vRow("PayeeContactNumber") = vContactNumber
        vRow("ContactName") = vCOntactName
        vRow("PayeeAddressNumber") = vAddress
      Next
      mvTA.PPADataSet.Tables.Add(vTable)
      mvTA.PurchaseOrderScheduleChanged = False
      vSetCheckBoxColumn = True
    End If
    If vSetCheckBoxColumn Then
      Dim vCopy As DataTable = mvTA.PPADataSet.Tables("DataRow").Copy
      vCopy.Columns.Add(New DataColumn("PayByBacs1", GetType(System.Boolean)))
      If vCopy.Columns.Contains("PayByBacs") Then

        For Each vRow As DataRow In vCopy.Rows
          vRow("PayByBacs1") = (BooleanValue(vRow("PayByBacs").ToString)).ToString
        Next
        vCopy.Columns.Remove("PayByBacs")
        vCopy.Columns("PayByBacs1").ColumnName = "PayByBacs"
        mvTA.PPADataSet.Tables.Remove("DataRow")
        mvTA.PPADataSet.Tables.Add(vCopy.DefaultView.ToTable)
      End If
    End If
  End Sub

  Private Sub SetPPDLines(ByVal pDataSet As DataSet)
    Dim vLineNumber As Integer

    'called from processData to clear the code out
    If pDataSet.Tables.Contains("PPDLine") Then
      Dim vPPDTable As DataTable = pDataSet.Tables("PPDLine")
      Dim vPPD As DataRow = Nothing
      If Not mvTA.PPDDataSet.Tables.Contains("DataRow") Then    'First analysis line
        Dim vNewTable As DataTable = vPPDTable.Copy
        vNewTable.TableName = "DataRow"
        mvTA.PPDDataSet.Tables.Add(vNewTable)                   'Add the table in
        vPPD = vNewTable.Rows(0)
        vLineNumber = IntegerValue(vPPD.Item("LineNumber").ToString)

      Else
        Dim vTable As DataTable = mvTA.PPDDataSet.Tables("DataRow")
        Dim vFound As Boolean
        For Each vExistingRow As DataRow In vTable.Rows
          If vExistingRow("LineNumber").ToString = vPPDTable.Rows(0)("LineNumber").ToString Then
            'Row already exists, update it
            vExistingRow.ItemArray = vPPDTable.Rows(0).ItemArray
            vFound = True
            vLineNumber = IntegerValue(vExistingRow.Item("LineNumber").ToString)

            Exit For
          End If
        Next
        If Not vFound Then
          vTable.Rows.Add(vPPDTable.Rows(0).ItemArray)
          vPPD = vPPDTable.Rows(0)
          vLineNumber = IntegerValue(vPPD.Item("LineNumber").ToString)
        End If
      End If
      SetPPDEditable(vLineNumber)
      mvTA.SetPPDLineTotal()
    End If
  End Sub

  Private Sub SetTraderAnalysisLines(ByVal pDataSet As DataSet, ByVal pIsEditAnalysisLine As Boolean)
    If pIsEditAnalysisLine = False AndAlso mvTA.PayPlanPayMethod = True AndAlso pDataSet.Tables.Contains("TraderAnalysisLine") Then
      'remove the existing lines and add the ones returned from the server as
      'the payment plan is being created from an unbalanced Transaction. 
      'TraderAnalysisLines that are PaymentPlanDetails would be removed
      If mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then mvTA.AnalysisDataSet.Tables.Remove("DataRow")
    End If

    'called from processData to clear the code out
    If pDataSet.Tables.Contains("TraderAnalysisLine") Then
      Dim vAnalysis As DataTable = pDataSet.Tables("TraderAnalysisLine")
      Dim vAnalysisLineNumber As Integer
      Dim vAnalysisRow As DataRow = Nothing
      SetDataColumnType(vAnalysis, "LineNumber", GetType(Int32))
      If Not mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then    'First analysis line
        Dim vNewTable As DataTable = vAnalysis.Copy
        vNewTable.TableName = "DataRow"
        mvTA.AnalysisDataSet.Tables.Add(vNewTable)                   'Add the table in
        vAnalysisRow = vNewTable.Rows(0)
      Else
        Dim vTable As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
        For Each vNewRow As DataRow In vAnalysis.Rows
          For Each vExistingRow As DataRow In vTable.Rows
            If vExistingRow("LineNumber").ToString = vNewRow("LineNumber").ToString Then
              'Row already exists, update it
              vExistingRow.ItemArray = vNewRow.ItemArray
              'just to mark that we have already looked at this row
              vNewRow("LineNumber") = "-1"
              If mvTA.TransactionType = "SALE" Then vAnalysisRow = vNewRow
              Exit For
            End If
          Next
        Next

        For Each vNewRow As DataRow In vAnalysis.Rows
          If vNewRow("LineNumber").ToString <> "-1" Then
            If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPostageAndPacking Then
              If vNewRow("PostagePacking").ToString = "Y" Then
                'Deleting all existing postage lines
                For vPostageRow As Integer = 0 To vTable.Rows.Count - 1
                  If vTable.Rows(vPostageRow).Item("PostagePacking").ToString = "Y" Then
                    vTable.Rows(vPostageRow).Delete()
                    Exit For
                  End If
                Next
              End If
            End If
            vTable.Rows.Add(vNewRow.ItemArray)
            If mvTA.TransactionType <> "SALE" Then
              vAnalysisRow = vNewRow 'BR12386: Need to set vAnalysisRow for any new row expect 'SALE'
              vAnalysisLineNumber = CInt(vAnalysisRow("LineNumber")) 'BR21183 We want the new row in mvTA.AnalysisDataSet, this is what will be shown to the user. Row not present in table yet so get its line number      
            End If
          End If
        Next
        vTable.DefaultView.Sort = "LineNumber"
        mvTA.AnalysisDataSet.Tables.Remove("DataRow")
        mvTA.AnalysisDataSet.Tables.Add(vTable.DefaultView.ToTable)
        If mvTA.TransactionType <> "SALE" Then
          'BR21183 We want the new row in mvTA.AnalysisDataSet, select by line number. Essential for Batch lead, foreign currency batch transactions where a payment is not the first analysis line
          vAnalysisRow = mvTA.AnalysisDataSet.Tables("DataRow").Select(String.Format("LineNumber={0}", vAnalysisLineNumber.ToString())).FirstOrDefault()
        End If
      End If

      If mvCurrentPage.PageType = CareServices.TraderPageType.tpOutstandingScheduledPayments AndAlso vAnalysisRow IsNot Nothing Then
        If mvTA.MultiCurrency() And mvTA.TransactionType = "PAYM" Then
          vAnalysisRow("Amount") = mvTA.CalcCurrencyAmount(CDbl(vAnalysisRow("Amount")), False).ToString("0.00")
        End If
      End If
      SetAnalysisEditable(0)
      mvTA.SetLineTotal()
      If vAnalysisRow IsNot Nothing Then
        Select Case mvCurrentPage.PageType
          Case CareServices.TraderPageType.tpAccommodationBooking
            Dim vEventInfo As CareEventInfo = mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("EventNumber").CareEventInfo
            vEventInfo.SetBookingInfo(CInt(vAnalysisRow("RoomBookingNumber")), CInt(vAnalysisRow("Quantity")), mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("ContactNumber").ContactInfo)
            Dim vForm As New frmEventSet(Me, vEventInfo, CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations)
            vForm.ShowDialog(Me)
            ShowInformationMessage(InformationMessages.ImAccommodationBookingReference, vEventInfo.BookingNumber.ToString)
            RefreshCardSet(RefreshTypes.rtEventBooking, IntegerValue(mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("ContactNumber").Text), vAnalysisRow)
          Case CareServices.TraderPageType.tpEventBooking
            Dim vEventInfo As CareEventInfo = mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("EventNumber").CareEventInfo
            vEventInfo.SetBookingInfo(CInt(vAnalysisRow("EventBookingNumber")), CInt(vAnalysisRow("Quantity")), mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("ContactNumber").ContactInfo)
            If vEventInfo.NameAttendees = True Then
              Dim vForm As New frmEventSet(Me, vEventInfo, CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates)
              vForm.ShowDialog(Me)
              'now do the activity or relationship information if the groups are set-up
              If vEventInfo.ActivityGroup.Length > 0 OrElse vEventInfo.RelationshipGroup.Length > 0 Then
                Dim vIndex As Integer = 1
                Dim vEventDelegates As CollectionList(Of EventDelegateInfo) = vForm.GetEventDelegates
                For Each vEventDelegate As EventDelegateInfo In vEventDelegates
                  Dim vContactInfo As New ContactInfo(ContactInfo.ContactTypes.ctContact, EntityGroup.DefaultContactGroupCode)
                  ShowDelegateDataSheet(Me, vContactInfo, "D", mvTA.TransactionSource, vEventInfo.ActivityGroup, vEventInfo.RelationshipGroup,
                                        GetInformationMessage(ControlText.FrmDelegateSupplementaryInfo, vEventDelegate.ContactName,
                                        vIndex.ToString, vEventDelegates.Count.ToString), vEventDelegate, True)
                  vIndex += 1
                Next
              End If
            End If
            ShowInformationMessage(InformationMessages.ImEventBookingReference, vEventInfo.BookingNumber.ToString)
            RefreshCardSet(RefreshTypes.rtEventBooking, IntegerValue(mvCurrentPage.EditPanel.FindPanelControl(Of TextLookupBox)("ContactNumber").Text), vAnalysisRow)
          Case CareServices.TraderPageType.tpExamBooking
            Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") And mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0 Then
              Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(0)
              If vRow.Item("ExamBookingId").ToString.Length > 0 Then
                Dim vList As New ParameterList(True)
                vList.IntegerValue("ExamBookingId") = CInt(vRow.Item("ExamBookingId").ToString)
                Dim vExamUnitTable As DataTable = ExamsDataHelper.SelectExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamBookingUnits, vList)
                For Each vExamBookingUnitRow As DataRow In vExamUnitTable.Rows
                  If vExamBookingUnitRow.Item("ActivityGroup").ToString.Length > 0 Then
                    ShowExamCandidateDataSheet(Me, vContactInfo, "C", mvTA.TransactionSource, vExamBookingUnitRow.Item("ActivityGroup").ToString, vExamBookingUnitRow.Item("ExamUnitDescription").ToString, CInt(vExamBookingUnitRow.Item("ExamBookingUnitId").ToString), True)
                  End If
                Next
              End If
            End If
          Case CareServices.TraderPageType.tpTransactionAnalysisSummary
            If mvTA.TransactionType = "SALE" Then
              mvTA.StockSales = BooleanValue(vAnalysisRow.Item("StockSale").ToString)
              If mvTA.StockSales Then mvTA.SetStockTransactionValues(IntegerValue(vAnalysisRow("StockTransactionID").ToString), IntegerValue(vAnalysisRow("Issued").ToString), vAnalysisRow("ProductCode").ToString, vAnalysisRow("WarehouseCode").ToString, IntegerValue(vAnalysisRow("Quantity").ToString))
            End If
          Case CareNetServices.TraderPageType.tpServiceBooking
            ShowInformationMessage(InformationMessages.ImServiceBookingReference, vAnalysisRow("ServiceBookingNumber").ToString)
        End Select
        If (mvCurrentPage.PageType = CareNetServices.TraderPageType.tpComments OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpTransactionAnalysisSummary) _
          AndAlso mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment AndAlso pDataSet.Tables("Result").Columns.Contains("BatchNumber") _
          AndAlso pDataSet.Tables("Result").Columns.Contains("TransactionNumber") Then
          RefreshCardSet(RefreshTypes.rtEventBooking, mvTA.OriginalPayerContactNumber, vAnalysisRow)
        End If
      End If
    End If
  End Sub

  Private Sub SetCMTPPDetails(ByVal pDataSet As DataSet)
    If mvTA.TransactionType = "MEMC" Then
      If pDataSet.Tables.Contains("OldCMTPPDLine") Then
        Dim vOldPPDDetailsTable As DataTable = pDataSet.Tables("OldCMTPPDLine")
        If vOldPPDDetailsTable IsNot Nothing Then
          If vOldPPDDetailsTable.Columns.Contains("CMTProrateCost") = False Then vOldPPDDetailsTable.Columns.Add("CMTProrateCost")
          If vOldPPDDetailsTable.Columns.Contains("CMTExcessPaymentType") = False Then vOldPPDDetailsTable.Columns.Add("CMTExcessPaymentType")
          If mvTA.CMTOldPPDDataSet.Tables.Contains("DataRow") Then mvTA.CMTOldPPDDataSet.Tables.Remove("DataRow")
          Dim vNewTable As New DataTable
          vNewTable = vOldPPDDetailsTable.Copy()
          vNewTable.TableName = "DataRow"
          mvTA.CMTOldPPDDataSet.Tables.Add(vNewTable)
        End If
      End If

      If pDataSet.Tables.Contains("PPDLine") Then
        Dim vNewPPDDetailsTable As DataTable = pDataSet.Tables("PPDLine")
        If vNewPPDDetailsTable IsNot Nothing Then
          If vNewPPDDetailsTable.Columns.Contains("CMTProrateCost") = False Then vNewPPDDetailsTable.Columns.Add("CMTProrateCost")
          If mvTA.CMTNewPPDDataSet.Tables.Contains("DataRow") Then mvTA.CMTNewPPDDataSet.Tables.Remove("DataRow")
          Dim vNewTable As New DataTable
          vNewTable = vNewPPDDetailsTable.Copy()
          vNewTable.TableName = "DataRow"
          mvTA.CMTNewPPDDataSet.Tables.Add(vNewTable)
        End If
      End If
    End If
  End Sub
#End Region

#Region " Display Grids "

  Private Sub FillInvoices(ByVal pDGR As DisplayGrid, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, Optional ByVal pContactNumber As String = "")
    Dim vList As ParameterList = New ParameterList(True)
    Dim vIndex As Integer = 0

    vList("Company") = pCompany
    vList("SalesLedgerAccount") = pSalesLedgerAccount
    mvOSInvoices = New CollectionList(Of InvoiceInfo)

    If pContactNumber.Length > 0 Then vList("ContactNumber") = pContactNumber
    With pDGR
      Dim vDataSet1 As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactOutstandingInvoices, mvTA.PayerContactNumber, vList)
      If vDataSet1 IsNot Nothing Then
        Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet1)
        If vTable IsNot Nothing Then
          For Each vRow As DataRow In vTable.Rows
            Dim vInvoice1 As InvoiceInfo = New InvoiceInfo(vRow)
            mvOSInvoices.Add(vInvoice1.Key, vInvoice1)
            vIndex = vIndex + 1
          Next
          vTable.Columns.Remove("PayCheck")
          vTable.Columns.Add(New DataColumn("PayCheck", GetType(System.Boolean)))
        End If
      End If
      .Populate(vDataSet1)
      If .RowCount > 0 Then
        .SetCellsEditable()
        .SetCellsReadOnly()
        .SetCheckBoxColumn("PayCheck")
        .SetButtonColumn("PayButton", "Pay ")
        .SetColumnWritable("InvoiceDisputeCode")
      End If
    End With

    'Now get the cash invoices/credit notes
    vIndex = 0
    Dim vSLAccount As String = pSalesLedgerAccount
    If String.IsNullOrEmpty(pContactNumber) = True OrElse mvTA.PayerContactNumber.Equals(IntegerValue(pContactNumber)) = False Then
      Dim vCCList As ParameterList = New ParameterList(True)
      vCCList("Company") = mvTA.CACompany
      vCCList("ContactNumber") = mvTA.PayerContactNumber.ToString
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vCCList)
      If vTable IsNot Nothing Then
        vTable.DefaultView.RowFilter = "Company = '" & mvTA.CACompany & "'"
        vSLAccount = vTable.Rows(0)("SalesLedgerAccount").ToString
      End If
    End If
    mvCashInvoices = New CollectionList(Of InvoiceInfo)
    Dim vInvoice As InvoiceInfo = New InvoiceInfo(Today.ToString, (mvTA.TransactionAmount - mvTA.CurrentLineTotal), "C", mvTA.PayerContactNumber, mvTA.PayerAddressNumber, vSLAccount)
    mvCashInvoices.Add(vInvoice.Key, vInvoice)
    vIndex = vIndex + 1

    If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment AndAlso mvTA.BatchNumber > 0 AndAlso (mvTA.TransactionAmount - mvTA.CurrentLineTotal) > 0 Then
      vList.IntegerValue("BatchNumber") = mvTA.BatchNumber
      vList.IntegerValue("TransactionNumber") = mvTA.TransactionNumber
    End If

    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCashInvoices, mvTA.PayerContactNumber, vList)
    If vDataSet IsNot Nothing Then
      Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          vInvoice = New InvoiceInfo(vRow)
          mvCashInvoices.Add(vInvoice.Key, vInvoice)
          vIndex = vIndex + 1
        Next
      End If
    End If
  End Sub

  Private Function GetLowestFutureLineNumber() As Integer
    Dim vLineNumber As Integer = 0

    If mvPPSDGR IsNot Nothing AndAlso mvTA.TransactionType = "MEMB" Then
      For vRow As Integer = 0 To mvPPSDGR.RowCount - 1
        If mvPPSDGR.GetValue(vRow, "TimeStatus") = "F" Then vLineNumber = vRow + 1
        If vLineNumber > 0 Then Exit For
      Next
    End If

    Return vLineNumber

  End Function

  Private Sub GetCollectionBoxes(ByVal pEPL As EditPanel)
    Dim vCollectionPisNumber As String = pEPL.GetValue("PisNumber")
    Dim vList As New ParameterList(True)
    vList.IntegerValue("CollectionNumber") = IntegerValue(pEPL.GetValue("AppealCollectionNumber"))
    If vCollectionPisNumber.Length > 0 Then vList.IntegerValue("CollectionPisNumber") = IntegerValue(vCollectionPisNumber)
    Dim vDS As DataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionBoxesForPayment, vList)
    Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDS)
    If vDT IsNot Nothing Then
      vDT.Columns.Remove("Pay")
      vDT.Columns.Add(New DataColumn("Pay", GetType(System.Boolean)))
    End If
    SetCBLines(vDS)
    mvCBXDGR.Populate(mvTA.CollectionBoxDataSet)

    If mvCBXDGR.RowCount > 0 Then
      With mvCBXDGR
        .SetCellsEditable()
        .SetCellsReadOnly()
        .SetCheckBoxColumn("Pay")
      End With
    End If
  End Sub

  Private Sub GetCollectionPISNumbers(ByVal pEPL As EditPanel, ByVal pCollectionNumber As String)
    Dim vList As New ParameterList(True)
    vList("CollectionNumber") = pCollectionNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtCollectionPISNumbers, vList))
    Dim vCombo As ComboBox = pEPL.FindComboBox("PisNumber")
    vCombo.DataSource = Nothing
    vCombo.SelectedText = ""
    If vTable IsNot Nothing Then
      If pEPL.FindTextLookupBox("AppealCollectionNumber").GetDataRowItem("CollectionType") <> "U" Then
        'Not an UnMannedCollection so add a blank row
        Dim vDataRow As DataRow = vTable.NewRow
        For vIndex As Integer = 0 To vTable.Columns.Count - 1
          vDataRow.Item(vIndex) = ""
        Next
        vTable.Rows.InsertAt(vDataRow, 0)
      End If
      With vCombo
        .ValueMember = "CollectionPISNumber"
        .DisplayMember = "PISNumber"
        .DataSource = vTable
      End With
    End If
  End Sub

  Private Sub mvDGR_CheckBoxClicked(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String)
    If mvInvoicesDGR.GetCheckBoxValue(pRow, pCol) = pValue Then
      Dim vShowDepositMsg As Boolean
      'key for the invoices collection
      Dim vKey As String = InvoiceInfo.KeyValue(mvInvoicesDGR.GetValue(pRow, "BatchNumber"), mvInvoicesDGR.GetValue(pRow, "TransactionNumber"))
      Dim vCheckPayment As Boolean = mvOSInvoices(vKey).AmountPaid = 0 AndAlso mvOSInvoices(vKey).DepositAmount > 0
      Dim vEPL As EditPanel = mvCurrentPage.EditPanel
      If pCol = 1 Then
        Dim vUnallocated As Double = DoubleValue(vEPL.GetValue("CurrentUnAllocated"))
        If CBool(pValue) Then
          If mvOSInvoices(vKey).HasPayments Then 'don't allow the checkbox to be ticked if the Pay button has already been used
            mvInvoicesDGR.SetValue(pRow, "PayCheck", False)
            ShowInformationMessage(InformationMessages.ImInvoiceManualPayments)
          ElseIf vUnallocated > 0 Then
            mvOSInvoices(vKey).ClearPayments()
            Dim vAmountUsed As Double
            If (mvOSInvoices(vKey).InvoiceAmount - mvOSInvoices(vKey).AmountPaid) >= vUnallocated Then
              'Amount due on the invoice is more than available, so allocate all that is available
              mvOSInvoices(vKey).AddPayment(mvCashInvoices(0), DoubleValue(vEPL.GetValue("CurrentUnAllocated")))
              mvCashInvoices(0).AmountUsed = mvCashInvoices(0).AmountUsed + vUnallocated
            Else
              'Amount due on the invoice is less than available, so allocate all what is due
              vAmountUsed = (mvOSInvoices(vKey).InvoiceAmount - mvOSInvoices(vKey).AmountPaid)
              mvCashInvoices(0).AmountUsed += vAmountUsed
              mvOSInvoices(vKey).AddPayment(mvCashInvoices(0), vAmountUsed)
            End If
            mvInvoicesDGR.SetValue(pRow, "AmountPaid", mvOSInvoices(vKey).AmountPaid.ToString)
            vEPL.SetValue("CurrentUnAllocated", (mvCashInvoices(0).InvoiceAmount - (mvCashInvoices(0).AmountPaid + mvCashInvoices(0).AmountUsed)).ToString("0.00"))
            If vCheckPayment Then vShowDepositMsg = mvOSInvoices(vKey).AmountPaid < Val(mvOSInvoices(vKey).DepositAmount)
          Else
            mvInvoicesDGR.SetValue(pRow, "PayCheck", False)
            ShowInformationMessage(InformationMessages.ImAllCurrPaymentAllocated)
          End If
        Else
          'Unchecking the check-box take the payment away
          mvOSInvoices(vKey).AmountPaid = mvOSInvoices(vKey).AmountPaid - mvOSInvoices(vKey).GetCashPaid
          mvCashInvoices(0).AmountUsed = mvCashInvoices(0).AmountUsed - mvOSInvoices(vKey).GetCashPaid
          mvInvoicesDGR.SetValue(pRow, "AmountPaid", mvOSInvoices(vKey).AmountPaid.ToString("0.00"))
          vEPL.SetValue("CurrentUnAllocated", (mvCashInvoices(0).InvoiceAmount - (mvCashInvoices(0).AmountPaid + mvCashInvoices(0).AmountUsed)).ToString("0.00"))
          mvOSInvoices(vKey).ClearPayments(False)
        End If
      End If
      If vShowDepositMsg Then ShowInformationMessage(InformationMessages.ImAmountLessThanDepositAmount, mvOSInvoices(vKey).AmountPaid.ToString("N"), mvOSInvoices(vKey).DepositAmount.ToString("N"))
    End If

  End Sub

  Private Sub mvInvoicesDGR_ButtonClicked(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCol As Integer)
    Dim vCheckPayment As Boolean
    Dim vAmountPaid As Double
    Dim vKey As String
    Dim vShowDepositMsg As Boolean

    vKey = InvoiceInfo.KeyValue(mvInvoicesDGR.GetValue(pRow, "BatchNumber"), mvInvoicesDGR.GetValue(pRow, "TransactionNumber"))
    vCheckPayment = mvOSInvoices(vKey).AmountPaid = 0 And mvOSInvoices(vKey).DepositAmount > 0
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    If mvInvoicesDGR.GetValue(pRow, "PayCheck") IsNot Nothing AndAlso mvInvoicesDGR.GetValue(pRow, "PayCheck").Length > 0 AndAlso CBool(mvInvoicesDGR.GetValue(pRow, "PayCheck")) Then
      ShowInformationMessage(InformationMessages.ImInvoiceHasAutoPayment)
    Else
      Dim vSLAccount As String = vEPL.GetValue("SalesLedgerAccount")
      Dim vForm As frmInvoicePayment = New frmInvoicePayment(mvOSInvoices(vKey), mvCashInvoices, vSLAccount)
      If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        mvCashInvoices = vForm.CashInvoices
        mvOSInvoices.Remove(vKey)
        mvOSInvoices.Add(vKey, vForm.Invoice)
      End If
      vAmountPaid = mvOSInvoices(vKey).AmountPaid
      mvInvoicesDGR.SetValue(pRow, "AmountPaid", vAmountPaid.ToString)
      vEPL.SetValue("CurrentUnAllocated", (mvCashInvoices(InvoiceInfo.KeyValue("0", "0")).InvoiceAmount - mvCashInvoices(InvoiceInfo.KeyValue("0", "0")).AmountUsed).ToString("0.00"))
      If vCheckPayment Then vShowDepositMsg = vAmountPaid < Val(mvOSInvoices(vKey).DepositAmount)
    End If
    If vShowDepositMsg Then ShowInformationMessage(InformationMessages.ImAmountLessThanDepositAmount, mvOSInvoices(vKey).AmountPaid.ToString("N"), mvOSInvoices(vKey).DepositAmount.ToString("N"))
  End Sub

  Private Sub mvInvoicesDGR_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    Dim vUpdateDC As Boolean
    Dim vDisputeCode As String = ""

    With mvInvoicesDGR
      'if invoice dispute code then validate the dispute code.
      If pCol = .GetColumn("InvoiceDisputeCode") Then
        vDisputeCode = .GetValue(pRow, "InvoiceDisputeCode")
        If vDisputeCode IsNot Nothing AndAlso vDisputeCode.Length > 0 Then
          If mvValidDisputeCodes Is Nothing OrElse Not mvValidDisputeCodes.ContainsKey(vDisputeCode) Then
            Dim vlist As New ParameterList(True)
            vlist("InvoiceDisputeCode") = vDisputeCode
            Dim vResult As ParameterList = DataHelper.GetLookupItem(CareNetServices.XMLLookupDataTypes.xldtInvoiceDisputeCodes, vlist)
            If vResult.ContainsKey("InvoiceDisputeCode") Then
              'valid dispute codes are held as a collectionlist so that we do not need to go to the server all the time
              If mvValidDisputeCodes Is Nothing Then mvValidDisputeCodes = New CollectionList(Of String)
              mvValidDisputeCodes.Add(vDisputeCode, vDisputeCode)
              vUpdateDC = True
            Else
              ShowWarningMessage(InformationMessages.ImInvalidValue)
            End If
          Else
            vUpdateDC = True
          End If
        Else
          vUpdateDC = True
          vDisputeCode = ""
        End If
      End If
    End With
    Dim vKey As String = InvoiceInfo.KeyValue(mvInvoicesDGR.GetValue(pRow, "BatchNumber"), mvInvoicesDGR.GetValue(pRow, "TransactionNumber"))
    If vUpdateDC Then mvOSInvoices(vKey).DisputeCode = vDisputeCode
  End Sub

  Private Sub mvMembersDGR_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer)
    mvCurrentRow = pRow
    SetMembersGridButtons(pRow)
  End Sub

  Private Sub mvOPSDGR_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    'Only column that can be changed is RevisedAmount
    Dim vPPBalance As Double = DoubleValue(mvCurrentPage.EditPanel.GetValue("Balance"))
    Dim vAmountOS As Double
    Dim vCol As Integer
    Dim vExpBalance As Double
    Dim vRow As Integer
    Dim vTotal As Double
    Dim vNewAmountOS As Double
    Dim vRevAmount As Double

    Try
      With mvOPSDGR
        If pCol = .GetColumn("RevisedAmount") Then
          'First update current Row
          'Get total of all AmountOutstanding up to current Row (but not including current Row)
          vCol = .GetColumn("AmountOutstanding")
          For vRow = 0 To pRow - 1
            vTotal += DoubleValue(.GetValue(vRow, vCol))
          Next

          'ExpectedBalance = PPBalance - total of AmountOutstanding
          vExpBalance = FixTwoPlaces(vPPBalance - vTotal)

          'Set AmountOutstanding
          vCol = .GetColumn("AmountOutstanding")
          vAmountOS = DoubleValue(.GetValue(pRow, vCol))
          If pValue.Length > 0 Then
            'We have set the RevisedAmount
            If DoubleValue(pValue) > vExpBalance Then
              'RevisedAmount > ExpectedBalance so restrict to ExpectedBalance
              ShowInformationMessage(InformationMessages.ImRevisedAmountGTPPAmount, vExpBalance.ToString())
              pValue = vExpBalance.ToString()
              .SetValue(vRow, pCol, pValue.ToString())

            End If
            .SetValue(pRow, vCol, pValue.ToString)
          Else
            'No RevisedAmount so reset AmountOutstanding to OrigAmountDue
            .SetValue(pRow, vCol, .GetValue(pRow, "OrigAmountDue"))
          End If
          vNewAmountOS = DoubleValue(.GetValue(pRow, vCol))

          'Set ExpectedBalance on current Row
          vExpBalance = FixTwoPlaces(vExpBalance - vNewAmountOS)
          vCol = .GetColumn("ExpectedBalance")
          .SetValue(pRow, vCol, IIf(vExpBalance < 0, "0", vExpBalance.ToString).ToString)

          vTotal = 0
          'Now update all subsequent rows
          If pRow < (.RowCount - 1) Then
            vCol = .GetColumn("AmountOutstanding")
            For vRow = 0 To pRow
              vTotal += DoubleValue(.GetValue(vRow, vCol))
            Next
            vCol = .GetColumn("RevisedAmount")
            For vRow = (pRow + 1) To (.RowCount - 1)
              If .GetValue(vRow, vCol) IsNot Nothing AndAlso .GetValue(vRow, vCol).Length > 0 Then
                vTotal += DoubleValue(.GetValue(vRow, vCol))
              End If
            Next

            'Set vLineAmount to be the AmountOutstanding for all subsequent rows without a RevisedAmount
            Dim vLineAmount As Double = FixTwoPlaces((vPPBalance - vTotal) / ((.RowCount - 1) - pRow))
            If vLineAmount < 0 Then vLineAmount = 0
            If vLineAmount > 0 AndAlso (FixTwoPlaces(vLineAmount * ((.RowCount - 1) - pRow))) <> FixTwoPlaces(vPPBalance - vTotal) Then vLineAmount += 0.01

            'Now update the other Rows
            For vRow = (pRow + 1) To (.RowCount - 1)
              vCol = .GetColumn("RevisedAmount")
              vRevAmount = DoubleValue(.GetValue(vRow, vCol))
              If .GetValue(vRow, vCol) IsNot Nothing AndAlso .GetValue(vRow, vCol).Length > 0 Then
                'Just update the ExpectedBalance
                vExpBalance = FixTwoPlaces(vExpBalance - vRevAmount)
                .SetValue(vRow, "ExpectedBalance", vExpBalance.ToString)
              Else
                'No RevisedAmount so update AmountOutstanding & ExpectedBalance
                If (vTotal + vLineAmount) > vPPBalance Then
                  'Final line is for more than the amount outstanding
                  vLineAmount = FixTwoPlaces(vPPBalance - vTotal)
                End If
                vCol = .GetColumn("AmountOutstanding")
                vAmountOS = DoubleValue(.GetValue(pRow, vCol))
                If vRow = (.RowCount - 1) Then
                  If vLineAmount > vExpBalance Then vLineAmount = vExpBalance
                End If
                .SetValue(vRow, vCol, vLineAmount.ToString)
                vCol = .GetColumn("ExpectedBalance")
                vExpBalance = FixTwoPlaces(vExpBalance - vLineAmount)
                .SetValue(vRow, vCol, vExpBalance.ToString)
                vTotal += vLineAmount
              End If
            Next
          End If

          'Now total up all the lines and display to user
          vTotal = 0
          vCol = .GetColumn("AmountOutstanding")
          For vRow = 0 To (.RowCount - 1)
            vTotal += DoubleValue(.GetValue(vRow, vCol))
          Next
          mvCurrentPage.EditPanel.SetValue("AmountOutstanding", FixTwoPlaces(vTotal).ToString("0.00"))
        End If
      End With
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try

  End Sub

  Private Sub mvOSPDGR_CheckBoxClicked(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String)
    PayScheduledPayment(pRow, pCol, pValue)
  End Sub

  Private Sub mvPPSDGR_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer)
    SetPPDEditable(IntegerValue(mvPPSDGR.GetValue(pRow, "LineNumber")))
  End Sub

  Private Sub mvTASDGR_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer)
    SetAnalysisEditable(pRow)
  End Sub

  Private Sub mvPPADGR_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer)
    If pRow >= 0 Then SetPPAEditable(pRow)
  End Sub
  Private Sub mvPPADGR_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    If mvPPADGR.GetColumn("PayeeContactNumber") = pCol Then
      Dim vValid As Boolean = True
      vValid = IntegerValue(pValue) > 0
      If vValid Then
        Dim vList As New ParameterList(True)
        vList("ContactNumber") = pValue
        Dim vDS As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftContacts, vList)
        If vDS Is Nothing OrElse vDS.Tables.Contains("DataRow") = False Then vValid = False
      End If
      If vValid Then
        mvPPADGR_ContactSelected(IntegerValue(pValue), pRow)
      Else
        mvPPADGR.SetValue(pRow, pCol, pOldValue)
      End If
    End If
  End Sub
  Private Function ValidateBankDetails(ByVal pPopPaymentType As String) As Boolean
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPopPaymentMethods)
    For Each vRow As DataRow In vDataTable.Rows
      If vRow.Item("PopPaymentMethod").ToString = pPopPaymentType Then
        If vRow.Item("RequiresBankDetails").ToString = "Y" Then
          Return True
        Else
          Return False
        End If
      End If
    Next
    Return False
  End Function
  Private Sub mvPPADGR_ButtonClicked(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCol As Integer)
    Dim vContactNumber As Integer = FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, New ParameterList(True), Me, True, False)
    If vContactNumber > 0 Then
      mvPPADGR.SetValue(pRow, "PayeeContactNumber", vContactNumber.ToString)
      mvPPADGR_ContactSelected(vContactNumber, pRow)
    End If
  End Sub
  Private Sub mvPPADGR_ContactSelected()
    mvPPADGR_ContactSelected(0, -1)
  End Sub
  Private Sub mvPPADGR_ContactSelected(ByVal pContactNumber As Integer, ByVal pRow As Integer)

    Dim vContactInfo As ContactInfo = Nothing
    Dim vAddressColl As New CollectionList(Of DataTable)
    Dim vAddressTable As DataTable = Nothing
    Dim vContactsTable As DataTable = Nothing
    Dim vLastContactNumber As Integer

    If pRow >= 0 Then
      'Using finder or typing in the contact number
      vContactInfo = New ContactInfo(pContactNumber)
      vAddressTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, pContactNumber))
      If vAddressTable IsNot Nothing Then vAddressColl.Add(pContactNumber.ToString, vAddressTable)
    Else
      'Get Address data for each different contact in the grid
      vContactsTable = mvTA.PPADataSet.Tables("DataRow").DefaultView.ToTable
      vContactsTable.DefaultView.Sort = "PayeeContactNumber,PayeeAddressNumber"
      For Each vRow As DataRow In vContactsTable.DefaultView.ToTable.Rows
        If vLastContactNumber <> IntegerValue(vRow("PayeeContactNumber").ToString) Then
          vLastContactNumber = IntegerValue(vRow("PayeeContactNumber").ToString)
          Dim vList As New ParameterList(True)
          vList("SystemColumns") = "N"
          vAddressTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vLastContactNumber, vList))
          If vAddressTable IsNot Nothing Then vAddressColl.Add(vLastContactNumber.ToString, vAddressTable)
        End If
      Next
    End If

    Dim vExisting As Boolean = vAddressColl.Count > 0
    'Set address data
    For Each vTable As DataTable In vAddressColl
      Dim vItemsData() As String = {}
      Dim vItemsValues() As String = {}
      Dim vIndex As Integer = 1
      If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
        For Each vRow As DataRow In vTable.Rows
          If BooleanValue(vRow("Historical").ToString) = False OrElse BooleanValue(vRow("Default").ToString) Then
            Array.Resize(vItemsData, vIndex)
            vItemsData.SetValue(vRow("AddressNumber").ToString, vIndex - 1)
            If pRow >= 0 Then
              mvPPADGR.SetValue(pRow, "ContactName", vContactInfo.ContactName)
              If BooleanValue(vRow("Default").ToString) Then mvPPADGR.SetValue(pRow, "PayeeAddressNumber", vRow("AddressNumber").ToString)
              mvPPADGR.SetValue(pRow, "PayByBacs", False)
            End If
            Array.Resize(vItemsValues, vIndex)
            vItemsValues.SetValue(vRow("AddressLine").ToString, vIndex - 1)
            vIndex += 1
          End If
        Next
      End If
      If pRow >= 0 Then
        mvPPADGR.SetComboBoxCell(pRow, mvPPADGR.GetColumn("PayeeAddressNumber"), vItemsValues, vItemsData)
      ElseIf vExisting Then
        vLastContactNumber = IntegerValue(vAddressColl.FindKey(vTable))
        vContactInfo = New ContactInfo(vLastContactNumber)
        Dim vRowIndex As Integer = 0
        For Each vRow As DataRow In vContactsTable.Rows
          If vLastContactNumber = IntegerValue(vRow("PayeeContactNumber").ToString) Then
            mvPPADGR.SetValue(vRowIndex, "ContactName", vContactInfo.ContactName)
            mvPPADGR.SetComboBoxCell(vRowIndex, mvPPADGR.GetColumn("PayeeAddressNumber"), vItemsValues, vItemsData)
            If mvPPADGR.GetValue(vRowIndex, "AuthorisationStatus").Length > 0 AndAlso BooleanValue(mvPPADGR.GetValue(vRowIndex, "ReadyForPayment")) Then
              mvPPADGR.SetCellsReadOnly(vRowIndex, mvPPADGR.GetColumn("PayeeAddressNumber"), True, True)
            End If
          End If
          vRowIndex += 1
        Next
      Else
        mvPPADGR.SetComboBoxColumn("PayeeAddressNumber", vItemsValues, vItemsData)
      End If
    Next
  End Sub

  Private Sub mvPPADGR_PaymentMethodSelection()

    Dim vContactInfo As ContactInfo = Nothing
    Dim vAddressColl As New CollectionList(Of DataTable)
    Dim vAddressTable As DataTable = Nothing
    Dim vPaymentMethod As DataTable = Nothing
    Dim vItemsData() As String = {}
    Dim vItemsValues() As String = {}

    For vRow As Integer = 0 To mvPPADGR.MaxGridRows - 1
      vPaymentMethod = mvTA.PPADataSet.Tables("DataRow").DefaultView.ToTable

      Dim vDTPayMethod As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPopPaymentMethods)

      Dim vIndex As Integer = 0
      For Each vPayMethodRow As DataRow In vDTPayMethod.Rows
        Array.Resize(vItemsData, vIndex + 1)
        Array.Resize(vItemsValues, vIndex + 1)
        vItemsData.SetValue(vPayMethodRow.Item("PopPaymentMethod"), vIndex)
        vItemsValues.SetValue(vPayMethodRow.Item("PopPaymentMethodDesc"), vIndex)
        mvPPADGR.SetComboBoxCell(vRow, mvPPADGR.GetColumn("PopPaymentMethod"), vItemsValues, vItemsData)
        If vPaymentMethod.Columns.Contains("PopPaymentMethod") AndAlso vPaymentMethod.Rows(vRow).Item("PopPaymentMethod").ToString.Length > 0 Then
          mvPPADGR.SetValue(vRow, "PopPaymentMethod", vPaymentMethod.Rows(vRow).Item("PopPaymentMethod").ToString)
        Else
          mvPPADGR.SetValue(vRow, "PopPaymentMethod", AppValues.ControlValue(AppValues.ControlValues.pop_def_payment_method))
        End If
        vIndex += 1
      Next
    Next

  End Sub

  Private Sub dgrInvoiceGrid_RowDoubleClicked(ByVal sender As Object, ByVal pRow As Integer)
    If mvInvoiceGrid.GetValue(pRow, "EventNumber").ToString.Length > 0 Then
      Dim vEventNumber As Integer = CInt(mvInvoiceGrid.GetValue(pRow, "EventNumber"))
      FormHelper.ShowEventIndex(vEventNumber)
    End If
  End Sub

  Private Sub dgrInvoiceGrid_CheckBoxClicked(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String)
    If BooleanValue(mvInvoiceGrid.GetCheckBoxValue(pRow, pCol)) = BooleanValue(pValue) Then
      Dim vAllSelected As Nullable(Of Boolean)
      If BooleanValue(pValue) = False AndAlso BooleanValue(mvCurrentPage.EditPanel.GetValue("SelectAll")) Then
        vAllSelected = False
      ElseIf BooleanValue(pValue) AndAlso BooleanValue(mvCurrentPage.EditPanel.GetValue("SelectAll")) = False Then
        Dim vFound As Boolean
        For vRow As Integer = 0 To mvInvoiceGrid.RowCount - 1
          If BooleanValue(mvInvoiceGrid.GetValue(vRow, pCol)) = False AndAlso vRow <> pRow Then
            vFound = True
            Exit For
          End If
        Next
        If vFound = False Then vAllSelected = True
      End If
      If vAllSelected.HasValue Then
        mvSuppressEvents = True
        mvCurrentPage.EditPanel.SetValue("SelectAll", CBoolYN(vAllSelected.Value))
        mvSuppressEvents = False
      End If
    End If
  End Sub

  Private Sub mvPPADGR_CheckBoxClicked(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String)
    If BooleanValue(mvPPADGR.GetCheckBoxValue(pRow, pCol)) = BooleanValue(pValue) Then
      If BooleanValue(pValue) Then
        If Not ValidateDefaultContactAccount(pRow) Then
          ShowErrorMessage(InformationMessages.ImPPAPayByBacsNoDefaultBankAccount)
          mvPPADGR.SetValue(pRow, pCol, False)
        End If
      End If
    End If
  End Sub

  Private Sub SetPPAEditable(ByVal pRow As Integer)
    With mvCurrentPage.EditPanel
      If mvPPADGR.GetValue(pRow, "Amount") IsNot Nothing Then
        .SetValue("Amount", mvPPADGR.GetValue(pRow, "Amount"))
        .SetValue("Percentage", mvPPADGR.GetValue(pRow, "Percentage"))
        .SetValue("DueDate", mvPPADGR.GetValue(pRow, "DueDate"))
        .SetValue("LatestExpectedDate", mvPPADGR.GetValue(pRow, "LatestExpectedDate"))
        .FindCheckBox("AuthorisationRequired").Checked = mvPPADGR.GetValue(pRow, "AuthorisationRequired") = "Y"
        If Not mvPPADGR.ActiveColumn = mvPPADGR.GetColumn("PayByBacs") Then
          If mvTA.POPercentage Then .FindTextBox("Percentage").Focus() Else .FindTextBox("Amount").Focus()
        End If
        Dim vEnabled As Boolean = True
        If mvTA.PurchaseOrderNumber > 0 AndAlso ((mvPPADGR.GetValue(pRow, "PostedOn") IsNot Nothing AndAlso mvPPADGR.GetValue(pRow, "PostedOn").Length > 0) _
          OrElse (mvPPADGR.GetValue(pRow, "AuthorisationStatus") IsNot Nothing AndAlso mvPPADGR.GetValue(pRow, "AuthorisationStatus").Length > 0)) Then vEnabled = False
        If vEnabled Then
          .EnableControl("Amount", Not mvTA.POPercentage)
          .EnableControl("Percentage", mvTA.POPercentage)
          .EnableControlList("DueDate,LatestExpectedDate,AuthorisationRequired", vEnabled)
        Else
          .EnableControlList("Amount,Percentage,DueDate,LatestExpectedDate,AuthorisationRequired", vEnabled)
        End If

        If .PanelInfo.PanelItems.Exists("PoPaymentType") Then .EnableControl("PoPaymentType", vEnabled)
        If .PanelInfo.PanelItems.Exists("DistributionCode") Then .EnableControl("DistributionCode", vEnabled)
        If .PanelInfo.PanelItems.Exists("NominalAccount") Then .EnableControl("NominalAccount", vEnabled)
        If .PanelInfo.PanelItems.Exists("SeparatePayment") Then .EnableControl("SeparatePayment", vEnabled)
        If .PanelInfo.PanelItems.Exists("Checkbox") Then .EnableControl("Checkbox", vEnabled)

        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.po_pay_disable_nominal_acc, False) = True AndAlso .PanelInfo.PanelItems.Exists("NominalAccount") Then
          .EnableControl("NominalAccount", False)
        End If

      Else
        .SetValue("Amount", "0.00")
        If mvTA.POPercentage Then .SetValue("Percentage", "0") Else .SetValue("Percentage", "")
        .FindCheckBox("AuthorisationRequired").Checked = True
        .SetValue("DueDate", mvTA.TransactionDate)
        .SetValue("LatestExpectedDate", mvTA.TransactionDate)
        If Not mvPPADGR.ActiveColumn = mvPPADGR.GetColumn("PayByBacs") Then
          If mvTA.POPercentage Then .FindTextBox("Percentage").SelectAll() Else .FindTextBox("Amount").SelectAll()
        End If
      End If
      .FindDateTimePicker("LatestExpectedDate").Checked = True
      If mvPPADGR.GetValue(pRow, "PoPaymentType") IsNot Nothing AndAlso mvPPADGR.GetValue(pRow, "PoPaymentType").Length > 0 Then
        If .PanelInfo.PanelItems.Exists("PoPaymentType") Then .SetValue("PoPaymentType", mvPPADGR.GetValue(pRow, "PoPaymentType"))
      End If
      If mvPPADGR.GetValue(pRow, "DistributionCode") IsNot Nothing AndAlso mvPPADGR.GetValue(pRow, "DistributionCode").Length > 0 Then
        If .PanelInfo.PanelItems.Exists("DistributionCode") Then .SetValue("DistributionCode", mvPPADGR.GetValue(pRow, "DistributionCode"))
      End If
      If mvPPADGR.GetValue(pRow, "NominalAccount") IsNot Nothing AndAlso mvPPADGR.GetValue(pRow, "NominalAccount").Length > 0 Then
        If .PanelInfo.PanelItems.Exists("NominalCode") Then .SetValue("NominalAccount", mvPPADGR.GetValue(pRow, "NominalAccount"))
      End If
      If mvPPADGR.GetValue(pRow, "SeparatePayment") IsNot Nothing AndAlso .PanelInfo.PanelItems.Exists("SepratePayment") Then .FindCheckBox("SeparatePayment").Checked = mvPPADGR.GetValue(pRow, "SeparatePayment") = "Y"

      If pRow <> mvPPADGR.MaxGridRows - 1 Then
        cmdNext.Enabled = True
        cmdFinished.Enabled = False
      End If

    End With

  End Sub
  Private Sub PayScheduledPayment(ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String)
    Dim vAmount As Double
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    With mvOSPDGR
      'Still have an amount outstanding
      If mvOSPDGR.GetCheckBoxValue(pRow, pCol) = pValue Then
        If CBool(pValue) Then   'if the checkbox has been checked
          If DoubleValue(vEPL.GetValue("AmountOutstanding")) > 0 Then
            If (.GetValue(pRow, "ScheduledPaymentStatus") = "V" Or .GetValue(pRow, "ScheduleCreationReason") = "AP") Then
              'Pay the full amount if possible
              vAmount = DoubleValue(vEPL.GetValue("AmountOutstanding"))
              If vAmount > DoubleValue(.GetValue(pRow, "AmountOutstanding")) Then vAmount = DoubleValue(.GetValue(pRow, "AmountOutstanding"))
            Else
              vAmount = DoubleValue(.GetValue(pRow, "AmountOutstanding"))
            End If
            If vAmount > DoubleValue(vEPL.GetValue("AmountOutstanding")) Then
              vAmount = DoubleValue(vEPL.GetValue("AmountOutstanding"))
            End If
            'Reduce the amount oustanding on this line
            .SetValue(pRow, "AmountOutstanding", (DoubleValue(.GetValue(pRow, "AmountOutstanding")) - vAmount).ToString("0.00"))
            'Update the payment amount to show the amount being paid
            .SetValue(pRow, "PaymentAmount", vAmount.ToString("0.00"))
            vEPL.SetValue("AmountOutstanding", (DoubleValue(vEPL.GetValue("AmountOutstanding")) - vAmount).ToString("0.00"))
          Else
            'There is nothing outstanding so uncheck the control
            .SetValue(pRow, "CheckValue", False)
          End If
        Else    'vbUnchecked
          vAmount = DoubleValue(.GetValue(pRow, "PaymentAmount"))
          If vAmount > 0 Then
            'Increase the amount outstanding on this line
            .SetValue(pRow, "AmountOutstanding", (DoubleValue(.GetValue(pRow, "AmountOutstanding")) + vAmount).ToString("0.00"))
            'Update the payment amount to show nothing being paid
            .SetValue(pRow, "PaymentAmount", "0.00")
            vEPL.SetValue("AmountOutstanding", (DoubleValue(vEPL.GetValue("AmountOutstanding")) + vAmount).ToString("0.00"))
          End If
        End If
        cmdNext.Enabled = (DoubleValue(vEPL.GetValue("AmountOutstanding")) = 0)
        If DoubleValue(vEPL.GetValue("AmountOutstanding")) > 0 Then
          'Enable cmdNext if all payments allocated so that remainder can be allocated
          Dim vEnable As Boolean = True
          If Not CBool(pValue) Then
            vEnable = False
          Else
            For vRow As Integer = 0 To mvOSPDGR.RowCount - 1

              If vRow <> pRow AndAlso CBool(pValue) AndAlso Not CBool(mvOSPDGR.GetValue(vRow, "CheckValue")) Then
                vEnable = False 'An unallocated payment
                Exit For
              End If
            Next
          End If
          cmdNext.Enabled = vEnable
        End If
      End If
    End With
  End Sub

  Private Sub SetAnalysisEditable(ByVal pRow As Integer)
    Dim vDataRow As DataRow = Nothing
    Dim vLineType As String = ""
    Dim vDelete As Boolean
    Dim vEdit As Boolean

    If mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionAnalysisSummary Then
      If mvTASDGR.RowCount > 0 Then
        If pRow >= 0 Then
          vDataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(pRow)
          vLineType = vDataRow.Item("TraderLineType").ToString
        End If

        Select Case vLineType
          Case "A", "L", "N", "R", "U", "V", "VC", "I", "K"
            'Accomodation,InvoiceAllocation,InvoicePayment,SundryCreditNote,UnallocatedSalesLedgerCash,ServiceBooking,ServiceBookingCredit,Exam,SundryCreditNoteInvoiceAllocation
            vEdit = False
            vDelete = True
          Case "Q"
            vEdit = False
            vDelete = mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.None
          Case "E"
            'Event
            vEdit = (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment)
            vDelete = True
          Case "B"
            'LegacyBequestReceipt,Incentive
            vEdit = False
            vDelete = True
          Case "P"
            'ProductSale
            vEdit = Not (vDataRow.Item("PostagePacking").ToString = "Y")
            vDelete = True
            If mvTA.BatchLedApp = True AndAlso mvTA.EditExistingTransaction Then
              If mvTA.BatchInfo.Picked <> "N" Then
                If vDataRow.Item("StockSale").ToString = "Y" Then
                  vEdit = False
                  vDelete = False
                End If
              End If
            ElseIf (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment) Then
              'BR11000: Disable DELETE button for stock sale in re-analysis
              'Do not allow deletion of an existing stock sale line
              If BooleanValue(vDataRow.Item("StockSale").ToString) Then
                vDelete = False
              End If
            End If
          Case "CC", "DD", "SO", "CCU", "DDU", "SOU"
            'CreditCardAuthority,DirectDebit,StandingOrder,'CreditCardAuthorityUpdate,DirectDebitUpdate,StandingOrderUpdate
            'Currently unsupported - once supported these need to be vDelete = True
            vEdit = False
            vDelete = True
          Case "X"
            'EventPricingMatrixLine
            vEdit = False
            vDelete = False
          Case "AS", "AA"
            vEdit = False
            vDelete = True
          Case "O", "M", "C"
            'PaymentPlan,Membership,Covenant
            If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.CashBatchConfirmation Then
              'BR15547: Prevent editing and deleting order payment lines when confirming a provisional cash transaction
              vEdit = False
              vDelete = False
            Else
              vEdit = (vLineType.Length > 0)
              vDelete = (vLineType.Length > 0)
            End If
          Case Else
            vEdit = (vLineType.Length > 0)
            vDelete = (vLineType.Length > 0)
        End Select
        cmdEdit.Enabled = vEdit
        If vDelete Then
          If mvTA.BatchInfo IsNot Nothing Then
            If ((mvTA.BatchLedApp AndAlso mvTA.BatchInfo.PostedToCashBook) AndAlso mvTA.BatchInfo.BatchType <> CareServices.BatchTypes.CreditSales AndAlso
                 mvTA.BatchInfo.BatchType <> CareServices.BatchTypes.StandingOrder AndAlso mvTA.BatchInfo.BatchType <> CareServices.BatchTypes.BankStatement AndAlso
                 AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cb_delete_transactions) = False) _
                 OrElse vLineType = "E" AndAlso mvTA.FinancialAdjustment <> BatchInfo.AdjustmentTypes.None Then
              vDelete = False
            End If
          End If
        End If
        cmdDelete.Enabled = vDelete
      Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
      End If
      If (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment) AndAlso DataChanged() Then cmdFinished.Enabled = True
    ElseIf mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPurchaseOrderSummary AndAlso CanAmendPurchaseOrderAmount = False Then
      cmdDelete.Enabled = False
      cmdNext.Enabled = False
    End If
  End Sub

  Private Sub SetPPDEditable(ByVal pLineNumber As Integer)
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanSummary Then
      If pLineNumber >= 0 AndAlso mvTA.PPDDataSet IsNot Nothing AndAlso mvPPSDGR.RowCount > 0 Then
        Dim vEdit As Boolean = True
        Dim vDelete As Boolean = True
        If (pLineNumber = 1 OrElse pLineNumber = GetLowestFutureLineNumber()) AndAlso (mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "CMEM" OrElse mvTA.TransactionType = "MEMC") Then
          'This is the current or future Membership charging line - this can not be deleted
          vDelete = False
        ElseIf ((mvTA.ApplicationType = ApplicationTypes.atTransaction AndAlso mvTA.TransactionType = "LOAN") OrElse ((mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance) And mvTA.PaymentPlan IsNot Nothing AndAlso mvTA.PaymentPlan.PlanType = PaymentPlanInfo.ppType.pptLoan)) Then
          Dim vAccruesInterest As Boolean = BooleanValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "AccruesInterest"))
          Dim vLoanInterest As Boolean = BooleanValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "LoanInterest"))
          If (pLineNumber = 1 OrElse (pLineNumber = 0 AndAlso IntegerValue(mvPPSDGR.GetValue(mvPPSDGR.CurrentRow, "LineNumber")) > 0)) AndAlso vAccruesInterest = True Then vDelete = False 'Loan capital product
          If vLoanInterest Then 'Loan Interest product
            vEdit = False
            vDelete = False
          End If
        End If
        cmdEdit.Enabled = vEdit
        cmdDelete.Enabled = vDelete
      Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
      End If
    End If
  End Sub

#End Region

#Region " External Events "

  Private Sub EPL_GetInitialCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByRef pList As ParameterList)
    Select Case pParameterName
      Case "ContactGroup"
        Dim vEpl As EditPanel = DirectCast(sender, CDBNETCL.EditPanel)
        If vEpl.PanelInfo.PanelItems(pParameterName).TableName = "service_controls" Then
          pList = New ParameterList(True)
          pList("ServiceGroup") = "Y"
        End If
      Case "PoPaymentType"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("NonHistoricPopType") = "Y"
      Case "ExamCentreCode", "ExamUnitCode"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("Trader") = "Y"
      Case "ExamSessionCode"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("NonSessionBased") = "Y"
        pList("Trader") = "Y"
    End Select
  End Sub

  Private Sub EPL_GetCodeRestrictions(ByVal pSender As Object, ByVal pParameterName As String, ByVal pList As ParameterList)
    Select Case pParameterName
      Case "AppealCollectionNumber"
        If mvTA.BatchLedApp = True AndAlso mvTA.BatchNumber > 0 Then
          pList("BankAccount") = mvTA.BatchInfo.BankAccount
        End If
      Case "AffiliatedMemberNumber"
        If mvTA.MembershipNumber > 0 Then pList("MembershipNumber") = mvTA.MembershipNumber.ToString
        'clearing the value so that this cannot be used again if the user changes the member number, the membership number should get reset in the epl_membership_selected for the new member 
        mvTA.MembershipNumber = 0
      Case "BequestNumber"
        pList("LegacyNumber") = DirectCast(pSender, EditPanel).GetValue("LegacyNumber")
      Case "BlockBookingNumber"
        pList("EventNumber") = DirectCast(pSender, EditPanel).GetValue("EventNumber")
        pList("RoomType") = DirectCast(pSender, EditPanel).GetValue("RoomType")
      Case "BookingNumber", "EventBookingNumber"
        pList("ContactNumber") = DirectCast(pSender, EditPanel).GetValue("ContactNumber")
      Case "InterestBookingNumber"
        pList("ContactNumber") = DirectCast(pSender, EditPanel).GetValue("ContactNumber")
        pList("BookingStatus") = "I,T"
      Case "CommunicationNumber"
        pList("ContactNumber") = DirectCast(pSender, EditPanel).GetValue("ContactNumber")
      Case "EventNumber"
        pList("EventGroup") = GetPageEventGroup(DirectCast(pSender, EditPanel))
        If mvCurrentPage.PageType <> CareServices.TraderPageType.tpAccommodationBooking Then
          pList("Booking") = "Y"
        End If
      Case "MemberNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection AndAlso mvTA.CMTMemberNumber.Length > 0 Then
          pList("CancellationReason") = ""
        End If
        If mvTA.MembershipNumber > 0 Then
          pList("MembershipNumber") = mvTA.MembershipNumber.ToString
          'clearing the value so that this cannot be used again if the user changes the member number, the membership number should get reset in the epl_membership_selected for the new member 
          mvTA.MembershipNumber = 0
        End If
        If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
          pList("CreditListReconciliation") = "Y"
        End If
      Case "OptionNumber"
        pList("EventNumber") = DirectCast(pSender, EditPanel).GetValue("EventNumber")
      Case "PaymentPlanNumber"
        If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
          pList("CreditListReconciliation") = "Y"
        End If
      Case "Product"
        Select Case mvTA.TransactionType
          Case "ACOM"
            pList("FindProductType") = "A"        'Accommodation
          Case "COLP", "DONR", "GAYE"
            pList("FindProductType") = "F"        'all flags bar donation set to 'N'
          Case "CRDN"
            pList("FindProductType") = "N"        'Non Membership
          Case "DONS"
            pList("FindProductType") = "O"        'donation or sponsorship event
          Case "EVNT"
            pList("FindProductType") = "E"        'Event
          Case "MEMB", "MEMC"
            'Member Creation and CMT
            If (mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance) Then
              If mvCurrentRow >= 0 AndAlso (mvCurrentRow = 0 OrElse DirectCast(IntegerValue(mvPPSDGR.GetValue(mvCurrentRow, "PPDLineType")), CareServices.PaymentPlanDetailTypes) = CareServices.PaymentPlanDetailTypes.ppdltCharge) Then
                pList("FindProductType") = "M"        'Membership
              Else
                pList("FindProductType") = "N"        'Non-Membership
              End If
            Else
              pList("FindProductType") = "M"        'Membership
            End If
            If (mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance AndAlso mvTA.TransactionType = "MEMB") AndAlso String.Compare(DataHelper.GetClientCode, "RHS", True) = 0 AndAlso (mvCurrentRow = 0 OrElse mvCurrentRow = 1) Then
              pList("FindProductType") = "M"
            End If
          Case "SALE"
            If mvTA.GiftInKind OrElse mvTA.SaleOrReturn OrElse mvTA.Voucher OrElse mvTA.CAFCard Then pList("FindProductType") = "Z" Else pList("FindProductType") = "P" 'Product (no Donations)
          Case "SUBS"
            pList("FindProductType") = "C"        'Subscription
            'not specifying FindProductType flag means products with no flags set 
        End Select
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderProducts Then pList("FindProductType") = "S"
        If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPostageAndPacking Then pList("FindProductType") = "G"
        If mvTA.SalesGroup.Length > 0 Then pList("SalesGroup") = mvTA.SalesGroup
        If mvTA.TransactionSource.Length > 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then pList("ProductSource") = mvTA.TransactionSource
      Case "ProductNumber"
        pList("ContactNumber") = mvCurrentPage.EditPanel.GetValue("ContactNumber")
        pList("BatchType") = AppValues.GetBatchTypeCode(CareServices.BatchTypes.SaleOrReturn)
        pList("FindTransactionType") = "V"
      Case "Rate"
        Select Case mvCurrentPage.PageType
          Case CareServices.TraderPageType.tpEventBooking
            pList("EventNumber") = DirectCast(pSender, EditPanel).GetValue("EventNumber")
            pList("BookingDate") = AppValues.TodaysDate()
          Case CareNetServices.TraderPageType.tpGiveAsYouEarnEntry
            pList("CurrentPrice") = "0"
          Case CareServices.TraderPageType.tpPaymentPlanProducts, CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts Then pList("CurrencyCode") = mvTA.DefaultCurrencyCode
            pList("LoanInterest") = "N"
        End Select
        If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking Then
          pList("ContactNumber") = DirectCast(pSender, EditPanel).GetValue("BookingContactNumber")
        Else
          Dim vContactNumber As String = DirectCast(pSender, EditPanel).GetOptionalValue("ContactNumber")
          If vContactNumber.Length > 0 Then pList("ContactNumber") = vContactNumber
        End If

        If (Not (pList.Contains("CurrencyCode"))) Then
          If mvTA.BatchInfo IsNot Nothing AndAlso mvTA.BatchInfo.CurrencyCode IsNot Nothing Then
            pList("CurrencyCode") = mvTA.BatchInfo.CurrencyCode
          ElseIf mvTA.BatchCurrencyCode IsNot Nothing AndAlso mvTA.BatchCurrencyCode.Length > 0 Then
            pList("CurrencyCode") = mvTA.BatchCurrencyCode
          End If
        End If
      Case "SalesLedgerAccount"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpInvoicePayments Then
          If mvTA.PayerContactNumber > 0 Then
            'Do not include ContactNumber
            pList("Company") = mvTA.CACompany
          End If
        Else
          Dim vContactNumber As String = mvCurrentPage.EditPanel.GetValue("ContactNumber")
          If vContactNumber.Length > 0 Then
            pList("ContactNumber") = vContactNumber
            pList("Company") = mvTA.CSCompany
          End If
        End If
      Case "Appeal"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderDetails Then
          If mvCurrentPage.EditPanel.FindTextLookupBox("Campaign", False) IsNot Nothing AndAlso mvCurrentPage.EditPanel.GetValue("Campaign").ToString.Length > 0 Then
            pList("Campaign") = mvCurrentPage.EditPanel.GetValue("Campaign").ToString
          End If
        End If
      Case "Segment"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderDetails Then
          If mvCurrentPage.EditPanel.FindTextLookupBox("Campaign", False) IsNot Nothing AndAlso mvCurrentPage.EditPanel.FindTextLookupBox("Appeal", False) IsNot Nothing AndAlso mvCurrentPage.EditPanel.GetValue("Campaign").ToString.Length > 0 AndAlso mvCurrentPage.EditPanel.GetValue("Appeal").ToString.Length > 0 Then
            pList("Campaign") = mvCurrentPage.EditPanel.GetValue("Campaign").ToString
            pList("Appeal") = mvCurrentPage.EditPanel.GetValue("Appeal").ToString
          End If
        End If
      Case "RelatedContactNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking Then
          GetServiceModifiers(True, True)
        End If
      Case "ServiceContactNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking Then
          Dim vEpl As EditPanel = CType(pSender, EditPanel)
          Dim vList As New ParameterList(True)
          Dim vServiceControl As DataRow = Nothing
          vList("ContactGroup") = vEpl.GetValue("ContactGroup")
          If vList("ContactGroup").Length > 0 Then
            vServiceControl = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtServiceControl, vList)
          End If

          If vServiceControl IsNot Nothing Then
            'Set up the finder type and defaults 
            pList("FinderType") = vServiceControl("FinderType").ToString
            pList("GeographicalRegionType") = vServiceControl("GeographicalRegiontype").ToString
            pList("ContactGroup") = vServiceControl("ContactGroup").ToString
            pList("LockContactGroup") = "Y" 'Restrict selection to the current service
            If vEpl.FindPanelControl("StartDate", False) IsNot Nothing Then
              pList("Date") = vEpl.GetValue("StartDate")
            End If
            If vEpl.FindPanelControl("EndDate", False) IsNot Nothing Then
              pList("Date2") = vEpl.GetValue("EndDate")
            End If
          End If
        End If
      Case "ServiceBookingNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails Then
          If mvTraderPages(CareNetServices.TraderPageType.tpTransactionDetails.ToString).EditPanel.FindTextLookupBox("ContactNumber", False) IsNot Nothing AndAlso GetPageValue(CareNetServices.TraderPageType.tpTransactionDetails, "ContactNumber").Length > 0 Then
            pList("ContactNumber") = GetPageValue(CareNetServices.TraderPageType.tpTransactionDetails, "ContactNumber")
          End If
        End If
      Case "ContactNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership Then
          Dim vEpl As EditPanel = CType(pSender, EditPanel)
          If vEpl.FindPanelControl("Joined", False) IsNot Nothing AndAlso vEpl.GetValue("Joined").Length > 0 Then
            pList("Joined") = vEpl.GetValue("Joined")
            vEpl.FindTextLookupBox("MembershipType").ComboBox.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, pList)
          End If
        End If
      Case "ExamUnitCode"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpExamBooking Then
          pList("Trader") = "Y"
        End If
      Case "ExamSessionCode"
        pList("NonSessionBased") = "Y"
    End Select
  End Sub

  Private Sub EPL_ShowMessage(ByVal pSender As Object, ByVal pMessage As String)
    sbp.Text = pMessage
  End Sub

  Public Sub HandleContactCreated(ByVal pContactInfo As ContactInfo)
    With pContactInfo
      If mvNewContacts Is Nothing Then mvNewContacts = New CollectionList(Of ContactInfo)
      mvNewContacts.Add(.ContactNumber.ToString, pContactInfo)
      UserHistory.AddContactHistoryNode(.ContactNumber, .ContactName, .ContactGroup)
      'Default source code with source from new contact
      Dim vEPL As EditPanel = mvCurrentPage.EditPanel
      If vEPL IsNot Nothing AndAlso .ContactCreated = True AndAlso FindControl(vEPL, "Source", False) IsNot Nothing Then
        If vEPL.GetValue("Source").Length = 0 Then SetValueRaiseChanged(vEPL, "Source", .Source)
      End If
    End With
  End Sub

  Private Sub EPL_CheckedItemsChanged(ByVal sender As Object)
    CalculateExamBookingPrice(mvCurrentPage.EditPanel)
  End Sub

  Private Sub EPL_MembershipSelected(ByVal pSender As Object, ByVal pParameterName As String, ByVal pMembershipNumber As Integer)
    If pParameterName = "AffiliatedMemberNumber" Then
      mvTA.MembershipNumber = pMembershipNumber
    Else
      mvTA.MembershipNumber = pMembershipNumber
    End If
    If FindControl(mvCurrentPage.EditPanel, "MemberNumber", False) IsNot Nothing Then mvCurrentPage.EditPanel.PanelInfo.PanelItems("MemberNumber").LastValue = "" 'to force a value changed event
    If FindControl(mvCurrentPage.EditPanel, "AffiliatedMemberNumber", False) IsNot Nothing Then mvCurrentPage.EditPanel.PanelInfo.PanelItems("AffiliatedMemberNumber").LastValue = "" 'to force a value changed event
  End Sub
  Private Sub EPL_ProductNumberSelected(ByVal pSender As Object, ByVal pList As ParameterList, ByVal pProductNumber As Integer, ByVal pContactNumber As Integer)
    Dim vEPL As EditPanel = DirectCast(pSender, EditPanel)
    Dim vValid As Boolean = True
    With vEPL
      If pList IsNot Nothing Then
        .SetValue("ProvisionalBatchNumber", pList("ProvisionalBatchNumber"))
        .SetValue("ProvisionalTransNumber", pList("ProvisionalTransNumber"))
        .SetValue("Amount", pList("Amount"))
        .SetValue("TransactionDate", pList("TransactionDate"))
        .SetValue("Reference", pList("Reference"))
      ElseIf pProductNumber > 0 OrElse pContactNumber > 0 Then
        Dim vList As New ParameterList(True)
        If pProductNumber > 0 Then vList.IntegerValue("ProductNumber") = pProductNumber Else vList.IntegerValue("ContactNumber") = pContactNumber
        Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProductNumbers, vList)
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
          Dim vRow As DataRow = vDataTable.Rows(0)
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpConfirmProvisionalTransactions Then
            vRow("ContactNumber") = .GetValue("ContactNumber")
            If pContactNumber > 0 Then
              If vDataTable.Rows.Count = 1 Then .Populate(vRow) Else vEPL.FindTextLookupBox("ProductNumber").OpenFinder()
            Else
              .Populate(vRow)
            End If

          ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionDetails Then
            .SetValue("ContactNumber", vRow("ContactNumber").ToString, False, True)
          End If
        Else
          vValid = False
        End If
      Else
        If mvCurrentPage.PageType <> CareServices.TraderPageType.tpTransactionDetails AndAlso vEPL.GetValue("ProvisionalBatchNumber").Length = 0 Then vValid = False
      End If
      If vValid = False Then
        .ClearControlList("ProvisionalBatchNumber,ProvisionalTransNumber,Amount,TransactionDate,Reference")
        .SetErrorField("ProductNumber", InformationMessages.ImInvalidProductNumber, True)
      Else
        .SetErrorField("ProductNumber", "")
      End If
    End With
  End Sub

  Private Sub CardAuthorisationComplete(sender As Object, e As EventArgs)
    If Me.CardAuthoriser.IsAuthorised AndAlso
       Not Me.CardAuthoriser.IsCancelled Then
      ProcessFinish()
    ElseIf Me.CardAuthoriser.IsCancelled Then
      Me.Close()
      If mvCancelSagepay Then InitCardtAuthorisation()
      mvCancelSagepay = False
    Else
      InitCardtAuthorisation()
    End If
  End Sub

#End Region

#Region " Edit Panel Interaction "

  Private Function GetPageEventGroup(ByVal pEPL As EditPanel) As String
    If FindControl(pEPL, "EventGroup", False) IsNot Nothing Then
      Return pEPL.GetValue("EventGroup")
    Else
      Return AppValues.DefaultEventGroupCode
    End If
  End Function

  Private Function GetPageValue(ByVal pTraderPageType As CareServices.TraderPageType, ByVal pParameterName As String) As String
    Dim vPageValue As String = ""
    If mvTraderPages.ContainsKey(pTraderPageType.ToString) Then
      Dim vPage As TraderPage = mvTraderPages(pTraderPageType.ToString)
      If vPage IsNot Nothing Then
        vPageValue = vPage.EditPanel.GetValue(pParameterName)
      End If
    End If
    Return vPageValue
  End Function

  Private Function GetOptionalPageValue(ByVal pTraderPageType As CareServices.TraderPageType, ByVal pParameterName As String) As String
    Dim vPageValue As String = ""
    If mvTraderPages.ContainsKey(pTraderPageType.ToString) Then
      Dim vPage As TraderPage = mvTraderPages(pTraderPageType.ToString)
      If vPage IsNot Nothing Then
        vPageValue = vPage.EditPanel.GetOptionalValue(pParameterName)
      End If
    End If
    Return vPageValue
  End Function

  Private Sub SetPageValue(ByVal pTraderPageType As CareServices.TraderPageType, ByVal pParameterName As String, ByVal pValue As String, Optional ByVal pCheckDefaultsSet As Boolean = False)
    If mvTraderPages.ContainsKey(pTraderPageType.ToString) Then
      Dim vPage As TraderPage = mvTraderPages(pTraderPageType.ToString)
      If vPage IsNot Nothing Then
        If (pCheckDefaultsSet AndAlso vPage.DefaultsSet) Or Not pCheckDefaultsSet Then
          vPage.EditPanel.SetValue(pParameterName, pValue)
        End If
      End If
    End If
  End Sub

  Private Sub SetValueRaiseChanged(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, Optional ByVal pDisable As Boolean = False)
    pEPL.SetValue(pParameterName, pValue, pDisable)
    EPL_ValueChanged(pEPL, pParameterName, pValue)
  End Sub

  Private Sub SetValueRaiseChanged(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, ByVal pDisable As Boolean, ByVal pForceEvent As Boolean, ByVal pReportError As Boolean, ByVal pUpdateLastValue As Boolean)
    pEPL.SetValue(pParameterName, pValue, pDisable, pForceEvent, pReportError, pUpdateLastValue)
    EPL_ValueChanged(pEPL, pParameterName, pValue)
  End Sub

  Private Function DataChanged() As Boolean
    For Each vTraderPage As TraderPage In mvTraderPages
      If vTraderPage.EditPanel.DataChanged Then Return True
    Next
    Return False
  End Function

#End Region

#Region " Page Initialization "

  Private Sub ClearPageDefaults(ByVal pPageType As CareServices.TraderPageType)
    For Each vPage As TraderPage In mvTraderPages
      If vPage.PageType = pPageType Then vPage.DefaultsSet = False
    Next
  End Sub

  Private Sub SetButtons(ByVal pRow As DataRow)
    'Handle the setting of the various buttons based on the return values
    cmdNext.Enabled = pRow("NextButton").ToString = "Enabled"
    If pRow("PreviousButton").ToString = "CheckPM1" Then
      cmdPrevious.Enabled = mvTraderPages(CareServices.TraderPageType.tpPaymentMethod1.ToString).MenuCount > 1
    Else
      cmdPrevious.Enabled = pRow("PreviousButton").ToString = "Enabled"
    End If
    cmdFinished.Enabled = pRow("FinishedButton").ToString = "Enabled"

    cmdEdit.Visible = pRow("EditButton").ToString = "Visible"
    cmdDelete.Visible = pRow("DeleteButton").ToString = "Visible"
    bpl.RepositionButtons()

    If Me.AcceptButton Is cmdNext And cmdNext.Enabled = False Then
      'Adjust Accept Button (i.e. the button that Enter will mimic) 
      If cmdFinished.Enabled Then
        Me.AcceptButton = cmdFinished
      End If
    End If

  End Sub

  Private Sub SetDefaults(ByVal pRow As DataRow)
    With mvCurrentPage
      .EditPanel.Populate(pRow)
      If FindControl(.EditPanel, "Source", False) IsNot Nothing AndAlso pRow.Table.Columns.Contains("Source") Then
        SetValueRaiseChanged(.EditPanel, "Source", pRow("Source").ToString)
      End If
      If FindControl(.EditPanel, "Product", False) IsNot Nothing AndAlso pRow.Table.Columns.Contains("Product") Then
        SetValueRaiseChanged(.EditPanel, "Product", pRow("Product").ToString)
      End If
      Select Case .PageType
        Case CareServices.TraderPageType.tpAccommodationBooking
          If FindControl(.EditPanel, "EventGroup", False) IsNot Nothing Then .EditPanel.SetValue("EventGroup", AppValues.DefaultEventGroupCode)
        Case CareServices.TraderPageType.tpAmendMembership
          'Restrict MembershipType selection to either the Membership or Associate Membership Type
          Dim vRow As DataRow = Nothing
          If mvTA.TransactionType = "MEMB" Then
            vRow = mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).EditPanel.FindTextLookupBox("MembershipType").GetDataRow
          ElseIf mvTA.TransactionType = "MEMC" Then
            vRow = mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel.FindTextLookupBox("MembershipType").GetDataRow
          End If
          If vRow IsNot Nothing Then
            Dim vRestriction As String = "MembershipType {0} '" & vRow.Item("MembershipType").ToString & "'"
            If vRow.Item("AssociateMembershipType").ToString.Length > 0 Then
              vRestriction = String.Format(vRestriction, "IN (")
              vRestriction &= ",'" & vRow.Item("AssociateMembershipType").ToString & "')"
            Else
              vRestriction = String.Format(vRestriction, "=")
            End If
            'Before applying the restriction on MembershipType, store MembershipType code
            'and re-set it on the page as adding the filter may have changed the value
            Dim vMembType As String = .EditPanel.GetValue("MembershipType")
            .EditPanel.FindTextLookupBox("MembershipType").SetFilter(vRestriction, True)
            SetValueRaiseChanged(.EditPanel, "MembershipType", vMembType)
          End If

        Case CareServices.TraderPageType.tpBankDetails
          .EditPanel.PanelInfo.PanelItems("AccountName").Mandatory = True
          If .EditPanel.GetValue("SortCode").Length > 0 Then
            If .EditPanel.PanelInfo.PanelItems.Exists("AccountNumber") Then
              .EditPanel.FindPanelControl("AccountNumber").Focus()
            Else
              SetBankDetails(.EditPanel, "SortCode", .EditPanel.GetValue("SortCode"), mvTA.AlbacsBankDetails)
            End If
          End If
          If mvTA.ApplicationType <> ApplicationTypes.atCreditListReconciliation Then .EditPanel.SetValue("Reference", mvTA.TransactionReference)

        Case CareNetServices.TraderPageType.tpBatchInvoiceSummary
          If FindControl(.EditPanel, "PrintPreview", False) IsNot Nothing Then
            If mvTA.InvoiceDocument.Length = 0 Then
              'If there is no InvoiceDocument then we cannot preview invoices
              .EditPanel.SetValue("PrintPreview", "N", True)
            End If
          End If
        Case CareServices.TraderPageType.tpCardDetails
          If WebBasedCardAuthoriser.IsAvailable AndAlso mvTA.OnlineCCAuthorisation Then
            InitCardtAuthorisation()
          Else
            .EditPanel.EnableControlList("Issuer,IssueNumber,ValidDate", False)
            .EditPanel.EnableControl("SecurityCode", mvTA.OnlineCCAuthorisation)
            .EditPanel.PanelInfo.PanelItems("SecurityCode").Mandatory = False
            EPL_ValueChanged(.EditPanel, "CreditOrDebitCard", .EditPanel.GetValue("CreditOrDebitCard"))
            If mvTA.CAFCard Then
              .EditPanel.EnableControl("CreditOrDebitCard", False)
              'CAF Card Expiry Date no longer required - default Expiry Date for CAF Cards only to be current month + 50 years
              If mvTA.TransactionPaymentMethod = "CAFC" Then
                .EditPanel.SetValue("ExpiryDate", Today.AddYears(50).ToString("MMyy"))
              End If
            End If
          End If


        Case CareServices.TraderPageType.tpChangeMembershipType
          If mvTA.ContactVATCategory.Length = 0 Then
            Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
            If vContactInfo IsNot Nothing Then mvTA.ContactVATCategory = vContactInfo.VATCategory
          End If
          If BooleanValue(.EditPanel.GetValue("GiftMembership")) = False Then
            .EditPanel.EnableControlList("OneYearGift,PackToDonor,GiftCardStatus_N,GiftCardStatus_B,GiftCardStatus_W,GiverContactNumber,AffiliatedMemberNumber", False)
            .EditPanel.PanelInfo.PanelItems("AffiliatedMemberNumber").Mandatory = False
          End If
          If Not (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.enter_member_number) = True And AppValues.ConfigurationValue(AppValues.ConfigurationValues.member_number_format) = "char_seq_integer") Then
            .EditPanel.EnableControl("MemberNumber", False)
          End If
          'BR12426: add any restriction to membership types
          'changes added for membership type categories restrictions to pass through parameters to restrict for membership type categories if no transition records
          mvTA.CMTOriginalMemberJoined = pRow.Item("OriginalJoined").ToString
          Dim vList As New ParameterList(True, True)
          vList("ContactNumber") = mvTA.CMTMemberContactNumber.ToString
          vList("Joined") = .EditPanel.GetValue("Joined").ToString
          Dim vDT As DataTable = DataHelper.MembershipTypeTransitionsTableRestricted(mvTA.PaymentPlan.PayPlanMembershipTypeCode, vList)
          If (vDT Is Nothing OrElse vDT.Rows.Count = 0) Then
            'no transitions have been found - check if there are transitions for this membership type without limiting through membership categories
            Dim vTable1 As DataTable = DataHelper.MembershipTypeTransitionsTable(mvTA.PaymentPlan.PayPlanMembershipTypeCode)
            If vTable1.Rows.Count = 0 Then
              'none have been found so need to change the datasource
              vDT = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList, True)
            End If
          End If
          .EditPanel.FindPanelControl(Of TextLookupBox)("MembershipType").ComboBox.DataSource = vDT
          If pRow.Table.Columns.Contains("CMTEarliestDate") Then
            If FindControl(.EditPanel, "CMTDate", False) IsNot Nothing Then
              Dim vDTP As DateTimePicker = .EditPanel.FindPanelControl(Of DateTimePicker)("CMTDate")
              If IsDate(pRow("CMTEarliestDate").ToString) Then vDTP.MinDate = CDate(pRow("CMTEarliestDate").ToString)
              If IsDate(pRow("CMTLatestDate").ToString) Then vDTP.MaxDate = CDate(pRow("CMTLatestDate").ToString)
            End If
          End If
          If pRow.Table.Columns.Contains("DisableWriteOffMissedPayments") Then
            If BooleanValue(pRow.Item("DisableWriteOffMissedPayments").ToString) Then .EditPanel.EnableControl("WriteOffMissedPayments", False)
          End If

        Case CareServices.TraderPageType.tpCollectionPayments
          If .EditPanel.GetValue("AppealCollectionNumber").Length = 0 Then
            'Clear PisNumber Combo (otherwise entering subsequent payment still has previous values displayed)
            Dim vCombo As ComboBox = .EditPanel.FindComboBox("PisNumber")
            vCombo.DataSource = Nothing
            vCombo.SelectedText = ""
          End If

        Case CareServices.TraderPageType.tpContactSelection
          If mvTA.ChangeMembershipType = True AndAlso mvTA.CMTMemberNumber.Length > 0 Then
            .EditPanel.SetValue("MemberNumber", mvTA.CMTMemberNumber)
            EPL_ValidateItem(.EditPanel, "MemberNumber", mvTA.CMTMemberNumber, True)
            EPL_ValueChanged(.EditPanel, "MemberNumber", mvTA.CMTMemberNumber)
          End If
          If mvTA.ApplicationType = ApplicationTypes.atMaintenance Or mvTA.ApplicationType = ApplicationTypes.atConversion Then
            .EditPanel.PanelInfo.PanelItems("ContactNumber").Mandatory = False
            .EditPanel.PanelInfo.PanelItems("AddressNumber").Mandatory = False
            If (mvMembConv OrElse mvTA.ApplicationStartPoint = TraderApplication.TraderApplicationStartPoint.taspRightMouse) AndAlso mvTA.PaymentPlan IsNot Nothing Then
              'Conversion w/o PP details or Conversion with PP details but not adding DD/SO/CCCA
              'Need to empty the grdPPS grid ready for an auto payment method to be added
              'or come in from the popup menu
              mvMembConv = False
              SetValueRaiseChanged(.EditPanel, "PaymentPlanNumber", mvTA.PaymentPlan.PaymentPlanNumber.ToString)
              EPL_ValidateItem(.EditPanel, "PaymentPlanNumber", mvTA.PaymentPlan.PaymentPlanNumber.ToString, True)
              ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
            End If
          End If
          If mvTA.PayerContactNumber > 0 Then SetValueRaiseChanged(.EditPanel, "ContactNumber", mvTA.PayerContactNumber.ToString)

        Case CareServices.TraderPageType.tpCreditCustomer
          .EditPanel.EnableControlList("CreditCategory,StopCode,CreditLimit,CustomerType", AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciNewCreditCustomer))
          If pRow.Table.Columns.Contains("ContactNumber") Then .EditPanel.PanelInfo.PanelItems("ContactNumber").LastValue = "" 'to force a value changed event

        Case CareServices.TraderPageType.tpCreditCardAuthority, CareServices.TraderPageType.tpDirectDebit
          .EditPanel.EnableControl("ClaimDay", AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D")
          If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
            Dim vType As String = IIf(.PageType = CareServices.TraderPageType.tpDirectDebit, "DD", "CC").ToString
            Dim vDT As DataTable = DirectCast(DirectCast(mvCurrentPage.EditPanel.FindPanelControl("ClaimDay"), ComboBox).DataSource, DataTable)
            If vDT IsNot Nothing Then
              DefaultClaimDay(vDT, mvCurrentPage.EditPanel.GetValue("BankAccount"), vType)
              vDT.DefaultView.RowFilter = "BankAccount = '" & mvCurrentPage.EditPanel.GetValue("BankAccount") & "' AND ClaimType = '" & vType & "'"
            Else
              .EditPanel.SetErrorField("ClaimDay", InformationMessages.ImNoClaimDays)
            End If
            If FindControl(.EditPanel, "ClaimDay", False) IsNot Nothing Then
              Dim vClaimDay As ComboBox = .EditPanel.FindComboBox("ClaimDay")
              If mvTA.PaymentPlan Is Nothing OrElse mvTA.PaymentPlan.Existing = False OrElse Not pRow.Table.Columns.Contains("ClaimDay") Then
                If vClaimDay.Items.Count > 0 Then
                  vClaimDay.SelectedIndex = 0
                Else 'if there were items in the data table but none of them were for the correct type, then it will not have caught this scenario above so set error field correctly
                  .EditPanel.SetErrorField("ClaimDay", InformationMessages.ImNoClaimDays)
                End If
              Else
                vClaimDay.SelectedValue = pRow("ClaimDay")
              End If
            End If
            .EditPanel.PanelInfo.PanelItems("ClaimDay").DisableClear = True
          Else
            .EditPanel.SetValue("ClaimDay", "")
          End If
          If .PageType = CareServices.TraderPageType.tpDirectDebit Then
            If .EditPanel.GetValue("SortCode").Length > 0 Then SetBankDetails(.EditPanel, "SortCode", .EditPanel.GetValue("SortCode"))
            .EditPanel.PanelInfo.PanelItems("AccountName").Mandatory = True
            If FindControl(.EditPanel, "DateSigned", False) IsNot Nothing Then
              'Ideally this would be done server-side when we get the controls but need to ensure that existing data with null DateSigned does not automatically get set to Today without the user setting it
              If .EditPanel.PanelInfo.PanelItems("DateSigned").Visible AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_dd_signed_date_mandatory, False) Then .EditPanel.PanelInfo.PanelItems("DateSigned").Mandatory = True
              .EditPanel.FindDateTimePicker("DateSigned").MaxDate = Today
            End If
          End If

          .EditPanel.EnableControlList("ContactNumber,AddressNumber,Amount", mvTA.TransactionType <> "APAY")
          If .PageType = CareNetServices.TraderPageType.tpCreditCardAuthority Then
            'BR13986: Setting the following Mandatory as these are mandatory in Credit Card Maintenance
            .EditPanel.PanelInfo.PanelItems("CreditCardType").Mandatory = True
            .EditPanel.PanelInfo.PanelItems("CreditCardNumber").Mandatory = True
          End If

        Case CareServices.TraderPageType.tpEventBooking
          If FindControl(.EditPanel, "EventGroup", False) IsNot Nothing Then .EditPanel.SetValue("EventGroup", AppValues.DefaultEventGroupCode)
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ev_retain_trader_booking_dets) Then
            If FindControl(.EditPanel, "EventNumber", False) IsNot Nothing Then .EditPanel.SetValue("EventNumber", mvTA.EventNumber)
            Dim vRate As String = mvTA.EventBookingRate 'Store as could get reset by setting Option Number
            If FindControl(.EditPanel, "OptionNumber", False) IsNot Nothing Then SetValueRaiseChanged(.EditPanel, "OptionNumber", mvTA.BookingOptionNumber)
            If FindControl(.EditPanel, "Rate", False) IsNot Nothing Then SetValueRaiseChanged(.EditPanel, "Rate", vRate)
          End If

        Case CareNetServices.TraderPageType.tpExamBooking
          Dim vControl As Control = FindControl(.EditPanel, "ExamUnitId", False)
          If vControl IsNot Nothing Then
            DirectCast(vControl, ExamSelector).InitForTrader(ExamSelector.SelectionType.Courses, String.Empty)
          End If
          SetValueRaiseChanged(.EditPanel, "ExamSessionCode", pRow("ExamSessionCode").ToString)
          SetValueRaiseChanged(.EditPanel, "ExamUnitCode", pRow("ExamUnitCode").ToString)

        Case CareNetServices.TraderPageType.tpLoans
          .EditPanel.FindTextLookupBox("PaymentFrequency").SetFilter("Frequency = '12' And Interval = '1'", True)   'Only allow monthly installments
          If mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance Then
            .EditPanel.EnableControl("OrderDate", False)
            If .EditPanel.GetValue("LoanTerm").Length = 0 Then .EditPanel.EnableControl("LoanTerm", False)
            If .EditPanel.GetValue("FixedMonthlyAmount").Length = 0 Then .EditPanel.EnableControl("FixedMonthlyAmount", False)
            If pRow.Table.Columns.Contains("TransactionType") Then mvTA.TransactionType = pRow("TransactionType").ToString
            mvTA.PPBalance = DoubleValue(pRow("Balance").ToString)
            mvTA.LoanAmount = DoubleValue(pRow("LoanAmount").ToString)
          End If

        Case CareServices.TraderPageType.tpMembership
          .EditPanel.SetValue("GiftCardStatus_N", "N")
          .EditPanel.EnableControlList("OneYearGift,PackToDonor,GiftCardStatus_N,GiftCardStatus_B,GiftCardStatus_W,GiverContactNumber,AffiliatedMemberNumber", False)
          If pRow.Table.Columns.Contains("Branch") Then
            Dim vBranch As String = pRow.Item("Branch").ToString
            .EditPanel.SetValue("Branch", "")
            .EditPanel.SetValue("Branch", vBranch)
          End If
          If Not (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.enter_member_number) = True And AppValues.ConfigurationValue(AppValues.ConfigurationValues.member_number_format) = "char_seq_integer") Then
            .EditPanel.EnableControl("MemberNumber", False)
          End If
          If .EditPanel.GetValue("ContactNumber").Length > 0 Then
            Dim vValid As Boolean = True
            EPL_ValidateItem(.EditPanel, "ContactNumber", .EditPanel.GetValue("ContactNumber"), vValid)
            If vValid = True AndAlso .EditPanel.GetValue("Joined").Length > 0 Then
              Dim vList As New ParameterList(True, True)
              vList("ContactNumber") = .EditPanel.GetValue("ContactNumber").ToString
              vList("Joined") = .EditPanel.GetValue("Joined").ToString
              .EditPanel.FindTextLookupBox("MembershipType").ComboBox.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList, True)
            End If
          End If
          .EditPanel.FindTextLookupBox("MembershipType").SetFilter("AllowAsFirstType <> 'N'", False, True)
          mvSavedMembershipType = String.Empty
        Case CareServices.TraderPageType.tpMembershipPayer
          If BooleanValue(pRow.Item("ForceGiftMembership").ToString) Then
            'Server has determined that GiftMembership flag MUST be set, so set it now
            mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel.SetValue("GiftMembership", "Y")
          End If
          If pRow.Table.Columns.Contains("ContactNumber") AndAlso pRow.Table.Columns.Contains("AddressNumber") Then
            'MembershipPayer page is the CMT page showing the payer of the membership so update TraderApplication class to hold the correct payer details
            mvTA.SetPayerContact(IntegerValue(pRow.Item("ContactNumber").ToString), IntegerValue(pRow.Item("AddressNumber").ToString))
          End If

        Case CareServices.TraderPageType.tpLegacyBequestReceipt
          .EditPanel.EnableControlList("ExpectedValue,EstimatedOutstanding", False)

        Case CareServices.TraderPageType.tpOutstandingScheduledPayments
          If pRow.Table.Columns.Contains("PaymentAmount") Then .EditPanel.SetValue("AmountDue", DoubleValue(pRow("PaymentAmount").ToString).ToString("0.00"))
          If pRow.Table.Columns.Contains("AmountOutstanding") Then .EditPanel.SetValue("AmountOutstanding", DoubleValue(pRow("AmountOutstanding").ToString).ToString("0.00"))
        Case CareServices.TraderPageType.tpPaymentPlanDetails
          .EditPanel.PanelInfo.PanelItems("RenewalAmount").Mandatory = False
          .EditPanel.EnableControl("RenewalAmount", False)
          If .EditPanel.FindPanelControl("UseAsFirstAmount").Visible Then
            'Check whether control should actually be visible (this is the CheckBox control)
            If pRow.Item("UseAsFirstAmount").ToString = "I" Then
              .EditPanel.SetValue("UseAsFirstAmount", "N", True)
              .EditPanel.SetControlVisible("UseAsFirstAmount", False)
            End If
          Else
            .EditPanel.SetValue("UseAsFirstAmount", "N")   'Control is hidden so it should not be checked
          End If
          If .EditPanel.FindTextBox("FirstAmount").Visible Then
            'Check whether control should actually be visible
            .EditPanel.SetControlVisible("FirstAmount", (BooleanValue(pRow.Item("FirstAmountVisible").ToString)))
            If BooleanValue(pRow.Item("FirstAmountVisible").ToString) Then
              If .EditPanel.FindTextLookupBox("PaymentFrequency").GetDataRowItem("Frequency") = "1" Then .EditPanel.SetValue("FirstAmount", "", True)
            End If
          End If
          .EditPanel.PanelInfo.PanelItems("OrderDate").LastValue = .EditPanel.GetValue("OrderDate")   'Stops EPL_ValueChanged being called when tabing out of field and value has not actually changed
          If FindControl(.EditPanel, "OneOffPayment", False) IsNot Nothing Then
            .EditPanel.SetValue("OneOffPayment", "N")
            If mvTA.EditExistingTransaction = False AndAlso (mvTA.TransactionType = "DONR" OrElse mvTA.TransactionType = "SUBS") AndAlso (mvTA.PPPaymentMethod = "DD" OrElse mvTA.PPPaymentMethod = "SO" OrElse mvTA.PPPaymentMethod = "CCCA") Then
              'New RegularDonation / Subscription paid by DD/SO/CCCA
              .EditPanel.EnableControl("OneOffPayment", True)
            Else
              'Everything else
              .EditPanel.EnableControl("OneOffPayment", False)
            End If
          End If
          'Handle StartMonth control
          Dim vStartMonthCombo As ComboBox = TryCast(FindControl(.EditPanel, "StartMonth", False), ComboBox)
          If vStartMonthCombo IsNot Nothing AndAlso vStartMonthCombo.Visible = True Then
            Dim vStartDate As Date = Date.Parse(.EditPanel.GetValue("OrderDate"))
            If pRow.Table.Columns.Contains("FixedStartDate") Then
              If pRow.Item("FixedStartDate").ToString.Length > 0 Then
                vStartDate = Date.Parse(pRow.Item("FixedStartDate").ToString)
                .EditPanel.SetValue("OrderDate", vStartDate.ToString(AppValues.DateFormat))
                EPL_ValueChanged(.EditPanel, "OrderDate", vStartDate.ToString(AppValues.DateFormat))
              End If
            End If
            If IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.payment_plan_minimum_term, "0")) > 0 Then
              'Set ExpiryDate value and minimum value
              Dim vExpiryDate As Date = vStartDate.AddYears(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.payment_plan_minimum_term, "0"))).AddDays(-1)
              .EditPanel.SetValue("ExpiryDate", vExpiryDate.ToString(AppValues.DateFormat))
              Dim vExpiryDateDtp As DateTimePicker = DirectCast(FindControl(.EditPanel, "ExpiryDate"), DateTimePicker)
              vExpiryDateDtp.MinDate = vExpiryDate
            End If
            If .EditPanel.GetValue("RenewalAmount").Length = 0 Then .EditPanel.SetValue("RenewalAmount", "0")
          End If

        Case CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
          If mvTA.ContactVATCategory.Length = 0 Then
            Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
            If vContactInfo IsNot Nothing Then mvTA.ContactVATCategory = vContactInfo.VATCategory
          End If
          If FindControl(.EditPanel, "EffectiveDate", False) IsNot Nothing AndAlso .EditPanel.FindDateTimePicker("EffectiveDate").Checked Then
            .EditPanel.EnableControl("EffectiveDate", True)
            Dim vValue As DateTime = .EditPanel.GetDateTimeValue("EffectiveDate")
            If DataHelper.GetTableFromDataSet(mvTA.PPDDataSet) IsNot Nothing Then
              For Each vRow As DataRow In DataHelper.GetTableFromDataSet(mvTA.PPDDataSet).Rows
                If vRow IsNot Nothing AndAlso vRow("EffectiveDate").ToString.Length > 0 AndAlso vValue < DateTime.Parse(vRow("EffectiveDate").ToString) Then
                  vValue = DateTime.Parse(vRow("EffectiveDate").ToString)
                End If
              Next
            End If
            .EditPanel.SetValue("EffectiveDate", vValue.ToString(AppValues.DateFormat))
          Else
            .EditPanel.EnableControl("EffectiveDate", False)
          End If
          If (mvTA.PaymentPlan.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment) = PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment AndAlso FindControl(.EditPanel, "FirstAmount", False) IsNot Nothing Then
            .EditPanel.EnableControl("FirstAmount", False)
          End If

        Case CareNetServices.TraderPageType.tpPaymentPlanProducts
          'Defaults for Valid From/To
          Dim vControl As Control = .EditPanel.FindPanelControl("ValidFrom", False)
          If vControl IsNot Nothing Then
            If mvTraderPages.ContainsKey(CareNetServices.TraderPageType.tpPaymentPlanDetails.ToString) Then
              If mvTraderPages(CareNetServices.TraderPageType.tpPaymentPlanDetails.ToString).EditPanel.FindPanelControl("OrderDate", False) IsNot Nothing Then
                .EditPanel.SetValue("ValidFrom", GetPageValue(CareNetServices.TraderPageType.tpPaymentPlanDetails, "OrderDate"))
              End If
            End If
          End If
          vControl = .EditPanel.FindPanelControl("ValidTo", False)
          If vControl IsNot Nothing Then
            If mvTraderPages.ContainsKey(CareNetServices.TraderPageType.tpPaymentPlanDetails.ToString) Then
              If mvTraderPages(CareNetServices.TraderPageType.tpPaymentPlanDetails.ToString).EditPanel.FindPanelControl("ExpiryDate", False) IsNot Nothing Then
                .EditPanel.SetValue("ValidTo", GetPageValue(CareNetServices.TraderPageType.tpPaymentPlanDetails, "ExpiryDate"))
              End If
            End If
          End If

        Case CareServices.TraderPageType.tpPaymentPlanMaintenance
          mvTA.TransactionType = pRow.Item("TransactionType").ToString
          .EditPanel.EnableControl("GiverContactNumber", mvTA.PaymentPlan.GiftMembership)
          If (mvTA.PaymentPlan.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment) = PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment AndAlso FindControl(.EditPanel, "FirstAmount", False) IsNot Nothing Then
            .EditPanel.EnableControl("FirstAmount", False)  'Disable as this will be automatically re-calculated when the PaymentPlan is saved
          End If
          .EditPanel.EnableControl("GiftMembership", False)
          If pRow.Table.Columns.Contains("OneYearGiftEnabled") AndAlso pRow.Item("OneYearGiftEnabled").ToString = "N" AndAlso FindControl(.EditPanel, "OneYearGift", False) IsNot Nothing Then
            'Disable the OneYearGift control
            .EditPanel.EnableControl("OneYearGift", False)
          End If

        Case CareServices.TraderPageType.tpProductDetails
          'Handle the situation where the default product and or rate is not valid for the page type
          'e.g The application or batch has a default donation product and we go to the product sale page
          'First ensure Product is enabled (for Stock Sales we may have disabled it)
          If (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment OrElse (mvTA.EditExistingTransaction AndAlso mvTA.ContactVATCategory.Length = 0)) Then
            Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
            If vContactInfo IsNot Nothing Then mvTA.ContactVATCategory = vContactInfo.VATCategory
          End If

          If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment Then
            .EditPanel.PanelInfo.PanelItems("ContactNumber").Mandatory = True
            .EditPanel.PanelInfo.PanelItems("AddressNumber").Mandatory = True
          End If

          .EditPanel.EnableControl("Product", True)
          Dim vProduct As TextLookupBox = .EditPanel.FindTextLookupBox("Product")
          Dim vRate As TextLookupBox = .EditPanel.FindTextLookupBox("Rate")
          Dim vQuantity As TextBox = .EditPanel.FindTextBox("Quantity")
          Dim vAmount As Double = DoubleValue(.EditPanel.GetValue("Amount"))
          If vProduct.Text.Length > 0 AndAlso vProduct.IsValid = False Then
            vProduct.Text = ""
            vRate.FillComboWithRestriction("")
            vQuantity.Text = ControlText.TxtOne
            vAmount = 0
          Else
            Dim vRateCode As String = vRate.Text
            vRate.Focus()     'Change the focus to Rate to force Product to be validated
            If vRateCode.Length > 0 Then vRate.Text = vRateCode 'Validating Product may have cleared the Rate
            If vRate.Text.Length > 0 AndAlso vRate.IsValid = False Then
              vRate.Text = ""
            ElseIf vRate.Text.Length > 0 Then
              EPL_ValueChanged(.EditPanel, "Rate", vRate.Text)
            End If
          End If
          vProduct.Focus()
          If vAmount > 0 And mvTA.LinePrice = 0 Then
            'Setting the Rate will have cleared the Amount for a 0-priced Rate, so reset
            SetValueRaiseChanged(.EditPanel, "Amount", vAmount.ToString("0.00"))
          End If
          'Clear combo
          If vProduct.Text.Length = 0 Then
            Dim vCombo As ComboBox = TryCast(FindControl(.EditPanel, "Warehouse", False), ComboBox)
            If vCombo IsNot Nothing Then
              vCombo.DataSource = Nothing
              vCombo.SelectedValue = ""
              vCombo.Enabled = False
            End If
          End If
          If mvTA.TransactionType = "DONS" Then
            If mvTA.LastDeceasedContactNumber > 0 Then
              .EditPanel.SetValue("DeceasedContactNumber", mvTA.LastDeceasedContactNumber.ToString)
              'Select the first visible checkbox
              If .EditPanel.PanelInfo.PanelItems("LineTypeG").Visible Then
                .EditPanel.SetValue("LineTypeG", "Y")
              ElseIf .EditPanel.PanelInfo.PanelItems("LineTypeH").Visible Then
                .EditPanel.SetValue("LineTypeH", "Y")
              ElseIf .EditPanel.PanelInfo.PanelItems("LineTypeS").Visible Then
                .EditPanel.SetValue("LineTypeS", "Y")
              End If
            Else
              .EditPanel.EnableControl("DeceasedContactNumber", False)
            End If
          End If
          If FindControl(.EditPanel, "CreditedContactNumber", False) IsNot Nothing Then .EditPanel.EnableControl("CreditedContactNumber", False)
          If FindControl(.EditPanel, "ServiceBookingNumber", False) IsNot Nothing AndAlso FindControl(.EditPanel, "ServiceBookingNumber", False).Visible Then
            Dim vSBExisting As Boolean = False
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0 Then
              For Each vRow As DataRow In mvTA.AnalysisDataSet.Tables("DataRow").Rows
                If vRow.Item("TraderLineType").ToString = "V" Then
                  vSBExisting = True
                  Exit For
                End If
              Next
            End If
            .EditPanel.EnableControl("ServiceBookingNumber", mvTA.TransactionType <> "DONS" AndAlso mvTA.ServiceBookingAnalysis AndAlso mvServiceBookingNumber = 0 AndAlso Not vSBExisting)
          End If
          If mvTA.DeliveryContactNumber > 0 AndAlso mvTA.DeliveryAddressNumber > 0 Then
            .EditPanel.SetValue("ContactNumber", mvTA.DeliveryContactNumber.ToString)
            .EditPanel.SetValue("AddressNumber", mvTA.DeliveryAddressNumber.ToString)
          End If

        Case CareServices.TraderPageType.tpStandingOrder
          If .EditPanel.GetValue("SortCode").Length > 0 Then SetBankDetails(.EditPanel, "SortCode", .EditPanel.GetValue("SortCode"))
          .EditPanel.EnableControlList("ContactNumber,AddressNumber,Amount", mvTA.TransactionType <> "APAY")

        Case CareServices.TraderPageType.tpTransactionDetails
          Dim vContactInfo As ContactInfo = .EditPanel.FindTextLookupBox("ContactNumber").ContactInfo
          If vContactInfo IsNot Nothing Then mvTA.ContactVATCategory = vContactInfo.VATCategory
          If .EditPanel.GetValue("Amount").Length > 0 Then EPL_ValueChanged(.EditPanel, "Amount", .EditPanel.GetValue("Amount"))
          'If we have a CreditSale then disable Contact/Address fields to prevent user changing them
          'This is to ensure that the transaction payer is the Credit Customer contact (See also SetPageControls)
          With .EditPanel
            .EnableControl("ContactNumber", (mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN"))
            .EnableControl("AddressNumber", (mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN"))
          End With
          'BR14206: Date of Birth on TRD page not populated when coming from CCU page, forcing ValueChanged event to populate controls (inc. DOB) from Contact Info object
          If mvTA.TransactionPaymentMethod = "CRED" OrElse mvTA.TransactionPaymentMethod = "CQIN" OrElse mvTA.TransactionPaymentMethod = "CCIN" Then EPL_ValueChanged(.EditPanel, "ContactNumber", .EditPanel.GetValue("ContactNumber"))
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_default_gift_aid_elig) = True And mvTA.AutoSetAmount = True And mvTA.TransactionAmount >= mvTA.GiftAidMinimum And Not (mvTA.Voucher Or mvTA.CAFCard Or mvTA.GiftInKind Or mvTA.SaleOrReturn) Then
            .EditPanel.SetValue("EligibleForGiftAid", "Y", False)
          Else
            .EditPanel.SetValue("EligibleForGiftAid", "N", False)
          End If
          .EditPanel.EnableControl("EligibleForGiftAid", Not (mvTA.Voucher OrElse mvTA.CAFCard OrElse mvTA.GiftInKind OrElse mvTA.SaleOrReturn))
          If pRow.Table.Columns.Contains("Campaign") Then
            Dim vValue As String = pRow("Campaign").ToString
            If pRow.Table.Columns.Contains("Appeal") Then vValue &= pRow("Appeal").ToString
            If vValue.Length > 0 Then .EditPanel.SetValue("Source", vValue)
          End If
          If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
            .EditPanel.SetValue("Amount", mvTA.TransactionAmount.ToString("0.00"))
            .EditPanel.SetValue("TransactionDate", mvTA.TransactionDate)
          End If
          'BR20526
          If mvCurrentPage.PageCode = "TRD" AndAlso mvTA.Voucher AndAlso mvTA.TransactionPaymentMethod = "VOUC" Then
            'Check the Additional_Reference_1 + 2 Fields are visable, otherwise trader won't work.
            If Not DirectCast(FindControl(Me, "AdditionalReference1").Tag, PanelItem).Visible Or Not DirectCast(FindControl(Me, "AdditionalReference2").Tag, PanelItem).Visible Then
              ShowInformationMessage(InformationMessages.ImAddAdditionalReferenceFields)
            End If
          End If
        Case CareServices.TraderPageType.tpPurchaseOrderDetails
          With .EditPanel
            If pRow.Table.Columns.Contains("DistributionMethod") Then
              .FindRadioButton("DistributionMethod_" & pRow("DistributionMethod").ToString).Checked = True
            Else
              .FindRadioButton("DistributionMethod_S").Checked = True
            End If
            If pRow.Table.Columns.Contains("PaymentAsPercentage") Then .FindCheckBox("PaymentAsPercentage").Checked = BooleanValue(pRow("PaymentAsPercentage").ToString)
            mvTA.SetPurchaseOrderType(.FindTextLookupBox("PurchaseOrderType").GetDataRow)
            mvTA.PONumberOfPayments = IntegerValue(.FindTextBox("NumberOfPayments").Text)
            mvTA.OriginalPayerContactNumber = IntegerValue(.FindTextLookupBox("ContactNumber").Text)
            If .PanelInfo.PanelItems.Exists("CurrencyCode") Then
              .PanelInfo.PanelItems("CurrencyCode").Mandatory = .PanelInfo.PanelItems("CurrencyCode").Visible
              If mvTA.DefaultCurrencyCode.Length > 0 Then .SetValue("CurrencyCode", mvTA.DefaultCurrencyCode) 'BR19514
            End If
            If mvTA.PPADataSet.Tables.Contains("DataRow") Then
              Dim vDisabled As Boolean
              For Each vRow As DataRow In mvTA.PPADataSet.Tables("DataRow").Rows
                If (vRow("AuthorisationStatus").ToString.Length > 0 AndAlso BooleanValue(vRow("ReadyForPayment").ToString)) OrElse vRow("PostedOn").ToString.Length > 0 Then
                  .EnableControlList("Amount,StartDate,PurchaseOrderType,NumberOfPayments,PayeeContactNumber,PayeeAddressNumber,CurrencyCode", False)
                  If mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments Then
                    .EnableControlList("ContactNumber,PaymentAsPercentage", False)
                  ElseIf mvTA.PurchaseOrderType = PurchaseOrderTypes.AdHocPayments Then
                    .EnableControl("Amount", True)
                  End If
                  vDisabled = True
                  Exit For
                End If
              Next
              If Not vDisabled AndAlso mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments Then
                'RegularPayments but none of the payments has been authorised yet
                .EnableControlList("NumberOfPayments,PaymentAsPercentage", False)
                .FindTextLookupBox("PaymentFrequency").SetFilter("Frequency = '1'", True)
                .SetValue("PaymentFrequency", pRow("PaymentFrequency").ToString)
              End If
            ElseIf mvTA.PurchaseOrderNumber > 0 Then
              .EnableControl("NumberOfPayments", False)
            End If
            If DoubleValue(.GetValue("Amount")) > 0 Then .SetValue("Amount", DoubleValue(.GetValue("Amount")).ToString("0.00"))
            'Only set the User details for POD. It can be added for all of the trader applicaiton
            'but this will need expensive tesing and this can not be done at this point in time. 
            'Something to be added to Dev Led enhancement
            .SetUserDefaults()
          End With

        Case CareServices.TraderPageType.tpPurchaseInvoiceDetails
          .EditPanel.FindDateTimePicker("PurchaseInvoiceDate").Checked = True
          If .EditPanel.PanelInfo.PanelItems.Exists("CurrencyCode") Then
            .EditPanel.PanelInfo.PanelItems("CurrencyCode").Mandatory = .EditPanel.PanelInfo.PanelItems("CurrencyCode").Visible
          End If
        Case CareServices.TraderPageType.tpPurchaseOrderProducts, CareServices.TraderPageType.tpPurchaseInvoiceProducts
          .EditPanel.SetValue("Quantity", "1")
          mvTA.EditLineNumber = 0
          Dim vLineItemName As String = "PILineItem"
          If .PageType = CareServices.TraderPageType.tpPurchaseOrderProducts Then
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_retain_po_dist_code) Then .EditPanel.SetValue("DistributionCode", AppValues.LastDistributionCode)
            .EditPanel.SetValue("NominalAccount", AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_purchase_order_line_account, ""))
            If .EditPanel.FindComboBox("Warehouse") IsNot Nothing Then .EditPanel.FindComboBox("Warehouse").DataSource = Nothing
            vLineItemName = "POLineItem"
          End If
          SetPOILineItemControls(.EditPanel, vLineItemName, "")
        Case CareServices.TraderPageType.tpPurchaseOrderPayments
          .EditPanel.EnableControl("Amount", Not mvTA.POPercentage)
          .EditPanel.EnableControl("Percentage", mvTA.POPercentage)
          If AppValues.ControlValue(AppValues.ControlTables.purchase_order_controls, AppValues.ControlValues.po_payment_type).Length > 0 AndAlso
             .EditPanel.PanelInfo.PanelItems.Exists("PoPaymentType") Then
            .EditPanel.FindTextLookupBox("PoPaymentType").Text = AppValues.ControlValue(AppValues.ControlTables.purchase_order_controls, AppValues.ControlValues.po_payment_type)
          End If

          If AppValues.ControlValue(AppValues.ControlTables.purchase_order_controls, AppValues.ControlValues.distribution_code).Length > 0 AndAlso
            .EditPanel.PanelInfo.PanelItems.Exists("DistributionCode") Then
            .EditPanel.FindTextLookupBox("DistributionCode").Text = AppValues.ControlValue(AppValues.ControlTables.purchase_order_controls, AppValues.ControlValues.distribution_code)
          End If

          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.po_pay_disable_nominal_acc, False) = True AndAlso .EditPanel.PanelInfo.PanelItems.Exists("NominalAccount") Then
            .EditPanel.EnableControl("NominalAccount", False)
          End If
          SetPPAEditable(0)
        Case CareServices.TraderPageType.tpPurchaseOrderCancellation
          .EditPanel.PanelInfo.PanelItems("CancellationReason").Mandatory = True
          .EditPanel.FindDateTimePicker("CancelledOn").ShowCheckBox = False
        Case CareServices.TraderPageType.tpChequeNumberAllocation
          .EditPanel.PanelInfo.PanelItems("ChequeReferenceNumber").Mandatory = True
          .EditPanel.PanelInfo.PanelItems("ChequeReferenceNumber2").Mandatory = True
          .EditPanel.PanelInfo.PanelItems("ChequeNumber").Mandatory = True
          .EditPanel.PanelInfo.PanelItems("ChequeNumber2").Mandatory = True
        Case CareServices.TraderPageType.tpChequeReconciliation
          .EditPanel.FindDateTimePicker("ReconciledOn").ShowCheckBox = False
          .EditPanel.EnableControlList("Amount,ContactNumber", False)
        Case CareServices.TraderPageType.tpActivityEntry
          If mvTA.PayerContactNumber > 0 Then
            SetValueRaiseChanged(.EditPanel, "ContactNumber", mvTA.PayerContactNumber.ToString, False, False, True, True)
          End If
          If FindControl(.EditPanel, "Source", False) IsNot Nothing Then
            Dim vSource As String = ""
            If pRow("DefaultSource").ToString.Length > 0 Then
              vSource = pRow("DefaultSource").ToString
            Else
              If mvTA.MaintenanceOnly Then
                vSource = mvTA.SourceCode
              Else
                vSource = mvTA.TransactionSource
              End If
            End If
            If vSource.Length > 0 Then
              SetValueRaiseChanged(.EditPanel, "Source", vSource, False, False, True, True)
            End If
          End If
        Case CareServices.TraderPageType.tpSuppressionEntry, CareServices.TraderPageType.tpSetStatus, CareServices.TraderPageType.tpGiftAidDeclaration,
        CareServices.TraderPageType.tpAddressMaintenance, CareServices.TraderPageType.tpGoneAway
          If mvTA.PayerContactNumber > 0 Then SetValueRaiseChanged(.EditPanel, "ContactNumber", mvTA.PayerContactNumber.ToString)
          If FindControl(.EditPanel, "Source", False) IsNot Nothing Then
            Dim vSource As String = ""
            If .PageType = CareServices.TraderPageType.tpGiftAidDeclaration AndAlso mvTA.TransactionSource.Length > 0 Then
              vSource = mvTA.TransactionSource
            End If
            If vSource.Length > 0 Then SetValueRaiseChanged(.EditPanel, "Source", vSource)

          End If
          If .PageType = CareServices.TraderPageType.tpGiftAidDeclaration Then
            .EditPanel.SetValue("DeclarationDate", Today.ToString, , , False)
            .EditPanel.SetValue("DeclarationType", "Y") 'Donations
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ga_membership_tax_reclaim) Then .EditPanel.SetValue("DeclarationType2", "Y", , , False) Else .EditPanel.SetValue("DeclarationType2", "N", True, , False) 'Members
            'BR19026
            .EditPanel.SetValue("Method", "W")
            If FindControl(.EditPanel, "StartDate", False) IsNot Nothing Then
              Dim vDR As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetGiftAidData(CareServices.XMLGiftAidDataSelectionTypes.xgdtGiftAidEarliestStartDate, 0))
              Dim vStartDate As String = vDR.Item("GiftAidEarliestStartDate").ToString
              Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.ga_declaration_start_date)
                Case "GIFTAID"
                  .EditPanel.SetValue("StartDate", vStartDate)
                Case Else
                  .EditPanel.SetValue("StartDate", AppValues.TodaysDate)
              End Select
              .EditPanel.FindDateTimePicker("StartDate").MinDate = DateValue(vStartDate)
              .EditPanel.SetValue("EndDate", DateValue(.EditPanel.GetValue("StartDate")).AddDays(1).ToString)
              .EditPanel.FindDateTimePicker("EndDate").MinDate = DateValue(vStartDate)
              .EditPanel.FindDateTimePicker("EndDate").Checked = False

            End If
          ElseIf .PageType = CareServices.TraderPageType.tpSetStatus AndAlso mvTA.DefaultStatus.Length > 0 Then
            .EditPanel.SetValue("Status2", mvTA.DefaultStatus)
          End If
        Case CareServices.TraderPageType.tpPayments
          If mvTA.TransactionType.Length = 0 Then mvTA.TransactionType = "PAYM"
        Case CareServices.TraderPageType.tpConfirmProvisionalTransactions
          SetValueRaiseChanged(.EditPanel, "ContactNumber", pRow("ContactNumber").ToString)
        Case CareServices.TraderPageType.tpBatchInvoiceProduction
          Dim vDaysFromEventStart As Integer = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.invoice_date_from_event_start))
          If vDaysFromEventStart > 0 Then
            .EditPanel.FindDateTimePicker("ToDate").Value = Today.AddDays(vDaysFromEventStart)
            .EditPanel.FindDateTimePicker("ToDate").Checked = True
          End If
          Dim vDisplayInvoices As CheckBox = TryCast(FindControl(.EditPanel, "DisplayInvoices", False), CheckBox)
          If vDisplayInvoices IsNot Nothing Then vDisplayInvoices.Checked = True
          Dim vPartPaid As CheckBox = TryCast(FindControl(.EditPanel, "PartPaidOnly", False), CheckBox)
          If vPartPaid IsNot Nothing AndAlso DirectCast(vPartPaid.Tag, PanelItem).Visible Then vPartPaid.Checked = True
        Case CareNetServices.TraderPageType.tpPostageAndPacking
          SetValueRaiseChanged(.EditPanel, "Product", pRow("CarriageProduct").ToString)
          SetValueRaiseChanged(.EditPanel, "Rate", pRow.Item("CarriageRate").ToString)
          If DoubleValue(pRow.Item("CarriagePrice").ToString) > 0 Then
            .EditPanel.SetValue("Percentage", "")
            .EditPanel.SetValue("Amount2", DoubleValue(pRow.Item("CarriagePrice").ToString).ToString("0.00"))
            .EditPanel.EnableControl("Percentage", False)
            .EditPanel.EnableControl("Amount2", False)
          Else
            .EditPanel.SetValue("Percentage", DoubleValue(pRow.Item("Percentage").ToString).ToString("0.00"))
            .EditPanel.EnableControl("Percentage", True)
            .EditPanel.EnableControl("Amount2", True)
          End If
          mvCarraigePercentage = DoubleValue(pRow.Item("Percentage").ToString)
          .EditPanel.SetValue("Amount", DoubleValue(pRow.Item("TransactionAmount").ToString).ToString("0.00"))
          .EditPanel.SetValue("Amount2", DoubleValue(pRow.Item("PAPAmount").ToString).ToString("0.00"))
        Case CareNetServices.TraderPageType.tpTokenSelection
          Dim vParameters As New ParameterList(True)
          vParameters.Add("ContactNumber", mvTA.PayerContactNumber)
          Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactCreditCards, vParameters)
          Dim vTokenList As ListBox = DirectCast(.EditPanel.Controls("TokenDesc"), ListBox)
          If vDataTable IsNot Nothing Then
            vTokenList.DisplayMember = "TokenDesc"
            vTokenList.ValueMember = "TokenId"
            vTokenList.DataSource = vDataTable
          Else
            vTokenList.DataSource = Nothing
          End If
          .EditPanel.PanelInfo.PanelItems("TokenDescription").Mandatory = False
          .EditPanel.SetControlVisible("TokenDescription", False)
      End Select
      .DefaultsSet = True
      .EditPanel.DataChanged = False
      ValidateDefaults()
    End With
  End Sub

  'Private Sub SetWebPageForTNSHosted()
  '  Dim vParams As New ParameterList(True)
  '  Dim vBatchCategory As String
  '  If mvTA.BatchNumber = 0 Then
  '    'Non batch led
  '    vBatchCategory = mvTA.BatchCategory
  '  Else
  '    'Batch led
  '    vBatchCategory = mvTA.BatchInfo.BatchCategory
  '  End If

  '  vParams("BatchCategory") = vBatchCategory
  '  vParams("Amount") = Convert.ToString(mvTA.TransactionAmount)
  '  vParams("ContactNumber") = Convert.ToString(mvTA.PayerContactNumber)
  '  vParams("AddressNumber") = Convert.ToString(mvTA.PayerAddressNumber)
  '  vParams("Description") = Convert.ToString(mvTA.TransactionType)

  '  Dim vResult As DataRow = Nothing
  '  Try
  '    vResult = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams)
  '    'Set authorised to false as it might have set to true in case of a card transaction declined
  '    mvCardAuthorised = False
  '  Catch vEx As CareException
  '    If vEx.ErrorNumber = CareException.ErrorNumbers.enTNSHostedPaymentNotSetUp Then ShowErrorMessage(vEx.Message)
  '  End Try


  '  If vResult IsNot Nothing AndAlso vResult("CardDetailsPageUrl") IsNot Nothing Then
  '    Dim vCardDetailsPage As New StringBuilder(vResult("CardDetailsPageUrl").ToString)
  '    vCardDetailsPage.Append("&")
  '    vCardDetailsPage.Append("Trader=")
  '    vCardDetailsPage.Append("Y")
  '    vCardDetailsPage.Append("&")
  '    vCardDetailsPage.Append("ContactNumber=")
  '    vCardDetailsPage.Append(mvTA.PayerContactNumber)
  '    vCardDetailsPage.Append("&")
  '    vCardDetailsPage.Append("Amount=")
  '    vCardDetailsPage.Append(mvTA.TransactionAmount)
  '    vCardDetailsPage.Append("&")
  '    vCardDetailsPage.Append("BatchCategory=")
  '    vCardDetailsPage.Append(vBatchCategory)

  '    Dim vWebBrowser As WebBrowser = mvCurrentPage.EditPanel.FindPanelControl(Of WebBrowser)("None2", False)
  '    If vWebBrowser IsNot Nothing Then
  '      vWebBrowser.Stop()
  '      vWebBrowser.Navigate(vCardDetailsPage.ToString)
  '      vWebBrowser.Refresh(WebBrowserRefreshOption.Completely)
  '      vWebBrowser.AllowNavigation = True
  '    End If
  '  End If

  'End Sub

  Private Sub InitCardtAuthorisation()
    If mvLastPage.EditPanel.PanelInfo.PanelItems.Exists("NewToken") AndAlso DirectCast(mvLastPage.EditPanel.FindPanelControl("NewToken"), CheckBox).Checked Then
      Me.CardAuthoriser.CreateToken = True
    ElseIf (mvLastPage.EditPanel.PanelInfo.PanelItems.Exists("TokenDesc") AndAlso DirectCast(mvLastPage.EditPanel.Controls("TokenDesc"), ListBox).SelectedIndex > -1) Then
      Dim vTokenList As ListBox = DirectCast(mvLastPage.EditPanel.Controls("TokenDesc"), ListBox)
      Me.CardAuthoriser.Token = vTokenList.SelectedValue.ToString
    End If

    Me.CardAuthoriser.RequestAuthorisation(mvTA.PayerContactNumber,
                                           mvTA.PayerAddressNumber,
                                           mvTA.TransactionType,
                                           CInt(mvTA.TransactionAmount * 100),
                                           If(mvTA.BatchNumber = 0, mvTA.BatchCategory, mvTA.BatchInfo.BatchCategory), Me.mvTA.MerchantDetailnumber)

  End Sub

  Private Sub SetWebPageForSagePayHosted()
    Dim vParams As New ParameterList(True)
    Dim vBatchCategory As String

    If mvTA.BatchNumber = 0 Then
      'Non batch led
      vBatchCategory = mvTA.BatchCategory
    Else
      'Batch led
      vBatchCategory = mvTA.BatchInfo.BatchCategory
    End If

    With vParams
      .Add("BatchCategory", vBatchCategory)
      .Add("Amount", Convert.ToString(mvTA.TransactionAmount))
      .Add("ContactNumber", mvTA.PayerContactNumber)
      .Add("AddressNumber", mvTA.PayerAddressNumber)
      .Add("Description", mvTA.TransactionType)
      .Add("MakeRequest", "Y")
      .Add("SmartClient", "Y")

      If mvLastPage.EditPanel.PanelInfo.PanelItems.Exists("NewToken") AndAlso DirectCast(mvLastPage.EditPanel.FindPanelControl("NewToken"), CheckBox).Checked Then
        vParams("CreateToken") = "Y"
      ElseIf (mvLastPage.EditPanel.PanelInfo.PanelItems.Exists("TokenDesc") AndAlso DirectCast(mvLastPage.EditPanel.Controls("TokenDesc"), ListBox).SelectedIndex > -1) Then
        Dim vTokenList As ListBox = DirectCast(mvLastPage.EditPanel.Controls("TokenDesc"), ListBox)
        .Add("Token", vTokenList.SelectedValue.ToString)
      End If
    End With
    Dim vResult As DataRow = Nothing
    Try
      vResult = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams)
      'Set authorised to false as it might have set to true in case of a card transaction declined
      mvCardAuthorised = False
      If vResult IsNot Nothing AndAlso vResult("GatewayFormUrl") IsNot Nothing Then
        Dim vWebBrowser As WebBrowser = mvCurrentPage.EditPanel.FindPanelControl(Of WebBrowser)("None2")
        If vWebBrowser IsNot Nothing Then
          vWebBrowser.Navigate(vResult("GatewayFormUrl").ToString)
          vWebBrowser.Refresh(WebBrowserRefreshOption.Completely)
          vWebBrowser.AllowNavigation = True
        End If
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enTNSHostedPaymentNotSetUp, CareException.ErrorNumbers.enSagePayHostedNotSetup,
          CareException.ErrorNumbers.enConnectionFailure, CareException.ErrorNumbers.enInvalidRequest
          ShowErrorMessage(vEx.Message)
      End Select
    End Try
  End Sub


  Private Sub SetPage(ByVal pPageType As CareServices.TraderPageType)
    If Not mvTraderPages.ContainsKey(pPageType.ToString) Then
      Throw New CareException(CareException.ErrorNumbers.enPageNotFound)
    End If
    Dim vNewPage As TraderPage = mvTraderPages(pPageType.ToString)
    If Not mvCurrentPage Is Nothing Then
      mvCurrentPage.EditPanel.Visible = False
      mvLastPage = mvCurrentPage
    End If
    Dim vValid As Boolean = True
    If vNewPage.Menu Then
      Select Case vNewPage.PageType
        Case CareServices.TraderPageType.tpPaymentMethod1
          If mvTA.PayMethodsAtEnd = True Then
            'mvPM1Caption = "&To Be Decided"
            vNewPage.EditPanel.EnableControl("CASH", Not mvEventWLPriceZeroed)
            vNewPage.EditPanel.EnableControl("CHEQ", Not mvEventWLPriceZeroed)
            vNewPage.EditPanel.EnableControl("VOUC", Not mvEventWLPriceZeroed)
            vNewPage.EditPanel.EnableControl("GFIK", Not mvEventWLPriceZeroed)
            vNewPage.EditPanel.EnableControl("POST", Not mvEventWLPriceZeroed)
          End If
          mvFirstTimeOnPM1 = True
        Case CareServices.TraderPageType.tpPaymentMethod2
          Dim vButton As Button = Nothing
          If FindControl(vNewPage.EditPanel, "CURR", False) IsNot Nothing Then
            vButton = vNewPage.EditPanel.FindPanelControl(Of Button)("CURR")
          End If
          Dim vPM1Button As Button = Nothing
          If vButton IsNot Nothing Then
            'Reset Text from Current to Cash etc.
            Select Case mvTA.TransactionPaymentMethod
              Case "CASH", "CHEQ", "POST", "CARD", "CRED", "CQIN", "CCIN"
                vPM1Button = mvTraderPages(CareServices.TraderPageType.tpPaymentMethod1.ToString).EditPanel.FindPanelControl(Of Button)(mvTA.TransactionPaymentMethod, False)
            End Select
            If vPM1Button IsNot Nothing Then
              If mvFirstTimeOnPM1 Then
                If mvTA.PayMethodsAtEnd Then
                Else
                  vButton.Text = vPM1Button.Text
                End If
              End If
            End If
            mvFirstTimeOnPM1 = False
          End If
        Case CareServices.TraderPageType.tpPaymentMethod3
          'Turn off all covenant related items
          SetButtonVisible(vNewPage.EditPanel, "COVT", False)
          SetButtonVisible(vNewPage.EditPanel, "CVDD", False)
          SetButtonVisible(vNewPage.EditPanel, "CVSO", False)
          SetButtonVisible(vNewPage.EditPanel, "CVCC", False)
          'Show or hide the other buttons as required
          SetButtonVisible(vNewPage.EditPanel, "DIRD", Not mvTA.PaymentPlan.HasAutoPaymentMethod)
          SetButtonVisible(vNewPage.EditPanel, "STDO", Not mvTA.PaymentPlan.HasAutoPaymentMethod)
          SetButtonVisible(vNewPage.EditPanel, "CCCA", Not mvTA.PaymentPlan.HasAutoPaymentMethod)
          SetButtonVisible(vNewPage.EditPanel, "MEMB", mvTA.Memberships AndAlso (mvTA.PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptMember) AndAlso (mvTA.PaymentPlan.OneOffPayment = False) AndAlso mvTA.PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptLoan)  ' And (mvEnableCovenant = True Or mvEnableAutoPay = True)) 
          SetButtonVisible(vNewPage.EditPanel, "MAINT", mvTA.PayPlanConvMaintenance)

        Case CareServices.TraderPageType.tpTransactionAnalysis
          If mvTA.CAFCard OrElse mvTA.Voucher Then
            vNewPage.EditPanel.EnableControl("MEMB", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ga_membership_tax_reclaim))
            vNewPage.EditPanel.EnableControl("PAYM", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ga_membership_tax_reclaim))
          Else
            If vNewPage.MenuCount = 1 Then
              mvNewPageType = vNewPage.PageType
              vNewPage.EditPanel.DoTraderChoice()
              mvNewPageType = New Nullable(Of CareServices.TraderPageType)
              vValid = Not (mvTA.PaymentPlan Is Nothing AndAlso mvTA.TransactionType = "MEMC" AndAlso sbp.Text.Length > 0) 'warning message at bottom of form
              If vValid = False OrElse mvIsInvalidCMT Then
                vNewPage = mvCurrentPage
              End If
            End If
          End If
          If mvTA.TransactionPaymentMethod = "CRED" OrElse mvTA.TransactionPaymentMethod = "CQIN" OrElse mvTA.TransactionPaymentMethod = "CCIN" Then
            'Hide buttons that are for options that cannot be paid by Invoice
            SetButtonVisible(vNewPage.EditPanel, "LOAN", False) 'Cannot pay for a loan by Invoice
            SetButtonVisible(vNewPage.EditPanel, "INVC", False) 'Cannot pay an invoice using an invoice
          End If
          'Hide Covenant related items
          SetButtonVisible(vNewPage.EditPanel, "CMEM", False) 'Membership with Covenant
          SetButtonVisible(vNewPage.EditPanel, "CSUB", False) 'Subscription with Covenant
          SetButtonVisible(vNewPage.EditPanel, "CDON", False) 'Regular Donation with Covenant
          If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.CashBatchConfirmation Then
            'BR15574: When confirming a provisional cash transaction hide Payments button
            SetButtonVisible(vNewPage.EditPanel, "PAYM", False)
          End If
        Case CareNetServices.TraderPageType.tpPaymentPlanFromUnbalanceTransaction
          mvTPPDone = True
          'When adding a Donation, cannot use the Payment Plan option so disable the button
          vNewPage.EditPanel.EnableControl("PLAN", (mvTA.TransactionType.Equals("DONS", StringComparison.InvariantCultureIgnoreCase) = False))
      End Select
      If vValid Then
        If vNewPage.PageType <> CareNetServices.TraderPageType.tpPaymentPlanFromUnbalanceTransaction Then vNewPage.EditPanel.FormatButtons()
        cmdNext.Enabled = False
      Else
        'reset buttons
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
      End If
    ElseIf vNewPage.SummaryPage Then
      Select Case vNewPage.PageType
        Case CareServices.TraderPageType.tpInvoicePayments
          If mvInvoicesDGR Is Nothing Then
            mvInvoicesDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("OSInvoices"), DisplayGrid)
            AddHandler mvInvoicesDGR.CheckBoxClicked, AddressOf mvDGR_CheckBoxClicked
            AddHandler mvInvoicesDGR.ButtonClicked, AddressOf mvInvoicesDGR_ButtonClicked
            AddHandler mvInvoicesDGR.ValueChanged, AddressOf mvInvoicesDGR_ValueChanged
          End If
          Dim vList As ParameterList = New ParameterList(True)
          vList("Company") = mvTA.CACompany
          vList("ContactNumber") = mvTA.PayerContactNumber.ToString
          Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vList)
          If vTable IsNot Nothing Then
            vTable.DefaultView.RowFilter = "Company = '" & mvTA.CACompany & "'"
            vList("SalesLedgerAccount") = vTable.Rows(0)("SalesLedgerAccount").ToString
            FillInvoices(mvInvoicesDGR, mvTA.CACompany, vTable.Rows(0)("SalesLedgerAccount").ToString, mvTA.PayerContactNumber.ToString)
          Else
            vList("SalesLedgerAccount") = ""
            mvInvoicesDGR.Clear()
          End If
          Dim vCurrentPayment As Double = mvTA.CalcCurrencyAmount(mvTA.TransactionAmount - mvTA.CurrentLineTotal, True)
          vNewPage.EditPanel.SetValue("CurrentPayment", vCurrentPayment.ToString("N"))
          vNewPage.EditPanel.SetValue("CurrentUnAllocated", vCurrentPayment.ToString("N"))
          vNewPage.EditPanel.SetValue("SalesLedgerAccount", vList("SalesLedgerAccount").ToString)
          'set the last value so that tabbing out of the field does not raise the value changed again
          vNewPage.EditPanel.PanelInfo.PanelItems("SalesLedgerAccount").LastValue = vList("SalesLedgerAccount").ToString

        Case CareServices.TraderPageType.tpMembershipMembersSummary
          If mvMembersDGR Is Nothing Then
            mvMembersDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("Members"), DisplayGrid)
            mvMembersDGR.AllowSorting = False
            AddHandler mvMembersDGR.RowSelected, AddressOf mvMembersDGR_RowSelected
          End If
          mvMembersDGR.Populate(mvTA.MembersDataSet)
          vNewPage.EditPanel.SetValue("MemberCount", mvTA.MemberCount.ToString)
          vNewPage.EditPanel.SetValue("CurrentMembers", mvTA.CurrentMembers.ToString)
          If mvAddBtn Is Nothing Then
            mvAddBtn = DirectCast(vNewPage.EditPanel.FindPanelControl("cmdAdd"), Button)
            AddHandler mvAddBtn.Click, AddressOf cmdAdd_Click
          End If
          If mvAmendBtn Is Nothing Then
            mvAmendBtn = DirectCast(vNewPage.EditPanel.FindPanelControl("cmdAmend"), Button)
            AddHandler mvAmendBtn.Click, AddressOf cmdAmend_Click
          End If
          If mvFindBtn Is Nothing Then
            mvFindBtn = DirectCast(vNewPage.EditPanel.FindPanelControl("cmdFind"), Button)
            AddHandler mvFindBtn.Click, AddressOf cmdFind_Click
          End If
          If mvRemoveBtn Is Nothing Then
            mvRemoveBtn = DirectCast(vNewPage.EditPanel.FindPanelControl("cmdRemove"), Button)
            AddHandler mvRemoveBtn.Click, AddressOf cmdRemove_Click
          End If
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          If mvPPSDGR Is Nothing Then
            mvPPSDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("Details"), DisplayGrid)
            mvPPSDGR.AllowSorting = False
            AddHandler mvPPSDGR.RowSelected, AddressOf mvPPSDGR_RowSelected
          End If
          mvPPSDGR.Populate(mvTA.PPDDataSet)
          If mvPPSDGR.RowCount > 0 Then SetPPDEditable(0)
          vNewPage.EditPanel.SetValue("PPBalance", mvTA.PPBalance.ToString("0.00"))
          vNewPage.EditPanel.SetValue("PPDTotal", mvTA.CurrentPPDLineTotal.ToString("0.00"))
          If mvTA.ApplicationType = ApplicationTypes.atConversion Or mvTA.ApplicationType = ApplicationTypes.atMaintenance Then mvTA.PaymentPlan.PPDTotalAmount = mvTA.CurrentPPDAmount.ToString

        Case CareServices.TraderPageType.tpTransactionAnalysisSummary
          If mvTASDGR Is Nothing Then
            mvTASDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("Analysis"), DisplayGrid)
            mvTASDGR.AllowSorting = False
            AddHandler mvTASDGR.RowSelected, AddressOf mvTASDGR_RowSelected
          End If
          mvTASDGR.Populate(mvTA.AnalysisDataSet)
          vNewPage.EditPanel.SetValue("TransactionAmount", mvTA.TransactionAmount.ToString("N"))
          vNewPage.EditPanel.SetValue("CurrentLineTotal", mvTA.CurrentLineTotal.ToString("N"))
          vNewPage.EditPanel.SetValue("DepositAmount", mvTA.CSDepositAmount.ToString("N"))
          If mvTA.EditExistingTransaction = True Then
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then
              Dim vDT As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
              If vDT.Rows.Count > 0 AndAlso vDT.Columns.Contains("TraderTransactionType") Then
                mvTA.TransactionType = vDT.Rows(0).Item("TraderTransactionType").ToString
              End If
            End If
          End If
          'Call SetPageControls to set the DepositAmount field visibility
          SetPageControls(vNewPage)
        Case CareServices.TraderPageType.tpPurchaseOrderSummary
          If mvPOSDGR Is Nothing Then
            mvPOSDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("PurchaseOrder"), DisplayGrid)
            mvPOSDGR.AllowSorting = False
          End If
          mvPOSDGR.Populate(mvTA.POSDataSet)
          vNewPage.EditPanel.SetValue("POBalance", mvTA.PPBalance.ToString("0.00"))
          vNewPage.EditPanel.SetValue("POSTotal", mvTA.CurrentPPDLineTotal.ToString("0.00"))
        Case CareServices.TraderPageType.tpPurchaseInvoiceSummary
          If mvPISDGR Is Nothing Then
            mvPISDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("PurchaseInvoice"), DisplayGrid)
            mvPISDGR.AllowSorting = False
          End If
          If mvTA.PISDataSet.Tables.Contains("DataRow") AndAlso mvTA.PISDataSet.Tables("DataRow").Rows.Count > 0 Then
            mvPISDGR.Populate(mvTA.PISDataSet)
          End If
          vNewPage.EditPanel.SetValue("PIBalance", mvTA.PPBalance.ToString("0.00"))
          vNewPage.EditPanel.SetValue("PISTotal", mvTA.CurrentPPDLineTotal.ToString("0.00"))
        Case CareNetServices.TraderPageType.tpStatementList
          If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
            Dim vContinue As Boolean = True
            Dim vList As New ParameterList(True)
            If mvStatementGrid Is Nothing Then
              'We have not yet asked the user to select a date
              mvTA.CreditListRecAdditionalCriteria = Nothing
              If mvStatementGrid Is Nothing Then mvStatementGrid = DirectCast(vNewPage.EditPanel.FindPanelControl("StatementDisplayGrid"), DisplayGrid)
              Dim vAPI As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptCLRStatementDate, Nothing, Nothing)
              vAPI.ShowDialog()
              vList = vAPI.ReturnList
              If vList IsNot Nothing AndAlso vList.Count > 0 Then
                'User has selected a date. Check for un-reconciled transactions on that date.
                mvTA.StatementDate = vList("BatchDate")
                For Each vItem As DictionaryEntry In vList
                  Select Case vItem.Key.ToString
                    Case "Culture", "Database", "UserLogname"
                      'Ignore
                    Case "BatchDate"
                      'Ignore as already set
                    Case Else
                      mvTA.CreditListRecAdditionalCriteria(vItem.Key.ToString) = vItem.Value.ToString
                  End Select
                Next
                If Not (mvTA.BatchNumber > 0 AndAlso mvTA.TransactionNumber > 0) Then
                  vContinue = False
                End If
              Else
                'User clicked cancel
                Me.Close()
                vContinue = False
              End If
            End If

            'Need to add a check here as the set page method is getting called twice for the first page
            'and we dont want to ask the user the same question twice
            If vContinue AndAlso mvTA.StatementDate.Length > 0 Then
              vList("BatchDate") = mvTA.StatementDate
              vList("BankAccount") = mvTA.BatchInfo.BankAccount
              vList("TransactionCode") = "99"
              vList("ReconciledStatus") = "U"
              vList("SystemColumns") = "Y"
              If mvTA.CreditListRecAdditionalCriteria IsNot Nothing AndAlso mvTA.CreditListRecAdditionalCriteria.Count > 0 Then
                For Each vItem As DictionaryEntry In mvTA.CreditListRecAdditionalCriteria
                  vList(vItem.Key.ToString) = vItem.Value.ToString
                Next
              End If
              Dim vDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtBankTransactions, vList)
              mvStatementGrid.Populate(vDataSet)
              If mvStatementGrid.RowCount = 0 Then
                'No un-reconciled transactions on specified date
                ShowInformationMessage(InformationMessages.ImNoUnreconciledTransactions)
                mvCloseMe = True
                Me.Close()
              End If
            End If
          End If
      End Select
    Else
      Select Case vNewPage.PageType
        Case CareServices.TraderPageType.tpBatchInvoiceSummary
          If mvInvoiceGrid Is Nothing Then
            mvInvoiceGrid = DirectCast(vNewPage.EditPanel.FindPanelControl("InvoicesDisplayGrid"), DisplayGrid)
            AddHandler mvInvoiceGrid.RowDoubleClicked, AddressOf dgrInvoiceGrid_RowDoubleClicked
            AddHandler mvInvoiceGrid.CheckBoxClicked, AddressOf dgrInvoiceGrid_CheckBoxClicked
          End If
          With mvInvoiceGrid
            .Populate(mvTA.BatchInvoicesDataSet)
            If mvInvoiceGrid.RowCount > 0 Then
              .SetCellsEditable()
              .SetCellsReadOnly()
              .SetCheckBoxColumn("Print")
            End If
          End With
          mvSuppressEvents = True
          vNewPage.EditPanel.SetValue("SelectAll", "N")
          mvSuppressEvents = False
        Case CareServices.TraderPageType.tpCollectionPayments
          If mvCBXDGR Is Nothing Then
            mvCBXDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("CollectionBoxes"), DisplayGrid)
          End If
          mvCBXDGR.Populate(mvTA.CollectionBoxDataSet)
          If mvCBXDGR.RowCount > 0 Then
            With mvCBXDGR
              .SetCellsEditable()
              .SetCellsReadOnly()
              .SetCheckBoxColumn("Pay")
            End With
          End If
        Case CareServices.TraderPageType.tpScheduledPayments
          If mvOPSDGR Is Nothing Then
            mvOPSDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("ScheduledPayments"), DisplayGrid)
            AddHandler mvOPSDGR.ValueChanged, AddressOf mvOPSDGR_ValueChanged
          End If
          mvOPSDGR.AllowSorting = False
          mvOPSDGR.Populate(mvTA.OPSDataSet)
          mvOPSDGR.SetCellsEditable()
          mvOPSDGR.SetCellsReadOnly()
          If mvTA.TransactionType <> "LOAN" Then mvOPSDGR.SetCellsReadOnly(, mvOPSDGR.GetColumn("RevisedAmount"), False) 'Make RevisedAmount writable
        Case CareServices.TraderPageType.tpOutstandingScheduledPayments
          mvOSPDGR = New DisplayGrid
          mvOSPDGR = DirectCast(vNewPage.EditPanel.FindPanelControl("ScheduledPaymentNumber"), DisplayGrid)
          If mvOSPDGR IsNot Nothing Then
            RemoveHandler mvOSPDGR.CheckBoxClicked, AddressOf mvOSPDGR_CheckBoxClicked
            AddHandler mvOSPDGR.CheckBoxClicked, AddressOf mvOSPDGR_CheckBoxClicked
            With mvOSPDGR
              .Populate(mvTA.OSPDataSet)
              If .RowCount > 0 Then
                .SetCellsEditable()
                .SetCellsReadOnly()
                .SetCheckBoxColumn("CheckValue")
              End If
            End With
          End If
        Case CareServices.TraderPageType.tpPurchaseOrderPayments
          SetPPALines(mvTA.PPADataSet)
          If mvPPADGR Is Nothing Then
            mvPPADGR = DirectCast(vNewPage.EditPanel.FindPanelControl("PPAScheduledPayments"), DisplayGrid)
            AddHandler mvPPADGR.RowSelected, AddressOf mvPPADGR_RowSelected
            AddHandler mvPPADGR.ButtonClicked, AddressOf mvPPADGR_ButtonClicked
            AddHandler mvPPADGR.ValueChanged, AddressOf mvPPADGR_ValueChanged
            AddHandler mvPPADGR.CheckBoxClicked, AddressOf mvPPADGR_CheckBoxClicked
          End If
          With mvPPADGR
            .MaxGridRows = If(CanAmendPurchaseOrderAmount, mvTA.PONumberOfPayments, mvTA.PPADataSet.Tables("DataRow").Rows.Count)
            .AllowSorting = False
            .HeaderLines = 2
            .Populate(mvTA.PPADataSet)
            .SetCellsEditable()
            .SetCellsReadOnly()
            .SetSelectionPolicy(FarPoint.Win.Spread.Model.SelectionPolicy.Single)
            .SetColumnWritable("PayeeContactNumber")
            .SetButtonColumn("Finder", "?")

            For vRow As Integer = 0 To mvPPADGR.MaxGridRows - 1
              If mvPPADGR.GetValue(vRow, "AuthorisationStatus").Length > 0 AndAlso BooleanValue(mvPPADGR.GetValue(vRow, "ReadyForPayment")) Then
                mvPPADGR.SetCellsReadOnly(vRow, mvPPADGR.GetColumn("PayByBacs"), True, True)
                mvPPADGR.SetCellsReadOnly(vRow, mvPPADGR.GetColumn("PayeeContactNumber"), True, True)
                mvPPADGR.SetCellsReadOnly(vRow, mvPPADGR.GetColumn("ContactName"), True, True)
                mvPPADGR.SetCellsReadOnly(vRow, mvPPADGR.GetColumn("Finder"), True, True)
                mvPPADGR.SetCellsReadOnly(vRow, mvPPADGR.GetColumn("PopPaymentMethod"), True, True)
              End If
            Next
            mvPPADGR_ContactSelected()
            mvPPADGR_PaymentMethodSelection()
            .SelectRow(0, True)
          End With
          vNewPage.DefaultsSet = False
        Case CareServices.TraderPageType.tpSetStatus
          vNewPage.EditPanel.SetReadOnly("Status", True)
        Case CareServices.TraderPageType.tpBatchInvoiceProduction
          vNewPage.DefaultsSet = False
        Case CareNetServices.TraderPageType.tpCardDetails
          If vNewPage.DefaultsSet = False AndAlso mvTA.OnlineCCAuthorisation AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type) = "SCXLVPCSCP" Then
            vNewPage.EditPanel.FindTextBox("AuthorisationCode").MaxLength = 6 'This is the max allowed value for SecureCXL
          ElseIf mvTA.OnlineCCAuthorisation AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type) = "PROTX" Then
            vNewPage.EditPanel.PanelInfo.PanelItems("CreditCardType").Mandatory = True
          ElseIf mvTA.OnlineCCAuthorisation AndAlso WebBasedCardAuthoriser.IsAvailable Then
            vNewPage.DefaultsSet = False
          End If

        Case CareNetServices.TraderPageType.tpAdvancedCMT
          If mvCMTOldPPD Is Nothing Then
            If vNewPage.EditPanel.FindPanelControl("CMTDetailLines") IsNot Nothing Then
              Dim vDGRS As DisplayGrids = DirectCast(vNewPage.EditPanel.FindPanelControl("CMTDetailLines"), DisplayGrids)
              mvCMTOldPPD = vDGRS.GetDisplayGrid1
              mvCMTNewPPD = vDGRS.GetDisplayGrid2
              mvCMTOldPPD.AllowSorting = False
              mvCMTNewPPD.AllowSorting = False
              AddHandler mvCMTOldPPD.ValueChanged, AddressOf mvCMTOldPPD_ValueChanged
              AddHandler mvCMTNewPPD.ValueChanged, AddressOf mvCMTNewPPD_ValueChanged
            End If
          End If
          mvCMTOldPPD.Populate(mvTA.CMTOldPPDDataSet)
          mvCMTNewPPD.Populate(mvTA.CMTNewPPDDataSet)
          If mvTA.CMTOldPPDDataSet.Tables.Count > 0 AndAlso mvTA.CMTOldPPDDataSet.Tables("DataRow").Rows.Count > 0 Then
            'Add Excess Payments combo
            With mvCMTOldPPD
              .SetCellsEditable()
              .SetCellsReadOnly()
              Dim vValues() As String = {}
              Dim vData() As String = {}
              Dim vCol As Integer = .GetColumn("CMTExcessPaymentType")
              Dim vCodeCol As Integer = .GetColumn("CMTExcessPaymentTypeCode")
              Dim vLineTypeCol As Integer = .GetColumn("PPDLineType")
              If vCol >= 0 Then
                Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCMTExcessPayments)  'CmtExcessPayment,CmtExcessPaymentDesc,CmtExcessPaymentType
                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then Throw New CareException(CareException.ErrorNumbers.enAdvancedCMTIncorrectlySetUp)
                Dim vIndex As Integer = 1
                For Each vRow As DataRow In vDT.Rows
                  Array.Resize(vValues, vIndex)
                  Array.Resize(vData, vIndex)
                  vValues.SetValue(vRow(1).ToString, vIndex - 1)
                  vData.SetValue(vRow(2).ToString, vIndex - 1)
                  vIndex += 1
                Next
                For vRow As Integer = 0 To .RowCount - 1
                  If CType(IntegerValue(.GetValue(vRow, vLineTypeCol)), PaymentPlanInfo.PaymentPlanDetailTypes) <> PaymentPlanInfo.PaymentPlanDetailTypes.OtherCharge Then
                    .SetComboBoxCell(vRow, vCol, vValues, vData)
                    .SetValue(vRow, vCol, .GetValue(vRow, vCodeCol))
                  End If
                Next
                'For other type lines only allow 'Carry Forward'
                If vLineTypeCol >= 0 Then
                  vDT.DefaultView.RowFilter = "CmtExcessPaymentType = 'C'"
                  Dim vTT As DataTable = vDT.DefaultView.ToTable()
                  If vTT.Rows.Count = 1 Then
                    vValues = {}
                    vData = {}
                    Array.Resize(vValues, 1)
                    Array.Resize(vData, 1)
                    vValues.SetValue(vTT.Rows(0).Item(1).ToString, 0)
                    vData.SetValue(vTT.Rows(0).Item(2).ToString, 0)
                    For vRow As Integer = 0 To .RowCount - 1
                      If CType(IntegerValue(.GetValue(vRow, vLineTypeCol)), PaymentPlanInfo.PaymentPlanDetailTypes) = PaymentPlanInfo.PaymentPlanDetailTypes.OtherCharge Then
                        .SetComboBoxCell(vRow, vCol, vValues, vData)
                        .SetValue(vRow, vCol, "C")
                      End If
                    Next
                  End If
                End If
                .SetPreferredColumnWidth(vCol)
              End If
            End With
          End If
          If mvTA.CMTOldPPDDataSet.Tables.Count > 0 OrElse mvTA.CMTNewPPDDataSet.Tables.Count > 0 Then
            'Add Prorate combo
            Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCMTProrateCosts)
            If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
              Dim vValues() As String = {}
              Dim vData() As String = {}
              Dim vIndex As Integer = 1
              For Each vRow As DataRow In vDT.Rows
                Array.Resize(vValues, vIndex)
                Array.Resize(vData, vIndex)
                vData.SetValue(vRow(0).ToString, vIndex - 1)
                vValues.SetValue(vRow(1).ToString, vIndex - 1)
                vIndex += 1
              Next
              If mvTA.CMTOldPPDDataSet.Tables("DataRow").Rows.Count > 0 Then
                With mvCMTOldPPD
                  .SetCellsEditable()
                  Dim vCol As Integer = mvCMTOldPPD.GetColumn("CMTProrateCost")
                  Dim vCodeCol As Integer = mvCMTOldPPD.GetColumn("CMTProrateCostCode")
                  Dim vLineTypeCol As Integer = .GetColumn("PPDLineType")
                  If vCol >= 0 Then
                    For vRowIndex As Integer = 0 To .RowCount - 1
                      If CType(IntegerValue(.GetValue(vRowIndex, vLineTypeCol)), PaymentPlanInfo.PaymentPlanDetailTypes) <> PaymentPlanInfo.PaymentPlanDetailTypes.OtherCharge Then
                        .SetComboBoxCell(vRowIndex, vCol, vValues, vData)
                        .SetValue(vRowIndex, vCol, .GetValue(vRowIndex, vCodeCol))
                      End If
                    Next
                    .SetPreferredColumnWidth(vCol)
                  End If
                End With
              End If
              If mvTA.CMTNewPPDDataSet.Tables("DataRow").Rows.Count > 0 Then
                With mvCMTNewPPD
                  .SetCellsEditable()
                  Dim vCol As Integer = .GetColumn("CMTProrateCost")
                  Dim vCodeCol As Integer = .GetColumn("CMTProrateCostCode")
                  If vCol >= 0 Then
                    .SetComboBoxColumn("CMTProrateCost", vValues, vData)
                    For vRowIndex As Integer = 0 To .RowCount - 1
                      .SetValue(vRowIndex, vCol, .GetValue(vRowIndex, vCodeCol))
                    Next
                    .SetPreferredColumnWidth(vCol)
                  End If
                End With
              End If
              'For other type lines only allow 'Prorate'
              vDT.DefaultView.RowFilter = "LookupCode='F'"
              Dim vTT As DataTable = vDT.DefaultView.ToTable
              If vTT.Rows.Count = 1 Then
                vValues = {}
                vData = {}
                Array.Resize(vValues, 1)
                Array.Resize(vData, 1)
                vData.SetValue(vTT.Rows(0).Item(0).ToString, 0)
                vValues.SetValue(vTT.Rows(0).Item(1).ToString, 0)
                If mvTA.CMTOldPPDDataSet.Tables("DataRow").Rows.Count > 0 Then
                  With mvCMTOldPPD
                    Dim vCol As Integer = .GetColumn("CMTProrateCost")
                    Dim vCodeCol As Integer = .GetColumn("CMTProrateCostCode")
                    Dim vLineTypeCol As Integer = .GetColumn("PPDLineType")
                    If vCol >= 0 AndAlso vLineTypeCol >= 0 Then
                      For vRow As Integer = 0 To .RowCount - 1
                        If CType(IntegerValue(.GetValue(vRow, vLineTypeCol)), PaymentPlanInfo.PaymentPlanDetailTypes) = PaymentPlanInfo.PaymentPlanDetailTypes.OtherCharge Then
                          .SetComboBoxCell(vRow, vCol, vValues, vData)
                          .SetValue(vRow, vCol, "F")
                        End If
                      Next
                    End If
                  End With
                End If
                If mvTA.CMTNewPPDDataSet.Tables("DataRow").Rows.Count > 0 Then
                  With mvCMTNewPPD
                    Dim vCol As Integer = .GetColumn("CMTProrateCost")
                    Dim vLineTypeCol As Integer = .GetColumn("PPDLineType")
                    If vCol >= 0 AndAlso vLineTypeCol >= 0 Then
                      For vRow As Integer = 0 To .RowCount - 1
                        If CType(IntegerValue(.GetValue(vRow, vLineTypeCol)), PaymentPlanInfo.PaymentPlanDetailTypes) = PaymentPlanInfo.PaymentPlanDetailTypes.OtherCharge Then
                          .SetComboBoxCell(vRow, vCol, vValues, vData)
                        End If
                      Next
                    End If
                  End With
                End If
              End If
            Else
              Throw New CareException(CareException.ErrorNumbers.enAdvancedCMTIncorrectlySetUp)
            End If
          End If
        Case CareNetServices.TraderPageType.tpPaymentPlanProducts
          If Me.mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanSummary Then
            'PaymentPlanProducts will contain it's previous values or defaults, so clear fields required for new product. 
            vNewPage.EditPanel.FindPanelControl(Of TextLookupBox)("Product").Text = String.Empty
            vNewPage.EditPanel.FindPanelControl(Of TextLookupBox)("Rate").Text = String.Empty
            vNewPage.EditPanel.FindPanelControl(Of TextBox)("Quantity").Text = "1"
            vNewPage.EditPanel.FindPanelControl(Of TextBox)("Amount").Clear()
            vNewPage.EditPanel.FindPanelControl(Of TextBox)("Balance").Clear()
          End If
      End Select
      SetPageControls(vNewPage)
      vNewPage.EditPanel.FillDeferredCombos(vNewPage.EditPanel)
    End If
    If mvTA.OnlineCCAuthorisation Then prgBar.Visible = (vNewPage.PageType = CareNetServices.TraderPageType.tpCardDetails)
    mvCurrentPage = vNewPage
    If (mvCurrentPage.Menu = True AndAlso mvCurrentPage.MenuCount = 1) OrElse (mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentMethod2 AndAlso mvTA.TransactionType = "MEMB" AndAlso (mvTA.TransactionPaymentMethod = "CAFC" OrElse mvTA.TransactionPaymentMethod = "VOUC")) Then
      'Don't set it as visible yet since we will step over it (Since CAF transactions only support donations & memberships don't show the PM2 options)
    Else
      mvCurrentPage.EditPanel.Visible = True
      If vValid Then mvCurrentPage.EditPanel.Focus()
      'Clear any messages and reset edit/delete buttons appropriate for the grid being displayed
      'This needs to be done here now that mvCurrentPage has been set.
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          EPL_ShowMessage(mvCurrentPage.EditPanel, "")
          SetPPDEditable(1)
        Case CareServices.TraderPageType.tpTransactionAnalysisSummary
          EPL_ShowMessage(mvCurrentPage.EditPanel, "")
          SetAnalysisEditable(0)
      End Select
    End If
    If System.Diagnostics.Debugger.IsAttached = True OrElse AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_trader_page_display) = True Then
      Me.Text = mvTA.Description & " - Page Code: " & mvCurrentPage.PageCode
    End If
  End Sub

  Private Sub SetButtonVisible(ByVal pPanel As EditPanel, ByVal pParameterName As String, ByVal pVisible As Boolean)
    Dim vControl As Control = FindControl(pPanel, pParameterName, False)
    If vControl IsNot Nothing Then
      vControl.Enabled = pVisible
      vControl.Visible = pVisible
      If pVisible Then
        If vControl.Height = 0 AndAlso vControl.Tag IsNot Nothing Then
          Dim vPanelItem As PanelItem = DirectCast(vControl.Tag, PanelItem)
          vControl.Height = vPanelItem.ControlHeight
        End If
      Else
        'Cannot use visible property as if form not shown yet control is not visible so set height to zero
        vControl.Height = 0
      End If
    End If
  End Sub

  Private Sub SetPageControls(ByVal pPage As TraderPage)
    'This method should only be used to adjust the controls on the page when they should change according to the 
    'type of transaction e.g. Product Details which changes depending on donation or product sale
    Select Case pPage.PageType
      Case CareServices.TraderPageType.tpAmendEventBooking
        pPage.EditPanel.EnableControlList("ContactNumber,AddressNumber,EventNumber,OptionNumber,BookingNumber,Product,Rate", False)
      Case CareServices.TraderPageType.tpProductDetails
        Dim vDonation As Boolean
        If mvTA.TransactionType = "DONS" Then vDonation = True
        With pPage.EditPanel
          .SetControlVisible("DeceasedContactNumber", vDonation AndAlso .PanelInfo.PanelItems("DeceasedContactNumber").Visible)
          .PanelInfo.PanelItems("DeceasedContactNumber").Mandatory = vDonation
          If FindControl(pPage.EditPanel, "LineTypeG", False) IsNot Nothing Then
            'Page does not have the option buttons
            .SetControlVisible("LineTypeG", vDonation AndAlso .PanelInfo.PanelItems("LineTypeG").Visible)
            .SetControlVisible("LineTypeH", vDonation AndAlso .PanelInfo.PanelItems("LineTypeH").Visible)
            .SetControlVisible("LineTypeS", vDonation AndAlso .PanelInfo.PanelItems("LineTypeS").Visible)
            .EnableControl("CreditedContactNumber", False)
          Else
            .SetControlVisible("LineType_G", vDonation AndAlso .PanelInfo.PanelItems("LineType").Visible)
            .SetControlVisible("LineType_H", vDonation AndAlso .PanelInfo.PanelItems("LineType2").Visible)
            .SetControlVisible("LineType_S", vDonation AndAlso .PanelInfo.PanelItems("LineType3").Visible)
            If FindControl(pPage.EditPanel, "CreditedContactNumber", False) IsNot Nothing Then .SetControlVisible("CreditedContactNumber", False)
          End If

          .SetControlVisible("GrossAmount", mvTA.PayerHasDiscount)
          .SetControlVisible("Discount", mvTA.PayerHasDiscount)
          .EnableControl("Warehouse", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_stock_multiple_warehouses))
          If FindControl(pPage.EditPanel, "EventBookingNumber", False) IsNot Nothing Then
            If vDonation = True Then
              .SetControlVisible("EventBookingNumber", False)
            ElseIf mvTA.EventMultipleAnalysis = False OrElse mvTA.EventBookingDataSet.Tables.Count > 1 Then
              .EnableControl("EventBookingNumber", False)
              .SetControlVisible("EventBookingNumber", .PanelInfo.PanelItems("EventBookingNumber").Visible)
            Else
              .SetControlVisible("EventBookingNumber", .PanelInfo.PanelItems("EventBookingNumber").Visible)
            End If
            If mvTA.EventNumber.Length > 0 Then .SetControlVisible("EventBookingNumber", False)
          End If
          If FindControl(pPage.EditPanel, "ServiceBookingNumber", False) IsNot Nothing Then .SetControlVisible("ServiceBookingNumber", Not vDonation AndAlso .PanelInfo.PanelItems("ServiceBookingNumber").Visible)
        End With
      Case CareServices.TraderPageType.tpTransactionDetails
        'If we have a CreditSale then disable Contact/Address fields to prevent user changing them
        'This is to ensure that the transaction payer is the Credit Customer contact (See also SetDefaults)
        With pPage.EditPanel
          .EnableControl("ContactNumber", (mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN"))
          .EnableControl("AddressNumber", (mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN"))
        End With
        Dim referenceTextBox As TextBox = DirectCast(pPage.EditPanel.FindPanelControl("Reference"), TextBox)
        If referenceTextBox IsNot Nothing Then
          referenceTextBox.AllowDrop = True
          RemoveHandler referenceTextBox.DragOver, AddressOf txt_DragEnter
          RemoveHandler referenceTextBox.DragDrop, AddressOf txt_DragDrop
          AddHandler referenceTextBox.DragOver, AddressOf txt_DragEnter
          AddHandler referenceTextBox.DragDrop, AddressOf txt_DragDrop
        End If
      Case CareNetServices.TraderPageType.tpTransactionAnalysisSummary
        'Make the DepositAmount field visible when Payment Method is Credit Sale, when not using 'Pay Methods At End'
        'where we have the Company Credit Controls 'Deposit Percentage' value greater than 0 and the Trader Sundry Credit Notes flag is not set
        With pPage.EditPanel
          Dim vVisible As Boolean = Not mvTA.PayMethodsAtEnd And mvTA.TransactionPaymentMethod = "CRED" And mvTA.CSDepositPercentage > 0 And Not mvTA.CreditNotes
          .SetControlVisible("DepositAmount", vVisible)
          Dim vWasHidden As Boolean = .PanelInfo.PanelItems("DepositAmount").Hidden
          .PanelInfo.PanelItems("DepositAmount").Hidden = Not vVisible
          If Not vVisible AndAlso Not vWasHidden Then
            'DepositAmount textbox now hidden so adjust top position of Analysis Grid control below to move up
            Dim vOffset As Integer = .PanelInfo.PanelItems("Analysis").ControlTop - .PanelInfo.PanelItems("DepositAmount").ControlTop
            .AdjustItem(.PanelInfo.PanelItems("Analysis"), vOffset)
          ElseIf vVisible AndAlso vWasHidden Then
            'DepositAmount textbox was hidden now un-hidden so adjust top position of Analysis Grid control below to move back down
            Dim vOffset As Integer = .PanelInfo.PanelItems("Analysis").ControlTop - .PanelInfo.PanelItems("CurrentLineTotal").ControlTop
            .AdjustItem(.PanelInfo.PanelItems("Analysis"), -vOffset)
          End If
        End With
    End Select
  End Sub

  Private Sub SetPageValuesForEditing()
    'Populate the controls on the required page for editing an analysis line
    'DO NOT USE - Use SetDefaults
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    Dim vDonation As Boolean

    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpAmendEventBooking
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        If vRow IsNot Nothing Then
          With vRow
            vEPL.SetValue("ContactNumber", .Item("DeliveryContactNumber").ToString, True)
            vEPL.SetValue("AddressNumber", .Item("DeliveryAddressNumber").ToString, True)
            vEPL.SetValue("EventNumber", .Item("EventNumber").ToString, True)
            vEPL.SetValue("BookingNumber", .Item("EventBookingNumber").ToString, True)
            vEPL.SetValue("OptionNumber", .Item("BookingOptionNumber").ToString, True)
            SetValueRaiseChanged(vEPL, "Product", .Item("ProductCode").ToString, True)
            SetValueRaiseChanged(vEPL, "Rate", .Item("Rate").ToString, True)
            Dim vAdultQty As String = .Item("AdultQuantity").ToString
            Dim vChildQty As String = .Item("ChildQuantity").ToString
            If FindControl(vEPL, "AdultQuantity", False) IsNot Nothing Then
              vEPL.SetValue("AdultQuantity", vAdultQty)
              vEPL.SetValue("ChildQuantity", vChildQty)
            End If
            If vAdultQty.Length > 0 OrElse vChildQty.Length > 0 Then
              vEPL.SetValue("Quantity", (IntegerValue(vAdultQty) + IntegerValue(vChildQty)).ToString, True)
            Else
              vEPL.SetValue("Quantity", .Item("Quantity").ToString)
            End If
            If FindControl(vEPL, "StartTime", False) IsNot Nothing Then
              vEPL.SetValue("StartTime", "07:00")
              vEPL.SetValue("EndTime", "20:00")
              vEPL.SetValue("StartTime", .Item("StartTime").ToString)
              vEPL.SetValue("EndTime", .Item("EndTime").ToString)
            End If
            vEPL.SetValue("Amount", .Item("Amount").ToString)
          End With
          If vEPL.FindTextLookupBox("EventNumber").CareEventInfo.EventPricingMatrix.Length > 0 Then vEPL.EnableControl("Amount", False)
          If vEPL.GetValue("AdultQuantity").Length > 0 OrElse vEPL.GetValue("ChildQuantity").Length > 0 Then vEPL.EnableControl("Quantity", False)
          vEPL.SetErrorField("Amount", "")     'Amount may have been validated and an error set before we populated it
          If mvTA.TransactionType.Length = 0 Then mvTA.TransactionType = "EVNT"
        End If
      Case CareServices.TraderPageType.tpCollectionPayments
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          SetValueRaiseChanged(vEPL, "AppealCollectionNumber", .Item("CollectionNumber").ToString)    'This will pre-populate other fields as well
          vEPL.FindTextBox("Amount").Focus()                                                          'This forces validation etc. of AppealCollectionNumber & PisNumber and prevents grid data from being changed
          vEPL.SetValue("PisNumber", .Item("CollectionPisNumber").ToString)
          If vEPL.GetValue("PisNumber").Length > 0 Then GetCollectionBoxes(vEPL)
          vEPL.SetValue("Amount", .Item("Amount").ToString)
          vEPL.SetValue("DeceasedContactNumber", .Item("DeceasedContactNumber").ToString)
          Dim vBoxNumbers As String = .Item("CollectionBoxNumbers").ToString
          If vBoxNumbers.Length > 0 Then
            Dim vBoxNumber() As String = vBoxNumbers.Split(","c)
            For vIndex As Integer = 0 To vBoxNumber.GetUpperBound(0)
              For vRowIndex As Integer = 0 To mvCBXDGR.RowCount - 1
                With mvCBXDGR
                  If IntegerValue(.GetValue(vRowIndex, "CollectionBoxNumber")) = IntegerValue(vBoxNumber(vIndex)) Then
                    .SetValue(vRowIndex, "Pay", "True")
                    Exit For
                  End If
                End With
              Next
            Next
          End If
        End With

      Case CareServices.TraderPageType.tpPaymentPlanProducts, CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
        Dim vRow As DataRow = mvTA.GetDataSetLine(mvTA.PPDDataSet, mvCurrentPPDLine)
        vEPL.SetValue("Product", vRow.Item("Product").ToString, IntegerValue(vRow.Item("SubscriptionNumber").ToString) > 0)
        vEPL.PanelInfo.PanelItems("Product").LastValue = vRow.Item("Product").ToString  'BR18968 - to ensure Value Changed functionality works correctly in Validating Handeler
        vEPL.FindPanelControl(Of TextLookupBox)("Product").OriginalText = vRow.Item("Product").ToString  'BR18968 -
        vEPL.SetValue("Rate", vRow.Item("Rate").ToString)
        vEPL.PanelInfo.PanelItems("Rate").LastValue = vRow.Item("Rate").ToString
        vEPL.FindPanelControl(Of TextLookupBox)("Rate").OriginalText = vRow.Item("Rate").ToString
        If vEPL.FindPanelControl(Of TextLookupBox)("Product").GetDataRow IsNot Nothing Then
          mvTA.ProductVATCategory = vEPL.FindPanelControl(Of TextLookupBox)("Product").GetDataRowItem("ProductVATCategory")
        End If
        If vEPL.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow IsNot Nothing Then
          Dim vRateRow As DataRow = vEPL.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow()
          Dim vCurrentPrice As Double
          'This is in PPN page of Copy of Payment Plan Maintenance
          If vRateRow("UseModifiers").ToString = "Y" Then
            'J1563: Use the TodaysDate if the TransactionDate is not set
            Dim vModifierDate As Date = CDate(If(mvTA.TransactionDate.Length > 0, mvTA.TransactionDate, AppValues.TodaysDate))
            vCurrentPrice = DataHelper.GetModifierPrice(vRateRow("Product").ToString, vRateRow("Rate").ToString, vModifierDate, IntegerValue(mvTA.PayerContactNumber), BooleanValue(vRateRow.Item("VatExclusive").ToString))
          Else
            vCurrentPrice = DoubleValue(vEPL.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow("CurrentPrice").ToString)
          End If
          mvTA.LinePrice = vCurrentPrice
          mvTA.LinePriceVATEx = BooleanValue(vRateRow.Item("VatExclusive").ToString)
          vEPL.EnableControl("Amount", (mvTA.LinePrice = 0))
        End If
        vEPL.SetValue("DistributionCode", vRow.Item("DistributionCode").ToString)
        vEPL.SetValue("Quantity", vRow.Item("Quantity").ToString, , , , True)
        If (mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "CMEM") OrElse
            (mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance) Then
        End If
        'For VAT-Exclusive Rates need to enable the NetFixedAmount if we have it, otherwise use the default Amount (GrossFixedAmount)
        Dim vHasNetFixedAmount As Boolean
        If FindControl(vEPL, "NetFixedAmount", False) IsNot Nothing AndAlso vEPL.PanelInfo.PanelItems("NetFixedAmount").Visible Then
          vHasNetFixedAmount = True
        End If
        If vHasNetFixedAmount Then
          If mvTA.LinePriceVATEx AndAlso mvTA.LinePrice = 0 Then
            'NetFixedAmount is visible so enable it and disable Amount
            vEPL.EnableControl("NetFixedAmount", True)
            vEPL.SetValue("NetFixedAmount", vRow.Item("NetFixedAmount").ToString, , , , True)
            vEPL.EnableOrSetValueDisable("Amount", False, vRow.Item("Amount").ToString)
          Else
            'Always disable NetFixedAmount 
            vEPL.EnableOrSetValueDisable("NetFixedAmount", False, "")
            vEPL.SetValue("Amount", vRow.Item("Amount").ToString, , , , True)
          End If
        Else
          vEPL.SetValue("Amount", vRow.Item("Amount").ToString, , , , True)
        End If
        vEPL.SetValue("Balance", vRow.Item("Balance").ToString)
        vEPL.SetValue("Arrears", vRow.Item("Arrears").ToString)
        vEPL.SetValue("Source", vRow.Item("Source").ToString)
        vEPL.SetValue("DespatchMethod", vRow.Item("DespatchMethod").ToString)
        vEPL.SetValue("ContactNumber", vRow.Item("ContactNumber").ToString)
        vEPL.SetValue("AddressNumber", vRow.Item("AddressNumber").ToString)
        If FindControl(vEPL, "TimeStatus", False) IsNot Nothing Then
          vEPL.SetValue("TimeStatus", vRow.Item("TimeStatus").ToString)
        End If
        If FindControl(vEPL, "CommunicationNumber", False) IsNot Nothing Then
          vEPL.SetValue("CommunicationNumber", vRow.Item("CommunicationNumber").ToString, IntegerValue(vRow.Item("SubscriptionNumber").ToString) = 0)
        End If
        If FindControl(vEPL, "EffectiveDate", False) IsNot Nothing Then
          vEPL.SetValue("EffectiveDate", vRow.Item("EffectiveDate").ToString, True)
        End If
        'Do not allow editing of the Product if it is for a future membership type
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
          'Maintenance
          If mvCurrentPPDLine = GetLowestFutureLineNumber() Then
            vEPL.EnableControl("Product", False)
          End If
          Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
          mvTA.ContactVATCategory = vContactInfo.VATCategory
        End If
        If vRow.Table.Columns.Contains("AccruesInterest") = True AndAlso (BooleanValue(vRow.Item("AccruesInterest").ToString) = True OrElse BooleanValue(vRow.Item("LoanInterest").ToString) = True) Then
          'Loan Capital or Interest product so user not allowed to change the product
          vEPL.EnableControl("Product", False)
          'For Loan Capital product only allow the FixedAmount to be changed
          If BooleanValue(vRow.Item("AccruesInterest").ToString) = True AndAlso (mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance) Then
            vEPL.EnableControlList("Rate,DistributionCode,Quantity,Arrears,DespatchMethod,Source,ContactNumber,AddressNumber,TimeStatus,CommunicationNumber,EffectiveDate,ValidFrom,ValidTo", False)
            If FindControl(vEPL, "NetFixedAmount", False) IsNot Nothing AndAlso vEPL.GetValue("NetFixedAmount").Length = 0 Then
              vEPL.EnableControl("NetFixedAmount", False)
            End If
            If vEPL.GetValue("Amount").Length = 0 Then
              vEPL.EnableControl("Amount", False)
            End If
          End If
        End If
        If vRow.Table.Columns.Contains("IncentiveLineType") AndAlso vRow.Item("IncentiveLineType").ToString.Length > 1 AndAlso vRow.Item("IncentiveLineType").ToString.Substring(1, 1).ToUpper = "I" Then
          'Special initial period incentive - quantity has been set to the plan term in months so set to 1 and do not allow it to be changed
          vEPL.SetValue("Quantity", "1", True)
        Else
          vEPL.EnableControl("Quantity", True)
        End If
        If FindControl(vEPL, "ValidFrom", False) IsNot Nothing AndAlso vEPL.PanelInfo.PanelItems("ValidFrom").Visible Then
          vEPL.SetValue("ValidFrom", String.Empty, True)
          If vRow.Table.Columns.Contains("ValidFrom") AndAlso Not String.IsNullOrEmpty(vRow.Item("ValidFrom").ToString()) Then
            vEPL.SetValue("ValidFrom", vRow.Item("ValidFrom").ToString(), True)
          End If
          vEPL.EnableControl("ValidFrom", True)
        End If
        If FindControl(vEPL, "ValidTo", False) IsNot Nothing AndAlso vEPL.PanelInfo.PanelItems("ValidTo").Visible Then
          vEPL.SetValue("ValidTo", String.Empty, True)
          If vRow.Table.Columns.Contains("ValidTo") AndAlso Not String.IsNullOrEmpty(vRow.Item("ValidTo").ToString()) Then
            vEPL.SetValue("ValidTo", vRow.Item("ValidTo").ToString, True)
          End If
          vEPL.EnableControl("ValidTo", True)
        End If
      Case CareServices.TraderPageType.tpPayments
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          If IntegerValue(.Item("PaymentPlanNumber").ToString) > 0 Then
            mvTA.PaymentPlan = New PaymentPlanInfo(IntegerValue(.Item("PaymentPlanNumber").ToString))
            mvTA.TransactionType = "PAYM"
          End If
          SetDefaults(vRow)
        End With

      Case CareServices.TraderPageType.tpProductDetails
        If (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment OrElse mvTA.EditExistingTransaction) Then
          Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
          If vContactInfo IsNot Nothing Then mvTA.ContactVATCategory = vContactInfo.VATCategory
        End If
        'What about if the controls are not on the page?
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          mvTA.TransactionType = .Item("TraderTransactionType").ToString
          Select Case mvTA.TransactionType
            Case "D", "F", "G", "S", "H", "DONS"
              vDonation = True
            Case "P"
              If .Item("ProductCode").ToString.Length > 0 Then
                'Check if Donation Product
                Dim vList As New ParameterList(True)
                vList("Product") = .Item("ProductCode").ToString
                If Not (mvTA.EditExistingTransaction OrElse mvTA.ExistingAdjustmentTran) Then vList("FindProductType") = "B"
                Dim vDS As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftProducts, vList)
                vDonation = (DataHelper.GetTableFromDataSet(vDS).Rows(0).Item("Donation").ToString = "Y")
                mvTA.StockSales = BooleanValue(DataHelper.GetTableFromDataSet(vDS).Rows(0).Item("StockItem").ToString)
              End If
          End Select
          If mvTA.GiftInKind Then
            Select Case mvTA.TransactionType
              Case "P", "H"
                vDonation = False
                mvTA.TransactionType = "SALE"
            End Select
          End If
          If vDonation = True OrElse mvTA.StockSales = True Then
            mvTA.TransactionType = IIf(vDonation = True, "DONS", "SALE").ToString
            SetPageControls(mvCurrentPage)   'mvTraderApplication.TransactionType has changed so Page controls may need to change
          End If
          If vDonation = True Then
            vEPL.FindComboBox("Warehouse").DataSource = Nothing
            vEPL.SetValue("LastStockCount", "", True)
          End If
          If IntegerValue(.Item("ProductNumber").ToString) > 0 Then vEPL.SetValue("ProductNumber", .Item("ProductNumber").ToString)
          vEPL.SetValue("Quantity", .Item("Quantity").ToString)
          SetValueRaiseChanged(vEPL, "Product", .Item("ProductCode").ToString)
          vEPL.EnableControl("Product", (Not (mvTA.TransactionType = "SALE" AndAlso mvTA.StockSales = True)))
          mvOldProductCode = .Item("ProductCode").ToString
          vEPL.FindTextLookupBox("Rate").Focus()        'Force validation of Product
          SetValueRaiseChanged(vEPL, "Rate", .Item("Rate").ToString)
          mvOldRate = .Item("Rate").ToString
          vEPL.FindTextLookupBox("Product").Focus()     'Force validation of Rate
          vEPL.SetValue("DistributionCode", .Item("DistributionCode").ToString)
          If IntegerValue(.Item("DeceasedContactNumber").ToString) > 0 Then vEPL.SetValue("DeceasedContactNumber", .Item("DeceasedContactNumber").ToString)
          Dim vGotCContact As Boolean = (FindControl(vEPL, "CreditedContactNumber", False) IsNot Nothing)
          If FindControl(vEPL, "LineTypeG", False) IsNot Nothing Then
            Select Case .Item("TraderLineType").ToString
              Case "D"
                If vGotCContact Then vEPL.SetValue("LineTypeH", "Y")
                vEPL.SetValue("LineTypeG", "Y")
              Case "F"
                If vGotCContact Then vEPL.SetValue("LineTypeS", "Y")
                vEPL.SetValue("LineTypeG", "Y")
              Case "G", "H", "S"
                vEPL.SetValue("LineType" & .Item("TraderLineType").ToString, "Y")
              Case Else
                vEPL.EnableControl("DeceasedContactNumber", (IntegerValue(.Item("DeceasedContactNumber").ToString) > 0))
            End Select
            If vGotCContact AndAlso IntegerValue(.Item("CreditedContactNumber").ToString) > 0 Then vEPL.SetValue("CreditedContactNumber", .Item("CreditedContactNumber").ToString)
          Else
            Select Case .Item("TraderLineType").ToString
              Case "H", "S", "G"
                vEPL.SetValue("LineType_" & .Item("TraderLineType").ToString, .Item("TraderLineType").ToString)
              Case Else
                vEPL.SetValue("LineType_G", "G")
            End Select
            If vGotCContact Then vEPL.SetValue("CreditedContactNumber", "", True)
          End If
          vEPL.SetValue("Source", .Item("Source").ToString)
          Dim vAmount As Double = DoubleValue(.Item("Amount").ToString)
          vEPL.SetValue("VatAmount", DoubleValue(.Item("VatAmount").ToString).ToString("0.00"), , , False)
          If mvTA.ShowVATExclusiveAmount Then
            'Display Amount less VAT
            vAmount -= DoubleValue(.Item("VATAmount").ToString)
          End If
          vEPL.SetValue("Amount", vAmount.ToString("0.00"))
          If .Item("GrossAmount").ToString.Length > 0 Then vEPL.SetValue("GrossAmount", DoubleValue(.Item("GrossAmount").ToString).ToString("0.00"))
          If .Item("Discount").ToString.Length > 0 Then vEPL.SetValue("Discount", DoubleValue(.Item("Discount").ToString).ToString("0.00"))
          vEPL.SetValue("When", .Item("LineDate").ToString)
          vEPL.SetValue("Notes", .Item("Notes").ToString)
          vEPL.SetValue("DespatchMethod", .Item("DespatchMethod").ToString)

          'BR14890: SC financial adjusting/analyse transactions are creating a zero for contact & address number on BTA which is incorrect, 
          'if NULL then should remain NULL and not zero.   
          'Only set the delivery contact and address number if they were present in the original transaction
          If IntegerValue(.Item("DeliveryContactNumber")) > 0 Then
            vEPL.SetValue("ContactNumber", .Item("DeliveryContactNumber").ToString)
            vEPL.SetValue("AddressNumber", .Item("DeliveryAddressNumber").ToString)
          End If

          If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment Then
            vEPL.PanelInfo.PanelItems("ContactNumber").Mandatory = IntegerValue(.Item("DeliveryContactNumber")) > 0
            vEPL.PanelInfo.PanelItems("AddressNumber").Mandatory = IntegerValue(.Item("DeliveryContactNumber")) > 0
          End If

          If .Item("SalesContactNumber").ToString.Length > 0 Then vEPL.SetValue("SalesContactNumber", .Item("SalesContactNumber").ToString)
          If .Item("WarehouseCode").ToString.Length > 0 Then vEPL.SetValue("Warehouse", .Item("WarehouseCode").ToString)
          If mvTA.StockSales Then mvTA.SetStockTransactionValues(IntegerValue(.Item("StockTransactionID").ToString), IntegerValue(.Item("Issued").ToString), .Item("ProductCode").ToString, .Item("WarehouseCode").ToString, IntegerValue(.Item("Quantity").ToString))
          If vRow.Table.Columns.Contains("EventBookingNumber") AndAlso FindControl(vEPL, "EventBookingNumber", False) IsNot Nothing Then
            vEPL.SetValue("EventBookingNumber", .Item("EventBookingNumber").ToString, True, True)
          End If
          If vRow.Table.Columns.Contains("ServiceBookingNumber") AndAlso FindControl(vEPL, "ServiceBookingNumber", False) IsNot Nothing AndAlso FindControl(vEPL, "ServiceBookingNumber", False).Visible Then
            Dim vSBExisting As Boolean = False
            For vLineNumber As Integer = 0 To mvCurrentRow
              'Check previous TAS Lines for Service Booking Line (TraderLineType 'V')
              If vLineNumber = mvCurrentRow Then Exit For
              Dim vDataRow As DataRow = mvTA.GetDataSetLine(mvTA.AnalysisDataSet, vLineNumber + 1)
              If vDataRow IsNot Nothing Then
                With vDataRow
                  If .Item("TraderLineType").ToString = "V" Then
                    vSBExisting = True
                    Exit For
                  End If
                End With
              End If
            Next
            vEPL.SetValue("ServiceBookingNumber", .Item("ServiceBookingNumber").ToString)
            vEPL.EnableControl("ServiceBookingNumber", Not vSBExisting AndAlso mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.None)
          End If
        End With

      Case CareServices.TraderPageType.tpPurchaseInvoiceProducts
        Dim vRow As DataRow = mvTA.GetDataSetLine(mvTA.PISDataSet, mvCurrentPPDLine)
        With vRow
          vEPL.SetValue("LineItem", .Item("LineItem").ToString)
          vEPL.SetValue("LinePrice", .Item("LinePrice").ToString)
          vEPL.SetValue("Quantity", .Item("Quantity").ToString)
          vEPL.SetValue("Amount", .Item("Amount").ToString)
          SetPOILineItemControls(vEPL, "PILineItem", vRow.Item("LineItem").ToString)
        End With

      Case CareServices.TraderPageType.tpPurchaseOrderProducts
        Dim vRow As DataRow = mvTA.GetDataSetLine(mvTA.POSDataSet, mvCurrentPPDLine)
        With vRow
          vEPL.SetValue("LineItem", .Item("LineItem").ToString)
          vEPL.SetValue("LinePrice", .Item("LinePrice").ToString)
          vEPL.SetValue("Quantity", .Item("Quantity").ToString)
          vEPL.SetValue("Amount", .Item("Amount").ToString)
          vEPL.SetValue("NominalAccount", .Item("NominalAccount").ToString)
          vEPL.SetValue("DistributionCode", .Item("DistributionCode").ToString)
          If vEPL.FindTextLookupBox("Product", False) IsNot Nothing Then
            vEPL.SetValue("Product", .Item("Product").ToString)
            EPL_ValueChanged(vEPL, "Product", .Item("Product").ToString)
            If .Item("Product").ToString.Length = 0 Then vEPL.FindComboBox("Warehouse").DataSource = Nothing
          End If
          SetPOILineItemControls(vEPL, "POLineItem", vRow.Item("LineItem").ToString)
          If Not CanAmendPurchaseOrderAmount Then vEPL.EnableControlList("LinePrice,Amount,Quantity", False)
        End With
      Case CareNetServices.TraderPageType.tpSetStatus
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          vEPL.SetValue("ContactNumber", .Item("DeliveryContactNumber").ToString)
          EPL_ValueChanged(vEPL, "ContactNumber", .Item("DeliveryContactNumber").ToString)
          vEPL.SetValue("Status2", .Item("ContactStatus").ToString)
        End With
      Case CareNetServices.TraderPageType.tpCancelGiftAidDeclaration
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          vEPL.SetValue("DeclarationNumber", .Item("DeclarationNumber").ToString)
          vEPL.SetValue("CancellationReason", .Item("CancellationReason").ToString)
          vEPL.SetValue("CancellationSource", .Item("Source").ToString)
        End With
      Case CareNetServices.TraderPageType.tpCancelPaymentPlan
        Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
        With vRow
          vEPL.SetValue("PaymentPlanNumber", .Item("PaymentPlanNumber").ToString)
          vEPL.SetValue("CancellationReason", .Item("CancellationReason").ToString)
          vEPL.SetValue("CancellationSource", .Item("Source").ToString)
        End With
    End Select
    If mvLastPage.PageType = CareServices.TraderPageType.tpTransactionAnalysisSummary Then
      Dim vRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
      Dim vTransactionType As String = vRow.Item("TraderTransactionType").ToString
      Select Case vTransactionType
        Case "M", "C", "O", "MEMB", "MEMC", "SUBS", "DONR", "CMEM", "CSUB", "CDON"
          mvOldSourceCode = vRow.Item("Source").ToString
        Case "SALE", "DONS", "G", "S", "H", "P", "CSRT"
          Dim vLineType As String = vRow.Item("TraderLineType").ToString()
          Select Case vLineType
            Case "E", "A", "V", "VC", "B"
            Case Else
              mvOldSourceCode = vRow.Item("Source").ToString
          End Select
      End Select
    End If

  End Sub


  Private Sub SetValuesForExistingTransaction(ByVal pDataSet As DataSet)
    'Populate the controls for each page used in the transaction
    Dim vPage As TraderPage
    Dim vTable As DataTable = pDataSet.Tables(0)
    Dim vRow As DataRow = vTable.Rows(0)
    Dim vBankDetails As Boolean
    Dim vControlName As String
    Dim vPageCode As String
    Dim vFound As Boolean

    If mvTA.EditExistingTransaction OrElse mvTA.ExistingAdjustmentTran Then
      For vIndex As Integer = 0 To vTable.Columns.Count - 1
        vPageCode = vTable.Columns(vIndex).ColumnName.Substring(0, 3)
        vControlName = vTable.Columns(vIndex).ColumnName
        If vTable.Columns(vIndex).ColumnName.Contains("_") Then
          vControlName = vControlName.Substring(4)
          If vPageCode = "CCU" AndAlso vControlName = "TermsFrom" Then
            vControlName &= "_" & vRow(vIndex).ToString
          End If
          vFound = False
          For Each vPage In mvTA.Pages
            If vPage.PageCode = vPageCode Then
              If vPage.DefaultsSet = False Then
                vPage.EditPanel.FillDeferredCombos(vPage.EditPanel)
                vPage.DefaultsSet = True
              End If
              If FindControl(vPage.EditPanel, vControlName, False) IsNot Nothing Then vPage.EditPanel.SetValue(vControlName, MultiLine(vRow(vIndex).ToString))
              If vPage.PageType = CareServices.TraderPageType.tpBankDetails Then vBankDetails = True
              vFound = True
            End If
            vPage.EditPanel.DataChanged = False
            If vFound Then Exit For
          Next
        Else
          'TraderApplication values
          Select Case vTable.Columns(vIndex).ColumnName
            Case "TransactionDate"
              mvTA.TransactionDate = vRow.Item(vIndex).ToString
            Case "TransactionAmount"
              mvTA.TransactionAmount = DoubleValue(vRow.Item(vIndex).ToString)
              mvTA.OriginalTransactionAmount = DoubleValue(vRow.Item(vIndex).ToString)
            Case "TransactionCurrencyAmount"
              mvTA.OriginalTransactionCurrencyAmount = DoubleValue(vRow.Item(vIndex).ToString)
            Case "TransactionPaymentMethod"
              mvTA.TransactionPaymentMethod = vRow.Item(vIndex).ToString
            Case "TransactionSource"
              mvTA.TransactionSource = vRow.Item(vIndex).ToString
            Case "PayerContactNumber"
              mvTA.SetPayerContact(IntegerValue(vRow.Item("PayerContactNumber").ToString), IntegerValue(vRow.Item("PayerAddressNumber").ToString))
          End Select
        End If
      Next
      If vBankDetails Then
        mvNextPageCode = "BKD"
        vPage = GetTraderPage(CareServices.TraderPageType.tpBankDetails)
        SetBankDetails(vPage.EditPanel, "SortCode", vPage.EditPanel.GetValue("SortCode"), "N")  'BR13853: Don't verify the account at this stage
        SetBankDetails(vPage.EditPanel, "AccountNumber", vPage.EditPanel.GetValue("AccountNumber"), mvTA.AlbacsBankDetails)
      End If

      If vTable.Columns.Contains("CDC_CreditOrDebitCard") Then
        mvNextPageCode = "CDC"
        vPage = GetTraderPage(CareServices.TraderPageType.tpCardDetails)
        SetCreditOrDebitCard(vPage.EditPanel, vPage.EditPanel.GetValue("CreditOrDebitCard"))
      End If

      'Set totals
      mvTA.SetLineTotal()
      vPage = mvTA.Pages(CareServices.TraderPageType.tpTransactionAnalysisSummary.ToString)
      If vPage IsNot Nothing Then vPage.EditPanel.SetValue("TransactionAmount", mvTA.TransactionAmount.ToString("N"))
    End If

  End Sub

  ''' <summary>Gets the required Trader Page. Throws a Page Not Found Exception if Trader has not been configured to use that page.</summary>
  Private Function GetTraderPage(ByVal pPageType As CareServices.TraderPageType) As TraderPage
    Dim vPage As TraderPage = Nothing
    If mvTA.Pages.ContainsKey(pPageType.ToString) Then
      vPage = mvTA.Pages(pPageType.ToString)
    Else
      Throw New CareException(CareException.ErrorNumbers.enPageNotFound)
    End If
    Return vPage
  End Function

#End Region

#Region " Pre-Server Processing "

  Private Sub AddAnalysisLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      'BR15742 padded zeros to line number in order to set the collection in ascending order of linenumber
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("TraderAnalysisLine" & vRow("LineNumber").ToString.PadLeft(5, CChar("0"))) = vRow
      Next
    End If
  End Sub

  ''' <summary>
  ''' Adds the Currently selected row on Trader Form TAS to the Parameter List. If no data exists this does nothing.
  ''' </summary>
  ''' <param name="plist">List of Parameters to be sent to the server</param>
  ''' <remarks>Originally intended for Cash Batch Maintenance where this data is passed back and forth between the Analysis form and Transactions form, via the Analysis Summary form 
  ''' so that Source can be synchronised. The Analysis Data needs to be passed to the server to facilitate this, but normal Trader operation doesn't want to do this, as ther cannot be any 
  ''' Analysis data present when the Transaction is being created.
  ''' </remarks>
  Private Sub AddSelectedAnalysisLine(ByVal plist As ParameterList)
    Dim vParameterName As String
    If mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0 Then
      vParameterName = "TraderAnalysisLine" & mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow).Item("LineNumber").ToString().PadLeft(5, CChar("0"))
      If Not plist.ContainsKey(vParameterName) Then
        plist.ObjectValue(vParameterName) = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
      End If
      mvTA.EditLineNumber = IntegerValue(mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow).Item("LineNumber").ToString())
    End If

  End Sub

  Private Sub AddCollectionBoxData(ByVal pList As ParameterList)
    'Retrieve comma-separated list of Box Numbers & Amounts
    Dim vBoxAmounts As New StringBuilder
    Dim vBoxNumbers As New StringBuilder
    Dim vFound As Boolean = False
    Dim vPay As Object

    With mvCBXDGR
      For vIndex As Integer = 0 To .RowCount - 1
        vPay = .GetValue(vIndex, "Pay")
        If vPay IsNot Nothing AndAlso CBool(vPay.ToString) = True Then
          vFound = True
          vBoxAmounts.Append(.GetValue(vIndex, "Amount"))
          vBoxNumbers.Append(.GetValue(vIndex, "CollectionBoxNumber"))
        End If
      Next
    End With
    If vFound Then
      pList("CollectionBoxAmounts") = vBoxAmounts.ToString
      pList("CollectionBoxNumbers") = vBoxNumbers.ToString
    End If
  End Sub

  Private Sub AddInvoiceLines(ByVal pList As ParameterList)
    Dim vIndex As Integer = 1
    If mvOSInvoices IsNot Nothing Then
      'add the outstanding invoices/these are the invoices we are paying
      For Each vInvoice As InvoiceInfo In mvOSInvoices
        If vInvoice.GetAmountPaid > 0 Then
          pList.InvoiceValue("InvoiceLine" & vIndex.ToString) = vInvoice
          vIndex += 1
        End If
      Next
    End If
    If mvCashInvoices IsNot Nothing Then
      'add the cash invpoices/credit notes. There are the invoices that may have been used to pay the above
      For Each vInvoice As InvoiceInfo In mvCashInvoices
        If vInvoice.InvoiceNumber > 0 AndAlso vInvoice.AmountUsed > 0 Then
          pList.InvoiceValue("InvoiceLine" & vIndex.ToString) = vInvoice
          vIndex += 1
        End If
      Next
    End If
  End Sub

  Private Sub AddMemberLines(ByVal pList As ParameterList)
    Dim vMaxLineNumber As Integer = 0
    Dim vTable As DataTable = mvTA.MembersDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("MemberLine" & vRow("LineNumber").ToString) = vRow
        If IntegerValue(vRow("LineNumber").ToString) > vMaxLineNumber Then vMaxLineNumber = IntegerValue(vRow.Item("LineNumber").ToString)
      Next
      pList.Add("MaxMemberLine", vMaxLineNumber.ToString)
    End If
  End Sub

  Private Function AddMembershipData(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True
    Dim vPage As TraderPage
    If mvTA.TransactionType = "MEMC" Then
      vPage = mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString)
    Else
      vPage = mvTraderPages(CareServices.TraderPageType.tpMembership.ToString)
    End If
    If vPage.DefaultsSet Then
      Dim vList As New ParameterList
      'Only add Page controls if the page has been used
      vValid = vPage.EditPanel.AddValuesToList(vList, False, EditPanel.AddNullValueTypes.anvtAll)
      For Each vItem As DictionaryEntry In vList
        pList(vItem.Key.ToString) = vItem.Value.ToString
      Next
      If Not vValid Then Return False
    End If

    If vValid AndAlso mvTA.MemberContactToAdd > 0 Then
      pList.IntegerValue("FinderContactNumber") = mvTA.MemberContactToAdd
    End If

    Return vValid

  End Function

  Private Sub AddOPSLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.OPSDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("OPSLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub
  ''' <summary>
  ''' BR19606 For Transaction History, Analysis followed by Edit or Delete will change the Order Payment Schedule, when Edit or Delete are clicked. This is the original order payment history before the change.
  ''' It is used to restore the Order Payment Schedule if Analysis is Cancelled. Using Samrt Client as a temporary store for the stateless web services.
  ''' This takes the stored OPS and sends it to the server
  ''' </summary>
  ''' <param name="pList">List of Parameters being passed to the server, the OPS is added to the list.</param>
  ''' <remarks></remarks>
  Private Sub AddOriginalOPSLines(pList As ParameterList)
    If mvTA.OriginalOPS IsNot Nothing Then
      pList.ObjectValue("OriginalOPSLine") = mvTA.OriginalOPS.Rows(0)
    End If
  End Sub

  Private Sub AddRemovedSchPaymentLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.RemovedSchPaymentsDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("RemovedSchPaymentLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub

  Private Sub AddOSPLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.OSPDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("OPSLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub

  Private Sub AddPPDLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.PPDDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      vTable.DefaultView.Sort = "LineNumber DESC"
      Dim vTT As DataTable = vTable.DefaultView.ToTable
      For Each vRow As DataRow In vTT.Rows
        pList.ObjectValue("PPDLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub

  Private Sub AddPOSLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.POSDataSet.Tables("DataRow")
    If vTable IsNot Nothing Then
      vTable.DefaultView.Sort = "LineNumber DESC"
      Dim vTT As DataTable = vTable.DefaultView.ToTable
      For Each vRow As DataRow In vTT.Rows
        pList.ObjectValue("POSLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub

  Private Sub AddPISLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.PISDataSet.Tables("DataRow")
    If vTable IsNot Nothing Then
      vTable.DefaultView.Sort = "LineNumber DESC"
      Dim vTT As DataTable = vTable.DefaultView.ToTable
      For Each vRow As DataRow In vTT.Rows
        pList.ObjectValue("PISLine" & vRow("LineNumber").ToString) = vRow
      Next
    End If
  End Sub

  Private Sub AddSelectedInvoices(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.BatchInvoicesDataSet.Tables("DataRow")
    If vTable IsNot Nothing Then
      Dim vCopyTable As DataTable = vTable.Copy   'Make a copy of the Table before applying the filter
      vCopyTable.DefaultView.RowFilter = "Print = 'True'"
      Dim vFilterTable As DataTable = vCopyTable.DefaultView.ToTable
      Dim vBatchNumbers As New StringBuilder
      Dim vTransNumbers As New StringBuilder
      Dim vEventNumbers As New StringBuilder
      Dim vCompany As New StringBuilder
      For Each vRow As DataRow In vFilterTable.Rows
        If vBatchNumbers.Length > 0 Then vBatchNumbers.Append(",")
        vBatchNumbers.Append(vRow("BatchNumber").ToString)
        If vTransNumbers.Length > 0 Then vTransNumbers.Append(",")
        vTransNumbers.Append(vRow("TransactionNumber").ToString)
        If vEventNumbers.Length > 0 Then vEventNumbers.Append(",")
        If vRow("Company").ToString.Length > 0 AndAlso vCompany.Length = 0 Then
          vCompany.Append(vRow("Company").ToString)
        End If
      Next

      pList("BatchNumbers") = vBatchNumbers.ToString
      pList("TransactionNumbers") = vTransNumbers.ToString
      pList("Company") = vCompany.ToString
    End If
  End Sub

  Private Function AddTransactionData(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True
    If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpTransactionAnalysisSummary.ToString) Then
      If mvTA.TransactionAmount <> mvTA.CurrentLineTotal Then
        If mvTA.EditExistingTransaction Then
          If mvTA.BatchInfo.PostedToCashBook = True AndAlso (mvTA.BatchInfo.BatchType = CareNetServices.BatchTypes.Cash OrElse mvTA.BatchInfo.BatchType = CareNetServices.BatchTypes.CashWithInvoice) AndAlso (mvTA.TransactionAmount <> mvTA.CurrentLineTotal) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cb_edit_transaction_amount) = False Then
            ShowInformationMessage(InformationMessages.ImCannotChangeTransAmount)
            Return False
          End If
        End If
        'BR18557 - Added tpTransactionDeatils condition.
        If (mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionAnalysisSummary OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpTransactionDetails) OrElse
           (mvTA.PayMethodsAtEnd AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpCreditCustomer AndAlso
            (mvTA.TransactionPaymentMethod = "CQIN" OrElse mvTA.TransactionPaymentMethod = "CCIN")) Then
          vValid = CheckTransactionTotal()
        End If
      End If
      If vValid Then
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpProductDetails.ToString) AndAlso mvTA.TransactionPaymentMethod = "CASH" _
          AndAlso mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.max_donation_unknown_address).Length > 0 Then
          Dim vTotalDonationAmount As Double
          For Each vRow As DataRow In mvTA.AnalysisDataSet.Tables("DataRow").Select("TraderTransactionType = 'DONS'")
            vTotalDonationAmount += DoubleValue(vRow("Amount").ToString)
          Next
          vValid = IsTotalDonationValid(vTotalDonationAmount, mvTA.PayerContactNumber, mvTA.PayerAddressNumber)
        End If
      End If
      If vValid Then
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpProductDetails.ToString) AndAlso (mvTA.TransactionPaymentMethod = "CQIN" OrElse mvTA.TransactionPaymentMethod = "CCIN") _
          AndAlso mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then
          Dim vCSUnderPaymentProduct As String = AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_under_payment_product)
          Dim vCSOverPaymentProduct As String = AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_over_payment_product)
          If vCSUnderPaymentProduct.Length > 0 AndAlso vCSOverPaymentProduct.Length > 0 Then
            If mvTA.AnalysisDataSet.Tables("DataRow").Select(String.Format("ProductCode = '{0}' OR ProductCode = '{1}'", vCSUnderPaymentProduct, vCSOverPaymentProduct)).Length > 1 Then
              ShowErrorMessage(InformationMessages.ImOnlyOneCSUnderOverPayProduct)
              vValid = False
            End If
          End If
        End If
      End If
      If vValid Then
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpActivityEntry.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpActivityEntry.ToString).DefaultsSet Then
          Dim vADS As ActivityDataSheet = TryCast(FindControl(mvTraderPages(CareServices.TraderPageType.tpActivityEntry.ToString).EditPanel, "Activity", False), ActivityDataSheet)
          If vADS IsNot Nothing AndAlso vADS.Initialised Then vADS.AddActivities(pList)
        End If
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpSuppressionEntry.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpSuppressionEntry.ToString).DefaultsSet Then
          Dim vSDS As SuppressionDataSheet = TryCast(FindControl(mvTraderPages(CareServices.TraderPageType.tpSuppressionEntry.ToString).EditPanel, "MailingSuppression", False), SuppressionDataSheet)
          If vSDS IsNot Nothing AndAlso vSDS.Initialised Then vSDS.AddSuppressions(pList)
        End If
      End If
    End If

    If vValid AndAlso mvTraderPages.ContainsKey(CareServices.TraderPageType.tpPaymentPlanSummary.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpPaymentPlanSummary.ToString).DefaultsSet Then
      'only do this if we have been to the pps page atleast once
      vValid = AddPPSTransactionData(pList)
    ElseIf vValid AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderSummary OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseInvoiceSummary) AndAlso mvCurrentPage.DefaultsSet Then
      vValid = CheckPOISummaryBalance(pList)
    End If
    If vValid AndAlso (mvTraderPages.ContainsKey(CareServices.TraderPageType.tpDirectDebit.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpDirectDebit.ToString).DefaultsSet) _
       Or (mvTraderPages.ContainsKey(CareServices.TraderPageType.tpStandingOrder.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpStandingOrder.ToString).DefaultsSet) _
       Or (mvTraderPages.ContainsKey(CareServices.TraderPageType.tpCreditCardAuthority.ToString) AndAlso mvTraderPages(CareServices.TraderPageType.tpCreditCardAuthority.ToString).DefaultsSet) Then
      If mvTA.ApplicationType = ApplicationTypes.atConversion Or mvTA.ApplicationType = ApplicationTypes.atMaintenance Then 'Maint
        Dim vConfig As AppValues.ConfigurationOptions = Nothing
        Dim vMessage As String = "" ' length of vm,essage is also being used to determine if one of the auto page is actually the current page
        Select Case mvCurrentPage.PageType
          Case CareServices.TraderPageType.tpDirectDebit
            vConfig = AppValues.ConfigurationOptions.fp_dd_set_next_payment_due
            vMessage = QuestionMessages.QmUpdatePPDSourceWithDDSource
          Case CareServices.TraderPageType.tpStandingOrder
            vConfig = AppValues.ConfigurationOptions.fp_so_set_next_payment_due
            vMessage = QuestionMessages.QmUpdatePPDSourceWithSOSource
          Case CareServices.TraderPageType.tpCreditCardAuthority
            vConfig = AppValues.ConfigurationOptions.fp_cc_set_next_payment_due
            vMessage = QuestionMessages.QmUpdatePPDSourceWithCCASource
        End Select
        If vMessage.Length > 0 Then
          If AppValues.ConfigurationOption(vConfig) Then
            Dim vStartDate As Date = New DateHelper(mvCurrentPage.EditPanel.GetValue("StartDate")).DateValue
            If vStartDate > Now And mvTA.PaymentPlan.NextPaymentDue < vStartDate Then
              Dim vResetNPDToStart As Boolean = False
              Dim vPaymentFrequency As DataRow = Nothing
              If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpPaymentPlanMaintenance.ToString) Then vPaymentFrequency = mvTraderPages(CareServices.TraderPageType.tpPaymentPlanMaintenance.ToString).EditPanel.FindTextLookupBox("PaymentFrequency").GetDataRow()
              If mvTA.PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptMember AndAlso
                ((vPaymentFrequency IsNot Nothing AndAlso vPaymentFrequency("Frequency").ToString = "1" AndAlso vPaymentFrequency("Interval").ToString = "1") OrElse
                (mvTA.PaymentPlan.PaymentFreq.Frequency = 1 AndAlso mvTA.PaymentPlan.PaymentFreq.Interval = 1)) Then
                'Regular monthly donation Payment Plan - do not ask user, just update the Payment Plan
                ShowInformationMessage(InformationMessages.ImNPDRDWillBeUpdated, mvTA.PaymentPlan.NextPaymentDue.ToString, vStartDate.ToString)    'The Next Payment Due is currently set to {0} Start Date is {1} - The Next Payment Due and Renewal dates will be updated to the Start Date.
                vResetNPDToStart = True
              Else
                'Ask user if they wish to update the Payment Plan
                If ShowQuestion(QuestionMessages.QmOverrideNPDWithStartDate, MessageBoxButtons.YesNo, mvTA.PaymentPlan.NextPaymentDue.ToString, vStartDate.ToString) = System.Windows.Forms.DialogResult.Yes Then    'The Next Payment Due is currently set to {0} Start Date is {1} - Override Next Payment Due with Start Date?
                  vResetNPDToStart = True
                End If
              End If
              pList("ResetNPDToStartDate") = IIf(vResetNPDToStart, "Y", "N").ToString
            End If
          End If
          'Update Source???
          If mvTA.ApplicationType = ApplicationTypes.atConversion Then
            Dim vSource As String = mvCurrentPage.EditPanel.GetValue("Source")
            If Len(vMessage) > 0 AndAlso Len(vSource) > 0 AndAlso vSource <> mvTA.PaymentPlan.Source Then
              'Sources are different so may need to update Payment Plan Details source
              Dim vUpdateSource As Boolean = False
              Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_ppd_use_auto_pmnt_source)
                Case "HU"
                  'Always update the source without asking the user
                  vUpdateSource = True
                Case "HN"
                  'Never update the source without asking the user
                  vUpdateSource = False
                Case Else
                  'Either 'SN' or 'SU' (default) - ask user whether to update the source
                  vUpdateSource = (ShowQuestion(vMessage, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
              End Select
              pList("UpdatePPDSource") = IIf(vUpdateSource, "Y", "N").ToString
            End If
          End If
        End If
      End If
    End If

    If vValid Then
      'Add page controls for pages that have been used and are required from cmdFinished
      For Each vPage As TraderPage In mvTraderPages
        Select Case vPage.PageType
          Case CareServices.TraderPageType.tpBankDetails, CareServices.TraderPageType.tpCardDetails, CareServices.TraderPageType.tpComments,
               CareServices.TraderPageType.tpCreditCustomer, CareServices.TraderPageType.tpTransactionDetails,
               CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpPaymentPlanDetails,
               CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpCreditCardAuthority, CareServices.TraderPageType.tpStandingOrder,
               CareServices.TraderPageType.tpPaymentPlanMaintenance, CareServices.TraderPageType.tpChangeMembershipType, CareServices.TraderPageType.tpMembershipPayer,
               CareServices.TraderPageType.tpPurchaseOrderDetails, CareServices.TraderPageType.tpPurchaseInvoiceDetails, CareServices.TraderPageType.tpLoans
            Dim vList As New ParameterList
            If vPage.DefaultsSet Then
              'Only add Page controls if the page has been used
              vValid = vPage.EditPanel.AddValuesToList(vList, False, EditPanel.AddNullValueTypes.anvtAll)
              If vPage.PageType = CareNetServices.TraderPageType.tpComments Then
                If mvTA.TransactionNote.Length > 0 Then
                  'As batch transaction notes can be maintained from transaction details or comments page types, the list value should be taken as transaction note from tader application because
                  'this holds the latest value from both these page types.
                  pList("Notes") = mvTA.TransactionNote
                End If
              End If
              For Each vItem As DictionaryEntry In vList
                pList(vPage.PageCode & "_" & vItem.Key.ToString) = vItem.Value.ToString
              Next
              If Not vValid Then Return False
              Select Case vPage.PageType
                Case CareServices.TraderPageType.tpBankDetails, CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpStandingOrder
                  pList("NewBank") = CBoolYN(mvTA.NewBank)
                  pList.IntegerValue("BankDetailsNumber") = mvTA.BankDetailsNumber
                Case CareServices.TraderPageType.tpCreditCardAuthority
                  pList.IntegerValue("CreditCardDetailsNumber") = mvTA.CreditCardDetailsNumber
                Case CareServices.TraderPageType.tpCreditCustomer
                  If mvTA.NewCreditCustomer Then
                    pList("UpdateCreditCustomerAddress") = "Y"
                  Else
                    pList("UpdateCreditCustomerAddress") = CBoolYN(mvTA.SaveCreditCustomerAddressChange)
                  End If
                  pList("StorePaymentTerms") = CBoolYN(mvTA.SavePaymentTerms)
                  If pList("CCU_TermsPeriod").Equals("Y") Then
                    pList("CCU_TermsPeriod") = "M"
                  Else
                    pList("CCU_TermsPeriod") = "D"
                  End If
                Case CareServices.TraderPageType.tpTransactionDetails, CareServices.TraderPageType.tpMembership
                  If vPage.PageType = CareServices.TraderPageType.tpTransactionDetails Then
                    If vList.Contains("AdditionalReference1") Then pList(vPage.PageCode & "_AdditionalReference1Caption") = vPage.EditPanel.FindLabel("AdditionalReference1").Text
                    If vList.Contains("AdditionalReference2") Then pList(vPage.PageCode & "_AdditionalReference2Caption") = vPage.EditPanel.FindLabel("AdditionalReference2").Text
                  End If
              End Select
            End If
          Case CareServices.TraderPageType.tpContactSelection
            'Special case for this page - we do not normally need anything from here but for non-financial we may need the mailing code if it was set
            If vPage.DefaultsSet = True AndAlso FindControl(vPage.EditPanel, "Mailing", False) IsNot Nothing Then
              Dim vMailingCode As String = vPage.EditPanel.GetValue("Mailing")
              pList(vPage.PageCode & "_Mailing") = vMailingCode
            End If
        End Select
      Next
      'BR13392
      If mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atTransaction Then
        If pList.ValueIfSet("TRD_Mailing").Length = 0 AndAlso pList.ValueIfSet("CSE_Mailing").Length = 0 Then
          AddCMDMailingCode(pList)
          pList("TRD_Mailing") = pList.ValueIfSet("Mailing")
        End If
      End If
      '
      If pList.Contains("TRD_Notes") Then
        'The COM page notes has been maintained on TRD page and possibly COM page, ensure the list contains the latest changed value which is held in transaction note property of the trader application 
        pList("COM_Notes") = mvTA.TransactionNote
      End If
    End If

    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpPaymentPlanMaintenance
        pList("PPDTotal") = mvTA.CurrentPPDLineTotal.ToString
    End Select

    Return vValid
  End Function

  Private Function AddPPSTransactionData(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True

    If mvTA.PPDLines = 0 Then    'No detail lines exist in the grid
      vValid = False
      ShowWarningMessage(InformationMessages.ImPayPlanNeedsOneDetailLine)
    ElseIf mvTA.CurrentPPDLineTotal < 0 Then
      'The sum of Payment Plan Details is less than zero!
      vValid = False
      ShowWarningMessage(InformationMessages.ImPayPlanBalanceNegativeAmount)
    Else
      With mvTA
        If .PPBalance <> .CurrentPPDLineTotal Then      'Totals are the same?
          Dim vWOMissedPayments As Boolean
          If (mvTA.ApplicationType = ApplicationTypes.atMaintenance OrElse (mvTA.ApplicationType = ApplicationTypes.atConversion AndAlso mvTA.PPPaymentType = "MAINT")) AndAlso mvTraderPages.ContainsKey(CareNetServices.TraderPageType.tpPaymentPlanMaintenance.ToString) Then
            Dim vWOString As String = GetOptionalPageValue(CareNetServices.TraderPageType.tpPaymentPlanMaintenance, "WriteOffMissedPayments")
            If vWOString.Length = 0 Then
              vWOMissedPayments = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pp_wo_missed_payments, False)
            Else
              vWOMissedPayments = BooleanValue(vWOString)
            End If
          End If
          Dim vResult As System.Windows.Forms.DialogResult
          If .PPBalance = 0 Then
            vResult = ShowQuestion(QuestionMessages.QmPPBalanceNotSet, MessageBoxButtons.YesNo, .CurrentPPDLineTotal.ToString("N"))
          ElseIf vWOMissedPayments Then
            vResult = ShowQuestion(QuestionMessages.QmPPBalanceNotMatchWithWO, MessageBoxButtons.YesNo, .PPBalance.ToString("N"), .CurrentPPDLineTotal.ToString("N"))
          Else
            vResult = ShowQuestion(QuestionMessages.QmPPBalanceNotMatch, MessageBoxButtons.YesNo, .PPBalance.ToString("N"), .CurrentPPDLineTotal.ToString("N"))
          End If
          If vResult = System.Windows.Forms.DialogResult.Yes Then
            If .ApplicationType = ApplicationTypes.atMaintenance Then '"MAINT"
              If .PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptLoan Then SetPageValue(CareServices.TraderPageType.tpPaymentPlanMaintenance, "Balance", .CurrentPPDLineTotal.ToString) 'Loans do not have this page.
              mvTA.PPBalance = .CurrentPPDLineTotal
              pList("PPBalance") = mvTA.PPBalance.ToString
            Else
              SetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Balance", .CurrentPPDLineTotal.ToString)
              mvTA.PPBalance = .CurrentPPDLineTotal
              pList("PPBalance") = mvTA.PPBalance.ToString
              If mvTA.PPAmount.Length > 0 Then
                SetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Amount", .CurrentPPDLineTotal.ToString)
                mvTA.PPAmount = .CurrentPPDLineTotal.ToString
                pList("FixedAmount") = mvTA.PPBalance.ToString
              End If
            End If
          Else
            vValid = False
          End If
        End If

        If mvTA.PaymentPlan IsNot Nothing AndAlso (mvTA.PPNumbersCreated Is Nothing OrElse Not mvTA.PPNumbersCreated.ContainsKey(mvTA.PaymentPlan.PaymentPlanNumber.ToString)) Then
          'it has not been created just now
          With mvTA.PaymentPlan
            If .Existing And mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanSummary Then
              Dim vBalance As Double
              If mvTA.ApplicationType = ApplicationTypes.atConversion AndAlso mvTA.TransactionPaymentMethod = "MEMB" Then
                vBalance = DoubleValue(GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Balance"))
              ElseIf .PlanType = PaymentPlanInfo.ppType.pptLoan Then
                vBalance = mvTA.PPBalance
              Else
                vBalance = DoubleValue(GetPageValue(CareServices.TraderPageType.tpPaymentPlanMaintenance, "Balance"))
              End If
              If vBalance = 0 And .Balance <> 0 And .TermUnits = PaymentPlanInfo.OrderTermUnits.otuNone Then
                'Payment plan now has zero balance & renewal-pending = N
                Dim vMessage As String = InformationMessages.ImBalSetToZero
                Dim vNewDate As Date
                If .RenewalDate < Today Or .RenewalPending Then 'voldrenewalpending
                  'Either the Renewal Date is in the past OR the Balance is being set to zero after R&R has been run but before the Renewal Date has been reached.
                  'So both the Renewal Date and Next Payment Due should be rolled forward.
                  vNewDate = .CalculateRenewalDate(.RenewalDate, True)         'Use the same code to recalculate the Renewal Date as is used by Batch Processing and OPS Regeneration
                  vMessage = vMessage & " " & GetInformationMessage(QuestionMessages.QmSetRDNPDTo, vNewDate.ToString)
                Else
                  'Otherwise we simply want to write off the remainder of this year's balance and move Next Payment Due forward to the Renewal Date
                  vNewDate = .RenewalDate
                  vMessage = vMessage & GetInformationMessage(QuestionMessages.QmSetNPDTo, vNewDate.ToString)
                End If
                If ShowQuestion(vMessage, MessageBoxButtons.YesNo, vNewDate.ToString) = System.Windows.Forms.DialogResult.Yes Then
                  pList("UpdateDates") = "Y"
                  pList("RenewalDate") = vNewDate.ToString
                  pList("NextPaymentDue") = vNewDate.ToString
                End If
              End If
            End If
          End With


          If vValid Then
            If .ApplicationType = ApplicationTypes.atMaintenance Then '"MAINT"
              Dim vAmount As Double = mvTA.CurrentPPDAmount
              Dim vPPAmount As String = ""
              If .PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptLoan Then vPPAmount = GetPageValue(CareServices.TraderPageType.tpPaymentPlanMaintenance, "Amount")
              'If the Amounts of the Payment Plan Details have been changed...
              If DoubleValue(vPPAmount) <> vAmount Then
                '...and the Amount has been set on the Payment Plan and the new Amounts on the Payment Plan Details do not equals the old Amounts then then set PaymentPlan.Amount to the new Amounts - the old Amounts
                If vPPAmount.Length > 0 AndAlso (vAmount <> DoubleValue(mvTA.PaymentPlan.OriginalPPDFixedAmount)) AndAlso (vAmount > DoubleValue(vPPAmount)) Then
                  vPPAmount = vAmount.ToString
                  SetPageValue(CareServices.TraderPageType.tpPaymentPlanMaintenance, "Amount", vPPAmount)
                End If
              End If
              Dim vArrears As Double = mvTA.CurrentPPDArrears
              If mvTA.PaymentPlan.Arrears <> vArrears Then      'Arrears are the same?           'mvArrears
                If ShowQuestion(QuestionMessages.QmPPArrearsToPPDArrears, MessageBoxButtons.YesNo, mvTA.PaymentPlan.Arrears.ToString, vArrears.ToString) = System.Windows.Forms.DialogResult.Yes Then 'mvArrears    'Payment Plan Arrears: {0} does not match the Details Arrears total: {1}\r\n\r\nDo you want to change the Payment Plan Arrears to {1}?
                  'The s/w used to attempt to find an arrears field on the PPM page and would only set the
                  'payment plan arrears if the field was found.  However, there has never been an arrears
                  'field on the PPM page, so that code has been removed.
                  pList("Arrears") = vArrears.ToString                  'mvArrears
                Else
                  vValid = False
                End If
              End If
            End If
          End If
        End If
      End With
    End If
    Return vValid
  End Function

  Private Sub AddPPALines(ByVal pList As ParameterList, ByRef pType As CareServices.TraderProcessDataTypes)
    Dim vUpdateAll As Boolean = False
    Dim vColl As New CollectionList(Of String)
    'Add Purchase Order Payment Lines
    If mvTA.PPADataSet.Tables.Contains("DataRow") Then
      Dim vCurrentRow As DataRow = mvTA.PPADataSet.Tables("DataRow").Rows(mvPPADGR.ActiveRow)
      Dim vMsg As String = ""
      If mvTA.POPercentage Then
        If DoubleValue(pList("Percentage")) = 0 Then vMsg = InformationMessages.ImPPAPaymentPercentageZero
      Else
        If DoubleValue(pList("Amount")) = 0 Then vMsg = InformationMessages.ImPPAPaymentAmountZero
      End If

      If pList.ContainsKey("Checkbox") AndAlso pList("Checkbox") = "Y" Then
        If ShowQuestion(QuestionMessages.QmUpdatePopPaymentLines, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          Dim vDataTable As DataTable = mvTA.PPADataSet.Tables("DataRow")
          Dim vIndex As Integer = 0
          For Each vRow As DataRow In vDataTable.Rows
            If vRow("AuthorisationStatus").ToString.Length = 0 Then
              vRow("PoPaymentType") = pList("PoPaymentType")
              vRow("DistributionCode") = pList("DistributionCode")
              vRow("NominalAccount") = pList("NominalAccount")
              vRow("SeparatePayment") = pList("SeparatePayment")
            End If
          Next
        End If
        If mvCurrentPage.EditPanel.PanelInfo.PanelItems.Exists("Checkbox") Then mvCurrentPage.EditPanel.FindCheckBox("Checkbox").Checked = False
      End If

      'Check all the payment schedule has PurchaseOrderType specified
      'when it is mandatory
      If pList.ContainsKey("POD_PurchaseOrderType") Then
        Dim vParams As New ParameterList(True)
        vParams("PurchaseOrderType") = pList("POD_PurchaseOrderType")

        Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPurchaseOrderTypes, vParams)
        If vDT IsNot Nothing AndAlso vDT.Columns.Contains("RequiresPoPaymentType") Then
          If vDT.Rows(0)("RequiresPoPaymentType").ToString = "Y" Then
            Dim vDataTable As DataTable = mvTA.PPADataSet.Tables("DataRow")
            Dim vIndex As Integer = 0
            For Each vRow As DataRow In vDataTable.Rows
              If (vRow("AuthorisationStatus").ToString.Length = 0 OrElse BooleanValue(vRow("ReadyForPayment").ToString) = False) Then
                If vRow("PoPaymentType").ToString.Length = 0 Then
                  vMsg = InformationMessages.ImPaymentTypeMandatory
                  cmdNext.Enabled = True
                  Exit For
                End If
              End If
            Next
          End If
        End If
      End If

      Dim vRowNumber As Integer = mvPPADGR.ActiveRow + 1
      Dim vTotal As Double
      Dim vNoDefaultCAMessage As New StringBuilder

      If vMsg.Length = 0 Then
        vCurrentRow("DueDate") = pList("DueDate")
        vCurrentRow("LatestExpectedDate") = pList("LatestExpectedDate")
        Dim vOldAmount As String = vCurrentRow("Amount").ToString
        vCurrentRow("Amount") = pList("Amount")
        vCurrentRow("Percentage") = pList("Percentage")
        vCurrentRow("AuthorisationRequired") = pList("AuthorisationRequired")
        If pList.ContainsKey("PoPaymentType") Then vCurrentRow("PoPaymentType") = pList("PoPaymentType")
        If pList.ContainsKey("DistributionCode") Then vCurrentRow("DistributionCode") = pList("DistributionCode")
        If pList.ContainsKey("NominalAccount") Then vCurrentRow("NominalAccount") = pList("NominalAccount")
        If pList.ContainsKey("SeparatePayment") Then vCurrentRow("SeparatePayment") = pList("SeparatePayment")
        If vRowNumber = mvPPADGR.MaxGridRows Then
          Dim vTable As DataTable = mvTA.PPADataSet.Tables("DataRow")
          Dim vRowIndex As Integer = 0
          For Each vRow As DataRow In vTable.Rows
            If mvTA.POPercentage Then
              vTotal += DoubleValue(vRow("Percentage").ToString)
            Else
              vTotal += DoubleValue(vRow("Amount").ToString)
            End If
            If (vRow("AuthorisationStatus").ToString.Length = 0 OrElse BooleanValue(vRow("ReadyForPayment").ToString) = False) AndAlso mvTA.PPADataSet.Tables("DataRow").Columns.Contains("PayByBacs") AndAlso BooleanValue(vRow("PayByBacs").ToString) OrElse
              (mvTA.PPADataSet.Tables("DataRow").Columns.Contains("PopPaymentMethod") AndAlso mvPPADGR.GetValue(vRowIndex, "PopPaymentMethod").Length > 0 AndAlso ValidateBankDetails(mvPPADGR.GetValue(vRowIndex, "PopPaymentMethod"))) Then
              If Not ValidateDefaultContactAccount(vRowIndex) Then
                If vNoDefaultCAMessage.Length = 0 Then vNoDefaultCAMessage.Append(InformationMessages.ImPPAMultiplePayByBacsNoDefaultBankAccount)
                If Not vColl.ContainsKey(mvPPADGR.GetValue(vRowIndex, "PayeeContactNumber")) Then
                  vColl.Add(mvPPADGR.GetValue(vRowIndex, "PayeeContactNumber"), mvPPADGR.GetValue(vRowIndex, "ContactName"))
                End If
              End If
            End If
            vRowIndex += 1
          Next
          If vNoDefaultCAMessage.Length = 0 Then
            vTotal = FixTwoPlaces(vTotal)
            If mvTA.POPercentage Then
              If vTotal <> 100 Then vMsg = InformationMessages.ImPPABalancePercentageNotMatch
            Else
              If mvTA.PPBalance <> vTotal Then
                If CanAmendPurchaseOrderAmount Then
                  vMsg = InformationMessages.ImPPABalanceNotMatch
                Else
                  If ShowQuestion(String.Format(QuestionMessages.QmConfirmPORegularPaymentChange, vOldAmount, pList("Amount").ToString), MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                    If mvTA.OldPORegularPaymentAmount.HasValue Then
                      If mvTA.OldPORegularPaymentAmount = DoubleValue(pList("Amount").ToString) Then mvTA.OldPORegularPaymentAmount = Nothing
                    Else
                      'Only set this once
                      mvTA.OldPORegularPaymentAmount = DoubleValue(vOldAmount)
                    End If
                    'Now update the trader values so that the user can see the updated values in every page
                    mvTA.PPBalance = vTotal
                    mvTA.CurrentPPDLineTotal = vTotal
                    mvTraderPages(CareNetServices.TraderPageType.tpPurchaseOrderSummary.ToString).EditPanel.SetValue("POBalance", mvTA.PPBalance.ToString("0.00"))
                    mvTraderPages(CareNetServices.TraderPageType.tpPurchaseOrderSummary.ToString).EditPanel.SetValue("POSTotal", mvTA.CurrentPPDLineTotal.ToString("0.00"))
                    Dim vPODetail As DataRow = mvTA.POSDataSet.Tables("DataRow").Rows(0)
                    vPODetail("Amount") = DoubleValue(vPODetail("Amount").ToString) - DoubleValue(vPODetail("Balance").ToString) + DoubleValue(vCurrentRow("Amount").ToString)
                    vPODetail("Balance") = vCurrentRow("Amount")
                    SetPageValue(CareNetServices.TraderPageType.tpPurchaseOrderDetails, "Amount", mvTA.CurrentPPDLineTotal.ToString("0.00"))
                  Else
                    vCurrentRow("Amount") = vOldAmount  'Reset the amount and do nothing
                    Exit Sub
                  End If
                End If
              End If
            End If
          End If
          If vMsg.Length = 0 AndAlso vNoDefaultCAMessage.Length = 0 Then
            Dim vTT As DataTable = vTable.DefaultView.ToTable
            For Each vRow As DataRow In vTT.Rows
              vRow("ContactName") = ""
              pList.ObjectValue("PPALine" & vRow("PaymentNumber").ToString) = vRow
            Next

            cmdNext.Enabled = False
            cmdFinished.Enabled = True
            vTable.DefaultView.ApplyDefaultSort = True
          End If
          vRowNumber -= 1
        End If
      Else
        vRowNumber -= 1 'Don't change the selected row as the current row has an error.
      End If
      mvPPADGR.SelectRow(vRowNumber, True)
      SetPPAEditable(mvPPADGR.ActiveRow)

      If vMsg.Length > 0 OrElse vNoDefaultCAMessage.Length > 0 Then
        If vMsg.Length > 0 Then
          ShowInformationMessage(vMsg, vTotal.ToString, mvTA.PPBalance.ToString)
        Else
          vNoDefaultCAMessage.AppendLine()
          vNoDefaultCAMessage.AppendLine()
          For vIndex As Integer = 0 To vColl.Count - 1
            vNoDefaultCAMessage.AppendLine(vColl.ItemKey(vIndex).ToString & " (" & vColl(vIndex).ToString & ")")
          Next
          ShowErrorMessage(vNoDefaultCAMessage.ToString)
        End If
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then pType = CareServices.TraderProcessDataTypes.tpdtNextPage
      End If
    End If
  End Sub

  Private Sub AddGiftAid()
    Dim vNewList As New ParameterList(True)
    If mvCurrentPage.EditPanel.AddValuesToList(vNewList) Then
      Dim vDonation As Boolean = mvCurrentPage.EditPanel.FindCheckBox("DeclarationType").Checked
      Dim vMembers As Boolean = mvCurrentPage.EditPanel.FindCheckBox("DeclarationType2").Checked
      vNewList("DeclarationType") = "M"
      If vDonation AndAlso vMembers Then
        vNewList("DeclarationType") = "A"
      ElseIf vDonation Then
        vNewList("DeclarationType") = "D"
      End If
      If vNewList.Contains("DeclarationType2") Then vNewList.Remove("DeclarationType2")
      If mvNonFinancialBatchNumber > 0 Then
        vNewList.IntegerValue("NonFinancialBatchNumber") = mvNonFinancialBatchNumber
        vNewList.IntegerValue("NonFinancialTransactionNumber") = mvNonFinancialTransactionNumber
      End If
      mvTA.DeclarationNumber = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctGiftAidDeclarations, vNewList).IntegerValue("DeclarationNumber")
      Dim vAdhocMessage As String = AppValues.OutstandingMessage
      If Not String.IsNullOrWhiteSpace(vAdhocMessage) Then
        ShowQuestion(vAdhocMessage, MessageBoxButtons.OK)
      End If
    End If
  End Sub

  Private Sub AddEventBookingLines(ByVal pList As ParameterList)
    If mvTA.EventBookingDataSet.Tables.Count > 0 AndAlso mvTA.EventBookingDataSet.Tables.Contains("DataRow") Then
      Dim vDT As DataTable = DataHelper.GetTableFromDataSet(mvTA.EventBookingDataSet)
      If vDT IsNot Nothing Then
        Dim vLineNumber As Integer
        Dim vInterestOnly As Boolean
        If pList.ContainsKey("InterestOnly") Then vInterestOnly = BooleanValue(pList("InterestOnly"))
        For Each vRow As DataRow In vDT.Rows
          vLineNumber += 1
          If vInterestOnly Then
            vRow("Amount") = "0.00"
            vRow("VatAmount") = "0.00"
          End If
          pList.ObjectValue("EventBookingLine" & vLineNumber.ToString) = vRow
        Next
      End If
    End If
  End Sub

  Private Sub AddExamBookingLines(ByVal pList As ParameterList)
    If mvTA.ExamBookingDataSet.Tables.Count > 0 AndAlso mvTA.ExamBookingDataSet.Tables.Contains("DataRow") Then
      Dim vDT As DataTable = DataHelper.GetTableFromDataSet(mvTA.ExamBookingDataSet)
      If vDT IsNot Nothing Then
        Dim vUseAlternateRate As Boolean
        Dim vPrimaryRate As String = ""
        Dim vAlternateRate As String = ""
        For Each vRow As DataRow In vDT.Rows
          If vRow.Item("AlternateRate").ToString.Length > 0 Then
            vPrimaryRate = vRow.Item("RateDesc").ToString
            vAlternateRate = vRow.Item("AlternateRateDesc").ToString
            vUseAlternateRate = True
            Exit For
          End If
        Next
        If vUseAlternateRate Then
          If ShowQuestion(QuestionMessages.QmCandidateMaxAttempts, MessageBoxButtons.YesNo, vPrimaryRate, vAlternateRate) = vbYes Then
            For Each vRow As DataRow In vDT.Rows
              If vRow.Item("AlternateRate").ToString.Length > 0 Then
                Dim vAmount As Double = DoubleValue(pList("Amount").ToString)
                vAmount -= DoubleValue(vRow.Item("Amount").ToString)
                vAmount += DoubleValue(vRow.Item("AlternateAmount").ToString)
                pList("Amount") = vAmount.ToString
                vRow.Item("Rate") = vRow.Item("AlternateRate")
                vRow.Item("Amount") = vRow.Item("AlternateAmount")
                vRow.Item("VATAmount") = vRow.Item("AlternateVATAmount")
              End If
            Next
          End If
        End If
        Dim vLineNumber As Integer
        For Each vRow As DataRow In vDT.Rows
          vLineNumber += 1
          pList.ObjectValue("ExamBookingLine" & vLineNumber.ToString) = vRow
        Next
      End If
    End If
  End Sub

  Private Sub AddCMTLines(ByVal pList As ParameterList)
    Dim vTable As DataTable = mvTA.CMTOldPPDDataSet.Tables("DataRow")
    If Not vTable Is Nothing Then
      vTable.DefaultView.Sort = "DetailNumber DESC"
      Dim vTT As DataTable = vTable.DefaultView.ToTable
      For Each vRow As DataRow In vTT.Rows
        pList.ObjectValue("OldCMTPPDLine" & vRow("DetailNumber").ToString) = vRow
      Next
    End If
    vTable = mvTA.CMTNewPPDDataSet.Tables("DataRow")
    If vTable IsNot Nothing Then
      vTable.DefaultView.Sort = "DetailNumber DESC"
      Dim vTT As DataTable = vTable.DefaultView.ToTable
      For Each vRow As DataRow In vTT.Rows
        pList.ObjectValue("PPDLine" & vRow("DetailNumber").ToString) = vRow
      Next
    End If
  End Sub

  ''' <summary>Perform any additional processing required for CMT.</summary>
  Private Function AddAdditionalCMTData(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True

    If BooleanValue(mvTA.PaymentPlan.DirectDebitStatus) Then
      If mvTA.Pages.ContainsKey(CareNetServices.TraderPageType.tpChangeMembershipType.ToString) Then
        Dim vCMTEPL As EditPanel = mvTA.Pages(CareNetServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel
        Dim vMTRow As DataRow = vCMTEPL.FindTextLookupBox("MembershipType").GetDataRow
        If vMTRow.Item("MembersPerOrder").ToString = "1" Then
          Dim vMembershipTypeCode As String = vCMTEPL.GetValue("MembershipType")
          Dim vMembershipNumber As Integer = 0
          If mvTA.MembersDataSet.Tables(0).Rows.Count >= 1 Then
            For Each vMemberRow As DataRow In DataHelper.GetTableFromDataSet(mvTA.MembersDataSet).Rows
              If vMemberRow.Item("MembershipType").ToString.Equals(vMembershipTypeCode, System.StringComparison.CurrentCultureIgnoreCase) Then
                vMembershipNumber = IntegerValue(vMemberRow.Item("MembershipNumber").ToString)
              End If
              If vMembershipNumber > 0 Then Exit For
            Next
          End If

          Dim vDDCheckList As New ParameterList(True, True)
          vDDCheckList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
          vDDCheckList.IntegerValue("MembershipNumber") = vMembershipNumber
          vDDCheckList("CancellationReason") = vCMTEPL.GetValue("CancellationReason")
          vDDCheckList("MembershipType") = vMembershipTypeCode
          Dim vDDReturnList As ParameterList = DataHelper.CanChangeDDPayer(vDDCheckList)
          If BooleanValue(vDDReturnList.ValueIfSet("CanChangeDDPayer")) Then
            mvTA.Pages(CareNetServices.TraderPageType.tpMembershipPayer.ToString).EditPanel.EnableControlList("ContactNumber,AddressNumber", False)
            Dim vDDResult As DialogResult = ShowQuestion(QuestionMessages.QmCancelMemberMoveDD, MessageBoxButtons.YesNoCancel, vDDReturnList("DirectDebitPayerName"), vDDReturnList("DirectDebitNewPayerName"))
            If vDDResult = System.Windows.Forms.DialogResult.Cancel Then
              vValid = False    'User cancelled so don't continue
            Else
              pList("ChangeDDPayer") = If(vDDResult = System.Windows.Forms.DialogResult.Yes, "Y", "N")
              If vDDResult = System.Windows.Forms.DialogResult.Yes Then pList("DirectDebitNewPayerContactNumber") = vDDReturnList("DirectDebitNewPayerContactNumber")
            End If
          End If
        End If
        If pList.ContainsKey("ChangeDDPayer") = False Then pList("ChangeDDPayer") = "N" 'Set to 'N' so that server know this has been checked.
      Else
        'Can't find CMT page so cannot continue
        vValid = False
      End If
    End If

    Return vValid

  End Function

  Private Function CheckTransactionTotal() As Boolean
    Dim vResult As System.Windows.Forms.DialogResult
    Dim vReturn As Boolean = True

    'to check that the totals match when the user clicks finished on the TAS Page
    If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment Then 'Or mvFinancialAdjustment = atMove Then
      ShowInformationMessage(InformationMessages.ImCannotFinishAdjustment, mvTA.CurrentLineTotal.ToString("0.00"), mvTA.TransactionAmount.ToString("0.00"))
      vResult = System.Windows.Forms.DialogResult.Cancel
    Else
      If (mvTA.TransactionPaymentMethod = "CQIN" OrElse mvTA.TransactionPaymentMethod = "CCIN") Then
        Dim vShowUnderOverPayment As Boolean = True
        If mvTA.TransactionAmount = 0 Then
          vResult = ShowQuestion(QuestionMessages.QmTransactionInvoicePayAmountNotSet, MessageBoxButtons.YesNo, mvTA.CurrentLineTotal.ToString("0.00"))
          If vResult = System.Windows.Forms.DialogResult.No Then
            vResult = System.Windows.Forms.DialogResult.Cancel
          Else
            vShowUnderOverPayment = False
          End If
        ElseIf mvTA.CurrentLineTotal > mvTA.TransactionAmount Then
          vResult = ShowQuestion(QuestionMessages.QmTransactionInvoiceUnderPayAmount, MessageBoxButtons.YesNo, mvTA.TransactionAmount.ToString("0.00"), mvTA.CurrentLineTotal.ToString("0.00"), ((mvTA.CurrentLineTotal - mvTA.TransactionAmount) * -1).ToString("0.00"))
          If vResult = System.Windows.Forms.DialogResult.Yes AndAlso AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_under_payment_product).Length = 0 Then
            ShowErrorMessage(InformationMessages.ImInvoiceUnderPayProductNotSet)
          End If
        Else
          vResult = ShowQuestion(QuestionMessages.QmTransactionInvoiceOverPayAmount, MessageBoxButtons.YesNo, mvTA.TransactionAmount.ToString("0.00"), mvTA.CurrentLineTotal.ToString("0.00"), (mvTA.TransactionAmount - mvTA.CurrentLineTotal).ToString("0.00"))
          If vResult = System.Windows.Forms.DialogResult.Yes AndAlso AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_over_payment_product).Length = 0 Then
            ShowErrorMessage(InformationMessages.ImInvoiceOverPayProductNotSet)
          End If
        End If
        If vShowUnderOverPayment AndAlso vResult = System.Windows.Forms.DialogResult.Yes Then
          SetPage(CareNetServices.TraderPageType.tpTransactionAnalysis)
          mvTA.TransactionType = "SALE"
          Dim vErrorNumber As CDBNETCL.CareException.ErrorNumbers = Nothing
          ProcessData(CareNetServices.TraderProcessDataTypes.tpdtNextPage, False, vErrorNumber)
          If vErrorNumber <> CareException.ErrorNumbers.enPageNotFound Then
            If mvTA.CurrentLineTotal > mvTA.TransactionAmount Then
              mvCurrentPage.EditPanel.SetValue("Product", AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_under_payment_product))
              mvCurrentPage.EditPanel.SetValue("Rate", AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_under_payment_rate))
              mvCurrentPage.EditPanel.SetValue("Amount", ((mvTA.CurrentLineTotal - mvTA.TransactionAmount) * -1).ToString)
            Else
              mvCurrentPage.EditPanel.SetValue("Product", AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_over_payment_product))
              mvCurrentPage.EditPanel.SetValue("Rate", AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.invoice_over_payment_rate))
              mvCurrentPage.EditPanel.SetValue("Amount", (mvTA.TransactionAmount - mvTA.CurrentLineTotal).ToString)
            End If
            ProcessData(CareNetServices.TraderProcessDataTypes.tpdtNextPage, True)  'Show the summary page
          End If
        End If
        If vShowUnderOverPayment Then vResult = System.Windows.Forms.DialogResult.Cancel 'Always stay on the page as the user needs to see what item has been item
      ElseIf mvTA.CreatesTransaction AndAlso (Not mvTA.PayPlanPayMethod OrElse mvTPPDone) Then
        If mvTA.AutoSetAmount Then
          vResult = System.Windows.Forms.DialogResult.Yes
        Else
          If mvTA.TransactionAmount = 0 Then
            vResult = ShowQuestion(QuestionMessages.QmTransactionAmountNotSet, MessageBoxButtons.YesNoCancel, mvTA.CurrentLineTotal.ToString("0.00"))
          Else
            vResult = System.Windows.Forms.DialogResult.No
            Dim vTotalDiff As Double = mvTA.TransactionAmount - mvTA.CurrentLineTotal
            If mvTA.DonationProduct.Length > 0 AndAlso vTotalDiff > 0 Then 'need to check for existance of donation product here
              vResult = ShowQuestion(QuestionMessages.QmTransactionAmountAddDonation, MessageBoxButtons.YesNoCancel, mvTA.TransactionAmount.ToString("0.00"), mvTA.CurrentLineTotal.ToString("0.00"), vTotalDiff.ToString("0.00")) 'The Transaction Amount {0} does not match the current line total {1} & vbCrLf & vbCrLf & Do you want to add a donation amount of {2}?
              If vResult = vbYes Then
                'we need to add the donation for the amount of vTotalDiff (money from the transaction that's not allocated yet)
                mvTA.TransactionDonationAmount = vTotalDiff
              End If
            End If
            If vResult = System.Windows.Forms.DialogResult.No Then
              vResult = ShowQuestion(QuestionMessages.QmTransactionAmountNotMatch, MessageBoxButtons.YesNoCancel, mvTA.TransactionAmount.ToString("0.00"), mvTA.CurrentLineTotal.ToString("0.00"))
            End If
          End If
        End If
      End If
    End If
    If vResult = vbCancel Then
      vReturn = False
    ElseIf vResult = vbYes Then
      mvTA.TransactionAmount = mvTA.CurrentLineTotal
      Dim vPage As TraderPage = mvTraderPages(CareServices.TraderPageType.tpTransactionDetails.ToString)
      vPage.EditPanel.SetValue("Amount", mvTA.TransactionAmount.ToString("0.00"))
    End If
    Return vReturn
  End Function

  Private Function GetAdditionalValues(ByVal pList As ParameterList, ByVal pType As CareServices.TraderProcessDataTypes) As Boolean
    Dim vValid As Boolean = True
    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpAccommodationBooking
        If pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso pType <> CareServices.TraderProcessDataTypes.tpdtPreviousPage Then vValid = ValidateAccommodationBooking(pList)
      Case CareServices.TraderPageType.tpEventBooking
        If pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso pType <> CareServices.TraderProcessDataTypes.tpdtPreviousPage Then vValid = ValidateEventBooking(pList)
      Case CareServices.TraderPageType.tpExamBooking
        If pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso pType <> CareServices.TraderProcessDataTypes.tpdtPreviousPage Then vValid = ValidateExamBooking(pList)
      Case CareServices.TraderPageType.tpInvoicePayments
        pList("Company") = mvTA.CACompany
        pList("InvoicePaymentAmount") = mvCurrentPage.EditPanel.GetValue("CurrentPayment")
        pList("CurrentUnAllocated") = mvCurrentPage.EditPanel.GetValue("CurrentUnAllocated")
      Case CareServices.TraderPageType.tpOutstandingScheduledPayments
        pList("AmountOutstanding") = mvCurrentPage.EditPanel.GetValue("AmountOutstanding")
        If (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment) Then pList.IntegerValue("LineNumber") = mvTA.EditLineNumber
        If mvTraderPages.ContainsKey(CareServices.TraderPageType.tpPayments.ToString) Then
          Dim vPage As TraderPage = mvTraderPages(CareServices.TraderPageType.tpPayments.ToString)
          If vPage IsNot Nothing Then
            Return vPage.EditPanel.AddValuesToList(pList, True, EditPanel.AddNullValueTypes.anvtAll)
          End If
        End If
      Case CareServices.TraderPageType.tpPaymentPlanProducts, CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
        pList.IntegerValue("PPDProductNumbersCount") = mvTA.PPDProductNumbersCount()
        If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanProducts AndAlso mvCurrentPPDLine >= 0 AndAlso mvTA.EditLineNumber > 0 Then
          Dim vRow As DataRow = mvTA.GetDataSetLine(mvTA.PPDDataSet, mvCurrentPPDLine)
          If vRow IsNot Nothing Then
            If vRow.Table.Columns.Contains("IncentiveLineType") Then
              Dim vIncentiveLineType As String = vRow.Item("IncentiveLineType").ToString
              pList("IncentiveLineType") = vIncentiveLineType
              pList("IncentiveIgnoreProductAndRate") = vRow.Item("IncentiveIgnoreProductAndRate").ToString
              If vIncentiveLineType.Length > 1 Then
                If vIncentiveLineType.Substring(1, 1).ToUpper = "I" Then pList("Quantity") = vRow.Item("Quantity").ToString 'Use original quantity as it is the incentive period (current page value is 1 & disabled)
              End If
            End If
          End If
        ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
          If mvTA.PPDDataSet.Tables("DataRow") IsNot Nothing Then
            Dim vTable As DataTable = mvTA.PPDDataSet.Tables("DataRow")
            If vTable.Columns.Contains("CreatedBy") Then
              Dim vDetailNumber As Integer = mvTA.EditLineNumber
              'keep createdby and createdon the same for the amended line
              For Each vExistingRow As DataRow In vTable.Rows
                If IntegerValue(vExistingRow("DetailNumber").ToString) = vDetailNumber AndAlso pList.ContainsKey("CreatedBy") = False Then
                  pList.Add("CreatedBy", vExistingRow("CreatedBy").ToString)
                  pList.Add("CreatedOn", vExistingRow("CreatedOn").ToString)
                End If
              Next
            End If
          End If
        End If
        If mvTA.PaymentPlanDetailsPricing.GotPricingData Then mvTA.PaymentPlanDetailsPricing.GetDataAsParameterlist(pList)
      Case CareServices.TraderPageType.tpTransactionDetails
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_fp_prevent_future_date) AndAlso
           Date.Compare(CDate(mvCurrentPage.EditPanel.GetValue("TransactionDate")), Today) > 0 Then
          mvCurrentPage.EditPanel.SetErrorField("TransactionDate", InformationMessages.ImTransactionDateInFuture)
          vValid = False
        End If
      Case CareServices.TraderPageType.tpPurchaseOrderCancellation
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then vValid = ShowQuestion(QuestionMessages.QmConfirmPOCancellation, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes
      Case CareServices.TraderPageType.tpConfirmProvisionalTransactions
        pList("BatchNumber") = mvCurrentPage.EditPanel.GetValue("ProvisionalBatchNumber")
        pList("TransactionNumber") = mvCurrentPage.EditPanel.GetValue("ProvisionalTransNumber")
      Case CareServices.TraderPageType.tpBatchInvoiceSummary
      Case CareNetServices.TraderPageType.tpServiceBooking
        pList("SBGrossQty") = mvTA.SBGrossQty.ToString("0.00")
        pList("SBGrossAmount") = mvTA.SBGrossAmount.ToString("0.00")
        pList("SBEntitlementQty") = mvTA.SBEntitlementQty.ToString("0.00")
        pList("VATExclusive") = CBoolYN(mvTA.LinePriceVATEx)
        pList("ContactDiscount") = CBoolYN(mvTA.PayerHasDiscount)
      Case CareNetServices.TraderPageType.tpStatementList
        If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
          If mvStatementGrid IsNot Nothing AndAlso mvStatementGrid.RowCount > 0 Then
            With mvStatementGrid
              mvTA.BankTransactionLineNumber = IntegerValue(.GetValue(.CurrentRow, "LineNumber"))
              mvTA.PayersSortCode = .GetValue(.CurrentRow, "PayersSortCode")
              mvTA.PayersAccountNumber = .GetValue(.CurrentRow, "PayersAccountNumber")
              mvTA.PayersName = .GetValue(.CurrentRow, "PayersName")
              mvTA.ReferenceNumber = .GetValue(.CurrentRow, "ReferenceNumber")
              mvTA.OriginalTransactionAmount = DoubleValue(.GetValue(.CurrentRow, "Amount"))
            End With
          Else
            vValid = False
          End If
        End If
      Case CareNetServices.TraderPageType.tpPaymentMethod2, CareNetServices.TraderPageType.tpTransactionAnalysisSummary
        pList("TPPDone") = CBoolYN(mvTPPDone)
    End Select

    If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
      'Required to initialise and update the bank transaction
      pList("StatementDate") = mvTA.StatementDate
      pList("BankTransactionLineNumber") = mvTA.BankTransactionLineNumber.ToString
      pList("PayersSortCode") = mvTA.PayersSortCode
      pList("PayersAccountNumber") = mvTA.PayersAccountNumber
      pList("PayersName") = mvTA.PayersName
      pList("ReferenceNumber") = mvTA.ReferenceNumber
      pList("OriginalAmount") = mvTA.OriginalTransactionAmount.ToString("0.00")
      pList("BankPaymentMethod") = mvTA.BankPaymentMethod
      pList("Notes") = mvTA.TransactionNote
    End If

    Select Case mvTA.TransactionType
      Case "CRDN"
        pList("SalesLedgerAccount") = GetPageValue(CareServices.TraderPageType.tpCreditCustomer, "SalesLedgerAccount")
      Case "PAYM"
        If Not pList.Contains("MemberNumber") Then pList("MemberNumber") = mvTA.MemberNumber
        If Not pList.Contains("CovenantNumber") Then pList.IntegerValue("CovenantNumber") = mvTA.CovenantNumber
        If mvTA.PaymentPlan IsNot Nothing Then
          If Not pList.Contains("PaymentPlanNumber") OrElse (pList.Contains("PaymentPlanNumber") AndAlso pList("PaymentPlanNumber").Length = 0) Then pList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
        Else
        End If
      Case "APAY"
        If Not pList.Contains("PaymentPlanNumber") AndAlso mvPPNumber > 0 Then pList.IntegerValue("PaymentPlanNumber") = mvPPNumber
        If Not pList.Contains("NonFinancialBatchNumber") AndAlso mvNonFinancialBatchNumber > 0 Then
          pList.IntegerValue("NonFinancialBatchNumber") = mvNonFinancialBatchNumber
          pList.IntegerValue("NonFinancialTransactionNumber") = mvNonFinancialTransactionNumber
        End If
      Case "MEMB"
        If Not pList.Contains("NonFinancialBatchNumber") AndAlso mvNonFinancialBatchNumber > 0 Then
          pList.IntegerValue("NonFinancialBatchNumber") = mvNonFinancialBatchNumber
          pList.IntegerValue("NonFinancialTransactionNumber") = mvNonFinancialTransactionNumber
        End If
      Case "DONS", "SALE"
        If mvTA.LinkToFundraisingPayments AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso
           (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished) _
           AndAlso AppValues.ControlValue(AppValues.ControlValues.fundraising_payment_type).Length > 0 Then
          If (mvTA.TransactionType = "SALE" AndAlso mvTA.GiftInKind) OrElse mvTA.TransactionType = "DONS" Then
            Dim vList As New ParameterList(True)
            vList("ContactNumber") = pList("ContactNumber")
            If Not mvTA.GiftInKind Then vList("FundraisingPaymentType") = AppValues.ControlValue(AppValues.ControlValues.fundraising_payment_type)
            vList.IntegerValue("TraderApplication") = mvTA.ApplicationNumber
            Dim vNumber As Integer
            If ShowQuestion(QuestionMessages.QmLinkToFundraisingPayment, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vNumber = FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftFundraisingPaymentScheduleFinder, vList, Me)
              If vNumber > 0 Then pList.IntegerValue("ScheduledPaymentNumber") = vNumber
            End If
            mvTA.ScheduledPaymentNumber = vNumber
            If vNumber = 0 AndAlso pList.Contains("ScheduledPaymentNumber") Then
              pList.Remove("ScheduledPaymentNumber")
              mvTA.ScheduledPaymentNumber = 0
            End If
          End If
        End If
    End Select
    If mvTA.PaymentPlan IsNot Nothing AndAlso (mvTA.ApplicationType = ApplicationTypes.atMaintenance _
      OrElse mvTA.ChangeMembershipType = True OrElse mvTA.ApplicationType = ApplicationTypes.atConversion) Then
      pList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
    End If

    If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
      'on finished we just need to know if we created any payment plans
      pList("PaymentPlanCreated") = CBoolYN(mvTA.PPNumbersCreated.Count > 0)
      If mvTA.NonFinancialBatch Then
        Dim vCreateMailingDocument As Boolean
        If mvTA.DeclarationNumber > 0 Then
          vCreateMailingDocument = True
        Else
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(mvTA.AnalysisDataSet)
          If vTable IsNot Nothing AndAlso vTable.Columns.Contains("TraderLineType") Then
            For Each vRow As DataRow In vTable.Rows
              If vRow("TraderLineType").ToString = "GP" Then
                vCreateMailingDocument = True
                Exit For
              End If
            Next
          End If
        End If
        If vCreateMailingDocument Then pList("CreateMailingDocument") = "Y"
      End If
    Else
      ' on any other button we need to know if this particular PP has been created within this transaction
      If mvTA.PaymentPlan IsNot Nothing Then
        If mvTA.PPNumbersCreated.ContainsKey(mvTA.PaymentPlan.PaymentPlanNumber.ToString) Then
          pList("PaymentPlanCreated") = "Y"
        Else
          pList("PaymentPlanCreated") = "N"
        End If
      End If
      If pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso mvTA.PPNumbersCreated.Count > 0 Then
        Dim vPPNumbersCreated As String = ""
        For vPPCreated As Integer = 0 To mvTA.PPNumbersCreated.Count - 1
          Dim vPPNumber As Integer = mvTA.PPNumbersCreated.Item(vPPCreated)
          If vPPNumbersCreated.Length = 0 Then
            vPPNumbersCreated = vPPNumber.ToString
          Else
            vPPNumbersCreated = String.Concat(vPPNumbersCreated, ",", vPPNumber)
          End If
        Next
        If vPPNumbersCreated.Length > 0 Then
          pList.Add("PaymentPlansToDelete", vPPNumbersCreated)
        End If
      ElseIf pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpGiftAidDeclaration OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpAddressMaintenance) Then
        pList.Add("NonFinancialBatchNumber", mvNonFinancialBatchNumber)
        pList.Add("NonFinancialTransactionNumber", mvNonFinancialTransactionNumber)
      End If
    End If

    If Not (mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanProducts OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance) Then
      mvTA.ClearPaymentPlanDetailsPricing()
    End If

    Return vValid
  End Function

  Private Sub GetEditValues(ByVal pType As CareServices.TraderProcessDataTypes, ByVal pList As ParameterList)
    'Find the DateRow to be edited etc.
    If (pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Or pType = CareServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine) Then
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          Dim vDataRow As DataRow = mvTA.GetDataSetLine(mvTA.PPDDataSet, mvCurrentPPDLine)
          pList.ObjectValue("PPDLine" & vDataRow("LineNumber").ToString) = vDataRow
          If pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then
            mvTA.EditLineNumber = IntegerValue(vDataRow("LineNumber").ToString)
            mvTA.PPDMemberOrPayer = vDataRow("MemberOrPayer").ToString
          End If
        Case CareServices.TraderPageType.tpTransactionAnalysisSummary
          'Clicked 'Edit' or 'Delete' from TAS page
          Dim vDataRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
          pList.ObjectValue("TraderAnalysisLine" & vDataRow("LineNumber").ToString) = vDataRow
          If pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then
            mvTA.EditLineNumber = IntegerValue(vDataRow("LineNumber").ToString)
            If IntegerValue(vDataRow("ScheduledPaymentNumber")) > 0 Then mvTA.ScheduledPaymentNumber = IntegerValue(vDataRow("ScheduledPaymentNumber"))
          End If
          If mvTA.EditExistingTransaction = False Then mvTA.TransactionLines -= 1
        Case CareServices.TraderPageType.tpPurchaseOrderSummary
          Dim vDataRow As DataRow = mvTA.GetDataSetLine(mvTA.POSDataSet, mvCurrentPPDLine)
          pList.ObjectValue("POSLine" & vDataRow("LineNumber").ToString) = vDataRow
          If pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then mvTA.EditLineNumber = IntegerValue(vDataRow("LineNumber").ToString)
        Case CareServices.TraderPageType.tpPurchaseInvoiceSummary
          Dim vDataRow As DataRow = mvTA.GetDataSetLine(mvTA.PISDataSet, mvCurrentPPDLine)
          pList.ObjectValue("PISLine" & vDataRow("LineNumber").ToString) = vDataRow
          If pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then mvTA.EditLineNumber = IntegerValue(vDataRow("LineNumber").ToString)
      End Select
    ElseIf pType = CareServices.TraderProcessDataTypes.tpdtNextPage Then
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpProductDetails, CareServices.TraderPageType.tpPayments, CareServices.TraderPageType.tpOutstandingScheduledPayments
          'Clicked 'Next from PRD page whilst editing an existing transaction
          If mvTA.EditLineNumber > 0 AndAlso (mvCurrentPage.PageType <> CareServices.TraderPageType.tpOutstandingScheduledPayments OrElse mvTA.FinancialAdjustment <> BatchInfo.AdjustmentTypes.Adjustment) Then
            'Editing an existing line
            pList.IntegerValue("TransactionLines") = mvTA.EditLineNumber - 1
          Else
            'Adding a new line
            pList.IntegerValue("TransactionLines") = mvTA.TransactionLines
          End If
          If mvTA.ScheduledPaymentNumber > 0 Then pList.IntegerValue("ScheduledPaymentNumber") = mvTA.ScheduledPaymentNumber
        Case CareServices.TraderPageType.tpPaymentPlanProducts, CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
          If mvTA.EditLineNumber > 0 Then
            'Editing an existing line
            pList.IntegerValue("PPDLines") = mvTA.EditLineNumber - 1
          Else
            'Adding a new line
            pList.IntegerValue("PPDLines") = mvTA.PPDLines
          End If
        Case CareServices.TraderPageType.tpPurchaseOrderProducts, CareServices.TraderPageType.tpPurchaseInvoiceProducts
          If mvTA.EditLineNumber > 0 Then pList.IntegerValue("LineNumber") = mvTA.EditLineNumber
      End Select
    ElseIf pType = CareServices.TraderProcessDataTypes.tpdtAmendMemberSummary Then
      Dim vDataRow As DataRow = mvTA.MembersDataSet.Tables("DataRow").Rows(mvCurrentRow)
      pList.ObjectValue("MemberLine" & vDataRow("LineNumber").ToString) = vDataRow
      mvTA.EditMemberLineNumber = IntegerValue(vDataRow("LineNumber").ToString)
    ElseIf pType = CareServices.TraderProcessDataTypes.tpdtDeletePaymentPlanLine Then
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpPaymentPlanSummary
          Dim vDataRow As DataRow = mvTA.GetDataSetLine(mvTA.PPDDataSet, mvCurrentPPDLine)
          pList.ObjectValue("PPDLine" & vDataRow("LineNumber").ToString) = vDataRow
      End Select
    End If

  End Sub


  Private Function CheckPOISummaryBalance(ByVal pList As ParameterList) As Boolean
    With mvTA
      If .PPBalance <> .CurrentPPDLineTotal Then      'Totals are the same?
        Dim vResult As System.Windows.Forms.DialogResult
        If .PPBalance = 0 Then
          vResult = ShowQuestion(QuestionMessages.QmPOSBalanceNotSet, MessageBoxButtons.YesNo, .CurrentPPDLineTotal.ToString)
        Else
          vResult = ShowQuestion(QuestionMessages.QmPOSBalanceNotMatch, MessageBoxButtons.YesNo, .PPBalance.ToString, .CurrentPPDLineTotal.ToString)
        End If
        If vResult = System.Windows.Forms.DialogResult.Yes Then
          If mvTA.ApplicationType = ApplicationTypes.atPurchaseOrder Then
            'Jira 664: Update Purchase Order Details Amount
            SetPageValue(CareNetServices.TraderPageType.tpPurchaseOrderDetails, "Amount", .CurrentPPDLineTotal.ToString)
            SetPageValue(CareServices.TraderPageType.tpPurchaseOrderSummary, "POBalance", .CurrentPPDLineTotal.ToString)
          Else
            SetPageValue(CareServices.TraderPageType.tpPurchaseInvoiceSummary, "PIBalance", .CurrentPPDLineTotal.ToString)
          End If
          mvTA.PPBalance = .CurrentPPDLineTotal
          pList("PPBalance") = mvTA.PPBalance.ToString
          Return True
        Else
          Return False
        End If
      Else
        Return True
      End If
    End With
  End Function
  Private Function GetIncentivesTable(ByVal pPaymentIncentive As Boolean, ByVal pSource As String, ByVal pAmount As Double, ByRef pCanAddEnclosure As Boolean) As DataTable
    If mvTA.IncentiveDataSet Is Nothing OrElse pPaymentIncentive Then
      Dim vSourceCode As String = String.Empty
      Dim vMemberPage As CareServices.TraderPageType = CareServices.TraderPageType.tpMembership
      If (mvTA.TransactionType = "MEMB" Or mvTA.TransactionType = "CMEM" Or mvTA.TransactionType = "MEMC") And mvTA.ConversionShowPPD = False Then
        If mvTA.TransactionType = "MEMC" Then vMemberPage = CareServices.TraderPageType.tpChangeMembershipType
        vSourceCode = GetPageValue(vMemberPage, "Source")
      Else
        vSourceCode = GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Source")
      End If
      If mvTA.ApplicationType = ApplicationTypes.atConversion Then
        'Get source code from auto payment method page
        Select Case mvCurrentPage.PageType
          Case CareNetServices.TraderPageType.tpCreditCardAuthority, CareNetServices.TraderPageType.tpDirectDebit, CareNetServices.TraderPageType.tpStandingOrder
            If FindControl(mvCurrentPage.EditPanel, "Source", False) IsNot Nothing Then vSourceCode = mvCurrentPage.EditPanel.GetValue("Source")
        End Select
      End If
      If Len(pSource) > 0 Then vSourceCode = pSource

      Dim vReason As String = ""
      Dim vDDReason As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.dd_reason)
      Dim vCCReason As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.cc_reason)
      Dim vSOReason As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.so_reason)

      If pPaymentIncentive Then
        vReason = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.payment_reason)
        If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpTransactionDetails, "Source")
      Else
        Select Case mvTA.TransactionType
          Case "MEMB", "CMEM"
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.reason_is_grade, True) Then
              vReason = GetPageValue(CareServices.TraderPageType.tpMembership, "MembershipType")
            Else
              vReason = AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.reason_for_despatch)
            End If
          Case Else
            Select Case mvTA.PPPaymentType
              Case "DIRD"
                vReason = vDDReason
              Case "CCCA"
                vReason = vCCReason
              Case "STDO"
                vReason = vSOReason
              Case Else
                vReason = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.o_reason)
            End Select
        End Select

        If mvTA.ApplicationType = ApplicationTypes.atConversion OrElse mvTA.ApplicationType = ApplicationTypes.atMaintenance OrElse mvTA.TransactionType = "MEMC" Then
          Select Case mvCurrentPage.PageType
            Case CareServices.TraderPageType.tpDirectDebit
              vReason = vDDReason
            Case CareServices.TraderPageType.tpCreditCardAuthority
              vReason = vCCReason
            Case CareServices.TraderPageType.tpStandingOrder
              vReason = vSOReason
            Case CareServices.TraderPageType.tpMembershipPayer
              If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.reason_is_grade, True) Then vReason = GetPageValue(vMemberPage, "MembershipType")
          End Select
        End If
      End If
      Dim vPayMethodReason As String = ""
      If vReason <> vCCReason And vReason <> vDDReason And vReason <> vSOReason Then
        If mvTA.PPPaymentType = "DIRD" OrElse mvTA.PPPaymentType = "CCCA" OrElse mvTA.PPPaymentType = "STDO" Then
          Select Case mvTA.PPPaymentType
            Case "DIRD"
              vPayMethodReason = vDDReason
            Case "CCCA"
              vPayMethodReason = vCCReason
            Case Else
              vPayMethodReason = vSOReason
          End Select
        End If
      End If

      If vSourceCode.Length > 0 AndAlso vReason.Length > 0 Then
        If mvTA.ContactVATCategory.Length = 0 Then
          Dim vControl As TextLookupBox
          vControl = mvCurrentPage.EditPanel.FindTextLookupBox("ContactNumber", False)
          If vControl Is Nothing Then vControl = mvTA.Pages(CareServices.TraderPageType.tpTransactionDetails.ToString).EditPanel.FindTextLookupBox("ContactNumber", False)
          'BR16642: If Contact control ContactInfo not set in TransactionDetails page use ContactSelection Contact control ContactInfo
          If vControl Is Nothing OrElse vControl.ContactInfo Is Nothing Then vControl = mvTA.Pages(CareNetServices.TraderPageType.tpContactSelection.ToString).EditPanel.FindTextLookupBox("ContactNumber", False)
          If vControl IsNot Nothing AndAlso vControl.ContactInfo IsNot Nothing Then
            mvTA.ContactVATCategory = vControl.ContactInfo.VATCategory
          End If
        End If
        Dim vList As New ParameterList(True)
        vList("Source") = vSourceCode
        vList("ReasonForDespatch") = vReason
        vList.AddItemIfValueSet("PayMethodReason", vPayMethodReason)
        vList("VatCategory") = mvTA.ContactVATCategory
        If pPaymentIncentive Then vList("Amount") = pAmount.ToString
        Dim vForm As New frmIncentives
        mvTA.IncentiveDataSet = vForm.GetIncentivesData(vList, True, True, pPaymentIncentive, pCanAddEnclosure)
      Else
        Return Nothing
      End If
    End If
    If mvTA.IncentiveDataSet.Tables.Contains("DataRow") Then
      Return mvTA.IncentiveDataSet.Tables("DataRow")
    Else
      Return Nothing
    End If
  End Function
  Private Sub AddIncentivesLines(ByVal pList As ParameterList, ByVal pPaymentIncentive As Boolean, ByVal pSource As String, ByVal pAmount As Double)
    Dim vCanAddEnclosure As Boolean
    Dim vTable As DataTable = GetIncentivesTable(pPaymentIncentive, pSource, pAmount, vCanAddEnclosure)
    If vTable IsNot Nothing Then
      Dim vIndex As Integer = 1
      For Each vRow As DataRow In vTable.Rows
        pList.ObjectValue("IncentiveLine" & vIndex.ToString) = vRow
        vIndex += 1
      Next
      If vCanAddEnclosure AndAlso mvTA.PreFulfilledIncentives AndAlso pPaymentIncentive = False AndAlso mvTA.FulfilIncentives = False Then
        If ShowQuestion(QuestionMessages.QmFulfilIncentives, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          pList("FulfilIncentives") = "Y"
          mvTA.FulfilIncentives = True
        End If
      End If
      If pPaymentIncentive Then mvTA.IncentiveDataSet = Nothing
    End If
  End Sub

  Private ReadOnly Property CanAddIncentives() As Boolean
    Get
      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpProductDetails, CareServices.TraderPageType.tpEventBooking, CareServices.TraderPageType.tpAccommodationBooking,
             CareServices.TraderPageType.tpServiceBooking, CareServices.TraderPageType.tpOutstandingScheduledPayments,
             CareServices.TraderPageType.tpInvoicePayments, CareServices.TraderPageType.tpLegacyBequestReceipt,
             CareServices.TraderPageType.tpConfirmProvisionalTransactions
          Return mvTA.TransactionSource IsNot Nothing AndAlso mvTA.TransactionSource.Length > 0
        Case Else
          Return False
      End Select
    End Get
  End Property

  Private ReadOnly Property CanAmendPurchaseOrderAmount() As Boolean
    'Only returns False when 1. Regular Payments, 2. Existing PO and 3. Atleast one payment has been authorised.
    Get
      Return Not (mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments AndAlso mvTA.PurchaseOrderNumber > 0 AndAlso
             mvTraderPages(CareNetServices.TraderPageType.tpPurchaseOrderDetails.ToString).EditPanel.FindTextLookupBox("PurchaseOrderType").Enabled = False)
    End Get
  End Property
#End Region

#Region " Page Processing "

  Private Sub ProcessData(ByVal pType As CareServices.TraderProcessDataTypes, ByVal pAddValues As Boolean, Optional ByRef pErrorNumber As CDBNETCL.CareException.ErrorNumbers = Nothing)
    Dim vResetProgressBar As Boolean
    Try
      Dim vList As New ParameterList(True)
      Dim vValid As Boolean = True
      'Jira 785: Ignore page validation errors when cancelling transaction
      Dim vValidate As Boolean = Not (pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction)
      If pAddValues Then
        If pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpAmendEventBooking OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking) Then
          'These pages have StartTime & EndTime on them that refer the the Event StartDate & EndDate, so need to add these two in order to validate the times correctly
          Dim vCareEventInfo As CareEventInfo = mvCurrentPage.EditPanel.FindTextLookupBox("EventNumber", False).CareEventInfo
          If vCareEventInfo IsNot Nothing Then
            vList("StartDate") = vCareEventInfo.StartDate.ToShortDateString
            vList("EndDate") = vCareEventInfo.EndDate.ToShortDateString
          End If
        End If
        PreValidateItems() 'The data might have been changed outside of the trader page
        If Not (pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpOutstandingScheduledPayments AndAlso (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment OrElse mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.EventAdjustment)) _
          And Not (pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceSummary) _
          And Not (pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceProduction) _
          And Not (pType = CareNetServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation) Then
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpCardDetails Then
            If mvTA.OnlineCCAuthorisation AndAlso Me.CardAuthoriser IsNot Nothing Then
              If Me.CardAuthoriser.IsAuthorised Then
                Me.CardAuthoriser.SetServerValues(vList)
                vValid = True
              Else
                vValid = False
              End If
            Else
              vValid = mvCurrentPage.EditPanel.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
            End If
            If Not vValid AndAlso
               pType = CareNetServices.TraderProcessDataTypes.tpdtFinished AndAlso
               mvTA.OnlineCCAuthorisation AndAlso
               Not mvTA.RequireCCAuthorisation AndAlso
               ShowQuestion(QuestionMessages.QmBypassCcAuthorisation, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vList("BypassCcAuthorisation") = "Y"
              vValid = True
            End If
          Else
            vValid = mvCurrentPage.EditPanel.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
          End If
        End If
        If vValid = False AndAlso mvTA.PPNumbersCreated.Count > 0 AndAlso pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails _
          OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpAccommodationBooking OrElse
          mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership OrElse
          mvCurrentPage.PageType = CareServices.TraderPageType.tpPayments Then
            Dim vDataSet As DataSet
            GetAdditionalValues(vList, pType)
            vList.IntegerValue("CurrentPageType") = mvCurrentPage.PageType
            vDataSet = DataHelper.ProcessTraderData(pType, mvTA.ApplicationNumber, vList)
            Dim vRow As DataRow = vDataSet.Tables("Result").Rows(0)
            If vRow.Table.Columns.Contains("InformationMessage") Then
              ShowInformationMessage(vRow.Item("InformationMessage").ToString)
            End If
            vValidate = True
          End If
        End If
        If vValid = True AndAlso mvTA.ApplicationType = ApplicationTypes.atBatchInvoiceGeneration AndAlso pType = CareNetServices.TraderProcessDataTypes.tpdtCancelTransaction Then
          If mvExtApplication IsNot Nothing Then mvExtApplication.CloseExternalApplication(False)
        End If
        'moved this call outside the ifpaddvalues condition
        If vValid = True OrElse Not vValidate Then
          Select Case mvCurrentPage.PageType
            Case CareServices.TraderPageType.tpCollectionPayments
              If pType = CareServices.TraderProcessDataTypes.tpdtNextPage Then vValid = ValidateCollectionPayment()
            Case CareServices.TraderPageType.tpContactSelection
              If mvTA.TransactionType = "MEMC" AndAlso mvTA.PaymentPlan IsNot Nothing Then vValid = ValidateCMT()
            Case CareServices.TraderPageType.tpCreditCustomer
              If mvTA.CreditTermsChanged = True AndAlso mvTA.PayerContactNumber > 0 Then
                Dim vResponse As DialogResult = ShowQuestion(QuestionMessages.QmCreditTermsChanged, MessageBoxButtons.YesNoCancel)
                mvTA.SavePaymentTerms = (vResponse = System.Windows.Forms.DialogResult.Yes)
                vValid = (vResponse <> System.Windows.Forms.DialogResult.Cancel)
              End If
            Case CareServices.TraderPageType.tpProductDetails
              If mvTA.TransactionType = "SALE" AndAlso (Not pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction) Then vValid = ValidateProductSale()
            Case CareServices.TraderPageType.tpAddressMaintenance
              If pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
                MaintainContactAddresses(CInt(vList("ContactNumber")), CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, Me, vList, True)
              End If
            Case CareServices.TraderPageType.tpPaymentPlanSummary
              If pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvTA.ApplicationType = ApplicationTypes.atMaintenance Or mvTA.PayPlanConvMaintenance) Then
                If mvTA.PPDDataSet.Tables("DataRow") IsNot Nothing Then
                  Dim vTable As DataTable = mvTA.PPDDataSet.Tables("DataRow")
                  If vTable.Columns.Contains("EffectiveDate") Then
                    For Each vExistingRow As DataRow In vTable.Rows
                      If vValid = True AndAlso vExistingRow("EffectiveDate").ToString.Length > 0 Then
                        If Date.Parse(vExistingRow("EffectiveDate").ToString) > Today Then
                          ShowInformationMessage(InformationMessages.ImNewDetailLine)
                          vValid = False
                        End If
                      End If
                    Next
                  End If
                End If
              End If
            Case CareServices.TraderPageType.tpEventBooking
              Dim vSalesContactNumber As TextLookupBox = mvCurrentPage.EditPanel.FindTextLookupBox("SalesContactNumber")
              mvTA.SalesContactNumber = IntegerValue(vSalesContactNumber.Text)
              If pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction AndAlso vList.ContainsKey("InterestOnly") = True AndAlso IntegerValue(vList("InterestBookingNumber")) > 0 Then
                vValid = (ShowQuestion(QuestionMessages.QmConfirmConvertInterestBooking, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
              End If
            Case CareNetServices.TraderPageType.tpPostageAndPacking
              If mvTA.StockTransactionID > 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_fp_postage_warn_stock) _
                AndAlso vList("Product").Length = 0 Then
                If ShowQuestion(QuestionMessages.QmAddCarriageCost, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then vValid = False
              End If
              If vList("Product").Length = 0 Then mvPAPDone = True
              If IntegerValue(vList("Amount2")) > mvTA.TransactionAmount Then
                If ShowQuestion(QuestionMessages.QmInvalidPAPAmount, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then vValid = False
              End If
          End Select
        End If
        If vValid OrElse Not vValidate Then mvTA.SaveApplicationValues(mvCurrentPage, vList)
      End If

      If vValid OrElse Not vValidate Then vValid = GetAdditionalValues(vList, pType)
      If vValid Then vValid = ValidatePage()
      If vValid OrElse Not vValidate Then
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceProduction Then
          If vList.Contains("DisplayInvoices") AndAlso vList("DisplayInvoices") = "N" Then
            pType = CareServices.TraderProcessDataTypes.tpdtFinished
          ElseIf Not vList.Contains("DisplayInvoices") AndAlso vList.Contains("Company") Then
            pType = CareServices.TraderProcessDataTypes.tpdtFinished
          End If
        End If
        mvTA.GetApplicationValues(vList)     'Add the application items which may be required
        If (mvTA.TransactionType = "MEMB" Or mvTA.TransactionType = "MEMC") And
           ((pType = CareServices.TraderProcessDataTypes.tpdtNextPage And mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary) _
           OrElse (pType = CareServices.TraderProcessDataTypes.tpdtAddMemberSummary) _
           OrElse (pType = CareServices.TraderProcessDataTypes.tpdtPreviousPage AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails)) Then
          vValid = AddMembershipData(vList)
          If mvTA.TransactionType = "MEMC" Then
            If vValid Then vValid = AddTransactionData(vList)
          End If
        ElseIf pType = CareServices.TraderProcessDataTypes.tpdtFinished OrElse
               (mvTA.PayMethodsAtEnd AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpCreditCustomer AndAlso pType = CareNetServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvTA.TransactionPaymentMethod = "CCIN") Then
          vValid = AddTransactionData(vList)
          If mvTA.TransactionDonationAmount > 0 Then vList.Add("AddDonationToTrans", mvTA.TransactionDonationAmount.ToString)
        ElseIf mvTA.TransactionType = "MEMC" AndAlso ((pType = CareServices.TraderProcessDataTypes.tpdtAddMemberSummary) OrElse (pType = CareServices.TraderProcessDataTypes.tpdtNextPage And mvCurrentPage.PageType = CareServices.TraderPageType.tpAmendMembership)) Then
          If vValid = True AndAlso mvCurrentRow >= 0 Then
            vList("MembershipNumber") = mvMembersDGR.GetValue(mvCurrentRow, "MembershipNumber")
          End If
        ElseIf mvTA.TransactionType = "APAY" AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpCreditCardAuthority OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpDirectDebit OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpStandingOrder) Then
          If vValid Then vValid = AddTransactionData(vList)
        End If
        If mvTA.AnalysisDataSet.Tables("DataRow") IsNot Nothing Then
          If mvTA.AnalysisDataSet.Tables("DataRow").Select("TraderLineType='P' AND ProductCode = '" & mvTA.ProductCode & "' AND Rate = '" & mvTA.RateCode & "'").Length > 0 Then
            vList("GetDefaultProductAndRate") = "N"
          Else
            vList("GetDefaultProductAndRate") = "Y"
          End If
        End If

        If vValid = True AndAlso mvTA.TransactionType = "MEMC" AndAlso pType = CareNetServices.TraderProcessDataTypes.tpdtFinished AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpMembershipPayer Then
          vValid = AddAdditionalCMTData(vList)
        End If
      End If
      If vValid OrElse Not vValidate Then
        GetEditValues(pType, vList)

        '=====================================================
        'Add any DataRows that need sending back to the server
        '=====================================================
        Dim vProcessData As Boolean = True
        'Analysis(lines
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished OrElse
           pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction OrElse
           (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso vList.ContainsKey("TransactionDateChanged")) OrElse
           (mvTA.PayPlanPayMethod AndAlso
            Not pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine AndAlso
            Not pType = CareNetServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine) Then
          'If we have finished the transaction / we are going to cancel / transaction date has changed then send back all the analysis lines we have so far
          AddAnalysisLines(vList)
          If pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
            AddOriginalOPSLines(vList) 'BR19606 return the stored, unchanged OPS to the server, if present, but only for Cancel
          End If
        End If
        'PaymentPlans
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
          Select Case mvTA.TransactionType
            Case "DONR", "MEMB", "SUBS", "CMEM", "CDON", "CSUB", "LOAN"
              AddPPDLines(vList)
            Case "SALE", "EVNT", "ACOM", "SRVC"
              'required if we're creating a payment plan from an unbalanced transaction
              If mvTA.PayPlanPayMethod Then AddPPDLines(vList)
          End Select
        End If
        AddRemovedSchPaymentLines(vList)
        'Payment Schedule
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished And mvCurrentPage.PageType = CareServices.TraderPageType.tpScheduledPayments Then
          AddOPSLines(vList)
        End If
        If pType = CareServices.TraderProcessDataTypes.tpdtNextPage And mvCurrentPage.PageType = CareServices.TraderPageType.tpOutstandingScheduledPayments Then
          AddOSPLines(vList)
        End If
        'Members
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpAdvancedCMT)) OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished OrElse pType = CareServices.TraderProcessDataTypes.tpdtAddMemberSummary Then
          AddMemberLines(vList)
        End If
        'Invoices
        If pType = CareServices.TraderProcessDataTypes.tpdtNextPage And mvCurrentPage.PageType = CareServices.TraderPageType.tpInvoicePayments Then
          AddInvoiceLines(vList)
        End If
        'Collection Boxes
        If pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpCollectionPayments Then
          AddCollectionBoxData(vList)
        End If
        'Purchase Orders
        If (pType = CareServices.TraderProcessDataTypes.tpdtFinished AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderSummary OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderPayments)) _
            OrElse (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderProducts OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderDetails)) Then
          AddPOSLines(vList)
        End If
        'Purchase Invoices
        If (pType = CareServices.TraderProcessDataTypes.tpdtFinished AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseInvoiceSummary) _
           OrElse (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseInvoiceProducts OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseInvoiceDetails)) Then
          AddPISLines(vList)
        End If
        'Gift Aid Declaration
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished) AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpGiftAidDeclaration Then
          AddGiftAid()
        End If
        'Purchase Order Payments
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished) AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderPayments Then
          AddPPALines(vList, pType)
          If pType = CareServices.TraderProcessDataTypes.tpdtNextPage Then vProcessData = False
        End If
        'Event Booking Lines (When using Event Pricing Matrix only)
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso (mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpAmendEventBooking)) Then
          AddEventBookingLines(vList)
        End If
        'Service Bookings
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished) AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking Then
          AddServiceBooking(vList)
        End If
        'Exam Booking Lines
        If (pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpExamBooking) Then
          AddExamBookingLines(vList)
        End If
        'CMT Lines (for Advanced CMT only)
        If (pType = CareNetServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpAdvancedCMT) OrElse (pType = CareNetServices.TraderProcessDataTypes.tpdtFinished AndAlso mvTA.TransactionType = "MEMC") Then
          AddCMTLines(vList)
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpAdvancedCMT Then AddTransactionData(vList)
        End If
        If (mvTA.AddIncentivesLinesRequired AndAlso pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary) _
              OrElse (mvTA.CheckIncentives AndAlso ((pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso
              (CanAddIncentives OrElse ((mvTA.TransactionType = "MEMB" AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership) OrElse mvTA.TransactionType = "CMEM" _
              OrElse (mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts OrElse
              (mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails AndAlso mvTA.PPDDataSet.Tables.Contains("DataRow") AndAlso mvTA.PPDDataSet.Tables("DataRow").Rows.Count > 0)))) _
              OrElse pType = CareServices.TraderProcessDataTypes.tpdtFinished))) Then
          Dim vPaymentIncentive As Boolean = pType = CareServices.TraderProcessDataTypes.tpdtNextPage AndAlso CanAddIncentives
          Dim vSource As String = ""
          Dim vAmount As Double = 0
          If vPaymentIncentive Then
            If FindControl(mvCurrentPage.EditPanel, "Source", False) IsNot Nothing Then vSource = mvCurrentPage.EditPanel.GetValue("Source")
            If FindControl(mvCurrentPage.EditPanel, "Amount", False) IsNot Nothing Then vAmount = DoubleValue(mvCurrentPage.EditPanel.GetValue("Amount"))
          End If
          If mvTA.PPIncentivesCompleted = False OrElse (mvTA.PPIncentivesCompleted AndAlso ((mvTA.TransactionType = "MEMB" AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership) OrElse vPaymentIncentive)) Then
            If mvTA.PPIncentivesCompleted And vPaymentIncentive = False Then
              'Reset Incentive Data once a PP is created
              mvTA.PPIncentivesCompleted = False
              mvTA.IncentiveDataSet = Nothing
            End If
            If mvTA.TransactionType = "MEMB" AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership _
            AndAlso ((vList.Contains("NumberOfMembers") AndAlso IntegerValue(vList("NumberOfMembers")) > 1) OrElse (vList.Contains("MaxFreeAssociates") AndAlso IntegerValue(vList("MaxFreeAssociates")) > 0)) Then
              'BR16407: Delay the display of incentive form and adding incentive lines when the MembershipMembersSummary page will be the next page
              mvTA.AddIncentivesLinesRequired = True
            Else
              mvTA.AddIncentivesLinesRequired = False
              AddIncentivesLines(vList, vPaymentIncentive, vSource, vAmount)
            End If
          End If
        End If
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceSummary Then
          AddSelectedInvoices(vList)
        End If
        ' BR19597 - Pass the Analysis Data (PRD) to the Transaction form (TRD) from the TAS screen, and back. Required to effect changes in Source in TRD when editing. 
        If pType = CareNetServices.TraderProcessDataTypes.tpdtPreviousPage AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpTransactionAnalysisSummary Then
          AddSelectedAnalysisLine(vList)
        End If
        If pType = CareNetServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpTransactionDetails Then
          AddSelectedAnalysisLine(vList)
        End If
        '=======================
        'End processing DataRows
        '=======================

        'All set - so send the data to the server and process the returns by saving any analysis that is returned
        vList.IntegerValue("CurrentPageType") = mvCurrentPage.PageType
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished AndAlso Not String.IsNullOrEmpty(mvTA.LinkToCommunication) Then
          Dim vCommNumber As Integer = FormHelper.GetCommunicationsNumber
          If vCommNumber > 0 Then vList.IntegerValue("LinkCommNumber") = vCommNumber
        End If

        Dim vDataSet As DataSet
        If vProcessData AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpBatchInvoiceSummary AndAlso pType = CareNetServices.TraderProcessDataTypes.tpdtFinished Then
          Dim vPrintPreview As Boolean
          If FindControl(mvCurrentPage.EditPanel, "PrintPreview", False) IsNot Nothing Then vPrintPreview = BooleanValue(mvCurrentPage.EditPanel.GetValue("PrintPreview"))
          If vPrintPreview Then
            'Preview the Invoices and do nothing else
            vDataSet = DataHelper.ProcessTraderData(pType, mvTA.ApplicationNumber, vList)
            If vDataSet IsNot Nothing Then
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vDataSet)
              GenerateMailmerge(CareNetServices.TraderMailmergeType.tmtInvoice, vRow)
              vProcessData = False
            End If
          End If
        End If
        If vProcessData Then
          Dim vAppointmentSet As ContactAppointmentSet = Nothing
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking AndAlso pType = CareServices.TraderProcessDataTypes.tpdtNextPage Then
            vAppointmentSet = CalendarApplication.CheckCalendarConflict(CalendarApplication.CalendarUpdateType.AddEventBooking, vList)
          ElseIf pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then    'First analysis line
              Dim vTable As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
              For Each vExistingRow As DataRow In vTable.Rows
                If vExistingRow("TraderLineType").ToString.Equals("E") AndAlso vExistingRow.Table.Columns.Contains("EventBookingNumber") AndAlso vExistingRow("EventBookingNumber").ToString.Length > 0 Then
                  Dim vEventBookingList As New ParameterList(True)
                  vEventBookingList("BookingNumber") = vExistingRow("EventBookingNumber").ToString
                  vAppointmentSet = CalendarApplication.CheckCalendarDelete(CalendarApplication.CalendarDeleteType.EventBooking, vEventBookingList)
                End If
              Next
            End If
          ElseIf pType = CareServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine Then
            Dim vEventRow As DataRow = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow)
            If vEventRow("TraderLineType").ToString.Equals("E") AndAlso vEventRow.Table.Columns.Contains("EventBookingNumber") AndAlso vEventRow("EventBookingNumber").ToString.Length > 0 Then
              Dim vEventBookingList As New ParameterList(True)
              vEventBookingList("BookingNumber") = vEventRow("EventBookingNumber").ToString
              vAppointmentSet = CalendarApplication.CheckCalendarDelete(CalendarApplication.CalendarDeleteType.EventBooking, vEventBookingList)
            End If
          End If
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpTransactionAnalysisSummary And pType = CareNetServices.TraderProcessDataTypes.tpdtFinished Then
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0 Then
              If mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count - 1).Item("TraderTransactionType").ToString = "P&P" OrElse mvPAPDone Then
                vList("PAPDone") = "Y"
              End If
            End If
          End If

          If pType = CareNetServices.TraderProcessDataTypes.tpdtFinished AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpCardDetails AndAlso mvTA.OnlineCCAuthorisation AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type) = "SCXLVPCSCP" Then
            vList.IntegerValue("TraderApplication") = mvTA.ApplicationNumber
            Dim vProcess As New AsyncProcessHandler(pType, vList)
            EPL_ShowMessage(mvCurrentPage.EditPanel, "Authorising Credit Card")
            vResetProgressBar = True
            vDataSet = vProcess.GetDataSetFromResult(prgBar)
          ElseIf pType = CareNetServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpEventBooking AndAlso vList.ValueIfSet("BookingNumber").Length > 0 Then
            vList.IntegerValue("TraderApplication") = mvTA.ApplicationNumber
            Dim vProcess As New AsyncProcessHandler(pType, vList)
            vDataSet = vProcess.GetDataSetFromResult
          Else
            vDataSet = DataHelper.ProcessTraderData(pType, mvTA.ApplicationNumber, vList)

            If pType = CareNetServices.TraderProcessDataTypes.tpdtNextPage AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpBatchInvoiceProduction AndAlso mvTA.mvUnpostedBatchMsgInPrint = True Then
              If vList.Contains("StartBatch") AndAlso vList.IntegerValue("StartBatch") > 0 AndAlso vList.Contains("EndBatch") Then
                Dim vBatchCount As Integer = DataHelper.GetUnpostedBatchCount(vList.IntegerValue("StartBatch"), vList.IntegerValue("EndBatch"), "CS")
                If vBatchCount > 0 Then
                  ShowInformationMessage(InformationMessages.ImUnpostedBatchCount, vList("StartBatch"), vList("EndBatch"), vBatchCount.ToString)
                End If
              End If
            End If
          End If
          If vAppointmentSet IsNot Nothing AndAlso vList.ContainsKey("EventNumber") Then
            If pType = CareServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine OrElse pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
              Dim vParams As New ParameterList
              vParams("EventNumber") = vList("EventNumber")
              CalendarApplication.DeleteAppointment(CalendarApplication.CalendarDeleteType.EventBooking, vAppointmentSet, vParams)
            Else
              Dim vParams As New ParameterList
              vParams.Add("EventNumber", vList("EventNumber"))
              CalendarApplication.UpdateAppointment(CalendarApplication.CalendarUpdateType.AddEventBooking, vAppointmentSet, vParams)
            End If
          End If
          SetAnalysisEditable(CInt(IIf((mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0), 0, -1)))
          SetPPDEditable(CInt(IIf(mvPPSDGR IsNot Nothing AndAlso mvPPSDGR.RowCount > 0, 0, -1)))

          'Process DataSets
          SetTraderAnalysisLines(vDataSet, (pType = CareNetServices.TraderProcessDataTypes.tpdtEditAnalysisLine))
          SetPPDLines(vDataSet)
          SetOPSLines(vDataSet)
          If mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.Adjustment Then 'Only backup the OPs if an Adjustment is being made
            SetOriginalOPSLine(vDataSet) 'BR19606 Get the unchanged OPS if present, and store it.
          End If
          SetRemovedSchPayments(vDataSet)
          SetMemberLines(pType, vDataSet)
          SetOSPLines(vDataSet)
          SetPOSLines(vDataSet)
          SetPPALines(vDataSet)
          SetPISLines(vDataSet)
          SetBatchInvoices(vDataSet)
          SetCMTPPDetails(vDataSet)

          'Handle the setting  of the various buttons based on the return values
          Dim vRow As DataRow = vDataSet.Tables("Result").Rows(0)
          SetButtons(vRow)

          If vRow.Table.Columns.Contains("NonFinancialBatchNumber") Then
            mvNonFinancialBatchNumber = IntegerValue(vRow("NonFinancialBatchNumber"))
            mvNonFinancialTransactionNumber = IntegerValue(vRow("NonFinancialTransactionNumber"))
          End If
          If vRow.Table.Columns.Contains("PaymentPlanNumber") Then
            mvPPNumber = IntegerValue(vRow("PaymentPlanNumber"))
          End If
          If vRow.Table.Columns.Contains("BankDetailsNumber") Then
            mvTA.BankDetailsNumber = IntegerValue(vRow("BankDetailsNumber"))
          End If
          If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceSummary AndAlso vRow.Table.Columns.Contains("PrintJobNumber") Then
              vList("PrintJobNumber") = vRow("PrintJobNumber").ToString
              vList("Company") = vRow("Company").ToString
            End If
            ProcessFinishedReturns(vDataSet, vList)
          ElseIf pType = CareServices.TraderProcessDataTypes.tpdtDeleteAnalysisLine Then
            If vRow.Table.Columns.Contains("EventNumber") Then
              'Deleted an Event or Room Booking, so refresh Events data
              RefreshCardSet(RefreshTypes.rtEventBooking, 0, mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow))
            End If
            SetAnalysisEditable(CInt(IIf(mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0, 0, -1)))
            If vRow.Table.Columns.Contains("InformationMessage") Then
              'Display message confirming item deleted etc.
              ShowInformationMessage(vRow.Item("InformationMessage").ToString)
            End If
          ElseIf pType = CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
            If mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then    'First analysis line
              Dim vTable As DataTable = mvTA.AnalysisDataSet.Tables("DataRow")
              For Each vExistingRow As DataRow In vTable.Rows
                'if event lines then refresh the event card set
                If vExistingRow("TraderLineType").ToString.Equals("E") AndAlso vExistingRow.Table.Columns.Contains("EventBookingNumber") AndAlso vExistingRow("EventBookingNumber").ToString.Length > 0 Then
                  RefreshCardSet(RefreshTypes.rtEventBooking, 0, vExistingRow)
                End If
              Next
            End If
            If vRow.Table.Columns.Contains("InformationMessage") Then
              'Display message confirming item deleted etc.
              ShowInformationMessage(vRow.Item("InformationMessage").ToString)
            End If
          ElseIf (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse (pType = CareServices.TraderProcessDataTypes.tpdtPreviousPage AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails)) AndAlso mvTA.StockSales = True Then
            mvTA.StockSales = False   'Reset flag to show that we have finished the Stock Sale
            mvTA.TransactionDateChanged = False   'Reset flag so that we will not re-calculate VAT again         
          ElseIf pType = CareNetServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then
            If vDataSet.Tables.Contains("PPDLine") Then
              Dim vPPDTable As DataTable = vDataSet.Tables("PPDLine")
              mvTA.PaymentPlanDetailsPricing.InitFromDataRow(vPPDTable.Rows(0))
            End If
          ElseIf pType = CareNetServices.TraderProcessDataTypes.tpdtEditTransaction Then
            If vRow.Table.Columns.Contains("COM_Notes") Then
              mvTA.TransactionNote = vRow("COM_Notes").ToString
            End If
          End If
          ProcessNextMove(pType, vDataSet)
        End If
      End If
    Catch vEx As CareException
      pErrorNumber = vEx.ErrorNumber
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enQuantityExceedsMaximum, CareException.ErrorNumbers.enEventBooking,
             CareException.ErrorNumbers.enCCAuthorisationFailed,
             CareException.ErrorNumbers.enCannotBookJointToEvent, CareException.ErrorNumbers.enBookingBatchNotPosted,
             CareException.ErrorNumbers.enCannotBookEvent, CareException.ErrorNumbers.enPPMaxProdNumbers, CareException.ErrorNumbers.enMaxProdNumbers,
             CareException.ErrorNumbers.enDDPrefixNoAlpha, CareException.ErrorNumbers.enDDSuffixNoAlpha, CareException.ErrorNumbers.enDDReferenceSameCharacters,
             CareException.ErrorNumbers.enDDReferenceSixAlphas, CareException.ErrorNumbers.enStandingOrderBalanceWrong, CareException.ErrorNumbers.enNoOPSForPP,
             CareException.ErrorNumbers.enPaymentScheduleNotCreated, CareException.ErrorNumbers.enCMTUnableRemoveIncentive, CareException.ErrorNumbers.enReversalTransTypeNotSetUp,
             CareException.ErrorNumbers.enTraderApplicationInvalid, CareException.ErrorNumbers.enNoInvoicesMatchCriteria, CareException.ErrorNumbers.enCreditCustomerMissing,
             CareException.ErrorNumbers.enSetInvoiceNumberDuplicateRecord, CareException.ErrorNumbers.enMissingClaimDates, CareException.ErrorNumbers.enCardRejectedAsOverCeilingLimit,
             CareException.ErrorNumbers.enCardHasExpired, CareException.ErrorNumbers.enAuthorisationHasBeenRefused, CareException.ErrorNumbers.enMerchantNumberNotSetUp,
             CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout, CareException.ErrorNumbers.enNoAuthorisationLevel, CareException.ErrorNumbers.enCannotFindLoanPayment,
             CareException.ErrorNumbers.enExamBooking, CareException.ErrorNumbers.enScheduleClashInBooking, CareException.ErrorNumbers.enScheduleClashExistingBooking,
             CareException.ErrorNumbers.enPaymentPlanBalanceCannotBeNegative, CareException.ErrorNumbers.enPaymentPlanRenewalAmountCannotBeNegative,
             CareException.ErrorNumbers.enCMTMemberTypeRefundProductNotSet, CareException.ErrorNumbers.enCMTEntitlementRefundProductNotSet,
             CareException.ErrorNumbers.enDirectDebitReferenceNotUnique, CareException.ErrorNumbers.enCannotUseInitialPeriodIncentives, CareException.ErrorNumbers.enCannotUseInitialPeriodIncentivesAtConversion,
             CareException.ErrorNumbers.enExamMultipleSessionInBooking, CareException.ErrorNumbers.enTraderUnsupportedFeature, CareException.ErrorNumbers.enNoExchangeRate,
             CareException.ErrorNumbers.enCannotEditAuthorisedTransaction
          ShowInformationMessage(vEx.Message)
          If vEx.ErrorNumber = CareException.ErrorNumbers.enCCAuthorisationFailed AndAlso
             Me.CardAuthoriser IsNot Nothing Then
            InitCardtAuthorisation()
          End If
        Case CareException.ErrorNumbers.enCanUpdatePPFixedAmount, CareException.ErrorNumbers.enMustUpdatePPFixedAmount
          If vEx.ErrorNumber = CareException.ErrorNumbers.enCanUpdatePPFixedAmount Then
            mvTA.CMTUpdatePPFixedAmount = ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes
          Else
            If ShowQuestion(vEx.Message, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then mvTA.CMTUpdatePPFixedAmount = True
          End If
          If mvTA.CMTUpdatePPFixedAmount.HasValue Then ProcessData(pType, pAddValues)
        Case CareException.ErrorNumbers.enAppointmentConflict
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking Then
            'Double bookings are allowed for service bookings
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_allow_double_booking, False) Then
              Dim vMsg As String = vEx.Message & Environment.NewLine & QuestionMessages.QmDoubleBook
              mvTA.ConfirmCalendarConflict = ShowQuestion(vMsg, MessageBoxButtons.YesNo) = DialogResult.Yes
              If mvTA.ConfirmCalendarConflict Then ProcessData(pType, True)
            Else
              ShowInformationMessage(vEx.Message)
            End If
          Else
            ShowInformationMessage(vEx.Message)
          End If
        Case CareException.ErrorNumbers.enAddressChangePrompt
          mvTA.ChangeBranchWithAddress = CBoolYN(ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
          ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
        Case CareException.ErrorNumbers.enTraderCreateCommLink
          mvTA.CreateCommLink = CBoolYN(ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
          ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
        Case CareException.ErrorNumbers.enGADDatesOverlap
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpGiftAidDeclaration Then mvCurrentPage.EditPanel.SetErrorField("StartDate", vEx.Message, True)
        Case CareException.ErrorNumbers.enContactAccountChangePrompt
          mvTA.CreateContactAccount = CBoolYN(ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
          ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
        Case CareException.ErrorNumbers.enPaymentPlanNotFound
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enPageNotFound
          ShowInformationMessage(vEx.Message, mvNextPageCode)
          If mvCurrentPage IsNot Nothing AndAlso mvCurrentPage.PageType.Equals(CareServices.TraderPageType.tpTransactionAnalysisSummary) AndAlso mvTA.EditExistingTransaction = True Then
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
          End If
        Case CareException.ErrorNumbers.enInvalidSBStartDays, CareException.ErrorNumbers.enNoFinancialPeriod
          ShowWarningMessage(vEx.Message)
        Case CareException.ErrorNumbers.enInvalidSBDuration
          If DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualDatabaseAdministrator OrElse DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualSupervisor Then
            Dim vMsg As String = vEx.Message & Environment.NewLine & QuestionMessages.QmContinue
            mvTA.ConfirmSBDuration = ShowQuestion(vMsg, MessageBoxButtons.YesNo) = DialogResult.Yes
            If mvTA.ConfirmSBDuration Then ProcessData(pType, True)
          Else
            ShowErrorMessage(vEx.Message)
          End If
        Case CareException.ErrorNumbers.enInvalidShortStay
          If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciServiceControlRestriction) Then
            Dim vMsg As String = vEx.Message & Environment.NewLine & QuestionMessages.QmContinue
            mvTA.ConfirmSBShortStay = ShowQuestion(vMsg, MessageBoxButtons.YesNo) = DialogResult.Yes
            If mvTA.ConfirmSBShortStay Then ProcessData(pType, True)
          Else
            ShowErrorMessage(vEx.Message)
          End If
        Case CareException.ErrorNumbers.enInvalidEventBookingLine, CareException.ErrorNumbers.enInvalidServiceBookingLine, CareException.ErrorNumbers.enInvalidAccomodationBookingLine
          ShowErrorMessage(vEx.Message)
        Case CareException.ErrorNumbers.enInvoiceAllocationError, CareException.ErrorNumbers.enUnallocateCreditNote, CareException.ErrorNumbers.enAllocateOrUnallocateCreditNote
          If vEx.ErrorNumber = CareException.ErrorNumbers.enInvoiceAllocationError OrElse vEx.ErrorNumber = CareException.ErrorNumbers.enUnallocateCreditNote Then
            If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              mvTA.AllocationsChecked = True
              If vEx.ErrorNumber = CareException.ErrorNumbers.enUnallocateCreditNote Then mvTA.UnallocateCreditNote = True
            End If
          ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enAllocateOrUnallocateCreditNote Then
            Dim vDialogueResult As System.Windows.Forms.DialogResult = ShowQuestion(vEx.Message, MessageBoxButtons.YesNoCancel)
            If vDialogueResult <> System.Windows.Forms.DialogResult.Cancel Then
              mvTA.AllocationsChecked = True
              mvTA.UnallocateCreditNote = vDialogueResult = System.Windows.Forms.DialogResult.No
            End If
          End If
          If mvTA.AllocationsChecked Then
            ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
            mvTA.AllocationsChecked = False
            mvTA.UnallocateCreditNote = False
          End If
        Case CareException.ErrorNumbers.enAdvancedCMTIncorrectlySetUp, CareException.ErrorNumbers.enCannotGetAnalysis
          ShowErrorMessage(vEx.Message)
          Me.Close()    'User cannot do anything so close Trader
        Case CareException.ErrorNumbers.enPaymentPlanBalanceChangedButNotDetails
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enCannotUseHistoryOnlyBankAccount
          ShowInformationMessage(vEx.Message)
        Case Else
          DataHelper.HandleException(vEx)
      End Select

    Catch vException As Exception

      DataHelper.HandleException(vException)
    Finally
      If vResetProgressBar Then
        prgBar.Value = 0
        EPL_ShowMessage(mvCurrentPage.EditPanel, "")
      End If
    End Try
  End Sub

#End Region

#Region " Post-Server Processing "

  Private Sub AddCMDMailingCode(ByVal pList As ParameterList)
    Dim vMailing As String = GetPageValue(CareServices.TraderPageType.tpTransactionDetails, "Mailing")
    If vMailing.Length = 0 Then vMailing = GetPageValue(CareServices.TraderPageType.tpContactSelection, "Mailing")
    If Len(vMailing) = 0 Then
      Dim vSourceCode As String = ""
      If mvTA.NonFinancialBatch And mvTA.BatchLedApp Then vSourceCode = mvTA.TransactionSource
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpDirectDebit, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpStandingOrder, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpCreditCardAuthority, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpPaymentPlanMaintenance, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpGiftAidDeclaration, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareServices.TraderPageType.tpGiveAsYouEarnEntry, "Source")
      If vSourceCode.Length = 0 Then vSourceCode = GetPageValue(CareNetServices.TraderPageType.tpLoans, "Source")
      If vSourceCode.Length = 0 Then
        'If at this stage the mailing code still hasn't been identified and the application is defined as requiring a mailing code, then prompt for a mailing code
        If mvTA.MailingCodeMandatory Then
          Dim vAPI As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptGetMailingCode, Nothing, Nothing)
          While vAPI.ShowDialog <> System.Windows.Forms.DialogResult.OK
            '  
          End While
          Dim vReturnList As ParameterList = vAPI.ReturnList
          pList("Mailing") = vReturnList("Mailing")
        End If
      Else
        'Now get the tyl attribute from the sources table
        Dim vList As New ParameterList(True)
        vList("Source") = vSourceCode
        pList("Mailing") = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtSources, vList).Item("ThankYouLetter").ToString
      End If
    Else
      pList("Mailing") = vMailing
    End If

  End Sub
  Private Function CheckBatchTotals() As Boolean
    'Check Batch Totals etc. to see if more transactions are to be entered
    Dim vMsg As New StringBuilder()
    Dim vTotalDiff As Double
    Dim vTransDiff As Integer
    Dim vResult As System.Windows.Forms.DialogResult = System.Windows.Forms.DialogResult.Yes

    If mvTA.EditExistingTransaction = True OrElse mvTA.ExistingAdjustmentTran Then
      vResult = System.Windows.Forms.DialogResult.No
    Else
      'New Transaction
      If mvTA.BatchLedApp Then mvTA.BatchInfo = New BatchInfo(mvTA.BatchNumber) 'Refresh
      If mvTA.BatchLedApp = True AndAlso mvTA.BatchNumber > 0 Then
        With mvTA.BatchInfo
          If .NumberOfEntries > 0 Then vTransDiff = (.NumberOfTransactions - .NumberOfEntries)
          If .BatchTotal > 0 Then vTotalDiff = (.TransactionTotal - .BatchTotal)
          If .BatchTotal > 0 OrElse .NumberOfEntries > 0 Then
            If vTransDiff <> 0 Then
              vMsg.AppendLine(InformationMessages.ImNumberTransactionsDoesNotMatch)
              vMsg.AppendLine()
            ElseIf .NumberOfEntries > 0 Then
              vMsg.AppendLine(InformationMessages.ImNumberOfTransactionsMatches)
              vMsg.AppendLine()
            End If
            If vTotalDiff <> 0 Then
              vMsg.AppendLine(InformationMessages.ImTransactionTotalDoesNotMatch)
              vMsg.AppendLine()
            ElseIf .BatchTotal > 0 Then
              vMsg.AppendLine(InformationMessages.ImTransactionTotalsMatch)
              vMsg.AppendLine()
            End If
          End If

          If vTransDiff = 0 OrElse vTotalDiff = 0 Then
            If vTransDiff = 0 AndAlso .BatchTypeCode = "CV" AndAlso .NumberOfEntries = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.cv_max_number_of_vouchers, "60")) Then
              vResult = System.Windows.Forms.DialogResult.No
              If vMsg.Length > 0 Then ShowInformationMessage(vMsg.ToString)
            Else
              vMsg.AppendLine(QuestionMessages.QmAddMoreEntries)
              vResult = ShowQuestion(vMsg.ToString, MessageBoxButtons.YesNo)
            End If
          End If
        End With
      ElseIf mvTA.BatchNumber > 0 Then
        vResult = System.Windows.Forms.DialogResult.No
      End If
    End If
    Return (vResult = System.Windows.Forms.DialogResult.Yes)
  End Function

  Private Sub CMDActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFileName As String)
    mvCMDFileName = pFileName
  End Sub

  Private Sub CreateCMD(ByVal pDataSet As DataSet, ByVal plist As ParameterList)
    Dim vRow As DataRow = pDataSet.Tables("Result").Rows(0)
    If vRow.Table.Columns.Contains("CreateMailingDocument") AndAlso vRow.Item("CreateMailingDocument").ToString = "Y" Then
      Dim vShowParagraphs As DialogResult = System.Windows.Forms.DialogResult.Yes
      If vRow.Table.Columns.Contains("ContactWarningSuppressionsPrompt") AndAlso vRow.Table.Columns.Contains("WarningSuppressions") AndAlso vRow.Item("ContactWarningSuppressionsPrompt").ToString = "Y" AndAlso vRow.Item("WarningSuppressions").ToString.Length > 0 Then
        vShowParagraphs = ShowQuestion(QuestionMessages.QmWarningSuppressions, MessageBoxButtons.YesNo, vRow.Item("WarningSuppressions").ToString)
      End If
      If vShowParagraphs = System.Windows.Forms.DialogResult.Yes Then
        'Retrieve the matching paragraphs
        Dim vTransTable As DataTable = DataHelper.GetTableFromDataSet(pDataSet)
        Dim vCMDList As New ParameterList(True)
        AddCMDMailingCode(vCMDList)
        vCMDList.Item("ContactNumber") = plist("PayerContactNumber")
        vCMDList.Item("ExistingTransaction") = CBoolYN(mvTA.EditExistingTransaction)
        vCMDList.Item("NewPayerContact") = "N"
        If mvNewContacts IsNot Nothing AndAlso mvNewContacts.ContainsKey(vCMDList.Item("ContactNumber")) Then vCMDList.Item("NewPayerContact") = "Y"
        If vTransTable.Columns.Contains("BatchNumber") AndAlso IntegerValue(vTransTable.Rows(0).Item("BatchNumber").ToString) > 0 Then vCMDList.Item("BatchNumber") = vTransTable.Rows(0).Item("BatchNumber").ToString
        If vTransTable.Columns.Contains("TransactionNumber") AndAlso IntegerValue(vTransTable.Rows(0).Item("TransactionNumber").ToString) > 0 Then vCMDList.Item("TransactionNumber") = vTransTable.Rows(0).Item("TransactionNumber").ToString
        If mvTA.PPNumbersCreated.Count > 0 Then
          vCMDList.Item("PaymentPlanNumber") = mvTA.PPNumbersCreated(0).ToString
          vCMDList.Item("PaymentPlanCreated") = "Y"
        Else
          If vRow.Table.Columns.Contains("PaymentPlanNumber") Then vCMDList.Item("PaymentPlanNumber") = vRow("PaymentPlanNumber").ToString
        End If
        If mvTA.DeclarationNumber > 0 Then vCMDList.IntegerValue("DeclarationNumber") = mvTA.DeclarationNumber
        If mvTA.AutoPaymentCreated Then vCMDList.Item("AutoPaymentCreated") = "Y"
        If mvTA.PaymentPlan IsNot Nothing Then vCMDList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
        Dim vCount As Integer
        Do
          Dim vCMDDataSet As DataSet = DataHelper.GetMailingDocumentParagraphs(vCMDList)
          'Display the matching paragraphs
          Dim vParagraphsTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
          If Not mvTA.BypassMailingParagraphs AndAlso vParagraphsTable IsNot Nothing AndAlso
             vParagraphsTable.Columns.Contains("DisplayParagraphs") AndAlso vParagraphsTable.Rows(0).Item("DisplayParagraphs").ToString = "Y" Then
            DocumentApplication = New WordApplication
            AddHandler DocumentApplication.ActionComplete, AddressOf CMDActionComplete
            Dim vTransDocumentType As frmTransactionDocument.TransactionDocumentTypes = frmTransactionDocument.TransactionDocumentTypes.tdtTransaction
            If vCount > 0 Then vTransDocumentType = frmTransactionDocument.TransactionDocumentTypes.tdtPaymentPlan
            Dim vForm As frmTransactionDocument = New frmTransactionDocument(vTransDocumentType, vCMDDataSet, vCMDList)
            vForm.ShowDialog()
            'need to edit the document when Edit is pressed
            vCMDDataSet = vForm.DataSet
          End If
          'Create the mailing document
          If DataHelper.GetTableFromDataSet(vCMDDataSet) IsNot Nothing Then
            If vCMDList.Contains("EarliestFulfilmentDate") = False Then vCMDList("EarliestFulfilmentDate") = AppValues.TodaysDate
            If mvCMDFileName.Length = 0 Then
              Dim vSelectedParagraphs As New StringBuilder
              Dim vCMDTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
              If vCMDTable IsNot Nothing Then
                For Each vCMDRow As DataRow In vCMDTable.Rows
                  If BooleanValue(vCMDRow.Item("Include").ToString) Then
                    If vSelectedParagraphs.Length > 0 Then vSelectedParagraphs.Append(",")
                    vSelectedParagraphs.Append(vCMDRow.Item("ParagraphNumber"))
                  End If
                Next
              End If
              If vSelectedParagraphs.Length = 0 Then
                vCMDList("SelectedParagraphs") = "0"
              Else
                vCMDList("SelectedParagraphs") = vSelectedParagraphs.ToString
              End If
            End If
            vCMDDataSet = DataHelper.AddContactMailingDocument(vCMDList)
            If mvCMDFileName.Length > 0 Then
              Dim vResultRow As DataRow = vCMDDataSet.Tables("Result").Rows(0)
              DataHelper.UpdateContactMailingDocumentFile(IntegerValue(vResultRow.Item("MailingDocumentNumber").ToString), mvCMDFileName)
            End If
          End If
          vCount += 1
          If mvTA.PaymentPlans.Count > 0 AndAlso mvTA.PaymentPlans.Count > vCount Then
            With vCMDList
              .Item("PaymentPlanNumber") = mvTA.PaymentPlans(vCount).PaymentPlanNumber.ToString
              If .Contains("BatchNumber") Then .Remove("BatchNumber")
              If .Contains("TransactionNumber") Then .Remove("TransactionNumber")
              If .Contains("DeclarationNumber") Then .Remove("DeclarationNumber")
            End With
          End If
        Loop While vCount < mvTA.PaymentPlans.Count
      End If
    End If
  End Sub

  Private Sub GenerateMailmerge(ByVal pMailmergeType As CareServices.TraderMailmergeType, ByVal pRow As DataRow)
    Dim vFileName As String = DataHelper.GetTempFile(".csv")
    Dim vList As New ParameterList(True)
    Dim vDocument As String
    Dim vPrintPreview As Boolean = False

    Select Case pMailmergeType
      Case CareServices.TraderMailmergeType.tmtInvoice
        If pRow.Table.Columns.Contains("PrintJobNumber") Then
          vList("PrintJobNumber") = pRow("PrintJobNumber").ToString
          vList("Company") = pRow("Company").ToString
          If pRow.Table.Columns.Contains("PrintPreview") Then
            vPrintPreview = BooleanValue(pRow("PrintPreview").ToString)
            vList("PrintPreview") = CBoolYN(vPrintPreview)
            If pRow.Table.Columns.Contains("InvoiceNumbersAdded") Then vList("InvoiceNumbersAdded") = pRow("InvoiceNumbersAdded").ToString
          End If
        Else
          vList.IntegerValue("FromInvoiceNumber") = 0
          vList.IntegerValue("ToInvoiceNumber") = 0
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpBatchInvoiceProduction Then
            vList("Company") = mvCurrentPage.EditPanel.GetValue("Company")
            Dim vFromDate As DateTimePicker = TryCast(FindControl(mvCurrentPage.EditPanel, "FromDate", False), DateTimePicker)
            If Not vFromDate Is Nothing Then
              vList("FromDate") = mvCurrentPage.EditPanel.GetValue("FromDate")
              vList("ToDate") = mvCurrentPage.EditPanel.GetValue("ToDate")
            End If
            If mvCurrentPage.EditPanel.GetValue("InvoiceNumber").Length > 0 Then
              vList.IntegerValue("FromInvoiceNumber") = CInt(mvCurrentPage.EditPanel.GetValue("InvoiceNumber"))
              vList.IntegerValue("ToInvoiceNumber") = CInt(mvCurrentPage.EditPanel.GetValue("InvoiceNumber2"))
            End If
            Dim vPartPaid As CheckBox = TryCast(FindControl(mvCurrentPage.EditPanel, "PartPaidOnly", False), CheckBox)
            If vPartPaid IsNot Nothing Then vList("PartPaidOnly") = CBoolYN(vPartPaid.Checked)
            Dim vRunType As String = ""
            If GetInvoicePrintRunType(mvCurrentPage.EditPanel, vRunType) Then vList("RunType") = vRunType
            EPL_ShowMessage(mvCurrentPage.EditPanel, InformationMessages.ImGeneratingInvoices)
          Else
            vList("BatchNumber") = pRow.Item("BatchNumber").ToString
            vList("TransactionNumber") = pRow.Item("TransactionNumber").ToString
            vList("Company") = mvTA.CSCompany
            vList("InstantPrint") = "Y"
            If mvTA.EditExistingTransaction Then vList("ExistingTransaction") = "Y"
          End If
        End If
        vList.IntegerValue("TraderApplication") = mvTA.ApplicationNumber
        vDocument = mvTA.InvoiceDocument
      Case CareNetServices.TraderMailmergeType.tmtPaymentPlan
        vList("PaymentPlanNumber") = pRow("PaymentPlanNumber").ToString
        vDocument = mvTA.PaymentPlanDocument
      Case Else
        vList("BatchNumber") = pRow.Item("BatchNumber").ToString
        vList("TransactionNumber") = pRow.Item("TransactionNumber").ToString
        If pMailmergeType = CareServices.TraderMailmergeType.tmtReceipt Then
          vDocument = mvTA.ReceiptDocument
        Else   'CareServices.TraderMailmergeType.tmtProvisionalCash
          vDocument = mvTA.ProvisionalCashTransactionDocument
        End If
    End Select

    If DataHelper.GetTraderMailingFile(pMailmergeType, vList, vFileName) Then
      'We have the mailmerge file, now need to perform the mailmerge
      If ((pMailmergeType = CareServices.TraderMailmergeType.tmtInvoice And mvTA.InvoiceDocument.Length > 0) OrElse (pMailmergeType <> CareServices.TraderMailmergeType.tmtInvoice)) Then
        'This will run Invoice Mailmerge with or without a document
        'And Receipt / ProvisionalCash only if there is a document
        vList = New ParameterList(True)
        vList("StandardDocument") = vDocument
        Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vList).Rows(0)
        mvExtApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
        If vPrintPreview Then
          mvExtApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vFileName, False, True, False, True)
        Else
          mvExtApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vFileName, True)
          mvExtApplication.CloseExternalApplication(False)
        End If
      ElseIf pMailmergeType = CareServices.TraderMailmergeType.tmtInvoice Then
        'If no Invoice document then inform user
        ShowInformationMessage(InformationMessages.ImManualInvoiceMerge, vFileName)
      End If
    End If
  End Sub

  Private Sub ProcessFinishedReturns(ByVal pDataSet As DataSet, ByVal pList As ParameterList)
    Dim vRow As DataRow = pDataSet.Tables("Result").Rows(0)
    If (vRow.Table.Columns.Contains("PrintInvoice") = True AndAlso vRow.Item("PrintInvoice").ToString = "Y") Then
      'Set InvoiceNumbers etc. and perform the mailmerge
      Dim vPrintInvoice As DialogResult = System.Windows.Forms.DialogResult.Yes
      If mvCurrentPage.PageType <> CareServices.TraderPageType.tpBatchInvoiceProduction Then vPrintInvoice = ShowQuestion(QuestionMessages.QmCreateInvoice, MessageBoxButtons.YesNo, (IIf(mvTA.TransactionType = "CRDN", "credit note", "invoice")).ToString)
      If vPrintInvoice = System.Windows.Forms.DialogResult.Yes Then
        GenerateMailmerge(CareServices.TraderMailmergeType.tmtInvoice, vRow)
      End If
    End If
    If pList.Contains("PrintJobNumber") AndAlso pList("PrintJobNumber").Length > 0 Then
      GenerateMailmerge(CareServices.TraderMailmergeType.tmtInvoice, vRow)
    End If
    If mvTA.ReceiptDocument.Length > 0 AndAlso (vRow.Table.Columns.Contains("PrintReceipt") = True AndAlso vRow.Item("PrintReceipt").ToString = "Y") Then
      'Only print Receipt if we have a Receipt document
      If ShowQuestion(QuestionMessages.QmPrintReceipt, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        GenerateMailmerge(CareServices.TraderMailmergeType.tmtReceipt, vRow)
      End If
    End If
    If mvTA.ProvisionalCashTransactionDocument.Length > 0 AndAlso (vRow.Table.Columns.Contains("PrintProvisionalCashDoc") = True AndAlso vRow.Item("PrintProvisionalCashDoc").ToString = "Y") Then
      'Only print ProvisionalCashDocument if we have a ProvisionalCashDocument
      If ShowQuestion(QuestionMessages.QmPrintProvisionalCashDoc, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        GenerateMailmerge(CareServices.TraderMailmergeType.tmtProvisionalCash, vRow)
      End If
    End If
    'BR16458 - Gift Aid Message
    If (vRow.Table.Columns.Contains("OneOffGADMessage") = True AndAlso vRow.Item("OneOffGADMessage").ToString = "Y") Then
      ShowInformationMessage(InformationMessages.ImGiftAidDeclarationReAnalysis)
    End If
    'Create the Contact Mailing Document
    CreateCMD(pDataSet, pList)
  End Sub

  Private Sub ProcessNextMove(ByVal pType As CareServices.TraderProcessDataTypes, ByVal pDataSet As DataSet)
    'called from the end of processdata
    Dim vRow As DataRow = pDataSet.Tables("Result").Rows(0)
    If vRow.Table.Columns.Contains("NextPageCode") Then
      mvNextPageCode = vRow("NextPageCode").ToString
    End If
    Dim vPaymentPlanNumber As Integer = 0
    If pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
      Dim vMsg As New StringBuilder
      If vRow.Table.Columns.Contains("PaymentPlanNumber") Then
        vPaymentPlanNumber = IntegerValue(vRow("PaymentPlanNumber").ToString)
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImPPCreated, vPaymentPlanNumber.ToString))
        mvTA.SetPaymentPlanCreated(vPaymentPlanNumber)
      End If
      If vRow.Table.Columns.Contains("MemberNumber") Then
        If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImMemberCreated, vRow("MemberNumber").ToString))
      End If
      If vRow.Table.Columns.Contains("LoanNumber") Then
        If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImLoanCreated, vRow("LoanNumber").ToString))
      End If
      If vRow.Table.Columns.Contains("DirectDebitNumber") Then
        If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImDDCreated, vRow("DirectDebitNumber").ToString))
        If vRow.Table.Columns.Contains("Reference") AndAlso vRow("DirectDebitNumber").ToString <> vRow("Reference").ToString Then
          vMsg.AppendLine()
          vMsg.AppendLine(GetInformationMessage(InformationMessages.ImReference, vRow("Reference").ToString))
        End If
        mvTA.AutoPaymentCreated = True
      End If
      If vRow.Table.Columns.Contains("BankersOrderNumber") Then
        If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImSOCreated, vRow("BankersOrderNumber").ToString))
        mvTA.AutoPaymentCreated = True
      End If
      If vRow.Table.Columns.Contains("CardAuthorityNumber") Then
        If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImCCCACreated, vRow("CardAuthorityNumber").ToString))
        mvTA.AutoPaymentCreated = True
      End If
      If mvTA.TransactionType = "MEMC" AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipPayer Then
        vMsg.AppendLine(InformationMessages.ImCMTMemberChanged)
        If vRow.Table.Columns.Contains("BatchNumber") Then
          vMsg.AppendLine(GetInformationMessage(InformationMessages.ImCMTInAdvanceBatch, vRow("BatchNumber").ToString))
        End If
      End If
      If vRow.Table.Columns.Contains("PurchaseOrderNumber") AndAlso IntegerValue(vRow("PurchaseOrderNumber")) > 0 Then
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImPOCreated, vRow("PurchaseOrderNumber").ToString))
      End If
      If vRow.Table.Columns.Contains("PurchaseInvoiceNumber") AndAlso IntegerValue(vRow("PurchaseInvoiceNumber")) > 0 Then
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImPICreated, vRow("PurchaseInvoiceNumber").ToString))
      End If
      If vRow.Table.Columns.Contains("TotalCheques") Then
        vMsg.AppendLine(GetInformationMessage(InformationMessages.ImTotalNumberOfChequesAllocated, vRow("TotalCheques").ToString))
      End If
      Dim vWarningMessage As New StringBuilder
      If vRow.Table.Columns.Contains("WarningMessage") Then
        vWarningMessage.AppendLine(vRow("WarningMessage").ToString)
      End If
      If vMsg.ToString.Length > 0 Then ShowInformationMessage(vMsg.ToString)
      If vWarningMessage.ToString.Length > 0 Then ShowWarningMessage(vWarningMessage.ToString)
    End If

    If mvTA.PaymentPlanDocument.Length > 0 AndAlso vPaymentPlanNumber > 0 Then
      'Only print PaymentPlanDocument if we have a PaymentPlanDocument
      If ShowQuestion(QuestionMessages.QmPrintPaymentPlanDocument, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        GenerateMailmerge(CareNetServices.TraderMailmergeType.tmtPaymentPlan, vRow)
      End If
    End If

    Dim vPageType As CareServices.TraderPageType = CType(vRow.Item("NextPageType"), CareServices.TraderPageType)
    If vPageType = CareServices.TraderPageType.tpNone Then
      If pType <> CareServices.TraderProcessDataTypes.tpdtFinished AndAlso pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
        ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, False)
      Else
        If pType <> CareServices.TraderProcessDataTypes.tpdtCancelTransaction Then
          'We are at the end of the transaction so show the reference and setup for the next transaction
          If mvTA.EditExistingTransaction = False Then
            If vRow.Table.Columns.Contains("BatchNumber") AndAlso vRow.Table.Columns.Contains("TransactionNumber") AndAlso mvTA.ShowTransactionReference Then
              ShowInformationMessage(InformationMessages.ImTransactionReference, vRow("BatchNumber").ToString, vRow("TransactionNumber").ToString)
            End If
          Else
            'update the transaction total on the batch info, so that it can be picked up by the transactions form
            If vRow.Table.Columns.Contains("BatchNumber") AndAlso vRow.Table.Columns.Contains("TransactionNumber") AndAlso mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.None Then
              If mvTA.BatchInfo IsNot Nothing Then
                mvTA.BatchInfo.UpdateTransactionTotal(mvTA.TransactionAmount - mvTA.OriginalTransactionAmount)
              End If
            End If
          End If

          If (mvTA.ApplicationStartPoint <> TraderApplication.TraderApplicationStartPoint.taspRightMouse) AndAlso CheckBatchTotals() Then
            If mvTA.BatchLedApp Then
              Dim vContactInfo As ContactInfo = Nothing
              Dim vControl As Control = FindControl(mvTraderPages(CareServices.TraderPageType.tpTransactionDetails.ToString).EditPanel, "ContactNumber", False)
              If vControl IsNot Nothing AndAlso TypeOf (vControl) Is TextLookupBox Then vContactInfo = CType(vControl, TextLookupBox).ContactInfo
              If vContactInfo Is Nothing OrElse vContactInfo.ContactNumber = 0 Then
                vContactInfo = New ContactInfo(mvTA.PayerContactNumber)
              End If
              Dim vAddressLine As String = String.Empty
              vControl = FindControl(mvTraderPages(CareServices.TraderPageType.tpTransactionDetails.ToString).EditPanel, "AddressNumber", False)
              If vControl IsNot Nothing AndAlso TypeOf (vControl) Is TextLookupBox Then vAddressLine = CType(vControl, TextLookupBox).Description
              mvTransactionsForm.NewTransactionAdded(vContactInfo, mvTA.CalcCurrencyAmount(mvTA.TransactionAmount, True), mvTA.TransactionAmount, mvTA.TransactionDate, vAddressLine)
            End If
            SetPage(mvTA.MainPageType)
            mvTA.NewTransaction()
            For Each vPage As TraderPage In mvTraderPages
              If vPage.PageType <> CareNetServices.TraderPageType.tpStatementList Then
                vPage.EditPanel.Clear()
              End If
              vPage.DefaultsSet = False
            Next
            mvNewContacts = Nothing
            mvCMDFileName = ""
            If mvTA.ApplicationType = ApplicationTypes.atConversion AndAlso ((mvTA.ConversionShowPPD = False And mvTA.PayPlanConvMaintenance = False) _
               Or ((mvTA.ConversionShowPPD = True Or mvTA.PayPlanConvMaintenance = True) And mvMembConv)) AndAlso mvTA.PaymentPlan IsNot Nothing Then
              'reinitialise the payment plan class to get the latest changes just done.
              Dim vPPNo As Integer = mvTA.PaymentPlan.PaymentPlanNumber
              mvTA.PaymentPlan = Nothing
              mvTA.PaymentPlan = New PaymentPlanInfo(vPPNo)
            Else
              mvTA.PaymentPlan = Nothing
            End If
            ProcessData(CareServices.TraderProcessDataTypes.tpdtFirstPage, False)
          Else
            'Finished editing existing transaction so close
            'clear the trader app so that the cancel (called on form close) does not pickup any existing analysis to cancel
            mvTA.NewTransaction()
            mvConfirmCancel = False
            Me.Close()
          End If
          mvEventWLPriceZeroed = False
        End If
      End If
    Else
      SetPage(vPageType)
      If Not mvCurrentPage.Menu AndAlso Not mvCurrentPage.SummaryPage AndAlso
        (pType = CareServices.TraderProcessDataTypes.tpdtNextPage OrElse
         pType = CareServices.TraderProcessDataTypes.tpdtFinished OrElse
         pType = CareServices.TraderProcessDataTypes.tpdtAmendMemberSummary OrElse
           (pType = CareServices.TraderProcessDataTypes.tpdtFirstPage And
             (vPageType = CareServices.TraderPageType.tpTransactionDetails OrElse
              vPageType = CareServices.TraderPageType.tpContactSelection OrElse
              vPageType = CareServices.TraderPageType.tpPurchaseOrderDetails OrElse
              vPageType = CareServices.TraderPageType.tpPurchaseInvoiceDetails OrElse
              vPageType = CareServices.TraderPageType.tpPurchaseOrderCancellation OrElse
              vPageType = CareServices.TraderPageType.tpChequeNumberAllocation OrElse
              vPageType = CareServices.TraderPageType.tpChequeReconciliation OrElse
              vPageType = CareServices.TraderPageType.tpGiveAsYouEarn OrElse
              vPageType = CareServices.TraderPageType.tpBatchInvoiceProduction OrElse
              vPageType = CareServices.TraderPageType.tpPostTaxPGPayment))) Then
        If mvCurrentPage.DefaultsSet = False Then
          SetDefaults(vRow)
        Else
          ResetValuesFromReturns(vRow)
        End If
      ElseIf mvCurrentPage.SummaryPage Then
        'mark the page as visited
        mvCurrentPage.DefaultsSet = True
      Else
        If vPageType = CareNetServices.TraderPageType.tpPaymentPlanFromUnbalanceTransaction Then
          If mvCurrentPage.DefaultsSet = False Then SetDefaults(vRow)
        End If
      End If
      If pType = CareServices.TraderProcessDataTypes.tpdtEditTransaction Then SetValuesForExistingTransaction(pDataSet)
      If pType = CareServices.TraderProcessDataTypes.tpdtEditAnalysisLine Then
        'Set values on each page
        SetPageValuesForEditing()
        mvTA.TransactionLines -= 1
      ElseIf vPageType = CareServices.TraderPageType.tpTransactionAnalysisSummary OrElse vPageType = CareServices.TraderPageType.tpPaymentPlanSummary Then
        'Reset to show that we are no longer editing an analysis line
        mvTA.EditLineNumber = 0
        mvTA.EditPPDetailNumber = 0
        mvTA.EditPPDSubscriptionNumber = 0
        mvTA.PPDMemberOrPayer = ""
        If pType = CareServices.TraderProcessDataTypes.tpdtFinished AndAlso mvTA.MembersDataSet.Tables.Contains("DataRow") Then
          mvTA.MembersDataSet.Tables.Remove("DataRow")
        End If
      ElseIf vPageType = CareServices.TraderPageType.tpMembershipMembersSummary Then
        'Reset to show that we are no longer editing a member line
        mvTA.EditMemberLineNumber = 0
      End If
      If vPageType = CareServices.TraderPageType.tpMembershipMembersSummary Then SetMembersGridButtons(mvMembersDGR.CurrentRow)
      If pType = CareServices.TraderProcessDataTypes.tpdtNextPage Or pType = CareServices.TraderProcessDataTypes.tpdtFinished Then
        If vPageType = CareServices.TraderPageType.tpTransactionAnalysisSummary AndAlso (mvTA.EditExistingTransaction = False OrElse mvTA.FinancialAdjustment <> BatchInfo.AdjustmentTypes.None) Then
          'We are going to TransactionAnalysisSummary page, now see if we in fact need to skip it
          If (mvTA.TransactionAmount = mvTA.CurrentLineTotal And mvTA.ConfirmAnalysis = False) _
          OrElse (mvTA.AutoSetAmount = True And mvTA.ConfirmAnalysis = False) Then
            'Process Finished button
            ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
          End If
        End If

        If vPageType = CareServices.TraderPageType.tpPaymentPlanSummary AndAlso mvTA.ApplicationType <> ApplicationTypes.atMaintenance _
        AndAlso Not mvTA.ConfirmPayPlanDetails AndAlso mvTA.PayPlanConvMaintenance = False Then  '"MAINT"
          'If not in Maintenance Type App then skip summary if totals are the same
          If mvTA.PPBalance = mvTA.CurrentPPDLineTotal Then      'Totals are the same
            'Process Finished button
            ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
          End If
        End If

        If vPageType = CareServices.TraderPageType.tpOutstandingScheduledPayments Then
          'We are going to OutstandingScheduledPayments page, now see if we in fact need to skip it
          'Process Finished button
          If DoubleValue(mvCurrentPage.EditPanel.GetValue("AmountDue").ToString) > 0 AndAlso DoubleValue(mvCurrentPage.EditPanel.GetValue("AmountOutstanding").ToString) = 0 Then
            ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
          End If
        End If

        If vPageType = CareServices.TraderPageType.tpMembershipMembersSummary Then
          If (mvTA.MemberCount = mvTA.CurrentMembers) AndAlso cmdNext.Enabled = True AndAlso mvTA.TransactionType = "MEMC" Then
            'Process Next button
            ProcessData(CareServices.TraderProcessDataTypes.tpdtNextPage, True)
          End If
        End If

        If vPageType = CareServices.TraderPageType.tpMembershipPayer Then
          'Process Finished button
          ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
        End If

        If (vPageType = CareServices.TraderPageType.tpPurchaseOrderSummary OrElse vPageType = CareServices.TraderPageType.tpPurchaseInvoiceSummary) AndAlso mvTA.ApplicationType <> ApplicationTypes.atMaintenance Then  '"MAINT"
          'If not in Maintenance Type App then skip summary if totals are the same
          If mvTA.PPBalance = mvTA.CurrentPPDLineTotal AndAlso mvTA.PurchaseOrderNumber = 0 Then      'Totals are the same and we are not amending
            'Process Finished button
            ProcessData(CareServices.TraderProcessDataTypes.tpdtFinished, True)
          End If
        End If
      End If
      SetAnalysisEditable(CInt(IIf((mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count > 0), 0, -1)))
      SetPPDEditable(CInt(IIf((mvTA.PPDDataSet.Tables.Contains("DataRow") AndAlso mvTA.PPDDataSet.Tables("DataRow").Rows.Count > 0), 1, -1)))
      If vPageType = CareNetServices.TraderPageType.tpTransactionDetails AndAlso mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count() >= mvCurrentRow + 1 Then
        Dim vSourceLookupBox As TextLookupBox = mvCurrentPage.EditPanel.FindTextLookupBox("Source", False)
        If vSourceLookupBox IsNot Nothing Then
          vSourceLookupBox.Text = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow).Item("Source").ToString()
        End If
      End If
      If (vPageType = CareNetServices.TraderPageType.tpTransactionDetails And pType = CareNetServices.TraderProcessDataTypes.tpdtPreviousPage) _
        Or (vPageType = CareNetServices.TraderPageType.tpComments) _
        Or (vPageType = CareNetServices.TraderPageType.tpTransactionDetails And mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation) Then
        Dim vNotes As TextBox = mvCurrentPage.EditPanel.FindPanelControl(Of TextBox)("Notes", False)
        If vNotes IsNot Nothing Then
          vNotes.Text = mvTA.TransactionNote
        End If
      End If
    End If
  End Sub

  Private Sub RefreshCardSet(ByVal pRefreshType As RefreshTypes, ByVal pContactNumber As Integer, ByVal pRow As DataRow)
    Select Case pRefreshType
      Case RefreshTypes.rtEventBooking
        If pRow.Table.Columns.Contains("EventBookingNumber") AndAlso pRow("EventBookingNumber").ToString.Length > 0 Then
          Dim vEventNumber As Integer = IntegerValue(pRow("EventBookingNumber").ToString) \ 10000
          If vEventNumber > 0 Then
            MainHelper.RefreshEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookings, vEventNumber)
            MainHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, pContactNumber)
          End If
        End If
    End Select
  End Sub

  Private Sub ResetValuesFromReturns(ByVal pRow As DataRow)
    '#**************************************************************************************************************************
    'This method should be used for values that need to be reset with values returned from the server after a Previous-Next
    'When we do a Next for the first time, it should go to SetDefaults
    'this method has been designed to handle special cases that may need to update selective fields on a page following a previous next
    'pRow is the Datarow that is returned from the server
    '#**************************************************************************************************************************
    If mvCurrentPage.DefaultsSet Then
      Dim vEPL As EditPanel = mvCurrentPage.EditPanel
      With pRow
        Select Case mvCurrentPage.PageType
          Case CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpCreditCardAuthority
            If .Table.Columns.Contains("StartDate") Then vEPL.SetValue("StartDate", .Item("StartDate").ToString)
          Case CareServices.TraderPageType.tpMembership
            Dim vList As ParameterList = GetMembershipPrices(vEPL)
            If vList.Contains("PrimaryRate") Then
              SetValueRaiseChanged(vEPL, "Rate", vList("PrimaryRate"))
            ElseIf vEPL.GetValue("MembershipType").Length > 0 Then
              If mvTA.TransactionType <> "MEMC" OrElse (mvTA.TransactionType = "MEMC" AndAlso (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_renew_at_same_rate) = True OrElse mvTA.PaymentPlan.DetermineMembershipPeriod() <> PaymentPlanInfo.MembershipPeriodTypes.mptSubsequentPeriod)) Then
                SetValueRaiseChanged(vEPL, "Rate", vEPL.FindTextLookupBox("MembershipType").GetDataRow.Item("FirstPeriodsRate").ToString)
              Else
                SetValueRaiseChanged(vEPL, "Rate", vEPL.FindTextLookupBox("MembershipType").GetDataRow.Item("SubsequentPeriodsRate").ToString)
              End If
            End If
          Case CareServices.TraderPageType.tpPaymentPlanDetails
            If .Table.Columns.Contains("OrderTerm") Then vEPL.SetValue("OrderTerm", .Item("OrderTerm").ToString)
            If .Table.Columns.Contains("OrderDate") Then
              vEPL.SetValue("OrderDate", .Item("OrderDate").ToString)
              'Raise change handler to get the updated price as the date might have changed
              If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.recalculate_membership_balance) AndAlso mvTA.mvChangedStartDate.Length > 0 Then
                vEPL.SetValue("OrderDate", mvTA.mvChangedStartDate)
                SetValueRaiseChanged(vEPL, "OrderDate", vEPL.GetValue("OrderDate"))
              Else
                vEPL.SetValue("OrderDate", .Item("OrderDate").ToString)
              End If
            End If
            If .Table.Columns.Contains("ExpiryDate") Then vEPL.SetValue("ExpiryDate", .Item("ExpiryDate").ToString)
            If .Table.Columns.Contains("PaymentFrequency") Then vEPL.SetValue("PaymentFrequency", .Item("PaymentFrequency").ToString)
            If .Table.Columns.Contains("ReasonForDespatch") Then vEPL.SetValue("ReasonForDespatch", .Item("ReasonForDespatch").ToString)
            If .Table.Columns.Contains("Source") Then vEPL.SetValue("Source", .Item("Source").ToString)
          Case CareServices.TraderPageType.tpPaymentPlanProducts
            SetValueRaiseChanged(vEPL, "Product", vEPL.GetValue("Product"))
          Case CareNetServices.TraderPageType.tpPostageAndPacking
            If pRow.Table.Columns.Contains("TransactionAmount") Then
              Dim vOldAmount As Double = DoubleValue(vEPL.GetValue("Amount"))
              Dim vNewAmount As Double = DoubleValue(pRow("TransactionAmount").ToString)
              If vOldAmount.Equals(vNewAmount) = False Then
                'Transaction Amount has changed so set the new amount
                vEPL.SetValue("Amount", vNewAmount.ToString("0.00"), True)
                If Not String.IsNullOrEmpty(vEPL.GetValue("Percentage")) Then
                  SetValueRaiseChanged(vEPL, "Percentage", vEPL.GetValue("Percentage"))   'Plus re-calculate the percentage amount
                End If
              End If
            End If
          Case CareServices.TraderPageType.tpTransactionDetails
            vEPL.EnableControl("EligibleForGiftAid", Not (mvTA.FinancialAdjustment <> BatchInfo.AdjustmentTypes.None AndAlso mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.GIKConfirmation))
        End Select
      End With
    End If
  End Sub

  Private Function AddStockMovement(ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer, ByVal pReverseStock As Boolean, ByVal pPlaceonBackOrder As Boolean) As Boolean
    Dim vBusyCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      Dim vLines As Integer = mvTA.TransactionLines
      Dim vValid As Boolean = True

      If mvTA.StockSales = True Then
        vList("Product") = pProductCode
        vList("Warehouse") = pWarehouseCode
        vList("Quantity") = pQuantity.ToString
        vList.IntegerValue("StockTransactionID") = mvTA.StockTransactionID
        If pReverseStock Then vList("StockReversal") = "Y"
        If mvTA.EditExistingTransaction = True AndAlso pReverseStock = True AndAlso pPlaceonBackOrder = False Then
          'Add Batch & Transaction Numbers
          vList.IntegerValue("BatchNumber") = mvTA.BatchNumber
          vList.IntegerValue("TransactionNumber") = mvTA.TransactionNumber
          vList("ExistingTransaction") = "Y"
          vLines -= 1
        End If
        vList.IntegerValue("TransactionLines") = vLines
        Dim vReturnList As ParameterList = DataHelper.ProcessStockMovement(CareServices.ProcessStockMovementType.psmtAdd, vList)

        If vReturnList IsNot Nothing Then
          'Store Stock values
          mvTA.SetStockTransactionValues(vReturnList.IntegerValue("StockTransactionID"), vReturnList.IntegerValue("StockIssued"), pProductCode, pWarehouseCode, pQuantity)
          vList.IntegerValue("StockTransactionID") = mvTA.StockTransactionID

          Dim vStockIssued As Integer = mvTA.StockIssued
          If vReturnList.Contains("StockMessage") Then
            'There was insufficient Stock so may need to place on Back Order
            If ShowQuestion(vReturnList("StockMessage"), MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
              'Reverse out the Stock already issued and set as invalid
              vValid = False
              vList("StockReversal") = "Y"
              vList.IntegerValue("Quantity") = mvTA.StockIssued
              vReturnList = DataHelper.ProcessStockMovement(CareServices.ProcessStockMovementType.psmtAdd, vList)
            Else
              'Add a Back Order StockMovement
              vList("BackOrder") = "Y"
              vList.IntegerValue("Quantity") = 0
              vReturnList = DataHelper.ProcessStockMovement(CareServices.ProcessStockMovementType.psmtAdd, vList)
            End If
            If vReturnList IsNot Nothing Then mvTA.SetStockTransactionValues(vReturnList.IntegerValue("StockTransactionID"), vStockIssued, pProductCode, pWarehouseCode, pQuantity)
          ElseIf pReverseStock = True AndAlso pPlaceonBackOrder = True Then
            'Now that Stock has been reversed, put it on Back Order (this is used when delivery date is in the future)
            vList("StockReversal") = "N"
            vList("BackOrder") = "Y"
            vList.IntegerValue("Quantity") = 0
            vReturnList = DataHelper.ProcessStockMovement(CareServices.ProcessStockMovementType.psmtAdd, vList)
            If vReturnList IsNot Nothing Then mvTA.SetStockTransactionValues(vReturnList.IntegerValue("StockTransactionID"), vReturnList.IntegerValue("StockIssued"), pProductCode, pWarehouseCode, pQuantity)
          End If
        End If

        Return vValid

      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try

  End Function

  Private Function AddStockMovement(ByVal pProductCode As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer) As Boolean
    Return AddStockMovement(pProductCode, pWarehouseCode, pQuantity, False, False)
  End Function

  Private Sub DeleteStockMovement(ByVal pUpdateStockLevels As Boolean)
    Dim vBusyCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("StockTransactionID") = mvTA.StockTransactionID
      vList("ExistingTransaction") = CBoolYN(mvTA.EditExistingTransaction)
      vList("UpdateStockLevel") = CBoolYN(pUpdateStockLevels)

      DataHelper.ProcessStockMovement(CareServices.ProcessStockMovementType.psmtDeleteAll, vList)
      mvTA.StockSales = False

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub MaintainContactAddresses(ByVal pContactNumber As Integer, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pParentForm As MaintenanceParentForm, ByRef pList As ParameterList, ByVal pShowDialog As Boolean)
    Dim vContactInfo As New ContactInfo(pContactNumber)
    Dim vAddressNumber As Integer = CInt(pList("AddressNumber"))
    If pList.Contains("AddressNumber") Then pList.Remove("AddressNumber")
    Dim vDataSet As DataSet = DataHelper.GetContactData(pType, pContactNumber, pList)
    Dim vRowNumber As Integer = 0
    For Each vDataRow As DataRow In vDataSet.Tables("DataRow").Rows
      If CInt(vDataRow.Item("AddressNumber")) = vAddressNumber Then Exit For
      vRowNumber += 1
    Next
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(pParentForm, vContactInfo, pType, vDataSet, True, vRowNumber, CareServices.XMLMaintenanceControlTypes.xmctNone, pList, pShowDialog, mvNonFinancialBatchNumber, mvNonFinancialTransactionNumber)
    If pShowDialog Then
      vForm.ShowDialog(pParentForm)
      pList("TraderAddressUpdated") = "Y"
    Else
      vForm.Show()
    End If
    If Not vForm.ReturnList Is Nothing AndAlso vForm.ReturnList.Contains("AddressNumber") Then
      pList("AddressNumber") = vForm.ReturnList("AddressNumber")
    Else
      pList("AddressNumber") = vAddressNumber.ToString
    End If
  End Sub

  Private Function ValidateDefaultContactAccount(ByVal pRow As Integer) As Boolean
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactBankAccounts, IntegerValue(mvPPADGR.GetValue(pRow, "PayeeContactNumber")))
    If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
      For Each vDataRow As DataRow In vDataSet.Tables("DataRow").Rows
        If BooleanValue(vDataRow("DefaultAccount").ToString) AndAlso Not BooleanValue(vDataRow("HistoryOnly").ToString) Then
          Return True
        End If
      Next
    End If
    Return False
  End Function

  Private Sub AddServiceBooking(ByVal pList As ParameterList)
    Dim vResult As ParameterList
    pList("ServiceBookingCredits") = CBoolYN(mvTA.ServiceBookingCredits)
    pList("ConfirmSBDuration") = CBoolYN(mvTA.ConfirmSBDuration)
    pList("ConfirmSBShortStay") = CBoolYN(mvTA.ConfirmSBShortStay)
    pList("ConfirmCalendarConflict") = CBoolYN(mvTA.ConfirmCalendarConflict)
    pList("NewQuantity") = mvTA.SBNewQuantity.ToString
    vResult = DataHelper.AddServiceBooking(pList)
    mvServiceBookingNumber = IntegerValue(vResult("ServiceBookingNumber"))
    pList("ServiceBookingNumber") = mvServiceBookingNumber.ToString
  End Sub

  Private Sub mvCMTOldPPD_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    'Combo value has changed
    If pValue.Length = 1 Then
      Dim vCol As Integer = -1
      With mvCMTOldPPD
        If pCol = .GetColumn("CMTProrateCost") Then
          vCol = .GetColumn("CMTProrateCostCode")
        ElseIf pCol = .GetColumn("CMTExcessPaymentType") Then
          vCol = .GetColumn("CMTExcessPaymentTypeCode")
        End If
        If vCol >= 0 Then .SetValue(pRow, vCol, pValue)
      End With
    End If
  End Sub

  Private Sub mvCMTNewPPD_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String)
    'Combo value has changed
    If pValue.Length = 1 Then
      Dim vCol As Integer = -1
      With mvCMTNewPPD
        If pCol = .GetColumn("CMTProrateCost") Then vCol = .GetColumn("CMTProrateCostCode")
        If vCol >= 0 Then .SetValue(pRow, vCol, pValue)
      End With
    End If
  End Sub
  ''' <summary>
  ''' The last point before a page is displayed.
  ''' </summary>
  ''' <remarks>If you need to make changes to the data on the first page of a trader application as a result of navigating backwards do it here</remarks>
  Private Sub PreProcessPage()
    ' Cash Batch Maintenance - Reflect change to Source on the PRD Page in the TRD Page, if the Analysis dataset contains data the we have got here by navingating backwards from TAS.
    ' BR19597
    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpTransactionDetails
        If mvTA.AnalysisDataSet.Tables.Contains("DataRow") AndAlso mvTA.AnalysisDataSet.Tables("DataRow").Rows.Count() >= mvCurrentRow + 1 Then
          Dim vSourceLookupBox As TextLookupBox = mvCurrentPage.EditPanel.FindTextLookupBox("Source", False)
          If vSourceLookupBox IsNot Nothing Then
            vSourceLookupBox.Text = mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow).Item("Source").ToString()
          End If
        End If
    End Select
  End Sub
#End Region

  Private Property CardAuthoriser As WebBasedCardAuthoriser = Nothing

End Class
