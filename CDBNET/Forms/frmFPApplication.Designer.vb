<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFPApplication
  Inherits ThemedForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFPApplication))
    Me.tbpCurrency = New System.Windows.Forms.TabPage()
    Me.pnlCurrencyBAs = New CDBNETCL.PanelEx()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.txtCurrency = New CDBNETCL.TextLookupBox()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.txtBatchType = New CDBNETCL.TextLookupBox()
    Me.Label8 = New System.Windows.Forms.Label()
    Me.txtBankAccount = New CDBNETCL.TextLookupBox()
    Me.cmdAdd = New System.Windows.Forms.Button()
    Me.cmdRemove = New System.Windows.Forms.Button()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.tbpRestrictions = New System.Windows.Forms.TabPage()
    Me.pnlRestrictions = New CDBNETCL.PanelEx()
    Me.chkMembersOnly = New System.Windows.Forms.CheckBox()
    Me.lblSalesGroup = New System.Windows.Forms.Label()
    Me.txtSalesGroup = New CDBNETCL.TextLookupBox()
    Me.lblAlbacsBankDetails = New System.Windows.Forms.Label()
    Me.cboAlbacsBankDetails = New System.Windows.Forms.ComboBox()
    Me.chkIncludeConfirmedTransaction = New System.Windows.Forms.CheckBox()
    Me.chkIncludeProvisionalTransaction = New System.Windows.Forms.CheckBox()
    Me.chkIncludeProvPaymentPlan = New System.Windows.Forms.CheckBox()
    Me.chkForceMailingCode = New System.Windows.Forms.CheckBox()
    Me.chkForceDistributionCode = New System.Windows.Forms.CheckBox()
    Me.chkSalesContactMandatory = New System.Windows.Forms.CheckBox()
    Me.tbpDocuments = New System.Windows.Forms.TabPage()
    Me.pnlDocuments = New CDBNETCL.PanelEx()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.txtInvoiceDoc = New CDBNETCL.TextLookupBox()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.txtReceiptDoc = New CDBNETCL.TextLookupBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.txtPayPlanDoc = New CDBNETCL.TextLookupBox()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.txtCreditStmtDoc = New CDBNETCL.TextLookupBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.txtProvCashDoc = New CDBNETCL.TextLookupBox()
    Me.tbpBank = New System.Windows.Forms.TabPage()
    Me.pnlBankAccount = New CDBNETCL.PanelEx()
    Me.lblCashAccount = New System.Windows.Forms.Label()
    Me.txtCashAccount = New CDBNETCL.TextLookupBox()
    Me.lblCreditCardAccount = New System.Windows.Forms.Label()
    Me.txtCreditCardAccount = New CDBNETCL.TextLookupBox()
    Me.lblDebitCardAccount = New System.Windows.Forms.Label()
    Me.txtDebitCardAccount = New CDBNETCL.TextLookupBox()
    Me.lblCreditSaleAccount = New System.Windows.Forms.Label()
    Me.txtCreditSaleAccount = New CDBNETCL.TextLookupBox()
    Me.lblStandingOrderAccount = New System.Windows.Forms.Label()
    Me.txtStandingOrderAccount = New CDBNETCL.TextLookupBox()
    Me.lblDirectDebitAccount = New System.Windows.Forms.Label()
    Me.txtDirectDebitAccount = New CDBNETCL.TextLookupBox()
    Me.lblCCAAccount = New System.Windows.Forms.Label()
    Me.txtCCAAccount = New CDBNETCL.TextLookupBox()
    Me.lblCAFAndVoucherAccount = New System.Windows.Forms.Label()
    Me.txtCAFAndVoucherAccount = New CDBNETCL.TextLookupBox()
    Me.tbpDefaults = New System.Windows.Forms.TabPage()
    Me.tbpCarriage = New System.Windows.Forms.TabControl()
    Me.tbpBatches = New System.Windows.Forms.TabPage()
    Me.pnlBatches = New CDBNETCL.PanelEx()
    Me.lblBatchCategory = New System.Windows.Forms.Label()
    Me.txtBatchCategory = New CDBNETCL.TextLookupBox()
    Me.lblBatchAnalysis = New System.Windows.Forms.Label()
    Me.txtBatchAnalysisCode = New CDBNETCL.TextLookupBox()
    Me.tbpTransactions = New System.Windows.Forms.TabPage()
    Me.pnlTransactions = New CDBNETCL.PanelEx()
    Me.lblCreditCategory = New System.Windows.Forms.Label()
    Me.txtCreditCategory = New CDBNETCL.TextLookupBox()
    Me.chkDefaultSourceFromLastMailing = New System.Windows.Forms.CheckBox()
    Me.lblSource = New System.Windows.Forms.Label()
    Me.txtSource = New CDBNETCL.TextLookupBox()
    Me.lblSalesPerson = New System.Windows.Forms.Label()
    Me.txtSalesPerson = New CDBNETCL.TextLookupBox()
    Me.lblLinkToCommunication = New System.Windows.Forms.Label()
    Me.txtLinkToCommunication = New CDBNETCL.TextLookupBox()
    Me.lblTransactionOrigin = New System.Windows.Forms.Label()
    Me.txtTransactionOrigin = New CDBNETCL.TextLookupBox()
    Me.chkLinkMALToEvent = New System.Windows.Forms.CheckBox()
    Me.chkLinkMALToService = New System.Windows.Forms.CheckBox()
    Me.chkLinkAnalysisLines = New System.Windows.Forms.CheckBox()
    Me.chkInvoicePrintPreview = New System.Windows.Forms.CheckBox()
    Me.tbpAnalysisDefaults = New System.Windows.Forms.TabPage()
    Me.pnlAnalysisSub = New CDBNETCL.PanelEx()
    Me.lblProduct = New System.Windows.Forms.Label()
    Me.txtProduct = New CDBNETCL.TextLookupBox()
    Me.lblRate = New System.Windows.Forms.Label()
    Me.txtRate = New CDBNETCL.TextLookupBox()
    Me.lblDonationProduct = New System.Windows.Forms.Label()
    Me.txtDonationProduct = New CDBNETCL.TextLookupBox()
    Me.lblAnotherRate = New System.Windows.Forms.Label()
    Me.txtDonationRate = New CDBNETCL.TextLookupBox()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
    Me.pnlCarriage = New CDBNETCL.PanelEx()
    Me.lblProductCarriage = New System.Windows.Forms.Label()
    Me.txtCarriageProduct = New CDBNETCL.TextLookupBox()
    Me.lblRateCarriage = New System.Windows.Forms.Label()
    Me.txtCarriageRate = New CDBNETCL.TextLookupBox()
    Me.lblPercentage = New System.Windows.Forms.Label()
    Me.txtPercentage = New System.Windows.Forms.TextBox()
    Me.tbpMembers = New System.Windows.Forms.TabPage()
    Me.pnlMembers = New CDBNETCL.PanelEx()
    Me.lblBranch = New System.Windows.Forms.Label()
    Me.txtBranch = New CDBNETCL.TextLookupBox()
    Me.tbpExams = New System.Windows.Forms.TabPage()
    Me.pnlExams = New CDBNETCL.PanelEx()
    Me.lblExamSession = New System.Windows.Forms.Label()
    Me.txtExamSession = New CDBNETCL.TextLookupBox()
    Me.lblExamUnit = New System.Windows.Forms.Label()
    Me.txtExamUnit = New CDBNETCL.TextLookupBox()
    Me.tbpAnalysis = New System.Windows.Forms.TabPage()
    Me.pnlAnalysis = New CDBNETCL.PanelEx()
    Me.TabAnalysis = New System.Windows.Forms.TabControl()
    Me.tbpSales = New System.Windows.Forms.TabPage()
    Me.pnlSales = New CDBNETCL.PanelEx()
    Me.chkConfirmSale = New System.Windows.Forms.CheckBox()
    Me.chkConfirmCollection = New System.Windows.Forms.CheckBox()
    Me.chkServiceBookingCredit = New System.Windows.Forms.CheckBox()
    Me.chkDonation = New System.Windows.Forms.CheckBox()
    Me.chkEventBooking = New System.Windows.Forms.CheckBox()
    Me.chkExamBooking = New System.Windows.Forms.CheckBox()
    Me.chkAccomodationBooking = New System.Windows.Forms.CheckBox()
    Me.chkServiceBooking = New System.Windows.Forms.CheckBox()
    Me.chkProduct = New System.Windows.Forms.CheckBox()
    Me.tbpPaymentPlans = New System.Windows.Forms.TabPage()
    Me.pnlPaymentPlans = New CDBNETCL.PanelEx()
    Me.chkDirectDebit = New System.Windows.Forms.CheckBox()
    Me.chkCreditCardAuthority = New System.Windows.Forms.CheckBox()
    Me.chkNoPaymentRequired = New System.Windows.Forms.CheckBox()
    Me.chkDisplayScheduledPayment = New System.Windows.Forms.CheckBox()
    Me.chkStandingOrder = New System.Windows.Forms.CheckBox()
    Me.chkCovenantedSubscription = New System.Windows.Forms.CheckBox()
    Me.chkCovenantedDonation = New System.Windows.Forms.CheckBox()
    Me.chkCOvenantedMembership = New System.Windows.Forms.CheckBox()
    Me.chkSubscription = New System.Windows.Forms.CheckBox()
    Me.chkRegularDonation = New System.Windows.Forms.CheckBox()
    Me.chkPayment = New System.Windows.Forms.CheckBox()
    Me.chkMembershipType = New System.Windows.Forms.CheckBox()
    Me.chkMembership = New System.Windows.Forms.CheckBox()
    Me.chkLoan = New System.Windows.Forms.CheckBox()
    Me.tbpSalesLedger = New System.Windows.Forms.TabPage()
    Me.pnlSalesLedger = New CDBNETCL.PanelEx()
    Me.chkInvoicePrintUnpostedBatches = New System.Windows.Forms.CheckBox()
    Me.chkUnpostedBatchMsgInPrint = New System.Windows.Forms.CheckBox()
    Me.chkDateRangeMsgInPrint = New System.Windows.Forms.CheckBox()
    Me.chkAutoCreateCreditCustomer = New System.Windows.Forms.CheckBox()
    Me.chkSundryCreditNote = New System.Windows.Forms.CheckBox()
    Me.chkInvoicePayment = New System.Windows.Forms.CheckBox()
    Me.tbpMaintenance = New System.Windows.Forms.TabPage()
    Me.pnlMaintenance = New CDBNETCL.PanelEx()
    Me.chkGoneAway = New System.Windows.Forms.CheckBox()
    Me.chkAddressMaintenance = New System.Windows.Forms.CheckBox()
    Me.chkStatus = New System.Windows.Forms.CheckBox()
    Me.txtStatus = New CDBNETCL.TextLookupBox()
    Me.chkActivity = New System.Windows.Forms.CheckBox()
    Me.txtActivity = New CDBNETCL.TextLookupBox()
    Me.chkSuppression = New System.Windows.Forms.CheckBox()
    Me.txtSuppression = New CDBNETCL.TextLookupBox()
    Me.chkGiftAidDeclaration = New System.Windows.Forms.CheckBox()
    Me.chkCancelGiftAidDeclaration = New System.Windows.Forms.CheckBox()
    Me.txtCancelGiftAidDeclaration = New CDBNETCL.TextLookupBox()
    Me.chkCancelPaymentPlan = New System.Windows.Forms.CheckBox()
    Me.chkPayrollGiving = New System.Windows.Forms.CheckBox()
    Me.chkAutoPaymentMaintenance = New System.Windows.Forms.CheckBox()
    Me.tbpLegacies = New System.Windows.Forms.TabPage()
    Me.pnlLegacies = New CDBNETCL.PanelEx()
    Me.chkLegacyReceipt = New System.Windows.Forms.CheckBox()
    Me.tabMain = New CDBNETCL.TabControl()
    Me.tbpGeneral = New System.Windows.Forms.TabPage()
    Me.pnlGeneral = New CDBNETCL.PanelEx()
    Me.lblType = New System.Windows.Forms.Label()
    Me.txtType = New CDBNETCL.TextLookupBox()
    Me.txtAppDesc = New System.Windows.Forms.TextBox()
    Me.lblMenu = New System.Windows.Forms.Label()
    Me.txtApplication = New System.Windows.Forms.TextBox()
    Me.lblApplication = New System.Windows.Forms.Label()
    Me.TabGeneral1 = New System.Windows.Forms.TabControl()
    Me.tbpOptions = New System.Windows.Forms.TabPage()
    Me.plnOption = New CDBNETCL.PanelEx()
    Me.chkAnalysisComments = New System.Windows.Forms.CheckBox()
    Me.chkConfirmDetails = New System.Windows.Forms.CheckBox()
    Me.chkAutoSetAmount = New System.Windows.Forms.CheckBox()
    Me.chkPrefulfilledIncentives = New System.Windows.Forms.CheckBox()
    Me.chkTransactionComment = New System.Windows.Forms.CheckBox()
    Me.chkMaintainPaymentPlan = New System.Windows.Forms.CheckBox()
    Me.chkConfirmProduct = New System.Windows.Forms.CheckBox()
    Me.chkConfirmAnalysis = New System.Windows.Forms.CheckBox()
    Me.ChkPaymentMethod = New System.Windows.Forms.CheckBox()
    Me.chkNonFinancialBatch = New System.Windows.Forms.CheckBox()
    Me.chkBankDetails = New System.Windows.Forms.CheckBox()
    Me.chkShowReference = New System.Windows.Forms.CheckBox()
    Me.chkCarriage = New System.Windows.Forms.CheckBox()
    Me.chkConfirmCarriage = New System.Windows.Forms.CheckBox()
    Me.chkBypass = New System.Windows.Forms.CheckBox()
    Me.chkSelectBatch = New System.Windows.Forms.CheckBox()
    Me.chkForeignCurrency = New System.Windows.Forms.CheckBox()
    Me.tbpMethods = New System.Windows.Forms.TabPage()
    Me.pnlPaymentMethods = New CDBNETCL.PanelEx()
    Me.chkCCWithInvoice = New System.Windows.Forms.CheckBox()
    Me.chkChequeWithInvoice = New System.Windows.Forms.CheckBox()
    Me.chkDebitCard = New System.Windows.Forms.CheckBox()
    Me.chkGiftInKind = New System.Windows.Forms.CheckBox()
    Me.chkPostalOrder = New System.Windows.Forms.CheckBox()
    Me.chkSaleOrReturn = New System.Windows.Forms.CheckBox()
    Me.chkCreditCard = New System.Windows.Forms.CheckBox()
    Me.chkCAFCard = New System.Windows.Forms.CheckBox()
    Me.chkCheque = New System.Windows.Forms.CheckBox()
    Me.chkPaymentPlan = New System.Windows.Forms.CheckBox()
    Me.chkCreditSale = New System.Windows.Forms.CheckBox()
    Me.chkVoucher = New System.Windows.Forms.CheckBox()
    Me.chkCash = New System.Windows.Forms.CheckBox()
    Me.TabPage2 = New System.Windows.Forms.TabPage()
    Me.lblAutoGADHelp = New CDBNETCL.InfoLabel()
    Me.newDeclarationGroup = New System.Windows.Forms.Panel()
    Me.InfoLabel1 = New CDBNETCL.InfoLabel()
    Me.gadMethodGroup = New System.Windows.Forms.Panel()
    Me.Label9 = New System.Windows.Forms.Label()
    Me.gadMethodElectronic = New System.Windows.Forms.RadioButton()
    Me.gadMethodWritten = New System.Windows.Forms.RadioButton()
    Me.gadMethodOral = New System.Windows.Forms.RadioButton()
    Me.gadSource = New CDBNETCL.TextLookupBox()
    Me.gadSourceLabel = New System.Windows.Forms.Label()
    Me.autoGiftAidDeclaration = New System.Windows.Forms.CheckBox()
    Me.TabPage3 = New System.Windows.Forms.TabPage()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.Label10 = New System.Windows.Forms.Label()
    Me.InfoLabel2 = New CDBNETCL.InfoLabel()
    Me.txtMerchantRetailNumber = New CDBNETCL.TextLookupBox()
    Me.chkRequireAuthorisation = New System.Windows.Forms.CheckBox()
    Me.chkOnlineAuth = New System.Windows.Forms.CheckBox()
    Me.TabPage4 = New System.Windows.Forms.TabPage()
    Me.pnlAlerts = New System.Windows.Forms.Panel()
    Me.pnlAlertsGrid = New System.Windows.Forms.Panel()
    Me.dgrAlerts = New CDBNETCL.DisplayGrid()
    Me.ilAlerts = New CDBNETCL.InfoLabel()
    Me.chkContactAlerts = New System.Windows.Forms.CheckBox()
    Me.bplAlerts = New CDBNETCL.ButtonPanel()
    Me.cmdAddAlert = New System.Windows.Forms.Button()
    Me.cmdAddAlertLink = New System.Windows.Forms.Button()
    Me.cmdDeleteAlertLink = New System.Windows.Forms.Button()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdCopy = New System.Windows.Forms.Button()
    Me.cmdRevert = New System.Windows.Forms.Button()
    Me.cmdDesign = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.tbpCurrency.SuspendLayout()
    Me.pnlCurrencyBAs.SuspendLayout()
    Me.tbpRestrictions.SuspendLayout()
    Me.pnlRestrictions.SuspendLayout()
    Me.tbpDocuments.SuspendLayout()
    Me.pnlDocuments.SuspendLayout()
    Me.tbpBank.SuspendLayout()
    Me.pnlBankAccount.SuspendLayout()
    Me.tbpDefaults.SuspendLayout()
    Me.tbpCarriage.SuspendLayout()
    Me.tbpBatches.SuspendLayout()
    Me.pnlBatches.SuspendLayout()
    Me.tbpTransactions.SuspendLayout()
    Me.pnlTransactions.SuspendLayout()
    Me.tbpAnalysisDefaults.SuspendLayout()
    Me.pnlAnalysisSub.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.pnlCarriage.SuspendLayout()
    Me.tbpMembers.SuspendLayout()
    Me.pnlMembers.SuspendLayout()
    Me.tbpExams.SuspendLayout()
    Me.pnlExams.SuspendLayout()
    Me.tbpAnalysis.SuspendLayout()
    Me.pnlAnalysis.SuspendLayout()
    Me.TabAnalysis.SuspendLayout()
    Me.tbpSales.SuspendLayout()
    Me.pnlSales.SuspendLayout()
    Me.tbpPaymentPlans.SuspendLayout()
    Me.pnlPaymentPlans.SuspendLayout()
    Me.tbpSalesLedger.SuspendLayout()
    Me.pnlSalesLedger.SuspendLayout()
    Me.tbpMaintenance.SuspendLayout()
    Me.pnlMaintenance.SuspendLayout()
    Me.tbpLegacies.SuspendLayout()
    Me.pnlLegacies.SuspendLayout()
    Me.tabMain.SuspendLayout()
    Me.tbpGeneral.SuspendLayout()
    Me.pnlGeneral.SuspendLayout()
    Me.TabGeneral1.SuspendLayout()
    Me.tbpOptions.SuspendLayout()
    Me.plnOption.SuspendLayout()
    Me.tbpMethods.SuspendLayout()
    Me.pnlPaymentMethods.SuspendLayout()
    Me.TabPage2.SuspendLayout()
    Me.newDeclarationGroup.SuspendLayout()
    Me.gadMethodGroup.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.Panel1.SuspendLayout()
    Me.TabPage4.SuspendLayout()
    Me.pnlAlerts.SuspendLayout()
    Me.pnlAlertsGrid.SuspendLayout()
    Me.bplAlerts.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'tbpCurrency
    '
    Me.tbpCurrency.Controls.Add(Me.pnlCurrencyBAs)
    Me.tbpCurrency.Location = New System.Drawing.Point(4, 26)
    Me.tbpCurrency.Name = "tbpCurrency"
    Me.tbpCurrency.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpCurrency.Size = New System.Drawing.Size(692, 454)
    Me.tbpCurrency.TabIndex = 7
    Me.tbpCurrency.Text = "Currency BAs"
    Me.tbpCurrency.UseVisualStyleBackColor = True
    '
    'pnlCurrencyBAs
    '
    Me.pnlCurrencyBAs.BackColor = System.Drawing.Color.Transparent
    Me.pnlCurrencyBAs.Controls.Add(Me.Label6)
    Me.pnlCurrencyBAs.Controls.Add(Me.txtCurrency)
    Me.pnlCurrencyBAs.Controls.Add(Me.Label7)
    Me.pnlCurrencyBAs.Controls.Add(Me.txtBatchType)
    Me.pnlCurrencyBAs.Controls.Add(Me.Label8)
    Me.pnlCurrencyBAs.Controls.Add(Me.txtBankAccount)
    Me.pnlCurrencyBAs.Controls.Add(Me.cmdAdd)
    Me.pnlCurrencyBAs.Controls.Add(Me.cmdRemove)
    Me.pnlCurrencyBAs.Controls.Add(Me.dgr)
    Me.pnlCurrencyBAs.Location = New System.Drawing.Point(0, 0)
    Me.pnlCurrencyBAs.Name = "pnlCurrencyBAs"
    Me.pnlCurrencyBAs.Size = New System.Drawing.Size(700, 422)
    Me.pnlCurrencyBAs.TabIndex = 68
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(6, 37)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(52, 13)
    Me.Label6.TabIndex = 63
    Me.Label6.Text = "Currency:"
    '
    'txtCurrency
    '
    Me.txtCurrency.ActiveOnly = False
    Me.txtCurrency.BackColor = System.Drawing.SystemColors.Control
    Me.txtCurrency.CustomFormNumber = 0
    Me.txtCurrency.Description = ""
    Me.txtCurrency.EnabledProperty = True
    Me.txtCurrency.ExamCentreId = 0
    Me.txtCurrency.ExamCentreUnitId = 0
    Me.txtCurrency.ExamUnitLinkId = 0
    Me.txtCurrency.HasDependancies = False
    Me.txtCurrency.IsDesign = False
    Me.txtCurrency.Location = New System.Drawing.Point(12, 57)
    Me.txtCurrency.MaxLength = 32767
    Me.txtCurrency.MultipleValuesSupported = False
    Me.txtCurrency.Name = "txtCurrency"
    Me.txtCurrency.OriginalText = Nothing
    Me.txtCurrency.PreventHistoricalSelection = False
    Me.txtCurrency.ReadOnlyProperty = False
    Me.txtCurrency.Size = New System.Drawing.Size(408, 24)
    Me.txtCurrency.TabIndex = 60
    Me.txtCurrency.TextReadOnly = False
    Me.txtCurrency.TotalWidth = 408
    Me.txtCurrency.ValidationRequired = True
    Me.txtCurrency.WarningMessage = Nothing
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(9, 96)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(65, 13)
    Me.Label7.TabIndex = 64
    Me.Label7.Text = "Batch Type:"
    '
    'txtBatchType
    '
    Me.txtBatchType.ActiveOnly = False
    Me.txtBatchType.BackColor = System.Drawing.SystemColors.Control
    Me.txtBatchType.CustomFormNumber = 0
    Me.txtBatchType.Description = ""
    Me.txtBatchType.EnabledProperty = True
    Me.txtBatchType.ExamCentreId = 0
    Me.txtBatchType.ExamCentreUnitId = 0
    Me.txtBatchType.ExamUnitLinkId = 0
    Me.txtBatchType.HasDependancies = False
    Me.txtBatchType.IsDesign = False
    Me.txtBatchType.Location = New System.Drawing.Point(12, 116)
    Me.txtBatchType.MaxLength = 32767
    Me.txtBatchType.MultipleValuesSupported = False
    Me.txtBatchType.Name = "txtBatchType"
    Me.txtBatchType.OriginalText = Nothing
    Me.txtBatchType.PreventHistoricalSelection = False
    Me.txtBatchType.ReadOnlyProperty = False
    Me.txtBatchType.Size = New System.Drawing.Size(408, 24)
    Me.txtBatchType.TabIndex = 61
    Me.txtBatchType.TextReadOnly = False
    Me.txtBatchType.TotalWidth = 408
    Me.txtBatchType.ValidationRequired = True
    Me.txtBatchType.WarningMessage = Nothing
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(9, 156)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(78, 13)
    Me.Label8.TabIndex = 65
    Me.Label8.Text = "Bank Account:"
    '
    'txtBankAccount
    '
    Me.txtBankAccount.ActiveOnly = False
    Me.txtBankAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtBankAccount.CustomFormNumber = 0
    Me.txtBankAccount.Description = ""
    Me.txtBankAccount.EnabledProperty = True
    Me.txtBankAccount.ExamCentreId = 0
    Me.txtBankAccount.ExamCentreUnitId = 0
    Me.txtBankAccount.ExamUnitLinkId = 0
    Me.txtBankAccount.HasDependancies = False
    Me.txtBankAccount.IsDesign = False
    Me.txtBankAccount.Location = New System.Drawing.Point(12, 176)
    Me.txtBankAccount.MaxLength = 32767
    Me.txtBankAccount.MultipleValuesSupported = False
    Me.txtBankAccount.Name = "txtBankAccount"
    Me.txtBankAccount.OriginalText = Nothing
    Me.txtBankAccount.PreventHistoricalSelection = False
    Me.txtBankAccount.ReadOnlyProperty = False
    Me.txtBankAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtBankAccount.TabIndex = 62
    Me.txtBankAccount.TextReadOnly = False
    Me.txtBankAccount.TotalWidth = 408
    Me.txtBankAccount.ValidationRequired = True
    Me.txtBankAccount.WarningMessage = Nothing
    '
    'cmdAdd
    '
    Me.cmdAdd.Location = New System.Drawing.Point(84, 225)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
    Me.cmdAdd.TabIndex = 63
    Me.cmdAdd.Text = "Add"
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'cmdRemove
    '
    Me.cmdRemove.Location = New System.Drawing.Point(204, 225)
    Me.cmdRemove.Name = "cmdRemove"
    Me.cmdRemove.Size = New System.Drawing.Size(75, 23)
    Me.cmdRemove.TabIndex = 64
    Me.cmdRemove.Text = "Remove"
    Me.cmdRemove.UseVisualStyleBackColor = True
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowColumnResize = True
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgr.Location = New System.Drawing.Point(426, 18)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(246, 203)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'tbpRestrictions
    '
    Me.tbpRestrictions.Controls.Add(Me.pnlRestrictions)
    Me.tbpRestrictions.Location = New System.Drawing.Point(4, 26)
    Me.tbpRestrictions.Name = "tbpRestrictions"
    Me.tbpRestrictions.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpRestrictions.Size = New System.Drawing.Size(692, 454)
    Me.tbpRestrictions.TabIndex = 6
    Me.tbpRestrictions.Text = "Restrictions"
    Me.tbpRestrictions.UseVisualStyleBackColor = True
    '
    'pnlRestrictions
    '
    Me.pnlRestrictions.BackColor = System.Drawing.Color.Transparent
    Me.pnlRestrictions.Controls.Add(Me.chkMembersOnly)
    Me.pnlRestrictions.Controls.Add(Me.lblSalesGroup)
    Me.pnlRestrictions.Controls.Add(Me.txtSalesGroup)
    Me.pnlRestrictions.Controls.Add(Me.lblAlbacsBankDetails)
    Me.pnlRestrictions.Controls.Add(Me.cboAlbacsBankDetails)
    Me.pnlRestrictions.Controls.Add(Me.chkIncludeConfirmedTransaction)
    Me.pnlRestrictions.Controls.Add(Me.chkIncludeProvisionalTransaction)
    Me.pnlRestrictions.Controls.Add(Me.chkIncludeProvPaymentPlan)
    Me.pnlRestrictions.Controls.Add(Me.chkForceMailingCode)
    Me.pnlRestrictions.Controls.Add(Me.chkForceDistributionCode)
    Me.pnlRestrictions.Controls.Add(Me.chkSalesContactMandatory)
    Me.pnlRestrictions.Location = New System.Drawing.Point(0, 0)
    Me.pnlRestrictions.Name = "pnlRestrictions"
    Me.pnlRestrictions.Size = New System.Drawing.Size(700, 420)
    Me.pnlRestrictions.TabIndex = 63
    '
    'chkMembersOnly
    '
    Me.chkMembersOnly.AutoSize = True
    Me.chkMembersOnly.Location = New System.Drawing.Point(18, 16)
    Me.chkMembersOnly.Name = "chkMembersOnly"
    Me.chkMembersOnly.Size = New System.Drawing.Size(93, 17)
    Me.chkMembersOnly.TabIndex = 0
    Me.chkMembersOnly.Text = "Members Only"
    Me.chkMembersOnly.UseVisualStyleBackColor = True
    '
    'lblSalesGroup
    '
    Me.lblSalesGroup.AutoSize = True
    Me.lblSalesGroup.Location = New System.Drawing.Point(15, 58)
    Me.lblSalesGroup.Name = "lblSalesGroup"
    Me.lblSalesGroup.Size = New System.Drawing.Size(65, 13)
    Me.lblSalesGroup.TabIndex = 1
    Me.lblSalesGroup.Text = "Sales Group"
    '
    'txtSalesGroup
    '
    Me.txtSalesGroup.ActiveOnly = False
    Me.txtSalesGroup.BackColor = System.Drawing.SystemColors.Control
    Me.txtSalesGroup.CustomFormNumber = 0
    Me.txtSalesGroup.Description = ""
    Me.txtSalesGroup.EnabledProperty = True
    Me.txtSalesGroup.ExamCentreId = 0
    Me.txtSalesGroup.ExamCentreUnitId = 0
    Me.txtSalesGroup.ExamUnitLinkId = 0
    Me.txtSalesGroup.HasDependancies = False
    Me.txtSalesGroup.IsDesign = False
    Me.txtSalesGroup.Location = New System.Drawing.Point(162, 58)
    Me.txtSalesGroup.MaxLength = 32767
    Me.txtSalesGroup.MultipleValuesSupported = False
    Me.txtSalesGroup.Name = "txtSalesGroup"
    Me.txtSalesGroup.OriginalText = Nothing
    Me.txtSalesGroup.PreventHistoricalSelection = False
    Me.txtSalesGroup.ReadOnlyProperty = False
    Me.txtSalesGroup.Size = New System.Drawing.Size(408, 24)
    Me.txtSalesGroup.TabIndex = 1
    Me.txtSalesGroup.TextReadOnly = False
    Me.txtSalesGroup.TotalWidth = 408
    Me.txtSalesGroup.ValidationRequired = True
    Me.txtSalesGroup.WarningMessage = Nothing
    '
    'lblAlbacsBankDetails
    '
    Me.lblAlbacsBankDetails.AutoSize = True
    Me.lblAlbacsBankDetails.Location = New System.Drawing.Point(15, 97)
    Me.lblAlbacsBankDetails.Name = "lblAlbacsBankDetails"
    Me.lblAlbacsBankDetails.Size = New System.Drawing.Size(102, 13)
    Me.lblAlbacsBankDetails.TabIndex = 3
    Me.lblAlbacsBankDetails.Text = "Albacs Bank Details"
    '
    'cboAlbacsBankDetails
    '
    Me.cboAlbacsBankDetails.FormattingEnabled = True
    Me.cboAlbacsBankDetails.Location = New System.Drawing.Point(162, 97)
    Me.cboAlbacsBankDetails.Name = "cboAlbacsBankDetails"
    Me.cboAlbacsBankDetails.Size = New System.Drawing.Size(121, 21)
    Me.cboAlbacsBankDetails.TabIndex = 2
    '
    'chkIncludeConfirmedTransaction
    '
    Me.chkIncludeConfirmedTransaction.AutoSize = True
    Me.chkIncludeConfirmedTransaction.Enabled = False
    Me.chkIncludeConfirmedTransaction.Location = New System.Drawing.Point(22, 160)
    Me.chkIncludeConfirmedTransaction.Name = "chkIncludeConfirmedTransaction"
    Me.chkIncludeConfirmedTransaction.Size = New System.Drawing.Size(175, 17)
    Me.chkIncludeConfirmedTransaction.TabIndex = 3
    Me.chkIncludeConfirmedTransaction.Text = "Include Confirmed Transactions"
    Me.chkIncludeConfirmedTransaction.UseVisualStyleBackColor = True
    '
    'chkIncludeProvisionalTransaction
    '
    Me.chkIncludeProvisionalTransaction.AutoSize = True
    Me.chkIncludeProvisionalTransaction.Enabled = False
    Me.chkIncludeProvisionalTransaction.Location = New System.Drawing.Point(22, 187)
    Me.chkIncludeProvisionalTransaction.Name = "chkIncludeProvisionalTransaction"
    Me.chkIncludeProvisionalTransaction.Size = New System.Drawing.Size(179, 17)
    Me.chkIncludeProvisionalTransaction.TabIndex = 4
    Me.chkIncludeProvisionalTransaction.Text = "Include Provisional Transactions"
    Me.chkIncludeProvisionalTransaction.UseVisualStyleBackColor = True
    '
    'chkIncludeProvPaymentPlan
    '
    Me.chkIncludeProvPaymentPlan.AutoSize = True
    Me.chkIncludeProvPaymentPlan.Location = New System.Drawing.Point(22, 214)
    Me.chkIncludeProvPaymentPlan.Name = "chkIncludeProvPaymentPlan"
    Me.chkIncludeProvPaymentPlan.Size = New System.Drawing.Size(183, 17)
    Me.chkIncludeProvPaymentPlan.TabIndex = 5
    Me.chkIncludeProvPaymentPlan.Text = "Include Provisional Payment Plan"
    Me.chkIncludeProvPaymentPlan.UseVisualStyleBackColor = True
    '
    'chkForceMailingCode
    '
    Me.chkForceMailingCode.AutoSize = True
    Me.chkForceMailingCode.Location = New System.Drawing.Point(22, 243)
    Me.chkForceMailingCode.Name = "chkForceMailingCode"
    Me.chkForceMailingCode.Size = New System.Drawing.Size(197, 17)
    Me.chkForceMailingCode.TabIndex = 6
    Me.chkForceMailingCode.Text = "Force Mailing Code to be Mandatory"
    Me.chkForceMailingCode.UseVisualStyleBackColor = True
    '
    'chkForceDistributionCode
    '
    Me.chkForceDistributionCode.AutoSize = True
    Me.chkForceDistributionCode.Location = New System.Drawing.Point(22, 270)
    Me.chkForceDistributionCode.Name = "chkForceDistributionCode"
    Me.chkForceDistributionCode.Size = New System.Drawing.Size(216, 17)
    Me.chkForceDistributionCode.TabIndex = 7
    Me.chkForceDistributionCode.Text = "Force Distribution Code to be Mandatory"
    Me.chkForceDistributionCode.UseVisualStyleBackColor = True
    '
    'chkSalesContactMandatory
    '
    Me.chkSalesContactMandatory.AutoSize = True
    Me.chkSalesContactMandatory.Location = New System.Drawing.Point(22, 297)
    Me.chkSalesContactMandatory.Name = "chkSalesContactMandatory"
    Me.chkSalesContactMandatory.Size = New System.Drawing.Size(202, 17)
    Me.chkSalesContactMandatory.TabIndex = 8
    Me.chkSalesContactMandatory.Text = "Force Sales Contact to be Mandatory"
    Me.chkSalesContactMandatory.UseVisualStyleBackColor = True
    '
    'tbpDocuments
    '
    Me.tbpDocuments.Controls.Add(Me.pnlDocuments)
    Me.tbpDocuments.Location = New System.Drawing.Point(4, 26)
    Me.tbpDocuments.Name = "tbpDocuments"
    Me.tbpDocuments.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpDocuments.Size = New System.Drawing.Size(692, 454)
    Me.tbpDocuments.TabIndex = 5
    Me.tbpDocuments.Text = "Documents"
    Me.tbpDocuments.UseVisualStyleBackColor = True
    '
    'pnlDocuments
    '
    Me.pnlDocuments.BackColor = System.Drawing.Color.Transparent
    Me.pnlDocuments.Controls.Add(Me.Label4)
    Me.pnlDocuments.Controls.Add(Me.txtInvoiceDoc)
    Me.pnlDocuments.Controls.Add(Me.Label5)
    Me.pnlDocuments.Controls.Add(Me.txtReceiptDoc)
    Me.pnlDocuments.Controls.Add(Me.Label2)
    Me.pnlDocuments.Controls.Add(Me.txtPayPlanDoc)
    Me.pnlDocuments.Controls.Add(Me.Label3)
    Me.pnlDocuments.Controls.Add(Me.txtCreditStmtDoc)
    Me.pnlDocuments.Controls.Add(Me.Label1)
    Me.pnlDocuments.Controls.Add(Me.txtProvCashDoc)
    Me.pnlDocuments.Location = New System.Drawing.Point(0, 0)
    Me.pnlDocuments.Name = "pnlDocuments"
    Me.pnlDocuments.Size = New System.Drawing.Size(689, 412)
    Me.pnlDocuments.TabIndex = 68
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(12, 19)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(68, 13)
    Me.Label4.TabIndex = 61
    Me.Label4.Text = "Invoice Doc."
    '
    'txtInvoiceDoc
    '
    Me.txtInvoiceDoc.ActiveOnly = False
    Me.txtInvoiceDoc.BackColor = System.Drawing.SystemColors.Control
    Me.txtInvoiceDoc.CustomFormNumber = 0
    Me.txtInvoiceDoc.Description = ""
    Me.txtInvoiceDoc.EnabledProperty = True
    Me.txtInvoiceDoc.ExamCentreId = 0
    Me.txtInvoiceDoc.ExamCentreUnitId = 0
    Me.txtInvoiceDoc.ExamUnitLinkId = 0
    Me.txtInvoiceDoc.HasDependancies = False
    Me.txtInvoiceDoc.IsDesign = False
    Me.txtInvoiceDoc.Location = New System.Drawing.Point(193, 19)
    Me.txtInvoiceDoc.MaxLength = 32767
    Me.txtInvoiceDoc.MultipleValuesSupported = False
    Me.txtInvoiceDoc.Name = "txtInvoiceDoc"
    Me.txtInvoiceDoc.OriginalText = Nothing
    Me.txtInvoiceDoc.PreventHistoricalSelection = False
    Me.txtInvoiceDoc.ReadOnlyProperty = False
    Me.txtInvoiceDoc.Size = New System.Drawing.Size(408, 24)
    Me.txtInvoiceDoc.TabIndex = 58
    Me.txtInvoiceDoc.TextReadOnly = False
    Me.txtInvoiceDoc.TotalWidth = 408
    Me.txtInvoiceDoc.ValidationRequired = True
    Me.txtInvoiceDoc.WarningMessage = Nothing
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(12, 60)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(70, 13)
    Me.Label5.TabIndex = 60
    Me.Label5.Text = "Receipt Doc."
    '
    'txtReceiptDoc
    '
    Me.txtReceiptDoc.ActiveOnly = False
    Me.txtReceiptDoc.BackColor = System.Drawing.SystemColors.Control
    Me.txtReceiptDoc.CustomFormNumber = 0
    Me.txtReceiptDoc.Description = ""
    Me.txtReceiptDoc.EnabledProperty = True
    Me.txtReceiptDoc.ExamCentreId = 0
    Me.txtReceiptDoc.ExamCentreUnitId = 0
    Me.txtReceiptDoc.ExamUnitLinkId = 0
    Me.txtReceiptDoc.HasDependancies = False
    Me.txtReceiptDoc.IsDesign = False
    Me.txtReceiptDoc.Location = New System.Drawing.Point(193, 60)
    Me.txtReceiptDoc.MaxLength = 32767
    Me.txtReceiptDoc.MultipleValuesSupported = False
    Me.txtReceiptDoc.Name = "txtReceiptDoc"
    Me.txtReceiptDoc.OriginalText = Nothing
    Me.txtReceiptDoc.PreventHistoricalSelection = False
    Me.txtReceiptDoc.ReadOnlyProperty = False
    Me.txtReceiptDoc.Size = New System.Drawing.Size(408, 24)
    Me.txtReceiptDoc.TabIndex = 59
    Me.txtReceiptDoc.TextReadOnly = False
    Me.txtReceiptDoc.TotalWidth = 408
    Me.txtReceiptDoc.ValidationRequired = True
    Me.txtReceiptDoc.WarningMessage = Nothing
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 103)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(75, 13)
    Me.Label2.TabIndex = 65
    Me.Label2.Text = "Pay Plan Doc."
    '
    'txtPayPlanDoc
    '
    Me.txtPayPlanDoc.ActiveOnly = False
    Me.txtPayPlanDoc.BackColor = System.Drawing.SystemColors.Control
    Me.txtPayPlanDoc.CustomFormNumber = 0
    Me.txtPayPlanDoc.Description = ""
    Me.txtPayPlanDoc.EnabledProperty = True
    Me.txtPayPlanDoc.ExamCentreId = 0
    Me.txtPayPlanDoc.ExamCentreUnitId = 0
    Me.txtPayPlanDoc.ExamUnitLinkId = 0
    Me.txtPayPlanDoc.HasDependancies = False
    Me.txtPayPlanDoc.IsDesign = False
    Me.txtPayPlanDoc.Location = New System.Drawing.Point(193, 103)
    Me.txtPayPlanDoc.MaxLength = 32767
    Me.txtPayPlanDoc.MultipleValuesSupported = False
    Me.txtPayPlanDoc.Name = "txtPayPlanDoc"
    Me.txtPayPlanDoc.OriginalText = Nothing
    Me.txtPayPlanDoc.PreventHistoricalSelection = False
    Me.txtPayPlanDoc.ReadOnlyProperty = False
    Me.txtPayPlanDoc.Size = New System.Drawing.Size(408, 24)
    Me.txtPayPlanDoc.TabIndex = 60
    Me.txtPayPlanDoc.TextReadOnly = False
    Me.txtPayPlanDoc.TotalWidth = 408
    Me.txtPayPlanDoc.ValidationRequired = True
    Me.txtPayPlanDoc.WarningMessage = Nothing
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 142)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(84, 13)
    Me.Label3.TabIndex = 63
    Me.Label3.Text = "Credit Stmt Doc."
    '
    'txtCreditStmtDoc
    '
    Me.txtCreditStmtDoc.ActiveOnly = False
    Me.txtCreditStmtDoc.BackColor = System.Drawing.SystemColors.Control
    Me.txtCreditStmtDoc.CustomFormNumber = 0
    Me.txtCreditStmtDoc.Description = ""
    Me.txtCreditStmtDoc.EnabledProperty = True
    Me.txtCreditStmtDoc.ExamCentreId = 0
    Me.txtCreditStmtDoc.ExamCentreUnitId = 0
    Me.txtCreditStmtDoc.ExamUnitLinkId = 0
    Me.txtCreditStmtDoc.HasDependancies = False
    Me.txtCreditStmtDoc.IsDesign = False
    Me.txtCreditStmtDoc.Location = New System.Drawing.Point(193, 142)
    Me.txtCreditStmtDoc.MaxLength = 32767
    Me.txtCreditStmtDoc.MultipleValuesSupported = False
    Me.txtCreditStmtDoc.Name = "txtCreditStmtDoc"
    Me.txtCreditStmtDoc.OriginalText = Nothing
    Me.txtCreditStmtDoc.PreventHistoricalSelection = False
    Me.txtCreditStmtDoc.ReadOnlyProperty = False
    Me.txtCreditStmtDoc.Size = New System.Drawing.Size(408, 24)
    Me.txtCreditStmtDoc.TabIndex = 61
    Me.txtCreditStmtDoc.TextReadOnly = False
    Me.txtCreditStmtDoc.TotalWidth = 408
    Me.txtCreditStmtDoc.ValidationRequired = True
    Me.txtCreditStmtDoc.WarningMessage = Nothing
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 181)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(82, 13)
    Me.Label1.TabIndex = 67
    Me.Label1.Text = "Prov Cash Doc."
    '
    'txtProvCashDoc
    '
    Me.txtProvCashDoc.ActiveOnly = False
    Me.txtProvCashDoc.BackColor = System.Drawing.SystemColors.Control
    Me.txtProvCashDoc.CustomFormNumber = 0
    Me.txtProvCashDoc.Description = ""
    Me.txtProvCashDoc.EnabledProperty = True
    Me.txtProvCashDoc.ExamCentreId = 0
    Me.txtProvCashDoc.ExamCentreUnitId = 0
    Me.txtProvCashDoc.ExamUnitLinkId = 0
    Me.txtProvCashDoc.HasDependancies = False
    Me.txtProvCashDoc.IsDesign = False
    Me.txtProvCashDoc.Location = New System.Drawing.Point(193, 181)
    Me.txtProvCashDoc.MaxLength = 32767
    Me.txtProvCashDoc.MultipleValuesSupported = False
    Me.txtProvCashDoc.Name = "txtProvCashDoc"
    Me.txtProvCashDoc.OriginalText = Nothing
    Me.txtProvCashDoc.PreventHistoricalSelection = False
    Me.txtProvCashDoc.ReadOnlyProperty = False
    Me.txtProvCashDoc.Size = New System.Drawing.Size(408, 24)
    Me.txtProvCashDoc.TabIndex = 62
    Me.txtProvCashDoc.TextReadOnly = False
    Me.txtProvCashDoc.TotalWidth = 408
    Me.txtProvCashDoc.ValidationRequired = True
    Me.txtProvCashDoc.WarningMessage = Nothing
    '
    'tbpBank
    '
    Me.tbpBank.Controls.Add(Me.pnlBankAccount)
    Me.tbpBank.Location = New System.Drawing.Point(4, 26)
    Me.tbpBank.Name = "tbpBank"
    Me.tbpBank.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpBank.Size = New System.Drawing.Size(692, 454)
    Me.tbpBank.TabIndex = 4
    Me.tbpBank.Text = "Bank Account"
    Me.tbpBank.UseVisualStyleBackColor = True
    '
    'pnlBankAccount
    '
    Me.pnlBankAccount.BackColor = System.Drawing.Color.Transparent
    Me.pnlBankAccount.Controls.Add(Me.lblCashAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtCashAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblCreditCardAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtCreditCardAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblDebitCardAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtDebitCardAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblCreditSaleAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtCreditSaleAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblStandingOrderAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtStandingOrderAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblDirectDebitAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtDirectDebitAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblCCAAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtCCAAccount)
    Me.pnlBankAccount.Controls.Add(Me.lblCAFAndVoucherAccount)
    Me.pnlBankAccount.Controls.Add(Me.txtCAFAndVoucherAccount)
    Me.pnlBankAccount.Location = New System.Drawing.Point(0, 0)
    Me.pnlBankAccount.Name = "pnlBankAccount"
    Me.pnlBankAccount.Size = New System.Drawing.Size(703, 426)
    Me.pnlBankAccount.TabIndex = 62
    '
    'lblCashAccount
    '
    Me.lblCashAccount.AutoSize = True
    Me.lblCashAccount.Location = New System.Drawing.Point(24, 25)
    Me.lblCashAccount.Name = "lblCashAccount"
    Me.lblCashAccount.Size = New System.Drawing.Size(74, 13)
    Me.lblCashAccount.TabIndex = 49
    Me.lblCashAccount.Text = "Cash Account"
    '
    'txtCashAccount
    '
    Me.txtCashAccount.ActiveOnly = False
    Me.txtCashAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtCashAccount.CustomFormNumber = 0
    Me.txtCashAccount.Description = ""
    Me.txtCashAccount.EnabledProperty = True
    Me.txtCashAccount.ExamCentreId = 0
    Me.txtCashAccount.ExamCentreUnitId = 0
    Me.txtCashAccount.ExamUnitLinkId = 0
    Me.txtCashAccount.HasDependancies = False
    Me.txtCashAccount.IsDesign = False
    Me.txtCashAccount.Location = New System.Drawing.Point(205, 25)
    Me.txtCashAccount.MaxLength = 32767
    Me.txtCashAccount.MultipleValuesSupported = False
    Me.txtCashAccount.Name = "txtCashAccount"
    Me.txtCashAccount.OriginalText = Nothing
    Me.txtCashAccount.PreventHistoricalSelection = False
    Me.txtCashAccount.ReadOnlyProperty = False
    Me.txtCashAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtCashAccount.TabIndex = 46
    Me.txtCashAccount.TextReadOnly = False
    Me.txtCashAccount.TotalWidth = 408
    Me.txtCashAccount.ValidationRequired = True
    Me.txtCashAccount.WarningMessage = Nothing
    '
    'lblCreditCardAccount
    '
    Me.lblCreditCardAccount.AutoSize = True
    Me.lblCreditCardAccount.Location = New System.Drawing.Point(24, 66)
    Me.lblCreditCardAccount.Name = "lblCreditCardAccount"
    Me.lblCreditCardAccount.Size = New System.Drawing.Size(102, 13)
    Me.lblCreditCardAccount.TabIndex = 48
    Me.lblCreditCardAccount.Text = "Credit Card Account"
    '
    'txtCreditCardAccount
    '
    Me.txtCreditCardAccount.ActiveOnly = False
    Me.txtCreditCardAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtCreditCardAccount.CustomFormNumber = 0
    Me.txtCreditCardAccount.Description = ""
    Me.txtCreditCardAccount.EnabledProperty = True
    Me.txtCreditCardAccount.ExamCentreId = 0
    Me.txtCreditCardAccount.ExamCentreUnitId = 0
    Me.txtCreditCardAccount.ExamUnitLinkId = 0
    Me.txtCreditCardAccount.HasDependancies = False
    Me.txtCreditCardAccount.IsDesign = False
    Me.txtCreditCardAccount.Location = New System.Drawing.Point(205, 66)
    Me.txtCreditCardAccount.MaxLength = 32767
    Me.txtCreditCardAccount.MultipleValuesSupported = False
    Me.txtCreditCardAccount.Name = "txtCreditCardAccount"
    Me.txtCreditCardAccount.OriginalText = Nothing
    Me.txtCreditCardAccount.PreventHistoricalSelection = False
    Me.txtCreditCardAccount.ReadOnlyProperty = False
    Me.txtCreditCardAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtCreditCardAccount.TabIndex = 47
    Me.txtCreditCardAccount.TextReadOnly = False
    Me.txtCreditCardAccount.TotalWidth = 408
    Me.txtCreditCardAccount.ValidationRequired = True
    Me.txtCreditCardAccount.WarningMessage = Nothing
    '
    'lblDebitCardAccount
    '
    Me.lblDebitCardAccount.AutoSize = True
    Me.lblDebitCardAccount.Location = New System.Drawing.Point(24, 109)
    Me.lblDebitCardAccount.Name = "lblDebitCardAccount"
    Me.lblDebitCardAccount.Size = New System.Drawing.Size(100, 13)
    Me.lblDebitCardAccount.TabIndex = 53
    Me.lblDebitCardAccount.Text = "Debit Card Account"
    '
    'txtDebitCardAccount
    '
    Me.txtDebitCardAccount.ActiveOnly = False
    Me.txtDebitCardAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtDebitCardAccount.CustomFormNumber = 0
    Me.txtDebitCardAccount.Description = ""
    Me.txtDebitCardAccount.EnabledProperty = True
    Me.txtDebitCardAccount.ExamCentreId = 0
    Me.txtDebitCardAccount.ExamCentreUnitId = 0
    Me.txtDebitCardAccount.ExamUnitLinkId = 0
    Me.txtDebitCardAccount.HasDependancies = False
    Me.txtDebitCardAccount.IsDesign = False
    Me.txtDebitCardAccount.Location = New System.Drawing.Point(205, 109)
    Me.txtDebitCardAccount.MaxLength = 32767
    Me.txtDebitCardAccount.MultipleValuesSupported = False
    Me.txtDebitCardAccount.Name = "txtDebitCardAccount"
    Me.txtDebitCardAccount.OriginalText = Nothing
    Me.txtDebitCardAccount.PreventHistoricalSelection = False
    Me.txtDebitCardAccount.ReadOnlyProperty = False
    Me.txtDebitCardAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtDebitCardAccount.TabIndex = 48
    Me.txtDebitCardAccount.TextReadOnly = False
    Me.txtDebitCardAccount.TotalWidth = 408
    Me.txtDebitCardAccount.ValidationRequired = True
    Me.txtDebitCardAccount.WarningMessage = Nothing
    '
    'lblCreditSaleAccount
    '
    Me.lblCreditSaleAccount.AutoSize = True
    Me.lblCreditSaleAccount.Location = New System.Drawing.Point(24, 148)
    Me.lblCreditSaleAccount.Name = "lblCreditSaleAccount"
    Me.lblCreditSaleAccount.Size = New System.Drawing.Size(106, 13)
    Me.lblCreditSaleAccount.TabIndex = 51
    Me.lblCreditSaleAccount.Text = "Credit Sales Account"
    '
    'txtCreditSaleAccount
    '
    Me.txtCreditSaleAccount.ActiveOnly = False
    Me.txtCreditSaleAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtCreditSaleAccount.CustomFormNumber = 0
    Me.txtCreditSaleAccount.Description = ""
    Me.txtCreditSaleAccount.EnabledProperty = True
    Me.txtCreditSaleAccount.ExamCentreId = 0
    Me.txtCreditSaleAccount.ExamCentreUnitId = 0
    Me.txtCreditSaleAccount.ExamUnitLinkId = 0
    Me.txtCreditSaleAccount.HasDependancies = False
    Me.txtCreditSaleAccount.IsDesign = False
    Me.txtCreditSaleAccount.Location = New System.Drawing.Point(205, 148)
    Me.txtCreditSaleAccount.MaxLength = 32767
    Me.txtCreditSaleAccount.MultipleValuesSupported = False
    Me.txtCreditSaleAccount.Name = "txtCreditSaleAccount"
    Me.txtCreditSaleAccount.OriginalText = Nothing
    Me.txtCreditSaleAccount.PreventHistoricalSelection = False
    Me.txtCreditSaleAccount.ReadOnlyProperty = False
    Me.txtCreditSaleAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtCreditSaleAccount.TabIndex = 49
    Me.txtCreditSaleAccount.TextReadOnly = False
    Me.txtCreditSaleAccount.TotalWidth = 408
    Me.txtCreditSaleAccount.ValidationRequired = True
    Me.txtCreditSaleAccount.WarningMessage = Nothing
    '
    'lblStandingOrderAccount
    '
    Me.lblStandingOrderAccount.AutoSize = True
    Me.lblStandingOrderAccount.Location = New System.Drawing.Point(24, 187)
    Me.lblStandingOrderAccount.Name = "lblStandingOrderAccount"
    Me.lblStandingOrderAccount.Size = New System.Drawing.Size(121, 13)
    Me.lblStandingOrderAccount.TabIndex = 57
    Me.lblStandingOrderAccount.Text = "Standing Order Account"
    '
    'txtStandingOrderAccount
    '
    Me.txtStandingOrderAccount.ActiveOnly = False
    Me.txtStandingOrderAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtStandingOrderAccount.CustomFormNumber = 0
    Me.txtStandingOrderAccount.Description = ""
    Me.txtStandingOrderAccount.EnabledProperty = True
    Me.txtStandingOrderAccount.ExamCentreId = 0
    Me.txtStandingOrderAccount.ExamCentreUnitId = 0
    Me.txtStandingOrderAccount.ExamUnitLinkId = 0
    Me.txtStandingOrderAccount.HasDependancies = False
    Me.txtStandingOrderAccount.IsDesign = False
    Me.txtStandingOrderAccount.Location = New System.Drawing.Point(205, 187)
    Me.txtStandingOrderAccount.MaxLength = 32767
    Me.txtStandingOrderAccount.MultipleValuesSupported = False
    Me.txtStandingOrderAccount.Name = "txtStandingOrderAccount"
    Me.txtStandingOrderAccount.OriginalText = Nothing
    Me.txtStandingOrderAccount.PreventHistoricalSelection = False
    Me.txtStandingOrderAccount.ReadOnlyProperty = False
    Me.txtStandingOrderAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtStandingOrderAccount.TabIndex = 50
    Me.txtStandingOrderAccount.TextReadOnly = False
    Me.txtStandingOrderAccount.TotalWidth = 408
    Me.txtStandingOrderAccount.ValidationRequired = True
    Me.txtStandingOrderAccount.WarningMessage = Nothing
    '
    'lblDirectDebitAccount
    '
    Me.lblDirectDebitAccount.AutoSize = True
    Me.lblDirectDebitAccount.Location = New System.Drawing.Point(24, 228)
    Me.lblDirectDebitAccount.Name = "lblDirectDebitAccount"
    Me.lblDirectDebitAccount.Size = New System.Drawing.Size(106, 13)
    Me.lblDirectDebitAccount.TabIndex = 56
    Me.lblDirectDebitAccount.Text = "Direct Debit Account"
    '
    'txtDirectDebitAccount
    '
    Me.txtDirectDebitAccount.ActiveOnly = False
    Me.txtDirectDebitAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtDirectDebitAccount.CustomFormNumber = 0
    Me.txtDirectDebitAccount.Description = ""
    Me.txtDirectDebitAccount.EnabledProperty = True
    Me.txtDirectDebitAccount.ExamCentreId = 0
    Me.txtDirectDebitAccount.ExamCentreUnitId = 0
    Me.txtDirectDebitAccount.ExamUnitLinkId = 0
    Me.txtDirectDebitAccount.HasDependancies = False
    Me.txtDirectDebitAccount.IsDesign = False
    Me.txtDirectDebitAccount.Location = New System.Drawing.Point(205, 228)
    Me.txtDirectDebitAccount.MaxLength = 32767
    Me.txtDirectDebitAccount.MultipleValuesSupported = False
    Me.txtDirectDebitAccount.Name = "txtDirectDebitAccount"
    Me.txtDirectDebitAccount.OriginalText = Nothing
    Me.txtDirectDebitAccount.PreventHistoricalSelection = False
    Me.txtDirectDebitAccount.ReadOnlyProperty = False
    Me.txtDirectDebitAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtDirectDebitAccount.TabIndex = 51
    Me.txtDirectDebitAccount.TextReadOnly = False
    Me.txtDirectDebitAccount.TotalWidth = 408
    Me.txtDirectDebitAccount.ValidationRequired = True
    Me.txtDirectDebitAccount.WarningMessage = Nothing
    '
    'lblCCAAccount
    '
    Me.lblCCAAccount.AutoSize = True
    Me.lblCCAAccount.Location = New System.Drawing.Point(24, 271)
    Me.lblCCAAccount.Name = "lblCCAAccount"
    Me.lblCCAAccount.Size = New System.Drawing.Size(71, 13)
    Me.lblCCAAccount.TabIndex = 61
    Me.lblCCAAccount.Text = "CCA Account"
    '
    'txtCCAAccount
    '
    Me.txtCCAAccount.ActiveOnly = False
    Me.txtCCAAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtCCAAccount.CustomFormNumber = 0
    Me.txtCCAAccount.Description = ""
    Me.txtCCAAccount.EnabledProperty = True
    Me.txtCCAAccount.ExamCentreId = 0
    Me.txtCCAAccount.ExamCentreUnitId = 0
    Me.txtCCAAccount.ExamUnitLinkId = 0
    Me.txtCCAAccount.HasDependancies = False
    Me.txtCCAAccount.IsDesign = False
    Me.txtCCAAccount.Location = New System.Drawing.Point(205, 271)
    Me.txtCCAAccount.MaxLength = 32767
    Me.txtCCAAccount.MultipleValuesSupported = False
    Me.txtCCAAccount.Name = "txtCCAAccount"
    Me.txtCCAAccount.OriginalText = Nothing
    Me.txtCCAAccount.PreventHistoricalSelection = False
    Me.txtCCAAccount.ReadOnlyProperty = False
    Me.txtCCAAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtCCAAccount.TabIndex = 52
    Me.txtCCAAccount.TextReadOnly = False
    Me.txtCCAAccount.TotalWidth = 408
    Me.txtCCAAccount.ValidationRequired = True
    Me.txtCCAAccount.WarningMessage = Nothing
    '
    'lblCAFAndVoucherAccount
    '
    Me.lblCAFAndVoucherAccount.AutoSize = True
    Me.lblCAFAndVoucherAccount.Location = New System.Drawing.Point(24, 310)
    Me.lblCAFAndVoucherAccount.Name = "lblCAFAndVoucherAccount"
    Me.lblCAFAndVoucherAccount.Size = New System.Drawing.Size(134, 13)
    Me.lblCAFAndVoucherAccount.TabIndex = 59
    Me.lblCAFAndVoucherAccount.Text = "CAF and Voucher Account"
    '
    'txtCAFAndVoucherAccount
    '
    Me.txtCAFAndVoucherAccount.ActiveOnly = False
    Me.txtCAFAndVoucherAccount.BackColor = System.Drawing.SystemColors.Control
    Me.txtCAFAndVoucherAccount.CustomFormNumber = 0
    Me.txtCAFAndVoucherAccount.Description = ""
    Me.txtCAFAndVoucherAccount.EnabledProperty = True
    Me.txtCAFAndVoucherAccount.ExamCentreId = 0
    Me.txtCAFAndVoucherAccount.ExamCentreUnitId = 0
    Me.txtCAFAndVoucherAccount.ExamUnitLinkId = 0
    Me.txtCAFAndVoucherAccount.HasDependancies = False
    Me.txtCAFAndVoucherAccount.IsDesign = False
    Me.txtCAFAndVoucherAccount.Location = New System.Drawing.Point(205, 310)
    Me.txtCAFAndVoucherAccount.MaxLength = 32767
    Me.txtCAFAndVoucherAccount.MultipleValuesSupported = False
    Me.txtCAFAndVoucherAccount.Name = "txtCAFAndVoucherAccount"
    Me.txtCAFAndVoucherAccount.OriginalText = Nothing
    Me.txtCAFAndVoucherAccount.PreventHistoricalSelection = False
    Me.txtCAFAndVoucherAccount.ReadOnlyProperty = False
    Me.txtCAFAndVoucherAccount.Size = New System.Drawing.Size(408, 24)
    Me.txtCAFAndVoucherAccount.TabIndex = 53
    Me.txtCAFAndVoucherAccount.TextReadOnly = False
    Me.txtCAFAndVoucherAccount.TotalWidth = 408
    Me.txtCAFAndVoucherAccount.ValidationRequired = True
    Me.txtCAFAndVoucherAccount.WarningMessage = Nothing
    '
    'tbpDefaults
    '
    Me.tbpDefaults.Controls.Add(Me.tbpCarriage)
    Me.tbpDefaults.Location = New System.Drawing.Point(4, 26)
    Me.tbpDefaults.Name = "tbpDefaults"
    Me.tbpDefaults.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpDefaults.Size = New System.Drawing.Size(692, 454)
    Me.tbpDefaults.TabIndex = 3
    Me.tbpDefaults.Text = "Defaults"
    Me.tbpDefaults.UseVisualStyleBackColor = True
    '
    'tbpCarriage
    '
    Me.tbpCarriage.Controls.Add(Me.tbpBatches)
    Me.tbpCarriage.Controls.Add(Me.tbpTransactions)
    Me.tbpCarriage.Controls.Add(Me.tbpAnalysisDefaults)
    Me.tbpCarriage.Controls.Add(Me.TabPage1)
    Me.tbpCarriage.Controls.Add(Me.tbpMembers)
    Me.tbpCarriage.Controls.Add(Me.tbpExams)
    Me.tbpCarriage.Location = New System.Drawing.Point(7, 6)
    Me.tbpCarriage.Name = "tbpCarriage"
    Me.tbpCarriage.SelectedIndex = 0
    Me.tbpCarriage.Size = New System.Drawing.Size(684, 410)
    Me.tbpCarriage.TabIndex = 0
    '
    'tbpBatches
    '
    Me.tbpBatches.Controls.Add(Me.pnlBatches)
    Me.tbpBatches.Location = New System.Drawing.Point(4, 22)
    Me.tbpBatches.Name = "tbpBatches"
    Me.tbpBatches.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpBatches.Size = New System.Drawing.Size(676, 384)
    Me.tbpBatches.TabIndex = 0
    Me.tbpBatches.Text = "Batches"
    Me.tbpBatches.UseVisualStyleBackColor = True
    '
    'pnlBatches
    '
    Me.pnlBatches.BackColor = System.Drawing.Color.Transparent
    Me.pnlBatches.Controls.Add(Me.lblBatchCategory)
    Me.pnlBatches.Controls.Add(Me.txtBatchCategory)
    Me.pnlBatches.Controls.Add(Me.lblBatchAnalysis)
    Me.pnlBatches.Controls.Add(Me.txtBatchAnalysisCode)
    Me.pnlBatches.Location = New System.Drawing.Point(0, 0)
    Me.pnlBatches.Name = "pnlBatches"
    Me.pnlBatches.Size = New System.Drawing.Size(684, 381)
    Me.pnlBatches.TabIndex = 32
    '
    'lblBatchCategory
    '
    Me.lblBatchCategory.AutoSize = True
    Me.lblBatchCategory.Location = New System.Drawing.Point(21, 15)
    Me.lblBatchCategory.Name = "lblBatchCategory"
    Me.lblBatchCategory.Size = New System.Drawing.Size(80, 13)
    Me.lblBatchCategory.TabIndex = 30
    Me.lblBatchCategory.Text = "Batch Category"
    '
    'txtBatchCategory
    '
    Me.txtBatchCategory.ActiveOnly = False
    Me.txtBatchCategory.BackColor = System.Drawing.SystemColors.Control
    Me.txtBatchCategory.CustomFormNumber = 0
    Me.txtBatchCategory.Description = ""
    Me.txtBatchCategory.EnabledProperty = True
    Me.txtBatchCategory.ExamCentreId = 0
    Me.txtBatchCategory.ExamCentreUnitId = 0
    Me.txtBatchCategory.ExamUnitLinkId = 0
    Me.txtBatchCategory.HasDependancies = False
    Me.txtBatchCategory.IsDesign = False
    Me.txtBatchCategory.Location = New System.Drawing.Point(177, 15)
    Me.txtBatchCategory.MaxLength = 32767
    Me.txtBatchCategory.MultipleValuesSupported = False
    Me.txtBatchCategory.Name = "txtBatchCategory"
    Me.txtBatchCategory.OriginalText = Nothing
    Me.txtBatchCategory.PreventHistoricalSelection = False
    Me.txtBatchCategory.ReadOnlyProperty = False
    Me.txtBatchCategory.Size = New System.Drawing.Size(408, 24)
    Me.txtBatchCategory.TabIndex = 28
    Me.txtBatchCategory.TextReadOnly = False
    Me.txtBatchCategory.TotalWidth = 408
    Me.txtBatchCategory.ValidationRequired = True
    Me.txtBatchCategory.WarningMessage = Nothing
    '
    'lblBatchAnalysis
    '
    Me.lblBatchAnalysis.AutoSize = True
    Me.lblBatchAnalysis.Location = New System.Drawing.Point(21, 56)
    Me.lblBatchAnalysis.Name = "lblBatchAnalysis"
    Me.lblBatchAnalysis.Size = New System.Drawing.Size(104, 13)
    Me.lblBatchAnalysis.TabIndex = 31
    Me.lblBatchAnalysis.Text = "Batch Analysis Code"
    '
    'txtBatchAnalysisCode
    '
    Me.txtBatchAnalysisCode.ActiveOnly = False
    Me.txtBatchAnalysisCode.BackColor = System.Drawing.SystemColors.Control
    Me.txtBatchAnalysisCode.CustomFormNumber = 0
    Me.txtBatchAnalysisCode.Description = ""
    Me.txtBatchAnalysisCode.EnabledProperty = True
    Me.txtBatchAnalysisCode.ExamCentreId = 0
    Me.txtBatchAnalysisCode.ExamCentreUnitId = 0
    Me.txtBatchAnalysisCode.ExamUnitLinkId = 0
    Me.txtBatchAnalysisCode.HasDependancies = False
    Me.txtBatchAnalysisCode.IsDesign = False
    Me.txtBatchAnalysisCode.Location = New System.Drawing.Point(177, 56)
    Me.txtBatchAnalysisCode.MaxLength = 32767
    Me.txtBatchAnalysisCode.MultipleValuesSupported = False
    Me.txtBatchAnalysisCode.Name = "txtBatchAnalysisCode"
    Me.txtBatchAnalysisCode.OriginalText = Nothing
    Me.txtBatchAnalysisCode.PreventHistoricalSelection = False
    Me.txtBatchAnalysisCode.ReadOnlyProperty = False
    Me.txtBatchAnalysisCode.Size = New System.Drawing.Size(408, 24)
    Me.txtBatchAnalysisCode.TabIndex = 29
    Me.txtBatchAnalysisCode.TextReadOnly = False
    Me.txtBatchAnalysisCode.TotalWidth = 408
    Me.txtBatchAnalysisCode.ValidationRequired = True
    Me.txtBatchAnalysisCode.WarningMessage = Nothing
    '
    'tbpTransactions
    '
    Me.tbpTransactions.Controls.Add(Me.pnlTransactions)
    Me.tbpTransactions.Location = New System.Drawing.Point(4, 22)
    Me.tbpTransactions.Name = "tbpTransactions"
    Me.tbpTransactions.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpTransactions.Size = New System.Drawing.Size(676, 384)
    Me.tbpTransactions.TabIndex = 1
    Me.tbpTransactions.Text = "Transactions"
    Me.tbpTransactions.UseVisualStyleBackColor = True
    '
    'pnlTransactions
    '
    Me.pnlTransactions.BackColor = System.Drawing.Color.Transparent
    Me.pnlTransactions.Controls.Add(Me.lblCreditCategory)
    Me.pnlTransactions.Controls.Add(Me.txtCreditCategory)
    Me.pnlTransactions.Controls.Add(Me.chkDefaultSourceFromLastMailing)
    Me.pnlTransactions.Controls.Add(Me.lblSource)
    Me.pnlTransactions.Controls.Add(Me.txtSource)
    Me.pnlTransactions.Controls.Add(Me.lblSalesPerson)
    Me.pnlTransactions.Controls.Add(Me.txtSalesPerson)
    Me.pnlTransactions.Controls.Add(Me.lblLinkToCommunication)
    Me.pnlTransactions.Controls.Add(Me.txtLinkToCommunication)
    Me.pnlTransactions.Controls.Add(Me.lblTransactionOrigin)
    Me.pnlTransactions.Controls.Add(Me.txtTransactionOrigin)
    Me.pnlTransactions.Controls.Add(Me.chkLinkMALToEvent)
    Me.pnlTransactions.Controls.Add(Me.chkLinkMALToService)
    Me.pnlTransactions.Controls.Add(Me.chkLinkAnalysisLines)
    Me.pnlTransactions.Controls.Add(Me.chkInvoicePrintPreview)
    Me.pnlTransactions.Location = New System.Drawing.Point(0, 0)
    Me.pnlTransactions.Name = "pnlTransactions"
    Me.pnlTransactions.Size = New System.Drawing.Size(680, 381)
    Me.pnlTransactions.TabIndex = 41
    '
    'lblCreditCategory
    '
    Me.lblCreditCategory.AutoSize = True
    Me.lblCreditCategory.Location = New System.Drawing.Point(16, 220)
    Me.lblCreditCategory.Name = "lblCreditCategory"
    Me.lblCreditCategory.Size = New System.Drawing.Size(79, 13)
    Me.lblCreditCategory.TabIndex = 40
    Me.lblCreditCategory.Tag = ""
    Me.lblCreditCategory.Text = "Credit Category"
    '
    'txtCreditCategory
    '
    Me.txtCreditCategory.ActiveOnly = False
    Me.txtCreditCategory.BackColor = System.Drawing.SystemColors.Control
    Me.txtCreditCategory.CustomFormNumber = 0
    Me.txtCreditCategory.Description = ""
    Me.txtCreditCategory.EnabledProperty = True
    Me.txtCreditCategory.ExamCentreId = 0
    Me.txtCreditCategory.ExamCentreUnitId = 0
    Me.txtCreditCategory.ExamUnitLinkId = 0
    Me.txtCreditCategory.HasDependancies = False
    Me.txtCreditCategory.IsDesign = False
    Me.txtCreditCategory.Location = New System.Drawing.Point(197, 213)
    Me.txtCreditCategory.MaxLength = 32767
    Me.txtCreditCategory.MultipleValuesSupported = False
    Me.txtCreditCategory.Name = "txtCreditCategory"
    Me.txtCreditCategory.OriginalText = Nothing
    Me.txtCreditCategory.PreventHistoricalSelection = False
    Me.txtCreditCategory.ReadOnlyProperty = False
    Me.txtCreditCategory.Size = New System.Drawing.Size(408, 24)
    Me.txtCreditCategory.TabIndex = 39
    Me.txtCreditCategory.Tag = ""
    Me.txtCreditCategory.TextReadOnly = False
    Me.txtCreditCategory.TotalWidth = 408
    Me.txtCreditCategory.ValidationRequired = True
    Me.txtCreditCategory.WarningMessage = Nothing
    '
    'chkDefaultSourceFromLastMailing
    '
    Me.chkDefaultSourceFromLastMailing.AutoSize = True
    Me.chkDefaultSourceFromLastMailing.Location = New System.Drawing.Point(18, 18)
    Me.chkDefaultSourceFromLastMailing.Name = "chkDefaultSourceFromLastMailing"
    Me.chkDefaultSourceFromLastMailing.Size = New System.Drawing.Size(182, 17)
    Me.chkDefaultSourceFromLastMailing.TabIndex = 1
    Me.chkDefaultSourceFromLastMailing.Text = "Default Source From Last Mailing"
    Me.chkDefaultSourceFromLastMailing.UseVisualStyleBackColor = True
    '
    'lblSource
    '
    Me.lblSource.AutoSize = True
    Me.lblSource.Location = New System.Drawing.Point(15, 55)
    Me.lblSource.Name = "lblSource"
    Me.lblSource.Size = New System.Drawing.Size(41, 13)
    Me.lblSource.TabIndex = 33
    Me.lblSource.Text = "Source"
    '
    'txtSource
    '
    Me.txtSource.ActiveOnly = False
    Me.txtSource.BackColor = System.Drawing.SystemColors.Control
    Me.txtSource.CustomFormNumber = 0
    Me.txtSource.Description = ""
    Me.txtSource.EnabledProperty = True
    Me.txtSource.ExamCentreId = 0
    Me.txtSource.ExamCentreUnitId = 0
    Me.txtSource.ExamUnitLinkId = 0
    Me.txtSource.HasDependancies = False
    Me.txtSource.IsDesign = False
    Me.txtSource.Location = New System.Drawing.Point(196, 55)
    Me.txtSource.MaxLength = 32767
    Me.txtSource.MultipleValuesSupported = False
    Me.txtSource.Name = "txtSource"
    Me.txtSource.OriginalText = Nothing
    Me.txtSource.PreventHistoricalSelection = False
    Me.txtSource.ReadOnlyProperty = False
    Me.txtSource.Size = New System.Drawing.Size(408, 24)
    Me.txtSource.TabIndex = 2
    Me.txtSource.TextReadOnly = False
    Me.txtSource.TotalWidth = 408
    Me.txtSource.ValidationRequired = True
    Me.txtSource.WarningMessage = Nothing
    '
    'lblSalesPerson
    '
    Me.lblSalesPerson.AutoSize = True
    Me.lblSalesPerson.Location = New System.Drawing.Point(15, 96)
    Me.lblSalesPerson.Name = "lblSalesPerson"
    Me.lblSalesPerson.Size = New System.Drawing.Size(69, 13)
    Me.lblSalesPerson.TabIndex = 32
    Me.lblSalesPerson.Text = "Sales Person"
    '
    'txtSalesPerson
    '
    Me.txtSalesPerson.ActiveOnly = False
    Me.txtSalesPerson.BackColor = System.Drawing.SystemColors.Control
    Me.txtSalesPerson.CustomFormNumber = 0
    Me.txtSalesPerson.Description = ""
    Me.txtSalesPerson.EnabledProperty = True
    Me.txtSalesPerson.ExamCentreId = 0
    Me.txtSalesPerson.ExamCentreUnitId = 0
    Me.txtSalesPerson.ExamUnitLinkId = 0
    Me.txtSalesPerson.HasDependancies = False
    Me.txtSalesPerson.IsDesign = False
    Me.txtSalesPerson.Location = New System.Drawing.Point(196, 96)
    Me.txtSalesPerson.MaxLength = 32767
    Me.txtSalesPerson.MultipleValuesSupported = False
    Me.txtSalesPerson.Name = "txtSalesPerson"
    Me.txtSalesPerson.OriginalText = Nothing
    Me.txtSalesPerson.PreventHistoricalSelection = False
    Me.txtSalesPerson.ReadOnlyProperty = False
    Me.txtSalesPerson.Size = New System.Drawing.Size(408, 24)
    Me.txtSalesPerson.TabIndex = 3
    Me.txtSalesPerson.TextReadOnly = False
    Me.txtSalesPerson.TotalWidth = 408
    Me.txtSalesPerson.ValidationRequired = True
    Me.txtSalesPerson.WarningMessage = Nothing
    '
    'lblLinkToCommunication
    '
    Me.lblLinkToCommunication.AutoSize = True
    Me.lblLinkToCommunication.Location = New System.Drawing.Point(15, 139)
    Me.lblLinkToCommunication.Name = "lblLinkToCommunication"
    Me.lblLinkToCommunication.Size = New System.Drawing.Size(114, 13)
    Me.lblLinkToCommunication.TabIndex = 37
    Me.lblLinkToCommunication.Text = "Link to Communication"
    '
    'txtLinkToCommunication
    '
    Me.txtLinkToCommunication.ActiveOnly = False
    Me.txtLinkToCommunication.BackColor = System.Drawing.SystemColors.Control
    Me.txtLinkToCommunication.CustomFormNumber = 0
    Me.txtLinkToCommunication.Description = ""
    Me.txtLinkToCommunication.EnabledProperty = True
    Me.txtLinkToCommunication.ExamCentreId = 0
    Me.txtLinkToCommunication.ExamCentreUnitId = 0
    Me.txtLinkToCommunication.ExamUnitLinkId = 0
    Me.txtLinkToCommunication.HasDependancies = False
    Me.txtLinkToCommunication.IsDesign = False
    Me.txtLinkToCommunication.Location = New System.Drawing.Point(196, 137)
    Me.txtLinkToCommunication.MaxLength = 32767
    Me.txtLinkToCommunication.MultipleValuesSupported = False
    Me.txtLinkToCommunication.Name = "txtLinkToCommunication"
    Me.txtLinkToCommunication.OriginalText = Nothing
    Me.txtLinkToCommunication.PreventHistoricalSelection = False
    Me.txtLinkToCommunication.ReadOnlyProperty = False
    Me.txtLinkToCommunication.Size = New System.Drawing.Size(408, 24)
    Me.txtLinkToCommunication.TabIndex = 4
    Me.txtLinkToCommunication.TextReadOnly = False
    Me.txtLinkToCommunication.TotalWidth = 408
    Me.txtLinkToCommunication.ValidationRequired = True
    Me.txtLinkToCommunication.WarningMessage = Nothing
    '
    'lblTransactionOrigin
    '
    Me.lblTransactionOrigin.AutoSize = True
    Me.lblTransactionOrigin.Location = New System.Drawing.Point(15, 185)
    Me.lblTransactionOrigin.Name = "lblTransactionOrigin"
    Me.lblTransactionOrigin.Size = New System.Drawing.Size(93, 13)
    Me.lblTransactionOrigin.TabIndex = 35
    Me.lblTransactionOrigin.Text = "Transaction Origin"
    '
    'txtTransactionOrigin
    '
    Me.txtTransactionOrigin.ActiveOnly = False
    Me.txtTransactionOrigin.BackColor = System.Drawing.SystemColors.Control
    Me.txtTransactionOrigin.CustomFormNumber = 0
    Me.txtTransactionOrigin.Description = ""
    Me.txtTransactionOrigin.EnabledProperty = True
    Me.txtTransactionOrigin.ExamCentreId = 0
    Me.txtTransactionOrigin.ExamCentreUnitId = 0
    Me.txtTransactionOrigin.ExamUnitLinkId = 0
    Me.txtTransactionOrigin.HasDependancies = False
    Me.txtTransactionOrigin.IsDesign = False
    Me.txtTransactionOrigin.Location = New System.Drawing.Point(196, 178)
    Me.txtTransactionOrigin.MaxLength = 32767
    Me.txtTransactionOrigin.MultipleValuesSupported = False
    Me.txtTransactionOrigin.Name = "txtTransactionOrigin"
    Me.txtTransactionOrigin.OriginalText = Nothing
    Me.txtTransactionOrigin.PreventHistoricalSelection = False
    Me.txtTransactionOrigin.ReadOnlyProperty = False
    Me.txtTransactionOrigin.Size = New System.Drawing.Size(408, 24)
    Me.txtTransactionOrigin.TabIndex = 5
    Me.txtTransactionOrigin.TextReadOnly = False
    Me.txtTransactionOrigin.TotalWidth = 408
    Me.txtTransactionOrigin.ValidationRequired = True
    Me.txtTransactionOrigin.WarningMessage = Nothing
    '
    'chkLinkMALToEvent
    '
    Me.chkLinkMALToEvent.AutoSize = True
    Me.chkLinkMALToEvent.Location = New System.Drawing.Point(18, 256)
    Me.chkLinkMALToEvent.Name = "chkLinkMALToEvent"
    Me.chkLinkMALToEvent.Size = New System.Drawing.Size(239, 17)
    Me.chkLinkMALToEvent.TabIndex = 6
    Me.chkLinkMALToEvent.Text = "Link Multiple Analysis Lines to Event Booking"
    Me.chkLinkMALToEvent.UseVisualStyleBackColor = True
    '
    'chkLinkMALToService
    '
    Me.chkLinkMALToService.AutoSize = True
    Me.chkLinkMALToService.Location = New System.Drawing.Point(18, 288)
    Me.chkLinkMALToService.Name = "chkLinkMALToService"
    Me.chkLinkMALToService.Size = New System.Drawing.Size(247, 17)
    Me.chkLinkMALToService.TabIndex = 7
    Me.chkLinkMALToService.Text = "Link Multiple Analysis Lines to Service Booking"
    Me.chkLinkMALToService.UseVisualStyleBackColor = True
    '
    'chkLinkAnalysisLines
    '
    Me.chkLinkAnalysisLines.AutoSize = True
    Me.chkLinkAnalysisLines.Location = New System.Drawing.Point(18, 320)
    Me.chkLinkAnalysisLines.Name = "chkLinkAnalysisLines"
    Me.chkLinkAnalysisLines.Size = New System.Drawing.Size(233, 17)
    Me.chkLinkAnalysisLines.TabIndex = 8
    Me.chkLinkAnalysisLines.Text = "Link Analysis Lines to Fundraising Payments"
    Me.chkLinkAnalysisLines.UseVisualStyleBackColor = True
    '
    'chkInvoicePrintPreview
    '
    Me.chkInvoicePrintPreview.AutoSize = True
    Me.chkInvoicePrintPreview.Location = New System.Drawing.Point(18, 352)
    Me.chkInvoicePrintPreview.Name = "chkInvoicePrintPreview"
    Me.chkInvoicePrintPreview.Size = New System.Drawing.Size(177, 17)
    Me.chkInvoicePrintPreview.TabIndex = 38
    Me.chkInvoicePrintPreview.Text = "Preview Invoices before printing"
    Me.chkInvoicePrintPreview.UseVisualStyleBackColor = True
    '
    'tbpAnalysisDefaults
    '
    Me.tbpAnalysisDefaults.Controls.Add(Me.pnlAnalysisSub)
    Me.tbpAnalysisDefaults.Location = New System.Drawing.Point(4, 22)
    Me.tbpAnalysisDefaults.Name = "tbpAnalysisDefaults"
    Me.tbpAnalysisDefaults.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpAnalysisDefaults.Size = New System.Drawing.Size(676, 384)
    Me.tbpAnalysisDefaults.TabIndex = 2
    Me.tbpAnalysisDefaults.Text = "Analysis"
    Me.tbpAnalysisDefaults.UseVisualStyleBackColor = True
    '
    'pnlAnalysisSub
    '
    Me.pnlAnalysisSub.BackColor = System.Drawing.Color.Transparent
    Me.pnlAnalysisSub.Controls.Add(Me.lblProduct)
    Me.pnlAnalysisSub.Controls.Add(Me.txtProduct)
    Me.pnlAnalysisSub.Controls.Add(Me.lblRate)
    Me.pnlAnalysisSub.Controls.Add(Me.txtRate)
    Me.pnlAnalysisSub.Controls.Add(Me.lblDonationProduct)
    Me.pnlAnalysisSub.Controls.Add(Me.txtDonationProduct)
    Me.pnlAnalysisSub.Controls.Add(Me.lblAnotherRate)
    Me.pnlAnalysisSub.Controls.Add(Me.txtDonationRate)
    Me.pnlAnalysisSub.Location = New System.Drawing.Point(0, 0)
    Me.pnlAnalysisSub.Name = "pnlAnalysisSub"
    Me.pnlAnalysisSub.Size = New System.Drawing.Size(680, 381)
    Me.pnlAnalysisSub.TabIndex = 46
    '
    'lblProduct
    '
    Me.lblProduct.AutoSize = True
    Me.lblProduct.Location = New System.Drawing.Point(21, 20)
    Me.lblProduct.Name = "lblProduct"
    Me.lblProduct.Size = New System.Drawing.Size(44, 13)
    Me.lblProduct.TabIndex = 41
    Me.lblProduct.Text = "Product"
    '
    'txtProduct
    '
    Me.txtProduct.ActiveOnly = False
    Me.txtProduct.BackColor = System.Drawing.SystemColors.Control
    Me.txtProduct.CustomFormNumber = 0
    Me.txtProduct.Description = ""
    Me.txtProduct.EnabledProperty = True
    Me.txtProduct.ExamCentreId = 0
    Me.txtProduct.ExamCentreUnitId = 0
    Me.txtProduct.ExamUnitLinkId = 0
    Me.txtProduct.HasDependancies = False
    Me.txtProduct.IsDesign = False
    Me.txtProduct.Location = New System.Drawing.Point(202, 20)
    Me.txtProduct.MaxLength = 32767
    Me.txtProduct.MultipleValuesSupported = False
    Me.txtProduct.Name = "txtProduct"
    Me.txtProduct.OriginalText = Nothing
    Me.txtProduct.PreventHistoricalSelection = False
    Me.txtProduct.ReadOnlyProperty = False
    Me.txtProduct.Size = New System.Drawing.Size(408, 24)
    Me.txtProduct.TabIndex = 38
    Me.txtProduct.TextReadOnly = False
    Me.txtProduct.TotalWidth = 408
    Me.txtProduct.ValidationRequired = True
    Me.txtProduct.WarningMessage = Nothing
    '
    'lblRate
    '
    Me.lblRate.AutoSize = True
    Me.lblRate.Location = New System.Drawing.Point(21, 61)
    Me.lblRate.Name = "lblRate"
    Me.lblRate.Size = New System.Drawing.Size(30, 13)
    Me.lblRate.TabIndex = 40
    Me.lblRate.Text = "Rate"
    '
    'txtRate
    '
    Me.txtRate.ActiveOnly = False
    Me.txtRate.BackColor = System.Drawing.SystemColors.Control
    Me.txtRate.CustomFormNumber = 0
    Me.txtRate.Description = ""
    Me.txtRate.EnabledProperty = True
    Me.txtRate.ExamCentreId = 0
    Me.txtRate.ExamCentreUnitId = 0
    Me.txtRate.ExamUnitLinkId = 0
    Me.txtRate.HasDependancies = False
    Me.txtRate.IsDesign = False
    Me.txtRate.Location = New System.Drawing.Point(202, 61)
    Me.txtRate.MaxLength = 32767
    Me.txtRate.MultipleValuesSupported = False
    Me.txtRate.Name = "txtRate"
    Me.txtRate.OriginalText = Nothing
    Me.txtRate.PreventHistoricalSelection = False
    Me.txtRate.ReadOnlyProperty = False
    Me.txtRate.Size = New System.Drawing.Size(408, 24)
    Me.txtRate.TabIndex = 39
    Me.txtRate.TextReadOnly = False
    Me.txtRate.TotalWidth = 408
    Me.txtRate.ValidationRequired = True
    Me.txtRate.WarningMessage = Nothing
    '
    'lblDonationProduct
    '
    Me.lblDonationProduct.AutoSize = True
    Me.lblDonationProduct.Location = New System.Drawing.Point(21, 104)
    Me.lblDonationProduct.Name = "lblDonationProduct"
    Me.lblDonationProduct.Size = New System.Drawing.Size(90, 13)
    Me.lblDonationProduct.TabIndex = 45
    Me.lblDonationProduct.Text = "Donation Product"
    '
    'txtDonationProduct
    '
    Me.txtDonationProduct.ActiveOnly = False
    Me.txtDonationProduct.BackColor = System.Drawing.SystemColors.Control
    Me.txtDonationProduct.CustomFormNumber = 0
    Me.txtDonationProduct.Description = ""
    Me.txtDonationProduct.EnabledProperty = True
    Me.txtDonationProduct.ExamCentreId = 0
    Me.txtDonationProduct.ExamCentreUnitId = 0
    Me.txtDonationProduct.ExamUnitLinkId = 0
    Me.txtDonationProduct.HasDependancies = False
    Me.txtDonationProduct.IsDesign = False
    Me.txtDonationProduct.Location = New System.Drawing.Point(202, 104)
    Me.txtDonationProduct.MaxLength = 32767
    Me.txtDonationProduct.MultipleValuesSupported = False
    Me.txtDonationProduct.Name = "txtDonationProduct"
    Me.txtDonationProduct.OriginalText = Nothing
    Me.txtDonationProduct.PreventHistoricalSelection = False
    Me.txtDonationProduct.ReadOnlyProperty = False
    Me.txtDonationProduct.Size = New System.Drawing.Size(408, 24)
    Me.txtDonationProduct.TabIndex = 40
    Me.txtDonationProduct.TextReadOnly = False
    Me.txtDonationProduct.TotalWidth = 408
    Me.txtDonationProduct.ValidationRequired = True
    Me.txtDonationProduct.WarningMessage = Nothing
    '
    'lblAnotherRate
    '
    Me.lblAnotherRate.AutoSize = True
    Me.lblAnotherRate.Location = New System.Drawing.Point(21, 150)
    Me.lblAnotherRate.Name = "lblAnotherRate"
    Me.lblAnotherRate.Size = New System.Drawing.Size(30, 13)
    Me.lblAnotherRate.TabIndex = 43
    Me.lblAnotherRate.Text = "Rate"
    '
    'txtDonationRate
    '
    Me.txtDonationRate.ActiveOnly = False
    Me.txtDonationRate.BackColor = System.Drawing.SystemColors.Control
    Me.txtDonationRate.CustomFormNumber = 0
    Me.txtDonationRate.Description = ""
    Me.txtDonationRate.EnabledProperty = True
    Me.txtDonationRate.ExamCentreId = 0
    Me.txtDonationRate.ExamCentreUnitId = 0
    Me.txtDonationRate.ExamUnitLinkId = 0
    Me.txtDonationRate.HasDependancies = False
    Me.txtDonationRate.IsDesign = False
    Me.txtDonationRate.Location = New System.Drawing.Point(202, 143)
    Me.txtDonationRate.MaxLength = 32767
    Me.txtDonationRate.MultipleValuesSupported = False
    Me.txtDonationRate.Name = "txtDonationRate"
    Me.txtDonationRate.OriginalText = Nothing
    Me.txtDonationRate.PreventHistoricalSelection = False
    Me.txtDonationRate.ReadOnlyProperty = False
    Me.txtDonationRate.Size = New System.Drawing.Size(408, 24)
    Me.txtDonationRate.TabIndex = 41
    Me.txtDonationRate.TextReadOnly = False
    Me.txtDonationRate.TotalWidth = 408
    Me.txtDonationRate.ValidationRequired = True
    Me.txtDonationRate.WarningMessage = Nothing
    '
    'TabPage1
    '
    Me.TabPage1.Controls.Add(Me.pnlCarriage)
    Me.TabPage1.Location = New System.Drawing.Point(4, 22)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(676, 384)
    Me.TabPage1.TabIndex = 3
    Me.TabPage1.Text = "Carriage"
    Me.TabPage1.UseVisualStyleBackColor = True
    '
    'pnlCarriage
    '
    Me.pnlCarriage.BackColor = System.Drawing.Color.Transparent
    Me.pnlCarriage.Controls.Add(Me.lblProductCarriage)
    Me.pnlCarriage.Controls.Add(Me.txtCarriageProduct)
    Me.pnlCarriage.Controls.Add(Me.lblRateCarriage)
    Me.pnlCarriage.Controls.Add(Me.txtCarriageRate)
    Me.pnlCarriage.Controls.Add(Me.lblPercentage)
    Me.pnlCarriage.Controls.Add(Me.txtPercentage)
    Me.pnlCarriage.Location = New System.Drawing.Point(0, 0)
    Me.pnlCarriage.Name = "pnlCarriage"
    Me.pnlCarriage.Size = New System.Drawing.Size(680, 381)
    Me.pnlCarriage.TabIndex = 51
    '
    'lblProductCarriage
    '
    Me.lblProductCarriage.AutoSize = True
    Me.lblProductCarriage.Location = New System.Drawing.Point(18, 19)
    Me.lblProductCarriage.Name = "lblProductCarriage"
    Me.lblProductCarriage.Size = New System.Drawing.Size(44, 13)
    Me.lblProductCarriage.TabIndex = 45
    Me.lblProductCarriage.Text = "Product"
    '
    'txtCarriageProduct
    '
    Me.txtCarriageProduct.ActiveOnly = False
    Me.txtCarriageProduct.BackColor = System.Drawing.SystemColors.Control
    Me.txtCarriageProduct.CustomFormNumber = 0
    Me.txtCarriageProduct.Description = ""
    Me.txtCarriageProduct.EnabledProperty = True
    Me.txtCarriageProduct.ExamCentreId = 0
    Me.txtCarriageProduct.ExamCentreUnitId = 0
    Me.txtCarriageProduct.ExamUnitLinkId = 0
    Me.txtCarriageProduct.HasDependancies = False
    Me.txtCarriageProduct.IsDesign = False
    Me.txtCarriageProduct.Location = New System.Drawing.Point(199, 19)
    Me.txtCarriageProduct.MaxLength = 32767
    Me.txtCarriageProduct.MultipleValuesSupported = False
    Me.txtCarriageProduct.Name = "txtCarriageProduct"
    Me.txtCarriageProduct.OriginalText = Nothing
    Me.txtCarriageProduct.PreventHistoricalSelection = False
    Me.txtCarriageProduct.ReadOnlyProperty = False
    Me.txtCarriageProduct.Size = New System.Drawing.Size(408, 24)
    Me.txtCarriageProduct.TabIndex = 42
    Me.txtCarriageProduct.TextReadOnly = False
    Me.txtCarriageProduct.TotalWidth = 408
    Me.txtCarriageProduct.ValidationRequired = True
    Me.txtCarriageProduct.WarningMessage = Nothing
    '
    'lblRateCarriage
    '
    Me.lblRateCarriage.AutoSize = True
    Me.lblRateCarriage.Location = New System.Drawing.Point(18, 60)
    Me.lblRateCarriage.Name = "lblRateCarriage"
    Me.lblRateCarriage.Size = New System.Drawing.Size(30, 13)
    Me.lblRateCarriage.TabIndex = 44
    Me.lblRateCarriage.Text = "Rate"
    '
    'txtCarriageRate
    '
    Me.txtCarriageRate.ActiveOnly = False
    Me.txtCarriageRate.BackColor = System.Drawing.SystemColors.Control
    Me.txtCarriageRate.CustomFormNumber = 0
    Me.txtCarriageRate.Description = ""
    Me.txtCarriageRate.EnabledProperty = True
    Me.txtCarriageRate.ExamCentreId = 0
    Me.txtCarriageRate.ExamCentreUnitId = 0
    Me.txtCarriageRate.ExamUnitLinkId = 0
    Me.txtCarriageRate.HasDependancies = False
    Me.txtCarriageRate.IsDesign = False
    Me.txtCarriageRate.Location = New System.Drawing.Point(199, 60)
    Me.txtCarriageRate.MaxLength = 32767
    Me.txtCarriageRate.MultipleValuesSupported = False
    Me.txtCarriageRate.Name = "txtCarriageRate"
    Me.txtCarriageRate.OriginalText = Nothing
    Me.txtCarriageRate.PreventHistoricalSelection = False
    Me.txtCarriageRate.ReadOnlyProperty = False
    Me.txtCarriageRate.Size = New System.Drawing.Size(408, 24)
    Me.txtCarriageRate.TabIndex = 43
    Me.txtCarriageRate.TextReadOnly = False
    Me.txtCarriageRate.TotalWidth = 408
    Me.txtCarriageRate.ValidationRequired = True
    Me.txtCarriageRate.WarningMessage = Nothing
    '
    'lblPercentage
    '
    Me.lblPercentage.AutoSize = True
    Me.lblPercentage.Location = New System.Drawing.Point(18, 107)
    Me.lblPercentage.Name = "lblPercentage"
    Me.lblPercentage.Size = New System.Drawing.Size(62, 13)
    Me.lblPercentage.TabIndex = 49
    Me.lblPercentage.Text = "Percentage"
    '
    'txtPercentage
    '
    Me.txtPercentage.Location = New System.Drawing.Point(199, 102)
    Me.txtPercentage.Name = "txtPercentage"
    Me.txtPercentage.Size = New System.Drawing.Size(96, 20)
    Me.txtPercentage.TabIndex = 44
    '
    'tbpMembers
    '
    Me.tbpMembers.Controls.Add(Me.pnlMembers)
    Me.tbpMembers.Location = New System.Drawing.Point(4, 22)
    Me.tbpMembers.Name = "tbpMembers"
    Me.tbpMembers.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpMembers.Size = New System.Drawing.Size(676, 384)
    Me.tbpMembers.TabIndex = 4
    Me.tbpMembers.Text = "Members"
    Me.tbpMembers.UseVisualStyleBackColor = True
    '
    'pnlMembers
    '
    Me.pnlMembers.BackColor = System.Drawing.Color.Transparent
    Me.pnlMembers.Controls.Add(Me.lblBranch)
    Me.pnlMembers.Controls.Add(Me.txtBranch)
    Me.pnlMembers.Location = New System.Drawing.Point(0, 0)
    Me.pnlMembers.Name = "pnlMembers"
    Me.pnlMembers.Size = New System.Drawing.Size(683, 381)
    Me.pnlMembers.TabIndex = 29
    '
    'lblBranch
    '
    Me.lblBranch.AutoSize = True
    Me.lblBranch.Location = New System.Drawing.Point(12, 21)
    Me.lblBranch.Name = "lblBranch"
    Me.lblBranch.Size = New System.Drawing.Size(41, 13)
    Me.lblBranch.TabIndex = 28
    Me.lblBranch.Text = "Branch"
    '
    'txtBranch
    '
    Me.txtBranch.ActiveOnly = False
    Me.txtBranch.BackColor = System.Drawing.SystemColors.Control
    Me.txtBranch.CustomFormNumber = 0
    Me.txtBranch.Description = ""
    Me.txtBranch.EnabledProperty = True
    Me.txtBranch.ExamCentreId = 0
    Me.txtBranch.ExamCentreUnitId = 0
    Me.txtBranch.ExamUnitLinkId = 0
    Me.txtBranch.HasDependancies = False
    Me.txtBranch.IsDesign = False
    Me.txtBranch.Location = New System.Drawing.Point(139, 21)
    Me.txtBranch.MaxLength = 32767
    Me.txtBranch.MultipleValuesSupported = False
    Me.txtBranch.Name = "txtBranch"
    Me.txtBranch.OriginalText = Nothing
    Me.txtBranch.PreventHistoricalSelection = False
    Me.txtBranch.ReadOnlyProperty = False
    Me.txtBranch.Size = New System.Drawing.Size(408, 24)
    Me.txtBranch.TabIndex = 27
    Me.txtBranch.TextReadOnly = False
    Me.txtBranch.TotalWidth = 408
    Me.txtBranch.ValidationRequired = True
    Me.txtBranch.WarningMessage = Nothing
    '
    'tbpExams
    '
    Me.tbpExams.Controls.Add(Me.pnlExams)
    Me.tbpExams.Location = New System.Drawing.Point(4, 22)
    Me.tbpExams.Name = "tbpExams"
    Me.tbpExams.Size = New System.Drawing.Size(676, 384)
    Me.tbpExams.TabIndex = 5
    Me.tbpExams.Text = "Exams"
    Me.tbpExams.UseVisualStyleBackColor = True
    '
    'pnlExams
    '
    Me.pnlExams.BackColor = System.Drawing.Color.Transparent
    Me.pnlExams.Controls.Add(Me.lblExamSession)
    Me.pnlExams.Controls.Add(Me.txtExamSession)
    Me.pnlExams.Controls.Add(Me.lblExamUnit)
    Me.pnlExams.Controls.Add(Me.txtExamUnit)
    Me.pnlExams.Location = New System.Drawing.Point(-3, 2)
    Me.pnlExams.Name = "pnlExams"
    Me.pnlExams.Size = New System.Drawing.Size(683, 381)
    Me.pnlExams.TabIndex = 30
    '
    'lblExamSession
    '
    Me.lblExamSession.AutoSize = True
    Me.lblExamSession.Location = New System.Drawing.Point(12, 21)
    Me.lblExamSession.Name = "lblExamSession"
    Me.lblExamSession.Size = New System.Drawing.Size(44, 13)
    Me.lblExamSession.TabIndex = 28
    Me.lblExamSession.Text = "Session"
    '
    'txtExamSession
    '
    Me.txtExamSession.ActiveOnly = False
    Me.txtExamSession.BackColor = System.Drawing.SystemColors.Control
    Me.txtExamSession.CustomFormNumber = 0
    Me.txtExamSession.Description = ""
    Me.txtExamSession.EnabledProperty = True
    Me.txtExamSession.ExamCentreId = 0
    Me.txtExamSession.ExamCentreUnitId = 0
    Me.txtExamSession.ExamUnitLinkId = 0
    Me.txtExamSession.HasDependancies = False
    Me.txtExamSession.IsDesign = False
    Me.txtExamSession.Location = New System.Drawing.Point(139, 21)
    Me.txtExamSession.MaxLength = 32767
    Me.txtExamSession.MultipleValuesSupported = False
    Me.txtExamSession.Name = "txtExamSession"
    Me.txtExamSession.OriginalText = Nothing
    Me.txtExamSession.PreventHistoricalSelection = False
    Me.txtExamSession.ReadOnlyProperty = False
    Me.txtExamSession.Size = New System.Drawing.Size(408, 24)
    Me.txtExamSession.TabIndex = 27
    Me.txtExamSession.TextReadOnly = False
    Me.txtExamSession.TotalWidth = 408
    Me.txtExamSession.ValidationRequired = True
    Me.txtExamSession.WarningMessage = Nothing
    '
    'lblExamUnit
    '
    Me.lblExamUnit.AutoSize = True
    Me.lblExamUnit.Location = New System.Drawing.Point(12, 51)
    Me.lblExamUnit.Name = "lblExamUnit"
    Me.lblExamUnit.Size = New System.Drawing.Size(40, 13)
    Me.lblExamUnit.TabIndex = 30
    Me.lblExamUnit.Text = "Course"
    '
    'txtExamUnit
    '
    Me.txtExamUnit.ActiveOnly = False
    Me.txtExamUnit.BackColor = System.Drawing.SystemColors.Control
    Me.txtExamUnit.CustomFormNumber = 0
    Me.txtExamUnit.Description = ""
    Me.txtExamUnit.EnabledProperty = True
    Me.txtExamUnit.ExamCentreId = 0
    Me.txtExamUnit.ExamCentreUnitId = 0
    Me.txtExamUnit.ExamUnitLinkId = 0
    Me.txtExamUnit.HasDependancies = False
    Me.txtExamUnit.IsDesign = False
    Me.txtExamUnit.Location = New System.Drawing.Point(139, 51)
    Me.txtExamUnit.MaxLength = 32767
    Me.txtExamUnit.MultipleValuesSupported = False
    Me.txtExamUnit.Name = "txtExamUnit"
    Me.txtExamUnit.OriginalText = Nothing
    Me.txtExamUnit.PreventHistoricalSelection = False
    Me.txtExamUnit.ReadOnlyProperty = False
    Me.txtExamUnit.Size = New System.Drawing.Size(408, 24)
    Me.txtExamUnit.TabIndex = 29
    Me.txtExamUnit.TextReadOnly = False
    Me.txtExamUnit.TotalWidth = 408
    Me.txtExamUnit.ValidationRequired = True
    Me.txtExamUnit.WarningMessage = Nothing
    '
    'tbpAnalysis
    '
    Me.tbpAnalysis.Controls.Add(Me.pnlAnalysis)
    Me.tbpAnalysis.Location = New System.Drawing.Point(4, 26)
    Me.tbpAnalysis.Name = "tbpAnalysis"
    Me.tbpAnalysis.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpAnalysis.Size = New System.Drawing.Size(692, 454)
    Me.tbpAnalysis.TabIndex = 2
    Me.tbpAnalysis.Text = "Analysis"
    Me.tbpAnalysis.UseVisualStyleBackColor = True
    '
    'pnlAnalysis
    '
    Me.pnlAnalysis.BackColor = System.Drawing.Color.Transparent
    Me.pnlAnalysis.Controls.Add(Me.TabAnalysis)
    Me.pnlAnalysis.Location = New System.Drawing.Point(0, 0)
    Me.pnlAnalysis.Name = "pnlAnalysis"
    Me.pnlAnalysis.Size = New System.Drawing.Size(693, 420)
    Me.pnlAnalysis.TabIndex = 1
    '
    'TabAnalysis
    '
    Me.TabAnalysis.Controls.Add(Me.tbpSales)
    Me.TabAnalysis.Controls.Add(Me.tbpPaymentPlans)
    Me.TabAnalysis.Controls.Add(Me.tbpSalesLedger)
    Me.TabAnalysis.Controls.Add(Me.tbpMaintenance)
    Me.TabAnalysis.Controls.Add(Me.tbpLegacies)
    Me.TabAnalysis.Location = New System.Drawing.Point(7, 22)
    Me.TabAnalysis.Name = "TabAnalysis"
    Me.TabAnalysis.SelectedIndex = 0
    Me.TabAnalysis.Size = New System.Drawing.Size(684, 394)
    Me.TabAnalysis.TabIndex = 0
    '
    'tbpSales
    '
    Me.tbpSales.Controls.Add(Me.pnlSales)
    Me.tbpSales.Location = New System.Drawing.Point(4, 22)
    Me.tbpSales.Name = "tbpSales"
    Me.tbpSales.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpSales.Size = New System.Drawing.Size(676, 368)
    Me.tbpSales.TabIndex = 0
    Me.tbpSales.Text = "Sales"
    Me.tbpSales.UseVisualStyleBackColor = True
    '
    'pnlSales
    '
    Me.pnlSales.BackColor = System.Drawing.Color.Transparent
    Me.pnlSales.Controls.Add(Me.chkConfirmSale)
    Me.pnlSales.Controls.Add(Me.chkConfirmCollection)
    Me.pnlSales.Controls.Add(Me.chkServiceBookingCredit)
    Me.pnlSales.Controls.Add(Me.chkDonation)
    Me.pnlSales.Controls.Add(Me.chkEventBooking)
    Me.pnlSales.Controls.Add(Me.chkExamBooking)
    Me.pnlSales.Controls.Add(Me.chkAccomodationBooking)
    Me.pnlSales.Controls.Add(Me.chkServiceBooking)
    Me.pnlSales.Controls.Add(Me.chkProduct)
    Me.pnlSales.Location = New System.Drawing.Point(0, 0)
    Me.pnlSales.Name = "pnlSales"
    Me.pnlSales.Size = New System.Drawing.Size(676, 361)
    Me.pnlSales.TabIndex = 14
    '
    'chkConfirmSale
    '
    Me.chkConfirmSale.AutoSize = True
    Me.chkConfirmSale.Location = New System.Drawing.Point(18, 235)
    Me.chkConfirmSale.Name = "chkConfirmSale"
    Me.chkConfirmSale.Size = New System.Drawing.Size(196, 17)
    Me.chkConfirmSale.TabIndex = 12
    Me.chkConfirmSale.Text = "Confirm Sale or Return Transactions"
    Me.chkConfirmSale.UseVisualStyleBackColor = True
    '
    'chkConfirmCollection
    '
    Me.chkConfirmCollection.AutoSize = True
    Me.chkConfirmCollection.Location = New System.Drawing.Point(18, 271)
    Me.chkConfirmCollection.Name = "chkConfirmCollection"
    Me.chkConfirmCollection.Size = New System.Drawing.Size(121, 17)
    Me.chkConfirmCollection.TabIndex = 13
    Me.chkConfirmCollection.Text = "Collection Payments"
    Me.chkConfirmCollection.UseVisualStyleBackColor = True
    '
    'chkServiceBookingCredit
    '
    Me.chkServiceBookingCredit.AutoSize = True
    Me.chkServiceBookingCredit.Location = New System.Drawing.Point(18, 197)
    Me.chkServiceBookingCredit.Name = "chkServiceBookingCredit"
    Me.chkServiceBookingCredit.Size = New System.Drawing.Size(134, 17)
    Me.chkServiceBookingCredit.TabIndex = 11
    Me.chkServiceBookingCredit.Text = "Service Booking Credit"
    Me.chkServiceBookingCredit.UseVisualStyleBackColor = True
    '
    'chkDonation
    '
    Me.chkDonation.AutoSize = True
    Me.chkDonation.Location = New System.Drawing.Point(18, 51)
    Me.chkDonation.Name = "chkDonation"
    Me.chkDonation.Size = New System.Drawing.Size(69, 17)
    Me.chkDonation.TabIndex = 7
    Me.chkDonation.Text = "Donation"
    Me.chkDonation.UseVisualStyleBackColor = True
    '
    'chkEventBooking
    '
    Me.chkEventBooking.AutoSize = True
    Me.chkEventBooking.Location = New System.Drawing.Point(18, 87)
    Me.chkEventBooking.Name = "chkEventBooking"
    Me.chkEventBooking.Size = New System.Drawing.Size(96, 17)
    Me.chkEventBooking.TabIndex = 8
    Me.chkEventBooking.Text = "Event Booking"
    Me.chkEventBooking.UseVisualStyleBackColor = True
    '
    'chkExamBooking
    '
    Me.chkExamBooking.AutoSize = True
    Me.chkExamBooking.Location = New System.Drawing.Point(18, 307)
    Me.chkExamBooking.Name = "chkExamBooking"
    Me.chkExamBooking.Size = New System.Drawing.Size(94, 17)
    Me.chkExamBooking.TabIndex = 8
    Me.chkExamBooking.Text = "Exam Booking"
    Me.chkExamBooking.UseVisualStyleBackColor = True
    '
    'chkAccomodationBooking
    '
    Me.chkAccomodationBooking.AutoSize = True
    Me.chkAccomodationBooking.Location = New System.Drawing.Point(18, 124)
    Me.chkAccomodationBooking.Name = "chkAccomodationBooking"
    Me.chkAccomodationBooking.Size = New System.Drawing.Size(136, 17)
    Me.chkAccomodationBooking.TabIndex = 9
    Me.chkAccomodationBooking.Text = "Accomodation Booking"
    Me.chkAccomodationBooking.UseVisualStyleBackColor = True
    '
    'chkServiceBooking
    '
    Me.chkServiceBooking.AutoSize = True
    Me.chkServiceBooking.Location = New System.Drawing.Point(18, 160)
    Me.chkServiceBooking.Name = "chkServiceBooking"
    Me.chkServiceBooking.Size = New System.Drawing.Size(104, 17)
    Me.chkServiceBooking.TabIndex = 10
    Me.chkServiceBooking.Text = "Service Booking"
    Me.chkServiceBooking.UseVisualStyleBackColor = True
    '
    'chkProduct
    '
    Me.chkProduct.AutoSize = True
    Me.chkProduct.Location = New System.Drawing.Point(18, 15)
    Me.chkProduct.Name = "chkProduct"
    Me.chkProduct.Size = New System.Drawing.Size(63, 17)
    Me.chkProduct.TabIndex = 6
    Me.chkProduct.Text = "Product"
    Me.chkProduct.UseVisualStyleBackColor = True
    '
    'tbpPaymentPlans
    '
    Me.tbpPaymentPlans.Controls.Add(Me.pnlPaymentPlans)
    Me.tbpPaymentPlans.Location = New System.Drawing.Point(4, 22)
    Me.tbpPaymentPlans.Name = "tbpPaymentPlans"
    Me.tbpPaymentPlans.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpPaymentPlans.Size = New System.Drawing.Size(676, 368)
    Me.tbpPaymentPlans.TabIndex = 1
    Me.tbpPaymentPlans.Text = "Payment Plans"
    Me.tbpPaymentPlans.UseVisualStyleBackColor = True
    '
    'pnlPaymentPlans
    '
    Me.pnlPaymentPlans.BackColor = System.Drawing.Color.Transparent
    Me.pnlPaymentPlans.Controls.Add(Me.chkDirectDebit)
    Me.pnlPaymentPlans.Controls.Add(Me.chkCreditCardAuthority)
    Me.pnlPaymentPlans.Controls.Add(Me.chkNoPaymentRequired)
    Me.pnlPaymentPlans.Controls.Add(Me.chkDisplayScheduledPayment)
    Me.pnlPaymentPlans.Controls.Add(Me.chkStandingOrder)
    Me.pnlPaymentPlans.Controls.Add(Me.chkCovenantedSubscription)
    Me.pnlPaymentPlans.Controls.Add(Me.chkCovenantedDonation)
    Me.pnlPaymentPlans.Controls.Add(Me.chkCOvenantedMembership)
    Me.pnlPaymentPlans.Controls.Add(Me.chkSubscription)
    Me.pnlPaymentPlans.Controls.Add(Me.chkRegularDonation)
    Me.pnlPaymentPlans.Controls.Add(Me.chkPayment)
    Me.pnlPaymentPlans.Controls.Add(Me.chkMembershipType)
    Me.pnlPaymentPlans.Controls.Add(Me.chkMembership)
    Me.pnlPaymentPlans.Controls.Add(Me.chkLoan)
    Me.pnlPaymentPlans.Location = New System.Drawing.Point(0, 0)
    Me.pnlPaymentPlans.Name = "pnlPaymentPlans"
    Me.pnlPaymentPlans.Size = New System.Drawing.Size(676, 365)
    Me.pnlPaymentPlans.TabIndex = 27
    '
    'chkDirectDebit
    '
    Me.chkDirectDebit.AutoSize = True
    Me.chkDirectDebit.Location = New System.Drawing.Point(454, 56)
    Me.chkDirectDebit.Name = "chkDirectDebit"
    Me.chkDirectDebit.Size = New System.Drawing.Size(82, 17)
    Me.chkDirectDebit.TabIndex = 23
    Me.chkDirectDebit.Text = "Direct Debit"
    Me.chkDirectDebit.UseVisualStyleBackColor = True
    '
    'chkCreditCardAuthority
    '
    Me.chkCreditCardAuthority.AutoSize = True
    Me.chkCreditCardAuthority.Location = New System.Drawing.Point(454, 92)
    Me.chkCreditCardAuthority.Name = "chkCreditCardAuthority"
    Me.chkCreditCardAuthority.Size = New System.Drawing.Size(122, 17)
    Me.chkCreditCardAuthority.TabIndex = 24
    Me.chkCreditCardAuthority.Text = "Credit Card Authority"
    Me.chkCreditCardAuthority.UseVisualStyleBackColor = True
    '
    'chkNoPaymentRequired
    '
    Me.chkNoPaymentRequired.AutoSize = True
    Me.chkNoPaymentRequired.Location = New System.Drawing.Point(454, 128)
    Me.chkNoPaymentRequired.Name = "chkNoPaymentRequired"
    Me.chkNoPaymentRequired.Size = New System.Drawing.Size(130, 17)
    Me.chkNoPaymentRequired.TabIndex = 25
    Me.chkNoPaymentRequired.Text = "No Payment Required"
    Me.chkNoPaymentRequired.UseVisualStyleBackColor = True
    '
    'chkDisplayScheduledPayment
    '
    Me.chkDisplayScheduledPayment.AutoSize = True
    Me.chkDisplayScheduledPayment.Location = New System.Drawing.Point(454, 164)
    Me.chkDisplayScheduledPayment.Name = "chkDisplayScheduledPayment"
    Me.chkDisplayScheduledPayment.Size = New System.Drawing.Size(158, 17)
    Me.chkDisplayScheduledPayment.TabIndex = 26
    Me.chkDisplayScheduledPayment.Text = "Display Scheduled Payment"
    Me.chkDisplayScheduledPayment.UseVisualStyleBackColor = True
    '
    'chkStandingOrder
    '
    Me.chkStandingOrder.AutoSize = True
    Me.chkStandingOrder.Location = New System.Drawing.Point(454, 20)
    Me.chkStandingOrder.Name = "chkStandingOrder"
    Me.chkStandingOrder.Size = New System.Drawing.Size(97, 17)
    Me.chkStandingOrder.TabIndex = 22
    Me.chkStandingOrder.Text = "Standing Order"
    Me.chkStandingOrder.UseVisualStyleBackColor = True
    '
    'chkCovenantedSubscription
    '
    Me.chkCovenantedSubscription.AutoSize = True
    Me.chkCovenantedSubscription.Location = New System.Drawing.Point(227, 56)
    Me.chkCovenantedSubscription.Name = "chkCovenantedSubscription"
    Me.chkCovenantedSubscription.Size = New System.Drawing.Size(145, 17)
    Me.chkCovenantedSubscription.TabIndex = 20
    Me.chkCovenantedSubscription.Text = "Covenanted Subscription"
    Me.chkCovenantedSubscription.UseVisualStyleBackColor = True
    '
    'chkCovenantedDonation
    '
    Me.chkCovenantedDonation.AutoSize = True
    Me.chkCovenantedDonation.Location = New System.Drawing.Point(227, 92)
    Me.chkCovenantedDonation.Name = "chkCovenantedDonation"
    Me.chkCovenantedDonation.Size = New System.Drawing.Size(130, 17)
    Me.chkCovenantedDonation.TabIndex = 21
    Me.chkCovenantedDonation.Text = "Covenanted Donation"
    Me.chkCovenantedDonation.UseVisualStyleBackColor = True
    '
    'chkCOvenantedMembership
    '
    Me.chkCOvenantedMembership.AutoSize = True
    Me.chkCOvenantedMembership.Location = New System.Drawing.Point(227, 20)
    Me.chkCOvenantedMembership.Name = "chkCOvenantedMembership"
    Me.chkCOvenantedMembership.Size = New System.Drawing.Size(144, 17)
    Me.chkCOvenantedMembership.TabIndex = 19
    Me.chkCOvenantedMembership.Text = "Covenanted Membership"
    Me.chkCOvenantedMembership.UseVisualStyleBackColor = True
    '
    'chkSubscription
    '
    Me.chkSubscription.AutoSize = True
    Me.chkSubscription.Location = New System.Drawing.Point(15, 56)
    Me.chkSubscription.Name = "chkSubscription"
    Me.chkSubscription.Size = New System.Drawing.Size(84, 17)
    Me.chkSubscription.TabIndex = 15
    Me.chkSubscription.Text = "Subscription"
    Me.chkSubscription.UseVisualStyleBackColor = True
    '
    'chkRegularDonation
    '
    Me.chkRegularDonation.AutoSize = True
    Me.chkRegularDonation.Location = New System.Drawing.Point(15, 92)
    Me.chkRegularDonation.Name = "chkRegularDonation"
    Me.chkRegularDonation.Size = New System.Drawing.Size(109, 17)
    Me.chkRegularDonation.TabIndex = 16
    Me.chkRegularDonation.Text = "Regular Donation"
    Me.chkRegularDonation.UseVisualStyleBackColor = True
    '
    'chkPayment
    '
    Me.chkPayment.AutoSize = True
    Me.chkPayment.Location = New System.Drawing.Point(15, 164)
    Me.chkPayment.Name = "chkPayment"
    Me.chkPayment.Size = New System.Drawing.Size(67, 17)
    Me.chkPayment.TabIndex = 17
    Me.chkPayment.Text = "Payment"
    Me.chkPayment.UseVisualStyleBackColor = True
    '
    'chkMembershipType
    '
    Me.chkMembershipType.AutoSize = True
    Me.chkMembershipType.Location = New System.Drawing.Point(15, 200)
    Me.chkMembershipType.Name = "chkMembershipType"
    Me.chkMembershipType.Size = New System.Drawing.Size(150, 17)
    Me.chkMembershipType.TabIndex = 18
    Me.chkMembershipType.Text = "Change Membership Type"
    Me.chkMembershipType.UseVisualStyleBackColor = True
    '
    'chkMembership
    '
    Me.chkMembership.AutoSize = True
    Me.chkMembership.Location = New System.Drawing.Point(15, 20)
    Me.chkMembership.Name = "chkMembership"
    Me.chkMembership.Size = New System.Drawing.Size(83, 17)
    Me.chkMembership.TabIndex = 14
    Me.chkMembership.Text = "Membership"
    Me.chkMembership.UseVisualStyleBackColor = True
    '
    'chkLoan
    '
    Me.chkLoan.AutoSize = True
    Me.chkLoan.Location = New System.Drawing.Point(15, 128)
    Me.chkLoan.Name = "chkLoan"
    Me.chkLoan.Size = New System.Drawing.Size(50, 17)
    Me.chkLoan.TabIndex = 27
    Me.chkLoan.Text = "Loan"
    Me.chkLoan.UseVisualStyleBackColor = True
    '
    'tbpSalesLedger
    '
    Me.tbpSalesLedger.Controls.Add(Me.pnlSalesLedger)
    Me.tbpSalesLedger.Location = New System.Drawing.Point(4, 22)
    Me.tbpSalesLedger.Name = "tbpSalesLedger"
    Me.tbpSalesLedger.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpSalesLedger.Size = New System.Drawing.Size(676, 368)
    Me.tbpSalesLedger.TabIndex = 2
    Me.tbpSalesLedger.Text = "Sales Ledger"
    Me.tbpSalesLedger.UseVisualStyleBackColor = True
    '
    'pnlSalesLedger
    '
    Me.pnlSalesLedger.BackColor = System.Drawing.Color.Transparent
    Me.pnlSalesLedger.Controls.Add(Me.chkInvoicePrintUnpostedBatches)
    Me.pnlSalesLedger.Controls.Add(Me.chkUnpostedBatchMsgInPrint)
    Me.pnlSalesLedger.Controls.Add(Me.chkDateRangeMsgInPrint)
    Me.pnlSalesLedger.Controls.Add(Me.chkAutoCreateCreditCustomer)
    Me.pnlSalesLedger.Controls.Add(Me.chkSundryCreditNote)
    Me.pnlSalesLedger.Controls.Add(Me.chkInvoicePayment)
    Me.pnlSalesLedger.Location = New System.Drawing.Point(0, 0)
    Me.pnlSalesLedger.Name = "pnlSalesLedger"
    Me.pnlSalesLedger.Size = New System.Drawing.Size(676, 365)
    Me.pnlSalesLedger.TabIndex = 21
    '
    'chkInvoicePrintUnpostedBatches
    '
    Me.chkInvoicePrintUnpostedBatches.AutoSize = True
    Me.chkInvoicePrintUnpostedBatches.Location = New System.Drawing.Point(18, 192)
    Me.chkInvoicePrintUnpostedBatches.Name = "chkInvoicePrintUnpostedBatches"
    Me.chkInvoicePrintUnpostedBatches.Size = New System.Drawing.Size(192, 17)
    Me.chkInvoicePrintUnpostedBatches.TabIndex = 23
    Me.chkInvoicePrintUnpostedBatches.Text = "Print Invoices in Unposted Batches"
    Me.chkInvoicePrintUnpostedBatches.UseVisualStyleBackColor = True
    '
    'chkUnpostedBatchMsgInPrint
    '
    Me.chkUnpostedBatchMsgInPrint.AutoSize = True
    Me.chkUnpostedBatchMsgInPrint.Location = New System.Drawing.Point(18, 124)
    Me.chkUnpostedBatchMsgInPrint.Name = "chkUnpostedBatchMsgInPrint"
    Me.chkUnpostedBatchMsgInPrint.Size = New System.Drawing.Size(185, 17)
    Me.chkUnpostedBatchMsgInPrint.TabIndex = 22
    Me.chkUnpostedBatchMsgInPrint.Text = "Unposted Batch Message In Print"
    Me.chkUnpostedBatchMsgInPrint.UseVisualStyleBackColor = True
    '
    'chkDateRangeMsgInPrint
    '
    Me.chkDateRangeMsgInPrint.AutoSize = True
    Me.chkDateRangeMsgInPrint.Checked = True
    Me.chkDateRangeMsgInPrint.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkDateRangeMsgInPrint.Location = New System.Drawing.Point(18, 158)
    Me.chkDateRangeMsgInPrint.Name = "chkDateRangeMsgInPrint"
    Me.chkDateRangeMsgInPrint.Size = New System.Drawing.Size(166, 17)
    Me.chkDateRangeMsgInPrint.TabIndex = 22
    Me.chkDateRangeMsgInPrint.Text = "Date Range Message In Print"
    Me.chkDateRangeMsgInPrint.UseVisualStyleBackColor = True
    '
    'chkAutoCreateCreditCustomer
    '
    Me.chkAutoCreateCreditCustomer.AutoSize = True
    Me.chkAutoCreateCreditCustomer.Location = New System.Drawing.Point(18, 90)
    Me.chkAutoCreateCreditCustomer.Name = "chkAutoCreateCreditCustomer"
    Me.chkAutoCreateCreditCustomer.Size = New System.Drawing.Size(159, 17)
    Me.chkAutoCreateCreditCustomer.TabIndex = 21
    Me.chkAutoCreateCreditCustomer.Text = "Auto Create Credit Customer"
    Me.chkAutoCreateCreditCustomer.UseVisualStyleBackColor = True
    '
    'chkSundryCreditNote
    '
    Me.chkSundryCreditNote.AutoSize = True
    Me.chkSundryCreditNote.Location = New System.Drawing.Point(18, 56)
    Me.chkSundryCreditNote.Name = "chkSundryCreditNote"
    Me.chkSundryCreditNote.Size = New System.Drawing.Size(115, 17)
    Me.chkSundryCreditNote.TabIndex = 20
    Me.chkSundryCreditNote.Text = "Sundry Credit Note"
    Me.chkSundryCreditNote.UseVisualStyleBackColor = True
    '
    'chkInvoicePayment
    '
    Me.chkInvoicePayment.AutoSize = True
    Me.chkInvoicePayment.Location = New System.Drawing.Point(18, 20)
    Me.chkInvoicePayment.Name = "chkInvoicePayment"
    Me.chkInvoicePayment.Size = New System.Drawing.Size(105, 17)
    Me.chkInvoicePayment.TabIndex = 19
    Me.chkInvoicePayment.Text = "Invoice Payment"
    Me.chkInvoicePayment.UseVisualStyleBackColor = True
    '
    'tbpMaintenance
    '
    Me.tbpMaintenance.Controls.Add(Me.pnlMaintenance)
    Me.tbpMaintenance.Location = New System.Drawing.Point(4, 22)
    Me.tbpMaintenance.Name = "tbpMaintenance"
    Me.tbpMaintenance.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpMaintenance.Size = New System.Drawing.Size(676, 368)
    Me.tbpMaintenance.TabIndex = 3
    Me.tbpMaintenance.Text = "Maintenance"
    Me.tbpMaintenance.UseVisualStyleBackColor = True
    '
    'pnlMaintenance
    '
    Me.pnlMaintenance.BackColor = System.Drawing.Color.Transparent
    Me.pnlMaintenance.Controls.Add(Me.chkGoneAway)
    Me.pnlMaintenance.Controls.Add(Me.chkAddressMaintenance)
    Me.pnlMaintenance.Controls.Add(Me.chkStatus)
    Me.pnlMaintenance.Controls.Add(Me.txtStatus)
    Me.pnlMaintenance.Controls.Add(Me.chkActivity)
    Me.pnlMaintenance.Controls.Add(Me.txtActivity)
    Me.pnlMaintenance.Controls.Add(Me.chkSuppression)
    Me.pnlMaintenance.Controls.Add(Me.txtSuppression)
    Me.pnlMaintenance.Controls.Add(Me.chkGiftAidDeclaration)
    Me.pnlMaintenance.Controls.Add(Me.chkCancelGiftAidDeclaration)
    Me.pnlMaintenance.Controls.Add(Me.txtCancelGiftAidDeclaration)
    Me.pnlMaintenance.Controls.Add(Me.chkCancelPaymentPlan)
    Me.pnlMaintenance.Controls.Add(Me.chkPayrollGiving)
    Me.pnlMaintenance.Controls.Add(Me.chkAutoPaymentMaintenance)
    Me.pnlMaintenance.Location = New System.Drawing.Point(0, 0)
    Me.pnlMaintenance.Name = "pnlMaintenance"
    Me.pnlMaintenance.Size = New System.Drawing.Size(678, 365)
    Me.pnlMaintenance.TabIndex = 28
    '
    'chkGoneAway
    '
    Me.chkGoneAway.AutoSize = True
    Me.chkGoneAway.Location = New System.Drawing.Point(15, 16)
    Me.chkGoneAway.Name = "chkGoneAway"
    Me.chkGoneAway.Size = New System.Drawing.Size(81, 17)
    Me.chkGoneAway.TabIndex = 14
    Me.chkGoneAway.Text = "Gone Away"
    Me.chkGoneAway.UseVisualStyleBackColor = True
    '
    'chkAddressMaintenance
    '
    Me.chkAddressMaintenance.AutoSize = True
    Me.chkAddressMaintenance.Location = New System.Drawing.Point(259, 16)
    Me.chkAddressMaintenance.Name = "chkAddressMaintenance"
    Me.chkAddressMaintenance.Size = New System.Drawing.Size(129, 17)
    Me.chkAddressMaintenance.TabIndex = 15
    Me.chkAddressMaintenance.Text = "Address Maintenance"
    Me.chkAddressMaintenance.UseVisualStyleBackColor = True
    '
    'chkStatus
    '
    Me.chkStatus.AutoSize = True
    Me.chkStatus.Location = New System.Drawing.Point(15, 43)
    Me.chkStatus.Name = "chkStatus"
    Me.chkStatus.Size = New System.Drawing.Size(56, 17)
    Me.chkStatus.TabIndex = 16
    Me.chkStatus.Text = "Status"
    Me.chkStatus.UseVisualStyleBackColor = True
    '
    'txtStatus
    '
    Me.txtStatus.ActiveOnly = False
    Me.txtStatus.BackColor = System.Drawing.SystemColors.Control
    Me.txtStatus.CustomFormNumber = 0
    Me.txtStatus.Description = ""
    Me.txtStatus.EnabledProperty = True
    Me.txtStatus.ExamCentreId = 0
    Me.txtStatus.ExamCentreUnitId = 0
    Me.txtStatus.ExamUnitLinkId = 0
    Me.txtStatus.HasDependancies = False
    Me.txtStatus.IsDesign = False
    Me.txtStatus.Location = New System.Drawing.Point(259, 39)
    Me.txtStatus.MaxLength = 32767
    Me.txtStatus.MultipleValuesSupported = False
    Me.txtStatus.Name = "txtStatus"
    Me.txtStatus.OriginalText = Nothing
    Me.txtStatus.PreventHistoricalSelection = False
    Me.txtStatus.ReadOnlyProperty = False
    Me.txtStatus.Size = New System.Drawing.Size(408, 24)
    Me.txtStatus.TabIndex = 24
    Me.txtStatus.TextReadOnly = False
    Me.txtStatus.TotalWidth = 408
    Me.txtStatus.ValidationRequired = True
    Me.txtStatus.WarningMessage = Nothing
    '
    'chkActivity
    '
    Me.chkActivity.AutoSize = True
    Me.chkActivity.Location = New System.Drawing.Point(15, 70)
    Me.chkActivity.Name = "chkActivity"
    Me.chkActivity.Size = New System.Drawing.Size(60, 17)
    Me.chkActivity.TabIndex = 17
    Me.chkActivity.Text = "Activity"
    Me.chkActivity.UseVisualStyleBackColor = True
    '
    'txtActivity
    '
    Me.txtActivity.ActiveOnly = False
    Me.txtActivity.BackColor = System.Drawing.SystemColors.Control
    Me.txtActivity.CustomFormNumber = 0
    Me.txtActivity.Description = ""
    Me.txtActivity.EnabledProperty = True
    Me.txtActivity.ExamCentreId = 0
    Me.txtActivity.ExamCentreUnitId = 0
    Me.txtActivity.ExamUnitLinkId = 0
    Me.txtActivity.HasDependancies = False
    Me.txtActivity.IsDesign = False
    Me.txtActivity.Location = New System.Drawing.Point(259, 70)
    Me.txtActivity.MaxLength = 32767
    Me.txtActivity.MultipleValuesSupported = False
    Me.txtActivity.Name = "txtActivity"
    Me.txtActivity.OriginalText = Nothing
    Me.txtActivity.PreventHistoricalSelection = False
    Me.txtActivity.ReadOnlyProperty = False
    Me.txtActivity.Size = New System.Drawing.Size(408, 24)
    Me.txtActivity.TabIndex = 25
    Me.txtActivity.TextReadOnly = False
    Me.txtActivity.TotalWidth = 408
    Me.txtActivity.ValidationRequired = True
    Me.txtActivity.WarningMessage = Nothing
    '
    'chkSuppression
    '
    Me.chkSuppression.AutoSize = True
    Me.chkSuppression.Location = New System.Drawing.Point(15, 97)
    Me.chkSuppression.Name = "chkSuppression"
    Me.chkSuppression.Size = New System.Drawing.Size(84, 17)
    Me.chkSuppression.TabIndex = 18
    Me.chkSuppression.Text = "Suppression"
    Me.chkSuppression.UseVisualStyleBackColor = True
    '
    'txtSuppression
    '
    Me.txtSuppression.ActiveOnly = False
    Me.txtSuppression.BackColor = System.Drawing.SystemColors.Control
    Me.txtSuppression.CustomFormNumber = 0
    Me.txtSuppression.Description = ""
    Me.txtSuppression.EnabledProperty = True
    Me.txtSuppression.ExamCentreId = 0
    Me.txtSuppression.ExamCentreUnitId = 0
    Me.txtSuppression.ExamUnitLinkId = 0
    Me.txtSuppression.HasDependancies = False
    Me.txtSuppression.IsDesign = False
    Me.txtSuppression.Location = New System.Drawing.Point(259, 100)
    Me.txtSuppression.MaxLength = 32767
    Me.txtSuppression.MultipleValuesSupported = False
    Me.txtSuppression.Name = "txtSuppression"
    Me.txtSuppression.OriginalText = Nothing
    Me.txtSuppression.PreventHistoricalSelection = False
    Me.txtSuppression.ReadOnlyProperty = False
    Me.txtSuppression.Size = New System.Drawing.Size(408, 24)
    Me.txtSuppression.TabIndex = 26
    Me.txtSuppression.TextReadOnly = False
    Me.txtSuppression.TotalWidth = 408
    Me.txtSuppression.ValidationRequired = True
    Me.txtSuppression.WarningMessage = Nothing
    '
    'chkGiftAidDeclaration
    '
    Me.chkGiftAidDeclaration.AutoSize = True
    Me.chkGiftAidDeclaration.Location = New System.Drawing.Point(15, 124)
    Me.chkGiftAidDeclaration.Name = "chkGiftAidDeclaration"
    Me.chkGiftAidDeclaration.Size = New System.Drawing.Size(117, 17)
    Me.chkGiftAidDeclaration.TabIndex = 19
    Me.chkGiftAidDeclaration.Text = "Gift Aid Declaration"
    Me.chkGiftAidDeclaration.UseVisualStyleBackColor = True
    '
    'chkCancelGiftAidDeclaration
    '
    Me.chkCancelGiftAidDeclaration.AutoSize = True
    Me.chkCancelGiftAidDeclaration.Location = New System.Drawing.Point(15, 151)
    Me.chkCancelGiftAidDeclaration.Name = "chkCancelGiftAidDeclaration"
    Me.chkCancelGiftAidDeclaration.Size = New System.Drawing.Size(153, 17)
    Me.chkCancelGiftAidDeclaration.TabIndex = 20
    Me.chkCancelGiftAidDeclaration.Text = "Cancel Gift Aid Declaration"
    Me.chkCancelGiftAidDeclaration.UseVisualStyleBackColor = True
    '
    'txtCancelGiftAidDeclaration
    '
    Me.txtCancelGiftAidDeclaration.ActiveOnly = False
    Me.txtCancelGiftAidDeclaration.BackColor = System.Drawing.SystemColors.Control
    Me.txtCancelGiftAidDeclaration.CustomFormNumber = 0
    Me.txtCancelGiftAidDeclaration.Description = ""
    Me.txtCancelGiftAidDeclaration.EnabledProperty = True
    Me.txtCancelGiftAidDeclaration.ExamCentreId = 0
    Me.txtCancelGiftAidDeclaration.ExamCentreUnitId = 0
    Me.txtCancelGiftAidDeclaration.ExamUnitLinkId = 0
    Me.txtCancelGiftAidDeclaration.HasDependancies = False
    Me.txtCancelGiftAidDeclaration.IsDesign = False
    Me.txtCancelGiftAidDeclaration.Location = New System.Drawing.Point(259, 151)
    Me.txtCancelGiftAidDeclaration.MaxLength = 32767
    Me.txtCancelGiftAidDeclaration.MultipleValuesSupported = False
    Me.txtCancelGiftAidDeclaration.Name = "txtCancelGiftAidDeclaration"
    Me.txtCancelGiftAidDeclaration.OriginalText = Nothing
    Me.txtCancelGiftAidDeclaration.PreventHistoricalSelection = False
    Me.txtCancelGiftAidDeclaration.ReadOnlyProperty = False
    Me.txtCancelGiftAidDeclaration.Size = New System.Drawing.Size(408, 24)
    Me.txtCancelGiftAidDeclaration.TabIndex = 27
    Me.txtCancelGiftAidDeclaration.TextReadOnly = False
    Me.txtCancelGiftAidDeclaration.TotalWidth = 408
    Me.txtCancelGiftAidDeclaration.ValidationRequired = True
    Me.txtCancelGiftAidDeclaration.WarningMessage = Nothing
    '
    'chkCancelPaymentPlan
    '
    Me.chkCancelPaymentPlan.AutoSize = True
    Me.chkCancelPaymentPlan.Location = New System.Drawing.Point(15, 178)
    Me.chkCancelPaymentPlan.Name = "chkCancelPaymentPlan"
    Me.chkCancelPaymentPlan.Size = New System.Drawing.Size(127, 17)
    Me.chkCancelPaymentPlan.TabIndex = 21
    Me.chkCancelPaymentPlan.Text = "Cancel Payment Plan"
    Me.chkCancelPaymentPlan.UseVisualStyleBackColor = True
    '
    'chkPayrollGiving
    '
    Me.chkPayrollGiving.AutoSize = True
    Me.chkPayrollGiving.Location = New System.Drawing.Point(15, 205)
    Me.chkPayrollGiving.Name = "chkPayrollGiving"
    Me.chkPayrollGiving.Size = New System.Drawing.Size(90, 17)
    Me.chkPayrollGiving.TabIndex = 22
    Me.chkPayrollGiving.Text = "Payroll Giving"
    Me.chkPayrollGiving.UseVisualStyleBackColor = True
    '
    'chkAutoPaymentMaintenance
    '
    Me.chkAutoPaymentMaintenance.AutoSize = True
    Me.chkAutoPaymentMaintenance.Location = New System.Drawing.Point(15, 232)
    Me.chkAutoPaymentMaintenance.Name = "chkAutoPaymentMaintenance"
    Me.chkAutoPaymentMaintenance.Size = New System.Drawing.Size(157, 17)
    Me.chkAutoPaymentMaintenance.TabIndex = 23
    Me.chkAutoPaymentMaintenance.Text = "Auto Payment Maintenance"
    Me.chkAutoPaymentMaintenance.UseVisualStyleBackColor = True
    '
    'tbpLegacies
    '
    Me.tbpLegacies.Controls.Add(Me.pnlLegacies)
    Me.tbpLegacies.Location = New System.Drawing.Point(4, 22)
    Me.tbpLegacies.Name = "tbpLegacies"
    Me.tbpLegacies.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpLegacies.Size = New System.Drawing.Size(676, 368)
    Me.tbpLegacies.TabIndex = 4
    Me.tbpLegacies.Text = "Legacies"
    Me.tbpLegacies.UseVisualStyleBackColor = True
    '
    'pnlLegacies
    '
    Me.pnlLegacies.BackColor = System.Drawing.Color.Transparent
    Me.pnlLegacies.Controls.Add(Me.chkLegacyReceipt)
    Me.pnlLegacies.Location = New System.Drawing.Point(0, 0)
    Me.pnlLegacies.Name = "pnlLegacies"
    Me.pnlLegacies.Size = New System.Drawing.Size(683, 365)
    Me.pnlLegacies.TabIndex = 16
    '
    'chkLegacyReceipt
    '
    Me.chkLegacyReceipt.AutoSize = True
    Me.chkLegacyReceipt.Location = New System.Drawing.Point(15, 20)
    Me.chkLegacyReceipt.Name = "chkLegacyReceipt"
    Me.chkLegacyReceipt.Size = New System.Drawing.Size(101, 17)
    Me.chkLegacyReceipt.TabIndex = 15
    Me.chkLegacyReceipt.Text = "Legacy Receipt"
    Me.chkLegacyReceipt.UseVisualStyleBackColor = True
    '
    'tabMain
    '
    Me.tabMain.Controls.Add(Me.tbpGeneral)
    Me.tabMain.Controls.Add(Me.tbpAnalysis)
    Me.tabMain.Controls.Add(Me.tbpDefaults)
    Me.tabMain.Controls.Add(Me.tbpBank)
    Me.tabMain.Controls.Add(Me.tbpDocuments)
    Me.tabMain.Controls.Add(Me.tbpRestrictions)
    Me.tabMain.Controls.Add(Me.tbpCurrency)
    Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tabMain.ItemSize = New System.Drawing.Size(99, 22)
    Me.tabMain.Location = New System.Drawing.Point(0, 0)
    Me.tabMain.Name = "tabMain"
    Me.tabMain.SelectedIndex = 0
    Me.tabMain.Size = New System.Drawing.Size(700, 484)
    Me.tabMain.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tabMain.TabIndex = 0
    '
    'tbpGeneral
    '
    Me.tbpGeneral.Controls.Add(Me.pnlGeneral)
    Me.tbpGeneral.Location = New System.Drawing.Point(4, 26)
    Me.tbpGeneral.Name = "tbpGeneral"
    Me.tbpGeneral.Padding = New System.Windows.Forms.Padding(1)
    Me.tbpGeneral.Size = New System.Drawing.Size(692, 454)
    Me.tbpGeneral.TabIndex = 8
    Me.tbpGeneral.Text = "General"
    Me.tbpGeneral.UseVisualStyleBackColor = True
    '
    'pnlGeneral
    '
    Me.pnlGeneral.BackColor = System.Drawing.Color.Transparent
    Me.pnlGeneral.Controls.Add(Me.lblType)
    Me.pnlGeneral.Controls.Add(Me.txtType)
    Me.pnlGeneral.Controls.Add(Me.txtAppDesc)
    Me.pnlGeneral.Controls.Add(Me.lblMenu)
    Me.pnlGeneral.Controls.Add(Me.txtApplication)
    Me.pnlGeneral.Controls.Add(Me.lblApplication)
    Me.pnlGeneral.Controls.Add(Me.TabGeneral1)
    Me.pnlGeneral.Location = New System.Drawing.Point(0, 0)
    Me.pnlGeneral.Name = "pnlGeneral"
    Me.pnlGeneral.Size = New System.Drawing.Size(696, 414)
    Me.pnlGeneral.TabIndex = 18
    '
    'lblType
    '
    Me.lblType.AutoSize = True
    Me.lblType.Location = New System.Drawing.Point(21, 62)
    Me.lblType.Name = "lblType"
    Me.lblType.Size = New System.Drawing.Size(31, 13)
    Me.lblType.TabIndex = 16
    Me.lblType.Text = "Type"
    '
    'txtType
    '
    Me.txtType.ActiveOnly = False
    Me.txtType.BackColor = System.Drawing.SystemColors.Control
    Me.txtType.CustomFormNumber = 0
    Me.txtType.Description = ""
    Me.txtType.EnabledProperty = True
    Me.txtType.ExamCentreId = 0
    Me.txtType.ExamCentreUnitId = 0
    Me.txtType.ExamUnitLinkId = 0
    Me.txtType.HasDependancies = False
    Me.txtType.IsDesign = False
    Me.txtType.Location = New System.Drawing.Point(111, 62)
    Me.txtType.MaxLength = 32767
    Me.txtType.MultipleValuesSupported = False
    Me.txtType.Name = "txtType"
    Me.txtType.OriginalText = Nothing
    Me.txtType.PreventHistoricalSelection = False
    Me.txtType.ReadOnlyProperty = False
    Me.txtType.Size = New System.Drawing.Size(560, 24)
    Me.txtType.TabIndex = 15
    Me.txtType.TextReadOnly = False
    Me.txtType.TotalWidth = 408
    Me.txtType.ValidationRequired = True
    Me.txtType.WarningMessage = Nothing
    '
    'txtAppDesc
    '
    Me.txtAppDesc.Location = New System.Drawing.Point(320, 23)
    Me.txtAppDesc.MaxLength = 50
    Me.txtAppDesc.Name = "txtAppDesc"
    Me.txtAppDesc.Size = New System.Drawing.Size(352, 20)
    Me.txtAppDesc.TabIndex = 14
    '
    'lblMenu
    '
    Me.lblMenu.AutoSize = True
    Me.lblMenu.Location = New System.Drawing.Point(240, 23)
    Me.lblMenu.Name = "lblMenu"
    Me.lblMenu.Size = New System.Drawing.Size(58, 13)
    Me.lblMenu.TabIndex = 13
    Me.lblMenu.Text = "Menu Text"
    '
    'txtApplication
    '
    Me.txtApplication.Enabled = False
    Me.txtApplication.Location = New System.Drawing.Point(111, 23)
    Me.txtApplication.Name = "txtApplication"
    Me.txtApplication.Size = New System.Drawing.Size(96, 20)
    Me.txtApplication.TabIndex = 12
    '
    'lblApplication
    '
    Me.lblApplication.AutoSize = True
    Me.lblApplication.Location = New System.Drawing.Point(21, 23)
    Me.lblApplication.Name = "lblApplication"
    Me.lblApplication.Size = New System.Drawing.Size(59, 13)
    Me.lblApplication.TabIndex = 11
    Me.lblApplication.Text = "Application"
    '
    'TabGeneral1
    '
    Me.TabGeneral1.Controls.Add(Me.tbpOptions)
    Me.TabGeneral1.Controls.Add(Me.tbpMethods)
    Me.TabGeneral1.Controls.Add(Me.TabPage2)
    Me.TabGeneral1.Controls.Add(Me.TabPage3)
    Me.TabGeneral1.Controls.Add(Me.TabPage4)
    Me.TabGeneral1.ItemSize = New System.Drawing.Size(58, 21)
    Me.TabGeneral1.Location = New System.Drawing.Point(24, 102)
    Me.TabGeneral1.Name = "TabGeneral1"
    Me.TabGeneral1.SelectedIndex = 0
    Me.TabGeneral1.Size = New System.Drawing.Size(652, 298)
    Me.TabGeneral1.TabIndex = 17
    '
    'tbpOptions
    '
    Me.tbpOptions.Controls.Add(Me.plnOption)
    Me.tbpOptions.Location = New System.Drawing.Point(4, 25)
    Me.tbpOptions.Name = "tbpOptions"
    Me.tbpOptions.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpOptions.Size = New System.Drawing.Size(644, 269)
    Me.tbpOptions.TabIndex = 0
    Me.tbpOptions.Text = "Options"
    Me.tbpOptions.UseVisualStyleBackColor = True
    '
    'plnOption
    '
    Me.plnOption.BackColor = System.Drawing.Color.Transparent
    Me.plnOption.Controls.Add(Me.chkAnalysisComments)
    Me.plnOption.Controls.Add(Me.chkConfirmDetails)
    Me.plnOption.Controls.Add(Me.chkAutoSetAmount)
    Me.plnOption.Controls.Add(Me.chkPrefulfilledIncentives)
    Me.plnOption.Controls.Add(Me.chkTransactionComment)
    Me.plnOption.Controls.Add(Me.chkMaintainPaymentPlan)
    Me.plnOption.Controls.Add(Me.chkConfirmProduct)
    Me.plnOption.Controls.Add(Me.chkConfirmAnalysis)
    Me.plnOption.Controls.Add(Me.ChkPaymentMethod)
    Me.plnOption.Controls.Add(Me.chkNonFinancialBatch)
    Me.plnOption.Controls.Add(Me.chkBankDetails)
    Me.plnOption.Controls.Add(Me.chkShowReference)
    Me.plnOption.Controls.Add(Me.chkCarriage)
    Me.plnOption.Controls.Add(Me.chkConfirmCarriage)
    Me.plnOption.Controls.Add(Me.chkBypass)
    Me.plnOption.Controls.Add(Me.chkSelectBatch)
    Me.plnOption.Controls.Add(Me.chkForeignCurrency)
    Me.plnOption.Location = New System.Drawing.Point(-4, 3)
    Me.plnOption.Name = "plnOption"
    Me.plnOption.Size = New System.Drawing.Size(652, 266)
    Me.plnOption.TabIndex = 0
    '
    'chkAnalysisComments
    '
    Me.chkAnalysisComments.AutoSize = True
    Me.chkAnalysisComments.Location = New System.Drawing.Point(449, 52)
    Me.chkAnalysisComments.Name = "chkAnalysisComments"
    Me.chkAnalysisComments.Size = New System.Drawing.Size(116, 17)
    Me.chkAnalysisComments.TabIndex = 31
    Me.chkAnalysisComments.Text = "Analysis Comments"
    Me.chkAnalysisComments.UseVisualStyleBackColor = True
    '
    'chkConfirmDetails
    '
    Me.chkConfirmDetails.AutoSize = True
    Me.chkConfirmDetails.Location = New System.Drawing.Point(449, 92)
    Me.chkConfirmDetails.Name = "chkConfirmDetails"
    Me.chkConfirmDetails.Size = New System.Drawing.Size(96, 17)
    Me.chkConfirmDetails.TabIndex = 32
    Me.chkConfirmDetails.Text = "Confirm Details"
    Me.chkConfirmDetails.UseVisualStyleBackColor = True
    '
    'chkAutoSetAmount
    '
    Me.chkAutoSetAmount.AutoSize = True
    Me.chkAutoSetAmount.Location = New System.Drawing.Point(449, 133)
    Me.chkAutoSetAmount.Name = "chkAutoSetAmount"
    Me.chkAutoSetAmount.Size = New System.Drawing.Size(106, 17)
    Me.chkAutoSetAmount.TabIndex = 34
    Me.chkAutoSetAmount.Text = "Auto Set Amount"
    Me.chkAutoSetAmount.UseVisualStyleBackColor = True
    '
    'chkPrefulfilledIncentives
    '
    Me.chkPrefulfilledIncentives.AutoSize = True
    Me.chkPrefulfilledIncentives.Location = New System.Drawing.Point(449, 170)
    Me.chkPrefulfilledIncentives.Name = "chkPrefulfilledIncentives"
    Me.chkPrefulfilledIncentives.Size = New System.Drawing.Size(126, 17)
    Me.chkPrefulfilledIncentives.TabIndex = 35
    Me.chkPrefulfilledIncentives.Text = "Prefulfilled Incentives"
    Me.chkPrefulfilledIncentives.UseVisualStyleBackColor = True
    '
    'chkTransactionComment
    '
    Me.chkTransactionComment.AutoSize = True
    Me.chkTransactionComment.Location = New System.Drawing.Point(449, 16)
    Me.chkTransactionComment.Name = "chkTransactionComment"
    Me.chkTransactionComment.Size = New System.Drawing.Size(134, 17)
    Me.chkTransactionComment.TabIndex = 30
    Me.chkTransactionComment.Text = "Transaction Comments"
    Me.chkTransactionComment.UseVisualStyleBackColor = True
    '
    'chkMaintainPaymentPlan
    '
    Me.chkMaintainPaymentPlan.AutoSize = True
    Me.chkMaintainPaymentPlan.Location = New System.Drawing.Point(13, 210)
    Me.chkMaintainPaymentPlan.Name = "chkMaintainPaymentPlan"
    Me.chkMaintainPaymentPlan.Size = New System.Drawing.Size(134, 17)
    Me.chkMaintainPaymentPlan.TabIndex = 29
    Me.chkMaintainPaymentPlan.Text = "Maintain Payment Plan"
    Me.chkMaintainPaymentPlan.UseVisualStyleBackColor = True
    '
    'chkConfirmProduct
    '
    Me.chkConfirmProduct.AutoSize = True
    Me.chkConfirmProduct.Location = New System.Drawing.Point(210, 52)
    Me.chkConfirmProduct.Name = "chkConfirmProduct"
    Me.chkConfirmProduct.Size = New System.Drawing.Size(101, 17)
    Me.chkConfirmProduct.TabIndex = 25
    Me.chkConfirmProduct.Text = "Confirm Product"
    Me.chkConfirmProduct.UseVisualStyleBackColor = True
    '
    'chkConfirmAnalysis
    '
    Me.chkConfirmAnalysis.AutoSize = True
    Me.chkConfirmAnalysis.Location = New System.Drawing.Point(210, 92)
    Me.chkConfirmAnalysis.Name = "chkConfirmAnalysis"
    Me.chkConfirmAnalysis.Size = New System.Drawing.Size(102, 17)
    Me.chkConfirmAnalysis.TabIndex = 26
    Me.chkConfirmAnalysis.Text = "Confirm Analysis"
    Me.chkConfirmAnalysis.UseVisualStyleBackColor = True
    '
    'ChkPaymentMethod
    '
    Me.ChkPaymentMethod.AutoSize = True
    Me.ChkPaymentMethod.Location = New System.Drawing.Point(210, 133)
    Me.ChkPaymentMethod.Name = "ChkPaymentMethod"
    Me.ChkPaymentMethod.Size = New System.Drawing.Size(145, 17)
    Me.ChkPaymentMethod.TabIndex = 27
    Me.ChkPaymentMethod.Text = "Payment Methods at End"
    Me.ChkPaymentMethod.UseVisualStyleBackColor = True
    '
    'chkNonFinancialBatch
    '
    Me.chkNonFinancialBatch.AutoSize = True
    Me.chkNonFinancialBatch.Location = New System.Drawing.Point(210, 170)
    Me.chkNonFinancialBatch.Name = "chkNonFinancialBatch"
    Me.chkNonFinancialBatch.Size = New System.Drawing.Size(122, 17)
    Me.chkNonFinancialBatch.TabIndex = 28
    Me.chkNonFinancialBatch.Text = "Non Financial Batch"
    Me.chkNonFinancialBatch.UseVisualStyleBackColor = True
    '
    'chkBankDetails
    '
    Me.chkBankDetails.AutoSize = True
    Me.chkBankDetails.Location = New System.Drawing.Point(210, 16)
    Me.chkBankDetails.Name = "chkBankDetails"
    Me.chkBankDetails.Size = New System.Drawing.Size(86, 17)
    Me.chkBankDetails.TabIndex = 24
    Me.chkBankDetails.Text = "Bank Details"
    Me.chkBankDetails.UseVisualStyleBackColor = True
    '
    'chkShowReference
    '
    Me.chkShowReference.AutoSize = True
    Me.chkShowReference.Location = New System.Drawing.Point(13, 52)
    Me.chkShowReference.Name = "chkShowReference"
    Me.chkShowReference.Size = New System.Drawing.Size(106, 17)
    Me.chkShowReference.TabIndex = 19
    Me.chkShowReference.Text = "Show Reference"
    Me.chkShowReference.UseVisualStyleBackColor = True
    '
    'chkCarriage
    '
    Me.chkCarriage.AutoSize = True
    Me.chkCarriage.Location = New System.Drawing.Point(13, 92)
    Me.chkCarriage.Name = "chkCarriage"
    Me.chkCarriage.Size = New System.Drawing.Size(65, 17)
    Me.chkCarriage.TabIndex = 20
    Me.chkCarriage.Text = "Carriage"
    Me.chkCarriage.UseVisualStyleBackColor = True
    '
    'chkConfirmCarriage
    '
    Me.chkConfirmCarriage.AutoSize = True
    Me.chkConfirmCarriage.Location = New System.Drawing.Point(13, 133)
    Me.chkConfirmCarriage.Name = "chkConfirmCarriage"
    Me.chkConfirmCarriage.Size = New System.Drawing.Size(103, 17)
    Me.chkConfirmCarriage.TabIndex = 21
    Me.chkConfirmCarriage.Text = "Confirm Carriage"
    Me.chkConfirmCarriage.UseVisualStyleBackColor = True
    '
    'chkBypass
    '
    Me.chkBypass.AutoSize = True
    Me.chkBypass.Location = New System.Drawing.Point(13, 170)
    Me.chkBypass.Name = "chkBypass"
    Me.chkBypass.Size = New System.Drawing.Size(117, 17)
    Me.chkBypass.TabIndex = 22
    Me.chkBypass.Text = "Bypass Paragraphs"
    Me.chkBypass.UseVisualStyleBackColor = True
    '
    'chkSelectBatch
    '
    Me.chkSelectBatch.AutoSize = True
    Me.chkSelectBatch.Location = New System.Drawing.Point(13, 16)
    Me.chkSelectBatch.Name = "chkSelectBatch"
    Me.chkSelectBatch.Size = New System.Drawing.Size(87, 17)
    Me.chkSelectBatch.TabIndex = 18
    Me.chkSelectBatch.Text = "Select Batch"
    Me.chkSelectBatch.UseVisualStyleBackColor = True
    '
    'chkForeignCurrency
    '
    Me.chkForeignCurrency.AutoSize = True
    Me.chkForeignCurrency.Location = New System.Drawing.Point(449, 92)
    Me.chkForeignCurrency.Name = "chkForeignCurrency"
    Me.chkForeignCurrency.Size = New System.Drawing.Size(106, 17)
    Me.chkForeignCurrency.TabIndex = 33
    Me.chkForeignCurrency.Text = "Foreign Currency"
    Me.chkForeignCurrency.UseVisualStyleBackColor = True
    Me.chkForeignCurrency.Visible = False
    '
    'tbpMethods
    '
    Me.tbpMethods.Controls.Add(Me.pnlPaymentMethods)
    Me.tbpMethods.Location = New System.Drawing.Point(4, 25)
    Me.tbpMethods.Name = "tbpMethods"
    Me.tbpMethods.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpMethods.Size = New System.Drawing.Size(644, 269)
    Me.tbpMethods.TabIndex = 1
    Me.tbpMethods.Text = "Payment Methods"
    Me.tbpMethods.UseVisualStyleBackColor = True
    '
    'pnlPaymentMethods
    '
    Me.pnlPaymentMethods.BackColor = System.Drawing.Color.Transparent
    Me.pnlPaymentMethods.Controls.Add(Me.chkCCWithInvoice)
    Me.pnlPaymentMethods.Controls.Add(Me.chkChequeWithInvoice)
    Me.pnlPaymentMethods.Controls.Add(Me.chkDebitCard)
    Me.pnlPaymentMethods.Controls.Add(Me.chkGiftInKind)
    Me.pnlPaymentMethods.Controls.Add(Me.chkPostalOrder)
    Me.pnlPaymentMethods.Controls.Add(Me.chkSaleOrReturn)
    Me.pnlPaymentMethods.Controls.Add(Me.chkCreditCard)
    Me.pnlPaymentMethods.Controls.Add(Me.chkCAFCard)
    Me.pnlPaymentMethods.Controls.Add(Me.chkCheque)
    Me.pnlPaymentMethods.Controls.Add(Me.chkPaymentPlan)
    Me.pnlPaymentMethods.Controls.Add(Me.chkCreditSale)
    Me.pnlPaymentMethods.Controls.Add(Me.chkVoucher)
    Me.pnlPaymentMethods.Controls.Add(Me.chkCash)
    Me.pnlPaymentMethods.Location = New System.Drawing.Point(0, 0)
    Me.pnlPaymentMethods.Name = "pnlPaymentMethods"
    Me.pnlPaymentMethods.Size = New System.Drawing.Size(644, 269)
    Me.pnlPaymentMethods.TabIndex = 0
    '
    'chkCCWithInvoice
    '
    Me.chkCCWithInvoice.AutoSize = True
    Me.chkCCWithInvoice.Location = New System.Drawing.Point(170, 100)
    Me.chkCCWithInvoice.Name = "chkCCWithInvoice"
    Me.chkCCWithInvoice.Size = New System.Drawing.Size(100, 17)
    Me.chkCCWithInvoice.TabIndex = 22
    Me.chkCCWithInvoice.Text = "CC with Invoice"
    Me.chkCCWithInvoice.UseVisualStyleBackColor = True
    '
    'chkChequeWithInvoice
    '
    Me.chkChequeWithInvoice.AutoSize = True
    Me.chkChequeWithInvoice.Location = New System.Drawing.Point(18, 100)
    Me.chkChequeWithInvoice.Name = "chkChequeWithInvoice"
    Me.chkChequeWithInvoice.Size = New System.Drawing.Size(123, 17)
    Me.chkChequeWithInvoice.TabIndex = 18
    Me.chkChequeWithInvoice.Text = "Cheque with Invoice"
    Me.chkChequeWithInvoice.UseVisualStyleBackColor = True
    '
    'chkDebitCard
    '
    Me.chkDebitCard.AutoSize = True
    Me.chkDebitCard.Location = New System.Drawing.Point(170, 140)
    Me.chkDebitCard.Name = "chkDebitCard"
    Me.chkDebitCard.Size = New System.Drawing.Size(76, 17)
    Me.chkDebitCard.TabIndex = 23
    Me.chkDebitCard.Text = "Debit Card"
    Me.chkDebitCard.UseVisualStyleBackColor = True
    '
    'chkGiftInKind
    '
    Me.chkGiftInKind.AutoSize = True
    Me.chkGiftInKind.Location = New System.Drawing.Point(347, 100)
    Me.chkGiftInKind.Name = "chkGiftInKind"
    Me.chkGiftInKind.Size = New System.Drawing.Size(78, 17)
    Me.chkGiftInKind.TabIndex = 26
    Me.chkGiftInKind.Text = "Gift In Kind"
    Me.chkGiftInKind.UseVisualStyleBackColor = True
    '
    'chkPostalOrder
    '
    Me.chkPostalOrder.AutoSize = True
    Me.chkPostalOrder.Location = New System.Drawing.Point(18, 140)
    Me.chkPostalOrder.Name = "chkPostalOrder"
    Me.chkPostalOrder.Size = New System.Drawing.Size(84, 17)
    Me.chkPostalOrder.TabIndex = 19
    Me.chkPostalOrder.Text = "Postal Order"
    Me.chkPostalOrder.UseVisualStyleBackColor = True
    '
    'chkSaleOrReturn
    '
    Me.chkSaleOrReturn.AutoSize = True
    Me.chkSaleOrReturn.Location = New System.Drawing.Point(507, 60)
    Me.chkSaleOrReturn.Name = "chkSaleOrReturn"
    Me.chkSaleOrReturn.Size = New System.Drawing.Size(96, 17)
    Me.chkSaleOrReturn.TabIndex = 28
    Me.chkSaleOrReturn.Text = "Sale Or Return"
    Me.chkSaleOrReturn.UseVisualStyleBackColor = True
    '
    'chkCreditCard
    '
    Me.chkCreditCard.AutoSize = True
    Me.chkCreditCard.Location = New System.Drawing.Point(170, 60)
    Me.chkCreditCard.Name = "chkCreditCard"
    Me.chkCreditCard.Size = New System.Drawing.Size(78, 17)
    Me.chkCreditCard.TabIndex = 21
    Me.chkCreditCard.Text = "Credit Card"
    Me.chkCreditCard.UseVisualStyleBackColor = True
    '
    'chkCAFCard
    '
    Me.chkCAFCard.AutoSize = True
    Me.chkCAFCard.Location = New System.Drawing.Point(347, 60)
    Me.chkCAFCard.Name = "chkCAFCard"
    Me.chkCAFCard.Size = New System.Drawing.Size(71, 17)
    Me.chkCAFCard.TabIndex = 25
    Me.chkCAFCard.Text = "CAF Card"
    Me.chkCAFCard.UseVisualStyleBackColor = True
    '
    'chkCheque
    '
    Me.chkCheque.AutoSize = True
    Me.chkCheque.Location = New System.Drawing.Point(18, 60)
    Me.chkCheque.Name = "chkCheque"
    Me.chkCheque.Size = New System.Drawing.Size(63, 17)
    Me.chkCheque.TabIndex = 17
    Me.chkCheque.Text = "Cheque"
    Me.chkCheque.UseVisualStyleBackColor = True
    '
    'chkPaymentPlan
    '
    Me.chkPaymentPlan.AutoSize = True
    Me.chkPaymentPlan.Location = New System.Drawing.Point(507, 20)
    Me.chkPaymentPlan.Name = "chkPaymentPlan"
    Me.chkPaymentPlan.Size = New System.Drawing.Size(91, 17)
    Me.chkPaymentPlan.TabIndex = 27
    Me.chkPaymentPlan.Text = "Payment Plan"
    Me.chkPaymentPlan.UseVisualStyleBackColor = True
    '
    'chkCreditSale
    '
    Me.chkCreditSale.AutoSize = True
    Me.chkCreditSale.Location = New System.Drawing.Point(170, 20)
    Me.chkCreditSale.Name = "chkCreditSale"
    Me.chkCreditSale.Size = New System.Drawing.Size(82, 17)
    Me.chkCreditSale.TabIndex = 20
    Me.chkCreditSale.Text = "Credit Sales"
    Me.chkCreditSale.UseVisualStyleBackColor = True
    '
    'chkVoucher
    '
    Me.chkVoucher.AutoSize = True
    Me.chkVoucher.Location = New System.Drawing.Point(347, 20)
    Me.chkVoucher.Name = "chkVoucher"
    Me.chkVoucher.Size = New System.Drawing.Size(66, 17)
    Me.chkVoucher.TabIndex = 24
    Me.chkVoucher.Text = "Voucher"
    Me.chkVoucher.UseVisualStyleBackColor = True
    '
    'chkCash
    '
    Me.chkCash.AutoSize = True
    Me.chkCash.Location = New System.Drawing.Point(18, 20)
    Me.chkCash.Name = "chkCash"
    Me.chkCash.Size = New System.Drawing.Size(50, 17)
    Me.chkCash.TabIndex = 16
    Me.chkCash.Text = "Cash"
    Me.chkCash.UseVisualStyleBackColor = True
    '
    'TabPage2
    '
    Me.TabPage2.Controls.Add(Me.lblAutoGADHelp)
    Me.TabPage2.Controls.Add(Me.newDeclarationGroup)
    Me.TabPage2.Controls.Add(Me.autoGiftAidDeclaration)
    Me.TabPage2.Location = New System.Drawing.Point(4, 25)
    Me.TabPage2.Name = "TabPage2"
    Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage2.Size = New System.Drawing.Size(644, 269)
    Me.TabPage2.TabIndex = 2
    Me.TabPage2.Text = "Gift Aid Declarations"
    Me.TabPage2.UseVisualStyleBackColor = True
    '
    'lblAutoGADHelp
    '
    Me.lblAutoGADHelp.Location = New System.Drawing.Point(314, 21)
    Me.lblAutoGADHelp.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.lblAutoGADHelp.Name = "lblAutoGADHelp"
    Me.lblAutoGADHelp.Size = New System.Drawing.Size(20, 17)
    Me.lblAutoGADHelp.Style = CDBNETCL.InfoLabelStyle.Tooltip
    Me.lblAutoGADHelp.TabIndex = 4
    Me.lblAutoGADHelp.Text = resources.GetString("lblAutoGADHelp.Text")
    '
    'newDeclarationGroup
    '
    Me.newDeclarationGroup.Controls.Add(Me.InfoLabel1)
    Me.newDeclarationGroup.Controls.Add(Me.gadMethodGroup)
    Me.newDeclarationGroup.Controls.Add(Me.gadSource)
    Me.newDeclarationGroup.Controls.Add(Me.gadSourceLabel)
    Me.newDeclarationGroup.Enabled = False
    Me.newDeclarationGroup.Location = New System.Drawing.Point(39, 38)
    Me.newDeclarationGroup.Name = "newDeclarationGroup"
    Me.newDeclarationGroup.Size = New System.Drawing.Size(592, 225)
    Me.newDeclarationGroup.TabIndex = 3
    '
    'InfoLabel1
    '
    Me.InfoLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.InfoLabel1.ForeColor = System.Drawing.SystemColors.GrayText
    Me.InfoLabel1.Location = New System.Drawing.Point(170, 120)
    Me.InfoLabel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.InfoLabel1.Name = "InfoLabel1"
    Me.InfoLabel1.Size = New System.Drawing.Size(413, 35)
    Me.InfoLabel1.Style = CDBNETCL.InfoLabelStyle.Label
    Me.InfoLabel1.TabIndex = 4
    Me.InfoLabel1.Text = "If no Source is specified then the Gift Aid Declaration will be created with the " &
    "Source of the first Batch Transaction Analysis line."
    '
    'gadMethodGroup
    '
    Me.gadMethodGroup.Controls.Add(Me.Label9)
    Me.gadMethodGroup.Controls.Add(Me.gadMethodElectronic)
    Me.gadMethodGroup.Controls.Add(Me.gadMethodWritten)
    Me.gadMethodGroup.Controls.Add(Me.gadMethodOral)
    Me.gadMethodGroup.Location = New System.Drawing.Point(3, 3)
    Me.gadMethodGroup.Name = "gadMethodGroup"
    Me.gadMethodGroup.Size = New System.Drawing.Size(219, 78)
    Me.gadMethodGroup.TabIndex = 3
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(3, 7)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(46, 13)
    Me.Label9.TabIndex = 3
    Me.Label9.Text = "Method:"
    '
    'gadMethodElectronic
    '
    Me.gadMethodElectronic.AutoSize = True
    Me.gadMethodElectronic.Location = New System.Drawing.Point(75, 52)
    Me.gadMethodElectronic.Name = "gadMethodElectronic"
    Me.gadMethodElectronic.Size = New System.Drawing.Size(72, 17)
    Me.gadMethodElectronic.TabIndex = 2
    Me.gadMethodElectronic.TabStop = True
    Me.gadMethodElectronic.Text = "Electronic"
    Me.gadMethodElectronic.UseVisualStyleBackColor = True
    '
    'gadMethodWritten
    '
    Me.gadMethodWritten.AutoSize = True
    Me.gadMethodWritten.Location = New System.Drawing.Point(75, 7)
    Me.gadMethodWritten.Name = "gadMethodWritten"
    Me.gadMethodWritten.Size = New System.Drawing.Size(59, 17)
    Me.gadMethodWritten.TabIndex = 0
    Me.gadMethodWritten.TabStop = True
    Me.gadMethodWritten.Text = "Written"
    Me.gadMethodWritten.UseVisualStyleBackColor = True
    '
    'gadMethodOral
    '
    Me.gadMethodOral.AutoSize = True
    Me.gadMethodOral.Location = New System.Drawing.Point(75, 30)
    Me.gadMethodOral.Name = "gadMethodOral"
    Me.gadMethodOral.Size = New System.Drawing.Size(44, 17)
    Me.gadMethodOral.TabIndex = 1
    Me.gadMethodOral.TabStop = True
    Me.gadMethodOral.Text = "Oral"
    Me.gadMethodOral.UseVisualStyleBackColor = True
    '
    'gadSource
    '
    Me.gadSource.ActiveOnly = False
    Me.gadSource.BackColor = System.Drawing.SystemColors.Control
    Me.gadSource.CustomFormNumber = 0
    Me.gadSource.Description = ""
    Me.gadSource.EnabledProperty = True
    Me.gadSource.ExamCentreId = 0
    Me.gadSource.ExamCentreUnitId = 0
    Me.gadSource.ExamUnitLinkId = 0
    Me.gadSource.HasDependancies = False
    Me.gadSource.IsDesign = False
    Me.gadSource.Location = New System.Drawing.Point(78, 90)
    Me.gadSource.MaxLength = 3
    Me.gadSource.MultipleValuesSupported = False
    Me.gadSource.Name = "gadSource"
    Me.gadSource.OriginalText = Nothing
    Me.gadSource.PreventHistoricalSelection = False
    Me.gadSource.ReadOnlyProperty = False
    Me.gadSource.Size = New System.Drawing.Size(478, 24)
    Me.gadSource.TabIndex = 2
    Me.gadSource.TextReadOnly = False
    Me.gadSource.TotalWidth = 408
    Me.gadSource.ValidationRequired = True
    Me.gadSource.WarningMessage = Nothing
    '
    'gadSourceLabel
    '
    Me.gadSourceLabel.AutoSize = True
    Me.gadSourceLabel.Location = New System.Drawing.Point(10, 90)
    Me.gadSourceLabel.Name = "gadSourceLabel"
    Me.gadSourceLabel.Size = New System.Drawing.Size(44, 13)
    Me.gadSourceLabel.TabIndex = 1
    Me.gadSourceLabel.Text = "Source:"
    '
    'autoGiftAidDeclaration
    '
    Me.autoGiftAidDeclaration.AutoSize = True
    Me.autoGiftAidDeclaration.Location = New System.Drawing.Point(26, 17)
    Me.autoGiftAidDeclaration.Name = "autoGiftAidDeclaration"
    Me.autoGiftAidDeclaration.Size = New System.Drawing.Size(221, 17)
    Me.autoGiftAidDeclaration.TabIndex = 0
    Me.autoGiftAidDeclaration.Text = "Automatically Create Gift Aid Declarations"
    Me.autoGiftAidDeclaration.UseVisualStyleBackColor = True
    '
    'TabPage3
    '
    Me.TabPage3.Controls.Add(Me.Panel1)
    Me.TabPage3.Controls.Add(Me.chkRequireAuthorisation)
    Me.TabPage3.Controls.Add(Me.chkOnlineAuth)
    Me.TabPage3.Location = New System.Drawing.Point(4, 25)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage3.Size = New System.Drawing.Size(644, 269)
    Me.TabPage3.TabIndex = 3
    Me.TabPage3.Text = "Credit Cards"
    Me.TabPage3.UseVisualStyleBackColor = True
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.Label10)
    Me.Panel1.Controls.Add(Me.InfoLabel2)
    Me.Panel1.Controls.Add(Me.txtMerchantRetailNumber)
    Me.Panel1.Location = New System.Drawing.Point(6, 63)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(632, 197)
    Me.Panel1.TabIndex = 38
    '
    'Label10
    '
    Me.Label10.AutoSize = True
    Me.Label10.Location = New System.Drawing.Point(16, 26)
    Me.Label10.Name = "Label10"
    Me.Label10.Size = New System.Drawing.Size(128, 13)
    Me.Label10.TabIndex = 14
    Me.Label10.Text = "Merchant Retail Number :"
    '
    'InfoLabel2
    '
    Me.InfoLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.InfoLabel2.ForeColor = System.Drawing.SystemColors.GrayText
    Me.InfoLabel2.Location = New System.Drawing.Point(19, 55)
    Me.InfoLabel2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.InfoLabel2.Name = "InfoLabel2"
    Me.InfoLabel2.Size = New System.Drawing.Size(553, 117)
    Me.InfoLabel2.Style = CDBNETCL.InfoLabelStyle.Label
    Me.InfoLabel2.TabIndex = 5
    Me.InfoLabel2.Text = resources.GetString("InfoLabel2.Text")
    '
    'txtMerchantRetailNumber
    '
    Me.txtMerchantRetailNumber.ActiveOnly = False
    Me.txtMerchantRetailNumber.BackColor = System.Drawing.SystemColors.Control
    Me.txtMerchantRetailNumber.CustomFormNumber = 0
    Me.txtMerchantRetailNumber.Description = ""
    Me.txtMerchantRetailNumber.EnabledProperty = True
    Me.txtMerchantRetailNumber.ExamCentreId = 0
    Me.txtMerchantRetailNumber.ExamCentreUnitId = 0
    Me.txtMerchantRetailNumber.ExamUnitLinkId = 0
    Me.txtMerchantRetailNumber.HasDependancies = False
    Me.txtMerchantRetailNumber.IsDesign = False
    Me.txtMerchantRetailNumber.Location = New System.Drawing.Point(164, 26)
    Me.txtMerchantRetailNumber.MaxLength = 32767
    Me.txtMerchantRetailNumber.MultipleValuesSupported = False
    Me.txtMerchantRetailNumber.Name = "txtMerchantRetailNumber"
    Me.txtMerchantRetailNumber.OriginalText = Nothing
    Me.txtMerchantRetailNumber.PreventHistoricalSelection = False
    Me.txtMerchantRetailNumber.ReadOnlyProperty = False
    Me.txtMerchantRetailNumber.Size = New System.Drawing.Size(408, 24)
    Me.txtMerchantRetailNumber.TabIndex = 0
    Me.txtMerchantRetailNumber.TextReadOnly = False
    Me.txtMerchantRetailNumber.TotalWidth = 408
    Me.txtMerchantRetailNumber.ValidationRequired = True
    Me.txtMerchantRetailNumber.WarningMessage = Nothing
    '
    'chkRequireAuthorisation
    '
    Me.chkRequireAuthorisation.AutoSize = True
    Me.chkRequireAuthorisation.Enabled = False
    Me.chkRequireAuthorisation.Location = New System.Drawing.Point(6, 40)
    Me.chkRequireAuthorisation.Name = "chkRequireAuthorisation"
    Me.chkRequireAuthorisation.Size = New System.Drawing.Size(127, 17)
    Me.chkRequireAuthorisation.TabIndex = 37
    Me.chkRequireAuthorisation.Text = "Require Authorisation"
    Me.chkRequireAuthorisation.UseVisualStyleBackColor = True
    '
    'chkOnlineAuth
    '
    Me.chkOnlineAuth.AutoSize = True
    Me.chkOnlineAuth.Location = New System.Drawing.Point(6, 17)
    Me.chkOnlineAuth.Name = "chkOnlineAuth"
    Me.chkOnlineAuth.Size = New System.Drawing.Size(144, 17)
    Me.chkOnlineAuth.TabIndex = 23
    Me.chkOnlineAuth.Text = "On-Line CC Authorisation"
    Me.chkOnlineAuth.UseVisualStyleBackColor = True
    '
    'TabPage4
    '
    Me.TabPage4.Controls.Add(Me.pnlAlerts)
    Me.TabPage4.Location = New System.Drawing.Point(4, 25)
    Me.TabPage4.Name = "TabPage4"
    Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage4.Size = New System.Drawing.Size(644, 269)
    Me.TabPage4.TabIndex = 4
    Me.TabPage4.Text = "Alerts"
    Me.TabPage4.UseVisualStyleBackColor = True
    '
    'pnlAlerts
    '
    Me.pnlAlerts.Controls.Add(Me.pnlAlertsGrid)
    Me.pnlAlerts.Controls.Add(Me.ilAlerts)
    Me.pnlAlerts.Controls.Add(Me.chkContactAlerts)
    Me.pnlAlerts.Controls.Add(Me.bplAlerts)
    Me.pnlAlerts.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlAlerts.Location = New System.Drawing.Point(3, 3)
    Me.pnlAlerts.Name = "pnlAlerts"
    Me.pnlAlerts.Size = New System.Drawing.Size(638, 263)
    Me.pnlAlerts.TabIndex = 0
    '
    'pnlAlertsGrid
    '
    Me.pnlAlertsGrid.Controls.Add(Me.dgrAlerts)
    Me.pnlAlertsGrid.Location = New System.Drawing.Point(3, 100)
    Me.pnlAlertsGrid.Name = "pnlAlertsGrid"
    Me.pnlAlertsGrid.Size = New System.Drawing.Size(632, 126)
    Me.pnlAlertsGrid.TabIndex = 39
    '
    'dgrAlerts
    '
    Me.dgrAlerts.AccessibleName = "Display Grid"
    Me.dgrAlerts.ActiveColumn = 0
    Me.dgrAlerts.AllowColumnResize = True
    Me.dgrAlerts.AllowSorting = True
    Me.dgrAlerts.AutoSetHeight = True
    Me.dgrAlerts.AutoSetRowHeight = False
    Me.dgrAlerts.DisplayTitle = Nothing
    Me.dgrAlerts.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrAlerts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrAlerts.Location = New System.Drawing.Point(0, 0)
    Me.dgrAlerts.MaintenanceDesc = Nothing
    Me.dgrAlerts.MaxGridRows = 8
    Me.dgrAlerts.MultipleSelect = False
    Me.dgrAlerts.Name = "dgrAlerts"
    Me.dgrAlerts.RowCount = 10
    Me.dgrAlerts.ShowIfEmpty = True
    Me.dgrAlerts.Size = New System.Drawing.Size(632, 126)
    Me.dgrAlerts.SuppressHyperLinkFormat = False
    Me.dgrAlerts.TabIndex = 1
    '
    'ilAlerts
    '
    Me.ilAlerts.Location = New System.Drawing.Point(10, 40)
    Me.ilAlerts.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.ilAlerts.Name = "ilAlerts"
    Me.ilAlerts.Size = New System.Drawing.Size(608, 55)
    Me.ilAlerts.Style = CDBNETCL.InfoLabelStyle.Label
    Me.ilAlerts.TabIndex = 38
    Me.ilAlerts.Text = "Generic Contact Alerts and Sticky Notes will only be run when the option is ticke" &
    "d." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "The Alerts shown in the grid below are specific to this Trader Application" &
    " and will always be run."
    '
    'chkContactAlerts
    '
    Me.chkContactAlerts.AutoSize = True
    Me.chkContactAlerts.Location = New System.Drawing.Point(13, 16)
    Me.chkContactAlerts.Name = "chkContactAlerts"
    Me.chkContactAlerts.Size = New System.Drawing.Size(246, 17)
    Me.chkContactAlerts.TabIndex = 37
    Me.chkContactAlerts.Text = "Show Generic Contact Alerts and Sticky Notes"
    Me.chkContactAlerts.UseVisualStyleBackColor = True
    '
    'bplAlerts
    '
    Me.bplAlerts.Controls.Add(Me.cmdAddAlert)
    Me.bplAlerts.Controls.Add(Me.cmdAddAlertLink)
    Me.bplAlerts.Controls.Add(Me.cmdDeleteAlertLink)
    Me.bplAlerts.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bplAlerts.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bplAlerts.Location = New System.Drawing.Point(0, 224)
    Me.bplAlerts.Name = "bplAlerts"
    Me.bplAlerts.Size = New System.Drawing.Size(638, 39)
    Me.bplAlerts.TabIndex = 1
    '
    'cmdAddAlert
    '
    Me.cmdAddAlert.Location = New System.Drawing.Point(160, 6)
    Me.cmdAddAlert.Name = "cmdAddAlert"
    Me.cmdAddAlert.Size = New System.Drawing.Size(96, 27)
    Me.cmdAddAlert.TabIndex = 3
    Me.cmdAddAlert.Text = "Add Alert"
    Me.cmdAddAlert.UseVisualStyleBackColor = True
    '
    'cmdAddAlertLink
    '
    Me.cmdAddAlertLink.Location = New System.Drawing.Point(271, 6)
    Me.cmdAddAlertLink.Name = "cmdAddAlertLink"
    Me.cmdAddAlertLink.Size = New System.Drawing.Size(96, 27)
    Me.cmdAddAlertLink.TabIndex = 2
    Me.cmdAddAlertLink.Text = "Add Link"
    Me.cmdAddAlertLink.UseVisualStyleBackColor = True
    '
    'cmdDeleteAlertLink
    '
    Me.cmdDeleteAlertLink.Location = New System.Drawing.Point(382, 6)
    Me.cmdDeleteAlertLink.Name = "cmdDeleteAlertLink"
    Me.cmdDeleteAlertLink.Size = New System.Drawing.Size(96, 27)
    Me.cmdDeleteAlertLink.TabIndex = 1
    Me.cmdDeleteAlertLink.Text = "Delete Link"
    Me.cmdDeleteAlertLink.UseVisualStyleBackColor = True
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdCopy)
    Me.bpl.Controls.Add(Me.cmdRevert)
    Me.bpl.Controls.Add(Me.cmdDesign)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdOk)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 445)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(700, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdCopy
    '
    Me.cmdCopy.Location = New System.Drawing.Point(2, 6)
    Me.cmdCopy.Name = "cmdCopy"
    Me.cmdCopy.Size = New System.Drawing.Size(96, 27)
    Me.cmdCopy.TabIndex = 0
    Me.cmdCopy.Text = "&Copy"
    Me.cmdCopy.UseVisualStyleBackColor = True
    '
    'cmdRevert
    '
    Me.cmdRevert.Location = New System.Drawing.Point(102, 6)
    Me.cmdRevert.Name = "cmdRevert"
    Me.cmdRevert.Size = New System.Drawing.Size(96, 27)
    Me.cmdRevert.TabIndex = 1
    Me.cmdRevert.Text = "&Revert"
    Me.cmdRevert.UseVisualStyleBackColor = True
    '
    'cmdDesign
    '
    Me.cmdDesign.Location = New System.Drawing.Point(202, 6)
    Me.cmdDesign.Name = "cmdDesign"
    Me.cmdDesign.Size = New System.Drawing.Size(96, 27)
    Me.cmdDesign.TabIndex = 2
    Me.cmdDesign.Text = "&Design"
    Me.cmdDesign.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(302, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 3
    Me.cmdNew.Text = "&New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(402, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 4
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(502, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 5
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(602, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 6
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmFPApplication
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    Me.ClientSize = New System.Drawing.Size(700, 484)
    Me.ControlBox = False
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.tabMain)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmFPApplication"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Trader Application"
    Me.tbpCurrency.ResumeLayout(False)
    Me.pnlCurrencyBAs.ResumeLayout(False)
    Me.pnlCurrencyBAs.PerformLayout()
    Me.tbpRestrictions.ResumeLayout(False)
    Me.pnlRestrictions.ResumeLayout(False)
    Me.pnlRestrictions.PerformLayout()
    Me.tbpDocuments.ResumeLayout(False)
    Me.pnlDocuments.ResumeLayout(False)
    Me.pnlDocuments.PerformLayout()
    Me.tbpBank.ResumeLayout(False)
    Me.pnlBankAccount.ResumeLayout(False)
    Me.pnlBankAccount.PerformLayout()
    Me.tbpDefaults.ResumeLayout(False)
    Me.tbpCarriage.ResumeLayout(False)
    Me.tbpBatches.ResumeLayout(False)
    Me.pnlBatches.ResumeLayout(False)
    Me.pnlBatches.PerformLayout()
    Me.tbpTransactions.ResumeLayout(False)
    Me.pnlTransactions.ResumeLayout(False)
    Me.pnlTransactions.PerformLayout()
    Me.tbpAnalysisDefaults.ResumeLayout(False)
    Me.pnlAnalysisSub.ResumeLayout(False)
    Me.pnlAnalysisSub.PerformLayout()
    Me.TabPage1.ResumeLayout(False)
    Me.pnlCarriage.ResumeLayout(False)
    Me.pnlCarriage.PerformLayout()
    Me.tbpMembers.ResumeLayout(False)
    Me.pnlMembers.ResumeLayout(False)
    Me.pnlMembers.PerformLayout()
    Me.tbpExams.ResumeLayout(False)
    Me.pnlExams.ResumeLayout(False)
    Me.pnlExams.PerformLayout()
    Me.tbpAnalysis.ResumeLayout(False)
    Me.pnlAnalysis.ResumeLayout(False)
    Me.TabAnalysis.ResumeLayout(False)
    Me.tbpSales.ResumeLayout(False)
    Me.pnlSales.ResumeLayout(False)
    Me.pnlSales.PerformLayout()
    Me.tbpPaymentPlans.ResumeLayout(False)
    Me.pnlPaymentPlans.ResumeLayout(False)
    Me.pnlPaymentPlans.PerformLayout()
    Me.tbpSalesLedger.ResumeLayout(False)
    Me.pnlSalesLedger.ResumeLayout(False)
    Me.pnlSalesLedger.PerformLayout()
    Me.tbpMaintenance.ResumeLayout(False)
    Me.pnlMaintenance.ResumeLayout(False)
    Me.pnlMaintenance.PerformLayout()
    Me.tbpLegacies.ResumeLayout(False)
    Me.pnlLegacies.ResumeLayout(False)
    Me.pnlLegacies.PerformLayout()
    Me.tabMain.ResumeLayout(False)
    Me.tbpGeneral.ResumeLayout(False)
    Me.pnlGeneral.ResumeLayout(False)
    Me.pnlGeneral.PerformLayout()
    Me.TabGeneral1.ResumeLayout(False)
    Me.tbpOptions.ResumeLayout(False)
    Me.plnOption.ResumeLayout(False)
    Me.plnOption.PerformLayout()
    Me.tbpMethods.ResumeLayout(False)
    Me.pnlPaymentMethods.ResumeLayout(False)
    Me.pnlPaymentMethods.PerformLayout()
    Me.TabPage2.ResumeLayout(False)
    Me.TabPage2.PerformLayout()
    Me.newDeclarationGroup.ResumeLayout(False)
    Me.newDeclarationGroup.PerformLayout()
    Me.gadMethodGroup.ResumeLayout(False)
    Me.gadMethodGroup.PerformLayout()
    Me.TabPage3.ResumeLayout(False)
    Me.TabPage3.PerformLayout()
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.TabPage4.ResumeLayout(False)
    Me.pnlAlerts.ResumeLayout(False)
    Me.pnlAlerts.PerformLayout()
    Me.pnlAlertsGrid.ResumeLayout(False)
    Me.bplAlerts.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents tbpCurrency As System.Windows.Forms.TabPage
  Friend WithEvents tbpRestrictions As System.Windows.Forms.TabPage
  Friend WithEvents tbpDocuments As System.Windows.Forms.TabPage
  Friend WithEvents tbpBank As System.Windows.Forms.TabPage
  Friend WithEvents tbpDefaults As System.Windows.Forms.TabPage
  Friend WithEvents tbpAnalysis As System.Windows.Forms.TabPage
  Friend WithEvents tabMain As CDBNETCL.TabControl
  Friend WithEvents tbpGeneral As System.Windows.Forms.TabPage
  Friend WithEvents TabGeneral1 As System.Windows.Forms.TabControl
  Friend WithEvents tbpOptions As System.Windows.Forms.TabPage
  Friend WithEvents tbpMethods As System.Windows.Forms.TabPage
  Friend WithEvents lblType As System.Windows.Forms.Label
  Friend WithEvents txtType As CDBNETCL.TextLookupBox
  Friend WithEvents txtAppDesc As System.Windows.Forms.TextBox
  Friend WithEvents lblMenu As System.Windows.Forms.Label
  Friend WithEvents txtApplication As System.Windows.Forms.TextBox
  Friend WithEvents lblApplication As System.Windows.Forms.Label
  Friend WithEvents TabAnalysis As System.Windows.Forms.TabControl
  Friend WithEvents tbpSales As System.Windows.Forms.TabPage
  Friend WithEvents tbpPaymentPlans As System.Windows.Forms.TabPage
  Friend WithEvents chkServiceBookingCredit As System.Windows.Forms.CheckBox
  Friend WithEvents chkDonation As System.Windows.Forms.CheckBox
  Friend WithEvents chkEventBooking As System.Windows.Forms.CheckBox
  Friend WithEvents chkExamBooking As System.Windows.Forms.CheckBox
  Friend WithEvents chkAccomodationBooking As System.Windows.Forms.CheckBox
  Friend WithEvents chkServiceBooking As System.Windows.Forms.CheckBox
  Friend WithEvents chkProduct As System.Windows.Forms.CheckBox
  Friend WithEvents tbpSalesLedger As System.Windows.Forms.TabPage
  Friend WithEvents tbpMaintenance As System.Windows.Forms.TabPage
  Friend WithEvents tbpLegacies As System.Windows.Forms.TabPage
  Friend WithEvents chkConfirmSale As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmCollection As System.Windows.Forms.CheckBox
  Friend WithEvents chkDirectDebit As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreditCardAuthority As System.Windows.Forms.CheckBox
  Friend WithEvents chkNoPaymentRequired As System.Windows.Forms.CheckBox
  Friend WithEvents chkDisplayScheduledPayment As System.Windows.Forms.CheckBox
  Friend WithEvents chkStandingOrder As System.Windows.Forms.CheckBox
  Friend WithEvents chkCovenantedSubscription As System.Windows.Forms.CheckBox
  Friend WithEvents chkCovenantedDonation As System.Windows.Forms.CheckBox
  Friend WithEvents chkCOvenantedMembership As System.Windows.Forms.CheckBox
  Friend WithEvents chkSubscription As System.Windows.Forms.CheckBox
  Friend WithEvents chkRegularDonation As System.Windows.Forms.CheckBox
  Friend WithEvents chkPayment As System.Windows.Forms.CheckBox
  Friend WithEvents chkMembershipType As System.Windows.Forms.CheckBox
  Friend WithEvents chkMembership As System.Windows.Forms.CheckBox
  Friend WithEvents chkSundryCreditNote As System.Windows.Forms.CheckBox
  Friend WithEvents chkInvoicePayment As System.Windows.Forms.CheckBox
  Friend WithEvents chkCancelPaymentPlan As System.Windows.Forms.CheckBox
  Friend WithEvents chkPayrollGiving As System.Windows.Forms.CheckBox
  Friend WithEvents chkCancelGiftAidDeclaration As System.Windows.Forms.CheckBox
  Friend WithEvents chkStatus As System.Windows.Forms.CheckBox
  Friend WithEvents chkActivity As System.Windows.Forms.CheckBox
  Friend WithEvents chkSuppression As System.Windows.Forms.CheckBox
  Friend WithEvents chkGiftAidDeclaration As System.Windows.Forms.CheckBox
  Friend WithEvents chkGoneAway As System.Windows.Forms.CheckBox
  Friend WithEvents chkAutoPaymentMaintenance As System.Windows.Forms.CheckBox
  Friend WithEvents txtStatus As CDBNETCL.TextLookupBox
  Friend WithEvents chkAddressMaintenance As System.Windows.Forms.CheckBox
  Friend WithEvents txtCancelGiftAidDeclaration As CDBNETCL.TextLookupBox
  Friend WithEvents txtSuppression As CDBNETCL.TextLookupBox
  Friend WithEvents txtActivity As CDBNETCL.TextLookupBox
  Friend WithEvents chkLegacyReceipt As System.Windows.Forms.CheckBox
  Friend WithEvents tbpCarriage As System.Windows.Forms.TabControl
  Friend WithEvents tbpBatches As System.Windows.Forms.TabPage
  Friend WithEvents tbpTransactions As System.Windows.Forms.TabPage
  Friend WithEvents tbpAnalysisDefaults As System.Windows.Forms.TabPage
  Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
  Friend WithEvents lblBatchAnalysis As System.Windows.Forms.Label
  Friend WithEvents lblBatchCategory As System.Windows.Forms.Label
  Friend WithEvents txtBatchAnalysisCode As CDBNETCL.TextLookupBox
  Friend WithEvents txtBatchCategory As CDBNETCL.TextLookupBox
  Friend WithEvents chkDefaultSourceFromLastMailing As System.Windows.Forms.CheckBox
  Friend WithEvents tbpMembers As System.Windows.Forms.TabPage
  Friend WithEvents lblBranch As System.Windows.Forms.Label
  Friend WithEvents txtBranch As CDBNETCL.TextLookupBox
  Friend WithEvents lblLinkToCommunication As System.Windows.Forms.Label
  Friend WithEvents txtLinkToCommunication As CDBNETCL.TextLookupBox
  Friend WithEvents lblTransactionOrigin As System.Windows.Forms.Label
  Friend WithEvents txtTransactionOrigin As CDBNETCL.TextLookupBox
  Friend WithEvents lblSource As System.Windows.Forms.Label
  Friend WithEvents lblSalesPerson As System.Windows.Forms.Label
  Friend WithEvents txtSalesPerson As CDBNETCL.TextLookupBox
  Friend WithEvents txtSource As CDBNETCL.TextLookupBox
  Friend WithEvents chkLinkAnalysisLines As System.Windows.Forms.CheckBox
  Friend WithEvents chkLinkMALToService As System.Windows.Forms.CheckBox
  Friend WithEvents chkLinkMALToEvent As System.Windows.Forms.CheckBox
  Friend WithEvents lblDonationProduct As System.Windows.Forms.Label
  Friend WithEvents txtDonationProduct As CDBNETCL.TextLookupBox
  Friend WithEvents lblAnotherRate As System.Windows.Forms.Label
  Friend WithEvents txtDonationRate As CDBNETCL.TextLookupBox
  Friend WithEvents lblProduct As System.Windows.Forms.Label
  Friend WithEvents lblRate As System.Windows.Forms.Label
  Friend WithEvents txtRate As CDBNETCL.TextLookupBox
  Friend WithEvents txtProduct As CDBNETCL.TextLookupBox
  Friend WithEvents lblProductCarriage As System.Windows.Forms.Label
  Friend WithEvents lblRateCarriage As System.Windows.Forms.Label
  Friend WithEvents txtCarriageRate As CDBNETCL.TextLookupBox
  Friend WithEvents txtCarriageProduct As CDBNETCL.TextLookupBox
  Friend WithEvents txtPercentage As System.Windows.Forms.TextBox
  Friend WithEvents lblPercentage As System.Windows.Forms.Label
  Friend WithEvents lblCCAAccount As System.Windows.Forms.Label
  Friend WithEvents txtCCAAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblCAFAndVoucherAccount As System.Windows.Forms.Label
  Friend WithEvents txtCAFAndVoucherAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblStandingOrderAccount As System.Windows.Forms.Label
  Friend WithEvents lblDirectDebitAccount As System.Windows.Forms.Label
  Friend WithEvents txtDirectDebitAccount As CDBNETCL.TextLookupBox
  Friend WithEvents txtStandingOrderAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblDebitCardAccount As System.Windows.Forms.Label
  Friend WithEvents txtDebitCardAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblCreditSaleAccount As System.Windows.Forms.Label
  Friend WithEvents txtCreditSaleAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblCashAccount As System.Windows.Forms.Label
  Friend WithEvents lblCreditCardAccount As System.Windows.Forms.Label
  Friend WithEvents txtCreditCardAccount As CDBNETCL.TextLookupBox
  Friend WithEvents txtCashAccount As CDBNETCL.TextLookupBox
  Friend WithEvents lblSalesGroup As System.Windows.Forms.Label
  Friend WithEvents txtSalesGroup As CDBNETCL.TextLookupBox
  Friend WithEvents chkForceDistributionCode As System.Windows.Forms.CheckBox
  Friend WithEvents chkIncludeConfirmedTransaction As System.Windows.Forms.CheckBox
  Friend WithEvents chkIncludeProvisionalTransaction As System.Windows.Forms.CheckBox
  Friend WithEvents chkForceMailingCode As System.Windows.Forms.CheckBox
  Friend WithEvents chkIncludeProvPaymentPlan As System.Windows.Forms.CheckBox
  Friend WithEvents chkMembersOnly As System.Windows.Forms.CheckBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents txtProvCashDoc As CDBNETCL.TextLookupBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents txtPayPlanDoc As CDBNETCL.TextLookupBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents txtCreditStmtDoc As CDBNETCL.TextLookupBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents txtReceiptDoc As CDBNETCL.TextLookupBox
  Friend WithEvents txtInvoiceDoc As CDBNETCL.TextLookupBox
  Friend WithEvents lblAlbacsBankDetails As System.Windows.Forms.Label
  Friend WithEvents cboAlbacsBankDetails As System.Windows.Forms.ComboBox
  Friend WithEvents cmdRemove As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents txtBankAccount As CDBNETCL.TextLookupBox
  Friend WithEvents txtBatchType As CDBNETCL.TextLookupBox
  Friend WithEvents txtCurrency As CDBNETCL.TextLookupBox
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdRevert As System.Windows.Forms.Button
  Friend WithEvents cmdDesign As System.Windows.Forms.Button
  Friend WithEvents cmdCopy As System.Windows.Forms.Button
  Friend WithEvents plnOption As CDBNETCL.PanelEx
  Friend WithEvents chkAnalysisComments As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmDetails As System.Windows.Forms.CheckBox
  Friend WithEvents chkAutoSetAmount As System.Windows.Forms.CheckBox
  Friend WithEvents chkPrefulfilledIncentives As System.Windows.Forms.CheckBox
  Friend WithEvents chkTransactionComment As System.Windows.Forms.CheckBox
  Friend WithEvents chkMaintainPaymentPlan As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmProduct As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmAnalysis As System.Windows.Forms.CheckBox
  Friend WithEvents ChkPaymentMethod As System.Windows.Forms.CheckBox
  Friend WithEvents chkNonFinancialBatch As System.Windows.Forms.CheckBox
  Friend WithEvents chkBankDetails As System.Windows.Forms.CheckBox
  Friend WithEvents chkOnlineAuth As System.Windows.Forms.CheckBox
  Friend WithEvents chkShowReference As System.Windows.Forms.CheckBox
  Friend WithEvents chkCarriage As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmCarriage As System.Windows.Forms.CheckBox
  Friend WithEvents chkBypass As System.Windows.Forms.CheckBox
  Friend WithEvents chkSelectBatch As System.Windows.Forms.CheckBox
  Friend WithEvents chkForeignCurrency As System.Windows.Forms.CheckBox
  Friend WithEvents pnlPaymentMethods As CDBNETCL.PanelEx
  Friend WithEvents chkDebitCard As System.Windows.Forms.CheckBox
  Friend WithEvents chkGiftInKind As System.Windows.Forms.CheckBox
  Friend WithEvents chkPostalOrder As System.Windows.Forms.CheckBox
  Friend WithEvents chkSaleOrReturn As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreditCard As System.Windows.Forms.CheckBox
  Friend WithEvents chkCAFCard As System.Windows.Forms.CheckBox
  Friend WithEvents chkCheque As System.Windows.Forms.CheckBox
  Friend WithEvents chkPaymentPlan As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreditSale As System.Windows.Forms.CheckBox
  Friend WithEvents chkVoucher As System.Windows.Forms.CheckBox
  Friend WithEvents chkCash As System.Windows.Forms.CheckBox
  Friend WithEvents pnlGeneral As CDBNETCL.PanelEx
  Friend WithEvents pnlAnalysis As CDBNETCL.PanelEx
  Friend WithEvents pnlCurrencyBAs As CDBNETCL.PanelEx
  Friend WithEvents pnlRestrictions As CDBNETCL.PanelEx
  Friend WithEvents pnlDocuments As CDBNETCL.PanelEx
  Friend WithEvents pnlBankAccount As CDBNETCL.PanelEx
  Friend WithEvents pnlBatches As CDBNETCL.PanelEx
  Friend WithEvents pnlTransactions As CDBNETCL.PanelEx
  Friend WithEvents pnlAnalysisSub As CDBNETCL.PanelEx
  Friend WithEvents pnlCarriage As CDBNETCL.PanelEx
  Friend WithEvents pnlMembers As CDBNETCL.PanelEx
  Friend WithEvents pnlSales As CDBNETCL.PanelEx
  Friend WithEvents pnlPaymentPlans As CDBNETCL.PanelEx
  Friend WithEvents pnlSalesLedger As CDBNETCL.PanelEx
  Friend WithEvents pnlMaintenance As CDBNETCL.PanelEx
  Friend WithEvents pnlLegacies As CDBNETCL.PanelEx
  Friend WithEvents chkSalesContactMandatory As System.Windows.Forms.CheckBox
  Friend WithEvents chkInvoicePrintPreview As System.Windows.Forms.CheckBox
  Friend WithEvents chkLoan As System.Windows.Forms.CheckBox
  Friend WithEvents chkDateRangeMsgInPrint As System.Windows.Forms.CheckBox
  Friend WithEvents chkUnpostedBatchMsgInPrint As System.Windows.Forms.CheckBox
  Friend WithEvents chkAutoCreateCreditCustomer As System.Windows.Forms.CheckBox
  Friend WithEvents lblCreditCategory As System.Windows.Forms.Label
  Friend WithEvents txtCreditCategory As CDBNETCL.TextLookupBox
  Friend WithEvents chkCCWithInvoice As System.Windows.Forms.CheckBox
  Friend WithEvents chkChequeWithInvoice As System.Windows.Forms.CheckBox
  Friend WithEvents tbpExams As System.Windows.Forms.TabPage
  Friend WithEvents pnlExams As CDBNETCL.PanelEx
  Friend WithEvents lblExamSession As System.Windows.Forms.Label
  Friend WithEvents txtExamSession As CDBNETCL.TextLookupBox
  Friend WithEvents lblExamUnit As System.Windows.Forms.Label
  Friend WithEvents txtExamUnit As CDBNETCL.TextLookupBox
  Friend WithEvents chkInvoicePrintUnpostedBatches As System.Windows.Forms.CheckBox
  Friend WithEvents chkRequireAuthorisation As System.Windows.Forms.CheckBox
  Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
  Friend WithEvents autoGiftAidDeclaration As System.Windows.Forms.CheckBox
  Friend WithEvents gadSource As CDBNETCL.TextLookupBox
  Friend WithEvents gadSourceLabel As System.Windows.Forms.Label
  Friend WithEvents gadMethodElectronic As System.Windows.Forms.RadioButton
  Friend WithEvents gadMethodOral As System.Windows.Forms.RadioButton
  Friend WithEvents gadMethodWritten As System.Windows.Forms.RadioButton
  Friend WithEvents gadMethodGroup As Panel
  Friend WithEvents Label9 As Label
  Friend WithEvents newDeclarationGroup As Panel
  Friend WithEvents InfoLabel1 As InfoLabel
  Friend WithEvents lblAutoGADHelp As InfoLabel
  Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents txtMerchantRetailNumber As CDBNETCL.TextLookupBox
  Friend WithEvents Label10 As System.Windows.Forms.Label
  Friend WithEvents InfoLabel2 As CDBNETCL.InfoLabel
  Friend WithEvents TabPage4 As TabPage
  Friend WithEvents pnlAlerts As Panel
  Friend WithEvents chkContactAlerts As CheckBox
  Friend WithEvents bplAlerts As ButtonPanel
  Friend WithEvents cmdAddAlert As Button
  Friend WithEvents cmdAddAlertLink As Button
  Friend WithEvents cmdDeleteAlertLink As Button
  Friend WithEvents ilAlerts As InfoLabel
  Friend WithEvents pnlAlertsGrid As Panel
  Friend WithEvents dgrAlerts As DisplayGrid
End Class
