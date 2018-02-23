Public Class frmFPApplication

  Private Const CT_TEXTBOX As Integer = 1
  Private Const CT_LABEL As Integer = 2
  Private Const CT_MASKEDBOX As Integer = 3
  Private Const CT_COMBOBOX As Integer = 4
  Private Const CT_CHECKBOX As Integer = 5
  Private Const CT_DTPICKER As Integer = 6
  Private Const CT_MEMO As Integer = 7

  Private mvPFAppNumber As Integer = 0
  Private mvCurrencyCode As Boolean       'Flag that the 'currency_codes' table exists
  Private mvPrefulfilledIncentives As Boolean = True 'Flag that the 'prefulfilled_incentives' attribute exists
  Private mvContactAlerts As Boolean = True 'Flag that the 'contact_alerts' attribute exists
  Private mvScheduledPayments As Boolean = True 'Flag that the 'display_scheduled_payments' attribute exists
  Private mvConfirmSRTransactions As Boolean = True 'Flag that the 'confirm_sr_transactions' attribute exists
  Private mvOnLineCCAuthorisation As Boolean = True 'Flag that the 'online_cc_authorisation' attribute exists
  Private mvPayPlanConvMaintenance As Boolean = True 'Flag that the 'pp_converion_incl_maintenance' attribute exists
  Private mvLinkToCommunication As Boolean = True 'Flag that the 'link_to_communication' attribute exists
  Private mvCollectionPayments As Boolean = True 'Flag that the 'collection_payments' attribute exists
  Private mvBatchAnalysisCode As Boolean = True 'Flag that the 'batch_analysis_code' attribute exists
  Private mvEventMultipleAnalysis As Boolean = True 'Flag that the 'event_multiple_analysis' attribute exists
  Private mvFPTransactionOrigin As Boolean = True 'Flag that the 'transaction_origin' attribute exists
  Private mvServiceBookingAnalysis As Boolean = True 'Flag that the 'service_booking_analysis' attribute exists
  Private mvAlbacsBankDetails As Boolean = True 'Flag that the 'albacs_bank_details' attribute exists
  Private mvFundraisingPayments As Boolean = True 'Flag that the 'fundraising_payment_schedule' table exists
  Private mvInvoicePrintUnpostedBatches As Boolean = True 'Flag that new invoice_print_unposted_batches attribute exists
  Private mvTable As DataTable
  Private mvLoading As Boolean
  Private mvEditMode As Boolean
  Private mvApplicationType As String
  Private mvDataSet As New DataSet
  Private mvReadOnly As Boolean = False
  Private mvPages As New Hashtable
  Friend WithEvents erp As System.Windows.Forms.ErrorProvider
  Private mvFirstErrorControl As Control
  Private mvBankAccountControls() As TextLookupBox

  Public Sub New(ByVal pFPAppNumber As Integer)
    Me.New(pFPAppNumber, False)
  End Sub
  Public Sub New(ByVal pFPAppNumber As Integer, ByVal pForceSave As Boolean)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    mvPFAppNumber = pFPAppNumber
    mvBankAccountControls = {txtCashAccount, txtCreditCardAccount, txtDebitCardAccount, txtCreditSaleAccount, txtStandingOrderAccount, txtDirectDebitAccount, txtCCAAccount, txtCAFAndVoucherAccount}
    InitializeControl(pForceSave)
    ' Add any initialization after the InitializeComponent() call.
  End Sub
  Private Sub InitializeControl(ByVal pForceSave As Boolean)
    Me.erp = New System.Windows.Forms.ErrorProvider(Me.components)
    If mvPFAppNumber > 0 Then
      mvEditMode = True
      If pForceSave Then
        cmdCopy.Enabled = False
        cmdNew.Enabled = False
        cmdCancel.Enabled = False
      End If
    Else
      mvEditMode = False
    End If
    AddCheckHandler(tabMain)
    SetTraderValues()
    Me.Width = Me.Width + 10
    Me.Width = Me.Width - 10
    SetControlTheme()
  End Sub

  Private Sub AddCheckHandler(ByVal pControl As Control)
    Dim vChk As CheckBox
    For Each vCon As Control In pControl.Controls
      If TypeName(vCon) = "CheckBox" Then
        vChk = DirectCast(vCon, CheckBox)
        AddHandler vChk.CheckedChanged, AddressOf CheckChangedHandler
      ElseIf TypeName(vCon) = "TabPage" Or TypeName(vCon) = "TabControl" Or TypeName(vCon) = "PanelEx" Then
        AddCheckHandler(vCon)
      End If
    Next
    AddHandler txtPercentage.KeyPress, AddressOf NumericKeyPressHandler
    AddHandler txtPercentage.Validating, AddressOf NumericReformatHandler
  End Sub


  Private Sub SetTraderValues(Optional ByVal pDisableLockedWarning As Boolean = False)
    Dim vState As Boolean
    Dim vList As New ParameterList(True)
    If mvEditMode Then

      vList("FpApplication") = mvPFAppNumber.ToString()
      mvTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTraderApplication, vList)
      If mvTable IsNot Nothing AndAlso mvTable.Rows.Count > 0 Then
        Dim vRow As DataRow = mvTable.Rows(0)
        mvApplicationType = vRow.Item("FpApplicationType").ToString
        mvLoading = True
        txtApplication.Text = vRow.Item("FpApplication").ToString
        txtAppDesc.Text = vRow.Item("FpApplicationDesc").ToString
        txtType.Text = vRow.Item("FpApplicationType").ToString
        mvCurrencyCode = BooleanValue(vRow.Item("CurrencyCodeExists").ToString)

        chkSelectBatch.Checked = BooleanValue(vRow.Item("BatchApplication").ToString)
        chkCash.Checked = BooleanValue(vRow.Item("PmCash").ToString)
        chkCheque.Checked = BooleanValue(vRow.Item("PmCheque").ToString)
        chkPostalOrder.Checked = BooleanValue(vRow.Item("PmPostalOrder").ToString)
        chkCreditCard.Checked = BooleanValue(vRow.Item("PmCreditCard").ToString)
        chkDebitCard.Checked = BooleanValue(vRow.Item("PmDebitCard").ToString)
        chkCreditSale.Checked = BooleanValue(vRow.Item("PmCredit").ToString)
        chkChequeWithInvoice.Checked = BooleanValue(vRow.Item("ChequeWithInvoice").ToString)
        chkCCWithInvoice.Checked = BooleanValue(vRow.Item("CcWithInvoice").ToString)
        chkBankDetails.Checked = BooleanValue(vRow.Item("BankDetails").ToString)
        chkTransactionComment.Checked = BooleanValue(vRow.Item("TransactionComments").ToString)
        chkMembership.Checked = BooleanValue(vRow.Item("Memberships").ToString)
        chkSubscription.Checked = BooleanValue(vRow.Item("Subscriptions").ToString)
        chkRegularDonation.Checked = BooleanValue(vRow.Item("DonationsRegular").ToString)
        chkProduct.Checked = BooleanValue(vRow.Item("ProductSales").ToString)
        chkDonation.Checked = BooleanValue(vRow.Item("DonationsOneOff").ToString)
        chkPayment.Checked = BooleanValue(vRow.Item("Payments").ToString)
        chkStandingOrder.Checked = BooleanValue(vRow.Item("StandingOrders").ToString)
        chkDirectDebit.Checked = BooleanValue(vRow.Item("DirectDebits").ToString)
        chkCreditCardAuthority.Checked = BooleanValue(vRow.Item("CreditCardAuthorities").ToString)
        chkCOvenantedMembership.Checked = BooleanValue(vRow.Item("CovenantMembership").ToString)
        chkCovenantedSubscription.Checked = BooleanValue(vRow.Item("CovenantSubscription").ToString)
        chkCovenantedDonation.Checked = BooleanValue(vRow.Item("CovenantDonationRegular").ToString)
        chkEventBooking.Checked = BooleanValue(vRow.Item("EventBooking").ToString)
        chkExamBooking.Checked = BooleanValue(vRow.Item("ExamBooking").ToString)
        chkAccomodationBooking.Checked = BooleanValue(vRow.Item("AccomodationBooking").ToString)
        chkInvoicePayment.Checked = BooleanValue(vRow.Item("InvoicePayments").ToString)
        mvReadOnly = (BooleanValue(vRow.Item("ReadOnly").ToString))

        chkShowReference.Checked = BooleanValue(vRow.Item("ShowTransactionReference").ToString)
        chkConfirmProduct.Checked = BooleanValue(vRow.Item("ConfirmDefaultProduct").ToString)
        chkConfirmAnalysis.Checked = BooleanValue(vRow.Item("ConfirmAnalysis").ToString)
        chkCarriage.Checked = BooleanValue(vRow.Item("Carriage").ToString)
        chkConfirmCarriage.Checked = BooleanValue(vRow.Item("ConfirmCarriage").ToString)
        chkAnalysisComments.Checked = BooleanValue(vRow.Item("AnalysisComments").ToString)
        chkNoPaymentRequired.Checked = BooleanValue(vRow.Item("NonPaidPaymentPlans").ToString)
        ChkPaymentMethod.Checked = BooleanValue(vRow.Item("PayMethodsAtEnd").ToString)
        chkServiceBooking.Checked = BooleanValue(vRow.Item("ServiceBookings").ToString)
        chkPaymentPlan.Checked = BooleanValue(vRow.Item("PayPlanPayMethod").ToString)
        chkSundryCreditNote.Checked = BooleanValue(vRow.Item("SundryCreditNotes").ToString)
        chkForeignCurrency.Checked = BooleanValue(vRow.Item("ForeignCurrency").ToString)
        chkMembershipType.Checked = BooleanValue(vRow.Item("ChangeMembership").ToString)
        chkMembersOnly.Checked = BooleanValue(vRow.Item("MembersOnly").ToString)
        chkAutoSetAmount.Checked = BooleanValue(vRow.Item("AutoSetAmount").ToString)
        chkServiceBookingCredit.Checked = BooleanValue(vRow.Item("ServiceBookingCredits").ToString)
        chkLegacyReceipt.Checked = BooleanValue(vRow.Item("LegacyReceipts").ToString)
        chkGoneAway.Checked = BooleanValue(vRow.Item("SetGoneAway").ToString)
        chkStatus.Checked = BooleanValue(vRow.Item("SetStatus").ToString)
        chkGiftAidDeclaration.Checked = BooleanValue(vRow.Item("GiftAidDeclaration").ToString)
        chkActivity.Checked = BooleanValue(vRow.Item("AddActivity").ToString)
        chkSuppression.Checked = BooleanValue(vRow.Item("AddSuppression").ToString)
        chkCancelPaymentPlan.Checked = BooleanValue(vRow.Item("CancelPaymentPlan").ToString)
        chkConfirmDetails.Checked = BooleanValue(vRow.Item("ConfirmDetails").ToString)
        chkIncludeProvisionalTransaction.Checked = BooleanValue(vRow.Item("IncludeProvisionalTrans").ToString)
        chkIncludeConfirmedTransaction.Checked = BooleanValue(vRow.Item("IncludeConfirmedTrans").ToString)
        chkVoucher.Checked = BooleanValue(vRow.Item("PmVoucher").ToString)
        chkCAFCard.Checked = BooleanValue(vRow.Item("PmCafCard").ToString)
        chkGiftInKind.Checked = BooleanValue(vRow.Item("PmGiftInkind").ToString)
        chkPayrollGiving.Checked = BooleanValue(vRow.Item("PayrollGiving").ToString)
        chkDefaultSourceFromLastMailing.Checked = BooleanValue(vRow.Item("SourceFromLastMailing").ToString)
        chkForceMailingCode.Checked = BooleanValue(vRow.Item("MailingCodeMandatory").ToString)
        chkForceDistributionCode.Checked = BooleanValue(vRow.Item("DistributionCodeMandatory").ToString)
        chkSalesContactMandatory.Checked = BooleanValue(vRow.Item("SalesContactMandatory").ToString)
        chkBypass.Checked = BooleanValue(vRow.Item("BypassMailingParagraphs").ToString)
        chkSaleOrReturn.Checked = BooleanValue(vRow.Item("PmSaleOrReturn").ToString)
        chkNonFinancialBatch.Checked = BooleanValue(vRow.Item("NonFinancialBatch").ToString)
        chkAddressMaintenance.Checked = BooleanValue(vRow.Item("AddressMaintenance").ToString)
        chkAutoPaymentMaintenance.Checked = BooleanValue(vRow.Item("AutoPaymentMaintenance").ToString)
        chkCancelGiftAidDeclaration.Checked = BooleanValue(vRow.Item("GiftAidCancellation").ToString)
        chkIncludeProvPaymentPlan.Checked = BooleanValue(vRow.Item("ProvisionalPaymentPlan").ToString)
        chkPrefulfilledIncentives.Checked = BooleanValue(vRow.Item("PrefulfilledIncentives").ToString)
        chkContactAlerts.Checked = BooleanValue(vRow.Item("ContactAlerts").ToString)
        chkDisplayScheduledPayment.Checked = BooleanValue(vRow.Item("DisplayScheduledPayments").ToString)
        chkSaleOrReturn.Checked = BooleanValue(vRow.Item("ConfirmSrTransactions").ToString)
        chkOnlineAuth.Checked = BooleanValue(vRow.Item("OnlineCcAuthorisation").ToString)
        chkRequireAuthorisation.Checked = (chkOnlineAuth.Checked AndAlso BooleanValue(vRow.Item("RequireCcAuthorisation").ToString))
        chkRequireAuthorisation.Enabled = chkOnlineAuth.Checked
        chkMaintainPaymentPlan.Checked = BooleanValue(vRow.Item("PpConversionInclMaintenance").ToString)
        chkConfirmCollection.Checked = BooleanValue(vRow.Item("CollectionPayments").ToString)
        chkLinkMALToEvent.Checked = BooleanValue(vRow.Item("EventMultipleAnalysis").ToString)
        chkLinkMALToService.Checked = BooleanValue(vRow.Item("ServiceBookingAnalysis").ToString)
        chkLinkAnalysisLines.Checked = BooleanValue(vRow.Item("LinkToFundraisingPayments").ToString)

        If mvTable.Columns.Contains("AutoCreateCreditCustomer") AndAlso vRow.Item("AutoCreateCreditCustomer").ToString = "Y" Then chkAutoCreateCreditCustomer.Checked = True Else chkAutoCreateCreditCustomer.Checked = False
        If mvTable.Columns.Contains("UnpostedBatchMsgInPrint") AndAlso vRow.Item("UnpostedBatchMsgInPrint").ToString = "Y" Then chkUnpostedBatchMsgInPrint.Checked = True Else chkUnpostedBatchMsgInPrint.Checked = False
        If mvTable.Columns.Contains("DateRangeMsgInPrint") AndAlso vRow.Item("DateRangeMsgInPrint").ToString = "Y" Then chkDateRangeMsgInPrint.Checked = True Else chkDateRangeMsgInPrint.Checked = False
        If vRow.Item("InvoicePrintPreviewDefault").ToString.Length > 0 Then
          chkInvoicePrintPreview.Checked = BooleanValue(vRow.Item("InvoicePrintPreviewDefault").ToString)
        Else
          'Control not in database so disable
          chkInvoicePrintPreview.Checked = False
          chkInvoicePrintPreview.Enabled = False
        End If
        If mvTable.Columns.Contains("Loans") AndAlso vRow.Item("Loans").ToString.Length > 0 Then
          chkLoan.Checked = BooleanValue(vRow.Item("Loans").ToString)
        Else
          'Control not in database so disable
          chkLoan.Checked = False
          chkLoan.Enabled = False
        End If
        If mvTable.Columns.Contains("InvoicePrintUnpostedBatches") = True AndAlso vRow.Item("InvoicePrintUnpostedBatches").ToString.Length > 0 Then
          chkInvoicePrintUnpostedBatches.Checked = BooleanValue(vRow.Item("InvoicePrintUnpostedBatches").ToString)
        Else
          'Control not in database so disable
          chkInvoicePrintUnpostedBatches.Checked = False
          chkInvoicePrintUnpostedBatches.Enabled = False
          mvInvoicePrintUnpostedBatches = False
        End If
        SetSourceDescription()
        vState = CBool(IIf(vRow.Item("ReadOnly").ToString() = "N", True, False))
        If Not vState AndAlso Not pDisableLockedWarning Then ShowInformationMessage(InformationMessages.ImApplicationLocked)
        cmdOk.Enabled = vState
        cmdDelete.Enabled = vState
        cmdDesign.Enabled = vState
        cmdRevert.Enabled = vState
      End If
      If chkIncludeProvPaymentPlan.Checked Then
        SetAdditionalPMChks(chkIncludeProvPaymentPlan)
        SetFieldsForProvisionalPayPlans()
      End If
      mvLoading = False
      If chkCarriage.Checked = False Then
        chkConfirmCarriage.Checked = False
        chkConfirmCarriage.Enabled = False
      End If
      EnableTypeChange(False)          'Cannot change type of existing application
      SetAllLookup() ' Setting all lookup boxex
      If mvTable.Columns.Contains("Loans") = False Then chkLoan.Enabled = False

      If mvTable.Columns.Contains("AutoGiftAidDeclaration") Then
        autoGiftAidDeclaration.Checked = BooleanValue(mvTable.Rows(0)("AutoGiftAidDeclaration").ToString)
        gadMethodOral.Checked = mvTable.Rows(0)("AutoGiftAidMethod").ToString.Equals("O", StringComparison.InvariantCultureIgnoreCase)
        gadMethodWritten.Checked = mvTable.Rows(0)("AutoGiftAidMethod").ToString.Equals("W", StringComparison.InvariantCultureIgnoreCase)
        gadMethodElectronic.Checked = mvTable.Rows(0)("AutoGiftAidMethod").ToString.Equals("E", StringComparison.InvariantCultureIgnoreCase)
        gadSource.Text = mvTable.Rows(0)("AutoGiftAidSource").ToString
      Else
        Me.autoGiftAidDeclaration.Enabled = False
      End If
      If mvTable.Columns.Contains("MerchantRetailNumber") Then
        txtMerchantRetailNumber.Text = mvTable.Rows(0)("MerchantRetailNumber").ToString
      End If
    Else
      cmdNew.Enabled = False
      cmdDelete.Enabled = False
      cmdRevert.Enabled = False
      cmdCopy.Enabled = False
      cmdAddAlert.Enabled = False
      cmdAddAlertLink.Enabled = False
      cmdDeleteAlertLink.Enabled = False
      SetAllLookup() ' Setting all lookup boxex
      SetDefaults()
      EnableTypeChange(True)          'Can change type of new application
    End If

    If AppValues.ConfigurationValue(AppValues.ConfigurationValues.option_covenants) <> "Yes" Then
      chkCovenantedDonation.Enabled = False
      chkCOvenantedMembership.Enabled = False
      chkCovenantedSubscription.Enabled = False

      chkCovenantedDonation.Checked = False
      chkCOvenantedMembership.Checked = False
      chkCovenantedSubscription.Checked = False
    End If

    SetDefaultFieldsState(chkStatus)
    SetDefaultFieldsState(chkCancelGiftAidDeclaration)
    SetDefaultFieldsState(chkActivity)
    SetDefaultFieldsState(chkSuppression)
    SetDefaultFieldsState(chkCancelPaymentPlan)

    If chkVoucher.Checked Then SetAdditionalPMChks(chkVoucher)
    If chkCAFCard.Checked Then SetAdditionalPMChks(chkCAFCard)
    If chkGiftInKind.Checked Then SetAdditionalPMChks(chkGiftInKind)
    If chkSaleOrReturn.Checked Then SetAdditionalPMChks(chkSaleOrReturn)
    If (mvApplicationType = "GAYEP" Or mvApplicationType = "POTPG") Then
      DisableBankAccountFields(txtStandingOrderAccount)
    End If

    If mvApplicationType <> "BINVG" Then
      chkInvoicePrintPreview.Checked = False
      chkInvoicePrintPreview.Enabled = False
      chkInvoicePrintUnpostedBatches.Checked = False
      chkInvoicePrintUnpostedBatches.Enabled = False
    End If

    If mvEditMode = True Then
      'Provisional cash so enable confirmed & provisional checkboxes
      If chkCash.Checked = True And chkIncludeProvisionalTransaction.Checked = True And chkCAFCard.Checked = False Then
        chkIncludeConfirmedTransaction.Enabled = True
        chkIncludeProvisionalTransaction.Enabled = True
      End If
      If chkCreditCard.Checked = False AndAlso chkDebitCard.Checked = False AndAlso chkCCWithInvoice.Checked = False Then
        chkOnlineAuth.Enabled = False
        chkOnlineAuth.Checked = False
      End If
      If (chkBankDetails.Checked = False Or mvApplicationType = "CLREC") And mvAlbacsBankDetails Then
        cboAlbacsBankDetails.SelectedIndex = 3
        cboAlbacsBankDetails.Enabled = False
      End If
    End If

    'Get linked Contact Alerts
    GetLinkedAlerts()

    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
  End Sub
  Private Sub SetAllLookup()
    InitTextLookupBox(txtCashAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtCreditCardAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtDebitCardAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtCreditSaleAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtDirectDebitAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtStandingOrderAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtCCAAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtCAFAndVoucherAccount, "bank_accounts", "bank_account")
    InitTextLookupBox(txtProduct, "products", "product")
    InitTextLookupBox(txtRate, "rates", "rate")
    InitTextLookupBox(txtDonationProduct, "products", "product")
    InitTextLookupBox(txtDonationRate, "rates", "rate")
    InitTextLookupBox(txtBatchCategory, "batch_categories", "batch_category")
    InitTextLookupBox(txtBatchAnalysisCode, "batch_analysis_codes", "batch_analysis_code")
    InitTextLookupBox(txtType, "fp_application_types", "fp_application_type")
    InitTextLookupBox(txtCarriageProduct, "products", "product")
    InitTextLookupBox(txtCarriageRate, "rates", "rate")
    InitTextLookupBox(txtBranch, "branches", "branch")
    InitTextLookupBox(txtSalesPerson, "sales_persons", "contact_number")
    InitTextLookupBox(txtInvoiceDoc, "standard_documents", "standard_document")
    InitTextLookupBox(txtReceiptDoc, "standard_documents", "standard_document")
    InitTextLookupBox(txtPayPlanDoc, "standard_documents", "standard_document")
    InitTextLookupBox(txtSalesGroup, "sales_groups", "sales_group")
    InitTextLookupBox(txtCreditStmtDoc, "standard_documents", "standard_document")
    InitTextLookupBox(txtProvCashDoc, "standard_documents", "standard_document")
    InitTextLookupBox(txtTransactionOrigin, "transaction_origins", "transaction_origin")
    InitTextLookupBox(txtExamSession, "exam_sessions", "exam_session_code")
    InitTextLookupBox(txtExamUnit, "exam_units", "exam_unit_code")
    cboAlbacsBankDetails.Items.Add(New LookupItem("W", "Warn"))
    cboAlbacsBankDetails.Items.Add(New LookupItem("E", "Error"))
    cboAlbacsBankDetails.Items.Add(New LookupItem("N", "None"))
    cboAlbacsBankDetails.Items.Add(New LookupItem("C", "Use Config"))

    If mvCurrencyCode Then
      InitTextLookupBox(txtBankAccount, "bank_accounts", "bank_account")
      InitTextLookupBox(txtBatchType, "batch_types", "batch_type")
      InitTextLookupBox(txtCurrency, "currency_codes", "currency_code")
      FillGrid()
    Else
      tabMain.TabPages.Remove(tbpCurrency)
    End If
    InitTextLookupBox(txtStatus, "statuses", "status")
    InitTextLookupBox(txtActivity, "activity_groups", "activity_group")
    InitTextLookupBox(txtSuppression, "suppression_groups", "suppression_group")
    InitTextLookupBox(txtCancelGiftAidDeclaration, "cancellation_reasons", "cancellation_reason")
    InitTextLookupBox(txtSource, "sources", "source")
    InitTextLookupBox(txtLinkToCommunication, "fp_applications", "link_to_communication")

    InitTextLookupBox(txtCreditCategory, "credit_categories", "credit_category")

    InitTextLookupBox(gadSource, "sources", "source")
    InitTextLookupBox(txtMerchantRetailNumber, "merchant_details", "merchant_retail_number")

    'vI = ComboSelect(txt(INDEX_CANC_REASON), cbo(INDEX_CANC_REASON), mvCancellationReasonsList)

    '' Setting values
    If mvTable IsNot Nothing AndAlso mvTable.Rows.Count > 0 Then
      Dim vRow As DataRow = mvTable.Rows(0)
      cboAlbacsBankDetails.SelectedIndex = getAlbacsIndex(vRow.Item("AlbacsBankDetails").ToString)
      txtProduct.Text = vRow.Item("Product").ToString
      txtRate.Text = vRow.Item("Rate").ToString
      txtSource.Text = vRow.Item("Source").ToString
      txtCashAccount.Text = vRow.Item("CaBankAccount").ToString
      txtCreditCardAccount.Text = vRow.Item("CcBankAccount").ToString
      txtDebitCardAccount.Text = vRow.Item("DcBankAccount").ToString
      txtCreditSaleAccount.Text = vRow.Item("CsBankAccount").ToString
      txtStandingOrderAccount.Text = vRow.Item("SoBankAccount").ToString
      txtDirectDebitAccount.Text = vRow.Item("DdBankAccount").ToString
      txtCCAAccount.Text = vRow.Item("CcaBankAccount").ToString
      txtCAFAndVoucherAccount.Text = vRow.Item("CVBankAccount").ToString

      txtCarriageProduct.Text = vRow.Item("CarriageProduct").ToString
      txtCarriageRate.Text = vRow.Item("CarriageRate").ToString

      txtReceiptDoc.Text = vRow.Item("ReceiptDocument").ToString
      txtPayPlanDoc.Text = vRow.Item("PaymentPlanDocument").ToString

      txtSalesGroup.Text = vRow.Item("SalesGroup").ToString
      txtCreditStmtDoc.Text = vRow.Item("CreditStatementDocument").ToString

      txtStatus.Text = vRow.Item("Status").ToString 'status

      txtSuppression.Text = vRow.Item("MailingSuppression").ToString
      txtDonationProduct.Text = vRow.Item("DonationProduct").ToString

      txtProvCashDoc.Text = vRow.Item("ProvisionalCashDocument").ToString
      txtTransactionOrigin.Text = vRow.Item("TransactionOrigin").ToString

      txtProduct.Text = vRow.Item("Product").ToString
      txtRate.Text = vRow.Item("Rate").ToString
      txtDonationProduct.Text = vRow.Item("DonationProduct").ToString
      txtDonationRate.Text = vRow.Item("DonationRate").ToString
      txtBatchCategory.Text = vRow.Item("BatchCategory").ToString
      txtBatchAnalysisCode.Text = vRow.Item("BatchAnalysisCode").ToString
      txtType.Text = vRow.Item("FpApplicationType").ToString
      txtCarriageProduct.Text = vRow.Item("CarriageProduct").ToString
      txtCarriageRate.Text = vRow.Item("CarriageRate").ToString
      Dim vSalesContact As Integer = IntegerValue(vRow.Item("DefaultSalesContact").ToString)
      If vSalesContact > 0 Then txtSalesPerson.Text = vSalesContact.ToString
      txtInvoiceDoc.Text = vRow.Item("InvoiceDocument").ToString
      txtReceiptDoc.Text = vRow.Item("ReceiptDocument").ToString
      txtPayPlanDoc.Text = vRow.Item("PaymentPlanDocument").ToString
      txtSalesGroup.Text = vRow.Item("SalesGroup").ToString
      txtCreditStmtDoc.Text = vRow.Item("CreditStatementDocument").ToString
      txtProvCashDoc.Text = vRow.Item("ProvisionalCashDocument").ToString
      txtTransactionOrigin.Text = vRow.Item("TransactionOrigin").ToString
      txtStatus.Text = vRow.Item("Status").ToString
      txtActivity.Text = vRow.Item("ActivityGroup").ToString
      txtSuppression.Text = vRow.Item("MailingSuppression").ToString
      txtCancelGiftAidDeclaration.Text = vRow.Item("CancellationReason").ToString
      txtSource.Text = vRow.Item("source").ToString
      txtBranch.Text = vRow.Item("DefaultMemberBranch").ToString
      txtLinkToCommunication.Text = vRow.Item("LinkToCommunication").ToString
      txtPercentage.Text = vRow.Item("CarriagePercentage").ToString
      If mvTable.Columns.Contains("CreditCategory") Then txtCreditCategory.Text = vRow.Item("CreditCategory").ToString
      txtExamSession.Text = vRow.Item("ExamSessionCode").ToString
      txtExamUnit.Text = vRow.Item("ExamUnitCode").ToString
    End If
  End Sub
  Private Sub GetCodeRestrictionsHandler(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As ParameterList)
    If pList Is Nothing Then pList = New ParameterList(True)
    Select Case DirectCast(sender, TextLookupBox).Name
      Case "txtSource"
        pList("Active") = "Y"
      Case "txtProduct"
        pList("Course") = "N"
        pList("Active") = "Y"
        pList("Accommodation") = "N"
      Case "txtDonationProduct"
        pList("Donation") = "Y"
        pList("Active") = "Y"
      Case "txtCarriageProduct"
        pList("Active") = "Y"
    End Select
  End Sub
  Private Sub GetInitialCodeRestrictionsHandler(ByVal sender As System.Object, ByVal pParameterName As System.String, ByRef pList As CDBNETCL.ParameterList)
    If pList Is Nothing Then pList = New ParameterList(True)

    Select Case DirectCast(sender, TextLookupBox).Name
      Case "txtInvoiceDoc"
        pList("MailmergeHeader") = "INV"
        pList("Active") = "Y"
      Case "txtReceiptDoc"
        pList("MailmergeHeader") = "RECPT"
        pList("Active") = "Y"
      Case "txtPayPlanDoc"
        pList("MailmergeHeader") = "PPLAN"
        pList("Active") = "Y"
      Case "txtCreditStmtDoc"
        pList("MailmergeHeader") = "CSTAT"
        pList("Active") = "Y"
      Case "txtProvCashDoc"
        pList("MailmergeHeader") = "PRVCSH"
        pList("Active") = "Y"
      Case "txtCurrency"
        pList("Restrict") = "Y"
      Case "txtActivity"
        pList("UsageCode") = "R"
      Case "txtType"
        pList("RestrictInAdd") = "Y"
      Case "txtExamSession"
        pList("NonSessionBased") = CBoolYN(True)
    End Select
  End Sub
  Private Sub FillGrid()

    Dim vColumnTable As DataTable = DataHelper.NewColumnTable
    If mvDataSet IsNot Nothing AndAlso mvDataSet.Tables("Column") Is Nothing Then
      DataHelper.AddDataColumn(vColumnTable, "CurrencyCode", "Currency Code")
      DataHelper.AddDataColumn(vColumnTable, "BatchType", "Batch Type")
      DataHelper.AddDataColumn(vColumnTable, "BankAccountDesc", "Bank Account")
      DataHelper.AddDataColumn(vColumnTable, "BankAccount", "Bank Account", "Char", "N")
      mvDataSet.Tables.Add(vColumnTable)
    End If
    Dim vDataTable As DataTable
    Dim vList As New ParameterList(True)
    vList("FpApplicationNumber") = mvPFAppNumber.ToString
    vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTraderAppBank, vList)
    If vDataTable IsNot Nothing Then
      mvDataSet.Tables.Add(vDataTable)
    End If
    If mvDataSet IsNot Nothing Then dgr.Populate(mvDataSet)
  End Sub
  Private Sub SetFieldsForProvisionalPayPlans(Optional ByVal pEnabled As Boolean = False)
    'General tab
    chkSelectBatch.Enabled = pEnabled
    chkCarriage.Enabled = pEnabled
    chkConfirmCarriage.Enabled = pEnabled
    chkBypass.Enabled = pEnabled
    chkConfirmProduct.Enabled = pEnabled
    chkConfirmAnalysis.Enabled = pEnabled
    ChkPaymentMethod.Enabled = pEnabled
    chkTransactionComment.Enabled = pEnabled
    chkAnalysisComments.Enabled = pEnabled
    chkAutoSetAmount.Enabled = pEnabled
    chkNonFinancialBatch.Enabled = pEnabled
    chkCash.Enabled = pEnabled
    chkCheque.Enabled = pEnabled
    chkPostalOrder.Enabled = pEnabled
    chkCreditSale.Enabled = pEnabled
    chkChequeWithInvoice.Enabled = pEnabled
    chkCCWithInvoice.Enabled = pEnabled
    chkCreditCard.Enabled = pEnabled
    chkDebitCard.Enabled = pEnabled
    chkVoucher.Enabled = pEnabled
    chkCAFCard.Enabled = pEnabled
    chkGiftInKind.Enabled = pEnabled
    chkSaleOrReturn.Enabled = pEnabled
    chkPaymentPlan.Enabled = pEnabled
    chkPrefulfilledIncentives.Enabled = pEnabled
    chkContactAlerts.Enabled = pEnabled
    'Analysis tab
    chkProduct.Enabled = pEnabled
    chkDonation.Enabled = pEnabled
    chkEventBooking.Enabled = pEnabled
    chkExamBooking.Enabled = pEnabled
    chkAccomodationBooking.Enabled = pEnabled
    chkServiceBooking.Enabled = pEnabled
    chkServiceBookingCredit.Enabled = pEnabled
    chkPayment.Enabled = pEnabled
    chkMembershipType.Enabled = pEnabled
    chkNoPaymentRequired.Enabled = pEnabled
    chkAddressMaintenance.Enabled = pEnabled
    chkCancelGiftAidDeclaration.Enabled = pEnabled
    ''Bank Acounts tab
    txtCAFAndVoucherAccount.Enabled = pEnabled
    ''Documents tab
    txtInvoiceDoc.Enabled = pEnabled
    txtReceiptDoc.Enabled = pEnabled
    txtCreditStmtDoc.Enabled = pEnabled
    ''Currency BA's tab
    txtCurrency.Enabled = pEnabled
    txtBatchType.Enabled = pEnabled
    txtBankAccount.Enabled = pEnabled
    ''Defaults tab
    txtCarriageProduct.Enabled = pEnabled
    txtCarriageRate.Enabled = pEnabled
    txtPercentage.Enabled = pEnabled

    If pEnabled = False Then
      'General tab
      chkSelectBatch.Checked = False
      chkCarriage.Checked = False
      chkBypass.Checked = False
      chkConfirmProduct.Checked = False
      chkConfirmAnalysis.Checked = False
      ChkPaymentMethod.Checked = False
      chkTransactionComment.Checked = False
      chkAnalysisComments.Checked = False
      chkAutoSetAmount.Checked = False
      chkNonFinancialBatch.Checked = False
      chkPaymentPlan.Checked = False
      chkCreditSale.Checked = False
      chkChequeWithInvoice.Checked = False
      chkCCWithInvoice.Checked = False
      chkInvoicePayment.Checked = False
      chkSundryCreditNote.Checked = False
      chkProduct.Checked = False
      chkDonation.Checked = False
      chkEventBooking.Checked = False
      chkExamBooking.Checked = False
      chkAccomodationBooking.Checked = False
      chkServiceBooking.Checked = False
      chkServiceBookingCredit.Checked = False
      chkPayment.Checked = False
      chkMembershipType.Checked = False
      chkCash.Checked = False
      chkCheque.Checked = False
      chkPostalOrder.Checked = False
      chkCreditCard.Checked = False
      chkDebitCard.Checked = False
      chkAddressMaintenance.Checked = False
      chkCancelGiftAidDeclaration.Checked = False
    End If
  End Sub
  Private Sub SetAppFieldsEnabled()


    Dim vPMTypes As Integer
    Dim vEnable As Boolean

    Select Case mvApplicationType
      Case "CLREC", "BSPOS"
        chkVoucher.Checked = False
        chkCAFCard.Checked = False
        chkGiftInKind.Checked = False
        chkSaleOrReturn.Checked = False
        chkSelectBatch.Checked = True
        chkCash.Checked = True
        chkCheque.Checked = False
        chkPostalOrder.Checked = False
        chkCreditCard.Checked = False
        chkDebitCard.Checked = False
        chkCreditSale.Checked = False
        chkChequeWithInvoice.Checked = False
        chkCCWithInvoice.Checked = False
        chkPaymentPlan.Checked = False
        chkSelectBatch.Enabled = False

        EnablePaymentMethods(False)

        If mvApplicationType = "CLREC" And mvAlbacsBankDetails Then
          cboAlbacsBankDetails.SelectedValue = "C"
          cboAlbacsBankDetails.Enabled = False
        End If
        AdjustTabs(mvApplicationType)
      Case "GAYEP", "POTPG"
        'General Tab
        chkSelectBatch.Enabled = False
        chkBankDetails.Enabled = False
        chkTransactionComment.Enabled = False
        chkConfirmProduct.Enabled = False
        chkConfirmAnalysis.Enabled = False
        chkAnalysisComments.Enabled = False
        chkConfirmCarriage.Enabled = False
        ChkPaymentMethod.Enabled = False
        chkAutoSetAmount.Enabled = False
        chkCarriage.Enabled = False
        chkConfirmDetails.Enabled = False
        chkBypass.Enabled = False
        chkNonFinancialBatch.Enabled = False
        chkPrefulfilledIncentives.Enabled = False
        EnablePaymentMethods(False)

        'Bank Accounts Tab is update from InitForm once all fields populated

        'Other Tabs
        tabMain.TabPages.Remove(tbpAnalysis)
        tabMain.TabPages.Remove(tbpDefaults)
        tabMain.TabPages.Remove(tbpDocuments)
        tabMain.TabPages.Remove(tbpRestrictions)
        AdjustTabs(mvApplicationType)
      Case Else
        EnablePaymentMethods(True)
        vPMTypes = GetPMTypeCount()
        If chkNonFinancialBatch.Checked Then vPMTypes = vPMTypes + 1
        If vPMTypes > 1 Then
          chkSelectBatch.Checked = False
          chkSelectBatch.Enabled = False
        Else
          If chkVoucher.Checked = False And chkIncludeProvPaymentPlan.Checked = False Then
            chkSelectBatch.Enabled = True
          End If
        End If
        If mvApplicationType = "TRANS" And mvEditMode = False Then chkIncludeProvPaymentPlan.Enabled = True
        If mvAlbacsBankDetails Then
          cboAlbacsBankDetails.Enabled = True
        End If
        AdjustTabs(mvApplicationType)
    End Select

    If GetPMTypeCount(GetPayMethodCountTypes.gpmctStandard) > 1 And chkCAFCard.Checked = False And chkIncludeProvPaymentPlan.Checked = False Then
      ChkPaymentMethod.Enabled = True
    Else
      ChkPaymentMethod.Checked = False
      ChkPaymentMethod.Enabled = False
    End If

    If mvApplicationType = "TRANS" Then
      chkIncludeProvPaymentPlan.Enabled = True
    Else
      chkIncludeProvPaymentPlan.Checked = False
      chkIncludeProvPaymentPlan.Enabled = False
    End If

    If mvApplicationType = "CNVRT" And mvPayPlanConvMaintenance = True Then
      chkMaintainPaymentPlan.Enabled = True
    Else
      chkMaintainPaymentPlan.Checked = False
      chkMaintainPaymentPlan.Enabled = False
    End If

    'Set the enabled status of the display scheduled payments checkbox
    vEnable = False

    If mvScheduledPayments And (mvApplicationType = "TRANS" Or mvApplicationType = "MAINT" Or (mvApplicationType = "CNVRT" And chkMaintainPaymentPlan.Checked = False)) And chkMembershipType.Checked = False Then
      If chkMembership.Checked Or chkSubscription.Checked Or _
      chkRegularDonation.Checked Or chkCOvenantedMembership.Checked Or _
      chkCovenantedSubscription.Checked Or chkCovenantedDonation.Checked Then
        vEnable = True
      End If
    End If
    If vEnable Then
      chkDisplayScheduledPayment.Enabled = True
    Else
      chkDisplayScheduledPayment.Enabled = False
      chkDisplayScheduledPayment.Checked = False
    End If

    'Handle Invoice Print controls
    If mvApplicationType = "BINVG" Then
      chkInvoicePrintPreview.Enabled = True
      chkInvoicePrintUnpostedBatches.Enabled = mvInvoicePrintUnpostedBatches
    Else
      chkInvoicePrintPreview.Checked = False
      chkInvoicePrintPreview.Enabled = False
      chkInvoicePrintUnpostedBatches.Checked = False
      chkInvoicePrintUnpostedBatches.Enabled = False
    End If

    'Handle Loans checkbox
    Dim vEnableLoans As Boolean
    If ((mvApplicationType = "TRANS" AndAlso GetPMTypeCount(GetPayMethodCountTypes.gpmctStandard) > 0) OrElse (mvApplicationType = "MAINT" OrElse mvApplicationType = "CNVRT")) AndAlso chkMembershipType.Checked = False Then
      vEnableLoans = True
    End If
    chkLoan.Enabled = vEnableLoans
    If vEnableLoans = False Then chkLoan.Checked = False

  End Sub
  ''' <summary>
  ''' This function is created to add the removed tabs again when transaction
  ''' type is changed from GAYEP or POTPG
  ''' </summary>
  ''' <param name="pTransType"></param>
  ''' <remarks></remarks>
  Private Sub AdjustTabs(ByVal pTransType As String)
    Select Case pTransType
      Case "GAYEP", "POTPG"
        If mvPages.Count = 0 Then
          mvPages.Add("analysis", tbpAnalysis)
          mvPages.Add("defaults", tbpDefaults)
          mvPages.Add("documents", tbpDocuments)
          mvPages.Add("restriction", tbpRestrictions)
        End If
      Case Else
        If mvPages.Count > 0 Then
          tabMain.TabPages.Clear()
          tabMain.TabPages.Add(tbpGeneral)
          tabMain.TabPages.Add(CType(mvPages("analysis"), TabPage))
          tabMain.TabPages.Add(CType(mvPages("defaults"), TabPage))
          tabMain.TabPages.Add(tbpBank)
          tabMain.TabPages.Add(CType(mvPages("documents"), TabPage))
          tabMain.TabPages.Add(CType(mvPages("restriction"), TabPage))
          tabMain.TabPages.Add(tbpCurrency)
          mvPages.Clear()
        End If
    End Select

  End Sub

  Private Sub EnablePaymentMethods(ByVal pValue As Boolean)
    If chkIncludeProvPaymentPlan.Checked = False Then
      chkCash.Enabled = pValue
      chkCheque.Enabled = pValue
      chkPostalOrder.Enabled = pValue
      chkCreditCard.Enabled = pValue
      chkDebitCard.Enabled = pValue
      chkCreditSale.Enabled = pValue
      chkChequeWithInvoice.Enabled = pValue
      chkCCWithInvoice.Enabled = pValue
      chkPaymentPlan.Enabled = pValue
      If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_card_sales_combined_claim) = "A" Or AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_card_sales_combined_claim) = "Y" Then
        chkDebitCard.Enabled = False
        chkDebitCard.Checked = False
      End If
      chkVoucher.Enabled = pValue
      chkCAFCard.Enabled = pValue
      chkGiftInKind.Enabled = pValue
      chkSaleOrReturn.Enabled = pValue
    End If
  End Sub

  Private Function GetPMTypeCount(Optional ByVal pType As GetPayMethodCountTypes = GetPayMethodCountTypes.gpmctAll) As Integer
    Dim vPMTypes As Integer
    If pType = GetPayMethodCountTypes.gpmctStandard Then
      If chkCash.Checked = True Then vPMTypes = 1
      If chkCheque.Checked = True Then vPMTypes = vPMTypes + 1
      If chkPostalOrder.Checked = True Then vPMTypes = vPMTypes + 1
    Else
      If chkCash.Checked = True Or chkCheque.Checked = True Or chkPostalOrder.Checked = True Then vPMTypes = 1
    End If
    If chkCreditCard.Checked = True Then vPMTypes = vPMTypes + 1
    If chkDebitCard.Checked = True Then vPMTypes = vPMTypes + 1
    If chkCreditSale.Checked = True Then vPMTypes = vPMTypes + 1
    If chkChequeWithInvoice.Checked = True Then vPMTypes = vPMTypes + 1
    If chkCCWithInvoice.Checked = True Then vPMTypes = vPMTypes + 1
    If pType = GetPayMethodCountTypes.gpmctAll Then
      If chkPaymentPlan.Checked = True Then vPMTypes = vPMTypes + 1
      If chkCAFCard.Checked = True Then vPMTypes = vPMTypes + 1
      If chkVoucher.Checked = True Then vPMTypes = vPMTypes + 1
      If chkGiftInKind.Checked = True Then vPMTypes = vPMTypes + 1
      If chkSaleOrReturn.Checked = True Then vPMTypes = vPMTypes + 1
    End If
    GetPMTypeCount = vPMTypes
  End Function

  Private Enum GetPayMethodCountTypes
    gpmctAll
    gpmctStandard
  End Enum
  Private Sub SetSourceDescription()
    txtSource.Text = mvTable.Rows(0).Item("source").ToString()
  End Sub

  Private Sub SetAdditionalPMChks(ByVal chkAdditionalPM As CheckBox)
    If chkAdditionalPM.Checked Then
      chkEventBooking.Checked = False
      chkExamBooking.Checked = False
      chkAccomodationBooking.Checked = False
      chkServiceBooking.Checked = False
      chkServiceBookingCredit.Checked = False
      chkMembershipType.Checked = False
      chkInvoicePayment.Checked = False
      chkSundryCreditNote.Checked = False
      chkLegacyReceipt.Checked = False
      If (chkAdditionalPM.Name = "chkIncludeProvPaymentPlan" And mvLoading = False) Or chkAdditionalPM.Name <> "chkIncludeProvPaymentPlan" Then
        chkCOvenantedMembership.Checked = False
        chkCovenantedSubscription.Checked = False
        chkCovenantedDonation.Checked = False
      End If
      If chkAdditionalPM.Name <> "chkIncludeProvPaymentPlan" Then
        chkSubscription.Checked = False
        chkRegularDonation.Checked = False
        chkStandingOrder.Checked = False
        chkDirectDebit.Checked = False
        chkCreditCardAuthority.Checked = False
        chkNoPaymentRequired.Checked = False
      End If
      chkGoneAway.Checked = False
      chkStatus.Checked = False
      chkActivity.Checked = False
      chkSuppression.Checked = False
      chkCancelPaymentPlan.Checked = False
      chkGiftAidDeclaration.Checked = False
      chkPayrollGiving.Checked = False
      chkAutoPaymentMaintenance.Checked = False

      chkEventBooking.Enabled = False
      chkExamBooking.Enabled = False
      chkAccomodationBooking.Enabled = False
      chkServiceBooking.Enabled = False
      chkServiceBookingCredit.Enabled = False
      chkMembershipType.Enabled = False

      If chkAdditionalPM.Name <> "chkIncludeProvPaymentPlan" Then
        chkSubscription.Enabled = False
        chkCOvenantedMembership.Enabled = False
        chkSubscription.Enabled = False
        chkRegularDonation.Enabled = False
        chkStandingOrder.Enabled = False
        chkDirectDebit.Enabled = False
        chkCreditCardAuthority.Enabled = False
      End If

      chkInvoicePayment.Enabled = False
      chkSundryCreditNote.Enabled = False
      chkLegacyReceipt.Enabled = False
      chkNoPaymentRequired.Enabled = False
      chkGoneAway.Enabled = False
      chkStatus.Enabled = False
      chkActivity.Enabled = False
      chkSuppression.Enabled = False
      chkCancelPaymentPlan.Enabled = False
      chkGiftAidDeclaration.Enabled = False
      chkPayrollGiving.Enabled = False
      chkAutoPaymentMaintenance.Enabled = False

      If chkAdditionalPM.Name = "chkGiftInKind" Or chkAdditionalPM.Name = "chkSaleOrReturn" Then
        chkDonation.Checked = False
        chkDonation.Enabled = False
        chkMembership.Checked = False
        chkMembership.Enabled = False
        chkPayment.Checked = False
        chkPayment.Enabled = False
        chkLoan.Checked = False
      Else
        If chkAdditionalPM.Name = "chkCAFCard" Then
          ChkPaymentMethod.Checked = False
          ChkPaymentMethod.Enabled = False
        End If
        If Not (chkAdditionalPM.Name = "chkCAFCard" OrElse chkAdditionalPM.Name = "chkVoucher") Then
          'BR14798: Allow the Analysis- Sales- Product option to be selected for CAF Card and CAF Voucher Payment Methods
          chkProduct.Checked = False
          chkProduct.Enabled = False
        End If
      End If
    Else
      chkProduct.Enabled = True
      chkEventBooking.Enabled = True
      chkExamBooking.Enabled = True
      chkAccomodationBooking.Enabled = True
      chkServiceBooking.Enabled = True
      chkServiceBookingCredit.Enabled = True
      chkSubscription.Enabled = True
      chkPayment.Enabled = True
      chkMembershipType.Enabled = True
      If AppValues.ConfigurationValue(AppValues.ConfigurationValues.option_covenants) = "Yes" Then
        chkCOvenantedMembership.Enabled = True
        chkCovenantedSubscription.Enabled = True
        chkCovenantedDonation.Enabled = True
      End If
      chkInvoicePayment.Enabled = True
      chkSundryCreditNote.Enabled = True
      chkLegacyReceipt.Enabled = True
      chkDonation.Enabled = True
      chkStandingOrder.Enabled = True
      chkDirectDebit.Enabled = True
      chkCreditCardAuthority.Enabled = True
      chkNoPaymentRequired.Enabled = True
      chkGoneAway.Enabled = True
      chkStatus.Enabled = True
      chkActivity.Enabled = True
      chkSuppression.Enabled = True
      chkCancelPaymentPlan.Enabled = True
      chkGiftAidDeclaration.Enabled = True
      chkPayrollGiving.Enabled = True
      chkLoan.Enabled = True

      If chkAdditionalPM.Name = "chkGiftInKind" Or chkAdditionalPM.Name = "chkSaleOrReturn" Then
        chkDonation.Enabled = True
        chkMembership.Enabled = True
      ElseIf chkAdditionalPM.Name = chkCAFCard.Name Then
        ChkPaymentMethod.Enabled = True
      End If
    End If
  End Sub
  Private Sub SetDefaultFieldsState(ByVal pChkBox As CheckBox)
    Dim vTextBox As TextLookupBox = Nothing
    Dim vMultiple As Boolean
    Dim vEnabled As Boolean

    Select Case pChkBox.Name
      Case "chkStatus"
        vTextBox = txtStatus
      Case "chkActivity"
        vTextBox = txtActivity
      Case "chkSuppression"
        vTextBox = txtSuppression
      Case "chkCancelPaymentPlan", "chkCancelGiftAidDeclaration"
        vTextBox = txtCancelGiftAidDeclaration
        vMultiple = True
    End Select

    If vMultiple Then
      vEnabled = chkCancelPaymentPlan.Checked Or chkCancelGiftAidDeclaration.Checked
    Else
      vEnabled = pChkBox.Checked
    End If
    vTextBox.Enabled = vEnabled
    If Not vTextBox.Enabled Then vTextBox.Text = ""

  End Sub
  Private Sub DisableBankAccountFields(ByVal pIgnoreField As TextLookupBox)
    'Disable fields if it contains a valid value
    If txtCashAccount.Name <> pIgnoreField.Name And txtCashAccount.Text.Length > 0 Then txtCashAccount.Enabled = False
    If txtCreditCardAccount.Name <> pIgnoreField.Name And txtCreditCardAccount.Text.Length > 0 Then txtCreditCardAccount.Enabled = False
    If txtDebitCardAccount.Name <> pIgnoreField.Name And txtDebitCardAccount.Text.Length > 0 Then txtDebitCardAccount.Enabled = False
    If txtCreditSaleAccount.Name <> pIgnoreField.Name And txtCreditSaleAccount.Text.Length > 0 Then txtCreditSaleAccount.Enabled = False
    If txtStandingOrderAccount.Name <> pIgnoreField.Name And txtStandingOrderAccount.Text.Length > 0 Then txtStandingOrderAccount.Enabled = False
    If txtDirectDebitAccount.Name <> pIgnoreField.Name And txtDirectDebitAccount.Text.Length > 0 Then txtDirectDebitAccount.Enabled = False
    If txtCAFAndVoucherAccount.Name <> pIgnoreField.Name And txtCAFAndVoucherAccount.Text.Length > 0 Then txtCAFAndVoucherAccount.Enabled = False
  End Sub
  Private Sub InitTextLookupBox(ByRef pTxtLookup As TextLookupBox, ByVal pValidationTable As String, ByVal pValidationAttribute As String, Optional ByVal pVal As String = "")

    Dim vParamList As New ParameterList(True)
    vParamList("TableName") = pValidationTable
    vParamList("FieldName") = pValidationAttribute
    vParamList("FieldType") = "C"  ' Character FieldType
    Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vParamList)
    vParams("AttributeName") = pValidationAttribute
    vParams("ValidationAttribute") = pValidationAttribute
    vParams("ValidationTable") = pValidationTable

    pTxtLookup.BackColor = Me.BackColor
    Dim vPanelItem As PanelItem = New PanelItem(pTxtLookup, pValidationAttribute)
    vPanelItem.InitFromMaintenanceData(vParams)
    Select Case vPanelItem.ParameterName
      Case "CancellationReason", "Status", "BatchCategory", "BatchAnalysisCode", "Source", "ContactNumber", "TransactionOrigin", _
        "Product", "Rate", "Branch", "StandardDocument", "SalesGroup", "CurrencyCode", "BatchType", "CreditCategory", _
        "ExamSessionCode", "ExamUnitCode", "MerchantRetailNumber"
        vPanelItem.Mandatory = False
      Case "BankAccount"
        If pTxtLookup.Name = "txtBankAccount" Then vPanelItem.Mandatory = False
    End Select
    pTxtLookup.Tag = vPanelItem
    AddHandler pTxtLookup.GetInitialCodeRestrictions, AddressOf GetInitialCodeRestrictionsHandler
    pTxtLookup.Init(vPanelItem, False, False)
    pTxtLookup.ActiveOnly = True
    pTxtLookup.TotalWidth = pTxtLookup.Width
    pTxtLookup.BackColor = Color.Transparent
    pTxtLookup.SetBounds(pTxtLookup.Location.X, pTxtLookup.Location.Y, 80, pTxtLookup.TextBox.Size.Height)

    AddHandler pTxtLookup.GetCodeRestrictions, AddressOf GetCodeRestrictionsHandler
    AddHandler pTxtLookup.Validating, AddressOf LookupValidatingHandler
    AddHandler pTxtLookup.TextChanged, AddressOf LookupChangedHandler

    pTxtLookup.Text = ""
    pTxtLookup.SetBounds(pTxtLookup.Location.X, pTxtLookup.Location.Y, 80, EditPanelInfo.DefaultHeight)
  End Sub
  Private Sub LookupChangedHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vParamTextBox As TextLookupBox = DirectCast(sender, TextLookupBox)
      If mvEditMode = False Then
        If vParamTextBox.Name = "txtType" Then
          mvApplicationType = txtType.Text
          If CanDesignApp() Then
            cmdDesign.Enabled = True
          Else
            cmdDesign.Enabled = False
          End If
          SetAppFieldsEnabled()
        End If
      End If
      If sender Is txtCurrency OrElse sender Is txtBatchType OrElse sender Is txtBankAccount Then
        cmdAdd.Enabled = IsAddValid()
      ElseIf IsBankAccountControl(vParamTextBox) AndAlso vParamTextBox.IsValid Then
        Dim vAccountsSet As Boolean = False
        For Each vControl As TextLookupBox In mvBankAccountControls
          If sender IsNot vControl AndAlso vControl.Text.Length > 0 Then
            vAccountsSet = True
            Exit For
          End If
        Next
        If vAccountsSet = False Then
          For Each vControl As TextLookupBox In mvBankAccountControls
            If vControl IsNot vParamTextBox Then
              vControl.Text = vParamTextBox.Text
            End If
          Next
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function IsBankAccountControl(ByVal pControl As Object) As Boolean
    For Each vControl As Control In mvBankAccountControls
      If pControl Is vControl Then Return True
    Next
    Return False
  End Function

  Private Sub LookupValidatingHandler(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
    Try
      Dim vTextLookupBox As TextLookupBox = DirectCast(sender, TextLookupBox)
      If ValidateControl(vTextLookupBox, DirectCast(vTextLookupBox.Tag, PanelItem), vTextLookupBox.Text) Then
        Select Case DirectCast(vTextLookupBox.Tag, PanelItem).ParameterName
          Case "ExamSessionCode"
            Dim vSessionId As Integer = vTextLookupBox.GetDataRowInteger("ExamSessionId")
            txtExamUnit.Text = ""
            txtExamUnit.FillComboWithRestriction(vSessionId.ToString)
        End Select
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub EnableTypeChange(ByVal pEnable As Boolean)
    txtType.Enabled = pEnable
    If CanDesignApp() AndAlso Not mvReadOnly Then
      cmdDesign.Enabled = True
    Else
      cmdDesign.Enabled = False
    End If
    SetAppFieldsEnabled()
  End Sub
  Private Sub SetDefaults()
    chkCash.Checked = True
    chkCheque.Checked = True
    chkPostalOrder.Checked = True
    chkBankDetails.Checked = True
    chkTransactionComment.Checked = True
    chkProduct.Checked = True
    chkShowReference.Checked = True
    chkConfirmProduct.Checked = True
    chkCarriage.Checked = True
    chkAnalysisComments.Checked = True
    chkConfirmCarriage.Checked = True
    chkIncludeConfirmedTransaction.Checked = True
    txtLinkToCommunication.Text = "N"
    cboAlbacsBankDetails.SelectedIndex = 3
    autoGiftAidDeclaration.Checked = False
    gadMethodOral.Checked = False
    gadMethodWritten.Checked = False
    gadMethodElectronic.Checked = True
    gadSource.Text = String.Empty
  End Sub

  Private Function CanDesignApp() As Boolean
    If mvApplicationType = "POPRT" Or mvApplicationType = "POGEN" Or mvApplicationType = "POCHQ" Then
      Return False
    Else
      Return True
    End If
  End Function
  Private Function getAlbacsIndex(ByVal pValue As String) As Integer
    Select Case pValue
      Case "w"
        Return 0
      Case "E"
        Return 1
      Case "N"
        Return 2
      Case "C"
        Return 3
    End Select

  End Function

  Private Sub cmdCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vReturnList As New ParameterList
      vList("FpApplicationNumber") = txtApplication.Text
      vReturnList = DataHelper.CopyTraderApplication(vList)
      If vReturnList("NewAppNumber").Length > 0 Then
        mvPFAppNumber = IntegerValue(vReturnList("NewAppNumber"))
        SetTraderValues(True)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vReturnList As New ParameterList
      If ConfirmDelete() Then
        vList("FpApplicationNumber") = txtApplication.Text
        vList("Flag") = "D"
        vReturnList = DataHelper.DeleteTraderApplication(vList)
        Me.Close()
      End If
    Catch vException As CareException
      Select Case vException.ErrorNumber
        Case CareException.ErrorNumbers.enConfigurationExist
          ShowInformationMessage(vException.Message)
        Case Else
          DataHelper.HandleException(vException)
      End Select
    End Try
  End Sub

  Private Sub cmdRevert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRevert.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vReturnList As New ParameterList
      If ShowQuestion(QuestionMessages.QmRevertApplication, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        vList("FpApplicationNumber") = txtApplication.Text
        vList("Flag") = "R"
        vReturnList = DataHelper.DeleteTraderApplication(vList)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub CheckChangedHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)

    Dim vOTypes As Integer
    Dim vOTypesVal As Boolean
    Dim vSalesTypes As Boolean
    Dim vMsg As String = ""
    Dim vInvalid As Boolean
    Dim vToken As Boolean

    Dim vParamChkBox As CheckBox = DirectCast(sender, CheckBox)

    Try
      If mvLoading Then Exit Sub
      If mvEditMode = True AndAlso vParamChkBox.Checked Then
        Select Case vParamChkBox.Name
          Case "chkCash", "chkCheque", "chkPostalOrder", "chkCreditCard", "chkDebitCard", "chkCreditSale", "chkPaymentPlan", "chkVoucher", "chkCAFCard", "chkGiftInKind", "chkSaleOrReturn", "chkCCWithInvoice"
            If vParamChkBox.Name <> "chkCreditSale" Then
              If chkSundryCreditNote.Checked Or chkServiceBookingCredit.Checked Then
                vParamChkBox.Checked = False
                If chkSundryCreditNote.Checked Then
                  'to do
                  vMsg = InformationMessages.ImPaymentMethodNotForSundry '"This payment method cannot be supported as this application supports Sundry Credit Notes"
                Else
                  'to do
                  vMsg = InformationMessages.ImPaymentMethodNotForServiceBookingCredits '"This payment method cannot be supported as this application supports Service Booking Credits"
                End If
                'to do
                ShowInformationMessage(vMsg)
              Else
                'mvPM1Changed = True
              End If
            Else
              'mvPM1Changed = True
            End If
            If chkConfirmCollection.Checked Then
              Select Case vParamChkBox.Name
                Case "chkCreditCard", "chkDebitCard", "chkCreditSale", "chkVoucher", "chkCAFCard", "chkGiftInKind", "chkSaleOrReturn"
                  If vParamChkBox.Checked Then
                    'to do
                    ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "collection payments") '"This payment method cannot be supported as this application supports Collection Payments"
                    vParamChkBox.Checked = False
                  End If
              End Select
            End If
            Select Case vParamChkBox.Name
              Case "chkCAFCard"
                If chkOnlineAuth.Checked Then
                  ShowInformationMessage(InformationMessages.ImTransactionPayCAFCCInvalid)  'CAF Card payment method invalid when using On-Line Credit Card Authorisation
                  vParamChkBox.Checked = False
                ElseIf chkDebitCard.Checked OrElse chkCreditCard.Checked OrElse chkCCWithInvoice.Checked Then
                  ShowInformationMessage(InformationMessages.ImTraderCAFCardandCardPayInvalid)  'CAF Card and Card Payment methods cannot be used together
                  vParamChkBox.Checked = False
                End If
              Case "chkDebitCard", "chkCreditCard", "chkCCWithInvoice"
                If chkCAFCard.Checked Then
                  ShowInformationMessage(InformationMessages.ImTraderCAFCardandCardPayInvalid)  'CAF Card and Card Payment methods cannot be used together
                  vParamChkBox.Checked = False
                End If
            End Select
          Case "chkMembership", "chkSubscription", "chkRegularDonation", "chkProduct", "chkDonation", "chkPayment", "chkCOvenantedMembership", "chkCovenantedSubscription", "chkCovenantedDonation", "chkEventBooking", "chkExamBooking", "chkAccomodationBooking", "chkInvoicePayment", "chkServiceBooking", "chkSundryCreditNote", "chkServiceBookingCredit", "chkLegacyReceipt", "chkStatus", "chkGoneAway", "chkGiftAidDeclaration", "chkActivity", "chkSuppression", "chkCancelPaymentPlan", "chkPayrollGiving", "chkLoan"
            If vParamChkBox.Name <> "chkSundryCreditNote" And chkSundryCreditNote.Checked Then
              vParamChkBox.Checked = False
              ShowInformationMessage(InformationMessages.ImPaymentMethodNotForSundry) '"This analysis option cannot be supported as this application supports Sundry Credit Notes"
            ElseIf vParamChkBox.Name <> "chkProduct" AndAlso vParamChkBox.Name <> "chkServiceBookingCredit" AndAlso chkServiceBookingCredit.Checked Then
              vParamChkBox.Checked = False
              ShowInformationMessage(InformationMessages.ImPaymentMethodNotForServiceBookingCredits) '"This analysis option cannot be supported as this application supports Service Booking Credits"
            Else
              'mvTAChanged = True
            End If
          Case "chkStandingOrder", "chkDirectDebit", "chkCreditCardAuthority"
            'mvPM2Changed = True
        End Select
      End If

      'Now allows provisional cash so check no incompatible settings
      Select Case vParamChkBox.Name
        Case "chkOnlineAuth"
          Me.chkRequireAuthorisation.Enabled = vParamChkBox.Checked
          Me.chkRequireAuthorisation.Checked = vParamChkBox.Checked
          If vParamChkBox.Checked Then
            If chkCAFCard.Checked Then
              ShowInformationMessage(InformationMessages.ImTransactionPayCAFCCInvalid) 'CAF Card payment method invalid when using On-Line Credit Card Authorisation
              vParamChkBox.Checked = False
            End If
          End If
        Case "chkCash", "chkCheque", "chkPostalOrder", "chkCreditCard", "chkDebitCard", "chkCreditSale", "chkPaymentPlan", "chkVoucher", "chkCAFCard", "chkGiftInKind", "chkSaleOrReturn"
          If chkVoucher.Checked = False And chkGiftInKind.Checked = False And chkCAFCard.Checked = False And chkSaleOrReturn.Checked = False Then
            If vParamChkBox.Name = "chkCash" Then
              'Enable/disable Confirmed/Provisional Trans checkboxes
              If vParamChkBox.Checked = True Then
                chkIncludeConfirmedTransaction.Enabled = True
                chkIncludeProvisionalTransaction.Enabled = True
              Else
                If chkIncludeProvisionalTransaction.Checked Then
                  chkIncludeProvisionalTransaction.Checked = False
                  chkIncludeConfirmedTransaction.Checked = False
                End If
                chkIncludeConfirmedTransaction.Enabled = False
                chkIncludeProvisionalTransaction.Enabled = False
              End If
            Else
              If chkCreditSale.Checked = True Then
                chkAutoCreateCreditCustomer.Enabled = True
                txtCreditCategory.Enabled = True
              Else
                If txtCreditCategory.ComboBox.Items.Count > 0 Then txtCreditCategory.ComboBox.SelectedIndex = 0
                txtCreditCategory.Text = ""
                txtCreditCategory.Enabled = False
                chkAutoCreateCreditCustomer.Enabled = False
                chkAutoCreateCreditCustomer.Checked = False
              End If

              If chkCash.Checked And chkIncludeProvisionalTransaction.Checked Then
                If vParamChkBox.Checked Then
                  vParamChkBox.Checked = False
                  'to do
                  ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Provisional Cash Batches") '"This payment method cannot be supported as the application supports Provisional Cash Batches"
                End If
              End If
            End If
          End If

        Case "chkIncludeConfirmedTransaction", "chkIncludeProvisionalTransaction"
          If vParamChkBox.Name = "chkIncludeProvisionalTransaction" And vParamChkBox.Enabled Then
            If vParamChkBox.Checked Then
              'Only allowed for Cash payment method
              chkIncludeConfirmedTransaction.Checked = False

              If chkCheque.Checked Then vInvalid = True
              If chkPostalOrder.Checked AndAlso vInvalid = False Then vInvalid = True
              If chkCreditCard.Checked AndAlso vInvalid = False Then vInvalid = True
              If chkDebitCard.Checked AndAlso vInvalid = False Then vInvalid = True
              If chkCreditSale.Checked AndAlso vInvalid = False Then vInvalid = True

              If vInvalid = False Then
                If chkVoucher.Checked Or chkCAFCard.Checked Or chkGiftInKind.Checked Or chkPaymentPlan.Checked Or chkSaleOrReturn.Checked Then
                  vInvalid = True
                End If
              End If
              If vInvalid Then
                ShowInformationMessage(InformationMessages.ImOptionNotSupportedCashPaymentIsValid) '"This option cannot be supported as only the Cash payment method is valid for Provisional Cash Transactions"
                chkIncludeProvisionalTransaction.Checked = False
                chkIncludeConfirmedTransaction.Checked = True
              End If
            End If
          ElseIf vParamChkBox.Name = "chkIncludeConfirmedTransaction" And vParamChkBox.Enabled Then
            If vParamChkBox.Checked Then chkIncludeProvisionalTransaction.Checked = False
          End If

      End Select

      'BR20526
      If vParamChkBox.Name = "chkVoucher" AndAlso vParamChkBox.Checked Then
        ShowInformationMessage(InformationMessages.ImAddAdditionalReferenceFields)
      End If
      If vParamChkBox.Name = "chkCheque" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        'to do
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkPostalOrder" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkDebitCard" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkCreditSale" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If

      Select Case vParamChkBox.Name
        Case "chkCash", "chkCheque", "chkPostalOrder", "chkCreditCard", "chkDebitCard", "chkCreditSale", "chkPaymentPlan", "chkVoucher", "chkGiftInKind", "chkCAFCard", "chkSaleOrReturn"
          If vParamChkBox.Checked Then
            If vParamChkBox.Name <> "chkVoucher" And chkVoucher.Checked Then
              vInvalid = True
              'to do
              vMsg = InformationMessages.ImPaymentMethodNotForVoucher '"This payment method cannot be supported as the application supports Voucher payments"
            ElseIf vParamChkBox.Name <> "chkGiftInKind" And chkGiftInKind.Checked Then
              vInvalid = True
              'to do
              vMsg = InformationMessages.ImPaymentMethodNotForGiftInKind '"This payment method cannot be supported as the application supports Gift in Kind payments"
            ElseIf vParamChkBox.Name <> "chkSaleOrReturn" And chkSaleOrReturn.Checked Then
              vInvalid = True
              'to do
              vMsg = InformationMessages.ImPaymentMethodNotForReturn '"This payment method cannot be supported as the application supports Sale or Return payments"
            End If
            If vInvalid Then
              vParamChkBox.Checked = False
              ShowInformationMessage(vMsg)
            End If
          Else
            'If last Confirmed PM switched off, switch off Confirmed chk.
            If chkCash.Checked = False And chkCheque.Checked = False And chkPostalOrder.Checked = False And chkCreditCard.Checked = False And chkDebitCard.Checked = False And chkCreditSale.Checked = False Then
              chkIncludeConfirmedTransaction.Checked = False
            End If
          End If

        Case "chkIncludeProvPaymentPlan"
          If vParamChkBox.Checked Then
            If chkIncludeProvisionalTransaction.Checked Then
              vInvalid = True
              vParamChkBox.Checked = False
              'to do
              ShowInformationMessage(InformationMessages.ImPaymentOptionNotForProvisionalTransactions) '"This option can not be chosen because this application supports Provisional Transactions"
            End If

            If Not vInvalid Then
              chkCreditSale.Enabled = False
              chkPayment.Enabled = False
              chkIncludeConfirmedTransaction.Checked = False
              chkNoPaymentRequired.Checked = True
              SetAdditionalPMChks(vParamChkBox)
              SetFieldsForProvisionalPayPlans()
            End If
          Else
            SetFieldsForProvisionalPayPlans(True)
          End If
      End Select

      Select Case vParamChkBox.Name
        Case "chkGiftInKind", "chkVoucher", "chkCAFCard", "chkSaleOrReturn"
          If vParamChkBox.Checked And chkConfirmSale.Checked Then
            vParamChkBox.Checked = False
            'to do
            ShowInformationMessage(InformationMessages.ImOptionNotSupportedSaleOrReturnSelected) '"This option cannot be supported because Confirm Sale or Return Transactions analysis option is selected"
          End If
          If vParamChkBox.Checked Then
            chkIncludeProvisionalTransaction.Checked = True
            If vParamChkBox.Name <> "chkCAFCard" Then
              chkCash.Checked = False
              chkCheque.Checked = False
              chkChequeWithInvoice.Checked = False
              chkPostalOrder.Checked = False
              chkCreditCard.Checked = False
              chkCCWithInvoice.Checked = False
              chkDebitCard.Checked = False
              chkCreditSale.Checked = False
              chkPaymentPlan.Checked = False
              chkCAFCard.Checked = False
              chkIncludeConfirmedTransaction.Checked = False
            Else
              If chkCash.Checked Then chkIncludeConfirmedTransaction.Checked = True

            End If
            chkIncludeConfirmedTransaction.Enabled = False
            chkIncludeProvisionalTransaction.Enabled = False
          Else
            If chkGiftInKind.Checked = False And chkVoucher.Checked = False And chkCAFCard.Checked = False And chkSaleOrReturn.Checked = False Then
              chkIncludeProvisionalTransaction.Checked = False
              If chkCash.Checked Then
                chkIncludeProvisionalTransaction.Enabled = True
                chkIncludeConfirmedTransaction.Enabled = True
              End If
            End If
          End If

        Case "chkCash", "chkCheque", "chkPostalOrder", "chkCreditCard", "chkDebitCard", "chkCreditSale", "chkPaymentPlan"
          If vParamChkBox.Checked Then
            chkIncludeConfirmedTransaction.Checked = True
          End If
      End Select

      If vParamChkBox.Name = "chkGiftInKind" Or vParamChkBox.Name = "chkCAFCard" Or vParamChkBox.Name = "chkVoucher" Or vParamChkBox.Name = "chkSaleOrReturn" Then
        If vParamChkBox.Name <> "chkVoucher" And chkVoucher.Checked Then
          '
        ElseIf vParamChkBox.Name <> "chkGiftInKind" And chkGiftInKind.Checked Then
          '
        ElseIf vParamChkBox.Name <> "chkSaleOrReturn" And chkSaleOrReturn.Checked Then
          '
        Else
          SetAdditionalPMChks(vParamChkBox)
        End If
      End If

      If vParamChkBox.Name = "chkPaymentPlan" And chkPaymentPlan.Checked Then
        If chkForeignCurrency.Checked Then
          chkPaymentPlan.Checked = False
          'to do
          ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        End If
      End If


      If vParamChkBox.Name = "chkVoucher" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
      End If
      If vParamChkBox.Name = "chkCAFCard" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
      End If
      If vParamChkBox.Name = "chkGiftInKind" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
      End If

      If vParamChkBox.Name = "chkIncludeProvPaymentPlan" AndAlso vParamChkBox.Checked Then
        If chkForeignCurrency.Checked Then
          vParamChkBox.Checked = False
          ShowInformationMessage(InformationMessages.ImProvPayPlanNotForForeignCurrency) '"Provisional Payment Plans can not be supported as the application supports Foreign Currency"
        End If
      End If

      If vParamChkBox.Name = "chkSaleOrReturn" And chkSaleOrReturn.Checked And chkForeignCurrency.Checked Then
        chkSaleOrReturn.Checked = False
        ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
      End If

      vToken = False
      If vParamChkBox.Name = "chkMembership" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkRegularDonation" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkProduct" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkDonation" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkStandingOrder" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkDirectDebit" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkCreditCardAuthority" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkCOvenantedMembership" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkCovenantedSubscription" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkCovenantedDonation" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkEventBooking" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkExamBooking" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkAccomodationBooking" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkInvoicePayment" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkServiceBooking" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vToken = False And vParamChkBox.Name = "chkSundryCreditNote" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If
      If vParamChkBox.Name = "chkLoan" And vParamChkBox.Checked And chkForeignCurrency.Checked Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImAnalysisNotForForeignCurrency) '"This analysis option cannot be supported as the application supports Foreign Currency"
        vToken = True
      End If

      If vParamChkBox.Name = "chkCreditCard" OrElse vParamChkBox.Name = "chkDebitCard" OrElse vParamChkBox.Name = "chkCCWithInvoice" Then
        If chkCreditCard.Checked OrElse chkDebitCard.Checked OrElse chkCCWithInvoice.Checked Then
          chkOnlineAuth.Enabled = True
          chkRequireAuthorisation.Enabled = chkOnlineAuth.Checked
        Else
          chkOnlineAuth.Enabled = False
          chkOnlineAuth.Checked = False
        End If
      End If


      'We currently enforce Applications to be Provisional OR Confirmed and set values:
      'if Gift in Kind/Voucher: Provisional, everything else: Confirmed.
      'Eventually we will allow mixed when there's a need, e.g. Pro-Forma TA 9/11
      SetAppFieldsEnabled()      'Deal with payment methods and batch led type

      If chkMembership.Checked Then vOTypes = vOTypes + 1
      If chkSubscription.Checked Then vOTypes = vOTypes + 1
      If chkRegularDonation.Checked Then vOTypes = vOTypes + 1
      If chkCOvenantedMembership.Checked Then vOTypes = vOTypes + 1
      If chkCovenantedSubscription.Checked Then vOTypes = vOTypes + 1
      If chkCovenantedDonation.Checked Then vOTypes = vOTypes + 1
      If chkPaymentPlan.Checked Then vOTypes = vOTypes + 1
      If chkLoan.Checked Then vOTypes += 1

      If vOTypes > 0 And chkForeignCurrency.Checked = False Then vOTypesVal = True
      chkStandingOrder.Enabled = vOTypesVal
      chkDirectDebit.Enabled = vOTypesVal
      chkCreditCardAuthority.Enabled = vOTypesVal
      If chkIncludeProvPaymentPlan.Checked = False Then chkNoPaymentRequired.Enabled = vOTypesVal

      If Not vOTypesVal Then
        chkStandingOrder.Checked = False
        chkDirectDebit.Checked = False
        chkCreditCardAuthority.Checked = False
        If chkIncludeProvPaymentPlan.Checked = False Then chkNoPaymentRequired.Checked = chkForeignCurrency.Checked
      End If

      If chkProduct.Checked Or chkDonation.Checked Or chkPayment.Checked Or chkEventBooking.Checked Or chkAccomodationBooking.Checked Or chkInvoicePayment.Checked Or chkServiceBooking.Checked Or chkServiceBookingCredit.Checked Then vSalesTypes = True
      If vParamChkBox.Name = "chkCarriage" Then
        If chkCarriage.Checked = False Then
          chkConfirmCarriage.Checked = False
          chkConfirmCarriage.Enabled = False
        Else
          chkConfirmCarriage.Checked = True
          chkConfirmCarriage.Enabled = True
        End If
      End If

      If vParamChkBox.Name = "ChkPaymentMethod" And ChkPaymentMethod.Checked Then
        If chkInvoicePayment.Checked And (txtCashAccount.Text <> txtCreditCardAccount.Text Or txtCashAccount.Text <> txtDirectDebitAccount.Text Or txtCreditCardAccount.Text <> txtDebitCardAccount.Text) Then
          ChkPaymentMethod.Checked = False
          ShowInformationMessage(InformationMessages.ImCashCreditDebitDifferent, "already supports Invoice Payments") '"This application cannot support this option as it already supports Invoice Payments, and the Cash, Credit Card & Debit Card Bank Accounts are different"
        End If
      End If

      If vParamChkBox.Name = "chkInvoicePayment" And chkInvoicePayment.Checked Then
        If ChkPaymentMethod.Checked And (txtCashAccount.Text <> txtCreditCardAccount.Text Or txtCashAccount.Text <> txtDirectDebitAccount.Text Or txtCreditCardAccount.Text <> txtDebitCardAccount.Text) Then
          chkInvoicePayment.Checked = False
          ShowInformationMessage(InformationMessages.ImCashCreditDebitDifferent, "supports the selection of the payment method at the end of a transaction") '"This application cannot support this option as it supports the selection of the payment method at the end of a transaction, and the Cash, Credit Card & Debit Card Bank Accounts are different"
        End If
      End If

      If vParamChkBox.Name = "chkPaymentPlan" And chkPaymentPlan.Checked Then
        If chkForeignCurrency.Checked Then
          chkPaymentPlan.Checked = False
          ShowInformationMessage(InformationMessages.ImPaymentMethodNotSupported, "Foreign Currency") '"This payment method cannot be supported as the application supports Foreign Currency"
        End If
      End If

      If vParamChkBox.Name = "chkSundryCreditNote" And chkSundryCreditNote.Checked Then
        If chkCreditSale.Checked = False Then 'credit sales payment method not supported
          chkSundryCreditNote.Checked = False
          ShowInformationMessage(InformationMessages.ImApplicationNotSupportThisOption, "does not support Credit Sales") '"This application cannot support this option as it does not support Credit Sales"
        ElseIf GetPMTypeCount() > 1 Then 'credit sales not the only payment method supported
          chkSundryCreditNote.Checked = False
          ShowInformationMessage(InformationMessages.ImApplicationNotSupportThisOption, "supports multiple payment methods") '"This application cannot support this option as it supports multiple payment methods"
        ElseIf vOTypes > 1 Or vSalesTypes Then 'any other analysis options chosen
          chkSundryCreditNote.Checked = False
          ShowInformationMessage(InformationMessages.ImApplicationNotSupportThisOption, "supports other analysis options") '"This application cannot support this option as it supports other analysis options"
        Else  'if credit notes supported then don't allow the app to be batch led
          chkSelectBatch.Checked = False
          chkSelectBatch.Enabled = False
        End If
      ElseIf (vParamChkBox.Name = "chkChequeWithInvoice" OrElse vParamChkBox.Name = "chkCCWithInvoice") AndAlso vParamChkBox.Checked AndAlso chkCreditSale.Checked = False Then
        vParamChkBox.Checked = False
        ShowInformationMessage(InformationMessages.ImApplicationNotSupportThisOption, "does not support Credit Sales") '"This application cannot support this option as it does not support Credit Sales"
      End If

      If vParamChkBox.Name = "chkVoucher" And chkVoucher.Checked Then
        'Force to be Batch led
        chkSelectBatch.Checked = True
        chkSelectBatch.Enabled = False
      End If

      If vParamChkBox.Name = "chkForeignCurrency" And chkForeignCurrency.Checked Then
        If chkSelectBatch.Checked = False Or (chkPayment.Checked = False And chkSubscription.Checked = False) Or (chkCash.Checked = False And chkCreditCard.Checked = False) Then
          vMsg = vbTab & InformationMessages.ImBatchPayAnalysisSubAnalysisCashOrCreditNotSelected & vbCrLf '"Either the Select Batch option, the Payments analysis option, the Subscriptions analysis option, the Cash payment method or the Credit Card payment method are not selected"
        End If

        vToken = False
        If chkCheque.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If
        If vToken = False And chkPostalOrder.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If
        If vToken = False And chkDebitCard.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCreditSale.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If

        vToken = False
        If chkVoucher.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCAFCard.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If
        If vToken = False And chkGiftInKind.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
          vToken = True
        End If


        If chkSaleOrReturn.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImPaymentMethodOtherThanCashCredit & vbCrLf
        End If


        vToken = False
        If chkMembership.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkRegularDonation.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkProduct.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkDonation.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkStandingOrder.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkDirectDebit.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCreditCardAuthority.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCOvenantedMembership.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCovenantedSubscription.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkCovenantedDonation.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkEventBooking.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkExamBooking.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkAccomodationBooking.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkInvoicePayment.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkServiceBooking.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False And chkSundryCreditNote.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If
        If vToken = False AndAlso chkLoan.Checked Then
          vMsg = vMsg & vbTab & InformationMessages.ImAnalysisOtherThanPaymentsAndSubscriptions & vbCrLf
          vToken = True
        End If

        If vMsg.Length > 0 Then
          vMsg = InformationMessages.ImOptionNotSupportedResons & vbCrLf & vbCrLf & vMsg
          chkForeignCurrency.Checked = False
          ShowInformationMessage(vMsg)
          'mvDataChanged = False
        End If
      End If

      If vParamChkBox.Name = "chkServiceBookingCredit" And chkServiceBookingCredit.Checked Then
        vToken = False
        If chkCash.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit '"Payment Methods other than Credit Sales are selected"
          vToken = True
        End If
        If vToken = False And chkCheque.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkPostalOrder.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkCreditCard.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkDebitCard.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkCreditSale.Checked = False Then chkCreditSale.Checked = True


        If chkPaymentPlan.Checked And vMsg.Length = 0 Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
        End If

        vToken = False
        If chkVoucher.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkCAFCard.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If
        If vToken = False And chkGiftInKind.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
          vToken = True
        End If

        If chkSaleOrReturn.Checked Then
          vMsg = InformationMessages.ImPayMethodOtherThanCredit
        End If

        vToken = False
        If chkMembership.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct '"Analysis options other than Product Sales are selected"
          vToken = True
        End If
        If vToken = False And chkSubscription.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkRegularDonation.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkDonation.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkPayment.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkStandingOrder.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkDirectDebit.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkCreditCardAuthority.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkCOvenantedMembership.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkCovenantedSubscription.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkCovenantedDonation.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkEventBooking.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkExamBooking.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkAccomodationBooking.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkInvoicePayment.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkServiceBooking.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkSundryCreditNote.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkGoneAway.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkStatus.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkGiftAidDeclaration.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkActivity.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkSuppression.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkCancelPaymentPlan.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False And chkLegacyReceipt.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If
        If vToken = False AndAlso chkLoan.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImAnalysisOtherThanProduct
          vToken = True
        End If

        If vMsg.Length > 0 Then
          If InStr(vMsg, vbCrLf) > 0 Then
            vMsg = InformationMessages.ImOptionNotSupportedResons & vbCrLf & vbCrLf & vbTab & vMsg
          Else
            vMsg = InformationMessages.ImOptionNotSupportedBecause & " " & vMsg
          End If
          chkServiceBookingCredit.Checked = False
          ShowInformationMessage(vMsg)
        End If
      End If

      If vParamChkBox.Name = "chkConfirmDetails" And vParamChkBox.Checked Then
        If Not vOTypesVal Then
          vParamChkBox.Checked = False
          ShowInformationMessage(InformationMessages.ImNotSupportCreationOfPaymentPlans) '"This application cannot support this option as it does not support the creation of payment plans"
        End If
      End If

      Select Case vParamChkBox.Name
        Case "chkStatus", "chkActivity", "chkSuppression", "chkSuppression", "chkCancelPaymentPlan", "chkCancelGiftAidDeclaration"
          SetDefaultFieldsState(vParamChkBox)
        Case Else
          'do nothing
      End Select

      If vParamChkBox.Name = "chkConfirmSale" And vParamChkBox.Checked Then
        If chkVoucher.Checked Then vMsg = InformationMessages.ImVoucherPaymentSelected '"Voucher payment method is selected"
        If chkCAFCard.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          ''''vMsg= vMsg & (InformationMessages.ImMethodSelected,)
          vMsg = vMsg & " " & InformationMessages.ImCAFSelected '"CAF Card payment method is selected"
        End If
        If chkGiftInKind.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImGiftInKindSelected '"Gift in Kind payment method is selected"
        End If
        If chkSaleOrReturn.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImSaleOrReturnSelected '"Sale or Return payment method is selected"
        End If
        If chkIncludeProvisionalTransaction.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImIncludeProvisionalTransactions '"Include Provisional Transactions restriction is selected"
        End If
        If chkIncludeProvPaymentPlan.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImIncludeProvisionalPayment '"Include Provisional Payment Plans restriction is selected"
        End If
        If vMsg.Length > 0 Then
          If InStr(vMsg, vbCrLf) > 0 Then
            vMsg = InformationMessages.ImOptionNotSupportedResons & vbCrLf & vbCrLf & vbTab & vMsg
          Else
            vMsg = InformationMessages.ImOptionNotSupportedParam & vMsg
          End If
          chkConfirmSale.Checked = False
          ShowInformationMessage(vMsg)
        End If
      Else
        Select Case vParamChkBox.Name
          Case "chkIncludeProvisionalTransaction", "chkIncludeProvPaymentPlan"
            If vParamChkBox.Checked And chkConfirmSale.Checked Then
              vParamChkBox.Checked = False
              ShowInformationMessage(InformationMessages.ImOptionNotSupportedSaleOrReturnSelected)
            End If
        End Select
      End If

      If vParamChkBox.Name = "chkConfirmCollection" And vParamChkBox.Checked Then
        vMsg = ""
        If chkCreditCard.Checked Then vMsg = "Credit Card payment method is selected"
        If chkDebitCard.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImDebitCardSelected '"Debit Card payment method is selected"
        End If
        If chkCreditSale.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImCreditSalesSelected '"Credit Sales payment method is selected"
        End If
        If chkVoucher.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImVoucherPaymentSelected '"Voucher payment method is selected"
        End If
        If chkCAFCard.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImCAFSelected '"CAF Card payment method is selected"
        End If
        If chkGiftInKind.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImGiftInKindSelected '"Gift in Kind payment method is selected"
        End If
        If chkSaleOrReturn.Checked Then
          If vMsg.Length > 0 Then vMsg = vMsg & vbCrLf & vbTab
          vMsg = vMsg & " " & InformationMessages.ImSaleOrReturnSelected '"Sale or Return payment method is selected"
        End If
        If vMsg.Length > 0 Then
          If InStr(vMsg, vbCrLf) > 0 Then
            vMsg = InformationMessages.ImOptionNotSupportedResons & vbCrLf & vbCrLf & vbTab & vMsg
          Else
            vMsg = InformationMessages.ImOptionNotSupportedBecause & " " & vMsg
          End If
          ShowInformationMessage(vMsg)
          chkConfirmCollection.Checked = False
        End If
      End If

      If vParamChkBox.Name = "chkLinkMALToEvent" And vParamChkBox.Checked Then
        If Not (chkEventBooking.Checked Or chkProduct.Checked Or chkSundryCreditNote.Checked) Then
          ShowInformationMessage(InformationMessages.ImOptionNotSupportedParam, "either Event Booking, Product Sale or Sundry Credit Note analysis options have not been selected")
          chkLinkMALToEvent.Checked = False
        End If
      ElseIf vParamChkBox.Name = "chkEventBooking" And vParamChkBox.Checked = False Then
        If chkProduct.Checked = False Then chkLinkMALToEvent.Checked = False
      ElseIf vParamChkBox.Name = "chkProduct" And vParamChkBox.Checked = False Then
        If chkEventBooking.Checked = False Then chkLinkMALToEvent.Checked = False
      ElseIf vParamChkBox.Name = "chkSundryCreditNote" And vParamChkBox.Checked = False Then
        chkLinkMALToEvent.Checked = False
      End If

      If vParamChkBox.Name = "chkLinkMALToService" And vParamChkBox.Checked Then
        If chkServiceBooking.Checked = False Then
          ShowInformationMessage(InformationMessages.ImOptionNotSupportedParam, "the Service Booking analysis option is not selected")
          chkLinkMALToService.Checked = False
        End If

      ElseIf vParamChkBox.Name = "chkServiceBooking" And vParamChkBox.Checked = False Then
        chkLinkMALToService.Checked = False
      End If

      If vParamChkBox.Name = "chkBankDetails" And mvAlbacsBankDetails Then
        If vParamChkBox.Checked And mvApplicationType <> "CLREC" Then
          cboAlbacsBankDetails.Enabled = True
        Else
          cboAlbacsBankDetails.SelectedIndex = 3
          cboAlbacsBankDetails.Enabled = False
        End If
      End If

      If vParamChkBox.Name = "chkLinkAnalysisLines" And vParamChkBox.Checked Then
        If chkDonation.Checked = False Then
          If chkGiftInKind.Checked = False Or chkProduct.Checked = False Then
            vMsg = ""
            vMsg = InformationMessages.ImOptionNotSupportedResons & vbCrLf & vbCrLf & vbTab
            vMsg = vMsg & vbCrLf & InformationMessages.ImDonationAnalysisSelected
            vMsg = vMsg & vbCrLf & InformationMessages.ImGiftInKindAndProductSalesSelected
            ShowInformationMessage(vMsg)
            chkLinkAnalysisLines.Checked = False
          End If
        End If
      ElseIf vParamChkBox.Name = "chkDonation" And vParamChkBox.Checked = False Then
        chkLinkAnalysisLines.Checked = False
      ElseIf (vParamChkBox.Name = "chkGiftInKind" Or vParamChkBox.Name = "chkProduct") And vParamChkBox.Checked = False Then
        chkLinkAnalysisLines.Checked = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function SaveTraderApp() As String
    If ValidateTrader() = False Then Return ""

    Dim vList As New ParameterList(True)
    If mvPFAppNumber > 0 Then
      If Not ConfirmUpdate() Then Return ""
      vList("FpApplicationNumber") = mvPFAppNumber.ToString()
    Else
      If Not ConfirmInsert() Then Return ""
    End If

    vList("FpApplicationDesc") = txtAppDesc.Text
    vList("FpApplicationType") = txtType.Text
    vList("BatchApplication") = CBoolYN(chkSelectBatch.Checked)
    vList("PmCash") = CBoolYN(chkCash.Checked)
    vList("PmCheque") = CBoolYN(chkCheque.Checked)
    vList("PmPostalOrder") = CBoolYN(chkPostalOrder.Checked)
    vList("PmCreditCard") = CBoolYN(chkCreditCard.Checked)
    vList("PmDebitCard") = CBoolYN(chkDebitCard.Checked)
    vList("PmCredit") = CBoolYN(chkCreditSale.Checked)
    vList("ChequeWithInvoice") = CBoolYN(chkChequeWithInvoice.Checked)
    vList("CcWithInvoice") = CBoolYN(chkCCWithInvoice.Checked)
    vList("BankDetails") = CBoolYN(chkBankDetails.Checked)
    vList("TransactionComments") = CBoolYN(chkTransactionComment.Checked)
    vList("Memberships") = CBoolYN(chkMembership.Checked)
    vList("Subscriptions") = CBoolYN(chkSubscription.Checked)
    vList("DonationsRegular") = CBoolYN(chkRegularDonation.Checked)
    vList("ProductSales") = CBoolYN(chkProduct.Checked)
    vList("DonationsOneOff") = CBoolYN(chkDonation.Checked)
    vList("Payments") = CBoolYN(chkPayment.Checked)
    vList("StandingOrders") = CBoolYN(chkStandingOrder.Checked)
    vList("DirectDebits") = CBoolYN(chkDirectDebit.Checked)
    vList("CreditCardAuthorities") = CBoolYN(chkCreditCardAuthority.Checked)
    vList("CovenantMembership") = CBoolYN(chkCOvenantedMembership.Checked)
    vList("CovenantSubscription") = CBoolYN(chkCovenantedSubscription.Checked)
    vList("CovenantDonationRegular") = CBoolYN(chkCovenantedDonation.Checked)
    vList("EventBooking") = CBoolYN(chkEventBooking.Checked)

    vList("ExamBooking") = CBoolYN(chkExamBooking.Checked)

    vList("AccomodationBooking") = CBoolYN(chkAccomodationBooking.Checked)
    vList("InvoicePayments") = CBoolYN(chkInvoicePayment.Checked)
    vList("Product") = txtProduct.Text
    vList("Rate") = txtRate.Text
    vList("Source") = txtSource.Text
    vList("CaBankAccount") = txtCashAccount.Text
    vList("CcBankAccount") = txtCreditCardAccount.Text
    vList("DcBankAccount") = txtDebitCardAccount.Text
    vList("CsBankAccount") = txtCreditSaleAccount.Text
    vList("SoBankAccount") = txtStandingOrderAccount.Text
    vList("DdBankAccount") = txtDirectDebitAccount.Text
    vList("CcaBankAccount") = txtCCAAccount.Text
    vList("ReadOnly") = "N"
    vList("ShowTransactionReference") = CBoolYN(chkShowReference.Checked)
    vList("ConfirmDefaultProduct") = CBoolYN(chkConfirmProduct.Checked)
    vList("ConfirmAnalysis") = CBoolYN(chkConfirmAnalysis.Checked)
    vList("Carriage") = CBoolYN(chkCarriage.Checked)
    vList("ConfirmCarriage") = CBoolYN(chkConfirmCarriage.Checked)
    vList("CarriageProduct") = txtCarriageProduct.Text
    vList("CarriageRate") = txtCarriageRate.Text
    vList("CarriagePercentage") = txtPercentage.Text
    vList("AnalysisComments") = CBoolYN(chkAnalysisComments.Checked)
    vList("NonPaidPaymentPlans") = CBoolYN(chkNoPaymentRequired.Checked)
    vList("PayMethodsAtEnd") = CBoolYN(ChkPaymentMethod.Checked)
    vList("ServiceBookings") = CBoolYN(chkServiceBooking.Checked)
    vList("DefaultSalesContact") = txtSalesPerson.Text
    vList("PayPlanPayMethod") = CBoolYN(chkPaymentPlan.Checked)
    vList("SundryCreditNotes") = CBoolYN(chkSundryCreditNote.Checked)
    vList("ForeignCurrency") = CBoolYN(chkForeignCurrency.Checked)
    vList("InvoiceDocument") = txtInvoiceDoc.Text
    vList("ReceiptDocument") = txtReceiptDoc.Text
    vList("PaymentPlanDocument") = txtPayPlanDoc.Text
    vList("ChangeMembership") = CBoolYN(chkMembershipType.Checked)
    vList("MembersOnly") = CBoolYN(chkMembersOnly.Checked)
    vList("SalesGroup") = txtSalesGroup.Text
    vList("CreditStatementDocument") = txtCreditStmtDoc.Text
    vList("AutoSetAmount") = CBoolYN(chkAutoSetAmount.Checked)
    vList("ServiceBookingCredits") = CBoolYN(chkServiceBookingCredit.Checked)
    vList("LegacyReceipts") = CBoolYN(chkLegacyReceipt.Checked)
    vList("SetGoneAway") = CBoolYN(chkGoneAway.Checked)
    vList("SetStatus") = CBoolYN(chkStatus.Checked)
    vList("Status") = txtStatus.Text
    vList("GiftAidDeclaration") = CBoolYN(chkGiftAidDeclaration.Checked)
    vList("AddActivity") = CBoolYN(chkActivity.Checked)
    vList("ActivityGroup") = txtActivity.Text
    vList("AddSuppression") = CBoolYN(chkSuppression.Checked)
    vList("CancelPaymentPlan") = CBoolYN(chkCancelPaymentPlan.Checked)
    vList("CancellationReason") = txtCancelGiftAidDeclaration.Text
    vList("MailingSuppression") = txtSuppression.Text
    vList("ConfirmDetails") = CBoolYN(chkConfirmDetails.Checked)
    vList("IncludeProvisionalTrans") = CBoolYN(chkIncludeProvisionalTransaction.Checked)
    vList("IncludeConfirmedTrans") = CBoolYN(chkIncludeConfirmedTransaction.Checked)
    vList("PmVoucher") = CBoolYN(chkVoucher.Checked)
    vList("PmCafCard") = CBoolYN(chkCAFCard.Checked)
    vList("PmGiftInKind") = CBoolYN(chkGiftInKind.Checked)
    vList("CvBankAccount") = txtCAFAndVoucherAccount.Text
    vList("DonationProduct") = txtDonationProduct.Text
    vList("DonationRate") = txtDonationRate.Text
    vList("PayrollGiving") = CBoolYN(chkPayrollGiving.Checked)
    vList("BatchCategory") = txtBatchCategory.Text
    vList("SourceFromLastMailing") = CBoolYN(chkDefaultSourceFromLastMailing.Checked)
    vList("MailingCodeMandatory") = CBoolYN(chkForceMailingCode.Checked)
    vList("DistributionCodeMandatory") = CBoolYN(chkForceDistributionCode.Checked)
    vList("SalesContactMandatory") = CBoolYN(chkSalesContactMandatory.Checked)
    vList("BypassMailingParagraphs") = CBoolYN(chkBypass.Checked)
    vList("PmSaleOrReturn") = CBoolYN(chkSaleOrReturn.Checked)
    vList("NonFinancialBatch") = CBoolYN(chkNonFinancialBatch.Checked)
    vList("AddressMaintenance") = CBoolYN(chkAddressMaintenance.Checked)
    vList("AutoPaymentMaintenance") = CBoolYN(chkAutoPaymentMaintenance.Checked)
    vList("GiftAidCancellation") = CBoolYN(chkCancelGiftAidDeclaration.Checked)
    vList("ProvisionalPaymentPlan") = CBoolYN(chkIncludeProvPaymentPlan.Checked)
    vList("DefaultMemberBranch") = txtBranch.Text
    vList("ProvisionalCashDocument") = txtProvCashDoc.Text
    vList("PrefulfilledIncentives") = CBoolYN(chkPrefulfilledIncentives.Checked)
    vList("ContactAlerts") = CBoolYN(chkContactAlerts.Checked)
    vList("DisplayScheduledPayments") = CBoolYN(chkDisplayScheduledPayment.Checked)
    vList("ConfirmSrTransactions") = CBoolYN(chkSaleOrReturn.Checked)
    vList("OnlineCcAuthorisation") = CBoolYN(chkOnlineAuth.Checked)
    vList("RequireCcAuthorisation") = CBoolYN(chkRequireAuthorisation.Checked)
    vList("PpConversionInclMaintenance") = CBoolYN(chkMaintainPaymentPlan.Checked)
    vList("LinkToCommunication") = txtLinkToCommunication.Text
    vList("CollectionPayments") = CBoolYN(chkConfirmCollection.Checked)
    vList("BatchAnalysisCode") = txtBatchAnalysisCode.Text
    vList("EventMultipleAnalysis") = CBoolYN(chkLinkMALToEvent.Checked)
    vList("TransactionOrigin") = txtTransactionOrigin.Text
    vList("ServiceBookingAnalysis") = CBoolYN(chkLinkMALToService.Checked)
    vList("AlbacsBankDetails") = DirectCast(cboAlbacsBankDetails.SelectedItem, LookupItem).LookupCode.ToString
    vList("LinkToFundraisingPayments") = CBoolYN(chkLinkAnalysisLines.Checked)
    vList("InvoicePrintPreviewDefault") = CBoolYN(chkInvoicePrintPreview.Checked)
    vList("Loans") = CBoolYN(chkLoan.Checked)
    vList("AutoCreateCreditCustomer") = CBoolYN(chkAutoCreateCreditCustomer.Checked)
    vList("UnpostedBatchMsgInPrint") = CBoolYN(chkUnpostedBatchMsgInPrint.Checked)
    vList("DateRangeMsgInPrint") = CBoolYN(chkDateRangeMsgInPrint.Checked)
    vList("CreditCategory") = txtCreditCategory.Text
    vList("ExamSessionCode") = txtExamSession.Text
    vList("ExamUnitCode") = txtExamUnit.Text
    vList("InvoicePrintUnpostedBatches") = CBoolYN(chkInvoicePrintUnpostedBatches.Checked)
    vList("AutoGiftAidDeclaration") = CBoolYN(autoGiftAidDeclaration.Checked)
    vList("AutoGiftAidMethod") = If(gadMethodOral.Checked,
                                    "O",
                                    If(gadMethodWritten.Checked,
                                       "W",
                                       "E"))
    vList("AutoGiftAidSource") = gadSource.Text
    vList("MerchantRetailNumber") = txtMerchantRetailNumber.Text
    Dim vReturnList As ParameterList = DataHelper.SaveTraderApplication(vList)
    vList = New ParameterList(True)
    vList("FpApplicationNumber") = vReturnList("FpApplication").ToString
    If mvDataSet IsNot Nothing AndAlso mvDataSet.Tables("DataRow") IsNot Nothing Then
      For vIncr As Integer = 0 To mvDataSet.Tables("DataRow").Rows.Count - 1
        If vIncr = 0 Then vList("Flag") = "Y" Else vList.Remove("Flag")
        vList("BankAccount") = mvDataSet.Tables("DataRow").Rows(vIncr).Item("BankAccount").ToString
        vList("CurrencyCode") = mvDataSet.Tables("DataRow").Rows(vIncr).Item("CurrencyCode").ToString
        vList("BatchType") = mvDataSet.Tables("DataRow").Rows(vIncr).Item("BatchType").ToString
        DataHelper.SaveTraderApplicationBank(vList)
      Next
    End If
    Return vReturnList("FpApplication").ToString
  End Function

  Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Try
      If SaveTraderApp().Length > 0 Then
        Me.Close()
      End If
    Catch vException As CareException
      Select Case vException.ErrorNumber
        Case CareException.ErrorNumbers.enAppCantUseDefaultProduct
          ShowInformationMessage(vException.Message)
        Case Else
          DataHelper.HandleException(vException)
      End Select
    End Try
  End Sub


  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    Try
      Dim vRow As DataRow
      If mvDataSet.Tables("DataRow") Is Nothing Then
        mvDataSet.Tables.Add("DataRow")
        mvDataSet.Tables("DataRow").Columns.Add("CurrencyCode", Type.GetType("System.String"))
        mvDataSet.Tables("DataRow").Columns.Add("BatchType", Type.GetType("System.String"))
        mvDataSet.Tables("DataRow").Columns.Add("BankAccountDesc", Type.GetType("System.String"))
        mvDataSet.Tables("DataRow").Columns.Add("BankAccount", Type.GetType("System.String"))
      End If
      vRow = mvDataSet.Tables("DataRow").NewRow()
      vRow.Item("CurrencyCode") = txtCurrency.Text
      vRow.Item("BatchType") = txtBatchType.Text
      vRow.Item("BankAccountDesc") = txtBankAccount.ComboBox.Text
      vRow.Item("BankAccount") = txtBankAccount.Text
      mvDataSet.Tables("DataRow").Rows.Add(vRow)
      dgr.Populate(mvDataSet)
      cmdAdd.Enabled = IsAddValid()
      cmdRemove.Enabled = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
    Try
      If dgr.RowCount > 0 Then
        dgr.DeleteRow(dgr.ActiveRow)
        If dgr.RowCount = 0 Then
          cmdRemove.Enabled = False
          If IsAddValid() Then
            cmdAdd.Enabled = True
          End If
        End If
        dgr.SelectRow(-1)
        cmdRemove.Enabled = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function ValidateTrader() As Boolean
    'Return True if all valid or False if not
    Dim vValid As Boolean = True
    If txtAppDesc.Text.Trim.Length = 0 Then
      vValid = False
      erp.SetError(txtAppDesc, InformationMessages.ImFieldMustNotBeBlank)
      tabMain.SelectedTab = tbpGeneral
    Else
      erp.SetError(txtAppDesc, "")
    End If
    If vValid Then
      mvFirstErrorControl = Nothing
      vValid = ValidateAllControls(Me.tabMain)
      If vValid = False Then
        Dim vTabPage As TabPage = TryCast(mvFirstErrorControl.Parent.Parent.Parent.Parent, TabPage)
        Dim vTabControl As System.Windows.Forms.TabControl = Nothing
        If vTabPage IsNot Nothing Then vTabControl = TryCast(vTabPage.Parent, System.Windows.Forms.TabControl)
        If vTabControl IsNot Nothing Then vTabControl.SelectedTab = vTabPage
        vTabPage = TryCast(mvFirstErrorControl.Parent.Parent, TabPage)
        If vTabPage IsNot Nothing Then vTabControl = TryCast(vTabPage.Parent, System.Windows.Forms.TabControl)
        If vTabControl IsNot Nothing Then vTabControl.SelectedTab = vTabPage
        mvFirstErrorControl.Focus()
        Return False
      End If
    End If

    If vValid And (mvApplicationType <> "GAYEP" And mvApplicationType <> "POTPG") Then
      If (chkCash.Checked = False And chkCheque.Checked = False And chkPostalOrder.Checked = False _
      And chkCreditCard.Checked = False And chkDebitCard.Checked = False And chkCreditSale.Checked = False _
      And chkVoucher.Checked = False And chkCAFCard.Checked = False _
      And chkGiftInKind.Checked = False And chkSaleOrReturn.Checked = False) Then
        'No actual payment methods defined
        If chkPaymentPlan.Checked Then
          If chkPayment.Checked Or chkInvoicePayment.Checked Or chkSundryCreditNote.Checked Then
            vValid = False
          End If
        ElseIf (chkProduct.Checked Or chkDonation.Checked Or chkPayment.Checked _
        Or chkEventBooking.Checked Or chkExamBooking.Checked Or chkAccomodationBooking.Checked Or chkInvoicePayment.Checked Or chkServiceBooking.Checked _
        Or chkSundryCreditNote.Checked Or chkServiceBookingCredit.Checked _
        Or chkLegacyReceipt.Checked) Then
          'The above items require a payment method and none is selected
          vValid = False
        ElseIf (chkMembership.Checked Or chkSubscription.Checked Or chkRegularDonation.Checked _
        Or chkCOvenantedMembership.Checked Or chkCovenantedSubscription.Checked _
        Or chkCovenantedDonation.Checked Or chkLoan.Checked) _
        And (chkStandingOrder.Checked = False And chkDirectDebit.Checked = False _
        And chkCreditCardAuthority.Checked = False And chkNoPaymentRequired.Checked = False Or chkConfirmCollection.Checked) Then
          'The above items require a payment method and none is selected
          vValid = False
        End If
      End If
      If Not vValid Then
        ShowInformationMessage(InformationMessages.ImNoPaymentMethods) '"No payment methods defined"
        tabMain.SelectedTab = tbpGeneral
      End If
    End If

    Dim vToken As Boolean
    If (mvApplicationType <> "GAYEP" And mvApplicationType <> "POTPG") Then
      If vValid Then
        vValid = False
        vToken = False
        If chkMembership.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkSubscription.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkRegularDonation.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkProduct.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkDonation.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkPayment.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkStandingOrder.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkDirectDebit.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkCreditCardAuthority.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkCOvenantedMembership.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkCovenantedSubscription.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkCovenantedDonation.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkEventBooking.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkExamBooking.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkAccomodationBooking.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkInvoicePayment.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkServiceBooking.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkSundryCreditNote.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkServiceBookingCredit.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkGoneAway.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkStatus.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkGiftAidDeclaration.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkActivity.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkSuppression.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkCancelPaymentPlan.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkLegacyReceipt.Checked Then
          vValid = True
          vToken = True
        End If
        If vToken = False And chkLoan.Checked Then
          vValid = True
          vToken = True
        End If

        If Not vValid Then vValid = chkMembershipType.Checked = True
        If Not vValid Then vValid = chkAddressMaintenance.Checked = True
        If Not vValid Then vValid = chkCancelGiftAidDeclaration.Checked = True
        If Not vValid Then vValid = chkAutoPaymentMaintenance.Checked = True
        If Not vValid Then vValid = chkPayrollGiving.Checked = True
        If Not vValid Then vValid = chkConfirmProduct.Checked = True
        If Not vValid Then
          ShowInformationMessage(InformationMessages.ImNoAnalysisDefined) '"No analysis items defined"
        End If
      End If
    End If


    If vValid Then
      If chkSundryCreditNote.Checked Then
        If txtCreditSaleAccount.Text = "" Then
          chkSundryCreditNote.Checked = False
          vValid = False
          ShowInformationMessage(InformationMessages.ImCreditSalesBankAccountNotSet) '"This application cannot support the Credit Notes option as the Credit Sales Bank Account has not been set."
        Else
          Dim vList As New ParameterList(True)
          vList("BankAccount") = txtCreditSaleAccount.Text
          If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctBankAccountCompanies, vList) = 0 Then
            chkSundryCreditNote.Checked = False
            vValid = False
            ShowInformationMessage(InformationMessages.ImNotSupportCreditNotes) '"This application cannot support the Credit Notes option as the company that owns the Credit Sales Bank Account does not have the Sundry Credit control product set up."
          End If
        End If
      End If

    End If

    'ensure that default product and default sales group are linked
    If vValid Then
      If (txtProduct.Text).Length > 0 And (txtSalesGroup.Text).Length > 0 Then
        Dim vSalesGroup As String
        Dim vList As New ParameterList(True)
        vList("Product") = txtProduct.TextBox.Text
        Dim vProductTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
        If vProductTable.Rows.Count > 0 Then
          vSalesGroup = vProductTable.Rows(0).Item("SalesGroup").ToString
          If vSalesGroup.Length > 0 Then
            If vSalesGroup <> txtSalesGroup.TextBox.Text Then
              vValid = False
              ShowInformationMessage(InformationMessages.ImProductNotInSalesGroup, vSalesGroup) 'The specified default Product is not in the specified default Sales Group - it is in Sales Group %s
            End If
          End If
        End If
      End If
    End If

    If vValid Then
      If chkActivity.Checked Then
        If (txtActivity.Text).Length = 0 Then
          vValid = False
          ShowInformationMessage(InformationMessages.ImActivityGroupMustBeSpecified) '"An Activity Group must be specified when the application is to support the adding of Activities"
        End If
      End If
    End If
    If vValid Then
      If chkSuppression.Checked Then
        If (txtSuppression.Text).Length = 0 Then
          vValid = False
          ShowInformationMessage(InformationMessages.ImSuppressionGroupMustBeSpecified) '"A Suppression Group must be specified when the application supports the adding of Suppressions"
        End If
      End If
    End If

    If vValid Then
      If chkOnlineAuth.Checked AndAlso chkCAFCard.Checked Then
        vValid = False
        ShowInformationMessage(InformationMessages.ImTransactionPayCAFCCInvalid) 'CAF Card payment method invalid when using On-Line Credit Card Authorisation
      End If
    End If
    If vValid Then
      If chkOnlineAuth.Checked = False AndAlso (chkDebitCard.Checked OrElse chkCreditCard.Checked OrElse chkCCWithInvoice.Checked) Then
        vValid = False
        ShowInformationMessage(InformationMessages.ImTransactionPayCardInvalid) 'Card payment method invalid when not using On-Line Credit Card Authorisation
      End If
    End If

    If vValid Then
      If chkConfirmDetails.Checked Then
        Dim vCount As Long

        If chkMembership.Checked Then vCount = vCount + 1
        If chkSubscription.Checked Then vCount = vCount + 1
        If chkRegularDonation.Checked Then vCount = vCount + 1
        If chkLoan.Checked Then vCount += 1

        If chkCOvenantedMembership.Checked Then vCount = vCount + 1
        If chkCovenantedSubscription.Checked Then vCount = vCount + 1
        If chkCovenantedDonation.Checked Then vCount = vCount + 1

        If vCount = 0 Then
          chkConfirmDetails.Checked = False
          vValid = False
          ShowInformationMessage(InformationMessages.ImConfirmDetailsOptionNotSupported) '"The Confirm Details option cannot be supported as the application does not support payment plans"
        End If
      End If
    End If
    If vValid Then ValidateTextLookups(Me, vValid)
    Return vValid
  End Function

  Private Sub ValidateTextLookups(ByVal pControl As Control, ByRef pValid As Boolean)
    For Each vControl As Control In pControl.Controls
      If TypeOf (vControl) Is TextLookupBox Then
        pValid = DirectCast(vControl, TextLookupBox).IsValid
        If pValid Then
          erp.SetError(vControl, "")
        Else
          erp.SetError(vControl, GetInformationMessage(InformationMessages.ImInvalidValue))
          Dim vParent As Control = vControl.Parent
          While vParent IsNot Nothing
            If TypeOf (vParent) Is TabPage Then
              Dim vTabpage As TabPage = DirectCast(vParent, TabPage)
              If TypeOf (vTabpage.Parent) Is System.Windows.Forms.TabControl Then
                DirectCast(vTabpage.Parent, System.Windows.Forms.TabControl).SelectedTab = vTabpage
              End If
            End If
            vParent = vParent.Parent
          End While
        End If
      Else
        ValidateTextLookups(vControl, pValid)
      End If
      If pValid = False Then Exit For
    Next
  End Sub

  Private Sub cmdDesign_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDesign.Click
    Try
      Dim vAppNumber As String = ""
      If mvPFAppNumber = 0 Then
        vAppNumber = SaveTraderApp()
        If vAppNumber.Length = 0 Then Exit Sub
        mvPFAppNumber = IntegerValue(vAppNumber)
      End If
      Dim vForm As New frmDesignTraderApp(mvPFAppNumber)
      vForm.ShowDialog()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try

  End Sub

  Private Sub cmdNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      mvPFAppNumber = 0
      mvTable = Nothing
      mvDataSet = Nothing
      mvEditMode = False
      txtAppDesc.Text = ""
      SetTraderValues()
      If dgr.DataRowCount > 0 Then
        dgr.ClearDataRows()
        dgr.DeleteRow(0)
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    Try
      If dgr.RowCount > 0 Then
        Dim vCurrency As String = dgr.GetValue(pRow, 0)
        Dim vBatchType As String = dgr.GetValue(pRow, 1)
        Dim vBankAccount As String = mvDataSet.Tables("DataRow").Rows(pDataRow).Item("BankAccount").ToString
        txtCurrency.TextBox.Text = vCurrency
        txtBatchType.TextBox.Text = vBatchType
        txtBankAccount.TextBox.Text = vBankAccount
        cmdRemove.Enabled = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Function IsAddValid() As Boolean
    Dim vAddEnabled As Boolean = False
    Dim vCurrency As String = txtCurrency.TextBox.Text
    Dim vBatchType As String = txtBatchType.TextBox.Text
    Dim vBankAccount As String = txtBankAccount.TextBox.Text
    If vCurrency.Length > 0 AndAlso txtCurrency.IsValid AndAlso _
      vBatchType.Length > 0 AndAlso txtBatchType.IsValid AndAlso _
      vBankAccount.Length > 0 AndAlso txtBankAccount.IsValid Then
      vAddEnabled = True
    End If
    Dim vIsAddValid As Boolean = True
    If dgr.RowCount > 0 Then
      For vRowCounter As Integer = 0 To dgr.RowCount - 1
        If vCurrency = dgr.GetValue(vRowCounter, 0) AndAlso vBatchType = dgr.GetValue(vRowCounter, 1) Then
          vIsAddValid = False
          Exit For
        Else
          vIsAddValid = True
        End If
      Next
    End If
    If vAddEnabled And vIsAddValid Then
      Return True
    Else
      Return False
    End If
  End Function

  Private Function ValidateControl(ByVal pControl As System.Windows.Forms.Control, ByVal pPanelItem As PanelItem, ByVal pValue As String) As Boolean
    Dim vValid As Boolean = True

    If pPanelItem.ValidationError Then
      vValid = False
    Else
      erp.SetError(pControl, "")                                'Clear any errors
    End If
    If TypeOf pControl Is TextLookupBox AndAlso DirectCast(pControl, TextLookupBox).IsValid = False Then
      erp.SetError(pControl, GetInformationMessage(InformationMessages.ImInvalidValue))
      vValid = False
      'ElseIf TypeOf pControl Is MaskedTextBox AndAlso pValue.Length > 0 AndAlso DirectCast(pControl, MaskedTextBox).MaskCompleted = False Then
      '  erp.SetError(pControl, GetInformationMessage(InformationMessages.ImInvalidValue))
      '  vValid = False
    End If
    If pControl.Enabled = True AndAlso vValid AndAlso pPanelItem.Mandatory AndAlso pValue.Length = 0 Then
      If TypeOf pControl Is RadioButton Then
        'Assume one of the radiobuttons is set
      Else
        erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldMandatory))
        vValid = False
      End If
    End If
    If pValue.Length > 0 Then
      If vValid AndAlso pPanelItem.MinimumValue.Length > 0 Then
        If pPanelItem.FieldType = PanelItem.FieldTypes.cftCharacter Then
          If pValue < pPanelItem.MinimumValue Then
            erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldNotLessThan, pPanelItem.MinimumValue))
            vValid = False
          End If
        Else
          If CDbl(pValue) < CDbl(pPanelItem.MinimumValue) Then
            Dim vErrorMsg As String = InformationMessages.ImFieldNotLessThan
            If pPanelItem.MinimumValueError.Length > 0 Then vErrorMsg = pPanelItem.MinimumValueError
            erp.SetError(pControl, GetInformationMessage(vErrorMsg, pPanelItem.MinimumValue))
            vValid = False
          End If
        End If
      End If
      If vValid AndAlso pPanelItem.MaximumValue.Length > 0 Then
        If pPanelItem.FieldType = PanelItem.FieldTypes.cftCharacter Then
          If pValue > pPanelItem.MaximumValue Then
            erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldNotGreaterThan, pPanelItem.MinimumValue))
            vValid = False
          End If
        Else
          If CDbl(pValue) > CDbl(pPanelItem.MaximumValue) Then
            Dim vErrorMsg As String = InformationMessages.ImFieldNotGreaterThan
            If pPanelItem.MaximumValueError.Length > 0 Then vErrorMsg = pPanelItem.MaximumValueError
            erp.SetError(pControl, GetInformationMessage(vErrorMsg, pPanelItem.MaximumValue))
            vValid = False
          End If
        End If
      End If
      If vValid AndAlso pPanelItem.Pattern.Length > 0 Then
        If Not (pPanelItem.AttributeName = "line_type" AndAlso pPanelItem.ControlType = PanelItem.ControlTypes.ctCheckBox) Then
          Dim vReg As New System.Text.RegularExpressions.Regex(pPanelItem.Pattern)
          If Not vReg.IsMatch(pValue) Then
            erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldMatchPattern, pPanelItem.Pattern))
          End If
        End If
      End If
      If vValid AndAlso pPanelItem.MinimumLength > 0 Then
        If pValue.Trim.Length < pPanelItem.MinimumLength Then
          erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldMinimumLength, pPanelItem.MinimumLength.ToString))
        End If
      End If
    End If
    If Not vValid Then
      If mvFirstErrorControl Is Nothing Then
        mvFirstErrorControl = pControl
      End If
    End If
    Return vValid
  End Function

  Private Function ValidateAllControls(ByVal pControl As Control) As Boolean
    Dim vValid As Boolean = True
    For Each vControl As Control In pControl.Controls
      If vControl.Controls.Count > 0 AndAlso Not TypeOf (vControl) Is TextLookupBox AndAlso Not TypeOf (vControl) Is TopicDataSheet AndAlso Not TypeOf (vControl) Is ReportBox AndAlso Not TypeOf (vControl) Is ColorSelector Then
        If ValidateAllControls(vControl) = False Then vValid = False
      ElseIf TypeOf vControl Is TextLookupBox Then
        Dim vTextLookupBox As TextLookupBox = DirectCast(vControl, TextLookupBox)
        If ValidateControl(vTextLookupBox, DirectCast(vTextLookupBox.Tag, PanelItem), vTextLookupBox.Text) = False Then vValid = False
      ElseIf TypeOf (vControl) Is TextBox Then
        If ValidateTextBox(CType(vControl, TextBox)) = False Then vValid = False
      End If
    Next
    Return vValid
  End Function

  Private Function ValidateTextBox(ByVal pControl As TextBox) As Boolean
    Dim vValid As Boolean = True
    Select Case pControl.Name
      Case "txtPercentage"
        erp.SetError(pControl, String.Empty)
        If pControl.Text.Length > 0 Then
          Dim vDouble As Double
          If Double.TryParse(pControl.Text, vDouble) Then
            If vDouble.CompareTo(0) < 0 Then
              erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldNotLessThan, "0.00"))
              vValid = False
            ElseIf vDouble.CompareTo(999.99) > 0 Then
              erp.SetError(pControl, GetInformationMessage(InformationMessages.ImFieldNotGreaterThan, "999.99"))
              vValid = False
            End If
          Else
            erp.SetError(pControl, InformationMessages.ImTraderCarriagePercentageInvalid)   'Carriage Percentage is not a valid percentage
            vValid = False
          End If
          If vValid = False Then
            If mvFirstErrorControl Is Nothing Then mvFirstErrorControl = pControl
          End If
        End If
    End Select

    Return vValid

  End Function

  Private Sub autoGiftAidDeclaration_CheckedChanged(sender As Object, e As EventArgs) Handles autoGiftAidDeclaration.CheckedChanged
    Me.newDeclarationGroup.Enabled = Me.autoGiftAidDeclaration.Checked
  End Sub

  Private Sub cmdAddAlert_Click(sender As Object, e As EventArgs) Handles cmdAddAlert.Click
    'Add a Contact Alert and link it to the Trader App.
    Try
      Dim vDefaults As New ParameterList
      vDefaults.Add("ContactAlertType", "F")
      vDefaults.Add("FpApplicationNumber", mvPFAppNumber.ToString())

      Dim vTableName As String = "contact_alerts"
      Dim vfrmTableEntry As New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vTableName, vDefaults, Nothing, False, True)
      vfrmTableEntry.Text = ControlText.FrmAddTo & Utilities.ProperName(vTableName)
      If vfrmTableEntry.ShowDialog() = DialogResult.OK Then
        'Add link
        If vDefaults.ContainsKey("ContactAlert") Then
          'Saving the ContactAlerts record changes vDefaults to be all the data added!!
          AddContactAlertLink(vDefaults("ContactAlert"))
          GetLinkedAlerts()
        End If
      End If

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdAddAlertLink_Click(sender As Object, e As EventArgs) Handles cmdAddAlertLink.Click
    'Link a Contact Alert to this Trader App.
    Try
      Dim vList As New ParameterList(True, True)
      vList("ContactAlertType") = "F"

      Dim vSF As New frmSimpleFinder()
      vSF.RestrictionsList = vList
      vSF.Init(CareNetServices.XMLLookupDataTypes.xldtAvailableAlerts, True)
      If vSF.ShowDialog() = DialogResult.OK Then
        AddContactAlertLink(vSF.ResultValue)
        GetLinkedAlerts()
      End If
    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(vCareEX.Message)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdDeleteAlertLink_Click(sender As Object, e As EventArgs) Handles cmdDeleteAlertLink.Click
    'Delete the link between a Contact Alert and this Trader App.
    Try
      Dim vRow As Integer = dgrAlerts.CurrentDataRow
      If vRow >= 0 Then
        If ConfirmDelete() Then
          Dim vList As New ParameterList(True, True)
          vList("ContactAlertLinkNumber") = dgrAlerts.GetValue(vRow, dgrAlerts.GetColumn("ContactAlertLinkNumber"))
          DataHelper.DeleteContactAlertLink(vList)
          'Update grid
          GetLinkedAlerts()
        End If
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub dgrAlerts_RowSelected(sender As Object, pRow As Integer, pDataRow As Integer) 
    cmdDeleteAlertLink.Enabled = (dgrAlerts.DataRowCount > 0)
  End Sub

  ''' <summary>Get the Contact Alerts that are linked to this Trader Application</summary>
  Private Sub GetLinkedAlerts()
    Dim vDS As DataSet = DataHelper.GetTraderAlerts(mvPFAppNumber)
    If vDS IsNot Nothing Then
      dgrAlerts.Populate(vDS)
      cmdDeleteAlertLink.Enabled = (mvPFAppNumber.CompareTo(0) > 0 AndAlso dgrAlerts.DataRowCount > 0)
    End If
  End Sub

  ''' <summary>Add a ContactAlertLink record to link this Trader Application with a Contact Alert</summary>
  Private Sub AddContactAlertLink(ByVal pContactAlert As String)
    'Add link
    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("FpApplicationNumber") = mvPFAppNumber
    vList.Add("ContactAlert", pContactAlert)
    DataHelper.AddContactAlertLink(vList)
  End Sub

  Private Sub dgrAlerts_CanCustomise(Sender As Object, pGridName As String) Handles dgrAlerts.CanCustomise
    GetLinkedAlerts()
  End Sub

End Class