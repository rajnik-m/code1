Public Class frmIncentiveMaintenance


  Private Enum imIncentiveTypes
    imitContactFulFilled = 1
    imitContactUnFulFilled
    imitPayPlanFulFilled
    imitPayPlanUnFulFilled
  End Enum

  Private mvIncentiveType As imIncentiveTypes
  Private mvOrderSource As String = String.Empty               'Order Source, used for reset
  Private mvOrderReason As String = String.Empty               'Order ReasonForDespatch, used for reset
  Private mvContact As Integer                 'Selected contact
  Private mvPayPlan As Integer                 'Selected PayPlan
  Private mvExtraSource As String = String.Empty               'Any extra Source
  Private mvExtraReason As String = String.Empty               'Any extra Reason for despatch
  Private mvSource As String = String.Empty                'Source for selection
  Private mvPP As PaymentPlanInfo           'Current PaymentPlan for paymentplan incentives
  Private mvReasonForDesp As String = String.Empty               'Reason for Despatch for selection
  Private mvCCReason As String = String.Empty               'CC_Reason from Financial Controls
  Private mvDDReason As String = String.Empty               'DD_Reason from Financial Controls
  Private mvSOReason As String = String.Empty               'SO_reason from financial controls
  Private mvOReason As String = String.Empty               'O_Reason from financial_controls
  Private mvMailing As String = String.Empty               'Mailing code for selection
  Private mvClose As Boolean = True                ' Close Form

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    Try
      SetControlTheme()

      Dim vList As New ParameterList()
      vList("ValidationTable") = "orders"
      vList("ValidationAttribute") = "order_number"
      InitTextLookupBox(txtLookupPayPlanNo, vList)

      vList.Clear()
      vList("ValidationTable") = "contacts"
      vList("ValidationAttribute") = "contact_number"
      InitTextLookupBox(txtLookupContactNo, vList)

      GetControlValues()
      ClearFields()
      optFulFilledIncentives.Checked = True
      cmdReset.Enabled = False

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub GetControlValues()
    'Get control values from Financial & Contact Controls
    mvCCReason = AppValues.ControlValue(AppValues.ControlValues.reason_for_despatch)
    mvDDReason = AppValues.ControlValue(AppValues.ControlValues.dd_reason)
    mvSOReason = AppValues.ControlValue(AppValues.ControlValues.so_reason)
    mvOReason = AppValues.ControlValue(AppValues.ControlValues.o_reason)
    mvReasonForDesp = AppValues.ControlValue(AppValues.ControlValues.reason_for_despatch)
  End Sub

  Private Sub InitTextLookupBox(ByVal pTxtLookup As TextLookupBox, ByVal pList As ParameterList)
    Dim vParamList As New ParameterList(True)
    vParamList("TableName") = pList("ValidationTable")
    vParamList("FieldName") = pList("ValidationAttribute")
    vParamList("FieldType") = "C"  ' Character FieldType
    Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vParamList)
    vParams("AttributeName") = pList("ValidationAttribute")
    vParams("ValidationAttribute") = pList("ValidationAttribute")
    vParams("ValidationTable") = pList("ValidationTable")
    If pList.Contains("RestrictionAttribute") Then vParams("RestrictionAttribute") = pList("RestrictionAttribute")

    Dim vPanelItem As PanelItem = New PanelItem(pTxtLookup, vParams("ValidationAttribute"))
    If vPanelItem.ParameterName = "ContactNumber" Then vPanelItem.SetValidationData("contacts", "contact_number")

    vPanelItem.InitFromMaintenanceData(vParams)
    pTxtLookup.Tag = vPanelItem
    pTxtLookup.Init(vPanelItem, False, False)
    pTxtLookup.TotalWidth = pTxtLookup.Width
    pTxtLookup.SetBounds(pTxtLookup.Location.X, pTxtLookup.Location.Y, 80, pTxtLookup.TextBox.Size.Height)
    AddHandler pTxtLookup.Validated, AddressOf LookupChangedHandler
  End Sub

  Private Sub LookupChangedHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If TypeOf (sender) Is TextLookupBox Then
        Dim vTextLookupBox As TextLookupBox = DirectCast(sender, TextLookupBox)
        Dim vPanelItem As PanelItem = DirectCast(vTextLookupBox.Tag, PanelItem)
        If vTextLookupBox Is txtLookupPayPlanNo Then
          If IntegerValue(txtLookupPayPlanNo.Text) > 0 Then GetPaymentPlan(IntegerValue(txtLookupPayPlanNo.Text))
        ElseIf vTextLookupBox Is txtLookupContactNo Then
          If IntegerValue(txtLookupContactNo.Text) > 0 Then GetContact(IntegerValue(txtLookupContactNo.Text))
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ClearFields()
    txtLookupContactNo.Text = ""
    txtLookupPayPlanNo.Text = ""
    optFulFilledIncentives.Checked = True
    tab.SelectedTab = tbpFind
    cmdReset.Enabled = False
    dgrResults.Clear()
  End Sub

  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
    Try

      Dim vEdit As Boolean
      Dim vReasonForDesp As String
      Dim vResponse As DialogResult
      Dim vSource As String

      'First delete existing records
      DeleteIncentives()
      If mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then
        vSource = mvOrderSource
        vReasonForDesp = mvOrderReason
      Else
        vSource = mvSource
        vReasonForDesp = AppValues.ControlValue(AppValues.ControlValues.reason_for_despatch)
      End If

      'Create basic incentives first
      vResponse = ShowQuestion(QuestionMessages.QmBasicPackage, MessageBoxButtons.YesNo)
      If vResponse = System.Windows.Forms.DialogResult.Yes Then vEdit = True

      Dim vList As New ParameterList(True)
      vList("Source") = vSource
      vList("ReasonForDespatch") = vReasonForDesp
      vList("VatCategory") = AppValues.ControlValue(AppValues.ControlValues.default_contact_vat_cat)
      vList("Type") = "P"
      vList("Basic") = "Y"

      If mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled And mvExtraSource.Length > 0 Then
        vList("ExtraSource") = mvExtraSource
        vList("ExtraReason") = mvExtraReason
      End If

      Dim vForm As New frmIncentives
      Dim vDS As DataSet = vForm.GetIncentivesData(vList, vEdit, False, True, False)
      Dim vDT As DataTable = Nothing
      If vDS IsNot Nothing Then vDT = DataHelper.GetTableFromDataSet(vDS)
      If vDT IsNot Nothing Then
        For Each vDR As DataRow In vDT.Rows
          If IntegerValue(vDR("Quantity")) > 0 Then AddNewUnFulfilled(vDR)
        Next
      End If

      'Now create optional incentives
      vList.Clear()
      vList = New ParameterList(True)
      vList("Source") = vSource
      vList("ReasonForDespatch") = vReasonForDesp
      vList("VatCategory") = AppValues.ControlValue(AppValues.ControlValues.default_contact_vat_cat)
      vList("Type") = "P"
      vList("Basic") = "N"
      If txtLookupPayPlanNo.Text.Length > 0 AndAlso mvExtraSource.Length > 0 Then
        vList("ExtraSource") = mvExtraSource
        vList("ExtraReason") = mvExtraReason
      End If
      vForm = New frmIncentives
      vDS = vForm.GetIncentivesData(vList, True, True, True, False)
      vDT = Nothing
      If vDS IsNot Nothing Then vDT = DataHelper.GetTableFromDataSet(vDS)
      If vDT IsNot Nothing Then
        For Each vDR As DataRow In vDT.Rows
          If IntegerValue(vDR("Quantity")) > 0 Then AddNewUnFulfilled(vDR)
        Next
      End If

      'Now that everything is done, re-select the records for the grid
      GetUnFulFilledIncentives()
      cmdReset.Enabled = True     'Re-enable so as to Reset if nothing was selected
      mvClose = False

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmIncentiveMaintenance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    Try
      If mvClose = False Then
        e.Cancel = True
      End If
      mvClose = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function GetUnFulFilledIncentives() As Boolean
    Dim vDataSet As DataSet
    Dim vFound As Boolean
    Dim vMessage As String

    Dim vList As New ParameterList(True)
    If mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled Then
      vList.IntegerValue("ContactNumber") = mvContact
      vMessage = InformationMessages.ImNoUnfulFilledIncForContact
      vList("Source") = mvSource
      vDataSet = DataHelper.GetContactPayPlanIncentives(CType(CareNetServices.XMLTransactionDataSelectionTypes.xtdtUnFulFilledContactIncentives, CareServices.XMLTransactionDataSelectionTypes), vList)
    Else
      vMessage = InformationMessages.ImNoUnFulFilledIncForPaymentPlan
      vList.IntegerValue("PaymentPlanNumber") = mvPayPlan
      vList("Source") = mvSource
      vList("ReasonForDespatch") = mvReasonForDesp
      vDataSet = DataHelper.GetContactPayPlanIncentives(CType(CareNetServices.XMLTransactionDataSelectionTypes.xtdtUnFulFilledPayPlanIncentives, CareServices.XMLTransactionDataSelectionTypes), vList)
    End If

    If vDataSet IsNot Nothing Then
      dgrResults.Populate(vDataSet)

      If dgrResults.DataRowCount > 0 Then
        vFound = True
        tab.SelectedTab = tbpResults
        dgrResults.Focus()
        cmdReset.Enabled = True
      Else
        dgrResults.Clear()
        cmdReset.Enabled = False
        ShowInformationMessage(vMessage)
      End If
    End If
    Return vFound
  End Function

  Private Sub GetFulFilledIncentives()
    Dim vMessage As String = String.Empty
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)

    If mvIncentiveType = imIncentiveTypes.imitContactFulFilled Then
      vMessage = InformationMessages.ImNoFulFilledIncForContact
      vList.IntegerValue("ContactNumber") = mvContact
      vList("Source") = mvSource
      vDataSet = DataHelper.GetContactPayPlanIncentives(CType(CareNetServices.XMLTransactionDataSelectionTypes.xtdtFulFilledContactIncentives, CareServices.XMLTransactionDataSelectionTypes), vList)
    Else
      vMessage = InformationMessages.ImNoFulFilledIncForPaymentPlan
      vList.IntegerValue("PaymentPlanNumber") = mvPayPlan
      vList("Source") = mvSource
      vDataSet = DataHelper.GetContactPayPlanIncentives(CType(CareNetServices.XMLTransactionDataSelectionTypes.xtdtFulFilledPayPlanIncentives, CareServices.XMLTransactionDataSelectionTypes), vList)
    End If

    If vDataSet IsNot Nothing Then
      dgrResults.Populate(vDataSet)

      If dgrResults.DataRowCount > 0 Then
        tab.SelectedTab = tbpResults
        dgrResults.Focus()
      Else
        dgrResults.Clear()
        ShowInformationMessage(vMessage)
      End If
    End If
  End Sub

  Private Sub GetUnFulFilledContactIncentives()
    Dim vFulFilled As Boolean
    Dim vMessage As String = String.Empty
    Dim vNoIncentives As Boolean
    Dim vList As New ParameterList(True)

    'First select the contact_incentive_responses
    vList.IntegerValue("ContactNumber") = mvContact
    vList("CheckDateFulfilled") = "N"
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactIncentiveResponses, vList)
    If vTable IsNot Nothing Then
      If vTable.Rows.Count > 0 Then
        If vTable.Rows(0)("DateFulFilled").ToString.Length > 0 Then
          vFulFilled = True
          If vTable.Rows(0)("source").ToString = mvSource Then
            vMessage = InformationMessages.imIncFulfilledForContact
          End If
        Else
          mvSource = vTable.Rows(0)("source").ToString
        End If
      Else
        vNoIncentives = True
      End If
    End If

    'Check for some incentive scheme products
    If (vFulFilled Or vNoIncentives) AndAlso vMessage.Length = 0 Then
      vList.Clear()
      vList = New ParameterList(True)
      vList("Source") = mvSource
      vList("ReasonForDispatch") = mvReasonForDesp
      Dim vTable1 As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtIncentiveSchemeProducts, vList)
      If vTable1 IsNot Nothing Then
        If vTable1.Rows.Count > 0 Then
          mvMailing = vTable1.Rows(0)("ThankYouLetter").ToString
        Else
          vMessage = InformationMessages.ImNoUnfulFilledIncForContact
        End If
      End If
    End If
    If vMessage.Length > 0 Then
      ShowInformationMessage(vMessage)
    Else
      GetUnFulFilledIncentives()
    End If
  End Sub

  Private Sub GetUnFulFilledPayPlanIncentives()
    Dim vContinue As Boolean
    Dim vFulFilled As Boolean
    Dim vMessage As String = String.Empty
    Dim vNoIncentives As Boolean
    Dim vDD As String
    Dim vSO As String
    Dim vCC As String
    Dim vMemPayPlan As Boolean
    Dim vDespReason As String
    Dim vList As New ParameterList(True)


    'First find some details from the order
    mvPP = New PaymentPlanInfo(mvPayPlan)

    vList.IntegerValue("PaymentPlanNumber") = mvPayPlan
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtOrders, vList)
    If vTable IsNot Nothing Then
      If vTable.Rows.Count > 0 Then
        mvPP = New PaymentPlanInfo(mvPayPlan)

        With mvPP
          mvSource = .Source
          mvOrderSource = .Source
          mvReasonForDesp = .ReasonForDespatch
          vDD = .DirectDebitStatus
          vSO = .StandingOrderStatus
          vCC = .CreditCardStatus
          Select Case .PlanType
            Case PaymentPlanInfo.ppType.pptMember
              vMemPayPlan = True
              mvOrderReason = .ReasonForDespatch
            Case PaymentPlanInfo.ppType.pptDD
              mvOrderReason = mvDDReason
            Case PaymentPlanInfo.ppType.pptSO
              mvOrderReason = mvSOReason
            Case PaymentPlanInfo.ppType.pptCCCA
              mvOrderReason = mvCCReason
            Case Else
              mvOrderReason = mvOReason
          End Select
        End With
        'Now check the new_orders table
        vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtNewOrders, vList)
        If vTable IsNot Nothing Then
          If vTable.Rows.Count > 0 Then
            vDespReason = vTable.Rows(0)("ReasonForDespatch").ToString
            If vTable.Rows(0)("DateFulFilled").ToString.Length > 0 Then
              vFulFilled = True
              If vMemPayPlan Then
                If (vDespReason = mvReasonForDesp) _
                Or ((vDD = "Y" And vDespReason = mvDDReason) _
                Or (vSO = "Y" And vDespReason = mvSOReason) _
                Or (vCC = "Y" And vDespReason = mvCCReason)) Then
                  vMessage = InformationMessages.ImIncFulFilledForPaymentPlan
                End If
              Else
                If (vDD = "Y" And vDespReason = mvDDReason) _
                Or (vSO = "Y" And vDespReason = mvSOReason) _
                Or (vCC = "Y" And vDespReason = mvCCReason) Then
                  vMessage = InformationMessages.ImIncFulFilledForPaymentPlan
                End If
              End If
            Else
              'Date fulfilled = null
              If vMemPayPlan Then
                If vDD = "Y" And vDespReason = mvDDReason Then mvOrderReason = mvDDReason
                If vSO = "Y" And vDespReason = mvSOReason Then mvOrderReason = mvSOReason
                If vCC = "Y" And vDespReason = mvCCReason Then mvOrderReason = mvCCReason
              End If
            End If
          Else
            vNoIncentives = True
          End If
        End If

        'Check for any extra source / reason for despatch
        If vDD = "Y" Or vSO = "Y" Or vCC = "Y" Then
          RetrieveAutoPM(mvOrderSource, mvOrderReason, mvPP)
        End If
      Else
        vMessage = InformationMessages.ImPPCanc
      End If
    End If

    If (vNoIncentives OrElse vFulFilled) AndAlso vMessage.Length = 0 Then
      vList = New ParameterList(True)
      vList("Source") = mvOrderSource
      vList("ReasonForDispatch") = mvOrderReason
      Dim vTable1 As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtIncentiveSchemeProducts, vList)
      If vTable1 IsNot Nothing Then
        If vTable1.Rows.Count > 0 Then
          vContinue = True
          mvMailing = vTable1.Rows(0)("Source").ToString
        End If
      End If

      'Now check against extra source
      If mvExtraSource.Length > 0 Then
        vList = New ParameterList(True)
        vList("Source") = mvExtraSource
        vList("ReasonForDispatch") = mvExtraReason
        vTable1 = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtIncentiveSchemeProducts, vList)
        If vTable1 Is Nothing OrElse vTable1.Rows.Count = 0 Then
          If Not (vContinue) Then
            vMessage = InformationMessages.ImNoUnFulFilledIncForPaymentPlan
          End If
        End If
      Else
        If Not (vContinue) Then vMessage = InformationMessages.ImNoUnFulFilledIncForPaymentPlan
      End If
    End If

    If vMessage.Length > 0 Then
      ShowInformationMessage(vMessage)
    Else
      GetUnFulFilledIncentives()
    End If
  End Sub

  Private Function GetContact(ByVal pContactNumber As Integer) As Boolean
    Try
      Dim vContact As New ContactInfo(pContactNumber)
      If vContact.ContactNumber > 0 Then
        mvPayPlan = 0
        mvContact = vContact.ContactNumber
        If vContact.NameGatheringSource.Length > 0 Then
          mvSource = vContact.NameGatheringSource
        Else
          mvSource = vContact.Source
        End If
      Else
        mvContact = 0
      End If
      Return True
    Catch
      mvPayPlan = 0
      mvContact = 0
      mvSource = ""
      Return False
    End Try
  End Function

  Private Function GetPaymentPlan(ByVal pPayPlanNo As Integer) As Boolean
    Dim vPP As PaymentPlanInfo = New PaymentPlanInfo(pPayPlanNo)
    mvPayPlan = pPayPlanNo
    mvContact = vPP.ContactNumber
    mvSource = CStr(IIf(vPP.ContactNumber > 0, vPP.Source, String.Empty))
    mvReasonForDesp = CStr(IIf(vPP.ContactNumber > 0, vPP.ReasonForDespatch, String.Empty))
    Return vPP.ContactNumber > 0
  End Function

  Private Sub RetrieveAutoPM(ByVal pOrderSource As String, ByVal pOrderReason As String, ByVal pPayPlan As PaymentPlanInfo)
    'Check autopayment methods for source & reason for despatch
    Dim vReason As String = String.Empty
    Dim vTable As DataTable = Nothing
    Dim vList As New ParameterList(True)
    vList.IntegerValue("PaymentPlanNumber") = mvPayPlan

    If pPayPlan.DirectDebitStatus = "Y" Then
      vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtDirectDebitSource, vList)
      vReason = mvDDReason
    ElseIf pPayPlan.StandingOrderStatus = "Y" Then
      vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBankersOrderSource, vList)
      vReason = mvSOReason
    ElseIf pPayPlan.CreditCardStatus = "Y" Then
      vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCardAuthoritiesSource, vList)
      vReason = mvCCReason
    End If

    If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
      If pOrderReason = vReason Then
        pOrderSource = vTable.Rows(0)("Source").ToString
        mvExtraSource = ""
        mvExtraReason = ""
      Else
        mvExtraSource = vTable.Rows(0)("Source").ToString
        mvExtraReason = vReason
      End If
    End If
  End Sub

  Private Sub AddNewUnFulfilled(ByVal pDR As DataRow)
    Dim vList As New ParameterList(True)

    vList.IntegerValue("ContactNumber") = mvContact
    If mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then

      'Find the correct contact number to use
      If mvPP.PlanType = PaymentPlanInfo.ppType.pptMember Then
        If pDR("ForWhom").ToString = "P" Then
          'Set to payers contact number
          vList.IntegerValue("ContactNumber") = mvPP.ContactNumber
        Else
          'Set to first members contact number where rfd = mem-type
          vList("GetContactFromPayPlan") = "Y"
        End If
      End If
    End If

    vList("Product") = pDR("Product").ToString
    vList("Quantity") = pDR("Quantity").ToString
    vList("ReasonForDespatch") = pDR("ReasonForDespatch").ToString

    If mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then
      vList.IntegerValue("PaymentPlanNumber") = mvPayPlan
      DataHelper.AddEnclosures(vList)
    ElseIf mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled Then
      vList("Source") = mvSource
      DataHelper.AddUnFulFilledContactIncentives(vList)
    End If
  End Sub

  Private Sub DeleteIncentives()
    'Delete all Incentives / Enclosures
    Dim vSource As String = String.Empty

    If mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled Then
      'Contact - Delete contact_incentives
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ContactNumber") = mvContact
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactIncentiveResponses, vList)

      If vTable IsNot Nothing Then
        If vTable.Rows.Count > 0 Then
          If vSource.Length > 0 Then
            If Strings.Left(vSource, 1) <> "'" Then vSource = "'" & vSource & "'"
            vSource = vSource & ",'" & vTable.Rows(0)("Source").ToString & "'"
          Else
            vSource = vTable.Rows(0)("Source").ToString
          End If
        End If
      End If

      If vSource.Length > 0 Then
        vList.Clear()
        vList = New ParameterList(True)
        vList.IntegerValue("ContactNumber") = mvContact
        vList("MultipleSources") = vSource
        DataHelper.DeleteContactIncentives(vList)
      End If

    ElseIf mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then
      'Payment Plan - Delete enclosures
      Dim vList As New ParameterList(True)
      vList.IntegerValue("PaymentPlanNumber") = mvPayPlan
      DataHelper.DeleteEnclosures(vList)
    End If
  End Sub

  Private Sub SetFinderFocus()
    tab.SelectedTab = tbpFind
    cmdReset.Enabled = False
  End Sub

  Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
    Try
      Dim vContinue As Boolean
      Dim vMessage As String = String.Empty

      If tab.SelectedTab IsNot tbpFind Then
        SetFinderFocus()
      Else
        If (txtLookupPayPlanNo.Text.Length > 0 AndAlso txtLookupContactNo.Text.Length > 0) _
        OrElse (txtLookupPayPlanNo.Text.Length = 0 AndAlso txtLookupContactNo.Text.Length = 0) Then
          'Both or no values entered
          vMessage = InformationMessages.ImEnterContactNoOrPaymentPlan
        Else
          vContinue = True
          If txtLookupContactNo.Text.Length > 0 Then
            If optFulFilledIncentives.Checked = True Then
              mvIncentiveType = imIncentiveTypes.imitContactFulFilled
            Else
              mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled
            End If
          Else
            'Pay Plan number entered
            If optFulFilledIncentives.Checked = True Then
              mvIncentiveType = imIncentiveTypes.imitPayPlanFulFilled
            Else
              mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled
            End If
          End If
        End If
        If vContinue Then
          SetDefaults()
        Else
          ShowInformationMessage(vMessage)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Public Sub SetDefaults()
    Dim vContinue As Boolean
    Dim vMessage As String = String.Empty

    mvExtraSource = ""
    mvExtraReason = ""
    If txtLookupContactNo.Text.Length > 0 Then
      'Contact
      If (IntegerValue(txtLookupContactNo.Text) <> mvContact) OrElse txtLookupContactNo.Label.Text.Length = 0 Then 'OrElse Len(lblDesc(icContact).Caption) = 0 Then
        vContinue = GetContact(IntegerValue(txtLookupContactNo.Text))
      Else
        vContinue = True
      End If
      If Not (vContinue) Then vMessage = InformationMessages.ImInvalidContactNumber
    ElseIf txtLookupPayPlanNo.Text.Length > 0 Then
      'Payment Plan
      If (IntegerValue(txtLookupPayPlanNo.Text) <> mvPayPlan) OrElse txtLookupPayPlanNo.Label.Text.Length = 0 Then
        vContinue = GetPaymentPlan(IntegerValue(txtLookupPayPlanNo.Text))
      Else
        vContinue = True
      End If
      If Not (vContinue) Then vMessage = InformationMessages.ImInvalidPaymentPlanEntered
    End If

    If vContinue Then
      If mvIncentiveType = imIncentiveTypes.imitContactFulFilled Or mvIncentiveType = imIncentiveTypes.imitPayPlanFulFilled Then
        GetFulFilledIncentives()
      ElseIf mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled Then
        GetUnFulFilledContactIncentives()
      ElseIf mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then
        GetUnFulFilledPayPlanIncentives()
      End If
    End If
    If Not (vContinue) Then ShowInformationMessage(vMessage)
  End Sub

  Private Sub ProcessFind()
    Dim vList As New ParameterList(True)
  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      ClearFields()
      txtLookupPayPlanNo.Focus()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub tab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
    Try
      If tab.SelectedTab Is tbpFind Then
        Me.AcceptButton = cmdFind
        cmdFind.Enabled = True
        cmdReset.Enabled = False
      Else
        cmdFind.Enabled = False
        If mvIncentiveType = imIncentiveTypes.imitContactUnFulFilled Or mvIncentiveType = imIncentiveTypes.imitPayPlanUnFulFilled Then
          cmdReset.Enabled = True
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class