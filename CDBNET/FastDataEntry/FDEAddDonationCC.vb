Friend Class FDEAddDonationCC
  Inherits CareFDEControl

  Private mvSourceCode As String = ""
  Private mvDefaultSourceCode As String           'From Add Transaction Details control
  Private mvDefaultDistributionCode As String     'From Source on Add Transaction Details control
  Private mvIncentiveScheme As String = ""
  Private mvIncentiveSequenceList As String = ""
  Private mvIncentiveQuantityList As String = ""
  Private mvProductDefaulted As Boolean
  Private mvType As CareNetServices.FDEControlTypes


  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
    mvSupportsSourceChanged = True
    mvSupportsTransactionDateChanged = True
    mvSupportsClearBankDetails = True
    mvType = pType
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
    mvSupportsSourceChanged = True
    mvSupportsTransactionDateChanged = True
    mvSupportsClearBankDetails = True
    mvType = pType
  End Sub
  Friend Overrides Sub SetDefaults()
    MyBase.SetDefaults()
    If mvRetainSource = False Then mvDefaultSourceCode = ""
    mvProductDefaulted = False      'Reset this flag for 2nd and subsequent transactions
    epl.FillDeferredCombos(epl)
    epl.SetValue("Quantity", "1")   'Set the quantity before raising event for Product and Rate
    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then
      SetDefaultProductAndRate()
    Else
      mvProductDefaulted = True
    End If
    Dim vProduct As String = epl.GetValue("Product")
    Dim vRate As String = epl.GetValue("Rate")
    If vProduct.Length > 0 Then SetValueRaiseChanged(epl, "Product", vProduct)
    If vRate.Length > 0 Then SetValueRaiseChanged(epl, "Rate", vRate)
    Dim vPaymentMethod As String = epl.GetValue("PaymentMethod")
    If vPaymentMethod.Length > 0 Then SetValueRaiseChanged(epl, "PaymentMethod", vPaymentMethod)
    Dim vCardNumberNotRequired As CheckBox = epl.FindPanelControl(Of CheckBox)("CardNumberNotRequired", False)
    If vCardNumberNotRequired IsNot Nothing AndAlso
       vCardNumberNotRequired.Checked Then
      SetValueRaiseChanged(epl, "CardNumberNotRequired", "Y")
    End If
    epl.EnableControl("DeceasedContactNumber", False)
    If mvInitialSettings.Length > 0 Then
      Dim vList As New ParameterList
      vList.FillFromValueList(mvInitialSettings)
      If WebBasedCardAuthoriser.IsAvailable Then
        epl.EnableControl("SecurityCode", False)
      Else
        If vList.ContainsKey("OnlineCCAuthorisation") AndAlso BooleanValue(vList("OnlineCCAuthorisation")) Then
          epl.PanelInfo.PanelItems("SecurityCode").Mandatory = False
        Else
          epl.EnableControl("SecurityCode", False)
        End If
      End If
    End If
    If mvProductValidation.LinePrice = 0 Then epl.SetValue("Amount", "")
    If mvDefaultDistributionCode IsNot Nothing AndAlso FindControl(epl, "DistributionCode_" & mvDefaultDistributionCode, False) IsNot Nothing Then SetValueRaiseChanged(epl, "DistributionCode_" & mvDefaultDistributionCode, mvDefaultDistributionCode)
    epl.DataChanged = False
  End Sub

  Friend Overrides Sub RefreshSource(ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
    MyBase.RefreshSource(pSourceCode, pDistributionCode, pIncentiveScheme)
    If mvDefaultSourceCode Is Nothing OrElse mvDefaultSourceCode.Length = 0 Then
      mvDefaultSourceCode = pSourceCode
      mvDefaultDistributionCode = pDistributionCode
      mvProductDefaulted = False
    End If
    mvSourceCode = pSourceCode
    mvIncentiveScheme = pIncentiveScheme
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
    'BR13463: Set DistributionCode after SetDefaultProductAndRate as it might be reset by setting the default product and rate
    SetDefaultProductAndRate()
    If FindControl(epl, "DistributionCode_" & pDistributionCode, False) IsNot Nothing Then SetValueRaiseChanged(epl, "DistributionCode_" & pDistributionCode, pDistributionCode)
    mvProductDefaulted = True
  End Sub

  Friend Overrides Sub RefreshTransactionDate(ByVal pTransactionDate As String)
    MyBase.RefreshTransactionDate(pTransactionDate)
    mvProductValidation.TransactionDate = pTransactionDate
  End Sub

  Friend Overrides Sub ResetIncentives()
    MyBase.ResetIncentives()
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
  End Sub

  Friend Overrides Sub ResetBankDetails()
    mvBankDetailsNumber = 0
    If FindControl(Me, "SortCode", False) IsNot Nothing Then
      epl.ClearControlList("SortCode,AccountNumber")
    End If
  End Sub
  Friend Overrides Function BuildParameterList(ByRef pList As CDBNETCL.ParameterList) As Boolean
    Dim vValid As Boolean = True
    Dim vAmount As String = epl.GetValue("Amount")
    epl.PanelInfo.PanelItems("Product").Mandatory = (vAmount.Length > 0)
    epl.PanelInfo.PanelItems("Rate").Mandatory = (vAmount.Length > 0)
    If vAmount.Length > 0 Then
      If epl.GetValue("Quantity").Length = 0 Then epl.SetValue("Quantity", "1")
      ClearOptionButtonError("DistributionCode")
      vValid = MyBase.BuildParameterList(pList)
      If pList.ContainsKey("DistributionCode") = False Then
        Dim vParameterName As String = "DistributionCode"
        If MyBase.GetMandatoryOptionButton(pList, vParameterName) = False Then
          vValid = epl.SetErrorField(vParameterName, GetInformationMessage(InformationMessages.ImFieldMandatory))
        End If
      End If
      Dim vDonCCList As New ParameterList
      vDonCCList.FillFromValueList(mvInitialSettings)
      If vDonCCList.ContainsKey("OnlineCCAuthorisation") Then
        If BooleanValue(vDonCCList("OnlineCCAuthorisation")) Then pList("GetAuthorisation") = "Y"
      End If
      If mvContactInfo.SelectedAddressNumber = 0 Then mvContactInfo.SelectedAddressNumber = mvContactInfo.AddressNumber
      pList.IntegerValue("DeliveryContactNumber") = mvContactInfo.ContactNumber
      pList.IntegerValue("DeliveryAddressNumber") = mvContactInfo.SelectedAddressNumber
      pList("BankAccount") = vDonCCList("BankAccount")
      If mvProductValidation.LineVATAmount = 0 AndAlso mvProductValidation.LineVATPercentage() > 0 Then
        'VAT was not recalculated so do it now
        mvProductValidation.LineVATAmount = AppHelper.CalculateVATAmount((DoubleValue(epl.GetValue("Amount")) * DoubleValue(epl.GetValue("Quantity"))), mvProductValidation.LineVATPercentage)
      End If
      pList("VatAmount") = mvProductValidation.LineVATAmount.ToString
      Dim vLineType As String = "P"
      If epl.GetValue("LineTypeG") = "Y" Then
        vLineType = "G"
      ElseIf epl.GetValue("LineTypeH") = "Y" Then
        vLineType = "H"
      ElseIf epl.GetValue("LineTypeS") = "Y" Then
        vLineType = "S"
      End If
      pList("LineType") = vLineType
      If pList.ContainsKey("DistributionCodeLookupGroup") Then pList.Remove("DistributionCodeLookupGroup")
      If mvIncentiveSequenceList.Length > 0 Then pList("IncentiveSequence") = mvIncentiveSequenceList
      If mvIncentiveQuantityList.Length > 0 Then pList("IncentiveQuantity") = mvIncentiveQuantityList
      If mvBankDetailsNumber > 0 Then pList.IntegerValue("BankDetailsNumber") = mvBankDetailsNumber
    ElseIf epl.GetValue("CreditCardNumber").Length > 0 Then
      vValid = epl.ValidateControl("CreditCardType")
      If epl.ValidateControl("CardExpiryDate") = False Then vValid = False
      If epl.ValidateControl("IssueNumber") = False Then vValid = False
      If epl.ValidateControl("CardStartDate") = False Then vValid = False
    End If
    If Not pList.Contains("PaymentMethod") AndAlso epl.GetValue("PaymentMethod").Length > 0 Then
      pList.Add("PaymentMethod", epl.GetValue("PaymentMethod"))
    End If
    Return vValid
  End Function

  Friend Overrides Function CheckIncentives(ByRef pList As ParameterList) As Boolean
    Dim vAmount As String = epl.GetValue("Amount")
    Dim vCheckIncentives As Boolean = MyBase.CheckIncentives(pList)
    If vAmount.Length > 0 AndAlso (mvSourceCode.Length > 0 AndAlso mvIncentiveScheme.Length > 0) AndAlso (mvIncentiveSequenceList.Length = 0 AndAlso mvIncentiveQuantityList.Length = 0) Then
      Dim vRFD As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.payment_reason)
      Dim vList As New ParameterList
      epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly)
      vCheckIncentives = (vRFD.Length > 0)
      If vCheckIncentives Then
        If pList Is Nothing Then pList = New ParameterList()
        pList("Source") = mvSourceCode
        pList("ReasonForDespatch") = vRFD
        pList("Amount") = epl.GetValue("Amount")
        If mvContactInfo IsNot Nothing Then pList("VatCategory") = mvContactInfo.VATCategory
      End If
    End If
    Return vCheckIncentives
  End Function

  Friend Overrides Sub AddIncentives(ByVal pSequenceNumbers As String, ByVal pQuantity As String)
    MyBase.AddIncentives(pSequenceNumbers, pQuantity)
    mvIncentiveSequenceList = pSequenceNumbers
    mvIncentiveQuantityList = pQuantity
  End Sub

  Friend Overrides ReadOnly Property CanSubmit() As Boolean
    Get
      Dim vCanSubmit As Boolean = MyBase.CanSubmit
      If epl.GetValue("Amount").Length > 0 Then vCanSubmit = True
      Return vCanSubmit
    End Get
  End Property

  Private Sub SetDefaultProductAndRate()
    If mvProductDefaulted = False AndAlso mvDefaultSourceCode IsNot Nothing AndAlso mvDefaultSourceCode.Length > 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then
      Dim vList As New ParameterList(True, True)
      vList.FillFromValueList(mvDefaultSettings)
      vList("FdeUserControl") = mvUserControlName
      vList("Source") = mvDefaultSourceCode
      Dim vReturnList As ParameterList = DataHelper.GetFastDataEntryModuleDefaults(vList)
      If vReturnList IsNot Nothing Then
        If vReturnList.Contains("Product") Then SetValueRaiseChanged(epl, "Product", vReturnList("Product"), , True)
        If vReturnList.Contains("Rate") Then SetValueRaiseChanged(epl, "Rate", vReturnList("Rate"), , True)
      End If
      mvProductDefaulted = True
    End If
  End Sub

  Friend Overrides Sub GetPaymentMethodParameters(ByVal pList As CDBNETCL.ParameterList)
    MyBase.GetPaymentMethodParameters(pList)
    Dim vCCNumber As String = epl.GetValue("CreditCardNumber")
    If vCCNumber.Length > 0 Then
      pList("CreditCardNumber") = vCCNumber
      pList("CreditCardType") = epl.GetValue("CreditCardType")
      pList("CardExpiryDate") = epl.GetValue("CardExpiryDate")
      pList("IssueNumber") = epl.GetValue("IssueNumber")
      pList("CardStartDate") = epl.GetValue("CardStartDate")
      Dim vDonCCList As New ParameterList
      vDonCCList.FillFromValueList(mvInitialSettings)
      If vDonCCList.ContainsKey("OnlineCCAuthorisation") Then
        If BooleanValue(vDonCCList("OnlineCCAuthorisation")) Then
          pList("CardSecurityCode") = epl.GetValue("CardSecurityCode")
          pList("GetAuthorisation") = "Y"
        End If
      End If
      If vDonCCList.ContainsKey("BankAccount") Then pList("BankAccount") = vDonCCList("BankAccount")
    End If
  End Sub

  Friend Overrides Sub GetCodeRestrictions(ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList)
    MyBase.GetCodeRestrictions(pParameterName, pList)
    Select Case pParameterName
      Case "Product"
        If mvType = CareNetServices.FDEControlTypes.AddDonationCC Then
          pList("FindProductType") = "O"        'donation or sponsorship event
          If mvSourceCode Is Nothing Then mvSourceCode = ""
          If mvSourceCode.Length > 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then pList("ProductSource") = mvSourceCode
        ElseIf mvType = CareNetServices.FDEControlTypes.ProductSale Then
          pList("FindProductType") = "Z"       'need type for non donation and non product types
          If mvSourceCode Is Nothing Then mvSourceCode = ""
          If mvSourceCode.Length > 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then pList("ProductSource") = mvSourceCode

        End If
End Select
  End Sub

End Class
