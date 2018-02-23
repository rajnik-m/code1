Friend Class FDEAddRegularDonation
  Inherits CareFDEControl

  Friend PPDDataSet As DataSet  'Holds the Detail Lines
  Friend DetailLineDefaults As ParameterList  'Default Source, Distribution Code, Quantity, Delivery Contact and Address

  Private mvDefaultSourceCode As String = ""         'From Add Transaction Details control
  Private mvIncentiveScheme As String = ""
  Private mvIncentiveSequenceList As String = ""
  Private mvIncentiveQuantityList As String = ""
  Private mvIncentivesProcessed As Boolean

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsSourceChanged = True
    mvSupportsTransactionDateChanged = True
    DirectCast(FindControl(epl, "DetailLines"), DisplayGrid).AllowSorting = False
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsSourceChanged = True
    mvSupportsTransactionDateChanged = True
    DirectCast(FindControl(epl, "DetailLines"), DisplayGrid).AllowSorting = False
  End Sub
  Friend Sub SetBalance(ByVal pBalance As String)

    SetDefaults()
  End Sub
  Friend Overrides Sub SetDefaults()
    epl.FillDeferredCombos(epl)
    MyBase.SetDefaults()
    If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) <> "D" Then epl.EnableControl("ClaimDay", False)
    Dim vList As New ParameterList
    vList.FillFromValueList(mvDefaultSettings)
    Dim vValue As String = ""
    If vList.ContainsKey("PaymentFrequency") Then vValue = vList("PaymentFrequency")
    If vValue.Length > 0 Then SetValueRaiseChanged(epl, "PaymentFrequency_" & vValue, vValue)
    Dim vClaimDay As String = epl.GetValue("ClaimDay")
    vValue = epl.GetValue("BankAccount")
    SetValueRaiseChanged(epl, "BankAccount", vValue)
    If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then SetValueRaiseChanged(epl, "ClaimDay", vClaimDay)
    With epl.PanelInfo.PanelItems
      .Item("ReasonForDespatch").Mandatory = True
      .Item("SortCode").Mandatory = False
      .Item("AccountNumber").Mandatory = False
      .Item("AccountName").Mandatory = False
      .Item("BranchName").Mandatory = False
      .Item("StartDate").Mandatory = False
      .Item("ClaimDay").Mandatory = False
    End With
    If vList.ContainsKey("MandateType") = False Then epl.SetValue("MandateType", "") 'This is the value for 'Unknown'
    epl.SetErrorField("ClaimDay", "")
    SetDDStartDate(Today, "StartDate")
    epl.SetValue("ReasonForDespatch", AppValues.ControlValue(AppValues.ControlValues.o_reason))

    If PPDDataSet Is Nothing OrElse Not PPDDataSet.Tables.Contains("Column") Then
      'Build the PPDDataSet
      If PPDDataSet Is Nothing Then PPDDataSet = New DataSet
      Dim vTable As DataTable = DataHelper.NewColumnTable
      DataHelper.AddDataColumn(vTable, "Product", "Product")
      DataHelper.AddDataColumn(vTable, "Rate", "Rate")
      DataHelper.AddDataColumn(vTable, "DistributionCode", "Distribution Code")
      DataHelper.AddDataColumn(vTable, "Quantity", "Quantity", "Long")
      DataHelper.AddDataColumn(vTable, "Amount", "Amount", "Number")
      DataHelper.AddDataColumn(vTable, "Balance", "Balance", "Number")
      DataHelper.AddDataColumn(vTable, "Source", "Source")
      DataHelper.AddDataColumn(vTable, "ContactNumber", "Contact Number")
      DataHelper.AddDataColumn(vTable, "AddressNumber", "Address Number")
      DataHelper.AddDataColumn(vTable, "ModifierActivity", "Activity")
      DataHelper.AddDataColumn(vTable, "ModifierActivityValue", "Activity Value")
      DataHelper.AddDataColumn(vTable, "ModifierActivityQuantity", "Activity Quantity", "Long")
      DataHelper.AddDataColumn(vTable, "ModifierActivityDate", "Activity Date", "Date")
      DataHelper.AddDataColumn(vTable, "ModifierPrice", "Modifier Price", "Number")
      DataHelper.AddDataColumn(vTable, "ModifierPerItem", "Per Item")
      DataHelper.AddDataColumn(vTable, "UnitPrice", "Unit Price", "Number")
      DataHelper.AddDataColumn(vTable, "ProRated", "Pro-Rated?")
      DataHelper.AddDataColumn(vTable, "NetAmount", "Net", "Number")
      DataHelper.AddDataColumn(vTable, "VatAmount", "VAT", "Number")
      DataHelper.AddDataColumn(vTable, "GrossAmount", "Gross", "Number")
      'Hidden columns
      DataHelper.AddDataColumn(vTable, "LineNumber", "Line Number", , "N")
      DataHelper.AddDataColumn(vTable, "Arrears", "Arrears", "Number", "N")
      DataHelper.AddDataColumn(vTable, "VatRate", "Vat Rate", "Char", "N")
      DataHelper.AddDataColumn(vTable, "VatPercentage", "VAT %", "Number", "N")
      PPDDataSet.Tables.Add(vTable)
      PPDDataSet.Tables.Add(DataHelper.NewDataTable(PPDDataSet.Tables("Column")))
    End If

    AddDefaultDetailLine(vList, False)  'Used when Next button is pressed
    Dim vGrid As DisplayGrid = DirectCast(FindControl(epl, "DetailLines"), DisplayGrid)
    vGrid.Populate(PPDDataSet, True)

    If Not mvRetainSource Then  'Do not reset the detail line defaults if Next button is pressed
      DetailLineDefaults = New ParameterList
      DetailLineDefaults.IntegerValue("Quantity") = 1
    End If
    epl.DataChanged = False
  End Sub

  Private Sub AddDefaultDetailLine(ByVal pList As ParameterList, ByVal pRePopulateGrid As Boolean)
    If PPDDataSet.Tables("DataRow").Rows.Count > 0 Then PPDDataSet.Tables("DataRow").Rows.Clear()
    If pList.ContainsKey("Product") AndAlso pList.ContainsKey("Rate") AndAlso mvContactInfo IsNot Nothing Then
      'Set default detail line only when the contact is selected
      Dim vQuantity As Integer = 1
      Dim vBalance As Double = mvProductValidation.GetAmount(pList("Product"), pList("Rate"), vQuantity, epl)
      Dim vRow As DataRow = PPDDataSet.Tables("DataRow").NewRow
      With vRow
        .Item("Product") = pList("Product")
        .Item("Rate") = pList("Rate")
        .Item("DistributionCode") = DetailLineDefaults.ValueIfSet("DistributionCode")
        .Item("Quantity") = vQuantity
        .Item("Amount") = ""
        .Item("Balance") = vBalance.ToString("0.00")
        .Item("Source") = DetailLineDefaults.ValueIfSet("Source")
        .Item("ContactNumber") = DetailLineDefaults.ValueIfSet("ContactNumber")
        .Item("AddressNumber") = DetailLineDefaults.ValueIfSet("AddressNumber")
        .Item("LineNumber") = 1
      End With
      Dim vPPDPricing As PaymentPlanDetailsPricing = DataHelper.GetModifierPriceData(pList("Product"), pList("Rate"), Today, mvContactInfo.ContactNumber)
      If vPPDPricing IsNot Nothing AndAlso vPPDPricing.GotPricingData Then
        vPPDPricing.PopulateRowWithPricingData(vRow)
      End If
      PPDDataSet.Tables("DataRow").Rows.Add(vRow)
      If pRePopulateGrid Then
        'Re-populate grid to set correct column widths when adding the default line on selection of a contact
        Dim vGrid As DisplayGrid = DirectCast(FindControl(epl, "DetailLines"), DisplayGrid)
        vGrid.Populate(PPDDataSet)
      End If
      epl.SetValue("Balance", vBalance.ToString("0.00"))
    End If
  End Sub

  Friend Overrides Sub RefreshContactData(ByVal pContactInfo As CDBNETCL.ContactInfo)
    MyBase.RefreshContactData(pContactInfo)
    DetailLineDefaults.IntegerValue("ContactNumber") = pContactInfo.ContactNumber
    DetailLineDefaults.IntegerValue("AddressNumber") = pContactInfo.AddressNumber
    If PPDDataSet.Tables("DataRow").Rows.Count = 0 Then
      'Add the default line if not already added. Delivery Contact would need to be changed manually
      Dim vList As New ParameterList
      vList.FillFromValueList(mvDefaultSettings)
      AddDefaultDetailLine(vList, True)
    End If
  End Sub

  Friend Overrides Sub RefreshSource(ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
    MyBase.RefreshSource(pSourceCode, pDistributionCode, pIncentiveScheme)
    DetailLineDefaults("Source") = pSourceCode
    DetailLineDefaults("DistributionCode") = pDistributionCode
    For Each vRow As DataRow In PPDDataSet.Tables("DataRow").Rows
      vRow("Source") = pSourceCode  'Always change the source on all detail lines
      If vRow("DistributionCode").ToString.Length = 0 Then vRow("DistributionCode") = pDistributionCode 'Only change the distribution code if not already set
    Next
    'Set incentive values
    mvDefaultSourceCode = pSourceCode
    mvIncentiveScheme = pIncentiveScheme
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
    mvIncentivesProcessed = False
  End Sub

  Friend Overrides Sub RefreshTransactionDate(ByVal pTransactionDate As String)
    MyBase.RefreshTransactionDate(pTransactionDate)
    mvProductValidation.TransactionDate = pTransactionDate
  End Sub

  Friend Overrides Sub ResetIncentives()
    MyBase.ResetIncentives()
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
    mvIncentivesProcessed = False
  End Sub

  Friend Overrides Function CheckIncentives(ByRef pList As ParameterList) As Boolean
    Dim vCheckIncentives As Boolean = MyBase.CheckIncentives(pList)
    Dim vAmount As String = epl.GetValue("Balance")
    If vAmount.Length > 0 AndAlso mvDefaultSourceCode.Length > 0 AndAlso mvIncentiveScheme.Length > 0 AndAlso mvIncentivesProcessed = False AndAlso PPDDataSet.Tables("DataRow").Rows.Count > 0 Then
      Dim vRFD As String = epl.GetValue("ReasonForDespatch")
      vCheckIncentives = (vRFD.Length > 0)
      If vCheckIncentives Then
        If pList Is Nothing Then pList = New ParameterList()
        pList("Source") = mvDefaultSourceCode
        pList("ReasonForDespatch") = vRFD
        pList("Amount") = epl.GetValue("Balance")
        If mvContactInfo IsNot Nothing Then pList("VatCategory") = mvContactInfo.VATCategory
      End If
    End If
    Return vCheckIncentives
  End Function

  Friend Overrides Sub AddIncentives(ByVal pSequenceNumbers As String, ByVal pQuantity As String)
    MyBase.AddIncentives(pSequenceNumbers, pQuantity)
    'Incentives will be added once for one payment plan
    mvIncentiveSequenceList = pSequenceNumbers
    mvIncentiveQuantityList = pQuantity
    mvIncentivesProcessed = True
  End Sub

  Friend Overrides ReadOnly Property CanSubmit() As Boolean
    Get
      Dim vCanSubmit As Boolean = MyBase.CanSubmit
      If PPDDataSet.Tables("DataRow").Rows.Count > 0 Then vCanSubmit = True
      Return vCanSubmit
    End Get
  End Property

  Friend Overrides Function BuildParameterList(ByRef pList As CDBNETCL.ParameterList) As Boolean
    Dim vValid As Boolean = True
    If CanSubmit() Then
      pList.FillFromValueList(mvInitialSettings)
      vValid = MyBase.BuildParameterList(pList)
      If vValid Then
        epl.SetErrorField("DetailLines", "")
        For Each vRow As DataRow In PPDDataSet.Tables("DataRow").Rows
          If vRow("Balance").ToString() = "0.00" Then
            Dim vList As New ParameterList(True)
            Dim vProductRow As DataRow
            vList("Product") = vRow("Product").ToString()
            vList("SystemColumns") = "N"
            vProductRow = DataHelper.GetRowFromDataSet(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftProducts, vList))
            If vProductRow IsNot Nothing Then
              If vProductRow("Donation").ToString = "Y" Then
                epl.SetErrorField("DetailLines", GetInformationMessage(InformationMessages.ImDonationBalanceCannotBeZero))
                vValid = False
              End If
            End If
          End If
        Next
        If pList.Contains("Product") Then pList.Remove("Product")
        If pList.Contains("Rate") Then pList.Remove("Rate")
        Dim vName As String = ""
        If pList.ContainsKey("PaymentFrequency") = False Then
          For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
            If vPanelItem.OptionParameterName = "PaymentFrequency" Then
              vName = vPanelItem.OptionParameterName & "_" & vPanelItem.OptionButtonValue
              Exit For
            End If
          Next
          If vName.Length > 0 Then epl.SetErrorField(vName, GetInformationMessage(InformationMessages.ImFieldMandatory))
          vValid = False
        End If

      End If
      If vValid Then
        If epl.GetValue("SortCode").Length > 0 AndAlso epl.GetValue("AccountNumber").Length > 0 Then
          pList("PaymentMethod") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dd)
          If mvBankDetailsNumber > 0 Then pList.IntegerValue("BankDetailsNumber") = mvBankDetailsNumber
          If mvNewBank Then
            pList("NewBank") = "Y"
          Else
            'Existing Bank so remove BranchName as it cannot be changed
            If pList.ContainsKey("BranchName") Then pList.Remove("BranchName")
          End If
        Else
          pList("PaymentMethod") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash)
        End If
        If pList.ContainsKey("StartDate") Then
          If pList.ContainsKey("AccountNumber") Then pList("AutoPayStartDate") = pList("StartDate") 'This is the DD StartDate & only include if we have an AccountNumber
        End If
        pList("StartDate") = AppValues.TodaysDate
        pList("OrderDate") = epl.GetValue("OrderDate")
        If pList.ContainsKey("Amount") Then pList("FixedAmount") = pList("Amount")
        If pList.ContainsKey("DistributionCodeLookupGroup") Then pList.Remove("DistributionCodeLookupGroup")
        If mvIncentiveSequenceList Is Nothing Then mvIncentiveSequenceList = ""
        If mvIncentiveQuantityList Is Nothing Then mvIncentiveQuantityList = ""
        If mvIncentiveSequenceList.Length > 0 Then pList("IncentiveSequence") = mvIncentiveSequenceList
        If mvIncentiveQuantityList.Length > 0 Then pList("IncentiveQuantity") = mvIncentiveQuantityList
        'Get Detail Lines
        Dim vLine As Integer = 1
        For Each vRow As DataRow In PPDDataSet.Tables("DataRow").Rows
          vRow("Arrears") = "0.00"  'Set arrears
          vRow("LineNumber") = vLine  'Set the Line Number of each line
          pList.ObjectValue("PPDLine" & vLine.ToString) = vRow
          vLine += 1
        Next
      End If
    End If
    Return vValid
  End Function

  Friend Overrides Sub RefreshDonationBalance()
    Dim vBalance As Double = 0
    Dim vGrid As DisplayGrid = DirectCast(FindControl(epl, "DetailLines"), DisplayGrid)

    For vIndex As Integer = 0 To vGrid.RowCount - 1
      vBalance = vBalance + CDbl(vGrid.GetValue(vIndex, "Balance"))
    Next

    epl.SetValue("Balance", vBalance.ToString("0.00"))
  End Sub

 
End Class
