Friend Class FDEAddMemberDD
  Inherits CareFDEControl

  Private mvMemberTypes As DataTable
  Private mvMembershipType As String = ""
  Private mvNumberOfMembers As Integer
  Private mvMaxFreeAssociates As Integer
  Private mvBranchMember As String = ""
  Private mvSourceCode As String = ""
  Private mvDistributionCode As String = ""
  Private mvIncentiveScheme As String = ""
  Private mvIncentiveSequenceList As String = ""
  Private mvIncentiveQuantityList As String = ""
  Private mvMemberContact2 As Integer
  Private mvMemberAddress2 As Integer
  Private mvSummaryMembersTable As DataTable

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsSourceChanged = True
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsSourceChanged = True
  End Sub

  Friend Overrides Sub RefreshContactData(ByVal pContactInfo As CDBNETCL.ContactInfo)
    MyBase.RefreshContactData(pContactInfo)
    mvContactInfo.SelectedAddressNumber = mvContactInfo.AddressNumber
    If BooleanValue(epl.GetValue("GiftMembership")) = False Then
      SetValueRaiseChanged(epl, "MemberContactNumber", mvContactInfo.ContactNumber.ToString)
      SetValueRaiseChanged(epl, "MemberAddressNumber", mvContactInfo.SelectedAddressNumber.ToString)
    End If
    If mvMembershipType.Length = 0 Then
      Dim vList As New ParameterList
      vList.FillFromValueList(mvDefaultSettings)
      If vList.ContainsKey("MembershipType") Then mvMembershipType = vList("MembershipType")
    End If
    If mvMembershipType.Length > 0 Then
      SetValueRaiseChanged(epl, "MembershipType_" & mvMembershipType, mvMembershipType)
    End If
    If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctJoint OrElse mvNumberOfMembers = 2 Then
      mvMemberContact2 = 0
      mvMemberAddress2 = 0
      FindContacts()
    End If
  End Sub

  Friend Overrides Sub RefreshAddressData(ByVal pAddressNumber As Integer)
    MyBase.RefreshAddressData(pAddressNumber)
    mvContactInfo.SelectedAddressNumber = pAddressNumber
    If BooleanValue(epl.GetValue("GiftMembership")) = False Then
      SetValueRaiseChanged(epl, "MemberAddressNumber", mvContactInfo.SelectedAddressNumber.ToString)
    End If
  End Sub

  Friend Overrides Sub ResetIncentives()
    MyBase.ResetIncentives()
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
  End Sub

  Friend Overrides Sub SetDefaults()
    epl.FillDeferredCombos(epl)
    MyBase.SetDefaults()

    If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) <> "D" Then
      epl.EnableControl("ClaimDay", False)
    End If

    Dim vValue As String = ""
    Dim vList As New ParameterList
    vList.FillFromValueList(mvDefaultSettings)
    If vList.ContainsKey("MembershipType") Then vValue = vList("MembershipType")
    If vValue.Length > 0 Then
      SetValueRaiseChanged(epl, "MembershipType_" & vValue, vValue)
      mvMembershipType = vValue
    End If
    vValue = ""
    If vList.ContainsKey("PaymentFrequency") Then vValue = vList("PaymentFrequency")
    If vValue.Length > 0 Then SetValueRaiseChanged(epl, "PaymentFrequency_" & vValue, vValue)
    Dim vClaimDay As String = epl.GetValue("ClaimDay")
    vValue = epl.GetValue("BankAccount")
    SetValueRaiseChanged(epl, "BankAccount", vValue)

    If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
      SetValueRaiseChanged(epl, "ClaimDay", vClaimDay)
    End If
    epl.PanelInfo.PanelItems("ClaimDay").Mandatory = False
    If vList.ContainsKey("MandateType") = False Then epl.SetValue("MandateType", "") 'This is the value for 'Unknown'
    vList = New ParameterList(True, True)
    vList.FillFromValueList(mvDefaultSettings)
    vList("FdeUserControl") = mvUserControlName
    Dim vReturnlist As ParameterList = DataHelper.GetFastDataEntryModuleDefaults(vList)
    If vReturnlist.ContainsKey("Joined") Then epl.SetValue("Joined", vReturnlist("Joined"))

    epl.EnableControlList("PackToDonor,PackToMember,OneYearGift", BooleanValue(epl.GetValue("GiftMembership")))
    epl.SetErrorField("ClaimDay", "")
    SetDDStartDate(Today, "StartDate")
    epl.DataChanged = False
  End Sub

  Protected Overrides Sub ProcessMemberValuesChanged(ByVal pEPL As CDBNETCL.EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    MyBase.ProcessMemberValuesChanged(pEPL, pParameterName, pValue)
    Select Case pParameterName
      Case "GiftMembership"
        Dim vGiftMem As Boolean = BooleanValue(pValue)
        With pEPL
          If vGiftMem Then
            '.SetValue("GiftCardStatus_B", "B")
            If mvContactInfo IsNot Nothing AndAlso IntegerValue(.GetValue("MemberContactNumber")) = mvContactInfo.ContactNumber Then
              'Payer is same as member but it is now a gift membership so clear all the member data
              .SetValue("MemberContactNumber", "")
              .SetValue("MemberAddressNumber", "")
              '.SetValue("DateOfBirth", "")
            End If
          Else
            '.SetValue("GiftCardStatus_N", "N")
            .SetValue("OneYearGift", "N")
            .SetValue("PackToDonor", "N")
            '.SetValue("GiverContactNumber", "")
            .SetValue("PackToMember", "N", , , False)
          End If
          .EnableControlList("OneYearGift,PackToDonor,GiftCardStatus_N,GiftCardStatus_B,GiftCardStatus_W,GiverContactNumber,PackToMember", vGiftMem)
          SetEligibleForGiftAid(pEPL)
        End With

      Case "MemberContactNumber"
        mvMemberContact2 = 0
        mvMemberAddress2 = 0
        If mvNumberOfMembers <> 1 OrElse BooleanValue(epl.GetValue("GiftMembership")) Then
          FindContacts()
        End If
        SetEligibleForGiftAid(pEPL)

        If pEPL.GetValue("Joined").Length > 0 AndAlso pValue.Length > 0 Then
          Dim vList As New ParameterList(True)
          vList("Joined") = pEPL.GetValue("Joined")
          vList("ContactNumber") = pValue
          mvMemberTypes = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
          If mvMemberTypes.Rows.Count > 0 Then
            For i As Integer = 0 To pEPL.PanelInfo.PanelItems.Count - 1
              If pEPL.PanelInfo.PanelItems.Item(i).AttributeName = "membership_type" Then
                With pEPL.PanelInfo.PanelItems.Item(i)
                  For Each vRow As DataRow In mvMemberTypes.Rows
                    If vRow.Item(0).ToString = .OptionButtonValue.ToString Then
                      If .Visible Then pEPL.FindPanelControl(Of RadioButton)("MembershipType_" & .OptionButtonValue).Enabled = True
                      Exit For
                    Else
                      pEPL.FindPanelControl(Of RadioButton)("MembershipType_" & .OptionButtonValue).Enabled = False
                    End If
                  Next
                End With
              End If
            Next
          End If
        End If

      Case "MembershipType"
        If pValue.Length > 0 Then
          mvMembershipType = pValue
          If mvMemberTypes Is Nothing Then
            mvMemberTypes = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes)
          End If
          Dim vRow As DataRow = Nothing
          If mvMemberTypes IsNot Nothing Then
            mvMemberTypes.DefaultView.RowFilter = "MembershipType = '" & pValue & "'"
            vRow = mvMemberTypes.DefaultView.ToTable.Rows(0)
          End If
          If vRow IsNot Nothing Then
            With pEPL
              mvNumberOfMembers = IntegerValue(vRow.Item("MembersPerOrder").ToString)
              mvMaxFreeAssociates = IntegerValue(vRow.Item("MaxFreeAssociates").ToString)
              .SetValue("GiftMembership", vRow.Item("PayerRequired").ToString)
              If vRow.Item("PayerRequired").ToString = "M" Then
                '.EnableControl("AffiliatedMemberNumber", True)
                .SetValue("GiftMembership", "N", True)
                .SetValue("OneYearGift", "N", True)
                .SetValue("PackToDonor", "N", True)
                '.SetValue("GiftCardStatus_N", "Y", True)
                '.EnableControl("GiftCardStatus_B", False)
                '.EnableControl("GiftCardStstatus_W", False)
              Else
                '.SetValue("AffiliatedMemberNumber", "", True)
                '.SetErrorField("AffiliatedMemberNumber", "")    'Clears any error
                SetValueRaiseChanged(pEPL, "GiftMembership", vRow.Item("PayerRequired").ToString)
                .EnableControl("GiftMembership", (vRow.Item("PayerRequired").ToString = "N"))
                If FindControl(pEPL, "PackToMember", False) IsNot Nothing Then
                  If BooleanValue(.GetValue("GiftMembership")) = False Then .SetValue("PackToMember", "N")
                  .EnableControl("PackToMember", (.GetValue("GiftMembership") = "Y"))
                End If
              End If
              'pEPL.EnableControl("AffiliatedMemberNumber", (vRow.Item("PayerRequired").ToString = "M"))
              'pEPL.PanelInfo.PanelItems("AffiliatedMemberNumber").Mandatory = (vRow.Item("PayerRequired").ToString = "M")
              mvBranchMember = vRow.Item("BranchMembership").ToString
              '.EnableControl("BranchMember", (vRow.Item("BranchMembership").ToString = "Y"))
              If .GetValue("Branch").Length = 0 Then
                pEPL.SetErrorField("Branch", "", False)
                If BooleanValue(vRow.Item("UsePositionBranch").ToString) AndAlso Not String.IsNullOrEmpty(pEPL.GetValue("MemberContactNumber")) Then
                  'Get Branch from first selected contact position
                  Dim vList As New ParameterList(True, True)
                  vList("Current") = "Y"
                  Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, IntegerValue(pEPL.GetValue("MemberContactNumber")), vList))
                  If Not vDT Is Nothing AndAlso vDT.Rows.Count > 0 Then
                    Dim vAddressNumber As Integer = IntegerValue(vDT.Rows(0).Field(Of String)("AddressNumber"))
                    If vAddressNumber > 0 Then
                      vList = New ParameterList(True, True)
                      vList.IntegerValue("AddressNumber") = vAddressNumber
                      Dim vAddressRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, IntegerValue(pEPL.GetValue("MemberContactNumber")), vList))
                      If Not vAddressRow Is Nothing Then
                        pEPL.SetValue("Branch", vAddressRow("Branch").ToString)
                      End If
                    End If
                  End If
                End If
                If String.IsNullOrEmpty(.GetValue("Branch")) AndAlso Not String.IsNullOrEmpty(pEPL.GetValue("MemberContactNumber")) Then
                  'Get Branch from members address
                  Dim vList As New ParameterList(True, True)
                  vList("ContactNumber") = pEPL.GetValue("MemberContactNumber")
                  vList("AddressNumber") = pEPL.GetValue("MemberAddressNumber")
                  Dim vCRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, IntegerValue(pEPL.GetValue("MemberContactNumber")), vList))
                  If vCRow IsNot Nothing Then
                    Dim vContactInfo As ContactInfo = New ContactInfo(vCRow)
                    pEPL.SetValue("Branch", vContactInfo.Branch)
                  End If
                End If
              End If
              SetValueRaiseChanged(epl, "Product", vRow.Item("FirstPeriodsProduct").ToString, True)
              SetValueRaiseChanged(epl, "Rate", vRow.Item("FirstPeriodsRate").ToString)

              If vRow.Item("MembersPerOrder").ToString = "2" AndAlso mvContactInfo IsNot Nothing Then
                If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctJoint Then
                  'Payer is joint contact so set member to be first selected individual contact
                  Dim vDataTable As DataTable = FormHelper.GetIndividualsFromJointContact(mvContactInfo.ContactNumber)
                  If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
                    .SetValue("MemberContactNumber", vDataTable.Rows(0).Item("ContactNumber").ToString)
                  End If
                End If
                FindContacts()
              End If
              SetEligibleForGiftAid(pEPL)
            End With
          End If
        End If

      Case "Joined"
        If IsDate(pValue) Then
          SetDDStartDate(CDate(pValue), "StartDate")
          SetEligibleForGiftAid(pEPL)
        End If

      Case "OneYearGift", "PackToDonor"
        If pValue = "Y" Then pEPL.SetValue("GiftMembership", "Y")

    End Select
  End Sub

  Protected Overrides Sub ValidateMembersPage(ByVal pEPL As CDBNETCL.EditPanel, ByRef pValid As Boolean)
    MyBase.ValidateMembersPage(pEPL, pValid)
    If mvMemberTypes Is Nothing Then
      mvMemberTypes = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes)
    End If
    Dim vRow As DataRow = Nothing
    If mvMemberTypes IsNot Nothing Then
      mvMemberTypes.DefaultView.RowFilter = "MembershipType = '" & mvMembershipType & "'"
      If mvMemberTypes.DefaultView.ToTable.Rows.Count > 0 Then vRow = mvMemberTypes.DefaultView.ToTable.Rows(0)
    End If

    If vRow Is Nothing Then
      '
    Else
      Dim vContactNumber As Integer = IntegerValue(pEPL.GetValue("MemberContactNumber"))
      Dim vDateOfBirth As String = pEPL.FindPanelControl(Of TextLookupBox)("MemberContactNumber").ContactInfo.DateOfBirth
      Dim vGiftMembership As Boolean = BooleanValue(pEPL.GetValue("GiftMembership"))
      Dim vJoined As String = pEPL.GetValue("Joined")
      Dim vMaxJnrAge As Integer = IntegerValue(vRow.Item("MaxJuniorAge").ToString)
      Dim vDOB As Date
      If vMaxJnrAge > 0 Then
        'Check member is of a suitable age for this membership type
        If Date.TryParse(vDateOfBirth, vDOB) Then
          If DateAdd(DateInterval.Year, DoubleValue((vMaxJnrAge * -1).ToString), Now) > vDOB Then
            pValid = pEPL.SetErrorField("MemberContactNumber", InformationMessages.ImContactTooOldForMembership)
          End If
        End If
      End If
      If pValid AndAlso vRow.Item("MembershipLevel").ToString = "A" Then
        Dim vJoinedDate As Date
        If Date.TryParse(vDateOfBirth, vDOB) = True AndAlso Date.TryParse(vJoined, vJoinedDate) = True Then
          If DateAdd(DateInterval.Year, AppValues.JuniorAgeLimit, vDOB) > vJoinedDate Then
            pValid = pEPL.SetErrorField("MemberContactNumber", String.Format(InformationMessages.ImMemberTooYoung, AppValues.JuniorAgeLimit))
          End If
        End If
      End If
      If pValid = True AndAlso BooleanValue(vRow.Item("SingleMembership").ToString) = True Then
        'Check for contact having other live memberships with single_membership flag set.
        Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMemberships, vContactNumber))
        If vDT IsNot Nothing Then
          vDT.DefaultView.RowFilter = "SingleMembership = 'Y' And len(CancelledOn) = 0"
          If vDT.DefaultView.Count > 0 Then
            pValid = pEPL.SetErrorField("MemberContactNumber", InformationMessages.ImContactAlreadyMember)
          End If
        End If
      End If
      If pValid = True AndAlso vGiftMembership = True Then
        'Check Member/GiverContact are not Payer for gifted membership
        If vContactNumber = mvContactInfo.ContactNumber Then
          pValid = pEPL.SetErrorField("MemberContactNumber", InformationMessages.ImPayerNotMemberForGiftMembership)
        End If
      End If
      If pValid = True AndAlso mvNumberOfMembers <> 1 Then
        If mvMemberContact2 = 0 Then
          pValid = epl.SetErrorField("MemberContactNumber", GetInformationMessage(InformationMessages.ImIncorrectNumberofMembers, mvNumberOfMembers.ToString))
        End If
      End If
    End If

  End Sub

  Friend Overrides Function BuildParameterList(ByRef pList As CDBNETCL.ParameterList) As Boolean
    Dim vValid As Boolean = True
    If CanSubmit() Then
      pList.FillFromValueList(mvInitialSettings)
      ClearOptionButtonError("DistributionCode")
      vValid = MyBase.BuildParameterList(pList)
      If vValid Then
        Dim vName As String = ""
        If pList.ContainsKey("MembershipType") = False Then
          For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
            If vPanelItem.OptionParameterName = "MembershipType" Then
              vName = vPanelItem.OptionParameterName & "_" & vPanelItem.OptionButtonValue
              Exit For
            End If
          Next
          If vName.Length > 0 Then epl.SetErrorField(vName, GetInformationMessage(InformationMessages.ImFieldMandatory))
          vValid = False
        End If
        If pList.ContainsKey("PaymentFrequency") = False Then
          vName = ""
          For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
            If vPanelItem.OptionParameterName = "PaymentFrequency" Then
              vName = vPanelItem.OptionParameterName & "_" & vPanelItem.OptionButtonValue
              Exit For
            End If
          Next
          If vName.Length > 0 Then epl.SetErrorField(vName, GetInformationMessage(InformationMessages.ImFieldMandatory))
          vValid = False
        End If
        If pList.ContainsKey("DistributionCode") = False Then
          vName = "DistributionCode"
          If MyBase.GetMandatoryOptionButton(pList, vName) = False Then
            vValid = epl.SetErrorField(vName, GetInformationMessage(InformationMessages.ImFieldMandatory))
          End If
        End If
      End If
      If vValid Then
        pList("MemberContactNumber") = epl.GetValue("MemberContactNumber")
        pList("MemberAddressNumber") = epl.GetValue("MemberAddressNumber")
        If mvNumberOfMembers = 2 Then
          pList.IntegerValue("MemberContactNumber2") = mvMemberContact2
          pList.IntegerValue("MemberAddressNumber2") = mvMemberAddress2
        End If
        pList("ReasonForDespatch") = pList("MembershipType")
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
        pList("StartDate") = pList("Joined")
        If pList.ContainsKey("DistributionCodeLookupGroup") Then pList.Remove("DistributionCodeLookupGroup")
        If mvIncentiveSequenceList Is Nothing Then mvIncentiveSequenceList = ""
        If mvIncentiveQuantityList Is Nothing Then mvIncentiveQuantityList = ""
        If mvIncentiveSequenceList.Length > 0 Then pList("IncentiveSequence") = mvIncentiveSequenceList
        If mvIncentiveQuantityList.Length > 0 Then pList("IncentiveQuantity") = mvIncentiveQuantityList
      End If
    End If
    Return vValid
  End Function

  Friend Overrides Sub RefreshSource(ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
    MyBase.RefreshSource(pSourceCode, pDistributionCode, pIncentiveScheme)
    mvSourceCode = pSourceCode
    mvDistributionCode = pDistributionCode
    mvIncentiveScheme = pIncentiveScheme
    mvIncentiveSequenceList = ""
    mvIncentiveQuantityList = ""
  End Sub

  Friend Overrides Function CheckIncentives(ByRef pList As ParameterList) As Boolean
    Dim vCheckIncentives As Boolean = MyBase.CheckIncentives(pList)
    Dim vCanSubmit As Boolean = CanSubmit() AndAlso HasCompleted = False
    If vCanSubmit = True AndAlso (mvSourceCode.Length > 0 AndAlso mvIncentiveScheme.Length > 0) AndAlso (mvIncentiveSequenceList.Length = 0 AndAlso mvIncentiveQuantityList.Length = 0) Then
      Dim vList As New ParameterList
      epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly)
      Dim vRFD As String = ""
      If vList.ContainsKey("MembershipType") Then vRFD = vList("MembershipType")
      vCheckIncentives = (vRFD.Length > 0)
      If vCheckIncentives Then
        If pList Is Nothing Then pList = New ParameterList()
        pList("Source") = mvSourceCode
        pList("ReasonForDespatch") = vRFD
        If mvContactInfo IsNot Nothing Then pList("VatCategory") = mvContactInfo.VATCategory
        If epl.GetValue("SortCode").Length > 0 AndAlso epl.GetValue("AccountNumber").Length > 0 Then pList("PayMethodReason") = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.dd_reason)
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
      Dim vJoined As String = epl.GetValue("Joined")
      vCanSubmit = (vJoined.Length > 0)
      If vCanSubmit Then
        Dim vMemberType As String = MyBase.GetOptionButtonValue("MembershipType")
        vCanSubmit = (vMemberType.Length > 0)
      End If
      Return vCanSubmit
    End Get
  End Property

  Private Sub FindContacts()
    If mvMembershipType.Length > 0 Then
      If epl.GetValue("MemberContactNumber").Length > 0 Then
        Dim vList As New ParameterList(True, True)
        vList.IntegerValue("PayerContactNumber") = mvContactInfo.ContactNumber
        vList("ContactNumber") = epl.GetValue("MemberContactNumber")
        vList("AddressNumber") = epl.GetValue("MemberAddressNumber")
        vList("MembershipType") = mvMembershipType
        vList("GiftMembership") = epl.GetValue("GiftMembership")
        vList("Branch") = epl.GetValue("Branch")
        vList("BranchMember") = mvBranchMember
        vList("Joined") = epl.GetValue("Joined")
        vList("Applied") = Today.ToShortDateString
        vList("DistributionCode") = mvDistributionCode
        vList.IntegerValue("NumberOfMembers") = mvNumberOfMembers
        Dim vDS As DataSet = DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipSummaryMembers, 1, vList)
        Dim vDT As DataTable = Nothing
        mvSummaryMembersTable = Nothing
        If vDS IsNot Nothing Then vDT = DataHelper.GetTableFromDataSet(vDS)
        Dim vCount As Integer
        If vDT IsNot Nothing Then
          Dim vMemberContactNumber As Integer = IntegerValue(epl.GetValue("MemberContactNumber"))
          mvSummaryMembersTable = vDT
          For Each vRow As DataRow In vDT.Rows
            vCount += 1
            If mvMemberContact2 = 0 Or mvMemberContact2 = vMemberContactNumber Then
              mvMemberContact2 = IntegerValue(vRow("ContactNumber").ToString)
              mvMemberAddress2 = IntegerValue(epl.GetValue("MemberAddressNumber"))
            End If
          Next
        End If
        If vCount <> mvNumberOfMembers Then
          mvMemberContact2 = 0
          mvMemberAddress2 = 0
          mvSummaryMembersTable = Nothing
          vList("ContactTypeCode") = mvContactInfo.ContactTypeCode
          vList("ContactGroup") = mvContactInfo.ContactGroup
          vList("Source") = mvSourceCode
          Dim vForm As New frmSelectItems(vDS, frmSelectItems.SelectItemsTypes.sitMembershipSummaryMembers, vList)
          If vForm.ShowDialog() = DialogResult.OK Then
            'All members are in the DataSet
            vCount = 0
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDS)
            mvSummaryMembersTable = vTable
            For Each vRow As DataRow In vTable.Rows
              If vCount = 0 Then
                epl.SetValue("MemberContactNumber", vRow("ContactNumber").ToString)
                epl.SetValue("MemberAddressNumber", vRow("AddressNumber").ToString)
              Else
                mvMemberContact2 = IntegerValue(vRow("ContactNumber").ToString)
                mvMemberAddress2 = IntegerValue(vRow("AddressNumber").ToString)
              End If
              vCount += 1
            Next
          End If
        End If
        If mvSummaryMembersTable IsNot Nothing AndAlso mvSummaryMembersTable.DataSet.Tables.Contains("DataRow") AndAlso mvSummaryMembersTable.DataSet.Tables("DataRow").Rows.Count > 0 Then
          If mvSummaryMembersTable.Columns.Contains("SequenceNumber") Then mvSummaryMembersTable.Columns("SequenceNumber").ColumnName = "LineNumber"
          If mvSummaryMembersTable.Columns.Contains("Surname") Then mvSummaryMembersTable.Columns.Remove("Surname")
          If mvSummaryMembersTable.Columns.Contains("Name") Then mvSummaryMembersTable.Columns.Remove("Name")
        End If
        If vCount <> mvNumberOfMembers Then
          epl.SetErrorField("MemberContactNumber", GetInformationMessage(InformationMessages.ImIncorrectNumberofMembers, mvNumberOfMembers.ToString))
        End If
      End If
    End If
  End Sub

  Friend Overrides Sub GetCodeRestrictions(ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList)
    MyBase.GetCodeRestrictions(pParameterName, pList)
    Select Case pParameterName
      Case "Product"
        pList("FindProductType") = "M"        'Membership
    End Select
  End Sub

  Private Sub SetEligibleForGiftAid(ByVal pEPL As EditPanel)
    With pEPL
      If FindControl(pEPL, "EligibleForGiftAid", False) IsNot Nothing Then
        If mvMembershipType.Length > 0 AndAlso .FindPanelControl(Of DateTimePicker)("Joined").Enabled AndAlso IsDate(.GetValue("Joined").ToString) Then
          Dim vAdultMembershipLevel As Boolean = False
          Dim vMTRow As DataRow = Nothing
          If mvMemberTypes IsNot Nothing Then
            mvMemberTypes.DefaultView.RowFilter = "MembershipType = '" & mvMembershipType & "'"
            vMTRow = mvMemberTypes.DefaultView.ToTable.Rows(0)
          End If
          If vMTRow IsNot Nothing AndAlso vMTRow.Item("MembershipLevel") IsNot Nothing AndAlso vMTRow.Item("MembershipLevel").ToString.Length > 0 Then
            vAdultMembershipLevel = (vMTRow.Item("MembershipLevel").ToString = "A")
          End If
          If vAdultMembershipLevel OrElse IntegerValue(.GetValue("MemberContactNumber").ToString) > 0 Then
            Dim vParams As New ParameterList(True)
            vParams("MembershipType") = mvMembershipType
            If Not vAdultMembershipLevel Then vParams("MemberContactNumber") = .GetValue("MemberContactNumber").ToString
            If .GetValue("Branch").ToString.Length > 0 Then vParams("Branch") = .GetValue("Branch").ToString
            vParams("GiftMembership") = .GetValue("GiftMembership").ToString
            vParams("OneYearGift") = .GetValue("OneYearGift").ToString
            vParams("Joined") = .GetValue("Joined").ToString
            If Not vAdultMembershipLevel Then
              If IsDate(.FindPanelControl(Of TextLookupBox)("MemberContactNumber").ContactInfo.DateOfBirth) Then
                vParams("DateOfBirth") = .FindPanelControl(Of TextLookupBox)("MemberContactNumber").ContactInfo.DateOfBirth
              End If
              If mvSummaryMembersTable IsNot Nothing AndAlso
                 mvSummaryMembersTable.DataSet.Tables.Contains("DataRow") AndAlso
                 mvSummaryMembersTable.DataSet.Tables("DataRow").Rows.Count > 0 Then
                For Each vRow As DataRow In mvSummaryMembersTable.Rows
                  vParams.ObjectValue("MemberLine" & vRow("LineNumber").ToString) = vRow
                Next
              End If
            End If
            vParams("TraderTransactionType") = "MEMB"
            vParams("FastDataEntry") = "Y"
            vParams = DataHelper.ProcessTraderPPEligibleForGiftAid(vParams)
            If vParams.Contains("EligibleForGiftAid") Then
              .EnableControl("EligibleForGiftAid", True)
              .SetValue("EligibleForGiftAid", "Y")
            Else
              .SetValue("EligibleForGiftAid", "N", True)
            End If
          Else
            .SetValue("EligibleForGiftAid", "N", True)
          End If
        Else
          .SetValue("EligibleForGiftAid", "N", True)
        End If
      End If
    End With
  End Sub
End Class
