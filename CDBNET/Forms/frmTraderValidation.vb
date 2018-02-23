Partial Public Class frmTrader
  'Trader Validation file
  'Please do not use for anything but trader validation only: Keep it tidy (alphabetical and regioned)
#Region " Validation "

  Private mvSavedMembershipType As String = ""

  Private Sub CheckForCMTPriceChange(ByVal pEPL As EditPanel)
    'Membership has been Renewed; check whether there has been a Rate price increase since Renewals and if so, give user option whether to use Current/New Price.
    If pEPL.GetValue("MembershipType") <> mvTA.CMTPrevMembershipTypeCode Then mvTA.CMTPriceDate = "" 'Reset this in case the MembershipType has changed

    With mvTA.PaymentPlan
      If (.RenewalPending = True AndAlso (.RenewalDate.CompareTo(Now.Date) > 0)) AndAlso (pEPL.GetValue("MembershipType") <> mvTA.CMTPrevMembershipTypeCode) Then
        'Renewals has been run and RenewalDate > Today
        Dim vMemCurrentPrice As Double
        Dim vMemFuturePrice As Double
        Dim vMemPriceChangeDate As Date
        Dim vGotMembershipData As Boolean = False
        Dim vEntitlementPriceChange As Boolean = False

        Dim vList As New ParameterList(True, True)
        vList("MembershipType") = pEPL.GetValue("MembershipType")
        vList("Product") = pEPL.GetValue("Product")
        vList("Rate") = pEPL.GetValue("Rate")
        vList.IntegerValue("MemberContactNumber") = mvTA.CMTMemberContactNumber
        If mvTA.CMTMemberContactNumber = 0 Then vList.IntegerValue("MemberContactNumber") = mvTA.PayerContactNumber
        vList("RenewalDate") = .RenewalDate.ToString(AppValues.DateFormat)
        Dim vPricesDT As DataTable = DataHelper.GetMembershipCMTPrices(vList)
        If vPricesDT IsNot Nothing AndAlso vPricesDT.Rows.Count > 0 Then
          For Each vRow As DataRow In vPricesDT.Rows
            If BooleanValue(vRow.Item("IsEntitlement").ToString) = False Then
              vMemCurrentPrice = DoubleValue(vRow.Item("CurrentPrice").ToString)
              vMemFuturePrice = DoubleValue(vRow.Item("FuturePrice").ToString)
              vMemPriceChangeDate = CDate(vRow.Item("PriceChangeDate").ToString)
              vGotMembershipData = True
            End If
          Next
        End If

        If vGotMembershipData Then
          Dim vEntPriceChangeDate As Date
          For Each vEntRow As DataRow In vPricesDT.Rows 'vDT.Rows
            If BooleanValue(vEntRow.Item("IsEntitlement").ToString) Then
              vEntPriceChangeDate = CDate(vEntRow.Item("PriceChangeDate").ToString)
              'If ((.PaymentMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dd) And vEntRow.Item("AddCondition").ToString = "DD")) OrElse (vEntRow.Item("AddCondition").ToString <> "DD") Then
              If (vEntPriceChangeDate.CompareTo(Now.Date) > 0) AndAlso (DoubleValue(vEntRow.Item("CurrentPrice").ToString) <> DoubleValue(vEntRow.Item("FuturePrice").ToString)) AndAlso (.RenewalDate.CompareTo(vEntPriceChangeDate) >= 0) Then
                'PriceChangeDate > Today and Current & Future Prices are different and RenewalDate >= PriceChangeDate
                vEntitlementPriceChange = True
              End If
            End If
          Next

          Dim vAppParams As frmApplicationParameters
          If (vEntitlementPriceChange = True) OrElse (vMemPriceChangeDate.CompareTo(Now.Date) > 0) AndAlso (.RenewalDate.CompareTo(vMemPriceChangeDate) >= 0) Then
            'Entitlements price changed OR Membership Type PriceChangeDate > Today and RenewalDate >= PriceChangeDate
            vList = New ParameterList()
            vList("MembershipType") = pEPL.GetValue("MembershipType")
            vList("RenewalDate") = .RenewalDate.ToString(AppValues.DateFormat)
            If vEntitlementPriceChange Then
              vAppParams = New frmApplicationParameters(CareServices.FunctionParameterTypes.fptCMTEntitlementPriceChange, vList, Nothing)
            Else
              vList("PriceChangeDate") = vMemPriceChangeDate.ToString(AppValues.DateFormat)
              vList("NewPrice") = vMemFuturePrice.ToString("F")
              vList("CurrentPrice") = vMemCurrentPrice.ToString("F")
              vAppParams = New frmApplicationParameters(CareServices.FunctionParameterTypes.fptCMTPriceChange, vList, Nothing)
            End If
            vAppParams.ShowDialog()
            If vAppParams.DialogResult = System.Windows.Forms.DialogResult.OK Then
              If vAppParams.ReturnList("RunType") = "N" Then
                mvTA.CMTPriceDate = vMemPriceChangeDate.ToString(AppValues.DateFormat)
              End If
            End If
          End If
        End If
      End If
    End With
  End Sub
  Private Shared Function PPFieldsBlank(ByVal pEPL As EditPanel) As Boolean
    Dim vValue As New StringBuilder
    With vValue
      If FindControl(pEPL, "PaymentPlanNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("PaymentPlanNumber"))
      If FindControl(pEPL, "MemberNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("MemberNumber"))
      If FindControl(pEPL, "BankersOrderNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("BankersOrderNumber"))
      If FindControl(pEPL, "DirectDebitNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("DirectDebitNumber"))
      If FindControl(pEPL, "CreditCardAuthorityNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("CreditCardAuthorityNumber"))
      If FindControl(pEPL, "CovenantNumber", False) IsNot Nothing Then .Append(pEPL.GetValue("CovenantNumber"))
    End With
    Return vValue.Length = 0
  End Function
  Private Function ValidateCardNumber(ByVal pEPL As EditPanel, ByVal pParameterName As String) As Boolean
    Dim vValid As Boolean = True
    With pEPL
      If .GetValue(pParameterName).Length > 0 Then
        Dim vCCType As EditPanel.CreditCardValidationTypes = EditPanel.CreditCardValidationTypes.ccvtStandard
        If pParameterName = "CreditCardNumber" AndAlso pEPL.GetValue("AuthorityType_C") = "C" Then vCCType = EditPanel.CreditCardValidationTypes.ccvtCAF
        If mvTA.TransactionPaymentMethod = "CAFC" Then vCCType = EditPanel.CreditCardValidationTypes.ccvtCAF
        Select Case pEPL.ValidCreditCardNumber(pEPL.GetValue(pParameterName), vCCType)
          Case EditPanel.CreditCardValidationStatus.ccvsInvalidNumber
            vValid = .SetErrorField(pParameterName, InformationMessages.ImInvalidCardNumber, True)
          Case EditPanel.CreditCardValidationStatus.ccvsNotNumeric
            vValid = .SetErrorField(pParameterName, InformationMessages.ImCardNumberNotNumeric, True)
          Case EditPanel.CreditCardValidationStatus.ccvsValid
            .SetErrorField(pParameterName, "")
        End Select
      End If
    End With
    Return vValid
  End Function
  Private Function ValidateCMT() As Boolean
    Dim vEPL As EditPanel = Nothing
    Dim vList As New ParameterList(True, True)
    Dim vRow As DataRow = Nothing
    Dim vParamName As String = "MembershipType"
    Dim vValid As Boolean = True

    'Figure out the correct field to place any error against
    vEPL = mvCurrentPage.EditPanel
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then
      If vEPL.GetValue("MemberNumber").Length > 0 Then
        vParamName = "MemberNumber"
      ElseIf vEPL.GetValue("PaymentPlanNumber").Length > 0 Then
        vParamName = "PaymentPlanNumber"
      ElseIf vEPL.GetValue("BankersOrderNumber").Length > 0 Then
        vParamName = "BankersOrderNumber"
      ElseIf vEPL.GetValue("CreditCardAuthorityNumber").Length > 0 Then
        vParamName = "CreditCardAuthorityNumber"
      ElseIf vEPL.GetValue("DirectDebitNumber").Length > 0 Then
        vParamName = "DirectDebitNumber"
      ElseIf vEPL.GetValue("CovenantNumber").Length > 0 Then
        vParamName = "CovenantNumber"
      Else
        vParamName = "ContactNumber"
      End If
    End If

    vEPL = mvCurrentPage.EditPanel
    If mvTA.PaymentPlan.PlanType <> PaymentPlanInfo.ppType.pptMember Then
      vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTNotAMembership, True)
    Else
      Dim vMembershipNumber As Integer = 0
      If vParamName = "MemberNumber" AndAlso vEPL.FindTextLookupBox("MemberNumber").GetDataRow IsNot Nothing Then
        vMembershipNumber = vEPL.FindTextLookupBox("MemberNumber").GetDataRowInteger("MembershipNumber")
      End If
      If vMembershipNumber = 0 Then
        'Find the MembershipNumber
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then
          Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMemberships, mvTA.CMTMemberContactNumber, vList))
          If vDT IsNot Nothing Then
            For Each vDataRow As DataRow In vDT.Rows
              If mvTA.PaymentPlan.PaymentPlanNumber = IntegerValue(vDataRow.Item("PaymentPlanNumber").ToString) Then
                If vMembershipNumber > 0 Then vValid = False
                'If vValid Then mvTA.SetCMTValues(mvTA.CMTMemberContactNumber, IntegerValue(vDataRow.Item("MembershipNumber").ToString))
                vMembershipNumber = IntegerValue(vDataRow.Item("MembershipNumber").ToString)
                If vValid = False Then Exit For
              End If
            Next
          Else
            vValid = False
          End If
          If vValid = False Then
            vEPL.SetErrorField(vParamName, InformationMessages.ImNoUniqueMembership, True)
          End If
        End If
      End If

      If vValid = True AndAlso mvCurrentPage.PageType = CareNetServices.TraderPageType.tpContactSelection Then
        If vEPL.GetValue("ContactNumber").Length > 0 Then
          If vEPL.FindTextLookupBox("ContactNumber").ContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctJoint Then
            vValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImJointContactCannotBeMember, True)
          End If
        End If
      End If

      If vValid = True AndAlso vMembershipNumber > 0 Then
        vList.IntegerValue("MembershipNumber") = vMembershipNumber
        vRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, mvTA.CMTMemberContactNumber, vList))
      End If

      If vRow IsNot Nothing Then
        With mvTA.PaymentPlan
          'Check for UnprocessedPayments
          If .Balance > 0 AndAlso .UnprocessedPayments > 0 Then
            vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTUnprocessedPayments, True)
          End If
          'Check if renewal needs to be run
          If vValid AndAlso (.RenewalPeriodEnd < Date.Today) Then
            vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTRenewalRequired, True)
          End If
          'Check if it is a group membership
          If vValid = True AndAlso (IntegerValue(vRow.Item("MembersPerOrder").ToString) = 0 AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_me_allow_group_change) = False) Then
            vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTCannotChangeGroupMship, True)
          End If
          'Check if renewed for future change
          If vValid Then
            If (.PayPlanMembershipTypeCode <> vRow.Item("MembershipType").ToString) Then
              vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTFutureChangeRenewal, True)
            ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection AndAlso vRow.Item("FutureMembershipType").ToString.Length > 0 Then
              If ShowQuestion(QuestionMessages.QmCMTRemoveFutureMembership, MessageBoxButtons.YesNo, vRow.Item("FutureMembershipTypeDesc").ToString, vRow.Item("FutureChangeDate").ToString) = System.Windows.Forms.DialogResult.No Then
                vValid = vEPL.SetErrorField(vParamName, InformationMessages.ImCMTTutureChangeExists, True)
              End If
            End If
          End If
        End With
      End If

      If mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType AndAlso mvTA.PaymentPlan.CMTProportionBalance <> PaymentPlanInfo.CMTProportionBalanceTypes.cmtNone Then
        If (mvTA.PaymentPlan.PayPlanMembershipTypeCode = vEPL.GetValue("MembershipType")) AndAlso (mvTA.PaymentPlan.MembershipRateCode = vEPL.GetValue("Rate")) Then
          vValid = vEPL.SetErrorField("Rate", InformationMessages.ImCMTSameMemberTypeAndRate, True)
        End If
      End If
    End If

    'Jira 1455: When CMT validation error occurs force trader to stay on the current page (prevent moving to Transaction Analysis page)
    mvIsInvalidCMT = Not vValid

    Return vValid

  End Function

  Private Sub ValidateDefaults()
    With mvCurrentPage
      Select Case .PageType
        Case CareServices.TraderPageType.tpPaymentPlanDetails
          EPL_ValidateItem(.EditPanel, "PaymentFrequency", .EditPanel.GetValue("PaymentFrequency"), True)
      End Select
    End With
  End Sub

  Private Function ValidateEventBooking(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True
    Dim vWaitingList As Boolean
    Dim vInterested As Boolean

    Dim vContactInfo As ContactInfo = mvCurrentPage.EditPanel.FindTextLookupBox("ContactNumber").ContactInfo
    Dim vEventInfo As CareEventInfo = mvCurrentPage.EditPanel.FindTextLookupBox("EventNumber").CareEventInfo
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    If IntegerValue(vEPL.GetValue("BookingNumber")) > 0 Then
      '17680: Where AllocationsChecked = True an error occurred with invoice allocations returned to the client. This question will have already been asked in previous iteration so prevent re-display.
      If mvTA.AllocationsChecked = False Then
        vValid = ShowQuestion(QuestionMessages.QmCancelOriginalBooking, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes
      End If
    End If

    If vValid Then
      If vEventInfo.EligibilityCheckRequired Then
        Dim vForm As New frmEventEligibility(vContactInfo, vEventInfo, mvTA.TransactionSource)
        If vForm.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then
          Return False
        End If
      End If

      vWaitingList = vEPL.GetOptionalValue("WaitingList") = "Y"
      vInterested = vEPL.GetOptionalValue("InterestOnly") = "Y"

      If Not vEventInfo.ChargeForWaiting Then
        Dim vEBPrice As Double = DoubleValue(vEPL.GetValue("Amount"))
        If vWaitingList And vEBPrice > 0 Then
          'User has checked Waiting List; only allow this if paying by CS/CC
          If mvTA.PayMethodsAtEnd = False AndAlso mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CARD" AndAlso
             mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN" Then
            vValid = False
            ShowInformationMessage(InformationMessages.ImUnchargedWLPayMethod)
          End If
        End If
      End If

      Dim vRow As DataRow = mvCurrentPage.EditPanel.FindTextLookupBox("OptionNumber").GetDataRow
      If vRow("PickSessions").ToString = "Y" Then
        vEventInfo.BookingOptionDesc = vRow("OptionDesc").ToString
        vEventInfo.PickSessionsCount = IntegerValue(vRow("NumberOfSessions").ToString)
        vEventInfo.BookingOptionNumber = IntegerValue(vRow("OptionNumber").ToString)
        Dim vForm As New frmEventSet(Me, vEventInfo, CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions)
        If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
          pList("SessionNumbers") = vForm.SelectedSessions.CSList
        Else
          vValid = False
        End If
      End If
    End If
    If vValid Then
      Dim vBookingType As String
      If vInterested Then
        vBookingType = "Y"
        pList("BookingStatus") = "I"
      Else
        Dim vBookingMessage As String
        Dim vList As New ParameterList(True)
        vList("Quantity") = vEPL.GetValue("Quantity")
        vList("OptionNumber") = vEPL.GetValue("OptionNumber")
        If FindControl(vEPL, "WaitingList", False) IsNot Nothing Then vList("WaitingList") = vEPL.GetValue("WaitingList")
        vList("EventNumber") = vEPL.GetValue("EventNumber")
        If pList.ContainsKey("SessionNumbers") Then vList("SessionNumbers") = pList("SessionNumbers")

        Dim vDataSet As DataSet = DataHelper.CheckEventBooking(vList)
        If vDataSet IsNot Nothing Then
          vBookingType = Trim(vDataSet.Tables("Result").Rows(0)("CanAddBooking").ToString)
          vBookingMessage = vDataSet.Tables("Result").Rows(0)("BookingMessage").ToString
          Select Case vBookingType
            Case "W"
              If Not vWaitingList Then
                If ShowQuestion(vBookingMessage, MessageBoxButtons.YesNo) <> System.Windows.Forms.DialogResult.Yes Then
                  vBookingType = "N"
                  vValid = False
                Else
                  If Not vEventInfo.ChargeForWaiting AndAlso DoubleValue(vEPL.GetValue("Amount")) > 0 AndAlso mvTA.PayMethodsAtEnd = False AndAlso
                     mvTA.TransactionPaymentMethod <> "CRED" AndAlso mvTA.TransactionPaymentMethod <> "CARD" AndAlso
                     mvTA.TransactionPaymentMethod <> "CQIN" AndAlso mvTA.TransactionPaymentMethod <> "CCIN" Then
                    'User answered Yes in response to Waiting List prompt; only allow this if paying by CS/CC
                    vValid = False
                    vBookingType = "N"
                    ShowInformationMessage(InformationMessages.ImUnchargedWLPayMethod)
                  Else
                    vValid = True
                    vWaitingList = True 'Set this flag to True as this will be used to set the amount as 0.
                    pList("WaitingList") = "Y"
                    If FindControl(vEPL, "WaitingList", False) IsNot Nothing Then vEPL.SetValue("WaitingList", "Y")
                  End If
                End If
              End If
            Case "N"
              ShowInformationMessage(vBookingMessage)
              vValid = False
          End Select
        Else
          vValid = True
        End If
      End If

      'If vBookingType = "N" Then
      '  vStatus = ""
      'Else
      '  If Not vInterested Then
      '    If vEventInfo.External Then
      '      vStatus = CareEventInfo.EventBookingStatuses.ebsExternal
      '    Else
      '      If vBookingType = "W" Then
      '        pList("WaitingList") = "Y"
      '        Select Case vStatus
      '          Case gvEnv.GetBookingStatusCode(ebsBookedAndPaid)
      '            pStatus = gvEnv.GetBookingStatusCode(ebsWaitingPaid)
      '          Case gvEnv.GetBookingStatusCode(ebsBookedCreditSale)
      '            pStatus = gvEnv.GetBookingStatusCode(ebsWaitingCreditSale)
      '          Case gvEnv.GetBookingStatusCode(ebsBookedInvoiced)                '            pStatus = gvEnv.GetBookingStatusCode(ebsWaitingInvoiced)
      '          Case Else
      '            pStatus = gvEnv.GetBookingStatusCode(ebsWaiting)
      '        End Select
      '      Else
      '        pWaiting = False
      '        Desired Status is supplied by Trader but not booking via Events.
      '        If Len(pStatus) = 0 Then pStatus = gvEnv.GetBookingStatusCode(ebsBooked)
      '      End If
      '    End If
      '  End If
      '  vContact.Init(gvEnv, pContactNumber, pAddressNumber)
      '  vEventBooking = vEvent.AddEventBooking(vContact, pAddressNumber, pQty, pOption, gvEnv.GetBookingStatus(pStatus), pRate, pSessionList, pNotes, , , , pConvertInterestedBooking)
      '  If vEventBooking Is Nothing Then
      '    pMsg = vEvent.LastBookingMessage
      '    BookEvent = True                'Error
      '  Else
      '    pBookingNumber = vEventBooking.BookingNumber
      '  End If
      'End If
      'Exit Function
      If pList.ContainsKey("WaitingList") AndAlso pList("WaitingList") = "Y" Then
        If mvTA.TransactionPaymentMethod = "CRED" Then
          pList("BookingStatus") = "A"
        Else
          pList("BookingStatus") = "P"
        End If
      ElseIf pList.ContainsKey("InterestOnly") AndAlso pList("InterestOnly") = "Y" Then
        pList("BookingStatus") = "I"
      ElseIf mvTA.TransactionPaymentMethod = "CRED" Then
        pList("BookingStatus") = "S"
      Else
        pList("BookingStatus") = "B"
      End If

      If vValid AndAlso vWaitingList AndAlso Not vEventInfo.ChargeForWaiting Then
        If IntegerValue(vEPL.GetValue("Amount")) > 0 Then
          mvEventWLPriceZeroed = True
          vEPL.SetValue("Amount", "0")
          'set the amount in the param list to 0 as well, as the param list would already have been populated by now
          If pList.ContainsKey("Amount") Then pList.IntegerValue("Amount") = 0
        End If
      End If
      If vValid AndAlso vInterested Then
        vEPL.SetValue("Amount", "0")
        'set the amount in the param list to 0 as well, as the param list would already have been populated by now
        If pList.ContainsKey("Amount") Then pList.IntegerValue("Amount") = 0
      End If
    End If

    Return vValid
  End Function

  Private Function ValidateExamBooking(ByVal pList As ParameterList) As Boolean
    'TODO EXAMS 
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    Dim vControl As Control = FindControl(vEPL, "ExamUnitId", False)
    If vControl IsNot Nothing Then
      Dim vChangedList As List(Of ChangedItem) = DirectCast(vControl, ExamSelector).GetChangedList()
      If vChangedList.Count = 0 Then
        ShowInformationMessage("No Items have been selected")
        Return False
      End If

      ' Validate Exam Session Home/Overseas Closing Date
      Dim vExamCentreCodeLookUp As TextLookupBox = mvCurrentPage.EditPanel.FindTextLookupBox("ExamCentreCode")
      If vEPL.GetValue("ExamCentreCode").Length > 0 AndAlso vExamCentreCodeLookUp.IsValid AndAlso vExamCentreCodeLookUp.GetDataRow() IsNot Nothing Then
        Dim vClosingDateStr As String = vExamCentreCodeLookUp.GetDataRowItem("ClosingDate")
        Dim vClosingDate As Date
        If Date.TryParse(vClosingDateStr, vClosingDate) Then
          If vClosingDate < Date.Today Then
            If ShowQuestion(QuestionMessages.QmConfirmExamBookingSessionExpired, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then Return False
          End If
        End If
      End If

      Dim vUnits As New StringBuilder
      Dim vUnitLinks As New StringBuilder
      Dim vAddComma As Boolean
      For Each vChange As ChangedItem In vChangedList
        If vAddComma Then vUnits.Append(",") : vUnitLinks.Append(",")
        vUnits.Append(vChange.Item.UnitID)
        vUnitLinks.Append(vChange.Item.LinkID)
        vAddComma = True
      Next
      pList("ExamUnits") = vUnits.ToString
      pList("ExamUnitLinks") = vUnitLinks.ToString
    End If
    Dim vUnit As TextLookupBox = vEPL.FindTextLookupBox("ExamUnitCode")
    pList.IntegerValue("ExamUnitId") = vUnit.GetDataRowInteger("ExamUnitId")
    pList.IntegerValue("ExamUnitLinkId") = vUnit.GetDataRowInteger("ExamUnitLinkId")
    pList("Amount") = vEPL.GetValue("Amount")

    Dim vDataSet As DataSet = ExamsDataHelper.GetExamEligibilityChecks(pList)
    Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
    If Not vDataTable Is Nothing Then
      If vDataTable.Rows.Count > 0 Then
        Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litExamEligibility)
        If vForm.ShowDialog() = DialogResult.Cancel Then Return False
      End If
    End If
    Dim vSession As TextLookupBox = vEPL.FindTextLookupBox("ExamSessionCode")
    pList.IntegerValue("ExamSessionId") = vSession.GetDataRowInteger("ExamSessionId")
    Dim vCentre As TextLookupBox = vEPL.FindTextLookupBox("ExamCentreCode")
    pList.IntegerValue("ExamCentreId") = vCentre.GetDataRowInteger("ExamCentreId")
    Return True
  End Function

  Private Function ValidateAccommodationBooking(ByVal pList As ParameterList) As Boolean
    Dim vValid As Boolean = True

    If mvTA.TransactionPaymentMethod = "CRED" Then
      pList("BookingStatus") = "S"
    Else
      pList("BookingStatus") = "B"
    End If

    Return vValid
  End Function

  Private Function ValidateCollectionPayment() As Boolean
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    Dim vAmount As Double = 0
    Dim vCollectorNumber As Integer = 0
    Dim vCollectionType As CampaignItem.AppealTypes
    Dim vPayAmount As Double = DoubleValue(vEPL.GetValue("Amount"))
    Dim vPisCollector As Integer = 0
    Dim vValid As Boolean = True

    Select Case vEPL.FindTextLookupBox("AppealCollectionNumber").GetDataRowItem("CollectionType")
      Case "H"
        vCollectionType = CampaignItem.AppealTypes.atH2HCollection
      Case "M"
        vCollectionType = CampaignItem.AppealTypes.atMannedCollection
      Case Else
        vCollectionType = CampaignItem.AppealTypes.atUnMannedCollection
    End Select

    'Set Collectors
    If vCollectionType = CampaignItem.AppealTypes.atMannedCollection OrElse vCollectionType = CampaignItem.AppealTypes.atH2HCollection Then
      vCollectorNumber = IntegerValue(vEPL.GetValue("DeceasedContactNumber"))
      If vCollectorNumber = 0 Then vCollectorNumber = mvTA.PayerContactNumber
    End If

    Dim vPay As Object
    If vCollectionType = CampaignItem.AppealTypes.atMannedCollection Then
      For vIndex As Integer = 0 To mvCBXDGR.RowCount - 1
        vPay = mvCBXDGR.GetValue(vIndex, "Pay")   'This will come back as either Nothing or True
        If vPay IsNot Nothing AndAlso CBool(vPay) = True Then
          vAmount += DoubleValue(GetStringValue(mvCBXDGR.GetValue(vIndex, "Amount")))
        End If
      Next
      If vPayAmount <> vAmount Then
        vValid = vEPL.SetErrorField("Amount", String.Format(InformationMessages.ImPaymentAmountNotBoxAmount, vPayAmount, vAmount))
      End If
      If vValid Then
        'Validate Collector
        For vIndex As Integer = 0 To mvCBXDGR.RowCount - 1
          vPay = mvCBXDGR.GetValue(vIndex, "Pay")   'This will come back as either Nothing or True
          If vPay IsNot Nothing AndAlso CBool(vPay) = True Then
            vPisCollector = IntegerValue(GetStringValue(mvCBXDGR.GetValue(vIndex, "ContactNumber")))
            If vPisCollector <> vCollectorNumber Then vValid = False
          End If
          If vValid = False Then Exit For
        Next
        If vValid = False Then
          vEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImPayerNotBoxCollector)
        End If
      End If
    Else
      'Unmanned (Boxes are optional) / House-2-House Collection (No boxes)
      Dim vColPISNumber As Integer = IntegerValue(vEPL.GetValue("PisNumber"))
      Dim vList As New ParameterList(True, True)
      If vColPISNumber > 0 Then vList.IntegerValue("CollectionPisNumber") = vColPISNumber

      If (vCollectionType = CampaignItem.AppealTypes.atUnMannedCollection AndAlso vColPISNumber < 1) OrElse (vCollectionType = CampaignItem.AppealTypes.atH2HCollection AndAlso vColPISNumber < 0) Then
        vValid = vEPL.SetErrorField("PisNumber", "Invalid Entry", True)
      ElseIf vCollectionType = CampaignItem.AppealTypes.atUnMannedCollection Then
        vList("CollectionNumber") = vEPL.GetValue("AppealCollectionNumber")
        Dim vStringAmount As String = ""
        Dim vDR As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, vList))
        If vDR IsNot Nothing Then
          vStringAmount = vDR.Item("Amount").ToString
        End If
        If vStringAmount.Length > 0 Then
          vAmount = DoubleValue(vStringAmount)
          If vPayAmount <> vAmount Then
            If ShowQuestion(QuestionMessages.QmPaymentAmountNotPISAmount, MessageBoxButtons.YesNo, vPayAmount.ToString, vAmount.ToString) = System.Windows.Forms.DialogResult.No Then
              vValid = vEPL.SetErrorField("Amount", String.Format(InformationMessages.ImPaymentAmountNotPISAmount, vPayAmount, vAmount))
            End If
          End If
        End If
      End If
      If vValid = True Then
        If vCollectionType = CampaignItem.AppealTypes.atUnMannedCollection Then
          'Check only 1 collection box has been selected
          Dim vCount As Integer = 0
          For vIndex As Integer = 0 To mvCBXDGR.RowCount - 1
            vPay = mvCBXDGR.GetValue(vIndex, "Pay")   'This will come back as either Nothing or True
            If vPay IsNot Nothing AndAlso CBool(vPay) = True Then
              vCount += 1
            End If
          Next
          If vCount > 1 Then vValid = vEPL.SetErrorField("AppealCollectionNumber", InformationMessages.ImUnmanndedCollectionOneCollectionBox)
        ElseIf vCollectionType = CampaignItem.AppealTypes.atH2HCollection Then
          vList("CollectionNumber") = vEPL.GetValue("AppealCollectionNumber")
          Dim vDR As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectionPIS, vList))
          If vDR IsNot Nothing Then
            vPisCollector = IntegerValue(vDR.Item("ContactNumber").ToString)
          End If
          If vPisCollector > 0 AndAlso (vPisCollector <> vCollectorNumber) Then
            vValid = vEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImPayerNotPISCollector)
          End If
        End If
      End If
    End If

    Return vValid

  End Function

  Private Function ValidateMember(Optional ByVal pGridRow As Integer = -1) As Boolean
    'tpMembershipMembersSummary page will use pGridRow to validate each row
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel
    Dim vRow As DataRow
    Dim vAgeOverride As String
    Dim vBranchMember As Boolean
    Dim vContactNumber As Integer
    Dim vControlName As String
    Dim vDateOfBirth As String
    Dim vGiftMembership As Boolean
    Dim vGiverContact As String = ""
    Dim vJoined As String
    Dim vValid As Boolean = True

    If mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary Then
      Dim vList As New ParameterList(True, True)
      vList("MembershipType") = mvMembersDGR.GetValue(pGridRow, "MembershipType")
      vRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList).Rows(0)
      With mvMembersDGR
        vAgeOverride = .GetValue(pGridRow, "AgeOverride")
        vBranchMember = BooleanValue(.GetValue(pGridRow, "BranchMember"))
        vContactNumber = IntegerValue(.GetValue(pGridRow, "ContactNumber"))
        vDateOfBirth = .GetValue(pGridRow, "DateOfBirth")
        vGiftMembership = BooleanValue(.GetValue(pGridRow, "GiftMembership"))
        vJoined = .GetValue(pGridRow, "Joined")
      End With
    Else
      vRow = vEPL.FindTextLookupBox("MembershipType").GetDataRow
      With vEPL
        vAgeOverride = .GetValue("AgeOverride")
        vBranchMember = BooleanValue(.GetValue("BranchMember"))
        vContactNumber = IntegerValue(.GetValue("ContactNumber"))
        vDateOfBirth = .GetValue("DateOfBirth")
        vJoined = .GetValue("Joined")
        If mvCurrentPage.PageType <> CareServices.TraderPageType.tpAmendMembership Then
          vGiftMembership = BooleanValue(.GetValue("GiftMembership"))
          vGiverContact = .GetValue("GiverContactNumber")
        End If
      End With
    End If

    If vRow IsNot Nothing Then
      Dim vMaxJnrAge As Integer = IntegerValue(vRow.Item("MaxJuniorAge").ToString)
      Dim vDOB As Date
      If vMaxJnrAge > 0 Then
        'Check member is of a suitable age for this membership type
        If Date.TryParse(vDateOfBirth, vDOB) Then
          If vAgeOverride.Length > 0 Then
            If IntegerValue(vAgeOverride) > vMaxJnrAge Then
              vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "AgeOverride").ToString
              vValid = vEPL.SetErrorField(vControlName, String.Format(InformationMessages.ImAgeOverrideInvalid, vMaxJnrAge))
            End If
          Else
            If DateAdd(DateInterval.Year, DoubleValue((vMaxJnrAge * -1).ToString), Now) > vDOB Then
              vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "ContactNumber").ToString
              vValid = vEPL.SetErrorField(vControlName, InformationMessages.ImContactTooOldForMembership)
            End If
          End If
        Else
          vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "DateOfBirth").ToString
          vValid = vEPL.SetErrorField(vControlName, InformationMessages.ImDOBMustBeSpecifiedForMembership)
        End If
      End If
      If vValid AndAlso vRow.Item("MembershipLevel").ToString = "A" Then
        Dim vJoinedDate As Date
        Dim vContinue As Boolean
        If Date.TryParse(vDateOfBirth, vDOB) = True Then
          If mvTA.TransactionType = "MEMC" Then
            vJoinedDate = Today
            vContinue = True
          Else
            vContinue = Date.TryParse(vJoined, vJoinedDate)
          End If
          If vContinue = True AndAlso DateAdd(DateInterval.Year, AppValues.JuniorAgeLimit, vDOB) > vJoinedDate Then
            vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "DateOfBirth").ToString
            vValid = vEPL.SetErrorField(vControlName, String.Format(InformationMessages.ImMemberTooYoung, AppValues.JuniorAgeLimit))
          End If
        End If
      End If
      If vValid And mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary Then
        If mvMembersDGR.GetValue(pGridRow, "Branch").Length = 0 Then
          vValid = vEPL.SetErrorField("CurrentMembers", InformationMessages.ImBranchMembershipNotAvailable)
        End If
      End If
      If vValid AndAlso BooleanValue(vRow.Item("BranchMembership").ToString) = False Then
        'BranchMembership can only be set of membership type allows branch mwmbership
        vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "BranchMember").ToString
        If vBranchMember Then vValid = vEPL.SetErrorField(vControlName, InformationMessages.ImBranchMembershipNotAvailable)
      End If
      If vValid = True AndAlso BooleanValue(vRow.Item("SingleMembership").ToString) = True Then
        'Check for contact having other live memberships with single_membership flag set.
        Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMemberships, vContactNumber))
        If vDT IsNot Nothing Then
          vDT.DefaultView.RowFilter = "SingleMembership = 'Y' And len(CancelledOn) = 0"
          If vDT.DefaultView.Count > 0 Then
            If (mvTA.TransactionType = "MEMC" AndAlso IntegerValue(vDT.DefaultView.Item(0)("PaymentPlanNumber").ToString) <> mvTA.PaymentPlan.PaymentPlanNumber) OrElse mvTA.TransactionType <> "MEMC" Then
              'CMT - Only error if the PaymentPlanNumber is different
              'Not CMT - always error
              vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "ContactNumber").ToString
              vValid = vEPL.SetErrorField(vControlName, InformationMessages.ImContactAlreadyMember)
            End If
          End If
        End If
      End If
      If vValid = True AndAlso vGiftMembership = True Then
        'Check Member/GiverContact are not Payer for gifted membership
        If vContactNumber = mvTA.PayerContactNumber Then
          vControlName = IIf(mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary, "CurrentMembers", "ContactNumber").ToString
          vValid = vEPL.SetErrorField(vControlName, InformationMessages.ImPayerNotMemberForGiftMembership)
        End If
        If vValid AndAlso vGiverContact.Length > 0 Then
          If (IntegerValue(vGiverContact) = vContactNumber) OrElse (IntegerValue(vGiverContact) = mvTA.PayerContactNumber) Then
            vValid = vEPL.SetErrorField("GiverContactNumber", InformationMessages.ImGiverCannotBeMemberOrPayer)
          End If
        End If
      End If
      If vValid = True AndAlso vRow.Item("PayerRequired").ToString = "M" Then
        'Check AffiliatedMember is Payer
        If (mvCurrentPage.PageType <> CareServices.TraderPageType.tpMembershipMembersSummary) AndAlso (mvCurrentPage.PageType <> CareServices.TraderPageType.tpAmendMembership) Then
          If vEPL.GetValue("AffiliatedMemberNumber").Length > 0 Then
            Dim vAMRow As DataRow = vEPL.FindTextLookupBox("AffiliatedMemberNumber").GetDataRow
            Dim vContactNo As Integer
            If vAMRow IsNot Nothing Then
              vContactNo = IntegerValue(vAMRow.Item("ContactNumber").ToString)
              If vContactNo <> mvTA.PayerContactNumber Then
                'Affiliated Member is the payer
                mvTA.SetPayerContact(vContactNo, IntegerValue(vAMRow.Item("AddressNumber").ToString))
              End If
            End If
            If vContactNo = 0 OrElse mvTA.PayerContactNumber = 0 Then
              vValid = vEPL.SetErrorField("AffiliatedMemberNumber", String.Format(InformationMessages.ImPayerCannotBeFoundFromMember, vEPL.GetValue("AffiliatedMemberNumber")))
            End If
          End If
        End If
      End If
      If vValid AndAlso vRow.Table.Columns.Contains("PricesCount") AndAlso vRow.Item("PricesCount").ToString.Length > 0 AndAlso mvTA.PayMethodsAtEnd Then
        vValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImMembershipPricesPayMethodsAtEnd)
      End If
    End If

    If vValid AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpMembershipMembersSummary Then
      vEPL.SetErrorField("CurrentMembers", String.Empty)
    End If

    Return vValid

  End Function

  Private Function ValidateMemberNumber(ByVal pMemberNumber As String) As Boolean
    Dim vChar As String
    Dim vCharPos As Integer
    Dim vIntPos As Integer
    Dim vLen As Integer
    Dim vValid As Boolean = True
    Dim vPos As Integer

    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.check_member_number) Then
      vLen = pMemberNumber.Length
      If vLen > 0 Then
        'Check member number starts with a character between A and Z
        vChar = pMemberNumber.Substring(0, 1)
        If IsNumeric(vChar) Then vValid = False
        If vValid Then
          If vChar < "A" Or vChar > "Z" Then vValid = False
        End If

        If vValid Then
          vPos = 1  'Start at 2nd character
          'Find position of first number (Member Number must contain at least 1 number)
          While vPos <= (vLen - 1) And vIntPos = 0
            vChar = pMemberNumber.Substring(vPos, 1)
            If IsNumeric(vChar) Then vIntPos = vPos
            vPos = vPos + 1
          End While
          'Error if we have 9 characters without any numbers
          If vIntPos = 0 And (vPos - 1 = 8) Then vValid = False

          If vValid Then
            'Ensure that no characters appear after the first number
            While vPos <= (vLen - 1) And vCharPos = 0
              vChar = pMemberNumber.Substring(vPos, 1)
              If IsNumeric(vChar) = False Then vCharPos = vPos
              vPos = vPos + 1
            End While
            If vCharPos > 0 Then vValid = False
          End If
        End If
      End If
    End If

    Return vValid

  End Function

  Private Function ValidatePage() As Boolean
    'This should only be used for validations that cannot be done under EPL_ValidateALLItems
    Dim vValid As Boolean = True
    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpScheduledPayments
        If mvTA.TransactionType <> "LOAN" Then
          'as these fields are labels, they never get validated under epl_validateallitems
          Dim vPPBalance As Double = DoubleValue(mvCurrentPage.EditPanel.GetValue("Balance"))
          If vPPBalance <> DoubleValue(mvCurrentPage.EditPanel.GetValue("AmountOutstanding")) Then
            vValid = mvCurrentPage.EditPanel.SetErrorField("AmountOutstanding", GetInformationMessage(InformationMessages.ImOPSOutstandingMustEqualBalance, FixTwoPlaces(vPPBalance).ToString))
          End If
        End If
    End Select
    Return vValid
  End Function

  Private Function ValidateProductSale() As Boolean
    'Validate stock product sales when 'Next' is clicked
    Dim vValid As Boolean = True

    If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.StockSales = True Then
      Dim vEPL As EditPanel = mvCurrentPage.EditPanel
      Dim vWarehouse As String = ""
      If FindControl(vEPL, "Warehouse", False) IsNot Nothing Then vWarehouse = vEPL.GetValue("Warehouse")
      'First ensure that all StockMovements have been created
      If mvTA.StockValuesChanged(vEPL.GetValue("Product"), vWarehouse, IntegerValue(vEPL.GetValue("Quantity")), True) Then
        vValid = AddStockMovement(vEPL.GetValue("Product"), vWarehouse, IntegerValue(vEPL.GetValue("Quantity")))
      End If

      If vValid Then
        Dim vWhen As String = vEPL.GetValue("When")
        If Not IsDate(vWhen) Then vWhen = mvTA.TransactionDate
        'Secondly, check for Delivery date in the future
        Dim vDeliveryDate As Date = Date.Parse(vWhen)
        If vDeliveryDate.CompareTo(Now.Date) > 0 Then
          'DeliveryDate is after Today
          If ShowQuestion(QuestionMessages.QmStockFutureDeliveryDate, MessageBoxButtons.YesNo, mvTA.StockIssued.ToString) = System.Windows.Forms.DialogResult.Yes Then
            vValid = AddStockMovement(vEPL.GetValue("Product"), vWarehouse, mvTA.StockIssued, True, True)
          Else
            vValid = vEPL.SetErrorField("When", InformationMessages.ImStockCannotBeFutureDate)
          End If
        End If
      Else
        vValid = vEPL.SetErrorField("Quantity", GetInformationMessage(InformationMessages.ImInsufficientStock), True)
      End If
    End If
    Return vValid
  End Function

  Private Sub ValidateQuantity(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean)
    Dim vRate As String = ""
    pEPL.SetErrorField(pParameterName, "")
    ProductValidation.ValidateQuantity(pEPL, pParameterName, pValue, pValid, mvCurrentPage.PageType, vRate)

    If vRate.Length > 0 Then
      If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking AndAlso pEPL.GetValue("Product").Length = 0 Then Exit Sub
      SetValueRaiseChanged(pEPL, "Rate", vRate)
    End If

    If pValid = True AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.StockSales = True Then
      Dim vWarehouseCode As String = ""
      If FindControl(pEPL, "Warehouse", False) IsNot Nothing Then vWarehouseCode = pEPL.GetValue("Warehouse")
      If AddStockMovement(pEPL.GetValue("Product"), vWarehouseCode, IntegerValue(pValue)) Then
        'Refresh Warehouses and stock count
        If Len(vWarehouseCode) > 0 Then
          GetWarehouses(pEPL, pEPL.GetValue("Product"))
          pEPL.SetValue("Warehouse", vWarehouseCode)
        End If
        Dim vProductList As New ParameterList(True, True)
        vProductList("Product") = pEPL.GetValue("Product")
        vProductList("FindProductType") = "S"
        vProductList("SystemColumns") = "N"
        If FindControl(pEPL, "LastStockCount", False) IsNot Nothing Then pEPL.SetValue("LastStockCount", mvTotalStock.ToString)
      Else
        pValid = pEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImInsufficientStock), True)
      End If
    End If
    If pValid Then
      SetAmount(pEPL)       'Update the amount since the quantity has changed
    End If
    If pValid Then
      If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts Then
        If pEPL.GetValue("CommunicationNumber").Length > 0 And IntegerValue(pValue) > 1 Then
          pEPL.SetErrorField("Quantity", InformationMessages.ImQuantity1ForSubscriptions)
          pValid = False
        End If
      End If
    End If
  End Sub

  Private Sub ValidatePTPGTotals(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean)
    Dim vAmount As Double = DoubleValue(pValue)

    If mvCurrentPage.PageType = CareServices.TraderPageType.tpGiveAsYouEarn Then
      If pParameterName = "DonorTotal" Then     'Got the donor_total, work out the government_total
        If vAmount > 0 Then
          Dim vGovernmentTotal As Double = vAmount * (DoubleValue(AppValues.ControlValue(AppValues.ControlValues.government_percentage)) / 100)
          If pEPL.GetDoubleValue("GovernmentTotal") <> vGovernmentTotal Then pEPL.SetValue("GovernmentTotal", vGovernmentTotal.ToString("0.00"))
        End If
      ElseIf pParameterName = "GovernmentTotal" Then
        If vAmount > 0 Then
          'Got the government_total, work out the donor_total
          If DoubleValue(AppValues.ControlValue(AppValues.ControlValues.government_percentage)) > 0 Then
            vAmount = vAmount * (100 / DoubleValue(AppValues.ControlValue(AppValues.ControlValues.government_percentage)))
          Else
            vAmount = vAmount
          End If
          If pEPL.GetDoubleValue("DonorTotal") <> vAmount Then pEPL.SetValue("DonorTotal", vAmount.ToString("0.00"))
        End If
      End If
      pEPL.SetValue("Amount", "")     'Only populate the amount field after all 4 gaye fields populated
      If pEPL.GetValue("DonorTotal").Length > 0 Then
        vAmount = pEPL.GetDoubleValue("DonorTotal")
        If pEPL.GetValue("EmployerTotal").Length > 0 Then
          vAmount += pEPL.GetDoubleValue("EmployerTotal")
          If pEPL.GetValue("GovernmentTotal").Length > 0 Then
            vAmount += pEPL.GetDoubleValue("GovernmentTotal")
            If pEPL.GetValue("AdminFeesTotal").Length > 0 Then
              vAmount += pEPL.GetDoubleValue("AdminFeesTotal")
              pEPL.SetValue("Amount", vAmount.ToString("0.00"))
            End If
          End If
        End If
      End If
    ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpPostTaxPGPayment Then
      'Just need to populate the Transaction Total
      pEPL.SetValue("Amount", "")
      vAmount = pEPL.GetDoubleValue("DonorTotal")
      vAmount += pEPL.GetDoubleValue("EmployerTotal")
      pEPL.SetValue("Amount", vAmount.ToString("0.00"))
    End If
  End Sub

  Private Function ValidatePledge(ByVal pEPL As EditPanel, ByVal pTLB As TextLookupBox) As Boolean
    Dim vTransactionDate As String = pEPL.GetValue("TransactionDate")
    If vTransactionDate.Length > 0 Then
      Dim vStartDate As String = pTLB.GetDataRowItem("StartDate")
      Dim vCancelledOn As String = pTLB.GetDataRowItem("CancelledOn")
      If IsDate(vStartDate) Then
        If CDate(vTransactionDate) < CDate(vStartDate) Then
          pEPL.SetErrorField("TransactionDate", String.Format("The transaction date is before the pledge start date of {0}", vStartDate), True)
          Return False
        End If
      End If
      If IsDate(vCancelledOn) Then
        If CDate(vTransactionDate) > CDate(vCancelledOn) Then
          pEPL.SetErrorField("TransactionDate", String.Format("The transaction date is after the pledge cancellation date of {0}", vCancelledOn), True)
          Return False
        End If
      End If
    End If
    Return True
  End Function

  Private Sub ValidateDeceasedContact(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Dim vInMemoriam As Boolean
    Select Case pParameterName
      Case "DeceasedContactNumber"
        'Only came here because it is InMemoriam
        vInMemoriam = True
      Case "LineType"
        vInMemoriam = (pValue = "G")
      Case "LineTypeG"
        vInMemoriam = BooleanValue(pValue)
      Case Else
        If pParameterName.Length >= 8 AndAlso pParameterName.StartsWith("LineType") Then
          'LineTypeH, LineTypeS
          vInMemoriam = BooleanValue(pEPL.GetValue("LineTypeG"))
        End If
    End Select

    If vInMemoriam Then
      Try
        Dim vDecdContactNumber As String = pEPL.GetValue("DeceasedContactNumber")
        If vDecdContactNumber.Length > 0 Then
          Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, IntegerValue(vDecdContactNumber))
          Dim vDeceasedStatus As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.deceased_status)
          If vRow IsNot Nothing AndAlso vRow("Status").ToString <> vDeceasedStatus Then
            pEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImDeceasedContact, True)
          Else
            pEPL.SetErrorField("DeceasedContactNumber", "")
          End If
        End If
      Catch vCareException As CareException
        If vCareException.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
          pEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImCannotFindContact, True)
        End If
      End Try
    End If
  End Sub

  Public Function ValidateTransactionPaymentMethod(ByVal pTransactionPaymentMethod As String) As Boolean
    Select Case pTransactionPaymentMethod
      Case "CAFC"
        If mvTA.OnlineCCAuthorisation Then
          ShowInformationMessage(InformationMessages.ImTransactionPayCAFCCInvalid) 'CAF Card payment method invalid when using On-Line Credit Card Authorisation
          Return False
        End If
      Case "CARD", "CCIN"
        If Not mvTA.OnlineCCAuthorisation Then
          ShowInformationMessage(InformationMessages.ImTransactionPayCardInvalid) 'Card payment method invalid when not using On-Line Credit Card Authorisation
          Return False
        End If
    End Select
    Return True
  End Function

#End Region

#Region " Value Changed "

  Private Sub GetServiceModifiers(ByVal pSelect As Boolean, ByVal pCreateNew As Boolean)
    Dim vBookingContact As TextLookupBox = Nothing
    Dim vServiceControl As DataRow = Nothing
    Dim vList As ParameterList = Nothing
    Dim vSelectionResult As DialogResult

    With mvCurrentPage.EditPanel
      If .FindPanelControl("BookingContactNumber", False) IsNot Nothing AndAlso
        .FindPanelControl("ContactGroup", False) IsNot Nothing AndAlso
        .FindPanelControl("RelatedContactNumber", False) IsNot Nothing AndAlso
        .FindPanelControl("Quantity", False) IsNot Nothing Then

        .SetValue("RelatedContactNumber", String.Empty)
        .SetValue("Quantity", String.Empty)

        vBookingContact = .FindTextLookupBox("BookingContactNumber")

        vList = New ParameterList(True)
        vList("ContactGroup") = .GetValue("ContactGroup")
        If vList("ContactGroup").Length > 0 Then
          vServiceControl = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtServiceControl, vList)
        End If

        If vBookingContact.HasValidValue AndAlso vServiceControl IsNot Nothing Then
          Dim vDefaultTime As String = String.Empty
          Dim vDateControl As Control = .FindPanelControl("StartDate", False)
          Dim vDate As Date
          'Set default start date using value defined in the service control table
          If vDateControl IsNot Nothing AndAlso vServiceControl("DefaultStartTime").ToString.Length > 0 Then
            vDate = CType(vDateControl, DateTimePicker).Value
            vDefaultTime = vServiceControl("DefaultStartTime").ToString
            vDate = New Date(vDate.Year, vDate.Month, vDate.Day, IntegerValue(Strings.Left(vDefaultTime, 2)), IntegerValue(Strings.Right(vDefaultTime, 2)), 0)
            .SetValue("StartDate", vDate.ToString(AppValues.DateTimeFormat))
          End If

          'Set default end date using value defined in the service control table
          vDateControl = .FindPanelControl("EndDate", False)
          If vDateControl IsNot Nothing AndAlso vServiceControl("DefaultEndTime").ToString.Length > 0 Then
            vDate = CType(vDateControl, DateTimePicker).Value
            vDefaultTime = vServiceControl("DefaultEndTime").ToString
            vDate = New Date(vDate.Year, vDate.Month, vDate.Day, IntegerValue(Strings.Left(vDefaultTime, 2)), IntegerValue(Strings.Right(vDefaultTime, 2)), 0)
            .SetValue("EndDate", vDate.ToString(AppValues.DateTimeFormat))
          End If

          If vServiceControl("ModifierGroup").ToString.Length > 0 AndAlso vServiceControl("ModifierRelationship").ToString.Length > 0 Then
            vList("Relationship") = vServiceControl("ModifierRelationship").ToString
            vList("ContactGroup") = vServiceControl("ModifierGroup").ToString
            Dim vContactLinks As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, IntegerValue(vBookingContact.Text), vList)

            'mvResultIndex = vRelatedIndex
            Dim vContactLink As DataRow = Nothing
            Dim vLinksTable As DataTable = DataHelper.GetTableFromDataSet(vContactLinks)
            If vLinksTable IsNot Nothing Then vContactLink = vLinksTable.Rows(0)

            If vContactLink IsNot Nothing Then
              If pSelect Then
                vList = New ParameterList(True)
                Dim vFrmSelect As New frmSelectListItem(vContactLinks, frmSelectListItem.ListItemTypes.litConGroup)
                vSelectionResult = vFrmSelect.ShowDialog(Me)
                If vSelectionResult = DialogResult.OK Then
                  Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vContactLinks)
                  .SetValue("RelatedContactNumber", vTable.Rows(vFrmSelect.SelectedRow)("ContactNumber").ToString)
                End If
              Else
                .SetValue("RelatedContactNumber", vContactLink("ContactNumber").ToString)
              End If
            Else
              'A value of DialogResult.No in frmSelectListItem means that the New button is clicked
              'Display the form to create a new link as there are no records for the 
              vSelectionResult = DialogResult.No
            End If

            If vSelectionResult = DialogResult.No And pCreateNew Then
              vList = New ParameterList
              Dim vEntityGroup As New EntityGroup(vServiceControl("ModifierGroup").ToString, EntityGroup.EntityGroupTypes.egtContactGroup)
              vList("ContactGroup") = vEntityGroup.Code
              Dim vContactNumber As Integer = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctContact, vList, Me)
              If vContactNumber > 0 Then
                'Create a relation between the booker and the newly added contact
                .SetValue("RelatedContactNumber", vContactNumber.ToString)
                vList = New ParameterList(True)
                vList("ContactNumber") = .GetValue("BookingContactNumber")
                vList("ContactNumber2") = vContactNumber.ToString
                vList("Relationship") = vServiceControl("ModifierRelationship").ToString
                DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
              End If
            End If

            'Fetch data from contact categories table to get the quantity for the modifier
            Dim vModifierQty As DataRow = Nothing
            vList = New ParameterList(True)
            vList("ContactNumber") = .GetValue("RelatedContactNumber")
            If vList("ContactNumber").Length > 0 Then
              vList("Activity") = vServiceControl("ModifierActivity").ToString
              vList("ActivityValue") = vServiceControl("ModifierActivityValue").ToString
              vModifierQty = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtContactCategories, vList)
            End If

            If vModifierQty IsNot Nothing Then
              .SetValue("Quantity", vModifierQty("Quantity").ToString)
            Else
              .SetValue("Quantity", "1")
            End If
          Else
            .SetValue("Quantity", "1")
          End If
        Else
          If pSelect Then
            If Not vBookingContact.HasValidValue Then
              .SetErrorField("BookingContactNumber", InformationMessages.ImSelectTheContactMakingTheBooking, True)    'Select the contact making the booking
            Else
              .SetErrorField("ContactGroup", InformationMessages.ImSelectTheServiceBeingSold, True)    'Select the service being sold
            End If
          End If
        End If
      End If
    End With
  End Sub

  Private Function CardAuthorisationRequired(ByVal pTransactionAmount As Double) As Boolean
    Dim vCCType As EditPanel.CreditCardValidationTypes = EditPanel.CreditCardValidationTypes.ccvtStandard
    If mvTA.TransactionPaymentMethod = "CAFC" Then vCCType = EditPanel.CreditCardValidationTypes.ccvtCAF
    Dim vFloorLimit As Double
    If vCCType = EditPanel.CreditCardValidationTypes.ccvtStandard Then
      vFloorLimit = DoubleValue(AppValues.ControlValue(AppValues.ControlValues.noncaf_floor_limit))
    Else
      vFloorLimit = DoubleValue(AppValues.ControlValue(AppValues.ControlValues.caf_floor_limit))
    End If
    Return (pTransactionAmount > vFloorLimit)
  End Function

  Private Sub ClearPPNumberChanged(ByVal pEPL As EditPanel, ByVal pExcludeParam As String)
    With pEPL
      If .GetValue(pExcludeParam).Length > 0 Then
        If pExcludeParam <> "MemberNumber" Then .SetValue("MemberNumber", "", False, False, False, True)
        If pExcludeParam <> "CovenantNumber" Then .SetValue("CovenantNumber", "", False, False, False, True)
        If pExcludeParam <> "BankersOrderNumber" Then .SetValue("BankersOrderNumber", "", False, False, False, True)
        If pExcludeParam <> "DirectDebitNumber" Then .SetValue("DirectDebitNumber", "", False, False, False, True)
        If pExcludeParam <> "CreditCardAuthorityNumber" Then .SetValue("CreditCardAuthorityNumber", "", False, False, False, True)
        If mvCurrentPage.PageType <> CareServices.TraderPageType.tpPayments Then
          'Do not clear PaymentPlanNumber if we are on the Payments page
          If pExcludeParam <> "PaymentPlanNumber" Then .SetValue("PaymentPlanNumber", "", False, False, False, True)
        End If
      End If
    End With
  End Sub

  Private Function GetContactNumberFromDataTable(ByVal pDataTable As DataTable) As Integer
    Dim vContactNumber As Integer = 0

    If pDataTable.Rows.Count > 0 Then
      Dim vRow As DataRow = pDataTable.Rows(0)    'Always get the first row (could be multiple rowrs)
      If vRow IsNot Nothing Then vContactNumber = IntegerValue(vRow.Item("ContactNumber").ToString)
    End If

    Return vContactNumber
  End Function
  Private Function GetMembershipPrices(ByVal pEPL As EditPanel) As ParameterList
    Dim vList As New ParameterList(True)
    If pEPL.GetValue("MembershipType").Length > 0 Then
      Dim vTLB As TextLookupBox = pEPL.FindTextLookupBox("MembershipType")
      If vTLB IsNot Nothing AndAlso vTLB.IsValid Then vList("MembershipType") = pEPL.GetValue("MembershipType") 'Ensure MembershipType is valid first
    End If
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership Then
      vList("PaymentMethod") = mvTA.PPPaymentMethod
    Else  'CMT - assume PaymentPlanInfo object is instantiated
      vList("PaymentMethod") = mvTA.PaymentPlan.PaymentMethod
    End If
    If pEPL.GetOptionalValue("PaymentFrequency").Length > 0 Then vList("PaymentFrequency") = pEPL.GetValue("PaymentFrequency")
    If mvTA.TransactionType = "MEMC" Then
      vList.IntegerValue("ContactNumber") = mvTA.CMTMemberContactNumber
      vList.IntegerValue("AddressNumber") = mvTA.CMTMemberAddressNumber
    Else
      If pEPL.GetOptionalValue("ContactNumber").Length > 0 Then vList("ContactNumber") = pEPL.GetValue("ContactNumber")
      If pEPL.GetOptionalValue("AddressNumber").Length > 0 Then vList("AddressNumber") = pEPL.GetValue("AddressNumber")
    End If
    If vList.Contains("MembershipType") AndAlso vList.Contains("PaymentMethod") AndAlso vList.Contains("PaymentFrequency") Then
      Dim vRestriction As New StringBuilder
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipPrices, vList)
      If vTable IsNot Nothing Then
        vList.Clear()
        For Each vRow As DataRow In vTable.Rows
          'construct an IN list 
          If vRestriction.Length > 0 Then vRestriction.Append(", ")
          vRestriction.Append("'")
          vRestriction.Append(vRow.Item("Rate").ToString)
          vRestriction.Append("'")
          If vRow.Item("Concessionary").ToString = "N" AndAlso Not vList.ContainsKey("PrimaryRate") Then vList("PrimaryRate") = vRow.Item("Rate").ToString
        Next
        If vRestriction.Length > 0 AndAlso vList.Contains("PrimaryRate") Then
          vRestriction.Insert(0, "Rate IN (")
          vRestriction.Append(")")
        ElseIf vRestriction.Length > 0 Then
          vRestriction.Length = 0
        End If
      End If
      pEPL.FindTextLookupBox("Rate").SetFilter(vRestriction.ToString)
    End If
    Return vList
  End Function
  Private Sub GetSpecialPrice(ByVal pEPL As EditPanel, ByVal pProduct As String)
    Dim vList As New ParameterList(True)
    vList("Product") = pProduct
    vList("ContactNumber") = mvTA.PayerContactNumber.ToString
    vList("AddressNumber") = mvTA.PayerAddressNumber.ToString
    Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtSpecialPrices, vList)
    If vRow IsNot Nothing Then SetValueRaiseChanged(pEPL, "Rate", vRow("Rate").ToString, True)
  End Sub

  Private Sub GetWarehouses(ByVal pEPL As EditPanel, ByVal pProduct As String)
    Dim vList As New ParameterList(True)
    vList("Product") = pProduct
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtProductWarehouses, vList))
    Dim vCombo As ComboBox = pEPL.FindComboBox("Warehouse")
    With vCombo
      If vTable IsNot Nothing Then
        .ValueMember = "Warehouse"
        .DisplayMember = "WarehouseStock"
        .DataSource = vTable
        mvTotalStock = 0
        For Each vRow As DataRow In vTable.Rows
          mvTotalStock = mvTotalStock + IntegerValue(vRow.Item("LastStockCount").ToString)
        Next
      Else
        .DataSource = Nothing
        .SelectedText = ""
      End If
    End With
  End Sub

  Private Sub ProcessMembersPageValuesChanged(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Select Case pParameterName
      Case "Amount"
        If mvTA.LinePrice = 0 Then ResetMembershipDetails(pEPL, pParameterName, pValue)
      Case "ContactNumber"
        'Member
        If pValue.Length > 0 Then
          Dim vMTCode As String = pEPL.GetValue("MembershipType")
          Dim vList As New ParameterList(True, True)
          vList("ContactNumber") = pValue
          vList("AddressNumber") = pEPL.GetValue("AddressNumber")
          Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, IntegerValue(pValue), vList))
          If vRow IsNot Nothing Then
            Dim vContactInfo As ContactInfo = New ContactInfo(vRow)
            pEPL.SetValue("DobEstimated", CBoolYN(vContactInfo.DOBEstimated))
            pEPL.SetValue("DateOfBirth", vContactInfo.DateOfBirth)
          End If
          SetRateFromMembershipPrices(pEPL)
          Dim vJoined As String = pEPL.GetValue("Joined")
          If vJoined.Length > 0 Then
            vList.Add("Joined", vJoined)
            pEPL.FindTextLookupBox("MembershipType").ComboBox.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList, True)
          End If
          pEPL.SetValue("MembershipType", "")
          SetValueRaiseChanged(pEPL, "MembershipType", vMTCode)
        End If
      Case "AddressNumber"
        If pValue.Length > 0 Then
          SetRateFromMembershipPrices(pEPL)
        End If
      Case "DobEstimated"
        If BooleanValue(pValue) Then
          If Not (Date.TryParse(pEPL.GetValue("DateOfBirth"), Nothing)) Then pEPL.SetValue("DateOfBirth", New Date(1901, 1, 1).ToString)
        End If
      Case "GiftMembership"
        Dim vGiftMem As Boolean = BooleanValue(pValue)
        With pEPL
          If vGiftMem Then
            If mvTA.TransactionType = "MEMC" Then
              .SetValue("GiftCardStatus_N", "N")      'CMT
            Else
              .SetValue("GiftCardStatus_B", "B")
              If IntegerValue(.GetValue("ContactNumber")) = mvTA.PayerContactNumber Then
                'Payer is same as member but it is now a gift membership so clear all the member data
                .SetErrorField("ContactNumber", "")    'Clear any error
                .SetValue("ContactNumber", "")
                .SetValue("AddressNumber", "")
                .SetValue("DateOfBirth", "")
              End If
            End If
            If BooleanValue(.GetValue("PackToDonor")) = False Then
              'The Pack To Donor checkbox may already be checked.
              'It could be that the Gift Membership checkbox has become checked BECAUSE Pack To Donor was checked.
              .SetValue("PackToDonor", CBoolYN(mvTA.PackToDonorDefault))
            End If
          Else
            .SetValue("GiftCardStatus_N", "N")
            .SetValue("OneYearGift", "N")
            .SetValue("PackToDonor", "N")
            .SetValue("GiverContactNumber", "")
            .SetValue("PackToMember", "N", , , False)
          End If
          .EnableControlList("OneYearGift,PackToDonor,GiftCardStatus_N,GiftCardStatus_B,GiftCardStatus_W,GiverContactNumber,PackToMember", vGiftMem)
        End With
      Case "GiftCardStatus"
        If mvTA.TransactionType = "MEMB" Then
          If pValue = "W" Then
            pEPL.EnableControlList("GiftFrom,GiftTo,GiftMessage", True)
          Else
            pEPL.EnableControlList("GiftFrom,GiftTo,GiftMessage", False)
          End If
        End If
      Case "Joined"
        ResetMembershipDetails(pEPL, pParameterName, pValue)
        If pValue.Length > 0 Then
          Dim vContactNumber As String = ""
          If mvTA.PayerContactNumber.ToString.Length > 0 Then
            vContactNumber = mvTA.PayerContactNumber.ToString
          ElseIf pEPL.GetValue("ContactNumber").Length > 0 Then
            vContactNumber = pEPL.GetValue("ContactNumber")
          End If
          If vContactNumber.Length > 0 Then
            Dim vMTCode As String = pEPL.GetValue("MembershipType")
            Dim vList As New ParameterList(True, True)
            vList.Add("ContactNumber", vContactNumber)
            vList.Add("Joined", pValue)
            Dim vUseTransitions As Boolean
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType Then
              Dim vDT As DataTable = DataHelper.MembershipTypeTransitionsTableRestricted(mvTA.PaymentPlan.PayPlanMembershipTypeCode, vList)
              If (vDT Is Nothing OrElse vDT.Rows.Count = 0) Then
                'no transitions have been found - check if there are tranistions for this membership type without limiting through memberhsip categories
                Dim vTable1 As DataTable = DataHelper.MembershipTypeTransitionsTable(mvTA.PaymentPlan.PayPlanMembershipTypeCode)
                If vTable1.Rows.Count = 0 Then
                  'none have been found so need to change the datasource
                  vDT = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList, True)
                Else
                  vUseTransitions = True
                End If
              End If
              pEPL.FindTextLookupBox("MembershipType").ComboBox.DataSource = vDT
            Else
              pEPL.FindTextLookupBox("MembershipType").ComboBox.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList, True)
            End If
            If vUseTransitions = False AndAlso vMTCode.Length > 0 Then
              pEPL.SetValue("MembershipType", "")
              SetValueRaiseChanged(pEPL, "MembershipType", vMTCode)
            End If
          End If
        End If
        If mvTA.TransactionType = "MEMC" Then
          'Changing Joined needs to re-calculate the Price if RateModifiers are being used
          If FindControl(pEPL, "Rate", False) IsNot Nothing Then
            Dim vRow As DataRow = pEPL.FindTextLookupBox("Rate").GetDataRow
            If vRow IsNot Nothing AndAlso BooleanValue(vRow.Item("UseModifiers").ToString) = True Then
              If pValue.Length = 0 Then pValue = AppValues.TodaysDate
              mvTA.LinePrice = DataHelper.GetModifierPrice(vRow("Product").ToString, vRow("Rate").ToString, CDate(pValue), IntegerValue(mvTA.PayerContactNumber))
              pEPL.SetValue("Amount", "")
              SetAmount(pEPL)
            End If
          End If
        End If
      Case "MembershipType"
        If mvSavedMembershipType <> pValue Or String.IsNullOrEmpty(pEPL.GetValue("Product")) OrElse
          String.IsNullOrWhiteSpace(pEPL.GetValue("Product")) Then
          If ResetMembershipDetails(pEPL, pParameterName, pValue) Then
            Dim vRow As DataRow = pEPL.FindTextLookupBox("MembershipType").GetDataRow
            If vRow IsNot Nothing Then
              With pEPL
                If pValue.Length > 0 Then
                  .SetValue("NumberOfMembers", "")
                  .SetValue("MaxFreeAssociates", "")
                  .SetValue("NumberOfMembers", vRow.Item("MembersPerOrder").ToString)
                  .SetValue("MaxFreeAssociates", IntegerValue(vRow.Item("MaxFreeAssociates").ToString).ToString)
                  If mvTA.TransactionType = "MEMC" Then
                    'CMT
                    .SetValue("GiftMembership", IIf(mvTA.PaymentPlan.GiftMembership = True, "Y", "N").ToString)
                    If mvTA.PaymentPlan.GiftMembership = False AndAlso vRow.Item("PayerRequired").ToString = "Y" Then
                      .SetValue("GiftMembership", "Y")
                    End If
                  Else
                    .SetValue("GiftMembership", vRow.Item("PayerRequired").ToString)
                  End If
                  If vRow.Item("PayerRequired").ToString = "M" Then
                    .EnableControl("AffiliatedMemberNumber", True)
                    .SetValue("GiftMembership", "N", True)
                    .SetValue("OneYearGift", "N", True)
                    .SetValue("PackToDonor", "N", True)
                    .SetValue("GiftCardStatus_N", "Y", True)
                    .EnableControl("GiftCardStatus_B", False)
                    .EnableControl("GiftCardStstatus_W", False)
                  Else
                    .SetValue("AffiliatedMemberNumber", "", True)
                    .SetErrorField("AffiliatedMemberNumber", "")    'Clears any error
                    If mvTA.TransactionType <> "MEMC" Then SetValueRaiseChanged(pEPL, "GiftMembership", vRow.Item("PayerRequired").ToString)
                    .EnableControl("GiftMembership", (vRow.Item("PayerRequired").ToString = "N"))
                    If FindControl(pEPL, "PackToMember", False) IsNot Nothing Then
                      If BooleanValue(.GetValue("GiftMembership")) = False Then .SetValue("PackToMember", "N")
                      .EnableControl("PackToMember", (.GetValue("GiftMembership") = "Y"))
                    End If
                  End If
                  pEPL.EnableControl("AffiliatedMemberNumber", (vRow.Item("PayerRequired").ToString = "M"))
                  pEPL.PanelInfo.PanelItems("AffiliatedMemberNumber").Mandatory = (vRow.Item("PayerRequired").ToString = "M")
                  If mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership AndAlso FindControl(pEPL, "PaymentFrequency", False) IsNot Nothing Then
                    If vRow.Table.Columns.Contains("PricesCount") AndAlso IntegerValue(vRow.Item("PricesCount").ToString) > 0 Then
                      .PanelInfo.PanelItems("PaymentFrequency").Mandatory = True
                    Else
                      .PanelInfo.PanelItems("PaymentFrequency").Mandatory = False
                      .SetErrorField("PaymentFrequency", "")   'Clear any error
                    End If
                  End If
                End If
                If mvTA.TransactionType = "MEMC" Then
                  'CMT
                  If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_renew_at_same_rate) = True OrElse mvTA.PaymentPlan.DetermineMembershipPeriod() <> PaymentPlanInfo.MembershipPeriodTypes.mptSubsequentPeriod Then
                    .SetValue("Product", vRow.Item("FirstPeriodsProduct").ToString, True)
                    .SetValue("Rate", vRow.Item("FirstPeriodsRate").ToString)
                  Else
                    .SetValue("Product", vRow.Item("SubsequentPeriodsProduct").ToString, True)
                    .SetValue("Rate", vRow.Item("SubsequentPeriodsRate").ToString)
                  End If
                  If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_renew_at_same_rate) = True Then SetRateFromMembershipPrices(pEPL)

                  'in most circumstances CMT will always use the original join date so we'll enable/disable the join date accordingly
                  Dim vOldMemTypeRow As DataRow = Nothing
                  If pValue <> mvTA.CMTPrevMembershipTypeCode Then
                    Dim vOldMTList As New ParameterList(True, True)
                    vOldMTList.Add("MembershipType", mvTA.PaymentPlan.PayPlanMembershipTypeCode)
                    vOldMemTypeRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vOldMTList).Rows(0)
                  End If
                  If vOldMemTypeRow IsNot Nothing Then
                    Dim vNewJoinedDate As String = Nothing
                    If (vOldMemTypeRow.Item("VotingRights").ToString = "N" AndAlso vRow("VotingRights").ToString = "Y") _
                    OrElse ((mvTA.PaymentPlan.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsNew) > 0 AndAlso DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(AppValues.TodaysDate), CDate(mvTA.PaymentPlan.RenewalDate)) >= 0) Then
                      .EnableControl("Joined", True)
                      If AppValues.ConfigurationValue(AppValues.ConfigurationValues.me_blank_joined_date).ToString = "Y" Then
                        vNewJoinedDate = ""
                      Else
                        vNewJoinedDate = AppValues.TodaysDate
                      End If
                    Else
                      'Following lines are commented to Fix BR15621 
                      '.EnableControl("Joined", False)
                      'If mvTA.CMTOriginalMemberJoined.Length > 0 Then vNewJoinedDate = mvTA.CMTOriginalMemberJoined
                    End If
                    If vNewJoinedDate IsNot Nothing AndAlso vNewJoinedDate <> .GetValue("Joined") Then
                      .SetValue("Joined", vNewJoinedDate)
                      ProcessMembersPageValuesChanged(pEPL, "Joined", vNewJoinedDate)
                    End If
                  End If
                Else
                  .SetValue("BranchMember", vRow.Item("BranchMembership").ToString)
                  .EnableControl("BranchMember", (vRow.Item("BranchMembership").ToString = "Y"))
                  If .GetValue("Branch").Length = 0 Then
                    pEPL.SetErrorField("Branch", "", False)
                    If BooleanValue(vRow.Item("UsePositionBranch").ToString) AndAlso Not String.IsNullOrEmpty(pEPL.GetValue("ContactNumber")) Then
                      'Get Branch from first selected contact position
                      Dim vList As New ParameterList(True, True)
                      vList("Current") = "Y"
                      Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, IntegerValue(pEPL.GetValue("ContactNumber")), vList))
                      If Not vDT Is Nothing AndAlso vDT.Rows.Count > 0 Then
                        Dim vAddressNumber As Integer = IntegerValue(vDT.Rows(0).Field(Of String)("AddressNumber"))
                        If vAddressNumber > 0 Then
                          vList = New ParameterList(True, True)
                          vList.IntegerValue("AddressNumber") = vAddressNumber
                          Dim vAddressRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, IntegerValue(pEPL.GetValue("ContactNumber")), vList))
                          If Not vAddressRow Is Nothing Then
                            pEPL.SetValue("Branch", vAddressRow("Branch").ToString)
                          End If
                        End If
                      End If
                    End If
                    If String.IsNullOrEmpty(.GetValue("Branch")) AndAlso Not String.IsNullOrEmpty(pEPL.GetValue("ContactNumber")) Then
                      'Get Branch from members address
                      Dim vList As New ParameterList(True, True)
                      vList("ContactNumber") = pEPL.GetValue("ContactNumber")
                      vList("AddressNumber") = pEPL.GetValue("AddressNumber")
                      Dim vCRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, IntegerValue(pEPL.GetValue("ContactNumber")), vList))
                      If vCRow IsNot Nothing Then
                        Dim vContactInfo As ContactInfo = New ContactInfo(vCRow)
                        pEPL.SetValue("Branch", vContactInfo.Branch)
                      End If
                    End If
                  End If
                  .SetValue("Product", vRow.Item("FirstPeriodsProduct").ToString, True)
                  If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_retain_product_rate) = True AndAlso (mvTA.LastMembershipType IsNot Nothing AndAlso mvTA.LastMembershipType.Length > 0 AndAlso mvTA.LastMembershipType = pValue) Then
                    .SetValue("Rate", mvTA.LastMembershipRate)
                  Else
                    If Not SetRateFromMembershipPrices(pEPL) Then .SetValue("Rate", vRow.Item("FirstPeriodsRate").ToString)
                  End If
                End If
                .SetErrorField("Product", "")   'Clear any error
                .SetErrorField("Rate", "")      'Clear any error
                EPL_ValueChanged(pEPL, "Rate", .GetValue("Rate"))
                If vRow.Item("MembersPerOrder").ToString = "2" Then
                  Dim vContactInfo As New ContactInfo(mvTA.PayerContactNumber)
                  If vContactInfo.ContactType = ContactInfo.ContactTypes.ctJoint Then
                    'Payer is joint contact so set member to be first selected individual contact
                    If mvTA.TransactionType = "MEMC" Then
                      'We should never ever get here but a previous bug allowed it so if we do error as the data is invalid
                      pEPL.SetErrorField("MembershipType", InformationMessages.ImJointContactCannotBeMember, True)
                    Else
                      .SetErrorField("ContactNumber", "")    'Clear any error
                      Dim vDataTable As DataTable = FormHelper.GetIndividualsFromJointContact(vContactInfo.ContactNumber)
                      .SetValue("ContactNumber", GetContactNumberFromDataTable(vDataTable).ToString)
                    End If

                  End If
                End If
                If mvTA.TransactionType = "MEMC" Then
                  'Check for price changes since Renewals
                  CheckForCMTPriceChange(pEPL)
                  'Reset EligibelForGiftAid
                  Dim vList As New ParameterList(True, True)
                  vList.IntegerValue("TraderApplication") = mvTA.ApplicationNumber
                  vList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
                  vList("MembershipType") = pValue
                  vList("TraderTransactionType") = mvTA.TransactionType
                  vList("GiftMembership") = .GetValue("GiftMembership")
                  Dim vReturn As ParameterList = DataHelper.ProcessTraderPPEligibleForGiftAid(vList)
                  If vReturn("EligibleForGiftAid") = "D" Then
                    'Disable EligibleForGiftAid
                    .SetValue("EligibleForGiftAid", "N", True)
                  Else
                    'Enable EligibleForGiftAid and set value
                    .EnableControl("EligibleForGiftAid", True)
                    .SetValue("EligibleForGiftAid", vReturn("EligibleForGiftAid"))
                  End If
                  mvTA.CMTPrevMembershipTypeCode = pValue
                End If
              End With
            End If
          End If
        End If
        mvSavedMembershipType = pValue
      Case "NumberOfMembers", "MaxFreeAssociates"
        If ResetMembershipDetails(pEPL, pParameterName, pValue) Then
          If mvTA.PPDDataSet.Tables.Contains("DataRow") Then
            'NumberOfMembers/Associates has changed so remove PPDLines
            mvTA.PPDDataSet.Tables.Remove("DataRow")
          End If
        End If
      Case "OneYearGift", "PackToDonor"
        If pValue = "Y" Then pEPL.SetValue("GiftMembership", "Y")
      Case "PaymentFrequency"
        If Not SetRateFromMembershipPrices(pEPL) Then
          If pEPL.GetValue("MembershipType").Length > 0 Then
            If mvTA.TransactionType <> "MEMC" OrElse (mvTA.TransactionType = "MEMC" AndAlso (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.me_renew_at_same_rate) = True OrElse mvTA.PaymentPlan.DetermineMembershipPeriod() <> PaymentPlanInfo.MembershipPeriodTypes.mptSubsequentPeriod)) Then
              pEPL.SetValue("Rate", pEPL.FindTextLookupBox("MembershipType").GetDataRow.Item("FirstPeriodsRate").ToString)
            Else
              pEPL.SetValue("Rate", pEPL.FindTextLookupBox("MembershipType").GetDataRow.Item("SubsequentPeriodsRate").ToString)
            End If
          End If
        End If
      Case "Rate"
        ResetMembershipDetails(pEPL, pParameterName, pValue)
      Case "Source"
        ResetMembershipDetails(pEPL, pParameterName, pValue)
        mvTA.TransactionSource = pValue
    End Select
  End Sub

  Private Function SetRateFromMembershipPrices(ByVal pEPL As EditPanel) As Boolean
    Dim vList As ParameterList = GetMembershipPrices(pEPL)
    If vList.Contains("PrimaryRate") Then
      pEPL.SetValue("Rate", vList("PrimaryRate"), False, False, True, True)
      Return True
    End If
  End Function

  Private Sub ProcessPPNumbersChanged(ByVal pTraderPage As TraderPage, ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    If pValue.Length > 0 Then
      If (pTraderPage.PageType = CareServices.TraderPageType.tpMembership OrElse pTraderPage.PageType = CareServices.TraderPageType.tpChangeMembershipType) AndAlso pParameterName = "MemberNumber" Then
        'User entered MemberNumber (already validated in EPL_ValidateItem
        'Just need to check that PaymentPlan is not provisional
      Else
        If mvTA.PaymentPlan IsNot Nothing Then
          'set the balance here, as the ppbalance is otherwise only set if we change the balance on the PPM Page. we need PPBalace to set the balance on the pps page.
          mvTA.PPBalance = mvTA.PaymentPlan.Balance
          'If vtraderpage.PageType = CareServices.TraderPageType.tpTransactionDetails Then ShowMemberInfo(Nothing)
          If pValue <> pValue.Trim() Then
            pEPL.SetValue(pParameterName, pValue.Trim())
            pValue = pValue.Trim()
          End If
          If pValue.Length > 0 Then
            If pParameterName <> "AffiliatedMemberNumber" Then
              'If mvCurrPageType = tpContactSelection And mvTraderApplication.AppType = atMaintenance Then
              'Nothing to do
              ' Else
              ''Do not select provisional payment plans
              'vSQL = vSQL & " AND (provisional IS NULL OR provisional = 'N')"
              'End If
              'May need to make a 2nd selection if the Payment Plan is for an Organisation
              'vRecordSet = gvConn.GetRecordSet(vSQL)
              With mvTA.PaymentPlan
                If (pTraderPage.PageType <> CareServices.TraderPageType.tpMembership And pTraderPage.PageType <> CareServices.TraderPageType.tpChangeMembershipType) Then
                  Dim vPayerContactValid As Boolean = True
                  'mvSettingContact = True
                  If pTraderPage.PageType = CareServices.TraderPageType.tpPayments AndAlso pParameterName <> "PaymentPlanNumber" Then
                    'Set pUpdateLastValue to True so that the field is not validated as validating then clears the MemberNumber and it has already been validated as part of validating the MemberNumber
                    pEPL.SetValue("PaymentPlanNumber", .PaymentPlanNumber.ToString, False, False, True, True)
                  End If
                  If FindControl(pEPL, "ContactNumber", False) IsNot Nothing AndAlso FindControl(pEPL, "ContactNumber").Enabled Then
                    pEPL.SetValue("ContactNumber", .ContactNumber.ToString)
                    If pEPL.FindPanelControl(Of TextLookupBox)("ContactNumber", True).IsValid Then
                      EPL_ValueChanged(pEPL, "ContactNumber", pValue)
                    Else
                      pEPL.SetErrorField("ContactNumber", InformationMessages.ImInvalidValue)
                      vPayerContactValid = False
                    End If
                  End If
                  If vPayerContactValid Then
                    Dim vUpdateAmount As Boolean = True
                    Dim vBalance As Double
                    If .Balance = 0 Then
                      If pTraderPage.PageType <> CareServices.TraderPageType.tpContactSelection And pTraderPage.PageType <> CareServices.TraderPageType.tpCancelPaymentPlan Then
                        Dim vMsg As New StringBuilder
                        If pTraderPage.PageType = CareNetServices.TraderPageType.tpPayments Then
                          '20376: If adding a payment or invoice to an already fully-paid Payment Plan then raise a prompt to continue
                          vMsg.AppendLine(If(mvTA.PPPaymentType = "CRED", QuestionMessages.QmConfirmPPInAdvancePaymentCS, QuestionMessages.QmConfirmPPInAdvancePayment))

                          vMsg.AppendLine(String.Format(QuestionMessages.QmConfirmPPInAdvancePaymentBalance,
                                                        .NextPaymentDue.ToString(AppValues.DateFormat), .LastPaymentDate.ToString(AppValues.DateFormat), .LastPaymentAmount, .FrequencyAmount, vbCrLf))

                          If ShowQuestion(vMsg.ToString(), MessageBoxButtons.YesNo, pDefaultButton:=MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.No Then
                            pEPL.SetValue("MemberNumber", String.Empty, False, False, False, True)
                            pEPL.SetValue("CovenantNumber", String.Empty, False, False, False, True)
                            pEPL.SetValue("PaymentPlanNumber", String.Empty, False, False, False, True)
                            vUpdateAmount = False
                          End If
                        Else
                          vMsg.AppendLine(String.Format(InformationMessages.ImPPBalanceZero1, .NextPaymentDue.ToString(AppValues.DateFormat)))
                          vMsg.AppendLine(String.Format(InformationMessages.ImPPBalanceZero2, .LastPaymentDate.ToString(AppValues.DateFormat), .LastPaymentAmount, .FrequencyAmount))
                          ShowInformationMessage(vMsg.ToString())
                        End If
                      End If
                      vBalance = .FrequencyAmount
                    Else
                      vBalance = .Balance
                    End If
                    'code for doing this only on amount page
                    If vUpdateAmount AndAlso (mvTA.ApplicationType <> ApplicationTypes.atCreditListReconciliation AndAlso pTraderPage.PageType <> CareServices.TraderPageType.tpCancelPaymentPlan _
                    OrElse (mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation AndAlso pTraderPage.PageType = CareServices.TraderPageType.tpPayments AndAlso pEPL.GetDoubleValue("Amount") > vBalance)) Then
                      Dim vAmount As Double
                      If (InStr(AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_ppd_proportional_balance), "FULLPAYMENT") > 0 OrElse InStr(AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_ppd_proportional_balance), "NEW") > 0) AndAlso .StartDate = .RenewalDate And .FirstAmount > 0 Then
                        vAmount = .FirstAmount
                      Else
                        'vAmount = DoubleValue(mvtraderapplication.pprow("NextPaymentAmount").ToString)
                        If .PaymentFreq.Frequency <> 1 Then
                          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanOutstandingOPS, mvTA.PaymentPlan.PaymentPlanNumber))
                          If vTable IsNot Nothing Then
                            For Each vRow As DataRow In vTable.Rows
                              Select Case vRow.Item("ScheduledPaymentStatus").ToString
                                Case "D", "P", "V"
                                  vAmount = DoubleValue(vTable.Rows(0)("AmountOutstanding").ToString)
                              End Select
                              If vAmount > 0 Then Exit For
                            Next
                          End If
                          If vAmount = 0 Then vAmount = .FrequencyAmount
                        Else
                          vAmount = .FrequencyAmount
                        End If
                      End If
                      If vBalance > 0 And vBalance < vAmount Then
                        vAmount = vBalance
                      Else
                        If (mvTA.PaymentPlan.ProportionalBalanceSetting And (PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
                         And mvTA.PaymentPlan.StartDate = mvTA.PaymentPlan.RenewalDate And mvTA.PaymentPlan.FirstAmount > 0 Then
                          vAmount = mvTA.PaymentPlan.FirstAmount
                        Else
                          vAmount = .FrequencyAmount
                        End If
                      End If
                      vAmount = mvTA.CalcCurrencyAmount(vAmount, False)
                      If pTraderPage.PageType <> CareServices.TraderPageType.tpContactSelection Then
                        If vBalance > 0 And vBalance < vAmount Then
                          SetValueRaiseChanged(pEPL, "Amount", vBalance.ToString("0.00"))
                        Else
                          If vAmount = 0 Then
                            SetValueRaiseChanged(pEPL, "Amount", .NextPaymentAmount.ToString("0.00"))
                          Else
                            SetValueRaiseChanged(pEPL, "Amount", vAmount.ToString("0.00"))
                          End If
                        End If
                      End If
                    End If

                    If FindControl(pEPL, "SalesContactNumber", False) IsNot Nothing Then
                      If .SalesContactNumber > 0 Then
                        pEPL.SetValue("SalesContactNumber", mvTA.PaymentPlan.SalesContactNumber.ToString)
                      Else
                        If mvTA.SalesContactNumber > 0 Then pEPL.SetValue("SalesContactNumber", mvTA.SalesContactNumber.ToString)
                      End If
                    End If
                    If pTraderPage.PageType = CareServices.TraderPageType.tpCancelPaymentPlan Then
                      pEPL.SetValue("RenewalDate", .RenewalDate.ToString, True)
                      pEPL.SetValue("Balance", DoubleValue(.Balance.ToString).ToString("0.00"), True)
                    Else
                      If vUpdateAmount AndAlso FindControl(pEPL, "Balance", False) IsNot Nothing Then
                        'i dont know wha other page types dont have balnce. please add here
                        pEPL.SetValue("Balance", DoubleValue(.Balance.ToString).ToString("0.00"))
                      End If
                    End If
                  End If
                End If
              End With
            End If
          End If
        End If
      End If
      ClearPPNumberChanged(pEPL, pParameterName)
    End If
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection AndAlso Not pParameterName.Equals("ContactNumber") AndAlso Not pParameterName.Equals("AddressNumber") Then
      pEPL.EnableControlList("ContactNumber,AddressNumber", PPFieldsBlank(pEPL))
    End If
  End Sub

  Private Function ResetMembershipDetails(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String) As Boolean
    Dim vOrigValue As String
    Dim vReset As Boolean = True

    Select Case pParameterName
      Case "Amount"
        vOrigValue = mvTA.LastMembershipFixedAmount
      Case "Joined"
        vOrigValue = mvTA.LastMembershipJoinedDate
      Case "MaxFreeAssociates"
        vOrigValue = mvTA.LastMembershipNumberAssociates.ToString
      Case "NumberOfMembers"
        vOrigValue = mvTA.LastMembershipNumberMembers.ToString
      Case "Rate"
        vOrigValue = mvTA.LastMembershipRate
      Case "Source"
        vOrigValue = mvTA.LastMembershipSource
      Case Else   'MembershipType
        vOrigValue = mvTA.LastMembershipType
    End Select

    If (pValue <> vOrigValue) AndAlso mvTraderPages(CareServices.TraderPageType.tpPaymentPlanDetails.ToString).DefaultsSet = True AndAlso GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "Balance").Length > 0 Then
      If ShowQuestion(QuestionMessages.QmMembershipDetailsChanged, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        ClearPageDefaults(CareServices.TraderPageType.tpPaymentPlanDetails)
        mvTA.PPDDataSet.Tables.Clear()      'Clear out existing values
        mvTA.IncentiveDataSet = Nothing
        mvTA.FulfilIncentives = False
        mvTA.CMTPrevMembershipTypeCode = ""
        'Also need to reset RenewalAmount back to zero if held separatly to the page value
      Else
        'Need to reset field back to original value
        pEPL.SetValue(pParameterName, vOrigValue, False, False, True, True)
        vReset = False
      End If
    End If
    Return vReset
  End Function
  Private Function SumLineTypes(ByVal pLineType As String) As Double
    Dim vAmount As Double
    For Each vDataRow As DataRow In mvTA.AnalysisDataSet.Tables("DataRow").Rows
      If vDataRow("TradertransactionType").ToString = pLineType Then
        Select Case pLineType
          Case "SALE"
            If vDataRow.Item("StockSale").ToString = "Y" Then
              vAmount = vAmount + DoubleValue(vDataRow("Amount").ToString)
            End If
          Case Else
            vAmount = vAmount + DoubleValue(vDataRow("Amount").ToString)
        End Select
      End If
    Next
    Return vAmount
  End Function
  Private Sub SetAmount(ByVal pEPL As EditPanel, Optional ByVal pParameterName As String = "")
    Dim vQuantity As Double
    Dim vPrice As Double
    Dim vGrossAmount As Double
    Dim vAmount As Double
    Dim vDiscount As Double
    Dim vSpecialInitialPeriod As Boolean
    Dim vOrigQuantity As Double
    Dim vGrossEnabled As Boolean
    Dim vEntitlement As Integer
    Dim vDays As Integer
    Dim vCurrentQuantity As Integer

    If pParameterName.Length = 0 Then
      If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
        pParameterName = "Balance"
      ElseIf (mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.PayerHasDiscount) OrElse (mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking AndAlso mvTA.PayerHasDiscount AndAlso pEPL.FindPanelControl("Source").Visible) Then
        pParameterName = "GrossAmount"
      Else
        pParameterName = "Amount"
      End If
    End If

    Dim vHasPRDVATAmount As Boolean
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpPostageAndPacking Then
      If FindControl(pEPL, "Amount2", False) IsNot Nothing Then
        Dim vRateRow As DataRow = pEPL.FindPanelControl(Of TextLookupBox)("Rate").GetDataRow
        Dim vRatePrice As Double
        If vRateRow IsNot Nothing Then
          vRatePrice = FixTwoPlaces(DoubleValue(vRateRow.Item("CurrentPrice").ToString))
        End If
        If vRatePrice <> 0.0 Then
          pEPL.EnableControl("Amount2", False)
          If mvTA.LinePriceVATEx Then
            mvTA.LineVATAmount = FixTwoPlaces((vRatePrice * (mvTA.LineVATPercentage / 100)))
            vRatePrice = vRatePrice + mvTA.LineVATAmount
          End If
          pEPL.SetValue("Amount2", FixTwoPlaces(vRatePrice).ToString("0.00"))
          pEPL.SetValue("Percentage", "")
          pEPL.EnableControl("Percentage", False)
        Else
          pEPL.EnableControl("Percentage", True)
          If FindControl(pEPL, "Percentage", False) IsNot Nothing Then
            If String.IsNullOrEmpty(pEPL.GetValue("Percentage")) Then
              pEPL.SetValue("Percentage", mvCarraigePercentage.ToString("0.00"))
            End If
            mvTA.LinePrice = 0
            For Each vAmtRow As DataRow In mvTA.AnalysisDataSet.Tables("DataRow").Rows
              If vAmtRow.Item("TraderTransactionType").ToString = "SALE" Then
                mvTA.LinePrice = mvTA.LinePrice + DoubleValue(vAmtRow.Item("Amount").ToString)
              End If
            Next
            pEPL.EnableControl("Amount2", True)
            pEPL.SetValue("Amount2", ((SumLineTypes("SALE") + SumLineTypes("P")) * (DoubleValue(pEPL.GetValue("Percentage").ToString)) / 100).ToString("0.00"), pUpdateLastValue:=True)
          End If
        End If
      End If
    Else
      'If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts And mvPPLine > 0 Then
      '  If mvTraderApplication.PaymentPlanDetails.Exists(mvPPLine) Then
      '    vPPD = mvTraderApplication.PaymentPlanDetails.Item(Format$(mvPPLine))
      '    vSpecialInitialPeriod = vPPD.SpecialInitialPeriodIncentive
      '  End If
      'End If
      If Not vSpecialInitialPeriod And Not mvTA.FixedPrice Then
        If FindControl(pEPL, "Quantity", False) IsNot Nothing Then
          vQuantity = pEPL.GetDoubleValue("Quantity")
        Else
          'Quantity control does not exist on page
          vQuantity = 1
        End If
      Else
        vQuantity = 1
      End If
      'If GetControlIndex("from_date", 0, vFromIndex) Then
      '  If GetControlIndex("to_date", 0, vToIndex) Then
      '    If IsDate(meb(vFromIndex)) And IsDate(meb(vToIndex)) Then
      '      vDays = DateDiff("y", meb(vFromIndex), meb(vToIndex))
      '      If vDays > 0 Then vQuantity = vQuantity * vDays
      '    End If
      '  End If
      'End If

      If pEPL.FindPanelControl("StartDate", False) IsNot Nothing Then
        If pEPL.FindPanelControl("EndDate", False) IsNot Nothing Then
          Dim vFromDate As Date = pEPL.FindDateTimePicker("StartDate").Value.Date
          Dim vToDate As Date = pEPL.FindDateTimePicker("EndDate").Value.Date
          Dim vList As ParameterList
          vDays = CInt(DateDiff("y", vFromDate, vToDate))
          mvTA.SBNewQuantity = 0
          If vDays > 0 Then
            Dim vProductOffer As DataRow = Nothing

            If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking AndAlso (Not mvTA.PayerHasDiscount) _
             AndAlso pEPL.FindPanelControl("Product", False) IsNot Nothing AndAlso pEPL.FindPanelControl("Rate", False) IsNot Nothing Then
              vList = New ParameterList(True)
              vList("Product") = pEPL.GetValue("Product")
              vList("Rate") = pEPL.GetValue("Rate")
              If vList("Product").Length > 0 Then vProductOffer = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProductOffers, vList)
            End If

            Dim vActivityQty As Integer
            'Calculate amount of free entitlement
            If vProductOffer IsNot Nothing AndAlso IntegerValue(vProductOffer("ProductQuantity")) > 0 _
             AndAlso pEPL.FindPanelControl("BookingContactNumber") IsNot Nothing Then
              vList = New ParameterList(True)
              vList("ContactNumber") = mvTA.PayerContactNumber.ToString
              vList("Activity") = vProductOffer("Activity").ToString
              vList("ActivityValue") = vProductOffer("ActivityValue").ToString
              Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtContactCategories, vList)
              If vRow IsNot Nothing Then vCurrentQuantity = IntegerValue(vRow("Quantity"))
              vEntitlement = (vCurrentQuantity + vDays) \ IntegerValue(vProductOffer("ActivityQuantity"))
              vActivityQty = IntegerValue(vProductOffer("ActivityQuantity"))
            End If
            If pEPL.FindPanelControl("ActivityQuantity", False) IsNot Nothing Then
              If pEPL.FindPanelControl("ProductQuantity", False) IsNot Nothing Then
                If IntegerValue(pEPL.GetValue("ActivityQuantity")) <> vEntitlement Then
                  pEPL.SetValue("ActivityQuantity", vEntitlement.ToString)
                Else
                  vEntitlement = IntegerValue(pEPL.GetValue("ProductQuantity"))
                End If
                If vEntitlement > vDays Then vEntitlement = CInt(vDays)
                pEPL.SetValue("ProductQuantity", vEntitlement.ToString)
                vDays = vDays - vEntitlement
                mvTA.SBNewQuantity = vCurrentQuantity - (vEntitlement * vActivityQty)
                mvTA.SBNewQuantity = mvTA.SBNewQuantity + vDays
              End If
            End If
            vQuantity = vQuantity * vDays
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking AndAlso mvTA.FixedUnitRate Then vQuantity = 1
          End If
        End If
      End If

      If FindControl(pEPL, "LinePrice", False) IsNot Nothing Then mvTA.LinePrice = DoubleValue(pEPL.GetValue("LinePrice"))

      vPrice = mvTA.LinePrice

      If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts Or mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
        '  With mvTraderApplication.PaymentPlan
        '    'Pro-rate product price if necessary
        Dim vList As New ParameterList(True)
        vList("Amount") = vPrice.ToString
        vList("Product") = mvCurrentPage.EditPanel.GetValue("Product")
        vList("Rate") = mvCurrentPage.EditPanel.GetValue("Rate")
        Dim vContactNumber As Integer
        If mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "CMEM" Then
          vContactNumber = IntegerValue(GetPageValue(CareNetServices.TraderPageType.tpMembership, "ContactNumber"))
        End If
        If vContactNumber = 0 Then vContactNumber = mvTA.PayerContactNumber
        vList("ContactNumber") = vContactNumber.ToString
        If mvTA.ApplicationType = ApplicationTypes.atMaintenance Then
          '      If .PlanType = PaymentPlanInfo.ppType.pptMember And .FixedRenewalCycle And .PreviousRenewalCycle _
          '      And (.ProportionalBalanceSetting And (PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
          '      And .StartDate = .RenewalDate Then
          '        .LoadMembers()
          '        vPrice = .GetProrataBalance(vPrice, .Member.Joined)
          '      ElseIf (.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsExisting) > 0 Then
          vList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
          If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance AndAlso FindControl(mvCurrentPage.EditPanel, "EffectiveDate", False) IsNot Nothing Then
            vList("EffectiveDate") = mvCurrentPage.EditPanel.GetValue("EffectiveDate")
          End If
          Dim vPFCode As String = GetPageValue(CareNetServices.TraderPageType.tpPaymentPlanMaintenance, "PaymentFrequency")
          If Not String.IsNullOrWhiteSpace(vPFCode) Then vList("PaymentFrequency") = vPFCode
          Dim vReturnList As ParameterList = DataHelper.GetDetailBalance(vList)
          vPrice = DoubleValue(vReturnList("Balance").ToString)     'This includes VAT
          If mvTA.LinePriceVATEx = True AndAlso vReturnList.ContainsKey("NetBalance") AndAlso DoubleValue(vReturnList("NetBalance").ToString) <> 0 Then vPrice = DoubleValue(vReturnList("NetBalance").ToString)
          mvTA.PaymentPlanDetailsPricing.InitFromParameterList(vReturnList)
        ElseIf mvTA.ApplicationType = ApplicationTypes.atTransaction Then
          Dim vPP As PaymentPlanInfo = mvTA.PaymentPlan
          If vPP Is Nothing Then vPP = New PaymentPlanInfo
          'Adding a new PPD when creating a new membership
          If (mvTA.TransactionType = "MEMB" Or mvTA.TransactionType = "MEMC") AndAlso vPP.FixedRenewalCycle AndAlso vPP.PreviousRenewalCycle _
             AndAlso (vPP.ProportionalBalanceSetting And (PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsNew)) > 0 Then
            If vPP.PaymentPlanNumber > 0 Then
              vList.IntegerValue("PaymentPlanNumber") = vPP.PaymentPlanNumber
            Else
              vList("Joined") = GetPageValue(CareServices.TraderPageType.tpMembership, "Joined")
              vList("MembershipType") = mvTA.LastMembershipType
              vList("Balance") = mvTA.PPBalance.ToString
              vList("PaymentFrequency") = GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "PaymentFrequency")
              vList("Term") = GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "OrderTerm")
              vList("PaymentMethod") = mvTA.PPPaymentMethod
            End If
            Dim vReturnList As ParameterList = DataHelper.GetDetailBalance(vList)
            vPrice = DoubleValue(vReturnList("Balance").ToString)       'This includes VAT
            If mvTA.LinePriceVATEx = True AndAlso vReturnList.ContainsKey("NetBalance") Then vPrice = DoubleValue(vReturnList("NetBalance").ToString)
            mvTA.PaymentPlanDetailsPricing.InitFromParameterList(vReturnList)
          Else
            SetPPDPricingData(pEPL)
          End If
        End If
        '  End With
      End If

      vOrigQuantity = vQuantity
      If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails Then
        vHasPRDVATAmount = FindControl(pEPL, "VatAmount", False) IsNot Nothing
        If mvTA.PayerHasDiscount = True AndAlso mvTA.LinePrice = 0 AndAlso pParameterName = "Amount" Then
          'The price is zero (e.g. Donation) and the user has changed the price
          'We will use the user-entered price so that we can set the GrossAmount (Discount is to remain at zero)
          'vQuantity is re-set to 1 otherwise these calculations could change the user-entered Amount
          vQuantity = 1
          vPrice = DoubleValue(pEPL.GetValue("Amount"))
        End If
      End If

      If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking Then
        vDays = vDays + vEntitlement
        vQuantity = IntegerValue(pEPL.GetValue("Quantity")) * vDays
        If mvTA.FixedUnitRate Then vQuantity = 1
        mvTA.SBGrossQty = vQuantity
        If mvTA.LinePriceVATEx Then
          mvTA.SBGrossAmount = FixTwoPlaces(Int(FixTwoPlaces(((vPrice * vQuantity) + ((vPrice * vQuantity) * (mvTA.LineVATPercentage / 100))) * 100)) / 100)
        Else
          mvTA.SBGrossAmount = FixTwoPlaces(vPrice * vQuantity)
        End If
        mvTA.SBEntitlementQty = IntegerValue(pEPL.GetValue("Quantity")) * vEntitlement
        If mvTA.ServiceBookingCredits Then
          mvTA.SBGrossQty = -mvTA.SBGrossQty
          mvTA.SBGrossAmount = -mvTA.SBGrossAmount
          mvTA.SBEntitlementQty = -mvTA.SBEntitlementQty
        End If
        If mvTA.PayerHasDiscount AndAlso (pEPL.FindPanelControl("Source", False) IsNot Nothing AndAlso pEPL.FindPanelControl("Source").Visible) Then
          If pEPL.FindPanelControl("GrossAmount", False) IsNot Nothing Then
            pEPL.SetValue("GrossAmount", mvTA.SBGrossAmount.ToString("0.00"))  'Gross amount needs to hold the full amount incl. vat
            'vDiscount = FixTwoPlaces(CDbl(IIf(mvTA.LinePrice = 0, mvTA.LinePrice, (FixTwoPlaces(mvTA.SBGrossAmount) * FixTwoPlaces(mvTA.DiscountPercentage) / 100))))
            'Discount should be calculated on the line amount and not on the gross amount 
            vDiscount = (FixTwoPlaces(mvTA.LinePrice * vQuantity) * FixTwoPlaces(mvTA.DiscountPercentage)) / 100
          End If
          If pEPL.FindPanelControl("Discount", False) IsNot Nothing Then
            pEPL.SetValue("Discount", vDiscount.ToString("0.00"))
          End If
        End If
      End If

      If mvTA.LinePriceVATEx Then
        'Added this condition as the vat amount should be calculated on the discounted amount and not on the original amount
        If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking Then
          mvTA.LineVATAmount = FixTwoPlaces((((vPrice - vDiscount) * vQuantity) * (mvTA.LineVATPercentage / 100)))
        Else
          mvTA.LineVATAmount = FixTwoPlaces(((vPrice * vQuantity) * (mvTA.LineVATPercentage / 100)))
          vGrossAmount = FixTwoPlaces((vPrice * vQuantity) + mvTA.LineVATAmount)
          If vHasPRDVATAmount = True AndAlso mvTA.ShowVATExclusiveAmount Then
            pEPL.SetValue(pParameterName, FixTwoPlaces(vPrice * vQuantity).ToString("0.00"), pUpdateLastValue:=True) 'Display price without VAT
          Else
            pEPL.SetValue(pParameterName, vGrossAmount.ToString("0.00"), pUpdateLastValue:=True)                                'Display price with VAT
          End If
        End If
      Else
        vGrossAmount = FixTwoPlaces(vPrice * vQuantity)
        pEPL.SetValue(pParameterName, vGrossAmount.ToString("0.00"), pUpdateLastValue:=True)
        mvTA.LineVATAmount = AppHelper.CalculateVATAmount((vPrice * vQuantity), mvTA.LineVATPercentage)
      End If
      If vHasPRDVATAmount Then pEPL.SetValue("VatAmount", mvTA.LineVATAmount.ToString("0.00"))

      If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpServiceBooking AndAlso pEPL.FindPanelControl("Amount", False) IsNot Nothing Then
        If mvTA.LinePriceVATEx Then
          pEPL.SetValue("Amount", (FixTwoPlaces(vPrice - vDiscount + mvTA.LineVATAmount) * vQuantity).ToString("0.00"))
        Else
          pEPL.SetValue("Amount", (FixTwoPlaces(vPrice - vDiscount) * vQuantity).ToString("0.00"))
        End If
      End If

      If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails And mvTA.PayerHasDiscount Then
        pEPL.SetValue("GrossAmount", vGrossAmount.ToString("0.00"))  'Gross amount needs to hold the full amount incl. vat        
        If mvTA.LinePrice <> 0 Then
          If mvTA.LinePriceVATEx Then
            vDiscount = FixTwoPlaces((vPrice * vQuantity) * mvTA.DiscountPercentage / 100)
          Else
            vDiscount = FixTwoPlaces(vGrossAmount * mvTA.DiscountPercentage / 100)
          End If
        End If
        pEPL.SetValue("Discount", vDiscount.ToString("0.00"))
        If mvTA.LinePriceVATEx Then
          vAmount = FixTwoPlaces(((vPrice * vQuantity) - vDiscount) + (((vPrice * vQuantity) - vDiscount) * (mvTA.LineVATPercentage / 100)))
        Else
          vAmount = FixTwoPlaces(vGrossAmount - vDiscount)
        End If
        pEPL.SetValue("Amount", vAmount.ToString("0.00"))
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.LinePrice = 0 Then pEPL.PanelInfo.PanelItems("Amount").ValueChanged(vAmount.ToString)
        'Amount has changed so need to re-calculate the VAT
        mvTA.LineVATAmount = AppHelper.CalculateVATAmount(vAmount, mvTA.LineVATPercentage)   'Amount currently includes VAT
        If mvTA.ShowVATExclusiveAmount Then pEPL.SetValue("Amount", FixTwoPlaces(vAmount - mvTA.LineVATAmount).ToString("0.00")) 'Update Amount to exclude VAT
        If vHasPRDVATAmount Then pEPL.SetValue("VatAmount", mvTA.LineVATAmount.ToString("0.00"))
        If mvTA.LinePrice = 0 Then vQuantity = vOrigQuantity
      End If

      Select Case mvCurrentPage.PageType
        Case CareServices.TraderPageType.tpChangeMembershipType, CareServices.TraderPageType.tpMembership,
           CareServices.TraderPageType.tpPaymentPlanProducts, CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
          Dim vAmountField As String = pParameterName
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpChangeMembershipType OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpMembership Then
            pEPL.EnableControl(pParameterName, vGrossAmount = 0)
          Else
            If mvTA.LinePrice <> 0 Then
              pEPL.SetValue("Amount", "", True, False, True, True)
            Else
              pEPL.EnableControl("Amount", True)
            End If
          End If
          'For VAT-Exclusive Rates need to enable the NetFixedAmount if we have it, otherwise use the default Amount (GrossFixedAmount)
          Dim vHasNetFixedAmount As Boolean
          If FindControl(pEPL, "NetFixedAmount", False) IsNot Nothing AndAlso pEPL.PanelInfo.PanelItems("NetFixedAmount").Visible = True Then vHasNetFixedAmount = True
          If vHasNetFixedAmount Then
            If mvTA.LinePriceVATEx = True AndAlso mvTA.LinePrice = 0 Then
              'NetFixedAmount is visible so enable it and disable Amount
              pEPL.EnableControl("NetFixedAmount", True)
              SetValueRaiseChanged(pEPL, "NetFixedAmount", pEPL.GetValue("Amount"))  'Set NetFixedAmount to be the Amount
              pEPL.EnableOrSetValueDisable("Amount", False, "")
            Else
              'Always disable NetFixedAmount 
              pEPL.EnableOrSetValueDisable("NetFixedAmount", False, "")
            End If
          End If
          If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanProducts OrElse mvCurrentPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then SetPPDPricingData(pEPL)

        Case CareServices.TraderPageType.tpProductDetails,
           CareServices.TraderPageType.tpEventBooking,
           CareServices.TraderPageType.tpAccommodationBooking,
           CareServices.TraderPageType.tpServiceBooking
          If (mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.PayerHasDiscount) OrElse (mvCurrentPage.PageType = CareServices.TraderPageType.tpServiceBooking AndAlso mvTA.PayerHasDiscount AndAlso pEPL.FindPanelControl("Source").Visible) Then
            vGrossEnabled = (mvTA.LinePrice = 0)
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.LinePrice = 0 Then vGrossEnabled = (DoubleValue(pEPL.GetValue("Amount")) = 0)
            pEPL.EnableControl("GrossAmount", vGrossEnabled)
            pEPL.EnableControl("DiscountAmount", (mvTA.LinePrice = 0) Or AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_discount_amend))
            pEPL.EnableControl("Amount", mvTA.LinePrice = 0)
          End If
          pEPL.EnableControl(pParameterName, mvTA.LinePrice = 0)
          pEPL.EnableControl("VatAmount", False)
      End Select

      'If mvCurrentPage.PageType = tpPaymentPlanProducts Or mvCurrentPage.PageType = tpPaymentPlanDetailsMaintenance Then
      '  If GetControlIndex("amount", 0, vAttrIndex) Then
      '    'Only allow Amount to be set where the price = 0
      '    If vPrice = 0 Then
      '      txtN(vAttrIndex).Enabled = True
      '    Else
      '      txtN(vAttrIndex).Text = gvNull
      '      txtN(vAttrIndex).Enabled = False
      '    End If
      '  End If
      'End If
      'If mvCurrentPage.PageType = tpPaymentPlanDetailsMaintenance And vPrice > 0 Then
      '  'PPMaint and price is greater than zero so set balance to include any arrears
      '  If GetControlIndex("arrears", 0, vOtherAttrIndex) And GetControlIndex(pAttr, 0, vAttrIndex) Then
      '    vArrears = Val(txtN(vOtherAttrIndex).Text)
      '    'If we have some arrears then add this to the balance
      '    If vArrears > 0 Then txtN(vAttrIndex).Text = Format$(Val(txtN(vAttrIndex).Text) + vArrears)
      '  End If
      'End If
    End If
  End Sub

  Private Sub SetBankDetails(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    SetBankDetails(pEPL, pParameterName, pValue, "C")
  End Sub

  Private Sub SetBankDetails(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, ByVal pAlbacsBankDetails As String)
    Dim vPayerContactNumber As Integer
    Dim vPayerAddressNumber As Integer
    AppHelper.SetBankDetails(pEPL, pParameterName, pValue, mvTA.EditExistingTransaction, mvTA.BankDetailsNumber, mvTA.CreateContactAccount, mvTA.NewBank, vPayerContactNumber, vPayerAddressNumber, pAlbacsBankDetails)

    If mvCurrentPage.PageType <> CareServices.TraderPageType.tpStandingOrder And mvCurrentPage.PageType <> CareServices.TraderPageType.tpDirectDebit Then
      If pValue.Length > 0 Then
        pEPL.PanelInfo.PanelItems("Reference").Mandatory = True
      Else
        pEPL.SetValue("Reference", "")
        pEPL.PanelInfo.PanelItems("Reference").Mandatory = False
      End If
    End If

    If vPayerContactNumber > 0 Then
      Dim vUpdatePayerContact As Boolean = False
      If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
        vUpdatePayerContact = True
      ElseIf mvTA.ApplicationType = ApplicationTypes.atTransaction Then
        Dim vPage As TraderPage = mvTA.TraderPage(pEPL)
        vUpdatePayerContact = Not (vPage.PageType = CareNetServices.TraderPageType.tpDirectDebit OrElse vPage.PageType = CareNetServices.TraderPageType.tpStandingOrder)
      End If
      If vUpdatePayerContact Then
        mvTA.SetPayerContact(vPayerContactNumber, vPayerAddressNumber)
      End If
    End If


  End Sub

  Private Sub SetCreditCustomerAccount(ByVal pEPL As EditPanel, ByVal pByContact As Boolean, ByVal pValue As String, ByRef pValid As Boolean)
    Dim vList As New ParameterList(True)
    If pByContact Then
      vList("ContactNumber") = pValue
      vList("Company") = mvTA.CSCompany
    Else
      vList("SalesLedgerAccount") = pValue
    End If
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vList))
    Dim vFound As Boolean
    If vTable IsNot Nothing Then
      Dim vRow As DataRow = vTable.Rows(0)
      If Not pByContact Then
        For Each vRow In vTable.Rows
          If vRow("Company").ToString = mvTA.CSCompany AndAlso CInt(vRow("ContactNumber")) = mvTA.PayerContactNumber Then
            vFound = True
            Exit For
          End If
        Next
        If Not vFound Then
          vRow = vTable.Rows(0)
          If vRow("Company").ToString <> mvTA.CSCompany Then
            pEPL.SetErrorField("SalesLedgerAccount", GetInformationMessage(InformationMessages.ImSLAWrongCompany, mvTA.CSCompany), True)
            pValid = False
            Return
          End If
          If CInt(vRow("ContactNumber")) <> mvTA.PayerContactNumber Then
            pEPL.SetErrorField("SalesLedgerAccount", InformationMessages.ImSLAWrongContact, True)
            pValid = False
            Return
          End If
        End If
      End If
      With pEPL
        If (CInt(vRow("ContactNumber")) <> mvTA.PayerContactNumber) AndAlso mvTA.PayMethodsAtEnd = False Then mvTA.Pages(CareServices.TraderPageType.tpTransactionDetails.ToString).DefaultsSet = False 'Contact has changed so need to update tpTransactionDetails page
        mvTA.SetPayerContact(CInt(vRow("ContactNumber")), CInt(vRow("AddressNumber")))
        .SetValue("AddressNumber", vRow("AddressNumber").ToString)
        .SetValue("SalesLedgerAccount", vRow("SalesLedgerAccount").ToString)
        Dim vTermsNumber As String = vRow("TermsNumber").ToString
        If vTermsNumber.Length = 0 Then vTermsNumber = mvTA.CSTermsNumber
        .SetValue("TermsNumber", vTermsNumber)
        Dim vTermsPeriod As String = vRow("TermsPeriod").ToString
        If vTermsPeriod.Length = 0 Then vTermsPeriod = mvTA.CSTermsPeriod
        If vTermsPeriod = "M" Then vTermsPeriod = "Y"
        .SetValue("TermsPeriod", vTermsPeriod)
        Dim vTermsFrom As String = vRow("TermsFrom").ToString
        If vTermsFrom.Length = 0 Then vTermsFrom = mvTA.CSTermsFrom
        .SetValue("TermsFrom_" & vTermsFrom, vTermsFrom)
        .SetValue("CreditCategory", vRow("CreditCategory").ToString)
        .SetValue("StopCode", vRow("StopCode").ToString)
        .SetValue("CreditLimit", vRow("CreditLimit").ToString)
        .SetValue("CustomerType", vRow("CustomerType").ToString)
        .SetValue("Outstanding", vRow("Outstanding").ToString)
        .SetValue("OnOrder", vRow("OnOrder").ToString)
        vFound = True
        If vRow("StopCode").ToString.Length > 0 AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciRemoveCreditStopCode) = False Then
          .SetErrorField("SalesLedgerAccount", InformationMessages.ImSLAAccountOnStop, True)
        End If
        .EnableControl("StopCode", AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciRemoveCreditStopCode))
        mvTA.CreditTermsChanged = False
        If mvTA.EditExistingTransaction Then
          'Need to reset DefaultsSet on TRD page
          Dim vPage As TraderPage = mvTraderPages(CareServices.TraderPageType.tpTransactionDetails.ToString)
          If vPage IsNot Nothing Then vPage.DefaultsSet = False
        End If
      End With
    End If
    If Not vFound Then
      Dim vCreateNewCreditCustomer As Boolean
      If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciNewCreditCustomer) Then
        Dim vCreateCustomer As Boolean
        If pByContact Then
          If mvTA.mvAutoCreateCreditCust Then
            vCreateCustomer = True
          ElseIf ShowQuestion(QuestionMessages.QmNewCreditCustomer, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            vCreateCustomer = True
          Else
            vCreateCustomer = False
          End If
          If vCreateCustomer Then
            If (CInt(pEPL.GetValue("ContactNumber")) <> mvTA.PayerContactNumber) AndAlso mvTA.PayMethodsAtEnd = False Then mvTA.Pages(CareServices.TraderPageType.tpTransactionDetails.ToString).DefaultsSet = False 'Contact has changed so need to update tpTransactionDetails page
            mvTA.SetPayerContact(CInt(pEPL.GetValue("ContactNumber")), CInt(pEPL.GetValue("AddressNumber")))
            vCreateNewCreditCustomer = True
          End If
        Else
          If ShowQuestion(QuestionMessages.QmNewCreditAccount, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then vCreateNewCreditCustomer = True
        End If
      End If
      If vCreateNewCreditCustomer Then
        mvTA.NewCreditCustomer = True
        With pEPL
          .ClearControlList("CreditCategory,StopCode,CreditLimit,CustomerType,Outstanding,OnOrder")
          .EnableControlList("CreditCategory,StopCode,CreditLimit,CustomerType", True)
          .SetValue("SalesLedgerAccount", pValue)
          .SetValue("TermsNumber", mvTA.CSTermsNumber)
          .SetValue("TermsPeriod", CBoolYN(mvTA.CSTermsPeriod = "M"))
          .SetValue("TermsFrom_" & mvTA.CSTermsFrom, mvTA.CSTermsFrom)
          .SetValue("CreditCategory", mvTA.mvCreditCategory)
          Dim vContactInfo As ContactInfo = .FindTextLookupBox("ContactNumber").ContactInfo
          If vContactInfo IsNot Nothing Then .SetValue("CustomerType", vContactInfo.ContactTypeCode)
        End With
        mvTA.CreditTermsChanged = False
      Else
        mvTA.NewCreditCustomer = False
        If pByContact Then
          pEPL.SetErrorField("ContactNumber", InformationMessages.ImContactNotCreditCustomer, True)
        Else
          pEPL.SetErrorField("SalesLedgerAccount", InformationMessages.ImInvalidSalesLedgerAccount, True)
        End If
        pValid = False
      End If
    End If
  End Sub

  Private Sub SetCreditOrDebitCard(ByVal pEPL As EditPanel, ByVal pValue As String)
    'J397: ValidDate is not mandatory for both Credit and Debit cards
    Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_card_sales_combined_claim)
      Case "A", "Y"
        pEPL.EnableControlList("Issuer,IssueNumber,ValidDate", True)
      Case Else   '"N" or not set
        pEPL.EnableControlList("Issuer,IssueNumber,ValidDate", pValue = "D")
    End Select
    If pValue.Length > 0 Then
      'J807: Always disable AuthorisationCode when using OnlineAuthorisation and the authorisation type is SecureCXL
      Dim vEnable As Boolean = Not (mvTA.OnlineCCAuthorisation AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type) = "SCXLVPCSCP")
      If pValue = "C" Or pValue = "D" Then
        'CreditCard or DebitCard
        'Check card FloorLimit to enable/diable AuthorisationCode
        pEPL.EnableControl("AuthorisationCode", vEnable AndAlso CardAuthorisationRequired(mvTA.TransactionAmount))
      Else
        pEPL.EnableControl("AuthorisationCode", vEnable)
      End If
    Else
      'No value for CreditOrDebitCard so disable the AuthorisationCode field
      pEPL.EnableControl("AuthorisationCode", False)
    End If

  End Sub

  Private Sub SetMailingFromAmount(ByVal pEPL As EditPanel, ByVal pAmount As String)
    Dim vAmount As Double = DoubleValue(pAmount)
    Dim vFound As Boolean
    Dim vTable As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardLetterBreaks)
    If vTable IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        If CType(vRow("Low"), Double) <= vAmount And CType(vRow("High"), Double) >= vAmount Then
          pEPL.SetValue("Mailing", vRow("Mailing").ToString)
          vFound = True
          Exit For
        End If
      Next
      If Not vFound Then
        If mvTA.DefaultMailingType = DefaultMailingTypes.dmtLetterBreaksOrSource Then
          pEPL.SetValue("Mailing", mvCurrentPage.EditPanel.FindTextLookupBox("Source").GetDataRowItem("ThankYouLetter"))
        Else
          pEPL.SetValue("Mailing", "")
        End If
      End If
    End If
  End Sub

  Private Sub SetOtherContactInfo(ByVal pContactInfo As ContactInfo)
    Dim vControl As Control

    vControl = FindControl(mvCurrentPage.EditPanel, "Status", False)
    If vControl IsNot Nothing AndAlso vControl.Visible Then vControl.Text = pContactInfo.StatusDesc
    vControl = FindControl(mvCurrentPage.EditPanel, "Activity", False)
    If vControl IsNot Nothing AndAlso vControl.Visible Then
      Dim vList As New ParameterList(True)
      vList("Current") = "Y"
      DirectCast(vControl, DisplayGrid).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCategories, pContactInfo.ContactNumber, vList))
    End If
    vControl = FindControl(mvCurrentPage.EditPanel, "Relationship", False)
    If vControl IsNot Nothing AndAlso vControl.Visible Then
      DirectCast(vControl, DisplayGrid).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, pContactInfo.ContactNumber))
    End If
    vControl = FindControl(mvCurrentPage.EditPanel, "TransactionNumber", False)
    If vControl IsNot Nothing AndAlso vControl.Visible Then
      DirectCast(vControl, DisplayGrid).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, pContactInfo.ContactNumber))
    End If
    If mvTA.SourceFromLastMailing Then
      vControl = FindControl(mvCurrentPage.EditPanel, "Source", False)
      If vControl IsNot Nothing Then
        Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactSourceFromLastMailing, pContactInfo.ContactNumber))
        If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
          DirectCast(vControl, TextLookupBox).Text = vTable.Rows(0).Item("Source").ToString
          vControl = FindControl(mvCurrentPage.EditPanel, "Mailing", False)
          If vControl IsNot Nothing Then
            DirectCast(vControl, TextLookupBox).Text = vTable.Rows(0).Item("Mailing").ToString
          End If
        End If
      End If
    End If
    vControl = FindControl(mvCurrentPage.EditPanel, "DateOfBirth", False)
    If vControl IsNot Nothing Then mvCurrentPage.EditPanel.SetValue("DateOfBirth", pContactInfo.DateOfBirth, False)
    vControl = FindControl(mvCurrentPage.EditPanel, "ContactStatus", False)
    If vControl IsNot Nothing Then mvCurrentPage.EditPanel.SetValue("ContactStatus", pContactInfo.Status, False)
    vControl = FindControl(mvCurrentPage.EditPanel, "StatusReason", False)
    If vControl IsNot Nothing Then mvCurrentPage.EditPanel.SetValue("StatusReason", pContactInfo.StatusReason, False)
  End Sub

  Private Sub SetMemberContactInfo(ByVal pContactInfo As ContactInfo)
    If mvTA.TransactionType = "MEMB" AndAlso mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).DefaultsSet = True Then
      Dim vEPL As EditPanel = mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).EditPanel
      If BooleanValue(vEPL.GetValue("GiftMembership")) = False Then
        Dim vContactNumber As Integer = IntegerValue(vEPL.GetValue("ContactNumber"))
        If vContactNumber <> pContactInfo.ContactNumber Then
          'Reset Contact details
          vEPL.SetValue("ContactNumber", pContactInfo.ContactNumber.ToString)
          vEPL.SetValue("AddressNumber", pContactInfo.AddressNumber.ToString)
          vEPL.SetValue("Branch", pContactInfo.Branch)
          vEPL.SetValue("DateOfBirth", pContactInfo.DateOfBirth)
          vEPL.SetValue("DobEstimated", CBoolYN(pContactInfo.DOBEstimated))
        End If
      End If
    End If
  End Sub

  Private Sub SetPPDMaintenanceBalance(ByVal pEPL As EditPanel, ByVal pAmount As String)
    If FindControl(pEPL, "Balance", False) IsNot Nothing Then
      If pEPL.FindTextLookupBox("Product").IsValid AndAlso pEPL.FindTextLookupBox("Rate").IsValid Then
        Dim vList As New ParameterList(True)
        vList("Amount") = pAmount
        vList("Product") = pEPL.GetValue("Product")
        vList("Rate") = pEPL.GetValue("Rate")
        vList.IntegerValue("PaymentPlanNumber") = mvTA.PaymentPlan.PaymentPlanNumber
        If FindControl(pEPL, "EffectiveDate", False) IsNot Nothing Then vList("EffectiveDate") = pEPL.GetValue("EffectiveDate")
        Dim vContactNumber As Integer = mvTA.PayerContactNumber
        vList("ContactNumber") = vContactNumber.ToString
        Dim vPFCode As String = GetPageValue(CareNetServices.TraderPageType.tpPaymentPlanMaintenance, "PaymentFrequency")
        If Not String.IsNullOrWhiteSpace(vPFCode) Then vList("PaymentFrequency") = vPFCode
        Dim vReturnList As ParameterList = DataHelper.GetDetailBalance(vList)
        pEPL.SetValue("Balance", DoubleValue(vReturnList("Balance").ToString).ToString("0.00"))
        mvTA.PaymentPlanDetailsPricing.InitFromParameterList(vReturnList)
      End If
    End If
  End Sub

  Private Function GetStringValue(ByVal pObject As Object) As String
    If pObject Is Nothing Then
      Return ""
    Else
      Return pObject.ToString
    End If
  End Function

  Private Sub CalculateEventBookingPrice(ByVal pEPL As EditPanel, ByVal pDisplayPricingBreakdown As Boolean)
    Dim vDone As Boolean
    Try
      pEPL.SetErrorField("Amount", "")
      Dim vEventInfo As CareEventInfo = pEPL.FindTextLookupBox("EventNumber").CareEventInfo
      If vEventInfo IsNot Nothing AndAlso vEventInfo.EventNumber > 0 AndAlso vEventInfo.EventPricingMatrix.Length > 0 Then
        'Need certain data in order to calculate the Price:
        'MANDATORY EventNumber,Quantity,ContactNumber
        'OPTIONAL  AdultQuantity,ChildQuantity,StartTime,EndTime
        Dim vBookingPrice As Double
        Dim vList As New ParameterList(True, True)
        vList.IntegerValue("EventNumber") = vEventInfo.EventNumber
        vList.IntegerValue("ContactNumber") = IntegerValue(pEPL.GetValue("ContactNumber"))
        vList.IntegerValue("Quantity") = IntegerValue(pEPL.GetValue("Quantity"))
        If FindControl(pEPL, "StartTime", False) IsNot Nothing Then
          vList("StartTime") = pEPL.GetValue("StartTime")
          vList("EndTime") = pEPL.GetValue("EndTime")
        End If
        Dim vGotQuantities As Boolean
        If FindControl(pEPL, "AdultQuantity", False) IsNot Nothing Then
          'Only include these values if the fields are visible
          If Not (pEPL.FindTextBox("AdultQuantity").Visible = False AndAlso pEPL.FindTextBox("ChildQuantity").Visible = False) Then
            'One or both controls are visible so use their values
            vList.IntegerValue("AdultQuantity") = IntegerValue(pEPL.GetValue("AdultQuantity"))
            vList.IntegerValue("ChildQuantity") = IntegerValue(pEPL.GetValue("ChildQuantity"))
            vGotQuantities = True
          End If
        End If
        If vList.IntegerValue("EventNumber") > 0 AndAlso vList.IntegerValue("ContactNumber") > 0 AndAlso vList.IntegerValue("Quantity") > 0 Then
          If (vGotQuantities = True AndAlso vList.IntegerValue("ChildQuantity") > 0 AndAlso pEPL.GetValue("AdultQuantity").Length > 0) OrElse vGotQuantities = False Then
            'Jira 383: pass flag to CalculateEventBookingPricefromMatrix to set whether to return event booking price breakdown
            Dim vDisplayPricingBreakdown As Boolean = False
            If pDisplayPricingBreakdown AndAlso FindControl(pEPL, "DisplayPricingBreakdown", False) IsNot Nothing AndAlso BooleanValue(pEPL.GetValue("DisplayPricingBreakdown")) Then
              vList("DisplayPricingBreakdown") = "Y"
              vDisplayPricingBreakdown = True
            End If
            Dim vResult As ParameterList = DataHelper.CalculateEventBookingPricefromMatrix(vList)
            vBookingPrice = DoubleValue(vResult("TotalBookingPrice"))
            pEPL.SetValue("Amount", vBookingPrice.ToString("0.00"), True)
            DataHelper.SetEventBookingLines(mvTA.EventBookingDataSet, vResult)
            If vDisplayPricingBreakdown Then
              Dim vEventPricingDataSet As DataSet = DataHelper.SetEventPricingLines(vResult)
              If vEventPricingDataSet.Tables("DataRow").Rows.Count > 0 Then
                Dim vfrmEventPricingBreakdown As New frmSelectItems(vEventPricingDataSet, frmSelectItems.SelectItemsTypes.sitEventBookingPricingBreakdown, vResult)
                vfrmEventPricingBreakdown.ShowDialog()
              End If
            End If
            vDone = True
          End If
        End If
      Else
        SetAmount(pEPL)
        vDone = True
      End If
    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enEventPricingMatrixCannotBeFound Then
        ShowErrorMessage(vCareEX.Message)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      If vDone = False Then
        pEPL.SetValue("Amount", "", True)
        pEPL.SetErrorField("Amount", InformationMessages.ImUnableToCalcBookingAmount, True)
      End If
    End Try
  End Sub

  Private Sub CalculateEventBookingPrice(ByVal pEPL As EditPanel)
    CalculateEventBookingPrice(pEPL, True)
  End Sub

  Private Sub CalculateExamBookingPrice(ByVal pEPL As EditPanel)
    Dim vDone As Boolean
    Try
      pEPL.SetErrorField("Amount", "")
      Dim vList As New ParameterList(True, True)

      Dim vControl As Control = FindControl(pEPL, "ExamUnitId", False)
      If vControl IsNot Nothing Then
        Dim vChangedList As List(Of ChangedItem) = DirectCast(vControl, ExamSelector).GetChangedList()
        If vChangedList.Count = 0 Then
          'Amount is zero
          pEPL.SetValue("Amount", "0.00", True)
          ExamsDataHelper.SetExamBookingLines(mvTA.ExamBookingDataSet, New ParameterList)
          Exit Sub
        End If
        Dim vUnits As New StringBuilder
        Dim vAddComma As Boolean
        For Each vChange As ChangedItem In vChangedList
          If vAddComma Then vUnits.Append(",")
          vUnits.Append(vChange.Item.UnitID)
          vAddComma = True
        Next
        vList("ExamUnits") = vUnits.ToString
      End If
      Dim vCentreCode As String = pEPL.GetValue("ExamCentreCode")
      If vCentreCode.Length > 0 Then vList("ExamCentreCode") = vCentreCode
      vList("TransactionDate") = mvTA.TransactionDate
      vList.IntegerValue("ContactNumber") = IntegerValue(pEPL.GetValue("ContactNumber"))
      Dim vResult As ParameterList = ExamsDataHelper.CalculateExamBookingPrice(vList)
      Dim vBookingPrice As Double = DoubleValue(vResult("TotalBookingPrice"))
      pEPL.SetValue("Amount", vBookingPrice.ToString("0.00"), True)
      ExamsDataHelper.SetExamBookingLines(mvTA.ExamBookingDataSet, vResult)
      vDone = True
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      If Not vDone Then
        pEPL.SetValue("Amount", "", True)
        pEPL.SetErrorField("Amount", InformationMessages.ImUnableToCalcBookingAmount, True)
      End If
    End Try
  End Sub

  Private Sub SetCreditedContact(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Dim vCContactParam As String = "CreditedContactNumber"
    Dim vCreditedContact As TextLookupBox = pEPL.FindTextLookupBox(vCContactParam, False)
    If vCreditedContact IsNot Nothing AndAlso vCreditedContact.Visible = True Then
      Dim vEnabled As Boolean
      If pParameterName = "LineTypeG" Then
        If BooleanValue(pValue) Then
          'Only enable CreditedContactNumber if HardCredit or SoftCredit are checked
          If BooleanValue(pEPL.GetValue("LineTypeH")) = True OrElse BooleanValue(pEPL.GetValue("LineTypeS")) = True Then vEnabled = True
        Else
          'InMemoriam is unchecked so clear & disable CreditedContactNumber
          vEnabled = False
        End If
      Else
        'If current control is checked then only enable CreditedContactNumber if InMemoriam is checked.
        Dim vOtherType As String = "LineType" & IIf(pParameterName = "LineTypeH", "S", "H").ToString
        If BooleanValue(pValue) = True OrElse BooleanValue(pEPL.GetValue(vOtherType)) = True Then
          If BooleanValue(pEPL.GetValue("LineTypeG")) Then vEnabled = True
        End If
      End If
      pEPL.PanelInfo.PanelItems(vCContactParam).Mandatory = vEnabled
      If vEnabled = False Then
        pEPL.SetValue(vCContactParam, "")
        pEPL.SetErrorField(vCContactParam, "")
      End If
      pEPL.EnableControl(vCContactParam, vEnabled)
    End If
  End Sub

  ''' <summary>Populate Membership information on the <see cref="CareNetServices.TraderPageType.tpTransactionDetails">Transaction Details</see> page when the Member Number has been set.</summary>
  ''' <param name="pParameterName">Name of control being processed.</param>
  ''' <param name="pValue">Value of control being processed.</param>
  Private Sub ShowMemberInfo(ByVal pParameterName As String, ByVal pValue As String)
    Dim vEPL As EditPanel = mvCurrentPage.EditPanel

    If mvCurrentPage.PageType.Equals(CareNetServices.TraderPageType.tpTransactionDetails) _
    AndAlso vEPL.FindPanelControl(Of TextBox)("MemberContactNumber", False) IsNot Nothing Then
      'Always clear the controls to start with
      vEPL.ClearControlList("MemberContactNumber,MemberAddressNumber,MembershipTypeDesc,MembershipCardExpires,RenewalDate,NextPaymentDue,PaymentMethodDesc,StatusDesc")

      If (String.IsNullOrWhiteSpace(pParameterName) = False AndAlso String.IsNullOrWhiteSpace(pValue) = False) _
      AndAlso pParameterName.Equals("MemberNumber", StringComparison.InvariantCultureIgnoreCase) Then
        Dim vMemberTLB As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)(pParameterName, False)
        If vMemberTLB IsNot Nothing AndAlso mvTA.PaymentPlan IsNot Nothing AndAlso vMemberTLB.GetDataRow IsNot Nothing Then
          Dim vMemberContactNumber As Integer = vMemberTLB.GetDataRowInteger("ContactNumber")
          Dim vMemberAddressNumber As Integer = vMemberTLB.GetDataRowInteger("AddressNumber")
          Dim vContactInfo As ContactInfo = Nothing
          If mvTA.PaymentPlan.ContactNumber.Equals(vMemberContactNumber) = False OrElse mvTA.PaymentPlan.AddressNumber.Equals(vMemberAddressNumber) = False Then
            vContactInfo = New ContactInfo(vMemberContactNumber, vMemberAddressNumber)
          Else
            vContactInfo = New ContactInfo(mvTA.PaymentPlan.ContactNumber, mvTA.PaymentPlan.AddressNumber)
          End If
          Dim vList As New ParameterList(True, True)
          vList.IntegerValue("ContactNumber") = vContactInfo.ContactNumber
          vList.IntegerValue("MembershipNumber") = vMemberTLB.GetDataRowInteger("MembershipNumber")
          Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, vContactInfo.ContactNumber, vList))
          Dim vDR As DataRow = Nothing
          If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then vDR = vDT.Rows(0)
          If vDR IsNot Nothing Then
            'We have all the required data so populate the controls
            vEPL.SetValue("MemberContactNumber", vContactInfo.ContactName)
            vEPL.SetValue("MemberAddressNumber", vContactInfo.AddressLine)
            vEPL.SetValue("MembershipTypeDesc", vDR.Item("MembershipTypeDesc").ToString)
            vEPL.SetValue("MembershipCardExpires", vDR.Item("MembershipCardExpires").ToString)
            vEPL.SetValue("RenewalDate", vDR.Item("RenewalDate").ToString)
            vEPL.SetValue("NextPaymentDue", vDR.Item("NextPaymentDue").ToString)
            vEPL.SetValue("PaymentMethodDesc", vDR.Item("PaymentMethodDesc").ToString)
            vEPL.SetValue("StatusDesc", vContactInfo.StatusDesc)    'This is the status of the Member Contact and not the Payer Contact (Payer Contact is displayed separately)
          End If
        End If
      End If
    End If
  End Sub

  Private Sub ShowAlerts(ByVal pEPL As EditPanel, ByVal pContactInfo As ContactInfo)
    If mvTA.ContactAlerts = True OrElse mvTA.TraderAlerts = True Then
      If mvTA.ContactAlerts Then
        pContactInfo.ShowLastStickyNote()
      End If
      Dim vDS As DataSet = pContactInfo.GetContactAlertsForTrader(mvTA.ContactAlerts, mvTA.TraderAlerts, mvTA.ApplicationNumber)
      If vDS IsNot Nothing AndAlso vDS.Tables.Count > 0 AndAlso vDS.Tables(0) IsNot Nothing Then
        Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDS)
        Dim vErrorDT As DataTable = Nothing
        Dim vWarnDT As DataTable = Nothing
        '(1) Get the errors
        vDT.DefaultView.RowFilter = "AlertMessageType = 'E'"
        If vDT.DefaultView.Count > 0 Then
          vErrorDT = vDT.DefaultView.ToTable()
        End If
        '(2) Get the warnings
        vDT.DefaultView.RowFilter = "AlertMessageType = 'W'"
        If vDT.DefaultView.Count > 0 Then
          vWarnDT = vDT.DefaultView.ToTable()
        End If
        '(3) Show the errors
        If vErrorDT IsNot Nothing Then
          Dim vErrorMsg As New StringBuilder()
          For Each vRow As DataRow In vErrorDT.Rows
            vErrorMsg.AppendLine(vRow.Item("AlertMessageDesc").ToString())
          Next
          If vErrorMsg.ToString.Length() > 0 Then
            pEPL.SetErrorField("ContactNumber", InformationMessages.ImInvalidValue, True)
            ShowErrorMessage(vErrorMsg.ToString)
          End If
        End If
        '(4) Finally show the warnings
        If vWarnDT IsNot Nothing Then
          vDS.Tables.Remove(vDT)
          vDS.Tables.Add(vWarnDT)
          Dim vForm As New frmSelectListItem(vDS, frmSelectListItem.ListItemTypes.litAlerts)
          vForm.Text = String.Format("{0}: {1} ({2})", ControlText.FrmContactAlerts, pContactInfo.ContactName, pContactInfo.ContactNumber)
          vForm.ShowDialog()
        End If
      End If
    End If
  End Sub

#End Region

#Region " External Events "

  Private Sub EPL_ContactSelected(ByVal sender As Object, ByVal pContactNumber As Integer)
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub EPL_ValidateItem(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean)
    Dim vEPL As EditPanel = DirectCast(pSender, EditPanel)
    Select Case mvCurrentPage.PageType
      Case CareServices.TraderPageType.tpAccommodationBooking
        Select Case pParameterName
          Case "EventNumber"
            SetValueRaiseChanged(vEPL, "RoomType", "")
            SetValueRaiseChanged(vEPL, "BlockBookingNumber", "")
            SetValueRaiseChanged(vEPL, "Product", "")
            SetValueRaiseChanged(vEPL, "Rate", "")
            SetValueRaiseChanged(vEPL, "Amount", "")
            vEPL.SetErrorField("BlockBookingNumber", "")
            If pValue.Length > 0 Then
              Dim vEventInfo As CareEventInfo = vEPL.FindPanelControl(Of TextLookupBox)("EventNumber").CareEventInfo
              If vEventInfo IsNot Nothing Then
                If vEventInfo.NumberOfAttendees >= vEventInfo.MaximumAttendees Then ShowInformationMessage(InformationMessages.ImEventFullyBooked)
                If vEventInfo.BookingsClosed Then
                  pValid = False
                  vEPL.SetErrorField("EventNumber", InformationMessages.ImEventBookingsClosed, True)
                End If
                If pValid = True Then
                  With vEPL
                    Dim vTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("RoomType")
                    If vTextLookupBox IsNot Nothing AndAlso vTextLookupBox.Visible Then
                      vTextLookupBox.FillComboWithRestriction(vEventInfo.EventNumber.ToString)
                      Dim vProductTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("Product")
                      If vProductTextLookupBox.IsValid = False Then
                        vEPL.SetErrorField("BlockBookingNumber", InformationMessages.ImAccommodationBBInvalidProduct, False)
                      End If
                      If vTextLookupBox.GetDataRow() Is Nothing Then
                        pValid = False
                        vEPL.SetErrorField("EventNumber", InformationMessages.ImEventNoAccommodation, True)
                      End If
                    End If
                  End With
                End If
              End If
            End If
          Case "RoomType"
            Dim vEventInfo As CareEventInfo = vEPL.FindPanelControl(Of TextLookupBox)("EventNumber").CareEventInfo
            If vEventInfo IsNot Nothing Then
              With vEPL
                Dim vTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("BlockBookingNumber")
                If vTextLookupBox IsNot Nothing AndAlso vTextLookupBox.Visible Then
                  Dim vList As ParameterList = New ParameterList(True)
                  vList("EventNumber") = vEventInfo.EventNumber.ToString
                  vList("RoomType") = pValue
                  Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBlockBookings, vList)
                  If vDataTable.Rows.Count = 1 Then
                    vEPL.SetErrorField("BlockBookingNumber", "")
                    SetValueRaiseChanged(vEPL, "BlockBookingNumber", vDataTable.Rows(0).Item("BlockBookingNumber").ToString)
                    Dim vProductTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("Product")
                    If vProductTextLookupBox.IsValid = False Then
                      vEPL.SetErrorField("BlockBookingNumber", InformationMessages.ImAccommodationBBInvalidProduct, False)
                    End If
                  End If
                End If
              End With
            End If
          Case "BlockBookingNumber"
            Dim vBBNTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("BlockBookingNumber")
            If vBBNTextLookupBox IsNot Nothing Then
              Dim vProductTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("Product")
              If vProductTextLookupBox IsNot Nothing Then
                SetValueRaiseChanged(vEPL, "Product", "")
                vEPL.SetErrorField("BlockBookingNumber", "")
                vProductTextLookupBox.Text = vBBNTextLookupBox.GetDataRowItem("Product")
                If vProductTextLookupBox.IsValid = False Then
                  vEPL.SetErrorField("BlockBookingNumber", InformationMessages.ImAccommodationBBInvalidProduct, False)
                End If
                Dim vTextLookupBox As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("Rate")
                If vTextLookupBox IsNot Nothing AndAlso vTextLookupBox.Visible Then
                  vTextLookupBox.FillComboWithRestriction(vBBNTextLookupBox.GetDataRowItem("Product"))
                  SetValueRaiseChanged(vEPL, "Rate", vBBNTextLookupBox.GetDataRowItem("Rate"))
                End If
              End If
              Dim vDTP As DateTimePicker = vEPL.FindPanelControl(Of DateTimePicker)("FromDate")
              If vDTP IsNot Nothing AndAlso Not String.IsNullOrEmpty(vBBNTextLookupBox.GetDataRowItem("FromDate")) Then
                vDTP.Value = Convert.ToDateTime(vBBNTextLookupBox.GetDataRowItem("FromDate")).Date
              End If
              vDTP = vEPL.FindPanelControl(Of DateTimePicker)("ToDate")
              If vDTP IsNot Nothing AndAlso Not String.IsNullOrEmpty(vBBNTextLookupBox.GetDataRowItem("ToDate")) Then
                vDTP.Value = Convert.ToDateTime(vBBNTextLookupBox.GetDataRowItem("ToDate")).Date
              End If
            End If
        End Select

      Case CareServices.TraderPageType.tpBankDetails
        If pParameterName = "AccountNumber" Then
          Dim vValidate As Boolean = (pValue.Length > 0)
          If vValidate = False Then
            vValidate = vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory
          End If
          If vValidate Then
            pValid = vEPL.ValidateAccountNumber(pValue)
          End If
          If pValid = False Then vEPL.SetErrorField("AccountNumber", InformationMessages.ImInvalidAccountNumber, True)
        End If

      Case CareServices.TraderPageType.tpCollectionPayments
        If pParameterName = "PisNumber" Then
          GetCollectionBoxes(vEPL)
        End If

      Case CareServices.TraderPageType.tpContactSelection
        Select Case pParameterName
          Case "ContactNumber"
            If IntegerValue(pValue) > 0 Then
              If vEPL.FindTextLookupBox("ContactNumber").ContactInfo.OwnershipAccessLevel <> ContactInfo.OwnershipAccessLevels.oalWrite Then
                pValid = vEPL.SetErrorField("ContactNumber", GetInformationMessage(InformationMessages.ImNoContactAccess), True)
              End If
            End If
        End Select

      Case CareServices.TraderPageType.tpActivityEntry, CareServices.TraderPageType.tpSuppressionEntry, CareServices.TraderPageType.tpSetStatus, CareServices.TraderPageType.tpGiftAidDeclaration
        Select Case pParameterName
          Case "ContactNumber"
            If IntegerValue(pValue) > 0 Then
              Dim vContactInfo As ContactInfo = vEPL.FindTextLookupBox("ContactNumber").ContactInfo
              If vContactInfo.OwnershipAccessLevel <> ContactInfo.OwnershipAccessLevels.oalWrite Then
                pValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImNoContactAccess, True)
              End If
              If pValid AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpGiftAidDeclaration Then
                If vContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
                  pValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImCannotAddOrganisationGAD, True)
                Else
                  If vContactInfo.AddressNumber = IntegerValue(AppValues.ControlValue(AppValues.ControlValues.non_address_number)) Then
                    pValid = False
                  ElseIf vContactInfo.Status = AppValues.ControlValue(AppValues.ControlValues.gone_away_status) Then
                    pValid = False
                  Else
                    pValid = True
                  End If
                  If pValid = False Then
                    vEPL.SetErrorField("ContactNumber", InformationMessages.ImContactGoneAwayOrUnknownAddress, True)
                  End If
                End If
              End If
            End If
          Case "StartDate", "EndDate"
            Dim vStartDate As String = vEPL.GetValue("StartDate")
            Dim vEndDate As String = vEPL.GetValue("EndDate")
            If vEndDate.Length > 0 Then
              If DateValue(vEndDate) < DateValue(vStartDate) Then
                vEPL.SetValue("EndDate", vStartDate)
                pValid = vEPL.SetErrorField("EndDate", InformationMessages.ImEndDateOnAfterStartDate)
              Else
                vEPL.SetErrorField("EndDate", "")
              End If
            End If
        End Select

      Case CareServices.TraderPageType.tpCreditCardAuthority, CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpStandingOrder
        If mvCurrentPage.PageType <> CareServices.TraderPageType.tpStandingOrder Then
          If pParameterName = "ClaimDay" Then
            If DirectCast(vEPL.FindPanelControl("ClaimDay"), ComboBox).Items.Count = 0 Then
              pValid = vEPL.SetErrorField("ClaimDay", InformationMessages.ImNoClaimDays)
            End If
          End If
        End If
        If pValid = True AndAlso pParameterName = "IbanNumber" Then
          Dim vControl As MaskedTextBox = DirectCast(FindControl(Me, "IbanNumber", False), MaskedTextBox)
          If vControl IsNot Nothing Then
            Try
              Dim vResult As ParameterList = DataHelper.CheckIbanNumber(vControl.Text)
            Catch vException As CareException
              Select Case vException.ErrorNumber
                Case CareException.ErrorNumbers.enInvalidIbanNumber
                  pValid = vEPL.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
              End Select
            End Try
          End If
        End If

        If pValid = True AndAlso pParameterName = "AccountNumber" Then
          Dim vValidate As Boolean = (pValue.Length > 0)
          If vValidate = False Then
            vValidate = vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory
          End If
          If vValidate Then
            pValid = vEPL.ValidateAccountNumber(pValue)
          End If
          If pValid = False Then vEPL.SetErrorField("AccountNumber", InformationMessages.ImInvalidAccountNumber, True)
        End If
        If pParameterName = "StartDate" Then
          Dim vOrderDate As Date
          If mvTA.ApplicationType = ApplicationTypes.atConversion Or mvTA.ApplicationType = ApplicationTypes.atMaintenance Then
            vOrderDate = mvTA.PaymentPlan.StartDate
          ElseIf mvTA.TransactionType = "LOAN" Then
            vOrderDate = New DateHelper(GetPageValue(CareServices.TraderPageType.tpLoans, "OrderDate"), DateHelper.DateHelperNullTypes.dhntNothing).DateValue
          Else
            vOrderDate = New DateHelper(GetPageValue(CareServices.TraderPageType.tpPaymentPlanDetails, "OrderDate"), DateHelper.DateHelperNullTypes.dhntNothing).DateValue
          End If
          If (New DateHelper(pValue, DateHelper.DateHelperNullTypes.dhntNothing).DateValue) < vOrderDate Then
            vEPL.SetErrorField("StartDate", GetInformationMessage(InformationMessages.ImStartDateLTPPStartDate, vOrderDate.ToString), True)
            pValid = False
          End If
        End If

      Case CareServices.TraderPageType.tpCreditCustomer
        Select Case pParameterName
          Case "ContactNumber"
            If IntegerValue(pValue) > 0 Then SetCreditCustomerAccount(vEPL, True, pValue, pValid)
          Case "SalesLedgerAccount"
            If pValue.Length > 0 Then SetCreditCustomerAccount(vEPL, False, pValue, pValid)
        End Select

      Case CareServices.TraderPageType.tpEventBooking, CareServices.TraderPageType.tpAmendEventBooking
        Select Case pParameterName
          Case "EventNumber"
            If pValue.Length > 0 AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking Then
              Dim vEventInfo As CareEventInfo = vEPL.FindTextLookupBox("EventNumber").CareEventInfo
              If vEventInfo IsNot Nothing Then
                With vEPL
                  If FindControl(vEPL, "EventGroup", False) Is Nothing OrElse vEventInfo.EventGroup = GetPageEventGroup(vEPL) Then
                    .SetDateTimeValue("StartDate", vEventInfo.StartDate)
                    .SetValue("EventReference", vEventInfo.EventReference)
                    .SetValue("Location", vEventInfo.Location)
                    .SetValue("VenueDesc", vEventInfo.VenueDesc)
                    If FindControl(vEPL, "StartTime", False) IsNot Nothing Then
                      .SetDateTimeValue("StartTime", vEventInfo.StartTime, True)
                      .SetDateTimeValue("EndTime", vEventInfo.EndTime, True)
                    End If
                    .EnableControl("StartDate", False)
                    If vEventInfo.UserIsOwner Then
                      If vEventInfo.Booking Then
                        If Not vEventInfo.BookingsClosed Then
                          If vEventInfo.NumberOfAttendees >= vEventInfo.MaximumAttendees Then ShowInformationMessage(InformationMessages.ImEventFullyBooked)
                          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions, vEventInfo.EventNumber))
                          If vTable IsNot Nothing AndAlso vTable.Rows.Count = 1 Then
                            SetValueRaiseChanged(vEPL, "OptionNumber", vTable.Rows(0).Item("OptionNumber").ToString)
                          Else
                            .SetValue("OptionNumber", "")
                            .PanelInfo.PanelItems("OptionNumber").ValueChanged("")
                            .SetValue("Product", "")
                            .SetValue("Rate", "")
                          End If
                          If FindControl(vEPL, "WaitingList", False) IsNot Nothing Then
                            If vEventInfo.MaximumOnWaitingList - vEventInfo.NumberOnWaitingList > 0 Then
                              .EnableControl("WaitingList", True)
                            Else
                              .SetValue("WaitingList", "N", True)
                            End If
                            'End If
                          End If
                        Else
                          vEPL.SetErrorField(pParameterName, InformationMessages.ImEventBookingsClosed, True)
                          pValid = False
                        End If
                      Else
                        vEPL.SetErrorField(pParameterName, InformationMessages.ImEventNotBooking, True)
                        pValid = False
                      End If
                    Else
                      vEPL.SetErrorField(pParameterName, InformationMessages.ImEventNotOwner, True)
                      pValid = False
                    End If
                  Else
                    vEPL.SetErrorField(pParameterName, String.Format(InformationMessages.ImEventWrongGroup, vEPL.FindTextLookupBox("EventGroup").Description), True)
                    pValid = False
                  End If
                End With
              End If
            End If

          Case "AdultQuantity", "ChildQuantity"
            Dim vLinkedParameter As String = ""
            Dim vLinkedQuantity As String = ""
            Dim vQuantity As Integer = IntegerValue(pValue)

            vEPL.SetErrorField(pParameterName, "")
            If pParameterName = "AdultQuantity" Then
              vLinkedParameter = "ChildQuantity"
            Else
              vLinkedParameter = "AdultQuantity"
            End If
            vEPL.SetErrorField(vLinkedParameter, "")
            Dim vLinkExists As Boolean = FindControl(vEPL, vLinkedParameter, False) IsNot Nothing
            ValidateQuantity(vEPL, pParameterName, pValue, pValid)
            If vLinkExists Then vQuantity += IntegerValue(vEPL.GetValue(vLinkedParameter))
            If pValid Then
              If Not IntegerValue(pValue) > 0 Then pParameterName = vLinkedParameter
              ValidateQuantity(vEPL, pParameterName, vQuantity.ToString, pValid)
            End If

          Case "EndTime", "StartTime"
            vEPL.SetErrorField("StartTime", "")
            vEPL.SetErrorField("EndTime", "")
            Dim vEventInfo As CareEventInfo = vEPL.FindTextLookupBox("EventNumber").CareEventInfo
            If vEventInfo IsNot Nothing Then
              Dim vStart As New DateHelper(vEventInfo.StartDate.ToShortDateString, vEPL.GetValue("StartTime"))
              Dim vEnd As New DateHelper(vEventInfo.EndDate.ToShortDateString, vEPL.GetValue("EndTime"))
              If vStart.DateValue > vEnd.DateValue Then pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImSessionBookingStartGTEnd)
            End If

        End Select
      Case CareServices.TraderPageType.tpInvoicePayments
        Select Case pParameterName
          Case "SalesLedgerAccount"
            If pValue.Length > 0 Then
              Dim vList As ParameterList = New ParameterList(True)
              vList("Company") = mvTA.CACompany
              vList("SalesLedgerAccount") = pValue
              Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vList)
              If vRow Is Nothing Then
                vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImSLAWrongCompany, mvTA.CACompany), True)
                pValid = False
              End If
            End If
            FillInvoices(DirectCast(mvCurrentPage.EditPanel.FindPanelControl("OSInvoices"), DisplayGrid), mvTA.CACompany, pValue)
            If pValid Then EPL_ValueChanged(vEPL, "CurrentPayment", vEPL.GetValue("CurrentPayment"))
        End Select

      Case CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpChangeMembershipType
        Select Case pParameterName
          Case "AffiliatedMemberNumber"
            'Contact must be a member and only have 1 live membership
            If pValue.Length > 0 Then
              Dim vTLB As TextLookupBox = vEPL.FindTextLookupBox("AffiliatedMemberNumber")
              If vTLB.GetDataRow("CancellationReason").ToString.Length > 0 Then pValid = vEPL.SetErrorField("AffiliatedMemberNumber", InformationMessages.ImMembershipCanc, True)
              If pValid Then
                Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMemberships, IntegerValue(vTLB.GetDataRow("ContactNumber").ToString)))
                vDT.DefaultView.RowFilter = "CancelledOn =''"
                If vDT.DefaultView.Count > 1 Then pValid = vEPL.SetErrorField("AffiliatedMemberNumber", InformationMessages.ImGT1LiveMembForAffiliated, True)
              End If
              Dim vRow As DataRow = vTLB.GetDataRow
              If pValid AndAlso vRow IsNot Nothing Then
                If pValid Then mvTA.SetPayerContact(IntegerValue(vRow.Item("ContactNumber").ToString), IntegerValue(vRow.Item("AddressNumber").ToString))
              End If
            End If
          Case "ContactNumber"
            'This is the member and can not be a joint contact
            If pValue.Length > 0 Then
              Dim vContactInfo As New ContactInfo(IntegerValue(pValue))
              If vContactInfo.ContactType = ContactInfo.ContactTypes.ctJoint Then
                pValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImJointContactCannotBeMember, True)
              End If
              Dim vDeceasedStatus As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.deceased_status)
              If vContactInfo.Status = vDeceasedStatus Then
                pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImNoMembershipWithDeceasedContact, True)
              End If
            End If
          Case "GiverContactNumber"
            If pValue.Length > 0 Then
              If (IntegerValue(pValue) = mvTA.PayerContactNumber) OrElse (IntegerValue(pValue) = IntegerValue(vEPL.GetValue("ContactNumber"))) Then
                pValid = vEPL.SetErrorField("GiverContactNumber", InformationMessages.ImGiverCannotBeMemberOrPayer)
              End If
            End If
          Case "MaxFreeAssociates"
            Dim vRow As DataRow = vEPL.FindTextLookupBox("MembershipType").GetDataRow
            If vRow IsNot Nothing Then
              If vRow.Item("AssociateMembershipType").ToString.Length > 0 Then
                If IntegerValue(pValue) < 1 Then
                  pValid = vEPL.SetErrorField("MaxFreeAssociates", InformationMessages.ImNumberOfAssocGreaterZero, True)
                Else
                  If IntegerValue(pValue) > IntegerValue(vRow.Item("MaxFreeAssociates").ToString) Then
                    If ShowQuestion(InformationMessages.ImNumberAssocGreaterThanFree, MessageBoxButtons.OKCancel, vRow.Item("MaxFreeAssociates").ToString) = System.Windows.Forms.DialogResult.Cancel Then
                      'Set MaxFreeAssociates from MembershipType and set focus back to control
                      vEPL.SetValue("MaxFreeAssociates", vRow.Item("MaxFreeAssociates").ToString)
                      pValid = False
                      Dim vTextBox As TextBox = vEPL.FindTextBox("MaxFreeAssociates")
                      If vTextBox IsNot Nothing AndAlso vTextBox.Visible Then
                        vTextBox.Focus()
                      End If
                    End If
                  End If
                End If
              Else
                If IntegerValue(pValue) > 0 Then
                  pValid = vEPL.SetErrorField("MaxFreeAssociates", InformationMessages.ImNoAssocMembershipType, True)
                End If
              End If
            End If
          Case "MemberNumber"
            If pValue.Length > 0 AndAlso ValidateMemberNumber(pValue) = False Then
              pValid = vEPL.SetErrorField("MemberNumber", InformationMessages.ImInvalidMemberNumberFormat, True)
            End If
          Case "NumberOfMembers"
            If IntegerValue(pValue) < 1 Then
              pValid = vEPL.SetErrorField("NumberOfMembers", InformationMessages.ImNumberOfMembersNotZero)
            Else
              Dim vRow As DataRow = vEPL.FindTextLookupBox("MembershipType").GetDataRow
              If vRow IsNot Nothing Then
                If BooleanValue(vRow.Item("SetNumberOfMembers").ToString) = False Then
                  If IntegerValue(vRow.Item("MembersPerOrder").ToString) > 0 AndAlso (IntegerValue(pValue) > IntegerValue(vRow.Item("MembersPerOrder").ToString)) Then
                    pValid = vEPL.SetErrorField("NumberOfMembers", String.Format(InformationMessages.ImNumberOfMembersExceeded, vRow.Item("MembersPerOrder").ToString), True)
                  End If
                End If
              End If
            End If
          Case "Rate"
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType Then
              If mvTA.PaymentPlan.CMTProportionBalance <> PaymentPlanInfo.CMTProportionBalanceTypes.cmtNone AndAlso (mvTA.PaymentPlan.PayPlanMembershipTypeCode = vEPL.GetValue("MembershipType")) Then
                If pValue = mvTA.PaymentPlan.MembershipRateCode Then
                  pValid = vEPL.SetErrorField("Rate", InformationMessages.ImCMTSameMemberTypeAndRate, True)
                End If
              End If
            End If
          Case "AddressNumber" 'BR18614-Need to update Branch if address is changed
            If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpMembership Then
              Dim vList As New ParameterList(True)
              vList.Add("AddressNumber", pValue)
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetAddressData(CareServices.XMLAddressDataSelectionTypes.xadtAddressInformation, vList))
              If vRow IsNot Nothing AndAlso CStr(vRow("Branch")) IsNot Nothing Then vEPL.FindTextLookupBox("Branch").SetComboString(CStr(vRow("Branch")))
            End If
        End Select

      Case CareServices.TraderPageType.tpPaymentPlanDetails
        Select Case pParameterName
          Case "Balance"
            If vEPL.GetValue("Amount").Length > 0 AndAlso DoubleValue(vEPL.GetValue("Amount")) <> DoubleValue(vEPL.GetValue("Balance")) Then
              vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImBalanceMustEqualAmount), True)
              pValid = False
            End If
          Case "FirstAmount"
            'Value must be between 0.01 and the Balance
            If (mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "CMEM") AndAlso pValue.Length > 0 Then
              If DoubleValue(pValue) < 0.01 OrElse DoubleValue(pValue) > DoubleValue(vEPL.GetValue("Balance")) Then
                pValid = vEPL.SetErrorField("FirstAmount", GetInformationMessage(InformationMessages.ImFirstAmountInvalid), True)
              End If
            End If
          Case "OrderTerm"
            If IntegerValue(pValue) < 1 Then
              vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImPPTermLTOne), True)
              pValid = False
            End If
          Case "OrderDate"
            If mvTA.TransactionType.Equals("MEMB", StringComparison.InvariantCultureIgnoreCase) _
            AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.recalculate_membership_balance) AndAlso IsDate(pValue) Then
              If vEPL.GetValue("PaymentFrequency").Length = 0 Then
                vEPL.SetErrorField("PaymentFrequency", GetInformationMessage(InformationMessages.ImFieldMustNotBeBlank))
              End If
              Try
                GetMemberBalanceAndRenewal(vEPL, True)
                mvTA.mvChangedStartDate = vEPL.GetValue("OrderDate")
              Catch vException As CareException
                Select Case vException.ErrorNumber
                  Case CareException.ErrorNumbers.enMembershipStartDateInvalid
                    ShowErrorMessage(vException.Message)
                    pValid = False
                  Case Else
                    ShowErrorMessage(vException.Message)
                    pValid = False
                End Select
              End Try
            End If
        End Select

      Case CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
        Select Case pParameterName
          Case "Amount"
            If mvTA.LinePrice = 0 Then
              If pValue.Length = 0 Then
                vEPL.SetErrorField("Amount", InformationMessages.ImFixedAmountNotSet)
                pValid = False
              Else
                Dim vFixedAmount As Double = DoubleValue(pValue)
                Dim vBalance As Double = vEPL.GetDoubleValue("Balance")
                Dim vArrears As Double = vEPL.GetDoubleValue("Arrears")
                If ((vFixedAmount >= 0 AndAlso (vFixedAmount < FixTwoPlaces(vBalance - vArrears))) OrElse (vFixedAmount < 0 AndAlso (vFixedAmount > FixTwoPlaces(vBalance + vArrears)))) Then
                  vEPL.SetErrorField("Amount", InformationMessages.ImFixedAmtLTBalArrears)
                  pValid = False
                End If
              End If
              If pValid AndAlso DoubleValue(pValue) <> 0 AndAlso ((mvTA.PaymentPlan.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsExisting) > 0 OrElse (mvTA.PaymentPlan.ProportionalBalanceSetting And PaymentPlanInfo.ProportionalBalanceConfigSettings.pbcsExistingPF) > 0) Then 'And (mvTraderApplication.PaymentPlan.ProportionalBalanceSetting.Equals("pbcsExisting")) Then
                'Fixed Amount > 0 so set Balance to Proportion of Fixed Amount
                If ((mvTA.EditLineNumber > 0 And vEPL.GetDoubleValue("Balance") = 0) OrElse DoubleValue(pValue) = 0) Then 'Dont set the balance to amount if editing a payment plan detail line with balance 0
                Else
                  SetPPDMaintenanceBalance(vEPL, pValue)
                End If
              End If
            End If

          Case "Balance"
            If mvTA.LinePrice = 0 AndAlso vEPL.GetDoubleValue("Amount") < 0 Then
              Dim vFixedAmount As Double = vEPL.GetDoubleValue("Amount")
              Dim vBalance As Double = DoubleValue(pValue)
              Dim vArrears As Double = vEPL.GetDoubleValue("Arrears")
              If ((vBalance > 0) OrElse (vFixedAmount > FixTwoPlaces(vBalance + vArrears))) Then
                vEPL.SetErrorField("Balance", InformationMessages.ImFixedAmtLTBalArrears)
                pValid = False
              End If
            End If
          Case "EffectiveDate"
            ValidateEffectiveDate(vEPL, pValue, pValid)
        End Select


      Case CareServices.TraderPageType.tpPaymentPlanMaintenance
        Select Case pParameterName
          Case "Amount"
            If DoubleValue(pValue) = 0 AndAlso mvTA.PaymentPlan.StandingOrderStatus.Equals("Y") Then
              pValid = False
              vEPL.SetErrorField("Amount", InformationMessages.ImFixedAmtNotSetForSO, True)
            End If
        End Select

      Case CareServices.TraderPageType.tpPaymentPlanProducts
        Select Case pParameterName
          Case "Balance"
            If mvTA.LinePrice = 0 AndAlso vEPL.GetDoubleValue("Amount") < 0 Then
              Dim vFixedAmount As Double = vEPL.GetDoubleValue("Amount")
              Dim vBalance As Double = DoubleValue(pValue)
              Dim vArrears As Double = vEPL.GetDoubleValue("Arrears")
              If ((vBalance > 0) OrElse (vFixedAmount > FixTwoPlaces(vBalance + vArrears))) Then
                vEPL.SetErrorField("Balance", InformationMessages.ImFixedAmtLTBalArrears)
                pValid = False
              End If
            End If

          Case "Product"
            Dim vRow As DataRow = vEPL.FindTextLookupBox("Product").GetDataRow()
            If vRow IsNot Nothing Then
              Dim vSubs As Boolean = vRow("Subscription").ToString = "Y"
              vEPL.EnableControl("CommunicationNumber", vSubs)
            End If
        End Select

      Case CareServices.TraderPageType.tpProductDetails
        Select Case pParameterName
          Case "CreditedContactNumber"
            If pValue.Length > 0 Then
              Dim vDecdContact As Integer = IntegerValue(vEPL.GetValue("DeceasedContactNumber"))
              If vDecdContact > 0 AndAlso (vDecdContact = IntegerValue(pValue)) Then
                pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImDeceasedAndCreditedContactSame)
              End If
            End If
          Case "DeceasedContactNumber"
            Dim vInMemoriam As Boolean
            If FindControl(vEPL, "LineType_G", False) IsNot Nothing Then
              vInMemoriam = (vEPL.GetValue("LineType_G") = "G")
            Else
              vInMemoriam = BooleanValue(vEPL.GetValue("LineTypeG"))
            End If
            mvTA.LastDeceasedContactNumber = IntegerValue(pValue)
            If vInMemoriam = True AndAlso pValue.Length > 0 Then ValidateDeceasedContact(vEPL, pParameterName, pValue)
            If FindControl(vEPL, "CreditedContactNumber", False) IsNot Nothing Then
              Dim vCContact As Integer = IntegerValue(vEPL.GetValue("CreditedContactNumber"))
              If vCContact > 0 AndAlso (vCContact = IntegerValue(pValue)) Then
                pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImDeceasedAndCreditedContactSame)
              End If
            End If
          Case "Warehouse"
            If mvTA.StockSales Then
              'Only create StockMovement if we have previously added one and we have now changed the Warehouse
              If mvTA.StockValuesChanged(vEPL.GetValue("Product"), pValue, IntegerValue(vEPL.GetValue("Quantity")), False) Then
                pValid = AddStockMovement(vEPL.GetValue("Product"), pValue, IntegerValue(vEPL.GetValue("Quantity")))
              End If
            End If
        End Select

        If mvOldProductCode IsNot Nothing AndAlso vEPL.GetValue("Product") <> mvOldProductCode Then
          If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cb_edit_product_rate, False) AndAlso mvTA.EditExistingTransaction = True _
           AndAlso mvTA.BatchInfo.PostedToCashBook AndAlso (mvTA.BatchInfo.BatchType = CareServices.BatchTypes.Cash OrElse mvTA.BatchInfo.BatchType = CareServices.BatchTypes.CashWithInvoice) Then
            vEPL.SetErrorField("Product", InformationMessages.ImCannotChangeProductRate, True)
          End If
        End If
        If mvOldRate IsNot Nothing AndAlso vEPL.GetValue("Rate") <> mvOldRate Then
          If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cb_edit_product_rate, False) AndAlso mvTA.EditExistingTransaction = True _
           AndAlso mvTA.BatchInfo.PostedToCashBook AndAlso (mvTA.BatchInfo.BatchType = CareServices.BatchTypes.Cash OrElse mvTA.BatchInfo.BatchType = CareServices.BatchTypes.CashWithInvoice) Then
            vEPL.SetErrorField("Rate", InformationMessages.ImCannotChangeProductRate, True)
          End If
        End If

      Case CareServices.TraderPageType.tpGiveAsYouEarn
        Select Case pParameterName
          Case "DonorTotal", "EmployerTotal", "GovernmentTotal", "AdminFeesTotal"
            ValidatePTPGTotals(vEPL, pParameterName, pValue, pValid)

          Case "GayePledgeNumber"
            Dim vTLB As TextLookupBox = DirectCast(vEPL.FindPanelControl(pParameterName), TextLookupBox)
            Dim vDataRow As DataRow = vTLB.GetDataRow
            If vDataRow IsNot Nothing Then
              vEPL.SetValue("ContactNumber", vDataRow("ContactNumber").ToString)
              vEPL.EnableControlList("ContactNumber,AddressNumber", False)
              vEPL.SetValue("PostCashBook", vDataRow("PostBatchesToCashBook").ToString)
              If ValidatePledge(vEPL, vTLB) Then
                If vEPL.GetValue("Source").Length = 0 Then vEPL.SetValue("Source", vDataRow("Source").ToString)
              End If
            Else
              vEPL.SetErrorField(pParameterName, "Invalid Pledge Number", True)
            End If
        End Select

      Case CareServices.TraderPageType.tpPostTaxPGPayment
        Select Case pParameterName
          Case "DonorTotal", "EmployerTotal"
            ValidatePTPGTotals(vEPL, pParameterName, pValue, pValid)

          Case "PledgeNumber"
            Dim vTLB As TextLookupBox = DirectCast(vEPL.FindPanelControl(pParameterName), TextLookupBox)
            Dim vDataRow As DataRow = vTLB.GetDataRow
            If vDataRow IsNot Nothing Then
              vEPL.SetValue("ContactNumber", vDataRow("ContactNumber").ToString)
              vEPL.EnableControlList("ContactNumber,AddressNumber", False)
              If ValidatePledge(vEPL, vTLB) Then
                If vEPL.GetValue("Source").Length = 0 Then vEPL.SetValue("Source", vDataRow("Source").ToString)
              End If
            Else
              vEPL.SetErrorField(pParameterName, "Invalid Pledge Number", True)
            End If

        End Select
      Case CareNetServices.TraderPageType.tpPostageAndPacking
        If vEPL.GetValue("Product").Length > 0 Then
          'Rate mandatory if postage and packing product selected.
          If vEPL.GetValue("Rate").Length = 0 Then vEPL.SetErrorField("Rate", "A rate must be entered", True)
        End If
        If vEPL.GetValue("Rate").Length > 0 Then
          If vEPL.GetValue("Product").Length = 0 Then vEPL.SetErrorField("Product", "A product must be entered", True)
        End If

      Case CareNetServices.TraderPageType.tpServiceBooking
        Select Case pParameterName
          Case "ContactGroup"
            If pValue.Length > 0 Then
              pValid = False
              Dim vContactGroup As TextLookupBox = vEPL.FindTextLookupBox(pParameterName)
              Dim vEntityGroup As EntityGroup = Nothing
              If DataHelper.ContactAndOrganisationGroups.ContainsKey(pValue) Then
                vEntityGroup = DataHelper.ContactAndOrganisationGroups(pValue)
                If vEntityGroup.Code = pValue Then
                  GetServiceModifiers(False, False)
                  pValid = True
                End If
              End If

              If Not pValid Then
                vEPL.SetValue(pParameterName, String.Empty)
                vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidValueCheck, True)
              End If
            End If
          Case "BookingContactNumber"
            GetServiceModifiers(False, False)
          Case "ServiceContactNumber"
            If pValue.Length > 0 Then
              Dim vServiceControl As DataRow = Nothing
              Dim vList As New ParameterList(True)
              Dim vContactGroup As String = vEPL.GetOptionalValue("ContactGroup")
              'if the service group is selected then get the service controls
              If vContactGroup.Length > 0 Then
                vList = New ParameterList(True)
                vList("ContactGroup") = vContactGroup
                vServiceControl = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtServiceControl, vList)
              End If
              If vServiceControl IsNot Nothing Then
                pValid = False
                Dim vContact As TextLookupBox = vEPL.FindTextLookupBox(pParameterName)
                If vContact.IsValid Then
                  If vContact.ContactInfo.ContactGroup = vServiceControl("ContactGroup").ToString Then
                    vList = New ParameterList(True)
                    vList("ContactNumber") = pValue
                    Dim vServiceProduct As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtServiceProductContacts, vList)
                    If vServiceProduct IsNot Nothing Then
                      mvTA.FixedUnitRate = BooleanValue(vServiceProduct("FixedUnitRate").ToString)
                      If vEPL.FindPanelControl("Product", False) IsNot Nothing Then
                        vEPL.SetValue("Product", vServiceProduct("Product").ToString)
                      End If
                      If vEPL.FindPanelControl("Rate", False) IsNot Nothing Then
                        vEPL.SetValue("Rate", vServiceProduct("Rate").ToString)
                      End If
                      pValid = True
                    End If
                  Else
                    'This is not a valid contact for the service group
                    vEPL.SetValue(pParameterName, String.Empty)
                  End If
                End If
                If Not pValid Then vEPL.SetErrorField(pParameterName, InformationMessages.ImServiceProductNotFound, True) 'Service Product Not Found
              Else
                vEPL.SetValue(pParameterName, String.Empty)
                vEPL.SetErrorField(pParameterName, InformationMessages.ImServiceNotDefined, True)  'Service Not Defined
              End If
            End If
          Case "StartDate", "EndDate"
            Dim vFromDate As Date
            Dim vToDate As Date
            Dim vDateControl As Control = Nothing
            vDateControl = vEPL.FindPanelControl("StartDate", False)
            If vDateControl IsNot Nothing Then vFromDate = CType(vDateControl, DateTimePicker).Value
            vDateControl = vEPL.FindPanelControl("EndDate", False)
            If vDateControl IsNot Nothing Then vToDate = CType(vDateControl, DateTimePicker).Value
            If vToDate <= vFromDate Then
              If pParameterName <> "EndDate" Then vEPL.SetValue("EndDate", vFromDate.ToString(AppValues.DateTimeFormat)) 'Changing start date
              vEPL.SetErrorField(pParameterName, InformationMessages.ImStartDateGTEndDate) 'Start date cannot be greater then End date
            ElseIf Not mvTA.FixedUnitRate Then
              Dim vAttr As String = String.Empty
              Dim vSource As Control = vEPL.FindPanelControl("Source", False)
              If mvTA.PayerHasDiscount AndAlso (vSource IsNot Nothing AndAlso vSource.Visible) Then
                vAttr = "GrossAmount"
              Else
                vAttr = "Amount"
              End If
              SetAmount(vEPL, vAttr)    'Update the amount since the quantity has changed
            End If
          Case "ProductQuantity"
            Dim vPrice As Double
            Dim vAttr As String = String.Empty
            If vEPL.FindPanelControl("ActivityQuantity", False) IsNot Nothing Then
              If IntegerValue(pValue) > IntegerValue(vEPL.GetValue("ActivityQuantity")) Then
                vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidFreeEntitlementQty, True)  'Cannot use more than free entitlement
                pValid = False
              End If
            End If
            If pValid Then
              vPrice = IntegerValue(vEPL.GetValue("Amount"))
              Dim vSource As Control = vEPL.FindPanelControl("Source", False)
              If mvTA.PayerHasDiscount AndAlso (vSource IsNot Nothing AndAlso vSource.Visible) Then
                vAttr = "GrossAmount"
              Else
                vAttr = "Amount"
              End If
              SetAmount(vEPL, vAttr)    'Update the amount since the quantity has changed
              'If mvTransPrice = 0 Then
              If mvTA.TransactionAmount = 0 Then
                If vEPL.FindPanelControl("Amount", False) IsNot Nothing Then
                  vEPL.SetValue("Amount", vPrice.ToString("0.00"))
                  mvTA.SBGrossAmount = vPrice
                End If
              End If
            End If
        End Select
        If Not mvTA.ServiceBookingCredits Then
          If pParameterName = "ContactGroup" OrElse pParameterName = "ServiceContactNumber" _
           OrElse pParameterName = "StartDate" OrElse pParameterName = "EndDate" Then
            'Clear the flags that are used to check if the warning mesgs have been displayed to the user
            'as the conditions used to set them have changed
            mvTA.ConfirmSBDuration = False
            mvTA.ConfirmSBShortStay = False
            mvTA.ConfirmCalendarConflict = False
          End If
        End If
    End Select

    'Now process items which appear on many pages
    Select Case pParameterName
      Case "CardNumber", "CreditCardNumber"
        pValid = ValidateCardNumber(vEPL, pParameterName)
      Case "ContactNumber"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpGiftAidDeclaration Then
          Dim vContactInfo As ContactInfo = vEPL.FindTextLookupBox(pParameterName).ContactInfo
          If vContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctJoint Then
            pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImGiftAidInvalidContactTypeJoint, True)
          ElseIf vContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctOrganisation Then
            pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImGiftAidInvalidContactTypeOrganisation, True)
          End If
        End If
      Case "ExpiryDate"
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails Then
          'PaymentPlan ExpiryDate
          If New DateHelper(pValue, DateHelper.DateHelperNullTypes.dhntNothing).DateValue < New DateHelper(vEPL.GetValue("OrderDate"), DateHelper.DateHelperNullTypes.dhntNothing).DateValue Then
            pValid = vEPL.SetErrorField("ExpiryDate", InformationMessages.ImExpiryDateBeforeStartDate, True)
          End If
          If pValid = True Then
            If New DateHelper(pValue, DateHelper.DateHelperNullTypes.dhntNothing).DateValue < Now.Date Then
              pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImExpiryDateMustBeInFuture), True)
            End If
          End If
        Else
          'Credit/Debit card ExpiryDate
          If (New DateHelper(pValue, DateHelper.DateHelperCardDateType.dhcdtExpiryDate).DateValue) < DateTime.Now Then
            pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImExpiryDateMustBeInFuture), True)
          End If
        End If
      Case "DateOfBirth"
        Dim vDOB As Date
        If Date.TryParse(pValue, vDOB) Then
          If vDOB > Today Then
            pValid = vEPL.SetErrorField("DateOfBirth", InformationMessages.ImDOBCannotBeInFuture, True)
          End If
        End If

      Case "MemberNumber", "CovenantNumber", "PaymentPlanNumber", "BankersOrderNumber", "DirectDebitNumber", "CreditCardAuthorityNumber"
        Dim vTLB As TextLookupBox
        Dim vDataRow As DataRow = Nothing
        ClearPPNumberChanged(vEPL, pParameterName)
        If PPFieldsBlank(vEPL) Then mvTA.PaymentPlan = Nothing
        If pParameterName = "MemberNumber" Then
          'MemberNumber field on MEM/CMT pages are TextBoxes and do not need to do this
          vTLB = TryCast(vEPL.FindPanelControl(pParameterName), TextLookupBox)
          If vTLB IsNot Nothing Then
            vDataRow = vTLB.GetDataRow
          End If
        Else
          vTLB = DirectCast(vEPL.FindPanelControl(pParameterName), TextLookupBox)
          vDataRow = vTLB.GetDataRow
        End If
        If vDataRow IsNot Nothing Then
          Dim vPPNo As Integer
          If vDataRow.Table.Columns.Contains("PaymentPlanNumber") Then
            vPPNo = IntegerValue(vDataRow.Item("PaymentPlanNumber").ToString)
          Else
            vPPNo = IntegerValue(vDataRow.Item("OrderNumber").ToString)
          End If
          'If mvTA.PaymentPlan Is Nothing Then
          mvTA.PaymentPlan = New PaymentPlanInfo(vPPNo)
          'End If
          'If vtraderpage.PageType = CareServices.TraderPageType.tpTransactionDetails Then ShowMemberInfo(Nothing)
          Dim vError As String = ""
          If pValue.Length > 0 Then
            If pParameterName <> "AffiliatedMemberNumber" Then
              If vDataRow IsNot Nothing Then
                With mvTA.PaymentPlan
                  If .Existing Then
                    pValid = True
                    If pParameterName.Equals("MemberNumber") And pValue.Length > 0 Then ' = "M" Then
                      If vDataRow.Item("CancellationReason").ToString.Length > 0 Then
                        vError = InformationMessages.ImMembershipCanc
                        pValid = False
                      End If
                    End If
                    If .CancellationReason.Length > 0 Then
                      vError = InformationMessages.ImPPCanc
                      pValid = False
                    End If
                    If pValid And (mvCurrentPage.PageType <> CareServices.TraderPageType.tpMembership And mvCurrentPage.PageType <> CareServices.TraderPageType.tpChangeMembershipType) Then
                      If mvCurrentPage.PageType = CareServices.TraderPageType.tpTransactionDetails Or mvCurrentPage.PageType = CareServices.TraderPageType.tpPayments Then
                        If .PaymentScheduleAmendedOn.ToString.Length = 0 Then
                          pValid = False
                          vError = InformationMessages.ImPPNoSchedule
                        End If
                      End If
                      If pValid = True And (mvTA.ApplicationType = ApplicationTypes.atMaintenance Or mvTA.ApplicationType = ApplicationTypes.atConversion) Then
                        'Maintenance/Conversion
                        If .UnprocessedPayments > 0 Then
                          vError = InformationMessages.ImPPUnprocessedPayments
                          pValid = False
                        End If
                      End If
                    End If
                    If pValid Then
                      If mvCurrentPage.PageType = CareServices.TraderPageType.tpPayments And (mvTA.TransactionPaymentMethod = "CAFC" Or mvTA.TransactionPaymentMethod = "VOUC") Then
                        Dim vMsg As String = ""
                        If .PlanType = PaymentPlanInfo.ppType.pptMember Then
                          pValid = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ga_membership_tax_reclaim)
                          vMsg = InformationMessages.ImCAFNotForMemb
                        Else
                          If .NonDonationDetails Then
                            pValid = False
                            vMsg = InformationMessages.ImCAFNotForNonDonation
                          End If
                        End If
                        If Not pValid Then vError = vMsg
                      End If
                    End If
                    If pValid Then
                      'Handle membership info here
                      'If mvPaymentType = "M" And vtraderpage.PageType = CareServices.TraderPageType.tpTransactionDetails Then ShowMemberInfo(vRecordSet)
                      If .FrequencyAmount = 0 And (mvCurrentPage.PageType <> CareServices.TraderPageType.tpContactSelection And mvCurrentPage.PageType <> CareServices.TraderPageType.tpCancelPaymentPlan) Then
                        vError = InformationMessages.ImFreqAmtisZero
                        pValid = False
                      Else
                        If mvTA.TransactionType = "MEMC" Or mvTA.ChangeMembershipType _
                         AndAlso mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then
                          Dim vMembershipNo As Integer = 0
                          If pParameterName = "MemberNumber" Then vMembershipNo = IntegerValue(vDataRow.Item("MembershipNumber").ToString)
                          mvTA.SetCMTValues(IntegerValue(vDataRow.Item("ContactNumber").ToString), IntegerValue(vDataRow.Item("AddressNumber").ToString))
                          If mvTA.TransactionType = "MEMC" Then pValid = ValidateCMT()
                        End If
                        'vepl.SetValue ("Amount"Format$(CalcCurrencyAmount(vepl.getvalue("Amount"), False), "Fixed")
                        'BR9137/9046 (Trader version 1015) changed txt_Change event so that it only processes an active control
                        'but here the amount text box is not active and so ValRequired was being left set to False
                        'when in fact it needs to be True otherwise lblTASAmount is not set  (SAS 05/04/2005)
                        'mvControls(mvCurrPageFirst + vAttrIndex).ValRequired = True
                        'vValid = ValidateControl(vAttrIndex)

                        'mvBalance = vBalance
                        'If vtraderpage.PageType <> CareServices.TraderPageType.tpCancelPaymentPlan Then mvPaymentNumber = txt(pIndex)
                        'vValid = True
                        'If mvPaymentType = "C" Then
                        '  mvCovAmount = vRecordSet.Fields.Item("covenanted_amount").DoubleValue
                        '  mvCovDepositedDeed = vRecordSet.Fields.Item("deposited_deed").Bool
                        'ElseIf vRecordSet.Fields.Item("order_type").Value = "C" Then
                        '  vRecordSet2 = gvConn.GetRecordSet("SELECT covenanted_amount,deposited_deed FROM covenants WHERE order_number = " & vRecordSet.Fields.Item("order_number").Value)
                        '  If vRecordSet2.Fetch = rssOK Then
                        '    mvCovAmount = vRecordSet2.Fields.Item("covenanted_amount").DoubleValue
                        '    mvCovDepositedDeed = vRecordSet2.Fields.Item("deposited_deed").Bool
                        '  Else
                        '    vValid = False
                        '    vError = LoadString(29010)    'Could find the covenants record related to the specified payment plan
                        '  End If
                        '  vRecordSet2.CloseRecordSet()
                        'Else
                        '  mvCovAmount = 0
                        '  mvCovDepositedDeed = False
                        'End If
                      End If
                    End If
                    If mvCurrentPage.PageType = CareServices.TraderPageType.tpContactSelection Then
                      Dim vEnableAutoPay As Boolean = Not (mvTA.PaymentPlan.DirectDebitStatus = "Y" Or mvTA.PaymentPlan.StandingOrderStatus = "Y" Or mvTA.PaymentPlan.CreditCardStatus = "Y")
                      If mvTA.TransactionType = "APAY" AndAlso vEnableAutoPay Then
                        pValid = vEPL.SetErrorField("PaymentPlanNumber", GetInformationMessage(InformationMessages.ImPaymentPlanCancelledOrNonexistentAutoPayMethod), True)
                      End If
                      If mvTA.ApplicationType = ApplicationTypes.atConversion Then
                        If Not vEnableAutoPay AndAlso mvTA.PaymentPlan.PlanType = PaymentPlanInfo.ppType.pptMember AndAlso Not mvTA.PayPlanConvMaintenance Then
                          pValid = vEPL.SetErrorField("PaymentPlanNumber", GetInformationMessage(InformationMessages.ImPPExistingMembAndAutoPay), True)
                        ElseIf mvTA.PaymentPlan.GiftMembership AndAlso mvTA.PaymentPlan.OneYearGift AndAlso mvTA.PaymentPlan.Balance = 0 Then
                          pValid = vEPL.SetErrorField("PaymentPlanNumber", GetInformationMessage(InformationMessages.ImPaymentPlanCannotConvertOneYearGift), True)
                        ElseIf mvTA.PaymentPlan.OneOffPayment = True AndAlso vEnableAutoPay = False AndAlso mvTA.PayPlanConvMaintenance = False Then
                          pValid = vEPL.SetErrorField("PaymentPlanNumber", GetInformationMessage(InformationMessages.ImPaymentPlanCannotConvertOneOff), True)
                        End If
                      End If
                      If pValid = True AndAlso mvTA.PaymentPlan.PlanType = PaymentPlanInfo.ppType.pptLoan AndAlso mvTA.Loans = False Then
                        pValid = vEPL.SetErrorField("PaymentPlanNumber", InformationMessages.ImTraderNotSupportLoans)
                      End If
                      'to clear the error
                      If pValid Then vEPL.SetErrorField("PaymentPlanNumber", "")
                    End If
                    If pValid = True AndAlso mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
                      If .StandingOrderStatus <> "Y" Then
                        pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImPPNoLIveSO, True)
                      End If
                    End If
                  Else
                    pValid = False
                    vError = InformationMessages.ImPaymentPlanCancelledOrNonexistentAutoPayMethod
                  End If
                End With
              End If
            End If
          Else
            pValid = True     'Missing value
          End If
          If pValid = False AndAlso vError.Length > 0 Then vEPL.SetErrorField(pParameterName, vError, True)
        End If
      Case "MembershipType"
        If pValue.Length > 0 Then
          Dim vRow As DataRow = vEPL.FindTextLookupBox("MembershipType").GetDataRow
          If vRow IsNot Nothing Then
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType Then
              If (pValue = mvTA.PaymentPlan.PayPlanMembershipTypeCode) AndAlso mvTA.PaymentPlan.CMTProportionBalance = PaymentPlanInfo.CMTProportionBalanceTypes.cmtNone Then
                pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImCMTSameMemberType, True)
              End If
              If BooleanValue(vRow.Item("ApprovalMembership").ToString) Then
                pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImCMTCannotChangeToApproval, True)
              End If
            End If
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.reason_is_grade) Then
              If vRow.Item("ReasonForDespatch").ToString.Length = 0 Then
                pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImMemberTypeNotReasonForDespatch, True)
              End If
            End If
            If mvCurrentPage.PageType <> CareServices.TraderPageType.tpAmendMembership AndAlso vRow.Item("AssociateType").ToString.Length > 0 Then
              pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImMemberTypeIsAssociateType, True)
            End If
            If vRow.Item("SubsequentMembershipType").ToString.Length > 0 AndAlso vRow.Item("SubsequentTrigger").ToString = "C" Then
              pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImFMTCatTriggerUnsupported)
            End If
            If vRow.Item("AssociateMembershipType").ToString.Length > 0 Then
              Dim vList As New ParameterList(True, True)
              vList.Add("MembershipType", vRow.Item("MembershipType").ToString)
              Dim vAssocRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList).Rows(0)
              If vAssocRow IsNot Nothing Then
                If vAssocRow.Item("SubsequentMembershipType").ToString.Length > 0 AndAlso vAssocRow.Item("SubsequentTrigger").ToString = "C" Then
                  pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImAssocFMTCatTriggerUnsupported)
                End If
              End If
            End If
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpMembership OrElse mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType Then
              If vRow.Item("PayerRequired").ToString = "M" Then
                Dim vAMControl As Control = FindControl(vEPL, "AffiliatedMemberNumber")
                If vAMControl IsNot Nothing AndAlso vAMControl.Visible = False Then
                  pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImAffiliatedMemberNotVisible, True)
                End If
              End If
            End If
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpChangeMembershipType Then
              Dim vList As New ParameterList(True, True)
              vList.Add("MembershipType", mvTA.PaymentPlan.PayPlanMembershipTypeCode)
              Dim vOldMemTypeRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList).Rows(0)
              If vOldMemTypeRow IsNot Nothing Then
                If vOldMemTypeRow.Item("FixedCycle").ToString.Length > 0 AndAlso vRow.Item("FixedCycle").ToString.Length > 0 Then
                  'Can't CMT from a MT w/ a fixed cycle to one w/ a different fixed cycle
                  If vOldMemTypeRow.Item("FixedCycle").ToString <> vRow.Item("FixedCycle").ToString Then
                    pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImCMTFixedCycleChanged, True)
                  End If
                Else
                  'Can't CMT from a MT w/ a fixed cycle to one w/out a fixed cycle
                  'Similarly, can't CMT from a MT w/out a fixed cycle to one w/ a fixed cycle
                  If Not (vOldMemTypeRow.Item("FixedCycle").ToString.Length = 0 AndAlso vRow.Item("FixedCycle").ToString.Length = 0) Then
                    pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImCMTFixedCycleChanged, True)
                  End If
                End If
                If pValid = True AndAlso mvTA.PaymentPlan.HasAutoPaymentMethod Then
                  'Check Annual flag
                  If vOldMemTypeRow.Item("Annual").ToString <> vRow.Item("Annual").ToString Then
                    pValid = vEPL.SetErrorField("MembershipType", InformationMessages.ImCMTCannotChangeAnnual, True)
                  End If
                End If
              End If
            End If
          End If
        End If

      Case "PaymentFrequency"
        Dim vRow As DataRow = vEPL.FindTextLookupBox("PaymentFrequency").GetDataRow()
        If vRow IsNot Nothing Then
          If mvTA.TransactionType = "DONR" And IntegerValue(vRow.Item("Frequency").ToString) <> 1 Then
            vEPL.SetErrorField("PaymentFrequency", InformationMessages.ImRegDonPaidByInst, True)
          Else
            pValid = True
          End If
        End If

      Case "Quantity"
        ValidateQuantity(vEPL, pParameterName, pValue, pValid)

      Case "Source"
        If mvOldSourceCode IsNot Nothing AndAlso vEPL.GetValue("Source") <> mvOldSourceCode Then
          If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cb_edit_source, False) AndAlso mvTA.EditExistingTransaction = True _
           AndAlso mvTA.BatchInfo.PostedToCashBook Then
            vEPL.SetErrorField("Source", InformationMessages.ImCannotChangeSource, True)
          End If
        End If

    End Select
  End Sub

  Private Sub EPL_ValidateAllItems(ByVal pSender As Object, ByVal pList As ParameterList, ByRef pValid As Boolean)
    Dim vEPL As EditPanel = DirectCast(pSender, EditPanel)
    'get the page off the epl as this could be called when the currentpage is different
    Dim vTraderPage As TraderPage = mvTA.TraderPage(vEPL)

    Select Case vTraderPage.PageType
      Case CareServices.TraderPageType.tpAmendMembership
        pValid = ValidateMember()

      Case CareServices.TraderPageType.tpBatchInvoiceProduction
        'InvoiceNumber Range
        If (vEPL.GetValue("InvoiceNumber").Length = 0 And vEPL.GetValue("InvoiceNumber2").Length > 0) _
        Or (vEPL.GetValue("InvoiceNumber").Length > 0 And vEPL.GetValue("InvoiceNumber2").Length = 0) Then
          pValid = False
          vEPL.SetErrorField("InvoiceNumber", InformationMessages.ImInvoiceNumberBlank)
        ElseIf vEPL.GetValue("InvoiceNumber").Length > 0 And vEPL.GetValue("InvoiceNumber2").Length > 0 Then
          If CInt(vEPL.GetValue("InvoiceNumber")) > CInt(vEPL.GetValue("InvoiceNumber2")) Then
            pValid = False
            vEPL.SetErrorField("InvoiceNumber2", InformationMessages.ImInvoiceRangeInvalid)
          End If
        End If
        'Date Range
        Dim vDate As DateTimePicker = TryCast(FindControl(vEPL, "FromDate", False), DateTimePicker)
        If Not vDate Is Nothing Then
          If pValid = True AndAlso vEPL.GetValue("FromDate").Length = 0 AndAlso vEPL.GetValue("ToDate").Length = 0 Then
            If mvTA.mvDateRangeMsgInPrint = True Then
              If ShowQuestion(QuestionMessages.QmDateRangeNotSet, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                pValid = False
              End If
            End If
          End If
        End If
        'BatchNumber Range
        If FindControl(vEPL, "StartBatch", False) IsNot Nothing AndAlso FindControl(vEPL, "EndBatch", False) IsNot Nothing Then
          If (vEPL.GetValue("StartBatch").Length = 0 And vEPL.GetValue("EndBatch").Length > 0) _
          Or (vEPL.GetValue("StartBatch").Length > 0 And vEPL.GetValue("EndBatch").Length = 0) Then
            pValid = False
            vEPL.SetErrorField("StartBatch", InformationMessages.ImInvoiceBatchNumberBlank)
          ElseIf vEPL.GetValue("StartBatch").Length > 0 And vEPL.GetValue("EndBatch").Length > 0 Then
            If CInt(vEPL.GetValue("StartBatch")) > CInt(vEPL.GetValue("EndBatch")) Then
              pValid = False
              vEPL.SetErrorField("EndBatch", InformationMessages.ImInvoiceBatchRangeInvalid)
            End If
          End If
        End If
        'RunType
        Dim vRunType As String = ""
        If FindControl(vEPL, "RunType_N", False) IsNot Nothing Then vEPL.SetErrorField("RunType_N", "") 'Clear any error first
        If GetInvoicePrintRunType(vEPL, vRunType) Then
          'Control is visible so set as invalid if no value
          If vRunType.Length = 0 Then pValid = vEPL.SetErrorField("RunType_N", GetInformationMessage(InformationMessages.ImFieldMandatory))
        End If

      Case CareServices.TraderPageType.tpCardDetails
        'Authorisation
        Dim vTextBox As TextBox = vEPL.FindTextBox("AuthorisationCode")
        If vTextBox.Enabled = True AndAlso vTextBox.Text.Length = 0 Then
          If vEPL.GetValue("CreditOrDebitCard") = "C" Then
            pValid = False
            vEPL.SetErrorField("AuthorisationCode", InformationMessages.ImCardAuthorisation)
          End If
        End If
        'ExpiryDate
        If vEPL.GetValue("ExpiryDate").Length > 0 Then
          Dim vDateTime As DateTime = New DateHelper(vEPL.GetValue("ExpiryDate"), DateHelper.DateHelperCardDateType.dhcdtExpiryDate).DateValue
          If Year(vDateTime) = 1 Then
            pValid = False
            vEPL.SetErrorField("ExpiryDate", InformationMessages.ImExpiryDateMustBeInFuture)
          End If
        End If
        'SecurityCode
        vTextBox = TryCast(FindControl(vEPL, "SecurityCode", False), TextBox) 'Jira 683: Handle SecurityCode not being on the Trader page
        If vTextBox IsNot Nothing AndAlso vTextBox.Enabled Then
          If vTextBox.Text.Length > 0 AndAlso vTextBox.Text.Length < 3 Then
            pValid = False
            vEPL.SetErrorField("SecurityCode", InformationMessages.ImInvalidCardSecurityCode)
          End If
        End If
        'ValidDate
        If vEPL.GetValue("ValidDate").Length > 0 Then
          Dim vDateTime As DateTime = New DateHelper(vEPL.GetValue("ValidDate"), DateHelper.DateHelperCardDateType.dhcdtValidDate).DateValue
          If Year(vDateTime) = 1 Then
            pValid = False
            vEPL.SetErrorField("ValidDate", InformationMessages.ImInvalidCardValidDate)
          End If
        End If
      Case CareServices.TraderPageType.tpChangeMembershipType
        pValid = ValidateCMT()
        If pValid = True AndAlso vEPL.GetValue("Joined").Length = 0 Then
          pValid = vEPL.SetErrorField("Joined", InformationMessages.ImFieldMandatory)
        End If
        If pValid = True Then
          Dim vRow As DataRow = vEPL.FindTextLookupBox("MembershipType").GetDataRow
          If vRow IsNot Nothing Then
            If vRow.Item("AssociateMembershipType").ToString.Length > 0 Then
              If IntegerValue(vEPL.GetValue("MaxFreeAssociates")) < 1 Then
                pValid = vEPL.SetErrorField("MaxFreeAssociates", InformationMessages.ImNumberOfAssocGreaterZero)

              End If
            End If
          End If
        End If
      Case CareServices.TraderPageType.tpContactSelection
        If vEPL.GetValue("PaymentPlanNumber").Length = 0 AndAlso vEPL.GetValue("MemberNumber").Length = 0 _
        AndAlso vEPL.GetOptionalValue("BankersOrderNumber").Length = 0 AndAlso vEPL.GetOptionalValue("DirectDebitNumber").Length = 0 _
        AndAlso vEPL.GetOptionalValue("CreditCardAuthorityNumber").Length = 0 And vEPL.GetOptionalValue("CovenantNumber").Length = 0 Then
          If vEPL.GetValue("ContactNumber").Length = 0 Then
            pValid = False
            vEPL.SetErrorField("ContactNumber", InformationMessages.ImContPPMNNotSpecified)
          Else
            'No payment plan but we have a contact number
            If mvTA.ApplicationType = ApplicationTypes.atMaintenance Or mvTA.ApplicationType = ApplicationTypes.atConversion Then
              vEPL.SetErrorField("PaymentPlanNumber", InformationMessages.ImPPorMNNotSpecified)
              pValid = False
            ElseIf mvTA.ApplicationType = ApplicationTypes.atTransaction AndAlso mvTA.TransactionType = "APAY" Then
              vEPL.SetErrorField("PaymentPlanNumber", InformationMessages.ImPPMNCNSODDOrCCANotSpecified)
              pValid = False
            End If
          End If
        ElseIf IntegerValue(vEPL.GetValue("PaymentPlanNumber")) > 0 Then
          If (mvTA.ApplicationType = ApplicationTypes.atMaintenance OrElse mvTA.ApplicationType = ApplicationTypes.atConversion) AndAlso mvTA.PaymentPlan IsNot Nothing AndAlso mvTA.PaymentPlan.PlanType = PaymentPlanInfo.ppType.pptLoan AndAlso mvTA.Loans = False Then
            vEPL.SetErrorField("PaymentPlanNumber", InformationMessages.ImTraderNotSupportLoans)
            pValid = False
          End If
        End If

      Case CareServices.TraderPageType.tpCreditCustomer
        If vEPL.GetValue("StopCode").Length > 0 Then
          vEPL.SetErrorField("SalesLedgerAccount", InformationMessages.ImSLAAccountOnStop)
          pValid = False
        End If
        If pList("TermsPeriod") = "Y" Then
          pList("TermsPeriod") = "M"
        Else
          pList("TermsPeriod") = "D"
        End If

      Case CareNetServices.TraderPageType.tpDirectDebit, CareNetServices.TraderPageType.tpStandingOrder
        If vEPL.PanelInfo.PanelItems.Exists("IbanNumber") AndAlso vEPL.GetOptionalValue("IbanNumber").Length > 0 Then
          Dim vControl As MaskedTextBox = DirectCast(FindControl(Me, "IbanNumber", False), MaskedTextBox)
          If vControl IsNot Nothing Then
            Try
              Dim vResult As ParameterList = DataHelper.CheckIbanNumber(vEPL.GetOptionalValue("IbanNumber"))
            Catch vException As CareException
              pValid = vEPL.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
            End Try
          End If
        End If
        If pValid Then
          pValid = Utilities.AccountDetailsEntryValid(vEPL, pValid)
        End If
      Case CareServices.TraderPageType.tpEventBooking
        With vEPL
          If .GetOptionalValue("InterestOnly") = "Y" Then
            If .GetOptionalValue("WaitingList") = "Y" Then
              pValid = vEPL.SetErrorField("InterestOnly", InformationMessages.ImEventCannotAddInterestToWaitingList)
            ElseIf .GetValue("InterestBookingNumber").Length > 0 Then
              pValid = .SetErrorField("InterestOnly", InformationMessages.ImCannotConvertInterestToInterest)
            End If
          End If
        End With

      Case CareNetServices.TraderPageType.tpExamBooking
        If vEPL.GetValue("ExamUnitCode").Length = 0 Then
          pValid = vEPL.SetErrorField("ExamUnitCode", InformationMessages.ImFieldMandatory, True)
        End If
        If vEPL.GetValue("ExamSessionCode").Length = 0 Then
          pValid = vEPL.SetErrorField("ExamSessionCode", InformationMessages.ImFieldMandatory, True)
        End If

      Case CareNetServices.TraderPageType.tpLoans
        If vEPL.GetValue("FixedMonthlyAmount").Length = 0 AndAlso vEPL.GetValue("LoanTerm").Length = 0 Then
          pValid = vEPL.SetErrorField("FixedMonthlyAmount", "Either Monthly Payment Amount or Loan Term must be set")
        End If

      Case CareServices.TraderPageType.tpMembership
        pValid = ValidateMember()
        If pValid = True AndAlso vEPL.GetValue("Joined").Length = 0 Then
          pValid = vEPL.SetErrorField("Joined", InformationMessages.ImFieldMandatory)
        End If
        'Check if Source has changed
        If pValid = True AndAlso mvTA.DiscountPercentage > 0 Then
          If (vEPL.GetValue("Source") <> mvTA.LastMembershipSource) AndAlso mvTA.PPBalance > 0 AndAlso mvTraderPages(CareServices.TraderPageType.tpPaymentPlanDetails).DefaultsSet = True Then
            If ShowQuestion(QuestionMessages.QmMemberSourceChanged, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              ClearPageDefaults(CareServices.TraderPageType.tpPaymentPlanDetails)
              'Reset RenewalAmount is held separately from page value
            Else
              vEPL.SetValue("Source", mvTA.LastMembershipSource)
              pValid = False
            End If
          End If
        End If
      Case CareServices.TraderPageType.tpMembershipMembersSummary
        If mvTA.MemberCount = mvTA.CurrentMembers Then
          For vRow As Integer = 0 To mvMembersDGR.RowCount - 1
            pValid = ValidateMember(vRow)
            If pValid = False Then mvMembersDGR.SelectRow(vRow)
            If pValid = False Then Exit For
          Next
        Else
          pValid = False
        End If
      Case CareServices.TraderPageType.tpMembershipPayer
        'ContactNumber
        Dim vRow As Integer = mvMembersDGR.FindRow(mvMembersDGR.GetColumn("ContactNumber"), vEPL.GetValue("ContactNumber"))
        If vRow >= 0 Then
          'Payer is also a Member
          If BooleanValue(mvTraderPages(CareServices.TraderPageType.tpChangeMembershipType.ToString).EditPanel.GetValue("GiftMembership")) Then
            pValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImPayerNotMemberForGiftMembership)
          End If
        End If
      Case CareServices.TraderPageType.tpOutstandingScheduledPayments
        Dim vAmtOutstanding As Double = DoubleValue(vEPL.GetValue("AmountOutstanding"))
        'If Len(lbl(vAttrIndex).Caption) = 0 Then vValid = False
        If vAmtOutstanding > 0 Then
          'Not all of the payment has been allocated against the payment schedule
          'Ask user if they wish to allocate the remainder
          If ShowQuestion(QuestionMessages.QmAmtToInAdv, MessageBoxButtons.YesNo, vAmtOutstanding.ToString) = System.Windows.Forms.DialogResult.No Then
            pValid = False
            vEPL.SetErrorField("AmountOutstanding", GetInformationMessage(InformationMessages.ImAmtStillUnAllocated, vAmtOutstanding.ToString), True)
          End If
        End If
      Case CareServices.TraderPageType.tpPayments
        If Len(vEPL.GetValue("MemberNumber") & vEPL.GetValue("PaymentPlanNumber") & vEPL.GetValue("CovenantNumber")) = 0 Then
          pValid = vEPL.SetErrorField("PaymentPlanNumber", InformationMessages.ImPPMembOrCovNoMandatory)
        ElseIf DoubleValue(vEPL.GetValue("Amount")) = 0 Then
          pValid = vEPL.SetErrorField("Amount", InformationMessages.ImAmountLTOne)
        ElseIf vEPL.GetValue("AcceptAsFull") = "Y" Then
          If mvTA.PaymentPlan IsNot Nothing Then
            Dim vAmount As Double = DoubleValue(vEPL.GetValue("Amount"))
            Dim vBalance As Double
            If mvTA.PaymentPlan.Balance > 0 Then
              vBalance = mvTA.PaymentPlan.Balance
            Else
              vBalance = mvTA.PaymentPlan.FrequencyAmount
            End If
            If Not (vAmount < vBalance And vAmount <> mvTA.PaymentPlan.FrequencyAmount) Then    ' mvFrequencyAmount) Then
              pValid = vEPL.SetErrorField("AcceptAsFull", InformationMessages.ImAmtGTBal)
            End If
          End If
        End If
      Case CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
        If mvTA.ApplicationType = ApplicationTypes.atMaintenance And mvCurrentRow = 0 And (mvTA.TransactionType = "MEMB" Or mvTA.TransactionType = "CMEM") Then '"MAINT"
          'If Maintaining a Membership Order and you change the rate of Line 1 then warn Activities may need to be changed
          'CHUI runs Catagory Maint!!
          If vEPL.GetValue("Rate") <> vEPL.FindTextLookupBox("Rate").OriginalText Then ShowInformationMessage(InformationMessages.ImMembRateChangeActivities)
        End If
      Case CareServices.TraderPageType.tpProductDetails
        If DoubleValue(vEPL.GetValue("Amount")) = 0 Then
          Dim vRow As DataRow = vEPL.FindTextLookupBox("Product").GetDataRow()
          If vRow IsNot Nothing Then
            Dim vDonation As Boolean = vRow("Donation").ToString = "Y"
            If vDonation Then pValid = vEPL.SetErrorField("Amount", InformationMessages.ImDonationAmountCannotBeZero)
          End If
        End If
      Case CareServices.TraderPageType.tpPaymentPlanProducts, CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
        If mvTA.ApplicationType <> ApplicationTypes.atMaintenance And mvTA.ApplicationType <> ApplicationTypes.atConversion Then
          If DoubleValue(vEPL.GetValue("Balance")) = 0 Then
            Dim vRow As DataRow = vEPL.FindTextLookupBox("Product").GetDataRow()
            If vRow IsNot Nothing Then
              Dim vDonation As Boolean = vRow("Donation").ToString = "Y"
              If vDonation Then pValid = vEPL.SetErrorField("Balance", InformationMessages.ImDonationBalanceCannotBeZero)
            End If
          End If
        End If
      Case CareServices.TraderPageType.tpPaymentPlanMaintenance
        If (mvTA.TransactionPaymentMethod = "STDO" OrElse mvTA.PaymentPlan.StandingOrderStatus = "Y") AndAlso DoubleValue(vEPL.GetValue("Amount").ToString) = 0 Then
          pValid = vEPL.SetErrorField("Amount", InformationMessages.ImFixedAmtNotSetForSO)
        End If
      Case CareServices.TraderPageType.tpChequeNumberAllocation
        Dim vFirstRef As Integer = IntegerValue(vEPL.GetValue("ChequeReferenceNumber"))
        Dim vLastRef As Integer = IntegerValue(vEPL.GetValue("ChequeReferenceNumber2"))
        Dim vFirstNo As Integer = IntegerValue(vEPL.GetValue("ChequeNumber"))
        Dim vLastNo As Integer = IntegerValue(vEPL.GetValue("ChequeNumber2"))
        If vLastRef < vFirstRef Or vLastNo < vFirstNo Then
          pValid = vEPL.SetErrorField("ChequeReferenceNumber", InformationMessages.ImInvalidRange)
        ElseIf vLastRef - vFirstRef <> vLastNo - vFirstNo Then
          pValid = vEPL.SetErrorField("ChequeReferenceNumber", InformationMessages.ImRangeNotMatched)
        End If
      Case CareServices.TraderPageType.tpActivityEntry
        Dim vADS As ActivityDataSheet = TryCast(FindControl(vEPL, "Activity", False), ActivityDataSheet)
        If vADS IsNot Nothing Then pValid = vADS.ValidateActivities(vEPL.GetValue("Source"))
      Case CareServices.TraderPageType.tpSuppressionEntry
        Dim vSDS As SuppressionDataSheet = TryCast(FindControl(vEPL, "MailingSuppression", False), SuppressionDataSheet)
        If vSDS IsNot Nothing Then pValid = vSDS.ValidateSuppressions
      Case CareServices.TraderPageType.tpBatchInvoiceSummary
        If mvInvoiceGrid.FindRow("Print", "True") < 0 Then
          'nothing selected
          ShowQuestion(InformationMessages.ImNoInvoicesSelected, MessageBoxButtons.OK)
          pValid = False
        End If
      Case CareServices.TraderPageType.tpBankDetails
        If mvTA.ApplicationType = ApplicationTypes.atCreditListReconciliation Then
          'Check if there are multiple contact accounts with the same account number and sort code
          If Not mvAccountSelected Then SetBankDetails(vEPL, "AccountNumber", vEPL.GetValue("AccountNumber"), mvTA.AlbacsBankDetails)
        Else
          AccountNoVerify(vEPL, vEPL.GetValue("SortCode"), vEPL.GetValue("AccountNumber"), False, mvTA.AlbacsBankDetails)
        End If
        If pValid Then
          pValid = Utilities.AccountDetailsEntryValid(vEPL, pValid)
          Utilities.SetMandatoryControlForBankDetails(vEPL)
          If vEPL.PanelInfo.PanelItems.Exists("IbanNumber") AndAlso vEPL.GetValue("IbanNumber").Length > 0 Then
            Try
              Dim vResult As ParameterList = DataHelper.CheckIbanNumber(vEPL.GetValue("IbanNumber"))
            Catch vException As CareException
              pValid = vEPL.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
            End Try
          End If
        End If

      Case CareNetServices.TraderPageType.tpServiceBooking
        If mvTA.TransactionAmount = 0 Then mvTA.SBGrossAmount = DoubleValue(vEPL.GetValue("Amount"))
        If vEPL.GetDateTimeValue("EndDate") <= vEPL.GetDateTimeValue("StartDate") Then
          pValid = False
          vEPL.SetErrorField("EndDate", InformationMessages.ImStartDateGTEndDate) 'Start date cannot be greater then End date
        End If
    End Select

    'Now process items which appear on many pages
    If pValid Then
      Select Case vTraderPage.PageType
        Case CareServices.TraderPageType.tpChangeMembershipType,
           CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpProductDetails
          Dim vMsg As String = ""
          Dim vAmount As Double = DoubleValue(vEPL.GetValue("Amount").ToString)
          Dim vList As New ParameterList(True, True)
          vList("Product") = vEPL.GetValue("Product")
          vList("Rate") = vEPL.GetValue("Rate")
          Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtRates, vList)
          Dim vLowLimit As String = vRow.Item("CurrentPriceLowerLimit").ToString
          Dim vUpLimit As String = vRow.Item("CurrentPriceUpperLimit").ToString
          Dim vCurrentPrice As Double
          If vRow("UseModifiers").ToString = "Y" Then
            If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpProductDetails Then
              vCurrentPrice = DataHelper.GetModifierPrice(vRow("Product").ToString, vRow("Rate").ToString, CDate(mvTA.TransactionDate), IntegerValue(mvTA.PayerContactNumber))
            Else
              Dim vstrModifierDate As String
              If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpChangeMembershipType Then
                vstrModifierDate = vEPL.GetValue("Joined")
              Else
                vstrModifierDate = mvTA.TransactionDate
                If String.IsNullOrEmpty(vstrModifierDate) Then 'BR17988 - Transaction date may not be present if no payment method is specified
                  vstrModifierDate = Today.ToShortDateString()
                End If
              End If
              Dim vModifierDate As Date = CDate(vstrModifierDate)
              mvTA.SetPaymentPlanDetailsPricing(DataHelper.GetModifierPriceData(vRow("Product").ToString, vRow("Rate").ToString, vModifierDate, IntegerValue(mvTA.PayerContactNumber)))
              vCurrentPrice = mvTA.PaymentPlanDetailsPricing.Price(False)
            End If
          Else
            vCurrentPrice = DoubleValue(vRow.Item("CurrentPrice").ToString())
          End If
          If (vLowLimit.Length > 0 OrElse vUpLimit.Length > 0) AndAlso vCurrentPrice = 0 Then
            If (vLowLimit.Length > 0 AndAlso vUpLimit.Length > 0) Then
              If (vAmount < DoubleValue(vLowLimit)) OrElse (vAmount > DoubleValue(vUpLimit)) Then
                If BooleanValue(vRow.Item("UpperLowerPriceMandatory").ToString) Then
                  ShowWarningMessage(CDBNETCL.My.Resources.InformationMessages.ImPriceOutsidelimits, vLowLimit, vUpLimit)
                  pValid = False
                Else
                  vMsg = GetInformationMessage(QuestionMessages.QmPriceOutsideLimits, vLowLimit, vUpLimit)
                End If
              End If

            ElseIf vUpLimit.Length > 0 Then
              If vAmount > DoubleValue(vUpLimit) Then vMsg = GetInformationMessage(QuestionMessages.QmPriceTooHigh, vUpLimit)
            Else
              If vAmount < DoubleValue(vLowLimit) Then vMsg = GetInformationMessage(QuestionMessages.QmPriceTooLow, vLowLimit)
            End If
            If vMsg.Length > 0 Then
              If ShowQuestion(vMsg, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                pValid = False
              End If
            End If
          End If
      End Select
    End If
  End Sub

  Private Sub EPL_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String)
    Dim vEPL As EditPanel = DirectCast(pSender, EditPanel)
    Dim vTraderPage As TraderPage = mvTA.TraderPage(vEPL)
    '************************************************************************************************************************
    'handle values by Page Type
    '************************************************************************************************************************
    Select Case vTraderPage.PageType
      Case CareServices.TraderPageType.tpAccommodationBooking
        Select Case pParameterName
          Case "BlockBookingNumber"
            Dim vBBNTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("BlockBookingNumber")
            If FindControl(vEPL, "Product", False) IsNot Nothing AndAlso FindControl(vEPL, "Rate", False) IsNot Nothing Then
              SetValueRaiseChanged(vEPL, "Product", vBBNTextLookupBox.GetDataRowItem("Product"))
              Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("Rate")
              vTextLookupBox.FillComboWithRestriction(vBBNTextLookupBox.GetDataRowItem("Product"))
              SetValueRaiseChanged(vEPL, "Rate", vBBNTextLookupBox.GetDataRowItem("Rate"))
            End If
            If FindControl(vEPL, "FromDate", False) IsNot Nothing AndAlso FindControl(vEPL, "ToDate", False) IsNot Nothing Then
              vEPL.SetValue("FromDate", vBBNTextLookupBox.GetDataRowItem("FromDate"))
              vEPL.SetValue("ToDate", vBBNTextLookupBox.GetDataRowItem("ToDate"))
            End If
        End Select

      Case CareServices.TraderPageType.tpActivityEntry
        If (pParameterName = "ContactNumber" OrElse pParameterName = "Source") Then
          PopulateActivityDataSheet(pValue, vEPL, pParameterName)
        End If

      Case CareServices.TraderPageType.tpAmendMembership
        Select Case pParameterName
          Case "DobEstimated"
            If BooleanValue(pValue) Then
              If Not (Date.TryParse(vEPL.GetValue("DateOfBirth"), Nothing)) Then vEPL.SetValue("DateOfBirth", New Date(1901, 1, 1).ToString)
            End If
        End Select

      Case CareServices.TraderPageType.tpBankDetails
        'Set the bank details controls visible based on the config
        Utilities.SetMandatoryControlForBankDetails(vEPL)
        Select Case pParameterName
          Case "AccountNumber", "SortCode"
            SetBankDetails(vEPL, pParameterName, pValue, mvTA.AlbacsBankDetails)
            'Setting this flag to true will prevent the SelectContact dialog from popping up 
            'when the user moves to the next page if the user has already selected the account number
            If pParameterName = "AccountNumber" Then mvAccountSelected = True
          Case "IbanNumber"
            Dim vControl As MaskedTextBox = DirectCast(FindControl(Me, "IbanNumber", False), MaskedTextBox)
            Dim vValid As Boolean = True
            If vControl IsNot Nothing Then
              Try
                Dim vResult As ParameterList = DataHelper.CheckIbanNumber(vControl.Text)
              Catch vException As CareException
                vValid = vEPL.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
              End Try
              If vValid Then SetBankDetails(vEPL, pParameterName, pValue, mvTA.AlbacsBankDetails)
            End If

        End Select

      Case CareNetServices.TraderPageType.tpBatchInvoiceProduction
        Select Case pParameterName
          Case "InvoiceNumber", "InvoiceNumber2", "FromDate", "ToDate"
            Dim vRunType As String = ""
            If GetInvoicePrintRunType(vEPL, vRunType) Then
              If vRunType.Length = 0 Then
                'Controls not set yet so set default value
                If pParameterName = "InvoiceNumber" OrElse pParameterName = "InvoiceNumber2" Then
                  vEPL.SetValue("RunType_R", "R")  'Reprint Invoices
                Else
                  vEPL.SetValue("RunType_N", "N")  'New Invoices
                End If
              End If
            End If
        End Select

      Case CareServices.TraderPageType.tpCardDetails
        Select Case pParameterName
          Case "CreditOrDebitCard"
            SetCreditOrDebitCard(vEPL, pValue)
        End Select

      Case CareServices.TraderPageType.tpBatchInvoiceSummary
        If pParameterName = "SelectAll" AndAlso mvSuppressEvents = False AndAlso mvTA.BatchInvoicesDataSet.Tables.Contains("DataRow") Then
          Dim vLine As Integer
          If BooleanValue(pValue) Then
            For vLine = 0 To mvInvoiceGrid.RowCount - 1
              mvInvoiceGrid.SetValue(vLine, "Print", "True")
            Next
          Else
            For vLine = 0 To mvInvoiceGrid.RowCount - 1
              mvInvoiceGrid.SetValue(vLine, "Print", "False")
            Next
          End If
        End If
      Case CareServices.TraderPageType.tpCancelGiftAidDeclaration
        Select Case pParameterName
          Case "DeclarationNumber"
            vEPL.ClearControlList("DeclarationDate,DeclarationType,StartDate,EndDate")
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetGiftAidData(CType(CareNetServices.XMLGiftAidDataSelectionTypes.xgdtGIftAidCancellationInfo, CareServices.XMLGiftAidDataSelectionTypes), IntegerValue(pValue)))
            If vRow IsNot Nothing AndAlso vRow("CancellationReason").ToString.Length = 0 Then
              If vRow("CanCancel").ToString = "Y" Then
                vEPL.SetValue("DeclarationDate", vRow("DeclarationDate").ToString)
                vEPL.SetValue("DeclarationType", vRow("DeclarationTypeDesc").ToString)
                vEPL.SetValue("StartDate", vRow("StartDate").ToString)
                vEPL.SetValue("EndDate", vRow("EndDate").ToString)
                vEPL.SetValue("ContactNumber", vRow("ContactNumber").ToString, True)
              Else
                vEPL.SetErrorField(pParameterName, vRow("CanCancel").ToString, True)
              End If
            Else
              vEPL.SetErrorField(pParameterName, "Invalid or Cancelled Declaration", True)
            End If
        End Select

      Case CareServices.TraderPageType.tpChequeNumberAllocation
        If pValue.Length > 0 Then
          Dim vList As New ParameterList(True, True)

          Select Case pParameterName
            Case "ChequeReferenceNumber", "ChequeReferenceNumber2"
              vList.IntegerValue("ChequeReferenceNumber") = IntegerValue(pValue)
            Case "ChequeNumber", "ChequeNumber2"
              vList.IntegerValue("ChequeNumber") = IntegerValue(pValue)
          End Select
          If vList.Contains("ChequeReferenceNumber") Or vList.Contains("ChequeNumber") Then
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtChequeInformation, 0, vList))

            If (pParameterName = "ChequeNumber" OrElse pParameterName = "ChequeNumber2") AndAlso vRow IsNot Nothing Then
              vEPL.SetErrorField(pParameterName, InformationMessages.ImChequeAlreadyAllocated, True)
            ElseIf (pParameterName = "ChequeReferenceNumber" OrElse pParameterName = "ChequeReferenceNumber2") Then
              If vRow IsNot Nothing Then
                If vRow("ChequeNumber").ToString.Length > 0 Then vEPL.SetErrorField(pParameterName, InformationMessages.ImChequeAlreadyEntered, True)
              Else
                vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidChequeReferenceNumber, True)
              End If
            End If
          End If
        End If

      Case CareServices.TraderPageType.tpChequeReconciliation
        Select Case pParameterName
          Case "ChequeNumber"
            If pValue.Length > 0 Then
              Dim vList As New ParameterList(True)
              vList("ChequeNumber") = pValue
              vList("ChequeReconciliation") = "Y"
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtChequeInformation, 0, vList))
              If vRow Is Nothing Then
                vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidChequeNumber, True)
              Else
                If vRow("ReconciledOn").ToString.Length > 0 Then
                  vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImChequeAlreadyReconciledOn, vRow("ReconciledOn").ToString), True)
                Else
                  vEPL.SetValue("Amount", vRow("Amount").ToString)
                  vEPL.SetValue("ContactNumber", vRow("ContactNumber").ToString)
                End If
              End If
            End If
        End Select

      Case CareServices.TraderPageType.tpCollectionPayments
        Select Case pParameterName
          Case "AppealCollectionNumber"
            With vEPL.FindTextLookupBox(pParameterName).GetDataRow
              SetValueRaiseChanged(vEPL, "Product", .Item("ProductCode").ToString, True)
              SetValueRaiseChanged(vEPL, "Rate", .Item("RateCode").ToString, True)
              SetValueRaiseChanged(vEPL, "Source", .Item("SourceCode").ToString, True)
              SetValueRaiseChanged(vEPL, "BankAccount", .Item("BankAccountCode").ToString, True)
            End With
            GetCollectionPISNumbers(vEPL, pValue)
            GetCollectionBoxes(vEPL)
        End Select

      Case CareServices.TraderPageType.tpContactSelection
        With vEPL
          If pValue.Length > 0 Then
            Dim vTLB As TextLookupBox = DirectCast(.FindPanelControl(pParameterName), TextLookupBox)
            If vTLB.GetDataRow IsNot Nothing Then
              .SetValue("ContactNumber", vTLB.GetDataRow.Item("ContactNumber").ToString)
              .SetValue("AddressNumber", vTLB.GetDataRow.Item("AddressNumber").ToString)
            End If
          End If
          If Not pParameterName.Equals("ContactNumber") AndAlso Not pParameterName.Equals("AddressNumber") Then .EnableControlList("ContactNumber,AddressNumber", pValue.Length <= 0)
        End With

      Case CareServices.TraderPageType.tpCreditCardAuthority
        Select Case pParameterName
          Case "BankAccount"
            If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
              Dim vCombo As ComboBox = DirectCast(vEPL.FindPanelControl("ClaimDay"), ComboBox)
              If vCombo IsNot Nothing Then
                Dim vDT As DataTable = DirectCast(vCombo.DataSource, DataTable)
                vDT.DefaultView.RowFilter = "BankAccount = '" & pValue & "' AND ClaimType = 'DD'"
                vEPL.ValidateControl("ClaimDay")
              End If
            End If
          Case "AuthorityType"
            ValidateCardNumber(vEPL, "CreditCardNumber")
        End Select

      Case CareServices.TraderPageType.tpCreditCustomer
        Select Case pParameterName
          Case "TermsFrom", "TermsPeriod", "TermsNumber"
            mvTA.CreditTermsChanged = True
            If pParameterName.Equals("TermsPeriod") Then
              Dim vPanelItem As PanelItem = DirectCast(FindControl(vEPL, "TermsNumber").Tag, PanelItem)
              If vEPL.GetValue("TermsPeriod").Equals("Y") Then
                vPanelItem.MinimumValue = "1"
                vPanelItem.MaximumValue = "12"
              Else
                vPanelItem.MinimumValue = "0"
                vPanelItem.MaximumValue = "31"
              End If
              vEPL.ValidateControl("TermsNumber")
            End If
          Case "AddressNumber"
            mvTA.CreditCustomerAddressChanged = True
            Dim vContactNumber As Integer = CInt(vEPL.GetValue("ContactNumber"))
            If mvTA.PayerContactNumber = vContactNumber Then mvTA.SetPayerContact(vContactNumber, CInt(pValue))
          Case "CreditCategory"
            vEPL.SetValue("CreditLimit", vEPL.FindTextLookupBox(pParameterName).GetDataRowItem("CreditLimit"))
        End Select

      Case CareServices.TraderPageType.tpDirectDebit, CareServices.TraderPageType.tpStandingOrder
        'Set the bank details controls visible based on the config
        Utilities.SetMandatoryControlForBankDetails(vEPL)

        Select Case pParameterName
          Case "AccountNumber"
            If (AppValues.DefaultCountryCode <> "CH" AndAlso AppValues.DefaultCountryCode <> "NL") AndAlso vTraderPage.PageType = CareServices.TraderPageType.tpDirectDebit Then
              If pValue.Length < 8 Then
                pValue = "00000000" & pValue
                pValue = pValue.Remove(1, (pValue.Length - 8))
                vEPL.SetValue("AccountNumber", pValue)
              End If
            End If
            SetBankDetails(vEPL, pParameterName, pValue)
          Case "BankAccount"
            If vTraderPage.PageType = CareServices.TraderPageType.tpDirectDebit AndAlso AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
              Dim vCombo As ComboBox = vEPL.FindPanelControl(Of ComboBox)("ClaimDay")
              If vCombo IsNot Nothing Then
                Dim vDT As DataTable = DirectCast(vCombo.DataSource, DataTable)
                DefaultClaimDay(vDT, pValue, "DD")
                vEPL.ValidateControl("ClaimDay")
              End If
            End If
            If pValue.Length > 0 Then
              'Defaults already set so Bank Account has been changed
              If GetOptionalPageValue(CareNetServices.TraderPageType.tpPaymentPlanDetails, "StartMonth").Length = 0 Then
                'Reset the DD/SO start date - but not if using the StartMonth as that date hasn't changed
                Dim vList As New ParameterList(True, True)
                If vTraderPage.PageType.Equals(CareNetServices.TraderPageType.tpDirectDebit) Then
                  vList("PaymentMethod") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dd)
                Else
                  vList("PaymentMethod") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_so)
                End If
                vList("BankAccount") = pValue
                Dim vPPStartDate As String = String.Empty
                If mvTA.TransactionType.ToUpper.Equals("LOAN") Then
                  vPPStartDate = GetPageValue(CareNetServices.TraderPageType.tpLoans, "OrderDate")
                Else
                  vPPStartDate = GetPageValue(CareNetServices.TraderPageType.tpPaymentPlanDetails, "OrderDate")
                  If mvTA.TransactionType.ToUpper.Equals("MEMB") Then
                    vList("MembershipType") = GetPageValue(CareNetServices.TraderPageType.tpMembership, "MembershipType")
                    vList("Joined") = GetPageValue(CareNetServices.TraderPageType.tpMembership, "Joined")
                  End If
                End If
                If String.IsNullOrEmpty(vPPStartDate) Then vPPStartDate = AppValues.TodaysDate
                vList("AutoPayDate") = vPPStartDate
                vPPStartDate = DataHelper.GetPaymentPlanAutoPayDate(vList).ToString(AppValues.DateFormat)
                vEPL.SetValue("StartDate", vPPStartDate)
              End If
            End If
          Case "SortCode", "IbanNumber"
            Dim vValid As Boolean = True
            If pParameterName = "IbanNumber" Then
              Dim vControl As MaskedTextBox = DirectCast(FindControl(Me, "IbanNumber", False), MaskedTextBox)
              If vControl IsNot Nothing Then
                Try
                  Dim vResult As ParameterList = DataHelper.CheckIbanNumber(pValue)
                Catch vException As CareException
                  vValid = vEPL.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
                End Try
              End If
              'If vEPL.PanelInfo.PanelItems.Exists("AccountNumber") AndAlso vEPL.PanelInfo.PanelItems("AccountNumber").Visible = True Then
              '  If vControl.Text.Length > 0 Then
              '    vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory = False
              '    vEPL.PanelInfo.PanelItems("SortCode").Mandatory = False
              '    vEPL.SetErrorField("AccountNumber", "")
              '    vEPL.SetErrorField("SortCode", "")
              '  Else
              '    vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory = True
              '    vEPL.PanelInfo.PanelItems("SortCode").Mandatory = True
              '  End If
              'End If
            End If
            If vValid Then SetBankDetails(vEPL, pParameterName, pValue)
        End Select

      Case CareServices.TraderPageType.tpEventBooking, CareServices.TraderPageType.tpAmendEventBooking
        Select Case pParameterName
          Case "InterestBookingNumber"
            'Convert From Interested Booking
            If pValue.Length > 0 Then
              Dim vRow As DataRow = vEPL.FindTextLookupBox("InterestBookingNumber").GetDataRow
              If vRow IsNot Nothing Then
                SetValueRaiseChanged(vEPL, "EventNumber", vRow("EventNumber").ToString, True)
                SetValueRaiseChanged(vEPL, "OptionNumber", vRow("OptionNumber").ToString)
                SetValueRaiseChanged(vEPL, "Quantity", vRow("Quantity").ToString)
                SetValueRaiseChanged(vEPL, "SalesContactNumber", vRow("SalesContactNumber").ToString)
                SetValueRaiseChanged(vEPL, "Notes", vRow("Notes").ToString)
                If vRow.Table.Columns.Contains("AdultQuantity") = True AndAlso FindControl(vEPL, "AdultQuantity", False) IsNot Nothing Then
                  SetValueRaiseChanged(vEPL, "AdultQuantity", vRow("AdultQuantity").ToString)
                  SetValueRaiseChanged(vEPL, "ChildQuantity", vRow("ChildQuantity").ToString)
                End If
                If vRow.Table.Columns.Contains("StartTime") = True AndAlso FindControl(vEPL, "StartTime", False) IsNot Nothing Then
                  SetValueRaiseChanged(vEPL, "StartTime", vRow("StartTime").ToString)
                  SetValueRaiseChanged(vEPL, "EndTime", vRow("EndTime").ToString)
                End If
              End If
            Else
              vEPL.EnableControl("EventNumber", True)
            End If
          Case "EventGroup"
            Dim vEventNumber As Integer = IntegerValue(vEPL.GetValue("EventNumber"))
            If vEventNumber > 0 Then
              vEPL.SetErrorField("EventNumber", "")
              EPL_ValidateItem(vEPL, "EventNumber", vEventNumber.ToString, True) 'Re-validate the Event
            End If
          Case "OptionNumber"
            SetValueRaiseChanged(vEPL, "Product", vEPL.FindTextLookupBox(pParameterName).GetDataRowItem("ProductCode"), True)
            SetValueRaiseChanged(vEPL, "Rate", vEPL.FindTextLookupBox(pParameterName).GetDataRowItem("RateCode"))
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpEventBooking Then
              If Not vEPL.FindTextLookupBox("Rate").IsValid Then
                Dim vDataTable As DataTable = DirectCast(vEPL.FindTextLookupBox("Rate").ComboBox.DataSource, DataTable)
                If vDataTable IsNot Nothing Then
                  'Rate combo have only one rate then set that rate.
                  If vDataTable.Rows.Count = 2 Then
                    SetValueRaiseChanged(vEPL, "Rate", vDataTable.Rows(1).Item("Rate").ToString)
                  Else
                    SetValueRaiseChanged(vEPL, "Rate", "")
                  End If
                Else
                  SetValueRaiseChanged(vEPL, "Rate", "")
                End If
              End If
            End If
            Dim vMinBookings As Integer = vEPL.FindTextLookupBox(pParameterName).GetDataRowInteger("MinimumBookings")
            Dim vMaxBookings As Integer = vEPL.FindTextLookupBox(pParameterName).GetDataRowInteger("MaximumBookings")
            If vEPL.GetValue("OptionNumber").Length > 0 Then
              If vMaxBookings < vEPL.GetDoubleValue("Quantity") Then
                SetValueRaiseChanged(vEPL, "Quantity", CStr(vMaxBookings))
                vEPL.PanelInfo.PanelItems("Quantity").LastValue = CStr(vMaxBookings)
              ElseIf vMinBookings > vEPL.GetDoubleValue("Quantity") Then
                SetValueRaiseChanged(vEPL, "Quantity", CStr(vMinBookings))
                vEPL.PanelInfo.PanelItems("Quantity").LastValue = CStr(vMinBookings)
              End If
            End If
            vEPL.SetErrorField("Quantity", "")
            Dim vEventInfo As CareEventInfo = vEPL.FindTextLookupBox("EventNumber").CareEventInfo
            If vEventInfo IsNot Nothing Then vEPL.EnableControl("Rate", vEventInfo.EventPricingMatrix.Length = 0)
            mvTA.BookingOptionNumber = pValue
            'Jira 383: Amount field cleared through setting Product/ Rate so need to re-populate Amount field
            'but not trigger display of Pricing Information (for Pricing Matrix Event) as Amount value not changed.
            CalculateEventBookingPrice(vEPL, False)
          Case "EventReference"
            If pValue.Length > 0 Then
              Dim vList As New ParameterList(True)
              vList("EventReference") = pValue
              If vEPL.GetValue("EventNumber").Length > 0 Then vList("EventNumber") = vEPL.GetValue("EventNumber")
              Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftEvents, vList))
              If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
                'This may return multiple Rows, but we always want the last Event (Rich Client orders by start_date desc; the finder orders by start_date)
                Dim vRow As DataRow = vDT.Rows(vDT.Rows.Count - 1)
                SetValueRaiseChanged(vEPL, "EventNumber", vRow("EventNumber").ToString)
              Else
                vEPL.SetValue("EventNumber", "")
              End If
            End If
          Case "EventNumber"
            mvTA.EventNumber = pValue
            If vTraderPage.PageType = CareServices.TraderPageType.tpEventBooking AndAlso FindControl(vEPL, "DisplayPricingBreakdown", False) IsNot Nothing AndAlso FindControl(vEPL, "DisplayPricingBreakdown", False).Visible Then
              Dim vEventPricingMatrix As Boolean = vEPL.FindTextLookupBox("EventNumber").CareEventInfo.EventPricingMatrix.Length > 0
              vEPL.EnableControl("DisplayPricingBreakdown", vEventPricingMatrix)
              vEPL.SetValue("DisplayPricingBreakdown", CBoolYN(vEventPricingMatrix))
            End If
            CalculateEventBookingPrice(vEPL)
          Case "Rate"
            mvTA.EventBookingRate = pValue
          Case "AdultQuantity", "ChildQuantity"
            vEPL.SetErrorField("Quantity", "")
            Dim vLinkedParameter As String = ""
            Dim vLinkedQuantity As String = ""
            If pParameterName = "AdultQuantity" Then
              vLinkedParameter = "ChildQuantity"
            Else
              vLinkedParameter = "AdultQuantity"
            End If
            If FindControl(vEPL, vLinkedParameter, False) IsNot Nothing Then vLinkedQuantity = vEPL.GetValue(vLinkedParameter)
            If pValue.Length > 0 OrElse vLinkedQuantity.Length > 0 Then
              Dim vQuantity As Integer = (IntegerValue(pValue) + IntegerValue(vLinkedQuantity))
              vEPL.SetValue("Quantity", vQuantity.ToString, True)
            Else
              vEPL.EnableControl("Quantity", True)
              vEPL.SetValue("Quantity", "1")
            End If
            CalculateEventBookingPrice(vEPL)
            'SetAmount(vEPL)
          Case "ContactNumber", "Quantity", "StartTime", "EndTime"
            CalculateEventBookingPrice(vEPL)
        End Select

      Case CareNetServices.TraderPageType.tpExamBooking
        ProcessExamSelections(vEPL, pParameterName, pValue)

      Case CareServices.TraderPageType.tpGiftAidDeclaration
        Select Case pParameterName
          Case "DeclarationType", "DeclarationType2"
            If vEPL.FindCheckBox("DeclarationType").Checked = False AndAlso vEPL.FindCheckBox("DeclarationType2").Checked = False Then
              vEPL.SetErrorField(pParameterName, InformationMessages.ImDeclarationTypeNotSpecified, True)
              'vEPL.FindCheckBox(pParameterName).Checked = True
            Else
              vEPL.SetErrorField("DeclarationType", "")
              vEPL.SetErrorField("DeclarationType2", "")
            End If
          Case "Method"
            vEPL.EnableControl("ConfirmedOn", vEPL.FindRadioButton("Method_O").Checked)
        End Select

      Case CareServices.TraderPageType.tpInvoicePayments
        If pParameterName.Equals("CurrentPayment") Then
          If mvCashInvoices IsNot Nothing Then
            Dim vPrevUnallocated As Double = DoubleValue(vEPL.GetValue("CurrentUnAllocated"))
            Dim vUnallocated As Double = DoubleValue(pValue) - mvCashInvoices("0-0").AmountUsed
            If vUnallocated < 0 Then
              vEPL.SetErrorField("CurrentPayment", GetInformationMessage(InformationMessages.ImPaymentGTInvAmtUsed, mvCashInvoices("0-0").AmountUsed.ToString), True)
            Else
              vEPL.SetValue("CurrentUnAllocated", FixTwoPlaces(vUnallocated).ToString("0.00"))
              mvCashInvoices("0-0").InvoiceAmount = DoubleValue(pValue)
            End If
          Else
            vEPL.SetValue("CurrentUnAllocated", FixTwoPlaces(DoubleValue(pValue)).ToString("0.00"))
          End If
        End If

      Case CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpChangeMembershipType
        ProcessMembersPageValuesChanged(vEPL, pParameterName, pValue)

      Case CareServices.TraderPageType.tpProductDetails
        'Changes made to display currency code next to the amount text box
        If vEPL.FindLabel("Amount") IsNot Nothing And Not mvCurrentPage.DefaultsSet Then
          If mvOrgAmountText Is Nothing Then
            mvOrgAmountText = vEPL.FindLabel("Amount").Text
          End If
          If vEPL.FindLabel("Amount").Text IsNot mvOrgAmountText Then vEPL.FindLabel("Amount").Text = mvOrgAmountText
        End If
        Select Case pParameterName
          Case "Discount"
            Dim vAmount As Double
            If mvTA.LinePriceVATEx Then
              vAmount = FixTwoPlaces(vEPL.GetDoubleValue("GrossAmount") - (vEPL.GetDoubleValue("Discount") * (1 + (mvTA.LineVATPercentage / 100))))
            Else
              vAmount = FixTwoPlaces(vEPL.GetDoubleValue("GrossAmount") - vEPL.GetDoubleValue("Discount"))
            End If
            If vAmount > 0 Then
              vEPL.SetValue("Amount", vAmount.ToString("0.00"))
              If FindControl(vEPL, "VatAmount", False) IsNot Nothing Then
                mvTA.LineVATAmount = AppHelper.CalculateVATAmount(vAmount, mvTA.LineVATPercentage)
                vEPL.SetValue("VatAmount", mvTA.LineVATAmount.ToString("0.00"))
                If mvTA.ShowVATExclusiveAmount Then vEPL.SetValue("Amount", FixTwoPlaces(vAmount - mvTA.LineVATAmount).ToString("0.00"))
              End If
            End If
          Case "LineType", "LineTypeG"
            Dim vInMemoriam As Boolean
            If pParameterName = "LineType" Then
              vInMemoriam = (pValue = "G")
            Else
              vInMemoriam = BooleanValue(pValue)
            End If
            vEPL.SetErrorField("DeceasedContactNumber", "")
            If vInMemoriam Then
              'Remove validation
              vEPL.EnableControl("DeceasedContactNumber", True)
              If vEPL.GetValue("DeceasedContactNumber").Length > 0 Then ValidateDeceasedContact(vEPL, pParameterName, pValue)
            Else
              If (pParameterName = "LineType") OrElse (pParameterName = "LineTypeG" AndAlso (BooleanValue(vEPL.GetValue("LineTypeH")) = True OrElse BooleanValue(vEPL.GetValue("LineTypeS")) = True)) Then
                ValidateDeceasedContact(vEPL, pParameterName, pValue)
              Else
                vEPL.SetValue("DeceasedContactNumber", "", True)
              End If
            End If
            If pParameterName = "LineTypeG" Then
              Dim vCreditedContact As TextLookupBox = vEPL.FindTextLookupBox("CreditedContactNumber", False)
              If (vCreditedContact Is Nothing OrElse vCreditedContact.Visible = False) AndAlso BooleanValue(pValue) = True Then
                'If the CreditedContactNumber is hidden or not here then un-check HardCredit (LineTypeH) & SoftCredit (LineTypeS)
                vEPL.SetValue("LineTypeH", "N")
                vEPL.SetValue("LineTypeS", "N")
              End If
              SetCreditedContact(vEPL, pParameterName, pValue)
            End If
          Case "LineTypeH", "LineTypeS"
            'Both cannot be set
            Dim vCreditedContact As TextLookupBox = vEPL.FindTextLookupBox("CreditedContactNumber", False)
            Dim vOtherControl As String = "LineType" & IIf(pParameterName = "LineTypeH", "S", "H").ToString
            vEPL.SetErrorField("DeceasedContactNumber", "")
            If BooleanValue(pValue) Then
              Dim vDecdContact As String = vEPL.GetValue("DeceasedContactNumber")
              vEPL.SetValue(vOtherControl, "N")
              vEPL.EnableControl("DeceasedContactNumber", True)
              If vDecdContact.Length > 0 Then vEPL.SetValue("DeceasedContactNumber", vDecdContact)
              vEPL.PanelInfo.PanelItems("DeceasedContactNumber").Mandatory = True
              'If the CreditedContactNumber is hidden or not here then un-check LineTypeG (InMemoriam)
              If vCreditedContact Is Nothing OrElse vCreditedContact.Visible = False Then vEPL.SetValue("LineTypeG", "N")
              ValidateDeceasedContact(vEPL, pParameterName, pValue)
            ElseIf BooleanValue(vEPL.GetValue(vOtherControl)) = False AndAlso BooleanValue(vEPL.GetValue("LineTypeG")) = False Then
              vEPL.SetValue("DeceasedContactNumber", "", True)
            Else
              ValidateDeceasedContact(vEPL, pParameterName, pValue)
            End If
            SetCreditedContact(vEPL, pParameterName, pValue)
          Case "Warehouse"
            vEPL.SetErrorField("Quantity", String.Empty)
            mvTA.WarehouseChanged = True
        End Select

      Case CareServices.TraderPageType.tpPaymentPlanDetails
        Select Case pParameterName
          Case "OrderDate"
            If vTraderPage.DefaultsSet Then
              Dim vInc As Integer
              Select Case mvTA.TransactionType
                Case "CMEM", "CDON", "CSUB"
                  'vInc = Val(GetPageValue(tpCovenant, "covenant_term"))
                  'If vInc < 4 Then vInc = 4
                Case "MEMB"
                  If mvTA.ApplicationType = ApplicationTypes.atConversion Then
                    vInc = 0
                  Else
                    vInc = 99
                  End If

                  If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.recalculate_membership_balance) Then
                    Dim vValid As Boolean
                    EPL_ValidateItem(vEPL, "OrderDate", vEPL.GetValue("OrderDate"), vValid)
                    'If vValid Then GetMemberBalanceAndRenewal(vEPL)
                  End If

                Case "SALE", "EVNT", "ACOM", "SRVC"    'When called from TPP page
                  vInc = 1
                Case "SUBS", "DONR"
                  Dim vStartMonthCbo As ComboBox = TryCast(FindControl(vEPL, "StartMonth", False), ComboBox)
                  If vStartMonthCbo IsNot Nothing AndAlso vStartMonthCbo.Visible = True Then
                    vInc = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.payment_plan_minimum_term, "0"))
                  End If
                  If vInc = 0 Then vInc = 99
                Case Else
                  vInc = 99
              End Select
              Dim vOrderDate As DateTime = New DateHelper(CDate(pValue)).DateValue
              Dim vExpiryDate As Date = (vOrderDate.AddYears(vInc)).AddDays(-1)
              vEPL.SetValue("ExpiryDate", vExpiryDate.ToString)
              If mvTA.TransactionType = "MEMB" Then
                If AppValues.ConfigurationValue(AppValues.ConfigurationValues.me_membership_price_date) = "START_DATE" Then
                  'StartDate has changed so update the Balance & RenewalAmount
                  Try
                    GetMemberBalanceAndRenewal(vEPL)
                  Catch vException As CareException
                    Select Case vException.ErrorNumber
                      Case CareException.ErrorNumbers.enMembershipStartDateInvalid
                        ShowErrorMessage(vException.Message)
                      Case Else
                        ShowErrorMessage(vException.Message)
                    End Select
                  End Try

                End If
              End If
            End If
          Case "PaymentFrequency"
            If vTraderPage.DefaultsSet AndAlso mvTA.TransactionType = "MEMB" AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.me_membership_price_date) = "START_DATE" Then
              'Payment frequency has changed so update the Balance & RenewalAmount in case it didn't happen on update of start date (nfpcare-1470)
              Try
                GetMemberBalanceAndRenewal(vEPL)
              Catch vException As CareException
                Select Case vException.ErrorNumber
                  Case CareException.ErrorNumbers.enMembershipStartDateInvalid
                    ShowErrorMessage(vException.Message)
                  Case Else
                    ShowErrorMessage(vException.Message)
                End Select
              End Try
            End If
        End Select

      Case CareServices.TraderPageType.tpPurchaseOrderCancellation
        Select Case pParameterName
          Case "PurchaseOrderNumber", "PurchaseOrderNumber2"
            ValidatePurchaseOrderNumber(vEPL, CareServices.TraderPageType.tpPurchaseOrderCancellation, pValue)
        End Select

      Case CareServices.TraderPageType.tpPurchaseOrderDetails, CareServices.TraderPageType.tpPurchaseInvoiceDetails
        Select Case pParameterName
          Case "PurchaseOrderNumber"
            ValidatePurchaseOrderNumber(vEPL, vTraderPage.PageType, pValue)
          Case "ContactNumber", "UseDifferentPayee"
            If pParameterName = "ContactNumber" Then vEPL.SetValue("PayeeContactNumber", pValue)

            'If the user different payee check Box is Checked then clear the default contact Payee
            If vTraderPage.PageType = CareNetServices.TraderPageType.tpPurchaseOrderDetails AndAlso
              FindControl(vEPL, "UseDifferentPayee", False) IsNot Nothing AndAlso vEPL.FindCheckBox("UseDifferentPayee").Checked Then
              vEPL.SetValue("PayeeContactNumber", "")
            Else
              If FindControl(vEPL, "PayeeContactNumber", False) IsNot Nothing Then vEPL.SetValue("PayeeContactNumber", vEPL.GetValue("ContactNumber"))
            End If

            If pParameterName = "ContactNumber" Then
              If vTraderPage.PageType = CareServices.TraderPageType.tpPurchaseOrderDetails Then mvTA.PurchaseOrderScheduleChanged = True
              mvTA.ContactVATCategory = vEPL.FindTextLookupBox("ContactNumber").ContactInfo.VATCategory
            End If
          Case "PurchaseOrderType"
            mvTA.SetPurchaseOrderType(vEPL.FindTextLookupBox("PurchaseOrderType").GetDataRow)
            vEPL.PanelInfo.PanelItems("NumberOfPayments").Mandatory = mvTA.PurchaseOrderType = PurchaseOrderTypes.PaymentSchedule
            vEPL.SetErrorField("NumberOfPayments", "")
            vEPL.FindTextBox("NumberOfPayments").Enabled = mvTA.PurchaseOrderType = PurchaseOrderTypes.PaymentSchedule
            Dim vPayFreq As TextLookupBox = vEPL.FindTextLookupBox("PaymentFrequency")
            vPayFreq.Text = ""
            vPayFreq.SetFilter(If(mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments, "Frequency = '1' OR Frequency=''", ""), True, True)
            vEPL.PanelInfo.PanelItems("PaymentFrequency").Mandatory = mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments
            mvTA.PurchaseOrderScheduleChanged = True  'Always recreate the schedule - change from a non regular payment to a regular payment or vice versa
            vEPL.SetValue("NumberOfPayments", If(mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments, "1", ""))
            vEPL.EnableControl("PaymentAsPercentage", Not mvTA.PurchaseOrderType = PurchaseOrderTypes.RegularPayments)

          Case "NumberOfPayments"
            mvTA.PurchaseOrderScheduleChanged = True
          Case "Campaign"
            If vEPL.FindTextLookupBox("Appeal", False) IsNot Nothing Then vEPL.SetValue("Appeal", "")
            If vEPL.FindTextLookupBox("Segment", False) IsNot Nothing Then vEPL.SetValue("Segment", "")
          Case "Appeal"
            If vEPL.FindTextLookupBox("Segment", False) IsNot Nothing Then vEPL.SetValue("Segment", "")
          Case "PayeeContactNumber", "PayeeAddressNumber"
            If vTraderPage.PageType = CareServices.TraderPageType.tpPurchaseOrderDetails Then mvTA.PurchaseOrderScheduleChanged = True
        End Select

      Case CareServices.TraderPageType.tpPurchaseOrderPayments
        Select Case pParameterName
          Case "DueDate", "LatestExpectedDate"
            Dim vStartDate As Date = Date.Parse(GetPageValue(CareServices.TraderPageType.tpPurchaseOrderDetails, "StartDate"))
            Dim vDueDate As Date = Date.Parse(vEPL.GetValue("DueDate"))
            If vDueDate < vStartDate Then vEPL.SetErrorField("DueDate", InformationMessages.ImDueDateBeforeStartDate, True)
            If vEPL.GetValue("LatestExpectedDate").Length > 0 AndAlso Date.Parse(vEPL.GetValue("LatestExpectedDate")) < vDueDate Then vEPL.SetValue("LatestExpectedDate", vDueDate.ToString(AppValues.DateFormat))
          Case "PoPaymentType"
            Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("PoPaymentType")
            If vTextLookupBox IsNot Nothing AndAlso vTextLookupBox.IsValid AndAlso vEPL.PanelInfo.PanelItems.Exists("PoPaymentType") AndAlso
             vEPL.GetValue("PoPaymentType").Length > 0 AndAlso vEPL.PanelInfo.PanelItems.Exists("DistributionCode") AndAlso vEPL.PanelInfo.PanelItems.Exists("NominalAccount") Then

              Dim vDistributionCode As String = vTextLookupBox.GetDataRowItem("DistributionCode")
              If vDistributionCode.Length > 0 Then vEPL.SetValue("DistributionCode", vDistributionCode)

              Dim vNominalAccount As String = vTextLookupBox.GetDataRowItem("NominalAccount")
              If vNominalAccount.Length > 0 Then vEPL.SetValue("NominalAccount", vNominalAccount)
            End If
        End Select

      Case CareServices.TraderPageType.tpPurchaseOrderProducts, CareServices.TraderPageType.tpPurchaseInvoiceProducts
        Select Case pParameterName
          Case "LinePrice"
            mvTA.LinePrice = DoubleValue(pValue)
            vEPL.SetValue("Amount", pValue)
            vEPL.SetErrorField("Amount", "", False)
        End Select

      Case CareServices.TraderPageType.tpSetStatus
        If pParameterName = "ContactNumber" Then vTraderPage.EditPanel.SetValue("Status", vEPL.FindTextLookupBox("ContactNumber").ContactInfo.Status)

      Case CareServices.TraderPageType.tpSuppressionEntry
        If pParameterName = "ContactNumber" Then PopulateSuppressionDataSheet(pValue, vEPL)

      Case CareServices.TraderPageType.tpTransactionDetails
        Select Case pParameterName
          Case "Amount"
            If Not (mvTA.Voucher OrElse mvTA.CAFCard OrElse mvTA.GiftInKind OrElse mvTA.SaleOrReturn OrElse
              mvTA.FinancialAdjustment <> BatchInfo.AdjustmentTypes.None AndAlso
              mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.GIKConfirmation) AndAlso
              DoubleValue(pValue) >= mvTA.GiftAidMinimum Then vEPL.SetValue("EligibleForGiftAid", "Y")
            If mvTA.DefaultMailingType = DefaultMailingTypes.dmtLetterBreaks Or mvTA.DefaultMailingType = DefaultMailingTypes.dmtLetterBreaksOrSource Then
              SetMailingFromAmount(vEPL, pValue)
            End If
          Case "AddressNumber"
            vEPL.SetValue("MailingAddressNumber", pValue)
          Case "ContactNumber"
            vEPL.SetValue("MailingContactNumber", pValue)
            Dim vContactInfo As ContactInfo = vEPL.FindPanelControl(Of TextLookupBox)(pParameterName).ContactInfo
            mvTA.ContactVATCategory = vContactInfo.VATCategory
            SetOtherContactInfo(vContactInfo)
            SetMemberContactInfo(vContactInfo)
            ShowAlerts(vEPL, vContactInfo)
          Case "Receipt"
            vEPL.SetValue("Mailing", "")
          Case "Source"
            If pValue.Length > 0 Then
              Dim vSourceControl As TextLookupBox = vEPL.FindTextLookupBox("Source")
              mvTA.TransactionDistributionCode = vSourceControl.GetDataRowItem("DistributionCode")
              mvTA.SourceDiscountPercentage = DoubleValue(vSourceControl.GetDataRowItem("DiscountPercentage"))
              If mvTA.DefaultMailingType <> DefaultMailingTypes.dmtLetterBreaks Then
                vEPL.SetValue("Mailing", vSourceControl.GetDataRowItem("ThankYouLetter"))
              End If
              SetPageValue(CareServices.TraderPageType.tpEventBooking, "DistributionCode", vSourceControl.GetDataRowItem("DistributionCode"), True)
              SetPageValue(CareServices.TraderPageType.tpProductDetails, "DistributionCode", vSourceControl.GetDataRowItem("DistributionCode"), True)
            End If
          Case "TransactionDate"
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_fp_prevent_future_date) AndAlso
             Date.Compare(CDate(pValue), Today) > 0 Then
              vEPL.SetErrorField("TransactionDate", InformationMessages.ImTransactionDateInFuture)
            Else
              'Date is valid
              If mvTASDGR IsNot Nothing AndAlso mvTASDGR.RowCount > 0 Then mvTA.TransactionDateChanged = True
              'We have some TAS lines so may need to re-calculate VAT
            End If
          Case "DateOfBirth"                    'BR11822
            mvTA.DOBChanged = vEPL.GetValue("DateOfBirth") <> mvCurrentPage.EditPanel.FindTextLookupBox("ContactNumber").ContactInfo.DateOfBirth
          Case "ProductNumber"
            EPL_ProductNumberSelected(mvCurrentPage.EditPanel, Nothing, IntegerValue(pValue), 0)
        End Select
      Case CareServices.TraderPageType.tpConfirmProvisionalTransactions
        Select Case pParameterName
          Case "ProductNumber"
            EPL_ProductNumberSelected(mvCurrentPage.EditPanel, Nothing, IntegerValue(pValue), 0)
          Case "ContactNumber"
            vEPL.ClearControlList("ProductNumber,ProvisionalBatchNumber,ProvisionalTransNumber,Amount,TransactionDate,Reference")
            EPL_ProductNumberSelected(mvCurrentPage.EditPanel, Nothing, 0, IntegerValue(pValue))
        End Select
      Case CareServices.TraderPageType.tpGiveAsYouEarn, CareServices.TraderPageType.tpPostTaxPGPayment
        Select Case pParameterName
          Case "Source"
            If pValue.Length > 0 Then
              Dim vSourceControl As TextLookupBox = vEPL.FindTextLookupBox("Source")
              mvTA.TransactionDistributionCode = vSourceControl.GetDataRowItem("DistributionCode")
              If mvTA.DefaultMailingType <> DefaultMailingTypes.dmtLetterBreaks Then
                vEPL.SetValue("Mailing", vSourceControl.GetDataRowItem("ThankYouLetter"))
              End If
            End If
          Case "TransactionDate"
            Dim vPledgeName As String = ""
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpGiveAsYouEarn Then
              vPledgeName = "GayePledgeNumber"
            ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpPostTaxPGPayment Then
              vPledgeName = "PledgeNumber"
            End If
            If vPledgeName.Length > 0 Then
              Dim vTLB As TextLookupBox = DirectCast(vEPL.FindPanelControl(vPledgeName), TextLookupBox)
              If vTLB.GetDataRow IsNot Nothing Then
                ValidatePledge(vEPL, vTLB)
              End If
            End If
        End Select
      Case CareServices.TraderPageType.tpLegacyBequestReceipt
        Select Case pParameterName
          Case "LegacyNumber"
            SetValueRaiseChanged(vEPL, "BequestNumber", "")
            Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("LegacyNumber")
            Dim vBequestNumber As String = ""
            If vTextLookupBox.GetDataRow() IsNot Nothing Then
              Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests, vTextLookupBox.GetDataRowInteger("ContactNumber"))
              If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows.Count = 1 Then vBequestNumber = vDataSet.Tables("DataRow").Rows.Item(0).Item("BequestNumber").ToString
            End If
            SetValueRaiseChanged(vEPL, "BequestNumber", vBequestNumber)
          Case "BequestNumber", "Amount"
            Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("BequestNumber")
            If vTextLookupBox.GetDataRow() IsNot Nothing Then
              If pParameterName = "BequestNumber" Then vEPL.SetValue("ExpectedValue", vTextLookupBox.GetDataRowItem("ExpectedValue"))
              vEPL.SetValue("EstimatedOutstanding", FixTwoPlaces(DoubleValue(vTextLookupBox.GetDataRowItem("EstimatedOutstanding")) - DoubleValue(vEPL.GetValue("Amount"))).ToString("0.00"))
              If vEPL.GetDoubleValue("EstimatedOutstanding") < 0 Then vEPL.SetValue("EstimatedOutstanding", FixTwoPlaces(0).ToString("0.00"))
            End If
        End Select

      Case CareNetServices.TraderPageType.tpLoans
        Select Case pParameterName
          Case "FixedMonthlyAmount"
            vEPL.EnableControl("LoanTerm", (pValue.Length = 0))
          Case "LoanAmount"
            Dim vPPBalance As Double = DoubleValue(pValue)
            If mvTA.ApplicationType = ApplicationTypes.atMaintenance OrElse mvTA.ApplicationType = ApplicationTypes.atConversion Then
              vPPBalance = FixTwoPlaces(mvTA.PaymentPlan.Balance + (DoubleValue(pValue) - mvTA.LoanAmount))
            End If
            mvTA.PPBalance = vPPBalance
            mvTraderPages(CareServices.TraderPageType.tpPaymentPlanSummary.ToString).EditPanel.SetValue("PPBalance", vPPBalance.ToString)
          Case "LoanTerm"
            vEPL.EnableControl("FixedMonthlyAmount", (pValue.Length = 0))
        End Select

      Case CareServices.TraderPageType.tpPostageAndPacking
        Select Case pParameterName
          Case "Amount2"
            'Where the Postage and Packing Amount (Amount2) is changed and the Transaction Total is greater than zero recalculate the Percentage otherwise set to zero
            If DoubleValue(vEPL.GetValue("Amount2")) > 0 AndAlso DoubleValue(vEPL.GetValue("Amount")) > 0 Then
              vEPL.SetValue("Percentage", ((DoubleValue(vEPL.GetValue("Amount2")) / DoubleValue(vEPL.GetValue("Amount"))) * 100).ToString("0.00"), pUpdateLastValue:=True)
            Else
              vEPL.SetValue("Percentage", "0.00")
            End If

          Case "Percentage"
            If vEPL.FindPanelControl(Of TextBox)("Percentage").Enabled Then
              If DoubleValue(pValue) > 0 Then
                SetAmount(vEPL)
              Else
                If vEPL.GetValue("Product").Length = 0 AndAlso vEPL.GetValue("Rate").Length = 0 AndAlso DoubleValue(vEPL.GetValue("Amount")) > 0 Then vEPL.SetValue("Percentage", "")
              End If
            End If

          Case "Rate"
            If pValue.Length = 0 Then
              If String.IsNullOrEmpty(vEPL.GetValue("Product")) Then
                'Clear the P&P data
                vEPL.ClearControlList("Amount2,Percentage")
              End If
            End If
        End Select

      Case CareNetServices.TraderPageType.tpServiceBooking
        Select Case pParameterName
          Case "Source"
            Dim vGrossAmt As Control = vEPL.FindPanelControl("GrossAmount", False)
            If vGrossAmt IsNot Nothing AndAlso vEPL.FindPanelControl("Amount", False) IsNot Nothing Then
              If (mvTA.PayerHasDiscount AndAlso vGrossAmt.Visible = False) OrElse
               (mvTA.PayerHasDiscount = False AndAlso vGrossAmt.Visible) Then
                'Field properties need changing
                vEPL.SetControlVisible("GrossAmount", mvTA.PayerHasDiscount)
                vEPL.SetControlVisible("Discount", mvTA.PayerHasDiscount)
                vEPL.EnableControl("Discount", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_discount_amend))
              End If
            End If
          Case "ServiceContactNumber"
            'Jira 1344: Call SetValueRaiseChanged for Rate, which will call SetAmount, on changing the Service Contact Number because Service Contact Number may be linked to a service_products record
            'with a different fixed_unit_rate flag (TraderApplication.FixedUnitRate) which will affect the Amount calculation
            Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("Rate")
            If vTextLookupBox IsNot Nothing AndAlso vTextLookupBox.Text.Length > 0 AndAlso vTextLookupBox.IsValid Then
              SetValueRaiseChanged(vEPL, "Rate", vTextLookupBox.Text)
            End If
        End Select
    End Select

    '************************************************************************************************************************
    'Now process items which appear on many pages
    '************************************************************************************************************************
    Select Case pParameterName
      Case "Amount", "NetFixedAmount"
        If pParameterName = "NetFixedAmount" Then
          'Add any VAT before setting the Balance
          Dim vVATAmount As Double = FixTwoPlaces((DoubleValue(pValue) * (mvTA.LineVATPercentage / 100)))
          pValue = FixTwoPlaces(DoubleValue(pValue) + vVATAmount).ToString
        End If
        Select Case vTraderPage.PageType
          Case CareServices.TraderPageType.tpProductDetails
            Dim vHasPRDVATAmount As Boolean = FindControl(vEPL, "VatAmount", False) IsNot Nothing
            If mvTA.LinePrice = 0 Then
              If mvTA.PayerHasDiscount Then
                SetAmount(vEPL, pParameterName)
              Else
                'Just need to set VAT Amount
                If DoubleValue(pValue) = 0 Then
                  mvTA.LineVATAmount = 0    'Ensure VAT Amount gets reset back to zero
                Else
                  If mvTA.ShowVATExclusiveAmount Then
                    mvTA.LineVATAmount = FixTwoPlaces((DoubleValue(vEPL.GetValue("Amount")) * (mvTA.LineVATPercentage / 100)))
                  Else
                    mvTA.LineVATAmount = AppHelper.CalculateVATAmount(DoubleValue(vEPL.GetValue("Amount")), mvTA.LineVATPercentage)
                  End If
                End If
                If vHasPRDVATAmount Then vEPL.SetValue("VatAmount", mvTA.LineVATAmount.ToString("0.00"), True)
              End If
            End If
          Case CareServices.TraderPageType.tpPaymentPlanDetails, CareServices.TraderPageType.tpPaymentPlanProducts
            vEPL.SetValue("Balance", pValue)
            If pParameterName = "NetFixedAmount" Then vEPL.SetValue("Amount", pValue)
            If vTraderPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanProducts Then SetPPDPricingData(vEPL)
          Case CareServices.TraderPageType.tpPaymentPlanMaintenance
            If DoubleValue(pValue) > 0 And (DoubleValue(pValue) <> mvTA.PaymentPlan.Amount) Then
              'Amount is set and has changed
              'Balance = Amount - (Renewal Amount - Balance)
              Dim vBalance As Double = (DoubleValue(pValue) - (mvTA.PaymentPlan.RenewalAmount - mvTA.PaymentPlan.Balance))
              If vBalance < 0 Then vBalance = 0 'Balance cannot be negative
              SetValueRaiseChanged(vEPL, "Balance", vBalance.ToString("0.00"))
              vEPL.PanelInfo.PanelItems("Balance").LastValue = vBalance.ToString("0.00")    'Set the LastValue so we know if the Balance gets changed back to it's original value
            End If
          Case CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance
            If pParameterName = "NetFixedAmount" Then vEPL.SetValue("Amount", pValue)
            SetPPDPricingData(vEPL)
        End Select
      Case "Balance"
        Select Case vTraderPage.PageType
          Case CareServices.TraderPageType.tpPaymentPlanProducts
            mvTA.PaymentPlanDetailsPricing.UpdateBalance(DoubleValue(pValue))
            If mvTA.LinePrice = 0 Then
              'For VAT-Exclusive Rates need to enable the NetFixedAmount if we have it, otherwise use the default Amount (GrossFixedAmount)
              Dim vParamName As String = "Amount"
              If mvTA.LinePriceVATEx = True AndAlso FindControl(vEPL, "NetFixedAmount", False) IsNot Nothing AndAlso vEPL.PanelInfo.PanelItems("NetFixedAmount").Visible = True Then
                vEPL.SetValue("Amount", pValue)
                vParamName = "NetFixedAmount"
                Dim vVATAmount As Double = AppHelper.CalculateVATAmount(DoubleValue(pValue), mvTA.LineVATPercentage)
                pValue = FixTwoPlaces(DoubleValue(pValue) - vVATAmount).ToString
              End If
              SetValueRaiseChanged(vEPL, vParamName, pValue)
            End If
          Case CareServices.TraderPageType.tpPaymentPlanMaintenance
            mvTA.PPBalance = DoubleValue(pValue)
          Case CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance
            'If FindControl(vEPL, "Amount", False) IsNot Nothing AndAlso mvTA.LinePrice = 0 AndAlso vEPL.GetValue("Amount").Length = 0 Then
            '  vEPL.SetValue("Amount", pValue)
            'End If
            mvTA.PaymentPlanDetailsPricing.UpdateBalance(DoubleValue(pValue))
            If mvTA.LinePrice = 0 Then
              Dim vHasNetFixedAmount As Boolean
              If FindControl(vEPL, "NetFixedAmount", False) IsNot Nothing AndAlso vEPL.PanelInfo.PanelItems("NetFixedAmount").Visible = True Then vHasNetFixedAmount = True
              If (vHasNetFixedAmount = False AndAlso FindControl(vEPL, "Amount", False) IsNot Nothing AndAlso vEPL.GetValue("Amount").Length = 0) OrElse (vHasNetFixedAmount = True AndAlso vEPL.GetValue("NetFixedAmount").Length = 0) Then
                If vHasNetFixedAmount = True AndAlso mvTA.LinePriceVATEx = True Then
                  vEPL.SetValue("NetFixedAmount", FixTwoPlaces(DoubleValue(pValue) - mvTA.LineVATAmount).ToString)
                Else
                  If vEPL.GetValue("Amount").Length = 0 Then vEPL.SetValue("Amount", pValue) 'Amount should only be updated to balance if the value is for amount is not set 
                End If
                SetPPDPricingData(vEPL)
              End If
            End If
        End Select
      Case "CreditCardNumber"
        mvTA.CreditCardDetailsNumber = 0
        Dim vCCType As EditPanel.CreditCardValidationTypes = EditPanel.CreditCardValidationTypes.ccvtStandard
        If mvTA.TransactionPaymentMethod = "CAFC" Then vCCType = EditPanel.CreditCardValidationTypes.ccvtCAF
        Dim vList As New ParameterList(True)
        vList("CreditCardNumber") = pValue
        Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCards, IntegerValue(vEPL.GetValue("ContactNumber").ToString), vList, True)
        If vRow IsNot Nothing Then
          vEPL.SetValue("ExpiryDate", vRow("ExpiryDate").ToString)
          vEPL.SetValue("Issuer", vRow("Issuer").ToString)
          vEPL.SetValue("AccountName", vRow("AccountName").ToString)
          vEPL.SetValue("CreditCardType", vRow("CreditCardType").ToString)
          mvTA.CreditCardDetailsNumber = IntegerValue(vRow("CreditCardDetailsNumber").ToString)
        Else
          vEPL.SetValue("ExpiryDate", "")
          vEPL.SetValue("Issuer", "")
          vEPL.SetValue("AccountName", vEPL.FindTextLookupBox("ContactNumber").Description)
        End If
      Case "PaymentFrequency"
        Dim vPFRow As DataRow = vEPL.FindTextLookupBox("PaymentFrequency").GetDataRow
        If vTraderPage.PageType = CareServices.TraderPageType.tpPaymentPlanMaintenance _
          OrElse (vTraderPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetails AndAlso (mvTA.TransactionType = "MEMB" OrElse mvTA.TransactionType = "CMEM")) Then
          If FindControl(vEPL, "FirstAmount", False) IsNot Nothing AndAlso FindControl(vEPL, "FirstAmount").Visible Then
            'Can only amend the FirstAmount if Frequency is not 1
            If vPFRow.Item("Frequency").ToString = "1" Then
              vEPL.SetValue("FirstAmount", "", True)    'Clear value and disable
            Else
              vEPL.EnableControl("FirstAmount", True)   'Enable
            End If
          End If
        End If
        If mvTA.PaymentPlan IsNot Nothing Then mvTA.PaymentPlan.InitPaymentFrequency(vPFRow)
      Case "Product"
        Dim vRow As DataRow = vEPL.FindTextLookupBox(pParameterName).GetDataRow
        If vRow IsNot Nothing Then
          If FindControl(vEPL, "Rate", False) IsNot Nothing Then
            vEPL.EnableControl("Rate", True)
            If vEPL.FindTextLookupBox("Rate").GetDataRow IsNot Nothing AndAlso vEPL.FindTextLookupBox("Rate").GetDataRow.Table.Rows.Count = 2 Then
              vEPL.EnableControl("Rate", False)
            End If
          End If
          If CInt(vRow("SpecialPriceCount")) > 0 Then GetSpecialPrice(vEPL, pValue)
          mvTA.CheckQuantityBreaks = CInt(vRow("QuantityBreakCount")) > 0
          If FindControl(vEPL, "Quantity", False) IsNot Nothing Then
            Dim vQuantity As Integer = IntegerValue(vEPL.GetValue("Quantity"))
            Dim vNewQuantity As Integer = vQuantity
            If vRow("SalesQuantity").ToString.Length > 0 Then
              Dim vSalesQuantity As Integer = IntegerValue(vRow("SalesQuantity").ToString)
              If vQuantity <> vSalesQuantity Then vNewQuantity = vSalesQuantity
            End If
            If vRow("MinimumQuantity").ToString.Length > 0 Then
              Dim vMinQuantity As Integer = IntegerValue(vRow("MinimumQuantity").ToString)
              If vQuantity < vMinQuantity Then vNewQuantity = vMinQuantity
            End If
            If vRow("MaximumQuantity").ToString.Length > 0 Then
              Dim vMaxQuantity As Integer = IntegerValue(vRow("MaximumQuantity").ToString)
              If vQuantity > vMaxQuantity Then vNewQuantity = vMaxQuantity
            End If
            If vNewQuantity <> vQuantity Then
              SetValueRaiseChanged(vEPL, "Quantity", vNewQuantity.ToString)
              Dim vValid As Boolean = True
              EPL_ValidateItem(vEPL, "Quantity", vNewQuantity.ToString, vValid)
            End If
          End If
          If FindControl(vEPL, "DespatchMethod", False) IsNot Nothing Then
            'Ensure page contains DespatchMethod control as not all pages will
            vEPL.SetValue("DespatchMethod", vRow.Item("DespatchMethod").ToString)
          End If
          mvTA.StockSales = BooleanValue(vRow.Item("StockItem").ToString)
          If FindControl(vEPL, "Warehouse", False) IsNot Nothing Then
            'Ensure page contains Warehouse
            Dim vEnable As Boolean = False
            If mvTA.TransactionType = "SALE" Or mvCurrentPage.PageType = CareServices.TraderPageType.tpPurchaseOrderProducts Then
              If mvTA.StockSales Then
                'Get the Warehouses
                GetWarehouses(vEPL, pValue)
              Else
                vEPL.FindComboBox("Warehouse").DataSource = Nothing
              End If
              If mvTA.StockSales = True AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_stock_multiple_warehouses) = True Then
                'If more than one warehouse returned
                If vEPL.FindComboBox("Warehouse").Items.Count > 1 Then
                  vEnable = True
                Else
                  vEnable = False
                End If
              End If
              vEPL.SetValue("Warehouse", vRow.Item("Warehouse").ToString)
              If mvTA.StockSales Then
                'If we have changed the Product code then may need to add a StockMovement
                If mvTA.StockValuesChanged(pValue, vRow.Item("Warehouse").ToString, IntegerValue(vEPL.GetValue("Quantity")), False) Then
                  Dim vValid As Boolean = True
                  vValid = AddStockMovement(pValue, vRow.Item("Warehouse").ToString, IntegerValue(vEPL.GetValue("Quantity")))
                  If vValid Then 'BR13907
                    vEPL.SetErrorField("Quantity", "")
                  Else
                    vEPL.SetErrorField("Quantity", GetInformationMessage(InformationMessages.ImInsufficientStock), True)
                  End If
                End If
              End If
            End If
            vEPL.EnableControl("Warehouse", vEnable)
          End If
          If FindControl(vEPL, "BookingNumber", False) IsNot Nothing Then
            If mvTA.TransactionType = "SALE" AndAlso (mvTA.EventMultipleAnalysis = False OrElse mvTA.EventBookingDataSet.Tables.Count > 1) Then vEPL.SetValue("BookingNumber", "", True)
          End If
          If mvTA.TransactionType = "SALE" AndAlso mvTA.StockSales AndAlso FindControl(vEPL, "LastStockCount", False) IsNot Nothing Then
            If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails Then
              Dim vQuantity As Integer
              If FindControl(vEPL, "Quantity", False) IsNot Nothing Then
                vQuantity = IntegerValue(vEPL.GetValue("Quantity"))
              End If
              mvTotalStock = mvTotalStock - vQuantity
              vEPL.SetValue("LastStockCount", mvTotalStock.ToString)
            Else
              'in case last stock count control used elsewhere
              vEPL.SetValue("LastStockCount", vRow.Item("LastStockCount").ToString)
            End If
          ElseIf mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.TransactionType <> "SALE" Then
            'Specifically set focus to Rate otherwise it gets left on Product (next field is Warehouse and we have just disabled it) and user has to tab out a 2nd time
            vEPL.FindTextLookupBox("Rate").Focus()
          End If
          ProductNumberAllocation(vEPL, vRow)
        End If
        If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails Then
          'BR20678 - After changing product the focus is expected to go to Rate. If there is only one Rate for the Product, Rate will be protected and cannot accept focus.
          '          To cover single and muliple Rate behaviour, find the next control after Product that can accept focus a set focus to that control. 
          Dim vProductcontrol As Control = FindControl(vEPL, "Product", False)
          Dim vNextCanFocusControl As Control = vEPL.FindNextControlForFocus(vProductcontrol)
          If vNextCanFocusControl IsNot Nothing Then
            vNextCanFocusControl.Focus()
          End If
        End If
      Case "Rate"
        Dim vRow As DataRow = vEPL.FindTextLookupBox(pParameterName).GetDataRow
        If vRow IsNot Nothing Then
          If vRow("UseModifiers").ToString = "Y" Then
            Dim vModifierDate As Date
            If mvCurrentPage.PageType = CareNetServices.TraderPageType.tpChangeMembershipType AndAlso vEPL.GetValue("Joined").Length > 0 Then
              vModifierDate = CDate(vEPL.GetValue("Joined"))
            Else
              vModifierDate = CDate(If(mvTA.TransactionDate.Length > 0, mvTA.TransactionDate, AppValues.TodaysDate))
            End If
            mvTA.SetPaymentPlanDetailsPricing(DataHelper.GetModifierPriceData(vRow("Product").ToString, vRow("Rate").ToString, vModifierDate, IntegerValue(mvTA.PayerContactNumber), vRow("VATExclusive").ToString = "Y"))
            mvTA.LinePrice = mvTA.PaymentPlanDetailsPricing.Price(vRow("VATExclusive").ToString = "Y")
          Else
            If vRow("PriceChangeDate").ToString <> "" AndAlso DateValue(vRow("PriceChangeDate").ToString) <= Today Then
              mvTA.LinePrice = DoubleValue(vRow("FuturePrice").ToString)
            Else
              mvTA.LinePrice = DoubleValue(vRow("CurrentPrice").ToString)
            End If
          End If
          mvTA.LinePriceVATEx = vRow("VATExclusive").ToString = "Y"
          mvTA.FixedPrice = vRow.Table.Columns.Contains("FixedPrice") AndAlso vRow("FixedPrice").ToString = "Y"
        End If
        If vTraderPage.PageType <> CareServices.TraderPageType.tpGiveAsYouEarnEntry Then
          Dim vProductRow As DataRow = vEPL.FindTextLookupBox("Product").GetDataRow
          If vProductRow IsNot Nothing Then
            mvTA.ProductVATCategory = vProductRow("ProductVATCategory").ToString
            If (vTraderPage.PageType <> CareServices.TraderPageType.tpPostageAndPacking AndAlso
              vTraderPage.PageType <> CareServices.TraderPageType.tpCollectionPayments) Then
              Dim vAmountParam As String = ""
              If vTraderPage.PageType = CareServices.TraderPageType.tpPaymentPlanProducts OrElse vTraderPage.PageType = CareServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
                vAmountParam = "Balance"
              ElseIf (vTraderPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso mvTA.PayerHasDiscount) OrElse (vTraderPage.PageType = CareServices.TraderPageType.tpServiceBooking AndAlso mvTA.PayerHasDiscount AndAlso vEPL.FindPanelControl("Source").Visible) Then
                vAmountParam = "GrossAmount"
              Else
                vAmountParam = "Amount"
              End If
              vEPL.SetValue(vAmountParam, "")
              vEPL.SetValue("VatAmount", "", , , False)
            End If
            SetAmount(vEPL)
          End If
        End If
        If vTraderPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanProducts OrElse vTraderPage.PageType = CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance Then
          'Dont allow the user to change the balance for percentage discounts
          'Seperate rates will have to be setup if different discount rates are required
          If vRow("PriceIsPercentage").ToString.Length > 0 Then
            If vEPL.FindPanelControl("Balance", False) IsNot Nothing Then vEPL.EnableControl("Balance", vRow("PriceIsPercentage").ToString = "N")
          End If
        End If
        'Changes made to display currency code next to the amount text box
        If vTraderPage.PageType = CareNetServices.TraderPageType.tpProductDetails AndAlso vEPL.FindLabel("Amount", False) IsNot Nothing Then
          If vRow.Table.Columns.Contains("CurrencyCode") AndAlso Not mvTA.DefaultCurrencyCode = vRow("CurrencyCode").ToString Then
            Dim vAmountText As String = vEPL.FindLabel("Amount").Text
            Dim vEndString As String = "):"
            If vAmountText.Contains(vEndString) Then
              vAmountText = vAmountText.Substring(0, vAmountText.IndexOf(CChar("(")))
            End If
            If vAmountText.EndsWith(CChar(":")) Then vAmountText = vAmountText.TrimEnd(CChar(":"))
            vEPL.FindLabel("Amount").Text = vAmountText.TrimEnd + " (" + vRow("CurrencyCode").ToString + vEndString
          Else
            If mvOrgAmountText.Length = 0 Then
              mvOrgAmountText = vEPL.FindLabel("Amount").Text
            End If
            vEPL.FindLabel("Amount").Text = mvOrgAmountText
          End If
        End If
      Case "MemberNumber", "CovenantNumber", "PaymentPlanNumber", "AffiliatedMemberNumber", "BankersOrderNumber", "DirectDebitNumber", "CreditCardAuthorityNumber"
        ProcessPPNumbersChanged(vTraderPage, vEPL, pParameterName, pValue)
        If vTraderPage.PageType.Equals(CareNetServices.TraderPageType.tpTransactionDetails) Then ShowMemberInfo(pParameterName, pValue)
      Case "Source"
        If vEPL.FindTextLookupBox("Source").GetDataRow IsNot Nothing Then
          mvTA.CheckIncentives = vEPL.FindTextLookupBox("Source").GetDataRowItem("IncentiveScheme").Length > 0
          If vTraderPage.PageType <> CareServices.TraderPageType.tpServiceBooking Then
            mvTA.TransactionSource = pValue
            If pValue.Length > 0 Then
              Dim vDistributionTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("DistributionCode", False)
              If vDistributionTextLookupBox IsNot Nothing AndAlso String.IsNullOrEmpty(vDistributionTextLookupBox.Text) Then
                ' BR19597 - Do not default the Distribution Code if it already has a value. Use control value not mvTA property. The property may not be initialised but the control will, if present.
                mvTA.TransactionDistributionCode = vEPL.FindTextLookupBox("Source").GetDataRow.Item("DistributionCode").ToString
                Select Case vTraderPage.PageType
                  Case CareServices.TraderPageType.tpMembership, CareServices.TraderPageType.tpChangeMembershipType, CareServices.TraderPageType.tpGiveAsYouEarnEntry _
                    , CareServices.TraderPageType.tpProductDetails, CareNetServices.TraderPageType.tpPaymentPlanProducts, CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance
                    vEPL.SetValue("DistributionCode", mvTA.TransactionDistributionCode, , , False)
                End Select
              End If
            End If
          End If
          Dim vOldDiscountPercentage As Double = mvTA.DiscountPercentage
          Dim vRow As DataRow = vEPL.FindTextLookupBox("Source").GetDataRow
          If vRow.Table.Columns.Contains("DiscountPercentage") Then mvTA.SourceDiscountPercentage = DoubleValue(vRow("DiscountPercentage").ToString)
          If (mvCurrentPage.PageType <> CareNetServices.TraderPageType.tpPaymentPlanProducts AndAlso mvCurrentPage.PageType <> CareNetServices.TraderPageType.tpPaymentPlanDetailsMaintenance) _
          AndAlso vOldDiscountPercentage <> mvTA.DiscountPercentage Then
            'Only do this if the discount percentage has changed and we are not on Payment Plan details pages (they don't have gross amount fields etc.)
            If FindControl(vEPL, "Rate", False) IsNot Nothing AndAlso vEPL.FindTextLookupBox("Rate").IsValid Then
              Dim vDataTable As DataTable = DirectCast(vEPL.FindTextLookupBox("Rate").ComboBox.DataSource, DataTable)
              If vDataTable IsNot Nothing Then
                SetValueRaiseChanged(vEPL, "Rate", vEPL.FindTextLookupBox("Rate").Text)
              End If
              Dim vGrossAmt As Control = vEPL.FindPanelControl("GrossAmount", False)
              If vGrossAmt IsNot Nothing AndAlso vEPL.FindPanelControl("Amount", False) IsNot Nothing Then
                If (mvTA.PayerHasDiscount AndAlso vGrossAmt.Visible = False) OrElse
                 (mvTA.PayerHasDiscount = False AndAlso vGrossAmt.Visible) Then
                  'Field properties need changing
                  vEPL.SetControlVisible("GrossAmount", mvTA.PayerHasDiscount)
                  vEPL.SetControlVisible("Discount", mvTA.PayerHasDiscount)
                  vEPL.EnableControl("Discount", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_discount_amend))
                End If
              End If
            End If
          End If
        End If
        Select Case vTraderPage.PageType
          Case CareNetServices.TraderPageType.tpTransactionDetails
            If mvTA.AnalysisDataSet.Tables.Count > 0 AndAlso mvTA.AnalysisDataSet.Tables.Contains("DataRow") Then
              Me.mvTA.AnalysisDataSet.Tables("DataRow").Rows(mvCurrentRow).Item("Source") = Me.mvTA.TransactionSource
            End If
          Case CareNetServices.TraderPageType.tpProductDetails
            Me.mvTA.TransactionSource = vEPL.FindTextLookupBox("Source").Text
        End Select
      Case "NewToken"
        If FindControl(vEPL, "TokenDescription", False) IsNot Nothing Then
          If BooleanValue(pValue) Then
            vEPL.SetControlVisible("TokenDescription", True)
            mvTA.TokenDescription = vEPL.GetValue("TokenDescription")
          Else
            vEPL.SetControlVisible("TokenDescription", False)
            mvTA.TokenDescription = String.Empty
          End If
          If FindControl(vEPL, "TokenDesc", False) IsNot Nothing Then
            DirectCast(FindControl(vEPL, "TokenDesc", False), ListBox).ClearSelected()
            mvTA.TokenDescription = String.Empty
          End If
        End If
      Case "TokenDesc"
        If FindControl(vEPL, "NewToken", False) IsNot Nothing AndAlso vEPL.FindCheckBox("NewToken").Checked Then
          vEPL.FindCheckBox("NewToken").Checked = False
          mvTA.TokenDescription = vEPL.GetValue("TokenDescription")
        End If

    End Select
    If pValue.Length > 0 Then vEPL.PanelInfo.PanelItems(pParameterName).LastUsedValue = pValue 'BR12225: Added support for Ctrl+L key
  End Sub

  Private Sub GetMemberBalanceAndRenewal(ByVal pEpl As EditPanel)
    GetMemberBalanceAndRenewal(pEpl, False)
  End Sub
  Private Sub GetMemberBalanceAndRenewal(ByVal pEpl As EditPanel, ByVal pStartDateChanged As Boolean)
    Dim vList As New ParameterList(True, True)
    If mvTraderPages(CareServices.TraderPageType.tpMembership.ToString).EditPanel.AddValuesToList(vList) Then
      pEpl.AddValuesToList(vList)
      mvTA.GetApplicationValues(vList)
      If mvTA.IncentiveDataSet IsNot Nothing Then AddIncentivesLines(vList, False, "", 0)
      Select Case mvTA.PPPaymentType
        Case "CCCA", "DIRD", "STDO"
          vList("PPPaymentMethod") = mvTA.PPPaymentMethod
          vList("BankAccount") = mvTA.CABankAccount   'Use a default value
          If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
            vList("ClaimDay") = "1"   'Use a default value
          End If
      End Select
      'check that the required parameters exist and have values before calling webservice
      If vList.ValueIfSet("PaymentFrequency").Length > 0 AndAlso vList.ValueIfSet("Source").Length > 0 AndAlso vList.ValueIfSet("OrderDate").Length > 0 Then
        Dim vReturnList As New ParameterList
        If pStartDateChanged Then
          If Not vList.ContainsKey("MembershipRate") Then vList.Add("MembershipRate", vList.ValueIfSet("Rate"))
          If Not vList.ContainsKey("PaymentPlanDetails") Then vList.Add("PaymentPlanDetails", "Y")
          Dim vDataSet As DataSet = DataHelper.GetDataSetFromResult(DataHelper.GetMemberBalance(vList))
          If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("PPDLine") Then
            If mvTA.PPDDataSet.Tables.Contains("DataRow") Then mvTA.PPDDataSet.Tables.Remove("DataRow")
            SetPPDLines(vDataSet)
          End If

          If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("Result") Then
            pEpl.SetValue("Balance", vDataSet.Tables("Result").Rows(0).Item("MemberBalance").ToString)
          End If
          vReturnList.FillFromXMLString(DataHelper.GetMemberRenewalBalance(vList))
          pEpl.SetValue("RenewalAmount", vReturnList.Item("MemberRenewalAmount").ToString)
        Else
          vReturnList = DataHelper.GetMemberBalanceAndRenewal(vList)
          pEpl.SetValue("Balance", vReturnList.Item("MemberBalance").ToString)
          pEpl.SetValue("RenewalAmount", vReturnList.Item("MemberRenewalAmount").ToString)
        End If
      End If
    End If
  End Sub

  Private Sub ProductNumberAllocation(ByVal pEPL As EditPanel, ByVal pRow As DataRow)
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails AndAlso FindControl(pEPL, "ProductNumber", False) IsNot Nothing Then
      Dim vList As New ParameterList(True)
      If pEPL.PanelInfo.PanelItems("Product").LastUsedValue IsNot Nothing AndAlso pEPL.GetValue("Product") <> pEPL.PanelInfo.PanelItems("Product").LastUsedValue Then
        If pEPL.GetValue("ProductNumber").Length > 0 Then
          vList("Product") = pEPL.PanelInfo.PanelItems("Product").LastUsedValue
          vList("ProductNumber") = pEPL.GetValue("ProductNumber")
          vList("Reallocate") = "Y"
          DataHelper.ProcessProductNumberAllocation(vList)
          pEPL.SetValue("ProductNumber", "")
        End If
      End If
      If pRow.Item("UsesProductNumbers").ToString = "Y" Then
        If pEPL.GetValue("ProductNumber").Length = 0 Then
          vList = New ParameterList(True)
          vList("Product") = pEPL.GetValue("Product")
          Dim vFind As Boolean
          Dim vProductNumber As Integer
          Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_prompt_allocate_prod_number)
            Case "Always"
              If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctProductNumbers, vList) > 0 Then
                'Existing Product Numbers available for re-allocation
                Dim vMsgResult As DialogResult = ShowQuestion(QuestionMessages.QmConfirmProducReAllocation, MessageBoxButtons.YesNo)
                Select Case vMsgResult
                  Case System.Windows.Forms.DialogResult.Yes
                    While vFind = False
                      If vList.Contains("ProductNumber") Then vList.Remove("ProductNumber")
                      Dim vReturnList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptReAllocateProductNumber, vList)
                      If vReturnList.Count > 0 Then
                        vProductNumber = vReturnList.IntegerValue("ProductNumber")
                        vList.IntegerValue("ProductNumber") = vProductNumber
                        If DataHelper.ProcessProductNumberAllocation(vList).Contains("ProductNumber") Then vFind = True Else ShowErrorMessage(InformationMessages.ImInvalidOrUsedProductNumber)
                      Else
                        'User clicked Cancel
                        vFind = True
                      End If
                    End While
                  Case System.Windows.Forms.DialogResult.No
                    'Get next Product Number from the Product
                    vFind = True
                End Select
              Else
                'just get next number
                vFind = True
              End If
            Case "No Prompt"
              vFind = True
            Case Else
              '
          End Select

          If vFind Then
            vList("GetNextNumber") = "Y"
            If vProductNumber = 0 Then vProductNumber = DataHelper.ProcessProductNumberAllocation(vList).IntegerValue("ProductNumber")
            pEPL.SetValue("ProductNumber", vProductNumber.ToString)
          End If
        End If
        'Disable Quantity field
        If pRow.Item("SalesQuantity").ToString > "Y" AndAlso (mvTA.FinancialAdjustment = BatchInfo.AdjustmentTypes.GIKConfirmation OrElse mvTA.TransactionType = "CSRT") Then pEPL.FindTextBox("Quantity").Enabled = False
      End If
    End If
  End Sub
  Private Sub PopulateActivityDataSheet(ByVal pValue As String, ByVal vEPL As EditPanel, ByVal pParameterName As String)
    Dim vADS As ActivityDataSheet = TryCast(FindControl(vEPL, "Activity", False), ActivityDataSheet)
    If vADS IsNot Nothing Then
      Dim vContactInfo As ContactInfo = vEPL.FindTextLookupBox("ContactNumber").ContactInfo
      Dim vSource As String = ""
      If pParameterName = "ContactNumber" Then
        vSource = vEPL.GetValue("Source")
      ElseIf pParameterName = "Source" Then
        vSource = pValue
      End If
      If vContactInfo IsNot Nothing AndAlso vSource.Length > 0 AndAlso vContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse Then
        Dim vList As New ParameterList(True)
        vList("UsageCode") = "R"
        vList("Source") = vSource
        vList(vContactInfo.ContactGroupParameterName) = vContactInfo.ContactGroup
        If mvTA.DefaultActivityGroup.Length > 0 Then vList("ActivityGroup") = mvTA.DefaultActivityGroup
        Dim vTable As DataTable
        vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtActivityDataSheet, vList)
        If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
          vADS.Init(vContactInfo, mvTA.DefaultActivityGroup, vTable, vSource, ActivityDataSheet.ActivityDataSheetTypes.adstTrader)
        Else
          ShowInformationMessage(InformationMessages.ImNoActivities)
        End If
      End If
    End If
  End Sub
  Private Sub PopulateSuppressionDataSheet(ByVal pValue As String, ByVal vEPL As EditPanel)
    Dim vSDS As SuppressionDataSheet = TryCast(FindControl(vEPL, "MailingSuppression", False), SuppressionDataSheet)
    If vSDS IsNot Nothing Then
      Dim vContactInfo As ContactInfo = vEPL.FindTextLookupBox("ContactNumber").ContactInfo
      If vContactInfo IsNot Nothing Then vSDS.Init(vContactInfo, mvTA.DefaultSuppression, SuppressionDataSheet.SuppressionDataSheetTypes.sdstTrader)
    End If
  End Sub
  Private Sub ValidateEffectiveDate(ByVal pEPL As EditPanel, ByVal pValue As String, ByRef pValid As Boolean)
    If pValue.Length > 0 Then
      Dim pMessage As String = ""
      If CDate(pValue) < Today Then
        pMessage = InformationMessages.ImEffectiveDateCurrent
      ElseIf CDate(pValue) > Today.AddYears(1) Then
        pMessage = InformationMessages.ImEffectiveDateYear
      Else
        If DataHelper.GetTableFromDataSet(mvTA.PPDDataSet) IsNot Nothing Then
          For Each vRow As DataRow In DataHelper.GetTableFromDataSet(mvTA.PPDDataSet).Rows
            If vRow IsNot Nothing AndAlso vRow("EffectiveDate").ToString.Length > 0 AndAlso CDate(pValue) < CDate(vRow("EffectiveDate").ToString) Then
              pMessage = InformationMessages.ImEffectiveDateExisting
              Exit For
            End If
          Next
        End If
      End If
      If pMessage.Length > 0 Then
        pEPL.SetErrorField("EffectiveDate", pMessage, True)
        pValid = False
      Else
        Dim vPrice As Double = mvTA.LinePrice
        If vPrice = 0 Then vPrice = DoubleValue(pEPL.GetOptionalValue("Amount"))
        SetPPDMaintenanceBalance(pEPL, vPrice.ToString)
      End If
    End If
  End Sub

  Private Sub ValidatePurchaseOrderNumber(ByVal pEPL As EditPanel, ByVal pPageType As CareServices.TraderPageType, ByVal pValue As String)
    Select Case pPageType
      Case CareServices.TraderPageType.tpPurchaseInvoiceDetails
        If pValue.Length > 0 Then
          Dim vList As New ParameterList(True)
          vList("PurchaseOrderNumber") = pEPL.GetValue("PurchaseOrderNumber")
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtPurchaseOrderInformation, IntegerValue(pEPL.GetValue("PurchaseOrderNumber"))))
          If vTable IsNot Nothing Then
            Dim vContactInfo As ContactInfo
            For Each vRow As DataRow In vTable.Rows
              vContactInfo = New ContactInfo(IntegerValue(vRow("ContactNumber")))
              mvTA.PurchaseInvoiceNumber = IntegerValue(vRow("PurchaseInvoiceNumber"))
              If BooleanValue(vRow("PaymentSchedule").ToString) OrElse BooleanValue(vRow("AdHocPayments").ToString) OrElse
                BooleanValue(vRow("RegularPayments").ToString) Then
                pEPL.SetErrorField("PurchaseOrderNumber", InformationMessages.ImCannotCreatePurchaseInvoice, True)
              ElseIf mvTA.PurchaseInvoiceNumber > 0 Then
                If vRow("ChequeReferenceNumber").ToString.Length > 0 Then
                  pEPL.SetErrorField("PurchaseOrderNumber", InformationMessages.ImChequeAlreadyProcessed, True)
                ElseIf vRow("BacsProcessed").ToString = "P" Then
                  'BR13667: BACS Processed so prevent edit of Purchase Invoice
                  pEPL.SetErrorField("PurchaseOrderNumber", InformationMessages.ImBacsAlreadyProcessed, True)
                Else
                  If ShowQuestion(QuestionMessages.QmPurchaseInvoiceProcessed, MessageBoxButtons.YesNo, vContactInfo.ContactName) = System.Windows.Forms.DialogResult.Yes Then
                    vRow = DataHelper.GetRowFromDataSet(DataHelper.GetPurchaseInvoiceData(CareServices.XMLPurchaseInvoiceDataSelectionTypes.xodtPurchaseInvoiceInformation, IntegerValue(mvTA.PurchaseInvoiceNumber)))
                    pEPL.Populate(vRow)
                    If pEPL.PanelInfo.PanelItems.Exists("CurrencyCode") Then pEPL.EnableControl("CurrencyCode", False)
                    'ElseIf ShowQuestion(QuestionMessages.QmConfirmDeletePreviousEntry, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                    'vList.IntegerValue("PurchaseInvoiceNumber") = mvTA.PurchaseInvoiceNumber
                    'DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctPurchaseInvoice, vList)
                  End If
                End If
              ElseIf vRow("CancellationReason").ToString.Length > 0 Then
                pEPL.SetErrorField("PurchaseOrderNumber", GetInformationMessage(InformationMessages.ImChequeAlreadyCancelledOnBy, vRow("CancelledOn").ToString, vContactInfo.ContactName), True)
              Else
                pEPL.Populate(vRow)
                If pEPL.PanelInfo.PanelItems.Exists("CurrencyCode") Then pEPL.EnableControl("CurrencyCode", False)
              End If
              Exit For
            Next
          Else
            pEPL.SetErrorField("PurchaseOrderNumber", "Invalid Number", True)
          End If
        Else
          If pEPL.PanelInfo.PanelItems.Exists("CurrencyCode") Then pEPL.EnableControl("CurrencyCode", True)
        End If
      Case CareServices.TraderPageType.tpPurchaseOrderCancellation
        If IntegerValue(pEPL.GetValue("PurchaseOrderNumber2")) > 0 AndAlso IntegerValue(pEPL.GetValue("PurchaseOrderNumber")) > IntegerValue(pEPL.GetValue("PurchaseOrderNumber2")) Then
          pEPL.SetErrorField("PurchaseOrderNumber", InformationMessages.ImInvalidRange, True)
        Else
          pEPL.SetErrorField("PurchaseOrderNumber", "")
        End If
    End Select

  End Sub

  Private Sub SetPOILineItemControls(ByVal vEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Dim vFound As Boolean
    With vEPL
      If FindControl(vEPL, pParameterName, False) IsNot Nothing Then
        Dim vCombo As ComboBox = .FindComboBox(pParameterName)
        If vCombo.Visible Then
          If .PanelInfo.PanelItems("LineItem").Visible Then
            If vCombo.Items.Count > 0 Then vFound = True
          Else
            vFound = True
          End If
        End If
        .SetControlVisible(pParameterName, vFound)
        If vFound Then
          .FindComboBox(pParameterName).DropDownStyle = ComboBoxStyle.DropDown
          .SetValue(pParameterName, pValue)
        End If
      End If
      If .PanelInfo.PanelItems.Exists(pParameterName) Then .PanelInfo.PanelItems(pParameterName).Mandatory = vFound
      vEPL.SetControlVisible("LineItem", Not vFound)
    End With
  End Sub

  Private Sub PreValidateItems()
    'Cleare the error flag on the fields which could have become Valid as the data could have been chaged outside the trader page
    If mvCurrentPage.PageType = CareServices.TraderPageType.tpProductDetails Then
      Dim vInMemoriam As Boolean
      If FindControl(mvCurrentPage.EditPanel, "LineType_G", False) IsNot Nothing Then
        vInMemoriam = (mvCurrentPage.EditPanel.GetValue("LineType_G") = "G")
      Else
        vInMemoriam = BooleanValue(mvCurrentPage.EditPanel.GetValue("LineTypeG"))
      End If
      If vInMemoriam = True Then ValidateDeceasedContact(mvCurrentPage.EditPanel, "DeceasedContactNumber", "")
    End If
  End Sub

  ''' <summary>Get the RunType for Invoice printing</summary>
  ''' <param name="pEPL">The EditPanel the controls are on</param>
  ''' <param name="pValue">The value currently set</param>
  ''' <returns>True if the control is visible, otherwise False</returns>
  Private Function GetInvoicePrintRunType(ByVal pEPL As EditPanel, ByRef pValue As String) As Boolean
    pValue = ""
    Dim vVisible As Boolean = False
    Dim vControl As Control = Nothing
    vControl = FindControl(pEPL, "RunType_N", False)
    If vControl IsNot Nothing Then
      vVisible = vControl.Visible
      pValue = pEPL.GetValue("RunType_N")
      pEPL.SetErrorField("RunType_N", "") 'Clear any error
    End If
    If pValue.Length = 0 Then
      vControl = FindControl(pEPL, "RunType_R", False)
      If vControl IsNot Nothing Then
        If vVisible = False Then vVisible = vControl.Visible
        pValue = pEPL.GetValue("RunType_R")
      End If
    End If
    If pValue.Length = 0 Then
      vControl = FindControl(pEPL, "RunType_A", False)
      If vControl IsNot Nothing Then
        If vVisible = False Then vVisible = vControl.Visible
        pValue = pEPL.GetValue("RunType_A")
      End If
    End If
    Return vVisible
  End Function

  Private Sub SetPPDPricingData(ByVal pEPL As EditPanel)
    Dim vUnitPrice As Double = mvTA.LinePrice
    If mvTA.LinePrice = 0 Then vUnitPrice = DoubleValue(pEPL.GetValue("Amount"))
    Dim vPPDPrice As Double = DoubleValue(pEPL.GetValue("Balance"))
    Dim vVATAmount As Double = AppHelper.CalculateVATAmount(vPPDPrice, mvTA.LineVATPercentage)   'Explicitly calculate VAT here because if LinePrice is zero the LineVATAmount is zero
    mvTA.PaymentPlanDetailsPricing.SetPricing(vUnitPrice, vPPDPrice, DoubleValue(pEPL.GetValue("Quantity")), vVATAmount, False, mvTA.LineVATRate, mvTA.LineVATPercentage)
  End Sub

  Private Sub ProcessExamSelections(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Dim vReselectUnits As Boolean = False

    Dim vCentreCode As String = pEPL.GetValue("ExamCentreCode")
    Dim vCourseCode As String = pEPL.GetValue("ExamUnitCode")
    Dim vSessionCode As String = pEPL.GetValue("ExamSessionCode")
    Dim vStudyMode As String = pEPL.GetValue("StudyMode")
    Dim vSessionID As Integer = 0

    Dim vCentreTLB As TextLookupBox = pEPL.FindTextLookupBox("ExamCentreCode")
    Dim vCourseTLB As TextLookupBox = pEPL.FindTextLookupBox("ExamUnitCode")
    Dim vSessionTLB As TextLookupBox = pEPL.FindTextLookupBox("ExamSessionCode")
    Dim vSessonIdString As String = String.Empty

    Select Case pParameterName
      Case "ContactNumber"
        mvProcessingExams = False
        vReselectUnits = True
        If FindControl(pEPL, "ContactReference", False) IsNot Nothing Then
          pEPL.SetValue("ContactReference", "")   'Clear any existing value
          Dim vContactInfo As ContactInfo = pEPL.FindTextLookupBox("ContactNumber").ContactInfo
          If vContactInfo.ContactReference.Length > 0 Then pEPL.SetValue("ContactReference", vContactInfo.ContactReference) 'If we have a value, then set it
          vSessionCode = pEPL.GetValue("ExamSessionCode")
          If vSessionCode.Length > 0 Then vSessionID = vSessionTLB.GetDataRowInteger("ExamSessionId")
        End If

      Case "ExamCentreCode", "ExamSessionCode", "ExamUnitCode"
        If pParameterName = "ExamSessionCode" AndAlso FindControl(pEPL, "CourseStartDate", False) IsNot Nothing Then
          'If we have a Session selected, clear & disable CourseStartDate
          vSessonIdString = vSessionTLB.GetDataRowItem("ExamSessionId")
          vSessionID = vSessionTLB.GetDataRowInteger("ExamSessionId")
          If vSessionID > 0 Then pEPL.SetValue("CourseStartDate", "")
          pEPL.EnableControl("CourseStartDate", (vSessionID = 0))
        End If

        If mvProcessingExams = False Then     'Do this to prevent processing ValueChanged events multiple times as the controls below are re-populated
          Dim vList As New ParameterList(True, True)
          vList("Trader") = "Y"
          vSessionID = 0
          If vCentreCode.Length > 0 Then vList.IntegerValue("ExamCentreId") = vCentreTLB.GetDataRowInteger("ExamCentreId")
          If vCourseCode.Length > 0 Then vList.IntegerValue("ExamUnitId") = vCourseTLB.GetDataRowInteger("ExamUnitId")
          If vSessionCode.Length > 0 Then
            vList("ExamSessionCode") = vSessionCode
            vSessonIdString = vSessionTLB.GetDataRowItem("ExamSessionId")
            vSessionID = vSessionTLB.GetDataRowInteger("ExamSessionId")
            If vSessionID > 0 Then vList.IntegerValue("ExamSessionId") = vSessionID
          End If
          vList("NonSessionBased") = "Y" 'Used to add 'Non-Session Based' to the Sessions

          mvProcessingExams = True
          'Re-select all combo data
          vSessionTLB.FillComboWithRestriction(vSessionID.ToString, "", False, vList)
          vList.Remove("NonSessionBased")
          vCourseTLB.FillComboWithRestriction(vSessionID.ToString, "", False, vList)
          vCentreTLB.FillComboWithRestriction(vSessionID.ToString, "", False, vList)
          'Re-set field values
          If vSessionCode.Length > 0 Then pEPL.SetValue("ExamSessionCode", vSessionCode)
          If vCourseCode.Length > 0 Then pEPL.SetValue("ExamUnitCode", vCourseCode)
          If vCentreCode.Length > 0 Then pEPL.SetValue("ExamCentreCode", vCentreCode)
          mvProcessingExams = False
          vReselectUnits = True
        End If
      Case "StudyMode"
        vReselectUnits = True
    End Select

    If vReselectUnits Then
      pEPL.SetValue("Amount", "")       'Clear any existing amount if reselecting
      pEPL.SetErrorField("ExamCentreCode", "")
      If vSessionID > 0 Then
        pEPL.PanelInfo.PanelItems("ExamCentreCode").Mandatory = True
      Else
        'If there is no Session selected, this is only mandatory if customised to be mandatory
        pEPL.PanelInfo.PanelItems("ExamCentreCode").Mandatory = pEPL.PanelInfo.PanelItems("ExamCentreCode").OriginalMandatory
      End If
      Dim vUnitId As Integer = vCourseTLB.GetDataRowInteger("ExamUnitId")
      Dim vUnitLinkId As Integer = vCourseTLB.GetDataRowInteger("ExamUnitLinkId")
      Dim vContactNumber As Integer = IntegerValue(pEPL.GetValue("ContactNumber"))
      Dim vCentreId As Integer = vCentreTLB.GetDataRowInteger("ExamCentreId")
      Dim vControl As Control = FindControl(pEPL, "ExamUnitId", False)
      If vSessionCode.Length > 0 Then vSessionID = vSessionTLB.GetDataRowInteger("ExamSessionId")
      'NFPNGEX 43 Modified to always select even if this results in clearing the treeview
      DirectCast(vControl, ExamSelector).InitForTrader(ExamSelector.SelectionType.BookingCourses, vSessonIdString, vUnitLinkId, vContactNumber, vStudyMode, vCentreId)
    End If
    If vCourseCode.Length > 0 Then CalculateExamBookingPrice(pEPL)

  End Sub
  Private Sub txt_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
    Dim vContactInfo As ContactInfo
    If DragDataValid(e, vContactInfo) Then
      Dim vTextBox As TextBox = DirectCast(sender, TextBox)
      vTextBox.Text = vContactInfo.ContactNumber.ToString
      vTextBox.Focus()
    End If

  End Sub
  Private Sub txt_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
    If DragDataValid(e) Then e.Effect = DragDropEffects.Copy
  End Sub

  Private Function DragDataValid(ByVal e As System.Windows.Forms.DragEventArgs, Optional ByRef pContactInfo As ContactInfo = Nothing) As Boolean
    Dim vTypeName As String = GetType(ContactInfo).FullName
    pContactInfo = CType(e.Data.GetData(vTypeName), ContactInfo)
    If pContactInfo IsNot Nothing Then DragDataValid = True
  End Function

#End Region

End Class
