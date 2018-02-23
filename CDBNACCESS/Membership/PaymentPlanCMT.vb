Namespace Access

  Partial Public Class PaymentPlan
    'This class only handles CMT

    Private Structure CMTExcessPaymentDetail
      Friend OriginalProductCode As String
      Friend OriginalRateCode As String
      Friend ExcessPaymentType As CmtExcessPayment.CMTExcessPaymentTypes
      Friend ExcessPaymentAmount As Double
      Friend AdjustmentProductCode As String
      Friend AdjustmentRateCode As String

      Friend Sub New(ByVal pOriginalProductCode As String, ByVal pOriginalRateCode As String, ByVal pExcessPaymentType As CmtExcessPayment.CMTExcessPaymentTypes, ByVal pExcessPaymentAmount As Double, ByVal pAdjustmentProductCode As String, ByVal pAdjustmentRateCode As String)
        OriginalProductCode = pOriginalProductCode
        OriginalRateCode = pOriginalRateCode
        ExcessPaymentType = pExcessPaymentType
        ExcessPaymentAmount = pExcessPaymentAmount
        AdjustmentProductCode = pAdjustmentProductCode
        AdjustmentRateCode = pAdjustmentRateCode
      End Sub
    End Structure

    Private Class CMTExcessPaymentProductDetail
      Private mvProductCode As String
      Private mvRateCode As String

      Friend Sub New(ByVal pProductCode As String, ByVal pRateCode As String, ByVal pPrice As Double, ByVal pBalance As Double, ByVal pAmountPaid As Double)
        mvProductCode = pProductCode
        mvRateCode = pRateCode
        Price = pPrice
        Balance = pBalance
        AmountPaid = pAmountPaid
      End Sub

      Friend Property Price As Double
      Friend Property Balance As Double
      Friend Property AmountPaid As Double

      Friend ReadOnly Property ProductCode As String
        Get
          Return mvProductCode
        End Get
      End Property

      Friend ReadOnly Property RateCode As String
        Get
          Return mvRateCode
        End Get
      End Property
    End Class

    Private mvCMTOldPricing As CMTPricing
    Private mvCMTNewPricing As CMTPricing
    Private mvCMTNewTermMonths As Integer

    Private mvCMTInAdvanceDetails As CollectionList(Of CMTExcessPaymentDetail)
    Private mvCMTRefundDetails As CollectionList(Of CMTExcessPaymentDetail)
    Private mvCMTProductDetails As CollectionList(Of CMTExcessPaymentProductDetail)            'Holds all the chargeable detail line product/rates etc.for handling excess payments

    Friend Sub SetupCMT(ByVal pParams As CDBParameters, ByVal pTransaction As TraderTransaction, ByVal pCMTDate As Date, ByVal pJoinedDate As Date, ByVal pGotCMTData As Boolean)
      mvProcessCMT = True
      Dim vPFCode As String = PaymentFrequencyCode
      Dim vPayFrequency As PaymentFrequency = mvEnv.GetPaymentFrequency(vPFCode)

      mvCMTOldPricing = New CMTPricing(mvEnv, Payer, pCMTDate, pJoinedDate, MembershipType, vPayFrequency, CMTPricing.CMTOldOrNewMemberType.OldMembershipType)

      If pParams.HasValue("PaymentFrequency") Then vPFCode = pParams("PaymentFrequency").Value
      vPayFrequency = mvEnv.GetPaymentFrequency(vPFCode)
      Dim vNewMemberType As MembershipType = mvEnv.MembershipType(pParams("MembershipType").Value)
      mvCMTNewPricing = New CMTPricing(mvEnv, Payer, pCMTDate, pJoinedDate, vNewMemberType, vPayFrequency, CMTPricing.CMTOldOrNewMemberType.NewMembershipType)

      'Get the old detail lines
      mvCMTOldPricing.SetupMembershipDetails(Details, pGotCMTData)
      'Get the new detail lines
      If pParams.Exists("Source") = False AndAlso pParams.Exists("MemberSource") = True Then pParams.Add("Source", pParams("MemberSource").Value)
      If pGotCMTData = False AndAlso mvSecondCMT = False Then
        Dim vCMTRenewalAmount As Double
        pTransaction.TraderPPDLines.Clear()
        Dim vNewPrice As Double = pTransaction.PaymentPlan.GetMemberBalance(pParams, pTransaction, vPFCode, vCMTRenewalAmount)
      End If
      mvCMTNewPricing.SetupMembershipDetails(pTransaction.TraderPPDLines, pGotCMTData)

      'Do the prorating
      Dim vFullTermMonths As Integer
      Dim vChargeTermMonths As Integer
      Dim vOldTypeMonths As Integer
      Dim vNewTypeMonths As Integer
      Dim vUseProrateMonths As Boolean
      CalculateCMTTerm(pCMTDate, vPayFrequency, pJoinedDate, vChargeTermMonths, vFullTermMonths, vOldTypeMonths, vNewTypeMonths, vUseProrateMonths)
      mvCMTNewTermMonths = vNewTypeMonths
      If pGotCMTData = False Then
        mvCMTOldPricing.ProrateCosts(CMTProportionBalance, vOldTypeMonths, vFullTermMonths, vChargeTermMonths, vUseProrateMonths, CanUseAdvancedCMT(vNewMemberType))
        mvCMTNewPricing.ProrateCosts(CMTProportionBalance, vNewTypeMonths, vFullTermMonths, vChargeTermMonths, False, CanUseAdvancedCMT(vNewMemberType))
      End If
    End Sub

    ''' <summary>Calculate the number of months for the old and new membership types.</summary>
    ''' <param name="pCMTDate">Date the prorating will be effective.</param>
    ''' <param name="pNewPayFrequency">The new PaymentFrequency. Only used when prorating is by frequency amounts.</param>
    ''' <param name="pJoinedDate">The members joined date.</param>
    ''' <param name="pChargeTermMonths">The total number of months to be charged for.</param>
    ''' <param name="pFullTermMonths">The total number of months in the current renewal period.</param>
    ''' <param name="pOldTypeMonths">The number of months applicable for the old membership type.</param>
    ''' <param name="pNewTypeMonths">The number of months applicable or the new membership type.</param>
    ''' <param name="pUseProrateMonths">Boolean flag indicating whether to use <paramref name="pChargeTermMonths">pChargeTermMonths</paramref>.</param>
    ''' <remarks></remarks>
    Friend Sub CalculateCMTTerm(ByVal pCMTDate As Date, ByVal pNewPayFrequency As PaymentFrequency, ByRef pJoinedDate As Date, ByRef pChargeTermMonths As Integer, ByRef pFullTermMonths As Integer, ByRef pOldTypeMonths As Integer, ByRef pNewTypeMonths As Integer, ByRef pUseProrateMonths As Boolean)
      Dim vRenewalDate As Date = CDate(RenewalPeriodEnd)
      If IsMultipleCMT() = True AndAlso mvFirstCMT = False AndAlso mvSecondCMT = False AndAlso RenewalPending = True Then
        'Multiple CMT and as neither mvFirstCMT or mvSecondCMT are set we must be in the first so re-set the dates
        vRenewalDate = CDate(RenewalDate)
      End If
      Dim vStartDate As Date = CDate(CalculateRenewalDate(vRenewalDate.ToString(CAREDateFormat), False))

      pFullTermMonths = Term
      If pFullTermMonths < 0 Then pFullTermMonths = 1 ''I'-type incentive
      If pFullTermMonths > 0 Then pFullTermMonths = (pFullTermMonths * 12)
      pChargeTermMonths = pFullTermMonths

      Select Case CMTProportionBalance
        Case CMTProportionBalanceTypes.cmtFrequencyAmounts
          Dim vNumberOfPayments As Integer
          Dim vTotalNumberOfPayments As Integer
          'Calculate number of payment frequencies for the old type
          Dim vPayFreqCode As String = PaymentFrequencyCode
          Dim vOldPayFrequency As PaymentFrequency = mvEnv.GetPaymentFrequency(vPayFreqCode)
          If vOldPayFrequency.Frequency = 1 AndAlso vOldPayFrequency.Interval = 12 Then
            'Annual payer so never charge
            vNumberOfPayments = 1
            vTotalNumberOfPayments = 1    'I.e. 1 - 1 (0) of 1 payments remaining
          Else
            vTotalNumberOfPayments = GetNumberOfFrequencyAmounts(vOldPayFrequency, vRenewalDate.ToString(CAREDateFormat), vNumberOfPayments, pCMTDate)
          End If
          pOldTypeMonths = GetMonthsRemaining(vOldPayFrequency, vTotalNumberOfPayments, (vTotalNumberOfPayments - vNumberOfPayments))
          'Calculate number of payment frequencies for the new type
          If pNewPayFrequency.Frequency = 1 AndAlso pNewPayFrequency.Interval = 12 Then
            'Annual payer so always charge full annual amount
            vTotalNumberOfPayments = 1
            vNumberOfPayments = 1    'I.e. 1 of 1 payments remaining
          Else
            Dim vMinDueDate As Nullable(Of Date)
            If vOldPayFrequency.PaymentFrequencyCode <> pNewPayFrequency.PaymentFrequencyCode Then
              If ((vOldPayFrequency.Frequency * vOldPayFrequency.Interval) = 12 AndAlso (pNewPayFrequency.Frequency * pNewPayFrequency.Interval) = 12) _
              AndAlso (vOldPayFrequency.Frequency > 1 AndAlso pNewPayFrequency.Frequency > vOldPayFrequency.Frequency) Then
                'Payment Frequency changed from (e.g.) quarterly to monthly so calculate payments from next installment date for old pay frequency
                If CDate(NextPaymentDue) > pCMTDate Then pCMTDate = CDate(NextPaymentDue)
                If pOldTypeMonths > 0 Then
                  Dim vMinNPD As Date = TermStartDate.AddMonths(pOldTypeMonths)
                  If vMinNPD > pCMTDate Then
                    vMinDueDate = vMinNPD
                  End If
                End If
              End If
            End If
            vTotalNumberOfPayments = GetNumberOfFrequencyAmounts(pNewPayFrequency, vRenewalDate.ToString(CAREDateFormat), vNumberOfPayments, pCMTDate, vMinDueDate)
          End If
          pNewTypeMonths = GetMonthsRemaining(pNewPayFrequency, vTotalNumberOfPayments, vNumberOfPayments)

        Case CMTProportionBalanceTypes.cmtMonths
          pNewTypeMonths = GetProRataMonthsRemaining(pFullTermMonths, pCMTDate, vRenewalDate)
          pOldTypeMonths = (pFullTermMonths - pNewTypeMonths)

        Case Else   'cmtNone
          pOldTypeMonths = 0
          pNewTypeMonths = pFullTermMonths

      End Select

      pUseProrateMonths = False
      If CMTProportionBalance <> CMTProportionBalanceTypes.cmtNone Then
        'If first year is not full term (fixed renewal memberships) then may need to calculate costs accordingly
        Dim vJoinDateString As String = pJoinedDate.ToString(CAREDateFormat)
        pUseProrateMonths = GetCMTProRateNumberOfMonths(pNewTypeMonths, vRenewalDate.ToString(CAREDateFormat), pChargeTermMonths, vJoinDateString)
        pJoinedDate = CDate(vJoinDateString)  'Date may have changed above
      End If
      If pChargeTermMonths > pFullTermMonths Then pChargeTermMonths = pFullTermMonths

    End Sub

    Private Sub ProcessCMTForSave(ByVal pParams As CDBParameters, ByVal pPPParams As CDBParameters, ByVal pTraderTransaction As TraderTransaction, ByRef pReApplyCMT As Boolean, ByVal pAdditionalParams As CDBParameters)
      'Only to be called from SavePaymentPlan
      Dim vCMTDate As Date = CDate(pPPParams.OptionalValue("CMTDate", TodaysDate))
      Dim vWriteOffMissedPayments As Boolean = pPPParams.ParameterExists("WriteOffMissedPayments").Bool

      LoadMembers()
      SetDetailLineTypesForSC(True)
      LoadSubscriptions()
      mvJoinDate = CDate(pPPParams("Joined").Value)
      TraderBatchCategory = pParams.ParameterExists("TRD_BatchCategory").Value

      Dim vNewMembershipType As MembershipType = mvEnv.MembershipType((pPPParams("MembershipType").Value))
      Dim vAdvancedCMT As Boolean = CanUseAdvancedCMT(vNewMembershipType)

      Dim vGotCMTData As Boolean
      If vAdvancedCMT = True AndAlso pTraderTransaction.CMTOldPPDLines.Count > 0 Then
        'We have come back from Trader with all the CMT data for old & new detail lines
        'First we need to replace the current details with those from the client as they contain all the pro-rating data
        'The new detail lines are in the TraderTransaction.TraderPPDLines collection
        AddDetails(pTraderTransaction.CMTOldPPDLines, True)
        vGotCMTData = True
      End If
      SetupCMT(pPPParams, pTraderTransaction, vCMTDate, mvJoinDate, vGotCMTData)

      'Set local variables for the prices etc.
      'Old membership type
      Dim vFullOldMshipPrice As Double
      Dim vProratedOldMshipPrice As Double
      Dim vProratedOldMshipPricelessOther As Double
      Dim vOtherLinesPrice As Double
      Dim vOldMshipBalance As Double
      Dim vOldMShipArrears As Double
      'New membership type
      Dim vFullNewMshipPrice As Double
      Dim vProratedNewMShipPrice As Double
      Dim vNewMshipBalance As Double

      Dim vFixedAmount As String = ""
      Dim vNewPrice As Double
      Dim vNewRenewalAmount As Double

      If mvSecondCMT Then
        vFullOldMshipPrice = pAdditionalParams("FullOldMshipPrice2").DoubleValue
        vProratedOldMshipPrice = 0   'pAdditionalParams("ProratedOldMshipPrice2").DoubleValue
        vOtherLinesPrice = pAdditionalParams("OtherLinesPrice2").DoubleValue
        vProratedOldMshipPricelessOther = 0   'FixTwoPlaces(vProratedOldMshipPrice - vOtherLinesPrice)
        vOldMshipBalance = pAdditionalParams("OldMshipBalance2").DoubleValue
        vOldMShipArrears = pAdditionalParams("OldMShipArrears2").DoubleValue
        vFixedAmount = pAdditionalParams("Amount2").Value
        'vFullNewMshipPrice = pAdditionalParams("FullNewMshipPrice2").DoubleValue
        'vProratedNewMShipPrice = vFullNewMshipPrice   'pAdditionalParams("ProratedNewMshipPrice2").DoubleValue
        'vNewMshipBalance = pAdditionalParams("NewMshipBalance2").DoubleValue
        vNewPrice = pAdditionalParams("NewPrice2").DoubleValue
        vNewRenewalAmount = pAdditionalParams("NewRenewalAmount2").DoubleValue
        With mvCMTNewPricing
          vFullNewMshipPrice = FixTwoPlaces(.MembershipPrice + .EntitlementPrice)
          vProratedNewMShipPrice = FixTwoPlaces(.ProratedMembershipPrice + .ProratedEntitlementPrice)
          vNewMshipBalance = FixTwoPlaces(.MembershipBalance + .EntitlementBalance)
        End With
        vWriteOffMissedPayments = False         'Cannot do this on the second CMT
        'The price may have been changed for next year - Only check of vOtherPrice
        If Amount.Length = 0 Then GetDetailBalance(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge, 0, 0, mvCMTOldPricing.OtherDetailsPrice, CDate(NextPaymentDue), "")
      Else
        With mvCMTOldPricing
          vFullOldMshipPrice = FixTwoPlaces(.MembershipPrice + .EntitlementPrice + .OtherDetailsPrice)
          vProratedOldMshipPrice = FixTwoPlaces(.ProratedMembershipPrice + .ProratedEntitlementPrice + .ProratedOtherDetailsPrice)
          vProratedOldMshipPricelessOther = FixTwoPlaces(vProratedOldMshipPrice - .ProratedOtherDetailsPrice)
          vOldMshipBalance = FixTwoPlaces(.MembershipBalance + .EntitlementBalance + .OtherDetailsBalance)
          vOldMShipArrears = FixTwoPlaces(.MembershipArrears + .EntitlementArrears + .OtherDetailsArrears)
          vOtherLinesPrice = .OtherDetailsPrice
          vFixedAmount = .MembershipFixedAmount
        End With

        With mvCMTNewPricing
          vFullNewMshipPrice = FixTwoPlaces(.MembershipPrice + .EntitlementPrice)
          vProratedNewMShipPrice = FixTwoPlaces(.ProratedMembershipPrice + .ProratedEntitlementPrice)
          vNewMshipBalance = FixTwoPlaces(.MembershipBalance + .EntitlementBalance)
        End With
        vNewPrice = vFullNewMshipPrice
        vNewRenewalAmount = FixTwoPlaces(vFullNewMshipPrice + vOtherLinesPrice)

        'Add the values to pAdditionalParams so that if required can be used for SecondCMT
        pAdditionalParams.Add("FullOldMshipPrice2", vFullOldMshipPrice)
        pAdditionalParams.Add("ProratedOldMshipPrice2", vProratedOldMshipPrice)
        pAdditionalParams.Add("OtherLinesPrice2", vOtherLinesPrice)
        pAdditionalParams.Add("OldMshipBalance2", vOldMshipBalance)
        pAdditionalParams.Add("OldMShipArrears2", vOldMShipArrears)
        pAdditionalParams.Add("Amount2", vFixedAmount)
        pAdditionalParams.Add("NewPrice2", vNewPrice)
        pAdditionalParams.Add("FullNewMshipPrice2", vFullNewMshipPrice)
        pAdditionalParams.Add("ProratedNewMshipPrice2", vProratedNewMShipPrice)
        pAdditionalParams.Add("NewMshipBalance2", vNewMshipBalance)
        pAdditionalParams.Add("NewRenewalAmount2", vNewRenewalAmount)
      End If

      'Get the member charge line (detail line 1) and setup correct rates on MembershipType's
      Dim vMemberChargingLine As PaymentPlanDetail = CType(Details.Item(GetDetailKeyFromLineNo(1)), PaymentPlanDetail)
      Dim vNewMemberChargingLine As PaymentPlanDetail = pTraderTransaction.TraderPPDLines("1")
      If DetermineMembershipPeriod() = MembershipPeriodTypes.mptFirstPeriod Then
        If MembershipType.FirstPeriodsRate <> vMemberChargingLine.RateCode Then MembershipType.UserDefinedFirstRate = vMemberChargingLine.RateCode
        If vNewMembershipType.FirstPeriodsRate <> vNewMemberChargingLine.RateCode Then vNewMembershipType.UserDefinedFirstRate = vNewMemberChargingLine.RateCode
      Else
        If MembershipType.SubsequentPeriodsRate <> vMemberChargingLine.RateCode Then MembershipType.UserDefinedSubsequentRate = vMemberChargingLine.RateCode
        If vNewMembershipType.SubsequentPeriodsRate <> vNewMemberChargingLine.RateCode Then vNewMembershipType.UserDefinedSubsequentRate = vNewMemberChargingLine.RateCode
      End If

      'Check any FixedAmount
      Dim vMultipleCMT As Boolean = IsMultipleCMT
      Dim vCMTProrateBalance As Boolean = CanProrateCMTBalance(vNewMembershipType)
      Dim vUpdatePPFixedAmount As Nullable(Of Boolean)
      'If Amount is set, must restrict balance to the Amount
      If Amount.Length > 0 AndAlso Val(Amount) <> (vNewPrice + vOtherLinesPrice) Then
        If pParams.Exists("UpdatePPFixedAmount") = False Then
          If vMultipleCMT Then
            'The user must have accepted to update PP Fixed Amount
            RaiseError(DataAccessErrors.daeMustUpdatePPFixedAmount, Amount, String.Format("{0:F}", vNewPrice + vOtherLinesPrice))
          Else
            'Ask user if PP Fixed Amount should be updated
            RaiseError(DataAccessErrors.daeCanUpdatePPFixedAmount, Amount, String.Format("{0:F}", vNewPrice + vOtherLinesPrice))
          End If
        End If
        If pParams("UpdatePPFixedAmount").Bool = False Then
          If vMultipleCMT Then RaiseError(DataAccessErrors.daeMustUpdatePPFixedAmount, Amount, String.Format("{0:F}", vNewPrice + vOtherLinesPrice))
          vUpdatePPFixedAmount = False
          vNewPrice = (DoubleValue(Amount) - vOtherLinesPrice)
          If vNewPrice < 0 Then
            vOtherLinesPrice = vOtherLinesPrice + vNewPrice
            vNewPrice = 0
          End If
        Else
          vUpdatePPFixedAmount = True
        End If
      ElseIf mvSecondCMT AndAlso pParams.Exists("UpdatePPFixedAmount") Then 'Re-use the value for SecondCMT
        vUpdatePPFixedAmount = pParams("UpdatePPFixedAmount").Bool
      End If

      'Calculate AmountPaid
      Dim vAmountPaid As Double
      If vCMTProrateBalance Then
        'Calculate amount paid using OPH
        Dim vRenewalDate As String = RenewalDate
        If RenewalPending Then vRenewalDate = CalculateRenewalDate(vRenewalDate, True)
        Dim vStartDate As String = CalculateRenewalDate(vRenewalDate, False)
        If vMultipleCMT Then
          If mvFirstCMT = False AndAlso mvSecondCMT = False AndAlso RenewalPending = True Then
            'We are preparing to run multiple CMT's so use dates for first CMT (RenewalDate is end of term)
            vRenewalDate = RenewalDate
            vStartDate = CalculateRenewalDate(vRenewalDate, False)
          End If
        End If
        For Each vOPH As OrderPaymentHistory In PaymentHistory(PaymentHistoryOrderByTypes.phobtTransDatePaymentNumber, 0, 0, 0, "NULL", vStartDate, vRenewalDate)
          vAmountPaid += vOPH.Amount
        Next vOPH
        vAmountPaid = FixTwoPlaces(vAmountPaid)

        'Condition added to check if the amount for the  payment plan is paid in advance and transaction date in financial history does not fall between
        'the due date specified in the order payment schedule then we should check the original renewal amount and the balance on the membership to find out
        'the amount that is paid against the old membership
        If FixedRenewalCycle AndAlso PreviousRenewalCycle _
        AndAlso (ProportionalBalanceSetting And (ProportionalBalanceConfigSettings.pbcsFullPayment + ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
        AndAlso vAmountPaid = 0 Then
          'Re-check payments but this time only for in-advance payments that have been processed
          'A processed in-advance payment will show 0 paid due to balancing -'ve & +'ve payments so get the original +'ve payment as well
          For Each vOPH As OrderPaymentHistory In PaymentHistory(PaymentHistoryOrderByTypes.phobtTransDatePaymentNumber, 0, 0, 0, "B", vStartDate, vRenewalDate)
            vAmountPaid += vOPH.Amount
          Next vOPH
          vAmountPaid = FixTwoPlaces(vAmountPaid)
          If vAmountPaid < 0 Then vAmountPaid = 0
        End If

        If (FixedRenewalCycle AndAlso PreviousRenewalCycle _
        AndAlso (ProportionalBalanceSetting And ProportionalBalanceConfigSettings.pbcsFullPayment) = ProportionalBalanceConfigSettings.pbcsFullPayment) Then
          If (StartDate = RenewalDate Or DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(StartDate), CDate(RenewalDate)) <= 1) Then
            'Within first year or renewed to second year & no payments received
            If vAmountPaid > 0 Then
              If (DoubleValue(FirstAmount) = RenewalAmount) Then
                vAmountPaid = FixTwoPlaces(vAmountPaid + InAdvance)
                If vAmountPaid > DoubleValue(FirstAmount) Then vAmountPaid = DoubleValue(FirstAmount)
              End If
            ElseIf RenewalPending = True AndAlso DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(StartDate), CDate(RenewalDate)) = 1 AndAlso (Balance < RenewalAmount) Then
              vAmountPaid = FixTwoPlaces(RenewalAmount - Balance)
            End If
          End If
        End If

        'Look for any write-offs
        Dim vPreviousWriteOffs As Double = SumWriteOffAmounts(vStartDate, vRenewalDate)
        If vPreviousWriteOffs > 0 Then vAmountPaid = FixTwoPlaces(vAmountPaid + vPreviousWriteOffs)

        'Try and handle price changes
        If DetailsContainPricingInfo(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement) = False Then
          'Price has been calculated from the Rates table and so could now be different due to price changes
          'A difference in the calculated price could result in detail lines remaining that should have been removed.
          vFullOldMshipPrice = vOldMshipBalance
          If RenewalPending = False AndAlso vAmountPaid > 0 Then
            vFullOldMshipPrice = FixTwoPlaces(vFullOldMshipPrice + vAmountPaid)
          End If
        End If
      Else
        'No prorating in use
        If FixedRenewalCycle AndAlso PreviousRenewalCycle _
        AndAlso (ProportionalBalanceSetting And (ProportionalBalanceConfigSettings.pbcsFullPayment + ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
        AndAlso (CDate(StartDate) = CDate(RenewalDate) OrElse DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(StartDate), CDate(RenewalDate)) <= System.Math.Abs(Term)) Then
          'Fixed cycle with pro-rating - Membership still in first year
          If CDate(StartDate) = CDate(RenewalDate) Then
            'No payments have been made
            vAmountPaid = 0
          Else
            'vAmountPaid = current balance + sum of OPH (non-I-status records)
            Dim vOldMshipPrice As Double = vOldMshipBalance
            'vOldMshipPrice = vOldMshipBalance + vOtherBalance
            If RenewalPending = False Then
              'If RenewalPending is True then no payments have been made so far this year so don't check payment history
              For Each vOPH As OrderPaymentHistory In PaymentHistory(PaymentHistoryOrderByTypes.phobtTransDatePaymentNumber, 0, 0, 0, "NULL")
                vOldMshipPrice = vOldMshipPrice + vOPH.Amount
              Next vOPH
            End If
            'BR 8768: Note: only need to take account of that paid off Membership since existing Balance is put on "other" lines.
            'Also take account of pro-rating on other line which does not include any Donation amounts.
            'vOtherPrice = DoubleValue(vAmount) + GetProrataBalance(vOtherPrice - DoubleValue(vAmount), mvJoinDate.ToString(CAREDateFormat))
            'vOldMshipPrice = vOldMshipPrice - vOtherPrice
            'vAmountPaid = vOldMshipPrice - vOldMshipBalance
            vAmountPaid = FixTwoPlaces(vFullOldMshipPrice - vOldMshipBalance)
          End If
        Else
          'BR 8768: Note: only need to take account of that paid off Membership since existing Balance is put on "other" lines
          vAmountPaid = FixTwoPlaces(vFullOldMshipPrice - vOldMshipBalance)
        End If
      End If

      If vMultipleCMT OrElse mvSecondCMT Then
        'We have run the renewals and reminders and there are some payments left for current year
        If mvSecondCMT = False Then
          mvFirstCMT = True
          pReApplyCMT = True
          If Not vCMTProrateBalance Then vAmountPaid = FixTwoPlaces(vFullOldMshipPrice - (vOldMshipBalance - vFullOldMshipPrice)) 'vOldMshipPrice - (vOldMshipBalance - vOldMshipPrice)
          mvCMTOldPricing.ProrateFirstCMTCosts(Balance, RenewalAmount, vAmountPaid)
          With mvCMTOldPricing
            vFullOldMshipPrice = FixTwoPlaces(.MembershipPrice + .EntitlementPrice + .OtherDetailsPrice)
            vProratedOldMshipPrice = FixTwoPlaces(.ProratedMembershipPrice + .ProratedEntitlementPrice + .ProratedOtherDetailsPrice)
            vProratedOldMshipPricelessOther = FixTwoPlaces(vProratedOldMshipPrice - .ProratedOtherDetailsPrice)
            vOldMshipBalance = FixTwoPlaces(.MembershipBalance + .EntitlementBalance + .OtherDetailsBalance)
            'vOldMShipArrears = FixTwoPlaces(.MembershipArrears + .EntitlementArrears + .OtherDetailsArrears)
            vOldMShipArrears = 0    'Arrears relates to the following year
            vOtherLinesPrice = .OtherDetailsPrice
            vFixedAmount = .MembershipFixedAmount
          End With
          pAdditionalParams("OldMShipArrears2").Value = "0"
          RenewalDate = CalculateRenewalDate(RenewalDate, False)
        ElseIf mvSecondCMT Then
          vAmountPaid = 0
          vOldMshipBalance = 0
          'vOtherBalance = 0
        End If
      ElseIf vAmountPaid < 0 Then
        vAmountPaid = 0 'Just in case!!
      End If

      'Calculate any over-paid amount
      Dim vOverpaidAmount As Double = FixTwoPlaces(vAmountPaid - FixTwoPlaces(vProratedOldMshipPrice + vProratedNewMShipPrice))
      If vOverpaidAmount < 0 Then vOverpaidAmount = 0

      'Remove old Detail lines
      Dim vNewPriceForCMTBalance As Double
      Dim vPaymentDiff As Double
      If (Not vUpdatePPFixedAmount.HasValue OrElse vUpdatePPFixedAmount.Value = True) AndAlso vCMTProrateBalance = True Then
        Dim vBalanceWrittenOff As Double = 0
        If vWriteOffMissedPayments Then
          Dim vOrigPaymentFrequency As PaymentFrequency = mvEnv.GetPaymentFrequency(mvClassFields.Item(PaymentPlanFields.ofPaymentFrequency).Value)
          vBalanceWrittenOff = WriteOffMissedPayments(pPPParams, vOrigPaymentFrequency, Today, True)
          pPPParams.Add("MissedPaymentsWrittenOff", vBalanceWrittenOff)
        End If

        Dim vOldOtherPrice As Double = mvCMTOldPricing.ProratedOtherDetailsPrice
        vPaymentDiff = FixTwoPlaces((vProratedOldMshipPricelessOther - vAmountPaid) - vBalanceWrittenOff)   'FixTwoPlaces(((vOldMshipPrice - vOldOtherPrice) - vAmountPaid) - vBalanceWrittenOff)
        Dim vNewPPDTotalBal As Double = vProratedNewMShipPrice
        If vWriteOffMissedPayments = True AndAlso vPaymentDiff >= 0 Then
          If ((PaymentFrequencyFrequency * PaymentFrequencyInterval) = 12 AndAlso PaymentFrequencyFrequency > 1) AndAlso mvSecondCMT = False Then
            'Everything written off so just remove all old detail lines and any remainder that could not be writtin-off add to the new member cost
            vNewPPDTotalBal = FixTwoPlaces(vNewPPDTotalBal + vPaymentDiff)
            vPaymentDiff = 0
          End If
        End If

        'Now need to reduce the PPD Balances
        'First lets hold relevant details in a collection for use later if there is an excess payment to be handled
        mvCMTProductDetails = New CollectionList(Of CMTExcessPaymentProductDetail)
        If vAdvancedCMT = False Then
          If vPaymentDiff > 0 Then
            'Reset Balances on old detail lines (these will need to be kept as they were not fully paid)
            mvCMTOldPricing.SetOldTypeDetailBalances(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, vPaymentDiff, vProratedOldMshipPricelessOther)
          End If
        Else
          If vBalanceWrittenOff > 0 Then mvCMTOldPricing.WriteOffOldTypeDetailBalances(vBalanceWrittenOff)
          mvCMTOldPricing.SetOldTypeDetailBalances(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, vOldMshipBalance, vProratedOldMshipPricelessOther)
        End If
        Dim vCMTOldPrice As Double
        For Each vOldPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
          With vOldPPD
            Select Case vOldPPD.CMTProrateLineType
              Case Access.MembershipType.CMTProrateCosts.FullCharge
                vCMTOldPrice = .GetFullPrice()
              Case Access.MembershipType.CMTProrateCosts.NoCharge
                vCMTOldPrice = 0
              Case Else
                vCMTOldPrice = .GetProratedPrice()
            End Select
            Dim vCMTProductDetail As New CMTExcessPaymentProductDetail(.ProductCode, .RateCode, vCMTOldPrice, .Balance, .AmountPaid)
            mvCMTProductDetails.Add((mvCMTProductDetails.Count + 1).ToString, vCMTProductDetail)
          End With
        Next
        If vAdvancedCMT = False AndAlso mvSecondCMT = False Then RemoveCMTLines(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, pPPParams("CancellationReason").Value, (vPaymentDiff > 0), False) 'Any detail lines with an outstanding Balance will be kept
        Dim vNewOtherPrice As Double = mvCMTNewPricing.ProratedOtherDetailsPrice
        If (mvCMTNewTermMonths > 0 AndAlso mvCMTNewTermMonths < 12) OrElse vNewPPDTotalBal > 0 Then mvCMTNewPricing.SetNewTypeDetailBalances(vNewPPDTotalBal, mvCMTNewTermMonths)
        'If vAdvancedCMT Then RemoveCMTLines(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, pPPParams("CancellationReason").Value, (vPaymentDiff > 0)) 'Any detail lines with an outstanding Balance will be kept

        If mvFirstCMT AndAlso CMTProportionBalance = CMTProportionBalanceTypes.cmtMonths Then
          For Each vPPDetail As PaymentPlanDetail In Details
            If (vPPDetail.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge) > 0 Then
              Dim vPPDetailAmount As Double
              If vPPDetail.Amount.Length > 0 Then
                vPPDetailAmount = DoubleValue(vPPDetail.Amount)
              Else
                vPPDetailAmount = vPPDetail.GetFullPrice()
              End If
              If vPPDetail.Balance > vPPDetailAmount Then vPPDetail.Balance = FixTwoPlaces(vPPDetail.Balance - vPPDetailAmount)
            End If
          Next
        End If

        If FixedRenewalCycle AndAlso PreviousRenewalCycle _
        AndAlso (ProportionalBalanceSetting And (ProportionalBalanceConfigSettings.pbcsFullPayment)) = ProportionalBalanceConfigSettings.pbcsFullPayment _
        AndAlso (CDate(StartDate) = CDate(RenewalDate) OrElse DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(StartDate), CDate(RenewalDate)) <= 1) Then
          With mvCMTOldPricing
            vNewPrice = If(DoubleValue(FirstAmount) = RenewalAmount, FixTwoPlaces(.MembershipBalance + .EntitlementBalance + .OtherDetailsBalance), vProratedOldMshipPrice)
          End With
        Else
          vNewPrice = FixTwoPlaces(vProratedOldMshipPricelessOther + vOtherLinesPrice + vOldMShipArrears)     'FixTwoPlaces(.ProratedMembershipPrice + .ProratedEntitlementPrice + .ProratedOtherDetailsPrice)
        End If
        With mvCMTNewPricing
          vNewPrice = FixTwoPlaces(vNewPrice + .ProratedMembershipPrice + .ProratedEntitlementPrice)
        End With
        If mvCMTNewTermMonths > 0 OrElse vNewPPDTotalBal > 0 Then
          'Reset vAmountPaid to just be the amount paid for the new Membership detail lines
          ' If CMTProportionBalance = CMTProportionBalanceTypes.cmtFrequencyAmounts Then
          '  In this situation, any ppdltOtherCharge lines are re-set back to full cost even if they had been paid so have to take this into account
          '  vAmountPaid will be amount paid - (amount allocated to pro-rated old m/ship cost - the amount already allocated as part of the old m/ship figure) i.e. just deduct the amount that relates to the m/ship fees/entitlements etc.
          '  vAmountPaid = FixTwoPlaces(vAmountPaid - vProratedOldMshipPricelessOther)
          ' Else
          '  vAmountPaid will be amount paid - amount allocated to pro-rated old m/ship cost - (full cost of ppdltOtherCharge lines - their outstanding balance - the amount already allocated as part of the old m/ship figure)
          vAmountPaid = FixTwoPlaces(vAmountPaid - vProratedOldMshipPricelessOther)
          vAmountPaid = FixTwoPlaces(vAmountPaid - (vOtherLinesPrice - mvCMTOldPricing.OtherDetailsBalance))
          ' End If
        Else
          If CMTProportionBalance = CMTProportionBalanceTypes.cmtFrequencyAmounts AndAlso mvFirstCMT = True Then
            'vAmountPaid will remian unchanged
            vFullNewMshipPrice = FixTwoPlaces(vFullOldMshipPrice + vOldOtherPrice)
          Else
            Dim vDiff As Double = 0
            If Balance > 0 Then vDiff = FixTwoPlaces(vFullOldMshipPrice - vAmountPaid)
            If vDiff < 0 Then vDiff = 0
            vAmountPaid = FixTwoPlaces(vFullOldMshipPrice - vOtherLinesPrice - vDiff)   'Use the normal price of the new membership to set the balance to zero
          End If
          vNewPrice = FixTwoPlaces(vNewPrice + vOtherLinesPrice)
          If vNewPrice < FixTwoPlaces(vFullOldMshipPrice + vPaymentDiff) Then vNewPriceForCMTBalance = FixTwoPlaces(vFullOldMshipPrice + vPaymentDiff)
          'If vWriteOffMissedPayments = True AndAlso CMTProportionBalance = CMTProportionBalanceTypes.cmtFrequencyAmounts Then vFullOldMshipPrice = vNewPrice
        End If
        If vAmountPaid < 0 Then vAmountPaid = 0
        Dim vFullOtherPrice As Double = 0
        If vOldOtherPrice <> 0 Then
          'Need to increase vFullNewMshipPrice by the full annual amount of 'Other'-type detail lines
          For Each vPPD As PaymentPlanDetail In mvDetails
            If (vPPD.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge) = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge Then
              With vPPD
                If .Amount.Length > 0 Then
                  vFullOtherPrice += DoubleValue(.Amount)
                ElseIf .HasPriceInfo = True AndAlso .UnitPrice <> 0 Then
                  vFullOtherPrice += .UnitPrice
                Else
                  vFullOtherPrice += vPPD.Price(vCMTDate, mvEnv.VATRate(vPPD.Product.ProductVatCategory, Payer.VATCategory), vPPD.Quantity, False, True)
                End If
                If vMultipleCMT AndAlso mvFirstCMT AndAlso .Balance <> mvCMTOldPricing.OtherDetailsBalance Then
                  'If the PP has been renewed and the Other detail line balance doesn't equal the pre-renewal Other detail line balance
                  'then reset balance to pre-renewal Other detail line balance value such that first CMT figures are setup correctly
                  .Balance = mvCMTOldPricing.OtherDetailsBalance
                End If
              End With
            End If
          Next
        End If
        vFullNewMshipPrice = FixTwoPlaces(vFullNewMshipPrice + vFullOtherPrice)
      Else
        'Remove the Old Membership Lines - handle any entitlements
        RemoveCMTLines(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, pPPParams("CancellationReason").Value, False, False) 'Any detail lines with an outstanding Balance will be kept
        'Add other price as it is not already added to the price - new cost of membership is just the cost of the new membership type plus any additional lines from the old type
        vNewPrice = FixTwoPlaces(vNewPrice + vOtherLinesPrice)
      End If
      vNewRenewalAmount = vNewPrice
      If vNewPriceForCMTBalance = 0 Then vNewPriceForCMTBalance = vNewPrice

      'Add any remaining PPDLines to the new lines to be added
      If mvSecondCMT = False AndAlso vAdvancedCMT = False Then 'Do not add the lines again when applying SecondCMT
        Dim vNewDetailNumber As Integer
        vNewDetailNumber = pTraderTransaction.TraderPPDLines.Count
        For Each vPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
          vNewDetailNumber = vNewDetailNumber + 1
          pTraderTransaction.TraderPPDLines.AddItem(vPPD, CStr(vNewDetailNumber))
        Next
      End If

      'Calculate new Membership Balance
      Dim vNewArrears As Double
      Dim vCostDiff As Double = 0
      If mvFirstCMT = True AndAlso vAdvancedCMT = False AndAlso mvCMTOldPricing.MembershipPrice > mvCMTNewPricing.MembershipPrice Then
        'For a First CMT (i.e. CMT after renewal and before the renewal date), if cost is decreasing then reduce amount paid by the difference
        vCostDiff = FixTwoPlaces(mvCMTOldPricing.MembershipPrice - mvCMTNewPricing.MembershipPrice)
      End If
      Dim vNewBalance As Double = CalculateNewCMTBalance(vAdvancedCMT, vCMTDate, vNewPrice, (vAmountPaid - vCostDiff), vUpdatePPFixedAmount, vNewArrears)
      If vAdvancedCMT = False AndAlso mvSecondCMT = True Then RemoveCMTLines(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, pPPParams("CancellationReason").Value, False, False) 'Any detail lines with an outstanding Balance will be kept

      'Add new detail lines to collection for use when dealing with excess payments
      If mvCMTProductDetails Is Nothing Then mvCMTProductDetails = New CollectionList(Of CMTExcessPaymentProductDetail) 'Should never happen, but just in case!
      Dim vCMTNewPrice As Double
      For Each vNewPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
        With vNewPPD
          Select Case .CMTProrateLineType
            Case Access.MembershipType.CMTProrateCosts.FullCharge
              vCMTNewPrice = .GetFullPrice()
            Case Access.MembershipType.CMTProrateCosts.NoCharge
              vCMTNewPrice = 0
            Case Else
              vCMTNewPrice = .GetProratedPrice()
          End Select
          Dim vCMTProductDetail As New CMTExcessPaymentProductDetail(.ProductCode, .RateCode, vCMTNewPrice, .Balance, 0)
          mvCMTProductDetails.Add((mvCMTProductDetails.Count + 1).ToString, vCMTProductDetail)
        End With
      Next

      mvCMTInAdvanceDetails = New CollectionList(Of CMTExcessPaymentDetail)
      mvCMTRefundDetails = New CollectionList(Of CMTExcessPaymentDetail)
      If vAdvancedCMT Then
        vOverpaidAmount = 0
        For Each vPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
          With vPPD
            If .ExcessPaymentAmount <> 0 Then
              Select Case .CMTExcessPaymentType
                Case CmtExcessPayment.CMTExcessPaymentTypes.Refund, CmtExcessPayment.CMTExcessPaymentTypes.Retain
                  If .CMTAdjustmentProductCode.Length = 0 Then
                    If .DetailType = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Then
                      RaiseError(DataAccessErrors.daeCMTMemberTypeRefundProductNotSet, MembershipTypeCode)
                    Else
                      RaiseError(DataAccessErrors.daeCMTEntitlementRefundProductNotSet, .ProductCode)
                    End If
                  End If
                  mvCMTRefundDetails.Add(.DetailNumber.ToString, New CMTExcessPaymentDetail(.ProductCode, .RateCode, .CMTExcessPaymentType, .ExcessPaymentAmount, .CMTAdjustmentProductCode, .CMTAdjustmentRateCode))
                  vOverpaidAmount = FixTwoPlaces(vOverpaidAmount + .ExcessPaymentAmount)
                Case Else   'CmtExcessPayment.CMTExcessPaymentTypes.CarryForward, CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse
                  'Any excess from a carry forward becomes in-advance
                  mvCMTInAdvanceDetails.Add(.DetailNumber.ToString, New CMTExcessPaymentDetail(.ProductCode, .RateCode, .CMTExcessPaymentType, .ExcessPaymentAmount, "", ""))
                  vOverpaidAmount = FixTwoPlaces(vOverpaidAmount + .ExcessPaymentAmount)
              End Select
            End If
          End With
        Next

        'For an Advanced CMT must remove paid lines AFTER balances have been calculated for new lines
        Dim vKeepPaidFullChargeLines As Boolean = False     'Set this to True if paid old detail lines that have CMTProrateType = FullCharge are to be kept
        RemoveCMTLines(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, pPPParams("CancellationReason").Value, (vPaymentDiff > 0), vKeepPaidFullChargeLines) 'Any detail lines with an outstanding Balance will be kept
        'Add any remaining PPDLines to the new lines to be added
        Dim vNewDetailNumber As Integer
        vNewDetailNumber = pTraderTransaction.TraderPPDLines.Count
        For Each vPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
          vNewDetailNumber = vNewDetailNumber + 1
          pTraderTransaction.TraderPPDLines.AddItem(vPPD, CStr(vNewDetailNumber))
          vNewBalance += vPPD.Balance   'These now need to be included in the new Balance
        Next vPPD
        vNewBalance = FixTwoPlaces(vNewBalance)
      Else
        'Need to add in-advance amount when not using Advanced CMT
        If vOverpaidAmount < 0 Then vOverpaidAmount = 0
        If vOverpaidAmount > 0 Then
          'Add a single CMTExcessPaymentDetail record for the full over-payment amount (CMTExcessPaymentTypes.CarryForward results in an in-advance payment)
          'Set the Original Product & Rate to the first detail line values and leave the Adjustment Product & Rate unset
          Dim vCMTExcessPayment As New CMTExcessPaymentDetail(vMemberChargingLine.ProductCode, vMemberChargingLine.RateCode, CmtExcessPayment.CMTExcessPaymentTypes.CarryForward, vOverpaidAmount, "", "")
          mvCMTInAdvanceDetails.Add((mvCMTInAdvanceDetails.Count + 1).ToString, vCMTExcessPayment)
        End If
      End If
      If vCMTProrateBalance = True AndAlso vOverpaidAmount > 0 AndAlso mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReverseTransType).Length = 0 Then RaiseError(DataAccessErrors.daeReversalTransTypeNotSetUp)

      If MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfMonthlyTerm And vOldMshipBalance = 0 Then vNewBalance = 0
      'Get the Membership Card Expiry Date
      Dim vMshipCardExpires As String = ""
      For Each vMember As Member In CurrentMembers
        If Len(vMshipCardExpires) = 0 Then
          If Len(vMember.MembershipCardExpires) > 0 Then
            vMshipCardExpires = vMember.MembershipCardExpires
            Exit For
          End If
        End If
      Next vMember

      'Find joint contact
      Dim vJointContact As Contact = New Contact(mvEnv)
      vJointContact.Init()
      If vNewMembershipType.MembersPerOrder = 2 And pPPParams("NumberOfMembers").IntegerValue = 2 Then
        If pPPParams("OneYearGift").Bool _
          Or Not pPPParams("GiftMembership").Bool Then
          'The joint contact on one year gift will become the payer when the membership is renewed, so joint contact may be created here.
          vJointContact = pTraderTransaction.GetMembershipJointContact(vNewMembershipType.MembershipTypeCode, pPPParams("Source").Value)
          If vJointContact.ContactNumber > 0 _
            And Not pPPParams("GiftMembership").Bool Then
            'Ensure that Payer is set to the Joint Contact
            pPPParams("PayerContactNumber").Value = vJointContact.ContactNumber.ToString
            pPPParams("PayerAddressNumber").Value = vJointContact.Address.AddressNumber.ToString
            Dim vRow As Integer = 0
            For Each vPPDetail As PaymentPlanDetail In pTraderTransaction.TraderPPDLines
              vRow = vRow + 1
              If vRow = 1 Then
                If vNewMembershipType.ChargeIndividualMembers <> "Y" Then
                  'If charge individual members then the 1st detail line should go to the 1st member not the payer
                  vPPDetail.SetContactAndAddress(vJointContact.ContactNumber, vJointContact.Address.AddressNumber)
                End If
              ElseIf vPPDetail.MemberOrPayer = "P" Then
                vPPDetail.SetContactAndAddress(vJointContact.ContactNumber, vJointContact.Address.AddressNumber)
              End If
            Next vPPDetail
          End If
        End If
      End If
      If pPPParams("GiftMembership").Bool Then
        'Payer needs to be the contact from the tpMembershipPayer page
        'For vMembershipType.PayerRequired = "M" this has already been done
        If pPPParams("PayerContactNumber").IntegerValue <> pPPParams("MembershipPayerContactNumber").IntegerValue AndAlso pPPParams("MembershipPayerContactNumber").IntegerValue > 0 Then
          pPPParams("PayerContactNumber").Value = pPPParams("MembershipPayerContactNumber").Value
          pPPParams("PayerAddressNumber").Value = pPPParams("MembershipPayerAddressNumber").Value
        End If
      End If

      'Now add parameters for all these values
      With pPPParams
        .Add("AmountPaid", CDBField.FieldTypes.cftNumeric, vAmountPaid.ToString)
        .Add("NewBalance", CDBField.FieldTypes.cftNumeric, vNewBalance.ToString)
        .Add("NewPrice", CDBField.FieldTypes.cftNumeric, vNewPrice.ToString)
        .Add("OldMembershipBalance", CDBField.FieldTypes.cftNumeric, vOldMshipBalance.ToString)
        .Add("OldMembershipPrice", CDBField.FieldTypes.cftNumeric, vFullOldMshipPrice.ToString)   'vOldMshipPrice.ToString)
        .Add("OrigMembershipCardExpires", CDBField.FieldTypes.cftCharacter, vMshipCardExpires)
        .Add("OtherBalance", CDBField.FieldTypes.cftNumeric, mvCMTOldPricing.OtherDetailsBalance.ToString)    'vOtherBalance.ToString)
        .Add("OtherPrice", CDBField.FieldTypes.cftNumeric, vOtherLinesPrice.ToString)
        If .Exists("RenewalAmount") = False Then .Add("RenewalAmount", CDBField.FieldTypes.cftNumeric)
        .Item("RenewalAmount").Value = vNewRenewalAmount.ToString
        .Add("OverPaymentAmount", CDBField.FieldTypes.cftNumeric, vOverpaidAmount.ToString)
        .Add("CMTProRateBalance", CDBField.FieldTypes.cftCharacter, BooleanString(vCMTProrateBalance))
        .Add("MonthsRemaining", CDBField.FieldTypes.cftLong, mvCMTNewTermMonths.ToString)
        .Add("IsAdvancedCMT", BooleanString(vAdvancedCMT))
        If vUpdatePPFixedAmount.HasValue AndAlso vUpdatePPFixedAmount.Value = True Then .Add("FullMembershipPrice", CDBField.FieldTypes.cftNumeric, vFullNewMshipPrice.ToString) 'vFullPrice.ToString) 'Full price of Membership + Entitlements + Other charge lines (without any pro-rating applied)
        mvFullMembershipPriceCMT = vFullNewMshipPrice   'vFullPrice 'Always set it at the end of this method
      End With
    End Sub

    ''' <summary>Remove detail lines for the old membership type.</summary>
    Private Sub RemoveCMTLines(ByVal pDetailTypes As PaymentPlanDetail.PaymentPlanDetailTypes, ByVal pCancellationReason As String, ByVal pZeroBalanceLinesOnly As Boolean, ByVal pKeepPaidFullChargeLines As Boolean)
      Dim vRenewalDate As String = RenewalDate
      If RenewalPending Then vRenewalDate = CalculateRenewalDate(vRenewalDate, True)
      For Each vPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
        With vPPD
          If (.DetailType And pDetailTypes) > 0 Then
            If (.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement) > 0 And CurrentSubscriptions.Count() > 0 Then
              For Each vSubscription As Subscription In CurrentSubscriptions
                If vSubscription.ContactNumber = .ContactNumber And vSubscription.Product = .ProductCode And vSubscription.CancellationReason.Length = 0 Then
                  'Found the corresponding Subscription
                  vSubscription.Cancel(pCancellationReason)
                  Exit For
                End If
              Next vSubscription
            End If
            If .Arrears > 0 Then
              'Line to be removed is in arrears - can't remove it just flag as an arrears line
              .SetDetailArrears(vRenewalDate)
            ElseIf (.Balance > 0 AndAlso pZeroBalanceLinesOnly = True) Then
              .SetCMTLineNoRenewalRequired(vRenewalDate)
            ElseIf (.Balance = 0 AndAlso pKeepPaidFullChargeLines = True AndAlso .CMTProrateLineType = Access.MembershipType.CMTProrateCosts.FullCharge) Then
              .SetCMTLineNoRenewalRequired(vRenewalDate)
            Else
              mvCMTOldPricing.OldDetails.Remove((mvCMTOldPricing.GetOldDetailKeyFromLineNo(vPPD.DetailNumber)))
              mvCMTOldPricing.ReNumberOldDetailKeys()
            End If
          End If
        End With
      Next
    End Sub

        Private Function CalculateNewCMTBalance(ByVal pAdvancedCMT As Boolean, ByVal pCMTDate As Date, ByVal pNewPrice As Double, ByVal pAmountPaid As Double, ByVal pUpdatePPFixedAmount As Nullable(Of Boolean), ByRef pNewPPArrears As Double) As Double
            'BR18903 changes required in this routine??
            Dim vNewPPBalance As Double
            Dim vNewPPArrears As Double

            Dim vDiscountOS As Double
            'Pre-process to total discounts
            For Each vPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                If vPPD.Balance < 0 Then
                    vDiscountOS = FixTwoPlaces(vDiscountOS + (vPPD.Balance * -1))
                End If
            Next

            If pAdvancedCMT Then
                '(1) Deal with Carry Forward first
                Dim vExcessPayment As Double
                For Each vOldPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
                    If vOldPPD.ExcessPaymentAmount > 0 AndAlso vOldPPD.CMTExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.CarryForward Then
                        vExcessPayment = vOldPPD.ExcessPaymentAmount
                        If vOldPPD.DetailType = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge Then
                            'Find a Charge detail line in the new details collection
                            For Each vNewPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                                If vNewPPD.DetailType = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge _
                                AndAlso vNewPPD.Balance > 0 AndAlso vExcessPayment > 0 Then
                                    'Apply excess payment against outstanding balance
                                    vNewPPD.CMTApplyExcessPayment(vOldPPD)
                                    vExcessPayment = vOldPPD.ExcessPaymentAmount
                                    If vExcessPayment = 0 Then Exit For
                                End If
                            Next
                        ElseIf (vOldPPD.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive) > 0 Then
                            If vOldPPD.EntitlementSequenceNumber > 0 Then
                                'Find detail line in new collection with same sequence number
                                For Each vNewPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                                    If vNewPPD.EntitlementSequenceNumber = vOldPPD.EntitlementSequenceNumber _
                                    AndAlso (vNewPPD.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive) > 0 Then
                                        'Apply excess payment against outstanding balance
                                        vNewPPD.CMTApplyExcessPayment(vOldPPD)
                                        vExcessPayment = vOldPPD.ExcessPaymentAmount
                                        If vExcessPayment = 0 Then Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next

                '(2) Deal with Retain & Refund
                'For these just leave the excess payment amount as it is, so do nothing for now

                '(3) Finally deal with any Reanalyse
                For Each vOldPPD As PaymentPlanDetail In mvCMTOldPricing.OldDetails
                    If vOldPPD.ExcessPaymentAmount <> 0 AndAlso vOldPPD.CMTExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse Then
                        'Reanalyse the excess against any detail line
                        vExcessPayment = vOldPPD.ExcessPaymentAmount
                        For Each vNewPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                            With vNewPPD
                                If .Balance <> 0 Then
                                    'Apply excess payment against outstanding balance
                                    vNewPPD.CMTApplyExcessPayment(vOldPPD)
                                    vExcessPayment = vOldPPD.ExcessPaymentAmount
                                    If vExcessPayment = 0 Then Exit For
                                End If
                            End With
                        Next
                    End If
                Next
                'Re-calculate balances
                For Each vNewPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                    vNewPPBalance = FixTwoPlaces(vNewPPBalance + vNewPPD.Balance)
                    vNewPPArrears = FixTwoPlaces(vNewPPArrears + vNewPPD.Arrears)
                Next
            Else
                'Not an Advanced CMT
                'Now total the lines
                'Dim vPrice As Double
                Dim vOverPaid As Double
                Dim vPlanTotal As Double
                Dim vLinePaid As Double
                Dim vSecondCMTBalance As Double

                'Now total the lines
                Dim vDetailNumber As Integer = 0
                Dim vPPDBalance As Double
                Dim vPPDArrears As Double
                For Each vPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                    vDetailNumber += 1
                    vPPDBalance = vPPD.Balance
                    vPPDArrears = vPPD.Arrears

                    If mvSecondCMT Then
                        'For a 2nd CMT, use the current PPD Balance so that the end result is old + new balances
                        If vDetailNumber <= mvCMTOldPricing.OldDetails.Count Then
                            Dim vOldPPD As PaymentPlanDetail = CType(mvCMTOldPricing.OldDetails.Item(GetDetailKeyFromLineNo(vDetailNumber)), PaymentPlanDetail)
                            vPPDBalance = vOldPPD.Balance
                            vPPDArrears = vOldPPD.Arrears
                        End If
                        If vPPD.Amount.Length > 0 Then
                            vSecondCMTBalance = DoubleValue(vPPD.Amount)
                        Else
                            vSecondCMTBalance = vPPD.FullPrice(pCMTDate, Payer.VATCategory)
                        End If
                    End If

                    If vPPDBalance < 0 Then vDiscountOS = FixTwoPlaces(vDiscountOS + vPPDBalance)

                    'Check the Pay Plan Limit - ignore arrears
                    vPlanTotal = FixTwoPlaces(vPlanTotal + vPPDBalance - vPPDArrears)

                    If ((vPlanTotal - vDiscountOS) > pNewPrice AndAlso Not (mvFirstCMT AndAlso vPPD.DetailType = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge)) _
                    OrElse (pUpdatePPFixedAmount.HasValue AndAlso pUpdatePPFixedAmount.Value = False) Then
                        'We have reached the Pay Plan Limit
                        vOverPaid = FixTwoPlaces((vPlanTotal - vDiscountOS) - pNewPrice)
                        If vPPDBalance > vOverPaid Then
                            vPPDBalance = FixTwoPlaces(vPPDBalance - vOverPaid)
                        Else
                            vPPDBalance = 0
                        End If
                    ElseIf mvFirstCMT AndAlso CMTProportionBalance = CMTProportionBalanceTypes.cmtNone _
                    AndAlso vPPD.DetailType = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge Then
                        Dim vAmount As Double
                        If vPPD.Amount.Length > 0 Then
                            vAmount = DoubleValue(vPPD.Amount)
                        Else
                            vAmount = vPPD.FullPrice(pCMTDate, Payer.VATCategory)
                        End If
                        If vPPDBalance - vAmount > 0 Then
                            vPPDBalance = FixTwoPlaces(vPPDBalance - vAmount)
                        Else
                            vPPDBalance = 0
                        End If
                    End If
                    vPPD.Balance = vPPDBalance
                    vPPD.Arrears = vPPDArrears
                    'Now take off Amount Paid
                    If pAmountPaid > 0 Then
                        If vPPD.Balance >= pAmountPaid Then
                            vLinePaid = pAmountPaid
                        Else
                            vLinePaid = vPPD.Balance
                        End If
                        vPPD.Balance = FixTwoPlaces(vPPD.Balance - vLinePaid)
                        pAmountPaid = FixTwoPlaces(pAmountPaid - vLinePaid)
                        If vPPD.Arrears > 0 Then
                            If vPPD.Arrears >= vLinePaid Then
                                vPPD.Arrears = FixTwoPlaces(vPPD.Arrears - vLinePaid)
                            Else
                                vPPD.Arrears = 0
                            End If
                        End If
                    End If

                    'Now Keep a running total of balance and arrears
                    If mvSecondCMT Then
                        vPPD.Balance = FixTwoPlaces(vPPD.Balance + vSecondCMTBalance)
                    Else
                        vNewPPBalance = FixTwoPlaces(vNewPPBalance + vPPD.Balance)
                    End If
                    vNewPPArrears = FixTwoPlaces(vNewPPArrears + vPPD.Arrears)
                Next
            End If

            If vNewPPArrears > 0 Then
                'Ensure that any detail lines with arrears have the correct balance (may have just set the Balance to zero which is invalid)
                'Calculate amount of arrears that needs fixing
                Dim vArrearsToAdjust As Double = 0
                For Each vPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                    If vPPD.Arrears <> 0 Then
                        If FixTwoPlaces(vPPD.Balance - vPPD.Arrears) < 0 Then vArrearsToAdjust = FixTwoPlaces(vArrearsToAdjust + Math.Abs(vPPD.Balance - vPPD.Arrears))
                    End If
                Next
                If vArrearsToAdjust > 0 Then
                    'Reduce balances on lines without arrears
                    Dim vAdjBalance As Double = vArrearsToAdjust
                    For Each vPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                        If vPPD.Arrears = 0 Then
                            If vPPD.Balance > vPPD.Arrears Then
                                If vPPD.Balance - vPPD.Arrears >= vAdjBalance Then
                                    vPPD.Balance = FixTwoPlaces(vPPD.Balance - vAdjBalance)
                                    vAdjBalance = 0
                                Else
                                    vAdjBalance = FixTwoPlaces(vAdjBalance - vPPD.Balance)
                                    vPPD.Balance = 0
                                End If
                            End If
                        End If
                        If vAdjBalance = 0 Then Exit For
                    Next
                    'Increase balances on lines with arrears (when balance less than arrears)
                    For Each vPPD As PaymentPlanDetail In mvCMTNewPricing.NewDetails
                        If vPPD.Arrears > 0 Then
                            If vPPD.Balance < vPPD.Arrears Then
                                vArrearsToAdjust = FixTwoPlaces(vArrearsToAdjust - (vPPD.Arrears - vPPD.Balance))
                                vPPD.Balance = vPPD.Arrears
                            End If
                        End If
                    Next
                End If
            End If

            If mvSecondCMT Then vNewPPBalance = pNewPrice
            pNewPPArrears = vNewPPArrears
            Return vNewPPBalance
        End Function

    Private Sub ProcessCMTOverPayment(ByVal pOverPaymentAmount As Double, ByVal pSourceCode As String, ByVal pAdvancedCMT As Boolean)
      Dim vCompanyControl As New CompanyControl
      Dim vFH As New FinancialHistory
      Dim vOPH As New OrderPaymentHistory
      Dim vBatch As Batch = Nothing 'New adjustment Batch
      Dim vBT As BatchTransaction = Nothing 'New adjustment BT
      Dim vFHD As FinancialHistoryDetail
      Dim vFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vAmount As Double
      Dim vEligibleGiftAid As Boolean
      Dim vLineNumber As Integer
      Dim vNewBatchNumber As Integer
      Dim vReversed As Boolean
      Dim vTrans As Boolean

      Dim vRenewalDate As String = RenewalPeriodEnd
      Dim vStartDate As String = CalculateRenewalDate(vRenewalDate, False)
      Dim vTransactionDate As String = TodaysDate()
      Dim vPaymentNumber As Integer = PaymentNumber

      vFH.Init(mvEnv)
      vOPH.Init(mvEnv)

      'Select all OPH, FH & FHD
      Dim vAttrs As String = vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll).Replace("oph.batch_number,", "").Replace("oph.transaction_number,", "").Replace("oph.status", "oph.status AS oph_status")
      vAttrs &= "," & vFH.GetRecordSetFields(FinancialHistory.FinancialHistoryRecordSetTypes.fhrtDetail Or FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers)
      vAttrs = vAttrs.Replace("fh.amount", "fh.amount AS fh_amount").Replace("fh.posted", "fh.posted AS fh_posted")
      vAttrs &= ", bt.eligible_for_gift_aid"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vAttrs &= ", b.currency_code, b.currency_exchange_rate"
      Dim vOrderBy As String = "oph.payment_number DESC"

      Dim vAnsiJoins As New AnsiJoins
      With vAnsiJoins
        .Add("batch_transactions bt", "oph.batch_number", "bt.batch_number", "oph.transaction_number", "bt.transaction_number")
        .Add("financial_history fh", "bt.batch_number", "fh.batch_number", "bt.transaction_number", "fh.transaction_number")
        .Add("batches b", "fh.batch_number", "b.batch_number")
        .AddLeftOuterJoin("order_payment_schedule ops", "oph.order_number", "ops.order_number", "oph.scheduled_payment_number", "ops.scheduled_payment_number")
      End With

      Dim vSQLWhereFields As New CDBFields
      With vSQLWhereFields
        .Add("oph.order_number", PlanNumber)
        .Add("oph.status", CDBField.FieldTypes.cftCharacter, "")
        .Add("oph.scheduled_payment_number", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoNotEqual)
        .Add("ops.due_date", CDBField.FieldTypes.cftDate, vStartDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
        .Add("ops.due_date#2", CDBField.FieldTypes.cftDate, vRenewalDate, CDBField.FieldWhereOperators.fwoBetweenTo Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("oph.scheduled_payment_number#2", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("fh.transaction_date", CDBField.FieldTypes.cftDate, vStartDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
        .Add("fh.transaction_date#2", CDBField.FieldTypes.cftDate, vRenewalDate, CDBField.FieldWhereOperators.fwoBetweenTo Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "order_payment_history oph", vSQLWhereFields, vOrderBy, vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()

      Dim vOrigFHColl As New CollectionList(Of FinancialHistory) 'Original FH records
      Dim vOrigOPHColl As New CollectionList(Of OrderPaymentHistory) 'Original OPH records
      Dim vKey As String
      Dim vCurrencyCode As String = ""
      Dim vExchangeRate As Double

      'Put all selected data into 2 collections
      While vRS.Fetch() = True And (vAmount <= pOverPaymentAmount)
        vEligibleGiftAid = vRS.Fields("eligible_for_gift_aid").Bool
        If vRS.Fields.Exists("currency_code") Then
          vCurrencyCode = vRS.Fields("currency_code").Value
          vExchangeRate = vRS.Fields("currency_exchange_rate").DoubleValue
        End If
        If (vOPH.BatchNumber <> vRS.Fields("batch_number").IntegerValue) Or (vOPH.TransactionNumber <> vRS.Fields("transaction_number").IntegerValue) Or (vOPH.LineNumber <> vRS.Fields("line_number").IntegerValue) Then
          'Either Batch, Transaction or Line has changed
          vOPH = New OrderPaymentHistory
          vOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          vOPH.Status = vRS.Fields("oph_status").Value
          vKey = vOPH.Key
          If vOrigOPHColl.ContainsKey(vKey) Then vKey &= (vOrigOPHColl.Count + 1).ToString
          vOrigOPHColl.Add(vKey, vOPH)
          If (vFH.BatchNumber <> vRS.Fields("batch_number").IntegerValue) Or (vFH.TransactionNumber <> vRS.Fields("transaction_number").IntegerValue) Then
            'Batch or Transaction have changed
            vFH = New FinancialHistory
            vFH.InitFromRecordSet(mvEnv, vRS, FinancialHistory.FinancialHistoryRecordSetTypes.fhrtDetail Or FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers)
            vFH.Amount = vRS.Fields("fh_amount").DoubleValue
            vFH.Posted = vRS.Fields("fh_posted").Value
            vOrigFHColl.Add(vFH.Key, vFH)
          End If
          vAmount = vAmount + vOPH.Amount
        End If
      End While
      vRS.CloseRecordSet()

      'If we found some records, then process the negative side (records are in reverse order)
      'Create all the reversals of the original payments that are going to be adjusted
      vFH = New FinancialHistory
      vFH.Init(mvEnv)
      For vIndex As Integer = vOrigOPHColl.Count - 1 To 0 Step -1
        vOPH = vOrigOPHColl.Item(vIndex)  'CType(vOrigOPHColl.Item(vIndex), OrderPaymentHistory)
        If (vOPH.BatchNumber <> vFH.BatchNumber) Or (vOPH.TransactionNumber <> vFH.TransactionNumber) Then
          vFH = vOrigFHColl.Item((vOPH.BatchNumber.ToString.PadLeft(9, "0"c) & vOPH.TransactionNumber.ToString.PadLeft(4, "0"c)))
        End If
        If vBatch Is Nothing Then
          vBatch = New Batch(mvEnv)
          With vBatch
            .InitNewBatch(mvEnv)
            .BatchType = Batch.BatchTypes.FinancialAdjustment
            .BatchCategory = TraderBatchCategory
            .BankAccount = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBankAccount)
            .ReadyForBanking = True
            .PostedToCashBook = True
            .Picked = "C"
            .SetPayingInSlipPrinted(0)
            .LockBatch()
          End With
          vBT = New BatchTransaction(mvEnv)
          With vBT
            .InitFromBatch(mvEnv, vBatch)
            .ContactNumber = ContactNumber
            .AddressNumber = AddressNumber
            .TransactionDate = vTransactionDate
            .TransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReverseTransType)
            .PaymentMethod = PaymentMethod
            .Receipt = "N"
            .EligibleForGiftAid = vEligibleGiftAid
            .Notes = "CMT Automatic Adjustment"
          End With

          vCompanyControl.InitFromBankAccount(mvEnv, (vBatch.BankAccount))

          If mvEnv.Connection.InTransaction = False Then
            mvEnv.Connection.StartTransaction()
            vTrans = True
          End If

          'Save Batch & BT first
          vBT.SaveChanges()
          vBatch.Save(mvEnv.User.UserID, True)
          vNewBatchNumber = vBatch.BatchNumber
        End If

        vAmount = vAmount + vOPH.Amount
        vLineNumber = vLineNumber + 1
        If vFH.Reverse(vBatch, vBT, Batch.AdjustmentTypes.atAdjustment, (vOPH.LineNumber), False, 0, 0, True) Then
          'Create FH & FHD for negative side as a copy of the originals
          'This will add FHD records to FH.Details
          For Each vFHD In vFH.Details
            If (vOPH.LineNumber = vFHD.LineNumber) Then
              Dim vNewFHD As New FinancialHistoryDetail(mvEnv)
              With vNewFHD
                .CopyValues(vFHD)  'Cannot use Clone as this table does not have a primary key
                .Reverse(vBT.BatchNumber, vBT.TransactionNumber, vLineNumber)
                .Save(mvEnv.User.UserID, True, 0)
              End With
            End If
          Next
          vReversed = True
          vPaymentNumber = vPaymentNumber + 1
        End If
      Next

      'Create negative FH
      If vReversed Then
        Dim vNewFH As New FinancialHistory
        With vNewFH
          .Init(mvEnv)
          .CreateFromBatchTransaction(vBT)
          .Save()
        End With

        'Set new OPH as posted (created by the FH.Reverse)
        vFields = New CDBFields
        vFields.Add("posted", CDBField.FieldTypes.cftCharacter, "Y")
        vWhereFields = New CDBFields
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, vBT.BatchNumber)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, vBT.TransactionNumber)
        mvEnv.Connection.UpdateRecords("order_payment_history", vFields, vWhereFields)
      End If

      'Next, create positive BT & BTA
      If vReversed Then
        Dim vReversalAmount As Double = vBT.Amount
        PaymentNumber = vPaymentNumber    'reset

        vBT = New BatchTransaction(mvEnv)
        With vBT
          .InitFromBatch(mvEnv, vBatch)
          .ContactNumber = ContactNumber
          .AddressNumber = AddressNumber
          .TransactionDate = vTransactionDate
          .TransactionType = "P"
          .PaymentMethod = PaymentMethod
          .Receipt = "N"
          .EligibleForGiftAid = vEligibleGiftAid
          .Notes = "CMT Automatic Adjustment"
        End With

        'We may have just reversed more than the excess payment amount so re-add the payments for the difference
        If vReversalAmount > pOverPaymentAmount Then
          Dim vDiff As Double = Math.Abs(FixTwoPlaces(pOverPaymentAmount - vReversalAmount))
          Dim vPayAmount As Double
          vFH = New FinancialHistory()
          vFH.Init(mvEnv)
          If mvCMTProductDetails Is Nothing Then mvCMTProductDetails = New CollectionList(Of CMTExcessPaymentProductDetail) 'Shouldn't happen but just in case!
          If pAdvancedCMT = True AndAlso mvCMTProductDetails.Count > 0 Then
            'Ensure that the payments are allocated to the correct products
            'First adjust the paid figures to take account of the reversals
            For vIndex As Integer = vOrigOPHColl.Count - 1 To 0 Step -1
              vOPH = vOrigOPHColl.Item(vIndex)
              If vOPH.BatchNumber <> vFH.BatchNumber OrElse vOPH.TransactionNumber <> vFH.TransactionNumber Then
                vFH = vOrigFHColl.Item((vOPH.BatchNumber.ToString.PadLeft(9, "0"c) & vOPH.TransactionNumber.ToString.PadLeft(4, "0"c)))
                For Each vFHD In vFH.Details
                  For Each vCMTPD As CMTExcessPaymentProductDetail In mvCMTProductDetails
                    With vCMTPD
                      If vFHD.ProductCode = .ProductCode AndAlso vFHD.RateCode = .RateCode Then
                        .AmountPaid = FixTwoPlaces(.AmountPaid - vFHD.Amount)
                        Exit For
                      End If
                    End With
                  Next
                Next
              End If
            Next
            'Now create the positive side of the transaction using the correct product/rate/amount's
            Dim vOPHAmount As Double
            For vIndex As Integer = vOrigOPHColl.Count - 1 To 0 Step -1
              vOPH = vOrigOPHColl.Item(vIndex)
              vOPHAmount = vOPH.Amount
              For Each vCMTPD As CMTExcessPaymentProductDetail In mvCMTProductDetails
                With vCMTPD
                  vPayAmount = FixTwoPlaces(.Price - .Balance - .AmountPaid)
                  If vPayAmount <> 0 Then
                    If (vPayAmount > vOPHAmount) AndAlso .Price > 0 Then vPayAmount = vOPHAmount
                    If (vPayAmount > vDiff) AndAlso .Price > 0 Then vPayAmount = vDiff
                  End If
                  If vPayAmount <> 0 Then
                    Dim vCMTExcessPayDetails As New CMTExcessPaymentDetail(vCMTPD.ProductCode, vCMTPD.RateCode, CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse, vPayAmount, vCMTPD.ProductCode, vCMTPD.RateCode)
                    CreatePositiveBTAForOverPaidAmount(vCMTExcessPayDetails, vBT, vPayAmount, vCompanyControl, pSourceCode, IntegerValue(vOPH.ScheduledPaymentNumber))
                    .AmountPaid = FixTwoPlaces(.AmountPaid + vPayAmount)
                    vOPHAmount = FixTwoPlaces(vOPHAmount - vPayAmount)
                    vDiff = FixTwoPlaces(vDiff - vPayAmount)
                    If vOPHAmount = 0 OrElse vDiff = 0 Then Exit For
                  End If
                End With
              Next
              If vDiff = 0 Then Exit For
            Next
          Else
            'The original way of doing it - just allocate the money proportionally across the original products
            Dim vMultiplier As Double
            Dim vSumFHDAmount As Double
            For vIndex As Integer = vOrigOPHColl.Count - 1 To 0 Step -1
              vOPH = vOrigOPHColl.Item(vIndex)
              vMultiplier = (vDiff / vOPH.Amount)
              If vOPH.BatchNumber <> vFH.BatchNumber OrElse vOPH.TransactionNumber <> vFH.TransactionNumber Then
                vFH = vOrigFHColl.Item((vOPH.BatchNumber.ToString.PadLeft(9, "0"c) & vOPH.TransactionNumber.ToString.PadLeft(4, "0"c)))
                For Each vFHD In vFH.Details
                  vPayAmount = FixTwoPlaces(vFHD.Amount * vMultiplier)
                  If FixTwoPlaces(vSumFHDAmount + vFHD.Amount) = vFH.Amount Then
                    'This is the last FHD
                    vPayAmount = vDiff
                  End If
                  If vPayAmount > vDiff Then vPayAmount = vDiff
                  vDiff = FixTwoPlaces(vDiff - vPayAmount)
                  vSumFHDAmount = FixTwoPlaces(vSumFHDAmount + vFHD.Amount)
                  Dim vCMTExcessPayDetails As New CMTExcessPaymentDetail(vFHD.ProductCode, vFHD.RateCode, CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse, vPayAmount, vFHD.ProductCode, vFHD.RateCode)
                  CreatePositiveBTAForOverPaidAmount(vCMTExcessPayDetails, vBT, vPayAmount, vCompanyControl, pSourceCode, IntegerValue(vOPH.ScheduledPaymentNumber))
                  If vDiff <= 0 Then Exit For
                Next
              End If
              If vDiff <= 0 Then Exit For
            Next
          End If

          'Now ensure that original OPS are updated correctly
          Dim vOPS As OrderPaymentSchedule
          vPayAmount = Math.Abs(FixTwoPlaces(pOverPaymentAmount - vReversalAmount))
          For vIndex As Integer = vOrigOPHColl.Count - 1 To 0 Step -1
            vOPH = vOrigOPHColl.Item(vIndex)
            vOPS = New OrderPaymentSchedule
            vOPS.Init(mvEnv, IntegerValue(vOPH.ScheduledPaymentNumber))
            If vOPS.Existing Then
              Dim vOPSAmount As Double = vOPS.AmountDue
              If vOPSAmount > vPayAmount Then
                vOPSAmount = vPayAmount
              End If
              vPayAmount = FixTwoPlaces(vPayAmount - vOPSAmount)
              With vOPS
                .Update(.DueDate, vOPSAmount, 0, .ExpectedBalance, .ClaimDate, .RevisedAmount, OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments)
                .Save(mvEnv.User.UserID, True)
              End With
            End If
          Next
        End If

        'We now need to deal with all the over-payments
        Dim vOverPaymentAmount As Double = pOverPaymentAmount

        'Handle any Retains / Refunds first
        If mvCMTRefundDetails Is Nothing Then mvCMTRefundDetails = New CollectionList(Of CMTExcessPaymentDetail) 'Just in case!!
        'Add Refund products
        For Each vCMTExcessPayment As CMTExcessPaymentDetail In mvCMTRefundDetails
          If vCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.Refund AndAlso vOverPaymentAmount > 0 Then
            CreatePositiveBTAForOverPaidAmount(vCMTExcessPayment, vBT, vOverPaymentAmount, vCompanyControl, pSourceCode, 0)
            vOverPaymentAmount = FixTwoPlaces(vOverPaymentAmount - vCMTExcessPayment.ExcessPaymentAmount)
            If vOverPaymentAmount < 0 Then vOverPaymentAmount = 0
            If vOverPaymentAmount = 0 Then Exit For
          End If
        Next
        'Add Retain products
        If vOverPaymentAmount > 0 Then
          For Each vCMTExcessPayment As CMTExcessPaymentDetail In mvCMTRefundDetails
            If vCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.Retain AndAlso vOverPaymentAmount > 0 Then
              CreatePositiveBTAForOverPaidAmount(vCMTExcessPayment, vBT, vOverPaymentAmount, vCompanyControl, pSourceCode, 0)
              vOverPaymentAmount = FixTwoPlaces(vOverPaymentAmount - vCMTExcessPayment.ExcessPaymentAmount)
              If vOverPaymentAmount < 0 Then vOverPaymentAmount = 0
              If vOverPaymentAmount = 0 Then Exit For
            End If
          Next
        End If

        'Now deal with any in-advance payments
        If vOverPaymentAmount > 0 Then
          'Carry-forwards first
          For Each vCMTExcessPayment As CMTExcessPaymentDetail In mvCMTInAdvanceDetails
            If vCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.CarryForward AndAlso vOverPaymentAmount > 0 Then
              CreatePositiveBTAForOverPaidAmount(vCMTExcessPayment, vBT, vOverPaymentAmount, vCompanyControl, pSourceCode, 0)
              vOverPaymentAmount = FixTwoPlaces(vOverPaymentAmount - vCMTExcessPayment.ExcessPaymentAmount)
              If vOverPaymentAmount < 0 Then vOverPaymentAmount = 0
            End If
          Next
          'Finally any re-analysis
          For Each vCMTExcessPayment As CMTExcessPaymentDetail In mvCMTInAdvanceDetails
            If vCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse AndAlso vOverPaymentAmount > 0 Then
              CreatePositiveBTAForOverPaidAmount(vCMTExcessPayment, vBT, vOverPaymentAmount, vCompanyControl, pSourceCode, 0)
              vOverPaymentAmount = FixTwoPlaces(vOverPaymentAmount - vCMTExcessPayment.ExcessPaymentAmount)
              If vOverPaymentAmount < 0 Then vOverPaymentAmount = 0
            End If
          Next
        End If

        'Save everything
        vBT.SaveChanges()
        With vBatch
          .NumberOfEntries = 0    'Force to be zero so that the totals get updated
          .SetBatchTotals()
          .SetDetailComplete(Nothing, False)
          .SetBatchPosted(True, mvEnv.User.UserID)
          .Save(mvEnv.User.UserID, True)
          .UnLockBatch()
        End With
        vFH = New FinancialHistory
        With vFH
          .Init(mvEnv)
          .CreateFromBatchTransaction(vBT)
          .Save()
        End With
      End If

    End Sub

    Private Sub CreatePositiveBTAForOverPaidAmount(ByVal pCMTExcessPayment As CMTExcessPaymentDetail, ByVal pBT As BatchTransaction, ByVal pOverPaymentAmount As Double, ByVal pCompanyControl As CompanyControl, ByVal pSourceCode As String, ByVal pOrigOPSNumber As Integer)
      'Create a positive BTA for the over-payment
      Dim vAdjAmount As Double = pCMTExcessPayment.ExcessPaymentAmount
      If vAdjAmount > pOverPaymentAmount Then vAdjAmount = pOverPaymentAmount

      Dim vProduct As New Product(mvEnv)
      If pCMTExcessPayment.AdjustmentProductCode.Length > 0 Then
        vProduct.Init(pCMTExcessPayment.AdjustmentProductCode)
      Else
        vProduct.Init()
      End If
      Dim vRateCode As String = pCMTExcessPayment.AdjustmentRateCode
      If vProduct.Existing = False Then
        If pCompanyControl.InAdvanceProduct Is Nothing Then
          RaiseError(DataAccessErrors.daeCompanyControlsMissingInAdvanceProduct)
        Else
          vProduct = pCompanyControl.InAdvanceProduct
          vRateCode = pCompanyControl.InAdvanceRate
        End If
      End If

      Dim vPayPlanPayment As Boolean = False
      Dim vLineType As String = "P"
      If pCMTExcessPayment.AdjustmentProductCode.Length = 0 OrElse pOrigOPSNumber > 0 Then
        vLineType = "O"
        vPayPlanPayment = True
      End If

      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      vBTA.Init()
      Dim vParams As New CDBParameters()
      If pOrigOPSNumber > 0 AndAlso pBT.NextLineNumber > 1 Then
        'Add to previous BTA
        vBTA.Init(pBT.BatchNumber, pBT.TransactionNumber, 1)
      End If
      If vBTA.Existing Then
        With vParams
          .Add("Amount", FixTwoPlaces(vBTA.Amount + vAdjAmount))
          .Add("CurrencyAmount", FixTwoPlaces(vBTA.Amount + vAdjAmount))
        End With
        With pBT
          .Amount = FixTwoPlaces(.Amount + vAdjAmount)
          .CurrencyAmount = FixTwoPlaces(.CurrencyAmount + vAdjAmount)
          .LineTotal = FixTwoPlaces(.LineTotal + vAdjAmount)
          '.NextLineNumber = .NextLineNumber + 1
        End With
      Else
        With vParams
          .Add("LineType", vLineType)
          If pCMTExcessPayment.AdjustmentProductCode.Length > 0 AndAlso vPayPlanPayment = False Then
            .Add("Product", vProduct.ProductCode)
            .Add("Rate", vRateCode)
            .Add("VatRate", mvEnv.VATRate(vProduct.ProductVatCategory, Payer.VATCategory).VatRateCode)
            .Add("VatAmount", CalculateVATAmount(vAdjAmount, (mvEnv.VATRate(vProduct.ProductVatCategory, Payer.VATCategory).Percentage)))
            .Add("CurrencyVatAmount", .Item("VatAmount").DoubleValue)
          Else
            .Add("OrderNumber", PlanNumber)
          End If
          .Add("Amount", vAdjAmount)
          .Add("CurrencyAmount", vAdjAmount)
          .Add("Quantity", If(pCMTExcessPayment.AdjustmentProductCode.Length > 0, 1, 0))
          .Add("Source", pSourceCode)
        End With
        vBTA.InitFromTransaction(pBT)
      End If
      With vBTA
        .Update(vParams)    'Use Update instead of Create as Batch & Transaction numbers have already been set
        .Save(mvEnv.User.UserID)
      End With

      'Create positive FH & FHD for the over-payment
      With vParams
        .Clear()
        .Add("BatchNumber", vBTA.BatchNumber)
        .Add("TransactionNumber", vBTA.TransactionNumber)
        .Add("LineNumber", vBTA.LineNumber)
        .Add("Amount", vAdjAmount)
        .Add("Product", vProduct.ProductCode)
        .Add("Rate", vRateCode)
        .Add("Source", pSourceCode)
        .Add("Quantity", vBTA.Quantity)
        If vBTA.VatRate.Length > 0 Then
          .Add("VatRate", vBTA.VatRate)
        Else
          .Add("VatRate", mvEnv.VATRate(vProduct.ProductVatCategory, Payer.VATCategory).VatRateCode)
        End If
        .Add("VatAmount", CalculateVATAmount(vAdjAmount, (mvEnv.VATRate(vProduct.ProductVatCategory, Payer.VATCategory).Percentage)))
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
          .Add("CurrencyAmount", vAdjAmount)
          .Add("CurrencyVatAmount", .Item("VatAmount").DoubleValue)
        End If
        .Add("InvoicePayment", "N")
      End With
      Dim vFHD As New FinancialHistoryDetail(mvEnv)
      vFHD.Create(vParams)
      vFHD.Save(mvEnv.User.UserID, True)
      vBTA.Save() 'Force a save after the FHD line is created so that it does any post-posting processing.  This is normally done when the batch is posted.

      If pCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.CarryForward OrElse pCMTExcessPayment.ExcessPaymentType = CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse Then
        'Add OrderPaymenthistory
        Dim vCreateInAdvancePayment As Boolean = (pOrigOPSNumber = 0)   'If pOrigOPSNumber >0 then we are re-allocating the original payment back to the payment plan
        'Dim vOPSNumber As Integer = pOrigOPSNumber
        Dim vOPS As New OrderPaymentSchedule
        vOPS.Init(mvEnv, pOrigOPSNumber)
        If vOPS.Existing = False Then
          vOPS.Init(mvEnv)
          vOPS.CreateInAdvance(mvEnv, Me, pOverPaymentAmount)
          If vCreateInAdvancePayment AndAlso Balance > 0 AndAlso pOverPaymentAmount > 0 Then
            'We have an outstanding Balance
            If vOPS.ScheduleCreationReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance AndAlso InAdvance = 0 AndAlso CDate(vOPS.DueDate) = CDate(RenewalPeriodEnd) AndAlso vOPS.AmountDue = vOPS.AmountOutstanding Then
              'This provisional OPS record will be updated in order to accomodate the remaining balance
              With vOPS
                .Update(TodaysDate, .AmountDue, .AmountOutstanding, .ExpectedBalance, pCreationReason:=OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance)
                .Save(mvEnv.User.UserID, True)
              End With
              mvScheduledPayments = Nothing
              vOPS = New OrderPaymentSchedule
              vOPS.Init(mvEnv)
              vOPS.CreateInAdvance(mvEnv, Me, pOverPaymentAmount)
            End If
          End If
          If pOverPaymentAmount > vOPS.AmountOutstanding Then
            'Over-payment is more than AmountOutstanding so update the AmountDue & AmountOutstanding by the difference
            With vOPS
              .Update(.DueDate, FixTwoPlaces(.AmountDue + (pOverPaymentAmount - .AmountOutstanding)), FixTwoPlaces(.AmountOutstanding + (pOverPaymentAmount - .AmountOutstanding)), .ExpectedBalance)
            End With
          End If
        End If

        Dim vNewOPH As New OrderPaymentHistory
        vNewOPH.Init(mvEnv)
        If pOrigOPSNumber > 0 AndAlso pBT.NextLineNumber > 1 Then
          Dim vWhereFields As New CDBFields(New CDBField("batch_number", vBTA.BatchNumber))
          vWhereFields.Add("transaction_number", vBTA.TransactionNumber)
          vWhereFields.Add("line_number", vBTA.LineNumber)
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vNewOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll), "order_payment_history oph", vWhereFields)
          Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          If vRS.Fetch Then vNewOPH.InitFromRecordSet(mvEnv, vRS, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          vRS.CloseRecordSet()
        End If
        If vNewOPH.Existing Then
          vNewOPH.IncreasePaymentAmount(vAdjAmount)
        Else
          PaymentNumber = PaymentNumber + 1   'Increment payment number for OPH
          With vBTA
            vNewOPH.SetValues(.BatchNumber, .TransactionNumber, PaymentNumber, PlanNumber, vAdjAmount, .LineNumber, 0, vOPS.ScheduledPaymentNumber, True)
            If vCreateInAdvancePayment Then
              vNewOPH.Status = "I"
              mvClassFields.Item(PaymentPlanFields.ofInAdvance).DoubleValue = FixTwoPlaces(InAdvance + vAdjAmount)
            End If
          End With
        End If
        vNewOPH.Save()

        vOPS.AddPayment(vAdjAmount)
        vOPS.ProcessPayment()
        vOPS.Save(mvEnv.User.UserID, True)
      End If

    End Sub

    Public Sub SetCMTDetailBalances(ByRef pDetailTypes As PaymentPlanDetail.PaymentPlanDetailTypes, ByVal pBalance As Double)
      'Set new Balance on DetailLines
      Dim vDetail As PaymentPlanDetail = Nothing
      Dim vLinePrice As Double
      Dim vRemaining As Double

      vRemaining = pBalance

      'Apply pOldPrice from last to first
      vDetail = CType(mvEnv.GetPreviousItem(Details, vDetail), PaymentPlanDetail)
      While Not (vDetail Is Nothing)
        With vDetail
          If (.DetailType And pDetailTypes) > 0 Then
            vLinePrice = .Balance
            If vLinePrice >= vRemaining Then
              .Balance = vRemaining
              vRemaining = 0
            Else
              'Keep current Balance
              vRemaining = FixTwoPlaces(vRemaining - vLinePrice)
            End If
          End If
        End With
        vDetail = CType(mvEnv.GetPreviousItem(Details, vDetail), PaymentPlanDetail)
      End While

    End Sub

    ''' <summary>Can the PaymentPlan Balance be prorated by CMT?</summary>
    ''' <param name="pNewMembershipType">The MembershipType to change to.</param>
    Friend ReadOnly Property CanProrateCMTBalance(ByVal pNewMembershipType As MembershipType) As Boolean
      Get
        If ((MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfMonthlyTerm OrElse MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfWeeklyTerm) _
        OrElse (pNewMembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfMonthlyTerm OrElse pNewMembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfWeeklyTerm) _
        OrElse (CMTProportionBalance = CMTProportionBalanceTypes.cmtNone)) Then
          Return False
        Else
          Return True
        End If
      End Get
    End Property

    ''' <summary>Can this PaymentPlan go through an Advanced CMT?</summary>
    ''' <param name="pNewMembershipType">The MembershipType to change to.</param>
    Friend ReadOnly Property CanUseAdvancedCMT(ByVal pNewMembershipType As MembershipType) As Boolean
      Get
        Dim vAdvancedCMT As Boolean = mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlAdvancedCMT)
        If vAdvancedCMT Then
          If CanProrateCMTBalance(pNewMembershipType) = False Then vAdvancedCMT = False
          If IsMultipleCMT Then vAdvancedCMT = False
        End If
        Return vAdvancedCMT
      End Get
    End Property

    ''' <summary>Gets the payment schedule creation date for Change Membership Type (CMT).</summary>
    ''' <returns>Payment schedule creation date.</returns>
    ''' <remarks>When running in multiple-CMT mode the payment schedule creation date will differ
    ''' depending upon whether this is the first or second CMT.</remarks>
    Private Function GetCMTOPSCreationDate() As String
      Dim vOPSDate As Nullable(Of Date)
      If CMTProportionBalance <> CMTProportionBalanceTypes.cmtNone Then
        Dim vRenewalDate As Date = CDate(RenewalPeriodEnd)
        If mvSecondCMT = True Then
          vOPSDate = CDate(CalculateRenewalDate(vRenewalDate.ToString(CAREDateFormat), False))
        Else
          Dim vTotalPayments As Integer = 0
          Dim vMonths As Integer = 0
          If CMTProportionBalance = CMTProportionBalanceTypes.cmtFrequencyAmounts Then
            Dim vPFCode As String = PaymentFrequencyCode  'Hold the code in a variable as passing the property value to mvEnv.GetPaymentFrequency calls the property set and re-calculates the frequency amount!! 
            Dim vPaymentFequency As PaymentFrequency = mvEnv.GetPaymentFrequency(vPFCode)
            Dim vPaymentsRemaining As Integer
            Dim vChangeDate As Date = Today
            Dim vMinChangeDate As Date = GetNextInstalmentDueDate(vChangeDate)
            Dim vMinDueDate As Nullable(Of Date)
            If vMinChangeDate > vChangeDate Then vMinDueDate = vMinChangeDate
            vTotalPayments = GetNumberOfFrequencyAmounts(vPaymentFequency, vRenewalDate.ToString(CAREDateFormat), vPaymentsRemaining, vChangeDate, vMinDueDate)
            If vPaymentsRemaining > 0 Then vMonths = GetMonthsRemaining(vPaymentFequency, vTotalPayments, vPaymentsRemaining)
            If vPaymentFequency.OffsetMonths > 0 _
            OrElse vPaymentFequency.Frequency * vPaymentFequency.Interval < 12 And vPaymentFequency.Period = PaymentFrequency.PaymentFrequencyPeriods.pfpMonths Then
              Dim vOffset As Integer = vPaymentFequency.GetCalculatedOffsetMonths()
              If vOffset <> 0 Then vRenewalDate = vRenewalDate.AddMonths(vOffset)
            End If
          Else
            vTotalPayments = Term
            If TermUnits = OrderTermUnits.otuNone AndAlso vTotalPayments > 0 Then
              vTotalPayments = vTotalPayments * 12
            Else
              vTotalPayments = Math.Abs(vTotalPayments)
            End If
            vMonths = GetProRataMonthsRemaining(vTotalPayments, If(IsDate(Member.CMTDate), CDate(Member.CMTDate), Today), CDate(vRenewalDate))
          End If
          If vMonths > 0 Then
            vOPSDate = vRenewalDate.AddMonths((vMonths * -1))
          ElseIf mvFirstCMT Then
            vOPSDate = vRenewalDate
          End If
        End If
      End If

      If vOPSDate.HasValue = False Then vOPSDate = If(IsDate(Member.CMTDate), CDate(Member.CMTDate), Today)

      Return vOPSDate.Value.ToString(CAREDateFormat)

    End Function

    ''' <summary>Used by Change Membership Type to calculate how much was written off during the specified period.</summary>
    ''' <param name="pStartDate">Renewal period start date</param>
    ''' <param name="pRenewalDate">Renewal period end date</param>
    ''' <returns>Amount written off during the specified period.</returns>
    Private Function SumWriteOffAmounts(ByVal pStartDate As String, ByVal pRenewalDate As String) As Double
      Dim vWriteOff As Double = 0
      Dim vAttrs As String = "ops.scheduled_payment_number, (amount_due - " & mvEnv.Connection.DBIsNull("SUM(oph.amount)", "0") & ") AS wo_amount"

      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.AddLeftOuterJoin("order_payment_history oph", "ops.order_number", "oph.order_number", "ops.scheduled_payment_number", "oph.scheduled_payment_number")

      Dim vOPS As New OrderPaymentSchedule()
      Dim vWhereFields As New CDBFields(New CDBField("ops.order_number", PlanNumber))
      With vWhereFields
        .Add("due_date", CDBField.FieldTypes.cftDate, pStartDate, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoBetweenFrom)
        .Add("due_date#2", CDBField.FieldTypes.cftDate, pRenewalDate, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoBetweenTo)
        .Add("scheduled_payment_status", vOPS.SetPaymentStatus(OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsWrittenOff))
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "order_payment_schedule ops", vWhereFields, "", vAnsiJoins)
      vSQLStatement.GroupBy = "ops.scheduled_payment_number, amount_due"

      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRS.Fetch
        If vRS.Fields(2).DoubleValue > 0 Then vWriteOff += vRS.Fields(2).DoubleValue
      End While
      vRS.CloseRecordSet()
      vWriteOff = FixTwoPlaces(vWriteOff)

      Return vWriteOff

    End Function

    ''' <summary>Can the change of membership type for a Member allow the Direct Debit payer to change?</summary>
    ''' <param name="pMember">The <see cref="Member">Member</see> having the membership type changed.</param>
    ''' <param name="pCancellationReasonCode">The cancellation reason code for the change of membership type.</param>
    ''' <param name="pNewMembershipType">The new <see cref="MembershipType">MembershipType</see> for the chosen Member.</param>
    ''' <param name="pNewDDPayer">The <see cref="Contact">Contact</see> that will become the new Direct Debit payer. This will either be set to a valid <see cref="Contact">Contact</see> or Nothing.</param>
    ''' <returns>True if the Direct Debit payer can be changed, otherwise False.</returns>
    ''' <remarks>This is a subset of <see cref="PaymentPlan.CanChangeMemberDDPayer">CanChangeMemberDDPayer</see> to perform the additional checks required by Change Membership Type.</remarks>
    Public Function CanChangeCMTDDPayer(ByVal pMember As Member, ByVal pCancellationReasonCode As String, ByVal pNewMembershipType As MembershipType, ByRef pNewDDPayer As Contact) As Boolean
      Dim vCanChange As Boolean = False

      pNewDDPayer = Nothing
      Dim vOrigMembershipType As MembershipType = pMember.MembershipType
      If pMember.OriginalMembershipTypeCode.Length > 0 AndAlso pMember.MembershipTypeCode.Equals(pMember.OriginalMembershipTypeCode, System.StringComparison.CurrentCultureIgnoreCase) = False Then
        vOrigMembershipType = mvEnv.MembershipType(pMember.OriginalMembershipTypeCode)
      End If
      If DirectDebitStatus = ppYesNoCancel.ppYes AndAlso vOrigMembershipType.MembersPerOrder = 2 AndAlso pCancellationReasonCode.Length > 0 AndAlso pNewMembershipType.MembersPerOrder = 1 Then
        'Joint membership paid by DD moving to individual membership
        '(1) Find member to be removed from membership
        If CurrentMembers.Count = 0 Then LoadMembers()
        If CurrentMembers.Count > 1 Then
          Dim vOtherMember As Member = Nothing
          If vOrigMembershipType.IsAssociateType() = False AndAlso pNewMembershipType.IsAssociateType() = False Then
            Dim vFound As Boolean = False
            For Each vOtherMember In CurrentMembers
              If pMember.ContactNumber <> vOtherMember.ContactNumber AndAlso pMember.MembershipTypeCode.Equals(vOtherMember.MembershipTypeCode, System.StringComparison.CurrentCultureIgnoreCase) Then
                If vOtherMember.CancellationReason.Length = 0 Then vFound = True
              End If
              If vFound Then Exit For
            Next
            If vFound = False Then
              vOtherMember = New Member()
              vOtherMember.Init(mvEnv)
            End If
          End If
          If vOtherMember IsNot Nothing AndAlso vOtherMember.Existing Then
            '(2) Check the DD payer is either the member to be removed or a joint contact
            Dim vGotContact As Boolean = False
            Dim vDDPayer As Contact = DirectDebit.Payer
            If vDDPayer.ContactNumber = vOtherMember.ContactNumber Then
              vGotContact = True
            ElseIf vDDPayer.ContactType = Contact.ContactTypes.ctcJoint Then
              Dim vJointLinks As CollectionList(Of ContactLink) = vDDPayer.GetJointLinks()
              For Each vLink As ContactLink In vJointLinks
                If vLink.ContactNumber1 = vDDPayer.ContactNumber AndAlso vLink.ContactNumber2 = vOtherMember.ContactNumber Then
                  vGotContact = True
                End If
                If vGotContact Then Exit For
              Next
            End If
            '(3) Check cancellation reason is Move DD Member Cancel Reason from Membership Controls table
            If vGotContact = True AndAlso pCancellationReasonCode.Equals(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMoveDDMemberCancelReason), System.StringComparison.CurrentCultureIgnoreCase) Then
              vCanChange = True
              pNewDDPayer = pMember.Contact
            End If
          End If
        End If
      End If

      Return vCanChange

    End Function

  End Class

End Namespace

