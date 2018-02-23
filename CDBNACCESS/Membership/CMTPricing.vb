Imports CARE.Utilities.Common
''' <summary>Used by Change Membership Type (CMT) to calculate the price to be charged for each membership detail line.</summary>
Friend Class CMTPricing

  Friend Enum CMTOldOrNewMemberType
    OldMembershipType
    NewMembershipType
  End Enum

  Private mvEnv As CDBEnvironment
  Private mvPayer As Contact
  Private mvCMTOldNewMemberType As CMTOldOrNewMemberType
  Private mvOldPPDetailLines As Collection
  Private mvNewPPDetailLines As TraderPaymentPlanDetails
  Private mvAdvancedCMT As Boolean

  Private mvCMTDate As Date
  Private mvJoinedDate As Date
  Private mvMembershipType As MembershipType
  Private mvPaymentFrequency As PaymentFrequency

  Private mvMembershipPrice As Double
  Private mvEntitlementPrice As Double
  Private mvOtherLinesPrice As Double

  Private mvProratedMembershipPrice As Double
  Private mvProratedEntitlementPrice As Double
  Private mvProratedOtherLinesPrice As Double

  Private mvMembershipBalance As Double
  Private mvEntitlementBalance As Double
  Private mvOtherLinesBalance As Double

  Private mvMembershipArrears As Double
  Private mvEntitlementArrears As Double
  Private mvOtherLinesArrears As Double

  Private mvMembershipFixedAmount As String = ""

  Private mvOldTermMonths As Integer
  Private mvNewTermMonths As Integer
  Private mvFullTermMonths As Integer

  ''' <summary>Initialise new CMTPricing class with common defaults.</summary>
  ''' <param name="pPayer">Current Membership Payer.</param>
  ''' <param name="pCMTDate">Date CMT and prorating takes effect.</param>
  ''' <param name="pJoinedDate">New membership joined date</param>
  ''' <remarks></remarks>
  Friend Sub New(ByVal pEnv As CDBEnvironment, ByVal pPayer As Contact, ByVal pCMTDate As Date, ByVal pJoinedDate As Date, ByVal pMembershipType As MembershipType, ByVal pPaymentFrequency As PaymentFrequency, ByVal pCMTOldNewType As CMTOldOrNewMemberType)
    mvCMTOldNewMemberType = pCMTOldNewType
    mvEnv = pEnv
    mvPayer = pPayer
    mvCMTDate = pCMTDate
    mvJoinedDate = pJoinedDate
    mvMembershipType = pMembershipType
    mvPaymentFrequency = pPaymentFrequency
    mvOldPPDetailLines = Nothing
    mvNewPPDetailLines = Nothing
  End Sub

  Friend Sub SetupMembershipDetails(ByVal pPaymentPlanDetails As Collection, ByVal pSetProratedPrices As Boolean)
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      mvOldPPDetailLines = pPaymentPlanDetails
      If pSetProratedPrices Then
        'Prices have already been prorated so just get the values
        For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
          With vPPD
            Select Case .DetailType
              Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge
                mvMembershipPrice = FixTwoPlaces(mvMembershipPrice + .GetFullPrice)
                If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                  mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + .GetFullPrice)
                ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                  mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + .GetProratedPrice)
                End If
                mvMembershipBalance = FixTwoPlaces(mvMembershipBalance + .Balance)
                mvMembershipArrears = FixTwoPlaces(mvMembershipArrears + .Arrears)
                If .Amount.Length > 0 Then mvMembershipFixedAmount = FixTwoPlaces(DoubleValue(mvMembershipFixedAmount) + DoubleValue(.Amount)).ToString

              Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement, PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive
                mvEntitlementPrice = FixTwoPlaces(mvEntitlementPrice + .GetFullPrice)
                If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                  mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + .GetFullPrice)
                ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                  mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + .GetProratedPrice)
                End If
                mvEntitlementBalance = FixTwoPlaces(mvEntitlementBalance + .Balance)
                mvEntitlementArrears = FixTwoPlaces(mvEntitlementArrears + .Arrears)

              Case Else   'PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge
                Dim vFullPrice As Double = .GetFullPrice()
                If .HasPriceInfo = True AndAlso .UnitPrice <> 0 AndAlso .GrossAmount <> .UnitPrice Then
                  'Try and handle situations where detail line added with a lower amount for this year
                  vFullPrice = .GrossAmount
                End If
                mvOtherLinesPrice = FixTwoPlaces(mvOtherLinesPrice + vFullPrice)
                mvProratedOtherLinesPrice = FixTwoPlaces(mvProratedOtherLinesPrice + vFullPrice)
                mvOtherLinesBalance = FixTwoPlaces(mvOtherLinesBalance + .Balance)
                mvOtherLinesArrears = FixTwoPlaces(mvOtherLinesArrears + .Arrears)

            End Select
          End With
        Next
        mvAdvancedCMT = True
      End If
    End If
  End Sub
  Friend Sub SetupMembershipDetails(ByVal pPaymentPlanDetails As TraderPaymentPlanDetails, ByVal pSetProratedPrices As Boolean)
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.NewMembershipType Then
      mvNewPPDetailLines = pPaymentPlanDetails
      If pSetProratedPrices Then
        'Prices have already been prorated so just get the values
        For Each vPPD As PaymentPlanDetail In mvNewPPDetailLines
          With vPPD
            Select Case .DetailType
              Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge
                mvMembershipPrice = FixTwoPlaces(mvMembershipPrice + .GetFullPrice)
                If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                  mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + .GetFullPrice)
                ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                  mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + .GetProratedPrice)
                End If
                mvMembershipBalance = FixTwoPlaces(mvMembershipBalance + .Balance)
                If .Amount.Length > 0 Then mvMembershipFixedAmount = FixTwoPlaces(DoubleValue(mvMembershipFixedAmount) + DoubleValue(.Amount)).ToString

              Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement, PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive
                mvEntitlementPrice = FixTwoPlaces(mvEntitlementPrice + .GetFullPrice)
                If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                  mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + .GetFullPrice)
                ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                  mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + .GetProratedPrice)
                End If
                mvEntitlementBalance = FixTwoPlaces(mvEntitlementBalance + .Balance)

              Case Else   'PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge
                mvOtherLinesPrice = FixTwoPlaces(mvOtherLinesPrice + .GetFullPrice)
                mvProratedOtherLinesPrice = FixTwoPlaces(mvProratedOtherLinesPrice + .GetFullPrice())
                mvOtherLinesBalance = FixTwoPlaces(mvOtherLinesBalance + .Balance)

            End Select
          End With
        Next
        mvAdvancedCMT = True
      End If
    End If
  End Sub

  Friend Sub ProrateCosts(ByVal pCMTProrateBasis As PaymentPlan.CMTProportionBalanceTypes, ByVal pNumberOfMonths As Integer, ByVal pFullTermMonths As Integer, ByVal pChargeTermMonths As Integer, ByVal pUseProrateMonths As Boolean, ByVal pAdvancedCMT As Boolean)
    mvAdvancedCMT = pAdvancedCMT
    mvMembershipPrice = GetDetailPrice(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge, mvMembershipBalance, mvMembershipArrears)
    mvEntitlementPrice = GetDetailPrice(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive, mvEntitlementBalance, mvEntitlementArrears)
    mvOtherLinesPrice = GetDetailPrice(PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge Or PaymentPlanDetail.PaymentPlanDetailTypes.ppdltNotSet, mvOtherLinesBalance, mvOtherLinesArrears)
    mvProratedMembershipPrice = 0
    mvProratedEntitlementPrice = 0
    mvProratedOtherLinesPrice = 0

    Dim vNumberOfMonths As Integer = pNumberOfMonths
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType AndAlso pUseProrateMonths Then
      Dim vRemaining As Integer = pFullTermMonths - pNumberOfMonths
      vNumberOfMonths = pChargeTermMonths - vRemaining
      If vNumberOfMonths < 0 Then vNumberOfMonths = 0
    End If

    Select Case pCMTProrateBasis
      Case PaymentPlan.CMTProportionBalanceTypes.cmtMonths, PaymentPlan.CMTProportionBalanceTypes.cmtFrequencyAmounts
        'Prorate according to the number of months
        Dim vFullPrice As Double
        Dim vPrortedPrice As Double

        'Sort out the prices first
        If pCMTProrateBasis = PaymentPlan.CMTProportionBalanceTypes.cmtFrequencyAmounts Then
          'Change these figures to payment frequencies
          pFullTermMonths = CInt(pFullTermMonths / mvPaymentFrequency.Interval)
          vNumberOfMonths = CInt(vNumberOfMonths / mvPaymentFrequency.Interval)
          If mvPaymentFrequency.Frequency * mvPaymentFrequency.Interval < 12 AndAlso mvPaymentFrequency.Period = PaymentFrequency.PaymentFrequencyPeriods.pfpMonths Then
            pFullTermMonths = CInt(mvPaymentFrequency.Frequency * mvPaymentFrequency.Interval)
          End If
        End If
        mvFullTermMonths = pFullTermMonths
        If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
          mvOldTermMonths = vNumberOfMonths
          For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
            vPrortedPrice = vPPD.ProratedPrice(pFullTermMonths, vNumberOfMonths)
          Next
        Else
          mvNewTermMonths = vNumberOfMonths
          For Each vPPD As PaymentPlanDetail In mvNewPPDetailLines
            vPrortedPrice = vPPD.ProratedPrice(pFullTermMonths, vNumberOfMonths)
          Next
        End If

        If mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlAdvancedCMT) = True AndAlso mvAdvancedCMT = True Then
          mvMembershipPrice = 0
          mvEntitlementPrice = 0
          mvOtherLinesPrice = 0
          If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
            For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
              vFullPrice = vPPD.GetFullPrice()
              vPrortedPrice = vPPD.GetProratedPrice()
              With vPPD
                Select Case .DetailType
                  Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge
                    mvMembershipPrice = FixTwoPlaces(mvMembershipPrice + vFullPrice)
                    If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                      mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + vFullPrice)
                    ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                      mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + vPrortedPrice)
                    End If
                    If .Amount.Length > 0 Then mvMembershipFixedAmount = FixTwoPlaces(DoubleValue(mvMembershipFixedAmount) + DoubleValue(.Amount)).ToString

                  Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement, PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive
                    mvEntitlementPrice = FixTwoPlaces(mvEntitlementPrice + vFullPrice)
                    If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                      mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + vFullPrice)
                    ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                      mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + vPrortedPrice)
                    End If

                  Case Else 'PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge
                    'Try and handle situations where detail line added with a lower amount for this year
                    .SetCMTOtherLinePartYearPrice()
                    vFullPrice = .GetFullPrice()
                    vPrortedPrice = .GetProratedPrice()
                    mvOtherLinesPrice = FixTwoPlaces(mvOtherLinesPrice + vFullPrice)
                    mvProratedOtherLinesPrice = FixTwoPlaces(mvProratedOtherLinesPrice + vFullPrice)   'Never prorate
                End Select
              End With
            Next
          Else
            For Each vPPD As PaymentPlanDetail In mvNewPPDetailLines
              vFullPrice = vPPD.GetFullPrice()
              vPrortedPrice = vPPD.GetProratedPrice()
              With vPPD
                Select Case .DetailType
                  Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltCharge
                    mvMembershipPrice = FixTwoPlaces(mvMembershipPrice + vFullPrice)
                    If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                      mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + vFullPrice)
                    ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                      mvProratedMembershipPrice = FixTwoPlaces(mvProratedMembershipPrice + vPrortedPrice)
                    End If
                    If .Amount.Length > 0 Then mvMembershipFixedAmount = FixTwoPlaces(DoubleValue(mvMembershipFixedAmount) + DoubleValue(.Amount)).ToString

                  Case PaymentPlanDetail.PaymentPlanDetailTypes.ppdltEntitlement, PaymentPlanDetail.PaymentPlanDetailTypes.ppdltIncentive
                    mvEntitlementPrice = FixTwoPlaces(mvEntitlementPrice + vFullPrice)
                    If .CMTProrateLineType = MembershipType.CMTProrateCosts.FullCharge Then
                      mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + vFullPrice)
                    ElseIf .CMTProrateLineType = MembershipType.CMTProrateCosts.Prorate Then
                      mvProratedEntitlementPrice = FixTwoPlaces(mvProratedEntitlementPrice + vPrortedPrice)
                    End If

                  Case Else 'PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge
                    mvOtherLinesPrice = FixTwoPlaces(mvOtherLinesPrice + vFullPrice)
                    mvProratedOtherLinesPrice = FixTwoPlaces(mvProratedOtherLinesPrice + vFullPrice)   'Never prorate
                End Select
              End With
            Next
          End If
        Else
          mvProratedMembershipPrice = CalculateProrateAmount(mvMembershipPrice + mvEntitlementPrice, pFullTermMonths, vNumberOfMonths)
          mvProratedEntitlementPrice = 0  'CalculateProrateAmount(mvEntitlementPrice, pFullTermMonths, vNumberOfMonths)
          mvOtherLinesPrice = 0
          If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
            'Try and handle situations where detail line added with a lower amount for this year
            For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
              If (vPPD.DetailType And PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge) = PaymentPlanDetail.PaymentPlanDetailTypes.ppdltOtherCharge Then
                With vPPD
                  .SetCMTOtherLinePartYearPrice()
                  mvOtherLinesPrice = FixTwoPlaces(mvOtherLinesPrice + .GetFullPrice)
                  'mvProratedOtherLinesPrice = FixTwoPlaces(mvProratedOtherLinesPrice + vFullPrice)   'Never prorate
                End With
              End If
            Next
          End If
          mvProratedOtherLinesPrice = mvOtherLinesPrice   'CalculateProrateAmount(mvOtherLinesPrice, pFullTermMonths, vNumberOfMonths)
        End If

      Case Else     'cmtNone
        'No prorating
        mvOldTermMonths = 0
        mvNewTermMonths = pFullTermMonths
    End Select
  End Sub

  Friend Sub ProrateFirstCMTCosts(ByVal pBalance As Double, ByVal pRenewalAmount As Double, ByVal pAmountPaid As Double)
    'Full Price
    mvMembershipPrice = FixTwoPlaces(((pBalance - pRenewalAmount) + pAmountPaid))
    mvEntitlementPrice = 0
    'mvOtherLinesPrice stays the same
    If mvOtherLinesPrice <> 0 Then mvMembershipPrice = FixTwoPlaces((mvMembershipPrice - mvOtherLinesPrice)) 'price calculated above includes other lines
    'Balance
    mvOtherLinesBalance = FixTwoPlaces((mvOtherLinesBalance - mvOtherLinesPrice))
    If mvOtherLinesBalance < 0 Then mvOtherLinesBalance = 0
    mvMembershipBalance = FixTwoPlaces(((pBalance - pRenewalAmount) - mvOtherLinesBalance))
    If mvMembershipBalance < 0 Then mvMembershipBalance = 0
    mvEntitlementBalance = 0
    'Prorated Price
    mvProratedMembershipPrice = CalculateProrateAmount(mvMembershipPrice, mvFullTermMonths, mvOldTermMonths)
    mvProratedEntitlementPrice = 0
    mvProratedOtherLinesPrice = mvOtherLinesPrice
  End Sub

  Friend Sub SetOldTypeDetailBalances(ByVal pDetailType As PaymentPlanDetail.PaymentPlanDetailTypes, ByVal pNewOverallBalance As Double, ByVal pTotalProratedPrice As Double)
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      'Pro-rate from last to first
      Dim vPPD As PaymentPlanDetail
      For vIndex As Integer = mvOldPPDetailLines.Count To 1 Step -1
        vPPD = CType(mvOldPPDetailLines.Item(vIndex), PaymentPlanDetail)
        With vPPD
          If (.DetailType And pDetailType) > 0 Then
            .SetCMTOldTypeBalance(mvAdvancedCMT, pNewOverallBalance)
          End If
        End With
      Next
      If mvAdvancedCMT = False Then
        'Now setup correct pro-rated figures
        Dim vCalcProratedPrice As Double = 0
        For Each vPPD In mvOldPPDetailLines
          With vPPD
            If (.DetailType And pDetailType) > 0 Then
              vCalcProratedPrice = FixTwoPlaces(vCalcProratedPrice + .GetProratedPrice())
            End If
          End With
        Next
        If vCalcProratedPrice > FixTwoPlaces(mvMembershipPrice + 0.12) Then vCalcProratedPrice = mvMembershipPrice 'Multiple CMT's have changed price for the first year (taking account rounding differences)
        If vCalcProratedPrice <> pTotalProratedPrice Then
          Dim vDiff As Double = 0.01
          If vCalcProratedPrice > pTotalProratedPrice Then vDiff = vDiff * -1 'reduce figures
          For Each vPPD In mvOldPPDetailLines
            With vPPD
              If (.DetailType And pDetailType) > 0 Then
                .SetCMTOldMembershipProratedPrice(FixTwoPlaces(.GetProratedPrice() + vDiff))
                vCalcProratedPrice = FixTwoPlaces(vCalcProratedPrice + vDiff)
              End If
            End With
            If vCalcProratedPrice = pTotalProratedPrice Then Exit For
          Next
        End If
      End If
    End If
  End Sub

  Friend Sub WriteOffOldTypeDetailBalances(ByVal pWriteoffAmount As Double)
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      Dim vProportion As Integer = mvPaymentFrequency.Frequency
      Dim vWOProportionally As Boolean = mvEnv.GetConfigOption("fp_pay_proportional_details", False)
      While pWriteoffAmount > 0
        For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
          vPPD.CMTWriteOff(vProportion, vWOProportionally, pWriteoffAmount)
          If pWriteoffAmount = 0 Then Exit For
        Next
        If pWriteoffAmount > 0 Then vWOProportionally = False
      End While
    End If
  End Sub

  Friend Sub SetNewTypeDetailBalances(ByVal pBalance As Double, ByVal pMonthsRemaining As Integer)
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.NewMembershipType Then
      For Each vPPD As PaymentPlanDetail In mvNewPPDetailLines
        With vPPD
          .SetCMTNewTypeProrateBalance(mvAdvancedCMT, pBalance, pMonthsRemaining)
          pBalance = FixTwoPlaces(pBalance - .Balance)
        End With
      Next

      'If only 1 detail line then increase Balance on that line to correct figure
      If pBalance > 0 AndAlso mvNewPPDetailLines.Count = 1 Then
        Dim vPPD As PaymentPlanDetail = mvNewPPDetailLines.Item(1)
        vPPD.SetNewBalanceForCMT(FixTwoPlaces(vPPD.Balance + pBalance), mvAdvancedCMT)
        pBalance = FixTwoPlaces(pBalance - vPPD.Balance)
      End If

      If mvAdvancedCMT = False Then
        Dim vDiff As Double
        While pBalance > 0
          'Rounding errors could cause a difference
          For Each vPPD As PaymentPlanDetail In mvNewPPDetailLines
            If vPPD.Balance <> 0 Then
              If vPPD.Balance >= 0 Then
                vDiff = 0.01
              Else
                vDiff = -0.01   'Discount line
              End If
              vPPD.SetNewBalanceForCMT(FixTwoPlaces(vPPD.Balance + vDiff), mvAdvancedCMT)
              pBalance = FixTwoPlaces(pBalance - vDiff)
            End If
            If pBalance = 0 Then Exit For
          Next
        End While
      End If
    End If
  End Sub

  ''' <summary>Get the balance etc. for all detail lines of the specified line type.</summary>
  Private Function GetDetailPrice(ByVal pDetailTypes As PaymentPlanDetail.PaymentPlanDetailTypes, ByRef pBalance As Double, ByRef pArrears As Double) As Double
    Dim vFullPrice As Double
    Dim vBalance As Double
    Dim vArrears As Double
    Dim vFixedAmount As String = ""

    pBalance = 0
    pArrears = 0
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      For Each vDetail As PaymentPlanDetail In mvOldPPDetailLines
        With vDetail
          If (.DetailType And pDetailTypes) > 0 Then
            If .HasPriceInfo = True AndAlso .UnitPrice = 0 AndAlso .GrossAmount <> 0 AndAlso (.Amount.Length > 0 OrElse .NetFixedAmount.Length > 0) Then
              'In earler versions UnitPrice was incorrectly set to zero instead of the FixedAmount so update it here so that the prices come out right
              .CMTUpdatePreviousPriceData()
            End If
            vBalance += .Balance
            vArrears = +.Arrears
            vFullPrice = FixTwoPlaces(vFullPrice + .FullPrice(mvJoinedDate, mvPayer.VATCategory))
            If .Amount.Length > 0 Then vFixedAmount = FixTwoPlaces(DoubleValue(vFixedAmount) + DoubleValue(.Amount)).ToString
          End If
        End With
      Next
    Else
      For Each vDetail As PaymentPlanDetail In mvNewPPDetailLines
        With vDetail
          If (.DetailType And pDetailTypes) > 0 Then
            vBalance += .Balance
            vArrears = +.Arrears
            vFullPrice = FixTwoPlaces(vFullPrice + .FullPrice(mvJoinedDate, mvPayer.VATCategory))
            If .Amount.Length > 0 Then vFixedAmount = FixTwoPlaces(DoubleValue(vFixedAmount) + DoubleValue(.Amount)).ToString
          End If
        End With
      Next
    End If

    pBalance = vBalance
    pArrears = vArrears

    Return vFullPrice

  End Function

  Friend Sub ReNumberOldDetailKeys()
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      Dim vPPDetails As New Collection
      Dim vCount As Integer
      'Following removal of a Detail Number, ensure keys for mvOldPPDetailLines
      'run sequentially

      For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
        vCount = vCount + 1
        vPPDetails.Add(vPPD, CStr(vCount))
      Next
      mvOldPPDetailLines = vPPDetails
    End If
  End Sub

  Friend Function GetOldDetailKeyFromLineNo(ByVal pLineNumber As Integer) As String
    Dim vCount As Integer
    If mvCMTOldNewMemberType = CMTOldOrNewMemberType.OldMembershipType Then
      For Each vPPD As PaymentPlanDetail In mvOldPPDetailLines
        vCount = vCount + 1
        If vPPD.DetailNumber = pLineNumber Then
          Return vCount.ToString
          Exit For
        End If
      Next
    End If
    Return ""
  End Function

#Region " Properties "

  Friend ReadOnly Property MembershipPrice() As Double
    Get
      Return mvMembershipPrice
    End Get
  End Property

  Friend ReadOnly Property EntitlementPrice() As Double
    Get
      Return mvEntitlementPrice
    End Get
  End Property

  Friend ReadOnly Property OtherDetailsPrice() As Double
    Get
      Return mvOtherLinesPrice
    End Get
  End Property

  Friend ReadOnly Property MembershipFixedAmount() As String
    Get
      Return mvMembershipFixedAmount
    End Get
  End Property

  Friend ReadOnly Property MembershipBalance() As Double
    Get
      Return mvMembershipBalance
    End Get
  End Property

  Friend ReadOnly Property EntitlementBalance() As Double
    Get
      Return mvEntitlementBalance
    End Get
  End Property

  Friend ReadOnly Property OtherDetailsBalance() As Double
    Get
      Return mvOtherLinesBalance
    End Get
  End Property

  Friend ReadOnly Property ProratedMembershipPrice() As Double
    Get
      Return mvProratedMembershipPrice
    End Get
  End Property

  Friend ReadOnly Property ProratedEntitlementPrice() As Double
    Get
      Return mvProratedEntitlementPrice
    End Get
  End Property

  Friend ReadOnly Property ProratedOtherDetailsPrice() As Double
    Get
      Return mvProratedOtherLinesPrice
    End Get
  End Property

  Friend ReadOnly Property MembershipArrears() As Double
    Get
      Return mvMembershipArrears
    End Get
  End Property

  Friend ReadOnly Property EntitlementArrears() As Double
    Get
      Return mvEntitlementArrears
    End Get
  End Property

  Friend ReadOnly Property OtherDetailsArrears() As Double
    Get
      Return mvOtherLinesArrears
    End Get
  End Property

  Friend ReadOnly Property OldDetails() As Collection
    Get
      Return mvOldPPDetailLines
    End Get
  End Property

  Friend ReadOnly Property NewDetails() As TraderPaymentPlanDetails
    Get
      Return mvNewPPDetailLines
    End Get
  End Property

#End Region

End Class


