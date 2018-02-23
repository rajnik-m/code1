

Namespace Access
  Public Class TraderPaymentPlanDetails
    Implements System.Collections.IEnumerable

    'This is a collection of PaymentPlanDetail objects
    Private mvCol As New Collection

    Friend Function Add(ByRef pIndexKey As String) As PaymentPlanDetail
      'Create a new object
      Dim vPPD As New PaymentPlanDetail

      mvCol.Add(vPPD, pIndexKey)
      Add = vPPD
      vPPD = Nothing
    End Function

    Public Sub AddItem(ByVal pPPD As PaymentPlanDetail, ByVal pIndexKey As String)
      mvCol.Add(pPPD, pIndexKey)
    End Sub

    Default Public ReadOnly Property Item(ByVal pIndexKey As String) As PaymentPlanDetail
      Get
        Item = CType(mvCol.Item(pIndexKey), PaymentPlanDetail)
      End Get
    End Property
    Default Public ReadOnly Property Item(ByVal pIndexKey As Integer) As PaymentPlanDetail
      Get
        Item = CType(mvCol.Item(pIndexKey), PaymentPlanDetail)
      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public Function Exists(ByVal pIndexKey As String) As Boolean
      Return mvCol.Contains(pIndexKey)
    End Function

    Public Sub AddDetailLinesFromPaymentPlan(ByVal pPP As PaymentPlan, Optional ByVal pSetLineNumbers As Boolean = False)
      'Add existing PaymentPlanDetails to the collection
      Dim vPPD As PaymentPlanDetail
      Dim vKey As Integer
      Dim vLineNumber As Integer

      vKey = mvCol.Count()
      If pSetLineNumbers Then vLineNumber = CType(mvCol.Item(CStr(vKey)), PaymentPlanDetail).LineNumber
      For Each vPPD In pPP.Details
        If pSetLineNumbers Then
          vLineNumber = vLineNumber + 1
          vPPD.LineNumber = vLineNumber
        End If
        vKey = vKey + 1
        AddItem(vPPD, CStr(vKey))
      Next vPPD

    End Sub

    Public Sub AddDetailLinesToPaymentPlan(ByVal pEnv As CDBEnvironment, ByVal pPP As PaymentPlan, ByVal pUpdateType As PaymentPlan.PaymentPlanUpdateTypes, Optional ByVal pSubsStartOnMemberJoined As Boolean = False, Optional ByVal pMemberJoinedDate As String = "")
      'This will add the PaymentPlanDetail lines to the PaymentPlan
      'Each detail will be added to a new Collection in the order they are to be updated/created
      Dim vPPD As PaymentPlanDetail
      Dim vDetails As New Collection
      Dim vDetailNumber As Integer

      '-----------------------------------------------
      'Add existing PaymentPlanDetails (last to first)
      '-----------------------------------------------
      For vDetailNumber = mvCol.Count() To 1 Step -1
        vPPD = Item(vDetailNumber)
        If vPPD.Existing And vPPD.DetailNumber > 0 Then
          If vPPD.DetailNumber <> vDetailNumber Then vPPD.SetPaymentPlanAndDetailNumbers(pPP.PlanNumber, vDetailNumber)
          ProcessDetail(pEnv, pPP, vPPD, pUpdateType, pSubsStartOnMemberJoined, pMemberJoinedDate)
          vDetails.Add(vPPD, CStr(vDetailNumber))
        End If
      Next

      '-----------------------------------------------
      'Add new PaymentPlanDetails (first to last)
      '-----------------------------------------------
      vDetailNumber = 0
      For Each vPPD In mvCol
        vDetailNumber = vDetailNumber + 1
        If Not (vPPD.Existing And vPPD.DetailNumber > 0) Then
          vPPD.SetPaymentPlanAndDetailNumbers(pPP.PlanNumber, vDetailNumber)
          ProcessDetail(pEnv, pPP, vPPD, pUpdateType, pSubsStartOnMemberJoined, pMemberJoinedDate)
          vDetails.Add(vPPD, CStr(vDetailNumber))
        End If
      Next vPPD

      '-----------------------------------------------
      'Now add the details to the Payment Plan
      '-----------------------------------------------
      If mvCol.Count() > 0 Then pPP.AddDetails(vDetails)

    End Sub

    Public Sub SetCMTProRataBalances(ByVal pBalance As Double, ByVal pMonthsRemaining As Integer)
      'Used by CMT to set pro-rata balances
      Dim vPPD As PaymentPlanDetail
      Dim vDiff As Double
      Dim vPPDBalance As Double

      For Each vPPD In mvCol
        With vPPD
          .SetImportBalance(FixTwoPlaces((.Balance / 12) * pMonthsRemaining))
          vPPDBalance = FixTwoPlaces(vPPDBalance + .Balance)
        End With
      Next vPPD

      Dim vOldPPDBalance As Double = vPPDBalance
      If vPPDBalance <> pBalance Then
        'Rounding errors could cause a difference
        Do
          For Each vPPD In mvCol
            With vPPD
              vDiff = 0
              If vPPDBalance > pBalance Then
                vDiff = -0.01
              Else
                vDiff = 0.01
              End If
              If .Balance > 0 Then
                'If PPDetail is free then do not update the Balance
                .SetImportBalance(FixTwoPlaces(.Balance + vDiff))
                vPPDBalance = FixTwoPlaces(vPPDBalance + vDiff)
              End If
            End With
            If vPPDBalance = pBalance Then Exit For
          Next vPPD
        Loop While vPPDBalance <> pBalance AndAlso vPPDBalance <> vOldPPDBalance  'BR15002: Exit loop if we have not changed the vPPDBalance because none of the PPD has got a balance
      End If

    End Sub

    Private Sub ProcessDetail(ByVal pEnv As CDBEnvironment, ByVal pPP As PaymentPlan, ByVal pPPD As PaymentPlanDetail, ByVal pUpdateType As PaymentPlan.PaymentPlanUpdateTypes, Optional ByVal pSubsStartOnMemberJoined As Boolean = False, Optional ByVal pMemberJoinedDate As String = "")
      Dim vPPConversion As Boolean
      Dim vPPMaintenance As Boolean

      If (pUpdateType And PaymentPlan.PaymentPlanUpdateTypes.pputConversion) = PaymentPlan.PaymentPlanUpdateTypes.pputConversion Then
        vPPConversion = True
      ElseIf (pUpdateType And (PaymentPlan.PaymentPlanUpdateTypes.pputAddCreditCardAuthority + PaymentPlan.PaymentPlanUpdateTypes.pputAddDirectDebit + PaymentPlan.PaymentPlanUpdateTypes.pputAddStandingOrder)) > 0 Then
        vPPConversion = True
      ElseIf (pUpdateType And PaymentPlan.PaymentPlanUpdateTypes.pputPaymentPlan) = PaymentPlan.PaymentPlanUpdateTypes.pputPaymentPlan Then
        vPPMaintenance = True
      End If
      Dim vCMT As Boolean
      If (pUpdateType And PaymentPlan.PaymentPlanUpdateTypes.pputChangeMembershipType) = PaymentPlan.PaymentPlanUpdateTypes.pputChangeMembershipType Then vCMT = True
      'Process any Initial Period Incentives
      If pPPD.SpecialInitialPeriodIncentive Then
        Dim vIgnore As Boolean
        Dim vOldTerm As String = pPP.Term.ToString
        With pPP
          .Term = CInt((pPPD.Quantity * -1))
          pPPD.ResetIncentiveQuantity() 'Reset Quantity back to 1
          If .Existing Then
            If pPPD.IncentiveIgnoreProductAndRate AndAlso vPPConversion = True Then vIgnore = True
            .AdvanceRenewalDate(PaymentPlan.AdvanceRenewalDateTypes.ardtAutomatic, pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAutomaticRenewalDateChangeReason), System.Math.Abs(.Term), vCMT)
          End If
          If IntegerValue(vOldTerm) <> .Term Then 'the Payment Plan term has changed
            If .RenewalDateChangedOn <> TodaysDate() AndAlso .RenewalPending = False Then
              'Money already paid for the year then Renewal Date needs amending too
              ApplyIncentive(pPP, vOldTerm)
            End If
            If .RenewalDateChangedOn = TodaysDate() AndAlso vCMT = True Then
              'we're doing a CMT
              If .OldMembershipIncentive AndAlso .RenewalPending = False Then
                'Money already paid for the year but there was an incentive on the plan we're changing from so change Renewal Date to include incentive
                ApplyIncentive(pPP, vOldTerm)
              End If
            End If
          End If
        End With

        If vIgnore Then
          'If the chosen optional incentive was an I-type incentive where the Ignore flag is set
          'Need to reset PPD back to original state
          If pPP.DetailExists((pPPD.DetailNumber)) Then
            pPPD = CType(pPP.Details.Item(pPP.GetDetailKeyFromLineNo(pPPD.DetailNumber)), PaymentPlanDetail)
          End If
        End If
      End If

      'Process any Subscription valid from/to dates
      pPPD.ProcessSubscriptionValidDates(pPP, vPPMaintenance, vCMT, pSubsStartOnMemberJoined, pMemberJoinedDate)

    End Sub

    Private Sub ApplyIncentive(ByVal pPP As PaymentPlan, ByVal pOldTerm As String)
      'updates the renewal date to apply incentive to the current payment plan
      Dim vDate As String
      With pPP
        vDate = .RenewalDate
        'Take off old Term
        If IntegerValue(pOldTerm) < 0 Then
          vDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, IntegerValue(pOldTerm), CDate(vDate)))
        Else
          vDate = CDate(vDate).AddYears(-IntegerValue(pOldTerm)).ToString(CAREDateFormat)
        End If
        'Add on the new Term
        If .Term < 0 Then
          vDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(.Term), CDate(vDate)))
        Else
          vDate = CDate(vDate).AddYears(.Term).ToString(CAREDateFormat)
        End If
        If .NextPaymentDue = .RenewalDate Then .NextPaymentDue = vDate
        .RenewalDate = vDate
      End With
    End Sub

    Public Sub Remove(ByRef pIndexKey As String)
      mvCol.Remove(pIndexKey)
    End Sub

    Public Sub Clear()
      mvCol = New Collection
    End Sub

    Public Function Items() As Collection
      Return mvCol
    End Function

  End Class
End Namespace
