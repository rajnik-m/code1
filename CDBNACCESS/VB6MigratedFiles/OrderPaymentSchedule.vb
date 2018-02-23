Imports System.Linq

Namespace Access
  Public Class OrderPaymentSchedule
    Implements IDbLoadable, IDbSelectable

    Public Enum OrderPaymentScheduleRecordSetTypes 'These are bit values
      opsrtAll = &HFFFFS
      'ADD additional recordset types here
      opsrtMain = 1
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum OrderPaymentScheduleFields
      opsfAll = 0
      opsfScheduledPaymentNumber
      opsfOrderNumber
      opsfDueDate
      opsfClaimDate
      opsfAmountDue
      opsfAmountOutstanding
      opsfRevisedAmount
      opsfExpectedBalance
      opsfScheduledPaymentStatus
      opsfScheduleCreationReason
      opsfAmendedBy
      opsfAmendedOn
    End Enum

    Public Enum OrderPaymentSchedulePaymentStatus
      opspsDue = 1
      opspsPartPaid
      opspsFullyPaid
      opspsUnprocessedPayment
      opspsArrears
      opspsCancelled
      opspsSkippedPayment
      opspsProvisional
      opspsWrittenOff
      opspsMissedLoanPayment
    End Enum

    Public Enum OrderPaymentScheduleCreationReasons
      opscrInAdvance = 1
      opscrBatchPosting
      opscrAutoPayMethodCancel
      opscrChangeMembershipType
      opscrFinancialAdjustments
      opscrFutureMemberTypeChange
      opscrInitialDataSetup
      opscrNewPaymentPlan
      opscrPaymentPlanMaintenance
      opscrRenewalsReminders
      opscrUserAmendment
      opscrSkippedPayment 'Only used when skipping a provisional payment which will have a creation reason of In-Advance
      opscrAdvanceRenewalDate 'Only used when calling PaymentPlan.RegenerateScheduledPayments from PaymentPlan.AdvanceRenewalDate
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvPaymentAmount As Double
    Private mvLogPPChangesToOPS As Boolean

    Private mvSCCheckValue As Boolean
    Private mvOrderPaymentHistory As List(Of OrderPaymentHistory) 'The OrderPaymentHistories of this schedule. Plase use the Property.

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "order_payment_schedule"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("due_date", CDBField.FieldTypes.cftDate)
          .Add("claim_date", CDBField.FieldTypes.cftDate)
          .Add("amount_due", CDBField.FieldTypes.cftNumeric)
          .Add("amount_outstanding", CDBField.FieldTypes.cftNumeric)
          .Add("revised_amount", CDBField.FieldTypes.cftNumeric)
          .Add("expected_balance", CDBField.FieldTypes.cftNumeric)
          .Add("scheduled_payment_status")
          .Add("schedule_creation_reason")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentScheduleFields.opsfOrderNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedBy).PrefixRequired = True
        mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedOn).PrefixRequired = True

        mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvPaymentAmount = 0
      mvLogPPChangesToOPS = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As OrderPaymentScheduleFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).IntegerValue = 0 Then mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).IntegerValue = mvEnv.GetCachedControlNumber(CDBEnvironment.CachedControlNumberTypes.ccnPaymentSchedule)
      mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    Private Function GetCreationReason(ByVal pCreationReasonCode As String) As OrderPaymentScheduleCreationReasons
      Select Case pCreationReasonCode
        Case "AP"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance
        Case "BP"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrBatchPosting
        Case "CA"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrAutoPayMethodCancel
        Case "CM"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrChangeMembershipType
        Case "FA"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments
        Case "FM"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrFutureMemberTypeChange
        Case "IN"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrInitialDataSetup
        Case "NP"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrNewPaymentPlan
        Case "PM"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrPaymentPlanMaintenance
        Case "RR"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrRenewalsReminders
        Case "US"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrUserAmendment
        Case "SK"
          GetCreationReason = OrderPaymentScheduleCreationReasons.opscrSkippedPayment
      End Select
    End Function
    Friend Function SetCreationReason(ByVal pCreationReason As OrderPaymentScheduleCreationReasons) As String
      Select Case pCreationReason
        Case OrderPaymentScheduleCreationReasons.opscrInAdvance
          Return "AP"
        Case OrderPaymentScheduleCreationReasons.opscrBatchPosting
          Return "BP"
        Case OrderPaymentScheduleCreationReasons.opscrAutoPayMethodCancel
          Return "CA"
        Case OrderPaymentScheduleCreationReasons.opscrChangeMembershipType
          Return "CM"
        Case OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments
          Return "FA"
        Case OrderPaymentScheduleCreationReasons.opscrFutureMemberTypeChange
          Return "FM"
        Case OrderPaymentScheduleCreationReasons.opscrInitialDataSetup
          Return "IN"
        Case OrderPaymentScheduleCreationReasons.opscrNewPaymentPlan
          Return "NP"
        Case OrderPaymentScheduleCreationReasons.opscrPaymentPlanMaintenance
          Return "PM"
        Case OrderPaymentScheduleCreationReasons.opscrRenewalsReminders
          Return "RR"
        Case OrderPaymentScheduleCreationReasons.opscrUserAmendment
          Return "US"
        Case OrderPaymentScheduleCreationReasons.opscrSkippedPayment
          Return "SK"
        Case Else
          Return ""       'Added to fix compiler warning
      End Select
    End Function

    Private Function GetPaymentStatus(ByVal pPaymentStatusCode As String) As OrderPaymentSchedulePaymentStatus
      Select Case pPaymentStatusCode
        Case "A"
          Return OrderPaymentSchedulePaymentStatus.opspsArrears
        Case "C"
          Return OrderPaymentSchedulePaymentStatus.opspsCancelled
        Case "D"
          Return OrderPaymentSchedulePaymentStatus.opspsDue
        Case "F"
          Return OrderPaymentSchedulePaymentStatus.opspsFullyPaid
        Case "M"
          Return OrderPaymentSchedulePaymentStatus.opspsMissedLoanPayment
        Case "P"
          Return OrderPaymentSchedulePaymentStatus.opspsPartPaid
        Case "S"
          Return OrderPaymentSchedulePaymentStatus.opspsSkippedPayment
        Case "U"
          Return OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
        Case "V"
          Return OrderPaymentSchedulePaymentStatus.opspsProvisional
        Case "W"
          Return OrderPaymentSchedulePaymentStatus.opspsWrittenOff
      End Select
    End Function

    Friend Function SetPaymentStatus(ByVal pPaymentStatus As OrderPaymentSchedulePaymentStatus) As String
      Select Case pPaymentStatus
        Case OrderPaymentSchedulePaymentStatus.opspsArrears
          Return "A"
        Case OrderPaymentSchedulePaymentStatus.opspsCancelled
          Return "C"
        Case OrderPaymentSchedulePaymentStatus.opspsDue
          Return "D"
        Case OrderPaymentSchedulePaymentStatus.opspsFullyPaid
          Return "F"
        Case OrderPaymentSchedulePaymentStatus.opspsMissedLoanPayment
          Return "M"
        Case OrderPaymentSchedulePaymentStatus.opspsPartPaid
          Return "P"
        Case OrderPaymentSchedulePaymentStatus.opspsSkippedPayment
          Return "S"
        Case OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
          Return "U"
        Case OrderPaymentSchedulePaymentStatus.opspsProvisional
          Return "V"
        Case OrderPaymentSchedulePaymentStatus.opspsWrittenOff
          Return "W"
        Case Else
          Return ""       'Added to fix compiler warning
      End Select
    End Function

    Private Function ProcessReversal(ByVal pPP As PaymentPlan, ByVal pClaimDate As String, ByRef pUseClaimDate As Boolean) As OrderPaymentSchedule
      'For certain DD/CCCA payments see if the reversed payment can be allocated against a later OPS record
      'pUpdateClaim is used to show that we need to use this ClaimDate even though there was no existing OPS with that date
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vPPD As PaymentPlanDetail
      Dim vDonPP As Boolean
      Dim vFound As Boolean

      pUseClaimDate = False 'Make sure this always starts as False
      If pPP.PlanType <> CDBEnvironment.ppType.pptLoan AndAlso (pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes OrElse pPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) Then
        Select Case mvEnv.GetConfig("fp_arrears_claim_method_create")
          Case "NEXT_CLAIM"
            Select Case pPP.PlanType
              Case CDBEnvironment.ppType.pptCovenant, CDBEnvironment.ppType.pptMember
                'Do nothing
              Case Else
                If pPP.PaymentFrequencyFrequency = 1 Then
                  vDonPP = True
                Else
                  For Each vPPD In pPP.Details
                    vDonPP = vPPD.Product.Donation
                    If vDonPP = False Then Exit For
                  Next vPPD
                End If
            End Select

            If vDonPP = False Then
              If pClaimDate.Length > 0 Then
                'See if we already have an OPS record for this claim date
                For Each vOPS In pPP.ScheduledPayments
                  If Len(vOPS.ClaimDate) > 0 Then
                    If CDate(vOPS.ClaimDate) = CDate(pClaimDate) Then vFound = True
                  End If
                  If vFound Then Exit For
                Next vOPS
                If vFound = False Then
                  vOPS = Nothing
                  pUseClaimDate = True
                End If
              End If
            End If
        End Select
      End If
      ProcessReversal = vOPS 'Return the OPS record (if any)

    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As OrderPaymentScheduleRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = OrderPaymentScheduleRecordSetTypes.opsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ops")
      ElseIf (pRSType And OrderPaymentScheduleRecordSetTypes.opsrtMain) > 0 Then
        vFields = "ops.order_number AS ops_order_number"
        vFields = "scheduled_payment_number, ops.due_date, claim_date, amount_due, amount_outstanding, revised_amount, expected_balance, scheduled_payment_status, schedule_creation_reason"
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pScheduledPaymentNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pScheduledPaymentNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(OrderPaymentScheduleRecordSetTypes.opsrtAll) & " FROM order_payment_schedule ops WHERE scheduled_payment_number = " & pScheduledPaymentNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, OrderPaymentScheduleRecordSetTypes.opsrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As OrderPaymentScheduleRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(OrderPaymentScheduleFields.opsfScheduledPaymentNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And OrderPaymentScheduleRecordSetTypes.opsrtMain) > 0 Then
          .SetItem(OrderPaymentScheduleFields.opsfOrderNumber, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfDueDate, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfClaimDate, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfAmountDue, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfAmountOutstanding, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfRevisedAmount, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfExpectedBalance, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfScheduledPaymentStatus, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfScheduleCreationReason, vFields)
        End If
        If (pRSType And OrderPaymentScheduleRecordSetTypes.opsrtAll) = OrderPaymentScheduleRecordSetTypes.opsrtAll Then
          .SetItem(OrderPaymentScheduleFields.opsfAmendedBy, vFields)
          .SetItem(OrderPaymentScheduleFields.opsfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub InitFromOPS(ByVal pEnv As CDBEnvironment, ByVal pOPS As OrderPaymentSchedule)
      mvEnv = pEnv
      If mvClassFields Is Nothing Then InitClassFields() 'Only call this when its not already initialised
      mvExisting = True
      With mvClassFields
        .Item(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).SetValue = pOPS.ScheduledPaymentNumber.ToString
        .Item(OrderPaymentScheduleFields.opsfOrderNumber).SetValue = pOPS.PlanNumber.ToString
        .Item(OrderPaymentScheduleFields.opsfDueDate).SetValue = pOPS.DueDate
        .Item(OrderPaymentScheduleFields.opsfClaimDate).SetValue = pOPS.ClaimDate
        .Item(OrderPaymentScheduleFields.opsfAmountDue).SetValue = pOPS.AmountDue.ToString
        .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).SetValue = pOPS.AmountOutstanding.ToString
        .Item(OrderPaymentScheduleFields.opsfRevisedAmount).SetValue = pOPS.RevisedAmount
        .Item(OrderPaymentScheduleFields.opsfExpectedBalance).SetValue = pOPS.ExpectedBalance.ToString
        .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).SetValue = pOPS.ScheduledPaymentStatusCode
        .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).SetValue = pOPS.ScheduleCreationReasonCode
        .Item(OrderPaymentScheduleFields.opsfAmendedBy).SetValue = pOPS.AmendedBy
        .Item(OrderPaymentScheduleFields.opsfAmendedOn).SetValue = pOPS.AmendedOn
      End With
    End Sub


    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(OrderPaymentScheduleFields.opsfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Friend Sub Create(ByVal pPlanNumber As Integer, ByVal pDueDate As String, ByVal pAmountDue As Double, ByVal pAmountOutstanding As Double, ByVal pExpectedBalance As Double, ByVal pCreationReason As OrderPaymentScheduleCreationReasons, Optional ByVal pClaimDate As String = "", Optional ByVal pProvisional As Boolean = False)
      Dim vStatus As OrderPaymentSchedulePaymentStatus

      If pProvisional Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsProvisional
      Else
        vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
        If pAmountOutstanding > 0 Then
          If pAmountDue > pAmountOutstanding Then vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
        ElseIf pAmountOutstanding = 0 Then
          vStatus = OrderPaymentSchedulePaymentStatus.opspsFullyPaid
        End If
      End If

      With mvClassFields
        .Item(OrderPaymentScheduleFields.opsfOrderNumber).IntegerValue = pPlanNumber
        .Item(OrderPaymentScheduleFields.opsfDueDate).Value = pDueDate
        .Item(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue = pAmountDue
        .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = pAmountOutstanding 'pAmountDue
        .Item(OrderPaymentScheduleFields.opsfExpectedBalance).DoubleValue = pExpectedBalance
        .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(vStatus)
        .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(pCreationReason)
        If pClaimDate.Length > 0 Then .Item(OrderPaymentScheduleFields.opsfClaimDate).Value = pClaimDate
      End With

    End Sub

    Public Sub CreateInAdvance(ByVal pEnv As CDBEnvironment, ByVal pPP As PaymentPlan, ByVal pPaymentAmount As Double, Optional ByVal pAllowOverPayment As Boolean = True, Optional ByVal pUpdateAmountDueOnly As Boolean = False)
      'Find existing in-advance ops and update, Or create a new one
      'If pAllowOverPayment is set then an existing in-advance line can be updated to allow an over-payment (I.e. in Trader)
      'Otherwise a new in-advance record created for the next renewal period (used in Auto SO Rec)
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vAmount As Double
      Dim vDue As String
      Dim vFound As Boolean

      If pPP.RenewalPending Then
        vDue = pPP.CalculateRenewalDate(pPP.RenewalDate, True)
      Else
        vDue = pPP.RenewalDate
      End If

      'Use pPP.CalculateBalanceForPaymentSchedule so that the calculations for AmountDue are the
      'same as those used when pPP.RegenerateScheduledPayments has created the provisional OPS.
      vAmount = pPP.CalculateBalanceForPaymentSchedule(OrderPaymentScheduleCreationReasons.opscrInAdvance, vDue)

      'Find existing record first
      If pUpdateAmountDueOnly Then
        If pPP.ScheduledPayments(False).Count() > 0 Then
          '
        End If
        If pPP.Balance = 0 And pPP.PaymentFrequencyFrequency = 1 And (CDate(TodaysDate()) > CDate(pPP.RenewalDate)) Then
          While CDate(TodaysDate()) > CDate(vDue)
            vDue = pPP.CalculateRenewalDate(vDue, True)
          End While
        End If
      End If

      For Each vOPS In pPP.ScheduledPayments
        If (vOPS.ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance Or vOPS.ScheduledPaymentStatus = OrderPaymentSchedulePaymentStatus.opspsProvisional) Then
          If CDate(vOPS.DueDate) = CDate(vDue) Then
            If pAllowOverPayment = True Then
              vFound = True
            Else
              vFound = (vOPS.AmountOutstanding >= pPaymentAmount)
              If vFound = False Then vDue = pPP.CalculateRenewalDate(vDue, True)
            End If
          End If
        Else
          If CDate(vOPS.DueDate) = CDate(vDue) Then
            'Payment for this date is not Provisional so go to next date
            vDue = pPP.CalculateRenewalDate(vDue, True)
          End If
        End If
        If vFound Then Exit For
      Next vOPS

      If vFound Then
        If pUpdateAmountDueOnly Then
          If vOPS.AmountDue < pPaymentAmount Then
            With vOPS
              .Update(.DueDate, pPaymentAmount, ((.AmountOutstanding - .AmountDue) + pPaymentAmount), .ExpectedBalance, .ClaimDate, .RevisedAmount)
            End With
          End If
        ElseIf vOPS.AmountOutstanding < pPaymentAmount Then
          'Update AmountDue and AmountOutstanding
          With vOPS
            .Update(.DueDate, FixTwoPlaces(.AmountDue + pPaymentAmount), FixTwoPlaces(.AmountOutstanding + pPaymentAmount), .ExpectedBalance, .ClaimDate, .RevisedAmount)
          End With
          vOPS.Save()
        End If
        Init(mvEnv, (vOPS.ScheduledPaymentNumber)) 'Set this class to be the OPS we found
      Else
        'Create a new record and add to the collection after it is saved
        Create(pPP.PlanNumber, vDue, vAmount, vAmount, 0, OrderPaymentScheduleCreationReasons.opscrInAdvance, pPP.FindNextClaimDate(vDue), True)
        Save()
        pPP.ScheduledPayments.Add(Me, CStr(ScheduledPaymentNumber))
      End If

    End Sub

    Public Sub Update(ByVal pDueDate As String, ByVal pAmountDue As Double, ByVal pAmountOutstanding As Double, ByVal pExpectedBalance As Double, Optional ByVal pClaimDate As String = "", Optional ByVal pRevisedAmount As String = "", Optional ByVal pCreationReason As OrderPaymentScheduleCreationReasons = 0)
      With mvClassFields
        .Item(OrderPaymentScheduleFields.opsfDueDate).Value = pDueDate
        .Item(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue = pAmountDue
        .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = pAmountOutstanding
        .Item(OrderPaymentScheduleFields.opsfExpectedBalance).DoubleValue = pExpectedBalance
        .Item(OrderPaymentScheduleFields.opsfClaimDate).Value = pClaimDate
        .Item(OrderPaymentScheduleFields.opsfRevisedAmount).Value = pRevisedAmount

        'Update the scheduled payment status
        If .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = 0 Then
          .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsFullyPaid)
        Else
          .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsDue)
          If .Item(OrderPaymentScheduleFields.opsfRevisedAmount).Value.Length > 0 Then
            If .Item(OrderPaymentScheduleFields.opsfRevisedAmount).DoubleValue > .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue Then
              .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsPartPaid)
            End If
          Else
            If .Item(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue > .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue Then
              .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsPartPaid)
            End If
          End If
          If (.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).SetValue = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsSkippedPayment) AndAlso .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments)) AndAlso pCreationReason = 0 Then
            'A DataUpdate has updated the AmountOutstanding so need to re-set the ScheduledPaymentStatus back to SkippedPayment
            If .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).ValueChanged Then .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsSkippedPayment)
          End If
        End If

        If .Item(OrderPaymentScheduleFields.opsfRevisedAmount).ValueChanged Then
          'The user has changed the revised amount so set creation reason as user amendment
          If GetCreationReason(.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value) <> OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments Then
            .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrUserAmendment)
          End If
        End If

        If .Item(OrderPaymentScheduleFields.opsfAmountDue).ValueChanged = True And pCreationReason <> OrderPaymentScheduleCreationReasons.opscrInAdvance Then
          'Amount due has changed and not as a result of updating a provisional line
          If GetCreationReason(mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value) = OrderPaymentScheduleCreationReasons.opscrInAdvance And pCreationReason = OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments Then
            'FA is updating the provisional line, leave the status as provisional
            .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsProvisional)
          ElseIf GetCreationReason(.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value) = OrderPaymentScheduleCreationReasons.opscrInAdvance And pCreationReason = 0 Then
            'If this was a provisional line then keep the status as Provisional
            .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsProvisional)
          End If
        ElseIf pCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance And ScheduledPaymentStatus = OrderPaymentSchedulePaymentStatus.opspsDue Then
          .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsProvisional)
        End If

        If (.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).SetValue = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsDue) AndAlso .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).SetValue = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrSkippedPayment)) AndAlso pCreationReason = 0 Then
          If .Item(OrderPaymentScheduleFields.opsfClaimDate).ValueChanged = True And .Item(OrderPaymentScheduleFields.opsfAmountOutstanding).ValueChanged = False Then
            'A DataUpdate has corrected the claim date so we also need to update the ScheduledPaymentStatus & ScheduleCreationReason as well
            .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsSkippedPayment)
            .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments)
          End If
        End If

        If pCreationReason <> 0 Then
          .Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(pCreationReason)
        End If
      End With

    End Sub

    Friend Sub SetClaimDate(ByVal pNewClaimDate As String)
      If pNewClaimDate.Length > 0 Then
        'Only update if status is D, P or V
        Select Case GetPaymentStatus(mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)
          Case OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedulePaymentStatus.opspsPartPaid, OrderPaymentSchedulePaymentStatus.opspsProvisional
            mvClassFields(OrderPaymentScheduleFields.opsfClaimDate).Value = pNewClaimDate
        End Select
      Else
        mvClassFields(OrderPaymentScheduleFields.opsfClaimDate).Value = ""
        mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrAutoPayMethodCancel)
      End If
    End Sub

    Friend Sub SetCancelled()
      mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsCancelled)
    End Sub

    Public Sub SetUnProcessedPayment(ByVal pUnProcessed As Boolean, ByVal pAmount As Double)
      SetUnProcessedPayment(pUnProcessed, pAmount, False) 'Don't know if this is a Loan payment so assume not
    End Sub

    Public Sub SetUnProcessedPayment(ByVal pUnProcessed As Boolean, ByVal pAmount As Double, ByVal pLoanPayment As Boolean)
      Dim vStatus As OrderPaymentSchedulePaymentStatus

      mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue - pAmount

      If pLoanPayment Then
        'Handle over-payments
        If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue < 0 Then
          'Loan payment has over-paid this OPS so set the RevisedAmount
          mvClassFields(OrderPaymentScheduleFields.opsfRevisedAmount).Value = FixTwoPlaces(DoubleValue(RevisedAmount) + (AmountDue + System.Math.Abs(AmountOutstanding))).ToString
          mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).Value = "0.00"
        ElseIf pAmount < 0 AndAlso pUnProcessed = False AndAlso AmountOutstanding > 0 AndAlso AmountOutstanding = DoubleValue(RevisedAmount) Then
          'Was over-paid and now it isn't so remove RevisedAmount & reset AmountOutstanding to AmountDue
          mvClassFields(OrderPaymentScheduleFields.opsfRevisedAmount).Value = ""
          mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).Value = AmountDue.ToString
        End If
      End If

      vStatus = GetPaymentStatus(mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)
      If pUnProcessed Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
      Else
        If GetPaymentStatus(mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).SetValue) = OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment Then
          vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
          If mvClassFields(OrderPaymentScheduleFields.opsfRevisedAmount).Value.Length > 0 Then
            If pLoanPayment = True AndAlso AmountOutstanding = 0 AndAlso DoubleValue(RevisedAmount) = System.Math.Abs(pAmount) Then
              mvClassFields(OrderPaymentScheduleFields.opsfRevisedAmount).Value = ""
              vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
            ElseIf mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue <> mvClassFields(OrderPaymentScheduleFields.opsfRevisedAmount).DoubleValue Then
              vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
            End If
          Else
            If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue <> mvClassFields(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue Then
              vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
            End If
          End If
        ElseIf ScheduledPaymentStatus <> OrderPaymentSchedulePaymentStatus.opspsProvisional Then
          If (pAmount * -1) = mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue Then
            vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
          Else
            vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
          End If
        End If
      End If

      mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(vStatus)

    End Sub

    Public Sub ProcessPayment(Optional ByVal pPayPlanIsCancelled As Boolean = False)
      Dim vStatus As OrderPaymentSchedulePaymentStatus

      If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = 0 Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsFullyPaid
      ElseIf mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = mvClassFields(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
        If pPayPlanIsCancelled Then vStatus = OrderPaymentSchedulePaymentStatus.opspsCancelled 'If PP was cancelled before posting this payment, then leave Status as Cancelled
      Else
        vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
        If pPayPlanIsCancelled Then vStatus = OrderPaymentSchedulePaymentStatus.opspsCancelled 'If PP was cancelled before posting this payment, then leave Status as Cancelled
      End If

      If ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance Then
        'If the Status is now back to Due then re-set to Provisional (setting to Part-Paid or Fully Paid is OK)
        If vStatus = OrderPaymentSchedulePaymentStatus.opspsDue Then vStatus = OrderPaymentSchedulePaymentStatus.opspsProvisional
      ElseIf ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrSkippedPayment And ScheduledPaymentStatus = OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment Then
        'Skipping an FinancialAdjustment OPS by BACS Messaging will have set the CreationReason to Skipped
        'whilst leaving the Status as Unprocessed, so re-setCreationReason to FinancialAdjustment
        mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments)
        vStatus = OrderPaymentSchedulePaymentStatus.opspsSkippedPayment
      End If

      mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(vStatus)
    End Sub

    Public Function DeterminePaymentScheduleStatus(ByVal pAmountOutstanding As Double, ByVal pAmountDue As Double, ByVal pPayPlanIsCancelled As Boolean) As OrderPaymentSchedulePaymentStatus
      Dim vStatus As OrderPaymentSchedulePaymentStatus
      If pAmountOutstanding = 0 Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsFullyPaid
      ElseIf pAmountOutstanding = pAmountDue Then
        vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
        If pPayPlanIsCancelled Then vStatus = OrderPaymentSchedulePaymentStatus.opspsCancelled
      Else
        vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
        If pPayPlanIsCancelled Then vStatus = OrderPaymentSchedulePaymentStatus.opspsCancelled
      End If
      Return vStatus
    End Function

    Public Function WriteOff(ByVal pAmount As Double) As Double
      Dim vStatus As OrderPaymentSchedulePaymentStatus

      Dim vAmount As Double = pAmount
      If vAmount > mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue Then vAmount = mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue
      mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = FixTwoPlaces(mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue - vAmount)
      vStatus = GetPaymentStatus(mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)

      If vAmount >= 0 Then
        'Write off
        'Need to include 0 for a fully-paid record where the payment is unprocessed (AmountOutstanding = 0)
        If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = 0 Then
          Select Case GetPaymentStatus(mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)
            Case OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment, OrderPaymentSchedulePaymentStatus.opspsArrears
              vStatus = OrderPaymentSchedulePaymentStatus.opspsWrittenOff
            Case OrderPaymentSchedulePaymentStatus.opspsPartPaid
              vStatus = OrderPaymentSchedulePaymentStatus.opspsFullyPaid
          End Select
        Else
          'Leave status as it is
        End If
      Else
        'Reinstate write off
        If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = 0 Then
          vStatus = OrderPaymentSchedulePaymentStatus.opspsFullyPaid
        ElseIf (mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = mvClassFields(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue) Then
          vStatus = OrderPaymentSchedulePaymentStatus.opspsDue
        Else
          vStatus = OrderPaymentSchedulePaymentStatus.opspsPartPaid
        End If
      End If

      mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(vStatus)
      Return vAmount
    End Function

    Public Sub Reverse(ByVal pPP As PaymentPlan, ByVal pAmount As Double, Optional ByVal pReverseCurrentOPSOnly As Boolean = False)
      'If pReverseCurrentOPSOnly is set then just "reverse" this OPS only (i.e. do not update ClaimDate etc.)
      Dim vOPS As OrderPaymentSchedule
      Dim vClaimDate As String = ""
      Dim vFound As Boolean
      Dim vUseClaimDate As Boolean

      If pAmount < 0 Then pAmount = pAmount * -1 'Ensure that this is a positive amount
      If (pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes OrElse pPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) AndAlso pReverseCurrentOPSOnly = False AndAlso pPP.PlanType <> CDBEnvironment.ppType.pptLoan Then
        Dim vAutoPayMethod As PaymentPlan.ppAutoPayMethods
        If pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
          vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMDD
        Else
          vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMCCCA
        End If
        vClaimDate = mvEnv.GetPaymentPlanAutoPayDate(Today, vAutoPayMethod, pPP.AutoPayBankAccount).ToString(CAREDateFormat)

        If CDate(mvClassFields.Item(OrderPaymentScheduleFields.opsfDueDate).Value) > CDate(vClaimDate) Then
          'Ensure we are not about to set the ClaimDate to earlier than the DueDate
          'This will be a payment made in advance of the date expected
          pReverseCurrentOPSOnly = True
        End If

        If pReverseCurrentOPSOnly = False Then
          vClaimDate = pPP.FindNextClaimDate(vClaimDate)
          vOPS = ProcessReversal(pPP, vClaimDate, vUseClaimDate)
          If Not (vOPS Is Nothing) Then
            vFound = True
            'This payment now due to be claimed on same day as another OPS, so update the other this payment
            With vOPS
              .Update(.DueDate, (.AmountDue + pAmount), (.AmountOutstanding + pAmount), .ExpectedBalance, .ClaimDate, (If(Len(.RevisedAmount) > 0, (Val(.RevisedAmount) + pAmount).ToString, "")), CType(IIf(.ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance, 0, OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments), OrderPaymentScheduleCreationReasons))
              .Save()
            End With
          End If
        End If
      End If

      If vFound = False Then
        'This payment is now outstanding again so set as Unprocessed and update ClaimDate
        'Reset the ClaimDate only if there was not already an existing OPS with this ClaimDate
        SetUnProcessedPayment(True, (pAmount * -1))
        With mvClassFields
          If vUseClaimDate Then .Item(OrderPaymentScheduleFields.opsfClaimDate).Value = vClaimDate
          .Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment)
        End With
      Else
        'Just set this payment as Unprocessed
        SetUnProcessedPayment(True, 0)
      End If

      'Always set this CreationReason to FinancialAdjustments so that it can never be deleted
      If GetCreationReason(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value) <> OrderPaymentScheduleCreationReasons.opscrInAdvance Then
        mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments)
      End If

    End Sub

    Public Function ReverseHistoricPayment(ByVal pPP As PaymentPlan, ByVal pAmount As Double, Optional ByVal pSetAsUnprocessed As Boolean = True, Optional ByVal pReverseCurrentOPSOnly As Boolean = False) As Integer
      'Process a reverse of an OPH record from prior to payment schedules
      'Basically this must be allocated against the payment schedule so either find a payment or create one
      'Return the OPS Number of the scheduled payment for the reversal
      'If pReverseCurrentOPSOnly is set then just create a new OPS for the "reverse"
      Dim vOPS As OrderPaymentSchedule = Nothing
      Dim vAutoStart As String = ""
      Dim vClaimDate As String = ""
      Dim vDueDate As String
      Dim vFound As Boolean
      Dim vUseClaimDate As Boolean

      If pAmount < 0 Then pAmount = pAmount * -1 'Ensure that this is a positive amount
      If (pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or pPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) Then ' And pReverseCurrentOPSOnly = False Then
        'For DD/CCCA see if there is an existing OPS that this payment can be attached to (just like a "normal" reversal)
        If pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
          vAutoStart = pPP.DirectDebit.StartDate
        Else
          vAutoStart = pPP.CreditCardAuthority.StartDate
        End If
        If pReverseCurrentOPSOnly = False Then
          Dim vAutoPayMethod As PaymentPlan.ppAutoPayMethods
          If pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
            vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMDD
          Else
            vAutoPayMethod = PaymentPlan.ppAutoPayMethods.ppAPMCCCA
          End If
          vDueDate = mvEnv.GetPaymentPlanAutoPayDate(Today, vAutoPayMethod, pPP.AutoPayBankAccount).ToString(CAREDateFormat)
          vClaimDate = pPP.FindNextClaimDate(vDueDate)
          If CDate(vAutoStart) > CDate(vClaimDate) Then vClaimDate = pPP.FindNextClaimDate(vAutoStart)
          vOPS = ProcessReversal(pPP, vClaimDate, vUseClaimDate)
          If Not (vOPS Is Nothing) Then
            'We have found an OPS record that this payment can be attached to so update the OPS
            vFound = True
            With vOPS
              .Update(.DueDate, (.AmountDue + pAmount), (.AmountOutstanding + pAmount), .ExpectedBalance, .ClaimDate, If(Len(.RevisedAmount) > 0, (Val(.RevisedAmount) + pAmount).ToString, ""), CType(IIf(.ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance, 0, OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments), OrderPaymentScheduleCreationReasons))
              If pSetAsUnprocessed Then .SetUnProcessedPayment(True, 0)
              .Save()
            End With
          End If
        End If
      End If

      If vFound = False Then
        'There is no existing payment that can be used so create
        'Use a DueDate prior to the first OPS (pPP.ScheduledPayments only contains unpaid OPS)
        vDueDate = mvEnv.Connection.GetValue("SELECT MIN(due_date) FROM order_payment_schedule WHERE order_number = " & pPP.PlanNumber)
        If Not IsDate(vDueDate) Then vDueDate = pPP.NextPaymentDue
        If pPP.TermUnits = PaymentPlan.OrderTermUnits.otuWeekly Or pPP.PaymentFrequencyPeriod = PaymentFrequency.PaymentFrequencyPeriods.pfpDays Then
          vDueDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -pPP.PaymentFrequencyInterval, CDate(vDueDate)))
        Else
          vDueDate = AddMonths((pPP.RenewalDate), vDueDate, -pPP.PaymentFrequencyInterval)
        End If
        'Reset the ClaimDate as the "original" date is to be used (just like if there had been an OPS record)
        If vUseClaimDate = False And (pPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Or pPP.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) Then
          vClaimDate = pPP.FindNextClaimDate(vDueDate)
          If CDate(vAutoStart) > CDate(vClaimDate) Then vClaimDate = pPP.FindNextClaimDate(vAutoStart)
        End If

        vOPS = New OrderPaymentSchedule
        With vOPS
          .Init(mvEnv)
          .Create(pPP.PlanNumber, vDueDate, pAmount, pAmount, pPP.Balance, OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments, vClaimDate)
          If pSetAsUnprocessed Then .SetUnProcessedPayment(True, 0)
          .Save()
        End With
      End If
      ReverseHistoricPayment = vOPS.ScheduledPaymentNumber
    End Function

    'Public Sub SetArrears()
    '  'At renewals time, may need to set status to Arrears
    '  If mvClassFields(opsfAmountOutstanding).DoubleValue > 0 Then mvClassFields(opsfScheduledPaymentStatus).Value = SetPaymentStatus(opspsArrears)
    'End Sub

    Public Sub SkipPayment()
      If ScheduledPaymentStatus = OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment And ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrFinancialAdjustments Then
        'An unprocessed Financial Adjustment - set CreationReason to Skipped
        'Processing the payment will re-set it back to FinancialAdjustment
        mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrSkippedPayment)
      Else
        mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsSkippedPayment)
        If ScheduleCreationReason = OrderPaymentScheduleCreationReasons.opscrInAdvance Then mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrSkippedPayment)
      End If
    End Sub

    Friend Sub Delete()
      mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Function DeleteUnpaidScheduledPayment() As Boolean
      'If this OPS has not payments linked to it, then delete it and return True
      'Otherwise leave it alone and return False
      Dim vDelete As Boolean

      If mvExisting = True Then
        If mvEnv.Connection.GetCount("order_payment_history", Nothing, "scheduled_payment_number = " & ScheduledPaymentNumber & "  AND order_number = " & PlanNumber) = 0 Then
          Delete()
          vDelete = True
        End If
      End If
      DeleteUnpaidScheduledPayment = vDelete
    End Function

    Public Sub ProcessReanalysis(ByVal pAmount As Double)
      'Used in Trader during a Reanalysis of a Pay Plan payment
      mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue - pAmount
      If pAmount < 0 Then
        mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment)
      Else
        If AmountOutstanding > 0 Then
          mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsPartPaid)
          If RevisedAmount.Length > 0 Then
            If AmountOutstanding = Val(RevisedAmount) Then
              mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsDue)
            End If
          Else
            If AmountOutstanding = AmountDue Then
              mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsDue)
            End If
          End If
        Else
          mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(OrderPaymentSchedulePaymentStatus.opspsFullyPaid)
        End If
      End If

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(OrderPaymentScheduleFields.opsfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AmountDue() As Double
      Get
        AmountDue = mvClassFields.Item(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue
      End Get
    End Property

    Public ReadOnly Property AmountOutstanding() As Double
      Get
        AmountOutstanding = mvClassFields.Item(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue
      End Get
    End Property

    Public ReadOnly Property ClaimDate() As String
      Get
        ClaimDate = mvClassFields.Item(OrderPaymentScheduleFields.opsfClaimDate).Value
      End Get
    End Property

    Public ReadOnly Property DueDate() As String
      Get
        DueDate = mvClassFields.Item(OrderPaymentScheduleFields.opsfDueDate).Value
      End Get
    End Property

    Public ReadOnly Property ExpectedBalance() As Double
      Get
        ExpectedBalance = mvClassFields.Item(OrderPaymentScheduleFields.opsfExpectedBalance).DoubleValue
      End Get
    End Property

    Public ReadOnly Property PlanNumber() As Integer
      Get
        PlanNumber = mvClassFields.Item(OrderPaymentScheduleFields.opsfOrderNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RevisedAmount() As String
      Get
        RevisedAmount = mvClassFields.Item(OrderPaymentScheduleFields.opsfRevisedAmount).Value
      End Get
    End Property

    Public ReadOnly Property ScheduleCreationReasonCode() As String
      Get
        ScheduleCreationReasonCode = mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentNumber() As Integer
      Get
        ScheduledPaymentNumber = mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentStatusCode() As String
      Get
        ScheduledPaymentStatusCode = mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentStatus() As OrderPaymentSchedulePaymentStatus
      Get
        ScheduledPaymentStatus = GetPaymentStatus(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)
      End Get
    End Property

    Public ReadOnly Property ScheduleCreationReason() As OrderPaymentScheduleCreationReasons
      Get
        ScheduleCreationReason = GetCreationReason(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value)
      End Get
    End Property

    Public Property PaymentAmount() As Double
      Get
        PaymentAmount = mvPaymentAmount
      End Get
      Set(ByVal Value As Double)
        'Used by Trader to hold the amount of the payment
        mvPaymentAmount = Value
      End Set
    End Property

    Public ReadOnly Property HasPayments() As Boolean
      Get
        'Return True if there have been any payments, otherwise False
        If mvEnv.Connection.GetCount("order_payment_history", Nothing, "order_number = " & PlanNumber & " AND scheduled_payment_number = " & ScheduledPaymentNumber) > 0 Then
          HasPayments = True
        End If
      End Get
    End Property

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        Select Case pAttributeName
          Case "CheckValue"
            'This will come in as either True or False from OSP page (Check-box control)
            'or, Y or N from PPS page (no control)
            If Len(Value) > 1 Then
              mvSCCheckValue = CBool(Value)
            Else
              mvSCCheckValue = BooleanValue(Value)
            End If
          Case "PaymentAmount"
            mvPaymentAmount = Val(Value)
          Case "PaymentPlanNumber"
            mvClassFields.Item(OrderPaymentScheduleFields.opsfOrderNumber).Value = Value
          Case "ScheduledPaymentStatusDesc", "OrigAmountDue", "LineNumber" 'In case they come back from the client
            '
          Case Else
            mvClassFields.ItemValue(pAttributeName) = Value
        End Select
      End Set
    End Property

    Public Property SCCheckValue() As Boolean
      Get
        SCCheckValue = mvSCCheckValue
      End Get
      Set(ByVal Value As Boolean)
        mvSCCheckValue = Value
      End Set
    End Property

    ''' <summary>
    ''' The OrderPaymentHistories of this scheduled payment, empty of no payments made
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OrderPaymentHistory() As List(Of OrderPaymentHistory)
      Get
        If mvOrderPaymentHistory Is Nothing Then
          Me.OrderPaymentHistory = Me.GetOrderPaymentHistory()
        End If
        Return mvOrderPaymentHistory
      End Get
      Private Set(ByVal value As List(Of OrderPaymentHistory))
        mvOrderPaymentHistory = value
      End Set
    End Property


    Public Function GetDataAsParameters() As CDBParameters
      Dim vParams As New CDBParameters
      Dim vField As ClassField

      For Each vField In mvClassFields
        vParams.Add(ProperName((vField.Name)), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(vField.Value), vField.Value))
      Next
      'add paymentplan number as well, as we have to get away from using ordernumber
      vParams.Add("PaymentPlanNumber", CDBField.FieldTypes.cftLong, mvClassFields(OrderPaymentScheduleFields.opsfOrderNumber).Value)
      vParams.Add("OrigAmountDue", CDBField.FieldTypes.cftNumeric, FixedFormat(mvClassFields.Item(OrderPaymentScheduleFields.opsfAmountDue).DoubleValue))
      vParams.Add("ScheduledPaymentStatusDesc", CDBField.FieldTypes.cftCharacter, mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value)
      vParams.Add("PaymentAmount", CDBField.FieldTypes.cftNumeric, FixedFormat(mvPaymentAmount))
      vParams.Add("CheckValue", CDBField.FieldTypes.cftCharacter, BooleanString(mvSCCheckValue))
      GetDataAsParameters = vParams
    End Function

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      Select Case pAttributeName
        Case "CheckValue", "ScheduledPaymentStatusDesc"
          LineDataType = CDBField.FieldTypes.cftCharacter
        Case "PaymentAmount", "OrigAmountDue"
          LineDataType = CDBField.FieldTypes.cftNumeric
        Case "PaymentPlanNumber", "LineNumber"
          LineDataType = CDBField.FieldTypes.cftLong
        Case Else
          LineDataType = mvClassFields.ItemDataType(pAttributeName)
      End Select
    End Function

    Public Sub AddPayment(ByRef pPaymentAmount As Double)
      mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue - pPaymentAmount
      If mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue < 0 Then mvClassFields(OrderPaymentScheduleFields.opsfAmountOutstanding).DoubleValue = 0
      mvPaymentAmount = pPaymentAmount
    End Sub

    Friend Sub SetLoanMissedPayment()
      Dim vStatus As OrderPaymentSchedulePaymentStatus = ScheduledPaymentStatus
      Select Case ScheduledPaymentStatus
        Case OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedulePaymentStatus.opspsPartPaid
          vStatus = OrderPaymentSchedulePaymentStatus.opspsMissedLoanPayment
        Case OrderPaymentSchedulePaymentStatus.opspsProvisional
          vStatus = OrderPaymentSchedulePaymentStatus.opspsMissedLoanPayment
          mvClassFields(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value = SetCreationReason(OrderPaymentScheduleCreationReasons.opscrPaymentPlanMaintenance)
        Case OrderPaymentSchedulePaymentStatus.opspsUnprocessedPayment
          '?
      End Select
      mvClassFields(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value = SetPaymentStatus(vStatus)
    End Sub

    ''' <summary>Has a provisional record been updated to be non-provisional?</summary>
    Friend ReadOnly Property IsUpdateFromProvisional() As Boolean
      Get
        Dim vIsUpdate As Boolean = False
        If mvExisting = True AndAlso AmountOutstanding > 0 _
        AndAlso (mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).ValueChanged = True OrElse mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).ValueChanged = True) Then
          If GetPaymentStatus(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduledPaymentStatus).Value) <> OrderPaymentSchedulePaymentStatus.opspsProvisional _
          AndAlso GetCreationReason(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).SetValue) = OrderPaymentScheduleCreationReasons.opscrInAdvance _
          AndAlso GetCreationReason(mvClassFields.Item(OrderPaymentScheduleFields.opsfScheduleCreationReason).Value) <> OrderPaymentScheduleCreationReasons.opscrInAdvance Then
            vIsUpdate = True
          End If
        End If
        Return vIsUpdate
      End Get
    End Property

    Private mvPaymentPlan As PaymentPlan = Nothing
    Public ReadOnly Property PaymentPlan As PaymentPlan
      Get
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan
          mvPaymentPlan.Init(mvEnv, Me.mvClassFields(OrderPaymentScheduleFields.opsfOrderNumber).IntegerValue)
          If Not mvPaymentPlan.Existing Then
            Throw New InvalidOperationException("Cannot find the payment plan for this payment shedule record.")
          End If
        End If
        Return mvPaymentPlan
      End Get
    End Property

    Public ReadOnly Property PotentialClaimDates() As List(Of Date)
      Get
        If Not (Me.ScheduledPaymentStatusCode.Equals("D", StringComparison.InvariantCultureIgnoreCase) OrElse
           Me.ScheduledPaymentStatusCode.Equals("V", StringComparison.InvariantCultureIgnoreCase)) Then
          RaiseError(DataAccessErrors.daeClaimDateCannotChangeStatus, Me.ScheduledPaymentNumber.ToString())
        End If
        If Me.PaymentPlan.DirectDebit Is Nothing OrElse
           Not Me.PaymentPlan.DirectDebit.Existing Then
          RaiseError(DataAccessErrors.daeClaimDateCannotChangeNoDD, Me.ScheduledPaymentNumber.ToString(), Me.ScheduledPaymentStatusCode)
        End If

        If Me.ScheduleCreationReasonCode.Equals("FA", StringComparison.InvariantCultureIgnoreCase) Then
          'BR21254 Check that there is a payment that fully paid this scheduled payment
          Dim vHistory As OrderPaymentHistory = Me.OrderPaymentHistory.Find(Function(pOrderPaymentHistory As OrderPaymentHistory)
                                                                              If pOrderPaymentHistory.Amount = Me.AmountDue Then Return True
                                                                              Return False
                                                                            End Function)
          'BR21254 Check that the sum of all payments is zero i.e. the payment is due
          Dim vSumOfPayments = Aggregate vHist In Me.OrderPaymentHistory
                               Into Sum(vHist.Amount)
          If vHistory Is Nothing OrElse vSumOfPayments <> 0 Then
            'BR21254 The scheduled payment was not fully paid or fully reversed
            RaiseError(DataAccessErrors.daeClaimDateCannotChangeNotReconciled, Me.ScheduledPaymentNumber.ToString())
          End If
        End If

        Dim vResult As New List(Of Date)
        Dim vBankAccount As New BankAccount(mvEnv)
        vBankAccount.InitWithPrimaryKey(New CDBFields({New CDBField("bank_account",
                                                                    Me.PaymentPlan.DirectDebit.BankAccount)}))
        Select Case mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAutoPayClaimDateMethod)
          Case "A"
            For Each vClaimDate As ScheduledClaimDate In vBankAccount.GetClaimDates("DD")
              If CDate(vClaimDate.ClaimDate) >= Me.EarliestValidClaimDate AndAlso
                 CDate(vClaimDate.LatestDueDate) >= CDate(Me.DueDate) AndAlso
                 CDate(vClaimDate.ClaimDate) <= Me.NextPaymentDueDate Then
                vResult.Add(CDate(vClaimDate.ClaimDate))
              End If
            Next vClaimDate
          Case "D"
            Dim vClaimDays As New List(Of Integer)
            For Each vClaimDay As BankAccountClaimDay In vBankAccount.GetClaimDays("DD")
              vClaimDays.Add(vClaimDay.ClaimDay)
            Next vClaimDay
            Dim vTestDate As Date = Me.EarliestValidClaimDate
            While vTestDate < Me.NextPaymentDueDate
              If vClaimDays.Contains(vTestDate.Day) OrElse
                 (vTestDate.AddDays(1).Month > vTestDate.Month AndAlso
                  vClaimDays.Max > vTestDate.Day) Then
                Dim vNonWorkingDayAction As String = String.Empty
                For Each vBankAccountClaimDay As BankAccountClaimDay In vBankAccount.GetClaimDays("DD")
                  If vBankAccountClaimDay.ClaimDay = CInt(vTestDate.Day) Then
                    vNonWorkingDayAction = vBankAccountClaimDay.NonWorkingDayBehaviour
                  End If
                Next vBankAccountClaimDay
                If String.IsNullOrEmpty(vNonWorkingDayAction) Then
                  For Each vBankAccountClaimDay As BankAccountClaimDay In vBankAccount.GetClaimDays("DD")
                    If vBankAccountClaimDay.ClaimDay = vClaimDays.Max Then
                      vNonWorkingDayAction = vBankAccountClaimDay.NonWorkingDayBehaviour
                    End If
                  Next vBankAccountClaimDay
                End If
                If vNonWorkingDayAction.Equals("N", StringComparison.InvariantCultureIgnoreCase) Then
                  vResult.Add(vTestDate.NextWorkingDay(mvEnv))
                ElseIf vNonWorkingDayAction.Equals("P", StringComparison.InvariantCultureIgnoreCase) Then
                  vResult.Add(vTestDate.PreviousWorkingDay(mvEnv))
                Else
                  vResult.Add(vTestDate)
                End If
              End If
              vTestDate = vTestDate.AddDays(1)
            End While
        End Select
        If vResult.Count < 1 Then
          vResult.Add(CDate(Me.ClaimDate))
        End If
        Return New List(Of Date)(vResult.Distinct)
      End Get
    End Property

    Private mvEarliestValidClaimDate As Date = Date.MinValue
    Private mvEarliestValidClaimDateValid As Boolean = False
    Private ReadOnly Property EarliestValidClaimDate As Date
      Get
        If Not mvEarliestValidClaimDateValid Then
          mvEarliestValidClaimDate = If(Not String.IsNullOrWhiteSpace(Me.DueDate), CDate(Me.DueDate), Date.MinValue)
          Dim vClaimRunData As DataTable = New SQLStatement(mvEnv.Connection,
                                                              "Max(cru.to_date)",
                                                              "collection_run_audit cru",
                                                              New CDBFields({New CDBField("cru.process_type", "DD"),
                                                                             New CDBField("cru.bank_account", Me.PaymentPlan.DirectDebit.BankAccount)})).GetDataTable
          If vClaimRunData.Rows.Count > 0 AndAlso
             Not IsDBNull(vClaimRunData.Rows(0)(0)) AndAlso
             CDate(vClaimRunData.Rows(0)(0)).AddDays(1) > mvEarliestValidClaimDate Then
            mvEarliestValidClaimDate = CDate(vClaimRunData.Rows(0)(0)).AddDays(1)
          End If
          If Not String.IsNullOrWhiteSpace(Me.PaymentPlan.DirectDebit.StartDate) AndAlso
             CDate(Me.PaymentPlan.DirectDebit.StartDate) > mvEarliestValidClaimDate Then
            mvEarliestValidClaimDate = CDate(Me.PaymentPlan.DirectDebit.StartDate)
          End If
          mvEarliestValidClaimDateValid = True
        End If
        Return mvEarliestValidClaimDate
      End Get
    End Property

    Private mvNextPaymentDueDate As Date = Date.MaxValue
    Private mvNextPaymentDueDateValid As Boolean = False

    Public Sub New()

    End Sub

    Public Sub New(pEnv As CDBEnvironment)
      Me.Environment = pEnv
    End Sub

    Private ReadOnly Property NextPaymentDueDate As Date
      Get
        If Not mvNextPaymentDueDateValid Then
          Dim vSql As New StringBuilder
          vSql.AppendLine("SELECT (SELECT Min(ops1.due_date) ")
          vSql.AppendLine("        FROM   order_payment_schedule ops1 ")
          vSql.AppendLine("        WHERE  ops1.due_date > ops.due_date ")
          vSql.AppendLine("               AND ops1.order_number = ops.order_number) AS next_due_date, ")
          vSql.AppendLine("       ops.claim_date ")
          vSql.AppendLine("FROM   order_payment_schedule ops ")
          vSql.AppendFormat("WHERE  ops.scheduled_payment_number = {0}", Me.ScheduledPaymentNumber)
          Dim vData As DataTable = New SQLStatement(mvEnv.Connection, vSql.ToString).GetDataTable
          If vData.Rows.Count > 0 AndAlso
             Not IsDBNull(vData.Rows(0)("next_due_date")) AndAlso
             Not String.IsNullOrWhiteSpace(CStr(vData.Rows(0)("next_due_date"))) Then
            mvNextPaymentDueDate = CDate(vData.Rows(0)("next_due_date"))
          End If
          Dim vRenewalDate As Date = CDate(If(Me.PaymentPlan.RenewalPending,
                                              Me.PaymentPlan.CalculateRenewalDate(Me.PaymentPlan.RenewalDate, True),
                                              Me.PaymentPlan.RenewalDate))
          If vRenewalDate < mvNextPaymentDueDate Then
            mvNextPaymentDueDate = CDate(Me.PaymentPlan.RenewalDate)
          End If
          mvNextPaymentDueDateValid = True
        End If
        Return mvNextPaymentDueDate
      End Get
    End Property

    Public Sub AmendClaimDate(pClaimDate As Date)
      If Not Me.PotentialClaimDates.Contains(pClaimDate) Then
        Throw New InvalidOperationException(String.Format("Order Payment Schedule number {0} cannot have a claim date of {1:" & CAREDateFormat() & "}.", Me.ScheduledPaymentNumber, pClaimDate))
      End If
      Me.mvClassFields(OrderPaymentScheduleFields.opsfClaimDate).Value = pClaimDate.ToString(CAREDateFormat)
    End Sub

    Public Overrides Function ToString() As String
      Return String.Format("ScheduledPaymentNumber = {0}", Me.ScheduledPaymentNumber)
    End Function
    ''' <summary>
    ''' Gets the OrderPaymentHistory for this OrderPaymentSchedule
    ''' </summary>
    ''' <returns>A List of OrderPaymentHistory objects or an empty List of OrderPaymentHistory objects</returns>
    ''' <remarks></remarks>
    Private Function GetOrderPaymentHistory() As List(Of OrderPaymentHistory)

      Dim vList As List(Of OrderPaymentHistory) = New List(Of OrderPaymentHistory)
      If Me.Existing Then
        For Each vHistory As OrderPaymentHistory In Me.PaymentPlan.PaymentHistory(Access.PaymentPlan.PaymentHistoryOrderByTypes.phobtTransDatePaymentNumber)
          If vHistory.OrderNumber = Me.PlanNumber And CInt(vHistory.ScheduledPaymentNumber) = Me.ScheduledPaymentNumber Then
            vList.Add(vHistory)
          End If
        Next
      End If
      Return vList
    End Function

    Protected Friend ReadOnly Property ClassFields As ClassFields
      Get
        CheckClassFields()
        Return mvClassFields
      End Get
    End Property

    Public ReadOnly Property FieldNames As String Implements IDbSelectable.DbFieldNames
      Get
        Return Me.ClassFields.FieldNames(Me.Environment, Me.ClassFields.TableNameAndAlias)
      End Get
    End Property

    Public ReadOnly Property AliasedTableName As String Implements IDbSelectable.DbAliasedTableName
      Get
        Return Me.ClassFields.TableNameAndAlias
      End Get
    End Property

    Protected Sub CheckClassFields()
      If mvClassFields Is Nothing Then
        InitClassFields()
      End If
    End Sub

    Public Sub LoadFromRow(pRow As DataRow) Implements IDbLoadable.LoadFromRow
      InitClassFields()
      mvExisting = True
      Dim vName As String
      Dim vUseProperNames As Boolean = False
      vUseProperNames = pRow IsNot Nothing AndAlso Not pRow.Table.Columns.Cast(Of DataColumn).Any(Function(pColumn) pColumn.ColumnName.Contains("_"))
      For Each vClassField As ClassField In mvClassFields
        If vUseProperNames Then vName = vClassField.ProperName Else vName = vClassField.Name
        If pRow.Table.Columns.Contains(vName) Then
          vClassField.SetValue = pRow.Item(vName).ToString
        End If
      Next
    End Sub

    Public Property Environment As CDBEnvironment
      Get
        Return mvEnv
      End Get
      Private Set(pEnv As CDBEnvironment)
        mvEnv = pEnv
      End Set
    End Property

    Public ReadOnly Property IsInFirstYearOfSchedule As Boolean
      Get
        Return CDate(Me.DueDate) < CDate(Me.PaymentPlan.StartDate).AddYears(1)
      End Get
    End Property
  End Class
End Namespace
