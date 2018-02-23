Imports System.Linq
Namespace Access

  Public Class PaymentPlanChange
    Inherits CARERecord

    Private mvPaymentPlan As PaymentPlan
    Private mvChangeRecordsRequired As Boolean
    Private mvPreviousPaymentPlanBalance As Double
    Private mvDetails As List(Of PaymentPlanChangeDetail)
    Private mvExistingPayPlanDetails As List(Of PaymentPlanDetail)
    Private mvExistingSchedulePayments As List(Of OrderPaymentSchedule)

    '--------------------------------------------------
    'Enum defining all the fields
    '--------------------------------------------------
    Private Enum PaymentPlanChangeFields
      AllFields = 0
      PaymentPlanChangeNumber
      PaymentPlanNumber
      PaymentPlanChangeDate
      Amount
      JournalNumber
      TermStartDate
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields

        .Add("payment_plan_change_number", CDBField.FieldTypes.cftLong)
        .Add("payment_plan_number", CDBField.FieldTypes.cftLong)
        .Add("payment_plan_change_date", CDBField.FieldTypes.cftDate)
        .Add("amount", CDBField.FieldTypes.cftNumeric)
        .Add("journal_number", CDBField.FieldTypes.cftLong)
        .Add("term_start_date", CDBField.FieldTypes.cftDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPayPlanChangesTermStartDate)

        .Item(PaymentPlanChangeFields.PaymentPlanChangeNumber).PrimaryKey = True

        .Item(PaymentPlanChangeFields.PaymentPlanChangeNumber).PrefixRequired = True
        .Item(PaymentPlanChangeFields.PaymentPlanNumber).PrefixRequired = True

        .SetControlNumberField(PaymentPlanChangeFields.PaymentPlanChangeNumber, "PX")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy As Boolean
      Get
        Return True
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias As String
      Get
        Return "ppc"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName As String
      Get
        Return "payment_plan_changes"
      End Get
    End Property

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvPaymentPlan = Nothing
      mvDetails = Nothing
      mvExistingPayPlanDetails = Nothing
      mvPreviousPaymentPlanBalance = 0
    End Sub

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pPaymentPlan As PaymentPlan)
      Me.New(pEnv, pPaymentPlan, pPaymentPlan.Balance)
    End Sub
    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pPaymentPlan As PaymentPlan, ByVal pOldBalance As Double)
      MyBase.New(pEnv)
      Init()
      mvPaymentPlan = pPaymentPlan
      mvChangeRecordsRequired = mvEnv.GetConfigOption("fp_record_payment_plan_changes")
      If mvChangeRecordsRequired Then
        Dim vCSPayMethod As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSPayMethod)
        If vCSPayMethod.Length > 0 AndAlso mvPaymentPlan.PaymentMethod = vCSPayMethod Then mvChangeRecordsRequired = False 'Do not generate changes for CreditSales
      End If
      If mvChangeRecordsRequired Then
        mvExistingPayPlanDetails = New List(Of PaymentPlanDetail)
        If pPaymentPlan.Existing Then
          Dim vDetail As New PaymentPlanDetail
          vDetail.Init(mvEnv)
          Dim vAnsiJoins As New AnsiJoins
          vAnsiJoins.Add("products p", "od.product", "p.product")
          vAnsiJoins.Add("rates r", "od.product", "r.product", "od.rate", "r.rate")
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("order_number", pPaymentPlan.PlanNumber)
          Dim vSQL As New SQLStatement(mvEnv.Connection, vDetail.GetRecordSetFields(PaymentPlanDetail.PaymentPlanDetailRecordSetTypes.odrtAll), "order_details od", vWhereFields, "detail_number", vAnsiJoins)
          Dim vRS As CDBRecordSet = vSQL.GetRecordSet
          While vRS.Fetch() = True
            vDetail = New PaymentPlanDetail
            vDetail.InitFromRecordSet(mvEnv, vRS, PaymentPlanDetail.PaymentPlanDetailRecordSetTypes.odrtAll)
            mvExistingPayPlanDetails.Add(vDetail)
          End While
          vRS.CloseRecordSet()
          mvPreviousPaymentPlanBalance = pOldBalance
          mvExistingSchedulePayments = New List(Of OrderPaymentSchedule)
          For Each vOPS As OrderPaymentSchedule In mvPaymentPlan.ScheduledPayments(False)
            mvExistingSchedulePayments.Add(vOPS)
          Next
        End If
      End If
    End Sub

#Region "Properties"
    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property PaymentPlanChangeNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeFields.PaymentPlanChangeNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeFields.PaymentPlanNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PaymentPlanChangeDate() As Date
      Get
        Return CDate(mvClassFields(PaymentPlanChangeFields.PaymentPlanChangeDate).Value)
      End Get
    End Property
    Public ReadOnly Property Amount() As Double
      Get
        Return mvClassFields(PaymentPlanChangeFields.Amount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property JournalNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeFields.JournalNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TermStartDate() As String
      Get
        'Could be null
        Return mvClassFields.Item(PaymentPlanChangeFields.TermStartDate).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(PaymentPlanChangeFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(PaymentPlanChangeFields.AmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ChangeRecordsRequired() As Boolean
      Get
        Return mvChangeRecordsRequired
      End Get
    End Property
#End Region

#Region "Methods"

    Public Sub GenerateChanges(ByVal pSchedulePayments As Collection, ByVal pReason As OrderPaymentSchedule.OrderPaymentScheduleCreationReasons, ByVal pChangeDate As Date)
      If mvChangeRecordsRequired Then
        'This routine is used when the payment plan order payment schedule is regenerated
        'Initial payment plan creation
        'Payment plan maintenance
        'Renewals and reminders
        Debug.Print(String.Format("PPChange: Generating Changes for Payment Plan {0} - Reason {1}", mvPaymentPlan.PlanNumber, pReason.ToString))

        Dim vGenerateChanges As Boolean = True
        '(1) Calculate change in Payment Plan balance
        Debug.WriteLine(String.Format("Payment Plan Balance was '{0}', now '{1}'", mvPreviousPaymentPlanBalance.ToString("F"), mvPaymentPlan.Balance.ToString("F")))

        '(2) Calculate change in Payment Plan Details balance
        Dim vPPDOriginalBal As Double = Aggregate vPPD1 In mvExistingPayPlanDetails Into Sum(vPPD1.Balance)
        Dim vPPDNewBal As Double = 0
        For Each vPPD2 As PaymentPlanDetail In mvPaymentPlan.Details
          vPPDNewBal += vPPD2.Balance
        Next
        vPPDOriginalBal = FixTwoPlaces(vPPDOriginalBal)
        vPPDNewBal = FixTwoPlaces(vPPDNewBal)
        Debug.WriteLine(String.Format("Payment Plan Details Balance was '{0}'. now '{1}'", vPPDOriginalBal.ToString("F"), vPPDNewBal.ToString("F")))
        Dim vPPDChangeAmount As Double = 0
        If vPPDOriginalBal <> vPPDNewBal Then vPPDChangeAmount = FixTwoPlaces(vPPDNewBal - vPPDOriginalBal)

        '(3) Calculate total payment schedule due
        Dim vOPSOriginaloutstanding As Double = 0
        If mvExistingSchedulePayments IsNot Nothing Then vOPSOriginaloutstanding = Aggregate vOPS1 In mvExistingSchedulePayments Into Sum(vOPS1.AmountOutstanding)
        Dim vOPSTotal As Double = 0
        For Each vOPS As OrderPaymentSchedule In pSchedulePayments
          With vOPS
            Select Case .ScheduledPaymentStatus
              Case OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsDue, OrderPaymentSchedule.OrderPaymentSchedulePaymentStatus.opspsPartPaid
                If .ScheduleCreationReason <> OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrInAdvance Then vOPSTotal += .AmountOutstanding
            End Select
          End With
        Next
        vOPSOriginaloutstanding = FixTwoPlaces(vOPSOriginaloutstanding)
        vOPSTotal = FixTwoPlaces(vOPSTotal)
        Debug.WriteLine(String.Format("Payment Schedule was: {0}, now {1}", vOPSOriginaloutstanding.ToString("F"), vOPSTotal.ToString("F")))
        If Not (
                 mvPaymentPlan.FixedRenewalCycle And _
                (mvPaymentPlan.PreviousRenewalCycle Or (mvEnv.GetConfigOption("recalculate_membership_balance") And mvPaymentPlan.UseStartDateForTerm)) And _
                 mvPaymentPlan.MembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfAnnualTerm And _
                (mvPaymentPlan.ProportionalBalanceSetting And _
                (PaymentPlan.ProportionalBalanceConfigSettings.pbcsFullPayment + PaymentPlan.ProportionalBalanceConfigSettings.pbcsNew)) > 0 _
               ) Then
          'BR19786 - Under the above conditions, a payment plan not balancing is valid, so we want to generate changes. The condition appears in PaymentPlan.GetProRataRenewalAmount
          '(4) Compare the Payment Plan, Payment Plan Details and OPS balances
          If (mvPaymentPlan.Balance <> vPPDNewBal) Then
            'ERROR if they don't balance
            vGenerateChanges = False
          ElseIf ((mvPaymentPlan.Balance <> vOPSTotal) AndAlso (vPPDChangeAmount <> vOPSTotal)) Then
            'ERROR if OPS does not balance
            Select Case pReason
              Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrBatchPosting
                'Batch Posting is renewing the Payment Plan
                'If this payment plan frequency is 1 (e.g. regualr payment, annual payment) then current OPS has already been marked as paid so assume all totals are correct
                If mvPreviousPaymentPlanBalance = 0 AndAlso mvPaymentPlan.Balance > 0 _
                   AndAlso mvPaymentPlan.PaymentFrequencyFrequency = 1 Then
                  vOPSTotal = mvPaymentPlan.Balance
                End If

              Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrRenewalsReminders
                If mvPreviousPaymentPlanBalance > 0 OrElse mvPaymentPlan.Arrears <> 0 Then
                  'Need to include any unpaid OPS that is not due to be paid yet, or is arrears that has not been included in the new years payments
                  Dim vWhereFields As New CDBFields(New CDBField("order_number", mvPaymentPlan.PlanNumber))
                  With vWhereFields
                    .Add("due_date", CDBField.FieldTypes.cftDate, If(mvPaymentPlan.RenewalPending = True, mvPaymentPlan.RenewalDate, mvPaymentPlan.CalculateRenewalDate(mvPaymentPlan.RenewalDate, False)), CDBField.FieldWhereOperators.fwoLessThan)
                    .Add("amount_outstanding", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
                    .Add("scheduled_payment_status", CDBField.FieldTypes.cftCharacter, "'D','P'", CDBField.FieldWhereOperators.fwoIn)
                  End With
                  Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "SUM(amount_outstanding)", "order_payment_schedule ops", vWhereFields)
                  Dim vSumOS As Double = DoubleValue(vSQLStatement.GetValue())
                  vOPSTotal = FixTwoPlaces(vOPSTotal + vSumOS)
                End If
            End Select
            If ((mvPaymentPlan.Balance <> vOPSTotal) AndAlso (vPPDChangeAmount <> vOPSTotal)) Then vGenerateChanges = False
          End If
        End If
        If vGenerateChanges = False Then
          RaiseError(DataAccessErrors.daeCannotCalculatePPChanges, mvPaymentPlan.PlanNumber.ToString, mvPaymentPlan.Balance.ToString("F"), vPPDNewBal.ToString("F"), vOPSTotal.ToString("F"))
        Else
          'The amount due has changed so we need to write a change record
          'First let's figure out what order detail lines have changed
          'Get a total from the existing details for each combination of product and rate (which will give a nominal code)
          Dim vExistingItems As New SortedList(Of String, PaymentPlanProductRateItem)
          If pReason = OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrRenewalsReminders Then
            'Add any existing lines with an outstanding balance only
            For Each vDetailLine As PaymentPlanDetail In mvExistingPayPlanDetails
              If vDetailLine.IsValidOnDate(pChangeDate) AndAlso vDetailLine.Balance <> 0 Then
                Dim vKey As String = vDetailLine.ProductCode & "_" & vDetailLine.RateCode
                If vExistingItems.ContainsKey(vKey) Then
                  vExistingItems(vKey).Amount = FixTwoPlaces(vExistingItems(vKey).Amount + vDetailLine.Balance)
                Else
                  vExistingItems.Add(vKey, New PaymentPlanProductRateItem(vDetailLine.ProductCode, vDetailLine.RateCode, vDetailLine.Balance))
                End If
              End If
            Next
          Else
            For Each vDetailLine As PaymentPlanDetail In mvExistingPayPlanDetails
              If vDetailLine.IsValidOnDate(pChangeDate) Then
                Dim vKey As String = vDetailLine.ProductCode & "_" & vDetailLine.RateCode
                If vExistingItems.ContainsKey(vKey) Then
                  vExistingItems(vKey).Amount = FixTwoPlaces(vExistingItems(vKey).Amount + vDetailLine.Balance)
                Else
                  vExistingItems.Add(vKey, New PaymentPlanProductRateItem(vDetailLine.ProductCode, vDetailLine.RateCode, vDetailLine.Balance))
                End If
              End If
            Next
          End If
          'Get a total from the new details for each combination of product and rate (which will give a nominal code)
          Dim vNewItems As New SortedList(Of String, PaymentPlanProductRateItem)
          For Each vDetailLine As PaymentPlanDetail In mvPaymentPlan.Details
            If (mvPaymentPlan.PaymentFrequencyFrequency = 1 And mvPaymentPlan.PaymentFrequencyInterval = 12) Or _
              vDetailLine.IsValidOnDate(pChangeDate) Then
              'BR19786 - For Annual Payments ignore the valid on date, the previous payment will be valid to the day before the renewal date, any arrears will be missing
              Dim vKey As String = vDetailLine.ProductCode & "_" & vDetailLine.RateCode
              If vNewItems.ContainsKey(vKey) Then
                vNewItems(vKey).Amount = FixTwoPlaces(vNewItems(vKey).Amount + vDetailLine.Balance)
              Else
                vNewItems.Add(vKey, New PaymentPlanProductRateItem(vDetailLine.ProductCode, vDetailLine.RateCode, vDetailLine.Balance))
              End If
            End If
          Next

          'Now delete the existing amounts from the new item amounts to find the difference
          For Each vItem As KeyValuePair(Of String, PaymentPlanProductRateItem) In vExistingItems
            If vNewItems.ContainsKey(vItem.Key) Then
              vNewItems(vItem.Key).Amount = FixTwoPlaces(vNewItems(vItem.Key).Amount - vItem.Value.Amount)
            Else
              'This existing combination no longer exists so add it as a negative amount
              vNewItems.Add(vItem.Key, vItem.Value)
              vItem.Value.Amount = -vItem.Value.Amount
            End If
          Next
          'Anything left in the new items list now with a non zero amount is a change
          Dim vCalculatedChangeAmount As Double = 0
          For Each vItem As KeyValuePair(Of String, PaymentPlanProductRateItem) In vNewItems
            vCalculatedChangeAmount = FixTwoPlaces(vCalculatedChangeAmount + vItem.Value.Amount)
            Debug.Print(String.Format("PPChange: Adding change Item {0} amount {1}", vItem.Key, vItem.Value.Amount))
          Next

          If (vCalculatedChangeAmount <> vPPDChangeAmount) And mvEnv.Connection.DatabaseAccessMode = CDBConnection.cdbDataAccessMode.damTest Then
            'As a test we will not have actually deleted anything which means the records we think are delete get selected, so just balance the figures
            vPPDChangeAmount = vCalculatedChangeAmount
          End If

          Debug.Print(String.Format("PPChange: Actual Change Amount {0} Total Change Amount {1}", vPPDChangeAmount, vCalculatedChangeAmount))
          If vCalculatedChangeAmount <> vPPDChangeAmount Then
            RaiseError(DataAccessErrors.daeCannotCalculatePPDChanges, mvPaymentPlan.PlanNumber.ToString, vPPDChangeAmount.ToString("F"), vCalculatedChangeAmount.ToString("F"))
          End If

          'The amount due has changed so we need to write a change record and as many detail records as are required
          Me.SetControlNumber()
          mvClassFields(PaymentPlanChangeFields.PaymentPlanNumber).IntegerValue = mvPaymentPlan.PlanNumber
          mvClassFields(PaymentPlanChangeFields.PaymentPlanChangeDate).Value = TodaysDate()
          mvClassFields(PaymentPlanChangeFields.Amount).DoubleValue = vCalculatedChangeAmount
          Dim vTermStartDate As Date = mvPaymentPlan.TermStartDate
          Select Case pReason
            Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrChangeMembershipType
              If mvPaymentPlan.RenewalPending = True AndAlso CDate(mvPaymentPlan.RenewalDate) > Today Then
                If mvPaymentPlan.CMTProportionBalance <> PaymentPlan.CMTProportionBalanceTypes.cmtFrequencyAmounts Then
                  vTermStartDate = CDate(mvPaymentPlan.RenewalDate)   'CMT applies from the RenewalDate
                End If
              End If
            Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrNewPaymentPlan
              vTermStartDate = CDate(mvPaymentPlan.StartDate)
            Case OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrRenewalsReminders, OrderPaymentSchedule.OrderPaymentScheduleCreationReasons.opscrBatchPosting
              vTermStartDate = CDate(mvPaymentPlan.RenewalDate)
          End Select
          mvClassFields.Item(PaymentPlanChangeFields.TermStartDate).Value = vTermStartDate.ToString(CAREDateFormat)

          'Now need to figure out how to allocate the amount between the order details lines
          Dim vLineNumber As Integer = 1
          For Each vItem As KeyValuePair(Of String, PaymentPlanProductRateItem) In vNewItems
            If vItem.Value.Amount <> 0 Then
              If mvDetails Is Nothing Then mvDetails = New List(Of PaymentPlanChangeDetail)
              Dim vChangeDetail As New PaymentPlanChangeDetail(mvEnv)
              vChangeDetail.Init()
              vChangeDetail.Create(Me, vItem.Value.Product, vItem.Value.Rate, vLineNumber, vItem.Value.Amount)
              mvDetails.Add(vChangeDetail)
              vLineNumber += 1
            End If
          Next
        End If
      End If
    End Sub

    Public Sub GenerateChangesFromBalance(ByVal pChangeDate As Date, ByVal pNegative As Boolean)
      If mvChangeRecordsRequired Then
        'Used by payment plan cancellation and reinstatement
        'Will generate a set of changes using the balance from the existing payment plan detail lines
        'Get a total from the existing details for each combination of product and rate (which will give a nominal code)
        Dim vExistingItems As New SortedList(Of String, PaymentPlanProductRateItem)
        Dim vTotalBalance As Double
        For Each vDetailLine As PaymentPlanDetail In mvExistingPayPlanDetails
          If vDetailLine.IsValidOnDate(pChangeDate) AndAlso vDetailLine.Balance > 0 Then
            Dim vKey As String = vDetailLine.ProductCode & "_" & vDetailLine.RateCode
            If vExistingItems.ContainsKey(vKey) Then
              vExistingItems(vKey).Amount = FixTwoPlaces(vExistingItems(vKey).Amount + vDetailLine.Balance)
            Else
              vExistingItems.Add(vKey, New PaymentPlanProductRateItem(vDetailLine.ProductCode, vDetailLine.RateCode, vDetailLine.Balance))
            End If
            vTotalBalance = FixTwoPlaces(vTotalBalance + vDetailLine.Balance)
          End If
        Next
        If vExistingItems.Count > 0 Then
          'There were some existing valid detail lines with an outstanding balance
          If pNegative Then vTotalBalance = -vTotalBalance
          Me.SetControlNumber()
          mvClassFields(PaymentPlanChangeFields.PaymentPlanNumber).IntegerValue = mvPaymentPlan.PlanNumber
          mvClassFields(PaymentPlanChangeFields.PaymentPlanChangeDate).Value = TodaysDate()
          mvClassFields(PaymentPlanChangeFields.Amount).DoubleValue = vTotalBalance
          mvClassFields(PaymentPlanChangeFields.TermStartDate).Value = mvPaymentPlan.TermStartDate.ToString(CAREDateFormat)

          'Now need to figure out how to allocate the amount between the order details lines
          Dim vLineNumber As Integer = 1
          For Each vItem As KeyValuePair(Of String, PaymentPlanProductRateItem) In vExistingItems
            If vItem.Value.Amount <> 0 Then
              If mvDetails Is Nothing Then mvDetails = New List(Of PaymentPlanChangeDetail)
              Dim vChangeDetail As New PaymentPlanChangeDetail(mvEnv)
              vChangeDetail.Init()
              Dim vAmount As Double = vItem.Value.Amount
              If pNegative Then vAmount = -vAmount
              vChangeDetail.Create(Me, vItem.Value.Product, vItem.Value.Rate, vLineNumber, vAmount)
              mvDetails.Add(vChangeDetail)
              vLineNumber += 1
            End If
          Next
        End If
      End If
    End Sub

    Public Sub SaveChanges()
      If mvChangeRecordsRequired AndAlso mvDetails IsNot Nothing Then
        Me.Save()
        For Each vChangeDetail As PaymentPlanChangeDetail In mvDetails
          vChangeDetail.Save()
        Next
      End If
    End Sub


#End Region


#Region "Private Classes"


    Private Class PaymentPlanProductRateItem
      Private mvProduct As String
      Private mvRate As String
      Private mvAmount As Double

      Public Sub New(ByVal pProduct As String, ByVal pRate As String, ByVal pAmount As Double)
        mvProduct = pProduct
        mvRate = pRate
        mvAmount = pAmount
      End Sub

      Public ReadOnly Property Key As String
        Get
          Return mvProduct & "_" & mvRate
        End Get
      End Property

      Public Property Amount As Double
        Get
          Return mvAmount
        End Get
        Set(ByVal value As Double)
          mvAmount = value
        End Set
      End Property

      Public ReadOnly Property Product As String
        Get
          Return mvProduct
        End Get
      End Property

      Public ReadOnly Property Rate As String
        Get
          Return mvRate
        End Get
      End Property

    End Class


#End Region

  End Class
End Namespace