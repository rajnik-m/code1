

Namespace Access
  Public Class PaymentPlanDetail

    Public Enum PaymentPlanDetailRecordSetTypes 'These are bit values
      odrtAll = &HFFS
      'ADD additional recordset types here
      odrtMain = 1
      odrtProduct = &H100S
    End Enum

    Public Enum PaymentPlanDetailTypes
      ppdltAll = &HFFFFS
      ppdltNotSet = 1
      ppdltCharge = 2
      ppdltIncentive = 4
      ppdltEntitlement = 8
      ppdltOtherCharge = 16
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PaymentPlanDetailFields
      odfAll = 0
      odfOrderNumber
      odfContactNumber
      odfAddressNumber
      odfDetailNumber
      odfProduct
      odfRate
      odfDistributionCode
      odfQuantity
      odfAmount
      odfBalance
      odfArrears
      odfDespatchMethod
      odfTimeStatus
      odfProductNumber
      odfAmendedBy
      odfAmendedOn
      odfSource
      odfCreatedBy
      odfCreatedOn
      odfCommunicationNumber
      odfEffectiveDate
      odfValidFrom
      odfValidTo
      odfNetFixedAmount
      odfModifierActivity
      odfModifierActivityValue
      odfModifierActivityQuantity
      odfModifierActivityDate
      odfModifierPrice
      odfModifierPerItem
      odfUnitPrice
      odfProRated
      odfNetAmount
      odfVatAmount
      odfGrossAmount
      odfVatRate
      odfVatPercentage
    End Enum

    Public Enum SubscriptionDataTypes
      sdtNone = 0
      sdtContactNumber = 1
      sdtAddressNumber = 2
      sdtQuantity = 4
      sdtDespatchMethod = 8
      sdtCommunicationNumber = 16
    End Enum

    Private mvDetailArrears As Boolean
    Private mvSubscription As Boolean
    Private mvDonation As Boolean
    Private mvValidFrom As String
    Private mvValidTo As String
    Private mvProductRate As ProductRate
    Private mvPaymentPlan As PaymentPlan = Nothing

    Private mvPaymentBalance As Double
    Private mvDetailType As PaymentPlanDetailTypes

    Private mvAmended As Boolean
    Private mvProduct As Product
    Private mvSubscriptionNumber As Integer
    Private mvCancellationReason As String
    Private mvAmountPaid As Double = 0

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAmendedValid As Boolean

    'Extra variables required for Trader
    Private mvSpecialInitialPeriod As Boolean 'I-Type incentive
    Private mvUsesProductNumbers As Boolean
    Private mvIgnoreProductAndRate As Boolean
    Private mvMemberOrPayer As String
    Private mvIncentiveLineType As String
    Private mvIncentiveProductDesc As String
    Private mvLineNumber As Integer
    Private mvAccruesInterest As Boolean
    Private mvLoanInterest As Boolean

    'CMT
    Private mvFullPrice As Double
    Private mvProratedPrice As Double
    Private mvCMTProrateCostCode As String
    Private mvCMTExcessPaymentTypeCode As String
    Private mvEntitlementSequenceNumber As Integer
    Private mvCMTExcessPaymentAmount As Double
    Private mvCMTRefundProductCode As String
    Private mvCMTRefundRateCode As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="PaymentPlanDetail" /> class.
    ''' </summary>
    ''' <remarks>This constructor constructs a payment detail line without reference to the containing payment plan.  This 
    ''' makes it difficult for the payment detail line to know how to behave.  This contstructor will continue to work and
    ''' the payment plan will be instantiated from the database if required, but code should construct the class passing the
    ''' payment plan where possible.</remarks>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PaymentPlanDetail" /> class.
    ''' </summary>
    ''' <param name="pPaymentPlan">The containing payment plan.</param>
    Public Sub New(ByVal pPaymentPlan As PaymentPlan)
      mvPaymentPlan = pPaymentPlan
    End Sub

    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "order_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("detail_number", CDBField.FieldTypes.cftInteger)
          .Add("product")
          .Add("rate")
          .Add("distribution_code")
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBTAQuantityDecimal) Then
            .Add("quantity", CDBField.FieldTypes.cftNumeric)
          Else
            .Add("quantity", CDBField.FieldTypes.cftInteger)
          End If
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("balance", CDBField.FieldTypes.cftNumeric)
          .Add("arrears", CDBField.FieldTypes.cftNumeric)
          .Add("despatch_method")
          .Add("time_status")
          .Add("product_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("source")
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
          .Add("communication_number", CDBField.FieldTypes.cftLong)
          .Add("effective_date", CDBField.FieldTypes.cftDate)
          .Add("valid_from", CDBField.FieldTypes.cftDate)
          .Add("valid_to", CDBField.FieldTypes.cftDate)
          .Add("net_fixed_amount", CDBField.FieldTypes.cftNumeric)
          .Add("modifier_activity")
          .Add("modifier_activity_value")
          .Add("modifier_activity_quantity", CDBField.FieldTypes.cftNumeric)
          .Add("modifier_activity_date", CDBField.FieldTypes.cftDate)
          .Add("modifier_price", CDBField.FieldTypes.cftNumeric)
          .Add("modifier_per_item")
          .Add("unit_price", CDBField.FieldTypes.cftNumeric)
          .Add("pro_rated")
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
          .Add("vat_amount", CDBField.FieldTypes.cftNumeric).PrefixRequired = True
          .Add("gross_amount", CDBField.FieldTypes.cftNumeric).PrefixRequired = True
          .Add("vat_rate").PrefixRequired = True
          .Add("vat_percentage", CDBField.FieldTypes.cftNumeric).PrefixRequired = True
        End With

        mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedBy).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanDetailCreatedBy)
        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedOn).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanDetailCreatedBy)
        mvClassFields.Item(PaymentPlanDetailFields.odfCommunicationNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber)
        mvClassFields.Item(PaymentPlanDetailFields.odfEffectiveDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPPDetailsEffectiveDate)
        mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage)
        mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage)
        mvClassFields.Item(PaymentPlanDetailFields.odfNetFixedAmount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlansVatExcl)

        Dim vGotAttr As Boolean = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails)
        With mvClassFields
          .Item(PaymentPlanDetailFields.odfModifierActivity).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfModifierActivityValue).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfModifierActivityQuantity).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfModifierActivityDate).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfModifierPrice).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfModifierPerItem).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfUnitPrice).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfProRated).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfNetAmount).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfVatAmount).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfGrossAmount).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfVatPercentage).InDatabase = vGotAttr
          .Item(PaymentPlanDetailFields.odfVatRate).InDatabase = vGotAttr
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvAmendedValid = False
      mvExisting = False
      mvProduct = Nothing
      mvProductRate = New ProductRate(mvEnv)
      mvProductRate.Init()
      mvSpecialInitialPeriod = False
      mvUsesProductNumbers = False
      mvIgnoreProductAndRate = False
      mvMemberOrPayer = ""
      mvIncentiveLineType = ""
      mvIncentiveProductDesc = ""
      mvAccruesInterest = False
      mvLoanInterest = False
      mvFullPrice = 0
      mvProratedPrice = 0
      mvCMTProrateCostCode = ""
      mvCMTExcessPaymentTypeCode = ""
      mvEntitlementSequenceNumber = 0
      mvCMTExcessPaymentAmount = 0
      mvCMTRefundProductCode = ""
      mvCMTRefundRateCode = ""

    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(PaymentPlanDetailFields.odfTimeStatus).Value = "C"
      mvClassFields.Item(PaymentPlanDetailFields.odfArrears).DoubleValue = 0

      mvDetailType = PaymentPlanDetailTypes.ppdltNotSet
      mvDetailArrears = False
      mvSubscription = False
      mvValidFrom = ""
      mvValidTo = ""
      mvDonation = False
      mvProductRate = New ProductRate(mvEnv)
      mvProductRate.Init()
      mvPaymentBalance = 0
      mvAmended = False
      mvExisting = False
    End Sub

    Private Sub SetValid(ByRef pField As PaymentPlanDetailFields)
      'Add code here to ensure all values are valid before saving
      If Not mvAmendedValid Then
        If Len(mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value) = 0 Then mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value = TodaysDate()
        If Len(mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value) = 0 Then mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value = mvEnv.User.UserID
      End If

      If (pField = PaymentPlanDetailFields.odfAll) And Len(mvClassFields(PaymentPlanDetailFields.odfCreatedBy).Value) = 0 And mvExisting = False Then
        'Only set created by/on for new detail lines
        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedBy).Value = mvEnv.User.UserID
        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedOn).Value = TodaysDate()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PaymentPlanDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And PaymentPlanDetailRecordSetTypes.odrtMain) > 0 Then
        vFields = "od.order_number AS od_order_number,"
        vFields = vFields & "od.contact_number AS od_contact_number,"
        vFields = vFields & "od.address_number AS od_address_number,"
        vFields = vFields & "od.product AS od_product,"
        vFields = vFields & "detail_number,od.rate,od.quantity,"
        vFields = vFields & "od.amount AS od_amount,"
        vFields = vFields & "od.balance AS od_balance,"
        vFields = vFields & "od.arrears AS od_arrears,"
        vFields = vFields & "od.despatch_method,"
        If Not (pRSType And PaymentPlanDetailRecordSetTypes.odrtProduct) = PaymentPlanDetailRecordSetTypes.odrtProduct Then vFields = vFields & "p.subscription,p.donation,"
        vFields = vFields & "time_status,"
        vFields = vFields & "od.product_number AS od_product_number,uses_product_numbers,"
        vFields = vFields & "od.distribution_code AS od_distribution_code,"
        vFields = vFields & "od.source AS od_source,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlanDetailCreatedBy) Then vFields = vFields & "od.created_by AS od_created_by, od.created_on AS od_created_on,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) Then vFields = vFields & "od.communication_number AS od_communication_number,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPPDetailsEffectiveDate) Then vFields = vFields & "od.effective_date AS od_effective_date,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPayPlansVatExcl) Then vFields &= "net_fixed_amount,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then vFields &= "od.valid_from,od.valid_to,"
        vFields = vFields & "od.amended_on AS od_amended_on,"
        vFields = vFields & "od.amended_by AS od_amended_by,"
        vFields = vFields & "r.current_price,r.future_price,r.price_change_date,r.vat_exclusive,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then vFields = vFields & "r.price_is_percentage,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoans) Then
          If (pRSType And PaymentPlanDetailRecordSetTypes.odrtProduct) = 0 Then vFields &= "p.accrues_interest,"
          vFields &= "r.loan_interest,"
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifier) Then vFields &= "r.use_modifiers,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPaymentPlanHistoryDetails) Then
          vFields &= "modifier_activity, modifier_activity_value, modifier_activity_quantity, modifier_activity_date, modifier_price, modifier_per_item,"
          vFields &= " unit_price, pro_rated, net_amount, od.vat_amount, od.gross_amount, od.vat_rate, od.vat_percentage,"
        End If
      End If
      If (pRSType And PaymentPlanDetailRecordSetTypes.odrtProduct) > 0 Then
        If mvProduct Is Nothing Then mvProduct = New Product(mvEnv)
        mvProduct.Init()
        vFields = vFields & mvProduct.GetRecordSetFields(Product.ProductRecordSetTypes.prstMain)
      End If
      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pPlanNumber As Integer = 0, Optional ByRef pDetailNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pPlanNumber > 0 And pDetailNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PaymentPlanDetailRecordSetTypes.odrtAll) & " FROM order_details od,products p,rates r WHERE o.order_number = " & pPlanNumber & " AND detail_number = " & pDetailNumber & " AND p.product = od.product AND r.product = od.product and r.rate = od.rate ORDER BY detail_number")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PaymentPlanDetailRecordSetTypes.odrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PaymentPlanDetailRecordSetTypes)
      mvEnv = pEnv
      InitClassFields()
      mvExisting = True
      'Always include the primary key attributes
      mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).SetValue = CStr(pRecordSet.Fields("od_order_number").IntegerValue)
      mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).SetValue = pRecordSet.Fields("detail_number").Value
      'Modify below to handle each recordset type as required
      If (pRSType And PaymentPlanDetailRecordSetTypes.odrtMain) > 0 Then
        mvClassFields.Item(PaymentPlanDetailFields.odfContactNumber).SetValue = CStr(pRecordSet.Fields("od_contact_number").IntegerValue)
        mvClassFields.Item(PaymentPlanDetailFields.odfAddressNumber).SetValue = CStr(pRecordSet.Fields("od_address_number").IntegerValue)
        mvClassFields.Item(PaymentPlanDetailFields.odfProduct).SetValue = pRecordSet.Fields("od_product").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfRate).SetValue = pRecordSet.Fields("rate").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfDistributionCode).SetValue = pRecordSet.Fields.FieldExists("od_distribution_code").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).SetValue = pRecordSet.Fields("quantity").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfAmount).SetValue = pRecordSet.Fields("od_amount").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfBalance).SetValue = CStr(pRecordSet.Fields("od_balance").DoubleValue)
        mvClassFields.Item(PaymentPlanDetailFields.odfArrears).SetValue = CStr(pRecordSet.Fields("od_arrears").DoubleValue)
        mvClassFields.Item(PaymentPlanDetailFields.odfDespatchMethod).SetValue = pRecordSet.Fields("despatch_method").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfTimeStatus).SetValue = pRecordSet.Fields.FieldExists("time_status").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfProductNumber).SetValue = pRecordSet.Fields.FieldExists("od_product_number").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfSource).SetValue = pRecordSet.Fields.FieldExists("od_source").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedBy).SetValue = pRecordSet.Fields.FieldExists("od_created_by").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfCreatedOn).SetValue = pRecordSet.Fields.FieldExists("od_created_on").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfCommunicationNumber).SetValue = pRecordSet.Fields.FieldExists("od_communication_number").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfEffectiveDate).SetValue = pRecordSet.Fields.FieldExists("od_effective_date").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfNetFixedAmount).SetValue = pRecordSet.Fields.FieldExists("net_fixed_amount").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).SetValue = pRecordSet.Fields("od_amended_by").Value
        mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).SetValue = pRecordSet.Fields("od_amended_on").Value
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfValidFrom, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfValidTo, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfNetFixedAmount, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierActivity, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierActivityValue, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierActivityQuantity, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierActivityDate, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierPrice, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfModifierPerItem, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfUnitPrice, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfProRated, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfNetAmount, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfVatAmount, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfGrossAmount, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfVatRate, pRecordSet.Fields)
        mvClassFields.SetOptionalItem(PaymentPlanDetailFields.odfVatPercentage, pRecordSet.Fields)

        mvSubscription = pRecordSet.Fields("subscription").Bool
        mvDonation = pRecordSet.Fields("donation").Bool
        Dim vPriceIsPercentage As String = "N"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPriceIsPercentage) Then vPriceIsPercentage = pRecordSet.Fields("price_is_percentage").Value
        mvAccruesInterest = False
        mvLoanInterest = False
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoans) = True AndAlso pRecordSet.Fields.ContainsKey("loan_interest") = True Then
          mvAccruesInterest = pRecordSet.Fields("accrues_interest").Bool
          mvLoanInterest = pRecordSet.Fields("loan_interest").Bool
        End If
        mvProductRate.InitForPrice(ProductCode, RateCode, pRecordSet.Fields("current_price").DoubleValue, pRecordSet.Fields("future_price").DoubleValue, pRecordSet.Fields("price_change_date").Value, pRecordSet.Fields.FieldExists("use_modifiers").Value, vPriceIsPercentage, pRecordSet.Fields("vat_exclusive").Bool)
        If (pRSType And PaymentPlanDetailRecordSetTypes.odrtProduct) = 0 Then 'Temp fix to not do this in batch processing
          If mvEnv.Connection.GetCount("order_detail_arrears", Nothing, "order_number = " & mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).Value & " AND detail_number = " & mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).Value) > 0 Then mvDetailArrears = True
        End If
        mvDetailType = PaymentPlanDetailTypes.ppdltNotSet
        mvAmended = False
      End If
      If (pRSType And PaymentPlanDetailRecordSetTypes.odrtProduct) > 0 Then
        If mvProduct Is Nothing Then mvProduct = New Product(mvEnv)
        mvProduct.InitFromRecordSet(pRecordSet, Product.ProductRecordSetTypes.prstMain)
      End If

      'Extra bits required for Trader
      If mvClassFields.Item(PaymentPlanDetailFields.odfProductNumber).Value.Length > 0 Then mvUsesProductNumbers = True

    End Sub

    Public Sub Save()
      Dim vArrearsFields As New CDBFields

      SetValid(PaymentPlanDetailFields.odfAll)
      'Assume Pay Plan Details are always deleted on a PayPlan.Save
      'So we must always insert them again
      mvClassFields.ClearSetValues()
      mvEnv.Connection.InsertRecord("order_details", mvClassFields.UpdateFields)
      mvAmended = False

      If mvDetailArrears Then
        'insert the order_details_arrears record
        vArrearsFields.Add("order_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).Value)
        vArrearsFields.Add("detail_number", CDBField.FieldTypes.cftInteger, mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).Value)
        mvEnv.Connection.InsertRecord("order_detail_arrears", vArrearsFields)
      End If
    End Sub

    Public Sub Delete()
      If mvExisting Then mvEnv.Connection.DeleteRecords("order_details", mvClassFields.WhereFields)
    End Sub

    Public Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PaymentPlanDetailFields.odfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub SetDetailArrears(ByVal pRenewalDate As String)
      If mvDetailArrears = False AndAlso Arrears <> 0 Then
        mvClassFields.Item(PaymentPlanDetailFields.odfBalance).Value = mvClassFields.Item(PaymentPlanDetailFields.odfArrears).Value
        If IsDate(pRenewalDate) Then
          SetCMTLineNoRenewalRequired(pRenewalDate)
        Else
          mvClassFields.Item(PaymentPlanDetailFields.odfAmount).Value = "0"
        End If
        mvDetailArrears = True
      End If
    End Sub

    Friend Sub SetAmount(ByVal pAmount As String)
      If pAmount.Length > 0 Then
        mvClassFields(PaymentPlanDetailFields.odfAmount).Value = CStr(Val(pAmount))
      Else
        mvClassFields(PaymentPlanDetailFields.odfAmount).Value = ""
      End If
      mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    Public Sub SetCreated(ByRef pCreatedOn As String, ByRef pCreatedBy As String)
      mvClassFields.Item(PaymentPlanDetailFields.odfCreatedOn).Value = pCreatedOn
      mvClassFields.Item(PaymentPlanDetailFields.odfCreatedBy).Value = pCreatedBy
    End Sub

    Public Sub Create(ByVal pPlanNumber As Integer, ByVal pDetailNumber As Integer, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Double, ByVal pBalance As Double, Optional ByVal pGrossFixedAmount As String = "", Optional ByVal pArrears As Double = 0, Optional ByVal pDespatchMethod As String = "", Optional ByVal pSource As String = "", Optional ByVal pDistributionCode As String = "", Optional ByVal pCreatedBy As String = "", Optional ByVal pCreatedOn As String = "", Optional ByVal pTimeStatus As String = "C", Optional ByVal pProductNumber As String = "", Optional ByVal pCommunicationNumber As String = "", Optional ByVal pSubscription As Boolean = False, Optional ByVal pDetailArrears As Boolean = False, Optional ByVal pMemberOrPayer As String = "", Optional ByVal pUsesProductNumbers As Boolean = False, Optional ByVal pIncentiveLineType As String = "", Optional ByVal pIncentiveProductDesc As String = "", Optional ByVal pMProd As Boolean = False, Optional ByVal pNetFixedAmount As String = "")
      'Create a new PaymentPlanDetail
      'Parameters after pDetailArrears are only used by Trader
      'pMProd is used by Smart Client to identify Membership products

      With mvClassFields
        .Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue = pPlanNumber
        .Item(PaymentPlanDetailFields.odfDetailNumber).IntegerValue = pDetailNumber
        .Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue = pContactNumber
        .Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue = pAddressNumber
        .Item(PaymentPlanDetailFields.odfProduct).Value = pProductCode
        .Item(PaymentPlanDetailFields.odfRate).Value = pRate
        .Item(PaymentPlanDetailFields.odfQuantity).DoubleValue = pQuantity
        .Item(PaymentPlanDetailFields.odfBalance).DoubleValue = pBalance
        .Item(PaymentPlanDetailFields.odfArrears).DoubleValue = pArrears
        .Item(PaymentPlanDetailFields.odfTimeStatus).Value = pTimeStatus
        If pGrossFixedAmount.Length > 0 Then .Item(PaymentPlanDetailFields.odfAmount).Value = pGrossFixedAmount
        If pDespatchMethod.Length > 0 Then .Item(PaymentPlanDetailFields.odfDespatchMethod).Value = pDespatchMethod
        If pSource.Length > 0 Then .Item(PaymentPlanDetailFields.odfSource).Value = pSource
        If pDistributionCode.Length > 0 Then .Item(PaymentPlanDetailFields.odfDistributionCode).Value = pDistributionCode
        If pCreatedBy.Length > 0 Then
          .Item(PaymentPlanDetailFields.odfCreatedBy).Value = pCreatedBy
          .Item(PaymentPlanDetailFields.odfCreatedOn).Value = pCreatedOn
        End If
        If pProductNumber.Length > 0 Then .Item(PaymentPlanDetailFields.odfProductNumber).Value = pProductNumber
        If pCommunicationNumber.Length > 0 Then .Item(PaymentPlanDetailFields.odfCommunicationNumber).Value = pCommunicationNumber
        If pNetFixedAmount.Length > 0 Then .Item(PaymentPlanDetailFields.odfNetFixedAmount).Value = pNetFixedAmount
      End With
      mvSubscription = pSubscription
      mvDetailArrears = pDetailArrears

      'Extra fields required by Trader
      mvMemberOrPayer = pMemberOrPayer
      mvUsesProductNumbers = pUsesProductNumbers
      If pIncentiveLineType.Length > 0 Then
        mvIncentiveLineType = pIncentiveLineType
        mvSpecialInitialPeriod = (Mid(pIncentiveLineType, 2, 1) = "I") 'Initial Period Incentive
        mvIncentiveProductDesc = pIncentiveProductDesc

        If HasPriceInfo = False Then
          'This is an incentive line and as the price has not come directly from the ProductRate class we have no pricing
          mvProductRate = New ProductRate(mvEnv)
          mvProductRate.Init(pProductCode, pRate)
          Dim vPrice As Double = mvProductRate.Price(pContactNumber)
          SetModifierAndPriceData(mvProductRate.PaymentPlanDetailPricing)
          SetIncentivesModifierAndPriceData(False)
        End If

      End If
      If pMProd Then mvDetailType = PaymentPlanDetailTypes.ppdltCharge

    End Sub

    Public Sub CreateSC(ByRef pParams As CDBParameters)
      'Create a new PaymentPlanDetail (Smart Client & Web Services only)
      'Parameters after pDetailArrears are only used by Trader
      Dim vProduct As New Product(mvEnv)
      With mvClassFields
        .Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue = pParams.ParameterExists("PaymentPlanNumber").IntegerValue
        .Item(PaymentPlanDetailFields.odfDetailNumber).IntegerValue = pParams.ParameterExists("DetailNumber").IntegerValue
        .Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue = pParams("ContactNumber").IntegerValue
        .Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue = pParams("AddressNumber").IntegerValue
        .Item(PaymentPlanDetailFields.odfProduct).Value = pParams("Product").Value
        .Item(PaymentPlanDetailFields.odfRate).Value = pParams("Rate").Value
        .Item(PaymentPlanDetailFields.odfQuantity).DoubleValue = Val(pParams.OptionalValue("Quantity", "1"))
        .Item(PaymentPlanDetailFields.odfBalance).DoubleValue = pParams("Balance").DoubleValue
        .Item(PaymentPlanDetailFields.odfArrears).DoubleValue = pParams.ParameterExists("Arrears").DoubleValue
        If pParams.ParameterExists("DetailFixedAmount").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfAmount).Value = pParams("DetailFixedAmount").Value
        If pParams.ParameterExists("DespatchMethod").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfDespatchMethod).Value = pParams("DespatchMethod").Value
        If pParams.ParameterExists("Source").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfSource).Value = pParams("Source").Value
        If pParams.ParameterExists("DistributionCode").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfDistributionCode).Value = pParams("DistributionCode").Value
        If pParams.ParameterExists("CommunicationNumber").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfCommunicationNumber).Value = pParams("CommunicationNumber").Value
        If pParams.ParameterExists("EffectiveDate").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfEffectiveDate).Value = pParams("EffectiveDate").Value
        .Item(PaymentPlanDetailFields.odfValidFrom).Value = pParams.ParameterExists("ValidFrom").Value
        .Item(PaymentPlanDetailFields.odfValidTo).Value = pParams.ParameterExists("ValidTo").Value
        If pParams.ParameterExists("NetFixedAmount").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfNetFixedAmount).Value = pParams("NetFixedAmount").Value
        If pParams.Exists("ModifierActivity") Then
          .Item(PaymentPlanDetailFields.odfModifierActivity).Value = pParams("ModifierActivity").Value
          If pParams.ParameterExists("ModifierActivityValue").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfModifierActivityValue).Value = pParams("ModifierActivityValue").Value
          If pParams.ParameterExists("ModifierActivityQuantity").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfModifierActivityQuantity).DoubleValue = pParams("ModifierActivityQuantity").DoubleValue
          If pParams.ParameterExists("ModifierActivityDate").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfModifierActivityDate).Value = pParams("ModifierActivityDate").Value
          If pParams.ParameterExists("ModifierPrice").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfModifierPrice).DoubleValue = pParams("ModifierPrice").DoubleValue
          If pParams.ParameterExists("ModifierPerItem").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfModifierPerItem).Value = pParams("ModifierPerItem").Value
        End If
        If pParams.ParameterExists("UnitPrice").Value.Length > 0 Then
          .Item(PaymentPlanDetailFields.odfUnitPrice).DoubleValue = pParams("UnitPrice").DoubleValue
          .Item(PaymentPlanDetailFields.odfProRated).Bool = BooleanValue(pParams.OptionalValue("ProRated", "N"))
          .Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = DoubleValue(pParams.OptionalValue("NetAmount", "0.00"))
          .Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = DoubleValue(pParams.OptionalValue("VatAmount", "0.00"))
          .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = DoubleValue(pParams.OptionalValue("GrossAmount", "0.00"))
          If pParams.ParameterExists("VatRate").Value.Length > 0 Then .Item(PaymentPlanDetailFields.odfVatRate).Value = pParams("VatRate").Value
          .Item(PaymentPlanDetailFields.odfVatPercentage).DoubleValue = DoubleValue(pParams.OptionalValue("VatPercentage", "0"))
        End If
      End With
      SetCreated(pParams.OptionalValue("CreatedOn", (TodaysDate())), pParams.OptionalValue("CreatedBy", mvEnv.User.UserID))
      If pParams.Exists("AmendedBy") Then SetAmended(pParams("AmendedOn").Value, pParams("AmendedBy").Value) 'Probably only set by Import
      'odfTimeStatus is always set to 'C'

      vProduct.InitWithRate(mvEnv, pParams("Product").Value, pParams("Rate").Value)
      mvSubscription = vProduct.Subscription
      mvUsesProductNumbers = vProduct.UsesProductNumbers
      mvProductRate = vProduct.ProductRate

      If mvUsesProductNumbers Then
        If pParams.ParameterExists("ProductNumber").Value.Length > 0 Then mvClassFields.Item(PaymentPlanDetailFields.odfProductNumber).Value = pParams("ProductNumber").Value
      End If

      If mvSubscription Then
        If pParams.ParameterExists("SubscriptionNumber").IntegerValue > 0 Then mvSubscriptionNumber = pParams("SubscriptionNumber").IntegerValue
      End If

      'mvDetailArrears = pDetailArrears

      'Extra fields required by Trader
      mvMemberOrPayer = pParams.ParameterExists("MemberOrPayer").Value
      If pParams.ParameterExists("IncentiveLineType").Value.Length > 0 Then
        mvIncentiveLineType = pParams("IncentiveLineType").Value
        If mvIncentiveLineType.Length > 1 Then mvSpecialInitialPeriod = (mvIncentiveLineType.Substring(1, 1).ToUpper = "I")
      End If
      If pParams.ParameterExists("IncentiveIgnoreProductAndRate").Value.Length > 0 Then mvIgnoreProductAndRate = pParams("IncentiveIgnoreProductAndRate").Bool
      'If Len(pIncentiveLineType) Then
      '  mvIncentiveLineType = pIncentiveLineType
      '  mvSpecialInitialPeriod = (Mid$(pIncentiveLineType, 2, 1) = "I")   'Initial Period Incentive
      '  mvIncentiveProductDesc = pIncentiveProductDesc
      'End If
      mvAccruesInterest = vProduct.AccruesInterest
      mvLoanInterest = vProduct.ProductRate.LoanInterest

      If HasPriceInfo = False Then
        Dim vPPDPricing As New PaymentPlanDetailPricing(mvEnv)
        vPPDPricing.Init()
        vPPDPricing.CalculatePricing(If(String.IsNullOrWhiteSpace(Amount), Balance, DoubleValue(Amount)), If(String.IsNullOrWhiteSpace(Amount), Balance, DoubleValue(Amount)), ProductRate.VatExclusive, pParams.ParameterExists("EffectiveDate").Value, mvEnv.VATRate(vProduct.ProductVatCategory, Me.PaymentPlan.Payer.VATCategory), False)
        SetModifierAndPriceData(vPPDPricing)
      End If
    End Sub

    Public Sub Update(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pQuantity As Double, ByVal pBalance As Double, Optional ByVal pAmount As String = "", Optional ByVal pArrears As String = "", Optional ByVal pProductCode As String = "", Optional ByVal pRate As String = "", Optional ByVal pDespatchMethod As String = "", Optional ByVal pSource As String = "", Optional ByVal pDistributionCode As String = "", Optional ByVal pCommunicationNumber As String = "", Optional ByVal pSubscription As Boolean = False, Optional ByVal pMemberOrPayer As String = "", Optional ByVal pIncentiveLineType As String = "", Optional ByVal pIncentiveProductDesc As String = "", Optional ByVal pIncentiveIgnoreProductAndRate As Boolean = False, Optional ByVal pNetFixedAmount As String = "")
      'Parameters after pSubscription are only used by Trader
      With mvClassFields
        .Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue = pContactNumber
        .Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue = pAddressNumber
        .Item(PaymentPlanDetailFields.odfQuantity).DoubleValue = pQuantity
        .Item(PaymentPlanDetailFields.odfBalance).DoubleValue = pBalance
        If pArrears.Length > 0 Then .Item(PaymentPlanDetailFields.odfArrears).DoubleValue = CDbl(pArrears)
        .Item(PaymentPlanDetailFields.odfAmount).Value = pAmount
        .Item(PaymentPlanDetailFields.odfNetFixedAmount).Value = pNetFixedAmount
        If pProductCode.Length > 0 Then
          .Item(PaymentPlanDetailFields.odfProduct).Value = pProductCode
          mvProduct = Nothing
        End If
        If pRate.Length > 0 Then .Item(PaymentPlanDetailFields.odfRate).Value = pRate
        If pDespatchMethod.Length > 0 Then .Item(PaymentPlanDetailFields.odfDespatchMethod).Value = pDespatchMethod
        If pSource.Length > 0 Then .Item(PaymentPlanDetailFields.odfSource).Value = pSource
        If pDistributionCode.Length > 0 Then .Item(PaymentPlanDetailFields.odfDistributionCode).Value = pDistributionCode
        If pCommunicationNumber.Length > 0 Then .Item(PaymentPlanDetailFields.odfCommunicationNumber).Value = pCommunicationNumber
        If (.Item(PaymentPlanDetailFields.odfProduct).ValueChanged = True Or .Item(PaymentPlanDetailFields.odfRate).ValueChanged = True) And mvProductRate.IsValid Then
          'Product or Rate has changed so re-select pricing information as it was previously set and could now be wrong
          Dim vAttrs As String = "donation, current_price, future_price, price_change_date"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifier) Then vAttrs &= ", use_modifiers"
          Dim vAnsiJoins As New AnsiJoins
          vAnsiJoins.Add("rates r", "p.product", "r.product")
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "products p", New CDBFields(New CDBField("r.rate", RateCode)), "", vAnsiJoins)
          Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          If vRS.Fetch() = True Then
            With vRS
              mvDonation = .Fields("donation").Bool
              mvProductRate.InitForPrice(ProductCode, RateCode, .Fields("current_price").DoubleValue, .Fields("future_price").DoubleValue, .Fields("price_change_date").Value, .Fields.FieldExists("use_modifiers").Value)
            End With
          End If
          vRS.CloseRecordSet()
        End If
      End With
      mvSubscription = pSubscription

      'Extra fields required by Trader
      mvMemberOrPayer = pMemberOrPayer
      If pIncentiveLineType.Length > 0 Then
        'Incentives
        mvIncentiveLineType = pIncentiveLineType
        mvSpecialInitialPeriod = (Mid(pIncentiveLineType, 2, 1) = "I")
        mvIgnoreProductAndRate = pIncentiveIgnoreProductAndRate
        mvIncentiveProductDesc = pIncentiveProductDesc
        SetIncentivesModifierAndPriceData(pIncentiveIgnoreProductAndRate)
      End If

    End Sub

    Friend Sub SetPaymentPlanAndDetailNumbers(ByVal pPlanNumber As Integer, ByVal pDetailNumber As Integer)
      'This will set the correct numbers
      mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue = pPlanNumber
      mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).IntegerValue = pDetailNumber
    End Sub

    Public Sub SetSubscriptionValidFromTo(ByVal pValidFrom As String, ByVal pValidTo As String)
      If mvSubscription Then
        mvValidFrom = pValidFrom
        mvValidTo = pValidTo
      End If
    End Sub

    Public Sub SetContactAndAddress(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer)
      'Reset the Contact and Address Numbers
      mvClassFields.Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue = pContactNumber
      mvClassFields.Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue = pAddressNumber
    End Sub

    Public Sub ApplyDiscount(ByVal pContactDiscount As Boolean, ByVal pDiscountPercentage As Double)
      Dim vDiscount As Double

      If pContactDiscount Then
        vDiscount = FixTwoPlaces(mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue * (pDiscountPercentage / 100))
        mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = FixTwoPlaces(mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue - vDiscount)
      End If

    End Sub

    Friend Sub ResetIncentiveQuantity()
      'Reset quantity to 1 for Special Initial Period Incentive
      If mvSpecialInitialPeriod Then mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).IntegerValue = 1
    End Sub

    Friend Sub ProcessSubscriptionValidDates(ByVal pPP As PaymentPlan, ByVal pPPMaintenance As Boolean, ByVal pCMT As Boolean, Optional ByVal pSubsStartOnMemberJoined As Boolean = False, Optional ByVal pMemberJoinedDate As String = "")
      'Used to calculate the Subscription ValidFrom & ValidTo dates
      Dim vDate As String
      Dim vToDate As String

      vDate = ""
      vToDate = ""
      If mvSubscription = True And (mvExisting = False Or (mvExisting = True And Len(mvValidFrom) = 0)) Then
        'Only need to re-calculate the ValidFrom & ValidTo if not already set
        If pSubsStartOnMemberJoined Then
          'Set vDate as the appropriate base date on which to calculate vToDate
          If pPP.FixedRenewalCycle And Not pPP.PreviousRenewalCycle Then
            vDate = pPP.StartDate
          Else
            vDate = pMemberJoinedDate
          End If
          If pPP.SubsExtension Then
            vToDate = CDate(vDate).AddYears(99).ToString(CAREDateFormat)
          Else
            vToDate = vDate
            If pCMT Then vToDate = pPP.RenewalDate
            vToDate = CDate(vToDate).AddMonths(pPP.MembershipType.SuspensionGrace).ToString(CAREDateFormat)
          End If
          'Set vDate as the From Date that should be used
          vDate = pMemberJoinedDate
        Else
          If pPPMaintenance Then
            vDate = TodaysDate()
          Else
            vDate = pPP.StartDate
          End If
          If pPP.SubsExtension Then
            vToDate = CDate(vDate).AddYears(99).ToString(CAREDateFormat)
          Else
            If pPPMaintenance Then
              vToDate = pPP.RenewalDate
            Else
              vToDate = CDate(vDate).AddYears(pPP.Term).ToString(CAREDateFormat)
            End If
          End If
        End If
        'Take 1 day off
        vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, -1, CDate(vToDate)))
        'If the calculated vToDate is actually less than vDate then set vToDate = vDate
        If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vDate), CDate(vToDate)) < 0 Then vToDate = vDate
        SetSubscriptionValidFromTo(vDate, vToDate)
      End If

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(PaymentPlanDetailFields.odfAddressNumber).IntegerValue = Value
      End Set
    End Property

    Public Property Amended() As Boolean
      Get
        Dim vUpdateFields As CDBFields

        vUpdateFields = mvClassFields.UpdateFields
        If vUpdateFields.Count > 0 Then
          mvAmended = True
        Else
          mvAmended = False
        End If
        Amended = mvAmended
      End Get
      Set(ByVal Value As Boolean)
        mvAmended = Value
      End Set
    End Property
    Public ReadOnly Property FinancialAmended() As Boolean
      Get
        Dim vUpdateFields As CDBFields
        Dim vUpdateField As CDBField

        vUpdateFields = mvClassFields.UpdateFields
        If vUpdateFields.Count > 0 Then
          For Each vUpdateField In vUpdateFields
            Select Case vUpdateField.Name
              Case "product", "rate", "quantity", "balance", "amount", "arrears"
                FinancialAmended = True
            End Select
          Next vUpdateField
          mvAmended = True
        End If
      End Get
    End Property

    Public Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value = Value
      End Set
    End Property

    Public Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value = Value
      End Set
    End Property

    ''' <summary>Fixed Amount including VAT (i.e. Gross)</summary>
    Public Property Amount() As String
      Get
        Amount = mvClassFields.Item(PaymentPlanDetailFields.odfAmount).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfAmount).Value = Value
      End Set
    End Property

    Public Property Arrears() As Double
      Get
        Arrears = mvClassFields.Item(PaymentPlanDetailFields.odfArrears).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(PaymentPlanDetailFields.odfArrears).DoubleValue = Value
      End Set
    End Property

    Public Property Balance() As Double
      Get
        Balance = mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = Value
      End Set
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(PaymentPlanDetailFields.odfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public Property DespatchMethod() As String
      Get
        DespatchMethod = mvClassFields.Item(PaymentPlanDetailFields.odfDespatchMethod).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfDespatchMethod).Value = Value
      End Set
    End Property

    Public ReadOnly Property DetailArrears() As Boolean
      Get
        DetailArrears = mvDetailArrears
      End Get
    End Property

    Public Property DetailNumber() As Integer
      Get
        DetailNumber = mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(PaymentPlanDetailFields.odfDetailNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property DetailType() As PaymentPlanDetailTypes
      Get
        DetailType = mvDetailType
      End Get
      Set(ByVal Value As PaymentPlanDetailTypes)
        mvDetailType = Value
      End Set
    End Property

    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvClassFields.Item(PaymentPlanDetailFields.odfDistributionCode).Value
      End Get
    End Property

    Public ReadOnly Property PriceIsPercentage() As String
      Get
        Return mvProductRate.PriceIsPercentage
      End Get
    End Property

    Public ReadOnly Property ValidFrom() As String
      Get
        'Could be null
        Return mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value
      End Get
    End Property

    Public Sub SetValidFrom(pValidFrom As Date)
      mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value = pValidFrom.ToString(CAREDateFormat)
    End Sub

    Public ReadOnly Property ValidTo() As String
      Get
        'Could be null
        Return mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value
      End Get
    End Property

    Public Sub SetValidTo(pValidTo As Date)
      mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value = pValidTo.ToString(CAREDateFormat)
    End Sub

    Public ReadOnly Property EffectiveDate() As String
      Get
        'Could be null
        EffectiveDate = mvClassFields.Item(PaymentPlanDetailFields.odfEffectiveDate).Value
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property Donation() As Boolean
      Get
        Donation = mvDonation
      End Get
    End Property

    Public Property PaymentBalance() As Double
      Get
        PaymentBalance = mvPaymentBalance
      End Get
      Set(ByVal Value As Double)
        mvPaymentBalance = Value
      End Set
    End Property

    Public Property PlanNumber() As Integer
      Get
        PlanNumber = mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property ProductRateIsValid() As Boolean
      Get
        Return mvProductRate.IsValid
      End Get
    End Property

    Public ReadOnly Property PriceIsZero() As Boolean
      Get
        Return mvProductRate.PriceIsZero
      End Get
    End Property

    Public ReadOnly Property CurrentPrice() As Double
      Get
        Dim vCurrentPrice As Double
        Dim vValid As Boolean = True

        If mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value.Length > 0 AndAlso Today < CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value) Then vValid = False
        If vValid AndAlso (mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value.Length > 0 AndAlso Today > CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value)) Then vValid = False
        If vValid Then vCurrentPrice = mvProductRate.Price(ContactNumber)

        Return vCurrentPrice
      End Get
    End Property

    Public ReadOnly Property RenewalPrice(ByVal pFuture As Boolean, ByVal pVATRate As VatRate, ByVal pTransactionDate As Date) As Double
      Get
        Dim vRenewalPrice As Double
        Dim vValid As Boolean = True
        Dim vDate As Date = Today

        If pFuture Then vDate = ProductRate.PriceChangeDate
        If mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value.Length > 0 AndAlso vDate < CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value) Then vValid = False
        If vValid AndAlso (mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value.Length > 0 AndAlso vDate > CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value)) Then vValid = False

        If vValid Then
          vRenewalPrice = ProductRate.RenewalPrice(pFuture, ContactNumber, Quantity, pVATRate, pTransactionDate)
        Else
          ProductRate.PaymentPlanDetailPricing.CalculatePricing(UnitPrice, 0, VATExclusive, pTransactionDate.ToString(CAREDateFormat), pVATRate, False)
        End If
        Return vRenewalPrice
      End Get
    End Property
    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True this will return the actual price of this detail line,
    '''  otherwise it is either the current or future price depending on the date of the price change.
    '''   The price returned takes into account the <see cref="Quantity">Quantity</see>.
    '''   The calculated price will include any required VAT amount.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date) As Double
      Get
        Return Price(pDate, mvEnv.VATRate(Me.Product.ProductVatCategory, Me.PaymentPlan.Payer.VATCategory), 0, False)
      End Get
    End Property
    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <param name="pOverrideQuantity">The quantity of items being priced</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True this will return the actual price of this detail line,
    '''  otherwise it is either the current or future price depending on the date of the price change.
    '''   The price returned takes into account the <see cref="Quantity">Quantity</see> (or the <paramref name="pOverrideQuantity">override quantity</paramref> if not zero).
    '''   The calculated price will include any required VAT amount.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date, ByVal pOverrideQuantity As Double) As Double
      Get
        Return Price(pDate, mvEnv.VATRate(Me.Product.ProductVatCategory, Me.PaymentPlan.Payer.VATCategory), pOverrideQuantity, False, False)
      End Get
    End Property
    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <param name="pVatRate">The VAT rate to use in the pricing</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True this will return the actual price of this detail line,
    '''  otherwise it is either the current or future price depending on the date of the price change.
    '''   The price returned takes into account the <see cref="Quantity">Quantity</see>.
    '''   The calculated price will include any required VAT amount using the <paramref name="pVATRate">VAT rate supplied</paramref>.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date, ByVal pVATRate As VatRate) As Double
      Get
        Return Price(pDate, pVATRate, 0, False, False)
      End Get
    End Property
    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <param name="pOverrideQuantity">The quantity of items being priced</param>
    ''' <param name="pVatRate">The VAT rate to use in the pricing</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True this will return the actual price of this detail line,
    '''  otherwise it is either the current or future price depending on the date of the price change.
    '''   The price returned takes into account the <see cref="Quantity">Quantity</see>
    '''  (or the <paramref name="pOverrideQuantity">override quantity</paramref> if not zero).
    '''   The calculated price will include any required VAT amount using the <paramref name="pVATRate">VAT rate supplied</paramref>.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date, ByVal pVATRate As VatRate, ByVal pOverrideQuantity As Double) As Double
      Get
        Return Price(pDate, pVATRate, pOverrideQuantity, False, False)
      End Get
    End Property

    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <param name="pOverrideQuantity">The quantity of items being priced</param>
    ''' <param name="pVatRate">The VAT rate to use in the pricing</param>
    ''' <param name="pRenewals">Boolean flag indicating whether this is a Renewals process in which the price must be re-calculated</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True and <paramref name="pRenewals">it is not a renewal</paramref>
    ''' this will return the actual price of this detail line,  otherwise it is either the current or future
    ''' price depending on the date of the price change.   The price returned takes into account the 
    ''' <see cref="Quantity">Quantity</see> (or the <paramref name="pOverrideQuantity">override quantity</paramref> if not zero).
    '''  The calculated price will include any required VAT amount using the <paramref name="pVATRate">VAT rate supplied</paramref>.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date, ByVal pVATRate As VatRate, ByVal pOverrideQuantity As Double, ByVal pRenewals As Boolean) As Double
      Get
        Return Price(pDate, pVATRate, pOverrideQuantity, pRenewals, False)
      End Get
    End Property

    ''' <summary>The price for this payment plan detail line</summary>
    ''' <param name="pDate">Date the price is required for</param>
    ''' <param name="pOverrideQuantity">The quantity of items being priced</param>
    ''' <param name="pVatRate">The VAT rate to use in the pricing</param>
    ''' <param name="pRenewals">Boolean flag indicating whether this is a Renewals process in which the price must be re-calculated</param>
    ''' <param name="pCalcRenewalAmountOnly">Boolean flag indicating whether the price is for a Renewal Amount calculation in which the full annual amount is required</param>
    ''' <returns>Calculated Price including VAT</returns>
    ''' <remarks>If <see cref="HasPriceInfo">HasPriceInfo</see> is True and <paramref name="pRenewals">it is not a renewal</paramref>
    ''' this will return the actual price of this detail line,  otherwise it is either the current or future
    ''' price depending on the date of the price change.   The price returned takes into account the 
    ''' <see cref="Quantity">Quantity</see> (or the <paramref name="pOverrideQuantity">override quantity</paramref> if not zero).
    '''  The calculated price will include any required VAT amount using the <paramref name="pVATRate">VAT rate supplied</paramref>.</remarks>
    Public ReadOnly Property Price(ByVal pDate As Date, ByVal pVATRate As VatRate, ByVal pOverrideQuantity As Double, ByVal pRenewals As Boolean, ByVal pCalcRenewalAmountOnly As Boolean) As Double
      Get
        Dim vPrice As Double
        Dim vValid As Boolean = True

        If mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value.Length > 0 AndAlso pDate < CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidFrom).Value) Then vValid = False
        If vValid AndAlso (mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value.Length > 0 AndAlso pDate > CDate(mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value)) Then vValid = False

        If vValid Then
          Dim vQuantity As Double = Quantity
          If pOverrideQuantity <> 0 Then vQuantity = pOverrideQuantity
          If pRenewals = False AndAlso HasPriceInfo = True Then
            'Detail line contains the prices
            If Quantity <> vQuantity Then
              'Quantities have changed so calculate the Price (ensuring it includes VAT)
              vPrice = FixTwoPlaces(UnitPrice * vQuantity)
              If ProductRate.VatExclusive Then vPrice = FixTwoPlaces(vPrice + pVATRate.CalculateVATAmount(vPrice, ProductRate.VatExclusive, pDate.ToString(CAREDateFormat)))
            ElseIf pCalcRenewalAmountOnly = True Then
              vPrice = FixTwoPlaces(UnitPrice * Quantity)
              If ProductRate.VatExclusive Then vPrice = FixTwoPlaces(vPrice + pVATRate.CalculateVATAmount(vPrice, ProductRate.VatExclusive, pDate.ToString(CAREDateFormat)))
            Else
              'Return the Price with VAT
              vPrice = GrossAmount
            End If
          Else
            'Always calculate the Price
            If mvClassFields.Item(PaymentPlanDetailFields.odfAmount).Value.Length = 0 Then
              vPrice = ProductRate.Price(pDate, ContactNumber, vQuantity, pVATRate)
            Else
              'Restrict to the Amount specified on the line
              vPrice = mvClassFields.Item(PaymentPlanDetailFields.odfAmount).DoubleValue
            End If
          End If
        Else
          ProductRate.PaymentPlanDetailPricing.CalculatePricing(UnitPrice, 0, VATExclusive, pDate.ToString(CAREDateFormat), pVATRate, False)
        End If
        Return vPrice
      End Get
    End Property

    ''' <summary>
    ''' The containing payment plan for this detail line.
    ''' </summary>
    ''' <remarks>If the payment plan is pecified when the detail line is instantiated, this will be that payment plan. If not,
    ''' the payment plan will be constructed from the database using the order number. Once a payment plan has been constructed,
    ''' it will not be constructed again for the lifetime of the detail line.</remarks>
    Public ReadOnly Property PaymentPlan As PaymentPlan
      Get
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan()
          mvPaymentPlan.Init(mvEnv, Me.OrderNumber)
        End If
        Debug.Assert(mvPaymentPlan.OrderNumber = Me.OrderNumber)
        Return mvPaymentPlan
      End Get
    End Property

    ''' <summary>
    ''' The order number associated with the containing payment plan.
    ''' </summary>
    Public ReadOnly Property OrderNumber As Integer
      Get
        Return mvClassFields(PaymentPlanDetailFields.odfOrderNumber).IntegerValue
      End Get
    End Property
    ''' <summary>
    ''' Gets the prorated price.
    ''' </summary>
    ''' <param name="pStartDate">The start date of the proration period</param>
    ''' <param name="pEndDate">The end date of the proration period</param>
    ''' <returns>The prorated price for the protion of the term indicated</returns>
    ''' <remarks>The calculation is performed by taking the currently indicated price for the line item at the term start and
    ''' multiplying it by the proportion of months in the full term represented by the number of months between the 
    ''' <paramref name="pStartDate">start</paramref> and <paramref name="pEndDate">end</paramref> dates.  In practice,
    ''' multiplication is performed before division to ensure maximum accuracy.  The calcultion uses the <see cref="Price">Price</see>
    ''' property and <see cref="MonthsDifference">MonthsDifference</see> function in calculating the result.</remarks>
    Public ReadOnly Property ProratedPrice(ByVal pStartDate As Date, ByVal pEndDate As Date) As Double
      Get
        Debug.Assert(pStartDate <= pEndDate)
        Debug.Assert(pStartDate >= Me.PaymentPlan.TermStartDate)
        Debug.Assert(pEndDate <= Me.PaymentPlan.TermEndDate)
        If mvFullPrice > 0 Then
          mvProratedPrice = FixTwoPlaces((Price(Me.PaymentPlan.TermStartDate) * MonthsDifference(pStartDate, pEndDate)) / MonthsDifference(Me.PaymentPlan.TermStartDate, Me.PaymentPlan.TermEndDate))
        Else
          mvProratedPrice = 0
        End If
        Return mvProratedPrice
      End Get
    End Property

    ''' <summary>Gets the prorated price.</summary>
    ''' <param name="pFullTermMonths">The current number of months covered by the <see cref="FullPrice">FullPrice</see>.</param>
    ''' <param name="pNumberOfMonths">The number of months the <see cref="FullPrice">FullPrice</see> is to be prorated.</param>
    ''' <returns>Prorated Price</returns>
    Friend ReadOnly Property ProratedPrice(ByVal pFullTermMonths As Integer, ByVal pNumberOfMonths As Integer) As Double
      Get
        mvProratedPrice = CalculateProrateAmount(mvFullPrice, pFullTermMonths, pNumberOfMonths)
        Return mvProratedPrice
      End Get
    End Property

    Public ReadOnly Property Product() As Product
      Get
        If mvProduct Is Nothing Then
          mvProduct = New Product(mvEnv)
          mvProduct.Init(ProductCode)
        End If
        Product = mvProduct
      End Get
    End Property

    Public Sub InitNew(pProductCode As String, pRateCode As String, pDistributionCode As String)
      mvClassFields.Item(PaymentPlanDetailFields.odfProduct).Value = pProductCode
      mvClassFields.Item(PaymentPlanDetailFields.odfRate).Value = pRateCode
      mvClassFields.Item(PaymentPlanDetailFields.odfDistributionCode).Value = pDistributionCode
    End Sub

    Public Property ProductRate() As ProductRate
      Get
        If mvProductRate Is Nothing OrElse mvProductRate.IsValid = False Then
          mvProductRate = New ProductRate(mvEnv)
          mvProductRate.Init(ProductCode, RateCode)
        End If
        Return mvProductRate
      End Get
      Set(pValue As ProductRate)
        mvProductRate = pValue
        mvClassFields.Item(PaymentPlanDetailFields.odfProduct).Value = mvProductRate.ProductCode
        mvClassFields.Item(PaymentPlanDetailFields.odfRate).Value = mvProductRate.RateCode
      End Set
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfProduct).Value
      End Get
    End Property

    Public Property ProductNumber() As String
      Get
        ProductNumber = mvClassFields.Item(PaymentPlanDetailFields.odfProductNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfProductNumber).Value = Value
      End Set
    End Property

    Public Property Quantity() As Double
      Get
        Quantity = mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).DoubleValue
      End Get
      Set(ByVal Value As Double)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBTAQuantityDecimal) Then
          mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).DoubleValue = Value
        Else
          mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).IntegerValue = CInt(Value)
        End If
      End Set
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfRate).Value
      End Get
    End Property

    Public Property Subscription() As Boolean
      Get
        Subscription = mvSubscription
      End Get
      Set(ByVal Value As Boolean)
        mvSubscription = Value
      End Set
    End Property

    Public Property TimeStatus() As String
      Get
        TimeStatus = mvClassFields.Item(PaymentPlanDetailFields.odfTimeStatus).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfTimeStatus).Value = Value
      End Set
    End Property

    Public Property SubscriptionValidFrom() As String
      Get
        Return mvValidFrom
      End Get
      Set(ByVal Value As String)
        mvValidFrom = Value
      End Set
    End Property

    Public Property SubscriptionValidTo() As String
      Get
        Return mvValidTo
      End Get
      Set(ByVal Value As String)
        mvValidTo = Value
      End Set
    End Property

    ''' <summary>Calculate the VATAmount for the current Balance for the MailMege file.</summary>
    ''' <param name="pVATPercentage">VAT percentage to use</param>
    ''' <returns>VAT Amount appropriate for the current Balance</returns>
    ''' <remarks>This is only used by PaymentPlan.WriteMailMergeOutput</remarks>
    Friend Function CalculateVATAmount(ByVal pVATPercentage As Double) As Double
      Dim vVATAmount As Double = 0
      If CDbl(mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).Value) > 0 Then
        vVATAmount = Int(((CDbl(mvClassFields.Item(PaymentPlanDetailFields.odfBalance).Value) - (CDbl(mvClassFields.Item(PaymentPlanDetailFields.odfBalance).Value) / (1 + pVATPercentage / 100))) * 100) + 0.5) / 100
      Else
        vVATAmount = 0
      End If
    End Function

    Public Property Source() As String
      Get
        Source = mvClassFields.Item(PaymentPlanDetailFields.odfSource).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfSource).Value = Value
      End Set
    End Property

    Public ReadOnly Property CreatedBy() As String
      Get
        CreatedBy = mvClassFields.Item(PaymentPlanDetailFields.odfCreatedBy).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(PaymentPlanDetailFields.odfCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property SubscriptionDataAmended() As SubscriptionDataTypes
      Get
        Dim vClassField As ClassField
        Dim vSubType As SubscriptionDataTypes
        Dim vValue As SubscriptionDataTypes

        vSubType = SubscriptionDataTypes.sdtNone
        For Each vClassField In mvClassFields
          With vClassField
            Select Case .Name
              Case "contact_number"
                vValue = SubscriptionDataTypes.sdtContactNumber
              Case "address_number"
                vValue = SubscriptionDataTypes.sdtAddressNumber
              Case "quantity"
                vValue = SubscriptionDataTypes.sdtQuantity
              Case "despatch_method"
                vValue = SubscriptionDataTypes.sdtDespatchMethod
              Case "communication_number"
                vValue = SubscriptionDataTypes.sdtCommunicationNumber
              Case Else
                vValue = SubscriptionDataTypes.sdtNone
            End Select
            If .InDatabase And .ValueChanged Then vSubType = vSubType Or vValue
          End With
        Next vClassField
        SubscriptionDataAmended = vSubType
      End Get
    End Property

    Public ReadOnly Property SubscriptionNumber(Optional ByVal pExcludeExpired As Boolean = True) As Integer
      Get
        If Subscription And mvSubscriptionNumber = 0 Then GetSubscriptionData(pExcludeExpired)
        SubscriptionNumber = mvSubscriptionNumber
      End Get
    End Property

    Public Property CommunicationNumber() As String
      Get
        CommunicationNumber = mvClassFields.Item(PaymentPlanDetailFields.odfCommunicationNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfCommunicationNumber).Value = Value
      End Set
    End Property

    ''' <summary>Fixed Amount excluding VAT</summary>
    Public Property NetFixedAmount() As String
      Get
        'Could be null
        Return mvClassFields.Item(PaymentPlanDetailFields.odfNetFixedAmount).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(PaymentPlanDetailFields.odfNetFixedAmount).Value = Value
      End Set
    End Property

    Public ReadOnly Property SpecialInitialPeriodIncentive() As Boolean
      Get
        SpecialInitialPeriodIncentive = mvSpecialInitialPeriod
      End Get
    End Property

    Public ReadOnly Property MemberOrPayer() As String
      Get
        MemberOrPayer = mvMemberOrPayer
      End Get
    End Property

    Public ReadOnly Property UsesProductNumbers() As Boolean
      Get
        UsesProductNumbers = mvUsesProductNumbers
      End Get
    End Property

    Public ReadOnly Property IncentiveLineType() As String
      Get
        IncentiveLineType = mvIncentiveLineType
      End Get
    End Property

    Public ReadOnly Property IncentiveIgnoreProductAndRate() As Boolean
      Get
        IncentiveIgnoreProductAndRate = mvIgnoreProductAndRate
      End Get
    End Property

    Public ReadOnly Property IncentiveProductDesc() As String
      Get
        IncentiveProductDesc = mvIncentiveProductDesc
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvCancellationReason
      End Get
    End Property

    Public Property LineNumber() As Integer
      Get
        'only used for the smart client
        LineNumber = mvLineNumber
      End Get
      Set(ByVal Value As Integer)
        'only used for the smart client
        mvLineNumber = Value
      End Set
    End Property

    Public Property AmountPaid As Double
      Get
        Return mvAmountPaid
      End Get
      Set(ByVal value As Double)
        mvAmountPaid = value
      End Set
    End Property

    Friend Function FullPrice(ByVal pTermStartDate As Date, ByVal pContactCategory As String) As Double
      Dim vVatRate As VatRate = mvEnv.VATRate(Product.ProductVatCategory, pContactCategory)
      mvFullPrice = Price(pTermStartDate, vVatRate, Quantity, False, True)
      Return mvFullPrice
    End Function

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        Select Case pAttributeName
          Case "AccruesInterest"
            mvAccruesInterest = BooleanValue(Value)
          Case "DetailNumber"
            mvClassFields.ItemValue(pAttributeName) = Value
            If Val(Value) > 0 Then mvExisting = True
          Case "LineNumber"
            mvLineNumber = IntegerValue(Value)
          Case "LoanInterest"
            mvLoanInterest = BooleanValue(Value)
          Case "MemberOrPayer"
            mvMemberOrPayer = Value
          Case "Product"
            mvClassFields.ItemValue(pAttributeName) = Value
            If Product.Donation Then
              mvDonation = True
            ElseIf Product.Subscription Then
              mvSubscription = True
            End If
          Case "PaymentPlanNumber"
            mvClassFields(PaymentPlanDetailFields.odfOrderNumber).Value = Value
          Case "SubscriptionNumber"
            mvSubscriptionNumber = IntegerValue(Value)
          Case "SubsValidFrom"
            mvValidFrom = Value
          Case "SubsValidTo"
            mvValidTo = Value
          Case "PPDLineType"
            mvDetailType = CType(Value, PaymentPlanDetailTypes)
          Case "UsesProductNumbers"
            mvUsesProductNumbers = (Value = "Y")
          Case "IncentiveLineType"
            mvIncentiveLineType = Value
            mvSpecialInitialPeriod = (Mid(mvIncentiveLineType, 2, 1) = "I") 'Initial Period Incentive
          Case "IncentiveProductDesc"
            mvIncentiveProductDesc = Value
          Case "IncentiveIgnoreProductAndRate"
            mvIgnoreProductAndRate = (Value = "Y")
          Case "PriceIsPercentage", "DiscountPercentage"
            'ignore readonly field
          Case "FullPrice"
            mvFullPrice = DoubleValue(Value)
          Case "ProratedPrice"
            mvProratedPrice = DoubleValue(Value)
          Case "CMTProrateCostCode"
            mvCMTProrateCostCode = Value
          Case "CMTExcessPaymentTypeCode"
            mvCMTExcessPaymentTypeCode = Value
          Case "EntitlementSequenceNumber"
            mvEntitlementSequenceNumber = IntegerValue(Value)
          Case "CMTRefundProductCode"
            mvCMTRefundProductCode = Value
          Case "CMTRefundRateCode"
            mvCMTRefundRateCode = Value
          Case "ProductDesc", "RateDesc", "CMTProrateCost", "CMTExcessPaymentType", "ExcessAmount"
            'Do nothing
          Case Else
            mvClassFields.ItemValue(pAttributeName) = Value
        End Select
      End Set
    End Property

    Public Sub SetImportBalance(ByRef pBalance As Double)
      mvClassFields.Item(PaymentPlanDetailFields.odfBalance).Value = CStr(pBalance)
    End Sub

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(PaymentPlanDetailFields.odfAmendedOn).Value = pAmendedOn
      mvClassFields.Item(PaymentPlanDetailFields.odfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Sub SetSubscriptionNumber(ByVal pNewValue As Integer)
      mvSubscriptionNumber = pNewValue
    End Sub

    Public Sub GetSubscriptionData(Optional ByVal pExcludeExpired As Boolean = True, Optional ByVal pExcludeCancelled As Boolean = True)
      Dim vWhereFields As New CDBFields
      Dim vRS As CDBRecordSet

      If Subscription Then
        With vWhereFields
          .Add("order_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(PaymentPlanDetailFields.odfOrderNumber).IntegerValue)
          'The SetValue property is used for the following items because these values may have been changed but not yet saved by Payment Plan Maintenance.
          'For instance, if PPD.ContactNumber is changed from 1 to 2 there won't be a subscriptions record for the PP and contact 2...yet.
          'But the one for contact 1 still exists and that's the subscription number we want.
          .Add("contact_number", CDBField.FieldTypes.cftLong, Val(mvClassFields.Item(PaymentPlanDetailFields.odfContactNumber).SetValue))
          .Add("address_number", CDBField.FieldTypes.cftLong, Val(mvClassFields.Item(PaymentPlanDetailFields.odfAddressNumber).SetValue))
          .Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(PaymentPlanDetailFields.odfProduct).SetValue)
          .Add("quantity", CDBField.FieldTypes.cftInteger, Val(mvClassFields.Item(PaymentPlanDetailFields.odfQuantity).SetValue))
          .Add("despatch_method", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(PaymentPlanDetailFields.odfDespatchMethod).SetValue)
          If pExcludeCancelled Then .Add("cancellation_reason", CDBField.FieldTypes.cftCharacter)
          If pExcludeExpired Then .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End With
        vRS = mvEnv.Connection.GetRecordSet("SELECT subscription_number, valid_from, valid_to, cancellation_reason FROM subscriptions WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        With vRS
          If .Fetch() = True Then
            mvSubscriptionNumber = .Fields(1).IntegerValue 'subscription_number
            SubscriptionValidFrom = .Fields(2).Value 'valid_from
            SubscriptionValidTo = .Fields(3).Value 'valid_to
            mvCancellationReason = .Fields.Item(4).Value 'cancellation_reason
          End If
          .CloseRecordSet()
        End With
      End If
    End Sub

    Public Function GetDataAsParameters() As CDBParameters
      Dim vParams As New CDBParameters
      Dim vField As ClassField

      For Each vField In mvClassFields
        vParams.Add(ProperName((vField.Name)), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(vField.Value), vField.Value))
      Next vField
      vParams.Add("MemberOrPayer", CDBField.FieldTypes.cftCharacter, mvMemberOrPayer)
      vParams.Add("LineNumber", mvLineNumber)
      vParams.Add("SubscriptionNumber", SubscriptionNumber(False)) 'Same as Rich Client Trader
      vParams.Add("PPDLineType", mvDetailType)
      vParams.Add("UsesProductNumbers", CDBField.FieldTypes.cftCharacter, BooleanString(mvUsesProductNumbers))
      vParams.Add("SubsValidFrom", CDBField.FieldTypes.cftDate, mvValidFrom)
      vParams.Add("SubsValidTo", CDBField.FieldTypes.cftDate, mvValidTo)
      vParams.Add("IncentiveLineType", CDBField.FieldTypes.cftCharacter, mvIncentiveLineType)
      vParams.Add("IncentiveProductDesc", CDBField.FieldTypes.cftCharacter, mvIncentiveProductDesc)
      vParams.Add("IncentiveIgnoreProductAndRate", CDBField.FieldTypes.cftCharacter, BooleanString(mvIgnoreProductAndRate))
      vParams.Add("PriceIsPercentage", CDBField.FieldTypes.cftCharacter, mvProductRate.PriceIsPercentage)
      'Return the discount percent so the discount amount can be calculated on the client side when items are added/removed
      vParams.Add("DiscountPercentage", CDBField.FieldTypes.cftCharacter) 'Return empty discount rate. The rate wil be fetched from the rates table.
      'If mvProductRate.PriceIsPercentage.Length > 0 AndAlso mvProductRate.PriceIsPercentage <> "N" Then vParams("DiscountPercentage").Value = mvClassFields(PaymentPlanDetailFields.odfBalance).Value
      vParams.Add("AccruesInterest", CDBField.FieldTypes.cftCharacter, BooleanString(mvAccruesInterest))
      vParams.Add("LoanInterest", CDBField.FieldTypes.cftCharacter, BooleanString(mvLoanInterest))
      If mvProduct Is Nothing OrElse mvProduct.Existing = False Then
        mvProduct = New Product(mvEnv)
        mvProduct.InitWithRate(mvEnv, ProductCode, RateCode)
        mvProductRate = mvProduct.ProductRate
      End If
      If mvProductRate Is Nothing OrElse mvProductRate.Existing = False Then
        mvProductRate = New ProductRate(mvEnv)
        mvProductRate.Init(ProductCode, RateCode)
      End If
      vParams.Add("ProductDesc", mvProduct.ProductDesc)
      vParams.Add("RateDesc", mvProductRate.RateDesc)
      vParams.Add("FullPrice", mvFullPrice)
      vParams.Add("ProratedPrice", mvProratedPrice)
      vParams.Add("ExcessAmount", FixedFormat(FixTwoPlaces(mvFullPrice - mvProratedPrice)))
      vParams.Add("CMTProrateCostCode", mvCMTProrateCostCode)
      vParams.Add("CMTExcessPaymentTypeCode", mvCMTExcessPaymentTypeCode)
      vParams.Add("EntitlementSequenceNumber", mvEntitlementSequenceNumber)
      vParams.Add("CMTRefundProductCode", mvCMTRefundProductCode)
      vParams.Add("CMTRefundRateCode", mvCMTRefundRateCode)
      GetDataAsParameters = vParams
    End Function

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      Select Case pAttributeName
        Case "LineNumber", "SubscriptionNumber", "PPDLineType", "PaymentPlanNumber", "DiscountPercentage", "EntitlementSequenceNumber"
          LineDataType = CDBField.FieldTypes.cftLong
        Case "MemberOrPayer", "UsesProductNumbers", "IncentiveLineType", "IncentiveProductDesc", "IncentiveIgnoreProductAndRate", "PriceIsPercentage", "AccruesInterest", "LoanInterest", _
             "ProductDesc", "RateDesc", "CMTProrateCost", "CMTProrateCostCode", "CMTExcessPaymentType", "CMTExcessPaymentTypeCode", _
             "CMTRefundProductCode", "CMTRefundRateCode"
          LineDataType = CDBField.FieldTypes.cftCharacter
        Case "SubsValidFrom", "SubsValidTo"
          LineDataType = CDBField.FieldTypes.cftDate
        Case "FullPrice", "ProratedPrice", "ExcessAmount"
          LineDataType = CDBField.FieldTypes.cftNumeric
        Case Else
          LineDataType = mvClassFields.ItemDataType(pAttributeName)
      End Select
    End Function

    Friend Sub SetPrices()
      If Not mvProductRate.IsValid Then
        'No price information has been set
        Dim vProduct As New Product(mvEnv)
        vProduct.InitWithRate(mvEnv, ProductCode, RateCode)
        If vProduct.Existing Then
          mvSubscription = vProduct.Subscription
          mvUsesProductNumbers = vProduct.UsesProductNumbers
          mvProductRate = vProduct.ProductRate
          mvProduct = vProduct
        End If
      End If
    End Sub

    Public Sub SetEffectiveDate(ByVal pEffectiveDate As String)
      'Used by Rich Client Trader and Data Import only
      'Added this here as (i)  I didn't want to make the EffectiveDate property writable, and
      '                   (ii) The Create method would not take any extra parameters without putting them on multiple lines
      If Not (IsDate(pEffectiveDate)) Then pEffectiveDate = ""
      mvClassFields.Item(PaymentPlanDetailFields.odfEffectiveDate).Value = pEffectiveDate
    End Sub

    Friend ReadOnly Property VATExclusive() As Boolean
      Get
        Return mvProductRate.VatExclusive
      End Get
    End Property

    ''' <summary>Gets a boolean flag indicating whether the Product accrues interest.</summary>
    ''' <returns>True if the Product accrues interest, otherwise False.</returns>
    Friend ReadOnly Property AccruesInterest() As Boolean
      Get
        Return mvAccruesInterest
      End Get
    End Property

    ''' <summary>Gets a boolean flag identifying the ProductRate is for Loan Interest.</summary>
    ''' <returns>True if the ProductRate is for Loan Interest, otherwise False.</returns>
    Friend ReadOnly Property LoanInterest() As Boolean
      Get
        Return mvLoanInterest
      End Get
    End Property

    Friend ReadOnly Property ModifierActivity As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierActivity).Value
      End Get
    End Property

    Friend ReadOnly Property ModifierActivityValue As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierActivityValue).Value
      End Get
    End Property

    Friend ReadOnly Property ModifierActivityQuantity As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierActivityQuantity).DoubleValue
      End Get
    End Property

    Friend ReadOnly Property ModifierActivityDate As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierActivityDate).Value
      End Get
    End Property

    ''' <summary>The price of the detail calculated from the Rate Modifiers.  If the <see cref="ModifierPerItem" /> flag is set to 'Y' then this is the price per item.</summary>
    ''' <remarks>This is not set if multiple Rate Modifiers were used.</remarks>
    Friend ReadOnly Property ModifierPrice As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierPrice).DoubleValue
      End Get
    End Property

    ''' <summary>Flay indicating whether the <see cref="ModifierPrice" /> is the total price or the price for a single <see cref="ModifierActivityQuantity" />.</summary>
    ''' <remarks>This will be set to M if multiple Rate Modifiers were used.</remarks>
    Friend ReadOnly Property ModifierPerItem As String
      Get
        'Cound be Y/N/M/null
        Return mvClassFields.Item(PaymentPlanDetailFields.odfModifierPerItem).Value
      End Get
    End Property

    ''' <summary>The full price of this detail line for a single item before any discounts or pro-rating etc. have been applied.</summary>
    Friend ReadOnly Property UnitPrice As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfUnitPrice).DoubleValue
      End Get
    End Property

    Friend ReadOnly Property ProRated As Boolean
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfProRated).Bool
      End Get
    End Property

    ''' <summary>The net price of this detail line taking into account the quantity.  This is after discounts, pro-rating etc. have been applied.</summary>
    Friend ReadOnly Property NetAmount As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue
      End Get
    End Property

    ''' <summary>The VAT amount of this detail line taking into account the quantity and using the payer VAT category.  This is after discounts, pro-rating etc. have been applied.</summary>
    Friend ReadOnly Property VatAmount As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue
      End Get
    End Property

    ''' <summary>The gross price of this detail line taking into account the quantity.  This is after discounts, pro-rating etc. have been applied.</summary>
    Friend ReadOnly Property GrossAmount As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue
      End Get
    End Property

    Friend ReadOnly Property VatRateCode As String
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfVatRate).Value
      End Get
    End Property

    Friend ReadOnly Property VatPercentage As Double
      Get
        Return mvClassFields.Item(PaymentPlanDetailFields.odfVatPercentage).DoubleValue
      End Get
    End Property

    ''' <summary>Does the PaymentPlanDetail contain the Price information</summary>
    Friend ReadOnly Property HasPriceInfo As Boolean
      Get
        'We assume that if these are set then all relevant fields have been set
        Return (mvClassFields.Item(PaymentPlanDetailFields.odfUnitPrice).Value.Length > 0 AndAlso mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).Value.Length > 0)
      End Get
    End Property

    Public Sub SetModifierAndPriceData(ByVal pPPDPricing As PaymentPlanDetailPricing)
      SetModifierAndPriceData(pPPDPricing, False)
    End Sub
    Public Sub SetModifierAndPriceData(ByVal pPPDPricing As PaymentPlanDetailPricing, ByVal pIgnoreQuantity As Boolean)
      'Quantity property may not be the number of units. For some incentives it might hold the number of months for the first period of a membership. In this case the pIgnoreQuantity is set to true.  
      With mvClassFields
        If pPPDPricing.ModifierActivity.Length > 0 Then
          .Item(PaymentPlanDetailFields.odfModifierActivity).Value = pPPDPricing.ModifierActivity
          .Item(PaymentPlanDetailFields.odfModifierActivityValue).Value = pPPDPricing.ModifierActivityValue
          If pPPDPricing.ModifierPerItem <> "M" Then
            'Do not set quantity or price if multiple RateModifiers were used
            .Item(PaymentPlanDetailFields.odfModifierActivityQuantity).DoubleValue = pPPDPricing.ModifierActivityQuantity
            .Item(PaymentPlanDetailFields.odfModifierPrice).DoubleValue = pPPDPricing.ModifierPrice
          End If
          .Item(PaymentPlanDetailFields.odfModifierActivityDate).Value = pPPDPricing.ModifierActivityDate
          .Item(PaymentPlanDetailFields.odfModifierPerItem).Value = pPPDPricing.ModifierPerItem
        End If
        .Item(PaymentPlanDetailFields.odfUnitPrice).DoubleValue = pPPDPricing.UnitPrice
        .Item(PaymentPlanDetailFields.odfProRated).Bool = pPPDPricing.ProRated
        If Me.Quantity > 1 And Not pIgnoreQuantity Then
          'BR20025 Recalculate Nett Amount and  then Recalulate VatAmount and GrossAmount from Nett Amount where the quantity is greater than one.
          Dim vVATRate As New VatRate(mvEnv)
          mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(pPPDPricing.NetAmount * Me.Quantity)
          mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = FixTwoPlaces(vVATRate.CalculateVATAmount(mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue, Me.VATExclusive, VatPercentage))
          mvClassFields.Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = FixTwoPlaces(mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue + mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue)
        Else
          mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = pPPDPricing.NetAmount
          mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = pPPDPricing.VatAmount
          mvClassFields.Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = pPPDPricing.GrossAmount
        End If
        .Item(PaymentPlanDetailFields.odfVatRate).Value = pPPDPricing.VatRate
        .Item(PaymentPlanDetailFields.odfVatPercentage).DoubleValue = pPPDPricing.VatPercentage
      End With
      If mvIgnoreProductAndRate = False AndAlso (Balance = 0 AndAlso GrossAmount <> 0) Then
        'Could be adding a detail line with no charge at this time even though it is chargeable (e.g. Membership Entitlement not being charged in first year)
        mvClassFields.Item(PaymentPlanDetailFields.odfProRated).Bool = True
        mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = 0
        mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = 0
        mvClassFields.Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = 0
      End If
    End Sub

    Private Sub SetIncentivesModifierAndPriceData(ByVal pIgnoreProductAndRate As Boolean)
      If pIgnoreProductAndRate = False Then
        'Incentives do not use RateModifiers and PPD line has been updated to use the incentive product/rate
        'so clear any modifier data that may have been set
        With mvClassFields
          .Item(PaymentPlanDetailFields.odfModifierActivity).Value = ""
          .Item(PaymentPlanDetailFields.odfModifierActivityValue).Value = ""
          .Item(PaymentPlanDetailFields.odfModifierActivityQuantity).Value = ""
          .Item(PaymentPlanDetailFields.odfModifierActivityDate).Value = ""
          .Item(PaymentPlanDetailFields.odfModifierPerItem).Value = ""
          .Item(PaymentPlanDetailFields.odfModifierPrice).Value = ""
        End With
      End If
      'If the Balance has changed and we have VAT info then re-calculate Net/VAT/Gross
      If (Balance <> GrossAmount) AndAlso VatRateCode.Length > 0 AndAlso mvClassFields.Item(PaymentPlanDetailFields.odfVatPercentage).Value.Length > 0 Then
        mvClassFields.Item(PaymentPlanDetailFields.odfUnitPrice).DoubleValue = Balance
        mvClassFields.Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = Balance
        Dim vVATAmount As Double = 0
        If Balance <> 0 Then
          Dim vVATRate As VatRate = mvEnv.VATRate(VatRateCode)
          vVATAmount = vVATRate.CalculateVATAmount(Balance, False, VatPercentage)
        End If
        mvClassFields.Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(Balance - vVATAmount)
        mvClassFields.Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = vVATAmount
      End If
    End Sub

    Friend ReadOnly Property EntitlementSequenceNumber() As Integer
      Get
        Return mvEntitlementSequenceNumber
      End Get
    End Property

    Friend ReadOnly Property CMTProrateLineType() As MembershipType.CMTProrateCosts
      Get
        Return MembershipType.GetCMTProrateCosts(mvCMTProrateCostCode)
      End Get
    End Property

    Friend ReadOnly Property CMTExcessPaymentType() As CmtExcessPayment.CMTExcessPaymentTypes
      Get
        Return CmtExcessPayment.GetCMTExcessPaymentType(mvCMTExcessPaymentTypeCode)
      End Get
    End Property

    Friend Sub SetCMTData(ByVal pProrateCostCode As String, ByVal pExcessPaymentTypeCode As String, ByVal pSequenceNumber As Integer, ByVal pRefundProductCode As String, ByVal pRefundRateCode As String)
      mvCMTProrateCostCode = pProrateCostCode
      mvCMTExcessPaymentTypeCode = pExcessPaymentTypeCode
      mvEntitlementSequenceNumber = pSequenceNumber
      mvCMTRefundProductCode = pRefundProductCode
      mvCMTRefundRateCode = pRefundRateCode
    End Sub

    ''' <summary>Set the new Balance on a <see cref="PaymentPlanDetail">Detail</see> line when this Detail line is for the old membership type.</summary>
    ''' <param name="pAdvancedCMT">Is Advanced CMT being used?</param>
    ''' <param name="pOutstandingRemainingBalance">Sum outstanding Balances.</param>
    ''' <remarks>For an Advanced CMT, new Balances are calculated according to the <see cref="CMTProrateLineType">CMT Prorate Line Type</see>.</remarks>
    Friend Sub SetCMTOldTypeBalance(ByVal pAdvancedCMT As Boolean, ByRef pOutstandingRemainingBalance As Double)
      Dim vNewBalance As Double
      Dim vProrated As Boolean = True
      mvCMTExcessPaymentAmount = 0
      mvAmountPaid = 0
      If pAdvancedCMT Then
        mvAmountPaid = FixTwoPlaces(mvFullPrice - (Balance - Arrears))
        If mvAmountPaid < 0 Then mvAmountPaid = 0
        Select Case CMTProrateLineType
          Case MembershipType.CMTProrateCosts.FullCharge
            'Retain existing Balance
            vNewBalance = Balance
            vProrated = False
          Case MembershipType.CMTProrateCosts.NoCharge
            'Set Balance to zero and make payment amount an excess payment
            If mvAmountPaid <> 0 Then mvCMTExcessPaymentAmount = mvAmountPaid
            vNewBalance = 0
          Case MembershipType.CMTProrateCosts.Prorate
            'Prorate Balance
            vNewBalance = FixTwoPlaces(mvProratedPrice - mvAmountPaid)
            If vNewBalance < 0 Then
              mvCMTExcessPaymentAmount = Math.Abs(vNewBalance)
              vNewBalance = 0
            End If
        End Select
        If Arrears <> 0 Then
          If Arrears > vNewBalance Then vNewBalance = Arrears
        End If
        If vNewBalance > pOutstandingRemainingBalance Then vNewBalance = pOutstandingRemainingBalance
        pOutstandingRemainingBalance = FixTwoPlaces(pOutstandingRemainingBalance - vNewBalance)
      Else
        Dim vLineBalance As Double = Balance
        If vLineBalance >= pOutstandingRemainingBalance Then
          vNewBalance = pOutstandingRemainingBalance
          pOutstandingRemainingBalance = 0
        Else
          'Keep current Balance
          vNewBalance = vLineBalance
          pOutstandingRemainingBalance = FixTwoPlaces(pOutstandingRemainingBalance - vLineBalance)
        End If
        'Note - Pro-rating not on line basis so mvProratedPrice could be out by pennies due to rounding
      End If
      Dim vNewArrears As Double = Arrears
      If vNewArrears <> 0 Then
        If vNewArrears > vNewBalance Then vNewArrears = vNewBalance
      End If
      mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = vNewBalance
      mvClassFields.Item(PaymentPlanDetailFields.odfArrears).DoubleValue = vNewArrears
      If vProrated = True AndAlso HasPriceInfo = True Then
        'Update pricing data to show that this line has been pro-rated
        If mvProductRate Is Nothing OrElse mvProductRate.Existing = False Then
          mvProductRate = New ProductRate(mvEnv)
          mvProductRate.Init(ProductCode, RateCode)
        End If
        mvClassFields.Item(PaymentPlanDetailFields.odfProRated).Bool = True
        Dim vVATRate As VatRate = mvEnv.VATRate(VatRateCode)
        Dim vVATAmount As Double = vVATRate.CalculateVATAmount(mvProratedPrice, ProductRate.VatExclusive, VatPercentage)
        With mvClassFields
          .Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = vVATAmount
          If ProductRate.VatExclusive Then
            .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = FixTwoPlaces(mvProratedPrice + vVATAmount)
          Else
            .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = mvProratedPrice
          End If
          .Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(GrossAmount - VatAmount)
        End With
      End If
      mvAmendedValid = False
    End Sub

    ''' <summary>Set the new Balance on a <see cref="PaymentPlanDetail">Detail</see> line when this Detail line is for the new membership type.</summary>
    ''' <param name="pAdvancedCMT">Is Advanced CMT being used?</param>
    ''' <param name="pNewBalanceTotal">Sum balance of the new Detail lines</param>
    ''' <param name="pMonthsRemaining">The number of months remaining for the new membership type.</param>
    ''' <remarks>For an Advanced CMT, new Balances are calculated according to the <see cref="CMTProrateLineType">CMT Prorate Line Type</see>.</remarks>
    Friend Sub SetCMTNewTypeProrateBalance(ByVal pAdvancedCMT As Boolean, ByVal pNewBalanceTotal As Double, ByVal pMonthsRemaining As Integer)
      Dim vNewBalance As Double
      Dim vProrated As Boolean = True
      Dim vLinePrice As Double = mvProratedPrice
      Dim vBalanceChanged As Boolean
      If pAdvancedCMT Then
        Select Case CMTProrateLineType
          Case MembershipType.CMTProrateCosts.FullCharge
            vNewBalance = mvFullPrice
            vLinePrice = mvFullPrice
            vProrated = False
          Case MembershipType.CMTProrateCosts.NoCharge
            vNewBalance = 0
            vLinePrice = 0
          Case MembershipType.CMTProrateCosts.Prorate
            vNewBalance = mvProratedPrice
        End Select
        If vNewBalance > pNewBalanceTotal Then vNewBalance = pNewBalanceTotal
      Else
        vNewBalance = FixTwoPlaces((mvFullPrice / 12) * pMonthsRemaining)
        If vNewBalance > pNewBalanceTotal Then vNewBalance = pNewBalanceTotal
        vLinePrice = vNewBalance
      End If
      If Balance <> vNewBalance Then vBalanceChanged = True
      mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = vNewBalance
      If (vProrated = True OrElse vBalanceChanged = True) AndAlso HasPriceInfo = True Then
        'Update pricing data to show that this line has been pro-rated
        If mvProductRate Is Nothing OrElse mvProductRate.Existing = False Then
          mvProductRate = New ProductRate(mvEnv)
          mvProductRate.Init(ProductCode, RateCode)
        End If
        mvClassFields.Item(PaymentPlanDetailFields.odfProRated).Bool = True
        Dim vVATRate As VatRate = mvEnv.VATRate(VatRateCode)
        Dim vVATAmount As Double = vVATRate.CalculateVATAmount(vLinePrice, ProductRate.VatExclusive, VatPercentage)
        With mvClassFields
          .Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = vVATAmount
          If ProductRate.VatExclusive Then
            .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = FixTwoPlaces(vLinePrice + vVATAmount)
          Else
            .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = vLinePrice
          End If
          .Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(GrossAmount - VatAmount)
        End With
      End If
      mvAmendedValid = False
    End Sub

    Friend Sub CMTWriteOff(ByVal pProportion As Integer, ByVal pWOProportionally As Boolean, ByRef pWriteOffAmountTotal As Double)
      Dim vWriteOffAmount As Double = Balance
      If pWOProportionally Then
        vWriteOffAmount = FixTwoPlaces(mvFullPrice / pProportion)
      End If
      If vWriteOffAmount > pWriteOffAmountTotal Then vWriteOffAmount = pWriteOffAmountTotal
      If vWriteOffAmount >= Balance Then
        vWriteOffAmount = Balance
        Balance = 0
      Else
        Balance = FixTwoPlaces(Balance - vWriteOffAmount)
      End If
      pWriteOffAmountTotal = FixTwoPlaces(pWriteOffAmountTotal - vWriteOffAmount)
      mvAmendedValid = False
    End Sub

    ''' <summary>Set <see cref="ValidTo">Valid To</see> date to the renewal date and <see cref="UnitPrice">Unit Price</see> to zero so that this line does not get renewed.</summary>
    Friend Sub SetCMTLineNoRenewalRequired(ByVal pRenewalDate As String)
      pRenewalDate = CDate(pRenewalDate).AddDays(-1).ToString(CAREDateFormat)   'Deduct 1 day
      If IsDate(ValidTo) Then
        If CDate(ValidTo) > CDate(pRenewalDate) Then mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value = pRenewalDate
      Else
        mvClassFields.Item(PaymentPlanDetailFields.odfValidTo).Value = pRenewalDate
      End If
      If HasPriceInfo = True Then
        mvClassFields.Item(PaymentPlanDetailFields.odfUnitPrice).DoubleValue = 0
      Else
        mvClassFields.Item(PaymentPlanDetailFields.odfAmount).DoubleValue = 0
      End If
      mvAmendedValid = False
    End Sub

    Friend Sub CMTApplyExcessPayment(ByVal pOldPPD As PaymentPlanDetail)
      Dim vExcessPayment As Double = pOldPPD.ExcessPaymentAmount
      Dim vNewBalance As Double = Balance
      If vNewBalance > vExcessPayment Then
        vNewBalance = FixTwoPlaces(vNewBalance - vExcessPayment)
        vExcessPayment = 0
      Else
        vExcessPayment = FixTwoPlaces(vExcessPayment - vNewBalance)
        vNewBalance = 0
      End If
      mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = vNewBalance
      pOldPPD.AllocateCMTExcessPayment(vExcessPayment)
      mvAmendedValid = False
    End Sub

    Friend Sub SetNewBalanceForCMT(ByVal pNewBalance As Double, ByVal pAdvancedCMT As Boolean)
      'Pro-rating may have caused a rounding error which needs to be corrected.
      Dim vStartBalance As Double = Balance
      mvClassFields.Item(PaymentPlanDetailFields.odfBalance).DoubleValue = pNewBalance
      If pAdvancedCMT = False AndAlso HasPriceInfo Then
        If (vStartBalance = GrossAmount) AndAlso ProRated = True AndAlso ((UnitPrice * Quantity) > vStartBalance) Then
          With mvClassFields
            Dim vVATRate As VatRate = mvEnv.VATRate(VatRateCode)
            If mvProductRate Is Nothing OrElse mvProductRate.Existing = False Then
              mvProductRate = New ProductRate(mvEnv)
              mvProductRate.Init(ProductCode, RateCode)
            End If
            .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = pNewBalance
            .Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = vVATRate.CalculateVATAmount(pNewBalance, mvProductRate.VatExclusive, VatPercentage)
            .Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(pNewBalance - VatAmount)
          End With
        End If
      End If
      mvAmendedValid = False
    End Sub

    ''' <summary>Set the excess payment amount to the amount not allocated for late refunding etc.</summary>
    ''' <param name="pNewExcessPayment">Amount of excess payment not allocated</param>
    Friend Sub AllocateCMTExcessPayment(ByVal pNewExcessPayment As Double)
      mvCMTExcessPaymentAmount = pNewExcessPayment
    End Sub

    Friend ReadOnly Property GetFullPrice() As Double
      Get
        Return mvFullPrice
      End Get
    End Property

    Friend ReadOnly Property GetProratedPrice() As Double
      Get
        Return mvProratedPrice
      End Get
    End Property

    Friend ReadOnly Property ExcessPaymentAmount() As Double
      Get
        Return mvCMTExcessPaymentAmount
      End Get
    End Property

    Friend ReadOnly Property CMTAdjustmentProductCode() As String
      Get
        Select Case CMTExcessPaymentType
          Case CmtExcessPayment.CMTExcessPaymentTypes.Refund
            Return mvCMTRefundProductCode
          Case CmtExcessPayment.CMTExcessPaymentTypes.Retain
            Return mvClassFields.Item(PaymentPlanDetailFields.odfProduct).Value
          Case Else   'CmtExcessPayment.CMTExcessPaymentTypes.CarryForward, CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse
            Return ""
        End Select
      End Get
    End Property

    Friend ReadOnly Property CMTAdjustmentRateCode() As String
      Get
        Select Case CMTExcessPaymentType
          Case CmtExcessPayment.CMTExcessPaymentTypes.Refund
            Return mvCMTRefundRateCode
          Case CmtExcessPayment.CMTExcessPaymentTypes.Retain
            Return mvClassFields.Item(PaymentPlanDetailFields.odfRate).Value
          Case Else   'CmtExcessPayment.CMTExcessPaymentTypes.CarryForward, CmtExcessPayment.CMTExcessPaymentTypes.ReAnalyse
            Return ""
        End Select
      End Get
    End Property

    ''' <summary>Used by CMT to re-set prices to pro-rated figures when dealing with Detail line added part way through year.</summary>
    Friend Sub SetCMTOtherLinePartYearPrice()
      If (mvDetailType And PaymentPlanDetailTypes.ppdltOtherCharge) = PaymentPlanDetailTypes.ppdltOtherCharge Then
        If HasPriceInfo = True AndAlso UnitPrice <> 0 AndAlso GrossAmount <> UnitPrice Then
          'mvFullPrice = full annual cost but full price for this year is pro-rated, so reduce mvProratedPrice to be the correct figure
          'Full year=12,gross(part year)=8,pro-rate=5
          'vPrortedPrice = 8-(12-5) = 8-7 = 1
          Dim vPrortedPrice As Double = FixTwoPlaces(GrossAmount - (mvFullPrice - mvProratedPrice))
          If (GrossAmount > 0 AndAlso vPrortedPrice < 0) OrElse (GrossAmount < 0 AndAlso vPrortedPrice > 0) Then vPrortedPrice = 0
          mvFullPrice = GrossAmount
          mvProratedPrice = vPrortedPrice
        End If
      End If
    End Sub

    ''' <summary>Used by CMT to re-set prorted figures to account for rounding when advanced CMT is not being used.</summary>
    Friend Sub SetCMTOldMembershipProratedPrice(ByVal pProratedPrice As Double)
      Dim vVATRate As VatRate = mvEnv.VATRate(VatRateCode)
      Dim vVATAmount As Double = vVATRate.CalculateVATAmount(pProratedPrice, ProductRate.VatExclusive, VatPercentage)
      With mvClassFields
        .Item(PaymentPlanDetailFields.odfVatAmount).DoubleValue = vVATAmount
        If ProductRate.VatExclusive Then
          .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = FixTwoPlaces(pProratedPrice + vVATAmount)
        Else
          .Item(PaymentPlanDetailFields.odfGrossAmount).DoubleValue = pProratedPrice
        End If
        .Item(PaymentPlanDetailFields.odfNetAmount).DoubleValue = FixTwoPlaces(GrossAmount - VatAmount)
      End With
    End Sub

    Public Function IsValidOnDate(pDate As Date) As Boolean
      Dim vValid As Boolean = True
      If IsDate(ValidFrom) Then
        If CDate(ValidFrom) > pDate Then vValid = False
      End If
      If IsDate(ValidTo) Then
        If CDate(ValidTo) < pDate Then vValid = False
      End If
      Return vValid
    End Function

    ''' <summary>Used by CMT to ensure that the pricing data is correct when fixed amount is set.</summary>
    ''' <remarks>Earlier versions of the software sometimes incorrectly set the UnitPrice as zero when the ProductRate had a price of zero, 
    ''' when it should have been set to the FixedAmount.  This will correct the UnitPrice so that the CMT pro-rating calculations can be performed.</remarks>
    Friend Sub CMTUpdatePreviousPriceData()
      If HasPriceInfo = True AndAlso UnitPrice = 0 AndAlso GrossAmount <> 0 Then
        Dim vAmount As Nullable(Of Double)
        If Amount.Length > 0 Then
          vAmount = DoubleValue(Amount)
        ElseIf NetFixedAmount.Length > 0 Then
          vAmount = DoubleValue(NetFixedAmount)
        End If
        If vAmount.HasValue Then mvClassFields.Item(PaymentPlanDetailFields.odfUnitPrice).Value = vAmount.ToString
      End If
    End Sub

  End Class
End Namespace
