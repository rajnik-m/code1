Namespace Access

  Partial Public Class PaymentPlanDetailPricing
    Inherits CARERecord

    '--------------------------------------------------
    'Enum defining all the fields
    '--------------------------------------------------
    Private Enum PaymentPlanDetailPricingFields
      AllFields = 0
      ModifierActivity
      ModifierActivityValue
      ModifierActivityQuantity
      ModifierActivityDate
      ModifierPrice
      ModifierPerItem
      UnitPrice
      ProRated
      NetAmount
      VatAmount
      GrossAmount
      VatRate
      VatPercentage
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields

        .Add("modifier_activity")
        .Add("modifier_activity_value")
        .Add("modifier_activity_quantity", CDBField.FieldTypes.cftNumeric)
        .Add("modifier_activity_date", CDBField.FieldTypes.cftDate)
        .Add("modifier_price", CDBField.FieldTypes.cftNumeric)
        .Add("modifier_per_item")
        .Add("unit_price", CDBField.FieldTypes.cftNumeric)
        .Add("pro_rated")
        .Add("net_amount", CDBField.FieldTypes.cftNumeric)
        .Add("vat_amount", CDBField.FieldTypes.cftNumeric)
        .Add("gross_amount", CDBField.FieldTypes.cftNumeric)
        .Add("vat_rate")
        .Add("vat_percentage", CDBField.FieldTypes.cftNumeric)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ppdr"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "order_details"
      End Get
    End Property
    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      'Do Nothing
    End Sub
    Public Overrides Sub Update(pParameterList As CDBParameters)
      'Do Nothing
    End Sub

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

#Region "Properties"
    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ModifierActivity() As String
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierActivity).Value
      End Get
    End Property
    Public ReadOnly Property ModifierActivityValue() As String
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property ModifierActivityQuantity() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierActivityQuantity).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ModifierActivityDate() As String
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierActivityDate).Value
      End Get
    End Property
    Public ReadOnly Property ModifierPrice() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierPrice).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ModifierPerItem() As String
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ModifierPerItem).Value
      End Get
    End Property
    Public ReadOnly Property UnitPrice() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.UnitPrice).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ProRated() As Boolean
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.ProRated).Bool
      End Get
    End Property
    Public ReadOnly Property NetAmount() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.NetAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property VatAmount() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.VatAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property GrossAmount() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.GrossAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property VatRate() As String
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.VatRate).Value
      End Get
    End Property
    Public ReadOnly Property VatPercentage() As Double
      Get
        Return mvClassFields(PaymentPlanDetailPricingFields.VatPercentage).DoubleValue
      End Get
    End Property
#End Region

    Public Sub SetModifierData(ByVal pRateModifier As RateModifier, ByVal pModifierActivityQuantity As Double, ByVal pModifierActivityDate As String, ByVal pModifierPrice As Double, ByVal pUnitPrice As Double)
      If pModifierPrice <> 0 OrElse mvEnv.GetConfigOption("fp_zero_modifiers_significant", False) Then
        If Not mvExisting Then
          'Initialise values from Rate Modifier, Modifier Activity and Product Rate
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivity).Value = pRateModifier.Activity
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivityValue).Value = pRateModifier.ActivityValue
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivityQuantity).DoubleValue = pModifierActivityQuantity
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivityDate).Value = pModifierActivityDate
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierPrice).DoubleValue = pModifierPrice
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierPerItem).Value = pRateModifier.PerItem
          mvClassFields.Item(PaymentPlanDetailPricingFields.UnitPrice).DoubleValue = pUnitPrice
        Else
          mvClassFields.ClearItems()
          'Set Values to 'Multiple'
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivity).Value = "MULTI"
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierActivityValue).Value = "N/A"
          mvClassFields.Item(PaymentPlanDetailPricingFields.ModifierPerItem).Value = "M"
          mvClassFields.Item(PaymentPlanDetailPricingFields.UnitPrice).DoubleValue = pUnitPrice 'This is already the sum of the unit prices
        End If
        mvExisting = True
      End If
    End Sub

    Public Sub CalculatePricing(ByVal pUnitPrice As Double, ByVal pActualPrice As Double, ByVal pVatExclusive As Boolean, ByVal pTransactionDate As String, ByVal pVatRate As VatRate, ByVal pProRated As Boolean)
      mvClassFields.Item(PaymentPlanDetailPricingFields.UnitPrice).DoubleValue = pUnitPrice
      mvClassFields.Item(PaymentPlanDetailPricingFields.VatAmount).DoubleValue = pVatRate.CalculateVATAmount(pActualPrice, pVatExclusive, pTransactionDate)
      If pVatExclusive Then
        mvClassFields.Item(PaymentPlanDetailPricingFields.GrossAmount).DoubleValue = FixTwoPlaces(pActualPrice + VatAmount)
      Else
        mvClassFields.Item(PaymentPlanDetailPricingFields.GrossAmount).DoubleValue = pActualPrice
      End If
      mvClassFields.Item(PaymentPlanDetailPricingFields.NetAmount).DoubleValue = FixTwoPlaces(GrossAmount - VatAmount)

      mvClassFields.Item(PaymentPlanDetailPricingFields.ProRated).Bool = pProRated
      mvClassFields.Item(PaymentPlanDetailPricingFields.VatRate).Value = pVatRate.VatRateCode
      mvClassFields.Item(PaymentPlanDetailPricingFields.VatPercentage).DoubleValue = pVatRate.Percentage
    End Sub

    Public Sub GetDataAsParameters(ByRef pParams As CDBParameters)
      If pParams Is Nothing Then pParams = New CDBParameters
      For Each vField As ClassField In mvClassFields
        pParams.Add(ProperName((vField.Name)), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(vField.Value), vField.Value))
      Next
    End Sub

  End Class
End Namespace
