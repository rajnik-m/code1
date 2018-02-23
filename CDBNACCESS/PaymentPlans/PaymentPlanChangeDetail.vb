Namespace Access

  Public Class PaymentPlanChangeDetail
    Inherits CARERecord

    '--------------------------------------------------
    'Enum defining all the fields
    '--------------------------------------------------
    Private Enum PaymentPlanChangeDetailFields
      AllFields = 0
      PaymentPlanChangeNumber
      ChangeLineNumber
      PaymentPlanNumber
      Product
      Rate
      Amount
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields

        .Add("payment_plan_change_number", CDBField.FieldTypes.cftLong)
        .Add("change_line_number", CDBField.FieldTypes.cftLong)
        .Add("payment_plan_number", CDBField.FieldTypes.cftLong)
        .Add("product")
        .Add("rate")
        .Add("amount", CDBField.FieldTypes.cftNumeric)

        .Item(PaymentPlanChangeDetailFields.PaymentPlanChangeNumber).PrimaryKey = True
        .Item(PaymentPlanChangeDetailFields.ChangeLineNumber).PrimaryKey = True

        .Item(PaymentPlanChangeDetailFields.PaymentPlanChangeNumber).PrefixRequired = True
        .Item(PaymentPlanChangeDetailFields.PaymentPlanNumber).PrefixRequired = True

      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy As Boolean
      Get
        Return True
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias As String
      Get
        Return "ppcd"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName As String
      Get
        Return "payment_plan_change_details"
      End Get
    End Property

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
    Public ReadOnly Property PaymentPlanChangeNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.PaymentPlanNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ChangeLineNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.ChangeLineNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.PaymentPlanNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ProductCode() As String
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.Product).Value
      End Get
    End Property
    Public ReadOnly Property RateCode() As String
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.Rate).Value
      End Get
    End Property
    Public ReadOnly Property Amount() As Double
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.Amount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(PaymentPlanChangeDetailFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Methods"

    Public Overloads Sub Create(pPaymentPlanChange As PaymentPlanChange, pProduct As String, pRate As String, pLineNumber As Integer, pAmount As Double)
      mvClassFields(PaymentPlanChangeDetailFields.PaymentPlanChangeNumber).IntegerValue = pPaymentPlanChange.PaymentPlanChangeNumber
      mvClassFields(PaymentPlanChangeDetailFields.PaymentPlanNumber).IntegerValue = pPaymentPlanChange.PaymentPlanNumber
      mvClassFields(PaymentPlanChangeDetailFields.ChangeLineNumber).IntegerValue = pLineNumber
      mvClassFields(PaymentPlanChangeDetailFields.Product).Value = pProduct
      mvClassFields(PaymentPlanChangeDetailFields.Rate).Value = pRate
      mvClassFields(PaymentPlanChangeDetailFields.Amount).DoubleValue = pAmount
    End Sub

#End Region

  End Class
End Namespace