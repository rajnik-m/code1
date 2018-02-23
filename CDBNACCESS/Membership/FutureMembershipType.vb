Namespace Access

  Public Class FutureMembershipType
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum FutureMembershipTypeFields
      AllFields = 0
      MembershipNumber
      FutureMembershipType
      FutureChangeDate
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
        .Add("membership_number", CDBField.FieldTypes.cftInteger)
        .Add("future_membership_type")
        .Add("future_change_date", CDBField.FieldTypes.cftDate)
        .Add("product")
        .Add("rate")
        .Add("amount", CDBField.FieldTypes.cftNumeric)

        .Item(FutureMembershipTypeFields.MembershipNumber).PrimaryKey = True
        .Item(FutureMembershipTypeFields.MembershipNumber).PrefixRequired = True

        .Item(FutureMembershipTypeFields.Amount).PrefixRequired = True
        .Item(FutureMembershipTypeFields.FutureMembershipType).PrefixRequired = True
        .Item(FutureMembershipTypeFields.Product).PrefixRequired = True
        .Item(FutureMembershipTypeFields.Rate).PrefixRequired = True
      End With
    End Sub

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvMembershipType = Nothing
      mvProduct = Nothing
      mvProductRate = Nothing
    End Sub

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      FutureChangeDate = TodaysDate()
    End Sub

    Public Overrides Sub Update(pParameterList As CDBParameters)
      MyBase.Update(pParameterList)
      If mvClassFields.Item(FutureMembershipTypeFields.FutureMembershipType).ValueChanged Then mvMembershipType = Nothing
      If mvClassFields.Item(FutureMembershipTypeFields.Product).ValueChanged Then
        mvProduct = Nothing
        mvProductRate = Nothing
      End If
      If mvClassFields.Item(FutureMembershipTypeFields.Rate).ValueChanged Then mvProductRate = Nothing
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "fmt"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "member_future_type"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property MembershipNumber() As Integer
      Get
        Return mvClassFields(FutureMembershipTypeFields.MembershipNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MembershipTypeCode() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.FutureMembershipType).Value
      End Get
    End Property
    Public Property FutureChangeDate() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.FutureChangeDate).Value
      End Get
      Private Set(value As String)
        mvClassFields(FutureMembershipTypeFields.FutureChangeDate).Value = value
      End Set
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property FutureMembershipProductCode() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.Product).Value
      End Get
    End Property
    Public ReadOnly Property FutureMembershipRateCode() As String
      Get
        Return mvClassFields(FutureMembershipTypeFields.Rate).Value
      End Get
    End Property
    Public ReadOnly Property FutureMembershipAmount() As Double
      Get
        Return mvClassFields(FutureMembershipTypeFields.Amount).DoubleValue
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Private mvMembershipType As MembershipType
    Private mvProduct As Product
    Private mvProductRate As ProductRate

    Public ReadOnly Property FutureMembershipProduct() As Product
      Get
        If mvProduct Is Nothing Then
          mvProduct = New Product(mvEnv)
          mvProduct.Init(FutureMembershipProductCode)
        End If
        Return mvProduct
      End Get
    End Property

    Public ReadOnly Property FutureMembershipProductRate As ProductRate
      Get
        If mvProductRate Is Nothing Then
          mvProductRate = New ProductRate(mvEnv)
          mvProductRate.Init(FutureMembershipProductCode, FutureMembershipRateCode)
        End If
        Return mvProductRate
      End Get
    End Property

    Public ReadOnly Property MembershipType() As MembershipType
      Get
        If mvMembershipType Is Nothing Then
          If MembershipTypeCode.Length > 0 Then
            mvMembershipType = mvEnv.MembershipType(MembershipTypeCode)
          End If
        End If
        MembershipType = mvMembershipType
      End Get
    End Property

    ''' <summary>When updating a membership from one future type to another this is used to store the new future type details in order to update the Payment Plan.
    ''' This method will be deprecated at some point and should NOT be used.</summary>
    Public Sub InitFromType(ByVal pEnv As CDBEnvironment, ByVal pMembershipType As String, ByVal pMembershipNumber As Integer)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvClassFields.Item(FutureMembershipTypeFields.MembershipNumber).Value = pMembershipNumber.ToString
      mvClassFields.Item(FutureMembershipTypeFields.FutureMembershipType).Value = pMembershipType
    End Sub

    ''' <summary>Calculate the future renewal amount for this future membership type.</summary>
    ''' <param name="pMember">The <see cref="Member">Member</see> to be used when the price is to be calculated using Rate Modifiers.</param>
    ''' <returns>The calculated future renewal amount.</returns>
    ''' <remarks>The routine will just calculate the price and return it.  The price is not stored at all.</remarks>
    Public Function GetFutureRenewalAmount(ByVal pMember As Member) As Double
      If MembershipTypeCode.Length = 0 OrElse IsDate(FutureChangeDate) = False Then
        Throw New InvalidOperationException("FutureMembershipType has not been initialised")
      End If

      Dim vPriceCalculationDate As Date = CDate(FutureChangeDate)
      Dim vProductRate As ProductRate = FutureMembershipProductRate
      If FutureMembershipProductCode.Length = 0 OrElse FutureMembershipRateCode.Length = 0 Then
        'We don't have the Product / Rate codes so use defaults from the Membership Type
        vProductRate = New ProductRate(mvEnv)
        vProductRate.Init(MembershipType.FirstPeriodsProduct, MembershipType.FirstPeriodsRate)
      End If

      Dim vRenewalAmount As Double = 0

      'First calculate price of membership charging line
      vRenewalAmount = vProductRate.Price(vPriceCalculationDate, pMember.Contact)
      If vRenewalAmount = 0 Then vRenewalAmount = FutureMembershipAmount 'Price is fixed at this amount
      Dim vNonDiscountPrice As Double = vRenewalAmount      'This is the prices without any discounts used when calculating discounts

      'Second calculate price of each entitlement
      Dim vEntitlementPrice As Double = 0
      Dim vDiscount As Double = 0
      For Each vEntitlement As MembershipEntitlement In MembershipType.Entitlements(vProductRate.RateCode, True, True)
        vEntitlementPrice = vEntitlement.ProductRate.Price(vPriceCalculationDate, pMember.Contact)

        If vEntitlement.ProductRate.PriceIsPercentage.Equals("T", StringComparison.InvariantCultureIgnoreCase) Then
          'Calculate discount on the non-percentage total
          vDiscount = FixTwoPlaces(vEntitlementPrice)   'Calculated price is actually a discount percentage
          vEntitlementPrice = FixTwoPlaces(vNonDiscountPrice * (vDiscount / 100)) * -1
        ElseIf vEntitlement.ProductRate.PriceIsPercentage.Equals("P", StringComparison.InvariantCultureIgnoreCase) Then
          'Calculate discount on the previous total
          vDiscount = FixTwoPlaces(vEntitlementPrice)   'Calculated price is actually a discount percentage
          vEntitlementPrice = FixTwoPlaces(vRenewalAmount * (vDiscount / 100)) * -1
        Else
          vNonDiscountPrice += vEntitlementPrice
        End If

        vRenewalAmount += vEntitlementPrice
      Next

      'Third make sure renewal amount is just two-decimal places
      vRenewalAmount = FixTwoPlaces(vRenewalAmount)

      Return vRenewalAmount

    End Function

#End Region

  End Class
End Namespace
