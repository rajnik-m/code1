Namespace Access

  Partial Public Class ProductRate

    Private mvPackedProductsInitialised As Boolean
    Private mvPackedProducts As Collection

    Protected Overrides Sub ClearFields()
      mvPackedProductsInitialised = False
      mvPackedProducts = Nothing
      mvFinalPrice = 0
      mvFinalRenewalPrice = 0
      mvPrevContactNo = 0
      mvPrevTransDate = Nothing
      mvPrevRenewalCFDateType = RenewalCurrentFutureDateTypes.rcfdtNone
      mvNoVatExclusiveVatRequired = False
      mvPaymentPlanDetailPricing = Nothing
    End Sub

    Public Function GetRecordSetFieldsForProduct() As String
      Dim vFields As String = GetRecordSetFields()
      vFields = vFields.Replace(",r.amended_on", "")
      vFields = vFields.Replace(",r.amended_by", "")
      vFields = vFields.Replace(",history_only", "")
      vFields = vFields.Replace("r.product,", "")
      vFields = vFields.Replace(",r.web_publish", ",r.web_publish AS rate_web_publish") 'web_publish is also an attibute of products table
      Return vFields
    End Function

    Public ReadOnly Property PackedProducts() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vWhereFields As New CDBFields
        Dim vPackedProduct As New PackedProduct

        If mvPackedProductsInitialised = False Then
          vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, ProductCode, CDBField.FieldWhereOperators.fwoEqual)
          vWhereFields.Add("rate", CDBField.FieldTypes.cftCharacter, RateCode, CDBField.FieldWhereOperators.fwoEqual)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vPackedProduct.GetRecordSetFields(PackedProduct.PackedProductRecordSetTypes.pprtAll) & " FROM packed_products WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          While vRecordSet.Fetch
            vPackedProduct = New PackedProduct
            vPackedProduct.InitFromRecordSet(mvEnv, vRecordSet, PackedProduct.PackedProductRecordSetTypes.pprtAll)
            mvPackedProducts.Add(vPackedProduct)
          End While
          vRecordSet.CloseRecordSet()
          mvPackedProductsInitialised = True
        End If
        Return mvPackedProducts
      End Get
    End Property

    Public ReadOnly Property PackVatAmount(ByVal pPrice As Double, ByVal pQuantity As Integer, ByVal pContactVatCategory As String, ByVal pTransactionDate As String) As Double
      Get
        'Only relevant for Pack Products
        Dim vLinePrice As Double
        Dim vLineVATAmount As Double
        Dim vTotalVATAmount As Double
        Dim vPackedProduct As New PackedProduct
        Dim vFullPrice As Double
        Dim vPriceProportion As Double
        Dim vVRI As New VatRateIdentification

        'First sum up Full Price if Pack Products had been purchased individually
        Dim vVATRate As VatRate
        For Each vPackedProduct In PackedProducts
          vVATRate = mvEnv.VATRate(vPackedProduct.LinkProduct.ProductVatCategory, pContactVatCategory)
          vFullPrice = vFullPrice + vPackedProduct.BaseRate.Price(0, vVATRate)
        Next vPackedProduct

        'Next what proportion of Full Price is being charged?
        If vFullPrice > 0 Then
          vPriceProportion = pPrice / vFullPrice
          'Now work out individual line VAT Amounts
          For Each vPackedProduct In PackedProducts
            vVRI = New VatRateIdentification
            vVRI.Init(mvEnv, vPackedProduct.LinkProduct.ProductVatCategory, pContactVatCategory)
            vLinePrice = (vPackedProduct.BaseRate.Price(CDate(TodaysDate()), 0, pQuantity, vVRI.VATRate)) * vPriceProportion
            vLineVATAmount = Int(((vLinePrice - (vLinePrice / (1 + vVRI.VATRate.CurrentPercentage(pTransactionDate) / 100))) * 100) + 0.5) / 100
            vTotalVATAmount = vTotalVATAmount + vLineVATAmount
          Next vPackedProduct
        End If
        PackVatAmount = vTotalVATAmount
      End Get
    End Property

  End Class

End Namespace
