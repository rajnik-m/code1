Namespace Access

  Partial Public Class Product

    Public Enum ProductRecordSetTypes 'These are bit values
      prstMain = 1
      'ADD additional recordset types here
    End Enum

    Private mvAllRates As CollectionList(Of ProductRate)
    Private mvProductRate As ProductRate
    Private mvFixedUnitRate As Boolean

    Protected Overrides Sub ClearFields()
      mvProductRate = Nothing
      mvAllRates = Nothing
    End Sub

    Public Overloads Function GetRecordSetFields(ByVal pRSType As ProductRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And ProductRecordSetTypes.prstMain) > 0 Then
        vFields = "p.product,product_desc,stock_item,activity,activity_value,product_vat_category,donation,subscription,postage_packing,p.warehouse,last_stock_count"
        If mvClassFields(ProductFields.SponsorshipEvent).InDatabase Then vFields = vFields & ",sponsorship_event"
        If mvClassFields(ProductFields.EligibleForGiftAid).InDatabase Then vFields = vFields & ",p.eligible_for_gift_aid"
        If mvClassFields(ProductFields.PackProduct).InDatabase Then vFields = vFields & ",pack_product"
        If mvClassFields(ProductFields.EligibleForGiftAid).InDatabase Then
          vFields = Replace(vFields, "p.eligible_for_gift_aid", "p.eligible_for_gift_aid AS p_eligible_for_gift_aid")
        End If
        If mvClassFields(ProductFields.WebPublish).InDatabase Then vFields = vFields & ",p.web_publish"
        If mvClassFields(ProductFields.AccruesInterest).InDatabase Then vFields &= ",accrues_interest"
        If mvClassFields(ProductFields.ActivityDurationMonths).InDatabase Then vFields &= ",p.activity_duration_months"
      End If
      Return vFields
    End Function

    Public Overloads Sub InitFromRecordSet(ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ProductRecordSetTypes)
      Dim vFields As CDBFields

      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ProductFields.Product, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ProductRecordSetTypes.prstMain) > 0 Then
          .SetItem(ProductFields.ProductDesc, vFields)
          .SetItem(ProductFields.StockItem, vFields)
          .SetItem(ProductFields.Activity, vFields)
          .SetItem(ProductFields.ActivityValue, vFields)
          .SetItem(ProductFields.ProductVatCategory, vFields)
          .SetItem(ProductFields.Donation, vFields)
          .SetItem(ProductFields.Subscription, vFields)
          .SetItem(ProductFields.PostagePacking, vFields)
          .SetItem(ProductFields.Warehouse, vFields)
          .SetItem(ProductFields.LastStockCount, vFields)
          .SetOptionalItem(ProductFields.SponsorshipEvent, vFields)
          If .Item(ProductFields.EligibleForGiftAid).InDatabase Then
            If vFields.Exists("p_eligible_for_gift_aid") Then
              .Item(ProductFields.EligibleForGiftAid).SetValue = vFields("p_eligible_for_gift_aid").Value
            End If
          End If
          .SetOptionalItem(ProductFields.PackProduct, vFields)
          .SetOptionalItem(ProductFields.WebPublish, vFields)
          .SetOptionalItem(ProductFields.AccruesInterest, vFields)
          .SetOptionalItem(ProductFields.ActivityDurationMonths, vFields)
        End If
      End With
    End Sub

    Public Sub InitWithRate(ByVal pEnv As CDBEnvironment, ByVal pProduct As String, ByVal pRate As String, Optional ByVal pFixedUnitRate As Boolean = False)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      mvProductRate = New ProductRate(mvEnv)
      mvProductRate.Init()
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields() & "," & mvProductRate.GetRecordSetFieldsForProduct() & " FROM products p, rates r WHERE p.product = '" & pProduct & "' AND p.product = r.product AND r.rate = '" & pRate & "'")
      If vRecordSet.Fetch() Then
        InitFromRecordSet(vRecordSet)
        mvProductRate = New ProductRate(mvEnv)
        'Use the web_publish value for the rate and not for the product
        If vRecordSet.Fields.ContainsKey("rate_web_publish") Then vRecordSet.Fields("web_publish").Value = vRecordSet.Fields("rate_web_publish").Value
        mvProductRate.InitFromRecordSet(vRecordSet)
        mvFixedUnitRate = pFixedUnitRate
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property ProductRate() As ProductRate
      Get
        If mvProductRate Is Nothing Then mvProductRate = New ProductRate(mvEnv)
        Return mvProductRate
      End Get
    End Property

    Public ReadOnly Property FixedUnitRate() As Boolean
      Get
        FixedUnitRate = mvFixedUnitRate
      End Get
    End Property

    Public ReadOnly Property Rates() As CollectionList(Of ProductRate)
      Get
        Dim vProductRate As ProductRate
        If mvAllRates Is Nothing Then
          mvAllRates = New CollectionList(Of ProductRate)
          vProductRate = New ProductRate(mvEnv)
          vProductRate.Init()
          Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vProductRate.GetRecordSetFields() & " FROM rates r WHERE product = '" & ProductCode & "'")
          While vRecordSet.Fetch
            vProductRate.InitFromRecordSet(vRecordSet)
            mvAllRates.Add(vProductRate.RateCode, vProductRate)
            vProductRate = New ProductRate(mvEnv)
          End While
          vRecordSet.CloseRecordSet()
        End If
        Return mvAllRates
      End Get
    End Property

    Public ReadOnly Property IsMembershipProduct() As Boolean
      Get
        Dim vCount As Integer
        Dim vWhere As String
        vWhere = "first_periods_product = '" & mvClassFields(ProductFields.Product).Value & "'"
        vWhere = vWhere & " OR subsequent_periods_product = '" & mvClassFields(ProductFields.Product).Value & "'"
        vCount = mvEnv.Connection.GetCount("membership_types", Nothing, vWhere)
        Return vCount > 0
      End Get
    End Property

  End Class

End Namespace
