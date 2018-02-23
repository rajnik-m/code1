

Namespace Access
  Public Class PackedProduct

    Public Enum PackedProductRecordSetTypes 'These are bit values
      pprtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PackedProductFields
      ppfAll = 0
      ppfProduct
      ppfRate
      ppfLinkProductCode
      ppfBaseRateCode
      ppfAmendedBy
      ppfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvBaseRateInitialised As Boolean
    Private mvBaseRate As New ProductRate(mvEnv)
    Private mvLinkProductInitialised As Boolean
    Private mvLinkProduct As New Product(mvEnv)

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "packed_products"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("product")
          .Add("rate")
          .Add("link_product")
          .Add("base_rate")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(PackedProductFields.ppfProduct).SetPrimaryKeyOnly()
        mvClassFields.Item(PackedProductFields.ppfRate).SetPrimaryKeyOnly()
        mvClassFields.Item(PackedProductFields.ppfLinkProductCode).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As PackedProductFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(PackedProductFields.ppfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PackedProductFields.ppfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PackedProductRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PackedProductRecordSetTypes.pprtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pp")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pProduct As Integer = 0, Optional ByRef pRate As Integer = 0, Optional ByRef pLinkProductCode As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pProduct > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PackedProductRecordSetTypes.pprtAll) & " FROM packed_products pp WHERE product = " & pProduct & " AND rate = " & pRate & " AND link_product = " & pLinkProductCode)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PackedProductRecordSetTypes.pprtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PackedProductRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PackedProductFields.ppfProduct, vFields)
        .SetItem(PackedProductFields.ppfRate, vFields)
        .SetItem(PackedProductFields.ppfLinkProductCode, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PackedProductRecordSetTypes.pprtAll) = PackedProductRecordSetTypes.pprtAll Then
          .SetItem(PackedProductFields.ppfBaseRateCode, vFields)
          .SetItem(PackedProductFields.ppfAmendedBy, vFields)
          .SetItem(PackedProductFields.ppfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PackedProductFields.ppfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(PackedProductFields.ppfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PackedProductFields.ppfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BaseRateCode() As String
      Get
        BaseRateCode = mvClassFields.Item(PackedProductFields.ppfBaseRateCode).Value
      End Get
    End Property
    Public ReadOnly Property LinkProduct() As Product
      Get
        If mvLinkProductInitialised = False Then
          mvLinkProduct.Init(LinkProductCode)
          mvLinkProductInitialised = True
        End If
        LinkProduct = mvLinkProduct
      End Get
    End Property
    Public ReadOnly Property BaseRate() As ProductRate
      Get
        If mvBaseRateInitialised = False Then
          mvBaseRate.Init(LinkProductCode, BaseRateCode)
          mvBaseRateInitialised = True
        End If
        BaseRate = mvBaseRate
      End Get
    End Property
    Public ReadOnly Property LinkProductCode() As String
      Get
        LinkProductCode = mvClassFields.Item(PackedProductFields.ppfLinkProductCode).Value
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(PackedProductFields.ppfProduct).Value
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(PackedProductFields.ppfRate).Value
      End Get
    End Property
  End Class
End Namespace
