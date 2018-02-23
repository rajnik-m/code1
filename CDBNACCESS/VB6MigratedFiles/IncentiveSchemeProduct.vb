

Namespace Access
  Public Class IncentiveSchemeProduct

    Public Enum IncentiveSchemeProductRecordSetTypes 'These are bit values
      isprtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum IncentiveSchemeProductFields
      ispfAll = 0
      ispfIncentiveScheme
      ispfReasonForDespatch
      ispfSequenceNumber
      ispfIncentiveType
      ispfForWhom
      ispfIncentiveDesc
      ispfProduct
      ispfRate
      ispfQuantity
      ispfBasic
      ispfAmendedBy
      ispfAmendedOn
      ispfIgnoreProductAndRate
      ispfMinimumQuantity
      ispfMaximumQuantity
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvProductDesc As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "incentive_scheme_products"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("incentive_scheme")
          .Add("reason_for_despatch")
          .Add("sequence_number", CDBField.FieldTypes.cftInteger)
          .Add("incentive_type")
          .Add("for_whom")
          .Add("incentive_desc")
          .Add("product")
          .Add("rate")
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("basic")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("ignore_product_and_rate")
          .Add("minimum_quantity", CDBField.FieldTypes.cftInteger)
          .Add("maximum_quantity", CDBField.FieldTypes.cftInteger)
          .Item(IncentiveSchemeProductFields.ispfIncentiveScheme).SetPrimaryKeyOnly()
          .Item(IncentiveSchemeProductFields.ispfReasonForDespatch).SetPrimaryKeyOnly()
          .Item(IncentiveSchemeProductFields.ispfSequenceNumber).SetPrimaryKeyOnly()
          .Item(IncentiveSchemeProductFields.ispfMinimumQuantity).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIncentiveProductMinMax)
          .Item(IncentiveSchemeProductFields.ispfMaximumQuantity).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIncentiveProductMinMax)
          .Item(IncentiveSchemeProductFields.ispfIncentiveScheme).SetPrimaryKeyOnly()
          .Item(IncentiveSchemeProductFields.ispfReasonForDespatch).SetPrimaryKeyOnly()
          .Item(IncentiveSchemeProductFields.ispfSequenceNumber).SetPrimaryKeyOnly()

          .Item(IncentiveSchemeProductFields.ispfIncentiveScheme).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfReasonForDespatch).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfProduct).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfRate).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfMinimumQuantity).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfMaximumQuantity).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfAmendedBy).PrefixRequired = True
          .Item(IncentiveSchemeProductFields.ispfAmendedOn).PrefixRequired = True

        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As IncentiveSchemeProductFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(IncentiveSchemeProductFields.ispfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(IncentiveSchemeProductFields.ispfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Public Function GetRecordSetFields(ByVal pRSType As IncentiveSchemeProductRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = IncentiveSchemeProductRecordSetTypes.isprtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "isp")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pIncentiveScheme As String = "", Optional ByRef pReasonForDespatch As String = "", Optional ByRef pSequenceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pIncentiveScheme) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(IncentiveSchemeProductRecordSetTypes.isprtAll) & " FROM incentive_scheme_products isp WHERE incentive_scheme = '" & pIncentiveScheme & "' AND reason_for_despatch = '" & pReasonForDespatch & "' AND sequence_number = " & pSequenceNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, IncentiveSchemeProductRecordSetTypes.isprtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As IncentiveSchemeProductRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(IncentiveSchemeProductFields.ispfIncentiveScheme, vFields)
        .SetItem(IncentiveSchemeProductFields.ispfReasonForDespatch, vFields)
        .SetItem(IncentiveSchemeProductFields.ispfSequenceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And IncentiveSchemeProductRecordSetTypes.isprtAll) = IncentiveSchemeProductRecordSetTypes.isprtAll Then
          .SetItem(IncentiveSchemeProductFields.ispfIncentiveType, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfForWhom, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfIncentiveDesc, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfProduct, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfRate, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfQuantity, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfBasic, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfAmendedBy, vFields)
          .SetItem(IncentiveSchemeProductFields.ispfAmendedOn, vFields)
          .SetOptionalItem(IncentiveSchemeProductFields.ispfIgnoreProductAndRate, vFields)
          .SetOptionalItem(IncentiveSchemeProductFields.ispfMinimumQuantity, vFields)
          .SetOptionalItem(IncentiveSchemeProductFields.ispfMaximumQuantity, vFields)
        End If
      End With
      mvProductDesc = pRecordSet.Fields.FieldExists("product_desc").Value
    End Sub

    Public Sub InitProductQty(ByRef pEnv As CDBEnvironment, ByRef pProductCode As String, ByRef pQuantity As Integer)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields.Item(IncentiveSchemeProductFields.ispfProduct).Value = pProductCode
      mvClassFields.Item(IncentiveSchemeProductFields.ispfQuantity).IntegerValue = pQuantity
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(IncentiveSchemeProductFields.ispfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public ReadOnly Property IncentiveScheme() As String
      Get
        IncentiveScheme = mvClassFields.Item(IncentiveSchemeProductFields.ispfIncentiveScheme).Value
      End Get
    End Property

    Public ReadOnly Property ReasonForDespatch() As String
      Get
        ReasonForDespatch = mvClassFields.Item(IncentiveSchemeProductFields.ispfReasonForDespatch).Value
      End Get
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvClassFields.Item(IncentiveSchemeProductFields.ispfSequenceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property IncentiveType() As String
      Get
        IncentiveType = mvClassFields.Item(IncentiveSchemeProductFields.ispfIncentiveType).Value
      End Get
    End Property

    Public ReadOnly Property ForWhom() As String
      Get
        ForWhom = mvClassFields.Item(IncentiveSchemeProductFields.ispfForWhom).Value
      End Get
    End Property

    Public ReadOnly Property IncentiveDesc() As String
      Get
        IncentiveDesc = mvClassFields.Item(IncentiveSchemeProductFields.ispfIncentiveDesc).Value
      End Get
    End Property

    Public ReadOnly Property MinimumQuantity() As String
      Get
        MinimumQuantity = mvClassFields.Item(IncentiveSchemeProductFields.ispfMinimumQuantity).Value
      End Get
    End Property

    Public ReadOnly Property MaximumQuantity() As String
      Get
        MaximumQuantity = mvClassFields.Item(IncentiveSchemeProductFields.ispfMaximumQuantity).Value
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(IncentiveSchemeProductFields.ispfProduct).Value
      End Get
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(IncentiveSchemeProductFields.ispfRate).Value
      End Get
    End Property

    Public Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(IncentiveSchemeProductFields.ispfQuantity).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(IncentiveSchemeProductFields.ispfQuantity).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property Basic() As Boolean
      Get
        Basic = mvClassFields.Item(IncentiveSchemeProductFields.ispfBasic).Bool
      End Get
    End Property

    Public ReadOnly Property ProductDesc() As String
      Get
        ProductDesc = mvProductDesc
      End Get
    End Property

    Public ReadOnly Property IgnoreProductAndRate() As Boolean
      Get
        IgnoreProductAndRate = mvClassFields(IncentiveSchemeProductFields.ispfIgnoreProductAndRate).Bool
      End Get
    End Property
  End Class
End Namespace
