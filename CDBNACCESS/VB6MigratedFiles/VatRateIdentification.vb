

Namespace Access
  Public Class VatRateIdentification

    Public Enum VatRateIdentificationRecordSetTypes 'These are bit values
      vrirtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum VatRateIdentificationFields
      vrifAll = 0
      vrifProductVatCategory
      vrifContactVatCategory
      vrifVatRateCode
      vrifAmendedBy
      vrifAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Dim mvVatRateInitialised As Boolean
    Dim mvVatRate As New VatRate(mvEnv)

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "vat_rate_identification"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("product_vat_category")
          .Add("contact_vat_category")
          .Add("vat_rate")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(VatRateIdentificationFields.vrifProductVatCategory).SetPrimaryKeyOnly()
        mvClassFields.Item(VatRateIdentificationFields.vrifContactVatCategory).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As VatRateIdentificationFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(VatRateIdentificationFields.vrifAmendedOn).Value = TodaysDate()
      mvClassFields.Item(VatRateIdentificationFields.vrifAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As VatRateIdentificationRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = VatRateIdentificationRecordSetTypes.vrirtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "vri")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pProductVatCategory As String = "", Optional ByRef pContactVatCategory As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pProductVatCategory) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(VatRateIdentificationRecordSetTypes.vrirtAll) & " FROM vat_rate_identification vri WHERE product_vat_category = '" & pProductVatCategory & "' AND contact_vat_category = '" & pContactVatCategory & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, VatRateIdentificationRecordSetTypes.vrirtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As VatRateIdentificationRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(VatRateIdentificationFields.vrifProductVatCategory, vFields)
        .SetItem(VatRateIdentificationFields.vrifContactVatCategory, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And VatRateIdentificationRecordSetTypes.vrirtAll) = VatRateIdentificationRecordSetTypes.vrirtAll Then
          .SetItem(VatRateIdentificationFields.vrifVatRateCode, vFields)
          .SetItem(VatRateIdentificationFields.vrifAmendedBy, vFields)
          .SetItem(VatRateIdentificationFields.vrifAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(VatRateIdentificationFields.vrifAll)
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
        AmendedBy = mvClassFields.Item(VatRateIdentificationFields.vrifAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(VatRateIdentificationFields.vrifAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactVatCategory() As String
      Get
        ContactVatCategory = mvClassFields.Item(VatRateIdentificationFields.vrifContactVatCategory).Value
      End Get
    End Property

    Public ReadOnly Property ProductVatCategory() As String
      Get
        ProductVatCategory = mvClassFields.Item(VatRateIdentificationFields.vrifProductVatCategory).Value
      End Get
    End Property

    Public ReadOnly Property VatRateCode() As String
      Get
        VatRateCode = mvClassFields.Item(VatRateIdentificationFields.vrifVatRateCode).Value
      End Get
    End Property

    Public ReadOnly Property VATRate() As VatRate
      Get
        Dim vRecordSet As CDBRecordSet
        If mvVatRateInitialised = False Then
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT nominal_account, vat_rate, percentage, previous_percentage, rate_changed from vat_rates v WHERE v.vat_rate = '" & VatRateCode & "'")
          If vRecordSet.Fetch() = True Then
            mvVatRate.InitFromRecordSet(vRecordSet)
          End If
          mvVatRateInitialised = True
          vRecordSet.CloseRecordSet()
        End If
        VATRate = mvVatRate
      End Get
    End Property
  End Class
End Namespace
