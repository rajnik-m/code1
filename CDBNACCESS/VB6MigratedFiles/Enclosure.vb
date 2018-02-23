

Namespace Access
  Public Class Enclosure

    Public Enum EnclosureRecordSetTypes 'These are bit values
      encrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EnclosureFields
      efAll = 0
      efPaymentPlanNumber
      efContactNumber
      efProduct
      efQuantity
      efDateCreated
      efDateFulfilled
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvMailing As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "enclosures"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("product")
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("date_created", CDBField.FieldTypes.cftDate)
          .Add("date_fulfilled", CDBField.FieldTypes.cftDate)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EnclosureFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EnclosureRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EnclosureRecordSetTypes.encrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "e")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EnclosureRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And EnclosureRecordSetTypes.encrtAll) = EnclosureRecordSetTypes.encrtAll Then
          .SetItem(EnclosureFields.efPaymentPlanNumber, vFields)
          .SetItem(EnclosureFields.efContactNumber, vFields)
          .SetItem(EnclosureFields.efProduct, vFields)
          .SetItem(EnclosureFields.efQuantity, vFields)
          .SetItem(EnclosureFields.efDateCreated, vFields)
          .SetItem(EnclosureFields.efDateFulfilled, vFields)
        End If
      End With
    End Sub

    Public Sub InitDetails(ByRef pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pProductCode As String, ByRef pQuantity As Integer, ByRef pMailing As String)
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        If pContactNumber > 0 Then .Item(EnclosureFields.efContactNumber).IntegerValue = pContactNumber
        .Item(EnclosureFields.efProduct).Value = pProductCode
        .Item(EnclosureFields.efQuantity).IntegerValue = pQuantity
        .Item(EnclosureFields.efDateCreated).Value = TodaysDate()
      End With
      mvMailing = pMailing
    End Sub

    Public Sub SetPaymentPlan(ByRef pPaymentPlanNumber As Integer, ByRef pContactNumber As Integer)
      With mvClassFields
        .Item(EnclosureFields.efPaymentPlanNumber).IntegerValue = pPaymentPlanNumber
        If ContactNumber = 0 Then .Item(EnclosureFields.efContactNumber).IntegerValue = pContactNumber
      End With
    End Sub

    Public Sub SetFulfilled()
      mvClassFields.Item(EnclosureFields.efDateFulfilled).Value = TodaysDate()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EnclosureFields.efAll)
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

    Public ReadOnly Property Mailing() As String
      Get
        Mailing = mvMailing
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(EnclosureFields.efContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EnclosureFields.efContactNumber).IntegerValue = Value
      End Set
    End Property
    Public ReadOnly Property DateCreated() As String
      Get
        DateCreated = mvClassFields.Item(EnclosureFields.efDateCreated).Value
      End Get
    End Property

    Public ReadOnly Property DateFulfilled() As String
      Get
        DateFulfilled = mvClassFields.Item(EnclosureFields.efDateFulfilled).Value
      End Get
    End Property

    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(EnclosureFields.efPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(EnclosureFields.efProduct).Value
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(EnclosureFields.efQuantity).IntegerValue
      End Get
    End Property
  End Class
End Namespace
