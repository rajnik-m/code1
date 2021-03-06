Namespace Access

  Public Class DespatchNote
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum DespatchNoteFields
      AllFields = 0
      PickingListNumber
      DespatchNoteNumber
      InvoiceNumber
      DespatchMethod
      DespatchDate
      OrderDate
      DeliveryCharge
      CarrierReference
      BatchNumber
      TransactionNumber
      ContactNumber
      AddressNumber
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("picking_list_number", CDBField.FieldTypes.cftLong)
        .Add("despatch_note_number", CDBField.FieldTypes.cftInteger)
        .Add("invoice_number", CDBField.FieldTypes.cftLong)
        .Add("despatch_method")
        .Add("despatch_date", CDBField.FieldTypes.cftDate)
        .Add("order_date", CDBField.FieldTypes.cftDate)
        .Add("delivery_charge", CDBField.FieldTypes.cftNumeric)
        .Add("carrier_reference", CDBField.FieldTypes.cftMemo)
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)

        .Item(DespatchNoteFields.PickingListNumber).PrimaryKey = True

        .Item(DespatchNoteFields.DespatchNoteNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "dn"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "despatch_notes"
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
    Public ReadOnly Property PickingListNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.PickingListNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property DespatchNoteNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.DespatchNoteNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property InvoiceNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.InvoiceNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property DespatchMethod() As String
      Get
        Return mvClassFields(DespatchNoteFields.DespatchMethod).Value
      End Get
    End Property
    Public ReadOnly Property DespatchDate() As String
      Get
        Return mvClassFields(DespatchNoteFields.DespatchDate).Value
      End Get
    End Property
    Public ReadOnly Property OrderDate() As String
      Get
        Return mvClassFields(DespatchNoteFields.OrderDate).Value
      End Get
    End Property
    Public ReadOnly Property DeliveryCharge() As Double
      Get
        Return mvClassFields(DespatchNoteFields.DeliveryCharge).DoubleValue
      End Get
    End Property
    Public ReadOnly Property CarrierReference() As String
      Get
        Return mvClassFields(DespatchNoteFields.CarrierReference).Value
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(DespatchNoteFields.AddressNumber).IntegerValue
      End Get
    End Property
#End Region

  End Class
End Namespace
