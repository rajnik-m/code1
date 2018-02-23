

Namespace Access
  Public Class Organiser

    Public Enum OrganiserRecordSetTypes 'These are bit values
      orgtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum OrganiserFields
      ofAll = 0
      ofOrganiser
      ofOrganiserDesc
      ofContactNumber
      ofAddressNumber
      ofInvoiceContact
      ofInvoiceAddress
      ofAmendedBy
      ofAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvArrangementContact As Contact
    Private mvArrangementAddress As Address
    Private mvInvoiceContact As Contact
    Private mvInvoiceAddress As Address

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "organisers"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("organiser")
          .Add("organiser_desc")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("invoice_contact", CDBField.FieldTypes.cftLong)
          .Add("invoice_address", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(OrganiserFields.ofOrganiser).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As OrganiserFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(OrganiserFields.ofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(OrganiserFields.ofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As OrganiserRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = OrganiserRecordSetTypes.orgtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "o")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pOrganiser As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pOrganiser.Length > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(OrganiserRecordSetTypes.orgtAll) & " FROM organisers o WHERE organiser = '" & pOrganiser & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, OrganiserRecordSetTypes.orgtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As OrganiserRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(OrganiserFields.ofOrganiser, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And OrganiserRecordSetTypes.orgtAll) = OrganiserRecordSetTypes.orgtAll Then
          .SetItem(OrganiserFields.ofOrganiserDesc, vFields)
          .SetItem(OrganiserFields.ofContactNumber, vFields)
          .SetItem(OrganiserFields.ofAddressNumber, vFields)
          .SetItem(OrganiserFields.ofInvoiceContact, vFields)
          .SetItem(OrganiserFields.ofInvoiceAddress, vFields)
          .SetItem(OrganiserFields.ofAmendedBy, vFields)
          .SetItem(OrganiserFields.ofAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(OrganiserFields.ofAll)
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

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(OrganiserFields.ofAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(OrganiserFields.ofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(OrganiserFields.ofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(OrganiserFields.ofContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ArrangementContact() As Contact
      Get
        If mvArrangementContact Is Nothing Then
          If ContactNumber > 0 Then
            mvArrangementContact = New Contact(mvEnv)
            mvArrangementContact.Init(ContactNumber)
          End If
        End If
        ArrangementContact = mvArrangementContact
      End Get
    End Property
    Public ReadOnly Property ArrangementAddress() As Address
      Get
        If mvArrangementAddress Is Nothing Then
          If AddressNumber > 0 Then
            mvArrangementAddress = New Address(mvEnv)
            mvArrangementAddress.Init(AddressNumber)
          End If
        End If
        ArrangementAddress = mvArrangementAddress
      End Get
    End Property
    Public ReadOnly Property InvoiceContact() As Contact
      Get
        If mvInvoiceContact Is Nothing Then
          mvInvoiceContact = New Contact(mvEnv)
          mvInvoiceContact.Init(InvoiceContactNumber)
        End If
        InvoiceContact = mvInvoiceContact
      End Get
    End Property
    Public ReadOnly Property InvoiceAddress() As Address
      Get
        If mvInvoiceAddress Is Nothing Then
          mvInvoiceAddress = New Address(mvEnv)
          mvInvoiceAddress.Init(InvoiceAddressNumber)
        End If
        InvoiceAddress = mvInvoiceAddress
      End Get
    End Property

    Public ReadOnly Property InvoiceAddressNumber() As Integer
      Get
        InvoiceAddressNumber = mvClassFields.Item(OrganiserFields.ofInvoiceAddress).IntegerValue
      End Get
    End Property

    Public ReadOnly Property InvoiceContactNumber() As Integer
      Get
        InvoiceContactNumber = mvClassFields.Item(OrganiserFields.ofInvoiceContact).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OrganiserCode() As String
      Get
        OrganiserCode = mvClassFields.Item(OrganiserFields.ofOrganiser).Value
      End Get
    End Property

    Public ReadOnly Property OrganiserDesc() As String
      Get
        OrganiserDesc = mvClassFields.Item(OrganiserFields.ofOrganiserDesc).Value
      End Get
    End Property
  End Class
End Namespace
