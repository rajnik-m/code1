

Namespace Access
  Public Class ContactAddressUsage

    Public Enum ContactAddressUsageRecordSetTypes 'These are bit values
      caurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    Public Enum ContactAddresssUsageLinkTypes
      caultContact
      caultOrganisation
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ContactAddressUsageFields
      caufAll = 0
      caufAddressNumber
      caufContactNumber
      caufAddressUsage
      caufNotes
      caufAmendedBy
      caufAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvLinkType As ContactAddresssUsageLinkTypes

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("address_number", CDBField.FieldTypes.cftLong)
          If mvLinkType = ContactAddressUsage.ContactAddresssUsageLinkTypes.caultOrganisation Then
            .DatabaseTableName = "organisation_address_usages"
            .Add("organisation_number", CDBField.FieldTypes.cftLong)
          Else
            .DatabaseTableName = "contact_address_usages"
            .Add("contact_number", CDBField.FieldTypes.cftLong)
          End If
          .Add("address_usage")
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(ContactAddressUsageFields.caufAddressNumber).SetPrimaryKeyOnly()
          .Item(ContactAddressUsageFields.caufContactNumber).SetPrimaryKeyOnly()
          .Item(ContactAddressUsageFields.caufAddressUsage).SetPrimaryKeyOnly()
          .Item(ContactAddressUsageFields.caufAddressUsage).PrefixRequired = True
          .Item(ContactAddressUsageFields.caufAmendedBy).PrefixRequired = True
          .Item(ContactAddressUsageFields.caufAmendedOn).PrefixRequired = True

        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ContactAddressUsageFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ContactAddressUsageFields.caufAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ContactAddressUsageFields.caufAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ContactAddressUsageRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ContactAddressUsageRecordSetTypes.caurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cau")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pLinkType As ContactAddresssUsageLinkTypes, Optional ByRef pAddressNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0, Optional ByRef pAddressUsage As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Not mvClassFields Is Nothing And pLinkType <> mvLinkType Then mvClassFields = Nothing
      mvLinkType = pLinkType
      InitClassFields()
      If pAddressNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactAddressUsageRecordSetTypes.caurtAll) & " FROM " & mvClassFields.DatabaseTableName & " cau WHERE address_number = " & pAddressNumber & " AND " & mvClassFields(ContactAddressUsageFields.caufContactNumber).Name & " = " & pContactNumber & " AND address_usage = '" & pAddressUsage & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ContactAddressUsageRecordSetTypes.caurtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactAddressUsageRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ContactAddressUsageFields.caufAddressNumber, vFields)
        .SetItem(ContactAddressUsageFields.caufContactNumber, vFields)
        .SetItem(ContactAddressUsageFields.caufAddressUsage, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ContactAddressUsageRecordSetTypes.caurtAll) = ContactAddressUsageRecordSetTypes.caurtAll Then
          .SetItem(ContactAddressUsageFields.caufNotes, vFields)
          .SetItem(ContactAddressUsageFields.caufAmendedBy, vFields)
          .SetItem(ContactAddressUsageFields.caufAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pUsage As String, ByRef pNotes As String)
      With mvClassFields
        .Item(ContactAddressUsageFields.caufContactNumber).IntegerValue = pContactNumber
        .Item(ContactAddressUsageFields.caufAddressNumber).IntegerValue = pAddressNumber
        .Item(ContactAddressUsageFields.caufAddressUsage).Value = pUsage
        .Item(ContactAddressUsageFields.caufNotes).Value = pNotes
      End With
    End Sub

    Public Sub Update(ByRef pNotes As String)
      With mvClassFields
        .Item(ContactAddressUsageFields.caufNotes).Value = pNotes
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ContactAddressUsageFields.caufAll)
      mvClassFields.VerifyUnique(mvEnv.Connection)
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
        AddressNumber = mvClassFields.Item(ContactAddressUsageFields.caufAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AddressUsageCode() As String
      Get
        AddressUsageCode = mvClassFields.Item(ContactAddressUsageFields.caufAddressUsage).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(ContactAddressUsageFields.caufAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ContactAddressUsageFields.caufAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ContactAddressUsageFields.caufContactNumber).IntegerValue
      End Get
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(ContactAddressUsageFields.caufNotes).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ContactAddressUsageFields.caufNotes).Value = Value
      End Set
    End Property
  End Class
End Namespace
