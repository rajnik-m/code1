Namespace Access
  Public Class AddressUsage

    Public Enum AddressUsageRecordSetTypes 'These are bit values
      aurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum AddressUsageFields
      aufAll = 0
      aufAddressUsage
      aufAddressUsageDesc
      aufAmendedBy
      aufAmendedOn
      aufNotesMandatory
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "address_usages"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("address_usage")
          .Add("address_usage_desc")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("notes_mandatory")
        End With

        mvClassFields.Item(AddressUsageFields.aufAddressUsage).SetPrimaryKeyOnly()
        mvClassFields.Item(AddressUsageFields.aufNotesMandatory).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAUNotesMandatory)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As AddressUsageFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(AddressUsageFields.aufAmendedOn).Value = TodaysDate()
      mvClassFields.Item(AddressUsageFields.aufAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As AddressUsageRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AddressUsageRecordSetTypes.aurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "au")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAddressUsage As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pAddressUsage) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AddressUsageRecordSetTypes.aurtAll) & " FROM address_usages au WHERE address_usage = '" & pAddressUsage & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, AddressUsageRecordSetTypes.aurtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AddressUsageRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AddressUsageFields.aufAddressUsage, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AddressUsageRecordSetTypes.aurtAll) = AddressUsageRecordSetTypes.aurtAll Then
          .SetItem(AddressUsageFields.aufAddressUsageDesc, vFields)
          .SetItem(AddressUsageFields.aufAmendedBy, vFields)
          .SetItem(AddressUsageFields.aufAmendedOn, vFields)
          .SetOptionalItem(AddressUsageFields.aufNotesMandatory, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(AddressUsageFields.aufAll)
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

    Public ReadOnly Property AddressUsageCode() As String
      Get
        AddressUsageCode = mvClassFields.Item(AddressUsageFields.aufAddressUsage).Value
      End Get
    End Property

    Public ReadOnly Property AddressUsageDesc() As String
      Get
        AddressUsageDesc = mvClassFields.Item(AddressUsageFields.aufAddressUsageDesc).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(AddressUsageFields.aufAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(AddressUsageFields.aufAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property NotesMandatory() As Boolean
      Get
        NotesMandatory = mvClassFields.Item(AddressUsageFields.aufNotesMandatory).Bool
      End Get
    End Property
  End Class
End Namespace
