

Namespace Access
  Public Class CommunicationUsage

    Public Enum CommunicationUsageRecordSetTypes 'These are bit values
      curtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CommunicationUsageFields
      cufAll = 0
      cufCommunicationUsage
      cufCommunicationUsageDesc
      cufAmendedBy
      cufAmendedOn
      cufNotesMandatory
      cufSingleUsage
      cufDevice
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
          .DatabaseTableName = "communication_usages"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("communication_usage")
          .Add("communication_usage_desc")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("notes_mandatory")
          .Add("single_usage")
          .Add("device")
        End With

        mvClassFields.Item(CommunicationUsageFields.cufCommunicationUsage).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CommunicationUsageFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CommunicationUsageFields.cufAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CommunicationUsageFields.cufAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CommunicationUsageRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CommunicationUsageRecordSetTypes.curtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cu")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCommunicationUsage As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCommunicationUsage) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CommunicationUsageRecordSetTypes.curtAll) & " FROM communication_usages cu WHERE communication_usage = '" & pCommunicationUsage & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CommunicationUsageRecordSetTypes.curtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CommunicationUsageRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CommunicationUsageFields.cufCommunicationUsage, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CommunicationUsageRecordSetTypes.curtAll) = CommunicationUsageRecordSetTypes.curtAll Then
          .SetItem(CommunicationUsageFields.cufCommunicationUsageDesc, vFields)
          .SetItem(CommunicationUsageFields.cufAmendedBy, vFields)
          .SetItem(CommunicationUsageFields.cufAmendedOn, vFields)
          .SetItem(CommunicationUsageFields.cufNotesMandatory, vFields)
          .SetItem(CommunicationUsageFields.cufSingleUsage, vFields)
          .SetItem(CommunicationUsageFields.cufDevice, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CommunicationUsageFields.cufAll)
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
        AmendedBy = mvClassFields.Item(CommunicationUsageFields.cufAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CommunicationUsageFields.cufAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CommunicationUsageCode() As String
      Get
        CommunicationUsageCode = mvClassFields.Item(CommunicationUsageFields.cufCommunicationUsage).Value
      End Get
    End Property

    Public ReadOnly Property CommunicationUsageDesc() As String
      Get
        CommunicationUsageDesc = mvClassFields.Item(CommunicationUsageFields.cufCommunicationUsageDesc).Value
      End Get
    End Property

    Public ReadOnly Property NotesMandatory() As Boolean
      Get
        NotesMandatory = mvClassFields.Item(CommunicationUsageFields.cufNotesMandatory).Bool
      End Get
    End Property

    Public ReadOnly Property SingleUsage() As Boolean
      Get
        Return mvClassFields.Item(CommunicationUsageFields.cufSingleUsage).Bool
      End Get
    End Property

    Public ReadOnly Property Device() As String
      Get
        Return mvClassFields.Item(CommunicationUsageFields.cufDevice).Value
      End Get
    End Property
  End Class
End Namespace
