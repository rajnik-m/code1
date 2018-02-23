

Namespace Access
  Public Class EventRoomLink

    Public Enum EventRoomLinkRecordSetTypes 'These are bit values
      erlrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventRoomLinkFields
      erlfAll = 0
      erlfEventNumber
      erlfBlockBookingNumber
      erlfAmendedBy
      erlfAmendedOn
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
          .DatabaseTableName = "event_room_links"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("block_booking_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventRoomLinkFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventRoomLinkFields.erlfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventRoomLinkFields.erlfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventRoomLinkRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventRoomLinkRecordSetTypes.erlrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "erl")
      End If
      GetRecordSetFields = vFields
    End Function
    Public Sub Delete()
      mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBlockBookingNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBlockBookingNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventRoomLinkRecordSetTypes.erlrtAll) & " FROM event_room_links erl WHERE block_booking_number = " & pBlockBookingNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventRoomLinkRecordSetTypes.erlrtAll)
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
    Public Sub SetValuesFromBooking(ByVal pEventNumber As Integer, ByVal pRoomBlockBooking As RoomBlockBooking)
      mvClassFields.Item(EventRoomLinkFields.erlfEventNumber).Value = CStr(pEventNumber)
      mvClassFields.Item(EventRoomLinkFields.erlfBlockBookingNumber).Value = CStr(pRoomBlockBooking.BlockBookingNumber)
    End Sub
    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventRoomLinkRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And EventRoomLinkRecordSetTypes.erlrtAll) = EventRoomLinkRecordSetTypes.erlrtAll Then
          .SetItem(EventRoomLinkFields.erlfEventNumber, vFields)
          .SetItem(EventRoomLinkFields.erlfBlockBookingNumber, vFields)
          .SetItem(EventRoomLinkFields.erlfAmendedBy, vFields)
          .SetItem(EventRoomLinkFields.erlfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventRoomLinkFields.erlfAll)
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
        AmendedBy = mvClassFields.Item(EventRoomLinkFields.erlfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventRoomLinkFields.erlfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BlockBookingNumber() As Integer
      Get
        BlockBookingNumber = mvClassFields.Item(EventRoomLinkFields.erlfBlockBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventRoomLinkFields.erlfEventNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
