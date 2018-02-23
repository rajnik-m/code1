

Namespace Access
  Public Class RoomBookingLink

    Public Enum RoomBookingLinkRecordSetTypes 'These are bit values
      rblrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum RoomBookingLinkFields
      rblfAll = 0
      rblfRoomBookingNumber
      rblfRoomId
      rblfContactNumber
      rblfAddressNumber
      rblfRoomDate
      rblfNotes
      rblfAmendedBy
      rblfAmendedOn
      rblfRoomBookingLinkNumber
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
          .DatabaseTableName = "room_booking_links"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("room_booking_number", CDBField.FieldTypes.cftLong)
          .Add("room_id", CDBField.FieldTypes.cftInteger)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("room_date", CDBField.FieldTypes.cftDate)
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("room_booking_link_number", CDBField.FieldTypes.cftLong)
        End With

        mvClassFields.Item(RoomBookingLinkFields.rblfRoomBookingLinkNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As RoomBookingLinkFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(RoomBookingLinkFields.rblfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(RoomBookingLinkFields.rblfAmendedBy).Value = mvEnv.User.Logname
      If RoomBookingLinkNumber = 0 Then mvClassFields(RoomBookingLinkFields.rblfRoomBookingLinkNumber).IntegerValue = mvEnv.GetControlNumber("RL")
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As RoomBookingLinkRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = RoomBookingLinkRecordSetTypes.rblrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "rbl")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pRoomBookingLinkNumber As Integer = 0, Optional ByRef pRoomBookingNumber As Integer = 0, Optional ByRef pRoomId As Integer = 0, Optional ByRef pRoomDate As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pRoomBookingLinkNumber > 0 Or pRoomBookingNumber > 0 Then
        If pRoomBookingLinkNumber > 0 Then
          vWhereFields.Add("room_booking_link_number", CDBField.FieldTypes.cftLong, pRoomBookingLinkNumber)
        Else
          vWhereFields.Add("room_booking_number", CDBField.FieldTypes.cftLong, pRoomBookingNumber)
          vWhereFields.Add("room_id", CDBField.FieldTypes.cftLong, pRoomId)
          vWhereFields.Add("room_date", CDBField.FieldTypes.cftLong, pRoomDate)
        End If
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(RoomBookingLinkRecordSetTypes.rblrtAll) & " FROM room_booking_links rbl WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, RoomBookingLinkRecordSetTypes.rblrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As RoomBookingLinkRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(RoomBookingLinkFields.rblfRoomBookingLinkNumber, vFields)
        .SetItem(RoomBookingLinkFields.rblfRoomBookingNumber, vFields)
        .SetItem(RoomBookingLinkFields.rblfRoomId, vFields)
        .SetItem(RoomBookingLinkFields.rblfRoomDate, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And RoomBookingLinkRecordSetTypes.rblrtAll) = RoomBookingLinkRecordSetTypes.rblrtAll Then
          .SetItem(RoomBookingLinkFields.rblfContactNumber, vFields)
          .SetItem(RoomBookingLinkFields.rblfAddressNumber, vFields)
          .SetItem(RoomBookingLinkFields.rblfNotes, vFields)
          .SetItem(RoomBookingLinkFields.rblfAmendedBy, vFields)
          .SetItem(RoomBookingLinkFields.rblfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      If pParams.Exists("ContactNumber") Then mvClassFields.Item(RoomBookingLinkFields.rblfContactNumber).Value = pParams("ContactNumber").Value
      If pParams.Exists("AddressNumber") Then mvClassFields.Item(RoomBookingLinkFields.rblfAddressNumber).Value = pParams("AddressNumber").Value
      If pParams.Exists("Notes") Then mvClassFields.Item(RoomBookingLinkFields.rblfNotes).Value = pParams("Notes").Value
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(RoomBookingLinkFields.rblfAll)
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
        Return mvClassFields.Item(RoomBookingLinkFields.rblfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(RoomBookingLinkFields.rblfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(RoomBookingLinkFields.rblfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields.Item(RoomBookingLinkFields.rblfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(RoomBookingLinkFields.rblfNotes).Value
      End Get
    End Property

    Public ReadOnly Property RoomBookingLinkNumber() As Integer
      Get
        Return mvClassFields.Item(RoomBookingLinkFields.rblfRoomBookingLinkNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RoomBookingNumber() As Integer
      Get
        Return mvClassFields.Item(RoomBookingLinkFields.rblfRoomBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RoomDate() As String
      Get
        RoomDate = mvClassFields.Item(RoomBookingLinkFields.rblfRoomDate).Value
      End Get
    End Property

    Public ReadOnly Property RoomId() As Integer
      Get
        RoomId = CInt(mvClassFields.Item(RoomBookingLinkFields.rblfRoomId).Value)
      End Get
    End Property
  End Class
End Namespace
