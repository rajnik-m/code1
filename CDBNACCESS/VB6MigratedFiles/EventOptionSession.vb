

Namespace Access
  Public Class EventOptionSession

    Public Enum EventOptionSessionRecordSetTypes 'These are bit values
      osrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventOptionSessionFields
      osfAll = 0
      osfOptionNumber
      osfSessionNumber
      osfAllocation
      osfAmendedBy
      osfAmendedOn
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
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "option_sessions"
          .Add("option_number", CDBField.FieldTypes.cftLong)
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("allocation", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(EventOptionSessionFields.osfOptionNumber).SetPrimaryKeyOnly()
          .Item(EventOptionSessionFields.osfSessionNumber).SetPrimaryKeyOnly()

          .Item(EventOptionSessionFields.osfOptionNumber).PrefixRequired = True
          .Item(EventOptionSessionFields.osfSessionNumber).PrefixRequired = True
          .Item(EventOptionSessionFields.osfAmendedBy).PrefixRequired = True
          .Item(EventOptionSessionFields.osfAmendedOn).PrefixRequired = True
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventOptionSessionFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventOptionSessionFields.osfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventOptionSessionFields.osfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("Allocation") Then .Item(EventOptionSessionFields.osfAllocation).Value = pParams("Allocation").Value
      End With
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventOptionSessionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventOptionSessionRecordSetTypes.osrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "os")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pOptionNumber As Integer = 0, Optional ByRef pSessionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pOptionNumber > 0 And pSessionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventOptionSessionRecordSetTypes.osrtAll) & " FROM option_sessions os WHERE option_number=" & pOptionNumber & " AND session_number = " & pSessionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventOptionSessionRecordSetTypes.osrtAll)
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

    Friend Sub InitFromOptionSession(ByVal pOriginalEvent As CDBEvent, ByRef pOptionSession As EventOptionSession, ByRef pNewEvent As CDBEvent)
      With pOptionSession
        mvClassFields.Item(EventOptionSessionFields.osfSessionNumber).Value = CStr(pNewEvent.Sessions(pOriginalEvent.Sessions.IndexOf(pOriginalEvent.Sessions.Item(.SessionNumber.ToString))).SessionNumber)       'CStr(.SessionNumber)  'CStr(pNewEvent.BaseItemNumber + (.SessionNumber Mod 10000))
        mvClassFields.Item(EventOptionSessionFields.osfOptionNumber).Value = CStr(pNewEvent.BookingOptions(pOriginalEvent.BookingOptions.IndexOf(pOriginalEvent.BookingOptions.Item(.OptionNumber.ToString))).OptionNumber)    'CStr(pNewEvent.BaseItemNumber + (.OptionNumber Mod 10000))
        mvClassFields.Item(EventOptionSessionFields.osfAllocation).Value = CStr(.Allocation)
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventOptionSessionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And EventOptionSessionRecordSetTypes.osrtAll) = EventOptionSessionRecordSetTypes.osrtAll Then
          .SetItem(EventOptionSessionFields.osfOptionNumber, vFields)
          .SetItem(EventOptionSessionFields.osfSessionNumber, vFields)
          .SetItem(EventOptionSessionFields.osfAllocation, vFields)
          .SetItem(EventOptionSessionFields.osfAmendedBy, vFields)
          .SetItem(EventOptionSessionFields.osfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventOptionSessionFields.osfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByRef pOptionNumber As Integer, ByRef pSessionNumber As Integer, ByRef pAllocation As Double)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvClassFields.Item(EventOptionSessionFields.osfOptionNumber).IntegerValue = pOptionNumber
      mvClassFields.Item(EventOptionSessionFields.osfSessionNumber).IntegerValue = pSessionNumber
      mvClassFields.Item(EventOptionSessionFields.osfAllocation).DoubleValue = pAllocation
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property Allocation() As Double
      Get
        Allocation = mvClassFields.Item(EventOptionSessionFields.osfAllocation).DoubleValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventOptionSessionFields.osfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventOptionSessionFields.osfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property OptionNumber() As Integer
      Get
        OptionNumber = mvClassFields.Item(EventOptionSessionFields.osfOptionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SessionNumber() As Integer
      Get
        SessionNumber = mvClassFields.Item(EventOptionSessionFields.osfSessionNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
