

Namespace Access
  Public Class EventSessionActivity

    Public Enum EventSessionActivityRecordSetTypes 'These are bit values
      esartAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventSessionActivityFields
      esafAll = 0
      esafEventNumber
      esafSessionNumber
      esafActivity
      esafActivityValue
      esafAmendedBy
      esafAmendedOn
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
          .DatabaseTableName = "session_activities"
          .Add("event_number", CDBField.FieldTypes.cftInteger)
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("activity")
          .Add("activity_value")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields(EventSessionActivityFields.esafEventNumber).SetPrimaryKeyOnly()
        mvClassFields(EventSessionActivityFields.esafSessionNumber).SetPrimaryKeyOnly()
        mvClassFields(EventSessionActivityFields.esafActivity).SetPrimaryKeyOnly()
        mvClassFields(EventSessionActivityFields.esafActivityValue).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventSessionActivityFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventSessionActivityFields.esafAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventSessionActivityFields.esafAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventSessionActivityRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventSessionActivityRecordSetTypes.esartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "esa")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSessionNumber As Integer = 0, Optional ByRef pActivity As String = "", Optional ByRef pActivityValue As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pSessionNumber > 0 Then
        vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong, pSessionNumber)
        vWhereFields.Add("activity", CDBField.FieldTypes.cftCharacter, pActivity)
        vWhereFields.Add("activity_value", CDBField.FieldTypes.cftCharacter, pActivityValue)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventSessionActivityRecordSetTypes.esartAll) & " FROM " & mvClassFields.DatabaseTableName & " WHERE " & pEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventSessionActivityRecordSetTypes.esartAll)
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

    Public Sub Create(ByRef pSession As EventSession, ByRef pActivity As String, ByRef pActivityValue As String)
      mvClassFields.Item(EventSessionActivityFields.esafEventNumber).Value = CStr(pSession.EventNumber)
      mvClassFields.Item(EventSessionActivityFields.esafSessionNumber).Value = CStr(pSession.SessionNumber)
      mvClassFields.Item(EventSessionActivityFields.esafActivity).Value = pActivity
      mvClassFields.Item(EventSessionActivityFields.esafActivityValue).Value = pActivityValue
    End Sub

    Friend Sub InitFromSessionActivity(ByVal pOriginalEvent As CDBEvent, ByRef pSessionActitiy As EventSessionActivity, ByRef pNewEvent As CDBEvent)
      With pSessionActitiy
        mvClassFields.Item(EventSessionActivityFields.esafEventNumber).Value = CStr(pNewEvent.EventNumber)
        mvClassFields.Item(EventSessionActivityFields.esafSessionNumber).Value = CStr(pNewEvent.Sessions(pOriginalEvent.Sessions.IndexOf(pOriginalEvent.Sessions.Item(.SessionNumber.ToString))).SessionNumber) 'CStr(.SessionNumber) 'CStr(pNewEvent.BaseItemNumber + (.SessionNumber Mod 10000))
        mvClassFields.Item(EventSessionActivityFields.esafActivity).Value = .Activity
        mvClassFields.Item(EventSessionActivityFields.esafActivityValue).Value = .ActivityValue
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventSessionActivityRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And EventSessionActivityRecordSetTypes.esartAll) = EventSessionActivityRecordSetTypes.esartAll Then
          .SetItem(EventSessionActivityFields.esafEventNumber, vFields)
          .SetItem(EventSessionActivityFields.esafSessionNumber, vFields)
          .SetItem(EventSessionActivityFields.esafActivity, vFields)
          .SetItem(EventSessionActivityFields.esafActivityValue, vFields)
          .SetItem(EventSessionActivityFields.esafAmendedBy, vFields)
          .SetItem(EventSessionActivityFields.esafAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventSessionActivityFields.esafAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property Activity() As String
      Get
        Activity = mvClassFields.Item(EventSessionActivityFields.esafActivity).Value
      End Get
    End Property

    Public ReadOnly Property ActivityValue() As String
      Get
        ActivityValue = mvClassFields.Item(EventSessionActivityFields.esafActivityValue).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventSessionActivityFields.esafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventSessionActivityFields.esafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventSessionActivityFields.esafEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SessionNumber() As Integer
      Get
        SessionNumber = mvClassFields.Item(EventSessionActivityFields.esafSessionNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
