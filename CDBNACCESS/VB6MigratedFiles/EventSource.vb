

Namespace Access
  Public Class EventSource

    Public Enum EventSourceRecordSetTypes 'These are bit values
      esrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventSourceFields
      esfAll = 0
      esfEventNumber
      esfSource
      esfAmendedBy
      esfAmendedOn
    End Enum

    Private mvEvent As CDBEvent

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
          .DatabaseTableName = "event_sources"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("source")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(EventSourceFields.esfEventNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventSourceFields.esfSource).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventSourceFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventSourceFields.esfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventSourceFields.esfAmendedBy).Value = mvEnv.User.Logname
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
        AmendedBy = mvClassFields.Item(EventSourceFields.esfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventSourceFields.esfAmendedOn).Value
      End Get
    End Property

    Private ReadOnly Property EventHeader() As CDBEvent
      Get
        If mvEvent Is Nothing Then
          mvEvent = New CDBEvent(mvEnv)
          mvEvent.Init((mvClassFields(EventSourceFields.esfEventNumber).IntegerValue))
        End If
        EventHeader = mvEvent
      End Get
    End Property

    Public Property EventNumber() As Integer
      Get
        EventNumber = CInt(mvClassFields.Item(EventSourceFields.esfEventNumber).Value)
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventSourceFields.esfEventNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property Source() As String
      Get
        Source = mvClassFields.Item(EventSourceFields.esfSource).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventSourceFields.esfSource).Value = Value
      End Set
    End Property

    Public Function GetRecordSetFields(ByVal pRSType As EventSourceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventSourceRecordSetTypes.esrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "es")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pSource As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventSourceRecordSetTypes.esrtAll) & " FROM event_sources es WHERE event_number = " & pEventNumber & " AND source = '" & pSource & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventSourceRecordSetTypes.esrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventSourceRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventSourceFields.esfEventNumber, vFields)
        .SetItem(EventSourceFields.esfSource, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventSourceRecordSetTypes.esrtAll) = EventSourceRecordSetTypes.esrtAll Then
          .SetItem(EventSourceFields.esfAmendedBy, vFields)
          .SetItem(EventSourceFields.esfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventSourceFields.esfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      If EventHeader.Source = mvClassFields(EventSourceFields.esfSource).Value Then
        RaiseError(DataAccessErrors.daeCannotDeleteEventSource)
      End If
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(EventSourceFields.esfEventNumber).Value = pParams("EventNumber").Value
        .Item(EventSourceFields.esfSource).Value = pParams("Source").Value
      End With
    End Sub
  End Class
End Namespace
