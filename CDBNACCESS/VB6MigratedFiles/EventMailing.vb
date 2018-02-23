

Namespace Access
  Public Class EventMailing

    Public Enum EventMailingRecordSetTypes 'These are bit values
      emrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventMailingFields
      emfAll = 0
      emfEventNumber
      emfMailing
      emfAmendedBy
      emfAmendedOn
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
          .DatabaseTableName = "event_mailings"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("mailing")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(EventMailingFields.emfEventNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventMailingFields.emfMailing).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventMailingFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventMailingFields.emfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventMailingFields.emfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventMailingRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventMailingRecordSetTypes.emrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "em")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pMailing As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventNumber > 0 And Len(pMailing) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventMailingRecordSetTypes.emrtAll) & " FROM event_mailings em WHERE event_number = " & pEventNumber & " AND mailing = '" & pMailing & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventMailingRecordSetTypes.emrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventMailingRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventMailingFields.emfEventNumber, vFields)
        .SetItem(EventMailingFields.emfMailing, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventMailingRecordSetTypes.emrtAll) = EventMailingRecordSetTypes.emrtAll Then
          .SetItem(EventMailingFields.emfAmendedBy, vFields)
          .SetItem(EventMailingFields.emfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventMailingFields.emfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(EventMailingFields.emfEventNumber).Value = pParams("EventNumber").Value
        .Item(EventMailingFields.emfMailing).Value = pParams("Mailing").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
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
        AmendedBy = mvClassFields.Item(EventMailingFields.emfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventMailingFields.emfAmendedOn).Value
      End Get
    End Property

    Public Property EventNumber() As Integer
      Get
        EventNumber = CInt(mvClassFields.Item(EventMailingFields.emfEventNumber).Value)
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventMailingFields.emfEventNumber).Value = CStr(Value)
      End Set
    End Property
    Public ReadOnly Property Mailing() As String
      Get
        Mailing = mvClassFields.Item(EventMailingFields.emfMailing).Value
      End Get
    End Property
  End Class
End Namespace
