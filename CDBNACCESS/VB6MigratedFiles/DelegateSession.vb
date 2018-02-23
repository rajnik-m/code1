

Namespace Access
  Public Class DelegateSession

    Public Enum DelegateSessionRecordSetTypes 'These are bit values
      dsrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DelegateSessionFields
      dsfAll = 0
      dsfEventDelegateNumber
      dsfSessionNumber
      dsfAttended
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
          .DatabaseTableName = "delegate_sessions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_delegate_number", CDBField.FieldTypes.cftLong)
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("attended")
        End With

        mvClassFields.Item(DelegateSessionFields.dsfEventDelegateNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DelegateSessionFields.dsfSessionNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As DelegateSessionFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As DelegateSessionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DelegateSessionRecordSetTypes.dsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ds")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventDelegateNumber As Integer = 0, Optional ByRef pSessionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventDelegateNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DelegateSessionRecordSetTypes.dsrtAll) & " FROM delegate_sessions ds WHERE event_delegate_number = " & pEventDelegateNumber & " AND session_number = " & pSessionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DelegateSessionRecordSetTypes.dsrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DelegateSessionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DelegateSessionFields.dsfEventDelegateNumber, vFields)
        .SetItem(DelegateSessionFields.dsfSessionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DelegateSessionRecordSetTypes.dsrtAll) = DelegateSessionRecordSetTypes.dsrtAll Then
          .SetItem(DelegateSessionFields.dsfAttended, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAudit As Boolean = False)
      SetValid(DelegateSessionFields.dsfAll)
      mvClassFields.Save(mvEnv, mvExisting, "", pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property Attended() As Boolean
      Get
        Attended = mvClassFields.Item(DelegateSessionFields.dsfAttended).Bool
      End Get
    End Property

    Public ReadOnly Property EventDelegateNumber() As Integer
      Get
        EventDelegateNumber = CInt(mvClassFields.Item(DelegateSessionFields.dsfEventDelegateNumber).Value)
      End Get
    End Property

    Public ReadOnly Property SessionNumber() As Integer
      Get
        SessionNumber = CInt(mvClassFields.Item(DelegateSessionFields.dsfSessionNumber).Value)
      End Get
    End Property
    Public Sub Create(ByVal pEventDelegateNumber As Integer, ByVal pSessionNumber As Integer, Optional ByVal pAttended As String = "N")
      With mvClassFields
        .Item(DelegateSessionFields.dsfEventDelegateNumber).Value = CStr(pEventDelegateNumber)
        .Item(DelegateSessionFields.dsfSessionNumber).Value = CStr(pSessionNumber)
        .Item(DelegateSessionFields.dsfAttended).Value = pAttended
      End With
    End Sub
    Public Sub Update(ByVal pEventDelegateNumber As Integer, ByVal pSessionNumber As Integer, ByVal pAttended As String)
      With mvClassFields
        .Item(DelegateSessionFields.dsfEventDelegateNumber).Value = CStr(pEventDelegateNumber)
        .Item(DelegateSessionFields.dsfSessionNumber).Value = CStr(pSessionNumber)
        .Item(DelegateSessionFields.dsfAttended).Value = pAttended
      End With
    End Sub
    Public Sub Delete(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub
  End Class
End Namespace
