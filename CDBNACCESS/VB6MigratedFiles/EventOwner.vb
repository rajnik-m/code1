

Namespace Access
  Public Class EventOwner

    Public Enum EventOwnerRecordSetTypes 'These are bit values
      eowrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventOwnerFields
      eofAll = 0
      eofEventNumber
      eofDepartment
      eofAmendedBy
      eofAmendedOn
    End Enum

    Private mvDepartmentDescription As String

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
          .DatabaseTableName = "event_owners"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("department")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields.Item(EventOwnerFields.eofEventNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventOwnerFields.eofDepartment).SetPrimaryKeyOnly()

        mvClassFields.Item(EventOwnerFields.eofEventNumber).PrefixRequired = True
        mvClassFields.Item(EventOwnerFields.eofDepartment).PrefixRequired = True
        mvClassFields.Item(EventOwnerFields.eofAmendedBy).PrefixRequired = True
        mvClassFields.Item(EventOwnerFields.eofAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventOwnerFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventOwnerFields.eofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventOwnerFields.eofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventOwnerRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventOwnerRecordSetTypes.eowrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "eo")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pDepartment As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventNumber > 0 And Len(pDepartment) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventOwnerRecordSetTypes.eowrtAll) & " FROM event_owners eo WHERE event_number = " & pEventNumber & " AND department = '" & pDepartment & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventOwnerRecordSetTypes.eowrtAll)
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
    Friend Sub InitFromOwner(ByRef pEventOwner As EventOwner, ByRef pNewEvent As CDBEvent)
      With pEventOwner
        mvClassFields.Item(EventOwnerFields.eofEventNumber).Value = CStr(pNewEvent.EventNumber)
        mvClassFields.Item(EventOwnerFields.eofDepartment).Value = .Department
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventOwnerRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventOwnerFields.eofEventNumber, vFields)
        .SetItem(EventOwnerFields.eofDepartment, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventOwnerRecordSetTypes.eowrtAll) = EventOwnerRecordSetTypes.eowrtAll Then
          .SetItem(EventOwnerFields.eofAmendedBy, vFields)
          .SetItem(EventOwnerFields.eofAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByRef pEventNumber As Integer, ByRef pDepartment As String)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields(EventOwnerFields.eofEventNumber).IntegerValue = pEventNumber
      mvClassFields(EventOwnerFields.eofDepartment).Value = pDepartment
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventOwnerFields.eofAll)
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

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventOwnerFields.eofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventOwnerFields.eofAmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property DepartmentDescription() As String
      Get
        If Len(mvDepartmentDescription) = 0 Then
          mvDepartmentDescription = mvEnv.GetDescription("departments", "department", Department)
        End If
        DepartmentDescription = mvDepartmentDescription
      End Get
    End Property
    Public Property Department() As String
      Get
        Department = mvClassFields.Item(EventOwnerFields.eofDepartment).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(EventOwnerFields.eofDepartment).Value = Value
      End Set
    End Property

    Public Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventOwnerFields.eofEventNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventOwnerFields.eofEventNumber).IntegerValue = Value
      End Set
    End Property
  End Class
End Namespace
