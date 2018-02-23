

Namespace Access
  Public Class DepartmentNote

    Public Enum DepartmentNoteRecordSetTypes 'These are bit values
      dnrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DepartmentNoteFields
      dnfAll = 0
      dnfRecordType
      dnfUniqueId
      dnfDepartment
      dnfNotes
      dnfAmendedOn
      dnfAmendedBy
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
          .DatabaseTableName = "department_notes"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("record_type")
          .Add("unique_id", CDBField.FieldTypes.cftLong)
          .Add("department")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")

          .Item(DepartmentNoteFields.dnfRecordType).SetPrimaryKeyOnly()
          .Item(DepartmentNoteFields.dnfUniqueId).SetPrimaryKeyOnly()
          .Item(DepartmentNoteFields.dnfDepartment).SetPrimaryKeyOnly()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As DepartmentNoteFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(DepartmentNoteFields.dnfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(DepartmentNoteFields.dnfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As DepartmentNoteRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DepartmentNoteRecordSetTypes.dnrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dn")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pRecordType As String = "", Optional ByRef pUniqueId As Integer = 0, Optional ByRef pDepartment As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pRecordType) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DepartmentNoteRecordSetTypes.dnrtAll) & " FROM department_notes dn WHERE record_type = '" & pRecordType & "' AND unique_id = " & pUniqueId & " AND department = '" & pDepartment & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DepartmentNoteRecordSetTypes.dnrtAll)
        Else
          InitClassFields()
          SetDefaults()
          mvClassFields(DepartmentNoteFields.dnfRecordType).Value = pRecordType
          mvClassFields(DepartmentNoteFields.dnfUniqueId).IntegerValue = pUniqueId
          mvClassFields(DepartmentNoteFields.dnfDepartment).Value = pDepartment
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DepartmentNoteRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DepartmentNoteFields.dnfRecordType, vFields)
        .SetItem(DepartmentNoteFields.dnfUniqueId, vFields)
        .SetItem(DepartmentNoteFields.dnfDepartment, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DepartmentNoteRecordSetTypes.dnrtAll) = DepartmentNoteRecordSetTypes.dnrtAll Then
          .SetItem(DepartmentNoteFields.dnfNotes, vFields)
          .SetItem(DepartmentNoteFields.dnfAmendedOn, vFields)
          .SetItem(DepartmentNoteFields.dnfAmendedBy, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(DepartmentNoteFields.dnfAll)
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
        AmendedBy = mvClassFields.Item(DepartmentNoteFields.dnfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(DepartmentNoteFields.dnfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Department() As String
      Get
        Department = mvClassFields.Item(DepartmentNoteFields.dnfDepartment).Value
      End Get
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(DepartmentNoteFields.dnfNotes).MultiLineValue
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(DepartmentNoteFields.dnfNotes).Value = Value
      End Set
    End Property

    Public ReadOnly Property RecordType() As String
      Get
        RecordType = mvClassFields.Item(DepartmentNoteFields.dnfRecordType).Value
      End Get
    End Property

    Public ReadOnly Property UniqueId() As Integer
      Get
        UniqueId = mvClassFields.Item(DepartmentNoteFields.dnfUniqueId).IntegerValue
      End Get
    End Property
  End Class
End Namespace
