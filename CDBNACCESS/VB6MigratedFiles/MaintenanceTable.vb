

Namespace Access
  Public Class MaintenanceTable

    Public Enum MaintenanceTableRecordSetTypes 'These are bit values
      matrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum MaintenanceTableFields
      mtfAll = 0
      mtfTableName
      mtfTableNameDesc
      mtfSelectionAttribute
      mtfApplicationName
      mtfTableAmendedBy
      mtfTableAmendedOn
      mtfMaintenanceGroup
      mtfTableDesc
      mtfTableNotes
      mtfDefaultNotes
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvAdministratorNotes As String
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "maintenance_tables"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("table_name")
          .Add("table_name_desc")
          .Add("selection_attribute")
          .Add("application_name")
          .Add("table_amended_by")
          .Add("table_amended_on", CDBField.FieldTypes.cftTime)
          .Add("maintenance_group")
          .Add("table_desc")
          .Add("table_notes")
          .Add("default_notes")
        End With
        mvClassFields.Item(MaintenanceTableFields.mtfTableName).SetPrimaryKeyOnly()
        mvClassFields.Item(MaintenanceTableFields.mtfTableName).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As MaintenanceTableFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As MaintenanceTableRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = MaintenanceTableRecordSetTypes.matrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "mt")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pTableName As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pTableName) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(MaintenanceTableRecordSetTypes.matrtAll) & " FROM maintenance_tables mt WHERE table_name = '" & pTableName & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, MaintenanceTableRecordSetTypes.matrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As MaintenanceTableRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(MaintenanceTableFields.mtfTableName, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And MaintenanceTableRecordSetTypes.matrtAll) = MaintenanceTableRecordSetTypes.matrtAll Then
          .SetItem(MaintenanceTableFields.mtfTableNameDesc, vFields)
          .SetItem(MaintenanceTableFields.mtfSelectionAttribute, vFields)
          .SetItem(MaintenanceTableFields.mtfApplicationName, vFields)
          .SetItem(MaintenanceTableFields.mtfTableAmendedBy, vFields)
          .SetItem(MaintenanceTableFields.mtfTableAmendedOn, vFields)
          .SetItem(MaintenanceTableFields.mtfMaintenanceGroup, vFields)
          .SetOptionalItem(MaintenanceTableFields.mtfTableDesc, vFields)
          .SetOptionalItem(MaintenanceTableFields.mtfTableNotes, vFields)
          .SetOptionalItem(MaintenanceTableFields.mtfDefaultNotes, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(MaintenanceTableFields.mtfAll)
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

    Public ReadOnly Property ApplicationName() As String
      Get
        ApplicationName = mvClassFields.Item(MaintenanceTableFields.mtfApplicationName).Value
      End Get
    End Property

    Public ReadOnly Property DefaultNotes() As String
      Get
        DefaultNotes = mvClassFields.Item(MaintenanceTableFields.mtfDefaultNotes).Value
      End Get
    End Property

    Public ReadOnly Property MaintenanceGroup() As String
      Get
        MaintenanceGroup = mvClassFields.Item(MaintenanceTableFields.mtfMaintenanceGroup).Value
      End Get
    End Property

    Public ReadOnly Property SelectionAttribute() As String
      Get
        SelectionAttribute = mvClassFields.Item(MaintenanceTableFields.mtfSelectionAttribute).Value
      End Get
    End Property

    Public ReadOnly Property TableAmendedBy() As String
      Get
        TableAmendedBy = mvClassFields.Item(MaintenanceTableFields.mtfTableAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property TableAmendedOn() As String
      Get
        TableAmendedOn = mvClassFields.Item(MaintenanceTableFields.mtfTableAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property TableDesc() As String
      Get
        TableDesc = mvClassFields.Item(MaintenanceTableFields.mtfTableDesc).Value
      End Get
    End Property

    Public ReadOnly Property TableName() As String
      Get
        TableName = mvClassFields.Item(MaintenanceTableFields.mtfTableName).Value
      End Get
    End Property

    Public ReadOnly Property TableNameDesc() As String
      Get
        TableNameDesc = mvClassFields.Item(MaintenanceTableFields.mtfTableNameDesc).Value
      End Get
    End Property

    Public ReadOnly Property TableNotes() As String
      Get
        TableNotes = mvClassFields.Item(MaintenanceTableFields.mtfTableNotes).MultiLineValue
      End Get
    End Property

    Public Property AdministratorNotes() As String
      Get
        Dim vRS As CDBRecordSet

        If Len(mvAdministratorNotes) = 0 Then
          vRS = mvEnv.Connection.GetRecordSet("SELECT administrator_notes FROM table_notes WHERE table_name = '" & TableName & "'")
          If vRS.Fetch() = True Then
            mvAdministratorNotes = vRS.Fields("administrator_notes").MultiLine
          End If
          vRS.CloseRecordSet()
        End If
        AdministratorNotes = mvAdministratorNotes
      End Get
      Set(ByVal Value As String)
        Dim vUpdateFields As New CDBFields
        Dim vWhereFields As New CDBFields
        Dim vUpdate As Boolean

        If Value <> mvAdministratorNotes Then
          vUpdateFields.Add("administrator_notes", CDBField.FieldTypes.cftCharacter, Value)
          vWhereFields.Add("table_name", CDBField.FieldTypes.cftCharacter, TableName)
          If mvAdministratorNotes.Length > 0 Then
            vUpdate = True
          Else
            vUpdate = (mvEnv.Connection.GetCount("table_notes", vWhereFields) > 0)
          End If

          If vUpdate Then
            mvEnv.Connection.UpdateRecords("table_notes", vUpdateFields, vWhereFields)
          Else
            With vUpdateFields
              .Add("table_name", CDBField.FieldTypes.cftCharacter, TableName)
              .AddAmendedOnBy(mvEnv.User.Logname)
            End With
            mvEnv.Connection.InsertRecord("table_notes", vUpdateFields)
          End If
          mvAdministratorNotes = Value
        End If
      End Set
    End Property
  End Class
End Namespace
