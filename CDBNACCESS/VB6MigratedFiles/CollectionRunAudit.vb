

Namespace Access
  Public Class CollectionRunAudit

    Public Enum CollectionRunAuditRecordSetTypes 'These are bit values
      crartAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionRunAuditFields
      crafAll = 0
      crafProcessType
      crafLogname
      crafRunDate
      crafBankAccount
      crafFromDate
      crafToDate
      crafBatchNumber
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
          .DatabaseTableName = "collection_run_audit"
          .Add("process_type")
          .Add("logname")
          .Add("run_date", CDBField.FieldTypes.cftDate)
          .Add("bank_account")
          .Add("from_date", CDBField.FieldTypes.cftDate)
          .Add("to_date", CDBField.FieldTypes.cftDate)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CollectionRunAuditFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionRunAuditRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CollectionRunAuditRecordSetTypes.crartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cra")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionRunAuditRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionRunAuditRecordSetTypes.crartAll) = CollectionRunAuditRecordSetTypes.crartAll Then
          .SetItem(CollectionRunAuditFields.crafProcessType, vFields)
          .SetItem(CollectionRunAuditFields.crafLogname, vFields)
          .SetItem(CollectionRunAuditFields.crafRunDate, vFields)
          .SetItem(CollectionRunAuditFields.crafBankAccount, vFields)
          .SetItem(CollectionRunAuditFields.crafFromDate, vFields)
          .SetItem(CollectionRunAuditFields.crafToDate, vFields)
          .SetItem(CollectionRunAuditFields.crafBatchNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(CollectionRunAuditFields.crafAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(CollectionRunAuditFields.crafBankAccount).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafBankAccount).Value = Value
      End Set
    End Property

    Public Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(CollectionRunAuditFields.crafBatchNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CollectionRunAuditFields.crafBatchNumber).IntegerValue = Value
      End Set
    End Property

    Public Property FromDate() As String
      Get
        FromDate = mvClassFields.Item(CollectionRunAuditFields.crafFromDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafFromDate).Value = Value
      End Set
    End Property

    Public Property Logname() As String
      Get
        Logname = mvClassFields.Item(CollectionRunAuditFields.crafLogname).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafLogname).Value = Value
      End Set
    End Property

    Public Property ProcessType() As String
      Get
        ProcessType = mvClassFields.Item(CollectionRunAuditFields.crafProcessType).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafProcessType).Value = Value
      End Set
    End Property

    Public Property RunDate() As String
      Get
        RunDate = mvClassFields.Item(CollectionRunAuditFields.crafRunDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafRunDate).Value = Value
      End Set
    End Property

    Public Property ToDate() As String
      Get
        ToDate = mvClassFields.Item(CollectionRunAuditFields.crafToDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CollectionRunAuditFields.crafToDate).Value = Value
      End Set
    End Property
  End Class
End Namespace
