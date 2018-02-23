Namespace Access
  Public Class PISBankStatement

    Public Enum PisBankStatementRecordSetTypes 'These are bit values
      pbsrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PisBankStatementFields
      pbsfAll = 0
      pbsfPISBankStatementNumber
      pbsfStatementDate
      pbsfDataLoadDate
      pbsfNotes
      pbsfAmendedBy
      pbsfAmendedOn
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
          .DatabaseTableName = "pis_bank_statements"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("pis_bank_statement_number", CDBField.FieldTypes.cftLong)
          .Add("statement_date", CDBField.FieldTypes.cftDate)
          .Add("data_load_date", CDBField.FieldTypes.cftDate)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(PisBankStatementFields.pbsfPISBankStatementNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As PisBankStatementFields)
      'Add code here to ensure all values are valid before saving
      With mvClassFields
        If .Item(PisBankStatementFields.pbsfPISBankStatementNumber).IntegerValue = 0 Then .Item(PisBankStatementFields.pbsfPISBankStatementNumber).IntegerValue = mvEnv.GetControlNumber("PB")
        .Item(PisBankStatementFields.pbsfAmendedOn).Value = TodaysDate()
        .Item(PisBankStatementFields.pbsfAmendedBy).Value = mvEnv.User.Logname
      End With
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PisBankStatementRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PisBankStatementRecordSetTypes.pbsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pbs")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pPisBankStatementNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pPisBankStatementNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PisBankStatementRecordSetTypes.pbsrtAll) & " FROM pis_bank_statements pbs WHERE pis_bank_statement_number = " & pPisBankStatementNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PisBankStatementRecordSetTypes.pbsrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PisBankStatementRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PisBankStatementFields.pbsfPISBankStatementNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PisBankStatementRecordSetTypes.pbsrtAll) = PisBankStatementRecordSetTypes.pbsrtAll Then
          .SetItem(PisBankStatementFields.pbsfStatementDate, vFields)
          .SetItem(PisBankStatementFields.pbsfDataLoadDate, vFields)
          .SetItem(PisBankStatementFields.pbsfNotes, vFields)
          .SetItem(PisBankStatementFields.pbsfAmendedBy, vFields)
          .SetItem(PisBankStatementFields.pbsfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PisBankStatementFields.pbsfAll)
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
        AmendedBy = mvClassFields.Item(PisBankStatementFields.pbsfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PisBankStatementFields.pbsfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property DataLoadDate() As String
      Get
        DataLoadDate = mvClassFields.Item(PisBankStatementFields.pbsfDataLoadDate).Value
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(PisBankStatementFields.pbsfNotes).Value
      End Get
    End Property

    Public Property PisBankStatementNumber() As Integer
      Get
        PisBankStatementNumber = CInt(mvClassFields.Item(PisBankStatementFields.pbsfPISBankStatementNumber).Value)
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(PisBankStatementFields.pbsfPISBankStatementNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property StatementDate() As String
      Get
        StatementDate = mvClassFields.Item(PisBankStatementFields.pbsfStatementDate).Value
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pStatementDate As String, ByVal pNotes As String, ByVal pLoadDate As String)
      Init(pEnv)
      With mvClassFields
        .Item(PisBankStatementFields.pbsfStatementDate).Value = pStatementDate
        If Len(pNotes) > 0 Then .Item(PisBankStatementFields.pbsfNotes).Value = pNotes
        .Item(PisBankStatementFields.pbsfDataLoadDate).Value = pLoadDate
      End With
    End Sub
  End Class
End Namespace
