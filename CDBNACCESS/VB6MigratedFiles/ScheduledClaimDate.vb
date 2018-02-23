

Namespace Access
  Public Class ScheduledClaimDate

    Public Enum ScheduledClaimDateRecordSetTypes 'These are bit values
      scdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ScheduledClaimDateFields
      scdfAll = 0
      scdfBankAccount
      scdfClaimType
      scdfClaimDate
      scdfLatestDueDate
      scdfAmendedBy
      scdfAmendedOn
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
          .DatabaseTableName = "scheduled_claim_dates"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("bank_account")
          .Add("claim_type")
          .Add("claim_date", CDBField.FieldTypes.cftDate)
          .Add("latest_due_date", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(ScheduledClaimDateFields.scdfBankAccount).PrefixRequired = True

        mvClassFields.Item(ScheduledClaimDateFields.scdfBankAccount).SetPrimaryKeyOnly()
        mvClassFields.Item(ScheduledClaimDateFields.scdfClaimType).SetPrimaryKeyOnly()
        mvClassFields.Item(ScheduledClaimDateFields.scdfClaimDate).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ScheduledClaimDateFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ScheduledClaimDateFields.scdfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ScheduledClaimDateFields.scdfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ScheduledClaimDateRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ScheduledClaimDateRecordSetTypes.scdrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "scd")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBankAccount As String = "", Optional ByRef pClaimType As String = "", Optional ByRef pClaimDate As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pBankAccount) > 0 And Len(pClaimType) > 0 And Len(pClaimDate) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ScheduledClaimDateRecordSetTypes.scdrtAll) & " FROM scheduled_claim_dates scd WHERE bank_account = '" & pBankAccount & "' AND claim_type = '" & pClaimType & "' AND claim_date " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftDate, pClaimDate))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ScheduledClaimDateRecordSetTypes.scdrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ScheduledClaimDateRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ScheduledClaimDateFields.scdfBankAccount, vFields)
        .SetItem(ScheduledClaimDateFields.scdfClaimType, vFields)
        .SetItem(ScheduledClaimDateFields.scdfClaimDate, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ScheduledClaimDateRecordSetTypes.scdrtAll) = ScheduledClaimDateRecordSetTypes.scdrtAll Then
          .SetItem(ScheduledClaimDateFields.scdfLatestDueDate, vFields)
          .SetItem(ScheduledClaimDateFields.scdfAmendedBy, vFields)
          .SetItem(ScheduledClaimDateFields.scdfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ScheduledClaimDateFields.scdfAll)
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
        AmendedBy = mvClassFields.Item(ScheduledClaimDateFields.scdfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ScheduledClaimDateFields.scdfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(ScheduledClaimDateFields.scdfBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property ClaimDate() As String
      Get
        ClaimDate = mvClassFields.Item(ScheduledClaimDateFields.scdfClaimDate).Value
      End Get
    End Property

    Public ReadOnly Property ClaimType() As String
      Get
        ClaimType = mvClassFields.Item(ScheduledClaimDateFields.scdfClaimType).Value
      End Get
    End Property

    Public ReadOnly Property LatestDueDate() As String
      Get
        LatestDueDate = mvClassFields.Item(ScheduledClaimDateFields.scdfLatestDueDate).Value
      End Get
    End Property
  End Class
End Namespace
