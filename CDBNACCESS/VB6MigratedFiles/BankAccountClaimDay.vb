Namespace Access
  Public Class BankAccountClaimDay

    Public Enum BankAccountClaimDayRecordSetTypes 'These are bit values
      bacdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BankAccountClaimDayFields
      bacdfAll = 0
      bacdfBankAccount
      bacdfClaimType
      bacdfClaimDay
      bacdfNonWorkingDayBehaviour
      bacdfAmendedBy
      bacdfAmendedOn
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
          .DatabaseTableName = "bank_account_claim_days"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("bank_account")
          .Add("claim_type")
          .Add("claim_day", CDBField.FieldTypes.cftInteger)
          .Add("non_working_day_behaviour")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(BankAccountClaimDayFields.bacdfBankAccount).PrefixRequired = True

        mvClassFields.Item(BankAccountClaimDayFields.bacdfBankAccount).SetPrimaryKeyOnly()
        mvClassFields.Item(BankAccountClaimDayFields.bacdfClaimType).SetPrimaryKeyOnly()
        mvClassFields.Item(BankAccountClaimDayFields.bacdfClaimDay).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As BankAccountClaimDayFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(BankAccountClaimDayFields.bacdfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(BankAccountClaimDayFields.bacdfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BankAccountClaimDayRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BankAccountClaimDayRecordSetTypes.bacdrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bacd")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBankAccount As String = "", Optional ByRef pClaimType As String = "", Optional ByRef pClaimDay As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pBankAccount) > 0 And Len(pClaimType) > 0 And pClaimDay > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(BankAccountClaimDayRecordSetTypes.bacdrtAll) & " FROM bank_account_claim_days bacd WHERE bank_account = '" & pBankAccount & "' AND claim_type = '" & pClaimType & "' AND claim_day = " & pClaimDay)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BankAccountClaimDayRecordSetTypes.bacdrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BankAccountClaimDayRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BankAccountClaimDayFields.bacdfBankAccount, vFields)
        .SetItem(BankAccountClaimDayFields.bacdfClaimType, vFields)
        .SetItem(BankAccountClaimDayFields.bacdfClaimDay, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BankAccountClaimDayRecordSetTypes.bacdrtAll) = BankAccountClaimDayRecordSetTypes.bacdrtAll Then
          .SetItem(BankAccountClaimDayFields.bacdfNonWorkingDayBehaviour, vFields)
          .SetItem(BankAccountClaimDayFields.bacdfAmendedBy, vFields)
          .SetItem(BankAccountClaimDayFields.bacdfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(BankAccountClaimDayFields.bacdfAll)
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
        AmendedBy = mvClassFields.Item(BankAccountClaimDayFields.bacdfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(BankAccountClaimDayFields.bacdfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(BankAccountClaimDayFields.bacdfBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property ClaimDay() As Integer
      Get
        ClaimDay = mvClassFields.Item(BankAccountClaimDayFields.bacdfClaimDay).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ClaimType() As String
      Get
        ClaimType = mvClassFields.Item(BankAccountClaimDayFields.bacdfClaimType).Value
      End Get
    End Property

    Public ReadOnly Property NonWorkingDayBehaviour() As String
      Get
        Return mvClassFields.Item(BankAccountClaimDayFields.bacdfNonWorkingDayBehaviour).Value
      End Get
    End Property
  End Class
End Namespace
