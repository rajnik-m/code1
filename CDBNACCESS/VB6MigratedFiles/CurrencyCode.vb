

Namespace Access
  Public Class CurrencyCode

    Public Enum CurrencyCodeRecordSetTypes 'These are bit values
      ccrtAll = &HFFFFS
      'ADD additional recordset types here
      ccrtExchangeRate = 1
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CurrencyCodeFields
      ccfAll = 0
      ccfCurrencyCode
      ccfCurrencyCodeDesc
      ccfExchangeRate
      ccfAmendedBy
      ccfAmendedOn
      ccfBankAccount
      ccfBankAccountDesc
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
          .DatabaseTableName = "currency_codes"
          .Add("currency_code")
          .Add("currency_code_desc")
          .Add("exchange_rate")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("bank_account")
          .Add("bank_account_desc")
        End With

        mvClassFields.Item(CurrencyCodeFields.ccfCurrencyCode).SetPrimaryKeyOnly()

        mvClassFields.Item(CurrencyCodeFields.ccfCurrencyCode).PrefixRequired = True
        mvClassFields.Item(CurrencyCodeFields.ccfAmendedBy).PrefixRequired = True
        mvClassFields.Item(CurrencyCodeFields.ccfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CurrencyCodeFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CurrencyCodeFields.ccfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CurrencyCodeFields.ccfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CurrencyCodeRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CurrencyCodeRecordSetTypes.ccrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        For Each vClassField As ClassField In mvClassFields
          If Left(vClassField.Name, 12) <> "bank_account" Then
            If vClassField.PrefixRequired Then vFields = vFields & "cc."
            vFields = vFields & vClassField.Name & ","
          End If
        Next vClassField
      ElseIf pRSType = CurrencyCodeRecordSetTypes.ccrtExchangeRate Then
        vFields = "exchange_rate"
      End If
      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pRSType As CurrencyCodeRecordSetTypes = CurrencyCodeRecordSetTypes.ccrtAll, Optional ByVal pCurrencyCode As String = "", Optional ByVal pApplication As String = "", Optional ByVal pBatchType As String = "", Optional ByVal pDate As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If Len(pCurrencyCode) > 0 Then
        If Len(pDate) = 0 Then pDate = TodaysDate()
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(pRSType) & " FROM currency_codes cc, currency_rates cr WHERE cc.currency_code = '" & pCurrencyCode & "' AND cc.currency_code = cr.currency_code AND date_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pDate) & " AND date_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pDate) & " ORDER BY date_to DESC")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, pRSType)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
        If Len(pCurrencyCode) > 0 And Len(pApplication) > 0 And Len(pBatchType) > 0 Then
          vSQL = "SELECT faba.bank_account, ba.bank_account_desc FROM fp_application_bank_accounts faba, bank_accounts ba"
          vSQL = vSQL & " WHERE fp_application = '" & pApplication & "'"
          vSQL = vSQL & " AND batch_type = '" & pBatchType & "'"
          vSQL = vSQL & " AND currency_code = '" & pCurrencyCode & "'"
          vSQL = vSQL & " AND faba.bank_account = ba.bank_account"
          vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
          With vRecordSet
            If .Fetch() = True Then
              mvClassFields.SetItem(CurrencyCodeFields.ccfBankAccount, .Fields)
              mvClassFields.SetItem(CurrencyCodeFields.ccfBankAccountDesc, .Fields)
            End If
            .CloseRecordSet()
          End With
        End If
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CurrencyCodeRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CurrencyCodeRecordSetTypes.ccrtAll) = CurrencyCodeRecordSetTypes.ccrtAll Then
          .SetItem(CurrencyCodeFields.ccfCurrencyCode, vFields)
          .SetItem(CurrencyCodeFields.ccfCurrencyCodeDesc, vFields)
          .SetItem(CurrencyCodeFields.ccfExchangeRate, vFields)
          .SetItem(CurrencyCodeFields.ccfAmendedBy, vFields)
          .SetItem(CurrencyCodeFields.ccfAmendedOn, vFields)
        ElseIf (pRSType And CurrencyCodeRecordSetTypes.ccrtExchangeRate) = CurrencyCodeRecordSetTypes.ccrtExchangeRate Then
          .SetItem(CurrencyCodeFields.ccfExchangeRate, vFields)
        End If
      End With
    End Sub
    Public Sub Save()
      SetValid(CurrencyCodeFields.ccfAll)
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

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CurrencyCodeFields.ccfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CurrencyCodeFields.ccfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CurrencyCodeCode() As String
      Get
        CurrencyCodeCode = mvClassFields.Item(CurrencyCodeFields.ccfCurrencyCode).Value
      End Get
    End Property

    Public ReadOnly Property CurrencyCodeDesc() As String
      Get
        CurrencyCodeDesc = mvClassFields.Item(CurrencyCodeFields.ccfCurrencyCodeDesc).Value
      End Get
    End Property

    Public ReadOnly Property ExchangeRate() As Double
      Get
        ExchangeRate = mvClassFields.Item(CurrencyCodeFields.ccfExchangeRate).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BankAccount() As String
      Get
        BankAccount = mvClassFields.Item(CurrencyCodeFields.ccfBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property BankAccountDesc() As String
      Get
        BankAccountDesc = mvClassFields.Item(CurrencyCodeFields.ccfBankAccountDesc).Value
      End Get
    End Property
  End Class
End Namespace
