

Namespace Access
  Public Class CreditCustomer

    Public Enum CreditCustomerRecordSetTypes 'These are bit values
      ccurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CreditCustomerFields
      ccfAll = 0
      ccfContactNumber
      ccfAddressNumber
      ccfCompany
      ccfSalesLedgerAccount
      ccfCreditCategory
      ccfStopCode
      ccfCreditLimit
      ccfOutstanding
      ccfOnOrder
      ccfCustomerType
      ccfTermsNumber
      ccfTermsPeriod
      ccfTermsFrom
      ccfLastStatementDate
      ccfLastStatementClosingBalance
      ccfLastStatementNumber
      ccfStatementPeriod
      ccfAmendedBy
      ccfAmendedOn
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
          .DatabaseTableName = "credit_customers"
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("company")
          .Add("sales_ledger_account")
          .Add("credit_category")
          .Add("stop_code")
          .Add("credit_limit", CDBField.FieldTypes.cftNumeric)
          .Add("outstanding", CDBField.FieldTypes.cftNumeric)
          .Add("on_order", CDBField.FieldTypes.cftNumeric)
          .Add("customer_type")
          .Add("terms_number", CDBField.FieldTypes.cftInteger)
          .Add("terms_period")
          .Add("terms_from")
          .Add("last_statement_date", CDBField.FieldTypes.cftDate)
          .Add("last_statement_closing_balance", CDBField.FieldTypes.cftNumeric)
          .Add("last_statement_number", CDBField.FieldTypes.cftInteger)
          .Add("statement_period", CDBField.FieldTypes.cftInteger)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CreditCustomerFields.ccfContactNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(CreditCustomerFields.ccfCompany).SetPrimaryKeyOnly()
        mvClassFields.Item(CreditCustomerFields.ccfSalesLedgerAccount).SetPrimaryKeyOnly()

        mvClassFields.Item(CreditCustomerFields.ccfContactNumber).PrefixRequired = True
        mvClassFields.Item(CreditCustomerFields.ccfAddressNumber).PrefixRequired = True
        mvClassFields.Item(CreditCustomerFields.ccfSalesLedgerAccount).PrefixRequired = True
        mvClassFields.Item(CreditCustomerFields.ccfCompany).PrefixRequired = True
        mvClassFields.Item(CreditCustomerFields.ccfAmendedBy).PrefixRequired = True
        mvClassFields.Item(CreditCustomerFields.ccfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CreditCustomerFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CreditCustomerFields.ccfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CreditCustomerFields.ccfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CreditCustomerRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CreditCustomerRecordSetTypes.ccurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ccu")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pContactNumber As Integer = 0, Optional ByRef pCompany As String = "", Optional ByRef pSalesLedgerAccount As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pContactNumber > 0 Then vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      If Len(pCompany) > 0 Then vWhereFields.Add("company", CDBField.FieldTypes.cftCharacter, pCompany)
      If Len(pSalesLedgerAccount) > 0 Then vWhereFields.Add("sales_ledger_account", CDBField.FieldTypes.cftCharacter, pSalesLedgerAccount)
      If vWhereFields.Count > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(CreditCustomerRecordSetTypes.ccurtAll) & " FROM credit_customers ccu WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CreditCustomerRecordSetTypes.ccurtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CreditCustomerRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CreditCustomerFields.ccfCompany, vFields)
        .SetItem(CreditCustomerFields.ccfSalesLedgerAccount, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CreditCustomerRecordSetTypes.ccurtAll) = CreditCustomerRecordSetTypes.ccurtAll Then
          .SetItem(CreditCustomerFields.ccfContactNumber, vFields)
          .SetItem(CreditCustomerFields.ccfAddressNumber, vFields)
          .SetItem(CreditCustomerFields.ccfCreditCategory, vFields)
          .SetItem(CreditCustomerFields.ccfStopCode, vFields)
          .SetItem(CreditCustomerFields.ccfCreditLimit, vFields)
          .SetItem(CreditCustomerFields.ccfOutstanding, vFields)
          .SetItem(CreditCustomerFields.ccfOnOrder, vFields)
          .SetItem(CreditCustomerFields.ccfCustomerType, vFields)
          .SetItem(CreditCustomerFields.ccfTermsNumber, vFields)
          .SetItem(CreditCustomerFields.ccfTermsPeriod, vFields)
          .SetItem(CreditCustomerFields.ccfTermsFrom, vFields)
          .SetItem(CreditCustomerFields.ccfLastStatementDate, vFields)
          .SetItem(CreditCustomerFields.ccfLastStatementClosingBalance, vFields)
          .SetItem(CreditCustomerFields.ccfLastStatementNumber, vFields)
          .SetItem(CreditCustomerFields.ccfStatementPeriod, vFields)
          .SetItem(CreditCustomerFields.ccfAmendedBy, vFields)
          .SetItem(CreditCustomerFields.ccfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub InitCompanySalesLedgerAccount(ByVal pEnv As CDBEnvironment, ByVal pCompany As String, ByVal pSalesLedgerAccount As String)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CreditCustomerRecordSetTypes.ccurtAll) & " FROM credit_customers ccu WHERE company = '" & pCompany & "' AND sales_ledger_account = '" & pSalesLedgerAccount & "'")
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CreditCustomerRecordSetTypes.ccurtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CreditCustomerFields.ccfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub AdjustOutstanding(ByRef pAmount As Double)
      mvClassFields.Item(CreditCustomerFields.ccfOutstanding).DoubleValue = FixTwoPlaces(Outstanding + pAmount)
    End Sub

    Public Sub AdjustOnOrder(ByVal pAmount As Double)
      'Reduce OnOrder and do not allow to go below zero
      mvClassFields.Item(CreditCustomerFields.ccfOnOrder).DoubleValue = FixTwoPlaces(OnOrder - pAmount)
      If mvClassFields.Item(CreditCustomerFields.ccfOnOrder).DoubleValue < 0 Then mvClassFields.Item(CreditCustomerFields.ccfOnOrder).DoubleValue = 0
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(CreditCustomerFields.ccfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CreditCustomerFields.ccfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CreditCustomerFields.ccfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Company() As String
      Get
        Company = mvClassFields.Item(CreditCustomerFields.ccfCompany).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(CreditCustomerFields.ccfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreditCategory() As String
      Get
        CreditCategory = mvClassFields.Item(CreditCustomerFields.ccfCreditCategory).Value
      End Get
    End Property

    Public ReadOnly Property CreditLimit() As Double
      Get
        CreditLimit = mvClassFields.Item(CreditCustomerFields.ccfCreditLimit).DoubleValue
      End Get
    End Property

    Public ReadOnly Property CustomerType() As String
      Get
        CustomerType = mvClassFields.Item(CreditCustomerFields.ccfCustomerType).Value
      End Get
    End Property

    Public ReadOnly Property LastStatementClosingBalance() As Double
      Get
        LastStatementClosingBalance = mvClassFields.Item(CreditCustomerFields.ccfLastStatementClosingBalance).DoubleValue
      End Get
    End Property

    Public ReadOnly Property LastStatementDate() As String
      Get
        LastStatementDate = mvClassFields.Item(CreditCustomerFields.ccfLastStatementDate).Value
      End Get
    End Property

    Public ReadOnly Property LastStatementNumber() As Integer
      Get
        LastStatementNumber = mvClassFields.Item(CreditCustomerFields.ccfLastStatementNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OnOrder() As Double
      Get
        OnOrder = mvClassFields.Item(CreditCustomerFields.ccfOnOrder).DoubleValue
      End Get
    End Property

    Public ReadOnly Property Outstanding() As Double
      Get
        Outstanding = mvClassFields.Item(CreditCustomerFields.ccfOutstanding).DoubleValue
      End Get
    End Property

    Public ReadOnly Property SalesLedgerAccount() As String
      Get
        SalesLedgerAccount = mvClassFields.Item(CreditCustomerFields.ccfSalesLedgerAccount).Value
      End Get
    End Property

    Public ReadOnly Property StatementPeriod() As Integer
      Get
        StatementPeriod = mvClassFields.Item(CreditCustomerFields.ccfStatementPeriod).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StopCode() As String
      Get
        StopCode = mvClassFields.Item(CreditCustomerFields.ccfStopCode).Value
      End Get
    End Property

    Public ReadOnly Property TermsFrom() As String
      Get
        TermsFrom = mvClassFields.Item(CreditCustomerFields.ccfTermsFrom).Value
      End Get
    End Property

    Public ReadOnly Property TermsNumber() As String
      Get
        TermsNumber = mvClassFields.Item(CreditCustomerFields.ccfTermsNumber).Value
      End Get
    End Property

    Public ReadOnly Property TermsPeriod() As String
      Get
        TermsPeriod = mvClassFields.Item(CreditCustomerFields.ccfTermsPeriod).Value
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pCompany As String, ByVal pSalesLedgerAccount As String, ByVal pCreditCategory As String, ByVal pCreditLimit As Double, ByVal pOutStanding As String, ByVal pCustomerType As String, Optional ByVal pStopCode As String = "", Optional ByVal pTermsNumber As String = "", Optional ByVal pTermsPeriod As String = "", Optional ByVal pTermsFrom As String = "", Optional ByVal pOnOrder As Double = 0)
      Init(pEnv)
      With mvClassFields
        .Item(CreditCustomerFields.ccfContactNumber).IntegerValue = pContactNumber
        .Item(CreditCustomerFields.ccfAddressNumber).IntegerValue = pAddressNumber
        .Item(CreditCustomerFields.ccfCompany).Value = pCompany
        .Item(CreditCustomerFields.ccfSalesLedgerAccount).Value = pSalesLedgerAccount
        .Item(CreditCustomerFields.ccfCreditCategory).Value = pCreditCategory
        .Item(CreditCustomerFields.ccfCreditLimit).DoubleValue = pCreditLimit
        .Item(CreditCustomerFields.ccfOutstanding).Value = pOutStanding
        .Item(CreditCustomerFields.ccfCustomerType).Value = pCustomerType
        .Item(CreditCustomerFields.ccfStopCode).Value = pStopCode
        .Item(CreditCustomerFields.ccfTermsNumber).Value = pTermsNumber
        .Item(CreditCustomerFields.ccfTermsPeriod).Value = pTermsPeriod
        .Item(CreditCustomerFields.ccfTermsFrom).Value = pTermsFrom
        .Item(CreditCustomerFields.ccfOnOrder).DoubleValue = pOnOrder
      End With
      SetValid(CreditCustomerFields.ccfAll)
    End Sub

    Public Sub Update(ByVal pCreditCategory As String, ByVal pCreditLimit As Double, ByVal pCustomerType As String, ByVal pStopCode As String, ByVal pTermsNumber As String, ByVal pTermsPeriod As String, ByVal pTermsFrom As String, ByVal pOutStanding As String, Optional ByVal pAddressNumber As Integer = 0, Optional ByVal pOnOrder As String = "")
      With mvClassFields
        .Item(CreditCustomerFields.ccfCreditCategory).Value = pCreditCategory
        .Item(CreditCustomerFields.ccfCreditLimit).DoubleValue = pCreditLimit
        .Item(CreditCustomerFields.ccfCustomerType).Value = pCustomerType
        .Item(CreditCustomerFields.ccfStopCode).Value = pStopCode
        .Item(CreditCustomerFields.ccfTermsNumber).Value = pTermsNumber
        .Item(CreditCustomerFields.ccfTermsPeriod).Value = pTermsPeriod
        .Item(CreditCustomerFields.ccfTermsFrom).Value = pTermsFrom
        If pOutStanding.Length > 0 Then .Item(CreditCustomerFields.ccfOutstanding).Value = pOutStanding
        If pAddressNumber > 0 Then .Item(CreditCustomerFields.ccfAddressNumber).IntegerValue = pAddressNumber
        If pOnOrder.Length > 0 Then .Item(CreditCustomerFields.ccfOnOrder).Value = pOnOrder
      End With
    End Sub
  End Class
End Namespace
