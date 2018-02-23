

Namespace Access
  Public Class ContactAccount

    Public Enum ContactAccountRecordSetTypes 'These are bit values
      cartAll = &HFFS
      'ADD additional recordset types here
      cartNumber = 1
      cartDetails = 2
      ' BR11347
      cartContactNameAndAddressNumber = 4
      cartContactNumber = 8
      cartAmendedAlias = &H400S
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ContactAccountFields
      cafAll = 0
      cafBankDetailsNumber
      cafContactNumber
      cafSortCode
      cafAccountNumber
      cafAccountName
      cafAmendedBy
      cafAmendedOn
      cafBankPayerName
      cafNotes
      cafDefaultAccount
      cafHistoryOnly
      cafIbanNumber
      cafBicCode
    End Enum

    Public Enum IbanBankChanged
      NotChecked = 0
      NoIbanNumber
      SameIbanNumber
      SameIbanBank
      BankChanged
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAmendedValid As Boolean

    ' BR 11347
    Private mvAccountsForContact As ContactAccounts
    Private mvForenames As String
    Private mvSurname As String
    Private mvAddressNumber As Integer
    Private mvIbanNumber As String
    Private mvBicCode As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "contact_accounts"
          .Add("bank_details_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("sort_code")
          .Add("account_number")
          .Add("account_name")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("bank_payer_name")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("default_account")
          .Add("history_only")
          .Add("iban_number")
          .Add("bic_code")
        End With
        mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If

      mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).PrefixRequired = True
      mvClassFields.Item(ContactAccountFields.cafContactNumber).PrefixRequired = True
      mvClassFields.Item(ContactAccountFields.cafIbanNumber).PrefixRequired = True
      mvClassFields.Item(ContactAccountFields.cafBicCode).PrefixRequired = True
      mvClassFields.Item(ContactAccountFields.cafDefaultAccount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDefaultBankAccount)
      mvClassFields.Item(ContactAccountFields.cafHistoryOnly).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHistoryOnlyAccount)
      mvClassFields.Item(ContactAccountFields.cafIbanNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers)
      mvClassFields.Item(ContactAccountFields.cafBicCode).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers)

      mvAmendedValid = False
      mvExisting = False

      ' BR 11347
      mvForenames = ""
      mvSurname = ""
      mvAddressNumber = 0

    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(ContactAccountFields.cafDefaultAccount).Value = "N"
      mvClassFields.Item(ContactAccountFields.cafHistoryOnly).Value = "N"
    End Sub

    Private Sub SetValid(ByRef pField As ContactAccountFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).IntegerValue = 0 Then mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).IntegerValue = mvEnv.GetControlNumber("BD")
      If pField = ContactAccountFields.cafAll And Not mvAmendedValid Then
        mvClassFields.Item(ContactAccountFields.cafAmendedOn).Value = TodaysDate()
        mvClassFields.Item(ContactAccountFields.cafAmendedBy).Value = mvEnv.User.UserID
      End If
      If (pField = ContactAccountFields.cafAll Or pField = ContactAccountFields.cafBankPayerName) And Len(mvClassFields.Item(ContactAccountFields.cafBankPayerName).Value) = 0 Then
        mvClassFields.Item(ContactAccountFields.cafBankPayerName).Value = UCase(Left(mvClassFields.Item(ContactAccountFields.cafAccountName).Value, 18))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ContactAccountRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ContactAccountRecordSetTypes.cartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ca")
      Else
        vFields = "ca.bank_details_number"
        If (pRSType And ContactAccountRecordSetTypes.cartDetails) > 0 Then vFields = vFields & ",sort_code,account_number,account_name,iban_number,bic_code"
        If (pRSType And ContactAccountRecordSetTypes.cartAmendedAlias) > 0 Then vFields = vFields & ",ca.amended_on AS ca_amended_on, ca.amended_by AS ca_amended_by"
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBankDetailsNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pBankDetailsNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactAccountRecordSetTypes.cartAll) & " FROM contact_accounts ca WHERE bank_details_number = " & pBankDetailsNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ContactAccountRecordSetTypes.cartAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitByAccount(ByVal pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pSortCode As String, ByRef pAccountNumber As String)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactAccountRecordSetTypes.cartAll) & " FROM contact_accounts ca WHERE contact_number = " & pContactNumber & " AND sort_code = '" & pSortCode & "' AND account_number = '" & pAccountNumber & "'")
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, ContactAccountRecordSetTypes.cartAll)
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitByAccount(ByVal pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pIbanNumber As String)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactAccountRecordSetTypes.cartAll) & " FROM contact_accounts ca WHERE contact_number = " & pContactNumber & " AND iban_number = '" & pIbanNumber & "'")
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, ContactAccountRecordSetTypes.cartAll)
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub
    Public Sub InitDefaultAccountByContact(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactAccountRecordSetTypes.cartAll) & " FROM contact_accounts ca WHERE contact_number = " & pContactNumber & " AND default_account = 'Y'")
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, ContactAccountRecordSetTypes.cartAll)
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactAccountRecordSetTypes, Optional ByVal pForenames As String = "", Optional ByVal pSurname As String = "", Optional ByVal pAddressNo As Integer = 0)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ContactAccountFields.cafBankDetailsNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ContactAccountRecordSetTypes.cartDetails) > 0 Then
          .SetItem(ContactAccountFields.cafSortCode, vFields)
          .SetItem(ContactAccountFields.cafAccountNumber, vFields)
          .SetItem(ContactAccountFields.cafAccountName, vFields)
          .SetOptionalItem(ContactAccountFields.cafIbanNumber, vFields)
          .SetOptionalItem(ContactAccountFields.cafBicCode, vFields)
        End If
        If (pRSType And ContactAccountRecordSetTypes.cartAll) = ContactAccountRecordSetTypes.cartAll Then
          .SetItem(ContactAccountFields.cafContactNumber, vFields)
          .SetItem(ContactAccountFields.cafAmendedBy, vFields)
          .SetItem(ContactAccountFields.cafAmendedOn, vFields)
          .SetItem(ContactAccountFields.cafBankPayerName, vFields)
          .SetItem(ContactAccountFields.cafNotes, vFields)
          .SetOptionalItem(ContactAccountFields.cafDefaultAccount, vFields)
          .SetOptionalItem(ContactAccountFields.cafHistoryOnly, vFields)
        End If
        If (pRSType And ContactAccountRecordSetTypes.cartAmendedAlias) > 0 Then
          mvClassFields.Item(ContactAccountFields.cafAmendedBy).SetValue = vFields("ca_amended_by").Value
          mvClassFields.Item(ContactAccountFields.cafAmendedOn).SetValue = vFields("ca_amended_on").Value
        End If
        ' BR11347
        If (pRSType And ContactAccountRecordSetTypes.cartContactNumber) > 0 Then .SetItem(ContactAccountFields.cafContactNumber, vFields)

      End With

      ' BR11347
      If (pRSType And ContactAccountRecordSetTypes.cartContactNameAndAddressNumber) > 0 Then
        ContactForenames = pForenames
        ContactSurname = pSurname
        AddressNumber = pAddressNo
      End If

    End Sub

    Public Sub Save()
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vChanged As Boolean
      Dim vTrans As Boolean

      SetValid(ContactAccountFields.cafAll)

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDefaultBankAccount) Then
        If mvExisting Then
          'Update existing ContactAccount so check for DefaultAccount flag having changed
          vChanged = mvClassFields.Item("default_account").ValueChanged
        Else
          'New ContactAccount so always want to ensure that no other records have DefaultAccount set to 'Y'
          vChanged = True
        End If
      End If

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      If DefaultAccount = True And vChanged = True Then
        'DefaultAccount is set and previously it was not, so update ALL other ContactAccounts to set DefaultAccount = 'N'
        vWhereFields = New CDBFields
        vUpdateFields = New CDBFields
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add("default_account", CDBField.FieldTypes.cftCharacter, "Y")
        vWhereFields.Add("bank_details_number", BankDetailsNumber, CDBField.FieldWhereOperators.fwoNotEqual)
        vUpdateFields.Add("default_account", CDBField.FieldTypes.cftCharacter, "N")
        mvEnv.Connection.UpdateRecords("contact_accounts", vUpdateFields, vWhereFields, False)
      End If

      mvClassFields.Save(mvEnv, mvExisting, mvEnv.User.UserID, True)

      If vTrans Then mvEnv.Connection.CommitTransaction()

    End Sub

    Public Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pContactNumber As Integer, Optional ByRef pSortCode As String = "", Optional ByRef pAccountNumber As String = "", Optional ByRef pAccountName As String = "", Optional ByRef pBankPayerName As String = "", Optional ByRef pNotes As String = "", Optional ByVal pDefaultAccount As String = "", Optional ByVal pHistoryOnly As String = "", Optional ByVal pIbanNumber As String = "", Optional ByVal pBicCode As String = "")
      Dim vCount As Integer

      With mvClassFields
        .Item(ContactAccountFields.cafContactNumber).IntegerValue = pContactNumber
        If pSortCode.Length > 0 Then .Item(ContactAccountFields.cafSortCode).Value = Trim(Replace(pSortCode, "-", ""))
        If pAccountNumber.Length > 0 Then .Item(ContactAccountFields.cafAccountNumber).Value = pAccountNumber
        .Item(ContactAccountFields.cafAccountName).Value = pAccountName
        .Item(ContactAccountFields.cafBankPayerName).Value = pBankPayerName
        .Item(ContactAccountFields.cafNotes).Value = pNotes
        If pDefaultAccount.Length > 0 Then .Item(ContactAccountFields.cafDefaultAccount).Value = pDefaultAccount
        If pHistoryOnly.Length > 0 Then .Item(ContactAccountFields.cafHistoryOnly).Value = pHistoryOnly
        If pIbanNumber.Length > 0 Then .Item(ContactAccountFields.cafIbanNumber).Value = pIbanNumber
        If pBicCode.Length > 0 Then .Item(ContactAccountFields.cafBicCode).Value = pBicCode
      End With

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDefaultBankAccount) = True And Len(pDefaultAccount) = 0 Then
        'Creating the first ContactAccount will always set DefaultAccount to 'Y' (already set to 'N' in SetDefaults)
        vCount = mvEnv.Connection.GetCount("contact_accounts", Nothing, "contact_number = " & pContactNumber)
        If vCount = 0 Then mvClassFields.Item(ContactAccountFields.cafDefaultAccount).Value = "Y"
      End If
    End Sub

    Public Sub Delete(ByVal pLogname As String, ByVal pAudit As Boolean)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("bank_details_number", CDBField.FieldTypes.cftInteger, BankDetailsNumber)
      If mvEnv.Connection.GetCount("bankers_orders", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Standing Orders") 'Cannot Delete this Record as it is Referenced by %1
      ElseIf mvEnv.Connection.GetCount("direct_debits", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Direct Debits") 'Cannot Delete this Record as it is Referenced by %1
      ElseIf mvEnv.Connection.GetCount("financial_history", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Financial History") 'Cannot Delete this Record as it is Referenced by %1
      ElseIf mvEnv.Connection.GetCount("batch_transactions", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Batch Transactions") 'Cannot Delete this Record as it is Referenced by %1
      ElseIf mvEnv.Connection.GetCount("purchase_invoices", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Purchase Invoices") 'Cannot Delete this Record as it is Referenced by %1
      ElseIf mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbIbanBicNumbers) Then
        vWhereFields.Clear()
        vWhereFields.Add("previous_bank_details_number", BankDetailsNumber)
        If mvEnv.Connection.GetCount("direct_debits", vWhereFields) > 0 Then
          RaiseError(DataAccessErrors.daeReferencedInOtherTable, "Direct Debits") 'Cannot Delete this Record as it is Referenced by %1
        End If
      End If
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pLogname, pAudit)
    End Sub

    Public Sub Update(ByRef pAccountName As String, ByRef pBankPayerName As String, Optional ByRef pSortCode As String = "", Optional ByRef pAccountNumber As String = "", Optional ByRef pNotes As String = "", Optional ByVal pDefaultAccount As String = "", Optional ByVal pHistoryOnly As String = "", Optional ByVal pIbanNumber As String = "", Optional ByVal pBicCode As String = "")
      With mvClassFields
        .Item(ContactAccountFields.cafAccountName).Value = pAccountName
        .Item(ContactAccountFields.cafBankPayerName).Value = pBankPayerName
        .Item(ContactAccountFields.cafSortCode).Value = Trim(Replace(pSortCode, "-", ""))
        .Item(ContactAccountFields.cafAccountNumber).Value = pAccountNumber
        If Len(pNotes) > 0 Then .Item(ContactAccountFields.cafNotes).Value = pNotes
        If pDefaultAccount.Length > 0 Then .Item(ContactAccountFields.cafDefaultAccount).Value = pDefaultAccount
        If pHistoryOnly.Length > 0 Then .Item(ContactAccountFields.cafHistoryOnly).Value = pHistoryOnly
        .Item(ContactAccountFields.cafIbanNumber).Value = pIbanNumber
        .Item(ContactAccountFields.cafBicCode).Value = pBicCode
      End With
    End Sub
    Public Sub Update(ByVal pParams As CDBParameters)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False AndAlso vClassField.InDatabase Then
          If pParams.Exists(vClassField.ProperName) Then vClassField.Value = pParams(vClassField.ProperName).Value
        End If
      Next
    End Sub

    Public Sub ReplaceMissingRecord(ByVal pEnv As CDBEnvironment, ByRef pBankDetailsNumber As Integer, ByRef pContactNumber As Integer)
      'Used to initialise the class in order to replace a missing address record (called from integrity check)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).Value = CStr(pBankDetailsNumber)
      mvClassFields.Item(ContactAccountFields.cafContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(ContactAccountFields.cafSortCode).Value = "000000"
      mvClassFields.Item(ContactAccountFields.cafAccountNumber).Value = "UNKNOWN"
      Me.Save()
    End Sub

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(ContactAccountFields.cafAmendedOn).Value = pAmendedOn
      mvClassFields.Item(ContactAccountFields.cafAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Function IsPotentialDuplicate(ByRef pContactAc As ContactAccount, Optional ByVal pUseSortCode As Boolean = False) As Boolean
      ' BR 11347
      Dim vFirstBankAccount As String
      Dim vSecondBankAccount As String
      Dim vFirstSortCode As String
      Dim vSecondSortCode As String
      Dim vFirstIbanNumber As String
      Dim vSecondIbanNumber As String

      vFirstBankAccount = Trim(mvClassFields.Item(ContactAccountFields.cafAccountNumber).Value)
      vSecondBankAccount = Trim(pContactAc.AccountNumber)

      vFirstIbanNumber = Trim(mvClassFields.Item(ContactAccountFields.cafIbanNumber).Value)
      vSecondIbanNumber = Trim(pContactAc.IbanNumber)

      If vFirstIbanNumber.Length > 0 AndAlso vSecondIbanNumber.Length > 0 AndAlso StrComp(vFirstIbanNumber, vSecondIbanNumber, CompareMethod.Text) = 0 Then
        ' IbanNumber is duplicate
        IsPotentialDuplicate = True
      ElseIf StrComp(vFirstBankAccount, vSecondBankAccount, CompareMethod.Text) = 0 Then
        ' Now need to see if we must use sort code for comparison as well
        If pUseSortCode Then
          vFirstSortCode = Trim(mvClassFields.Item(ContactAccountFields.cafSortCode).Value)
          vSecondSortCode = Trim(pContactAc.SortCode)
          IsPotentialDuplicate = (StrComp(vFirstSortCode, vSecondSortCode, CompareMethod.Text) = 0)
        Else
          ' Nope just the A/c number is dandy
          IsPotentialDuplicate = True
        End If
      Else
        IsPotentialDuplicate = False
      End If
    End Function

    Private Sub SetAccountsForContact(ByVal pContactNumber As Integer)
      ' BR 11347
      Dim vContactAccount As New ContactAccount
      Dim vRS As CDBRecordSet
      Dim vSQL As String

      vContactAccount.Init(mvEnv)
      mvAccountsForContact = New ContactAccounts
      vSQL = "SELECT " & vContactAccount.GetRecordSetFields(ContactAccountRecordSetTypes.cartNumber) & " FROM contact_accounts ca"
      vSQL = vSQL & " WHERE ca.contact_number = " & pContactNumber
      vSQL = vSQL & " ORDER BY ca.contact_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        If vRS.Fields("iban_number").Value.Length > 0 Then
          vContactAccount = mvAccountsForContact.Add((vRS.Fields("iban_number").Value))
        Else
          vContactAccount = mvAccountsForContact.Add((vRS.Fields("account_number").Value))
        End If

        vContactAccount.InitFromRecordSet(mvEnv, vRS, ContactAccountRecordSetTypes.cartNumber)
      End While
      vRS.CloseRecordSet()
    End Sub

    Public Function ContactSurnamesMatch(ByVal pSurname As String, ByVal pUseSoundex As Boolean) As Boolean
      ' BR 11347/12014
      ContactSurnamesMatch = StrComp(pSurname, ContactSurname, CompareMethod.Text) = 0

      ' If we have already matched the surnames, no need to perform Soundex
      If Not ContactSurnamesMatch Then
        If pUseSoundex Then
          ' Need to pop in a quick soundex check
          ContactSurnamesMatch = (GetSoundexCode(pSurname) = GetSoundexCode(ContactSurname))
        End If
      End If

    End Function

    ''' <summary>Has the bank portion of the IBAN Number changed?</summary>
    Friend Function HasIbanBankChanged(ByVal pOldIbanNumber As String) As IbanBankChanged
      Dim vBankChanged As IbanBankChanged = IbanBankChanged.NotChecked
      If (mvClassFields.Item(ContactAccountFields.cafIbanNumber).Value.Length > 0 AndAlso pOldIbanNumber.Length > 0) Then
        vBankChanged = IbanBankChanged.SameIbanNumber
        If mvClassFields.Item(ContactAccountFields.cafIbanNumber).Value <> pOldIbanNumber Then
          'IBAN Numbers are different
          Dim vNewIbanCountry As CountryIbanNumber = mvEnv.GetCountryIban(IbanNumber)
          Dim vOldIbanCountry As CountryIbanNumber = mvEnv.GetCountryIban(pOldIbanNumber)
          If vOldIbanCountry.IbanCountry = vNewIbanCountry.IbanCountry Then
            vBankChanged = IbanBankChanged.SameIbanBank
            If vOldIbanCountry.GetBankCodeFromIban(pOldIbanNumber) <> vNewIbanCountry.GetBankCodeFromIban(IbanNumber) Then vBankChanged = IbanBankChanged.BankChanged
          Else
            'Different Country so bank must have changed
            vBankChanged = IbanBankChanged.BankChanged
          End If
        End If
      Else
        vBankChanged = IbanBankChanged.NoIbanNumber
      End If
      Return vBankChanged
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property AccountName() As String
      Get
        AccountName = mvClassFields.Item(ContactAccountFields.cafAccountName).Value
      End Get
      Set(ByVal Value As String)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(ContactAccountFields.cafAccountName).Value = Value
      End Set
    End Property

    Public Property AccountNumber() As String
      Get
        AccountNumber = mvClassFields.Item(ContactAccountFields.cafAccountNumber).Value
      End Get
      Set(ByVal Value As String)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(ContactAccountFields.cafAccountNumber).Value = Value
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(ContactAccountFields.cafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ContactAccountFields.cafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BankDetailsNumber() As Integer
      Get
        SetValid(ContactAccountFields.cafBankDetailsNumber)
        BankDetailsNumber = mvClassFields.Item(ContactAccountFields.cafBankDetailsNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BankPayerName() As String
      Get
        SetValid(ContactAccountFields.cafBankPayerName)
        BankPayerName = mvClassFields.Item(ContactAccountFields.cafBankPayerName).Value
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ContactAccountFields.cafContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        'Can be used by BACS Rejections
        mvClassFields.Item(ContactAccountFields.cafContactNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property DefaultAccount() As Boolean
      Get
        'Null = 'N'
        DefaultAccount = mvClassFields.Item(ContactAccountFields.cafDefaultAccount).Bool
      End Get
    End Property

    Public ReadOnly Property HistoryOnly() As Boolean
      Get
        HistoryOnly = mvClassFields.Item(ContactAccountFields.cafHistoryOnly).Bool
      End Get
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(ContactAccountFields.cafNotes).MultiLineValue
      End Get
      Set(ByVal Value As String)
        'Used by BACS Rejections
        mvClassFields.Item(ContactAccountFields.cafNotes).Value = Value
      End Set
    End Property

    Public ReadOnly Property FormattedSortCode() As String
      Get
        Dim vResult As String
        vResult = SortCode
        FormattedSortCode = Left(vResult, 2) & "-" & Mid(vResult, 3, 2) & "-" & Right(vResult, 2)
      End Get
    End Property

    Public Property SortCode() As String
      Get
        SortCode = mvClassFields.Item(ContactAccountFields.cafSortCode).Value
      End Get
      Set(ByVal Value As String)
        'Used by AutoSO Reconciliation Data Fix
        mvClassFields.Item(ContactAccountFields.cafSortCode).Value = Value
      End Set
    End Property

    Public ReadOnly Property AccountsForContact(Optional ByVal pPopulateCollection As Boolean = False) As ContactAccounts
      Get
        ' BR 11347
        If mvAccountsForContact Is Nothing Then
          mvAccountsForContact = New ContactAccounts
          If pPopulateCollection Then SetAccountsForContact(ContactNumber)
        End If
        AccountsForContact = mvAccountsForContact
      End Get
    End Property

    ' BR11347
    Public Property ContactForenames() As String
      Get
        ContactForenames = mvForenames
      End Get
      Set(ByVal Value As String)
        mvForenames = Value
      End Set
    End Property
    Public Property ContactSurname() As String
      Get
        ContactSurname = mvSurname
      End Get
      Set(ByVal Value As String)
        mvSurname = Value
      End Set
    End Property
    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvAddressNumber
      End Get
      Set(ByVal Value As Integer)
        mvAddressNumber = Value
      End Set
    End Property
    Public Property IbanNumber() As String
      Get
        IbanNumber = mvClassFields.Item(ContactAccountFields.cafIbanNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ContactAccountFields.cafIbanNumber).Value = Value
      End Set
    End Property
    Public Property BicCode() As String
      Get
        BicCode = mvClassFields.Item(ContactAccountFields.cafBicCode).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ContactAccountFields.cafBicCode).Value = Value
      End Set
    End Property
  End Class
End Namespace
