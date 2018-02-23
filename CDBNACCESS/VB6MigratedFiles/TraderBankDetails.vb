

Namespace Access
  Public Class TraderBankDetails

    Private mvBankDetailsNumber As Integer = 0

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pExistingTrans As Boolean, ByVal pFA As Boolean, ByVal pContactNumber As Integer, ByVal pBankDetailsNumber As Integer, ByVal pSortCode As String, ByVal pAccountNumber As String, ByVal pAccountName As String, Optional ByVal pBranchName As String = "", Optional ByVal pNewBank As Boolean = False, Optional ByVal pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0)
      Dim vParams As New CDBParameters
      With vParams
        .Add("SortCode", pSortCode)
        .Add("AccountNumber", pAccountNumber)
        .Add("AccountName", pAccountName)
        .Add("BranchName", pBranchName)
        .Add("ContactNumber", pContactNumber)
        .Add("BankDetailsNumber", pBankDetailsNumber)
      End With
      Init(pEnv, vParams, pExistingTrans, pFA, pNewBank, pBatchNumber, pTransactionNumber)
    End Sub
    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pParams As CDBParameters, ByVal pExistingTrans As Boolean, ByVal pFA As Boolean, ByVal pNewBank As Boolean, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vBank As Bank
      Dim vCA As ContactAccount
      Dim vFields As CDBFields
      Dim vHContact As Integer
      Dim vSQL As String

      mvBankDetailsNumber = pParams.ParameterExists("BankDetailsNumber").IntegerValue
      If mvBankDetailsNumber < 0 Then mvBankDetailsNumber = 0

      If ((pParams.ParameterExists("SortCode").Value.Length > 0 AndAlso pParams.ParameterExists("AccountNumber").Value.Length > 0) _
      OrElse (pParams.ParameterExists("IbanNumber").Value.Length > 0)) _
      AndAlso pParams.ParameterExists("AccountName").Value.Length > 0 Then
        'Must have AccountName and SortCode/AccountNumber or IbanNumber
        Dim vSortCode As String = pParams.ParameterExists("SortCode").Value
        If vSortCode.Length > 0 Then vSortCode = vSortCode.Replace("-", "")
        If vSortCode.Length = 0 Then pNewBank = False

        If pNewBank Then
          vBank = New Bank
          With vBank
            .Init(pEnv)
            .Create(vSortCode, pParams.ParameterExists("BranchName").Value)
            .Save()
          End With
        End If

        vCA = New ContactAccount
        vCA.Init(pEnv, mvBankDetailsNumber)

        If vCA.Existing = True And pExistingTrans = True And pFA = False Then
          'This is editing an existing transaction and not a financial adjustment
          'So need to check whether the transaction was originally against the holding contact
          'and if it was then either update or create the Contact Accounts
          vHContact = IntegerValue(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber))
          If vHContact > 0 And (vCA.ContactNumber = vHContact) Then
            vFields = New CDBFields
            vFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBatchNumber)
            vFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pTransactionNumber)
            vFields.Add("contact_number", CDBField.FieldTypes.cftLong, vHContact)
            If pEnv.Connection.GetCount("batch_transactions", vFields) > 0 Then
              'Original transaction was using holding Contact
              With vFields
                .Clear()
                .Add("bank_details_number", CDBField.FieldTypes.cftLong, mvBankDetailsNumber)
                .Add("contact_number", CDBField.FieldTypes.cftLong, vHContact)
                vSQL = pEnv.Connection.WhereClause(vFields) & " AND ("
                .Clear()
                .Add("batch_number#1", pBatchNumber, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
                .Add("transaction_number", pTransactionNumber, CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
                .Add("batch_number#2", pBatchNumber, CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoCloseBracket)
              End With
              vSQL = vSQL & pEnv.Connection.WhereClause(vFields) & " )"
              If pEnv.Connection.GetCount("batch_transactions", Nothing, vSQL) = 0 Then
                'No other transactions so update this CA record
                vCA.ContactNumber = pParams("ContactNumber").IntegerValue
              Else
                'Other transactions using this CA record so create a new one
                mvBankDetailsNumber = 0
              End If
            End If
          End If
        End If

        Dim vAccountName As String = pParams("AccountName").Value
        Dim vAccountNumber As String = pParams.ParameterExists("AccountNumber").Value
        Dim vIbanNumber As String = pParams.ParameterExists("IbanNumber").Value
        Dim vBicCode As String = pParams.ParameterExists("BicCode").Value
        Dim vBankPayerName As String = pParams.ParameterExists("BankPayerName").Value
        If vCA.Existing = True AndAlso vCA.AccountName = vAccountName Then
          'Name unchanged so leave as is
          vBankPayerName = vCA.BankPayerName
        ElseIf vBankPayerName.Length = 0 Then
          vBankPayerName = Substring(vAccountName, 0, 18)
        End If
        vBankPayerName = vBankPayerName.ToUpper   'Should always be in upper case
        If pParams.Exists("BankPayerName") = False Then pParams.Add("BankPayerName")
        pParams("BankPayerName").Value = vBankPayerName
        With vCA
          If mvBankDetailsNumber = 0 Then
            'Create a new record
            .Init(pEnv)
            .Create(pParams("ContactNumber").IntegerValue, vSortCode, vAccountNumber, vAccountName, vBankPayerName, pParams.ParameterExists("Notes").Value, , , vIbanNumber, vBicCode)
            .Save()
          ElseIf .Existing Then
            'Update existing record
            .Update(pParams)
            .SaveChanges(pEnv.User.UserID, True)
          End If
        End With
        mvBankDetailsNumber = vCA.BankDetailsNumber
      End If

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property BankDetailsNumber() As Integer
      Get
        BankDetailsNumber = mvBankDetailsNumber
      End Get
    End Property
  End Class
End Namespace
