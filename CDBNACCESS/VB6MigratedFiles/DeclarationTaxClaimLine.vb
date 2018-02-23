

Namespace Access
  Public Interface ITaxClaimLine
    ReadOnly Property CdNumber() As Integer
    ReadOnly Property ClaimNumber() As Integer
    ReadOnly Property DeclarationOrCovenantNumber() As String
    ReadOnly Property BatchNumber() As Integer
    ReadOnly Property TransactionNumber() As Integer
    ReadOnly Property LineNumber() As Integer
    Property NetAmount() As Double
    Property AmountClaimed() As Double
    Sub UpdateNetAmount(ByVal pNewValue As Double, ByRef pClaim As DeclarationTaxClaim)
    Sub Delete()
    Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
  End Interface

  Public Class DeclarationTaxClaimLine
    Implements ITaxClaimLine

    Public Enum DeclarationTaxClaimLineRecordSetTypes 'These are bit values
      dtclrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DeclarationTaxClaimLineFields
      dtclfAll = 0
      dtclfClaimNumber
      dtclfCdNumber
      dtclfDeclarationOrCovenantNumber
      dtclfBatchNumber
      dtclfTransactionNumber
      dtclfLineNumber
      dtclfNetAmount
      dtclfAmountClaimed
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvTaxPercent As Integer

    Public Sub New()

    End Sub

    Public Sub New(pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "declaration_tax_claim_lines"
          .TableAlias = "dtcl"
          .Add("claim_number", CDBField.FieldTypes.cftLong)
          .Add("cd_number", CDBField.FieldTypes.cftLong)
          .Add("declaration_or_covenant_number")
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfBatchNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfTransactionNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfLineNumber).PrefixRequired = True

        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfLineNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfDeclarationOrCovenantNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As DeclarationTaxClaimLineFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As DeclarationTaxClaimLineRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DeclarationTaxClaimLineRecordSetTypes.dtclrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dtcl")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pClaimNumber As Integer = 0, Optional ByRef pCdNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pClaimNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DeclarationTaxClaimLineRecordSetTypes.dtclrtAll) & " FROM declaration_tax_claim_lines dtcl WHERE claim_number = " & pClaimNumber & " AND cd_number = " & pCdNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DeclarationTaxClaimLineRecordSetTypes.dtclrtAll)
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
    Public Sub InitFromUnclaimed(ByVal pEnv As CDBEnvironment, ByVal pClaim As DeclarationTaxClaim, ByVal pUnClaimed As ITaxClaimLine, ByVal pTaxPercent As Integer)
      'pUnClaimed could be DeclarationTaxClaimLine or DeclarationPotentialLine
      Dim vAmount As Double

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvTaxPercent = pTaxPercent
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfClaimNumber).Value = CStr(pClaim.ClaimNumber)
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfCdNumber).IntegerValue = pUnClaimed.CdNumber
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfDeclarationOrCovenantNumber).Value = pUnClaimed.DeclarationOrCovenantNumber
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfBatchNumber).IntegerValue = pUnClaimed.BatchNumber
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfTransactionNumber).IntegerValue = pUnClaimed.TransactionNumber
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfLineNumber).IntegerValue = pUnClaimed.LineNumber
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfNetAmount).DoubleValue = pUnClaimed.NetAmount
      vAmount = pUnClaimed.AmountClaimed
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).Value = CStr(vAmount)
      pClaim.CalculatedTaxAmount = pClaim.CalculatedTaxAmount + vAmount
      pClaim.TotalNetAmount = pClaim.TotalNetAmount + NetAmount
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DeclarationTaxClaimLineRecordSetTypes, Optional ByRef pTaxPercent As Integer = 0)
      Dim vFields As CDBFields

      mvEnv = pEnv
      If pTaxPercent > 0 Then mvTaxPercent = pTaxPercent
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DeclarationTaxClaimLineFields.dtclfBatchNumber, vFields)
        .SetItem(DeclarationTaxClaimLineFields.dtclfTransactionNumber, vFields)
        .SetItem(DeclarationTaxClaimLineFields.dtclfLineNumber, vFields)
        .SetItem(DeclarationTaxClaimLineFields.dtclfNetAmount, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DeclarationTaxClaimLineRecordSetTypes.dtclrtAll) = DeclarationTaxClaimLineRecordSetTypes.dtclrtAll Then
          .SetItem(DeclarationTaxClaimLineFields.dtclfClaimNumber, vFields)
          .SetItem(DeclarationTaxClaimLineFields.dtclfCdNumber, vFields)
          .SetItem(DeclarationTaxClaimLineFields.dtclfDeclarationOrCovenantNumber, vFields)
          .SetItem(DeclarationTaxClaimLineFields.dtclfAmountClaimed, vFields)
        End If
      End With
    End Sub

    Public Sub InitWithPrimaryKey(ByVal pWhereFields As CDBFields)
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields(), mvClassFields.TableNameAndAlias, pWhereFields)
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(vRecordSet)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub
    Public Overridable Sub InitFromRecordSet(ByVal pRecordSet As CDBRecordSet)
      InitClassFields()
      Dim vFields As CDBFields = pRecordSet.Fields
      mvExisting = True
      For Each vClassField As ClassField In mvClassFields
        If vClassField.InDatabase AndAlso vClassField.FieldType <> CDBField.FieldTypes.cftBulk Then
          If vClassField.FieldType = CDBField.FieldTypes.cftBinary Then
            vClassField.SetValue = vFields(vClassField.Name).Value
            vClassField.ByteValue = vFields(vClassField.Name).ByteValue
          Else
            vClassField.SetValue = vFields(vClassField.Name).Value
          End If
        Else
          If vFields.Exists(vClassField.Name) Then
            vClassField.SetValue = vFields(vClassField.Name).Value
          End If
        End If
      Next
    End Sub
    Public Overridable Function GetRecordSetFields() As String
      If mvClassFields Is Nothing Then InitClassFields()
      Return mvClassFields.FieldNames(mvEnv, "dtcl")
    End Function
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False) Implements ITaxClaimLine.Save
      SetValid(DeclarationTaxClaimLineFields.dtclfAll)
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

    Public Property AmountClaimed() As Double Implements ITaxClaimLine.AmountClaimed
      Get
        AmountClaimed = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property BatchNumber() As Integer Implements ITaxClaimLine.BatchNumber
      Get
        BatchNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CdNumber() As Integer Implements ITaxClaimLine.CdNumber
      Get
        CdNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfCdNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ClaimNumber() As Integer Implements ITaxClaimLine.ClaimNumber
      Get
        ClaimNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfClaimNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DeclarationOrCovenantNumber() As String Implements ITaxClaimLine.DeclarationOrCovenantNumber
      Get
        DeclarationOrCovenantNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfDeclarationOrCovenantNumber).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer Implements ITaxClaimLine.LineNumber
      Get
        LineNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfLineNumber).IntegerValue
      End Get
    End Property

    Public Property NetAmount() As Double Implements ITaxClaimLine.NetAmount
      Get
        NetAmount = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfNetAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfNetAmount).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property TransactionNumber() As Integer Implements ITaxClaimLine.TransactionNumber
      Get
        TransactionNumber = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfTransactionNumber).IntegerValue
      End Get
    End Property

    Private Function CalculateAmountClaimed(ByVal pNetAmount As Double) As Double
      Dim vAmount As Double

      vAmount = FixTwoPlaces((pNetAmount / (100 - mvTaxPercent) * 100) - pNetAmount)
      CalculateAmountClaimed = vAmount
    End Function

    Public Sub Delete() Implements ITaxClaimLine.Delete
      mvEnv.Connection.DeleteRecords("declaration_tax_claim_lines", mvClassFields.WhereFields)
    End Sub

    Public Sub UpdateNetAmount(ByVal pNewValue As Double, ByRef pClaim As DeclarationTaxClaim) Implements ITaxClaimLine.UpdateNetAmount
      'This is called from the Tax Claim process to change the claim line amount
      'Also updates the claim itself with the new value
      Dim vOldClaimAmount As Double
      Dim vOldNetAmount As Double

      vOldClaimAmount = mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).DoubleValue
      vOldNetAmount = NetAmount
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).Value = CStr(CalculateAmountClaimed(pNewValue))
      mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfNetAmount).Value = CStr(pNewValue)
      pClaim.CalculatedTaxAmount = FixTwoPlaces(pClaim.CalculatedTaxAmount - (vOldClaimAmount - mvClassFields.Item(DeclarationTaxClaimLineFields.dtclfAmountClaimed).DoubleValue))
      pClaim.TotalNetAmount = pClaim.TotalNetAmount - (vOldNetAmount - NetAmount)
    End Sub
  End Class
End Namespace
