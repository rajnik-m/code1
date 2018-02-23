Namespace Access
  Public Class DeclarationPotentialLine
    Implements ITaxClaimLine

    Public Enum DeclarationPotentialLineRecordSetTypes 'These are bit values
      dplrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DeclarationPotentialLineFields
      dplfAll = 0
      dplfCdNumber
      dplfDeclarationOrCovenantNumber
      dplfContactNumber
      dplfBatchNumber
      dplfTransactionNumber
      dplfLineNumber
      dplfNetAmount
      dplfAmountClaimed
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvClaimNumber As Integer
    Private mvTaxPercent As Integer

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = Me.TableName
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("cd_number", CDBField.FieldTypes.cftLong)
          .Add("declaration_or_covenant_number")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationPotentialLineFields.dplfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationPotentialLineFields.dplfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationPotentialLineFields.dplfLineNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(DeclarationPotentialLineFields.dplfCdNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfContactNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfBatchNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfTransactionNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfLineNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As DeclarationPotentialLineFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    Private Function CalculateAmountClaimed(ByVal pNetAmount As Double) As Double
      Dim vAmount As Double

      vAmount = FixTwoPlaces((pNetAmount / (100 - mvTaxPercent) * 100) - pNetAmount)
      CalculateAmountClaimed = vAmount
    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As DeclarationPotentialLineRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DeclarationPotentialLineRecordSetTypes.dplrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dpl")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0, Optional ByRef pDeclarationOrCovenantNumber As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DeclarationPotentialLineRecordSetTypes.dplrtAll) & " FROM " & Me.TableName & " dpl WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber & " AND declaration_or_covenant_number = '" & pDeclarationOrCovenantNumber & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DeclarationPotentialLineRecordSetTypes.dplrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DeclarationPotentialLineRecordSetTypes, Optional ByRef pTaxPercent As Integer = 0, Optional ByRef pClaimNumber As Integer = 0)
      Dim vFields As CDBFields

      mvEnv = pEnv
      If pTaxPercent > 0 Then mvTaxPercent = pTaxPercent
      mvClaimNumber = pClaimNumber
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber, vFields)
        .SetItem(DeclarationPotentialLineFields.dplfBatchNumber, vFields)
        .SetItem(DeclarationPotentialLineFields.dplfTransactionNumber, vFields)
        .SetItem(DeclarationPotentialLineFields.dplfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DeclarationPotentialLineRecordSetTypes.dplrtAll) = DeclarationPotentialLineRecordSetTypes.dplrtAll Then
          .SetItem(DeclarationPotentialLineFields.dplfCdNumber, vFields)
          .SetItem(DeclarationPotentialLineFields.dplfContactNumber, vFields)
          .SetItem(DeclarationPotentialLineFields.dplfNetAmount, vFields)
          .SetItem(DeclarationPotentialLineFields.dplfAmountClaimed, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False) Implements ITaxClaimLine.Save
      SetValid(DeclarationPotentialLineFields.dplfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub InitFromUnclaimed(ByVal pEnv As CDBEnvironment, ByRef pClaim As DeclarationTaxClaim, ByRef pUnClaimed As DeclarationLinesUnclaimed, ByVal pTaxPercent As Integer)
      'pUnClaimed could be DeclarationTaxClaimLine or DeclarationPotentialLine
      Dim vAmount As Double

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvTaxPercent = pTaxPercent
      mvClassFields.Item(DeclarationPotentialLineFields.dplfCdNumber).Value = CStr(pUnClaimed.CdNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber).Value = pUnClaimed.DeclarationOrCovenantNumber
      mvClassFields.Item(DeclarationPotentialLineFields.dplfContactNumber).Value = CStr(pUnClaimed.ContactNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfBatchNumber).Value = CStr(pUnClaimed.BatchNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfTransactionNumber).Value = CStr(pUnClaimed.TransactionNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfLineNumber).Value = CStr(pUnClaimed.LineNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).Value = CStr(pUnClaimed.NetAmount)
      vAmount = CalculateAmountClaimed(pUnClaimed.NetAmount)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).Value = CStr(vAmount)
      pClaim.CalculatedTaxAmount = pClaim.CalculatedTaxAmount + vAmount
      pClaim.TotalNetAmount = pClaim.TotalNetAmount + pUnClaimed.NetAmount
      mvClaimNumber = pClaim.ClaimNumber
    End Sub

    Public Sub InitFromUnClaimedAdjustment(ByVal pEnv As CDBEnvironment, ByRef pClaim As DeclarationTaxClaim, ByRef pBTA As BatchTransactionAnalysis, ByVal pContactNumber As Integer, ByVal pTaxPercent As Integer)
      Dim vAmount As Double

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      mvTaxPercent = pTaxPercent
      mvClassFields.Item(DeclarationPotentialLineFields.dplfCdNumber).Value = pBTA.MemberNumber
      mvClassFields.Item(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber).Value = "D"
      mvClassFields.Item(DeclarationPotentialLineFields.dplfContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfBatchNumber).Value = CStr(pBTA.BatchNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfTransactionNumber).Value = CStr(pBTA.TransactionNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfLineNumber).Value = CStr(pBTA.LineNumber)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).Value = CStr(pBTA.Amount)
      vAmount = CalculateAmountClaimed(pBTA.Amount)
      mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).Value = CStr(vAmount)
      pClaim.CalculatedTaxAmount = pClaim.CalculatedTaxAmount + vAmount
      pClaim.TotalNetAmount = pClaim.TotalNetAmount + pBTA.Amount
      mvClaimNumber = pClaim.ClaimNumber
    End Sub

    Public Sub Delete() Implements ITaxClaimLine.Delete
      mvEnv.Connection.DeleteRecords(Me.TableName, mvClassFields.WhereFields)
    End Sub

    Public Sub UpdateNetAmount(ByVal pNewValue As Double, ByRef pClaim As DeclarationTaxClaim) Implements ITaxClaimLine.UpdateNetAmount
      'This is called from the Tax Claim process to change the claim line amount
      'Also updates the claim itself with the new value
      Dim vOldClaimAmount As Double
      Dim vOldNetAmount As Double

      vOldClaimAmount = mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).DoubleValue
      vOldNetAmount = mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).DoubleValue
      mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).Value = CStr(CalculateAmountClaimed(pNewValue))
      mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).Value = CStr(pNewValue)
      pClaim.CalculatedTaxAmount = FixTwoPlaces(pClaim.CalculatedTaxAmount - (vOldClaimAmount - mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).DoubleValue))
      pClaim.TotalNetAmount = pClaim.TotalNetAmount - (vOldNetAmount - NetAmount)
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
        AmountClaimed = mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationPotentialLineFields.dplfAmountClaimed).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property BatchNumber() As Integer Implements ITaxClaimLine.BatchNumber
      Get
        BatchNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CdNumber() As Integer Implements ITaxClaimLine.CdNumber
      Get
        CdNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfCdNumber).IntegerValue
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(DeclarationPotentialLineFields.dplfContactNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property DeclarationOrCovenantNumber() As String Implements ITaxClaimLine.DeclarationOrCovenantNumber
      Get
        DeclarationOrCovenantNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfDeclarationOrCovenantNumber).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer Implements ITaxClaimLine.LineNumber
      Get
        LineNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfLineNumber).IntegerValue
      End Get
    End Property

    Public Property NetAmount() As Double Implements ITaxClaimLine.NetAmount
      Get
        NetAmount = mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationPotentialLineFields.dplfNetAmount).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property TransactionNumber() As Integer Implements ITaxClaimLine.TransactionNumber
      Get
        TransactionNumber = mvClassFields.Item(DeclarationPotentialLineFields.dplfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ClaimNumber() As Integer Implements ITaxClaimLine.ClaimNumber
      Get
        ClaimNumber = mvClaimNumber
      End Get
    End Property

    Protected Overridable ReadOnly Property TableName As String
      Get
        Return "declaration_potential_lines"
      End Get
    End Property
  End Class
End Namespace
