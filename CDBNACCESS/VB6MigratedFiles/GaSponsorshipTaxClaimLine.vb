Namespace Access
  Public Class GaSponsorshipTaxClaimLine

    Public Enum GaSponsorshipTaxClaimLineRecordSetTypes 'These are bit values
      gstclrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GaSponsorshipTaxClaimLineFields
      gstclfAll = 0
      gstclfClaimNumber
      gstclfBatchNumber
      gstclfTransactionNumber
      gstclfLineNumber
      gstclfContactNumber
      gstclfNetAmount
      gstclfAmountClaimed
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
          .DatabaseTableName = "ga_sponsorship_tax_claim_lines"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("claim_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfLineNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As GaSponsorshipTaxClaimLineFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GaSponsorshipTaxClaimLineRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GaSponsorshipTaxClaimLineRecordSetTypes.gstclrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "gstcl")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GaSponsorshipTaxClaimLineRecordSetTypes.gstclrtAll) & " FROM ga_sponsorship_tax_claim_lines gstcl WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GaSponsorshipTaxClaimLineRecordSetTypes.gstclrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GaSponsorshipTaxClaimLineRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GaSponsorshipTaxClaimLineFields.gstclfBatchNumber, vFields)
        .SetItem(GaSponsorshipTaxClaimLineFields.gstclfTransactionNumber, vFields)
        .SetItem(GaSponsorshipTaxClaimLineFields.gstclfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GaSponsorshipTaxClaimLineRecordSetTypes.gstclrtAll) = GaSponsorshipTaxClaimLineRecordSetTypes.gstclrtAll Then
          .SetItem(GaSponsorshipTaxClaimLineFields.gstclfClaimNumber, vFields)
          .SetItem(GaSponsorshipTaxClaimLineFields.gstclfContactNumber, vFields)
          .SetItem(GaSponsorshipTaxClaimLineFields.gstclfNetAmount, vFields)
          .SetItem(GaSponsorshipTaxClaimLineFields.gstclfAmountClaimed, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GaSponsorshipTaxClaimLineFields.gstclfAll)
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

    Public ReadOnly Property AmountClaimed() As Double
      Get
        AmountClaimed = CDbl(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfAmountClaimed).Value)
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ClaimNumber() As Integer
      Get
        ClaimNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfClaimNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfContactNumber).Value)
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfLineNumber).Value)
      End Get
    End Property

    Public ReadOnly Property NetAmount() As Double
      Get
        NetAmount = CDbl(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfNetAmount).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimLineFields.gstclfTransactionNumber).Value)
      End Get
    End Property

    Public Sub InitFromUnclaimed(ByVal pEnv As CDBEnvironment, ByVal pClaim As GaSponsorshipTaxClaim, ByVal pUnClaimed As GaSponsorshipPotentialLine, ByVal pTaxPercent As Integer)

      Init(pEnv)
      With mvClassFields
        .Item(GaSponsorshipTaxClaimLineFields.gstclfClaimNumber).Value = CStr(pClaim.ClaimNumber)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfBatchNumber).Value = CStr(pUnClaimed.BatchNumber)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfTransactionNumber).Value = CStr(pUnClaimed.TransactionNumber)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfLineNumber).Value = CStr(pUnClaimed.LineNumber)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfContactNumber).Value = CStr(pUnClaimed.ContactNumber)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfNetAmount).Value = CStr(pUnClaimed.NetAmount)
        .Item(GaSponsorshipTaxClaimLineFields.gstclfAmountClaimed).Value = CStr(pUnClaimed.AmountClaimed)
      End With
      pClaim.UpdateClaimAmounts(pUnClaimed.NetAmount, pUnClaimed.AmountClaimed, pTaxPercent)
    End Sub
  End Class
End Namespace
