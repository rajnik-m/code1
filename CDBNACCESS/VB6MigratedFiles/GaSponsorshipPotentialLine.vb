Namespace Access
  Public Class GaSponsorshipPotentialLine

    Public Enum GaSponsorshipPotentialLineRecordSetTypes 'These are bit values
      gsplrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GaSponsorshipPotentialLineFields
      gsplfAll = 0
      gsplfBatchNumber
      gsplfTransactionNumber
      gsplfLineNumber
      gsplfContactNumber
      gsplfNetAmount
      gsplfAmountClaimed
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
          .DatabaseTableName = "ga_sponsorship_potential_lines"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfLineNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfBatchNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfTransactionNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfLineNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfContactNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfNetAmount).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As GaSponsorshipPotentialLineFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GaSponsorshipPotentialLineRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GaSponsorshipPotentialLineRecordSetTypes.gsplrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "gspl")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GaSponsorshipPotentialLineRecordSetTypes.gsplrtAll) & " FROM ga_sponsorship_potential_lines gspl WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GaSponsorshipPotentialLineRecordSetTypes.gsplrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GaSponsorshipPotentialLineRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GaSponsorshipPotentialLineFields.gsplfBatchNumber, vFields)
        .SetItem(GaSponsorshipPotentialLineFields.gsplfTransactionNumber, vFields)
        .SetItem(GaSponsorshipPotentialLineFields.gsplfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GaSponsorshipPotentialLineRecordSetTypes.gsplrtAll) = GaSponsorshipPotentialLineRecordSetTypes.gsplrtAll Then
          .SetItem(GaSponsorshipPotentialLineFields.gsplfContactNumber, vFields)
          .SetItem(GaSponsorshipPotentialLineFields.gsplfNetAmount, vFields)
          .SetItem(GaSponsorshipPotentialLineFields.gsplfAmountClaimed, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GaSponsorshipPotentialLineFields.gsplfAll)
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
        AmountClaimed = CDbl(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfAmountClaimed).Value)
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = CInt(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfContactNumber).Value)
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = CInt(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfLineNumber).Value)
      End Get
    End Property

    Public ReadOnly Property NetAmount() As Double
      Get
        NetAmount = CDbl(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfNetAmount).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfTransactionNumber).Value)
      End Get
    End Property

    Public Sub InitFromUnclaimed(ByVal pEnv As CDBEnvironment, ByVal pClaim As GaSponsorshipTaxClaim, ByRef pUnClaimed As GaSponsorshipLinesUnclaimed, ByVal pTaxPercent As Integer)
      Dim vAmount As Double

      Init(pEnv)
      With mvClassFields
        .Item(GaSponsorshipPotentialLineFields.gsplfBatchNumber).Value = CStr(pUnClaimed.BatchNumber)
        .Item(GaSponsorshipPotentialLineFields.gsplfTransactionNumber).Value = CStr(pUnClaimed.TransactionNumber)
        .Item(GaSponsorshipPotentialLineFields.gsplfLineNumber).Value = CStr(pUnClaimed.LineNumber)
        .Item(GaSponsorshipPotentialLineFields.gsplfContactNumber).Value = CStr(pUnClaimed.ContactNumber)
        .Item(GaSponsorshipPotentialLineFields.gsplfNetAmount).Value = CStr(pUnClaimed.NetAmount)
      End With

      vAmount = CalculateAmountClaimed(pUnClaimed.NetAmount, pTaxPercent)
      mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfAmountClaimed).Value = CStr(vAmount)
      pClaim.UpdateClaimAmounts(mvClassFields.Item(GaSponsorshipPotentialLineFields.gsplfNetAmount).DoubleValue, vAmount, pTaxPercent)

    End Sub

    Private Function CalculateAmountClaimed(ByVal pNetAmount As Double, ByVal pTaxPercent As Integer) As Double
      Dim vAmount As Double

      vAmount = FixTwoPlaces((pNetAmount / (100 - pTaxPercent) * 100) - pNetAmount)
      CalculateAmountClaimed = vAmount
    End Function

    Public Sub Delete()
      mvEnv.Connection.DeleteRecords("ga_sponsorship_potential_lines", mvClassFields.WhereFields)
    End Sub
  End Class
End Namespace
