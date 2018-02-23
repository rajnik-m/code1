Namespace Access
  Public Class GaSponsorshipLinesUnclaimed

    Public Enum GaSponsorshipLinesUnclaimedRecordSetTypes 'These are bit values
      gslurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GaSponsorshipLinesUnclaimedFields
      gslufAll = 0
      gslufBatchNumber
      gslufTransactionNumber
      gslufLineNumber
      gslufContactNumber
      gslufNetAmount
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
          .DatabaseTableName = "ga_sponsorship_lines_unclaimed"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufLineNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufBatchNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufTransactionNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufContactNumber).PrefixRequired = True
        mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufLineNumber).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As GaSponsorshipLinesUnclaimedFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GaSponsorshipLinesUnclaimedRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GaSponsorshipLinesUnclaimedRecordSetTypes.gslurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "gslu")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GaSponsorshipLinesUnclaimedRecordSetTypes.gslurtAll) & " FROM ga_sponsorship_lines_unclaimed gslu WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GaSponsorshipLinesUnclaimedRecordSetTypes.gslurtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GaSponsorshipLinesUnclaimedRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GaSponsorshipLinesUnclaimedFields.gslufBatchNumber, vFields)
        .SetItem(GaSponsorshipLinesUnclaimedFields.gslufTransactionNumber, vFields)
        .SetItem(GaSponsorshipLinesUnclaimedFields.gslufLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GaSponsorshipLinesUnclaimedRecordSetTypes.gslurtAll) = GaSponsorshipLinesUnclaimedRecordSetTypes.gslurtAll Then
          .SetItem(GaSponsorshipLinesUnclaimedFields.gslufContactNumber, vFields)
          .SetItem(GaSponsorshipLinesUnclaimedFields.gslufNetAmount, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GaSponsorshipLinesUnclaimedFields.gslufAll)
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

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = CInt(mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufContactNumber).Value)
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = CInt(mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufLineNumber).Value)
      End Get
    End Property

    Public ReadOnly Property NetAmount() As Double
      Get
        NetAmount = CDbl(mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufNetAmount).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(GaSponsorshipLinesUnclaimedFields.gslufTransactionNumber).Value)
      End Get
    End Property

    Public Sub Delete()
      mvEnv.Connection.DeleteRecords("ga_sponsorship_lines_unclaimed", mvClassFields.WhereFields)
    End Sub

    Public Sub CreateNewNegativeLines(ByVal pNewBatch As Integer, ByVal pNewTransaction As Integer, ByVal pNewline As Integer, ByVal pOrigBatch As Integer, ByVal pOrigTransaction As Integer, ByVal pOrigLine As Integer)
      Dim vSQL As String

      vSQL = "INSERT INTO ga_sponsorship_lines_unclaimed(batch_number,transaction_number,line_number,contact_number,net_amount)"
      vSQL = vSQL & " SELECT " & pNewBatch & "," & pNewTransaction & "," & pNewline & ", contact_number, (net_amount * -1)"
      vSQL = vSQL & " FROM ga_sponsorship_tax_claim_lines gstcl"
      vSQL = vSQL & " WHERE gstcl.batch_number = " & pOrigBatch & " AND gstcl.transaction_number = " & pOrigTransaction
      vSQL = vSQL & " AND line_number = " & pOrigLine

      mvEnv.Connection.ExecuteSQL(vSQL)
    End Sub
  End Class
End Namespace
