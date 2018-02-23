

Namespace Access
  Public Class DeclarationLinesUnclaimed

    Public Enum DeclarationLinesUnclaimedRecordSetTypes 'These are bit values
      dlurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DeclarationLinesUnclaimedFields
      dlufAll = 0
      dlufCdNumber
      dlufContactNumber
      dlufBatchNumber
      dlufTransactionNumber
      dlufLineNumber
      dlufDeclarationOrCovenantNumber
      dlufNetAmount
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
          .DatabaseTableName = "declaration_lines_unclaimed"
          .Add("cd_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("declaration_or_covenant_number")
          .Add("net_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufBatchNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufLineNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufContactNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufBatchNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufTransactionNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufLineNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufCdNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber).PrefixRequired = True
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufNetAmount).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvGiftAidDeclaration = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub
    Private Sub SetValid(ByRef pField As DeclarationLinesUnclaimedFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Sub Delete()
      mvEnv.Connection.DeleteRecords("declaration_lines_unclaimed", mvClassFields.WhereFields)
    End Sub

    Public Function GetRecordSetFields(ByVal pRSType As DeclarationLinesUnclaimedRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DeclarationLinesUnclaimedRecordSetTypes.dlurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dlu")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0, Optional ByRef pDecOrCovNumber As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pBatchNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DeclarationLinesUnclaimedRecordSetTypes.dlurtAll) & " FROM declaration_lines_unclaimed dlu WHERE batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber & " AND declaration_or_covenant_number = '" & pDecOrCovNumber & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DeclarationLinesUnclaimedRecordSetTypes.dlurtAll)
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
    Public Sub InitForNew(ByRef pEnv As CDBEnvironment, ByRef pCdNumber As Integer, ByRef pContactNumber As Integer, ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, ByRef pLineNumber As Integer, ByRef pType As String, ByRef pNetAmount As Double)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()

      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufCdNumber).Value = CStr(pCdNumber)
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufBatchNumber).Value = CStr(pBatchNumber)
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufTransactionNumber).Value = CStr(pTransactionNumber)
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufLineNumber).Value = CStr(pLineNumber)
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber).Value = pType
      mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufNetAmount).Value = CStr(pNetAmount)
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DeclarationLinesUnclaimedRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DeclarationLinesUnclaimedFields.dlufBatchNumber, vFields)
        .SetItem(DeclarationLinesUnclaimedFields.dlufTransactionNumber, vFields)
        .SetItem(DeclarationLinesUnclaimedFields.dlufLineNumber, vFields)
        .SetItem(DeclarationLinesUnclaimedFields.dlufNetAmount, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DeclarationLinesUnclaimedRecordSetTypes.dlurtAll) = DeclarationLinesUnclaimedRecordSetTypes.dlurtAll Then
          .SetItem(DeclarationLinesUnclaimedFields.dlufCdNumber, vFields)
          .SetItem(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber, vFields)
          .SetItem(DeclarationLinesUnclaimedFields.dlufContactNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(DeclarationLinesUnclaimedFields.dlufAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Public Sub InitFromOverPaidClaimLine(ByVal pEnv As CDBEnvironment, ByVal pClaimLine As ITaxClaimLine, ByVal pDeclaration As GiftAidDeclaration, ByVal pContactNumber As Integer, ByVal pAmount As Double)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      'First check not already existing
      Init(mvEnv, pClaimLine.BatchNumber, pClaimLine.TransactionNumber, pClaimLine.LineNumber, "D")
      If Existing = False Then
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufCdNumber).Value = CStr(pDeclaration.DeclarationNumber)
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufContactNumber).Value = CStr(pContactNumber)
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufBatchNumber).IntegerValue = pClaimLine.BatchNumber
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufTransactionNumber).IntegerValue = pClaimLine.TransactionNumber
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufLineNumber).IntegerValue = pClaimLine.LineNumber
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber).Value = "D"
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufNetAmount).Value = CStr(pAmount)
      End If
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
        BatchNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CdNumber() As Integer
      Get
        CdNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufCdNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DeclarationOrCovenantNumber() As String
      Get
        DeclarationOrCovenantNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufDeclarationOrCovenantNumber).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufLineNumber).IntegerValue
      End Get
    End Property

    Public Property NetAmount() As Double
      Get
        NetAmount = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufNetAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufNetAmount).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(DeclarationLinesUnclaimedFields.dlufTransactionNumber).IntegerValue
      End Get
    End Property

    Public Sub CreateNewNegativeLines(ByVal pNewBatch As Integer, ByVal pNewTransaction As Integer, ByVal pNewline As Integer, ByVal pOrigBatch As Integer, ByVal pOrigTransaction As Integer, Optional ByVal pOrigLine As Integer = 0)
      Dim vSQL As String
      Dim vIndex As Integer
      Dim vType As String
      Dim vWhereFields As New CDBFields
      Dim vDLU As New DeclarationLinesUnclaimed
      Dim vRS As CDBRecordSet
      Dim vDT As New CDBDataTable
      Dim vDR As CDBDataRow
      Dim vAdjustmentTotal As Double
      Dim vAdjustmentCount As Integer
      Dim vFirstRecord As Boolean
      Dim vAdjustmentReferences As Collection
      Dim vDeclarationNumber As Integer

      For vIndex = 0 To 1
        vType = If(vIndex = 0, "D", "C")
        vSQL = "SELECT cd_number,contact_number,net_amount,line_number"
        vSQL = vSQL & " FROM declaration_tax_claim_lines dtcl, " & If(vType = "D", "gift_aid_declarations gad", "covenants c")
        vSQL = vSQL & " WHERE dtcl.batch_number = " & pOrigBatch & " AND dtcl.transaction_number = " & pOrigTransaction
        If pOrigLine > 0 Then vSQL = vSQL & " AND line_number = " & pOrigLine
        vSQL = vSQL & " AND declaration_or_covenant_number = '" & vType & "' AND dtcl.cd_number = " & If(vType = "D", "gad.declaration_number", "c.covenant_number")

        vRS = mvEnv.Connection.GetRecordSet(vSQL)

        vFirstRecord = True
        While vRS.Fetch() = True
          vDeclarationNumber = vRS.Fields(1).IntegerValue
          If vFirstRecord = True And vType = "D" Then
            'BR 11524: We are looking at the first BTA for this Declaration, fill Data Table with any
            'GC adjustment records for this Contact.
            With vWhereFields
              .Clear()
              .Add("b.batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment), CDBField.FieldWhereOperators.fwoEqual)
              .Add("bt.batch_number", CDBField.FieldTypes.cftLong, "b.batch_number", CDBField.FieldWhereOperators.fwoEqual)
              .Add("bt.contact_number", CDBField.FieldTypes.cftLong, vRS.Fields(2).IntegerValue)
              .Add("bta.batch_number", CDBField.FieldTypes.cftLong, "bt.batch_number", CDBField.FieldWhereOperators.fwoEqual)
              .Add("bta.transaction_number", CDBField.FieldTypes.cftLong, "bt.transaction_number", CDBField.FieldWhereOperators.fwoEqual)
            End With
            vSQL = "SELECT bta.notes,bta.amount,bta.batch_number,bta.transaction_number,bta.line_number,bta.member_number FROM batches b, batch_transactions bt, batch_transaction_analysis bta WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY bta.batch_number, bta.transaction_number, bta.line_number"
            vDT.FillFromSQLDONOTUSE(mvEnv, vSQL, "Notes,Amount,Batch_Number,Transaction_Number,Line_Number,Member_Number")
            vFirstRecord = False
          End If

          vAdjustmentTotal = 0
          vAdjustmentCount = 0
          vAdjustmentReferences = New Collection

          For Each vDR In vDT.Rows
            If vDR.Item("Notes") = pOrigBatch & "/" & pOrigTransaction & "/" & vRS.Fields(4).IntegerValue Then
              vAdjustmentReferences.Add(vDR.Item("Batch_Number") & "/" & vDR.Item("Transaction_Number") & "/" & vDR.Item("Line_Number"))
              vAdjustmentCount = vAdjustmentCount + 1
              vAdjustmentTotal = vAdjustmentTotal + vDR.DoubleItem("Amount")
            End If
          Next vDR

          While vAdjustmentReferences.Count() > 0
            'Iterative loop to add totals on any Adjustments of Adjustments and add back in if they were in turn adjusted!
            For Each vDR In vDT.Rows
              If vDR.Item("Notes") = CStr(vAdjustmentReferences.Item(1)) Then
                vAdjustmentReferences.Add(vDR.Item("Batch_Number") & "/" & vDR.Item("Transaction_Number") & "/" & vDR.Item("Line_Number"))
                vAdjustmentCount = vAdjustmentCount + 1
                vAdjustmentTotal = vAdjustmentTotal + vDR.DoubleItem("Amount")
                'Override Declaration Number with that of Adjustment Transaction as it is the Declaration that
                'replaced the Original Declaration.
                vDeclarationNumber = IntegerValue(vDR.Item("Member_Number"))
              End If
            Next vDR
            vAdjustmentReferences.Remove((1))
          End While

          If vAdjustmentCount = 0 Or vAdjustmentTotal >= 0 Then
            'Either we have not previously adjusted this payment OR we have adjusted it but
            'balance of adjustments are positive so create a negative (reversal)
            With vDLU
              .InitForNew(mvEnv, vDeclarationNumber, (vRS.Fields(2).IntegerValue), pNewBatch, pNewTransaction, pNewline, vType, (vRS.Fields(3).DoubleValue * -1))
              .Save()
            End With
          End If
        End While
        vRS.CloseRecordSet()
      Next
    End Sub

    Private mvGiftAidDeclaration As GiftAidDeclaration = Nothing
    Public ReadOnly Property GiftAidDeclaration As GiftAidDeclaration
      Get
        If mvGiftAidDeclaration Is Nothing Then
          Dim vDeclaration As New GiftAidDeclaration
          vDeclaration.Init(mvEnv)
          Dim vSql As New SQLStatement(mvEnv.Connection,
                                       vDeclaration.GetRecordSetFields(Access.GiftAidDeclaration.GiftAidDeclarationRecordSetTypes.gadrtAll),
                                       vDeclaration.TableName,
                                       New CDBFields({New CDBField("declaration_number", Me.CdNumber)}))
          Dim vRs As CDBRecordSet = vSql.GetRecordSet
          If vRs.Fetch Then
            vDeclaration.InitFromRecordSet(mvEnv, vRs, Access.GiftAidDeclaration.GiftAidDeclarationRecordSetTypes.gadrtAll)
            mvGiftAidDeclaration = vDeclaration
          End If
          vRs.CloseRecordSet()
        End If
        Return mvGiftAidDeclaration
      End Get
    End Property

  End Class
End Namespace
