

Namespace Access
  Public Class OrderPaymentHistory

    Public Enum OrderPaymentHistoryRecordSetTypes 'These are bit values
      ophrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum OrderPaymentHistoryFields
      ophfAll = 0
      ophfBatchNumber
      ophfTransactionNumber
      ophflineNumber
      ophfPaymentNumber
      ophfOrderNumber
      ophfAmount
      ophfBalance
      ophfStatus
      ophfScheduledPaymentNumber
      ophfPosted
      ophfWriteOffLineAmount
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
          .DatabaseTableName = "order_payment_history"
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftInteger)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("payment_number", CDBField.FieldTypes.cftInteger)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("balance", CDBField.FieldTypes.cftNumeric)
          .Add("status")
          .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong)
          .Add("posted")
          .Add("write_off_line_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(OrderPaymentHistoryFields.ophfBatchNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfTransactionNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophflineNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfAmount).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfStatus).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfPosted).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfBalance).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfPaymentNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfOrderNumber).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfWriteOffLineAmount).PrefixRequired = True
        mvClassFields.Item(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments)
        mvClassFields.Item(OrderPaymentHistoryFields.ophfPosted).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataScheduledPayments)
        mvClassFields.Item(OrderPaymentHistoryFields.ophfWriteOffLineAmount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbBankTransactionsImport)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As OrderPaymentHistoryFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As OrderPaymentHistoryRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = OrderPaymentHistoryRecordSetTypes.ophrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "oph")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As OrderPaymentHistoryRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And OrderPaymentHistoryRecordSetTypes.ophrtAll) = OrderPaymentHistoryRecordSetTypes.ophrtAll Then
          .SetItem(OrderPaymentHistoryFields.ophfBatchNumber, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfTransactionNumber, vFields)
          .SetItem(OrderPaymentHistoryFields.ophflineNumber, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfPaymentNumber, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfOrderNumber, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfAmount, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfBalance, vFields)
          .SetItem(OrderPaymentHistoryFields.ophfStatus, vFields)
          .SetOptionalItem(OrderPaymentHistoryFields.ophfScheduledPaymentNumber, vFields)
          .SetOptionalItem(OrderPaymentHistoryFields.ophfPosted, vFields)
          .SetOptionalItem(OrderPaymentHistoryFields.ophfWriteOffLineAmount, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(OrderPaymentHistoryFields.ophfAll)
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

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(OrderPaymentHistoryFields.ophfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property Balance() As Double
      Get
        Balance = mvClassFields.Item(OrderPaymentHistoryFields.ophfBalance).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(OrderPaymentHistoryFields.ophfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(OrderPaymentHistoryFields.ophflineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OrderNumber() As Integer
      Get
        OrderNumber = mvClassFields.Item(OrderPaymentHistoryFields.ophfOrderNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PaymentNumber() As Integer
      Get
        PaymentNumber = mvClassFields.Item(OrderPaymentHistoryFields.ophfPaymentNumber).IntegerValue
      End Get
    End Property

    Public Property Status() As String
      Get
        Status = mvClassFields.Item(OrderPaymentHistoryFields.ophfStatus).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(OrderPaymentHistoryFields.ophfStatus).Value = Value
      End Set
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(OrderPaymentHistoryFields.ophfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentNumber() As String
      Get
        'Could be null
        ScheduledPaymentNumber = mvClassFields(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).Value
      End Get
    End Property

    Public ReadOnly Property Posted() As Boolean
      Get
        'Could be null
        Posted = (mvClassFields(OrderPaymentHistoryFields.ophfPosted).Value <> "N")
      End Get
    End Property

    Public ReadOnly Property WriteOffLineAmount() As Double
      Get
        Return mvClassFields(OrderPaymentHistoryFields.ophfWriteOffLineAmount).DoubleValue
      End Get
    End Property

    Public Sub SetWriteOffLineAmount(ByVal pWriteOffLineAmount As Double)
      mvClassFields(OrderPaymentHistoryFields.ophfWriteOffLineAmount).DoubleValue = pWriteOffLineAmount
    End Sub

    Public Sub SetValues(ByRef pBatchNo As Integer, ByRef pTransNo As Integer, ByRef pPaymentNo As Integer, ByRef pOrderNo As Integer, ByRef pAmount As Double, ByRef pLineNo As Integer, ByRef pBalance As Double, ByVal pScheduledPaymentNumber As Integer, Optional ByVal pPosted As Boolean = False, Optional ByVal pWriteOffLineAmount As Double = 0)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfBatchNumber).Value = CStr(pBatchNo)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfTransactionNumber).Value = CStr(pTransNo)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfPaymentNumber).Value = CStr(pPaymentNo)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfOrderNumber).Value = CStr(pOrderNo)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfAmount).Value = CStr(pAmount)
      mvClassFields.Item(OrderPaymentHistoryFields.ophflineNumber).Value = CStr(pLineNo)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfBalance).Value = CStr(pBalance)
      If pScheduledPaymentNumber > 0 Then
        mvClassFields.Item(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).Value = CStr(pScheduledPaymentNumber)
      End If
      mvClassFields.Item(OrderPaymentHistoryFields.ophfPosted).Bool = pPosted
      mvClassFields.Item(OrderPaymentHistoryFields.ophfWriteOffLineAmount).Value = pWriteOffLineAmount.ToString
    End Sub

    Public Sub SetPosted(ByVal pPosted As Boolean, Optional ByVal pBalance As Double = 0)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfPosted).Bool = pPosted
      mvClassFields.Item(OrderPaymentHistoryFields.ophfBalance).Value = CStr(pBalance)
    End Sub

    Public Sub Delete()
      If mvExisting Then mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Sub Reverse(ByVal pBatchNumber As Integer, ByVal pTransNumber As Integer, ByVal pLineNumber As Integer, ByVal pPaymentNumber As Integer, ByVal pScheduledPaymentNumber As Integer)
      'Reverse this OPH record
      With mvClassFields
        .ClearSetValues()
        .Item(OrderPaymentHistoryFields.ophfStatus).Value = ""
        .Item(OrderPaymentHistoryFields.ophfBatchNumber).IntegerValue = pBatchNumber
        .Item(OrderPaymentHistoryFields.ophfTransactionNumber).IntegerValue = pTransNumber
        .Item(OrderPaymentHistoryFields.ophflineNumber).IntegerValue = pLineNumber
        .Item(OrderPaymentHistoryFields.ophfPaymentNumber).IntegerValue = pPaymentNumber
        .Item(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).IntegerValue = pScheduledPaymentNumber
        .Item(OrderPaymentHistoryFields.ophfAmount).DoubleValue = (.Item(OrderPaymentHistoryFields.ophfAmount).DoubleValue * -1)
        .Item(OrderPaymentHistoryFields.ophfBalance).DoubleValue = (.Item(OrderPaymentHistoryFields.ophfBalance).DoubleValue * -1)
        .Item(OrderPaymentHistoryFields.ophfPosted).Bool = False
        If pScheduledPaymentNumber = 0 Then .Item(OrderPaymentHistoryFields.ophfScheduledPaymentNumber).Value = ""
        .Item(OrderPaymentHistoryFields.ophfWriteOffLineAmount).DoubleValue = (.Item(OrderPaymentHistoryFields.ophfWriteOffLineAmount).DoubleValue * -1)
      End With
      mvExisting = False
    End Sub

    Public Sub DeleteFromBatch(ByVal pWhereFields As CDBFields, ByVal pBatchType As Batch.BatchTypes)
      'Used when Deleting a Batch / BT / BTA to ensure that all OPH are deleted and OPS updated
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vOPH As New OrderPaymentHistory
      Dim vOPS As New OrderPaymentSchedule
      Dim vPP As New PaymentPlan
      Dim vRecordSet As CDBRecordSet

      pWhereFields.Add("posted", CDBField.FieldTypes.cftCharacter, "N")

      vOPH.Init(mvEnv)
      vPP.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE " & mvEnv.Connection.WhereClause(pWhereFields) & " ORDER BY line_number DESC")
      While vRecordSet.Fetch() = True
        vOPH.InitFromRecordSet(mvEnv, vRecordSet, OrderPaymentHistoryRecordSetTypes.ophrtAll)
        If vOPH.OrderNumber <> vPP.PlanNumber Then vPP.Init(mvEnv, (vOPH.OrderNumber))
        If vPP.PaymentNumber = vOPH.PaymentNumber Then
          If pBatchType = Batch.BatchTypes.DirectDebit And vPP.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes Then
            vBT = New BatchTransaction(mvEnv)
            vBT.Init((vOPH.BatchNumber), (vOPH.TransactionNumber))
            If vBT.TransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlFirstClaimTransactionType) Then
              vPP.DirectDebit.FirstClaim = True
              vPP.DirectDebit.SetAmended((TodaysDate()), mvEnv.User.Logname)
              vPP.DirectDebit.SaveChanges()
            End If
          End If
          vPP.PaymentNumber = vPP.PaymentNumber - 1
          vPP.SaveChanges()
        End If
        If vOPH.ScheduledPaymentNumber.Length > 0 Then
          vOPS.Init(mvEnv, CInt(vOPH.ScheduledPaymentNumber))
          vOPS.SetUnProcessedPayment(False, (vOPH.Amount * -1))
          vOPS.Save()
        End If
        vOPH.Delete()
      End While
      vRecordSet.CloseRecordSet()

      pWhereFields.Remove((pWhereFields.Count))
    End Sub

    Friend ReadOnly Property Key() As String
      Get
        Return mvClassFields.Item(OrderPaymentHistoryFields.ophfOrderNumber).Value & "|" & mvClassFields.Item(OrderPaymentHistoryFields.ophfPaymentNumber).Value
      End Get
    End Property

    Friend Sub IncreasePaymentAmount(ByVal pAmount As Double)
      mvClassFields.Item(OrderPaymentHistoryFields.ophfAmount).DoubleValue = FixTwoPlaces(Amount + pAmount)
    End Sub
  End Class
End Namespace
