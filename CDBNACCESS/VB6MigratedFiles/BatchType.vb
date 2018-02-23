Namespace Access
  Public Class BatchTypeData

    Public Enum BatchTypeRecordSetTypes 'These are bit values
      btyrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BatchTypeFields
      btfAll = 0
      btfBatchType
      btfBatchTypeDesc
      btfDefaultBankAccount
      btfPrintChequeList
      btfAmendedBy
      btfAmendedOn
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
          .DatabaseTableName = "batch_types"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("batch_type")
          .Add("batch_type_desc")
          .Add("default_bank_account")
          .Add("print_cheque_list")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(BatchTypeFields.btfBatchType).SetPrimaryKeyOnly()
        mvClassFields.Item(BatchTypeFields.btfBatchType).PrefixRequired = True
        mvClassFields.Item(BatchTypeFields.btfPrintChequeList).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPrintChequeList)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As BatchTypeFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(BatchTypeFields.btfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(BatchTypeFields.btfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BatchTypeRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BatchTypeRecordSetTypes.btyrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bty")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchType As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pBatchType) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(BatchTypeRecordSetTypes.btyrtAll) & " FROM batch_types bty WHERE batch_type = '" & pBatchType & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BatchTypeRecordSetTypes.btyrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BatchTypeRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BatchTypeFields.btfBatchType, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BatchTypeRecordSetTypes.btyrtAll) = BatchTypeRecordSetTypes.btyrtAll Then
          .SetItem(BatchTypeFields.btfBatchTypeDesc, vFields)
          .SetItem(BatchTypeFields.btfDefaultBankAccount, vFields)
          .SetOptionalItem(BatchTypeFields.btfPrintChequeList, vFields)
          .SetItem(BatchTypeFields.btfAmendedBy, vFields)
          .SetItem(BatchTypeFields.btfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(BatchTypeFields.btfAll)
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

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(BatchTypeFields.btfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(BatchTypeFields.btfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchTypeCode() As String
      Get
        BatchTypeCode = mvClassFields.Item(BatchTypeFields.btfBatchType).Value
      End Get
    End Property
    Public ReadOnly Property PrintChequeList() As Boolean
      Get
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPrintChequeList) Then
          PrintChequeList = mvClassFields.Item(BatchTypeFields.btfPrintChequeList).Bool
        Else
          If BatchTypeCode = Batch.GetBatchTypeCode(Batch.BatchTypes.Cash) OrElse BatchTypeCode = Batch.GetBatchTypeCode(Batch.BatchTypes.CashWithInvoice) Then
            PrintChequeList = True
          Else
            PrintChequeList = False
          End If
        End If
      End Get
    End Property
    Public ReadOnly Property BatchTypeDesc() As String
      Get
        BatchTypeDesc = mvClassFields.Item(BatchTypeFields.btfBatchTypeDesc).Value
      End Get
    End Property

    Public ReadOnly Property DefaultBankAccount() As String
      Get
        DefaultBankAccount = mvClassFields.Item(BatchTypeFields.btfDefaultBankAccount).Value
      End Get
    End Property
  End Class
End Namespace
