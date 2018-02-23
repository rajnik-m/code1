

Namespace Access
  Public Class CollectionPayment

    Public Enum CollectionPaymentRecordSetTypes 'These are bit values
      cpyrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionPaymentFields
      cpfAll = 0
      cpfCollectionPaymentNumber
      cpfCollectionNumber
      cpfCollectionPISNumber
      cpfBatchNumber
      cpfTransactionNumber
      cpfLineNumber
      cpfCollectionBoxNumber
      cpfAmount
      cpfAmendedBy
      cpfAmendedOn
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
          .DatabaseTableName = "collection_payments"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_payment_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("collection_pis_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("collection_box_number", CDBField.FieldTypes.cftLong)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by", CDBField.FieldTypes.cftCharacter)
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CollectionPaymentFields.cpfCollectionPaymentNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionPaymentFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields(CollectionPaymentFields.cpfCollectionPaymentNumber).IntegerValue = 0 Then mvClassFields(CollectionPaymentFields.cpfCollectionPaymentNumber).IntegerValue = mvEnv.GetControlNumber("CY")
      mvClassFields.Item(CollectionPaymentFields.cpfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectionPaymentFields.cpfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionPaymentRecordSetTypes) As String
      Dim vFields As String = ""

      'Modify below to add each recordset type as required
      If pRSType = CollectionPaymentRecordSetTypes.cpyrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cp")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionPaymentNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionPaymentNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionPaymentRecordSetTypes.cpyrtAll) & " FROM collection_payments cp WHERE collection_payment_number = " & pCollectionPaymentNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionPaymentRecordSetTypes.cpyrtAll)
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

    Public Sub InitFromBatch(ByVal pEnv As CDBEnvironment, ByVal pCollectionNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      vSQL = "SELECT " & GetRecordSetFields(CollectionPaymentRecordSetTypes.cpyrtAll) & " FROM collection_payments cp WHERE collection_number = " & pCollectionNumber
      vSQL = vSQL & " AND batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND line_number = " & pLineNumber
      vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CollectionPaymentRecordSetTypes.cpyrtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromCollPISAndBTA(ByVal pEnv As CDBEnvironment, ByRef pCollPIS As CollectionPIS, ByRef pBTA As BatchTransactionAnalysis, Optional ByRef pCBNumber As Integer = 0)

      mvEnv = pEnv
      'Set mvBatchTransaction = pBatchTransaction
      InitClassFields()
      SetDefaults()
      mvExisting = False
      With pBTA
        mvClassFields.Item(CollectionPaymentFields.cpfBatchNumber).Value = CStr(.BatchNumber)
        mvClassFields.Item(CollectionPaymentFields.cpfTransactionNumber).Value = CStr(.TransactionNumber)
        mvClassFields.Item(CollectionPaymentFields.cpfLineNumber).Value = CStr(.TransactionNumber)
        mvClassFields.Item(CollectionPaymentFields.cpfAmount).Value = CStr(.Amount)
      End With
      With pCollPIS
        mvClassFields.Item(CollectionPaymentFields.cpfCollectionNumber).Value = CStr(.CollectionNumber)
        mvClassFields.Item(CollectionPaymentFields.cpfCollectionPISNumber).Value = CStr(.CollectionPisNumber)
      End With
      If pCBNumber > 0 Then
        mvClassFields.Item(CollectionPaymentFields.cpfCollectionBoxNumber).IntegerValue = pCBNumber
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionPaymentRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionPaymentFields.cpfCollectionPaymentNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionPaymentRecordSetTypes.cpyrtAll) = CollectionPaymentRecordSetTypes.cpyrtAll Then
          .SetItem(CollectionPaymentFields.cpfCollectionNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfCollectionPISNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfBatchNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfTransactionNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfLineNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfCollectionBoxNumber, vFields)
          .SetItem(CollectionPaymentFields.cpfAmount, vFields)
          .SetItem(CollectionPaymentFields.cpfAmendedBy, vFields)
          .SetItem(CollectionPaymentFields.cpfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectionPaymentFields.cpfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pCollectionNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pAmount As Double, Optional ByVal pCollectionPISNumber As Integer = 0, Optional ByVal pCollectionBoxNumber As Integer = 0)
      With mvClassFields
        .Item(CollectionPaymentFields.cpfCollectionNumber).IntegerValue = pCollectionNumber
        .Item(CollectionPaymentFields.cpfBatchNumber).IntegerValue = pBatchNumber
        .Item(CollectionPaymentFields.cpfTransactionNumber).IntegerValue = pTransactionNumber
        .Item(CollectionPaymentFields.cpfLineNumber).IntegerValue = pLineNumber
        .Item(CollectionPaymentFields.cpfAmount).DoubleValue = pAmount
        If pCollectionPISNumber > 0 Then .Item(CollectionPaymentFields.cpfCollectionPISNumber).IntegerValue = pCollectionPISNumber
        If pCollectionBoxNumber > 0 Then .Item(CollectionPaymentFields.cpfCollectionBoxNumber).IntegerValue = pCollectionBoxNumber
      End With
    End Sub

    Public Sub Update(ByVal pCollectionNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, Optional ByVal pCollectionPISNumber As Integer = 0)
      With mvClassFields
        .Item(CollectionPaymentFields.cpfCollectionNumber).IntegerValue = pCollectionNumber
        .Item(CollectionPaymentFields.cpfBatchNumber).IntegerValue = pBatchNumber
        .Item(CollectionPaymentFields.cpfTransactionNumber).IntegerValue = pTransactionNumber
        .Item(CollectionPaymentFields.cpfLineNumber).IntegerValue = pLineNumber
        If .Item(CollectionPaymentFields.cpfCollectionPISNumber).IntegerValue <> pCollectionPISNumber Then
          'Field is optional
          .Item(CollectionPaymentFields.cpfCollectionPISNumber).Value = ""
          If pCollectionPISNumber > 0 Then .Item(CollectionPaymentFields.cpfCollectionPISNumber).IntegerValue = pCollectionPISNumber
        End If
      End With
    End Sub

    Public Sub Delete()
      mvEnv.Connection.DeleteRecords("collection_payments", mvClassFields.WhereFields)
    End Sub

    Public Sub Reverse(ByVal pNewBatchNumber As Integer, ByVal pNewTransactionNumber As Integer, Optional ByVal pNewLineNumber As Integer = 0)
      'Reverse the existing CollectionPayment
      'If pNewLineNumber set then new Payment has this line number, otherwise original line number is kept
      With mvClassFields
        .ClearSetValues()
        .Item(CollectionPaymentFields.cpfCollectionPaymentNumber).IntegerValue = 0
        .Item(CollectionPaymentFields.cpfBatchNumber).IntegerValue = pNewBatchNumber
        .Item(CollectionPaymentFields.cpfTransactionNumber).IntegerValue = pNewTransactionNumber
        .Item(CollectionPaymentFields.cpfAmount).DoubleValue = (.Item(CollectionPaymentFields.cpfAmount).DoubleValue * -1)
        If pNewLineNumber > 0 Then .Item(CollectionPaymentFields.cpfLineNumber).IntegerValue = pNewLineNumber
      End With
      mvExisting = False
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
        AmendedBy = mvClassFields.Item(CollectionPaymentFields.cpfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionPaymentFields.cpfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(CollectionPaymentFields.cpfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(CollectionPaymentFields.cpfBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = CInt(mvClassFields.Item(CollectionPaymentFields.cpfCollectionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectionBoxNumber() As Integer
      Get
        CollectionBoxNumber = mvClassFields.Item(CollectionPaymentFields.cpfCollectionBoxNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPaymentNumber() As Integer
      Get
        CollectionPaymentNumber = mvClassFields.Item(CollectionPaymentFields.cpfCollectionPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPisNumber() As Integer
      Get
        CollectionPisNumber = mvClassFields.Item(CollectionPaymentFields.cpfCollectionPISNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(CollectionPaymentFields.cpfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(CollectionPaymentFields.cpfTransactionNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
