

Namespace Access
  Public Class EventBookingTransaction

    Public Enum EventBookingTransactionRecordSetTypes 'These are bit values
      ebtrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventBookingTransactionFields
      ebtfAll = 0
      ebtfEventNumber
      ebtfBookingNumber
      ebtfBatchNumber
      ebtfTransactionNumber
      ebtfLineNumber
      ebtfAmendedBy
      ebtfAmendedOn
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
          .DatabaseTableName = "event_booking_transactions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("booking_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields(EventBookingTransactionFields.ebtfBookingNumber).SetPrimaryKeyOnly()
        mvClassFields(EventBookingTransactionFields.ebtfBatchNumber).SetPrimaryKeyOnly()
        mvClassFields(EventBookingTransactionFields.ebtfTransactionNumber).SetPrimaryKeyOnly()
        mvClassFields(EventBookingTransactionFields.ebtfLineNumber).SetPrimaryKeyOnly()

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventBookingTransactionFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventBookingTransactionFields.ebtfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventBookingTransactionFields.ebtfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventBookingTransactionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventBookingTransactionRecordSetTypes.ebtrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ebt")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pBookingNumber As Integer = 0, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If pBookingNumber > 0 And pBatchNumber > 0 And pTransactionNumber > 0 And pLineNumber > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(EventBookingTransactionRecordSetTypes.ebtrtAll) & " FROM event_booking_transactions ebt WHERE booking_number = " & pBookingNumber
        vSQL = vSQL & " AND batch_number = " & pBatchNumber & " And transaction_number = " & pTransactionNumber & " And line_number = " & pLineNumber
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventBookingTransactionRecordSetTypes.ebtrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventBookingTransactionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And EventBookingTransactionRecordSetTypes.ebtrtAll) = EventBookingTransactionRecordSetTypes.ebtrtAll Then
          .SetItem(EventBookingTransactionFields.ebtfEventNumber, vFields)
          .SetItem(EventBookingTransactionFields.ebtfBookingNumber, vFields)
          .SetItem(EventBookingTransactionFields.ebtfBatchNumber, vFields)
          .SetItem(EventBookingTransactionFields.ebtfTransactionNumber, vFields)
          .SetItem(EventBookingTransactionFields.ebtfLineNumber, vFields)
          .SetItem(EventBookingTransactionFields.ebtfAmendedBy, vFields)
          .SetItem(EventBookingTransactionFields.ebtfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(EventBookingTransactionFields.ebtfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Friend Sub Create(ByVal pParams As CDBParameters)
      With mvClassFields
        .Item(EventBookingTransactionFields.ebtfEventNumber).IntegerValue = pParams("EventNumber").IntegerValue
        .Item(EventBookingTransactionFields.ebtfBookingNumber).IntegerValue = pParams("BookingNumber").IntegerValue
      End With
      Update(pParams)
    End Sub

    Friend Sub Update(ByVal pParams As CDBParameters)
      With mvClassFields
        .Item(EventBookingTransactionFields.ebtfBatchNumber).IntegerValue = pParams("BatchNumber").IntegerValue
        .Item(EventBookingTransactionFields.ebtfTransactionNumber).IntegerValue = pParams("TransactionNumber").IntegerValue
        .Item(EventBookingTransactionFields.ebtfLineNumber).IntegerValue = pParams("LineNumber").IntegerValue
      End With
    End Sub

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pOldEBT As EventBookingTransaction, ByVal pNewBTA As BatchTransactionAnalysis)
      Init(pEnv)
      With mvClassFields
        .Item(EventBookingTransactionFields.ebtfEventNumber).IntegerValue = pOldEBT.EventNumber
        .Item(EventBookingTransactionFields.ebtfBookingNumber).IntegerValue = pOldEBT.BookingNumber
        .Item(EventBookingTransactionFields.ebtfBatchNumber).IntegerValue = pNewBTA.BatchNumber
        .Item(EventBookingTransactionFields.ebtfTransactionNumber).IntegerValue = pNewBTA.TransactionNumber
        .Item(EventBookingTransactionFields.ebtfLineNumber).IntegerValue = pNewBTA.LineNumber
      End With
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
        AmendedBy = mvClassFields.Item(EventBookingTransactionFields.ebtfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventBookingTransactionFields.ebtfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(EventBookingTransactionFields.ebtfBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property BookingNumber() As Integer
      Get
        BookingNumber = CInt(mvClassFields.Item(EventBookingTransactionFields.ebtfBookingNumber).Value)
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = CInt(mvClassFields.Item(EventBookingTransactionFields.ebtfEventNumber).Value)
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = CInt(mvClassFields.Item(EventBookingTransactionFields.ebtfLineNumber).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(EventBookingTransactionFields.ebtfTransactionNumber).Value)
      End Get
    End Property
  End Class
End Namespace
