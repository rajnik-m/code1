Namespace Access
  Public Class FundraisingEventAnalysis

    Public Enum FundraisingEventAnalysisRecordSetTypes 'These are bit values
      feartAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum FundraisingEventAnalysisFields
      feafAll = 0
      feafContactFundraisingNumber
      feafBatchNumber
      feafTransactionNumber
      feafLineNumber
      feafAmendedBy
      feafAmendedOn
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
          .DatabaseTableName = "fundraising_event_analysis"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_fundraising_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As FundraisingEventAnalysisFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(FundraisingEventAnalysisFields.feafAmendedOn).Value = TodaysDate()
      mvClassFields.Item(FundraisingEventAnalysisFields.feafAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As FundraisingEventAnalysisRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = FundraisingEventAnalysisRecordSetTypes.feartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "fea")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Create(ByRef pContactFundRaisingNumber As Integer, ByRef pBatchNumber As Integer, ByRef pTransactionNumber As Integer, ByRef pLineNumber As Integer)
      mvClassFields(FundraisingEventAnalysisFields.feafContactFundraisingNumber).IntegerValue = pContactFundRaisingNumber
      mvClassFields(FundraisingEventAnalysisFields.feafBatchNumber).IntegerValue = pBatchNumber
      mvClassFields(FundraisingEventAnalysisFields.feafTransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields(FundraisingEventAnalysisFields.feafLineNumber).IntegerValue = pLineNumber
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As FundraisingEventAnalysisRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And FundraisingEventAnalysisRecordSetTypes.feartAll) = FundraisingEventAnalysisRecordSetTypes.feartAll Then
          .SetItem(FundraisingEventAnalysisFields.feafContactFundraisingNumber, vFields)
          .SetItem(FundraisingEventAnalysisFields.feafBatchNumber, vFields)
          .SetItem(FundraisingEventAnalysisFields.feafTransactionNumber, vFields)
          .SetItem(FundraisingEventAnalysisFields.feafLineNumber, vFields)
          .SetItem(FundraisingEventAnalysisFields.feafAmendedBy, vFields)
          .SetItem(FundraisingEventAnalysisFields.feafAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(FundraisingEventAnalysisFields.feafAll)
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
        AmendedBy = mvClassFields.Item(FundraisingEventAnalysisFields.feafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(FundraisingEventAnalysisFields.feafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(FundraisingEventAnalysisFields.feafBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ContactFundraisingNumber() As Integer
      Get
        ContactFundraisingNumber = CInt(mvClassFields.Item(FundraisingEventAnalysisFields.feafContactFundraisingNumber).Value)
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = CInt(mvClassFields.Item(FundraisingEventAnalysisFields.feafLineNumber).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(FundraisingEventAnalysisFields.feafTransactionNumber).Value)
      End Get
    End Property
  End Class
End Namespace
