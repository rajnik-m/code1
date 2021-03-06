Namespace Access

  Public Class LegacyBequestReceipt
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum LegacyBequestReceiptFields
      AllFields = 0
      LegacyNumber
      BequestNumber
      ReceiptNumber
      Amount
      BatchNumber
      TransactionNumber
      LineNumber
      DateReceived
      Notes
      Status
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("legacy_number", CDBField.FieldTypes.cftLong)
        .Add("bequest_number", CDBField.FieldTypes.cftLong)
        .Add("receipt_number", CDBField.FieldTypes.cftLong)
        .Add("amount", CDBField.FieldTypes.cftNumeric)
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("line_number", CDBField.FieldTypes.cftInteger)
        .Add("date_received", CDBField.FieldTypes.cftDate)
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("status")

        .Item(LegacyBequestReceiptFields.ReceiptNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "lbr"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "legacy_bequest_receipts"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property LegacyNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.LegacyNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property BequestNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.BequestNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ReceiptNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.ReceiptNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Amount() As Double
      Get
        Return mvClassFields(LegacyBequestReceiptFields.Amount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LineNumber() As Integer
      Get
        Return mvClassFields(LegacyBequestReceiptFields.LineNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(LegacyBequestReceiptFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(LegacyBequestReceiptFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property DateReceived() As String
      Get
        Return mvClassFields(LegacyBequestReceiptFields.DateReceived).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(LegacyBequestReceiptFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property Status() As String
      Get
        Return mvClassFields(LegacyBequestReceiptFields.Status).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Public Sub InitFromBequest(ByVal pBequest As LegacyBequest, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, ByVal pAmount As Double, ByVal pDateReceived As String, ByVal pNotes As String)
      Init()
      With mvClassFields
        .Item(LegacyBequestReceiptFields.LegacyNumber).IntegerValue = pBequest.LegacyNumber
        .Item(LegacyBequestReceiptFields.BequestNumber).IntegerValue = pBequest.BequestNumber
        .Item(LegacyBequestReceiptFields.ReceiptNumber).IntegerValue = mvEnv.GetControlNumber("LR")
        .Item(LegacyBequestReceiptFields.BatchNumber).IntegerValue = pBatchNumber
        .Item(LegacyBequestReceiptFields.TransactionNumber).IntegerValue = pTransactionNumber
        .Item(LegacyBequestReceiptFields.LineNumber).IntegerValue = pLineNumber
        .Item(LegacyBequestReceiptFields.Amount).DoubleValue = pAmount
        .Item(LegacyBequestReceiptFields.DateReceived).Value = pDateReceived
        .Item(LegacyBequestReceiptFields.Notes).Value = pNotes
      End With
    End Sub

#End Region

  End Class
End Namespace
