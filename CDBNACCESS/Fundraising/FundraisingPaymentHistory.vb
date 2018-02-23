Namespace Access

  Public Class FundraisingPaymentHistory
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum FundraisingPaymentHistoryFields
      AllFields = 0
      ScheduledPaymentNumber
      BatchNumber
      TransactionNumber
      LineNumber
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong)
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("line_number", CDBField.FieldTypes.cftInteger)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "fph"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "fundraising_payment_history"
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
    Public ReadOnly Property ScheduledPaymentNumber() As Integer
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.ScheduledPaymentNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LineNumber() As Integer
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.LineNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(FundraisingPaymentHistoryFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"
    Public Sub CreateNewLink(ByVal pScheduledPaymentNo As Integer, ByVal pBatchNo As Integer, ByVal pTransactionNo As Integer, ByVal pLineNo As Integer, Optional ByVal pOrigBatchNo As Integer = 0, Optional ByVal pOrigTransNo As Integer = 0)
      If pOrigBatchNo > 0 And pOrigTransNo > 0 Then
        DeleteLinks(pOrigBatchNo, pOrigTransNo)
      End If
      With mvClassFields
        .Item(FundraisingPaymentHistoryFields.ScheduledPaymentNumber).IntegerValue = pScheduledPaymentNo
        .Item(FundraisingPaymentHistoryFields.BatchNumber).IntegerValue = pBatchNo
        .Item(FundraisingPaymentHistoryFields.TransactionNumber).IntegerValue = pTransactionNo
        .Item(FundraisingPaymentHistoryFields.LineNumber).IntegerValue = pLineNo
      End With
      Save()
    End Sub

    ''' <summary>Confirming a provisional transaction may have been configured so that the link to a Fundraising Request payment is no longer required. So this will remove those unwanted links.</summary>
    ''' <param name="pBatchNumber">Provisional Batch Number</param>
    ''' <param name="pTransactionNumber">Provisional Transaction Number</param>
    ''' <remarks>This will delete all the records for the Batch Number and Transaction Number.</remarks>
    Friend Sub DeleteLinks(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      If pBatchNumber > 0 AndAlso pTransactionNumber > 0 Then
        Dim vWhereFields As New CDBFields({New CDBField("batch_number", pBatchNumber), New CDBField("transaction_number", pTransactionNumber)})
        mvEnv.Connection.DeleteRecords("fundraising_payment_history", vWhereFields, False)
      End If
    End Sub

#End Region

  End Class
End Namespace