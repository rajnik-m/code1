Namespace Access

  Partial Public Class CreditSale
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum CreditSaleFields
      AllFields = 0
      ContactNumber
      AddressNumber
      BatchNumber
      TransactionNumber
      StockSale
      AddressTo
      SalesLedgerAccount
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("stock_sale")
        .Add("address_to")
        .Add("sales_ledger_account")

        .Item(CreditSaleFields.BatchNumber).PrimaryKey = True

        .Item(CreditSaleFields.TransactionNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cs"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "credit_sales"
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
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(CreditSaleFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(CreditSaleFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property BatchNumber() As Integer
      Get
        Return mvClassFields(CreditSaleFields.BatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return mvClassFields(CreditSaleFields.TransactionNumber).IntegerValue
      End Get
    End Property
    Public Property StockSale() As Boolean
      Get
        Return mvClassFields(CreditSaleFields.StockSale).Bool
      End Get
      Set(ByVal value As Boolean)
        mvClassFields(CreditSaleFields.StockSale).Bool = value
      End Set
    End Property
    Public ReadOnly Property AddressTo() As String
      Get
        Return mvClassFields(CreditSaleFields.AddressTo).Value
      End Get
    End Property
    Public ReadOnly Property SalesLedgerAccount() As String
      Get
        Return mvClassFields(CreditSaleFields.SalesLedgerAccount).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Overloads Sub Init(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      If pBatchNumber > 0 AndAlso pTransactionNumber > 0 Then
        CheckClassFields()
        Dim vWhereFields As New CDBFields
        vWhereFields.Add(mvClassFields(CreditSaleFields.BatchNumber).Name, pBatchNumber)
        vWhereFields.Add(mvClassFields(CreditSaleFields.TransactionNumber).Name, pTransactionNumber)
        InitWithPrimaryKey(vWhereFields)
      Else
        Init()
      End If
    End Sub

#End Region

  End Class
End Namespace