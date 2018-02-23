Namespace Access

  Partial Public Class CreditSale

    Public Overloads Sub Clone(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      mvClassFields.ClearSetValues()
      mvClassFields.Item(CreditSaleFields.BatchNumber).IntegerValue = pBatchNumber
      mvClassFields.Item(CreditSaleFields.TransactionNumber).IntegerValue = pTransactionNumber
      mvExisting = False
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Init()
      mvClassFields(CreditSaleFields.BatchNumber).LongValue = pBatchNumber
      mvClassFields(CreditSaleFields.TransactionNumber).LongValue = pTransactionNumber
    End Sub

    Public Overloads Sub Update(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pAddressTo As String, ByVal pSalesLedgerAccount As String, ByVal pStockSale As Boolean)
      mvClassFields(CreditSaleFields.ContactNumber).IntegerValue = pContactNumber
      mvClassFields(CreditSaleFields.AddressNumber).IntegerValue = pAddressNumber
      mvClassFields(CreditSaleFields.AddressTo).Value = pAddressTo
      mvClassFields(CreditSaleFields.SalesLedgerAccount).Value = pSalesLedgerAccount
      mvClassFields(CreditSaleFields.StockSale).Bool = pStockSale
    End Sub

  End Class

End Namespace
