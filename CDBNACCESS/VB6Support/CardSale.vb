Namespace Access

  Partial Public Class CardSale

    Public Overloads Sub Clone(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pNoClaimRequired As Boolean)
      mvClassFields.ClearSetValues()
      mvClassFields.Item(CardSaleFields.BatchNumber).IntegerValue = pBatchNumber
      mvClassFields.Item(CardSaleFields.TransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields.Item(CardSaleFields.NoClaimRequired).Bool = pNoClaimRequired
      mvExisting = False
    End Sub

    Public Sub CloneFromCardSale(ByVal pOrigCardSale As CardSale)
      With pOrigCardSale
        mvClassFields.Item(CardSaleFields.IssueNumber).Value = .IssueNumber
        mvClassFields.Item(CardSaleFields.ValidDate).Value = .ValidDate
        mvClassFields.Item(CardSaleFields.ExpiryDate).Value = .ExpiryDate
        mvClassFields.Item(CardSaleFields.AuthorisationCode).Value = .AuthorisationCode
        mvClassFields.Item(CardSaleFields.CreditCardType).Value = .CreditCardType
        mvClassFields.Item(CardSaleFields.NoClaimRequired).Bool = .NoClaimRequired
        'mvClassFields.Item(SecurityCode).Value = .SecurityCode
        mvClassFields.Item(CardSaleFields.TemplateNumber).Value = .TemplateNumber
      End With
    End Sub

    Public Sub SetImportDetails(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pCreditCardType As String, ByVal pCardNumber As String, ByVal pAuthorisationCode As String, ByVal pExpiryDate As String, Optional ByVal pIssueNumber As String = "", Optional ByVal pValidDate As String = "")
      mvClassFields.Item(CardSaleFields.BatchNumber).IntegerValue = pBatchNumber
      mvClassFields.Item(CardSaleFields.TransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields.Item(CardSaleFields.CreditCardType).Value = pCreditCardType
      mvClassFields.Item(CardSaleFields.AuthorisationCode).Value = pAuthorisationCode
      mvClassFields.Item(CardSaleFields.ExpiryDate).Value = pExpiryDate
      If pIssueNumber.Length > 0 Then mvClassFields.Item(CardSaleFields.IssueNumber).Value = pIssueNumber
      If pValidDate.Length > 0 Then mvClassFields.Item(CardSaleFields.ValidDate).Value = pValidDate
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Init()
      mvClassFields(CardSaleFields.BatchNumber).IntegerValue = pBatchNumber
      mvClassFields(CardSaleFields.TransactionNumber).IntegerValue = pTransactionNumber
    End Sub

    Public Sub SetSecurityCode(ByVal pSecurityCode As String)
      mvClassFields.Item(CardSaleFields.SecurityCode).Value = pSecurityCode
    End Sub

  End Class

End Namespace
