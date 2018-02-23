Imports Advanced.LanguageExtensions
Namespace Access

  Partial Public Class BatchTransactionAnalysis

    Private mvScheduledPayment As OrderPaymentSchedule

    Public Enum TransactionAnalysisAdditionalTypes
      taatNotSet 'Not read
      taatNone 'No additional type
      taatEventBooking
      taatAccommodationBooking
      taatServiceBooking
      taatServiceBookingCredit
      taatServiceBookingEntitlement
      taatLegacyBequest
      taatCollectionPayments
      taatServiceBookingTransaction
      taatEventBookingTransaction
      taatFundraisingPayment
      taatInvoicePaymentHistory
      taatExamBooking
      taatExamBookingTransaction
    End Enum

    Public Function GetRecordSetFieldsProduct() As String
      Return "bta.batch_number,bta.transaction_number,bta.line_number,bta.product,issued,product_number"
    End Function

    Public Sub InitFromRecordSetProduct(ByVal pRS As CDBRecordSet)
      MyBase.InitFromRecordSetFields(pRS, GetRecordSetFieldsProduct)
    End Sub

    Public Sub InitProductFromRecordSet(ByVal pRecordSet As CDBRecordSet, ByVal pRSType As Product.ProductRecordSetTypes)
      If mvProduct Is Nothing Then mvProduct = New Product(mvEnv)
      mvProduct.InitFromRecordSet(pRecordSet, pRSType)
    End Sub

    Public Property Processed() As Boolean
      Get
        Processed = mvProcessed
      End Get
      Set(ByVal Value As Boolean)
        mvProcessed = Value
      End Set
    End Property

    Public ReadOnly Property TransactionContainsEventBooking() As Boolean
      Get
        TransactionContainsEventBooking = mvTransContainsEventBooking
      End Get
    End Property

    Public ReadOnly Property AnalysisAdditionalType() As TransactionAnalysisAdditionalTypes
      Get
        AnalysisAdditionalType = mvAdditionalType
      End Get
    End Property

    Public ReadOnly Property AdditionalNumber() As Integer
      Get
        AdditionalNumber = mvAdditionalNumber
      End Get
    End Property
    Public ReadOnly Property AdditionalNumber2() As Integer
      Get
        AdditionalNumber2 = mvAdditionalNumber2
      End Get
    End Property
    Public ReadOnly Property AdditionalNumber3() As Integer
      Get
        AdditionalNumber3 = mvAdditionalNumber3
      End Get
    End Property

    Public Function GetVATPercentage(ByVal pTransDate As String) As String
      Dim vPercentage As String = ""
      If VatRate.Length > 0 Then
        Dim vVAT As VatRate = mvEnv.VATRate(VatRate, "")
        vPercentage = vVAT.CurrentPercentage(pTransDate).ToString
      End If
      Return vPercentage
    End Function

    Public ReadOnly Property Product() As Product
      Get
        If mvProduct Is Nothing Then
          mvProduct = New Product(mvEnv)
          mvProduct.Init(ProductCode)
        End If
        Return mvProduct
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentNumber() As String
      Get
        'Return String as it could be null
        'Value will be looked up in OrderPaymentHistory
        Dim vFields As New CDBFields

        If mvSchPaymentNumber.Length = 0 And mvClassFields.Item(BatchTransactionAnalysisFields.PaymentPlanNumber).LongValue > 0 Then
          Select Case mvClassFields.Item(BatchTransactionAnalysisFields.LineType).Value
            Case "O", "M", "C", "S", "G", "H"
              With vFields
                .Add("order_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchTransactionAnalysisFields.PaymentPlanNumber).LongValue)
                .Add("batch_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchTransactionAnalysisFields.BatchNumber).LongValue)
                .Add("transaction_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchTransactionAnalysisFields.TransactionNumber).LongValue)
                .Add("line_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(BatchTransactionAnalysisFields.LineNumber).LongValue)
                .Add("amount", CDBField.FieldTypes.cftNumeric, mvClassFields.Item(BatchTransactionAnalysisFields.Amount).DoubleValue)
              End With
              mvSchPaymentNumber = mvEnv.Connection.GetValue("SELECT scheduled_payment_number FROM order_payment_history WHERE " & mvEnv.Connection.WhereClause(vFields))
          End Select
        End If
        Return mvSchPaymentNumber
      End Get
    End Property

    Public Property ScheduledPayment As OrderPaymentSchedule
      Get
        If mvScheduledPayment Is Nothing AndAlso Me.ScheduledPaymentNumber.IsNullOrWhitespace = False Then
          Me.ScheduledPayment = GetScheduledPayment(Me.ScheduledPaymentNumber)
        End If
        Return mvScheduledPayment
      End Get
      Private Set(value As OrderPaymentSchedule)
        mvScheduledPayment = value
      End Set
    End Property

    Private Function GetScheduledPayment(vScheduledPaymentNumber As String) As OrderPaymentSchedule
      Dim vDummy As New OrderPaymentSchedule()
      Return CARERecordFactory.SelectInstanceByPrimaryKey(Of OrderPaymentSchedule)(Me.Environment, vScheduledPaymentNumber, vDummy.ClassFields)
    End Function

    Friend Sub ChangeSign()
      mvClassFields.Item(BatchTransactionAnalysisFields.Amount).DoubleValue = -mvClassFields.Item(BatchTransactionAnalysisFields.Amount).DoubleValue
      If mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).LongValue <> 0 Then mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).Value = CStr(-mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).LongValue)
      If mvClassFields.Item(BatchTransactionAnalysisFields.Issued).LongValue <> 0 Then mvClassFields.Item(BatchTransactionAnalysisFields.Issued).Value = CStr(-mvClassFields.Item(BatchTransactionAnalysisFields.Issued).LongValue)
      If mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).Value.Length > 0 Then mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).Value = CStr(-mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).DoubleValue)
      If mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).Value.Length > 0 Then mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).Value = CStr(-mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).DoubleValue)
      If mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).Value.Length > 0 Then mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).Value = CStr(-mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).DoubleValue)
    End Sub

    Public Sub CloneForFA(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer)
      mvExisting = False
      mvClassFields.Item(BatchTransactionAnalysisFields.BatchNumber).IntegerValue = pBatchNumber
      mvClassFields.Item(BatchTransactionAnalysisFields.TransactionNumber).IntegerValue = pTransactionNumber
      mvClassFields.Item(BatchTransactionAnalysisFields.LineNumber).IntegerValue = pLineNumber
      mvClassFields.Item(BatchTransactionAnalysisFields.AmendedBy).Value = mvEnv.User.UserID
      mvClassFields.Item(BatchTransactionAnalysisFields.AmendedOn).Value = TodaysDate()
      mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).IntegerValue = Quantity * -1
      mvClassFields.Item(BatchTransactionAnalysisFields.Issued).IntegerValue = Issued * -1
      mvClassFields.Item(BatchTransactionAnalysisFields.Amount).DoubleValue = Amount * -1
      mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).DoubleValue = VatAmount * -1
      If GrossAmount.Length > 0 Then
        mvClassFields.Item(BatchTransactionAnalysisFields.GrossAmount).DoubleValue = DoubleValue(GrossAmount) * -1
        mvClassFields.Item(BatchTransactionAnalysisFields.Discount).DoubleValue = DoubleValue(Discount) * -1
      End If
      If mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).Value.Length > 0 Then
        mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).DoubleValue = CurrencyAmount * -1
        mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).DoubleValue = CurrencyVatAmount * -1
      End If
      mvClassFields.ClearSetValues()
    End Sub

    Public Sub CloneForPartRefund(ByVal pBTA As BatchTransactionAnalysis, ByVal pQuantity As Integer, ByVal pIssued As Integer, ByVal pCurrencyCode As String, ByVal pCurrencyExchangeRate As Double, ByVal pRefundAmount As Double)

      Dim vAmount As Double
      Dim vCVatCat As String
      Dim vPrice As Double


      'Create new bta as copy of old bta except for quantity,issued & amounts
      CloneFromBTA(pBTA)

      'Quantity may have changed so re-calculate the amount(s)
      If ProductCode.Length > 0 AndAlso (pQuantity <> mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).LongValue) Then
        vPrice = Val(mvEnv.Connection.GetValue("SELECT current_price FROM rates WHERE product = '" & ProductCode & "' AND rate = '" & RateCode & "'"))
        If vPrice = 0 AndAlso pBTA.Amount > 0 AndAlso pBTA.Quantity > 0 Then
          'When product price is zero then user was allowed to use a different price value to calculate the amount. 
          '- We will calculate the price used by the user. 
          vPrice = pBTA.Amount / pBTA.Quantity
        End If
        vAmount = vPrice * pQuantity
      ElseIf pRefundAmount > 0 Then
        vAmount = pRefundAmount
        pQuantity = pBTA.Quantity
        pIssued = pBTA.Issued
      End If

      mvClassFields.Item(BatchTransactionAnalysisFields.Amount).Value = vAmount.ToString
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) = True AndAlso pCurrencyCode.Length > 0 Then
        mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).Value = FixTwoPlaces(vAmount * pCurrencyExchangeRate).ToString
      End If

      If mvClassFields.Item(BatchTransactionAnalysisFields.VatRate).Value.Length > 0 Then
        Dim vVatRate As VatRate
        Dim vVatAmount As Double = 0
        vCVatCat = mvEnv.Connection.GetValue("SELECT contact_vat_category FROM contacts WHERE contact_number = " & ContactNumber)
        vVatRate = mvEnv.VATRate(Product.ProductVatCategory, vCVatCat)
        vVatAmount = CalculateVATAmount(vAmount, (vVatRate.Percentage))
        mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).Value = vVatAmount.ToString
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) = True AndAlso pCurrencyCode.Length > 0 Then
          mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).Value = FixTwoPlaces(vVatAmount * pCurrencyExchangeRate).ToString
        End If
      End If

      mvClassFields.Item(BatchTransactionAnalysisFields.Quantity).Value = pQuantity.ToString
      If pBTA.IssuedSet Then mvClassFields.Item(BatchTransactionAnalysisFields.Issued).Value = pIssued.ToString
    End Sub


    Public Sub DeleteFromBatch()
      'This routine is only intended for use for provisional transactions from the AddEventBooking WEB service
      'It does not take into account stock or other issues with removing an analysis line
      Dim vBatch As New Batch(mvEnv)
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vWhereFields As New CDBFields
      Dim vTransaction As Boolean

      mvClassFields(BatchTransactionAnalysisFields.BatchNumber).PrimaryKey = True 'Set up for delete
      mvClassFields(BatchTransactionAnalysisFields.TransactionNumber).PrimaryKey = True
      mvClassFields(BatchTransactionAnalysisFields.LineNumber).PrimaryKey = True

      vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
      vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)

      vBatch.Init(BatchNumber)
      If Not mvEnv.Connection.InTransaction Then
        mvEnv.Connection.StartTransaction()
        vTransaction = True
      End If
      If mvEnv.Connection.GetCount("batch_transaction_analysis", vWhereFields) = 1 Then
        mvEnv.Connection.DeleteRecords("batch_transactions", vWhereFields)
        vBatch.AddTransactionAmount(-Amount, True, True, -1)
      Else
        vBT.Init(BatchNumber, TransactionNumber, False)
        vBT.Amount = vBT.Amount - Amount
        vBT.LineTotal = vBT.LineTotal - Amount
        vBT.CurrencyAmount = vBT.CurrencyAmount - CurrencyAmount
        vBT.Save()
        vBatch.AddTransactionAmount(-Amount, False, True, 1)
      End If
      mvClassFields.Delete(mvEnv.Connection)
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Friend Sub SetAdditionalType(ByVal pAdditionalType As TransactionAnalysisAdditionalTypes, Optional ByRef pNumber As Integer = 0, Optional ByRef pNumber2 As Integer = 0, Optional ByRef pNumber3 As Integer = 0)
      mvAdditionalType = pAdditionalType
      mvAdditionalNumber = pNumber
      mvAdditionalNumber2 = pNumber2
      mvAdditionalNumber3 = pNumber3
      mvTransContainsEventBooking = False
    End Sub

    Friend Sub SetEventBookingTransactionAdditionalType(ByVal pNumber As Integer, ByVal pNumber2 As Integer, Optional ByRef pNumber3 As Integer = 0, Optional ByVal pTransContainsBooking As Boolean = False)
      mvAdditionalType = TransactionAnalysisAdditionalTypes.taatEventBookingTransaction
      mvAdditionalNumber = pNumber
      mvAdditionalNumber2 = pNumber2
      mvAdditionalNumber3 = pNumber3
      mvTransContainsEventBooking = pTransContainsBooking
    End Sub

    Public Sub SetVATRate(ByVal pVATRate As String, ByVal pVATAmount As Double, ByVal pCurrencyVATAmount As Double)
      mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).DoubleValue = pVATAmount
      mvClassFields.Item(BatchTransactionAnalysisFields.VatRate).Value = pVATRate
      mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).DoubleValue = pCurrencyVATAmount
    End Sub

    Public Sub SetVATAmounts(ByRef pVATRate As VatRate, Optional ByVal pContactVatCategory As String = "", Optional ByVal pTransactionDate As String = "")
      Dim vVatAmount As Double
      Dim vCurrencyVATAmount As Double
      Dim vPackProductRate As New ProductRate(mvEnv)
      Dim vVatSet As Boolean
      Dim vUnitPrice As Double
      Dim vCurrencyUnitPrice As Double

      mvClassFields.Item(BatchTransactionAnalysisFields.VatRate).Value = pVATRate.VatRateCode
      'Can only do pack product calculation if Contact VAT Category has been supplied
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPackProducts) = True And pContactVatCategory.Length > 0 Then
        If Product.PackProduct = True Then
          vPackProductRate.Init(ProductCode, RateCode)
          If Quantity = 0 Then
            vUnitPrice = 0
            vCurrencyUnitPrice = 0
          Else
            vUnitPrice = Amount / Quantity
            vCurrencyUnitPrice = CurrencyAmount / Quantity
          End If
          vVatAmount = vPackProductRate.PackVatAmount(vUnitPrice, Quantity, pContactVatCategory, pTransactionDate)
          vCurrencyVATAmount = vPackProductRate.PackVatAmount(vCurrencyUnitPrice, Quantity, pContactVatCategory, pTransactionDate)
          vVatSet = True
        End If
      End If
      If Not vVatSet Then
        vVatAmount = Int(((Amount - (Amount / (1 + pVATRate.CurrentPercentage(pTransactionDate) / 100))) * 100) + 0.5) / 100
        vCurrencyVATAmount = Int(((CurrencyAmount - (CurrencyAmount / (1 + pVATRate.CurrentPercentage(pTransactionDate) / 100))) * 100) + 0.5) / 100
      End If
      mvClassFields.Item(BatchTransactionAnalysisFields.VatAmount).Value = CStr(vVatAmount)
      mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyVatAmount).Value = CStr(vCurrencyVATAmount)
    End Sub

    Public ReadOnly Property Key(Optional ByVal pMultipleTransactions As Boolean = False) As String
      Get
        Dim vKey As String
        If pMultipleTransactions Then
          vKey = BatchNumber & "|" & TransactionNumber & "|" & LineNumber & LineType & ProductCode
        Else
          vKey = LineNumber & LineType & ProductCode
        End If
        Return vKey
      End Get
    End Property

    Public ReadOnly Property IsFinancialAdjustment As Boolean
      Get
        Return Me.Reversal IsNot Nothing
      End Get
    End Property

    Public Sub AddAddressAndContactNumbers(ByVal vAddressNo As Integer, ByVal vContactNo As Integer)
      'BR11758 - New procedure to handle the setting of address and contact numbers
      Me.AddressNumber = vAddressNo
      Me.ContactNumber = vContactNo
    End Sub

    Public Sub AllocateBoxDonation(ByVal pDistributionCode As String, ByVal pCurrencyAmount As Double, ByVal pSterlingAmount As Double, ByVal pExchangeRate As Double)
      'Used by AllocateDonationsToBoxes Task
      Dim vVatRate As VatRate
      Dim vCVatCat As String
      If pCurrencyAmount < CurrencyAmount Then
        mvClassFields.Item(BatchTransactionAnalysisFields.CurrencyAmount).Value = CStr(pCurrencyAmount)
        mvClassFields.Item(BatchTransactionAnalysisFields.Amount).Value = CStr(pSterlingAmount)
        'Calculate VAT if required
        If VatAmount > 0 Then
          vCVatCat = mvEnv.Connection.GetValue("SELECT contact_vat_category FROM contacts WHERE contact_number = " & ContactNumber)
          vVatRate = mvEnv.VATRate(Product.ProductVatCategory, vCVatCat)
          SetVATAmounts(vVatRate)
        End If
      End If
      mvClassFields.Item(BatchTransactionAnalysisFields.DistributionCode).Value = pDistributionCode
    End Sub

    Public Sub CloneForBoxAllocations(ByVal pBTA As BatchTransactionAnalysis, ByVal pNewLineNumber As Integer, ByVal pNewAmount As Double, ByVal pNewCurrencyAmount As Double, ByVal pTransactionDate As String)
      'When allocating Boxes to Donations, may need to create a new BTA record
      Dim vVatRate As VatRate
      Dim vCVatCat As String

      mvExisting = False
      CloneFromBTA(pBTA) 'Clone all data
      'Set these specific values differently
      With mvClassFields
        .Item(BatchTransactionAnalysisFields.BatchNumber).IntegerValue = pBTA.BatchNumber
        .Item(BatchTransactionAnalysisFields.TransactionNumber).IntegerValue = pBTA.TransactionNumber
        .Item(BatchTransactionAnalysisFields.LineNumber).IntegerValue = pNewLineNumber
        .Item(BatchTransactionAnalysisFields.Amount).Value = CStr(pNewAmount)
        .Item(BatchTransactionAnalysisFields.CurrencyAmount).Value = CStr(pNewCurrencyAmount)
        .Item(BatchTransactionAnalysisFields.DistributionCode).Value = ""
      End With

      'Set VAT if required
      If pBTA.VatAmount > 0 Then
        vCVatCat = mvEnv.Connection.GetValue("SELECT contact_vat_category FROM contacts WHERE contact_number = " & ContactNumber)
        vVatRate = mvEnv.VATRate(Product.ProductVatCategory, vCVatCat)
        SetVATAmounts(vVatRate, "", pTransactionDate)
      End If
    End Sub

    Public Sub SetAmended(ByVal pAmendedOn As String, ByVal pAmendedBy As String)
      mvClassFields.Item(BatchTransactionAnalysisFields.AmendedOn).Value = pAmendedOn
      mvClassFields.Item(BatchTransactionAnalysisFields.AmendedBy).Value = pAmendedBy
      mvOverrideAmended = True
    End Sub

    Friend Function IsPartRefund() As Boolean
      Dim vIsPartRefund As Boolean = False

      Dim vWhereFields As New CDBFields(New CDBField("r.batch_number", BatchNumber))
      vWhereFields.Add("r.transaction_number", TransactionNumber)
      vWhereFields.Add("r.line_number", LineNumber)
      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("batch_transaction_analysis bta", "r.was_batch_number", "bta.batch_number", "r.was_transaction_number", "bta.transaction_number", "r.was_line_number", "bta.line_number", AnsiJoin.AnsiJoinTypes.InnerJoin)})

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "bta.batch_number, bta.transaction_number, bta.line_number, bta.amount", "reversals r", vWhereFields, "", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      If vRS.Fetch Then
        If vRS.Fields("amount").DoubleValue.CompareTo(Amount) > 0 Then vIsPartRefund = True    'Original amount > new amount so must be a part-refund
      End If
      vRS.CloseRecordSet()

      Return vIsPartRefund
    End Function
  End Class

End Namespace
