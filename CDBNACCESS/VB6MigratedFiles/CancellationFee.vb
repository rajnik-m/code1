

Namespace Access
  Public Class CancellationFee

    Public Enum CancellationFeeRecordSetTypes 'These are bit values
      cfrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CancellationFeeFields
      cffAll = 0
      cffCancellationReason
      cffMinimumDays
      cffMaximumDays
      cffProduct
      cffRate
      cffPercentage
      cffAmendedBy
      cffAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvProductRate As ProductRate

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "cancellation_fees"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("cancellation_reason")
          .Add("minimum_days", CDBField.FieldTypes.cftInteger)
          .Add("maximum_days", CDBField.FieldTypes.cftInteger)
          .Add("product")
          .Add("rate")
          .Add("percentage", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvProductRate = Nothing
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CancellationFeeFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CancellationFeeFields.cffAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CancellationFeeFields.cffAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CancellationFeeRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CancellationFeeRecordSetTypes.cfrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cf")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromBooking(ByVal pEnv As CDBEnvironment, ByVal pCancellationReason As String, ByVal pDate As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vNoDays As Integer

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      vNoDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, CDate(pDate)))
      If vNoDays < 0 Then vNoDays = 0
      vWhereFields.Add((mvClassFields.Item(CancellationFeeFields.cffCancellationReason).Name), pCancellationReason)
      vWhereFields.Add((mvClassFields.Item(CancellationFeeFields.cffMaximumDays).Name), vNoDays, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      vWhereFields.Add((mvClassFields.Item(CancellationFeeFields.cffMinimumDays).Name), vNoDays, CDBField.FieldWhereOperators.fwoLessThanEqual)

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CancellationFeeRecordSetTypes.cfrtAll) & " FROM " & mvClassFields.DatabaseTableName & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(mvEnv, vRecordSet, CancellationFeeRecordSetTypes.cfrtAll)
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CancellationFeeRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CancellationFeeRecordSetTypes.cfrtAll) = CancellationFeeRecordSetTypes.cfrtAll Then
          .SetItem(CancellationFeeFields.cffCancellationReason, vFields)
          .SetItem(CancellationFeeFields.cffMinimumDays, vFields)
          .SetItem(CancellationFeeFields.cffMaximumDays, vFields)
          .SetItem(CancellationFeeFields.cffProduct, vFields)
          .SetItem(CancellationFeeFields.cffRate, vFields)
          .SetItem(CancellationFeeFields.cffPercentage, vFields)
          .SetItem(CancellationFeeFields.cffAmendedBy, vFields)
          .SetItem(CancellationFeeFields.cffAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Function IsCancellationAmountRequired() As Boolean
      If Percentage = 0 Then
        If ProductRate.Existing Then
          If ProductRate.PriceIsZero Then IsCancellationAmountRequired = True
        End If
      End If
    End Function

    Private Function ProductRate() As ProductRate
      If mvProductRate Is Nothing Then
        mvProductRate = New ProductRate(mvEnv)
        mvProductRate.Init(ProductCode, RateCode)
      End If
      ProductRate = mvProductRate
    End Function

    Public Sub AddCancellationFeeTransaction(ByRef pBatch As Batch, ByRef pBT As BatchTransaction, ByRef pBTA As BatchTransactionAnalysis, ByRef pCancellationAmount As Double, Optional ByVal pCancellationFeeTrans As BatchTransaction = Nothing)
      Dim vNewBatch As New Batch(mvEnv)
      Dim vBT As New BatchTransaction(mvEnv)
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vAmount As Double
      Dim vWhereFields As New CDBFields
      Dim vTransNo As Integer
      Dim vTransaction As Boolean
      Dim vInvoice As New Invoice
      Dim vOrigCreditSale As New CreditSale(mvEnv)
      Dim vCreditSale As New CreditSale(mvEnv)
      Dim vCreditCustomer As New CreditCustomer
      Dim vBankAccount As New BankAccount(mvEnv)
      Dim vCSTerms As New CreditSalesTerms
      Dim vDate As Date
      Dim vStatus As String = ""
      Dim vOrigInvoice As New Invoice
      Dim vOrigCardSale As New CardSale(mvEnv)
      Dim vCardSale As New CardSale(mvEnv)

      'get the product and rate
      Dim vProduct As New Product(mvEnv)
      If ProductCode.Length > 0 Then
        vProduct.Init(ProductCode)
      Else
        vProduct.InitWithRate(mvEnv, pBTA.ProductCode, pBTA.RateCode)
      End If
      Dim vProductRate As ProductRate
      If ProductRate.Existing Then
        vProductRate = ProductRate()
      Else
        vProductRate = vProduct.ProductRate
      End If
      Dim vContact As New Contact(mvEnv)
      vContact.Init(pBT.ContactNumber)
      Dim vVATRate As VatRate = mvEnv.VATRate(vProduct.ProductVatCategory, vContact.VATCategory)

      'figure out the amount
      If Percentage > 0 Then
        vAmount = (pBTA.Amount * Percentage) / 100
      Else
        If ProductRate.Existing Then
          vAmount = ProductRate.Price(0, vVATRate)       'No modifier support for cancellation fees
          If vAmount = 0 Then
            If pCancellationAmount = 0 Then RaiseError(DataAccessErrors.daeCancellationFeeMissing)
            vAmount = pCancellationAmount
          End If
        End If
      End If
      Dim vVatAmount As Double
      vVatAmount = vVATRate.CalculateVATAmount(vAmount, vProductRate.VatExclusive, TodaysDate)

      'Find an open batch of the correct type
      vNewBatch.InitOpenBatch(pBatch)
      vTransNo = vNewBatch.AllocateTransactionNumber
      If Not mvEnv.Connection.InTransaction Then
        mvEnv.Connection.StartTransaction()
        vTransaction = True
      End If

      With vBT
        .InitFromBatch(mvEnv, vNewBatch, vTransNo) 'Set up new transaction for the batch
        .ContactNumber = pBT.ContactNumber
        .AddressNumber = pBT.AddressNumber
        .TransactionDate = TodaysDate()
        .TransactionType = pBT.TransactionType
        .PaymentMethod = pBT.PaymentMethod
        .Reference = pBT.Reference
        .Receipt = "N"
        .EligibleForGiftAid = False
        .Notes = "Cancellation Fee"
        With vBTA
          .InitFromTransaction(vBT) 'Set up new analysis line for the transaction
          .LineType = pBTA.LineType
          .Amount = vAmount
          .CurrencyAmount = vAmount
          .Source = pBTA.Source
          .ProductCode = vProduct.ProductCode
          .RateCode = vProductRate.RateCode
          .Quantity = 1
          .WhenValue = TodaysDate()
          .VatAmount = vVatAmount
          .VatRate = vVATRate.VatRateCode
          If pBTA.ContactNumber > 0 Then .ContactNumber = pBTA.ContactNumber
          If pBTA.AddressNumber > 0 Then .AddressNumber = pBTA.AddressNumber
          .Save() 'Insert the analysis line
        End With
        .SaveChanges() 'Insert the transaction
      End With
      vNewBatch.AddTransactionAmount(vAmount)
      Select Case vNewBatch.BatchType
        Case Batch.BatchTypes.CreditSales
          If mvEnv.GetConfigOption("fp_use_sales_ledger", True) = True Then
            vOrigInvoice.Init(mvEnv, (pBT.BatchNumber), (pBT.TransactionNumber))
            vOrigCreditSale.Init((pBT.BatchNumber), (pBT.TransactionNumber))
            vBankAccount.Init(vNewBatch.BankAccount)
            vCreditCustomer.Init(mvEnv, (pBT.ContactNumber), vBankAccount.Company, vOrigCreditSale.SalesLedgerAccount)
            vCSTerms.Init(mvEnv, (pBT.ContactNumber), (vBankAccount.Company), (vOrigCreditSale.SalesLedgerAccount))
            'Create Credit Sale
            vCreditSale.Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
            vCreditSale.Update(vCreditCustomer.ContactNumber, vCreditCustomer.AddressNumber, vOrigCreditSale.AddressTo, vCreditCustomer.SalesLedgerAccount, False)
            vCreditSale.Save()
            'Create Invoice
            vInvoice.Init(mvEnv)
            If vInvoice.CalcInvPayDue(vCSTerms.TermsFrom, vCSTerms.TermsPeriod, vCSTerms.TermsNumber, vBTA.BatchNumber, vBTA.TransactionNumber, CDate(vBT.TransactionDate), vDate) = True Then
              If vInvoice.GetInvPayStatus(Invoice.InvPayStatuses.InvNotPaid, vStatus) = True Then
                vInvoice.Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
                vInvoice.Update(0, vCreditCustomer.ContactNumber, vCreditCustomer.AddressNumber, vCreditCustomer.Company, vCreditCustomer.SalesLedgerAccount, 0, -1, vBT.TransactionDate, CStr(vDate), vStatus, "I")
                vInvoice.Save()
                vCreditCustomer.AdjustOutstanding(vBT.Amount) 'J641: Update the credit customer record when adding a new invoice
                vCreditCustomer.Save()
              End If
            End If
          End If 'Use Sales Ledger
        Case Batch.BatchTypes.CreditCard
          vOrigCardSale.Init((pBT.BatchNumber), (pBT.TransactionNumber))
          With vCardSale
            .Create(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
            .CloneFromCardSale(vOrigCardSale)
            .Save()
          End With
      End Select
      If vTransaction Then mvEnv.Connection.CommitTransaction()

      Select Case vNewBatch.BatchType
        Case Batch.BatchTypes.CreditCard
          If vCardSale.TemplateNumber.Length > 0 Then
            If mvEnv.Connection.InTransaction Then
              'Currently used by UpdateEventBooking web service
              pCancellationFeeTrans = vBT
            Else
              'Currently used by UpdateAccommodationBookin web service, CancelEventBooking web service and EventCancellation task
              Dim vCCA As New CreditCardAuthorisation
              vCCA.InitFromTransaction(mvEnv, vBT.BatchNumber, vBT.TransactionNumber)
              vCCA.ContactNumber = vBT.ContactNumber
              vCCA.AuthoriseTransaction(vCardSale, CreditCardAuthorisation.CreditCardAuthorisationTypes.ccatNormal, vBT.Amount, vBT.AddressNumber)
              vCardSale.Save()
            End If
          End If
      End Select
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CancellationFeeFields.cffAll)
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
        AmendedBy = mvClassFields.Item(CancellationFeeFields.cffAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CancellationFeeFields.cffAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(CancellationFeeFields.cffCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property MaximumDays() As Integer
      Get
        MaximumDays = mvClassFields.Item(CancellationFeeFields.cffMaximumDays).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MinimumDays() As Integer
      Get
        MinimumDays = mvClassFields.Item(CancellationFeeFields.cffMinimumDays).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Percentage() As Double
      Get
        Percentage = mvClassFields.Item(CancellationFeeFields.cffPercentage).DoubleValue
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(CancellationFeeFields.cffProduct).Value
      End Get
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(CancellationFeeFields.cffRate).Value
      End Get
    End Property
  End Class
End Namespace
