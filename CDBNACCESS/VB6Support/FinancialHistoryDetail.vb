

Namespace Access
  Partial Public Class FinancialHistoryDetail

    Private mvBatchTransactionAnalysis As BatchTransactionAnalysis

    Public Enum FinancialHistoryDetailRecordSetTypes 'These are bit values
      fhdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    ''' <summary>DO NOT USE - VB6 CONVERTED CODE ONLY.</summary>
    Public Overloads Function GetRecordSetFields(ByVal pRSType As FinancialHistoryDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = FinancialHistoryDetailRecordSetTypes.fhdrtAll Then
        vFields = "fhd.batch_number AS fhd_batch_number,"
        vFields = vFields & "fhd.transaction_number AS fhd_transaction_number,"
        vFields = vFields & "fhd.line_number AS fhd_line_number,"
        vFields = vFields & "fhd.amount AS fhd_amount,"
        vFields = vFields & "fhd.product AS fhd_product,"
        vFields = vFields & "fhd.rate AS fhd_rate,"
        vFields = vFields & "fhd.source AS fhd_source,"
        vFields = vFields & "fhd.quantity AS fhd_quantity,"
        vFields = vFields & "fhd.vat_rate AS fhd_vat_rate,"
        vFields = vFields & "fhd.vat_amount AS fhd_vat_amount,"
        vFields = vFields & "fhd.status AS fhd_status,"
        If mvClassFields.Item(FinancialHistoryDetailFields.SalesContactNumber).InDatabase Then vFields = vFields & "fhd.sales_contact_number AS fhd_sales_contact_number,"
        If mvClassFields.Item(FinancialHistoryDetailFields.InvoicePayment).InDatabase Then vFields = vFields & "fhd.invoice_payment AS fhd_invoice_payment,"
        If mvClassFields.Item(FinancialHistoryDetailFields.CurrencyAmount).InDatabase Then
          vFields = vFields & "fhd.currency_amount AS fhd_currency_amount,"
          vFields = vFields & "fhd.currency_vat_amount AS fhd_currency_vat_amount,"
        End If
        If mvClassFields.Item(FinancialHistoryDetailFields.DistributionCode).InDatabase Then vFields = vFields & "fhd.distribution_code AS fhd_distribution_code,"
      End If
      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      Return vFields
    End Function

    ''' <summary>DO NOT USE - VB6 CONVERTED CODE ONLY.</summary>
    Public Overloads Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pBatchNumber > 0 And pTransactionNumber > 0 And pLineNumber > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(FinancialHistoryDetailRecordSetTypes.fhdrtAll) & " FROM financial_history_details fhd WHERE fhd.batch_number = " & pBatchNumber & " AND fhd.transaction_number = " & pTransactionNumber & " AND fhd.line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, FinancialHistoryDetailRecordSetTypes.fhdrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Protected Overloads Sub Init()
      MyBase.Init()
    End Sub

    Public Property BatchTransactionAnalysis As BatchTransactionAnalysis
      Get
        If mvBatchTransactionAnalysis Is Nothing AndAlso Me.BatchNumber > 0 AndAlso Me.TransactionNumber > 0 AndAlso Me.LineNumber > 0 Then
          Me.BatchTransactionAnalysis = Me.GetRelatedInstance(Of BatchTransactionAnalysis)({FinancialHistoryDetailFields.BatchNumber,
                                                                                           FinancialHistoryDetailFields.TransactionNumber,
                                                                                           FinancialHistoryDetailFields.LineNumber})
        End If
        Return mvBatchTransactionAnalysis
      End Get
      Private Set(value As BatchTransactionAnalysis)
        mvBatchTransactionAnalysis = value
      End Set
    End Property



    ''' <summary>DO NOT USE - VB6 CONVERTED CODE ONLY.</summary>
    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As FinancialHistoryDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True

      With pRecordSet
        mvClassFields.Item(FinancialHistoryDetailFields.BatchNumber).SetValue = CStr(.Fields("fhd_batch_number").IntegerValue)
        mvClassFields.Item(FinancialHistoryDetailFields.TransactionNumber).SetValue = .Fields("fhd_transaction_number").Value
        mvClassFields.Item(FinancialHistoryDetailFields.LineNumber).SetValue = .Fields("fhd_line_number").Value
        mvClassFields.Item(FinancialHistoryDetailFields.Amount).SetValue = CStr(.Fields("fhd_amount").DoubleValue)
        mvClassFields.Item(FinancialHistoryDetailFields.Product).SetValue = .Fields("fhd_product").Value
        mvClassFields.Item(FinancialHistoryDetailFields.Rate).SetValue = .Fields("fhd_rate").Value
        mvClassFields.Item(FinancialHistoryDetailFields.Source).SetValue = .Fields("fhd_source").Value
        mvClassFields.Item(FinancialHistoryDetailFields.Quantity).SetValue = .Fields("fhd_quantity").Value
        mvClassFields.Item(FinancialHistoryDetailFields.VatRate).SetValue = .Fields("fhd_vat_rate").Value
        mvClassFields.Item(FinancialHistoryDetailFields.VatAmount).SetValue = CStr(.Fields("fhd_vat_amount").DoubleValue)
        mvClassFields.Item(FinancialHistoryDetailFields.Status).SetValue = .Fields("fhd_status").Value
        If mvClassFields.Item(FinancialHistoryDetailFields.SalesContactNumber).InDatabase Then mvClassFields.Item(FinancialHistoryDetailFields.SalesContactNumber).SetValue = .Fields("fhd_sales_contact_number").Value 'WAS LongValue
        If mvClassFields.Item(FinancialHistoryDetailFields.InvoicePayment).InDatabase Then mvClassFields.Item(FinancialHistoryDetailFields.InvoicePayment).SetValue = .Fields("fhd_invoice_payment").Value
        If mvClassFields.Item(FinancialHistoryDetailFields.CurrencyAmount).InDatabase Then
          mvClassFields.Item(FinancialHistoryDetailFields.CurrencyAmount).SetValue = CStr(.Fields("fhd_currency_amount").DoubleValue)
          mvClassFields.Item(FinancialHistoryDetailFields.CurrencyVatAmount).SetValue = CStr(.Fields("fhd_currency_vat_amount").DoubleValue)
        End If
        If mvClassFields.Item(FinancialHistoryDetailFields.DistributionCode).InDatabase Then mvClassFields.Item(FinancialHistoryDetailFields.DistributionCode).SetValue = .Fields("fhd_distribution_code").Value
      End With
    End Sub

    ''' <summary>DO NOT USE - VB6 CONVERTED CODE ONLY.</summary>
    Public Overloads Sub Save()
      SetValid()
      If mvExisting Then
        'WARNING - No UNIQUE KEY Can't update, must delete all and re-insert from FinancialHistory.Save
        System.Diagnostics.Debug.Assert(False, "")
        'mvEnv.Connection.UpdateRecords "financial_history_details", mvClassFields.UpdateFields, mvClassFields.WhereFields
      Else
        mvClassFields.ClearSetValues() 'do this so all fhd details are included in the insert
        mvEnv.Connection.InsertRecord("financial_history_details", mvClassFields.UpdateFields)
      End If
    End Sub

  End Class
End Namespace
