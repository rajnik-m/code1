Namespace Access

  Partial Public Class BatchTransaction

    'Public Enum BatchTransactionRecordSetTypes 'These are bit values
    '  btrtAll = &HFFS
    '  'ADD additional recordset types here
    '  btrtAllTT = &H1FFS
    'End Enum


    Private mvStockMovements As CDBCollection

    Protected Overrides Sub ClearFields()
      mvTransactionSign = ""
      mvContactVATCategory = ""
      mvContactType = Contact.ContactTypes.ctcContact
      mvStockMovements = Nothing
    End Sub

    'Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BatchTransactionRecordSetTypes)
    '  Dim vFields As CDBFields

    '  mvEnv = pEnv
    '  InitClassFields()
    '  vFields = pRecordSet.Fields
    '  mvExisting = True
    '  With mvClassFields
    '    'Modify below to handle each recordset type as required
    '    If (pRSType And BatchTransactionRecordSetTypes.btrtAll) = BatchTransactionRecordSetTypes.btrtAll Then
    '      .SetItem(BatchTransactionFields.BatchNumber, vFields)
    '      .SetItem(BatchTransactionFields.TransactionNumber, vFields)
    '      .SetItem(BatchTransactionFields.ContactNumber, vFields)
    '      .SetItem(BatchTransactionFields.AddressNumber, vFields)
    '      .SetItem(BatchTransactionFields.TransactionDate, vFields)
    '      .SetItem(BatchTransactionFields.TransactionType, vFields)
    '      .SetItem(BatchTransactionFields.BankDetailsNumber, vFields)
    '      .SetItem(BatchTransactionFields.Amount, vFields)
    '      .SetItem(BatchTransactionFields.CurrencyAmount, vFields)
    '      .SetItem(BatchTransactionFields.PaymentMethod, vFields)
    '      .SetItem(BatchTransactionFields.Reference, vFields)
    '      .SetItem(BatchTransactionFields.NextLineNumber, vFields)
    '      .SetItem(BatchTransactionFields.LineTotal, vFields)
    '      .SetItem(BatchTransactionFields.Mailing, vFields)
    '      .SetItem(BatchTransactionFields.Receipt, vFields)
    '      .SetItem(BatchTransactionFields.Notes, vFields)
    '      .SetItem(BatchTransactionFields.MailingContactNumber, vFields)
    '      .SetItem(BatchTransactionFields.MailingAddressNumber, vFields)
    '      .SetItem(BatchTransactionFields.AmendedBy, vFields)
    '      .SetItem(BatchTransactionFields.AmendedOn, vFields)
    '      .SetItem(BatchTransactionFields.EligibleForGiftAid, vFields)
    '      .SetOptionalItem(BatchTransactionFields.TransactionOrigin, vFields)
    '    End If
    '    If (pRSType And BatchTransactionRecordSetTypes.btrtAllTT) = BatchTransactionRecordSetTypes.btrtAllTT Then
    '      mvTransactionSign = vFields("transaction_sign").Value
    '      mvContactVATCategory = vFields("contact_vat_category").Value
    '      Select Case vFields("contact_type").Value
    '        Case "C"
    '          mvContactType = Contact.ContactTypes.ctcContact
    '        Case "O"
    '          mvContactType = Contact.ContactTypes.ctcOrganisation
    '        Case "J"
    '          mvContactType = Contact.ContactTypes.ctcJoint
    '      End Select
    '    End If
    '  End With
    'End Sub

    Friend ReadOnly Property ContactType() As Contact.ContactTypes
      Get
        ContactType = mvContactType
      End Get
    End Property
    Friend ReadOnly Property ContactVatCategory() As String
      Get
        ContactVatCategory = mvContactVATCategory
      End Get
    End Property
    Public ReadOnly Property TransactionSign() As String
      Get
        If mvTransactionSign.Length = 0 Then
          mvTransactionSign = mvEnv.Connection.GetValue("SELECT transaction_sign FROM transaction_types WHERE transaction_type = '" & TransactionType & "'")
        End If
        TransactionSign = mvTransactionSign
      End Get
    End Property

    Public ReadOnly Property IsFinancialAdjustment() As Boolean
      Get
        Dim vWhereFields As New CDBFields

        vWhereFields.Add("batch_number", BatchNumber)
        vWhereFields.Add("transaction_number", TransactionNumber)
        Return mvEnv.Connection.GetCount("reversals", vWhereFields) > 0
      End Get
    End Property

    Public Sub InitAnalysisFromRecordSets(ByVal pEnv As CDBEnvironment, ByVal pBTARS As CDBRecordSet, ByVal pBTARSP As CDBRecordSet)
      mvAnalysis = New CollectionList(Of BatchTransactionAnalysis)
      Dim vAdded As Boolean
      Dim vAnalysis As BatchTransactionAnalysis
      Do
        vAdded = False
        If pBTARS.Status Then
          If pBTARS.Fields("transaction_number").LongValue = TransactionNumber Then
            vAnalysis = AddAnalysisFromRecordSet(pBTARS)
            Select Case vAnalysis.LineType
              Case "P", "G"
                If vAnalysis.ProductCode.Length = 0 Then RaiseError(DataAccessErrors.daeProductInvalid)
            End Select
            If vAnalysis.ProductCode.Length > 0 Then
              'This should have product information with it
              If pBTARSP.Status And pBTARSP.Fields("transaction_number").LongValue = TransactionNumber And pBTARSP.Fields("line_number").LongValue = vAnalysis.LineNumber Then
                vAnalysis.InitProductFromRecordSet(pBTARSP, Product.ProductRecordSetTypes.prstMain)
                pBTARSP.Fetch()
              Else
                RaiseError(DataAccessErrors.daeProductInvalid, (vAnalysis.ProductCode))
              End If
            End If
            vAdded = True
            pBTARS.Fetch()
          End If
        End If
      Loop While pBTARS.Status And vAdded = True
      If mvAnalysis.Count() = 0 Then
        'No BTA have been added
        'There are a number of reasons for this - BT w/out BTA, BTA w/out BT, etc.
        RaiseError(DataAccessErrors.daeBTAndBTADoNotMatch, CStr(BatchNumber))
      End If
    End Sub

    Public ReadOnly Property StockMovements() As CDBCollection
      Get
        If mvStockMovements Is Nothing Then mvStockMovements = New CDBCollection
        Return mvStockMovements
      End Get
    End Property

    Public Sub InitAnalysisStockMovements()
      mvStockMovements = New CDBCollection
      Dim vStockMovement As New StockMovement
      vStockMovement.Init(mvEnv)
      Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vStockMovement.GetRecordSetFields(StockMovement.StockMovementRecordSetTypes.smrtAll) & " FROM stock_movements sm WHERE batch_number = " & BatchNumber & " AND transaction_number = " & TransactionNumber & " ORDER BY stock_movement_number DESC")
      With vRS
        While .Fetch
          vStockMovement = New StockMovement
          vStockMovement.InitFromRecordSet(mvEnv, vRS, StockMovement.StockMovementRecordSetTypes.smrtAll)
          If (vStockMovement.StockMovementReason <> mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonAwaitBackOrder) And vStockMovement.MovementQuantity <> 0) Or (vStockMovement.StockMovementReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonAwaitBackOrder) And vStockMovement.MovementQuantity = 0) Then
            If vStockMovement.StockMovementNumber <> 0 Then
              mvStockMovements.Add(vStockMovement, vStockMovement.StockMovementNumber.ToString)
            End If
          End If
        End While
        .CloseRecordSet()
      End With
    End Sub

    Public Sub InitDetailsFromFinancialHistory(ByVal pEnv As CDBEnvironment, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      Dim vFH As New FinancialHistory
      Dim vPaymentHistory As New Collection
      Dim vOPH As New OrderPaymentHistory
      Dim vFinancialLinks As New Collection
      Dim vFL As New FinancialLink
      Dim vBackOrderDetails As New Collection
      Dim vBOD As New BackOrderDetail
      Dim vOriginalBTAs As New Collection
      Dim vOriginalBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vRecordSet As CDBRecordSet
      Dim vRemovedOne As Boolean
      Dim vIndex As Integer
      Dim vBTA As BatchTransactionAnalysis
      Dim vFHD As FinancialHistoryDetail
      Dim vMultiplier As Integer
      Dim vUsedBOD As Boolean
      Dim vLines As New CDBParameters
      Dim vIPHRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vKey As String
      Dim vCC As New CompanyControl
      Dim vAddOverpayment As Boolean
      Dim vAddBTA As Boolean

      mvEnv = pEnv
      Init(pBatchNumber, pTransactionNumber)
      mvClassFields.Item(BatchTransactionFields.NextLineNumber).IntegerValue = 1
      vCC.Init(mvEnv)
      'create FH object w/ FHD collection
      vFH.Init(mvEnv, pBatchNumber, pTransactionNumber)
      TransactionDate = vFH.TransactionDate
      Do
        vRemovedOne = False
        For vIndex = 1 To vFH.Details.Count()
          If DirectCast(vFH.Details.Item(vIndex), FinancialHistoryDetail).Status <> FinancialHistory.FinancialHistoryStatus.fhsNormal Then
            vFH.Details.Remove(vIndex)
            vRemovedOne = True
            Exit For
          Else
            If vIndex = 1 Then vLines = New CDBParameters
            Dim vLineNumber As Integer = DirectCast(vFH.Details.Item(vIndex), FinancialHistoryDetail).LineNumber
            If Not vLines.Exists(vLineNumber.ToString) Then
              vLines.Add(vLineNumber.ToString, vLineNumber)
            End If
          End If
        Next
      Loop While vRemovedOne
      'create OPH collection
      vOPH.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOPH.GetRecordSetFields(OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll) & " FROM order_payment_history oph WHERE batch_number = " & vFH.BatchNumber & " AND transaction_number = " & vFH.TransactionNumber & " AND line_number IN (" & vLines.ItemList & ") ORDER BY line_number")
      With vRecordSet
        While .Fetch() = True
          vOPH = New OrderPaymentHistory
          vOPH.InitFromRecordSet(mvEnv, vRecordSet, OrderPaymentHistory.OrderPaymentHistoryRecordSetTypes.ophrtAll)
          vPaymentHistory.Add(vOPH)
        End While
        .CloseRecordSet()
      End With
      'create Original BTA collection
      vOriginalBTA.Init()
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOriginalBTA.GetRecordSetFields() & " FROM batch_transaction_analysis bta WHERE batch_number = " & vFH.BatchNumber & " AND transaction_number = " & vFH.TransactionNumber & " ORDER BY line_number")
      With vRecordSet
        While .Fetch() = True
          vOriginalBTA = New BatchTransactionAnalysis(mvEnv)
          vOriginalBTA.InitFromRecordSet(vRecordSet)
          If vOriginalBTA.LineType = "M" Or vOriginalBTA.LineType = "C" Then vOriginalBTA.LineType = "O"
          vOriginalBTAs.Add(vOriginalBTA)
        End While
        .CloseRecordSet()
      End With
      'create FL collection
      vFL.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vFL.GetRecordSetFields(FinancialLink.FinancialLinkRecordSetTypes.flrtAll) & " FROM financial_links WHERE donor_contact_number = " & vFH.ContactNumber & " AND batch_number = " & vFH.BatchNumber & " AND transaction_number = " & vFH.TransactionNumber & " AND line_number IN (" & vLines.ItemList & ") AND line_type IN ('G','S','H') ORDER BY line_number")
      With vRecordSet
        While .Fetch() = True
          vFL = New FinancialLink
          vFL.InitFromRecordSet(mvEnv, vRecordSet, FinancialLink.FinancialLinkRecordSetTypes.flrtAll)
          vFinancialLinks.Add(vFL)
        End While
        .CloseRecordSet()
      End With
      'create invoice payments record set
      vSQL = "SELECT fhd.line_number, i2.invoice_number, fhd.amount, fhd.source, i.record_type, i.invoice_number AS  paying_invoice_number, i2.sales_ledger_account"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then vSQL = vSQL & ",fhd.currency_amount,fhd.currency_vat_amount"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) Then vSQL = vSQL & ", iph.batch_number, iph.allocation_batch_number"
      vSQL = vSQL & " FROM financial_history_details fhd, invoice_payment_history iph, invoices i, invoices i2"
      vSQL = vSQL & " WHERE fhd.batch_number = " & vFH.BatchNumber
      vSQL = vSQL & " AND fhd.transaction_number = " & vFH.TransactionNumber
      vSQL = vSQL & " AND fhd.status IS NULL"
      vSQL = vSQL & " AND invoice_payment = 'Y'"
      vSQL = vSQL & " AND fhd.batch_number = iph.batch_number"
      vSQL = vSQL & " AND fhd.transaction_number = iph.transaction_number "
      vSQL = vSQL & " AND fhd.line_number= iph.line_number"
      vSQL = vSQL & " AND iph.batch_number = i.batch_number"
      vSQL = vSQL & " AND iph.transaction_number = i.transaction_number"
      vSQL = vSQL & " AND iph.invoice_number = i2.invoice_number"
      vSQL = vSQL & " ORDER BY iph.allocation_batch_number" & mvEnv.Connection.DBSortByNullsFirst & ", fhd.line_number" 'BR17007: Changed order by to include allocation_batch_number nulls first (to get N type line out first) to prevent issue whereby a transaction containing N & U type lines was not including the N type line in the mvAnalysisLines collection and the below RaiseError(DataAccessErrors.daeCannotGetAnalysis) was being generated.
      vIPHRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      'create BOD collection
      vBOD.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vBOD.GetRecordSetFields(BackOrderDetail.BackOrderDetailRecordSetTypes.bodrtAll) & " FROM back_order_details bod WHERE batch_number = " & vFH.BatchNumber & " AND transaction_number = " & vFH.TransactionNumber & " AND line_number IN (" & vLines.ItemList & ") ORDER BY line_number")
      With vRecordSet
        While .Fetch() = True
          vBOD = New BackOrderDetail
          vBOD.InitFromRecordSet(mvEnv, vRecordSet, BackOrderDetail.BackOrderDetailRecordSetTypes.bodrtAll)
          vBackOrderDetails.Add(vBOD)
        End While
        .CloseRecordSet()
      End With

      'If we got OPH data then get the Company Controls
      If vPaymentHistory.Count() > 0 Then
        vSQL = "SELECT " & vCC.GetRecordSetFields(CompanyControl.CompanyControlRecordSetTypes.cocrtAll) & " FROM batches b, bank_accounts ba, company_controls cc"
        vSQL = vSQL & " WHERE b.batch_number = " & pBatchNumber & " AND ba.bank_account = b.bank_account AND cc.company = ba.company"
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then vCC.InitFromRecordSet(mvEnv, vRecordSet, CompanyControl.CompanyControlRecordSetTypes.cocrtAll)
        vRecordSet.CloseRecordSet()
      End If

      'got all the data, now create BTA collection
      vLines = New CDBParameters
      If mvTransactionSign = "D" Then
        vMultiplier = -1
      Else
        vMultiplier = 1
      End If
      mvAnalysis = New CollectionList(Of BatchTransactionAnalysis)
      'start w/ OPH
      For Each vOPH In vPaymentHistory
        vAddBTA = True
        If Not vLines.Exists(vOPH.LineNumber.ToString) Then
          vBTA = New BatchTransactionAnalysis(mvEnv)
          vBTA.InitFromTransaction(Me)
          vBTA.LineType = "O"
          vBTA.LineNumber = vOPH.LineNumber 'Need to explicitly set this
          vBTA.PaymentPlanNumber = vOPH.OrderNumber
          vBTA.Amount = vOPH.Amount * vMultiplier
          For Each vFHD In vFH.Details
            If vFHD.LineNumber = vOPH.LineNumber Then
              If (vFHD.ProductCode = vCC.OverPaymentProductCode) And (vFHD.RateCode = vCC.OverPaymentRate) Then
                vAddOverpayment = True
                vAddBTA = Not (vBTA.Amount = vFHD.Amount)
              Else
                If vFHD.SalesContactNumber > 0 Then vBTA.SalesContactNumber = vFHD.SalesContactNumber
                If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
                  vBTA.CurrencyAmount = vBTA.CurrencyAmount + (vFHD.CurrencyAmount * vMultiplier)
                  vBTA.CurrencyVatAmount = vBTA.CurrencyVatAmount + (vFHD.CurrencyVATAmount * vMultiplier)
                End If
              End If
              vBTA.Source = vFHD.Source
            End If
          Next vFHD
          If vOPH.Balance <> 0 Then vBTA.AcceptAsFull = True
          For Each vFL In vFinancialLinks
            If vBTA.LineNumber = vFL.LineNumber Then
              vBTA.LineType = vFL.LineType
              If vFL.LineType = "H" And vBTA.PaymentPlanNumber > 0 Then
                vBTA.DeceasedContactNumber = vFL.DonorContactNumber
              Else
                vBTA.DeceasedContactNumber = vFL.ContactNumber
              End If
              Exit For
            End If
          Next vFL
          If vAddBTA Then mvAnalysis.Add(vBTA.LineNumber.ToString & vBTA.LineType & vBTA.ProductCode, vBTA)
          vLines.Add(vOPH.LineNumber.ToString, vOPH.LineNumber)
        End If
      Next vOPH
      'do IPH
      With vIPHRecordSet
        While .Fetch() = True
          If Not vLines.Exists(.Fields(1).Value) Then
            vBTA = New BatchTransactionAnalysis(mvEnv)
            vBTA.InitFromTransaction(Me)
            If .Fields(5).Value = "C" Then
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAllocationsOnIPH) = True AndAlso .Fields("allocation_batch_number").Value.Length > 0 Then
                If .Fields("batch_number").IntegerValue = .Fields("allocation_batch_number").IntegerValue Then
                  vBTA.LineType = "N"
                Else
                  vBTA.LineType = "U"
                End If
              Else
                If .Fields(6).IntegerValue = .Fields(2).IntegerValue Then     '6 = PayingInvoiceNumber, 2 = InvoiceNumber
                  vBTA.LineType = "U"
                Else
                  vBTA.LineType = "N"
                End If
              End If
            ElseIf .Fields(5).Value = "N" Then
              vBTA.LineType = "R" 'DON'T KNOW IF THIS IS 100% CORRECT!!!
            End If
            vBTA.LineNumber = .Fields("line_number").IntegerValue 'Need to explicitly set this
            vBTA.InvoiceNumber = .Fields(2).IntegerValue
            vBTA.Amount = .Fields(3).DoubleValue * vMultiplier
            vBTA.Source = .Fields(4).Value
            vBTA.MemberNumber = .Fields(7).Value
            vBTA.PaymentPlanNumber = .Fields("paying_invoice_number").IntegerValue
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
              vBTA.CurrencyAmount = .Fields(8).DoubleValue * vMultiplier
              vBTA.CurrencyVatAmount = .Fields(9).DoubleValue * vMultiplier
            End If
            For Each vFL In vFinancialLinks
              If vBTA.LineNumber = vFL.LineNumber Then
                vBTA.LineType = vFL.LineType
                If vFL.LineType = "H" And vBTA.PaymentPlanNumber > 0 Then
                  vBTA.DeceasedContactNumber = vFL.DonorContactNumber
                Else
                  vBTA.DeceasedContactNumber = vFL.ContactNumber
                End If
                Exit For
              End If
            Next vFL
            mvAnalysis.Add(vBTA.LineNumber.ToString & vBTA.LineType & vBTA.ProductCode, vBTA)
            vLines.Add(.Fields(1).IntegerValue.ToString, .Fields(1).IntegerValue)
          End If
        End While
        .CloseRecordSet()
      End With
      'now do FHD, excluding those lines already covered by OPH & IPH
      For Each vFHD In vFH.Details
        If Not vLines.Exists(CStr(vFHD.LineNumber)) Then
          'see if any BOD exists for this line or any earlier lines
          vUsedBOD = False
          For Each vBOD In vBackOrderDetails
            If Not vLines.Exists(CStr(vBOD.LineNumber)) Then
              If vBOD.LineNumber <= vFHD.LineNumber Then
                If vBOD.LineNumber < vFHD.LineNumber Then
                  ReverseBackOrder(vBOD)
                  vLines.Add(vBOD.LineNumber.ToString, vBOD.LineNumber)
                Else
                  ReverseBackOrder(vBOD)
                End If
                vUsedBOD = True
              End If
            End If
          Next vBOD
          If Not vUsedBOD Then
            vBTA = New BatchTransactionAnalysis(mvEnv)
            vBTA.InitFromTransaction(Me)
            vBTA.LineType = "P"
            vBTA.LineNumber = vFHD.LineNumber 'Need to explicitly set this
            'BR 7991: Explictly Check for Legacy Bequest Receipt
            'and Warehouse (BR8353)
            For Each vOriginalBTA In vOriginalBTAs
              If vFHD.LineNumber = vOriginalBTA.LineNumber Then
                Select Case vOriginalBTA.LineType
                  Case "B"
                    vBTA.LineType = "B"
                    vBTA.WhenValue = vOriginalBTA.WhenValue
                  Case "D", "F"
                    vBTA.LineType = vOriginalBTA.LineType
                  Case "P"
                    'Need to set the Warehouse from OriginalBTA as not contained in FHD
                    If vOriginalBTA.Warehouse.Length > 0 Then vBTA.Warehouse = vOriginalBTA.Warehouse
                  Case "U"
                    vBTA.LineType = "U"
                    vBTA.MemberNumber = vOriginalBTA.MemberNumber
                  Case "X"
                    vBTA.LineType = "X"
                  Case "Q"
                    vBTA.LineType = "Q"
                End Select
                Exit For
              End If
            Next vOriginalBTA
            'does this line exist in the financial_links table?
            For Each vFL In vFinancialLinks
              If vFHD.LineNumber = vFL.LineNumber Then
                If Not (vBTA.LineType = "D" Or vBTA.LineType = "F") Then vBTA.LineType = vFL.LineType
                If Not ((vBTA.LineType = "D" Or vBTA.LineType = "F") AndAlso (vFL.LineType = "H" Or vFL.LineType = "S")) Then
                  vBTA.DeceasedContactNumber = vFL.ContactNumber
                End If
                If Not (vBTA.LineType = "D" Or vBTA.LineType = "F") Then Exit For
              End If
            Next vFL
            vBTA.Source = vFHD.Source
            vBTA.ProductCode = vFHD.ProductCode
            vBTA.RateCode = vFHD.RateCode
            vBTA.DistributionCode = vFHD.DistributionCode
            If (vBTA.LineType = "D" OrElse vBTA.LineType = "F") Then
              vBTA.ContactNumber = vOriginalBTA.ContactNumber
              vBTA.AddressNumber = vOriginalBTA.AddressNumber
            Else
              vBTA.ContactNumber = vFH.ContactNumber
              vBTA.AddressNumber = vFH.AddressNumber
            End If
            If vFHD.SalesContactNumber > 0 Then vBTA.SalesContactNumber = vFHD.SalesContactNumber
            'changed to use setVATRate on 15th June 2001-Pooja
            'vBTA.VatRate = vFHD.VatRate
            vBTA.Amount = vFHD.Amount * vMultiplier
            'changed to use setVATRate on 15th June 2001-Pooja
            'vBTA.VatAmount = vFHD.VatAmount * vMultiplier
            vBTA.SetVATRate((vFHD.VatRate), vFHD.VatAmount * vMultiplier, vFHD.CurrencyVATAmount * vMultiplier)
            vBTA.Issued = IntegerValue(vFHD.Quantity) * vMultiplier
            vBTA.Quantity = IntegerValue(vFHD.Quantity) * vMultiplier
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
              vBTA.CurrencyAmount = vFHD.CurrencyAmount * vMultiplier
              'changed to use setVATRate on 15th June 2001-Pooja
              'vBTA.CurrencyVATAmount = vFHD.CurrencyVATAmount * vMultiplier
            End If
            mvAnalysis.Add(vBTA.LineNumber.ToString & vBTA.LineType & vBTA.ProductCode, vBTA)
          End If
          vLines.Add(vFHD.LineNumber.ToString, vFHD.LineNumber)
        End If
      Next vFHD
      'Now add any overpayments
      If vAddOverpayment Then
        For Each vFHD In vFH.Details
          If (vFHD.ProductCode = vCC.OverPaymentProductCode) And (vFHD.RateCode = vCC.OverPaymentRate) Then
            vBTA = New BatchTransactionAnalysis(mvEnv)
            With vBTA
              .InitFromTransaction(Me)
              .LineType = "P"
              .LineNumber = vFHD.LineNumber
              .Source = vFHD.Source
              .ProductCode = vFHD.ProductCode
              .RateCode = vFHD.RateCode
              .DistributionCode = vFHD.DistributionCode
              .Amount = vFHD.Amount
              .ContactNumber = vFH.ContactNumber
              .AddressNumber = vFH.AddressNumber
              If vFHD.SalesContactNumber > 0 Then .SalesContactNumber = vFHD.SalesContactNumber
              .SetVATRate((vFHD.VatRate), vFHD.VatAmount * vMultiplier, vFHD.CurrencyVATAmount * vMultiplier)
              .Quantity = IntegerValue(vFHD.Quantity) * vMultiplier
              .Issued = .Quantity
              If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
                .CurrencyAmount = vFHD.CurrencyAmount * vMultiplier
              End If
            End With
            mvAnalysis.Add(vBTA.LineNumber.ToString & vBTA.LineType & vBTA.ProductCode, vBTA)
          End If
        Next vFHD
      End If
      'Set Additional data
      InitAnalysisAdditionalData()

      'Check EventBooking lines
      For Each vBTA In mvAnalysis
        If vBTA.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBooking Then
          vBTA.MemberNumber = CStr(vBTA.AdditionalNumber)
        End If
      Next vBTA

      'now add details held only on the original BTA
      For Each vOriginalBTA In vOriginalBTAs
        vKey = CStr(vOriginalBTA.LineNumber) & vOriginalBTA.LineType & vOriginalBTA.ProductCode
        If mvAnalysis.ContainsKey(vKey) Then
          If vOriginalBTA.ContactNumber > 0 Then
            mvAnalysis.Item(vKey).ContactNumber = vOriginalBTA.ContactNumber
            mvAnalysis.Item(vKey).AddressNumber = vOriginalBTA.AddressNumber
          Else
            mvAnalysis.Item(vKey).ClearContactAndAddressNumbers()
          End If
        ElseIf vOriginalBTA.LineType <> "O" Then
          RaiseError(DataAccessErrors.daeCannotGetAnalysis)
        End If
      Next vOriginalBTA
      'lastly, do any remaining BOD
      For Each vBOD In vBackOrderDetails
        If Not vLines.Exists(CStr(vBOD.LineNumber)) Then
          ReverseBackOrder(vBOD)
        End If
      Next vBOD
      'build a product object for all the BTA lines
      InitProducts()
    End Sub

    Private Sub InitProducts()
      Dim vRecordSet As CDBRecordSet
      Dim vProduct As New Product(mvEnv)
      Dim vBTA As BatchTransactionAnalysis
      Dim vProducts As String = ""
      For Each vBTA In mvAnalysis
        If vBTA.ProductCode.Length > 0 Then
          If InStr(vProducts, "'" & vBTA.ProductCode & "', ") = 0 Then
            vProducts = vProducts & "'" & vBTA.ProductCode & "', "
          End If
        End If
      Next vBTA
      If vProducts.Length > 0 Then
        vProducts = Left(vProducts, Len(vProducts) - 2)
        vProduct.Init()
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vProduct.GetRecordSetFields(Product.ProductRecordSetTypes.prstMain) & " FROM products p WHERE product IN (" & vProducts & ")")
        With vRecordSet
          While .Fetch() = True
            For Each vBTA In mvAnalysis
              If vBTA.Product Is Nothing And vBTA.ProductCode = .Fields("product").Value Then
                vBTA.InitProductFromRecordSet(vRecordSet, Product.ProductRecordSetTypes.prstMain)
                Exit For
              End If
            Next vBTA
          End While
          .CloseRecordSet()
        End With
      End If
    End Sub

    Private Sub ReverseBackOrder(ByRef pBOD As BackOrderDetail)
      Dim vBTA As BatchTransactionAnalysis
      Dim vMultiplier As Integer
      If mvTransactionSign = "D" Then
        vMultiplier = -1
      Else
        vMultiplier = 1
      End If

      vBTA = New BatchTransactionAnalysis(mvEnv)
      vBTA.InitFromTransaction(Me)
      vBTA.LineType = "P"
      vBTA.LineNumber = pBOD.LineNumber 'Need to explicitly set this
      vBTA.Source = pBOD.Source
      vBTA.ProductCode = pBOD.Product
      vBTA.RateCode = pBOD.RateCode
      vBTA.ContactNumber = pBOD.ContactNumber
      vBTA.AddressNumber = pBOD.AddressNumber
      vBTA.WhenValue = pBOD.EarliestDelivery
      vBTA.DespatchMethod = pBOD.DespatchMethod
      vBTA.VatRate = pBOD.VatRate
      vBTA.Amount = (pBOD.UnitPrice * pBOD.Ordered) * vMultiplier
      vBTA.VatAmount = (pBOD.VatAmount * pBOD.Ordered) * vMultiplier
      vBTA.Issued = pBOD.Issued * vMultiplier
      vBTA.Quantity = pBOD.Ordered * vMultiplier
      '  vBTA.AmendedBy = mvEnv.User.Logname
      '  vBTA.AmendedOn = TodaysDate()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCurrencyCode) Then
        vBTA.CurrencyAmount = (pBOD.CurrencyUnitPrice * pBOD.Ordered) * vMultiplier
        vBTA.CurrencyVATAmount = (pBOD.CurrencyVATAmount * pBOD.Ordered) * vMultiplier
      End If
      vBTA.Warehouse = pBOD.Warehouse
      mvAnalysis.Add(vBTA.LineNumber.ToString & vBTA.LineType & vBTA.ProductCode, vBTA)
    End Sub

    Public Sub InitFromBatch(ByVal pEnv As CDBEnvironment, ByRef pBatch As Batch, Optional ByRef pTransactionNumber As Integer = 0, Optional ByVal pTransactionSign As String = "")
      mvEnv = pEnv
      Me.Batch = pBatch
      InitClassFields()
      SetDefaults()
      If pTransactionSign.Length > 0 Then mvTransactionSign = pTransactionSign
      mvExisting = False
      If pTransactionNumber = 0 Then
        mvClassFields.Item(BatchTransactionFields.TransactionNumber).Value = CStr(pBatch.GetNextTransactionNumber)
      Else
        mvClassFields.Item(BatchTransactionFields.TransactionNumber).Value = CStr(pTransactionNumber)
        pBatch.SetNextTransactionNumber(pTransactionNumber)
      End If
      mvClassFields.Item(BatchTransactionFields.BatchNumber).Value = CStr(pBatch.BatchNumber)
      mvClassFields.Item(BatchTransactionFields.Amount).Value = CStr(0)
      mvClassFields.Item(BatchTransactionFields.LineTotal).Value = CStr(0)
      mvClassFields.Item(BatchTransactionFields.NextLineNumber).Value = CStr(1)
    End Sub

    Public Function AdjustmentTransactionType(ByVal pOrigBatchType As Batch.BatchTypes, ByVal pOrigTransSign As String, ByVal pAdjustmentType As Batch.AdjustmentTypes) As String
      Dim vRS As CDBRecordSet
      Dim vNewTransactionType As String = ""
      Dim vWhere As String

      Select Case pOrigBatchType
        Case Batch.BatchTypes.CreditSales
          vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCSCreditTransType)
        Case Else
          If pAdjustmentType = Batch.AdjustmentTypes.atAdjustment Then
            vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAdjustmentTransType)
          ElseIf pAdjustmentType <> Batch.AdjustmentTypes.atAdjustment Then
            vNewTransactionType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReverseTransType)
          End If
      End Select
      If vNewTransactionType.Length > 0 Then
        Return vNewTransactionType
      Else
        Dim vAdjustmentType As String = ""
        'Use first valid opposite Transaction Type
        vWhere = "transaction_sign = '"
        If pAdjustmentType = Batch.AdjustmentTypes.atAdjustment Then
          vWhere = vWhere & pOrigTransSign
        Else
          vWhere = vWhere & If(pOrigTransSign = "C", "D", "C")
        End If
        vWhere = vWhere & "' AND negatives_allowed = 'Y'"
        vRS = mvEnv.Connection.GetRecordSet("SELECT transaction_type FROM transaction_types WHERE " & vWhere)
        If vRS.Fetch Then vAdjustmentType = vRS.Fields(1).Value
        vRS.CloseRecordSet()
        Return vAdjustmentType
      End If
    End Function

    Public Sub AdjustTransactionAmount(ByVal pAmount As Double)
      'This routine is used by WEB services to adjust the amount of a provisional transaction
      'and is used when deleting an analysis line - a negative amount is usually passed in
      mvClassFields(BatchTransactionFields.Amount).DoubleValue = Amount + pAmount
      mvClassFields(BatchTransactionFields.CurrencyAmount).DoubleValue = CurrencyAmount + pAmount
      mvClassFields(BatchTransactionFields.LineTotal).DoubleValue = LineTotal + pAmount
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      SaveChanges()
      mvExisting = True
      mvClassFields.SetSaved()
    End Sub

    Public Sub SaveChanges()
      If mvExisting = False AndAlso mvBatch IsNot Nothing Then
        With mvBatch
          .NumberOfTransactions = .NumberOfTransactions + 1
          .BatchTotal = .BatchTotal + mvClassFields.Item(BatchTransactionFields.Amount).DoubleValue
          If mvClassFields.Item(BatchTransactionFields.Amount).Value.Length > 0 And mvClassFields.Item(BatchTransactionFields.CurrencyAmount).Value.Length = 0 Then mvClassFields.Item(BatchTransactionFields.CurrencyAmount).Value = mvClassFields.Item(BatchTransactionFields.Amount).Value
          .CurrencyBatchTotal = .CurrencyBatchTotal + mvClassFields.Item(BatchTransactionFields.CurrencyAmount).DoubleValue
        End With
      End If
      MyBase.Save("", False, 0)
      mvExisting = True
    End Sub

    Public Sub SetAmended(ByVal pAmendedOn As String, ByVal pAmendedBy As String)
      mvClassFields.Item(BatchTransactionFields.AmendedOn).Value = pAmendedOn
      mvClassFields.Item(BatchTransactionFields.AmendedBy).Value = pAmendedBy
      mvOverrideAmended = True
    End Sub

    Public Function CanAddGiftAidDeclaration() As Boolean
      Dim vFields As New CDBFields
      Dim vGADec As New GiftAidDeclaration
      Dim vPayPlan As New PaymentPlan
      Dim vRS As CDBRecordSet
      Dim vAdd As Boolean
      Dim vSQL As String

      '(1) Check transaction OK
      If BatchNumber > 0 Then
        If EligibleForGiftAid Then
          vGADec.Init(mvEnv, pRaiseNoGAControlError:=False)
          vAdd = vGADec.GADControlsExists
          If vAdd Then
            If CDate(TransactionDate) >= CDate(vGADec.GiftAidEarliestStartDate) Then
              If PaymentMethod <> vGADec.CAFPaymentMethod Then
                vAdd = True
              End If
            End If
          End If
        End If
      End If

      '(2) Check for other Declarations linked to this payment
      If vAdd Then
        With vFields
          .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End With
        If mvEnv.Connection.GetCount("gift_aid_declarations", vFields) > 0 Then vAdd = False
      End If

      '(3) Check for Payment Plans & Donation Products
      If vAdd Then
        'Only exclude Gift Memberships if the only analysis lines are paying for Gift Memberships
        With vFields
          .Remove("contact_number")
          .Add("bta.member_number", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
          .Add("m.member_number", CDBField.FieldTypes.cftLong, "bta.member_number")
          .Add("o.order_number", CDBField.FieldTypes.cftLong, "m.order_number")
        End With

        vAdd = False
        vPayPlan.Init(mvEnv)
        vSQL = Replace(vPayPlan.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstAll), "sales_contact_number", "o.sales_contact_number")
        vSQL = "SELECT " & vSQL & " FROM batch_transaction_analysis bta, members m, orders o"
        vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vFields)
        vSQL = vSQL & " ORDER BY gift_membership, o.order_number"
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        vRS.Fetch()
        If vRS.Status Then
          Do
            vPayPlan.InitFromRecordSet(mvEnv, vRS, PaymentPlan.PayPlanRecordSetTypes.pprstAll)
            vAdd = vPayPlan.MembershipEligibleForGiftAid(TransactionDate)
            If vAdd = True Then Exit Do
            vRS.Fetch()
          Loop While vRS.Status = True
        End If
        vRS.CloseRecordSet()

        If vAdd = False Then
          With vFields
            .Remove("bta.member_number")
            .Remove("m.member_number")
            .Remove("o.order_number")
            .Add("bta.order_number", CDBField.FieldTypes.cftLong, "0", CDBField.FieldWhereOperators.fwoGreaterThan)
            .Add("o.order_number", CDBField.FieldTypes.cftLong, "bta.order_number")
          End With
          vPayPlan.Init(mvEnv)
          vSQL = Replace(vPayPlan.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstAll), "sales_contact_number", "o.sales_contact_number")
          vSQL = "SELECT " & vSQL & " FROM batch_transaction_analysis bta, orders o"
          vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vFields)
          vSQL = vSQL & " ORDER BY gift_membership, o.order_number"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          vRS.Fetch()
          If vRS.Status Then
            Do
              vPayPlan.InitFromRecordSet(mvEnv, vRS, PaymentPlan.PayPlanRecordSetTypes.pprstAll)
              If vPayPlan.PlanType = CDBEnvironment.ppType.pptMember Then
                vAdd = vPayPlan.MembershipEligibleForGiftAid(TransactionDate)
              Else
                vAdd = True
              End If
              If vAdd = True Then Exit Do
              vRS.Fetch()
            Loop While vRS.Status = True
          End If
          vRS.CloseRecordSet()
        End If

        If vAdd = False Then
          With vFields
            .Clear()
            .Add("batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
            .Add("bta.product", CDBField.FieldTypes.cftLong, "p.product")
            .Add("donation", CDBField.FieldTypes.cftCharacter, "Y")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductEligibleGA) Then
              .Add("p.eligible_for_gift_aid", CDBField.FieldTypes.cftCharacter, "Y")
            End If
          End With
          If mvEnv.Connection.GetCount("batch_transaction_analysis bta, products p", vFields) > 0 Then 'Donation products
            vAdd = True
          End If
        End If
        vRS.CloseRecordSet()
      End If
      Return vAdd
    End Function

  End Class

End Namespace
