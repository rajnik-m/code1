

Namespace Access
  Public Class TraderAnalysisLine
    'This is an individual Analysis Line in Trader

    Public Enum TraderAnalysisLineTypes
      taltAccomodation = 1 'A
      taltActivityEntry 'AA
      taltAddressUpdate 'ADDR
      taltAddSuppression 'AS
      taltLegacyBequestReceipt 'B
      taltCovenant 'C
      taltCreditCardAuthority 'CC
      taltCancelGiftAidDeclaration 'CG
      taltCreditCardAuthorityUpdate 'CCU
      taltCancelPaymentPlan 'CP
      taltDirectDebit 'DD
      taltDirectDebitUpdate 'DDU
      taltEvent 'E
      taltDeceased 'G
      taltGoneAway 'GA
      taltGiftAidDeclaration 'GD
      taltPayrollGivingPledge 'GP     Payroll Giving (Give As You Earn) Pledge
      taltHardCredit 'H
      taltIncentive 'I
      taltInvoiceAllocation 'L
      taltMembership 'M
      taltInvoicePayment 'N
      taltNoPayment 'NP
      taltPaymentPlan 'O      Order-type line
      taltProductSale 'P      Payment-type line
      taltSundryCreditNote 'R
      taltSoftCredit 'S
      taltStandingOrder 'SO
      taltStatus 'ST
      taltStandingOrderUpdate 'SOU
      taltUnallocatedSalesLedgerCash 'U
      taltServiceBooking 'V
      taltServiceBookingCredit 'VC
      taltServiceBookingEntitlement 'VE
      taltPostTaxPayrollGivingPayment 'PG
      taltCollectionPayment 'AP
      taltPreTaxPayrollGivingPayment 'PP
      taltEventPricingMatrixLine 'X
      taltInMemoriamHardCredit    'D
      taltInMemoriamSoftCredit    'F
      taltExamBooking 'Q
      taltSundryCreditNoteInvoiceAllocation 'K
    End Enum

    Private Enum TraderAnalysisLineFields
      talfLineNumber = 1
      talfTraderTransactionType
      talfTraderLineType
      talfPaymentPlanNumber
      talfProductCode
      talfRate
      talfDistributionCode
      talfQuantity
      talfSource
      talfGrossAmount
      talfDiscount
      talfAmount 'This is the Amount in the currency the Batch has been created in
      talfAcceptAsFull
      talfLineDate
      talfDespatchMethod
      talfDeliveryContactNumber
      talfDeliveryAddressNumber
      talfVatRate
      talfVATPercentage
      talfSalesContactNumber
      talfSalesLedgerAccount
      talfIssued
      talfStockSale
      talfNotes
      talfFinancialAdjustment
      talfWarehouseCode
      talfProductNumber
      talfGiverContactNumber
      talfScheduledPaymentNumber
      talfProvisionalBatchNumber
      talfProvisionalTransactionNumber
      talfProvisionalLineNumber
      talfDeceasedContactNumber
      talfPaymentPlanType
      talfStockMovementNumbers
      talfMemberNumber
      talfInvoiceNumber
      talfInvoiceNumberUsed
      talfEventBookingNumber
      talfRoomBookingNumber
      talfServiceBookingNumber
      talfLegacyNumber
      talfBequestNumber
      talfLegacyReceiptNumber
      talfCovenantNumber
      talfCancellationReason
      talfDeclarationNumber
      talfPGDonorID 'Payroll Giving Donor ID
      talfContactStatus
      talfPostagePacking
      talfActivityGroup
      talfSuppression
      talfAutoPaymentNumber
      talfIncentiveLineNumber
      talfVATAmount 'This is the VAT Amount in the currency the Batch has been created in
      talfPGPledgeNumber
      talfCollectionNumber
      talfCollectionPISNumber
      talfCollectionBankAccount
      talfCollectionBoxNumbers
      talfCollectionBoxAmounts
      talfStockTransactionID
      talfPriceVATExclusive
      talfEventNumber
      talfBookingOptionNumber
      talfAdultQuantity
      talfChildQuantity
      talfStartTime
      talfEndTime
      talfAmendedEventBookingNumber
      talfCreditedContactNumber
      talfInvoiceTypeUsed
      talfDepositAllowed  'Y/N flag set for Credit Sale Transaction Analysis Lines where the line is eligible/ in-eligible for a Deposit
      talfExamBookingNumber
      talfExamUnitId
      talfExamUnitProductId
    End Enum
    'Note: Currency Amounts for the Base Currency will be calculated as the BTA Line is created

    Public Enum TraderAnalysisLineInfo
      taliCreatesBTA
      taliHasStockMovement
      taliIsMaintenanceType
      taliIsInvoiceAllocation
    End Enum

    Private mvClassFields As ClassFields
    Private mvCreditedContact As Contact

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .Add("LineNumber", CDBField.FieldTypes.cftLong)
          .Add("TraderTransactionType")
          .Add("TraderLineType")
          .Add("PaymentPlanNumber", CDBField.FieldTypes.cftLong)
          .Add("ProductCode")
          .Add("Rate")
          .Add("DistributionCode")
          .Add("Quantity", CDBField.FieldTypes.cftLong)
          .Add("Source")
          .Add("GrossAmount")
          .Add("Discount")
          .Add("Amount", CDBField.FieldTypes.cftNumeric)
          .Add("AcceptAsFull")
          .Add("LineDate")
          .Add("DespatchMethod")
          .Add("DeliveryContactNumber", CDBField.FieldTypes.cftLong)
          .Add("DeliveryAddressNumber", CDBField.FieldTypes.cftLong)
          .Add("VatRate")
          .Add("VATPercentage", CDBField.FieldTypes.cftNumeric)
          .Add("SalesContactNumber", CDBField.FieldTypes.cftLong)
          .Add("SalesLedgerAccount")
          .Add("Issued")
          .Add("StockSale")
          .Add("Notes")
          .Add("FinancialAdjustment")
          .Add("WarehouseCode")
          .Add("ProductNumber", CDBField.FieldTypes.cftLong)
          .Add("GiverContactNumber", CDBField.FieldTypes.cftLong)
          .Add("ScheduledPaymentNumber", CDBField.FieldTypes.cftLong)
          .Add("ProvisionalBatchNumber", CDBField.FieldTypes.cftLong)
          .Add("ProvisionalTransactionNumber", CDBField.FieldTypes.cftLong)
          .Add("ProvisionalLineNumber", CDBField.FieldTypes.cftLong)
          .Add("DeceasedContactNumber", CDBField.FieldTypes.cftLong)
          .Add("PaymentPlanType")
          .Add("StockMovementNumbers")
          .Add("MemberNumber", CDBField.FieldTypes.cftLong)
          .Add("InvoiceNumber", CDBField.FieldTypes.cftLong)
          .Add("InvoiceNumberUsed", CDBField.FieldTypes.cftLong)
          .Add("EventBookingNumber", CDBField.FieldTypes.cftLong)
          .Add("RoomBookingNumber", CDBField.FieldTypes.cftLong)
          .Add("ServiceBookingNumber", CDBField.FieldTypes.cftLong)
          .Add("LegacyNumber", CDBField.FieldTypes.cftLong)
          .Add("BequestNumber", CDBField.FieldTypes.cftLong)
          .Add("LegacyReceiptNumber", CDBField.FieldTypes.cftLong)
          .Add("CovenantNumber", CDBField.FieldTypes.cftLong)
          .Add("CancellationReason")
          .Add("DeclarationNumber", CDBField.FieldTypes.cftLong)
          .Add("PGDonorID")
          .Add("ContactStatus")
          .Add("PostagePacking")
          .Add("ActivityGroup")
          .Add("Suppression")
          .Add("AutoPaymentNumber", CDBField.FieldTypes.cftLong)
          .Add("IncentiveLineNumber", CDBField.FieldTypes.cftLong)
          .Add("VATAmount", CDBField.FieldTypes.cftNumeric)
          .Add("PGPledgeNumber", CDBField.FieldTypes.cftLong)
          .Add("CollectionNumber", CDBField.FieldTypes.cftLong)
          .Add("CollectionPISNumber", CDBField.FieldTypes.cftLong)
          .Add("CollectionBankAccount")
          .Add("CollectionBoxNumbers")
          .Add("CollectionBoxAmounts")
          .Add("StockTransactionID", CDBField.FieldTypes.cftLong)
          .Add("PriceVATExclusive", CDBField.FieldTypes.cftCharacter)
          .Add("EventNumber", CDBField.FieldTypes.cftLong)
          .Add("BookingOptionNumber", CDBField.FieldTypes.cftLong)
          .Add("AdultQuantity", CDBField.FieldTypes.cftLong)
          .Add("ChildQuantity", CDBField.FieldTypes.cftLong)
          .Add("StartTime", CDBField.FieldTypes.cftTime)
          .Add("EndTime", CDBField.FieldTypes.cftTime)
          .Add("AmendedEventBookingNumber", CDBField.FieldTypes.cftLong)
          .Add("CreditedContactNumber", CDBField.FieldTypes.cftLong)
          .Add("InvoiceTypeUsed")
          .Add("DepositAllowed")
          .Add("ExamBookingId")
          .Add("ExamUnitId")
          .Add("ExamUnitProductId")
        End With
      Else
        mvClassFields.ClearItems()
      End If

    End Sub

    Public Function GetDataAsParameters() As CDBParameters
      Dim vParams As New CDBParameters
      Dim vField As ClassField

      For Each vField In mvClassFields
        If vField.Name = "SalesContactNumber" Then
          vParams.Add((vField.Name), (vField.FieldType), If(vField.IntegerValue > 0, vField.Value, "")) 'Do not output zero as it is invalid
        Else
          vParams.Add((vField.Name), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(vField.DoubleValue), vField.Value))
        End If
      Next vField
      GetDataAsParameters = vParams
    End Function

    Public Function GetAnalysisLineTypeCode(ByVal pAnalysisLineType As TraderAnalysisLineTypes) As String
      Dim vCode As String = ""
      Select Case pAnalysisLineType
        Case TraderAnalysisLineTypes.taltAccomodation
          vCode = "A"
        Case TraderAnalysisLineTypes.taltActivityEntry
          vCode = "AA"
        Case TraderAnalysisLineTypes.taltAddressUpdate
          vCode = "ADDR"
        Case TraderAnalysisLineTypes.taltAddSuppression
          vCode = "AS"
        Case TraderAnalysisLineTypes.taltLegacyBequestReceipt
          vCode = "B"
        Case TraderAnalysisLineTypes.taltCovenant
          vCode = "C"
        Case TraderAnalysisLineTypes.taltCreditCardAuthority
          vCode = "CC"
        Case TraderAnalysisLineTypes.taltCancelGiftAidDeclaration
          vCode = "CG"
        Case TraderAnalysisLineTypes.taltCreditCardAuthorityUpdate
          vCode = "CCU"
        Case TraderAnalysisLineTypes.taltCancelPaymentPlan
          vCode = "CP"
        Case TraderAnalysisLineTypes.taltDirectDebit
          vCode = "DD"
        Case TraderAnalysisLineTypes.taltDirectDebitUpdate
          vCode = "DDU"
        Case TraderAnalysisLineTypes.taltEvent
          vCode = "E"
        Case TraderAnalysisLineTypes.taltDeceased
          vCode = "G"
        Case TraderAnalysisLineTypes.taltGoneAway
          vCode = "GA"
        Case TraderAnalysisLineTypes.taltGiftAidDeclaration
          vCode = "GD"
        Case TraderAnalysisLineTypes.taltPayrollGivingPledge
          vCode = "GP"
        Case TraderAnalysisLineTypes.taltHardCredit
          vCode = "H"
        Case TraderAnalysisLineTypes.taltIncentive
          vCode = "I"
        Case TraderAnalysisLineTypes.taltInMemoriamHardCredit
          vCode = "D"
        Case TraderAnalysisLineTypes.taltInMemoriamSoftCredit
          vCode = "F"
        Case TraderAnalysisLineTypes.taltInvoiceAllocation
          vCode = "L"
        Case TraderAnalysisLineTypes.taltMembership
          vCode = "M"
        Case TraderAnalysisLineTypes.taltInvoicePayment
          vCode = "N"
        Case TraderAnalysisLineTypes.taltNoPayment
          vCode = "NP"
        Case TraderAnalysisLineTypes.taltPaymentPlan
          vCode = "O"
        Case TraderAnalysisLineTypes.taltProductSale
          vCode = "P"
        Case TraderAnalysisLineTypes.taltSundryCreditNote
          vCode = "R"
        Case TraderAnalysisLineTypes.taltSoftCredit
          vCode = "S"
        Case TraderAnalysisLineTypes.taltStandingOrder
          vCode = "SO"
        Case TraderAnalysisLineTypes.taltStandingOrderUpdate
          vCode = "SOU"
        Case TraderAnalysisLineTypes.taltStatus
          vCode = "ST"
        Case TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash
          vCode = "U"
        Case TraderAnalysisLineTypes.taltServiceBooking
          vCode = "V"
        Case TraderAnalysisLineTypes.taltServiceBookingCredit
          vCode = "VC"
        Case TraderAnalysisLineTypes.taltServiceBookingEntitlement
          vCode = "VE"
        Case TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment
          vCode = "PG"
        Case TraderAnalysisLineTypes.taltCollectionPayment
          vCode = "AP"
        Case TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
          vCode = "PP"
        Case TraderAnalysisLineTypes.taltEventPricingMatrixLine
          vCode = "X"
        Case TraderAnalysisLineTypes.taltExamBooking
          vCode = "Q"
        Case TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
          vCode = "K"
      End Select
      GetAnalysisLineTypeCode = vCode
    End Function

    Public Function GetAnalysisLineTypeFromCode(ByVal pAnalysisLineTypeCode As String) As TraderAnalysisLineTypes
      Dim vAnalysisLineType As TraderAnalysisLineTypes

      Select Case pAnalysisLineTypeCode
        Case "A"
          vAnalysisLineType = TraderAnalysisLineTypes.taltAccomodation
        Case "AA"
          vAnalysisLineType = TraderAnalysisLineTypes.taltActivityEntry
        Case "ADDR"
          vAnalysisLineType = TraderAnalysisLineTypes.taltAddressUpdate
        Case "AP"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCollectionPayment
        Case "AS"
          vAnalysisLineType = TraderAnalysisLineTypes.taltAddSuppression
        Case "B"
          vAnalysisLineType = TraderAnalysisLineTypes.taltLegacyBequestReceipt
        Case "C"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCovenant
        Case "CC"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCreditCardAuthority
        Case "CCU"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCreditCardAuthorityUpdate
        Case "CG"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCancelGiftAidDeclaration
        Case "CP"
          vAnalysisLineType = TraderAnalysisLineTypes.taltCancelPaymentPlan
        Case "D"
          vAnalysisLineType = TraderAnalysisLineTypes.taltInMemoriamHardCredit
        Case "DD"
          vAnalysisLineType = TraderAnalysisLineTypes.taltDirectDebit
        Case "DDU"
          vAnalysisLineType = TraderAnalysisLineTypes.taltDirectDebitUpdate
        Case "E"
          vAnalysisLineType = TraderAnalysisLineTypes.taltEvent
        Case "F"
          vAnalysisLineType = TraderAnalysisLineTypes.taltInMemoriamSoftCredit
        Case "G"
          vAnalysisLineType = TraderAnalysisLineTypes.taltDeceased
        Case "GA"
          vAnalysisLineType = TraderAnalysisLineTypes.taltGoneAway
        Case "GD"
          vAnalysisLineType = TraderAnalysisLineTypes.taltGiftAidDeclaration
        Case "GP"
          vAnalysisLineType = TraderAnalysisLineTypes.taltPayrollGivingPledge
        Case "H"
          vAnalysisLineType = TraderAnalysisLineTypes.taltHardCredit
        Case "I"
          vAnalysisLineType = TraderAnalysisLineTypes.taltIncentive
        Case "K"
          vAnalysisLineType = TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
        Case "L"
          vAnalysisLineType = TraderAnalysisLineTypes.taltInvoiceAllocation
        Case "M"
          vAnalysisLineType = TraderAnalysisLineTypes.taltMembership
        Case "N"
          vAnalysisLineType = TraderAnalysisLineTypes.taltInvoicePayment
        Case "NP"
          vAnalysisLineType = TraderAnalysisLineTypes.taltNoPayment
        Case "O"
          vAnalysisLineType = TraderAnalysisLineTypes.taltPaymentPlan
        Case "P"
          vAnalysisLineType = TraderAnalysisLineTypes.taltProductSale
        Case "PG"
          vAnalysisLineType = TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment
        Case "PP"
          vAnalysisLineType = TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
        Case "Q"
          vAnalysisLineType = TraderAnalysisLineTypes.taltExamBooking
        Case "R"
          vAnalysisLineType = TraderAnalysisLineTypes.taltSundryCreditNote
        Case "S"
          vAnalysisLineType = TraderAnalysisLineTypes.taltSoftCredit
        Case "SO"
          vAnalysisLineType = TraderAnalysisLineTypes.taltStandingOrder
        Case "SOU"
          vAnalysisLineType = TraderAnalysisLineTypes.taltStandingOrderUpdate
        Case "ST"
          vAnalysisLineType = TraderAnalysisLineTypes.taltStatus
        Case "U"
          vAnalysisLineType = TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash
        Case "V"
          vAnalysisLineType = TraderAnalysisLineTypes.taltServiceBooking
        Case "VC"
          vAnalysisLineType = TraderAnalysisLineTypes.taltServiceBookingCredit
        Case "VE"
          vAnalysisLineType = TraderAnalysisLineTypes.taltServiceBookingEntitlement
        Case "X"
          vAnalysisLineType = TraderAnalysisLineTypes.taltEventPricingMatrixLine
      End Select
      GetAnalysisLineTypeFromCode = vAnalysisLineType
    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Sub Init(ByVal pLineNumber As Integer, Optional ByVal pTransactionType As String = "")
      InitClassFields()
      mvClassFields.Item(TraderAnalysisLineFields.talfLineNumber).Value = CStr(pLineNumber)
      If pTransactionType.Length > 0 Then mvClassFields.Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = pTransactionType
    End Sub

    Public Sub InitFromBTA(ByVal pBTA As BatchTransactionAnalysis, ByVal pLineType As TraderAnalysisLineTypes, ByVal pFinancialAdjustment As Batch.AdjustmentTypes, ByVal pPayerContactNumber As Integer, ByVal pPayerAddressNumber As Integer, ByVal pBatchCurrencyCode As String, ByVal pTransDate As String, ByVal pStockMovementNumbers As String, ByVal pStockTransactionID As Integer)
      'Setup from an existing BTA
      Dim vLineType As TraderAnalysisLineTypes

      InitClassFields()
      vLineType = pLineType

      With mvClassFields
        'Check for additional types
        Select Case pBTA.AnalysisAdditionalType
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBooking
            vLineType = TraderAnalysisLineTypes.taltEvent
            .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = CStr(pBTA.AdditionalNumber)
            .Item(TraderAnalysisLineFields.talfEventNumber).Value = CStr(pBTA.AdditionalNumber2)
            .Item(TraderAnalysisLineFields.talfBookingOptionNumber).Value = CStr(pBTA.AdditionalNumber3)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatExamBooking
            vLineType = TraderAnalysisLineTypes.taltExamBooking
            .Item(TraderAnalysisLineFields.talfExamBookingNumber).Value = CStr(pBTA.AdditionalNumber)
            .Item(TraderAnalysisLineFields.talfExamUnitId).Value = CStr(pBTA.AdditionalNumber2)
            .Item(TraderAnalysisLineFields.talfExamUnitProductId).Value = CStr(pBTA.AdditionalNumber3)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatExamBookingTransaction
            vLineType = TraderAnalysisLineTypes.taltExamBooking
            .Item(TraderAnalysisLineFields.talfExamBookingNumber).Value = CStr(pBTA.AdditionalNumber)
            .Item(TraderAnalysisLineFields.talfExamUnitId).Value = CStr(pBTA.AdditionalNumber2)
            .Item(TraderAnalysisLineFields.talfExamUnitProductId).Value = CStr(pBTA.AdditionalNumber3)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatAccommodationBooking
            vLineType = TraderAnalysisLineTypes.taltAccomodation
            .Item(TraderAnalysisLineFields.talfRoomBookingNumber).Value = CStr(pBTA.AdditionalNumber)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBooking, BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBookingTransaction
            If pBTA.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBooking Then vLineType = TraderAnalysisLineTypes.taltServiceBooking
            .Item(TraderAnalysisLineFields.talfServiceBookingNumber).Value = CStr(pBTA.AdditionalNumber)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBookingCredit
            vLineType = TraderAnalysisLineTypes.taltServiceBookingCredit
            .Item(TraderAnalysisLineFields.talfServiceBookingNumber).Value = CStr(pBTA.AdditionalNumber)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatServiceBookingEntitlement
            vLineType = TraderAnalysisLineTypes.taltServiceBookingEntitlement
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatLegacyBequest
            vLineType = TraderAnalysisLineTypes.taltLegacyBequestReceipt
            .Item(TraderAnalysisLineFields.talfLegacyNumber).Value = CStr(pBTA.AdditionalNumber2)
            .Item(TraderAnalysisLineFields.talfBequestNumber).Value = CStr(pBTA.AdditionalNumber)
            .Item(TraderAnalysisLineFields.talfLegacyReceiptNumber).Value = CStr(pBTA.AdditionalNumber3)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatCollectionPayments
            vLineType = TraderAnalysisLineTypes.taltCollectionPayment
            .Item(TraderAnalysisLineFields.talfCollectionNumber).Value = CStr(pBTA.AdditionalNumber)
            .Item(TraderAnalysisLineFields.talfCollectionPISNumber).Value = CStr(pBTA.AdditionalNumber2)
            .Item(TraderAnalysisLineFields.talfCollectionBoxNumbers).Value = pBTA.MemberNumber
            If pBTA.DeceasedContactNumber > 0 Then .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = CStr(pBTA.DeceasedContactNumber)
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment
            .Item(TraderAnalysisLineFields.talfScheduledPaymentNumber).IntegerValue = pBTA.AdditionalNumber
          Case BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatInvoicePaymentHistory
            Select Case vLineType
              Case TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
                .Item(TraderAnalysisLineFields.talfInvoiceNumberUsed).Value = pBTA.AdditionalNumber.ToString
            End Select
        End Select

        'Set up specific fields according to the LineType
        Select Case vLineType
          Case TraderAnalysisLineTypes.taltDeceased, TraderAnalysisLineTypes.taltSoftCredit, TraderAnalysisLineTypes.taltHardCredit
            If pBTA.PaymentPlanNumber > 0 Then
              .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = CStr(pBTA.DeceasedContactNumber)
              If vLineType = TraderAnalysisLineTypes.taltHardCredit Then .Item(TraderAnalysisLineFields.talfGiverContactNumber).Value = CStr(pBTA.DeceasedContactNumber)
              If pBTA.MemberNumber.Length > 0 Then
                .Item(TraderAnalysisLineFields.talfMemberNumber).Value = pBTA.MemberNumber
                .Item(TraderAnalysisLineFields.talfPaymentPlanType).Value = "M"
              ElseIf pBTA.CovenantNumber > 0 Then
                .Item(TraderAnalysisLineFields.talfCovenantNumber).Value = CStr(pBTA.CovenantNumber)
                .Item(TraderAnalysisLineFields.talfPaymentPlanType).Value = "C"
              Else
                .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pBTA.PaymentPlanNumber)
                .Item(TraderAnalysisLineFields.talfPaymentPlanType).Value = "O"
              End If
            Else
              .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = CStr(pBTA.DeceasedContactNumber)
            End If
            If vLineType = TraderAnalysisLineTypes.taltHardCredit And pFinancialAdjustment = Batch.AdjustmentTypes.atGIKConfirmation Then
              .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).IntegerValue = pPayerContactNumber
            End If
          Case TraderAnalysisLineTypes.taltInMemoriamHardCredit, TraderAnalysisLineTypes.taltInMemoriamSoftCredit
            .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = pBTA.DeceasedContactNumber.ToString
          Case TraderAnalysisLineTypes.taltMembership
            .Item(TraderAnalysisLineFields.talfMemberNumber).Value = pBTA.MemberNumber
            .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pBTA.PaymentPlanNumber)
          Case TraderAnalysisLineTypes.taltCovenant
            .Item(TraderAnalysisLineFields.talfCovenantNumber).Value = CStr(pBTA.CovenantNumber)
            .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pBTA.PaymentPlanNumber)
          Case TraderAnalysisLineTypes.taltPaymentPlan
            .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pBTA.PaymentPlanNumber)
          Case TraderAnalysisLineTypes.taltInvoicePayment, TraderAnalysisLineTypes.taltInvoiceAllocation, TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash, TraderAnalysisLineTypes.taltSundryCreditNote, _
               TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
            .Item(TraderAnalysisLineFields.talfInvoiceNumber).Value = CStr(pBTA.InvoiceNumber)
            .Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value = pBTA.MemberNumber
          Case TraderAnalysisLineTypes.taltIncentive
            .Item(TraderAnalysisLineFields.talfIncentiveLineNumber).Value = CStr(pBTA.PaymentPlanNumber)
          Case TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment, TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment
            If Len(pBTA.MemberNumber) > 0 Then .Item(TraderAnalysisLineFields.talfPGPledgeNumber).Value = pBTA.MemberNumber
          Case TraderAnalysisLineTypes.taltEventPricingMatrixLine
            If pBTA.MemberNumber.Length > 0 Then .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = pBTA.MemberNumber
        End Select

        'Set up the remaining fields
        .Item(TraderAnalysisLineFields.talfLineNumber).Value = CStr(pBTA.LineNumber)
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(vLineType)
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = pBTA.LineType
        If pBTA.ProductCode.Length > 0 Then
          If pBTA.Product.PostagePacking Then
            .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = "P&P"
            .Item(TraderAnalysisLineFields.talfPostagePacking).Bool = True
          End If
          If pBTA.LineType <> "I" Then
            .Item(TraderAnalysisLineFields.talfStockSale).Bool = pBTA.Product.StockItem
            If pBTA.Product.StockItem = True Then
              If Len(pStockMovementNumbers) > 0 Then .Item(TraderAnalysisLineFields.talfStockMovementNumbers).Value = pStockMovementNumbers 'Used by Rich Client
              If pStockTransactionID > 0 Then .Item(TraderAnalysisLineFields.talfStockTransactionID).Value = CStr(pStockTransactionID) 'Used by Smart Client & Web Services
            End If
          End If
        End If
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pBTA.ProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pBTA.RateCode
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pBTA.DistributionCode
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pBTA.Quantity)
        .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pBTA.Issued)
        .Item(TraderAnalysisLineFields.talfNotes).Value = pBTA.Notes
        .Item(TraderAnalysisLineFields.talfSource).Value = pBTA.Source
        If pBatchCurrencyCode.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pBTA.CurrencyAmount)
        Else
          .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pBTA.Amount)
        End If
        If pBTA.GrossAmount.Length > 0 Then .Item(TraderAnalysisLineFields.talfGrossAmount).Value = pBTA.GrossAmount
        If pBTA.Discount.Length > 0 Then .Item(TraderAnalysisLineFields.talfDiscount).Value = pBTA.Discount
        .Item(TraderAnalysisLineFields.talfAcceptAsFull).Bool = pBTA.AcceptAsFull
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pBTA.WhenValue
        .Item(TraderAnalysisLineFields.talfDespatchMethod).Value = pBTA.DespatchMethod
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pBTA.ContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pBTA.AddressNumber)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pBTA.VatRate
        If pBTA.VatRate.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfVATPercentage).Value = pBTA.GetVATPercentage(pTransDate)
          .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(pBTA.VatAmount)
        End If
        .Item(TraderAnalysisLineFields.talfWarehouseCode).Value = pBTA.Warehouse
        .Item(TraderAnalysisLineFields.talfProductNumber).Value = pBTA.ProductNumber.ToString
        If pBTA.AnalysisAdditionalType <> BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatFundraisingPayment Then
          .Item(TraderAnalysisLineFields.talfScheduledPaymentNumber).Value = pBTA.ScheduledPaymentNumber
        End If
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = CStr(pBTA.SalesContactNumber)
        Select Case pLineType
          Case TraderAnalysisLineTypes.taltInMemoriamHardCredit, TraderAnalysisLineTypes.taltInMemoriamSoftCredit
            If pBTA.DeceasedContactNumber > 0 Then
              .Item(TraderAnalysisLineFields.talfCreditedContactNumber).Value = pBTA.ContactNumber.ToString
              .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = pPayerContactNumber.ToString
              .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = pPayerAddressNumber.ToString
            End If
          Case TraderAnalysisLineTypes.taltProductSale
            If pBTA.AnalysisAdditionalType = BatchTransactionAnalysis.TransactionAnalysisAdditionalTypes.taatEventBookingTransaction AndAlso pBTA.TransactionContainsEventBooking = False Then .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = pBTA.AdditionalNumber2.ToString
        End Select
      End With

    End Sub

    Public Sub AddAccomodationBooking(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pRoomBookingNumber As Integer, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltAccomodation)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfRoomBookingNumber).Value = CStr(pRoomBookingNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
      End With

    End Sub

    Public Sub AddActivity(ByVal pContactNumber As Integer, ByVal pActivityGroup As String)
      'Show the Activities added
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltActivityEntry)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfActivityGroup).Value = pActivityGroup
      End With
    End Sub

    Public Sub AddAddressUpdate(ByVal pContactNumber As Integer)
      'Show the Contacts' addresses updated
      mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltAddressUpdate)
      mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
    End Sub

    Public Sub AddAutoPaymentUpdate(ByVal pType As TraderAnalysisLineTypes, ByVal pPaymentPlanNumber As Integer, ByVal pAutoPaymentNumber As Integer)
      'Show the Auto Payment Method updated
      If pType = TraderAnalysisLineTypes.taltDirectDebitUpdate Or pType = TraderAnalysisLineTypes.taltStandingOrderUpdate Or pType = TraderAnalysisLineTypes.taltCreditCardAuthorityUpdate Then
        With mvClassFields
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(pType)
          .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pPaymentPlanNumber)
          .Item(TraderAnalysisLineFields.talfAutoPaymentNumber).Value = CStr(pAutoPaymentNumber)
        End With
      End If
    End Sub

    Public Sub AddCollectionPayment(ByVal pCollectionNumber As Integer, ByVal pProductCode As String, ByVal pRate As String, ByVal pSource As String, ByVal pAmount As Double, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pCollectionBankAccount As String, ByVal pPriceVATExclusive As Boolean, Optional ByVal pNotes As String = "", Optional ByVal pCollectionPISNumber As Integer = 0, Optional ByVal pDeceasedContactNumber As Integer = 0, Optional ByVal pCollectionBoxNumbers As String = "", Optional ByVal pCollectionBoxAmounts As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltCollectionPayment)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(1)
        .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(1)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfCollectionNumber).Value = CStr(pCollectionNumber)
        .Item(TraderAnalysisLineFields.talfCollectionBankAccount).Value = pCollectionBankAccount
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
        If pCollectionPISNumber > 0 Then .Item(TraderAnalysisLineFields.talfCollectionPISNumber).Value = CStr(pCollectionPISNumber)
        If pDeceasedContactNumber > 0 Then .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = CStr(pDeceasedContactNumber)
        If Len(pCollectionBoxNumbers) > 0 Then .Item(TraderAnalysisLineFields.talfCollectionBoxNumbers).Value = pCollectionBoxNumbers
        If Len(pCollectionBoxAmounts) > 0 Then
          .Item(TraderAnalysisLineFields.talfCollectionBoxAmounts).Value = pCollectionBoxAmounts
        Else
          .Item(TraderAnalysisLineFields.talfCollectionBoxAmounts).Value = CStr(pAmount)
        End If
      End With

    End Sub

    Public Sub AddContactStatus(ByVal pContactNumber As Integer, Optional ByVal pStatus As String = "", Optional ByVal pDate As String = "")
      'Show the Contact Status changed
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltStatus)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfContactStatus).Value = pStatus
        If pStatus.Length > 0 Then .Item(TraderAnalysisLineFields.talfLineDate).Value = pDate
      End With
    End Sub

    Public Sub AddEventBooking(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pEventBookingNumber As Integer, ByVal pPriveVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pNotes As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltEvent)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = CStr(pEventBookingNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriveVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
      End With

    End Sub

    Public Sub AddEventBookingPriceMatrixLine(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pVATRate As String, ByVal pVATAmount As Double, ByVal pVATPercentage As Double, ByVal pSource As String, ByVal pNotes As String, ByVal pEventBookingNumber As Integer, ByVal pDistributionCode As String)
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltEventPricingMatrixLine)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(pVATAmount)
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercentage)
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = True 'Price Matrix fees always exclude VAT
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
        .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = CStr(pEventBookingNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
      End With
    End Sub

    Public Sub AddExamBookingLine(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pExamBookingNumber As Integer, ByVal pExamUnitId As String, ByVal pExamUnitProductId As String, ByVal pPriveVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pNotes As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltExamBooking)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfExamBookingNumber).Value = CStr(pExamBookingNumber)
        .Item(TraderAnalysisLineFields.talfExamUnitId).Value = CStr(pExamUnitId)
        .Item(TraderAnalysisLineFields.talfExamUnitProductId).Value = CStr(pExamUnitProductId)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriveVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
      End With

    End Sub

    Public Sub AddGiftAidDeclaration(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pStartDate As String, ByVal pSource As String, Optional ByVal pNotes As String = "")
      'Show the Gift Aid Declaration added
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltGiftAidDeclaration)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pAddressNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pStartDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
      End With

    End Sub

    Public Sub AddGiftAidDeclarationCancellation(ByVal pDeclarationNumber As Integer, ByVal pContactNumber As Integer, ByVal pCancellationReason As String, ByVal pCancellationDate As String, Optional ByVal pCancellationSource As String = "")
      'Show the Gift Aid Declaration cancelled
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltCancelGiftAidDeclaration)
        .Item(TraderAnalysisLineFields.talfDeclarationNumber).Value = CStr(pDeclarationNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfCancellationReason).Value = pCancellationReason
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pCancellationDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pCancellationSource
      End With

    End Sub

    Public Sub AddGoneAway(ByVal pContactNumber As Integer)
      'Show Contact has GA status added
      mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltGoneAway)
      mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
    End Sub

    Public Sub AddIncentive(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pSource As String, ByVal pDate As String, ByVal pDespatchMethod As String, ByVal pDeliveryContactNumber As Integer, ByVal pDeliveryAddressNumber As Integer, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pIncentiveLineNumber As Integer)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltIncentive)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfAmount).Value = "0"
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pDate
        .Item(TraderAnalysisLineFields.talfDespatchMethod).Value = pDespatchMethod
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pDeliveryContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pDeliveryAddressNumber)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfIncentiveLineNumber).Value = CStr(pIncentiveLineNumber)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(0)
      End With

    End Sub

    Public Sub AddInvoicePayment(ByVal pInvoiceNumber As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pSalesLedgerAccount As String, Optional ByVal pCashInvoiceNumber As Integer = 0, Optional ByVal pInvoiceNumberUsed As Integer = 0, Optional ByVal pDistributionCode As String = "", Optional ByVal pContactNumber As Integer = 0, Optional ByVal pAddressNumber As Integer = 0, Optional ByVal pInvoiceTypeUsed As String = "")
      Dim vLineType As TraderAnalysisLineTypes
      Dim vTransType As String

      If pCashInvoiceNumber > 0 Then
        If pInvoiceTypeUsed = "N" Then
          'BR16409: For 'N' type Sundry Credit Notes set SundryCreditNoteInvoiceAllocation 'K' trader line type
          vLineType = TraderAnalysisLineTypes.taltSundryCreditNoteInvoiceAllocation
        Else
          vLineType = TraderAnalysisLineTypes.taltInvoiceAllocation
        End If
        vTransType = "ALL"
      Else
        vLineType = TraderAnalysisLineTypes.taltInvoicePayment
        vTransType = "INV"
      End If

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(vLineType)
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = vTransType
        .Item(TraderAnalysisLineFields.talfInvoiceNumber).Value = CStr(pInvoiceNumber)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value = pSalesLedgerAccount
        .Item(TraderAnalysisLineFields.talfQuantity).Value = "1"
        'Optional
        If pInvoiceNumberUsed > 0 Then .Item(TraderAnalysisLineFields.talfInvoiceNumberUsed).Value = CStr(pInvoiceNumberUsed)
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        'Need to store the Credit Customer's Contact and Address Numbers on the BTA as the payment could be from someone other than the customer.
        If pContactNumber > 0 Then
          Debug.Assert(pAddressNumber > 0)
          .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = pContactNumber.ToString
          .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = pAddressNumber.ToString
        End If
        If pInvoiceTypeUsed.Length > 0 Then .Item(TraderAnalysisLineFields.talfInvoiceTypeUsed).Value = pInvoiceTypeUsed
      End With

    End Sub

    Public Sub AddLegacyBequestReceipt(ByVal pLegacyNumber As Integer, ByVal pBequestNumber As Integer, ByVal pProductCode As String, ByVal pRate As String, ByVal pAmount As Double, ByVal pDate As String, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, Optional ByVal pNotes As String = "", Optional ByVal pDistributionCode As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltLegacyBequestReceipt)
        .Item(TraderAnalysisLineFields.talfLegacyNumber).Value = CStr(pLegacyNumber)
        .Item(TraderAnalysisLineFields.talfBequestNumber).Value = CStr(pBequestNumber)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = "1"
        .Item(TraderAnalysisLineFields.talfIssued).Value = "1"
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pAddressNumber)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        'Optional
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
      End With

    End Sub

    Public Sub AddNonPaymentLine(ByVal pPaymentPlanNumber As Integer, ByVal pSource As String, Optional ByVal pLineType As String = "", Optional ByVal pPaymentNumber As String = "")
      mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltNoPayment)
      mvClassFields.Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pPaymentPlanNumber)
      mvClassFields.Item(TraderAnalysisLineFields.talfSource).Value = pSource
      If pLineType.Length > 0 Then
        If pPaymentNumber.Length > 0 Then
          'Set the PaymentNumber but do not change the LineType
          Select Case GetAnalysisLineTypeFromCode(pLineType)
            Case TraderAnalysisLineTypes.taltMembership
              mvClassFields.Item(TraderAnalysisLineFields.talfMemberNumber).Value = pPaymentNumber
            Case TraderAnalysisLineTypes.taltCovenant
              mvClassFields.Item(TraderAnalysisLineFields.talfCovenantNumber).Value = pPaymentNumber
          End Select
        Else
          'There is no PaymentNumber so change the LineType
          mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value = pLineType
        End If
      End If
    End Sub

    Public Sub AddNonStockProductSale(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pDate As String, ByVal pSource As String, ByVal pDeliveryContactNumber As Integer, ByVal pDeliveryAddressNumber As Integer, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pVATAmount As Double, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDespatchMethod As String = "", Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pNotes As String = "", Optional ByVal pProductNumber As String = "", Optional ByVal pContactDiscount As Boolean = False, Optional ByVal pGrossAmount As Double = 0, Optional ByVal pDiscount As Double = 0, Optional ByVal pDeceasedContactNumber As String = "", Optional ByVal pDeceasedLineTypeCode As String = "", Optional ByVal pSalesLedgerAcount As String = "", Optional ByVal pServiceBookingNumber As Integer = 0, Optional ByVal pEventBookingNumber As Integer = 0, _
                                      Optional ByVal pFundScheduledPaymentNumber As Integer = 0, Optional ByVal pCreditedContactNumber As String = "")
      'Non-stock Products only

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltProductSale)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pDeliveryContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pDeliveryAddressNumber)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(pVATAmount)
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDespatchMethod).Value = pDespatchMethod
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
        .Item(TraderAnalysisLineFields.talfProductNumber).Value = pProductNumber
        If pContactDiscount Then
          .Item(TraderAnalysisLineFields.talfGrossAmount).Value = CStr(pGrossAmount)
          .Item(TraderAnalysisLineFields.talfDiscount).Value = CStr(pDiscount)
        End If
        If pDeceasedContactNumber.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = pDeceasedContactNumber
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = pDeceasedLineTypeCode 'D, F, G, H or S - depending upon whether it is a Hard Credit, Soft Credit etc.
          If TraderLineType = TraderAnalysisLineTypes.taltInMemoriamHardCredit OrElse TraderLineType = TraderAnalysisLineTypes.taltInMemoriamSoftCredit Then
            .Item(TraderAnalysisLineFields.talfCreditedContactNumber).Value = pCreditedContactNumber
          End If
        End If
        If pSalesLedgerAcount.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value = pSalesLedgerAcount
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltSundryCreditNote)
        End If
        If pServiceBookingNumber > 0 Then .Item(TraderAnalysisLineFields.talfServiceBookingNumber).IntegerValue = pServiceBookingNumber
        If pEventBookingNumber > 0 Then .Item(TraderAnalysisLineFields.talfEventBookingNumber).IntegerValue = pEventBookingNumber
        If pFundScheduledPaymentNumber > 0 Then .Item(TraderAnalysisLineFields.talfScheduledPaymentNumber).LongValue = pFundScheduledPaymentNumber
      End With

    End Sub

    Public Sub AddPaymentPlanCancellation(ByVal pPaymentPlanNumber As Integer, ByVal pCancellationReason As String, ByVal pCancellationDate As String, Optional ByVal pCancellationSource As String = "")
      'Show Payment Plan cancelled
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltCancelPaymentPlan)
        .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pPaymentPlanNumber)
        .Item(TraderAnalysisLineFields.talfCancellationReason).Value = pCancellationReason
        If Not (IsDate(pCancellationDate)) Then pCancellationDate = TodaysDate()
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pCancellationDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pCancellationSource
      End With

    End Sub

    Public Sub AddPaymentPlanPayment(ByVal pPaymentPlanNumber As Integer, ByVal pPaymentNumber As String, ByVal pScheduledPaymentNumber As Integer, ByVal pAmount As Double, ByVal pSource As String, Optional ByVal pAcceptAsFull As Boolean = False, Optional ByVal pDeceasedContactNumber As String = "", Optional ByVal pPayPlanType As String = "", Optional ByVal pGiverContactNumber As String = "", Optional ByVal pAdditionalLineTypeCode As String = "", Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value
        If pAdditionalLineTypeCode.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = pAdditionalLineTypeCode
          If pAdditionalLineTypeCode Like "[SGH]" Then
            .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = pAdditionalLineTypeCode
          End If
        End If
        .Item(TraderAnalysisLineFields.talfPaymentPlanNumber).Value = CStr(pPaymentPlanNumber)
        .Item(TraderAnalysisLineFields.talfScheduledPaymentNumber).Value = CStr(pScheduledPaymentNumber)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        If pAcceptAsFull Then .Item(TraderAnalysisLineFields.talfAcceptAsFull).Bool = True
        Select Case GetAnalysisLineTypeFromCode(.Item(TraderAnalysisLineFields.talfTraderLineType).Value)
          Case TraderAnalysisLineTypes.taltMembership
            .Item(TraderAnalysisLineFields.talfMemberNumber).Value = pPaymentNumber
          Case TraderAnalysisLineTypes.taltCovenant
            .Item(TraderAnalysisLineFields.talfCovenantNumber).Value = pPaymentNumber
        End Select
        If pDeceasedContactNumber.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = pDeceasedContactNumber
          .Item(TraderAnalysisLineFields.talfPaymentPlanType).Value = pPayPlanType
        End If
        If pGiverContactNumber.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltHardCredit)
          .Item(TraderAnalysisLineFields.talfGiverContactNumber).Value = pGiverContactNumber
        End If
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        If Val(pSalesContactNumber) > 0 Then .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
      End With

    End Sub

    Public Sub AddPayrollGivingPledge(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pStartDate As String, ByVal pAmount As Double, ByVal pDonorID As String, ByVal pProductCode As String, ByVal pRate As String, Optional ByVal pDistributionCode As String = "")
      'Show the Payroll Giving Pledge added
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltPayrollGivingPledge)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pAddressNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pStartDate
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfPGDonorID).Value = pDonorID
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
      End With

    End Sub

    Public Sub AddPostageAndPacking(ByVal pProductCode As String, ByVal pRate As String, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltProductSale)
        .Item(TraderAnalysisLineFields.talfPostagePacking).Bool = True
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = "P&P"
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = "1"
        .Item(TraderAnalysisLineFields.talfIssued).Value = "1"
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
      End With

    End Sub

    Public Sub AddPreTaxPGPayment(ByVal pProductCode As String, ByVal pRate As String, ByVal pAmount As Double, ByVal pSource As String, ByVal pDistributionCode As String, Optional ByVal pPledgeNumber As Integer = 0)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = "PP"
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltPreTaxPayrollGivingPayment)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = "1"
        .Item(TraderAnalysisLineFields.talfIssued).Value = "1"
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        If pPledgeNumber > 0 Then .Item(TraderAnalysisLineFields.talfPGPledgeNumber).Value = CStr(pPledgeNumber)
      End With

    End Sub

    Public Sub AddPostTaxPGPayment(ByVal pProductCode As String, ByVal pRate As String, ByVal pAmount As Double, ByVal pSource As String, ByVal pDistributionCode As String, Optional ByVal pPledgeNumber As Integer = 0)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = "PG"
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltPostTaxPayrollGivingPayment)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = "1"
        .Item(TraderAnalysisLineFields.talfIssued).Value = "1"
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        If pPledgeNumber > 0 Then .Item(TraderAnalysisLineFields.talfPGPledgeNumber).Value = CStr(pPledgeNumber)
      End With

    End Sub

    Public Sub AddStockProductSale(ByVal pProductCode As String, ByVal pRate As String, ByVal pWarehouseCode As String, ByVal pQuantity As Integer, ByVal pIssued As Integer, ByVal pAmount As Double, ByVal pDate As String, ByVal pSource As String, ByVal pDespatchMethod As String, ByVal pDeliveryContactNumber As Integer, ByVal pDeliveryAddressNumber As Integer, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pVATAmount As Double, ByVal pStockMovementNumbers As String, ByVal pVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pNotes As String = "", Optional ByVal pProductNumber As String = "", Optional ByVal pContactDiscount As Boolean = False, Optional ByVal pGrossAmount As Double = 0, Optional ByVal pDiscount As Double = 0, Optional ByVal pDeceasedContactNumber As String = "", Optional ByVal pDeceasedLineTypeCode As String = "", Optional ByVal pSalesLedgerAcount As String = "", Optional ByVal pStockTransactionID As Integer = 0, Optional ByVal pServiceBookingNumber As Integer = 0, Optional ByVal pEventBookingNumber As Integer = 0)
      'Stock Products Only

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltProductSale)
        .Item(TraderAnalysisLineFields.talfStockSale).Bool = True
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfWarehouseCode).Value = pWarehouseCode
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pIssued)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pDate
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfDespatchMethod).Value = pDespatchMethod
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pDeliveryContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pDeliveryAddressNumber)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfStockMovementNumbers).Value = pStockMovementNumbers '(Comma-separated list of StockMovementNumbers - Rich-client only)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(pVATAmount)
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        .Item(TraderAnalysisLineFields.talfNotes).Value = pNotes
        .Item(TraderAnalysisLineFields.talfProductNumber).Value = pProductNumber
        If pContactDiscount Then
          .Item(TraderAnalysisLineFields.talfGrossAmount).Value = CStr(pGrossAmount)
          .Item(TraderAnalysisLineFields.talfDiscount).Value = CStr(pDiscount)
        End If
        If pDeceasedContactNumber.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = pDeceasedContactNumber
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = pDeceasedLineTypeCode 'This code will depend upon whether it is a Hard Credit, Soft Credit etc.
        End If
        If pSalesLedgerAcount.Length > 0 Then
          .Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value = pSalesLedgerAcount
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltSundryCreditNote)
        End If
        If pStockTransactionID > 0 Then .Item(TraderAnalysisLineFields.talfStockTransactionID).Value = CStr(pStockTransactionID)
        If pServiceBookingNumber > 0 Then .Item(TraderAnalysisLineFields.talfServiceBookingNumber).IntegerValue = pServiceBookingNumber
        If pEventBookingNumber > 0 Then .Item(TraderAnalysisLineFields.talfEventBookingNumber).IntegerValue = pEventBookingNumber
      End With

    End Sub

    Public Sub AddServiceBooking(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pServiceBookingNumber As Integer, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pContactDiscount As Boolean = False, Optional ByVal pGrossAmount As Double = 0, Optional ByVal pDiscount As Double = 0)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltServiceBooking)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfServiceBookingNumber).Value = CStr(pServiceBookingNumber)
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        If pContactDiscount Then
          .Item(TraderAnalysisLineFields.talfGrossAmount).Value = CStr(pGrossAmount)
          .Item(TraderAnalysisLineFields.talfDiscount).Value = CStr(pDiscount)
        End If
      End With

    End Sub

    Public Sub AddServiceBookingCredit(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pServiceBookingNumber As Integer, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "", Optional ByVal pContactDiscount As Boolean = False, Optional ByVal pGrossAmount As Double = 0, Optional ByVal pDiscount As Double = 0)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltServiceBookingCredit)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(System.Math.Abs(pQuantity)) 'Could be negative
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(System.Math.Abs(pAmount)) 'Could be negative
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfServiceBookingNumber).Value = CStr(pServiceBookingNumber)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount((.Item(TraderAnalysisLineFields.talfAmount).DoubleValue), pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
        If pContactDiscount Then
          .Item(TraderAnalysisLineFields.talfGrossAmount).Value = CStr(System.Math.Abs(pGrossAmount)) 'Could be negative
          .Item(TraderAnalysisLineFields.talfDiscount).Value = CStr(System.Math.Abs(pDiscount)) 'Could be negative
        End If
      End With

    End Sub

    Public Sub AddServiceBookingEntitlementProduct(ByVal pProductCode As String, ByVal pRate As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pSource As String, ByVal pVATRate As String, ByVal pVATPercent As Double, ByVal pPriceVATExclusive As Boolean, Optional ByVal pDistributionCode As String = "", Optional ByVal pSalesContactNumber As String = "")

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltServiceBookingEntitlement)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pRate
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pVATRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfLineDate).Value = TodaysDate()
        If pQuantity > 0 Then .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pQuantity)
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(pAmount, pVATPercent))
        .Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = pPriceVATExclusive
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = pSalesContactNumber
      End With

    End Sub

    Public Sub AddSuppression(ByVal pContactNumber As Integer, ByVal pSuppression As String)
      'Show the Suppression added
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltAddSuppression)
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pContactNumber)
        .Item(TraderAnalysisLineFields.talfSuppression).Value = pSuppression
      End With

    End Sub

    Public Sub AddUnallocatedSalesledgerPayment(ByVal pAmount As Double, ByVal pSource As String, ByVal pSalesLedgerAccount As String, Optional ByVal pDistributionCode As String = "", Optional ByVal pContactNumber As Integer = 0, Optional ByVal pAddressNumber As Integer = 0)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltUnallocatedSalesLedgerCash)
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pAmount)
        .Item(TraderAnalysisLineFields.talfSource).Value = pSource
        .Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value = pSalesLedgerAccount
        'Optional
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pDistributionCode
        .Item(TraderAnalysisLineFields.talfTraderTransactionType).Value = "UNALL"
        'Need to store the Credit Customer's Contact and Address Numbers on the BTA as the payment could be from someone other than the customer.
        If pContactNumber > 0 Then
          Debug.Assert(pAddressNumber > 0)
          .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = pContactNumber.ToString
          .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = pAddressNumber.ToString
        End If
      End With

    End Sub

    Public Sub ConfirmProvisionalTransaction(ByVal pBTA As BatchTransactionAnalysis, ByVal pVATPercent As Double, Optional ByVal pContactDiscount As Boolean = False, Optional ByVal pStockSale As Boolean = False)

      With mvClassFields
        .Item(TraderAnalysisLineFields.talfTraderLineType).Value = GetAnalysisLineTypeCode(TraderAnalysisLineTypes.taltProductSale)
        .Item(TraderAnalysisLineFields.talfStockSale).Bool = pStockSale
        .Item(TraderAnalysisLineFields.talfProvisionalBatchNumber).Value = CStr(pBTA.BatchNumber)
        .Item(TraderAnalysisLineFields.talfProvisionalTransactionNumber).Value = CStr(pBTA.TransactionNumber)
        .Item(TraderAnalysisLineFields.talfProvisionalLineNumber).Value = CStr(pBTA.LineNumber)
        .Item(TraderAnalysisLineFields.talfProductCode).Value = pBTA.ProductCode
        .Item(TraderAnalysisLineFields.talfRate).Value = pBTA.RateCode
        .Item(TraderAnalysisLineFields.talfDistributionCode).Value = pBTA.DistributionCode
        .Item(TraderAnalysisLineFields.talfQuantity).Value = CStr(pBTA.Quantity)
        If pBTA.DeceasedContactNumber > 0 Then
          .Item(TraderAnalysisLineFields.talfDeceasedContactNumber).Value = CStr(pBTA.DeceasedContactNumber)
          .Item(TraderAnalysisLineFields.talfTraderLineType).Value = pBTA.LineType
        End If
        If pContactDiscount Then
          .Item(TraderAnalysisLineFields.talfGrossAmount).Value = pBTA.GrossAmount
          .Item(TraderAnalysisLineFields.talfDiscount).Value = pBTA.Discount
        End If
        .Item(TraderAnalysisLineFields.talfAmount).Value = CStr(pBTA.Amount)
        If pBTA.AcceptAsFull Then .Item(TraderAnalysisLineFields.talfAcceptAsFull).Bool = True
        .Item(TraderAnalysisLineFields.talfLineDate).Value = pBTA.WhenValue
        .Item(TraderAnalysisLineFields.talfDespatchMethod).Value = pBTA.DespatchMethod
        .Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pBTA.ContactNumber)
        .Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pBTA.AddressNumber)
        .Item(TraderAnalysisLineFields.talfVatRate).Value = pBTA.VatRate
        .Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercent)
        .Item(TraderAnalysisLineFields.talfSource).Value = pBTA.Source
        If pBTA.SalesContactNumber > 0 Then .Item(TraderAnalysisLineFields.talfSalesContactNumber).Value = CStr(pBTA.SalesContactNumber)
        .Item(TraderAnalysisLineFields.talfIssued).Value = CStr(pBTA.Quantity)
        .Item(TraderAnalysisLineFields.talfNotes).Value = pBTA.Notes
        .Item(TraderAnalysisLineFields.talfProductNumber).Value = pBTA.ProductNumber.ToString
        .Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(CalculateVATAmount(Val(CStr(pBTA.Amount)), Val(CStr(pVATPercent))))
      End With

    End Sub

    Friend Sub ResetIssuedForPP()
      If mvClassFields.Item(TraderAnalysisLineFields.talfPostagePacking).Bool = True Then
        mvClassFields.Item(TraderAnalysisLineFields.talfIssued).Value = CStr(0)
      End If
    End Sub

    Public Sub SetDeliveryContactAndAddress(ByVal pDeliveryContactNumber As Integer, ByVal pDeliveryAddressNumber As Integer)
      mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryContactNumber).Value = CStr(pDeliveryContactNumber)
      mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).Value = CStr(pDeliveryAddressNumber)
    End Sub

    Public Sub SetEventBookingInfo(ByVal pEventNumber As Integer, ByVal pOptionNumber As Integer, ByVal pAdultQuantity As String, ByVal pChildQuantity As String, ByVal pStartTime As String, ByVal pEndTime As String, ByVal pAmendedBookingNumber As Integer)
      With mvClassFields
        .Item(TraderAnalysisLineFields.talfEventNumber).Value = CStr(pEventNumber)
        .Item(TraderAnalysisLineFields.talfBookingOptionNumber).Value = CStr(pOptionNumber)
        If Len(pAdultQuantity) > 0 Then .Item(TraderAnalysisLineFields.talfAdultQuantity).Value = pAdultQuantity
        If Len(pChildQuantity) > 0 Then .Item(TraderAnalysisLineFields.talfChildQuantity).Value = pChildQuantity
        If Len(pStartTime) > 0 Then .Item(TraderAnalysisLineFields.talfStartTime).Value = pStartTime
        If Len(pEndTime) > 0 Then .Item(TraderAnalysisLineFields.talfEndTime).Value = pEndTime
        If pAmendedBookingNumber > 0 Then .Item(TraderAnalysisLineFields.talfAmendedEventBookingNumber).Value = CStr(pAmendedBookingNumber)
      End With
    End Sub

    Public Sub SetFinancialAdjustment()
      mvClassFields.Item(TraderAnalysisLineFields.talfFinancialAdjustment).Bool = True
    End Sub

    Friend Sub SetNewIncentiveLineNumber(ByVal pNewLineNumber As Integer)
      If TraderLineType = TraderAnalysisLineTypes.taltIncentive Then
        mvClassFields.Item(TraderAnalysisLineFields.talfIncentiveLineNumber).Value = CStr(pNewLineNumber)
      End If
    End Sub

    Friend Sub UpdateVAT(ByVal pVATPercentage As Double)
      'This is used when the TransactionDate in Trader was changed after the VAT had been calculated
      Dim vAmount As Double
      Dim vDiscountAmount As Double
      Dim vGrossAmount As Double
      Dim vVatAmount As Double

      If pVATPercentage <> mvClassFields.Item(TraderAnalysisLineFields.talfVATPercentage).DoubleValue Then
        vAmount = mvClassFields.Item(TraderAnalysisLineFields.talfAmount).DoubleValue
        vVatAmount = mvClassFields.Item(TraderAnalysisLineFields.talfVATAmount).DoubleValue

        If Len(mvClassFields.Item(TraderAnalysisLineFields.talfDiscount).Value) > 0 Then
          vDiscountAmount = mvClassFields.Item(TraderAnalysisLineFields.talfDiscount).DoubleValue
          vGrossAmount = mvClassFields.Item(TraderAnalysisLineFields.talfGrossAmount).DoubleValue
          If mvClassFields.Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = True Then
            'vAmount = Amount - VAT + Discount
            'vVAT = VAT on Amount - Discount
            'vGrossAmount = vAmount + VAT on vAmount
            vAmount = FixTwoPlaces(vAmount - vVatAmount + vDiscountAmount) 'Line price without VAT
            vVatAmount = FixTwoPlaces(vAmount * (pVATPercentage / 100)) 'VAT on line price
            vGrossAmount = FixTwoPlaces(vAmount + vVatAmount) 'Line price with VAT
            vVatAmount = FixTwoPlaces((vAmount - vDiscountAmount) * (pVATPercentage / 100)) 'VAT on line price after discount deducted
            vAmount = FixTwoPlaces(vAmount + vVatAmount - vDiscountAmount)
            mvClassFields.Item(TraderAnalysisLineFields.talfGrossAmount).DoubleValue = vGrossAmount
          Else
            'vGrossAmount = Line price including VAT
            'vAmount = Gross - Discount
            'vVAT = VAT element of Amount
            vVatAmount = CalculateVATAmount(vAmount, pVATPercentage)
          End If
        Else
          If mvClassFields.Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool = True Then
            'Price does not include VAT so need to deduct VAT first
            vAmount = FixTwoPlaces(vAmount - vVatAmount)
            vVatAmount = FixTwoPlaces(vAmount * (pVATPercentage / 100))
            vAmount = FixTwoPlaces(vAmount + vVatAmount)
          Else
            'Price includes VAT so just calculate the VAT element
            vVatAmount = CalculateVATAmount(vAmount, pVATPercentage)
          End If
        End If
        mvClassFields.Item(TraderAnalysisLineFields.talfVATPercentage).Value = CStr(pVATPercentage)
        mvClassFields.Item(TraderAnalysisLineFields.talfVATAmount).Value = CStr(vVatAmount)
        mvClassFields.Item(TraderAnalysisLineFields.talfAmount).DoubleValue = vAmount
      End If

    End Sub

    Friend Sub UpdateEventBookingDetails(ByVal pOldBookingNumber As Integer, ByVal pNewBookingNumber As Integer, ByVal pEventNumber As Integer)
      With mvClassFields
        If .Item(TraderAnalysisLineFields.talfEventBookingNumber).IntegerValue = pOldBookingNumber Then
          If .Item(TraderAnalysisLineFields.talfEventNumber).IntegerValue = pEventNumber Then .Item(TraderAnalysisLineFields.talfAmendedEventBookingNumber).Value = .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value
          .Item(TraderAnalysisLineFields.talfEventBookingNumber).Value = CStr(pNewBookingNumber)
        End If
      End With
    End Sub

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      LineDataType = mvClassFields.Item(pAttributeName).FieldType
    End Function

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        mvClassFields.Item(pAttributeName).Value = Value
      End Set
    End Property
    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(TraderAnalysisLineFields.talfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TraderTransactionTypeCode() As String
      Get
        TraderTransactionTypeCode = mvClassFields.Item(TraderAnalysisLineFields.talfTraderTransactionType).Value
      End Get
    End Property

    Public ReadOnly Property TraderLineTypeCode() As String
      Get
        TraderLineTypeCode = mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value
      End Get
    End Property

    Public ReadOnly Property TraderLineType() As TraderAnalysisLineTypes
      Get
        TraderLineType = GetAnalysisLineTypeFromCode(mvClassFields.Item(TraderAnalysisLineFields.talfTraderLineType).Value)
      End Get
    End Property

    Public ReadOnly Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(TraderAnalysisLineFields.talfPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(TraderAnalysisLineFields.talfProductCode).Value
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(TraderAnalysisLineFields.talfRate).Value
      End Get
    End Property

    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvClassFields.Item(TraderAnalysisLineFields.talfDistributionCode).Value
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(TraderAnalysisLineFields.talfQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(TraderAnalysisLineFields.talfSource).Value
      End Get
    End Property

    Public ReadOnly Property GrossAmount() As String
      Get
        'This will be null unless Discounts are being used
        GrossAmount = mvClassFields.Item(TraderAnalysisLineFields.talfGrossAmount).Value
      End Get
    End Property

    Public ReadOnly Property Discount() As String
      Get
        'This will be null unless Discounts are being used
        Discount = mvClassFields.Item(TraderAnalysisLineFields.talfDiscount).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(TraderAnalysisLineFields.talfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property AcceptAsFull() As Boolean
      Get
        AcceptAsFull = mvClassFields.Item(TraderAnalysisLineFields.talfAcceptAsFull).Bool
      End Get
    End Property

    Public ReadOnly Property LineDate() As String
      Get
        LineDate = mvClassFields.Item(TraderAnalysisLineFields.talfLineDate).Value
      End Get
    End Property

    Public ReadOnly Property DespatchMethod() As String
      Get
        DespatchMethod = mvClassFields.Item(TraderAnalysisLineFields.talfDespatchMethod).Value
      End Get
    End Property

    Public ReadOnly Property DeliveryContactNumber() As Integer
      Get
        DeliveryContactNumber = mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DeliveryAddressNumber() As Integer
      Get
        DeliveryAddressNumber = mvClassFields.Item(TraderAnalysisLineFields.talfDeliveryAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property VatRate() As String
      Get
        VatRate = mvClassFields.Item(TraderAnalysisLineFields.talfVatRate).Value
      End Get
    End Property

    Public ReadOnly Property VATPercentage() As String
      Get
        VATPercentage = mvClassFields.Item(TraderAnalysisLineFields.talfVATPercentage).Value
      End Get
    End Property

    Public ReadOnly Property SalesContactNumber() As Integer
      Get
        SalesContactNumber = mvClassFields.Item(TraderAnalysisLineFields.talfSalesContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SalesLedgerAccount() As String
      Get
        SalesLedgerAccount = mvClassFields.Item(TraderAnalysisLineFields.talfSalesLedgerAccount).Value
      End Get
    End Property

    Public ReadOnly Property Issued() As String
      Get
        'Issued may not have been set
        Issued = mvClassFields.Item(TraderAnalysisLineFields.talfIssued).Value
      End Get
    End Property

    Public ReadOnly Property StockSale() As Boolean
      Get
        StockSale = mvClassFields.Item(TraderAnalysisLineFields.talfStockSale).Bool
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(TraderAnalysisLineFields.talfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property FinancialAdjustment() As Boolean
      Get
        FinancialAdjustment = mvClassFields.Item(TraderAnalysisLineFields.talfFinancialAdjustment).Bool
      End Get
    End Property

    Public ReadOnly Property WarehouseCode() As String
      Get
        WarehouseCode = mvClassFields.Item(TraderAnalysisLineFields.talfWarehouseCode).Value
      End Get
    End Property

    Public ReadOnly Property ProductNumber() As Integer
      Get
        ProductNumber = mvClassFields.Item(TraderAnalysisLineFields.talfProductNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property GiverContactNumber() As Integer
      Get
        GiverContactNumber = mvClassFields.Item(TraderAnalysisLineFields.talfGiverContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ScheduledPaymentNumber() As Integer
      Get
        ScheduledPaymentNumber = mvClassFields.Item(TraderAnalysisLineFields.talfScheduledPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProvisionalBatchNumber() As Integer
      Get
        ProvisionalBatchNumber = mvClassFields.Item(TraderAnalysisLineFields.talfProvisionalBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProvisionalTransactionNumber() As Integer
      Get
        ProvisionalTransactionNumber = mvClassFields.Item(TraderAnalysisLineFields.talfProvisionalTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ProvisionalLineNumber() As Integer
      Get
        ProvisionalLineNumber = mvClassFields.Item(TraderAnalysisLineFields.talfProvisionalLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DeceasedContactNumber() As Integer
      Get
        DeceasedContactNumber = mvClassFields.Item(TraderAnalysisLineFields.talfDeceasedContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PaymentPlanTypeCode() As String
      Get
        PaymentPlanTypeCode = mvClassFields.Item(TraderAnalysisLineFields.talfPaymentPlanType).Value
      End Get
    End Property

    Public ReadOnly Property StockMovementNumbers() As String
      Get
        'This is a comma-separated list of numbers
        StockMovementNumbers = mvClassFields.Item(TraderAnalysisLineFields.talfStockMovementNumbers).Value
      End Get
    End Property

    Public ReadOnly Property MemberNumber() As String
      Get
        MemberNumber = mvClassFields.Item(TraderAnalysisLineFields.talfMemberNumber).Value
      End Get
    End Property

    Public ReadOnly Property InvoiceNumber() As Integer
      Get
        InvoiceNumber = mvClassFields.Item(TraderAnalysisLineFields.talfInvoiceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property InvoiceNumberUsed() As Integer
      Get
        InvoiceNumberUsed = mvClassFields.Item(TraderAnalysisLineFields.talfInvoiceNumberUsed).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(TraderAnalysisLineFields.talfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventBookingNumber() As Integer
      Get
        EventBookingNumber = mvClassFields.Item(TraderAnalysisLineFields.talfEventBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedEventBookingNumber() As Integer
      Get
        AmendedEventBookingNumber = mvClassFields.Item(TraderAnalysisLineFields.talfAmendedEventBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RoomBookingNumber() As Integer
      Get
        RoomBookingNumber = mvClassFields.Item(TraderAnalysisLineFields.talfRoomBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ServiceBookingNumber() As Integer
      Get
        ServiceBookingNumber = mvClassFields.Item(TraderAnalysisLineFields.talfServiceBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LegacyNumber() As Integer
      Get
        LegacyNumber = mvClassFields.Item(TraderAnalysisLineFields.talfLegacyNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BequestNumber() As Integer
      Get
        BequestNumber = mvClassFields.Item(TraderAnalysisLineFields.talfBequestNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LegacyReceiptNumber() As Integer
      Get
        LegacyReceiptNumber = mvClassFields.Item(TraderAnalysisLineFields.talfLegacyReceiptNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CovenantNumber() As Integer
      Get
        CovenantNumber = mvClassFields.Item(TraderAnalysisLineFields.talfCovenantNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(TraderAnalysisLineFields.talfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property DeclarationNumber() As Integer
      Get
        DeclarationNumber = mvClassFields.Item(TraderAnalysisLineFields.talfDeclarationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PayrollGivingDonorId() As String
      Get
        PayrollGivingDonorId = mvClassFields.Item(TraderAnalysisLineFields.talfPGDonorID).Value
      End Get
    End Property

    Public ReadOnly Property ContactStatus() As String
      Get
        ContactStatus = mvClassFields.Item(TraderAnalysisLineFields.talfContactStatus).Value
      End Get
    End Property

    Public ReadOnly Property PostagePacking() As Boolean
      Get
        PostagePacking = mvClassFields.Item(TraderAnalysisLineFields.talfPostagePacking).Bool
      End Get
    End Property

    Public ReadOnly Property ActivityGroup() As String
      Get
        ActivityGroup = mvClassFields.Item(TraderAnalysisLineFields.talfActivityGroup).Value
      End Get
    End Property

    Public ReadOnly Property Suppression() As String
      Get
        Suppression = mvClassFields.Item(TraderAnalysisLineFields.talfSuppression).Value
      End Get
    End Property

    Public ReadOnly Property AutoPaymentNumber() As Integer
      Get
        AutoPaymentNumber = mvClassFields.Item(TraderAnalysisLineFields.talfAutoPaymentNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property IncentiveLineNumber() As Integer
      Get
        IncentiveLineNumber = mvClassFields.Item(TraderAnalysisLineFields.talfIncentiveLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property VatAmount() As Double
      Get
        VatAmount = mvClassFields.Item(TraderAnalysisLineFields.talfVATAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property PayrollGivingPledgeNumber() As Integer
      Get
        PayrollGivingPledgeNumber = mvClassFields.Item(TraderAnalysisLineFields.talfPGPledgeNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(TraderAnalysisLineFields.talfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPisNumber() As Integer
      Get
        CollectionPisNumber = mvClassFields.Item(TraderAnalysisLineFields.talfCollectionPISNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionBankAccount() As String
      Get
        CollectionBankAccount = mvClassFields.Item(TraderAnalysisLineFields.talfCollectionBankAccount).Value
      End Get
    End Property

    Public ReadOnly Property CollectionBoxNumbers() As String
      Get
        'Comma separated list
        CollectionBoxNumbers = mvClassFields.Item(TraderAnalysisLineFields.talfCollectionBoxNumbers).Value
      End Get
    End Property

    Public ReadOnly Property CollectionBoxAmounts() As String
      Get
        'Comma separated list
        CollectionBoxAmounts = mvClassFields.Item(TraderAnalysisLineFields.talfCollectionBoxAmounts).Value
      End Get
    End Property

    Public ReadOnly Property StockTransactionID() As Integer
      Get
        'This is only used by Smart Client / Web Services
        StockTransactionID = mvClassFields.Item(TraderAnalysisLineFields.talfStockTransactionID).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PriceVATExclusive() As Boolean
      Get
        PriceVATExclusive = mvClassFields.Item(TraderAnalysisLineFields.talfPriceVATExclusive).Bool
      End Get
    End Property

    Public ReadOnly Property AdultQuantity() As Integer
      Get
        AdultQuantity = mvClassFields.Item(TraderAnalysisLineFields.talfAdultQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ChildQuantity() As Integer
      Get
        ChildQuantity = mvClassFields.Item(TraderAnalysisLineFields.talfChildQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StartTime() As String
      Get
        StartTime = mvClassFields.Item(TraderAnalysisLineFields.talfStartTime).Value
      End Get
    End Property

    Public ReadOnly Property EndTime() As String
      Get
        EndTime = mvClassFields.Item(TraderAnalysisLineFields.talfEndTime).Value
      End Get
    End Property

    Public ReadOnly Property CreditedContactNumber() As Integer
      Get
        Return mvClassFields.Item(TraderAnalysisLineFields.talfCreditedContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property InvoiceTypeUsed() As String
      Get
        Return mvClassFields.Item(TraderAnalysisLineFields.talfInvoiceTypeUsed).Value
      End Get
    End Property

    Public ReadOnly Property ExamBookingNumber() As Integer
      Get
        Return mvClassFields.Item(TraderAnalysisLineFields.talfExamBookingNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields.Item(TraderAnalysisLineFields.talfExamUnitId).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ExamUnitProductId() As Integer
      Get
        Return mvClassFields.Item(TraderAnalysisLineFields.talfExamUnitProductId).IntegerValue
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetTraderLineInfo(ByVal pTraderLineInfo As TraderAnalysisLineInfo) As Boolean

      Select Case pTraderLineInfo
        Case TraderAnalysisLineInfo.taliCreatesBTA 'Which Line Types will create a BTA record?
          Select Case TraderLineTypeCode
            Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "M", "N", "O", "P", "PG", "PP", "R", "S", "U", "V", "VC", "VE", "AP", "X", "Q", "K"
              GetTraderLineInfo = True
            Case "L"
              GetTraderLineInfo = InvoiceTypeUsed = "C"
          End Select
        Case TraderAnalysisLineInfo.taliHasStockMovement 'Which Line Types will create stock movements?
          Select Case TraderLineTypeCode
            Case "P"
              GetTraderLineInfo = True
          End Select
        Case TraderAnalysisLineInfo.taliIsMaintenanceType 'Which Line Types are Maintenance types?
          Select Case TraderLineTypeCode
            Case "AS", "GA", "ST", "CP", "AA", "GD", "GP", "CG", "ADDR"
              GetTraderLineInfo = True
          End Select
        Case TraderAnalysisLineInfo.taliIsInvoiceAllocation 'Which Line Types are Invoice Allocation types?
          Select Case TraderLineTypeCode
            Case "L", "K"
              GetTraderLineInfo = True
          End Select
      End Select

    End Function

    Public Sub SetLineNumber(ByVal pLineNumber As Integer, ByVal pAdjustmentType As Batch.AdjustmentTypes)
      'Only used in SCGetTransactionData for atAdjustment type to re-order the line numbers
      If pAdjustmentType = Batch.AdjustmentTypes.atAdjustment Then mvClassFields.Item(TraderAnalysisLineFields.talfLineNumber).IntegerValue = pLineNumber
    End Sub

    Friend Function CreditedContactDefaultAddressNumber(ByVal pEnv As CDBEnvironment) As Integer
      If mvCreditedContact Is Nothing Then
        mvCreditedContact = New Contact(pEnv)
        mvCreditedContact.Init(CreditedContactNumber)
      End If
      Return mvCreditedContact.Address.AddressNumber
    End Function

    Friend Sub SetDepositAllowed(ByVal pNewValue As Boolean)
      mvClassFields.Item(TraderAnalysisLineFields.talfDepositAllowed).Bool = pNewValue
    End Sub

  End Class
End Namespace
