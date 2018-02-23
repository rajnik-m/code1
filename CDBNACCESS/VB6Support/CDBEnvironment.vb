Namespace Access

  Partial Public Class CDBEnvironment

    Public Enum FunctionParameterTypes
      fptNone
      fptActionActivationDate
      fptUpdateStandingOrder
      fptPayingInSlipNumber
      fptCLRStatementDate
      fptInvoicePaymentDue
      fptReportNumber
      fptChangePayer
      fptEventProgrammeReport
      fptEventPersonnelReport
      fptEventAttendeeReport
      fptFinancialAdjustment
      fptCloseOrganisationSite
      fptMoveBranch
      fptMoveRegion
      fptUpdateContact
      fptAppealBudgetPeriod
      fptOwnershipGroup
      fptConfirmProvisionalTransaction
      fptCreateGiftAidDeclaration
      fptAuthorisePOPayment
      fptAdvanceRenewalDate
      fptFAPartRefund
      fptAttachmentList
      fptCancReason
      fptGetMailingCode
      fptPayPlanMissedPayments
      fptFAReverseRefundOptions
      fptStockSalesAnalysis
      fptSOReconciliationReport
      fptStockSalesAnalysisDetailed
      fptStockSalesAnalysisSummary
      fptReportDataSelection
      fptEventCancellationFAType
      fptChangeSubscriptionCommunication
      fptUpdatePaymentPlanDetailSource
      fptCMTPriceChange
      fptImportTraderApp
      fptScheduleTask
      fptLMAddressUsage
      fptAddCollectionBoxes
      fptCancellationReasonAndSource
      fptAllocatePISToEvent
      fptAllocatePISToDelegates
      fptDuplicateEvent
      fptMembershipReinstatement
      fptCancellationReasonSourceAndDate
      fptRemoveFutureMembershipType
      fptEditAppointment
      fptLeavePosition
      fptMovePosition
      fptCancelPaymentPlan
      fptPISPrinting
      fptScheduledJobDetails
      fptCMTEntitlementPriceChange
      fptNewMailingCode
      fptReAllocateProductNumber
      fptAddFastDataEntryPage
      fptCopyEventPricingMatrix
      fptEnterCancellationFee
      fptDuplicateSurvey
      fptPaymentPlanDocument
      'Add new values here
      fptMaxControls
      fptCopySegment
      fptSetChequeStatus
      fptLoadDataUpdates
      fptExamResultEntry
      fptListManagerRandomDataSample
      fptReCalculateLoanInterest
      fptExamChangeCentre
      fptCopyAppeal
      fptCLIBrowser
      fptDuplicateMeeting
      fptShareExamUnit
      fptPOPAnalysis
      fptExamCertificateReprint
      fptCopySegmentCriteria
      fptWorkstreamGroupActions
      fptActionChangeReasons
      fptExamScheduleWorkstreams
      fptSOCancellation
      fptExamCertificates
      fptStandardFields
      fptMiscFields
      fptSelectionTester
      fptFASLPartRefund
      fptGAReprintTaxClaim
    End Enum

    ''' <summary>
    ''' The enumerations for Payment Plan Types used by AddPaymentPlan
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ppType
      pptNull
      pptMember
      pptCovenant
      pptDD
      pptSO
      pptCCCA
      pptOther
      pptLoan
    End Enum

    Public Enum CreditCardValidationStatus
      ccvsValid
      ccvsInvalidNumber
      ccvsNotNumeric
    End Enum

    'Used as an optional parameter to CheckCCNumber
    Public Enum CreditCardValdationTypes
      ccvtStandard
      ccvtCAF
    End Enum

    Public Enum SelectionOptionSettings
      sosNone
      sosSome
      sosAll
    End Enum

    Public Enum MailOrderControlTypes
      moctCAF
      moctNonCAF
    End Enum

    Public Enum InvoicePayStatusTypes
      ipsFullyPaid
      ipsPartPaid
      ipsPaymentDue
      ipsDDCollectionPending
    End Enum

    Private mvInvoicePayStatus(4) As String
    Private mvMembershipTypes As CollectionList(Of MembershipType)
    Private mvBankAccounts As CollectionList(Of BankAccount)
    Private mvPaymentFrequencies As CollectionList(Of PaymentFrequency)
    Private mvCachedControlNumbers As CollectionList(Of CachedControlNumber)
    Private mvUniservInterface As UniservInterface
    Private mvPPOs As PostcodeProximityOrgs
    Private mvConnectionID As String

    Public Function GetInvoicePayStatusType(ByVal pStatus As String) As InvoicePayStatusTypes
      Dim vIndex As Integer
      BuildPayStatuses()
      For vIndex = 0 To mvInvoicePayStatus.Length - 1
        If mvInvoicePayStatus(vIndex) = pStatus Then Exit For
      Next
      Return CType(vIndex, InvoicePayStatusTypes)
    End Function

    Public Function GetInvoicePayStatus(ByVal pType As InvoicePayStatusTypes) As String
      BuildPayStatuses()
      GetInvoicePayStatus = mvInvoicePayStatus(pType)
    End Function

    Private Sub BuildPayStatuses()
      Dim vRecordSet As CDBRecordSet

      If Len(mvInvoicePayStatus(0)) = 0 Then
        vRecordSet = Connection.GetRecordSet("SELECT invoice_pay_status,not_paid,part_paid,fully_paid,pending_dd_payment FROM invoice_pay_statuses")
        While vRecordSet.Fetch()
          If vRecordSet.Fields(2).Bool Then mvInvoicePayStatus(InvoicePayStatusTypes.ipsPaymentDue) = vRecordSet.Fields(1).Value
          If vRecordSet.Fields(3).Bool Then mvInvoicePayStatus(InvoicePayStatusTypes.ipsPartPaid) = vRecordSet.Fields(1).Value
          If vRecordSet.Fields(4).Bool Then mvInvoicePayStatus(InvoicePayStatusTypes.ipsFullyPaid) = vRecordSet.Fields(1).Value
          If vRecordSet.Fields(5).Bool Then mvInvoicePayStatus(InvoicePayStatusTypes.ipsDDCollectionPending) = vRecordSet.Fields(1).Value
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    'Public Function GetBookingStatusCode(ByRef pNewValue As EventBooking.EventBookingStatuses) As String
    '  Dim vBookingStatus As String

    '  Select Case pNewValue
    '    Case EventBooking.EventBookingStatuses.ebsAmended
    '      vBookingStatus = "U"
    '    Case EventBooking.EventBookingStatuses.ebsBooked
    '      vBookingStatus = "F"
    '    Case EventBooking.EventBookingStatuses.ebsWaiting
    '      vBookingStatus = "W"
    '    Case EventBooking.EventBookingStatuses.ebsBookedTransfer
    '      vBookingStatus = "X"
    '    Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
    '      vBookingStatus = "B"
    '    Case EventBooking.EventBookingStatuses.ebsWaitingPaid
    '      vBookingStatus = "P"
    '    Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
    '      vBookingStatus = "Y"
    '    Case EventBooking.EventBookingStatuses.ebsBookedCreditSale
    '      vBookingStatus = "S"
    '    Case EventBooking.EventBookingStatuses.ebsWaitingCreditSale
    '      vBookingStatus = "A"
    '    Case EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer
    '      vBookingStatus = "R"
    '    Case EventBooking.EventBookingStatuses.ebsBookedInvoiced
    '      vBookingStatus = "V"
    '    Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced
    '      vBookingStatus = "O"
    '    Case EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer
    '      vBookingStatus = "D"
    '    Case EventBooking.EventBookingStatuses.ebsExternal
    '      vBookingStatus = "E"
    '    Case EventBooking.EventBookingStatuses.ebsCancelled
    '      vBookingStatus = "C"
    '    Case EventBooking.EventBookingStatuses.ebsInterested
    '      vBookingStatus = "I"
    '    Case EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
    '      vBookingStatus = "T"
    '  End Select
    '  GetBookingStatusCode = vBookingStatus
    'End Function


    'Public Function AddJournalRecord(ByVal pType As JournalTypes, ByVal pOperation As JournalOperations, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pSelect1 As Integer, optional ByVal pSelect2 As Integer = 0, optional ByVal pSelect3 As Integer = 0, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer) As Integer
    '  Return AddJournalRecord(pType, pOperation, pContactNumber, pAddressNumber, pSelect1, pSelect2, pSelect3, pBatchNumber, pTransactionNumber)
    'End Function

    Public Function EncryptCreditCardNumber(ByVal pCCNumber As String) As String
      Dim vEP As New EncryptionProvider
      Dim vCCNumber As String

      vCCNumber = pCCNumber
      If Len(pCCNumber) > 0 And GetDataStructureInfo(cdbDataStructureConstants.cdbDataCreditCardAVSCVV2) = True Then
        'Only Encrypt if the database has been upgraded
        vEP.Init(mvClientCode)
        If Left(pCCNumber, 1) <> "E" Then vCCNumber = "E" & vEP.Encrypt(pCCNumber) 'Encrypted number prefixed by 'E' so that we know it is encrypted
      End If
      Return vCCNumber
    End Function

    Public Function DecryptCreditCardNumber(ByVal pEncryptedCCNumber As String) As String
      Dim vEP As New EncryptionProvider
      Dim vCCNumber As String = ""

      If pEncryptedCCNumber.Length > 0 Then
        vEP.Init(mvClientCode)
        vCCNumber = pEncryptedCCNumber
        If Left(pEncryptedCCNumber, 1) = "E" Then vCCNumber = vEP.Decrypt(pEncryptedCCNumber.Substring(1)) 'Strip off the leading 'E' before decrypting
      End If
      Return vCCNumber
    End Function

    Public Function MembershipType(ByRef pCode As String) As MembershipType

      If mvMembershipTypes Is Nothing Then mvMembershipTypes = New CollectionList(Of MembershipType)
      If mvMembershipTypes.ContainsKey(pCode) Then
        Return mvMembershipTypes(pCode)
      Else
        Dim vMembershipType As MembershipType = New MembershipType(Me)
        vMembershipType.Init(pCode)
        mvMembershipTypes.Add(pCode, vMembershipType)
        Return vMembershipType
      End If
    End Function

    Public Function BankAccount(ByRef pCode As String) As BankAccount
      If mvBankAccounts Is Nothing Then mvBankAccounts = New CollectionList(Of BankAccount)
      If mvBankAccounts.ContainsKey(pCode) Then
        Return mvBankAccounts.Item(pCode)
      Else
        Dim vBankAccount As BankAccount = New BankAccount(Me)
        vBankAccount.Init(pCode)
        mvBankAccounts.Add(pCode, vBankAccount)
        Return vBankAccount
      End If
    End Function

    Friend Function GetPaymentFrequency(ByRef pCode As String) As PaymentFrequency
      If mvPaymentFrequencies Is Nothing Then
        mvPaymentFrequencies = New CollectionList(Of PaymentFrequency)
        Dim vPF As PaymentFrequency = New PaymentFrequency
        vPF.Init(Me)
        Dim vRecordSet As CDBRecordSet = Connection.GetRecordSet("SELECT " & vPF.GetRecordSetFields(PaymentFrequency.PaymentFrequencyRecordSetTypes.pfrtAll) & " FROM payment_frequencies ORDER BY payment_frequency")
        While vRecordSet.Fetch()
          vPF = New PaymentFrequency
          vPF.InitFromRecordSet(Me, vRecordSet, PaymentFrequency.PaymentFrequencyRecordSetTypes.pfrtAll)
          mvPaymentFrequencies.Add(vPF.PaymentFrequencyCode, vPF)
        End While
        vRecordSet.CloseRecordSet()
      End If
      If mvPaymentFrequencies.ContainsKey(pCode) Then
        Return mvPaymentFrequencies(pCode)
      Else
        RaiseError(DataAccessErrors.daePaymentFrequencyInvalid, pCode)
        Return Nothing       'Just to fix warning
      End If
    End Function

    Public Function GetNextItem(ByVal pCollection As Collection, ByVal pCurrentItem As Object) As Object
      'Return next item in the collection after the current item
      Dim vObj As Object = Nothing
      Dim vCurrFound As Boolean = False 'Current item found
      Dim vNextFound As Boolean = False 'Next item found

      If pCurrentItem Is Nothing Then vCurrFound = True 'Get first item in collection
      For Each vObj In pCollection
        If vCurrFound = True Then vNextFound = True
        If pCurrentItem IsNot Nothing Then
          If vObj Is pCurrentItem Then
            vCurrFound = True
          End If
        End If
        If vNextFound Then Exit For
      Next vObj
      If vNextFound = False Then vObj = Nothing 'Need to return Nothing if there were no further items in the collection
      Return vObj
    End Function

    Public Function GetPreviousItem(ByVal pCollection As Collection, ByVal pCurrentItem As Object) As Object
      'Return previous item in the collection before the current item
      Dim vPrevObj As Object = Nothing
      Dim vFound As Boolean
      'If pCurrentItem Is Nothing then the last item is required
      For Each vObj As Object In pCollection
        If Not (pCurrentItem Is Nothing) Then
          If vObj Is pCurrentItem Then vFound = True
        End If
        If vFound Then Exit For
        vPrevObj = vObj
      Next vObj
      Return vPrevObj
    End Function

    Public Function GetRegionNumber(ByVal pBranch As String) As Integer
      Dim vRegionNumber As Integer
      If pBranch.Length > 0 Then
        Dim vParentLink As String = GetControlValue(cdbControlConstants.cdbControlBranchParent)
        If vParentLink.Length > 0 Then
          Dim vRecordSet As CDBRecordSet = Connection.GetRecordSet("SELECT organisation_number_2 FROM branches b,organisation_links ol WHERE b.branch = '" & pBranch & "' AND ol.organisation_number_1 = b.organisation_number AND relationship = '" & vParentLink & "' AND (ol.historical IS NULL OR ol.historical = 'N')")
          If vRecordSet.Fetch() Then
            vRegionNumber = vRecordSet.Fields(1).LongValue
          End If
          vRecordSet.CloseRecordSet()
        End If
      End If
      Return vRegionNumber
    End Function

    Public Function GetBranchFromAddress(ByVal pAddressNumber As Integer) As String
      Return Connection.GetValue("SELECT branch FROM addresses WHERE address_number = " & pAddressNumber)
    End Function

    Public Sub GetCancellationInfo(ByRef pCancellationReason As String, ByRef pStatus As String, ByRef pDescription As String)
      pStatus = ""
      pDescription = ""
      If pCancellationReason.Length > 0 Then
        Dim vRecordSet As CDBRecordSet = Connection.GetRecordSet("SELECT cancellation_reason_desc, cr.status, reason_required FROM cancellation_reasons cr, statuses s WHERE cr.cancellation_reason = '" & pCancellationReason & "' AND cr.status = s.status")
        If vRecordSet.Fetch() Then
          If vRecordSet.Fields(3).Bool Then pDescription = vRecordSet.Fields(1).Value
          pStatus = vRecordSet.Fields(2).Value
        End If
        vRecordSet.CloseRecordSet()
      Else
        RaiseError(DataAccessErrors.daeInvalidCancellationReason)
      End If
    End Sub

    Public Sub AddAuditRecord(ByVal pType As AuditTypes, ByVal pTable As String, ByVal pSelect1 As Integer, ByVal pSelect2 As Integer, ByVal pFieldName As String, ByVal pOldValue As String, Optional ByVal pNewValue As String = "", Optional ByVal pOldFields As CDBFields = Nothing, Optional ByVal pNewFields As CDBFields = Nothing)
      Dim vOperation As String = ""
      Dim vFields As New CDBFields
      Dim vSelect1 As String = ""
      Dim vSelect2 As String = ""
      Dim vValues As String
      Dim vField As CDBField

      Select Case pType
        Case AuditTypes.audInsert
          vOperation = "insert"
        Case AuditTypes.audUpdate
          vOperation = "update"
        Case AuditTypes.audDelete
          vOperation = "delete"
      End Select
      If pSelect1 > 0 Then vSelect1 = pSelect1.ToString
      If pSelect2 > 0 Then vSelect2 = pSelect2.ToString

      Select Case AuditStyle
        Case AuditStyleTypes.ausAmendmentHistory, AuditStyleTypes.ausExtended
          With vFields
            .Add("operation", CDBField.FieldTypes.cftCharacter, vOperation)
            .Add("operation_date", CDBField.FieldTypes.cftTime, TodaysDateAndTime)
            .Add("table_name", CDBField.FieldTypes.cftCharacter, pTable)
            .Add("logname", CDBField.FieldTypes.cftCharacter, mvUser.Logname)
            .Add("select_1", CDBField.FieldTypes.cftLong, vSelect1)
            .Add("select_2", CDBField.FieldTypes.cftLong, vSelect2)
            If pFieldName.Length > 0 Then
              vValues = pFieldName & ":" & pOldValue & Chr(22) & vbCrLf & Chr(22) & pFieldName & ":" & pNewValue & Chr(22)
            Else
              vValues = ""
              If pType <> AuditTypes.audInsert Then
                vValues = "OLD" & Chr(22)
                For Each vField In pOldFields
                  vValues = vValues & vField.Name & ":" & vField.Value & Chr(22)
                Next vField
                vValues = vValues & vbCrLf
              End If
              If pType <> AuditTypes.audDelete Then
                vValues = vValues & "NEW" & Chr(22)
                For Each vField In pNewFields
                  vValues = vValues & vField.Name & ":" & vField.Value & Chr(22)
                Next vField
                vValues = vValues & vbCrLf
              End If
            End If
            .Add("data_values", CDBField.FieldTypes.cftMemo, vValues)
            Connection.InsertRecord("amendment_history", vFields)
          End With
      End Select
    End Sub

    Public Function GetFloorLimit(ByVal pType As MailOrderControlTypes) As Double
      Dim vSQL As String
      Dim vRecordSet As CDBRecordSet
      Dim vRecordMissing As Boolean

      'Get the floor limit amount of the CAF & NONCAF records in the mail_order_controls table
      vSQL = "SELECT floor_limit FROM mail_order_controls WHERE report_type = '"
      If pType = MailOrderControlTypes.moctNonCAF Then vSQL = vSQL & "NON"
      vSQL = vSQL & "CAF'"
      vRecordSet = Connection.GetRecordSet(vSQL)
      With vRecordSet
        If .Fetch() Then
          GetFloorLimit = .Fields(1).DoubleValue
        Else
          vRecordMissing = True
        End If
        .CloseRecordSet()
      End With
      If vRecordMissing Then RaiseError(DataAccessErrors.daeMissingMailOrderControlRecord, If(pType = MailOrderControlTypes.moctCAF, "CAF", "NONCAF"))
    End Function

    Public Sub CacheControlNumbers(ByVal pNumberType As CachedControlNumberTypes, ByVal pCount As Integer)
      Dim vType As String = ""
      Dim vIgnore As Boolean
      Dim vCCN As CachedControlNumber

      Select Case pNumberType
        Case CachedControlNumberTypes.ccnJournal
          vType = "CJ"
          If Not mvJournalInitialised Then InitJournal()
          If Not mvOptionJournal Then vIgnore = True
        Case CachedControlNumberTypes.ccnPaymentSchedule
          vType = "SP"
        Case CachedControlNumberTypes.ccnTimesheet
          vType = "TK"
        Case CachedControlNumberTypes.ccnAddress
          vType = "A"
        Case CachedControlNumberTypes.ccnContact
          vType = "C"
        Case CachedControlNumberTypes.ccnAddressLink
          vType = "AL"
        Case CachedControlNumberTypes.ccnPosition
          vType = "PN"
        Case CachedControlNumberTypes.ccnExamMarkingBatchDetail
          vType = "XMD"
        Case CachedControlNumberTypes.ccnExamCentreUnit
          vType = "XCU"
      End Select

      If Not vIgnore Then
        If mvCachedControlNumbers Is Nothing Then
          mvCachedControlNumbers = New CollectionList(Of CachedControlNumber)
        End If
        If mvCachedControlNumbers.ContainsKey(vType) Then
          vCCN = mvCachedControlNumbers(vType)
          vCCN.CheckAvailable(pCount)
        Else
          vCCN = New CachedControlNumber
          vCCN.Init(Me, vType, pCount)
          mvCachedControlNumbers.Add(vType, vCCN)
        End If
      End If
    End Sub

    Public Function GetExchangeRate(ByVal pCurrencyCode As String) As Double
      Dim vCurrencyCode As New CurrencyCode()
      vCurrencyCode.Init(Me, CurrencyCode.CurrencyCodeRecordSetTypes.ccrtExchangeRate, pCurrencyCode)
      If vCurrencyCode.ExchangeRate = 0 Then
        RaiseError(DataAccessErrors.daeNoExchangeRate, pCurrencyCode)
      Else
        Return vCurrencyCode.ExchangeRate
      End If
    End Function

    Public Function GetServiceBookingStatusCode(ByRef pStatus As ServiceBooking.ServiceBookingStatuses) As String
      Select Case pStatus
        Case ServiceBooking.ServiceBookingStatuses.sbsBooked
          Return "B"
        Case ServiceBooking.ServiceBookingStatuses.sbsCancelled
          Return "C"
        Case Else
          Return ""       'To remove warning
      End Select
    End Function

    Public Function GetAttributeLength(ByVal pTable As String, ByRef pAttr As String) As Integer
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("table_name", pTable)
      vWhereFields.Add("attribute_name", pAttr)
      Dim vSQL As New SQLStatement(Connection, "entry_length", "maintenance_attributes", vWhereFields)
      Return vSQL.GetIntegerValue
    End Function

    Public Function UniservInterface() As UniservInterface
      If mvUniservInterface Is Nothing Then
        mvUniservInterface = New UniservInterface
        mvUniservInterface.Init(Me)
      End If
      Return mvUniservInterface
    End Function

    Public Function LastUniservMessage() As String
      Dim vReturnValue As String = String.Empty 'BR21061. If Uniserv is not present, always return an empty string.
      If mvUniservInterface IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(mvUniservInterface.LastErrorMessage) Then
        vReturnValue = mvUniservInterface.LastErrorMessage
      End If
      Return vReturnValue
    End Function

    Public Function CheckCCNumber(ByVal pString As String, Optional ByVal pCardType As CreditCardValdationTypes = CreditCardValdationTypes.ccvtStandard) As CreditCardValidationStatus
      Dim vString As String
      Dim vCheckDigit As Integer
      Dim vTotal As Integer
      Dim vPos As Integer
      Dim vValue As Integer
      Dim vTopValue As Integer
      Dim vCalcDigit As Integer
      Dim vDouble As Boolean
      Dim vStatus As CreditCardValidationStatus

      'Remove spaces - there should not be any
      vString = pString.Replace(" ", "")
      vStatus = CreditCardValidationStatus.ccvsValid
      If vString.Length > 0 Then
        If Not IsNumeric(vString) Then
          vStatus = CreditCardValidationStatus.ccvsNotNumeric
        Else
          Select Case pCardType
            Case CreditCardValdationTypes.ccvtStandard
              vCheckDigit = IntegerValue(Right(vString, 1))
              vPos = vString.Length - 1
              vDouble = True
              While vPos > 0
                vValue = IntegerValue(Mid(vString, vPos, 1))
                If vDouble Then
                  vValue = vValue * 2
                  If vValue > 9 Then vValue = (vValue - 10) + 1
                End If
                vTotal = vTotal + vValue
                vDouble = Not vDouble
                vPos = vPos - 1
              End While
              vTopValue = ((CInt(Int(vTotal / 10))) + 1) * 10
              vCalcDigit = vTopValue - vTotal
              If vCalcDigit = 10 Then vCalcDigit = 0
            Case CreditCardValdationTypes.ccvtCAF
              If Left(vString, 6) <> "564193" Then
                vStatus = CreditCardValidationStatus.ccvsInvalidNumber
              ElseIf Len(vString) < 16 Then
                vStatus = CreditCardValidationStatus.ccvsInvalidNumber
              Else
                vString = Left(vString, Len(vString) - 2) 'remove last 2 digits = sequence number
                vString = Right(vString, Len(vString) - 6) 'remove first 6 digits = standard CAF card indicator
                vCheckDigit = IntegerValue(Left(vString, 1)) 'check digit is first digit
                vPos = 2
                vDouble = False
                While vPos < 9
                  vValue = IntegerValue(Mid(vString, vPos, 1))
                  If vDouble Then vValue = vValue * 2 'multiply the digits that appear in positions 2, 4 & 6 by 2
                  vTotal = vTotal + vValue 'sum the value of the digits
                  vDouble = Not vDouble
                  vPos = vPos + 1
                End While
                vCalcDigit = (vTotal Mod 9) + 1 'check digit = (remainder of vTotal/9) + 1
              End If
          End Select
          'if check digit != calculated check digit then invalid number
          If vStatus = CreditCardValidationStatus.ccvsValid And vCalcDigit <> vCheckDigit Then vStatus = CreditCardValidationStatus.ccvsInvalidNumber
        End If
      End If
      Return vStatus
    End Function

    ''' <summary>Checks the IBAN number for validitity.</summary>
    ''' <param name="pIbanNumber">Iban number to validate</param>
    ''' <returns>True if IBAN number is valid else False</returns>
    Public Function CheckIbanNumber(ByVal pIbanNumber As String) As Boolean
      Dim vValid As Boolean = False
      If pIbanNumber.Length > 0 Then
        Dim vCountryIbanNumber As CountryIbanNumber = Me.GetCountryIban(pIbanNumber)
        vValid = vCountryIbanNumber.ValidateIbanNumber(pIbanNumber)
      End If

      Return vValid
    End Function

    Public Function GetRecordSetParseSQL(ByRef pSQL As String) As CDBRecordSet
      Return GetRecordSetParseSQL(pSQL, CDBConnection.RecordSetOptions.None)
    End Function

    Public Function GetRecordSetParseSQL(ByRef pSQL As String, ByVal pOptions As CDBConnection.RecordSetOptions) As CDBRecordSet
      Dim vSQL As String = pSQL
      If InStr(1, pSQL, " print ", CompareMethod.Text) > 0 Then RaiseError(DataAccessErrors.daeDBSpecificSQL, "'PRINT'")
      If InStr(1, pSQL, "= null", CompareMethod.Text) > 0 Then RaiseError(DataAccessErrors.daeDBSpecificSQL, "'= NULL'")
      If InStr(pSQL, """") > 0 Then RaiseError(DataAccessErrors.daeDBSpecificSQL, "'Double Quotes'")
      If InStr(1, pSQL, " AS '", CompareMethod.Text) > 0 Then RaiseError(DataAccessErrors.daeDBSpecificSQL, "'AS with quoted alias'")

      If Connection.DBForceOrder.Length = 0 Then vSQL = vSQL.Replace("OPTION (FORCE ORDER)", " ")

      ReplaceAllODBCFunctions(vSQL)
      ReplaceAllSpecialColumns(vSQL)

      'TODO VB6 Conversion 
      Dim vNullsSortAtEnd As Boolean = Connection.NullsSortAtEnd
      Dim vCollateString As String = Connection.DBCollateString
      If vNullsSortAtEnd OrElse vCollateString.Length > 0 Then
        Dim vPos As Integer = InStr(1, vSQL, "ORDER BY", CompareMethod.Text)
        If vPos > 0 Then
          Dim vOrderClauses() As String = Split(Mid(vSQL, vPos + 9), ",")
          For vIndex As Integer = 0 To UBound(vOrderClauses)
            If vCollateString.Length > 0 Then
              'remove any leading or trailing spaces & table aliases
              Dim vOrderByAttr As String = vOrderClauses(vIndex).Trim
              Dim vPos2 As Integer = InStr(1, vOrderByAttr, ".")
              If vPos2 > 0 Then vOrderByAttr = Mid(vOrderByAttr, vPos2 + 1)
              If CanUseCollation(vOrderByAttr) Then
                vOrderClauses(vIndex) = vOrderClauses(vIndex) & vCollateString
              End If
            End If
            If vNullsSortAtEnd Then
              If InStr(1, vOrderClauses(vIndex), " DESC", CompareMethod.Text) <= 0 Then vOrderClauses(vIndex) = vOrderClauses(vIndex) & " Nulls First"
            End If
          Next
          vSQL = Left(vSQL, vPos - 1) & " ORDER BY " & Join(vOrderClauses, ",")
        End If
      End If

      If InStr(1, vSQL, "INNER JOIN", CompareMethod.Text) > 0 Or InStr(1, vSQL, "OUTER JOIN", CompareMethod.Text) > 0 Then
        Return Connection.GetRecordSetAnsiJoins(vSQL, 0, pOptions)
      Else
        Return Connection.GetRecordSet(vSQL, pOptions)
      End If
    End Function

    Public Sub ReplaceAllSpecialColumns(ByRef pSQL As String)
      ReplaceSpecialCol(pSQL, "current")
      ReplaceSpecialCol(pSQL, "number")
      ReplaceSpecialCol(pSQL, "reference")
      ReplaceSpecialCol(pSQL, "expression")
      ReplaceSpecialCol(pSQL, "function")
      ReplaceSpecialCol(pSQL, "when")
      ReplaceSpecialCol(pSQL, "primary")
      ReplaceSpecialCol(pSQL, "distributed")
    End Sub

    Private Function CanUseCollation(ByVal pFieldName As String) As Boolean
      Select Case pFieldName
        Case "surname", "forenames", "name"
          Return True
        Case "label_name", "initials", "sort_name", "search_name"
          Return False
          'do nothing...yet
          'these attributes may contain accented characters but are not currently used in any report SQL as ORDER BY attributes
        Case "activity_value_desc", "room_type_desc", "comment_priority_desc"
          Return False
          'noticed that these description attributes are used as ORDER BY attributes in report SQL
          'it is possible that some clients will use accented characters in these so need to prepared in case a problem is encountered
        Case Else
          Return False
      End Select
    End Function

    Private Sub ReplaceFNIfNull(ByRef pSQL As String)
      While pSQL.Contains("{fn ifnull")
        Dim vIsNull As String = Connection.DBIsNull("", "")
        Dim vPos As Integer = vIsNull.IndexOf("(")
        vIsNull = vIsNull.Substring(0, vPos)
        pSQL = pSQL.Replace("{fn ifnull", vIsNull)
        pSQL = pSQL.Replace("}", " ")
      End While
    End Sub

    Private Sub ReplaceFNYear(ByRef pSQL As String)
      While pSQL.Contains("{fn year")
        Dim vPos As Integer = pSQL.IndexOf("{fn year")
        If vPos >= 0 Then
          Dim vEndPos As Integer = pSQL.IndexOf("}", vPos) + 1
          Dim vText As String = pSQL.Substring(vPos, vEndPos - vPos)
          If vText.Length > 0 Then
            Dim vDateStartPos As Integer = vText.IndexOf("(") + 1
            Dim vDateEndPos As Integer = vText.IndexOf(")", vDateStartPos)
            Dim vDateString As String = vText.Substring(vDateStartPos, vDateEndPos - vDateStartPos)   'Gives us just the date portion of {fn year (...)}
            Dim vYearText As String = Connection.DBYear(vDateString)      'Returns the text to substitute
            pSQL = pSQL.Substring(0, vPos) & vYearText & pSQL.Substring(vEndPos)
          End If
        End If
      End While
    End Sub

    Private Sub ReplaceFunction(ByRef pSQL As String, ByVal pFunction As String)
      Dim vPos As Integer
      Dim vLen As Integer
      Dim vStart As Integer
      Dim vPos2 As Integer
      Dim vPos3 As Integer
      Dim vText As String

      vStart = 1
      vLen = pFunction.Length
      Do
        vPos = InStr(vStart, pSQL, pFunction & "(")
        If vPos > 0 Then
          vPos2 = InStr(vPos, pSQL, ")")
          If vPos2 > 0 Then
            If pFunction = "length" Then
              pSQL = Left(pSQL, vPos - 1) & Connection.DBLength(Mid(pSQL, vPos + vLen + 1, (vPos2 - vPos) - vLen - 1)) & Mid(pSQL, vPos2 + 1)
            ElseIf pFunction = "ADDMONTHS" Then
              vPos3 = InStr(vPos, pSQL, ",")
              vText = Connection.DBAddMonths(Mid(pSQL, vPos + vLen + 1, (vPos3 - vPos) - vLen - 1), Mid(pSQL, vPos3 + 1, (vPos2 - vPos3) - 1))
              pSQL = Left(pSQL, vPos - 1) & vText & Right(pSQL, Len(pSQL) - vPos2)
            ElseIf pFunction = "TODAY" Then
              vText = Connection.SQLLiteral("", CARE.Data.CDBField.FieldTypes.cftDate, "today")
              pSQL = Left(pSQL, vPos - 1) & vText & Right(pSQL, Len(pSQL) - vPos2)
            End If
          End If
        End If
        vStart = vPos2
      Loop While vPos > 0
    End Sub

    Private Sub ReplaceSpecialCol(ByRef pSQL As String, ByVal pName As String)
      Dim vPos As Integer
      Dim vDoReplace As Boolean
      Dim vChar As String

      Dim vStart As Integer = 1
      Dim vLen As Integer = pName.Length
      Do
        vPos = InStr(vStart, pSQL, pName)
        If vPos > 0 Then
          vDoReplace = True
          If vPos > 1 Then
            vChar = Mid(pSQL, vPos - 1, 1)
            If vChar = "_" Or (vChar >= "a" And vChar <= "z") Then vDoReplace = False
          End If
          vChar = Mid(pSQL, vPos + vLen, 1)
          If vChar = "_" Or (vChar >= "a" And vChar <= "z") Then vDoReplace = False
          If vDoReplace Then
            pSQL = Left(pSQL, vPos - 1) & Connection.DBSpecialCol("", pName) & Mid(pSQL, vPos + vLen)
          End If
          vStart = vPos + vLen
        End If
      Loop While vPos > 0
    End Sub

    Public ReadOnly Property PostcodeProximityOrganisations() As PostcodeProximityOrgs
      Get
        If mvPPOs Is Nothing Then
          mvPPOs = New PostcodeProximityOrgs
          mvPPOs.Init(Me)
        End If
        Return mvPPOs
      End Get
    End Property

    Public Function GenerateContactNumberCheckDigit(ByVal pContactNumber As Integer) As Integer
      'This will take a contact number and output the check-digit for it
      Dim vCount As Integer
      Dim vSum As Integer

      Dim vContactNumber As String = pContactNumber.ToString
      'Contact Number must be 8 digits.  If the number is less than eight digits, zeroes are concatenated to the left of the number until the number consists of eight digits.
      While vContactNumber.Length < 8
        vContactNumber = "0" & vContactNumber
      End While
      'Each digit of the number is weighted, where the leftmost number is given a weighting of 1, the next 2, and so on.
      While vContactNumber.Length > 0
        vCount = vCount + 1
        vSum = vSum + CInt(vContactNumber.Substring(0, 1)) * vCount
        If vContactNumber.Length = 1 Then
          vContactNumber = ""
        Else
          vContactNumber = vContactNumber.Substring(1)
        End If
      End While
      vCount = vSum Mod 10    'vCount / 10 only returning the remainder
      Return vCount
    End Function

    Public Function UpdateMailingHistory(ByVal pMailingNumber As Integer, ByVal pMailing As String, ByVal pSelectionSet As String, ByVal pTempTable As String, ByVal pMailingDesc As String, Optional ByVal pMailingFileName As String = "", Optional ByVal pUseMailingCode As Boolean = False, Optional ByVal pMailingDate As String = "", Optional ByVal pNumberInMailing As Integer = 0, Optional ByVal pNumberOfMailingCodes As Integer = 1, Optional ByVal pRevision As Integer = 0, Optional ByVal pMailingNotes As String = "", Optional ByVal pMailingHistoryNotes As String = "", Optional ByVal pCheckTempTable As Boolean = True, Optional pTopic As String = "", Optional pSubTopic As String = "", Optional pSubject As String = "") As Integer
      'pCheckTempTable is used to decide whether to check the temp table for the existance of the mailing attribute as it will fail if we are in a transaction on an Oracle database
      ' - Passing False to pCheckTempTable will assume that the mailing attribute does not exist (default is True so temp table will be checked)
      Dim vSQL As String
      Dim vInsertFields As New CDBFields
      Dim vDT As CDBDataTable = Nothing
      Dim vDR As CDBDataRow
      Dim vRowsInserted As Integer
      Dim vMailingOnTempTable As Boolean
      Dim vAllowDuplicates As Boolean = GetConfigOption("ml_allow_duplicates")
      If pTempTable.Length > 0 Then
        'create contact_mailings records from temp table
        vSQL = "INSERT INTO contact_mailings (mailing_number,contact_number,address_number)"
        vSQL = vSQL & " SELECT %1,contact_number,address_number FROM " & pTempTable
        If pCheckTempTable Then vMailingOnTempTable = Connection.AttributeExists(pTempTable, "mailing") And pNumberOfMailingCodes > 1
        If vMailingOnTempTable Then
          vSQL = vSQL & " WHERE mailing = '%2'"
          If pRevision > 0 Then vSQL = vSQL & " AND revision = " & pRevision
          vDT = New CDBDataTable
          If vAllowDuplicates Then
            vDT.FillFromSQLDONOTUSE(Me, "SELECT mailing FROM " & pTempTable)
          Else
            vDT.FillFromSQLDONOTUSE(Me, "SELECT DISTINCT mailing FROM " & pTempTable)
          End If
          vDT.AddColumn("number_in_mailing", CDBField.FieldTypes.cftLong)
          vDT.AddColumn("mailing_number", CDBField.FieldTypes.cftLong)
          Dim vFirstRow As Boolean = True
          For Each vDR In vDT.Rows
            If vFirstRow Then
              vDR.Item(3) = CStr(pMailingNumber)
              vFirstRow = False
            Else
              vDR.Item(3) = CStr(GetControlNumber("MA"))
            End If
            vDR.Item(2) = CStr(Connection.ExecuteSQL(Replace(Replace(vSQL, "%1", vDR.Item(3)), "%2", vDR.Item(1))))
            vRowsInserted = vRowsInserted + IntegerValue(vDR.Item(2))
          Next vDR
        Else
          vSQL = Replace(vSQL, "%1", CStr(pMailingNumber))
          If pUseMailingCode Then
            vSQL = vSQL & " WHERE mailing = '" & pMailing & "'"
          ElseIf pSelectionSet.Length > 0 Then
            If InStr(pSelectionSet, ",") > 0 Then
              vSQL = vSQL & " WHERE selection_set IN (" & pSelectionSet & ")"
            Else
              vSQL = vSQL & " WHERE selection_set = " & pSelectionSet
            End If
          End If
          If pRevision > 0 Then
            vSQL = vSQL & " AND revision = " & pRevision
          End If
          vRowsInserted = Connection.ExecuteSQL(vSQL)
        End If
      Else
        vRowsInserted = pNumberInMailing
      End If

      If pMailingDate.Length = 0 Then pMailingDate = TodaysDate()
      'create mailing_history record
      vInsertFields = New CDBFields
      With vInsertFields
        .Add("mailing")
        .Add("mailing_date", CDBField.FieldTypes.cftDate, pMailingDate)
        .Add("mailing_by", CDBField.FieldTypes.cftCharacter, User.Logname)
        .Add("number_in_mailing", CDBField.FieldTypes.cftLong)
        .Add("mailing_number", CDBField.FieldTypes.cftLong)
        .Add("mailing_filename", CDBField.FieldTypes.cftCharacter, pMailingFileName)
        If pMailingHistoryNotes.Length > 0 Then .Add("notes", CDBField.FieldTypes.cftMemo, pMailingHistoryNotes)
        If pTopic.Length > 0 Then .Add("topic", pTopic)
        If pSubTopic.Length > 0 Then .Add("sub_topic", pSubTopic)
        If pSubject.Length > 0 Then .Add("subject", pSubject)
      End With
      If vMailingOnTempTable Then
        For Each vDR In vDT.Rows
          With vInsertFields
            .Item(1).Value = vDR.Item(1)
            .Item(4).Value = vDR.Item(2)
            .Item(5).Value = vDR.Item(3)
          End With
          Connection.InsertRecord("mailing_history", vInsertFields)
        Next vDR
      Else
        With vInsertFields
          .Item(1).Value = pMailing
          .Item(4).Value = CStr(vRowsInserted)
          .Item(5).Value = CStr(pMailingNumber)
        End With
        Connection.InsertRecord("mailing_history", vInsertFields)
      End If

      'create mailings record, if req'd
      If pMailingDesc.Length > 0 Then
        If Connection.GetCount("mailings", Nothing, "mailing = '" & pMailing & "'") = 0 Then
          vInsertFields = New CDBFields
          With vInsertFields
            .Add("mailing", CDBField.FieldTypes.cftCharacter, pMailing)
            .Add("mailing_desc", CDBField.FieldTypes.cftCharacter, pMailingDesc)
            .Add("department", CDBField.FieldTypes.cftCharacter, User.Department)
            .Add("direction", CDBField.FieldTypes.cftCharacter, "O")
            .Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
            .Add("marketing", CDBField.FieldTypes.cftCharacter, "N")
            If Len(pMailingNotes) > 0 Then .Add("notes", CDBField.FieldTypes.cftMemo, pMailingNotes)
            .AddAmendedOnBy(User.Logname)
          End With
          Connection.InsertRecord("mailings", vInsertFields)
        End If
      End If
      Return vRowsInserted
    End Function

    Public Function UpdateEMailingHistory(ByVal pMailingNumber As Integer, ByVal pSelectionSet As String, ByVal pTempTable As String, ByVal pRevision As Long) As Integer
      Dim vSQL As String
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vRowsInserted As Integer

      'create contact_emailings records from temp table
      vSQL = "INSERT INTO contact_emailings (mailing_number,contact_number,address_number,communication_number)"
      vSQL = vSQL & " SELECT " & pMailingNumber & ",contact_number,address_number,communication_number FROM " & pTempTable
      If InStr(pSelectionSet, ",") > 0 Then
        vSQL = vSQL & " WHERE selection_set IN (" & pSelectionSet & ")"
      Else
        vSQL = vSQL & " WHERE selection_set = " & pSelectionSet
      End If
      vSQL = vSQL & " AND revision = " & pRevision
      vRowsInserted = Connection.ExecuteSQL(vSQL)

      vUpdateFields.Add("number_in_mailing", CDBField.FieldTypes.cftInteger, "number_in_mailing + " & vRowsInserted)
      vWhereFields.Add("mailing_number", pMailingNumber)
      Connection.UpdateRecords("mailing_history", vUpdateFields, vWhereFields)
      Return vRowsInserted
    End Function

    Private Class LookupData
      Public TableName As String
      Public AttributeName As String
      Public Items As SortedList(Of String, String)

      Public Sub New(ByVal pTableName As String, ByVal pAttributeName As String)
        TableName = pTableName
        AttributeName = pAttributeName
        Items = New SortedList(Of String, String)
      End Sub
    End Class

    Private mvLookupData As New List(Of LookupData)

    Public Function GetLookupData(ByVal pTableName As String, ByVal pAttributeName As String, ByVal pCode As String) As String
      Dim vFound As Boolean
      Dim vLookupData As LookupData = Nothing
      For Each vLookupData In mvLookupData
        If vLookupData.TableName = pTableName AndAlso vLookupData.AttributeName = pAttributeName Then
          vFound = True
        End If
      Next
      If Not vFound Then vLookupData = New LookupData(pTableName, pAttributeName)
      If vLookupData.Items.ContainsKey(pCode) Then
        Return vLookupData.Items(pCode)
      Else
        Dim vDescription As String = GetDescription(pTableName, pAttributeName, pCode)
        vLookupData.Items.Add(pCode, vDescription)
        Return vDescription
      End If
    End Function

    Public Enum GetDocumentLocations
      gdlStandardDocument
      gdlCommunicationsLog
      gdlContactMailingDocuments
      gdlMailingHistoryDocuments
    End Enum

    Function GetDocument(ByVal pDocumentLocation As GetDocumentLocations, ByVal pValue As String, ByVal pExternal As Boolean, ByVal pExtension As String, Optional ByRef pFileName As String = "") As String
      'Get the document either from the bulk attribute in the database or from an external location
      Dim vFileName As String = ""
      Dim vValue As String = ""
      Dim vSQL As SQLStatement = Nothing
      Dim vWhereFields As New CDBFields
      Select Case pDocumentLocation
        Case GetDocumentLocations.gdlCommunicationsLog
          vSQL = New SQLStatement(Connection, "document", "communications_log", New CDBField("communications_log_number", IntegerValue(pValue)))
        Case GetDocumentLocations.gdlContactMailingDocuments
          vSQL = New SQLStatement(Connection, "document_text", "contact_mailing_documents", New CDBField("mailing_document_number", IntegerValue(pValue)))
        Case GetDocumentLocations.gdlStandardDocument
          vSQL = New SQLStatement(Connection, "standard_document_text", "standard_documents", New CDBField("standard_document", pValue))
        Case GetDocumentLocations.gdlMailingHistoryDocuments
          vSQL = New SQLStatement(Connection, "document_text", "mailing_history_documents", New CDBField("mailing_number", pValue))
      End Select
      Dim vRecordSet As CDBRecordSet = Connection.GetRecordSet(vSQL, 0, CDBConnection.RecordSetOptions.NoDataTable)
      If vRecordSet.Fetch() Then
        If pExternal Then
          'if the document is stored externally get the filename
          vFileName = vRecordSet.Fields(1).Value
          If vFileName.Length > 0 Then
            vValue = My.Computer.FileSystem.ReadAllText(vFileName)
          Else
            vRecordSet.CloseRecordSet()
            RaiseError(DataAccessErrors.daeExternalFilenameInvalid)
          End If
          'Now check that it is valid
          If UCase(Left(vValue, 9)) <> "FILENAME=" Or Len(vValue) = 9 Then
          Else
            vValue = vValue.Substring(9)
            If My.Computer.FileSystem.FileExists(vValue) Then
              If Len(pFileName) > 0 Then
                FileCopy(vValue, pFileName)
                vFileName = pFileName
              Else
                vFileName = vValue
              End If
            Else
              vRecordSet.CloseRecordSet()
              RaiseError(DataAccessErrors.daeExternalFileNotFound, vValue)
            End If
          End If
        Else
          If Len(pFileName) > 0 Then
            My.Computer.FileSystem.MoveFile(vRecordSet.Fields(1).Value, pFileName)
            vFileName = pFileName
          Else
            vFileName = vRecordSet.Fields(1).Value
          End If
        End If
      End If
      vRecordSet.CloseRecordSet()
      Return vFileName
    End Function

    Public Function TempTableName(ByVal pID As String) As String
      Dim vTableName As String = "rpt_temp_" & pID & User.Logname
      If vTableName.Length > 24 Then vTableName = vTableName.Substring(0, 24)
      Return vTableName & Format$(Now, "hhMMss")
    End Function

    Public Function GetAccountReportCode(ByVal pCode As String) As String
      'pCode should be one of CB, NL, IT, BS, PI, CT, PS
      Dim vAccountsInterface As String = TruncateString(GetControlValue(cdbControlConstants.cdbControlAccountsInterface), 4)
      Return (pCode & vAccountsInterface).ToUpper
    End Function

    Public Function AccountsReportExists(ByVal pCode As String) As Boolean
      'pCode should be one of CB, NL, IT, BS, PI, CT, PS
      Dim vAccountReportCode As String = GetAccountReportCode(pCode)
      If Mid$(vAccountReportCode, 3) <> "NONE" Then
        If Mid$(vAccountReportCode, 3) = "CHAM" Then Mid(vAccountReportCode, 3) = "CSV1"
        Return ReportExists(vAccountReportCode)
      End If
    End Function

    Public Function ReportExists(ByVal pReportCode As String) As Boolean
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("report_code", pReportCode.ToUpper)
      Return Connection.GetCount("reports", vWhereFields) > 0
    End Function

    Public Function GetAccountsFileName(ByVal pDirectory As String, ByVal pPrefix As String, ByVal pPeriod As String, Optional ByVal pJournalNo As String = "", Optional ByVal pName As String = "", Optional ByVal pUseNameOnly As Boolean = False) As String
      Dim vFileName As String
      Dim vFullFileName As String
      Dim vFullXFileName As String
      Dim vFileVersion As Integer

      Dim vDirName As String
      If pDirectory.Length = 0 Then
        vDirName = GetConfig("default_accounts_directory", "C:\contacts\accounts")
      Else
        vDirName = pDirectory
      End If
      If Not My.Computer.FileSystem.DirectoryExists(vDirName) Then RaiseError(DataAccessErrors.daeInvalidDirectory, vDirName)

      If Len(vDirName) > 0 And Right$(vDirName, 1) <> "\" Then vDirName = vDirName & "\"
      If pName.Length > 0 And pUseNameOnly = True Then
        Return vDirName & pName
      Else
        vFileName = pPrefix & "_"
        If pPeriod.Length = 1 Then vFileName = vFileName & "0"
        vFileName = vFileName & pPeriod & "_"
        vFileName = vFileName & Today.ToString("MMdd") & "_"
        vFullFileName = vDirName & vFileName
        vFullXFileName = vDirName & "x_" & vFileName        'This file is created by scripts in CHAM once the file has been loaded
        If pJournalNo.Length = 0 Or pName.Length > 0 Then
          vFileVersion = 1
          If pName.Length > 0 Then
            pJournalNo = Right$("00000000" & pJournalNo, 8)
            While My.Computer.FileSystem.FileExists(vFullFileName & vFileVersion & "_" & pJournalNo & "_" & pName)
              vFileVersion = vFileVersion + 1
            End While
            GetAccountsFileName = vFullFileName & vFileVersion & "_" & pJournalNo & "_" & pName
          Else
            While My.Computer.FileSystem.FileExists(vFullFileName & vFileVersion) Or My.Computer.FileSystem.FileExists(vFullXFileName & vFileVersion)
              vFileVersion = vFileVersion + 1
            End While
            Return vFullFileName & vFileVersion
          End If
        Else
          Return vFullFileName & pJournalNo
        End If
      End If
    End Function

    Public ReadOnly Property ConnectionID() As String
      Get
        If String.IsNullOrEmpty(mvConnectionID) Then mvConnectionID = mvUser.Logname.Substring(0, 2).ToUpper & TodaysDate.Substring(0, 2) & Date.Now.ToString("hhmmss")
        Return mvConnectionID
      End Get
    End Property

    Public Function GetSortedListOfItems(ByVal pFields As String, ByVal pTable As String) As SortedList(Of String, String)
      Dim vSQL As New SQLStatement(Connection, pFields, pTable)
      Dim vList As New SortedList(Of String, String)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      While vRS.Fetch
        vList.Add(vRS.Fields(1).Value, vRS.Fields(2).Value)
      End While
      vRS.CloseRecordSet()
      Return vList
    End Function

    Public Sub ReplaceAllODBCFunctions(ByRef pSQL As String)
      ReplaceFunction(pSQL, "length")
      ReplaceFunction(pSQL, "ADDMONTHS")
      ReplaceFunction(pSQL, "TODAY")
      ReplaceFNIfNull(pSQL)
      ReplaceFNYear(pSQL)
    End Sub

  End Class

End Namespace
