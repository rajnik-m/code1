Namespace Access

  Partial Public Class MailingTemplate

    Public Sub InitFromMailing(ByVal pEnv As CDBEnvironment, ByVal pMailing As String)
      Dim vRecordSet As CDBRecordSet
      Dim vMTP As New MailingTemplateParagraph(mvEnv)

      mvEnv = pEnv
      vMTP.Init()
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields() & "," & vMTP.GetRecordSetFieldsDetail() & " FROM mailings m, mailing_templates mt, mailing_template_paragraphs mtp WHERE mailing = '" & pMailing & "' AND m.mailing_template = mt.mailing_template AND mt.mailing_template = mtp.mailing_template ORDER BY sequence_number")
      If vRecordSet.Fetch() Then
        InitFromRecordSet(vRecordSet)
        mvParagraphs = New CollectionList(Of MailingTemplateParagraph)
        Do
          vMTP = New MailingTemplateParagraph(mvEnv)
          vMTP.InitFromRecordSet(vRecordSet)
          mvParagraphs.Add(vMTP.ParagraphNumber.ToString, vMTP)
        Loop While vRecordSet.Fetch
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub CheckParagraphConditions(ByVal pExisting As Boolean, ByVal pContact As Contact, ByVal pBT As BatchTransaction, ByVal pPaymentPlan As PaymentPlan, ByVal pMailingCode As String, ByVal pGiftAidDeclaration As GiftAidDeclaration, ByVal pGAYEPledge As PreTaxPledge, ByVal pNewPayerContact As Boolean, ByVal pOrganisation As Organisation, Optional ByVal pSetBTReceipt As Boolean = True, Optional ByVal pPaymentPlanCreatedSmartClient As Boolean = False, Optional ByVal pAutoPaymentCreatedSmartClient As Boolean = False)
      Dim vMTP As MailingTemplateParagraph
      Dim vCheckTransaction As Boolean
      Dim vCheckAnalysis As Boolean
      Dim vBTA As BatchTransactionAnalysis
      Dim vPPD As PaymentPlanDetail
      Dim vWhereFields As New CDBFields
      Dim vActivityTable As String

      mvMailingDoc = New ContactMailingDocument(mvEnv)
      If pExisting Then 'Get the existing record if available
        mvMailingDoc.InitFromTransaction(pBT.BatchNumber, pBT.TransactionNumber)
      Else 'Else intialise a new one
        mvMailingDoc.Init()
      End If
      If pBT.Existing And pSetBTReceipt Then
        pBT.Receipt = "M"
        pBT.SaveChanges()
      End If

      'Now reset the template details in case it changed
      mvMailingDoc.InitFromTemplate(Me, pBT, pPaymentPlan, pMailingCode, pContact, pGiftAidDeclaration, pGAYEPledge, pNewPayerContact, pOrganisation)

      For Each vMTP In mvParagraphs
        vMTP.Include = False
        Select Case vMTP.ParagraphCondition
          Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcUserDepartmentCode
            If mvEnv.User.Department = vMTP.ControlValue Then vMTP.Include = True
          Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisDistributionCode, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisSourceCode, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisProduct, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisLineType, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcEnclosureProduct
            vCheckTransaction = True
            vCheckAnalysis = True
          Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcTransactionPaymentMethod, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcTransactionType
            vCheckTransaction = True
        End Select
      Next vMTP
      If pBT.Existing And vCheckTransaction And pBT.Analysis.Count() = 0 Then
        If vCheckAnalysis Then pBT.InitBatchTransactionAnalysis(pBT.BatchNumber, pBT.TransactionNumber)
      End If

      For Each vMTP In mvParagraphs
        With vMTP
          Select Case .ParagraphCondition
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcHasGiftAidDeclaration
              If Not pContact Is Nothing Then .Include = pContact.HasValidGiftAidDeclaration
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcNoGiftAidDeclaration
              If Not pContact Is Nothing Then .Include = Not (pContact.HasValidGiftAidDeclaration)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisDistributionCode
              For Each vBTA In pBT.Analysis
                If vBTA.DistributionCode = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vBTA
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisSourceCode
              For Each vBTA In pBT.Analysis
                If vBTA.Source = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vBTA
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisProduct
              For Each vBTA In pBT.Analysis
                If vBTA.ProductCode = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vBTA
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcAnalysisLineType
              For Each vBTA In pBT.Analysis
                If vBTA.LineType = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vBTA
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcTransactionPaymentMethod
              .Include = (pBT.PaymentMethod = .ControlValue)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcTransactionType
              .Include = (pBT.TransactionType = .ControlValue)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanProduct
              For Each vPPD In pPaymentPlan.Details
                If vPPD.ProductCode = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vPPD
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanDistributionCode
              For Each vPPD In pPaymentPlan.Details
                If vPPD.DistributionCode = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vPPD
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanSourceCode
              If pPaymentPlan.Source = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanPaymentMethod
              If pPaymentPlan.PaymentMethod = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanPaymentFrequency
              If pPaymentPlan.PaymentFrequencyCode = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcDirectDebitSourceCode
              If pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes And pPaymentPlan.AutoPaymentSource = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcStandingOrderSourceCode
              If pPaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes And pPaymentPlan.AutoPaymentSource = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcCardAuthoritySourceCode
              If pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes And pPaymentPlan.AutoPaymentSource = .ControlValue Then .Include = True
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcNewPaymentPlan
              .Include = pPaymentPlan.Created Or pPaymentPlanCreatedSmartClient
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcExistingPaymentPlan
              .Include = pPaymentPlan.PlanNumber > 0 And (pPaymentPlan.Created = False And pPaymentPlanCreatedSmartClient = False)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcNewDirectDebit
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcExistingDirectDebit
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = False And pAutoPaymentCreatedSmartClient = False)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcNewStandingOrder
              .Include = (pPaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcExistingStandingOrder
              .Include = (pPaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = False And pAutoPaymentCreatedSmartClient = False)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcNewCreditCardAuthority
              .Include = (pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcExistingCreditCardAuthority
              .Include = (pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = False And pAutoPaymentCreatedSmartClient = False)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcCAFStandingOrder
              .Include = (pPaymentPlan.StandingOrderStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCAF = True)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcCAFCreditCardAuthority
              .Include = (pPaymentPlan.CreditCardStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCAF = True)
              'Following added for WWF Fulfilment mods
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcHasOralGiftAidDeclaration
              If Not pContact Is Nothing Then .Include = pContact.HasValidGiftAidDeclaration(GiftAidDeclaration.GiftAidDeclarationMethods.gadmOral)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcHasWrittenGiftAidDeclaration
              If Not pContact Is Nothing Then .Include = pContact.HasValidGiftAidDeclaration(GiftAidDeclaration.GiftAidDeclarationMethods.gadmWritten)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcHasElectronicGiftAidDeclaration
              If Not pContact Is Nothing Then .Include = pContact.HasValidGiftAidDeclaration(GiftAidDeclaration.GiftAidDeclarationMethods.gadmElectronic)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcExistingCovenant
              .Include = (pPaymentPlan.CovenantStatus <> PaymentPlan.ppCovenant.ppcNo And Not pPaymentPlan.CovenantCreated)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcEnclosureProduct
              'Look for the specified product as an I-type line in the transaction
              For Each vBTA In pBT.Analysis
                If vBTA.LineType = "I" And vBTA.ProductCode = .ControlValue Then
                  .Include = True
                  Exit For
                End If
              Next vBTA
              'Now look for the specified product in contact_incentives
              If Not .Include And mvMailingDoc.NewContact Then
                vWhereFields = New CDBFields
                vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, mvMailingDoc.ContactNumber)
                vWhereFields.Add("date_fulfilled", CDBField.FieldTypes.cftDate)
                vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, .ControlValue)
                .Include = mvEnv.Connection.GetCount("contact_incentives", vWhereFields, "") > 0
              End If
              'Now look for the specified product in enclosures
              If Not .Include Then
                vWhereFields = New CDBFields
                vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, mvMailingDoc.PaymentPlanNumber)
                vWhereFields.Add("date_fulfilled", CDBField.FieldTypes.cftDate)
                vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, .ControlValue)
                .Include = mvEnv.Connection.GetCount("enclosures", vWhereFields, "") > 0
              End If
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcMembershipType
              .Include = (pPaymentPlan.PlanType = CDBEnvironment.ppType.pptMember And pPaymentPlan.MembershipTypeCode = .ControlValue)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcEventBookingPayment
              vWhereFields = New CDBFields
              vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, pBT.BatchNumber)
              vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, pBT.TransactionNumber)
              .Include = mvEnv.Connection.GetCount("event_bookings", vWhereFields, "") > 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcTransactionAmount
              .Include = (pBT.Amount >= Val(.ControlValue) And pBT.Amount <= Val(.ControlValue2))
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcCurrentActivityAndValue, MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcUpdateActivityAndValue
              vWhereFields = New CDBFields
              If Not pOrganisation Is Nothing Then
                vActivityTable = "organisation_categories"
                vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, pOrganisation.OrganisationNumber)
              Else
                vActivityTable = "contact_categories"
                vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContact.ContactNumber)
              End If
              vWhereFields.Add("activity", CDBField.FieldTypes.cftCharacter, .ControlValue)
              vWhereFields.Add("activity_value", CDBField.FieldTypes.cftCharacter, .ControlValue2)
              If .ParagraphCondition = MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcCurrentActivityAndValue Then
                'This is an attempt at including only those contact/organisation_categories records that are current, i.e. valid_to >= the current date.
                vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
              Else
                'This is an attempt at including only those contact/organisation_categories records that have just been created, i.e. valid_from <= the current date and amended_on = the current date.
                vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoLessThanEqual)
                vWhereFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
              End If
              .Include = mvEnv.Connection.GetCount(vActivityTable, vWhereFields, "") > 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcOwnershipGroup
              .Include = mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups And pContact.OwnershipGroup = .ControlValue
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcHasNominatedGiver
              .Include = Val(pPaymentPlan.GiverContactNumber) > 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPayerGeographicRegion
              vWhereFields = New CDBFields
              vWhereFields.Add("geographical_region_type", CDBField.FieldTypes.cftCharacter, .ControlValue)
              vWhereFields.Add("postcode", CDBField.FieldTypes.cftCharacter, pContact.Address.Postcode)
              vWhereFields.Add("geographical_region", CDBField.FieldTypes.cftCharacter, .ControlValue2)
              .Include = mvEnv.Connection.GetCount("address_geographical_regions", vWhereFields, "") > 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaperlessDirectDebitMandate
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True) And (pPaymentPlan.DirectDebit.MandateType = "P")
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcWrittenDirectDebitMandate
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True) And (pPaymentPlan.DirectDebit.MandateType = "W")
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcUnknownDirectDebitMandate
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.AutoPaymentCreated = True Or pAutoPaymentCreatedSmartClient = True) And Len(pPaymentPlan.DirectDebit.MandateType) = 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcMembershipTypePaidByDD
              'PP paid by DD and linked to membership with same type as control value 1
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppYes) And (pPaymentPlan.MembershipTypeCode = .ControlValue)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcMembershipTypePaidByNonDDPayment
              'PP not paid by DD and linked to membership with same type as control value 1
              .Include = (pPaymentPlan.DirectDebitStatus = PaymentPlan.ppYesNoCancel.ppNo) And (pPaymentPlan.MembershipTypeCode = .ControlValue)
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPayerRegisteredUser
              'payer on mailing doc is user in registered_users table
              vWhereFields = New CDBFields
              vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, mvMailingDoc.ContactNumber)
              .Include = mvEnv.Connection.GetCount("registered_users", vWhereFields, "") > 0
            Case MailingTemplateParagraph.MailingTemplateParagraphConditions.mtpcPaymentPlanProductRate
              'PP contains detail lines with product = control value 1, Rate = control value 2, created_on is todays date
              vWhereFields = New CDBFields
              vWhereFields.Add("order_number", CDBField.FieldTypes.cftInteger, pPaymentPlan.PlanNumber)
              vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, .ControlValue)
              vWhereFields.Add("rate", CDBField.FieldTypes.cftCharacter, .ControlValue2)
              vWhereFields.Add("created_on", CDBField.FieldTypes.cftDate, TodaysDate())
              .Include = mvEnv.Connection.GetCount("order_details", vWhereFields, "") > 0
          End Select
        End With
      Next vMTP
      SetSelectedParagraphs()
    End Sub

    Public Sub SetSelectedParagraphs()
      Dim vMTP As MailingTemplateParagraph
      Dim vSelectedList As String = ""

      For Each vMTP In mvParagraphs
        With vMTP
          If .Include Then
            If vSelectedList.Length > 0 Then vSelectedList = vSelectedList & ","
            vSelectedList = vSelectedList & .ParagraphNumber
          End If
        End With
      Next vMTP
      mvMailingDoc.SelectedParagraphs = vSelectedList
    End Sub

    Public ReadOnly Property ParagraphsDataTable() As CDBDataTable
      Get
        Dim vMTP As MailingTemplateParagraph
        Dim vDT As New CDBDataTable

        vDT.AddColumnsFromList("Include,ParagraphNumber,ParagraphDesc,BookmarkName,Mandatory")
        For Each vMTP In mvParagraphs
          With vMTP
            vDT.AddRowFromList(If(.Include, "Y,", "N,") & vMTP.ParagraphNumber.ToString & "," & vMTP.ParagraphDesc.Replace(",", "") & "," & vMTP.BookmarkName & "," & If(vMTP.Mandatory, "Y", "N"))
          End With
        Next vMTP
        Return vDT
      End Get
    End Property
  End Class
End Namespace
