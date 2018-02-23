Namespace Access

  Partial Public Class Campaign

    Private Enum SegmentCopyTables
      sctCostCentres
      sctTickBoxes
      sctProductAllocation
    End Enum

    Private mvAppeals As Collection

    Public Function CalculateActualIncome(ByRef pProcessAll As Boolean, Optional ByRef pAppealCode As String = "", Optional ByRef pSegmentCode As String = "", Optional ByRef pCollectionNumber As Integer = 0) As Double
      Dim vSQL As String
      Dim vCollSQL As String
      Dim vSegSQL As String
      Dim vRecordSet As CDBRecordSet
      Dim vIncome As Double
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vTable As String

      vUpdateFields.AddAmendedOnBy(mvEnv.User.Logname)
      vUpdateFields.Add("actual_income_date", CDBField.FieldTypes.cftDate, TodaysDate)
      vUpdateFields.Add("actual_income", CDBField.FieldTypes.cftNumeric)

      If pProcessAll Then
        'Everything for this Campaign
        '1. Initialise all Segments, Appeals & the header to 0 since some/all Appeals may not yet have been mailed
        vWhereFields.Add("campaign", CDBField.FieldTypes.cftCharacter, CampaignCode)
        vUpdateFields(4).Value = CStr(0)
        mvEnv.Connection.UpdateRecords("campaigns", vUpdateFields, vWhereFields)
        mvEnv.Connection.UpdateRecords("appeals", vUpdateFields, vWhereFields, False)
        mvEnv.Connection.UpdateRecords("segments", vUpdateFields, vWhereFields, False)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then mvEnv.Connection.UpdateRecords("appeal_collections", vUpdateFields, vWhereFields, False)

        '2. Set Actual Income on Segments from FHD
        vUpdateFields(4).FieldType = CDBField.FieldTypes.cftLong
        vUpdateFields(4).Value = "(SELECT" & mvEnv.Connection.DBIsNull("sum(fhd.amount)", "0") & "FROM financial_history_details fhd WHERE fhd.source = segments.source)"
        mvEnv.Connection.UpdateRecords("segments", vUpdateFields, vWhereFields, False)

        '3. Set Actual Income on Appeal Collections from FHD
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
          vUpdateFields(4).Value = "(SELECT" & mvEnv.Connection.DBIsNull("sum(fhd.amount)", "0") & "FROM collection_payments cp, financial_history_details fhd WHERE cp.collection_number = appeal_collections.collection_number AND fhd.batch_number = cp.batch_number AND fhd.transaction_number = cp.transaction_number AND fhd.line_number = cp.line_number)"
          mvEnv.Connection.UpdateRecords("appeal_collections", vUpdateFields, vWhereFields, False)
        End If

        '4. Set Actual Income on Appeals from Actual Income on Segments
        vSQL = "SELECT" & mvEnv.Connection.DBIsNull("sum(actual_income)", "0") & "FROM %1 x WHERE x.campaign = appeals.campaign AND x.appeal = appeals.appeal"
        vWhereFields.Add("appeal_type", CDBField.FieldTypes.cftCharacter, "S")
        vUpdateFields(4).Value = "(" & Replace(vSQL, "%1", "segments") & ")"
        mvEnv.Connection.UpdateRecords("appeals", vUpdateFields, vWhereFields, False)

        '5. Set Actual Income on Appeals from Actual Income on Appeal Collections
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
          vWhereFields(2).WhereOperator = CDBField.FieldWhereOperators.fwoNotEqual
          vUpdateFields(4).Value = "(" & Replace(vSQL, "%1", "appeal_collections") & ")"
          mvEnv.Connection.UpdateRecords("appeals", vUpdateFields, vWhereFields, False)
        End If

        '6. Set Actual Income on Campaigns from Actual Income on Appeals
        vWhereFields.Remove((2))
        vSQL = "SELECT SUM (actual_income)  AS  sum_amount FROM appeals"
        vSQL = vSQL & " WHERE campaign = '" & CampaignCode & "'"
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          vUpdateFields(4).Value = CStr(vRecordSet.Fields(1).DoubleValue)
          mvClassFields(CampaignFields.ActualIncome).Value = CStr(vRecordSet.Fields(1).DoubleValue)
          mvClassFields(CampaignFields.ActualIncomeDate).Value = TodaysDate()
          CalculateActualIncome = ActualIncome
          'Do not save the Campaign class as there may be other changes that are not to be saved
          mvEnv.Connection.UpdateRecords("campaigns", vUpdateFields, vWhereFields)
        End If
        vRecordSet.CloseRecordSet()

      Else
        'Calculate income for either Appeal, Segment or Appeal Collection and everything below
        vSQL = "SELECT" & mvEnv.Connection.DBIsNull("SUM (fhd.amount)", "0") & "AS  sum_amount FROM %1, financial_history_details fhd WHERE "
        vCollSQL = "cp.collection_number = %2 AND cp.batch_number = fhd.batch_number AND cp.transaction_number = fhd.transaction_number AND cp.line_number = fhd.line_number"
        vSegSQL = "s.campaign = '" & CampaignCode & "'"
        If pAppealCode.Length > 0 Then vSegSQL = vSegSQL & " AND s.appeal = '" & pAppealCode & "'"
        vSegSQL = vSegSQL & " %2 AND s.source = fhd.source"
        If pCollectionNumber = 0 And Len(pAppealCode) > 0 And Len(pSegmentCode) = 0 Then
          'Appeal and everything below
          vWhereFields.Add("campaign", CDBField.FieldTypes.cftCharacter, CampaignCode)
          vWhereFields.Add("appeal", CDBField.FieldTypes.cftCharacter, pAppealCode)
          '(i) Update Segments
          vTable = "segments"
          vUpdateFields(4).FieldType = CDBField.FieldTypes.cftLong
          vUpdateFields(4).Value = "(" & Replace(vSQL, "%1", "segments s") & Replace(vSegSQL, "%2", "AND s.segment = segments.segment") & ")"
          mvEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWhereFields, False)
          '(ii) Update Appeal Collections
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
            vTable = "appeal_collections"
            vUpdateFields(4).Value = "(" & Replace(vSQL, "%1", "collection_payments cp") & Replace(vCollSQL, "%2", "appeal_collections.collection_number") & ")"
            mvEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWhereFields, False)
          End If
          '(iii) Update Appeal
          vTable = "appeals"
          vSQL = "SELECT SUM(actual_income) AS sum_income FROM %1 WHERE campaign = '" & CampaignCode & "' AND appeal = '" & pAppealCode & "'"
          vRecordSet = mvEnv.Connection.GetRecordSet(Replace(vSQL, "%1", "segments"))
          If vRecordSet.Fetch() = True Then vIncome = vRecordSet.Fields(1).DoubleValue
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then
            vRecordSet.CloseRecordSet()
            vRecordSet = mvEnv.Connection.GetRecordSet(Replace(vSQL, "%1", "appeal_collections"))
            If vRecordSet.Fetch() = True Then vIncome = vIncome + vRecordSet.Fields(1).DoubleValue
          End If
          vRecordSet.CloseRecordSet()

        Else
          If pCollectionNumber > 0 Then
            'Appeal Collection Only
            vTable = "appeal_collections"
            vSQL = Replace(vSQL, "%1", "collection_payments cp")
            vSQL = vSQL & Replace(vCollSQL, "%2", CStr(pCollectionNumber))
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            If vRecordSet.Fetch() = True Then vIncome = vRecordSet.Fields(1).DoubleValue
            vRecordSet.CloseRecordSet()
            vWhereFields.Add("collection_number", CDBField.FieldTypes.cftLong, pCollectionNumber)
          Else
            'Campaign Only OR Segment Only
            'Note: Campaign and everything below is done above as pProcessAll = True.
            If pSegmentCode.Length > 0 Then
              vTable = "segments"
            Else
              vTable = "campaigns"
            End If
            vSQL = Replace(vSQL, "%1", "segments s")
            If pSegmentCode.Length > 0 Then
              vSQL = vSQL & Replace(vSegSQL, "%2", "AND s.segment = '" & pSegmentCode & "'")
            Else
              vSQL = vSQL & Replace(vSegSQL, "%2", "")
            End If
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            If vRecordSet.Fetch() = True Then vIncome = vRecordSet.Fields(1).DoubleValue
            vRecordSet.CloseRecordSet()
            With vWhereFields
              .Add("campaign", CDBField.FieldTypes.cftCharacter, CampaignCode)
              If pAppealCode.Length > 0 Then .Add("appeal", CDBField.FieldTypes.cftCharacter, pAppealCode)
              If pSegmentCode.Length > 0 Then .Add("segment", CDBField.FieldTypes.cftCharacter, pSegmentCode)
            End With
          End If
        End If
        vUpdateFields(4).FieldType = CDBField.FieldTypes.cftNumeric
        vUpdateFields(4).Value = CStr(vIncome)
        mvEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWhereFields)
      End If
      CalculateActualIncome = vIncome
    End Function

    Public Sub CopyAppeal(ByVal pSrcAppeal As String, ByVal pTgtCampaign As String, ByVal pTgtAppeal As String, ByVal pTgtAppealDesc As String, ByVal pCopyTickBoxes As Boolean, ByVal pCopyMailingCode As Boolean, ByVal pCopySourceCode As Boolean)
      Dim vAppeal As New Appeal(mvEnv)
      Dim vTgtAppeal As New Appeal(mvEnv)
      Dim vSegment As Segment
      Dim vAppealColl As AppealCollection
      Dim vParams As New CDBParameters

      vAppeal.Init((mvClassFields.Item(CampaignFields.Campaign).Value), pSrcAppeal, True, Appeal.SegmentSortOrders.ssoSegment, True, True)
      If vAppeal.Existing Then
        If vAppeal.AppealType <> Appeal.AppealTypes.aptySegmentedMailing Then
          If DoubleValue(vAppeal.MinimumRollForwardIncome) > 0 Then
            'Re-calculate the ActualIncome on Appeal and all Collections below
            CalculateActualIncome(False, (vAppeal.AppealCode))
          End If
          vAppeal.InitAppealCollections()
        End If

        With vParams
          .Add("Campaign", pTgtCampaign)
          .Add("Appeal", pTgtAppeal)
          .Add("AppealDesc", pTgtAppealDesc)
          .Add("MailingType", vAppeal.MailingType)
          .Add("BusinessType", vAppeal.AppealBusinessType)
          .Add("Manager", vAppeal.Manager)
          .Add("RequiredCount", vAppeal.RequiredCount)
          .Add("AppealDate", vAppeal.AppealDate)
          .Add("Notes", vAppeal.Notes)
          If vAppeal.AppealType = Appeal.AppealTypes.aptySegmentedMailing Then
            .Add("TargetIncome", vAppeal.TargetIncome)
          End If
          .Add("ExpenditureGroup", vAppeal.ExpenditureGroup)
          .Add("AppealCostCentre", vAppeal.AppealCostCentre)
          .Add("Topic", vAppeal.Topic)
          .Add("SubTopic", vAppeal.SubTopic)
          If vAppeal.AppealType = Appeal.AppealTypes.aptyHouseToHouseCollection Or vAppeal.AppealType = Appeal.AppealTypes.aptyMannedCollection Then
            'Do Not Copy
          Else
            .Add("TargetResponse", vAppeal.TargetResponse)
          End If
          .Add("ThankYouLetter", vAppeal.ThankYouLetter)
          .Add("JointStatusExclusions", vAppeal.JointStatusExclusions)
          .Add("CombineMail", BooleanString(vAppeal.CombineMail))
          .Add("BypassCount", BooleanString(vAppeal.BypassCount))
          .Add("CreateMailingHistory", BooleanString(vAppeal.CreateMailingHistory))
          .Add("DefaultDespatchQuantity", vAppeal.DefaultDespatchQuantity)
          .Add("ExpectedIncome", vAppeal.ExpectedIncome)
          .Add("GuaranteedIncome", vAppeal.GuaranteedIncome)
          .Add("MailJoints", BooleanString(vAppeal.MailJoints))
          .Add("MailJointsMethod", vAppeal.MailJointsMethodCode)
          .Add("BudgetedCount", vAppeal.BudgetedCount)
          .Add("Cost", vAppeal.Cost)
          .Add("FixedCost", vAppeal.FixedCost)
          .Add("DespatchMethod", vAppeal.DespatchMethod)
          .Add("TargetResponsePercentage", vAppeal.TargetResponsePercentage)
          .Add("AppealType", vAppeal.AppealTypeCode)
          If IsDate(vAppeal.EndDate) Then .Add("EndDate", vAppeal.EndDate)
          .Add("SegmentOrgSelectionOptions", BooleanString(vAppeal.SegmentOrgSelectionOptions))
          If vAppeal.CollectionsSource.Length > 0 Then .Add("Source", vAppeal.CollectionsSource)
          If vAppeal.CollectionsProduct.Length > 0 Then .Add("Product", vAppeal.CollectionsProduct)
          If vAppeal.CollectionsRate.Length > 0 Then .Add("Rate", vAppeal.CollectionsRate)
          If vAppeal.CollectionsBankAccount.Length > 0 Then .Add("BankAccount", vAppeal.CollectionsBankAccount)
          .Add("MinimumRollForwardIncome", vAppeal.MinimumRollForwardIncome)
          .Add("ReadyForConfirmation", BooleanString(vAppeal.ReadyForConfirmation)) '?
          .Add("ReadyForAcknowledgement", BooleanString(vAppeal.ReadyForAcknowledgement)) '?
        End With
        vTgtAppeal.Create(mvEnv, vParams)
        mvEnv.Connection.StartTransaction()
        vTgtAppeal.Save()
        If vAppeal.AppealType = Appeal.AppealTypes.aptySegmentedMailing Then
          'Copy the segments associated with the source appeal to the target appeal
          For Each vSegment In vAppeal.Segments
            CopySegment(vSegment, vTgtAppeal.Campaign, vTgtAppeal.AppealCode, vSegment.SegmentCode, vSegment.SegmentDesc, CStr(vSegment.SegmentSequence), pCopyTickBoxes, pCopyMailingCode, pCopySourceCode)
          Next vSegment
          'Copy the budgets associated with the source appeal to the target appeal
          CopyBudgets(vAppeal, pTgtCampaign, pTgtAppeal)
        Else
          'Copy the AppealCollections & related data associated with the source Appeal
          For Each vAppealColl In vAppeal.AppealCollections
            CopyAppealCollection(vAppealColl, vTgtAppeal.Campaign, vTgtAppeal.AppealCode, vAppealColl.Collection, vAppeal)
          Next vAppealColl
        End If
        mvEnv.Connection.CommitTransaction()
      Else
        RaiseError(DataAccessErrors.daeCannotFindAppeal)
      End If
    End Sub

    Public Function CopyAppealCollection(ByVal pSrcAppealCollection As AppealCollection, ByVal pTgtCampaignCode As String, ByVal pTgtAppealCode As String, ByVal pTgtCollectionCode As String, Optional ByVal pSrcAppeal As Appeal = Nothing, Optional ByRef pTgtCollectionDesc As String = "") As Integer
      Dim vTgtAppCollection As New AppealCollection(mvEnv)
      Dim vCopy As Boolean
      Dim vTgtAppeal As New Appeal(mvEnv)

      vCopy = True
      If pSrcAppealCollection.Existing Then
        If Len(pTgtCollectionDesc) = 0 Then
          'copying the whole appeal
          If pSrcAppeal.Existing Then
            'Copy Criteria is only used when copying a whole appeal, not just a collection
            Select Case pSrcAppealCollection.CopyCriteria
              Case AppealCollection.AppealCollectionCopyCriteria.acccCriteriaBased
                If pSrcAppealCollection.ActualIncome < DoubleValue(pSrcAppeal.MinimumRollForwardIncome) Then vCopy = False
              Case AppealCollection.AppealCollectionCopyCriteria.acccAlwaysExclude
                vCopy = False
              Case Else 'acccAlwaysInclude
                vCopy = True
            End Select
          Else
            RaiseError(DataAccessErrors.daeCannotFindAppeal)
          End If
        Else
          'copying only one collection
          'if copying a collection make sure the target appeal is of the right type
          vTgtAppeal.Init(pTgtCampaignCode, pTgtAppealCode)
          If pSrcAppealCollection.CollectionTypeCode <> vTgtAppeal.AppealTypeCode Then
            RaiseError(DataAccessErrors.daeInvalidAppealType)
          End If
        End If

        If vCopy Then
          Select Case pSrcAppealCollection.CollectionType
            Case AppealCollection.AppealCollectionType.actHouseToHouse
              If pSrcAppealCollection.H2hCollection.Existing = False Then
                RaiseError(DataAccessErrors.daeCannotFindAppealCollection)
              End If
            Case AppealCollection.AppealCollectionType.actManned
              If pSrcAppealCollection.MannedCollection.Existing = False Then
                RaiseError(DataAccessErrors.daeCannotFindAppealCollection)
              End If
            Case AppealCollection.AppealCollectionType.actUnmanned
              If pSrcAppealCollection.UnmannedCollection.Existing = False Then
                RaiseError(DataAccessErrors.daeCannotFindAppealCollection)
              End If
          End Select
          'Clone the Collections (this will also save everything)
          vTgtAppCollection.Clone(mvEnv, pSrcAppealCollection, pTgtCampaignCode, pTgtAppealCode, pTgtCollectionCode, pTgtCollectionDesc)
          CopyAppealCollection = vTgtAppCollection.CollectionNumber
        End If
      Else
        RaiseError(DataAccessErrors.daeCannotFindAppealCollection)
      End If
    End Function
    Public Sub CopySegment(ByVal pSrcSegment As Segment, ByVal pTgtCampaign As String, ByVal pTgtAppeal As String, ByVal pTgtSegment As String, ByVal pTgtSegmentDesc As String, ByVal pTgtSequence As String, ByVal pCopyTickBoxes As Boolean, ByVal pCopyMailingCode As Boolean, ByVal pCopySourceCode As Boolean)
      Dim vSourceCount As Integer
      Dim vMailingCount As Integer
      Dim vTgtSource As String
      Dim vTgtMailing As String
      Dim vTgtCriteria As Integer
      Dim vMailingWhereFields As New CDBFields
      Dim vMailingFields As New CDBFields
      Dim vSourceWhereFields As New CDBFields
      Dim vSourceFields As New CDBFields
      Dim vTgtSegment As New Segment
      Dim vTrans As Boolean

      If pSrcSegment.Existing Then
        vTgtMailing = pSrcSegment.DeriveSegmentSource(pTgtCampaign, pTgtAppeal, pTgtSegment)
        vSourceWhereFields.Add("source")
        vMailingWhereFields.Add("mailing")
        If mvEnv.Connection.GetCount("segments", Nothing, "campaign = '" & pTgtCampaign & "' AND appeal = '" & pTgtAppeal & "' AND segment_sequence = " & pTgtSequence) > 0 Then
          RaiseError(DataAccessErrors.daeSegmentSequenceExists)
        End If
        'Insert/Update the Source Code
        'If vSegment.Mailing <> vSegment.Source Then 'non default source on segment as defined by user
        If pCopySourceCode Then
          vTgtSource = pSrcSegment.SourceCode
        Else
          vTgtSource = vTgtMailing
          'Find out if source already exists
          vSourceWhereFields(1).Value = vTgtSource
          vSourceCount = mvEnv.Connection.GetCount("sources", vSourceWhereFields)
          With vSourceFields
            .Clear()
            .AddAmendedOnBy(mvEnv.User.Logname)
            .Add("source_desc", CDBField.FieldTypes.cftCharacter, pTgtSegmentDesc)
            .Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
            .Add("incentive_scheme", CDBField.FieldTypes.cftCharacter, pSrcSegment.IncentiveScheme)
            If Len(pSrcSegment.IncentiveTriggerLevel) > 0 Then .Add("incentive_trigger_level", IntegerValue(pSrcSegment.IncentiveTriggerLevel))
            .Add("thank_you_letter", CDBField.FieldTypes.cftCharacter, pSrcSegment.ThankYouLetter)
          End With
        End If

        'Insert/Update the Mailing Code
        If pCopyMailingCode Then
          vTgtMailing = pSrcSegment.Mailing
        Else
          vMailingWhereFields(1).Value = vTgtMailing
          vMailingCount = mvEnv.Connection.GetCount("mailings", vMailingWhereFields)
          With vMailingFields
            .Clear()
            .AddAmendedOnBy(mvEnv.User.Logname)
            .Add("mailing_desc", CDBField.FieldTypes.cftCharacter, pTgtSegmentDesc)
            .Add("direction", CDBField.FieldTypes.cftCharacter, pSrcSegment.Direction)
            .Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
            .Add("marketing", CDBField.FieldTypes.cftCharacter, "Y")
            .Add("department", CDBField.FieldTypes.cftCharacter, mvEnv.User.Department)
          End With
        End If

        'Begin the transaction from here
        If mvEnv.Connection.InTransaction = False Then
          mvEnv.Connection.StartTransaction()
          vTrans = True
        End If

        'Insert/Update the Source Code
        If Not pCopySourceCode Then
          If vSourceCount > 0 Then
            mvEnv.Connection.UpdateRecords("sources", vSourceFields, vSourceWhereFields)
          Else
            vSourceFields.Add("source", CDBField.FieldTypes.cftCharacter, vTgtSource)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vSourceFields.Add("source_number", CDBField.FieldTypes.cftLong, mvEnv.GetControlNumber("SR"))
            mvEnv.Connection.InsertRecord("sources", vSourceFields)
          End If
        End If

        'Insert/Update the Mailing Code
        If Not pCopyMailingCode Then
          If vMailingCount > 0 Then
            mvEnv.Connection.UpdateRecords("mailings", vMailingFields, vMailingWhereFields)
          Else
            vMailingFields.Add("mailing", CDBField.FieldTypes.cftCharacter, vTgtMailing)
            mvEnv.Connection.InsertRecord("mailings", vMailingFields)
          End If
        End If

        'BR13349: Get the source aappeal and segment codes as pSrcSegment object will be replaced by vTgtSegment
        Dim vSrcAppealCode As String = pSrcSegment.AppealCode
        Dim vSrcSegmentCode As String = pSrcSegment.SegmentCode
        vTgtSegment = pSrcSegment
        With vTgtSegment
          .Clone(pTgtCampaign, pTgtAppeal, pTgtSegment, pTgtSegmentDesc, IntegerValue(pTgtSequence), vTgtMailing, vTgtSource)
          .Save()
        End With

        'Copy the Criteria Set
        If pSrcSegment.CriteriaSet > 0 Then
          CopyCriteria(pSrcSegment.CriteriaSet, vTgtCriteria)
          If vTgtCriteria > -1 Then
            vTgtSegment.CriteriaSet = vTgtCriteria
            vTgtSegment.Save()
          End If
        End If
        'Copy the Segment Product Allocation
        CopySegmentTable(SegmentCopyTables.sctProductAllocation, vSrcAppealCode, vSrcSegmentCode, pTgtCampaign, pTgtAppeal, pTgtSegment)
        'Copy the Segment Cost Centres
        CopySegmentTable(SegmentCopyTables.sctCostCentres, vSrcAppealCode, vSrcSegmentCode, pTgtCampaign, pTgtAppeal, pTgtSegment)
        'Copy the Tick Boxes
        If pCopyTickBoxes Then
          CopySegmentTable(SegmentCopyTables.sctTickBoxes, vSrcAppealCode, vSrcSegmentCode, pTgtCampaign, pTgtAppeal, pTgtSegment)
        End If
        'Insert the new segment record
        If vTrans Then mvEnv.Connection.CommitTransaction()
      Else
        RaiseError(DataAccessErrors.daeCannotFindSegment)
      End If
    End Sub
    Private Sub CopyBudgets(ByRef pSrcAppeal As Appeal, ByRef pTgtCampaign As String, ByRef pTgtAppeal As String)
      Dim vBudget As AppealBudget
      Dim vBudgetDetail As AppealBudgetDetail
      Dim vTgtBudget As AppealBudget
      Dim vTgtBudgetDetail As AppealBudgetDetail

      For Each vBudget In pSrcAppeal.Budgets
        vTgtBudget = New AppealBudget
        vTgtBudget.Init(mvEnv)
        vTgtBudget.Create(pTgtCampaign, pTgtAppeal, vBudget.BudgetPeriod, vBudget.PeriodStartDate, vBudget.PeriodEndDate, CStr(vBudget.PeriodPercentage))
        vTgtBudget.Save()
        For Each vBudgetDetail In vBudget.AppealBudgetDetails
          vTgtBudgetDetail = New AppealBudgetDetail
          vTgtBudgetDetail.Create(mvEnv, (vTgtBudget.AppealBudgetNumber), (vBudgetDetail.Segment), (vBudgetDetail.ReasonForDespatch), (vBudgetDetail.ForecastUnits), (vBudgetDetail.BudgetedCosts), (vBudgetDetail.BudgetedIncome))
          vTgtBudgetDetail.Save()
        Next vBudgetDetail
      Next vBudget
    End Sub

    Public Sub CopyCriteria(ByVal pSrcCriteria As Integer, ByRef pTgtCriteria As Integer)
      Dim vCriteriaSet As New CriteriaSet
      Dim vCriteriaSetDetails As New CriteriaDetails

      vCriteriaSet.Init(mvEnv, pSrcCriteria)
      If vCriteriaSet.Existing Then
        vCriteriaSet.Clone(pTgtCriteria)
      Else
        vCriteriaSetDetails.Init(mvEnv, pSrcCriteria, 1)
        If vCriteriaSetDetails.Existing Then
          vCriteriaSet.Create(vCriteriaSetDetails.CriteriaSetNumber, "", "", "", "") 'This assigns the CriteriaSet object the criteria_set_number from the CriteriaSetDetails object because that's all that's need when cloning the criteria set details
          vCriteriaSet.Clone(pTgtCriteria, True)
        End If
      End If
    End Sub

    Private Sub CopySegmentTable(ByVal pCopyType As SegmentCopyTables, ByVal pSrcAppeal As String, ByVal pSrcSegment As String, ByVal pTgtCampaign As String, ByVal pTgtAppeal As String, ByVal pTgtSegment As String)
      Dim vRS As CDBRecordSet
      Dim vTable As String = ""
      Dim vAttrs As String = ""
      Dim vFields As CDBFields
      Dim vSetAmendedOnBy As Boolean

      Select Case pCopyType
        Case SegmentCopyTables.sctCostCentres
          vTable = "segment_cost_centres"
          vAttrs = "cost_centre,cost_centre_percentage"
        Case SegmentCopyTables.sctProductAllocation
          vTable = "segment_product_allocation"
          vAttrs = "amount_number,product,rate"
          vSetAmendedOnBy = True
        Case SegmentCopyTables.sctTickBoxes
          vTable = "tick_boxes"
          vAttrs = "tick_box_number,activity,activity_value,mailing_suppression"
          vSetAmendedOnBy = True
      End Select

      vRS = mvEnv.Connection.GetRecordSet("SELECT " & vAttrs & " FROM " & vTable & " WHERE campaign = '" & mvClassFields.Item(CampaignFields.Campaign).Value & "' AND appeal = '" & pSrcAppeal & "' AND segment = '" & pSrcSegment & "'")
      While vRS.Fetch() = True
        vFields = New CDBFields
        vFields.Clone(vRS.Fields)
        If vSetAmendedOnBy Then vFields.AddAmendedOnBy(mvEnv.User.Logname)
        vFields.Add("campaign", CDBField.FieldTypes.cftCharacter, pTgtCampaign)
        vFields.Add("appeal", CDBField.FieldTypes.cftCharacter, pTgtAppeal)
        vFields.Add("segment", CDBField.FieldTypes.cftCharacter, pTgtSegment)
        mvEnv.Connection.InsertRecord(vTable, vFields)
      End While
      vRS.CloseRecordSet()
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByVal pCampaign As String, ByRef pDescription As String, ByRef pStartDate As String, ByRef pEndDate As String, Optional ByRef pManager As String = "", Optional ByRef pBusinesType As String = "", Optional ByRef pStatus As String = "", Optional ByRef pStatusDate As String = "", Optional ByRef pStatusReason As String = "", Optional ByRef pNotes As String = "", Optional ByVal pTopic As String = "")
      Init()
      With mvClassFields
        .Item(CampaignFields.Campaign).Value = pCampaign
        .Item(CampaignFields.CampaignDesc).Value = pDescription
        .Item(CampaignFields.StartDate).Value = pStartDate
        .Item(CampaignFields.EndDate).Value = pEndDate
        .Item(CampaignFields.Notes).Value = pNotes
        .Item(CampaignFields.Manager).Value = pManager
        .Item(CampaignFields.CampaignBusinessType).Value = pBusinesType
        .Item(CampaignFields.CampaignStatus).Value = pStatus
        .Item(CampaignFields.CampaignStatusDate).Value = pStatusDate
        .Item(CampaignFields.CampaignStatusReason).Value = pStatusReason
        .Item(CampaignFields.Topic).Value = pTopic
      End With
    End Sub

    Public Overloads Sub Update(ByRef pDescription As String, ByRef pStartDate As String, ByRef pEndDate As String, ByRef pManager As String, ByRef pBusinesType As String, ByRef pStatus As String, ByRef pStatusDate As String, ByRef pStatusReason As String, ByRef pNotes As String, ByRef pActualIncome As String, ByRef pActualIncomeDate As String, ByVal pTopic As String)
      With mvClassFields
        .Item(CampaignFields.CampaignDesc).Value = pDescription
        .Item(CampaignFields.StartDate).Value = pStartDate
        .Item(CampaignFields.EndDate).Value = pEndDate
        .Item(CampaignFields.Notes).Value = pNotes
        .Item(CampaignFields.Manager).Value = pManager
        .Item(CampaignFields.CampaignBusinessType).Value = pBusinesType
        .Item(CampaignFields.CampaignStatus).Value = pStatus
        .Item(CampaignFields.CampaignStatusDate).Value = pStatusDate
        .Item(CampaignFields.CampaignStatusReason).Value = pStatusReason
        If pActualIncome.Length > 0 Then .Item(CampaignFields.ActualIncome).Value = pActualIncome
        .Item(CampaignFields.ActualIncomeDate).Value = pActualIncomeDate
        .Item(CampaignFields.Topic).Value = pTopic
      End With
    End Sub

    Public Overloads Sub Init(ByRef pCampaign As String, ByVal pInitAppeals As Boolean)
      Init(mvEnv, pCampaign, pInitAppeals)
    End Sub

    Public Overloads Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCampaign As String = "", Optional ByVal pInitAppeals As Boolean = False)
      mvEnv = pEnv
      If pCampaign.Length > 0 Then
        Init(pCampaign)
        If Existing AndAlso pInitAppeals Then InitAppeals()
      Else
        Init()
      End If
    End Sub

    Private Sub InitAppeals()
      Dim vAppeal As New Appeal(mvEnv)
      vAppeal.Init()
      mvAppeals = New Collection
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAppeal.GetRecordSetFields() & " FROM appeals ap WHERE campaign = '" & CampaignCode & "'")
      While vRecordSet.Fetch() = True
        vAppeal.InitFromRecordSet(mvEnv, vRecordSet, True, Appeal.SegmentSortOrders.ssoSegment)
        mvAppeals.Add(vAppeal)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      mvEnv.Connection.StartTransaction()

      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(Campaign.CampaignFields.Campaign).Name, CampaignCode)
      mvEnv.Connection.DeleteRecords("appeals", vWhereFields, False)
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      Dim vAppeal As New Appeal(mvEnv)
      vAppeal.DeleteAssociatedData(CampaignCode)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SumAppeal(ByVal pAppeal As Appeal)
      Dim vWhereFields As New CDBFields
      With vWhereFields
        .Add("campaign", CDBField.FieldTypes.cftCharacter, pAppeal.Campaign)
        .Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal.AppealCode)
      End With

      Dim vSQL As String
      If pAppeal.AppealType = Appeal.AppealTypes.aptySegmentedMailing Then
        'Update Appeals counts from Segments
        vSQL = "SELECT SUM(target_response) AS total_responses, SUM(target_income) AS total_income, SUM(actual_count) AS total_actuals FROM segments WHERE campaign = '" & CampaignCode & "' AND appeal = '" & pAppeal.AppealCode & "'"
      Else
        'Update AppealCollections actual collectors first
        CollectionCollectorsCount(pAppeal)
        'Update Appeals counts from AppealCollections
        vSQL = "SELECT SUM(target_collectors) AS total_collectors, SUM(target_income) AS total_income, SUM(actual_collectors) AS total_actuals FROM appeal_collections WHERE campaign = '" & CampaignCode & "' AND appeal = '" & pAppeal.AppealCode & "'"
      End If
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      Dim vResponses As Integer
      Dim vIncome As Double
      Dim vActuals As Integer
      With vRecordSet
        If .Fetch() Then
          vResponses = .Fields(1).LongValue
          vIncome = .Fields(2).DoubleValue
          vActuals = .Fields(3).LongValue
        End If
        .CloseRecordSet()
      End With
      Dim vUpdateFields As New CDBFields
      With vUpdateFields
        If pAppeal.AppealType <> Appeal.AppealTypes.aptyUnmannedCollection Then .Add("target_response", CDBField.FieldTypes.cftLong, vResponses)
        .Add("target_income", CDBField.FieldTypes.cftNumeric, vIncome)
        .Add("actual_count", CDBField.FieldTypes.cftLong, vActuals)
        .Add("actual_count_date", CDBField.FieldTypes.cftDate, TodaysDate)
      End With
      pAppeal.SetTotals(vIncome, vResponses, vActuals)
      'Do not save the Appeal class as there may be other changes that are not to be saved
      mvEnv.Connection.UpdateRecords("appeals", vUpdateFields, vWhereFields)
    End Sub

    Public Sub CollectionCollectorsCount(Optional ByVal pAppeal As Appeal = Nothing, Optional ByVal pCollection As AppealCollection = Nothing)
      'Must pass in either pAppeal OR pCollection
      Dim vCollectionType As AppealCollection.AppealCollectionType

      'Set the WhereFields and CollectionType
      Dim vWhereFields As New CDBFields
      If Not pAppeal Is Nothing Then
        vWhereFields.Add("campaign", CDBField.FieldTypes.cftCharacter, pAppeal.Campaign)
        vWhereFields.Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal.AppealCode)
        Select Case pAppeal.AppealType
          Case Appeal.AppealTypes.aptyMannedCollection
            vCollectionType = AppealCollection.AppealCollectionType.actManned
          Case Appeal.AppealTypes.aptyHouseToHouseCollection
            vCollectionType = AppealCollection.AppealCollectionType.actHouseToHouse
          Case Appeal.AppealTypes.aptyUnmannedCollection
            vCollectionType = AppealCollection.AppealCollectionType.actUnmanned
        End Select
      ElseIf Not pCollection Is Nothing Then
        'If pCollectionNumber > 0 Then
        vWhereFields.Add("collection_number", CDBField.FieldTypes.cftLong, pCollection.CollectionNumber)
        vCollectionType = pCollection.CollectionType
      End If

      Dim vUpdateFields As New CDBFields
      Select Case vCollectionType
        'Update the counts on the AppealCollections
        Case AppealCollection.AppealCollectionType.actManned, AppealCollection.AppealCollectionType.actHouseToHouse
          Dim vTable As String
          If vCollectionType = AppealCollection.AppealCollectionType.actManned Then
            vTable = "manned_collectors"
          Else
            vTable = "h2h_collectors"
          End If
          vUpdateFields.Add("actual_collectors", CDBField.FieldTypes.cftLong, mvEnv.Connection.ProcessAnsiJoins("(SELECT COUNT(DISTINCT collector_number) FROM " & vTable & " mc WHERE mc.collection_number = appeal_collections.collection_number AND mc.contact_number in (SELECT DISTINCT " & mvEnv.Connection.DBIsNull("fl.contact_number", "fh.contact_number") & " FROM collection_payments cp INNER JOIN financial_history fh ON fh.batch_number = cp.batch_number AND fh.transaction_number = cp.transaction_number INNER JOIN financial_history_details fhd ON fhd.batch_number = cp.batch_number AND fhd.transaction_number = cp.transaction_number AND fhd.line_number = cp.line_number LEFT OUTER JOIN financial_links fl ON fhd.batch_number = fl.batch_number AND fhd.transaction_number = fl.transaction_number AND fhd.line_number = fl.line_number WHERE mc.collection_number = cp.collection_number AND fhd.amount > 0 AND fh.status IS NULL))"))
        Case AppealCollection.AppealCollectionType.actUnmanned
          vUpdateFields.Add("actual_collectors", CDBField.FieldTypes.cftLong, "(SELECT COUNT(DISTINCT uc.collection_number) FROM unmanned_collections uc, collection_payments cp, financial_history fh, financial_history_details fhd WHERE uc.collection_number = appeal_collections.collection_number AND uc.collection_number = cp.collection_number AND fh.batch_number = cp.batch_number AND fh.transaction_number = cp.transaction_number AND fh.status IS NULL AND fhd.batch_number = fh.batch_number AND fhd.transaction_number = fh.transaction_number AND fhd.line_number = cp.line_number AND fhd.amount > 0)")
      End Select
      vUpdateFields.Add("actual_collectors_date", CDBField.FieldTypes.cftDate, TodaysDate)
      mvEnv.Connection.UpdateRecords("appeal_collections", vUpdateFields, vWhereFields, False)
      If Not pCollection Is Nothing Then
        pCollection.Init(pCollection.CollectionNumber)
      End If
    End Sub

    Public Overrides ReadOnly Property DataTable() As CDBDataTable
      Get
        Dim vDataTable As CDBDataTable = mvClassFields.DataTable
        vDataTable.Columns("CampaignStatusDate").Name = "StatusDate"
        vDataTable.Columns("CampaignStatusReason").Name = "StatusReason"
        vDataTable.Columns("CampaignBusinessType").Name = "BusinessType"
        vDataTable.Columns("ActualIncomeDate").Name = "LastUpdated"
        Return vDataTable
      End Get
    End Property
  End Class
End Namespace
