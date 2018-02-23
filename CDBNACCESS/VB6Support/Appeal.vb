Namespace Access

  Partial Public Class Appeal

    Private mvSegments As Collection
    Private mvBudgets As Collection
    Private mvAppealCollections As Collection
    Private mvSelectionOptionsSet As Boolean
    Private mvSelectionOptions As CDBEnvironment.SelectionOptionSettings

    Public Enum SegmentSortOrders
      ssoSequence
      ssoSegment
      ssoMailing
      ssoOutputGroup
    End Enum

    Public Enum MailJointMethods
      mjmAlways 'Always convert a selected individual to their associated joint.
      mjmBothContactsSelected 'Only convert a selected individual to their associated joint when the other individual has not been explicitly excluded from the appeal due to having data that matches the standard exclusion criteria
      mjmOneSelectedOneNotExcluded 'Only convert a selected individual to their associated joint when both individuals have been selected by the appeal
      mjmNone 'Never convert a selected individual to their associated joint.
    End Enum

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      Init()
      With mvClassFields
        .Item(AppealFields.Campaign).Value = pParams("Campaign").Value
        .Item(AppealFields.Appeal).Value = pParams("Appeal").Value
        If Not pParams.Exists("Locked") Then pParams.Add("Locked", CDBField.FieldTypes.cftCharacter, "N")
        Update(pParams)
      End With
    End Sub

    Public Overloads Sub Update(ByRef pParams As CDBParameters)
      With mvClassFields
        If pParams.Exists("AppealDesc") Then .Item(AppealFields.AppealDesc).Value = pParams("AppealDesc").Value
        If pParams.Exists("MailingType") Then .Item(AppealFields.MailingType).Value = pParams("MailingType").Value
        If pParams.Exists("RequiredCount") Then .Item(AppealFields.RequiredCount).Value = pParams("RequiredCount").Value
        If pParams.Exists("ActualCount") Then .Item(AppealFields.ActualCount).Value = pParams("ActualCount").Value
        If pParams.Exists("AppealDate") Then .Item(AppealFields.AppealDate).Value = pParams("AppealDate").Value
        If pParams.Exists("Notes") Then .Item(AppealFields.Notes).Value = pParams("Notes").Value
        If pParams.Exists("TargetIncome") Then .Item(AppealFields.TargetIncome).Value = pParams("TargetIncome").Value
        If pParams.Exists("ExpenditureGroup") Then .Item(AppealFields.ExpenditureGroup).Value = pParams("ExpenditureGroup").Value
        If pParams.Exists("Manager") Then .Item(AppealFields.Manager).Value = pParams("Manager").Value
        If pParams.Exists("AppealCostCentre") Then .Item(AppealFields.AppealCostCentre).Value = pParams("AppealCostCentre").Value
        If pParams.Exists("BusinessType") Then .Item(AppealFields.AppealBusinessType).Value = pParams("BusinessType").Value
        If pParams.Exists("Topic") Then .Item(AppealFields.Topic).Value = pParams("Topic").Value
        If pParams.Exists("SubTopic") Then .Item(AppealFields.SubTopic).Value = pParams("SubTopic").Value
        If pParams.Exists("TargetResponse") Then .Item(AppealFields.TargetResponse).Value = pParams("TargetResponse").Value
        If pParams.Exists("ThankYouLetter") Then .Item(AppealFields.ThankYouLetter).Value = pParams("ThankYouLetter").Value
        If pParams.Exists("MailJoints") Then .Item(AppealFields.MailJoints).Value = pParams("MailJoints").Value
        If pParams.Exists("JointStatusExclusions") Then .Item(AppealFields.JointStatusExclusions).Value = pParams("JointStatusExclusions").Value
        If pParams.Exists("CombineMail") Then .Item(AppealFields.CombineMail).Value = pParams("CombineMail").Value
        If pParams.Exists("BypassCount") Then .Item(AppealFields.BypassCount).Value = pParams("BypassCount").Value
        If pParams.Exists("CreateMailingHistory") Then .Item(AppealFields.CreateMailingHistory).Value = pParams("CreateMailingHistory").Value
        If pParams.Exists("ExpectedIncome") Then .Item(AppealFields.ExpectedIncome).Value = pParams("ExpectedIncome").Value
        If pParams.Exists("GuaranteedIncome") Then .Item(AppealFields.GuaranteedIncome).Value = pParams("GuaranteedIncome").Value
        If pParams.Exists("DefaultDespatchQuantity") Then .Item(AppealFields.DefaultDespatchQuantity).Value = pParams("DefaultDespatchQuantity").Value
        If pParams.Exists("DespatchMethod") Then .Item(AppealFields.DespatchMethod).Value = pParams("DespatchMethod").Value
        If pParams.Exists("TargetResponsePercentage") Then .Item(AppealFields.TargetResponsePercentage).Value = pParams("TargetResponsePercentage").Value
        If pParams.Exists("MailJointsMethod") Then .Item(AppealFields.MailJointsMethod).Value = pParams("MailJointsMethod").Value
        If pParams.Exists("BudgetedCount") Then .Item(AppealFields.BudgetedCount).Value = pParams("BudgetedCount").Value
        If pParams.Exists("Cost") Then .Item(AppealFields.Cost).Value = pParams("Cost").Value
        If pParams.Exists("FixedCost") Then .Item(AppealFields.FixedCost).Value = pParams("FixedCost").Value
        If pParams.Exists("ActualIncome") Then .Item(AppealFields.ActualIncome).Value = pParams("ActualIncome").Value
        If pParams.Exists("ActualIncomeDate") Then .Item(AppealFields.ActualIncomeDate).Value = pParams("ActualIncomeDate").Value
        If pParams.Exists("AppealType") Then .Item(AppealFields.AppealType).Value = pParams("AppealType").Value
        If pParams.Exists("EndDate") Then .Item(AppealFields.EndDate).Value = pParams("EndDate").Value
        If pParams.Exists("Source") Then .Item(AppealFields.CollectionsSource).Value = pParams("Source").Value
        If pParams.Exists("Product") Then .Item(AppealFields.CollectionsProduct).Value = pParams("Product").Value
        If pParams.Exists("Rate") Then .Item(AppealFields.CollectionsRate).Value = pParams("Rate").Value
        If pParams.Exists("BankAccount") Then .Item(AppealFields.CollectionsBankAccount).Value = pParams("BankAccount").Value
        If pParams.Exists("MinimumRollForwardIncome") Then .Item(AppealFields.MinimumRollForwardIncome).Value = pParams("MinimumRollForwardIncome").Value
        If pParams.Exists("ReadyForConfirmation") Then .Item(AppealFields.ReadyForConfirmation).Value = pParams("ReadyForConfirmation").Value
        If pParams.Exists("ReadyForAcknowledgement") Then .Item(AppealFields.ReadyForAcknowledgement).Value = pParams("ReadyForAcknowledgement").Value
        If pParams.Exists("ConfirmationProducedOn") Then .Item(AppealFields.ConfirmationProducedOn).Value = pParams("ConfirmationProducedOn").Value
        If pParams.Exists("LabelsProducedOn") Then .Item(AppealFields.LabelsProducedOn).Value = pParams("LabelsProducedOn").Value
        If pParams.Exists("ResourcesProducedOn") Then .Item(AppealFields.ResourcesProducedOn).Value = pParams("ResourcesProducedOn").Value
        If pParams.Exists("EndOfCollectionProducedOn") Then .Item(AppealFields.EndOfCollectionProducedOn).Value = pParams("EndOfCollectionProducedOn").Value
        If pParams.Exists("AcknowledgementProducedOn") Then .Item(AppealFields.AcknowledgementProducedOn).Value = pParams("AcknowledgementProducedOn").Value
        If pParams.Exists("ActualCountUpdatedOn") Then .Item(AppealFields.ActualCountDate).Value = pParams("ActualCountUpdatedOn").Value
        If pParams.Exists("ActualCountDate") Then .Item(AppealFields.ActualCountDate).Value = pParams("ActualCountDate").Value
        If pParams.Exists("ReminderProducedOn") Then .Item(AppealFields.ReminderProducedOn).Value = pParams("ReminderProducedOn").Value
        If pParams.Exists("MailJoints") Then .Item(AppealFields.MailJoints).Bool = pParams("MailJoints").Bool
        If pParams.Exists("CombineMail") Then .Item(AppealFields.CombineMail).Bool = pParams("CombineMail").Bool
        If pParams.Exists("BypassCount") Then .Item(AppealFields.BypassCount).Bool = pParams("BypassCount").Bool
        If pParams.Exists("CreateMailingHistory") Then .Item(AppealFields.CreateMailingHistory).Bool = pParams("CreateMailingHistory").Bool
        If pParams.Exists("Locked") Then .Item(AppealFields.Locked).Bool = pParams("Locked").Bool
        If pParams.Exists("SegmentOrgSelectionOptions") Then .Item(AppealFields.SegmentOrgSelectionOptions).Value = pParams("SegmentOrgSelectionOptions").Value
        If pParams.Exists("MasterAction") Then .Item(AppealFields.MasterAction).Value = pParams("MasterAction").Value
        If pParams.Exists("TotalExpenditure") Then .Item(AppealFields.TotalExpenditure).DoubleValue = pParams("TotalExpenditure").DoubleValue
        If pParams.Exists("ReturnOnInvestment") Then .Item(AppealFields.ReturnOnInvestment).DoubleValue = pParams("ReturnOnInvestment").DoubleValue
      End With
    End Sub

    Public Overloads Sub Init(Optional ByRef pCampaign As String = "", Optional ByRef pAppeal As String = "", Optional ByRef pInitSegments As Boolean = False, Optional ByRef pSortOrder As SegmentSortOrders = SegmentSortOrders.ssoSequence, Optional ByRef pInitBudgets As Boolean = False, Optional ByRef pGetSourceAndMailingData As Boolean = False)
      Dim vRecordSet As CDBRecordSet

      If Len(pCampaign) > 0 And Len(pAppeal) > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields() & " FROM appeals ap WHERE campaign = '" & pCampaign & "' AND appeal = '" & pAppeal & "'")
        If vRecordSet.Fetch() Then
          InitFromRecordSet(mvEnv, vRecordSet, pInitSegments, pSortOrder, pInitBudgets, pGetSourceAndMailingData)
        Else
          InitClassFields()
          SetDefaults()
          With mvClassFields
            .Item(Appeal.AppealFields.Campaign).Value = pCampaign
            .Item(Appeal.AppealFields.Appeal).Value = pAppeal
            .Item(Appeal.AppealFields.CreateMailingHistory).Bool = True
          End With
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, Optional ByRef pInitSegments As Boolean = False, Optional ByRef pSortOrder As SegmentSortOrders = SegmentSortOrders.ssoSequence, Optional ByRef pInitBudgets As Boolean = False, Optional ByRef pGetSourceAndMailingData As Boolean = False)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(Appeal.AppealFields.Campaign, vFields)
        .SetItem(Appeal.AppealFields.Appeal, vFields)
        'Modify below to handle each recordset type as required
        .SetItem(Appeal.AppealFields.AppealDesc, vFields)
        .SetItem(Appeal.AppealFields.MailingType, vFields)
        .SetItem(Appeal.AppealFields.RequiredCount, vFields)
        .SetItem(Appeal.AppealFields.ActualCount, vFields)
        .SetItem(Appeal.AppealFields.AppealDate, vFields)
        .SetItem(Appeal.AppealFields.Notes, vFields)
        .SetItem(Appeal.AppealFields.TargetIncome, vFields)
        .SetItem(Appeal.AppealFields.Locked, vFields)
        .SetItem(Appeal.AppealFields.ExpenditureGroup, vFields)
        .SetItem(Appeal.AppealFields.Manager, vFields)
        .SetItem(Appeal.AppealFields.AppealCostCentre, vFields)
        .SetItem(Appeal.AppealFields.AppealBusinessType, vFields)
        .SetItem(Appeal.AppealFields.Topic, vFields)
        .SetItem(Appeal.AppealFields.SubTopic, vFields)
        .SetItem(Appeal.AppealFields.TargetResponse, vFields)
        .SetItem(Appeal.AppealFields.VariableParameters, vFields)
        .SetItem(Appeal.AppealFields.OrgMailTo, vFields)
        .Item(Appeal.AppealFields.CreateMailingHistory).Bool = True
        .SetOptionalItem(Appeal.AppealFields.OrgMailWhere, vFields)
        .SetOptionalItem(Appeal.AppealFields.OrgMailRoles, vFields)
        .SetOptionalItem(Appeal.AppealFields.OrgMailAddrUsage, vFields)
        .SetOptionalItem(Appeal.AppealFields.OrgMailLabelName, vFields)
        .SetOptionalItem(Appeal.AppealFields.CreateMailingHistory, vFields)
        .SetItem(Appeal.AppealFields.ThankYouLetter, vFields)
        .SetItem(Appeal.AppealFields.MailJoints, vFields)
        .SetItem(Appeal.AppealFields.JointStatusExclusions, vFields)
        .SetItem(Appeal.AppealFields.CombineMail, vFields)
        .SetItem(Appeal.AppealFields.BypassCount, vFields)
        .SetItem(Appeal.AppealFields.DefaultDespatchQuantity, vFields)
        .SetItem(Appeal.AppealFields.AmendedOn, vFields)
        .SetItem(Appeal.AppealFields.AmendedBy, vFields)
        .SetItem(Appeal.AppealFields.MasterAction, vFields)
        .SetItem(Appeal.AppealFields.ExpectedIncome, vFields)
        .SetItem(Appeal.AppealFields.GuaranteedIncome, vFields)
        .SetOptionalItem(Appeal.AppealFields.MailJointsMethod, vFields)
        .SetOptionalItem(Appeal.AppealFields.BudgetedCount, vFields)
        .SetOptionalItem(Appeal.AppealFields.Cost, vFields)
        .SetOptionalItem(Appeal.AppealFields.FixedCost, vFields)
        .SetOptionalItem(Appeal.AppealFields.ActualIncome, vFields)
        .SetOptionalItem(Appeal.AppealFields.ActualIncomeDate, vFields)
        .SetOptionalItem(Appeal.AppealFields.TargetResponsePercentage, vFields)
        .SetOptionalItem(Appeal.AppealFields.DespatchMethod, vFields)
        .SetOptionalItem(Appeal.AppealFields.AppealType, vFields)
        .SetOptionalItem(Appeal.AppealFields.EndDate, vFields)
        .SetOptionalItem(Appeal.AppealFields.SegmentOrgSelectionOptions, vFields)
        .SetOptionalItem(Appeal.AppealFields.CollectionsSource, vFields)
        .SetOptionalItem(Appeal.AppealFields.CollectionsProduct, vFields)
        .SetOptionalItem(Appeal.AppealFields.CollectionsRate, vFields)
        .SetOptionalItem(Appeal.AppealFields.CollectionsBankAccount, vFields)
        .SetOptionalItem(Appeal.AppealFields.MinimumRollForwardIncome, vFields)
        .SetOptionalItem(Appeal.AppealFields.ReadyForConfirmation, vFields)
        .SetOptionalItem(Appeal.AppealFields.ReadyForAcknowledgement, vFields)
        .SetOptionalItem(Appeal.AppealFields.ConfirmationProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.LabelsProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.ResourcesProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.EndOfCollectionProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.AcknowledgementProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.ActualCountDate, vFields)
        .SetOptionalItem(Appeal.AppealFields.ReminderProducedOn, vFields)
        .SetOptionalItem(Appeal.AppealFields.TotalExpenditure, vFields)
        .SetOptionalItem(Appeal.AppealFields.ReturnOnInvestment, vFields)
        .SetOptionalItem(Appeal.AppealFields.TotalItemisedCost, vFields)
      End With
      If pInitSegments Then InitSegments(pSortOrder, pGetSourceAndMailingData)
      If pInitBudgets Then InitBudgets()
    End Sub

    Public Sub InitSegments(ByRef pSortOrder As SegmentSortOrders, Optional ByRef pGetSourceAndMailingData As Boolean = False)
      Dim vSegment As New Segment
      Dim vRecordSet As CDBRecordSet
      Dim vSortOrder As String = ""
      Dim vKey As String
      Dim vSQL As String

      Select Case pSortOrder
        Case SegmentSortOrders.ssoMailing
          vSortOrder = "mailing"
        Case SegmentSortOrders.ssoSegment
          vSortOrder = "segment"
        Case SegmentSortOrders.ssoSequence
          vSortOrder = "segment_sequence"
        Case SegmentSortOrders.ssoOutputGroup
          vSortOrder = "output_group"
      End Select

      mvSegments = New Collection
      vSegment.Init(mvEnv, "", "", "", True)
      vSQL = "SELECT " & vSegment.GetRecordSetFields(Segment.SegmentRecordSetTypes.srtAll)
      If pGetSourceAndMailingData Then vSQL = vSQL & ",source_desc,incentive_trigger_level,thank_you_letter,incentive_scheme,distribution_code,discount_percentage,mailing_desc,direction,m.notes AS m_notes,department"
      vSQL = vSQL & " FROM segments sg"
      If pGetSourceAndMailingData Then vSQL = vSQL & ",sources so, mailings m "
      vSQL = vSQL & " WHERE campaign = '" & mvClassFields.Item(Appeal.AppealFields.Campaign).Value & "' AND appeal = '" & mvClassFields.Item(Appeal.AppealFields.Appeal).Value & "'"
      If pGetSourceAndMailingData Then vSQL = vSQL & "AND sg.source = so.source AND sg.mailing = m.mailing "
      vSQL = vSQL & " ORDER BY " & vSortOrder
      vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      With vRecordSet
        While .Fetch()
          vSegment = New Segment
          vSegment.InitFromRecordSet(mvEnv, vRecordSet, Segment.SegmentRecordSetTypes.srtAll, pGetSourceAndMailingData)
          mvSegments.Add(vSegment, vSegment.CampaignCode & vSegment.AppealCode & vSegment.SegmentCode)
        End While
        .CloseRecordSet()
      End With
      vSortOrder = "s." & vSortOrder
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT campaign,appeal,segment,report_code,create_mailing_history,create_selection_set,create_output FROM segments s, mailings m, output_groups og WHERE campaign = '" & mvClassFields.Item(Appeal.AppealFields.Campaign).Value & "' AND appeal = '" & mvClassFields.Item(Appeal.AppealFields.Appeal).Value & "' AND s.mailing = m.mailing AND m.direction = 'O' AND s.output_group IS NOT NULL AND s.output_group = og.output_group ORDER BY " & vSortOrder)
      With vRecordSet
        While .Fetch()
          vKey = .Fields("campaign").Value & .Fields("appeal").Value & .Fields("segment").Value
          vSegment = DirectCast(mvSegments.Item(vKey), Segment)
          vSegment.InitOutputGroup(.Fields)
        End While
        .CloseRecordSet()
      End With
    End Sub

    Public Sub InitAppealCollections(Optional ByVal pCollectionNumber As Integer = 0)
      'pCollectionNumber only used when need to select an individual Collection (used by CopyAppealCollection)
      'Select - AppealCollections
      '       - H2hCollection / MannedCollection / UnMannedCollection
      '         - CollectionRegions
      '           - CollectionPoints
      Dim vRS As CDBRecordSet
      Dim vRSType As AppealCollection.AppealCollectionRecordSetTypes
      Dim vAppealCollection As New AppealCollection(mvEnv)
      Dim vCollPoint As New CollectionPoint
      Dim vCollRegion As New CollectionRegion
      Dim vWhereFields As New CDBFields
      Dim vSQL As String = ""
      Dim vTableName As String = ""
      Dim vAttrName As String = ""

      mvAppealCollections = New Collection
      vAppealCollection.Init()
      vCollPoint.Init(mvEnv)
      vCollRegion.Init(mvEnv)

      vWhereFields.Add("campaign", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(AppealFields.Campaign).Value)
      vWhereFields.Add("appeal", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(AppealFields.Appeal).Value)
      If pCollectionNumber > 0 Then vWhereFields.Add("ac.collection_number", CDBField.FieldTypes.cftLong, pCollectionNumber)

      Select Case AppealType
        Case AppealTypes.aptyHouseToHouseCollection
          vSQL = vAppealCollection.GetRecordSetFields(AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtHouseToHouse)
          vSQL = Replace(vSQL, "hc.collection_number,", "")
          vTableName = "h2h_collections hc"
          vRSType = AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtHouseToHouse
          vAttrName = "hc.collection_number"
        Case AppealTypes.aptyMannedCollection
          vSQL = vAppealCollection.GetRecordSetFields(AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtManned)
          vSQL = Replace(vSQL, "mc.collection_number,", "")
          vTableName = "manned_collections mc"
          vRSType = AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtManned
          vAttrName = "mc.collection_number"
        Case AppealTypes.aptyUnmannedCollection
          vSQL = vAppealCollection.GetRecordSetFields(AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtUnmanned)
          vSQL = Replace(vSQL, "uc.collection_number,", "")
          vTableName = "unmanned_collections uc"
          vRSType = AppealCollection.AppealCollectionRecordSetTypes.apcrtAll Or AppealCollection.AppealCollectionRecordSetTypes.apcrtUnmanned
          vAttrName = "uc.collection_number"
      End Select
      vRSType = vRSType Or AppealCollection.AppealCollectionRecordSetTypes.apcrtCollectionWithRegion
      vSQL = vSQL & ", " & Replace(Replace(Replace(vCollRegion.GetRecordSetFields(CollectionRegion.CollectionRegionRecordSetTypes.crertAll Or CollectionRegion.CollectionRegionRecordSetTypes.crertAllPlusPoints) & ",", "amended_by,", ""), "amended_on,", ""), "cr.collection_number,", "") 'Remove extra amended_by/on and cr.collection_number
      If Right(vSQL, 1) = "," Then vSQL = Left(vSQL, Len(vSQL) - 1)

      If Len(vSQL) > 0 Then
        vSQL = "SELECT " & vSQL
        vSQL = vSQL & " FROM appeal_collections ac INNER JOIN " & vTableName & " ON " & vAttrName & " = ac.collection_number"
        vSQL = vSQL & " LEFT OUTER JOIN collection_regions cr ON ac.collection_number = cr.collection_number"
        vSQL = vSQL & " LEFT OUTER JOIN collection_points cp ON cr.collection_region_number = cp.collection_region_number"
        vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY ac.collection_number"
        vRS = mvEnv.Connection.GetRecordSetAnsiJoins(vSQL)
        If vRS.Fetch Then
          While vRS.Status
            vAppealCollection = New AppealCollection(mvEnv)
            vAppealCollection.InitFromRecordSet(mvEnv, vRS, vRSType)
            mvAppealCollections.Add(vAppealCollection, CStr(vAppealCollection.CollectionNumber))
          End While
        End If
        vRS.CloseRecordSet()
      End If
    End Sub

    Public Sub InitBudgets()
      Dim vRecordSet As CDBRecordSet
      Dim vBudget As New AppealBudget

      mvBudgets = New Collection
      vBudget.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vBudget.GetRecordSetFields(AppealBudget.AppealBudgetRecordSetTypes.abrtAll) & " FROM appeal_budgets where campaign = '" & Campaign & "' AND appeal = '" & AppealCode & "' ORDER BY budget_period")
      While vRecordSet.Fetch()
        vBudget = New AppealBudget
        vBudget.InitFromRecordSet(mvEnv, vRecordSet, AppealBudget.AppealBudgetRecordSetTypes.abrtAll)
        mvBudgets.Add(vBudget, CStr(vBudget.AppealBudgetNumber))
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property Budgets() As Collection
      Get
        Budgets = mvBudgets
      End Get
    End Property

    Public ReadOnly Property Segments() As Collection
      Get
        Segments = mvSegments
      End Get
    End Property

    Public ReadOnly Property AppealCollections() As Collection
      Get
        If mvAppealCollections Is Nothing Then InitAppealCollections()
        AppealCollections = mvAppealCollections
      End Get
    End Property

    Public ReadOnly Property MailJointsMethod() As MailJointMethods
      Get
        Select Case mvClassFields.Item(AppealFields.MailJointsMethod).Value
          Case "A"
            Return MailJointMethods.mjmAlways
          Case "B"
            Return MailJointMethods.mjmBothContactsSelected
          Case "N"
            Return MailJointMethods.mjmOneSelectedOneNotExcluded
          Case ""
            If MailJoints Then
              Return MailJointMethods.mjmAlways
            Else
              Return MailJointMethods.mjmNone
            End If
        End Select
      End Get
    End Property

    Public ReadOnly Property SelectionOptions() As CDBEnvironment.SelectionOptionSettings
      Get
        Dim vCount As Integer

        If Not mvSelectionOptionsSet Then
          mvSelectionOptionsSet = True
          With mvClassFields
            If OrgMailTo.Length > 0 Then vCount += 1
            If OrgMailWhere.Length > 0 Then vCount += 1
            If OrgMailAddrUsage.Length > 0 Then vCount += 1
            If OrgMailRoles.Length > 0 Then vCount += 1
            If OrgMailLabelName.Length > 0 Then vCount += 1
            If VariableParameters.Length > 0 And VariableParameters <> "||" Then vCount += 1
          End With
          Select Case vCount
            Case 0
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosNone
            Case 6
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosAll
            Case Else
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosSome
          End Select
        End If
        Return mvSelectionOptions
      End Get
    End Property
    Public ReadOnly Property SegmentSelectionOptions() As CDBEnvironment.SelectionOptionSettings
      Get
        Dim vCount As Integer
        For Each vSegment As Segment In Segments
          If vSegment.SelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then vCount = vCount + 1
        Next vSegment
        If vCount = 0 Then
          Return CDBEnvironment.SelectionOptionSettings.sosNone
        ElseIf vCount = Segments.Count() Then
          Return CDBEnvironment.SelectionOptionSettings.sosAll
        Else
          Return CDBEnvironment.SelectionOptionSettings.sosSome
        End If
      End Get
    End Property

    Public Sub LockAppeal()
      mvClassFields.Item(AppealFields.Locked).Bool = True
    End Sub

    Public Sub UnlockAppeal()
      mvClassFields.Item(AppealFields.Locked).Bool = False
    End Sub

    Public Sub SetTotals(ByVal pTargetIncome As Double, ByVal pTargetResponse As Integer, ByVal pActualCount As Integer)
      mvClassFields(AppealFields.TargetIncome).DoubleValue = pTargetIncome
      If AppealType <> AppealTypes.aptyUnmannedCollection Then mvClassFields(AppealFields.TargetResponse).DoubleValue = pTargetResponse
      mvClassFields(AppealFields.ActualCount).DoubleValue = pActualCount
      mvClassFields(AppealFields.ActualCountDate).Value = TodaysDate()
    End Sub

    Public Sub UpdateOrganisationMailDetails(ByRef pParams As CDBParameters)
      With mvClassFields
        If pParams.Exists("OrgMailTo") Then .Item(AppealFields.OrgMailTo).Value = pParams("OrgMailTo").Value
        If pParams.Exists("OrgMailWhere") Then .Item(AppealFields.OrgMailWhere).Value = pParams("OrgMailWhere").Value
        If pParams.Exists("OrgAddressUsage") Then .Item(AppealFields.OrgMailAddrUsage).Value = pParams("OrgAddressUsage").Value
        If pParams.Exists("OrgLabelName") Then .Item(AppealFields.OrgMailLabelName).Value = pParams("OrgLabelName").Value
        If OrgMailTo = "O" Then
          If pParams.Exists("OrgRoles") Then .Item(AppealFields.OrgMailRoles).Value = pParams("OrgRoles").Value
        Else
          .Item(AppealFields.OrgMailRoles).Value = ""
        End If
      End With
    End Sub

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvSegments = New Collection
      mvBudgets = New Collection
      mvAppealCollections = Nothing
    End Sub

  End Class
End Namespace
