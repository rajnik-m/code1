Namespace Access
  Public Class Segment

    Public Enum SegmentRecordSetTypes 'These are bit values
      srtAll = &HFFFFS
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SegmentFields
      sfAll = 0
      sfCampaign
      sfAppeal
      sfSegment
      sfSegmentDesc
      sfSegmentType
      sfSegmentSequence
      sfCriteriaSet
      sfSelectionSet
      sfNotes
      sfOutputGroup
      sfRequiredCount
      sfActualCount
      sfActualIncome
      sfActualIncomeDate
      sfCountTime
      sfCost
      sfFixedCost
      sfRandom
      sfMailing
      sfScore
      sfSource
      sfAmendedBy
      sfTargetIncome
      sfTargetResponse
      sfSegmentDate
      sfSegmentTime
      sfSegmentMediaType
      sfSegmentMediaName
      sfSegmentGeographic
      sfSegmentAdPagePosition
      sfSegmentAdPublPosition
      sfSegmentAdSize
      sfSegmentAdImage
      sfSegmentAdColour
      sfSegmentLetterSize
      sfSegmentPhoneType
      sfAmendedOn
      sfMailingNumber
      sfUpdatedByRp
      sfDespatchQuantity
      sfSegmentCreative
      sfBudgetedCount
      sfOrgMailTo
      sfOrgMailWhere
      sfOrgMailRoles
      sfOrgMailAddrUsage
      sfOrgMailLabelName
      sfTotalItemisedCost
      sfTelemarketing
    End Enum

    Private Structure MailingValues
      Dim mvsMailing As String
      Dim mvsDescription As String
      Dim mvsDirection As String
      Dim mvsNotes As String
      Dim mvsDepartment As String
      Dim mvsLogname As String
    End Structure

    Private Structure SourceValues
      Dim svsSource As String
      Dim svsSourceDesc As String
      Dim svsIncentive As String
      Dim svsTrigger As String
      Dim svsThankYouLetter As String
      Dim svsDistributionCode As String
      Dim svsDiscountPercentage As String
    End Structure

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvReportCode As String
    Private mvCreateMailingHistory As Boolean
    Private mvCreateSelectionSet As Boolean
    Private mvCreateOutput As Boolean
    Private mvAmendedValid As Boolean
    Private mvSourceValuesSet As Boolean
    Private mvMailingValuesSet As Boolean
    Private mvMailingValues As MailingValues
    Private mvSourceValues As SourceValues
    Private mvCostCentres As Collection
    Private mvProducts As Collection
    Private mvOrgMailToDesc As String
    Private mvOrgMailWhereDesc As String
    Private mvRoleDescriptions As CDBDataTable
    Private mvSelectionOptions As CDBEnvironment.SelectionOptionSettings
    Private mvSelectionOptionsSet As Boolean
    Private mvPrefixRequired As Boolean

    Public ReadOnly Property CreateMailingHistory() As Boolean
      Get
        CreateMailingHistory = mvCreateMailingHistory
      End Get
    End Property
    Public ReadOnly Property CreateOutput() As Boolean
      Get
        CreateOutput = mvCreateOutput
      End Get
    End Property
    Public ReadOnly Property CreateSelectionSet() As Boolean
      Get
        CreateSelectionSet = mvCreateSelectionSet
      End Get
    End Property

    Public ReadOnly Property ReportCode() As String
      Get
        ReportCode = mvReportCode
      End Get
    End Property

    Public Property CriteriaSet() As Integer
      Get
        CriteriaSet = mvClassFields.Item(SegmentFields.sfCriteriaSet).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SegmentFields.sfCriteriaSet).Value = CStr(Value)
      End Set
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property
    Public ReadOnly Property ActualIncome() As Double
      Get
        ActualIncome = mvClassFields.Item(SegmentFields.sfActualIncome).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ActualIncomeDate() As String
      Get
        ActualIncomeDate = mvClassFields.Item(SegmentFields.sfActualIncomeDate).Value
      End Get
    End Property

    Public Property ActualCount() As Integer
      Get
        ActualCount = mvClassFields.Item(SegmentFields.sfActualCount).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SegmentFields.sfActualCount).Value = CStr(Value)
      End Set
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(SegmentFields.sfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SegmentFields.sfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AppealCode() As String
      Get
        AppealCode = mvClassFields.Item(SegmentFields.sfAppeal).Value
      End Get
    End Property

    Public ReadOnly Property BudgetedCount() As String
      Get
        BudgetedCount = mvClassFields.Item(SegmentFields.sfBudgetedCount).Value
      End Get
    End Property

    Public ReadOnly Property CampaignCode() As String
      Get
        CampaignCode = mvClassFields.Item(SegmentFields.sfCampaign).Value
      End Get
    End Property

    Public ReadOnly Property Cost() As Double
      Get
        Cost = mvClassFields.Item(SegmentFields.sfCost).DoubleValue
      End Get
    End Property

    Public ReadOnly Property CostCentres() As Collection
      Get
        Dim vCostCentre As New SegmentCostCentre
        Dim vRecordSet As CDBRecordSet

        If mvCostCentres Is Nothing Then
          mvCostCentres = New Collection
          vCostCentre.Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vCostCentre.GetRecordSetFields(SegmentCostCentre.SegmentCostCentreRecordSetTypes.sccrtAll) & " FROM segment_cost_centres WHERE campaign = '" & mvClassFields.Item(SegmentFields.sfCampaign).Value & "' AND appeal = '" & mvClassFields.Item(SegmentFields.sfAppeal).Value & "' AND segment = '" & mvClassFields.Item(SegmentFields.sfSegment).Value & "'")
          With vRecordSet
            While .Fetch() = True
              vCostCentre = New SegmentCostCentre
              vCostCentre.InitFromRecordSet(mvEnv, vRecordSet, SegmentCostCentre.SegmentCostCentreRecordSetTypes.sccrtAll)
              mvCostCentres.Add(vCostCentre, vCostCentre.CostCentre)
            End While
            .CloseRecordSet()
          End With
        End If
        CostCentres = mvCostCentres
      End Get
    End Property

    Public ReadOnly Property CountTime() As String
      Get
        CountTime = mvClassFields.Item(SegmentFields.sfCountTime).Value
      End Get
    End Property
    Public ReadOnly Property FixedCost() As Double
      Get
        FixedCost = mvClassFields.Item(SegmentFields.sfFixedCost).DoubleValue
      End Get
    End Property

    Public Property Mailing() As String
      Get
        Mailing = mvClassFields.Item(SegmentFields.sfMailing).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SegmentFields.sfMailing).Value = Value
      End Set
    End Property
    Public Property MailingNumber() As Integer
      Get
        MailingNumber = mvClassFields.Item(SegmentFields.sfMailingNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SegmentFields.sfMailingNumber).Value = CStr(Value)
      End Set
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(SegmentFields.sfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property OutputGroup() As String
      Get
        OutputGroup = mvClassFields.Item(SegmentFields.sfOutputGroup).Value
      End Get
    End Property

    Public ReadOnly Property ProductAllocations() As Collection
      Get
        Dim vProductAllocation As New SegmentProductAllocation
        Dim vRecordSet As CDBRecordSet

        If mvProducts Is Nothing Then
          mvProducts = New Collection
          vProductAllocation.Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vProductAllocation.GetRecordSetFields(SegmentProductAllocation.SegmentProductAllocationRecordSetTypes.spartAll) & " FROM segment_product_allocation WHERE campaign = '" & mvClassFields.Item(SegmentFields.sfCampaign).Value & "' AND appeal = '" & mvClassFields.Item(SegmentFields.sfAppeal).Value & "' AND segment = '" & mvClassFields.Item(SegmentFields.sfSegment).Value & "' ORDER BY amount_number")
          With vRecordSet
            While .Fetch() = True
              vProductAllocation = New SegmentProductAllocation
              vProductAllocation.InitFromRecordSet(mvEnv, vRecordSet, SegmentProductAllocation.SegmentProductAllocationRecordSetTypes.spartAll)
              mvProducts.Add(vProductAllocation, CStr(vProductAllocation.AmountNumber))
            End While
            .CloseRecordSet()
          End With
        End If
        ProductAllocations = mvProducts
      End Get
    End Property

    Public Property Random() As Boolean
      Get
        Random = mvClassFields.Item(SegmentFields.sfRandom).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(SegmentFields.sfRandom).Bool = Value
      End Set
    End Property

    Public ReadOnly Property RequiredCount() As Integer
      Get
        RequiredCount = mvClassFields.Item(SegmentFields.sfRequiredCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Score() As String
      Get
        Score = mvClassFields.Item(SegmentFields.sfScore).Value
      End Get
    End Property

    Public ReadOnly Property SegmentCode() As String
      Get
        SegmentCode = mvClassFields.Item(SegmentFields.sfSegment).Value
      End Get
    End Property

    Public ReadOnly Property SegmentAdColour() As String
      Get
        SegmentAdColour = mvClassFields.Item(SegmentFields.sfSegmentAdColour).Value
      End Get
    End Property

    Public ReadOnly Property SegmentAdImage() As String
      Get
        SegmentAdImage = mvClassFields.Item(SegmentFields.sfSegmentAdImage).Value
      End Get
    End Property

    Public ReadOnly Property SegmentAdPagePosition() As String
      Get
        SegmentAdPagePosition = mvClassFields.Item(SegmentFields.sfSegmentAdPagePosition).Value
      End Get
    End Property

    Public ReadOnly Property SegmentAdPublPosition() As String
      Get
        SegmentAdPublPosition = mvClassFields.Item(SegmentFields.sfSegmentAdPublPosition).Value
      End Get
    End Property

    Public ReadOnly Property SegmentAdSize() As String
      Get
        SegmentAdSize = mvClassFields.Item(SegmentFields.sfSegmentAdSize).Value
      End Get
    End Property

    Public ReadOnly Property SegmentCreative() As String
      Get
        SegmentCreative = mvClassFields.Item(SegmentFields.sfSegmentCreative).Value
      End Get
    End Property

    Public ReadOnly Property SegmentDate() As String
      Get
        SegmentDate = mvClassFields.Item(SegmentFields.sfSegmentDate).Value
      End Get
    End Property

    Public ReadOnly Property SegmentDesc() As String
      Get
        SegmentDesc = mvClassFields.Item(SegmentFields.sfSegmentDesc).Value
      End Get
    End Property

    Public ReadOnly Property SegmentGeographic() As String
      Get
        SegmentGeographic = mvClassFields.Item(SegmentFields.sfSegmentGeographic).Value
      End Get
    End Property

    Public ReadOnly Property SegmentLetterSize() As String
      Get
        SegmentLetterSize = mvClassFields.Item(SegmentFields.sfSegmentLetterSize).Value
      End Get
    End Property

    Public ReadOnly Property SegmentMediaName() As String
      Get
        SegmentMediaName = mvClassFields.Item(SegmentFields.sfSegmentMediaName).Value
      End Get
    End Property

    Public ReadOnly Property SegmentMediaType() As String
      Get
        SegmentMediaType = mvClassFields.Item(SegmentFields.sfSegmentMediaType).Value
      End Get
    End Property

    Public ReadOnly Property SegmentPhoneType() As String
      Get
        SegmentPhoneType = mvClassFields.Item(SegmentFields.sfSegmentPhoneType).Value
      End Get
    End Property

    Public ReadOnly Property SegmentSequence() As Integer
      Get
        SegmentSequence = mvClassFields.Item(SegmentFields.sfSegmentSequence).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SegmentTime() As String
      Get
        SegmentTime = mvClassFields.Item(SegmentFields.sfSegmentTime).Value
      End Get
    End Property

    Public ReadOnly Property SegmentType() As String
      Get
        SegmentType = mvClassFields.Item(SegmentFields.sfSegmentType).Value
      End Get
    End Property

    Public Property SelectionSet() As Integer
      Get
        SelectionSet = mvClassFields.Item(SegmentFields.sfSelectionSet).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SegmentFields.sfSelectionSet).Value = CStr(Value)
      End Set
    End Property
    Public ReadOnly Property SourceCode() As String
      Get
        SourceCode = mvClassFields.Item(SegmentFields.sfSource).Value
      End Get
    End Property

    Public ReadOnly Property TargetIncome() As Double
      Get
        TargetIncome = mvClassFields.Item(SegmentFields.sfTargetIncome).DoubleValue
      End Get
    End Property

    Public ReadOnly Property TargetResponse() As Integer
      Get
        TargetResponse = mvClassFields.Item(SegmentFields.sfTargetResponse).IntegerValue
      End Get
    End Property

    Public Property UpdatedByRp() As Boolean
      Get
        UpdatedByRp = mvClassFields.Item(SegmentFields.sfUpdatedByRp).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(SegmentFields.sfUpdatedByRp).Bool = Value
      End Set
    End Property

    Public ReadOnly Property DespatchQuantity() As Integer
      Get
        DespatchQuantity = mvClassFields.Item(SegmentFields.sfDespatchQuantity).IntegerValue
      End Get
    End Property
    Public ReadOnly Property OrgMailAddrUsage() As String
      Get
        OrgMailAddrUsage = mvClassFields.Item(SegmentFields.sfOrgMailAddrUsage).Value
      End Get
    End Property
    Public ReadOnly Property OrgMailLabelName() As String
      Get
        OrgMailLabelName = mvClassFields.Item(SegmentFields.sfOrgMailLabelName).Value
      End Get
    End Property
    Public ReadOnly Property OrgMailTo() As String
      Get
        OrgMailTo = mvClassFields.Item(SegmentFields.sfOrgMailTo).Value
      End Get
    End Property
    Public ReadOnly Property OrgMailWhere() As String
      Get
        OrgMailWhere = mvClassFields.Item(SegmentFields.sfOrgMailWhere).Value
      End Get
    End Property
    Public ReadOnly Property OrgMailRoles() As String
      Get
        OrgMailRoles = mvClassFields.Item(SegmentFields.sfOrgMailRoles).Value
      End Get
    End Property
    Public ReadOnly Property SelectionOptions() As CDBEnvironment.SelectionOptionSettings
      Get
        Dim vCount As Integer

        If Not mvSelectionOptionsSet Then
          mvSelectionOptionsSet = True
          With mvClassFields
            If Len(.Item(SegmentFields.sfOrgMailTo).Value) > 0 Then vCount = vCount + 1
            If Len(.Item(SegmentFields.sfOrgMailWhere).Value) > 0 Then vCount = vCount + 1
            If Len(.Item(SegmentFields.sfOrgMailAddrUsage).Value) > 0 Then vCount = vCount + 1
            If Len(.Item(SegmentFields.sfOrgMailRoles).Value) > 0 Then vCount = vCount + 1
            If Len(.Item(SegmentFields.sfOrgMailLabelName).Value) > 0 Then vCount = vCount + 1
          End With
          Select Case vCount
            Case 0
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosNone
            Case 5
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosAll
            Case Else
              mvSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosSome
          End Select
        End If
        SelectionOptions = mvSelectionOptions
      End Get
    End Property

    Public ReadOnly Property IncentiveScheme() As String
      Get
        IncentiveScheme = mvSourceValues.svsIncentive
      End Get
    End Property
    Public ReadOnly Property IncentiveTriggerLevel() As String
      Get
        IncentiveTriggerLevel = mvSourceValues.svsTrigger
      End Get
    End Property
    Public ReadOnly Property ThankYouLetter() As String
      Get
        ThankYouLetter = mvSourceValues.svsThankYouLetter
      End Get
    End Property
    Public ReadOnly Property Direction() As String
      Get
        Direction = mvMailingValues.mvsDirection
      End Get
    End Property

    Public ReadOnly Property MailingDescription() As String
      Get
        MailingDescription = mvMailingValues.mvsDescription
      End Get
    End Property

    Public ReadOnly Property SourceDescription() As String
      Get
        SourceDescription = mvSourceValues.svsSourceDesc
      End Get
    End Property

    Public ReadOnly Property TotalItemisedCost() As Double
      Get
        TotalItemisedCost = mvClassFields.Item(SegmentFields.sfTotalItemisedCost).DoubleValue
      End Get
    End Property

    Public ReadOnly Property Telemarketing() As Boolean
      Get
        Return mvClassFields.Item(SegmentFields.sfTelemarketing).Bool
      End Get
    End Property
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "segments"
          .Add("campaign")
          .Add("appeal")
          .Add("segment")
          .Add("segment_desc")
          .Add("segment_type")
          .Add("segment_sequence", CDBField.FieldTypes.cftInteger)
          .Add("criteria_set", CDBField.FieldTypes.cftLong)
          .Add("selection_set", CDBField.FieldTypes.cftLong)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("output_group")
          .Add("required_count", CDBField.FieldTypes.cftLong)
          .Add("actual_count", CDBField.FieldTypes.cftLong)
          .Add("actual_income", CDBField.FieldTypes.cftNumeric)
          .Add("actual_income_date", CDBField.FieldTypes.cftDate)
          .Add("count_time")
          .Add("cost", CDBField.FieldTypes.cftNumeric)
          .Add("fixed_cost", CDBField.FieldTypes.cftNumeric)
          .Add("random")
          .Add("mailing")
          .Add("score")
          .Add("source")
          .Add("amended_by")
          .Add("target_income", CDBField.FieldTypes.cftNumeric)
          .Add("target_response", CDBField.FieldTypes.cftLong)
          .Add("segment_date", CDBField.FieldTypes.cftDate)
          .Add("segment_time")
          .Add("segment_media_type")
          .Add("segment_media_name")
          .Add("segment_geographic")
          .Add("segment_ad_page_position")
          .Add("segment_ad_publ_position")
          .Add("segment_ad_size")
          .Add("segment_ad_image")
          .Add("segment_ad_colour")
          .Add("segment_letter_size")
          .Add("segment_phone_type")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("mailing_number", CDBField.FieldTypes.cftLong)
          .Add("updated_by_rp")
          .Add("despatch_quantity", CDBField.FieldTypes.cftInteger)
          .Add("segment_creative", CDBField.FieldTypes.cftCharacter)
          .Add("budgeted_count", CDBField.FieldTypes.cftCharacter)
          .Add("org_mail_to")
          .Add("org_mail_where")
          .Add("org_mail_roles")
          .Add("org_mail_addr_usage")
          .Add("org_mail_label_name")
          .Add("total_itemised_cost", CDBField.FieldTypes.cftNumeric)
          .Add("telemarketing").InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing)

          .Item(SegmentFields.sfCampaign).SetPrimaryKeyOnly()
          .Item(SegmentFields.sfAppeal).SetPrimaryKeyOnly()
          .Item(SegmentFields.sfSegment).SetPrimaryKeyOnly()
          .SetUniqueField(SegmentFields.sfCampaign)
          .SetUniqueField(SegmentFields.sfAppeal)
          .SetUniqueField(SegmentFields.sfSegment)

          .Item(SegmentFields.sfBudgetedCount).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignBudgets)
          .Item(SegmentFields.sfActualIncome).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignActualIncome)
          .Item(SegmentFields.sfActualIncomeDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignActualIncome)
          .Item(SegmentFields.sfTotalItemisedCost).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignItemisedCosts)
          If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataSegmentOrgSelectionOptions) Then
            .Item(SegmentFields.sfOrgMailTo).InDatabase = False
            .Item(SegmentFields.sfOrgMailWhere).InDatabase = False
            .Item(SegmentFields.sfOrgMailRoles).InDatabase = False
            .Item(SegmentFields.sfOrgMailAddrUsage).InDatabase = False
            .Item(SegmentFields.sfOrgMailLabelName).InDatabase = False
          End If

          If mvPrefixRequired Then
            .Item(SegmentFields.sfSource).PrefixRequired = True
            .Item(SegmentFields.sfMailing).PrefixRequired = True
            .Item(SegmentFields.sfNotes).PrefixRequired = True
            .Item(SegmentFields.sfAmendedBy).PrefixRequired = True
            .Item(SegmentFields.sfAmendedOn).PrefixRequired = True
            .Item(SegmentFields.sfMailingNumber).PrefixRequired = True
          End If
        End With
      Else
        mvClassFields.ClearItems()
      End If
      'UPGRADE_NOTE: Object mvProducts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvProducts = Nothing
      'UPGRADE_NOTE: Object mvCostCentres may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvCostCentres = Nothing
      ClearSourceValues()
      mvMailingValuesSet = False
      mvAmendedValid = False
      mvExisting = False
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As SegmentFields)
      'Add code here to ensure all values are valid before saving
      If pField = SegmentFields.sfAll And Not mvAmendedValid Then
        mvClassFields.Item(SegmentFields.sfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(SegmentFields.sfAmendedBy).Value = mvEnv.User.Logname
      End If
      If mvClassFields(SegmentFields.sfUpdatedByRp).Value = "" Then mvClassFields(SegmentFields.sfUpdatedByRp).Value = "N"
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SegmentRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SegmentRecordSetTypes.srtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sg")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCampaign As String = "", Optional ByRef pAppeal As String = "", Optional ByRef pSegment As String = "", Optional ByRef pGetSourceAndMailingData As Boolean = False)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      mvPrefixRequired = pGetSourceAndMailingData
      If Len(pCampaign) > 0 And Len(pAppeal) > 0 And Len(pSegment) > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(SegmentRecordSetTypes.srtAll)
        If pGetSourceAndMailingData Then vSQL = vSQL & ",source_desc,incentive_trigger_level,thank_you_letter,incentive_scheme,distribution_code,discount_percentage,mailing_desc,direction,m.notes AS m_notes,department"
        vSQL = vSQL & " FROM segments sg"
        If pGetSourceAndMailingData Then vSQL = vSQL & ", sources sc, mailings m"
        vSQL = vSQL & " WHERE campaign = '" & pCampaign & "' AND appeal = '" & pAppeal & "' AND segment = '" & pSegment & "'"
        If pGetSourceAndMailingData Then vSQL = vSQL & " AND sg.source = sc.source AND sg.mailing = m.mailing"
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SegmentRecordSetTypes.srtAll, pGetSourceAndMailingData)
        Else
          InitClassFields()
          SetDefaults()
          mvClassFields.Item(SegmentFields.sfCampaign).Value = pCampaign
          mvClassFields.Item(SegmentFields.sfAppeal).Value = pAppeal
          mvClassFields.Item(SegmentFields.sfSegment).Value = pSegment
          If pSegment = "Exclud" Then 'initalising standard exclusions dummy segment
            mvClassFields.Item(SegmentFields.sfSegmentDesc).Value = "Standard Exclusions"
          ElseIf pSegment = "SMSeg" Then  'initalising dummy segment one for Selection Manager mailings
            mvClassFields.Item(SegmentFields.sfSegmentDesc).Value = "Current Criteria"
            mvClassFields.Item(SegmentFields.sfSegmentSequence).Value = "1"
            mvClassFields.Item(SegmentFields.sfOutputGroup).Value = "SMOutputGroup"
            mvCreateMailingHistory = True
            mvCreateOutput = True
          End If
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromMailing(ByVal pEnv As CDBEnvironment, ByVal pMailing As String)
      mvEnv = pEnv
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields(SegmentRecordSetTypes.srtAll), mvClassFields.DatabaseTableName + " sg", New CDBField("Mailing", pMailing))
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet()
      If vRecordSet.Fetch() Then
        InitFromRecordSet(pEnv, vRecordSet, SegmentRecordSetTypes.srtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SegmentRecordSetTypes, Optional ByRef pGetSourceAndMailingData As Boolean = False)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SegmentFields.sfCampaign, vFields)
        .SetItem(SegmentFields.sfAppeal, vFields)
        .SetItem(SegmentFields.sfSegment, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SegmentRecordSetTypes.srtAll) = SegmentRecordSetTypes.srtAll Then
          .SetItem(SegmentFields.sfSegmentDesc, vFields)
          .SetItem(SegmentFields.sfSegmentType, vFields)
          .SetItem(SegmentFields.sfSegmentSequence, vFields)
          .SetItem(SegmentFields.sfCriteriaSet, vFields)
          .SetItem(SegmentFields.sfSelectionSet, vFields)
          .SetItem(SegmentFields.sfNotes, vFields)
          .SetItem(SegmentFields.sfOutputGroup, vFields)
          .SetItem(SegmentFields.sfRequiredCount, vFields)
          .SetItem(SegmentFields.sfActualCount, vFields)
          .SetOptionalItem(SegmentFields.sfActualIncome, vFields)
          .SetOptionalItem(SegmentFields.sfActualIncomeDate, vFields)
          .SetItem(SegmentFields.sfCountTime, vFields)
          .SetItem(SegmentFields.sfCost, vFields)
          .SetItem(SegmentFields.sfFixedCost, vFields)
          .SetItem(SegmentFields.sfRandom, vFields)
          .SetItem(SegmentFields.sfMailing, vFields)
          .SetItem(SegmentFields.sfScore, vFields)
          .SetItem(SegmentFields.sfSource, vFields)
          .SetItem(SegmentFields.sfAmendedBy, vFields)
          .SetItem(SegmentFields.sfTargetIncome, vFields)
          .SetItem(SegmentFields.sfTargetResponse, vFields)
          .SetItem(SegmentFields.sfSegmentDate, vFields)
          .SetItem(SegmentFields.sfSegmentTime, vFields)
          .SetItem(SegmentFields.sfSegmentMediaType, vFields)
          .SetItem(SegmentFields.sfSegmentMediaName, vFields)
          .SetItem(SegmentFields.sfSegmentGeographic, vFields)
          .SetItem(SegmentFields.sfSegmentAdPagePosition, vFields)
          .SetItem(SegmentFields.sfSegmentAdPublPosition, vFields)
          .SetItem(SegmentFields.sfSegmentAdSize, vFields)
          .SetItem(SegmentFields.sfSegmentAdImage, vFields)
          .SetItem(SegmentFields.sfSegmentAdColour, vFields)
          .SetItem(SegmentFields.sfSegmentLetterSize, vFields)
          .SetItem(SegmentFields.sfSegmentPhoneType, vFields)
          .SetItem(SegmentFields.sfAmendedOn, vFields)
          .SetItem(SegmentFields.sfMailingNumber, vFields)
          .SetItem(SegmentFields.sfUpdatedByRp, vFields)
          .SetOptionalItem(SegmentFields.sfDespatchQuantity, vFields)
          .SetOptionalItem(SegmentFields.sfSegmentCreative, vFields)
          .SetOptionalItem(SegmentFields.sfBudgetedCount, vFields)
          .SetOptionalItem(SegmentFields.sfOrgMailTo, vFields)
          .SetOptionalItem(SegmentFields.sfOrgMailWhere, vFields)
          .SetOptionalItem(SegmentFields.sfOrgMailRoles, vFields)
          .SetOptionalItem(SegmentFields.sfOrgMailAddrUsage, vFields)
          .SetOptionalItem(SegmentFields.sfOrgMailLabelName, vFields)
          .SetOptionalItem(SegmentFields.sfTotalItemisedCost, vFields)
          .SetOptionalItem(SegmentFields.sfTelemarketing, vFields)
        End If
        If .Item(SegmentFields.sfSelectionSet).Value = "" Then .Item(SegmentFields.sfSelectionSet).Value = CStr(mvEnv.GetControlNumber("SS"))
        If pGetSourceAndMailingData Then
          GetSourceValues(pRecordSet)
          GetMailingValues(pRecordSet)
        Else
          GetSourceValues()
          GetMailingValues()
        End If
      End With
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pParams As CDBParameters)
      'Only used by Web Services

      Init(pEnv)
      With mvClassFields
        .Item(SegmentFields.sfCampaign).Value = pParams("Campaign").Value
        .Item(SegmentFields.sfAppeal).Value = pParams("Appeal").Value
        .Item(SegmentFields.sfSegment).Value = pParams("Segment").Value
      End With
      Update(pParams)
      Dim vAppeal As New Appeal(pEnv)
      vAppeal.Init(pParams)
      If vAppeal.AppealCostCentre.Length > 0 Then
        Dim vCostCentre As New SegmentCostCentre()
        vCostCentre.Init(pEnv)
        vCostCentre.Create(pParams("Campaign").Value, pParams("Appeal").Value, pParams("Segment").Value, vAppeal.AppealCostCentre, "100")
        vCostCentre.Save()
      End If

    End Sub

    Public Sub Update(ByVal pParams As CDBParameters)
      'Used by Web Services only

      With mvClassFields
        If pParams.Exists("SegmentDesc") Then .Item(SegmentFields.sfSegmentDesc).Value = pParams("SegmentDesc").Value
        If pParams.Exists("SegmentType") Then .Item(SegmentFields.sfSegmentType).Value = pParams("SegmentType").Value
        If pParams.Exists("SegmentSequence") Then .Item(SegmentFields.sfSegmentSequence).Value = pParams("SegmentSequence").Value
        If pParams.HasValue("CriteriaSet") Then .Item(SegmentFields.sfCriteriaSet).Value = pParams("CriteriaSet").Value
        '  sfSelectionSet
        If pParams.Exists("Notes") Then .Item(SegmentFields.sfNotes).Value = pParams("Notes").Value
        If pParams.Exists("OutputGroup") Then .Item(SegmentFields.sfOutputGroup).Value = pParams("OutputGroup").Value
        If pParams.Exists("RequiredCount") Then .Item(SegmentFields.sfRequiredCount).Value = pParams("RequiredCount").Value
        If pParams.Exists("ActualCount") Then .Item(SegmentFields.sfActualCount).Value = pParams("ActualCount").Value
        '***If pParams.Exists("ActualIncome") Then .Item(sfActualIncome).Value = pParams("ActualIncome").Value
        '***If pParams.Exists("ActualIncomeDate") Then .Item(sfActualIncomeDate).Value = pParams("ActualincomeDate").Value
        '  sfCountTime
        If pParams.Exists("Cost") Then .Item(SegmentFields.sfCost).Value = pParams("Cost").Value
        If pParams.Exists("FixedCost") Then .Item(SegmentFields.sfFixedCost).Value = pParams("FixedCost").Value
        If pParams.Exists("Random") Then .Item(SegmentFields.sfRandom).Value = pParams("Random").Value
        If pParams.Exists("Mailing") Then .Item(SegmentFields.sfMailing).Value = pParams("Mailing").Value
        If pParams.Exists("Score") Then .Item(SegmentFields.sfScore).Value = pParams("Score").Value
        If pParams.Exists("Source") Then .Item(SegmentFields.sfSource).Value = pParams("Source").Value
        If pParams.Exists("TargetIncome") Then .Item(SegmentFields.sfTargetIncome).Value = pParams("TargetIncome").Value
        If pParams.Exists("TargetResponse") Then .Item(SegmentFields.sfTargetResponse).Value = pParams("TargetResponse").Value
        If pParams.Exists("SegmentDate") Then .Item(SegmentFields.sfSegmentDate).Value = pParams("SegmentDate").Value
        If pParams.Exists("SegmentTime") Then .Item(SegmentFields.sfSegmentTime).Value = pParams("SegmentTime").Value
        If pParams.Exists("SegmentMediaType") Then .Item(SegmentFields.sfSegmentMediaType).Value = pParams("SegmentMediaType").Value
        If pParams.Exists("SegmentMediaName") Then .Item(SegmentFields.sfSegmentMediaName).Value = pParams("SegmentMediaName").Value
        If pParams.Exists("SegmentGeographic") Then .Item(SegmentFields.sfSegmentGeographic).Value = pParams("SegmentGeographic").Value
        If pParams.Exists("SegmentAdPagePosition") Then .Item(SegmentFields.sfSegmentAdPagePosition).Value = pParams("SegmentAdPagePosition").Value
        If pParams.Exists("SegmentAdPublPosition") Then .Item(SegmentFields.sfSegmentAdPublPosition).Value = pParams("SegmentAdPublPosition").Value
        If pParams.Exists("SegmentAdSize") Then .Item(SegmentFields.sfSegmentAdSize).Value = pParams("SegmentAdSize").Value
        If pParams.Exists("SegmentAdImage") Then .Item(SegmentFields.sfSegmentAdImage).Value = pParams("SegmentAdImage").Value
        If pParams.Exists("SegmentAdColour") Then .Item(SegmentFields.sfSegmentAdColour).Value = pParams("SegmentAdColour").Value
        If pParams.Exists("SegmentLetterSize") Then .Item(SegmentFields.sfSegmentLetterSize).Value = pParams("SegmentLetterSize").Value
        If pParams.Exists("SegmentPhoneType") Then .Item(SegmentFields.sfSegmentPhoneType).Value = pParams("SegmentPhoneType").Value
        '  sfMailingNumber
        '  sfUpdatedByRp
        If pParams.Exists("DespatchQuantity") Then .Item(SegmentFields.sfDespatchQuantity).Value = pParams("DespatchQuantity").Value
        If pParams.Exists("SegmentCreative") Then .Item(SegmentFields.sfSegmentCreative).Value = pParams("SegmentCreative").Value
        If pParams.Exists("BudgetedCount") Then .Item(SegmentFields.sfBudgetedCount).Value = pParams("BudgetedCount").Value
        If pParams.Exists("Telemarketing") Then .Item(SegmentFields.sfTelemarketing).Value = pParams("Telemarketing").Value
        '  sfOrgMailTo
        '  sfOrgMailWhere
        '  sfOrgMailRoles
        '  sfOrgMailAddrUsage
        '  sfOrgMailLabelName
      End With

    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vTransaction As Boolean

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      DeleteAssociatedData(CampaignCode, AppealCode, SegmentCode)
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub DeleteAssociatedData(ByVal pCampaign As String, Optional ByVal pAppeal As String = "", Optional ByVal pSegment As String = "")
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vTransaction As Boolean
      Dim vAppeal As New Appeal(mvEnv)
      Dim vAppealBudget As New AppealBudget
      Dim vABD As New AppealBudgetDetail

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If
      With vWhereFields
        .Add("campaign", CDBField.FieldTypes.cftCharacter, pCampaign)
        If Len(pAppeal) > 0 Then .Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal)
        If Len(pSegment) > 0 Then .Add("segment", CDBField.FieldTypes.cftCharacter, pSegment)
      End With
      mvEnv.Connection.DeleteRecords("segment_cost_centres", vWhereFields, False)
      mvEnv.Connection.DeleteRecords("segment_product_allocation", vWhereFields, False)
      mvEnv.Connection.DeleteRecords("tick_boxes", vWhereFields, False)
      If CriteriaSet > 0 Then
        With vWhereFields
          .Clear()
          .Add("criteria_set", CDBField.FieldTypes.cftLong, CriteriaSet)
        End With
        mvEnv.Connection.DeleteRecords("criteria_set_details", vWhereFields, False)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.CDBDataViewNames) Then
          mvEnv.Connection.DeleteRecords("criteria_sets", vWhereFields, False)
          mvEnv.Connection.DeleteRecords("selection_steps", vWhereFields, False)
        End If
      End If
      If Len(SourceCode) > 0 Then
        If SourceCode = DeriveSegmentSource(pCampaign, pAppeal, pSegment) Then
          vFields.Add("history_only", CDBField.FieldTypes.cftCharacter, "Y")
          vFields.AddAmendedOnBy(mvEnv.User.Logname)
          vWhereFields.Clear()
          vWhereFields.Add("source", CDBField.FieldTypes.cftCharacter, SourceCode)
          mvEnv.Connection.UpdateRecords("sources", vFields, vWhereFields, False)
        End If
      End If
      If Len(Mailing) > 0 Then
        If Mailing = DeriveSegmentSource(pCampaign, pAppeal, pSegment) Then
          vFields.Clear()
          vFields.Add("history_only", CDBField.FieldTypes.cftCharacter, "Y")
          vFields.AddAmendedOnBy(mvEnv.User.Logname)
          vWhereFields.Clear()
          vWhereFields.Add("mailing", CDBField.FieldTypes.cftCharacter, Mailing)
          mvEnv.Connection.UpdateRecords("mailings", vFields, vWhereFields, False)
        End If
      End If

      If Len(pAppeal) > 0 Then
        vAppeal.Init(pCampaign, pAppeal, False, Appeal.SegmentSortOrders.ssoSequence, True)
        For Each vAppealBudget In vAppeal.Budgets
          For Each vABD In vAppealBudget.AppealBudgetDetails
            If Len(vABD.Segment) > 0 And vABD.Segment = pSegment Then
              vABD.Delete()
            End If
          Next vABD
        Next vAppealBudget
      End If
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Friend Sub Clone(ByVal pCampaignCode As String, ByVal pAppealCode As String, ByVal pSegmentCode As String, ByVal pSegmentDesc As String, ByVal pSegmentSequence As Integer, ByVal pMailingCode As String, ByVal pSourceCode As String)
      'Used by Web Services only

      With mvClassFields
        .Item(SegmentFields.sfCampaign).Value = pCampaignCode
        .Item(SegmentFields.sfAppeal).Value = pAppealCode
        .Item(SegmentFields.sfSegment).Value = pSegmentCode
        .Item(SegmentFields.sfSegmentDesc).Value = pSegmentDesc
        .Item(SegmentFields.sfSegmentSequence).Value = CStr(pSegmentSequence)
        .Item(SegmentFields.sfMailing).Value = pMailingCode
        .Item(SegmentFields.sfSource).Value = pSourceCode
        .Item(SegmentFields.sfSelectionSet).Value = ""
        .Item(SegmentFields.sfMailingNumber).Value = "" 'BR14739: Don't save the mailing number as the segment has not yet been mailed
        If RequiredCount > 0 Then
          'Leave as it is
        Else
          .Item(SegmentFields.sfRequiredCount).Value = "" 'Ensure it is null rather than zero
        End If
        .ClearSetValues()
      End With
      mvExisting = False
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields
      Dim vFields As New CDBFields
      Dim vMailingWhere As New CDBFields
      Dim vMailingCount As Integer
      Dim vSourceWhere As New CDBFields
      Dim vSourceCount As Integer
      Dim vTransaction As Boolean

      If mvMailingValuesSet Then
        vMailingWhere.Add("mailing", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsMailing)
        vMailingCount = mvEnv.Connection.GetCount("mailings", vMailingWhere)
      End If
      If mvSourceValuesSet Then
        vSourceWhere.Add("source", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsSource)
        vSourceCount = mvEnv.Connection.GetCount("sources", vSourceWhere)
      End If

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      SetValid(SegmentFields.sfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      With vWhereFields
        .Add("campaign", CDBField.FieldTypes.cftCharacter, CampaignCode)
        .Add("appeal", CDBField.FieldTypes.cftCharacter, AppealCode)
        .Add("segment", CDBField.FieldTypes.cftCharacter, SegmentCode)
      End With

      If mvMailingValuesSet Then
        With vFields
          .Clear()
          .Add("direction", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsDirection)
          .Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailingNotes) Then .Add("notes", CDBField.FieldTypes.cftMemo, mvMailingValues.mvsNotes)
          .AddAmendedOnBy(mvEnv.User.Logname)
          If vMailingCount > 0 Then
            If Len(mvMailingValues.mvsDescription) > 0 Then .Add("mailing_desc", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsDescription)
            mvEnv.Connection.UpdateRecords("mailings", vFields, vMailingWhere)
          Else
            .Add("mailing_desc", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsDescription)
            .Add("marketing", CDBField.FieldTypes.cftCharacter, "Y")
            .Add("department", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsDepartment)
            .Add("mailing", CDBField.FieldTypes.cftCharacter, mvMailingValues.mvsMailing)
            mvEnv.Connection.InsertRecord("mailings", vFields)
          End If
        End With
      End If

      If mvSourceValuesSet Then
        With vFields
          .Clear()
          .Add("incentive_scheme", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsIncentive)
          .Add("incentive_trigger_level", CDBField.FieldTypes.cftNumeric, mvSourceValues.svsTrigger)
          .Add("thank_you_letter", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsThankYouLetter)
          .AddAmendedOnBy(mvEnv.User.Logname)
          .Add("history_only", CDBField.FieldTypes.cftCharacter, "N")
          .Add("distribution_code", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsDistributionCode)
          .Add("discount_percentage", CDBField.FieldTypes.cftNumeric, mvSourceValues.svsDiscountPercentage)
          If vSourceCount > 0 Then
            If Len(mvSourceValues.svsSourceDesc) > 0 Then .Add("source_desc", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsSourceDesc)
            mvEnv.Connection.UpdateRecords("sources", vFields, vSourceWhere, False)
          Else
            .Add("source", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsSource)
            .Add("source_desc", CDBField.FieldTypes.cftCharacter, mvSourceValues.svsSourceDesc)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then .Add("source_number", CDBField.FieldTypes.cftLong, mvEnv.GetControlNumber("SR"))
            mvEnv.Connection.InsertRecord("sources", vFields)
          End If
        End With
      End If
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SetSourceValues(ByVal pSource As String, ByVal pSourceDesc As String, ByVal pIncentive As String, ByVal pTriggerLevel As Double, ByVal pThankYouLetter As String, ByVal pDistributionCode As String, ByVal pSourceDiscountPercentage As String)
      On Error GoTo SetSourceValuesError
      With mvSourceValues
        If Len(pSource) = 0 Then
          pSource = DeriveSegmentSource()
          mvClassFields(SegmentFields.sfSource).Value = pSource
          If Len(pSourceDesc) = 0 Then pSourceDesc = SegmentDesc
        Else
          mvClassFields(SegmentFields.sfSource).Value = pSource
        End If
        .svsSource = pSource
        .svsSourceDesc = pSourceDesc
        .svsIncentive = pIncentive
        .svsTrigger = CStr(pTriggerLevel)
        .svsThankYouLetter = pThankYouLetter
        .svsDistributionCode = pDistributionCode
        .svsDiscountPercentage = pSourceDiscountPercentage
      End With
      mvSourceValuesSet = True
      Exit Sub

SetSourceValuesError:
      mvSourceValuesSet = False
    End Sub

    Public Sub SetMailingValues(ByVal pMailing As String, ByVal pDescription As String, ByVal pDirection As String, ByVal pNotes As String, ByVal pDepartment As String, ByVal pLogname As String)
      On Error GoTo SetMailingValuesError

      With mvMailingValues
        If Len(pMailing) = 0 Then
          pMailing = DeriveSegmentSource()
          mvClassFields(SegmentFields.sfMailing).Value = pMailing
          If Len(pDescription) = 0 Then pDescription = SegmentDesc
          If Len(pNotes) = 0 Then pNotes = Notes
          If Len(pDepartment) = 0 Then pDepartment = mvEnv.User.Department
          If Len(pLogname) = 0 Then pLogname = mvEnv.User.Logname
        End If
        .mvsMailing = pMailing
        .mvsDescription = pDescription
        .mvsDirection = pDirection
        .mvsNotes = pNotes
        .mvsDepartment = pDepartment
        .mvsLogname = pLogname
      End With
      mvMailingValuesSet = True
      Exit Sub

SetMailingValuesError:
      mvMailingValuesSet = False
    End Sub

    Public Sub InitMailingData(ByRef pCreateOutput As Boolean, ByRef pMailingHistory As Boolean, ByRef pSelectionSetDesc As String)
      mvCreateOutput = pCreateOutput
      mvCreateMailingHistory = pMailingHistory
      If pSelectionSetDesc.Length > 0 Then
        mvCreateSelectionSet = True
        mvClassFields(SegmentFields.sfSegmentDesc).Value = pSelectionSetDesc
      Else
        mvCreateSelectionSet = False
      End If
    End Sub

    Public Sub InitOutputGroup(ByVal pFields As CDBFields)
      mvReportCode = pFields("report_code").Value
      mvCreateMailingHistory = pFields("create_mailing_history").Bool
      mvCreateSelectionSet = pFields("create_selection_set").Bool
      mvCreateOutput = pFields("create_output").Bool
    End Sub

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(SegmentFields.sfAmendedOn).Value = pAmendedOn
      mvClassFields.Item(SegmentFields.sfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Function DeriveSegmentSource(Optional ByVal pCampaign As String = "", Optional ByVal pAppeal As String = "", Optional ByVal pSegment As String = "") As String
      Dim vSource As String
      Dim vSum As Integer
      Dim vCounter As Integer
      Dim vTimesThru As Integer
      Dim vChar As String
      Dim vValue As Integer
      Dim vRemainder As Integer

      If Len(pCampaign) = 0 Then pCampaign = CampaignCode
      If Len(pAppeal) = 0 Then pAppeal = AppealCode
      If Len(pSegment) = 0 Then pSegment = SegmentCode

      If Len(pCampaign) = 3 And Len(pAppeal) = 3 And Len(pSegment) = 3 Then
        vSource = pCampaign & pAppeal & pSegment
        If mvEnv.GetConfigOption("ma_check_digit_on_source", True) Then
          vSum = 0
          vCounter = 1
          vTimesThru = 1
          While vCounter <= Len(vSource)
            vChar = Mid(vSource, vCounter, 1)
            If vChar Like "[A-Z]" Then
              'Convert chars to numbers A=10, B=11, etc.  (Asc("A") = 65)
              vValue = Asc(vChar) - 55
            Else
              vValue = IntegerValue(vChar)
            End If
            Select Case vTimesThru
              Case 1
                vSum = vSum + (vValue * 1)
              Case 2
                vSum = vSum + (vValue * 3)
              Case 3
                vSum = vSum + (vValue * 5)
              Case 4
                vSum = vSum + (vValue * 7)
                vTimesThru = 0
            End Select
            vCounter = vCounter + 1
            vTimesThru = vTimesThru + 1
          End While
          vRemainder = vSum Mod 10
          vSource = vSource & vRemainder.ToString("0")
        End If
      Else
        vSource = ""
      End If
      DeriveSegmentSource = vSource
    End Function

    Public Function DataTable() As CDBDataTable
      Dim vDataTable As CDBDataTable
      Dim vRow As CDBDataRow = Nothing

      vDataTable = mvClassFields.DataTable
      vDataTable.AddColumnsFromList("SourceDesc,IncentiveTriggerLevel,ThankYouLetter,IncentiveScheme,DistributionCode,DiscountPercentage,Direction")
      If Len(SourceCode) > 0 Then
        GetSourceValues()
        vRow = vDataTable.Rows.Item(0)
        With vRow
          .Item("SourceDesc") = mvSourceValues.svsSourceDesc
          .Item("IncentiveTriggerLevel") = mvSourceValues.svsTrigger
          .Item("ThankYouLetter") = mvSourceValues.svsThankYouLetter
          .Item("IncentiveScheme") = mvSourceValues.svsIncentive
          .Item("DistributionCode") = mvSourceValues.svsDistributionCode
          .Item("DiscountPercentage") = mvSourceValues.svsDiscountPercentage
        End With
      End If
      If Mailing.Length > 0 Then
        GetMailingValues()
        If vRow IsNot Nothing Then vRow.Item("Direction") = mvMailingValues.mvsDirection
      End If
      Return vDataTable
    End Function

    Private Sub GetSourceValues(Optional ByRef pRS As CDBRecordSet = Nothing)
      Dim vRecordSet As CDBRecordSet
      Dim vContinue As Boolean
      Dim vCloseRS As Boolean

      With mvSourceValues
        ClearSourceValues()
        If Len(SourceCode) > 0 Then
          If pRS Is Nothing Then
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT source_desc,incentive_trigger_level,thank_you_letter,incentive_scheme,distribution_code,discount_percentage FROM sources WHERE source = '" & SourceCode & "'")
            vContinue = vRecordSet.Fetch() = True
            vCloseRS = vContinue
          Else
            vRecordSet = pRS
            vContinue = True
          End If
          If vContinue Then
            .svsSource = SourceCode
            .svsSourceDesc = vRecordSet.Fields("source_desc").Value
            .svsTrigger = vRecordSet.Fields("incentive_trigger_level").Value
            .svsThankYouLetter = vRecordSet.Fields("thank_you_letter").Value
            .svsIncentive = vRecordSet.Fields("incentive_scheme").Value
            .svsDistributionCode = vRecordSet.Fields("distribution_code").Value
            .svsDiscountPercentage = vRecordSet.Fields("discount_percentage").Value
            If vCloseRS Then vRecordSet.CloseRecordSet()
          End If
        End If
      End With
    End Sub
    Private Sub GetMailingValues(Optional ByRef pRS As CDBRecordSet = Nothing)
      Dim vRecordSet As CDBRecordSet
      Dim vContinue As Boolean
      Dim vCloseRS As Boolean

      With mvMailingValues
        ClearMailingValues()
        If Len(Mailing) > 0 Then
          If pRS Is Nothing Then
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT mailing_desc,direction,notes AS m_notes,department FROM mailings WHERE mailing = '" & Mailing & "'")
            vContinue = vRecordSet.Fetch() = True
            vCloseRS = vContinue
          Else
            vRecordSet = pRS
            vContinue = True
          End If
          If vContinue Then
            .mvsMailing = Mailing
            .mvsDescription = vRecordSet.Fields.Item("mailing_desc").Value
            .mvsDirection = vRecordSet.Fields.Item("direction").Value
            .mvsNotes = vRecordSet.Fields.Item("m_notes").Value
            .mvsDepartment = vRecordSet.Fields.Item("department").Value
            If vCloseRS Then vRecordSet.CloseRecordSet()
          End If
        End If
      End With
    End Sub
    Private Sub ClearSourceValues()
      With mvSourceValues
        .svsSource = ""
        .svsSourceDesc = ""
        .svsTrigger = ""
        .svsThankYouLetter = ""
        .svsIncentive = ""
        .svsDistributionCode = ""
        .svsDiscountPercentage = ""
      End With
      mvSourceValuesSet = False
    End Sub
    Private Sub ClearMailingValues()
      With mvMailingValues
        .mvsMailing = ""
        .mvsDescription = ""
        .mvsDirection = ""
        .mvsNotes = ""
        .mvsDepartment = ""
        .mvsLogname = ""
      End With
      mvMailingValuesSet = False
    End Sub

    Public Sub SetOrgMailOptions(ByVal pOrgMailTo As String, ByVal pOrgMailWhere As String, ByVal pOrgMailAddrUsage As String, ByVal pOrgMailRoles As String, ByVal pOrgMailLabelName As String)
      With mvClassFields
        .Item(SegmentFields.sfOrgMailTo).Value = pOrgMailTo
        .Item(SegmentFields.sfOrgMailWhere).Value = pOrgMailWhere
        .Item(SegmentFields.sfOrgMailAddrUsage).Value = pOrgMailAddrUsage
        .Item(SegmentFields.sfOrgMailRoles).Value = pOrgMailRoles
        .Item(SegmentFields.sfOrgMailLabelName).Value = pOrgMailLabelName
      End With
    End Sub

    Public Sub SetReportCode(ByVal pReportCode As String)
      mvReportCode = pReportCode
    End Sub
  End Class
End Namespace
