Public Class frmCampaignSet
  Implements IPanelVisibility
  Implements IMainForm

  Private mvCampaignDataType As CareServices.XMLCampaignDataSelectionTypes
  Private mvEditing As Boolean
  Private mvCampaignItem As CampaignItem
  Private mvCampaignBusinessType As String
  Private mvMailingType As String
  Private mvMainMenu As MainMenu
  Private WithEvents mvCampaignMenu As CampaignMenu
  Private WithEvents mvFrmEditCriteria As frmEditCriteria
  Private mvMailingInfo As MailingInfo
  Private mvCampaignMarkHistorical As Boolean
  Private mvKeyValuesColl As DisplayGridKeyValues
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private mvActionNumber As Integer
  Private WithEvents mvActionMenu As ActionMenu

#Region "IPanelVisibility"

  Public Property PanelHasFocus() As Boolean Implements CDBNETCL.IPanelVisibility.PanelHasFocus
    Get
      Return sel.Focused
    End Get
    Set(ByVal value As Boolean)
      If value Then sel.Focus()
    End Set
  End Property

  Public Sub SetPanelVisibility() Implements CDBNETCL.IPanelVisibility.SetPanelVisibility
    'Nothing
  End Sub

#End Region

#Region "IMainForm"

  Public ReadOnly Property MainMenu() As MainMenu Implements IMainForm.MainMenu
    Get
      Return mvMainMenu
    End Get
  End Property

#End Region

  Public Sub New(ByVal pCampaignItem As CampaignItem, ByVal pParentForm As MaintenanceParentForm, ByVal pRestrictions As ParameterList)
    MyBase.New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pCampaignItem, pParentForm, pRestrictions)
  End Sub

  Private Sub InitialiseControls(ByVal pCampaignItem As CampaignItem, ByVal pParentForm As MaintenanceParentForm, ByVal pRestrictions As ParameterList)
    mvMainMenu = MainHelper.AddMainMenu(Me)
    mvCampaignItem = pCampaignItem
    mvCampaignMenu = New CampaignMenu(Nothing)
    mvCustomiseMenu = New CustomiseMenu
    mvParentForm = pParentForm
    mvSelectedRow = -1
    splMaint.Panel1Collapsed = True
    cmdSave.Enabled = False       'For the moment
    cmdDelete.Visible = False

    sel.SetCampaignRestrictions(pRestrictions)
    sel.Init(pCampaignItem)

    Me.Text = sel.Caption
    'sel.TreeContextMenu = mvCampaignMenu
    splTop.Panel1Collapsed = True         'No Header
    cmdNew.Visible = False
    cmdClose.Visible = False
    cmdDefault.Visible = False
    HideGrid()
    cmdOther.Visible = False
  End Sub

  Private Sub CampaignTabSelected(ByVal pSender As Object, ByVal pType As CareServices.XMLCampaignDataSelectionTypes, ByVal pCampaignItem As CampaignItem) Handles sel.CampaignTabSelected
    Dim vBusyCursor As New BusyCursor()

    Try
      If pCampaignItem.Appeal <> "" AndAlso pCampaignItem.AppealLocked Then    'BR11765: Dont display any data if the appeal is locked.
        epl.Visible = False
        dgr.Visible = False
        mvCampaignItem = pCampaignItem
        SetButtons(False, True)
        ShowInformationMessage(InformationMessages.ImAppealLocked)
      Else
        mvCampaignDataType = pType
        mvCampaignItem = pCampaignItem
        mvSelectedRow = -1
        RefreshCard()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Public Overrides Sub RefreshData()
    Dim vCancel As Boolean
    BeforeSelect(sel, vCancel)
    If Not vCancel Then RefreshCard()
  End Sub

  Private Sub RefreshCard()
    Dim vList As New ParameterList(True)
    Dim vShowGrid As Boolean
    Dim vSelectExisting As Boolean
    Dim vShowEditPanel As Boolean = True
    cmdDelete.Visible = False
    cmdClose.Visible = False
    cmdLink1.Visible = False
    cmdLink2.Visible = False
    epl.Visible = False
    Select Case mvCampaignDataType
      Case CareServices.XMLCampaignDataSelectionTypes.xcadtCampaign
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaign
        mvCampaignItem.FillParameterList(vList)
        vSelectExisting = mvCampaignItem.Existing
        cmdSave.Enabled = True
        If sel.HasDependants = False AndAlso vSelectExisting Then cmdDelete.Visible = True
        sel.TreeContextMenu = mvCampaignMenu
      Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal
        mvCampaignItem.FillParameterList(vList)
        vSelectExisting = mvCampaignItem.Existing
        If sel.HasDependants = False AndAlso vSelectExisting Then cmdDelete.Visible = True
        cmdSave.Visible = True
        sel.TreeContextMenu = mvCampaignMenu
      Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegment
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegment
        mvCampaignItem.FillParameterList(vList)
        vSelectExisting = mvCampaignItem.Existing
        If sel.HasDependants = False AndAlso vSelectExisting Then cmdDelete.Visible = True
        cmdSave.Visible = True
        sel.TreeContextMenu = mvCampaignMenu
      Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollection
        Select Case mvCampaignItem.AppealType
          Case CampaignItem.AppealTypes.atMannedCollection
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection
          Case CampaignItem.AppealTypes.atUnMannedCollection
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
          Case CampaignItem.AppealTypes.atH2HCollection
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
        End Select
        mvCampaignItem.FillParameterList(vList)
        vSelectExisting = mvCampaignItem.Existing
        If sel.HasDependants = False AndAlso vSelectExisting Then cmdDelete.Visible = True
        sel.TreeContextMenu = mvCampaignMenu
      Case CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppealActions, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCostCentres, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCosts, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectors, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectionPIS, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectionBoxes, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentProducts, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtTickBoxes, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtUnMannedCollectionBoxes, _
       CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignRoles
        mvCampaignItem.FillParameterList(vList)
        vSelectExisting = mvCampaignItem.Existing
        vShowGrid = True
        dgr.ContextMenuStrip = Nothing
        Select Case mvCampaignDataType
          Case CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppealActions
            mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction
            cmdLink1.Text = ControlText.CmdActionLinks
            cmdLink1.Visible = True
            cmdLink2.Text = ControlText.CmdActionSubjects
            cmdLink2.Visible = True
            mvActionMenu = New ActionMenu(Me)
            mvActionMenu.ActionType = ActionMenu.ActionTypes.CampaignActions
            If mvSelectedRow > -1 Then
              mvActionNumber = IntegerValue(dgr.GetValue(mvSelectedRow, "ActionNumber"))
              mvActionMenu.ActionNumber = mvActionNumber
              mvActionMenu.ActionStatus = dgr.GetValue(mvSelectedRow, "ActionStatus")
              mvActionMenu.MasterActionNumber = IntegerValue(dgr.GetValue(mvSelectedRow, "MasterAction"))
            End If
            dgr.ContextMenuStrip = mvActionMenu
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppealResources
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionResources
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCostCentres
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegmentCostCentre
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCosts
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaignCosts
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectors
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollectors
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectionPIS
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone
            vShowEditPanel = False
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectionBoxes
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentProducts
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaignSuppliers
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtTickBoxes
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTickBox
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtUnMannedCollectionBoxes
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes
          Case CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignRoles
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaignRoles
        End Select
        sel.TreeContextMenu = Nothing
    End Select

    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone Then
      epl.Visible = False
    Else
      If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
        epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))
        epl.FillDeferredCombos(epl)
        epl.ContextMenuStrip = mvCustomiseMenu
        mvCustomiseMenu.SetContext(Nothing, mvMaintenanceType, "")
        epl.Refresh()
      End If
      epl.Visible = True
    End If
    'populate any combos that need to be done, as this has to be done before populating the grids
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
        'fill the collectors combo
        Dim vTable As DataTable
        vList("CollectionNumber") = mvCampaignItem.CollectionNumber.ToString
        Dim vCollectorDataSet As DataSet
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes Then
          vCollectorDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors, vList)
        Else
          vCollectorDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectors, vList)
        End If
        vTable = DataHelper.GetTableFromDataSet(vCollectorDataSet)
        If vTable Is Nothing Then
          vTable = New DataTable
          vTable.Columns.AddRange(New DataColumn() {New DataColumn("ContactName"), New DataColumn("CollectorNumber")})
        End If
        epl.SetComboDataSource("CollectorNumber", "CollectorNumber", "ContactName", vTable, True)

        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes Then
          Dim vPISDataSet As DataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, vList)
          vTable = DataHelper.GetTableFromDataSet(vPISDataSet)
          If vTable Is Nothing Then
            vTable = New DataTable
            vTable.Columns.AddRange(New DataColumn() {New DataColumn("CollectionPisNumber"), New DataColumn("PisNumber")})
          End If
          Dim vDataRow As DataRow = vTable.NewRow
          vTable.Rows.InsertAt(vDataRow, 0)
          epl.SetComboDataSource("CollectionPisNumber", "CollectionPisNumber", "PisNumber", vTable, False)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink, _
         CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        vShowGrid = True
        cmdClose.Visible = True
        cmdLink1.Visible = False
        cmdLink2.Visible = False
        dgr.ContextMenuStrip = Nothing
    End Select
    Dim vDataSet As DataSet


    If vSelectExisting Then
      vDataSet = DataHelper.GetCampaignData(mvCampaignDataType, vList)
      If mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtCampaign Then
        If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows(0).Item("MarkHistorical").ToString() = "Y" Then
          mvCampaignMarkHistorical = True
        Else
          mvCampaignMarkHistorical = False
        End If
      End If
      mvCampaignItem.MarkHistorical = mvCampaignMarkHistorical
    Else
      vDataSet = New DataSet
    End If

    If vShowGrid Then
      ShowGrid()
      dgr.Populate(vDataSet)
    Else
      HideGrid()
    End If
    mvCampaignMenu.CriteriaSet = 0
    With epl
      If vDataSet.Tables.Contains("DataRow") Then
        mvEditing = True
        If vShowGrid Then
          SelectRow(0)
        Else
          .Populate(vDataSet.Tables("DataRow").Rows(0))
        End If
        SetButtons(vShowGrid)
        SetTabEditingControls(vDataSet)
      Else
        mvEditing = False
        epl.Clear()
        'Now Set Defaults for new item
        SetTabNotEditingControls()
        SetDefaults()
      End If
      'settings for both when editing or not editing
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS
          cmdOther.Text = ControlText.CmdCollectionBoxes
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions
          cmdOther.Text = ControlText.CmdPoints
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionResources
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).AppealType = CampaignItem.AppealTypes.atUnMannedCollection AndAlso sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).ResourcesProducedOn IsNot Nothing AndAlso IsDate(sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).ResourcesProducedOn) Then
            'the collection's resources produced on date is set. we should not allow any changes to the resources for this collection.
            epl.EnableControls(epl, False)
          End If
        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
          cmdOther.Text = ControlText.CmdShifts
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
          cmdOther.Text = ControlText.CmdAppealBudgetDetails
      End Select
      If mvCampaignItem.Existing = False AndAlso mvCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citCampaign Then .Focus()
    End With
    bpl.RepositionButtons()
    epl.DataChanged = False
    SetButtons(vShowGrid)
    epl.Visible = vShowEditPanel
  End Sub
  Private Sub ValidateAllItems(ByVal pSender As Object, ByVal pList As ParameterList, ByRef pValid As Boolean) Handles epl.ValidateAllItems
    Dim vList As New ParameterList(True)

    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
        ValidateItem(pSender, "Appeal", pList("Appeal"), pValid)
        Dim vParentCampaignItem As CampaignItem = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCampaign)
        pValid = CheckParentDates(vParentCampaignItem.StartDate, vParentCampaignItem.EndDate, epl.GetValue("AppealDate"), epl.GetValue("EndDate"), "AppealDate", "EndDate", InformationMessages.ImAppealDatesOutOfRange)
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
        If mvEditing = False AndAlso dgr.FindRow("BudgetPeriod", pList("BudgetPeriod")) >= 0 Then
          epl.SetErrorField("BudgetPeriod", InformationMessages.ImRecordAlreadyExists)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
        ValidateItem(pSender, "Campaign", pList("Campaign"), pValid)
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS
        If pList("Amount").Length > 0 Xor epl.GetValue("BankedBy").Length > 0 Then
          epl.SetErrorField("Amount", InformationMessages.ImBankedByFields)
          pValid = False
        ElseIf epl.GetValue("BankedOn").Length > 0 Xor pList("Amount").Length > 0 Then
          epl.SetErrorField("Amount", InformationMessages.ImBankedByFields)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
        If IsDate(epl.GetValue("StartTime")) Or IsDate(epl.GetValue("EndTime")) Then
          pValid = CheckParentDates(mvCampaignItem.StartTime, mvCampaignItem.EndTime, epl.GetValue("StartTime"), epl.GetValue("EndTime"), "StartTime", "EndTime", InformationMessages.ImShiftTimesOutOfRange)
          If pValid Then
            Dim vNewStartTime As DateTime = epl.GetDateTimeValue("StartTime")
            Dim vNewEndTime As DateTime = epl.GetDateTimeValue("EndTime")
            Dim vTimeAllocated As TimeSpan = vNewEndTime.Subtract(vNewStartTime)

            Dim vMaxTime As TimeSpan
            TimeSpan.TryParse(mvCampaignItem.AdditionalValues("TotalTime").ToString, vMaxTime)

            Dim vindex As Integer
            Dim vStartTime As DateTime
            Dim vEndTime As DateTime

            For vindex = 0 To dgr.DataRowCount - 1
              If vindex <> dgr.CurrentRow Then
                DateTime.TryParse(dgr.GetValue(vindex, "StartTime"), vStartTime)
                DateTime.TryParse(dgr.GetValue(vindex, "EndTime"), vEndTime)
                If (vNewStartTime >= vStartTime And vNewStartTime <= vEndTime) Or (vNewEndTime >= vStartTime And vNewEndTime <= vEndTime) Then
                  epl.SetErrorField("StartTime", InformationMessages.ImOverlappingShiftTimes)
                  pValid = False
                End If
                vTimeAllocated = vTimeAllocated + vEndTime.Subtract(vStartTime)
              End If
            Next
            If pValid AndAlso mvCampaignItem.AdditionalValues("TotalTime").ToString.Length > 0 AndAlso vTimeAllocated > vMaxTime Then ShowWarningMessage(InformationMessages.ImShiftTimesExceedsMaxTime)
          End If
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSegment
        ValidateItem(pSender, "Segment", pList("Segment"), pValid)
        Dim vParentCampaignItem As CampaignItem = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCampaign)
        pValid = CheckParentDates(vParentCampaignItem.StartDate, vParentCampaignItem.EndDate, epl.GetValue("SegmentDate"), epl.GetValue("SegmentDate"), "SegmentDate", "SegmentDate", InformationMessages.ImSegmentDatesOutOfRange)

        If pValid Then ValidateItem(pSender, "SegmentSequence", pList("SegmentSequence"), pValid)
        If pList("RequiredCount").Length > 0 AndAlso pList("Random") = "N" AndAlso pList("Score").Length = 0 Then
          epl.SetErrorField("Random", InformationMessages.ImRandomOrScore)
          pValid = False
        End If
        If pList("Direction") = "O" AndAlso pList("OutputGroup").Length = 0 Then
          epl.SetErrorField("OutputGroup", InformationMessages.ImOutputGroupMissing)
          pValid = False
        End If
        If pList("MailingType") = "SR" AndAlso IntegerValue(pList("DespatchQuantity")) <= 0 Then
          epl.SetErrorField("DespatchQuantity", InformationMessages.ImDespatchQuantitySR)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSegmentCostCentre
        If dgr.FindRow("CostCentre", pList("CostCentre")) >= 0 And Not mvEditing Then
          epl.SetErrorField("CostCentre", InformationMessages.ImRecordAlreadyExists)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct
        If dgr.FindRow("AmountNumber", pList("AmountNumber")) >= 0 And Not mvEditing Then
          epl.SetErrorField("AmountNumber", InformationMessages.ImRecordAlreadyExists)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctTickBox
        If dgr.FindRow("TickBoxNumber", pList("TickBoxNumber")) >= 0 And Not mvEditing Then
          epl.SetErrorField("TickBoxNumber", InformationMessages.ImRecordAlreadyExists)
          pValid = False
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
        ValidateItem(pSender, "Collection", pList("Collection"), pValid)
        Dim vParentCampaignItem As CampaignItem = sel.GetParentCampaignItem(mvCampaignItem.ParentItemType)
        Dim vParameter As String = ""
        Dim vEndDate As String = Nothing
        Dim vError As String = ""
        Dim vEndDateParam As String = Nothing
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection
            vParameter = "CollectionDate"
            vEndDateParam = "CollectionDate"
            vEndDate = epl.GetValue("CollectionDate")
            vError = InformationMessages.ImCollectionDatesOutOfRange
          Case CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
               CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
            vParameter = "StartDate"
            vEndDate = epl.GetValue("EndDate")
            vEndDateParam = "EndDate"
            vError = InformationMessages.ImCollectionDatesOutOfRange
        End Select
        pValid = CheckParentDates(vParentCampaignItem.StartDate, vParentCampaignItem.EndDate, epl.GetValue(vParameter), vEndDate, vParameter, vEndDateParam, vError)
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
        If epl.GetValue("ForecastUnits").Length = 0 AndAlso epl.GetValue("BudgetedCosts").Length = 0 AndAlso epl.GetValue("BudgetedIncome").Length = 0 Then
          pValid = epl.SetErrorField("ForecastUnits", InformationMessages.ImBudgetInformationRequired)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
        If mvEditing AndAlso epl.GetValue("TotalTime").Length > 0 Then
          GetPrimaryKeyValues(vList, dgr.CurrentRow, True)
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectorShifts, vList))
          If vTable IsNot Nothing Then
            Dim vDataRow As DataRow
            Dim vTimeAllocated As TimeSpan
            Dim vMaxTime As TimeSpan
            TimeSpan.TryParse(epl.GetValue("TotalTime"), vMaxTime)
            For Each vDataRow In vTable.Rows
              Dim vStartTime As DateTime
              Dim vEndTime As DateTime
              DateTime.TryParse(vDataRow.Item("StartTime").ToString, vStartTime)
              DateTime.TryParse(vDataRow.Item("EndTime").ToString, vEndTime)
              vTimeAllocated = vTimeAllocated + vEndTime.Subtract(vStartTime)
            Next
            If pValid AndAlso vTimeAllocated > vMaxTime Then ShowWarningMessage(InformationMessages.ImShiftTimesExceedsMaxTime)
          End If
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctCampaignRoles
        If dgr.FindRow("ContactNumber,CampaignRole", pList("ContactNumber").ToString & "," & pList("CampaignRole").ToString) >= 0 Then
          epl.SetErrorField("CampaignRole", InformationMessages.ImRecordAlreadyExists)
          pValid = False
        End If
    End Select
  End Sub
  Private Sub ValidateItem(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean) Handles epl.ValidateItem
    Dim vList As New ParameterList(True)

    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
        If pParameterName = "Appeal" AndAlso Not mvEditing Then
          vList("Campaign") = mvCampaignItem.Campaign
          vList(pParameterName) = pValue
          Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftCampaignAppeals, vList)
          If vDataSet IsNot Nothing AndAlso DataHelper.GetRowFromDataSet(vDataSet) IsNot Nothing Then
            epl.SetErrorField(pParameterName, InformationMessages.ImRecordAlreadyExists)
            pValid = False
          End If
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
        If pParameterName = "Campaign" AndAlso Not mvEditing Then
          vList(pParameterName) = pValue
          Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftCampaigns, vList)
          If vDataSet IsNot Nothing AndAlso DataHelper.GetRowFromDataSet(vDataSet) IsNot Nothing Then
            epl.SetErrorField(pParameterName, InformationMessages.ImRecordAlreadyExists)
            pValid = False
          End If
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSegment
        If pParameterName = "Segment" AndAlso Not mvEditing Then
          vList("Campaign") = mvCampaignItem.Campaign
          vList("Appeal") = mvCampaignItem.Appeal
          vList(pParameterName) = pValue
          Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftCampaignSegments, vList)
          If vDataSet IsNot Nothing AndAlso DataHelper.GetRowFromDataSet(vDataSet) IsNot Nothing Then
            epl.SetErrorField(pParameterName, InformationMessages.ImRecordAlreadyExists)
            pValid = False
          End If
        ElseIf pParameterName = "SegmentSequence" Then
          pValid = sel.NewSequenceValid(IntegerValue(pValue))
          If Not pValid Then epl.SetErrorField(pParameterName, InformationMessages.ImRecordAlreadyExists)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
        If pParameterName = "Collection" AndAlso Not mvEditing Then
          vList("Campaign") = mvCampaignItem.Campaign
          vList("Appeal") = mvCampaignItem.Appeal
          vList(pParameterName) = pValue
          Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftCampaignCollections, vList)
          If vDataSet IsNot Nothing AndAlso DataHelper.GetRowFromDataSet(vDataSet) IsNot Nothing Then
            epl.SetErrorField(pParameterName, InformationMessages.ImRecordAlreadyExists)
            pValid = False
          End If
        ElseIf pParameterName = "ContactNumber" AndAlso mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection Then
          'Ensure that the contact number is not an Organisation
          Dim vCtl As TextLookupBox = epl.FindTextLookupBox(pParameterName, False)
          Dim vContactInfo As ContactInfo = vCtl.ContactInfo
          epl.SetErrorField(pParameterName, "")
          If vContactInfo.ContactType <> ContactInfo.ContactTypes.ctContact Then
            epl.SetErrorField(pParameterName, InformationMessages.ImOrganisationNotValidAsContact)
            'Remove the description to force future validation (e.g. before save) to fail. Don't know why the epl.SetErrorField doesn't do that
            vCtl.Label.Text = ""
            pValid = False
          End If
        End If
    End Select
  End Sub

  Private Sub ShowGrid()
    dgr.Visible = True
    splRight.Panel1Collapsed = False
  End Sub
  Private Sub GetPrimaryKeyValues(ByVal pList As ParameterList, ByVal pRow As Integer, ByVal pForUpdate As Boolean)
    mvCampaignItem.FillParameterList(pList)

    mvKeyValuesColl = GetPrimaryKeyNames(mvMaintenanceType, pForUpdate)
    If mvKeyValuesColl.Count > 0 Then
      For Each vDGKeyValue As DisplayGridKeyValue In mvKeyValuesColl
        pList(vDGKeyValue.ParameterName) = dgr.GetValue(pRow, vDGKeyValue.GridColumnName)
      Next

      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctSegment
          If mvMailingType Is Nothing Then mvMailingType = ""
          pList("MailingType") = mvMailingType        'Used by validation in epl.AddValuesToList
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
          If Not pForUpdate Then pList.Remove("AppealBudgetNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctTickBox
          If Not pForUpdate Then pList.Remove("TickBoxNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct
          If Not pForUpdate Then pList.Remove("AmountNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctSegmentCostCentre
          If Not pForUpdate Then pList.Remove("CostCentre")
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions
          If Not pForUpdate Then pList.Remove("CollectionRegionNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors, CareServices.XMLMaintenanceControlTypes.xmctH2HCollectors
          If Not pForUpdate Then pList.Remove("CollectorNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes, CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes
          If Not pForUpdate Then pList.Remove("CollectionBoxNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
          If Not pForUpdate Then pList.Remove("CollectionPISNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionResources
          If Not pForUpdate Then pList.Remove("CollectionResourceNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
          If Not pForUpdate Then pList.Remove("CollectionPointNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
          If Not pForUpdate Then pList.Remove("CollectorShiftNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealResources
          If Not pForUpdate Then pList.Remove("AppealResourceNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
          If Not pForUpdate Then pList.Remove("CollectionPisNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCampaignSuppliers
          If Not pForUpdate Then pList.Remove("OrganisationNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
          If Not pForUpdate Then pList.Remove("AppealBudgetDetailsNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctCampaignRoles
          If Not pForUpdate Then pList.Remove("ContactCampaignRoleNumber")
      End Select
    End If
  End Sub
  Private Sub GetAdditionalKeyValues(ByVal pList As ParameterList)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
        pList("Campaign") = mvCampaignItem.Campaign
        If epl.GetValue("AppealType") = "S" Then
          pList("ReadyForConfirmation") = "N"
        Else
          pList("MailJoints") = "N"
          pList("CombineMail") = "N"
          pList("BypassCount") = "N"
          pList("CreateMailingHistory") = "N"
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSegment
        pList("Campaign") = mvCampaignItem.Campaign
        pList("Appeal") = mvCampaignItem.Appeal
        pList("AppealType") = mvCampaignItem.AppealTypeCode
        pList("MailingType") = mvMailingType        'Used by validation in epl.AddValuesToList
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget, _
           CareServices.XMLMaintenanceControlTypes.xmctAppealResources, _
           CareServices.XMLMaintenanceControlTypes.xmctTickBox, _
           CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct, _
           CareServices.XMLMaintenanceControlTypes.xmctSegmentCostCentre, _
           CareServices.XMLMaintenanceControlTypes.xmctCampaignSuppliers, _
           CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions, _
           CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors, _
           CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes, _
           CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, _
           CareServices.XMLMaintenanceControlTypes.xmctCollectionResources, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectors, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS, _
           CareServices.XMLMaintenanceControlTypes.xmctCampaignCosts, _
           CareServices.XMLMaintenanceControlTypes.xmctCampaignRoles, _
           CareNetServices.XMLMaintenanceControlTypes.xmctAction, _
           CareNetServices.XMLMaintenanceControlTypes.xmctActionLink, _
           CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic
        mvCampaignItem.FillParameterList(pList)
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
        mvCampaignItem.FillParameterList(pList)
        pList("MailJoints") = "N"
      Case CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
        pList.IntegerValue("CollectionNumber") = mvCampaignItem.CollectionNumber
        pList.IntegerValue("CollectionPISNumber") = IntegerValue(mvCampaignItem.AdditionalValues("CollectionPISNumber").ToString)
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
        mvCampaignItem.FillParameterList(pList)
        pList.IntegerValue("CollectionRegionNumber") = IntegerValue(mvCampaignItem.AdditionalValues("CollectionRegionNumber").ToString)
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
        mvCampaignItem.FillParameterList(pList)
        pList.IntegerValue("CollectorNumber") = IntegerValue(mvCampaignItem.AdditionalValues("CollectorNumber").ToString)
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
        mvCampaignItem.FillParameterList(pList)
        pList.IntegerValue("AppealBudgetNumber") = IntegerValue(mvCampaignItem.AdditionalValues("AppealBudgetNumber").ToString)
    End Select
  End Sub

  Protected Overrides Function ProcessSave(ByVal pDefault As Boolean, ByVal sender As System.Object) As Boolean 'Return true if saved
    Try
      Dim vList As New ParameterList(True)
      If mvEditing Then
        'If editing an existing record then get the primary key values
        GetPrimaryKeyValues(vList, mvSelectedRow, True)
      Else
        'For new records add in any additional key values
        GetAdditionalKeyValues(vList)
      End If
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS Then
        'Special case as we have a ListBox containing multiple values
        'We are updating CollectionBoxes to add CollectionPISNumber
        vList.Remove("CollectionNumber")
        vList.Add("CollectionBoxNumber", "")
        Dim vLB As ListBox = DirectCast(epl.FindPanelControl("BoxReference"), ListBox)
        If vLB.SelectedItems.Count > 0 Then
          'Use BoxReference ListBox
          For Each vRow As DataRowView In vLB.SelectedItems
            vList("CollectionBoxNumber") = vRow.Item("CollectionBoxNumber").ToString
            DataHelper.UpdateItem(mvMaintenanceType, vList)
          Next
          Return True
        Else
          epl.AddValuesToList(vList, True)    'This will set the ErrorProvider
        End If
      ElseIf epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
        If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction Then
          vList("FromCampaign") = "Y"
        End If

        If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or _
          mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
          vList("ActionNumber") = mvActionNumber.ToString
          vList("Notified") = "N"
        End If

        If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAppeal Then
          Dim vReadOnlyItems() As String = {"OrgMailTo", "OrgMailWhere", "OrgMailAddrUsage", "OrgMailLabelName", "OrgMailRoles", "VariableParameters"}
          For Each vItem As String In vReadOnlyItems
            If vList.ContainsKey(vItem) Then vList.Remove(vItem)
          Next
        End If

        'Update or Insert record
        If mvEditing Then
          If ConfirmUpdate() = False Then Exit Function
          mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
        Else
          If ConfirmInsert() = False Then Exit Function
          mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
        End If
        mvRefreshParent = True
        epl.DataChanged = False     'Data saved now

        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAction And Not mvEditing Then
          If DataHelper.UserInfo.ContactNumber > 0 And mvReturnList("ActionNumber").ToString.Length > 0 Then
            If ShowQuestion(QuestionMessages.QmNoActioners, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vList("ActionNumber") = mvReturnList("ActionNumber").ToString
              vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
              vList("ActionLinkType") = "A"
              vList("Notified") = "N"
              DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
            End If
          End If
        End If
        'Save the primary key values for Selection
        If Not mvEditing Then
          mvKeyValuesColl = GetPrimaryKeyNames(mvMaintenanceType)
          If mvKeyValuesColl.Count > 0 Then
            For Each vDGKeyValue As DisplayGridKeyValue In mvKeyValuesColl
              If mvReturnList IsNot Nothing AndAlso mvReturnList.Count > 0 Then
                If mvReturnList.Contains(vDGKeyValue.ParameterName) Then vDGKeyValue.Value = mvReturnList(vDGKeyValue.ParameterName)
              ElseIf vList.Count > 0 Then
                If vList.Contains(vDGKeyValue.ParameterName) Then vDGKeyValue.Value = vList(vDGKeyValue.ParameterName)
              End If
            Next
          End If
        End If
        RefreshTabSelector(vList)
        Return True
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
        Case CareException.ErrorNumbers.enNotEnoughQtyOnAppeal, _
             CareException.ErrorNumbers.enSpecifiedDataNotFound, _
             CareException.ErrorNumbers.enBudgetDetailsNotAllowed, _
             CareException.ErrorNumbers.enInvalidBudgetPeriodDates, _
             CareException.ErrorNumbers.enCollectionsErrorFromWebService, _
             CareException.ErrorNumbers.enOverlappingShiftTimes, _
             CareException.ErrorNumbers.enAppointmentConflict, _
             CareException.ErrorNumbers.enSegmentSourceCodeCannotBeDerived
          ShowInformationMessage(vEx.Message)
        Case Else
          Throw vEx
      End Select
    End Try
  End Function

  Protected Overrides Sub SetCommandsForNew()
    mvEditing = False
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
        epl.SetValue("ReadyForAcknowledgement", AppValues.DefaultReadyFlag)
        epl.SetValue("ReadyForConfirmation", AppValues.DefaultReadyFlag)
        epl.SetValue("Attended", AppValues.DefaultAttended)
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
        Dim vControl As Control = FindControl(epl, "PisNumber")
        If TryCast(vControl, ComboBox) IsNot Nothing Then
          Dim vList As ParameterList = New ParameterList(True)
          vList("BankAccount") = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).CollectionBankAccount
          Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPISNumbers, vList)
          epl.SetComboDataSource("PisNumber", "PISNumber", "PISNumber", vTable, False)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
        epl.SetValue("GeographicalRegionDesc", mvCampaignItem.AdditionalValues("GeographicalRegionDesc"))
        epl.SetDependancies("CollectionPointType")
      Case CareServices.XMLMaintenanceControlTypes.xmctH2HCollectors
        epl.SetValue("CollectorStatus", AppValues.ControlValue(AppValues.ControlValues.default_collector_status))
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionResources
        epl.SetValue("Product", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).FirstAppealResource)
        epl.SetValue("DespatchOn", AppValues.TodaysDate)
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
        epl.SetValue("ActionPriority", AppValues.DefaultActionPriority)
        Dim vTimeSpan As TimeSpan = AppValues.DefaultActionDuration
        If vTimeSpan.Days > 0 Then epl.SetValue("DurationDays", vTimeSpan.Days.ToString)
        If vTimeSpan.Hours > 0 Then epl.SetValue("DurationHours", vTimeSpan.Hours.ToString)
        If vTimeSpan.Minutes > 0 Then epl.SetValue("DurationMinutes", vTimeSpan.Minutes.ToString)
    End Select
    SetButtons(True)
    If epl.Visible Then epl.Focus()
  End Sub

  Protected Overrides Sub SetDefaults(Optional ByVal pInitialSetup As Boolean = True)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
        SetItemAsMaxColumn("BudgetPeriod")
      Case CareServices.XMLMaintenanceControlTypes.xmctSegmentCostCentre
        epl.EnableControl("CostCentre", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct
        SetItemAsMaxColumn("AmountNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctTickBox
        SetItemAsMaxColumn("TickBoxNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
        epl.SetValue("BudgetPeriod", mvCampaignItem.AdditionalValues("BudgetPeriod"), True)
        epl.EnableControl("Segment", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
        epl.SetValue("TotalTime", AppValues.DefaultTotalTime)
        Dim vDTP As DateTimePicker = epl.FindDateTimePicker("TotalTime")
        vDTP.Checked = False
      Case CareServices.XMLMaintenanceControlTypes.xmctCampaignSuppliers
        epl.EnableControl("OrganisationNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
        Dim vControl As Control = FindControl(epl, "PisNumber")
        Dim vTextLookupBox As TextLookupBox = TryCast(vControl, TextLookupBox)
        If vTextLookupBox IsNot Nothing Then
          DirectCast(vControl, TextLookupBox).ValidationRequired = True
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
        epl.SetValue("ActionPriority", AppValues.DefaultActionPriority)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
        epl.SetEntityLinkDefaults("ActionLinkType", "R")
    End Select
    epl.DataChanged = False
  End Sub

  Private Sub SetItemAsMaxColumn(ByVal pColumnName As String)

    epl.SetValue(pColumnName, CStr(dgr.MaxColumnValue(pColumnName) + 1))
    epl.EnableControl(pColumnName, True)
  End Sub

  Protected Overrides Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
          CampaignTabSelected(sel, CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, mvCampaignItem)
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
          CampaignTabSelected(sel, CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors, mvCampaignItem)
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
          CampaignTabSelected(sel, CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions, mvCampaignItem)
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
          CampaignTabSelected(sel, CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets, mvCampaignItem)
        Case CareServices.XMLMaintenanceControlTypes.xmctActionLink, CareServices.XMLMaintenanceControlTypes.xmctActionTopic
          CampaignTabSelected(sel, CareServices.XMLCampaignDataSelectionTypes.xcadtAppealActions, mvCampaignItem)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overrides Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vBusyCursor As New BusyCursor()

    Try
      'TODO Confirm Update or Insert 
      Dim vDefault As Boolean
      If mvCampaignItem.Appeal <> "" AndAlso mvCampaignItem.AppealLocked Then
        epl.Visible = False
        dgr.Visible = False
        SetButtons(False, True)
        ShowInformationMessage(InformationMessages.ImAppealLocked)
      Else
        If ProcessSave(vDefault, sender) Then
          If dgr.Visible Then     'Only one record
            RePopulateGrid()
          End If
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS Then RePopulateListBox(DirectCast(epl.FindPanelControl("BoxReference"), ListBox))
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink Or _
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic Then
            ProcessNew()
          End If
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub
  Protected Overrides Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      'TODO Confirm cancel changes
      mvReturnList = Nothing                  'Clear this as it should only be valid after a save (new/update)
      Dim vList As New ParameterList(True)
      GetPrimaryKeyValues(vList, mvSelectedRow, True)
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS Then
        'This will update the CollectionBox to remove the PIS Number (effectively deleting the link between PIS and Box)
        vList("CollectionBoxNumber") = dgr.GetValue(mvSelectedRow, "CollectionBoxNumber")
        vList.Item("CollectionPisNumber") = ""    'This is being set to Null
        If mvCampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection Then
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes, vList)
        ElseIf mvCampaignItem.AppealType = CampaignItem.AppealTypes.atUnMannedCollection Then
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes, vList)
        End If
      Else
        If Not ConfirmDelete() Then Exit Sub
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctAppeal, CareServices.XMLMaintenanceControlTypes.xmctSegment, _
               CareServices.XMLMaintenanceControlTypes.xmctH2HCollection, CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
               CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
            Dim vCampaignCopyType As CampaignCopyInfo.CampaignCopyTypes
            Select Case mvMaintenanceType
              Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
                vCampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctAppeal
              Case CareServices.XMLMaintenanceControlTypes.xmctSegment
                vCampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctSegment
              Case CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
                vCampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctH2HCollection
              Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection
                vCampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctMannedCollection
              Case CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
                vCampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctUnMannedCollection
            End Select
            If FormHelper.ClipboardContainsCampaignData(vCampaignCopyType, mvCampaignItem) Then
              Clipboard.Clear()
            End If
          Case CareNetServices.XMLMaintenanceControlTypes.xmctAction
            vList("FromCampaign") = "Y"
            vList("MasterActionNumber") = dgr.GetValue(mvSelectedRow, "MasterAction")
          Case CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic, _
              CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
            vList("FromCampaign") = "Y"
            vList("ActionNumber") = mvActionNumber.ToString
            Select Case dgr.GetValue(mvSelectedRow, "EntityType").ToUpper
              Case "D"
                'Document
                vList("DocumentNumber") = vList("ContactNumber")
                vList.Remove("ContactNumber")
              Case "N"
                'Exam Centre
                vList("ExamCentreId") = vList("ContactNumber")
                vList.Remove("ContactNumber")
              Case Else
                'Contact / Organisation
            End Select
        End Select
        DataHelper.DeleteItem(mvMaintenanceType, vList)
      End If
      mvRefreshParent = True
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
          Me.Close() 'Deleted the only item
        Case CareServices.XMLMaintenanceControlTypes.xmctAppeal, _
             CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctH2HCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctSegment
          sel.RemoveSelectedNode()
          If mvCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citAppeal Then mvCampaignItem.SegmentCount -= 1
        Case CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
          If dgr.Visible Then RePopulateGrid()
          RePopulateListBox(DirectCast(epl.FindPanelControl("BoxReference"), ListBox))
        Case Else
          If dgr.Visible Then RePopulateGrid()
          'do nothing...these are the nodes with grids. the delete should only delete a grid row.
      End Select
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enNotEnoughQtyOnAppeal, _
             CareException.ErrorNumbers.enSpecifiedDataNotFound, _
             CareException.ErrorNumbers.enCollectionsErrorFromWebService
          ShowInformationMessage(vEx.Message)
        Case Else
          Throw vEx
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      If dgr.Visible Then RePopulateGrid()
    End Try
  End Sub

  Private Sub BeforeSelect(ByVal pSender As Object, ByRef pCancel As Boolean) Handles sel.BeforeSelect
    Dim vChangeNode As Boolean = True
    If epl.DataChanged Then
      If cmdSave.Enabled AndAlso ConfirmSave() Then
        Dim vSave As Boolean = ProcessSave(False, pSender)
        pCancel = Not vSave
        vChangeNode = vSave
        pCancel = True
      Else
        'We have some data changed and we are going to cancel it - Check if it was a new appeal or segment
        If mvEditing = False Then
          epl.DataChanged = False
          'sel.RemoveSelectedNode()
        End If
      End If
    End If

    If vChangeNode AndAlso _
      (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaign _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegment _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollection) Then
      If Not mvCampaignItem.Existing And Not mvEditing Then
        sel.RemoveSelectedNode()
      End If
    End If
  End Sub

  'Private Sub RowChanging(ByRef pCancel As Boolean) Handles dgr.RowChanging
  '  Dim vChangeRow As Boolean = True

  '  If epl.DataChanged Then
  '    If ConfirmCancel() = False Then
  '      pCancel = True
  '      vChangeRow = False
  '    Else
  '      'We have some data changed and we are going to cancel it - Check if it was a new appeal or segment
  '      If mvEditing = False Then
  '        epl.DataChanged = False
  '        'sel.RemoveSelectedNode()
  '      End If
  '    End If
  '  End If

  '  'If vChangeRow AndAlso _
  '  '  (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaign _
  '  '   Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal _
  '  '   Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegment _
  '  '   Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection _
  '  '   Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection _
  '  '   Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollection) Then
  '  '  If Not mvCampaignItem.Existing And Not mvEditing Then
  '  '    sel.RemoveSelectedNode()
  '  '  End If
  '  'End If
  'End Sub

  Private Sub GetCodeRestrictions(ByVal pSender As Object, ByVal pParameterName As String, ByVal pList As ParameterList) Handles epl.GetCodeRestrictions
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctSegmentProduct
        pList("FindProductType") = "B"      'Donation or Product sale
      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionResources
        pList("Campaign") = mvCampaignItem.Campaign
        pList("Appeal") = mvCampaignItem.Appeal

      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctAppeal
        pList("FindProductType") = "D"      'Donation

      Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, _
           CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
        pList("BankAccount") = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).CollectionBankAccount
    End Select
  End Sub

  Private Sub RePopulateGrid()
    Dim vList As ParameterList = New ParameterList(True)
    Dim vDataType As CareServices.XMLCampaignDataSelectionTypes = mvCampaignDataType
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS Then vDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtContactCollectionBoxes

    GetAdditionalKeyValues(vList)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink
        dgr.Populate(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, mvActionNumber))
      Case CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        dgr.Populate(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvActionNumber))
      Case Else
        dgr.Populate(DataHelper.GetCampaignData(vDataType, vList))
    End Select

    If mvActionMenu IsNot Nothing And dgr.DataRowCount = 0 Then mvActionMenu.ActionNumber = 0

    If dgr.DataRowCount = 0 Then
      ProcessNew()
    Else
      'Select current row
      If Not mvEditing Then
        mvSelectedRow = dgr.FindRow(mvKeyValuesColl.GridColumnNames, mvKeyValuesColl.Values)
      Else
        If mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignRoles Then
          If mvReturnList IsNot Nothing Then mvSelectedRow = dgr.FindRow("ContactCampaignRoleNumber", mvReturnList("ContactCampaignRoleNumber"))
        End If
      End If
      If mvSelectedRow <= 0 Then mvSelectedRow = 0 'TODO Find the records which have just been added
      If mvSelectedRow > dgr.DataRowCount - 1 Then mvSelectedRow = dgr.DataRowCount - 1
      dgr.SelectRow(mvSelectedRow)
      mvEditing = True
      SelectRow(mvSelectedRow)
      bpl.RepositionButtons()
    End If
    SetButtons(True)
  End Sub

  Protected Overrides Sub SelectRow(ByVal pRow As Integer)
    Dim vBusyCursor As New BusyCursor

    Try
      If pRow >= 0 Then
        Dim vList As New ParameterList(True)
        vList("SystemColumns") = "N"                'Ensure we get all the columns
        Dim vCount As Integer = vList.Count
        Dim vDataRow As DataRow
        If mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS _
           And mvCampaignDataType <> CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments Then
          GetPrimaryKeyValues(vList, pRow, True)
          If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction Then
            vDataRow = DataHelper.GetRowFromDataSet(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionInformation, vList.IntegerValue("ActionNumber")))
          ElseIf mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or _
            mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
            vDataRow = Nothing
          Else
            vDataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(mvCampaignDataType, vList))
          End If
          If vList.Count > vCount Then
            If vDataRow IsNot Nothing Then epl.Populate(vDataRow)
            mvEditing = True
          End If
          mvSelectedRow = pRow
          Select Case mvCampaignDataType
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets
              epl.EnableControl("BudgetPeriod", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgetDetails
              epl.EnableControl("BudgetPeriod", False)
              epl.EnableControl("Segment", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtTickBoxes
              epl.EnableControl("TickBoxNumber", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentProducts
              epl.EnableControl("AmountNumber", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtCostCentres
              epl.EnableControl("CostCentre", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers
              epl.EnableControl("OrganisationNumber", False)
            Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources
              epl.EnableControl("AppealResources", False)
            Case CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppealActions
              Dim vActionNumber As Integer = IntegerValue(dgr.GetValue(pRow, "ActionNumber"))
              If vActionNumber > 0 Then mvActionNumber = vActionNumber
              If mvSelectedRow > -1 Then
                mvActionNumber = IntegerValue(dgr.GetValue(mvSelectedRow, "ActionNumber"))
                mvActionMenu.ActionNumber = mvActionNumber
                mvActionMenu.ActionStatus = dgr.GetValue(mvSelectedRow, "ActionStatus")
                mvActionMenu.MasterActionNumber = IntegerValue(dgr.GetValue(mvSelectedRow, "MasterAction"))
              End If
              If vActionNumber > 0 Then
                If (ActionRights(vActionNumber) And DataHelper.DocumentAccessRights.darDelete) = DataHelper.DocumentAccessRights.darDelete Then cmdDelete.Enabled = True Else cmdDelete.Enabled = False
                If (ActionRights(vActionNumber) And DataHelper.DocumentAccessRights.darEdit) = DataHelper.DocumentAccessRights.darEdit Then
                  cmdSave.Enabled = True
                  cmdLink1.Enabled = True
                  cmdLink2.Enabled = True
                Else
                  cmdSave.Enabled = False
                  cmdLink1.Enabled = False
                  cmdLink2.Enabled = False
                End If
              End If
          End Select
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS, _
                 CareServices.XMLMaintenanceControlTypes.xmctH2HCollectionPIS
              Dim vControl As Control = FindControl(epl, "PisNumber")
              Dim vComboBox As ComboBox = TryCast(vControl, ComboBox)
              If vComboBox IsNot Nothing Then
                vComboBox.DataSource = Nothing
                vComboBox.ValueMember = "LookupCode"
                vComboBox.DisplayMember = "LookupDesc"
                vComboBox.Items.Add(New LookupItem(vDataRow.Item("PISNumber").ToString, vDataRow.Item("PISNumber").ToString))
                epl.SetValue("PisNumber", vDataRow.Item("PISNumber").ToString, True)
                epl.DataChanged = False
              Else
                DirectCast(vControl, TextLookupBox).ValidationRequired = False
                vControl.Enabled = False
              End If
            Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
              epl.SetDependancies("CollectionPointType")

          End Select
        End If
      End If
      SetButtons(True)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Protected Overrides Sub frmCardMaintenance_Load(ByVal sender As Object, ByVal e As System.EventArgs)
    If Me.DesignMode Then Return
    If Not mvParentForm Is Nothing AndAlso mvParentForm.SizeMaintenanceForm Then
      Location = mvParentForm.Location
      Size = mvParentForm.Size              'Required here for Windows 2000
      mvParentForm.Enabled = False
    Else
      If Not mvParentForm Is Nothing Then mvParentForm.Enabled = False
      If MdiParent Is Nothing AndAlso MDIForm IsNot Nothing Then
        Location = MDIForm.PointToScreen(MDILocation(Width, Height))
      Else
        Location = MDILocation(Width, Height)
      End If
    End If
    mvSelectedRow = -1
    bpl.RepositionButtons()
  End Sub

  Private Sub mvCampaignMenu_ItemSelected(ByVal pMenuItem As CampaignMenu.CampaignMenuItems) Handles mvCampaignMenu.MenuSelected
    Try
      Select Case pMenuItem
        Case CampaignMenu.CampaignMenuItems.cmiNewAppeal
          sel.AddNode(CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal, mvCampaignItem.Code)

        Case CampaignMenu.CampaignMenuItems.cmiNewSegment
          sel.AddNode(CareServices.XMLCampaignDataSelectionTypes.xcadtSegment, mvCampaignItem.Code)

        Case CampaignMenu.CampaignMenuItems.cmiNewCollection
          sel.AddNode(CareServices.XMLCampaignDataSelectionTypes.xcadtCollection, mvCampaignItem.Code)

        Case CampaignMenu.CampaignMenuItems.cmiSumAppeal
          Dim vList As New ParameterList(True)
          mvCampaignItem.FillParameterList(vList)
          Dim vDataSet As DataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtSumAppeal, vList)
          If vDataSet IsNot Nothing Then
            Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(vDataSet)
            If vDataRow IsNot Nothing Then
              epl.Populate(vDataRow)
              SetTabEditingControls(vDataSet)
              ShowInformationMessage(InformationMessages.ImSumAppealComplete)
            End If
          End If

        Case CampaignMenu.CampaignMenuItems.cmiCalculateIncome
          Dim vList As New ParameterList(True)
          Dim vType As CareServices.XMLCampaignDataSelectionTypes = CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignIncome
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
              vType = CareServices.XMLCampaignDataSelectionTypes.xcadtAppealIncome
            Case CareServices.XMLMaintenanceControlTypes.xmctSegment
              vType = CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentIncome
            Case CareServices.XMLMaintenanceControlTypes.xmctH2HCollection, CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
              vType = CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionIncome
          End Select
          mvCampaignItem.FillParameterList(vList)
          Dim vDataSet As DataSet = DataHelper.GetCampaignData(vType, vList)
          If vDataSet IsNot Nothing Then
            Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(vType, vList))
            If vDataRow IsNot Nothing Then
              epl.Populate(vDataRow)
              SetTabEditingControls(vDataSet)
              epl.SetDependancies("Source")
              epl.DataChanged = False
            End If
          End If

        Case CampaignMenu.CampaignMenuItems.cmiAddCollectionBoxes
          'Display frmApplicationParameters then process the resulting ParameterList containing the results
          Dim vList As New ParameterList(True)
          Dim vDefaults As New ParameterList
          mvCampaignItem.FillParameterList(vList)
          Dim vDataRow As DataRow = DataHelper.GetCampaignItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollection, vList)
          If vDataRow IsNot Nothing Then vDefaults.Add("CollectionDesc", vDataRow.Item("CollectionDesc").ToString)
          vList = New ParameterList     'Clear the current List
          vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptAddCollectionBoxes, vDefaults)
          If vList IsNot Nothing AndAlso vList.Count > 0 Then
            'Add the data
            vList("CollectionNumber") = mvCampaignItem.CollectionNumber.ToString
            If mvCampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection Then
              'Manned Collection
              DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes, vList)
              sel.SetSelectionType(CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectionBoxes)
            Else
              'Unmanned Collection
              DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes, vList)
              sel.SetSelectionType(CareServices.XMLCampaignDataSelectionTypes.xcadtUnMannedCollectionBoxes)
            End If
          End If

        Case CampaignMenu.CampaignMenuItems.cmiCountCollectors
          Dim vList As New ParameterList(True)
          mvCampaignItem.FillParameterList(vList)
          Dim vDataSet As DataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectorsCount, vList)
          If vDataSet IsNot Nothing Then
            Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(vDataSet)
            If vDataRow IsNot Nothing Then
              epl.Populate(vDataRow)
              epl.SetDependancies("Source")
              SetTabEditingControls(vDataSet)
              epl.DataChanged = False
            End If
          End If

        Case CampaignMenu.CampaignMenuItems.cmiCopyData
          Clipboard.Clear()
          Clipboard.SetData(GetType(CampaignCopyInfo).FullName, GetCampaignCopyInfo())

        Case CampaignMenu.CampaignMenuItems.cmiCopyCriteria
          Clipboard.Clear()
          Clipboard.SetData(GetType(CampaignCopyInfo).FullName, GetCampaignCopyInfo(True))

        Case CampaignMenu.CampaignMenuItems.cmiPaste
          If Clipboard.ContainsData(GetType(CampaignCopyInfo).FullName) Then
            PasteItem(DirectCast(Clipboard.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo))
          End If

        Case CampaignMenu.CampaignMenuItems.cmiSegmentCriteria
          Dim vCriteriaSet As Integer = mvCampaignMenu.CriteriaSet
          Dim vAddCriteria As Boolean
          Dim vCriteriaSetDesc As String = ""
          If vCriteriaSet > 0 Then
            Dim vList As New ParameterList(True)
            vList.IntegerValue("CriteriaSet") = vCriteriaSet
            Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtCriteriaSets, vList)
            If vRow Is Nothing Then
              vAddCriteria = True
            Else
              vCriteriaSetDesc = vRow("CriteriaSetDesc").ToString
            End If
          Else
            vAddCriteria = True
          End If
          If vAddCriteria Then
            Dim vList As New ParameterList(True)
            vList("ApplicationName") = AppValues.MailingApplicationCode(CareServices.TaskJobTypes.tjtMailingRun)
            If vCriteriaSet > 0 Then vList.IntegerValue("CriteriaSetNumber") = vCriteriaSet
            vCriteriaSetDesc = String.Format("Segment Criteria {0}", mvCampaignItem.Code)
            vList("CriteriaSetDesc") = vCriteriaSetDesc
            Dim vReturnList As ParameterList = DataHelper.AddCriteriaSet(vList)
            vList = New ParameterList(True)
            vList("Campaign") = mvCampaignItem.Campaign
            vList("Appeal") = mvCampaignItem.Appeal
            vList("Segment") = mvCampaignItem.Segment
            mvCampaignMenu.CriteriaSet = vReturnList.IntegerValue("CriteriaSetNumber")
            vCriteriaSet = vReturnList.IntegerValue("CriteriaSetNumber")
            vList("CriteriaSet") = vReturnList("CriteriaSetNumber")
            DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSegment, vList)
          End If
          mvMailingInfo = New MailingInfo
          mvMailingInfo.Init(mvCampaignItem.ApplicationName, vCriteriaSet, True)
          mvMailingInfo.TaskType = CareNetServices.TaskJobTypes.tjtMailingRun
          mvFrmEditCriteria = New frmEditCriteria(mvMailingInfo, "Campaign Selection Set")
          mvFrmEditCriteria.ShowDialog()
          RefreshCard()

        Case CampaignMenu.CampaignMenuItems.cmiSegmentSteps
          Dim vCriteriaSet As Integer = mvCampaignMenu.CriteriaSet
          Dim vAddCriteria As Boolean
          Dim vCriteriaSetDesc As String = ""
          If vCriteriaSet > 0 Then
            Dim vList As New ParameterList(True)
            vList.IntegerValue("CriteriaSet") = vCriteriaSet
            Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtCriteriaSets, vList)
            If vRow Is Nothing Then
              vAddCriteria = True
            Else
              vCriteriaSetDesc = vRow("CriteriaSetDesc").ToString
            End If
          Else
            vAddCriteria = True
          End If
          If vAddCriteria Then
            Dim vList As New ParameterList(True)
            vList("ApplicationName") = AppValues.MailingApplicationCode(CareServices.TaskJobTypes.tjtMailingRun)
            If vCriteriaSet > 0 Then vList.IntegerValue("CriteriaSetNumber") = vCriteriaSet
            vCriteriaSetDesc = String.Format("Segment Criteria {0}", mvCampaignItem.Code)
            vList("CriteriaSetDesc") = vCriteriaSetDesc
            Dim vReturnList As ParameterList = DataHelper.AddCriteriaSet(vList)
            vList = New ParameterList(True)
            vList("Campaign") = mvCampaignItem.Campaign
            vList("Appeal") = mvCampaignItem.Appeal
            vList("Segment") = mvCampaignItem.Segment
            mvCampaignMenu.CriteriaSet = vReturnList.IntegerValue("CriteriaSetNumber")
            vList("CriteriaSet") = vReturnList("CriteriaSetNumber")
            DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSegment, vList)
          End If
          Dim vLM As New ListManager(mvCampaignMenu.CriteriaSet, vCriteriaSetDesc)
          vLM.ShowDialog()
          RefreshCard()

        Case CampaignMenu.CampaignMenuItems.cmiActions
          Dim vMasterAction As Integer = mvCampaignItem.AppealActionNumber
          If vMasterAction = 0 Then
            vMasterAction = FormHelper.NewActionFromTemplate(Me, 0)
            If vMasterAction > 0 Then
              'Update the Appeal
              Dim vList As New ParameterList(True)
              mvCampaignItem.FillParameterList(vList)
              vList("MasterAction") = vMasterAction.ToString
              mvReturnList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctAppeal, vList)
              mvCampaignItem.AppealActionNumber = vMasterAction
            End If
          End If
          If vMasterAction > 0 Then FormHelper.EditAction(vMasterAction, Me)

        Case CampaignMenu.CampaignMenuItems.cmiCountAppealOrSegment, CampaignMenu.CampaignMenuItems.cmiMailAppeal, CampaignMenu.CampaignMenuItems.cmiMailAppealParams 'refresh appeal or segment after count/mail
          Dim vParamList As New ParameterList(True)
          If DoCountOrMail(vParamList, pMenuItem) Then
            Dim vList As New ParameterList(True)
            mvCampaignItem.FillParameterList(vList)
            Dim vDataSet As DataSet = DataHelper.GetCampaignData(mvCampaignDataType, vList)
            If vDataSet IsNot Nothing Then
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vDataSet)
              If vRow IsNot Nothing Then
                epl.Populate(vRow)
                epl.DataChanged = False
              End If
              SetTabEditingControls(vDataSet)
            End If
          End If
      End Select

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enVarInMultipleAreas
          ShowInformationMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enVariableNameContainsInvalidCharacters
          ShowErrorMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try

  End Sub

  Private Function GetCampaignCopyInfo() As CampaignCopyInfo
    Return GetCampaignCopyInfo(False)
  End Function

  Private Function GetCampaignCopyInfo(ByVal pCopySegmentCriteria As Boolean) As CampaignCopyInfo
    Dim vAppealInfo As CampaignCopyInfo = Nothing
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
        vAppealInfo = New CampaignCopyInfo(mvCampaignItem, epl.GetValue("AppealDesc"), Nothing, Nothing, Nothing)
      Case CareServices.XMLMaintenanceControlTypes.xmctH2HCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
           CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
        vAppealInfo = New CampaignCopyInfo(mvCampaignItem, Nothing, epl.GetValue("Collection"), epl.GetValue("CollectionDesc"), Nothing)
      Case CareServices.XMLMaintenanceControlTypes.xmctSegment
        vAppealInfo = New CampaignCopyInfo(mvCampaignItem, Nothing, Nothing, Nothing, epl.GetValue("SegmentDesc"), pCopySegmentCriteria)
    End Select
    Return vAppealInfo
  End Function

  Private Sub PasteItem(CampaignCopyInfo As CampaignCopyInfo)
    Try
      Dim vList As ParameterList = New ParameterList(True)
      Dim vAppealInfo As CampaignCopyInfo = CampaignCopyInfo
      Dim vCopyAppeal As Boolean = False
      Dim vCopySegment As Boolean = False
      Dim vCopyCollection As Boolean = False
      Dim vNewCampaignItem As CampaignItem = Nothing
      Dim vProcess As Boolean

      vAppealInfo.FillParameterList(vList)
      vList("Campaign2") = mvCampaignItem.Campaign
      vList("CopyTickBoxes") = "N"
      vList("CopyMailingCode") = "N"
      vList("CopySourceCode") = "N"
      Dim vCampaignItem As CampaignItem = New CampaignItem(vList("Campaign2"), Nothing, Nothing)

      If (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaign _
         Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal) _
         AndAlso vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctAppeal Then
        vList("Appeal2") = vAppealInfo.Appeal
        vList("AppealDesc") = vAppealInfo.AppealDesc
        If String.Equals(vList("Appeal").ToString, vList("Appeal2").ToString) Then
          Dim vParamList As New ParameterList
          vParamList("Appeal2") = vList("Appeal2")
          vParamList("AppealDesc") = vList("AppealDesc")
          vParamList("CopyTickBoxes") = vList("CopyTickBoxes")
          vParamList("CopyMailingCode") = vList("CopyMailingCode")
          vParamList("CopySourceCode") = vList("CopySourceCode")
          vParamList("Campaign2") = vList("Campaign2")
          vParamList("CopyCampaignCode") = CampaignCopyInfo.Campaign
          Dim vReturnParamList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCopyAppeal, vParamList)
          If vReturnParamList IsNot Nothing AndAlso vReturnParamList.Count > 0 Then
            vList("Appeal2") = vReturnParamList("Appeal2")
            vList("AppealDesc") = vReturnParamList("AppealDesc")
            vList("CopyTickBoxes") = vReturnParamList("CopyTickBoxes")
            vList("CopyMailingCode") = vReturnParamList("CopyMailingCode")
            vList("CopySourceCode") = vReturnParamList("CopySourceCode")
            vProcess = True
          End If
        Else
          vProcess = True
        End If
        If vProcess Then
          DataHelper.CopyCampaignData(vList)
          vNewCampaignItem = New CampaignItem(vList("Campaign2") & "_" & vList("Appeal2") & "_" & vAppealInfo.AppealType, Nothing, Nothing)
          sel.Init(vCampaignItem, vNewCampaignItem)
        End If
      ElseIf ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollection _
             Or (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atH2HCollection)) _
             AndAlso vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctH2HCollection) _
                                                   OrElse _
             ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection _
             Or (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection)) _
             AndAlso vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctMannedCollection) _
                                                   OrElse _
             ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection _
             Or (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atUnMannedCollection)) _
             AndAlso vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctUnMannedCollection) Then
        'copy collection
        vList("Appeal2") = mvCampaignItem.Appeal
        vList("Collection2") = vAppealInfo.Collection
        vList("CollectionDesc") = vAppealInfo.CollectionDesc
        Dim vDataSet As DataSet = DataHelper.CopyCampaignData(vList)
        vNewCampaignItem = New CampaignItem(vList("Campaign2") & "_" & vList("Appeal2") & "_" & vAppealInfo.AppealType & "_" & vDataSet.Tables("Result").Rows(0).Item("CollectionNumber").ToString, Nothing, Nothing)
        sel.UpdateAppealData(vNewCampaignItem)
      ElseIf (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctSegment _
           Or (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppeal And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atSegment)) _
           AndAlso (vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctSegment OrElse vAppealInfo.CampaignCopyType = CDBNET.CampaignCopyInfo.CampaignCopyTypes.cctSegmentCriteria) Then
        If vAppealInfo.CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctSegment Then
          'copy segment
          vList("Appeal2") = mvCampaignItem.Appeal
          vList("Segment2") = vAppealInfo.Segment
          vList("SegmentDesc") = vAppealInfo.SegmentDesc
          If String.Equals(vList("Campaign").ToString, vList("Campaign2").ToString) Then
            Dim vParamList As New ParameterList
            vParamList("Segment2") = vList("Segment2")
            vParamList("SegmentDesc") = vList("SegmentDesc")
            vParamList("CopyTickBoxes") = vList("CopyTickBoxes")
            vParamList("CopyMailingCode") = vList("CopyMailingCode")
            vParamList("CopySourceCode") = vList("CopySourceCode")
            vParamList("CopyAppeal") = vAppealInfo.Appeal
            vParamList("Appeal2") = vList("Appeal2")
            Dim vReturnParamList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCopySegment, vParamList)
            If vReturnParamList IsNot Nothing AndAlso vReturnParamList.Count > 0 Then
              vList("Segment2") = vReturnParamList("Segment2")
              vList("SegmentDesc") = vReturnParamList("SegmentDesc")
              vList("CopyTickBoxes") = vReturnParamList("CopyTickBoxes")
              vList("CopyMailingCode") = vReturnParamList("CopyMailingCode")
              vList("CopySourceCode") = vReturnParamList("CopySourceCode")
              vProcess = True
            End If
          Else
            vProcess = True
          End If
          If vProcess Then
            DataHelper.CopyCampaignData(vList)
            vNewCampaignItem = New CampaignItem(vList("Campaign2") & "_" & vList("Appeal2") & "_" & vAppealInfo.AppealType & "_" & vList("Segment2"), Nothing, Nothing)
            sel.UpdateAppealData(vNewCampaignItem)
          End If
        ElseIf vAppealInfo.CampaignCopyType = CDBNET.CampaignCopyInfo.CampaignCopyTypes.cctSegmentCriteria Then
          'Copy Segment Criteria
          Dim vContinue As Boolean = True
          'Validate that we aren't about to create Criteria with Criteria Set Details and Criteria Selection Steps as we can't have both
          If vAppealInfo.HasCriteriaSetDetails AndAlso mvCampaignItem.HasCriteriaSelectionSteps Then
            'Cannot copy criteria with selection criteria onto criteria with selection steps 
            vContinue = False
            ShowInformationMessage(InformationMessages.ImCannotCopyCampaignCriteriaSetSCSS)
          ElseIf vAppealInfo.HasCriteriaSelectionSteps AndAlso mvCampaignItem.HasCriteriaSetDetails Then
            'Cannot copy criteria with selection steps onto criteria with selection criteria
            vContinue = False
            ShowInformationMessage(InformationMessages.ImCannotCopyCampaignCriteriaSetSSSC)
          End If
          If vContinue Then
            'Copy the Segment criteria set
            vList("Appeal2") = mvCampaignItem.Appeal
            vList("Segment2") = mvCampaignItem.Segment
            Dim vParamList As New ParameterList
            vParamList("Campaign") = vList("Campaign")
            vParamList("Appeal") = vList("Appeal")
            vParamList("Segment") = vList("Segment")
            vParamList("Campaign2") = vList("Campaign2")
            vParamList("Appeal2") = vList("Appeal2")
            vParamList("Segment2") = vList("Segment2")
            Dim vReturnParamList As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptCopySegmentCriteria, vParamList)
            If vReturnParamList IsNot Nothing AndAlso vReturnParamList.Count > 0 Then
              vProcess = True
            End If
            If vProcess Then
              Dim vResult As DataSet = DataHelper.CopyCampaignData(vList)
              If vResult.Tables IsNot Nothing AndAlso vResult.Tables(0).Rows.Count > 0 AndAlso vResult.Tables(0).Columns.Contains("CriteriaSet") AndAlso IntegerValue(vResult.Tables(0).Rows(0).Item("CriteriaSet").ToString) > 0 Then
                ShowInformationMessage(InformationMessages.ImCampaignSegmentCriteriaCopied)  'Segment Criteria copied successfully
              End If
              RefreshCard()
            End If
          End If
        End If
      End If
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        Throw vEx
      End If
    End Try
  End Sub

  Private Sub mvFrmEditCriteria_ProcessMailingCriteriaWithOptional(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As CDBNETCL.ParameterList, ByRef pSuccess As Boolean) Handles mvFrmEditCriteria.ProcessMailingCriteriaWithOptional
    mvMailingInfo.ProcessMailingCriteriaWithOptional(pMailingSelection, pCriteriaSet, pProcessVariables, pEditSegmentCriteria, pList, pSuccess)
  End Sub

  Private Sub mvFrmEditCriteria_ProcessSelection(ByVal pRunPhase As String, ByVal pList As ParameterList) Handles mvFrmEditCriteria.ProcessSelection
    Try
      mvMailingInfo.GenerateStatus = MailingInfo.MailingGenerateResult.mgrNone
      If mvMailingInfo.CriteriaRows = 0 Then
        ShowInformationMessage(InformationMessages.ImNoCriteria)
      ElseIf mvMailingInfo.CriteriaRows > 0 Then
        If pList Is Nothing Then pList = New ParameterList(True)
        pList.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
        pList.IntegerValue("Revision") = mvMailingInfo.Revision
        pList("ApplicationName") = mvMailingInfo.MailingTypeCode
        pList("RunPhase") = pRunPhase
        pList.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        pList.IntegerValue("ExclusionCriteria") = mvMailingInfo.ExclusionCriteriaSet
        pList("UseStandardExclusions") = CStr(IIf(mvMailingInfo.ExclusionCriteriaSet > 0, "Y", "N"))
        pList("OrgMailTo") = mvMailingInfo.OrganisationMailTo
        pList("OrgMailWhere") = mvMailingInfo.OrganisationMailWhere
        pList("OrgLabelName") = mvMailingInfo.OrganisationLabelName
        pList("OrgAddressUsage") = mvMailingInfo.OrganisationAddressUsage
        pList("OrgRoles") = mvMailingInfo.OrganisationRoles
        pList("OrgIncludeHistoricRoles") = CStr(IIf(mvMailingInfo.IncludeHistoricRoles = True, "Y", "N"))
        pList("GeneralMailing") = "Y"
        Dim vResults As ParameterList = DataHelper.ProcessMailingCount(pList)
        If vResults.Contains("MailingCount") Then mvMailingInfo.SelectionCount = vResults.IntegerValue("MailingCount")
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ProcessNew()
    dgr.SelectRow(-1)
    mvSelectedRow = -1
    epl.Clear()
    SetDefaults(False)
    SetCommandsForNew()
  End Sub
  Private Sub SetSourceFields()
    Dim vEnableFields As Boolean = True
    With epl
      If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then vEnableFields = False
      If Not mvEditing Then
        .PanelInfo.PanelItems("Source").Mandatory = False           'New segment source can be generated on server
        .SetValue("IncentiveTriggerLevel", "0.00")
        If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).TYL IsNot Nothing Then
          epl.SetValue("ThankYouLetter", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).TYL)
          'If epl.GetValue("ThankYouLetter").Length > 0 Then vEnableFields = False
        End If
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctSegment
            '
          Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
               CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
               CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
            If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Source IsNot Nothing Then epl.SetValue("Source", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Source)
            If epl.GetValue("Source").Length > 0 Then
              vEnableFields = False
            Else
              If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) = False Then
                .PanelInfo.PanelItems("Source").Mandatory = True
              End If
            End If
        End Select
      Else
        If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Source IsNot Nothing AndAlso sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Source.Length > 0 Then
          'disable the fields
          vEnableFields = False
        End If
      End If
      .EnableControl("Source", vEnableFields)
      '.EnableControl("DistributionCode", vEnableFields)
      '.EnableControl("IncentiveScheme", vEnableFields)
      '.EnableControl("IncentiveTriggerLevel", vEnableFields)
      '.EnableControl("DiscountPercentage", vEnableFields)
      If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_analysis_from_source) Then
        'If config set then enable ThankYouLetter control if it is blank
        If epl.GetValue("ThankYouLetter").Length = 0 Then vEnableFields = True
      End If
      '.EnableControl("ThankYouLetter", vEnableFields)
      epl.SetDependancies("Source")
    End With
  End Sub

  Private Sub SetTabEditingControls(ByVal pDataSet As DataSet)
    With epl
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
          .EnableControl("Campaign", False)
          mvCampaignBusinessType = .GetValue("BusinessType")
        Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
          .EnableControl("Appeal", False)
          mvMailingType = .GetValue("MailingType")
          .SetDependancies("AppealType")
          EPL_ValueChanged(epl, "AppealType", epl.GetValue("AppealType"))
          If sel.HasDependants Then .EnableControl("MailingType", False)
          If mvCampaignItem.AppealType = CampaignItem.AppealTypes.atSegment Then
            .SetDependancies("MailingType")
            .SetDependancies("RequiredCount")
            .SetDependancies("MailJoints")
            .EnableControl("DefaultDespatchQuantity", (.GetValue("MailingType") = "SR"))
          Else
            mvCampaignItem.SetIncomeFields(.GetValue("Source"), .GetValue("Product"), .GetValue("Rate"), .GetValue("BankAccount"), .GetValue("ThankYouLetter"))
            If sel.HasDependants Then
              ' if there are any collections under the appeal, and the follwing values are set then they should be allowed to be changed
              If .GetValue("Source").Length > 0 Then .EnableControl("Source", False)
              If .GetValue("Product").Length > 0 Then .EnableControl("Product", False)
              If .GetValue("Rate").Length > 0 Then .EnableControl("Rate", False)
              If .GetValue("BankAccount").Length > 0 Then .EnableControl("BankAccount", False)
            End If
            .EnableControl("DefaultDespatchQuantity", False)
          End If
          mvCampaignItem.AppealActionNumber = IntegerValue(pDataSet.Tables("DataRow").Rows(0).Item("MasterAction").ToString)
          .EnableControlList("OrgMailTo,OrgMailWhere,OrgMailAddrUsage", False)

          epl.DataChanged = False
          'save the appeals Income Fields in the campaign item

        Case CareServices.XMLMaintenanceControlTypes.xmctSegment
          mvCampaignMenu.CriteriaSet = IntegerValue(pDataSet.Tables("DataRow").Rows(0).Item("CriteriaSet").ToString)
          If mvCampaignMenu.CriteriaSet > 0 Then
            mvCampaignItem.CriteriaSet = mvCampaignMenu.CriteriaSet
          End If

          .EnableControl("Segment", False)
          .SetDependancies("RequiredCount")
          .EnableControl("DespatchQuantity", (mvCampaignItem.AppealMailingTypeCode = "SR"))
          .EnableControl("ActualCount", (.GetValue("Direction") = "I"))
          epl.DataChanged = False
          SetSourceFields()
          epl.DataChanged = False
        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
          epl.DataChanged = False
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Product IsNot Nothing Then
            epl.EnableControl("Product", False)
          End If
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Rate IsNot Nothing Then
            epl.EnableControl("Rate", False)
          End If
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).BankAccount IsNot Nothing Then
            epl.EnableControl("BankAccount", False)
          End If

          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
              'save the collection's resources produced on field in the campaign item
              mvCampaignItem.ResourcesProducedOn = pDataSet.Tables("DataRow").Rows(0).Item("ResourcesProducedOn").ToString
            Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection
              'save the collection's resources produced on field in the campaign item
              mvCampaignItem.SetCollectionTimes(pDataSet.Tables("DataRow").Rows(0).Item("StartTime").ToString, pDataSet.Tables("DataRow").Rows(0).Item("EndTime").ToString)
          End Select
          mvCampaignItem.CollectionBankAccount = epl.GetValue("BankAccount")
          SetSourceFields()
        Case CareServices.XMLMaintenanceControlTypes.xmctAppealResources
          If dgr.DataRowCount > 0 Then
            Dim vProduct As String = dgr.GetValue(dgr.CurrentRow, "Product")
            If vProduct.Length > 0 Then mvCampaignItem.FirstAppealResource = vProduct
          End If

        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes
          Dim vValue As Double = DoubleValue(dgr.GetValue(mvSelectedRow, "SumPayments"))
          If vValue > 0 Then
            'Disable fields
            epl.EnableControlList("CollectorNumber,Amount,CollectionPisNumber", False)
          End If
      End Select
    End With
  End Sub

  Private Sub SetTabNotEditingControls()
    With epl
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
          .SetValue("StartDate", AppValues.TodaysDate)
          .SetValue("EndDate", AppValues.TodaysDateAddYears(1))
        Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
          .SetValue("AppealDate", AppValues.TodaysDate)
          .SetValue("BusinessType", mvCampaignBusinessType)
          .SetValue("CreateMailingHistory", "Y")
          .SetValue("EndDate", AppValues.TodaysDate)
          .SetValue("BypassCount", "Y")
          'only allow appeal types where the user has create permissions
          Dim vTable As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtAppealTypes)
          vTable.DefaultView.RowFilter = "Access = 'C'"
          epl.SetComboDataSource("AppealType", "AppealType", "AppealTypeDesc", vTable, False)
          .SetValue("AppealType", AppValues.DefaultAppealType)
          .EnableControl("DefaultDespatchQuantity", False)

          .EnableControlList("OrgMailTo,OrgMailWhere,OrgMailAddrUsage", False)

        Case CareServices.XMLMaintenanceControlTypes.xmctSegment
          .SetValue("Direction", "O")
          .SetValue("Random", "N", True)
          .SetValue("Score", "", True)
          .SetValue("SegmentSequence", CStr(sel.MaxSequenceNumber + 1))
          SetSourceFields()
          .EnableControl("DespatchQuantity", (mvCampaignItem.AppealMailingTypeCode = "SR"))
        Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
             CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
          .SetValue("CopyCriteria", AppValues.DefaultCopyCriteria)
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Product IsNot Nothing Then
            epl.SetValue("Product", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Product)
            If epl.GetValue("Product").Length > 0 Then .EnableControl("Product", False)
          End If
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Rate IsNot Nothing Then
            epl.SetValue("Rate", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).Rate)
            If epl.GetValue("Rate").Length > 0 Then .EnableControl("Rate", False)
          End If
          If sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).BankAccount IsNot Nothing Then
            epl.SetValue("BankAccount", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).BankAccount)
            If epl.GetValue("BankAccount").Length > 0 Then .EnableControl("BankAccount", False)
          End If
          SetSourceFields()
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctH2HCollection _
          Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection Then
            epl.SetValue("StartDate", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).StartDate)
            epl.SetValue("EndDate", sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citAppeal).EndDate)
          End If
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection Then
            .SetValue("StartTime", AppValues.DefaultStartTime)
            .SetValue("EndTime", AppValues.DefaultEndTime)
          End If
      End Select
    End With
    'done this as a workaround for now. need to move all defaulting code from here to set defaults and then alwyas call processnew
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors Then
      ProcessNew()
    Else
      SetCommandsForNew()
    End If
  End Sub
  Protected Overrides Sub cmdLink_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    'This is the Boxes button for Paying In Slips
    If dgr.DataRowCount > 0 AndAlso mvSelectedRow >= 0 Then
      'Need this at the moment because clicking New does not disable Boxes button
      Try
        Dim vDataSet As New DataSet
        cmdClose.Visible = True
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS
            Dim vList As New ParameterList(True)
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
            AddHandler epl.PopulateListBox, AddressOf epl_PopulateListBox
            mvCampaignItem.AdditionalValues("CollectionPISNumber") = dgr.GetValue(mvSelectedRow, "CollectionPISNumber")
            cmdSave.Enabled = True
            If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
              epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))
            End If
            epl.Visible = True
            GetAdditionalKeyValues(vList)
            vDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtContactCollectionBoxes, vList)
            epl.Refresh()
          Case CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions
            Dim vList As New ParameterList(True)
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
            mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtPoints
            mvCampaignItem.AdditionalValues("CollectionRegionNumber") = dgr.GetValue(mvSelectedRow, "CollectionRegionNumber")
            mvCampaignItem.AdditionalValues("GeographicalRegionDesc") = dgr.GetValue(mvSelectedRow, "GeographicalRegionDesc")
            cmdSave.Enabled = True

            If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
              epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))
              epl.Refresh()
              epl.SetValue("GeographicalRegionDesc", mvCampaignItem.AdditionalValues("GeographicalRegionDesc"))
            End If
            epl.Visible = True
            GetAdditionalKeyValues(vList)
            vDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtPoints, vList)

          Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors
            Dim vList As New ParameterList(True)
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
            mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtCollectorShifts
            mvCampaignItem.AdditionalValues("CollectorNumber") = dgr.GetValue(mvSelectedRow, "CollectorNumber")
            mvCampaignItem.AdditionalValues("ContactName") = dgr.GetValue(mvSelectedRow, "ContactName")
            mvCampaignItem.AdditionalValues("TotalTime") = dgr.GetValue(mvSelectedRow, "TotalTime")
            mvCampaignItem.AdditionalValues("CollectionStartTime") = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).StartTime
            mvCampaignItem.AdditionalValues("CollectionEndTime") = sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCollection).EndTime
            cmdSave.Enabled = True

            If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
              epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))
              epl.Refresh()
            End If
            epl.Visible = True
            GetAdditionalKeyValues(vList)
            'Fill the regions combo
            Dim vTextLookup As TextLookupBox = epl.FindTextLookupBox("GeographicalRegion")
            vTextLookup.FillComboWithRestriction(mvCampaignItem.CollectionNumber.ToString)
            If vTextLookup.GetDataRow IsNot Nothing Then vTextLookup.SetFilter("Collectionregionnumber <> ''") 'Only set filter if we have some data
            vDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectorShifts, vList)

          Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudget
            Dim vList As New ParameterList(True)
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
            mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgetDetails
            mvCampaignItem.AdditionalValues("AppealBudgetNumber") = dgr.GetValue(mvSelectedRow, "AppealBudgetNumber")
            mvCampaignItem.AdditionalValues("BudgetPeriod") = dgr.GetValue(mvSelectedRow, "BudgetPeriod")
            cmdSave.Enabled = True

            If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
              epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))
              epl.Refresh()
            End If
            epl.Visible = True
            'Populate the Segment ComboBox
            vList("Campaign") = mvCampaignItem.Campaign
            vList("Appeal") = mvCampaignItem.Appeal
            epl.SetComboDataSource("Segment", "Segment", "SegmentDesc", DataHelper.GetTableFromDataSet(DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftCampaignSegments, vList)))
            'Select the data
            vList = New ParameterList(True)
            GetAdditionalKeyValues(vList)
            vDataSet = DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgetDetails, vList)

          Case CareServices.XMLMaintenanceControlTypes.xmctAction
            mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtNone
            If sender Is cmdLink1 Then
              mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink
              vDataSet = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, mvActionNumber)
            Else
              mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic
              vDataSet = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvActionNumber)
            End If
            RefreshCard()

        End Select
        ShowGrid()
        dgr.Populate(vDataSet)
        mvEditing = False
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
            If dgr.DataRowCount > 0 Then SelectRow(0)
          Case CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS
            mvEditing = (dgr.DataRowCount > 0)
          Case CareServices.XMLMaintenanceControlTypes.xmctCollectorShifts
            If dgr.DataRowCount > 0 Then SelectRow(0)
          Case CareServices.XMLMaintenanceControlTypes.xmctAppealBudgetDetails
            mvEditing = (dgr.DataRowCount > 0)
            If dgr.DataRowCount > 0 Then
              SelectRow(0)
            Else
              ProcessNew()
            End If
          Case CareServices.XMLMaintenanceControlTypes.xmctActionLink
            If dgr.DataRowCount > 0 Then cmdDelete.Enabled = True
          Case CareServices.XMLMaintenanceControlTypes.xmctActionTopic
            If dgr.DataRowCount > 0 Then cmdDelete.Enabled = True
        End Select
        cmdOther.Visible = False
        cmdDelete.Visible = CanDelete(mvMaintenanceType)
        SetButtons(True)

        If ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink Or _
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic) And _
          dgr.DataRowCount > 0) Then
          cmdDelete.Enabled = True
        End If

      Catch vException As Exception
        DataHelper.HandleException(vException)
      End Try
    End If
  End Sub
  Private Sub epl_PopulateListBox(ByVal sender As Object, ByVal pLB As ListBox, ByVal pPanelItem As PanelItem)
    RePopulateListBox(pLB)
  End Sub

  Private Sub RePopulateListBox(ByVal pLB As ListBox)
    Dim vList As New ParameterList(True)

    pLB.DataSource = Nothing
    pLB.DisplayMember = "BoxReference"
    pLB.ValueMember = "CollectionBoxNumber"
    vList.IntegerValue("CollectionNumber") = mvCampaignItem.CollectionNumber
    DataHelper.FillListBox(pLB, CareNetServices.XMLLookupDataTypes.xldtCollectionBoxReferences, False, vList)
  End Sub

  Private Sub EPL_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    ' these bits are done here as this form knows whether we are editing or not and the edit panel does not.
    Select Case pParameterName
      Case "AppealType"
        If Not mvEditing Then
          Select Case pValue
            Case "S"
              'if creating a new segmented appeal then set the mailing type to blank and enable it
              epl.SetValue("MailingType", "")
              epl.EnableControl("MailingType", True)
            Case Else
              'if creating a new non-segmented appeal then set the mailing type to 'MKTG' and disable it
              epl.SetValue("MailingType", "MKTG", True)
              If FindControl(epl, "BankAccount", False) IsNot Nothing Then
                'if creating a new non-segmented appeal then set the bank account from the default bank account control values
                Select Case pValue
                  Case "H"
                    epl.SetValue("BankAccount", AppValues.ControlValue(AppValues.ControlValues.default_h2h_bank_account))
                  Case "M"
                    epl.SetValue("BankAccount", AppValues.ControlValue(AppValues.ControlValues.default_manned_bank_account))
                  Case "U"
                    epl.SetValue("BankAccount", AppValues.ControlValue(AppValues.ControlValues.default_unmanned_bank_account))
                End Select
              End If
          End Select
        Else
          If pValue <> "S" Then
            epl.SetValue("MailingType", "MKTG", True)
          End If
        End If

      Case "AddressNumber", "OrganisationNumber"
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints Or _
               mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection Then
          SetFromOrganisationAndAddress()
        End If
      Case "OrgMailRoles"
        Dim vListBox As ListBox = DirectCast(pSender, ListBox)
        If vListBox.Items.Count > 0 AndAlso vListBox.Items(0).ToString.StartsWith("'") Then
          Dim vTable As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtRoles)
          For vIndex As Integer = 0 To vListBox.Items.Count - 1
            Dim vRows() As DataRow = vTable.Select("Role = " & vListBox.Items(vIndex).ToString)
            If vRows.Length = 1 Then
              vListBox.Items(vIndex) = vRows(0).Item("RoleDesc")
            End If
          Next
        End If
    End Select
  End Sub

  Private Sub SetFromOrganisationAndAddress()
    Dim vTextLookup As TextLookupBox = DirectCast(FindControl(Me, "OrganisationNumber", False), TextLookupBox)
    Dim vOrgName As String = ""
    Dim vAddressTown As String = ""
    Dim vDesc As String = ""
    If vTextLookup IsNot Nothing Then
      vOrgName = vTextLookup.Description
      If epl.GetValue("AddressNumber").Length > 0 Then
        Dim vCombo As ComboBox = epl.FindComboBox("AddressNumber")
        If vCombo IsNot Nothing Then
          Dim vTable As DataTable = DirectCast(vCombo.DataSource, DataTable)
          vAddressTown = vTable.Rows(vCombo.SelectedIndex).Item("Town").ToString()
        End If
      End If
    End If
    If vOrgName.Length > 0 Then
      If vAddressTown.Length > 0 Then
        vDesc = String.Concat(vOrgName, "- ", vAddressTown)
      Else
        vDesc = vOrgName
      End If

      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctCollectionPoints
          If Not mvEditing OrElse epl.GetValue("CollectionPoint").Length = 0 Then
            epl.SetValue("CollectionPoint", vDesc)
          End If
        Case CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection
          If Not mvEditing OrElse epl.GetValue("CollectionDesc").Length = 0 Then
            epl.SetValue("CollectionDesc", vDesc)
          End If
      End Select
    End If
  End Sub

  Private Sub SetButtons(ByVal pShowGrid As Boolean, Optional ByVal pAppealLocked As Boolean = False)
    If pAppealLocked = False Then
      If pShowGrid Then
        cmdDelete.Visible = CanDelete(mvMaintenanceType)
        cmdNew.Visible = (mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes _
                                And mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes _
                                And mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctAssignBoxesToPIS _
                                And mvCampaignDataType <> CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments)
        cmdSave.Visible = (mvCampaignDataType <> CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments)
      Else
        cmdNew.Visible = False
        cmdSave.Visible = True
      End If
      cmdOther.Visible = ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionPIS And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection) _
                          Or (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCollectionRegions And mvCampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection) _
                          Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectors _
                          Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAppealBudget)

      If pShowGrid Then
        If mvEditing Then
          cmdDelete.Enabled = CanDelete(mvMaintenanceType)
          cmdOther.Enabled = True
          cmdSave.Enabled = True
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTickBox Then
            Dim vTB As TextBox = epl.FindTextBox("TickBoxNumber")
            If vTB IsNot Nothing Then
              If IntegerValue(epl.PanelInfo.PanelItems.Item("TickBoxNumber").MaximumValue) > 0 AndAlso (dgr.RowCount >= IntegerValue(epl.PanelInfo.PanelItems.Item("TickBoxNumber").MaximumValue)) Then
                cmdNew.Enabled = False
              Else
                cmdNew.Enabled = True
              End If
            End If
          End If
        Else
          cmdDelete.Enabled = False
          cmdOther.Enabled = False
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollectionBoxes _
            Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollectionBoxes Then
            cmdSave.Enabled = False
          Else
            cmdSave.Enabled = True
          End If
        End If
      Else
        cmdDelete.Enabled = True
        cmdSave.Enabled = True
      End If
    Else
      cmdNew.Visible = False
      cmdSave.Visible = False
      cmdDelete.Visible = False
      cmdOther.Visible = False

    End If
    bpl.RepositionButtons()
  End Sub

  Private Sub RefreshTabSelector(ByVal pList As ParameterList)
    Dim vItem As CampaignItem
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCampaign Then
      vItem = New CampaignItem(pList("Campaign"), pList("StartDate"), pList("EndDate"))
    Else
      vItem = New CampaignItem(pList("Campaign"), sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCampaign).StartDate, sel.GetParentCampaignItem(CampaignItem.CampaignItemTypes.citCampaign).EndDate)
    End If
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctCampaign
        sel.Init(vItem)
        Me.Text = sel.Caption
      Case CareServices.XMLMaintenanceControlTypes.xmctAppeal
        Dim vSelectItem As CampaignItem = New CampaignItem(pList("Campaign") & "_" & pList("Appeal") & "_" & pList("AppealType"), pList("AppealDate"), pList("EndDate"))
        sel.Init(vItem, vSelectItem)
        Me.Text = sel.Caption
      Case CareServices.XMLMaintenanceControlTypes.xmctSegment
        mvCampaignItem = New CampaignItem(pList("Campaign") & "_" & pList("Appeal") & "_" & pList("AppealType") & "_" & pList("Segment"), pList("SegmentDate"), "")
        sel.UpdateAppealData(mvCampaignItem)
      Case CareServices.XMLMaintenanceControlTypes.xmctMannedCollection, _
       CareServices.XMLMaintenanceControlTypes.xmctUnMannedCollection, _
       CareServices.XMLMaintenanceControlTypes.xmctH2HCollection
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMannedCollection Then
          mvCampaignItem = New CampaignItem(pList("Campaign") & "_" & pList("Appeal") & "_" & pList("AppealType") & "_" & mvReturnList("CollectionNumber"), pList("CollectionDate"), "")
        Else
          mvCampaignItem = New CampaignItem(pList("Campaign") & "_" & pList("Appeal") & "_" & pList("AppealType") & "_" & mvReturnList("CollectionNumber"), pList("StartDate"), pList("EndDate"))
        End If
        sel.UpdateAppealData(mvCampaignItem)
    End Select
  End Sub

  Public Function CheckParentDates(ByVal pParentStartDate As String, ByVal pParentEndDate As String, ByVal pChildStartDate As String, ByVal pChildEndDate As String, ByVal pStartDateParam As String, ByVal pEndDateParam As String, ByVal pError As String) As Boolean
    Dim vCheckRequired As Boolean = True
    Dim vParentStartDate As DateTime = New DateHelper(pParentStartDate, DateHelper.DateHelperNullTypes.dhntMinimum).DateValue
    Dim vParentEndDate As DateTime = New DateHelper(pParentEndDate, DateHelper.DateHelperNullTypes.dhntMaximum).DateValue
    Dim vStartDate As DateTime
    If pChildStartDate.Length = 0 Then
      vCheckRequired = False
    Else
      vStartDate = New DateHelper(pChildStartDate, DateHelper.DateHelperNullTypes.dhntMinimum).DateValue
    End If
    Dim vEndDate As DateTime = New DateHelper(pChildEndDate, DateHelper.DateHelperNullTypes.dhntMaximum).DateValue
    Dim vValid As Boolean = True

    epl.SetErrorField(pStartDateParam, "")
    If pEndDateParam IsNot Nothing Then epl.SetErrorField(pEndDateParam, "")
    If vCheckRequired Then
      If vStartDate < vParentStartDate Then
        vValid = False
        epl.SetErrorField(pStartDateParam, pError)
      ElseIf vEndDate > vParentEndDate Then
        vValid = False
        epl.SetErrorField(pEndDateParam, pError)
      End If
    End If
    Return vValid
  End Function

  Private Sub frmCampaignSet_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs)
    If mvCampaignMenu IsNot Nothing Then mvCampaignMenu.Dispose()
    If mvCustomiseMenu IsNot Nothing Then mvCustomiseMenu.Dispose()
  End Sub

  Private Sub dgr_GetPrintParameters(ByVal pSender As Object, ByRef pJobName As String) Handles dgr.GetPrintParameters
    Dim vPrintCaption As New StringBuilder
    If mvCampaignItem IsNot Nothing Then
      vPrintCaption.Append(sel.Caption)
      If mvCampaignItem.ItemType <> CampaignItem.CampaignItemTypes.citCampaign Then
        vPrintCaption.Append(" Appeal: ")
        vPrintCaption.Append(mvCampaignItem.Appeal)
        If mvCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citSegment Then
          vPrintCaption.Append(" Segment: ")
          vPrintCaption.Append(mvCampaignItem.Segment)
        ElseIf mvCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citCollection Then
          vPrintCaption.Append(" Collection: ")
          vPrintCaption.Append(mvCampaignItem.CollectionNumber)
        End If
      End If
      vPrintCaption.Append(" - ")
      vPrintCaption.Append(sel.SelectedNodeText)
      pJobName = vPrintCaption.ToString
    End If
  End Sub

  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub epl_ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Handles epl.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function DoCountOrMail(ByVal pParamList As ParameterList, ByVal pMenuItem As CampaignMenu.CampaignMenuItems) As Boolean
    Dim vCursor As New BusyCursor()
    Dim vParamList As ParameterList = pParamList
    Dim vMenuItem As CampaignMenu.CampaignMenuItems = pMenuItem

    Try
      If Not vParamList.Contains("Campaign") Then
        vParamList.Add("Campaign", mvCampaignItem.Campaign)
        vParamList.Add("Appeal", mvCampaignItem.Appeal)
        vParamList.Add("Segment", mvCampaignItem.Segment)
        If vMenuItem = CampaignMenu.CampaignMenuItems.cmiMailAppeal Or vMenuItem = CampaignMenu.CampaignMenuItems.cmiMailAppealParams Then
          vParamList.Add("Mail", "Y")
        Else
          vParamList.Add("Mail", "N")
        End If

      End If
      vParamList("UseStandardExclusions") = "N"
      If IntegerValue(AppValues.ControlValue(AppValues.ControlTables.marketing_controls, AppValues.ControlValues.criteria_set)) > 0 Then
        Dim vResult As DialogResult = ShowQuestion(QuestionMessages.QmUseStandardExclusions, MessageBoxButtons.YesNoCancel)
        Select Case vResult
          Case System.Windows.Forms.DialogResult.Yes
            vParamList("UseStandardExclusions") = "Y"
          Case System.Windows.Forms.DialogResult.Cancel
            Return False
        End Select
      End If
      Dim vCheckORS As Boolean = False
      Try
        Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignMailOrCount, vParamList))
      Catch vEx As CareException
        Select Case vEx.ErrorNumber
          Case CareException.ErrorNumbers.enSegmentCriteriaMissingCount, _
               CareException.ErrorNumbers.enSegmentCriteriaMissingMailing, _
               CareException.ErrorNumbers.enAppealSegmentsCriteriaMissingCount, _
               CareException.ErrorNumbers.enAppealSegmentsCriteriaMissingMailing
            If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
              Return False
            Else
              vCheckORS = True
            End If
          Case CareException.ErrorNumbers.enSingleSegmentHasMultipleOrs, _
               CareException.ErrorNumbers.enSegmentHasMultipleOrs, _
               CareException.ErrorNumbers.enSegmentsHaveMultipleOrs
            vCheckORS = True
          Case Else
            Throw vEx
        End Select
      End Try
      If vCheckORS Then
        Try
          vParamList("SegmentCriteriaChecked") = "Y"
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCampaignMailOrCount, vParamList))
        Catch vEx As CareException
          Select Case vEx.ErrorNumber
            Case CareException.ErrorNumbers.enSingleSegmentHasMultipleOrs, _
                 CareException.ErrorNumbers.enSegmentHasMultipleOrs, _
                 CareException.ErrorNumbers.enSegmentsHaveMultipleOrs
              If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                Return False
              End If
            Case Else
              Throw vEx
          End Select
        End Try
        'Remove any parameters used for exceptions
        If vParamList.Contains("SegmentCriteriaChecked") Then vParamList.Remove("SegmentCriteriaChecked")
      End If
      vParamList("ApplicationName") = mvCampaignItem.ApplicationName

      Dim vRunMailing As Boolean = True
      If pMenuItem = CampaignMenu.CampaignMenuItems.cmiMailAppealParams Then vRunMailing = False

      Dim vRunMailResult As FormHelper.RunMailingResult
      vRunMailResult = FormHelper.RunMailing(CareServices.TaskJobTypes.tjtMailingRun, vParamList, vRunMailing)
      Select Case vRunMailResult
        Case FormHelper.RunMailingResult.MailingRunAsych
          'Don't unlock the appeal
          Return False
        Case FormHelper.RunMailingResult.NoMailingRun
          'No need to unlock appeal as the server code should do this   mvCampaignItem.AppealLocked = False
          Return False
        Case Else
          'No need to update the actual count date as the server should do this
          Return True
      End Select

    Finally
      vCursor.Dispose()
    End Try
  End Function
  Private Sub dgr_CanCustomise(ByVal pSender As Object, ByVal pRow As String) Handles dgr.CanCustomise
    RefreshCard()
  End Sub
  Private Sub UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel
    epl.ClearDataSources(epl)
    epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing, 0, ""))
    epl.FillDeferredCombos(epl)
    RefreshCard()
  End Sub
  Private Function ActionRights(ByVal pActionNumber As Integer) As DataHelper.DocumentAccessRights

    Dim vDataSet As DataSet = DataHelper.GetActionData(CDBNETCL.CareNetServices.XMLActionDataSelectionTypes.xadtActionInformation, pActionNumber)
    Dim vRights As DataHelper.DocumentAccessRights

    If vDataSet.Tables.Count > 0 Then
      Dim vRow As DataRow = vDataSet.Tables(0).Rows(0)
      vRights = CType(vRow.Item("ActionRights"), DataHelper.DocumentAccessRights)
    End If
    Return vRights
  End Function

  Private Sub mvActionMenu_DeleteAction(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvActionMenu.DeleteAction
    cmdDelete_Click(sender, e)
  End Sub

  Private Sub mvActionMenu_NewAction(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvActionMenu.NewAction
    ProcessNew()
  End Sub

  Private Sub mvActionMenu_NewActionFromTemplate(ByVal sender As Object, ByVal e As System.EventArgs, ByVal pActionNumber As Integer) Handles mvActionMenu.NewActionFromTemplate

    If pActionNumber > 0 Then
      'Update the Appeal
      Dim vList As New ParameterList(True)
      mvCampaignItem.FillParameterList(vList)
      'vList("ActionNumber") = pActionNumber.ToString
      vList("MasterAction") = pActionNumber.ToString
      mvReturnList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctAppeal, vList)
      mvCampaignItem.AppealActionNumber = pActionNumber
      RefreshCard()
    End If

  End Sub

  Private Sub mvActionMenu_RefreshCard(ByVal sender As Object) Handles mvActionMenu.RefreshCard
    RefreshCard()
  End Sub

  Private Sub AfterExpand(ByVal sender As Object) Handles sel.NodeAfterExpand
    If mvCampaignDataType = CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal AndAlso cmdDelete.Visible = False AndAlso sel.HasDependants = False AndAlso mvCampaignItem.Existing Then
      cmdDelete.Visible = True
      bpl.RepositionButtons()
    End If
  End Sub

#Region "Drag and Drop Support"

  Private mvNodeToDrag As TreeNode = Nothing

  Private Sub TreeView_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles sel.TreeMouseDown
    mvNodeToDrag = Nothing
    If e.Button = System.Windows.Forms.MouseButtons.Left AndAlso _
       sender.GetType Is GetType(VistaTreeView) Then
      Dim vNode As TreeNode = DirectCast(sender, VistaTreeView).GetNodeAt(e.X, e.Y)
      If vNode IsNot Nothing AndAlso _
         (CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppeal OrElse _
          CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtSegment) Then
        mvNodeToDrag = vNode
        DirectCast(sender, VistaTreeView).SelectedNode = vNode
      End If
    End If
  End Sub

  Private Sub TreeView_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles sel.TreeMouseMove
    If e.Button = System.Windows.Forms.MouseButtons.Left AndAlso _
       sender.GetType Is GetType(VistaTreeView) AndAlso
       mvNodeToDrag IsNot Nothing Then
      DirectCast(sender, VistaTreeView).AllowDrop = True
      DoDragDrop(GetCampaignCopyInfo(), DragDropEffects.Copy)
    Else
      mvNodeToDrag = Nothing
    End If
  End Sub

  Private Sub TreeView_DragOver(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles sel.TreeDragOver
    If sender.GetType Is GetType(VistaTreeView) AndAlso
      e.Data.GetDataPresent(GetType(CampaignCopyInfo).FullName) Then
      Dim vNode As TreeNode = DirectCast(sender, VistaTreeView).GetNodeAt(DirectCast(sender, VistaTreeView).PointToClient(New Point(e.X, e.Y)))
      If vNode IsNot Nothing AndAlso _
         ((DirectCast(e.Data.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo).CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctAppeal AndAlso
           CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtCampaign) OrElse
          (DirectCast(e.Data.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo).CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctSegment AndAlso
           CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppeal)) Then
        e.Effect = DragDropEffects.Copy
      Else
        e.Effect = DragDropEffects.None
      End If
    End If
  End Sub

  Private Sub TreeView_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles sel.TreeDragDrop
    If sender.GetType Is GetType(VistaTreeView) AndAlso _
       e.Data.GetDataPresent(GetType(CampaignCopyInfo).FullName) Then
      Dim vNode As TreeNode = DirectCast(sender, VistaTreeView).GetNodeAt(DirectCast(sender, VistaTreeView).PointToClient(New Point(e.X, e.Y)))
      If vNode IsNot Nothing AndAlso _
         ((DirectCast(e.Data.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo).CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctAppeal AndAlso
           CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtCampaign) OrElse
          (DirectCast(e.Data.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo).CampaignCopyType = CampaignCopyInfo.CampaignCopyTypes.cctSegment AndAlso
           CType(vNode.Tag, TabSelector.SelectionItem).CampaignSelectionType = CareNetServices.XMLCampaignDataSelectionTypes.xcadtAppeal)) Then
        vNode.TreeView.SelectedNode = vNode
        PasteItem(DirectCast(e.Data.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo))
      End If
    End If
  End Sub

#End Region
End Class