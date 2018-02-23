
Public Class frmEventSet
  Implements IPanelVisibility
  Implements IDashboardTabContainer
  Implements IMainForm

  Private mvEventDataType As CareServices.XMLEventDataSelectionTypes = CType(-1, CareServices.XMLEventDataSelectionTypes)
  Private mvGroupID As String
  Private mvReadOnlyPage As Boolean
  Private mvEditing As Boolean
  Private mvEventInfo As CareEventInfo
  Private mvSessionNumber As Integer = 0

  Private mvStandAlone As Boolean
  Private mvStandAloneParent As Form

  Private mvOptionNumber As Integer = 0           'Used by option sessions to hold booking option information
  Private mvPersonnelInfo As EventPersonnelInfo   'Used by personnel tasks if restricting to a personnel person
  Private mvBookingNumber As Integer = 0          'Used by allocate to identify booking
  Private mvBookingQuantity As Integer            'Used by allocate
  Private mvBookerContact As ContactInfo          'Used by allocate
  Private mvOptionDesc As String
  Private mvNumberOfSessions As Integer
  Private mvPickSessions As Boolean
  Private mvSessionMaximum As Integer
  Private mvSelectedSessionNumbers As ArrayListEx
  Private mvSessionsChanged As Boolean
  Private mvBookingList As ParameterList
  Private mvOptionMaximumBookings As Integer
  Private mvOptionMinimumBookings As Integer
  Private mvDashboard As DashboardTabControl
  Private mvTopicDataSheet As TopicDataSheet
  Private WithEvents mvHeader As FormHeader
  Private mvRoomsTable As DataTable
  Private mvContainsBookings As Boolean
  Private mvBookingPriceChange As Boolean
  Private mvMainMenu As MainMenu
  Private mvPrefix As String = ""
  Private mvActionNumber As Integer
  Private mvDefaultOptionNumber As Integer = 0

  Private WithEvents mvEventMenu As EventMenu
  Private WithEvents mvEventFinancialLinkMenu As EventFinancialLinkMenu
  Private WithEvents mvEventDelegateMenu As EventDelegateMenu
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private WithEvents mvActionMenu As ActionMenu

  Private mvKeyValuesColl As DisplayGridKeyValues

#Region "IPanelVisibility"

  Public Sub SetPanelVisibility() Implements IPanelVisibility.SetPanelVisibility
    If mvHeader IsNot Nothing AndAlso mvEventInfo IsNot Nothing AndAlso mvEventInfo.EventNumber > 0 Then RefreshHeader()
    splBottom.Panel1Collapsed = Not MainHelper.ShowSelectionPanel
  End Sub

  Public Property PanelHasFocus() As Boolean Implements IPanelVisibility.PanelHasFocus
    Get
      Return sel.Focused
    End Get
    Set(ByVal value As Boolean)
      If value Then
        sel.Focus()
      Else
        Me.Focus()
      End If
    End Set
  End Property

#End Region

#Region "IMainForm"

  Public ReadOnly Property MainMenu() As MainMenu Implements IMainForm.MainMenu
    Get
      Return mvMainMenu
    End Get
  End Property

#End Region

#Region "IDashboardTabContainer"

  Private Sub OpenDashboard() Implements CDBNETCL.IDashboardTabContainer.Open
    Dim vDashboard As DashboardTabControl = mvDashboard.CreateFromDatabase(Me)
    If vDashboard IsNot Nothing Then
      Dim vParent As Control = mvDashboard.Parent
      Me.SuspendLayout()
      vParent.Controls.Remove(mvDashboard)
      mvDashboard = vDashboard
      mvDashboard.Visible = False
      vParent.Controls.Add(mvDashboard)
      mvDashboard.SetItemID(mvEventInfo.EventNumber)
      mvDashboard.Visible = True
      Me.ResumeLayout()
    End If
  End Sub

  Private Sub SaveDashboard(ByVal pOptions As DashboardTabControl.SaveOptions) Implements CDBNETCL.IDashboardTabContainer.Save
    mvDashboard.SaveToDatabase(pOptions)
  End Sub

  Private Sub DeleteDashboard(ByVal pOptions As DashboardTabControl.DeleteOptions) Implements CDBNETCL.IDashboardTabContainer.Delete
    mvDashboard.DeleteFromDatabase(pOptions)
    OpenDashboard()     'This will either open another existing Dashboard or init a new Dashboard
  End Sub

  Private Sub NavigateHistoryItem(ByVal pHistoryEntityType As HistoryEntityTypes, ByVal pNumber As Integer) Implements CDBNETCL.IDashboardTabContainer.NavigateHistoryItem
    MainHelper.NavigateHistoryItem(pHistoryEntityType, pNumber)
  End Sub

  Private Sub ProcessSearch(ByVal pSearchText As String) Implements CDBNETCL.IDashboardTabContainer.ProcessSearch
    MainHelper.ProcessSearch(pSearchText)
  End Sub

  Private Sub SetBrowserMenu(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.SetBrowserMenu
    MainHelper.SetBrowserMenu(pSender, Me)
  End Sub

  Private Sub HistoryItemSelected(ByVal pSender As Object, ByVal pHistoryItem As CDBNETCL.UserHistoryItem, ByVal pDescription As String, ByVal pList As CDBNETCL.ArrayListEx) Implements CDBNETCL.IDashboardTabContainer.HistoryItemSelected
    MainHelper.HistoryItemSelected(pSender, pHistoryItem, pDescription, pList)
  End Sub

  Private Sub ActionItemSelected(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.ActionItemSelected
    MainHelper.ActionItemSelected(pSender)
  End Sub

  Private Sub SetActionMenu(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.SetActionMenu
    MainHelper.SetActionMenu(pSender, Me)
  End Sub

  Private Sub CalendarDoubleClickedHandler(ByVal pType As CalendarView.CalendarItemTypes, ByVal pDescription As String, ByVal pStart As Date, ByVal pEnd As Date, ByVal pUniqueID As Integer) Implements CDBNETCL.IDashboardTabContainer.CalendardItemDoubleClicked
    MainHelper.CalendarDoubleClicked(Me, pType, pDescription, pStart, pEnd, pUniqueID)
  End Sub

  Private Sub ProcessEditing(ByVal pType As CDBNETCL.DashboardDisplayPanel.MaintenanceTypes) Implements CDBNETCL.IDashboardTabContainer.ProcessEditing
    '
  End Sub

  Private Sub ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Implements IDashboardTabContainer.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber, True)
  End Sub

  Public Sub NavigateNewSelectionSet(pContactNumbers As String) Implements CDBNETCL.IDashboardTabContainer.NavigateNewSelectionSet
    FormHelper.CreateNewSelectionSet(pContactNumbers)
  End Sub

#End Region

  Public Sub New(ByVal pParentForm As MaintenanceParentForm)
    MyBase.New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pParentForm)
  End Sub

  Public Sub New(ByVal pParentForm As Form, ByVal pEventInfo As CareEventInfo, ByVal pEventDataType As CareServices.XMLEventDataSelectionTypes)
    MyBase.New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pParentForm, pEventInfo, pEventDataType)
  End Sub

  Private Sub InitialiseControls(ByVal pParentForm As Form, ByVal pEventInfo As CareEventInfo, ByVal pEventDataType As CareServices.XMLEventDataSelectionTypes)
    'Operating as a stand-alone popup dialog
    mvStandAlone = True
    mvStandAloneParent = pParentForm
    splTop.Panel1Collapsed = True         'No Header at present
    splBottom.Panel1Collapsed = True      'No selector
    epl.Visible = False
    splMaint.Panel1Collapsed = True       'No Combo 
    mvSelectedRow = -1
    cmdDelete.Visible = False
    cmdClose.Visible = False
    cmdNew.Visible = False
    cmdDefault.Visible = False
    HideGrid()
    With pEventInfo
      mvNumberOfSessions = .PickSessionsCount
      mvOptionDesc = .BookingOptionDesc
      mvBookingNumber = .BookingNumber
      mvBookingQuantity = .BookingQuantity
      mvBookerContact = .BookerContact
      mvOptionNumber = .BookingOptionNumber
    End With
    Select Case pEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
        Me.Text = ControlText.FrmSelectSessions
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
        Me.Text = ControlText.FrmNameAttendees
      Case CareNetServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
        Me.Text = ControlText.FrmReviewAccomodation
    End Select
    Init(pEventDataType, pEventInfo)
  End Sub

  Private Sub InitialiseControls(ByVal pParentForm As MaintenanceParentForm)
    splTop.Panel1Collapsed = True         'No Header at present
    mvHeader = New FormHeader
    mvHeader.Dock = DockStyle.Top         'BR12291: Used "Top" dock style instead of "Fill" to expand the header correctly
    splTop.Panel1.Controls.Add(mvHeader)
    epl.Visible = False
    splMaint.Panel1Collapsed = True       'No Combo 
    mvParentForm = pParentForm
    mvSelectedRow = -1
    cmdDelete.Visible = False
    cmdClose.Visible = False
    cmdNew.Visible = False
    cmdDefault.Visible = False
    mvMainMenu = MainHelper.AddMainMenu(Me)
    HideGrid()
    SetPanelVisibility()
  End Sub

  Public Sub Init(ByVal pType As CareServices.XMLEventDataSelectionTypes, ByVal pEventInfo As CareEventInfo, Optional ByVal pRetainPage As Boolean = False)
    Dim vEntityGroup As EntityGroup
    If DataHelper.EventGroups.ContainsKey(pEventInfo.EventGroup) Then
      vEntityGroup = DataHelper.EventGroups(pEventInfo.EventGroup)
      If vEntityGroup.ImageIndex < MainHelper.ImageProvider.NewTreeViewImages.Images.Count Then
        Me.Icon = Drawing.Icon.FromHandle(CType(MainHelper.ImageProvider.NewTreeViewImages.Images(vEntityGroup.ImageIndex), Bitmap).GetHicon)
        mvPrefix = vEntityGroup.GroupName & ": "
      Else
        Me.Icon = vEntityGroup.Icon
      End If
    Else
      Me.Icon = My.Resources.Events
    End If
    mvSelectedSessionNumbers = Nothing
    mvEventInfo = pEventInfo

    If mvStandAlone Then
      mvEventDataType = pType
    Else
      SettingsName = pEventInfo.EventGroup & "_CardSet"
      sel.Init(pEventInfo)
      mvEventMenu = New EventMenu(Me)
      mvEventMenu.EventInfo = mvEventInfo
      mvCustomiseMenu = New CustomiseMenu
      sel.TreeContextMenu = mvEventMenu
      If mvEventDataType <> -1 AndAlso pRetainPage AndAlso pEventInfo.EventNumber > 0 Then
        'Leave mvdatatype as it is
      Else
        mvEventDataType = pType
        sel.SetSelectionType(mvEventDataType)
      End If
      mvContainsBookings = pEventInfo.ContainsBookings
      RefreshHeader()
    End If
    RefreshSessions()
    dgr.AutoSetHeight = True
    RefreshCard()
    If mvEventInfo.EventNumber > 0 Then UserHistory.AddEventHistoryNode(mvEventInfo.EventNumber, mvEventInfo.EventName, mvEventInfo.EventGroup)
    mvBookingPriceChange = False
  End Sub

  Public Sub RefreshCard(Optional ByVal pMenuItem As EventMenu.EventMenuItems = EventMenu.EventMenuItems.emiNone)
    Try
      Dim vList As New ParameterList(True)
      Dim vShowGrid As Boolean = True
      Dim vHideSessions As Boolean = True
      Dim vShowEditPanel As Boolean = True
      Dim vSessionCount As Integer = 0

      cmdClose.Visible = False
      cmdOther.Visible = False
      cmdLink1.Visible = False
      cmdLink2.Visible = False
      mvRoomsTable = Nothing
      If mvHeader IsNot Nothing AndAlso mvEventInfo IsNot Nothing Then RefreshHeader()
      If mvEventDataType <> CareServices.XMLEventDataSelectionTypes.xedtEventDashboard AndAlso mvDashboard IsNot Nothing Then mvDashboard.Visible = False
      If mvEventDataType <> CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then mvSelectedSessionNumbers = Nothing
      Select Case mvEventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventAccommodation
        Case CareNetServices.XMLEventDataSelectionTypes.xedtEventActions
          'If mvSender Is Nothing Then
          mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction
          'End If
          cmdLink1.Text = ControlText.CmdActionLinks
          cmdLink1.Visible = True
          cmdLink2.Text = ControlText.CmdActionSubjects
          cmdLink2.Visible = True
          cmdLink1.Enabled = True
          cmdLink2.Enabled = True
          cmdNew.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation
          cmdClose.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
          cmdClose.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptions
          If mvEventInfo.MultiSession = True Then cmdOther.Visible = True
          cmdOther.Text = ControlText.CmdSessions
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptionSessions
          cmdClose.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookings
          vHideSessions = False
          cmdLink1.Visible = True
          cmdLink1.Text = ControlText.CmdSessions
          cmdOther.Visible = True
          cmdOther.Text = ControlText.CmdAllocate
        Case CareServices.XMLEventDataSelectionTypes.xedtEventContacts
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventContacts
        Case CareServices.XMLEventDataSelectionTypes.xedtEventCosts
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventCosts
        Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventDelegates
          vHideSessions = False
          mvBookingNumber = 0
        Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventInformation
          vShowGrid = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventMailings
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventMailings
        Case CareServices.XMLEventDataSelectionTypes.xedtEventOrganiser
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventOrganiser
          vShowGrid = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventOwners
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventOwners
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel
          vHideSessions = False
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ev_personnel_tasks_lookup) Then cmdOther.Visible = True
          cmdOther.Text = ControlText.CmdTasks
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPersonnelTasks
          vHideSessions = True
          If mvPersonnelInfo IsNot Nothing Then cmdClose.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPIS
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPIS

        Case CareServices.XMLEventDataSelectionTypes.xedtEventResources
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventResources
          vHideSessions = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventResults
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventResults
          vHideSessions = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventRoomBookings
          cmdOther.Visible = True
          cmdOther.Text = ControlText.CmdAllocate
        Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventRoomAllocation
          cmdClose.Visible = True
        Case CareServices.XMLEventDataSelectionTypes.xedtEventSessions
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSessions
        Case CareServices.XMLEventDataSelectionTypes.xedtEventSessionActivities
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSessionActivities
          vHideSessions = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventSessionTests
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSessionTests
          vHideSessions = False
        Case CareServices.XMLEventDataSelectionTypes.xedtEventSources
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSources
        Case CareServices.XMLEventDataSelectionTypes.xedtEventSubmissions
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSubmissions
        Case CareServices.XMLEventDataSelectionTypes.xedtEventTopics
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics
          If mvGroupID.Length > 0 Then
            vShowGrid = False
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventVenueBookings
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventVenueBookings
        Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventSessionCPD
          vHideSessions = False
        Case Else
          If mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctActionLink And
            mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctActionTopic Then
            mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone
          End If
      End Select
      epl.Visible = False
      splMaint.Panel1Collapsed = vHideSessions
      cmdDelete.Visible = CanDelete(mvMaintenanceType)
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone Then
        vShowEditPanel = False        'Delegates or Dashboard
      Else
        If (epl.PanelInfo Is Nothing) OrElse (epl.PanelInfo.MaintenanceType <> mvMaintenanceType) Then
          epl.ClearDataSources(epl)
          epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing, 0, mvEventInfo.EventGroup))
          epl.FillDeferredCombos(epl)
          epl.ContextMenuStrip = mvCustomiseMenu
          If mvCustomiseMenu IsNot Nothing Then mvCustomiseMenu.SetContext(Nothing, mvMaintenanceType, mvEventInfo.EventGroup)
        End If
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctActionTopic
            cmdClose.Visible = True
            cmdLink1.Visible = False
            cmdLink2.Visible = False
          Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
            cmdClose.Visible = True
            cmdLink1.Visible = False
            cmdLink2.Visible = False
            If epl.PanelInfo.PanelItems.Exists("ExamCentreDescription") Then epl.PanelInfo.PanelItems("ExamCentreDescription").Mandatory = False 'BR21326
          Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptionSessions
            Dim vSessions As New CollectionList(Of LookupItem)
            For Each vRow As DataRow In DirectCast(cboData.DataSource, DataTable).Rows
              If Not vRow.Item("SessionName").ToString.StartsWith("<") Then
                vSessions.Add(vRow.Item("SessionNumber").ToString, New LookupItem(vRow.Item("SessionNumber").ToString, vRow.Item("SessionName").ToString))
              End If
            Next
            epl.SetComboDataSource("SessionNumber", vSessions)

          Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
            Dim vList1 As New ParameterList(True)
            vList1.IntegerValue("OptionNumber") = mvOptionNumber
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions, mvEventInfo.EventNumber, vList1))
            If vTable IsNot Nothing Then epl.SetComboDataSource("SessionNumber", "SessionNumber", "SessionDesc", vTable, False)
          Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
            epl.FindTextLookupBox("ActivityGroup").FillComboWithRestriction("D")
            epl.FindTextLookupBox("RelationshipGroup").FillComboWithRestriction("D")
          Case CareServices.XMLMaintenanceControlTypes.xmctEventPIS
            'set the delegates combo
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventAttendees, mvEventInfo.EventNumber))
            If vTable Is Nothing Then
              vTable = New DataTable
              vTable.Columns.AddRange(New DataColumn() {New DataColumn("DelegateName"), New DataColumn("DelegateNumber")})
            End If
            epl.SetComboDataSource("EventDelegateNumber", "DelegateNumber", "DelegateName", vTable, True)

          Case CareServices.XMLMaintenanceControlTypes.xmctEventRoomBookings
            GetAvailableRoomTypes()
          Case CareServices.XMLMaintenanceControlTypes.xmctEventSessions
            epl.SetComboDataSource("VenueBookingNumber", "VenueNumber", "Description", DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventVenueBookings, mvEventInfo.EventNumber)))
          Case CareServices.XMLMaintenanceControlTypes.xmctEventResults
            Dim vTestList As New ParameterList(True)
            vTestList.IntegerValue("SessionNumber") = SessionNumber
            epl.SetComboDataSource("TestNumber", "TestNumber", "TestDesc", DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventSessionTests, mvEventInfo.EventNumber, vTestList)))
            vTestList.IntegerValue("BaseItemNumber") = mvEventInfo.BaseItemNumber
            epl.SetComboDataSource("ContactNumber", "ContactNumber", "ContactName", DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventCandidates, mvEventInfo.EventNumber, vTestList)))
          Case CareServices.XMLMaintenanceControlTypes.xmctEventBookings
            Dim vEventBookingOptionsDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions, mvEventInfo.EventNumber))
            mvDefaultOptionNumber = 0
            If vEventBookingOptionsDT IsNot Nothing AndAlso vEventBookingOptionsDT.Rows.Count = 1 Then
              mvDefaultOptionNumber = IntegerValue(vEventBookingOptionsDT.Rows(0).Item("OptionNumber"))
            End If
            epl.SetComboDataSource("OptionNumber", "OptionNumber", "OptionDesc", vEventBookingOptionsDT)
          Case CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel, CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation
            Dim vCombo As ComboBox = DirectCast(epl.FindComboBox("StandardPosition"), ComboBox)
            If vCombo IsNot Nothing Then
              If vCombo.Items.Count = 0 Then
                vCombo.Visible = False
                FindControl(epl, "Position", True).Visible = True
              Else
                vCombo.DropDownStyle = ComboBoxStyle.DropDown
              End If
            End If
            If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation And mvBookingNumber > 0 Then
              Dim vBookingList As New ParameterList(True)
              vBookingList.IntegerValue("BookingNumber") = mvBookingNumber
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookings, mvEventInfo.EventNumber, vBookingList))
              If vRow IsNot Nothing AndAlso IntegerValue(vRow("OrganisationNumber").ToString) > 0 Then
                Dim vTextLB As TextLookupBox = epl.FindTextLookupBox("ContactNumber")
                If vTextLB IsNot Nothing Then vTextLB.SetOrganisationContacts(IntegerValue(vRow("OrganisationNumber").ToString))
              End If
            End If
        End Select
        vShowEditPanel = True
      End If

      Dim vDataSet As DataSet
      If mvEventInfo.EventNumber > 0 Then
        If vHideSessions = False Then
          Dim vTable As DataTable = DirectCast(cboData.DataSource, DataTable)
          If (mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventSessionActivities OrElse mvEventDataType = CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD) _
             AndAlso mvEventInfo.MultiSession = True Then
            vSessionCount = vTable.Rows.Count - 1
            If vTable.DefaultView.RowFilter.Length = 0 Then
              vTable.DefaultView.RowFilter = String.Format("SessionNumber <> {0}", mvEventInfo.BaseItemNumber)
              If vSessionCount > 0 Then SelectComboBoxItem(cboData, vTable.DefaultView.Item(0).Item("SessionNumber").ToString, True)
            End If
          Else
            If vTable.DefaultView.RowFilter.Length > 0 OrElse (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel AndAlso SessionNumber = 0) Then 'BR12346 changed to add event personnel clause.
              vTable.DefaultView.RowFilter = ""
              SelectComboBoxItem(cboData, vTable.Rows(0).Item("SessionNumber").ToString, True)
            End If
          End If

          If ((SessionNumber <> mvEventInfo.BaseItemNumber) Or
              (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventResults)) Then vList.IntegerValue("SessionNumber") = SessionNumber
        End If

        Select Case mvEventDataType
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
            vList.IntegerValue("BookingNumber") = mvBookingNumber
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
            vList.IntegerValue("OptionNumber") = mvOptionNumber
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
            If mvBookingNumber > 0 Then
              vList.IntegerValue("BookingNumber") = mvBookingNumber
            Else
              vList.IntegerValue("BookingNumber") = 999999999             'Invalid number
            End If
            If mvSelectedSessionNumbers IsNot Nothing Then vList("SessionNumbers") = mvSelectedSessionNumbers.CSList
          Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
            vList.IntegerValue("RoomBookingNumber") = mvBookingNumber
          Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks
            If mvPersonnelInfo IsNot Nothing Then vList.IntegerValue("EventPersonnelNumber") = mvPersonnelInfo.PersonnelNumber
        End Select
        Select Case pMenuItem
          Case EventMenu.EventMenuItems.emiCalculateDelegateTotals, EventMenu.EventMenuItems.emiCalculateEventTotals, EventMenu.EventMenuItems.emiCalculateEventAndDelegateTotals
            Try
              Select Case pMenuItem
                Case EventMenu.EventMenuItems.emiCalculateDelegateTotals
                  vDataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventCalculateDelegateTotals, mvEventInfo.EventNumber, vList)
                Case EventMenu.EventMenuItems.emiCalculateEventTotals
                  vDataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventCalculateTotals, mvEventInfo.EventNumber, vList)
                Case EventMenu.EventMenuItems.emiCalculateEventAndDelegateTotals
                  vDataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventCalculateDelegateTotals, mvEventInfo.EventNumber, vList)
                  vDataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventCalculateTotals, mvEventInfo.EventNumber, vList)
              End Select
            Catch vCareEX As CareException
              If vCareEX.ErrorNumber = CareException.ErrorNumbers.enEventCalculateTotalsError Then
                'Errors occurred calculating the totals (most likely overflow errors) so refresh the screen and then display the error message
                Dim vErrorMessage As String = vCareEX.Message
                vList("SystemColumns") = "N"
                vDataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventHeaderInfo, mvEventInfo.EventNumber, vList)
                ShowInformationMessage(vErrorMessage)
              Else
                Throw vCareEX
              End If
            Catch vEX As Exception
              Throw vEX
            End Try
          Case Else
            If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventDashboard Then
              vShowGrid = False
              vDataSet = New DataSet
            Else
              If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
                vDataSet = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvActionNumber)
              ElseIf mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Then
                vDataSet = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, mvActionNumber)
              Else
                vDataSet = DataHelper.GetEventData(mvEventDataType, mvEventInfo.EventNumber, vList)
              End If

            End If
        End Select
      Else
        vDataSet = New DataSet
      End If

      If vShowGrid Then
        ShowGrid()
        Select Case mvEventDataType
          Case CareServices.XMLEventDataSelectionTypes.xedtEventFinancialHistory,
               CareServices.XMLEventDataSelectionTypes.xedtEventFinancialLinks
            dgr.MaxGridRows = 500
            splRight.Panel2Collapsed = True
            cmdSave.Visible = False
            cmdNew.Visible = False
            If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventFinancialLinks Then
              mvEventFinancialLinkMenu = New EventFinancialLinkMenu(Me)
              dgr.ContextMenuStrip = mvEventFinancialLinkMenu
            End If
          Case CareServices.XMLEventDataSelectionTypes.xedtEventActions
            mvActionMenu = New ActionMenu(Me)
            mvActionMenu.ActionType = ActionMenu.ActionTypes.EventActions
            mvActionMenu.MasterActionNumber = mvEventInfo.MasterAction
            dgr.ContextMenuStrip = mvActionMenu
          Case Else
            splRight.Panel2Collapsed = False
            dgr.MaxGridRows = DisplayTheme.DefaultMaxGridRows
            If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventAttendees Then
              mvEventDelegateMenu = New EventDelegateMenu(Me)
              dgr.ContextMenuStrip = mvEventDelegateMenu
              cmdSave.Visible = True
              cmdSave.Enabled = False
              cmdNew.Visible = False
            ElseIf mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventDocuments Then
              cmdSave.Visible = False
              cmdNew.Visible = False
            Else
              cmdSave.Visible = True
              cmdSave.Enabled = True
              cmdNew.Visible = True
              cmdNew.Enabled = True
              If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then
                mvEventFinancialLinkMenu = New EventFinancialLinkMenu(Me)
                dgr.ContextMenuStrip = mvEventFinancialLinkMenu
              End If
            End If

        End Select
        dgr.Populate(vDataSet)
        splRight.SplitterDistance = dgr.RequiredHeight
      Else
        HideGrid()
        cmdSave.Visible = True
        cmdSave.Enabled = True
        cmdNew.Visible = False
      End If

      If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventDashboard Then
        If mvTopicDataSheet IsNot Nothing Then mvTopicDataSheet.Visible = False
        If mvDashboard Is Nothing Then
          mvDashboard = New DashboardTabControl
          mvDashboard.Visible = False
          splRight.Panel2.Controls.Add(mvDashboard)
          mvDashboard.Dock = DockStyle.Fill
          mvDashboard.Init(Me, DashboardTypes.EventDashboardType, mvEventInfo.EventGroup)
          OpenDashboard()
        End If
        bpl.Visible = False
        mvDashboard.SetItemID(mvEventInfo.EventNumber)
        mvDashboard.Visible = True
        If epl.PanelInfo Is Nothing Then epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optDashboardTabName))
      ElseIf mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventTopics AndAlso mvGroupID.Length > 0 Then
        If mvDashboard IsNot Nothing Then mvDashboard.Visible = False
        If mvTopicDataSheet Is Nothing Then
          mvTopicDataSheet = New TopicDataSheet
          mvTopicDataSheet.Visible = False
          splRight.Panel2.Controls.Add(mvTopicDataSheet)
          mvTopicDataSheet.Dock = DockStyle.Fill
        End If
        mvTopicDataSheet.Init(mvEventInfo, mvGroupID, Nothing)
        mvTopicDataSheet.Visible = True
        mvTopicDataSheet.BringToFront()
        cmdDelete.Visible = False
      Else
        If mvDashboard IsNot Nothing Then mvDashboard.Visible = False
        If mvTopicDataSheet IsNot Nothing Then mvTopicDataSheet.Visible = False
        bpl.Visible = True
      End If

      With epl
        epl.Visible = vShowEditPanel
        If vDataSet.Tables.Contains("DataRow") Then
          If vShowGrid Then
            SelectRow(0)
          Else
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
            .Populate(vTable.Rows(0))
          End If
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
              If mvSelectedSessionNumbers Is Nothing Then
                mvSelectedSessionNumbers = New ArrayListEx
                For vRow As Integer = 0 To dgr.DataRowCount - 1
                  mvSelectedSessionNumbers.Add(IntegerValue(dgr.GetValue(vRow, "SessionNumber")))
                Next
                mvSessionsChanged = False
              End If
            Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
              epl.EnableControlList("MultiSession,Venue,VenueReference", False)
              If Not AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciEventDepartmentMaintenance) Then epl.EnableControl("Department", False)
              If IntegerValue(epl.GetValue("NumberOfAttendees")) > 0 OrElse IntegerValue(epl.GetValue("NumberOnWaitingList")) > 0 Then
                epl.EnableControlList("ChargeForWaiting,FreeOfCharge", False)
              End If
              If FindControl(epl, "EventPricingMatrix", False) IsNot Nothing Then
                Dim vEnabled As Boolean = False
                If DataHelper.EventGroups.ContainsKey(mvEventInfo.EventGroup) Then
                  vEnabled = DataHelper.EventGroups.Item(mvEventInfo.EventGroup).UseEventPricingMatrix
                End If
                If vEnabled = True Then vEnabled = Not (mvContainsBookings)
                If vEnabled Then
                  Dim vEPM As String = epl.GetValue("EventPricingMatrix")
                  epl.FindTextLookupBox("EventPricingMatrix").SetFilter(GetPricingMatrixFilter(""))
                  epl.SetValue("EventPricingMatrix", vEPM)
                End If
                epl.EnableControl("EventPricingMatrix", vEnabled)
              End If
              epl.SetDependancies("Organiser")
              epl.SetDependancies("ActivityGroup")
              cmdDelete.Enabled = True
            Case CareServices.XMLMaintenanceControlTypes.xmctEventOrganiser
              cmdDelete.Enabled = Not mvEventInfo.External
            Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
              SetDefaults()
          End Select
          mvEditing = True
        Else
          mvEditing = False
          epl.Clear()
          cmdDelete.Enabled = False
          SetDefaults(True)        'Set Defaults for new item
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
              With epl
                vList.Add("EventGroup", mvEventInfo.EventGroup)
                .Populate(vList)
                .EnableControlList("MultiSession,Venue,VenueReference,ChargeForWaiting,FreeOfCharge,EventPricingMatrix", True)
                .SetValue("StartDate", AppValues.TodaysDate)
                .SetValue("EndDate", AppValues.TodaysDate)
                .SetValue("StartTime", "09:00")
                .SetValue("EndTime", "17:30")
                .SetValue("WaitingListControlMethod", "A")
                .SetValue("ChargeForWaiting", "Y")
                .SetValue("CandidateNumberingMethod_I", "I")
                .SetValue("FirstCandidateNumber", "1")
                .SetValue("CandidateNumberBlockSize", "1000")
                .SetValue("NameAttendees", "Y")
                Dim vDefaultName As String = AppValues.DefaultName(mvEventInfo.EventGroup, "Event")
                Dim vDisable As Boolean
                If vDefaultName.Length = 0 Then vDisable = False Else vDisable = True
                .SetValue("EventDesc", vDefaultName, vDisable)
              End With
              If FindControl(epl, "EventPricingMatrix", False) IsNot Nothing Then
                Dim vEnabled As Boolean = False
                If DataHelper.EventGroups.ContainsKey(mvEventInfo.EventGroup) Then
                  vEnabled = DataHelper.EventGroups.Item(mvEventInfo.EventGroup).UseEventPricingMatrix
                End If
                If vEnabled Then epl.FindTextLookupBox("EventPricingMatrix").SetFilter(GetPricingMatrixFilter(""))
                epl.EnableControl("EventPricingMatrix", vEnabled)
              End If
            Case CareServices.XMLMaintenanceControlTypes.xmctEventSessions
              If mvEventInfo.MultiSession = False Then
                epl.EnableControls(False)
                cmdSave.Visible = False
                cmdNew.Visible = False
                cmdDelete.Visible = False
              End If
            Case CareServices.XMLMaintenanceControlTypes.xmctEventPIS
              SetCommandsForNew()
          End Select
        End If
        'If pCode.Length = 0 Then .Focus()
      End With
      If mvEventDataType = CareNetServices.XMLEventDataSelectionTypes.xedtEventActions Then
        Utilities.SetActionChangeReason(epl, mvEditing, False)
      End If
      CheckOwnership()
      If vShowEditPanel Then epl.DataChanged = False

      'We need to set the combox box value after the combobox is made visible.
      SetStandardPositions()

      bpl.RepositionButtons()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public Sub RefreshSessions()
    SessionNumber = 0
    If mvEventInfo.EventNumber > 0 Then
      cboData.ValueMember = "SessionNumber"
      cboData.DisplayMember = "SessionName"
      cboData.DataSource = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventSessionNames, mvEventInfo.EventNumber))
      cboData_SelectedIndexChanged(Me, New EventArgs)
    End If
  End Sub

  Public Sub RefreshHeader()
    If mvEventInfo.EventNumber > 0 AndAlso MainHelper.ShowHeaderPanel Then
      mvHeader.Populate(mvEventInfo)
      mvHeader.Visible = True
      splTop.Panel1Collapsed = False
      If splTop.Width > 0 Then splTop.SplitterDistance = mvHeader.Height
    Else
      splTop.Panel1Collapsed = True
    End If
    Me.Text = mvPrefix & mvEventInfo.EventDescription
  End Sub

  Public ReadOnly Property SelectedSessions() As ArrayListEx
    Get
      Return mvSelectedSessionNumbers
    End Get
  End Property

  Private Sub CheckOwnership()
    Dim vDisableEdit As Boolean
    Dim vDisableLink1 As Boolean
    Dim vDisableOther As Boolean
    Dim vDisableStatus As Boolean

    If mvEventInfo.UserIsOwner = False Then
      vDisableEdit = True
      vDisableLink1 = True
      vDisableOther = True
      vDisableStatus = True
    Else
      Select Case mvEventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
          vDisableEdit = Not mvEventInfo.CanBook
          If dgr.DataRowCount > 0 Then
            vDisableOther = epl.GetValue("BookingStatus") = ebsCancelled  'Allocate
            vDisableLink1 = Not mvPickSessions  'Sessions
          Else
            vDisableOther = True
            vDisableLink1 = True
            vDisableStatus = True
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
          vDisableEdit = Not mvEventInfo.CanBook
      End Select
    End If
    If vDisableEdit Then
      epl.EnableControls(False)
      cmdSave.Enabled = False
      If vDisableStatus = False And mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then
        Dim vComboBox As ComboBox = epl.FindComboBox("BookingStatus")
        If vComboBox.Items.Count > 1 Then
          vComboBox.Enabled = True
          cmdSave.Enabled = True
        End If
      End If
      If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then
        If epl.GetValue("BookingStatus") <> ebsCancelled Then
          epl.FindTextBox("Notes").Enabled = True
          cmdSave.Enabled = True
        Else
          epl.FindTextBox("Notes").Enabled = False
        End If
      End If
      cmdDelete.Enabled = False
      cmdNew.Enabled = False
    End If
    If vDisableOther Then cmdOther.Enabled = False
    If vDisableLink1 Then cmdLink1.Enabled = False
  End Sub

  Public Overrides ReadOnly Property CareEventInfo() As CareEventInfo
    Get
      Return mvEventInfo
    End Get
  End Property

  Private Sub EventTabSelected(ByVal pSender As Object, ByVal pType As CareServices.XMLEventDataSelectionTypes, ByVal pGroupID As String, ByVal pReadOnlyPage As Boolean) Handles sel.EventTabSelected
    Dim vCursor As New BusyCursor
    Try
      If mvEventDataType <> pType OrElse
        (mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventTopics AndAlso pGroupID <> mvGroupID) OrElse
        (mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks AndAlso mvPersonnelInfo IsNot Nothing) Then
        mvEventDataType = pType
        mvGroupID = pGroupID
        mvReadOnlyPage = pReadOnlyPage
        mvPersonnelInfo = Nothing             'No restriction for tasks any more
        RefreshCard()
      Else
        'Get here when the card is first opened due to an after_select on the TabSelector TreeView
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ShowGrid()
    dgr.Visible = True
    dgr.ContextMenuStrip = Nothing
    splRight.Panel1Collapsed = False
  End Sub
  Private Sub GetPrimaryKeyValues(ByVal pList As ParameterList, ByVal pRow As Integer, ByVal pForUpdate As Boolean)
    If mvEventInfo.EventNumber > 0 Then pList("EventNumber") = mvEventInfo.EventNumber.ToString

    mvKeyValuesColl = GetPrimaryKeyNames(mvMaintenanceType, pForUpdate)
    If mvKeyValuesColl.Count > 0 Then
      For Each vDGKeyValue As DisplayGridKeyValue In mvKeyValuesColl
        pList(vDGKeyValue.ParameterName) = dgr.GetValue(pRow, vDGKeyValue.GridColumnName)
      Next
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctEventBookings
          If mvSelectedSessionNumbers IsNot Nothing Then pList("SessionNumbers") = mvSelectedSessionNumbers.CSList
        Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
          pList.IntegerValue("BookingNumber") = mvBookingNumber
        Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptionSessions
          pList.IntegerValue("OptionNumber") = mvOptionNumber
        Case CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation
          pList.IntegerValue("BookingNumber") = mvBookingNumber
        Case CareServices.XMLMaintenanceControlTypes.xmctEventPersonnelTasks
          Dim vEPN As Integer = IntegerValue(dgr.GetValue(pRow, "EventPersonnelNumber"))
          If vEPN > 0 Then pList.IntegerValue("EventPersonnelNumber") = vEPN
        Case CareServices.XMLMaintenanceControlTypes.xmctEventResources
          pList("Product") = dgr.GetValue(pRow, "Product")
          pList("CopyTo") = dgr.GetValue(pRow, "CopyTo")
          If IntegerValue(dgr.GetValue(pRow, "ResourceNumber")) <= 0 Then pList.Remove("ResourceNumber")
        Case CareServices.XMLMaintenanceControlTypes.xmctEventRoomAllocation
          pList.IntegerValue("RoomBookingNumber") = mvBookingNumber
      End Select
    End If
  End Sub
  Private Sub GetAdditionalKeyValues(ByVal pList As ParameterList)
    If mvEventInfo.EventNumber > 0 Then pList("EventNumber") = mvEventInfo.EventNumber.ToString
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptionSessions
        pList.IntegerValue("OptionNumber") = mvOptionNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
        If mvBookingNumber > 0 Then
          pList.IntegerValue("BookingNumber") = mvBookingNumber
        Else
          pList.IntegerValue("BookingNumber") = 999999999             'Invalid number
        End If
        If mvSelectedSessionNumbers IsNot Nothing AndAlso mvSelectedSessionNumbers.Count > 0 Then
          pList("SessionNumbers") = mvSelectedSessionNumbers.CSList
        Else
          pList.IntegerValue("SessionNumbers") = 0
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation
        If mvBookingNumber > 0 Then pList.IntegerValue("BookingNumber") = mvBookingNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
        pList("EventGroup") = mvEventInfo.EventGroup
      Case CareServices.XMLMaintenanceControlTypes.xmctEventPersonnelTasks
        If mvPersonnelInfo IsNot Nothing Then pList.IntegerValue("EventPersonnelNumber") = mvPersonnelInfo.PersonnelNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSessionTests, CareServices.XMLMaintenanceControlTypes.xmctEventResults,
           CareServices.XMLMaintenanceControlTypes.xmctEventResources, CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel,
           CareServices.XMLMaintenanceControlTypes.xmctEventSessionActivities, CareNetServices.XMLMaintenanceControlTypes.xmctEventSessionCPD
        pList("SessionNumber") = SessionNumber.ToString
      Case CareServices.XMLMaintenanceControlTypes.xmctEventRoomAllocation
        pList.IntegerValue("RoomBookingNumber") = mvBookingNumber
    End Select
  End Sub
  Private Function EditingExistingRecord() As Boolean
    If NoEditingAllowed() Then
      Return False
    Else
      Return mvEditing
    End If
  End Function
  Private Function NoEditingAllowed() As Boolean
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSources, CareServices.XMLMaintenanceControlTypes.xmctEventMailings,
           CareServices.XMLMaintenanceControlTypes.xmctEventSessionActivities, CareServices.XMLMaintenanceControlTypes.xmctEventOwners,
           CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions, CareServices.XMLMaintenanceControlTypes.xmctActionLink,
           CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        Return True
    End Select
  End Function
  Private Sub ProcessNew()
    dgr.SelectRow(-1)
    mvSelectedRow = -1
    epl.Clear()
    SetDefaults(False)
    SetCommandsForNew()
  End Sub
  Protected Overrides Function ProcessSave(ByVal pDefault As Boolean, ByVal sender As System.Object) As Boolean 'Return true if saved
    Try
      Dim vList As New ParameterList(True)
      Dim vEditing As Boolean = EditingExistingRecord()

      If vEditing Then
        'If editing an existing record then get the primary key values
        GetPrimaryKeyValues(vList, mvSelectedRow, True)
      Else
        'For new records add in any additional key values
        GetAdditionalKeyValues(vList)
      End If
      If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink AndAlso vList.ContainsKey("EventNumber") Then
        vList.Remove("EventNumber")   'Adding an ActionLink does not require the EventNumber as we are not linking the Action to the Event.
      End If
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics AndAlso mvGroupID.Length > 0 Then
        If mvTopicDataSheet.ValidateTopics Then
          If mvTopicDataSheet.DataChanged = True AndAlso mvEventInfo.EventPricingMatrix.Length > 0 AndAlso mvContainsBookings = True Then ShowInformationMessage(InformationMessages.ImEventBookingPriceChange)
          mvTopicDataSheet.SaveTopics(mvEventInfo.EventNumber)
        End If
      ElseIf epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
        'Only allow booking against FOC booking option if the event is not FOC 
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookings AndAlso vEditing = False AndAlso Not mvEventInfo.FreeOfCharge Then
          If Not mvEventInfo.IsFocBookingOption(vList("OptionNumber")) Then
            ShowInformationMessage(InformationMessages.ImNotFOCBookingOption)
            Exit Function
          End If

          'Only allow a rate of zero for all the bookings that are booked against FOC booking option(s)
          If Not mvEventInfo.IfZeroRateProduct(vList) Then
            ShowInformationMessage(InformationMessages.ImInvalidRate)
            Exit Function
          End If
        ElseIf mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAction Then
          'BR20338 - Action Schedules - Creating a Schedule for an action with no duration behaves differently between contacts and exams and events
          If FindControl(epl, "ScheduledOn", False) IsNot Nothing AndAlso epl.GetValue("ScheduledOn").Length > 0 Then
            epl.GetScheduleDate(vList, CareNetServices.XMLActionScheduleTypes.xastGivenDate, (Not vEditing))
          End If
        End If
        'Update or Insert record
        If vEditing Then
          If ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookings) Or
    (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventRoomBookings)) AndAlso vList("BookingStatus") = ebsCancelled Then
            Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource)
            If vParams.Count > 0 Then
              For Each vValue As DictionaryEntry In vParams
                If Not vList.Contains(vValue.Key.ToString) Then vList.Add(vValue.Key.ToString, vValue.Value.ToString)
              Next
              If ConfirmUpdate() = False Then Exit Function
              mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
            Else
              Exit Function
            End If
          Else
            If mvEventInfo.EventPricingMatrix.Length > 0 AndAlso mvContainsBookings = True _
AndAlso ((mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventInformation And mvBookingPriceChange = True) OrElse (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics AndAlso mvGroupID.Length = 0)) Then
              'Some details have changed that may effect the price of existing bookings so warn user
              If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics Then mvBookingPriceChange = epl.DataChanged
              If mvBookingPriceChange Then ShowInformationMessage(InformationMessages.ImEventBookingPriceChange)
            End If
            mvBookingPriceChange = False

            If ConfirmUpdate() = False Then Exit Function
            Try
              mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
            Catch vCareEx As CareException
              If vCareEx.ErrorNumber = CareException.ErrorNumbers.enActivitiesLinkedToDelegate Then
                vList("Confirm") = CBoolYN(ShowQuestion(vCareEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
                mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
              Else
                Throw vCareEx
              End If
            Catch vEx As Exception
              Throw vEx
            End Try
          End If
        Else
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
              If mvSelectedSessionNumbers Is Nothing Then mvSelectedSessionNumbers = New ArrayListEx
              If Not mvSelectedSessionNumbers.Contains(vList.IntegerValue("SessionNumber")) Then
                If vList.ContainsKey("IgnoreTimeConflicts") AndAlso vList("IgnoreTimeConflicts") = "N" AndAlso mvSelectedSessionNumbers.Count > 0 Then
                  Dim vSessions As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventSessions, vList.IntegerValue("EventNumber")))
                  Dim vNewSessionRow As DataRow = Nothing
                  For Each vRow As DataRow In vSessions.Rows
                    If vRow.Item("SessionNumber").ToString = vList("SessionNumber") Then
                      vNewSessionRow = vRow
                      Exit For
                    End If
                  Next
                  If vNewSessionRow IsNot Nothing Then
                    Dim vStartDate As Date = CDate(vNewSessionRow.Item("StartDate").ToString & " " & vNewSessionRow.Item("StartTime").ToString)
                    Dim vEndDate As Date = CDate(vNewSessionRow.Item("EndDate").ToString & " " & vNewSessionRow.Item("EndTime").ToString)
                    For Each vRow As DataRow In vSessions.Rows
                      If mvSelectedSessionNumbers.Contains(IntegerValue(vRow.Item("SessionNumber").ToString)) Then
                        Dim vCheckStartDate As Date = CDate(vRow.Item("StartDate").ToString & " " & vRow.Item("StartTime").ToString)
                        Dim vCheckEndDate As Date = CDate(vRow.Item("EndDate").ToString & " " & vRow.Item("EndTime").ToString)
                        If vCheckStartDate < vEndDate And vCheckEndDate > vStartDate Then
                          Throw New CareException(GetInformationMessage(InformationMessages.ImSessionDateTimeConflict, vRow("SessionDesc").ToString), CareException.ErrorNumbers.enSessionDateTimeConflict)
                        End If
                      End If
                    Next
                  End If
                End If
                mvSelectedSessionNumbers.Add(vList.IntegerValue("SessionNumber"))
                mvSessionsChanged = True
              End If
            Case Else
              If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink Then
                vList("ActionNumber") = mvActionNumber.ToString
                vList("Notified") = "N"
              End If
              If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics Then
                If mvEventInfo.EventPricingMatrix.Length > 0 AndAlso mvContainsBookings = True And mvGroupID.Length = 0 Then
                  'Some details have changed that may effect the price of existing bookings so warn user
                  ShowInformationMessage(InformationMessages.ImEventBookingPriceChange)
                End If
              End If
              If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction Then
                vList("FromEvent") = "Y"
              End If
              mvBookingPriceChange = False
              If ConfirmInsert() = False Then Exit Function
              mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)

              Select Case mvMaintenanceType
                Case CareNetServices.XMLMaintenanceControlTypes.xmctAction
                  If mvEditing = False AndAlso DataHelper.UserInfo.ContactNumber > 0 And mvReturnList("ActionNumber").ToString.Length > 0 Then
                    If ShowQuestion(QuestionMessages.QmNoActioners, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                      vList("ActionNumber") = mvReturnList("ActionNumber").ToString
                      vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
                      vList("ActionLinkType") = "A"
                      vList("Notified") = "N"
                      DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
                    End If
                  End If

                Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
                  If epl.Recipients IsNot Nothing AndAlso epl.Recipients.Rows.Count > 1 Then
                    'This is PostPoint contacts - need to add Action Link to each Contact
                    'First Contact has already been done
                    For vIndex As Integer = 1 To epl.Recipients.Rows.Count - 1
                      vList("ContactNumber") = epl.Recipients.Rows(vIndex).Item("ContactNumber").ToString
                      mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
                    Next
                  End If

                Case CareNetServices.XMLMaintenanceControlTypes.xmctEventOrganiser
                  mvEditing = True
              End Select
          End Select
        End If
        mvRefreshParent = True
        epl.DataChanged = False     'Data saved now

        'Save the primary key values for Selection
        If Not vEditing Then
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

        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
            If mvEventInfo.EventNumber = 0 Then       'Added new event
              mvEventInfo = New CareEventInfo(mvReturnList.IntegerValue("EventNumber"), vList("EventGroup"))
              mvEventMenu.EventInfo = mvEventInfo
              UserHistory.AddEventHistoryNode(mvEventInfo.EventNumber, mvEventInfo.EventName, mvEventInfo.EventGroup)
              RefreshSessions()
              sel.SetSelectionType(mvEventDataType) ' BR13730
              RefreshCard()
            Else
              If mvEventInfo.MultiSession = False AndAlso IntegerValue(epl.GetValue("MaximumAttendees")) > mvEventInfo.MaximumAttendees AndAlso mvEventInfo.NumberOnWaitingList > 0 Then
                mvEventMenu_ItemSelected(EventMenu.EventMenuItems.emiProcessWaitingList)
              End If
              RefreshSessions()
              mvEventInfo.RefreshData()
            End If
            RefreshHeader()
          Case CareServices.XMLMaintenanceControlTypes.xmctEventSessions
            If vEditing AndAlso IntegerValue(epl.GetValue("MaximumAttendees")) > mvSessionMaximum AndAlso IntegerValue(dgr.GetValue(dgr.CurrentRow, "NumberOnWaitingList")) > 0 Then
              mvEventMenu_ItemSelected(EventMenu.EventMenuItems.emiProcessWaitingList)
            End If
            RefreshSessions()
          Case CareServices.XMLMaintenanceControlTypes.xmctEventBookings
            If mvReturnList.ContainsKey("ProcessWaitingList") Then
              mvEventMenu_ItemSelected(EventMenu.EventMenuItems.emiProcessWaitingList)
            Else
              mvEventInfo.RefreshData()                                                 'Changed number booked etc..
            End If
          Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptions
            If Not vEditing Then
              mvOptionNumber = mvReturnList.IntegerValue("OptionNumber")
              mvOptionDesc = vList("OptionDesc")
              mvPickSessions = vList("PickSessions").ToString = "Y"
            End If
            mvNumberOfSessions = vList.IntegerValue("NumberOfSessions")
            If FindControl(epl, "NumberOfSessions", False).Enabled AndAlso mvPickSessions Then
              vList = New ParameterList(True)
              vList.IntegerValue(("OptionNumber")) = mvOptionNumber
              Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions, mvEventInfo.EventNumber, vList))
              If (vTable Is Nothing OrElse vTable.Rows.Count <= mvNumberOfSessions) Then
                mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
                RefreshCard()
              End If
            End If
          Case CareServices.XMLMaintenanceControlTypes.xmctEventVenueBookings
            RefreshHeader()
          Case CareServices.XMLMaintenanceControlTypes.xmctEventOrganiser
            cmdDelete.Enabled = True
        End Select
        If NoEditingAllowed() Then ProcessNew()
        If dgr.Visible Then     'Only one record
          RePopulateGrid()
        End If
        Return True
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
        Case CareException.ErrorNumbers.enCannotResolveUserName
          ShowInformationMessage(vEx.Message)
          RefreshCard()
        Case CareException.ErrorNumbers.enEventParameterError, CareException.ErrorNumbers.enSessionDateTimeConflict,
             CareException.ErrorNumbers.enQuantityExceedsMaximum, CareException.ErrorNumbers.enEventBooking,
             CareException.ErrorNumbers.enAppointmentConflict,
             CareException.ErrorNumbers.enCannotUpdateDelegateOnEventPIS,
             CareException.ErrorNumbers.enEventSponsorshipInfoMissing,
             CareException.ErrorNumbers.enNotEnoughPIS,
             CareException.ErrorNumbers.enNotAllPISAllocated,
              CareException.ErrorNumbers.enInvalidActionDuration,
             CareException.ErrorNumbers.enCannotUpdateSponsorshipAsPayments,
             CareException.ErrorNumbers.enCannotDeleteDelegateAsPaidPIS,
             CareException.ErrorNumbers.enCannotBookJointToEvent,
             CareException.ErrorNumbers.enEventTopicMandatory,
             CareException.ErrorNumbers.enTooManyDelegates, CareException.ErrorNumbers.enCCAuthorisationFailed,
             CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enInvalidRateValue
          epl.SetErrorField("Rate", vEx.Message)
        Case Else
          Throw vEx
      End Select
    End Try
  End Function
  Protected Overrides Sub SetCommandsForNew()
    cmdDelete.Enabled = False
    mvEditing = False
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctEventPIS
        'set the PIS Number combo or textlookup
        Dim vControl As Control = FindControl(epl, "PisNumber")
        If TryCast(vControl, ComboBox) IsNot Nothing Then
          Dim vList As ParameterList = New ParameterList(True)
          vList("BankAccount") = AppValues.ControlValue(AppValues.ControlValues.pis_bank_account)
          Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPISNumbers, vList)
          epl.SetComboDataSource("PisNumber", "PISNumber", "PISNumber", vTable, False)
        End If
        epl.EnableControl("PisNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptions
        cmdLink1.Visible = False
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSessions
        cmdLink1.Visible = False
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
        epl.SetValue("ActionPriority", AppValues.DefaultActionPriority)
        Dim vTimeSpan As TimeSpan = AppValues.DefaultActionDuration
        If vTimeSpan.Days > 0 Then epl.SetValue("DurationDays", vTimeSpan.Days.ToString)
        If vTimeSpan.Hours > 0 Then epl.SetValue("DurationHours", vTimeSpan.Hours.ToString)
        If vTimeSpan.Minutes > 0 Then epl.SetValue("DurationMinutes", vTimeSpan.Minutes.ToString)
        Utilities.SetActionChangeReason(epl, False, False)
    End Select
  End Sub
  Protected Overrides Sub SetDefaults(Optional ByVal pInitialSetup As Boolean = True)
    Select Case mvMaintenanceType
      'BR 20528 : To enable Contact number field for Event Contact
      Case CareServices.XMLMaintenanceControlTypes.xmctEventContacts
        epl.EnableControl("ContactNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventAccommodation
        epl.SetDateTimeValue("FromDate", mvEventInfo.StartDate)
        epl.SetDateTimeValue("ToDate", mvEventInfo.EndDate)
        epl.EnableControlList("RoomType,FromDate,ToDate,Product,Rate,OrganisationNumber,AddressNumber,ContactNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptions
        epl.EnableControls(True)
        Dim vDisable As Boolean = Not mvEventInfo.MultiSession
        epl.SetValue("DeductFromEvent", "Y", vDisable)
        epl.SetValue("PickSessions", "N", vDisable)
        epl.SetValue("IssueEventResources", "Y", vDisable)
        epl.SetValue("NumberOfSessions", "1", vDisable)
        epl.SetValue("MaximumBookings", "1", mvEventInfo.EligibilityCheckRequired)
        If FindControl(epl, "MinimumBookings", False) IsNot Nothing Then
          epl.SetValue("MinimumBookings", "1", mvEventInfo.EligibilityCheckRequired, False, False, False)
        End If
        cmdOther.Enabled = False
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions
        epl.SetValue("OptionDesc", mvOptionDesc)
        epl.SetValue("NumberOfSessions", mvNumberOfSessions.ToString)
        epl.EnableControl("SessionNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookingOptionSessions
        epl.SetValue("OptionDesc", mvOptionDesc)
        epl.SetValue("NumberOfSessions", mvNumberOfSessions.ToString)
        Dim vDisabled As Boolean
        If mvPickSessions Then vDisabled = True
        epl.SetValue("Allocation", IIf(mvNumberOfSessions > 0, CDbl(100 / mvNumberOfSessions).ToString("#.00"), "100.00").ToString, vDisabled)
        epl.EnableControlList("SessionNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventBookings
        Dim vCanBook As Boolean = mvEventInfo.CanBook
        mvSelectedSessionNumbers = Nothing
        mvBookingNumber = 0
        epl.EnableControls(vCanBook)
        cmdSave.Enabled = vCanBook
        SetBookingStatusRestriction("", "")
        epl.SetValue("Quantity", "1")
        cmdOther.Enabled = False
        cmdLink1.Enabled = False
        If vCanBook AndAlso mvDefaultOptionNumber <> 0 Then
          epl.SetValue("OptionNumber", mvDefaultOptionNumber.ToString)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation
        epl.SetValue("Quantity", mvBookingQuantity.ToString, True)
        epl.SetValue("BookedBy", mvBookerContact.ContactName, True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventRoomBookings
        Dim vCanBook As Boolean = mvEventInfo.CanBook
        epl.EnableControls(vCanBook)
        cmdSave.Enabled = vCanBook
        With mvEventInfo
          epl.SetDateTimeValue("FromDate", .StartDate)
          epl.SetDateTimeValue("ToDate", .EndDate)
          Dim vValid As Boolean
          ValidateItem(epl, "FromDate", .StartDate.ToString(AppValues.DateFormat), vValid)
        End With
        SetBookingStatusRestriction("", "")
        cmdOther.Enabled = False
      Case CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel
        If SessionNumber > 0 AndAlso SessionNumber <> mvEventInfo.BaseItemNumber Then
          Dim vList As ParameterList = New ParameterList(True)
          vList.IntegerValue("SessionNumber") = SessionNumber
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventSessions, mvEventInfo.EventNumber, vList))
          epl.SetDateTimeValue("StartDate", CDate(vDataRow.Item("StartDate")))
          epl.SetDateTimeValue("EndDate", CDate(vDataRow.Item("EndDate")))
          epl.SetValue("StartTime", vDataRow.Item("StartTime").ToString)
          epl.SetValue("EndTime", vDataRow.Item("EndTime").ToString)
        Else
          epl.SetDateTimeValue("StartDate", mvEventInfo.StartDate)
          epl.SetDateTimeValue("EndDate", mvEventInfo.EndDate)
          epl.SetDateTimeValue("StartTime", mvEventInfo.StartTime)
          epl.SetDateTimeValue("EndTime", mvEventInfo.EndTime)
        End If
        cmdOther.Enabled = False
      Case CareServices.XMLMaintenanceControlTypes.xmctEventPersonnelTasks
        If mvPersonnelInfo IsNot Nothing Then
          mvPersonnelInfo.SetEPLValues(epl)
        Else
          With mvEventInfo
            epl.SetDateTimeValue("StartDate", .StartDate)
            epl.SetDateTimeValue("EndDate", .EndDate)
            epl.SetDateTimeValue("StartTime", .StartTime)
            epl.SetDateTimeValue("EndTime", .EndTime)
          End With
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctEventResources
        epl.EnableControlList("ResourceNumber,ResourceType,QuantityRequired,CopyTo,IssueBasis", True)
        'epl.EnableControlList("ResourceNumber,ResourceType,OrganisationNumber,AddressNumber,ContactNumber,ExternalResourceType,ObtainedOn,ReturnBy,ReturnedOn,TotalAmount,DueDate,Deposit,DepositPaidDate", True)
        cmdSave.Enabled = True
      Case CareServices.XMLMaintenanceControlTypes.xmctEventResults
        epl.EnableControlList("ContactNumber,TestNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSessions
        With mvEventInfo
          epl.SetValue("Subject", .Subject)
          epl.SetValue("SkillLevel", .SkillLevel)
          epl.SetValue("Location", .Location)
          epl.SetValue("MinimumAttendees", .MinimumAttendees.ToString)
          epl.SetValue("MaximumAttendees", .MaximumAttendees.ToString)
          epl.SetValue("MaximumOnWaitingList", .MaximumOnWaitingList.ToString)
          epl.SetValue("TargetAttendees", .TargetAttendees.ToString)
          epl.SetDateTimeValue("StartDate", .StartDate)
          epl.SetDateTimeValue("EndDate", .EndDate)
          epl.SetDateTimeValue("StartTime", .StartTime)
          epl.SetDateTimeValue("EndTime", .EndTime)
          epl.EnableControlList("StartDate,StartTime,EndDate,EndTime", True)
        End With
        If FindControl(epl, "CpdCategory", False) IsNot Nothing Then
          'Customised form still displaying CPD data so always disable
          epl.EnableControlList("CpdApprovalStatus,CpdDateApproved,CpdAwardingBody,CpdCategory,CpdYear,CpdPoints,CpdNotes", False)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSessionTests
        epl.EnableControl("TestNumber", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventSubmissions
        epl.SetDependancies("Forwarded")
        epl.SetDependancies("Returned")
      Case CareServices.XMLMaintenanceControlTypes.xmctEventTopics
        epl.EnableControlList("Topic,SubTopic", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
        epl.SetValue("NameAttendees", "Y", False)
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
        epl.SetValue("ActionPriority", AppValues.DefaultActionPriority)
        Dim vTimeSpan As TimeSpan = AppValues.DefaultActionDuration
        If vTimeSpan.Days > 0 Then epl.SetValue("DurationDays", vTimeSpan.Days.ToString)
        If vTimeSpan.Hours > 0 Then epl.SetValue("DurationHours", vTimeSpan.Hours.ToString)
        If vTimeSpan.Minutes > 0 Then epl.SetValue("DurationMinutes", vTimeSpan.Minutes.ToString)
        epl.SetValue("ScheduledOn", String.Empty)
        epl.SetValue("Deadline", String.Empty)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
        epl.SetEntityLinkDefaults("ActionLinkType", "R")
      Case CareNetServices.XMLMaintenanceControlTypes.xmctEventSessionCPD
        epl.SetValue("WebPublish", "N")
        epl.SetUserDefaults()
    End Select
    epl.DataChanged = False
  End Sub
  Private Sub SetItemAsMaxColumn(ByVal pColumnName As String)
    epl.SetValue(pColumnName, CStr(dgr.MaxColumnValue(pColumnName) + 1))
    epl.EnableControl(pColumnName, True)
  End Sub

  Protected Overrides Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      If epl.DataChanged Then
        If ConfirmSave() Then
          If ProcessSave(False, sender) = False Then Exit Sub
          If dgr.Visible Then     'Only one record
            RePopulateGrid()
          End If
        End If
      End If
      'Clear selection on display grid and set defaults for new record
      ProcessNew()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overrides Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try

      'TODO Confirm Update or Insert 
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookings AndAlso mvPickSessions AndAlso mvBookingNumber = 0 Then
        mvBookingList = New ParameterList(True)
        GetAdditionalKeyValues(mvBookingList)
        If epl.AddValuesToList(mvBookingList, True, EditPanel.AddNullValueTypes.anvtAll) Then
          'We have saved the booking information in the mvBookingList - now we need to choose the sessions
          mvSelectedSessionNumbers = Nothing
          mvOptionNumber = IntegerValue(mvBookingList("OptionNumber"))
          mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
          RefreshCard()
        End If
      Else
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation Then
          If FindControl(epl, "Position").Visible = False Then
            epl.SetValue("Position", epl.GetValue("StandardPosition"))
          End If
        End If
        If ProcessSave(False, sender) Then
          If dgr.Visible Then     'Only one record
            If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions Then
              ProcessNew()
            End If
          End If
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Protected Overrides Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      'TODO Confirm cancel changes 
      Dim vList As New ParameterList(True)
      GetPrimaryKeyValues(vList, mvSelectedRow, False)
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventOrganiser Then
        Dim vTextBox As TextLookupBox = epl.FindTextLookupBox("Organiser")
        vList("Organiser") = vTextBox.Text
      ElseIf mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventTopics Then
        If mvEventInfo.EventPricingMatrix.Length > 0 AndAlso mvContainsBookings = True Then
          ShowInformationMessage(InformationMessages.ImEventBookingPriceChange)
        End If
      ElseIf mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctAction Then
        vList("MasterActionNumber") = dgr.GetValue(mvSelectedRow, "MasterAction")
        vList("FromEvent") = "Y"
      ElseIf mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or
        mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
        vList("ActionNumber") = mvActionNumber.ToString
        If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Then
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
        End If
      End If
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventBookingSessions Then
        mvSelectedSessionNumbers.Remove(vList.IntegerValue("SessionNumber"))
        mvSessionsChanged = True
      Else
        If Not ConfirmDelete() Then Exit Sub
        Try
          DataHelper.DeleteItem(mvMaintenanceType, vList)
        Catch vCareEx As CareException
          If vCareEx.ErrorNumber = CareException.ErrorNumbers.enActivitiesLinkedToDelegate Then
            vList("Confirm") = CBoolYN(ShowQuestion(vCareEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes)
            DataHelper.DeleteItem(mvMaintenanceType, vList)
          ElseIf vCareEx.ErrorNumber = CareException.ErrorNumbers.enCannotDeletePersonnelRecord Then
            ShowInformationMessage(vCareEx.Message)
          Else
            Throw vCareEx
          End If
        Catch vEx As Exception
          Throw vEx
        End Try
      End If
      mvRefreshParent = True
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctEventInformation
          If mvEventInfo.EventNumber > 0 Then UserHistory.RemoveEventHistoryNode(mvEventInfo.EventNumber, mvEventInfo.EventGroup)
          Me.Close() 'Deleted the only item
          Close()
        Case CareServices.XMLMaintenanceControlTypes.xmctEventOrganiser
          RefreshCard()
      End Select
    Catch vException As CareException
      Select Case vException.ErrorNumber
        Case CareException.ErrorNumbers.enCannotDelete,
         CareException.ErrorNumbers.enCannotDeleteEventPISAsDelegates,
         CareException.ErrorNumbers.enCannotDeleteEventPISAsPayments,
         CareException.ErrorNumbers.enCannotDeleteSessionCPDDelegate
          ShowInformationMessage(vException.Message)
        Case Else
          DataHelper.HandleException(vException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      If dgr.Visible Then RePopulateGrid()
    End Try
  End Sub
  Protected Overrides Sub cmdLink_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vCancel As Boolean

    Try
      If epl.DataChanged Then
        vCancel = Not ProcessSave(False, sender)
      End If
      If Not vCancel Then
        Select Case mvEventDataType
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
            If sender Is cmdOther Then
              'Allocate button for bookings
              mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
              RefreshCard()
            Else
              'Sessions button for pick sessions booking
              mvSelectedSessionNumbers = Nothing
              mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
              RefreshCard()
            End If
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions
            If sender Is cmdLink1 Then
              'Copy
              cmdLink1.Visible = False
              Dim vControl As TextBox = epl.FindTextBox("OptionDesc")
              If vControl.Text.Length < vControl.MaxLength - 7 Then vControl.Text &= " (copy)"
              dgr.SelectRow(-1)
              mvSelectedRow = -1
              SetCommandsForNew()
              If ProcessSave(False, sender) Then RePopulateGrid()
            Else
              'Sessions button for booking options
              mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
              RefreshCard()
            End If
          Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
            'Tasks button for personnel
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks
            RefreshCard()
          Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
            'Allocate button for room bookings
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
            RefreshCard()
          Case CareServices.XMLEventDataSelectionTypes.xedtEventSessions
            cmdLink1.Visible = False
            Dim vControl As TextBox
            vControl = epl.FindTextBox("SessionDesc")
            If vControl.Text.Length < vControl.MaxLength - 7 Then vControl.Text &= " (copy)"
            vControl = epl.FindTextBox("LongDescription")
            vControl.Text &= " (copy)"
            dgr.SelectRow(-1)
            mvSelectedRow = -1
            SetCommandsForNew()
            If ProcessSave(False, sender) Then RePopulateGrid()
          Case CareServices.XMLEventDataSelectionTypes.xedtEventActions
            If dgr.DataRowCount > 0 Then
              mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtNone
              If sender Is cmdLink1 Then
                mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
              Else
                mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtNone
                mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic
              End If
              RefreshCard()
            End If
        End Select
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Protected Overrides Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      Select Case mvEventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
          If mvStandAlone Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
          Else
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings
            mvBookingNumber = 0
            RefreshCard()
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
          'Put checks in to ensure correct number of sessions
          If mvStandAlone Then Me.DialogResult = System.Windows.Forms.DialogResult.None
          If CanCloseBookingSessions() Then
            If mvStandAlone Then
              Me.DialogResult = System.Windows.Forms.DialogResult.OK
              Me.Close()
            Else
              If mvBookingNumber > 0 Then
                'If editing an existing booking then update it here
                If mvSelectedSessionNumbers IsNot Nothing AndAlso mvSessionsChanged Then
                  Dim vList As New ParameterList(True)
                  vList.IntegerValue("EventNumber") = mvEventInfo.EventNumber
                  vList.IntegerValue("BookingNumber") = mvBookingNumber
                  vList("SessionNumbers") = mvSelectedSessionNumbers.CSList
                  mvReturnList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctEventBookings, vList)
                End If
              Else
                'Creating a new booking
                mvBookingList("SessionNumbers") = mvSelectedSessionNumbers.CSList
                mvReturnList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctEventBookings, mvBookingList)
                mvEventInfo.RefreshData()                                       'Changed number booked etc..
              End If
              mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings
              RefreshCard()
            End If
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
          'Put checks in to ensure correct number of sessions and allocation
          If CanCloseBookingOptionSessions() Then
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions
            RefreshCard()
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks
          mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
          mvPersonnelInfo = Nothing
          RefreshCard()
        Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
          Dim vValid As Boolean = True
          Dim vList As New ParameterList(True)
          vList("RoomBookingNumber") = mvBookingNumber.ToString
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings, mvEventInfo.EventNumber, vList))
          If vDataRow.Item("EnforceAllocation").ToString = "Y" Then
            'Enforce Allocation is set for Room Type; ensure different Contact has been assigned to each place.
            Dim vContactNumber As Long
            For vRow As Integer = 0 To dgr.DataRowCount - 1
              vContactNumber = CLng(dgr.GetValue(vRow, "ContactNumber"))
              For vRow2 As Integer = vRow + 1 To dgr.DataRowCount - 1
                If CLng(dgr.GetValue(vRow2, "ContactNumber")) = vContactNumber Then vValid = False
              Next
            Next
          End If
          If vValid = False Then
            ShowInformationMessage(InformationMessages.ImEnforceAllocation)
            Me.DialogResult = System.Windows.Forms.DialogResult.None
            '            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
            '            epl.DataChanged = False
            '            RefreshCard()
          Else
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
            mvBookingNumber = 0
            RefreshCard()
          End If
        Case Else
          If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or
            mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
            mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventActions
            RefreshCard()
          Else
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
          End If

      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function CanCloseBookingSessions() As Boolean
    Dim vSessionCount As Integer = dgr.DataRowCount

    If vSessionCount <> mvNumberOfSessions And mvNumberOfSessions <> 0 Then
      ShowInformationMessage(InformationMessages.ImNumberOfSessions, CStr(mvNumberOfSessions))
    ElseIf vSessionCount = 0 And mvNumberOfSessions = 0 Then
      ShowInformationMessage(InformationMessages.ImNumberOfSessions, "1")
    Else
      Return True
    End If
  End Function

  Private Function CanCloseBookingOptionSessions() As Boolean
    Dim vSessionCount As Integer = dgr.DataRowCount

    If mvPickSessions AndAlso (vSessionCount < mvNumberOfSessions OrElse (vSessionCount = 0)) Then
      ShowInformationMessage(InformationMessages.ImPickSessionsCount, CStr(mvNumberOfSessions))
      Return False
    ElseIf mvPickSessions AndAlso (vSessionCount = mvNumberOfSessions) Then
      Return True
    ElseIf Not mvPickSessions AndAlso vSessionCount <> mvNumberOfSessions Then
      ShowInformationMessage(InformationMessages.ImNumberOfSessions, CStr(mvNumberOfSessions))
      Return False
    ElseIf Not mvPickSessions Then
      Dim vAllocation As Double
      For vRow As Integer = 0 To dgr.DataRowCount - 1
        vAllocation += DoubleValue(dgr.GetValue(vRow, "Allocation"))
      Next
      If FixTwoPlaces(vAllocation) <> 100 Then
        If ShowQuestion(QuestionMessages.QmAllocationTotalError, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
          Return True
        End If
      Else
        Return True
      End If
    Else
      Return True
    End If
  End Function

  Private Sub BeforeSelect(ByVal pSender As Object, ByRef pCancel As Boolean) Handles sel.BeforeSelect
    If mvEventInfo IsNot Nothing AndAlso mvEventInfo.EventNumber = 0 Then
      pCancel = True
    Else
      If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions AndAlso CanCloseBookingOptionSessions() = False Then
        pCancel = True
      Else
        Dim vDataChanged As Boolean
        If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventTopics AndAlso mvGroupID.Length > 0 Then
          vDataChanged = mvTopicDataSheet.DataChanged
        Else
          vDataChanged = epl.DataChanged
        End If
        If vDataChanged Then
          If ConfirmSave() Then
            pCancel = Not ProcessSave(False, pSender)
          Else
            'We have some data changed and we are going to cancel it
            If mvEditing = False Then
              epl.DataChanged = False
            End If
          End If
        End If
      End If
    End If
  End Sub

  Private Sub GetCodeRestrictions(ByVal pSender As Object, ByVal pParameterName As String, ByVal pList As ParameterList) Handles epl.GetCodeRestrictions
    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation
        pList("FindProductType") = "A"      'Accommodation
      Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation
        pList("FindProductType") = "O"      'Sponsorship Event
      Case CareServices.XMLEventDataSelectionTypes.xedtEventPIS
        pList("BankAccount") = AppValues.ControlValue(AppValues.ControlValues.pis_bank_account)
      Case Else
        pList("FindProductType") = "E"      'Event
    End Select
  End Sub

  Private Sub epl_GetInitialCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByRef pList As ParameterList) Handles epl.GetInitialCodeRestrictions
    Select Case mvEventDataType
      Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
        If pParameterName = "CpdCategoryType" Then
          If pList Is Nothing Then pList = New ParameterList(True)
          pList("WithoutCPDCycle") = "Y"  'This will make sure we do not restrict the types for CPD Cycles
        End If
    End Select
  End Sub

  Private Sub RePopulateGrid()
    Dim vList As ParameterList = New ParameterList(True)
    GetAdditionalKeyValues(vList)
    If (((SessionNumber > 0) AndAlso (SessionNumber <> mvEventInfo.BaseItemNumber)) Or
        (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventResults)) Then
      vList("SessionNumber") = SessionNumber.ToString
    ElseIf vList.ContainsKey("SessionNumber") Then
      vList.Remove("SessionNumber")
    End If
    Select Case mvMaintenanceType
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
        dgr.Populate(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, mvActionNumber))
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic
        dgr.Populate(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvActionNumber))
      Case Else
        dgr.Populate(DataHelper.GetEventData(mvEventDataType, mvEventInfo.EventNumber, vList))
    End Select

    If mvActionMenu IsNot Nothing And dgr.DataRowCount = 0 Then mvActionMenu.ActionNumber = 0

    If dgr.DataRowCount = 0 Then
      ProcessNew()
    Else
      If dgr.RequiredHeight > splRight.SplitterDistance Then splRight.SplitterDistance = dgr.RequiredHeight
      If Not mvEditing Then mvSelectedRow = dgr.FindRow(mvKeyValuesColl.GridColumnNames, mvKeyValuesColl.Values)
      'Select current row
      If mvSelectedRow <= 0 Then mvSelectedRow = 0
      If mvSelectedRow > dgr.DataRowCount - 1 Then mvSelectedRow = dgr.DataRowCount - 1
      dgr.SelectRow(mvSelectedRow)
      SelectRow(mvSelectedRow)
    End If

  End Sub
  Protected Overrides Sub SelectRow(ByVal pRow As Integer)
    Dim vCursor As New BusyCursor
    Try
      If dgr.DataRowCount > pRow Then     'JIRA1395 Modified to only select a record if there actually is some data
        Select Case mvEventDataType
          Case CareServices.XMLEventDataSelectionTypes.xedtEventSources, CareServices.XMLEventDataSelectionTypes.xedtEventMailings,
            CareServices.XMLEventDataSelectionTypes.xedtEventOwners, CareServices.XMLEventDataSelectionTypes.xedtEventSessionActivities,
            CareServices.XMLEventDataSelectionTypes.xedtEventFinancialHistory
            'No edit so no select
            cmdDelete.Enabled = CanDelete(mvMaintenanceType)
          Case CareServices.XMLEventDataSelectionTypes.xedtEventFinancialLinks
            cmdDelete.Enabled = False
            mvEventFinancialLinkMenu.SetContext(mvEventInfo, IntegerValue(dgr.GetValue(pRow, "BatchNumber")), IntegerValue(dgr.GetValue(pRow, "TransactionNumber")), IntegerValue(dgr.GetValue(pRow, "LineNumber")))
          Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingSessions
            epl.SetValue("OptionDesc", mvOptionDesc)
            epl.SetValue("NumberOfSessions", mvNumberOfSessions.ToString)
            cmdDelete.Enabled = True
            epl.SetValue("SessionNumber", dgr.GetValue(pRow, "SessionNumber"), True)
            epl.DataChanged = False
            'Case CareServices.XMLEventDataSelectionTypes.xedtEventActionTopics, CareServices.XMLEventDataSelectionTypes.xedtEventActionLinks
            '  cmdDelete.Visible = True
            '  If mvSelectedRow > -1 Then cmdDelete.Enabled = True
            '  mvEditing = False
          Case Else

            Dim vList As New ParameterList(True)
            vList("SystemColumns") = "N"                'Ensure we get all the columns
            Dim vCount As Integer = vList.Count
            GetPrimaryKeyValues(vList, pRow, True)
            Dim vDataRow As DataRow = Nothing
            If vList.Count > vCount Then
              If mvEventDataType = CareNetServices.XMLEventDataSelectionTypes.xedtEventActions Then
                If dgr.GetValue(pRow, "ActionNumber").Length > 0 Then mvActionNumber = CInt(dgr.GetValue(pRow, "ActionNumber"))
                If mvActionMenu IsNot Nothing Then
                  mvActionMenu.ActionStatus = dgr.GetValue(pRow, "ActionStatus")
                  If dgr.GetValue(pRow, "ActionNumber").Length > 0 Then mvActionMenu.ActionNumber = CInt(dgr.GetValue(pRow, "ActionNumber"))
                  If dgr.GetValue(pRow, "MasterAction").Length > 0 Then mvActionMenu.MasterActionNumber = CInt(dgr.GetValue(pRow, "MasterAction"))
                End If
                If vList.ContainsKey("ActionNumber") AndAlso vList("ActionNumber").Length > 0 Then vDataRow = DataHelper.GetRowFromDataSet(DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionInformation, vList.IntegerValue("ActionNumber")))

              Else
                If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
                  vDataRow = Nothing
                Else
                  vDataRow = DataHelper.GetRowFromDataSet(DataHelper.GetEventData(mvEventDataType, mvEventInfo.EventNumber, vList))
                End If

              End If
              If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Or mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings Then
                If vDataRow IsNot Nothing Then
                  SetBookingStatusRestriction(vDataRow.Item("BookingStatusCode").ToString, vDataRow.Item("BatchNumber").ToString)
                End If
              End If
              If vDataRow IsNot Nothing Then epl.Populate(vDataRow)
            End If
            cmdDelete.Enabled = CanDelete(mvMaintenanceType)

            Select Case mvEventDataType
              'BR 20528 : To disable Conatct number field while amending existing Event contact
              Case CareServices.XMLEventDataSelectionTypes.xedtEventContacts
                If vList.ContainsKey("ContactNumber") Then
                  Dim vTextLookupBox As TextLookupBox = epl.FindTextLookupBox("ContactNumber")
                  If vTextLookupBox IsNot Nothing Then
                    vTextLookupBox.Enabled = False
                  End If
                End If
              Case CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation
                Dim vNumberOfNights As Integer = SetNumberOfNights()
                Dim vQuantity As Integer = IntegerValue(dgr.GetValue(pRow, "NumberOfRooms"))
                Dim vNightsAvailable As Integer = IntegerValue(dgr.GetValue(pRow, "NightsAvailable"))
                Dim vEnable As Boolean = vNightsAvailable = vQuantity * vNumberOfNights
                epl.EnableControlList("RoomType,FromDate,ToDate,Product,Rate,OrganisationNumber,AddressNumber,ContactNumber", vEnable)
                cmdDelete.Enabled = vEnable
              Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees
                mvEventDelegateMenu.SetContext(mvEventInfo)
                cmdSave.Enabled = True
              Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
                mvBookingNumber = IntegerValue(dgr.GetValue(pRow, "BookingNumber"))
                Dim vCancelled As Boolean = epl.GetValue("BookingStatus") = ebsCancelled
                epl.EnableControls(Not vCancelled)
                cmdSave.Enabled = Not vCancelled
                cmdOther.Enabled = Not vCancelled
                mvBookingQuantity = IntegerValue(epl.GetValue("Quantity"))
                mvBookerContact = epl.FindTextLookupBox("ContactNumber").ContactInfo
                mvOptionNumber = IntegerValue(epl.GetValue("OptionNumber"))
                cmdLink1.Enabled = mvPickSessions AndAlso Not vCancelled
                If epl.GetValue("BookingStatus") = ebsBookedAndPaid Then
                  Dim vBookingStatus As ComboBox = epl.FindComboBox("BookingStatus")
                  vBookingStatus.Enabled = False
                  Dim vBookingOption As ComboBox = epl.FindComboBox("OptionNumber")
                  vBookingOption.Enabled = False
                  Dim vQuantity As TextBox = epl.FindTextBox("Quantity")
                  vQuantity.Enabled = False
                  Dim vRate As ComboBox = epl.FindComboBox("Rate")
                  vRate.Enabled = False
                  Dim vSalesContact As TextLookupBox = epl.FindTextLookupBox("SalesContactNumber")
                  vSalesContact.Enabled = False
                End If
                If Not mvPickSessions Or vCancelled Then mvSelectedSessionNumbers = Nothing
                mvEventFinancialLinkMenu.SetContext(mvEventInfo, mvEventDataType, vDataRow)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
                epl.SetValue("Quantity", mvBookingQuantity.ToString, True)
                If mvBookerContact IsNot Nothing Then epl.SetValue("BookedBy", mvBookerContact.ContactName, True)
                cmdNew.Enabled = dgr.DataRowCount < mvBookingQuantity
                If FindControl(epl, "StandardPosition").Visible = True Then
                  epl.SetValue("StandardPosition", epl.GetValue("Position"))
                End If
                epl.DataChanged = False
              Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions
                If IntegerValue(dgr.GetValue(pRow, "BookingCount")) > 0 Then
                  epl.EnableControls(False)
                  cmdOther.Enabled = False
                Else
                  If Not mvEventInfo.MultiSession Then
                    If epl.GetValue("NumberOfSessions") <> "1" Then
                      epl.SetValue("NumberOfSessions", "1", True)
                    Else
                      epl.EnableControl("NumberOfSessions", False)
                    End If
                    epl.EnableControlList("PickSessions,IssueEventResources,DeductFromEvent", False)
                  Else
                    mvOptionNumber = CInt(vList("OptionNumber"))
                    If vDataRow IsNot Nothing Then
                      mvOptionDesc = vDataRow.Item("OptionDesc").ToString
                      mvNumberOfSessions = IntegerValue(vDataRow.Item("NumberOfSessions").ToString)
                      mvPickSessions = vDataRow.Item("PickSessions").ToString = "Y"
                      cmdOther.Enabled = True
                    End If
                  End If
                  If mvEventInfo.EligibilityCheckRequired Then
                    If epl.GetValue("MaximumBookings") <> "1" Then
                      epl.SetValue("MaximumBookings", "1", True)
                    Else
                      epl.EnableControl("MaximumBookings", False)
                    End If
                    If FindControl(epl, "MinimumBookings", False) IsNot Nothing Then
                      If epl.GetValue("MinimumBookings") <> "1" Then
                        epl.SetValue("MinimumBookings", "1", True)
                      Else
                        epl.EnableControl("MinimumBookings", False)
                      End If
                    End If
                  End If
                End If
                If mvEventInfo.ContainsBookings AndAlso mvEventInfo.HasBookingsForBookingOption(dgr.GetValue(pRow, "OptionNumber")) Then
                  epl.EnableControl("FreeOfCharge", False)
                End If
                cmdLink1.Visible = True
                cmdLink1.Text = ControlText.MnuQBECopy
                bpl.RepositionButtons()

              Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions
                epl.SetValue("OptionDesc", mvOptionDesc)
                epl.SetValue("NumberOfSessions", mvNumberOfSessions.ToString)
                epl.EnableControl("SessionNumber", False)
                If mvPickSessions Then epl.EnableControl("Allocation", False)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
                mvPersonnelInfo = New EventPersonnelInfo(IntegerValue(dgr.GetValue(pRow, "EventPersonnelNumber")), epl)
                cmdOther.Enabled = True
                If FindControl(epl, "StandardPosition").Visible = True Then
                  epl.SetValue("StandardPosition", epl.GetValue("Position"))
                End If
                epl.DataChanged = False
              Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks
                If mvPersonnelInfo IsNot Nothing Then epl.EnableControlList("ContactNumber,AddressNumber", False)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventPIS
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

              Case CareServices.XMLEventDataSelectionTypes.xedtEventResources
                epl.EnableControl("ResourceType", False)
                epl.SetDependancies("ResourceNumber")
                epl.SetDependancies("ResourceType")
                If epl.GetValue("ResourceType") = "" Then
                  Dim vTextLookupBox As TextLookupBox = epl.FindTextLookupBox("ResourceNumber")
                  vTextLookupBox.Description = dgr.GetValue(dgr.CurrentRow, "StandardProductDesc")
                  cmdSave.Enabled = False
                Else
                  cmdSave.Enabled = True
                End If
              Case CareServices.XMLEventDataSelectionTypes.xedtEventResults
                epl.EnableControl("ContactNumber", False)
                epl.EnableControl("TestNumber", False)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
                mvBookingNumber = IntegerValue(dgr.GetValue(pRow, "BookingNumber"))
                SetNumberOfNights()
                Dim vCancelled As Boolean = epl.GetValue("BookingStatus") = ebsCancelled
                epl.EnableControls(Not vCancelled)
                cmdSave.Enabled = Not vCancelled
                cmdOther.Enabled = Not vCancelled
                mvBookingQuantity = IntegerValue(epl.GetValue("Quantity"))
                mvBookerContact = epl.FindTextLookupBox("ContactNumber").ContactInfo
                'BR13580: Disable the following controls as they are not yet supported for Update
                If Not vCancelled Then epl.EnableControlList("BlockBookingNumber,FromDate,ToDate,RoomType,Quantity", False)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
                epl.SetValue("Quantity", mvBookingQuantity.ToString, True)
                epl.SetValue("BookedBy", mvBookerContact.ContactName, True)
                epl.EnableControlList("Position,OrganisationName", False)
                cmdNew.Enabled = dgr.DataRowCount < mvBookingQuantity
              Case CareServices.XMLEventDataSelectionTypes.xedtEventSessions
                If mvEventInfo.MoveSessionDates = False Then
                  epl.EnableControlList("StartDate,StartTime,EndDate,EndTime", False)
                End If
                If FindControl(epl, "CpdCategory", False) IsNot Nothing Then
                  'Customised form still displaying CPD data so always disable
                  epl.EnableControlList("CpdApprovalStatus,CpdDateApproved,CpdAwardingBody,CpdCategory,CpdYear,CpdPoints,CpdNotes", False)
                End If
                If IntegerValue(dgr.GetValue(pRow, "NumberOfAttendees")) > 0 Or
                IntegerValue(dgr.GetValue(pRow, "NumberInterested")) > 0 Or
                IntegerValue(dgr.GetValue(pRow, "NumberOnWaitingList")) > 0 Then cmdDelete.Enabled = False
                mvSessionMaximum = IntegerValue(epl.GetValue("MaximumAttendees"))

                cmdLink1.Visible = True
                cmdLink1.Text = ControlText.MnuQBECopy

              Case CareServices.XMLEventDataSelectionTypes.xedtEventSessionTests
                epl.EnableControl("TestNumber", False)
                epl.SetDependancies("GradeDataType")
                epl.SetDependancies("MinimumValue")
                epl.SetDependancies("MaximumValue")
                epl.SetDependancies("Pattern")
              Case CareServices.XMLEventDataSelectionTypes.xedtEventTopics
                epl.EnableControlList("Topic,SubTopic", False)
              Case CareServices.XMLEventDataSelectionTypes.xedtEventVenueBookings
                SetVenueInformation()
              Case CareServices.XMLEventDataSelectionTypes.xedtEventSubmissions
                epl.SetDependancies("Forwarded")
                epl.SetDependancies("Returned")
              Case CareServices.XMLEventDataSelectionTypes.xedtEventDocuments
                'vNumber = SelectRowItemNumber(pDataRow, "DocumentNumber")
                'mvDocumentMenu.DocumentNumber = vNumber
                'mvContactInfo.SelectedDocumentNumber = vNumber
                'tbp.Text = ControlText.TbpPrecis
                'txt.Text = DataHelper.GetDocumentPrecis(vNumber)
                'txt.BackColor = SystemColors.Window
                'mvDataSet2 = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, vNumber)
                'dgr2.Populate(mvDataSet2)
                'mvDataSet3 = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentLinks, vNumber)
                'dgr3.Populate(mvDataSet3)
                'mvDocumentMenu.SetNotifyProcessed(dgr3)
                'dgr4.MaxGridRows = DisplayTheme.HistoryMaxGridRows
                'dgr4.Populate(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentHistory, vNumber))
                'vTabCount = 4
              Case CareServices.XMLEventDataSelectionTypes.xedtEventActions
                Dim vActionNumber As Integer = IntegerValue(dgr.GetValue(pRow, "ActionNumber"))
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
                  'cmdSave.Enabled = True
                  Utilities.SetActionChangeReason(epl, (vActionNumber > 0), False)
                End If

              Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
                AppHelper.CPDCategoryTypeValueChanged(epl, True, SessionNumber)
                epl.SetErrorField("CpdCategory", "")
                Dim vEnable As Boolean = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciEventCPDMaintenance)
                epl.EnableControls(vEnable)
                cmdSave.Enabled = vEnable
                cmdNew.Enabled = vEnable

            End Select
            If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionLink Or
               mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic Then
              cmdDelete.Visible = True
              If mvSelectedRow > -1 Then cmdDelete.Enabled = True
              mvEditing = False
            Else
              mvEditing = True
            End If
        End Select
        mvSelectedRow = pRow
        CheckOwnership()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ValidateAllItems(ByVal pSender As Object, ByVal pList As ParameterList, ByRef pValid As Boolean) Handles epl.ValidateAllItems
    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation
        If Date.Parse(epl.GetValue("ToDate")).Subtract(Date.Parse(epl.GetValue("FromDate"))).Days < 1 Then
          epl.SetErrorField("ToDate", InformationMessages.ImMustSelectOneNight)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingOptions
        If pList.IntegerValue("NumberOfSessions") < 1 And pList("PickSessions") = "N" Then
          epl.SetErrorField("NumberOfSessions", InformationMessages.ImNumberOfSessionsGT0)
          pValid = False
        End If
        If pList.ContainsKey("MinimumBookings") AndAlso pList.ContainsKey("MaximumBookings") AndAlso
           pList.IntegerValue("MinimumBookings") > pList.IntegerValue("MaximumBookings") Then
          epl.SetErrorField("MinimumBookings", InformationMessages.ImMinimumLTMax)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventCosts
        If DoubleValue(pList("TotalAmount").ToString) <> DoubleValue(pList("Deposit").ToString) + DoubleValue(pList("Balance").ToString) Then
          epl.SetErrorField("Balance", InformationMessages.ImBalanceTotalMinusDeposit)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventPIS
        If pList("Amount").Length > 0 Xor epl.GetValue("BankedBy").Length > 0 Then
          epl.SetErrorField("Amount", InformationMessages.ImBankedByFields)
          pValid = False
        ElseIf epl.GetValue("BankedOn").Length > 0 Xor pList("Amount").Length > 0 Then
          epl.SetErrorField("Amount", InformationMessages.ImBankedByFields)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
        Dim vSessionStartDateAndTime As DateHelper
        Dim vSessionStartDate As DateHelper
        Dim vPersonnelStartDateTime As DateHelper
        Dim vPersonnelStartDate As DateHelper
        Dim vSessionEndDateAndTime As DateHelper
        Dim vSessionEndDate As DateHelper
        Dim vPersonnelEndDateTime As DateHelper
        Dim vPersonnelEndDate As DateHelper
        If SessionNumber > 0 AndAlso SessionNumber <> mvEventInfo.BaseItemNumber Then
          Dim vList As ParameterList = New ParameterList(True)
          vList.IntegerValue("SessionNumber") = SessionNumber
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventSessions, mvEventInfo.EventNumber, vList))
          vSessionStartDateAndTime = New DateHelper(vDataRow.Item("StartTime").ToString)
          vSessionStartDate = New DateHelper(vDataRow.Item("StartDate").ToString)
          vSessionEndDateAndTime = New DateHelper(vDataRow.Item("EndTime").ToString)
          vSessionEndDate = New DateHelper(vDataRow.Item("EndDate").ToString)
        Else
          vSessionStartDateAndTime = New DateHelper(mvEventInfo.StartTime.ToString)
          vSessionStartDate = New DateHelper(mvEventInfo.StartDate.Date)
          vSessionEndDateAndTime = New DateHelper(mvEventInfo.EndTime.ToString)
          vSessionEndDate = New DateHelper(mvEventInfo.EndDate.Date)
        End If
        vPersonnelStartDateTime = New DateHelper(epl.GetValue("StartTime"))
        vPersonnelStartDate = New DateHelper(epl.GetValue("StartDate"))
        vPersonnelEndDateTime = New DateHelper(epl.GetValue("EndTime"))
        vPersonnelEndDate = New DateHelper(epl.GetValue("EndDate"))
        If vPersonnelStartDate.DateValue < vSessionStartDate.DateValue Then
          epl.SetErrorField("StartDate", InformationMessages.ImBeforeSessionStartdate)
          pValid = False
        End If
        If pValid AndAlso vPersonnelStartDateTime.DateValue < vSessionStartDateAndTime.DateValue Then
          epl.SetErrorField("StartTime", InformationMessages.ImBeforeSessionStartdate)
          pValid = False
        End If
        If pValid AndAlso vPersonnelEndDate.DateValue > vSessionEndDate.DateValue Then
          epl.SetErrorField("EndDate", InformationMessages.ImAfterSessionEndDate)
          pValid = False
        End If
        If pValid AndAlso vPersonnelEndDateTime.DateValue > vSessionEndDateAndTime.DateValue Then
          epl.SetErrorField("EndTime", InformationMessages.ImAfterSessionEndDate)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation, CareServices.XMLEventDataSelectionTypes.xedtEventSessions
        Dim vTemplate As Boolean = False
        Dim vNumberOfAttendees As Integer
        Dim vNumberOnWaitingList As Integer
        If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventInformation Then
          Dim vBaseSessionEndDate As Date
          Dim vBaseSessionEndDateTime As New DateHelper(epl.GetValue("EndDate"), epl.GetValue("EndTime"))
          If pList.Contains("MultiSession") AndAlso pList("MultiSession") = "Y" Then
            If Date.TryParse(pList("EndDate"), vBaseSessionEndDate) Then
              'get end dates of sessions
              If pList.Contains("EventNumber") AndAlso pList("EventNumber").ToString.Length > 0 Then
                Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventSessions, IntegerValue(pList("EventNumber"))))
                If Not vDataTable Is Nothing Then
                  For Each vRow As DataRow In vDataTable.Rows
                    If vBaseSessionEndDate < CDate(vRow.Item("EndDate").ToString) Then
                      epl.SetErrorField("EndDate", InformationMessages.ImCannotSetEndDate)
                      pValid = False
                    ElseIf vBaseSessionEndDate = CDate(vRow.Item("EndDate").ToString) Then
                      Dim vSessionEndDate As String = vRow.Item("EndDate").ToString
                      Dim vSessionEndTime As String = vRow.Item("EndTime").ToString
                      Dim vSessionEndDateTime As New DateHelper(vSessionEndDate, vSessionEndTime)
                      If vBaseSessionEndDateTime.DateValue < vSessionEndDateTime.DateValue Then
                        epl.SetErrorField("EndTime", InformationMessages.ImCannotSetEndDate)
                        pValid = False
                      End If
                    End If
                  Next
                End If
              End If
            End If
          End If
          If pList("External") = "Y" Then
            If pList("FreeOfCharge") = "Y" Then
              epl.SetErrorField("FreeOfCharge", InformationMessages.ImCannotSetExternalAndFree)
              pValid = False
            End If
            If pList("Organiser").Length = 0 Then                     'if external is set then enforce the organiser as required
              epl.SetErrorField("Organiser", InformationMessages.ImOrganiserRequired)
              pValid = False
            End If
            If IntegerValue(pList("MaximumOnWaitingList")) > 0 Then
              epl.SetErrorField("MaximumOnWaitingList", InformationMessages.ImNoExternalWaitingList)
              pValid = False
            End If
          End If
          Dim vBookingsClose As Date
          Dim vEndDate As Date
          If Date.TryParse(pList("BookingsClose"), vBookingsClose) And Date.TryParse(pList("EndDate"), vEndDate) Then
            If vBookingsClose > vEndDate Then
              epl.SetErrorField("BookingsClose", InformationMessages.ImBookingsCloseBeforeEnd)
              pValid = False
            End If
          End If
          If pList("EligibilityCheckRequired") = "Y" Then
            If pList("EligibilityCheckText").Length = 0 Then
              epl.SetErrorField("EligibilityCheckText", InformationMessages.ImEligibilityTextRequired)
              pValid = False
            End If
            If pList("DeferredBookingAct").Length = 0 Then
              epl.SetErrorField("DeferredBookingAct", InformationMessages.ImDeferredActRequired)
              pValid = False
            End If
            If pList("DeferredBookingActValue").Length = 0 Then
              epl.SetErrorField("DeferredBookingActValue", InformationMessages.ImDeferredActValueRequired)
              pValid = False
            End If
            If pList("RejectedBookingAct").Length = 0 Then
              epl.SetErrorField("RejectedBookingAct", InformationMessages.ImRejectedActRequired)
              pValid = False
            End If
            If pList("RejectedBookingActValue").Length = 0 Then
              epl.SetErrorField("RejectedBookingActValue", InformationMessages.ImRejectedActValueRequired)
              pValid = False
            End If
          End If
          If pList("Booking") = "Y" Then
            If pList.ContainsKey("EventPricingMatrix") AndAlso pList("EventPricingMatrix").Length > 0 Then
              If mvEventInfo.Booking = False AndAlso mvEventInfo.EventPricingMatrix = pList("EventPricingMatrix") Then
                'We have set the Bookings Allowed flag, if the Event has just been duplicated then the Matrix may be invalid
                ValidateItem(pSender, "EventPricingMatrix", pList("EventPricingMatrix"), pValid)
              End If
            End If
          End If
          vTemplate = pList("Template") = "Y"
          vNumberOfAttendees = IntegerValue(epl.GetValue("NumberOfAttendees"))
          vNumberOnWaitingList = IntegerValue(epl.GetValue("NumberOnWaitingList"))
        Else
          'Session only checks
          vTemplate = mvEventInfo.Template
          If dgr.DataRowCount > 0 AndAlso pList.ContainsKey("SessionNumber") Then
            vNumberOfAttendees = IntegerValue(dgr.GetValue(dgr.CurrentRow, "NumberOfAttendees"))
            vNumberOnWaitingList = IntegerValue(dgr.GetValue(dgr.CurrentRow, "NumberOnWaitingList"))
          Else
            vNumberOfAttendees = 0
            vNumberOnWaitingList = 0
          End If
          Dim vStartDate As New DateHelper(epl.GetValue("StartDate"), epl.GetValue("StartTime"))
          If vStartDate.DateValue < mvEventInfo.StartDateHelper.DateValue Then
            Dim vParameter As String
            If vStartDate.DateValue.Date = mvEventInfo.StartDateHelper.DateValue.Date Then
              vParameter = "StartTime"
            Else
              vParameter = "StartDate"
            End If
            epl.SetErrorField(vParameter, InformationMessages.ImInvalidSessionStart)
            pValid = False
          End If
          Dim vEndDate As New DateHelper(epl.GetValue("EndDate"), epl.GetValue("EndTime"))
          If vEndDate.DateValue > mvEventInfo.EndDateHelper.DateValue Then
            Dim vParameter As String
            If vEndDate.DateValue.Date = mvEventInfo.EndDateHelper.DateValue.Date Then
              vParameter = "EndTime"
            Else
              vParameter = "EndDate"
            End If
            epl.SetErrorField(vParameter, InformationMessages.ImInvalidSessionEnd)
            pValid = False
          End If
        End If
        If vTemplate AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_we_eve_from_template) = False Then
          Dim vStart As Date
          Date.TryParse(pList("StartDate"), vStart)
          If vStart.DayOfWeek = DayOfWeek.Saturday Or vStart.DayOfWeek = DayOfWeek.Sunday Then
            epl.SetErrorField("StartDate", InformationMessages.ImTemplateStartWeekday)
            pValid = False
          End If
          Dim vEnd As Date
          Date.TryParse(pList("EndDate"), vEnd)
          If vEnd.DayOfWeek = DayOfWeek.Saturday Or vEnd.DayOfWeek = DayOfWeek.Sunday Then
            epl.SetErrorField("EndDate", InformationMessages.ImTemplateEndWeekday)
            pValid = False
          End If
        End If
        If pList.IntegerValue("MaximumAttendees") < pList.IntegerValue("MinimumAttendees") Then
          epl.SetErrorField("MaximumAttendees", InformationMessages.ImMaxAttendeesGTMinimum)
          pValid = False
        End If
        If pList.IntegerValue("MaximumAttendees") < vNumberOfAttendees Then
          epl.SetErrorField("MaximumAttendees", InformationMessages.ImMaxAttendeesLTBookings)
          pValid = False
        End If
        If pList.IntegerValue("MaximumOnWaitingList") < vNumberOnWaitingList Then
          epl.SetErrorField("MaximumOnWaitingList", InformationMessages.ImMaxWaitingLTWaitingList)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventResults
        Dim vRow As DataRowView = DirectCast(epl.FindComboBox("TestNumber").SelectedItem, DataRowView)
        Dim vValue As String = epl.GetValue("TestResult")
        Dim vPattern As String = vRow.Item("Pattern").ToString
        Dim vMin As String = vRow.Item("MinimumValue").ToString
        Dim vMax As String = vRow.Item("MaximumValue").ToString
        Select Case vRow.Item("GradeDataType").ToString
          Case "C"
            If vPattern.Length > 0 Then
              If vPattern.IndexOf(vValue) < 0 Then pValid = False
              If Not pValid Then epl.SetErrorField("TestResult", GetInformationMessage(InformationMessages.ImResultLike, vPattern))
            Else
              If vMin.Length > 0 AndAlso vValue < vMin Then pValid = False
              If vMax.Length > 0 AndAlso vValue > vMax Then pValid = False
              If Not pValid Then epl.SetErrorField("TestResult", GetInformationMessage(InformationMessages.ImResultRange, vMin, vMax))
            End If
          Case "I"
            If IntegerValue(vValue) <> DoubleValue(vValue) Then pValid = False
            If vMin.Length > 0 AndAlso IntegerValue(vValue) < IntegerValue(vMin) Then pValid = False
            If vMax.Length > 0 AndAlso IntegerValue(vValue) > IntegerValue(vMax) Then pValid = False
            If Not pValid Then epl.SetErrorField("TestResult", GetInformationMessage(InformationMessages.ImResultIntegerRange, vMin, vMax))
          Case Else
            If vMin.Length > 0 AndAlso DoubleValue(vValue) < DoubleValue(vMin) Then pValid = False
            If vMax.Length > 0 AndAlso DoubleValue(vValue) > DoubleValue(vMax) Then pValid = False
            If Not pValid Then epl.SetErrorField("TestResult", GetInformationMessage(InformationMessages.ImResultRange, vMin, vMax))
        End Select
      Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
        If Date.Parse(epl.GetValue("ToDate")).Subtract(Date.Parse(epl.GetValue("FromDate"))).Days < 1 Then
          epl.SetErrorField("ToDate", InformationMessages.ImMustSelectOneNight)
          pValid = False
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventSessionTests
        Dim vMin As String = pList("MinimumValue")
        Dim vMax As String = pList("MaximumValue")
        If vMin.Length = 0 And vMax.Length = 0 And pList("Pattern").Length = 0 Then
          epl.SetErrorField("MinimumValue", InformationMessages.ImRangeOrPattern)
          pValid = False
        Else
          If (vMin.Length > 0 And vMax.Length = 0) Or (vMax.Length > 0 And vMin.Length = 0) Then
            epl.SetErrorField("MinimumValue", InformationMessages.ImCompleteRange)
            pValid = False
          End If
        End If
        If pValid Then
          Select Case pList("GradeDataType")
            Case "C"
              If vMin > vMax Then
                epl.SetErrorField("MinimumValue", InformationMessages.ImMinimumLTMax)
                pValid = False
              End If
            Case Else
              If DoubleValue(vMin) > DoubleValue(vMax) Then
                epl.SetErrorField("MinimumValue", InformationMessages.ImMinimumLTMax)
                pValid = False
              ElseIf pList("GradeDataType") = "I" Then
                If IntegerValue(vMin) <> DoubleValue(vMin) Or
                   IntegerValue(vMax) <> DoubleValue(vMax) Then
                  epl.SetErrorField("MinimumValue", InformationMessages.ImMinAndMaxWholeNumbers)
                End If
              End If
          End Select
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
        If IntegerValue(epl.GetValue("Quantity")) > mvOptionMaximumBookings Then
          epl.SetErrorField("Quantity", GetInformationMessage(InformationMessages.ImQtyGTMaxBookings, mvOptionMaximumBookings.ToString))
          pValid = False
        ElseIf IntegerValue(epl.GetValue("Quantity")) < mvOptionMinimumBookings Then
          epl.SetErrorField("Quantity", GetInformationMessage(InformationMessages.ImQtyLTMinBookings, mvOptionMinimumBookings.ToString))
          pValid = False
        End If

      Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
        If DoubleValue(epl.GetValue("CpdPoints")) + DoubleValue(epl.GetValue("CpdPoints2")) = 0 Then
          epl.SetErrorField("CpdPoints", InformationMessages.ImCPDPointsTotalCannotBeZero)
          pValid = False
        End If

    End Select
  End Sub
  Private Sub ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings, CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
        Select Case pParameterName
          Case "AddressNumber"
            Dim vRow As DataRow = Nothing
            If IntegerValue(pValue) > 0 Then
              vRow = DataHelper.GetRowFromDataSet(DataHelper.GetAddressData(CareServices.XMLAddressDataSelectionTypes.xadtOrganisationFromAddress, IntegerValue(pValue)))
            End If
            Dim vCurrentOrg As String = epl.GetValue("PositionOrganisation")
            If vRow IsNot Nothing Then
              Dim vOrgNumber As String = vRow.Item("OrganisationNumber").ToString
              If vOrgNumber <> vCurrentOrg Then epl.SetValue("PositionOrganisation", vOrgNumber, False, False, True, True)
            ElseIf vCurrentOrg <> "" Then
              epl.SetValue("PositionOrganisation", "", False, False, True, True)
            End If

          Case "PositionOrganisation"
            epl.SetValue("ContactNumber", pValue)
            Debug.Print("Position Organisation Changed")

          Case "OptionNumber"
            Dim vRow As DataRowView = DirectCast(epl.FindComboBox(pParameterName).SelectedItem, DataRowView)
            mvPickSessions = vRow.Item("PickSessions").ToString = "Y"
            mvNumberOfSessions = IntegerValue(vRow.Item("NumberOfSessions").ToString)
            mvOptionDesc = vRow.Item("OptionDesc").ToString
            mvOptionMinimumBookings = IntegerValue(vRow.Item("MinimumBookings").ToString)
            mvOptionMaximumBookings = IntegerValue(vRow.Item("MaximumBookings").ToString)
            Dim vTable As DataTable = Nothing
            If vRow.Item("ProductCode").ToString.Length > 0 Then
              Dim vList As New ParameterList(True)
              vList("Product") = vRow.Item("ProductCode").ToString
              If epl.GetValue("ContactNumber") <> "" Then vList("ContactNumber") = epl.GetValue("ContactNumber")
              vList("EventNumber") = mvEventInfo.EventNumber.ToString
              vList("BookingDate") = AppValues.TodaysDate
              vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRates, vList)
            End If

            Dim vRateCode As String = epl.GetValue("Rate")
            If vTable IsNot Nothing Then
              If vTable.Rows.Count = 1 Then
                vRateCode = vTable.Rows(0).Item("Rate").ToString
              ElseIf vTable.Rows.Count > 1 Then
                'if more than one rate exists for the booking option product then get the rate from the event booking option
                vRateCode = vRow.Item("RateCode").ToString
              End If
            End If
            epl.SetComboDataSource("Rate", "Rate", "RateDesc", vTable)
            If Not String.IsNullOrEmpty(vRateCode) Then
              epl.SetValue("Rate", vRateCode)
            End If

          Case "RoomType"
            Dim vRoomsAt As New CollectionList(Of LookupItem)
            If mvRoomsTable IsNot Nothing Then
              For Each vRow As DataRow In mvRoomsTable.Rows
                If vRow.Item("RoomType").ToString = pValue AndAlso Not vRoomsAt.ContainsKey(vRow.Item("BlockBookingNumber").ToString) Then
                  Dim vItem As New LookupItem(vRow.Item("BlockBookingNumber").ToString, vRow.Item("Organisation").ToString)
                  vRoomsAt.Add(vItem.LookupCode, vItem)
                End If
              Next
              epl.SetComboDataSource("BlockBookingNumber", vRoomsAt)
            End If
          Case "BlockBookingNumber"     'Rooms At
            If mvRoomsTable IsNot Nothing Then
              Dim vRoomType As String = epl.GetValue("RoomType")
              For Each vRow As DataRow In mvRoomsTable.Rows
                If vRow.Item("BlockBookingNumber").ToString = pValue Then
                  Dim vList As New ParameterList(True)
                  vList("Product") = vRow.Item("ProductCode").ToString
                  epl.SetComboDataSource("Rate", "Rate", "RateDesc", DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRates, vList))
                  epl.SetValue("Rate", vRow.Item("RateCode").ToString)
                End If
              Next
            End If
        End Select
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates, CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations
        If pParameterName = "AddressNumber" Then
          Dim vList As New ParameterList(True)
          vList("AddressNumber") = pValue
          Dim vDisable As Boolean = (mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookingAllocations)
          Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressPositionAndOrganisation, IntegerValue(epl.GetValue("ContactNumber")), vList))
          If vRow IsNot Nothing Then
            epl.SetValue("Position", vRow.Item("Position").ToString, vDisable)
            epl.SetValue("OrganisationName", vRow.Item("Name").ToString, vDisable)
          Else
            epl.SetValue("Position", "", vDisable)
            epl.SetValue("OrganisationName", "", vDisable)
          End If
          If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates Then
            If FindControl(epl, "StandardPosition").Visible = True Then
              epl.SetValue("StandardPosition", epl.GetValue("Position"))
            End If
          End If
        End If

      Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation
        Select Case pParameterName
          Case "EndDate", "EndTime", "StartDate", "StartTime", "SkillLevel", "Venue"
            If FindControl(epl, "EventPricingMatrix", False) IsNot Nothing Then
              Dim vEPM As String = epl.GetValue("EventPricingMatrix")
              If epl.FindTextLookupBox("EventPricingMatrix").Enabled = True Then
                If (pParameterName = "EndDate" OrElse pParameterName = "StartDate") AndAlso vEPM.Length > 0 Then
                  ValidateItem(epl, "EventPricingMatrix", vEPM, True)
                ElseIf pParameterName = "Venue" Then
                  epl.FindTextLookupBox("EventPricingMatrix").SetFilter(GetPricingMatrixFilter(pValue))
                  epl.SetValue("EventPricingMatrix", vEPM)    'Reset the original value as it may have been cleared
                  If epl.FindTextLookupBox("EventPricingMatrix").IsValid = False Then vEPM = ""
                  If pParameterName = "Venue" AndAlso pValue.Length > 0 Then
                    'Default Pricing Matrix
                    If vEPM.Length = 0 Then
                      vEPM = DefaultPricingMatrix(pValue)
                      epl.FindTextLookupBox("EventPricingMatrix").SetFilter(GetPricingMatrixFilter(pValue))   'Filtering has been changed
                      If vEPM.Length > 0 Then
                        epl.SetValue("EventPricingMatrix", vEPM)
                        ValidateItem(epl, "EventPricingMatrix", vEPM, True)
                      End If
                    End If
                  End If
                End If
              End If
              mvBookingPriceChange = True
            End If
        End Select

      Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
        Select Case pParameterName
          Case "CpdApprovalStatus"
            If pValue.Length > 0 Then
              Dim vRow As DataRow = epl.FindTextLookupBox("CpdApprovalStatus").GetDataRow()
              If vRow IsNot Nothing Then
                Dim vDateMandatory As Boolean = BooleanValue(vRow.Item("CpdApprovalDateRequired").ToString)
                epl.PanelInfo.PanelItems("CpdDateApproved").Mandatory = vDateMandatory
                If vDateMandatory = True AndAlso epl.GetValue("CpdDateApproved").Length = 0 Then
                  epl.SetErrorField("CpdDateApproved", InformationMessages.ImApprovalDateRequired)
                End If
              End If
            End If

          Case "CpdCategory"
            AppHelper.CPDCategoryValueChanged(epl, False)

          Case "CpdCategoryType"
            AppHelper.CPDCategoryTypeValueChanged(epl, EditingExistingRecord(), SessionNumber)

        End Select
    End Select
  End Sub

  Private Sub GetAddressRestrictionsHandler(ByVal sender As Object, ByRef pFilter As String) Handles epl.GetAddressRestrictions
    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_hide_historic_addresses, False) Then
      Select Case mvEventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings, CareServices.XMLEventDataSelectionTypes.xedtEventAttendees, CareServices.XMLEventDataSelectionTypes.xedtEventCurrentAttendees
          If mvEditing = False Then pFilter = "Historical <> 'Yes'"
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
          Dim vTextLookup As TextLookupBox = TryCast(sender, TextLookupBox)
          If vTextLookup IsNot Nothing Then
            If vTextLookup.OriginalText <> vTextLookup.Text Then
              pFilter = "Historical <> 'Yes'"
            End If
          End If
      End Select
    End If
  End Sub

  Private Sub ValidateItem(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean) Handles epl.ValidateItem

    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates
        If pParameterName = "ContactNumber" Then
          Dim vList As New ParameterList(True)
          'BR11718 - Need to populate the position and organisation values dependent on the config setting
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.default_delegate_position, False) Then
            ' Populate where we only have a single position for the contact, otherwise as normal
            Dim vTableCP As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, IntegerValue(epl.GetValue("ContactNumber")), vList))
            If vTableCP IsNot Nothing Then
              vTableCP.DefaultView.RowFilter = "Current='Yes'"
              If vTableCP.DefaultView.Count = 1 Then
                epl.SetValue("Position", vTableCP.DefaultView.Item(0).Item("Position").ToString)
                epl.SetValue("OrganisationName", vTableCP.DefaultView.Item(0).Item("ContactName").ToString)
                SetStandardPositions()
              End If
            End If
          End If
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventOrganiser
        If pParameterName = "Organiser" Then SetOrganiserContactInformation()
      Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation, CareServices.XMLEventDataSelectionTypes.xedtEventVenueBookings
        If pParameterName = "Venue" Then
          SetVenueInformation()
        ElseIf pParameterName = "EventPricingMatrix" Then
          Dim vRow As DataRow = epl.FindTextLookupBox("EventPricingMatrix").GetDataRow()
          If vRow IsNot Nothing AndAlso IsDate(vRow.Item("EventFeeStartDate").ToString) Then
            Dim vEventStart As Date = Date.Parse(epl.GetValue("StartDate"))
            Dim vFeeStart As Date = Date.Parse(vRow.Item("EventFeeStartDate").ToString)
            Dim vFeeEnd As Date = Date.Parse(vRow.Item("EventFeeEndDate").ToString)
            If Not (DateDiff(DateInterval.Day, vEventStart, vFeeStart) <= 0 AndAlso DateDiff(DateInterval.Day, vFeeEnd, vEventStart) <= 0) Then
              'EventPricingMatrix starts after Event Start or ends before Event Start
              pValid = epl.SetErrorField("EventPricingMatrix", InformationMessages.ImEventPricingMatrixInvalid, True)
            Else
              epl.SetErrorField("EventPricingMatrix", "")
              pValid = True
            End If
          End If
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings, CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation
        Select Case pParameterName
          Case "FromDate", "ToDate"
            SetNumberOfNights()
        End Select

      Case CareNetServices.XMLEventDataSelectionTypes.xedtEventSessionCPD
        Select Case pParameterName
          Case "CpdCategory"
            If AppHelper.IsCPDCategoryTypeAndCategoryAlreadyUsed(epl, dgr, EditingExistingRecord()) Then
              pValid = False
              epl.SetErrorField("CpdCategory", InformationMessages.ImCPDSessionCategoryAlreasyExists, True)
            End If
        End Select

    End Select
  End Sub

  Private Function SetNumberOfNights() As Integer
    Dim vNumberOfNights As Integer = Date.Parse(epl.GetValue("ToDate")).Subtract(Date.Parse(epl.GetValue("FromDate"))).Days
    epl.SetValue("NumberOfNights", vNumberOfNights.ToString)
    Return vNumberOfNights
  End Function

  Private Sub SetOrganiserContactInformation()
    Dim vTextBox As TextLookupBox = epl.FindTextLookupBox("Organiser")
    Dim vRow As DataRow = vTextBox.GetDataRow
    If vRow IsNot Nothing Then
      Dim vItems() As String = {"OrganiserContactName", "OrganiserContactAddressLine", "InvoiceContactName", "InvoiceContactAddressLine"}
      For Each vItem As String In vItems
        epl.SetValue(vItem, vRow.Item(vItem).ToString)
      Next
    End If
  End Sub

  Private Sub ControlDoubleClick(ByVal pSender As Object, ByVal pParameterName As String) Handles epl.ControlDoubleClick
    Dim vContactNumber As Integer
    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventOrganiser
        Dim vTextBox As TextLookupBox = epl.FindTextLookupBox("Organiser")
        Dim vRow As DataRow = vTextBox.GetDataRow
        If vRow IsNot Nothing Then
          Select Case pParameterName
            Case "InvoiceContactName"
              vContactNumber = IntegerValue(vRow.Item("InvoiceContactNumber").ToString)
            Case "OrganiserContactName"
              vContactNumber = IntegerValue(vRow.Item("OrganiserContactNumber").ToString)
          End Select
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventVenueBookings
        Dim vTextBox As TextLookupBox = epl.FindTextLookupBox("Venue")
        Dim vRow As DataRow = vTextBox.GetDataRow
        If vRow IsNot Nothing Then
          Select Case pParameterName
            Case "VenueContactName"
              vContactNumber = IntegerValue(vRow.Item("VenueContactNumber").ToString)
            Case "VenueOrganisationName"
              vContactNumber = IntegerValue(vRow.Item("VenueOrganisationNumber").ToString)
          End Select
        End If
    End Select
    If vContactNumber > 0 Then FormHelper.ShowContactCardIndex(vContactNumber)
  End Sub

  Private Sub SetVenueInformation()
    Dim vTextBox As TextLookupBox = epl.FindTextLookupBox("Venue")
    Dim vRow As DataRow = vTextBox.GetDataRow
    If vRow IsNot Nothing Then
      Dim vItems As New List(Of String)
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventInformation Then
        vItems.Add("Location")
      Else
        vItems.Add("VenueOrganisationName")
        vItems.Add("VenueAddress")
        vItems.Add("VenueContactName")
        vItems.Add("VenueTelephone")
        vItems.Add("VenueCapacity")
      End If
      For Each vItem As String In vItems
        If vItem = "VenueCapacity" Then
          If epl.Controls.ContainsKey("VenueCapacity") Then epl.SetValue(vItem, vRow.Item(vItem).ToString)
        Else
          epl.SetValue(vItem, vRow.Item(vItem).ToString)
        End If
      Next
    End If
  End Sub

  Private Const ebsBooked As String = "F"
  Private Const ebsWaiting As String = "W"
  Private Const ebsBookedTransfer As String = "X"
  Private Const ebsBookedAndPaid As String = "B"
  Private Const ebsWaitingPaid As String = "P"
  Private Const ebsBookedAndPaidTransfer As String = "Y"
  Private Const ebsBookedCreditSale As String = "S"
  Private Const ebsWaitingCreditSale As String = "A"
  Private Const ebsBookedCreditSaleTransfer As String = "R"
  Private Const ebsBookedInvoiced As String = "V"
  Private Const ebsWaitingInvoiced As String = "O"
  Private Const ebsBookedInvoicedTransfer As String = "D"
  Private Const ebsExternal As String = "E"
  Public Const ebsCancelled As String = "C"
  Private Const ebsInterested As String = "I"
  Private Const ebsAwaitingAcceptance As String = "T"
  Private Const ebsAmended As String = "U"

  Private Sub GetAvailableRoomTypes()
    Dim vList As New ParameterList(True)
    vList("SystemColumns") = "N"
    mvRoomsTable = DataHelper.GetTableFromDataSet(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventAccommodation, mvEventInfo.EventNumber, vList))
    Dim vRoomTypes As New CollectionList(Of LookupItem)
    If mvRoomsTable IsNot Nothing Then
      vRoomTypes.Add("", New LookupItem("", ""))
      For Each vRow As DataRow In mvRoomsTable.Rows
        If Not vRoomTypes.ContainsKey(vRow.Item("RoomType").ToString) Then
          Dim vItem As New LookupItem(vRow.Item("RoomType").ToString, vRow.Item("RoomTypeDesc").ToString)
          vRoomTypes.Add(vItem.LookupCode, vItem)
        End If
      Next
      epl.SetComboDataSource("RoomType", vRoomTypes)
    End If
  End Sub

  Private Sub SetBookingStatusRestriction(ByVal pStatus As String, ByVal pBatchNumber As String)
    Dim vComboBox As ComboBox = epl.FindComboBox("BookingStatus")
    Dim vTable As DataTable
    If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings Then
      vTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtAccomodationBookingStatuses)
    Else
      vTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtEventBookingStatuses)
    End If
    Dim vValidStatuses As New ArrayListEx
    If pStatus.Length = 0 Then
      If mvEventDataType = CareServices.XMLEventDataSelectionTypes.xedtEventBookings Then
        vValidStatuses.Add(ebsInterested)
        If mvEventInfo.External Then
          vValidStatuses.Add(ebsExternal)
        Else
          If mvEventInfo.MaximumOnWaitingList > 0 AndAlso mvEventInfo.NumberOnWaitingList < mvEventInfo.MaximumOnWaitingList Then vValidStatuses.Add(ebsWaiting)
          If mvEventInfo.MaximumAttendees > mvEventInfo.NumberOfAttendees Then vValidStatuses.Add(ebsBooked)
        End If
      Else
        vValidStatuses.Add(ebsBooked)
      End If
    Else
      vValidStatuses.Add(pStatus)
    End If
    Dim vCanCancel As Boolean = False
    Select Case pStatus
      Case ebsBooked
        If Not (mvEventInfo.FreeOfCharge Or mvEventInfo.HasFocBookingOption) Then vValidStatuses.Add(ebsBookedAndPaid)
        vCanCancel = True
      Case ebsExternal, ebsCancelled
        '
      Case ebsWaiting
        If Not (mvEventInfo.FreeOfCharge Or mvEventInfo.HasFocBookingOption) Then vValidStatuses.Add(ebsWaitingPaid)
        vCanCancel = True
      Case ebsBookedTransfer
        If Not (mvEventInfo.FreeOfCharge Or mvEventInfo.HasFocBookingOption) Then vValidStatuses.Add(ebsBookedAndPaidTransfer)
        vCanCancel = True
      Case ebsBookedAndPaid
        If IntegerValue(pBatchNumber) = 0 Then
          vValidStatuses.Add(ebsBooked)
          vCanCancel = True
        End If
      Case ebsWaitingPaid
        vCanCancel = IntegerValue(pBatchNumber) = 0
      Case ebsBookedAndPaidTransfer
        If IntegerValue(pBatchNumber) = 0 Then
          vValidStatuses.Add(ebsBookedTransfer)
          vCanCancel = True
        End If
      Case ebsBookedCreditSale
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsBookedInvoiced)
          vValidStatuses.Add(ebsBookedAndPaid)
        End If
      Case ebsWaitingCreditSale
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsWaitingInvoiced)
          vValidStatuses.Add(ebsWaitingPaid)
        End If
      Case ebsBookedCreditSaleTransfer
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsBookedInvoicedTransfer)
          vValidStatuses.Add(ebsBookedAndPaidTransfer)
        End If
      Case ebsBookedInvoiced
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsBookedAndPaid)
        End If
      Case ebsWaitingInvoiced
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsWaitingPaid)
        End If
      Case ebsBookedInvoicedTransfer
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_use_sales_ledger) = False Then
          vValidStatuses.Add(ebsBookedAndPaidTransfer)
        End If
      Case ebsInterested, ebsAwaitingAcceptance
        If pStatus = ebsInterested Then vValidStatuses.Add(ebsAwaitingAcceptance)
        vValidStatuses.Add(ebsCancelled)
        If mvEventInfo.External Then
          vValidStatuses.Add(ebsExternal)
        Else
          If (mvEventInfo.FreeOfCharge Or mvEventInfo.HasFocBookingOption) Then vValidStatuses.Add(ebsBooked)
        End If
    End Select
    If vCanCancel AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciCancelEventBooking) Then vValidStatuses.Add(ebsCancelled)
    Dim vItems As New List(Of LookupItem)
    For Each vRow As DataRow In vTable.Rows
      For Each vItem As String In vValidStatuses
        If vRow.Item("LookupCode").ToString = vItem Then
          vItems.Add(New LookupItem(vRow.Item("LookupCode").ToString, vRow.Item("LookupDesc").ToString))
        End If
      Next
    Next
    vComboBox.DataSource = vItems
    If pStatus.Length = 0 Then SelectComboBoxItem(vComboBox, ebsBooked)
  End Sub

  Private Function GetPricingMatrixFilter(ByVal pVenue As String) As String
    If pVenue.Length = 0 Then pVenue = epl.GetValue("Venue")
    Dim vFilter As String = "Venue = ''"
    If pVenue.Length > 0 Then vFilter &= " Or Venue = '{0}'"

    Return String.Format(vFilter, pVenue)
  End Function

  Private Function DefaultPricingMatrix(ByVal pVenue As String) As String
    Dim vDefault As String = ""
    Dim vFilter As String = "Venue = '{0}'"
    If pVenue.Length > 0 Then
      Dim vDT As DataTable = epl.FindTextLookupBox("EventPricingMatrix").GetDataRow.Table
      vDT.DefaultView.RowFilter = String.Format(vFilter, pVenue)
      Dim vFDT As DataTable = vDT.DefaultView.ToTable
      If vFDT.Rows.Count > 0 Then
        vDefault = vFDT.Rows(0).Item("EventPricingMatrix").ToString
      Else
        vDT.DefaultView.RowFilter = String.Format(vFilter, "") & " AND EventPricingMatrix <> ''"
        Dim vFDT2 As DataTable = vDT.DefaultView.ToTable
        If vFDT2.Rows.Count > 0 Then vDefault = vFDT2.Rows(0).Item("EventPricingMatrix").ToString
      End If
    End If
    Return vDefault
  End Function

  Protected Overrides Sub frmCardMaintenance_Load(ByVal sender As Object, ByVal e As System.EventArgs)
    If Me.DesignMode Then Return
    If mvStandAlone AndAlso mvStandAloneParent IsNot Nothing Then     'Called from Trader
      If mvStandAloneParent.MdiParent IsNot Nothing And Me.MdiParent Is Nothing Then
        Location = FormHelper.MDIPointToScreen(mvStandAloneParent.Location)
        SetStandardPositions()              'Setting the value for StandardPosition combobox 
      Else
        Location = mvStandAloneParent.Location
      End If
      Size = mvStandAloneParent.Size              'Required here for Windows 2000
    Else
      If mvParentForm IsNot Nothing AndAlso mvParentForm.SizeMaintenanceForm Then
        Location = mvParentForm.Location
        Size = mvParentForm.Size              'Required here for Windows 2000
        mvParentForm.Enabled = False
      Else
        If mvParentForm IsNot Nothing Then mvParentForm.Enabled = False
        If MdiParent Is Nothing AndAlso MDIForm IsNot Nothing Then
          Location = MDIForm.PointToScreen(MDILocation(Width, Height))
        Else
          Location = MDILocation(Width, Height)
        End If
      End If
      mvSelectedRow = -1
    End If
    bpl.RepositionButtons()
  End Sub

  Private Sub cboData_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboData.SelectedIndexChanged
    If cboData.SelectedIndex >= 0 Then
      If CInt(cboData.SelectedValue) <> SessionNumber Then
        Dim vCurrentSession As Integer = SessionNumber
        SessionNumber = CInt(cboData.SelectedValue)
        If vCurrentSession > 0 Then RefreshCard()
      End If
    End If
  End Sub

  Private Sub epl_ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Handles epl.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub epl_PopulateDisplayGrid(ByVal pSender As Object, ByVal pDGR As DisplayGrid, ByVal pPanelItem As PanelItem) Handles epl.PopulateDisplayGrid
    Select Case mvEventDataType
      Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees
        Dim vList As New ParameterList(True)
        vList("ContactNumber") = dgr.GetValue(dgr.CurrentRow, "ContactNumber")
        If pPanelItem.ParameterName = "Mailings" Then
          vList("EventNumber") = mvEventInfo.EventNumber.ToString
          pDGR.Populate(DataHelper.GetEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventDelegateMailings, mvEventInfo.EventNumber, vList))
        Else
          pDGR.Populate(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventDelegateIncome, mvEventInfo.EventNumber, vList))
        End If
      Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
        Dim vList As New ParameterList(True)
        vList("EventPersonnelNumber") = dgr.GetValue(dgr.CurrentRow, "EventPersonnelNumber")
        pDGR.Populate(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventPersonnelTasks, mvEventInfo.EventNumber, vList))
      Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
        Dim vList As New ParameterList(True)
        vList("BookingNumber") = dgr.GetValue(dgr.CurrentRow, "BookingNumber")
        pDGR.Populate(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookingTransactions, mvEventInfo.EventNumber, vList))
    End Select
  End Sub

  Private Sub dgr_PrintParameters(ByVal pSender As Object, ByRef pJobName As String) Handles dgr.GetPrintParameters
    Dim vPrintCaption As New StringBuilder
    If mvEventInfo IsNot Nothing Then
      vPrintCaption.Append(mvEventInfo.EventDescription)
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

  Private Sub dgr_DocumentSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDocumentNumber As Integer) Handles dgr.DocumentSelected
    Try
      FormHelper.EditDocument(pDocumentNumber, Me, Nothing)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmEventSet_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs)
    cboData.DataSource = Nothing
    If mvEventMenu IsNot Nothing Then mvEventMenu.Dispose()
    If mvCustomiseMenu IsNot Nothing Then mvCustomiseMenu.Dispose()
    If mvActionMenu IsNot Nothing Then mvActionMenu.Dispose()
  End Sub

  Private Sub mvEventMenu_ItemSelected(ByVal pMenuItem As EventMenu.EventMenuItems) Handles mvEventMenu.MenuSelected
    Try
      Select Case pMenuItem
        Case EventMenu.EventMenuItems.emiProcessWaitingList
          Dim vWaiting As New frmWaitingList(mvEventInfo)
          vWaiting.ShowDialog(Me)
          mvEventInfo.RefreshData()
        Case EventMenu.EventMenuItems.emiAllocatePISToEvent
          Dim vList As New ParameterList(True)
          Dim vDefaults As New ParameterList
          vList = New ParameterList     'Clear the current List
          vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptAllocatePISToEvent, vDefaults)
          If vList IsNot Nothing AndAlso vList.Count > 0 Then
            'Add the data
            vList("EventNumber") = mvEventInfo.EventNumber.ToString
            vList("IssueDate") = AppValues.TodaysDate
            Dim vResult As ParameterList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctEventPIS, vList)
            If vResult.Contains("NumberAdded") AndAlso vResult.IntegerValue("NumberAdded") < vList.IntegerValue("NumberToAdd") Then
              ShowInformationMessage(InformationMessages.ImNotAllPISAllocated, vResult("NumberAdded"))
            End If
          End If
        Case EventMenu.EventMenuItems.emiAllocatePISToDelegates
          Dim vList As New ParameterList(True)
          Dim vDefaults As New ParameterList
          vList = New ParameterList     'Clear the current List
          vDefaults("NumberToAdd") = AppValues.ControlValue(AppValues.ControlValues.pis_per_delegate)
          vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptAllocatePISToDelegates, vDefaults)
          If vList IsNot Nothing AndAlso vList.Count > 0 Then
            'Add the data
            vList("EventNumber") = mvEventInfo.EventNumber.ToString
            DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctAllocatePISToDelegates, vList)
          End If
        Case EventMenu.EventMenuItems.emiConfirmDelegates
          Dim vForm As New frmSelectItems(mvEventInfo)
          vForm.ShowDialog()

        Case EventMenu.EventMenuItems.emiDuplicateEvent
          Dim vDefaults As New ParameterList
          vDefaults("EventDesc") = mvEventInfo.EventDescription
          vDefaults("LongDescription") = mvEventInfo.LongDescription
          vDefaults("StartDate") = AppValues.TodaysDate
          Dim vList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateEvent, vDefaults)
          If vList.Count > 0 Then
            vList.Add("Template", CBoolYN(mvEventInfo.Template))
            Dim vResult As ParameterList = DataHelper.DuplicateEvent(mvEventInfo.EventNumber, vList)
            'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
            If vResult.Contains("PricingMatrixValid") AndAlso Not BooleanValue(vResult("PricingMatrixValid").ToString) Then
              ShowWarningMessage(InformationMessages.ImEventPricingMatrixInvalid)
            End If
            If vResult.Contains("EventNumber") Then
              mvEventInfo = New CareEventInfo(vResult.IntegerValue("EventNumber"), mvEventInfo.EventGroup)
              mvEventMenu.EventInfo = mvEventInfo
              mvContainsBookings = mvEventInfo.ContainsBookings
              mvBookingPriceChange = False
              UserHistory.AddEventHistoryNode(mvEventInfo.EventNumber, mvEventInfo.EventName, mvEventInfo.EventGroup)
              RefreshSessions()
              RefreshCard()
            End If
          End If
        Case EventMenu.EventMenuItems.emiNumberCandidates
          Dim vList As New ParameterList(True)
          Dim vEventNumber As Integer = mvEventInfo.EventNumber
          Dim vResult As DialogResult
          vList("EventNumber") = vEventNumber.ToString
          vResult = ShowQuestion(QuestionMessages.QmRenumberCandidates, MessageBoxButtons.YesNoCancel)
          If vResult = System.Windows.Forms.DialogResult.Yes Then
            vList("Renumber") = "Y"
          ElseIf vResult = System.Windows.Forms.DialogResult.No Then
            vList("Renumber") = "N"
          Else
            Exit Select
          End If
          DataHelper.RenumberEventCandidates(vList)
        Case EventMenu.EventMenuItems.emiNumberSessionBookings
          Dim vList As New ParameterList(True)
          Dim vEventNumber As Integer = mvEventInfo.EventNumber
          Dim vResult As DialogResult
          vList("EventNumber") = vEventNumber.ToString
          vResult = ShowQuestion(QuestionMessages.QmRenumberSessionBookings, MessageBoxButtons.YesNoCancel)
          If vResult = System.Windows.Forms.DialogResult.Yes Then
            vList("Renumber") = "Y"
          ElseIf vResult = System.Windows.Forms.DialogResult.No Then
            vList("Renumber") = "N"
          Else
            Exit Select
          End If
          DataHelper.RenumberSessionBookings(vList)
        Case EventMenu.EventMenuItems.emiLoanItems
          Dim vForm As New frmEventLoanItems(mvEventInfo.EventNumber)
          vForm.ShowDialog()
        Case EventMenu.EventMenuItems.emiAuthoriseExpenses
          Dim vForm As New frmAuthoriseExpenses(mvEventInfo.EventNumber)
          vForm.ShowDialog()
        Case EventMenu.EventMenuItems.emiIssueResources
          'TODO Remove NYI from above and un-comment the code once Event Issue Resources implementation is completed.
          Dim vList As New ParameterList(True)
          vList("ReportTotals") = "Y"
          vList("ReportProducts") = "Y"
          vList("IssueResources") = "Y"
          vList("EventNumber") = mvEventInfo.EventNumber.ToString
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtIssueResources, vList, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case EventMenu.EventMenuItems.emiCustomise
          Dim vParams As New ParameterList(True)
          vParams.Add("SelectionPages", "Y")
          vParams("DataSelectionType") = sel.DataSelectionType.ToString
          vParams("ParameterName") = "EventGroup"
          vParams("ParameterValue") = mvEventInfo.EventGroup
          Dim vDisplayList As New frmDisplayList(frmDisplayList.ListUsages.CustomiseDisplayList, vParams)
          If vDisplayList.ShowDialog() = DialogResult.OK Then
            RefreshSelectionPages()
          End If
        Case EventMenu.EventMenuItems.emiRevert
          Dim vParams As New ParameterList(True)
          Try
            If ShowQuestion(QuestionMessages.QmRevertModule, MessageBoxButtons.OKCancel) = DialogResult.OK Then
              vParams.Add("DataSelectionType", sel.DataSelectionType.ToString)
              vParams.Add("AccessMethod", "S")
              vParams.Add("EventGroup", mvEventInfo.EventGroup)
              vParams.Add("Logname", DataHelper.UserInfo.Logname.ToString)
              vParams.Add("Department", DataHelper.UserInfo.Department.ToString)
              vParams.Add("Client", DataHelper.GetClientCode())
              vParams.Add("WebPageItemNumber", "")
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctDisplayListItem, vParams)
              RefreshSelectionPages()
            End If
          Catch vEx As Exception
            DataHelper.HandleException(vEx)
          End Try
      End Select
      RefreshCard(pMenuItem)
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
        Case CareException.ErrorNumbers.enEventSponsorshipInfoMissing,
             CareException.ErrorNumbers.enNotEnoughPIS,
             CareException.ErrorNumbers.enNotAllPISAllocated
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enAppointmentConflict
          ShowInformationMessage(vEx.Message)
        Case Else
          Throw vEx
      End Select
    End Try
  End Sub

  Private Sub mvEventDelegateMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvEventDelegateMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiSupplementaryInformation
          Dim vSource As String
          Dim vContactNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "ContactNumber"))
          Dim vList As New ParameterList(True)
          vList("EventDelegateNumber") = dgr.GetValue(dgr.CurrentRow, "DelegateNumber")
          Dim vDataRow As DataRow = DataHelper.GetContactItem(CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates, vContactNumber, vList)
          If vDataRow IsNot Nothing Then
            Dim vEventDelegateInfo As New EventDelegateInfo(IntegerValue(vList("EventDelegateNumber")), vContactNumber, "", IntegerValue(vDataRow("BatchNumber")) > 0, vDataRow("TransactionSource").ToString)
            If vEventDelegateInfo.BookingTransactionExists Then
              vSource = vEventDelegateInfo.TransactionSource
            Else
              vSource = mvEventInfo.SourceCode
            End If
            Dim vContactInfo As New ContactInfo(vContactNumber)
            ShowDelegateDataSheet(Me, vContactInfo, "D", vSource, mvEventInfo.ActivityGroup, mvEventInfo.RelationshipGroup, vEventDelegateInfo.ContactName, vEventDelegateInfo, False)
          End If
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Private Sub mvEventFinancialLinkMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems) Handles mvEventFinancialLinkMenu.MenuSelected
    Try
      Dim vRefresh As Boolean
      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiGoToTransaction
          Dim vContactNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "ContactNumber"))
          Dim vForm As Form = FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vContactNumber, False)
          If vForm IsNot Nothing Then
            DirectCast(vForm, frmCardSet).SelectTransaction(mvEventFinancialLinkMenu.TargetBatchNumber, mvEventFinancialLinkMenu.TargetTransactionNumber, mvEventFinancialLinkMenu.TargetLineNumber)
          End If
        Case FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink2,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink3,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink4,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink5
          DataHelper.DeleteEventFinancialLink(mvEventInfo.EventNumber, mvEventFinancialLinkMenu.TargetBatchNumber, mvEventFinancialLinkMenu.TargetTransactionNumber, mvEventFinancialLinkMenu.TargetLineNumber)
          vRefresh = True
        Case BaseFinancialMenu.FinancialMenuItems.fmiAmendBooking
          FormHelper.ShowAmendEventBookingForm(Me, dgr, mvEventInfo)
        Case BaseFinancialMenu.FinancialMenuItems.fmiAnalysis
          Dim vList As New ParameterList(True, False)
          vList("SmartClient") = "Y"
          Dim vBatchNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "BatchNumber"))
          Dim vTransactionNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "TransactionNumber"))
          Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vBatchNumber, vTransactionNumber))
          If vRow IsNot Nothing Then
            Dim vContactNumber As Integer = IntegerValue(vRow.Item("ContactNumber").ToString)   'This is the payer Contact & could be different to the booker Contact
            vList.AddSystemColumns()
            vList.IntegerValue("BatchNumber") = vBatchNumber
            vList.IntegerValue("TransactionNumber") = vTransactionNumber
            Dim vPTRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vContactNumber, vList))
            If vPTRow IsNot Nothing Then
              Dim vTransDate As String = vPTRow.Item("TransactionDate").ToString
              Dim vTransSign As String = vPTRow.Item("TransactionSign").ToString
              Dim vStock As Boolean = BooleanValue(vPTRow.Item("ContainsStock").ToString)
              If vStock = False Then vStock = BooleanValue(vPTRow.Item("ContainsPostage").ToString)
              If BooleanValue(dgr.GetValue(dgr.CurrentRow, "CreditSale")) = True AndAlso BooleanValue(dgr.GetValue(dgr.CurrentRow, "InvoicePrinted")) = False Then
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atEventAdjustment, vList, vTransDate, vTransSign, vStock)
              Else
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atAdjustment, vList, vTransDate, vTransSign, vStock)
              End If
              MainHelper.RefreshEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookings, mvEventInfo.EventNumber)
              MainHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, vContactNumber)
            Else
              Throw New CareException(GetInformationMessage(InformationMessages.ImCannotFindFinancialHistoryDetails, vBatchNumber.ToString, vTransactionNumber.ToString), CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails)
            End If
          Else
            Throw New CareException(GetInformationMessage(InformationMessages.ImCannotFindFinancialHistoryDetails, vBatchNumber.ToString, vTransactionNumber.ToString), CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails)
          End If
      End Select
      If vRefresh Then RefreshCard()
    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails Then
        ShowInformationMessage(vCareEX.Message)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel
    epl.ClearDataSources(epl)
    epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing, 0, mvEventInfo.EventGroup))
    epl.FillDeferredCombos(epl)
    RefreshCard()
  End Sub

  Public Overrides ReadOnly Property EventDataType() As CareServices.XMLEventDataSelectionTypes
    Get
      Return mvEventDataType
    End Get
  End Property

  Public Overrides Sub RefreshData()
    RefreshCard()
  End Sub
  Private Sub dgr_CanCustomise(ByVal pSender As Object, ByVal pRow As String) Handles dgr.CanCustomise
    RefreshCard()
  End Sub

  'Private Sub sel_CanCustomise(ByVal Sender As Object, ByVal pResult As String) Handles sel.CanCustomise
  '  Dim vEntityGroup As EntityGroup = DataHelper.EventGroups(mvEventInfo.EventGroup)
  '  vEntityGroup.ResetSelectionPages()
  '  sel.Init(mvEventInfo, True)
  'End Sub

  Public Function GetEventDelegates() As CollectionList(Of EventDelegateInfo)
    Dim vEventDelegates As New CollectionList(Of EventDelegateInfo)

    For vIndex As Integer = 0 To dgr.DataRowCount - 1
      Dim vEventDelegate As New EventDelegateInfo(IntegerValue(dgr.GetValue(vIndex, "EventDelegateNumber")), IntegerValue(dgr.GetValue(vIndex, "ContactNumber")), dgr.GetValue(vIndex, "ContactName"))
      vEventDelegates.Add(dgr.GetValue(vIndex, "EventDelegateNumber"), vEventDelegate)
    Next
    Return vEventDelegates
  End Function

  Public Sub SetStandardPositions()
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventPersonnel Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEventDelegateAllocation Then
      If FindControl(epl, "StandardPosition").Visible = True Then
        epl.SetValue("StandardPosition", epl.GetValue("Position"))
      End If
      epl.DataChanged = False
    End If
  End Sub

  Private Sub RefreshSelectionPages()
    Dim vEntityGroup As EntityGroup = DataHelper.EventGroups(mvEventInfo.EventGroup)
    vEntityGroup.ResetSelectionPages()
    sel.Init(mvEventInfo, True)
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
      'Update the Event to set the MasterAction
      Dim vList As New ParameterList(True, True)
      GetAdditionalKeyValues(vList)
      vList.IntegerValue("MasterAction") = pActionNumber
      mvReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventInformation, vList)
      mvEventInfo.MasterAction = pActionNumber
      mvActionMenu.MasterActionNumber = pActionNumber
      RefreshData()
    End If
  End Sub

  Private Sub mvActionMenu_RefreshCard(ByVal sender As Object) Handles mvActionMenu.RefreshCard
    RefreshCard()
  End Sub

  Private Sub mvHeader_RefreshHeader(ByVal sender As Object) Handles mvHeader.RefreshHeader
    RefreshHeader()
  End Sub

  Public Sub EntitySelected(pSender As Object, pEntityNumber As Integer, Optional pEntityType As HistoryEntityTypes = CDBNETCL.HistoryEntityTypes.hetContacts) Implements IDashboardTabContainer.EntitySelected
    MainHelper.NavigateHistoryItem(pEntityType, pEntityNumber, True)
  End Sub

  Private Sub frmEventSet_Load(sender As Object, e As EventArgs) Handles Me.Load
    SetDefaults()
  End Sub

  Private Property SessionNumber As Integer
    Get
      Return mvSessionNumber
    End Get
    Set(value As Integer)
      mvSessionNumber = value
    End Set
  End Property

End Class
