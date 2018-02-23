Imports System.Runtime.InteropServices



Public Interface IMainForm
  ReadOnly Property MainMenu() As MainMenu
  'Sub ResetToolbar()
  'WriteOnly Property ToolbarTextPosition() As TextImageRelation
End Interface

Public Class MainHelper

#Region "DLL registrations"
  <DllImport("user32.dll")> _
  Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
  End Function

#End Region

#Region "MAIN Sub"

  Private Shared mvMDIMode As Boolean
  Private Shared mvMDIForm As frmMain
  Private Shared mvApplicationContext As MyApplicationContext
  Private Shared mvForms As New List(Of Form)
  Private Shared mvMainForms As New List(Of Form)
  Private Shared mvToolbarTextPosition As TextImageRelation = TextImageRelation.ImageBeforeText
  Private Shared mvImageProvider As New ImageProvider
  Private Shared mvShowSelectionPanel As Boolean = True
  Private Shared mvShowHeaderPanels As Boolean = True
  Private Shared mvShowToolbar As Boolean = True
  Private Shared mvLastContact As ContactInfo
  Private Shared mvPreTaskCount As Integer
  Private Shared mvPreNotificationCount As Integer
  Private Shared WithEvents mvTaskBarIcon As System.Windows.Forms.NotifyIcon
  Friend Shared WithEvents mvTaskBarIconMenu As TaskBarIconMenu
  Friend Shared WithEvents mvTimer As System.Timers.Timer
  Friend Shared WithEvents mvTaskTimer As System.Timers.Timer
  Private Shared mvTaskStatus As String = ""
  Private Shared mvNotificationText As String = ""
  Private Shared mvFormView As FormViews

  Public Shared Sub Start()
    AppValues.DragDropSupported = True
    ShowFinder = New ShowFinderDelegate(AddressOf FormHelper.ShowFinder)
    ShowCardIndex = New ShowCardIndexDelegate(AddressOf FormHelper.ShowCardIndex)
    DoRefreshData = New RefreshDataDelegate(AddressOf RefreshData)
    DoContactCreated = New ContactCreatedDelegate(AddressOf ContactCreated)
    DoEMailDeleted = New EMailDeletedDelegate(AddressOf EMailDeleted)
    DoGetCurrentContact = New GetCurrentContactDelegate(AddressOf GetCurrentContact)
    ShowCriteriaLists = New CriteriaListsDelegate(AddressOf FormHelper.ShowCriteriaLists)
    NavigateItem = New NavigateItemDelegate(AddressOf MainHelper.NavigateHistoryItem)
    DisplayTheme.InitFromSettings()

    For Each vArg As String In AppValues.CommandLineArguments
      If vArg.ToLower.Contains("checkdatabaseversion") Then
        If AppValues.ConfigurationValue(AppValues.ConfigurationValues.last_db_upgrade_version) <> My.Application.Info.Version.ToString(3) Then
          If ShowQuestion("The Database version {0} does not appear to match this version of the software {1}\r\n\r\nDo you want to continue?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2, AppValues.ConfigurationValue(AppValues.ConfigurationValues.last_db_upgrade_version), My.Application.Info.Version.ToString(3)) = DialogResult.No Then
            Exit Sub
          End If
        End If
      End If
    Next


    Dim vForm As Form
    If AppValues.RunAsThames Then
      vForm = New frmDashboard
      'vForm = New frmCardSet
      'DirectCast(vForm, frmCardSet).Init(Nothing, CareServices.XMLContactDataSelectionTypes.xcdtNone, New ContactInfo(ContactInfo.ContactTypes.ctContact, "CON"))
      mvApplicationContext = New MyApplicationContext(vForm)
    Else
      mvMDIMode = True
      mvMDIForm = New frmMain
      mvMDIForm.SuspendLayout()
      mvForms.Add(mvMDIForm)
      CurrentMainForm = mvMDIForm
      mvApplicationContext = New MyApplicationContext(mvMDIForm)
      mvMDIForm.ResumeLayout()
    End If
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
    InitialiseTaskTrayIcon()
    AddHandler PhoneApplication.PhoneInterface.IncomingCall, AddressOf IncomingPhoneCall
    InitServiceLocator()
    Application.Run(mvApplicationContext)
    ClearTaskTrayIcon()
  End Sub

#End Region

#Region "Main Properties and Methods"

  Public Shared ReadOnly Property MainForm() As Form
    Get
      If mvApplicationContext IsNot Nothing Then
        Return mvApplicationContext.MainForm
      Else
        Return Nothing
      End If
    End Get
  End Property

  Public Shared Sub SetMDIParent(ByVal pForm As Form)
    If mvMDIMode Then
      pForm.MdiParent = mvMDIForm
    Else
      mvForms.Add(pForm)
      AddHandler pForm.FormClosed, AddressOf FormClosed
    End If
  End Sub

  Private Shared Sub AddMainForm(ByVal pForm As Form)
    If Not mvMDIMode Then
      mvMainForms.Add(pForm)
      mvForms.Add(pForm)
      AddHandler pForm.Activated, AddressOf MainFormActivated
      AddHandler pForm.FormClosed, AddressOf MainFormClosed
    End If
  End Sub

  Private Shared Sub MainFormActivated(ByVal sender As Object, ByVal e As EventArgs)
    mvApplicationContext.MainForm = DirectCast(sender, Form)
    CurrentMainForm = mvApplicationContext.MainForm
  End Sub

  Private Shared Sub MainFormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
    mvMainForms.Remove(DirectCast(sender, Form))
    mvForms.Remove(DirectCast(sender, Form))
    If mvMainForms.Count > 0 Then
      mvApplicationContext.MainForm = mvMainForms(mvMainForms.Count - 1)
    Else
      AppValues.SaveWindowSizes()
      Settings.ShowNavPanel = NavigationPanel
      If AppHelper.FormView = FormViews.Classic Then
        Settings.ShowToolbar = ShowToolbar
        Settings.ShowStatusBar = StatusBar
      End If
      Settings.ShowHeaderPanel = ShowHeaderPanel
      Settings.ShowSelectionPanel = ShowSelectionPanel
      DirectCast(sender, IMainForm).MainMenu.SaveToolbarItems()
      CDBNETCL.Settings.Save()
      DataHelper.Logout("CD")
      End If
  End Sub

  Private Shared Sub FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
    mvForms.Remove(DirectCast(sender, Form))
  End Sub

  Public Shared Function AddMainMenu(ByVal pForm As Form) As MainMenu
    If mvMDIMode = False OrElse (mvMDIMode AndAlso TypeOf (pForm) Is frmMain) Then
      AddMainForm(pForm)
      Return New MainMenu(pForm)
    Else
      Return Nothing
    End If
  End Function

  Public Shared ReadOnly Property Forms() As Form()
    Get
      If mvMDIMode Then
        Return mvMDIForm.MdiChildren
      Else
        Return mvForms.ToArray
      End If
    End Get
  End Property

  Public Shared Sub CloseAll()
    Dim vForm As Form
    For Each vForm In Forms
      If vForm IsNot MainForm Then vForm.Close()
    Next
  End Sub

  Public Shared Function FindForms(Of FormType As Form)() As List(Of FormType)
    Dim vRtn As New List(Of FormType)
    For Each vOpen As Form In Forms
      If vOpen.GetType() = GetType(FormType) Then
        vRtn.Add(CType(vOpen, FormType))
      End If
    Next
    Return vRtn
  End Function



  Public Shared Sub RefreshData(ByVal pTaskJobType As CareServices.TaskJobTypes)
    Dim vCursor As New BusyCursor
    Try
      Select Case pTaskJobType
        Case CareServices.TaskJobTypes.tjtCashBookPosting, CareServices.TaskJobTypes.tjtPickingList, CareServices.TaskJobTypes.tjtConfirmStockAllocation, _
             CareServices.TaskJobTypes.tjtBatchUpdate, CareServices.TaskJobTypes.tjtPayingInSlips, CareServices.TaskJobTypes.tjtCAFProvisionalBatchClaim, _
             CareServices.TaskJobTypes.tjtCAFCardSalesReport, CareServices.TaskJobTypes.tjtCardSalesFile, CareServices.TaskJobTypes.tjtCCClaimReport, _
             CareServices.TaskJobTypes.tjtCCClaimFile, CareServices.TaskJobTypes.tjtDDClaimFile, CareServices.TaskJobTypes.tjtDDCreditFile, _
             CareServices.TaskJobTypes.tjtCardSalesReport
          For Each vForm As Form In Forms
            Dim vBatchProcessing As frmBatchProcessing = TryCast(vForm, frmBatchProcessing)                     'Refresh frmBatchProcessing 
            If vBatchProcessing IsNot Nothing Then
              vBatchProcessing.RefreshData()
              Exit For
            End If
          Next
      End Select
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub RefreshEventData(ByVal pType As CareServices.XMLEventDataSelectionTypes, ByVal pEventNumber As Integer)
    Dim vCursor As New BusyCursor
    Dim vForm As Form
    Dim vParentForm As MaintenanceParentForm
    Try
      For Each vForm In Forms
        vParentForm = TryCast(vForm, frmEventSet)
        If vParentForm IsNot Nothing Then
          If vParentForm.CareEventInfo.EventNumber = pEventNumber Then
            Select Case pType
              Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
                Select Case vParentForm.EventDataType
                  Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees, _
                       CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates, _
                       CareServices.XMLEventDataSelectionTypes.xedtEventBookings, _
                       CareServices.XMLEventDataSelectionTypes.xedtEventInformation, _
                       CareServices.XMLEventDataSelectionTypes.xedtEventRoomBookings
                    vParentForm.RefreshData()
                End Select
              Case Else
                If vParentForm.EventDataType = pType Then
                  vParentForm.RefreshData()
                End If
            End Select
          End If
        End If
      Next
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub RefreshHistoryData(ByVal pHistoryType As HistoryEntityTypes, ByVal pNumber As Integer)
    Dim vCursor As New BusyCursor
    Try
      Select Case pHistoryType
        Case HistoryEntityTypes.hetSelectionSets    'Currently used in Bulk Contact Deletion
          Dim vDataSet As DataSet = DataHelper.GetSelectionSetData(pNumber)
          UserHistory.UpdateSelectionSetData(vDataSet, pNumber)

          For Each vForm As Form In Forms
            Dim vSSForm As frmSelectionSet = TryCast(vForm, frmSelectionSet)
            If vSSForm IsNot Nothing AndAlso vSSForm.SelectionSetNumber = pNumber Then
              vSSForm.RefreshData(vDataSet)
            Else
              Dim vCardSet As frmCardSet = TryCast(vForm, frmCardSet)
              If vCardSet IsNot Nothing And vDataSet.Tables.Contains("DataRow") Then
                For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
                  If IntegerValue(vRow("ContactNumber")) = vCardSet.ContactInfo.ContactNumber Then vCardSet.Close()
                Next
              End If
            End If
          Next
        Case HistoryEntityTypes.hetContacts         'Currently used in Contact to Organisation Conversion
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetContacts, pNumber)
          For Each vForm As Form In Forms
            Dim vCardSet As frmCardSet = TryCast(vForm, frmCardSet)                     'Refresh any CardSet form
            If vCardSet IsNot Nothing AndAlso vCardSet.ContactInfo.ContactNumber = pNumber Then
              Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, pNumber)
              Dim vContactInfo As New ContactInfo(vDataSet.Tables("DataRow").Rows(0))
              If FormHelper.CheckAccessRights(vContactInfo) Then
                vCardSet.Init(vDataSet, CareServices.XMLContactDataSelectionTypes.xcdtNone, vContactInfo, True)
                vCardSet.BringToFront()
              End If
            End If
            Dim vFinder As frmFinder = TryCast(vForm, frmFinder)                        'Refresh any Finder form
            If vFinder IsNot Nothing AndAlso vFinder.SelectedContactNumber = pNumber Then
              vFinder.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctNone)
            End If
            Dim vNetwork As frmNetworkNew = TryCast(vForm, frmNetworkNew)                     'Refresh any Network form
            'If vNetwork IsNot Nothing AndAlso IntegerValue(vNetwork.tvw.SelectedNode.Tag) = pNumber Then vNetwork.Refresh()
            If vNetwork IsNot Nothing Then
              vNetwork.Refresh()
            End If
          Next
      End Select
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub ContactCreated(ByVal pContactInfo As CDBNETCL.ContactInfo)
    For Each vForm As Form In Forms
      If TypeOf (vForm) Is frmTrader Then
        DirectCast(vForm, frmTrader).HandleContactCreated(pContactInfo)
      End If
    Next
  End Sub

  Public Shared Sub SetControlColors()
    If mvMDIMode = True Then mvMDIForm.SetControlColors()
    For Each vForm As Form In Forms
      If TypeOf (vForm) Is IThemeSettable Then
        CType(vForm, IThemeSettable).SetControlTheme()
      End If
    Next
  End Sub

  Public Shared ReadOnly Property ImageProvider() As ImageProvider
    Get
      Return mvImageProvider
    End Get
  End Property

  Public Shared Property NavigationPanel() As Boolean
    Get
      If mvMDIMode Then
        Return mvMDIForm.NavigationPanel
      Else
        Return True
      End If
    End Get
    Set(ByVal value As Boolean)
      If mvMDIMode Then mvMDIForm.NavigationPanel = value
    End Set
  End Property

  Private Shared mvLastNavigationControl As INavigable

  Public Shared Sub RegisterForNavigation(ByVal pControl As INavigable)
    AddHandler pControl.GotNavigationFocus, AddressOf GotNavigationFocus
    AddHandler pControl.CancelNavigation, AddressOf CancelNavigation
  End Sub

  Public Shared Sub GotNavigationFocus(ByVal sender As INavigable)
    mvLastNavigationControl = sender
  End Sub

  Public Shared Sub CancelNavigation(ByVal sender As INavigable)
    RemoveHandler sender.GotNavigationFocus, AddressOf GotNavigationFocus
    RemoveHandler sender.CancelNavigation, AddressOf CancelNavigation
  End Sub

  Public Shared Sub NavigateToNext()
    If mvLastNavigationControl IsNot Nothing Then
      DirectCast(mvLastNavigationControl, INavigable).NavigateToNext()
    End If
  End Sub

  Public Shared Sub NavigateToPrevious()
    If mvLastNavigationControl IsNot Nothing Then
      DirectCast(mvLastNavigationControl, INavigable).NavigateToPrevious()
    End If
  End Sub

#End Region

#Region "Delegate Functions"

  Public Shared Sub RefreshData(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Integer)
    'ContactNumber 0 used to refresh for all Contacts
    Dim vCursor As New BusyCursor
    Dim vForm As Form
    Dim vParentForm As MaintenanceParent
    Try
      If mvMDIMode Then
        If MainForm IsNot Nothing Then
          For Each vForm In MainForm.MdiChildren
            vParentForm = TryCast(vForm, MaintenanceParent)
            If vParentForm IsNot Nothing Then
              If vParentForm IsNot Nothing AndAlso (pContactNumber = 0 OrElse vParentForm.ContactInfo.ContactNumber = pContactNumber) Then
                Refresh(pType, vParentForm, pContactNumber)
              End If
            End If
          Next
        End If
      Else
        For Each vForm In mvForms
          vParentForm = TryCast(vForm, MaintenanceParentForm)
          If vParentForm IsNot Nothing AndAlso (pContactNumber = 0 OrElse vParentForm.ContactInfo.ContactNumber = pContactNumber) Then
            Refresh(pType, vParentForm, pContactNumber)
          End If
        Next
      End If
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Shared Sub Refresh(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pParentForm As MaintenanceParent, ByVal pContactNumber As Integer)
    Select Case pType
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings
        Select Case pParentForm.ContactDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, _
               CareServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates, _
               CareServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings, _
               CareServices.XMLContactDataSelectionTypes.xcdtContactEventSessions
            pParentForm.RefreshData()
        End Select
      Case Else
        If pParentForm.ContactDataType = pType AndAlso (pContactNumber = 0 OrElse pParentForm.ContactInfo.ContactNumber = pContactNumber) Then
          pParentForm.RefreshData()
        End If
    End Select
  End Sub

  Private Shared Sub EMailDeleted(ByVal pEmailMessage As EMailMessage)
    For Each vForm As Form In Forms
      If TypeOf (vForm) Is frmInbox Then
        DirectCast(vForm, frmInbox).EMailDeleted(pEmailMessage.ID)
      End If
    Next
  End Sub

  Private Shared Function GetCurrentContact() As ContactInfo
    Return CurrentContact
  End Function

#End Region

#Region "All Main Form Methods"

  Public Shared Sub CloseAllForms()
    If mvMDIMode Then
      mvMDIForm.Close()
    Else
      While mvMainForms.Count > 0
        mvMainForms.Item(0).Close()
      End While
    End If
  End Sub

  Public Shared Property ShowToolbar() As Boolean
    Get
      Return mvShowToolbar
    End Get
    Set(ByVal Value As Boolean)
      For Each vForm As Form In mvForms
        If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.ToolBarChecked = Value
      Next
      mvShowToolbar = Value
    End Set
  End Property

  Public Shared Property ShowHeaderPanel() As Boolean
    Get
      Return mvShowHeaderPanels
    End Get
    Set(ByVal Value As Boolean)
      For Each vForm As Form In mvForms
        If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.HeaderPanelChecked = Value
      Next
      mvShowHeaderPanels = Value
      SetPanelVisibility()
    End Set
  End Property

  Public Shared Property ShowSelectionPanel() As Boolean
    Get
      Return mvShowSelectionPanel
    End Get
    Set(ByVal Value As Boolean)
      For Each vForm As Form In mvForms
        If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.SelectionPanelChecked = Value
      Next
      mvShowSelectionPanel = Value
      SetPanelVisibility()
    End Set
  End Property

  Private Shared Sub SetPanelVisibility()
    For Each vForm As Form In Forms
      If TypeOf (vForm) Is IPanelVisibility Then
        CType(vForm, IPanelVisibility).SetPanelVisibility()
      End If
    Next
  End Sub

  Public Shared Sub EnableTraderApplications(ByVal pEnable As Boolean)
    For Each vForm As Form In mvForms
      If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.EnableTraderApplications(pEnable)
    Next
  End Sub

  Public Shared Sub ResetToolbar()
    CDBNETCL.Settings.MainToolbarItems = ""
    For Each vForm As Form In mvForms
      If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.ResetToolbar()
    Next
  End Sub

  Public Shared Property ToolbarTextPosition() As TextImageRelation
    Get
      Return mvToolbarTextPosition
    End Get
    Set(ByVal pValue As TextImageRelation)
      If pValue <> mvToolbarTextPosition Then
        For Each vForm As Form In mvForms
          If TypeOf (vForm) Is IMainForm Then DirectCast(vForm, IMainForm).MainMenu.ToolbarTextPosition = pValue
        Next
      End If
      mvToolbarTextPosition = pValue
    End Set
  End Property

#End Region

#Region "MDI Only Methods"

  Public Shared ReadOnly Property CurrentContact() As ContactInfo
    Get
      If mvMDIMode Then
        Return mvMDIForm.CurrentContact
      Else
        Return mvLastContact
      End If
    End Get
  End Property

  Public Shared Sub SetStatusContact(ByVal pContactInfo As ContactInfo, ByVal pActive As Boolean)
    If mvMDIMode Then
      mvMDIForm.SetStatusContact(pContactInfo, pActive)
    Else
      If pContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
        If pActive Then
          mvLastContact = pContactInfo
        End If
      Else
        If pActive Then
          mvLastContact = pContactInfo
        End If
      End If
      If pActive = False AndAlso mvLastContact IsNot Nothing AndAlso pContactInfo.ContactNumber = mvLastContact.ContactNumber Then mvLastContact = Nothing
    End If
  End Sub

  Public Shared Property StatusBar() As Boolean
    Get
      If mvMDIMode Then
        Return mvMDIForm.StatusBar
      Else
        Return True
      End If
    End Get
    Set(ByVal value As Boolean)
      If mvMDIMode Then mvMDIForm.StatusBar = value
    End Set
  End Property

  Public Shared Sub UpdateNotificationIcon(ByVal pDataSet As DataSet)
    If mvMDIMode Then mvMDIForm.UpdateNotificationIcon(pDataSet)
    UpdateNotification(pDataSet) 'Notifications are now shown in Task Tray 
  End Sub

  Public Shared Sub SetBackgroundImage(ByVal pFilename As String, ByVal pLayout As ImageLayout)
    If mvMDIMode Then mvMDIForm.SetBackgroundImage(pFilename, pLayout)
  End Sub

  Public Shared Sub SetNotificationTime()
    SetNotificationTimer() 'Notifications are now shown in Task Tray and the MDI form if in MDI mode
  End Sub

  Public Shared Sub SetStatusMessage(ByVal pMessage As String)
    If mvMDIMode Then mvMDIForm.SetStatusMessage(pMessage)
  End Sub

  Public Shared Sub SetFlatMdiArea(setFlat As Boolean)
    Dim vMDIClient As MdiClient
    Dim vWindowLong As Integer

    For Each ctl As Control In MDIForm.Controls
      If TypeOf ctl Is MdiClient Then
        Try
          vMDIClient = CType(ctl, MdiClient)
          vWindowLong = User32API.GetWindowLong(ctl.Handle, GWL.EXSTYLE)
          If setFlat Then
            vWindowLong = CInt(vWindowLong And Not GWL.EX_CLIENTEDGE)
          Else
            vWindowLong = vWindowLong Or GWL.EX_CLIENTEDGE
          End If
          SetWindowLong(vMDIClient.Handle, GWL.EXSTYLE, vWindowLong)
          SetWindowPos(vMDIClient.Handle, IntPtr.Zero, 0, 0, 0, 0, CUInt(SWP.NOACTIVATE Or SWP.NOMOVE Or SWP.NOSIZE Or SWP.NOZORDER Or SWP.NOOWNERZORDER Or SWP.FRAMECHANGED))
          Exit For
        Catch exc As InvalidCastException
        End Try
      End If
    Next

  End Sub

#End Region

#Region "Settings methods"

  Friend Shared Sub SaveSettings()
    UpdateMySettings()
    My.Settings.Save()
  End Sub
  Friend Shared Sub UpgradeSettings()
    My.Settings.Upgrade()
    GetMySettings()
  End Sub

  Friend Shared Sub GetMySettings()
    With My.Settings
      CDBNETCL.Settings.AllowDatabaseSelection = .AllowDatabaseSelection
      CDBNETCL.Settings.BackGroundImage = .BackgroundImage
      CDBNETCL.Settings.BackgroundImageLayout = .BackgroundImageLayout
      CDBNETCL.Settings.DATABASE = .DATABASE
      CDBNETCL.Settings.HistoryDays = .HistoryDays
      CDBNETCL.Settings.LargeNavPanelIcons = .LargeNavPanelIcons
      CDBNETCL.Settings.LargeToolbarIcons = .LargeToolbarIcons
      CDBNETCL.Settings.LargeGridToolbarIcons = .LargeGridToolbarIcons
      CDBNETCL.Settings.MainToolbarItems = .MainToolbarItems
      CDBNETCL.Settings.NavPanelHistoryMode = .NavPanelHistoryMode
      CDBNETCL.Settings.NavPanelPinned = .NavPanelPinned
      CDBNETCL.Settings.NavPanelWidth = .NavPanelWidth
      CDBNETCL.Settings.NotificationPollingMinutes = .NotificationPollingMinutes
      CDBNETCL.Settings.NotifyActions = .NotifyActions
      CDBNETCL.Settings.NotifyDeadlines = .NotifyDeadlines
      CDBNETCL.Settings.NotifyDocuments = .NotifyDocuments
      CDBNETCL.Settings.NotifyMeetings = .NotifyMeetings
      CDBNETCL.Settings.ShowNavPanel = .ShowNavPanel
      CDBNETCL.Settings.ShowToolbar = .ShowToolbar
      CDBNETCL.Settings.UpgradeSettings = .UpgradeSettings
      CDBNETCL.Settings.WebServiceTimeout = .WebServiceTimeout
      CDBNETCL.Settings.WindowSizes = .WindowSizes
      CDBNETCL.Settings.MainToolbarItemsTipText = .MainToolBarTipText
      CDBNETCL.Settings.MainToolbarItemsText = .MainToolBarText
      CDBNETCL.Settings.MainToolbarTextPosition = .MainToolBarTextPosition
      CDBNETCL.Settings.PlainEditPanel = .PlainEditPanel
      CDBNETCL.Settings.ShowDashboard = .ShowDashboard
      CDBNETCL.Settings.DashboardSize = .DashboardSize
      CDBNETCL.Settings.DebugMode = .DebugMode
      CDBNETCL.Settings.FontThemeID = .FontThemeID
      CDBNETCL.Settings.AppearanceThemeID = .AppearanceThemeID
      CDBNETCL.Settings.ShowHeaderPanel = .ShowHeaderPanel
      CDBNETCL.Settings.ShowSelectionPanel = .ShowSelectionPanel
      CDBNETCL.Settings.ShowStatusBar = .ShowStatusBar
      CDBNETCL.Settings.TabIntoDisplayPanel = .TabIntoDisplayPanel
      CDBNETCL.Settings.TabIntoHeaderPanel = .TabIntoHeaderPanel
      CDBNETCL.Settings.ShowErrorsAsMsgBox = .ShowErrorsAsMsgBox
      CDBNETCL.Settings.DisplayDashboardAtLogin = .DisplayDashboardAtLogin
      CDBNETCL.Settings.ConfirmCancel = .ConfirmCancel
      CDBNETCL.Settings.ConfirmInsert = .ConfirmInsert
      CDBNETCL.Settings.ConfirmUpdate = .ConfirmUpdate
      CDBNETCL.Settings.ConfirmDelete = .ConfirmDelete
      CDBNETCL.Settings.TaskNotificationPollingSeconds = .TaskNotificationPollingSeconds
      CDBNETCL.Settings.HideHistoricNetwork = .HideHistoricNetwork
      CDBNETCL.Settings.FinderResultsMsgBox = .FinderResultsMsgBox
      CDBNETCL.Settings.HomeToolbar = .HomeToolbar
      CDBNETCL.Settings.SchemeID = .SchemeID
      CDBNETCL.Settings.FavouriteCommands = .FavouriteCommands
    End With
  End Sub

  Private Shared Sub UpdateMySettings()
    With My.Settings
      .BackgroundImage = CDBNETCL.Settings.BackGroundImage
      .BackgroundImageLayout = CDBNETCL.Settings.BackgroundImageLayout
      .DATABASE = CDBNETCL.Settings.DATABASE
      .HistoryDays = CDBNETCL.Settings.HistoryDays
      .LargeNavPanelIcons = CDBNETCL.Settings.LargeNavPanelIcons
      .LargeToolbarIcons = CDBNETCL.Settings.LargeToolbarIcons
      .LargeGridToolbarIcons = CDBNETCL.Settings.LargeGridToolbarIcons
      .MainToolbarItems = CDBNETCL.Settings.MainToolbarItems
      .NavPanelHistoryMode = CDBNETCL.Settings.NavPanelHistoryMode
      .NavPanelPinned = CDBNETCL.Settings.NavPanelPinned
      .NavPanelWidth = CDBNETCL.Settings.NavPanelWidth
      .NotificationPollingMinutes = CDBNETCL.Settings.NotificationPollingMinutes
      .NotifyActions = CDBNETCL.Settings.NotifyActions
      .NotifyDeadlines = CDBNETCL.Settings.NotifyDeadlines
      .NotifyDocuments = CDBNETCL.Settings.NotifyDocuments
      .NotifyMeetings = CDBNETCL.Settings.NotifyMeetings
      .ShowNavPanel = CDBNETCL.Settings.ShowNavPanel
      .ShowToolbar = CDBNETCL.Settings.ShowToolbar
      .UpgradeSettings = CDBNETCL.Settings.UpgradeSettings
      .WebServiceTimeout = CDBNETCL.Settings.WebServiceTimeout
      .WindowSizes = CDBNETCL.Settings.WindowSizes
      .MainToolBarTipText = CDBNETCL.Settings.MainToolbarItemsTipText
      .MainToolBarText = CDBNETCL.Settings.MainToolbarItemsText
      .MainToolBarTextPosition = CDBNETCL.Settings.MainToolbarTextPosition
      .PlainEditPanel = CDBNETCL.Settings.PlainEditPanel
      .ShowDashboard = CDBNETCL.Settings.ShowDashboard
      .DashboardSize = CDBNETCL.Settings.DashboardSize
      .DebugMode = CDBNETCL.Settings.DebugMode
      .FontThemeID = CDBNETCL.Settings.FontThemeID
      .AppearanceThemeID = CDBNETCL.Settings.AppearanceThemeID
      .ShowHeaderPanel = CDBNETCL.Settings.ShowHeaderPanel
      .ShowSelectionPanel = CDBNETCL.Settings.ShowSelectionPanel
      .ShowStatusBar = CDBNETCL.Settings.ShowStatusBar
      .TabIntoDisplayPanel = CDBNETCL.Settings.TabIntoDisplayPanel
      .TabIntoHeaderPanel = CDBNETCL.Settings.TabIntoHeaderPanel
      .ShowErrorsAsMsgBox = CDBNETCL.Settings.ShowErrorsAsMsgBox
      .DisplayDashboardAtLogin = CDBNETCL.Settings.DisplayDashboardAtLogin
      .ConfirmCancel = CDBNETCL.Settings.ConfirmCancel
      .ConfirmInsert = CDBNETCL.Settings.ConfirmInsert
      .ConfirmUpdate = CDBNETCL.Settings.ConfirmUpdate
      .ConfirmDelete = CDBNETCL.Settings.ConfirmDelete
      .TaskNotificationPollingSeconds = CDBNETCL.Settings.TaskNotificationPollingSeconds
      .HideHistoricNetwork = CDBNETCL.Settings.HideHistoricNetwork
      .FinderResultsMsgBox = CDBNETCL.Settings.FinderResultsMsgBox
      .HomeToolbar = CDBNETCL.Settings.HomeToolbar
      .SchemeID = CDBNETCL.Settings.SchemeID
      .FavouriteCommands = CDBNETCL.Settings.FavouriteCommands
    End With
  End Sub

#End Region

#Region "Dashboard Container Methods"

  Public Shared Sub NavigateHistoryItem(ByVal pHistoryEntityType As HistoryEntityTypes, ByVal pNumber As Integer, Optional ByVal pShowNewWindow As Boolean = False)
    Select Case pHistoryEntityType
      Case HistoryEntityTypes.hetContacts
        FormHelper.ShowContactCardIndex(pNumber, pShowNewWindow)
      Case HistoryEntityTypes.hetActions
        FormHelper.EditAction(pNumber)
      Case HistoryEntityTypes.hetDocuments
        FormHelper.EditDocument(pNumber)
      Case HistoryEntityTypes.hetEvents
        FormHelper.ShowEventIndex(pNumber)
      Case HistoryEntityTypes.hetExamCentres
        FormHelper.ShowExamIndex(pNumber, "N")
      Case HistoryEntityTypes.hetExamUnits
        FormHelper.ShowExamIndex(pNumber, "U")
      Case HistoryEntityTypes.hetExamCentreUnits
        FormHelper.ShowExamIndex(pNumber, "X")
      Case HistoryEntityTypes.hetMeetings
        FormHelper.EditMeeting(pNumber)
      Case HistoryEntityTypes.hetLegacies
        FormHelper.ShowLegacy(pNumber)
      Case HistoryEntityTypes.hetWorkstreams
        FormHelper.ShowWorkstreamIndex(pNumber)
      Case HistoryEntityTypes.hetContactPositions
        FormHelper.ShowContactPosition(pNumber, pShowNewWindow)
      Case Else
        'Do Nothing
    End Select
  End Sub

  Public Shared Sub SmartSearch(pSearchText As String)
    Dim vNumber As Integer = 0
    If Integer.TryParse(pSearchText, vNumber) Then
      MainHelper.NavigateHistoryItem(HistoryEntityTypes.hetContacts, vNumber)
    Else
      MainHelper.ProcessSearch(pSearchText)
    End If

  End Sub

  Public Shared Sub ProcessSearch(ByVal pSearchText As String)
    Dim vList As New ParameterList
    vList("SearchText") = pSearchText
    FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftTextSearch, vList)
  End Sub

  Public Shared Sub SetBrowserMenu(ByVal pSender As Object, ByVal pForm As MaintenanceParentForm)
    DirectCast(pSender, Control).ContextMenuStrip = New BrowserMenu(pForm)
    If TryCast(pSender, INavigable) IsNot Nothing Then
      RegisterForNavigation(DirectCast(pSender, INavigable))
    End If
  End Sub

  Public Shared Sub HistoryItemSelected(ByVal pSender As Object, ByVal pHistoryItem As CDBNETCL.UserHistoryItem, ByVal pDescription As String, ByVal pList As CDBNETCL.ArrayListEx)
    Dim vBrowserMenu As BrowserMenu = TryCast(DirectCast(pSender, Control).ContextMenuStrip, BrowserMenu)
    If vBrowserMenu IsNot Nothing Then
      vBrowserMenu.EntityType = pHistoryItem.HistoryEntityType
      vBrowserMenu.ItemNumber = pHistoryItem.Number
      vBrowserMenu.GroupCode = pHistoryItem.GroupCode
      vBrowserMenu.Favourite = pHistoryItem.Favourite
      vBrowserMenu.ItemDescription = pDescription
      vBrowserMenu.ItemList = pList
    End If
  End Sub

  Public Shared Sub SetActionMenu(ByVal pSender As Object, ByVal pForm As MaintenanceParentForm)
    DirectCast(pSender, Control).ContextMenuStrip = New ActionMenu(pForm)
  End Sub

  Public Shared Sub ActionItemSelected(ByVal pSender As Object)
    Dim vActionMenu As ActionMenu = TryCast(DirectCast(pSender, Control).ContextMenuStrip, ActionMenu)
    If vActionMenu IsNot Nothing Then
      Dim vGrd As DisplayGrid = TryCast(pSender, DisplayGrid)
      If vGrd IsNot Nothing Then
        Dim vRow As Integer = vGrd.CurrentRow
        vActionMenu.ActionNumber = IntegerValue(vGrd.GetValue(vRow, "ActionNumber"))
        vActionMenu.ActionStatus = vGrd.GetValue(vRow, "ActionStatus")
        vActionMenu.RelatedContactNumber = IntegerValue(vGrd.GetValue(vRow, "ContactNumber"))
      End If
    End If
  End Sub

  Public Shared Sub CalendarDoubleClicked(ByVal pForm As MaintenanceParentForm, ByVal pType As CalendarView.CalendarItemTypes, ByVal pDescription As String, ByVal pStart As Date, ByVal pEnd As Date, ByVal pUniqueID As Integer)
    Try
      Select Case pType
        Case CalendarView.CalendarItemTypes.catMeeting
          FormHelper.EditMeeting(pUniqueID, pForm)

        Case CalendarView.CalendarItemTypes.catOther
          Dim vList As New ParameterList
          vList("AppointmentDesc") = pDescription
          vList("StartDate") = pStart.ToString
          vList("EndDate") = pEnd.ToString
          Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctAppointments, DataHelper.UserContactInfo.ContactNumber, Nothing, vList)
          vForm.ShowDialog(pForm)

        Case CalendarView.CalendarItemTypes.catServiceBooking
          Dim vList As New ParameterList(True, True)
          vList("RecordType") = "S"
          vList.IntegerValue("UniqueId") = pUniqueID
          vList("StartDate") = pStart.ToString
          vList("EndDate") = pEnd.ToString
          Dim vSBDS As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAppointmentDetails, GetCurrentContact.ContactNumber, vList)
          If vSBDS IsNot Nothing Then
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vSBDS)
            If vRow IsNot Nothing Then
              Dim vContactNumber As Integer = IntegerValue(vRow.Item("ContactNumber").ToString)
              Dim vBatchNumber As Integer = IntegerValue(vRow.Item("BatchNumber").ToString)
              Dim vTransNumber As Integer = IntegerValue(vRow.Item("TransactionNumber").ToString)
              Dim vLineNumber As Integer = IntegerValue(vRow.Item("LineNumber").ToString)
              If vContactNumber > 0 AndAlso vBatchNumber > 0 Then
                Dim vForm As Form = FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vContactNumber, False)
                If vForm IsNot Nothing Then DirectCast(vForm, frmCardSet).SelectTransaction(vBatchNumber, vTransNumber, vLineNumber)
              End If
            End If
          End If
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

#End Region

#Region "TaskTrayIcon"
  Private Shared Sub InitialiseTaskTrayIcon()
    mvTaskBarIcon = New System.Windows.Forms.NotifyIcon()
    mvTaskBarIcon.Icon = CType(My.Resources.NG_icon, System.Drawing.Icon)
    mvTaskBarIconMenu = New TaskBarIconMenu()
    mvTaskBarIcon.Text = ControlText.TaskBarIconText
    mvTaskBarIcon.ContextMenuStrip = mvTaskBarIconMenu
    mvTaskBarIcon.Visible = True
    SetNotificationTimer()
    SetTaskNotificationTimer()
  End Sub

  Private Shared Sub ClearTaskTrayIcon()
    If mvTimer IsNot Nothing Then mvTimer.Stop()
    If mvTaskTimer IsNot Nothing Then mvTaskTimer.Stop()
    If mvTaskBarIcon.Visible Then mvTaskBarIcon.Visible = False
  End Sub

  Private Shared Sub mvTaskTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles mvTaskTimer.Elapsed
    Try
      AsynchGetTasknotifications()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Shared Sub AsynchGetTasknotifications()
    Dim vMI As New MethodInvoker(AddressOf GetTaskNotifications)
    vMI.BeginInvoke(Nothing, Nothing)
  End Sub

  Private Shared Sub GetTaskNotifications()
    Dim vList As New ParameterList(True)
    Dim vCount As Integer

    Try
      If Not FormHelper.TaskStatusForm Is Nothing Then
        FormHelper.TaskStatusForm.DoRefresh()
      Else
        vList("SubmittedBy") = DataHelper.UserInfo.Logname
        vList("JobStatus") = "C,H"
        Dim vDataSet As DataSet = DataHelper.GetJobScheduleData(vList)
        mvTaskStatus = ""

        If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
          vCount = vDataSet.Tables("DataRow").Rows.Count
        End If
        If vCount > 0 Then
          mvTaskStatus = "Task(s) - " + vCount.ToString
          mvTaskBarIconMenu.ShowTaskItems = True
          If Not mvTaskBarIcon.Visible Then
            mvTaskBarIcon.Visible = True
          End If
          If mvPreTaskCount <> vCount Then
            ShowTaskIconPopUp()
            mvPreTaskCount = vCount
          End If
        Else
          If mvTaskBarIconMenu.ShowNotificationItems = False Then
            mvTaskBarIcon.Visible = False
          Else
            mvTaskBarIconMenu.ShowTaskItems = False
          End If
          If mvTaskTimer IsNot Nothing Then mvTaskTimer.Stop()
        End If
      End If
    Catch vException As Exception
      If Diagnostics.Debugger.IsAttached Then
        Debug.Print(vException.Message)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Shared Sub ShowTaskIconPopUp()
    Dim vStringBuilder As String

    If mvNotificationText.Length > 0 And mvTaskStatus.Length > 0 Then
      vStringBuilder = mvNotificationText + vbCrLf + mvTaskStatus
    ElseIf mvNotificationText.Length > 0 Then
      vStringBuilder = mvNotificationText
    Else
      vStringBuilder = mvTaskStatus
    End If
    If vStringBuilder.Length > 0 Then mvTaskBarIcon.ShowBalloonTip(100, "Notifications", vStringBuilder, ToolTipIcon.Info)
  End Sub


  Public Shared Sub SetNotificationTimer()
    If mvTimer Is Nothing Then
      If Settings.NotificationPollingMinutes > 0 And DataHelper.UserContactInfo.ContactNumber > 0 Then
        mvTimer = New System.Timers.Timer((Settings.NotificationPollingMinutes * 60) * 1000)
        mvTimer.AutoReset = True
        mvTimer.Enabled = True
        AsynchGetNotifications()
      End If
    Else
      mvTimer.Stop()
      If Settings.NotificationPollingMinutes > 0 And DataHelper.UserContactInfo.ContactNumber > 0 Then
        mvTimer.Interval = (Settings.NotificationPollingMinutes * 60) * 1000
        mvTimer.Start()
      End If
    End If
    AsynchGetNotifications()
  End Sub

  Public Shared Sub SetTaskNotificationTimer()
    If mvTaskTimer Is Nothing Then
      If Settings.TaskNotificationPollingSeconds > 0 And DataHelper.UserContactInfo.ContactNumber > 0 Then
        mvTaskTimer = New System.Timers.Timer(Settings.TaskNotificationPollingSeconds * 1000)
        mvTaskTimer.AutoReset = True
        mvTaskTimer.Enabled = True
      End If
    Else
      mvTaskTimer.Stop()
      If Settings.NotificationPollingMinutes > 0 And DataHelper.UserContactInfo.ContactNumber > 0 Then
        mvTaskTimer.Interval = (Settings.TaskNotificationPollingSeconds) * 1000
        mvTaskTimer.Start()
      End If
    End If
    AsynchGetTasknotifications()
  End Sub

  Private Shared Sub mvTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles mvTimer.Elapsed
    Try
      AsynchGetNotifications()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Shared Sub AsynchGetNotifications()
    Dim vMI As New MethodInvoker(AddressOf GetNotifications)
    vMI.BeginInvoke(Nothing, Nothing)
  End Sub

  Private Shared Sub GetNotifications()
    Try
      If Not FormHelper.NotifyForm Is Nothing Then
        FormHelper.NotifyForm.DoRefresh()
      Else
        Dim vSelect As Boolean = (Settings.NotifyActions Or Settings.NotifyDeadlines Or Settings.NotifyDocuments Or Settings.NotifyMeetings) AndAlso DataHelper.UserInfo.ContactNumber > 0
        Dim vDataSet As DataSet = Nothing
        If vSelect Then
          Dim vList As New ParameterList(True)
          If Settings.NotifyActions Then vList("NotifyActions") = "Y"
          If Settings.NotifyDocuments Then vList("NotifyDocuments") = "Y"
          If Settings.NotifyDeadlines Then vList("NotifyDeadlines") = "Y"
          If Settings.NotifyMeetings Then vList("NotifyMeetings") = "Y"
          vDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactNotifications, DataHelper.UserInfo.ContactNumber, vList)
        End If
        UpdateNotificationIcon(vDataSet)
      End If
    Catch vException As Exception
      If Diagnostics.Debugger.IsAttached Then
        Debug.Print(vException.Message)
      End If
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Shared Sub mvTaskBarIconMenu_ShowTaskStatus(ByVal pSender As Object, ByVal pName As String) Handles mvTaskBarIconMenu.ShowTaskStatus
    If FormHelper.TaskStatusForm Is Nothing Then FormHelper.TaskStatusForm = New frmTaskStatus
    FormHelper.TaskStatusForm.Show()
  End Sub

  Private Shared Sub mvTaskBarIconMenu_ShowNotifications(ByVal pSender As Object, ByVal pName As String) Handles mvTaskBarIconMenu.ShowNotifications
    If FormHelper.NotifyForm Is Nothing Then FormHelper.NotifyForm = New frmNotify
    FormHelper.NotifyForm.Show()
  End Sub

  Public Shared Sub UpdateNotification(ByVal pDataSet As DataSet)
    Dim vCount As Integer
    Dim vDocuments As Integer
    Dim vActions As Integer
    Dim vMeetings As Integer
    Dim vEntityAlerts As Integer
    Dim vIndex As Integer

    mvNotificationText = ""
    If pDataSet IsNot Nothing AndAlso pDataSet.Tables.Contains("DataRow") Then
      vCount = pDataSet.Tables("DataRow").Rows.Count
      For Each vRow As DataRow In pDataSet.Tables("DataRow").Rows
        Select Case vRow.Item("ItemCode").ToString
          Case "A", "O"
            vActions += 1
          Case "D"
            vDocuments += 1
          Case "M"
            vMeetings += 1
          Case "I"
            vEntityAlerts += 1
        End Select
      Next
      vCount = vActions + vDocuments + vMeetings + vEntityAlerts
      If vCount > 0 Then vIndex = 1
      If vCount > 0 Then
        mvTaskBarIconMenu.ShowNotificationItems = True
        Dim vToolTipText As New StringBuilder
        If vDocuments > 0 Then vToolTipText.Append(String.Format(InformationMessages.ImNotificationsDocuments, vDocuments))
        If vActions > 0 Then
          If vToolTipText.Length > 0 Then vToolTipText.Append(", ")
          vToolTipText.Append(String.Format(InformationMessages.ImNotificationsActions, vActions))
        End If
        If vMeetings > 0 Then
          If vToolTipText.Length > 0 Then vToolTipText.Append(", ")
          vToolTipText.Append(String.Format(InformationMessages.ImNotificationsMeetings, vMeetings))
        End If
        If vEntityAlerts > 0 Then
          If vToolTipText.Length > 0 Then vToolTipText.Append(", ")
          vToolTipText.Append(String.Format(InformationMessages.ImNotificationsAlerts, vEntityAlerts))
        End If

        mvNotificationText = vToolTipText.ToString
        If Not mvTaskBarIcon.Visible Then mvTaskBarIcon.Visible = True
        If mvPreNotificationCount <> vCount Then
          ShowTaskIconPopUp()
          mvPreNotificationCount = vCount
        End If
      Else
        If mvTaskBarIconMenu.ShowTaskItems = False Then
          mvTaskBarIcon.Visible = False
        Else
          mvTaskBarIconMenu.ShowNotificationItems = False
        End If

      End If
    Else
      mvTaskBarIconMenu.ShowNotificationItems = False
    End If
  End Sub

  Private Shared Sub OnExitClick(ByVal sender As Object, ByVal e As EventArgs)

    mvTaskBarIcon.Visible = False
    Application.Exit()
  End Sub

#End Region

  Private Shared Sub mvTaskBarIcon_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvTaskBarIcon.BalloonTipClicked
    If mvPreTaskCount > 0 Then
      If FormHelper.TaskStatusForm Is Nothing Then FormHelper.TaskStatusForm = New frmTaskStatus
      FormHelper.TaskStatusForm.Show()
    ElseIf mvPreNotificationCount > 0 Then
      If FormHelper.NotifyForm Is Nothing Then FormHelper.NotifyForm = New frmNotify
      FormHelper.NotifyForm.Show()
    Else
    End If
  End Sub

#Region "Incoming Phone Call"

  Private Delegate Sub HandleIncomingCallDelegate(ByVal pNumber As String)

  Private Shared Sub IncomingPhoneCall(pPhoneNumber As String)
    If MainForm.InvokeRequired Then
      MainForm.BeginInvoke(New HandleIncomingCallDelegate(AddressOf IncomingPhoneCall), New Object() {pPhoneNumber})
    Else
      'First we need to find out who the phone number belongs to
      Dim vList As New ParameterList(True)
      vList("PhoneNumber") = pPhoneNumber
      vList("UseContactRestriction") = "Y"
      Dim vDataSet As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftOrganisations, vList)
      Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
      If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
        'One or more organisations
        'MessageBox.Show(String.Format("Found {0} Organisations for phone number {1}", vTable.Rows.Count, pPhoneNumber))
        Dim vForm As New frmCLIBrowser(vDataSet, True, pPhoneNumber)
        vForm.ShowDialog()
      Else
        vDataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
        vTable = DataHelper.GetTableFromDataSet(vDataSet)
        If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
          If vTable.Rows.Count = 1 Then
            Dim vForm As Form = FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString), False, True)
            If vForm IsNot Nothing Then
              Dim vTCRList As New ParameterList
              vTCRList("Direction") = "I"
              vTCRList("Precis") = GetInformationMessage(InformationMessages.ImTcrPrecis, pPhoneNumber, Today.ToString(AppValues.DateFormat), Now.ToString("HH:mm:ss"))
              Dim vTCRForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, vTCRList)
              vTCRForm.RelatedContact = DirectCast(vForm, frmCardSet).ContactInfo()
              vTCRForm.Show()
            End If
          Else
            'More than one contact
            'MessageBox.Show(String.Format("Found {0} Contacts for phone number {1}", vTable.Rows.Count, pPhoneNumber))
            Dim vForm As New frmCLIBrowser(vDataSet, False, pPhoneNumber)
            vForm.ShowDialog()
          End If
        Else
          'MessageBox.Show(String.Format("No record found for {0}", pPhoneNumber))
          'Found nobody
          Dim vForm As New frmCLIBrowser(Nothing, False, pPhoneNumber)
          vForm.ShowDialog()
        End If
      End If
    End If
  End Sub

#End Region

  Private Shared Sub InitServiceLocator()
    ServiceLocator.Instance.Register(Of IDialogService, DialogService)()
    ServiceLocator.Instance.Register(Of IMessageBoxService, MessageBoxService)()
    ServiceLocator.Instance.Register(Of IConfirmActionService, ConfirmActionService)()
    ServiceLocator.Instance.Register(Of ICustomiseService, CustomiseService)()
  End Sub

End Class
