Public Class frmDashboard
  Implements IDashboardTabContainer
  Implements IMainForm

  Private mvDashboard As DashboardTabControl
  Private mvMainMenu As MainMenu

#Region "IMainForm Interface"

  Public ReadOnly Property MainMenu() As MainMenu Implements IMainForm.MainMenu
    Get
      Return mvMainMenu
    End Get
  End Property

#End Region

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    mvMainMenu = MainHelper.AddMainMenu(Me)
    mvDashboard = dtc
    SettingsName = "Dashboard"
    MdiParent = MDIForm
    If AppValues.RunAsThames Then Me.Text = ControlText.FrmIRISNFPDashboard
    SetControlTheme()
    Try
      dtc.Init(Me, DashboardTypes.GeneralDashboardType, "")
      OpenDashboard()
      'vBrowser.URL = "http://localhost/ReportServer/Pages/ReportViewer.aspx?%2fReport+Project1%2fTop+50+Donations&rs:Command=Render"
      ''Dim vViewer As DashboardReportViewer
      ''vViewer = New DashboardReportViewer("SQL Reporting Services Report 2")
      ''vViewer.SetReport("http://localhost/ReportServer", "/Report Project1/Top 50 Donations")
      ''vDbc.AddItem(0, vViewer)
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

#Region "IDashboardTabContainer"

  Private Sub OpenDashboard() Implements CDBNETCL.IDashboardTabContainer.Open
    Dim vDashboard As DashboardTabControl = mvDashboard.CreateFromDatabase(Me)
    If vDashboard IsNot Nothing Then
      Me.SuspendLayout()
      Me.Controls.Remove(mvDashboard)
      mvDashboard = vDashboard
      Me.Controls.Add(mvDashboard)
      mvDashboard.SetItemID(0)
      mvDashboard.BringToFront()
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

  Private Sub ProcessEditing(ByVal pType As CDBNETCL.DashboardDisplayPanel.MaintenanceTypes) Implements CDBNETCL.IDashboardTabContainer.ProcessEditing

  End Sub

  Private Sub CalendarDoubleClickedHandler(ByVal pType As CalendarView.CalendarItemTypes, ByVal pDescription As String, ByVal pStart As Date, ByVal pEnd As Date, ByVal pUniqueID As Integer) Implements IDashboardTabContainer.CalendardItemDoubleClicked
    MainHelper.CalendarDoubleClicked(Me, pType, pDescription, pStart, pEnd, pUniqueID)
  End Sub

  Private Sub ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Implements IDashboardTabContainer.ContactSelected
    EntitySelected(pSender, pContactNumber, HistoryEntityTypes.hetContacts)
  End Sub

  Public Sub NavigateNewSelectionSet(pContactNumbers As String) Implements CDBNETCL.IDashboardTabContainer.NavigateNewSelectionSet
    FormHelper.CreateNewSelectionSet(pContactNumbers)
  End Sub

#End Region

  Public Sub EntitySelected(pSender As Object, pEntityNumber As Integer, Optional pEntityType As HistoryEntityTypes = CDBNETCL.HistoryEntityTypes.hetContacts) Implements IDashboardTabContainer.EntitySelected
    MainHelper.NavigateHistoryItem(pEntityType, pEntityNumber, True)
  End Sub
End Class