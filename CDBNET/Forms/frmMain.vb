Imports System.Linq
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Imports CDBNETXAML
Imports WPFControls = System.Windows.Controls
Imports CDBNETXAMLPages


Public Class frmMain
  Inherits PersistentForm
  Implements IMainForm, INotifyPropertyChanged

  Private mvDragging As Boolean = False
  Private mvDragCusorPoint As Point
  Friend WithEvents wpfMenuHeaderHost As Integration.ElementHost
  'Friend wpfMenuHeaderInstance As MenuHeader
  Friend WithEvents pnlExplorerTop As Panel
  Friend WithEvents lblDummyBorder As Label
  Private mvDragFormPoint As Point

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()
    Me.SuspendLayout()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls()
    mvFormDocker = New FormDocker(Me)
    Me.ResumeLayout()
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents pnlNavPanel As System.Windows.Forms.Panel
  Friend WithEvents splNavPanel As System.Windows.Forms.Splitter
  Friend WithEvents pnlTreeView As CDBNETCL.PanelEx
  Friend WithEvents tvw As CDBNETCL.VistaTreeView
  Friend WithEvents sbr As System.Windows.Forms.StatusBar
  Friend WithEvents sbpStatus As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpContact As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpOrganisation As System.Windows.Forms.StatusBarPanel
  Friend WithEvents Tim As System.Windows.Forms.Timer
  Friend WithEvents sbpNotify As System.Windows.Forms.StatusBarPanel
  Friend WithEvents tspNavPanel As System.Windows.Forms.ToolStrip
  Friend WithEvents tsbDateView As System.Windows.Forms.ToolStripButton
  Friend WithEvents tsbHistoryView As System.Windows.Forms.ToolStripButton
  Friend WithEvents tsbPin As System.Windows.Forms.ToolStripButton
  Friend WithEvents tss1 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents pnlMDIExplorer As System.Windows.Forms.Panel
  Friend WithEvents wpfMDIExplorerHost As System.Windows.Forms.Integration.ElementHost
  Friend WithEvents wpfTabStripHost As System.Windows.Forms.Integration.ElementHost
  Friend WithEvents wpfMDIExplorerInstance As New CDBNETXAMLPages.MDIExplorer

  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node1")
    Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node0", New System.Windows.Forms.TreeNode() {TreeNode1})
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
    Me.pnlNavPanel = New System.Windows.Forms.Panel()
    Me.pnlTreeView = New CDBNETCL.PanelEx()
    Me.tvw = New CDBNETCL.VistaTreeView()
    Me.tspNavPanel = New System.Windows.Forms.ToolStrip()
    Me.tsbDateView = New System.Windows.Forms.ToolStripButton()
    Me.tsbHistoryView = New System.Windows.Forms.ToolStripButton()
    Me.tss1 = New System.Windows.Forms.ToolStripSeparator()
    Me.tsbPin = New System.Windows.Forms.ToolStripButton()
    Me.splNavPanel = New System.Windows.Forms.Splitter()
    Me.sbr = New System.Windows.Forms.StatusBar()
    Me.sbpStatus = New System.Windows.Forms.StatusBarPanel()
    Me.sbpContact = New System.Windows.Forms.StatusBarPanel()
    Me.sbpOrganisation = New System.Windows.Forms.StatusBarPanel()
    Me.sbpNotify = New System.Windows.Forms.StatusBarPanel()
    Me.Tim = New System.Windows.Forms.Timer(Me.components)
    Me.pnlMDIExplorer = New System.Windows.Forms.Panel()
    Me.wpfMDIExplorerHost = New System.Windows.Forms.Integration.ElementHost()
    Me.wpfMDIExplorerInstance = New CDBNETXAMLPages.MDIExplorer()
    Me.wpfTabStripHost = New System.Windows.Forms.Integration.ElementHost()
    Me.pnlExplorerTop = New System.Windows.Forms.Panel()
    Me.lblDummyBorder = New System.Windows.Forms.Label()
    Me.pnlNavPanel.SuspendLayout()
    Me.pnlTreeView.SuspendLayout()
    Me.tspNavPanel.SuspendLayout()
    CType(Me.sbpStatus, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpContact, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpOrganisation, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpNotify, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnlMDIExplorer.SuspendLayout()
    Me.pnlExplorerTop.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnlNavPanel
    '
    Me.pnlNavPanel.Controls.Add(Me.pnlTreeView)
    Me.pnlNavPanel.Dock = System.Windows.Forms.DockStyle.Left
    Me.pnlNavPanel.Location = New System.Drawing.Point(0, 0)
    Me.pnlNavPanel.Name = "pnlNavPanel"
    Me.pnlNavPanel.Size = New System.Drawing.Size(192, 728)
    Me.pnlNavPanel.TabIndex = 1
    '
    'pnlTreeView
    '
    Me.pnlTreeView.BackColor = System.Drawing.Color.Transparent
    Me.pnlTreeView.Controls.Add(Me.tvw)
    Me.pnlTreeView.Controls.Add(Me.tspNavPanel)
    Me.pnlTreeView.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlTreeView.Location = New System.Drawing.Point(0, 0)
    Me.pnlTreeView.Name = "pnlTreeView"
    Me.pnlTreeView.Size = New System.Drawing.Size(192, 728)
    Me.pnlTreeView.TabIndex = 1
    '
    'tvw
    '
    Me.tvw.AccessibleName = "Navigation Tree"
    Me.tvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer))
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tvw.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tvw.HideSelection = False
    Me.tvw.Indent = 19
    Me.tvw.Location = New System.Drawing.Point(0, 27)
    Me.tvw.Name = "tvw"
    TreeNode1.ImageIndex = 2
    TreeNode1.Name = "Node1"
    TreeNode1.Text = "Node1"
    TreeNode2.ImageIndex = 1
    TreeNode2.Name = "Node0"
    TreeNode2.Text = "Node0"
    Me.tvw.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode2})
    Me.tvw.Size = New System.Drawing.Size(192, 701)
    Me.tvw.TabIndex = 0
    '
    'tspNavPanel
    '
    Me.tspNavPanel.AccessibleName = "Navigation Tool Bar"
    Me.tspNavPanel.ImageScalingSize = New System.Drawing.Size(20, 20)
    Me.tspNavPanel.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbDateView, Me.tsbHistoryView, Me.tss1, Me.tsbPin})
    Me.tspNavPanel.Location = New System.Drawing.Point(0, 0)
    Me.tspNavPanel.Name = "tspNavPanel"
    Me.tspNavPanel.Size = New System.Drawing.Size(192, 27)
    Me.tspNavPanel.TabIndex = 0
    Me.tspNavPanel.TabStop = True
    '
    'tsbDateView
    '
    Me.tsbDateView.AccessibleName = "Date View"
    Me.tsbDateView.CheckOnClick = True
    Me.tsbDateView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
    Me.tsbDateView.Image = CType(resources.GetObject("tsbDateView.Image"), System.Drawing.Image)
    Me.tsbDateView.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.tsbDateView.Name = "tsbDateView"
    Me.tsbDateView.Size = New System.Drawing.Size(24, 24)
    Me.tsbDateView.ToolTipText = "Date View"
    '
    'tsbHistoryView
    '
    Me.tsbHistoryView.AccessibleName = "History View"
    Me.tsbHistoryView.CheckOnClick = True
    Me.tsbHistoryView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
    Me.tsbHistoryView.Image = CType(resources.GetObject("tsbHistoryView.Image"), System.Drawing.Image)
    Me.tsbHistoryView.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.tsbHistoryView.Name = "tsbHistoryView"
    Me.tsbHistoryView.Size = New System.Drawing.Size(24, 24)
    Me.tsbHistoryView.ToolTipText = "History View"
    '
    'tss1
    '
    Me.tss1.Name = "tss1"
    Me.tss1.Size = New System.Drawing.Size(6, 27)
    '
    'tsbPin
    '
    Me.tsbPin.AccessibleName = "Pin Navigation Panel"
    Me.tsbPin.AccessibleRole = System.Windows.Forms.AccessibleRole.None
    Me.tsbPin.Checked = True
    Me.tsbPin.CheckOnClick = True
    Me.tsbPin.CheckState = System.Windows.Forms.CheckState.Checked
    Me.tsbPin.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
    Me.tsbPin.Image = CType(resources.GetObject("tsbPin.Image"), System.Drawing.Image)
    Me.tsbPin.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.tsbPin.Name = "tsbPin"
    Me.tsbPin.Size = New System.Drawing.Size(24, 24)
    Me.tsbPin.ToolTipText = "Pin Navigation Panel"
    '
    'splNavPanel
    '
    Me.splNavPanel.Location = New System.Drawing.Point(244, 737)
    Me.splNavPanel.Name = "splNavPanel"
    Me.splNavPanel.Size = New System.Drawing.Size(6, 0)
    Me.splNavPanel.TabIndex = 7
    Me.splNavPanel.TabStop = False
    '
    'sbr
    '
    Me.sbr.Location = New System.Drawing.Point(244, 704)
    Me.sbr.Name = "sbr"
    Me.sbr.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbpStatus, Me.sbpContact, Me.sbpOrganisation, Me.sbpNotify})
    Me.sbr.ShowPanels = True
    Me.sbr.Size = New System.Drawing.Size(772, 24)
    Me.sbr.SizingGrip = False
    Me.sbr.TabIndex = 3
    '
    'sbpStatus
    '
    Me.sbpStatus.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbpStatus.Name = "sbpStatus"
    Me.sbpStatus.Width = 254
    '
    'sbpContact
    '
    Me.sbpContact.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbpContact.Name = "sbpContact"
    Me.sbpContact.Width = 254
    '
    'sbpOrganisation
    '
    Me.sbpOrganisation.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbpOrganisation.Name = "sbpOrganisation"
    Me.sbpOrganisation.Width = 254
    '
    'sbpNotify
    '
    Me.sbpNotify.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpNotify.Name = "sbpNotify"
    Me.sbpNotify.Width = 10
    '
    'Tim
    '
    '
    'pnlMDIExplorer
    '
    Me.pnlMDIExplorer.BackColor = System.Drawing.Color.White
    Me.pnlMDIExplorer.Controls.Add(Me.wpfMDIExplorerHost)
    Me.pnlMDIExplorer.Dock = System.Windows.Forms.DockStyle.Left
    Me.pnlMDIExplorer.Location = New System.Drawing.Point(192, 0)
    Me.pnlMDIExplorer.Name = "pnlMDIExplorer"
    Me.pnlMDIExplorer.Size = New System.Drawing.Size(52, 728)
    Me.pnlMDIExplorer.TabIndex = 9
    Me.pnlMDIExplorer.Visible = False
    '
    'wpfMDIExplorerHost
    '
    Me.wpfMDIExplorerHost.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wpfMDIExplorerHost.Location = New System.Drawing.Point(0, 0)
    Me.wpfMDIExplorerHost.Name = "wpfMDIExplorerHost"
    Me.wpfMDIExplorerHost.Size = New System.Drawing.Size(52, 728)
    Me.wpfMDIExplorerHost.TabIndex = 0
    Me.wpfMDIExplorerHost.Text = "ElementHost1"
    Me.wpfMDIExplorerHost.Child = Me.wpfMDIExplorerInstance
    '
    'wpfTabStripHost
    '
    Me.wpfTabStripHost.AutoSize = True
    Me.wpfTabStripHost.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wpfTabStripHost.Location = New System.Drawing.Point(0, 0)
    Me.wpfTabStripHost.Margin = New System.Windows.Forms.Padding(0)
    Me.wpfTabStripHost.Name = "wpfTabStripHost"
    Me.wpfTabStripHost.Size = New System.Drawing.Size(772, 737)
    Me.wpfTabStripHost.TabIndex = 12
    Me.wpfTabStripHost.Child = Nothing
    '
    'pnlExplorerTop
    '
    Me.pnlExplorerTop.AutoSize = True
    Me.pnlExplorerTop.Controls.Add(Me.wpfTabStripHost)
    Me.pnlExplorerTop.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlExplorerTop.Location = New System.Drawing.Point(244, 0)
    Me.pnlExplorerTop.Name = "pnlExplorerTop"
    Me.pnlExplorerTop.Size = New System.Drawing.Size(772, 737)
    Me.pnlExplorerTop.TabIndex = 15
    '
    'lblDummyBorder
    '
    Me.lblDummyBorder.BackColor = System.Drawing.Color.Gainsboro
    Me.lblDummyBorder.Dock = System.Windows.Forms.DockStyle.Left
    Me.lblDummyBorder.Location = New System.Drawing.Point(250, 737)
    Me.lblDummyBorder.Name = "lblDummyBorder"
    Me.lblDummyBorder.Size = New System.Drawing.Size(1, 0)
    Me.lblDummyBorder.TabIndex = 17
    '
    'frmMain
    '
    Me.ClientSize = New System.Drawing.Size(1016, 728)
    Me.Controls.Add(Me.lblDummyBorder)
    Me.Controls.Add(Me.splNavPanel)
    Me.Controls.Add(Me.pnlExplorerTop)
    Me.Controls.Add(Me.sbr)
    Me.Controls.Add(Me.pnlMDIExplorer)
    Me.Controls.Add(Me.pnlNavPanel)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.IsMdiContainer = True
    Me.Name = "frmMain"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.pnlNavPanel.ResumeLayout(False)
    Me.pnlTreeView.ResumeLayout(False)
    Me.pnlTreeView.PerformLayout()
    Me.tspNavPanel.ResumeLayout(False)
    Me.tspNavPanel.PerformLayout()
    CType(Me.sbpStatus, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpContact, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpOrganisation, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpNotify, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnlMDIExplorer.ResumeLayout(False)
    Me.pnlExplorerTop.ResumeLayout(False)
    Me.pnlExplorerTop.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

#End Region

#Region "Enums and Module Variables"

  Private Enum TimerUsage
    tuNone
    tuMaximiseNavPanel
    tuMinimizeNavPanel
  End Enum

  Private mvTimerUsage As TimerUsage
  Private mvNavPanelWidth As Integer = 192
  Private mvCurrentContact As ContactInfo
  Private mvCurrentOrganisation As ContactInfo
  Private mvLastContact As ContactInfo
  Private mvNavPanelMinimized As Boolean

  Private WithEvents mvMainMenu As MainMenu
  Private WithEvents mvBrowserMenu As BrowserMenu
  Private WithEvents mvTreeViewSubclasser As SubClasser

  Private WithEvents mvExplorerViewModel As CDBNETXAMLPages.MDIExplorerViewModel
  Private WithEvents mvHomeScreenViewModel As CDBNETXAMLPages.HomeScreenViewModel
  Private mvShowClassicToolstrip As Boolean = False
  Private mvOpenFormsStrip As CDBNETXAMLPages.OpenFormsStrip
  Private mvOpenFormsViewModel As CDBNETXAMLPages.OpenFormsViewModel
  Private mvHomeToolbarViewModel As CDBNETXAMLPages.HomeToolbarViewModel
  Private WithEvents mvFormDocker As FormDocker

#End Region

#Region "Main Sub"

  <System.STAThread()>
  Public Shared Sub Main()

    Application.EnableVisualStyles()
    Application.DoEvents()

    AppDomain.CurrentDomain.SetPrincipalPolicy(Security.Principal.PrincipalPolicy.WindowsPrincipal)
    System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo(System.Threading.Thread.CurrentThread.CurrentCulture.Name)
    AppValues.SetResourceManagers()
    Settings.Save = New Settings.SaveSettingsDelegate(AddressOf MainHelper.SaveSettings)
    Settings.Upgrade = New Settings.UpgradeSettingsDelegate(AddressOf MainHelper.UpgradeSettings)
    MainHelper.GetMySettings()
    If My.Application.IsNetworkDeployed Then
      AppValues.Init(My.Application.CommandLineArgs, My.Application.IsNetworkDeployed, My.Application.Deployment.UpdateLocation)
    Else
      AppValues.Init(My.Application.CommandLineArgs, False, Nothing)
    End If
    DisplayTheme.PlainEditPanelTheme = Settings.PlainEditPanel
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psConnecting)
    Try
      If DataHelper.CheckVersion Then
        DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
        Dim vRun As Boolean = False
        If DataHelper.AuthenticatedUser.Length > 0 Then
          Try
            DataHelper.Login("CD", DataHelper.AuthenticatedUser, "none", AppValues.Database, DataHelper.AuthenticatedUser)
            vRun = True
          Catch vEx As CareException
            If vEx.ErrorNumber <> CareException.ErrorNumbers.enLoginFailed Then
              Throw
            End If
          End Try
        End If
        If Not vRun Then
          Dim vForm As New frmLogin("CD")
          If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then vRun = True
          vForm = Nothing
        End If
        If vRun Then
          DataHelper.ShowProgress(frmProgress.ProgressStatuses.psInitialising)
          Try
            DisplayTheme.InitFromSettings()
            MainHelper.Start()
          Catch vEx As Exception
            DataHelper.HandleException(vEx)
          End Try
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

#End Region

#Region "Settings methods"

  Public Sub SetBackgroundImage(ByVal pFilename As String, ByVal pLayout As ImageLayout)
    Me.BackgroundImageLayout = pLayout
    If pFilename.Length = 0 Then
      Me.BackgroundImage = Nothing
    Else
      Me.BackgroundImage = Image.FromFile(pFilename)
    End If
  End Sub

  Public Sub SetControlColors()
    Me.Font = DisplayTheme.FormFont
    tvw.Font = DisplayTheme.NavigationPanelFont
    wpfMDIExplorerHost.Font = DisplayTheme.NavigationPanelFont
  End Sub

#End Region

  Private Function GetQueryStringParameters() As Dictionary(Of String, String)
    Dim NameValueTable As New System.Collections.Generic.Dictionary(Of String, String)()

    If My.Application.IsNetworkDeployed Then
      Dim vUrl As String = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData(0)
      Dim vUri As Uri = New Uri(vUrl)
      'vUri.Host

      Dim QueryString As String = (New Uri(vUrl)).Query
      Dim NameValuePairs() As String = QueryString.Split("&"c)
      For Each NameValuePair As String In NameValuePairs
        Dim Vars() As String = NameValuePair.Split("="c)
        If (Not NameValueTable.ContainsKey(Vars(0))) Then
          NameValueTable.Add(Vars(0), Vars(1))
        End If
      Next
    End If
    Return NameValueTable
  End Function

#Region "Form Events and Initialisation"

  Private Sub InitialiseControls()
    Me.SuspendLayout()
    Me.tvw.SuspendLayout()
    Try
      SettingsName = "MainForm"
      MDIForm = Me
      mvBrowserMenu = New BrowserMenu(Nothing)
      tvw.ContextMenuStrip = mvBrowserMenu
      tvw.Indent = AppValues.TreeViewIndent
      'Dim vImageProvider As New ImageProvider
      'tvw.ImageList = vImageProvider.NewTreeViewImages

      tvw.ImageList = AppHelper.ImageProvider.NewTreeViewImages

      mvBrowserMenu.RemoveSupported = True
      If AppValues.DatabaseDescription.Length > 0 Then
        Me.Text = String.Format("{0} ({1})", ProductName, AppValues.DatabaseDescription)
      Else
        Me.Text = ProductName
      End If
      mvMainMenu = MainHelper.AddMainMenu(Me)
      SetControlColors()
    Finally
      Me.tvw.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub frmMain_DragOver(sender As Object, e As DragEventArgs) Handles Me.DragOver

  End Sub

  Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    AppValues.SaveWindowSizes()
    Settings.Save()
  End Sub

  Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.SuspendLayout()
    SetBackgroundImage(Settings.BackGroundImage, Settings.BackgroundImageLayout)
    MainHelper.ShowToolbar = Settings.ShowToolbar
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psCheckingHistory)
    mvTreeViewSubclasser = New SubClasser(tvw.Handle)
    mvTreeViewSubclasser.SubClass = True
    NavigationPanel = Settings.ShowNavPanel
    LargeNavPanelIcons = Settings.LargeNavPanelIcons
    mvNavPanelWidth = Settings.NavPanelWidth
    pnlNavPanel.Width = mvNavPanelWidth
    tsbPin.Checked = Settings.NavPanelPinned
    CheckMinimizeNavPanel()

    MainHelper.ShowHeaderPanel = Settings.ShowHeaderPanel
    MainHelper.ShowSelectionPanel = Settings.ShowSelectionPanel
    StatusBar = Settings.ShowStatusBar
    Dim vHistoryMode As HistoryViewModes = CType(Settings.NavPanelHistoryMode, HistoryViewModes)
    If vHistoryMode = HistoryViewModes.hvmDate Then
      tsbDateView.Checked = True
    Else
      tsbHistoryView.Checked = True
    End If
    UserHistory.ClearOldHistory()
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psLoadingHistory)
    UserHistory.GetHistoryItems(tvw, vHistoryMode)
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psCheckingNotifications)
    sbpNotify.Icon = My.Resources.iconNoNotifications
    sbpNotify.ToolTipText = InformationMessages.ImNotificationDisabled
    If NavigationPanel Then tvw.Focus()
    Me.KeyPreview = True
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
    If Settings.DisplayDashboardAtLogin AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDashboardView) Then
      Dim vForm As New frmDashboard()
      vForm.Show()
    End If
    If AppHelper.FormView = FormViews.Modern Then
      SetExplorerView()
    Else
      SetClassicView()
    End If
    Me.ResumeLayout()
  End Sub

  Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    PhoneApplication.DisableCallControl()
    Settings.NavPanelWidth = mvNavPanelWidth
    Dim vHistoryMode As HistoryViewModes
    If tsbDateView.Checked Then
      vHistoryMode = HistoryViewModes.hvmDate
    Else
      vHistoryMode = HistoryViewModes.hvmHistory
    End If
    Settings.NavPanelHistoryMode = CInt(vHistoryMode)
    Settings.NavPanelPinned = tsbPin.Checked
    Settings.ShowNavPanel = NavigationPanel
    Settings.LargeNavPanelIcons = LargeNavPanelIcons
    If AppHelper.FormView = FormViews.Classic Then
      Settings.ShowToolbar = MainHelper.ShowToolbar
      Settings.ShowStatusBar = StatusBar
    End If
    Settings.LargeToolbarIcons = mvMainMenu.LargeToolbarIcons
    Settings.ShowHeaderPanel = MainHelper.ShowHeaderPanel
    Settings.ShowSelectionPanel = MainHelper.ShowSelectionPanel
    If mvHomeToolbarViewModel IsNot Nothing Then
      Settings.HomeToolbar = mvHomeToolbarViewModel.SerializeToolbar()
    End If
    mvMainMenu.SaveToolbarItems()
    DataHelper.Logout("CD")
    System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeShutdown()
    Application.Exit()
  End Sub

  Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    If Not DocumentApplication Is Nothing Then DocumentApplication.ProcessAppActive()
  End Sub

  Private Sub frmMain_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MdiChildActivate
    If Me.ActiveMdiChild Is Nothing AndAlso NavigationPanel Then
      tvw.Focus()
    End If
    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("MdiChildren"))
  End Sub


#End Region

#Region "Status Bar Handling"

  Private mvSBRClickNotify As Boolean

  Public ReadOnly Property CurrentContact() As ContactInfo
    Get
      Return mvLastContact
    End Get
  End Property
  Public Property StatusBar() As Boolean
    Get
      Return sbr.Visible
    End Get
    Set(ByVal Value As Boolean)
      sbr.Visible = Value
      mvMainMenu.StatusBarChecked = Value
    End Set
  End Property
  Public Sub SetStatusMessage(ByVal pMessage As String)
    sbpStatus.Text = pMessage
  End Sub

  Public Sub SetStatusContact(ByVal pContactInfo As ContactInfo, ByVal pActive As Boolean)
    If pContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
      If pActive Then
        mvCurrentOrganisation = pContactInfo
        mvLastContact = pContactInfo
        sbpOrganisation.Text = pContactInfo.ContactName & " (" & pContactInfo.ContactNumber & ")"
      ElseIf Not mvCurrentOrganisation Is Nothing Then
        If pContactInfo.ContactNumber = mvCurrentOrganisation.ContactNumber Then mvCurrentOrganisation = Nothing
        sbpOrganisation.Text = ""
      End If
    Else
      If pActive Then
        mvCurrentContact = pContactInfo
        mvLastContact = pContactInfo
        sbpContact.Text = pContactInfo.ContactName & " (" & pContactInfo.ContactNumber & ")"
      ElseIf Not mvCurrentContact Is Nothing Then
        If pContactInfo.ContactNumber = mvCurrentContact.ContactNumber Then mvCurrentContact = Nothing
        sbpContact.Text = ""
      End If
    End If
    If pActive = False AndAlso mvLastContact IsNot Nothing AndAlso pContactInfo.ContactNumber = mvLastContact.ContactNumber Then mvLastContact = Nothing
  End Sub
  Private Sub sbr_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles sbr.MouseDown
    mvSBRClickNotify = False
    If e.X > sbpStatus.Width Then
      If e.X < (sbr.Width - sbpNotify.Width) Then
        If e.X > (sbpStatus.Width + sbpContact.Width) Then
          If Not mvCurrentOrganisation Is Nothing Then DoDragDrop(mvCurrentOrganisation, DragDropEffects.Copy) 'DragDrop Org
        Else
          If Not mvCurrentContact Is Nothing Then DoDragDrop(mvCurrentContact, DragDropEffects.Copy) 'DragDrop Contact
        End If
      Else
        If sbpNotify.Text.Length > 0 Then mvSBRClickNotify = True
      End If
    End If
  End Sub
  Private Sub sbr_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles sbr.DoubleClick
    If mvSBRClickNotify Then
      If FormHelper.NotifyForm Is Nothing Then FormHelper.NotifyForm = New frmNotify
      FormHelper.NotifyForm.Show()
    End If
  End Sub

#End Region

#Region "Toolbar Handling"



#End Region

#Region "Navigation Panel Handling"

  Public Property NavigationPanel() As Boolean
    Get
      Return pnlNavPanel.Visible
    End Get
    Set(ByVal Value As Boolean)
      pnlNavPanel.Visible = Value
      splNavPanel.Visible = Value
      mvMainMenu.NavigationPanelChecked = Value
    End Set
  End Property
  Private Property LargeNavPanelIcons() As Boolean
    Get
      Return tspNavPanel.ImageList Is AppHelper.ImageProvider.NavPanel32
    End Get
    Set(ByVal Value As Boolean)
      If Value Then
        tspNavPanel.ImageList = AppHelper.ImageProvider.NavPanel32
        tspNavPanel.ImageScalingSize = New Size(32, 32)
      Else
        tspNavPanel.ImageList = AppHelper.ImageProvider.NavPanel16
        tspNavPanel.ImageScalingSize = New Size(16, 16)
      End If
      tsbDateView.ImageIndex = 0
      tsbHistoryView.ImageIndex = 1
      tsbPin.ImageIndex = 2
    End Set
  End Property

  'Handles the subclassing of the TreeView control so that we can detect the user trying to use the scrollbar
  'If we don't handle this then if the panel is not pinned it will disappear while trying to scroll
  Public Sub TreeViewSubClasserCallBack(ByRef pMessage As Message) Handles mvTreeViewSubclasser.CallBackProc
    Const WM_NCMOUSEMOVE As Integer = &HA0
    Const WM_NCMOUSELEAVE As Integer = &H2A2
    Const WM_NCLBUTTONDOWN As Integer = &HA1

    Select Case pMessage.Msg
      Case WM_NCMOUSEMOVE
        Tim.Stop()
      'Debug.Print(String.Format("Got Tree View NCMouseMove LParam {0} WParam {1}", pMessage.LParam, pMessage.WParam))
      Case WM_NCLBUTTONDOWN
        Tim.Stop()
      'Debug.Print(String.Format("Got Tree View NCLButtonDown LParam {0} WParam {1}", pMessage.LParam, pMessage.WParam))
      Case WM_NCMOUSELEAVE
        If Control.MouseButtons = System.Windows.Forms.MouseButtons.None Then
          mvTimerUsage = TimerUsage.tuMinimizeNavPanel
          Tim.Interval = 500
          Tim.Start()
        End If
        'Debug.Print(String.Format("Got Tree View NCMouseLeave LParam {0} WParam {1}", pMessage.LParam, pMessage.WParam))
    End Select
  End Sub

  Private Sub NavPanelMaxStart(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlTreeView.MouseEnter
    'Debug.WriteLine("Mouse Enter on pnlTreeView (NavPanelMaxStart)")
    mvTimerUsage = TimerUsage.tuMaximiseNavPanel
    Tim.Interval = 500
    Tim.Start()
  End Sub
  Private Sub NavPanelMaxEnd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlTreeView.MouseLeave
    'Debug.WriteLine("Mouse Leave on pnlTreeView (NavPanelMaxEnd)")
    Tim.Stop()
  End Sub

  Private Sub NavPanelMinStart(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tvw.MouseLeave
    'Debug.WriteLine("Mouse Leave on tvw (NavPanelMinStart)")
    mvTimerUsage = TimerUsage.tuMinimizeNavPanel
    Tim.Interval = 500
    Tim.Start()
  End Sub
  Private Sub NavPanelMinEnd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tvw.MouseEnter
    'Debug.WriteLine("Mouse Enter on tvw (NavPanelMinEnd)")
    Tim.Stop()
  End Sub

  Private Sub Tim_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tim.Tick
    Tim.Stop()
    Select Case mvTimerUsage
      Case TimerUsage.tuMaximiseNavPanel
        CheckMaximizeNavPanel()
      Case TimerUsage.tuMinimizeNavPanel
        CheckMinimizeNavPanel()
    End Select
    mvTimerUsage = TimerUsage.tuNone
  End Sub
  Private Sub CheckMaximizeNavPanel()
    If Not tsbPin.Checked And mvNavPanelMinimized Then
      pnlNavPanel.Width = mvNavPanelWidth
      tvw.Visible = True
      splNavPanel.Enabled = True
      mvNavPanelMinimized = False
    End If
  End Sub
  Private Sub CheckMinimizeNavPanel()
    If Not tsbPin.Checked AndAlso Not mvNavPanelMinimized Then
      pnlNavPanel.Width = tspNavPanel.Height
      tvw.Visible = False
      splNavPanel.Enabled = False
      mvNavPanelMinimized = True
    End If
  End Sub
  Private Sub splNavPanel_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splNavPanel.SplitterMoved
    mvNavPanelWidth = pnlNavPanel.Width
  End Sub

#Region "Treeview Events"

  Private Sub tvw_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tvw.DoubleClick
    DoNavigationSelection()
  End Sub
  Private Sub tvw_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles tvw.ItemDrag
    Try
      Dim vNode As TreeNode = CType(e.Item, TreeNode)
      If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
        Dim vItem As UserHistoryItem = DirectCast(vNode.Tag, UserHistoryItem)
        DragDropHistoryItem(vItem)
      End If
    Catch vCareEx As CareException
      If vCareEx.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
        ShowInformationMessage(InformationMessages.ImCannotFindContact)      'not used for Document
      Else
        DataHelper.HandleException(vCareEx)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub DragDropHistoryItem(vItem As UserHistoryItem)
    If vItem.HistoryEntityType = HistoryEntityTypes.hetContacts And vItem.Number > 0 Then
      Dim vContactInfo As New ContactInfo(vItem.Number)
      DoDragDrop(vContactInfo, DragDropEffects.Copy)
    ElseIf vItem.HistoryEntityType = HistoryEntityTypes.hetDocuments And vItem.Number > 0 Then
      DoDragDrop(New DocumentInfo(vItem.Number, vItem.Description), DragDropEffects.Copy)
    ElseIf vItem.HistoryEntityType = HistoryEntityTypes.hetEvents And vItem.Number > 0 Then
      DoDragDrop(New CareEventInfo(vItem.Number), DragDropEffects.Copy)
    End If
  End Sub
  Private Sub tvw_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tvw.KeyPress
    If e.KeyChar = ControlChars.Cr Then DoNavigationSelection()
  End Sub
  Private Sub tvw_AfterExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterExpand
    Try
      Dim vNode As TreeNode = e.Node
      If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
        Dim vItem As UserHistoryItem = DirectCast(vNode.Tag, UserHistoryItem)
        If vItem.HistoryEntityType = HistoryEntityTypes.hetSelectionSets Then
          If vNode.GetNodeCount(False) = 1 Then
            Dim vChildNode As TreeNode = vNode.FirstNode        'Either this is a place holder or a top level Contacts or Organisations type node
            If vChildNode.Tag IsNot Nothing Then
              Dim vChildItem As UserHistoryItem = DirectCast(vChildNode.Tag, UserHistoryItem)
              If vChildItem.HistoryEntityType = HistoryEntityTypes.hetSelectionSetPlaceHolder Then
                UserHistory.ReadSelectionSetData(vNode, vItem.Number, DataHelper.GetSelectionSetData(vItem.Number)) 'Read the children
                vChildNode.Remove()
              End If
            End If
          End If
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub tvw_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterSelect
    Dim vNode As TreeNode = DirectCast(e.Node, TreeNode)
    If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
      Dim vItem As UserHistoryItem = DirectCast(vNode.Tag, UserHistoryItem)
      Dim vList As New ArrayListEx
      mvBrowserMenu.EntityType = vItem.HistoryEntityType
      mvBrowserMenu.ItemNumber = vItem.Number
      mvBrowserMenu.ItemDescription = vNode.Text
      mvBrowserMenu.GroupCode = vItem.GroupCode
      mvBrowserMenu.Favourite = vItem.Favourite
      If (vItem.HistoryEntityType = HistoryEntityTypes.hetActions Or vItem.HistoryEntityType = HistoryEntityTypes.hetDocuments) Then
        If vItem.Number = 0 Then
          For Each vChildNode As TreeNode In vNode.Nodes
            vItem = DirectCast(vChildNode.Tag, UserHistoryItem)
            If vItem.Number > 0 Then vList.Add(vItem.Number)
          Next
        ElseIf vItem.Number > 0 Then
          vList.Add(vItem.Number)
        End If
      End If
      mvBrowserMenu.ItemList = vList
    End If
  End Sub
  Private Sub tvw_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvw.MouseDown
    If e.Button = System.Windows.Forms.MouseButtons.Right Then
      Dim vNode As TreeNode = tvw.GetNodeAt(e.X, e.Y)
      tvw.SelectedNode = vNode
      If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
      End If
    End If
  End Sub

  Private Sub DoNavigationSelection()
    Dim vCursor As New BusyCursor
    Try
      Dim vNode As TreeNode = tvw.SelectedNode
      If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
        Dim vItem As UserHistoryItem = DirectCast(vNode.Tag, UserHistoryItem)
        If vItem.Number > 0 Then
          Select Case vItem.HistoryEntityType
            Case HistoryEntityTypes.hetContacts
              CheckMinimizeNavPanel()
              FormHelper.ShowContactCardIndex(vItem.Number)
            Case HistoryEntityTypes.hetActions
              CheckMinimizeNavPanel()
              FormHelper.EditAction(vItem.Number)
            Case HistoryEntityTypes.hetDocuments
              CheckMinimizeNavPanel()
              FormHelper.EditDocument(vItem.Number)
            Case HistoryEntityTypes.hetEvents
              CheckMinimizeNavPanel()
              FormHelper.ShowEventIndex(vItem.Number)
            Case HistoryEntityTypes.hetMeetings
              CheckMaximizeNavPanel()
              FormHelper.EditMeeting(vItem.Number)
            Case HistoryEntityTypes.hetWorkstreams
              CheckMaximizeNavPanel()
              FormHelper.ShowWorkstreamIndex(vItem.GroupCode, vItem.Number)
            Case Else
              'Do Nothing
          End Select
        End If
        Me.Cursor = Cursors.Default
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

#End Region

  Private Sub mvBrowserMenu_Remove(ByVal pEntityType As HistoryEntityTypes, ByVal pNumber As Integer) Handles mvBrowserMenu.Remove
    Try
      UserHistory.RemoveOtherHistoryNode(pEntityType, pNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub tsbPin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPin.Click
    Try
      If tsbPin.Checked Then
        pnlNavPanel.Width = mvNavPanelWidth
        tvw.Visible = True
        splNavPanel.Enabled = True
        mvNavPanelMinimized = False
      Else
        CheckMinimizeNavPanel()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub tsbDateView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDateView.Click
    Try
      tsbHistoryView.Checked = False
      UserHistory.GetHistoryItems(tvw, HistoryViewModes.hvmDate)
      CheckMaximizeNavPanel()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub tsbHistoryView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbHistoryView.Click
    tsbDateView.Checked = False
    UserHistory.GetHistoryItems(tvw, HistoryViewModes.hvmHistory)
    CheckMaximizeNavPanel()
  End Sub
  Private Sub tspNavPanel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tspNavPanel.DoubleClick
    Try
      CheckMaximizeNavPanel()
      LargeNavPanelIcons = Not LargeNavPanelIcons
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

#End Region

#Region "Notifications"

  'The notification is now mostly handled in the MainHelper class

  Private Delegate Sub UpdateNotificationIconDelegate(ByVal pDataSet As DataSet)

  Public Sub UpdateNotificationIcon(ByVal pDataSet As DataSet)
    Dim vCount As Integer
    Dim vDocuments As Integer
    Dim vActions As Integer
    Dim vMeetings As Integer
    Dim vEntityAlerts As Integer
    Dim vIndex As Integer

    If sbr.InvokeRequired Then
      sbr.BeginInvoke(New UpdateNotificationIconDelegate(AddressOf UpdateNotificationIcon), New Object() {pDataSet})
    Else
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
      End If
      If vCount > 0 Then
        sbpNotify.Icon = My.Resources.iconNotifications
        sbpNotify.Text = String.Format(InformationMessages.ImNotificationsItems, vCount)
        Dim vToolTipText As New StringBuilder
        vToolTipText.Append(InformationMessages.ImNotifications)
        vToolTipText.Append(" ")
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
        sbpNotify.ToolTipText = vToolTipText.ToString
      Else
        sbpNotify.Icon = My.Resources.iconNoNotifications
        sbpNotify.Text = ""
        sbpNotify.ToolTipText = InformationMessages.ImNoNotifications
      End If
    End If
  End Sub

#End Region

  Private Enum NavigationTargets
    Other
    NavigationPanel
    SelectionPanel
  End Enum

  Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    If e.Control AndAlso e.KeyValue = Keys.Return Then
      Dim vFormWithSelectionPanel As IPanelVisibility = Nothing
      If Me.ActiveMdiChild IsNot Nothing AndAlso TypeOf Me.ActiveMdiChild Is IPanelVisibility Then vFormWithSelectionPanel = DirectCast(Me.ActiveMdiChild, IPanelVisibility)
      Dim vTarget As NavigationTargets = NavigationTargets.Other

      If NavigationPanel AndAlso tvw.Focused Then
        If vFormWithSelectionPanel IsNot Nothing Then vTarget = NavigationTargets.SelectionPanel
        'Goto active form selection panel - or something on the form if there is not a selection panel
      ElseIf vFormWithSelectionPanel IsNot Nothing Then
        'If we have a form with a selection panel then
        If vFormWithSelectionPanel.PanelHasFocus = True Then
          'If the selection panel has the focus then Goto something on the form
        Else
          'If it does not have the focus then goto the nav panel if it is switched on - if not goto selection panel
          If NavigationPanel Then
            vTarget = NavigationTargets.NavigationPanel
          Else
            vTarget = NavigationTargets.SelectionPanel
          End If
        End If
      Else
        If NavigationPanel AndAlso tvw.Focused = False Then
          vTarget = NavigationTargets.NavigationPanel
        End If
      End If

      Select Case vTarget
        Case NavigationTargets.NavigationPanel
          tvw.Focus()
        Case NavigationTargets.SelectionPanel
          vFormWithSelectionPanel.PanelHasFocus = True
        Case Else
          If vFormWithSelectionPanel IsNot Nothing Then
            vFormWithSelectionPanel.PanelHasFocus = False
          ElseIf Me.ActiveMdiChild IsNot Nothing Then
            Me.ActiveMdiChild.Activate()
            Dim vControl As Control = DirectCast(Me.ActiveMdiChild, Control).GetNextControl(Nothing, True)
            If vControl IsNot Nothing Then vControl.Focus()
          End If
      End Select
      e.SuppressKeyPress = True
    End If
  End Sub

  Private Sub frmMain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
    'Debug.Print("e.keychar" & e.KeyChar)
  End Sub

  Public ReadOnly Property MainMenu() As MainMenu Implements IMainForm.MainMenu
    Get
      Return mvMainMenu
    End Get
  End Property

  Private Sub SetClassicView()
    Me.SuspendLayout()
    SetBackgroundColor(SystemColors.AppWorkspace)
    splNavPanel.BackColor = SystemColors.Control
    pnlNavPanel.Show()
    pnlMDIExplorer.Hide()
    pnlExplorerTop.Hide()
    Me.MainMenuStrip.Show()
    mvShowClassicToolstrip = Settings.ShowToolbar
    MainHelper.ShowToolbar = Settings.ShowToolbar
    Me.FormBorderStyle = FormBorderStyle.Sizable
    Me.Padding = New Padding(0)
    Me.ResumeLayout()
  End Sub

  Private Sub SetExplorerView()
    Me.SuspendLayout()
    InitialiseExplorerViewModel()
    MainHelper.SetFlatMdiArea(True)
    SetBackgroundColor(SystemColors.Window)
    mvShowClassicToolstrip = MainHelper.ShowToolbar
    MainHelper.ShowToolbar = False
    Me.MainMenuStrip.Hide()
    NavigationPanel = True
    pnlNavPanel.Width = 125
    pnlNavPanel.Hide()
    pnlMDIExplorer.Width = pnlNavPanel.Width
    If pnlMDIExplorer.Width < 260 Then
      pnlMDIExplorer.Width = 260
    End If

    Dim appThemeDic As New System.Windows.ResourceDictionary()
    appThemeDic.Source = New Uri("pack://application:,,,/CDBNETXAML;component/Themes/AppTheme.xaml")
    Dim newClr As Color = SystemColors.Control
    If appThemeDic.Contains("AppBackgroundBrush") Then
      Dim vBrush As System.Windows.Media.SolidColorBrush = TryCast(appThemeDic.Item("AppBackgroundBrush"), System.Windows.Media.SolidColorBrush)
      If vBrush IsNot Nothing Then
        newClr = Color.FromArgb(vBrush.Color.A, vBrush.Color.R, vBrush.Color.G, vBrush.Color.B)
      End If
    End If
      For Each ctl As Control In pnlMDIExplorer.Controls
        ctl.BackColor = newClr
      Next
      splNavPanel.BackColor = newClr
    splNavPanel.Left = pnlNavPanel.Width
    splNavPanel.Enabled = True
    sbr.Hide()
    pnlMDIExplorer.Show()
    pnlExplorerTop.Show()

    If Me.Visible = True AndAlso (MainHelper.Forms Is Nothing OrElse MainHelper.Forms.Count = 0) Then
      GoHome()
    End If

    Dim sz As Size = Me.Size
    Dim loc As Point = Me.Location

    Me.ResumeLayout()

  End Sub
  Private Sub MaximizeModernView()
    Me.Bounds = Screen.GetWorkingArea(Me)
    'Me.MaximizedBounds = Screen.GetWorkingArea(Me)
    Me.Location = Screen.GetWorkingArea(Me).Location
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
    Me.MaximumSize = New Size(My.Computer.Screen.WorkingArea.Size.Width, My.Computer.Screen.WorkingArea.Size.Height - 1)
    Me.WindowState = FormWindowState.Minimized
    Me.WindowState = FormWindowState.Maximized
  End Sub
  Private Sub InitialiseExplorerViewModel()
    If mvExplorerViewModel Is Nothing Then
      Dim explorerItems As CDBNETXAMLPages.MDIExplorerViewModelItems = ConstructExplorerUIItems()
      mvExplorerViewModel = New CDBNETXAMLPages.MDIExplorerViewModel(explorerItems)
      wpfMDIExplorerInstance.DataContext = mvExplorerViewModel
      ConstructExplorerMenus()
      'wpfMenuHeaderInstance.DataContext = mvMenuViewModel
      mvOpenFormsStrip = New OpenFormsStrip()
      wpfTabStripHost.Child = mvOpenFormsStrip
      mvOpenFormsViewModel = New OpenFormsViewModel(Me)
      mvOpenFormsViewModel.ExplorerCommand = New RelayCommand(Of ExplorerCommands, Object)(AddressOf DoExplorerCommand)
      mvOpenFormsViewModel.AllCommands = mvHomeScreenViewModel.AllCommands
    End If
    mvOpenFormsStrip.DataContext = mvOpenFormsViewModel
    mvExplorerViewModel.Refresh()
  End Sub

  Private Sub DoExplorerCommand(pCommand As ExplorerCommands, param1 As Object)
    Select Case pCommand
      Case CDBNETXAML.ExplorerCommands.GoHome
        GoHome()
      Case CDBNETXAML.ExplorerCommands.ShowMenuWindow
        ShowMenuWindow()
      Case CDBNETXAML.ExplorerCommands.SearchText
        If param1 IsNot Nothing AndAlso TypeOf (param1) Is String Then
          MainHelper.SmartSearch(param1.ToString())
        End If
      Case CDBNETXAML.ExplorerCommands.ShowClassicView
        SetClassicView()
      Case CDBNETXAML.ExplorerCommands.RequestAppClose
        Me.SuspendLayout()
        SetClassicView()
        Me.Close()
      Case CDBNETXAML.ExplorerCommands.NavigateHistory
        'tabStripHost.Show()
        If param1 IsNot Nothing AndAlso TypeOf (param1) Is UserHistoryItem Then
          NavigateHistoryItem(DirectCast(param1, UserHistoryItem))
        End If
      Case CDBNETXAML.ExplorerCommands.DragHistory
        Dim historyItem As UserHistoryItem = TryCast(param1, UserHistoryItem)
        If historyItem IsNot Nothing Then
          DragDropHistoryItem(historyItem)
        End If
      Case CDBNETXAML.ExplorerCommands.CloseMenu
        CloseMenuWindow()
        'tabStripHost.Show()
      'explorerViewModel_Command(CDBNETXAML.ExplorerCommands.GoHome)
      Case ExplorerCommands.CloseApp
        MainHelper.CloseAllForms()
      Case ExplorerCommands.MinimizeMDI
        Me.WindowState = If(Me.WindowState = FormWindowState.Minimized, FormWindowState.Maximized, FormWindowState.Minimized)
      Case ExplorerCommands.ShowDashboard
        MainMenu.Execute(CommandIndexes.cbiDashboard)
    End Select
  End Sub

  Private Sub SetBackgroundColor(ByVal color As System.Drawing.Color)
    Dim ctlMDI As MdiClient

    For Each ctl In Me.Controls
      If TypeOf ctl Is MdiClient Then
        Try
          ctlMDI = CType(ctl, MdiClient)
          ctlMDI.BackColor = color
          Exit For
        Catch exc As InvalidCastException
        End Try
      End If
    Next
  End Sub

  Private Sub SetTheme(themeName As String)
    Dim queryCriteria As New ParameterList(True)
    Dim setTheme As Boolean = False

    queryCriteria("XmlDataType") = "AT"   'Appearance themes
    Dim themeTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, queryCriteria, False)
    Dim results As DataRow() = Nothing

    If themeTable IsNot Nothing Then results = themeTable.Select("ItemDesc = '" + themeName + "'")

    If results IsNot Nothing AndAlso results.Count > 0 Then
      Settings.AppearanceThemeID = Convert.ToInt32(results(0)("XmlItemNumber").ToString())
      setTheme = True
    End If

    queryCriteria = New ParameterList(True)
    queryCriteria("XmlDataType") = "FT"   'Font themes
    themeTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, queryCriteria, True)

    results = Nothing
    If themeTable IsNot Nothing Then results = themeTable.Select("ItemDesc = '" + themeName + "'")

    If results IsNot Nothing AndAlso results.Count > 0 Then
      Settings.FontThemeID = Convert.ToInt32(results(0)("XmlItemNumber").ToString())
      setTheme = True
    End If

    If setTheme Then
      DisplayTheme.InitFromSettings()
      MainHelper.SetControlColors()
    End If

  End Sub

  '#Region "Explorer Events"

  Private Sub explorerViewModel_Command(command As CDBNETXAML.ExplorerCommands, Optional param1 As Object = Nothing, Optional param2 As Object = Nothing) Handles mvExplorerViewModel.SendExplorerCommand
    DoExplorerCommand(command, param1)
  End Sub

  Private Sub GoHome()
    'All commands to process when the user presses the Home button in Explorer view
    ShowMenuWindow()
  End Sub

  Private Sub NavigateHistoryItem(vItem As UserHistoryItem)
    MainHelper.NavigateHistoryItem(vItem.HistoryEntityType, vItem.Number)
  End Sub


  '#End Region

  Private Function ConstructExplorerUIItems() As CDBNETXAMLPages.MDIExplorerViewModelItems

    Dim vRtn As New CDBNETXAMLPages.MDIExplorerViewModelItems

    UserHistory.InitialiseForModernView(Function(item) mvBrowserMenu.InitialiseContextCommands(item))

    Return vRtn
  End Function

  Private Sub AddEntityGroup(masterGroup As CollectionList(Of EntityGroup), appendGroup As CollectionList(Of EntityGroup))
    If appendGroup IsNot Nothing And masterGroup IsNot Nothing Then
      For Each group As EntityGroup In appendGroup
        masterGroup.Add(group.Code, group)
      Next
    End If

  End Sub

  Private Sub ShowMenuWindow()

    If mvHomeScreenViewModel Is Nothing Then
      ConstructExplorerMenus()
    End If
    FormHelper.ShowExplorerMenuViewer(mvHomeScreenViewModel)
  End Sub

  Private Sub ConstructExplorerMenus()
    Dim vCommands As List(Of MenuToolbarCommand) = MainMenu.MenuItems.Cast(Of MenuToolbarCommand).Where(Function(item) item.HideItem = False).ToList()
    'If mvMenuViewModel Is Nothing Then
    '  mvMenuViewModel = New MenuViewModel(vCommands)
    '  AddHandler mvMenuViewModel.SendExplorerCommand, AddressOf explorerViewModel_Command
    'End If
    If mvHomeScreenViewModel Is Nothing Then
      mvHomeScreenViewModel = New HomeScreenViewModel(vCommands)
    End If
  End Sub

  Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

  Private Sub CloseMenuWindow()
    For Each frm As Form In Me.MdiChildren
      If TypeOf (frm) Is frmModernMenuViewer Then
        frm.Close()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("MdiChildren"))
        Exit For
      End If
    Next
  End Sub

  Private Sub titlePanel_MouseDown(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
    If e.LeftButton = System.Windows.Input.MouseButtonState.Pressed AndAlso e.ClickCount = 1 Then
      If Me.WindowState = FormWindowState.Maximized Then
        Me.WindowState = FormWindowState.Normal
        Dim vArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Me.Size = New Size(CInt(vArea.Size.Width * 0.8), CInt(vArea.Size.Height * 0.8))
      End If
      mvDragging = True
      Debug.Print(String.Format("{0}  Setting Dragging to True", Date.Now.ToString()))
      mvDragCusorPoint = System.Windows.Forms.Cursor.Position
      mvDragFormPoint = Me.Location
    End If
  End Sub



  Private Sub titlePanel_MouseMove(sender As Object, e As System.Windows.Input.MouseEventArgs)
    If mvDragging Then
      Dim vPointDifference As Point = Point.Subtract(System.Windows.Forms.Cursor.Position, New Size(mvDragCusorPoint))
      Dim vNewLoc As Point = Point.Add(mvDragFormPoint, New Size(vPointDifference))
      If vNewLoc.X < -10 Then vNewLoc.X = -10
      If vNewLoc.Y + Me.Width < 20 Then vNewLoc.Y = Me.Width - 20
      Me.Location = vNewLoc
    End If
  End Sub

  Private Sub titlePanel_MouseUp(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
    mvDragging = False
    Debug.Print(String.Format("{0}  Setting Dragging to False From MouseUp", Date.Now.ToString()))
  End Sub

  Private Sub title_Panel_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
    If Me.WindowState = FormWindowState.Maximized Then
      Me.WindowState = FormWindowState.Normal
    Else
      mvDragging = False
      Debug.Print(String.Format("{0}  Setting Dragging to False From DoubleClick", Date.Now.ToString()))
      MaximizeModernView()
    End If
  End Sub

  Private Sub titlePanel_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs)
    If e.LeftButton = System.Windows.Input.MouseButtonState.Released Then
      mvDragging = False
    Else
      'Sometimes the user is too fast and the code can't keep up with moving the screen under the mouse.  So you get a Mouse Leave.  Still not perfect but it's an improvement
      If mvDragging Then titlePanel_MouseMove(sender, e)
    End If
  End Sub

End Class

Public Class FormDocker
  Implements INotifyPropertyChanged

  Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

  Private mvCurrentChild As IMdiDockable
  Private mvMdiParent As Form
  Private mvMdiClient As MdiClient
  Private mvDockPosition As DockStyle = DockStyle.None
  Private mvMdiClientSize As Size = Nothing
  Private mvDraggingCaption As Boolean = False
  Private mvDockPositionChanged As Boolean = False

  Private mvDockedForms As Dictionary(Of IMdiDockable, DockStyle) = Nothing

  Private ReadOnly Property ActiveMdiChild As IMdiDockable
    Get
      Return TryCast(mvMdiParent.ActiveMdiChild, IMdiDockable)
    End Get
  End Property

  Private Property DockPosition() As DockStyle
    Get
      Return mvDockPosition
    End Get
    Set(value As DockStyle)
      DockPositionChanged = (mvDockPosition <> value)
      If DockPositionChanged Then
        mvDockPosition = value
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("DockPosition"))
      End If
    End Set
  End Property

  Private ReadOnly Property DockedForms() As Dictionary(Of IMdiDockable, DockStyle)
    Get
      If mvDockedForms Is Nothing Then
        mvDockedForms = New Dictionary(Of IMdiDockable, DockStyle)
      End If
      Return mvDockedForms
    End Get
  End Property

  Private Property DraggingCaption As Boolean
    Get
      Return mvDraggingCaption
    End Get
    Set(value As Boolean)
      mvDraggingCaption = value
    End Set
  End Property

  Private Property DockPositionChanged As Boolean
    Get
      Return mvDockPositionChanged
    End Get
    Set(value As Boolean)
      mvDockPositionChanged = value
    End Set
  End Property

  Public Sub New(pParentForm As Form)
    If pParentForm.IsMdiContainer Then
      mvMdiParent = pParentForm
      AddHandler mvMdiParent.MdiChildActivate, AddressOf ChildActivate
      For Each vCtl As Control In pParentForm.Controls
        If TypeOf vCtl Is MdiClient Then
          Try
            vCtl.SuspendLayout()
            AddHandler vCtl.Resize, AddressOf MdiClientResizing
            AddHandler vCtl.Paint, AddressOf PaintClientRectangle
            mvMdiClient = CType(vCtl, MdiClient)
            Dim method As Reflection.MethodInfo = DirectCast(vCtl, MdiClient).[GetType]().GetMethod("SetStyle", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            method.Invoke(DirectCast(vCtl, MdiClient), New [Object]() {ControlStyles.OptimizedDoubleBuffer, True})
            method.Invoke(DirectCast(vCtl, MdiClient), New [Object]() {ControlStyles.AllPaintingInWmPaint, True})
            method.Invoke(DirectCast(vCtl, MdiClient), New [Object]() {ControlStyles.UserPaint, True})
            Exit For
          Catch exc As InvalidCastException
          Finally
            vCtl.ResumeLayout()
          End Try
        End If
      Next
    Else
      Throw New Exception("Form Dock Parent is not MDI")
    End If
  End Sub

  Public Sub ChildActivate(sender As Object, e As System.EventArgs)
    If ActiveMdiChild IsNot Nothing AndAlso ActiveMdiChild IsNot mvCurrentChild Then
      mvCurrentChild = ActiveMdiChild
      HookChildEvents(ActiveMdiChild)
    End If
  End Sub

  Private Sub HookChildEvents(childForm As IMdiDockable)
    AddHandler childForm.CaptionDrag, AddressOf CaptionDrag
    AddHandler childForm.CaptionDragReleased, AddressOf CaptionDragReleased
    AddHandler DirectCast(childForm, Form).FormClosing, AddressOf FormClosing
  End Sub
  Private Sub UnhookChildEvents(childForm As IMdiDockable)
    RemoveHandler childForm.CaptionDrag, AddressOf CaptionDrag
    RemoveHandler childForm.CaptionDragReleased, AddressOf CaptionDragReleased
    RemoveHandler DirectCast(childForm, Form).FormClosing, AddressOf FormClosing
  End Sub

  Private Sub CaptionDrag(curPos As Point)
    If mvMdiClient IsNot Nothing Then
      'Remove docked form as dragging can cause re-sizing which will attempt to reposition any docked forms
      RemoveDockedForm(ActiveMdiChild)
      DraggingCaption = True
      RefreshMdiClient()

      Dim clientPos As Point = mvMdiClient.PointToClient(curPos)

      If clientPos.X = mvMdiClient.ClientRectangle.Left Then
        DockPosition = DockStyle.Left
      ElseIf clientPos.Y = mvMdiClient.ClientRectangle.Top Then
        DockPosition = DockStyle.Top
      ElseIf clientPos.X >= (mvMdiClient.ClientRectangle.Right - mvMdiClient.Margin.Right) Then
        DockPosition = DockStyle.Right
      Else
        DockPosition = DockStyle.None
      End If
    End If
  End Sub

  Private Sub CaptionDragReleased()
    DraggingCaption = False
    AddDockedForm(ActiveMdiChild, DockPosition)
    DockFromDockedForms()
    DockPosition = DockStyle.None
  End Sub

  Private Sub AddDockedForm(ByVal pForm As IMdiDockable, ByVal pDockPosition As DockStyle)
    If Not DockedForms.ContainsKey(pForm) Then
      DockedForms.Add(pForm, pDockPosition)
    End If
  End Sub

  Private Sub RemoveDockedForm(ByVal pForm As IMdiDockable)
    If DockedForms.ContainsKey(pForm) Then
      DockedForms.Remove(pForm)
    End If
  End Sub

  Private Sub FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
    If sender IsNot Nothing AndAlso TypeOf sender Is IMdiDockable Then
      RemoveDockedForm(CType(sender, IMdiDockable))
      UnhookChildEvents(CType(sender, IMdiDockable))
    End If
  End Sub

  Private Sub MdiClientResizing(ByVal sender As System.Object, ByVal e As System.EventArgs)
    RefreshMdiClient()
    DockFromDockedForms()
  End Sub

  Private Sub DockFromDockedForms()
    If DockedForms.Count > 0 Then
      For Each vKVP As KeyValuePair(Of IMdiDockable, DockStyle) In DockedForms
        DockForm(vKVP.Key, vKVP.Value)
      Next
    End If
  End Sub

  Private Sub DockForm(ByVal pForm As IMdiDockable, ByVal pDockPosition As DockStyle)
    If pForm IsNot Nothing Then
      If mvMdiClient IsNot Nothing Then
        mvMdiClient.Invalidate()
        Dim vForm As Form = DirectCast(pForm, Form)
        With vForm
          Select Case pDockPosition
            Case DockStyle.None
              'Do nothing
            Case DockStyle.Left
              .Left = mvMdiClient.ClientRectangle.Left
              .Top = mvMdiClient.ClientRectangle.Top
              .Width = CInt(mvMdiClient.ClientRectangle.Width / 2)
              .Height = mvMdiClient.ClientRectangle.Height
            Case DockStyle.Right
              .Left = CInt((mvMdiClient.ClientRectangle.Width / 2))
              .Top = mvMdiClient.ClientRectangle.Top
              .Width = CInt(mvMdiClient.ClientRectangle.Width / 2)
              .Height = mvMdiClient.ClientRectangle.Height
            Case DockStyle.Top
              .Left = mvMdiClient.ClientRectangle.Left
              .Top = mvMdiClient.ClientRectangle.Top
              .Width = mvMdiClient.ClientRectangle.Width
              .Height = mvMdiClient.ClientRectangle.Height
          End Select
        End With
      End If
    End If
  End Sub

  Private Sub FormDocker_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
    If e.PropertyName = "DockPosition" Then
      'If dock position changes we need to refresh and redraw the mdi client control
      RefreshMdiClient()
    End If
  End Sub

  Private Sub RefreshMdiClient()
    If mvMdiClient IsNot Nothing Then
      mvMdiClient.Refresh()
    End If
  End Sub

  Protected Sub PaintClientRectangle(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    If DockPosition <> DockStyle.None OrElse DraggingCaption Then
      Dim fillPadding As Integer = 2

      With mvMdiClient.ClientRectangle
        If DockPosition <> DockStyle.None Then
          Dim borderRectangle As Rectangle
          fillPadding = 0
          If DockPosition = DockStyle.Left Then
            borderRectangle = New Rectangle(.Left + fillPadding, .Top + fillPadding, CInt((.Width - (fillPadding * 2)) / 2), CInt(.Height - (fillPadding * 2)))
          ElseIf DockPosition = DockStyle.Right Then
            borderRectangle = New Rectangle(CInt((.Width / 2) + fillPadding), .Top + fillPadding, CInt((.Width - (fillPadding * 2)) / 2), CInt(.Height - (fillPadding * 2)))
          ElseIf DockPosition = DockStyle.Top Then
            borderRectangle = New Rectangle(.Left + fillPadding, .Top + fillPadding, CInt(.Width - (fillPadding * 2)), CInt(.Height - (fillPadding * 2)))
          End If
          Dim colour As Color = Color.FromArgb(90, 255, 255, 255)
          Dim b As New SolidBrush(colour)

          e.Graphics.FillRectangle(b, borderRectangle)
        Else
          'Dragging caption
          Dim colour As Color = Color.FromArgb(170, 255, 255, 255)
          Dim brush As New SolidBrush(colour)
          Dim dRWidth As Integer = 6
          Dim dRHeight As Integer = 40
          Dim rightBorderOffset As Integer = 1

          Dim borderLeftRectangle As New Rectangle(.Left + fillPadding, CInt((.Height - dRHeight) / 2), dRWidth, dRHeight)
          Dim borderRightRectangle As New Rectangle(.Right - (dRWidth + fillPadding + rightBorderOffset), CInt((.Height - dRHeight) / 2), dRWidth, dRHeight)
          Dim borderTopRectangle As New Rectangle(CInt((.Width - dRHeight) / 2), .Top + fillPadding, dRHeight, dRWidth)

          Dim borderRectangleCollection As New CollectionList(Of Rectangle)
          borderRectangleCollection.Add(DockStyle.Left.ToString, borderLeftRectangle)
          borderRectangleCollection.Add(DockStyle.Right.ToString, borderRightRectangle)
          borderRectangleCollection.Add(DockStyle.Top.ToString, borderTopRectangle)

          For Each borderRectangleItem As Rectangle In borderRectangleCollection
            DrawRoundedRectangle(e.Graphics, borderRectangleItem, CInt(borderRectangleItem.Height * 0.1), New Pen(brush), colour)
          Next
        End If
      End With
    End If
  End Sub

  Private Sub DrawRoundedRectangle(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal cornerRadius As Integer, ByVal drawPen As Pen, ByVal fillColour As Color)
    Dim strokeOffset As Integer = Convert.ToInt32(Math.Ceiling(drawPen.Width))
    bounds = Rectangle.Inflate(bounds, -strokeOffset, -strokeOffset)

    drawPen.EndCap = LineCap.Round
    drawPen.StartCap = LineCap.Round

    Dim gfxPath As New GraphicsPath()
    gfxPath.AddArc(bounds.X, bounds.Y, cornerRadius, cornerRadius, 180, 90)
    gfxPath.AddArc(bounds.X + bounds.Width - cornerRadius, bounds.Y, cornerRadius, cornerRadius, 270, 90)
    gfxPath.AddArc(bounds.X + bounds.Width - cornerRadius, bounds.Y + bounds.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90)
    gfxPath.AddArc(bounds.X, bounds.Y + bounds.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90)
    gfxPath.CloseAllFigures()

    g.FillPath(New SolidBrush(fillColour), gfxPath)
    g.DrawPath(drawPen, gfxPath)
  End Sub

End Class